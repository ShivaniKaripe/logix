<%@ Page CodeFile="..\ConnectorCB.vb" Debug="true" Inherits="ConnectorCB" Language="vb" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%
  ' *****************************************************************************
  ' * FILENAME: OfferValidation.aspx
  ' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' * Copyright © 2002 - 2009.  All rights reserved by:
  ' *
  ' * NCR Corporation
  ' * 1435 Win Hentschel Boulevard
  ' * West Lafayette, IN  47906
  ' * voice: 888-346-7199  fax: 765-464-1369
  ' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' *
  ' * PROJECT : NCR Advanced Marketing Solution
  ' *
  ' * MODULE  : Logix
  ' *
  ' * PURPOSE :
  ' *
  ' * NOTES   :
  ' *
  ' * Version : 7.3.1.138972
  ' *
  ' *****************************************************************************
%>
<script runat="server">
  Public Common As New Copient.CommonInc
  Public Connector As New Copient.ConnectorInc
  Public GZIP As New Copient.GZIPInc
  Public TextData As String
  Public IPL As Boolean
  Public LogFile As String
  Public FileStamp As String
  Public FileNum As Integer
  Public StartTime As Decimal
  Public DebugMode As Boolean = False
  Public MacAddress As String
  Public LocalServerIP As String
  Public LocalServerID As Long

  Function Handle_Post(ByVal LocalServerID As Long, ByVal LocationID As Long, ByRef Data As String) As Boolean

    Dim CompressedData As String
    Dim InboundData As String
    Dim FileData() As Byte
    Dim Checksum As String
    Dim MD5 As String
    Dim DataRetrieved As Boolean = False

    Try
      Send_Response_Header("Offer Validation", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
      MD5 = Request.QueryString("MD5")
      InboundData = ""
      If Request.Files.Count > 0 Then
        ReDim FileData(Request.Files(0).ContentLength)
        Request.Files(0).InputStream.Read(FileData, 0, Request.Files(0).ContentLength)
        'uncomment to view raw data
        'Send(Encoding.Default.GetString(FileData))
        CompressedData = Encoding.Default.GetString(FileData)
        FileData = Nothing
        InboundData = GZIP.DecompressString(CompressedData)
        CompressedData = Nothing
        'uncomment to view decompressed data
        'Send(Inbounddata)

        If (Not DebugMode) Then
          If (MD5 <> "") Then
            'if we got an MD5 checksum, make sure it matches the compressed data we received
            Checksum = Common.MD5(InboundData)
            If Checksum <> MD5 Then
              Common.Write_Log(LogFile, "Bad MD5 .. LocalServer sent ->" & MD5 & "     CentralServer computed ->" & Checksum)
              Data = "Bad MD5"
              DataRetrieved = False
            Else
              Common.Write_Log(LogFile, "Valid MD5 ->" & MD5)
              Common.Write_Log(LogFile, "GZip decompression successful ... size after unzipping is " & Format(Len(InboundData), "###,###,###,###,##0") & " bytes")
              Common.Write_Log(LogFile, "Received Data:" & vbCrLf & InboundData)
              Data = InboundData
              DataRetrieved = True
            End If
          Else
            If InboundData = "no data" Then
              Common.Write_Log(LogFile, "Server had no data to upload - no processing performed")
              Data = "ACK"
              DataRetrieved = False
            End If
          End If
        Else
          Data = InboundData
          DataRetrieved = True
        End If

      Else
        Common.Write_Log(LogFile, "No files were uploaded")
        Data = "No files were uploaded"
        DataRetrieved = False
      End If
    Catch ex As Exception
      Send(ex.ToString)
      Common.Error_Processor()
      Common.Write_Log(LogFile, ex.ToString())
    End Try

    Return DataRetrieved
  End Function

  Function ProcessData(ByVal data As String, ByVal LocationID As String) As String

    Dim arrRecs() As String = Nothing
    Dim cols() As String = Nothing
    Dim record As String = Nothing
    Dim i As Integer
    Dim row As DataRow = Nothing
    Dim dtGraphic As DataTable = Nothing
    Dim dtUserGroup As DataTable = Nothing
    Dim dtProdGroup As DataTable = Nothing
    Dim dtIncentive As DataTable = Nothing
    Dim responseText As String = ""
    Dim graphicKeys As New ArrayList(100)
    Dim userGroupKeys As New ArrayList(100)
    Dim prodGroupKeys As New ArrayList(100)
    Dim incentiveKeys As New ArrayList(100)
    Dim stopWatch As New System.Diagnostics.Stopwatch()
    Dim TempDate As DateTime
    Dim ErrMsg As String = ""

    Try
      stopWatch.Start()
      arrRecs = data.Split(vbCrLf)

      MacAddress = Trim(Request.QueryString("mac"))

      If MacAddress = "" Or MacAddress = "0" Then
        MacAddress = Trim(Request.UserHostAddress)
      End If
      LocalServerID = Common.Extract_Val(Request.QueryString("serial"))
      LocalServerIP = Common.Extract_Val(Request.QueryString("IP"))
      If LocalServerIP = "" Or LocalServerIP = "0" Then
        LocalServerIP = MacAddress & " IP from requesting browser. "
      End If

      For i = 0 To arrRecs.GetUpperBound(0)
        record = arrRecs(i)
        cols = record.Split(Chr(30))
        If (cols.Length > 0) Then
          Select Case cols(0).Trim().ToUpper
            Case "GRAPHIC"
              If (dtGraphic Is Nothing) Then
                dtGraphic = New DataTable
                dtGraphic.Columns.Add("LocationID", System.Type.GetType("System.Int64"))
                dtGraphic.Columns.Add("OnScreenAdID", System.Type.GetType("System.Int64"))
                dtGraphic.Columns.Add("DateValidated", System.Type.GetType("System.DateTime"))
                dtGraphic.Columns.Add("MD5", System.Type.GetType("System.String"))
                dtGraphic.Columns.Add("UpdateLevel", System.Type.GetType("System.Int32"))
              End If

              If (TryParseDateSent(cols(2), TempDate)) Then
                row = dtGraphic.NewRow()
                cols(0) = LocationID
                'row.ItemArray = cols
                row.Item("LocationID") = LocationID
                row.Item("OnScreenAdID") = cols(1)
                row.Item("DateValidated") = TempDate
                row.Item("MD5") = cols(3)
                row.Item("UpdateLevel") = cols(4)
                dtGraphic.Rows.Add(row)
                graphicKeys.Add(cols(1))
              Else
                If ErrMsg = "" Then
                  ErrMsg = "Incorrect date format found in offer validation record(s) " & vbCrLf & vbCrLf & _
                           "Process Info: Server Name=" & Environment.MachineName & " Invoking LocationID=" & LocationID & vbCrLf
                End If
                ErrMsg &= record
              End If

            Case "USERGROUP"
              If (dtUserGroup Is Nothing) Then
                dtUserGroup = New DataTable
                dtUserGroup.Columns.Add("LocationID", System.Type.GetType("System.Int64"))
                dtUserGroup.Columns.Add("CustomerGroupID", System.Type.GetType("System.Int64"))
                dtUserGroup.Columns.Add("DateValidated", System.Type.GetType("System.DateTime"))
                dtUserGroup.Columns.Add("CountReported", System.Type.GetType("System.Int32"))
                dtUserGroup.Columns.Add("CountActual", System.Type.GetType("System.Int32"))
                dtUserGroup.Columns.Add("UpdateLevel", System.Type.GetType("System.Int32"))
              End If

              If (TryParseDateSent(cols(2), TempDate)) Then
                row = dtUserGroup.NewRow()
                cols(0) = LocationID
                'row.ItemArray = cols
                row.Item("LocationID") = LocationID
                row.Item("CustomerGroupID") = cols(1)
                row.Item("DateValidated") = TempDate
                row.Item("CountActual") = DBNull.Value
                row.Item("CountReported") = cols(3)
                row.Item("UpdateLevel") = cols(4)
                dtUserGroup.Rows.Add(row)
                userGroupKeys.Add(cols(1))
              Else
                If ErrMsg = "" Then
                  ErrMsg = "Incorrect date format found in offer validation record(s) " & vbCrLf & vbCrLf & _
                           "Process Info: Server Name=" & Environment.MachineName & " Invoking LocationID=" & LocationID & vbCrLf
                End If
                ErrMsg &= record
              End If

            Case "PRODGROUP"
              If (dtProdGroup Is Nothing) Then
                dtProdGroup = New DataTable
                dtProdGroup.Columns.Add("LocationID", System.Type.GetType("System.Int64"))
                dtProdGroup.Columns.Add("ProductGroupID", System.Type.GetType("System.Int64"))
                dtProdGroup.Columns.Add("DateValidated", System.Type.GetType("System.DateTime"))
                dtProdGroup.Columns.Add("CountReported", System.Type.GetType("System.Int32"))
                dtProdGroup.Columns.Add("CountActual", System.Type.GetType("System.Int32"))
                dtProdGroup.Columns.Add("UpdateLevel", System.Type.GetType("System.Int32"))
              End If

              If (TryParseDateSent(cols(2), TempDate)) Then
                row = dtProdGroup.NewRow()
                cols(0) = LocationID
                'row.ItemArray = cols
                row.Item("LocationID") = LocationID
                row.Item("ProductGroupID") = cols(1)
                row.Item("DateValidated") = TempDate
                row.Item("CountActual") = DBNull.Value
                row.Item("CountReported") = cols(3)
                row.Item("UpdateLevel") = cols(4)
                dtProdGroup.Rows.Add(row)
                prodGroupKeys.Add(cols(1))
              Else
                If ErrMsg = "" Then
                  ErrMsg = "Incorrect date format found in offer validation record(s) " & vbCrLf & vbCrLf & _
                           "Process Info: Server Name=" & Environment.MachineName & " Invoking LocationID=" & LocationID & vbCrLf
                End If
                ErrMsg &= record
              End If

            Case "INCENTIVE"
              If (dtIncentive Is Nothing) Then
                dtIncentive = New DataTable
                dtIncentive.Columns.Add("LocationID", System.Type.GetType("System.Int64"))
                dtIncentive.Columns.Add("IncentiveID", System.Type.GetType("System.Int64"))
                dtIncentive.Columns.Add("DateValidated", System.Type.GetType("System.DateTime"))
                dtIncentive.Columns.Add("UpdateLevel", System.Type.GetType("System.Int32"))
              End If

              If (TryParseDateSent(cols(2), TempDate)) Then
                row = dtIncentive.NewRow()
                cols(0) = LocationID
                'row.ItemArray = cols
                row.Item("LocationID") = LocationID
                row.Item("IncentiveID") = cols(1)
                row.Item("DateValidated") = TempDate
                row.Item("UpdateLevel") = cols(3)
                dtIncentive.Rows.Add(row)
                incentiveKeys.Add(cols(1))
              Else
                If ErrMsg = "" Then
                  ErrMsg = "Incorrect date format found in offer validation record(s) " & vbCrLf & vbCrLf & _
                           "Process Info: Server Name=" & Environment.MachineName & " Invoking LocationID=" & LocationID & vbCrLf
                End If
                ErrMsg &= record
              End If

          End Select
        End If
      Next

      If Not (dtGraphic Is Nothing) Then WriteData(dtGraphic, "ValidGraphics", graphicKeys, LocationID)
      If Not (dtUserGroup Is Nothing) Then
        WriteData(dtUserGroup, "ValidCustomerGroups", userGroupKeys, LocationID)
        UpdateUserGroupCount(dtUserGroup, LocationID)
      End If
      If Not (dtProdGroup Is Nothing) Then
        WriteData(dtProdGroup, "ValidProductGroups", prodGroupKeys, LocationID)
        UpdateProdGroupCount(dtProdGroup, LocationID)
      End If
      If Not (dtIncentive Is Nothing) Then WriteData(dtIncentive, "ValidIncentives", incentiveKeys, LocationID)

      If ErrMsg <> "" Then
        responseText = ErrMsg
        Common.Error_Processor(ErrMsg)
        Common.Write_Log(LogFile, ErrMsg, True)
      End If

    Catch ex As Exception
      responseText = ex.ToString
      Common.Error_Processor()
      Common.Write_Log(LogFile, ex.ToString())
    Finally
    End Try
    Common.Write_Log(LogFile, "Processing Time (in ms): " & stopWatch.ElapsedMilliseconds)
    stopWatch.Stop()
    Return responseText
  End Function

  Sub UpdateProdGroupCount(ByVal dt As DataTable, ByVal LocationID As Integer)
    Dim dtProdGroups As DataTable
    Dim dtCount As DataTable
    Dim row As DataRow
    Dim ctTable As New Hashtable(500)
    Dim count As Integer = 0

    Try
      If Not (Common.LRTadoConn.State = ConnectionState.Open) Then Common.Open_LogixRT()

      Common.QueryStr = "select LocationID, ProductGroupID, CountActual from ValidProductGroups with (NoLock) " & _
                        "where CountActual is null and LocationID=" & LocationID & ";"
      dtProdGroups = Common.LRT_Select

      For Each row In dtProdGroups.Rows
        Common.QueryStr = "select count(*) as ItemCount from ProdGroupItems with (NoLock) " & _
                          "where ProductGroupID = " & Common.NZ(row.Item("ProductGroupID"), -1) & " and Deleted=0;"
        dtCount = Common.LRT_Select
        If (dtCount.Rows.Count > 0) Then
          count = Common.NZ(dtCount.Rows(0).Item("ItemCount"), 0)
        Else
          count = 0
        End If

        row.SetModified()
        row.Item("CountActual") = count
      Next

      BatchUpdatePG(dtProdGroups, dtProdGroups.Rows.Count)

    Catch ex As Exception
      Send(ex.ToString)
      Common.Error_Processor()
      Common.Write_Log(LogFile, ex.ToString())
    End Try

  End Sub

  Sub UpdateUserGroupCount(ByVal dt As DataTable, ByVal LocationID As Integer)
    Dim dtUsers As DataTable
    Dim dtCounts As DataTable
    Dim row As DataRow
    Dim ctTable As New Hashtable(500)
    Dim count As Integer = 0

    Try
      If Not (Common.LRTadoConn.State = ConnectionState.Open) Then Common.Open_LogixRT()
      If Not (Common.LXSadoConn.State = ConnectionState.Open) Then Common.Open_LogixXS()

      MacAddress = Trim(Request.QueryString("mac"))
      If MacAddress = "" Or MacAddress = "0" Then
        MacAddress = Trim(Request.UserHostAddress)
      End If
      LocalServerIP = Common.Extract_Val(Request.QueryString("IP"))
      If LocalServerIP = "" Or LocalServerIP = "0" Then
        LocalServerIP = MacAddress & " :IP from requesting browser. "
      End If

      Common.QueryStr = "select distinct GM.CustomerGroupID, count(*) as CustomerCount from GroupMembership GM with (NoLock) " & _
                        "inner join CustomerLocations CL with (NoLock) on CL.CustomerPK = GM.CustomerPK " & _
                        "and CL.LocationID=" & LocationID & " " & _
                        "where GM.Deleted = 0 Group by GM.CustomerGroupID;"
      dtCounts = Common.LXS_Select
      If (dtCounts.Rows.Count > 0) Then
        For Each row In dtCounts.Rows
          ctTable.Add(Common.NZ(row.Item("CustomerGroupID"), -1), _
                      Common.NZ(row.Item("CustomerCount"), 0))
        Next
      End If

      Common.QueryStr = "select LocationID, CustomerGroupID, CountActual from ValidCustomerGroups with (NoLock) " & _
                        "where CountActual is null and LocationID=" & LocationID & ";"
      dtUsers = Common.LRT_Select

      For Each row In dtUsers.Rows
        row.SetModified()
        If (ctTable.Contains(Common.NZ(row.Item("CustomerGroupID"), "-1"))) Then
          count = ctTable.Item(Common.NZ(row.Item("CustomerGroupID"), "-1"))
        Else
          count = 0
        End If
        row.Item("CountActual") = count
      Next

      BatchUpdate(dtUsers, dtUsers.Rows.Count)

    Catch ex As Exception
      Send(ex.ToString)
      Common.Error_Processor()
      Common.Write_Log(LogFile, "LocationID:" & LocationID & " Serial:" & LocalServerID & " MacAddress:" & MacAddress & " IP:" & LocalServerIP & " " & ex.ToString() & " server=" & Environment.MachineName)
    End Try
  End Sub

  Public Sub BatchUpdate(ByVal dataTable As DataTable, ByVal batchSize As Int32)
    Dim adapter As New SqlDataAdapter()

    'Set the UPDATE command and parameters.
    adapter.UpdateCommand = New SqlCommand( _
      "UPDATE ValidCustomerGroups with (RowLock) SET " _
      & "CountActual=@CountActual where LocationID=@LocationID and CustomerGroupID=@CustomerGroupID and CountActual is null;", _
      Common.LRTadoConn)
    adapter.UpdateCommand.Parameters.Add("@CountActual", _
      SqlDbType.BigInt, 8, "CountActual")
    adapter.UpdateCommand.Parameters.Add("@LocationID", _
      SqlDbType.BigInt, 8, "LocationID")
    adapter.UpdateCommand.Parameters.Add("@CustomerGroupID", _
      SqlDbType.BigInt, 8, "CustomerGroupID")
    adapter.UpdateCommand.UpdatedRowSource = UpdateRowSource.None

    ' Set the batch size.
    adapter.UpdateBatchSize = batchSize

    adapter.Update(dataTable)

  End Sub

  Sub BatchUpdatePG(ByVal dataTable As DataTable, ByVal batchSize As Int32)
    Dim adapter As New SqlDataAdapter()

    'Set the UPDATE command and parameters.
    adapter.UpdateCommand = New SqlCommand( _
      "UPDATE ValidProductGroups with (RowLock) SET " _
      & "CountActual=@CountActual where LocationID=@LocationID and ProductGroupID=@ProductGroupID and CountActual is null;", _
      Common.LRTadoConn)
    adapter.UpdateCommand.Parameters.Add("@CountActual", _
      SqlDbType.BigInt, 8, "CountActual")
    adapter.UpdateCommand.Parameters.Add("@LocationID", _
      SqlDbType.BigInt, 8, "LocationID")
    adapter.UpdateCommand.Parameters.Add("@ProductGroupID", _
      SqlDbType.BigInt, 8, "ProductGroupID")
    adapter.UpdateCommand.UpdatedRowSource = UpdateRowSource.None

    ' Set the batch size.
    adapter.UpdateBatchSize = batchSize

    adapter.Update(dataTable)

  End Sub

  Sub WriteData(ByVal dt As DataTable, ByVal TableName As String, ByVal Keys As ArrayList, ByVal LocationID As Integer)
    Dim bc As SqlBulkCopy = Nothing
    Dim keyList As New StringBuilder()

    Try
      If Not (Common.LRTadoConn.State = ConnectionState.Open) Then Common.Open_LogixRT()

      'Common.QueryStr = "Begin Transaction;"
      'Common.LRT_Execute()

      ' first delete pre-existing records matching the keys in the request data
      If Keys.Count > 0 Then
        Common.QueryStr = "delete from " & TableName & " with (RowLock) where  LocationID=" & LocationID & _
                          " and " & GetKeyFieldName(TableName) & " in (" & GetKeyList(Keys) & ");"
        Common.LRT_Execute()
      End If

      bc = New SqlBulkCopy(Common.LRTadoConn)
      bc.BatchSize = dt.Rows.Count
      bc.DestinationTableName = TableName
      bc.WriteToServer(dt)
      bc.Close()
      'Common.QueryStr = "Commit Transaction;"
      'Common.LRT_Execute()
    Catch ex As Exception
      'Common.QueryStr = "Rollback Transaction;"
      'Common.LRT_Execute()
      Send(ex.ToString)
      Common.Error_Processor()
      Common.Write_Log(LogFile, "LocationID:" & LocationID & " MacAddress:" & MacAddress & " IP:" & LocalServerIP & " " & ex.ToString() & " server:" & Environment.MachineName)
    End Try

  End Sub

  Function GetKeyFieldName(ByVal TableName As String) As String
    Dim keyFieldName As String = ""

    Select Case TableName
      Case "ValidGraphics"
        keyFieldName = "OnScreenAdID"
      Case "ValidCustomerGroups"
        keyFieldName = "CustomerGroupID"
      Case "ValidProductGroups"
        keyFieldName = "ProductGroupID"
      Case "ValidIncentives"
        keyFieldName = "IncentiveID"
    End Select

    Return keyFieldName
  End Function

  Function GetKeyList(ByVal keys As ArrayList) As String
    Dim keyList As New StringBuilder(300)
    Dim i As Integer = 0
    keyList.Append("-77")
    For i = 0 To keys.Count - 1
      keyList.Append(",")
      keyList.Append(keys(i).ToString)
    Next
    Return keyList.ToString
  End Function

  Private Function TryParseDateSent(ByVal dateString As String, ByRef RetDate As DateTime) As Boolean
    Dim Parsed As Boolean = True
    Dim Year, Month, Day As Integer
    Dim Hour, Minute, Second, Millisecond As Integer
    Dim UTCOffset As Integer

    Try
      ' Date should be in this format YYYY-MM-DD HH:mm:ss.fff(+/-)kk
      Year = Integer.Parse(Left(dateString, 4))
      Month = Integer.Parse(Mid(dateString, 6, 2))
      Day = Integer.Parse(Mid(dateString, 9, 2))
      Hour = Integer.Parse(Mid(dateString, 12, 2))
      Minute = Integer.Parse(Mid(dateString, 15, 2))
      Second = Integer.Parse(Mid(dateString, 18, 2))
      Millisecond = Common.Extract_Val(Mid(dateString, 21, 3))
      UTCOffset = Integer.Parse(Right(dateString, 3))

      RetDate = New Date(Year, Month, Day, Hour, Minute, Second, Millisecond)
      RetDate = RetDate.AddHours(-UTCOffset)
      RetDate = RetDate.ToLocalTime()
    Catch ex As Exception
      Parsed = False
      RetDate = Date.MinValue
    End Try

    Return Parsed
  End Function
</script>
<%
  Dim LocalServerID As Long
  Dim LocationID As Long
  Dim LastHeard As String
  Dim TotalTime As Decimal
  Dim ZipOutput As Boolean
  Dim Data As String = ""
  Dim Mode As String
  Dim BannerID As Integer

  Common.AppName = "OfferValidation.aspx"
  Response.Expires = 0

  On Error GoTo ErrorTrap

  MacAddress = Trim(Request.QueryString("mac"))

  If MacAddress = "" Or MacAddress = "0" Then
    MacAddress = Trim(Request.UserHostAddress)
  End If
  LocalServerID = Common.Extract_Val(Request.QueryString("serial"))
  LocalServerIP = Common.Extract_Val(Request.QueryString("IP"))
  If LocalServerIP = "" Or LocalServerIP = "0" Then
    LocalServerIP = MacAddress & " IP from requesting browser. "
  End If

  StartTime = DateAndTime.Timer

  LocalServerID = Common.Extract_Val(Request.QueryString("serial"))
  LogFile = "UE-OfferValidationLog-" & Format(LocalServerID, "00000") & "." & Common.Leading_Zero_Fill(Year(Today), 4) & Common.Leading_Zero_Fill(Month(Today), 2) & Common.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"

  LSVersion = Common.Extract_Val(Request.QueryString("lsversion"))
  LSBuild = Common.Extract_Val(Request.QueryString("lsbuild"))
  LastHeard = "1/1/1980"

  Common.Open_LogixRT()
  Common.Open_LogixXS()
  Common.Open_LogixWH()
  Common.Load_System_Info()
  Connector.Load_System_Info(Common)

  Common.Write_Log(LogFile, "---------------------------------------------------------------------------")
  Common.Write_Log(LogFile, "Processing Offer Validation  Process running on server:" & Environment.MachineName, True)

  ZipOutput = True
  Response.ContentType = "application/x-gzip"

  Connector.Get_LS_Info(Common, LocalServerID, LocationID, BannerID, LastHeard, 1, Request.UserHostAddress)

  DebugMode = (Request.QueryString("debug") = "1")
  If (DebugMode) Then LocationID = 107

  If LocationID = "0" Then
    Common.Write_Log(LogFile, Common.AppName & "    Invalid Serial Number:" & LocalServerID & " from Serial:" & LocalServerID & " MacAddress: " & MacAddress & " IP:" & LocalServerIP & "  Process running on server:" & Environment.MachineName, True)
    Send_Response_Header("OfferValidation - Invalid Serial Number", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
    Send("<br />Invalid Serial Number from IP: " & Trim(Request.UserHostAddress))
  Else
    Mode = UCase(Request.QueryString("mode"))
    Common.Write_Log(LogFile, "** " & Microsoft.VisualBasic.DateAndTime.Now & "  Serial: " & LocalServerID & " MacAddress:" & MacAddress & " IP:" & LocalServerIP & "   LocationID: " & LocationID & "  Mode: " & Mode & " server:" & Environment.MachineName)
    If (Mode = "0" OrElse DebugMode) Then
      If (Handle_Post(LocalServerID, LocationID, Data)) Then
        Dim responseText As String = ProcessData(Data, LocationID)
        If (responseText = "") Then
          Send("ACK")
        Else
          Send(responseText)
        End If
      Else
        Send(Data)
      End If
    Else
      Send_Response_Header("Invalid Request", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
      Send("Invalid Request")
    End If
  End If 'locationid="0"

  Common.Close_LogixRT()
  Common.Close_LogixXS()
  Common.Close_LogixWH()

  TotalTime = DateAndTime.Timer - StartTime
  Common.Write_Log(LogFile, "RunTime=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")

  Response.End()
ErrorTrap:
  Response.Write(Common.Error_Processor(, , "Serial:" & LocalServerID & " MacAddress:" & MacAddress & " LocalServerIP:" & LocalServerIP, Common.InstallationName))
  Common.Write_Log(LogFile, "***** An error occurred during processing! Please check the ErrorLog! *****")
  If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
  If Not (Common.LXSadoConn.State = ConnectionState.Closed) Then Common.Close_LogixXS()
  If Not (Common.LWHadoConn.State = ConnectionState.Closed) Then Common.Close_LogixWH()
  Common = Nothing
%>