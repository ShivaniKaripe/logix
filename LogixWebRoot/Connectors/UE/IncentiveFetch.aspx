<%@ Page CodeFile="..\ConnectorCB.vb" Debug="true" Inherits="ConnectorCB" Language="vb" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: IncentiveFetch.aspx
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
  Public OutboundBuffer As StringBuilder
  Public TextData As String
  Public LogFile As String
  Public FileLocation As String
  Public RequireLocations As Boolean = False

  ' -----------------------------------------------------------------------------------------------

  Sub SD(ByVal OutStr As String)
    OutboundBuffer.Append(OutStr & vbCrLf)
  End Sub

  ' -----------------------------------------------------------------------------------------------

  Sub SDb(ByVal OutStr As String)
    OutboundBuffer.Append(OutStr)
  End Sub

  ' -----------------------------------------------------------------------------------------------

  Function Parse_Bit(ByVal BooleanField As Boolean) As String
    If BooleanField Then
      Parse_Bit = "1"
    Else
      Parse_Bit = "0"
    End If
  End Function

  ' -----------------------------------------------------------------------------------------------

  Sub Construct_Table(ByVal TableName As String, ByVal Operation As String, ByVal DelimChar As String, ByVal LocalServerID As Long, ByVal LocationID As Long, ByVal dst As DataTable)

    Dim TempResults As String
    Dim NumRecs As Long
    Dim row As DataRow
    Dim SQLCol As DataColumn
    Dim TempOut As String
    Dim Index As Integer
    Dim FieldList As String
    Dim FileName As String = ""
    Dim RowsSent As Long

    TempOut = ""
    TempResults = ""
    NumRecs = 0
    FieldList = ""
    If dst.Rows.Count > 0 Then
      For Each SQLCol In dst.Columns
        If Not (FieldList = "") Then FieldList = FieldList & DelimChar
        FieldList = FieldList & SQLCol.ColumnName
      Next
      TempOut = "1:" & TableName & vbCrLf
      TempOut = TempOut & "2:" & Operation & vbCrLf
      TempOut = TempOut & "3:" & FieldList
      SD(TempOut)
      Common.Write_Log(LogFile, TempOut)

      RowsSent = 0
      For Each row In dst.Rows
        Index = 0
        TempResults = ""
        For Each SQLCol In dst.Columns
          If TableName = "ListOfFiles" And Index = 0 Then
            FileName = Common.NZ(row(Index), "")
          End If
          If Not (TempResults = "") Then
            TempResults = TempResults & DelimChar
          End If
          If SQLCol.DataType.Name = "Boolean" Then 'if it is a binary field
            TempResults = TempResults & Parse_Bit(Common.NZ(row(Index), 0))
          ElseIf SQLCol.DataType.Name = "Int32" Or SQLCol.DataType.Name = "Int64" Then 'if it is an Int or BigInt Field
            TempResults = TempResults & Common.NZ(row(Index), 0)
          Else 'else treat it as a string
            TempResults = TempResults & Common.NZ(row(Index), "")
          End If
          Index = Index + 1
        Next
        SD(TempResults)
        If TableName = "ListOfFiles" Then
          Common.Write_Log(LogFile, "<a href=""" & FileLocation & FileName & """ target=""_blank"">" & TempResults & "</a>")
        Else
          Common.Write_Log(LogFile, TempResults)
        End If
        RowsSent = RowsSent + 1
      Next
      SD("###")
      Common.Write_Log(LogFile, "###")
      Common.Write_Log(LogFile, "Rows Sent=" & RowsSent)
    End If

  End Sub

  ' -----------------------------------------------------------------------------------------------

  Sub Construct_Output(ByVal LocalServerID As Long, ByVal LocationID As Long, ByVal LastHeard As String)

    Dim OutStr As String
    Dim DelimChar As String
    Dim MustIPL As Boolean
    Dim IncentiveFetchURL As String
    Dim ImageFetchURL As String
    Dim dst As DataTable
    Dim MaxBatchSize As Integer
    Dim OperateAtEnterprise As Boolean = False
    Dim dtAffectedLocations As New DataTable()    
    Dim incentiveIdsNewOrChanged As New DataTable()
    Dim incentiveIdsDeleted As New DataTable()

    incentiveIdsNewOrChanged.Columns.Add("LongColumn")
    incentiveIdsDeleted.Columns.Add("LongColumn")
    OutStr = ""
    DelimChar = Chr(30)

    Common.Write_Log(LogFile, "Returned the following data:")

    Common.QueryStr = "dbo.pa_UE_CheckMustIPL"
    Common.Open_LRTsp()
    Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
    dst = Common.LRTsp_select
    Common.Close_LRTsp()
    If dst.Rows.Count > 0 Then
      MustIPL = Common.NZ(dst.Rows(0).Item("MustIPL"), True)
    End If
    dst = Nothing

    If MustIPL Then
      OutStr = "MustIPL"
    Else
      OutStr = "IncentiveFetch"
    End If

    OutStr = OutStr & "," & Connector.CSMajorVersion & "," & Connector.CSMinorVersion & "," & Connector.CSBuild & "," & Connector.CSBuildRevision & vbCrLf
    OutStr = OutStr & "LocationID=" & LocationID
    SD(OutStr)
    Common.Write_Log(LogFile, OutStr)

    If Not (MustIPL) Then

      Common.QueryStr = "dbo.pa_UE_CheckEnterprise"
      Common.Open_LRTsp()
      Common.LRTsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = LocationID
      Common.LRTsp.Parameters.Add("@Enterprise", SqlDbType.Bit).Direction = ParameterDirection.Output
      Common.LRTsp.ExecuteNonQuery()
      OperateAtEnterprise = Common.NZ(Common.LRTsp.Parameters("@Enterprise").Value, False)
      Common.Close_LRTsp()

      'see if the IncentiveFetchURL is set for this local server - if not, update it from the Clients table
            'AMS-1597: Updating incentive fetch url with ue systemoption 43      
            IncentiveFetchURL = Common.Fetch_UE_SystemOption(43)
            ImageFetchURL = "NotSpecified"
            Common.QueryStr = "dbo.pa_UE_CheckIFURL"
            Common.Open_LRTsp()
            Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
            dst = Common.LRTsp_select
            Common.Close_LRTsp()
            If dst.Rows.Count > 0 Then
                'IncentiveFetchURL = Common.NZ(dst.Rows(0).Item("IncentiveFetchURL"), "NotSpecified")
                ImageFetchURL = Common.NZ(dst.Rows(0).Item("ImageFetchURL"), "NotSpecified")
            End If
            dst = Nothing
            If Trim(IncentiveFetchURL) = "" Or IncentiveFetchURL = "NotSpecified" Then
                Common.QueryStr = "dbo.pa_UE_UpdateIncentFURL"
                Common.Open_LRTsp()
                Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
                Common.LRTsp.ExecuteNonQuery()
                Common.Close_LRTsp()
            End If
            If Trim(ImageFetchURL) = "" Or ImageFetchURL = "NotSpecified" Then
                Common.QueryStr = "dbo.pa_UE_UpdateImgFURL"
                Common.Open_LRTsp()
                Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
                Common.LRTsp.ExecuteNonQuery()
                Common.Close_LRTsp()
            End If
            FileLocation = Trim(IncentiveFetchURL)
            If Not (Right(FileLocation, 1) = "\") Then
                FileLocation = FileLocation & "\"
            End If

            'LocalServers
            'send the data from the LocalServers table
            Common.QueryStr = "dbo.pa_UE_IF_LocalServers"
            Common.Open_LRTsp()
            Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
            dst = Common.LRTsp_select
            Common.Close_LRTsp()
            Construct_Table("LocalServers", 1, DelimChar, LocalServerID, LocationID, dst)

            'Locations (retail locations)
            Common.QueryStr = "dbo.pa_UE_IF_LocationsActive"
            Common.Open_LRTsp()
            Common.LRTsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
            Common.LRTsp.Parameters.Add("@Lastheard", SqlDbType.DateTime).Value = LastHeard
            dst = Common.LRTsp_select
            Common.Close_LRTsp()
            Construct_Table("Locations", 1, DelimChar, LocalServerID, LocationID, dst)

            'send any newly deleted LocationLanguages records
            If Not (OperateAtEnterprise) Then

                Common.QueryStr = "select PKID from LocationLanguages with (NoLock) where Deleted=1 and LocationID=@LocationID  and LastUpdate>=@LastUpdate "
                Common.DBParameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
                Common.DBParameters.Add("@LastUpdate", SqlDbType.DateTime).Value = LastHeard

            Else
                Common.QueryStr = "select LL.PKID " & _
                          "from LocationLanguages as LL with (NoLock) Inner join Locations as L with (NoLock) on L.LocationID=LL.LocationID " & _
                          "where L.EngineID=9 and LL.Deleted=1 and LL.LastUpdate>=@LastUpdate "

                Common.DBParameters.Add("@LastUpdate", SqlDbType.DateTime).Value = LastHeard
            End If
            dst = Common.ExecuteQuery(Copient.DataBases.LogixRT)
            Construct_Table("LocationLanguages", 2, DelimChar, LocalServerID, LocationID, dst)

            'send any new or updated LocationLanguages records
            If Not (OperateAtEnterprise) Then
                Common.QueryStr = "select PKID, LocationID, LanguageID, Required from LocationLanguages with (NoLock) where Deleted=0 and LocationID=@LocationID  and LastUpdate>=@LastUpdate "
                Common.DBParameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
                Common.DBParameters.Add("@LastUpdate", SqlDbType.DateTime).Value = LastHeard

            Else
                Common.QueryStr = "select LL.PKID, LL.LocationID, LL.LanguageID, LL.Required " & _
                                "from LocationLanguages as LL with (NoLock) Inner join Locations as L with (NoLock) on L.LocationID=LL.LocationID " & _
                                "where L.EngineID=9 and LL.Deleted=0 and LL.LastUpdate>=@LastUpdate "
                Common.DBParameters.Add("@LastUpdate", SqlDbType.DateTime).Value = LastHeard

            End If
            dst = Common.ExecuteQuery(Copient.DataBases.LogixRT)
            Construct_Table("LocationLanguages", 1, DelimChar, LocalServerID, LocationID, dst)

            MaxBatchSize = Common.Extract_Val(Common.Fetch_UE_SystemOption(119))
            Common.QueryStr = "dbo.pa_UE_IF_GetFileList"
            Common.Open_LRTsp()
            Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
            Common.LRTsp.Parameters.Add("@MaxBatch", SqlDbType.Int).Value = MaxBatchSize
            dst = Common.LRTsp_select
            Common.Close_LRTsp()
            Dim tempOffersDst As DataTable = dst
            Construct_Table("ListOfFiles", 9, DelimChar, LocalServerID, LocationID, dst)
      
            If OperateAtEnterprise AndAlso RequireLocations Then
                GetIncentiveIdsBasedOnFileTypes(tempOffersDst, incentiveIdsNewOrChanged, incentiveIdsDeleted)

                dtAffectedLocations = GetAffectedLocationsForNewOrChangedOffers(Common, incentiveIdsNewOrChanged)
                dtAffectedLocations.Merge(GetAffectedLocationsForDeletedOffers(Common, incentiveIdsDeleted))
                dtAffectedLocations.Merge(GetOfferLocationsChangedPostDeploy(Common, incentiveIdsNewOrChanged))
                'dtAffectedLocations.Merge(GetOfferLocationsRemovedPostDeploy(Common, incentiveIdsForDeletedLocations))
                If dtAffectedLocations.Rows.Count > 0 Then
                    Construct_Table("FetchAffectedLocations", 5, DelimChar, LocalServerID, LocationID, dtAffectedLocations)
                End If
            End If
        End If

        SD("***") 'send the EOF marker
    End Sub
    Function GetOfferLocationsChangedPostDeploy(ByRef Common As Copient.CommonInc, ByRef incentiveIdsToIgnore As DataTable) As DataTable
        Dim dst As New DataTable()

        Common.Write_Log(LogFile, "Sending new locations for deployed offers")
        Common.QueryStr = "pt_GetIncentive_ChangedLocations_PostDeploy"
        Common.Open_LRTsp()
        Common.LRTsp.Parameters.Add("@ExceptIncentiveIdTT", SqlDbType.Structured).Value = incentiveIdsToIgnore
                    
        dst = Common.LRTsp_select()
        'If dst.Rows.Count > 0 Then
        '    dtAffectedLocations.Merge(dst)
        'End If
        Common.Close_LogixRT()
        
        Return dst
    End Function

    Function GetAffectedLocationsForDeletedOffers(ByRef Common As Copient.CommonInc, ByRef incentiveIdsDeleted As DataTable) As DataTable
        Dim dst As New DataTable()
        If incentiveIdsDeleted.Rows.Count > 0 Then
            Common.Write_Log(LogFile, "Sending affected locations for deleted offers as well along with incentive files")
            Common.QueryStr = "pt_Get_DeletedIncentiveLocations"
            Common.Open_LRTsp()
            Common.LRTsp.Parameters.Add("@IncentiveIdTT", SqlDbType.Structured).Value = incentiveIdsDeleted
                    
            dst = Common.LRTsp_select()
            'If dst.Rows.Count > 0 Then
            '    dtAffectedLocations.Merge(dst)
            'End If
            Common.Close_LogixRT()
        End If
        
        Return dst
    End Function
    Function GetAffectedLocationsForNewOrChangedOffers(ByRef Common As Copient.CommonInc, ByRef incentiveIdsNewOrChanged As DataTable) As DataTable
        Dim dst As New DataTable()
        If incentiveIdsNewOrChanged.Rows.Count > 0 Then
            Common.Write_Log(LogFile, "Sending affected locations for new or modified offers as well along with incentive files")
            'Common.QueryStr = "SELECT distinct LocationID FROM UE_IncentiveAllLocationsView WHERE IncentiveID in (" & incentiveIdsNewOrChanged.ToString() & ")"
            Common.QueryStr = "pt_GetIncentive_ChangedLocations"
            Common.Open_LRTsp()
            Common.LRTsp.Parameters.Add("@IncentiveIdTT", SqlDbType.Structured).Value = incentiveIdsNewOrChanged
                    
            dst = Common.LRTsp_select()
            'Construct_Table("FetchAffectedLocations", 5, DelimChar, LocalServerID, LocationID, dst)
            'If (dst.Rows.Count > 0) Then
            '    dtAffectedLocations.Merge(dst)
            'End If
            Common.Close_LRTsp()
        End If        
        Return dst
    End Function
    Sub GetIncentiveIdsBasedOnFileTypes(ByRef tempOffersDst As DataTable, ByRef incentiveIdsNewOrChanged As DataTable, ByRef incentiveIdsDeleted As DataTable)
        For Each row As DataRow In tempOffersDst.Rows
            Dim data As String() = row(0).ToString().Split(New Char() {"-"c})
            
            Select Case data(1)
                Case "A"
                    incentiveIdsNewOrChanged.NewRow()
                    incentiveIdsNewOrChanged.Rows.Add(data(2))
                    Exit Select
                Case "OD"
                    incentiveIdsDeleted.NewRow()
                    incentiveIdsDeleted.Rows.Add(data(2))
                    Exit Select
                Case "A-RD"
                    Exit Select
                Case "D-PD"
                    Exit Select
            End Select
        Next
    End Sub
    ' -----------------------------------------------------------------------------------------------

    Sub Process_ACK(ByVal LocalServerID As Integer, ByVal LocationID As Integer)

        Dim FileList As String
        'Dim FileArray() As String
        'Dim Index As Integer

        Common.Write_Log(LogFile, "Received IncentiveFetch ACK")
        'FileList = Request.QueryString("files")
        FileList = GetCgiValue("files")
        'FileArray = Split(FileList, ",")
        ''put quotes around all of the files in the list and then rebuild the list
        'FileList = ""
        'For Index = 0 To UBound(FileArray)
        '  FileList = FileList & "'" & Trim(FileArray(Index)) & "'"
        '  If Not (Index = UBound(FileArray)) Then
        '    FileList = FileList & ", "
        '  End If
        'Next

        Common.QueryStr = "Update LocalServers with (RowLock) set IncentiveLastHeard=getdate(), Lastheard=getdate() where LocalServerID=@LocalServerID "
        Common.DBParameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID

        Common.ExecuteNonQuery(Copient.DataBases.LogixRT)

        If Not (FileList = "") Then
            Common.QueryStr = "update UE_IncentiveDLBuffer with (RowLock) set WaitingACK=2 where WaitingACK=1 and LocalServerID=@LocalServerID  and FileName in (select items  from dbo.Split(@FileName,',') )"
            Common.DBParameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
            Common.DBParameters.Add("@FileName", SqlDbType.VarChar).Value = FileList

            Common.ExecuteNonQuery(Copient.DataBases.LogixRT)

        End If
        If FileList = "" Then FileList = "None"
        Common.Write_Log(LogFile, "Received ACK for files: " & FileList)

        Send_Response_Header("IncentiveFetch", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
        Send("ACK Received")

    End Sub

    ' -----------------------------------------------------------------------------------------------

    Sub Process_NAK(ByVal LocalServerID As Integer, ByVal LocationID As Integer)

        Dim ErrorMsg As String

        ErrorMsg = Trim(Request.QueryString("errormsg"))
        Common.Write_Log(LogFile, "Received NAK - ErrorMsg:" & ErrorMsg)
        Send_Response_Header("NAK Received", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)

    End Sub
</script>
<%
  '-----------------------------------------------------------------------------------------------
  'Main Code - Execution starts here

  Dim LocalServerID As Long
  Dim LocationID As Long
  Dim LastHeard As String
  Dim StartTime As Object
  Dim TotalTime As Object
  Dim Mode As String
  Dim RawRequest As String
  Dim IPAddress As String = ""
  Dim CompressedArray() As Byte
  Dim BannerID As Integer
  Dim LocalServerIP As String
  Dim MacAddress As String

  Common.AppName = "IncentiveFetch.aspx"
  Response.Expires = 0
  On Error GoTo ErrorTrap
  StartTime = DateAndTime.Timer
  IPAddress = Request.UserHostAddress

  LastHeard = "1/1/1980"
  LocalServerID = Common.Extract_Val(Request.QueryString("serial"))
  LSVersion = Common.Extract_Val(Request.QueryString("lsversion"))
  LSBuild = Common.Extract_Val(Request.QueryString("lsbuild"))
  Mode = UCase(Request.QueryString("mode"))
  LogFile = "UE-IncentiveFetchLog-" & Format(LocalServerID, "00000") & "." & Common.Leading_Zero_Fill(Year(Today), 4) & Common.Leading_Zero_Fill(Month(Today), 2) & Common.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"

  LocalServerIP = Trim(Request.QueryString("IP"))
  RequireLocations = Common.NZ(Request.QueryString("RequireLocations"), False)
  If LocalServerIP = "" Then LocalServerIP = Trim(Request.QueryString("ip"))
  If LocalServerIP = "" Then
    LocalServerIP = Trim(Request.UserHostAddress) & "(UserHostAddress)"
  End If

  MacAddress = Trim(Request.QueryString("mac"))
  If MacAddress = "" Then
    MacAddress = "Not Specified"
  End If

  Common.Open_LogixRT()
  Common.Load_System_Info()
  Connector.Load_System_Info(Common)
  Connector.Get_LS_Info(Common, LocalServerID, LocationID, BannerID, LastHeard, 1, IPAddress)

  Response.ContentType = "application/x-gzip"

  Common.Write_Log(LogFile, "----------------------------------------------------------------")
  Common.Write_Log(LogFile, "** " & Microsoft.VisualBasic.DateAndTime.Now & "  Serial: " & LocalServerID & "   LocationID: " & LocationID & "  Mode: " & Mode & "  LSVersion=" & LSVersion & "  LSBuild=" & LSBuild & "  Process running on server:" & Environment.MachineName & "  Requester IP Address: " & LocalServerIP & "   Requester Mac Address: " & MacAddress)

  If LocationID = 0 Then
    Common.Write_Log(LogFile, "Received Invalid Serial Number from IP: " & IPAddress)
    Send_Response_Header("Invalid Serial Number", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
  ElseIf Not (Connector.IsValidLocationForConnectorEngine(Common, LocationID, 9)) Then
    'the location calling TransDownload is not associated with the UE promoengine
    Common.Write_Log(LogFile, "This location is associated with a promotion engine other than UE.  Can not proceed.", True)
    Send_Response_Header("This location is associated with a promotion engine other than UE.  Can not proceed.", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
  Else
    Select Case Mode
      Case "ACK"
        Process_ACK(LocalServerID, LocationID)
      Case "NAK"
        Process_NAK(LocalServerID, LocationID)
      Case "FETCH"
        OutboundBuffer = New StringBuilder
        Construct_Output(LocalServerID, LocationID, LastHeard)
        Common.Write_Log(LogFile, "Starting GZip compression ... size before zipping is " & Format(OutboundBuffer.Length, "###,###,###,###,##0") & " bytes")
        TotalTime = DateAndTime.Timer - StartTime
        Common.Write_Log(LogFile, "Time elapsed before starting compression=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
        CompressedArray = Encoding.Default.GetBytes(GZIP.CompressString(OutboundBuffer.ToString))
        Response.BinaryWrite(CompressedArray)
        Common.Write_Log(LogFile, "GZip compression successful ... size after zipping is " & Format(UBound(CompressedArray) + 1, "###,###,###,###,##0") & " bytes")
        TotalTime = DateAndTime.Timer - StartTime
        Common.Write_Log(LogFile, "Time elapsed after finishing compression=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
      Case Else
        Send_Response_Header("Invalid Request - bad mode", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
        Common.Write_Log(LogFile, "Received invalid request!")
        Common.Write_Log(LogFile, "Query String: " & Request.RawUrl & vbCrLf & "Form Data: " & vbCrLf)
        RawRequest = Get_Raw_Form(Request.InputStream)
        Common.Write_Log(LogFile, RawRequest)
    End Select
  End If 'locationid="0"

  If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
%>
<%Response.End()
ErrorTrap:
  Response.Write(Common.Error_Processor(, "Process Info: Server Name=" & Environment.MachineName & " Invoking LocalServerID: " & LocalServerID & "  LocationID=" & LocationID & "  Requester IP Address: " & LocalServerIP & "  Requester Mac Address: " & MacAddress, , Common.InstallationName))
  Common.Write_Log(LogFile, "***** An error occurred during processing! Please check the ErrorLog! *****")
  If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
  Common = Nothing
%>