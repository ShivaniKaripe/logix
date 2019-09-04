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
    
    Dim OutStr As String = String.Empty
    Dim DelimChar As String
    Dim UpdateTime As String = ""
    Dim MustIPL As Boolean
    Dim IncentiveFetchOffline As Boolean
    Dim IncentiveFetchURL As String
    Dim ImageFetchURL As String
    Dim dst As DataTable
    Dim MaxBatchSize As Integer
    Dim OperateAtEnterprise As Boolean = False
    
    OutStr = ""
    DelimChar = Chr(30)
    If Common.Fetch_CPE_SystemOption(91) = "1" Then OperateAtEnterprise = True
    
    Common.Write_Log(LogFile, "Returned the following data:")
  
    Common.QueryStr = "dbo.pa_CPE_CheckMustIPL"
    Common.Open_LRTsp()
    Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
    dst = Common.LRTsp_select
    Common.Close_LRTsp()
    If dst.Rows.Count > 0 Then
      MustIPL = Common.NZ(dst.Rows(0).Item("MustIPL"), True)
    End If
    dst = Nothing
    
    If MustIPL Then OutStr = "MustIPL"
    
    Common.QueryStr = "dbo.pa_CPE_CheckIncentiveFetchOffline"
    Common.Open_LRTsp()
    Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
    dst = Common.LRTsp_select
    
    If dst.Rows.Count > 0 Then
      IncentiveFetchOffline = Common.NZ(dst.Rows(0).Item("IncentiveFetchOffline"), True)
    End If
    dst = Nothing

    If (IncentiveFetchOffline) Then OutStr = "IncentiveFetchOffline"
    
    If (String.IsNullOrEmpty(OutStr)) Then OutStr = "IncentiveFetch"
    
    OutStr = OutStr & "," & Connector.CSMajorVersion & "," & Connector.CSMinorVersion & "," & Connector.CSBuild & "," & Connector.CSBuildRevision & vbCrLf
    OutStr = OutStr & "LocationID=" & LocationID
    SD(OutStr)
    Common.Write_Log(LogFile, OutStr)
    
    If Not (MustIPL Or IncentiveFetchOffline) Then
      'see if the IncentiveFetchURL is set for this local server - if not, update it from the Clients table
      IncentiveFetchURL = "NotSpecified"
      ImageFetchURL = "NotSpecified"
      Common.QueryStr = "dbo.pa_CPE_CheckIFURL"
      Common.Open_LRTsp()
      Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
      dst = Common.LRTsp_select
      Common.Close_LRTsp()
      If dst.Rows.Count > 0 Then
        IncentiveFetchURL = Common.NZ(dst.Rows(0).Item("IncentiveFetchURL"), "NotSpecified")
        ImageFetchURL = Common.NZ(dst.Rows(0).Item("ImageFetchURL"), "NotSpecified")
      End If
      dst = Nothing
      If Trim(IncentiveFetchURL) = "" Or IncentiveFetchURL = "NotSpecified" Then
        Common.QueryStr = "dbo.pa_CPE_UpdateIncentFURL"
        Common.Open_LRTsp()
        Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
        Common.LRTsp.ExecuteNonQuery()
        Common.Close_LRTsp()
      End If
      If Trim(ImageFetchURL) = "" Or ImageFetchURL = "NotSpecified" Then
        Common.QueryStr = "dbo.pa_CPE_UpdateImgFURL"
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
      Common.QueryStr = "dbo.pa_CPE_IF_LocalServers"
      Common.Open_LRTsp()
      Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
      dst = Common.LRTsp_select
      Common.Close_LRTsp()
      Construct_Table("LocalServers", 1, DelimChar, LocalServerID, LocationID, dst)
      
      'Locations (retail locations)
      Common.QueryStr = "dbo.pa_CPE_IF_LocationsActive"
      Common.Open_LRTsp()
      Common.LRTsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
      Common.LRTsp.Parameters.Add("@Lastheard", SqlDbType.DateTime).Value = LastHeard
      dst = Common.LRTsp_select
      Common.Close_LRTsp()
      Construct_Table("Locations", 1, DelimChar, LocalServerID, LocationID, dst)
      'deleted locations records are sent from IncentiveAgent to all Local Servers
      
      'send any newly deleted LocationLanguages records
      If Not (OperateAtEnterprise) Then
        Common.QueryStr = "select PKID from LocationLanguages with (NoLock) where Deleted=1 and LocationID=" & LocationID & " and LastUpdate>='" & LastHeard & "';"
      Else
        Common.QueryStr = "select LL.PKID " & _
                          "from LocationLanguages as LL with (NoLock) Inner join Locations as L with (NoLock) on L.LocationID=LL.LocationID " & _
                          "where L.EngineID=2 and LL.Deleted=1 and LL.LastUpdate>='" & LastHeard & "';"
      End If
      dst = Common.LRT_Select
      Construct_Table("LocationLanguages", 2, DelimChar, LocalServerID, LocationID, dst)
      
      'send any new or updated LocationLanguages records
      If Not (OperateAtEnterprise) Then
        Common.QueryStr = "select PKID, LocationID, LanguageID, Required from LocationLanguages with (NoLock) where Deleted=0 and LocationID=" & LocationID & " and LastUpdate>='" & LastHeard & "';"
      Else
        Common.QueryStr = "select LL.PKID, LL.LocationID, LL.LanguageID, LL.Required " & _
                          "from LocationLanguages as LL with (NoLock) Inner join Locations as L with (NoLock) on L.LocationID=LL.LocationID " & _
                          "where L.EngineID=2 and LL.Deleted=0 and LL.LastUpdate>='" & LastHeard & "';"
      End If
      dst = Common.LRT_Select
      Construct_Table("LocationLanguages", 1, DelimChar, LocalServerID, LocationID, dst)

      
      MaxBatchSize = Common.Extract_Val(Common.Fetch_CPE_SystemOption(119))
      Common.QueryStr = "dbo.pa_CPE_IF_GetFileList"
      Common.Open_LRTsp()
      Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
      Common.LRTsp.Parameters.Add("@MaxBatch", SqlDbType.Int).Value = MaxBatchSize
      dst = Common.LRTsp_select
      Common.Close_LRTsp()
      Construct_Table("ListOfFiles", 9, DelimChar, LocalServerID, LocationID, dst)
    End If
    
      SD("***") 'send the EOF marker
  End Sub
  
  ' -----------------------------------------------------------------------------------------------
  
  Sub Process_ACK(ByVal LocalServerID As Integer, ByVal LocationID As Integer)
    
    Dim FileList As String
    Dim FileArray() As String
    Dim Index As Integer
    
    Common.Write_Log(LogFile, "Received IncentiveFetch ACK")
    'FileList = Request.QueryString("files")
    FileList = Request.Form("files")
    FileArray = Split(FileList, ",")
    'put quotes around all of the files in the list and then rebuild the list
    FileList = ""
    For Index = 0 To UBound(FileArray)
      FileList = FileList & "'" & Trim(FileArray(Index)) & "'"
      If Not (Index = UBound(FileArray)) Then
        FileList = FileList & ", "
      End If
    Next
    Common.QueryStr = "Update LocalServers with (RowLock) set IncentiveLastHeard=getdate(), Lastheard=getdate() where LocalServerID=" & LocalServerID & ";"
    Common.LRT_Execute()
    If Not (FileList = "") Then
      Common.QueryStr = "update CPE_IncentiveDLBuffer with (RowLock) set WaitingACK=2 where WaitingACK=1 and LocalServerID=" & LocalServerID & " and FileName in (" & FileList & ");"
      Common.LRT_Execute()
    End If
    If FileList = "" Then FileList = "None"
    Common.Write_Log(LogFile, "Received ACK for files: " & FileList)

    '  reset incentiveoffline flag  
    Common.QueryStr = "If Not Exists (Select PKID from CPE_IncentiveDLBuffer where WaitingACK = 1 and LocalServerID = " & LocalServerID &  ")" & _
                      "Begin " & _
                      "   Update LocalServers set incentivefetchoffline=0, incentivefetchnakcount=0 Where LocalServerID = " & LocalServerID & " " & _
                      "End"
    Common.LRT_Execute()
    
    Send_Response_Header("IncentiveFetch", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
    Send("ACK Received")
    
  End Sub
  
  ' -----------------------------------------------------------------------------------------------
  
  Sub Process_NAK(ByVal LocalServerID As Integer, ByVal LocationID As Integer)
    
    Dim ErrorMsg As String
        Dim ExtLocationCode As String = ""
        Dim LocationName As String = ""
        Dim rst As New DataTable
    
    ErrorMsg = Trim(Request.QueryString("errormsg"))
    Common.Write_Log(LogFile, "Received NAK - ErrorMsg:" & ErrorMsg)
    Send_Response_Header("NAK Received", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
    
        If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()
        Common.QueryStr = "select LocationName, ExtLocationCode from Locations with (NoLock) where LocationID=" & LocationID.ToString & ";"
        rst = Common.LRT_Select
        If rst.Rows.Count > 0 Then
            LocationName = Common.NZ(rst.Rows(0).Item("LocationName"), "")
            ExtLocationCode = Common.NZ(rst.Rows(0).Item("ExtLocationCode"), "0")
        End If

        Dim OutBuffer As String = "Local Server Incentive Fetch NAK Received" & vbCrLf
        OutBuffer = OutBuffer & "LocationID: " & LocationID.ToString() & vbCrLf
        OutBuffer = OutBuffer & "Location: " & ExtLocationCode & " " & LocationName & vbCrLf
        OutBuffer = OutBuffer & "LocalServerSerial: " & LocalServerID.ToString() & vbCrLf
        OutBuffer = OutBuffer & "ErrorMsg " & ErrorMsg & vbCrLf
        OutBuffer = OutBuffer & vbCrLf & "Subject: Incentive Fetch NAK Received"

        Common.Send_Email(Common.Get_Error_Emails(5), Common.SystemEmailAddress, "Incentive Fetch NAK Received Message - " & Common.InstallationName & "  Store #" & ExtLocationCode, OutBuffer)
    
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
    Dim OutBuffer As String = ""
    Dim LocationName As String = ""
    Dim ExtLocationCode As String = ""
    Dim rst As New DataTable
  
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
  LogFile = "CPE-IncentiveFetchLog-" & Format(LocalServerID, "00000") & "." & Common.Leading_Zero_Fill(Year(Today), 4) & Common.Leading_Zero_Fill(Month(Today), 2) & Common.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
  
  LocalServerIP = Trim(Request.QueryString("IP"))
  If LocalServerIP = "" Then LocalServerIP = Trim(Request.QueryString("ip"))
  If LocalServerIP = "" Then
    Common.Write_Log(LogFile, "Could not get IP from query. Analyzing client request for IP ...")
    LocalServerIP = Trim(Request.UserHostAddress)
  End If

  MacAddress = Trim(Request.QueryString("mac"))
  If MacAddress = "" Then
    MacAddress = "0"
  End If

  
  Common.Open_LogixRT()
  Common.Load_System_Info()
  Connector.Load_System_Info(Common)
  Connector.Get_LS_Info(Common, LocalServerID, LocationID, BannerID, LastHeard, 1, LocalServerIP)
  
  Common.Write_Log(LogFile, "----------------------------------------------------------------")
  Common.Write_Log(LogFile, "** " & Microsoft.VisualBasic.DateAndTime.Now & "  Serial: " & LocalServerID & "   LocationID: " & LocationID & "  Mode: " & Mode & "  LSVersion=" & LSVersion & "  LSBuild=" & LSBuild & "  Process running on server:" & Environment.MachineName & "  Requester IP Address: " & LocalServerIP & "   Requester Mac Address: " & MacAddress)
  
  If LocationID = 0 Then
    Common.Write_Log(LogFile, "Received Invalid Serial Number from IP: " & IPAddress)
    Send_Response_Header("Invalid Serial Number", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
        
        OutBuffer = "Incentie Fetch Received Invalid Serial from MacAddress:" & MacAddress & vbCrLf
        OutBuffer = OutBuffer & "LocationID: " & LocationID & vbCrLf
        OutBuffer = OutBuffer & "LocalServerSerial: " & LocalServerID & vbCrLf
        OutBuffer = OutBuffer & "IP: " & Trim(Request.UserHostAddress) & vbCrLf
        OutBuffer = OutBuffer & "Server: " & Environment.MachineName & vbCrLf
        OutBuffer = OutBuffer & vbCrLf & "Incentive Fetch Check Invalid Serial from MacAddress"
        Common.Send_Email(Common.Get_Error_Emails(5), Common.SystemEmailAddress, "Incentive Fetch Invalid Serial Message - " & Common.InstallationName & "  LocalServerID " & LocalServerID, OutBuffer)
  ElseIf Not (Connector.IsValidLocationForConnectorEngine(Common, LocationID, 2)) Then
    'the location calling IncentiveFetch is not associated with the CPE promoengine
    Common.Write_Log(LogFile, "This location is associated with a promotion engine other than CPE.  Can not proceed.", True)
    Send_Response_Header("This location is associated with a promotion engine other than CPE.  Can not proceed.", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
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

                If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()
                Common.QueryStr = "select LocationName, ExtLocationCode from Locations with (NoLock) where LocationID=" & LocationID & ";"
                rst = Common.LRT_Select
                If rst.Rows.Count > 0 Then
                    LocationName = Common.NZ(rst.Rows(0).Item("LocationName"), "")
                    ExtLocationCode = Common.NZ(rst.Rows(0).Item("ExtLocationCode"), "0")
                End If
    
                OutBuffer = "Incentive Fetch Received invalid request - bad mode" & vbCrLf
                OutBuffer = OutBuffer & "LocationID: " & LocationID.ToString() & vbCrLf
                OutBuffer = OutBuffer & "Location: " & ExtLocationCode & " " & LocationName & vbCrLf
                OutBuffer = OutBuffer & "LocalServerSerial: " & LocalServerID.ToString & vbCrLf
                OutBuffer = OutBuffer & vbCrLf & "Subject: Incentive Fetch Received invalid request"
                Common.Send_Email(Common.Get_Error_Emails(5), Common.SystemEmailAddress, "Incentive Fetch Invalid Request Message - " & Common.InstallationName & "  LocalServerID " & LocalServerID.ToString(), OutBuffer)
    
                
    End Select
  End If 'locationid="0"
  
  If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
%>
<%Response.End()
ErrorTrap:
  Response.Write(Common.Error_Processor(, "Process Info: Server Name=" & Environment.MachineName & " Invoking LocalServerID: " & LocalServerID & "  LocationID=" & LocationID & "  Requester IP Address: " & LocalServerIP & "  Requester Mac Address: " & MacAddress, , Common.InstallationName))
  Common.Write_Log(LogFile, "***** An error occurred during processing! Please check the ErrorLog! *****")
    
    If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()
    Common.QueryStr = "select LocationName, ExtLocationCode from Locations with (NoLock) where LocationID=" & LocationID & ";"
    rst = Common.LRT_Select
    If rst.Rows.Count > 0 Then
        LocationName = Common.NZ(rst.Rows(0).Item("LocationName"), "")
        ExtLocationCode = Common.NZ(rst.Rows(0).Item("ExtLocationCode"), "0")
    End If
    
    Dim ErrorMsg As String = "Incentive Fetch Error during Local Server Processing" & vbCrLf
    ErrorMsg = ErrorMsg & "LocationID: " & LocationID.ToString() & vbCrLf
    ErrorMsg = ErrorMsg & "Location: " & ExtLocationCode & " " & LocationName & vbCrLf
    ErrorMsg = ErrorMsg & "LocalServerSerial: " & LocalServerID.ToString & vbCrLf
    ErrorMsg = ErrorMsg & "MacAddress: " & MacAddress & vbCrLf
    ErrorMsg = ErrorMsg & "IP: " & LocalServerIP & vbCrLf
    ErrorMsg = ErrorMsg & vbCrLf & "Subject: Incentive Fetch Error Exception"
    Common.Send_Email(Common.Get_Error_Emails(5), Common.SystemEmailAddress, "Incentive FetchError in Processing Message - " & Common.InstallationName & "  LocalServerID " & LocalServerID.ToString(), ErrorMsg)
    
  If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
  Common = Nothing
%>
