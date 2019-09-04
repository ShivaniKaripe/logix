<%@ Page CodeFile="..\ConnectorCB.vb" Debug="true" Inherits="ConnectorCB" Language="vb" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.IO.Compression" %>
<%@ Import Namespace="System.Data.SQLClient" %>
<% 
  ' *****************************************************************************
  ' * FILENAME: IPL-Prefs.aspx 
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
  Public TextData As String
  Public IPL As Boolean
  Public LogFile As String
  Public FileStamp As String
  Public FileNum As Integer
  Public StartTime As Decimal
  Public TotalTime As Decimal
  Public ApplicationName As String
  Public ApplicationExtension As String
  Public gzStream As GZipStream = Nothing
  Public UncompressedSize As Long
  Public BufferedRecs As Long
  Public FlushTime As Decimal
  Public FlushStartTime As Decimal
  Public MacAddress As String
  Public LocalServerIP As String
  Public LSVerMajor As Integer
  Public LSVerMinor As Integer
  Public LSBuildMajor As Integer
  Public LSBuildMinor As Integer
  
  ' -------------------------------------------------------------------------------------------------
  
  Sub SD(ByVal OutStr As String)
    'PrintLine(FileNum, OutStr)
    Dim Bytes As Byte()
    Bytes = Encoding.Default.GetBytes(OutStr & vbCrLf)
    UncompressedSize = UncompressedSize + Bytes.Length
    gzStream.Write(Bytes, 0, Bytes.Length)
    Bytes = Nothing
    BufferedRecs = BufferedRecs + 1
    If BufferedRecs >= 5000 Then
      FlushStartTime = DateAndTime.Timer
      Response.Flush()
      FlushTime = FlushTime + (DateAndTime.Timer - FlushStartTime)
      BufferedRecs = 0
    End If
  End Sub
  
  ' -------------------------------------------------------------------------------------------------
  
  Sub SDb(ByVal OutStr As String)
    'Print(FileNum, OutStr)
    Dim Bytes As Byte()
    Bytes = Encoding.Default.GetBytes(OutStr)
    UncompressedSize = UncompressedSize + Bytes.Length
    gzStream.Write(Bytes, 0, Bytes.Length)
    Bytes = Nothing
    BufferedRecs = BufferedRecs + 1
    If BufferedRecs >= 5000 Then
      FlushStartTime = DateAndTime.Timer
      Response.Flush()
      BufferedRecs = 0
      FlushTime = FlushTime + (DateAndTime.Timer - FlushStartTime)
    End If
  End Sub
  
  ' -------------------------------------------------------------------------------------------------
  
  Function Parse_Bit(ByVal BooleanField As Boolean) As String
    If BooleanField Then
      Parse_Bit = "1"
    Else
      Parse_Bit = "0"
    End If
  End Function
  
  ' -------------------------------------------------------------------------------------------------
  
  Function Construct_Table(ByVal TableName As String, ByVal Operation As Integer, ByVal DelimChar As Integer, ByVal LocalServerID As String, ByVal LocationID As String, ByVal DBName As String) As String
    
    Dim TempResults As String
    Dim NumRecs As Long
    Dim TempOut As String
    Dim FieldList As String
    Dim DataBack As String
    Dim OperationType As Integer
    Dim QueryStartTime As Decimal
    Dim QueryTotalTime As Decimal
    Dim ConstructStartTime As Decimal
    Dim ConstructTotalTime As Decimal
    Dim reader As SqlDataReader
    
    'Send ("<!-- Table=" & TableName & " " & Format$(Time, "hh:mm:ss") & Format$(Timer - Fix(Timer), ".00") & " -->")
    
    TempOut = ""
    TempResults = ""
    NumRecs = 0
    FieldList = ""
    DataBack = ""
    
    Common.LRTTimeout = 14400
    Common.LWHTimeout = 14400
    Common.LXSTimeout = 14400
    
    QueryStartTime = DateAndTime.Timer
    If UCase(DBName) = "LXS" Then
      'dst = Common.LXS_Select
      Common.LXScmd = New System.Data.SqlClient.SqlCommand
      Common.LXScmd.Connection = Common.LXSadoConn
      Common.LXScmd.CommandType = CommandType.Text
      Common.LXScmd.CommandTimeout = Common.LXSTimeout
      Common.LXScmd.CommandText = Common.QueryStr
      reader = Common.LXScmd.ExecuteReader()
    ElseIf UCase(DBName) = "LWH" Then
      'dst = Common.LWH_Select
      Common.LWHcmd = New System.Data.SqlClient.SqlCommand
      Common.LWHcmd.Connection = Common.LWHadoConn
      Common.LWHcmd.CommandType = CommandType.Text
      Common.LWHcmd.CommandTimeout = Common.LWHTimeout
      Common.LWHcmd.CommandText = Common.QueryStr
      reader = Common.LWHcmd.ExecuteReader()
    Else
      'dst = Common.LRT_Select
      Common.LRTcmd = New System.Data.SqlClient.SqlCommand
      Common.LRTcmd.Connection = Common.LRTadoConn
      Common.LRTcmd.CommandType = CommandType.Text
      Common.LRTcmd.CommandTimeout = Common.LRTTimeout
      Common.LRTcmd.CommandText = Common.QueryStr
      reader = Common.LRTcmd.ExecuteReader()
    End If
    
    QueryTotalTime = DateAndTime.Timer - QueryStartTime
    ConstructStartTime = DateAndTime.Timer
    
    If reader.HasRows Then
      If UCase(TableName) = "CUSTOMERPREFERENCES" Then
        FieldList = "CustomerPK" & Chr(DelimChar) & "PreferenceID" & Chr(DelimChar) & "Value"
      End If
      OperationType = Operation
      If OperationType = 99 Then OperationType = 2
      'send the table header
      TempOut = "1:" & TableName & vbCrLf
      TempOut = TempOut & "2:" & Trim(Str(OperationType)) & vbCrLf
      TempOut = TempOut & "3:" & FieldList
      SD(TempOut)
      Common.Write_Log(LogFile, TempOut)
      NumRecs = 0
      
      If UCase(TableName) = "CUSTOMERPREFERENCES" And (Operation = 5) Then
        While reader.Read
          SD(reader.Item("Data"))
          NumRecs = NumRecs + 1
        End While
      End If
      SD("###")
      Common.Write_Log(LogFile, "# records: " & NumRecs)
      ConstructTotalTime = DateAndTime.Timer - ConstructStartTime
      TotalTime = DateAndTime.Timer - StartTime
      Common.Write_Log(LogFile, "Query took " & Int(QueryTotalTime) & Format$(QueryTotalTime - Fix(QueryTotalTime), ".000") & "(sec) - Constructing data took " & Int(ConstructTotalTime) & Format$(ConstructTotalTime - Fix(ConstructTotalTime), ".000") & "(sec) - Total elapsed time " & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".000") & "(sec)" & vbCrLf)
    End If  'reader.HasRows
    
    reader.Close()
    reader = Nothing
    If UCase(DBName) = "LXS" Then
      Common.LXScmd = Nothing
    ElseIf UCase(DBName) = "LWH" Then
      Common.LWHcmd = Nothing
    Else
      Common.LRTcmd = Nothing
    End If
    
    Construct_Table = DataBack
    
  End Function
  
  ' -------------------------------------------------------------------------------------------------
  
  Sub Construct_Output(ByVal LocalServerID As Long, ByVal LocationID As Long)
    
    Dim OutStr As String
    Dim TempOut As String
    Dim DelimChar As Integer
    
    DelimChar = 30
    TempOut = ""
    
    Common.Write_Log(LogFile, "Returned the following data:")
    OutStr = ""
    OutStr = OutStr & ApplicationName & "," & Connector.CSMajorVersion & "," & Connector.CSMinorVersion & "," & Connector.CSBuild & "," & Connector.CSBuildRevision & vbCrLf
    OutStr = OutStr & "D:" & vbCrLf
    OutStr = OutStr & "T:" & vbCrLf
    OutStr = OutStr & "LocationID=" & LocationID
    SD(OutStr)
    Common.Write_Log(LogFile, OutStr)
    
    'Prefs
    Common.QueryStr = "select convert(nvarchar, P.CustomerPK) + char(" & DelimChar & ") + convert(nvarchar, P.PreferenceID) + char(" & DelimChar & ") + isnull(P.Value, '') as Data " & _
             "from CustomerPreferences as P with (NoLock) Inner Join CustomerLocations as CL with (NoLock) on P.CustomerPK=CL.CustomerPK " & _
             "where CL.LocationID=" & LocationID & ";"
    OutStr = OutStr & Construct_Table("CustomerPreferences", 5, DelimChar, LocalServerID, LocationID, "LXS")
    
    SD("***") 'send the EOF marker
    
  End Sub
  
  
  ' -------------------------------------------------------------------------------------------------
  
  
  Sub Process_IPL(ByVal IPLTypeID As Integer, ByVal LocalServerID As Long, ByVal LocationID As Long, ByVal IPLSessionID As String, ByVal HistoryStartTime As DateTime)
    
    Dim MustIPL As Boolean
    Dim dst As DataTable
    
    'see if this server needs to be IPL'd - if so, send a response header
    MustIPL = False
    Common.QueryStr = "dbo.pa_CPE_CheckMustIPL"
    Common.Open_LRTsp()
    Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
    dst = Common.LRTsp_select
    Common.Close_LRTsp()
    If dst.Rows.Count > 0 Then
      MustIPL = Common.NZ(dst.Rows(0).Item("MustIPL"), True)
    End If
    dst = Nothing
    If MustIPL Then
      Send_Response_Header("MustIPL", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
    Else
        
      'record the starting of the IPL
      Common.QueryStr = "dbo.pa_IPL_HistoryStart"
      Common.Open_LRTsp()
      Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
      Common.LRTsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = LocationID
      Common.LRTsp.Parameters.Add("@IPLTypeID", SqlDbType.Int).Value = IPLTypeID
      Common.LRTsp.Parameters.Add("@StartTable", SqlDbType.VarChar, 200).Value = ""
      Common.LRTsp.Parameters.Add("@StartPK", SqlDbType.BigInt).Value = 0
      Common.LRTsp.Parameters.Add("@SessionID", SqlDbType.VarChar, 50).Value = IPLSessionID
      Common.LRTsp.Parameters.Add("@StartTime", SqlDbType.DateTime).Direction = ParameterDirection.Output
      Common.LRTsp.ExecuteNonQuery()
      HistoryStartTime = Common.LRTsp.Parameters("@StartTime").Value
      Common.Close_LRTsp()
        
      Common.Write_Log(LogFile, "Removing buffered TransDownload data")
      Common.QueryStr = "dbo.pc_CPE_IPLPurge_Prefs_Output"
      Common.Open_LXSsp()
      Common.LXSsp.CommandTimeout = 1200
      Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
      Common.LXSsp.ExecuteNonQuery()
      Common.Close_LXSsp()
      TotalTime = DateAndTime.Timer - StartTime
      Common.Write_Log(LogFile, "Finished removing buffered TransDownload data.  Total elapsed time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".000") & "(sec)")
       
      gzStream = New GZipStream(Response.OutputStream, CompressionMode.Compress, True)

      Construct_Output(LocalServerID, LocationID)
        
      TotalTime = DateAndTime.Timer - StartTime
      Common.Write_Log(LogFile, "Finished All Queries. Closing GZip Stream ... " & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".000") & "(sec)")
      If gzStream IsNot Nothing Then
        gzStream.Close()
        gzStream.Dispose()
        gzStream = Nothing
      End If
        
      TotalTime = DateAndTime.Timer - StartTime
      Common.Write_Log(LogFile, "GZip stream closed.  Flushing final records ... " & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".000") & "(sec)")
        
      FlushStartTime = DateAndTime.Timer
      Response.Flush()
      FlushTime = FlushTime + (DateAndTime.Timer - FlushStartTime)
      TotalTime = DateAndTime.Timer - StartTime
      Common.Write_Log(LogFile, "Total time to send all records to client=" & Int(FlushTime) & Format$(FlushTime - Fix(FlushTime), ".000") & "(sec)" & "  Total elapsed time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".000") & "(sec)")
        
      Common.Write_Log(LogFile, "Total uncompressed size = " & UncompressedSize)

      'update the history record for the IPL end time
      Common.QueryStr = "dbo.pa_IPL_HistoryEnd"
      Common.Open_LRTsp()
      Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
      Common.LRTsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = LocationID
      Common.LRTsp.Parameters.Add("@IPLTypeID", SqlDbType.Int).Value = IPLTypeID
      Common.LRTsp.Parameters.Add("@StartTime", SqlDbType.DateTime).Value = HistoryStartTime
      Common.LRTsp.Parameters.Add("@UncompressedSize", SqlDbType.BigInt).Value = UncompressedSize
      Common.LRTsp.Parameters.Add("@CompressedSize", SqlDbType.BigInt).Value = 0
      Common.LRTsp.ExecuteNonQuery()
      Common.Close_LRTsp()
    End If

  End Sub
  
  ' -------------------------------------------------------------------------------------------------
  
  Sub Process_ACK(ByVal LocalServerID As Long, ByVal LocationID As Long)
    
    Common.QueryStr = "Update LocalServers with (RowLock) set IncentiveLastHeard=getdate(), MustIPL=0 where LocalServerID='" & LocalServerID & "';"
    Common.LRT_Execute()
    Send_Response_Header(ApplicationName & " - ACK Received", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
    Send("ACK Received")
    Common.Write_Log(LogFile, "ACK Received")
    
  End Sub
  
  ' -------------------------------------------------------------------------------------------------
  
  Sub Process_NAK(ByVal LocalServerID As String, ByVal LocationID As String)
    
    Dim ErrorMsg As String
        
    LocalServerID = Common.Extract_Val(Request.QueryString("serial"))
    LocalServerIP = Trim(Request.QueryString("IP"))
    MacAddress = Trim(Request.QueryString("mac"))
    If MacAddress = "" Or MacAddress = "0" Then
      MacAddress = "0"
    End If
    If LocalServerIP = "" Or LocalServerIP = "0" Then
      Common.Write_Log(LogFile, "Could not get IP from query. Analyzing client request for IP ...")
      LocalServerIP = Trim(Request.UserHostAddress)
    End If
        
    ErrorMsg = Trim(Request.QueryString("errormsg"))
    Common.Write_Log(LogFile, "Serial:" & LocalServerID & " MacAddress:" & MacAddress & " IP:" & LocalServerIP & "  server:" & Environment.MachineName & " Received NAK - ErrorMsg:" & ErrorMsg)
    Send_Response_Header(ApplicationName & " - NAK Received", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
    
  End Sub

</script>

<%
  
  ' -------------------------------------------------------------------------------------------------
  ' Execution starts here
  
  Dim ServerSerial As String
  Dim LocalServerID As Long
  Dim LocationID As Long
  Dim LastHeard As String
  Dim Mode As String
  Dim HistoryStartTime As DateTime
  Dim IPLTypeID As Integer
  Dim dst As DataTable
  
  Dim IPLSessionID As String
  Dim BannerID As Integer
  Dim LSVerParts() As String
  
  
  IPLTypeID = 10
  ApplicationName = "IPL-Prefs"
  ApplicationExtension = ".aspx"
  Common.AppName = ApplicationName & ApplicationExtension
  Response.Expires = 0
  On Error GoTo ErrorTrap
  
  StartTime = DateAndTime.Timer
     
  LocalServerID = Common.Extract_Val(Request.QueryString("serial"))
  LogFile = "CPE-IncentiveFetchLog-" & Format(LocalServerID, "00000") & "." & Common.Leading_Zero_Fill(Year(Today), 4) & Common.Leading_Zero_Fill(Month(Today), 2) & Common.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
  LocalServerIP = Trim(Request.QueryString("IP"))
  MacAddress = Trim(Request.QueryString("mac"))
  If MacAddress = "" Or MacAddress = "0" Then
    MacAddress = "0"
  End If
  If LocalServerIP = "" Or LocalServerIP = "0" Then
    Common.Write_Log(LogFile, "Could not get IP from query. Analyzing client request for IP ...")
    LocalServerIP = Trim(Request.UserHostAddress)
  End If  
  
  LSVersion = Trim(Request.QueryString("lsversion"))
  LSVerMajor = 0
  LSVerMinor = 0
  If InStr(LSVersion, ".", CompareMethod.Binary) > 0 Then
    LSVerParts = Split(LSVersion, ".", , CompareMethod.Binary)
    LSVerMajor = Common.Extract_Val(LSVerParts(0))
    LSVerMinor = Common.Extract_Val(LSVerParts(1))
  End If
  LSBuild = Trim(Request.QueryString("lsbuild"))
  LSBuildMajor = 0
  LSBuildMinor = 0
  If InStr(LSBuild, ".", CompareMethod.Binary) > 0 Then
    LSVerParts = Split(LSBuild, ".", , CompareMethod.Binary)
    LSBuildMajor = Common.Extract_Val(LSVerParts(0))
    LSBuildMinor = Common.Extract_Val(LSVerParts(1))
  End If
  
  LastHeard = "1/1/1980"
  
  IPLSessionID = "0"
  If (LSVerMajor > 5) Or (LSVerMajor = 5 And LSVerMinor > 4) Or (LSVerMajor = 5 And LSVerMinor = 4 And LSBuildMajor >= 6) Then
    IPLSessionID = IIf(Request.QueryString("sessionid") <> "", Request.QueryString("sessionid"), "0")
  End If
  
  Common.Open_LogixRT()
  Common.Open_LogixXS()
  Common.Load_System_Info()
  Connector.Load_System_Info(Common)
  
  BufferedRecs = 0
  
  Common.Write_Log(LogFile, "---------------------------------------------------------------------------")
  
  Connector.Get_LS_Info(Common, LocalServerID, LocationID, BannerID, LastHeard, 1, LocalServerIP)
  
  If LocationID = "0" Then
    Common.Write_Log(LogFile, Common.AppName & "   Invalid Serial Number:" & LocalServerID & " from MacAddress:" & MacAddress & " IP:" & LocalServerIP & "  Process running on server:" & Environment.MachineName, True)
    Send_Response_Header(ApplicationName & " - Invalid Serial Number", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
  ElseIf Not (Connector.IsValidLocationForConnectorEngine(Common, LocationID, 2)) Then
    'the location calling IPL-Prefs is not associated with the CPE promoengine
    Common.Write_Log(LogFile, "This location is associated with a promotion engine other than CPE.  Can not proceed.", True)
    Send_Response_Header("This location is associated with a promotion engine other thanCPE.  Can not proceed.", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
  Else 'LocationID is valid
    Mode = UCase(Request.QueryString("mode"))
    Common.Write_Log(LogFile, "** " & Common.AppName & "   " & Microsoft.VisualBasic.DateAndTime.Now & "  CSVersion: " & Connector.CSMajorVersion & "." & Connector.CSMinorVersion & "b" & Connector.CSBuild & "r" & Connector.CSBuildRevision & "  LSVersion=" & LSVerMajor & "." & LSVerMinor & "b" & LSBuildMajor & "." & LSBuildMinor & "  Serial: " & LocalServerID & "   LocationID: " & LocationID & "  Mode: " & Mode & "  Process running on server:" & Environment.MachineName)
    Common.Write_Log(LogFile, "SessionID=" & IPLSessionID)
    
    'Check and see if the PrefMan integration is enabled, and if Pref Data Distribution is enabled
    If Common.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER) And Common.Fetch_CPE_SystemOption(128) = "1" Then
    
      If Mode = "ACK" Then
        Process_ACK(LocalServerID, LocationID)
      ElseIf Mode = "NAK" Then
        Process_NAK(LocalServerID, LocationID)
      ElseIf (Mode = "IPL") Then
        Process_IPL(IPLTypeID, LocalServerID, LocationID, IPLSessionID, HistoryStartTime)
      Else  'no matches for MODE=
        Send_Response_Header("Invalid Request", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
      End If
      
    Else 'Something is not enabled
      If Not (Common.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER)) Then
        'PrefMan integration is not enabled
        Common.Write_Log(LogFile, "PrefMan integration is not installed - NAK returned to caller")
        Process_NAK(LocalServerID, LocationID)
      Else
        'Pref Data Distribution is not enabled
        Common.Write_Log(LogFile, "Preference data distribution (CPESystemOptionID 128) is disabled - NAK returned to caller")
        Process_NAK(LocalServerID, LocationID)
      End If
    End If 'Is the PrefMan integration enabled and Pref Data Distribution enabled
    
  End If 'locationid="0"
    
  TotalTime = DateAndTime.Timer - StartTime
  Common.Write_Log(LogFile, "Closing database connections - Total elapsed time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".000") & "(sec)")
  
  Common.Close_LogixRT()
  Common.Close_LogixXS()
  
  TotalTime = DateAndTime.Timer - StartTime
  Common.Write_Log(LogFile, "RunTime=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".000") & "(sec)")
%>
<% 
  Response.End()
ErrorTrap:
  Response.Write(Common.Error_Processor(, "Serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " Process Info: Server Name=" & Environment.MachineName & " Invoking LocationID=" & LocationID, , Common.InstallationName))
  Common.Write_Log(LogFile, "***** An error occurred during processing! Please check the ErrorLog! *****")
  If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
  If Not (Common.LXSadoConn.State = ConnectionState.Closed) Then Common.Close_LogixXS()
  Common = Nothing
%>
