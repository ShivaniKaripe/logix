<%@ Page CodeFile="..\ConnectorCB.vb" Debug="true" Inherits="ConnectorCB" Language="vb" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SQLClient" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.IO.Compression" %>
<% 
  ' *****************************************************************************
  ' * FILENAME: IPL-Customers.aspx 
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
  Public IPL As Boolean
  Public LogFile As String
  Public StartTime As Decimal
  Public TotalTime As Decimal
  Public ApplicationName As String
  Public ApplicationExtension As String
  Public LastPK As Long
  Public LastTable As String
  Public StartPK As Long
  Public StartTable As String
  Public MaxBatchSize As Long
  Public EndPK As Long
  Public MacAddress As String
  Public LocalServerIP As String
  Public LSVerMajor As Integer
  Public LSVerMinor As Integer
  Public LSBuildMajor As Integer
  Public LSBuildMinor As Integer
  Public gzStream As GZipStream = Nothing
  Public UncompressedSize As Long = 0
  Public BufferedRecs As Long = 0
  Public FlushTime As Decimal = 0
  Public FlushStartTime As Decimal = 0

  ' the following variables are used for GZip streaming with version 5.19 and above
  Public GZIP As New Copient.GZIPInc
  Public TextData As String
  Public FileStamp As String
  Public FileNum As Integer
  Public ZipOutput As Boolean
  Public  MyCryptLib As New Copient.CryptLib
  ' -------------------------------------------------------------------------------------------------
  
  Sub SD(ByVal OutStr As String)
    If Not (ZipOutput) Then
      Dim Bytes As Byte()
      Bytes = Encoding.UTF8.GetBytes(OutStr & vbCrLf)
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
    Else
      PrintLine(FileNum, OutStr)
    End If
  End Sub
  
  ' -------------------------------------------------------------------------------------------------
  
  Sub SDb(ByVal OutStr As String)
    
    If Not (ZipOutput) Then
      Dim Bytes As Byte()
      Bytes = Encoding.UTF8.GetBytes(OutStr)
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
    Else
      Print(FileNum, OutStr)
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
    Dim FieldList As String
    Dim OperationType As Integer
    Dim QueryStartTime As Decimal
    Dim QueryTotalTime As Decimal
    Dim ConstructStartTime As Decimal
    Dim ConstructTotalTime As Decimal
        Dim reader As SqlDataReader
        
    Const BUFFERED_WRITE_SIZE As Integer = 10000
    'Send ("<!-- Table=" & TableName & " " & Format$(Time, "hh:mm:ss") & Format$(Timer - Fix(Timer), ".00") & " -->")
    
    TempResults = ""
    NumRecs = 0
    FieldList = ""
    
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
      If UCase(TableName) = "USERS" Then
        FieldList = "UserID" & Chr(DelimChar) & "ClientUserID1" & Chr(DelimChar) & "HHPrimaryID" & Chr(DelimChar) & "HHrec" & Chr(DelimChar) & "CustomerTypeID" & Chr(DelimChar) & "Employee" & Chr(DelimChar) & "CurrYearSTD" & Chr(DelimChar) & "LastYearSTD" & Chr(DelimChar) & "CustomerStatusID" & Chr(DelimChar) & "LastVisit" & Chr(DelimChar) & "AltIDOptOut" & Chr(DelimChar) & "AlternateID" & Chr(DelimChar) & "Verifier" & Chr(DelimChar) & "EmployeeID" & Chr(DelimChar) & "FirstName" & Chr(DelimChar) & "LastName" & Chr(DelimChar) & "AirmileMemberID" & Chr(DelimChar) & "Prefix" & Chr(DelimChar) & "Suffix"
      End If

      If UCase(TableName) = "CARDIDS" Then
        FieldList = "CardPK" & Chr(DelimChar) & "UserID" & Chr(DelimChar) & "ExtCardID" & Chr(DelimChar) & "CardStatusID" & Chr(DelimChar) & "CardTypeID"
      End If
      
      OperationType = Operation
      If OperationType = 99 Then OperationType = 2
      'send the table header
      Dim TempOut As String = "1:" & TableName & vbCrLf
      TempOut = TempOut & "2:" & Trim(Str(OperationType)) & vbCrLf
      TempOut = TempOut & "3:" & FieldList
      SD(TempOut)
      Common.Write_Log(LogFile, TempOut, True)
      NumRecs = 0
      
      Dim buf As New StringBuilder()
      
      If UCase(TableName) = "USERS" And Operation = 5 Then
        LastTable = "Users"
                While reader.Read()
                    Dim TempAppendLine As String = reader.Item("Customer_PK") & Chr(DelimChar) & Common.NZ(MyCryptLib.SQL_StringDecrypt(reader.Item("InitialCardIDOriginal").ToString()), 0) & Chr(DelimChar) & reader.Item("data")
                    buf.AppendLine(TempAppendLine)
                    If (buf.Length > BUFFERED_WRITE_SIZE) Then
                        SDb(buf.ToString())
                        buf.Clear()
                    End If
                    LastPK = reader.Item("CustomerPK")
                    EndPK = LastPK
                    NumRecs = NumRecs + 1
                End While
        If Not (NumRecs = MaxBatchSize) Or (MaxBatchSize = 0) Then
          LastTable = ""
          'EndPK = reader.Item("CustomerPK")
          LastPK = 0
        End If
      Else
        While reader.Read()
          buf.AppendLine(reader.Item("CardPK") & Chr(DelimChar) & reader.Item("CustomerPK") & Chr(DelimChar) & MyCryptLib.SQL_StringDecrypt(reader.Item("ExtCardIDOriginal")) & Chr(DelimChar) & reader.Item("CardStatusID") & Chr(DelimChar) & reader.Item("CardTypeID"))
          If ( buf.Length > BUFFERED_WRITE_SIZE )
            SDb( buf.ToString() )
            buf.Clear()
          End If
          NumRecs = NumRecs + 1
        End While
      End If

      buf.AppendLine("###")
      SDb( buf.ToString() )
      Common.Write_Log(LogFile, "# records: " & NumRecs, True)
      If Not (LastTable = "") Then
        Common.Write_Log(LogFile, "LastTable=" & TableName & "   LastPK=" & LastPK, True)
      End If
      ConstructTotalTime = DateAndTime.Timer - ConstructStartTime
      TotalTime = DateAndTime.Timer - StartTime
      Common.Write_Log(LogFile, "Query took " & Int(QueryTotalTime) & Format$(QueryTotalTime - Fix(QueryTotalTime), ".000") & "(sec) - Constructing data took " & Int(ConstructTotalTime) & Format$(ConstructTotalTime - Fix(ConstructTotalTime), ".000") & "(sec) - Total elapsed time " & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".000") & "(sec)" & vbCrLf, True)
    End If  'rst.recordcount=0
    
    reader.Close()
    reader = Nothing
    If UCase(DBName) = "LXS" Then
      Common.LXScmd = Nothing
    ElseIf UCase(DBName) = "LWH" Then
      Common.LWHcmd = Nothing
    Else
      Common.LRTcmd = Nothing
    End If
    
    Return ""
    
  End Function
  
  ' -------------------------------------------------------------------------------------------------
  
  Sub Construct_Output(ByVal LocalServerID As Long, ByVal LocationID As Long)
  
    Dim DelimChar As Integer
    Dim VerifyDef As String = ""
    Dim AltIDDef As String = ""
    Dim MaxBatchStr As String
    Dim OutStr As String = ""
    
    DelimChar = 30
    
    If MaxBatchSize > 0 Then
      MaxBatchStr = "top " & MaxBatchSize
    Else
      MaxBatchStr = ""
    End If
        
    If Not (ZipOutput) Then
      Common.Write_Log(LogFile, "Returned the following data:", True)
      If (LSVerMajor < 5) Or (LSVerMajor = 5 And LSVerMinor < 4) Or (LSVerMajor = 5 And LSVerMinor = 4 And LSBuildMajor < 6) Then
        OutStr = OutStr & ApplicationName & "," & Connector.CSMajorVersion & "," & Connector.CSMinorVersion & "," & Connector.CSBuild & "," & Connector.CSBuildRevision & vbCrLf
        OutStr = OutStr & "D:" & vbCrLf
        OutStr = OutStr & "T:" & vbCrLf
        OutStr = OutStr & "LocationID=" & LocationID
      Else
        OutStr = OutStr & ApplicationName & "," & Connector.CSMajorVersion & "," & Connector.CSMinorVersion & "," & Connector.CSBuild & "," & Connector.CSBuildRevision & vbCrLf
        OutStr = OutStr & "LocationID=" & LocationID & vbCrLf
      End If
      SD(OutStr)
      Common.Write_Log(LogFile, OutStr, True)
    Else
      If (LSVerMajor < 5) Or (LSVerMajor = 5 And LSVerMinor < 4) Or (LSVerMajor = 5 And LSVerMinor = 4 And LSBuildMajor < 6) Then
        OutStr = OutStr & ApplicationName & "," & Connector.CSMajorVersion & "," & Connector.CSMinorVersion & "," & Connector.CSBuild & "," & Connector.CSBuildRevision & vbCrLf
        OutStr = OutStr & "D:" & vbCrLf
        OutStr = OutStr & "T:" & vbCrLf
        OutStr = OutStr & "LocationID=" & LocationID
        SD(OutStr)
        Common.Write_Log(LogFile, OutStr, True)
      End If
    End If
    
    
    
    ' determine the table.column for the verifier
    VerifyDef = Common.Fetch_SystemOption(61)
    
    ' determine the table.column for the alternanteid
    AltIDDef = Common.Fetch_SystemOption(60)
    
    If (AltIDDef <> "") Then
      AltIDDef = "+ isnull(" & AltIDDef & ",'') + char(" & DelimChar & ") "
    Else
      AltIDDef = "+ '' + char(" & DelimChar & ") "
    End If

    If (VerifyDef <> "") Then
      VerifyDef = "+ isnull(" & VerifyDef & ",'')"
    Else
      VerifyDef = "+ ''"
    End If
    
    
    'send the data from the Customers table
    'Common.QueryStr = "select distinct C.CustomerPK as UserID, PrimaryExtID as ClientUserID1, HHPK as HHPrimaryID, CustomerTypeID as HHRec, Employee, CurrYearSTD, LastYearSTD, isnull(CardStatusID, 1) as CardStatusID " & _
    '           "from Customers as C with (NoLock) Inner Join CustomerLocations as CL with (NoLock) on C.CustomerPK=CL.CustomerPK " & _
    '           "where CL.LocationID=" & LocationID & ";"
    'Common.QueryStr = "select  " & MaxBatchStr & " RA.RewardAccumulationID, convert(nvarchar, RA.RewardAccumulationID) + char(" & DelimChar & ") + convert(nvarchar, RA.RewardOptionID) + char(" & DelimChar & ") + convert(nvarchar, RA.CustomerPK) + char(" & DelimChar & ") + convert(nvarchar, QtyPurchased) + char(" & DelimChar & ") + convert(nvarchar, TotalPrice) + char(" & DelimChar & ") + convert(varchar, AccumulationDate, 20) + char(" & DelimChar & ") + convert(nvarchar, isnull(LocalID, 0)) + char(" & DelimChar & ") + convert(nvarchar, isnull(ServerSerial, 0)) + char(" & DelimChar & ") + '2' as data " & _
        Common.QueryStr = "select distinct " & MaxBatchStr & " Customers.CustomerPK, convert(nvarchar, Customers.CustomerPK) as Customer_PK, InitialCardIDOriginal,  convert(nvarchar, isnull(HHPK, 0)) + char(" & DelimChar & ") + case convert(nvarchar, isnull(CustomerTypeID, 0)) when 2 then '0' else convert(nvarchar, isnull(CustomerTypeID, 0)) end   + char(" & DelimChar & ") + convert(nvarchar, isnull(CustomerTypeID, 0)) + char(" & DelimChar & ") + convert(nvarchar, isnull(Employee, 0)) + char(" & DelimChar & ") + convert(nvarchar, isnull(CurrYearSTD, 0)) + char(" & DelimChar & ") " & _
               "  + convert(nvarchar, isnull(LastYearSTD, 0)) + char(" & DelimChar & ") + convert(nvarchar, isnull(CustomerStatusID, 1)) + char(" & DelimChar & ") + convert(nvarchar, isnull(CL.LastVisit, '1/1/1980'), 120) + char(" & DelimChar & ") + convert(nvarchar, isnull(Customers.AltIDOptOut, '0')) + char(" & DelimChar & ") " & AltIDDef & VerifyDef & " + char(" & DelimChar & ") + isnull(EmployeeID, '') + char(" & DelimChar & ") + isnull(FirstName, '') + char(" & DelimChar & ") + isnull(LastName, '') + char(" & DelimChar & ") + isnull(CustomerExt.AirmileMemberID, '') + char(" & DelimChar & ") + isnull(Customers.Prefix, '') + char(" & DelimChar & ") + isnull(Customers.Suffix, '') as data " & _
               " from Customers with (NoLock) Inner Join CustomerLocations as CL with (NoLock) on Customers.CustomerPK=CL.CustomerPK left Join CustomerExt with (NoLock) on Customers.CustomerPK=CustomerExt.CustomerPK " & _
               "where CL.LocationID=" & LocationID & " "
    If StartPK > 0 Then
      Common.QueryStr = Common.QueryStr & "and Customers.CustomerPK>" & StartPK & " "
    End If
    Common.QueryStr = Common.QueryStr & "order by Customers.CustomerPK;"
    
    Construct_Table("Users", 5, DelimChar, LocalServerID, LocationID, "LXS")
    
    'send the data from the CardIDs table
        Common.QueryStr = "select distinct CardIDs.CustomerPK, CardIDs.CardPK, CardIDs.CustomerPK, CardIDs.ExtCardIDOriginal, CardIDs.CardStatusID, CardIDs.CardTypeID " & _
               " from CardIDs with (NoLock) Inner Join CustomerLocations as CL with (NoLock) on CardIDs.CustomerPK=CL.CustomerPK " & _
               "where CL.LocationID=" & LocationID & " "
    If StartPK > 0 Then
      Common.QueryStr = Common.QueryStr & "and CardIDs.CustomerPK >" & StartPK & " "
    End If
    If EndPK > 0 Then
      Common.QueryStr = Common.QueryStr & "and CardIDs.CustomerPK <=" & EndPK & " "
    End If
    
    Common.QueryStr = Common.QueryStr & "order by CardIDs.CustomerPK;"
    
    Construct_Table("CardIDs", 5, DelimChar, LocalServerID, LocationID, "LXS")
    
    SD("***") 'send the EOF marker
    
  End Sub
  
  ' -------------------------------------------------------------------------------------------------
  
  Sub Process_ACK(ByVal LocalServerID As Long, ByVal LocationID As Long)
  
    Common.QueryStr = "Update LocalServers with (RowLock) set IncentiveLastHeard=getdate(), MustIPL=0 where LocalServerID='" & LocalServerID & "';"
    Common.LRT_Execute()
    Send_Response_Header(ApplicationName & " - ACK Received", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
    Send("ACK Received")
    Common.Write_Log(LogFile, "ACK Received", True)
    
  End Sub
  
  ' -------------------------------------------------------------------------------------------------
  
  Sub Process_NAK(ByVal LocalServerID As String, ByVal LocationID As String)
  
    Dim ErrorMsg As String
    MacAddress = Trim(Request.QueryString("mac"))
    
    If MacAddress = "" Or MacAddress = "0" Then
      MacAddress = "0"
    End If
    LocalServerID = Common.Extract_Val(Request.QueryString("serial"))
    LocalServerIP = Trim(Request.QueryString("IP"))
    If LocalServerIP = "" Or LocalServerIP = "0" Then
      Common.Write_Log(LogFile, "Could not get IP from query. Analyzing client request for IP ...", True)
      LocalServerIP = Trim(Request.UserHostAddress)
    End If
   
    
    ErrorMsg = Trim(Request.QueryString("errormsg"))
    Common.Write_Log(LogFile, "Serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " Received NAK - ErrorMsg:" & ErrorMsg & " Server:" & Environment.MachineName & " server=" & Environment.MachineName, True)
    Send_Response_Header(ApplicationName & " - NAK Received", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
    
  End Sub

</script>

<%
  Dim ServerSerial As String
  Dim LocalServerID As Long
  Dim LocationID As Long
  Dim LastHeard As String
  Dim Mode As String
  Dim HistoryStartTime As DateTime
  Dim IPLTypeID As Integer
  Dim dst As DataTable
  Dim MustIPL As Boolean
  Dim OutStr As String
  Dim UncompressedSize As Long
  Dim IPLSessionID As String
  Dim IPLTempPath As String
  Dim BannerID As Integer
  Dim LSVerParts() As String
  
  ' the following variables are used for GZip streaming with version 5.19 and above
  Dim DataFile As String
  Dim ZipFile As String
  Dim FileStamp As String
  Dim BatchFile As String
  Dim LockFile As String
  Dim LaunchAttempts As Integer
  Dim ReturnCode As Long
  Dim WorkspaceFilePath As String
  Dim CompressedArray() As Byte
  Dim CompressedSize As Long
  
  
  IPLTypeID = 3
  ApplicationName = "IPL-Customers"
  ApplicationExtension = ".aspx"
  Common.AppName = ApplicationName & ApplicationExtension
  Response.Expires = 0
  On Error GoTo ErrorTrap
    
  LocalServerID = Common.Extract_Val(Request.QueryString("serial"))
  LogFile = "CPE-IncentiveFetchLog-" & Format(LocalServerID, "00000") & "." & Common.Leading_Zero_Fill(Year(Today), 4) & Common.Leading_Zero_Fill(Month(Today), 2) & Common.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
  
  MacAddress = Trim(Request.QueryString("mac"))    
  If MacAddress = "" Or MacAddress = "0" Then
    MacAddress = "0"
  End If
    
  LocalServerIP = Trim(Request.QueryString("IP"))
  If LocalServerIP = "" Or LocalServerIP = "0" Then
    Common.Write_Log(LogFile, "Could not get IP from query. Analyzing client request for IP ...", True)
    LocalServerIP = Trim(Request.UserHostAddress)
  End If
   
  StartTime = DateAndTime.Timer
  
  
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
  StartTable = ""
  StartPK = 0
  IPLSessionID = "0"
  If (LSVerMajor > 5) Or (LSVerMajor = 5 And LSVerMinor > 4) Or (LSVerMajor = 5 And LSVerMinor = 4 And LSBuildMajor >= 6) Then
    StartTable = IIf(Request.QueryString("starttable") <> "", Request.QueryString("starttable"), "")
    StartPK = Common.Extract_Val(Request.QueryString("startpk"))
    IPLSessionID = IIf(Request.QueryString("sessionid") <> "", Request.QueryString("sessionid"), "0")
  End If
  
  Common.Open_LogixRT()
  Common.Open_LogixXS()
  Common.Load_System_Info()
  Connector.Load_System_Info(Common)
  
  Common.Write_Log(LogFile, "---------------------------------------------------------------------------", True)
  
  If (LSVerMajor > 5) Or (LSVerMajor = 5 And LSVerMinor > 4) Or (LSVerMajor = 5 And LSVerMinor = 4 And LSBuildMajor >= 6) Then
    MaxBatchSize = Common.Extract_Val(Common.Fetch_CPE_SystemOption(60))
  Else
    MaxBatchSize = 0
  End If
  
  ' determine if GZip streaming should be used based on Logix version (5.19 and above)
  ZipOutput = Not ((LSVerMajor > 5) OrElse ((LSVerMajor = 5) AndAlso (LSVerMinor >= 19)))

  Connector.Get_LS_Info(Common, LocalServerID, LocationID, BannerID, LastHeard, 1, LocalServerIP)
  
  If LocationID = "0" Then
    Common.Write_Log(LogFile, Common.AppName & "    Invalid Serial Number:" & LocalServerID & " from MacAddress: " & MacAddress & " IP=" & LocalServerIP & "  Process running on server:" & Environment.MachineName, True)
    Send_Response_Header(ApplicationName & " - Invalid Serial Number", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
  ElseIf Not (Connector.IsValidLocationForConnectorEngine(Common, LocationID, 2)) Then
    'the location calling IPL-Customers is not associated with the CPE promoengine
    Common.Write_Log(LogFile, "This location is associated with a promotion engine other than CPE.  Can not proceed.", True)
    Send_Response_Header("This location is associated with a promotion engine other thanCPE.  Can not proceed.", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
  Else
    Mode = UCase(Request.QueryString("mode"))
    Common.Write_Log(LogFile, "** " & Common.AppName & "   " & Microsoft.VisualBasic.DateAndTime.Now & "  CSVersion: " & Connector.CSMajorVersion & "." & Connector.CSMinorVersion & "b" & Connector.CSBuild & "r" & Connector.CSBuildRevision & "  LSVersion=" & LSVerMajor & "." & LSVerMinor & "b" & LSBuildMajor & "." & LSBuildMinor & "  Serial: " & LocalServerID & "   LocationID: " & LocationID & "  Mode: " & Mode & "  Process running on server:" & Environment.MachineName, True)
    Common.Write_Log(LogFile, "SessionID=" & IPLSessionID & "    StartTable=" & StartTable & "    StartPK=" & StartPK, True)
    If Mode = "ACK" Then
      Process_ACK(LocalServerID, LocationID)
    ElseIf Mode = "NAK" Then
      Process_NAK(LocalServerID, LocationID)
    ElseIf (Mode = "IPL") Then
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
        Common.LRTsp.Parameters.Add("@StartTable", SqlDbType.VarChar, 200).Value = StartTable
        Common.LRTsp.Parameters.Add("@StartPK", SqlDbType.BigInt).Value = StartPK
        Common.LRTsp.Parameters.Add("@SessionID", SqlDbType.VarChar, 50).Value = IPLSessionID
        Common.LRTsp.Parameters.Add("@StartTime", SqlDbType.DateTime).Direction = ParameterDirection.Output
        Common.LRTsp.ExecuteNonQuery()
        HistoryStartTime = Common.LRTsp.Parameters("@StartTime").Value
        Common.Close_LRTsp()
        
        If StartTable = "" Then  'only perform these actions the first time IPL-PG is called
          Common.Write_Log(LogFile, "Removing buffered TransDownload data", True)
          Common.QueryStr = "dbo.pc_CPE_IPLPurge_Customers_Output"
          Common.Open_LXSsp()
          Common.LXSsp.CommandTimeout = 1200
          Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
          Common.LXSsp.ExecuteNonQuery()
          Common.Close_LXSsp()
          TotalTime = DateAndTime.Timer - StartTime
          Common.Write_Log(LogFile, "Finished removing buffered TransDownload data.  Total elapsed time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".000") & "(sec)", True)
        End If

        If Not (ZipOutput) Then
          gzStream = New GZipStream(Response.OutputStream, CompressionMode.Compress, True)

          Construct_Output(LocalServerID, LocationID)
          OutStr = ""
          If (LSVerMajor > 5) Or (LSVerMajor = 5 And LSVerMinor > 4) Or (LSVerMajor = 5 And LSVerMinor = 4 And LSBuildMajor >= 6) Then
            OutStr = OutStr & "LastTable: " & LastTable & vbCrLf
            OutStr = OutStr & "LastPK: "
            If Not (LastPK = 0) Then
              OutStr = OutStr & LastPK
            End If
            SD(OutStr)
          End If
          TotalTime = DateAndTime.Timer - StartTime
          Common.Write_Log(LogFile, "Finished All Queries=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".000") & "(sec)", True)
                
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
          
        Else
          ' ----------------------------------------------------------------------------------------------------------------------------------------
          FileStamp = LocalServerID & Right(Format(DateAndTime.Timer, "000"), 3)  'declared globally
          IPLTempPath = Trim(Common.Fetch_CPE_SystemOption(63))
          If IPLTempPath = "" Then
            IPLTempPath = Trim(Common.Fetch_SystemOption(29)) 'get the workspace file path and use that instead
          End If
          If Not (Right(IPLTempPath, 1) = "\") Then IPLTempPath = IPLTempPath & "\"
          DataFile = IPLTempPath & ApplicationName & "-" & FileStamp & ".txt"
        
          FileNum = FreeFile()
          Common.Write_Log(LogFile, "Creating temporary file: " & DataFile, True)
          FileOpen(FileNum, DataFile, OpenMode.Output)
          Construct_Output(LocalServerID, LocationID)
          FileClose(FileNum)
          TotalTime = DateAndTime.Timer - StartTime
          Common.Write_Log(LogFile, "Finished All Queries=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".000") & "(sec)", True)
        
          'sleep for half a second
          'System.Threading.Thread.Sleep(500)
        
          'Build the file header
          OutStr = ""
          If (LSVerMajor > 5) Or (LSVerMajor = 5 And LSVerMinor > 4) Or (LSVerMajor = 5 And LSVerMinor = 4 And LSBuildMajor >= 6) Then
            OutStr = OutStr & ApplicationName & "," & Connector.CSMajorVersion & "," & Connector.CSMinorVersion & "," & Connector.CSBuild & "," & Connector.CSBuildRevision & vbCrLf
            OutStr = OutStr & "LastTable: " & LastTable & vbCrLf
            OutStr = OutStr & "LastPK: "
            If Not (LastPK = 0) Then
              OutStr = OutStr & LastPK
            End If
            OutStr = OutStr & vbCrLf
            OutStr = OutStr & "LocationID=" & LocationID & vbCrLf
          End If
                
          'open the file and read the contents
          TextData = OutStr & My.Computer.FileSystem.ReadAllText(DataFile)
        
          'delete the data file
          My.Computer.FileSystem.DeleteFile(DataFile)
        
          If ZipOutput Then
            UncompressedSize = Len(TextData)
            Common.Write_Log(LogFile, "Starting GZip compression ... size before zipping is " & Format(UncompressedSize, "###,###,###,###,##0") & " bytes", True)
            CompressedArray = Encoding.Default.GetBytes(GZIP.CompressString(TextData))
            CompressedSize = UBound(CompressedArray) + 1
            Common.Write_Log(LogFile, "GZip compression successful ... size after zipping is " & Format(CompressedSize, "###,###,###,###,##0") & " bytes", True)
            TotalTime = DateAndTime.Timer - StartTime
            Common.Write_Log(LogFile, "Starting to return data to Local Server - Total elapsed time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".000") & "(sec)", True)
            Response.BinaryWrite(CompressedArray)
            CompressedArray = Nothing
            TotalTime = DateAndTime.Timer - StartTime
            Common.Write_Log(LogFile, "Finished returning data to Local Server - Total elapsed time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".000") & "(sec)", True)
          Else
            Sendb(TextData)
          End If
          TextData = Nothing
          ' ----------------------------------------------------------------------------------------------------------------------------------------
        End If
        
        'update the history record for the IPL end time
        Common.QueryStr = "dbo.pa_IPL_HistoryEnd"
        Common.Open_LRTsp()
        Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
        Common.LRTsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = LocationID
        Common.LRTsp.Parameters.Add("@IPLTypeID", SqlDbType.Int).Value = IPLTypeID
        Common.LRTsp.Parameters.Add("@StartTime", SqlDbType.DateTime).Value = HistoryStartTime
        Common.LRTsp.Parameters.Add("@UncompressedSize", SqlDbType.BigInt).Value = UncompressedSize
        If Not (ZipOutput) Then
          Common.LRTsp.Parameters.Add("@CompressedSize", SqlDbType.BigInt).Value = 0
        Else
          Common.LRTsp.Parameters.Add("@CompressedSize", SqlDbType.BigInt).Value = CompressedSize
        End If
        Common.LRTsp.ExecuteNonQuery()
        Common.Close_LRTsp()
      End If
    Else  'no matches for MODE=
        Send_Response_Header("Invalid Request", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
    End If
  End If 'locationid="0"
  TotalTime = DateAndTime.Timer - StartTime
  Common.Write_Log(LogFile, "Closing database connections - Total elapsed time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".000") & "(sec)", True)
  
  Common.Close_LogixRT()
  Common.Close_LogixXS()
  
  TotalTime = DateAndTime.Timer - StartTime
  Common.Write_Log(LogFile, "RunTime=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".000") & "(sec)", True)
%>
<% 
  Response.End()
ErrorTrap:
  Response.Write(Common.Error_Processor(, "Process Info: Server Name=" & Environment.MachineName & " serial=" & LocalServerID / 0 & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " Invoking LocationID=" & LocationID, , Common.InstallationName))
  Common.Write_Log(LogFile, "***** An error occurred during processing! Please check the ErrorLog! *****", True)
  If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
  If Not (Common.LXSadoConn.State = ConnectionState.Closed) Then Common.Close_LogixXS()
  Common = Nothing
%>
