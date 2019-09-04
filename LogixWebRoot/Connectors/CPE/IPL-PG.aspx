<%@ Page CodeFile="..\ConnectorCB.vb" Debug="true" Inherits="ConnectorCB" Language="vb" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.IO.Compression" %>
<% 
  ' *****************************************************************************
  ' * FILENAME: IPL-PG.aspx 
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
  Public OperateAtEnterprise As Boolean
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
    Dim TempOut As String
    Dim FieldList As String
    Dim DataBack As String
    Dim OperationType As Integer
    Dim QueryStartTime As Decimal
    Dim QueryTotalTime As Decimal
    Dim ConstructStartTime As Decimal
    Dim ConstructTotalTime As Decimal
    Dim reader As SqlDataReader
    
	  Const BUFFERED_WRITE_SIZE As Integer = 10000
    'Send ("<!-- Table=" & TableName & " " & Format$(Time, "hh:mm:ss") & Format$(Timer - Fix(Timer), ".00") & " -->")
    
    TempOut = ""
    TempResults = ""
    NumRecs = 0
    FieldList = ""
    DataBack = ""
    
    If (UCase(TableName) = "PRODUCTGROUPS" And (Operation = 5)) Then
      DataBack = "(-77"
    End If
    
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
      If UCase(TableName) = "PRODGROUPITEMS" Then
        FieldList = "ProductGroupID" & Chr(DelimChar) & "ExtProductID" & Chr(DelimChar) & "ProductTypeID"
      ElseIf UCase(TableName) = "PRODUCTGROUPS" Then
        FieldList = "ProductGroupID" & Chr(DelimChar) & "GroupName" & Chr(DelimChar) & "AnyProduct" & Chr(DelimChar) & "PointsNotApplyGroup" & Chr(DelimChar) & "NonDiscountableGroup" & Chr(DelimChar) & "UpdateLevel"
      End If
      OperationType = Operation
      If OperationType = 99 Then OperationType = 2

      If (UCase(TableName) <> "PRODUCTGROUPS") Or (UCase(TableName) = "PRODUCTGROUPS" And StartTable = "") Then
        'send the table header
        TempOut = "1:" & TableName & vbCrLf
        TempOut = TempOut & "2:" & Trim(Str(OperationType)) & vbCrLf
        TempOut = TempOut & "3:" & FieldList
        SD(TempOut)
        Common.Write_Log(LogFile, TempOut)
      Else
        Common.Write_Log(LogFile, "Running query to determine ProductGroups related to this local server:" & LocalServerID)
      End If
      NumRecs = 0
      
	  Dim buf As New StringBuilder()
	  
      If UCase(TableName) = "PRODUCTGROUPS" And (Operation = 5) Then
        While reader.Read
          DataBack = DataBack & ", " & reader.Item("ProductGroupID")
          If StartTable = "" Then
			buf.AppendLine( reader.Item("ProductGroupID") & Chr(DelimChar) & Common.NZ(reader.Item("GroupName"), "Unknown") & Chr(DelimChar) & Parse_Bit(Common.NZ(reader.Item("AnyProduct"), 0)) & Chr(DelimChar) & Parse_Bit(Common.NZ(reader.Item("PointsNotApplyGroup"), 0)) & Chr(DelimChar) & Parse_Bit(Common.NZ(reader.Item("NonDiscountableGroup"), 0)) & Chr(DelimChar) & Common.NZ(reader.Item("UpdateLevel"), ""))
			If ( buf.Length > BUFFERED_WRITE_SIZE )
			  SDb( buf.ToString() )
			  buf.Clear()
			End If
          End If
          NumRecs = NumRecs + 1
        End While
        
      ElseIf UCase(TableName) = "PRODGROUPITEMS" And (Operation = 5) Then
        LastTable = "ProdGroupItems"
        While reader.Read
          buf.AppendLine( reader.Item("data") )
          If ( buf.Length > BUFFERED_WRITE_SIZE )
            SDb( buf.ToString() )
            buf.Clear()
          End If
          LastPK = reader.Item("PKID")
          NumRecs = NumRecs + 1
        End While
        If Not (NumRecs = MaxBatchSize) Or (MaxBatchSize = 0) Then
          LastTable = ""
          LastPK = 0
        End If
        DataBack = "Sent Data"
      End If

      If Not (UCase(TableName) = "PRODUCTGROUPS") Or (UCase(TableName) = "PRODUCTGROUPS" And StartTable = "") Then
        buf.AppendLine("###")
        SDb( buf.ToString() )
      End If
      Common.Write_Log(LogFile, "# records: " & NumRecs)
      If Not (LastTable = "") Then
        Common.Write_Log(LogFile, "LastTable=" & TableName & "   LastPK=" & LastPK)
      End If
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
    
    DataBack = DataBack & ")"
    
    Construct_Table = DataBack
    
  End Function
  
  ' -------------------------------------------------------------------------------------------------
  
  Sub Construct_Output(ByVal LocalServerID As Long, ByVal LocationID As Long)
  
    Dim OutStr As String
    Dim TempOut As String
    Dim DelimChar As Integer
    Dim ProductGroupIDStr As String
    Dim MaxBatchStr As String
    
    LastTable = ""
    LastPK = 0
    DelimChar = 30
    TempOut = ""
    
    If MaxBatchSize > 0 Then
      MaxBatchStr = "top " & MaxBatchSize
    Else
      MaxBatchStr = ""
    End If
    
    Common.Write_Log(LogFile, "Returned the following data:")
    OutStr = ""
    If Not (ZipOutput) Then
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
      Common.Write_Log(LogFile, OutStr)
    Else
      If (LSVerMajor < 5) Or (LSVerMajor = 5 And LSVerMinor < 4) Or (LSVerMajor = 5 And LSVerMinor = 4 And LSBuildMajor < 6) Then
        OutStr = OutStr & ApplicationName & "," & Connector.CSMajorVersion & "," & Connector.CSMinorVersion & "," & Connector.CSBuild & "," & Connector.CSBuildRevision & vbCrLf
        OutStr = OutStr & "D:" & vbCrLf
        OutStr = OutStr & "T:" & vbCrLf
        OutStr = OutStr & "LocationID=" & LocationID
        SD(OutStr)
        Common.Write_Log(LogFile, OutStr)
      End If
    End If
      
    
    'ProductGroups

    Common.QueryStr = "SELECT ProductGroupID, GroupName, AnyProduct, PointsNotApplyGroup, NonDiscountableGroup, UpdateLevel INTO #tmp_pg FROM ( " & _
       "select PG.ProductGroupID, PG.Name as GroupName, PG.AnyProduct, PG.PointsNotApplyGroup, PG.NonDiscountableGroup, PG.UpdateLevel " & _
       "from ProductGroups as PG with (NoLock) " & _
       "Inner Join ProductGroupLocUpdate as PGLU with (NoLock) " & _
       "on PG.ProductGroupID=PGLU.ProductGroupID " & _
       "and PG.Deleted=0 " & _
       "Where (PGLU.LocationID=" & LocationID & " and PGLU.EngineID=2) " & _
       "union " & _
       "select PG.ProductGroupID, PG.Name as GroupName, PG.AnyProduct, PG.PointsNotApplyGroup, PG.NonDiscountableGroup, PG.UpdateLevel " & _
       "from ProductGroups as PG with (NoLock) " & _
       "where PG.AnyProduct=1 " & _
       "or PG.PointsNotApplyGroup=1 " & _
       "or PG.NonDiscountableGroup=1 " & _
       ") as aa "
    Common.LRT_Execute()

    Common.QueryStr = "SELECT ProductGroupID, GroupName, AnyProduct, PointsNotApplyGroup, NonDiscountableGroup, UpdateLevel FROM #tmp_pg"
    ProductGroupIDStr = Construct_Table("ProductGroups", 5, DelimChar, LocationID, LocationID, "LRT")
    
    'send the Entire List of products for each of the product groups we just sent
    Common.QueryStr = "select distinct PGI.PKID, convert(nvarchar,PGI.ProductGroupID) + char(" & DelimChar & ") + P.ExtProductID + char(" & DelimChar & ") + convert(nvarchar,P.ProductTypeID)as data " & _
       "INTO #tmp_pgi " & _
       "From #tmp_pg as PG " & _
       "Inner Join ProdGroupItems as PGI with (NoLock) " & _
       "on PG.ProductGroupID=PGI.ProductGroupID " & _
          "and PGI.Deleted=0 " & _
            "Inner Join Products as P with (NoLock, Index=PK_Products_1) " & _
       "on PGI.ProductID=P.ProductID "
    Common.LRT_Execute()

    'send the Entire List of products for each of the product groups we just sent
    Common.QueryStr = "select " & MaxBatchStr & " PKID, data from #tmp_pgi AS PGI "
    If StartPK > 0 Then
      Common.QueryStr = Common.QueryStr & "where PGI.PKID>" & StartPK & " "
    End If
    Common.QueryStr = Common.QueryStr & "order by PGI.PKID;"
    OutStr = OutStr & Construct_Table("ProdGroupItems", 5, DelimChar, LocalServerID, LocationID, "LRT")
    
    SD("***") 'send the EOF marker
    
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
    Common.Write_Log(LogFile, " Serial:" & LocalServerID & " MacAddress:" & MacAddress & " IP:" & LocalServerIP & " Received NAK - ErrorMsg:" & ErrorMsg & " server:" & Environment.MachineName)
    Send_Response_Header(ApplicationName & " - NAK Received", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
    
  End Sub

</script>

<%
  
  Dim LocalServerID As Long
  Dim LocationID As Long
  Dim LastHeard As String
  Dim Mode As String
  Dim HistoryStartTime As DateTime
  Dim IPLTypeID As Integer
  Dim dst As DataTable
  Dim MustIPL As Boolean
  Dim OutStr As String
  Dim IPLSessionID As String
  Dim IPLTempPath As String
  Dim BannerID As Integer
  Dim LSVerParts() As String
  
  ' the following variables are used for GZip streaming with version 5.19 and above
  Dim DataFile As String
  Dim FileStamp As String
  Dim CompressedArray() As Byte
  Dim UncompressedSize As Long
  Dim CompressedSize As Long
  
  IPLTypeID = 2
  ApplicationName = "IPL-PG"
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
  StartTable = ""
  StartPK = 0
  IPLSessionID = "0"
  If (LSVerMajor > 5) Or (LSVerMajor = 5 And LSVerMinor > 4) Or (LSVerMajor = 5 And LSVerMinor = 4 And LSBuildMajor >= 6) Then
    StartTable = IIf(Request.QueryString("starttable") <> "", Request.QueryString("starttable"), "")
    StartPK = Common.Extract_Val(Request.QueryString("startpk"))
    IPLSessionID = IIf(Request.QueryString("sessionid") <> "", Request.QueryString("sessionid"), "0")
  End If
  
  Common.Open_LogixRT()
  Common.Load_System_Info()
  Connector.Load_System_Info(Common)
  
  Common.Write_Log(LogFile, "---------------------------------------------------------------------------")
  
  If (LSVerMajor > 5) Or (LSVerMajor = 5 And LSVerMinor > 4) Or (LSVerMajor = 5 And LSVerMinor = 4 And LSBuildMajor >= 6) Then
    MaxBatchSize = Common.Extract_Val(Common.Fetch_CPE_SystemOption(60))
  Else
    MaxBatchSize = 0
  End If

  ' determine if GZip streaming should be used based on Logix version (5.19 and above)
  ZipOutput = Not ((LSVerMajor > 5) OrElse ((LSVerMajor = 5) AndAlso (LSVerMinor >= 19)))
  
  OperateAtEnterprise = False
  If Common.Fetch_CPE_SystemOption(91) = "1" Then
    OperateAtEnterprise = True
  End If
  
  Connector.Get_LS_Info(Common, LocalServerID, LocationID, BannerID, LastHeard, 1, LocalServerIP)
  
  If LocationID = "0" Then
    Common.Write_Log(LogFile, Common.AppName & "  Invalid Serial Number " & LocalServerID & " from MacAddress: " & MacAddress & " IP:" & LocalServerIP & "  Process running on server:" & Environment.MachineName, True)
    Send_Response_Header(ApplicationName & " - Invalid Serial Number", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
  ElseIf Not (Connector.IsValidLocationForConnectorEngine(Common, LocationID, 2)) Then
    'the location calling IPL-PG is not associated with the CPE promoengine
    Common.Write_Log(LogFile, "This location is associated with a promotion engine other than CPE.  Can not proceed.", True)
    Send_Response_Header("This location is associated with a promotion engine other than CPE.  Can not proceed.", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
  Else
    Mode = UCase(Request.QueryString("mode"))
    Common.Write_Log(LogFile, "** " & Common.AppName & "   " & Microsoft.VisualBasic.DateAndTime.Now & "  CSVersion: " & Connector.CSMajorVersion & "." & Connector.CSMinorVersion & "b" & Connector.CSBuild & "r" & Connector.CSBuildRevision & "  LSVersion=" & LSVerMajor & "." & LSVerMinor & "b" & LSBuildMajor & "." & LSBuildMinor & "   Serial: " & LocalServerID & "   LocationID: " & LocationID & "  Mode: " & Mode & "  Process running on server:" & Environment.MachineName)
    Common.Write_Log(LogFile, "SessionID=" & IPLSessionID & "    StartTable=" & StartTable & "    StartPK=" & StartPK)
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
        Common.LRTsp.Parameters.Add("@StartTable", SqlDbType.VarChar, 255).Value = StartTable
        Common.LRTsp.Parameters.Add("@StartPK", SqlDbType.BigInt).Value = StartPK
        Common.LRTsp.Parameters.Add("@SessionID", SqlDbType.VarChar, 50).Value = IPLSessionID
        Common.LRTsp.Parameters.Add("@StartTime", SqlDbType.DateTime).Direction = ParameterDirection.Output
        Common.LRTsp.ExecuteNonQuery()
        HistoryStartTime = Common.LRTsp.Parameters("@StartTime").Value
        Common.Close_LRTsp()
        
        If StartTable = "" Then  'only perform these actions the first time IPL-PG is called
          'Purge any ProductGroups that are no longer used in active offers from ProductGroupLocUpdate for this location
          TotalTime = DateAndTime.Timer - StartTime
          Common.Write_Log(LogFile, "Removing expired ProductGroupLocUpdate data - Total elapsed time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".000") & "(sec)")
          If OperateAtEnterprise Then
            Common.QueryStr = "pa_Purge_Expired_CPE_ProductGroupLocUpdate_Enterprise"
          Else
            Common.QueryStr = "pa_Purge_Expired_CPE_ProductGroupLocUpdate"
          End If
          Common.Open_LRTsp()
          Common.LRTsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = LocationID
          Common.LRTsp.ExecuteNonQuery()
          Common.Close_LRTsp()
          TotalTime = DateAndTime.Timer - StartTime
          Common.Write_Log(LogFile, "Finished removing expired ProductGroupLocUpdate data - Total elapsed time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".000") & "(sec)")
        End If
        
        If Not (ZipOutput) Then
          gzStream = New GZipStream(Response.OutputStream, CompressionMode.Compress, True)
          Construct_Output(LocalServerID, LocationID)
          If (LSVerMajor > 5) Or (LSVerMajor = 5 And LSVerMinor > 4) Or (LSVerMajor = 5 And LSVerMinor = 4 And LSBuildMajor >= 6) Then
            SD("LastTable: " & LastTable)
            OutStr = "LastPK: "
            If Not (LastPK = 0) Then OutStr = OutStr & LastPK
            SD(OutStr)
          End If
          TotalTime = DateAndTime.Timer - StartTime
          Common.Write_Log(LogFile, "Finished All Queries=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".000") & "(sec)")
        
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
          'Create the temporary file name
          FileStamp = LocalServerID & Right(Format(DateAndTime.Timer, "000"), 3)  'declared globally
          IPLTempPath = Trim(Common.Fetch_CPE_SystemOption(63))
          If IPLTempPath = "" Then
            IPLTempPath = Trim(Common.Fetch_SystemOption(29)) 'get the workspace file path and use that instead
          End If
          If Not (Right(IPLTempPath, 1) = "\") Then IPLTempPath = IPLTempPath & "\"
          DataFile = IPLTempPath & ApplicationName & "-" & FileStamp & ".txt"
        
          FileNum = FreeFile()
          Common.Write_Log(LogFile, "Creating temporary file: " & DataFile)
          FileOpen(FileNum, DataFile, OpenMode.Output)
          Construct_Output(LocalServerID, LocationID)
          FileClose(FileNum)
          TotalTime = DateAndTime.Timer - StartTime
          Common.Write_Log(LogFile, "Finished All Queries=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".000") & "(sec)")
        
          'sleep for half a second
          System.Threading.Thread.Sleep(500)
        
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
            Common.Write_Log(LogFile, "Starting GZip compression ... size before zipping is " & Format(UncompressedSize, "###,###,###,###,##0") & " bytes")
            CompressedArray = Encoding.Default.GetBytes(GZIP.CompressString(TextData))
            CompressedSize = UBound(CompressedArray) + 1
            Common.Write_Log(LogFile, "GZip compression successful ... size after zipping is " & Format(CompressedSize, "###,###,###,###,##0") & " bytes")
            TotalTime = DateAndTime.Timer - StartTime
            Common.Write_Log(LogFile, "Starting to return data to Local Server - Total elapsed time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".000") & "(sec)")
            Response.BinaryWrite(CompressedArray)
            CompressedArray = Nothing
            TotalTime = DateAndTime.Timer - StartTime
            Common.Write_Log(LogFile, "Finished returning data to Local Server - Total elapsed time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".000") & "(sec)")
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
  Common.Write_Log(LogFile, "Closing database connections - Total elapsed time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".000") & "(sec)")
  
  Common.Close_LogixRT()
  
  TotalTime = DateAndTime.Timer - StartTime
  Common.Write_Log(LogFile, "RunTime=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".000") & "(sec)")
%>
<% 
  Response.End()
ErrorTrap:
  Response.Write(Common.Error_Processor(, "Serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " Process Info: Server Name=" & Environment.MachineName & " Invoking LocationID=" & LocationID, , Common.InstallationName))
  Common.Write_Log(LogFile, "***** An error occurred during processing! Please check the ErrorLog! *****")
  If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
  Common = Nothing
  Response.End()
%>
