<%@ Page CodeFile="..\ConnectorCB.vb" Debug="true" Inherits="ConnectorCB" Language="vb" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: ReportingUpload.aspx 
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
    Public MacAddress As String
    Public LocalServerIP As String
  Dim TextData As String
  Dim LogFile As String
  Dim MD5 As String
  Dim StartTime As Object
  
  ' -------------------------------------------------------------------------------------------------
  
  Sub SD(ByVal OutStr As String)
    OutboundBuffer.Append(OutStr & vbCrLf)
  End Sub
  
  ' -------------------------------------------------------------------------------------------------
  
  Sub SDb(ByVal OutStr As String)
    OutboundBuffer.Append(OutStr)
  End Sub
  
  ' -----------------------------------------------------------------------------------------------
  
  ' Line#  Contains
  '   1:    TableName
  '   2:    1=Insert/Update  2=Delete  3=Image 4=CentralKey update
  '   3:    Column List
  '   4:    Row Count
  '  Rows of data follow
  
  'If Line# 2 is a 4 and the value for the CentralKey is a zero then the LocalServer should delete
  'the row with that corresponding LocalID - this is becuase the data contained in that record
  'is a duplicate of another record at the CentralServer - this should only occur rarely (if at all)
  
  Sub Handle_Post(ByVal LocalServerID As Long, ByVal LocationID As Long, ByVal Mode As String)
    
    Dim CompressedData As String
    Dim InboundData As String
    Dim FileData() As Byte
    Dim Checksum As String = ""
    Dim dst As DataTable
    Dim PrevMD5 As String = ""
    Dim DataSize As Long
    Dim Index As Long
    Dim SkipLineOne As Boolean
    Dim FileNum As Integer
    Dim SqlBulkPath As String
    Dim FileName As String
    Dim TimeStamp As String
    Dim FileType As Integer
    Dim FileVersion As String
    
    Send_Response_Header("ReportingUpload", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
        LocalServerID = Common.Extract_Val(Request.QueryString("serial"))
        LocalServerIP = Trim(Request.QueryString("IP"))
        MacAddress = Trim(Request.QueryString("mac"))
        If MacAddress = "" Or MacAddress = "0" Then
            MacAddress = Trim(Request.UserHostAddress)
        End If
        If LocalServerIP = "" Or LocalServerIP = "0" Then
            LocalServerIP = MacAddress & " IP from requesting browser. "
        End If
    FileVersion = Common.Extract_Val(Get_Page_Value("ver"))
    If FileVersion = 0 Then FileVersion = 1
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
      'uncommnet to view decompressed data
      'Send(Inbounddata)
    Else
      Common.Write_Log(LogFile, "No files were uploaded")
      Send("No files were uploaded")
      Exit Sub
    End If
    Common.Write_Log(LogFile, "GZip decompression successful ... size after unzipping is " & Format(Len(InboundData), "###,###,###,###,##0") & " bytes")
    
    If Not (MD5 = "") Then
      'if we got an MD5 checksum, make sure it matches the decommpressed data we recieved
      Checksum = Common.MD5(InboundData)
      If Checksum <> MD5 Then
                Common.Write_Log(LogFile, "Serial: " & LocalServerID & " MacAddress:" & MacAddress & " IP:" & LocalServerIP & " Bad MD5 .. LocalServer sent ->" & MD5 & "     CentralServer computed ->" & Checksum & "server:" & Environment.MachineName)
        Send("Bad MD5")
        Exit Sub
      End If
      Common.Write_Log(LogFile, "Valid MD5 ->" & MD5)
      
      
      'see if this MD5 is the same as the last set of data we processed
      If Mode = "TRANSHISTORY" Then  'This is the old TransHistory from pre of v5.6
        Common.QueryStr = "select RUTHMD5 as MD5 from LocalServers with (NoLock) where LocalServerID=" & LocalServerID & ";"
      ElseIf Mode = "TRANSREDEMPTION" Then  'TransRedemption is new as of v5.6 (used to be TransHistory)
        Common.QueryStr = "select RUTHMD5 as MD5 from LocalServers with (NoLock) where LocalServerID=" & LocalServerID & ";"
      ElseIf Mode = "IMPRESSIONS" Then
        Common.QueryStr = "select RUIMPMD5 as MD5 from LocalServers with (NoLock) where LocalServerID=" & LocalServerID & ";"
      ElseIf Mode = "NEWTRANSHIST" Then  'This is really TransHistory as of v5.6
        Common.QueryStr = "select RUNTHMD5 as MD5 from LocalServers with (NoLock) where LocalServerID=" & LocalServerID & ";"
      Else
        Common.QueryStr = "select '' as MD5;"
      End If
      dst = Common.LRT_Select
      If dst.Rows.Count > 0 Then
        PrevMD5 = Common.NZ(dst.Rows(0).Item("MD5"), "")
      End If
      dst = Nothing
            Common.Write_Log(LogFile, "Serial: " & LocalServerID & " MacAddress:" & MacAddress & " IP:" & LocalServerIP & " Previous MD5 ->" & PrevMD5 & " server:" & Environment.MachineName)
      If PrevMD5 = MD5 Then
        'this file was previously processed
                Common.Write_Log(LogFile, "Serial: " & LocalServerID & " MacAddress:" & MacAddress & " IP:" & LocalServerIP & " Dup MD5 .. LocalServer sent ->" & MD5 & "     This matched the MD5 of the previously received file!" & " server=" & Environment.MachineName)
        Send("Dup MD5")
        Exit Sub
      End If
    End If  'End - we recieved an MD5 from the local server
    
    If MD5 = "" Then
      If InboundData = "no data" Then
                Common.Write_Log(LogFile, "Serial: " & LocalServerID & " MacAddress:" & MacAddress & " IP:" & LocalServerIP & " Server had no data to upload - no processing performed")
        Send("ACK")
        Exit Sub
      End If
    Else
      If Right(InboundData, 7) = "no data" Then
        If Mode = "TRANSHISTORY" Or Mode = "TRANSREDEMPTION" Then  'This is really TransRedemption as of v5.6
          Common.QueryStr = "update LocalServers with (RowLock) set RUTHMD5='" & Checksum & "' where LocalServerID=" & LocalServerID & ";"
        ElseIf Mode = "IMPRESSIONS" Then
          Common.QueryStr = "update LocalServers with (RowLock) set RUIMPMD5='" & Checksum & "' where LocalServerID=" & LocalServerID & ";"
        ElseIf Mode = "NEWTRANSHIST" Then  'This is really TransHistory as of v5.6
          Common.QueryStr = "update LocalServers with (RowLock) set RUNTHMD5='" & Checksum & "' where LocalServerID=" & LocalServerID & ";"
        End If
        Common.LRT_Execute()
                Common.Write_Log(LogFile, "Serial: " & LocalServerID & " IPAddress:" & Trim(Request.UserHostAddress) & " Server had no data to upload - no processing performed")
        Send("ACK")
        Exit Sub
      End If
    End If
    
    Common.Write_Log(LogFile, "ver (file version) = " & FileVersion)
        Common.Write_Log(LogFile, "Serial: " & LocalServerID & " MacAddress:" & MacAddress & " IP:" & LocalServerIP & " Received Data:" & vbCrLf & InboundData)
    
    DataSize = Len(InboundData)
    Index = 1
    SkipLineOne = False
    If Not (MD5 = "") Then SkipLineOne = True
    
    'strip the date from the first line of the file
    If SkipLineOne Then
      Index = InStr(InboundData, vbCrLf, CompareMethod.Binary)
      Index = Index + 2 'add two characters to move past the CRLF
      InboundData = Right(InboundData, DataSize - Index + 1)
    End If
    
    SqlBulkPath = Common.Fetch_SystemOption(29)
    If SqlBulkPath = "" Then
            Common.Error_Processor("Serial: " & LocalServerID & " MacAddress:" & MacAddress & " IP:" & LocalServerIP & " SQL Bulk Insert Path not set!")
      Exit Sub
    End If
    If Not (SqlBulkPath.Substring(SqlBulkPath.Length - 1, 1) = "\") Then
      SqlBulkPath = SqlBulkPath & "\"
    End If
    TimeStamp = Common.Leading_Zero_Fill(Year(Today), 4) & Common.Leading_Zero_Fill(Month(Today), 2) & Common.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & Common.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Hour(Now()), 2) & Common.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Minute(Now()), 2) & Common.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Second(Now()), 2)
    
    If Mode = "TRANSHISTORY" Or Mode = "TRANSREDEMPTION" Then  'This is really TransRedemption as of v5.6
      FileName = "TR-" & LocationID & "-" & TimeStamp & ".txt"
      
      'write the data to a file
      FileNum = FreeFile()
      FileOpen(FileNum, SqlBulkPath & FileName, OpenMode.Append)
      PrintLine(FileNum, InboundData)
      FileClose(FileNum)
            Common.Write_Log(LogFile, "Serial: " & LocalServerID & " MacAddress:" & MacAddress & " IP:" & LocalServerIP & " Wrote data to file: " & SqlBulkPath & FileName)
      
      If Mode = "TRANSHISTORY" Then
        FileType = 1
      Else
        FileType = 4
      End If
      'Add a record for this file to PromoMoveInsertQueue
      Common.QueryStr = "dbo.pa_PromoMoveQueue_Insert"
      Common.Open_LWHsp()
      Common.LWHsp.Parameters.Add("@FileName", SqlDbType.NVarChar, 255).Value = FileName
      Common.LWHsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = 2
      Common.LWHsp.Parameters.Add("@Filetype", SqlDbType.Int).Value = FileType
      Common.LWHsp.Parameters.Add("@FileVersion", SqlDbType.Int).Value = FileVersion
      Common.LWHsp.ExecuteNonQuery()
      Common.Close_LWHsp()
      
    ElseIf Mode = "IMPRESSIONS" Then
      FileName = "IMP-" & LocationID & "-" & TimeStamp & ".txt"
      
      'write the data to a file
      FileNum = FreeFile()
      FileOpen(FileNum, SqlBulkPath & FileName, OpenMode.Append)
      PrintLine(FileNum, InboundData)
      FileClose(FileNum)
            Common.Write_Log(LogFile, "Serial: " & LocalServerID & " MacAddress:" & MacAddress & " IP:" & LocalServerIP & " Wrote data to file: " & SqlBulkPath & FileName)
      
      'Add a record for this file to PromoMoveInsertQueue
      Common.QueryStr = "dbo.pa_PromoMoveQueue_Insert"
      Common.Open_LWHsp()
      Common.LWHsp.Parameters.Add("@FileName", SqlDbType.NVarChar, 255).Value = FileName
      Common.LWHsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = 2
      Common.LWHsp.Parameters.Add("@Filetype", SqlDbType.Int).Value = 2
      Common.LWHsp.Parameters.Add("@FileVersion", SqlDbType.Int).Value = FileVersion
      Common.LWHsp.ExecuteNonQuery()
      Common.Close_LWHsp()
    ElseIf Mode = "NEWTRANSHIST" Then
      FileName = "TH-" & LocationID & "-" & TimeStamp & ".txt"
      
      'write the data to a file
      FileNum = FreeFile()
      FileOpen(FileNum, SqlBulkPath & FileName, OpenMode.Append)
      PrintLine(FileNum, InboundData)
      FileClose(FileNum)
      Common.Write_Log(LogFile, "Wrote data to file: " & SqlBulkPath & FileName)
      
      'Add a record for this file to PromoMoveInsertQueue
      Common.QueryStr = "dbo.pa_PromoMoveQueue_Insert"
      Common.Open_LWHsp()
      Common.LWHsp.Parameters.Add("@FileName", SqlDbType.NVarChar, 255).Value = FileName
      Common.LWHsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = 2
      Common.LWHsp.Parameters.Add("@Filetype", SqlDbType.Int).Value = 3
      Common.LWHsp.Parameters.Add("@FileVersion", SqlDbType.Int).Value = FileVersion
      Common.LWHsp.ExecuteNonQuery()
      Common.Close_LWHsp()
      
    Else
      Send("Invalid Mode!")
      Exit Sub
    End If
    
      If Not (MD5 = "") Then  'if we got an MD5 then
        If Mode = "TRANSHISTORY" Then  'This is really TransRedemption as of v5.6
          Common.QueryStr = "update LocalServers with (RowLock) set RUTHMD5='" & Checksum & "' where LocalServerID=" & LocalServerID & ";"
          Common.LRT_Execute()
        ElseIf Mode = "IMPRESSIONS" Then
          Common.QueryStr = "update LocalServers with (RowLock) set RUIMPMD5='" & Checksum & "' where LocalServerID=" & LocalServerID & ";"
          Common.LRT_Execute()
        ElseIf Mode = "NEWTRANSHIST" Then
          Common.QueryStr = "update LocalServers with (RowLock) set RUNTHMD5='" & Checksum & "' where LocalServerID=" & LocalServerID & ";"
          Common.LRT_Execute()
        End If
      End If
    
      Send("ACK")
    
  End Sub
</script>

<%
  '-----------------------------------------------------------------------------------------------
  'Main Code - Execution starts here
  
  Dim LocalServerID As Long
  Dim LocationID As Long
  Dim LastHeard As String
  Dim TotalTime As Object
  Dim ZipOutput As Boolean
  Dim DataFile As String
  Dim ZipFile As String
  Dim FileStamp As String
  Dim Mode As String
  Dim RawRequest As String
  Dim Index As Long
  Dim IPAddress As String = ""
  Dim CompressedArray() As Byte
  Dim dst As DataTable
  Dim ProcessOK As Boolean
  Dim SerialOK As Boolean
  Dim MustIPL As Boolean
  Dim BannerID As Integer
  
  Common.AppName = "ReportingUpload.aspx"
  Response.Expires = 0
    On Error GoTo ErrorTrap
    
  StartTime = DateAndTime.Timer
    LocalServerID = Common.Extract_Val(Request.QueryString("serial"))
    LogFile = "CPE-TransUpdateLog-" & Format(LocalServerID, "00000") & "." & Common.Leading_Zero_Fill(Year(Today), 4) & Common.Leading_Zero_Fill(Month(Today), 2) & Common.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
    LocalServerIP = Trim(Request.QueryString("IP"))
    MacAddress = Trim(Request.QueryString("mac"))
    If MacAddress = "" Or MacAddress = "0" Then
        MacAddress = "0"
    End If
    If LocalServerIP = "" Or LocalServerIP = "0" Then
        Common.Write_Log(LogFile, "Could not get IP from query. Analyzing client request for IP ...")
        LocalServerIP = Trim(Request.UserHostAddress)
    End If
  LastHeard = "1/1/1980"
 
  LSVersion = Common.Extract_Val(Request.QueryString("lsversion"))
  LSBuild = Common.Extract_Val(Request.QueryString("lsbuild"))
  MD5 = Trim(Request.QueryString("md5"))
  Mode = UCase(Request.QueryString("mode"))    
  
  Common.Open_LogixRT()
  Common.Open_LogixWH()
  Common.Load_System_Info()
  Connector.Load_System_Info(Common)
  'Connector.Get_LS_Info(Common, LocalServerID, LocationID, BannerID, LastHeard, 2, LocalServerIP)
  
  Common.Write_Log(LogFile, "----------------------------------------------------------------")
    Common.Write_Log(LogFile, "** " & Common.AppName & "  -  " & Microsoft.VisualBasic.DateAndTime.Now & "  Serial: " & LocalServerID & " MacAddress: " & MacAddress & " IP:" & LocalServerIP & "   Mode: " & Mode & "  Process running on server:" & Environment.MachineName)
  
  ProcessOK = True
  SerialOK = False
  If Not (Mode = "TRANSHISTORY" Or Mode = "TRANSREDEMPTION" Or Mode = "IMPRESSIONS" Or Mode = "NEWTRANSHIST") Then
    Send_Response_Header("Invalid Mode", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
        Common.Write_Log(LogFile, "Serial: " & LocalServerID & " Received invalid mode" & Mode & " from MacAddress:" & MacAddress & " IP:" & LocalServerIP & "server:" & Environment.MachineName)
    ProcessOK = False
  End If
  
  If ProcessOK Then
    Common.QueryStr = "dbo.pa_Gen_CheckSerial"
    Common.Open_LRTsp()
    Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
    dst = Common.LRTsp_select
    Common.Close_LRTsp()
    If dst.Rows.Count > 0 Then
      If Common.NZ(dst.Rows(0).Item("NumRecs"), 0) > 0 Then SerialOK = True
    End If
    dst = Nothing
    If Not (SerialOK) Then
      Send_Response_Header("Invalid SerialNumber", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
            Common.Write_Log(LogFile, "Returned: Invalid Serial:" & LocalServerID & vbCrLf)
      ProcessOK = False
    End If
  End If
  
  If ProcessOK Then
    If Request.QueryString("force") = "1" Then
      LocationID = Common.Extract_Val(Request.QueryString("locationid"))
      If LocationID = 0 Then
        Send_Response_Header("Missing LocationID", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
        Common.Write_Log(LogFile, "Returned: Missing LocationID" & vbCrLf)
        ProcessOK = False
      Else
        'we need to change the LocationID so that the data will go back down to the store that is currently servicing this location
        LocationID = -1 * LocationID
      End If
    Else
            'Common.QueryStr = "dbo.pa_CPE_CheckMustIPL"
            'Common.Open_LRTsp()
            'Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
            'dst = Common.LRTsp_select
            'Common.Close_LRTsp()
            'If dst.Rows.Count > 0 Then
            '    MustIPL = Common.NZ(dst.Rows(0).Item("MustIPL"), True)
            'End If
            'dst = Nothing
      
            'If MustIPL Then
            '    Send_Response_Header("MustIPL", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
            '    Common.Write_Log(LogFile, "Returned: MustIPL")
        Connector.Get_LS_Info(Common, LocalServerID, LocationID, BannerID, LastHeard, 2, LocalServerIP)
            ProcessOK = True
            'Else
            '    Connector.Get_LS_Info(Common, LocalServerID, LocationID, BannerID, LastHeard, 2, LocalServerIP)
            'End If
    End If
    
  End If 'ProcessOK
  
  If ProcessOK Then

    Common.Write_Log(LogFile, "LocationID=" & LocationID)
    If LocationID = "0" Then

      Send_Response_Header("Invalid Serial Number", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
      Common.Write_Log(LogFile, "Serial: " & LocalServerID & " Received invalid LocationID:" & LocationID & " parsed from MacAddress:" & MacAddress)

    ElseIf Not (Request.QueryString("force") = "1") AndAlso Not (Connector.IsValidLocationForConnectorEngine(Common, LocationID, 2)) Then
      'the location calling ReportingUpload is not associated with the CPE promoengine
      Common.Write_Log(LogFile, "This location is associated with a promotion engine other than CPE.  Can not proceed.", True)
      Send_Response_Header("This location is associated with a promotion engine other than CPE.  Can not proceed.", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)

    Else

      If Request.QueryString("force") = "1" Then
        Common.Write_Log(LogFile, "Processing FORCED UPLOAD")
      End If
      
      Handle_Post(LocalServerID, LocationID, Mode)
      
      TotalTime = DateAndTime.Timer - StartTime
      Common.Write_Log(LogFile, "Finished Processing")
      
      Common.QueryStr = "dbo.pa_CPE_RU_SetLastUpdate"
      Common.Open_LRTsp()
      Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
      Common.LRTsp.ExecuteNonQuery()
      Common.Close_LRTsp()
      
      TotalTime = DateAndTime.Timer - StartTime
      Common.Write_Log(LogFile, "RunTime=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
      
    End If
  End If 'ProcessOK
  
  If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
  If Not (Common.LWHadoConn.State = ConnectionState.Closed) Then Common.Close_LogixWH()
%>
<%
  Response.End()
ErrorTrap:
    Response.Write(Common.Error_Processor(, "Serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " Process Info: Server Name=" & Environment.MachineName & " Invoking LocationID=" & LocationID, , Common.InstallationName))
  Common.Write_Log(LogFile, "***** An error occurred during processing! Please check the ErrorLog! *****")
  If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
  If Not (Common.LWHadoConn.State = ConnectionState.Closed) Then Common.Close_LogixWH()
  Common = Nothing
%>
