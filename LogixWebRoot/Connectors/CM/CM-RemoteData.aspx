<%@ Page CodeFile="ConnectorCB.vb" Debug="true" Inherits="ConnectorCB" Language="vb" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: CM-RemoteData.aspx
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
  
  Sub Handle_Post(ByVal LocalServerID As Long, ByVal LocationID As Long)
    
    Dim UploadedData As String
    Dim FileData() As Byte
    Dim Checksum As String = ""
    Dim dst As DataTable
    Dim PrevMD5 As String = ""
    Dim FilePath As String
    Dim FileName As String = ""
    Dim RemoteDataType As Integer
    Dim FileVersion As String
    Dim FileStrData As String = ""
    

    Send_Response_Header("CM-RemoteData", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
    'See what type of data is being uploaded
    RemoteDataType = Common.Extract_Val(Trim(Get_Page_Value("rdtype")))
    If RemoteDataType = 0 Then
      Common.Write_Log(LogFile, "rdtype was not specified")
      Send("rdtype not specified")
      Exit Sub
    End If
    
    'This line is here to work around an issue.  The CM local servers
    'are incorrectly sending rdtype as 1 instead of 2.  This will be
    'fixed in a later rev of the CM store level code.  After that
    'fix is in place, then this line can be removed.  - MM 6/3/08
    'If RemoteDataType = 1 Then RemoteDataType = 2
    
    FileVersion = Common.Extract_Val(Trim(Get_Page_Value("ver")))
    
    FilePath = ""
    
    If (RemoteDataType = 1) Then
      FilePath = ""
      FilePath = Common.Fetch_SystemOption(29)
      If FilePath = "" Then
        Common.Error_Processor("Serial=" & LocalServerID & " IPAddress=" & Trim(Request.UserHostAddress) & " SQL Bulk Insert Path not set!")
        Exit Sub
      End If
      If Not (FilePath.Substring(FilePath.Length - 1, 1) = "\") Then
        FilePath = FilePath & "\"
      End If
      Dim TimeStamp As String = Common.Leading_Zero_Fill(Year(Today), 4) & Common.Leading_Zero_Fill(Month(Today), 2) & Common.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & Common.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Hour(Now()), 2) & Common.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Minute(Now()), 2) & Common.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Second(Now()), 2)
      FileName = "ID-" & LocationID & "-" & TimeStamp & ".txt"
    Else
      'Fetch the filepath from LogixRT for this data type
      Common.QueryStr = "select OutputPath from RemoteDataOptions as RDO with (NoLock) where RemoteDataTypeID=" & RemoteDataType & ";"
      dst = Common.LRT_Select
      If dst.Rows.Count > 0 Then
        FilePath = Trim(dst.Rows(0).Item("OutputPath"))
      End If
      dst = Nothing
      If FilePath = "" Then
        Common.Write_Log(LogFile, "rdtype invalid or OutputPath not specified.  rdtype=" & RemoteDataType)
        Send("rdtype is invalid or OutputPath not specified")
        Exit Sub
      End If
      If Not (Right(FilePath, 1) = "\") Then
        FilePath = FilePath & "\"
      End If
    End If
    
    UploadedData = ""
    If Request.Files.Count > 0 Then
      If RemoteDataType = 1 Then
        ReDim FileData(Request.Files(0).ContentLength - 1)
        Request.Files(0).InputStream.Read(FileData, 0, Request.Files(0).ContentLength)
        'If StyleTypeID = 0 Then
        'Else
        '  If FileName = "" Then 'if we weren't explicity told what the filename was in the fname parameter, then use the filename in the file collection
        '    FileName = Request.Files(0).FileName
        '  End If
        'End If
        Dim CompressedData As String = Encoding.Default.GetString(FileData)
        Common.Write_Log(LogFile, "File receipt successful ... file size is " & Format(Len(CompressedData), "###,###,###,###,##0") & " bytes")
        Common.Write_Log(LogFile, "ver (file version) = " & FileVersion)
        FileData = Nothing
        FileStrData = GZIP.DecompressString(CompressedData)
        CompressedData = Nothing
        FileData = Nothing
      Else
        ReDim FileData(Request.Files(0).ContentLength - 1)
        Request.Files(0).InputStream.Read(FileData, 0, Request.Files(0).ContentLength)
        If FileName = "" Then FileName = Request.Files(0).FileName
        'uncomment to view raw data
        'Send(Encoding.Default.GetString(FileData))
        'UploadedData = Encoding.Default.GetString(FileData)
      End If
    Else
      Common.Write_Log(LogFile, "No files were uploaded")
      Send("No files were uploaded")
      Exit Sub
    End If
    Common.Write_Log(LogFile, "File receipt successful ... file size is " & Format(Len(UploadedData), "###,###,###,###,##0") & " bytes")
    
    If Not (MD5 = "") Then
      'if we got an MD5 checksum, make sure it matches the commpressed data we recieved
      Checksum = Common.MD5(Encoding.Default.GetString(FileData))
      If Checksum <> MD5 Then
        Common.Write_Log(LogFile, "Bad MD5 .. LocalServer sent ->" & MD5 & "     CentralServer computed ->" & Checksum)
        Send("Bad MD5")
        Exit Sub
      End If
      Common.Write_Log(LogFile, "Valid MD5 ->" & MD5)
      
      'see if this MD5 is the same as the last set of data we processed
      Common.QueryStr = "select RemoteDataMD5 from LocalServers with (NoLock) where LocalServerID=" & LocalServerID & ";"
      dst = Common.LRT_Select
      If dst.Rows.Count > 0 Then
        PrevMD5 = Common.NZ(dst.Rows(0).Item("RemoteDataMD5"), "")
      End If
      dst = Nothing
      Common.Write_Log(LogFile, "Previous MD5 ->" & PrevMD5)
      If PrevMD5 = MD5 Then
        'this file was previously processed
        Common.Write_Log(LogFile, "Dup MD5 .. LocalServer sent ->" & MD5 & "     This matched the MD5 of the previously received file!")
        Send("Dup MD5")
        Exit Sub
      End If
    End If  'End - we recieved an MD5 from the local server
    
    If RemoteDataType = 1 Then
      'write the file out to the drive here
      Dim FileNum As Integer = FreeFile()
      FileOpen(FileNum, FilePath & FileName, OpenMode.Output)
      Print(FileNum, FileStrData)
      FileClose(FileNum)
      
      'this is Issuance data that needs to be inserted into the LogixEX database
      Common.Open_LogixEX()
      Common.QueryStr = "dbo.pt_IssuanceInsertQueue_Insert"
      Common.Open_LEXsp()
      Common.LEXsp.Parameters.Add("@FileName", SqlDbType.NVarChar, 255).Value = FileName
      Common.LEXsp.Parameters.Add("@FileVersion", SqlDbType.VarChar, 10).Value = FileVersion
      Common.LEXsp.ExecuteNonQuery()
      Common.Close_LEXsp()
      Common.Close_LogixEX()
      Common.Write_Log(LogFile, "Queued file " & FilePath & FileName)
    Else
      'write the file out to the drive here
      My.Computer.FileSystem.WriteAllBytes(FilePath & FileName, FileData, False)
      Common.Write_Log(LogFile, "Wrote data to file: " & FilePath & FileName)
      My.Computer.FileSystem.WriteAllText(FilePath & FileName & ".ok", FilePath & FileName, False)
    End If
    
    If Not (MD5 = "") Then  'if we got an MD5 then
      Common.QueryStr = "update LocalServers with (RowLock) set RemoteDataMD5='" & Checksum & "' where LocalServerID=" & LocalServerID & ";"
      Common.LRT_Execute()
    End If
    
    Send("ACK")
    
  End Sub
</script>

<%
  '-----------------------------------------------------------------------------------------------
  'Main Code - Execution starts here
  
  Dim ExtLocationCode As String
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
  Dim TempLocationID As Long
  
  Common.AppName = "CM-RemoteData.aspx"
  Response.Expires = 0
  On Error GoTo ErrorTrap
  StartTime = Microsoft.VisualBasic.DateAndTime.Timer
  IPAddress = Request.UserHostAddress
  
  LastHeard = "1/1/1980"
  ExtLocationCode = Trim(Get_Page_Value("extlocationcode"))
  LocalServerID = 0
  LocationID = 0
  MD5 = Trim(Get_Page_Value("md5"))
  LogFile = "CM-RemoteDataLog-" & Format(LocalServerID, "00000") & "." & Common.Leading_Zero_Fill(Year(Today), 4) & Common.Leading_Zero_Fill(Month(Today), 2) & Common.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
  
  Common.Open_LogixRT()
  Common.Load_System_Info()
  Connector.Load_System_Info(Common)
  
  Common.Write_Log(LogFile, "----------------------------------------------------------------")
  Common.Write_Log(LogFile, "** " & Common.AppName & "  -  " & Microsoft.VisualBasic.DateAndTime.Now & "  Serial: " & LocalServerID)
  
  ProcessOK = True
  SerialOK = False
  
  If ProcessOK Then
    If ExtLocationCode = "" Then
      Send_Response_Header("No ExtLocationCode specified", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
      Common.Write_Log(LogFile, "No ExtLocationCode specified from IP(" & IPAddress & ")")
      ProcessOK = False
    End If
  End If
  
  If ProcessOK Then
    Common.QueryStr = "dbo.pa_CM_Gen_CheckExtLocationCode"
    Common.Open_LRTsp()
    Common.LRTsp.Parameters.Add("@ExtLocationCode", SqlDbType.NVarChar).Value = ExtLocationCode
    dst = Common.LRTsp_select
    Common.Close_LRTsp()
    If dst.Rows.Count > 0 Then
      LocalServerID = Common.NZ(dst.Rows(0).Item("LocalServerID"), 0)
      LocationID = Common.NZ(dst.Rows(0).Item("LocationID"), 0)
    End If
    dst = Nothing
    If LocalServerID = 0 Or LocationID = 0 Then
      Send_Response_Header("Invalid LocationID or no associated LocalServerID for ExtLocationCode '" & ExtLocationCode & "'", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
      Common.Write_Log(LogFile, "Invalid LocationID or no associated LocalServerID for ExtLocationCode '" & ExtLocationCode & "' from IP(" & IPAddress & ")")
      ProcessOK = False
    End If
  End If
  
 
  If ProcessOK Then
    Common.Write_Log(LogFile, "ExtLocationCode='" & ExtLocationCode & "'   LocationID=" & LocationID & "   LocalServerID=" & LocalServerID)
      
    Handle_Post(LocalServerID, LocationID)
      
    TotalTime = Microsoft.VisualBasic.DateAndTime.Timer - StartTime
    Common.Write_Log(LogFile, "Finished Processing")
      
    Common.QueryStr = "dbo.pa_CPE_RemoteData_SetLastUpdate"
    Common.Open_LRTsp()
    Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
    Common.LRTsp.ExecuteNonQuery()
    Common.Close_LRTsp()
      
    TotalTime = Microsoft.VisualBasic.DateAndTime.Timer - StartTime
    Common.Write_Log(LogFile, "RunTime=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
  End If 'ProcessOK
  
  If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
%>
<%
  Response.End()
ErrorTrap:
  Response.Write(Common.Error_Processor(, "Process Info: Server Name=" & Environment.MachineName & " Invoking LocationID=" & LocationID, , Common.InstallationName))
  Common.Write_Log(LogFile, "***** An error occurred during processing! Please check the ErrorLog! *****")
  If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
  Common = Nothing
%>
