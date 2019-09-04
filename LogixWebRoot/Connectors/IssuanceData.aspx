<%@ Page CodeFile="ConnectorCB.vb" Debug="true" Inherits="ConnectorCB" Language="vb" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: IssuanceData.aspx
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

  Public Function CountCharacter(ByVal value As String, ByVal ch As Char) As Integer
    Dim cnt As Integer = 0
    For Each c As Char In value
      If c = ch Then cnt += 1
    Next
    Return cnt
  End Function
  ' -----------------------------------------------------------------------------------------------

  ' Line#  Contains
  '   1:    TableName
  '   2:    1=Insert/Update  2=Delete  3=Image 4=CentralKey update
  '   3:    Column List
  '   4:    Row Count
  '  Rows of data follow

  'If Line# 2 is a 4 and the value for the CentralKey is a zero then the LocalServer should delete
  'the row with that corresponding LocalID - this is because the data contained in that record
  'is a duplicate of another record at the CentralServer - this should only occur rarely (if at all)

  Sub Handle_Post(ByVal LocalServerID As Long, ByVal LocationID As Long)

    Dim FileData() As Byte
    Dim FileStrData As String
    Dim CompressedData As String
    Dim Checksum As String = ""
    Dim dst As DataTable
    Dim PrevMD5 As String = ""
    Dim FilePath As String
    Dim FileName As String
    Dim RemoteDataType As Integer
    Dim StyleTypeID As Integer
    Dim StartPoint As Long
    Dim FileNum As Integer
    Dim TimeStamp As String
    Dim FileVersion As String
    Dim DateValue As Date
    Dim strGUID As String

    Send_Response_Header("IssuanceData", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
    FileName = ""
    'See what type of data is being uploaded
    FileVersion = Common.Extract_Val(Trim(Get_Page_Value("ver")))
    RemoteDataType = Common.Extract_Val(Trim(Get_Page_Value("rdtype")))
    StyleTypeID = Common.Extract_Val(Trim(Get_Page_Value("style")))
    If RemoteDataType = 0 Then
      Common.Write_Log(LogFile, "rdtype was not specified")
      Send("rdtype not specified")
      Exit Sub
    End If

    If StyleTypeID = 0 Then
      'multiple styles are included in the payload
      'since there are multiple types, there's no file path to fetch
      'data will be inserted into the database - Issuance (LogixEX)

      FilePath = ""
      FilePath = Common.Fetch_SystemOption(29)
      If FilePath = "" Then
        Common.Error_Processor("Serial=" & LocalServerID & " IPAddress=" & Trim(Request.UserHostAddress) & " SQL Bulk Insert Path not set!")
        Exit Sub
      End If
      If Not (FilePath.Substring(FilePath.Length - 1, 1) = "\") Then
        FilePath = FilePath & "\"
      End If
      DateValue = Date.Now
      TimeStamp = DateValue.ToString("yyyyMMddHHmmss")
      strGUID = System.Guid.NewGuid.ToString()
      FileName = "ID-" & LocationID & "-" & TimeStamp & "-" & strGUID & ".txt"
    Else

      FilePath = ""
      'Fetch the filepath from LogixRT for this data type
      Common.QueryStr = "select isnull(OutputPath, '') as OutputPath from RemoteDataOptions as RDO with (NoLock) where RemoteDataTypeID=@RemoteDataType;"
      Common.DBParameters.Add("@RemoteDataType", SqlDbType.Int).Value = RemoteDataType
      dst = Common.ExecuteQuery(Copient.DataBases.LogixRT)
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

      FileName = Get_Page_Value("fname")
    End If

    If Request.Files.Count > 0 Then
      ReDim FileData(Request.Files(0).ContentLength - 1)
      Request.Files(0).InputStream.Read(FileData, 0, Request.Files(0).ContentLength)
      If StyleTypeID = 0 Then
      Else
        If FileName = "" Then 'if we weren't explicitly told what the filename was in the fname parameter, then use the filename in the file collection
          FileName = Request.Files(0).FileName
        End If
      End If
      CompressedData = Encoding.Default.GetString(FileData)
      Common.Write_Log(LogFile, "File receipt successful ... file size is " & Format(Len(CompressedData), "###,###,###,###,##0") & " bytes")
      Common.Write_Log(LogFile, "ver (file version) = " & FileVersion)
      FileData = Nothing
      FileStrData = GZIP.DecompressString(CompressedData)
      CompressedData = Nothing
      FileData = Nothing
    Else
      Common.Write_Log(LogFile, "No files were uploaded")
      Send("No files were uploaded")
      Exit Sub
    End If
    Common.Write_Log(LogFile, "GZip decompression successful ... size after unzipping is " & Format(Len(FileStrData), "###,###,###,###,##0") & " bytes")
    Common.Write_Log(LogFile, "File Contents: " & vbCrLf & FileStrData)
    If FileStrData = "no data" Then
      Common.Write_Log(LogFile, "Server had 'no data' to upload - no processing performed")
      Send("ACK")
      Exit Sub
    End If

    If Not (MD5 = "") Then
      'if we got an MD5 checksum, make sure it matches the compressed data we received
      Checksum = Common.MD5(FileStrData)
      If Checksum <> MD5 Then
        Common.Write_Log(LogFile, "Bad MD5 .. LocalServer sent ->" & MD5 & "     CentralServer computed ->" & Checksum)
        Send("Bad MD5")
        Exit Sub
      End If
      Common.Write_Log(LogFile, "Valid MD5 ->" & MD5)

      'see if this MD5 is the same as the last set of data we processed
      Common.QueryStr = "select RemoteDataMD5 from LocalServers with (NoLock) where LocalServerID=@LocalServerID;"
      Common.DBParameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
      dst = Common.ExecuteQuery(Copient.DataBases.LogixRT)
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
    End If  'End - we received an MD5 from the local server

    If Not (MD5 = "") Then 'strip out the first line (which should be the date stamp)
      StartPoint = InStr(FileStrData, vbCrLf, vbBinaryCompare) + 2
      If Len(FileStrData) > StartPoint Then
        FileStrData = Mid(FileStrData, StartPoint)
      Else
        Common.Write_Log(LogFile, "The file contained nothing but the timestamp - no processing performed")
        Send("ACK")
        Exit Sub
      End If
    End If

    'write the file out to the drive here
    FileNum = FreeFile()
    FileOpen(FileNum, FilePath & FileName, OpenMode.Output)
    Print(FileNum, FileStrData)
    FileClose(FileNum)

    Common.Write_Log(LogFile, "Wrote data to file: " & FilePath & FileName & "  (" & Format(Len(FileStrData), "###,###,###,###,##0") & " bytes)")
    If RemoteDataType = 1 Then
      'This came up because of a bug in issuance fmt files. During v6 there were 38 fields in total, and later a new field was added without
      'incrementing the version number of fmt file. So, had to add this check to count the number of fields and select the fmt file based on
      'count.
      If FileVersion = 6
        Dim S As String = FileStrData.Split(new String() {Environment.NewLine}, StringSplitOptions.None)(0)
        If CountCharacter(S, ","C) = 38
          FileVersion = "6.1"
        End If          
      End If      
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
    End If

    Common.QueryStr = "update LocalServers with (RowLock) set RemoteDataMD5=@Checksum where LocalServerID=@LocalServerID;"
    Common.DBParameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
    Common.DBParameters.Add("@Checksum", SqlDbType.NVarChar).Value = Checksum

    Common.ExecuteNonQuery(Copient.DataBases.LogixRT)

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
  Dim Mode As String
  Dim dst As DataTable
  Dim ProcessOK As Boolean
  Dim SerialOK As Boolean
  'Dim MustIPL As Boolean
  Dim TempLocationID As Long
  Dim BannerID As Integer

  Common.AppName = "IssuanceData.aspx"
  Response.Expires = 0
  On Error GoTo ErrorTrap
  StartTime = DateAndTime.Timer
  'IPAddress = Request.UserHostAddress

  LocalServerID = Common.Extract_Val(Request.QueryString("serial"))
  LogFile = "IssuanceDataLog-" & Format(LocalServerID, "00000") & "." & Common.Leading_Zero_Fill(Year(Today), 4) & Common.Leading_Zero_Fill(Month(Today), 2) & Common.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
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
  LocalServerID = Common.Extract_Val(Get_Page_Value("serial"))
  LSVersion = Common.Extract_Val(Get_Page_Value("lsversion"))
  LSBuild = Common.Extract_Val(Get_Page_Value("lsbuild"))
  MD5 = Trim(Get_Page_Value("md5"))
  Mode = UCase(Get_Page_Value("mode"))
  If Mode = "" Then Mode = "FETCH"

  Common.Open_LogixRT()
  Common.Load_System_Info()
  Connector.Load_System_Info(Common)

  Common.Write_Log(LogFile, "----------------------------------------------------------------")
  Common.Write_Log(LogFile, "** " & Common.AppName & "  -  " & Microsoft.VisualBasic.DateAndTime.Now & "  Serial: " & LocalServerID & " MacAddress: " & MacAddress & " IP:" & LocalServerIP & "   Mode: " & Mode & "  Process running on server:" & Environment.MachineName)

  ProcessOK = True
  SerialOK = False
  If Not (Mode = "FETCH") Then
    Send_Response_Header("Invalid Mode", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
    Common.Write_Log(LogFile, "Received invalid mode: " & Mode & " from MacAddress:" & MacAddress & " IP:" & LocalServerIP & "  serial:" & LocalServerID & " server: " & Environment.MachineName)
    ProcessOK = False
  End If

  If ProcessOK Then
    Common.QueryStr = "dbo.pa_CPE_Gen_CheckSerial"
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
      Common.Write_Log(LogFile, "Returned: Invalid Serial: " & LocalServerID & " MacAddress:" & MacAddress & " IP:" & LocalServerIP & " Server: " & Environment.MachineName & vbCrLf)
      ProcessOK = False
    End If
  End If

  If ProcessOK Then
    Common.Write_Log(LogFile, "serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " server=" & Environment.MachineName)
    Connector.Get_LS_Info(Common, LocalServerID, LocationID, BannerID, LastHeard, 2, LocalServerIP)
    If Get_Page_Value("force") = "1" Then
      TempLocationID = Common.Extract_Val(Get_Page_Value("locationid"))
      If TempLocationID = 0 Then
        Send_Response_Header("Missing LocationID", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
        Common.Write_Log(LogFile, "Returned: Missing LocationID" & vbCrLf)
        ProcessOK = False
      Else
        'if the FORCE is coming from the server that is already servicing the location, don't change the LocationID
        If Not (TempLocationID = LocationID) Then
          'we need to change the LocationID so that the data will go back down to the store that is currently servicing this location
          LocationID = -1 * TempLocationID
        End If
      End If
    Else
      'CLOUDSOL-482: Avoid sending mustIPL status to CPE while upload
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
      ProcessOK = True

      'End If
    End If
  End If 'ProcessOK

  If ProcessOK Then
    Common.Write_Log(LogFile, "LocationID=" & LocationID)
    If LocationID = "0" Then
      Send_Response_Header("Invalid Serial Number", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
      Common.Write_Log(LogFile, "Received invalid LocationID parsed from MacAddress:" & MacAddress & " IP" & LocalServerIP & " Serial:" & LocalServerID & " server=" & Environment.MachineName)
    Else
      If Get_Page_Value("force") = "1" Then
        Common.Write_Log(LogFile, "Processing FORCED UPLOAD")
      End If

      Handle_Post(LocalServerID, LocationID)

      TotalTime = DateAndTime.Timer - StartTime
      Common.Write_Log(LogFile, "Finished Processing")

      Common.QueryStr = "dbo.pa_CPE_RemoteData_SetLastUpdate"
      Common.Open_LRTsp()
      Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
      Common.LRTsp.ExecuteNonQuery()
      Common.Close_LRTsp()

      TotalTime = DateAndTime.Timer - StartTime
      Common.Write_Log(LogFile, "RunTime=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
    End If
  End If 'ProcessOK

  If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
%>
<%
  Response.End()
ErrorTrap:
  Response.Write(Common.Error_Processor(, "Serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & "Process Info: Server Name=" & Environment.MachineName & " Invoking LocationID=" & LocationID, , Common.InstallationName))
  Common.Write_Log(LogFile, "***** An error occurred during processing! Please check the ErrorLog! *****")
  If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
  Common = Nothing
%>