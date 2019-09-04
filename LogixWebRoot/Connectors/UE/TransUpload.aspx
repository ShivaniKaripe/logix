<%@ Page CodeFile="..\ConnectorCB.vb" Debug="true" Inherits="ConnectorCB" Language="vb" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: TransUpload.aspx
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
    Public MyCryptLib As New Copient.CryptLib
  Public Connector As New Copient.ConnectorInc
  Public GZIP As New Copient.GZIPInc
  Dim LogFile As String
  Dim MD5 As String
  Dim StartTime As Object
  Dim LSVerMajor As Integer
  Dim LSVerMinor As Integer
  Dim ExternalIds As Dictionary(Of String, String) = New Dictionary(Of String, String)
  ' -----------------------------------------------------------------------------------------------

  Sub Clear_Waiting_ACK(ByVal LocalServerID As Long)
    Common.Write_Log(LogFile, "Clearing any data waiting ACK")
    'Get rid of any old data that might have been hanging out from a previously failed upload
    Common.QueryStr = "pa_CPE_TU_PurgeWaitingACK"
    Common.Open_LXSsp()
    Common.LXSsp.Parameters.Add("@WaitingACK", SqlDbType.Int).Value = LocalServerID
    Common.LXSsp.ExecuteNonQuery()
    Common.Close_LXSsp()

  End Sub

  ' -----------------------------------------------------------------------------------------------

  Sub Send_Failure_Response(ByVal ErrorMsg As String, ByVal LocalServerID As Long, ByVal LocationID As Long, ByVal LocalServerIP As String, ByVal MacAddress As String)

    ErrorMsg = "Processing not completed - Error occurred:  " & vbCrLf & ErrorMsg
    Common.Write_Log(LogFile, ErrorMsg)
    Send("NAK")
    Response.Write(Common.Error_Processor(, "Serial:" & LocalServerID & "  Process Info: server:" & Environment.MachineName & " Invoking LocationID=" & LocationID & "  Requester IP Address: " & LocalServerIP & "  Requester Mac Address: " & MacAddress & vbCrLf & ErrorMsg, , Common.InstallationName))

  End Sub
  
  Function GetExternalId (ByVal LocalId As String, ByVal ServerSerial As Integer, ByVal Create As Boolean) As String
    Dim ExternalId As String = ""    
    Dim Key As String = ServerSerial.ToString() & "," & LocalId
    If Create
      Dim epoch  As Date = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)  
      ExternalId = Convert.ToUInt64((DateAndTime.Now.ToUniversalTime() - epoch).TotalMilliseconds).ToString()  
      ExternalIds.Add(Key, ExternalId)
    Else
      ' Check if the record is newly created?
      Common.QueryStr = "SELECT ExternalId FROM " &
                        "StoredValue " &
                        "WHERE LocalId = " & LocalId & " AND ServerSerial = " & ServerSerial & " " &
                        "UNION ALL " &
                        "SELECT Col11 AS ExternalId FROM " &
                        "CPE_UploadTemp_SV " &
                        "WHERE Col1 = " & LocalId & " AND Col2 = " & ServerSerial
      Dim data As DataTable = Common.LXS_Select()
      If data.Rows.Count > 0
        ExternalId = data.Rows(0)(0).ToString()
      ElseIf ExternalIds.ContainsValue(Key) Then
        ExternalId = ExternalIds.Item(Key)
      End If
    End If
    Return ExternalId
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

  Sub Handle_Post(ByVal LocalServerID As Long, ByVal LocationID As Long, ByVal LocalServerIP As String, ByVal MacAddress As String, ByVal IPLSeqNum As Integer)

    Dim CompressedData As String
    Dim InboundData As String
    Dim FileData() As Byte
    Dim Checksum As String = ""
    Dim dst As DataTable
    Dim PrevMD5 As String = ""
    Dim DataSize As Long
    Dim Index As Long
    Dim LineOne As Boolean
    Dim EndPoint As Long
    Dim DataStr As String
    Dim SQLParams() As String
    Dim TableNum As Integer
    Dim OperationType As Integer
    Dim NumParams As Integer 'the number of parameters that should be in SQLParams based on the TableNum and OperationType
    Dim ParamsMin As Integer
    Dim ParamsMax As Integer
    Dim ParamIndex As Integer
    Dim ColName As String
    Dim RA_OverThreshold_Expected As Boolean = False

    Send_Response_Header("TransUpload", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
    InboundData = ""
    If LSVerMajor >= 5 And LSVerMinor >= 10 Then
      RA_OverThreshold_Expected = True
    End If
    Common.Write_Log(LogFile, "RA_OverThreshold_Expected=" & RA_OverThreshold_Expected)

    If Request.Files.Count > 0 Then
      ReDim FileData(Request.Files(0).ContentLength - 1)
      Request.Files(0).InputStream.Read(FileData, 0, Request.Files(0).ContentLength)
      'uncomment to view raw data
      'Send(Encoding.Default.GetString(FileData))
      CompressedData = Encoding.Default.GetString(FileData)
      FileData = Nothing
      InboundData = GZIP.DecompressString(CompressedData)
      CompressedData = Nothing
      'uncomment to view decompressed data
      'Send(Inbounddata)
    Else
      Common.Write_Log(LogFile, "No files were uploaded")
      Send("No files were uploaded")
      Exit Sub
    End If
    Common.Write_Log(LogFile, "GZip decompression successful ... size after unzipping is " & Format(Len(InboundData), "###,###,###,###,##0") & " bytes")

    If Not (MD5 = "") Then
      'if we got an MD5 checksum, make sure it matches the compressed data we received
      Checksum = Common.MD5(InboundData)
      If Checksum <> MD5 Then
        Common.Write_Log(LogFile, "Bad MD5 .. LocalServer sent ->" & MD5 & "     CentralServer computed ->" & Checksum)
        Send("Bad MD5")
        Exit Sub
      End If
      Common.Write_Log(LogFile, "Valid MD5 ->" & MD5)

      'see if this MD5 is the same as the last set of data we processed
      Common.QueryStr = "select TUMD5 from LocalServers with (NoLock) where LocalServerID=" & LocalServerID & ";"
      dst = Common.LRT_Select
      If dst.Rows.Count > 0 Then
        PrevMD5 = Common.NZ(dst.Rows(0).Item("TUMD5"), "")
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

    If MD5 = "" Then
      If InboundData = "no data" Then
        Common.Write_Log(LogFile, "Server had no data to upload - no processing performed")
        Send("ACK")
        Exit Sub
      End If
    Else
      If Right(InboundData, 7) = "no data" Then
        Common.QueryStr = "update LocalServers with (RowLock) set TUMD5='" & Checksum & "' where LocalServerID=" & LocalServerID & ";"
        Common.LRT_Execute()
        Common.Write_Log(LogFile, "Server had no data to upload - no processing performed")
        Send("ACK")
        Exit Sub
      End If
    End If

    Clear_Waiting_ACK(LocalServerID)
    Common.Write_Log(LogFile, "Received Data:" & vbCrLf & InboundData)

    DataSize = Len(InboundData)
    Index = 1
    LineOne = False
    If Not (MD5 = "") Then LineOne = True
    'Write_Log LocationID, "MD5=" & MD5 & "   Lineone=" & LineOne
    While Index < DataSize
      EndPoint = InStr(Index, InboundData, vbCrLf, vbBinaryCompare)
      DataStr = Mid(InboundData, Index, EndPoint - Index)
      Index = EndPoint + 2 'move past the CRLF
      'Write_Log LocationID, "LineOne=" & LineOne & "   ->" & DataStr
      If Not (LineOne) Then 'skip the first line if the uploaded file had an MD5 checksum
        'parse the line of data apart and execute a standard stored procedure
        SQLParams = Split(DataStr, Chr(9), -1, vbBinaryCompare)
        If UBound(SQLParams) < 1 Then
          Send_Failure_Response("Unable to retrieve at least two columns of data from line" & vbCrLf & "  Data in error = '" & DataStr & "'", LocalServerID, LocationID, LocalServerIP, MacAddress)
          Clear_Waiting_ACK(LocalServerID)
          Exit Sub
        Else
          TableNum = SQLParams(0)
          OperationType = SQLParams(1)
          NumParams = UBound(SQLParams) - 1
          'NumParams is the number of uploaded columns (not including the TableNum and Operations parameters)
          'ParamsMin is the minimum number of expected parameters (Col1 - ColN)
          'ParamsMax is the maximum number of expected parameters (Col1 - ColN)
          ParamsMin = 0
          ParamsMax = 0
          Select Case TableNum
            Case "1"
              Common.QueryStr = "dbo.pa_CPE_TU_InsertData_RD"
              If OperationType = 1 Then
                'NumParams = 8
                ParamsMin = 7
                ParamsMax = 13  'For backwards compatibility for v5.6 rollout  LogixTransNum won't be received from stores with an older version
              End If
            Case "2"
              If OperationType = 1 Then
                Common.QueryStr = "dbo.pa_CPE_TU_InsertData_RA_N"
                If RA_OverThreshold_Expected Then
                  'OverThreshold IS sent from the stores as of 5.10b01
                  ParamsMin = 17
                  ParamsMax = 18
                Else
                  'For backwards compatibility for v5.9 rollout.  OverThreshold won't be sent from the stores
                  ParamsMin = 8
                  ParamsMax = 14
                End If
              End If
              If OperationType = 2 Then
                Common.QueryStr = "dbo.pa_CPE_TU_InsertData_RA_OD"
                ParamsMin = 8
                ParamsMax = 15  'For backwards compatibility for v5.6 rollout  LogixTransNumEarned and LogixTransNumConsumed won't be received from stores with an older version
              End If
              If OperationType = 3 Then
                Common.QueryStr = "dbo.pa_CPE_TU_InsertData_RA_ND"
                ParamsMin = 8
                ParamsMax = 14  'For backwards compatibility for v5.6 rollout  LogixTransNum won't be received from stores with an older version
                'NumParams = 9
              End If
            Case "3"
              Common.QueryStr = "dbo.pa_CPE_TU_InsertData_GM"
              If OperationType = 1 Then
                ParamsMin = 5
                ParamsMax = 6  'For backwards compatibility for v5.6 rollout  LogixTransNum won't be received from stores with an older version
                'NumParams = 6
              End If
            Case "4"
              Common.QueryStr = "dbo.pa_CPE_TU_InsertData_PA"
              If OperationType = 1 Then
                ParamsMin = 5
                ParamsMax = 12  'For backwards compatibility for v5.6 rollout  CustomerTypeID and LogixTransNum won't be received from stores with an older version
                'NumParams = 5
              End If
            Case "5"
              Common.QueryStr = "dbo.pa_CPE_TU_InsertData_CR"
              If OperationType = 1 Then
                ParamsMin = 7
                ParamsMax = 9  'For backwards compatibility for v5.6 rollout  LogixTransNum won't be received from stores with an older version
                'NumParams = 8
              End If
            Case "6"
              Common.QueryStr = "dbo.pa_CPE_TU_InsertData_UL_Type2"
              If OperationType = 1 Then
                ParamsMin = 3
                ParamsMax = 6
              End If
            Case "9"
              Common.QueryStr = "dbo.pa_CPE_TU_InsertData_YB"
              If OperationType = 1 Then
                ParamsMin = 4
                ParamsMax = 5
                'NumParams = 5
              End If
            Case "10"
              If OperationType = 1 Then
                ParamsMin = 11
                ParamsMax = 17  'For backwards compatibility for v5.6 rollout  LogixTransNum won't be received from stores with an older version
                'NumParams = 12
                Common.QueryStr = "dbo.pa_CPE_TU_InsertData_SV"
              End If
              If OperationType = 3 Then
                ParamsMin = 11
                ParamsMax = 18  'For backwards compatibility for v5.6 rollout  LogixTransNum won't be received from stores with an older version
                'NumParams = 12
                Common.QueryStr = "dbo.pa_CPE_TU_InsertData_SVUpdated"
              End If
            Case "11"
              Common.QueryStr = "dbo.pa_CPE_TU_InsertData_Pref"
              If OperationType = 1 Or OperationType = 2 Then
                ParamsMin = 4
                ParamsMax = 4
              End If
            Case "12"
              Common.QueryStr = "dbo.pa_CPE_TU_InsertData_PrefMV"
              If OperationType = 1 Or OperationType = 2 Then
                ParamsMin = 4
                ParamsMax = 4
              End If
            Case "13"
              Common.QueryStr = "dbo.pa_UE_InsertData_PointsPending"
              If OperationType = 1 Then
                ParamsMin = 7
                ParamsMax = 7
              End If
            Case "14"
              Common.QueryStr = "dbo.pa_UE_InsertData_RewardDistributionPending"
              If OperationType = 1 Then
                ParamsMin = 6
                ParamsMax = 6
              End If
            Case "15"
              Common.QueryStr = "dbo.pa_UE_InsertData_RewardLimitVariablesPending"
              If OperationType = 1 Then
                ParamsMin = 6
                ParamsMax = 6
              End If
            Case "16"
              Common.QueryStr = "dbo.pt_PendingDeleteByCartID"
              If OperationType = 1 Then
                ParamsMin = 2
                ParamsMax = 2
              End If
          End Select
          If ParamsMin = 0 And ParamsMax = 0 Then 'illegal TableNum/Operation combination
            Send_Failure_Response("Operation " & OperationType & " is not supported for TableNum " & TableNum & vbCrLf & "  Data in error = '" & DataStr & "'", LocalServerID, LocationID, LocalServerIP, MacAddress)
            Clear_Waiting_ACK(LocalServerID)
            Exit Sub
          ElseIf NumParams < ParamsMin Then 'the expected number of columns was not sent
            Send_Failure_Response("Received less than the expected number of parameters for TableNum " & TableNum & " - only received " & NumParams & ", expected at least " & ParamsMin & " - row could not be processed." & vbCrLf & "  Data in error = '" & DataStr & "'", LocalServerID, LocationID, LocalServerIP, MacAddress)
            Clear_Waiting_ACK(LocalServerID)
            Exit Sub
          ElseIf NumParams > ParamsMax Then 'the expected number of columns was not sent
            Send_Failure_Response("Received more than the expected number of parameters for TableNum " & TableNum & " - received " & NumParams & ", but expected no more than " & ParamsMax & " - row could not be processed." & vbCrLf & "  Data in error = '" & DataStr & "'", LocalServerID, LocationID, LocalServerIP, MacAddress)
            Clear_Waiting_ACK(LocalServerID)
            Exit Sub
          Else
            Common.Open_LXSsp()
            If TableNum = 2 Then
              Common.LXSsp.Parameters.Add("@TableNum", SqlDbType.Int).Value = SQLParams(0)
              Common.LXSsp.Parameters.Add("@Operation", SqlDbType.Int).Value = SQLParams(1)
            Else
              Common.LXSsp.Parameters.Add("@TableNum", SqlDbType.VarChar, 4).Value = SQLParams(0)
              Common.LXSsp.Parameters.Add("@Operation", SqlDbType.VarChar, 2).Value = SQLParams(1)
            End If
            If Not (TableNum = 9) Then 'add the IPLSeqNum parameter for all stored procedure calls except pa_CPE_TU_InsertData_YB
              Common.LXSsp.Parameters.Add("@IPLSeqNum", SqlDbType.Int).Value = IPLSeqNum
            End If
            For ParamIndex = 2 To UBound(SQLParams)
              If (TableNum = 11 Or TableNum = 12) Then
                If ParamIndex <= NumParams Then    'SQLParams has last item as empty always, skipping that
                  ColName = "@Col" & Trim(Str(ParamIndex - 1))
                  Common.LXSsp.Parameters.Add(ColName, SqlDbType.VarChar, 255).Value = SQLParams(ParamIndex)
                End If
              Else
                If TableNum = 10 And ParamIndex > 12 Then
                  'skip these Col12 and Col13 for StoredValue as they aren't sent by TransUpload
                  ColName = "@Col" & Trim(Str(ParamIndex + 1))
                ElseIf TableNum = 2 And OperationType = 1 And ParamIndex > 11 And Not (RA_OverThreshold_Expected) Then
                  'Col11 (OverThreshold) won't be sent by the local servers - skip over this column
                  ColName = "@Col" & Trim(Str(ParamIndex))
                ElseIf TableNum = 4 And OperationType = 1 And ParamIndex > 7 Then
                  'Col7 (SourceTypeID) won't be sent by the local servers - skip over this column
                  ColName = "@Col" & Trim(Str(ParamIndex))
                Else
                  ColName = "@Col" & Trim(Str(ParamIndex - 1))
                End If
                ' Adding External ID support so that cross-shopping would work between EntComm and Broker
                If TableNum = 10 AndAlso OperationType = 1 AndAlso ParamIndex = 12 AndAlso String.IsNullOrEmpty(SQLParams(12))
                                    Common.LXSsp.Parameters.Add(ColName, SqlDbType.VarChar, 400).Value = MyCryptLib.SQL_StringEncrypt(GetExternalId(SQLParams(2), Convert.ToInt32(SQLParams(3)), True))
                Else If TableNum = 10 AndAlso OperationType = 3 AndAlso ParamIndex = 5 AndAlso String.IsNullOrEmpty(SQLParams(5))
                                    Common.LXSsp.Parameters.Add(ColName, SqlDbType.VarChar, 400).Value = MyCryptLib.SQL_StringEncrypt(GetExternalId(SQLParams(2), Convert.ToInt32(SQLParams(3)), False))
                Else
                  Common.LXSsp.Parameters.Add(ColName, SqlDbType.VarChar, 255).Value = SQLParams(ParamIndex)                  
                End If
              End If
            Next
			
            If TableNum = 13 And OperationType = 1 Then
              Common.QueryStr = "select PromoVarID from PointsPrograms with (NoLock) where Deleted=0 and ProgramID=" & SQLParams(3) & ";"
              dst = Common.LRT_Select
              If dst.Rows.Count > 0 Then
                 Common.LXSsp.Parameters.Add("PromoVarID", SqlDbType.BigInt).Value = Common.NZ(dst.Rows(0).Item("PromoVarID"), "")
              End If
            ElseIf TableNum = 14 and OperationType = 1 Then
              Common.QueryStr = "select O.DistPeriodVarID from Offers as O join CPE_Incentives as I on I.ClientOfferID=O.OfferID with (NoLock) where O.Deleted=0 and I.Deleted=0 and I.IncentiveID=" & SQLParams(3) & ";"
              dst = Common.LRT_Select
              If dst.Rows.Count > 0 Then
                 Common.LXSsp.Parameters.Add("PromoVarID", SqlDbType.BigInt).Value = Common.NZ(dst.Rows(0).Item("DistPeriodVarID"), "")
              End If
            End If
			
            Common.LXSsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
            Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = LocationID
            Common.LXSsp.Parameters.Add("@WaitingACK", SqlDbType.Int).Value = LocalServerID
            Common.LXSsp.ExecuteNonQuery()
            Common.Close_LXSsp()
          End If

        End If
      End If
      LineOne = False
    End While
    If Not (MD5 = "") Then  'if we got an MD5 then
      Common.QueryStr = "update LocalServers with (RowLock) set TUMD5='" & Checksum & "' where LocalServerID=" & LocalServerID & ";"
      Common.LRT_Execute()
    End If
    'Release the inserted data for processing by the TransUpdate Agents
    Common.QueryStr = "pa_CPE_TU_ReleaseData"
    Common.Open_LXSsp()
    Common.LXSsp.Parameters.Add("@WaitingACK", SqlDbType.Int).Value = LocalServerID
    Common.LXSsp.ExecuteNonQuery()
    Common.Close_LXSsp()

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
  Dim MustIPL As Boolean
  Dim TempLocationID As Long
  Dim BannerID As Integer
  Dim LocalServerIP As String = ""
  Dim MacAddress As String = ""
  Dim LSVerParts() As String
  Dim IPLSeqNum As Integer

  Common.AppName = "TransUpload.aspx"
  Response.Expires = 0
  On Error GoTo ErrorTrap
  StartTime = DateAndTime.Timer

  LocalServerIP = Trim(Request.QueryString("IP"))
  If LocalServerIP = "" Then LocalServerIP = Trim(Request.QueryString("ip"))
  If LocalServerIP = "" Then
    LocalServerIP = Trim(Request.UserHostAddress) & "(UserHostAddress)"
  End If

  MacAddress = Trim(Request.QueryString("mac"))
  If MacAddress = "" Then
    MacAddress = "Not Specified"
  End If

  LastHeard = "1/1/1980"
  LocalServerID = Common.Extract_Val(Request.QueryString("serial"))
  IPLSeqNum = Common.Extract_Val(GetCgiValue("iplseqnum"))
  LSVersion = Request.QueryString("lsversion")
  LSVerMajor = 0
  LSVerMinor = 0
  If InStr(LSVersion, ".", CompareMethod.Binary) > 0 Then
    LSVerParts = Split(LSVersion, ".", , CompareMethod.Binary)
    LSVerMajor = Common.Extract_Val(LSVerParts(0))
    LSVerMinor = Common.Extract_Val(LSVerParts(1))
  End If
  LSBuild = Common.Extract_Val(Request.QueryString("lsbuild"))
  MD5 = Trim(Request.QueryString("md5"))
  Mode = UCase(Request.QueryString("mode"))
  If Mode = "" Then Mode = "FETCH"

  LogFile = "UE-TransUpdateLog-" & Format(LocalServerID, "00000") & "." & Common.Leading_Zero_Fill(Year(Today), 4) & Common.Leading_Zero_Fill(Month(Today), 2) & Common.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"

  Common.Open_LogixRT()
  Common.Open_LogixXS()
  Common.Load_System_Info()
  Connector.Load_System_Info(Common)
  'Connector.Get_LS_Info(Common, LocalServerID, LocationID, BannerID, LastHeard, 2, IPAddress)

  Common.Write_Log(LogFile, "----------------------------------------------------------------")
  Common.Write_Log(LogFile, "** " & Common.AppName & "  -  " & Microsoft.VisualBasic.DateAndTime.Now & "  Serial: " & LocalServerID & "  Mode: " & Mode & "  Process running on server:" & Environment.MachineName & "   LSVersion=" & LSVersion & "   LSVerMajor=" & LSVerMajor & "  LSVerMinor=" & LSVerMinor & "  Requester IP Address: " & LocalServerIP & "   Requester Mac Address: " & MacAddress & "  Requester IPL Sequence Num: " & IPLSeqNum)

  ProcessOK = True
  SerialOK = False
  If Not (Mode = "FETCH") Then
    Send_Response_Header("Invalid Mode", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
    Common.Write_Log(LogFile, "Received invalid mode from IP(" & LocalServerIP & ")")
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
      Common.Write_Log(LogFile, "Returned: Invalid Serial" & vbCrLf)
      ProcessOK = False
    End If
  End If

  If ProcessOK Then
    Connector.Get_LS_Info(Common, LocalServerID, LocationID, BannerID, LastHeard, 2, LocalServerIP)
    If Request.QueryString("force") = "1" Then
      TempLocationID = Common.Extract_Val(Request.QueryString("locationid"))
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
        Common.Write_Log(LogFile, "Returned: MustIPL")
        ProcessOK = False
      End If
    End If
  End If 'ProcessOK

  If ProcessOK Then
    Common.Write_Log(LogFile, "LocationID=" & LocationID)
    If LocationID = "0" Then
      Send_Response_Header("Invalid Serial Number", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
      Common.Write_Log(LogFile, "Received invalid LocationID parsed from IP(" & LocalServerIP & ")")
    ElseIf Not (Request.QueryString("force") = "1") AndAlso Not (Connector.IsValidLocationForConnectorEngine(Common, LocationID, 9)) Then
      'the location calling TransUpload is not associated with the UE promoengine
      Common.Write_Log(LogFile, "This location is associated with a promotion engine other than UE.  Can not proceed.", True)
      Send_Response_Header("This location is associated with a promotion engine other than UE.  Can not proceed.", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
    Else
      If Request.QueryString("force") = "1" Then
        Common.Write_Log(LogFile, "Processing FORCED UPLOAD")
      End If

      Handle_Post(LocalServerID, LocationID, LocalServerIP, MacAddress, IPLSeqNum)

      TotalTime = DateAndTime.Timer - StartTime
      Common.Write_Log(LogFile, "Finished Processing")

      Common.QueryStr = "dbo.pa_CPE_TU_SetLastUpdate"
      Common.Open_LRTsp()
      Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
      Common.LRTsp.ExecuteNonQuery()
      Common.Close_LRTsp()

      TotalTime = DateAndTime.Timer - StartTime
      Common.Write_Log(LogFile, "RunTime=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
    End If
  End If 'ProcessOK

  If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
  If Not (Common.LXSadoConn.State = ConnectionState.Closed) Then Common.Close_LogixXS()
%>
<%
  Response.End()
ErrorTrap:
  Response.Write(Common.Error_Processor(, "Serial:" & LocalServerID & "  Process Info: server:" & Environment.MachineName & " Invoking LocationID=" & LocationID & "  Requester IP Address: " & LocalServerIP & "  Requester Mac Address: " & MacAddress, , Common.InstallationName))
  Common.Write_Log(LogFile, "***** An error occurred during processing! Please check the ErrorLog! *****")
  If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
  If Not (Common.LXSadoConn.State = ConnectionState.Closed) Then Common.Close_LogixXS()
  Common = Nothing
%>