<%@ Page CodeFile="..\ConnectorCB.vb" Debug="true" Inherits="ConnectorCB" Language="vb" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: Messaging.aspx
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

  Public Structure LocationInfo
    Dim LocationID As Long
    Dim LocationName As String
    Dim ExtLocationCode As String
  End Structure

  Public Common As New Copient.CommonInc
  Public Connector As New Copient.ConnectorInc
  Public MyAltID As New Copient.AlternateID
  Public CAM As New Copient.CAM
  Public TextData As String
  Public IPL As Boolean
  Public LogFile As String
  Dim StartTime As Object
  Dim TotalTime As Object
  Public LSVerMajor As Integer
  Public LSVerMinor As Integer
  Public LSBuildMajor As Integer
  Public LSBuildMinor As Integer

  '------------------------------------------------------------------------------------------------

  Function Fetch_Location_Info(ByVal LocationID As String) As LocationInfo

    Dim Location As LocationInfo
    Dim dst As DataTable

    Location.LocationID = LocationID
    Location.LocationName = "Unknown"
    Location.ExtLocationCode = "0"
    Common.QueryStr = "select LocationName, ExtLocationCode from Locations with (NoLock) where LocationID=" & LocationID & ";"
    dst = Common.LRT_Select
    If dst.Rows.Count > 0 Then
      Location.LocationName = Common.NZ(dst.Rows(0).Item("LocationName"), "")
      Location.ExtLocationCode = Common.NZ(dst.Rows(0).Item("ExtLocationCode"), "0")
    End If

    Return Location

  End Function

  '------------------------------------------------------------------------------------------------

  Function Build_STD_Message(ByVal Location As LocationInfo, ByVal LocalServerID As Long, ByVal LocalServerIP As String) As String

    Dim OutBuffer As String = ""
    Dim UpdateTime As String = ""

    UpdateTime = Microsoft.VisualBasic.DateAndTime.Today & " " & Microsoft.VisualBasic.DateAndTime.TimeOfDay

    OutBuffer = "Local Server System Message" & vbCrLf
    OutBuffer = OutBuffer & "Location: " & Location.LocationName & vbCrLf
    OutBuffer = OutBuffer & "LocationID: " & Location.LocationID & vbCrLf
    OutBuffer = OutBuffer & "Serial: " & LocalServerID & vbCrLf
    OutBuffer = OutBuffer & "IP Address: " & LocalServerIP & vbCrLf
    OutBuffer = OutBuffer & "Report received: " & UpdateTime & vbCrLf
    OutBuffer = OutBuffer & "Subject: " & GetCgiValue("subject")

    Return OutBuffer

  End Function

  '------------------------------------------------------------------------------------------------

  Sub Echo_System_Msg(ByVal LocalServerID As Long, ByVal LocationID As Long, ByVal LocalServerIP As String)

    Dim OutBuffer As String
    Dim Location As LocationInfo

    Location = Fetch_Location_Info(LocationID)
    OutBuffer = Build_STD_Message(Location, LocalServerID, LocalServerIP)

    Common.Write_Log(LogFile, OutBuffer, True)
    Common.Send_Email(Common.Get_Error_Emails(4), Common.SystemEmailAddress, "System Message - " & Common.InstallationName & "  Store #" & Location.ExtLocationCode, OutBuffer)
    Send_Response_Header("SystemMsg", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
    Send("ACK")

  End Sub

  '------------------------------------------------------------------------------------------------

  Sub Echo_Health_Msg(ByVal LocalServerID As Long, ByVal LocationID As Long, ByVal LocalServerIP As String)

    Dim OutBuffer As String
    Dim Location As LocationInfo

    Location = Fetch_Location_Info(LocationID)
    OutBuffer = Build_STD_Message(Location, LocalServerID, LocalServerIP)

    Common.Write_Log(LogFile, OutBuffer, True)
    Common.Send_Email(Common.Get_Error_Emails(7), Common.SystemEmailAddress, "Health Server Message - " & Common.InstallationName & "  Store #" & Location.ExtLocationCode, OutBuffer)
    Send_Response_Header("HealthMsg", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
    Send("ACK")

  End Sub

  '------------------------------------------------------------------------------------------------

  Sub Send_Invalid_Response(ByVal Mode As String, ByVal LocalServerID As Long, ByVal MacAddress As String, ByVal LocalServerIP As String, ByVal ValidLocalServer As Boolean)

    Dim RawRequest As String
    Dim ResponseText As String

    If Not (ValidLocalServer) Then
      ResponseText = "Invalid LocalServerID (serial)"
    Else
      ResponseText = "Invalid Request - bad mode"
    End If

    Send_Response_Header(ResponseText, Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
    Send("NAK")
    Common.Write_Log(LogFile, ResponseText & "    Mode: " & Mode & "  LSVersion=" & LSVerMajor & "." & LSVerMinor & "b" & LSBuildMajor & "." & LSBuildMinor & "  Process running on server:" & Environment.MachineName & "  Serial=" & LocalServerID & "  MacAddress=" & MacAddress & " IP=" & LocalServerIP)
    Common.Write_Log(LogFile, "Form Data: " & vbCrLf)
    RawRequest = Get_Raw_Form(Request.InputStream)
    Common.Write_Log(LogFile, RawRequest)
  End Sub

  Sub Echo_SetMustIPL_Msg(ByVal LocalServerID As Long, ByVal LocationID As Long)
    Dim localServers As New Copient.LocalServers(Common)
    Dim isTargetServerPrimary As Nullable(Of Boolean) = Nothing
    Dim targetServerID As Int32
    Dim isMustIPl As Nullable(Of Boolean) = Nothing
    Dim isRequestedServerPrimary As Nullable(Of Boolean) = Nothing
    Dim mustIPL As String = Nothing
    Try
      isRequestedServerPrimary = localServers.VerifyPrimaryServer(LocalServerID, LocationID)
      targetServerID = Common.Extract_Val(GetCgiValue("iplserverid"))
      isTargetServerPrimary = localServers.VerifyPrimaryServer(targetServerID, LocationID)

      mustIPL = GetCgiValue("value")

      If (String.IsNullOrEmpty(mustIPL) OrElse (mustIPL.ToUpper() <> "ON" AndAlso mustIPL.ToUpper() <> "OFF")) Then
        isMustIPl = Nothing
      ElseIf (mustIPL.ToUpper() = "ON") Then
        isMustIPl = 1
      ElseIf (mustIPL.ToUpper() = "OFF") Then
        isMustIPl = 0
      End If

      Common.Write_Log(LogFile, "MustIPL=" & mustIPL & " request received for ServerID:" & targetServerID.ToString() & " from ServerID:" & LocalServerID.ToString() & " having Location ID:" & LocationID.ToString(), True)

      If (isRequestedServerPrimary IsNot Nothing AndAlso isTargetServerPrimary IsNot Nothing AndAlso isRequestedServerPrimary = True AndAlso isTargetServerPrimary = False AndAlso LocalServerID <> targetServerID AndAlso isMustIPl IsNot Nothing) Then
        localServers.SetServerMustIPLFlag(isMustIPl, targetServerID)
        Send_Response_Header("MustIPLMsg", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
        Send("ACK")
        Common.Write_Log(LogFile, "MustIPL request processed for ServerID:" & targetServerID.ToString() & ",Requested ServerID:" & LocalServerID.ToString(), True)
      ElseIf (isRequestedServerPrimary Is Nothing OrElse isTargetServerPrimary Is Nothing) Then
        Send_Response_Header("Requested/Targeted server is not registered with store.", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
        Send("NAK")
        Common.Write_Log(LogFile, "Requested/Targeted server is not registered with store.", True)
      ElseIf (isRequestedServerPrimary = False) Then
        Send_Response_Header("Requested server is not active", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
        Send("NAK")
        Common.Write_Log(LogFile, "Requested server is not active", True)
      ElseIf (LocalServerID = targetServerID) Then
        Send_Response_Header("Requested server can not set mustipl for itself", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
        Send("NAK")
        Common.Write_Log(LogFile, "Requested server can not set mustipl for itself", True)
      ElseIf (isTargetServerPrimary = True) Then
        Send_Response_Header("Target server is already active; can not set mustipl for active server", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
        Send("NAK")
        Common.Write_Log(LogFile, "Target server is already active; can not set mustipl for active server", True)
      ElseIf (isMustIPl Is Nothing) Then
        Send_Response_Header("Value of MustIPL should be on/off", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
        Send("NAK")
        Common.Write_Log(LogFile, "Value of MustIPL should be on/off", True)
      End If

    Catch ex As Exception
      Send_Response_Header("Error occurred during request processing on Logix server.", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
      Send("NAK")
      Common.Write_Log(LogFile, "Error occurred during MUSTIPL request. Error Description: - " & ex.ToString(), True)
    End Try

  End Sub

  Sub Echo_FailOverFailBackMsg(ByVal mode As String, ByVal localServerID As Integer, ByVal locationID As Integer, ByVal IP As String)

    If Not ValidateVersionCompatibilityForFailoverModes() Then
      Return
    End If

    Dim localServers As New Copient.LocalServers(Common)
    Dim Location As LocationInfo
    Dim emailSub As String = String.Empty
    Dim subject As String = Request.QueryString("subject")

    If ValidateSubject(subject) Then

      Location = Fetch_Location_Info(locationID)
      emailSub = "Failover Server Message - " & Common.InstallationName & ", Store: " & Location.ExtLocationCode & " failed {0}"

      Select Case mode
        Case "FAILOVER"
          localServers.SetFailOverServer(localServerID, locationID, IP, subject)
          emailSub = String.Format(emailSub, "over")
        Case "FAILBACK"
          localServers.ReSetServerFromFailOver(locationID)
          emailSub = String.Format(emailSub, "back")
      End Select

      Common.Send_Email(Common.Get_Error_Emails(10001), Common.SystemEmailAddress, emailSub, Build_STD_Message(Location, localServerID, IP))
      Send_Response_Header(mode, Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
      Send("ACK")
      Common.Write_Log(LogFile, emailSub, True)

    End If
  End Sub

  Function ValidateVersionCompatibilityForFailoverModes() As Boolean

    If (LSVerMajor = 5 AndAlso LSVerMinor >= 19) OrElse (LSVerMajor >= 6) Then
      Return True
    Else
      Send_Response_Header("Invalid Version - Minimum version for Failover modes is 5.19", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
      Send("NAK")
      Return False
    End If

  End Function

  Function ValidateSubject(ByVal subject As String) As Boolean
    ValidateSubject = True

    Dim maxlength As Integer
    If String.IsNullOrEmpty(subject) Then
      Send_Response_Header("Invalid Subject - No Value", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
      Send("NAK")
      Return False
    ElseIf subject.Length > IIf(Integer.TryParse(Common.Fetch_SystemOption(151), maxlength), maxlength, 0) Then
      Send_Response_Header("Invalid Subject - Greater than Maximum Character Length(" & maxlength & ")", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
      Send("NAK")
      Return False
    End If
  End Function
</script>
<%
  '-----------------------------------------------------------------------------------------------
  'Main Code - Execution starts here

  Dim LocalServerID As Long
  Dim LocationID As Long
  Dim BannerID As Integer
  Dim Mode As String
  Dim LSVerParts() As String
  Dim MacAddress As String
  Dim LocalServerIP As String
  Dim LastHeard As String
  Dim ValidLocalServer As Boolean = False

  Common.AppName = "Messaging.aspx"
  Response.Expires = 0
  On Error GoTo ErrorTrap
  StartTime = DateAndTime.Timer

  LastHeard = "1/1/1980"

  MacAddress = Trim(GetCgiValue("mac"))
  If MacAddress = "" Then
    MacAddress = "Not Specified"
  End If

  LocalServerID = Common.Extract_Val(GetCgiValue("serial"))
  LocalServerIP = GetCgiValue("IP")
  If LocalServerIP = "" Then
    LocalServerIP = "IP from requestor: " & Trim(Request.UserHostAddress)
  End If

  LSVersion = Trim(GetCgiValue("lsversion"))
  LSVerMajor = 0
  LSVerMinor = 0
  If InStr(LSVersion, ".", CompareMethod.Binary) > 0 Then
    LSVerParts = Split(LSVersion, ".", , CompareMethod.Binary)
    LSVerMajor = Common.Extract_Val(LSVerParts(0))
    LSVerMinor = Common.Extract_Val(LSVerParts(1))
  End If
  LSBuild = Trim(GetCgiValue("lsbuild"))
  LSBuildMajor = 0
  LSBuildMinor = 0
  If InStr(LSBuild, ".", CompareMethod.Binary) > 0 Then
    LSVerParts = Split(LSBuild, ".", , CompareMethod.Binary)
    LSBuildMajor = Common.Extract_Val(LSVerParts(0))
    LSBuildMinor = Common.Extract_Val(LSVerParts(1))
  End If

  Mode = UCase(GetCgiValue("mode"))
  LogFile = "UE-MessagingLog-" & Format(LocalServerID, "00000") & "." & Common.Leading_Zero_Fill(Year(Today), 4) & Common.Leading_Zero_Fill(Month(Today), 2) & Common.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"

  Common.Write_Log(LogFile, "----------------------------------------------------------------")
  Common.Write_Log(LogFile, "** " & Microsoft.VisualBasic.DateAndTime.Now & "  Request|" & Request.Url.ToString())

  Common.Open_LogixRT()
  Common.Load_System_Info()
  Connector.Load_System_Info(Common)

  Connector.Get_LS_Info(Common, LocalServerID, LocationID, BannerID, LastHeard, 1, Request.UserHostAddress)
  ValidLocalServer = Connector.IsValidLocalServer(Common, LocalServerID)
  Common.Write_Log(LogFile, "** " & Microsoft.VisualBasic.DateAndTime.Now & "  Serial: " & LocalServerID & "   LocationID: " & LocationID & "  Mode: " & Mode & "  LSVersion=" & LSVerMajor & "." & LSVerMinor & "b" & LSBuildMajor & "." & LSBuildMinor & "  Process running on server:" & Environment.MachineName & " with MacAddress=" & MacAddress & " IP=" & LocalServerIP)

  If Not (ValidLocalServer) Then
    Send_Invalid_Response(Mode, LocalServerID, MacAddress, LocalServerIP, ValidLocalServer)
  ElseIf Not (Connector.IsValidLocationForConnectorEngine(Common, LocationID, 9)) Then
    'the location calling TransDownload is not associated with the UE promoengine
    Common.Write_Log(LogFile, "This location is associated with a promotion engine other than UE.  Can not proceed.", True)
    Send_Response_Header("This location is associated with a promotion engine other than UE.  Can not proceed.", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
    Send("NAK")
  Else
    Select Case Mode
      Case "SYSTEM"
        Echo_System_Msg(LocalServerID, LocationID, LocalServerIP)
        TotalTime = DateAndTime.Timer - StartTime
        Common.Write_Log(LogFile, "SystemMsg Run time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
      Case "HEALTH"
        Echo_Health_Msg(LocalServerID, LocationID, LocalServerIP)
        TotalTime = DateAndTime.Timer - StartTime
        Common.Write_Log(LogFile, "HealthMsg Run time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
      Case "SETMUSTIPL"
        Echo_SetMustIPL_Msg(LocalServerID, LocationID)
        TotalTime = DateAndTime.Timer - StartTime
        Common.Write_Log(LogFile, "SetMustIPL Run time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
      Case "FAILOVER", "FAILBACK"
        Common.Write_Log(LogFile, Mode & " request received from Serial:" & LocalServerID.ToString(), True)
        Echo_FailOverFailBackMsg(Mode, LocalServerID, LocationID, LocalServerIP)
        TotalTime = DateAndTime.Timer - StartTime
        Common.Write_Log(LogFile, Mode & " Run time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
      Case Else
        Send_Invalid_Response(Mode, LocalServerID, MacAddress, LocalServerIP, ValidLocalServer)
    End Select
  End If

  If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()

%>
<%
  Response.End()
ErrorTrap:
  Response.Write(Common.Error_Processor(, "Serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " Process Info: Server Name=" & Environment.MachineName & " Invoking LocationID=" & LocationID, , Common.InstallationName))
  Send("NAK")
  Common.Write_Log(LogFile, "***** An error occurred during processing! Please check the ErrorLog! *****")
  Common = Nothing
%>