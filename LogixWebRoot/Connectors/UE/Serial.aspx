<%@ Page CodeFile="..\ConnectorCB.vb" Debug="true" Inherits="ConnectorCB" Language="vb" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
    ' *****************************************************************************
    ' * FILENAME: Serial.aspx
    ' *~~~~~~~~~~~~pu~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
    Public LogFile As String
    Dim StartTime As Object
    Dim TotalTime As Object
    Public MacAddress As String
    Public LSVersionNum As Copient.ConnectorInc.LSVersionRec
    Dim DelimChar As String = Chr(30)
    Dim LocationLib As Copient.LocalServers
    Dim dataTransferLib As CPE_LS_DataTransferLib.LS_DataTransferLib

    '------------------------------------------------------------------------------------------------

    ''' <summary>
    ''' This function fetches the clc parameter from the web request and looks for a location with a matching ExtLocationCode that is associated with the UE promo engine.
    ''' If a LocationID is not found, the function logs an error message and returns a response header indicating the error to the requestor.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function Resolve_Location() As Integer

        Dim DT As DataTable
        Dim LocationID As Long = 0
        Dim EngineID As Integer = -1
        Dim ClientLocationCode As String

        ClientLocationCode = Trim(GetCgiValue("clc"))
        If ClientLocationCode = "" Then
            Send_Response_Header("ClientLocationCode (clc) was not provided", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
            Common.Write_Log(LogFile, "ClientLocationCode (clc) was not provided!", True)
            Response.End()
        End If

        'ok, we have a client location code, let see if we can find a matching LogixRT.Locations record
        Common.QueryStr = "select LocationID, isnull(EngineID, -1) as EngineID from Locations with (NoLock) where ExtLocationCode='" & Common.Parse_Quotes(ClientLocationCode) & "' and Deleted=0;"
        DT = Common.LRT_Select
        If DT.Rows.Count > 0 Then
            EngineID = Common.NZ(DT.Rows(0).Item("EngineID"), 0)
            If Not (EngineID = 9) Then
                Send_Response_Header("This location code is not associated with the UE engine", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
                Common.Write_Log(LogFile, "Invalid ClientLocationCode - clc=" & ClientLocationCode & " - not associated with the UE engine.", True)
                Response.End()
            Else
                LocationID = DT.Rows(0).Item("LocationID")
            End If
        End If
        If LocationID = 0 Then
            Send_Response_Header("Invalid Location Code", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
            Common.Write_Log(LogFile, "Invalid ClientLocationCode - clc: '" & ClientLocationCode & "'", True)
            Response.End()
        Else
            Common.Write_Log(LogFile, "Found LocationID " & LocationID & " for CLC:'" & ClientLocationCode & "'", True)
        End If

        Return LocationID

    End Function

    '------------------------------------------------------------------------------------------------

    ''' <summary>
    ''' This function fetches the MAC parameter from the web request and looks for a LocalServer with a matching MacAddress.
    ''' If a record with a matching MacAddress is not found, then a new LocalServers record is created.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function Resolve_Server() As Integer

        Dim DT As DataTable
        Dim LocalServerID As Integer = 0
        Dim IPAddress As String = ""

        MacAddress = Trim(GetCgiValue("mac"))
        IPAddress = GetCgiValue("IP")
        Dim regmac As String = "^([0-9A-Fa-f]{2}[:-]){5}([0-9A-Fa-f]{2})$"
        Dim regex As Regex = New Regex(regmac)
        'first, make sure a MAC addressed was passed to us. If it wasn't, we can't proceed
        If MacAddress = "" Then
            Send_Response_Header("Missing MAC Address", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
            Common.Write_Log(LogFile, "Missing MAC Address!", True)
            Response.End()
        ElseIf Not regex.IsMatch(MacAddress) Then
            Send_Response_Header("Invalid Mac Address! Please provide the valid Mac address", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
            Common.Write_Log(LogFile, "Invalid Mac Address! Please provide the valid Mac address", True)
            Response.End()
        End If

        'see if we have a LocalServers record for this MacAddress
        Common.QueryStr = "select LocalServerID from LocalServers with (NoLock) where MacAddress='" & Common.Parse_Quotes(MacAddress) & "';"
        DT = Common.LRT_Select
        If DT.Rows.Count > 0 Then
            LocalServerID = DT.Rows(0).Item("LocalServerID")
        End If
        If LocalServerID = 0 Then 'create a LocalServers record for this MacAddress
            LocalServerID = Connector.Create_UE_LocalServer(Common, MacAddress, IPAddress)
            Common.Write_Log(LogFile, "Created LocalServers record (" & LocalServerID & ") for MacAddress '" & MacAddress & "'", True)
        Else
            'update LocalServers.LastHeard
            Common.Write_Log(LogFile, "Found Serial=" & LocalServerID & " for MacAddress:" & MacAddress, True)
            Common.QueryStr = "Update LocalServers with (RowLock) set LastHeard=getdate(), LastIP='" & Left(Common.Parse_Quotes(IPAddress), 15) & "' where LocalServerID=" & LocalServerID.ToString & ";"
            Common.LRT_Execute()
        End If

        Return LocalServerID

    End Function

    '-----------------------------------------------------------------------------------------------

    Sub Fetch_Serial()

        Dim LocationID As Long
        Dim LocalServerID As Integer
        Dim OldLocalServerID As Integer
        Dim OldLocalServerMustIPL As Boolean
        Dim OldLocationID As Long
        Dim ServerSwitch As Boolean = False
        Dim MustIPL As Boolean = False
        Dim dt As New DataTable
        Dim FirstServer As Boolean  'this will be used in the future when UE needs to know the difference between a MustIPL and a NewIPL

        'add code to set MUSTIPL mode if oldLocationID for LocalServerID <> NewLocationID

        'these functions will never return with out a valid LocationID and LocalServerID
        LocationID = Resolve_Location()
        LocalServerID = Resolve_Server()  'returns the LocalServerID associated with the MAC address request parameter, and creates a new local server record if necessary

        'grab the LocationID that is associated with the MAC address passed to us in the call
        OldLocationID = Connector.Fetch_UE_CurrentLSLocation(Common, LocalServerID)
        'grab the LocalServerID that is associated with the CLC (client location code) that was passed to us in the call
        OldLocalServerID = Connector.Fetch_UE_ActiveServer(Common, LocationID)
        'grab the MustIPL state of the OldLocalServerID
        OldLocalServerMustIPL = Connector.Fetch_UEServer_MustIPLStatus(Common, OldLocalServerID)
        If OldLocalServerID <> LocalServerID Then
            Common.Write_Log(LogFile, "Changing the server handling LocationID " & LocationID & " to from LocalServerID=" & OldLocalServerID & "  to LocalServerID=" & LocalServerID & ";", True)
            ServerSwitch = True
            'make the local server that invoked the request to serial.aspx the new active server for the store
            FirstServer = Connector.Associate_UELS_Location(Common, LocalServerID, LocationID, Copient.ConnectorInc.AssociationType.ActiveServer)
            'update failover status of the secondary server to 1.It can cause issue when the MAC address is changed.
            If Not FirstServer Then
                Common.QueryStr = "select FailoverServer from localservers where localserverid=" & OldLocalServerID
                dt=Common.LRT_Select()
                If dt.Rows(0).Item("FailoverServer") = 0 Then
                    Common.QueryStr = "Update localservers set FailoverServer=1 where localserverid=" & LocalServerID
                    Common.LRT_Execute()
                    End If
                End If
           
                'set the OldLocalServer to MustIPL=1  This is a new business rule as of 5/21/2013
                Connector.Set_LS_MustIPLStatus(Common, OldLocalServerID, True)
                'if the LocalServer is no longer servicing the same location it had previously, then we must set it to MustIPL=1
                If OldLocationID <> LocationID Then
                    Common.Write_Log(LogFile, "LocalServerID=" & LocalServerID.ToString & " is changing from servicing LocationID=" & OldLocationID.ToString & " to LocationID=" & LocationID.ToString, True)
                    Common.QueryStr = "Update LocalServers set MustIPL=1 where LocalServerID=" & LocalServerID & ";"
                    Common.LRT_Execute()
                End If
            Else
                Common.Write_Log(LogFile, "Server re-requested serial number - Location and MacAddress have not changed", True)
            End If

            'If this is the Active server for a location, also store it in the Locations table (for backward compatibility with other connectors/UEOfferAgent)
            Connector.Set_LegacyActiveServer_Location(Common, LocalServerID, LocationID)

            'if the new active server is in a MustIPL state, then we should return MustIPL to the caller
            If Connector.Fetch_UEServer_MustIPLStatus(Common, LocalServerID) Then
                MustIPL = True
            End If
            'if the active server for the store is switching and the previously active server was in a MustIPL state, then we should return MustIPL to the caller
            If Not (MustIPL) AndAlso (ServerSwitch) AndAlso (OldLocalServerMustIPL) Then
                MustIPL = True
                Connector.Set_LS_MustIPLStatus(Common, LocalServerID, True)
            End If

            Common.Write_Log(LogFile, "Fetched serial number '" & LocalServerID & "' for Local Server at ClientLocationCode '" & Trim(GetCgiValue("clc")) & "'  MustIPL=" & MustIPL, True)
            If MustIPL Then
                Send_Response_Header("MustIPL", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
            Else
                Send_Response_Header("Serial", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
            End If
            ' @todo: Send LastHeard for Active Server
            Send(LocalServerID)

    End Sub

    '-----------------------------------------------------------------------------------------------

    ' Do most of the serial work but this server doesn't take over communications for the store.
    Sub Fetch_Maintenance()

        Dim LocationID As Long
        Dim LocalServerID As String
        Dim FirstServer As Boolean  'this will be used in the future when UE needs to know the difference between a MustIPL and a NewIPL
        Dim DT As DataTable
        Dim LastHeard As DateTime  'The date/time the active server for the location communicated with Logix
        Dim UTCLastHeard As DateTime
        Dim TimeStamp As DateTime
        Dim UTCTimeStamp As DateTime
        Dim ActiveMustIPL As Integer
        Dim ActiveLocalServerID As Integer
        Dim LocalServerLocation As Copient.ConnectorInc.UEServerLocationRec  'Contains information about LocalServer's association with a location
        Dim BaseTime As DateTime = "1/1/1970 00:00:00"
        Dim dtLocalServers As New DataTable()
        Dim LocalServerIP As String
        Dim dtLocalIDSeed As New DataTable()
        'Business rule definition:
        'if the location for the LocalServer in central doesn't match the location for the CLC
        ' - move the server to the location associated with the CLC
        'associate the local server with the location (as a secondary) if the server is not associated with any location

        'these functions will never return with out a valid LocationID and LocalServerID
        LocationID = Resolve_Location()   'returns the LocationID associated with the CLC (client location code) request parameter.  Appends if the CLS is invalid.
        LocalServerID = Resolve_Server()  'returns the LocalServerID associated with the MAC address request parameter, and creates a new local server record if necessary

        'fetch the location information for the local server and make sure it's location matches the CLC
        LocalServerLocation = Connector.Fetch_UELS_LocationInfo(Common, LocalServerID)
        If LocalServerLocation.LocationID <> LocationID Then
            'The LocalServer doesn't match the CLC passed to us, or is not associated with any location
            'Associate the local server to the CLC that was passed to us and set it up as a secondary server
            FirstServer = Connector.Associate_UELS_Location(Common, LocalServerID, LocationID, Copient.ConnectorInc.AssociationType.SecondaryServer)
            Connector.Set_LS_MustIPLStatus(Common, LocalServerID, True)
            Common.Write_Log(LogFile, "Associated LocalServer " & LocalServerID & " with LocationID " & LocationID & " as a secondary server", True)
        End If

        'fetch the ActiveLocalServer for this location
        ActiveLocalServerID = Connector.Fetch_UE_ActiveServer(Common, LocationID)

        ' Look up the MustIPL state, LastHeard, and current timestamp of the active server.
        Common.QueryStr = "SELECT MustIPL, LastHeard, getdate() as timestamp, getUTCDate() as UTCtimestamp FROM LocalServers WHERE LocalServerID=" & ActiveLocalServerID & ";"
        DT = Common.LRT_Select()

        ' There may not be an active server registered for this store.
        If DT.Rows.Count = 0 Then
            Common.QueryStr = "SELECT getdate() As timestamp, getUTCDate() as UTCtimestamp;"
            DT = Common.LRT_Select()
            ActiveMustIPL = -1
            LastHeard = New DateTime(1970, 1, 1, 0, 0, 0)
            TimeStamp = DT.Rows(0).Item("timestamp")
            UTCTimeStamp = DT.Rows(0).Item("UTCtimestamp")
            UTCLastHeard = LastHeard
        Else
            ActiveMustIPL = DT.Rows(0).Item("MustIPL")
            If ActiveMustIPL <> 0 Then
                ActiveMustIPL = 1
            End If
            LastHeard = DT.Rows(0).Item("LastHeard")
            TimeStamp = DT.Rows(0).Item("timestamp")
            UTCTimeStamp = DT.Rows(0).Item("UTCtimestamp")
            UTCLastHeard = DateAdd(DateInterval.Second, (UTCTimeStamp - TimeStamp).TotalSeconds, LastHeard)
        End If

        TimeStamp = DateAdd(DateInterval.Second, (UTCTimeStamp - TimeStamp).TotalSeconds, TimeStamp)

        Send_Response_Header("Serial", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
        Send(LocalServerID & DelimChar & IIf(Connector.Fetch_UEServer_MustIPLStatus(Common, LocalServerID), "1", "0") & DelimChar & ActiveLocalServerID & DelimChar & (UTCLastHeard - BaseTime).TotalSeconds & DelimChar & ActiveMustIPL & DelimChar & (TimeStamp - BaseTime).TotalSeconds)

        'Send(LocalServerID & " " & callermustipl & " " & ActiveLocalServerID & " " & (LastHeard - DateTime.MinValue).TotalSeconds & " " & ActiveMustIPL & " " & (TimeStamp - DateTime.MinValue).TotalSeconds)

        LocalServerIP = Request.QueryString("IP")
        If LocalServerIP = "" Then
            LocalServerIP = Trim(Request.UserHostAddress)
        End If

        LocationLib = New Copient.LocalServers(Common)
        dataTransferLib = New CPE_LS_DataTransferLib.LS_DataTransferLib(Common, LocalServerID, LocationID, LocalServerIP)

        ' Construct data from LocalServers table
        dtLocalServers = LocationLib.GetLocalServerDetailForMntService(LocalServerID)
        dataTransferLib.Construct_Table("LocalServers", "1", LocalServerID, LocationID, dtLocalServers)

        ' Construct data from pa_CPE_LocalID_Seeds
        dtLocalIDSeed = LocationLib.GetLocalIDSeeds(LocalServerID)
        dataTransferLib.Construct_Table("LocalIDSeeds", "5", LocalServerID, LocationID, dtLocalIDSeed)

        dataTransferLib.Construct_EOF()
        Send(dataTransferLib.Output)
    End Sub
</script>
<%
    '-----------------------------------------------------------------------------------------------
    'Main Code - Execution starts here

    Dim LocalServerID As Long
    Dim LocationID As Long
    Dim LastHeard As String
    Dim Mode As String
    Dim RawRequest As String
    Dim LocalServerIP As String

    Common.AppName = "Serial.aspx"
    Response.Expires = 0
    On Error GoTo ErrorTrap
    StartTime = DateAndTime.Timer

    LastHeard = "1/1/1980"

    MacAddress = Trim(Request.QueryString("mac"))
    If MacAddress = "" Then
        MacAddress = "Not Specified"
    End If

    LocalServerIP = Request.QueryString("IP")
    If LocalServerIP = "" Then
        LocalServerIP = "IP from requestor: " & Trim(Request.UserHostAddress)
    End If

    LSVersionNum = Connector.Convert_LSVersion(Common, Trim(Request.QueryString("lsversion")), Trim(Request.QueryString("lsbuild")))

    Mode = UCase(Request.QueryString("mode"))
    'this is serial.aspx, the whole purpose is to return a LocalServerID, so at this point, since we don't know what the local serverID will be, logging has to go into a file for LocalServerID 0
    LogFile = "UE-SerialLog-" & Format(0, "00000") & "." & Common.Leading_Zero_Fill(Year(Today), 4) & Common.Leading_Zero_Fill(Month(Today), 2) & Common.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"

    Common.Open_LogixRT()
    Common.Load_System_Info()
    Connector.Load_System_Info(Common)

    Common.Write_Log(LogFile, "----------------------------------------------------------------", True)
    Common.Write_Log(LogFile, "** " & Microsoft.VisualBasic.DateAndTime.Now & "  Mode: " & Mode & "  LSVersion=" & LSVersionNum.LSMajorVersion & "." & LSVersionNum.LSMinorVersion & "b" & LSVersionNum.BuildMajorVersion & "." & LSVersionNum.BuildMinorVersion & "  Process running on server:" & Environment.MachineName & "  Serial=" & LocalServerID & "  MacAddress=" & MacAddress & " IP=" & LocalServerIP, True)

    'make sure we have a MAC address
    If Trim(Request.QueryString("mac")) = "" Then
        Send_Response_Header("Missing MAC Address", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
        Common.Write_Log(LogFile, "Missing MAC Address! - can not continue", True)
        Exit Sub
    End If

    Select Case Mode
        Case "SRL"
            Fetch_Serial()
        Case "MNT"
            Fetch_Maintenance()
        Case Else
            Send_Response_Header("Invalid Request - bad mode", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
            Common.Write_Log(LogFile, "Received invalid request!", True)
            Common.Write_Log(LogFile, "Query String: " & Request.RawUrl & vbCrLf & "Form Data: " & vbCrLf, True)
            RawRequest = Get_Raw_Form(Request.InputStream)
            Common.Write_Log(LogFile, RawRequest, True)
    End Select
    TotalTime = DateAndTime.Timer - StartTime
    Common.Write_Log(LogFile, "Serial Run time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)", True)

    If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()

%>
<%
    Response.End()
ErrorTrap:
    Response.Write(Common.Error_Processor(, "Serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " Process Info: Server Name=" & Environment.MachineName & " Invoking LocationID=" & LocationID, , Common.InstallationName))
    Common.Write_Log(LogFile, "***** An error occurred during processing! Please check the ErrorLog! *****", True)
    Common = Nothing
%>