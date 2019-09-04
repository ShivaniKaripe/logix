<%@ Page CodeFile="..\ConnectorCB.vb" Debug="true" Inherits="ConnectorCB" Language="vb" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: ServerList.aspx
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
  Public LocalServerIP As String
  Public MacAddress As String
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

  Function Parse_Bit(ByVal BooleanField As Boolean) As String
    If BooleanField Then
      Parse_Bit = "1"
    Else
      Parse_Bit = "0"
    End If
  End Function

  ' -----------------------------------------------------------------------------------------------

  Sub Construct_Output(ByVal LocalServerID As Long)

    Dim OutStr As String
    Dim dst As DataTable
    Dim row As DataRow
    Dim dst2 As DataTable
    Dim ConditionStr As String
    Dim FoundData As Boolean = False
    Dim FetchRunID As Boolean = False
    Dim LastRunID As Long

    ConditionStr = ""
    If Not (LocalServerID = 0) Then
      ConditionStr = " and LS.LocalServerID=" & LocalServerID & " "
      Common.Write_Log(LogFile, "Generating data for specified LocalServerID=" & LocalServerID)
      FetchRunID = True
    Else
      Common.Write_Log(LogFile, "No server specified - generating data for all local servers")
    End If

    Common.Write_Log(LogFile, "Returned the following data:")
    SD("ServerList")
    Common.Write_Log(LogFile, "ServerList")

    OutStr = ""
    Common.QueryStr = "select LS.LocalServerID, isnull(LS.LocationID, 0) as LocationID, isnull(LS.FailoverServer, 0) as FailoverServer, " & _
                      "  isnull(LS.LastHeard, '1/1/1980') as LastHeard, isnull(LS.LastIP, '') as LastIP, isnull(L.ExtLocationCode, '') as ExtLocationCode " & _
                      "from LocalServers  as LS with (NoLock) Left Join Locations as L with (NoLock) on LS.LocationID=L.LocationID " & _
                      "where LastIP is not null and (L.EngineID=2 or L.EngineID is NULL) " & _
                      "and ((LS.LocationID>0 and L.LocationID is not null) or IsNull(LS.LocationID,0)=0) " & ConditionStr & _
                      "order by LocalServerID;"
    dst = Common.LRT_Select
    If dst.Rows.Count > 0 Then
      FoundData = True
    End If
    For Each row In dst.Rows
      LastRunID = 0
      If FetchRunID = True Then
        Common.QueryStr = "select max(RunID) as LastRunID from LS_HealthHistory with (NoLock) where LocalServerID=" & LocalServerID & ";"
        dst2 = Common.LWH_Select
        If dst2.Rows.Count > 0 Then
          LastRunID = Common.NZ(dst2.Rows(0).Item("LastRunID"), 0)
        End If
      End If
      OutStr = row.Item("LocalServerID") & "," & row.Item("LocationID") & "," & Parse_Bit(row.Item("FailoverServer")) & "," & _
               row.Item("LastHeard") & "," & LastRunID & "," & Trim(row.Item("LastIP")) & "," & Trim(row.Item("ExtLocationCode")).Replace(",", " ")
      SD(OutStr)
      'Common.Write_Log(LogFile, OutStr)
    Next
    If Not (FoundData) Then
      SD("no data")
      Common.Write_Log(LogFile, "no data")
    End If

  End Sub
</script>
<%

  '-----------------------------------------------------------------------------------------------
  'Main Code - Execution starts here

  Dim LocalServerID As Long
  Dim LocationID As Long
  Dim LastHeard As String
  Dim TotalTime As Object
  Dim Mode As String = ""
  Dim CompressedArray() As Byte
  Dim BannerID As Integer

  Common.AppName = "ServerList.aspx"
  Response.Expires = 0
  On Error GoTo ErrorTrap
  StartTime = DateAndTime.Timer
  LocalServerID = Common.Extract_Val(Request.QueryString("serial"))
  LogFile = "UE-LS-HealthLog." & Common.Leading_Zero_Fill(Year(Today), 4) & Common.Leading_Zero_Fill(Month(Today), 2) & Common.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
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
  LocalServerID = Common.Extract_Val(Request.QueryString("serial"))
  'Mode = UCase(Request.QueryString("mode"))

  Common.Open_LogixRT()
  Common.Open_LogixWH()
  Common.Load_System_Info()
  Connector.Load_System_Info(Common)
  Connector.Get_LS_Info(Common, LocalServerID, LocationID, BannerID, LastHeard, 2, LocalServerIP)

  Common.Write_Log(LogFile, "----------------------------------------------------------------")
  Common.Write_Log(LogFile, "** " & Microsoft.VisualBasic.DateAndTime.Now & "  Serial: " & LocalServerID & " IPAddress: " & (Trim(Request.UserHostAddress)) & "   LocationID: " & LocationID & "  Mode: " & Mode & " server: " & Environment.MachineName)

  Response.ContentType = "application/x-gzip"

  OutboundBuffer = New StringBuilder
  Construct_Output(LocalServerID)
  Common.Write_Log(LogFile, "Starting GZip compression ... size before zipping is " & Format(OutboundBuffer.Length, "###,###,###,###,##0") & " bytes")
  TotalTime = DateAndTime.Timer - StartTime
  Common.Write_Log(LogFile, "Serial: " & LocalServerID & " IPAddress:" & Trim(Request.UserHostAddress) & " Time elapsed before starting compression=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
  CompressedArray = Encoding.Default.GetBytes(GZIP.CompressString(OutboundBuffer.ToString))
  Response.BinaryWrite(CompressedArray)
  Common.Write_Log(LogFile, "GZip compression successful ... size after zipping is " & Format(UBound(CompressedArray) + 1, "###,###,###,###,##0") & " bytes")
  TotalTime = DateAndTime.Timer - StartTime
  Common.Write_Log(LogFile, "Time elapsed after finishing compression=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")

  If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
  If Not (Common.LWHadoConn.State = ConnectionState.Closed) Then Common.Close_LogixWH()

%>
<%Response.End()
ErrorTrap:
  Response.Write(Common.Error_Processor(, "Serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " Process Info: Server Name=" & Environment.MachineName & " Invoking LocationID=" & LocationID, , Common.InstallationName))
  Common.Write_Log(LogFile, "***** An error occurred during processing! Please check the ErrorLog! *****")
  If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
  If Not (Common.LWHadoConn.State = ConnectionState.Closed) Then Common.Close_LogixWH()
  Common = Nothing
%>