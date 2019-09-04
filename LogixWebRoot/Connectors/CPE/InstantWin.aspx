<%@ Page CodeFile="..\ConnectorCB.vb" Debug="true" Inherits="ConnectorCB" Language="vb" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: InstantWin.aspx 
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
  Public MyAltID As New Copient.AlternateID
  Public CAM As New Copient.CAM
  Public TextData As String
  Public IPL As Boolean
  Public LogFile As String
  Dim StartTime As Object
  Dim TotalTime As Object

  ' -----------------------------------------------------------------------------------------------
  
  Function Parse_Bit(ByVal BooleanField As Boolean) As String
    If BooleanField Then
      Parse_Bit = "1"
    Else
      Parse_Bit = "0"
    End If
  End Function
  
  ' -----------------------------------------------------------------------------------------------
  
  Function Construct_Table(ByVal TableName As String, ByVal Operation As String, ByVal DelimChar As Integer, ByVal LocalServerID As Long, ByVal LocationID As Long, ByVal dst As DataTable) As String
    
    Dim TempResults As String
    Dim NumRecs As Long
    Dim row As DataRow
    Dim SQLCol As DataColumn
    Dim TempOut As String
    Dim Index As Integer
    Dim FieldList As String
    Dim LineOut As String
    
    TempOut = ""
    TempResults = ""
    NumRecs = 0
    FieldList = ""
    If dst.Rows.Count > 0 Then
      For Each SQLCol In dst.Columns
        If Not (FieldList = "") Then FieldList = FieldList & Chr(DelimChar)
        FieldList = FieldList & SQLCol.ColumnName
      Next
      For Each row In dst.Rows
        Index = 0
        LineOut = ""
        For Each SQLCol In dst.Columns
          If Not (LineOut = "") Then
            LineOut = LineOut & Chr(DelimChar)
          End If
          If SQLCol.DataType.Name = "Boolean" Then 'if it is a binary field 
            LineOut = LineOut & Parse_Bit(Common.NZ(row(Index), 0))
          ElseIf SQLCol.DataType.Name = "Int32" Or SQLCol.DataType.Name = "Int64" Then 'if it is an Int or BigInt Field
            LineOut = LineOut & Common.NZ(row(Index), 0)
          ElseIf SQLCol.DataType.Name = "DateTime" Then 'time, so format with zone
            LineOut = LineOut & Format(Common.NZ(row(Index), Now()), "yyyy-MM-dd hh:mm:sszzz")
          Else 'else treat it as a string
            LineOut = LineOut & Common.NZ(row(Index), "")
          End If
          Index = Index + 1
        Next
        TempResults = TempResults & LineOut & vbCrLf
        NumRecs = NumRecs + 1
      Next
      TempOut = TempOut & "1:" & TableName & vbCrLf
      TempOut = TempOut & "2:" & Operation & vbCrLf
      TempOut = TempOut & "3:" & FieldList & vbCrLf
      TempOut = TempOut & "4:" & NumRecs & vbCrLf
      'Common.Write_Log(LogFile, TempOut)
      TempOut = TempOut & TempResults
    End If
    
    Construct_Table = TempOut
    
  End Function
   
  ' -----------------------------------------------------------------------------------------------
  
</script>

<%
  '-----------------------------------------------------------------------------------------------
  'Main Code - Execution starts here
  
  Dim LocalServerID As Long
  Dim TempLocationID As Long
  Dim LocationID As Long
  Dim TriggerID As Long
  Dim BannerID As Integer
  Dim LastHeard As String
  Dim ZipOutput As Boolean
  Dim DataFile As String
  Dim ZipFile As String
  Dim FileStamp As String
  Dim Mode As String
  Dim RawRequest As String
  Dim Index As Long
  Dim ReturnValue As Long
  Dim OutStr As String
  Dim OperateAtEnterprise As Integer
  Dim MacAddress As String
  Dim LocalServerIP As String
  
  Common.AppName = "InstantWin.aspx"
  Response.Expires = 0
  On Error GoTo ErrorTrap
    
  LocalServerID = Common.Extract_Val(Request.QueryString("serial"))
  LogFile = "CPE-InstantWinLog-" & Format(LocalServerID, "00000") & "." & Common.Leading_Zero_Fill(Year(Today), 4) & Common.Leading_Zero_Fill(Month(Today), 2) & Common.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
    
  LocalServerIP = Trim(Request.QueryString("IP"))
  MacAddress = Trim(Request.QueryString("mac"))
  If MacAddress = "" Or MacAddress = "0" Then
    MacAddress = "0"
  End If
  If LocalServerIP = "" Or LocalServerIP = "0" Then
    Common.Write_Log(LogFile, "Could not get IP from query. Analyzing client request for IP ...")
    LocalServerIP = Trim(Request.UserHostAddress)
  End If
  StartTime = DateAndTime.Timer
  
  LastHeard = "1/1/1980"
  LSVersion = Common.Extract_Val(Request.QueryString("lsversion"))
  LSBuild = Common.Extract_Val(Request.QueryString("lsbuild"))
  TriggerID = Common.Extract_Val(Request.QueryString("triggerid"))
  TempLocationID = Common.Extract_Val(Request.QueryString("locationid"))  
  
  Common.Open_LogixRT()
  Connector.Load_System_Info(Common)
  
  OperateAtEnterprise = Common.Fetch_CPE_SystemOption(91)
  
  If OperateAtEnterprise <> 0 Then
    LocationID = TempLocationID
  Else
     LocationID = 0
  End If
  
  If TriggerID = 0 Then
    Common.Write_Log(LogFile, "----------------------------------------------------------------")
    Common.Write_Log(LogFile, "** " & Microsoft.VisualBasic.DateAndTime.Now & "LSVersion=" & LSVersion & "  LSBuild=" & LSBuild & "  Process running on server:" & Environment.MachineName)
    Common.Write_Log(LogFile, "Invalid TriggerID " & TriggerID & "from Serial:" & LocalServerID & " MacAddress:" & MacAddress & " IP:" & LocalServerIP & "on server:" & Environment.MachineName & vbCrLf)
    ReturnValue = -1
  Else
    Connector.Get_LS_Info(Common, LocalServerID, LocationID, BannerID, LastHeard, 1, LocalServerIP, true)
    Common.Write_Log(LogFile, "----------------------------------------------------------------")
    Common.Write_Log(LogFile, "** " & Microsoft.VisualBasic.DateAndTime.Now & "  Serial: " & LocalServerID & "   LocationID: " & LocationID & " MacAddress:" & MacAddress & " IP:" & LocalServerIP & "  LSVersion=" & LSVersion & "  LSBuild=" & LSBuild & "  Process running on server:" & Environment.MachineName)

    If LocationID = 0 Then
      Common.Write_Log(LogFile, "Invalid Serial - associated LocationID not found")
      Common.Write_Log(LogFile, "Serial:" & LocalServerID & " with MacAddress: " & MacAddress & " IP:" & LocalServerIP & " Received invalid request from !" & " server:" & Environment.MachineName)
      Common.Write_Log(LogFile, "Query String: " & Request.RawUrl & vbCrLf & "Form Data: " & vbCrLf)
      RawRequest = Get_Raw_Form(Request.InputStream)
      Common.Write_Log(LogFile, RawRequest)
      Send_Response_Header("Invalid Serial Number", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
      ReturnValue = -1
    ElseIf Not (Connector.IsValidLocationForConnectorEngine(Common, LocationID, 2)) Then
      'the location calling InstantWin is not associated with the CPE promoengine
      Common.Write_Log(LogFile, "This location is associated with a promotion engine other than CPE.  Can not proceed.", True)
      Common.Write_Log(LogFile, "Query String: " & Request.RawUrl & vbCrLf & "Form Data: " & vbCrLf)
      RawRequest = Get_Raw_Form(Request.InputStream)
      Common.Write_Log(LogFile, RawRequest)
      Send_Response_Header("This location is associated with a promotion engine other thanCPE.  Can not proceed.", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
      ReturnValue = -1
    Else
      Common.QueryStr = "dbo.pt_CPE_IncentiveEIWTriggersUsed_Insert"
      Common.Open_LRTsp()
      Common.LRTsp.Parameters.Add("@TriggerID", SqlDbType.BigInt).Value = TriggerID
      Common.LRTsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = LocationID
      Common.LRTsp.Parameters.Add("@Return", SqlDbType.BigInt).Direction = ParameterDirection.Output
      Common.LRTsp.ExecuteNonQuery()
      ReturnValue = Common.LRTsp.Parameters("@Return").Value
      Common.Close_LRTsp()
    End If 'locationid="0"
  End If
  
  If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
  
  Send_Response_Header("InstantWin", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
  TotalTime = DateAndTime.Timer - StartTime
  
  If ReturnValue = "-1" Then
    Common.Write_Log(LogFile, "** Serial:" & LocalServerID & " MacAddress:" & (Trim(Request.UserHostAddress)) & " LocationID:" & LocationID & " NAK for TriggerID: " & TriggerID)
    OutStr = "NAK" & vbCrLf
  Else
    Common.Write_Log(LogFile, "** LocationID:" & LocationID & " ACK for TriggerID: " & TriggerID)
    OutStr = "ACK" & vbCrLf
  End If
  
  OutStr = OutStr & "L:Central Execute Time: " & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)" & vbCrLf
  Send(OutStr)
%>
<%
  Response.End()
ErrorTrap:
  Response.Write(Common.Error_Processor(, "Serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP:" & LocalServerIP & " Process Info: Server Name=" & Environment.MachineName & " Invoking LocationID=" & LocationID, , Common.InstallationName))
  Common.Write_Log(LogFile, "***** An error occurred during processing! Please check the ErrorLog! *****")
  Common = Nothing
%>
