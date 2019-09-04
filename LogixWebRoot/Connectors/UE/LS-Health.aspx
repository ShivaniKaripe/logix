<%@ Page CodeFile="..\ConnectorCB.vb" Debug="true" Inherits="ConnectorCB" Language="vb" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Xml" %>
<%
  ' *****************************************************************************
  ' * FILENAME: LS-Health.aspx
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
  Public TextData As String
  Public IPL As Boolean
  Public LogFile As String
  Public FileStamp As String
  Public FileNum As Integer
  Public StartTime As Decimal
  Public MacAddress As String
  Public LocalServerIP As String
  Public LocalServerID As String
  Private Const DUP_PRIMARY_KEY As Integer = 2627

  Function Handle_Post(ByRef Data As String) As Boolean

    Dim CompressedData As String
    Dim InboundData As String
    Dim FileData() As Byte
    Dim DataRetrieved As Boolean = False

    Try
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
        'uncomment to view decompressed data
        'Send(InboundData)

        Data = InboundData
        DataRetrieved = True

      Else
        Common.Write_Log(LogFile, "No files were uploaded")
        Data = "No files were uploaded"
        DataRetrieved = False
      End If
    Catch ex As Exception
      Send(ex.ToString)
      Common.Error_Processor()
      Common.Write_Log(LogFile, ex.ToString())
    End Try

    Return DataRetrieved
  End Function

  Function ProcessData(ByVal data As String) As String

    Dim responseText As String = ""
    Dim stopWatch As New System.Diagnostics.Stopwatch()
    Dim XmlDoc As New XmlDocument()
    Dim ServerNodes As XmlNodeList = Nothing
    Dim LSNode As XmlNode = Nothing
    Dim dtLSHistory As DataTable = Nothing
    Dim dtLSErrors As DataTable = Nothing

    Try
      stopWatch.Start()
      XmlDoc.LoadXml(data)

      'create the local server history datatable
      dtLSHistory = New DataTable
      dtLSHistory.Columns.Add("LocalServerID", System.Type.GetType("System.Int32"))
      dtLSHistory.Columns.Add("RunID", System.Type.GetType("System.Int32"))
      dtLSHistory.Columns.Add("RunDate", System.Type.GetType("System.DateTime"))
      dtLSHistory.Columns.Add("HealthStatusID", System.Type.GetType("System.Int32"))
      dtLSHistory.Columns.Add("Sev1", System.Type.GetType("System.Int32"))
      dtLSHistory.Columns.Add("Sev10", System.Type.GetType("System.Int32"))
      dtLSHistory.Columns.Add("HealthSeverityID", System.Type.GetType("System.Int32"))

      ' create the local server error datatable
      dtLSErrors = New DataTable
      dtLSErrors.Columns.Add("LocalServerID", System.Type.GetType("System.Int32"))
      dtLSErrors.Columns.Add("RunID", System.Type.GetType("System.Int32"))
      dtLSErrors.Columns.Add("ErrorID", System.Type.GetType("System.Int32"))
      dtLSErrors.Columns.Add("SectionID", System.Type.GetType("System.Int32"))
      dtLSErrors.Columns.Add("TagID", System.Type.GetType("System.Int32"))
      dtLSErrors.Columns.Add("HealthSeverityID", System.Type.GetType("System.Int32"))
      dtLSErrors.Columns.Add("ErrorText", System.Type.GetType("System.String"))
      dtLSErrors.Columns.Add("StatusFlag", System.Type.GetType("System.Boolean"))

      ServerNodes = XmlDoc.SelectNodes("/EXCEPTIONS/SERVER")
      Common.Write_Log(LogFile, "Data Size: " & data.Length & "    Server Count: " & ServerNodes.Count, True)
      For Each LSNode In ServerNodes
        ProcessServerNode(LSNode, dtLSHistory, dtLSErrors)
      Next
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
      ' update the local servers table with the server summary info
      PopulateLocalServers(dtLSHistory)

      If (dtLSHistory.Rows.Count > 0) Then WriteData(dtLSHistory, "LS_HealthHistory", responseText)
      If (dtLSErrors.Rows.Count > 0) Then WriteData(dtLSErrors, "LS_HealthErrors", responseText)

    Catch ex As Exception
      responseText = ex.ToString
      Common.Error_Processor()
      Common.Write_Log(LogFile, ex.ToString() & "Serial:" & LocalServerID & " MacAddress:" & MacAddress & " IP:" & LocalServerIP & " server=" & Environment.MachineName)
    Finally
    End Try

    Common.Write_Log(LogFile, "Processing Time (in ms): " & stopWatch.ElapsedMilliseconds)
    stopWatch.Stop()
    Return responseText
  End Function

  Private Sub ProcessServerNode(ByVal ServerNode As XmlNode, ByRef dtLSHistory As DataTable, ByRef dtLSErrors As DataTable)
    Dim attrib As XmlAttribute = Nothing
    Dim i As Integer = 0
    Dim LocalServerID As Integer

    ' read the attributes and store them in the server datatable
    LocalServerID = PopulateServerTag(ServerNode.Attributes, dtLSHistory)

    If (ServerNode.HasChildNodes) Then
      ' read the error child nodes and store them in the exception datatable
      PopulateErrorTags(ServerNode, dtLSErrors)
    Else
      ' clear all health error alerts reported for this local server
      ClearHealthAlerts(LocalServerID)
    End If

  End Sub

  Private Function PopulateServerTag(ByVal attributes As XmlAttributeCollection, ByRef dtLSHistory As DataTable) As Integer
    Dim LSID, RunID, StatusID, SeverityID As Integer
    Dim RunDate As Date
    Dim Sev1, Sev10 As Integer
    Dim row As DataRow

    Integer.TryParse(attributes("SERIAL").Value, LSID)
    Integer.TryParse(attributes("RUNID").Value, RunID)
    Integer.TryParse(attributes("STATUS").Value, StatusID)
    Integer.TryParse(attributes("SEVERITY").Value, SeverityID)
    Integer.TryParse(attributes("sev1").Value, Sev1)
    Integer.TryParse(attributes("sev10").Value, Sev10)

    ' validate that the date is sent and correctly formatted
    If attributes("DATE") Is Nothing OrElse attributes("DATE").Value.Trim = "" Then
      Throw New Exception("No Run Date attribute sent from the Health server for Serial " & LSID & " RunID: " & RunID)
    Else
      If Not Date.TryParse(attributes("DATE").Value, RunDate) Then
        Throw New Exception("Invalid date format of: " & attributes("DATE").Value & " for Serial " & LSID & " RunID: " & RunID)
      End If

      If RunDate < Date.Now.AddYears(-10) Then
        Throw New Exception("The Run Date sent from the Health Server is over 10 years old for Serial " & LSID & " RunID: " & RunID)
      End If
    End If

    row = dtLSHistory.NewRow()
    row.Item("LocalServerID") = LSID
    row.Item("RunID") = RunID
    row.Item("RunDate") = RunDate
    row.Item("HealthStatusID") = StatusID
    row.Item("HealthSeverityID") = SeverityID
    row.Item("Sev1") = Sev1
    row.Item("Sev10") = Sev10

    dtLSHistory.Rows.Add(row)

    Return LSID
  End Function

  Private Sub PopulateErrorTags(ByVal ServerNode As XmlNode, ByRef dtLSErrors As DataTable)
    Dim i As Integer = 0
    Dim RunID, LSID As Integer
    Dim ErrorID, SeverityID, TagID, SectionID As Integer
    Dim row As DataRow = Nothing

    Integer.TryParse(ServerNode.Attributes("SERIAL").Value, LSID)
    Integer.TryParse(ServerNode.Attributes("RUNID").Value, RunID)

    For i = 0 To ServerNode.ChildNodes.Count - 1
      row = dtLSErrors.NewRow
      row.Item("LocalServerID") = LSID
      row.Item("RunID") = RunID
      Integer.TryParse(ServerNode.ChildNodes(i).Attributes("ERRORID").Value, ErrorID)
      row.Item("ErrorID") = ErrorID
      Integer.TryParse(ServerNode.ChildNodes(i).Attributes("SECTIONID").Value, SectionID)
      row.Item("SectionID") = SectionID
      Integer.TryParse(ServerNode.ChildNodes(i).Attributes("TAGID").Value, TagID)
      row.Item("TagID") = TagID
      Integer.TryParse(ServerNode.ChildNodes(i).Attributes("SEVERITY").Value, SeverityID)
      row.Item("HealthSeverityID") = SeverityID
      row.Item("ErrorText") = ServerNode.ChildNodes(i).Attributes("TEXT").Value
      row.Item("StatusFlag") = 0
      dtLSErrors.Rows.Add(row)
    Next

  End Sub

  Private Sub PopulateLocalServers(ByVal dtLSHistory As DataTable)
    Dim row As DataRow
    Dim lsRow As DataRow
    Dim dt As New DataTable

    dt.Columns.Add("LocalServerID", System.Type.GetType("System.Int32"))
    dt.Columns.Add("LastRunID", System.Type.GetType("System.Int32"))
    dt.Columns.Add("Sev1", System.Type.GetType("System.Int32"))
    dt.Columns.Add("Sev10", System.Type.GetType("System.Int32"))

    For Each row In dtLSHistory.Rows
      lsRow = dt.NewRow()
      lsRow.Item("LocalServerID") = Common.NZ(row.Item("LocalServerID"), 0)
      lsRow.Item("LastRunID") = Common.NZ(row.Item("RunID"), 0)
      lsRow.Item("Sev1") = Common.NZ(row.Item("Sev1"), 0)
      lsRow.Item("Sev10") = Common.NZ(row.Item("Sev10"), 0)
      dt.Rows.Add(lsRow)
      dt.Rows(dt.Rows.Count - 1).AcceptChanges()
      dt.Rows(dt.Rows.Count - 1).SetModified()
    Next

    If (dt.Rows.Count > 0) Then
      BatchUpdate(dt, dt.Rows.Count)
    End If

  End Sub

  Private Sub ClearHealthAlerts(ByVal LocalServerID As Integer)

    If Not (Common.LWHadoConn.State = ConnectionState.Open) Then Common.Open_LogixWH()

    Common.QueryStr = "delete from LS_HealthAlerts with (RowLock) where LocalServerID = " & LocalServerID
    Common.LWH_Execute()

  End Sub

  Sub WriteData(ByVal dt As DataTable, ByVal TableName As String, ByRef responseText As String)
    Dim bc As SqlBulkCopy = Nothing

    Try
      If Not (Common.LWHadoConn.State = ConnectionState.Open) Then Common.Open_LogixWH()

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

      bc = New SqlBulkCopy(Common.LWHadoConn)
      bc.BatchSize = dt.Rows.Count
      bc.DestinationTableName = TableName
      bc.WriteToServer(dt)
      bc.Close()
    Catch sqlEx As SqlException
      If (sqlEx.Number = DUP_PRIMARY_KEY) Then
        responseText = "DUP"
        Common.Write_Log(LogFile, "Duplicate Run ID sent to " & TableName)
      Else
        Send(sqlEx.ToString)
        Common.Error_Processor()
        Common.Write_Log(LogFile, "serial:" & LocalServerID & " MacAddres:" & MacAddress & " IP:" & LocalServerIP & " " & sqlEx.ToString() & " server:" & Environment.MachineName)
      End If
    Catch ex As Exception
      Send(ex.ToString)
      Common.Error_Processor()
      Common.Write_Log(LogFile, ex.ToString() & "Serial:" & LocalServerID & " MacAddress:" & MacAddress & " IP:" & LocalServerIP & " server:" & Environment.MachineName)
    End Try

  End Sub

  Public Sub BatchUpdate(ByVal dataTable As DataTable, ByVal batchSize As Int32)
    Dim adapter As New SqlDataAdapter()
    Dim exceptionRows() As DataRow
    Dim row As DataRow
    Dim LogText As String = ""

    Try
      'Set the UPDATE command and parameters.
      adapter.UpdateCommand = New SqlCommand( _
        "UPDATE LocalServers with (RowLock) SET " _
        & "LastRunID=@LastRunID, Sev1Errors=@Sev1, Sev10Errors=@Sev10 where LocalServerID=@LocalServerID;", _
        Common.LRTadoConn)
      adapter.UpdateCommand.Parameters.Add("@LastRunID", _
        SqlDbType.Int, 4, "LastRunID")
      adapter.UpdateCommand.Parameters.Add("@Sev1", _
        SqlDbType.Int, 4, "Sev1")
      adapter.UpdateCommand.Parameters.Add("@Sev10", _
        SqlDbType.Int, 4, "Sev10")
      adapter.UpdateCommand.Parameters.Add("@LocalServerID", _
        SqlDbType.Int, 4, "LocalServerID")
      adapter.UpdateCommand.UpdatedRowSource = UpdateRowSource.None
      adapter.UpdateCommand.Connection = Common.LRTadoConn

      ' Set the batch size.
      adapter.UpdateBatchSize = batchSize

      adapter.Update(dataTable)
    Catch dbcEx As DBConcurrencyException
      If dbcEx.RowCount > 0 Then
        ReDim exceptionRows(dbcEx.RowCount - 1)
        dbcEx.CopyToRows(exceptionRows)
        If exceptionRows IsNot Nothing AndAlso exceptionRows.Length > 0 Then
          For Each row In exceptionRows
            LogText = "Local Server " & Common.NZ(row.Item("LocalServerID"), "") & " not found in LogixRT.LocalServers table. )"
            Common.Write_Log(LogFile, LogText, True)
          Next
        End If
      End If
    Catch ex As Exception
      Common.Write_Log(LogFile, ex.ToString, True)
    End Try

  End Sub

  Function GetOptions() As String
    Dim OptionBuf As New StringBuilder()
    Dim dt As DataTable
    Dim i, rowUBound As Integer

    Send("ACK")

    Common.QueryStr = "select Convert(nvarchar, OptionID) + ':' + OptionValue as OptionToken " & _
                      "from HealthServerOptions with (NoLock) order by OptionID;"
    dt = Common.LRT_Select()
    rowUBound = dt.Rows.Count - 1

    For i = 0 To rowUBound
      OptionBuf.Append(Common.NZ(dt.Rows(i).Item("OptionToken"), ""))
      OptionBuf.Append(vbCrLf)
    Next

    Return OptionBuf.ToString
  End Function

  Function GetMaxRunID() As String
    Dim MaxID As Long
    Dim dt As DataTable

    Send("ACK")

    Common.QueryStr = "select max(runid) as latestrunid from ls_healthhistory;"
    dt = Common.LWH_Select()

    If (dt.Rows.Count > 0) Then
      MaxID = Common.NZ(dt.Rows(0).Item("latestrunid"), 0)
    Else
      MaxID = 0
    End If

    Return MaxID.ToString
  End Function
</script>
<%
  Dim TotalTime As Decimal
  Dim Data As String = ""
  Dim Mode As String = ""

  Common.AppName = "LS-Health.aspx"
  Response.Expires = 0
  On Error GoTo ErrorTrap

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

  StartTime = DateAndTime.Timer

  Common.Open_LogixRT()
  Common.Open_LogixXS()
  Common.Open_LogixWH()
  Common.Load_System_Info()
  Connector.Load_System_Info(Common)

  Common.Write_Log(LogFile, "---------------------------------------------------------------------------")
  Common.Write_Log(LogFile, "Processing Local Server Health  Process running on server:" & Environment.MachineName, True)

  Mode = Request.QueryString("mode")

  Send_Response_Header("Server Health", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)

  If (Mode IsNot Nothing AndAlso Mode.ToUpper = "GETOPTIONS") Then
    Send(GetOptions)
  ElseIf (Mode IsNot Nothing AndAlso Mode.ToUpper = "GETMAXRUNID") Then
    Send(GetMaxRunID)
  ElseIf (Mode IsNot Nothing AndAlso Mode.ToUpper = "PROCESS") Then
    If (Handle_Post(Data)) Then
      Dim responseText As String = ProcessData(Data)
      If (responseText = "") Then
        Send("ACK")
      Else
        Send(responseText)
      End If
    Else
      Send("NAK")
      Send(Data)
    End If
  Else
    Send("NAK")
  End If

  Common.Close_LogixRT()
  Common.Close_LogixXS()
  Common.Close_LogixWH()

  TotalTime = DateAndTime.Timer - StartTime
  Common.Write_Log(LogFile, "Serial:" & LocalServerID & " MacAddress:" & MacAddress & " IP:" & LocalServerIP & " RunTime:" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")

  Response.End()
ErrorTrap:
  Response.Write(Common.Error_Processor(, , , Common.InstallationName) & "Serial:" & LocalServerID & " MacAddress:" & MacAddress & " IP:" & LocalServerIP & " server:" & Environment.MachineName)
  Common.Write_Log(LogFile, "***** An error occurred during processing! Please check the ErrorLog! *****")
  If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
  If Not (Common.LXSadoConn.State = ConnectionState.Closed) Then Common.Close_LogixXS()
  If Not (Common.LWHadoConn.State = ConnectionState.Closed) Then Common.Close_LogixWH()
  Common = Nothing
%>