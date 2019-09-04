<%@ WebService Language="VB" Class="Service" %>

Imports System
Imports System.Web
Imports System.Web.Services
Imports System.Xml.Schema
Imports System.Xml
Imports System.Data
Imports System.IO

<WebService(Namespace:="http://www.copienttech.com/IssuanceService/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class Service
  Inherits System.Web.Services.WebService
  ' version:7.3.1.138972.Official Build (SUSDAY10202)
    Private MyCommon As New Copient.CommonInc
    Private MyCryptLib As New Copient.CryptLib
  Private LogFile As String = "IssuanceWSLog." & Date.Now.ToString("yyyyMMdd") & ".txt"
  Private Const CONNECTOR_ID As Integer = 41
  Private Const MAX_RECORDS As Integer = 10000

  Public Enum ReturnCodes As Integer
    SUCCESS = 0
    INVALID_GUID = 1
    INVALID_DATE = 2
    INVALID_CRITERIA = 3
    MISSING_EXT_CRM_INTERFACE_ID = 4
    INVALID_LASTPKID = 5
    INVALID_EXTCRMINTERFACE = 6
    APPLICATION_EXCEPTION = 9999
  End Enum

  <WebMethod()> _
  Public Function GetIssuanceColumns(ByVal GUID As String) As String
    Dim ColBuf As New StringBuilder("")
    Dim RetCode As ReturnCodes = ReturnCodes.SUCCESS
    Dim RetMsg As String = ""
    Dim dt As DataTable
    Dim col As DataColumn
    Dim i As Integer = 0
    Dim Done As Boolean = False

    Try
      If Not IsValidGUID(GUID, "GetIssuanceColumns") Then
        RetCode = ReturnCodes.INVALID_GUID
        RetMsg = "ERROR: " & GetReturnCodeText(RetCode) & ControlChars.CrLf & "DESC: " & GUID
      Else
        If MyCommon.LEXadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixEX()

        While Not Done
          Try
            MyCommon.QueryStr = "select top 1 * from Issuance" & Date.Now.AddDays(i).ToString("yyyyMMdd") & " with (NoLock)"
            dt = MyCommon.LEX_Select
            If dt.Columns.Count > 0 Then
              For Each col In dt.Columns
                If ColBuf.Length > 0 Then ColBuf.Append(",")
                ColBuf.Append(col.ColumnName)
              Next
            End If
            Done = True
          Catch ex As Exception
            ' search back through the last ten days, if not found then return error code
            If i < -9 Then
              Done = True
              RetCode = ReturnCodes.INVALID_DATE
              RetMsg = "ERROR: " & GetReturnCodeText(RetCode) & ControlChars.CrLf & "DESC: Unable to find any issuance tables for the last 10 days"
            Else
              ' look for the previous days table
              i -= 1
            End If
          End Try
        End While

      End If
    Catch ex As Exception
      RetCode = ReturnCodes.APPLICATION_EXCEPTION
      RetMsg = "ERROR: " & GetReturnCodeText(RetCode) & ControlChars.CrLf & "DESC: " & ex.ToString
    Finally
      If MyCommon.LEXadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixEX()
    End Try

    If RetCode <> ReturnCodes.SUCCESS Then
      ColBuf = New StringBuilder(RetMsg)
    End If

    Return ColBuf.ToString
  End Function

  <WebMethod()> _
  Public Function GetNextAsString(ByVal GUID As String, ByVal Criteria As String) As String
    Dim CriteriaXML As New XmlDocument
    Dim RetXML As New XmlDocument
    Dim RetCode As ReturnCodes = ReturnCodes.SUCCESS
    Dim RetMsg As String = ""

    Try
      CriteriaXML.LoadXml(Criteria)
    Catch ex As Exception
      RetCode = ReturnCodes.INVALID_CRITERIA
      RetMsg = "Request failed. Invalid Criteria: Exception=" & ex.ToString
      MyCommon.Write_Log(LogFile, RetMsg, True)
    End Try

    If RetCode = ReturnCodes.SUCCESS Then
      RetXML = GetNext(GUID, CriteriaXML)
    Else
      RetXML = GetErrorXML(GetReturnCodeText(RetCode), RetMsg)
    End If

    Return RetXML.OuterXml
  End Function

  <WebMethod()> _
  Public Function GetNext(ByVal GUID As String, ByVal Criteria As XmlDocument) As XmlDocument
    Dim XmlDoc As New XmlDocument()
    Dim ms As New MemoryStream
    Dim Writer As New XmlTextWriter(ms, System.Text.Encoding.UTF8)
    Dim RetCode As ReturnCodes = ReturnCodes.SUCCESS
    Dim dt As DataTable = New DataTable()
    Dim dt2 As DataTable = New DataTable()
    Dim row As DataRow
    Dim ExtCRMInterface As String = ""
    Dim ColumnList As String = ""
    Dim ValidColumnList As String = ""
    Dim TableDate As Date
    Dim LastPKID As Long = 0
    Dim NewLastPKID As Long = 0
    Dim NewTableDate As Date
    Dim RecordLimit As Integer = 0
    Dim RecordCount As Integer = 0
    Dim DayOffset As Integer = 0
    Dim attrib As XmlAttribute = Nothing
    Dim LoopCtr As Integer = 0

    Try
      If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
      If MyCommon.LEXadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixEX()

      If Not IsValidGUID(GUID, "GetNext") Then
        RetCode = ReturnCodes.INVALID_GUID
      ElseIf Not IsValidXmlDocument("IssuanceCriteria.xsd", Criteria) Then
        RetCode = ReturnCodes.INVALID_CRITERIA
      Else
        ExtCRMInterface = ParseExtCRMInterface(Criteria, RetCode)
        ColumnList = ParseColumns(GUID, Criteria)
        MyCommon.Write_Log(LogFile, "columnList: " & ColumnList, True)
        RecordLimit = ParseRecordLimit(Criteria, MAX_RECORDS)
        MyCommon.Write_Log(LogFile, "RecordLimit: " & RecordLimit, True)
        GetStartPosition(ExtCRMInterface, TableDate, LastPKID, GUID)
        MyCommon.Write_Log(LogFile, "TableDate: " & TableDate & " LastPKID: " & LastPKID, True)
        NewLastPKID = LastPKID
        NewTableDate = TableDate

        ' if everything validated then let's write the XML document
        If RetCode = ReturnCodes.SUCCESS Then
          Writer.Formatting = Formatting.Indented
          Writer.Indentation = 4

          Writer.WriteStartDocument()
          Writer.WriteStartElement("IssuanceData")

          ' ensure that the table is still in the LogixEX database
          Try
            If Not IsValidTable("Issuance" & TableDate.ToString("yyyyMMdd")) Then
              TableDate = TableDate.AddDays(1)
              While Not IsValidTable("Issuance" & TableDate.ToString("yyyyMMdd"))
                MyCommon.Write_Log(LogFile, "Checking for table: " & "Issuance" & TableDate.ToString("yyyyMMdd"), True)
                LoopCtr += 1
                TableDate = TableDate.AddDays(1)
                If TableDate > Date.Now OrElse LoopCtr > 3650 Then Exit While
              End While
            Else
              MyCommon.Write_Log(LogFile, "Table is valid: " & "Issuance" & TableDate.ToString("yyyyMMdd"), True)
            End If
          Catch ex As Exception
            MyCommon.Write_Log(LogFile, ex.ToString, True)
          End Try

          Try
            ValidColumnList = CleanColumnList(ColumnList, "Issuance" & TableDate.ToString("yyyyMMdd"))
            MyCommon.QueryStr = "select top " & RecordLimit & " PKID as LastPKID, " & _
                                "  '" & TableDate.ToString("yyyyMMdd") & "' as LastTableDate, " & ValidColumnList & " " & _
                                "from Issuance" & TableDate.ToString("yyyyMMdd") & " with (NoLock) " & _
                                "where ExtCRMInterface=@ExtCRMInterface and PKID > @LastPKID " & _
                                "order by LastPKID"
            MyCommon.DBParameters.Add("@ExtCRMInterface", SqlDbType.NVarChar, 100).Value = ExtCRMInterface
            MyCommon.DBParameters.Add("@LastPKID", SqlDbType.Int).Value = LastPKID
            dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixEX)
            RecordCount = dt.Rows.Count
          Catch ex As Exception
            ' The table for the tabledate doesn't exist
            RecordCount = 0
          End Try

          ' load up data from the subsequent days tables, if necessary
          DayOffset = 1
          While RecordCount < RecordLimit AndAlso DayOffset <= 3650 AndAlso TableDate.AddDays(DayOffset) <= Date.Now
            Try
              ValidColumnList = CleanColumnList(ColumnList, "Issuance" & TableDate.AddDays(DayOffset).ToString("yyyyMMdd"))
              MyCommon.QueryStr = "select top " & (RecordLimit - RecordCount) & " PKID as LastPKID, " & _
                                  "  '" & TableDate.AddDays(DayOffset).ToString("yyyyMMdd") & "' as LastTableDate, " & ValidColumnList & " " & _
                                  "from Issuance" & TableDate.AddDays(DayOffset).ToString("yyyyMMdd") & " with (NoLock) " & _
                                  "where ExtCRMInterface=@ExtCRMInterface " & _
                                  "order by LastPKID;"
              MyCommon.DBParameters.Add("@ExtCRMInterface", SqlDbType.NVarChar, 100).Value = ExtCRMInterface
              dt2 = MyCommon.ExecuteQuery(Copient.DataBases.LogixEX)
              dt.Merge(dt2)
              NewTableDate = TableDate.AddDays(DayOffset)
              RecordCount = dt.Rows.Count
            Catch ex As Exception
              MyCommon.Write_Log(LogFile, "Table Issuance" & TableDate.AddDays(DayOffset).ToString("yyyyMMdd") & " not found.", True)
              ' the table doesn't exist...increment dayoffset by 1 and check if the TableDate is equal to CurrentDate. If not then keep looking for
              ' next tables.
            End Try
            DayOffset += 1
          End While

          If RecordCount > 0 Then
            ' write each of the rows to the XML file
            For Each row In dt.Rows
              WriteIssuanceRow(Writer, row)
            Next

            ' get the LastPKID for later update - reset to 0 if the last table found has no records in it yet.
            If NewTableDate.ToString("yyyyMMdd") > MyCommon.NZ(dt.Rows(dt.Rows.Count - 1).Item("LastTableDate"), "!") Then
              NewLastPKID = 0
            Else
              NewLastPKID = MyCommon.NZ(dt.Rows(dt.Rows.Count - 1).Item("LastPKID"), 0)
            End If
          Else
            NewLastPKID = 0
          End If

          Writer.WriteEndElement() ' end IssuanceData
          Writer.WriteEndDocument()
          Writer.Flush()
        End If
      End If

      If RetCode = ReturnCodes.SUCCESS Then
        ms.Seek(0, SeekOrigin.Begin)
        XmlDoc.Load(ms)

        attrib = XmlDoc.CreateAttribute("statusCode")
        attrib.Value = GetReturnCodeText(RetCode)
        attrib = XmlDoc.SelectSingleNode("//IssuanceData").Attributes.Append(attrib)

        attrib = XmlDoc.CreateAttribute("message")
        attrib.Value = IIf(RecordCount = 0, "No new issuance data to report", RecordCount & " rows of issuance data returned ") & " for " & ExtCRMInterface & "."
        attrib = XmlDoc.SelectSingleNode("//IssuanceData").Attributes.Append(attrib)

        ' update the table with the last PK retrieved by this call
        If RecordCount > 0 Then
          UpdateLastPKID(ExtCRMInterface, NewLastPKID, RecordCount, NewTableDate, GUID)
        End If
      Else
        ' write the ErrorXML
        MyCommon.Write_Log(LogFile, "Request failed. RetCode=" & GetReturnCodeText(RetCode), True)
        XmlDoc = GetErrorXML(GetReturnCodeText(RetCode), "")
      End If

    Catch ex As Exception
      ' write the ErrorXML
      MyCommon.Write_Log(LogFile, "Request failed. Exception=" & ex.ToString, True)
      XmlDoc = GetErrorXML(GetReturnCodeText(ReturnCodes.APPLICATION_EXCEPTION), ex.ToString)
    Finally
      If Not Writer Is Nothing Then
        Writer.Close()
      End If
      If Not ms Is Nothing Then
        ms.Close()
        ms.Dispose()
        ms = Nothing
      End If

      If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
      If MyCommon.LEXadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixEX()
    End Try

    Return XmlDoc
  End Function

  <WebMethod()> _
  Public Function SetLastPosition(ByVal GUID As String, ByVal ExtCRMInterface As String, _
                                  ByVal LastTableDate As Date, ByVal LastPKID As Long) As String
    Dim RetCode As ReturnCodes = ReturnCodes.SUCCESS
    Dim RetMsg As String = ""
    Dim dt As DataTable

    Try
      If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
      If MyCommon.LEXadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixEX()

      If Not IsValidGUID(GUID, "SetLastPosition") Then
        RetCode = ReturnCodes.INVALID_GUID
      ElseIf LastPKID < 0 Then
        RetCode = ReturnCodes.INVALID_LASTPKID
      ElseIf Not IsValidExtCRMInterface(ExtCRMInterface) Then
        RetCode = ReturnCodes.INVALID_EXTCRMINTERFACE
      Else
        MyCommon.QueryStr = "dbo.pa_Issuance_SetPosition"
        MyCommon.Open_LEXsp()
        MyCommon.LEXsp.Parameters.Add("@ExtCRMInterface", SqlDbType.NVarChar, 100).Value = ExtCRMInterface
        MyCommon.LEXsp.Parameters.Add("@LastTableDate", SqlDbType.DateTime).Value = LastTableDate
        MyCommon.LEXsp.Parameters.Add("@LastPKID", SqlDbType.BigInt).Value = LastPKID
        MyCommon.LEXsp.Parameters.Add("@GUID", SqlDbType.NVarChar, 36).Value = GUID
        MyCommon.LEXsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
        dt = MyCommon.LEXsp_select
        RetCode = MyCommon.LEXsp.Parameters("@Status").Value
        MyCommon.Close_LEXsp()
      End If
      If RetCode = ReturnCodes.SUCCESS Then
        RetMsg = "OK"
      Else
        RetMsg = "ERROR: " & GetReturnCodeText(RetCode) & ControlChars.CrLf & "DESC: Unable to update the IssuanceRetrieval table for source " & ExtCRMInterface
      End If
    Catch ex As Exception
      RetMsg = "ERROR: " & GetReturnCodeText(ReturnCodes.APPLICATION_EXCEPTION) & ControlChars.CrLf & "DESC: " & ex.ToString
    Finally
      If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
      If MyCommon.LEXadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixEX()
    End Try

    Return RetMsg
  End Function

  Private Function IsValidExtCRMInterface(ByVal ExtCRMInterface As String) As Boolean
    MyCommon.QueryStr = "Select Name from ExtCRMInterfaces where Name = @Name"
    MyCommon.DBParameters.Add("@Name", SqlDbType.NVarChar).Value = ExtCRMInterface
    Dim dt As DataTable = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
    Return dt.Rows.Count > 0
  End Function

  Private Function IsValidGUID(ByVal GUID As String, ByVal MethodName As String) As Boolean
    Dim IsValid As Boolean = False
    Dim ConnInc As New Copient.ConnectorInc
    Dim MsgBuf As New StringBuilder()

    Try
      If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
      IsValid = ConnInc.IsValidConnectorGUID(MyCommon, CONNECTOR_ID, GUID)
    Catch ex As Exception
      IsValid = False
    End Try

    ' Log the call
    Try
      If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
      MsgBuf.Append(IIf(IsValid, "Validated call to ", "Invalid call to "))
      MsgBuf.Append(MethodName)
      MsgBuf.Append(" from GUID: ")
      MsgBuf.Append(GUID)
      MsgBuf.Append(" and IP: " & HttpContext.Current.Request.UserHostAddress)
      MyCommon.Write_Log(LogFile, MsgBuf.ToString, True)
    Catch ex As Exception
      ' ignore
    End Try

    Return IsValid
  End Function

  Private Function IsValidTable(ByVal TableName As String) As Boolean
    Dim IsValid As Boolean = False
    Dim dt As DataTable

    If MyCommon.LEXadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixEX()

    Try
      MyCommon.QueryStr = "select top 1 * from " & TableName & " with (NoLock)"
      dt = MyCommon.LEX_Select
      IsValid = True
    Catch ex As Exception
      IsValid = False
    End Try

    Return IsValid
  End Function

  Private Function IsValidXmlDocument(ByVal sXsdFileName As String, ByVal XmlDoc As XmlDocument) As Boolean

    Dim Settings As XmlReaderSettings
    Dim xr As XmlReader = Nothing
    Dim ms As New MemoryStream()
    Dim sMsg As String = ""
    Dim bValid As Boolean = True
    Dim xsdPath As String

    Try
      If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

      xsdPath = MyCommon.Get_Install_Path & "AgentFiles\" & sXsdFileName
      If sXsdFileName <> "" AndAlso XmlDoc IsNot Nothing Then
        Settings = New XmlReaderSettings()
        Settings.Schemas.Add(Nothing, xsdPath)
        Settings.ValidationType = ValidationType.Schema
        Settings.IgnoreComments = True
        Settings.IgnoreProcessingInstructions = True
        Settings.IgnoreWhitespace = True

        XmlDoc.Save(ms)
        ms.Seek(0, SeekOrigin.Begin)

        xr = XmlReader.Create(ms, Settings)
        Do While (xr.Read())
          'Console.WriteLine("NodeType: " & xr.NodeType.ToString & " - " & xr.LocalName & " Depth: " & xr.Depth.ToString)
        Loop
        bValid = True
      Else
        bValid = False
      End If
    Catch eXmlSch As XmlSchemaException
      sMsg = "(Xml Schema Validation Error Line: " & eXmlSch.LineNumber.ToString & " - Col: " & eXmlSch.LinePosition.ToString & ") " & eXmlSch.Message
      bValid = False
    Catch eXml As XmlException
      sMsg = "(Xml Error Line: " & eXml.LineNumber.ToString & " - Col: " & eXml.LinePosition.ToString & ") " & eXml.Message
      bValid = False
    Catch exApp As ApplicationException
      sMsg = "Application Error: " & exApp.ToString
      bValid = False
    Catch ex As Exception
      sMsg = "Error: " & ex.ToString
      bValid = False
    Finally
      If Not xr Is Nothing Then
        xr.Close()
      End If
      If Not ms Is Nothing Then
        ms.Close()
        ms.Dispose()
        ms = Nothing
      End If
    End Try

    ' Log Error if one exists
    If (sMsg <> "") Then
      MyCommon.Write_Log(LogFile, sMsg, True)
    End If

    Return bValid
  End Function

  Private Function ParseExtCRMInterface(ByVal Criteria As XmlDocument, ByRef RetCode As ReturnCodes) As String
    Dim ExtCRMInterface As String = ""
    Dim TempNode As XmlNode

    ' only validate the external CRM interface if there are current no errors
    If Criteria IsNot Nothing AndAlso RetCode = ReturnCodes.SUCCESS Then
      TempNode = Criteria.SelectSingleNode("/SearchCriteria/ExtCRMInterface")
      If TempNode Is Nothing OrElse TempNode.InnerText.Trim = "" Then
        RetCode = ReturnCodes.MISSING_EXT_CRM_INTERFACE_ID
      Else
        ExtCRMInterface = TempNode.InnerText.Trim
      End If
    End If

    Return ExtCRMInterface
  End Function

  Private Function ParseRecordLimit(ByVal Criteria As XmlDocument, ByVal MaxRecordCount As Integer) As Integer
    Dim RecordLimit As Integer = 0
    Dim TempNode As XmlNode

    If Criteria IsNot Nothing Then
      TempNode = Criteria.SelectSingleNode("/SearchCriteria/RecordLimit")
      If TempNode Is Nothing Then
        RecordLimit = MaxRecordCount
      Else
        Integer.TryParse(TempNode.InnerText.Trim, RecordLimit)
      End If
    End If

    If RecordLimit <= 0 Or RecordLimit > MaxRecordCount Then
      RecordLimit = MaxRecordCount
    End If

    Return RecordLimit
  End Function

  Private Function ParseColumns(ByVal GUID As String, ByVal Criteria As XmlDocument) As String
    Dim TempNodeList As XmlNodeList
    Dim TempNode As XmlNode
    Dim Columns(-1) As String
    Dim ColumnList As String = ""
    Dim FullList As String = ""
    Dim i As Integer

    FullList = GetIssuanceColumns(GUID)

    ' only validate the column list if there are no current errors
    If Criteria IsNot Nothing Then
      TempNodeList = Criteria.SelectNodes("/SearchCriteria/Columns/Name")
      If TempNodeList IsNot Nothing AndAlso TempNodeList.Count > 0 Then
        ReDim Columns(TempNodeList.Count - 1)
        i = 0
        For Each TempNode In TempNodeList
          Columns(i) = TempNode.InnerText.Trim
          i += 1
        Next
        ColumnList = String.Join(",", Columns)
      End If
    End If

    If ColumnList = "" Or ColumnList.Trim = "," Then
      ColumnList = FullList
    End If

    Return ColumnList
  End Function

  Private Sub GetStartPosition(ByVal ExtCRMInterface As String, ByRef TableDate As Date, ByRef LastPKID As Long, ByVal GUID As String)
    Dim dt As DataTable
    Dim dt1 As DataTable

    If MyCommon.LEXadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixEX()

    TableDate = Date.Now
    LastPKID = 0

    MyCommon.QueryStr = "select LastTableDate,LastPKID, LastCalled, LastRecordCount from IssuanceRetrieval with (NoLock) " & _
                        "where ExtCRMInterface=@ExtCRMInterface and GUID=@GUID;"
    MyCommon.DBParameters.Add("@ExtCRMInterface", SqlDbType.NVarChar, 100).Value = ExtCRMInterface
    MyCommon.DBParameters.Add("@GUID", SqlDbType.NVarChar, 36).Value = GUID
    dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixEX)
    If dt.Rows.Count > 0 Then
      TableDate = MyCommon.NZ(dt.Rows(0).Item("LastTableDate"), Date.Now)
      LastPKID = MyCommon.NZ(dt.Rows(0).Item("LastPKID"), 0)
    Else
      MyCommon.QueryStr = "select LastTableDate,LastPKID, LastCalled, LastRecordCount from IssuanceRetrieval with (NoLock) " & _
                      "where ExtCRMInterface=@ExtCRMInterface;"
      MyCommon.DBParameters.Add("@ExtCRMInterface", SqlDbType.NVarChar, 100).Value = ExtCRMInterface
      dt1 = MyCommon.ExecuteQuery(Copient.DataBases.LogixEX)
      If dt1.Rows.Count > 0 Then
        TableDate = MyCommon.NZ(dt1.Rows(0).Item("LastTableDate"), Date.Now)
        LastPKID = MyCommon.NZ(dt1.Rows(0).Item("LastPKID"), 0)
      End If
    End If

  End Sub

  Private Sub WriteIssuanceRow(ByRef Writer As XmlWriter, ByRef row As DataRow)
    Dim dt As DataTable
    Dim i As Integer = 0
    Dim Value As String = ""
    Dim TempDate As Date
    Dim ColType As System.Type

    If row IsNot Nothing Then
      dt = row.Table
      Writer.WriteStartElement("Row")

      ' skip the first two columns (LastPKID, LastTableDate) as these are used only by the application and not sent to the client
      For i = 2 To dt.Columns.Count - 1
        ColType = dt.Columns(i).DataType
        If IsDBNull(row.Item(i)) AndAlso (ColType Is GetType(Long) OrElse ColType Is GetType(Decimal) OrElse ColType Is GetType(Integer)) Then
          Value = 0
        ElseIf IsDBNull(row.Item(i)) AndAlso (ColType Is GetType(DateTime)) Then
          Value = "1980-01-01T00:00:00"
        ElseIf IsDBNull(row.Item(i)) Then
          Value = ""
        ElseIf ColType Is GetType(Boolean) Then
          Value = row.Item(i).ToString.ToLower
        Else
          If dt.Columns(i).DataType Is GetType(DateTime) Then
            TempDate = row.Item(i)
            Value = TempDate.ToString("yyyy-MM-ddTHH:mm:ss")
          Else
                        If UCase(dt.Columns(i).ColumnName) = "PRIMARYEXTID" Or UCase(dt.Columns(i).ColumnName) = "RESOLVEDCUSTOMERID" or UCase(dt.Columns(i).ColumnName) = "HHID" Then
                            Value = MyCryptLib.SQL_StringDecrypt(row.Item(i).ToString).Trim
                        Else
                            Value = row.Item(i).ToString.Trim
                        End If
                    End If
        End If

        Writer.WriteElementString(dt.Columns(i).ColumnName, Value)
      Next
      Writer.WriteEndElement() ' end row
    End If
  End Sub

  Private Function UpdateLastPKID(ByVal ExtCRMInterface As String, ByVal LastPKID As Long, _
                                  ByVal LastRecordCount As Integer, ByVal LastTableDate As Date, ByVal GUID As String) As Boolean
    Dim Updated As Boolean = False
    Dim dt As DataTable

    If MyCommon.LEXadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixEX()

    MyCommon.QueryStr = "select ExtCRMInterface, GUID from IssuanceRetrieval with (NoLock) " & _
                    "where ExtCRMInterface = @ExtCRMInterface and GUID=@GUID;"
    MyCommon.DBParameters.Add("@ExtCRMInterface", SqlDbType.NVarChar, 100).Value = ExtCRMInterface
    MyCommon.DBParameters.Add("@GUID", SqlDbType.NVarChar, 36).Value = GUID
    dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixEX)
    If dt.Rows.Count > 0 Then
      MyCommon.QueryStr = "update IssuanceRetrieval with (RowLock) set LastPKID=@LastPKID, LastCalled=getdate(), LastRecordCount=@LastRecordCount," & _
                            "LastTableDate=@LastTableDate where ExtCRMInterface = @ExtCRMInterface and GUID=@GUID;"
      MyCommon.DBParameters.Add("@LastPKID", SqlDbType.BigInt).Value = LastPKID
      MyCommon.DBParameters.Add("@LastRecordCount", SqlDbType.Int).Value = LastRecordCount
      MyCommon.DBParameters.Add("@LastTableDate", SqlDbType.DateTime).Value = LastTableDate.ToString("yyyy-MM-dd")
      MyCommon.DBParameters.Add("@ExtCRMInterface", SqlDbType.NVarChar, 100).Value = ExtCRMInterface
      MyCommon.DBParameters.Add("@GUID", SqlDbType.NVarChar, 36).Value = GUID
      MyCommon.ExecuteNonQuery(Copient.DataBases.LogixEX)
    Else
      MyCommon.QueryStr = "insert into IssuanceRetrieval with (RowLock) (LastPKID, LastCalled, LastRecordCount, LastTableDate, ExtCRMInterface, GUID)" & _
                          "values (@LastPKID, getdate(), @LastRecordCount,@LastTableDate,@ExtCRMInterface, @GUID);"
      MyCommon.DBParameters.Add("@LastPKID", SqlDbType.BigInt).Value = LastPKID
      MyCommon.DBParameters.Add("@LastRecordCount", SqlDbType.Int).Value = LastRecordCount
      MyCommon.DBParameters.Add("@LastTableDate", SqlDbType.DateTime).Value = LastTableDate.ToString("yyyy-MM-dd")
      MyCommon.DBParameters.Add("@ExtCRMInterface", SqlDbType.NVarChar, 100).Value = ExtCRMInterface
      MyCommon.DBParameters.Add("@GUID", SqlDbType.NVarChar, 36).Value = GUID
      MyCommon.ExecuteNonQuery(Copient.DataBases.LogixEX)
    End If
    Updated = (MyCommon.RowsAffected > 0)

    Return Updated
  End Function

  Private Function GetErrorXML(ByVal StatusCode As String, ByVal Message As String) As XmlDocument
    Dim XmlDoc As New XmlDocument
    Dim XmlBuf As New StringBuilder()

    XmlBuf.Append("<IssuanceData statusCode=""" & StatusCode & """")
    If Message.Trim <> "" Then
      XmlBuf.Append(" message=""" & Message & """")
    End If
    XmlBuf.Append(">")
    XmlBuf.Append("</IssuanceData>")

    XmlDoc.LoadXml(XmlBuf.ToString)

    Return XmlDoc
  End Function

  Private Function GetReturnCodeText(ByVal RetCode As ReturnCodes) As String
    Dim RetCodeText As String = "SUCCESS"

    Select Case RetCode
      Case ReturnCodes.SUCCESS
        RetCodeText = "SUCCESS"
      Case ReturnCodes.APPLICATION_EXCEPTION
        RetCodeText = "APPLICATION_EXCEPTION"
      Case ReturnCodes.INVALID_CRITERIA
        RetCodeText = "INVALID_CRITERIA"
      Case ReturnCodes.INVALID_DATE
        RetCodeText = "INVALID_DATE"
      Case ReturnCodes.INVALID_GUID
        RetCodeText = "INVALID_GUID"
      Case ReturnCodes.MISSING_EXT_CRM_INTERFACE_ID
        RetCodeText = "MISSING_EXT_CRM_INTERFACE_ID"
      Case ReturnCodes.INVALID_LASTPKID
        RetCodeText = "INVALID_LASTPKID"
      Case ReturnCodes.INVALID_EXTCRMINTERFACE
        RetCodeText = "INVALID_EXTCRMINTERFACE"
    End Select
    Return RetCodeText
  End Function

  Private Function CleanColumnList(ByVal ColumnList As String, ByVal TableName As String) As String
    Dim CleanColBuf As New StringBuilder(200)
    Dim RequestedCols(-1) As String
    Dim TableCols(-1) As String
    Dim dt As DataTable
    Dim i, j As Integer
    Dim ColFound As Boolean = False

    Try
      If ColumnList IsNot Nothing AndAlso TableName IsNot Nothing Then
        ' put the requested columns in an array for later comparison to the table columns array
        RequestedCols = ColumnList.Split(",")

        ' get all the valid table columns and put them in an array
        If MyCommon.LEXadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixEX()
        MyCommon.QueryStr = "select top 1 * from " & TableName & " with (NoLock)"
        dt = MyCommon.LEX_Select

        If dt.Columns.Count > 0 Then
          ReDim TableCols(dt.Columns.Count - 1)
          For i = 0 To TableCols.GetUpperBound(0)
            TableCols(i) = dt.Columns(i).ColumnName
          Next
        End If

        ' build up clean column list for valid table columns and log columns that aren't present in the table
        For i = 0 To RequestedCols.GetUpperBound(0)
          ColFound = False
          For j = 0 To TableCols.GetUpperBound(0)
            If RequestedCols(i) = TableCols(j) Then
              ColFound = True
              Exit For
            End If
          Next

          If Not ColFound Then
            MyCommon.Write_Log(LogFile, "Column " & RequestedCols(i) & " does not exist in table " & TableName, True)
          Else
            If CleanColBuf.Length > 0 Then CleanColBuf.Append(",")
            CleanColBuf.Append(RequestedCols(i))
          End If
        Next

      End If
    Catch ex As Exception
      Throw ex
    End Try

    Return CleanColBuf.ToString
  End Function

End Class