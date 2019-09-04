<%@ WebService Language="VB" Class="Service" %>

Imports System
Imports System.Web
Imports System.Web.Services
Imports System.Xml.Schema
Imports System.Xml
Imports System.Data
Imports System.IO
Imports Copient.CryptLib

<WebService(Namespace:="http://www.copienttech.com/CustomerData/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class Service
  Inherits System.Web.Services.WebService
  ' version:7.3.1.138972.Official Build (SUSDAY10202)
  Private MyCommon As New Copient.CommonInc
  Private LogFile As String = "CustomerDataWSLog." & Date.Now.ToString("yyyyMMdd") & ".txt"
Private MyCryptlib As New Copient.CryptLib
  Public Enum StatusCodes As Integer
    SUCCESS = 0
    INVALID_GUID = 1
    TIME_LOCKED = 2
    APPLICATION_EXCEPTION = 9999
  End Enum

  Private Enum UPDATE_TYPES As Integer
    UNPROCESSED = 1
    ALREADY_PROCESSED = 2
    DATE_RANGE = 3
  End Enum

  <WebMethod()> _
  Public Function GetLogixCustomerUpdates(ByVal GUID As String) As XmlDocument
    Dim RetXmlDoc As New XmlDocument
    Dim Args(-1) As Object
    Dim RetCode As StatusCodes = StatusCodes.SUCCESS
    Dim RetMsg As String = ""

    If GUID.Trim = "" Then
      'Invalid GUID
      RetCode = StatusCodes.INVALID_GUID
      RetMsg = "Invalid GUID"
      MyCommon.Write_Log(LogFile, RetMsg, True)
      RetXmlDoc = New XmlDocument()
      RetXmlDoc.LoadXml(GetErrorXML(RetCode, RetMsg))
      Return RetXmlDoc
      Exit Function
    ElseIf GUID.Contains("'") Or GUID.Contains(Chr(34)) = True Or GUID.Contains(">") = True Then
      'Invalid GUID
      RetCode = StatusCodes.INVALID_GUID
      RetMsg = "Invalid GUID"
      MyCommon.Write_Log(LogFile, RetMsg, True)
      RetXmlDoc = New XmlDocument()
      RetXmlDoc.LoadXml(GetErrorXML(RetCode, RetMsg))
      Return RetXmlDoc
      Exit Function
    End If

    RetXmlDoc = WriteCustomerUpdates(GUID, UPDATE_TYPES.UNPROCESSED, Args)

    Return RetXmlDoc
  End Function

  <WebMethod()> _
  Public Function GetLogixCustomerUpdatesByBatch(ByVal GUID As String, ByVal BatchID As String) As XmlDocument
    Dim RetXmlDoc As New XmlDocument
    Dim Args(0) As Object
    Dim RetCode As StatusCodes = StatusCodes.SUCCESS
    Dim RetMsg As String = ""

    If GUID.Trim = "" Then
      'Invalid GUID
      RetCode = StatusCodes.INVALID_GUID
      RetMsg = "Invalid GUID"
      MyCommon.Write_Log(LogFile, RetMsg, True)
      RetXmlDoc = New XmlDocument()
      RetXmlDoc.LoadXml(GetErrorXML(RetCode, RetMsg))
      Return RetXmlDoc
      Exit Function
    ElseIf GUID.Contains("'") Or GUID.Contains(Chr(34)) = True Or GUID.Contains(">") = True Then
      'Invalid GUID
      RetCode = StatusCodes.INVALID_GUID
      RetMsg = "Invalid GUID"
      MyCommon.Write_Log(LogFile, RetMsg, True)
      RetXmlDoc = New XmlDocument()
      RetXmlDoc.LoadXml(GetErrorXML(RetCode, RetMsg))
      Return RetXmlDoc
      Exit Function
    End If

    If BatchID.Trim = "" Then
      'BatchID not provided
      RetCode = StatusCodes.APPLICATION_EXCEPTION
      RetMsg = "Invalid BatchID"
      MyCommon.Write_Log(LogFile, RetMsg, True)
      RetXmlDoc = New XmlDocument()
      RetXmlDoc.LoadXml(GetErrorXML(RetCode, RetMsg))
      Return RetXmlDoc
            Exit Function
        ElseIf BatchID.Contains("'") Or BatchID.Contains(Chr(34)) = True Or BatchID.Contains(">") = True Then
            'Invalid BatchID
            RetCode = StatusCodes.APPLICATION_EXCEPTION
            RetMsg = "Invalid BatchID"
            MyCommon.Write_Log(LogFile, RetMsg, True)
            RetXmlDoc = New XmlDocument()
            RetXmlDoc.LoadXml(GetErrorXML(RetCode, RetMsg))
            Return RetXmlDoc
            Exit Function
    End If

    Args(0) = BatchID
    RetXmlDoc = WriteCustomerUpdates(GUID, UPDATE_TYPES.ALREADY_PROCESSED, Args)

    Return RetXmlDoc
  End Function

  <WebMethod()> _
  Public Function GetLogixCustomerUpdatesByDates(ByVal GUID As String, ByVal StartDateTime As String, ByVal EndDateTime As String) As XmlDocument
    Dim RetXmlDoc As New XmlDocument
    Dim Args(1) As Object

    'check if the input parameters are valid for update by dates
    Dim RetCode As StatusCodes = StatusCodes.SUCCESS
    Dim RetMsg As String = ""

    If GUID.Trim = "" Then
      'Invalid GUID
      RetCode = StatusCodes.INVALID_GUID
      RetMsg = "Invalid GUID"
      MyCommon.Write_Log(LogFile, RetMsg, True)
      RetXmlDoc = New XmlDocument()
      RetXmlDoc.LoadXml(GetErrorXML(RetCode, RetMsg))
      Return RetXmlDoc
      Exit Function
    ElseIf GUID.Contains("'") Or GUID.Contains(Chr(34)) = True Or GUID.Contains(">") = True Then
      'Invalid GUID
      RetCode = StatusCodes.INVALID_GUID
      RetMsg = "Invalid GUID"
      MyCommon.Write_Log(LogFile, RetMsg, True)
      RetXmlDoc = New XmlDocument()
      RetXmlDoc.LoadXml(GetErrorXML(RetCode, RetMsg))
      Return RetXmlDoc
      Exit Function
    End If

    If Not IsValidDate(StartDateTime) OrElse Not IsValidDate(EndDateTime) Then
      RetCode = StatusCodes.APPLICATION_EXCEPTION
      If Not IsValidDate(StartDateTime) Then
        RetMsg = "StartDateTime is invalid."
      Else
        RetMsg = "EndDateTime is invalid."
      End If
      MyCommon.Write_Log(LogFile, RetMsg, True)
      RetXmlDoc = New XmlDocument()
      RetXmlDoc.LoadXml(GetErrorXML(RetCode, RetMsg))
    Else
      Args(0) = System.DateTime.Parse(StartDateTime)
      Args(1) = System.DateTime.Parse(EndDateTime)
      RetXmlDoc = WriteCustomerUpdates(GUID, UPDATE_TYPES.DATE_RANGE, Args)
    End If
    Return RetXmlDoc
  End Function

  Private Function WriteCustomerUpdates(ByVal GUID As String, ByVal UpdateType As UPDATE_TYPES, ByVal Args() As Object) As XmlDocument
    Dim RetXmlDoc As New XmlDocument
    Dim ms As New MemoryStream
    Dim Writer As New XmlTextWriter(ms, System.Text.Encoding.UTF8)
    Dim RetCode As StatusCodes = StatusCodes.SUCCESS
    Dim RetMsg As String = ""
    Dim BatchGUID As String = System.Guid.NewGuid.ToString
    Dim dtBatch As DataTable = Nothing
    Dim row As DataRow
    Dim TotalRecs, BatchCount, Remaining As Integer
    Dim ShowBatchID As Boolean = True

    Try
      If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
      If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

      ' validate the request
      If Not IsValidGUID(GUID, "GetLogixCustomerUpdates") Then
        RetCode = StatusCodes.INVALID_GUID
        RetMsg = "GUID " & GUID & " is not valid for the Customer Update web service."
      End If

      ' check if this connector is within the time range of being called.
      If RetCode <> StatusCodes.INVALID_GUID AndAlso IsConnectorLocked(GUID) Then
        RetCode = StatusCodes.TIME_LOCKED
        RetMsg = "CustomerData web service was called more than once within a 10 second time range."
      End If

      If RetCode = StatusCodes.SUCCESS Then
        Writer.Formatting = Formatting.Indented
        Writer.Indentation = 4

        Writer.WriteStartDocument()

        Select Case UpdateType
          Case UPDATE_TYPES.UNPROCESSED
            MyCommon.QueryStr = "dbo.pa_CustomerData_ChangedRecords"
            MyCommon.Open_LXSsp()
            MyCommon.LXSsp.Parameters.Add("@BatchGUID", SqlDbType.NVarChar, 36).Value = Left(BatchGUID, 36)
            MyCommon.LXSsp.Parameters.Add("@RecordsToProcess", SqlDbType.Int).Direction = ParameterDirection.Output
            dtBatch = MyCommon.LXSsp_select
            Integer.TryParse(MyCommon.LXSsp.Parameters("@RecordsToProcess").Value, TotalRecs)
            MyCommon.Close_LXSsp()
            BatchCount = dtBatch.Rows.Count
            Remaining = TotalRecs - BatchCount
          Case UPDATE_TYPES.ALREADY_PROCESSED
            MyCommon.QueryStr = "dbo.pa_CustomerData_BatchRecords"
            MyCommon.Open_LXSsp()
            MyCommon.LXSsp.Parameters.Add("@BatchGUID", SqlDbType.NVarChar, 36).Value = Left(Args(0).ToString, 36)
            dtBatch = MyCommon.LXSsp_select
            MyCommon.Close_LXSsp()
            BatchCount = dtBatch.Rows.Count
            BatchGUID = Args(0).ToString
          Case UPDATE_TYPES.DATE_RANGE
            MyCommon.QueryStr = "dbo.pa_CustomerData_RecordsByDate"
            MyCommon.Open_LXSsp()
            MyCommon.LXSsp.Parameters.Add("@StartDate", SqlDbType.DateTime).Value = Args(0)
            MyCommon.LXSsp.Parameters.Add("@EndDate", SqlDbType.DateTime).Value = Args(1)
            dtBatch = MyCommon.LXSsp_select
            MyCommon.Close_LXSsp()
            BatchCount = dtBatch.Rows.Count
            ShowBatchID = False
        End Select

        Writer.WriteStartElement("CustomerData")
        Writer.WriteStartElement("Response")
        Writer.WriteAttributeString("returnCode", "SUCCESS")
        Writer.WriteElementString("Message", RetMsg)
        ' write the batch ID at the top level if all the records in the response share the same batch ID
        If ShowBatchID Then
          Writer.WriteElementString("BatchID", BatchGUID)
        End If
        Writer.WriteElementString("BatchCount", BatchCount.ToString)
        Writer.WriteElementString("Remaining", Remaining.ToString)
        Writer.WriteEndElement() ' end Response

        If dtBatch IsNot Nothing Then
          For Each row In dtBatch.Rows
            WriteCustomerRecord(Writer, row, UpdateType)
          Next
        End If

        Writer.WriteEndElement() ' end CustomerData
        Writer.WriteEndDocument()
        Writer.Flush()

        ms.Seek(0, SeekOrigin.Begin)
        RetXmlDoc.Load(ms)
      Else
        MyCommon.Write_Log(LogFile, RetMsg, True)
        RetXmlDoc = New XmlDocument()
        RetXmlDoc.LoadXml(GetErrorXML(RetCode, RetMsg))
      End If

    Catch ex As Exception
      MyCommon.Write_Log(LogFile, ex.ToString, True)

      RetXmlDoc = New XmlDocument()
      RetXmlDoc.LoadXml(GetErrorXML(StatusCodes.APPLICATION_EXCEPTION, ex.ToString))
    Finally
      If Not Writer Is Nothing Then
        Writer.Close()
      End If
      If Not ms Is Nothing Then
        ms.Close()
        ms.Dispose()
        ms = Nothing
      End If

      UpdateLastCalled(GUID)

      If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
      If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
    End Try

    Return RetXmlDoc
  End Function

  Private Sub WriteCustomerRecord(ByRef Writer As XmlTextWriter, ByVal row As DataRow, ByVal UpdateType As UPDATE_TYPES)
    Dim CustomerPK As Long
    Dim TempDate As Date

    CustomerPK = MyCommon.NZ(row.Item("CustomerPK"), 0)
    If IsDBNull(row.Item("LastUpdate")) Then
      TempDate = New Date(1980, 1, 1)
    Else
      TempDate = row.Item("LastUpdate")
    End If

    If CustomerPK > 0 Then
      Writer.WriteStartElement("Customer")
      Writer.WriteAttributeString("id", CustomerPK)

      Writer.WriteStartElement("ChangeDetails")
      Writer.WriteElementString("AdminUserID", MyCommon.NZ(row.Item("AdminUserID"), 0))
      Writer.WriteElementString("ChangeDate", FormatDateXmlString(row, "LastUpdate"))
      Writer.WriteElementString("ChangeID", MyCommon.NZ(row.Item("EditPK"), 0))
      If UpdateType = UPDATE_TYPES.DATE_RANGE Then
        Writer.WriteElementString("BatchID", MyCommon.NZ(row.Item("BatchGUID"), ""))
      End If
      Writer.WriteEndElement() ' end change details

      ' write the customer data found in the customers table
      Writer.WriteStartElement("Data")
      Writer.WriteElementString("CustomerPK", CustomerPK)
      Writer.WriteElementString("FirstName", MyCommon.NZ(row.Item("FirstName"), ""))
      Writer.WriteElementString("LastName", MyCommon.NZ(row.Item("LastName"), ""))
      Writer.WriteElementString("Employee", MyCommon.NZ(row.Item("Employee"), "false").ToString.ToLower)
      Writer.WriteElementString("CurrYearSTD", MyCommon.NZ(row.Item("CurrYearSTD"), "0"))
      Writer.WriteElementString("LastYearSTD", MyCommon.NZ(row.Item("LastYearSTD"), "0"))
      Writer.WriteElementString("Password", MyCryptlib.SQL_StringDecrypt(MyCommon.NZ(row.Item("Password"), "")))
      Writer.WriteElementString("HHPK", MyCommon.NZ(row.Item("HHPK"), "0"))
      Writer.WriteElementString("EnrollmentDate", MyCommon.NZ(row.Item("EnrollmentDate"), ""))

      Writer.WriteElementString("LastComm", FormatDateXmlString(row, "LastComm"))
      Writer.WriteElementString("CommFrequency", MyCommon.NZ(row.Item("CommFrequency"), "0"))
      Writer.WriteElementString("CommEmail", MyCommon.NZ(row.Item("CommEmail"), "false").ToString.ToLower)
      Writer.WriteElementString("CommPrint", MyCommon.NZ(row.Item("CommPrint"), "false").ToString.ToLower)
      Writer.WriteElementString("BannerID", MyCommon.NZ(row.Item("BannerID"), "0"))
      Writer.WriteElementString("AltID", MyCommon.NZ(row.Item("AltID"), ""))
      Writer.WriteElementString("Verifier", MyCommon.NZ(row.Item("Verifier"), ""))
      Writer.WriteElementString("AltIDOptOut", MyCommon.NZ(row.Item("AltIDOptOut"), ""))
      Writer.WriteElementString("TestCard", MyCommon.NZ(row.Item("TestCard"), "false").ToString.ToLower)
      Writer.WriteElementString("MiddleName", MyCommon.NZ(row.Item("MiddleName"), ""))
      Writer.WriteElementString("EmployeeID", MyCommon.NZ(row.Item("EmployeeID"), ""))
      Writer.WriteElementString("CustomerStatusID", MyCommon.NZ(row.Item("CustomerStatusID"), "1"))
      Writer.WriteElementString("CustomerTypeID", MyCommon.NZ(row.Item("CustomerTypeID"), "0"))
      Writer.WriteEndElement() ' end Data

      ' write the extended customer data found in CustomerExt
      Writer.WriteStartElement("ExtendedData")
      Writer.WriteElementString("CustomerPK", CustomerPK)
      Writer.WriteElementString("Address", MyCommon.NZ(row.Item("Address"), ""))
      Writer.WriteElementString("City", MyCommon.NZ(row.Item("City"), ""))
      Writer.WriteElementString("State", MyCommon.NZ(row.Item("State"), ""))
      Writer.WriteElementString("Zip", MyCommon.NZ(row.Item("ZIP"), ""))
            Writer.WriteElementString("Phone", MyCryptlib.SQL_StringDecrypt(MyCommon.NZ(row.Item("Phone"), "")))
            Writer.WriteElementString("Email", MyCryptlib.SQL_StringDecrypt(MyCommon.NZ(row.Item("Email"), "")))
      Writer.WriteElementString("Country", MyCommon.NZ(row.Item("Country"), ""))
            Writer.WriteElementString("DOB", MyCryptlib.SQL_StringDecrypt(MyCommon.NZ(row.Item("DOB"), "")))
      Writer.WriteEndElement() ' end ExtendedData

      Writer.WriteEndElement() ' end customer
    End If

  End Sub

  Private Function GetErrorXML(ByVal Code As StatusCodes, ByVal Message As String) As String
    Dim ErrorXml As New StringBuilder("<?xml version=""1.0"" encoding=""utf-8""?>")

    ErrorXml.Append("<CustomerData>")
    ErrorXml.Append("<Response returnCode=""")
    Select Case Code
      Case StatusCodes.INVALID_GUID
        ErrorXml.Append("INVALID_GUID")
      Case StatusCodes.TIME_LOCKED
        ErrorXml.Append("TIME_LOCKED")
      Case Else
        ' treat everything else as an application exception
        ErrorXml.Append("APPLICATION_EXCEPTION")
    End Select
    ErrorXml.Append(""">")

    If Message <> "" Then
      ErrorXml.Append("  <Message>" & Message & "</Message>")
    End If

    ErrorXml.Append("  <BatchCount>0</BatchCount>")
    ErrorXml.Append("  <Remaining>0</Remaining>")
    ErrorXml.Append("</Response>")
    ErrorXml.Append("</CustomerData>")

    Return ErrorXml.ToString
  End Function

  Private Function IsValidGUID(ByVal GUID As String, ByVal MethodName As String) As Boolean
    Dim IsValid As Boolean = False
    Dim ConnInc As New Copient.ConnectorInc
    Dim MsgBuf As New StringBuilder()

    Try
      If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
      IsValid = ConnInc.IsValidConnectorGUID(MyCommon, 43, GUID)
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

  Private Function IsValidDate(ByVal inputDate As String) As Boolean
    Dim dt As DateTime
    Dim IsDate As Boolean = False
    Try
      If DateTime.TryParse(inputDate, dt) Then
        If dt >= SqlTypes.SqlDateTime.MinValue And dt <= SqlTypes.SqlDateTime.MaxValue Then
          IsDate = True
        End If
      End If
    Catch
    End Try
    Return IsDate
  End Function

  Private Function IsConnectorLocked(ByVal GUID As String) As Boolean
    Dim Locked As Boolean = False

    If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
    MyCommon.QueryStr = "dbo.pa_ConnectorGUIDs_IsTimeLocked"
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add("@ConnectorID", SqlDbType.Int).Value = 43
    MyCommon.LRTsp.Parameters.Add("@ConnectorGUID", SqlDbType.NVarChar, 36).Value = GUID
    MyCommon.LRTsp.Parameters.Add("@LockSeconds", SqlDbType.Int).Value = 10 ' set to every 10 seconds
    MyCommon.LRTsp.Parameters.Add("@Locked", SqlDbType.Bit).Direction = ParameterDirection.Output
    MyCommon.LRTsp.ExecuteNonQuery()
    Locked = MyCommon.LRTsp.Parameters("@Locked").Value

    Return Locked
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

  Private Sub UpdateLastCalled(ByVal GUID As String)
    Try
      If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
      MyCommon.QueryStr = "Update ConnectorGUIDs with (RowLock) set LastCalled=getdate() " & _
                          "where ConnectorID=43 and GUID=@GUID"
      MyCommon.DBParameters.Add("@GUID", SqlDbType.NVarChar).Value = GUID
      MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)

    Catch ex As Exception
      ' ignore exception
      MyCommon.Write_Log(LogFile, ex.ToString, True)
    End Try

  End Sub

  Private Function FormatDateXmlString(ByVal row As DataRow, ByVal ColumnName As String) As String
    Dim TempDate As Date

    If IsDBNull(row.Item(ColumnName)) Then
      TempDate = New Date(1980, 1, 1)
    Else
      TempDate = row.Item(ColumnName)
    End If

    Return TempDate.ToString("yyyy-MM-ddTHH:mm:ss")
  End Function

End Class