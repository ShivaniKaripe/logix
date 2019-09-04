<%@ WebService Language="VB" Class="Service" %>
' version:7.3.1.138972.Official Build (SUSDAY10202) 

Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.Odbc
Imports System.IO
Imports System.Xml
Imports System.Text
Imports System.Security.Cryptography
Imports Copient.CommonInc
Imports Copient.ImportXml


<WebService(Namespace:="http://www.copienttech.com/CmExtOfferConnector/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class Service
  Inherits System.Web.Services.WebService

  Private Const sVersion As String = "5.7b01"
  Private Const sLogFileName As String = "CmExtOfferConnector"
  Private Const sOkStatus1 As String = "<Response Status=""Ok"" Version="""
  Private Const sOkStatus2 As String = """ Method="""
  Private Const sOkStatus3 As String = """ />" & vbCrLf
  Private Const sAppName As String = "CmExtOfferConnector"
  Private Const sCmEngine As String = "CM"
  Private Const lUserId As Long = 1
  Private Const iLanguageId As Integer = 1
  Private Const bDebugLogOn As Boolean = True

  Private sInputForLog As String = ""
  Private sCurrentMethod As String = ""
  Private sInstallationName As String = ""
  Private LogixCommon As Copient.CommonInc
  Private eDefaultErrorType As ErrorType = ErrorType.General
  Private DebugStartTimes As New ArrayList()


  Private Enum DebugState
    BeginTime = -1
    CurrentTime = 0
    EndTime = 1
  End Enum

  Private Enum ErrorType
    SqlServer = 0
    General = 1
  End Enum

  Private Enum MessageType
    Info = 0
    Warning = 1
    AppError = 2
    SysError = 3
    Debug = 4
  End Enum

  <WebMethod()> _
  Public Function AboutThisService() As String
    Dim sResponse As String

    sResponse = "Logix CmExtOfferConnector (Version " & sVersion & ")"

    Return sResponse
  End Function

  <WebMethod()> _
  Public Function OfferMaintenance(ByVal EngineId As String, ByVal OfferXml As String) As String
    Dim sXml As String

    sXml = ProcessExtOffer(EngineId, OfferXml)

    Return sXml
  End Function

  <WebMethod()> _
  Public Function CustomerMaintenance(ByVal EngineId As String, ByVal CustomerXml As String) As String
    Dim sXml As String

    sXml = ProcessExtOfferCustomer(EngineId, CustomerXml)

    Return sXml
  End Function

  <WebMethod()> _
  Public Function OfferStatus(ByVal StatusType As String, ByVal ParameterXml As String) As String
    Dim sXml As String

    sXml = GetExtOfferStatus(StatusType, ParameterXml)

    Return sXml
  End Function

  <WebMethod()> _
  Public Function StoredValueCoupons(ByVal EngineId As String, ByVal CouponXml As String) As String
    Dim sXml As String

    sXml = ProcessStoredValueCoupons(EngineId, CouponXml)

    Return sXml
  End Function

  Private Function ProcessExtOffer(ByVal sEngineId As String, ByVal sOfferXml As String) As String
    Dim bStatus As Boolean
    Dim sErrorMsg As String
    Dim sOfferId As String
    Dim sExtOfferId As String
    Dim sResponse As String
    Dim Import As Copient.ImportXml
    Dim sw As StringWriter = Nothing
    Dim Writer As XmlWriter = Nothing
    Dim SettingsWrite As XmlWriterSettings

    sCurrentMethod = "OfferMaintenance"
    sInputForLog = "*(Type=Input) (Method=OfferMaintenance)] - (EngineId='" & sEngineId & "' - " & sOfferXml & ")"
    LogixCommon = New Copient.CommonInc
    LogixCommon.AppName = sAppName
    Import = New Copient.ImportXml

    eDefaultErrorType = ErrorType.SqlServer
    LogixCommon.Open_LogixRT()
    LogixCommon.Load_System_Info()
    sInstallationName = LogixCommon.InstallationName
    eDefaultErrorType = ErrorType.General

    WriteDebug("OfferMaintenance", DebugState.BeginTime)

    Try
      If sEngineId = sCmEngine Then
        bStatus = Import.ProcessExternalOffer("", sOfferXml, 1, 1)
        If bStatus Then
          sOfferId = Import.GetOfferId
          sExtOfferId = Import.GetOriginalOfferId

          sw = New StringWriter()
          SettingsWrite = New XmlWriterSettings()
          SettingsWrite.Indent = True
          Writer = XmlWriter.Create(sw, SettingsWrite)
          Writer.WriteStartDocument()
          Writer.WriteStartElement("Response")
          Writer.WriteAttributeString("Status", "Ok")
          Writer.WriteAttributeString("Version", sVersion)
          Writer.WriteAttributeString("Method", sCurrentMethod)
          Writer.WriteElementString("LogixOfferID", sOfferId)
          Writer.WriteEndElement() ' Response
          Writer.WriteEndDocument()
          Writer.Flush()
          sResponse = sw.ToString
        Else
          sErrorMsg = Import.GetStatusMsg
          Throw New ApplicationException(sErrorMsg)
        End If
      Else
        Throw New ApplicationException("Invalid EngineId: " & sEngineId)
      End If
    Catch exApp As ApplicationException
      sResponse = BuildErrorXml(exApp.Message, ErrorType.General, False)
    Catch ex As Exception
      sResponse = BuildErrorXml(ex.Message, ErrorType.SqlServer, True)
    Finally
      If Not Writer Is Nothing Then
        Writer.Close()
      End If
      If Not sw Is Nothing Then
        sw.Close()
        sw.Dispose()
      End If
      WriteDebug("OfferMaintenance", DebugState.EndTime)
      If LogixCommon.LWHadoConn.State <> ConnectionState.Closed Then
        LogixCommon.Close_LogixWH()
      End If
      If LogixCommon.LXSadoConn.State <> ConnectionState.Closed Then
        LogixCommon.Close_LogixXS()
      End If
      If LogixCommon.LRTadoConn.State <> ConnectionState.Closed Then
        LogixCommon.Close_LogixRT()
      End If
    End Try

    Return sResponse
  End Function

  Private Function ProcessExtOfferCustomer(ByVal sEngineId As String, ByVal sCustomerXml As String) As String
    Dim bStatus As Boolean
    Dim sErrorMsg As String
    Dim sResponse As String
    Dim Import As Copient.ImportXml
    Dim sw As StringWriter = Nothing
    Dim Writer As XmlWriter = Nothing
    Dim SettingsWrite As XmlWriterSettings

    sCurrentMethod = "CustomerMaintenance"
    sInputForLog = "*(Type=Input) (Method=CustomerMaintenance)] - (EngineId='" & sEngineId & "' - " & sCustomerXml & ")"
    LogixCommon = New Copient.CommonInc
    LogixCommon.AppName = sAppName
    Import = New Copient.ImportXml

    eDefaultErrorType = ErrorType.SqlServer
    LogixCommon.Open_LogixRT()
    LogixCommon.Load_System_Info()
    sInstallationName = LogixCommon.InstallationName
    eDefaultErrorType = ErrorType.General

    WriteDebug("CustomerMaintenance", DebugState.BeginTime)

    Try
      If sEngineId = sCmEngine Then
        bStatus = Import.ProcessExternalCustomerBinding(sCustomerXml, 1, 1)
        If bStatus Then
          sw = New StringWriter()
          SettingsWrite = New XmlWriterSettings()
          SettingsWrite.Indent = True
          Writer = XmlWriter.Create(sw, SettingsWrite)
          Writer.WriteStartDocument()
          Writer.WriteStartElement("Response")
          Writer.WriteAttributeString("Status", "Ok")
          Writer.WriteAttributeString("Version", sVersion)
          Writer.WriteAttributeString("Method", sCurrentMethod)
          Writer.WriteEndElement() ' Response
          Writer.WriteEndDocument()
          Writer.Flush()
          sResponse = sw.ToString
        Else
          sErrorMsg = Import.GetStatusMsg
          Throw New ApplicationException(sErrorMsg)
        End If
      Else
        Throw New ApplicationException("Invalid EngineId: " & sEngineId)
      End If
    Catch exApp As ApplicationException
      sResponse = BuildErrorXml(exApp.Message, ErrorType.General, False)
    Catch ex As Exception
      sResponse = BuildErrorXml(ex.Message, ErrorType.SqlServer, True)
    Finally
      If Not Writer Is Nothing Then
        Writer.Close()
      End If
      If Not sw Is Nothing Then
        sw.Close()
        sw.Dispose()
      End If
      WriteDebug("CustomerMaintenance", DebugState.EndTime)
      If LogixCommon.LWHadoConn.State <> ConnectionState.Closed Then
        LogixCommon.Close_LogixWH()
      End If
      If LogixCommon.LXSadoConn.State <> ConnectionState.Closed Then
        LogixCommon.Close_LogixXS()
      End If
      If LogixCommon.LRTadoConn.State <> ConnectionState.Closed Then
        LogixCommon.Close_LogixRT()
      End If
    End Try

    Return sResponse
  End Function

  Private Function GetExtOfferStatus(ByVal sType As String, ByVal sParameterXml As String) As String
    Dim sResponse As String
    Dim sExtOfferId As String
    Dim iValid As Integer
    Dim iWatch As Integer
    Dim iWarn As Integer
    Dim bComponents As Boolean
    Dim sr As StringReader = Nothing
    Dim SettingsRead As XmlReaderSettings
    Dim xr As XmlReader = Nothing
    Dim xrSubTree As XmlReader = Nothing
    Dim sw As StringWriter = Nothing
    Dim Writer As XmlWriter = Nothing
    Dim SettingsWrite As XmlWriterSettings

    sCurrentMethod = "OfferStatus"
    sInputForLog = "*(Type=Input) (Method=OfferStatus)] - (EngineId='" & sType & "' - " & sParameterXml & ")"
    LogixCommon = New Copient.CommonInc
    LogixCommon.AppName = sAppName

    eDefaultErrorType = ErrorType.SqlServer
    LogixCommon.Open_LogixRT()
    LogixCommon.Load_System_Info()
    sInstallationName = LogixCommon.InstallationName
    eDefaultErrorType = ErrorType.General

    WriteDebug("OfferStatus", DebugState.BeginTime)

    Try
      If sType.ToLower = "offerhealth" Then
        sr = New StringReader(sParameterXml)

        SettingsRead = New XmlReaderSettings()
        xr = XmlReader.Create(sr, SettingsRead)
        xr.ReadToFollowing("OfferStatus")
        If xr.EOF Then
          Throw New ApplicationException("No 'OfferStatus' root element")
        End If
        xr.ReadToFollowing("OfferID")
        If xr.EOF Then
          Throw New ApplicationException("No 'OfferID' element")
        End If
        sExtOfferId = xr.ReadElementContentAsString

        ValidateCmOffer(sExtOfferId, iValid, iWatch, iWarn, bComponents)

        sw = New StringWriter()
        SettingsWrite = New XmlWriterSettings()
        SettingsWrite.Indent = True
        Writer = XmlWriter.Create(sw, SettingsWrite)
        Writer.WriteStartDocument()
        Writer.WriteStartElement("Response")
        Writer.WriteAttributeString("Status", "Ok")
        Writer.WriteAttributeString("Version", sVersion)
        Writer.WriteAttributeString("Method", sCurrentMethod)
        Writer.WriteElementString("Valid", iValid.ToString)
        Writer.WriteElementString("Watch", iWatch.ToString)
        Writer.WriteElementString("Warning", iWarn.ToString)
        Writer.WriteElementString("ProductGroupsOk", bComponents.ToString.ToLower)
        Writer.WriteEndElement() ' Response
        Writer.WriteEndDocument()
        Writer.Flush()
        sResponse = sw.ToString
      Else
        sResponse = BuildErrorXml("Invalid Status Type: " & sType, eDefaultErrorType, False)
      End If
    Catch exApp As ApplicationException
      sResponse = BuildErrorXml(exApp.Message, ErrorType.General, False)
    Catch ex As Exception
      sResponse = BuildErrorXml(ex.Message, ErrorType.SqlServer, True)
    Finally
      If Not xr Is Nothing Then
        xr.Close()
      End If
      If Not sr Is Nothing Then
        sr.Close()
        sr.Dispose()
      End If
      If Not Writer Is Nothing Then
        Writer.Close()
      End If
      If Not sw Is Nothing Then
        sw.Close()
        sw.Dispose()
      End If
      WriteDebug("OfferStatus", DebugState.EndTime)
      If LogixCommon.LWHadoConn.State <> ConnectionState.Closed Then
        LogixCommon.Close_LogixWH()
      End If
      If LogixCommon.LXSadoConn.State <> ConnectionState.Closed Then
        LogixCommon.Close_LogixXS()
      End If
      If LogixCommon.LRTadoConn.State <> ConnectionState.Closed Then
        LogixCommon.Close_LogixRT()
      End If
    End Try
    Return sResponse
  End Function

  Private Function ProcessStoredValueCoupons(ByVal sEngineId As String, ByVal sCouponXml As String) As String
    Dim bStatus As Boolean
    Dim sErrorMsg As String
    Dim sResponse As String
    Dim Import As Copient.ImportXml
    Dim sw As StringWriter = Nothing
    Dim Writer As XmlWriter = Nothing
    Dim SettingsWrite As XmlWriterSettings
    Dim sInputTemp As String
    
    If (sCouponXml.Length > 200) Then
      sInputTemp = sCouponXml.Substring(0, 200) & "*** TRUNCATED ***"
    Else
      sInputTemp = sCouponXml
    End If

    sCurrentMethod = "StoredValueCoupons"
    sInputForLog = "*(Type=Input) (Method=StoredValueCoupons)] - (EngineId='" & sEngineId & "' - " & sInputTemp & ")"
    LogixCommon = New Copient.CommonInc
    LogixCommon.AppName = sAppName
    Import = New Copient.ImportXml

    eDefaultErrorType = ErrorType.SqlServer
    LogixCommon.Open_LogixRT()
    LogixCommon.Load_System_Info()
    sInstallationName = LogixCommon.InstallationName
    eDefaultErrorType = ErrorType.General

    WriteDebug("StoredValueCoupons", DebugState.BeginTime)

    Try
      If sEngineId = sCmEngine Then
        bStatus = Import.ProcessCoupons("", sCouponXml, 1, 1, True)
        If bStatus Then
          sw = New StringWriter()
          SettingsWrite = New XmlWriterSettings()
          SettingsWrite.Indent = True
          Writer = XmlWriter.Create(sw, SettingsWrite)
          Writer.WriteStartDocument()
          Writer.WriteStartElement("Response")
          Writer.WriteAttributeString("Status", "Ok")
          Writer.WriteAttributeString("Version", sVersion)
          Writer.WriteAttributeString("Method", sCurrentMethod)
          Writer.WriteEndElement() ' Response
          Writer.WriteEndDocument()
          Writer.Flush()
          sResponse = sw.ToString
        Else
          sErrorMsg = Import.GetStatusMsg
          Throw New ApplicationException(sErrorMsg)
        End If
      Else
        Throw New ApplicationException("Invalid EngineId: " & sEngineId)
      End If
    Catch exApp As ApplicationException
      sResponse = BuildErrorXml(exApp.Message, ErrorType.General, False)
    Catch ex As Exception
      sResponse = BuildErrorXml(ex.Message, ErrorType.SqlServer, True)
    Finally
      If Not Writer Is Nothing Then
        Writer.Close()
      End If
      If Not sw Is Nothing Then
        sw.Close()
        sw.Dispose()
      End If
      WriteDebug("StoredValueCoupons", DebugState.EndTime)
      If LogixCommon.LWHadoConn.State <> ConnectionState.Closed Then
        LogixCommon.Close_LogixWH()
      End If
      If LogixCommon.LXSadoConn.State <> ConnectionState.Closed Then
        LogixCommon.Close_LogixXS()
      End If
      If LogixCommon.LRTadoConn.State <> ConnectionState.Closed Then
        LogixCommon.Close_LogixRT()
      End If
    End Try

    Return sResponse
  End Function

  Private Sub ValidateCmOffer(ByVal sExtOfferID As String, ByRef iValid As Integer, ByRef iWatch As Integer, ByRef iWarn As Integer, ByRef bComponents As Boolean)
    Dim dtValid, dtComponents As DataTable
    Dim rowOK(), rowWaiting(), rowWatches(), rowWarnings() As DataRow
    Dim rowComp As DataRow
    Dim objTemp As Object
    Dim GraceHours As Integer
    Dim GraceHoursWarn As Integer
    Dim ComponentsValid As Boolean = True
    Dim dt As DataTable
    Dim i64OfferId As Int64 = 0

    LogixCommon.QueryStr = "select O.OfferID from Offers O with (NoLock) " & _
                    "where O.Deleted=0 and O.InboundCRMEngineId=3 and O.ExtOfferID='" & sExtOfferID & "'"
    dt = LogixCommon.LRT_Select
    If (dt.Rows.Count > 0) Then
      i64OfferId = LogixCommon.NZ(dt.Rows(0).Item("OfferID"), 0)
    End If
    If i64OfferId = 0 Then
      Throw New ApplicationException("Avenu Offer: '" & sExtOfferID & "' not found!")
    End If


    objTemp = LogixCommon.Fetch_CM_SystemOption(10)
    If Not (Integer.TryParse(objTemp.ToString, GraceHours)) Then
      GraceHours = 4
    End If

    objTemp = LogixCommon.Fetch_CM_SystemOption(11)
    If Not (Integer.TryParse(objTemp.ToString, GraceHoursWarn)) Then
      GraceHoursWarn = 24
    End If

    LogixCommon.QueryStr = "dbo.pa_CM_ValidationReport_Offer"
    LogixCommon.Open_LRTsp()
    LogixCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.Int).Value = i64OfferId
    LogixCommon.LRTsp.Parameters.Add("@GraceHours", SqlDbType.Int).Value = GraceHours
    LogixCommon.LRTsp.Parameters.Add("@GraceHoursWarn", SqlDbType.Int).Value = GraceHoursWarn

    dtValid = LogixCommon.LRTsp_select()

    rowOK = dtValid.Select("Status=0", "LocationName")
    rowWaiting = dtValid.Select("Status=1", "LocationName")
    rowWatches = dtValid.Select("Status=2", "LocationName")
    rowWarnings = dtValid.Select("Status=3", "LocationName")
    LogixCommon.Close_LRTsp()

    LogixCommon.QueryStr = "dbo.pa_CM_ValidationReport_OfferComponents"
    LogixCommon.Open_LRTsp()
    LogixCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.Int).Value = i64OfferId
    dtComponents = LogixCommon.LRTsp_select

    For Each rowComp In dtComponents.Rows
      ComponentsValid = IsValidCmComponent(rowComp)
      If (Not ComponentsValid) Then Exit For
    Next

    iValid = rowOK.Length
    iWatch = rowWatches.Length + rowWaiting.Length
    iWarn = rowWarnings.Length
    bComponents = ComponentsValid

    ' Note: The code below may be over kill since it adds overhead to performance time, but
    '       thought it would be nice to keep the "Offer Health" page in sync!
    ' Update the Offer Validation Summary table with the most current validation information
    LogixCommon.QueryStr = "dbo.pa_UpdateValidationSummary"
    LogixCommon.Open_LRTsp()
    LogixCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = i64OfferId
    LogixCommon.LRTsp.Parameters.Add("@ValidLocations", SqlDbType.Int).Value = iValid
    LogixCommon.LRTsp.Parameters.Add("@WatchLocations", SqlDbType.Int).Value = iWatch
    LogixCommon.LRTsp.Parameters.Add("@WarningLocations", SqlDbType.Int).Value = iWarn
    LogixCommon.LRTsp.Parameters.Add("@ComponentsValid", SqlDbType.Bit).Value = IIf(bComponents, 1, 0)
    LogixCommon.LRTsp.ExecuteNonQuery()
    LogixCommon.Close_LRTsp()

  End Sub

  Private Function IsValidCmComponent(ByVal rowComp As DataRow) As Boolean
    Dim RecordType As String = ""
    Dim ID As Integer
    Dim StoredProcName As String = ""
    Dim IDParmName As String = ""
    Dim dtValid As DataTable
    Dim rowWarnings() As DataRow
    Dim objTemp As Object
    Dim GraceHours As Integer
    Dim GraceHoursWarn As Integer
    Dim bValid As Boolean = True
    Dim RequiresValidation As Boolean = False

    objTemp = LogixCommon.Fetch_CM_SystemOption(10)
    If Not (Integer.TryParse(objTemp.ToString, GraceHours)) Then
      GraceHours = 4
    End If

    objTemp = LogixCommon.Fetch_CM_SystemOption(11)
    If Not (Integer.TryParse(objTemp.ToString, GraceHoursWarn)) Then
      GraceHoursWarn = 24
    End If

    RecordType = LogixCommon.NZ(rowComp.Item("RecordType"), "")
    ID = LogixCommon.NZ(rowComp.Item("ID"), -1)

    Select Case RecordType
      Case "term.customergroup"
        StoredProcName = "dbo.pa_CM_ValidationReport_CustGroup"
        IDParmName = "@CustomerGroupID"
        RequiresValidation = IIf(ID = 1 OrElse ID = 2, False, True)
      Case "term.productgroup"
        StoredProcName = "dbo.pa_CM_ValidationReport_ProdGroup"
        IDParmName = "@ProductGroupID"
        RequiresValidation = IIf(ID = 1, False, True)
    End Select

    If (RequiresValidation) Then
      LogixCommon.QueryStr = StoredProcName
      LogixCommon.Open_LRTsp()
      LogixCommon.LRTsp.Parameters.Add(IDParmName, SqlDbType.Int).Value = ID
      LogixCommon.LRTsp.Parameters.Add("@GraceHours", SqlDbType.Int).Value = GraceHours
      LogixCommon.LRTsp.Parameters.Add("@GraceHoursWarn", SqlDbType.Int).Value = GraceHoursWarn

      dtValid = LogixCommon.LRTsp_select()
      rowWarnings = dtValid.Select("Status=3", "LocationName")
      If rowWarnings.Length > 0 Then
        bValid = False
      Else
        bValid = True
      End If
      LogixCommon.Close_LRTsp()
    Else
      bValid = True
    End If

    Return bValid
  End Function

  Private Function BuildErrorXml(ByVal sText As String, ByVal eErrType As ErrorType, ByVal bSystemError As Boolean) As String
    Dim sXml As String
    Dim iErrType As Integer
    Dim sLogText As String
    Dim eMsgType As MessageType
    Dim sw As StringWriter = Nothing
    Dim Writer As XmlWriter = Nothing
    Dim SettingsWrite As XmlWriterSettings

    iErrType = eErrType
    If bSystemError Then
      eMsgType = MessageType.SysError
      Try
        LogixCommon.Error_Processor(, sCurrentMethod, sAppName, sInstallationName)
      Catch
      End Try
    Else
      eMsgType = MessageType.AppError
    End If
    sLogText = WriteLog(sText, eMsgType)

    sw = New StringWriter()
    SettingsWrite = New XmlWriterSettings()
    SettingsWrite.Indent = True
    Writer = XmlWriter.Create(sw, SettingsWrite)
    Writer.WriteStartDocument()
    Writer.WriteStartElement("Response")
    Writer.WriteAttributeString("Status", "Error")
    Writer.WriteAttributeString("Version", sVersion)
    Writer.WriteAttributeString("Method", sCurrentMethod)
    Writer.WriteStartElement("Error")
    Writer.WriteElementString("Type", iErrType.ToString)
    Writer.WriteElementString("Text", sLogText)
    Writer.WriteEndElement() ' Error
    Writer.WriteEndElement() ' Response
    Writer.WriteEndDocument()
    Writer.Flush()
    sXml = sw.ToString

    If Not Writer Is Nothing Then
      Writer.Close()
    End If
    If Not sw Is Nothing Then
      sw.Close()
      sw.Dispose()
    End If

    Return sXml
  End Function

  Private Sub WriteDebug(ByVal sText As String, ByVal mode As DebugState)
    If bDebugLogOn Then
      Dim TotalSeconds As Double
      Dim sIndent As String
      Select Case mode
        Case DebugState.BeginTime
          ' first call
          DebugStartTimes.Add(Now)
          If DebugStartTimes.Count = 1 Then
            WriteLog("------------------------------------------------------------------------", MessageType.Debug)
          Else
            sIndent = New String(" "c, 2 * (DebugStartTimes.Count - 1))
            sText = sIndent & sText
          End If
          sText = sText & " - Begin"
        Case DebugState.EndTime
          ' last call
          If DebugStartTimes.Count > 0 Then
            TotalSeconds = Now.Subtract(DebugStartTimes(DebugStartTimes.Count - 1)).TotalSeconds
            sIndent = New String(" "c, 2 * (DebugStartTimes.Count - 1))
            sText = sIndent & sText & " - End elasped time: " & Format(TotalSeconds, "00.000") & "(sec)"
            DebugStartTimes.RemoveAt(DebugStartTimes.Count - 1)
          End If
        Case Else
          ' interim call
          If DebugStartTimes.Count > 0 Then
            TotalSeconds = Now.Subtract(DebugStartTimes(DebugStartTimes.Count - 1)).TotalSeconds
            sIndent = New String(" "c, 2 * (DebugStartTimes.Count - 1))
            sText = sIndent & sText & " - Current elasped time: " & Format(TotalSeconds, "0.000") & "(sec)"
          End If
      End Select
      WriteLog(sText, MessageType.Debug)
    End If
  End Sub

  Private Function WriteLog(ByVal sText As String, ByVal eType As MessageType) As String
    Dim sFileName As String
    Dim sLogText As String = ""

    Try

      sFileName = sLogFileName & Format(Date.Now, "yyyyMMdd") & ".txt"

      If eType = MessageType.Debug Then
        sLogText = "[" & Format(Date.Now, "MM/dd/yyyy HH:mm:ss.fffzzz") & " (Type=" & eType.ToString & ")] " & sText
      Else
        sText = sText.Replace(ControlChars.CrLf, " ")
        If sInputForLog.Length > 0 Then
          sLogText = "[" & Format(Date.Now, "MM/dd/yyyy HH:mm:ss.fffzzz") & sInputForLog & ControlChars.CrLf
          sInputForLog = ""
        End If
        sLogText = sLogText & "[" & Format(Date.Now, "MM/dd/yyyy HH:mm:ss.fffzzz") & " (Type=" & eType.ToString & ") (Method=" & sCurrentMethod & ")] " & sText
      End If

      LogixCommon.Write_Log(sFileName, sLogText)
    Catch ex As Exception
      Try
        LogixCommon.Error_Processor(, "WriteLog", sAppName, sInstallationName)
      Catch
      End Try
      sText += " (WriteLog Error: " & ex.Message & ")"
    End Try

    Return sText
  End Function

End Class