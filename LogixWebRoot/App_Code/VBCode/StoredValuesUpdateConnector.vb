Imports Microsoft.VisualBasic
Imports System.Web
Imports System.Web.Services
Imports System.Xml
Imports System.Data
Imports System.IO
<WebService(Namespace:="http://www.copienttech.com/StoredValuesUpdate/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class StoredValuesUpdateConnector
    Inherits System.Web.Services.WebService

    Private Common As New Copient.CommonInc
    Private ConnectorInc As New Copient.ConnectorInc
    Private SVLogFile As String = "StoredValuesUpdateWSLog." & Date.Now.ToString("yyyyMMdd") & ".txt"

    Private Const CONNECTOR_ID As Integer = 55

    Public Enum StatusCodes As Integer
        SUCCESS = 1
        INVALID_GUID = -1
        INVALID_XML_DOCUMENT = -2
        APPLICATION_EXCEPTION = -9999
    End Enum

    <WebMethod()> _
    Public Function AdjustStoredValues(ByVal GUID As String, ByVal UpdateStr As String) As XmlDocument

        Dim MethodName As String = "AdjustStoredValues"
        Dim ResponseXML As New XmlDocument
        Dim sw As StringWriter = Nothing
        Dim Writer As XmlWriter = Nothing
        Dim UpdateXML As New XmlDocument
        Dim StatusDesc As String = ""
        Dim ValidXMLDoc As Boolean
        Dim xsdName As String = "StoredValuesUpdate.xsd"
        Dim xsdError As String = ""
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim SVAdj As Copient.StoredValue

        On Error GoTo ErrorHandler
        UpdateStr = UpdateStr.Replace(vbLf, Environment.NewLine)
        Init_ResponseXML(sw, Writer)
        Open_Databases()

        If Not (IsValidGUID(GUID, MethodName, Common)) Then
            Generate_Status_XML(Writer, MethodName, StatusCodes.INVALID_GUID, "GUID '" & GUID & "' is not valid for the Stored Values Update web service.", False)
        Else
            ConnectorInc = New Copient.ConnectorInc()
            If ConnectorInc.ConvertStringToXML(UpdateStr, UpdateXML) Then
                ValidXMLDoc = ConnectorInc.IsValidXmlDocument(Common, xsdName, UpdateXML, xsdError)
                If ValidXMLDoc Then
                    SVAdj = New Copient.StoredValue(Common)
                    'process the XML that was passed in
                    StatusDesc = SVAdj.Process_StoredValuesUpdateXML(UpdateStr, RetCode)
                    Generate_Status_XML(Writer, MethodName, RetCode, StatusDesc, (RetCode = StatusCodes.SUCCESS))
                    Copient.Logger.Write_Log(SVLogFile, StatusDesc, True)
                Else
                    Generate_Status_XML(Writer, MethodName, StatusCodes.INVALID_XML_DOCUMENT, xsdError, False)
                End If
            Else
                Generate_Status_XML(Writer, MethodName, StatusCodes.INVALID_XML_DOCUMENT, "Invalid XML Document", False)
            End If

        End If

        GoTo Finish

ErrorHandler:
        Init_ResponseXML(sw, Writer)
        Common.Error_Processor()
        Generate_Status_XML(Writer, MethodName, StatusCodes.APPLICATION_EXCEPTION, "An error occurred while processing - please see the error log!", False)
        Copient.Logger.Write_Log(SVLogFile, "An error occurred while processing - please see the log!", True)

Finish:
        Close_ResponseXML(Writer)
        Writer.Flush()
        Writer.Close()
        ResponseXML.LoadXml(sw.ToString)
        Close_Databases()

        Return ResponseXML
    End Function



    Private Sub Open_Databases()
        If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()
        If Common.LXSadoConn.State = ConnectionState.Closed Then Common.Open_LogixXS()
    End Sub



    Private Sub Close_Databases()
        If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
        If Not (Common.LXSadoConn.State = ConnectionState.Closed) Then Common.Close_LogixXS()
    End Sub



    Private Sub Init_ResponseXML(ByRef sw As StringWriter, ByRef Writer As XmlWriter)

        Init_XMLWriter(sw, Writer)
        Writer.WriteStartDocument()
        Writer.WriteStartElement("StoredValuesAdjust")
    End Sub



    Private Sub Init_XMLWriter(ByRef sw As StringWriter, ByRef Writer As XmlWriter)

        Dim Settings As XmlWriterSettings

        sw = New StringWriter()
        Settings = New XmlWriterSettings()
        Settings.Indent = True
        Settings.Encoding = System.Text.UTF8Encoding.UTF8
        Settings.ConformanceLevel = ConformanceLevel.Auto

        Writer = XmlWriter.Create(sw, Settings)

    End Sub



    Private Sub Close_ResponseXML(ByRef Writer As XmlWriter)
        Writer.WriteEndElement()
        Writer.WriteEndDocument()
    End Sub



    Private Sub Generate_Status_XML(ByRef Writer As XmlWriter, ByVal MethodName As String, ByVal StatusCode As StatusCodes, ByVal StatusDescription As String, ByVal IsSuccess As Boolean)
        Writer.WriteStartElement("Status")
        Writer.WriteAttributeString("operation", MethodName)
        Writer.WriteAttributeString("success", IsSuccess.ToString.ToLower)
        Writer.WriteAttributeString("responseCode", StatusCode.ToString)
        Writer.WriteAttributeString("message", StatusDescription)
        Writer.WriteEndElement() 'Status
    End Sub



    Private Function IsValidGUID(ByVal GUID As String, ByVal MethodName As String, ByRef Common As Copient.CommonInc) As Boolean
        Dim IsValid As Boolean = False
        Dim ConnInc As New Copient.ConnectorInc
        Dim MsgBuf As New StringBuilder()

        Try
            IsValid = ConnInc.IsValidConnectorGUID(Common, CONNECTOR_ID, GUID)
        Catch ex As Exception
            IsValid = False
        End Try

        ' Log the call
        Try
            MsgBuf.Append(IIf(IsValid, "Validated call to ", "Invalid call to "))
            MsgBuf.Append(MethodName)
            MsgBuf.Append(" from GUID: ")
            MsgBuf.Append(GUID)
            MsgBuf.Append(" and IP: " & HttpContext.Current.Request.UserHostAddress)
            Copient.Logger.Write_Log(SVLogFile, MsgBuf.ToString, True)
        Catch ex As Exception
            ' ignore
        End Try

        Return IsValid
    End Function



End Class
