<%@ WebService Language="VB" Class="Service" %>

Imports System
Imports System.Web
Imports System.Web.Services
Imports System.Xml
Imports System.Data
Imports System.IO

<WebService(Namespace:="http://www.copienttech.com/CustomerUpdate/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class Service
  Inherits System.Web.Services.WebService
  ' version:7.3.1.138972.Official Build (SUSDAY10202)
  Private LogFile As String = "CustomerUpdateWSLog." & Date.Now.ToString("yyyyMMdd") & ".txt"
  
  Public Enum StatusCodes As Integer
    SUCCESS = 0
    INVALID_GUID = 1
    INVALID_CUSTOMERID = 2
    INVALID_CUSTOMERTYPEID = 3
    INVALID_CUSTOMERGROUPID = 4
    INVALID_MODE = 5
    NOTFOUND_CUSTOMER = 6
    NOTFOUND_HOUSEHOLD = 7
    NOTFOUND_CAM = 8
    INVALID_XML_DOCUMENT = 9
    MALFORMED_XML_FOR_OPERATION = 10
    OPERATION_TAG_LIMIT_EXCEEDED = 11
    INVALID_UNIQUE_CONSTRAINTS=12
    APPLICATION_EXCEPTION = 9999
  End Enum
  
  Public Enum CustomerTypes As Integer
    CARDHOLDER = 0
    HOUSEHOLD = 1
    CAM = 2
    ALTERNATEID = 3
  End Enum

  ' should this be a system option?
  Public Const OPERATION_TAG_LIMIT As Integer = 5
  
  Public Function ReadAll(ByVal memStream As MemoryStream) As String
    Dim pos As Long = memStream.Position
    memStream.Position = 0
    Dim reader As New StreamReader(memStream)
    Dim str = reader.ReadToEnd()
    ' Reset the position so that subsequent writes are correct.    
    memStream.Position = pos
    Return str
    End Function
    
  <WebMethod()> _
  Public Function Update(ByVal GUID As String, ByVal CustomerXML As String) As String
    Dim MyCommon As New Copient.CommonInc
    Dim MyMassUpdate As Copient.MassUpdate
    Dim Parameters As New Copient.MassUpdate.ProcessingInstructions
    Dim RetXmlDoc As New XmlDocument
    Dim RetXmlStr As String = ""
    Dim RetCode As StatusCodes = StatusCodes.SUCCESS
    Dim RetMsg As String = ""
    Dim XsdType As Integer = 0
    Dim CardID As String = ""
    Dim CustomerXmlDoc As New XmlDocument()
    Dim ConnInc As New Copient.ConnectorInc
    Dim TagCt As Integer = 0
    Dim bFlag As Boolean = False
    Dim node1 As XmlNode 
    Dim root As XmlElement

    
    Try
      If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
      If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
      
      'Remove all invalid unicode characters 
      ' Encode the XML string in a UTF-8 byte array     
      Dim encodedString As Byte() = Encoding.UTF8.GetBytes(CustomerXML)
      ' Put the byte array into a stream and rewind it to the beginning     
      Dim ms As New MemoryStream(encodedString)
      ms.Flush()
      ms.Position = 0      
      CustomerXML = ReadAll(ms)
      
      MyMassUpdate = New Copient.MassUpdate(MyCommon, Copient.MassUpdate.CALLER.CUSTOMER_UPDATE_WS)
      MyMassUpdate.EnableLogging(LogFile)
      
      ' validate the request
      If Not IsValidGUID(GUID, "Update", MyCommon) Then
        RetCode = StatusCodes.INVALID_GUID
        RetMsg = "GUID " & GUID & " is not valid for the Customer Update web service."
      ElseIf Not ConnInc.ConvertStringToXML(CustomerXML, CustomerXmlDoc) Then
        RetCode = StatusCodes.INVALID_XML_DOCUMENT
        RetMsg = "CustomerXML parameter is not a valid XML Document"
      ElseIf Not IsValidTagCount(CustomerXmlDoc, TagCt) Then
        RetCode = StatusCodes.OPERATION_TAG_LIMIT_EXCEEDED
        RetMsg = "Tag limit of " & OPERATION_TAG_LIMIT & " was exceeded. This request includes " & TagCt & " operational tags."
      End If
      
      MyCommon.Write_Log(LogFile, "TagCt: " & TagCt, True)
      
      If RetCode = StatusCodes.SUCCESS Then
        Parameters.Mode = Copient.MassUpdate.PROCESS_MODE.XML_STRING
        Parameters.CustomerXML = CustomerXmlDoc
        Parameters.Caller = Copient.MassUpdate.CALLER.CUSTOMER_UPDATE_WS
        RetXmlDoc = MyMassUpdate.ProcessCustomerXML(Parameters)
        If RetXmlDoc IsNot Nothing Then
          RetXmlStr = RetXmlDoc.OuterXml
          bFlag = RetXmlStr.Contains("Invalid Unique Constraints")
          'Add statuscode for invalid unique constraints
          If bFlag Then 
           root= RetXmlDoc.DocumentElement  
           node1 = RetXmlDoc.CreateNode(XmlNodeType.Element, "StatusCode", "")  
           node1.InnerText = StatusCodes.INVALID_UNIQUE_CONSTRAINTS
           root.AppendChild(node1)
           RetXmlStr = RetXmlDoc.OuterXml
          End If
        Else
          RetXmlStr = GetErrorXML(StatusCodes.APPLICATION_EXCEPTION, "Error encountered during processing customer XML")
        End If
      Else
        RetXmlStr = GetErrorXML(RetCode, RetMsg)
        MyCommon.Write_Log(LogFile, "Error:  " & RetCode.ToString & " - " & RetMsg, True)
      End If
    Catch argEx As ArgumentException
      Select Case argEx.ParamName.ToUpper
        Case "CUSTOMERXML"
          RetXmlStr = GetErrorXML(StatusCodes.INVALID_XML_DOCUMENT, argEx.Message)
        Case Else
          RetXmlStr = GetErrorXML(StatusCodes.APPLICATION_EXCEPTION, argEx.ToString)
      End Select
    Catch ex As Exception
      RetXmlStr = GetErrorXML(StatusCodes.APPLICATION_EXCEPTION, ex.ToString)
      MyCommon.Write_Log(LogFile, "Error:  " & StatusCodes.APPLICATION_EXCEPTION.ToString & " - " & ex.ToString, True)
    Finally
      If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
      If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
    End Try
    
    Return RetXmlStr
  End Function

  
  Private Function GetErrorXML(ByVal Code As StatusCodes, ByVal Message As String) As String
    Dim ErrorXml As New StringBuilder("<?xml version=""1.0"" encoding=""utf-8""?>")

    ErrorXml.Append("<CustomerUpdate returnCode=""")
    Select Case Code
      Case StatusCodes.INVALID_GUID
        ErrorXml.Append("INVALID_GUID")
      Case StatusCodes.INVALID_XML_DOCUMENT
        ErrorXml.Append("INVALID_XML_DOCUMENT")
      Case StatusCodes.OPERATION_TAG_LIMIT_EXCEEDED
        ErrorXml.Append("OPERATION_TAG_LIMIT_EXCEEDED")
      Case Else
        ' treat everything else as an application exception
        ErrorXml.Append("APPLICATION_EXCEPTION")
    End Select
    ErrorXml.Append(""">")

    If Message <> "" Then
      ErrorXml.Append("<Message>" & Message & "</Message>")
    End If

    ErrorXml.Append("</CustomerUpdate>")

    Return ErrorXml.ToString
  End Function
  
  Private Function IsValidGUID(ByVal GUID As String, ByVal MethodName As String, ByRef MyCommon As Copient.CommonInc) As Boolean
    Dim LocalCommon as New Copient.CommonInc
    Dim IsValid As Boolean = False
    Dim ConnInc As New Copient.ConnectorInc
    Dim MsgBuf As New StringBuilder()
    
    Try
      LocalCommon.Open_LogixRT()
      IsValid = ConnInc.IsValidConnectorGUID(LocalCommon, 45, GUID)
    Catch ex As Exception
      IsValid = False
    Finally
      LocalCommon.Close_LogixRT()
    End Try
    
    ' Log the call
    Try
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

  
  Private Function IsValidTagCount(ByRef CustomerXMLDoc As XmlDocument, ByRef TagCount As Integer) As Boolean
    Dim RootNode As XmlNode = Nothing
    Dim ValidCount As Boolean = False
    
    TagCount = 0
    If CustomerXMLDoc IsNot Nothing Then
      RootNode = CustomerXMLDoc.SelectSingleNode("/CustomerUpdate")
      
      If RootNode IsNot Nothing Then
        TagCount = RootNode.ChildNodes.Count
      End If
    End If
    
    ValidCount = (TagCount <= OPERATION_TAG_LIMIT)

    Return ValidCount
  End Function

End Class