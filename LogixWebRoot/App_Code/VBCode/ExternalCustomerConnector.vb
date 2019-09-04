Imports System
Imports System.Web
Imports System.Web.Services
Imports System.Xml.Serialization
Imports System.Xml.Schema
Imports System.Xml
Imports System.Data
Imports System.IO
Imports Copient.CommonInc

<WebService(Namespace:="http://www.copienttech.com/LogixCustomerService/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class ExternalCustomerConnector
  Inherits System.Web.Services.WebService

  ' Return error codes
  Public Const ERROR_NONE As Integer = 0
  Public Const ERROR_ADD_CUSTOMER_FAILED As Integer = 1
  Public Const ERROR_ADD_CUSTOMER_TO_OFFER_FAILED As Integer = 2
  Public Const ERROR_XML_INVALID_DOC As Integer = 3
  Public Const ERROR_XML_EMPTY_DOC As Integer = 4
  Public Const ERROR_XML_BAD_SOURCE_ID As Integer = 5
  Public Const ERROR_XSD_NOT_FOUND As Integer = 6
  Public Const ERROR_APPLICATION As Integer = 7
  Public Const ERROR_INVALID_CUSTOMER_ID As Integer = 8
  Public Const ERROR_INVALID_OFFER_ID As Integer = 9

  Public Enum CUSTOMER_TYPE As Integer
    CUSTOMER = 0
    HOUSEHOLD = 1
  End Enum

  Public Enum CARD_STATUS As Integer
    ACTIVE = 1
    INACTIVE = 2
    CANCELED = 3
    EXPIRED = 4
    LOST_STOLEN = 5
    DEFAULT_CARD = 6
  End Enum

  Private Enum RESPONSE_TYPES As Integer
    ADD_CUSTOMER = 1
    UPDATE_OFFER = 2
    REMOVE_OFFER = 3
    ADD_BANNER = 4
    REMOVE_BANNER = 5
    UPDATE_CLIENT_ID = 6
  End Enum


    Private MyCommon As New Copient.CommonInc
    Private MyCryptLib As New Copient.CryptLib

  <WebMethod()> _
  Public Function AddCustomer(ByVal customerXml As String, ByVal ExternalSourceID As Integer) As String
        Dim ExtCustomerID As String = ""
        Dim firstName As String = ""
        Dim lastName As String = ""
        Dim CustomerPK As Long
    Dim ErrorCode As Integer = ERROR_NONE
    Dim ErrorMsg As String = ""
    Dim custXmlDoc As New XmlDocument
    Dim Updated As Boolean = False

    Try
      If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
      If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

      If IsValidDocument(customerXml, "WS_Customer.xsd", ErrorCode, ErrorMsg) Then

        custXmlDoc.LoadXml(customerXml)

                ExtCustomerID = GetCustomerExtID(custXmlDoc, ErrorCode, ErrorMsg)
                firstName = GetCustomerFirstName(custXmlDoc, ErrorCode, ErrorMsg)
                lastName = GetCustomerLastName(custXmlDoc, ErrorCode, ErrorMsg)

                ' Add the customer (or return the existing CustomerPK for this ExtCustomerID)
                CustomerPK = HandleAddCustomerRec(ExtCustomerID, firstName, lastName)

        ' update the customer data
        If CustomerPK > 0 Then
          Updated = UpdateCustomerRec(custXmlDoc, ExtCustomerID, CustomerPK, ErrorCode, ErrorMsg)
          If Not Updated AndAlso ErrorMsg = "" Then
            ErrorCode = ERROR_ADD_CUSTOMER_FAILED
            ErrorMsg = "Customer " & ExtCustomerID & " added, but failed to update customer record with name information."
          End If

          Updated = UpdateCustomerExt(custXmlDoc, ExtCustomerID, CustomerPK, ErrorCode, ErrorMsg)
          If Not Updated AndAlso ErrorMsg = "" Then
            ErrorCode = ERROR_ADD_CUSTOMER_FAILED
            ErrorMsg = "Customer " & ExtCustomerID & " added, but failed to update customer extension record."
          End If
        End If

      Else
        ErrorCode = ERROR_XML_INVALID_DOC
      End If
    Catch ex As Exception
      ErrorCode = ERROR_APPLICATION
      ErrorMsg = "Add Customer encountered the following error: " & ex.ToString
    Finally
      If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
      If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
    End Try
        If ErrorMsg <> ""
            Return ErrorMsg
            Else
            Return "Add Customer"
        End If
    
  End Function


  <WebMethod()> _
  Public Function GetCustomer(ByVal CustomerID As String, ByVal CustomerType As Integer) As String
    Return "Get Customer"
  End Function


  Private Function IsValidDocument(ByVal customerXml As String, ByVal xsdName As String, ByRef ErrorCode As Integer, ByRef ErrorMsg As String) As Boolean
    Dim sXsdFileName As String = ""
    Dim Settings As XmlReaderSettings
    Dim xr As XmlReader = Nothing
    Dim bValid As Boolean = True

    If (customerXml IsNot Nothing OrElse customerXml.Trim <> "") Then
      sXsdFileName = LoadXsdFileName(xsdName, ErrorCode, ErrorMsg)

      If (ErrorMsg = "") Then
        Try

          Settings = New XmlReaderSettings()
          Settings.Schemas.Add(Nothing, sXsdFileName)
          Settings.ValidationType = ValidationType.Schema
          Settings.IgnoreComments = True
          Settings.IgnoreProcessingInstructions = True
          Settings.IgnoreWhitespace = True
          xr = XmlReader.Create(New StringReader(customerXml), Settings)

          Do While (xr.Read())
            'Console.WriteLine("NodeType: " & xr.NodeType.ToString & " - " & xr.LocalName & " Depth: " & xr.Depth.ToString)
          Loop
          bValid = True
        Catch eXmlSch As XmlSchemaException
          ErrorCode = ERROR_XML_INVALID_DOC
          ErrorMsg = "(Xml Schema Validation Error Line: " & eXmlSch.LineNumber.ToString & " - Col: " & eXmlSch.LinePosition.ToString & ") " & eXmlSch.Message
          bValid = False
        Catch eXml As XmlException
          ErrorCode = ERROR_XML_INVALID_DOC
          ErrorMsg = "(Xml Error Line: " & eXml.LineNumber.ToString & " - Col: " & eXml.LinePosition.ToString & ") " & eXml.Message
          bValid = False
        Catch exApp As ApplicationException
          ErrorCode = ERROR_APPLICATION
          ErrorMsg = "Application Error: " & exApp.ToString
          bValid = False
        Catch ex As Exception
          ErrorCode = ERROR_APPLICATION
          ErrorMsg = "Error: " & ex.ToString
          bValid = False
        Finally
          If Not xr Is Nothing Then
            xr.Close()
          End If
        End Try

      End If
    Else
      bValid = False
      ErrorCode = ERROR_XML_EMPTY_DOC
      ErrorMsg = "XML document is empty"
    End If
       
    Return bValid
  End Function

  Private Function LoadXsdFileName(ByVal xsdName As String, ByRef ErrorCode As Integer, ByRef ErrorMsg As String) As String
    Dim xsdFileName As String = ""

    If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

    xsdFileName = MyCommon.Get_Install_Path & "AgentFiles\" & xsdName

    ' ensure that the schema exists on the server
    If xsdFileName <> "" And Not System.IO.File.Exists(xsdFileName) Then
      ErrorCode = ERROR_XSD_NOT_FOUND
      ErrorMsg = "XSD file not found: " & xsdFileName
      xsdFileName = ""
    End If

    Return xsdFileName
  End Function
    Private Function GetCustomerFirstName(ByVal custXmlDoc As XmlDocument, ByRef ErrorCode As Integer, ByRef ErrorMsg As String) As String
        Dim firstName As String = ""
        Dim firstNameNode As XmlNode

        Try
            firstNameNode = custXmlDoc.SelectSingleNode("//CustomerService/AddCustomer/FirstName")
            If (firstNameNode IsNot Nothing) Then
                firstName = firstNameNode.InnerText
            End If

        Catch ex As Exception
            ErrorCode = ERROR_APPLICATION
            ErrorMsg = "Error encountered in GetCustomerFirstName : " & ex.ToString
        End Try

        Return firstName
    End Function
    Private Function GetCustomerLastName(ByVal custXmlDoc As XmlDocument, ByRef ErrorCode As Integer, ByRef ErrorMsg As String) As String
        Dim lastName As String = ""
        Dim lastNameNode As XmlNode

        Try
            lastNameNode = custXmlDoc.SelectSingleNode("//CustomerService/AddCustomer/LastName")
            If (lastNameNode IsNot Nothing) Then
                lastName = lastNameNode.InnerText
            End If

        Catch ex As Exception
            ErrorCode = ERROR_APPLICATION
            ErrorMsg = "Error encountered in GetCustomerLastName : " & ex.ToString
        End Try

        Return lastName
    End Function
    Private Function GetCustomerExtID(ByVal custXmlDoc As XmlDocument, ByRef ErrorCode As Integer, ByRef ErrorMsg As String) As String
    Dim CustomerExtID As String = ""
    Dim IdNode As XmlNode

    Try
      IdNode = custXmlDoc.SelectSingleNode("//CustomerService/AddCustomer/CustomerID")
      If (IdNode IsNot Nothing) Then
        CustomerExtID = IdNode.InnerText
      End If

      If CustomerExtID.Trim = "" Then
        ErrorCode = ERROR_INVALID_CUSTOMER_ID
        ErrorMsg = "No Customer ID found in XML document"
      End If

    Catch ex As Exception
      ErrorCode = ERROR_APPLICATION
      ErrorMsg = "Error encountered in GetCustomerExtID : " & ex.ToString
    End Try

    Return CustomerExtID
  End Function

    Private Function HandleAddCustomerRec(ByVal ExtCustomerID As String, ByVal firstName As String, ByVal lastName As String) As Long
        Dim CustomerPK As Long

        If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
        ExtCustomerID = MyCommon.Pad_ExtCardID(ExtCustomerID, 0)

        MyCommon.QueryStr = "dbo.pa_ServiceCustomers_Insert"
        MyCommon.Open_LXSsp()
        MyCommon.LXSsp.Parameters.Add("@ExtCardID", SqlDbType.NVarChar, 400).Value = MyCryptLib.SQL_StringEncrypt(ExtCustomerID)
        MyCommon.LXSsp.Parameters.Add("@CardTypeID", SqlDbType.Int).Value = 0
        MyCommon.LXSsp.Parameters.Add("@CustomerTypeID", SqlDbType.Int).Value = 0
        MyCommon.LXSsp.Parameters.Add("@FirstName", SqlDbType.NVarChar, 256).Value = firstName
        MyCommon.LXSsp.Parameters.Add("@LastName", SqlDbType.NVarChar, 256).Value = lastName
        MyCommon.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
        MyCommon.LXSsp.ExecuteNonQuery()
        CustomerPK = MyCommon.LXSsp.Parameters("@CustomerPK").Value
        MyCommon.Close_LXSsp()

        Return CustomerPK
    End Function

    Private Function UpdateCustomerRec(ByVal custXmlDoc As XmlDocument, ByVal CustomerID As String, _
                                     ByVal CustomerPK As Long, ByRef ErrorCode As Integer, ByRef ErrorMsg As String) As Boolean
    Dim Updated As Boolean = False
    Dim TempStr As String = ""
    Dim CustTypeID As Integer = CUSTOMER_TYPE.CUSTOMER
    Dim CustNode As XmlNode = Nothing
    Dim TempNode As XmlNode = Nothing
    Dim SqlBuf As New StringBuilder()
    Dim ValTable As New Hashtable(15)
    Dim i As Integer = 0
    Dim enumerator As IDictionaryEnumerator
    Dim Cols() As String = New String() {"FirstName", "LastName", "Employee", "CardStatusID", "CustomerTypeID"}
    Dim ColDataTypes() As System.TypeCode = New System.TypeCode() _
                                        {TypeCode.String, TypeCode.String, TypeCode.Boolean, TypeCode.Int32, TypeCode.Int32}

    Try
      If (MyCommon.LXSadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixXS()

      CustNode = custXmlDoc.SelectSingleNode("//CustomerService/AddCustomer")

      If CustNode IsNot Nothing Then
        For i = 0 To Cols.GetUpperBound(0)
          PutChildNodeValue(CustNode, Cols(i), ColDataTypes(i), ValTable)
        Next

        TempStr = GetChildNodeValue(CustNode, "CardStatusID")
        If TempStr IsNot Nothing AndAlso TempStr.Trim <> "" Then
          ValTable.Add("CardStatusID", GetCardStatusID(TempStr).ToString)
        End If

        TempStr = GetChildNodeValue(CustNode, "CustomerTypeID")
        If TempStr IsNot Nothing AndAlso TempStr.Trim <> "" Then
          Integer.TryParse(GetCustomerTypeID(TempStr), CustTypeID)
          ValTable.Add("CustomerTypeID", CustTypeID.ToString)
        End If

        If (ValTable.Count > 0) Then
          SqlBuf.Append("update Customers with (RowLock) set ")

          enumerator = ValTable.GetEnumerator
          While enumerator.MoveNext()
            SqlBuf.Append(enumerator.Key.ToString)
            SqlBuf.Append("=")
            SqlBuf.Append(enumerator.Value.ToString)
            SqlBuf.Append(",")
          End While

          If CustTypeID = CUSTOMER_TYPE.HOUSEHOLD Then SqlBuf.Append(" HHPK=0,")
          SqlBuf.Append("CPEStoreSendFlag = 1 where CustomerPK=" & CustomerPK & ";")

          MyCommon.QueryStr = SqlBuf.ToString
          MyCommon.LRT_Execute()

          Updated = (MyCommon.RowsAffected > 0)
          MyCommon.Close_LogixRT()
        Else
          ' nothing to update
          Updated = True
        End If

      End If

    Catch ex As Exception
      ErrorCode = ERROR_ADD_CUSTOMER_FAILED
      ErrorMsg = "Customer: " & CustomerID & " failed to update due to the following reasons: " & ex.ToString
      Updated = False
    End Try

    Return Updated
  End Function

  Private Function UpdateCustomerExt(ByVal custXmlDoc As XmlDocument, ByVal CustomerID As String, _
                                     ByVal CustomerPK As Long, ByRef ErrorCode As Integer, ByRef ErrorMsg As String) As Boolean
    Dim Updated As Boolean = False
    Dim dt As DataTable
    Dim TempStr As String = ""
    Dim CustTypeID As Integer = 0
    Dim BirthDate As Date = Nothing
    Dim DOB As String = ""
    Dim CustNode As XmlNode = Nothing
    Dim SqlBuf As New StringBuilder()
    Dim ValTable As New Hashtable(15)
    Dim i As Integer = 0
    Dim enumerator As IDictionaryEnumerator
    Dim Cols() As String = New String() {"Phone", "Address", "City", _
                                         "State", "ZIP", "Email", "Country"}
    Dim ColDataTypes() As System.TypeCode = New System.TypeCode() _
                                        {TypeCode.String, TypeCode.String, TypeCode.String, _
                                         TypeCode.String, TypeCode.String, TypeCode.String, TypeCode.String}

    Try
      If (MyCommon.LXSadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixXS()

      CustNode = custXmlDoc.SelectSingleNode("//CustomerService/AddCustomer")

      If CustNode IsNot Nothing Then
        For i = 0 To Cols.GetUpperBound(0)
          PutChildNodeValue(CustNode, Cols(i), ColDataTypes(i), ValTable)
        Next

        If ValTable.Count > 0 AndAlso ValTable.ContainsKey(Cols(4)) Then
          TempStr = ValTable.Item(Cols(4))
          If (TempStr IsNot Nothing AndAlso TempStr.ToUpper = "HOUSEHOLD") Then CustTypeID = 1
        End If

        ' format the Date of Birth(DOB) to comply with column datatype in CustomerExt.
        TempStr = GetChildNodeValue(CustNode, "DOB")
        If (TempStr IsNot Nothing) Then
          If Date.TryParse(TempStr, BirthDate) Then
            DOB = BirthDate.ToString("MMddYYYY")
          End If
        End If

        If (ValTable.Count > 0) Then
          ' check if an CustomerExt record already exists, if so update it otherwise add one
          MyCommon.QueryStr = "select CustomerPK from CustomerExt with (NoLock) where CustomerPK =" & CustomerPK
          dt = MyCommon.LXS_Select

          If (dt.Rows.Count > 0) Then
            SqlBuf.Append("update CustomerExt with (RowLock) set ")

            i = 0
            enumerator = ValTable.GetEnumerator
            While enumerator.MoveNext()
              If (i > 0) Then SqlBuf.Append(",")
              SqlBuf.Append(enumerator.Key.ToString)
              SqlBuf.Append("=")
              SqlBuf.Append(enumerator.Value.ToString)
              i += 1
            End While

            If (DOB <> "") Then
              If (i > 0) Then SqlBuf.Append(",")
              SqlBuf.Append("DOB='" & DOB & "' ")
            End If

            SqlBuf.Append(" where CustomerPK=" & CustomerPK & ";")
          Else
            SqlBuf.Append("insert into CustomerExt with (RowLock) (")

            enumerator = ValTable.GetEnumerator
            While enumerator.MoveNext()
              SqlBuf.Append(enumerator.Key.ToString)
              SqlBuf.Append(",")
            End While

            If (DOB <> "") Then SqlBuf.Append("DOB, ")

            SqlBuf.Append("CustomerPK) values (")

            enumerator.Reset()
            While enumerator.MoveNext()
              SqlBuf.Append(enumerator.Value.ToString)
              SqlBuf.Append(",")
            End While
            If (DOB <> "") Then SqlBuf.Append(DOB & ",")
            SqlBuf.Append(CustomerPK)
            SqlBuf.Append(");")

          End If

          MyCommon.QueryStr = SqlBuf.ToString
          MyCommon.LRT_Execute()

          Updated = (MyCommon.RowsAffected > 0)
          MyCommon.Close_LogixRT()
        Else
          ' nothing to update
          Updated = True
        End If

      End If

    Catch ex As Exception
      ErrorCode = ERROR_ADD_CUSTOMER_FAILED
      ErrorMsg = "Customer: " & CustomerID & " failed to update due to the following reasons: " & ex.ToString
      Updated = False
    End Try

    Return Updated
  End Function

  Private Function GetCustomerPK(ByVal ExtCustomerID As String) As Long
    Dim dt As DataTable
    Dim CustomerPK As Long = 0
    If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
        'Private function and no references found.
    MyCommon.QueryStr = "select CustomerPK from Customers with (NoLock) where PrimaryExtID='" & ExtCustomerID & "' order by CustomerTypeID;"
    dt = MyCommon.LXS_Select
    If (dt.Rows.Count > 0) Then
      CustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)
    End If

    Return CustomerPK
  End Function


  Private Sub UpdateAttributes(ByRef Node As XmlNode, ByVal Added As Boolean, ByVal Removed As Boolean, ByVal ErrorCode As Integer, ByVal ErrorMsg As String)

    If Node IsNot Nothing Then
      WriteAttributesToNode(Node, "added", IIf(Added, "true", "false"))
      WriteAttributesToNode(Node, "removed", IIf(Removed, "true", "false"))
      If (ErrorCode > 0) Then
        WriteAttributesToNode(Node, "errorcode", ErrorCode.ToString)
      End If
      If (ErrorMsg.Trim <> "") Then
        WriteAttributesToNode(Node, "errormessage", ErrorMsg)
      End If
    End If
  End Sub

  Private Sub WriteAttributesToNode(ByRef Node As XmlNode, ByVal AttribName As String, ByVal AttribVal As String)
    Dim TempAttrib, NewAttrib As XmlAttribute
    Dim xmlDoc As New XmlDocument

    TempAttrib = Node.Attributes(AttribName)
    ' create the attribute if it doesn't already exist to provide feedback to caller
    If (TempAttrib Is Nothing) Then
      NewAttrib = xmlDoc.CreateAttribute(AttribName)
      TempAttrib = Node.Attributes.Append(NewAttrib)
    End If

    If TempAttrib IsNot Nothing Then
      TempAttrib.InnerText = AttribVal
    End If

  End Sub

  Private Function GetChildNodeValue(ByVal ParentNode As XmlNode, ByVal ChildNodeName As String) As String
    Dim Value As String = Nothing
    Dim ChildNode As XmlNode

    If Not (ParentNode Is Nothing) Then
      ChildNode = ParentNode.SelectSingleNode(ChildNodeName)
      If Not (ChildNode Is Nothing) Then
        Value = ChildNode.InnerText
      End If
    End If

    Return Value
  End Function

  Private Function PutChildNodeValue(ByVal ParentNode As XmlNode, ByVal ChildNodeName As String, ByVal ColDataType As System.TypeCode, ByRef ValTable As Hashtable) As String
    Dim Value As String = Nothing
    Dim ChildNode As XmlNode
    Dim bTemp As Boolean

    If Not (ParentNode Is Nothing) Then
      ChildNode = ParentNode.SelectSingleNode(ChildNodeName)
      If Not (ChildNode Is Nothing) Then
        Value = ChildNode.InnerText
        Select Case ColDataType
          Case TypeCode.Boolean
            Boolean.TryParse(Value, bTemp)
            Value = IIf(bTemp, "1", "0")
          Case TypeCode.Char, TypeCode.String
            Value = "'" & Value & "'"
          Case TypeCode.DateTime
            Value = "convert(datetime, '" & Value & "')"
                    Case Else
                        If UCase(ChildNodeName) = "InitialCardID" Then
                            Value = MyCryptLib.SQL_StringEncrypt(Value)
                        Else
                            Value = Value
                        End If
                End Select
        ValTable.Add(ChildNodeName, Value)
      End If
    End If

    Return Value
  End Function

  Private Function GetCardStatusID(ByVal Code As String) As Integer
    Dim CardStatusID As Integer

    Select Case Code.Trim.ToUpper
      Case "ACTIVE"
        CardStatusID = CARD_STATUS.ACTIVE
      Case "INACTIVE"
        CardStatusID = CARD_STATUS.INACTIVE
      Case "CANCELED"
        CardStatusID = CARD_STATUS.CANCELED
      Case "EXPIRED"
        CardStatusID = CARD_STATUS.EXPIRED
      Case "LOST_STOLEN"
        CardStatusID = CARD_STATUS.LOST_STOLEN
      Case "DEFAULT_CARD"
        CardStatusID = CARD_STATUS.DEFAULT_CARD
      Case Else
        CardStatusID = CARD_STATUS.ACTIVE
    End Select

    Return CardStatusID
  End Function

  Private Function GetCustomerTypeID(ByVal Code As String) As Integer
    Dim CustomerTypeID As Integer

    Select Case Code.Trim.ToUpper
      Case "CUSTOMER"
        CustomerTypeID = CUSTOMER_TYPE.CUSTOMER
      Case "HOUSEHOLD"
        CustomerTypeID = CUSTOMER_TYPE.HOUSEHOLD
      Case Else
        CustomerTypeID = CUSTOMER_TYPE.CUSTOMER
    End Select

    Return CustomerTypeID
  End Function

End Class

