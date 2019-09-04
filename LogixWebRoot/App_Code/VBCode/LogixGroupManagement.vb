Imports System
Imports System.Web
Imports System.Web.Services
Imports System.Xml.Serialization
Imports System.Xml.Schema
Imports System.Xml
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Collections.Generic
Imports System.Collections.Specialized.NameValueCollection
Imports System.Threading.Thread
Imports System.Security.SecurityElement

<WebService(Namespace:="http://www.copienttech.com/LogixGroupManagement/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class LogixGroupManagement
    Inherits System.Web.Services.WebService

    Private MyCommon As New Copient.CommonInc
    Private MyCryptLib As New Copient.CryptLib
    Private MyMassUpdate As Copient.MassUpdate
    Private MyImport As New Copient.ImportXml
    Private MyCGroupOp As New Copient.CustomerGroupOperations
    Private MyProdGroupOp As New Copient.ProductGroupOperations
    Private MyLocGroupOp As New Copient.Location(MyCommon)
    'Private MyLog As New Copient.Log4Net 
    'Private WSthread As Integer = System.Threading.Thread.CurrentThread.ManagedThreadId

    Private LGMLogFile As String = "LogixGroupManagementWSLog." & Date.Now.ToString("yyyyMMdd") & ".txt"
    Private Const CONNECTOR_ID As Integer = 56

    Public Const OPERATION_TAG_LIMIT As Integer = 1

    Public Enum StatusCodes As Integer
        SUCCESS = 0
        INVALID_GUID = 1
        INVALID_CUSTOMERGROUPID = 2
        INVALID_CUSTOMERGROUPNAME = 3
        INVALID_CARDID = 4
        INVALID_CARDTYPEID = 5
        INVALID_CUSTOMERNAME = 6
        INVALID_OPERATIONTYPE = 7
        INVALID_XML_DOCUMENT = 8
        INVALID_CRITERIA_XML = 9
        INVALID_INCOMPLETEPROCESSCUSTDATA = 10
        INVALID_LOCATIONGROUPID = 11
        INVALID_LOCATIONNAME = 12
        INVALID_LOCATIONCODE = 13
        INVALID_BANNER = 14
        INVALID_INCOMPLETEPROCESSLOCDATA = 15
        INVALID_EXTPRODGROUPID = 16
        INVALID_PRODGROUPNAME = 17
        INVALID_PRODUCTTYPEID = 18
        INVALID_PRODGROUPID = 19
        INVALID_PRODUCTID = 20
        INVALID_INCOMPLETEPROCESSPRODDATA = 21
        FAILED_OPTIN = 22
        INVALID_DESCRIPTION = 23
        APPLICATION_EXCEPTION = 9999
    End Enum

#Region "Common"

    Private Sub InitApp()
        MyCommon.AppName = "LogixGroupManagement.asmx"

        Try
        Catch eXmlSch As XmlSchemaException
        Catch ex As Exception

        End Try
    End Sub

    Private Function IsValidGUID(ByVal GUID As String, ByVal MethodName As String) As Boolean
        Dim LocalCommon As New Copient.CommonInc
        LocalCommon.AppName = "LogixGroupManagement.asmx"
        Dim IsValid As Boolean = False
        Dim ConnInc As New Copient.ConnectorInc
        Dim MsgBuf As New StringBuilder()
        Try
            LocalCommon.Open_LogixRT()
            IsValid = ConnInc.IsValidConnectorGUID(LocalCommon, CONNECTOR_ID, GUID)
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
            Copient.Logger.Write_Log(LGMLogFile, MsgBuf.ToString, True)
        Catch ex As Exception
            ' ignore
        End Try

        Return IsValid
    End Function


#End Region


#Region "SingleCustomerGroup"

    Private Function isValidCustomerCardData(ByVal ExtCardID As String, ByVal Isnumeric As Boolean) As Boolean
        Dim lExtCardID As Long = 0
        Try
            If Isnumeric = True Then
                lExtCardID = Convert.ToInt64(ExtCardID)
            End If
            isValidCustomerCardData = True
        Catch ex As Exception
            isValidCustomerCardData = False
        End Try
    End Function

    Private Function _ProcessCardInCustomerGroup(ByVal GUID As String, ByVal ExtGroupID As String, ByVal GroupID As Long, ByVal GroupName As String, _
                                                ByVal ExtCardId As String, ByVal CardTypeID As String, _
                                                ByVal OperationType As String, ByVal ByExtID As Boolean) As System.Data.DataSet
        Dim ResultSet As New System.Data.DataSet("LogixGroupManagement")
        Dim MethodName As String = "ProcessCardInCustomerGroup"
        Dim dtStatus As DataTable
        Dim row As System.Data.DataRow
        Dim ErrorMsg As String = ""
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim RetMsgLog As String = ""
        Dim Name As String = ""
        Dim ErrorMessage As String = ""
        Dim CamStatus As Boolean = False

        Dim NumericOnly As Boolean = False
        Dim CardValidationResponse As Copient.CommonInc.CardValidationResponse

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        If isValidInputText(GUID) = False Then
            'GUID Value is empty
            RetCode = StatusCodes.INVALID_GUID
            RetMsg = "Failure. Invalid GUID"
            row = dtStatus.NewRow()
            row.Item("StatusCode") = RetCode
            row.Item("Description") = RetMsg
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            Copient.Logger.Write_Log(LGMLogFile, RetMsgLog, True)
            _ProcessCardInCustomerGroup = ResultSet
            Exit Function
        End If
        If ByExtID AndAlso (isValidInputText(ExtGroupID) = False Or ExtGroupID.Trim = "0") Then
            'ExtGroupID Value is empty or invalid
            RetCode = StatusCodes.INVALID_CUSTOMERGROUPID
            RetMsg = "Failure. Invalid Customer GroupID"
            row = dtStatus.NewRow()
            row.Item("StatusCode") = RetCode
            row.Item("Description") = RetMsg

            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            Copient.Logger.Write_Log(LGMLogFile, RetMsgLog, True)
            _ProcessCardInCustomerGroup = ResultSet
            Exit Function
        End If
        If isValidInputText(ExtCardId) = False Then
            'Card ID Value is empty or invalid
            RetCode = StatusCodes.INVALID_CARDID
            RetMsg = "Failure. Invalid Customer CardID"
            row = dtStatus.NewRow()
            row.Item("StatusCode") = RetCode
            row.Item("Description") = RetMsg
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            Copient.Logger.Write_Log(LGMLogFile, RetMsgLog, True)
            _ProcessCardInCustomerGroup = ResultSet
            Exit Function
        End If 'ExtCardId

        'Validate ExtCardId and CardType id
        If MyCommon.AllowToProcessCustomerCard(ExtCardId, CardTypeID, CardValidationResponse) = False Then
            If (CardValidationResponse = Copient.CommonInc.CardValidationResponse.CARDTYPENOTFOUND Or CardValidationResponse = Copient.CommonInc.CardValidationResponse.INVALIDCARDTYPEFORMAT) Then
                RetCode = StatusCodes.INVALID_CARDTYPEID
            ElseIf (CardValidationResponse = Copient.CommonInc.CardValidationResponse.CARDIDNOTNUMERIC Or CardValidationResponse = Copient.CommonInc.CardValidationResponse.INVALIDCARDFORMAT) Then
                RetCode = StatusCodes.INVALID_CARDID
            ElseIf (CardValidationResponse = Copient.CommonInc.CardValidationResponse.ERROR_APPLICATION) Then
                RetCode = StatusCodes.APPLICATION_EXCEPTION
            End If
            RetMsg = MyCommon.CardValidationResponseMessage(ExtCardId, CardTypeID, CardValidationResponse, RetMsgLog)

            row = dtStatus.NewRow()
            row.Item("StatusCode") = RetCode
            row.Item("Description") = RetMsg

            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            If (String.IsNullOrWhiteSpace(RetMsgLog)) Then
                Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
            Else
                Copient.Logger.Write_Log(LGMLogFile, RetMsgLog, True)
            End If
            _ProcessCardInCustomerGroup = ResultSet
            Exit Function
        End If

        If isValidInputText(OperationType) = False Then
            row = dtStatus.NewRow()
            RetCode = StatusCodes.INVALID_OPERATIONTYPE
            RetMsg = "Failure. Invalid Operation"
            row.Item("StatusCode") = RetCode
            row.Item("Description") = RetMsg
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
            _ProcessCardInCustomerGroup = ResultSet
            Exit Function
        End If

        If isValidInputText(GroupName, False) = False Then
            row = dtStatus.NewRow()
            RetCode = StatusCodes.INVALID_CUSTOMERGROUPNAME
            RetMsg = "Failure. Invalid Customer Group Name"
            row.Item("StatusCode") = RetCode
            row.Item("Description") = RetMsg
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
            _ProcessCardInCustomerGroup = ResultSet
            Exit Function
        End If

        OperationType = OperationType.ToLower
        If OperationType <> "augment" AndAlso OperationType <> "remove" AndAlso OperationType <> "replace" Then
            row = dtStatus.NewRow()
            RetCode = StatusCodes.INVALID_OPERATIONTYPE
            RetMsg = "Failure. Invalid Operation"
            row.Item("StatusCode") = RetCode
            row.Item("Description") = RetMsg
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
            _ProcessCardInCustomerGroup = ResultSet
            Exit Function
        End If

        Try
            Using MyCommon.LXSadoConn
                If MyCommon.LXSadoConn.State = ConnectionState.Closed Then
                    MyCommon.Open_LogixXS()
                End If
                If Not IsValidGUID(GUID, MethodName) Then
                    RetCode = StatusCodes.INVALID_GUID
                    RetMsg = "Failure. Invalid GUID"
                Else
                    If RetCode = StatusCodes.SUCCESS Then
                        Dim CustGroupExists As Boolean
                        If ByExtID Then
                            CustGroupExists = MyCGroupOp.CustomerGroupExists(ExtGroupID, GroupID, Name)
                        Else
                            CustGroupExists = MyCGroupOp.CustomerGroupExists(GroupID, Name)
                        End If
                        If CustGroupExists Then
                            If Not (GroupName = "") AndAlso Not (GroupName = Name) Then
                                MyCGroupOp.CustomerGroupNameExists(LGMLogFile, GroupID, GroupName, RetCode, RetMsg)
                            End If
                        Else
                            If Not (GroupName = "") Then
                                MyCGroupOp.CustomerGroupNameExists(LGMLogFile, GroupID, GroupName, RetCode, RetMsg)
                            Else
                                RetCode = StatusCodes.INVALID_CUSTOMERGROUPNAME
                                RetMsg = "Failure. GroupName is not provided"
                            End If
                            If RetCode = StatusCodes.SUCCESS Then
                                If OperationType.ToUpper <> "REMOVE" Then
                                    If ByExtID = True Then
                                        GroupID = MyCGroupOp.CreateCustomerGroupByExtID(LGMLogFile, ExtGroupID, ErrorMessage, GroupName)
                                    Else
                                        GroupID = MyCGroupOp.CreateCustomerGroupByLogixID(LGMLogFile, CamStatus, ErrorMessage, GroupName)
                                    End If
                                End If
                            End If
                        End If

                    End If
                    If RetCode = StatusCodes.SUCCESS Then
                        ProcessBulkCustGroupData(ExtCardId & "," & CardTypeID, GroupID, OperationType, RetCode, RetMsg, RetMsg)
                    Else
                        Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                    End If
                End If
            End Using
        Catch ex As Exception
            RetCode = StatusCodes.APPLICATION_EXCEPTION
            RetMsg = "Failure. Application " & ex.ToString
        Finally
            If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count = 0 Then
                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)
            End If
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
        End Try
        Return ResultSet
    End Function

#End Region

#Region "MultipleCustomerGroup"

    Private Function _ProcessMultipleCardInCustomerGroup(ByVal GUID As String, _
                                                      ByVal CustomerXML As String, ByVal ByExtID As Boolean) As String

        Dim LocalCGroupOp As New Copient.CustomerGroupOperations
        Dim LocalImport As New Copient.ImportXml
        Dim ResultSet As New System.Data.DataSet("LogixGroupManagement")
        Dim MethodName As String = "_ProcessMultipleCardInCustomerGroup"
        Dim RetXmlDoc As New XmlDocument
        Dim RetXmlStr As String = ""
        Dim ConnInc As New Copient.ConnectorInc
        Dim CustomerXmlDoc As New XmlDocument()
        Dim bRTConnectionOpened As Boolean = False
        Dim bXSConnectionOpened As Boolean = False

        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim ReturnCode As StatusCodes = StatusCodes.SUCCESS ' This is what will eventually be the ReturnCode attribute. It will only be a success if all of the status codes are success
        Dim RetMsg As String = ""
        Dim RetMsgLog As String = ""
        Dim ExtGroupID As String = ""
        Dim GroupName As String = ""
        Dim Name As String = ""
        Dim Operation As String = ""
        Dim CustList As String = ""
        Dim Description As String = ""

        Dim ErrorMsg As String = ""


        Dim CustomerGroupID As Long

        If isValidInputText(GUID) = False Then
            'GUID Value is empty
            RetCode = StatusCodes.INVALID_GUID
            RetMsg = "Failure. Invalid GUID"
            'RetXmlStr = GetCustomerListErrorXML(RetCode, RetMsg, "CustomerGroupUpdate")
            RetXmlStr = GetGroupResponseXML(RetCode, "", "CustomerGroupUpdate", RetMsg)
            Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
            _ProcessMultipleCardInCustomerGroup = RetXmlStr
            Exit Function
        End If
        If CustomerXML = "" Then
            'CustomerXML is empty
            RetCode = StatusCodes.INVALID_XML_DOCUMENT
            RetMsg = "Failure. Invalid CustomerXML"
            'RetXmlStr = GetCustomerListErrorXML(RetCode, RetMsg, "CustomerGroupUpdate")
            RetXmlStr = GetGroupResponseXML(RetCode, "", "CustomerGroupUpdate", RetMsg)
            Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
            _ProcessMultipleCardInCustomerGroup = RetXmlStr
            Exit Function
        End If
        If ByExtID = True Then
            If CustomerXML.Contains("ImportByLogixID") Then
                'CustomerXML is empty
                RetCode = StatusCodes.INVALID_XML_DOCUMENT
                RetMsg = "Failure. Invalid CustomerXML"
                'RetXmlStr = GetCustomerListErrorXML(RetCode, RetMsg, "CustomerGroupUpdate")
                RetXmlStr = GetGroupResponseXML(RetCode, "", "CustomerGroupUpdate", RetMsg)
                Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                _ProcessMultipleCardInCustomerGroup = RetXmlStr
                Exit Function
            End If
        Else
            If CustomerXML.Contains("ImportByExternalID") Then
                'CustomerXML is empty
                RetCode = StatusCodes.INVALID_XML_DOCUMENT
                RetMsg = "Failure. Invalid CustomerXML"
                'RetXmlStr = GetCustomerListErrorXML(RetCode, RetMsg, "CustomerGroupUpdate")
                RetXmlStr = GetGroupResponseXML(RetCode, "", "CustomerGroupUpdate", RetMsg)
                Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                _ProcessMultipleCardInCustomerGroup = RetXmlStr
                Exit Function
            End If
        End If


        Try
            'Remove all invalid Unicode characters 
            ' Encode the XML string in a UTF-8 byte array     
            Dim encodedString As Byte() = Encoding.UTF8.GetBytes(CustomerXML)
            ' Put the byte array into a stream and rewind it to the beginning     
            Dim ms As New MemoryStream(encodedString)
            ms.Flush()
            ms.Position = 0
            CustomerXML = MyCommon.ReadAll(ms)

            If Not IsValidGUID(GUID, MethodName) Then
                RetCode = StatusCodes.INVALID_GUID
                RetMsg = "GUID " & GUID & " is not valid for the LogixGroupManagement web service."
                Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
            ElseIf Not ConnInc.ConvertStringToXML(CustomerXML, CustomerXmlDoc) Then
                'If Not ConnInc.ConvertStringToXML(CustomerXML, CustomerXmlDoc) Then
                RetCode = StatusCodes.INVALID_XML_DOCUMENT
                RetMsg = "CustomerXML parameter is not a valid XML Document"
                Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
            ElseIf Not LocalImport.ValidateXml("CustomerGroupUpdate.xsd", CustomerXML) Then
                RetCode = StatusCodes.INVALID_CRITERIA_XML
                RetMsg = "XML document sent does not confirm to the CustomerGroupUpdate.xsd schema."
                Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
            End If

            If RetCode = StatusCodes.SUCCESS Then
                Dim rootNode As XmlNode = CustomerXmlDoc.DocumentElement
                Dim nodeList As XmlNodeList
                If ByExtID = True Then
                    nodeList = rootNode.SelectNodes("/CustomerGroupUpdate/ImportByExternalID")
                Else
                    nodeList = rootNode.SelectNodes("/CustomerGroupUpdate/ImportByLogixID")
                End If

                For Each node In nodeList
                    RetCode = StatusCodes.SUCCESS

                    CustomerGroupID = 0
                    ResultSet = New DataSet()
                    ResultSet.ReadXml(New XmlNodeReader(CustomerXmlDoc))


                    If Not node.Item("ID") Is Nothing Then
                        If ByExtID = True Then
                            ExtGroupID = node.Item("ID").InnerText
                        Else
                            CustomerGroupID = node.Item("ID").InnerText
                        End If
                    Else
                        ExtGroupID = ""
                        CustomerGroupID = 0
                    End If
                    If Not node.Item("Name") Is Nothing Then
                        GroupName = node.Item("Name").InnerText
                    Else
                        GroupName = ""
                    End If
                    If Not node.Item("Operation") Is Nothing Then
                        Operation = node.Item("Operation").InnerText
                    Else
                        Operation = ""
                    End If

                    'CustListNodes = node.SelectNodes("BulkData")
                    CustList = node.Item("BulkData").InnerText

                    If (ByExtID = True AndAlso (ExtGroupID Is Nothing OrElse ExtGroupID.Trim = "")) Then
                        RetCode = StatusCodes.INVALID_CUSTOMERGROUPID
                        RetMsg = "Failure. CustomerGroupId is not Provided."
                    End If


                    If Operation Is Nothing OrElse Operation.Trim = "" OrElse (Operation <> "augment" AndAlso Operation <> "replace" AndAlso Operation <> "remove") Then
                        RetCode = StatusCodes.INVALID_OPERATIONTYPE
                        RetMsg = "Failure. OperationType is not Provided"
                    End If
                    If RetCode = StatusCodes.SUCCESS Then

                        If ByExtID = True Then
                            If LocalCGroupOp.CustomerGroupExists(ExtGroupID, CustomerGroupID, Name) Then
                                If Not (GroupName = "") AndAlso Not (GroupName = Name) Then
                                    LocalCGroupOp.CustomerGroupNameExists(LGMLogFile, CustomerGroupID, GroupName, RetCode, RetMsg)
                                End If
                            Else
                                If Not (GroupName = "") Then
                                    LocalCGroupOp.CustomerGroupNameExists(LGMLogFile, CustomerGroupID, GroupName, RetCode, RetMsg)
                                Else
                                    RetCode = StatusCodes.INVALID_CUSTOMERGROUPNAME
                                    RetMsg = "Failure. Customer group name is not provided"
                                End If

                                If Operation.ToUpper <> "REMOVE" And RetCode = StatusCodes.SUCCESS Then
                                    CustomerGroupID = LocalCGroupOp.CreateCustomerGroupByExtID(LGMLogFile, ExtGroupID, ErrorMsg, GroupName)
                                ElseIf RetCode = StatusCodes.SUCCESS Then
                                    RetCode = StatusCodes.INVALID_CUSTOMERGROUPID
                                    RetMsg = "Failure. New Customer group cannot be created for operation type - Remove"
                                End If
                            End If
                        Else
                            If LocalCGroupOp.CustomerGroupExists(CustomerGroupID, Name) Then
                                If Not (GroupName = "") AndAlso Not (GroupName = Name) Then
                                    LocalCGroupOp.CustomerGroupNameExists(LGMLogFile, CustomerGroupID, GroupName, RetCode, RetMsg)
                                End If
                            Else
                                If Not (GroupName = "") Then
                                    LocalCGroupOp.CustomerGroupNameExists(LGMLogFile, CustomerGroupID, GroupName, RetCode, RetMsg)
                                Else
                                    RetCode = StatusCodes.INVALID_CUSTOMERGROUPNAME
                                    RetMsg = "Failure. Customer group name is not provided"
                                End If
                                If Operation.ToUpper <> "REMOVE" And RetCode = StatusCodes.SUCCESS Then
                                    CustomerGroupID = LocalCGroupOp.CreateCustomerGroupByLogixID(LGMLogFile, False, ErrorMsg, GroupName)
                                End If
                            End If
                        End If
                        'Customer group creation logic - Ends here 

                        If RetCode = StatusCodes.SUCCESS Then
                            If ByExtID = True Then
                                If Not (ResultSet.Tables("ImportByExternalID") Is Nothing) Then
                                    AddColumnToTable(ResultSet, "ImportByExternalID", "statuscode", Type.GetType("System.String"))
                                    AddColumnToTable(ResultSet, "ImportByExternalID", "Description", Type.GetType("System.String"))
                                End If
                            Else
                                If Not (ResultSet.Tables("ImportByLogixID") Is Nothing) Then
                                    AddColumnToTable(ResultSet, "ImportByLogixID", "statuscode", Type.GetType("System.String"))
                                    AddColumnToTable(ResultSet, "ImportByLogixID", "Description", Type.GetType("System.String"))
                                End If
                            End If
                            'Process start for individual customers
                            ProcessBulkCustGroupData(CustList, CustomerGroupID, Operation, RetCode, RetMsg, RetMsgLog)
                        Else
                            'Unable to process the individual customer records
                            Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                        End If
                    Else
                        'RetXmlStr = GetCustomerListErrorXML(RetCode, RetMsg, "CustomerGroupUpdate")
                        Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                    End If

                    RetXmlStr &= GetInputResponseXML(RetCode, RetMsg, ByExtID, node)
                    If Not RetCode = StatusCodes.SUCCESS Then
                        ReturnCode = RetCode
                    End If

                Next
            End If
            If Not RetCode = StatusCodes.SUCCESS Then
                ReturnCode = RetCode

            End If
            RetXmlStr = GetGroupResponseXML(ReturnCode, RetXmlStr, "CustomerGroupUpdate", RetMsg)

        Catch ex As Exception
            RetCode = StatusCodes.APPLICATION_EXCEPTION
            RetMsg = "Failure. Application exception - " & ex.Message.ToString
            RetXmlStr = GetCustomerListErrorXML(RetCode, RetMsg, "CustomerGroupUpdate")
        End Try
        Return RetXmlStr
    End Function

    Private Function GetGroupResponseXML(ByVal ReturnCode As StatusCodes, ByVal InputRespXML As String, ByVal RootNode As String, Optional ByVal ErrorMessage As String = "") As String
        Dim OutputXML As New StringBuilder("<?xml version=""1.0"" encoding=""utf-8""?>")

        OutputXML.Append("<" & RootNode)
        If InputRespXML = "" Then
            OutputXML.Append(" returnCode=""")
            Select Case ReturnCode
                Case StatusCodes.SUCCESS
                    OutputXML.Append("SUCCESS")
                Case StatusCodes.INVALID_GUID
                    OutputXML.Append("INVALID_GUID")
                Case StatusCodes.INVALID_CUSTOMERGROUPID
                    OutputXML.Append("INVALID_CUSTOMERGROUPID")
                Case StatusCodes.INVALID_XML_DOCUMENT
                    OutputXML.Append("INVALID_XML_DOCUMENT")
                Case StatusCodes.INVALID_CRITERIA_XML
                    OutputXML.Append("INVALID_CRITERIA_XML")
                Case StatusCodes.INVALID_OPERATIONTYPE
                    OutputXML.Append("INVALID_OPERATIONTYPE")
                Case StatusCodes.INVALID_CUSTOMERGROUPNAME
                    OutputXML.Append("INVALID_CUSTOMERGROUPNAME")
                Case StatusCodes.INVALID_INCOMPLETEPROCESSCUSTDATA
                    OutputXML.Append("INVALID_INCOMPLETEPROCESSCUSTDATA")
                Case StatusCodes.INVALID_CARDID
                    OutputXML.Append("INVALID_CARDID")
                Case StatusCodes.INVALID_CARDTYPEID
                    OutputXML.Append("INVALID_CARDTYPEID")
                Case StatusCodes.INVALID_LOCATIONNAME
                    OutputXML.Append("LOCATION_GROUP_NOT_FOUND")
                Case StatusCodes.INVALID_LOCATIONGROUPID
                    OutputXML.Append("INVALID_LOCATIONGROUPID")
                Case StatusCodes.INVALID_BANNER
                    OutputXML.Append("INVALID_BANNER")
                Case StatusCodes.INVALID_INCOMPLETEPROCESSLOCDATA
                    OutputXML.Append("INVALID_INCOMPLETEPROCESSLOCDATA")
                Case StatusCodes.INVALID_EXTPRODGROUPID
                    OutputXML.Append("INVALID_EXTPRODGROUPID")
                Case StatusCodes.INVALID_PRODGROUPNAME
                    OutputXML.Append("INVALID_PRODGROUPNAME")
                Case StatusCodes.INVALID_PRODUCTTYPEID
                    OutputXML.Append("INVALID_PRODUCTTYPEID")
                Case StatusCodes.INVALID_PRODGROUPID
                    OutputXML.Append("INVALID_PRODGROUPID")
                Case StatusCodes.INVALID_PRODUCTID
                    OutputXML.Append("INVALID_PRODUCTID")
                Case StatusCodes.INVALID_INCOMPLETEPROCESSPRODDATA
                    OutputXML.Append("INVALID_INCOMPLETEPROCESSPRODDATA")
                Case StatusCodes.FAILED_OPTIN
                    OutputXML.Append("FAILED_OPTIN")
                Case StatusCodes.INVALID_DESCRIPTION
                    OutputXML.Append("INVALID_DESCRIPTION")
                Case Else
                    ' treat everything else as an application exception
                    OutputXML.Append("APPLICATION_EXCEPTION")
            End Select
            OutputXML.Append("""><ErrorMessage>" & Escape(ErrorMessage) & "</ErrorMessage>")
        Else
            OutputXML.Append(">")
        End If

        OutputXML.Append("" & InputRespXML)
        OutputXML.Append("</" & RootNode & ">")
        Return OutputXML.ToString
    End Function

    Private Function GetInputResponseXML(ByVal Code As StatusCodes, ByVal Description As String, ByVal ByExtID As Boolean, ByVal node As XmlNode) As String
        Dim OutputXML As New StringBuilder("")
        Dim RootNode As String
        If ByExtID Then
            RootNode = "ImportByExternalID"
        Else
            RootNode = "ImportByLogixID"
        End If
        OutputXML.Append("<" & RootNode & " statuscode=""")
        Select Case Code
            Case StatusCodes.SUCCESS
                OutputXML.Append("SUCCESS")
            Case StatusCodes.INVALID_GUID
                OutputXML.Append("INVALID_GUID")
            Case StatusCodes.INVALID_CUSTOMERGROUPID
                OutputXML.Append("INVALID_CUSTOMERGROUPID")
            Case StatusCodes.INVALID_XML_DOCUMENT
                OutputXML.Append("INVALID_XML_DOCUMENT")
            Case StatusCodes.INVALID_CRITERIA_XML
                OutputXML.Append("INVALID_CRITERIA_XML")
            Case StatusCodes.INVALID_OPERATIONTYPE
                OutputXML.Append("INVALID_OPERATIONTYPE")
            Case StatusCodes.INVALID_CUSTOMERGROUPNAME
                OutputXML.Append("INVALID_CUSTOMERGROUPNAME")
            Case StatusCodes.INVALID_INCOMPLETEPROCESSCUSTDATA
                OutputXML.Append("INVALID_INCOMPLETEPROCESSCUSTDATA")
            Case StatusCodes.INVALID_CARDID
                OutputXML.Append("INVALID_CARDID")
            Case StatusCodes.INVALID_CARDTYPEID
                OutputXML.Append("INVALID_CARDTYPEID")
            Case StatusCodes.INVALID_LOCATIONNAME
                OutputXML.Append("LOCATION_GROUP_NOT_FOUND")
            Case StatusCodes.INVALID_LOCATIONGROUPID
                OutputXML.Append("INVALID_LOCATIONGROUPID")
            Case StatusCodes.INVALID_BANNER
                OutputXML.Append("INVALID_BANNER")
            Case StatusCodes.INVALID_INCOMPLETEPROCESSLOCDATA
                OutputXML.Append("INVALID_INCOMPLETEPROCESSLOCDATA")
            Case StatusCodes.INVALID_EXTPRODGROUPID
                OutputXML.Append("INVALID_EXTPRODGROUPID")
            Case StatusCodes.INVALID_PRODGROUPNAME
                OutputXML.Append("INVALID_PRODGROUPNAME")
            Case StatusCodes.INVALID_PRODUCTTYPEID
                OutputXML.Append("INVALID_PRODUCTTYPEID")
            Case StatusCodes.INVALID_PRODGROUPID
                OutputXML.Append("INVALID_PRODGROUPID")
            Case StatusCodes.INVALID_PRODUCTID
                OutputXML.Append("INVALID_PRODUCTID")
            Case StatusCodes.INVALID_INCOMPLETEPROCESSPRODDATA
                OutputXML.Append("INVALID_INCOMPLETEPROCESSPRODDATA")
            Case StatusCodes.FAILED_OPTIN
                OutputXML.Append("FAILED_OPTIN")
            Case StatusCodes.INVALID_DESCRIPTION
                OutputXML.Append("INVALID_DESCRIPTION")
            Case Else
                ' treat everything else as an application exception
                OutputXML.Append("APPLICATION_EXCEPTION")
        End Select
        OutputXML.Append("""  Description = """ & Escape(Description) & """>")
        For Each childNode In node.ChildNodes
            OutputXML.Append("<" & childNode.Name & ">" & childNode.InnerText & "</" & childNode.Name & ">")
        Next
        OutputXML.Append("</" & RootNode & ">")

        Return OutputXML.ToString
    End Function

    Public Sub ProcessBulkCustGroupData(ByVal CustList As String, ByVal CustomerGroupID As Integer, ByVal Operation As String _
                                             , ByRef RetCode As Integer, ByRef RetMsg As String, ByRef RetMsgLog As String)
        Dim LocalCGroupOp As New Copient.CustomerGroupOperations
        Dim LocalCommon As New Copient.CommonInc
        Dim LocalLookup As New Copient.CustomerLookup
        LocalCommon.AppName = "LogixGroupManagement.asmx"
        Dim CustListData() As String = Nothing
        Dim CustomerData() As String = Nothing

        Dim querystr As String
        Dim dst As DataTable
        Dim Status As Integer
        Dim customerParameterList As New List(Of SqlParameter)
        Dim customerParameter As SqlParameter
        Dim cnt As Integer = 0

        Try
            Dim interfaceOption As String = LocalCommon.Fetch_InterfaceOption(11)
            LocalCommon.Open_LogixRT()
            LocalCommon.Open_LogixXS()
            If CustList IsNot Nothing AndAlso CustList.Trim > "" Then
                CustList = CustList.Replace(vbCrLf, vbLf)
                CustList = CustList.Replace(vbLf, vbCrLf)
                CustListData = CustList.Split(ControlChars.CrLf)

                LocalCommon.QueryStr = "IF OBJECT_ID('tempdb..#TempCardPK') IS NOT NULL BEGIN drop table #TempCardPK END "
                LocalCommon.LRT_Execute()

                querystr = "Create Table #TempCardPK  ([TempPK] int PRIMARY KEY IDENTITY," & _
                     "[CustomerPK] BIGINT NOT NULL)"
                For Each customer In CustListData
                    CustomerData = customer.Trim.Split(",")
                    Dim ExtCardID As String = Trim(CustomerData(0))
                    Dim CardTypeID As Integer
                    If CustomerData.Length = 1 Then
                        If ExtCardID.Length = 0 Then
                            Copient.Logger.Write_Log(LGMLogFile, "Skipping blank customer entry", True)
                            Continue For
                        Else
                            RetCode = StatusCodes.INVALID_INCOMPLETEPROCESSCUSTDATA
                            RetMsg = "Failure. Invalid Customer Bulk Data."
                            Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                            Return
                        End If
                    ElseIf CustomerData.Length <> 2 Then
                        RetCode = StatusCodes.INVALID_INCOMPLETEPROCESSCUSTDATA
                        RetMsg = "Failure. Invalid Customer Bulk Data."
                        Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                        Return
                    End If

                    If Integer.TryParse(CustomerData(1), CardTypeID) = False Then
                        RetCode = StatusCodes.INVALID_CARDTYPEID
                        RetMsg = "Failure. Invalid card type: " & CustomerData(1) & " for card: " & ExtCardID
                        RetMsgLog = "Failure. Invalid card type: " & CustomerData(1) & " for card: " & Copient.MaskHelper.MaskCard(ExtCardID, CustomerData(1))
                        Copient.Logger.Write_Log(LGMLogFile, RetMsgLog, True)
                        Return
                    Else
                        LocalCommon.QueryStr = "select CardTypeID from CardTypes with (NoLock) " & _
                                                "where CardTypeID=@CardTypeID and CustTypeID=0;"
                        LocalCommon.DBParameters.Add("@CardTypeID", SqlDbType.Int).Value = CardTypeID
                        dst = LocalCommon.ExecuteQuery(Copient.DataBases.LogixXS)
                        If dst.Rows.Count = 0 Then
                            RetCode = StatusCodes.INVALID_CARDTYPEID
                            RetMsg = "Failure. Invalid card type: " & CustomerData(1) & " for card: " & ExtCardID
                            RetMsgLog = "Failure. Invalid card type: " & CustomerData(1) & " for card: " & Copient.MaskHelper.MaskCard(ExtCardID, CustomerData(1))
                            Copient.Logger.Write_Log(LGMLogFile, RetMsgLog, True)
                            Return
                        End If
                    End If

                    If (LocalCGroupOp.CheckCamCard(CustomerGroupID)) Then
                        CardTypeID = 2
                    End If

                    ExtCardID = LocalCommon.Pad_ExtCardID(ExtCardID, CardTypeID)

                    LocalCommon.QueryStr = "Select CustomerPK from CardIDs where ExtCardID = @ExtCardID and CardTypeID = @CardTypeID"
                    LocalCommon.DBParameters.Add("@ExtCardID", SqlDbType.NVarChar).Value = MyCryptLib.SQL_StringEncrypt(ExtCardID, True)
                    LocalCommon.DBParameters.Add("@CardTypeID", SqlDbType.Int).Value = CardTypeID
                    dst = LocalCommon.ExecuteQuery(Copient.DataBases.LogixXS)

                    If dst.Rows.Count > 0 Then
                        querystr &= "insert into #TempCardPK values(@CustomerPK" & cnt & "); "
                        customerParameter = New SqlParameter("@CustomerPK" & cnt, SqlDbType.BigInt)
                        customerParameter.Value = dst.Rows(0).Item("CustomerPK")
                        customerParameterList.Add(customerParameter)
                    Else
                        If interfaceOption = "0" Then
                            RetCode = StatusCodes.INVALID_CARDID
                            RetMsg = "Failure. Card ID: " & ExtCardID & " of type " & CardTypeID & " does not exist."
                            RetMsgLog = "Failure. Card ID: " & Copient.MaskHelper.MaskCard(ExtCardID, CardTypeID) & " of type " & CardTypeID & " does not exist."
                            Copient.Logger.Write_Log(LGMLogFile, RetMsgLog, True)
                            Return
                        Else
                            'Create New Customer
                            If Operation.ToUpper() <> "REMOVE" Then
                                Dim CustomerPK As Int64
                                Dim NewCustomer As New Copient.Customer
                                LocalCommon.QueryStr = "Select NumericOnly from CardTypes where CardTypeID = @CardTypeID"
                                LocalCommon.DBParameters.Add("@CardTypeID", SqlDbType.Int).Value = CardTypeID
                                dst = LocalCommon.ExecuteQuery(Copient.DataBases.LogixXS)
                                If dst.Rows.Count > 0 Then
                                    If dst.Rows(0).Item("NumericOnly") = True Then
                                        'If Not tryParseInt(ExtCardID) Then

                                        If (Not IsNumeric(ExtCardID)) Then
                                            RetCode = StatusCodes.INVALID_CARDID
                                            RetMsg = "Failure. Could not add card " & ExtCardID & " because it contains non-numeric characters"
                                            RetMsgLog = "Failure. Could not add card " & Copient.MaskHelper.MaskCard(ExtCardID, CardTypeID) & " CardTypeID " & CardTypeID & " because it contains non-numeric characters"
                                            Copient.Logger.Write_Log(LGMLogFile, RetMsgLog, True)
                                            Return
                                        End If

                                        If (CLng(ExtCardID) <= 0) Then
                                            RetCode = StatusCodes.INVALID_CARDID
                                            RetMsg = "Failure. Invalid Card ID: " & ExtCardID
                                            Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                                            Return
                                        End If
                                    End If
                                End If

                                NewCustomer.AddCard(New Copient.Card(ExtCardID, CardTypeID))
                                LocalLookup.AddCustomer(NewCustomer, RetCode)
                                If RetCode = StatusCodes.SUCCESS Then
                                    CustomerPK = LocalLookup.FindCustomerPKFromExtID(ExtCardID, CardTypeID)
                                    Copient.Logger.Write_Log(LGMLogFile, "Created Customer " & Copient.MaskHelper.MaskCard(ExtCardID, CardTypeID) & " of type " & CardTypeID & " was created (PK=" & CustomerPK & ")", True)

                                    'Add to table
                                    querystr &= "insert into #TempCardPK values(@CustomerPK" & cnt & "); "
                                    customerParameter = New SqlParameter("@CustomerPK" & cnt, SqlDbType.BigInt)
                                    customerParameter.Value = CustomerPK
                                    customerParameterList.Add(customerParameter)
                                ElseIf RetCode = StatusCodes.INVALID_CARDID Then
                                    RetMsg = "Card failed Check Digit validation for Card Id: " & " (" & ExtCardID & ")"
                                    RetMsgLog = "Card failed Check Digit validation for Card Id: " & " (" & Copient.MaskHelper.MaskCard(ExtCardID, CardTypeID) & ") CardTypeID: (" & CardTypeID & ")"
                                    Copient.Logger.Write_Log(LGMLogFile, RetMsgLog, True)
                                    Return
                                Else
                                    RetCode = StatusCodes.APPLICATION_EXCEPTION
                                    RetMsg = "Could not create a new customer " & ExtCardID & " of type " & CardTypeID
                                    RetMsgLog = "Could not create a new customer " & Copient.MaskHelper.MaskCard(ExtCardID, CardTypeID) & " of type " & CardTypeID
                                    Copient.Logger.Write_Log(LGMLogFile, RetMsgLog, True)
                                    Return
                                End If
                            Else
                                RetCode = StatusCodes.INVALID_CARDID
                                RetMsg = "Failure. Cannot remove card: " & ExtCardID & " of type " & CardTypeID & " because it does not exist"
                                RetMsgLog = "Failure. Cannot remove card: " & Copient.MaskHelper.MaskCard(ExtCardID, CardTypeID) & " of type " & CardTypeID & " because it does not exist"
                                Copient.Logger.Write_Log(LGMLogFile, RetMsgLog, True)
                                Return
                            End If

                        End If
                    End If
                    cnt = cnt + 1
                Next
                For Each sqlParam In customerParameterList
                    LocalCommon.DBParameters.Add(sqlParam)
                Next
                LocalCommon.QueryStr = querystr
                LocalCommon.QueryStr &= "declare @Status int = 0; "
                If Operation = "augment" Then
                    LocalCommon.QueryStr &= "exec pt_LGM_CustomerGroup_Insert  @CustomerGroupID = @CustomerGroupID, @Status = @Status OUTPUT; Select @Status 'Status';"
                    LocalCommon.QueryStr &= " drop table #TempCardPK;"
                    LocalCommon.DBParameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = CustomerGroupID
                    dst = LocalCommon.ExecuteQuery(Copient.DataBases.LogixXS)
                    If dst.Rows.Count > 0 Then
                        Status = dst.Rows(0).Item("Status")
                        If Status = 0 Or Status = -1 Then
                            RetCode = StatusCodes.SUCCESS
                            RetMsg = "Cards Successfully added to Customer Group " & CustomerGroupID
                            RetMsgLog = RetMsg
                        ElseIf Status = -2 Then
                            'Since the cards are checked earlier, it should never get to this point. 
                            RetCode = StatusCodes.INVALID_INCOMPLETEPROCESSCUSTDATA
                            RetMsg = "Failure. Tried to add a customer that does not exist. To the CustomerGroup. No cards were added"
                            RetMsgLog = RetMsg
                        End If
                    End If
                ElseIf Operation = "replace" Then
                    LocalCommon.QueryStr &= "exec pt_LGM_CustomerGroup_Replace  @CustomerGroupID = @CustomerGroupID, @Status = @Status OUTPUT; Select @Status 'Status';"
                    LocalCommon.QueryStr &= " drop table #TempCardPK;"
                    LocalCommon.DBParameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = CustomerGroupID
                    dst = LocalCommon.ExecuteQuery(Copient.DataBases.LogixXS)
                    If dst.Rows.Count > 0 Then
                        Status = dst.Rows(0).Item("Status")
                        If Status = 0 Then
                            RetCode = StatusCodes.SUCCESS
                            RetMsg = "Successfully replaced cards in Customer Group " & CustomerGroupID
                            RetMsgLog = RetMsg
                        ElseIf Status = -2 Then
                            'Since the cards are checked earlier, it should never get to this point. 
                            RetCode = StatusCodes.INVALID_INCOMPLETEPROCESSCUSTDATA
                            RetMsg = "Failure. Tried to add a customer that does not exist. To the CustomerGroup. No cards were replaced"
                            RetMsgLog = RetMsg
                        End If
                    End If
                ElseIf Operation = "remove" Then
                    LocalCommon.QueryStr &= "exec pt_LGM_CustomerGroup_Remove @CustomerGroupID = @CustomerGroupID, @Status = @Status OUTPUT; Select @Status 'Status';"
                    LocalCommon.QueryStr &= " drop table #TempCardPK;"
                    LocalCommon.DBParameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = CustomerGroupID
                    dst = LocalCommon.ExecuteQuery(Copient.DataBases.LogixXS)
                    If dst.Rows.Count > 0 Then
                        Status = dst.Rows(0).Item("Status")
                        If Status = 0 Then
                            RetCode = StatusCodes.SUCCESS
                            RetMsg = "Cards Successfully removed from Customer Group " & CustomerGroupID
                            RetMsgLog = RetMsg
                        ElseIf Status = -1 Then
                            'Since the cards are checked earlier, it should never get to this point. 
                            RetCode = StatusCodes.INVALID_INCOMPLETEPROCESSCUSTDATA
                            RetMsg = "Failure. Tried to remove a customer that does not exist. No cards were removed"
                            RetMsgLog = RetMsg
                        End If
                    End If
                End If
            Else
                RetCode = StatusCodes.INVALID_INCOMPLETEPROCESSCUSTDATA
                RetMsg = "Failure. BulkData is empty"
                RetMsgLog = RetMsg
            End If
        Catch ex As Exception
            RetCode = StatusCodes.APPLICATION_EXCEPTION
            RetMsg = ex.ToString()
            RetMsgLog = RetMsg
        Finally
            LocalCommon.QueryStr = "IF OBJECT_ID('tempdb..#TempCardPK') IS NOT NULL BEGIN drop table #TempCardPK END "
            LocalCommon.LRT_Execute()
            LocalCommon.Close_LogixRT()
            LocalCommon.Close_LogixXS()
        End Try
        Copient.Logger.Write_Log(LGMLogFile, RetMsgLog, True)
    End Sub


#End Region

    Private Function isValidInputText(ByVal pvalue As String, Optional ByVal pvalueMandatory As Boolean = True) As Boolean
        Dim blnResult As Boolean = False
        Try
            If pvalueMandatory = True Then
                If pvalue = "" Then
                    blnResult = False
                End If
            Else
                blnResult = True
            End If
            If pvalue <> "" Then
                If pvalue.Contains("'") = True OrElse pvalue.Contains(Chr(34)) = True Then
                    blnResult = False
                Else
                    blnResult = True
                End If
            End If
        Catch ex As Exception
            blnResult = False
        End Try
        Return blnResult
    End Function

#Region "SingleLocationGroup"

    Private Function _ProcessLocationInLocationGroup(ByVal GUID As String, ByVal EXTGroupID As String, ByVal GroupName As String, ByVal Description As String, ByVal ExtLocationID As String, ByVal LocationName As String, ByVal ExtBannerID As String, ByVal operation As String) As DataSet
        Dim LocationGroupId As Long
        Dim Name As String = ""
        Dim StoreList As String = ""
        Dim OfferId As Long = -1
        Dim dtStatus As DataTable
        Dim row As DataRow
        Dim dt As DataTable
        Dim BannersEnabled As Boolean = False
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim BannerID As Long
        Dim EngineType As Integer
        Dim ErrorMessage As String = ""
        Dim resultset As New DataSet

        EngineType = GetInstalledEngine(LGMLogFile)
        StoreList = ExtLocationID & "," & LocationName


        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        If isValidInputText(GUID) = False Then
            'GUID Value is empty
            RetCode = StatusCodes.INVALID_GUID
            RetMsg = "Failure. Invalid GUID"
            row = dtStatus.NewRow()
            row.Item("StatusCode") = RetCode
            row.Item("Description") = RetMsg
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            resultset.Tables.Add(dtStatus.Copy())
            Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
            _ProcessLocationInLocationGroup = resultset
            Exit Function
        End If
        If isValidInputText(EXTGroupID) = False Then
            'Card ID Value is empty or invalid
            RetCode = StatusCodes.INVALID_LOCATIONGROUPID
            RetMsg = "Failure. Invalid Location GroupID"
            row = dtStatus.NewRow()
            row.Item("StatusCode") = RetCode
            row.Item("Description") = RetMsg
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            resultset.Tables.Add(dtStatus.Copy())
            Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
            _ProcessLocationInLocationGroup = resultset
            Exit Function
        End If 'ExtCardId
        If isValidInputText(ExtLocationID) = False Then
            'Card ID Value is empty or invalid
            RetCode = StatusCodes.INVALID_LOCATIONCODE
            RetMsg = "Failure. Invalid Location Code"
            row = dtStatus.NewRow()
            row.Item("StatusCode") = RetCode
            row.Item("Description") = RetMsg
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            resultset.Tables.Add(dtStatus.Copy())
            Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
            _ProcessLocationInLocationGroup = resultset
            Exit Function
        End If 'ExtCardId
        If isValidInputText(operation) = False Then
            RetCode = StatusCodes.INVALID_OPERATIONTYPE
            RetMsg = "Failure. Invalid Operation"
            row = dtStatus.NewRow()
            row.Item("StatusCode") = RetCode
            row.Item("Description") = RetMsg
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            resultset.Tables.Add(dtStatus.Copy())
            Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
            _ProcessLocationInLocationGroup = resultset
            Exit Function
        End If
        If isValidInputText(GroupName, False) = False Then
            RetCode = StatusCodes.INVALID_CUSTOMERGROUPNAME
            RetMsg = "Failure. Invalid Location Group Name"
            row = dtStatus.NewRow()
            row.Item("StatusCode") = RetCode
            row.Item("Description") = RetMsg
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            resultset.Tables.Add(dtStatus.Copy())
            Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
            _ProcessLocationInLocationGroup = resultset
            Exit Function
        End If
        operation = operation.ToLower
        If operation <> "augment" And operation <> "remove" And operation <> "replace" Then
            RetCode = StatusCodes.INVALID_OPERATIONTYPE
            RetMsg = "Failure. Invalid Operation"
            row = dtStatus.NewRow()
            row.Item("StatusCode") = RetCode
            row.Item("Description") = RetMsg
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            resultset.Tables.Add(dtStatus.Copy())
            Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
            _ProcessLocationInLocationGroup = resultset
            Exit Function
        End If
        Try
            ''''for opening the database connections
            Using MyCommon.LRTadoConn
                If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
                BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")


                If Not IsValidGUID(GUID, "ProcessLocationInLocationGroupByExtID") Then
                    RetCode = StatusCodes.INVALID_GUID
                    RetMsg = "Failure. Invalid GUID."
                ElseIf BannersEnabled = True Then
                    If isValidInputText(ExtBannerID) = False Then
                        RetCode = StatusCodes.INVALID_BANNER
                        RetMsg = "Failure. Banners are enabled but location group was not sent with EXTBannerID - can not process the group"
                    Else
                        MyCommon.QueryStr = "SELECT BannerID FROM Banners with (NoLock) WHERE EXTBannerID = @ExtBannerID AND Deleted=0"
                        MyCommon.DBParameters.Add("@ExtBannerID", SqlDbType.NVarChar).Value = ExtBannerID
                        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                        If (dt.Rows.Count > 0) Then
                            BannerID = MyCommon.NZ(dt.Rows(0).Item("BannerID"), 0)
                            If BannerID > 0 Then
                                MyCommon.QueryStr = "select EngineID from BannerEngines BE with (NoLock) where BannerID=@BannerID"
                                MyCommon.DBParameters.Add("@BannerID", SqlDbType.Int).Value = BannerID
                                dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                                If (dt.Rows.Count > 0) Then
                                    EngineType = MyCommon.NZ(dt.Rows(0).Item("EngineID"), EngineType)
                                End If
                            End If
                        Else
                            RetCode = StatusCodes.INVALID_BANNER
                            RetMsg = "Failure. Invalid BannerID."
                        End If
                    End If
                End If
                'for all input validations

                If RetCode = StatusCodes.SUCCESS Then
                    If MyLocGroupOp.LocationGroupExists(EXTGroupID, LocationGroupId, Name) Then
                        If Not (GroupName = "") AndAlso Not (GroupName = Name) Then
                            MyLocGroupOp.LocationGroupNameExists(LGMLogFile, LocationGroupId, GroupName, RetCode, RetMsg)
                            If RetMsg <> "" Then
                                MyLocGroupOp.UpdateLocationGroupDescription(LGMLogFile, LocationGroupId, Description)
                            End If
                        End If
                    Else
                        If Not (GroupName = "") Then
                            MyLocGroupOp.LocationGroupNameExists(LGMLogFile, LocationGroupId, GroupName, RetCode, RetMsg)
                        Else
                            RetCode = StatusCodes.INVALID_LOCATIONNAME
                            RetMsg = "Failure. GroupName is not provided"
                        End If
                        If RetCode = StatusCodes.SUCCESS Then
                            If operation.ToUpper <> "REMOVE" Then
                                LocationGroupId = MyLocGroupOp.CreateLocationGroupByExtID(EXTGroupID, BannerID, Description, ErrorMessage, EngineType, GroupName)
                            End If
                        End If
                    End If
                End If
                If RetCode = StatusCodes.SUCCESS Then
                    ProcessLocationGroup(LocationGroupId, StoreList, BannerID, ErrorMessage, operation, RetCode, RetMsg)
                Else
                    Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                End If
            End Using
        Catch ex As Exception
            RetCode = StatusCodes.APPLICATION_EXCEPTION
            RetMsg = "Failure. Application " & ex.ToString
        Finally
            If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count = 0 Then
                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                resultset.Tables.Add(dtStatus.Copy())
            End If
        End Try
        Return resultset
    End Function

    Private Function _ProcessLocationInLocationGroup(ByVal GUID As String, ByVal LocationGroupID As Long, ByVal GroupName As String, ByVal Description As String, ByVal ExtLocationID As String, ByVal LocationName As String, ByVal ExtBannerID As String, ByVal operation As String) As DataSet
        'Dim LocationGroupId As Long
        Dim Name As String = ""
        Dim StoreList As String = ""
        Dim OfferId As Long = -1
        Dim dtStatus As DataTable
        Dim row As DataRow
        Dim dt As DataTable
        Dim BannersEnabled As Boolean = False
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim BannerID As Long
        Dim EngineType As Integer
        Dim ErrorMessage As String = ""
        Dim resultset As New DataSet

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        If isValidInputText(GUID) = False Then
            'GUID Value is empty
            RetCode = StatusCodes.INVALID_GUID
            RetMsg = "Failure. Invalid GUID"
            row = dtStatus.NewRow()
            row.Item("StatusCode") = RetCode
            row.Item("Description") = RetMsg
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            resultset.Tables.Add(dtStatus.Copy())
            Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
            _ProcessLocationInLocationGroup = resultset
            Exit Function
        End If
        If LocationGroupID = -1 Then
            RetCode = StatusCodes.INVALID_LOCATIONGROUPID
            RetMsg = "Failure. Invalid Location GroupID"
            row = dtStatus.NewRow()
            row.Item("StatusCode") = RetCode
            row.Item("Description") = RetMsg
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            resultset.Tables.Add(dtStatus.Copy())
            Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
            _ProcessLocationInLocationGroup = resultset
            Exit Function
        End If
        If isValidInputText(ExtLocationID) = False Then
            'Card ID Value is empty or invalid
            RetCode = StatusCodes.INVALID_LOCATIONCODE
            RetMsg = "Failure. Invalid Location Code"
            row = dtStatus.NewRow()
            row.Item("StatusCode") = RetCode
            row.Item("Description") = RetMsg
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            resultset.Tables.Add(dtStatus.Copy())
            Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
            _ProcessLocationInLocationGroup = resultset
            Exit Function
        End If 'ExtCardId
        If isValidInputText(operation) = False Then
            RetCode = StatusCodes.INVALID_OPERATIONTYPE
            RetMsg = "Failure. Invalid Operation"
            row = dtStatus.NewRow()
            row.Item("StatusCode") = RetCode
            row.Item("Description") = RetMsg
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            resultset.Tables.Add(dtStatus.Copy())
            Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
            _ProcessLocationInLocationGroup = resultset
            Exit Function
        End If
        If isValidInputText(GroupName, False) = False Then
            RetCode = StatusCodes.INVALID_CUSTOMERGROUPNAME
            RetMsg = "Failure. Invalid Location Group Name"
            row = dtStatus.NewRow()
            row.Item("StatusCode") = RetCode
            row.Item("Description") = RetMsg
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            resultset.Tables.Add(dtStatus.Copy())
            Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
            _ProcessLocationInLocationGroup = resultset
            Exit Function
        End If
        operation = operation.ToLower
        If operation <> "augment" AndAlso operation <> "remove" AndAlso operation <> "replace" Then
            RetCode = StatusCodes.INVALID_OPERATIONTYPE
            RetMsg = "Failure. Invalid Operation"
            row = dtStatus.NewRow()
            row.Item("StatusCode") = RetCode
            row.Item("Description") = RetMsg
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            resultset.Tables.Add(dtStatus.Copy())
            Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
            _ProcessLocationInLocationGroup = resultset
            Exit Function
        End If
        Try
            ''''for opening the database connections
            Using MyCommon.LRTadoConn
                If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
                EngineType = GetInstalledEngine(LGMLogFile)
                StoreList = ExtLocationID & "," & LocationName
                BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
                If Not IsValidGUID(GUID, "ProcessLocationInLocationGroupByExtID") Then
                    RetCode = StatusCodes.INVALID_GUID
                    RetMsg = "Failure. Invalid GUID."
                ElseIf BannersEnabled = True Then
                    If isValidInputText(ExtBannerID) = False Then
                        RetCode = StatusCodes.INVALID_BANNER
                        RetMsg = "Failure. Banners are not enabled but location group was sent with EXTBannerID - can not process the group"
                    Else
                        MyCommon.QueryStr = "SELECT BannerID FROM Banners with (NoLock) WHERE EXTBannerID = @ExtBannerID AND Deleted=0"
                        MyCommon.DBParameters.Add("@ExtBannerID", SqlDbType.NVarChar).Value = ExtBannerID
                        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                        If (dt.Rows.Count > 0) Then
                            BannerID = MyCommon.NZ(dt.Rows(0).Item("BannerID"), 0)
                            If BannerID > 0 Then
                                MyCommon.QueryStr = "select EngineID from BannerEngines BE with (NoLock) where BannerID=@BannerID"
                                MyCommon.DBParameters.Add("@BannerID", SqlDbType.Int).Value = BannerID
                                dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                                If (dt.Rows.Count > 0) Then
                                    EngineType = MyCommon.NZ(dt.Rows(0).Item("EngineID"), EngineType)
                                End If
                            End If
                        Else
                            RetCode = StatusCodes.INVALID_BANNER
                            RetMsg = "Failure. Invalid BannerID."
                        End If
                    End If
                End If

                If RetCode = StatusCodes.SUCCESS Then
                    If MyLocGroupOp.LocationGroupExists(LocationGroupID, Name) Then
                        If Not (GroupName = "") AndAlso Not (GroupName = Name) Then
                            MyLocGroupOp.LocationGroupNameExists(LGMLogFile, LocationGroupID, GroupName, RetCode, RetMsg)
                            If RetMsg <> "" Then
                                MyLocGroupOp.UpdateLocationGroupDescription(LGMLogFile, LocationGroupID, Description)
                            End If
                        End If
                    Else
                        If Not (GroupName = "") Then
                            MyLocGroupOp.LocationGroupNameExists(LGMLogFile, LocationGroupID, GroupName, RetCode, RetMsg)
                        Else
                            RetCode = StatusCodes.INVALID_LOCATIONNAME
                            RetMsg = "Failure. GroupName is not provided"
                        End If
                        If RetCode = StatusCodes.SUCCESS Then
                            If operation.ToUpper <> "REMOVE" Then
                                LocationGroupID = MyLocGroupOp.CreateLocationGroupByLogixID(BannerID, Description, ErrorMessage, EngineType, GroupName, RetCode, RetMsg)
                            End If
                        End If
                    End If
                End If

                If RetCode = StatusCodes.SUCCESS Then
                    ProcessLocationGroup(LocationGroupID, StoreList, BannerID, ErrorMessage, operation, RetCode, RetMsg)
                Else
                    Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                End If
            End Using
        Catch ex As Exception
            RetCode = StatusCodes.APPLICATION_EXCEPTION
            RetMsg = "Failure. Application " & ex.ToString
        Finally
            If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count = 0 Then
                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                resultset.Tables.Add(dtStatus.Copy())
            End If
        End Try
        Return resultset
    End Function

    Private Function GetInstalledEngine(ByVal sLogFile As String) As Integer

        Dim dst As DataTable
        Dim EngineType As Integer
        Try
            'Using MyCommon.LRTadoConn
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            MyCommon.QueryStr = "Select EngineID from PromoEngines with (NoLock) where Installed = 1 and DefaultEngine = 1;"
            dst = MyCommon.LRT_Select()
            If (dst.Rows.Count > 0) Then
                EngineType = dst.Rows(0).Item("EngineID")
            End If
            'End Using
        Catch ex As Exception
            Copient.Logger.Write_Log(sLogFile, "" & ex.Message & "", True)
        End Try

        Return EngineType

    End Function

    Private Sub ProcessLocationGroup(ByVal LocationGroupID As Long, ByVal StoreList As String, ByVal BannerID As Long, _
                                         ByRef ErrorMessage As String, ByVal Operation As String, _
                                         ByRef RetCode As Integer, ByRef RetMsg As String)
        Dim Stores() As String = Nothing
        Dim StoreData() As String = Nothing
        Dim querystr As String
        Dim dst As DataTable
        Dim CurrencyID As Integer = 0
        Dim Status As Integer
        Dim locationParameter As SqlParameter
        Dim locationParameterList As New List(Of SqlParameter)

        Try
            ' Using MyCommon.LRTadoConn
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If StoreList IsNot Nothing AndAlso StoreList.Trim > "" Then
                StoreList = StoreList.Replace(vbCrLf, vbLf)
                StoreList = StoreList.Replace(vbLf, vbCrLf)
                Stores = StoreList.Split(ControlChars.CrLf)
                MyCommon.QueryStr = "IF OBJECT_ID('tempdb..#TempLocPK') IS NOT NULL BEGIN drop table #TempLocPK END "
                MyCommon.LRT_Execute()
                querystr = "Create Table #TempLocPK  ([TempPK] int PRIMARY KEY IDENTITY," & _
                    "[LocationID] BIGINT NOT NULL)"
                For Each store In Stores
                    StoreData = store.Trim.Split(",")
                    Dim StoreExtID As String = Trim(StoreData(0))
                    Dim StoreName As String
                    If StoreData.Length = 1 Then
                        If StoreExtID.Length = 0 Then
                            Copient.Logger.Write_Log(LGMLogFile, "Skipping blank location entry", True)
                            Continue For
                        Else
                            RetCode = StatusCodes.INVALID_INCOMPLETEPROCESSLOCDATA
                            RetMsg = "Failure. Invalid Location Bulk Data."
                            Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                            Return
                        End If
                    ElseIf StoreData.Length <> 2 Then
                        RetCode = StatusCodes.INVALID_INCOMPLETEPROCESSLOCDATA
                        RetMsg = "Failure. Invalid Location Bulk Data."
                        Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                        Return
                    End If

                    StoreName = Trim(StoreData(1))


                    If StoreExtID <> "" Then
                        ' determine if the location already exists
                        MyCommon.QueryStr = "select LocationID from Locations with (NoLock) where Deleted=0 and (ExtLocationCode = @StoreExtID)"
                        MyCommon.DBParameters.Add("@StoreExtID", SqlDbType.NVarChar).Value = StoreExtID
                    Else
                        RetCode = StatusCodes.INVALID_INCOMPLETEPROCESSLOCDATA
                        RetMsg = "Failure. Invalid Location Bulk Data."
                        Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                        Return
                    End If

                    dst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)

                    If dst.Rows.Count > 0 Then
                        querystr &= "insert into #TempLocPK values(@LocationID" & StoreExtID & "); "
                        'MyCommon.DBParameters.Add("@LocationID",SqlDbType.BigInt).Value = MyCommon.NZ(dst.Rows(0).Item("LocationID"), 0)
                        locationParameter = New SqlParameter("@LocationID" & StoreExtID, SqlDbType.BigInt)
                        locationParameter.Value = MyCommon.NZ(dst.Rows(0).Item("LocationID"), 0)
                        locationParameterList.Add(locationParameter)

                    Else

                        If MyCommon.Fetch_InterfaceOption(84) = "0" Then
                            RetCode = StatusCodes.INVALID_LOCATIONCODE
                            RetMsg = "Failure. Store: " & StoreName & " with code " & StoreExtID & " does not exist."
                            Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                            Return
                        Else

                            'Create New Location
                            If Operation.ToUpper <> "REMOVE" Then
                                If StoreExtID <> "" Then
                                    MyCommon.QueryStr = "select LocationID from Locations with (NoLock) where Deleted=0 and (ExtLocationCode =@StoreExtID)"
                                    MyCommon.DBParameters.Add("@StoreExtID", SqlDbType.NVarChar).Value = StoreExtID

                                    dst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                                    If dst.Rows.Count > 0 Then
                                        RetCode = StatusCodes.INVALID_LOCATIONNAME
                                        RetMsg = "Failure. Cannot create a new store with name " & StoreName & " and ExtID of " & StoreExtID & " because a store already exists with that name or id"
                                        Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                                        Return
                                    End If

                                    If CurrencyID = 0 Then
                                        MyCommon.QueryStr = "select case isnumeric(OptionValue) when 1 then cast(OptionValue as int) else 1 end AS DefaultCurrencyID from UE_SystemOptions where OptionID=137;"
                                        dst = MyCommon.LRT_Select()
                                        If (dst.Rows.Count = 1) Then
                                            For Each row In dst.Rows
                                                CurrencyID = row.Item("DefaultCurrencyID")
                                            Next
                                        End If
                                    End If
                                    Dim LocationID As Int64
                                    MyCommon.QueryStr = "dbo.pt_Locations_Insert"
                                    MyCommon.Open_LRTsp()
                                    MyCommon.LRTsp.Parameters.Add("@ExtLocationCode", SqlDbType.NVarChar, 20).Value = StoreExtID
                                    MyCommon.LRTsp.Parameters.Add("@LocationName", SqlDbType.NVarChar, 100).Value = StoreName
                                    MyCommon.LRTsp.Parameters.Add("@LocationTypeID", SqlDbType.Int).Value = 1
                                    MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 4000).Value = ""
                                    MyCommon.LRTsp.Parameters.Add("@Address1", SqlDbType.NVarChar, 200).Value = ""
                                    MyCommon.LRTsp.Parameters.Add("@Address2", SqlDbType.NVarChar, 200).Value = ""
                                    MyCommon.LRTsp.Parameters.Add("@City", SqlDbType.NVarChar, 100).Value = ""
                                    MyCommon.LRTsp.Parameters.Add("@State", SqlDbType.NVarChar, 50).Value = ""
                                    MyCommon.LRTsp.Parameters.Add("@Zip", SqlDbType.NVarChar, 20).Value = ""
                                    MyCommon.LRTsp.Parameters.Add("@CountryID", SqlDbType.Int).Value = 1
                                    MyCommon.LRTsp.Parameters.Add("@ContactName", SqlDbType.NVarChar, 200).Value = ""
                                    MyCommon.LRTsp.Parameters.Add("@PhoneNumber", SqlDbType.NVarChar, 40).Value = ""
                                    MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = GetInstalledEngine(LGMLogFile)
                                    MyCommon.LRTsp.Parameters.Add("@TimeZone", SqlDbType.NVarChar, 20).Value = ""
                                    MyCommon.LRTsp.Parameters.Add("@CurrencyID", SqlDbType.Int).Value = CurrencyID
                                    MyCommon.LRTsp.Parameters.Add("@UOMSetID", SqlDbType.Int).Value = 0
                                    MyCommon.LRTsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Direction = ParameterDirection.Output

                                    MyCommon.LRTsp.ExecuteNonQuery()

                                    If (Not IsDBNull(MyCommon.LRTsp.Parameters("@LocationID").Value)) Then
                                        LocationID = MyCommon.LRTsp.Parameters("@LocationID").Value
                                        ' if applicable, assign the first pre-set banners to this location
                                        If (MyCommon.Fetch_SystemOption(66) = "1" AndAlso LocationID > 0 AndAlso BannerID > 0) Then
                                            MyCommon.QueryStr = "update Locations with (RowLock) set BannerID = @BannerID " & _
                                                 "where LocationID = @LocationID and (BannerID is null or BannerID = 0)"
                                            MyCommon.DBParameters.Add("@BannerID", SqlDbType.Int).Value = BannerID
                                            MyCommon.DBParameters.Add("@LocationID", SqlDbType.BigInt).Value = LocationID
                                            MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
                                        End If
                                    End If
                                    MyCommon.Close_LRTsp()
                                    Copient.Logger.Write_Log(LGMLogFile, "Created store with name: " & StoreName & " and Code: " & StoreExtID, True)
                                    'Add to table
                                    querystr &= "insert into #TempLocPK values(@LocationID" & StoreExtID & "); "
                                    locationParameter = New SqlParameter("@LocationID" & StoreExtID, SqlDbType.BigInt)
                                    locationParameter.Value = LocationID
                                    locationParameterList.Add(locationParameter)
                                ElseIf StoreExtID = "" Then
                                    RetCode = StatusCodes.INVALID_LOCATIONCODE
                                    RetMsg = "Failure. Cannot create a new store because the ExtID is blank"
                                    Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                                    Return
                                Else
                                    RetCode = StatusCodes.INVALID_LOCATIONNAME
                                    RetMsg = "Failure. Cannot create a new store because the store name is blank"
                                    Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                                    Return
                                End If
                            Else
                                RetCode = StatusCodes.INVALID_LOCATIONCODE
                                RetMsg = "Failure. Cannot remove store: " & StoreName & " because it does not exist"
                                Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                                Return
                            End If
                        End If
                    End If

                Next
                For Each param In locationParameterList
                    MyCommon.DBParameters.Add(param)
                Next
                MyCommon.QueryStr = querystr
                MyCommon.QueryStr &= "declare @Status int = 0;"
                If Operation = "augment" Then
                    MyCommon.QueryStr &= "exec pt_LGM_LocationGroup_Insert @LocationGroupID = @LocationGroupID, @Status = @Status OUTPUT; Select @Status 'Status';"
                    MyCommon.QueryStr &= " drop table #TempLocPK;"
                    MyCommon.DBParameters.Add("@LocationGroupID", SqlDbType.BigInt).Value = LocationGroupID
                    dst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                    If dst.Rows.Count > 0 Then
                        Status = dst.Rows(0).Item("Status")
                        If Status = 0 Or Status = -1 Then
                            RetCode = StatusCodes.SUCCESS
                            RetMsg = "Stores Successfully added to Location Group " & LocationGroupID
                        ElseIf Status = -2 Then
                            'Since the cards are checked earlier, it should never get to this point. 
                            RetCode = StatusCodes.INVALID_INCOMPLETEPROCESSLOCDATA
                            RetMsg = "Failure. Tried to add a store that does not exist. To the LocationGroup. No stores were added"
                        End If
                    End If
                ElseIf Operation.ToLower = "replace" Then
                    MyCommon.QueryStr &= "exec pt_LGM_LocationGroup_Replace  @LocationGroupID = @LocationGroupID, @Status = @Status OUTPUT; Select @Status 'Status';"
                    MyCommon.QueryStr &= " drop table #TempLocPK;"
                    MyCommon.DBParameters.Add("@LocationGroupID", SqlDbType.BigInt).Value = LocationGroupID
                    dst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                    If dst.Rows.Count > 0 Then
                        Status = dst.Rows(0).Item("Status")
                        If Status = 0 Or Status = -1 Then
                            RetCode = StatusCodes.SUCCESS
                            RetMsg = "Stores Successfully added to Location Group " & LocationGroupID
                        ElseIf Status = -2 Then
                            'Since the cards are checked earlier, it should never get to this point. 
                            RetCode = StatusCodes.INVALID_INCOMPLETEPROCESSLOCDATA
                            RetMsg = "Failure. Tried to add a store that does not exist. To the LocationGroup. No stores were added"
                        End If
                    End If
                ElseIf Operation.ToLower = "remove" Then
                    MyCommon.QueryStr &= "exec pt_LGM_LocationGroup_Remove  @LocationGroupID = @LocationGroupID, @Status = @Status OUTPUT; Select @Status 'Status';"
                    MyCommon.QueryStr &= " drop table #TempLocPK;"
                    MyCommon.DBParameters.Add("@LocationGroupID", SqlDbType.BigInt).Value = LocationGroupID
                    dst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                    If dst.Rows.Count > 0 Then
                        Status = dst.Rows(0).Item("Status")
                        If Status = 0 Then
                            RetCode = StatusCodes.SUCCESS
                            RetMsg = "Stores Successfully removed from Location Group " & LocationGroupID
                        ElseIf Status = -2 Then
                            'Since the cards are checked earlier, it should never get to this point. 
                            RetCode = StatusCodes.INVALID_INCOMPLETEPROCESSLOCDATA
                            RetMsg = "Failure. Tried to remove a store that does not exist in the in the LocationGroup. No stores were removed"
                        End If
                    End If
                End If
            Else
                RetCode = StatusCodes.INVALID_INCOMPLETEPROCESSLOCDATA
                RetMsg = "Failure. BulkData is empty"
            End If

            'End Using

        Catch ex As Exception
            RetCode = StatusCodes.APPLICATION_EXCEPTION
            RetMsg = ex.ToString()
        Finally

            MyCommon.QueryStr = "IF OBJECT_ID('tempdb..#TempLocPK') IS NOT NULL BEGIN drop table #TempLocPK END "
            MyCommon.LRT_Execute()

            If MyCommon.LRTadoConn.State = ConnectionState.Open Then MyCommon.Close_LogixRT()

        End Try
        Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
    End Sub

#End Region

#Region "MutipleLocationGroup"
    Private Function _ProcessMulLocationsInLocationGroup(ByVal GUID As String, ByVal LocXML As String) As String

        Dim resultset As New DataSet
        Dim RetXmlDoc As New XmlDocument
        Dim RetXmlStr As String = ""
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim ReturnCode As StatusCodes = StatusCodes.SUCCESS ' This is what will eventually be the ReturnCode attribute. It will only be a success if all of the status codes are success
        Dim RetMsg As String = ""
        Dim LocationXmlDoc As New XmlDocument()
        Dim ConnInc As New Copient.ConnectorInc
        Dim ExtGroupID As String = ""
        Dim GroupName As String = ""
        Dim Name As String = ""
        Dim ErrorMessage As String = ""
        Dim Description As String = ""
        Dim dt As DataTable
        Dim ExtBannerID As String = ""
        Dim Operation As String = ""
        Dim BannersEnabled As Boolean = False
        Dim BannerID As Long
        Dim EngineType As Integer = -1
        Dim LocationGroupID As Long
        Dim StoreList As String = ""


        If isValidInputText(GUID) = False Then
            'GUID Value is empty
            RetCode = StatusCodes.INVALID_GUID
            RetMsg = "Failure. Invalid GUID"
            RetXmlStr = GetGroupResponseXML(RetCode, "", "LocationGroupUpdate", RetMsg)
            Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
            _ProcessMulLocationsInLocationGroup = RetXmlStr
            Exit Function
        End If
        If LocXML = "" Then
            'CustomerXML is empty
            RetCode = StatusCodes.INVALID_XML_DOCUMENT
            RetMsg = "Failure. Invalid Location XML"
            RetXmlStr = GetGroupResponseXML(RetCode, "", "LocationGroupUpdate", RetMsg)
            Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
            _ProcessMulLocationsInLocationGroup = RetXmlStr
            Exit Function
        End If
        Try
            Using MyCommon.LRTadoConn
                If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

                BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
                'Remove all invalid Unicode characters 
                ' Encode the XML string in a UTF-8 byte array     
                Dim encodedString As Byte() = Encoding.UTF8.GetBytes(LocXML)
                ' Put the byte array into a stream and rewind it to the beginning     
                Dim ms As New MemoryStream(encodedString)
                ms.Flush()
                ms.Position = 0
                LocXML = MyCommon.ReadAll(ms)

                ' validate the request
                If Not IsValidGUID(GUID, "Update") Then
                    RetCode = StatusCodes.INVALID_GUID
                    RetMsg = "GUID " & GUID & " is not valid for the this web service."
                    Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                ElseIf Not MyCommon.ConvertStringToXML(LocXML, LocationXmlDoc) Then
                    RetCode = StatusCodes.INVALID_XML_DOCUMENT
                    RetMsg = "LocationXML parameter is not a valid XML Document"
                    Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                ElseIf Not MyImport.ValidateXml("LocationGroupUpdate.xsd", LocXML) Then
                    RetCode = StatusCodes.INVALID_CRITERIA_XML
                    RetMsg = "XML document sent does not conform to the LocationGroupUpdate.xsd schema."
                    Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                End If
                If RetCode = StatusCodes.SUCCESS Then
                    Dim rootNode As XmlNode = LocationXmlDoc.DocumentElement
                    Dim nodeList As XmlNodeList

                    nodeList = rootNode.SelectNodes("/LocationGroupUpdate/ImportByExternalID")

                    For Each node In nodeList
                        RetCode = StatusCodes.SUCCESS
                        ErrorMessage = ""
                        Name = ""
                        LocationGroupID = Nothing
                        resultset = New DataSet()
                        resultset.ReadXml(New XmlNodeReader(LocationXmlDoc))

                        If Not node.Item("ID") Is Nothing Then
                            ExtGroupID = node.Item("ID").InnerText
                        Else
                            ExtGroupID = ""
                        End If
                        If Not node.Item("Name") Is Nothing Then
                            GroupName = node.Item("Name").InnerText
                        Else
                            GroupName = ""
                        End If
                        If Not node.Item("Operation") Is Nothing Then
                            Operation = node.Item("Operation").InnerText
                        Else
                            Operation = ""
                        End If
                        If Not node.Item("ExtBannerID") Is Nothing Then
                            ExtBannerID = node.Item("ExtBannerID").InnerText
                        Else
                            ExtBannerID = ""
                        End If
                        If Not node.Item("BulkData") Is Nothing Then
                            StoreList = node.Item("BulkData").InnerText
                        Else
                            StoreList = ""
                        End If


                        If (ExtGroupID Is Nothing OrElse ExtGroupID.Trim = "") Then
                            RetCode = StatusCodes.INVALID_LOCATIONGROUPID
                            RetMsg = "Failure. LocationGroupId is not Provided."
                        ElseIf Operation Is Nothing OrElse Operation.Trim = "" Then
                            RetCode = StatusCodes.INVALID_OPERATIONTYPE
                            RetMsg = "Failure. OperationType is not Provided"
                        ElseIf BannersEnabled = True Then
                            If isValidInputText(ExtBannerID) = False Then
                                RetCode = StatusCodes.INVALID_BANNER
                                RetMsg = "Failure. Banners are not enabled but location group was sent with EXTBannerID - can not process the group"
                            Else
                                MyCommon.QueryStr = "SELECT BannerID FROM Banners with (NoLock) WHERE EXTBannerID = @ExtBannerID AND Deleted=0"
                                MyCommon.DBParameters.Add("@ExtBannerID", SqlDbType.NVarChar).Value = ExtBannerID
                                dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                                If (dt.Rows.Count > 0) Then
                                    BannerID = MyCommon.NZ(dt.Rows(0).Item("BannerID"), 0)

                                    If BannerID > 0 Then
                                        MyCommon.QueryStr = "select EngineID from BannerEngines BE with (NoLock) where BannerID=@BannerID"
                                        MyCommon.DBParameters.Add("@BannerID", SqlDbType.Int).Value = BannerID
                                        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                                        If (dt.Rows.Count > 0) Then
                                            EngineType = MyCommon.NZ(dt.Rows(0).Item("EngineID"), EngineType)
                                        End If
                                    End If
                                Else
                                    RetCode = StatusCodes.INVALID_BANNER
                                    RetMsg = "Failure. Invalid BannerID."
                                End If
                            End If
                        End If

                        If RetCode = StatusCodes.SUCCESS Then
                            If MyLocGroupOp.LocationGroupExists(ExtGroupID, LocationGroupID, Name) Then
                                If Not (GroupName = "") AndAlso Not (GroupName = Name) Then
                                    MyLocGroupOp.LocationGroupNameExists(LGMLogFile, LocationGroupID, GroupName, RetCode, RetMsg)
                                    If RetCode = StatusCodes.SUCCESS Then
                                        MyLocGroupOp.UpdateLocationGroupDescription(LGMLogFile, LocationGroupID, Description)
                                    End If
                                End If
                            Else
                                If Not (GroupName = "") Then
                                    MyLocGroupOp.LocationGroupNameExists(LGMLogFile, LocationGroupID, GroupName, RetCode, RetMsg)
                                Else
                                    RetCode = StatusCodes.INVALID_LOCATIONNAME
                                    RetMsg = "Failure. GroupName is not provided"
                                End If
                                If EngineType = -1 Then
                                    MyCommon.QueryStr = "Select top 1 EngineID from PromoEngines where Installed = 1 and DefaultEngine = 1"
                                    dt = MyCommon.LRT_Select

                                    If dt.Rows.Count > 0 Then
                                        EngineType = dt.Rows(0).Item("EngineID")
                                    Else
                                        RetCode = StatusCodes.APPLICATION_EXCEPTION
                                        RetMsg = "Failure. Could not find a default promotion engine. "
                                        Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                                    End If

                                End If
                                If RetCode = StatusCodes.SUCCESS Then
                                    If Operation.ToUpper <> "REMOVE" Then
                                        LocationGroupID = MyLocGroupOp.CreateLocationGroupByExtID(ExtGroupID, BannerID, Description, ErrorMessage, EngineType, GroupName)
                                    Else
                                        RetCode = StatusCodes.INVALID_OPERATIONTYPE
                                        RetMsg = "Failure. Location group does not exist and operation type is remove."
                                    End If
                                Else
                                    ' RetXmlStr = GetLocationListErrorXML(RetCode, RetMsg)
                                    Copient.Logger.Write_Log(LGMLogFile, ErrorMessage, True)
                                End If
                            End If
                        Else
                            '  RetXmlStr = GetLocationListErrorXML(RetCode, RetMsg)
                            Copient.Logger.Write_Log(LGMLogFile, ErrorMessage, True)
                        End If

                        If RetCode = StatusCodes.SUCCESS Then
                            If Not (resultset.Tables("ImportByExternalID") Is Nothing) Then
                                AddColumnToTable(resultset, "ImportByExternalID", "statuscode", Type.GetType("System.String"))
                                AddColumnToTable(resultset, "ImportByExternalID", "Description", Type.GetType("System.String"))
                            End If

                            ProcessLocationGroup(LocationGroupID, StoreList, BannerID, ErrorMessage, Operation, RetCode, RetMsg)

                        Else
                            'RetXmlStr = GetLocationListErrorXML(RetCode, RetMsg, "LocationGroupUpdate")
                            Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                        End If

                        RetXmlStr &= GetInputResponseXML(RetCode, RetMsg, True, node)
                        If Not RetCode = StatusCodes.SUCCESS Then
                            ReturnCode = RetCode
                        End If

                    Next

                Else
                    'RetXmlStr = GetLocationListErrorXML(RetCode, RetMsg, "LocationGroupUpdate")
                    Copient.Logger.Write_Log(LGMLogFile, ErrorMessage, True)
                End If


            End Using
            If Not RetCode = StatusCodes.SUCCESS Then
                ReturnCode = RetCode
            End If
            RetXmlStr = GetGroupResponseXML(ReturnCode, RetXmlStr, "LocationGroupUpdate", RetMsg)

        Catch ex As Exception
            RetCode = StatusCodes.APPLICATION_EXCEPTION
            RetMsg = "Failure. Application " & ex.ToString
            RetXmlStr = GetLocationListErrorXML(RetCode, RetMsg, "LocationGroupUpdate")
        Finally
        End Try
        Return RetXmlStr
    End Function

    Private Function _ProcessMulLocationsInLocationGroupByLogixID(ByVal GUID As String, ByVal LocXML As String) As String

        Dim resultset As New DataSet
        Dim RetXmlDoc As New XmlDocument
        Dim RetXmlStr As String = ""
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim ReturnCode As StatusCodes = StatusCodes.SUCCESS ' This is what will eventually be the ReturnCode attribute. It will only be a success if all of the status codes are success
        Dim RetMsg As String = ""
        Dim LocationXmlDoc As New XmlDocument()
        Dim ConnInc As New Copient.ConnectorInc
        Dim ExtGroupID As String = ""
        Dim GroupName, Name, ErrorMessage As String
        Dim Description As String = ""
        Dim dt As DataTable
        Dim ExtBannerID As String = ""
        Dim Operation As String = ""
        Dim BannersEnabled As Boolean = False
        Dim BannerID As Long
        Dim EngineType As Integer
        Dim LocationGroupID As Long
        Dim StoreList As String = ""

        If isValidInputText(GUID) = False Then
            'GUID Value is empty
            RetCode = StatusCodes.INVALID_GUID
            RetMsg = "Failure. Invalid GUID"
            RetXmlStr = GetGroupResponseXML(RetCode, "", "LocationGroupUpdate", RetMsg)
            Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
            _ProcessMulLocationsInLocationGroupByLogixID = RetXmlStr
            Exit Function
        End If
        If LocXML = "" Then
            'CustomerXML is empty
            RetCode = StatusCodes.INVALID_XML_DOCUMENT
            RetMsg = "Failure. Invalid Location XML"
            RetXmlStr = GetGroupResponseXML(RetCode, "", "LocationGroupUpdate", RetMsg)
            Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
            _ProcessMulLocationsInLocationGroupByLogixID = RetXmlStr
            Exit Function
        End If

        Try
            Using MyCommon.LRTadoConn
                If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

                BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
                'Remove all invalid Unicode characters 
                ' Encode the XML string in a UTF-8 byte array     
                Dim encodedString As Byte() = Encoding.UTF8.GetBytes(LocXML)
                ' Put the byte array into a stream and rewind it to the beginning     
                Dim ms As New MemoryStream(encodedString)
                ms.Flush()
                ms.Position = 0
                LocXML = MyCommon.ReadAll(ms)
                ' validate the request
                If Not IsValidGUID(GUID, "Update") Then
                    RetCode = StatusCodes.INVALID_GUID
                    RetMsg = "GUID " & GUID & " is not valid for the this web service."
                    Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                ElseIf Not MyCommon.ConvertStringToXML(LocXML, LocationXmlDoc) Then
                    RetCode = StatusCodes.INVALID_XML_DOCUMENT
                    RetMsg = "LocationXML parameter is not a valid XML Document"
                    Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                ElseIf Not MyImport.ValidateXml("LocationGroupUpdate.xsd", LocXML) Then
                    RetCode = StatusCodes.INVALID_CRITERIA_XML
                    RetMsg = "XML document sent does not conform to the LocationGroupUpdate.xsd schema."
                    Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                End If

                If RetCode = StatusCodes.SUCCESS Then
                    Dim rootNode As XmlNode = LocationXmlDoc.DocumentElement
                    Dim nodeList As XmlNodeList

                    nodeList = rootNode.SelectNodes("/LocationGroupUpdate/ImportByLogixID")

                    For Each node In nodeList
                        RetCode = StatusCodes.SUCCESS
                        ErrorMessage = ""
                        Name = ""
                        LocationGroupID = Nothing
                        resultset = New DataSet()
                        resultset.ReadXml(New XmlNodeReader(LocationXmlDoc))


                        If Not node.Item("ID") Is Nothing Then
                            LocationGroupID = node.Item("ID").InnerText
                        Else
                            LocationGroupID = ""
                        End If
                        If Not node.Item("Name") Is Nothing Then
                            GroupName = node.Item("Name").InnerText
                        Else
                            GroupName = ""
                        End If
                        If Not node.Item("ExtBannerID") Is Nothing Then
                            ExtBannerID = node.Item("ExtBannerID").InnerText
                        Else
                            ExtBannerID = ""
                        End If
                        If Not node.Item("Operation") Is Nothing Then
                            Operation = node.Item("Operation").InnerText
                        Else
                            Operation = ""
                        End If
                        If Not node.Item("BulkData") Is Nothing Then
                            StoreList = node.Item("BulkData").InnerText
                        Else
                            StoreList = ""
                        End If


                        If Operation Is Nothing OrElse Operation.Trim = "" Then
                            RetCode = StatusCodes.INVALID_OPERATIONTYPE
                            RetMsg = "Failure. OperationType is not Provided"
                        ElseIf BannersEnabled = True Then
                            If isValidInputText(ExtBannerID) = False Then
                                RetCode = StatusCodes.INVALID_BANNER
                                RetMsg = "Failure. Banners are not enabled but location group was sent with EXTBannerID - can not process the group"
                            Else
                                MyCommon.QueryStr = "SELECT BannerID FROM Banners with (NoLock) WHERE EXTBannerID = @ExtBannerID AND Deleted=0"
                                MyCommon.DBParameters.Add("@ExtBannerID", SqlDbType.NVarChar).Value = ExtBannerID
                                dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                                If (dt.Rows.Count > 0) Then
                                    BannerID = MyCommon.NZ(dt.Rows(0).Item("BannerID"), 0)
                                    If BannerID > 0 Then
                                        MyCommon.QueryStr = "select EngineID from BannerEngines BE with (NoLock) where BannerID=@BannerID"
                                        MyCommon.DBParameters.Add("@BannerID", SqlDbType.Int).Value = BannerID
                                        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                                        If (dt.Rows.Count > 0) Then
                                            EngineType = MyCommon.NZ(dt.Rows(0).Item("EngineID"), EngineType)
                                        End If
                                    End If
                                Else
                                    RetCode = StatusCodes.INVALID_BANNER
                                    RetMsg = "Failure. Invalid BannerID."
                                End If
                            End If
                        End If

                        If RetCode = StatusCodes.SUCCESS Then
                            If MyLocGroupOp.LocationGroupExists(LocationGroupID, Name) Then
                                If Not (GroupName = "") AndAlso Not (GroupName = Name) Then
                                    MyLocGroupOp.LocationGroupNameExists(LGMLogFile, LocationGroupID, GroupName, RetCode, RetMsg)
                                    If RetCode = StatusCodes.SUCCESS Then
                                        MyLocGroupOp.UpdateLocationGroupDescription(LGMLogFile, LocationGroupID, Description)
                                    End If
                                End If
                            Else
                                If Not (GroupName = "") Then
                                    MyLocGroupOp.LocationGroupNameExists(LGMLogFile, LocationGroupID, GroupName, RetCode, RetMsg)
                                Else
                                    RetCode = StatusCodes.INVALID_LOCATIONNAME
                                    RetMsg = "Failure. GroupName is not provided"
                                End If
                                If RetCode = StatusCodes.SUCCESS Then
                                    If Operation.ToUpper <> "REMOVE" Then
                                        LocationGroupID = MyLocGroupOp.CreateLocationGroupByLogixID(BannerID, Description, ErrorMessage, EngineType, GroupName, RetCode, RetMsg)
                                    Else
                                        RetCode = StatusCodes.INVALID_OPERATIONTYPE
                                        RetMsg = "Failure. Location group does not exist and operation type is remove."
                                    End If
                                Else
                                    Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                                End If
                            End If
                        Else
                            Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                        End If

                        If RetCode = StatusCodes.SUCCESS Then
                            If Not (resultset.Tables("ImportByLogixID") Is Nothing) Then
                                AddColumnToTable(resultset, "ImportByLogixID", "statuscode", Type.GetType("System.String"))
                                AddColumnToTable(resultset, "ImportByLogixID", "Description", Type.GetType("System.String"))
                            End If

                            ProcessLocationGroup(LocationGroupID, StoreList, BannerID, ErrorMessage, Operation, RetCode, RetMsg)
                        Else
                            Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                        End If

                        RetXmlStr &= GetInputResponseXML(RetCode, RetMsg, False, node)
                        If Not RetCode = StatusCodes.SUCCESS Then
                            ReturnCode = RetCode
                        End If
                    Next
                Else
                    Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                End If
            End Using
            If Not RetCode = StatusCodes.SUCCESS Then
                ReturnCode = RetCode
            End If
            RetXmlStr = GetGroupResponseXML(ReturnCode, RetXmlStr, "LocationGroupUpdate", RetMsg)

        Catch ex As Exception
            RetCode = StatusCodes.APPLICATION_EXCEPTION
            RetMsg = "Failure. Application " & ex.ToString
            RetXmlStr = GetLocationListErrorXML(RetCode, RetMsg, "LocationGroupUpdate")
        Finally
        End Try
        Return RetXmlStr
    End Function
#End Region

#Region "WebMethods"

#Region "Single Customer Group"

    <WebMethod()> _
    Public Function ProcessCardInCustGroupByExtGrpID(ByVal GUID As String, ByVal ExtGroupID As String, ByVal GroupName As String, _
                                           ByVal ExtCardId As String, ByVal CardTypeID As String, ByVal OperationType As String) As DataSet
        InitApp()
        Return _ProcessCardInCustomerGroup(GUID, ExtGroupID, -1, GroupName, ExtCardId, CardTypeID, OperationType, True)
    End Function

    <WebMethod()> _
    Public Function ProcessCardInCustGroupByLogixID(ByVal GUID As String, ByVal GroupID As String, ByVal GroupName As String, _
                                           ByVal ExtCardId As String, ByVal CardTypeID As String, ByVal OperationType As String) As DataSet
        InitApp()
        Dim iGroupID As Integer = -1
        Try
            If (GroupID <> "") Then
                iGroupID = CInt(GroupID)
            End If
        Catch ex As Exception

        End Try
        Return _ProcessCardInCustomerGroup(GUID, "", iGroupID, GroupName, ExtCardId, CardTypeID, OperationType, False)
    End Function

#End Region

#Region "Multiple Customer Group"

    <WebMethod()> _
    Public Function ProcessMultipleCardsInCustGroupByExtGrpID(ByVal GUID As String, ByVal CustomerXML As String) As String
        InitApp()
        Return _ProcessMultipleCardInCustomerGroup(GUID, CustomerXML, True)
    End Function

    <WebMethod()> _
    Public Function ProcessMultipleCardsInCustGroupByLogixID(ByVal GUID As String, ByVal CustomerXML As String) As String
        InitApp()
        Return _ProcessMultipleCardInCustomerGroup(GUID, CustomerXML, False)
    End Function

#End Region

#Region "Single Location Group"

    <WebMethod()> _
    Public Function ProcessLocInLocGroupByExtGrpID(ByVal GUID As String, ByVal EXTGroupID As String, ByVal Name As String, ByVal Description As String, ByVal ExtLocationID As String, ByVal LocationName As String, ByVal ExtBannerID As String, ByVal OperationType As String) As DataSet

        Return _ProcessLocationInLocationGroup(GUID, EXTGroupID, Name, Description, ExtLocationID, LocationName, ExtBannerID, OperationType)

    End Function

    <WebMethod()> _
    Public Function ProcessLocInLocGroupByLogixID(ByVal GUID As String, ByVal GroupID As String, ByVal Name As String, ByVal Description As String, ByVal ExtLocationID As String, ByVal LocationName As String, ByVal ExtBannerID As String, ByVal OperationType As String) As DataSet

        Dim iGroupID As Long
        If Long.TryParse(GroupID, iGroupID) Then
            Return _ProcessLocationInLocationGroup(GUID, iGroupID, Name, Description, ExtLocationID, LocationName, ExtBannerID, OperationType)
        Else
            Return _ProcessLocationInLocationGroup(GUID, -1, Name, Description, ExtLocationID, LocationName, ExtBannerID, OperationType)
        End If

    End Function

#End Region

#Region "Multiple Location Group"

    <WebMethod()> _
    Public Function ProcessMultipleLocInLocGroupByExtGrpID(ByVal GUID As String, ByVal LocationXML As String) As String
        InitApp()
        Return _ProcessMulLocationsInLocationGroup(GUID, LocationXML)
    End Function

    <WebMethod()> _
    Public Function ProcessMultipleLocInLocGroupByLogixID(ByVal GUID As String, ByVal LocationXML As String) As String
        InitApp()
        Return _ProcessMulLocationsInLocationGroupByLogixID(GUID, LocationXML)
    End Function

#End Region


#End Region

#Region "SingleProduct"

    'Validating Single Product inputs
    Private Function SingleProductValidations(ByVal Guid As String, ByVal GroupID As String, ByVal GroupName As String, ByVal ExtProductID As String, _
                                                 ByVal Description As String, ByVal OperationType As String, ByVal GroupIDType As String, _
                                                 ByRef StatusCode As String, ByRef MessageDescription As String, Optional ByVal ProductTypeId As Integer = -1) As Boolean

        Dim MethodName As String = "ProcessProductInProductGroup"
        If (Guid Is Nothing OrElse String.IsNullOrEmpty(Guid.Trim)) Then
            'If String.IsNullOrWhiteSpace(Guid) Then
            StatusCode = StatusCodes.INVALID_GUID
            MessageDescription = "Failure. Invalid GUID"
            Return False
        End If
        If Guid.Contains("'") = True Or Guid.Contains(Chr(34)) = True Then
            StatusCode = StatusCodes.INVALID_GUID
            MessageDescription = "Failure. Invalid GUID"
            Return False
        End If

        If GroupID.Contains("'") = True Or GroupID.Contains(Chr(34)) = True Then
            StatusCode = StatusCodes.INVALID_EXTPRODGROUPID
            MessageDescription = "Failure. Invalid Product Group ID"
            Return False
        End If
        If GroupIDType = "LogixID" Then
            If Not IsNumeric(GroupID) And Not String.IsNullOrEmpty(GroupID.Trim) Then
                StatusCode = StatusCodes.INVALID_PRODGROUPID
                MessageDescription = "Failure. Invalid Logix Group ID"
                Return False
            End If
        Else
            If (GroupID Is Nothing OrElse String.IsNullOrEmpty(GroupID.Trim)) Then
                'If String.IsNullOrWhiteSpace(GroupID) Then
                StatusCode = StatusCodes.INVALID_EXTPRODGROUPID
                MessageDescription = "Failure. Invalid Product Group ID"
                Return False
            End If
        End If
        If GroupName.Contains("'") = True Or GroupName.Contains(Chr(34)) = True Then
            StatusCode = StatusCodes.INVALID_PRODGROUPNAME
            MessageDescription = "Failure. Invalid Product Group Name"
            Return False
        End If
        If (ExtProductID Is Nothing OrElse String.IsNullOrEmpty(ExtProductID.Trim)) Then
            'If String.IsNullOrWhiteSpace(ExtProductID) Then
            StatusCode = StatusCodes.INVALID_PRODUCTID
            MessageDescription = "Failure. Invalid External Product ID"
            Return False
        End If
        If ExtProductID.Contains("'") = True Or ExtProductID.Contains(Chr(34)) = True Then
            StatusCode = StatusCodes.INVALID_PRODUCTID
            MessageDescription = "Failure. Invalid External Product ID"
            Return False
        End If

        If Not IsPositiveID(ExtProductID) Then
            StatusCode = StatusCodes.INVALID_PRODUCTID
            MessageDescription = "Failure. Invalid External Product ID"
            Return False
        End If

        If Not CleanUPC(ExtProductID) Then
            StatusCode = StatusCodes.INVALID_PRODUCTID
            MessageDescription = "Failure. Invalid External Product ID"
            Return False
        End If

        If ProductTypeId = -1 Then
            StatusCode = StatusCodes.INVALID_PRODUCTTYPEID
            MessageDescription = "Failure. Invalid Product Type ID"
            Return False
        End If
        If Description.Contains("'") = True Or Description.Contains(Chr(34)) = True Then
            StatusCode = StatusCodes.INVALID_DESCRIPTION
            MessageDescription = "Failure. Invalid Product Description"
            Return False
        End If
        'If String.IsNullOrWhiteSpace(OperationType) Then
        If (OperationType Is Nothing OrElse String.IsNullOrEmpty(OperationType.Trim)) Then
            StatusCode = StatusCodes.INVALID_OPERATIONTYPE
            MessageDescription = "Failure. Invalid Operation Type"
            Return False
        End If
        OperationType = OperationType.ToUpper
        If OperationType <> "AUGMENT" And OperationType <> "REMOVE" And OperationType <> "REPLACE" Then
            StatusCode = StatusCodes.INVALID_OPERATIONTYPE
            MessageDescription = "Failure. Invalid Operation Type"
            Return False
        End If
        If Not IsValidGUID(Guid, MethodName) Then
            StatusCode = StatusCodes.INVALID_GUID
            MessageDescription = "Failure. Invalid GUID"
            Return False
        End If

        REM MyCommon.QueryStr = "Select ProductID from Products where ExtProductID = " & ExtProductID & " ProductTypeID = " & ProductTypeID 
        REM dst = MyCommon.LRT_Select 
        REM If dst.Rows.Count = 0 Then
        REM StatusCode = StatusCodes.INVALIDPRODUCTID
        REM MessageDescription = "Failure. Cannot 


        Return True
    End Function

    Public Function IsPositiveID(ByVal ID As String) As Boolean
        Dim isValid As Boolean = True
        Try
            If IsNumeric(ID) Then
                Dim iID As Integer = -1
                iID = CInt(ID)
                If (iID < 1) Then
                    isValid = False
                End If
            End If
        Catch ex As Exception
            isValid = False
        End Try
        Return isValid
    End Function

    Public Function CleanUPC(ByVal InString As String) As Boolean
        Dim z As Integer
        Dim IsClean As Boolean = True

        If InString IsNot Nothing Then
            For z = 0 To InString.Length - 1
                If (InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_-", InString(z))) Then
                    IsClean = True
                Else
                    Return False
                End If
            Next
        End If
        Return IsClean
    End Function


    'Creation of Groups
    Private Function CreateGroups(ByVal GroupID As String, ByVal GroupName As String, ByVal OperationType As String, ByVal GroupIDType As String,
                                  ByRef StatusCode As String, ByRef MessageDescription As String, ByRef ProductGroupID As Long) As Boolean
        Dim result As Boolean



        If GroupIDType = "LogixID" Then
            If String.IsNullOrEmpty(GroupID.Trim) Then
                GroupID = 0
            End If

            result = MyProdGroupOp.CreateProductGroupByLogixID(Convert.ToInt64(GroupID), Trim(GroupName), Trim(OperationType), ProductGroupID, MessageDescription)
        Else
            result = MyProdGroupOp.CreateProductGroupByExternalID(Trim(GroupID), Trim(GroupName), Trim(OperationType), ProductGroupID, MessageDescription)
        End If
        If Not result Then
            StatusCode = StatusCodes.INVALID_PRODGROUPNAME
            MessageDescription = "Failure. " & MessageDescription
            Copient.Logger.Write_Log(LGMLogFile, "Failure. " & MessageDescription, True)
            Return False
        Else
            Copient.Logger.Write_Log(LGMLogFile, "Success. " & MessageDescription, True)

        End If
        Return True
    End Function

    Private Function _ProcessProductInProductGroup(ByVal Guid As String, ByVal GroupID As String, ByVal GroupName As String, ByVal ExtProductID As String,
                                                 ByVal Description As String, ByVal OperationType As String, ByVal GroupIDType As String, Optional ByVal ProductTypeId As Integer = -1) As System.Data.DataSet

        Dim ProductGroupID As Long = 0
        Dim dtStatus As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim StatusCode, MessageDescription As String
        StatusCode = String.Empty
        MessageDescription = String.Empty
        Dim ResultSet As New System.Data.DataSet("ProcessProductInProductGroup")
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = String.Empty
        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        If Not SingleProductValidations(Guid, GroupID, GroupName, ExtProductID, Description, OperationType, GroupIDType,
                                                 StatusCode, MessageDescription, ProductTypeId) Then
            row = dtStatus.NewRow()

            row.Item("StatusCode") = StatusCode
            row.Item("Description") = MessageDescription
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            Copient.Logger.Write_Log(LGMLogFile, MessageDescription, True)
            _ProcessProductInProductGroup = ResultSet
            Exit Function
        End If
        Try
            Using MyCommon.LRTadoConn
                If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                    MyCommon.Open_LogixRT()
                End If



                'Creation of Groups
                If Not CreateGroups(GroupID, GroupName, OperationType, GroupIDType, StatusCode, MessageDescription, ProductGroupID) Then
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = StatusCode
                    row.Item("Description") = MessageDescription
                    dtStatus.Rows.Add(row)
                    dtStatus.AcceptChanges()
                    ResultSet.Tables.Add(dtStatus.Copy())
                    Copient.Logger.Write_Log(LGMLogFile, MessageDescription, True)
                    _ProcessProductInProductGroup = ResultSet
                    Exit Function
                End If

                ProcessProductGroup(ExtProductID & "," & ProductTypeId & "," & Description, ProductGroupID, OperationType, RetCode, RetMsg)

                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                _ProcessProductInProductGroup = ResultSet
            End Using
        Catch ex As Exception
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure. Application " & ex.ToString
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
        End Try
        Return ResultSet
    End Function
    'Web Method used to Process Product by ExtID
    <WebMethod()>
    Public Function ProcessProdInProdGroupByExtGrpID(ByVal GUID As String, ByVal ExtGroupID As String, ByVal GroupName As String, ByVal ExtProductID As String,
                                             ByVal ProductTypeId As String, ByVal Description As String, ByVal OperationType As String) As System.Data.DataSet
        Dim iProductTypeId As Integer = -1
        Try
            InitApp()
            If Not String.IsNullOrEmpty(ProductTypeId) Then
                iProductTypeId = CInt(ProductTypeId)
            Else
                iProductTypeId = 1
            End If
        Catch ex As Exception
            iProductTypeId = -1
        End Try
        Return _ProcessProductInProductGroup(GUID, ExtGroupID, GroupName, ExtProductID, Description, OperationType, "ExtID", iProductTypeId)
    End Function
    'Web Method used to Process Product by logixID
    <WebMethod()>
    Public Function ProcessProdInProdGroupByLogixID(ByVal GUID As String, ByVal LogixGroupID As String, ByVal GroupName As String, ByVal ExtProductID As String,
                                             ByVal ProductTypeId As String, ByVal Description As String, ByVal OperationType As String) As System.Data.DataSet
        Dim iProductTypeId As Integer = -1
        Try
            InitApp()
            If Not String.IsNullOrEmpty(ProductTypeId) Then
                iProductTypeId = CInt(ProductTypeId)
            Else
                iProductTypeId = 1
            End If
        Catch ex As Exception
            iProductTypeId = -1
        End Try
        Return _ProcessProductInProductGroup(GUID, LogixGroupID, GroupName, ExtProductID, Description, OperationType, "LogixID", iProductTypeId)
    End Function

#End Region

#Region "MultipleProduct"
    <WebMethod()>
    Public Function ProcessMultipleProdInProdGroupByExtGrpID(ByVal GUID As String, ByVal ProductXML As String) As String
        InitApp()
        Return _ProcessMultipleProductsInProductGroup(GUID, ProductXML, "ExtID")

    End Function
    <WebMethod()>
    Public Function ProcessMultipleProdInProdGroupByLogixID(ByVal GUID As String, ByVal ProductXML As String) As String
        InitApp()
        Return _ProcessMultipleProductsInProductGroup(GUID, ProductXML, "LogixID")

    End Function

    'Validating Multiple Products inputs
    Private Function MultipleProductValidations(ByVal Guid As String, ByVal ProductXML As String, ByRef StatusCode As String,
                                                ByRef MessageDescription As String, ByRef ProductXmlDoc As XmlDocument) As Boolean

        Dim MethodName As String = "ProcessMultipleProductsInProductGroup"

        If (Guid Is Nothing OrElse String.IsNullOrEmpty(Guid.Trim)) Then
            StatusCode = StatusCodes.INVALID_GUID
            MessageDescription = "Failure. Invalid GUID"
            Return False
        End If

        If Guid.Contains("'") = True Or Guid.Contains(Chr(34)) = True Then
            StatusCode = StatusCodes.INVALID_GUID
            MessageDescription = "Failure. Invalid GUID"
            Return False
        End If

        'If String.IsNullOrWhiteSpace(ProductXML) Then
        If (ProductXML Is Nothing OrElse String.IsNullOrEmpty(ProductXML.Trim)) Then
            StatusCode = StatusCodes.INVALID_XML_DOCUMENT
            MessageDescription = "Failure. Invalid Product XML"
            Return False
        End If

        'Remove all invalid unicode characters 
        ' Encode the XML string in a UTF-8 byte array     
        Dim encodedString As Byte() = Encoding.UTF8.GetBytes(ProductXML)
        ' Put the byte array into a stream and rewind it to the beginning     
        Dim ms As New MemoryStream(encodedString)
        ms.Flush()
        ms.Position = 0
        ProductXML = MyCommon.ReadAll(ms)

        If Not IsValidGUID(Guid, MethodName) Then
            StatusCode = StatusCodes.INVALID_GUID
            MessageDescription = "GUID " & Guid & " is not valid for the LogixGroupManagement web service."
            Copient.Logger.Write_Log(LGMLogFile, MessageDescription, True)
            Return False
        ElseIf Not MyCommon.ConvertStringToXML(ProductXML, ProductXmlDoc) Then
            'If Not ConnInc.ConvertStringToXML(CustomerXML, CustomerXmlDoc) Then
            StatusCode = StatusCodes.INVALID_XML_DOCUMENT
            MessageDescription = "ProductXML parameter is not a valid XML Document"
            Copient.Logger.Write_Log(LGMLogFile, MessageDescription, True)
            Return False
        ElseIf Not MyImport.ValidateXml("ProductGroupUpdate.xsd", ProductXML) Then
            StatusCode = StatusCodes.INVALID_CRITERIA_XML
            MessageDescription = "XML document sent does not conform to the ProductGroupUpdate.xsd schema."
            Copient.Logger.Write_Log(LGMLogFile, MessageDescription, True)
            Return False
        End If

        Return True
    End Function

    Private Function _ProcessMultipleProductsInProductGroup(ByVal GUID As String, ByVal ProductXML As String, ByVal GroupIDType As String) As String
        Dim StatusCode, MessageDescription As String
        StatusCode = String.Empty
        MessageDescription = String.Empty
        Dim ResultSet As System.Data.DataSet = Nothing
        Dim row As System.Data.DataRow
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim ReturnCode As StatusCodes = StatusCodes.SUCCESS ' This is what will eventually be the ReturnCode attribute. It will only be a success if all of the status codes are success
        Dim RetMsg As String = String.Empty
        Dim ProductXmlDoc As New XmlDocument()
        Dim ExtGroupID As String = String.Empty
        Dim GroupName As String = String.Empty
        Dim Operation As String = String.Empty
        Dim ProdList As String = String.Empty
        Dim dtprocessprodgroup As DataTable
        Dim ProdListData() As String = Nothing
        Dim ProductData() As String = Nothing
        Dim iRowCnt As Integer = -1
        Dim ProductGroupID As Long = 0
        Dim ProductTypeID As Integer = 1
        Dim ProductDesc As String = String.Empty
        Dim ErrMsg As String = String.Empty
        Dim ProdLength As Integer = 0
        Dim ProdCount As Integer = 0
        Dim RetXmlStr As String = String.Empty
        Dim tagID As String = String.Empty
        Dim RetXmlDoc As New XmlDocument
        Dim DelFlag As Boolean = True
        Dim rootNode As XmlNode
        Dim nodeList As XmlNodeList
        Dim ByExtID As Boolean

        If Not MultipleProductValidations(GUID, ProductXML, StatusCode, MessageDescription, ProductXmlDoc) Then
            RetCode = StatusCode
            RetMsg = MessageDescription
            RetXmlStr = GetGroupResponseXML(RetCode, "", "ProductGroupUpdate", RetMsg)
            Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
            _ProcessMultipleProductsInProductGroup = RetXmlStr
            Exit Function
        End If

        If GroupIDType = "LogixID" Then
            ByExtID = False
            If ProductXML.Contains("ImportByExternalID") Then
                'ProductXML is empty
                RetCode = StatusCodes.INVALID_XML_DOCUMENT
                RetMsg = "Failure. Invalid ProductXML"
                RetXmlStr = GetGroupResponseXML(RetCode, "", "ProductGroupUpdate", RetMsg)
                Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                _ProcessMultipleProductsInProductGroup = RetXmlStr
                Exit Function
            End If
        Else
            ByExtID = True
            If ProductXML.Contains("ImportByLogixID") Then
                'ProductXML is empty
                RetCode = StatusCodes.INVALID_XML_DOCUMENT
                RetMsg = "Failure. Invalid ProductXML"
                RetXmlStr = GetGroupResponseXML(RetCode, "", "ProductGroupUpdate", RetMsg)
                Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                _ProcessMultipleProductsInProductGroup = RetXmlStr
                Exit Function
            End If
        End If

        Try
            Using MyCommon.LRTadoConn
                If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                    MyCommon.Open_LogixRT()
                End If
                rootNode = ProductXmlDoc.DocumentElement

                If ByExtID = True Then
                    nodeList = rootNode.SelectNodes("/ProductGroupUpdate/ImportByExternalID")
                Else
                    nodeList = rootNode.SelectNodes("/ProductGroupUpdate/ImportByLogixID")
                End If
                For Each node In nodeList
                    If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                        MyCommon.Open_LogixRT()
                    End If
                    ProductGroupID = 0
                    RetCode = StatusCodes.SUCCESS
                    MessageDescription = ""
                    ResultSet = New DataSet()
                    ResultSet.ReadXml(New XmlNodeReader(ProductXmlDoc))
                    If Not node.Item("ID") Is Nothing Then
                        ExtGroupID = node.Item("ID").InnerText
                    End If
                    If Not node.Item("Name") Is Nothing Then
                        GroupName = node.Item("Name").InnerText
                    End If
                    If Not node.Item("Operation") Is Nothing Then
                        Operation = node.Item("Operation").InnerText
                    End If
                    If Not node.Item("BulkData") Is Nothing Then
                        ProdList = node.Item("BulkData").InnerText
                    Else
                        ProdList = ""
                    End If


                    If (ExtGroupID Is Nothing OrElse String.IsNullOrEmpty(ExtGroupID.Trim)) Then
                        If Not GroupIDType = "LogixID" Then
                            RetCode = StatusCodes.INVALID_PRODUCTID
                            RetMsg = "Failure. Product Group Id is not Provided."
                        End If
                    End If
                    If Operation Is Nothing OrElse String.IsNullOrEmpty(Operation.Trim) OrElse (Operation <> "augment" And Operation <> "replace" And Operation <> "remove") Then
                        RetCode = StatusCodes.INVALID_OPERATIONTYPE
                        RetMsg = "Failure. Operation Type is not Provided"
                    End If
                    If RetCode = StatusCodes.SUCCESS Then
                        'Creation of Groups
                        If Not CreateGroups(ExtGroupID, GroupName, Operation, GroupIDType, StatusCode, MessageDescription, ProductGroupID) Then
                            RetCode = StatusCode
                            RetMsg = MessageDescription

                            GoTo BuildErrorXML
                        End If


                        If Not (ResultSet.Tables(tagID) Is Nothing) Then
                            AddColumnToTable(ResultSet, tagID, "statuscode", Type.GetType("System.String"))
                            AddColumnToTable(ResultSet, tagID, "Description", Type.GetType("System.String"))
                        End If

                        dtprocessprodgroup = New DataTable("ProductGroups")
                        dtprocessprodgroup.Columns.Add("Status", System.Type.GetType("System.String"))
                        row = dtprocessprodgroup.NewRow()

                        ProcessProductGroup(ProdList, ProductGroupID, Operation, RetCode, RetMsg)

                    End If

BuildErrorXML:
                    RetXmlStr &= GetInputResponseXML(RetCode, RetMsg, ByExtID, node)
                    If Not RetCode = StatusCodes.SUCCESS Then
                        ReturnCode = RetCode
                    End If
                Next
            End Using
            If Not RetCode = StatusCodes.SUCCESS Then
                ReturnCode = RetCode
            End If
            RetXmlStr = GetGroupResponseXML(ReturnCode, RetXmlStr, "ProductGroupUpdate", RetMsg)


        Catch ex As Exception
            RetCode = StatusCodes.APPLICATION_EXCEPTION
            RetMsg = "Failure. Application " & ex.ToString
            'RetXmlStr = GetProductListErrorXML(RetCode, RetMsg, "ProductGroupUpdate")
            RetXmlStr = GetGroupResponseXML(ReturnCode, "", "ProductGroupUpdate", RetMsg)
            _ProcessMultipleProductsInProductGroup = RetXmlStr
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try
        Return RetXmlStr
    End Function

    Private Function BuildErrorMessage(ByVal ErrMsg As String, ByRef Data As String) As String
        If Not String.IsNullOrEmpty(ErrMsg) Then
            ErrMsg = ErrMsg & ", " & Data
        Else
            ErrMsg = Data
        End If
        Return ErrMsg
    End Function

    Private Sub ProcessProductGroup(ByVal ProdList As String, ByVal ProductGroupID As Integer, ByVal Operation As String, ByRef RetCode As Integer, ByRef RetMsg As String)
        Dim ItemList() As String = Nothing
        Dim ItemData() As String = Nothing
        Dim querystr As String
        Dim dst As DataTable
        Dim Status As Integer
        Dim ProductID As Integer
        Dim productParameter As SqlParameter
        Dim productParameterList As New List(Of SqlParameter)

        Try
            'Using MyCommon.LRTadoConn
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If ProdList IsNot Nothing AndAlso ProdList.Trim > "" Then
                ProdList = ProdList.Replace(vbCrLf, vbLf)
                ProdList = ProdList.Replace(vbLf, vbCrLf)
                ItemList = ProdList.Split(ControlChars.CrLf)
                MyCommon.QueryStr = "IF OBJECT_ID('tempdb..#TempProdPK') IS NOT NULL BEGIN drop table #TempProdPK END "
                MyCommon.LRT_Execute()
                querystr = "create table #TempProdPK([TempPK] int PRIMARY KEY IDENTITY," &
                     "[ProductID] bigint NOT NULL)"
                For Each prod In ItemList

                    ItemData = prod.Trim.Split(",")
                    Dim ExtProductID As String = Trim(ItemData(0))
                    Dim ProductTypeID As Integer
                    Dim ProductDesc As String

                    If ItemData.Length = 1 Then
                        If ExtProductID.Length = 0 Then
                            Copient.Logger.Write_Log(LGMLogFile, "Skipping blank item entry", True)
                            Continue For
                        Else
                            RetCode = StatusCodes.INVALID_INCOMPLETEPROCESSPRODDATA
                            RetMsg = "Failure. Invalid Product Bulk Data."
                            Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                            Return
                        End If
                    ElseIf ItemData.Length <> 3 Then
                        RetCode = StatusCodes.INVALID_INCOMPLETEPROCESSPRODDATA
                        RetMsg = "Failure. Invalid Product Bulk Data."
                        Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                        Return
                    End If

                    ProductDesc = Trim(ItemData(2))
                    If Integer.TryParse(ItemData(1), ProductTypeID) = False Then
                        RetCode = StatusCodes.INVALID_CARDTYPEID
                        RetMsg = "Failure. Invalid product type: " & ItemData(1)
                        Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                        Return
                    Else
                        MyCommon.QueryStr = "select ProductTypeID from ProductTypes with (NoLock) " &
                                            "where ProductTypeID=@ProductTypeID"
                        MyCommon.DBParameters.Add("@ProductTypeID", SqlDbType.Int).Value = ProductTypeID
                        dst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                        If dst.Rows.Count = 0 Then
                            RetCode = StatusCodes.INVALID_CARDTYPEID
                            RetMsg = "Failure. Invalid product type: " & ItemData(1)
                            Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                            Return
                        End If
                    End If

                    ExtProductID = MyProdGroupOp.PadExternalProductID(ExtProductID, ProductTypeID)

                    '				 If ExtProductID <> "" Then
                    ' determine if the location already exists
                    MyCommon.QueryStr = "select ProductID from Products with (NoLock) where (ExtProductID = @ExtProductID and " &
                                        " ProductTypeID = @ProductTypeID)"
                    MyCommon.DBParameters.Add("@ExtProductID", SqlDbType.NVarChar).Value = ExtProductID
                    MyCommon.DBParameters.Add("@ProductTypeID", SqlDbType.Int).Value = ProductTypeID
                    dst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                    If dst.Rows.Count > 0 Then
                        querystr &= "insert into #TempProdPK values(" & MyCommon.NZ(dst.Rows(0).Item("ProductID"), 0) & "); "
                        'MyCommon.DBParameters.Add("@ProductID",SqlDbType.BigInt).Value = MyCommon.NZ(dst.Rows(0).Item("ProductID"), 0)
                        'productParameter = New SqlParameter("@ProdID" & ExtProductID, SqlDbType.BigInt)
                        'productParameter.Value = MyCommon.NZ(dst.Rows(0).Item("ProductID"), 0)
                        'productParameterList.Add(productParameter)
                    Else
                        If MyCommon.Fetch_SystemOption(150) = 0 Then
                            RetCode = StatusCodes.INVALID_PRODUCTID
                            RetMsg = "Failure. Product " & ExtProductID & " cannot be added to group " & ProductGroupID & " because it does not exist"
                            Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                            Return
                        Else
                            'Create New Product
                            If Operation.ToLower <> "remove" Then
                                If MyCommon.Fetch_SystemOption(97) = 1 Then
                                    If Not IsNumeric(ExtProductID) Then
                                        RetCode = StatusCodes.INVALID_PRODUCTID
                                        RetMsg = "Failure. Cannot add item " & ExtProductID & " to group because it contains non-numeric characters"
                                        Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                                        Return
                                    End If
                                    If CLng(ExtProductID) <= 0 Then
                                        RetCode = StatusCodes.INVALID_PRODUCTID
                                        RetMsg = "Failure. Cannot add item " & ExtProductID & " to group because produtct id must be positive."
                                        Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                                        Return
                                    End If

                                End If

                                MyCommon.QueryStr = "dbo.pa_PUA_UpdateProduct"
                                MyCommon.Open_LRTsp()
                                MyCommon.LRTsp.Parameters.Add("@ExtProductID", SqlDbType.NVarChar, 120).Value = ExtProductID
                                MyCommon.LRTsp.Parameters.Add("@ProductTypeID", SqlDbType.Int).Value = ProductTypeID
                                MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 200).Value = ProductDesc
                                MyCommon.LRTsp.Parameters.Add("@ProductID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                                MyCommon.LRTsp.ExecuteNonQuery()
                                ProductID = MyCommon.LRTsp.Parameters("@ProductID").Value
                                MyCommon.Close_LRTsp()
                                Copient.Logger.Write_Log(LGMLogFile, "Created product " & ExtProductID, True)
                                'Add to table
                                querystr &= "insert into #TempProdPK values(" & ProductID & "); "
                                'productParameter = New SqlParameter("@ProdID" & ExtProductID, SqlDbType.BigInt)
                                'productParameter.Value = ProductID
                                'productParameterList.Add(productParameter)
                            Else
                                RetCode = StatusCodes.INVALID_PRODUCTID
                                RetMsg = "Failure. Cannot remove item: " & ExtProductID & " because it does not exist"
                                Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
                                Return
                            End If
                        End If
                    End If
                Next
                'For Each param In productParameterList
                '    MyCommon.DBParameters.Add(param)
                'Next
                MyCommon.QueryStr = querystr
                MyCommon.LRT_Execute()
                If Operation = "augment" Then

                    MyCommon.QueryStr = "dbo.pt_LGM_ProductGroup_Insert"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt, 20).Value = ProductGroupID
                    MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
                    MyCommon.LRTsp.CommandTimeout = 2400
                    MyCommon.LRTsp.ExecuteNonQuery()
                    Status = MyCommon.LRTsp.Parameters("@Status").Value
                    MyCommon.Close_LRTsp()
                    MyCommon.QueryStr = " drop table #TempProdPK;"
                    MyCommon.LRT_Execute()
                    If Status = 0 Or Status = -1 Then
                        RetCode = StatusCodes.SUCCESS
                        RetMsg = "Product Successfully added to Product Group " & ProductGroupID
                    ElseIf Status = -2 Then
                        'Since the cards are checked earlier, it should never get to this point. 
                        RetCode = StatusCodes.INVALID_INCOMPLETEPROCESSLOCDATA
                        RetMsg = "Failure. Tried to add the product that does not exist to the Product Group. No products were added"
                    End If
                ElseIf Operation = "replace" Then

                    MyCommon.QueryStr = "dbo.pt_LGM_ProductGroup_Replace"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt, 20).Value = ProductGroupID
                    MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
                    MyCommon.LRTsp.CommandTimeout = 2400
                    MyCommon.LRTsp.ExecuteNonQuery()
                    Status = MyCommon.LRTsp.Parameters("@Status").Value
                    MyCommon.Close_LRTsp()
                    MyCommon.QueryStr = " drop table #TempProdPK;"
                    MyCommon.LRT_Execute()
                    If Status = 0 Then
                        RetCode = StatusCodes.SUCCESS
                        RetMsg = "Successfully replaces products in Product Group " & ProductGroupID
                    ElseIf Status = -2 Then
                        'Since the cards are checked earlier, it should never get to this point. 
                        RetCode = StatusCodes.INVALID_INCOMPLETEPROCESSLOCDATA
                        RetMsg = "Failure. Tried to add a non-existent product to the Product Group. No products were replaced"
                    End If
                ElseIf Operation = "remove" Then

                    MyCommon.QueryStr = "dbo.pt_LGM_ProductGroup_Remove"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt, 20).Value = ProductGroupID
                    MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
                    MyCommon.LRTsp.CommandTimeout = 2400
                    MyCommon.LRTsp.ExecuteNonQuery()
                    Status = MyCommon.LRTsp.Parameters("@Status").Value
                    MyCommon.Close_LRTsp()
                    MyCommon.QueryStr = " drop table #TempProdPK;"
                    MyCommon.LRT_Execute()
                    If Status = 0 Then
                        RetCode = StatusCodes.SUCCESS
                        RetMsg = "Successfully removed products from Product Group " & ProductGroupID
                    ElseIf Status = -1 Then
                        'Since the cards are checked earlier, it should never get to this point. 
                        RetCode = StatusCodes.INVALID_INCOMPLETEPROCESSLOCDATA
                        RetMsg = "Failure. Some products were not in Product Group. No products were removed"
                    End If
                End If

            Else
                RetCode = StatusCodes.INVALID_INCOMPLETEPROCESSLOCDATA
                RetMsg = "Failure. BulkData is empty"
            End If
            '	End Using
        Catch ex As Exception
            RetCode = StatusCodes.APPLICATION_EXCEPTION
            RetMsg = ex.ToString
        Finally
            MyCommon.QueryStr = "IF OBJECT_ID('tempdb..#TempProdPK') IS NOT NULL BEGIN drop table #TempProdPK END "
            MyCommon.LRT_Execute()
            If MyCommon.LRTadoConn.State = ConnectionState.Open Then MyCommon.Close_LogixRT()

        End Try
        Copient.Logger.Write_Log(LGMLogFile, RetMsg, True)
    End Sub
#End Region

    <WebMethod()> _
    Public Function GetCustomerListByGroupID(ByVal GUID As String, ByVal CustomerGroupID As String, ByVal CardTypeID As String) As XmlDocument
        Dim SessionXml As New XmlDocument()
        Dim ms As New MemoryStream
        Dim Writer As New XmlTextWriter(ms, System.Text.Encoding.UTF8)
        Dim rst As DataTable
        Dim row As DataRow
        Dim ErrorXML As String = ""
        Dim query As String = ""
        Dim iCardTypeID As Integer = Nothing
        Dim maxCustomerRet As Integer = Nothing
        Dim MsgBuf As New StringBuilder()
        Try

            Dim iCustomerGroupID As Integer = -1
            Try
                iCustomerGroupID = CInt(CustomerGroupID)
            Catch ex As Exception
                iCustomerGroupID = -1
            End Try

            If IsValidGUID(GUID, "GetCustomerList") Then
                If CardTypeID = "" OrElse Integer.TryParse(CardTypeID, iCardTypeID) Then
                    Writer.Formatting = Formatting.Indented
                    Writer.Indentation = 4

                    Writer.WriteStartDocument()
                    Writer.WriteStartElement("CustomerGroupList")
                    Writer.WriteAttributeString("returnCode", "SUCCESS")
                    Writer.WriteAttributeString("responseTime", Date.Now.ToString("yyyy-MM-ddTHH:mm:ss"))
                    Writer.WriteAttributeString("xmlns", "xsi", Nothing, "http://www.w3.org/2001/XMLSchema-instance")

                    ''ct = MyCommon.NZ(ct,-1)
                    Using MyCommon.LRTadoConn
                        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
                        MyCommon.QueryStr = "select CustomerGroupID, Name, ExtGroupID from CustomerGroups where Deleted = 0 and CustomerGroupID = @CustomerGroupID"
                        MyCommon.DBParameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = iCustomerGroupID
                        rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                    End Using

          If (rst.Rows.Count > 0) Then
            Using MyCommon.LXSadoConn
              If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
              Writer.WriteElementString("CustomerGroupID", iCustomerGroupID)
              Writer.WriteElementString("ExtGroupID", MyCommon.NZ(rst.Rows(0).Item("ExtGroupID"), ""))
              Writer.WriteElementString("Name", MyCommon.NZ(rst.Rows(0).Item("Name"), ""))
              Writer.WriteStartElement("CardIDs")
				try
					maxCustomerRet = MyCommon.Fetch_InterfaceOption(105)
				catch e as Exception
					maxCustomerRet = 2000000000
				end try
                            query = "select CARDS.ExtCardIDOriginal as ExtCardID, CARDS.CardTypeID from CardIDs CARDS with (NoLock) " & _
                       "inner join  (select top " & maxCustomerRet & " CustomerPK, MembershipID from GroupMembership with (NoLock)" & _
                       "where CustomerGroupID=@CustomerGroupID and Deleted=0) as GM on GM.CustomerPK = CARDS.CustomerPK"
              
              MyCommon.DBParameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = iCustomerGroupID
              If (CardTypeID <> "") Then
                query = query & " where CardTypeID = @CardTypeID"
                MyCommon.DBParameters.Add("@CardTypeID", SqlDbType.Int).Value = iCardTypeID
              End If
              MyCommon.QueryStr = query
              rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)
              If (rst.Rows.Count > 0) Then
                For Each row In rst.Rows
                  Writer.WriteStartElement("Card")
                                    Writer.WriteElementString("ExtCardID", MyCryptLib.SQL_StringDecrypt(row.Item("ExtCardID").ToString()))
                                    Writer.WriteElementString("CardTypeID", MyCommon.NZ(row.Item("CardTypeID"), ""))
                                    Writer.WriteEndElement() 'Card
                                Next
                                MsgBuf.Append("Returned " & rst.Rows.Count & " Customer IDs from GroupID " & CustomerGroupID)
                            End If
                            Writer.WriteEndElement() 'CardIDs
                        End Using
                    Else
                        ' customer group not found
                        ErrorXML = GetCustomerListErrorXML(StatusCodes.INVALID_CUSTOMERGROUPID, "Customer Group ID " & CustomerGroupID & " was not found.")
                    End If

                    Writer.WriteEndElement() ' end customer group
                    Writer.WriteEndDocument()
                    Writer.Flush()
                Else
                    ' Send back Invalid CardTypeID return code
                    ErrorXML = GetCustomerListErrorXML(StatusCodes.INVALID_CARDTYPEID, "CardTypeID '" & CardTypeID & "' is invalid.")
                End If
            Else
                ' Send back Invalid GUID return code
                ErrorXML = GetCustomerListErrorXML(StatusCodes.INVALID_GUID, "GUID: " & GUID & " is invalid")
            End If

        Catch ex As Exception
            ErrorXML = GetCustomerListErrorXML(StatusCodes.APPLICATION_EXCEPTION, ex.ToString())
        Finally
            If ErrorXML = "" Then
                ms.Seek(0, SeekOrigin.Begin)
                SessionXml.Load(ms)
            Else
                SessionXml.LoadXml(ErrorXML)
            End If

            If Not Writer Is Nothing Then
                Writer.Close()
            End If
            If Not ms Is Nothing Then
                ms.Close()
                ms.Dispose()
                ms = Nothing
            End If

            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try
        Copient.Logger.Write_Log(LGMLogFile, MsgBuf.ToString, True)
        Return SessionXml
    End Function

    <WebMethod()> _
    Public Function GetCustomerListByExtGroupID(ByVal GUID As String, ByVal ExtGroupID As String, ByVal Name As String, ByVal CardTypeID As String) As XmlDocument
        Dim SessionXml As New XmlDocument()
        Dim ms As New MemoryStream
        Dim Writer As New XmlTextWriter(ms, System.Text.Encoding.UTF8)
        Dim rst As DataTable
        Dim row As DataRow
        Dim ErrorXML As String = ""
        Dim query As String
        Dim iCardTypeID As Integer = Nothing
        Dim maxCustomerRet As Integer = Nothing
        Dim MsgBuf As New StringBuilder()
        Try



            If IsValidGUID(GUID, "GetCustomerList") Then
                If (CardTypeID = "" OrElse Integer.TryParse(CardTypeID, iCardTypeID)) Then
                    Writer.Formatting = Formatting.Indented
                    Writer.Indentation = 4

                    Writer.WriteStartDocument()
                    Writer.WriteStartElement("CustomerGroupList")
                    Writer.WriteAttributeString("returnCode", "SUCCESS")
                    Writer.WriteAttributeString("responseTime", Date.Now.ToString("yyyy-MM-ddTHH:mm:ss"))
                    Writer.WriteAttributeString("xmlns", "xsi", Nothing, "http://www.w3.org/2001/XMLSchema-instance")
                    Using MyCommon.LRTadoConn
                        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
                        query = "select CustomerGroupID, Name, ExtGroupID from CustomerGroups where Deleted = 0 "
                        If (ExtGroupID <> "") Then
                            query = query & "and ExtGroupID = @ExtGroupID"
                            MyCommon.DBParameters.Add("@ExtGroupID", SqlDbType.NVarChar).Value = ExtGroupID
                        End If

                        If (Name <> "") Then
                            query = query & " and Name = @Name"
                            MyCommon.DBParameters.Add("@Name", SqlDbType.NVarChar).Value = Name
                        End If

                        If (ExtGroupID = "" And Name = "") Then
                            query = query & " and CustomerGroupID = -1" 'return 0 rows
                        End If

                        MyCommon.QueryStr = query
                        rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                    End Using

			If (rst.Rows.Count > 0) Then
			  Writer.WriteElementString("CustomerGroupID", rst.Rows(0).Item("CustomerGroupID"))
			  Writer.WriteElementString("ExtGroupID", MyCommon.NZ(rst.Rows(0).Item("ExtGroupID"), ""))
			  Writer.WriteElementString("Name", MyCommon.NZ(rst.Rows(0).Item("Name"), ""))
			  Writer.WriteStartElement("CardIDs")
				try
					maxCustomerRet = MyCommon.Fetch_InterfaceOption(105)
				catch e as Exception
					maxCustomerRet = 2000000000
				end try
			  Using MyCommon.LXSadoConn
				If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
                            query = "select CARDS.ExtCardIDOriginal as ExtCardID, CARDS.CardTypeID from CardIDs CARDS with (NoLock) " & _
                         "inner join  (select top " & maxCustomerRet & " CustomerPK, MembershipID from GroupMembership with (NoLock) " & _
                         "where CustomerGroupID=@CustomerGroupID and Deleted=0) as GM on GM.CustomerPK = CARDS.CustomerPK"
				MyCommon.DBParameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = MyCommon.NZ(rst.Rows(0).Item("CustomerGroupID"), -1)
				If (CardTypeID <> "") Then
				  query = query & " where CardTypeID = @CardTypeID"
				  MyCommon.DBParameters.Add("@CardTypeID", SqlDbType.Int).Value = CardTypeID
				End If
				MyCommon.QueryStr = query
				rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)
				If (rst.Rows.Count > 0) Then
				  For Each row In rst.Rows
					Writer.WriteStartElement("Card")
                                    Writer.WriteElementString("ExtCardID", MyCryptLib.SQL_StringDecrypt(row.Item("ExtCardID").ToString()))
                                    Writer.WriteElementString("CardTypeID", MyCommon.NZ(row.Item("CardTypeID"), ""))
                                    Writer.WriteEndElement() 'Card
                                Next
                                MsgBuf.Append("Returned " & rst.Rows.Count & " Customer IDs from ExtGroupID " & ExtGroupID)
                            End If
                        End Using
                        Writer.WriteEndElement() 'CardIDs
                    Else
                        ' customer group not found
                        ErrorXML = GetCustomerListErrorXML(StatusCodes.INVALID_CUSTOMERGROUPNAME, "Customer Group with Name: " & Name & " and/or ExtGroupID: " & ExtGroupID & " was not found.")
                    End If

                    Writer.WriteEndElement() ' end customer group
                    Writer.WriteEndDocument()
                    Writer.Flush()
                Else
                    ' Send back Invalid CardTypeID return code
                    ErrorXML = GetCustomerListErrorXML(StatusCodes.INVALID_CARDTYPEID, "CardTypeID '" & CardTypeID & "' is invalid.")
                End If
            Else
                ' Send back Invalid GUID return code
                ErrorXML = GetCustomerListErrorXML(StatusCodes.INVALID_GUID, "GUID: " & GUID & " is invalid")
            End If

            If ErrorXML = "" Then
                ms.Seek(0, SeekOrigin.Begin)
                SessionXml.Load(ms)
            Else
                SessionXml.LoadXml(ErrorXML)
            End If

        Catch ex As Exception
            ErrorXML = GetCustomerListErrorXML(StatusCodes.APPLICATION_EXCEPTION, ex.ToString())
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
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try
        Copient.Logger.Write_Log(LGMLogFile, MsgBuf.ToString, True)
        Return SessionXml
    End Function

    <WebMethod()> _
    Public Function GetProductListByGroupID(ByVal GUID As String, ByVal ProductGroupID As String) As XmlDocument
        Dim SessionXml As New XmlDocument()
        Dim ms As New MemoryStream
        Dim Writer As New XmlTextWriter(ms, System.Text.Encoding.UTF8)
        Dim rst As DataTable
        Dim row As DataRow
        Dim ErrorXML As String = ""
        Try

            Dim iProductGroupID As Integer = -1
            Try
                iProductGroupID = CInt(ProductGroupID)
            Catch ex As Exception
                iProductGroupID = -1
            End Try

            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            If IsValidGUID(GUID, "GetProductList") Then
                Writer.Formatting = Formatting.Indented
                Writer.Indentation = 4

                Writer.WriteStartDocument()
                Writer.WriteStartElement("ProductGroupList")
                Writer.WriteAttributeString("returnCode", "SUCCESS")
                Writer.WriteAttributeString("responseTime", Date.Now.ToString("yyyy-MM-ddTHH:mm:ss"))
                Writer.WriteAttributeString("xmlns", "xsi", Nothing, "http://www.w3.org/2001/XMLSchema-instance")

                Using MyCommon.LRTadoConn
                    If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
                    MyCommon.QueryStr = "select ProductGroupID, Name, ExtGroupID from ProductGroups where Deleted = 0 and ProductGroupID = @ProductGroupID"
                    MyCommon.DBParameters.Add("@ProductGroupID", SqlDbType.BigInt).Value = iProductGroupID
                    rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)

                    If (rst.Rows.Count > 0) Then
                        Writer.WriteElementString("ProductGroupID", iProductGroupID)
                        Writer.WriteElementString("ExtGroupID", MyCommon.NZ(rst.Rows(0).Item("ExtGroupID"), ""))
                        Writer.WriteElementString("Name", MyCommon.NZ(rst.Rows(0).Item("Name"), ""))
                        Writer.WriteStartElement("Products")

                        MyCommon.QueryStr = "select P.ExtProductID, P.ProductTypeID from Products P with (NoLock) " & _
                                 "inner join (select ProductGroupID, ProductID from ProdGroupItems with (NoLock) " & _
                                 "where ProductGroupID = @ProductGroupID and Deleted = 0) as PGI on PGI.ProductID = P.ProductID"
                        MyCommon.DBParameters.Add("@ProductGroupID", SqlDbType.BigInt).Value = iProductGroupID
                        rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)

                        If (rst.Rows.Count > 0) Then
                            For Each row In rst.Rows
                                Writer.WriteStartElement("Item")
                                Writer.WriteElementString("ExtProductID", MyCommon.NZ(row.Item("ExtProductID"), ""))
                                Writer.WriteElementString("ProductTypeID", MyCommon.NZ(row.Item("ProductTypeID"), ""))
                                Writer.WriteEndElement() 'Item
                            Next
                        End If
                        Writer.WriteEndElement() 'Products
                    Else
                        ' Product group not found
                        ErrorXML = GetProductListErrorXML(StatusCodes.INVALID_PRODGROUPID, "Product Group ID " & ProductGroupID & " was not found.")
                    End If
                End Using
                Writer.WriteEndElement() ' end Product group
                Writer.WriteEndDocument()
                Writer.Flush()

            Else
                ' Send back Invalid GUID return code
                ErrorXML = GetProductListErrorXML(StatusCodes.INVALID_GUID, "GUID: " & GUID & " is invalid")
            End If

            If ErrorXML = "" Then
                ms.Seek(0, SeekOrigin.Begin)
                SessionXml.Load(ms)
            Else
                SessionXml.LoadXml(ErrorXML)
            End If

        Catch ex As Exception
            ErrorXML = GetProductListErrorXML(StatusCodes.APPLICATION_EXCEPTION, ex.ToString())
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
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try

        Return SessionXml
    End Function

    <WebMethod()> _
    Public Function GetProductListByExtGroupID(ByVal GUID As String, ByVal ExtGroupID As String, ByVal Name As String) As XmlDocument
        Dim SessionXml As New XmlDocument()
        Dim ms As New MemoryStream
        Dim Writer As New XmlTextWriter(ms, System.Text.Encoding.UTF8)
        Dim rst As DataTable
        Dim row As DataRow
        Dim ErrorXML As String = ""
        Dim query As String
        Try

            If IsValidGUID(GUID, "GetProductList") Then
                Writer.Formatting = Formatting.Indented
                Writer.Indentation = 4

                Writer.WriteStartDocument()
                Writer.WriteStartElement("ProductGroupList")
                Writer.WriteAttributeString("returnCode", "SUCCESS")
                Writer.WriteAttributeString("responseTime", Date.Now.ToString("yyyy-MM-ddTHH:mm:ss"))
                Writer.WriteAttributeString("xmlns", "xsi", Nothing, "http://www.w3.org/2001/XMLSchema-instance")

                Using MyCommon.LRTadoConn
                    If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
                    query = "select ProductGroupID, Name, ExtGroupID from ProductGroups where Deleted = 0 "
                    If (ExtGroupID <> "") Then
                        query = query & "and ExtGroupID = @ExtGroupID"
                        MyCommon.DBParameters.Add("@ExtGroupID", SqlDbType.NVarChar).Value = ExtGroupID
                    End If

                    If (Name <> "") Then
                        query = query & " and Name = @Name"
                        MyCommon.DBParameters.Add("@Name", SqlDbType.NVarChar).Value = Name
                    End If

                    If (ExtGroupID = "" And Name = "") Then
                        query = query & "and ProductGroupID = -1" 'return 0 rows
                    End If
                    MyCommon.QueryStr = query
                    rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)


                    If (rst.Rows.Count > 0) Then
                        Writer.WriteElementString("ProductGroupID", rst.Rows(0).Item("ProductGroupID"))
                        Writer.WriteElementString("ExtGroupID", MyCommon.NZ(rst.Rows(0).Item("ExtGroupID"), ""))
                        Writer.WriteElementString("Name", MyCommon.NZ(rst.Rows(0).Item("Name"), ""))
                        Writer.WriteStartElement("Products")

                        MyCommon.QueryStr = "select P.ExtProductID, P.ProductTypeID from Products P with (NoLock) " & _
                                 "inner join (select ProductGroupID, ProductID from ProdGroupItems with (NoLock) " & _
                                 "where ProductGroupID = @ProductGroupID and Deleted = 0) as PGI on PGI.ProductID = P.ProductID"
                        MyCommon.DBParameters.Add("@ProductGroupID", SqlDbType.BigInt).Value = rst.Rows(0).Item("ProductGroupID")
                        rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)

                        If (rst.Rows.Count > 0) Then
                            For Each row In rst.Rows
                                Writer.WriteStartElement("Item")
                                Writer.WriteElementString("ExtProductID", MyCommon.NZ(row.Item("ExtProductID"), ""))
                                Writer.WriteElementString("ProductTypeID", MyCommon.NZ(row.Item("ProductTypeID"), ""))
                                Writer.WriteEndElement() 'Item
                            Next
                        End If
                        Writer.WriteEndElement() 'Products
                    Else
                        ' Product group not found
                        ErrorXML = GetProductListErrorXML(StatusCodes.INVALID_PRODGROUPNAME, "Product Group with Name: " & Name & " or ExtGroupID: " & ExtGroupID & " was not found.")
                    End If
                End Using
                Writer.WriteEndElement() ' end Product group
                Writer.WriteEndDocument()
                Writer.Flush()

            Else
                ' Send back Invalid GUID return code
                ErrorXML = GetProductListErrorXML(StatusCodes.INVALID_GUID, "GUID: " & GUID & " is invalid")
            End If

            If ErrorXML = "" Then
                ms.Seek(0, SeekOrigin.Begin)
                SessionXml.Load(ms)
            Else
                SessionXml.LoadXml(ErrorXML)
            End If

        Catch ex As Exception
            ErrorXML = GetProductListErrorXML(StatusCodes.APPLICATION_EXCEPTION, ex.ToString())
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
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try

        Return SessionXml
    End Function

    <WebMethod()> _
    Public Function GetLocationListByGroupID(ByVal GUID As String, ByVal LocationGroupID As String) As XmlDocument
        Dim SessionXml As New XmlDocument()
        Dim ms As New MemoryStream
        Dim Writer As New XmlTextWriter(ms, System.Text.Encoding.UTF8)
        Dim rst As DataTable
        Dim row As DataRow
        Dim ErrorXML As String = ""
        Try
            Dim iLocationGroupID As Integer = -1
            Try
                iLocationGroupID = CInt(LocationGroupID)
            Catch ex As Exception
                iLocationGroupID = -1
            End Try

            If IsValidGUID(GUID, "GetProductList") Then
                Writer.Formatting = Formatting.Indented
                Writer.Indentation = 4

                Writer.WriteStartDocument()
                Writer.WriteStartElement("LocationGroupList")
                Writer.WriteAttributeString("returnCode", "SUCCESS")
                Writer.WriteAttributeString("responseTime", Date.Now.ToString("yyyy-MM-ddTHH:mm:ss"))
                Writer.WriteAttributeString("xmlns", "xsi", Nothing, "http://www.w3.org/2001/XMLSchema-instance")
                Using MyCommon.LRTadoConn
                    If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
                    MyCommon.QueryStr = "select LocationGroupID, Name, ExtGroupID from LocationGroups where Deleted = 0 and LocationGroupID = @LocationGroupID"
                    MyCommon.DBParameters.Add("@LocationGroupID", SqlDbType.BigInt).Value = iLocationGroupID
                    rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                    If (rst.Rows.Count > 0) Then
                        Writer.WriteElementString("LocationGroupID", iLocationGroupID)
                        Writer.WriteElementString("ExtGroupID", MyCommon.NZ(rst.Rows(0).Item("ExtGroupID"), ""))
                        Writer.WriteElementString("Name", MyCommon.NZ(rst.Rows(0).Item("Name"), ""))
                        Writer.WriteStartElement("Locations")
                        MyCommon.QueryStr = "select loc.LocationID, loc.ExtLocationCode, loc.LocationName from Locations loc with (NoLock) " & _
                                 "inner join (select LocationGroupID, LocationID from LocGroupItems with (NoLock) " & _
                                 "where LocationGroupID = @LocationGroupID and Deleted =0) as LGI on LGI.LocationID = loc.LocationID"
                        MyCommon.DBParameters.Add("@LocationGroupID", SqlDbType.BigInt).Value = iLocationGroupID
                        rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                        If (rst.Rows.Count > 0) Then
                            For Each row In rst.Rows
                                Writer.WriteStartElement("Store")
                                Writer.WriteElementString("LocationID", MyCommon.NZ(row.Item("LocationID"), ""))
                                Writer.WriteElementString("ExtLocationCode", MyCommon.NZ(row.Item("ExtLocationCode"), ""))
                                Writer.WriteElementString("LocationName", MyCommon.NZ(row.Item("LocationName"), ""))
                                Writer.WriteEndElement() 'Store
                            Next
                        End If
                        Writer.WriteEndElement() 'Locations
                    Else
                        ' Location group not found
                        ErrorXML = GetLocationListErrorXML(StatusCodes.INVALID_LOCATIONGROUPID, "Location Group ID " & LocationGroupID & " was not found.")
                    End If
                End Using
                Writer.WriteEndElement() ' end Location group List
                Writer.WriteEndDocument()
                Writer.Flush()

            Else
                ' Send back Invalid GUID return code
                ErrorXML = GetLocationListErrorXML(StatusCodes.INVALID_GUID, "GUID: " & GUID & " is invalid")
            End If

            If ErrorXML = "" Then
                ms.Seek(0, SeekOrigin.Begin)
                SessionXml.Load(ms)
            Else
                SessionXml.LoadXml(ErrorXML)
            End If

        Catch ex As Exception
            ErrorXML = GetLocationListErrorXML(StatusCodes.APPLICATION_EXCEPTION, ex.ToString())
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
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try

        Return SessionXml
    End Function

    <WebMethod()> _
    Public Function GetLocationListByExtGroupID(ByVal GUID As String, ByVal ExtGroupID As String, ByVal Name As String) As XmlDocument
        Dim SessionXml As New XmlDocument()
        Dim ms As New MemoryStream
        Dim Writer As New XmlTextWriter(ms, System.Text.Encoding.UTF8)
        Dim rst As DataTable
        Dim row As DataRow
        Dim ErrorXML As String = ""
        Dim query As String
        Try


            If IsValidGUID(GUID, "GetProductList") Then
                Writer.Formatting = Formatting.Indented
                Writer.Indentation = 4

                Writer.WriteStartDocument()
                Writer.WriteStartElement("LocationGroupList")
                Writer.WriteAttributeString("returnCode", "SUCCESS")
                Writer.WriteAttributeString("responseTime", Date.Now.ToString("yyyy-MM-ddTHH:mm:ss"))
                Writer.WriteAttributeString("xmlns", "xsi", Nothing, "http://www.w3.org/2001/XMLSchema-instance")

                Using MyCommon.LRTadoConn
                    If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
                    query = "select LocationGroupID, Name, ExtGroupID from LocationGroups where Deleted = 0 "
                    If (ExtGroupID <> "") Then
                        query = query & "and ExtGroupID = @ExtGroupID"
                        MyCommon.DBParameters.Add("@ExtGroupID", SqlDbType.NVarChar, 20).Value = ExtGroupID
                    End If

                    If (Name <> "") Then
                        query = query & " and Name = @Name "
                        MyCommon.DBParameters.Add("@Name", SqlDbType.NVarChar, 200).Value = Name
                    End If

                    If (ExtGroupID = "" And Name = "") Then
                        query = query & "and LocationGroupID = -1" 'return 0 rows
                    End If

                    MyCommon.QueryStr = query
                    rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                    If (rst.Rows.Count > 0) Then
                        Writer.WriteElementString("LocationGroupID", rst.Rows(0).Item("LocationGroupID"))
                        Writer.WriteElementString("ExtGroupID", MyCommon.NZ(rst.Rows(0).Item("ExtGroupID"), ""))
                        Writer.WriteElementString("Name", MyCommon.NZ(rst.Rows(0).Item("Name"), ""))
                        Writer.WriteStartElement("Locations")
                        MyCommon.QueryStr = "select loc.LocationID, loc.ExtLocationCode, loc.LocationName from Locations loc with (NoLock) " & _
                                 "inner join (select LocationGroupID, LocationID from LocGroupItems with (NoLock) " & _
                                 "where LocationGroupID = @LocationGroupID and Deleted =0) as LGI on LGI.LocationID = loc.LocationID"
                        MyCommon.DBParameters.Add("@LocationGroupID", SqlDbType.BigInt).Value = rst.Rows(0).Item("LocationGroupID")
                        rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                        If (rst.Rows.Count > 0) Then
                            For Each row In rst.Rows
                                Writer.WriteStartElement("Store")
                                Writer.WriteElementString("LocationID", MyCommon.NZ(row.Item("LocationID"), ""))
                                Writer.WriteElementString("ExtLocationCode", MyCommon.NZ(row.Item("ExtLocationCode"), ""))
                                Writer.WriteElementString("LocationName", MyCommon.NZ(row.Item("LocationName"), ""))
                                Writer.WriteEndElement() 'Store
                            Next
                        End If
                        Writer.WriteEndElement() 'Locations
                    Else
                        ' Location group not found
                        ErrorXML = GetLocationListErrorXML(StatusCodes.INVALID_LOCATIONNAME, "Location Group with Name: " & Name & " or ExtGroupID: " & ExtGroupID & " was not found.")
                    End If
                End Using
                Writer.WriteEndElement() ' end Location group List
                Writer.WriteEndDocument()
                Writer.Flush()

            Else
                ' Send back Invalid GUID return code
                ErrorXML = GetLocationListErrorXML(StatusCodes.INVALID_GUID, "GUID: " & GUID & " is invalid")
            End If

            If ErrorXML = "" Then
                ms.Seek(0, SeekOrigin.Begin)
                SessionXml.Load(ms)
            Else
                SessionXml.LoadXml(ErrorXML)
            End If

        Catch ex As Exception
            ErrorXML = GetLocationListErrorXML(StatusCodes.APPLICATION_EXCEPTION, ex.ToString())
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
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try

        Return SessionXml
    End Function

    'This webmethod  returns the external group id as well as the logix group id, given the name of the customer group. 
    <WebMethod()> _
    Public Function GetCustomerGroupIDByName(ByVal GUID As String, ByVal CustomerGroupName As String) As XmlDocument
        Dim SessionXml As New XmlDocument()
        Dim ms As New MemoryStream
        Dim Writer As New XmlTextWriter(ms, System.Text.Encoding.UTF8)
        Dim rst As DataTable
        Dim ErrorXML As String = ""
        Dim query As String
        Dim iCardTypeID As Integer = Nothing
        Dim maxCustomerRet As Integer = Nothing
        Dim MsgBuf As New StringBuilder()
        Dim ExtGroupID As String

        Try

            If IsValidGUID(GUID, "GetCustomerList") Then
                Writer.Formatting = Formatting.Indented
                Writer.Indentation = 4
                Writer.WriteStartDocument()
                Writer.WriteStartElement("CustomerGroupList")
                Writer.WriteAttributeString("returnCode", "SUCCESS")
                Writer.WriteAttributeString("responseTime", Date.Now.ToString("yyyy-MM-ddTHH:mm:ss"))
                Writer.WriteAttributeString("xmlns", "xsi", Nothing, "http://www.w3.org/2001/XMLSchema-instance")
                Using MyCommon.LRTadoConn
                    If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
                    query = "select CustomerGroupID, ExtGroupID from CustomerGroups where Name=@Name and Deleted = 0 "
                    MyCommon.DBParameters.Add("@Name", SqlDbType.NVarChar).Value = CustomerGroupName
                    MyCommon.QueryStr = query
                    rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                End Using

                If (rst.Rows.Count > 0) Then
                    ExtGroupID = MyCommon.NZ(rst.Rows(0).Item("ExtGroupID"), "")
                    Writer.WriteElementString("CustomerGroupID", rst.Rows(0).Item("CustomerGroupID"))
                    Writer.WriteElementString("ExtGroupID", ExtGroupID)
                Else
                    ' customer group not found
                    ErrorXML = GetCustomerListErrorXML(StatusCodes.INVALID_CUSTOMERGROUPNAME, "Customer Group with Name: " & CustomerGroupName & " was not found.")
                End If

                Writer.WriteEndElement() ' end customer group
                Writer.WriteEndDocument()
                Writer.Flush()

            Else
                ' Send back Invalid GUID return code
                ErrorXML = GetCustomerListErrorXML(StatusCodes.INVALID_GUID, "GUID: " & GUID & " is invalid")
            End If

            If ErrorXML = "" Then
                ms.Seek(0, SeekOrigin.Begin)
                SessionXml.Load(ms)
            Else
                SessionXml.LoadXml(ErrorXML)
            End If

        Catch ex As Exception
            ErrorXML = GetCustomerListErrorXML(StatusCodes.APPLICATION_EXCEPTION, ex.ToString())
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
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try
        Copient.Logger.Write_Log(LGMLogFile, MsgBuf.ToString, True)
        Return SessionXml
    End Function

    'This web method returns all location groups defined in the system, along with the associated locations. 
    <WebMethod()> _
    Public Function GetAllLocationGroupDetails(ByVal GUID As String) As XmlDocument
        Dim SessionXml As New XmlDocument()
        Dim ms As New MemoryStream
        Dim Writer As New XmlTextWriter(ms, System.Text.Encoding.UTF8)
        Dim rst As DataTable
        Dim lgrpRow As DataRow
        Dim ErrorXML As String = ""
        Dim CurrLocationGroupID As String = ""
        Dim PrevLocationGroupID As String = ""
        Dim maxLocationsRet As String = ""
        Dim topLimit As String = 1500
        Try

            If IsValidGUID(GUID, "GetProductList") Then
                Writer.Formatting = Formatting.Indented
                Writer.Indentation = 4
                Writer.WriteStartDocument()
                Writer.WriteStartElement("LocationGroupList")
                Writer.WriteAttributeString("returnCode", "SUCCESS")
                Writer.WriteAttributeString("responseTime", Date.Now.ToString("yyyy-MM-ddTHH:mm:ss"))
                Writer.WriteAttributeString("xmlns", "xsi", Nothing, "http://www.w3.org/2001/XMLSchema-instance")
                ' Try
                maxLocationsRet = MyCommon.Fetch_SystemOption(269)
                topLimit = If((maxLocationsRet <> String.Empty And maxLocationsRet <> "0"), "Top " & maxLocationsRet, String.Empty)
                Using MyCommon.LRTadoConn
                    If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

                    MyCommon.QueryStr = "select LocationGroupID,ExtGroupID , Name from LocationGroups   with (NoLock)  where LocationGroupID =1"
                    rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                    lgrpRow = rst.Rows(0)
                    Writer.WriteStartElement("LocationGroup")
                    Writer.WriteElementString("LocationGroupID", CurrLocationGroupID)
                    Writer.WriteElementString("ExtGroupID", MyCommon.NZ(lgrpRow.Item("ExtGroupID"), ""))
                    Writer.WriteElementString("Name", MyCommon.NZ(lgrpRow.Item("Name"), ""))
                    Writer.WriteStartElement("Locations")
                    MyCommon.QueryStr = "select " & topLimit & " LocationID, ExtLocationCode, LocationName from locations  with (NoLock)"

                    rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                    If (rst.Rows.Count > 0) Then

                        For Each lgrpRow In rst.Rows

                            Writer.WriteStartElement("Store")
                            Writer.WriteElementString("LocationID", MyCommon.NZ(lgrpRow.Item("LocationID"), ""))
                            Writer.WriteElementString("ExtLocationCode", MyCommon.NZ(lgrpRow.Item("ExtLocationCode"), ""))
                            Writer.WriteElementString("LocationName", MyCommon.NZ(lgrpRow.Item("LocationName"), ""))
                            Writer.WriteEndElement() 'Store
                        Next
                    Else
                        ' Location group not found
                        Writer.WriteElementString("Store", "No Stores")
                    End If

                End Using
                Writer.WriteEndElement() ' end Location group List
                Writer.WriteEndDocument()
                Writer.Flush()
            Else
                ' Send back Invalid GUID return code
                ErrorXML = GetLocationListErrorXML(StatusCodes.INVALID_GUID, "GUID: " & GUID & " is invalid")
            End If

            If ErrorXML = "" Then
                ms.Seek(0, SeekOrigin.Begin)
                SessionXml.Load(ms)
            Else
                SessionXml.LoadXml(ErrorXML)
            End If

        Catch ex As Exception
            ErrorXML = GetLocationListErrorXML(StatusCodes.APPLICATION_EXCEPTION, ex.ToString())
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
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try
        Return SessionXml
    End Function

    Private Function GetCustomerListErrorXML(ByVal Code As StatusCodes, Optional ByVal ErrorMsg As String = "", Optional ByVal RootNode As String = "") As String
        Dim ErrorXml As New StringBuilder("<?xml version=""1.0"" encoding=""utf-8""?>")
        If String.IsNullOrEmpty(RootNode) Then
            RootNode = "CustomerGroupList"
        End If
        ErrorXml.Append("<" & RootNode & " returnCode=""")
        Select Case Code
            Case StatusCodes.INVALID_GUID
                ErrorXml.Append("INVALID_GUID")
            Case StatusCodes.INVALID_CUSTOMERGROUPID
                ErrorXml.Append("CUSTOMER_GROUP_NOT_FOUND")
            Case StatusCodes.INVALID_CUSTOMERGROUPNAME
                ErrorXml.Append("CUSTOMER_GROUP_NOT_FOUND")
            Case StatusCodes.INVALID_XML_DOCUMENT
                ErrorXml.Append("INVALID_XML_DOCUMENT")
            Case StatusCodes.INVALID_CRITERIA_XML
                ErrorXml.Append("INVALID_CRITERIA_XML")
            Case StatusCodes.INVALID_CUSTOMERGROUPID
                ErrorXml.Append("INVALID_CUSTOMERGROUPID")
            Case StatusCodes.INVALID_OPERATIONTYPE
                ErrorXml.Append("INVALID_OPERATIONTYPE")
            Case StatusCodes.INVALID_CUSTOMERGROUPNAME
                ErrorXml.Append("INVALID_CUSTOMERGROUPNAME")
            Case StatusCodes.INVALID_INCOMPLETEPROCESSCUSTDATA
                ErrorXml.Append("INVALID_INCOMPLETEPROCESSCUSTDATA")
            Case StatusCodes.INVALID_CARDID
                ErrorXml.Append("INVALID_CARDID")
            Case StatusCodes.INVALID_CARDTYPEID
                ErrorXml.Append("INVALID_CARDTYPEID")
            Case Else
                ' treat everything else as an application exception
                ErrorXml.Append("APPLICATION_EXCEPTION")
        End Select
        ErrorXml.Append(""" responseTime=""" & Date.Now.ToString("yyyy-MM-ddTHH:mm:ss") & """ >")
        ErrorXml.Append("  <ErrorMessage>" & Escape(ErrorMsg) & "</ErrorMessage>")
        ErrorXml.Append("</" & RootNode & ">")

        Return ErrorXml.ToString
    End Function

    Private Function GetProductListErrorXML(ByVal Code As StatusCodes, Optional ByVal ErrorMsg As String = "" _
                                  , Optional ByVal RootNode As String = "") As String

        Dim ErrorXml As New StringBuilder("<?xml version=""1.0"" encoding=""utf-8""?>")

        If String.IsNullOrEmpty(RootNode) Then
            RootNode = "ProductGroupList"
        End If
        ErrorXml.Append("<" & RootNode & " returnCode=""")
        Select Case Code
            Case StatusCodes.INVALID_GUID
                ErrorXml.Append("INVALID_GUID")
            Case StatusCodes.INVALID_XML_DOCUMENT
                ErrorXml.Append("INVALID_XML_DOCUMENT")
            Case StatusCodes.INVALID_CRITERIA_XML
                ErrorXml.Append("INVALID_CRITERIA_XML")
            Case StatusCodes.INVALID_OPERATIONTYPE
                ErrorXml.Append("INVALID_OPERATIONTYPE")
            Case StatusCodes.INVALID_EXTPRODGROUPID
                ErrorXml.Append("INVALID_EXTPRODGROUPID")
            Case StatusCodes.INVALID_PRODGROUPNAME
                ErrorXml.Append("INVALID_PRODGROUPNAME")
            Case StatusCodes.INVALID_PRODUCTTYPEID
                ErrorXml.Append("INVALID_PRODUCTTYPEID")
            Case StatusCodes.INVALID_PRODGROUPID
                ErrorXml.Append("INVALID_PRODGROUPID")
            Case StatusCodes.INVALID_PRODUCTID
                ErrorXml.Append("INVALID_PRODUCTID")
            Case StatusCodes.INVALID_INCOMPLETEPROCESSPRODDATA
                ErrorXml.Append("INVALID_INCOMPLETEPROCESSPRODDATA")
            Case StatusCodes.FAILED_OPTIN
                ErrorXml.Append("FAILED_OPTIN")
            Case StatusCodes.INVALID_DESCRIPTION
                ErrorXml.Append("INVALID_DESCRIPTION")
            Case Else
                ' treat everything else as an application exception
                ErrorXml.Append("APPLICATION_EXCEPTION")
        End Select
        ErrorXml.Append(""" responseTime=""" & Date.Now.ToString("yyyy-MM-ddTHH:mm:ss") & """ >")
        ErrorXml.Append("  <ErrorMessage>" & ErrorMsg & "</ErrorMessage>")
        ErrorXml.Append("</" & RootNode & ">")
        Return ErrorXml.ToString
    End Function

    Private Function GetLocationListErrorXML(ByVal Code As StatusCodes, Optional ByVal ErrorMsg As String = "", Optional ByVal RootNode As String = "") As String
        Dim ErrorXml As New StringBuilder("<?xml version=""1.0"" encoding=""utf-8""?>")
        If String.IsNullOrEmpty(RootNode) Then
            RootNode = "LocationGroupList"
        End If
        ErrorXml.Append("<" & RootNode & " returnCode=""")
        Select Case Code
            Case StatusCodes.INVALID_GUID
                ErrorXml.Append("INVALID_GUID")
            Case StatusCodes.INVALID_LOCATIONGROUPID
                ErrorXml.Append("LOCATION_GROUP_NOT_FOUND")
            Case StatusCodes.INVALID_LOCATIONNAME
                ErrorXml.Append("LOCATION_GROUP_NOT_FOUND")
            Case StatusCodes.INVALID_XML_DOCUMENT
                ErrorXml.Append("INVALID_XML_DOCUMENT")
            Case StatusCodes.INVALID_CRITERIA_XML
                ErrorXml.Append("INVALID_CRITERIA_XML")
            Case StatusCodes.INVALID_LOCATIONGROUPID
                ErrorXml.Append("INVALID_LOCATIONGROUPID")
            Case StatusCodes.INVALID_OPERATIONTYPE
                ErrorXml.Append("INVALID_OPERATIONTYPE")
            Case StatusCodes.INVALID_BANNER
                ErrorXml.Append("INVALID_BANNER")
            Case StatusCodes.INVALID_INCOMPLETEPROCESSLOCDATA
                ErrorXml.Append("INVALID_INCOMPLETEPROCESSLOCDATA")
            Case Else
                ' treat everything else as an application exception
                ErrorXml.Append("APPLICATION_EXCEPTION")
        End Select
        ErrorXml.Append(""" responseTime=""" & Date.Now.ToString("yyyy-MM-ddTHH:mm:ss") & """ >")
        ErrorXml.Append("  <ErrorMessage>" & ErrorMsg & "</ErrorMessage>")
        ErrorXml.Append("</" & RootNode & ">")

        Return ErrorXml.ToString
    End Function

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub AddColumnToTable(ByRef DS As DataSet, ByVal TableName As String, _
                           ByVal ColumnName As String, ByVal ColType As System.Type)

        If Not DS.Tables(TableName).Columns.Contains(ColumnName) Then
            DS.Tables(TableName).Columns.Add(ColumnName, ColType)
        End If
    End Sub

    Private Function CreateReturnXMLDoc(ByRef DS As DataSet, ByVal xsdName As String) As XmlDocument
        Dim xmlDoc As XmlDocument = Nothing
        Dim Written As Boolean = False
        Dim sw As New System.IO.StringWriter
        'Dim schemas As New XmlSchemaSet
        Dim root As XmlElement
        Dim xsiNS As String = "http://www.w3.org/2001/XMLSchema-instance"

        Try
            If (DS IsNot Nothing AndAlso DS.Tables.Count > 0) Then
                For Each table As DataTable In DS.Tables
                    If table.Columns.Contains("statuscode") Then table.Columns("statuscode").ColumnMapping = MappingType.Attribute
                    If table.Columns.Contains("description") Then table.Columns("description").ColumnMapping = MappingType.Attribute
                Next

                DS.WriteXml(sw)
                xmlDoc = New XmlDocument
                xmlDoc.LoadXml(sw.ToString)

                root = xmlDoc.DocumentElement
                root.SetAttribute("xmlns:xsi", xsiNS)
                root.SetAttribute("noNamespaceSchemaLocation", xsiNS, xsdName)

            End If

        Catch ex As Exception
            xmlDoc = Nothing
        Finally
            If sw IsNot Nothing Then
                sw.Close()
                sw.Dispose()
            End If
        End Try

        Return xmlDoc
    End Function
End Class