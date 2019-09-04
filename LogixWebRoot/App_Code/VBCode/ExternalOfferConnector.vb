Imports System.Web.Services
Imports System.Xml.Schema
Imports System.Xml
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports Newtonsoft.Json
Imports Copient.CommonInc
Imports Copient.CPEOffer
Imports Copient.CMOffer
Imports CMS.AMS.Models
Imports CMS.AMS
Imports CMS.AMS.Contract
Imports Copient
Imports System.Threading

Public Delegate Function DetectOfferCollisionDelegate(OfferID As Long) As AMSResult(Of Integer)
<WebService(Namespace:="http://www.copienttech.com/ExternalOfferConnector/")>
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Public Class ExternalOfferConnector
    Inherits System.Web.Services.WebService
    Dim iOfferID_CDS As Long, DeploymentType_CDS As DEPLOYMENT_TYPES
    Dim _PGIDforOCD As Long = -1
    Dim _OfferNameforOCD As String = String.Empty
    ' Return error codes
    Private Enum ERROR_CODES As Integer
        ERROR_NONE = 0
        ERROR_IMPORT_FAILED = 1
        ERROR_MISSING_CLIENT_ID = 2
        ERROR_XML_INVALID_DOC = 3
        ERROR_XML_EMPTY_DOC = 4
        ERROR_XML_BAD_SOURCE_ID = 5
        ERROR_XSD_NOT_FOUND = 6
        ERROR_APPLICATION = 7
        ERROR_OFFER_ALREADY_EXISTS = 8
        ERROR_OFFER_DOES_NOT_EXIST = 9
        ERROR_OFFER_UPDATE_FAILED = 10
        ERROR_OFFER_REMOVE_FAILED = 11
        ERROR_BANNER_ALREADY_ASSSIGNED = 12
        ERROR_BANNER_ADD_FAILED = 13
        ERROR_BANNER_NOT_ASSIGNED = 14
        ERROR_BANNER_REMOVE_FAILED = 15
        ERROR_BANNERS_NOT_ENABLED = 16
        ERROR_OFFER_ID_CHANGE_FAILED = 17
        ERROR_INVALID_CUSTOMER_ID = 18
        ERROR_ADD_PRODUCT_FAILED = 19
        ERROR_REMOVE_PRODUCT_FAILED = 20
        ERROR_OFFER_FAILED_TO_DEPLOY = 21
        ERROR_INVALID_DISCOUNT_TYPE = 22
        ERROR_INCOMPLETE_OFFER_UPDATED = 24
        ERROR_MAX_OFFER_EXCEEDED = 25
        ERROR_INVALID_CHARGEBACK_VENDOR = 26
        ERROR_INCORRECT_ENGINE_TYPE = 27
        ERROR_BANNER_ENGINE_ID_NOT_SAME_AS_OFFER = 28
        ERROR_BANNER_ID_NOT_FOUND = 29
        ERROR_OFFER_ENGINE_ID_INVALID = 30
        ERROR_PRODUCT_DOES_NOT_EXIST = 31
        ERROR_INVALID_AMOUNT_TYPE = 32
        ERROR_INVALID_DISCOUNT_SCORECARD_ID = 33
        ERROR_DISABLED_COMPONENT = 34
        ERROR_INVALID_CARD_TYPE_ID = 35
        ERROR_ADD_CUSTOMERS_FAILED = 36
        ERROR_REMOVE_CUSTOMERS_FAILED = 37
        ERROR_INVALID_FORMAT = 38
        ERROR_INVALID_CHARGEBACK_DEPT = 39
        ERROR_INVALID_MULTIPLE_BANNERED_OFFER = 40
        ERROR_INVALID_BANNER_LOCATION = 41
        ERROR_INVALID_WEIGHT_VOLUMN_TYPE = 42
        ERROR_INVALID_CRMENGINE = 43
        ERROR_INVALID_ANY_CUSTOMER_OFFER = 44
        ERROR_CLIP_BUNDLE_FAILED = 45
        ERROR_UNCLIP_BUNDLE_FAILED = 46
        ERROR_OFFER_EXPIRED = 47
        ERROR_INVALID_PRODUCT_CONDITION_QUANTITY = 48
        ERROR_INVALID_CURRENCYID = 49
        ERROR_PERCENT_OFF_WITH_BEST_DEAL = 50 ' used in client specific implementation reserving the value here.
        ERROR_INVALID_PROGRAMUSES = 51
        ERROR_INVALID_OFFERTYPE = 52
        ERROR_INVALID_CUSTOMER_APPROVAL = 53

    End Enum

    Private Enum RESPONSE_TYPES As Integer
        ADD_OFFER = 1
        UPDATE_OFFER = 2
        REMOVE_OFFER = 3
        ADD_BANNER = 4
        REMOVE_BANNER = 5
        UPDATE_CLIENT_ID = 6
        ADD_CUSTOMER_OFFER = 7
        REMOVE_CUSTOMER_OFFER = 8
        ADD_PRODUCT_OFFER = 9
        REMOVE_PRODUCT_OFFER = 10
        GET_CUSTOMER_OFFERS = 11
        GET_OFFER_CUSTOMERS = 12
        GET_OFFER = 13
        ADD_CUSTOMERS_OFFER = 14
        REMOVE_CUSTOMERS_OFFER = 15
        ADD_CUSTOMERS = 16
        REMOVE_CUSTOMERS = 17
        CLIP_BUNDLE = 18
        UNCLIP_BUNDLE = 19
    End Enum

    Private Enum DEPLOYMENT_TYPES As Integer
        IMMEDIATE = 1
        DEFERRED = 2
    End Enum

    Private Enum EOC_LOG_TYPE As Integer
        ADD_CUSTOMER_TO_GROUP = 1
        REMOVE_CUSTOMER_FROM_GROUP = 2
        ADD_PRODUCT_TO_GROUP = 3
        REMOVE_PRODUCT_FROM_GROUP = 4
    End Enum

    Private Enum ENGINE_ID As Integer
        CM = 0
        CPE = 2
        UE = 9
    End Enum

    Private Structure CUSTOMERS_QUEUE_DATA
        Dim FileName As String
        Dim ExtInterfaceID As Integer
        Dim EngineID As Integer
        Dim ResponseType As RESPONSE_TYPES
        Dim FormatFileName As String
        Dim TreatAsClipData As Boolean
    End Structure

    Private Const CONNECTOR_ID As Integer = 22
    Private Const BANNER_OPT_ID As Integer = 66

    Private MyCommon As New Copient.CommonInc
    Private CreatedProductGroups As New ArrayList(3)
    Private CreatedCustomerGroups As New ArrayList(3)
    Private MyCmOffer As New Copient.CMOffer(MyCommon, False)
    Private MyCryptLib As New Copient.CryptLib
    Private messagingService As New MessagingService
    Private lUserId As Long = 1
    Private ExtInterfaceID As Long
    Private AcceptanceLogFile As String = "EOCAcceptanceLog." & Date.Now.ToString("yyyyMMdd") & ".txt"
    Private RejectionLogFile As String = "EOCRejectionLog." & Date.Now.ToString("yyyyMMdd") & ".txt"
    Private RejectionLogFileUnmasked As String = "EOCRejectionLog.Unmasked." & Date.Now.ToString("yyyyMMdd") & ".txt"

    Private EOCLogFile As String = "EOCErrorLogs." & Date.Now.ToString("yyyyMMdd") & ".txt"
    Private sErrorMethod As String = String.Empty
    Private methodName As String = String.Empty
    Private sErrorSubMethod As String = String.Empty


    Private Function GetDefaultCardTypeID() As String
        Return MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(30))
    End Function


    <WebMethod()>
    Public Function AddOffer(ByVal ExternalSourceID As String, ByVal OfferXml As String) As String
        Dim ErrorMsg As String = ""
        Dim ErrorCode As ERROR_CODES = ERROR_CODES.ERROR_NONE
        Dim EngineID As Integer
        Dim OfferXmlDoc As New XmlDocument

        Dim OfferFields As NameValueCollection
        Dim MyCpeOffer As New Copient.CPEOffer
        Dim OfferID As Long = -1
        Dim ClientOfferID As String = ""
        Dim OfferExists As Boolean = False
        Dim LogixID As Long = 0
        Dim Added As Boolean = False
        Dim AutoDeploy As Boolean = False
        Dim DefaultBannerID As Integer = 0
        Dim OfferType As String = ""
        Dim EngineType As String = ""
        Dim OverMaxOffers As Boolean = False
        Dim MaxOffers As Long = 0
        Dim DefaultAsLogixID As Boolean = False
        Dim ClientOfferIDSent As Boolean = False
        Dim IsClientIDCreated As Boolean = False
        Dim amsresult As AMSResult(Of Boolean) = New AMSResult(Of Boolean)
        sErrorMethod = "(AddOffer)"
        methodName = "AddOffer"
        CurrentRequest.Resolver.AppName = "External Offer Connector"
        Dim offerDeploymentValidator As IOfferDeploymentValidator = CurrentRequest.Resolver.Resolve(Of IOfferDeploymentValidator)()
        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

            ' Strip out Unicode characters in <Name> and <Description>
            OfferXml = CheckOfferName(OfferXml)
            OfferXml = CheckOfferDescription(OfferXml)

            Try
                OfferXmlDoc.LoadXml(OfferXml)
                TryParseAttributeValue(OfferXmlDoc, "Offer", "engine", EngineType)
            Catch ex As Exception
                ErrorMsg = ex.ToString
                ErrorCode = ERROR_CODES.ERROR_XML_INVALID_DOC
            End Try

            If ErrorMsg = "" Then
                EngineID = GetEngineID(EngineType)
                ExtInterfaceID = GetExtInterfaceID(ExternalSourceID, AutoDeploy, DefaultAsLogixID)

                OverMaxOffers = ExceedsMaximumOffers(ExtInterfaceID, MaxOffers)

                If ExtInterfaceID > 0 AndAlso EngineID > -1 AndAlso Not OverMaxOffers Then
                    Try
                        TryParseAttributeValue(OfferXmlDoc, "Offer", "type", OfferType)
                    Catch ex As Exception
                        ErrorMsg = ex.ToString
                        ErrorCode = ERROR_CODES.ERROR_XML_INVALID_DOC
                        OfferType = ""
                    End Try

                    If OfferType <> "" Then
                        ' validate that xmlDoc is valid and well-formed against appropriate schema
                        If IsValidDocument(OfferType, ErrorCode, ErrorMsg, EngineID, OfferXmlDoc) Then
                            ClientOfferIDSent = TryParseElementValue(OfferXmlDoc, "//Offer/ClientOfferID", ClientOfferID)
                            If DefaultAsLogixID OrElse (ClientOfferIDSent AndAlso ClientOfferID <> "") Then
                                ' check if the offer already exists before importing it.
                                If (ClientOfferIDSent AndAlso ClientOfferID <> "") Then
                                    Try
                                        OfferExists = DoesOfferExist(ClientOfferID, ExtInterfaceID, LogixID)
                                        If OfferExists Then
                                            ErrorCode = ERROR_CODES.ERROR_OFFER_ALREADY_EXISTS
                                            ErrorMsg = "Offer " & ClientOfferID & " already exists (LogixID: " & LogixID & ")"
                                        End If
                                    Catch ex As Exception
                                        ErrorCode = ERROR_CODES.ERROR_APPLICATION
                                        ErrorMsg = ex.ToString
                                    End Try
                                End If

                                If (Not OfferExists AndAlso ErrorMsg = "") Then
                                    If Not AreValidComponents(OfferXmlDoc, EngineID, ErrorMsg) Then
                                        ErrorCode = ERROR_CODES.ERROR_DISABLED_COMPONENT
                                    End If
                                End If

                                If (Not OfferExists AndAlso ErrorMsg = "") Then
                                    If Not IsValidDiscountType(ExternalSourceID, EngineID, OfferXmlDoc, ErrorMsg) Then
                                        ErrorCode = ERROR_CODES.ERROR_INVALID_DISCOUNT_TYPE
                                    End If
                                End If

                                If (Not OfferExists AndAlso ErrorMsg = "") Then
                                    If EngineID = 2 OrElse EngineID = 9 Then
                                        If Not IsValidChargebackVendor(ExternalSourceID, OfferXmlDoc, ErrorMsg) Then
                                            ErrorCode = ERROR_CODES.ERROR_INVALID_CHARGEBACK_VENDOR
                                        End If
                                    End If
                                End If

                                If (Not OfferExists AndAlso ErrorMsg = "") Then
                                    If Not IsValidAmountType(ExternalSourceID, OfferXmlDoc, ErrorMsg) Then
                                        ErrorCode = ERROR_CODES.ERROR_INVALID_AMOUNT_TYPE
                                    End If
                                End If

                                If (Not OfferExists AndAlso ErrorMsg = "") Then
                                    If EngineID = 2 OrElse EngineID = 9 Then
                                        If Not IsValidDiscountScorecard(ExternalSourceID, OfferXmlDoc, ErrorMsg) Then
                                            ErrorCode = ERROR_CODES.ERROR_INVALID_DISCOUNT_SCORECARD_ID
                                        End If
                                    End If
                                End If

                                If (Not OfferExists AndAlso ErrorMsg = "") Then
                                    If Not IsValidChargebackDept(OfferXmlDoc, EngineID, ErrorMsg) Then
                                        ErrorCode = ERROR_CODES.ERROR_INVALID_CHARGEBACK_DEPT
                                    End If
                                End If

                                If (Not OfferExists AndAlso ErrorMsg = "") Then
                                    If Not AreValidBanners(OfferXmlDoc, EngineID, False, ErrorMsg) Then
                                        ErrorCode = ERROR_CODES.ERROR_BANNER_ID_NOT_FOUND
                                    End If
                                End If

                                If (Not OfferExists AndAlso ErrorMsg = "") Then
                                    If Not IsValidBannerAssignment(OfferXmlDoc, EngineID, ErrorMsg) Then
                                        ErrorCode = ERROR_CODES.ERROR_INVALID_MULTIPLE_BANNERED_OFFER
                                    End If
                                End If

                                If (Not OfferExists AndAlso ErrorMsg = "") Then
                                    If Not AreValidBannerLocations(OfferXmlDoc, EngineID, ErrorMsg) Then
                                        ErrorCode = ERROR_CODES.ERROR_INVALID_BANNER_LOCATION
                                    End If
                                End If

                                If (Not OfferExists AndAlso ErrorMsg = "") Then
                                    If Not IsValidCRMEngine(OfferXmlDoc, EngineID, ErrorMsg) Then
                                        ErrorCode = ERROR_CODES.ERROR_INVALID_CRMENGINE
                                    End If
                                End If

                                If (Not OfferExists AndAlso ErrorMsg = "") Then
                                    If Not IsValidAnyCustomerOffer(OfferXmlDoc, EngineID, ErrorMsg) Then
                                        ErrorCode = ERROR_CODES.ERROR_INVALID_ANY_CUSTOMER_OFFER
                                    End If
                                End If

                                If (Not OfferExists AndAlso ErrorMsg = "") Then
                                    If Not IsValidProductCondition(OfferXmlDoc, EngineID, ErrorMsg) Then
                                        ErrorCode = ERROR_CODES.ERROR_INVALID_PRODUCT_CONDITION_QUANTITY
                                    End If
                                End If

                                If (Not OfferExists AndAlso ErrorMsg = "") Then
                                    If Not IsValidOfferType(OfferXmlDoc, EngineID, ErrorMsg) Then
                                        ErrorCode = ERROR_CODES.ERROR_INVALID_OFFERTYPE
                                    End If
                                End If
                                'validate customer approval condition
                                If (Not OfferExists AndAlso ErrorMsg = "") Then
                                    If Not IsValidCustomerApprovalCondition(OfferXmlDoc, EngineID, ErrorMsg) Then
                                        ErrorCode = ERROR_CODES.ERROR_INVALID_CUSTOMER_APPROVAL
                                    End If
                                End If

                                ' Added  For CurrencyID 
                                If (Not OfferExists AndAlso ErrorMsg = "") Then
                                    If Not IsvalidCurrencyId(OfferXmlDoc, EngineID, ErrorMsg) Then
                                        ErrorCode = ERROR_CODES.ERROR_INVALID_CURRENCYID
                                    End If
                                End If

                                If (ClientOfferIDSent AndAlso ClientOfferID <> "") Then
                                    If (Not OfferExists AndAlso ErrorMsg = "") Then
                                        If (Not InsertClientOfferId(ClientOfferID, ExtInterfaceID, IsClientIDCreated)) Then
                                            DoesOfferExist(ClientOfferID, ExtInterfaceID, LogixID)
                                            ErrorCode = ERROR_CODES.ERROR_OFFER_ALREADY_EXISTS
                                            ErrorMsg = "Offer " & ClientOfferID & " already exists (LogixID: " & LogixID & ")"
                                        End If
                                    End If
                                End If

                                If (Not OfferExists AndAlso ErrorMsg = "") Then
                                    OfferFields = GetCreateOfferFields(OfferXmlDoc, ExtInterfaceID, EngineID)
                                    Select Case EngineID
                                        Case 2, 9
                                            ' CPE or UE - create the offer header record
                                            OfferID = MyCpeOffer.CreateOffer(OfferFields, 1, ErrorMsg)
                                            If DefaultAsLogixID AndAlso OfferFields("ClientOfferID").Trim = "" Then
                                                ClientOfferID = OfferID
                                                SetClientOfferIDAsLogixID(OfferID, EngineID)
                                            End If

                                            If (OfferID <= 0) Then
                                                ErrorCode = ERROR_CODES.ERROR_IMPORT_FAILED
                                                ErrorMsg &= "Offer import failed."
                                            Else
                                                OfferFields = GetSaveOfferFields(OfferXmlDoc, OfferID, ExtInterfaceID, EngineID)
                                                MyCpeOffer.SaveOffer(OfferFields, 1, ErrorMsg, "AddOffer")
                                                _OfferNameforOCD = OfferFields.Get("form_name")
                                                ProcessCustomerConditions(OfferID, OfferXmlDoc, ErrorMsg, methodName)
                                                ProcessProductConditions(OfferID, OfferXmlDoc, ErrorMsg)
                                                ProcessTrackableCouponConditions(OfferXmlDoc, OfferID, EngineID, ErrorMsg)
                                                ProcessDiscount(OfferID, OfferXmlDoc, ExternalSourceID, EngineID, ErrorMsg)
                                                ProcessOfferLocations(OfferID, OfferXmlDoc, ErrorMsg, EngineID)
                                                ProcessOfferTerminals(OfferID, OfferXmlDoc, ErrorMsg, EngineID)

                                                HandleOfferTypes(ExternalSourceID, OfferID, OfferXmlDoc)
                                                HandleBestDealSetting(ExternalSourceID, OfferID)
                                                HandleVendorChargeback(OfferID)

                                                EnsureCustomerCondition(OfferID)
                                                'Check for Add Household to Group system option to automatically set Enable Householding flag (mg185181)
                                                If (MyCommon.Fetch_SystemOption(138) = "1") Then
                                                    MyCommon.QueryStr = "Update CPE_RewardOptions set HHEnable=1 where IncentiveID=" & OfferID
                                                    MyCommon.LRT_Execute()
                                                End If
                                                If ErrorMsg = "" Then
                                                    Added = True
                                                Else
                                                    ' an error occurred during offer building, so remove the entire offer.
                                                    CpeRemoveOffer(ClientOfferID, ExtInterfaceID, OfferID, ErrorCode, ErrorMsg)
                                                    MyCpeOffer.RemoveAllCustomerConditions(OfferID, 1)
                                                    MyCpeOffer.RemoveAllProductConditions(OfferID, 1)
                                                    MyCpeOffer.RemoveAllOfferLocations(OfferID, 1, True)
                                                    MyCpeOffer.RemoveAllOfferTerminals(OfferID, 1, True)
                                                    RemoveCreatedGroups()
                                                    OfferID = 0
                                                    ErrorCode = ERROR_CODES.ERROR_IMPORT_FAILED
                                                End If
                                            End If
                                        Case 0
                                            ' CM 
                                            OfferID = MyCmOffer.CreateOffer(OfferFields, lUserId, ErrorMsg)
                                            If DefaultAsLogixID AndAlso OfferFields("ClientOfferID").Trim = "" Then
                                                ClientOfferID = OfferID
                                                SetClientOfferIDAsLogixID(OfferID, EngineID)
                                            End If

                                            If (OfferID <= 0) Then
                                                ErrorCode = ERROR_CODES.ERROR_IMPORT_FAILED
                                                ErrorMsg &= "Offer import failed."
                                            Else
                                                OfferFields = GetSaveOfferFields(OfferXmlDoc, OfferID, ExtInterfaceID, EngineID)
                                                MyCmOffer.SaveOffer(OfferFields, lUserId, ErrorMsg)

                                                CmProcessCustomerConditions(OfferID, OfferXmlDoc, ErrorMsg)
                                                CmProcessProductConditions(OfferID, OfferXmlDoc, ErrorMsg, AutoDeploy)
                                                CmProcessDiscount(ExternalSourceID, OfferID, OfferXmlDoc, ErrorMsg)
                                                CmProcessOfferLocations(OfferID, OfferXmlDoc, ErrorMsg)
                                                CmProcessOfferTerminals(OfferID, OfferXmlDoc, ErrorMsg)

                                                HandleOfferTypes(ExternalSourceID, OfferID, OfferXmlDoc)
                                                'HandleBestDealSetting(ExternalSourceID, OfferID)
                                                'HandleVendorChargeback(OfferID)

                                                'EnsureCustomerCondition(OfferID)

                                                If ErrorMsg = "" Then
                                                    Added = True
                                                Else
                                                    ' an error occurred during offer building, so remove the entire offer.
                                                    CmRemoveOffer(ClientOfferID, ExtInterfaceID, OfferID, ErrorCode, ErrorMsg)
                                                    MyCmOffer.RemoveAllCustomerConditions(lUserId)
                                                    MyCmOffer.RemoveAllProductConditions(lUserId)
                                                    MyCmOffer.RemoveAllOfferLocations(lUserId, True)
                                                    MyCmOffer.RemoveAllOfferTerminals(lUserId, True)
                                                    RemoveCreatedGroups()
                                                    OfferID = 0
                                                    ErrorCode = ERROR_CODES.ERROR_IMPORT_FAILED
                                                End If
                                            End If
                                        Case Else
                                            ErrorCode = ERROR_CODES.ERROR_XML_BAD_SOURCE_ID
                                            ErrorMsg = "Unrecognized EngineID: " & EngineID & " ExternalSourceID: " & ExternalSourceID
                                    End Select

                                End If
                            Else
                                ErrorCode = ERROR_CODES.ERROR_MISSING_CLIENT_ID
                                ErrorMsg = "No client offer ID sent in Offer XML Document"
                            End If
                        Else
                            ErrorCode = ERROR_CODES.ERROR_XML_INVALID_DOC
                        End If
                    End If

                ElseIf OverMaxOffers Then
                    ErrorCode = ERROR_CODES.ERROR_MAX_OFFER_EXCEEDED
                    ErrorMsg = "Offer not added as this would exceed the maximum number of " & MaxOffers &
                               " offers from this external source (" & ExternalSourceID & ")"
                ElseIf EngineID <= -1 Then
                    ErrorCode = ERROR_CODES.ERROR_OFFER_ENGINE_ID_INVALID
                    ErrorMsg = "Unrecognized Engine: " & EngineType & " ExternalSourceID: " & ExternalSourceID
                Else
                    ErrorCode = ERROR_CODES.ERROR_XML_BAD_SOURCE_ID
                    ErrorMsg = "Unrecognized Engine: " & EngineType & " ExternalSourceID: " & ExternalSourceID
                End If

                If ErrorCode = ERROR_CODES.ERROR_NONE AndAlso OfferID > 0 Then
                    If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
                    MyCommon.Activity_Log(3, OfferID, 1, Copient.PhraseLib.Lookup("history.offer-create", 1))
                End If

                If AutoDeploy And ErrorCode = ERROR_CODES.ERROR_NONE And ErrorMsg = "" Then
                    If (EngineID = 9 Or EngineID = 2) Then
                        amsresult = offerDeploymentValidator.ValidateCPEOffer(OfferID, False, False)
                    ElseIf (EngineID = 0) Then
                        amsresult = offerDeploymentValidator.ValidateCMOffer(OfferID, False, False)
                    End If
                    If (amsresult.ResultType = AMSResultType.Success) Then
                        If (amsresult.Result = True) Then
                            If IsValidLocationCurrencyMatch(OfferID, EngineID, ErrorMsg) Then
                                If Not DeployOffer(OfferID, EngineID, OfferXmlDoc, RESPONSE_TYPES.ADD_OFFER) Then
                                    ErrorCode = ERROR_CODES.ERROR_OFFER_FAILED_TO_DEPLOY
                                    ErrorMsg = "Offer " & ClientOfferID & " was added but failed to automatically deploy."
                                End If
                            Else
                                ErrorCode = ERROR_CODES.ERROR_OFFER_FAILED_TO_DEPLOY
                                ErrorMsg = "Offer " & ClientOfferID & " was added but failed to automatically deploy." + ErrorMsg
                            End If
                        Else
                            ErrorCode = ERROR_CODES.ERROR_OFFER_FAILED_TO_DEPLOY
                            ErrorMsg = Copient.PhraseLib.Lookup(amsresult.MessageString, 1)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Added = False
            ErrorCode = ERROR_CODES.ERROR_APPLICATION
            ErrorMsg = "Application error encountered: " & ex.ToString
        Finally

            If (IsClientIDCreated) AndAlso Not String.IsNullOrEmpty(ErrorMsg) Then
                DeleteClientID(ClientOfferID, ExtInterfaceID)
            End If
            AppendToLog(ExternalSourceID, "AddOffer", OfferXml, "ClientOfferID=" & ClientOfferID, ErrorCode, ErrorMsg)
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try

        Return WriteOfferXmlResponse(OfferID, ClientOfferID, ErrorCode, ErrorMsg, RESPONSE_TYPES.ADD_OFFER, Added)
    End Function

    <WebMethod()>
    Public Function RemoveOffer(ByVal ExternalSourceID As String, ByVal ClientOfferID As String) As String
        Dim Removed As Boolean = False
        Dim ErrorMsg As String = ""
        Dim ErrorCode As ERROR_CODES = ERROR_CODES.ERROR_NONE
        Dim LogixID As Long
        Dim AutoDeploy As Boolean = False
        Dim DefaultAsLogixID As Boolean = False
        Dim OfferExists As Boolean = False
        Dim OfferEngineID As Long = -1

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

            ExtInterfaceID = GetExtInterfaceID(ExternalSourceID, AutoDeploy)

            If ExtInterfaceID > 0 Then
                OfferExists = DoesOfferExist(ClientOfferID, ExtInterfaceID, LogixID, OfferEngineID)
                If OfferExists Then
                    If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
                    Select Case OfferEngineID
                        Case 0
                            Removed = CmRemoveOffer(ClientOfferID, ExtInterfaceID, LogixID, ErrorCode, ErrorMsg)
                        Case 2, 9
                            Removed = CpeRemoveOffer(ClientOfferID, ExtInterfaceID, LogixID, ErrorCode, ErrorMsg)
                        Case Else
                            ErrorCode = ERROR_CODES.ERROR_OFFER_ENGINE_ID_INVALID
                            ErrorMsg = "Invalid Offer EngineID: " & OfferEngineID
                    End Select
                Else
                    ErrorCode = ERROR_CODES.ERROR_OFFER_DOES_NOT_EXIST
                    ErrorMsg = "Offer " & ClientOfferID & " does not exist."
                End If
            Else
                ErrorCode = ERROR_CODES.ERROR_XML_BAD_SOURCE_ID
                ErrorMsg = "Unrecognized ExternalSourceID: " & ExternalSourceID
            End If

            If (ErrorCode = ERROR_CODES.ERROR_NONE AndAlso Removed) Then
                MyCommon.Activity_Log(3, LogixID, 1, Copient.PhraseLib.Lookup("history.offer-delete", 1))
            End If
        Catch ex As Exception
            Removed = False
            ErrorCode = ERROR_CODES.ERROR_APPLICATION
            ErrorMsg = "Application error encountered: " & ex.ToString
        Finally
            AppendToLog(ExternalSourceID, "RemoveOffer", "", "ClientOffer=" & ClientOfferID, ErrorCode, ErrorMsg)
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try

        Return WriteOfferXmlResponse(LogixID, ClientOfferID, ErrorCode, ErrorMsg, RESPONSE_TYPES.REMOVE_OFFER, Removed)
    End Function

    <WebMethod()>
    Public Function UpdateOffer(ByVal ExternalSourceID As String, ByVal OfferXml As String) As String
        Dim ErrorMsg As String = ""
        Dim ErrorCode As ERROR_CODES = ERROR_CODES.ERROR_NONE
        Dim EngineID As Integer
        Dim OfferXmlDoc As New XmlDocument

        Dim OfferFields As NameValueCollection
        Dim MyCpeOffer As New Copient.CPEOffer
        Dim OfferID As Long = -1
        Dim ClientOfferID As String = ""
        Dim OfferExists As Boolean = False
        Dim Updated As Boolean = False
        Dim AutoDeploy As Boolean = False
        Dim DefaultBannerID As Integer = 0
        Dim OfferType As String = ""
        Dim EngineType As String = ""
        Dim OfferEngineID As Long = -1
        methodName = "UpdateOffer"
        CurrentRequest.Resolver.AppName = "External Offer Connector"

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

            ' Strip out Unicode characters in <Name> and <Description>
            OfferXml = CheckOfferName(OfferXml)
            OfferXml = CheckOfferDescription(OfferXml)

            Try
                OfferXmlDoc.LoadXml(OfferXml)
                TryParseAttributeValue(OfferXmlDoc, "Offer", "engine", EngineType)
            Catch ex As Exception
                ErrorMsg = ex.ToString
                ErrorCode = ERROR_CODES.ERROR_XML_INVALID_DOC
            End Try

            If ErrorMsg = "" Then
                EngineID = GetEngineID(EngineType)
                ExtInterfaceID = GetExtInterfaceID(ExternalSourceID, AutoDeploy)

                If ExtInterfaceID > 0 AndAlso (EngineID = 2 Or EngineID = 0 Or EngineID = 9) Then
                    Try
                        TryParseAttributeValue(OfferXmlDoc, "Offer", "type", OfferType)
                    Catch ex As Exception
                        ErrorCode = ERROR_CODES.ERROR_XML_INVALID_DOC
                    End Try

                    If OfferType <> "" Then
                        ' validate that xmlDoc is valid and well-formed against appropriate schema
                        If IsValidDocument(OfferType, ErrorCode, ErrorMsg, EngineID, OfferXmlDoc) Then

                            If (TryParseElementValue(OfferXmlDoc, "//Offer/ClientOfferID", ClientOfferID) AndAlso ClientOfferID <> "") Then
                                ' check if the offer already exists before importing it.
                                Try
                                    OfferExists = DoesOfferExist(ClientOfferID, ExtInterfaceID, OfferID, OfferEngineID)
                                    If Not OfferExists Then
                                        ErrorCode = ERROR_CODES.ERROR_OFFER_DOES_NOT_EXIST
                                        ErrorMsg = "Offer " & ClientOfferID & " does not exist in Logix"
                                    ElseIf EngineID <> OfferEngineID Then
                                        ErrorCode = ERROR_CODES.ERROR_INCORRECT_ENGINE_TYPE
                                        ErrorMsg = "Offer " & ClientOfferID & " is not for Engine type: " & EngineType
                                    End If
                                Catch ex As Exception
                                    ErrorCode = ERROR_CODES.ERROR_APPLICATION
                                    ErrorMsg = ex.ToString
                                End Try

                                If (OfferExists AndAlso ErrorMsg = "") Then
                                    If Not AreValidComponents(OfferXmlDoc, EngineID, ErrorMsg) Then
                                        ErrorCode = ERROR_CODES.ERROR_DISABLED_COMPONENT
                                    End If
                                End If

                                If (OfferExists AndAlso ErrorMsg = "") Then
                                    If Not IsValidDiscountType(ExternalSourceID, EngineID, OfferXmlDoc, ErrorMsg) Then
                                        ErrorCode = ERROR_CODES.ERROR_INVALID_DISCOUNT_TYPE
                                    End If
                                End If

                                If (OfferExists AndAlso ErrorMsg = "") Then
                                    If EngineID = 2 OrElse EngineID = 9 Then
                                        If Not IsValidChargebackVendor(ExternalSourceID, OfferXmlDoc, ErrorMsg) Then
                                            ErrorCode = ERROR_CODES.ERROR_INVALID_CHARGEBACK_VENDOR
                                        End If
                                    End If
                                End If

                                If (OfferExists AndAlso ErrorMsg = "") Then
                                    If Not IsValidAmountType(ExternalSourceID, OfferXmlDoc, ErrorMsg) Then
                                        ErrorCode = ERROR_CODES.ERROR_INVALID_AMOUNT_TYPE
                                    End If
                                End If

                                If (OfferExists AndAlso ErrorMsg = "") Then
                                    If EngineID = 2 OrElse EngineID = 9 Then
                                        If Not IsValidDiscountScorecard(ExternalSourceID, OfferXmlDoc, ErrorMsg) Then
                                            ErrorCode = ERROR_CODES.ERROR_INVALID_DISCOUNT_SCORECARD_ID
                                        End If
                                    End If
                                End If

                                If (OfferExists AndAlso ErrorMsg = "") Then
                                    If Not IsValidChargebackDept(OfferXmlDoc, EngineID, ErrorMsg) Then
                                        ErrorCode = ERROR_CODES.ERROR_INVALID_CHARGEBACK_DEPT
                                    End If
                                End If

                                If (OfferExists AndAlso ErrorMsg = "") Then
                                    If Not AreValidBanners(OfferXmlDoc, EngineID, False, ErrorMsg) Then
                                        ErrorCode = ERROR_CODES.ERROR_BANNER_ID_NOT_FOUND
                                    End If
                                End If

                                If (OfferExists AndAlso ErrorMsg = "") Then
                                    If Not IsValidBannerAssignment(OfferXmlDoc, EngineID, ErrorMsg) Then
                                        ErrorCode = ERROR_CODES.ERROR_INVALID_MULTIPLE_BANNERED_OFFER
                                    End If
                                End If

                                If (OfferExists AndAlso ErrorMsg = "") Then
                                    If Not AreValidBannerLocations(OfferXmlDoc, EngineID, ErrorMsg) Then
                                        ErrorCode = ERROR_CODES.ERROR_INVALID_BANNER_LOCATION
                                    End If
                                End If

                                If (OfferExists AndAlso ErrorMsg = "") Then
                                    If Not IsValidCRMEngine(OfferXmlDoc, EngineID, ErrorMsg) Then
                                        ErrorCode = ERROR_CODES.ERROR_INVALID_CRMENGINE
                                    End If
                                End If

                                If (OfferExists AndAlso ErrorMsg = "") Then
                                    If Not IsValidAnyCustomerOffer(OfferXmlDoc, EngineID, ErrorMsg) Then
                                        ErrorCode = ERROR_CODES.ERROR_INVALID_ANY_CUSTOMER_OFFER
                                    End If
                                End If
                                If (OfferExists AndAlso ErrorMsg = "") Then
                                    If IsOfferExpiredAndLocked(OfferID, EngineID, ErrorMsg) Then
                                        ErrorCode = ERROR_CODES.ERROR_OFFER_EXPIRED
                                    End If
                                End If

                                If (Not OfferExists AndAlso ErrorMsg = "") Then
                                    If Not IsValidProductCondition(OfferXmlDoc, EngineID, ErrorMsg) Then
                                        ErrorCode = ERROR_CODES.ERROR_INVALID_PRODUCT_CONDITION_QUANTITY
                                    End If
                                End If

                                If (OfferExists AndAlso ErrorMsg = "") Then
                                    If Not IsValidCustomerApprovalCondition(OfferXmlDoc, EngineID, ErrorMsg) Then
                                        ErrorCode = ERROR_CODES.ERROR_INVALID_CUSTOMER_APPROVAL
                                    End If
                                End If

                                If (OfferExists AndAlso ErrorMsg = "") Then
                                    OfferFields = GetSaveOfferFields(OfferXmlDoc, OfferID, ExtInterfaceID, EngineID)
                                    Select Case EngineID
                                        Case 2, 9
                                            MyCpeOffer.SaveOffer(OfferFields, 1, ErrorMsg, "UpdateOffer")

                                            ProcessCustomerConditions(OfferID, OfferXmlDoc, ErrorMsg, methodName)
                                            ProcessProductConditions(OfferID, OfferXmlDoc, ErrorMsg)
                                            ProcessTrackableCouponConditions(OfferXmlDoc, OfferID, EngineID, ErrorMsg)
                                            ProcessDiscount(OfferID, OfferXmlDoc, ExternalSourceID, EngineID, ErrorMsg)
                                            ProcessOfferLocations(OfferID, OfferXmlDoc, ErrorMsg, EngineID)
                                            ProcessOfferTerminals(OfferID, OfferXmlDoc, ErrorMsg, EngineID)

                                            HandleOfferTypes(ExternalSourceID, OfferID, OfferXmlDoc)
                                            HandleBestDealSetting(ExternalSourceID, OfferID)
                                            HandleVendorChargeback(OfferID)

                                            EnsureCustomerCondition(OfferID)

                                            If ErrorMsg = "" Then
                                                Updated = True
                                            Else
                                                ErrorCode = ERROR_CODES.ERROR_INCOMPLETE_OFFER_UPDATED
                                            End If
                                        Case 0
                                            MyCmOffer.SaveOffer(OfferFields, lUserId, ErrorMsg)

                                            CmProcessCustomerConditions(OfferID, OfferXmlDoc, ErrorMsg)
                                            CmProcessProductConditions(OfferID, OfferXmlDoc, ErrorMsg)
                                            CmProcessDiscount(ExternalSourceID, OfferID, OfferXmlDoc, ErrorMsg)
                                            CmProcessOfferLocations(OfferID, OfferXmlDoc, ErrorMsg)
                                            CmProcessOfferTerminals(OfferID, OfferXmlDoc, ErrorMsg)

                                            HandleOfferTypes(ExternalSourceID, OfferID, OfferXmlDoc)
                                            'HandleBestDealSetting(ExternalSourceID, OfferID)
                                            'HandleVendorChargeback(OfferID)
                                            'EnsureCustomerCondition(OfferID)

                                            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

                                            If ErrorMsg = "" Then
                                                Updated = True
                                            Else
                                                ErrorCode = ERROR_CODES.ERROR_INCOMPLETE_OFFER_UPDATED
                                            End If
                                    End Select
                                End If

                            Else
                                ErrorCode = ERROR_CODES.ERROR_MISSING_CLIENT_ID
                                ErrorMsg = "No client offer ID sent in Offer XML Document"
                            End If

                        Else
                            ErrorCode = ERROR_CODES.ERROR_XML_INVALID_DOC
                        End If
                    End If

                Else
                    ErrorCode = ERROR_CODES.ERROR_XML_BAD_SOURCE_ID
                    ErrorMsg = "Unrecognized EngineID: " & EngineID & " ExternalSourceID: " & ExternalSourceID
                End If

                If (ErrorCode = ERROR_CODES.ERROR_NONE OrElse ErrorCode = ERROR_CODES.ERROR_INCOMPLETE_OFFER_UPDATED) AndAlso OfferID > 0 Then
                    MyCommon.Activity_Log(3, OfferID, 1, Copient.PhraseLib.Lookup("history.offer-edit", 1))
                End If

                If AutoDeploy And ErrorCode = ERROR_CODES.ERROR_NONE Then
                    If Not DeployOffer(OfferID, EngineID, OfferXmlDoc, RESPONSE_TYPES.UPDATE_OFFER) Then
                        ErrorCode = ERROR_CODES.ERROR_OFFER_FAILED_TO_DEPLOY
                        ErrorMsg = "Offer " & ClientOfferID & " was added but failed to automatically deploy."
                    End If
                End If
            End If
        Catch ex As Exception
            Updated = False
            ErrorCode = ERROR_CODES.ERROR_APPLICATION
            ErrorMsg = "Application error encountered: " & ex.ToString
        Finally
            AppendToLog(ExternalSourceID, "UpdateOffer", OfferXml, "", ErrorCode, ErrorMsg)
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try

        Return WriteOfferXmlResponse(OfferID, ClientOfferID, ErrorCode, ErrorMsg, RESPONSE_TYPES.UPDATE_OFFER, Updated)
    End Function

    <WebMethod()>
    Public Function GetOffer(ByVal ExternalSourceID As String, ByVal ClientOfferID As String) As String
        Dim ResponseXml As String = ""
        Dim OfferXml As String = ""
        Dim ErrorXml As New StringBuilder()
        Dim MyCpeOffer As New Copient.CPEOffer
        Dim LogixID As Long
        Dim AutoDeploy As Boolean
        Dim ErrorCode As ERROR_CODES = ERROR_CODES.ERROR_NONE
        Dim ErrorMsg As String = ""
        Dim OfferEngineId As Long = -1

        Try
            ExtInterfaceID = GetExtInterfaceID(ExternalSourceID, AutoDeploy)
            If (ExtInterfaceID > 0) Then
                If DoesOfferExist(ClientOfferID, ExtInterfaceID, LogixID, OfferEngineId) Then
                    Select Case OfferEngineId
                        Case 0
                            ErrorCode = ERROR_CODES.ERROR_OFFER_ENGINE_ID_INVALID
                            ErrorMsg = "This method does not currently handle CM offers - Offer EngineID: " & OfferEngineId
                        Case 2, 9
                            MyCpeOffer.SetEngineID(OfferEngineId)
                            OfferXml = MyCpeOffer.GetOfferXML(LogixID, ErrorMsg)
                        Case Else
                            ErrorCode = ERROR_CODES.ERROR_OFFER_ENGINE_ID_INVALID
                            ErrorMsg = "Invalid Offer EngineID: " & OfferEngineId
                    End Select
                Else
                    ErrorCode = ERROR_CODES.ERROR_OFFER_DOES_NOT_EXIST
                    ErrorMsg = "Offer " & ClientOfferID & " does not exist in Logix."
                End If
            Else
                ErrorCode = ERROR_CODES.ERROR_XML_BAD_SOURCE_ID
                ErrorMsg = "Unrecognized ExternalSourceID: " & ExternalSourceID
            End If

        Catch ex As Exception
            ErrorCode = ERROR_CODES.ERROR_APPLICATION
            ErrorMsg = "Application error encountering: " & ex.ToString
        End Try

        If (ErrorCode = ERROR_CODES.ERROR_NONE) Then
            ResponseXml = OfferXml.Replace("encoding=""utf-16""", "encoding=""utf-8""")
            OfferXml = Nothing
        Else
            ' create error xml
            ErrorXml.Append("<?xml version=""1.0"" encoding=""UTF-8""?>" & ControlChars.CrLf)
            ErrorXml.Append("<ExternalOfferConnector>" & ControlChars.CrLf)
            ErrorXml.Append("  <Offer id=""" & ClientOfferID & """ logixId=""" & LogixID & """ operation=""" & RESPONSE_TYPES.GET_OFFER.ToString & """ success=""false"" />" & ControlChars.CrLf)
            ErrorXml.Append("  <Error code=""" & ErrorCode.ToString & """ message=""" & ErrorMsg & """ />" & ControlChars.CrLf)
            ErrorXml.Append("</ExternalOfferConnector>")

            ResponseXml = ErrorXml.ToString
        End If

        AppendToLog(ExternalSourceID, "GetOffer", "", "ClientOfferID=" & ClientOfferID, ErrorCode, ErrorMsg)

        Return ResponseXml
    End Function

    <WebMethod()>
    Public Function AddBannerToOffer(ByVal ExternalSourceID As String, ByVal ClientOfferID As String, ByVal BannerID As Integer) As String
        Dim Added As Boolean = False
        Dim LogixID As Long = 0
        Dim dt As DataTable
        Dim ErrorCode As ERROR_CODES = ERROR_CODES.ERROR_NONE
        Dim ErrorMsg As String = ""
        Dim AutoDeploy As Boolean = False
        Dim OfferEngineID As Long = -1
        Dim BannerEngineID As Long

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

            ExtInterfaceID = GetExtInterfaceID(ExternalSourceID, AutoDeploy)

            If (ExtInterfaceID > 0) Then
                If (MyCommon.Fetch_SystemOption(66) = "1") Then
                    If DoesOfferExist(ClientOfferID, ExtInterfaceID, LogixID, OfferEngineID) Then
                        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
                        MyCommon.QueryStr = "select BannerID from BannerOffers with (NoLock) " &
                                            "where BannerID = @BannerID and OfferID = @OfferID"
                        MyCommon.DBParameters.Add("@BannerID", SqlDbType.Int).Value = BannerID
                        MyCommon.DBParameters.Add("@OfferID", SqlDbType.Int).Value = LogixID
                        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                        If (dt.Rows.Count > 0) Then
                            ErrorCode = ERROR_CODES.ERROR_BANNER_ALREADY_ASSSIGNED
                            ErrorMsg = "Banner " & BannerID & " is already assigned for offer " & ClientOfferID
                        Else
                            MyCommon.QueryStr = "select B.BannerID, BE.EngineID from Banners B with (NoLock) " &
                                                "inner join BannerEngines BE with (NoLock) on BE.BannerID=B.BannerID " &
                                                "where B.BannerID = @BannerID"
                            MyCommon.DBParameters.Add("@BannerID", SqlDbType.Int).Value = BannerID
                            dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                            If (dt.Rows.Count > 0) Then
                                BannerEngineID = MyCommon.NZ(dt.Rows(0).Item("EngineID"), 0)
                                If OfferEngineID = BannerEngineID Then
                                    MyCommon.QueryStr = "insert into BannerOffers with (RowLock) (BannerID, OfferID) values (@BannerID, @OfferID)"
                                    MyCommon.DBParameters.Add("@BannerID", SqlDbType.Int).Value = BannerID
                                    MyCommon.DBParameters.Add("@OfferID", SqlDbType.Int).Value = LogixID
                                    MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
                                    Added = (MyCommon.RowsAffected > 0)
                                Else
                                    ErrorCode = ERROR_CODES.ERROR_BANNER_ENGINE_ID_NOT_SAME_AS_OFFER
                                    ErrorMsg = "EngineID for Banner " & BannerID & " does not match EngineID of offer " & ClientOfferID
                                End If
                            Else
                                ErrorCode = ERROR_CODES.ERROR_BANNER_ID_NOT_FOUND
                                ErrorMsg = "Banner " & BannerID & " is not defined in Logix"
                            End If

                        End If
                    Else
                        ErrorCode = ERROR_CODES.ERROR_OFFER_DOES_NOT_EXIST
                        ErrorMsg = "Offer " & ClientOfferID & " does not exist"
                        Added = False
                    End If
                Else
                    ErrorCode = ERROR_CODES.ERROR_BANNERS_NOT_ENABLED
                    ErrorMsg = "Banners are not enabled."
                    Added = False
                End If
            Else
                ErrorCode = ERROR_CODES.ERROR_XML_BAD_SOURCE_ID
                ErrorMsg = "Unrecognized ExternalSourceID: " & ExternalSourceID
            End If

            If (ErrorCode = ERROR_CODES.ERROR_NONE AndAlso Added) Then
                MyCommon.Activity_Log(3, LogixID, 1, Copient.PhraseLib.Lookup("history.offer-edit", 1))
            End If

        Catch ex As Exception
            ErrorCode = ERROR_CODES.ERROR_BANNER_ADD_FAILED
            ErrorMsg = "Adding Banner " & BannerID & " to offer " & ClientOfferID & "failed for following reason: " & ex.ToString
            Added = False
        Finally
            AppendToLog(ExternalSourceID, "AddBannerToOffer", "", "ClientOfferID=" & ClientOfferID & "; BannerID=" & BannerID, ErrorCode, ErrorMsg)
            MyCommon.Close_LogixRT()
        End Try

        Return WriteOfferXmlResponse(LogixID, ClientOfferID, ErrorCode, ErrorMsg, RESPONSE_TYPES.ADD_BANNER, Added)
    End Function

    <WebMethod()>
    Public Function RemoveBannerFromOffer(ByVal ExternalSourceID As String, ByVal ClientOfferID As String, ByVal BannerID As Integer) As String
        Dim Removed As Boolean = False
        Dim LogixID As Long = 0
        Dim dt As DataTable
        Dim ErrorCode As ERROR_CODES = ERROR_CODES.ERROR_NONE
        Dim ErrorMsg As String = ""
        Dim AutoDeploy As Boolean = False

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

            ExtInterfaceID = GetExtInterfaceID(ExternalSourceID, AutoDeploy)

            If (ExtInterfaceID > 0) Then
                If (MyCommon.Fetch_SystemOption(66) = "1") Then
                    If DoesOfferExist(ClientOfferID, ExtInterfaceID, LogixID) Then
                        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
                        MyCommon.QueryStr = "select BannerID from BannerOffers where BannerID = @BannerID and OfferID = @OfferID "
                        MyCommon.DBParameters.Add("@BannerID", SqlDbType.Int).Value = BannerID
                        MyCommon.DBParameters.Add("@OfferID", SqlDbType.Int).Value = LogixID
                        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                        If (dt.Rows.Count > 0) Then
                            MyCommon.QueryStr = "delete from BannerOffers with (RowLock) where BannerID = @BannerID and OfferID = @OfferID"
                            MyCommon.DBParameters.Add("@BannerID", SqlDbType.Int).Value = BannerID
                            MyCommon.DBParameters.Add("@OfferID", SqlDbType.Int).Value = LogixID
                            MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
                            Removed = (MyCommon.RowsAffected > 0)
                        Else
                            ErrorCode = ERROR_CODES.ERROR_BANNER_NOT_ASSIGNED
                            ErrorMsg = "Banner " & BannerID & " is not assigned to offer " & ClientOfferID
                        End If
                    Else
                        ErrorCode = ERROR_CODES.ERROR_OFFER_DOES_NOT_EXIST
                        ErrorMsg = "Offer " & ClientOfferID & " does not exist"
                    End If
                Else
                    ErrorCode = ERROR_CODES.ERROR_BANNERS_NOT_ENABLED
                    ErrorMsg = "Banners are not enabled."
                    Removed = False
                End If
            Else
                ErrorCode = ERROR_CODES.ERROR_XML_BAD_SOURCE_ID
                ErrorMsg = "Unrecognized ExternalSourceID: " & ExternalSourceID
            End If

            If (ErrorCode = ERROR_CODES.ERROR_NONE AndAlso Removed) Then
                MyCommon.Activity_Log(3, LogixID, 1, Copient.PhraseLib.Lookup("history.offer-edit", 1))
            End If

        Catch ex As Exception
            ErrorCode = ERROR_CODES.ERROR_BANNER_REMOVE_FAILED
            ErrorMsg = "Removing Banner " & BannerID & " from offer " & ClientOfferID & "failed for following reason: " & ex.ToString
            Removed = False
        Finally
            AppendToLog(ExternalSourceID, "RemoveBannerFromOffer", "", "ClientOfferID=" & ClientOfferID & "; BannerID=" & BannerID, ErrorCode, ErrorMsg)
            MyCommon.Close_LogixRT()
        End Try

        Return WriteOfferXmlResponse(LogixID, ClientOfferID, ErrorCode, ErrorMsg, RESPONSE_TYPES.REMOVE_BANNER, Removed)
    End Function


    <WebMethod()>
    Public Function UpdateClientOfferID(ByVal ExternalSourceID As String, ByVal ClientOfferID As String, ByVal NewClientOfferID As String) As String
        Dim LogixID, NewLogixID As Long
        Dim ErrorCode As ERROR_CODES = ERROR_CODES.ERROR_NONE
        Dim ErrorMsg As String = ""
        Dim Updated As Boolean = False
        Dim OfferExists As Boolean = False
        Dim NewOfferIdExists As Boolean = False
        Dim AutoDeploy As Boolean = False
        Dim OfferEngineID As Long = -1

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

            ExtInterfaceID = GetExtInterfaceID(ExternalSourceID, AutoDeploy)

            If (ExtInterfaceID > 0) Then
                OfferExists = DoesOfferExist(ClientOfferID, ExtInterfaceID, LogixID, OfferEngineID)
                If OfferExists Then
                    If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

                    NewOfferIdExists = DoesOfferExist(NewClientOfferID, ExtInterfaceID, NewLogixID)
                    If Not NewOfferIdExists Then
                        Select Case OfferEngineID
                            Case 0
                                MyCommon.QueryStr = "update Offers set ExtOfferID = @ExtOfferID " &
                                                    "where OfferID = @OfferID and Deleted=0"
                                MyCommon.DBParameters.Add("@ExtOfferID", SqlDbType.NVarChar).Value = NewClientOfferID.ConvertBlankIfNothing
                                MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = LogixID
                            Case 2, 9
                                MyCommon.QueryStr = "update CPE_Incentives set ClientOfferID = @NewClientOfferID " &
                                                    "where ClientOfferID = @ClientOfferID and InboundCRMEngineID = @InboundCRMEngineID and Deleted=0"
                                MyCommon.DBParameters.Add("@NewClientOfferID", SqlDbType.NVarChar).Value = NewClientOfferID.ConvertBlankIfNothing
                                MyCommon.DBParameters.Add("@ClientOfferID", SqlDbType.NVarChar).Value = ClientOfferID.ConvertBlankIfNothing
                                MyCommon.DBParameters.Add("@InboundCRMEngineID", SqlDbType.Int).Value = ExtInterfaceID
                            Case Else
                                ErrorCode = ERROR_CODES.ERROR_OFFER_ENGINE_ID_INVALID
                                ErrorMsg = "Invalid Offer EngineID: " & OfferEngineID
                        End Select

                        If ErrorMsg = "" Then
                            MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
                            Updated = (MyCommon.RowsAffected > 0)
                            If Not Updated Then
                                ErrorCode = ERROR_CODES.ERROR_OFFER_ID_CHANGE_FAILED
                                ErrorMsg = "Unable to change Client OfferID during database update statement: " & MyCommon.QueryStr
                            End If
                        End If
                    Else
                        ErrorCode = ERROR_CODES.ERROR_OFFER_ALREADY_EXISTS
                        ErrorMsg = "An offer with the ID " & NewClientOfferID & " already exists in Logix."
                    End If
                Else
                    ErrorCode = ERROR_CODES.ERROR_OFFER_DOES_NOT_EXIST
                    ErrorMsg = "Offer " & ClientOfferID & " does not exist."
                End If
            Else
                ErrorCode = ERROR_CODES.ERROR_XML_BAD_SOURCE_ID
                ErrorMsg = "Unrecognized ExternalSourceID: " & ExternalSourceID
            End If

            If (ErrorCode = ERROR_CODES.ERROR_NONE AndAlso Updated) Then
                MyCommon.Activity_Log(3, LogixID, 1, Copient.PhraseLib.Lookup("history.offer-edit", 1))
            End If

        Catch ex As Exception
            ErrorCode = ERROR_CODES.ERROR_APPLICATION
            ErrorMsg = "Application Error encountered: " & ex.ToString
        Finally
            AppendToLog(ExternalSourceID, "UpdateClientOfferID", "", "ClientOfferID=" & ClientOfferID & "; NewClientOfferID=" & NewClientOfferID, ErrorCode, ErrorMsg)
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try

        Return WriteOfferXmlResponse(LogixID, NewClientOfferID, ErrorCode, ErrorMsg, RESPONSE_TYPES.UPDATE_CLIENT_ID, Updated)
    End Function

    <WebMethod()>
    Private Function HandleAddCustomerCardToOffer(ByVal ExternalSourceID As String, ByVal CardID As String,
                                         ByVal CardTypeID As String, ByVal ClientOfferID As String, ByVal CalledBy As String) As String
        Dim Added As Boolean = False
        Dim LogixID As Long
        Dim ErrorCode As ERROR_CODES = ERROR_CODES.ERROR_NONE
        Dim ErrorMsg As String = ""
        Dim CustomerPK As Long = 0
        Dim AutoDeploy As Boolean = False
        Dim OfferEngineID As Long = -1
        Dim CardTypeIDint As Integer
        Dim Response As CardValidationResponse
        Dim Offers As IOffer
        Dim MyPoints As Copient.Points
        Dim StoredValue As Copient.StoredValue
        Dim PointsAdj, PointsBalance, PointsRequired As Decimal
        Dim PointsBalanceOriginal, PointsPending As Decimal
        Dim IncludePendingPoints As Boolean = False
        Dim ErrorMsgLog As String = ""
        Dim ErrorMsgLogUnmasked As String = ""
        Try

            Offers = CurrentRequest.Resolver.Resolve(Of IOffer)()
            MyPoints = CurrentRequest.Resolver.Resolve(Of Copient.Points)()
            StoredValue = CurrentRequest.Resolver.Resolve(Of Copient.StoredValue)()

            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            ExtInterfaceID = GetExtInterfaceID(ExternalSourceID, AutoDeploy)

            If (ExtInterfaceID > 0) Then
                If DoesOfferExist(ClientOfferID, ExtInterfaceID, LogixID, OfferEngineID) Then
                    If OfferEngineID = 0 Or OfferEngineID = 2 Or OfferEngineID = 9 Then
                        If MyCommon.AllowToProcessCustomerCard(CardID, CardTypeID, Response) Then
                            CardTypeIDint = Convert.ToInt32(CardTypeID)
                            CardID = MyCommon.Pad_ExtCardID(CardID, CardTypeIDint)
                            CustomerPK = GetCustomerPK(CardID, CardTypeIDint, True)

                            If (CustomerPK > 0) Then
                                ' ensure that eligibility conditions of this offer are met for this customer
                                Dim Offer As New Models.Offer
                                Offer.OfferID = LogixID
                                Offer.EngineID = OfferEngineID
                                Offer = Offers.LoadOfferDetails(Offer, LoadOfferOptions.AllRegularConditions)
                                If (Offer.CustomerGroupConditions IsNot Nothing) Then
                                    Offers.SatisfyCustomerCondition(Offer.CustomerGroupConditions, CustomerPK, ExtInterfaceID, CardID)
                                End If

                                If MyCommon.Fetch_InterfaceOption(102) = "1" Then
                                    IncludePendingPoints = True
                                End If
                                For Each PointCondition As Models.PointsCondition In Offer.PointsProgramConditions
                                    PointsBalance = MyPoints.GetBalance(CustomerPK, PointCondition.ProgramID)
                                    If IncludePendingPoints Then
                                        ' We can't lock the involved tables and processes, so narrow the window as much as possible
                                        ' by checking the balance a second time.
                                        Do
                                            PointsBalanceOriginal = PointsBalance
                                            PointsPending = MyPoints.GetInProcessAdjustment(CustomerPK, PointCondition.ProgramID)

                                            PointsBalance = MyPoints.GetBalance(CustomerPK, PointCondition.ProgramID)
                                            ' Check to see if original balance has NOT changed since getting it the first time
                                            If PointsBalance = PointsBalanceOriginal Then
                                                PointsBalance += PointsPending
                                                Exit Do
                                            End If
                                        Loop
                                    End If
                                    PointsRequired = PointCondition.Quantity
                                    PointsAdj = Decimal.Subtract(PointsRequired, PointsBalance)
                                    ' GrantTypeID 1 =
                                    ' GrantTypeID 2 >
                                    Dim PointToAdjust As Decimal = 0.0
                                    If PointsAdj <> 0 Then
                                        If (OfferEngineID = 0) Then
                                            If (PointCondition.GrantTypeID = 1) Then
                                                PointToAdjust = PointsAdj
                                            End If
                                            If (PointCondition.GrantTypeID = 2) Then
                                                If (PointsAdj > 0) Then PointToAdjust = PointsAdj
                                            End If
                                        End If
                                        If ((OfferEngineID = 2 Or OfferEngineID = 9) And PointsAdj > 0) Then
                                            PointToAdjust = PointsAdj
                                        End If
                                        If PointToAdjust <> 0 Then
                                            Offers.SatisfyPointsCondition(PointCondition, CustomerPK, PointToAdjust)
                                        End If
                                    End If
                                Next

                                For Each StoredValueCondition As SVCondition In Offer.SVProgramConditions
                                    PointsBalance = StoredValue.GetQuantityBalance(CustomerPK, StoredValueCondition.ProgramID)
                                    PointsRequired = StoredValueCondition.Quantity
                                    PointsAdj = Decimal.Subtract(PointsRequired, PointsBalance)
                                    If PointsAdj > 0 Then
                                        StoredValue.AdjustStoredValue(1, StoredValueCondition.ProgramID, CustomerPK, PointsAdj.ToString)
                                    End If
                                Next
                                Added = True
                                SendDataToExchange(CustomerPK, Offer) 'Send the group membership data to RabbitMQ exchange to be forwarded to Store customer proxy
                            ElseIf CustomerPK = -2 Then
                                ErrorCode = ERROR_CODES.ERROR_INVALID_CUSTOMER_ID
                                ErrorMsg = "Customer: " & CardID & " is not in valid format."
                                ErrorMsgLog = "Customer: " & Copient.MaskHelper.MaskCard(CardID, -1) & " is not in valid format."
                                ErrorMsgLogUnmasked = "Customer: " & CardID & " is not in valid format."
                            Else
                                ErrorCode = ERROR_CODES.ERROR_INVALID_CUSTOMER_ID
                                ErrorMsg = "Customer: " & CardID & " does not exist and failed to create a new customer record."
                                ErrorMsgLog = "Customer: " & Copient.MaskHelper.MaskCard(CardID, CardTypeID) & " does not exist and failed to create a new customer record."
                                ErrorMsgLogUnmasked = "Customer: " & CardID & " does not exist and failed to create a new customer record."
                            End If
                        ElseIf Response = CardValidationResponse.INVALIDCARDTYPEFORMAT Or Response = CardValidationResponse.CARDTYPENOTFOUND Then
                            ErrorCode = ERROR_CODES.ERROR_INVALID_CARD_TYPE_ID
                            ErrorMsg = "Card type ID: " & CardTypeID & " is an invalid card type."
                        ElseIf Response = CardValidationResponse.CARDIDNOTNUMERIC Then
                            ErrorCode = ERROR_CODES.ERROR_INVALID_CUSTOMER_ID
                            ErrorMsg = "Card ID: " & CardID & " should be numeric."
                            ErrorMsgLog = "Card ID: " & Copient.MaskHelper.MaskCard(CardID, CardTypeID) & " should be numeric."
                            ErrorMsgLogUnmasked = "Card ID: " & CardID & " should be numeric."
                        ElseIf Response = CardValidationResponse.INVALIDCARDFORMAT Then
                            ErrorCode = ERROR_CODES.ERROR_INVALID_CUSTOMER_ID
                            ErrorMsg = "Card ID: " & CardID & " is invalid."
                            ErrorMsgLog = "Card ID: " & Copient.MaskHelper.MaskCard(CardID, CardTypeID) & " is invalid."
                            ErrorMsgLogUnmasked = "Card ID: " & CardID & " is invalid."
                        ElseIf Response = CardValidationResponse.ERROR_APPLICATION Then
                            ErrorCode = ERROR_CODES.ERROR_APPLICATION
                            ErrorMsg = "Application Error encountered."
                        End If
                    Else
                        ErrorCode = ERROR_CODES.ERROR_OFFER_ENGINE_ID_INVALID
                        ErrorMsg = "Invalid Offer EngineID: " & OfferEngineID
                    End If
                Else
                    ErrorCode = ERROR_CODES.ERROR_OFFER_DOES_NOT_EXIST
                    ErrorMsg = "Offer " & ClientOfferID & " does not exist."
                End If
            Else
                ErrorCode = ERROR_CODES.ERROR_XML_BAD_SOURCE_ID
                ErrorMsg = "Unrecognized ExternalSourceID: " & ExternalSourceID
            End If
        Catch ex As Exception
            ErrorCode = ERROR_CODES.ERROR_APPLICATION
            ErrorMsg = "Application Error encountered: " & ex.ToString
        Finally
            If (String.IsNullOrWhiteSpace(ErrorMsgLog)) Then
                AppendToLog(ExternalSourceID, CalledBy, "", "ClientOfferID=" & ClientOfferID & "; CustomerID=" & Copient.MaskHelper.MaskCard(CardID, CardTypeID), ErrorCode, ErrorMsg)
                AppendToLog(ExternalSourceID, CalledBy, "", "ClientOfferID=" & ClientOfferID & "; CustomerID=" & CardID, ErrorCode, ErrorMsgLogUnmasked, True)
            Else
                AppendToLog(ExternalSourceID, CalledBy, "", "ClientOfferID=" & ClientOfferID & "; CustomerID=" & Copient.MaskHelper.MaskCard(CardID, CardTypeID), ErrorCode, ErrorMsgLog)
                AppendToLog(ExternalSourceID, CalledBy, "", "ClientOfferID=" & ClientOfferID & "; CustomerID=" & CardID, ErrorCode, ErrorMsgLogUnmasked, True)
            End If


            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try

        Return WriteOfferXmlResponse(LogixID, ClientOfferID, ErrorCode, ErrorMsg, RESPONSE_TYPES.ADD_CUSTOMER_OFFER, Added)

    End Function

    <WebMethod()>
    Public Function AddCustomersToOffer(ByVal ExternalSourceID As String, ByVal CustomerIDs As String, ByVal ClientOfferID As String) As String
        Dim Added As Boolean = False
        Dim LogixID As Long
        Dim ErrorCode As ERROR_CODES = ERROR_CODES.ERROR_NONE
        Dim ErrorMsg As String = ""
        Dim CustomerPK As Long = 0
        Dim AutoDeploy As Boolean = False
        Dim OfferEngineID As Long = -1

        Try

            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            ExtInterfaceID = GetExtInterfaceID(ExternalSourceID, AutoDeploy)

            If (ExtInterfaceID > 0) Then
                If DoesOfferExist(ClientOfferID, ExtInterfaceID, LogixID, OfferEngineID) Then
                    If OfferEngineID = 0 Or OfferEngineID = 2 Or OfferEngineID = 9 Then
                        CustomerIDs = CustomerIDs.Replace("|", ControlChars.CrLf)
                        SatisfyCustomerCondition(CustomerIDs, LogixID, OfferEngineID, ErrorCode, ErrorMsg)
                        Added = (ErrorCode = ERROR_CODES.ERROR_NONE)
                    Else
                        ErrorCode = ERROR_CODES.ERROR_OFFER_ENGINE_ID_INVALID
                        ErrorMsg = "Invalid Offer EngineID: " & OfferEngineID
                    End If
                Else
                    ErrorCode = ERROR_CODES.ERROR_OFFER_DOES_NOT_EXIST
                    ErrorMsg = "Offer " & ClientOfferID & " does not exist."
                End If
            Else
                ErrorCode = ERROR_CODES.ERROR_XML_BAD_SOURCE_ID
                ErrorMsg = "Unrecognized ExternalSourceID: " & ExternalSourceID
            End If
        Catch ex As Exception
            ErrorCode = ERROR_CODES.ERROR_APPLICATION
            ErrorMsg = "Application Error encountered: " & ex.ToString
        Finally
            AppendToLog(ExternalSourceID, "AddCustomersToOffer", "", "ClientOfferID=" & ClientOfferID, ErrorCode, ErrorMsg)
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try

        Return WriteOfferXmlResponse(LogixID, ClientOfferID, ErrorCode, ErrorMsg, RESPONSE_TYPES.ADD_CUSTOMERS_OFFER, Added)

    End Function

    <WebMethod()>
    Public Function AddCustomers(ByVal ExternalSourceID As String, ByVal CustomerOfferData As String) As String
        Return HandleOfferCustomers(ExternalSourceID, CustomerOfferData, RESPONSE_TYPES.ADD_CUSTOMERS, False, True)
    End Function

    <WebMethod()>
    Public Function ClipBundle(ByVal ExternalSourceID As String, ByVal CustomerOfferData As String) As String
        Return HandleOfferCustomers(ExternalSourceID, CustomerOfferData, RESPONSE_TYPES.CLIP_BUNDLE, True)
    End Function

    <WebMethod()>
    Public Function AddCustomerToOffer(ByVal ExternalSourceID As String, ByVal CustomerID As String, ByVal ClientOfferID As String) As String
        ' Use the default CardTypeID
        StartUp()
        Return HandleAddCustomerCardToOffer(ExternalSourceID, CustomerID, GetDefaultCardTypeID(), ClientOfferID, "AddCustomerToOffer")
        Shutdown()
    End Function

    <WebMethod()>
    Public Function AddCustomerCardToOffer(ByVal ExternalSourceID As String, ByVal CardID As String, ByVal CardTypeID As String, ByVal ClientOfferID As String) As String
        StartUp()
        Return HandleAddCustomerCardToOffer(ExternalSourceID, CardID, CardTypeID, ClientOfferID, "AddCustomerCardToOffer")
        Shutdown()
    End Function

    <WebMethod()>
    Private Function HandleRemoveCustomerCardFromOffer(ByVal ExternalSourceID As String, ByVal CardID As String,
                                              ByVal CardTypeID As String, ByVal ClientOfferID As String, ByVal CalledBy As String) As String
        Dim Removed As Boolean = False
        Dim LogixID As Long
        Dim ErrorCode As ERROR_CODES = ERROR_CODES.ERROR_NONE
        Dim ErrorMsg As String = ""
        Dim ErrorMsgLog As String = ""
        Dim CustomerPK As Long = 0
        Dim AutoDeploy As Boolean = False
        Dim OfferEngineID As Long = -1
        Dim CardTypeIDint As Integer
        Dim Offers As IOffer
        Dim ErrorMsgLogUnmasked As String = ""

        Try

            Offers = CurrentRequest.Resolver.Resolve(Of IOffer)()
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            ExtInterfaceID = GetExtInterfaceID(ExternalSourceID, AutoDeploy)

            If (ExtInterfaceID > 0) Then
                If DoesOfferExist(ClientOfferID, ExtInterfaceID, LogixID, OfferEngineID) Then
                    If (Integer.TryParse(CardTypeID, CardTypeIDint)) Then
                        CardID = MyCommon.Pad_ExtCardID(CardID, CardTypeIDint)
                    End If
                    If (Integer.TryParse(CardTypeID, CardTypeIDint)) Then
                        CustomerPK = GetCustomerPK(CardID, CardTypeIDint, False)

                        If (CustomerPK > 0) Then
                            ' remove the customer from the group so the customer will no longer qualify for it.
                            Select Case OfferEngineID
                                Case 0
                                    Offers.RemoveCustomerConditionCM(LogixID, CustomerPK, ExtInterfaceID, CardID, lUserId)
                                    Removed = True
                                Case 2, 9
                                    Offers.RemoveCustomerCondition(LogixID, CustomerPK, ExtInterfaceID, CardID)
                                    Removed = True
                                Case Else
                                    ErrorCode = ERROR_CODES.ERROR_OFFER_ENGINE_ID_INVALID
                                    ErrorMsg = "Invalid Offer EngineID: " & OfferEngineID
                            End Select
                        Else
                            If CustomerPK = -1 Then
                                ErrorCode = ERROR_CODES.ERROR_INVALID_CARD_TYPE_ID
                                ErrorMsg = "Card type ID: " & CardTypeID & " is an invalid card type."
                            ElseIf CustomerPK = -2 Then
                                ErrorCode = ERROR_CODES.ERROR_INVALID_CUSTOMER_ID
                                ErrorMsg = "Customer: " & CardID & " is not in valid format."
                                ErrorMsgLog = "Customer: " & Copient.MaskHelper.MaskCard(CardID, -1) & " is not in valid format."
                                ErrorMsgLogUnmasked = "Customer: " & CardID & " is not in valid format."
                            Else
                                ErrorCode = ERROR_CODES.ERROR_INVALID_CUSTOMER_ID
                                ErrorMsg = "Customer: " & CardID & " (Card type ID: " & CardTypeID & ") does not exist in Logix."
                                ErrorMsgLog = "Customer: " & Copient.MaskHelper.MaskCard(CardID, CardTypeID) & " (Card type ID: " & CardTypeID & ") does not exist in Logix."
                                ErrorMsgLogUnmasked = "Customer: " & CardID & " (Card type ID: " & CardTypeID & ") does not exist in Logix."
                            End If
                        End If
                    Else
                        ErrorCode = ERROR_CODES.ERROR_INVALID_CARD_TYPE_ID
                        ErrorMsg = "Card type ID: " & CardTypeID & " is an invalid card type."
                    End If
                Else
                    ErrorCode = ERROR_CODES.ERROR_OFFER_DOES_NOT_EXIST
                    ErrorMsg = "Offer " & ClientOfferID & " does not exist in Logix."
                End If
            Else
                ErrorCode = ERROR_CODES.ERROR_XML_BAD_SOURCE_ID
                ErrorMsg = "Unrecognized ExternalSourceID: " & ExternalSourceID
            End If

        Catch ex As Exception
            ErrorCode = ERROR_CODES.ERROR_APPLICATION
            ErrorMsg = "Application Error encountered: " & ex.ToString
        Finally
            If (Not String.IsNullOrWhiteSpace(ErrorMsgLog)) Then
                AppendToLog(ExternalSourceID, CalledBy, "", "ClientOfferID=" & ClientOfferID &
                                          "; CardID=" & Copient.MaskHelper.MaskCard(CardID, CardTypeID) & "; CardTypeID=" & CardTypeID, ErrorCode, ErrorMsgLog)
                AppendToLog(ExternalSourceID, CalledBy, "", "ClientOfferID=" & ClientOfferID &
                                          "; CardID=" & CardID & "; CardTypeID=" & CardTypeID, ErrorCode, ErrorMsgLogUnmasked, True)
            Else
                AppendToLog(ExternalSourceID, CalledBy, "", "ClientOfferID=" & ClientOfferID &
                                          "; CardID=" & Copient.MaskHelper.MaskCard(CardID, CardTypeID) & "; CardTypeID=" & CardTypeID, ErrorCode, ErrorMsg)
                AppendToLog(ExternalSourceID, CalledBy, "", "ClientOfferID=" & ClientOfferID &
                                          "; CardID=" & CardID & "; CardTypeID=" & CardTypeID, ErrorCode, ErrorMsgLogUnmasked, True)
            End If


            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try

        Return WriteOfferXmlResponse(LogixID, ClientOfferID, ErrorCode, ErrorMsg, RESPONSE_TYPES.REMOVE_CUSTOMER_OFFER, Removed)
    End Function

    <WebMethod()>
    Public Function RemoveCustomerFromOffer(ByVal ExternalSourceID As String, ByVal CustomerID As String, ByVal ClientOfferID As String) As String
        ' Use the default CardTypeID
        StartUp()
        Return HandleRemoveCustomerCardFromOffer(ExternalSourceID, CustomerID, GetDefaultCardTypeID(), ClientOfferID, "RemoveCustomerFromOffer")
        Shutdown()
    End Function

    <WebMethod()>
    Public Function RemoveCustomerCardFromOffer(ByVal ExternalSourceID As String, ByVal CustomerID As String, ByVal CardTypeID As String, ByVal ClientOfferID As String) As String
        StartUp()
        Return HandleRemoveCustomerCardFromOffer(ExternalSourceID, CustomerID, CardTypeID, ClientOfferID, "RemoveCustomerCardFromOffer")
        Shutdown()
    End Function
    Private Sub SendDataToExchange(ByVal custPK As Long, ByVal offer As Models.Offer)
        Dim locations As List(Of String)

        locations = GetCustomerLocations(custPK) ' get locations that the customer has visited to
        If locations.Count > 0 Then

            Dim custDataToSend As Models.CustGroupMembership = New Models.CustGroupMembership()
            Dim routingKey As String

            routingKey = String.Join(".", locations.ToArray()) ' send location ids as the routing key to the RabbitMQ exchange
            custDataToSend.custPK = custPK
            custDataToSend.customerGroups = New List(Of CustGroups)

            For Each conditionDetail As CustomerConditionDetails In offer.CustomerGroupConditions.IncludeCondition
                Dim gMembership As Models.CustGroups = New CustGroups()
                gMembership.customerGroupID = conditionDetail.CustomerGroup.CustomerGroupID
                gMembership.manual = GetManualFlagForGroupMembership(custPK, gMembership.customerGroupID).ToString.ToLower()
                custDataToSend.customerGroups.Add(gMembership)
            Next
            Dim dataToSend As String = JsonConvert.SerializeObject(custDataToSend)
            ConnectToMessageServer(dataToSend, routingKey)
        Else
            Copient.Logger.Write_Log(RejectionLogFile, "There are no locations associated to the customer to send the group membership data", True)
        End If

    End Sub
    Private Function GetManualFlagForGroupMembership(ByVal custPK As Long, ByVal cgID As Long) As Boolean
        Dim manual As Boolean = False
        Dim dt As DataTable
        MyCommon.QueryStr = "select Manual from GroupMembership where CustomerPK=@custPK and CustomerGroupID=@cgID"
        MyCommon.DBParameters.Add("@custPK", SqlDbType.BigInt).Value = custPK
        MyCommon.DBParameters.Add("@cgID", SqlDbType.BigInt).Value = cgID
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            manual = dt.Rows(0).Item("Manual")
        End If
        Return manual
    End Function
    Private Sub ConnectToMessageServer(ByVal dataToSend As String, ByVal routingKey As String)
        Try
            Dim connection = messagingService.Connect(GetHost(), GetPort(), GetUsername(), GetPassword(), "ams.exchange.couponclip", routingKey, "", "topic", False)
            If connection.ResultType <> AMSResultType.Success Then
                Copient.Logger.Write_Log(RejectionLogFile, connection.MessageString, True)
            Else
                messagingService.Send(dataToSend) ' publish the data if the connection is successful
            End If
        Catch ex As Exception
            Copient.Logger.Write_Log(RejectionLogFile, ex.ToString(), True)
        Finally
            If messagingService IsNot Nothing Then messagingService.Disconnect() 'disconnect from the server
        End Try

    End Sub
    Private Function GetCustomerLocations(ByVal custPK As Integer) As List(Of String)
        Dim locations As List(Of String) = New List(Of String)
        Dim dt As DataTable
        MyCommon.QueryStr = "select LocationID from CustomerLocations where CustomerPK=@custPK"
        MyCommon.DBParameters.Add("@custPK", SqlDbType.Int).Value = custPK
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            locations = dt.Rows.Cast(Of DataRow).Select(Function(dr) dr(0).ToString).ToList
        End If
        Return locations
    End Function
    Private Function GetSystemOptionValueOrDefault(ByVal defaultValue As String, ByVal optionID As Integer) As String
        Dim value As String
        If optionID = 163 Then
            Dim MyCryptLib As New Copient.CryptLib()
            value = MyCryptLib.SQL_StringDecrypt(MyCommon.Fetch_UE_SystemOption(optionID))
        Else
            value = MyCommon.Fetch_UE_SystemOption(optionID)
        End If

        If Not String.IsNullOrEmpty(value) Then
            Return value
        Else
            Return defaultValue
        End If
    End Function
    Private Function GetHost() As String
        Dim defaultHost As String = "localhost"
        Dim optionID As Integer = 160
        Return GetSystemOptionValueOrDefault(defaultHost, optionID)
    End Function
    Private Function GetPort() As String
        Dim defaultPort As String = "5672"
        Dim optionID As Integer = 161
        Return GetSystemOptionValueOrDefault(defaultPort, optionID)
    End Function
    Private Function GetUsername() As String
        Dim defaultUsername As String = "ams"
        Dim optionID As Integer = 162
        Return GetSystemOptionValueOrDefault(defaultUsername, optionID)
    End Function
    Private Function GetPassword() As String
        Dim defaultPassword As String = "ncr.ams"
        Dim optionID As Integer = 163
        Return GetSystemOptionValueOrDefault(defaultPassword, optionID)
    End Function
    Private Function GetRoutingKey() As String
        Dim defaultRoutingKey As String = "ams.routing_key.download"
        Dim optionID As Integer = 165
        Return GetSystemOptionValueOrDefault(defaultRoutingKey, optionID)
    End Function
    Private Sub StartUp()
        CurrentRequest.Resolver.AppName = "External Offer Connector"
        Dim common As CMS.AMS.Common = CurrentRequest.Resolver.Resolve(Of CMS.AMS.Common)()
        If (common.LRT_Connection_State() = ConnectionState.Closed) Then common.Open_LogixRT()
        If (common.LXS_Connection_State() = ConnectionState.Closed) Then common.Open_LogixXS()

        common.Set_AppInfo()
        CurrentRequest.Resolver.RegisterInstance(Of Copient.CommonInc)(MyCommon)
    End Sub

    Private Sub Shutdown()
        Dim common As CMS.AMS.Common = CurrentRequest.Resolver.Resolve(Of CMS.AMS.Common)()
        If (common.LRT_Connection_State() <> ConnectionState.Closed) Then common.Close_LogixRT()
        If (common.LXS_Connection_State() <> ConnectionState.Closed) Then common.Close_LogixXS()
    End Sub

    <WebMethod()>
    Public Function RemoveCustomersFromOffer(ByVal ExternalSourceID As String, ByVal CustomerIDs As String, ByVal ClientOfferID As String) As String
        Dim Removed As Boolean = False
        Dim LogixID As Long
        Dim ErrorCode As ERROR_CODES = ERROR_CODES.ERROR_NONE
        Dim ErrorMsg As String = ""
        Dim CustomerPK As Long = 0
        Dim AutoDeploy As Boolean = False
        Dim OfferEngineID As Long = -1
        Dim Log As Boolean = False
        Dim ErrorMsgLog As String = ""

        Try

            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            ExtInterfaceID = GetExtInterfaceID(ExternalSourceID, AutoDeploy)

            If (ExtInterfaceID > 0) Then
                If (CustomerIDs Is Nothing OrElse String.IsNullOrEmpty(CustomerIDs.Trim)) Then
                    ErrorCode = ERROR_CODES.ERROR_INVALID_CUSTOMER_ID
                    ErrorMsg = "Invalid Customer Id"
                    Return WriteOfferXmlResponse(LogixID, ClientOfferID, ErrorCode, ErrorMsg, RESPONSE_TYPES.REMOVE_CUSTOMERS_OFFER, Removed)
                ElseIf CustomerIDs.Contains("'") = True Or CustomerIDs.Contains(Chr(34)) = True Then
                    ErrorCode = ERROR_CODES.ERROR_INVALID_CUSTOMER_ID   '
                    ErrorMsg = "Invalid Customer Id"
                    Return WriteOfferXmlResponse(LogixID, ClientOfferID, ErrorCode, ErrorMsg, RESPONSE_TYPES.REMOVE_CUSTOMERS_OFFER, Removed)
                End If

                If DoesOfferExist(ClientOfferID, ExtInterfaceID, LogixID, OfferEngineID) Then
                    CustomerIDs = ValidateCustId(CustomerIDs.Trim, ExternalSourceID, ClientOfferID, ErrorCode, ErrorMsg, Log)
                    If Not String.IsNullOrEmpty(CustomerIDs) Then
                        If OfferEngineID = Copient.CommonInc.InstalledEngines.CM Then
                            CMRemoveCustomersFromCondition(CustomerIDs, LogixID, ErrorCode, ErrorMsg)
                        Else
                            CPERemoveCustomersFromCondition(CustomerIDs, LogixID, ErrorCode, ErrorMsg)
                        End If
                        Removed = (ErrorCode = ERROR_CODES.ERROR_NONE)
                    Else
                        ErrorCode = ERROR_CODES.ERROR_INVALID_CUSTOMER_ID
                        ErrorMsg = "Invalid Customer Id"
                    End If
                Else
                    ErrorCode = ERROR_CODES.ERROR_OFFER_DOES_NOT_EXIST
                    ErrorMsg = "Offer " & ClientOfferID & " does not exist."
                    ErrorMsgLog = "Offer " & ClientOfferID & " does not exist."
                End If
            Else
                ErrorCode = ERROR_CODES.ERROR_XML_BAD_SOURCE_ID
                ErrorMsg = "Unrecognized ExternalSourceID: " & ExternalSourceID
            End If

        Catch ex As Exception
            ErrorCode = ERROR_CODES.ERROR_APPLICATION
            ErrorMsg = "Application Error encountered: " & ex.ToString
        Finally
            If Not Log Then
                If (Not String.IsNullOrWhiteSpace(ErrorMsgLog)) Then
                    AppendToLog(ExternalSourceID, "RemoveCustomersFromOffer", "", "ClientOfferID=" & ClientOfferID, ErrorCode, ErrorMsgLog)
                Else
                    AppendToLog(ExternalSourceID, "RemoveCustomersFromOffer", "", "ClientOfferID=" & ClientOfferID, ErrorCode, ErrorMsg)
                End If

            End If
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try

        Return WriteOfferXmlResponse(LogixID, ClientOfferID, ErrorCode, ErrorMsg, RESPONSE_TYPES.REMOVE_CUSTOMERS_OFFER, Removed)
    End Function

    <WebMethod()>
    Public Function RemoveCustomers(ByVal ExternalSourceID As String, ByVal CustomerOfferData As String) As String
        Return HandleOfferCustomers(ExternalSourceID, CustomerOfferData, RESPONSE_TYPES.REMOVE_CUSTOMERS, False, True)
    End Function

    <WebMethod()>
    Public Function UnclipBundle(ByVal ExternalSourceID As String, ByVal CustomerOfferData As String) As String
        Return HandleOfferCustomers(ExternalSourceID, CustomerOfferData, RESPONSE_TYPES.UNCLIP_BUNDLE, True)
    End Function

    <WebMethod()>
    Public Function AddProductToOffer(ByVal ExternalSourceID As String, ByVal ClientProductID As String, ByVal ClientOfferID As String) As String
        Dim ErrorCode As ERROR_CODES = ERROR_CODES.ERROR_NONE
        Dim ErrorMsg As String = ""
        Dim dt, dt2 As DataTable
        Dim row As DataRow
        Dim LogixID As Long
        Dim ExcludedGroup As Boolean
        Dim ProductGroupID As Integer
        Dim ProductDesc As String = ""
        Dim ProductID As Long
        Dim OutputStatus As Integer
        Dim AutoDeploy As Boolean = False
        Dim OfferEngineId As Long = -1
        Dim NumericOnly As Boolean = False
        Dim ValidIdentifier As Boolean = False

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

            ExtInterfaceID = GetExtInterfaceID(ExternalSourceID, AutoDeploy)
            NumericOnly = (MyCommon.Fetch_SystemOption(97) = "1")
            ValidIdentifier = (Not NumericOnly) OrElse IsStringNumeric(ClientProductID)

            If (ValidIdentifier AndAlso ExtInterfaceID > 0) Then
                If DoesOfferExist(ClientOfferID, ExtInterfaceID, LogixID, OfferEngineId) Then
                    ClientProductID = HandleProductIdPadding(ClientProductID)

                    Select Case OfferEngineId
                        Case 0
                            MyCommon.QueryStr = "Select distinct OC.LinkID as ProductGroupID, 0 as ExcludedProducts " &
                                                "from OfferConditions OC with (NoLock) " &
                                                "inner join ProductGroups PG with (NoLock) on PG.ProductGroupId=OC.LinkId " &
                                                "where OC.Deleted=0 And OC.ConditionTypeId=2 And PG.Deleted=0 And OC.OfferID = @OfferID " &
                                                "and PG.AnyProduct=0 " &
                                                "union " &
                                                "Select distinct OC.ExcludedID as ProductGroupID, 1 as ExcludedProducts " &
                                                "from OfferConditions OC with (NoLock) " &
                                                "inner join ProductGroups PG with (NoLock) on PG.ProductGroupId=OC.ExcludedID " &
                                                "where OC.Deleted=0 And OC.ConditionTypeId=2 And PG.Deleted=0 And OC.OfferID = @OfferID " &
                                                "and PG.AnyProduct=0 " &
                                                "union " &
                                                "Select distinct ORW.ProductGroupId as ProductGroupID, 0 as ExcludedProducts " &
                                                "from OfferRewards ORW with (NoLock) " &
                                                "inner join ProductGroups PG with (NoLock) on PG.ProductGroupId=ORW.ProductGroupId " &
                                                "where ORW.Deleted=0 And PG.Deleted=0 And ORW.OfferID = @OfferID " &
                                                "union " &
                                                "Select distinct ORW.ExcludedProdGroupId as ProductGroupID, 1 as ExcludedProducts " &
                                                "from OfferRewards ORW with (NoLock) " &
                                                "inner join ProductGroups PG with (NoLock) on PG.ProductGroupId=ORW.ExcludedProdGroupId " &
                                                "where ORW.Deleted=0 And PG.Deleted=0 And ORW.OfferID = @OfferID " &
                                                "order by ExcludedProducts desc;"
                            MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = LogixID
                        Case 2, 9
                            MyCommon.QueryStr = "Select IPG.ProductGroupID, IPG.ExcludedProducts from CPE_IncentiveProductGroups IPG with (NoLock) " &
                                                "inner join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID = IPG.RewardOptionID " &
                                                "inner join ProductGroups PG with (NoLock) on PG.ProductGroupID = IPG.ProductGroupID " &
                                                "where IPG.Deleted=0 and PG.Deleted=0 and PG.AnyProduct=0 and RO.Deleted=0 and RO.IncentiveID = @IncentiveID"
                            MyCommon.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = LogixID
                        Case Else
                            ErrorCode = ERROR_CODES.ERROR_OFFER_ENGINE_ID_INVALID
                            ErrorMsg = "Invalid Offer EngineID: " & OfferEngineId
                    End Select

                    If ErrorMsg = "" Then
                        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                        If (dt.Rows.Count > 0) Then
                            For Each row In dt.Rows
                                ProductGroupID = MyCommon.NZ(row.Item("ProductGroupID"), 0)
                                ExcludedGroup = MyCommon.NZ(row.Item("ExcludedProducts"), False)

                                MyCommon.QueryStr = "select ProductID, Description from Products with (NoLock) where ExtProductID = @ExtProductID "
                                MyCommon.DBParameters.Add("@ExtProductID", SqlDbType.NVarChar).Value = ClientProductID
                                dt2 = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                                If (dt2.Rows.Count > 0) Then
                                    ProductID = MyCommon.NZ(dt2.Rows(0).Item("ProductID"), 0)
                                    ProductDesc = MyCommon.NZ(dt2.Rows(0).Item("Description"), "")
                                End If

                                MyCommon.QueryStr = "Select PKID from ProdGroupItems where ProductGroupID = @ProductGroupID " &
                                                    "  and ProductID = @ProductID and Deleted=0"
                                MyCommon.DBParameters.Add("@ProductGroupID", SqlDbType.BigInt).Value = ProductGroupID
                                MyCommon.DBParameters.Add("@ProductID", SqlDbType.BigInt).Value = ProductID
                                dt2 = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)

                                ' remove the product from the excluded group or add the product to the included group
                                If (dt2.Rows.Count > 0 AndAlso ExcludedGroup) Then
                                    MyCommon.Open_LogixRT()
                                    MyCommon.QueryStr = "dbo.pt_ProdGroupItems_DeleteItem"
                                    MyCommon.Open_LRTsp()
                                    MyCommon.LRTsp.Parameters.Add("@ExtProductID", SqlDbType.NVarChar, 120).Value = ClientProductID
                                    MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt).Value = ProductGroupID
                                    MyCommon.LRTsp.Parameters.Add("@ProductTypeID", SqlDbType.Int).Value = 1 ' UPC
                                    MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
                                    MyCommon.LRTsp.ExecuteNonQuery()
                                    OutputStatus = MyCommon.LRTsp.Parameters("@Status").Value
                                    MyCommon.Close_LRTsp()
                                    If (OutputStatus <> 0) Then
                                        ErrorCode = ERROR_CODES.ERROR_ADD_PRODUCT_FAILED
                                        If ErrorMsg <> "" Then ErrorMsg &= "; "
                                        ErrorMsg &= "Error Encountering while attempting to remove ProductID " & ClientProductID & " from offer " & ClientOfferID
                                    Else
                                        MyCommon.Activity_Log(5, ProductGroupID, 1, Copient.PhraseLib.Lookup("history.pgroup-remove", 1) & " " & ClientProductID)
                                    End If
                                ElseIf (dt2.Rows.Count = 0 AndAlso Not ExcludedGroup) Then
                                    AddProductToGroup(ProductGroupID, ClientProductID, ProductDesc, ClientOfferID, ErrorCode, ErrorMsg)
                                End If

                                If ErrorCode = ERROR_CODES.ERROR_NONE Then
                                    MyCommon.QueryStr = "update productgroups with (RowLock) set updatelevel=updatelevel+1, lastupdate=getdate() where ProductGroupID = @ProductGroupID"
                                    MyCommon.DBParameters.Add("@ProductGroupID", SqlDbType.BigInt).Value = ProductGroupID
                                    MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
                                End If
                            Next
                        Else
                            ProductGroupID = CreateProductGroup(LogixID, ClientOfferID, False, ErrorCode, ErrorMsg)
                            If (ProductGroupID > 0) Then
                                MyCommon.QueryStr = "select Description from Products with (NoLock) where ExtProductID = @ExtProductID "
                                MyCommon.DBParameters.Add("@ExtProductID", SqlDbType.NVarChar).Value = ClientProductID.ConvertBlankIfNothing
                                dt2 = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                                If (dt2.Rows.Count > 0) Then
                                    ProductDesc = MyCommon.NZ(dt2.Rows(0).Item("Description"), "")
                                End If

                                AddProductToGroup(ProductGroupID, ClientProductID, ProductDesc, ClientOfferID, ErrorCode, ErrorMsg)
                            End If
                        End If
                    End If
                Else
                    ErrorCode = ERROR_CODES.ERROR_OFFER_DOES_NOT_EXIST
                    ErrorMsg = "Offer " & ClientOfferID & " does not exist."
                End If
            ElseIf Not ValidIdentifier Then
                ErrorCode = ERROR_CODES.ERROR_INVALID_FORMAT
                ErrorMsg = "Product identifier " & ClientProductID & " contains non-numeric characters. " &
                           "This installation's configuration setting for products only allow numeric characters in the product identifier."
            Else
                ErrorCode = ERROR_CODES.ERROR_XML_BAD_SOURCE_ID
                ErrorMsg = "Unrecognized ExternalSourceID: " & ExternalSourceID
            End If

        Catch ex As Exception
            ErrorCode = ERROR_CODES.ERROR_APPLICATION
            ErrorMsg = "Application Error encountered: " & ex.ToString
        Finally
            AppendToLog(ExternalSourceID, "AddProductToOffer", "", "ClientOfferID=" & ClientOfferID & "; ClientProductID=" & ClientProductID, ErrorCode, ErrorMsg)
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try

        Return WriteOfferXmlResponse(LogixID, ClientOfferID, ErrorCode, ErrorMsg, RESPONSE_TYPES.ADD_PRODUCT_OFFER, (ErrorCode = ERROR_CODES.ERROR_NONE))
    End Function

    <WebMethod()>
    Public Function RemoveProductFromOffer(ByVal ExternalSourceID As String, ByVal ClientProductID As String, ByVal ClientOfferID As String) As String
        Dim ErrorCode As ERROR_CODES = ERROR_CODES.ERROR_NONE
        Dim ErrorMsg As String = ""
        Dim dt, dt2 As DataTable
        Dim row As DataRow
        Dim LogixID, ROID As Long
        Dim ExcludedGroup, IsAnyProduct As Boolean
        Dim ProductGroupID, ExcludedGroupID As Integer
        Dim ProductID As Long
        Dim ProductDesc As String = ""
        Dim OutputStatus As Integer
        Dim AutoDeploy As Boolean = False
        Dim OfferEngineId As Long = -1

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

            ExtInterfaceID = GetExtInterfaceID(ExternalSourceID, AutoDeploy)

            If (ExtInterfaceID > 0) Then
                If DoesOfferExist(ClientOfferID, ExtInterfaceID, LogixID, OfferEngineId) Then
                    Select Case OfferEngineId
                        Case 0
                            CmRemoveProductFromOffer(ClientProductID, ClientOfferID, LogixID, ErrorCode, ErrorMsg)
                        Case 2, 9
                            ' validate that the product exists
                            ClientProductID = HandleProductIdPadding(ClientProductID)
                            MyCommon.QueryStr = "select ProductID, Description from Products with (NoLock) where ExtProductID = @ExtProductID "
                            MyCommon.DBParameters.Add("@ExtProductID", SqlDbType.NVarChar).Value = ClientProductID.ConvertBlankIfNothing
                            dt2 = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                            If (dt2.Rows.Count > 0) Then
                                ProductID = MyCommon.NZ(dt2.Rows(0).Item("ProductID"), 0)
                                ProductDesc = MyCommon.NZ(dt2.Rows(0).Item("Description"), "")

                                MyCommon.QueryStr = "Select IPG.ProductGroupID, IPG.ExcludedProducts, IPG.RewardOptionID, PG.AnyProduct from CPE_IncentiveProductGroups IPG with (NoLock) " &
                                                    "inner join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID = IPG.RewardOptionID " &
                                                    "inner join ProductGroups PG with (NoLock) on PG.ProductGroupID = IPG.ProductGroupID " &
                                                    "where IPG.Deleted=0 and PG.Deleted=0 and RO.Deleted=0 and RO.IncentiveID = @IncentiveID"
                                MyCommon.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = LogixID
                                dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                                If (dt.Rows.Count > 0) Then
                                    For Each row In dt.Rows
                                        ProductGroupID = MyCommon.NZ(row.Item("ProductGroupID"), 0)
                                        ExcludedGroup = MyCommon.NZ(row.Item("ExcludedProducts"), False)
                                        IsAnyProduct = MyCommon.NZ(row.Item("AnyProduct"), False)
                                        ROID = MyCommon.NZ(row.Item("RewardOptionID"), 0)



                                        MyCommon.QueryStr = "Select PKID from ProdGroupItems where ProductGroupID = @ProductGroupID " &
                                                            "  and ProductID = @ProductID and Deleted=0"
                                        MyCommon.DBParameters.Add("@ProductGroupID", SqlDbType.BigInt).Value = ProductGroupID
                                        MyCommon.DBParameters.Add("@ProductID", SqlDbType.BigInt).Value = ProductID
                                        dt2 = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)

                                        ' remove the product from the included group
                                        If (dt2.Rows.Count > 0 AndAlso Not ExcludedGroup) Then
                                            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
                                            MyCommon.QueryStr = "dbo.pt_ProdGroupItems_DeleteItem"
                                            MyCommon.Open_LRTsp()
                                            MyCommon.LRTsp.Parameters.Add("@ExtProductID", SqlDbType.NVarChar, 120).Value = ClientProductID
                                            MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt).Value = ProductGroupID
                                            MyCommon.LRTsp.Parameters.Add("@ProductTypeID", SqlDbType.Int).Value = 1 ' UPC
                                            MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
                                            MyCommon.LRTsp.ExecuteNonQuery()
                                            OutputStatus = MyCommon.LRTsp.Parameters("@Status").Value
                                            MyCommon.Close_LRTsp()
                                            If (OutputStatus <> 0) Then
                                                ErrorCode = ERROR_CODES.ERROR_REMOVE_PRODUCT_FAILED
                                                If ErrorMsg <> "" Then ErrorMsg &= "; "
                                                ErrorMsg &= "Error Encountering while attempting to remove ProductID " & ClientProductID & " from offer " & ClientOfferID
                                            Else
                                                MyCommon.Activity_Log(5, ProductGroupID, 1, Copient.PhraseLib.Lookup("history.pgroup-remove", 1) & " " & ClientProductID)
                                            End If
                                        ElseIf (dt2.Rows.Count = 0 AndAlso IsAnyProduct) Then
                                            ' check if there is a pre-existing excluded group, if not then create one and assign product to it
                                            MyCommon.QueryStr = "Select IPG.ProductGroupID from CPE_IncentiveProductGroups IPG with (NoLock) " &
                                                                "inner join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID = IPG.RewardOptionID " &
                                                                "inner join ProductGroups PG with (NoLock) on PG.ProductGroupID = IPG.ProductGroupID " &
                                                                "where IPG.Deleted=0 and IPG.ExcludedProducts=1 and PG.Deleted=0 and PG.AnyProduct=0 and RO.Deleted=0 and RO.IncentiveID = @IncentiveID"
                                            MyCommon.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = LogixID
                                            dt2 = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)

                                            If (dt2.Rows.Count > 0) Then
                                                ExcludedGroupID = MyCommon.NZ(dt2.Rows(0).Item("ProductGroupID"), 0)
                                            Else
                                                ' create new product group to be used as an exclusion group for this offer
                                                MyCommon.QueryStr = "dbo.pt_ProductGroups_Insert"
                                                MyCommon.Open_LRTsp()
                                                MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 200).Value = "Excluded Product Group for Offer " & LogixID
                                                MyCommon.LRTsp.Parameters.Add("@AnyProduct", SqlDbType.Bit).Value = 0
                                                MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                                                MyCommon.LRTsp.ExecuteNonQuery()
                                                ExcludedGroupID = MyCommon.LRTsp.Parameters("@ProductGroupID").Value
                                                MyCommon.Close_LRTsp()

                                                If (ExcludedGroupID > 0) Then
                                                    MyCommon.Activity_Log(5, ExcludedGroupID, 1, Copient.PhraseLib.Lookup("history.pgroup-create", 1))

                                                    ' assign the excluded product group to the offer
                                                    MyCommon.QueryStr = "insert into CPE_IncentiveProductGroups (RewardOptionID,ProductGroupID,ExcludedProducts,Deleted,LastUpdate,RequiredFromTemplate,TCRMAStatusFlag,Disqualifier) " &
                                                                        " values(@RewardOptionID, @ProductGroupID, 1, 0, getdate(), 0, 3, 0)"
                                                    MyCommon.DBParameters.Add("@RewardOptionID", SqlDbType.BigInt).Value = ROID
                                                    MyCommon.DBParameters.Add("@ProductGroupID", SqlDbType.Int).Value = ExcludedGroupID
                                                    MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
                                                    If (MyCommon.RowsAffected > 0) Then
                                                        MyCommon.Activity_Log(3, LogixID, 1, "Added condition for excluded product group " & ExcludedGroupID)
                                                    End If
                                                Else
                                                    ErrorCode = ERROR_CODES.ERROR_REMOVE_PRODUCT_FAILED
                                                    ErrorMsg = "Error encountered while attempting to create an exclusion group for offer " & ClientOfferID
                                                End If
                                            End If



                                            If (ExcludedGroupID > 0) Then
                                                MyCommon.QueryStr = "dbo.pa_ProdGroupItems_ManualInsert"
                                                MyCommon.Open_LRTsp()
                                                MyCommon.LRTsp.Parameters.Add("@ExtProductID", SqlDbType.NVarChar, 120).Value = ClientProductID
                                                MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt).Value = ExcludedGroupID
                                                MyCommon.LRTsp.Parameters.Add("@ProductTypeID", SqlDbType.Int).Value = 1 ' UPC
                                                MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 200).Value = ProductDesc
                                                MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
                                                MyCommon.LRTsp.Parameters.Add("@ProductStatus", SqlDbType.Int).Direction = ParameterDirection.Output
                                                MyCommon.LRTsp.ExecuteNonQuery()
                                                OutputStatus = MyCommon.LRTsp.Parameters("@Status").Value
                                                MyCommon.Close_LRTsp()

                                                MyCommon.Activity_Log(5, ExcludedGroupID, 1, Copient.PhraseLib.Lookup("history.pgroup-add", 1) & " " & ClientProductID)
                                            End If
                                        End If

                                        If ErrorCode = ERROR_CODES.ERROR_NONE Then
                                            MyCommon.QueryStr = "update productgroups with (RowLock) set updatelevel=updatelevel+1 where ProductGroupID = @ProductGroupID"
                                            MyCommon.DBParameters.Add("@ProductGroupID", SqlDbType.BigInt).Value = ProductGroupID
                                            MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
                                        End If
                                    Next
                                End If
                            Else
                                ErrorCode = ERROR_CODES.ERROR_PRODUCT_DOES_NOT_EXIST
                                ErrorMsg = "Client Product " & ClientProductID & " does not exist in Logix."
                            End If
                        Case Else
                            ErrorCode = ERROR_CODES.ERROR_OFFER_ENGINE_ID_INVALID
                            ErrorMsg = "Invalid Offer EngineID: " & OfferEngineId
                    End Select
                Else
                    ErrorCode = ERROR_CODES.ERROR_OFFER_DOES_NOT_EXIST
                    ErrorMsg = "Offer " & ClientOfferID & " does not exist in Logix."
                End If
            Else
                ErrorCode = ERROR_CODES.ERROR_XML_BAD_SOURCE_ID
                ErrorMsg = "Unrecognized ExternalSourceID: " & ExternalSourceID
            End If

        Catch ex As Exception
            ErrorCode = ERROR_CODES.ERROR_APPLICATION
            ErrorMsg = "Application Error encountered: " & ex.ToString
        Finally
            AppendToLog(ExternalSourceID, "RemoveProductFromOffer", "", "ClientOfferID=" & ClientOfferID & "; ClientProductID=" & ClientProductID, ErrorCode, ErrorMsg)
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try

        Return WriteOfferXmlResponse(LogixID, ClientOfferID, ErrorCode, ErrorMsg, RESPONSE_TYPES.REMOVE_PRODUCT_OFFER, (ErrorCode = ERROR_CODES.ERROR_NONE))

    End Function

    <WebMethod()>
    Public Function GetCustomerOffers(ByVal ExternalSourceID As String, ByVal CustomerID As String, ByVal IncludeAnyCardholders As Boolean) As String
        Dim dt, dtAssigned, dtAddOffers As DataTable
        Dim row As DataRow
        Dim sortedRows() As DataRow
        Dim CgXml As String = ""
        Dim rowCount As Integer
        Dim CustomerPK As Long = 0
        Dim reader As SqlDataReader = Nothing
        Dim ds As New DataSet()
        Dim ErrorCode As ERROR_CODES = ERROR_CODES.ERROR_NONE
        Dim ErrorMsg As String = ""
        Dim ResponseXml As New StringBuilder("<?xml version=""1.0"" encoding=""UTF-8""?>" & ControlChars.CrLf & "<ExternalOfferConnector>" & ControlChars.CrLf)
        Dim ErrorXml As New StringBuilder()
        Dim CustomerFound As Boolean = False
        Dim AutoDeploy As Boolean = False
        Dim DefaultCardType As Integer = 0
        Dim ErrorMsgLog As String = ""
        Dim ErrorMsgLogUnmasked As String = ""

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            ExtInterfaceID = GetExtInterfaceID(ExternalSourceID, AutoDeploy)

            If (ExtInterfaceID > 0) Then

                DefaultCardType = GetDefaultCardTypeID()
                CustomerID = MyCommon.Pad_ExtCardID(CustomerID, DefaultCardType)
                MyCommon.QueryStr = "select CustomerPK from CardIDs C with (NoLock) where ExtCardID = @ExtCardID and CardTypeID = @CardTypeID"
                MyCommon.DBParameters.Add("@ExtCardID", SqlDbType.NVarChar).Value = MyCryptLib.SQL_StringEncrypt(CustomerID, True)
                MyCommon.DBParameters.Add("@CardTypeID", SqlDbType.Int).Value = DefaultCardType
                dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)

                If (dt.Rows.Count > 0) Then
                    CustomerFound = True
                    CustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)

                    MyCommon.QueryStr = "select CustomerGroupID from groupmembership with (NoLock) where customerpk = @customerpk and deleted=0"
                    MyCommon.DBParameters.Add("@customerpk", SqlDbType.BigInt).Value = CustomerPK
                    dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)

                    CgXml = "<customergroups>"
                    If IncludeAnyCardholders Then CgXml &= "<id>1</id><id>2</id> "
                    rowCount = dt.Rows.Count
                    If rowCount > 0 Then
                        For Each row In dt.Rows
                            CgXml &= "<id>" & MyCommon.NZ(row.Item("CustomerGroupID"), "") & "</id>"
                        Next
                    End If
                    CgXml &= "</customergroups>"

                    MyCommon.QueryStr = "dbo.pa_CustomerOffersCurrentAndAdd"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@cgXml", SqlDbType.Xml).Value = CgXml
                    MyCommon.LRTsp.Parameters.Add("@IsEmployee", SqlDbType.Bit).Value = 0
                    MyCommon.LRTsp.Parameters.Add("@ShowAdd", SqlDbType.Bit).Value = 0
                    reader = MyCommon.LRTsp.ExecuteReader

                    dtAssigned = New DataTable
                    ds.Tables.Add(dtAssigned)
                    dtAddOffers = New DataTable
                    ds.Tables.Add(dtAddOffers)

                    ds.Load(reader, LoadOption.OverwriteChanges, Nothing, New DataTable() {dtAssigned, dtAddOffers})

                    MyCommon.Close_LRTsp()
                    reader.Close()

                Else
                    ErrorCode = ERROR_CODES.ERROR_INVALID_CUSTOMER_ID
                    ErrorMsg = "Customer: " & CustomerID & " (Card type ID: " & DefaultCardType & ") does not exist in Logix"
                    ErrorMsgLog = "Customer: " & Copient.MaskHelper.MaskCard(CustomerID, GetDefaultCardTypeID()) & " (Card type ID: " & DefaultCardType & ") does not exist in Logix"
                    ErrorMsgLogUnmasked = "Customer: " & CustomerID & " (Card type ID: " & DefaultCardType & ") does not exist in Logix"
                    Throw New Exception(ErrorMsg)
                End If

                ResponseXml.Append("<Customer id=""" & CustomerID & """ operation=""" & RESPONSE_TYPES.GET_CUSTOMER_OFFERS.ToString & """ success=""true"">" & ControlChars.CrLf)
                ResponseXml.Append("  <Offers>" & ControlChars.CrLf)
                If (dtAssigned IsNot Nothing AndAlso dtAssigned.Rows.Count > 0) Then
                    sortedRows = dtAssigned.Select("", "ExtOfferID")
                    For Each row In sortedRows
                        ResponseXml.Append("    <Offer id=""" & CleanXmlString(MyCommon.NZ(row.Item("ExtOfferID"), "")) & """ ")
                        ResponseXml.Append("logixId=""" & MyCommon.NZ(row.Item("OfferID"), "") & """ ")
                        ResponseXml.Append("name=""" & CleanXmlString(MyCommon.NZ(row.Item("Name"), "")) & """ ")
                        ResponseXml.Append("/>" & ControlChars.CrLf)
                    Next
                End If
                ResponseXml.Append("</Offers>" & ControlChars.CrLf)
                ResponseXml.Append("</Customer>" & ControlChars.CrLf)
            Else
                ErrorCode = ERROR_CODES.ERROR_XML_BAD_SOURCE_ID
                ErrorMsg = "Unrecognized ExternalSourceID: " & ExternalSourceID
                Throw New Exception(ErrorMsg)
            End If

        Catch ex As Exception
            If (ErrorCode = ERROR_CODES.ERROR_NONE) Then
                ErrorCode = ERROR_CODES.ERROR_APPLICATION
                ErrorMsg = "Application Error encountered: " & ex.ToString
            End If

            ' create error xml
            ErrorXml.Append("<?xml version=""1.0"" encoding=""UTF-8""?>" & ControlChars.CrLf)
            ErrorXml.Append("<ExternalOfferConnector>" & ControlChars.CrLf)
            ErrorXml.Append("  <Customer id=""" & CustomerID & """ operation=""" & RESPONSE_TYPES.GET_CUSTOMER_OFFERS.ToString & """ success=""false"" />" & ControlChars.CrLf)
            ErrorXml.Append("  <Error code=""" & ErrorCode.ToString & """ message=""" & ErrorMsg & """ />" & ControlChars.CrLf)
            ErrorXml.Append("</ExternalOfferConnector>")
        Finally
            If (Not String.IsNullOrWhiteSpace(ErrorMsgLog)) Then
                AppendToLog(ExternalSourceID, "GetCustomerOffers", "", "CustomerID=" & Copient.MaskHelper.MaskCard(CustomerID, GetDefaultCardTypeID()) & "; Includes AnyCardholder Offers=" & IncludeAnyCardholders, ErrorCode, ErrorMsgLog)
                AppendToLog(ExternalSourceID, "GetCustomerOffers", "", "CustomerID=" & CustomerID & "; Includes AnyCardholder Offers=" & IncludeAnyCardholders, ErrorCode, ErrorMsgLogUnmasked, True)
            Else
                AppendToLog(ExternalSourceID, "GetCustomerOffers", "", "CustomerID=" & Copient.MaskHelper.MaskCard(CustomerID, GetDefaultCardTypeID()) & "; Includes AnyCardholder Offers=" & IncludeAnyCardholders, ErrorCode, ErrorMsg)
                AppendToLog(ExternalSourceID, "GetCustomerOffers", "", "CustomerID=" & CustomerID & "; Includes AnyCardholder Offers=" & IncludeAnyCardholders, ErrorCode, ErrorMsgLogUnmasked, True)
            End If
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try

        ResponseXml.Append("</ExternalOfferConnector>")

        If (ErrorXml.Length > 0) Then
            Return ErrorXml.ToString
        Else
            Return ResponseXml.ToString
        End If

    End Function

    <WebMethod()>
    Public Function GetOfferCustomers(ByVal ExternalSourceID As String, ByVal ClientOfferID As String) As String
        Dim ErrorCode As ERROR_CODES = ERROR_CODES.ERROR_NONE
        Dim ErrorMsg As String = ""
        Dim ResponseXml As New StringBuilder("<?xml version=""1.0"" encoding=""UTF-8""?>" & ControlChars.CrLf & "<ExternalOfferConnector>" & ControlChars.CrLf)
        Dim ErrorXml As New StringBuilder()
        Dim CustomerGroupID As Integer
        Dim LogixID As Long
        Dim dt, dt2 As DataTable
        Dim row, row2 As DataRow
        Dim IsExcludedGroup As Boolean
        Dim AutoDeploy As Boolean = False
        Dim OfferEngineId As Long = -1

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
            ExtInterfaceID = GetExtInterfaceID(ExternalSourceID, AutoDeploy)

            If (ExtInterfaceID > 0) Then
                If DoesOfferExist(ClientOfferID, ExtInterfaceID, LogixID, OfferEngineId) Then
                    Select Case OfferEngineId
                        Case 0
                            MyCommon.QueryStr = "Select OC.LinkID as CustomerGroupID, 0 as ExcludedUsers, CG.AnyCardholder, CG.AnyCustomer, CG.NewCardholders " &
                                                "from OfferConditions OC with (NoLock) " &
                                                "inner join CustomerGroups CG with (NoLock) on CG.CustomerGroupID = OC.LinkID " &
                                                "where Oc.Deleted=0 And OC.ConditionTypeId=1 And CG.Deleted = 0 And OC.OfferID = @OfferID " &
                                                "union " &
                                                "Select OC.ExcludedID as CustomerGroupID, 1 as ExcludedUsers, CG.AnyCardholder, CG.AnyCustomer, CG.NewCardholders " &
                                                "from OfferConditions OC with (NoLock) " &
                                                "inner join CustomerGroups CG with (NoLock) on CG.CustomerGroupID = OC.ExcludedID " &
                                                "where Oc.Deleted=0 And OC.ConditionTypeId=1 And CG.Deleted=0 And OC.OfferID = @OfferID " &
                                                "order by AnyCardholder desc, AnyCustomer desc, NewCardholders desc, ExcludedUsers desc"
                            MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = LogixID
                        Case 2, 9
                            MyCommon.QueryStr = "Select ICG.CustomerGroupID, ICG.ExcludedUsers, CG.AnyCardholder, CG.AnyCustomer, CG.NewCardholders from CPE_IncentiveCustomerGroups ICG with (NoLock) " &
                                                "inner join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID = ICG.RewardOptionID " &
                                                "inner join CustomerGroups CG with (NoLock) on CG.CustomerGroupID = ICG.CustomerGroupID " &
                                                "where ICG.Deleted=0 and CG.Deleted=0 and RO.Deleted=0 and RO.IncentiveID = @IncentiveID " &
                                                "order by AnyCardholder desc, AnyCustomer desc, NewCardholders desc, ExcludedUsers desc"
                            MyCommon.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = LogixID
                        Case Else
                            ErrorCode = ERROR_CODES.ERROR_OFFER_ENGINE_ID_INVALID
                            ErrorMsg = "Invalid Offer EngineID: " & OfferEngineId
                    End Select
                    If ErrorMsg = "" Then
                        ResponseXml.Append("  <Offer id=""" & ClientOfferID & """ logixId=""" & LogixID & """ operation=""" & RESPONSE_TYPES.GET_OFFER_CUSTOMERS.ToString & """ success=""true"">" & ControlChars.CrLf)
                        ResponseXml.Append("    <Customers>" & ControlChars.CrLf)

                        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                        If (dt.Rows.Count > 0) Then
                            For Each row In dt.Rows
                                CustomerGroupID = MyCommon.NZ(row.Item("CustomerGroupID"), 0)
                                IsExcludedGroup = MyCommon.NZ(row.Item("ExcludedUsers"), False)

                                If (MyCommon.NZ(row.Item("AnyCardholder"), False)) Then
                                    ResponseXml.Append("      <Customer id=""0"" anycardholder=""true"" />" & ControlChars.CrLf)
                                ElseIf (MyCommon.NZ(row.Item("AnyCustomer"), False)) Then
                                    ResponseXml.Append("      <Customer id=""-1"" anycustomer=""true"" />" & ControlChars.CrLf)
                                ElseIf (MyCommon.NZ(row.Item("NewCardholders"), False)) Then
                                    ResponseXml.Append("      <Customer id==""-2"" newcardholders=""true"" />" & ControlChars.CrLf)
                                Else
                                    MyCommon.QueryStr = "select ExtCardIDOriginal as ExtCardID from CardIDs with (NoLock) where CustomerPK in " &
                                                        "(select CustomerPK from GroupMembership with (NoLock) where CustomerGroupID = @CustomerGroupID and Deleted=0)"
                                    MyCommon.DBParameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = CustomerGroupID
                                    dt2 = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)
                                    For Each row2 In dt2.Rows
                                        ResponseXml.Append("      <Customer id=""" & MyCommon.NZ(MyCryptLib.SQL_StringDecrypt(row2.Item("ExtCardID").ToString()), "0") & """ ")
                                        If IsExcludedGroup Then ResponseXml.Append(" excluded=""true"" ")
                                        ResponseXml.Append("/>" & ControlChars.CrLf)
                                    Next
                                End If
                            Next
                        End If
                        ResponseXml.Append("    </Customers>" & ControlChars.CrLf)
                        ResponseXml.Append("  </Offer>" & ControlChars.CrLf)
                    End If
                Else
                    ErrorCode = ERROR_CODES.ERROR_OFFER_DOES_NOT_EXIST
                    ErrorMsg = "Offer " & ClientOfferID & " does not exist in Logix."
                    Throw New Exception(ErrorMsg)
                End If
            Else
                ErrorCode = ERROR_CODES.ERROR_XML_BAD_SOURCE_ID
                ErrorMsg = "Unrecognized ExternalSourceID: " & ExternalSourceID
                Throw New Exception(ErrorMsg)
            End If
        Catch ex As Exception
            If (ErrorCode = ERROR_CODES.ERROR_NONE) Then
                ErrorCode = ERROR_CODES.ERROR_APPLICATION
                ErrorMsg = "Application Error encountered: " & ex.ToString
            End If

            ' create error xml
            ErrorXml.Append("<?xml version=""1.0"" encoding=""UTF-8""?>" & ControlChars.CrLf)
            ErrorXml.Append("<ExternalOfferConnector>" & ControlChars.CrLf)
            ErrorXml.Append("  <Offer id=""" & ClientOfferID & """ logixId=""" & LogixID & """ operation=""" & RESPONSE_TYPES.GET_OFFER_CUSTOMERS.ToString & """ success=""false"" />" & ControlChars.CrLf)
            ErrorXml.Append("  <Error code=""" & ErrorCode.ToString & """ message=""" & ErrorMsg & """ />" & ControlChars.CrLf)
            ErrorXml.Append("</ExternalOfferConnector>")
        Finally
            AppendToLog(ExternalSourceID, "GetOfferCustomers", "", "ClientOfferID=" & ClientOfferID, ErrorCode, ErrorMsg)
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try

        ResponseXml.Append("</ExternalOfferConnector>")

        If (ErrorXml.Length > 0) Then
            Return ErrorXml.ToString
        Else
            Return ResponseXml.ToString
        End If

    End Function
    Private Sub ChangeFormatToDateTime(ByVal offerXmlDoc As XmlDocument)
        Dim tempNode As XmlNode
        Dim offerNode As XmlNode = offerXmlDoc.SelectSingleNode("Offer")
        Dim dateTimeFormat As String = "yyyy-MM-dd'T'HH:mm:ss"

        tempNode = offerNode.SelectSingleNode("StartDate")
        If tempNode IsNot Nothing Then tempNode.InnerText = Date.Parse(tempNode.InnerText).ToString(dateTimeFormat)

        tempNode = offerNode.SelectSingleNode("EndDate")
        If tempNode IsNot Nothing Then tempNode.InnerText = Date.Parse(tempNode.InnerText).ToString(dateTimeFormat)

        tempNode = offerNode.SelectSingleNode("EligibilityStartDate")
        If tempNode IsNot Nothing Then tempNode.InnerText = Date.Parse(tempNode.InnerText).ToString(dateTimeFormat)

        tempNode = offerNode.SelectSingleNode("EligibilityEndDate")
        If tempNode IsNot Nothing Then tempNode.InnerText = Date.Parse(tempNode.InnerText).ToString(dateTimeFormat)

        tempNode = offerNode.SelectSingleNode("TestingStartDate")
        If tempNode IsNot Nothing Then tempNode.InnerText = Date.Parse(tempNode.InnerText).ToString(dateTimeFormat)

        tempNode = offerNode.SelectSingleNode("TestingEndDate")
        If tempNode IsNot Nothing Then tempNode.InnerText = Date.Parse(tempNode.InnerText).ToString(dateTimeFormat)
    End Sub

    Dim XMLValidationErrMsg As New System.Text.StringBuilder("")
    Dim XMLValidationWarningMsg As New System.Text.StringBuilder("")
    Public Sub OnValidationCallBack(ByVal sender As Object, ByVal args As ValidationEventArgs)
        If args.Severity = XmlSeverityType.Warning Then
            XMLValidationWarningMsg.AppendLine(vbTab & "Warning: Matching schema not found.  No validation occurred." & Convert.ToString(args.Message))
        Else
            XMLValidationErrMsg.AppendLine(vbTab & "Validation error: " & Convert.ToString(args.Message))
        End If
    End Sub
    Private Function IsValidDocument(ByVal OfferType As String,
                                     ByRef ErrorCode As ERROR_CODES, ByRef ErrorMsg As String, ByVal EngineID As Integer, ByVal OfferXmlDoc As XmlDocument) As Boolean
        Dim XsdFileName As String = ""
        Dim Settings As XmlReaderSettings
        Dim xr As XmlReader = Nothing
        Dim bValid As Boolean = True
        '  Dim OfferXmlDoc As New XmlDocument
        Dim isTimeEnabledForUEOffers As Boolean = (EngineID = Engines.UE) AndAlso (MyCommon.Fetch_UE_SystemOption(200) = "1")

        Try
            ' only CPE validates against a specific XSD 
            If (OfferXmlDoc IsNot Nothing AndAlso OfferXmlDoc.SelectSingleNode("Offer") IsNot Nothing) Then
                XsdFileName = LoadXsdFileName(OfferType, ErrorCode, ErrorMsg, isTimeEnabledForUEOffers)
                If isTimeEnabledForUEOffers Then
                    ChangeFormatToDateTime(OfferXmlDoc)
                End If

                If (ErrorMsg = "") Then

                    Settings = New XmlReaderSettings()
                    Settings.Schemas.Add(Nothing, XsdFileName)
                    Settings.ValidationType = ValidationType.Schema
                    Settings.IgnoreComments = True
                    Settings.IgnoreProcessingInstructions = True
                    Settings.IgnoreWhitespace = True
                    OfferXmlDoc.Schemas = Settings.Schemas
                    OfferXmlDoc.Validate(AddressOf OnValidationCallBack)
                    If Not XMLValidationErrMsg.ToString() = "" Then
                        ErrorCode = ERROR_CODES.ERROR_XML_INVALID_DOC
                        ErrorMsg = XMLValidationErrMsg.ToString()
                        bValid = False
                    Else
                        If Not XMLValidationWarningMsg.ToString() = "" Then
                            Copient.Logger.Write_Log(AcceptanceLogFile, XMLValidationWarningMsg.ToString(), True)
                        End If
                        bValid = True
                    End If

                Else
                    bValid = False
                End If
            Else
                bValid = False
                ErrorCode = ERROR_CODES.ERROR_XML_EMPTY_DOC
                ErrorMsg = "XML document is empty"
            End If

        Catch eXmlSch As XmlSchemaException
            ErrorCode = ERROR_CODES.ERROR_XML_INVALID_DOC
            ErrorMsg = "(Xml Schema Validation Error Line: " & eXmlSch.LineNumber.ToString & " - Col: " & eXmlSch.LinePosition.ToString & ") " & eXmlSch.Message
            bValid = False
        Catch eXml As XmlException
            ErrorCode = ERROR_CODES.ERROR_XML_INVALID_DOC
            ErrorMsg = "(Xml Error Line: " & eXml.LineNumber.ToString & " - Col: " & eXml.LinePosition.ToString & ") " & eXml.Message
            bValid = False
        Catch exApp As ApplicationException
            ErrorCode = ERROR_CODES.ERROR_APPLICATION
            ErrorMsg = "Application Error: " & exApp.ToString
            bValid = False
        Catch ex As Exception
            ErrorCode = ERROR_CODES.ERROR_APPLICATION
            ErrorMsg = "Error: " & ex.ToString
            bValid = False
        Finally
            If Not xr Is Nothing Then
                xr.Close()
            End If
        End Try

        ' AMSPS-3044: Rather than remove the following lines, I am commenting them out.  I need to leave them in (commented), as a 
        ' warning / reminder to anyone who might be tempted to put something like this back into the code.  Even though the 3 tags that
        ' were checked here when CM is the engine, are in fact not used by CM, they are part of the schema, and customers have built
        ' offer templates based on that schema.  That is, their offers began to be rejected by the EOC as part of AMS 6.2 -- it is 
        ' not backward compatible.  These tags, while not used by CM, must not throw an error by the EOC!
        '        If EngineID = Engines.CM Then
        '            bValid = ValidateTagsForCM(OfferXmlDoc, ErrorMsg)
        '        End If

        Return bValid
    End Function

    Private Function GetEngineID(ByRef sEngineType As String) As Integer
        Dim EngineID As Integer = -1
        Dim dt As DataTable

        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
        If sEngineType = "" Then
            MyCommon.QueryStr = "select EngineID, Description from PromoEngines with (NoLock) where DefaultEngine=1 and Installed=1"
        Else
            MyCommon.QueryStr = "select EngineID, Description from PromoEngines with (NoLock) where Description = @Description and Installed=1"
            MyCommon.DBParameters.Add("@Description", SqlDbType.NVarChar).Value = sEngineType.ConvertBlankIfNothing
        End If
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        If (dt.Rows.Count > 0) Then
            EngineID = MyCommon.NZ(dt.Rows(0).Item("EngineID"), -1)
            If sEngineType = "" Then
                sEngineType = MyCommon.NZ(dt.Rows(0).Item("Description"), "No engines installed")
            End If
        End If

        Return EngineID
    End Function

    Private Function GetExtInterfaceID(ByVal ExternalSourceID As String, ByRef AutoDeploy As Boolean, Optional ByRef DefaultAsLogixID As Boolean = False) As Integer
        Dim dt As DataTable

        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

        MyCommon.QueryStr = "select ExtInterfaceID, AutoDeploy, DefaultAsLogixID from ExtCRMInterfaces with (NoLock) where ExtCode = @ExtCode and Deleted=0"
        MyCommon.DBParameters.Add("@ExtCode", SqlDbType.NVarChar).Value = ExternalSourceID.ConvertBlankIfNothing
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        If (dt.Rows.Count > 0) Then
            ExtInterfaceID = MyCommon.NZ(dt.Rows(0).Item("ExtInterfaceID"), -1)
            AutoDeploy = MyCommon.NZ(dt.Rows(0).Item("AutoDeploy"), False)
            DefaultAsLogixID = MyCommon.NZ(dt.Rows(0).Item("DefaultAsLogixID"), False)
        End If

        Return ExtInterfaceID
    End Function

    Private Function GetDefaultReceiptMessage(ByVal ExternalSourceID As String, ByVal IsMfg As Boolean) As String
        Dim dt As DataTable
        Dim sMessage As String = ""

        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

        MyCommon.QueryStr = "select MfgDefaultReceiptMessage, NonMfgDefaultReceiptMessage from ExtCRMInterfaces with (NoLock) where ExtCode = @ExtCode and Deleted=0"
        MyCommon.DBParameters.Add("@ExtCode", SqlDbType.NVarChar).Value = ExternalSourceID.ConvertBlankIfNothing
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        If (dt.Rows.Count > 0) Then
            If IsMfg Then
                sMessage = MyCommon.NZ(dt.Rows(0).Item("MfgDefaultReceiptMessage"), "")
            Else
                sMessage = MyCommon.NZ(dt.Rows(0).Item("NonMfgDefaultReceiptMessage"), "")
            End If
        End If

        Return sMessage
    End Function

    Private Function GetCRMEngineID(ByVal CRMEngineExtCode As String) As Integer
        Dim dt As DataTable
        Dim CRMEngineID As Integer = -1

        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

        MyCommon.QueryStr = "select ExtInterfaceID from ExtCRMInterfaces with (NoLock) where ExtCode = @ExtCode"
        MyCommon.DBParameters.Add("@ExtCode", SqlDbType.NVarChar).Value = CRMEngineExtCode.ConvertBlankIfNothing
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        If (dt.Rows.Count > 0) Then
            CRMEngineID = MyCommon.NZ(dt.Rows(0).Item("ExtInterfaceID"), -1)
        End If

        Return CRMEngineID
    End Function

    Private Function ShouldSendIssuance(ByVal ExtInterfaceID As Integer) As Boolean
        Dim dt As DataTable
        Dim SendIssuance As Boolean = False

        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

        MyCommon.QueryStr = "select EnableIssuance from ExtCRMInterfaces with (NoLock) where ExtInterfaceID = @ExtInterfaceID and Deleted=0"
        MyCommon.DBParameters.Add("@ExtInterfaceID", SqlDbType.Int).Value = ExtInterfaceID
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        If (dt.Rows.Count > 0) Then
            SendIssuance = MyCommon.NZ(dt.Rows(0).Item("EnableIssuance"), False)
        End If

        Return SendIssuance
    End Function

    Private Function LoadXsdFileName(ByVal EngineID As Integer, ByRef ErrorCode As ERROR_CODES, ByRef ErrorMsg As String) As String
        Dim xsdFileName As String = ""

        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

        ' get the appropriate schema based on the external source ID
        Select Case EngineID
            Case 2, 9
                xsdFileName = MyCommon.Get_Install_Path & "AgentFiles\" & "PromoCPE.xsd"
            Case Else
                xsdFileName = ""
                ErrorCode = ERROR_CODES.ERROR_XML_BAD_SOURCE_ID
                ErrorMsg = "Unrecognized Engine Type"
        End Select

        ' ensure that the schema exists on the server
        If xsdFileName <> "" And Not System.IO.File.Exists(xsdFileName) Then
            ErrorCode = ERROR_CODES.ERROR_XSD_NOT_FOUND
            ErrorMsg = "XSD file not found: " & xsdFileName
            xsdFileName = ""
        End If

        Return xsdFileName
    End Function

    Private Function TryParseElementValue(ByVal OfferXmlDoc As XmlDocument, ByVal XPath As String, ByRef ElementValue As String) As Boolean
        Dim SingleNode As XmlNode = Nothing
        Dim IsParsed As Boolean = False

        SingleNode = OfferXmlDoc.SelectSingleNode(XPath)
        If (SingleNode IsNot Nothing) Then
            ElementValue = SingleNode.InnerText
            IsParsed = True
        End If

        Return IsParsed
    End Function
    Private Function TryParseAttributeValue(ByVal OfferXmlDoc As XmlDocument, ByVal ElementName As String,
                                            ByVal AttributeName As String, ByRef AttributeValue As String) As Boolean
        Dim SingleNode As XmlNode = Nothing
        Dim Attrib As XmlAttribute = Nothing
        Dim IsParsed As Boolean = False

        SingleNode = OfferXmlDoc.SelectSingleNode(ElementName)
        If (SingleNode IsNot Nothing) Then
            Attrib = SingleNode.Attributes(AttributeName)
            If Attrib IsNot Nothing Then
                AttributeValue = Attrib.InnerText
                IsParsed = True
            End If
        End If

        Return IsParsed
    End Function
    ' AMSPS-3044: Rather than remove the following lines, I am commenting them out.  I need to leave them in (commented), as a 
    ' warning / reminder to anyone who might be tempted to put something like this back into the code.  Even though the 3 tags that
    ' were checked here when CM is the engine, are in fact not used by CM, they are part of the schema, and customers have built
    ' offer templates based on that schema.  That is, their offers began to be rejected by the EOC as part of AMS 6.2 -- it is 
    ' not backward compatible.  These tags, while not used by CM, must not throw an error by the EOC!
    '    Private Function ValidTagExists(ByVal OfferXmlDoc As XmlDocument, ByVal XPath As String) As Boolean
    '        Dim Valid As Boolean = True
    '        Dim SingleNode As XmlNode = Nothing

    '        SingleNode = OfferXmlDoc.SelectSingleNode(XPath)
    '        If (SingleNode IsNot Nothing) Then
    '            Valid = False
    '        End If

    '        Return Valid
    '    End Function
    '    Private Function ValidateTagsForCM(ByVal OfferXmlDoc As XmlDocument, ByRef ErrorMsg As String) As Boolean
    '        Dim Valid As Boolean = True
    '        Dim SingleNode As XmlNode = Nothing
    '        Dim ReportImpPath As String = "//Offer/ReportImpressions"
    '        Dim ReportRedemPath As String = "//Offer/ReportRedemptions"
    '        Dim PromoClassIDPath As String = "//Offer/PromoClassID"

    '        Valid = ValidTagExists(OfferXmlDoc, ReportImpPath)
    '        If Not Valid Then
    '            ErrorMsg = "ReportImpressions tag is not valid for CM Engine"
    '        Else
    '            Valid = ValidTagExists(OfferXmlDoc, ReportRedemPath)
    '            If Not Valid Then
    '               ErrorMsg = "ReportRedemptions tag is not valid for CM Engine"
    '            Else
    '                Valid = ValidTagExists(OfferXmlDoc, PromoClassIDPath)
    '                If Not Valid Then
    '                    ErrorMsg = "PromoClassID tag is not valid for CM Engine"
    '                End If
    '            End If
    '        End If


    '        Return Valid
    '    End Function
    Private Function LoadXsdFileName(ByVal OfferType As String, ByRef ErrorCode As ERROR_CODES, ByRef ErrorMsg As String, ByVal isTimeEnableForUEOffers As Boolean) As String
        Dim xsdFileName As String = ""

        If (OfferType IsNot Nothing AndAlso OfferType <> "") Then
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

            Select Case OfferType
                Case "CENTS_OFF", "PERCENT_OFF", "PRICE_POINT_ITEMS", "CENTS_OFF_WEIGHT_VOLUME", "PRICE_POINT_WEIGHT_VOLUME"
                    xsdFileName = MyCommon.Get_Install_Path & "AgentFiles\ExternalOffers\" & IIf(isTimeEnableForUEOffers, "DiscountedUE.xsd", "Discounted.xsd")
                Case "FREE"
                    xsdFileName = MyCommon.Get_Install_Path & "AgentFiles\ExternalOffers\" & IIf(isTimeEnableForUEOffers, "FreeItemUE.xsd", "FreeItem.xsd")
                Case Else
                    ErrorCode = ERROR_CODES.ERROR_XSD_NOT_FOUND
                    ErrorMsg = "XSD not found that matches offer type of " & OfferType
                    xsdFileName = ""
            End Select

            ' ensure that the schema exists on the server
            If xsdFileName <> "" And Not System.IO.File.Exists(xsdFileName) Then
                ErrorCode = ERROR_CODES.ERROR_XSD_NOT_FOUND
                ErrorMsg = "XSD file not found: " & xsdFileName
                xsdFileName = ""
            End If

        End If

        Return xsdFileName
    End Function

    Private Function IsManufacturerCouponSet(ByVal OfferXmlDoc As XmlDocument, ByRef MfgCouponValue As Boolean) As Boolean
        Dim MfgSet As Boolean = False
        Dim mfgOfferNode As XmlNode

        mfgOfferNode = OfferXmlDoc.SelectSingleNode("//Offer/IsManufacturerCoupon")
        If (mfgOfferNode IsNot Nothing) Then
            MfgSet = Boolean.TryParse(mfgOfferNode.InnerText, MfgCouponValue)
        End If

        Return MfgSet
    End Function

    Private Function IsStoreCouponSet(ByVal OfferXmlDoc As XmlDocument, ByRef StoreCouponValue As Boolean) As Boolean
        Dim StoreCouponSet As Boolean = False
        Dim storeCouponNode As XmlNode

        storeCouponNode = OfferXmlDoc.SelectSingleNode("//Offer/IsStoreCoupon")
        If (storeCouponNode IsNot Nothing) Then
            StoreCouponSet = Boolean.TryParse(storeCouponNode.InnerText, StoreCouponValue)
        End If

        Return StoreCouponSet
    End Function

    Private Function GetCreateOfferFields(ByVal OfferXmlDoc As XmlDocument, ByVal ExtInterfaceID As Long, ByVal EngineID As Integer) As NameValueCollection
        Dim OfferFields As New NameValueCollection(7)
        Dim Name As String = ""
        Dim ClientOfferID As String = ""
        Dim BannerIDs(-1) As Long
        Dim i As Integer
        Dim CRMEngineExtCode As String = ""
        Dim AutoSendOutbound As Boolean = False
        Dim SendOutbound As Boolean = False

        If (OfferXmlDoc IsNot Nothing) Then
            TryParseElementValue(OfferXmlDoc, "//Offer/Name", Name)
            OfferFields.Add("Name", Name)
            OfferFields.Add("crmEngineID", ExtInterfaceID.ToString)
            OfferFields.Add("EngineID", EngineID.ToString())
            TryParseElementValue(OfferXmlDoc, "//Offer/ClientOfferID", ClientOfferID)
            OfferFields.Add("ClientOfferID", ClientOfferID)
            If ((EngineID = 0) And (MyCommon.Fetch_CM_SystemOption(114) = "1")) Then
                TryParseElementValue(OfferXmlDoc, "//Offer/CRMEngineExtCode", CRMEngineExtCode)
                OfferFields.Add("CRMEngineExtCode", CRMEngineExtCode)
                TryParseElementValue(OfferXmlDoc, "//Offer/Actions/SendOutbound", SendOutbound)
                OfferFields.Add("SendOutbound", SendOutbound)
            End If

            BannerIDs = GetBannerIDs(OfferXmlDoc)
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            For i = 0 To BannerIDs.GetUpperBound(0)
                If (BannerIDs(i) > 0) Then
                    OfferFields.Add("banner", BannerIDs(i).ToString)
                End If
            Next
        End If

        Return OfferFields
    End Function

    Private Function GetSaveOfferFields(ByVal OfferXmlDoc As XmlDocument, ByVal OfferID As Long, ByVal ExtInterfaceID As Long, ByVal EngineID As Integer) As NameValueCollection
        Dim OfferFields As New NameValueCollection(18)
        Dim CustomLimitNode As XmlNode = Nothing
        Dim TempVal As String = ""
        Dim TempDate As Date
        Dim BannerIDs(-1) As Long
        Dim i As Integer
        Dim Limit As String = ""
        Dim Period As String = ""
        Dim PeriodType As String = ""
        Dim Frequency As String = ""
        Dim MfgCouponValue As String = ""
        Dim IssuanceValue As String = ""
        Dim MyCpeOffer As New Copient.CPEOffer
        Dim VendorID As Integer = 0
        Dim VendorName As String = ""
        Dim Priority As Integer = 1
        Dim StoreCouponValue As String = ""

        If (OfferXmlDoc IsNot Nothing) Then
            OfferFields.Add("OfferID", OfferID.ToString)
            If TryParseElementValue(OfferXmlDoc, "//Offer/StartDate", TempVal) Then
                If Date.TryParse(TempVal, TempDate) Then
                    OfferFields.Add("productionstart", TempDate.ToString("yyyy-MM-ddTHH:mm:ss"))
                End If
            End If
            If TryParseElementValue(OfferXmlDoc, "//Offer/EndDate", TempVal) Then
                If Date.TryParse(TempVal, TempDate) Then
                    OfferFields.Add("productionend", TempDate.ToString("yyyy-MM-ddTHH:mm:ss"))
                End If
            End If
            If ((MyCommon.Fetch_CM_SystemOption(85) = "1" And EngineID = 0) Or (MyCommon.Fetch_UE_SystemOption(143) = "1" And EngineID = 9)) Then
                If TryParseElementValue(OfferXmlDoc, "//Offer/DisplayStartDate", TempVal) Then
                    If Date.TryParse(TempVal, TempDate) Then
                        OfferFields.Add("displaystartdate", TempDate.ToString("yyyy-MM-ddTHH:mm:ss"))
                    End If
                End If
                If TryParseElementValue(OfferXmlDoc, "//Offer/DisplayEndDate", TempVal) Then
                    If Date.TryParse(TempVal, TempDate) Then
                        OfferFields.Add("displayenddate", TempDate.ToString("yyyy-MM-ddTHH:mm:ss"))
                    End If
                End If
            End If
            If TryParseElementValue(OfferXmlDoc, "//Offer/EligibilityStartDate", TempVal) Then
                If Date.TryParse(TempVal, TempDate) Then
                    OfferFields.Add("eligibilitystart", TempDate.ToString("yyyy-MM-ddTHH:mm:ss"))
                End If
            End If
            If TryParseElementValue(OfferXmlDoc, "//Offer/EligibilityEndDate", TempVal) Then
                If Date.TryParse(TempVal, TempDate) Then
                    OfferFields.Add("eligibilityend", TempDate.ToString("yyyy-MM-ddTHH:mm:ss"))
                End If
            End If

            If TryParseElementValue(OfferXmlDoc, "//Offer/TestingStartDate", TempVal) Then
                If Date.TryParse(TempVal, TempDate) Then
                    OfferFields.Add("testingstart", TempDate.ToString("yyyy-MM-ddTHH:mm:ss"))
                End If
            End If
            If TryParseElementValue(OfferXmlDoc, "//Offer/TestingEndDate", TempVal) Then
                If Date.TryParse(TempVal, TempDate) Then
                    OfferFields.Add("testingend", TempDate.ToString("yyyy-MM-ddTHH:mm:ss"))
                End If
            End If

            If TryParseElementValue(OfferXmlDoc, "//Offer/Name", TempVal) Then OfferFields.Add("form_name", TempVal)
            If TryParseElementValue(OfferXmlDoc, "//Offer/Description", TempVal) Then OfferFields.Add("form_description", TempVal)
            If TryParseElementValue(OfferXmlDoc, "//Offer/IsManufacturerCoupon", TempVal) Then
                If TempVal.ToUpper = "TRUE" OrElse TempVal = "1" Then MfgCouponValue = "on"
                If TempVal.ToUpper = "FALSE" OrElse TempVal = "0" Then MfgCouponValue = "off"
                OfferFields.Add("mfgCoupon", MfgCouponValue)
            End If

            If TryParseElementValue(OfferXmlDoc, "//Offer/ChargebackVendorCode", TempVal) Then
                TryParseAttributeValue(OfferXmlDoc, "//Offer/ChargebackVendorCode", "name", VendorName)
                VendorID = MyCpeOffer.GetVendorID(TempVal, VendorName)
                If VendorID > 0 AndAlso IsChargeableVendor(VendorID) Then
                    OfferFields.Add("vendor", VendorID)
                End If
            End If

            If TryParseElementValue(OfferXmlDoc, "//Offer/VendorCouponCode", TempVal) Then OfferFields.Add("vendorCouponCode", TempVal)

            If TryParseElementValue(OfferXmlDoc, "//Offer/ReportImpressions", TempVal) Then
                If TempVal.ToUpper = "TRUE" OrElse TempVal = "1" Then TempVal = "on"
                OfferFields.Add("reportingimp", TempVal)
            Else
                ' use the default system option value
                OfferFields.Add("reportingimp", IIf(MyCommon.Fetch_CPE_SystemOption(84) = "1", "on", ""))
            End If

            If TryParseElementValue(OfferXmlDoc, "//Offer/ReportRedemptions", TempVal) Then
                If TempVal.ToUpper = "TRUE" OrElse TempVal = "1" Then TempVal = "on"
                OfferFields.Add("reportingred", TempVal)
            Else
                ' use the default system option value
                OfferFields.Add("reportingred", IIf(MyCommon.Fetch_CPE_SystemOption(85) = "1", "on", ""))
            End If

            If TryParseElementValue(OfferXmlDoc, "//Offer/PromoClassID", TempVal) Then OfferFields.Add("form_Category", TempVal)

            If TryParseElementValue(OfferXmlDoc, "//Offer/SendIssuance", TempVal) Then
                If TempVal.ToUpper = "TRUE" OrElse TempVal = "1" Then IssuanceValue = "on"
                OfferFields.Add("issuance", IssuanceValue)
            Else
                ' get the value from the external source's enable issuance default value.
                OfferFields.Add("issuance", IIf(ShouldSendIssuance(ExtInterfaceID), "on", ""))
            End If

            If TryParseElementValue(OfferXmlDoc, "//Offer/CRMEngineExtCode", TempVal) Then
                TempVal = GetCRMEngineID(TempVal)
                OfferFields.Add("crmEngineID", TempVal)
            End If

            ' save the offer's limit
            If TryParseElementValue(OfferXmlDoc, "//Offer/Limits/Frequency", Frequency) Then
                Select Case Frequency
                    Case "NO_LIMIT"
                        OfferFields.Add("limit3", "0")
                        OfferFields.Add("limit3period", "0")
                        OfferFields.Add("P3DistTimeType", "2")
                    Case "ONCE_PER_TRANSACTION"
                        OfferFields.Add("limit3", "1")
                        OfferFields.Add("limit3period", "1")
                        OfferFields.Add("P3DistTimeType", "2")
                    Case "ONCE_PER_DAY"
                        OfferFields.Add("limit3", "1")
                        OfferFields.Add("limit3period", "1")
                        OfferFields.Add("P3DistTimeType", "1")
                    Case "ONCE_PER_WEEK"
                        OfferFields.Add("limit3", "1")
                        OfferFields.Add("limit3period", "7")
                        OfferFields.Add("P3DistTimeType", "1")
                    Case Else
                        ' treat all others as once per offer
                        OfferFields.Add("limit3", "1")
                        If EngineID = 0 Then
                            OfferFields.Add("limit3period", "-1")
                        Else
                            OfferFields.Add("limit3period", "3650")
                        End If
                        OfferFields.Add("P3DistTimeType", "1")
                End Select
            Else
                CustomLimitNode = OfferXmlDoc.SelectSingleNode("//Offer/Limits/CustomLimit")
                If CustomLimitNode IsNot Nothing Then
                    TryParseElementValue(OfferXmlDoc, "//Offer/Limits/CustomLimit/Limit", Limit)
                    OfferFields.Add("limit3", Limit)
                    TryParseElementValue(OfferXmlDoc, "//Offer/Limits/CustomLimit/Period", Period)
                    OfferFields.Add("limit3period", Period)
                    TryParseElementValue(OfferXmlDoc, "//Offer/Limits/CustomLimit/PeriodType", PeriodType)
                    Select Case PeriodType
                        Case "DAYS_SINCE_INCENTIVE_START", "DAYS_ROLLING"
                            PeriodType = "1"
                        Case "HOURS_SINCE_LAST_AWARDED", "PER_TRANSACTION"
                            PeriodType = "2"
                        Case Else ' default to hours (per_transaction for UE)
                            PeriodType = "2"
                    End Select
                    OfferFields.Add("P3DistTimeType", PeriodType)
                End If
            End If

            BannerIDs = GetBannerIDs(OfferXmlDoc)
            If BannerIDs.Length > 0 Then
                OfferFields.Add("bannerschanged", "true")
                For i = 0 To BannerIDs.GetUpperBound(0)
                    If (BannerIDs(i) > 0) Then
                        OfferFields.Add("banner", BannerIDs(i).ToString)
                        OfferFields.Add("bannerids", BannerIDs(i).ToString)
                    End If
                Next
            End If
        End If

        If EngineID = Copient.CommonInc.InstalledEngines.UE Then
            If Integer.TryParse(MyCommon.Fetch_InterfaceOption(56), Priority) Then
                OfferFields.Add("priority", Priority.ToString)
            Else
                If (MyCommon.IsEngineInstalled(9) AndAlso (MyCommon.Fetch_UE_SystemOption(146) = "1")) Then
                    OfferFields.Add("priority", "50")
                Else
                    OfferFields.Add("priority", "1")
                End If
            End If
            If TryParseElementValue(OfferXmlDoc, "//Offer/EnableCollisionDetection", TempVal) AndAlso Boolean.TryParse(TempVal, Nothing) Then
                OfferFields.Add("enablecollisiondetection", TempVal)
            End If

        End If

        If TryParseElementValue(OfferXmlDoc, "//Offer/IsStoreCoupon", TempVal) Then
            If TempVal.ToUpper = "TRUE" OrElse TempVal = "1" Then StoreCouponValue = "on"
            If TempVal.ToUpper = "FALSE" OrElse TempVal = "0" Then StoreCouponValue = "off"
            OfferFields.Add("storeCoupon", StoreCouponValue)
        End If

        If TryParseElementValue(OfferXmlDoc, "//Offer/CurrencyID", TempVal) Then
            OfferFields.Add("currencyId", TempVal)
        End If

        Return OfferFields
    End Function
    Private Sub ProcessTrackableCouponConditions(ByVal OfferXmlDoc As XmlDocument, ByVal OfferID As Long, ByVal EngineID As Integer, ByRef ErrorMessage As String)
        Dim ExtProgramIDNodes, TCListNodes As XmlNodeList
        Dim MyCpeOffer As New Copient.CPEOffer
        Dim ExtProgramID As String = ""
        Dim AddedProgramID As Boolean = False
        Dim AddedCoupons As Boolean = False
        Dim dtTC As DataTable = Nothing
        Dim TrackableCoupon As XmlNode
        Dim ErrorCode As ERROR_CODES = ERROR_CODES.ERROR_NONE
        Dim AcceptanceLogFile As String = "EOCAcceptanceLog." & Date.Now.ToString("yyyyMMdd") & ".txt"
        Dim status As Boolean
        Dim dtDup As DataTable = Nothing
        CurrentRequest.Resolver.AppName = "External Offer Connector"
        Dim offerService As IOffer = CurrentRequest.Resolver.Resolve(Of OfferService)()
        Try
            TrackableCoupon = OfferXmlDoc.SelectSingleNode("//Offer/Conditions/TrackableCoupon")
            TrackableCoupon = OfferXmlDoc.SelectSingleNode("//Offer/Conditions/TrackableCoupon")
            If TrackableCoupon IsNot Nothing AndAlso TrackableCoupon.ChildNodes.Count > 0 Then
                If (TrackableCoupon.Attributes("MaxRedeemCount") IsNot Nothing AndAlso TrackableCoupon.Attributes("MaxRedeemCount").Value IsNot Nothing AndAlso TrackableCoupon.Attributes("MaxRedeemCount").Value > 0) Then
                    status = TrackableCoupon.Attributes("MaxRedeemCount").Value
                End If
                If EngineID = InstalledEngines.UE Then
                    ExtProgramIDNodes = OfferXmlDoc.SelectNodes("//Offer/Conditions/TrackableCoupon/ExtProgramID")
                    TCListNodes = OfferXmlDoc.SelectNodes("//Offer/Conditions/TrackableCoupon/TCList")
                    If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
                    If (ExtProgramIDNodes IsNot Nothing AndAlso ExtProgramIDNodes.Count > 0) Then
                        If (ExtProgramIDNodes(0).InnerText IsNot "") Then
                            ExtProgramID = ExtProgramIDNodes(0).InnerText
                        End If
                        Dim programId As Integer = 0
                        Dim maxRedeemCount As Integer = 0
                        Dim extProgramIDExists As Boolean = MyCpeOffer.CheckExtProgramIDExists(ExtProgramID, programId, maxRedeemCount, ErrorMessage) 'check if the ExtProgramID exists or not
                        If (Not extProgramIDExists AndAlso maxRedeemCount <= 0) Then
                            If (status) Then
                                maxRedeemCount = TrackableCoupon.Attributes("MaxRedeemCount").Value
                                If (maxRedeemCount > 255) Then
                                    ErrorCode = ERROR_CODES.ERROR_INVALID_PROGRAMUSES
                                    ErrorMessage &= "; " & "Trackable coupon program MaxRedeemCount can not be more than 255."
                                End If
                            Else
                                maxRedeemCount = 1
                            End If
                        Else
                            If (status) Then
                                Dim wMsg As String = "Program is already existing in the system, MaxRedeemCount for existing program can not be overwritten"
                                Copient.Logger.Write_Log(AcceptanceLogFile, wMsg, True)
                            End If
                        End If
                        Dim programUsedInOtherOffer As Boolean = False
                        If programId > 0 Then
                            'Validate if the tcp being imported is already in use with some offer. 
                            Dim amsResult As AMSResult(Of List(Of CMS.AMS.Models.Offer)) = offerService.GetofferByTrackableProgramID(programId)
                            If amsResult.Result IsNot Nothing AndAlso amsResult.Result.Count > 0 Then
                                'TCP exists and is associated with some offer                                
                                'Offer being ADDED OR UPDATED using UpdateOffer
                                If amsResult.Result(0).OfferID <> OfferID Then
                                    programUsedInOtherOffer = True
                                End If
                            End If
                        End If
                        If programUsedInOtherOffer = False Then
                            Dim dt As New DataTable()
                            If (TCListNodes IsNot Nothing AndAlso TCListNodes.Count > 0) Then
                                ExtProgramID = ExtProgramIDNodes(0).InnerText
                                dt = MyCpeOffer.ConvertXmlNodeListToDataTable(OfferXmlDoc.SelectNodes("//Offer/Conditions/TrackableCoupon/TCList"), ExtProgramID, programId, maxRedeemCount, ErrorMessage)
                            End If
                            If (dt.Rows.Count > 0) Then
                                'Validate if the coupon is unique, coupon length,alphanumeric. 
                                dt = MyCpeOffer.ValidateTrackableCoupon(dt, ErrorMessage)
                            End If
                            If extProgramIDExists Then

                                AddedCoupons = MyCpeOffer.AddTrackableCoupons(dt, ExtProgramID, programId, maxRedeemCount, ErrorMessage)
                            Else
                                'ExtProgramID does not exists, so create it and add the coupons
                                programId = MyCpeOffer.CreateTrackableCouponsProgram(MyCommon, ExtProgramID, ExtProgramID, "New Trackable Coupons Program created for ExtProgramId: " & ExtProgramID, maxRedeemCount)
                                If programId > 0 Then
                                    AddedProgramID = True
                                End If
                                If dt.Rows.Count > 0 Then
                                    AddedCoupons = MyCpeOffer.AddTrackableCoupons(dt, ExtProgramID, programId, 1, ErrorMessage)
                                End If
                            End If

                            MyCommon.QueryStr = "Select TCC.TrackableCouponConditionID " &
                                "FROM OfferRegularConditions ORC with (NoLock) " +
                                "INNER JOIN Conditions C with (NoLock) ON C.ConditionID = ORC.ConditionID AND ORC.OfferID = @OfferID AND ORC.EngineID = @EngineID " &
                                "AND C.ConditionTypeID = 15 AND C.Deleted = 0 " &
                        "INNER JOIN TrackableCouponsCondition TCC with (NoLock) on C.ConditionID = TCC.ConditionID " &
                                "INNER JOIN TrackableCouponProgram TCP with (NoLock) on TCP.ProgramID = TCC.ProgramID AND TCP.Deleted = 0"
                            MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
                            MyCommon.DBParameters.Add("@EngineID", SqlDbType.Int).Value = EngineID

                            dtTC = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                            If (dtTC.Rows.Count > 0) Then
                                MyCommon.QueryStr = "delete from TrackableCouponsCondition with (RowLock) where ConditionID = @ConditionID"
                                MyCommon.DBParameters.Add("@ConditionID", SqlDbType.Int).Value = MyCommon.NZ(dtTC.Rows(0).Item("TrackableCouponConditionID"), 0)
                                MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
                            End If

                            MyCpeOffer.CreateTrackableCouponCondition(EngineID, OfferID, programId)

                            If (AddedProgramID AndAlso AddedCoupons) Or AddedCoupons Then
                                MyCpeOffer.UpdateTrackableCouponProgramLastLoadStatus("Coupon codes added successfully " & DateTime.Now, programId, True)
                            ElseIf AddedProgramID Then
                                MyCpeOffer.UpdateTrackableCouponProgramLastLoadStatus("External program created successfully " & DateTime.Now, programId, True)
                            Else
                                MyCpeOffer.UpdateTrackableCouponProgramLastLoadStatus("An error occurred while processing Couponcodes", programId)
                            End If
                        Else
                            ErrorMessage &= "; " & "Trackable coupon program id is associated with other offer. A trackable coupon program can be used in only one offer."
                        End If
                    Else
                        ErrorMessage &= "; " & "No external programId provided. Cannot process Trackable Coupon condition."
                    End If
                Else
                    ErrorMessage &= "Trackable coupons are supported only in Universal Engine."
                End If
            End If

        Catch ex As Exception
            ErrorMessage &= "; " & ex.ToString
        Finally
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try

    End Sub

    Private Sub ProcessProductConditions(ByVal OfferID As Long, ByVal OfferXmlDoc As XmlDocument, ByRef ErrorMessage As String)
        Dim ProductGroupNodes, ProductListNodes As XmlNodeList
        Dim ProdNode As XmlNode
        Dim ProductGroupID As Long
        Dim ProductGroupIDs(-1) As Long
        Dim ProductList As String = ""
        Dim Added As Boolean = False
        Dim MyCpeOffer As New Copient.CPEOffer
        Dim Condition As Copient.OfferProductGroup
        Dim MyCommon As New Copient.CommonInc

        Try
            If OfferXmlDoc IsNot Nothing Then
                If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

                ProductGroupNodes = OfferXmlDoc.SelectNodes("//Offer/Conditions/Product/ProductGroupID")
                ProductListNodes = OfferXmlDoc.SelectNodes("//Offer/Conditions/Product/ProductList")

                If (ProductGroupNodes IsNot Nothing AndAlso ProductGroupNodes.Count > 0) OrElse (ProductListNodes IsNot Nothing AndAlso ProductListNodes.Count > 0) Then
                    ' if any new product conditions are being sent, then remove the existing ones.
                    MyCpeOffer.RemoveAllProductConditions(OfferID, 1)

                    ' attempt to use an existing product group by ID, if it doesn't exist a new product group will be created
                    If ProductGroupNodes IsNot Nothing Then
                        For Each ProdNode In ProductGroupNodes
                            Long.TryParse(ProdNode.InnerText, ProductGroupID)
                            Condition = BuildProductCondition(OfferID, ProductGroupID, ProdNode)
                            Added = MyCpeOffer.AddProductCondition(Condition, ErrorMessage)
                            ' store the created customer group in case we need to remove it later in processing due to offer failure  
                            If Added AndAlso ProductGroupID <> Condition.GetProductGroupID Then
                                CreatedProductGroups.Add(Condition.GetProductGroupID)
                            End If
                        Next
                    End If

                    ' create a new product group and place the products in the list into the group.
                    If ProductListNodes IsNot Nothing Then
                        For Each ProdNode In ProductListNodes
                            ProductList = ProdNode.InnerText
                            Condition = BuildProductCondition(OfferID, 0, ProdNode)
                            If ProdNode.Attributes("name") IsNot Nothing Then
                                Added = MyCpeOffer.AddProductCondition(Condition, ErrorMessage, ProductList, ProdNode.Attributes("name").InnerText)
                            Else
                                Added = MyCpeOffer.AddProductCondition(Condition, ErrorMessage, ProductList)
                            End If
                            ' store the created customer group in case we need to remove it later in processing due to offer failure  
                            If Added Then CreatedProductGroups.Add(Condition.GetProductGroupID)
                        Next
                    End If
                End If

                ' check to see if at least one product condition exists, if not then create one.
                If Not Added Then
                    ProductGroupIDs = MyCpeOffer.GetConditionalProductGroups(OfferID, ErrorMessage)
                    If ProductGroupIDs.Length = 0 Then
                        Condition = New Copient.OfferProductGroup(OfferID, 0)
                        Added = MyCpeOffer.AddProductCondition(Condition, ErrorMessage)
                        ' store the created customer group in case we need to remove it later in processing due to offer failure  
                        If Added Then CreatedProductGroups.Add(Condition.GetProductGroupID)
                    End If
                End If
            End If
        Catch ex As Exception
            ErrorMessage &= "; " & ex.ToString
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try

    End Sub

    Private Sub CmProcessProductConditions(ByVal OfferID As Long, ByVal OfferXmlDoc As XmlDocument, ByRef ErrorMessage As String, Optional ByVal AutoDeploy As Boolean = False)
        Dim ProductGroupNodes, ProductListNodes As XmlNodeList
        Dim ProdNode As XmlNode
        Dim ProductGroupID As Long
        Dim ProductGroupIDs(-1) As Long
        Dim ProductList As String = ""
        Dim Added As Boolean = False
        Dim AddedProductGroupID As Long = 0
        Dim Condition As Copient.OfferProductGroup

        Try
            If OfferXmlDoc IsNot Nothing Then
                If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

                ProductGroupNodes = OfferXmlDoc.SelectNodes("//Offer/Conditions/Product/ProductGroupID")
                ProductListNodes = OfferXmlDoc.SelectNodes("//Offer/Conditions/Product/ProductList")

                If (ProductGroupNodes IsNot Nothing AndAlso ProductGroupNodes.Count > 0) OrElse (ProductListNodes IsNot Nothing AndAlso ProductListNodes.Count > 0) Then
                    ' if any new product conditions are being sent, then remove the existing ones.
                    MyCmOffer.RemoveAllProductConditions(lUserId)

                    ' attempt to use an existing product group by ID, if it doesn't exist a new product group will be created
                    If ProductGroupNodes IsNot Nothing Then
                        For Each ProdNode In ProductGroupNodes
                            Long.TryParse(ProdNode.InnerText, ProductGroupID)
                            Condition = BuildProductCondition(OfferID, ProductGroupID, ProdNode)
                            Added = MyCmOffer.AddProductCondition(Condition, ErrorMessage, AutoDeploy)
                            ' store the created product group in case we need to remove it later in processing due to offer failure  
                            If Added AndAlso ProductGroupID <> Condition.GetProductGroupID Then
                                CreatedProductGroups.Add(Condition.GetProductGroupID)
                            End If
                        Next
                    End If

                    ' create a new product group and place the products in the list into the group.
                    If ProductListNodes IsNot Nothing Then
                        For Each ProdNode In ProductListNodes
                            ProductList = ProdNode.InnerText
                            Condition = BuildProductCondition(OfferID, 0, ProdNode)
                            If ProdNode.Attributes("name") IsNot Nothing Then
                                Added = MyCmOffer.AddProductCondition(Condition, ErrorMessage, ProductList, ProdNode.Attributes("name").InnerText)
                            Else
                                Added = MyCmOffer.AddProductCondition(Condition, ErrorMessage, ProductList)
                            End If
                            ' store the created customer group in case we need to remove it later in processing due to offer failure  
                            If Added Then CreatedProductGroups.Add(Condition.GetProductGroupID)
                        Next
                    End If
                End If

                ' check to see if at least one product condition exists, if not then create one.
                'If Not Added Then
                '  ProductGroupIDs = MyCmOffer.GetConditionalProductGroups(ErrorMessage)
                '  If ProductGroupIDs.Length = 0 Then
                '    Condition = New Copient.OfferProductGroup(OfferID, 0)
                '    Added = MyCmOffer.AddProductCondition(Condition, ErrorMessage)
                '    ' store the created customer group in case we need to remove it later in processing due to offer failure  
                '    If Added Then CreatedProductGroups.Add(Condition.GetProductGroupID)
                '  End If
                'End If
            End If
        Catch ex As Exception
            ErrorMessage &= "; " & ex.ToString
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try

    End Sub

    Private Function BuildProductCondition(ByVal OfferID As Long, ByVal ProductGroupID As Long, ByVal ProductNode As XmlNode) As Copient.OfferProductGroup
        Dim Condition As Copient.OfferProductGroup
        Dim Quantity As Decimal
        Dim UnitType As Integer
        Dim FullPrice As Integer
        Dim ClearanceState As Integer
        Dim ClearanceLevel As Integer

        Condition = New Copient.OfferProductGroup(OfferID, ProductGroupID)

        If ProductNode IsNot Nothing Then
            If ProductNode.Attributes("quantity") IsNot Nothing Then
                Decimal.TryParse(ProductNode.Attributes("quantity").InnerText, Quantity)
                Condition.SetQuantity(Quantity)
            End If

            If ProductNode.Attributes("unitType") IsNot Nothing Then
                Select Case ProductNode.Attributes("unitType").InnerText
                    Case "ITEMS"
                        UnitType = 1
                    Case "DOLLARS"
                        UnitType = 2
                    Case "WEIGHT_VOLUME"
                        UnitType = 3
                End Select
                Condition.SetUnitType(UnitType)
            End If

            If ProductNode.SelectSingleNode("//Offer/Conditions/Product/FullPrice") IsNot Nothing Then
                Integer.TryParse(ProductNode.SelectSingleNode("//Offer/Conditions/Product/FullPrice").InnerText, FullPrice)
                Condition.SetFullPrice(FullPrice)
            End If
            If ProductNode.SelectSingleNode("//Offer/Conditions/Product/ClearanceState") IsNot Nothing Then
                Integer.TryParse(ProductNode.SelectSingleNode("//Offer/Conditions/Product/ClearanceState").InnerText, ClearanceState)
                Condition.SetClearanceState(ClearanceState)
            End If
            If ProductNode.SelectSingleNode("//Offer/Conditions/Product/ClearanceLevel") IsNot Nothing Then
                Integer.TryParse(ProductNode.SelectSingleNode("//Offer/Conditions/Product/ClearanceLevel").InnerText, ClearanceLevel)
                Condition.SetClearanceLevel(ClearanceLevel)
            End If
        End If

        Return Condition
    End Function
    Private Sub LogWarningMsg(ByVal sMsg As String)
        Dim wMsg As String = String.Empty
        If sErrorMethod.Length > 0 Then
            wMsg = sErrorMethod
        End If
        If sErrorSubMethod.Length > 0 Then
            wMsg += sErrorSubMethod & " - "
        End If
        wMsg = wMsg & sMsg
        Copient.Logger.Write_Log(EOCLogFile, wMsg, True)
    End Sub
    Private Sub ProcessDiscount(ByVal OfferID As Long, ByVal OfferXmlDoc As XmlDocument,
                                ByVal ExternalSourceID As String, ByVal EngineID As Int32, ByRef ErrorMessage As String)
        Dim MyCommon As New Copient.CommonInc
        Dim ProductNodes As XmlNodeList
        Dim ProdNode As XmlNode
        Dim ProductGroupID As Long = 0
        Dim ProductGroupIDs(-1) As Long
        Dim ProductList As String = ""
        Dim MyCpeOffer As New Copient.CPEOffer
        Dim MyDiscount As New Copient.Discount
        Dim TempValue As String = "'"
        Dim TempInt As Integer
        Dim TempDec As Decimal
        Dim TempBool As Boolean = False
        Dim CreateEmptyGroup As Boolean = False
        Dim SameAsProductCon As Boolean = False
        Dim NewDiscountID As Long
        Dim ChargebackDeptID As Integer = -1
        Dim BannerID As Integer = -1
        Dim sReceiptDescription As String
        Try
            sErrorSubMethod = "(ProcessDiscount)"
            If OfferXmlDoc IsNot Nothing Then
                If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
                MyCpeOffer.SetEngineID(EngineID)
                ' if a discount already exists then load it up
                MyDiscount = MyCpeOffer.GetOfferDiscount(OfferID, ErrorMessage)
                If MyDiscount.GetDiscountID <= 0 Then MyDiscount = New Copient.Discount

                ' discount type
                If TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/DiscountType", TempValue) Then
                    Select Case TempValue
                        Case "ITEM_LEVEL"
                            MyDiscount.SetDiscountTypeID(Copient.Discount.DISCOUNT_TYPES.ITEM)
                        Case "DEPARTMENT_LEVEL"
                            MyDiscount.SetDiscountTypeID(Copient.Discount.DISCOUNT_TYPES.DEPARTMENT)
                        Case "BASKET_LEVEL"
                            MyDiscount.SetDiscountTypeID(Copient.Discount.DISCOUNT_TYPES.BASKET)
                        Case Else ' treat as item_level
                            MyDiscount.SetDiscountTypeID(Copient.Discount.DISCOUNT_TYPES.ITEM)
                    End Select
                End If

                ' product group(s)
                ProdNode = OfferXmlDoc.SelectSingleNode("//Offer/Rewards/Discount/SameAsConditionProducts")
                If ProdNode IsNot Nothing Then
                    If Boolean.TryParse(ProdNode.InnerText(), SameAsProductCon) AndAlso SameAsProductCon Then
                        ProductGroupIDs = MyCpeOffer.GetConditionalProductGroups(OfferID, ErrorMessage)
                        If (ProductGroupIDs.Length > 0) Then
                            ' reward product groups only allow one product group, so use the first one
                            MyDiscount.SetDiscountProductGroup(ProductGroupIDs(0))
                        End If
                    End If
                Else
                    ProductNodes = OfferXmlDoc.SelectNodes("//Offer/Rewards/Discount/ProductGroupID")
                    If ProductNodes IsNot Nothing AndAlso ProductNodes.Count > 0 Then
                        For Each ProdNode In ProductNodes
                            Long.TryParse(ProdNode.InnerText, ProductGroupID)
                            If MyCpeOffer.ProductGroupExists(ProductGroupID, MyCommon) Then
                                MyDiscount.SetDiscountProductGroup(ProductGroupID)
                            End If
                        Next
                    End If

                    ProductNodes = OfferXmlDoc.SelectNodes("//Offer/Rewards/Discount/ProductList")
                    If ProductNodes IsNot Nothing AndAlso ProductNodes.Count > 0 Then
                        For Each ProdNode In ProductNodes
                            ProductList = ProdNode.InnerText
                            'If ProductList <> "" Then
                            If ProdNode.Attributes("name") IsNot Nothing Then
                                ProductGroupID = MyCpeOffer.CreateProductGroup(OfferID, MyCommon, ErrorMessage, ProductList, ProdNode.Attributes("name").InnerText)
                            Else
                                ProductGroupID = MyCpeOffer.CreateProductGroup(OfferID, MyCommon, ErrorMessage, ProductList)
                            End If
                            If ProductList.Trim() <> String.Empty Then _PGIDforOCD = ProductGroupID
                            MyDiscount.SetDiscountProductGroup(ProductGroupID)
                            'End If
                        Next
                    End If
                End If

                ' determine if there is already a discounted product group assigned to the offer ,if not create one
                If MyDiscount.GetDiscountProductGroup <= 0 Then
                    ProductGroupID = MyCpeOffer.CreateProductGroup(OfferID, MyCommon, ErrorMessage)
                    MyDiscount.SetDiscountProductGroup(ProductGroupID)
                End If

                ' distribution variables
                If TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/AmountType", TempValue) Then
                    Select Case TempValue
                        Case "FIXED_AMOUNT_OFF"
                            MyDiscount.SetAmountTypeID(Copient.Discount.AMOUNT_TYPES.FIXED_AMOUNT_OFF)
                        Case "PERCENT_OFF"
                            MyDiscount.SetAmountTypeID(Copient.Discount.AMOUNT_TYPES.PERCENTAGE_OFF)
                        Case "FREE"
                            MyDiscount.SetAmountTypeID(Copient.Discount.AMOUNT_TYPES.FREE)
                        Case "PRICE_POINT_ITEMS"
                            MyDiscount.SetAmountTypeID(Copient.Discount.AMOUNT_TYPES.PRICE_POINT_ITEMS)
                        Case "FIXED_AMOUNT_OFF_WEIGHT_VOLUME"
                            MyDiscount.SetAmountTypeID(Copient.Discount.AMOUNT_TYPES.FIXED_AMOUNT_WT_VOL)
                        Case "PRICE_POINT_WEIGHT_VOLUME"
                            MyDiscount.SetAmountTypeID(Copient.Discount.AMOUNT_TYPES.PRICE_POINT_WT_VOL)
                        Case Else ' treat as fixed amount off
                            MyDiscount.SetAmountTypeID(Copient.Discount.AMOUNT_TYPES.FIXED_AMOUNT_OFF)
                    End Select
                End If

                If TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/DiscountAmount", TempValue) Then
                    If Decimal.TryParse(TempValue, TempDec) Then
                        MyDiscount.SetDiscountAmount(TempDec)
                    End If
                End If

                If TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/DollarLimit", TempValue) Then
                    If Decimal.TryParse(TempValue, TempDec) Then
                        MyDiscount.SetDollarLimit(TempDec)
                    End If
                End If

                If TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/ItemLimit", TempValue) Then
                    If Integer.TryParse(TempValue, TempInt) Then
                        MyDiscount.SetItemLimit(TempInt)
                    End If
                End If

                If TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/BestDeal", TempValue) Then
                    If Boolean.TryParse(TempValue, TempBool) Then
                        MyDiscount.SetBestDeal(TempBool)
                    End If
                Else
                    ' use the default
                    MyDiscount.SetBestDeal(MyCommon.Fetch_CPE_SystemOption(17) = "1")
                End If

                If TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/ScorecardID", TempValue) Then
                    If Integer.TryParse(TempValue, TempInt) Then
                        MyDiscount.SetScorecardID(TempInt)
                    End If
                End If

                If TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/ScorecardName", TempValue) Then
                    MyDiscount.SetScorecardName(Left(TempValue, 200))
                End If

                If TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/ScorecardText", TempValue) Then
                    MyDiscount.SetScorecardDesc(Left(TempValue, 40))
                End If

                MyDiscount.SetComputeDiscount(True)
                ' ComputeDiscount field is only applied when the amount type is fixed amount weight/volume is sent; otherwise it's ignored
                If MyDiscount.GetAmountTypeID = Copient.Discount.AMOUNT_TYPES.FIXED_AMOUNT_WT_VOL Then
                    If TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/ComputeDiscount", TempValue) Then
                        If Boolean.TryParse(TempValue, TempBool) Then
                            MyDiscount.SetComputeDiscount(TempBool)
                        End If
                    End If
                End If
                If EngineID = 9 AndAlso TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/SVProgramID", TempValue) Then
                    MyCommon.QueryStr = "select SVTypeID, Isnull(SVPPES.AllowAnyCustomer,0) as AllowAnyCustomer from StoredValuePrograms SVP left join SVProgramsPromoEngineSettings SVPPES with (NoLock) on SVP.SVProgramID =SVPPES.SVProgramID where SVP.Deleted=0 and SVP.SVProgramID =" & TempValue
                    Dim rstTemp As DataTable = MyCommon.LRT_Select
                    If (rstTemp.Rows.Count <= 0) Then
                        LogWarningMsg("Stored value program ID:|" & TempValue & "| does not exists.")
                    Else
                        If (rstTemp.Rows(0)("SVTypeID") = "2") Then
                            Dim bAnyCustomer As Boolean = IsAnyCustomerOffer(OfferXmlDoc, EngineID)
                            If ((bAnyCustomer = False) OrElse (bAnyCustomer = True And rstTemp.Rows(0)("AllowAnyCustomer") = "1")) Then
                                MyDiscount.SetSVProgramID(TempValue)
                            Else
                                LogWarningMsg("Stored value program ID:|" & TempValue & "|  is not Any Customer enabled.")
                            End If
                        Else
                            LogWarningMsg("Stored value program ID:|" & TempValue & "|  is not monetory. Only Monetory Stored Value Program allowed in discount reward.")
                        End If
                    End If
                End If

                If MyDiscount.GetAmountTypeID = Copient.Discount.AMOUNT_TYPES.FIXED_AMOUNT_WT_VOL _
                OrElse MyDiscount.GetAmountTypeID = Copient.Discount.AMOUNT_TYPES.PRICE_POINT_WT_VOL Then
                    If TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/WeightVolumeLimit", TempValue) Then
                        If Decimal.TryParse(TempValue, TempDec) Then
                            MyDiscount.SetWeightLimit(TempDec)
                        End If
                    End If
                End If

                If EngineID = 9 AndAlso TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/PriceFilter", TempValue) Then
                    Dim IsValid As Boolean = False
                    If Integer.TryParse(TempValue, TempDec) Then
                        If TempDec >= 0 Then
                            MyCommon.QueryStr = "SELECT ucl.ClearanceLevelValue AS id FROM dbo.UE_ClearanceLevels ucl WITH (NOLOCK) WHERE ucl.ClearanceLevelValue =  @PriceFilter" &
                                                                                " UNION SELECT PriceFilterId as id FROM  dbo.UE_PriceFilter pf WITH (NOLOCK) WHERE pf.PriceFilterId = @PriceFilter"
                            MyCommon.DBParameters.Add("@PriceFilter", SqlDbType.Int).Value = Integer.Parse(TempDec)
                            Dim dt As DataTable = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                            If (dt.Rows.Count > 0) Then
                                MyDiscount.SetPriceFilter(TempDec)
                                IsValid = True
                            End If
                        End If
                    End If
                    If Not IsValid Then
                        Copient.Logger.Write_Log(RejectionLogFile, "Discount Price Filter value:'" & TempValue & "' is not valid. Price filter expects a valid numeric value.", True)
                    End If
                End If
                ' chargeback dept. - apply the interface options for the appropriate discount type default
                '                    only in non-bannered installations.
                If MyCommon.Fetch_SystemOption(66) <> "1" Then
                    Select Case MyDiscount.GetDiscountTypeID
                        Case Copient.Discount.DISCOUNT_TYPES.ITEM
                            TempValue = MyCommon.Fetch_InterfaceOption(36)
                        Case Copient.Discount.DISCOUNT_TYPES.DEPARTMENT
                            TempValue = MyCommon.Fetch_InterfaceOption(37)
                        Case Copient.Discount.DISCOUNT_TYPES.BASKET
                            TempValue = MyCommon.Fetch_InterfaceOption(38)
                    End Select
                    ChargebackDeptID = LookupChargebackDeptID(TempValue)
                End If

                If TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/ChargebackDept", TempValue) AndAlso Not String.IsNullOrEmpty(TempValue) Then
                    ChargebackDeptID = LookupChargebackDeptID(TempValue)
                End If

                If ChargebackDeptID > -1 Then
                    MyDiscount.SetChargeBackDeptID(ChargebackDeptID)
                End If

                ' Overwrite receipt description with externaL source default?
                sReceiptDescription = ""
                If MyCommon.Fetch_InterfaceOption(55) = "1" Then
                    Dim MfgIsSet As Boolean
                    Dim MfgCouponValue As Boolean = False
                    MfgIsSet = IsManufacturerCouponSet(OfferXmlDoc, MfgCouponValue)
                    sReceiptDescription = GetDefaultReceiptMessage(ExternalSourceID, MfgCouponValue)
                End If
                If sReceiptDescription = "" Then
                    TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/ReceiptDescription", sReceiptDescription)
                End If
                If sReceiptDescription <> "" Then
                    MyDiscount.SetReceiptDescription(Left(sReceiptDescription, 100))
                End If

                Dim ROID As Long = MyCpeOffer.GetOfferROID(OfferID, MyCommon)
                ' if the discount already existed then save the changes, otherwise create a new one
                If MyDiscount.GetDiscountID > 0 Then
                    MyCpeOffer.SaveDiscount(OfferID, 1, MyDiscount, ErrorMessage, ROID)
                Else
                    NewDiscountID = MyCpeOffer.AddDiscount(OfferID, 1, MyDiscount, ErrorMessage, ROID)
                    MyDiscount.SetDiscountID(NewDiscountID)
                End If

                ' overwrite discount settings with connector-specified option value settings.
                ApplyDiscountOptions(ExternalSourceID, OfferXmlDoc, MyDiscount.GetDiscountID, MyDiscount)

            End If

        Catch ex As Exception
            ErrorMessage &= "; " & ex.ToString
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try

    End Sub

    Private Sub CmProcessDiscount(ByVal ExternalSourceID As String, ByVal OfferID As Long, ByVal OfferXmlDoc As XmlDocument, ByRef ErrorMessage As String)
        Dim ProductNodes As XmlNodeList
        Dim ProdNode As XmlNode
        Dim ProductGroupID As Long = 0
        Dim ProductGroupIDs(-1) As Long
        Dim ProductList As String = ""
        Dim MyDiscount As Copient.CmDiscount
        Dim TempValue As String = "'"
        Dim TempDec As Decimal
        Dim TempBool As Boolean = False
        Dim CreateEmptyGroup As Boolean = False
        Dim IsMfgCoupon As Boolean = False
        Dim ProductType As Integer
        Dim iRollupQty As Integer = 0
        Dim iItemLimitQty As Integer = 0
        Dim ChargebackDeptID As Integer = -1
        Dim BannerID As Integer = -1
        Dim sReceiptDescription As String
        Dim iApplyToQty As Integer = 1

        Try
            If OfferXmlDoc IsNot Nothing Then
                If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

                ' if a discount already exists then load it up
                MyDiscount = MyCmOffer.GetOfferDiscount(ErrorMessage)
                If MyDiscount.GetDiscountID <= 0 Then MyDiscount = New Copient.CmDiscount

                IsMfgCoupon = GetManufacturerCoupon(ExternalSourceID, OfferXmlDoc)
                MyDiscount.SetSponsorID(IIf(IsMfgCoupon, 1, 0))

                ' discount amount
                If TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/DiscountAmount", TempValue) Then
                    If Decimal.TryParse(TempValue, TempDec) Then
                        MyDiscount.SetDiscountAmount(TempDec)
                    End If
                End If

                ' discount type
                If TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/DiscountType", TempValue) Then
                    Select Case TempValue
                        Case "ITEM_LEVEL"
                            ' CM item level with product types of UPC
                            MyDiscount.SetDiscountTypeID(Copient.CmDiscount.DISCOUNT_TYPES.ITEM)
                        Case "DEPARTMENT_LEVEL"
                            ' CM item level with product types of Department
                            MyDiscount.SetDiscountTypeID(Copient.CmDiscount.DISCOUNT_TYPES.DEPARTMENT)
                        Case "BASKET_LEVEL"
                            ' CM transaction level
                            MyDiscount.SetDiscountTypeID(Copient.CmDiscount.DISCOUNT_TYPES.BASKET)
                        Case Else ' treat as item_level
                            MyDiscount.SetDiscountTypeID(Copient.CmDiscount.DISCOUNT_TYPES.ITEM)
                    End Select
                End If

                If MyDiscount.GetDiscountTypeID = Copient.CmDiscount.DISCOUNT_TYPES.BASKET Then
                    ' Transaction level, so no product group
                    MyDiscount.SetDiscountProductGroup(0)
                    If TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/AmountType", TempValue) Then
                        Select Case TempValue
                            Case "FIXED_AMOUNT_OFF"
                                MyDiscount.SetAmountTypeID(Copient.CmDiscount.AMOUNT_TYPES.TRX_AMOUNT_OFF)
                            Case "PERCENT_OFF"
                                MyDiscount.SetAmountTypeID(Copient.CmDiscount.AMOUNT_TYPES.TRX_PERCENTAGE_OFF)
                            Case Else ' treat as fixed amount off
                                MyDiscount.SetAmountTypeID(Copient.CmDiscount.AMOUNT_TYPES.TRX_AMOUNT_OFF)
                        End Select
                    End If
                Else
                    ' Item level, so get product group(s)
                    ProdNode = OfferXmlDoc.SelectSingleNode("//Offer/Rewards/Discount/SameAsConditionProducts")
                    If ProdNode IsNot Nothing Then
                        ProductGroupIDs = MyCmOffer.GetConditionalProductGroups(ErrorMessage)
                        If (ProductGroupIDs.Length > 0) Then
                            ' reward product groups only allow one product group, so use the first one
                            MyDiscount.SetDiscountProductGroup(ProductGroupIDs(0))
                        End If
                    Else
                        ProductNodes = OfferXmlDoc.SelectNodes("//Offer/Rewards/Discount/ProductGroupID")
                        If ProductNodes IsNot Nothing AndAlso ProductNodes.Count > 0 Then
                            For Each ProdNode In ProductNodes
                                Long.TryParse(ProdNode.InnerText, ProductGroupID)
                                If MyCmOffer.ProductGroupExists(ProductGroupID) Then
                                    MyDiscount.SetDiscountProductGroup(ProductGroupID)
                                End If
                            Next
                        End If

                        ProductNodes = OfferXmlDoc.SelectNodes("//Offer/Rewards/Discount/ProductList")
                        If ProductNodes IsNot Nothing AndAlso ProductNodes.Count > 0 Then
                            For Each ProdNode In ProductNodes
                                ProductList = ProdNode.InnerText
                                If MyDiscount.GetDiscountTypeID = Copient.CmDiscount.DISCOUNT_TYPES.DEPARTMENT Then
                                    ProductType = 2
                                Else
                                    ProductType = 0
                                End If
                                If ProdNode.Attributes("name") IsNot Nothing Then
                                    ProductGroupID = MyCmOffer.CreateProductGroup(ErrorMessage, ProductList, ProductType, ProdNode.Attributes("name").InnerText)
                                Else
                                    ProductGroupID = MyCmOffer.CreateProductGroup(ErrorMessage, ProductList, ProductType)
                                End If
                                MyDiscount.SetDiscountProductGroup(ProductGroupID)
                            Next
                        End If
                    End If

                    ' determine if there is already a discounted product group assigned to the offer ,if not create one
                    If MyDiscount.GetDiscountProductGroup <= 0 Then
                        ProductGroupID = MyCmOffer.CreateProductGroup(ErrorMessage)
                        MyDiscount.SetDiscountProductGroup(ProductGroupID)
                    End If
                    If TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/AmountType", TempValue) Then
                        Select Case TempValue
                            Case "FIXED_AMOUNT_OFF"
                                MyDiscount.SetAmountTypeID(Copient.CmDiscount.AMOUNT_TYPES.FIXED_AMOUNT_OFF)
                            Case "PERCENT_OFF"
                                MyDiscount.SetAmountTypeID(Copient.CmDiscount.AMOUNT_TYPES.PERCENTAGE_OFF)
                            Case "FREE"
                                MyDiscount.SetAmountTypeID(Copient.CmDiscount.AMOUNT_TYPES.FREE)
                            Case "PRICE_POINT_ITEMS"
                                MyDiscount.SetAmountTypeID(Copient.CmDiscount.AMOUNT_TYPES.PRICE_POINT_ITEMS)
                            Case "FIXED_AMOUNT_OFF_WEIGHT_VOLUME"
                                If TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/AmountWeightVolumeType", TempValue) Then
                                    Select Case TempValue
                                        Case "WEIGHT"
                                            MyDiscount.SetAmountTypeID(Copient.CmDiscount.AMOUNT_TYPES.FIXED_AMOUNT_WT)
                                        Case "VOLUME"
                                            MyDiscount.SetAmountTypeID(Copient.CmDiscount.AMOUNT_TYPES.FIXED_AMOUNT_VOL)
                                        Case Else ' Default to weight
                                            MyDiscount.SetAmountTypeID(Copient.CmDiscount.AMOUNT_TYPES.FIXED_AMOUNT_WT)
                                    End Select
                                Else
                                    MyDiscount.SetAmountTypeID(Copient.CmDiscount.AMOUNT_TYPES.FIXED_AMOUNT_WT)
                                End If
                            Case "PRICE_POINT_WEIGHT_VOLUME"
                                If TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/AmountWeightVolumeType", TempValue) Then
                                    Select Case TempValue
                                        Case "WEIGHT"
                                            MyDiscount.SetAmountTypeID(Copient.CmDiscount.AMOUNT_TYPES.PRICE_POINT_WT)
                                        Case "VOLUME"
                                            Throw New ApplicationException("CM does not support Price Point per volume.")
                                        Case Else ' Default to weight
                                            MyDiscount.SetAmountTypeID(Copient.CmDiscount.AMOUNT_TYPES.PRICE_POINT_WT)
                                    End Select
                                Else
                                    MyDiscount.SetAmountTypeID(Copient.CmDiscount.AMOUNT_TYPES.PRICE_POINT_WT)
                                End If
                            Case Else ' treat as fixed amount off
                                MyDiscount.SetAmountTypeID(Copient.CmDiscount.AMOUNT_TYPES.FIXED_AMOUNT_OFF)
                        End Select
                    End If

                    ' compare reward product group to the product groups for conditions
                    ProductGroupID = MyDiscount.GetDiscountProductGroup()
                    If ProductGroupID > 0 Then
                        iRollupQty = MyCmOffer.GetQuantityForConditionalProductGroup(ProductGroupID)
                        If TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/ItemLimit", TempValue) Then
                            If Not Integer.TryParse(TempValue, iItemLimitQty) Then
                                iItemLimitQty = 0
                            End If
                        End If

                        If TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/ApplyToQuantity", TempValue) Then
                            If Not Integer.TryParse(TempValue, iApplyToQty) Then
                                iApplyToQty = 1
                            End If
                        End If

                        If iItemLimitQty > 0 Then
                            If (MyDiscount.GetAmountTypeID = Copient.CmDiscount.AMOUNT_TYPES.PRICE_POINT_ITEMS) _
                             Or (MyDiscount.GetAmountTypeID = Copient.CmDiscount.AMOUNT_TYPES.PRICE_POINT_WT) Then
                                Throw New ApplicationException("CM does not support Item Limits for Price Point related amount types.")
                            End If

                            ' convert CPE item limit to CM special pricing
                            MyDiscount.SetItemLimitQty(iItemLimitQty)
                            If iRollupQty > 1 Then
                                If iRollupQty > iItemLimitQty Then
                                    MyDiscount.SetSpecialPricingQty(iRollupQty)
                                Else
                                    MyDiscount.SetSpecialPricingQty(iItemLimitQty)
                                End If
                            Else
                                MyDiscount.SetSpecialPricingQty(iItemLimitQty)
                            End If
                        Else
                            If (iRollupQty > 0) Then
                                ' If reward product group = condition product group, rollup condition to reward
                                MyDiscount.SetTriggerQty(iRollupQty)
                                MyCmOffer.RemoveProductConditionForGroup(ProductGroupID, lUserId)
                            End If

                            If iRollupQty = -1 Then
                                MyDiscount.SetApplyToLimit(iApplyToQty)
                                MyDiscount.SetTriggerQty(iApplyToQty)
                            Else
                                If iApplyToQty > iRollupQty Then
                                    Throw New ApplicationException("The ApplyToQuantity (" & iApplyToQty & ") can not be greater than the product quantity (" & iRollupQty & ")!")
                                Else
                                    MyDiscount.SetApplyToLimit(iApplyToQty)
                                End If
                            End If
                        End If

                        If MyDiscount.GetAmountTypeID = Copient.CmDiscount.AMOUNT_TYPES.FREE Then
                            MyDiscount.SetDiscountAmount(0.01)
                        End If
                    End If
                End If

                If TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/DollarLimit", TempValue) Then
                    If Decimal.TryParse(TempValue, TempDec) Then
                        MyDiscount.SetRewardLimit(TempDec)
                        MyDiscount.SetRewardLimitTypeID(2)
                    End If
                End If

                If TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/WeightVolumeLimit", TempValue) Then
                    If Decimal.TryParse(TempValue, TempDec) Then
                        MyDiscount.SetRewardLimit(TempDec)
                        If TryParseAttributeValue(OfferXmlDoc, "//Offer/Rewards/Discount/WeightVolumeLimit", "type", TempValue) Then
                            If TempValue = "VOLUME" Then
                                MyDiscount.SetRewardLimitTypeID(4)
                            Else
                                MyDiscount.SetRewardLimitTypeID(3)
                            End If
                        Else
                            MyDiscount.SetRewardLimitTypeID(3)
                        End If
                    End If
                End If

                If TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/BestDeal", TempValue) Then
                    If Boolean.TryParse(TempValue, TempBool) Then
                        MyDiscount.SetBestDeal(TempBool)
                    End If
                End If

                ' chargeback dept. - apply the interface options for the appropriate discount type default
                '                    only in non-bannered installations.
                If MyCommon.Fetch_SystemOption(66) <> "1" Then
                    Select Case MyDiscount.GetDiscountTypeID
                        Case Copient.CmDiscount.DISCOUNT_TYPES.ITEM
                            TempValue = MyCommon.Fetch_InterfaceOption(36)
                        Case Copient.CmDiscount.DISCOUNT_TYPES.DEPARTMENT
                            TempValue = MyCommon.Fetch_InterfaceOption(37)
                        Case Copient.CmDiscount.DISCOUNT_TYPES.BASKET
                            TempValue = MyCommon.Fetch_InterfaceOption(38)
                    End Select
                    ChargebackDeptID = LookupChargebackDeptID(TempValue)
                End If

                If TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/ChargebackDept", TempValue) Then
                    ChargebackDeptID = LookupChargebackDeptID(TempValue)
                End If

                If ChargebackDeptID > -1 Then
                    MyDiscount.SetChargeBackDeptID(ChargebackDeptID)
                End If

                ' Overwrite receipt description with externaL source default?
                sReceiptDescription = ""
                If MyCommon.Fetch_InterfaceOption(55) = "1" Then
                    Dim MfgIsSet As Boolean
                    Dim MfgCouponValue As Boolean = False
                    MfgIsSet = IsManufacturerCouponSet(OfferXmlDoc, MfgCouponValue)
                    sReceiptDescription = GetDefaultReceiptMessage(ExternalSourceID, MfgCouponValue)
                End If
                If sReceiptDescription = "" Then
                    TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/ReceiptDescription", sReceiptDescription)
                End If
                If sReceiptDescription <> "" Then
                    MyDiscount.SetPrintLineText(Left(sReceiptDescription, 100))
                End If

                If TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/PromoteToTransLevel", TempValue) Then
                    If Boolean.TryParse(TempValue, TempBool) Then
                        MyDiscount.SetPromoteToTransLevel(TempBool)
                    End If
                Else
                    ' PromoteToTransLevel is NOT set explicitly, so set for Brookshire's special case
                    If MyDiscount.GetDiscountTypeID = Copient.CmDiscount.DISCOUNT_TYPES.ITEM Then
                        If MyDiscount.GetAmountTypeID = Copient.CmDiscount.AMOUNT_TYPES.FIXED_AMOUNT_OFF Then
                            ' check if auto promote to transaction level is enabled
                            If MyCommon.Fetch_CM_SystemOption(40) = "1" Then
                                MyDiscount.SetPromoteToTransLevel(True)
                            End If
                        End If
                    End If
                End If

                If ErrorMessage = "" Then
                    ' if the discount already existed then save the changes, otherwise create a new one
                    If MyDiscount.GetDiscountID > 0 Then
                        MyCmOffer.SaveDiscount(lUserId, MyDiscount, ErrorMessage)
                    Else
                        MyCmOffer.AddDiscount(lUserId, MyDiscount, ErrorMessage)
                    End If
                End If

                If MyDiscount.GetSpecialPricingQty > 0 Then
                    MyCmOffer.HandleSpecialPricingLimit(lUserId, MyDiscount.GetSpecialPricingQty, ErrorMessage)
                End If

            End If
        Catch exApp As ApplicationException
            ErrorMessage = exApp.ToString
        Catch ex As Exception
            ErrorMessage &= "; " & ex.ToString
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try

    End Sub

    Private Sub ProcessOfferLocations(ByVal OfferID As Long, ByVal OfferXmlDoc As XmlDocument, ByRef ErrorMessage As String, Optional ByVal EngineID As Integer = -1)
        Dim MyCommon As New Copient.CommonInc
        Dim LocGroupNodes, LocListNodes As XmlNodeList
        Dim LocNode As XmlNode
        Dim LocationGroupID As Long
        Dim StoreList As String = ""
        Dim Added As Boolean = False
        Dim MyCpeOffer As New Copient.CPEOffer
        Dim BannerIDs(-1) As Long
        Dim RetCode As Copient.CPEOffer.RETURN_CODES = RETURN_CODES.UNSET
        Dim OfferLocations(-1) As Copient.OfferLocation

        Try
            LocGroupNodes = OfferXmlDoc.SelectNodes("//Offer/Stores/StoreGroupID")
            LocListNodes = OfferXmlDoc.SelectNodes("//Offer/Stores/StoreList")

            If (LocGroupNodes IsNot Nothing AndAlso LocGroupNodes.Count > 0) OrElse (LocListNodes IsNot Nothing AndAlso LocListNodes.Count > 0) Then
                If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

                ' if any new offerlocations are being sent, then remove the existing ones.
                MyCpeOffer.RemoveAllOfferLocations(OfferID, 1, True)

                BannerIDs = GetBannerIDs(OfferXmlDoc)

                If OfferXmlDoc IsNot Nothing Then
                    If LocGroupNodes IsNot Nothing Then
                        For Each LocNode In LocGroupNodes
                            Long.TryParse(LocNode.InnerText, LocationGroupID)
                            If MyCpeOffer.LocationGroupExists(LocationGroupID, MyCommon) Then
                                RetCode = MyCpeOffer.AddOfferLocation(OfferID, LocationGroupID, False, 1)
                                If (RetCode = RETURN_CODES.APPLICATION_ERROR) Then
                                    Throw New ApplicationException(String.Format("An error occurred adding locations"))
                                End If
                                Added = (RetCode = RETURN_CODES.ADDED OrElse RetCode = RETURN_CODES.ALREADY_EXISTS)
                            End If
                        Next
                    End If

                    If LocListNodes IsNot Nothing Then
                        For Each LocNode In LocListNodes
                            StoreList = LocNode.InnerText
                            'If StoreList <> "" Then
                            If LocNode.Attributes("name") IsNot Nothing Then
                                LocationGroupID = MyCpeOffer.CreateLocationGroup(OfferID, BannerIDs, MyCommon, ErrorMessage, StoreList, LocNode.Attributes("name").InnerText, EngineID)
                            Else
                                LocationGroupID = MyCpeOffer.CreateLocationGroup(OfferID, BannerIDs, MyCommon, ErrorMessage, StoreList, "", EngineID)
                            End If
                            RetCode = MyCpeOffer.AddOfferLocation(OfferID, LocationGroupID, False, 1)
                            If (RetCode = RETURN_CODES.APPLICATION_ERROR) Then
                                Throw New ApplicationException(String.Format("An error occurred adding locations"))
                            End If
                            Added = (RetCode = RETURN_CODES.ADDED OrElse RetCode = RETURN_CODES.ALREADY_EXISTS)
                            'End If
                        Next
                    End If

                End If


                ' if an offer location doesn't already exist, then create an empty one
                If Not Added Then
                    OfferLocations = MyCpeOffer.GetOfferLocations(OfferID)
                    If OfferLocations.Length = 0 Then
                        RetCode = MyCpeOffer.AddOfferLocation(OfferID, 0, False, 1)
                        If (RetCode = RETURN_CODES.APPLICATION_ERROR) Then
                            Throw New ApplicationException(String.Format("An error occurred adding locations"))
                        End If
                    End If
                End If

            End If

        Catch ex As Exception
            ErrorMessage &= "; " & ex.ToString
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try

    End Sub

    Private Sub CmProcessOfferLocations(ByVal OfferID As Long, ByVal OfferXmlDoc As XmlDocument, ByRef ErrorMessage As String)
        Dim LocGroupNodes, LocListNodes As XmlNodeList
        Dim LocNode As XmlNode
        Dim LocationGroupID As Long
        Dim StoreList As String = ""
        Dim Added As Boolean = False
        Dim bExcludeGroup As Boolean
        Dim BannerIDs(-1) As Long
        Dim bBannerOK As Boolean
        Dim RetCode As Copient.CMOffer.CM_RETURN_CODES = CM_RETURN_CODES.UNSET
        Dim OfferLocations(-1) As Copient.OfferLocation

        Try
            LocGroupNodes = OfferXmlDoc.SelectNodes("//Offer/Stores/StoreGroupID")
            LocListNodes = OfferXmlDoc.SelectNodes("//Offer/Stores/StoreList")

            If (LocGroupNodes IsNot Nothing AndAlso LocGroupNodes.Count > 0) OrElse (LocListNodes IsNot Nothing AndAlso LocListNodes.Count > 0) Then
                If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

                ' if any new offerlocations are being sent, then remove the existing ones.
                MyCmOffer.RemoveAllOfferLocations(lUserId, True)

                BannerIDs = GetBannerIDs(OfferXmlDoc)

                If OfferXmlDoc IsNot Nothing Then
                    If LocGroupNodes IsNot Nothing Then
                        For Each LocNode In LocGroupNodes
                            If Not Long.TryParse(LocNode.InnerText, LocationGroupID) Then LocationGroupID = 0
                            If LocationGroupID > 0 Then
                                If MyCmOffer.LocationGroupExists(LocationGroupID) Then
                                    bBannerOK = CmCheckLocationGroupBanner(LocationGroupID, BannerIDs, ErrorMessage)
                                    If bBannerOK Then
                                        If LocNode.Attributes("excluded") IsNot Nothing Then
                                            If Not Boolean.TryParse(LocNode.Attributes("excluded").InnerText, bExcludeGroup) Then bExcludeGroup = False
                                        Else
                                            bExcludeGroup = False
                                        End If
                                        RetCode = MyCmOffer.AddOfferLocation(LocationGroupID, bExcludeGroup, lUserId)
                                        Added = True
                                    End If
                                End If
                            End If
                        Next
                    End If

                    If LocListNodes IsNot Nothing Then
                        For Each LocNode In LocListNodes
                            StoreList = LocNode.InnerText
                            If LocNode.Attributes("name") IsNot Nothing Then
                                LocationGroupID = MyCmOffer.CreateLocationGroup(BannerIDs, ErrorMessage, StoreList, LocNode.Attributes("name").InnerText)
                            Else
                                LocationGroupID = MyCmOffer.CreateLocationGroup(BannerIDs, ErrorMessage, StoreList)
                            End If
                            bBannerOK = CmCheckLocationGroupBanner(LocationGroupID, BannerIDs, ErrorMessage)
                            If bBannerOK Then
                                If LocNode.Attributes("excluded") IsNot Nothing Then
                                    If Not Boolean.TryParse(LocNode.Attributes("excluded").InnerText, bExcludeGroup) Then bExcludeGroup = False
                                Else
                                    bExcludeGroup = False
                                End If
                                RetCode = MyCmOffer.AddOfferLocation(LocationGroupID, bExcludeGroup, lUserId)
                                Added = True
                            End If
                        Next
                    End If

                End If

                ' if an offer location doesn't already exist, then create one for all locations
                If Not Added Then
                    OfferLocations = MyCmOffer.GetOfferLocations()
                    If OfferLocations.Length = 0 Then
                        RetCode = MyCmOffer.AddOfferLocation(1, False, lUserId)
                    End If
                End If

            End If

        Catch ex As Exception
            ErrorMessage &= "; " & ex.ToString
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try

    End Sub

    Private Sub ProcessOfferTerminals(ByVal OfferID As Long, ByVal OfferXmlDoc As XmlDocument, ByRef ErrorMessage As String, ByVal EngineID As Int32)
        Dim MyCommon As New Copient.CommonInc
        Dim TermIDNodes, TermNameNodes As XmlNodeList
        Dim TermNode As XmlNode
        Dim TermName As String = ""
        Dim TerminalTypeID As Long
        Dim Excluded As Boolean
        Dim MyCpeOffer As New Copient.CPEOffer
        Dim BannerIDs(-1) As Long
        Dim BannerID As Long = 0
        Dim RetCode As Copient.CPEOffer.RETURN_CODES = RETURN_CODES.UNSET

        Try

            If OfferXmlDoc IsNot Nothing Then
                TermIDNodes = OfferXmlDoc.SelectNodes("//Offer/Terminals/ID")
                TermNameNodes = OfferXmlDoc.SelectNodes("//Offer/Terminals/Name")

                If (TermIDNodes IsNot Nothing AndAlso TermIDNodes.Count > 0) OrElse (TermNameNodes IsNot Nothing AndAlso TermNameNodes.Count > 0) Then
                    If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

                    BannerIDs = GetBannerIDs(OfferXmlDoc)

                    ' remove all existing offer terminals so that the new ones are the only ones assigned to the offer
                    MyCpeOffer.RemoveAllOfferTerminals(OfferID, 1, True)

                    If TermIDNodes IsNot Nothing Then
                        For Each TermNode In TermIDNodes
                            Long.TryParse(TermNode.InnerText, TerminalTypeID)
                            If TermNode.Attributes("excluded") IsNot Nothing Then
                                Excluded = TermNode.Attributes("excluded").InnerText
                            Else
                                Excluded = False
                            End If
                            'MyCommon.Write_Log("aaa.txt", "[ID]   OfferID=" & OfferID & ", TerminalTypeID=" & TerminalTypeID & ", Excluded=" & Excluded)
                            RetCode = MyCpeOffer.AddOfferTerminal(OfferID, TerminalTypeID, Excluded, 1, EngineID)
                        Next
                    End If

                    If TermNameNodes IsNot Nothing Then
                        For Each TermNode In TermNameNodes
                            TermName = TermNode.InnerText
                            ' determine if the terminal name already exists; if not, create it
                            TerminalTypeID = MyCpeOffer.GetTerminalIdFromName(TermName, MyCommon, EngineID)
                            If BannerIDs.Length > 0 Then
                                BannerID = BannerIDs(0)
                            End If
                            If TerminalTypeID = 0 Then
                                TerminalTypeID = MyCpeOffer.CreateTerminal(TermName, BannerID, 1, EngineID)
                            End If
                            If TerminalTypeID > 0 Then
                                If TermNode.Attributes("excluded") IsNot Nothing Then
                                    Excluded = TermNode.Attributes("excluded").InnerText
                                Else
                                    Excluded = False
                                End If
                                'MyCommon.Write_Log("aaa.txt", "[NAME] OfferID=" & OfferID & ", TerminalTypeID=" & TerminalTypeID & ", Excluded=" & Excluded)
                                RetCode = MyCpeOffer.AddOfferTerminal(OfferID, TerminalTypeID, Excluded, 1, EngineID)
                            End If
                        Next
                    End If
                End If

            End If

        Catch ex As Exception
            ErrorMessage &= "; " & ex.ToString
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try

    End Sub

    Private Sub CmProcessOfferTerminals(ByVal OfferID As Long, ByVal OfferXmlDoc As XmlDocument, ByRef ErrorMessage As String)
        Dim TermIDNodes, TermNameNodes, TermCodeNodes As XmlNodeList
        Dim TermNode As XmlNode
        Dim TermName As String = ""
        Dim TermCode As String = ""
        Dim bExcludeTerminal As Boolean = False
        Dim TerminalTypeID As Long
        Dim BannerIDs(-1) As Long
        Dim BannerID As Long = 0
        Dim bBannerOK As Boolean
        Dim RetCode As Copient.CMOffer.CM_RETURN_CODES = CM_RETURN_CODES.UNSET

        Try

            If OfferXmlDoc IsNot Nothing Then
                TermIDNodes = OfferXmlDoc.SelectNodes("//Offer/Terminals/ID")
                TermNameNodes = OfferXmlDoc.SelectNodes("//Offer/Terminals/Name")
                TermCodeNodes = OfferXmlDoc.SelectNodes("//Offer/Terminals/Code")

                If (TermIDNodes IsNot Nothing AndAlso TermIDNodes.Count > 0) _
                    OrElse (TermNameNodes IsNot Nothing AndAlso TermNameNodes.Count > 0) _
                    OrElse (TermCodeNodes IsNot Nothing AndAlso TermCodeNodes.Count > 0) Then
                    If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

                    MyCmOffer.RemoveAllOfferTerminals(lUserId, True)

                    BannerIDs = GetBannerIDs(OfferXmlDoc)

                    If TermIDNodes IsNot Nothing Then
                        For Each TermNode In TermIDNodes
                            If Not Long.TryParse(TermNode.InnerText, TerminalTypeID) Then TerminalTypeID = 0
                            If TerminalTypeID > 0 Then
                                bBannerOK = CmCheckTerminalBanner(TerminalTypeID, BannerIDs, ErrorMessage)
                                If bBannerOK Then
                                    If TermNode.Attributes("excluded") IsNot Nothing Then
                                        If Not Boolean.TryParse(TermNode.Attributes("excluded").InnerText, bExcludeTerminal) Then bExcludeTerminal = False
                                    Else
                                        bExcludeTerminal = False
                                    End If
                                    RetCode = MyCmOffer.AddOfferTerminal(TerminalTypeID, bExcludeTerminal, lUserId)
                                End If
                            End If
                        Next
                    End If

                    If TermNameNodes IsNot Nothing Then
                        For Each TermNode In TermNameNodes
                            TermName = TermNode.InnerText
                            ' determine if the terminal name already exists, if not create it
                            TerminalTypeID = MyCmOffer.GetTerminalIdFromName(TermName)
                            If TerminalTypeID > 0 Then
                                bBannerOK = CmCheckTerminalBanner(TerminalTypeID, BannerIDs, ErrorMessage)
                                If bBannerOK Then
                                    If TermNode.Attributes("excluded") IsNot Nothing Then
                                        If Not Boolean.TryParse(TermNode.Attributes("excluded").InnerText, bExcludeTerminal) Then bExcludeTerminal = False
                                    Else
                                        bExcludeTerminal = False
                                    End If
                                    RetCode = MyCmOffer.AddOfferTerminal(TerminalTypeID, bExcludeTerminal, lUserId)
                                End If
                            Else
                                ErrorMessage &= "; " & " Terminal with name '" & TermName & "' does not exist. Must provide terminal 'Code' to add a new terminal for a CM offer."
                            End If
                        Next
                    End If

                    If TermCodeNodes IsNot Nothing Then
                        If BannerIDs.Length > 0 Then BannerID = BannerIDs(0)
                        For Each TermNode In TermCodeNodes
                            TermCode = TermNode.InnerText
                            ' determine if the external terminal code already exists, if not create it
                            TerminalTypeID = MyCmOffer.GetTerminalIdFromCode(TermCode)

                            If TerminalTypeID = 0 Then
                                If TermNode.Attributes("name") IsNot Nothing Then
                                    TermName = TermNode.Attributes("name").InnerText
                                Else
                                    TermName = "External Code: " & TermCode
                                End If
                                TerminalTypeID = MyCmOffer.CreateTerminal(TermName, TermCode, BannerID, lUserId)
                            End If

                            If TerminalTypeID > 0 Then
                                bBannerOK = CmCheckTerminalBanner(TerminalTypeID, BannerIDs, ErrorMessage)
                                If bBannerOK Then
                                    If TermNode.Attributes("excluded") IsNot Nothing Then
                                        If Not Boolean.TryParse(TermNode.Attributes("excluded").InnerText, bExcludeTerminal) Then bExcludeTerminal = False
                                    Else
                                        bExcludeTerminal = False
                                    End If
                                    RetCode = MyCmOffer.AddOfferTerminal(TerminalTypeID, bExcludeTerminal, lUserId)
                                End If
                            End If
                        Next
                    End If
                End If

            End If

        Catch ex As Exception
            ErrorMessage &= "; " & ex.ToString
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try

    End Sub

    Private Function CmCheckTerminalBanner(ByVal TerminalTypeID As Long, ByVal BannerIDs As Long(), ByRef ErrorMessage As String) As Boolean
        Dim bBannerOK As Boolean = True
        Dim TerminalBannerId As Long
        Dim i As Integer

        If BannerIDs.Length > 0 Then
            If TerminalTypeID > 0 Then
                TerminalBannerId = MyCmOffer.GetTerminalBannerId(TerminalTypeID)
                If TerminalBannerId > 0 Then
                    bBannerOK = False
                    For i = 0 To BannerIDs.Length - 1
                        If TerminalBannerId = BannerIDs(i) Then
                            bBannerOK = True
                            Exit For
                        End If
                    Next
                    If Not bBannerOK Then
                        ErrorMessage &= "; " & " The Banner ID '" & TerminalBannerId & "' for Terminal ID '" & TerminalTypeID & "' is not included with offer"
                    End If
                End If
            End If
        End If
        Return bBannerOK
    End Function

    Private Function CmCheckLocationGroupBanner(ByVal LocationGroupID As Long, ByVal BannerIDs As Long(), ByRef ErrorMessage As String) As Boolean
        Dim bBannerOK As Boolean = True
        Dim LocationGroupBannerId As Long
        Dim i As Integer

        If BannerIDs.Length > 0 Then
            If LocationGroupID > 0 Then
                LocationGroupBannerId = MyCmOffer.GetLocationGroupBannerId(LocationGroupID)
                If LocationGroupBannerId > 0 Then
                    bBannerOK = False
                    For i = 0 To BannerIDs.Length - 1
                        If LocationGroupBannerId = BannerIDs(i) Then
                            bBannerOK = True
                            Exit For
                        End If
                    Next
                    If Not bBannerOK Then
                        ErrorMessage &= "; " & " The Banner ID '" & LocationGroupBannerId & "' for Location Group ID '" & LocationGroupID & "' is not included with offer"
                    End If
                End If
            End If
        End If
        Return bBannerOK
    End Function

    Private Sub ProcessCustomerConditions(ByVal OfferID As Long, ByVal OfferXmlDoc As XmlDocument, ByRef ErrorMessage As String, ByVal methodName As String)
        Dim CustGroupNodes, CustListNodes As XmlNodeList
        Dim CustNode As XmlNode
        Dim CustomerGroupID As Long
        Dim CustomerGroupIDs(-1) As Long
        Dim CustomerList As String = ""
        Dim Added As Boolean = False
        Dim MyCpeOffer As New Copient.CPEOffer
        Dim Condition As Copient.OfferCustomerGroup
        Dim MyCommon As New Copient.CommonInc

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If OfferXmlDoc IsNot Nothing Then

                CustGroupNodes = OfferXmlDoc.SelectNodes("//Offer/Conditions/Customer/CustomerGroupID")
                CustListNodes = OfferXmlDoc.SelectNodes("//Offer/Conditions/Customer/CustomerList")

                If (CustGroupNodes IsNot Nothing AndAlso CustGroupNodes.Count > 0) OrElse (CustListNodes IsNot Nothing AndAlso CustListNodes.Count > 0) Then
                    ' if any new customer conditions are being sent, then remove the existing ones.
                    MyCpeOffer.RemoveAllCustomerConditions(OfferID, 1)

                    ' attempt to use an existing customer group by ID, if it doesn't exist a new customer group will be created
                    If CustGroupNodes IsNot Nothing Then
                        For Each CustNode In CustGroupNodes
                            Long.TryParse(CustNode.InnerText, CustomerGroupID)
                            Condition = New Copient.OfferCustomerGroup(OfferID, CustomerGroupID)
                            Added = MyCpeOffer.AddCustomerCondition(Condition, ErrorMessage)
                            ' store the created customer group in case we need to remove it later in processing due to offer failure  
                            If Added AndAlso CustomerGroupID <> Condition.GetCustomerGroupID Then
                                CreatedCustomerGroups.Add(Condition.GetCustomerGroupID)
                            End If
                        Next
                    End If

                    ' create a new customer group and place the customers in the list into the group.
                    If CustListNodes IsNot Nothing Then
                        For Each CustNode In CustListNodes
                            CustomerList = CustNode.InnerText
                            Condition = New Copient.OfferCustomerGroup(OfferID, 0)
                            If CustNode.Attributes("name") IsNot Nothing Then
                                Added = MyCpeOffer.AddCustomerCondition(Condition, ErrorMessage, CustomerList, CustNode.Attributes("name").InnerText)
                            Else
                                Added = MyCpeOffer.AddCustomerCondition(Condition, ErrorMessage, CustomerList)
                            End If
                            ' store the created customer group in case we need to remove it later in processing due to offer failure  
                            If Added Then CreatedCustomerGroups.Add(Condition.GetCustomerGroupID)
                        Next
                    End If

                    ' add customer approval condition
                    Dim custApprovalNode = OfferXmlDoc.SelectSingleNode("//Offer/Conditions/Customer/CustomerApproval")
                    If custApprovalNode IsNot Nothing AndAlso custApprovalNode.InnerXml <> "" Then
                        Added = MyCpeOffer.AddCustomerApprovalRecord(OfferXmlDoc, OfferID, ErrorMessage, methodName)
                    End If

                    ' check to see if at least one customer condition exists, if not then create one.
                    If Not Added Then
                        CustomerGroupIDs = MyCpeOffer.GetConditionalCustomerGroups(OfferID, ErrorMessage)
                        If CustomerGroupIDs.Length = 0 Then
                            Condition = New Copient.OfferCustomerGroup(OfferID, 0)
                            Added = MyCpeOffer.AddCustomerCondition(Condition, ErrorMessage)
                            ' store the created customer group in case we need to remove it later in processing due to offer failure  
                            If Added Then CreatedCustomerGroups.Add(Condition.GetCustomerGroupID)
                        End If
                    End If

                End If
            End If
        Catch ex As Exception
            ErrorMessage &= ex.ToString
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try

    End Sub

    Private Sub CmProcessCustomerConditions(ByVal OfferID As Long, ByVal OfferXmlDoc As XmlDocument, ByRef ErrorMessage As String)
        Dim CustGroupNodes, CustListNodes As XmlNodeList
        Dim CustNode As XmlNode
        Dim CustomerGroupID As Long
        Dim CustomerGroupIDs(-1) As Long
        Dim CustomerList As String = ""
        Dim Added As Boolean = False
        Dim Condition As Copient.OfferCustomerGroup
        Dim bNoErrorsSoFar As Boolean = True

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

            If OfferXmlDoc IsNot Nothing Then

                CustGroupNodes = OfferXmlDoc.SelectNodes("//Offer/Conditions/Customer/CustomerGroupID")
                CustListNodes = OfferXmlDoc.SelectNodes("//Offer/Conditions/Customer/CustomerList")

                If (CustGroupNodes IsNot Nothing AndAlso CustGroupNodes.Count > 0) OrElse (CustListNodes IsNot Nothing AndAlso CustListNodes.Count > 0) Then
                    ' if any new customer conditions are being sent, then remove the existing ones.
                    MyCmOffer.RemoveAllCustomerConditions(lUserId)

                    ' attempt to use an existing customer group by ID, if it doesn't exist a new customer group will be created
                    If CustGroupNodes IsNot Nothing AndAlso CustGroupNodes.Count > 0 Then
                        If CustGroupNodes.Count = 1 Then
                            For Each CustNode In CustGroupNodes
                                Long.TryParse(CustNode.InnerText, CustomerGroupID)
                                Condition = New Copient.OfferCustomerGroup(OfferID, CustomerGroupID)
                                Added = MyCmOffer.AddCustomerCondition(Condition, ErrorMessage)
                                ' store the created customer group in case we need to remove it later in processing due to offer failure  
                                If Added AndAlso CustomerGroupID <> Condition.GetCustomerGroupID Then
                                    CreatedCustomerGroups.Add(Condition.GetCustomerGroupID)
                                End If
                            Next
                        Else
                            ErrorMessage &= " CM allows only 1 Customer condition! "
                            bNoErrorsSoFar = False
                        End If
                    End If

                    ' create a new customer group and place the customers in the list into the group.
                    If bNoErrorsSoFar AndAlso CustListNodes IsNot Nothing AndAlso CustListNodes.Count > 0 Then
                        If Added Then
                            ErrorMessage &= " CM allows only 1 Customer condition! "
                            bNoErrorsSoFar = False
                        Else
                            If CustListNodes.Count = 1 Then
                                For Each CustNode In CustListNodes
                                    CustomerList = CustNode.InnerText
                                    Condition = New Copient.OfferCustomerGroup(OfferID, 0)
                                    If CustNode.Attributes("name") IsNot Nothing Then
                                        Added = MyCmOffer.AddCustomerCondition(Condition, ErrorMessage, CustomerList, CustNode.Attributes("name").InnerText)
                                    Else
                                        Added = MyCmOffer.AddCustomerCondition(Condition, ErrorMessage, CustomerList)
                                    End If
                                    ' store the created customer group in case we need to remove it later in processing due to offer failure  
                                    If Added Then CreatedCustomerGroups.Add(Condition.GetCustomerGroupID)
                                Next
                            Else
                                ErrorMessage &= " CM allows only 1 Customer condition! "
                                bNoErrorsSoFar = False
                            End If
                        End If
                    End If
                End If

                ' check to see if at least one customer condition exists, if not then create one.
                If bNoErrorsSoFar AndAlso (Not Added) Then
                    CustomerGroupIDs = MyCmOffer.GetConditionalCustomerGroups(ErrorMessage)
                    If CustomerGroupIDs.Length = 0 Then
                        Condition = New Copient.OfferCustomerGroup(OfferID, 0)
                        Added = MyCmOffer.AddCustomerCondition(Condition, ErrorMessage)
                        ' store the created customer group in case we need to remove it later in processing due to offer failure  
                        If Added Then CreatedCustomerGroups.Add(Condition.GetCustomerGroupID)
                    End If
                End If

            End If
        Catch ex As Exception
            ErrorMessage &= ex.ToString
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try

    End Sub

    Private Function GetBannerIDs(ByVal OfferXmlDoc As XmlDocument) As Long()
        Dim BannerIDs(-1) As Long
        Dim BannerNode As XmlNode = Nothing
        Dim NameNodes, ExtIdNodes, IdNodes As XmlNodeList
        Dim i As Integer

        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

        BannerNode = OfferXmlDoc.SelectSingleNode("//Offer/Banners")
        If BannerNode IsNot Nothing Then
            IdNodes = BannerNode.SelectNodes("//Offer/Banners/BannerID")
            If IdNodes IsNot Nothing AndAlso IdNodes.Count > 0 Then
                ReDim BannerIDs(IdNodes.Count - 1)
                For i = 0 To BannerIDs.GetUpperBound(0)
                    Long.TryParse(IdNodes.Item(i).InnerText, BannerIDs(i))
                Next
            Else
                ExtIdNodes = BannerNode.SelectNodes("//Offer/Banners/ExtBannerID")
                If (ExtIdNodes IsNot Nothing AndAlso ExtIdNodes.Count > 0) Then
                    BannerIDs = LookupBannerIDs(ExtIdNodes, "ExtBannerID")
                Else
                    NameNodes = BannerNode.SelectNodes("//Offer/Banners/Name")
                    If NameNodes IsNot Nothing AndAlso NameNodes.Count > 0 Then
                        BannerIDs = LookupBannerIDs(NameNodes, "Name")
                    End If
                End If
            End If
        End If

        Return BannerIDs
    End Function

    Private Function LookupBannerIDs(ByVal BannerNodes As XmlNodeList, ByVal ColumnName As String) As Long()
        Dim BannerIDs(-1) As Long
        Dim i As Integer
        Dim WhereClause As New StringBuilder()
        Dim TempValue As String = ""
        Dim dt As DataTable

        For i = 0 To BannerNodes.Count - 1
            TempValue = BannerNodes(i).InnerText.Trim
            If (TempValue <> "") Then
                If WhereClause.Length > 0 Then WhereClause.Append(",")
                WhereClause.Append("'")
                WhereClause.Append(TempValue)
                WhereClause.Append("'")
            End If
        Next
        If WhereClause.Length > 0 Then
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            MyCommon.QueryStr = "select BannerID from Banners with (NoLock) " &
                                "where Deleted=0 and " & ColumnName & " in (" & WhereClause.ToString & ");"
            dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            If dt.Rows.Count > 0 Then
                ReDim BannerIDs(dt.Rows.Count - 1)
                For i = 0 To dt.Rows.Count - 1
                    BannerIDs(i) = MyCommon.NZ(dt.Rows(i).Item("BannerID"), 0)
                Next
            End If
        End If

        Return BannerIDs
    End Function

    Private Function WriteOfferXmlResponse(ByVal OfferID As Long, ByVal ClientOfferID As String,
                                           ByVal ErrorCode As ERROR_CODES, ByVal ErrorMsg As String,
                                           ByVal ResponseType As RESPONSE_TYPES, ByVal ResponseFlag As Boolean, Optional ByVal OutputFileName As String = "") As String
        Dim sw As StringWriter = Nothing
        Dim Settings As XmlWriterSettings
        Dim Writer As XmlWriter
        Dim xmlStr As String = ""

        sw = New StringWriter()
        Settings = New XmlWriterSettings()
        Settings.Encoding = Encoding.UTF8
        Settings.Indent = True
        Settings.IndentChars = ControlChars.Tab

        Settings.NewLineChars = ControlChars.CrLf
        Settings.NewLineHandling = NewLineHandling.Replace

        Writer = XmlWriter.Create(sw, Settings)

        Writer.WriteStartDocument()
        Writer.WriteStartElement("ExternalOfferConnector")

        Writer.WriteStartElement("Offer")
        Writer.WriteAttributeString("id", ClientOfferID)
        Writer.WriteAttributeString("logixId", OfferID.ToString)
        Writer.WriteAttributeString("operation", ResponseType.ToString)
        If (Not String.IsNullOrEmpty(OutputFileName)) Then Writer.WriteAttributeString("clipfilename", OutputFileName)
        Writer.WriteAttributeString("success", ResponseFlag.ToString.ToLower)
        Writer.WriteEndElement() 'Offer

        If (ErrorMsg.Trim <> "") Then
            Writer.WriteStartElement("Error")
            Writer.WriteAttributeString("code", ErrorCode.ToString)
            Writer.WriteAttributeString("message", ErrorMsg)
            Writer.WriteEndElement() 'Error
        End If

        Writer.WriteEndElement() 'ExternalOfferConnector
        Writer.WriteEndDocument()

        Writer.Flush()
        Writer.Close()

        ' workaround for problem where encoding is always set to utf-16 no matter
        ' what you set for the encoding in the XMLWriterSettings.Encoding 
        xmlStr = sw.ToString
        If (xmlStr IsNot Nothing) Then

            xmlStr = xmlStr.Replace("encoding=""utf-16""", "encoding=""utf-8""")
        End If

        Return xmlStr
    End Function

    Private Function DoesOfferExist(ByVal ClientOfferID As String, ByVal ExtInterfaceID As Integer, ByRef LogixID As Long) As Boolean
        Dim OfferExists As Boolean = False
        Dim OfferEngineId As Long = -1

        OfferExists = DoesOfferExist(ClientOfferID, ExtInterfaceID, LogixID, OfferEngineId)

        Return OfferExists
    End Function

    Private Function InsertClientOfferId(ByVal ClientOfferID As String, ByVal ExtInterfaceID As Long, ByRef IsCreated As Boolean) As Boolean
        Dim IsClientOfferIdInserted As Boolean = True
        Try
            If (MyCommon.LRTadoConn.State = Data.ConnectionState.Closed) Then MyCommon.Open_LogixRT()
            MyCommon.QueryStr = "dbo.pt_ExtOfferID_Insert"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@ExtOfferID", SqlDbType.NVarChar, 20).Value = ClientOfferID
            MyCommon.LRTsp.Parameters.Add("@InboundCRMEngineID", SqlDbType.Int).Value = ExtInterfaceID
            MyCommon.LRTsp.ExecuteNonQuery()
            MyCommon.Close_LRTsp()
            IsCreated = True
        Catch ex As SqlException
            If ex.Number = 2627 Or ex.Message.StartsWith("Unique key violation") Then
                IsClientOfferIdInserted = False
            Else
                Throw ex
            End If
        Catch ex As Exception
            Throw ex
        End Try
        Return IsClientOfferIdInserted
    End Function
    Private Sub DeleteClientID(ByVal ClientOfferID As String, ByVal ExtInterfaceID As Long)
        If (MyCommon.LRTadoConn.State = Data.ConnectionState.Closed) Then MyCommon.Open_LogixRT()
        MyCommon.QueryStr = "DELETE FROM ExtOfferIDs WHERE ExtOfferID=@ExtOfferID AND InboundCRMEngineID=@InboundCRMEngineID"
        MyCommon.DBParameters.Add("@ExtOfferID", SqlDbType.NVarChar, 20).Value = ClientOfferID
        MyCommon.DBParameters.Add("@InboundCRMEngineID", SqlDbType.Int).Value = ExtInterfaceID
        MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
    End Sub

    Private Function DoesOfferExist(ByVal ClientOfferID As String, ByVal ExtInterfaceID As Integer, ByRef LogixID As Long, ByRef OfferEngineID As Long) As Boolean
        Dim OfferExists As Boolean = False
        Dim sQuery As String
        Dim dt As DataTable = Nothing

        Try
            If (MyCommon.LRTadoConn.State = Data.ConnectionState.Closed) Then MyCommon.Open_LogixRT()

            sQuery = "select IncentiveID as OfferID, EngineId from CPE_Incentives with (NoLock) where ClientOfferID = @ClientOfferID " &
                     "and InboundCRMEngineID = @InboundCRMEngineID and Deleted=0"

            If MyCommon.IsEngineInstalled(0) Then
                MyCommon.QueryStr = "select OfferID, EngineID from Offers with (NoLock) where ExtOfferID = @ExtOfferID " &
                                    "and InboundCRMEngineID = @InboundCRMEngineID and Deleted=0 " & " union " & sQuery
                MyCommon.DBParameters.Add("@ExtOfferID", SqlDbType.NVarChar).Value = ClientOfferID
            Else
                MyCommon.QueryStr = sQuery
            End If
            MyCommon.DBParameters.Add("@ClientOfferID", SqlDbType.NVarChar).Value = ClientOfferID
            MyCommon.DBParameters.Add("@InboundCRMEngineID", SqlDbType.Int).Value = ExtInterfaceID
            dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            If (dt.Rows.Count > 0) Then
                OfferExists = True
                LogixID = MyCommon.NZ(dt.Rows(0).Item("OfferID"), 0)
                OfferEngineID = MyCommon.NZ(dt.Rows(0).Item("EngineID"), 0)
            End If

        Catch ex As Exception
            Throw ex
        End Try
        Return OfferExists

    End Function

    Private Function CpeOfferUpdated(ByVal OfferXml As String, ByVal LogixID As Long, ByRef ErrorCode As ERROR_CODES, ByRef ErrorMsg As String) As Boolean
        Dim Updated As Boolean = False
        Dim xmlDoc As New XmlDocument
        Dim HeaderNode As XmlNode

        Try
            If (MyCommon.LRTadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixRT()

            xmlDoc.LoadXml(OfferXml)

            HeaderNode = xmlDoc.SelectSingleNode("//Offer/Header")
            If (HeaderNode IsNot Nothing) Then
                Updated = UpdateCpeHeader(HeaderNode, LogixID)
            End If

        Catch ex As Exception
            ErrorCode = ERROR_CODES.ERROR_OFFER_UPDATE_FAILED
            ErrorMsg = "Logix Offer: " & LogixID & " failed to update due to the following reasons: " & ex.ToString
            Updated = False
        End Try

        Return Updated
    End Function

    Private Function UpdateCpeHeader(ByVal HeaderNode As XmlNode, ByVal LogixID As Long) As Boolean
        Dim Updated As Boolean = False
        Dim SqlBuf As New StringBuilder()
        Dim ValTable As New Hashtable(15)
        Dim i As Integer = 0
        Dim enumerator As IDictionaryEnumerator
        Dim HeaderCols() As String = New String() {"ClientOfferID", "IncentiveName", "Description", "Priority",
                                                   "StartDate", "EndDate", "EligibilityStartDate", "EligibilityEndDate",
                                                   "TestingStartDate", "TestingEndDate", "Reporting", "DisabledOnCFW",
                                                   "AllowOptOut", "EmployeesOnly", "EmployeesExcluded", "DeferCalcToEOS",
                                                   "ExportToEDW"}
        Dim ColDataTypes() As System.TypeCode = New System.TypeCode() {TypeCode.String, TypeCode.String, TypeCode.String, TypeCode.Int32,
                                                   TypeCode.DateTime, TypeCode.DateTime, TypeCode.DateTime, TypeCode.DateTime,
                                                   TypeCode.DateTime, TypeCode.DateTime, TypeCode.Boolean, TypeCode.Boolean,
                                                   TypeCode.Boolean, TypeCode.Boolean, TypeCode.Boolean, TypeCode.Boolean,
                                                   TypeCode.Boolean}

        For i = 0 To HeaderCols.GetUpperBound(0)
            PutChildNodeValue(HeaderNode, HeaderCols(i), ColDataTypes(i), ValTable)
        Next

        If (ValTable.Count > 0) Then
            SqlBuf.Append("update CPE_Incentives set ")

            enumerator = ValTable.GetEnumerator
            While enumerator.MoveNext()
                SqlBuf.Append(enumerator.Key.ToString)
                SqlBuf.Append("=")
                SqlBuf.Append(enumerator.Value.ToString)
                SqlBuf.Append(", ")
            End While
            SqlBuf.Append("LastUpdate=getdate(), StatusFlag=1 where IncentiveID = @IncentiveID and Deleted=0")

            If (MyCommon.LRTadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixRT()

            MyCommon.QueryStr = SqlBuf.ToString
            MyCommon.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = LogixID
            MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)

            Updated = (MyCommon.RowsAffected > 0)
        Else
            ' when there's nothing changed then the update is technically successful
            Updated = True
        End If

        Return Updated
    End Function

    Private Function CpeRemoveOffer(ByVal ClientOfferID As String, ByVal ExtInterfaceID As Integer, ByRef LogixID As Long, ByRef ErrorCode As ERROR_CODES, ByRef ErrorMsg As String) As Boolean
        Dim Removed As Boolean = False
        Dim dt As DataTable

        Try
            If (MyCommon.LRTadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixRT()

            ' ensure that the offer exists
            If LogixID <= 0 Then
                MyCommon.QueryStr = "select top 1 IncentiveID from CPE_Incentives with (NoLock) where ClientOfferID = @ClientOfferID and InboundCRMEngineID = @InboundCRMEngineID and Deleted=0"
                MyCommon.DBParameters.Add("@ClientOfferID", SqlDbType.NVarChar).Value = ClientOfferID.ConvertBlankIfNothing
                MyCommon.DBParameters.Add("@InboundCRMEngineID", SqlDbType.Int).Value = ExtInterfaceID
            Else
                MyCommon.QueryStr = "select top 1 IncentiveID from CPE_Incentives with (NoLock) where IncentiveID = @IncentiveID and InboundCRMEngineID = @InboundCRMEngineID and Deleted=0"
                MyCommon.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = LogixID
                MyCommon.DBParameters.Add("@InboundCRMEngineID", SqlDbType.Int).Value = ExtInterfaceID
            End If
            dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            If dt.Rows.Count > 0 Then
                LogixID = MyCommon.NZ(dt.Rows(0).Item("IncentiveID"), -1)
            End If

            If LogixID > 0 Then
                MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=2,Deleted=1,UpdateLevel=UpdateLevel+1 where IncentiveID = @IncentiveID"
                MyCommon.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = LogixID
                MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
                Removed = (MyCommon.RowsAffected > 0)

                ' Mark the shadow table offer as deleted as well.
                MyCommon.QueryStr = "update CPE_ST_Incentives with (RowLock) set Deleted=1, UpdateLevel = (select UpdateLevel from CPE_Incentives with (NoLock) where IncentiveID = @IncentiveID) where IncentiveID = @IncentiveID"
                MyCommon.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = LogixID
                MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)

                'remove the banners assigned to this offer
                If (MyCommon.Fetch_SystemOption(66) = "1") Then
                    MyCommon.QueryStr = "delete from BannerOffers with (RowLock) where OfferID = @OfferID"
                    MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = LogixID
                    MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
                End If

            Else
                Removed = False
                ErrorCode = ERROR_CODES.ERROR_OFFER_DOES_NOT_EXIST
                ErrorMsg = "Offer " & ClientOfferID & " does not exist in Logix"
            End If

        Catch ex As Exception
            ErrorCode = ERROR_CODES.ERROR_OFFER_REMOVE_FAILED
            ErrorMsg = "Removal of Offer " & ClientOfferID & " failed due to " & ex.ToString
            Removed = False
        End Try

        Return Removed
    End Function

    Private Function CmRemoveOffer(ByVal ClientOfferID As String, ByVal ExtInterfaceID As Integer, ByRef LogixID As Long, ByRef ErrorCode As ERROR_CODES, ByRef ErrorMsg As String) As Boolean
        Dim Removed As Boolean = False
        Dim dt As DataTable

        Try
            If (MyCommon.LRTadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixRT()

            ' ensure that the offer exists
            If LogixID <= 0 Then
                MyCommon.QueryStr = "select top 1 OfferID from Offers with (NoLock) where ExtOfferID = @ExtOfferID and InboundCRMEngineID = @InboundCRMEngineID and Deleted=0"
                MyCommon.DBParameters.Add("@ExtOfferID", SqlDbType.NVarChar).Value = ClientOfferID.ConvertBlankIfNothing
                MyCommon.DBParameters.Add("@InboundCRMEngineID", SqlDbType.Int).Value = ExtInterfaceID
            Else
                MyCommon.QueryStr = "select top 1 OfferID from Offers with (NoLock) where offerID = @OfferID and InboundCRMEngineID = @InboundCRMEngineID and Deleted=0"
                MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = LogixID
                MyCommon.DBParameters.Add("@InboundCRMEngineID", SqlDbType.Int).Value = ExtInterfaceID
            End If
            dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            If dt.Rows.Count > 0 Then
                LogixID = MyCommon.NZ(dt.Rows(0).Item("OfferID"), -1)
            End If

            If LogixID > 0 Then
                MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=2,Deleted=1,UpdateLevel=UpdateLevel+1 where OfferID = @OfferID"
                MyCommon.DBParameters.Add("@OfferID", SqlDbType.Int).Value = LogixID
                MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
                Removed = (MyCommon.RowsAffected > 0)

                'remove the banners assigned to this offer
                If (MyCommon.Fetch_SystemOption(66) = "1") Then
                    MyCommon.QueryStr = "delete from BannerOffers with (RowLock) where OfferID = @OfferID"
                    MyCommon.DBParameters.Add("@OfferID", SqlDbType.Int).Value = LogixID
                    MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
                End If

            Else
                Removed = False
                ErrorCode = ERROR_CODES.ERROR_OFFER_DOES_NOT_EXIST
                ErrorMsg = "Offer " & ClientOfferID & " does not exist in Logix"
            End If

        Catch ex As Exception
            ErrorCode = ERROR_CODES.ERROR_OFFER_REMOVE_FAILED
            ErrorMsg = "Removal of Offer " & ClientOfferID & " failed due to " & ex.ToString
            Removed = False
        End Try

        Return Removed
    End Function

    Private Sub RemoveCreatedGroups()
        Dim i As Integer

        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

        If CreatedCustomerGroups.Count > 0 Then
            For i = 0 To CreatedCustomerGroups.Count - 1
                If Long.Parse(CreatedCustomerGroups.Item(i)) > 2 Then
                    MyCommon.QueryStr = "dbo.pt_CustomerGroups_Delete"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = Long.Parse(CreatedCustomerGroups.Item(i))
                    MyCommon.LRTsp.ExecuteNonQuery()
                    MyCommon.Close_LRTsp()
                End If
            Next
        End If

        If CreatedProductGroups.Count > 0 Then
            For i = 0 To CreatedProductGroups.Count - 1
                If Long.Parse(CreatedProductGroups.Item(i)) > 1 Then
                    MyCommon.QueryStr = "dbo.pt_ProductGroups_Delete"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt).Value = Long.Parse(CreatedProductGroups.Item(i))
                    MyCommon.LRTsp.ExecuteNonQuery()
                    MyCommon.Close_LRTsp()
                End If
            Next
        End If

    End Sub

    Private Sub EnsureCustomerCondition(ByVal OfferID As Long)
        Dim dt As DataTable
        Dim ROID As Long
        Dim NewCg As Long

        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()


        MyCommon.QueryStr = "select IncentiveCustomerID, CustomerGroupID from CPE_IncentiveCustomerGroups ICG with (NoLock) " &
                            "inner join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID=ICG.RewardOptionID " &
                            "where ICG.Deleted=0 and ICG.ExcludedUsers =0 and RO.Deleted=0 and RO.IncentiveID = @IncentiveID"
        MyCommon.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = OfferID
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)

        ' create a blank customer group if one doesn't already exist for this offer
        If (dt.Rows.Count = 0) Then
            MyCommon.QueryStr = "dbo.pt_CustomerGroups_Insert"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 200).Value = "Customer group for " & OfferID
            MyCommon.LRTsp.Parameters.Add("@CustomerGroupID", SqlDbType.BigInt).Direction = ParameterDirection.Output
            MyCommon.LRTsp.Parameters.Add("@CAMCustomerGroup", SqlDbType.Bit).Value = 0
            MyCommon.LRTsp.ExecuteNonQuery()
            NewCg = MyCommon.LRTsp.Parameters("@CustomerGroupID").Value
            MyCommon.Close_LRTsp()

            If (NewCg > 0) Then
                MyCommon.Activity_Log(4, NewCg, 1, Copient.PhraseLib.Lookup("history.cgroup-create", 1))

                MyCommon.QueryStr = "select RewardOptionID from CPE_RewardOptions RO where Deleted=0 and IncentiveID = @IncentiveID"
                MyCommon.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = OfferID
                dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                If (dt.Rows.Count > 0) Then
                    ROID = MyCommon.NZ(dt.Rows(0).Item("RewardOptionID"), -1)
                    ' add that group to the offer
                    MyCommon.QueryStr = "insert into CPE_IncentiveCustomerGroups with (RowLock) (RewardOptionID,CustomerGroupID,ExcludedUsers,Deleted,LastUpdate,RequiredFromTemplate,TCRMAStatusFlag) " &
                                  "values(@RewardOptionID, @CustomerGroupID, 0, 0,getdate(), 0, 3)"
                    MyCommon.DBParameters.Add("@RewardOptionID", SqlDbType.BigInt).Value = ROID
                    MyCommon.DBParameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = NewCg
                    MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)

                    If (MyCommon.RowsAffected > 0) Then
                        MyCommon.Activity_Log(3, OfferID, 1, "Added customer condition from web service")
                    End If

                End If
            End If

        End If
    End Sub

    'Returns -1 when CardTypeID is invalid
    'Returns -2 when (Card, CardTypeID) is invalid
    'Returns 0 when not found
    'Returns >0 when CreateIfNotFound=True and there are no exceptions.
    Private Function GetCustomerPK(ByVal ExtCardID As String, ByVal CardTypeID As Integer, ByVal CreateIfNotFound As Boolean) As Long
        Dim CustomerPK As Long = 0
        Dim dt As DataTable

        If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

        If IsValidCardType(CardTypeID) Then
            If IsValidCard(CardTypeID, ExtCardID) Then
                ExtCardID = MyCommon.Pad_ExtCardID(MyCommon.Parse_Quotes(ExtCardID), CardTypeID)
                If CreateIfNotFound Then

                    Dim CustomerTypeID = CardTypeID

                    'AMS-13772
                    'We only set CustTypeID for Customer card, Household card and CAM card.
                    'For all other cards the type must be 0 since 0 = individual card, 1 = household card, 2 = CAM card. 
                    If CardTypeID > 2 Then
                        CustomerTypeID = 0
                    End If

                    MyCommon.QueryStr = "dbo.pa_EOC_GetOrCreateCustomer"
                    MyCommon.Open_LXSsp()
                    MyCommon.LXSsp.Parameters.Add("@ExtCardID", SqlDbType.NVarChar, 400).Value = MyCryptLib.SQL_StringEncrypt(ExtCardID, True)
                    MyCommon.LXSsp.Parameters.Add("@CardTypeID", SqlDbType.Int).Value = CardTypeID
                    MyCommon.LXSsp.Parameters.Add("@CustomerTypeID", SqlDbType.Int).Value = CustomerTypeID
                    MyCommon.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Direction = ParameterDirection.Output
                    MyCommon.LXSsp.Parameters.Add("@ExtCardIDOriginal", SqlDbType.NVarChar, 400).Value = MyCryptLib.SQL_StringEncrypt(ExtCardID)
                    MyCommon.LXSsp.ExecuteNonQuery()
                    CustomerPK = MyCommon.LXSsp.Parameters("@CustomerPK").Value
                    MyCommon.Close_LXSsp()
                Else
                    ' lookup the pk for this card number
                    MyCommon.QueryStr = "select CustomerPK from CardIDs with (NoLock) where CardTypeID = @CardTypeID and ExtCardID = @ExtCardID"
                    MyCommon.DBParameters.Add("@CardTypeID", SqlDbType.Int).Value = CardTypeID
                    MyCommon.DBParameters.Add("@ExtCardID", SqlDbType.NVarChar).Value = MyCryptLib.SQL_StringEncrypt(ExtCardID.ConvertBlankIfNothing, True)
                    dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)
                    If dt.Rows.Count > 0 Then
                        CustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)
                    End If
                End If
            Else
                Return -2
            End If
        Else
            Return -1
        End If

        Return CustomerPK
    End Function

    Private Function GetProductID(ByVal ClientProductID As String) As Long
        Dim ProductID As Long
        Dim dt As DataTable

        MyCommon.QueryStr = "select ProductID from Products with (NoLock) where ExtProductID = @ExtProductID"
        MyCommon.DBParameters.Add("@ExtProductID", SqlDbType.NVarChar).Value = ClientProductID.ConvertBlankIfNothing
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        If (dt.Rows.Count > 0) Then
            ProductID = MyCommon.NZ(dt.Rows(0).Item("ProductID"), 0)
        End If

        Return ProductID
    End Function

    Private Sub SatisfyCustomerCondition(ByVal CustomerIDs As String, ByVal LogixID As Long,
                                       ByVal EngineID As Long, ByRef ErrorCode As ERROR_CODES,
                                       ByVal ErrorMsg As String)
        Dim dt As DataTable
        Dim row As DataRow
        Dim CustomerGroupID As Long
        Dim ExcludedGroup As Boolean
        Dim Success As Boolean
        Dim MyCPEOffer As New Copient.CPEOffer
        Dim ErrorMessage As String = ""
        Dim OperationType As Copient.CPEOffer.OPERATION_TYPES
        Dim CPEQueueData As New Copient.CPEOffer.QUEUE_DATA
        Dim CMQueueData As New Copient.CMOffer.QUEUE_DATA

        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
        If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

        ' find all the user-defined customer groups attached to this offer and remove the customer from the excluded groups
        ' and add the customer to all the included groups
        MyCommon.QueryStr = GetOfferCustomerGroupSQL(EngineID)
        MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = LogixID
        If MyCommon.QueryStr <> "" Then
            dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            If (dt.Rows.Count > 0) Then
                For Each row In dt.Rows
                    Success = False
                    CustomerGroupID = MyCommon.NZ(row.Item("CustomerGroupID"), 0)
                    ExcludedGroup = MyCommon.NZ(row.Item("ExcludedUsers"), False)

                    If ExcludedGroup Then
                        OperationType = Copient.CPEOffer.OPERATION_TYPES.REMOVE_FROM_GROUP
                    Else
                        OperationType = Copient.CPEOffer.OPERATION_TYPES.ADD_TO_GROUP
                    End If

                    Select Case EngineID
                        Case Copient.CommonInc.InstalledEngines.CM
                            With CMQueueData
                                .CustomerGroupID = CustomerGroupID
                                .ExtInterfaceID = ExtInterfaceID
                                .FullReplace = False
                                .OperationType = OperationType
                            End With
                            Success = MyCmOffer.PopulateCustomerGroup(CMQueueData, CustomerIDs, ErrorMessage)
                        Case Else
                            With CPEQueueData
                                .CustomerGroupID = CustomerGroupID
                                .ExtInterfaceID = ExtInterfaceID
                                .FullReplace = False
                                .OperationType = OperationType
                            End With
                            Success = MyCPEOffer.PopulateCustomerGroup(CPEQueueData, CustomerIDs, ErrorMessage)
                    End Select

                    If Not Success Then
                        ErrorCode = ERROR_CODES.ERROR_ADD_CUSTOMERS_FAILED
                        If ExcludedGroup Then
                            ErrorMsg &= "Failed to remove customers from the offer's exclusion group ID " & CustomerGroupID
                        Else
                            ErrorMsg &= "Failed to add customer to the offer's conditional customer group ID " & CustomerGroupID
                        End If
                    End If
                Next
            End If
        End If

    End Sub

    Private Sub CPERemoveCustomersFromCondition(ByVal CustomerIDs As String, ByVal LogixID As Long,
                                              ByRef ErrorCode As ERROR_CODES, ByRef ErrorMsg As String)
        Dim dt As DataTable
        Dim row As DataRow
        Dim rows() As DataRow
        Dim CgID, ExcludedCgID As Long
        Dim MyCPEOffer As New Copient.CPEOffer
        Dim QueueData As New Copient.CPEOffer.QUEUE_DATA
        Dim Success As Boolean
        Dim ErrorMessage As String = ""
        Dim ROID As Long

        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
        If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

        ' find all the customer groups attached to this offer
        MyCommon.QueryStr = "Select ICG.CustomerGroupID, ICG.RewardOptionID, CG.AnyCardholder, ICG.ExcludedUsers from CPE_IncentiveCustomerGroups ICG with (NoLock) " &
                            "inner join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID = ICG.RewardOptionID " &
                            "inner join CustomerGroups CG with (NoLock) on CG.CustomerGroupID = ICG.CustomerGroupID " &
                            "where ICG.Deleted=0 and CG.Deleted=0 and RO.Deleted=0 and RO.IncentiveID = @IncentiveID " &
                            "  and CG.AnyCustomer<>1 and CG.NewCardholders<>1"
        MyCommon.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = LogixID
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)

        If (dt.Rows.Count > 0) Then
            ROID = MyCommon.NZ(dt.Rows(0).Item("RewardOptionID"), 0)

            ' determine if the offer is targeted to any cardholders
            rows = dt.Select("AnyCardholder=1")
            If rows.Length > 0 Then
                ' determine if there is an exclusion group, if not then create one
                rows = dt.Select("ExcludedUsers=1")
                If rows.Length = 0 Then
                    ' add a exclusion customer group to this offer
                    MyCommon.QueryStr = "dbo.pt_CustomerGroups_Insert"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 200).Value = "Excluded Group for Offer " & LogixID.ToString
                    MyCommon.LRTsp.Parameters.Add("@CustomerGroupID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                    MyCommon.LRTsp.Parameters.Add("@CAMCustomerGroup", SqlDbType.Bit).Value = 0
                    MyCommon.LRTsp.ExecuteNonQuery()
                    ExcludedCgID = MyCommon.LRTsp.Parameters("@CustomerGroupID").Value
                    MyCommon.Close_LRTsp()

                    ' assign new excluded group to this offer
                    MyCommon.QueryStr = "insert into CPE_IncentiveCustomerGroups with (RowLock) (RewardOptionID,CustomerGroupID,ExcludedUsers,Deleted,LastUpdate,RequiredFromTemplate, TCRMAStatusFlag) " &
                                        "values(@RewardOptionID, @CustomerGroupID, 1, 0, getdate(), 0, 3)"
                    MyCommon.DBParameters.Add("@RewardOptionID", SqlDbType.Int).Value = ROID
                    MyCommon.DBParameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = ExcludedCgID
                    MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)

                    ' set the offer in a modified state
                    MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID = @LastUpdatedByAdminID where IncentiveID = @IncentiveID"
                    MyCommon.DBParameters.Add("@LastUpdatedByAdminID", SqlDbType.Int).Value = lUserId
                    MyCommon.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = LogixID
                    MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
                Else
                    ExcludedCgID = MyCommon.NZ(rows(0).Item("CustomerGroupID"), 0)
                End If

                If ExcludedCgID > 0 Then
                    ' add the customers to the exclusion group
                    With QueueData
                        .CustomerGroupID = ExcludedCgID
                        .ExtInterfaceID = ExtInterfaceID
                        .FullReplace = False
                        .OperationType = Copient.CPEOffer.OPERATION_TYPES.ADD_TO_GROUP
                    End With
                    Success = MyCPEOffer.PopulateCustomerGroup(QueueData, CustomerIDs, ErrorMessage)
                End If
            Else
                ' remove customers from the targeted customer group(s)
                rows = dt.Select("AnyCardholder<>1 and ExcludedUsers<>1")
                For Each row In rows
                    CgID = MyCommon.NZ(row.Item("CustomerGroupID"), 0)
                    With QueueData
                        .CustomerGroupID = CgID
                        .ExtInterfaceID = ExtInterfaceID
                        .FullReplace = False
                        .OperationType = Copient.CPEOffer.OPERATION_TYPES.REMOVE_FROM_GROUP
                    End With
                    Success = MyCPEOffer.PopulateCustomerGroup(QueueData, CustomerIDs, ErrorMessage)
                Next
            End If
        End If

    End Sub

    Private Sub CMRemoveCustomersFromCondition(ByVal CustomerIDs As String, ByVal LogixID As Long,
                                               ByRef ErrorCode As ERROR_CODES, ByRef ErrorMsg As String)
        Dim dt As DataTable
        Dim row As DataRow
        Dim rows() As DataRow
        Dim CgID, ExcludedCgID As Long
        Dim Success As Boolean
        Dim ErrorMessage As String = ""
        Dim ConditionID As Long
        Dim QueueData As New Copient.CMOffer.QUEUE_DATA

        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
        If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

        ' find all the customer groups attached to this offer
        MyCommon.QueryStr = "Select OC.LinkID as IncludedID, CG.AnyCardholder, OC.ExcludedID, OC.ConditionID " &
                            "from OfferConditions OC with (NoLock) " &
                            "inner join CustomerGroups CG with (NoLock) on CG.CustomerGroupID=OC.LinkID " &
                            "where OC.Deleted=0 And OC.ConditionTypeId=1 And CG.Deleted=0 And OC.OfferID = @offerid " &
                            "and CG.AnyCustomer=0 and CG.NewCardholders=0 "
        MyCommon.DBParameters.Add("@offerid", SqlDbType.BigInt).Value = LogixID
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        If (dt.Rows.Count > 0) Then
            ' determine if the offer is targeted to any cardholders
            rows = dt.Select("AnyCardholder=1")
            If rows.Length > 0 Then
                ExcludedCgID = MyCommon.NZ(rows(0).Item("ExcludedID"), 0)
                ConditionID = MyCommon.NZ(rows(0).Item("ConditionID"), 0)

                ' determine if there is an exclusion group, if not then create one
                If ExcludedCgID = 0 Then
                    ' add a exclusion customer group to this offer
                    MyCommon.QueryStr = "dbo.pt_CustomerGroups_Insert"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 200).Value = "Excluded Group for Offer " & LogixID.ToString
                    MyCommon.LRTsp.Parameters.Add("@CustomerGroupID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                    MyCommon.LRTsp.Parameters.Add("@CAMCustomerGroup", SqlDbType.Bit).Value = 0
                    MyCommon.LRTsp.ExecuteNonQuery()
                    ExcludedCgID = MyCommon.LRTsp.Parameters("@CustomerGroupID").Value
                    MyCommon.Close_LRTsp()
                Else
                    ExcludedCgID = MyCommon.NZ(rows(0).Item("CustomerGroupID"), 0)
                End If

                If ExcludedCgID > 0 Then
                    ' assign new excluded group to this offer
                    MyCommon.QueryStr = "update OfferConditions with (RowLock) " &
                                        "set ExcludedID = @ExcludedID where ConditionID = @ConditionID "
                    MyCommon.DBParameters.Add("@ExcludedID", SqlDbType.BigInt).Value = ExcludedCgID
                    MyCommon.DBParameters.Add("@ConditionID", SqlDbType.BigInt).Value = ConditionID

                    ' set the offer in a modified state
                    MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID = @LastUpdatedByAdminID where offerid = @offerid "
                    MyCommon.DBParameters.Add("@LastUpdatedByAdminID", SqlDbType.Int).Value = lUserId
                    MyCommon.DBParameters.Add("@offerid", SqlDbType.BigInt).Value = LogixID
                    MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)

                    ' add the customers to the exclusion group
                    With QueueData
                        .CustomerGroupID = ExcludedCgID
                        .ExtInterfaceID = ExtInterfaceID
                        .FullReplace = False
                        .OperationType = Copient.CMOffer.OPERATION_TYPES.ADD_TO_GROUP
                    End With
                    Success = MyCmOffer.PopulateCustomerGroup(QueueData, CustomerIDs, ErrorMessage)
                End If
            Else
                ' remove customers from the targeted customer group(s)
                rows = dt.Select("AnyCardholder<>1")
                For Each row In rows
                    CgID = MyCommon.NZ(rows(0).Item("IncludedID"), 0)
                    With QueueData
                        .CustomerGroupID = CgID
                        .ExtInterfaceID = ExtInterfaceID
                        .FullReplace = False
                        .OperationType = Copient.CMOffer.OPERATION_TYPES.REMOVE_FROM_GROUP
                    End With
                    Success = MyCmOffer.PopulateCustomerGroup(QueueData, CustomerIDs, ErrorMessage)
                Next
            End If
        End If

    End Sub

    Private Sub CmRemoveProductFromOffer(ByVal ClientProductID As String, ByVal ClientOfferId As String, ByVal LogixID As Long, ByRef ErrorCode As ERROR_CODES, ByRef ErrorMsg As String)
        Dim dt, dt2 As DataTable
        Dim row As DataRow
        Dim IncludedPgID, ExcludedPgID As Long
        Dim ID As Long
        Dim AnyProduct As Boolean
        Dim Updated, Added, bReward As Boolean
        Dim ProductID As Long
        Dim ProductDesc As String = ""
        Dim OutputStatus As Integer


        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
        If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

        ClientProductID = HandleProductIdPadding(ClientProductID)
        MyCommon.QueryStr = "select ProductID, Description from Products with (NoLock) where ExtProductID = @ExtProductID "
        MyCommon.DBParameters.Add("@ExtProductID", SqlDbType.NVarChar).Value = ClientProductID.ConvertBlankIfNothing
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        If (dt.Rows.Count > 0) Then
            ProductID = MyCommon.NZ(dt.Rows(0).Item("ProductID"), 0)
            ProductDesc = MyCommon.NZ(dt.Rows(0).Item("Description"), "")

            ' find all the user defined product groups attached to conditions & rewards for this offer
            MyCommon.QueryStr = "select OC.LinkID as IncludedID, PG.AnyProduct, OC.ExcludedID, OC.ConditionID as ID, 0 as Reward " &
                                "from OfferConditions OC with (NoLock) " &
                                "inner join ProductGroups PG with (NoLock) on PG.ProductGroupId = OC.LinkId " &
                                "where PG.Deleted=0 and OC.Deleted=0 and OC.ConditionTypeId=2 and OC.OfferID = @OfferID" &
                                " union " &
                                "select ORW.ProductGroupId as IncludedID, PG.AnyProduct, ORW.ExcludedProdGroupId as ExcludedID, ORW.RewardID as ID, 1 as Reward " &
                                "from OfferRewards ORW with (NoLock) " &
                                "inner join ProductGroups PG with (NoLock) on PG.ProductGroupId = ORW.ProductGroupId " &
                                "where PG.Deleted=0 and ORW.Deleted=0 and ORW.OfferID = @OfferID" &
                                " order by Reward, ID"
            MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = LogixID
            dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            If (dt.Rows.Count > 0) Then
                For Each row In dt.Rows
                    Updated = False
                    Added = False
                    IncludedPgID = MyCommon.NZ(row.Item("IncludedID"), 0)
                    AnyProduct = MyCommon.NZ(row.Item("AnyProduct"), False)
                    ExcludedPgID = MyCommon.NZ(row.Item("ExcludedID"), 0)
                    ID = MyCommon.NZ(row.Item("ID"), 0)
                    bReward = MyCommon.NZ(row.Item("Reward"), 0)

                    If IncludedPgID > 0 And Not AnyProduct Then
                        MyCommon.QueryStr = "Select PKID from ProdGroupItems where ProductGroupID=@ProductGroupID" &
                                            "  and ProductID = @ProductID and Deleted=0"
                        MyCommon.DBParameters.Add("@ProductID", SqlDbType.BigInt).Value = ProductID
                        MyCommon.DBParameters.Add("@ProductGroupID", SqlDbType.BigInt).Value = IncludedPgID
                        dt2 = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                        ' remove the product from the included group
                        If (dt2.Rows.Count > 0) Then
                            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
                            MyCommon.QueryStr = "dbo.pt_ProdGroupItems_DeleteItem"
                            MyCommon.Open_LRTsp()
                            MyCommon.LRTsp.Parameters.Add("@ExtProductID", SqlDbType.NVarChar, 120).Value = ClientProductID
                            MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt).Value = IncludedPgID
                            MyCommon.LRTsp.Parameters.Add("@ProductTypeID", SqlDbType.Int).Value = 1 ' UPC
                            MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
                            MyCommon.LRTsp.ExecuteNonQuery()
                            OutputStatus = MyCommon.LRTsp.Parameters("@Status").Value
                            MyCommon.Close_LRTsp()
                            If (OutputStatus <> 0) Then
                                ErrorCode = ERROR_CODES.ERROR_REMOVE_PRODUCT_FAILED
                                If ErrorMsg <> "" Then ErrorMsg &= "; "
                                ErrorMsg &= "Error Encountered while attempting to remove ProductID " & ClientProductID & " from offer " & ClientOfferId
                            Else
                                MyCommon.QueryStr = "update ProductGroups with (RowLock) set updatelevel=updatelevel+1,LastUpdate=getdate() where ProductGroupID = @ProductGroupID"
                                MyCommon.DBParameters.Add("@ProductGroupID", SqlDbType.BigInt).Value = IncludedPgID
                                MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
                                MyCommon.Activity_Log(5, IncludedPgID, 1, Copient.PhraseLib.Lookup("history.pgroup-remove", 1) & " " & ClientProductID)
                            End If
                        End If
                    Else
                        If (ExcludedPgID = 0) Then
                            ' create new product group to be used as an exclusion group for this offer
                            MyCommon.QueryStr = "dbo.pt_ProductGroups_Insert"
                            MyCommon.Open_LRTsp()
                            MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 200).Value = "Excluded Product Group for Offer " & LogixID
                            MyCommon.LRTsp.Parameters.Add("@AnyProduct", SqlDbType.Bit).Value = 0
                            MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                            MyCommon.LRTsp.ExecuteNonQuery()
                            ExcludedPgID = MyCommon.LRTsp.Parameters("@ProductGroupID").Value
                            MyCommon.Close_LRTsp()

                            If (ExcludedPgID > 0) Then
                                MyCommon.Activity_Log(5, ExcludedPgID, 1, Copient.PhraseLib.Lookup("history.pgroup-create", 1))

                                If bReward Then
                                    MyCommon.QueryStr = "update OfferRewards with (RowLock) " &
                                                        "set ExcludedProdGroupId = @ExcludedProdGroupId where RewardID = @RewardID"
                                    MyCommon.DBParameters.Add("@ExcludedProdGroupId", SqlDbType.BigInt).Value = ExcludedPgID
                                    MyCommon.DBParameters.Add("@RewardID", SqlDbType.BigInt).Value = ID
                                Else
                                    MyCommon.QueryStr = "update OfferConditions with (RowLock) " &
                                                        "set ExcludedID = @ExcludedId where ConditionID = @ConditionID"
                                    MyCommon.DBParameters.Add("@ExcludedId", SqlDbType.BigInt).Value = ExcludedPgID
                                    MyCommon.DBParameters.Add("@ConditionID", SqlDbType.BigInt).Value = ID
                                End If
                                MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
                                If (MyCommon.RowsAffected > 0) Then
                                    MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID = @LastUpdatedByAdminID where offerid=@offerid"
                                    MyCommon.DBParameters.Add("@LastUpdatedByAdminID", SqlDbType.Int).Value = lUserId
                                    MyCommon.DBParameters.Add("@offerid", SqlDbType.BigInt).Value = LogixID
                                    MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
                                    If (MyCommon.RowsAffected > 0) Then
                                        If bReward Then
                                            MyCommon.Activity_Log(3, LogixID, 1, "Updated reward with excluded product group " & ExcludedPgID)
                                        Else
                                            MyCommon.Activity_Log(3, LogixID, 1, "Updated condition with excluded product group " & ExcludedPgID)
                                        End If
                                    Else
                                        ErrorCode = ERROR_CODES.ERROR_REMOVE_PRODUCT_FAILED
                                        ErrorMsg = "Error encountered while attempting to update status for offer " & ClientOfferId
                                    End If
                                Else
                                    ErrorCode = ERROR_CODES.ERROR_REMOVE_PRODUCT_FAILED
                                    If bReward Then
                                        ErrorMsg = "Error encountered while attempting to add an exclusion group to reward for offer " & ClientOfferId
                                    Else
                                        ErrorMsg = "Error encountered while attempting to add an exclusion group to condition for offer " & ClientOfferId
                                    End If
                                End If
                            Else
                                ErrorCode = ERROR_CODES.ERROR_REMOVE_PRODUCT_FAILED
                                ErrorMsg = "Error encountered while attempting to create an exclusion group for offer " & ClientOfferId
                            End If
                        End If

                        ' if there is an exclusion group for this condition then add the product to it
                        If (ExcludedPgID > 0) Then
                            ' see if the product is already a member of the exclusion group, if not then add the product to it
                            MyCommon.QueryStr = "Select PKID from ProdGroupItems where ProductGroupID = @ProductGroupID" &
                                                "  and ProductID = @ProductID and Deleted=0"
                            MyCommon.DBParameters.Add("@ProductGroupID", SqlDbType.BigInt).Value = ExcludedPgID
                            MyCommon.DBParameters.Add("@ProductID", SqlDbType.BigInt).Value = ProductID
                            dt2 = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                            If (dt2.Rows.Count = 0) Then
                                MyCommon.QueryStr = "insert into ProdGroupItems with (RowLock) " &
                                                    "(ProductGroupID, ProductID, Manual, Deleted, CMOAStatusFlag, TCRMAStatusFlag, CPEStatusFlag, LastUpdate) " &
                                                    "values " &
                                                    "(@ProductGroupID, @ProductID, 1, 0, 2, 2,2,getdate())"
                                MyCommon.DBParameters.Add("@ProductGroupID", SqlDbType.BigInt).Value = ExcludedPgID
                                MyCommon.DBParameters.Add("@ProductID", SqlDbType.BigInt).Value = ProductID
                                MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
                                If (MyCommon.RowsAffected > 0) Then
                                    MyCommon.QueryStr = "Update ProductGroups with (RowLock) set updatelevel=updatelevel+1,LastUpdate=getdate() where ProductGroupID = @ProductGroupID"
                                    MyCommon.DBParameters.Add("@ProductGroupID", SqlDbType.BigInt).Value = ExcludedPgID
                                    MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
                                    MyCommon.Activity_Log(5, ExcludedPgID, 1, Copient.PhraseLib.Lookup("history.pgroup-add", 1) & " " & ClientProductID)
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        Else
            ErrorCode = ERROR_CODES.ERROR_PRODUCT_DOES_NOT_EXIST
            ErrorMsg = "Client Product " & ClientProductID & " does not exist in Logix"
        End If


    End Sub

    Private Sub HandleLocationsAndTerminals(ByVal LogixID As Long)
        Dim dt As DataTable
        Dim TerminalTypeID As Long
        Dim LocationGroupID As Long

        If (MyCommon.LRTadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixRT()

        ' add Any Location as a default to the offer location
        MyCommon.QueryStr = "select LocationGroupID from LocationGroups LG with (NoLock) Where AllLocations=1 and Deleted=0"
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        If (dt.Rows.Count > 0) Then
            LocationGroupID = MyCommon.NZ(dt.Rows(0).Item("LocationGroupID"), 0)
            If (LocationGroupID > 0) Then
                AddOfferLocation(LogixID, LocationGroupID, False)
            End If
        End If

        ' add Any Terminal as a default for the terminal type
        MyCommon.QueryStr = "select TerminalTypeID from TerminalTypes TT with (NoLock) Where AnyTerminal=1 and Deleted=0 and (EngineID=2 or EngineID=9)"
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        If (dt.Rows.Count > 0) Then
            TerminalTypeID = MyCommon.NZ(dt.Rows(0).Item("TerminalTypeID"), 0)
            If (TerminalTypeID > 0) Then
                AddOfferTerminal(LogixID, TerminalTypeID, False)
            End If
        End If
    End Sub

    Private Sub HandleBannerAssignments(ByVal LogixID As Long)
        Dim dt As DataTable
        Dim TempStr As String = ""
        Dim DefaultBannerID As Integer = 0

        If (MyCommon.LRTadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixRT()

        If (MyCommon.Fetch_SystemOption(66) = "1") Then
            TempStr = MyCommon.Fetch_CPE_SystemOption(61)
            Integer.TryParse(TempStr, DefaultBannerID)

            If (DefaultBannerID > 0) Then
                MyCommon.QueryStr = "select BannerID from Banners with (NoLock) where BannerID = 1 and Deleted=0"
                dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                If (dt.Rows.Count = 0) Then
                    DefaultBannerID = 0
                End If

                If (DefaultBannerID > 0) Then
                    ' assign offer to banner
                    MyCommon.QueryStr = "insert into BannerOffers with (RowLock) (BannerID, OfferID) values (@BannerID, @OfferID)"
                    MyCommon.DBParameters.Add("@BannerID", SqlDbType.Int).Value = DefaultBannerID
                    MyCommon.DBParameters.Add("@OfferID", SqlDbType.Int).Value = LogixID
                    MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)

                    ' assign banner id to the offer's location groups that are unassigned
                    MyCommon.QueryStr = "update LocationGroups with (RowLock) set BannerID = @BannerID" &
                                        "  where LocationGroupID in " &
                                        "    ( select LG.LocationGroupID from OfferLocations OL with (NoLock) " &
                                        "      inner join LocationGroups LG with (NoLock) on LG.LocationGroupID = OL.LocationGroupID " &
                                        "      where OL.OfferID = @OfferID and OL.Deleted=0 and LG.Deleted=0 and (LG.BannerID is null or LG.BannerID = 0) " &
                                        "    )"
                    MyCommon.DBParameters.Add("@BannerID", SqlDbType.Int).Value = DefaultBannerID
                    MyCommon.DBParameters.Add("@OfferID", SqlDbType.Int).Value = LogixID
                    MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)

                    ' assign locations to a banner
                    MyCommon.QueryStr = "update Locations with (RowLock) set BannerID = @BannerID where LocationID in " &
                                        "  (select LOC.LocationID from OfferLocations OL with (NoLock) " &
                                        "    inner join LocGroupItems LGI with (NoLock) on LGI.LocationGroupID = OL.LocationGroupID " &
                                        "    inner join Locations LOC with (NoLock) on LOC.LocationID = LGI.LocationID " &
                                        "    where OL.OfferID = @OfferID and OL.Deleted=0 and LGI.Deleted=0 and LOC.Deleted=0 and (LOC.BannerID is null or LOC.BannerID = 0) " &
                                        "   )"
                    MyCommon.DBParameters.Add("@BannerID", SqlDbType.Int).Value = DefaultBannerID
                    MyCommon.DBParameters.Add("@OfferID", SqlDbType.Int).Value = LogixID
                    MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)

                End If
            End If
        End If

    End Sub

    Private Function GetDefaultBanner() As Integer
        Dim dt As DataTable
        Dim TempStr As String = ""
        Dim DefaultBannerID As Integer = 0

        If (MyCommon.LRTadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixRT()

        If (MyCommon.Fetch_SystemOption(66) = "1") Then
            TempStr = MyCommon.Fetch_CPE_SystemOption(61)
            Integer.TryParse(TempStr, DefaultBannerID)

            If (DefaultBannerID > 0) Then
                MyCommon.QueryStr = "select BannerID from Banners with (NoLock) where BannerID = @BannerID and Deleted=0"
                MyCommon.DBParameters.Add("@BannerID", SqlDbType.Int).Value = DefaultBannerID
                dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                If (dt.Rows.Count = 0) Then
                    DefaultBannerID = 0
                End If
            End If
        End If

        Return DefaultBannerID
    End Function

    Private Sub AddOfferLocation(ByVal LogixID As Long, ByVal LocationGroupID As Long, ByVal Excluded As Boolean)

        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

        MyCommon.QueryStr = "pt_OfferLocations_Insert"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = LogixID
        MyCommon.LRTsp.Parameters.Add("@LocationGroupID", SqlDbType.BigInt).Value = LocationGroupID
        MyCommon.LRTsp.Parameters.Add("@Excluded", SqlDbType.Bit).Value = IIf(Excluded, 1, 10)
        MyCommon.LRTsp.Parameters.Add("@TCRMAStatusFlag", SqlDbType.Bit).Value = 3
        MyCommon.LRTsp.Parameters.Add("@PKID", SqlDbType.BigInt).Direction = ParameterDirection.Output
        MyCommon.LRTsp.ExecuteNonQuery()
        MyCommon.Close_LRTsp()

        MyCommon.Activity_Log(3, LogixID, 1, Copient.PhraseLib.Lookup("history.offer-addstore", 1))

    End Sub

    Private Sub AddOfferTerminal(ByVal LogixID As Long, ByVal TerminalTypeID As Long, ByVal Excluded As Boolean)

        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

        MyCommon.QueryStr = "insert into OfferTerminals with (RowLock) (OfferID,TerminalTypeID,LastUpdate) values(@OfferID,@TerminalTypeID,getdate())"
        MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = LogixID
        MyCommon.DBParameters.Add("@TerminalTypeID", SqlDbType.Int).Value = TerminalTypeID
        MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)

        MyCommon.Activity_Log(3, LogixID, 1, Copient.PhraseLib.Lookup("history.offer-addterminal", 1))

    End Sub

    Private Function DeployOffer(ByVal LogixID As Long, ByVal EngineId As Integer,
                                 ByVal OfferXML As XmlDocument, ByVal OperationType As RESPONSE_TYPES) As Boolean
        Dim Deployed As Boolean = False
        Dim DeploymentType As DEPLOYMENT_TYPES = DEPLOYMENT_TYPES.IMMEDIATE
        Dim DeferDeployTagValue As Boolean
        Dim DeployPostCollisionDetection As Boolean = False
        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

            ' if the defer deploy tag was sent in the offer XML then use that value; otherwise, use the interface option value
            If IsDeferDeploySet(OfferXML, OperationType, DeferDeployTagValue) Then
                DeploymentType = IIf(DeferDeployTagValue, DEPLOYMENT_TYPES.DEFERRED, DEPLOYMENT_TYPES.IMMEDIATE)
            Else
                DeploymentType = GetDeploymentType(OperationType)
            End If

            Select Case EngineId
                Case 0
                    If DeploymentType = DEPLOYMENT_TYPES.IMMEDIATE Then
                        MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=2,UpdateLevel=UpdateLevel+1 where OfferID = @OfferID"
                        MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = LogixID
                    ElseIf DeploymentType = DEPLOYMENT_TYPES.DEFERRED Then
                        MyCommon.QueryStr = "update Offers with (RowLock) set DeployDeferred=1, UpdateLevel=UpdateLevel+1 where OfferID = @OfferID"
                        MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = LogixID
                    End If
                Case 2
                    If DeploymentType = DEPLOYMENT_TYPES.IMMEDIATE Then
                        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=2,UpdateLevel=UpdateLevel+1, LastDeployClick=getdate(), DeployDeferred=0, LastUpdate=getdate(), LastUpdatedByAdminID=1 where IncentiveID = @IncentiveID"
                        MyCommon.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = LogixID
                    ElseIf DeploymentType = DEPLOYMENT_TYPES.DEFERRED Then
                        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set DeployDeferred=1, UpdateLevel=UpdateLevel+1, LastDeployClick=getdate(), LastUpdate=getdate() where IncentiveID = @IncentiveID"
                        MyCommon.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = LogixID
                    End If
                Case 9
                    Dim OfferService As IOffer = CurrentRequest.Resolver.Resolve(Of IOffer)()
                    Dim ProductGroupService As IProductGroupService = CurrentRequest.Resolver.Resolve(Of IProductGroupService)()
                    Dim isCollisionEnabled As AMSResult(Of Boolean) = OfferService.IsCollisionDetectionEnabled(Engines.UE, LogixID)
                    If isCollisionEnabled.ResultType = AMSResultType.Success AndAlso isCollisionEnabled.Result = True Then
                        If (_PGIDforOCD = -1) Then
                            Dim collisiondetectionservice As ICollisionDetectionService = CurrentRequest.Resolver.Resolve(Of ICollisionDetectionService)()
                            iOfferID_CDS = LogixID
                            DeploymentType_CDS = DeploymentType
                            AddHandler collisiondetectionservice.OnDetectionComplete, AddressOf DetectOfferCollisionCallBack
                            collisiondetectionservice.DetectOfferCollisionAsync(Of Int32)(LogixID)
                            Copient.Logger.Write_Log(AcceptanceLogFile, String.Format("Collision Detection Initiated for Offer ID: {0}", LogixID), True)
                            DeployPostCollisionDetection = True
                            MyCommon.QueryStr = ""
                            Exit Select
                        Else
                            Dim _ProdGroupName As AMSResult(Of String) = ProductGroupService.GetProductGroupName(_PGIDforOCD)
                            'Change log text below according to system option of automatic collision detection on product group change when that implementation is done
                            Copient.Logger.Write_Log(AcceptanceLogFile, String.Format("Collision Detection has not been performed for the offer: '{0}'. Collision Detection will be automatically performed for the offer when Process Product Group agent will be processing Product Group: '{1}'", _OfferNameforOCD, _ProdGroupName.Result), True)
                        End If
                    End If

                    If DeploymentType = DEPLOYMENT_TYPES.IMMEDIATE Then
                        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=2,UpdateLevel=UpdateLevel+1, LastDeployClick=getdate(), DeployDeferred=0, LastUpdate=getdate(), LastUpdatedByAdminID=1 where IncentiveID = @IncentiveID"
                        MyCommon.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = LogixID
                    ElseIf DeploymentType = DEPLOYMENT_TYPES.DEFERRED Then
                        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set DeployDeferred=1, UpdateLevel=UpdateLevel+1, LastDeployClick=getdate(), LastUpdate=getdate() where IncentiveID = @IncentiveID"
                        MyCommon.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = LogixID
                    End If
                Case Else
                    MyCommon.QueryStr = ""
            End Select

            If MyCommon.QueryStr <> "" Then
                MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
                If (MyCommon.RowsAffected > 0) Then
                    Deployed = True
                    MyCommon.Activity_Log(3, LogixID, 1, Copient.PhraseLib.Lookup("history.offer-deploy", 1))
                End If
            ElseIf (EngineId = 9 AndAlso DeployPostCollisionDetection = True) Then
                Deployed = True
                MyCommon.Activity_Log(3, LogixID, 1, Copient.PhraseLib.Lookup("history.offer-awaitingcollisiondetection", 1))
            End If

        Catch ex As Exception
            Deployed = False
        End Try

        Return Deployed
    End Function

    Public Sub DetectOfferCollisionCallBack(CollisionCount As AMSResult(Of Int32))
        Dim bCloseConn As Boolean = False
        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
            bCloseConn = True
            MyCommon.Open_LogixRT()
        End If
        If CollisionCount.ResultType = AMSResultType.Success AndAlso CollisionCount.Result = 0 Then
            Copient.Logger.Write_Log(AcceptanceLogFile, String.Format("Collision Detection Completed for Offer ID: {0}", iOfferID_CDS), True)
            MyCommon.Activity_Log(3, iOfferID_CDS, 1, String.Format(Copient.PhraseLib.Lookup("history.offer-deploy-nocollisionsfound", 1), iOfferID_CDS))
            If DeploymentType_CDS = DEPLOYMENT_TYPES.IMMEDIATE Then
                MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=2,UpdateLevel=UpdateLevel+1, LastDeployClick=getdate(), DeployDeferred=0, LastUpdate=getdate(), LastUpdatedByAdminID=1 where IncentiveID = @IncentiveID"
                MyCommon.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = iOfferID_CDS
            ElseIf DeploymentType_CDS = DEPLOYMENT_TYPES.DEFERRED Then
                MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set DeployDeferred=1, UpdateLevel=UpdateLevel+1, LastDeployClick=getdate(), LastUpdate=getdate() where IncentiveID = @IncentiveID"
                MyCommon.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = iOfferID_CDS
            End If
            MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
            If (MyCommon.RowsAffected > 0) Then
                MyCommon.Activity_Log(3, iOfferID_CDS, 1, Copient.PhraseLib.Lookup("history.offer-deploy", 1))
            End If
        Else
            Threading.ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf NotificationWorker), New Object() {iOfferID_CDS})
            Copient.Logger.Write_Log(AcceptanceLogFile, String.Format("Collision Detection Completed for Offer ID: {0}", iOfferID_CDS), True)
            MyCommon.Activity_Log(3, iOfferID_CDS, 1, String.Format(Copient.PhraseLib.Lookup("history.offer-deployfailed-collisionsfound", 1), iOfferID_CDS))
        End If
        If (MyCommon.LRTadoConn.State <> ConnectionState.Closed AndAlso bCloseConn = True) Then MyCommon.Close_LogixRT()
    End Sub

    Public Sub NotificationWorker(ByVal State As Object)
        Dim obj As Object() = State
        Dim iOfferID_CDS As Long = obj(0)
        Try
            Dim resolverbuilder As New WebRequestResolverBuilder()
            CurrentRequest.Resolver = resolverbuilder.GetResolver()
            resolverbuilder.Build()
            CurrentRequest.Resolver.AppName = "External Offer Connector"
            Dim collisiondetectionservice As ICollisionDetectionService = CurrentRequest.Resolver.Resolve(Of ICollisionDetectionService)(CurrentRequest.Resolver.AppName)
            collisiondetectionservice.SendNotifications(iOfferID_CDS)
        Catch ex As Exception
            Copient.Logger.Write_Log(RejectionLogFile, "Error: " + ex.ToString, True)
        End Try
    End Sub




    Private Function GetDeploymentType(ByVal OperationType As RESPONSE_TYPES) As DEPLOYMENT_TYPES
        Dim Type As DEPLOYMENT_TYPES = DEPLOYMENT_TYPES.IMMEDIATE
        Dim OptionValue As String = ""

        Try
            ' lookup the appropriate interface option based on the type of operation
            Select Case OperationType
                Case RESPONSE_TYPES.ADD_OFFER
                    OptionValue = MyCommon.Fetch_InterfaceOption(19)
                Case RESPONSE_TYPES.UPDATE_OFFER
                    OptionValue = MyCommon.Fetch_InterfaceOption(20)
            End Select

            ' convert the option value to the deployment type
            Select Case OptionValue
                Case "1"
                    Type = DEPLOYMENT_TYPES.IMMEDIATE
                Case "2"
                    Type = DEPLOYMENT_TYPES.DEFERRED
            End Select

        Catch ex As Exception
            Type = DEPLOYMENT_TYPES.IMMEDIATE
        End Try

        Return Type
    End Function

    Private Function CreateProductGroup(ByVal LogixID As Long, ByVal ClientOfferID As String, ByVal Excluded As Boolean,
                                        ByRef ErrorCode As ERROR_CODES, ByRef ErrorMsg As String) As Long
        Dim ProductGroupID As Long = 0
        Dim ROID As Long = 0
        Dim dt As DataTable

        ' no product group currently exists for this offer, so create one and add this product to it.
        ' create new product group to be used as an exclusion group for this offer
        Try
            If (MyCommon.LRTadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixRT()

            MyCommon.QueryStr = "dbo.pt_ProductGroups_Insert"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 200).Value = "Product Group for Offer " & LogixID
            MyCommon.LRTsp.Parameters.Add("@AnyProduct", SqlDbType.Bit).Value = 0
            MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt).Direction = ParameterDirection.Output
            MyCommon.LRTsp.ExecuteNonQuery()
            ProductGroupID = MyCommon.LRTsp.Parameters("@ProductGroupID").Value
            MyCommon.Close_LRTsp()

            If (ProductGroupID > 0) Then
                MyCommon.Activity_Log(5, ProductGroupID, 1, Copient.PhraseLib.Lookup("history.pgroup-create", 1))

                ' get the ROID for this offer
                MyCommon.QueryStr = "select RewardOptionID from CPE_RewardOptions RO with (NoLock) where Deleted=0 and IncentiveID = @IncentiveID"
                MyCommon.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = LogixID
                dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                If (dt.Rows.Count > 0) Then
                    ' assign the excluded product group to the offer
                    ROID = MyCommon.NZ(dt.Rows(0).Item("RewardOptionID"), -1)
                    MyCommon.QueryStr = "insert into CPE_IncentiveProductGroups (RewardOptionID,ProductGroupID,ExcludedProducts,Deleted,LastUpdate,RequiredFromTemplate,TCRMAStatusFlag,Disqualifier) " &
                                        " values(@RewardOptionID, @ProductGroupID, @ExcludedProducts, 0, getdate(), 0, 3, 0)"
                    MyCommon.DBParameters.Add("@RewardOptionID", SqlDbType.BigInt).Value = ROID
                    MyCommon.DBParameters.Add("@ProductGroupID", SqlDbType.Int).Value = ProductGroupID
                    MyCommon.DBParameters.Add("@ExcludedProducts", SqlDbType.Bit).Value = IIf(Excluded, 1, 0)
                    MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
                    If (MyCommon.RowsAffected > 0) Then
                        MyCommon.Activity_Log(3, LogixID, 1, "Added condition for product group " & ProductGroupID)
                    End If
                Else
                    ErrorCode = ERROR_CODES.ERROR_ADD_PRODUCT_FAILED
                    ErrorMsg = "Error encountered while attempting to create an product group for offer " & ClientOfferID
                End If

            Else
                ErrorCode = ERROR_CODES.ERROR_ADD_PRODUCT_FAILED
                ErrorMsg = "Error encountered while attempting to create an product group for offer " & ClientOfferID
            End If
        Catch ex As Exception
            ProductGroupID = 0
        End Try

        Return ProductGroupID
    End Function

    Private Function AddProductToGroup(ByVal ProductGroupID As Long, ByVal ClientProductID As String, ByVal ProductDesc As String,
                                       ByVal ClientOfferID As String, ByRef ErrorCode As ERROR_CODES, ByRef ErrorMsg As String) As Boolean
        Dim Added As Boolean = False
        Dim OutputStatus As Integer = 0

        Try
            If (MyCommon.Fetch_CM_SystemOption(82) = "1" AndAlso MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM) = True) Then
                If (CheckItemCode(ClientProductID, ErrorMsg) = False) Then
                    ErrorCode = ERROR_CODES.ERROR_ADD_PRODUCT_FAILED
                    Return False
                    Exit Function
                End If
            End If

            MyCommon.QueryStr = "dbo.pa_ProdGroupItems_ManualInsert"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@ExtProductID", SqlDbType.NVarChar, 120).Value = ClientProductID
            MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt).Value = ProductGroupID
            MyCommon.LRTsp.Parameters.Add("@ProductTypeID", SqlDbType.Int).Value = 1 ' UPC
            MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 200).Value = ProductDesc
            MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
            MyCommon.LRTsp.Parameters.Add("@ProductStatus", SqlDbType.Int).Direction = ParameterDirection.Output
            MyCommon.LRTsp.ExecuteNonQuery()
            OutputStatus = MyCommon.LRTsp.Parameters("@Status").Value
            MyCommon.Close_LRTsp()

            '      If (OutputStatus <> 0) Then
            ' ErrorCode = ERROR_CODES.ERROR_ADD_PRODUCT_FAILED
            ' If ErrorMsg <> "" Then ErrorMsg &= "; "
            ' ErrorMsg &= "Error encountered while attempting to add ProductID " & ClientProductID & " to offer " & ClientOfferID
            ' Else
            MyCommon.Activity_Log(5, ProductGroupID, 1, Copient.PhraseLib.Lookup("history.pgroup-add", 1) & " " & ClientProductID)
            ' End If
            Added = True
        Catch ex As Exception
            Added = False
        End Try

        Return Added
    End Function

    Public Function CheckItemCode(ByRef itemCode As String, Optional ByRef StatusString As String = "") As Boolean
        Dim itemVal As New Copient.ItemCodeValidation
        Dim bRetVal As Boolean = False
        bRetVal = itemVal.ValidateItemCode(itemCode, StatusString)
        itemVal = Nothing
        CheckItemCode = bRetVal
    End Function

    ' AMS-6260: Modified method name to handle both Manufacuturer Coupon and Store coupon while adding and updating EOC offer
    Private Sub HandleOfferTypes(ByVal ExternalSourceID As String, ByVal LogixID As Long, ByVal OfferXmlDoc As XmlDocument)

        If (MyCommon.LRTadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixRT()

        ' if the manufacturer coupon is not set in the XML document, then use the default value
        ' set for the external source
        If GetManufacturerCoupon(ExternalSourceID, OfferXmlDoc) Then
            MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set ManufacturerCoupon=1 where IncentiveID = @IncentiveID"
            MyCommon.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = LogixID
            MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
        End If

        ' if the store coupon is not set in the XML document, then use the default value
        ' set for the external source
        If GetStoreCoupon(ExternalSourceID, OfferXmlDoc) Then
            MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StoreCoupon=1 where IncentiveID = @IncentiveID"
            MyCommon.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = LogixID
            MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
        End If

    End Sub

    Private Sub HandleBestDealSetting(ByVal ExternalSourceID As String, ByVal LogixID As Long)
        Dim dt As DataTable
        Dim row As DataRow
        Dim IsMfgCoupon As Boolean = False

        If (MyCommon.LRTadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixRT()

        ' when an offer is set as a manufacturer coupon offer, best deal is not available and should be turned off for the offer
        MyCommon.QueryStr = "select ManufacturerCoupon from CPE_Incentives with (NoLock) where IncentiveID = @IncentiveID"
        MyCommon.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = LogixID
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        If dt.Rows.Count > 0 Then
            IsMfgCoupon = MyCommon.NZ(dt.Rows(0).Item("ManufacturerCoupon"), False)
            If IsMfgCoupon Then
                MyCommon.QueryStr = "select DISC.DiscountID from CPE_Discounts DISC with (NoLock) " &
                                    "inner join CPE_Deliverables DEL with (NoLock) on DEL.OutputID = DISC.DiscountID and DEL.DeliverableTypeID=2 " &
                                    "   and DEL.RewardOptionPhase=3 and DEL.Deleted=0 " &
                                    "inner join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID = DEL.RewardOptionId and RO.Deleted=0 " &
                                    "where DISC.Deleted=0 and RO.IncentiveID = @IncentiveID"
                MyCommon.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = LogixID
                dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                If dt.Rows.Count Then
                    For Each row In dt.Rows
                        MyCommon.QueryStr = "update CPE_Discounts with (RowLock) set BestDeal=0 where DiscountID = @DiscountID"
                        MyCommon.DBParameters.Add("@DiscountID", SqlDbType.BigInt).Value = MyCommon.NZ(row.Item("DiscountID"), 0)
                        MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
                    Next
                End If
            End If
        End If

    End Sub

    Private Sub CmHandleBestDealSetting(ByVal ExternalSourceID As String, ByVal LogixID As Long)
        Dim dt As DataTable
        Dim row As DataRow
        Dim IsMfgCoupon As Boolean = False

        If (MyCommon.LRTadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixRT()

        ' when an offer is set as a manufacturer coupon offer, best deal is not available and should be turned off for the offer
        MyCommon.QueryStr = "select SponsorID from OfferRewards with (NoLock) " &
                            "where deleted=0 and RewardTypeId=1 and OfferID = @OfferID"
        MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = LogixID
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        If dt.Rows.Count > 0 Then
            IsMfgCoupon = MyCommon.NZ(dt.Rows(0).Item("SponsorID"), False)
            If IsMfgCoupon Then
                MyCommon.QueryStr = "select DISC.DiscountID from CPE_Discounts DISC with (NoLock) " &
                                    "inner join CPE_Deliverables DEL with (NoLock) on DEL.OutputID = DISC.DiscountID and DEL.DeliverableTypeID=2 " &
                                    "   and DEL.RewardOptionPhase=3 and DEL.Deleted=0 " &
                                    "inner join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID = DEL.RewardOptionId and RO.Deleted=0 " &
                                    "where DISC.Deleted=0 and RO.IncentiveID = @IncentiveID"
                MyCommon.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = LogixID
                dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                If dt.Rows.Count Then
                    For Each row In dt.Rows
                        MyCommon.QueryStr = "update CPE_Discounts with (RowLock) set BestDeal=0 where DiscountID = @DiscountID"
                        MyCommon.DBParameters.Add("@DiscountID", SqlDbType.Int).Value = MyCommon.NZ(row.Item("DiscountID"), 0)
                        MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
                    Next
                End If
            End If
        End If

    End Sub

    ' If the vendor chargeback is blank then set it to the AnyVendor (unspecified) vendor by default.
    Private Sub HandleVendorChargeback(ByVal LogixID As Long)
        Dim dt As DataTable
        Dim UnspecifiedVendorID As Integer = 1
        Dim NeedsAssignment As Boolean = False

        ' check if the chargeback vendor id is a valid and chargeable vendor, if not then assign to the Unspecified vendor
        MyCommon.QueryStr = "select IsNull(ChargeBackVendorID,0) as VendorID from CPE_Incentives INC with (NoLock) " &
                            "inner join Vendors V with (NoLock) on V.VendorID = INC.ChargebackVendorID " &
                            "where V.Chargeable=1 and V.Deleted=0 and INC.Deleted=0 and INC.IncentiveID = @IncentiveID"
        MyCommon.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = LogixID
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        If dt.Rows.Count > 0 Then
            NeedsAssignment = (MyCommon.NZ(dt.Rows(0).Item("VendorID"), 0) = 0)
        Else
            NeedsAssignment = True
        End If

        If NeedsAssignment Then
            MyCommon.QueryStr = "select VendorID from Vendors with (NoLock) where AnyVendor=1 and Deleted=0"
            dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            If (dt.Rows.Count > 0) Then
                UnspecifiedVendorID = MyCommon.NZ(dt.Rows(0).Item("VendorID"), 0)
                If UnspecifiedVendorID > 0 Then
                    MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set ChargebackVendorID = @ChargebackVendorID " &
                                        "  where IncentiveID = @IncentiveID and Deleted=0 "
                    MyCommon.DBParameters.Add("@ChargebackVendorID", SqlDbType.Int).Value = UnspecifiedVendorID
                    MyCommon.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = LogixID
                    MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
                End If
            End If
        End If

    End Sub

    Private Function GetManufacturerCoupon(ByVal ExternalSourceID As String, ByVal OfferXmlDoc As XmlDocument) As Boolean
        Dim IsMfgCoupon As Boolean
        Dim dt As DataTable

        If Not IsManufacturerCouponSet(OfferXmlDoc, IsMfgCoupon) Then
            ' check if this external source has a default value set for the manufacturer coupon
            MyCommon.QueryStr = "select DefaultAsMfgCoupon from ExtCRMInterfaces with (NoLock) where ExtCode = @ExtCode " &
                            "  and Editable=1 and Active=1 and Deleted=0"
            MyCommon.DBParameters.Add("@ExtCode", SqlDbType.NVarChar).Value = ExternalSourceID.ConvertBlankIfNothing
            dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            If dt.Rows.Count > 0 Then
                IsMfgCoupon = MyCommon.NZ(dt.Rows(0).Item("DefaultAsMfgCoupon"), False)
            End If
        End If

        Return IsMfgCoupon
    End Function

    Private Function GetStoreCoupon(ByVal ExternalSourceID As String, ByVal OfferXmlDoc As XmlDocument) As Boolean
        Dim IsStoreCoupon As Boolean
        Dim dt As DataTable

        If Not IsStoreCouponSet(OfferXmlDoc, IsStoreCoupon) Then
            ' check if this external source has a default value set for store coupon
            MyCommon.QueryStr = "select DefaultAsStoreCoupon from ExtCRMInterfaces with (NoLock) where ExtCode = @ExtCode " &
                            "  and Editable=1 and Active=1 and Deleted=0"
            MyCommon.DBParameters.Add("@ExtCode", SqlDbType.NVarChar).Value = ExternalSourceID.ConvertBlankIfNothing
            dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            If dt.Rows.Count > 0 Then
                IsStoreCoupon = MyCommon.NZ(dt.Rows(0).Item("DefaultAsStoreCoupon"), False)
            End If
        End If

        Return IsStoreCoupon
    End Function

    ' when an offer is designated as a manufacturer coupon offer, only item-level and basket-level discounts are valid 
    Private Function IsValidDiscountType(ByVal ExternalSourceID As String, ByVal EngineID As Integer, ByVal OfferXmlDoc As XmlDocument, ByRef ErrorMsg As String) As Boolean
        Dim ValidDiscType As Boolean = True
        Dim IsMfgCoupon As Boolean = False
        Dim DiscType As String = ""
        Dim ItemLimit As String = ""
        Dim AmountType As String = ""
        Dim HasAnyProduct As Boolean = False
        Dim PGID As Long

        IsMfgCoupon = GetManufacturerCoupon(ExternalSourceID, OfferXmlDoc)

        If OfferXmlDoc IsNot Nothing Then
            If TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/DiscountType", DiscType) Then
                Select Case DiscType
                    Case "ITEM_LEVEL"
                        ' ensure that any product is not selected for an item-level discount
                        If TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/ProductGroupID", PGID) AndAlso PGID = 1 Then
                            ValidDiscType = False
                            ErrorMsg = "Item-level discount types do not permit Any Product to be selected"
                        Else
                            ValidDiscType = True
                        End If
                    Case "BASKET_LEVEL"
                        ' only Fixed Amount off discounts are allowable for basket-level discounts
                        If TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/AmountType", AmountType) Then
                            If IsMfgCoupon AndAlso AmountType <> "FIXED_AMOUNT_OFF" Then
                                ValidDiscType = False
                                ErrorMsg = "Only discount types of Fixed amount off are permitted for " &
                                           "basket-level manufacturer coupon offers"
                            ElseIf (EngineID <> 0) AndAlso (Not IsAnyProductDiscount(OfferXmlDoc)) Then
                                ValidDiscType = False
                                ErrorMsg = "Only the product condition of Any Product is allowable for  " &
                                           "basket-level manufacturer coupon offers"
                            ElseIf (EngineID <> 0) AndAlso TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/ItemLimit", ItemLimit) Then
                                ValidDiscType = False
                                ErrorMsg = "Basket-level discount types do not permit item limits"
                            Else
                                ValidDiscType = True
                            End If
                        Else
                            ValidDiscType = False
                            ErrorMsg = "Only discount types of Fixed amount off are permitted for " &
                                       "basket-level manufacturer coupon offers"
                        End If
                    Case "DEPARTMENT_LEVEL"
                        If IsMfgCoupon Then
                            ValidDiscType = False
                            ErrorMsg = "Department level manufacturer coupon offers are not permitted."
                        Else
                            If TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/ProductGroupID", PGID) AndAlso PGID = 1 Then
                                ValidDiscType = False
                                ErrorMsg = "Department-level discount types do not permit Any Product to be selected"
                            ElseIf (EngineID <> 0) AndAlso TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/ItemLimit", ItemLimit) Then
                                ValidDiscType = False
                                ErrorMsg = "Department-level discount types do not permit item limits"
                            Else
                                ValidDiscType = True
                            End If
                        End If
                    Case "GROUP_LEVEL"
                        If TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/AmountType", AmountType) Then
                            ' only Fixed Amount off discounts are allowable for group-level discounts
                            If IsMfgCoupon AndAlso AmountType <> "FIXED_AMOUNT_OFF" Then
                                ValidDiscType = False
                                ErrorMsg = "Only discount types of Fixed amount off are permitted for " &
                                           "group-level manufacturer coupon offers"
                            ElseIf TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/ProductGroupID", PGID) AndAlso PGID = 1 Then
                                ' ensure that any product is not selected for an group-level discount
                                ValidDiscType = False
                                ErrorMsg = "Group-level discount types do not permit Any Product to be selected"
                            ElseIf (EngineID <> 0) AndAlso TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/ItemLimit", ItemLimit) Then
                                ValidDiscType = False
                                ErrorMsg = "Group-level discount types do not permit item limits"
                            Else
                                ValidDiscType = True
                            End If
                        Else
                            ValidDiscType = False
                            ErrorMsg = "Only discount types of Fixed amount off are permitted for " &
                                       "group-level manufacturer coupon offers"
                        End If
                    Case Else
                        ValidDiscType = False
                        ErrorMsg = "Only Item-level and Basket-level manufacturer coupon offers are permitted."
                End Select
            End If
        End If

        Return ValidDiscType
    End Function

    ' when an offer is a price point items offer, offer must have a discounted product group, item-level discount type only
    Private Function IsValidAmountType(ByVal ExternalSourceID As String, ByVal OfferXmlDoc As XmlDocument, ByRef ErrorMsg As String) As Boolean
        Dim ValidAmtType As Boolean = True
        Dim DiscType As String = ""
        Dim AmountType As String = ""

        If OfferXmlDoc IsNot Nothing AndAlso TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/AmountType", AmountType) Then

            'AL-4258
            If TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/DiscountType", DiscType) AndAlso DiscType = "DEPARTMENT_LEVEL" Then
                If Copient.commonShared.Contains(AmountType, "FIXED_AMOUNT_OFF", "PERCENT_OFF") = False Then
                    ValidAmtType = False
                    ErrorMsg = "Invalid AmountType-" & AmountType & " for Department level discount."
                    Return ValidAmtType
                End If
            End If

            ' offers of price point (items) amount type should not use the Any Product group and should be item level discount types
            Select Case AmountType
                Case "PRICE_POINT_ITEMS", "PRICE_POINT_WEIGHT_VOLUME"
                    If IsAnyProductDiscount(OfferXmlDoc) Then
                        ValidAmtType = False
                        ErrorMsg = "Price Point (items) discount amount types do not permit assignment of the Any Product product group."
                    ElseIf DiscType <> "ITEM_LEVEL" Then
                        ValidAmtType = False
                        ErrorMsg = "Price Point (items) discount type must be set to ITEM_LEVEL."
                    End If
                Case Else
                    ValidAmtType = True
            End Select
        End If

        Return ValidAmtType
    End Function

    ' validate that a chargeback vendor code is currently in Logix and chargeable if this is a manufacturer coupon offer.
    Private Function IsValidChargebackVendor(ByVal ExternalSourceID As String, ByVal OfferXmlDoc As XmlDocument,
                                             ByRef ErrorMessage As String) As Boolean
        Dim ValidVendor As Boolean = False
        Dim IsMfgCoupon, IsChargeable As Boolean
        Dim VendorID As Integer = 0
        Dim MyCpeOffer As New Copient.CPEOffer
        Dim TempVal As String = ""
        Dim VendorName As String = ""

        If TryParseElementValue(OfferXmlDoc, "//Offer/ChargebackVendorCode", TempVal) Then
            TryParseAttributeValue(OfferXmlDoc, "//Offer/ChargebackVendorCode", "name", VendorName)
            VendorID = MyCpeOffer.GetVendorID(TempVal, VendorName)
            If VendorID > 0 Then
                IsChargeable = IsChargeableVendor(VendorID)
                If IsChargeable Then
                    ValidVendor = True
                Else
                    ValidVendor = False
                    ErrorMessage = "Chargeback Vendor Code " & TempVal & " is not designated as a chargeable vendor in Logix."
                End If
            Else
                ValidVendor = False
                ErrorMessage = "Chargeback Vendor Code " & TempVal & " is not a valid code stored in Logix."
            End If
        Else
            IsMfgCoupon = GetManufacturerCoupon(ExternalSourceID, OfferXmlDoc)
            If IsMfgCoupon Then
                ValidVendor = False
                ErrorMessage = "Chargeback Vendor Code is a required field for all offers designated as Manufacturer coupons."
            Else
                ValidVendor = True
            End If
        End If

        Return ValidVendor
    End Function

    Private Function IsValidDiscountScorecard(ByVal ExternalSourceID As String, ByVal OfferXmlDoc As XmlDocument, ByRef ErrorMsg As String) As Boolean
        Dim ValidScorecard As Boolean = True
        Dim TempStr As String = ""
        Dim ScorecardID As Integer
        Dim dt As DataTable

        ' if a discount scorecard is sent in the XML, then validate it
        If OfferXmlDoc IsNot Nothing AndAlso TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/ScorecardID", TempStr) AndAlso Integer.TryParse(TempStr, ScorecardID) Then
            MyCommon.QueryStr = "select ScorecardID from Scorecards with (NoLock) " &
                                "where ScorecardTypeID=3 and Deleted=0 and ScorecardID = @ScorecardID "
            MyCommon.DBParameters.Add("@ScorecardID", SqlDbType.Int).Value = ScorecardID
            dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            ValidScorecard = (dt.Rows.Count > 0)
            If Not ValidScorecard Then
                ErrorMsg = "Discount Scorecard ID " & ScorecardID & " is invalid."
            End If
        End If

        Return ValidScorecard
    End Function

    Private Function IsValidChargebackDept(ByVal OfferXmlDoc As XmlDocument, ByVal EngineID As Integer,
                                           ByRef ErrorMsg As String) As Boolean
        Dim ValidDept As Boolean = True
        Dim ExtDept As String = ""
        Dim ChargebackDeptID As Integer = -1

        TryParseElementValue(OfferXmlDoc, "//Offer/Rewards/Discount/ChargebackDept", ExtDept)

        If ExtDept.Trim <> "" Then
            ChargebackDeptID = LookupChargebackDeptID(ExtDept)
            ValidDept = (ChargebackDeptID > -1)

            If Not ValidDept Then
                ErrorMsg = "Chargeback Department " & ExtDept & " is invalid"
            End If
        End If

        Return ValidDept
    End Function

    Private Function LookupChargebackDeptID(ByVal ExtDeptID As String) As Integer
        Dim ChargebackDeptID As Integer = -1

        If (Not String.IsNullOrEmpty(ExtDeptID)) Then

            MyCommon.QueryStr = "select ChargebackDeptID from ChargebackDepts with (NoLock) where Deleted = 0 and ExternalID = @ExternalID"
            MyCommon.DBParameters.Add("@ExternalID", SqlDbType.NVarChar).Value = ExtDeptID
            Dim dt As DataTable = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            If dt.Rows.Count > 0 Then
                ChargebackDeptID = MyCommon.NZ(dt.Rows(0).Item("ChargebackDeptID"), -1)
            End If
        End If

        Return ChargebackDeptID
    End Function

    Private Function IsChargeableVendor(ByVal VendorID As Integer) As Boolean
        Dim IsChargeable As Boolean = False
        Dim dt As DataTable

        If (MyCommon.LRTadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixRT()

        MyCommon.QueryStr = "select Chargeable from Vendors with (NoLock) where Deleted=0 and VendorID = @VendorID"
        MyCommon.DBParameters.Add("@VendorID", SqlDbType.Int).Value = VendorID
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        If dt.Rows.Count > 0 Then
            IsChargeable = MyCommon.NZ(dt.Rows(0).Item("Chargeable"), False)
        End If

        Return IsChargeable
    End Function

    Private Function ExceedsMaximumOffers(ByVal ExtInterfaceID As Integer, ByRef MaxOffers As Long) As Boolean
        Dim OverLimit As Boolean = False
        Dim dt As DataTable
        Dim CurrentOffers As Integer = 0

        Try
            If (MyCommon.LRTadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixRT()
            MyCommon.QueryStr = "select MaxOffers from ExtCRMInterfaces with (NoLock) where ExtInterfaceID = @ExtInterfaceID and deleted=0"
            MyCommon.DBParameters.Add("@ExtInterfaceID", SqlDbType.Int).Value = ExtInterfaceID
            dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            If dt.Rows.Count > 0 Then
                MaxOffers = MyCommon.NZ(dt.Rows(0).Item("MaxOffers"), 0)
                If MaxOffers < 0 Then
                    ' a negative value effectively disables the offer submission for the external source.
                    OverLimit = True
                ElseIf MaxOffers = 0 Then
                    ' 0 value indicates that there is an unlimited number of offers that this external source may submit
                    OverLimit = False
                ElseIf MaxOffers > 0 Then
                    MyCommon.QueryStr = "select count(*) as IncentiveCount from CPE_Incentives as INC with (NoLock) " &
                                        "where Deleted=0 and InboundCRMEngineID = @InboundCRMEngineID " &
                                        "  and DateAdd(d, 1, EndDate) >= getdate()"
                    MyCommon.DBParameters.Add("@InboundCRMEngineID", SqlDbType.Int).Value = ExtInterfaceID
                    dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                    If dt.Rows.Count > 0 Then
                        CurrentOffers = (MyCommon.NZ(dt.Rows(0).Item("IncentiveCount"), 0))
                        OverLimit = (CurrentOffers >= MaxOffers)
                    End If
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try

        Return OverLimit
    End Function

    Private Function IsAnyProductDiscount(ByVal OfferXmlDoc As XmlDocument) As Boolean
        Dim AnyProduct As Boolean = False
        Dim nodeList As XmlNodeList
        Dim PGID As Long = 0
        Dim HasProductList As Boolean = False
        Dim PGCount As Integer
        Dim SameAsProductCon As Boolean

        If OfferXmlDoc IsNot Nothing Then
            nodeList = OfferXmlDoc.SelectNodes("//Offer/Rewards/Discount/ProductList")
            HasProductList = (nodeList IsNot Nothing AndAlso nodeList.Count > 0)

            nodeList = OfferXmlDoc.SelectNodes("//Offer/Rewards/Discount/ProductGroupID")
            If nodeList IsNot Nothing Then
                PGCount = nodeList.Count
                If PGCount = 1 Then Long.TryParse(nodeList.Item(0).InnerText, PGID)
            End If

            nodeList = OfferXmlDoc.SelectNodes("//Offer/Rewards/Discount/SameAsConditionProducts")
            If nodeList IsNot Nothing AndAlso nodeList.Count > 0 Then
                Boolean.TryParse(nodeList.Item(0).InnerText(), SameAsProductCon)
                If SameAsProductCon AndAlso Not HasProductList Then
                    ' find the product condition
                    nodeList = OfferXmlDoc.SelectNodes("//Offer/Conditions/Product/ProductList")
                    HasProductList = (nodeList IsNot Nothing AndAlso nodeList.Count > 0)

                    If Not HasProductList Then
                        nodeList = OfferXmlDoc.SelectNodes("//Offer/Conditions/Product/ProductGroupID")
                        If nodeList IsNot Nothing Then
                            PGCount += nodeList.Count
                            If PGCount = 1 Then Long.TryParse(nodeList.Item(0).InnerText, PGID)
                        End If
                    End If
                End If
            End If

            ' it's only Any Product if that's the only product for both the discount and the condition
            If Not HasProductList AndAlso PGCount = 1 AndAlso PGID = 1 Then
                AnyProduct = True
            End If

        End If

        Return AnyProduct
    End Function

    Private Sub SetClientOfferIDAsLogixID(ByVal LogixID As Long, ByVal EngineID As Integer)

        Dim setClientOfferIdStatement As String

        Select Case EngineID
            Case 2, 9
                setClientOfferIdStatement = "update CPE_Incentives with (RowLock) set ClientOfferID = CAST(@LogixID as nvarchar(max)) where IncentiveID = @LogixID"
            Case 0
                setClientOfferIdStatement = "update Offers with (RowLock) set ExtOfferID = CAST(@LogixID as nvarchar(max)) where OfferID = @LogixID"
            Case Else
                Return
        End Select

        If (MyCommon.LRTadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixRT()
        MyCommon.QueryStr = setClientOfferIdStatement
        MyCommon.DBParameters.Add("@LogixID", SqlDbType.BigInt).Value = LogixID
        MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
    End Sub

    Private Sub ApplyDiscountOptions(ByVal ExternalSourceID As String, ByVal OfferXmlDoc As XmlDocument, ByVal DiscountID As Long, Optional ByVal MyDiscount As Copient.Discount = Nothing)
        Dim AllowNegOptID As Integer
        Dim FlexNegOptID As Integer
        Dim IsMfgCoupon As Boolean
        Dim AllowNeg, FlexNeg As Boolean
        Dim AllowNegVal, FlexNegVal As Boolean
        Dim SVProgramID As Integer = 0
        Try
            If (MyCommon.LRTadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixRT()

            IsMfgCoupon = GetManufacturerCoupon(ExternalSourceID, OfferXmlDoc)
            AllowNegOptID = IIf(IsMfgCoupon, 1, 2)
            FlexNegOptID = IIf(IsMfgCoupon, 3, 4)

            AllowNeg = (MyCommon.Fetch_InterfaceOption(AllowNegOptID) = "1")
            FlexNeg = (MyCommon.Fetch_InterfaceOption(FlexNegOptID) = "1")

            AllowNegVal = IIf(AllowNeg, 1, 0)
            FlexNegVal = IIf(FlexNeg, 1, 0)
            MyCommon.QueryStr = "update CPE_Discounts with (RowLock) " &
                                " set AllowNegative = @AllowNegative," &
                                "     FlexNegative = @FlexNegative" &
                                " where DiscountID = @DiscountID"
            MyCommon.DBParameters.Add("@AllowNegative", SqlDbType.Bit).Value = AllowNegVal
            MyCommon.DBParameters.Add("@FlexNegative", SqlDbType.Bit).Value = FlexNegVal
            MyCommon.DBParameters.Add("@DiscountID", SqlDbType.Int).Value = DiscountID
            MyCommon.DBParameters.Add("@SVProgramID", SqlDbType.Int).Value = SVProgramID
            MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

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
                        Value = Value
                End Select
                ValTable.Add(ChildNodeName, Value)
            End If
        End If

        Return Value
    End Function

    Private Function HandleCustomerIdPadding(ByVal CustomerID As String) As String
        Dim PaddedStr As String = CustomerID
        Dim IdLength As Integer = 0

        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

        PaddedStr = MyCommon.Pad_ExtCardID(CustomerID, Copient.commonShared.CardTypes.CUSTOMER)
        Return PaddedStr
    End Function

    Private Function HandleProductIdPadding(ByVal ProductID As String) As String
        Dim PaddedStr As String = ProductID
        Dim IdLength As Integer = 0
        Dim rst As DataTable

        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
        MyCommon.QueryStr = "select paddingLength from producttypes where producttypeid=1"
        rst = MyCommon.LRT_Select
        If rst IsNot Nothing Then
            IdLength = Convert.ToInt32(rst.Rows(0).Item("paddingLength"))
        End If
        'IdLength = MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(52))
        If (IdLength > 0) Then
            PaddedStr = ProductID.PadLeft(IdLength, "0")
        End If

        Return PaddedStr
    End Function

    Private Function CleanXmlString(ByVal sUnformattedXml As String) As String
        Dim CleanStr As String = ""

        If (sUnformattedXml IsNot Nothing) Then
            CleanStr = sUnformattedXml
            CleanStr = CleanStr.Replace("&", "&amp;")
            CleanStr = CleanStr.Replace("""", "&quot;")
            CleanStr = CleanStr.Replace("<", "&lt;")
        End If

        Return CleanStr
    End Function

    Private Sub AppendToLog(ByVal ExternalSourceID As String, ByVal FunctionName As String, ByVal OfferXML As String,
                          ByVal LogText As String, ByVal ErrorCode As ERROR_CODES, ByVal ErrorMsg As String, Optional ByVal IsUnmaskedString As Boolean = False)
        Dim LogOfferXML As Boolean
        Dim LogString As String = ""
        Dim AndLogText As String = ""

        If LogText <> "" Then AndLogText = LogText & ";"

        If (IsUnmaskedString = True) Then
            If ErrorCode = ERROR_CODES.ERROR_NONE Then
                Return
            End If
            LogString = "Error=" & FunctionName & "; " & AndLogText & " ExternalSourceID=" & ExternalSourceID & "; Error_encountered=" & ErrorCode.ToString & " " & ErrorMsg & "; Server=" & Environment.MachineName '& "; IP = " & IP
            LogOfferXML = (MyCommon.Fetch_InterfaceOption(24) = "1")
            If LogOfferXML AndAlso OfferXML <> "" Then LogString = LogString & vbCrLf & OfferXML
            Copient.Logger.Write_Log(RejectionLogFileUnmasked, LogString, True)
            Return
        End If

        Select Case ErrorCode
            Case ERROR_CODES.ERROR_NONE
                LogString = "Success=" & FunctionName & "; " & AndLogText & " ExternalSourceID=" & ExternalSourceID & "; Server= " & Environment.MachineName

                LogOfferXML = (MyCommon.Fetch_InterfaceOption(23) = "1")

                If LogOfferXML AndAlso OfferXML <> "" Then LogString = LogString & vbCrLf & OfferXML

                Copient.Logger.Write_Log(AcceptanceLogFile, LogString, True)
            Case Else
                LogString = "Error=" & FunctionName & "; " & AndLogText & " ExternalSourceID=" & ExternalSourceID & "; Error_encountered=" & ErrorCode.ToString & " " & ErrorMsg & "; Server=" & Environment.MachineName '& "; IP = " & IP

                LogOfferXML = (MyCommon.Fetch_InterfaceOption(24) = "1")

                If LogOfferXML AndAlso OfferXML <> "" Then LogString = LogString & vbCrLf & OfferXML

                Copient.Logger.Write_Log(RejectionLogFile, LogString, True)
        End Select

    End Sub

    Private Function IsDeferDeploySet(ByVal OfferXmlDoc As XmlDocument, ByVal OperationType As RESPONSE_TYPES, ByRef DeferDeployValue As Boolean) As Boolean
        Dim DeferSet As Boolean = False
        Dim deferNode As XmlNode = Nothing

        Select Case OperationType
            Case RESPONSE_TYPES.ADD_OFFER
                deferNode = OfferXmlDoc.SelectSingleNode("//Offer/Actions/DeferDeploymentForAdd")
            Case RESPONSE_TYPES.UPDATE_OFFER
                deferNode = OfferXmlDoc.SelectSingleNode("//Offer/Actions/DeferDeploymentForUpdate")
        End Select

        If (deferNode IsNot Nothing) Then
            DeferSet = Boolean.TryParse(deferNode.InnerText, DeferDeployValue)
        End If

        Return DeferSet
    End Function

    Private Function GetOfferCustomerGroupSQL(ByVal EngineID As Integer) As String
        Dim cgSQL As String = ""

        Select Case EngineID
            Case 0
                cgSQL = "Select OC.LinkID as CustomerGroupID, 0 as ExcludedUsers " &
                        "from OfferConditions OC with (NoLock) " &
                        "inner join CustomerGroups CG with (NoLock) on CG.CustomerGroupID=OC.LinkID " &
                        "where OC.Deleted=0 And OC.ConditionTypeId=1 And CG.Deleted=0 And OC.OfferID = @OfferID " &
                        "and CG.AnyCustomer=0 and CG.AnyCardholder=0 and CG.NewCardholders=0 " &
                        "union " &
                        "Select OC.ExcludedID as CustomerGroupID, 1 as ExcludedUsers " &
                        "from OfferConditions OC with (NoLock) " &
                        "inner join CustomerGroups CG with (NoLock) on CG.CustomerGroupID=OC.ExcludedID " &
                        "where OC.Deleted=0 And OC.ConditionTypeId=1 And CG.Deleted=0 And OC.OfferID = @OfferID " &
                        " and CG.AnyCustomer=0 and CG.AnyCardholder=0 and CG.NewCardholders=0 " &
                        "order by ExcludedUsers desc;"
            Case 2, 9
                cgSQL = "Select ICG.CustomerGroupID, ICG.ExcludedUsers from CPE_IncentiveCustomerGroups ICG with (NoLock) " &
                        "inner join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID = ICG.RewardOptionID " &
                        "inner join CustomerGroups CG with (NoLock) on CG.CustomerGroupID = ICG.CustomerGroupID " &
                        "where ICG.Deleted=0 and CG.Deleted=0 and RO.Deleted=0 and RO.IncentiveID = @OfferID " &
                        " and ICG.CustomerGroupID not in (1,2) and CG.NewCardholders<>1 " &
                        "order by ExcludedUsers desc;"
            Case Else
                cgSQL = ""
        End Select

        Return cgSQL
    End Function

    Private Function GetOfferCustomerGroup(ByVal EngineID As Integer, ByVal LogixID As Long) As Long
        Dim CGID As Long = 0
        Dim dt As DataTable
        Dim rows() As DataRow

        ' find the primary conditional customer group assigned to this offer
        MyCommon.QueryStr = GetOfferCustomerGroupSQL(EngineID)
        MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = LogixID
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        If dt.Rows.Count > 0 Then
            rows = dt.Select("ExcludedUsers=0")
            If rows.Length > 0 Then
                CGID = MyCommon.NZ(rows(0).Item("CustomerGroupID"), 0)
            End If
        End If

        Return CGID
    End Function

    ' checks to determine if all offer components sent in the xml are enabled in the instance of Logix
    Private Function AreValidComponents(ByVal xmlDoc As XmlDocument, ByVal EngineID As Integer, ByRef ErrorMsg As String) As Boolean
        Dim Valid As Boolean = True
        Dim dt As DataTable
        Dim i As Integer = 0
        Dim ComponentTypeID, LinkID As Integer
        Dim CMLinkIDs() As Integer = {1, 2, 1}  ' Customer, Product, Discount
        Dim CPELinkIDs() As Integer = {1, 2, 2} ' Customer, Product, Discount
        Dim LinkIDs(-1) As Integer

        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

        ' find all the disabled offer components
        MyCommon.QueryStr = "select ComponentTypeID, LinkID from PromoEngineComponentTypes with (NoLock) " &
                            "where EngineID = @EngineID and Enabled=0 order by ComponentTypeID"
        MyCommon.DBParameters.Add("@EngineID", SqlDbType.Int).Value = EngineID
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        If dt.Rows.Count > 0 Then

            ' assign the correct link ids based on the engine type
            Select Case EngineID
                Case 0
                    LinkIDs = CMLinkIDs
                Case Else
                    LinkIDs = CPELinkIDs
            End Select

            ' check if the disabled component are present in the offer XML 
            While i < dt.Rows.Count AndAlso Valid
                ComponentTypeID = MyCommon.NZ(dt.Rows(i).Item("ComponentTypeID"), 0)
                LinkID = MyCommon.NZ(dt.Rows(i).Item("LinkID"), 0)

                Select Case ComponentTypeID
                    Case 1 ' Conditions 
                        If LinkID = LinkIDs(0) AndAlso HasComponent(xmlDoc, "//Offer/Conditions/Customer") Then
                            ErrorMsg = "Customer conditions are disabled this instance of Logix."
                            Valid = False
                        ElseIf LinkID = LinkIDs(1) AndAlso HasComponent(xmlDoc, "//Offer/Conditions/Product") Then
                            ErrorMsg = "Product conditions are disabled this instance of Logix."
                            Valid = False
                        End If
                    Case 2 ' Rewards
                        If LinkID = LinkIDs(2) AndAlso HasComponent(xmlDoc, "//Offer/Rewards/Discount") Then
                            ErrorMsg = "Discount rewards are disabled this instance of Logix."
                            Valid = False
                        End If
                End Select
                i += 1
            End While
        End If

        Return Valid
    End Function

    ' determines if a particular xPath exists in the xmlDoc parameter
    Private Function HasComponent(ByVal xmlDoc As XmlDocument, ByVal xPath As String) As Boolean
        Dim Present As Boolean = False
        Dim nList As XmlNodeList = Nothing

        nList = xmlDoc.SelectNodes(xPath)
        If nList IsNot Nothing AndAlso nList.Count > 0 Then
            Present = True
        End If

        Return Present
    End Function

    Private Function IsValidCardType(ByVal CardTypeID As Integer) As Boolean
        Dim ValidType As Boolean = False
        Dim dt As DataTable

        If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

        MyCommon.QueryStr = "select CardTypeID from CardTypes with (NoLock) where CardTypeID = @CardTypeID"
        MyCommon.DBParameters.Add("@CardTypeID", SqlDbType.Int).Value = CardTypeID
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)
        ValidType = (dt.Rows.Count > 0)

        Return ValidType
    End Function

    Private Function IsValidCard(ByVal CardTypeID As Integer, ByVal Card As String) As Boolean
        If (String.IsNullOrEmpty(Card)) Then
            Return False
        End If

        If (CardTypeID = 6) Then
            'Email format check.
            Return Regex.IsMatch(Card, "([\w-+]+(?:\.[\w-+]+)*@(?:[\w-]+\.)+[a-zA-Z]{2,7})")
        End If
        'For all other card types as of now we return true.
        Return True
    End Function

    Private Function HandleOfferCustomers(ByVal ExternalSourceID As String, ByVal CustomerOfferData As String,
                                          ByVal ResponseType As RESPONSE_TYPES, Optional ByVal TreatAsClipData As Boolean = False, Optional ByVal TreatAsCustomers As Boolean = False) As String
        Dim Queued As Boolean = False
        Dim ErrorCode As ERROR_CODES = ERROR_CODES.ERROR_NONE
        Dim ErrorMsg As String = ""
        Dim CustomerPK As Long = 0
        Dim AutoDeploy As Boolean = False
        Dim OfferEngineID As Long = -1
        Dim FileName As String = ""
        Dim QueueData As New CUSTOMERS_QUEUE_DATA
        Dim EngineID As Integer = -1
        Dim FormatFileName As String = ""
        Dim sQuery As String = ""
        Dim dt As DataTable = Nothing
        Dim engineid1 As String = ""
        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            ExtInterfaceID = GetExtInterfaceID(ExternalSourceID, AutoDeploy)
            If (Not TreatAsCustomers) Then
                EngineID = GetEngineID("")
            Else
                Dim clientOfferID1 = CustomerOfferData.Substring(0, CustomerOfferData.IndexOf(","))
                engineid1 = "select EngineId from CPE_Incentives with (NoLock) where ClientOfferID = '" & clientOfferID1 & "' " &
                     "and InboundCRMEngineID=" & ExtInterfaceID & " and Deleted=0;"
                MyCommon.QueryStr = engineid1
                dt = MyCommon.LRT_Select
                If (dt.Rows.Count > 0) Then
                    OfferEngineID = MyCommon.NZ(dt.Rows(0).Item("EngineID"), 0)
                    EngineID = Convert.ToInt32(OfferEngineID)
                End If
            End If

            If (ExtInterfaceID > 0 AndAlso EngineID > -1) Then
                If Not IsValidFormat(CustomerOfferData, FormatFileName, ErrorMsg, TreatAsClipData) Then
                    ErrorCode = ERROR_CODES.ERROR_INVALID_FORMAT
                Else
                    FileName = WriteOfferCustomerFile(CustomerOfferData, ErrorCode, ErrorMsg)
                    If ErrorCode = ERROR_CODES.ERROR_NONE AndAlso FileName.Trim <> "" Then
                        With QueueData
                            .FileName = FileName
                            .ExtInterfaceID = ExtInterfaceID
                            .EngineID = EngineID
                            .ResponseType = ResponseType
                            .FormatFileName = FormatFileName
                            .TreatAsClipData = TreatAsClipData
                        End With
                        QueueOfferCustomers(QueueData, ErrorCode, ErrorMsg)
                    End If
                End If

                Queued = (ErrorCode = ERROR_CODES.ERROR_NONE)
            ElseIf EngineID = -1 Then
                ErrorCode = ERROR_CODES.ERROR_INCORRECT_ENGINE_TYPE
                ErrorMsg = "EngineID could not be determined for this request."
            Else
                ErrorCode = ERROR_CODES.ERROR_XML_BAD_SOURCE_ID
                ErrorMsg = "Unrecognized ExternalSourceID: " & ExternalSourceID
            End If
        Catch ex As Exception
            ErrorCode = ERROR_CODES.ERROR_APPLICATION
            ErrorMsg = "Application Error encountered: " & ex.ToString
        Finally
            Dim method As String = "<unknown method>"
            If ResponseType = RESPONSE_TYPES.ADD_CUSTOMERS Then
                method = "AddCustomers"
            ElseIf ResponseType = RESPONSE_TYPES.REMOVE_CUSTOMERS Then
                method = "RemoveCustomers"
            ElseIf ResponseType = RESPONSE_TYPES.CLIP_BUNDLE Then
                method = "ClipBundle"
            ElseIf ResponseType = RESPONSE_TYPES.UNCLIP_BUNDLE Then
                method = "UnclipBundle"
            End If
            AppendToLog(ExternalSourceID, method, "", "", ErrorCode, ErrorMsg)
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try

        Return WriteOfferXmlResponse(0, "", ErrorCode, ErrorMsg, ResponseType, Queued, FileName)

    End Function

    Private Function IsValidFormat(ByRef str As String, ByRef FormatFileName As String, ByRef ErrorMsg As String, Optional ByVal TreatAsClipData As Boolean = False) As Boolean
        Dim RetVal As Boolean = False
        Dim i As Integer
        Dim c As Char
        Dim prevC As Char = ""
        Dim strUbound As Integer = 0
        Dim commaCt, lineCommaCt As Integer
        Dim firstLine As Boolean = True
        Dim RecDelimited As Boolean = False
        Dim RecCount As Integer = 1

        ErrorMsg = ""

        strUbound = (str.Length - 1)
        For i = 0 To strUbound
            RetVal = False
            RecDelimited = False

            c = str.Chars(i)
            If c = "," Then
                commaCt += 1
                ' if first line then determine how many commas should be on each line
                If firstLine Then lineCommaCt += 1
            ElseIf c = "|" Then
                If commaCt <> lineCommaCt Then
                    ErrorMsg = "Record number " & RecCount & " has an incorrect number of fields."
                End If

                ' reset the comma count for the next line
                commaCt = 0
                RecCount += 1
                firstLine = False
                RecDelimited = True
                RetVal = True
            End If
            prevC = c
        Next

        ' account for the case that the last line may not have a record delimiter
        If Not RecDelimited Then
            If commaCt = lineCommaCt Then
                RetVal = True
            Else
                RetVal = False
                ErrorMsg = "Record number " & RecCount & " has an incorrect number of fields."
            End If
        End If

        Select Case lineCommaCt
            Case 1
                FormatFileName = "UsersOffer.fmt"
            Case 2
                FormatFileName = IIf(TreatAsClipData, "UsersClipOffer.fmt", "UsersTypeOffer.fmt")
            Case Else
                FormatFileName = ""
        End Select

        Return RetVal
    End Function

    Private Function IsValidBannerAssignment(ByVal OfferXMLDoc As XmlDocument, ByVal EngineID As Integer,
                                             ByRef ErrorMsg As String) As Boolean
        Dim ValidBanner As Boolean = True
        Dim BannerIDs As Long()

        If MyCommon.Fetch_SystemOption(66) = "1" Then
            'BannerIDs = LoadAllOfferBanners(OfferXMLDoc, LogixID)
            BannerIDs = GetBannerIDs(OfferXMLDoc)

            ' check to see if multiple banners for offers is enabled when more than one is sent
            If MyCommon.Fetch_SystemOption(67) <> "1" AndAlso BannerIDs.Length > 1 Then
                ValidBanner = False
                ErrorMsg = "Multiple banner assignments to an offer are not permitted."
            Else
                ValidBanner = AreValidBanners(OfferXMLDoc, EngineID, True, ErrorMsg)
            End If
        End If

        Return ValidBanner
    End Function

    Private Function AreValidBanners(ByVal OfferXMLDoc As XmlDocument, ByVal EngineID As Integer,
                                     ByVal bValidateAllBanners As Boolean, ByRef ErrorMsg As String) As Boolean
        Dim bValidBanners As Boolean = True
        Dim BannerID As Long
        Dim ValidBannerID As Long
        Dim BannerNode As XmlNode = Nothing
        Dim NameNodes, ExtIdNodes, IdNodes As XmlNodeList
        Dim i As Integer
        Dim sTempValue As String
        Dim bAllBanners As Boolean = False

        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

        If MyCommon.Fetch_SystemOption(66) = "1" Then
            BannerNode = OfferXMLDoc.SelectSingleNode("//Offer/Banners")
            If BannerNode IsNot Nothing Then
                IdNodes = BannerNode.SelectNodes("//Offer/Banners/BannerID")
                If IdNodes IsNot Nothing AndAlso IdNodes.Count > 0 Then
                    For i = 0 To IdNodes.Count - 1
                        Long.TryParse(IdNodes.Item(i).InnerText, BannerID)
                        ValidBannerID = GetValidBannerID("BannerID", BannerID.ToString(), EngineID, bAllBanners)
                        If ValidBannerID = 0 Then
                            bValidBanners = False
                            ErrorMsg = "Banner with BannerID = " & BannerID & " does not exist!"
                            Exit For
                        Else
                            If bValidateAllBanners AndAlso bAllBanners AndAlso IdNodes.Count > 1 Then
                                bValidBanners = False
                                ErrorMsg = "Banner with BannerID = " & BannerID & " has 'All Banners' set, so multiple banners are not allowed!"
                                Exit For
                            End If
                        End If
                    Next
                Else
                    ExtIdNodes = BannerNode.SelectNodes("//Offer/Banners/ExtBannerID")
                    If (ExtIdNodes IsNot Nothing AndAlso ExtIdNodes.Count > 0) Then
                        For i = 0 To ExtIdNodes.Count - 1
                            sTempValue = ExtIdNodes(i).InnerText.Trim
                            If (sTempValue <> "") Then
                                ValidBannerID = GetValidBannerID("ExtBannerID", sTempValue, EngineID, bAllBanners)
                                If ValidBannerID = 0 Then
                                    bValidBanners = False
                                    ErrorMsg = "Banner with ExtBannerID = " & sTempValue & " does not exist!"
                                    Exit For
                                End If
                            Else
                                If bValidateAllBanners AndAlso bAllBanners AndAlso IdNodes.Count > 1 Then
                                    bValidBanners = False
                                    ErrorMsg = "Banner with ExtBannerID = " & sTempValue & " has 'All Banners' set, so multiple banners are not allowed!"
                                    Exit For
                                End If
                            End If
                        Next
                    Else
                        NameNodes = BannerNode.SelectNodes("//Offer/Banners/Name")
                        If NameNodes IsNot Nothing AndAlso NameNodes.Count > 0 Then
                            For i = 0 To NameNodes.Count - 1
                                sTempValue = NameNodes(i).InnerText.Trim
                                If (sTempValue <> "") Then
                                    ValidBannerID = GetValidBannerID("Name", sTempValue, EngineID, bAllBanners)
                                    If ValidBannerID = 0 Then
                                        bValidBanners = False
                                        ErrorMsg = "Banner with Name = " & sTempValue & " does not exist!"
                                        Exit For
                                    Else
                                        If bValidateAllBanners AndAlso bAllBanners AndAlso IdNodes.Count > 1 Then
                                            bValidBanners = False
                                            ErrorMsg = "Banner with Name = " & sTempValue & " has 'All Banners' set, so multiple banners are not allowed!"
                                            Exit For
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If
            Else
                bValidBanners = False
                ErrorMsg = "You must select a banner for the offer."
            End If
        End If

        Return bValidBanners
    End Function

    Private Function GetValidBannerID(ByVal sSearchColumn As String, ByVal sSearchText As String, ByVal EngineID As Integer, ByRef bAllBanners As Boolean) As Long
        Dim lValidBannerID As Long = 0
        Dim bAddSingleQuotes As Boolean = True
        Dim sWhere As String
        Dim dt As DataTable

        bAllBanners = False
        If MyCommon.Fetch_SystemOption(66) = "1" Then
            If sSearchColumn.ToLower() = "bannerid" Then
                sWhere = "where B.Deleted=0 and B.BannerID = @BannerID and BE.EngineID = @EngineID "
            Else
                sWhere = "where B.Deleted=0 and B." & sSearchColumn & "= @BannerID and BE.EngineID = @EngineID "
            End If
            MyCommon.QueryStr = "select B.BannerID, B.AllBanners from Banners B with (NoLock) " &
                                "inner join BannerEngines BE with (NoLock) on BE.BannerID=B.BannerID " & sWhere
            MyCommon.DBParameters.Add("@BannerID", SqlDbType.NVarChar).Value = sSearchText
            MyCommon.DBParameters.Add("@EngineID", SqlDbType.Int).Value = EngineID
            dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            If (dt.Rows.Count > 0) Then
                lValidBannerID = MyCommon.NZ(dt.Rows(0).Item("BannerID"), 0)
                bAllBanners = MyCommon.NZ(dt.Rows(0).Item("AllBanners"), False)
            End If
        End If

        Return lValidBannerID
    End Function

    Private Function LoadAllOfferBanners(ByVal OfferXMLDoc As XmlDocument, ByVal LogixID As Long) As Long()
        Dim BannerIDs(-1) As Long
        Dim BannerID As Long = -1
        Dim BannerList As String = ""
        Dim dt As DataTable
        Dim i As Integer

        BannerIDs = GetBannerIDs(OfferXMLDoc)

        ' load up existing banners for this offer
        If LogixID > 0 Then
            MyCommon.QueryStr = "select distinct BannerID from BannerOffers with (NoLock) " &
                                "where OfferID = @OfferID"
            MyCommon.DBParameters.Add("@OfferID", SqlDbType.Int).Value = LogixID
            If BannerIDs.Length > 0 Then
                BannerList = ConvertToCSV(BannerIDs)
                If BannerList <> "" Then MyCommon.QueryStr &= " and BannerID not in (" & BannerList & ");"
            End If

            dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            If dt.Rows.Count > 0 Then
                ReDim Preserve BannerIDs(BannerIDs.Length + dt.Rows.Count - 1)
                For i = 0 To dt.Rows.Count - 1
                    BannerIDs(i) = MyCommon.NZ(dt.Rows(i).Item("BannerID"), -1)
                Next
            End If
        End If

        Return BannerIDs
    End Function

    Private Function AreValidBannerLocations(ByVal OfferXMLDoc As XmlDocument, ByVal EngineID As Integer,
                                             ByRef ErrorMsg As String) As Boolean
        Dim Valid As Boolean = True
        Dim BannerIDs(-1) As Long
        Dim BannerID As Long
        Dim BannerList As String = ""
        Dim LocGroupNodes, LocListNodes As XmlNodeList
        Dim LocNode As XmlNode
        Dim LocationGroupID As Long
        Dim StoreList As String = ""
        Dim Locations(-1) As String
        Dim Fields(-1) As String
        Dim LocationID As Long
        Dim dt As DataTable
        Dim i As Integer
        Dim ExtLocCode As String = ""
        Dim LocName As String = ""
        Dim CharsToTrim() As Char = {Chr(0), Chr(8), Chr(9), Chr(10),
                                     Chr(13), Chr(32)}

        If MyCommon.Fetch_SystemOption(66) = "1" Then
            'BannerIDs = LoadAllOfferBanners(OfferXMLDoc, LogixID)
            BannerIDs = GetBannerIDs(OfferXMLDoc)
            BannerList = ExcludeAllBannersConvertToCSV(BannerIDs)

            LocGroupNodes = OfferXMLDoc.SelectNodes("//Offer/Stores/StoreGroupID")
            LocListNodes = OfferXMLDoc.SelectNodes("//Offer/Stores/StoreList")

            If OfferXMLDoc IsNot Nothing Then
                ' validate any StoreGroupID tags
                If LocGroupNodes IsNot Nothing Then
                    If BannerList <> "" Then
                        For Each LocNode In LocGroupNodes
                            Long.TryParse(LocNode.InnerText, LocationGroupID)
                            ' All Locations is always ok
                            If LocationGroupID <> 1 Then
                                MyCommon.QueryStr = "select BannerID from LocationGroups with (NoLock) " &
                                                    "where LocationGroupID = @LocationGroupID" &
                                                    " and BannerID in (select items from SPLIT(@BannerID,','))"
                                MyCommon.DBParameters.Add("@BannerID", SqlDbType.NVarChar).Value = BannerList
                                MyCommon.DBParameters.Add("@LocationGroupID", SqlDbType.BigInt).Value = LocationGroupID
                                dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                                If dt.Rows.Count = 0 Then
                                    ErrorMsg = "StoreGroupID " & LocationGroupID & " does not belong " &
                                                   "to one of the offer's assigned banner" & IIf(BannerIDs.Length = 1, "", "s") & " " &
                                                   "(ID" & IIf(BannerIDs.Length = 1, "", "s") & "=" & BannerList & ")"
                                    Return False
                                End If
                            End If
                        Next
                    End If
                End If

                ' validate any StoreList tags
                If LocListNodes IsNot Nothing Then
                    For Each LocNode In LocListNodes
                        StoreList = LocNode.InnerText
                        If StoreList <> "" Then
                            StoreList = StoreList.Replace(vbCrLf, vbLf)
                            StoreList = StoreList.Replace(vbLf, vbCrLf)
                            Locations = StoreList.Split(ControlChars.CrLf)
                            For i = 0 To Locations.GetUpperBound(0)
                                Fields = Locations(i).Split(",")
                                If Fields.Length = 2 Then
                                    ExtLocCode = Fields(0).Trim(CharsToTrim)
                                    LocName = Fields(1).Trim(CharsToTrim)
                                ElseIf Fields.Length > 2 Then
                                    ExtLocCode = ""
                                    LocName = ""
                                End If

                                If ExtLocCode <> "" AndAlso LocName <> "" Then
                                    ' determine if the location already exists
                                    MyCommon.QueryStr = "select LocationID, BannerID from Locations with (NoLock) " &
                                                        "where Deleted=0 and ExtLocationCode = @ExtLocationCode "
                                    MyCommon.DBParameters.Add("@ExtLocationCode", SqlDbType.NVarChar).Value = ExtLocCode
                                    dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                                    If dt.Rows.Count > 0 Then
                                        LocationID = MyCommon.NZ(dt.Rows(0).Item("LocationID"), 0)
                                        BannerID = MyCommon.NZ(dt.Rows(0).Item("BannerID"), -1)
                                        If Not IsValueInArray(BannerIDs, BannerID) Then
                                            ErrorMsg = "Location " & ExtLocCode & " with name " & LocName & " (LocationID=" & LocationID & " ) " &
                                                       "is " & IIf(BannerID = -1, " unassigned to a banner ", "assigned to BannerID " & BannerID) & ". " &
                                                       "Locations should belong to the offer's assigned banner" & IIf(BannerIDs.Length = 1, "", "s") &
                                                       " (ID=" & IIf(BannerIDs.Length = 1, "", "s") & BannerList & ")"
                                            Return False
                                        End If
                                    End If
                                End If
                            Next

                        End If
                    Next
                End If
            End If

        End If

        Return Valid
    End Function

    Private Function IsValueInArray(ByVal arr As Long(), ByVal value As Long) As Boolean
        Dim Exists As Boolean = False

        If arr IsNot Nothing AndAlso arr.Length > 0 Then
            For Each checkLong As Long In arr
                If value = checkLong Then
                    Exists = True
                    Exit For
                End If
            Next
        End If

        Return Exists
    End Function

    Private Function WriteOfferCustomerFile(ByVal CustomerOfferData As String, ByRef ErrorCode As ERROR_CODES,
                                            ByRef ErrorMessage As String) As String
        Dim FileName As String = ""
        Dim SQLBulkPath As String = ""

        SQLBulkPath = MyCommon.Fetch_SystemOption(29)
        If SQLBulkPath.Trim = "" Then
            ErrorCode = ERROR_CODES.ERROR_ADD_CUSTOMERS_FAILED
            ErrorMessage &= "Workspace path file is empty. A valid path is needed to create a customer file for upload."
        Else
            If Not (SQLBulkPath.Substring(SQLBulkPath.Length - 1, 1) = "\") Then SQLBulkPath = SQLBulkPath & "\"
            FileName = "OC_" & Guid.NewGuid.ToString & ".txt"
            Try
                File.WriteAllText(SQLBulkPath & FileName, CustomerOfferData)
            Catch ex As Exception
                FileName = ""
                ErrorCode = ERROR_CODES.ERROR_ADD_CUSTOMERS_FAILED
                ErrorMessage = ex.ToString
            End Try
        End If

        Return FileName
    End Function

    Private Function QueueOfferCustomers(ByVal QueueData As CUSTOMERS_QUEUE_DATA,
                                         ByRef ErrorCode As ERROR_CODES, ByVal ErrorMsg As String) As Boolean
        Dim Queued As Boolean
        Dim OperationType As Integer
        Dim TreatAsClipData As Boolean = False
        If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

        TreatAsClipData = IIf(QueueData.TreatAsClipData, True, False)
        OperationType = IIf((QueueData.ResponseType = RESPONSE_TYPES.ADD_CUSTOMERS Or QueueData.ResponseType = RESPONSE_TYPES.CLIP_BUNDLE), 1, 2)
        MyCommon.QueryStr = "Insert into OfferCustomerInsertQueue (FileName, UploadTime, StatusFlag, " &
                            " ExtInterfaceID, EngineID, OperationType, FormatFileName, TreatAsClipData) " &
                            " values (@FileName, getdate(), 0, @ExtInterfaceID, " &
                            "@EngineID, @OperationType, @FormatFileName, " &
                            "@TreatAsClipData)"
        MyCommon.DBParameters.Add("@FileName", SqlDbType.NVarChar).Value = (QueueData.FileName).ConvertBlankIfNothing
        MyCommon.DBParameters.Add("@ExtInterfaceID", SqlDbType.Int).Value = QueueData.ExtInterfaceID
        MyCommon.DBParameters.Add("@EngineID", SqlDbType.Int).Value = QueueData.EngineID
        MyCommon.DBParameters.Add("@OperationType", SqlDbType.Int).Value = OperationType
        MyCommon.DBParameters.Add("@FormatFileName", SqlDbType.NVarChar).Value = (QueueData.FormatFileName).ConvertBlankIfNothing
        MyCommon.DBParameters.Add("@TreatAsClipData", SqlDbType.Bit).Value = IIf(TreatAsClipData, 1, 0)
        MyCommon.ExecuteNonQuery(Copient.DataBases.LogixXS)

        Queued = (MyCommon.RowsAffected = 1)
        If Not Queued Then
            If (QueueData.ResponseType = RESPONSE_TYPES.CLIP_BUNDLE) Then
                ErrorCode = IIf(OperationType = 1, ERROR_CODES.ERROR_CLIP_BUNDLE_FAILED, ERROR_CODES.ERROR_UNCLIP_BUNDLE_FAILED)
            ElseIf (QueueData.ResponseType = RESPONSE_TYPES.ADD_CUSTOMERS) Then
                ErrorCode = IIf(OperationType = 1, ERROR_CODES.ERROR_ADD_CUSTOMERS_FAILED, ERROR_CODES.ERROR_REMOVE_CUSTOMERS_FAILED)
            End If
            ErrorMsg = "Failed to write file " & QueueData.FileName & " to the OfferCustomerInsertQueue for processing"
        End If

        Return Queued
    End Function

    Private Function IsStringNumeric(ByVal str As String) As Boolean
        Dim NumericOnly As Boolean = False
        Dim i As Integer
        Dim strUbound As Integer

        If str IsNot Nothing Then
            strUbound = (str.Length - 1)
            For i = 0 To strUbound
                NumericOnly = Char.IsDigit(str.Chars(i))
                If Not NumericOnly Then Exit For
            Next
        End If

        Return NumericOnly
    End Function

    Private Function IsValidCRMEngine(ByVal OfferXmlDoc As XmlDocument, ByVal EngineID As Integer, ByRef ErrorMsg As String) As Boolean
        Dim ValidEngine As Boolean = True
        Dim CRMEngineExtCode As String = ""
        Dim dt As DataTable

        TryParseElementValue(OfferXmlDoc, "//Offer/CRMEngineExtCode", CRMEngineExtCode)

        If CRMEngineExtCode.Trim <> "" Then
            MyCommon.QueryStr = "select ExtInterfaceID from ExtCRMInterfaces with (NoLock) " &
                                "where Deleted=0 and Active=1 and OutboundEnabled=1 and ExtCode = @ExtCode"
            MyCommon.DBParameters.Add("@ExtCode", SqlDbType.NVarChar).Value = CRMEngineExtCode.ConvertBlankIfNothing
            dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            ValidEngine = (dt.Rows.Count > 0)
            If Not ValidEngine Then
                ErrorMsg = "CRM engine '" & CRMEngineExtCode & "' is invalid"
            End If
        End If

        Return ValidEngine
    End Function

    Private Function IsAnyCustomerOffer(ByVal OfferXmlDoc As XmlDocument, ByVal EngineID As Integer) As Boolean
        Dim AnyCustOffer As Boolean = False
        Dim TempStr As String = ""
        Dim CGID As Long = 0

        If TryParseElementValue(OfferXmlDoc, "//Offer/Conditions/Customer/CustomerGroupID", TempStr) Then
            If TempStr IsNot Nothing AndAlso Long.TryParse(TempStr, CGID) AndAlso CGID = 1 Then
                AnyCustOffer = True
            End If
        End If

        Return AnyCustOffer
    End Function
    Private Function IsValidProductCondition(ByVal OfferXmlDoc As XmlDocument, ByVal EngineID As Integer, ByRef ErrorMsg As String) As Boolean
        Dim Valid As Boolean = True
        Dim ProductGroupNodes As XmlNodeList
        Dim ProdNode As XmlNode
        Dim Quantity As Decimal
        If (EngineID = Copient.CommonInc.InstalledEngines.CPE OrElse EngineID = Copient.CommonInc.InstalledEngines.UE) Then
            If OfferXmlDoc IsNot Nothing Then
                ProductGroupNodes = OfferXmlDoc.SelectNodes("//Offer/Conditions/Product/ProductGroupID")
                If (ProductGroupNodes IsNot Nothing AndAlso ProductGroupNodes.Count > 0) Then
                    If ProductGroupNodes IsNot Nothing Then
                        For Each ProdNode In ProductGroupNodes
                            If ProdNode IsNot Nothing Then
                                If ProdNode.Attributes("quantity") IsNot Nothing Then
                                    Decimal.TryParse(ProdNode.Attributes("quantity").InnerText, Quantity)
                                    If (Quantity <= 0) Then
                                        Valid = False
                                        ErrorMsg = "Invalid product condition quantity value."
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
            End If
        End If
        Return Valid
    End Function

    Private Function IsValidCustomerApprovalCondition(ByVal offerXmlDoc As XmlDocument, ByVal engineID As Integer, ByRef errMsg As String) As Boolean

        Dim custAppNode As XmlNode = offerXmlDoc.SelectSingleNode("//Offer/Conditions/Customer/CustomerApproval")
        Dim mesgDescNode, appTypeNode As XmlNode

        If offerXmlDoc IsNot Nothing AndAlso custAppNode IsNot Nothing Then
            If (engineID = InstalledEngines.CPE) Then
                errMsg = "Customer Approval condition is not for Engine type : CPE"
                Return False
            ElseIf (engineID = InstalledEngines.UE) Then

                mesgDescNode = offerXmlDoc.SelectSingleNode("//Offer/Conditions/Customer/CustomerApproval/MessageDescription")
                appTypeNode = offerXmlDoc.SelectSingleNode("//Offer/Conditions/Customer/CustomerApproval/ApprovalType")

                If mesgDescNode IsNot Nothing Then
                    If mesgDescNode.InnerText = "" Then
                        errMsg = "Message Description for Customer Approval cannot be blank."
                        Return False
                    End If
                End If
                If appTypeNode IsNot Nothing Then
                    If appTypeNode.InnerText < 1 OrElse appTypeNode.InnerText > 2 Then 'Approval Type currently has two values : 1-Once for offer 2-each offer redemption
                        errMsg = "Invalid Approval Type for Customer Approval Condition."
                        Return False
                    End If
                End If

            End If
        End If

        Return True

    End Function

    Private Function IsValidOfferType(ByVal OfferXmlDoc As XmlDocument, ByVal EngineID As Integer, ByRef ErrorMsg As String) As Boolean
        Dim Valid As Boolean = True
        Dim manuCoupon As Boolean = False
        Dim storeCoupon As Boolean = False
        Dim manuCouponNode As XmlNode = Nothing
        Dim storeCouponNode As XmlNode = Nothing

        If (EngineID = Copient.CommonInc.InstalledEngines.UE) Then
            If OfferXmlDoc IsNot Nothing Then
                manuCouponNode = OfferXmlDoc.SelectSingleNode("//Offer/IsManufacturerCoupon")
                storeCouponNode = OfferXmlDoc.SelectSingleNode("//Offer/IsStoreCoupon")
                If manuCouponNode IsNot Nothing AndAlso storeCouponNode IsNot Nothing Then
                    Boolean.TryParse(manuCouponNode.InnerText, manuCoupon)
                    Boolean.TryParse(storeCouponNode.InnerText, storeCoupon)
                    If manuCoupon = True AndAlso storeCoupon = True Then
                        ErrorMsg = "Offer can be of any one type - Standard Offer, Store Coupon or Manufacturer Coupon."
                        Valid = False
                    End If
                End If
            End If
        End If

        Return Valid
    End Function


    Private Function IsValidAnyCustomerOffer(ByVal OfferXmlDoc As XmlDocument, ByVal EngineID As Integer, ByRef ErrorMsg As String) As Boolean
        Dim Valid As Boolean = True
        Dim TempStr As String = ""
        Dim Frequency As String = ""
        Dim CustomLimitNode As XmlNode
        Dim Period As Integer = 0
        Dim PeriodType As Integer = 0

        If (EngineID = Copient.CommonInc.InstalledEngines.CPE OrElse EngineID = Copient.CommonInc.InstalledEngines.UE) AndAlso IsAnyCustomerOffer(OfferXmlDoc, EngineID) Then
            ' first, check that report impressions is turned off
            TryParseElementValue(OfferXmlDoc, "//Offer/ReportImpressions", TempStr)
            If TempStr = "1" Then TempStr = "TRUE"

            If TempStr IsNot Nothing AndAlso TempStr.ToUpper = "TRUE" Then
                Valid = False
                ErrorMsg = "Report Impressions is an invalid setting for an Any Customer offer."
            End If

            ' then, check the limit
            If TryParseElementValue(OfferXmlDoc, "//Offer/Limits/Frequency", Frequency) Then
                Select Case Frequency
                    Case "NO_LIMIT", "ONCE_PER_TRANSACTION"
                        ' these are the only allowable values for an Any Customer Offer
                    Case Else
                        Valid = False
                        ErrorMsg = "Invalid Limit Frequency tag value for an Any Customer offer.  NO_LIMIT and ONCE_PER_TRANSACTION are the only valid values."
                End Select
            Else
                CustomLimitNode = OfferXmlDoc.SelectSingleNode("//Offer/Limits/CustomLimit")
                If CustomLimitNode IsNot Nothing Then
                    TryParseElementValue(OfferXmlDoc, "//Offer/Limits/CustomLimit/Period", TempStr)
                    Integer.TryParse(TempStr, Period)

                    TryParseElementValue(OfferXmlDoc, "//Offer/Limits/CustomLimit/PeriodType", TempStr)
                    Select Case TempStr
                        Case "DAYS_SINCE_INCENTIVE_START"
                            PeriodType = 1
                        Case "HOURS_SINCE_LAST_AWARDED"
                            PeriodType = 2
                        Case Else ' default to hours
                            PeriodType = 2
                    End Select

                    If PeriodType = 1 OrElse (PeriodType = 2 AndAlso Period > 1) Then
                        Valid = False
                        ErrorMsg = "Invalid CustomLimit tag value for an Any Customer offer. PeriodType should not equal " &
                                   "either DAYS_SINCE_INCENTIVE_START or HOURS_SINCE_LAST_AWARDED with a Period greater than 1."
                    End If
                End If
            End If

        End If

        Return Valid
    End Function

    Private Function IsOfferExpiredAndLocked(ByVal LogixOfferID As Integer, ByVal EngineID As Integer, ByRef ErrorMsg As String) As Boolean
        Dim Valid As Boolean = False
        Dim dtIncentives As DataTable
        Dim IsLockedEnable As Boolean
        Dim EndDate As Date
        Dim IsExpireLocked As Boolean
        If (EngineID = Copient.CommonInc.InstalledEngines.CPE) Then
            IsLockedEnable = (MyCommon.Extract_Val(MyCommon.Fetch_CPE_SystemOption(80)) = 1)
        ElseIf (EngineID = Copient.CommonInc.InstalledEngines.UE) Then
            IsLockedEnable = (MyCommon.Extract_Val(MyCommon.Fetch_UE_SystemOption(80)) = 1)
        End If
        MyCommon.QueryStr = "select StartDate, EndDate, TestingStartDate, TestingEndDate, UpdateLevel, EngineID,ExpireLocked  from CPE_Incentives with (NoLock) where IncentiveID=" & LogixOfferID
        dtIncentives = MyCommon.LRT_Select()
        If (dtIncentives.Rows.Count > 0) Then
            If (IsDBNull(dtIncentives.Rows(0).Item("EndDate"))) Then
                EndDate = Date.Parse("1900-01-01")
            Else
                EndDate = CDate(dtIncentives.Rows(0).Item("EndDate"))
            End If
            If (IsDBNull(dtIncentives.Rows(0).Item("ExpireLocked"))) Then
                IsExpireLocked = False
            Else
                IsExpireLocked = dtIncentives.Rows(0).Item("ExpireLocked")
            End If
        End If
        If IsLockedEnable AndAlso IsExpireLocked AndAlso EndDate < Now Then
            Valid = True
            ErrorMsg = "The offer has expired and cannot be updated."
        End If
        Return Valid
    End Function
    Private Function ConvertToCSV(ByVal values As Long()) As String
        Dim csv As String = ""
        Dim i As Integer

        If values IsNot Nothing And values.Length > 0 Then
            For i = 0 To values.GetUpperBound(0)
                If i > 0 Then csv &= ","
                csv &= values(i).ToString
            Next
        End If

        Return csv
    End Function

    Private Function ExcludeAllBannersConvertToCSV(ByVal BannerIds As Long()) As String
        Dim csv As String = ""
        Dim i As Integer
        Dim dt As DataTable

        If BannerIds IsNot Nothing And BannerIds.Length > 0 Then
            For i = 0 To BannerIds.GetUpperBound(0)
                ' get list of banners that are flagged as "All Banners"
                MyCommon.QueryStr = "select BannerID from Banners with (NoLock) where Deleted=0 and AllBanners=0 and BannerId = @BannerId"
                MyCommon.DBParameters.Add("@BannerId", SqlDbType.Int).Value = BannerIds(i)
                dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                If dt.Rows.Count > 0 Then
                    If i > 0 Then csv &= ","
                    csv &= BannerIds(i).ToString
                End If
            Next
        End If

        Return csv
    End Function

    Private Function ValidateCustId(ByRef CustomerIDs As String, ByVal ExternalSourceID As String, ByVal ClientOfferID As String,
                                      ByRef ErrorCode As ERROR_CODES, ByRef ErrorMsg As String, ByRef Log As Boolean) As String
        Dim CustIdTypeArray As String()
        Dim CustId As String = String.Empty
        Dim CardType As String = String.Empty
        Dim CardTypeIDint As Integer
        Dim CustomerPK As Long = 0
        Dim AllCustIds As New StringBuilder
        Dim TotCust As Integer = 0
        Dim CustCount As Integer = 0
        Dim ErrMsg As String = String.Empty
        Dim result As String = String.Empty
        Dim ErrorMsgLog As String = String.Empty
        Dim ErrorMsgLogUnmasked As String = String.Empty

        CustIdTypeArray = CustomerIDs.Split("|")
        TotCust = CustIdTypeArray.Length
        For Each Custitem In CustIdTypeArray
            If Custitem.Contains(",") Then
                CustId = Mid(Custitem, 1, InStr(Custitem, ",") - 1)
                CardType = Mid(Custitem, InStr(Custitem, ",") + 1, Custitem.Length)
            Else
                CustId = Custitem
                CardType = GetDefaultCardTypeID()
            End If
            CustId = MyCommon.Pad_ExtCardID(CustId, CardType)
            If (Integer.TryParse(CardType, CardTypeIDint)) Then
                CustomerPK = GetCustomerPK(CustId, CardTypeIDint, False)
                If (CustomerPK > 0) Then
                    AllCustIds.Append(Custitem)
                    AllCustIds.Append(ControlChars.CrLf)
                    CustCount += 1
                ElseIf CustomerPK = -2 Then
                    ErrorCode = ERROR_CODES.ERROR_INVALID_CUSTOMER_ID
                    ErrorMsg = "Customer: " & CustId & " is not in valid format."
                    ErrorMsgLog = "Customer: " & Copient.MaskHelper.MaskCard(CustId, -1) & " is not in valid format."
                    ErrorMsgLogUnmasked = "Customer: " & CustId & " is not in valid format."
                Else
                    If Not String.IsNullOrEmpty(ErrMsg) Then
                        ErrMsg = ErrMsg & ", " & CustId
                        ErrorMsgLog = ErrorMsgLog & ", " & Copient.MaskHelper.Mask(CustId, CardType)
                        ErrorMsgLogUnmasked = ErrorMsgLogUnmasked & ", " & CustId
                    Else
                        ErrMsg = CustId
                        ErrorMsgLog = Copient.MaskHelper.Mask(CustId, CardType)
                        ErrorMsgLogUnmasked = CustId
                    End If
                End If
            Else
                If Not String.IsNullOrEmpty(ErrMsg) Then
                    ErrMsg = ErrMsg & ", " & CustId
                    ErrorMsgLog = ErrorMsgLog & ", " & Copient.MaskHelper.Mask(CustId, CardType)
                    ErrorMsgLogUnmasked = ErrorMsgLogUnmasked & ", " & CustId
                Else
                    ErrMsg = CustId
                    ErrorMsgLog = Copient.MaskHelper.Mask(CustId, CardType)
                End If
            End If
        Next

        If TotCust = CustCount Then
        ElseIf CustCount = 0 Then
            ErrMsg = vbNewLine & "Customerid(s) are listed: " & ErrMsg
            AppendToLog(ExternalSourceID, "RemoveCustomersFromOffer", "", "ClientOfferID=" & ClientOfferID, ERROR_CODES.ERROR_INVALID_CUSTOMER_ID, ErrorMsgLog)
            AppendToLog(ExternalSourceID, "RemoveCustomersFromOffer", "", "ClientOfferID=" & ClientOfferID, ERROR_CODES.ERROR_INVALID_CUSTOMER_ID, ErrorMsgLogUnmasked, True)
            Log = True
        Else
            ErrMsg = vbNewLine & "Customerid(s) are listed: " & ErrMsg
            AppendToLog(ExternalSourceID, "RemoveCustomersFromOffer", "", "ClientOfferID=" & ClientOfferID, ERROR_CODES.ERROR_INVALID_CUSTOMER_ID, ErrorMsgLog)
            AppendToLog(ExternalSourceID, "RemoveCustomersFromOffer", "", "ClientOfferID=" & ClientOfferID, ERROR_CODES.ERROR_INVALID_CUSTOMER_ID, ErrorMsgLogUnmasked, True)
        End If

        If Not AllCustIds Is Nothing Then
            result = AllCustIds.ToString.Trim
        End If

        Return result
    End Function

    Private Function CheckOfferName(ByVal InString As String, Optional ByVal AdditionalValidCharacters As String = "") As String
        Dim tmpString As String = ""
        Dim z As Integer
        Dim n1, n2 As Integer

        If InString IsNot Nothing Then
            ' Remove unicode characters in <Name> if any
            n1 = InStr(1, InString, "<Name>", CompareMethod.Text)
            If n1 > 0 Then
                n1 = n1 + 5
                n2 = InStr(n1, InString, "</Name>", CompareMethod.Text)
                n2 = n2 - 1
                If (n2 > 0 And n2 > n1) Then
                    tmpString = InString.Substring(0, n1)
                    'Copient.Logger.Write_Log(AcceptanceLogFile, n1.ToString & "   " & n2.ToString, True)
                    For z = n1 To n2 - 1
                        If (InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789.#$*-&%@!?/:;+() " & AdditionalValidCharacters, InString(z))) Then
                            tmpString = tmpString & InString(z)
                        End If
                    Next
                    tmpString = tmpString & InString.Substring(n2, (InString.Length - n2))

                End If
            End If

        End If

        CheckOfferName = tmpString
    End Function

    Private Function CheckOfferDescription(ByVal InString As String, Optional ByVal AdditionalValidCharacters As String = "") As String
        Dim tmpString As String = ""
        Dim z As Integer
        Dim n1, n2 As Integer

        If InString IsNot Nothing Then
            ' Remove unicode characters in <Description> if any
            n1 = InStr(1, InString, "<Description>", CompareMethod.Text)
            If n1 > 0 Then
                n1 = n1 + 12
                n2 = InStr(n1, InString, "</Description>", CompareMethod.Text)
                n2 = n2 - 1
                If (n2 > 0 And n2 > n1) Then
                    tmpString = InString.Substring(0, n1)
                    'Copient.Logger.Write_Log(AcceptanceLogFile, n1.ToString & "   " & n2.ToString, True)
                    For z = n1 To n2 - 1
                        If (InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789.#$*-&%@!?/:;+ " & AdditionalValidCharacters, InString(z))) Then
                            tmpString = tmpString & InString(z)
                        End If
                    Next
                    tmpString = tmpString & InString.Substring(n2, (InString.Length - n2))

                End If
            End If

        End If
        If (String.IsNullOrEmpty(tmpString)) Then
            'Offer description is empty.
            tmpString = InString
        End If
        CheckOfferDescription = tmpString
    End Function

    Private Function IsValidLocationCurrencyMatch(ByVal OfferID As Integer, ByVal EngineID As Integer, ByRef ErrorMsg As String) As Boolean
        Dim Valid As Boolean = True
        Dim TempVal As String = ""
        Dim dst As DataTable
        Dim ErrorTextList As String = ""
        If (EngineID = Copient.CommonInc.InstalledEngines.UE) Then

            'If Fetch_UE_SystemOption(136) = "1" Then 'is multi-currency enabled?  no need to run this query if there has already been a violation of the requirements
            'query to see if there are any locations joined to the offer that use a currencyID that is different than RewardOptions.CurrencyID 

            MyCommon.QueryStr = "select L.LocationID, isnull(L.LocationName, '') as LocationName ,L.CurrencyID " &
                               "  from Locations as L with (NoLock) Inner Join LocGroupItems as LGI with (NoLock) on L.LocationID=LGI.LocationID and LGI.Deleted=0 and L.Deleted=0 " &
                               " Inner Join OfferLocations as OL with (NoLock) on OL.LocationGroupID=LGI.LocationGroupID and OL.Deleted=0 " &
                               " Inner Join CPE_RewardOptions as RO with (NoLock) on RO.IncentiveID=OL.OfferID and RO.Deleted=0 and RO.TouchResponse=0 " &
                               " Where  L.CurrencyID<>RO.CurrencyID and RO.CurrencyID>0 and RO.IncentiveID = @IncentiveID "
            MyCommon.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = OfferID
            dst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            If dst.Rows.Count > 0 Then
                ErrorMsg = "The following targeted locations do not support the currency selected for this offer: " 'The following targeted locations do not support the currency selected for this offer: 
                ErrorTextList = ""
                Valid = False
                For Each row In dst.Rows
                    If Not (ErrorTextList = "") Then ErrorTextList = ErrorTextList & ", "
                    ErrorTextList = ErrorTextList & row.Item("LocationName")
                Next
                If Len(ErrorTextList) > 500 Then ErrorTextList = Left(ErrorTextList, 500) & " ..."
                ErrorMsg = ErrorMsg & ErrorTextList
            End If
            dst = Nothing
            ' End If

            ' End If

        End If
        Return Valid

    End Function

    Private Function IsvalidCurrencyId(ByVal OfferXmlDoc As XmlDocument, ByVal EngineID As Integer, ByRef ErrorMsg As String) As Boolean
        Dim ValidCurrency As Boolean = True
        Dim CurrencyID As String = ""
        Dim dt As DataTable

        TryParseElementValue(OfferXmlDoc, "//Offer/CurrencyID", CurrencyID)

        If CurrencyID.Trim <> "" Then
            MyCommon.QueryStr = "Select  Currencyid from Currencies where Currencyid=@CurrencyId"
            MyCommon.DBParameters.Add("@CurrencyId", SqlDbType.NVarChar).Value = CurrencyID.ConvertBlankIfNothing
            dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            ValidCurrency = (dt.Rows.Count > 0)
            If Not ValidCurrency Then
                ErrorMsg = "Offer CurrencyID '" & CurrencyID & "' is invalid"
            End If
        End If

        Return ValidCurrency

    End Function

End Class
