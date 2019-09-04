<%@ WebService Language="VB" Class="ShellFuelTargetedOffers" %>
Imports System
Imports System.Web
Imports System.Web.Services
Imports System.Xml.Serialization
Imports System.Xml.Schema
Imports System.Xml
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Thread
Imports System.Threading
Imports System.Collections.Generic
Imports System.Collections.Specialized.NameValueCollection
Imports Copient.CommonInc
Imports Copient.StoredValue
Imports Copient.CustomerInquiry
Imports Copient
Imports Copient.commonShared

<WebService(Namespace:="http://www.copienttech.com/ShellFuelTargetedOffers/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class ShellFuelTargetedOffers
  Inherits System.Web.Services.WebService

  Implements ISyncShellTargetedOffersBinding
  
    Private MyCommon As New Copient.CommonInc
    Private MyCryptLib As New Copient.CryptLib
  Private ShellWSLogFile As String = "ShellWSLog." & Date.Now.ToString("yyyyMMdd") & ".txt"

  'Time Tracking Logging
  Private Const sVersion As String = "7.3.1.138972"
  Private Const sAppName As String = "ShellFuelTargetedOffers.asmx"
  Private Const sLogFileName As String = "ShellFuelTargetedOffers"
  Private Const scDashes As String = "------------------------------------"
  Private Const bDebugLogOn As Boolean = True
  Private Const sDateFormat As String = "yyyy-MM-ddTHH:mm:ss"
  Private Enum DebugState
    BeginTime = -1
    CurrentTime = 0
    EndTime = 1
  End Enum
  Private DebugStartTimes As New ArrayList()
  'bTimeLogOn sets or clears logging for 
  Private bTimeLogOn As Boolean = IIf(System.Configuration.ConfigurationManager.AppSettings("ElapsedTimeLoggingForShellWS").ToUpper = "TRUE", true, false)
  Private Enum MessageType
    Info = 0
    Warning = 1
    AppError = 2
    SysError = 3
    Debug = 4
  End Enum
  Private sInputForLog As String = ""
  Private sLogLines As String = ""
  'Private sLogText As String = ""
  Private sCurrentMethod As String = ""
  Private sInstallationName As String = ""
  Private bTimeLogHasBeenCalled As Boolean = False
  Private bTimeLogLastCall As Boolean = False
  
  Public Structure ThirdPartyTransaction
    Dim ArrangementID As String
    Dim CustomerPk As Long
    Dim LoyaltyCardNumber As String
    Dim RedeemableQuantity As Integer
    Dim SVProgramID As Integer
    Dim CouponID As String
    Dim extLocId As String
    Dim DiscountAmount As Decimal
    Dim POSTimeStamp As Date
    Dim ProcessTransaction As Boolean
    Dim OfferRewardAmount As Decimal
    Dim AdjustAmount As Integer
  End Structure
  Public Structure TransactionFields
    Dim extLocCode As String
    Dim TransactionNumber As String
    Dim POSDateTime As DateTime
    Dim CustomerPk As Long
    Dim ExtCardID As String
    Dim CardType As Integer
    Dim HHID As String
    Dim TerminalNum As String
    Dim DiscountTotal As Decimal
    Dim OfferID As Long
    Dim SVProgramID As Integer
    Dim SVProgramQuantity As Integer
  End Structure
  Public Structure Items
    Dim ItemID As String
    Dim Qualifier As String
    Dim Type As String
  End Structure

  Public objItems As Items
  Public MaxNoOfDistribution As Integer
  <WebMethod()> _
  Public Function GetNetworkManagement(ByVal NWMRequest As NWMType) As NWMType Implements ISyncShellTargetedOffersBinding.GetNetworkManagement
    WriteDebug("GetNetworkManagement WebService Called", DebugState.CurrentTime)
    WriteDebug("GetNetworkManagement WebService Returning, Timer Stopped", DebugState.EndTime)
    Return NWMRequest

  End Function
  <WebMethod()> _
  Public Function GetTargetedOffers(ByVal GetTargetedOffers1 As CustomerType) As CustomerType Implements ISyncShellTargetedOffersBinding.GetTargetedOffers
    WriteDebug("GetTargetedOffers WebService Called", DebugState.CurrentTime)
    Dim MethodName As String = "GetTargetedOffers"
    Dim output As StringBuilder = New StringBuilder()
    Dim sw As New IO.StringWriter()    
    Dim CustomerXmlDoc As New XmlDocument()
        Dim ExtCardID As String
        Dim i As Integer
        Dim dtOffersByProduct As DataTable
        Dim logixproductid As Long = 0
        Dim ProdGroupIDs As String = ""
        Dim GallonLimit As Decimal = 0
        Dim CustSVBalance As Integer
        Dim DistributionLimit As Integer
        Dim OfferID As Long
        Dim QtyForIncentive As Integer
        Dim l_pgID As Long
        Dim CardExists, SiteExists As Boolean
        Dim HouseholdPK As Long = 0 ' for now assuming that the provided card will always be customer card (card type 0)
        Dim OffersAvailable As Boolean
        Dim oThirdPartyTrans As ThirdPartyTransaction
        Dim shellxml As String
        Dim pointsredeemed As Integer = 0

        Dim LocalCommon As New Copient.CommonInc
        LocalCommon.AppName = "ShellFuelTargetedOffers.asmx"

        Try
    
            'For Debug
            shellxml = GetXML(GetTargetedOffers1)
            If System.Configuration.ConfigurationManager.AppSettings("EnhancedLoggingForShellWS").ToUpper = "TRUE" Then
                Copient.Logger.Write_Log(ShellWSLogFile, "Actual XML for Message Type 1:" & vbCrLf & shellxml, True)
            End If

            'If IsValidCaller(MethodName) Then 'this check has been disabled per request from customer.  Reinstate as soon as possible
            If True Then

                If LocalCommon.LRTadoConn.State = ConnectionState.Closed Then LocalCommon.Open_LogixRT()
                If LocalCommon.LXSadoConn.State = ConnectionState.Closed Then LocalCommon.Open_LogixXS()


                'Dim x As XmlSerializer = (CreateOverrider("ShowTargetedOffers"))
                'Dim ns As New XmlSerializerNamespaces()
                'x.Serialize(sw, GetTargetedOffers1)
                'shellxml = sw.ToString

                objItems = GetItems(shellxml)
                sw.Close()


                'ConnInc.ConvertStringToXML(shellxml, CustomerXmlDoc)
                'Dim ShowTargtedOffers As CustomerType

                'Dim rootNode As XmlNode = CustomerXmlDoc.DocumentElement
                'For Each Elemnt As System.Xml.XmlElement In CustomerXmlDoc.DocumentElement.ChildNodes
                '  For i = 0 To Elemnt.ChildNodes.Count - 1
                '    If Elemnt.ChildNodes(i).Name = "LoyaltyAccountID" Then
                '      ExtCardID = Elemnt.ChildNodes(i).InnerText
                '    End If
                '    If Elemnt.ChildNodes(i).Name = "Item" Then
                '      UPC = Elemnt.ChildNodes(i).ChildNodes.ItemOf(0).InnerText
                '    End If
                '  Next

                'Next

                'Checks for the existence of the card

                ExtCardID = GetTargetedOffers1.LoyaltyAccount.LoyaltyAccountID
                oThirdPartyTrans.extLocId = GetTargetedOffers1.ShoppingBasket.TransactionLink.Site.SiteID
                Copient.Logger.Write_Log(ShellWSLogFile, "AuthRequest came with bTimeLogOn: " & bTimeLogOn & " LoyaltyCardId: " & MaskHelper.MaskCard(ExtCardID, commonShared.CardTypes.CUSTOMER) & " for CustomerPK " & oThirdPartyTrans.CustomerPk & " , SiteID: " & oThirdPartyTrans.extLocId & "", True)

                CardExists = DoesCardExist(ExtCardID, oThirdPartyTrans.CustomerPk)
                SiteExists = DoesSiteExist(oThirdPartyTrans.extLocId)

                oThirdPartyTrans.ArrangementID = System.Guid.NewGuid.ToString

                WriteDebug("Created ArrangementID: " & oThirdPartyTrans.ArrangementID, DebugState.CurrentTime)

                GetTargetedOffers1.SecurityAgreement = New SecurityArrangementType

                GetTargetedOffers1.SecurityAgreement.ArrangementID = oThirdPartyTrans.ArrangementID
                GetTargetedOffers1.SecurityAgreement.ExpirationDate = Nothing

                If SiteExists Then
                    If CardExists Then

                        'generate arrangement id

                        oThirdPartyTrans.LoyaltyCardNumber = ExtCardID
                        '
                        'Lets say, the SV Program for fuel is 27  -- 8896
                        '
                        WriteDebug("Database Call - Get System Option 158 request", DebugState.CurrentTime)
                        l_pgID = LocalCommon.Fetch_CPE_SystemOption(158)
                        WriteDebug("Database Call - Get System Option 158 response", DebugState.CurrentTime)

                        'Offers by SVProgram

                        'MyCommon.QueryStr = "SELECT DISTINCT I.IncentiveName as Name, I.IncentiveID as OfferID, 2 as EngineID, DISC.DiscountAmount  as 'OfferRewardAmount', DISC.WeightLimit, ISVP.QtyForIncentive, DSV.Quantity, ISVP.SVProgramID from CPE_IncentiveStoredValuePrograms ISVP with (NoLock)  " & _
                        '                  "INNER JOIN CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID=ISVP.RewardOptionID  " & _
                        '                  "INNER JOIN CPE_Incentives I with (NoLock) on I.IncentiveID = RO.IncentiveID  " & _
                        '                  "INNER JOIN StoredValuePrograms SVP with (NoLock) on ISVP.SVProgramID = SVP.SVProgramID  " & _
                        '                  "WHERE ISVP.SVProgramID=" & l_pgID & " and ISVP.Deleted = 0 and RO.Deleted=0 and I.Deleted=0 and I.IsTemplate=0 and SVP.Deleted=0  " & _
                        '                  "UNION  " & _
                        '                  "SELECT DISTINCT I.IncentiveName as Name, I.IncentiveID as OfferID, 2 as EngineID, DISC.DiscountAmount  as 'OfferRewardAmount', DISC.WeightLimit, DSV.Quantity from CPE_DeliverableStoredValue DSV with (NoLock) " & _
                        '                  "INNER JOIN CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID=DSV.RewardOptionID  " & _
                        '                  "INNER JOIN CPE_Incentives I with (NoLock) on I.IncentiveID = RO.IncentiveID  " & _
                        '                  "WHERE DSV.Deleted=0 and RO.Deleted=0 and I.Deleted=0 and I.IsTemplate=0 and SVProgramID=" & l_pgID & " " & _
                        '                  "UNION " & _
                        '                  "SELECT DISTINCT I.IncentiveName as Name, I.IncentiveID as OfferID, 2 as EngineID, DISC.DiscountAmount  as 'OfferRewardAmount', DISC.WeightLimit, DSV.Quantity from CPE_Deliverables D with (NoLock) " & _
                        '                  "INNER JOIN CPE_RewardOptions RO on RO.RewardOptionID = D.RewardOptionID and RO.Deleted=0 " & _
                        '                  "INNER JOIN CPE_Incentives I on I.IncentiveID = RO.IncentiveID and I.Deleted = 0 " & _
                        '                  "INNER JOIN CPE_Discounts DISC on DISC.DiscountID = D.OutputID and DISC.Deleted=0 " & _
                        '                  "WHERE D.DeliverableTypeID=2 and D.Deleted=0 and DISC.AmountTypeID=7 and DISC.SVProgramID = " & l_pgID & " "

                        LocalCommon.QueryStr = "SELECT DISTINCT I.IncentiveName as Name, I.IncentiveID as OfferID, 2 as EngineID, DISC.DiscountAmount  as 'OfferRewardAmount', DISC.WeightLimit, ISVP.QtyForIncentive, DSV.Quantity, SVP.SVProgramID, I.P3DistQtyLimit FROM CPE_IncentiveStoredValuePrograms ISVP with (NoLock) " & _
                                            "INNER JOIN CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID=ISVP.RewardOptionID  " & _
                                            "INNER JOIN CPE_Incentives I with (NoLock) on I.IncentiveID = RO.IncentiveID  " & _
                                            "INNER JOIN OfferTerminals OT with (NoLock) on OT.OfferID = I.IncentiveID " & _
                                            "INNER JOIN TerminalTypes TT with (NoLock) on TT.TerminalTypeID = OT.TerminalTypeID " & _
                                            "INNER JOIN StoredValuePrograms SVP with (NoLock) on ISVP.SVProgramID = SVP.SVProgramID  " & _
                                            "INNER JOIN CPE_DeliverableStoredValue DSV with (NoLock) on RO.RewardOptionID=DSV.RewardOptionID " & _
                                            "INNER JOIN CPE_Deliverables D with (NoLock) on RO.RewardOptionID=D.RewardOptionID " & _
                                            "INNER JOIN CPE_Discounts DISC on DISC.DiscountID = D.OutputID and DISC.Deleted=0 " & _
                                            "WHERE ISVP.SVProgramID=" & l_pgID & " and ISVP.Deleted = 0 and RO.Deleted=0 and I.Deleted=0 and I.IsTemplate=0 and SVP.Deleted=0 and DSV.Deleted=0 " & _
                                            "and DSV.SVProgramID=" & l_pgID & " and D.DeliverableTypeID=2 and D.Deleted=0 And I.EndDate >= GetDate() and TT.FuelProcessing = 1"

                        WriteDebug("Database Call - dtOffersByProduct request", DebugState.CurrentTime)
                        dtOffersByProduct = LocalCommon.LRT_Select
                        WriteDebug("Database Call - dtOffersByProduct response", DebugState.CurrentTime)
                        If dtOffersByProduct.Rows.Count > 0 Then
                            OffersAvailable = True
                            For i = 0 To dtOffersByProduct.Rows.Count - 1
                                OfferID = dtOffersByProduct.Rows(i).Item("OfferID")
                                oThirdPartyTrans.OfferRewardAmount = dtOffersByProduct.Rows(i).Item("OfferRewardAmount")
                                GallonLimit = dtOffersByProduct.Rows(i).Item("WeightLimit")
                                QtyForIncentive = dtOffersByProduct.Rows(i).Item("QtyForIncentive")
                                oThirdPartyTrans.SVProgramID = dtOffersByProduct.Rows(i).Item("SVProgramID")
                                oThirdPartyTrans.RedeemableQuantity = dtOffersByProduct.Rows(i).Item("Quantity")
                                DistributionLimit = dtOffersByProduct.Rows(i).Item("P3DistQtyLimit")
                            Next
                            Copient.Logger.Write_Log(ShellWSLogFile, "Offer " & OfferID & " is available")
                        Else
                            'No offers available
                            GetTargetedOffers1.ResponseCodeSpecified = True
                            GetTargetedOffers1.ResponseCode = ResponseCodeType.KBValidAccountWithoutDiscount
                            GetTargetedOffers1.ResponseDescription = "No offers available"
                            Copient.Logger.Write_Log(ShellWSLogFile, "There are no offers available for Customer card " & MaskHelper.MaskCard(ExtCardID, commonShared.CardTypes.CUSTOMER) & " for CustomerPK " & oThirdPartyTrans.CustomerPk, True)
                            WriteDebug("GetTargetedOffers WebService Returning, timer stopped", DebugState.EndTime)
                            Return GetTargetedOffers1
                            Exit Function
                        End If

                        If OffersAvailable Then
                            If Not IsSufficientBalance(oThirdPartyTrans.CustomerPk, oThirdPartyTrans.SVProgramID, QtyForIncentive, CustSVBalance) Then 'offer available. check for the sufficient balance in customer sv bucket expecting only one offer for fuel SVProgram      
                                GetTargetedOffers1.ResponseCodeSpecified = True
                                GetTargetedOffers1.ResponseCode = ResponseCodeType.KBValidAccountWithoutDiscount
                                GetTargetedOffers1.ResponseDescription = "InsufficientBalance"
                                Copient.Logger.Write_Log(ShellWSLogFile, "Customer card " & MaskHelper.MaskCard(ExtCardID, commonShared.CardTypes.CUSTOMER) & " for CustomerPK " & oThirdPartyTrans.CustomerPk & " does not have sufficient stored value prog balance", True)
                                WriteDebug("GetTargetedOffers WebService Returning, timer stopped", DebugState.EndTime)
                                Copient.Logger.Write_Log(ShellWSLogFile, "Actual XML for Message Type 2:" & vbCrLf & shellxml, True)
                                Return GetTargetedOffers1
                                Exit Function
                            Else

                                oThirdPartyTrans.DiscountAmount = ComputeDiscountAmount(oThirdPartyTrans.OfferRewardAmount, QtyForIncentive, CustSVBalance, DistributionLimit)

                                oThirdPartyTrans.AdjustAmount = ComputeAdjustmentAmount(oThirdPartyTrans.DiscountAmount, oThirdPartyTrans.OfferRewardAmount, oThirdPartyTrans.RedeemableQuantity)
                                pointsredeemed = oThirdPartyTrans.AdjustAmount * -1
                                oThirdPartyTrans.CouponID = oThirdPartyTrans.SVProgramID.ToString().PadLeft(3, "0") & pointsredeemed.ToString().PadLeft(5, "0") & OfferID.ToString().PadLeft(6, "0")
                                oThirdPartyTrans.POSTimeStamp = GetTargetedOffers1.POSTimestamp



                                'Populate the table which will be used for next method call for this transaction (ArrangementID) and for creation of reconciliation file
                                PopulateThirdPartyTransactions(oThirdPartyTrans)

                                'response back with required items
                                GetTargetedOffers1.ResponseCodeSpecified = True
                                GetTargetedOffers1.ResponseCode = ResponseCodeType.KAValidAccountWithDiscount
                                GetTargetedOffers1.ResponseDescription = "Valid account with discount"

                                GetTargetedOffers1.Promotion = New PromotionIDType
                                GetTargetedOffers1.Promotion.Item = New ItemType
                                GetTargetedOffers1.Promotion.Item.Quantity = New QuantityType
                                GetTargetedOffers1.Promotion.DiscountAmount = New MonetaryAmountType
                                GetTargetedOffers1.Promotion.DiscountAmount.Amount = New MonetaryAmountTypeAmount
                                GetTargetedOffers1.Promotion.QuantityLimit = New QuantityType
                                'GetTargetedOffers1.ShoppingBasket = New BasketType
                                GetTargetedOffers1.LoyaltyAccount = New LoyaltyAccountType

                                GetTargetedOffers1.Promotion.CouponID = oThirdPartyTrans.CouponID
                                GetTargetedOffers1.Promotion.DiscountAmount.Amount.Value = Decimal.Round(oThirdPartyTrans.DiscountAmount, 2)
                                GetTargetedOffers1.Promotion.Item.ItemID = objItems.ItemID
                                GetTargetedOffers1.Promotion.Item.Qualifier = objItems.Qualifier
                                GetTargetedOffers1.Promotion.Item.Type = IIf(objItems.Type.ToUpper = "UPC", 0, 0)
                                GetTargetedOffers1.Promotion.Item.Quantity.Units = 0
                                GetTargetedOffers1.Promotion.Item.Quantity.UnitOfMeasureCode = UnitOfMeasureCodeCommonData.GLL
                                GetTargetedOffers1.Promotion.QuantityLimit.Units = Decimal.Round(GallonLimit, 2)
                                GetTargetedOffers1.LoyaltyAccount.LoyaltyAccountID = ReFormatCard(ExtCardID)
                            End If
                        End If

                    Else
                        'send a failure response
                        GetTargetedOffers1.ResponseCodeSpecified = True
                        GetTargetedOffers1.ResponseCode = ResponseCodeType.KCUnknownAccount
                        GetTargetedOffers1.ResponseDescription = "UnknownAccount"
                        Copient.Logger.Write_Log(ShellWSLogFile, "The Customer card " & MaskHelper.MaskCard(ExtCardID, commonShared.CardTypes.CUSTOMER) & " for CustomerPK " & oThirdPartyTrans.CustomerPk & " not found", True)
                    End If
                Else
                    'send a failure response
                    GetTargetedOffers1.ResponseCodeSpecified = True
                    GetTargetedOffers1.ResponseCode = ResponseCodeType.KDUnknownSite
                    GetTargetedOffers1.ResponseDescription = "UnknownSite"
                    Copient.Logger.Write_Log(ShellWSLogFile, "The Site " & oThirdPartyTrans.extLocId & " not found", True)
                End If
            Else
                GetTargetedOffers1.ResponseCodeSpecified = True
                GetTargetedOffers1.ResponseCode = ResponseCodeType.KGInvalidMessageFormat
                GetTargetedOffers1.ResponseDescription = "The Caller can not be validated"
                Copient.Logger.Write_Log(ShellWSLogFile, "The Caller can not be validated", True)
            End If

        Catch ex As Exception
            GetTargetedOffers1.ResponseCodeSpecified = True
            GetTargetedOffers1.ResponseCode = ResponseCodeType.KGInvalidMessageFormat
            GetTargetedOffers1.ResponseDescription = "Exception occured during processing"
            Copient.Logger.Write_Log(ShellWSLogFile, ex.Message, True)
        Finally
            If LocalCommon.LRTadoConn.State = ConnectionState.Open Then LocalCommon.Close_LogixRT()
            If LocalCommon.LXSadoConn.State = ConnectionState.Open Then LocalCommon.Close_LogixXS()

        End Try

        If System.Configuration.ConfigurationManager.AppSettings("EnhancedLoggingForShellWS").ToUpper = "TRUE" Then
            shellxml = GetXML(GetTargetedOffers1) 'For Debug Purpose
            Copient.Logger.Write_Log(ShellWSLogFile, "Actual XML for Message Type 2:" & vbCrLf & shellxml, True)
        Else
            Copient.Logger.Write_Log(ShellWSLogFile, "AuthResponse to LoyaltyCardId: " & MaskHelper.MaskCard(ExtCardID, commonShared.CardTypes.CUSTOMER) & " for CustomerPK " & oThirdPartyTrans.CustomerPk & " , SiteID: " & oThirdPartyTrans.extLocId & " with couponid: " & oThirdPartyTrans.CouponID & " and ArrangementID: " & oThirdPartyTrans.ArrangementID & "", True)
        End If
        WriteDebug("GetTargetedOffers WebService Returning, timer stopped", DebugState.EndTime)
        Return GetTargetedOffers1
    End Function

    <WebMethod()> _
    Public Function ProcessTargetedOffers(ByVal ProcessTargetedOffers1 As CustomerType) As CustomerType Implements ISyncShellTargetedOffersBinding.ProcessTargetedOffers
        WriteDebug("ProcessTargetedOffers WebService Called", DebugState.CurrentTime)
        Dim MethodName As String = "ProcessTargetedOffers"
        Dim oTransFields As TransactionFields
        Dim sw As New IO.StringWriter()
        Dim dt As DataTable
        Dim MySV As New StoredValue
        Dim RetMsg As String = ""
        Dim AdminUserID As Long = 1
        Dim SVProgramID As Integer
        Dim CustomerPK As Long
        Dim AdjustAmount As Integer
        Dim CouponID As String
        Dim Shellxml As String
        Dim Comments As String = "Updating customer stored value points for shell transaction at "
        Dim ArrangementIDFound As Boolean = False
        Dim LoyaltyCardNumber As String = ""

        'ProcessTargetedOffers1.ShoppingBasket = New BasketType
        'ProcessTargetedOffers1.ShoppingBasket.TransactionLink = New OldTransactionLinkCommonData
        'ProcessTargetedOffers1.Promotion = New PromotionIDType
        'ProcessTargetedOffers1.Promotion.DiscountAmount = New MonetaryAmountType
        'ProcessTargetedOffers1.Promotion.DiscountAmount.Amount = New MonetaryAmountTypeAmount
        Dim LocalCommon As New Copient.CommonInc
        LocalCommon.AppName = "ShellFuelTargetedOffers.asmx"

        Try
            'For Debug Purpose

            If System.Configuration.ConfigurationManager.AppSettings("EnhancedLoggingForShellWS").ToUpper = "TRUE" Then
                Shellxml = GetXML(ProcessTargetedOffers1)
                Copient.Logger.Write_Log(ShellWSLogFile, "Actual XML for Message Type 3:" & vbCrLf & Shellxml, True)
            End If

            'If IsValidCaller(MethodName) Then 'this check has been disabled per request from customer.  Reinstate as soon as possible
            If True Then

                If LocalCommon.LRTadoConn.State = ConnectionState.Closed Then LocalCommon.Open_LogixRT()
                If LocalCommon.LXSadoConn.State = ConnectionState.Closed Then LocalCommon.Open_LogixXS()
                If LocalCommon.LWHadoConn.State = ConnectionState.Closed Then LocalCommon.Open_LogixWH()
                'fetch the user Shell
                AdminUserID = GetAdminUserID()

                Dim x As XmlSerializer = (CreateOverrider("AcknowledgeTargetedOffers"))
                x.Serialize(sw, ProcessTargetedOffers1)
                Shellxml = sw.ToString
                sw.Close()

                'check if total discount amount is non zero. it indicates that the discount has been redeemed at shell
                If ProcessTargetedOffers1.TotalDiscountAmount.Amount.Value <> 0 Then
                    'check for the existence arrangementid in ThirdPartyTransactions table
                    Copient.Logger.Write_Log(ShellWSLogFile, "Redemption request came for ArrangementID:" & ProcessTargetedOffers1.SecurityAgreement.ArrangementID, True)
                    LocalCommon.QueryStr = "select SVProgramID, CustomerPK, LoyaltyCardNumber, SiteID, CouponID, MaxDistributions, AdjustAmount  from ThirdPartyTransactions with (NoLock) where ArrangementID = '" & ProcessTargetedOffers1.SecurityAgreement.ArrangementID & "'"
                    WriteDebug("Database Call - ThirdPartyTransactions for ArrangementID " & ProcessTargetedOffers1.SecurityAgreement.ArrangementID & " - request", DebugState.CurrentTime)
                    dt = LocalCommon.LRT_Select
                    WriteDebug("Database Call - ThirdPartyTransactions for ArrangementID " & ProcessTargetedOffers1.SecurityAgreement.ArrangementID & " - response", DebugState.CurrentTime)

                    If dt.Rows.Count > 0 Then ' ArrangementID found, make adjustment
                        ArrangementIDFound = True
                        Copient.Logger.Write_Log(ShellWSLogFile, "ArrangementID found, Processing for adjustment", True)
                        CouponID = dt.Rows(0).Item("CouponID")
                        SVProgramID = dt.Rows(0).Item("SVProgramID")
                        CustomerPK = dt.Rows(0).Item("CustomerPK")
                        LoyaltyCardNumber = MyCryptLib.SQL_StringDecrypt(dt.Rows(0).Item("LoyaltyCardNumber").ToString())
                        'SVQuantity = dt.Rows(0).Item("SVProgramQuantity")
                        oTransFields.extLocCode = dt.Rows(0).Item("SiteID")
                        MaxNoOfDistribution = dt.Rows(0).Item("MaxDistributions")
                        'RewardAmount = dt.Rows(0).Item("OfferRewardAmount")
                        'MaxDiscountAmount = dt.Rows(0).Item("DiscountAmount")
                        AdjustAmount = dt.Rows(0).Item("AdjustAmount")
                        RetMsg = MySV.AdjustStoredValue(AdminUserID, SVProgramID, CustomerPK, AdjustAmount.ToString(), Comments)
                        'RetMsg = MySV.AdjustStoredValue(AdminUserID, SVProgramID, CustomerPK, AdjustAmount.ToString())
                        If RetMsg = "" Then
                            Copient.Logger.Write_Log(ShellWSLogFile, "Adjustment is successful  with adjust amount:  " & AdjustAmount & " for ArrangementID: " & ProcessTargetedOffers1.SecurityAgreement.ArrangementID & "", True)
                            ProcessTargetedOffers1.ResponseCodeSpecified = True
                            ProcessTargetedOffers1.ResponseCode = ResponseCodeType.KAValidAccountWithDiscount
                            ProcessTargetedOffers1.ResponseDescription = "valid account with discount"

                            'Update the flag in ThirdPartyTransactions table, which indicates successfull adjustment for a transaction

                            LocalCommon.QueryStr = "Update ThirdPartyTransactions with (RowLock) set ProcessTransaction = 1, DiscountAmount=" & ProcessTargetedOffers1.TotalDiscountAmount.Amount.Value & ", SVProgramQuantity = " & AdjustAmount & "  where ArrangementID = '" & ProcessTargetedOffers1.SecurityAgreement.ArrangementID & "' "
                            WriteDebug("Database Call - Update ThirdPartyTransactions for ArrangementID " & ProcessTargetedOffers1.SecurityAgreement.ArrangementID & " - request", DebugState.CurrentTime)
                            LocalCommon.LRT_Execute()
                            WriteDebug("Database Call - Update ThirdPartyTransactions for ArrangementID " & ProcessTargetedOffers1.SecurityAgreement.ArrangementID & " - request", DebugState.CurrentTime)

                            'Update the customer ytd savings to the given discount amount
                            UpdateSTD(CustomerPK, ProcessTargetedOffers1.TotalDiscountAmount.Amount.Value)

                            'Populate the transactions tables
                            oTransFields.CustomerPk = CustomerPK
                            oTransFields.ExtCardID = LoyaltyCardNumber
                            oTransFields.TransactionNumber = ProcessTargetedOffers1.ShoppingBasket.TransactionLink.SequenceNumber
                            oTransFields.POSDateTime = ProcessTargetedOffers1.POSTimestamp
                            oTransFields.DiscountTotal = ProcessTargetedOffers1.TotalDiscountAmount.Amount.Value
                            oTransFields.TerminalNum = IIf(ProcessTargetedOffers1.ShoppingBasket.TransactionLink.WorkstationID IsNot Nothing, ProcessTargetedOffers1.ShoppingBasket.TransactionLink.WorkstationID, "")
                            oTransFields.OfferID = CouponID.Substring(SVProgramID.ToString.Length)
                            oTransFields.SVProgramID = SVProgramID
                            oTransFields.SVProgramQuantity = AdjustAmount
                            PopulatetransactionTables(oTransFields)

                        Else
                            Copient.Logger.Write_Log(ShellWSLogFile, RetMsg, True)
                        End If
                    Else
                        ProcessTargetedOffers1.ResponseCodeSpecified = True
                        ProcessTargetedOffers1.ResponseCode = ResponseCodeType.KEInvalidArrangementID
                        ProcessTargetedOffers1.ResponseDescription = "Invalid ArrangementID"
                        Copient.Logger.Write_Log(ShellWSLogFile, "Invalid ArrangementID: " & ProcessTargetedOffers1.SecurityAgreement.ArrangementID & " for loyalty cardid: " & MaskHelper.MaskCard(ProcessTargetedOffers1.LoyaltyAccount.LoyaltyAccountID, commonShared.CardTypes.CUSTOMER) & " of type ID " & commonShared.CardTypes.CUSTOMER & ", siteid: " & ProcessTargetedOffers1.ShoppingBasket.TransactionLink.Site.SiteID & "", True)
                    End If
                End If
            End If

        Catch ex As Exception
            Copient.Logger.Write_Log(ShellWSLogFile, ex.Message, True)
        Finally
            If LocalCommon.LRTadoConn.State = ConnectionState.Open Then LocalCommon.Close_LogixRT()
            If LocalCommon.LXSadoConn.State = ConnectionState.Open Then LocalCommon.Close_LogixXS()
            If LocalCommon.LWHadoConn.State = ConnectionState.Open Then LocalCommon.Close_LogixWH()
        End Try

        If System.Configuration.ConfigurationManager.AppSettings("EnhancedLoggingForShellWS").ToUpper = "TRUE" Then
            Shellxml = GetXML(ProcessTargetedOffers1)
            Copient.Logger.Write_Log(ShellWSLogFile, "Actual XML for Message Type 4:" & vbCrLf & Shellxml, True)
        End If

        WriteDebug("ProcessTargetedOffers WebService Returning, timer stopped", DebugState.EndTime)

        Return ProcessTargetedOffers1

    End Function
  Public Sub PopulatetransactionTables(ByVal oTransFields As TransactionFields)
    WriteDebug("IN PopulatetransactionTables", DebugState.CurrentTime)

    Dim LocalCommon As New Copient.CommonInc
    LocalCommon.AppName = "ShellFuelTargetedOffers.asmx"
    Try
      If LocalCommon.LWHadoConn.State = ConnectionState.Closed Then LocalCommon.Open_LogixWH()
      Dim PosTimestamp As String
      Dim CustomerDetails(1) As String
      'bifurcate the PosDateTime into two parts (Date MMDD (4),Time hhmmss (6))
      PosTimestamp = oTransFields.POSDateTime.Month.ToString() & oTransFields.POSDateTime.Day.ToString() & oTransFields.POSDateTime.Hour.ToString() & _
        oTransFields.POSDateTime.Minute.ToString() & oTransFields.POSDateTime.Second.ToString()

      CustomerDetails = GetCustomerDetails(oTransFields.CustomerPk)


      LocalCommon.QueryStr = "pt_ThirdPartyTransTable_Insert"
      WriteDebug("Database Call - execute pt_ThirdPartyTransTable_Insert START", DebugState.CurrentTime)
      LocalCommon.Open_LWHsp()
            LocalCommon.LWHsp.Parameters.Add("@ExtCardID", SqlDbType.NVarChar, 26).Value = oTransFields.ExtCardID
      LocalCommon.LWHsp.Parameters.Add("@CardType", SqlDbType.NVarChar, 6).Value = CustomerDetails(0)
      LocalCommon.LWHsp.Parameters.Add("@HHID", SqlDbType.NVarChar, 26).Value = IIf(CustomerDetails(1) IsNot Nothing, CustomerDetails(1), 0)
      LocalCommon.LWHsp.Parameters.Add("@TransactionNumber", SqlDbType.NVarChar, 6).Value = oTransFields.TransactionNumber
      LocalCommon.LWHsp.Parameters.Add("@SiteID", SqlDbType.NVarChar, 20).Value = oTransFields.extLocCode
      LocalCommon.LWHsp.Parameters.Add("@DiscountAmount", SqlDbType.Decimal).Value = oTransFields.DiscountTotal
      LocalCommon.LWHsp.Parameters.Add("@postimestamp", SqlDbType.NVarChar, 12).Value = PosTimestamp
      LocalCommon.LWHsp.Parameters.Add("@POSDateTime", SqlDbType.DateTime).Value = oTransFields.POSDateTime
      LocalCommon.LWHsp.Parameters.Add("@TerminalNum", SqlDbType.NVarChar, 12).Value = oTransFields.TerminalNum
      LocalCommon.LWHsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = oTransFields.OfferID
      LocalCommon.LWHsp.Parameters.Add("@SVProgramID", SqlDbType.Int).Value = oTransFields.SVProgramID
      LocalCommon.LWHsp.Parameters.Add("@SVProgramQuantity", SqlDbType.Int).Value = oTransFields.SVProgramQuantity
      LocalCommon.LWHsp.ExecuteNonQuery()
      LocalCommon.Close_LWHsp()
      WriteDebug("Database Call - execute pt_ThirdPartyTransTable_Insert END", DebugState.CurrentTime)

    Catch ex As Exception
      Copient.Logger.Write_Log(ShellWSLogFile, ex.Message, True)
    Finally
        If LocalCommon.LWHadoConn.State = ConnectionState.Open Then LocalCommon.Close_LogixWH()
    End Try

  End Sub
  Public Function CreateOverrider(ByVal ElementName As String) As XmlSerializer
    WriteDebug("IN CreateOverrider", DebugState.CurrentTime)

    ' Create an XmlAttributes to override the default root element. 
    Dim myXmlAttributes As New XmlAttributes()

    ' Create an XmlRootAttribute and set its element name and namespace. 
    Dim myXmlRootAttribute As New XmlRootAttribute()
    myXmlRootAttribute.ElementName = ElementName
    myXmlRootAttribute.Namespace = "http://schema.aholdusa.com/loyalty/fuel/LoyaltyManager"

    ' Set the XmlRoot property to the XmlRoot object.
    myXmlAttributes.XmlRoot = myXmlRootAttribute
    Dim myXmlAttributeOverrides As New XmlAttributeOverrides()

    ' Add the XmlAttributes object to the XmlAttributeOverrides object.
    myXmlAttributeOverrides.Add(GetType(CustomerType), myXmlAttributes)

    ' Create the Serializer, and return it. 
    Dim myXmlSerializer As New XmlSerializer(GetType(CustomerType), myXmlAttributeOverrides)
    Return myXmlSerializer
  End Function
  Public Function GetXML(ByVal otype As CustomerType) As String
    WriteDebug("IN GetXML", DebugState.CurrentTime)

    Dim sw As New IO.StringWriter()
    Dim ShellXML As String = ""
    Try
      Dim y As New XmlSerializer(GetType(CustomerType)) ' For Debug purpose
      y.Serialize(sw, otype) 'For Debug Purpose
      ShellXML = sw.ToString
      sw.Close()
    Catch ex As Exception

    End Try
    Return ShellXML
  End Function
  
  Public Function DoesCardExist(ByRef ExtCardID As String, ByRef CustomerPK As Long) As Boolean
    WriteDebug("IN DoesCardExist1", DebugState.CurrentTime)

    Dim LocalCommon As New Copient.CommonInc
    LocalCommon.AppName = "ShellFuelTargetedOffers.asmx"

    Dim Exists As Boolean = False
    Dim dt As DataTable
    Dim IsCardAltID As Boolean = False
    Dim IDLength As Integer = 0
    'If card starts with 1 AND if card length is 10 digits, append leading 22. 
    'If card starts with 22 AND if card length is 13 digits parse out the first 12 digits  
    'If the card starts with 6007 AND the digits 7-8 is 22, parse out the 7th through 18th digit.

    'lookup for customerPK
    Try
      If LocalCommon.LRTadoConn.State = ConnectionState.Closed Then LocalCommon.Open_LogixRT()
      If LocalCommon.LXSadoConn.State = ConnectionState.Closed Then LocalCommon.Open_LogixXS()

	  ExtCardID = ExtCardID.Trim()

      If ExtCardID.StartsWith("22") AndAlso ExtCardID.Length = 13 Then
        ExtCardID = ExtCardID.Substring(0, 12)
      ElseIf ExtCardID.StartsWith("600700") AndAlso ExtCardID.Substring(6, 2) = "22" Then
        ExtCardID = ExtCardID.Substring(6, 12)
      ElseIf ExtCardID.StartsWith("44") AndAlso ExtCardID.Length = 12 Then
        ExtCardID = ExtCardID.Substring(0, 11)
      ElseIf ExtCardID.StartsWith("600700") AndAlso ExtCardID.Substring(6, 2) = "44" Then
        ExtCardID = ExtCardID.Substring(6, 11)
      ElseIf ExtCardID.Length = 10 Then
        ExtCardID = ExtCardID & "0000"
        IsCardAltID = True
      End If
      
	  If Not IsCardAltID Then
	    Integer.TryParse(LocalCommon.Fetch_SystemOption(53), IDLength)	  
        If (IDLength > 0) Then
          ExtCardID = LocalCommon.Leading_Zero_Fill(ExtCardID, IDLength)      
        End If
	  End If
      

            Copient.Logger.Write_Log(ShellWSLogFile, "The formatted Cardid is : " & MaskHelper.MaskCard(ExtCardID, commonShared.CardTypes.CUSTOMER) & " for CustomerPK " & CustomerPK, True)

      If Not IsCardAltID Then
        Exists = DoesCardExist(ExtCardID, 0, CustomerPK)
      Else
        Exists = DoesCardExist(ExtCardID, 3, CustomerPK)        
      End If


      If IsCardAltID Then
        'get the corresponding loyalty card
        LocalCommon.QueryStr = "select CardPK, ExtCardID from CardIDs with (NoLock) where CustomerPK = " & CustomerPK & " and CardTypeID = 0"
        WriteDebug("Database Call - does this card exist?  request", DebugState.CurrentTime)
        dt = LocalCommon.LXS_Select
        WriteDebug("Database Call - does this card exist?  response", DebugState.CurrentTime)
        If dt.Rows.Count > 0 Then
          ExtCardID = MyCryptLib.SQL_StringDecrypt(dt.Rows(0).Item("ExtCardID").ToString())
        End If
      End If
      
            If Exists Then Copient.Logger.Write_Log(ShellWSLogFile, "Card " & MaskHelper.MaskCard(ExtCardID, commonShared.CardTypes.CUSTOMER) & " for CustomerPK " & CustomerPK & " Found", True)
    Catch ex As Exception
      Copient.Logger.Write_Log(ShellWSLogFile, ex.Message, True)
    Finally
      If LocalCommon.LRTadoConn.State = ConnectionState.Open Then LocalCommon.Close_LogixRT()
      If LocalCommon.LXSadoConn.State = ConnectionState.Open Then LocalCommon.Close_LogixXS()
    End Try

    Return Exists
  End Function
  Public Function DoesSiteExist(ByVal ExtLocId As String) As Boolean
    WriteDebug("IN DoesSiteExist", DebugState.CurrentTime)

    Dim LocalCommon As New Copient.CommonInc
    LocalCommon.AppName = "ShellFuelTargetedOffers.asmx"

    Dim exist As Boolean = False
    Dim dst As DataTable

  Try
    If LocalCommon.LRTadoConn.State = ConnectionState.Closed Then LocalCommon.Open_LogixRT()

    WriteDebug("Database Call - does this site exist?  request", DebugState.CurrentTime)
    LocalCommon.QueryStr = "select LocationID from Locations with (NoLock) where ExtLocationCode='" & ExtLocId & "' and Deleted=0;"
    WriteDebug("Database Call - does this site exist?  response", DebugState.CurrentTime)
    dst = LocalCommon.LRT_Select
    If dst.Rows.Count > 0 Then
      exist = True      
      Copient.Logger.Write_Log(ShellWSLogFile, "Site " & ExtLocId & " Found", True)
    End If
  Catch ex As Exception
    Copient.Logger.Write_Log(ShellWSLogFile, ex.Message, True)
  Finally
    If LocalCommon.LRTadoConn.State = ConnectionState.Open Then LocalCommon.Close_LogixRT()
  End Try

    Return exist
  End Function

  Public Function DoesCardExist(ByVal ExtCardID As String, ByVal CardTypeID As Integer, ByRef CustomerPK As Long) As Boolean
    WriteDebug("IN DoesCardExist2", DebugState.CurrentTime)

    Dim dt As DataTable
    Dim Exists As Boolean = False

    Dim LocalCommon As New Copient.CommonInc
    LocalCommon.AppName = "ShellFuelTargetedOffers.asmx"
  Try
    If LocalCommon.LXSadoConn.State = ConnectionState.Closed Then LocalCommon.Open_LogixXS()

    'Find the Customer PK
    LocalCommon.QueryStr = "select CustomerPK from CardIDs with (NoLock) " & _
                        "where ExtCardID='" & MyCryptLib.SQL_StringEncrypt(ExtCardID) & "' and CardTypeID = " & CardTypeID & ""

    WriteDebug("Database Call - does this card exist?  request", DebugState.CurrentTime)
    dt = LocalCommon.LXS_Select
    WriteDebug("Database Call - does this card exist?  response", DebugState.CurrentTime)
    If dt.Rows.Count > 0 Then
      Exists = True
      CustomerPK = dt.Rows(0).Item("CustomerPK")
    End If
  Catch ex As Exception
    Copient.Logger.Write_Log(ShellWSLogFile, ex.Message, True)
  Finally
    If LocalCommon.LXSadoConn.State = ConnectionState.Open Then LocalCommon.Close_LogixXS()
  End Try
    Return Exists

  End Function
  Public Function IsSufficientBalance(ByVal customerpk As Long, ByVal SVProgramID As Integer, ByVal SVProgramQuantity As Integer, ByRef CustomerSVBalance As Integer) As Boolean
    WriteDebug("IN IsSufficientBalance", DebugState.CurrentTime)

    Dim SufficientBalance As Boolean = False

    'Dim MyLookup As New Copient.CustomerLookup
    Dim LookupRetCode As Copient.CustomerAbstract.RETURN_CODE
    Dim Balances(-1) As Copient.CustomerAbstract.StoredValueBalance

    Try

      Balances = GetCustomerSVBalances(customerpk, False, LookupRetCode)

      If LookupRetCode = Copient.CustomerLookup.RETURN_CODE.OK Then
        For i As Integer = 0 To Balances.GetUpperBound(0)
          If Balances(i).SVProgramID = SVProgramID Then
            If Balances(i).Balance >= SVProgramQuantity Then
              SufficientBalance = True
              CustomerSVBalance = Balances(i).Balance
            End If
          End If
        Next
      End If
      ' load all stored value programs for lookup
    Catch ex As Exception
      Copient.Logger.Write_Log(ShellWSLogFile, ex.Message, True)
    End Try

    Return SufficientBalance
  End Function
  Private Sub PopulateThirdPartyTransactions(ByVal ThirdPartyTrans As ThirdPartyTransaction)
    WriteDebug("IN PopulateThirdPartyTransactions", DebugState.CurrentTime)

    Dim LocalCommon As New Copient.CommonInc
    LocalCommon.AppName = "ShellFuelTargetedOffers.asmx"

      If LocalCommon.LRTadoConn.State = ConnectionState.Closed Then LocalCommon.Open_LogixRT()

    Try
      WriteDebug("Database Call - execute pt_ThirdPartyTransactions_Insert START", DebugState.CurrentTime)
      LocalCommon.QueryStr = "pt_ThirdPartyTransactions_Insert"
      LocalCommon.Open_LRTsp()
      LocalCommon.LRTsp.Parameters.Add("@ArrangementID", SqlDbType.NVarChar, 36).Value = ThirdPartyTrans.ArrangementID
      LocalCommon.LRTsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = ThirdPartyTrans.CustomerPk
      LocalCommon.LRTsp.Parameters.Add("@LoyaltyCardNumber", SqlDbType.NVarChar, 400).Value = MyCryptLib.SQL_StringEncrypt(ThirdPartyTrans.LoyaltyCardNumber)
      LocalCommon.LRTsp.Parameters.Add("@SVProgramQuantity", SqlDbType.Int).Value = ThirdPartyTrans.RedeemableQuantity
      LocalCommon.LRTsp.Parameters.Add("@SVProgramID", SqlDbType.Int).Value = ThirdPartyTrans.SVProgramID
      LocalCommon.LRTsp.Parameters.Add("@CouponID", SqlDbType.NVarChar, 20).Value = ThirdPartyTrans.CouponID
      LocalCommon.LRTsp.Parameters.Add("@SiteID", SqlDbType.NVarChar, 20).Value = ThirdPartyTrans.extLocId
      LocalCommon.LRTsp.Parameters.Add("@OfferRewardAmount", SqlDbType.Decimal).Value = ThirdPartyTrans.OfferRewardAmount
      LocalCommon.LRTsp.Parameters.Add("@DiscountAmount", SqlDbType.Decimal).Value = ThirdPartyTrans.DiscountAmount
      LocalCommon.LRTsp.Parameters.Add("@POSTimeStamp", SqlDbType.DateTime).Value = ThirdPartyTrans.POSTimeStamp
      LocalCommon.LRTsp.Parameters.Add("@MaxDistributions", SqlDbType.Int).Value = MaxNoOfDistribution
      LocalCommon.LRTsp.Parameters.Add("@AdjustAmount", SqlDbType.Int).Value = ThirdPartyTrans.AdjustAmount
      LocalCommon.LRTsp.ExecuteNonQuery()
      WriteDebug("Database Call - execute pt_ThirdPartyTransactions_Insert END", DebugState.CurrentTime)
      Copient.Logger.Write_Log(ShellWSLogFile, "ThirdPartyTransaction populated with ArrangementID:" & ThirdPartyTrans.ArrangementID, True)
      'LocalCommon.LRTsp.ExecuteNonQuery()
      LocalCommon.Close_LRTsp()
    Catch ex As Exception
      Copient.Logger.Write_Log(ShellWSLogFile, ex.Message, True)
    Finally
      If LocalCommon.LRTadoConn.State = ConnectionState.Open Then LocalCommon.Close_LogixRT()
    End Try

  End Sub
  Private Function IsValidCaller(ByVal MethodName As String) As Boolean
    WriteDebug("IN IsValidCaller", DebugState.CurrentTime)

    Dim bIsValidCaller As Boolean = False
    Dim MsgBuf As New StringBuilder()
    Dim ClientIPAddress As String
    Dim SName As String

    Try
      ClientIPAddress = HttpContext.Current.Request.UserHostName


      ClientIPAddress = HttpContext.Current.Request.ServerVariables("SERVER_NAME")

      Dim computer_name As String
      Dim CallerName As String

      computer_name = System.Net.Dns.Resolve(HttpContext.Current.Request.ServerVariables("remote_addr")).HostName
      CallerName = computer_name

      Dim servername() As String = System.Configuration.ConfigurationManager.AppSettings("ServerName").Split(",")

      For Each SName In servername
        SName = SName.Trim
        If String.Compare(UCase(SName), UCase(CallerName)) = 0 Then
          bIsValidCaller = True         
          Exit For        
        End If
      Next
      If Not bIsValidCaller Then
        MsgBuf.Append("Could not validate call to:")
        MsgBuf.Append(MethodName)
        MsgBuf.Append("CallerName: " & CallerName)        
        Copient.Logger.Write_Log(ShellWSLogFile, MsgBuf.ToString, True)
      End If
    Catch ex As Exception
      Copient.Logger.Write_Log(ShellWSLogFile, ex.Message, True)
    End Try

    Return bIsValidCaller
  End Function
  Public Function GetCustomerSVBalances(ByVal CustomerPK As Long, ByVal IncludeZeroBalances As Boolean, _
                                      ByRef ReturnCode As Copient.CustomerAbstract.RETURN_CODE) As Copient.CustomerLookup.StoredValueBalance()
    Dim Balances(-1) As Copient.CustomerLookup.StoredValueBalance
    Dim SVBal As Copient.CustomerLookup.StoredValueBalance
    Dim dt, dt2 As DataTable
    Dim i, RowCt As Integer
    Dim HHPK As Long
    WriteDebug("IN GetCustomerSVBalances", DebugState.CurrentTime)

    Dim LocalCommon As New Copient.CommonInc
    LocalCommon.AppName = "ShellFuelTargetedOffers.asmx"

    Try
      If LocalCommon.LXSadoConn.State = ConnectionState.Closed Then LocalCommon.Open_LogixXS()
      If LocalCommon.LRTadoConn.State = ConnectionState.Closed Then LocalCommon.Open_LogixRT()
      ReturnCode = Copient.CustomerAbstract.RETURN_CODE.OK


      ' check if the customer is householded, if so then use the household CustomerPK
      HHPK = GetCustomerHHPK(CustomerPK)
      If HHPK > 0 Then CustomerPK = HHPK

      ' Find the stored value balance 
      LocalCommon.QueryStr = "select SVProgramID, Sum(IsNull(QtyEarned,0)) - Sum(IsNull(QtyUsed,0)) as Quantity " & _
                          "from StoredValue with (NoLock) " & _
                          "where CustomerPK=" & CustomerPK & " and StatusFlag=1 and ExpireDate >= getdate() and Deleted=0 " & _
                          "group by SVProgramID "
      WriteDebug("Database Call - get customer sv balances sum  request", DebugState.CurrentTime)
      dt = LocalCommon.LXS_Select()
      WriteDebug("Database Call - get customer sv balances sum  response", DebugState.CurrentTime)

      If Not IncludeZeroBalances Then
        LocalCommon.QueryStr &= "having (Sum(IsNull(QtyEarned,0)) - Sum(IsNull(QtyUsed,0))) > 0;"
      End If

      dt = LocalCommon.LXS_Select
      RowCt = dt.Rows.Count
      If RowCt > 0 Then
        ReDim Balances(RowCt - 1)
        For i = 0 To RowCt - 1
          SVBal = New Copient.CustomerLookup.StoredValueBalance
          SVBal.SVProgramID = LocalCommon.NZ(dt.Rows(i).Item("SVProgramID"), 0)
          SVBal.Units = LocalCommon.NZ(dt.Rows(i).Item("Quantity"), 0)

          WriteDebug("Database Call - get customer sv program type  request", DebugState.CurrentTime)
          ' get the type of stored value program to determine the balance 
          LocalCommon.QueryStr = "select case when SVTypeID > 1 then  convert(decimal(12,3), " & SVBal.Units & "  * Value) " & _
                            "       else CONVERT(int, " & SVBal.Units & " * Value) end as Balance " & _
                            "from StoredValuePrograms with (NoLock) " & _
                            "where SVProgramID = " & SVBal.SVProgramID & " and Deleted=0;"
          dt2 = LocalCommon.LRT_Select
          WriteDebug("Database Call - get customer sv program type  response", DebugState.CurrentTime)
          If dt2.Rows.Count > 0 Then
            SVBal.Balance = LocalCommon.NZ(dt2.Rows(0).Item("Balance"), 0)
          End If

          Balances(i) = SVBal
        Next
      End If
    Catch ex As Exception
      Copient.Logger.Write_Log(ShellWSLogFile, ex.Message, True)
    Finally
      If LocalCommon.LXSadoConn.State = ConnectionState.Open Then LocalCommon.Close_LogixXS()
      If LocalCommon.LRTadoConn.State = ConnectionState.Open Then LocalCommon.Close_LogixRT()
    End Try
    Return Balances


  End Function
  Public Function GetCustomerHHPK(ByVal CustomerPK As Long) As Long
    Dim HHPK As Long = 0
    Dim dt As DataTable
    WriteDebug("IN GetCustomerHHPK", DebugState.CurrentTime)

    Dim LocalCommon As New Copient.CommonInc
    LocalCommon.AppName = "ShellFuelTargetedOffers.asmx"

    Try
      If LocalCommon.LXSadoConn.State = ConnectionState.Closed Then LocalCommon.Open_LogixXS()

      LocalCommon.QueryStr = "select HHPK from Customers with (NoLock) where CustomerPK =" & CustomerPK

      WriteDebug("Database Call - get customer HHPK  request", DebugState.CurrentTime)
      dt = LocalCommon.LXS_Select
      WriteDebug("Database Call - get customer HHPK  response", DebugState.CurrentTime)
      If dt.Rows.Count > 0 Then
        HHPK = LocalCommon.NZ(dt.Rows(0).Item("HHPK"), 0)
      End If
    Catch ex As Exception
      Copient.Logger.Write_Log(ShellWSLogFile, ex.Message, True)
    Finally
      If LocalCommon.LXSadoConn.State = ConnectionState.Open Then LocalCommon.Close_LogixXS()
    End Try

    
    Return HHPK
  End Function
  Public Function ComputeDiscountAmount(ByVal rewardAmount As Decimal, ByVal SVProQuantity As Integer, ByVal CustSVBalance As Integer, ByVal DistributionLimit As Integer) As Decimal
WriteDebug("IN ComputeDiscountAmount", DebugState.CurrentTime)

    Dim DiscountGiven As Decimal
    Try
      MaxNoOfDistribution = CustSVBalance \ SVProQuantity
      If MaxNoOfDistribution <= DistributionLimit Then
        DiscountGiven = MaxNoOfDistribution * rewardAmount
      Else
        DiscountGiven = DistributionLimit * rewardAmount
      End If
    Catch ex As Exception
      Copient.Logger.Write_Log(ShellWSLogFile, ex.Message, True)
    End Try
    
    Return DiscountGiven

  End Function

  Public Function ComputeAdjustmentAmount(ByVal MaximumDiscount As Decimal, ByVal OfferRewardAmount As Decimal, ByVal SVQuantity As Integer) As Integer
    WriteDebug("IN ComputeAdjustmentAmount", DebugState.CurrentTime)

    Dim AdjustAmount As Integer
    Dim NoOfDiscounts As Decimal

    Try

      NoOfDiscounts = MaximumDiscount / OfferRewardAmount
      AdjustAmount = NoOfDiscounts * SVQuantity
    Catch ex As Exception
      Copient.Logger.Write_Log(ShellWSLogFile, ex.Message, True)
    End Try
    Return AdjustAmount
  End Function
  Public Sub UpdateSTD(ByVal CustomerPK As Long, ByVal TotalDiscount As Decimal)
    WriteDebug("IN UpdateSTD", DebugState.CurrentTime)

    Dim LocalCommon As New Copient.CommonInc
    LocalCommon.AppName = "ShellFuelTargetedOffers.asmx"

    Dim HHPK As Long

    Try
      If LocalCommon.LXSadoConn.State = ConnectionState.Closed Then LocalCommon.Open_LogixXS()

      HHPK = GetCustomerHHPK(CustomerPK)
      If HHPK > 0 Then CustomerPK = HHPK

      LocalCommon.QueryStr = "Update Customers set CurrYearSTD = CurrYearSTD + " & TotalDiscount & " where CustomerPK=" & CustomerPK
      WriteDebug("Database Call - Update CurrYearSTD where customerpk = " & CustomerPK & "  - request", DebugState.CurrentTime)
      LocalCommon.LXS_Execute()
      WriteDebug("Database Call - Update CurrYearSTD where customerpk = " & CustomerPK & "  - response", DebugState.CurrentTime)
      Copient.Logger.Write_Log(ShellWSLogFile, "Updated YTD savings", True)
    Catch ex As Exception
      Copient.Logger.Write_Log(ShellWSLogFile, ex.Message, True)
    Finally
      If LocalCommon.LXSadoConn.State = ConnectionState.Open Then LocalCommon.Close_LogixXS()
    End Try

  End Sub
  Private Function GetItems(ByVal shellxml As String) As Items
    WriteDebug("IN GetItems", DebugState.CurrentTime)

    Dim XMLDOc As New XmlDocument()
    'Dim ItemNode As XmlNode

    Try
      XMLDOc.LoadXml(shellxml)

      For Each Element As System.Xml.XmlElement In XMLDOc.SelectNodes("//*")

        If Element.Name.ToUpper = "ITEM" Then
          For Each childnode As System.Xml.XmlNode In Element
            If childnode.Name.ToUpper = "ITEMID" Then
              objItems.ItemID = childnode.InnerText
            ElseIf childnode.Name.ToUpper = "TYPE" Then
              objItems.Type = childnode.InnerText
            ElseIf childnode.Name.ToUpper = "QUALIFIER" Then
              objItems.Qualifier = childnode.InnerText
            End If
          Next
        End If
      Next
  

    Catch ex As Exception
      Copient.Logger.Write_Log(ShellWSLogFile, ex.Message, True)
    End Try

    Return objItems
  End Function

  Private Function GetCustomerDetails(ByVal customerpk As Long) As String()
    WriteDebug("IN GetCustomerDetails", DebugState.CurrentTime)

    Dim dt As DataTable
    Dim CustomerDetails(2) As String
    Dim HHPK As Long

    Dim LocalCommon As New Copient.CommonInc
    LocalCommon.AppName = "ShellFuelTargetedOffers.asmx"
  Try
    If LocalCommon.LXSadoConn.State = ConnectionState.Closed Then LocalCommon.Open_LogixXS()

    WriteDebug("Database Call - Get Customer Details for customerPK " & customerpk & "  - request", DebugState.CurrentTime)
    LocalCommon.QueryStr = "select ExtCardId, CardTypeID from CardIds with (NoLock) where CustomerPK =" & customerpk

    dt = LocalCommon.LXS_Select

    If dt.Rows.Count > 0 Then


      CustomerDetails(0) = LocalCommon.NZ(dt.Rows(0).Item("CardTypeID"), 0)

    End If


    HHPK = GetCustomerHHPK(customerpk)
    If HHPK > 0 Then
      LocalCommon.QueryStr = "select ExtCardId from CardIds with (NoLock) where CustomerPK =" & HHPK
      dt = LocalCommon.LXS_Select
      WriteDebug("Database Call - Get Customer Details for customerPK " & customerpk & "  - response", DebugState.CurrentTime)
      If dt.Rows.Count > 0 Then
                    'CustomerDetails(0) = 1
                    'No need to decrypt as this is a private function and is called from PopulatetransactionTables where we pass encrypted value.
        CustomerDetails(1) = IIf(IsDBNull(dt.Rows(0).Item("ExtCardId")), 0,MyCryptLib.SQL_StringDecrypt(dt.Rows(0).Item("ExtCardId").ToString()))

      End If
    End If
  Catch ex As Exception
    Copient.Logger.Write_Log(ShellWSLogFile, ex.Message, True)
  Finally
    If LocalCommon.LXSadoConn.State = ConnectionState.Open Then LocalCommon.Close_LogixXS()
  End Try
    Return CustomerDetails
  End Function
  
  Private Function GetAdminUserID() As long
    WriteDebug("IN GetAdminUserID", DebugState.CurrentTime)

    Dim LocalCommon As New Copient.CommonInc
    LocalCommon.AppName = "ShellFuelTargetedOffers.asmx"
	  Dim dt As DataTable
	  Dim AdminUserID As Long = 1

  Try
    If LocalCommon.LRTadoConn.State = ConnectionState.Closed Then LocalCommon.Open_LogixRT()

	  LocalCommon.QueryStr = "select AdminUserID from AdminUsers with (NoLock) where UserName ='Shell'"
    WriteDebug("Database Call - GetAdminUserID  request", DebugState.CurrentTime)
	  dt = LocalCommon.LRT_Select
    WriteDebug("Database Call - GetAdminUserID  response", DebugState.CurrentTime)
	  If dt.Rows.Count > 0 Then
	    AdminUserID = LocalCommon.NZ(dt.Rows(0).Item("AdminUserID"), 1)	
	  End If
  Catch ex As Exception
    Copient.Logger.Write_Log(ShellWSLogFile, ex.Message, True)
  Finally
    If LocalCommon.LRTadoConn.State = ConnectionState.Open Then LocalCommon.Close_LogixRT()
  End Try    
    Return AdminUserID
  End Function
  
  Public Function ReFormatCard(ByVal CardID As String) As String
    Dim FormattedCard As String = ""
    If CardID.Length > 10 Then
      FormattedCard = CardID.Substring(CardID.Length - 10)
    End If
    Return FormattedCard
  End Function

  Private Sub WriteDebug(ByVal sText As String, ByVal mode As DebugState)
    If bTimeLogOn Then
      If (bTimeLogHasBeenCalled  AndAlso mode <> DebugState.EndTime) Then
        mode = DebugState.CurrentTime
      Else If (Not bTimeLogHasBeenCalled AndAlso mode <> DebugState.EndTime)
        mode = DebugState.BeginTime
        bTimeLogHasBeenCalled = True
      End If

      Dim TotalSeconds As Double
      Dim sIndent As String
      Select Case mode
        Case DebugState.BeginTime
          ' first call
          DebugStartTimes.Clear()
          DebugStartTimes.Add(Now)
          If DebugStartTimes.Count = 1 Then
            Dim sIPAddress As String
            WriteLog(scDashes, MessageType.Debug)
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
            sText = sIndent & sText & " - End elapsed time: " & Format(TotalSeconds, "00.0000") & "(sec)" & vbCrLf
            DebugStartTimes.RemoveAt(DebugStartTimes.Count - 1)
            bTimeLogLastCall = True
            bTimeLogHasBeenCalled = False
          End If
        Case Else
          ' interim call
          If DebugStartTimes.Count > 0 Then
            TotalSeconds = Now.Subtract(DebugStartTimes(DebugStartTimes.Count - 1)).TotalSeconds
            sIndent = New String(" "c, 2 * (DebugStartTimes.Count - 1))
            sText = sIndent & sText & " - Current elapsed time: " & Format(TotalSeconds, "00.0000") & "(sec)"
          End If
      End Select
      WriteLog(sText, MessageType.Debug)
    End If
  End Sub

  Private Sub WriteLog(ByVal sText As String, ByVal eType As MessageType)
    Dim sFileName As String
    Dim sLogText As String = ""
    Dim sEnableLogging As String
    Dim TotalSeconds As Double
    Dim sIndent As String


    If eType = MessageType.Debug Then
      sLogText = "[" & Format(Date.Now, "hh:mm:sszzz") & " (Type=" & eType.ToString & ")]  {ThreadID: " & AppDomain.GetCurrentThreadId.ToString() & "}" & sText
    Else
      If eType <> MessageType.Info Then
        sText = sText.Replace(ControlChars.CrLf, " ")
      End If
      If sInputForLog.Length > 0 Then
        sLogText = "[" & Format(Date.Now, "hh:mm:sszzz") & " (Type=" & eType.ToString & ")]  {ThreadID: " & AppDomain.GetCurrentThreadId.ToString() & "}" & sInputForLog & ControlChars.CrLf
        sInputForLog = ""
      End If
      sLogText = sLogText & "[" & Format(Date.Now, "hh:mm:sszzz") & " (Type=" & eType.ToString & ")]  {ThreadID: " & AppDomain.GetCurrentThreadId.ToString() & "}" & sText
    End If
    sLogLines = sLogLines & sLogText '& vbCrLf
      Try
        Copient.Logger.Write_Log(ShellWSLogFile, sLogLines, True)
        sLogLines = ""
      Catch ex As Exception
        Try
          MyCommon.Error_Processor(, "WriteLog", sAppName, sInstallationName)
        Catch
        End Try
      Finally
        If (bTimeLogLastCall) Then
          If DebugStartTimes.Count > 0 Then
            Try
              sLogText = "[" & Format(Date.Now, "hh:mm:sszzz") & " (Type=Debug)] "
              TotalSeconds = Now.Subtract(DebugStartTimes(DebugStartTimes.Count - 1)).TotalSeconds
              sIndent = New String(" "c, 2 * (DebugStartTimes.Count - 1))
              sLogLines = sLogText & sIndent & "** Log entry time to write " & " - End elapsed time: " & Format(TotalSeconds, "00.0000") & "(sec)"
              DebugStartTimes.RemoveAt(DebugStartTimes.Count - 1)
              Copient.Logger.Write_Log(ShellWSLogFile, sLogLines, True)
            Catch
            End Try
          End If
        End If
      End Try
  End Sub

End Class
'1. Do we need to check that is the location associated with offer??  -- No, as per the latest shell locations design
'2. Expired offers?
'3. HouseHolding?
'4.'check for the existence and get back the corresponding loyalty card no. will only check for alt id card types?
'5. There will be only 1 offer then what is the point of having a query to find the offers? can we not have some system option/configuration from where we can get the offer id. 
'6. Check for the existence and get back the corresponding loyalty card no. will only check for alt id card types?
