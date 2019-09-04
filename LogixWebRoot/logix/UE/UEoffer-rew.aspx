<%@ Page Language="vb" Debug="true" CodeFile="ue-cb.vb" Inherits="UECB" %>

<%@ Import Namespace="Copient.Localization" %>
<%@ Import Namespace="CMS.AMS.Models" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) 
%>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Text.RegularExpressions" %>
<%@ Import Namespace="CMS.AMS" %>
<%@ Import Namespace="CMS.AMS.Contract" %>
<%@ Import Namespace="System.Globalization" %>
<%
    ' *****************************************************************************
    ' * FILENAME: UEoffer-rew.aspx 
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

    Dim CopientFileName As String = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
    Dim CopientFileVersion As String = "7.3.1.138972"
    Dim CopientProject As String = "Copient Logix"
    Dim CopientNotes As String = ""
    Dim SystemCacheData As ICacheData
    Dim MyCommon As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    Dim MyCpe As New Copient.CPEOffer
    Dim rst As DataTable
    Dim row As DataRow
    Dim rst2 As DataTable
    Dim row2 As DataRow
    Dim rst3 As DataTable
    Dim row3 As DataRow
    Dim rstPT As DataTable
    Dim rowPT As DataRow
    Dim Value As Decimal
    Dim OfferID As Long
    Dim RewardID As Long
    Dim DeliverableID As Long
    Dim PKID As Long
    Dim MessageID As Long = 0
    Dim FrankID As Long = 0
    Dim DiscountID As Long = 0
    Dim PassThruRewardID As Long = 0
    Dim ParentROID As Long = 0
    Dim ROID As Long = 0
    Dim Name As String = ""
    Dim DeleteGraphicURL As String = ""
    Dim AddTouchPtURL As String = ""
    Dim UrlTokens As String = ""
    Dim DeliverableType As Integer
    Dim AddOptionArray As New BitArray(8, True)
    Dim MessageTypeLabel As String = ""
    Dim index As Integer = 0
    Dim i As Integer = 0
    Dim DeleteBtnDisabled As String = ""
    Dim Details As StringBuilder
    Dim IsTemplate As Boolean = False
    Dim IsTemplateVal As String = "Not"
    Dim ActiveSubTab As Integer = 24
    Dim FromTemplate As Boolean = False
    Dim Disallow_Rewards As Boolean = False
    Dim RewardsCount As Integer = 0
    Dim IsCustomerAssigned As Boolean = False
    Dim AccumEligible As Boolean = False
    Dim infoMessage As String = ""
    Dim modMessage As String = ""
    Dim Handheld As Boolean = False
    Dim Rewards As String() = Nothing
    Dim LockedStatus As String() = Nothing
    Dim LoopCtr As Integer = 0
    Dim RewardDisabled As String = ""
    Dim BannersEnabled As Boolean = True
    Dim StatusCode As Copient.LogixInc.STATUS_FLAGS = Copient.LogixInc.STATUS_FLAGS.STATUS_UNKNOWN
    Dim StatusText As String = ""
    Dim TierLevels As Integer = 1
    Dim t As Integer = 1
    Dim IsFooterOffer As Boolean = False
    Dim DiscountWorthy As Boolean = False
    Dim EngineID As Integer = 2
    Dim EngineSubTypeID As Integer = 0
    Dim TempQuerystr As String
    Dim Localizer As Copient.Localization
    Dim LocalizerService As CMS.AMS.LocalizationService
    Dim AmountTypeID As Integer
    Dim DefaultLanguageID As Integer = 0
    Dim DeferCalcToTotal As Boolean = False
    Dim m_PassThruReward As IPassThroughRewards
    Dim GiftCardID As Long = 0
    Dim ProximityMessageID As Long = 0
    Dim DeleteAllowed As Boolean = True
    Dim m_EditOfferRegardlessOfBuyer As Boolean = True
    Dim EPMInstalled As Boolean = False
    Dim objPreferenceService As IPreferenceRewardService
    Dim objCouponService As ICouponRewardService
    Dim m_PreferenceService As IPreferenceService
    Dim IntegrationVals As New Copient.CommonInc.IntegrationValues
    Dim PreferenceRewardID As Int32 = 0
    Dim m_DiscountPGService As IDiscountRewardService
    Dim m_CustCondService As ICustomerGroupCondition
    Dim disableGiftCard As Boolean = False
    Dim custAppResult As AMSResult(Of CustomerApproval) = New AMSResult(Of CustomerApproval)()
    Dim restrictRewardforRPOS As Boolean = False
    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If

    Response.Expires = 0
    MyCommon.AppName = "UEoffer-rew.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
    Localizer = New Copient.Localization(MyCommon)

    Integer.TryParse(MyCommon.Fetch_SystemOption(1), DefaultLanguageID)
    BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")

    CurrentRequest.Resolver.AppName = MyCommon.AppName
    m_PassThruReward = CurrentRequest.Resolver.Resolve(Of IPassThroughRewards)()
    objPreferenceService = CurrentRequest.Resolver.Resolve(Of IPreferenceRewardService)()
    m_PreferenceService = CurrentRequest.Resolver.Resolve(Of IPreferenceService)()
    objCouponService = CurrentRequest.Resolver.Resolve(Of ICouponRewardService)()
    m_DiscountPGService = CurrentRequest.Resolver.Resolve(Of IDiscountRewardService)()
    m_CustCondService = CurrentRequest.Resolver.Resolve(Of ICustomerGroupCondition)()
    SystemCacheData = CurrentRequest.Resolver.Resolve(Of ICacheData)()
    OfferID = Request.QueryString("OfferID")
    RewardID = Request.QueryString("RewardID")
    DeliverableID = MyCommon.Extract_Val(Request.QueryString("DeliverableID"))
    PKID = MyCommon.Extract_Val(Request.QueryString("PKID"))
    MessageID = MyCommon.Extract_Val(Request.QueryString("MessageID"))
    FrankID = MyCommon.Extract_Val(Request.QueryString("FrankID"))
    DiscountID = MyCommon.Extract_Val(Request.QueryString("DiscountID"))
    PassThruRewardID = MyCommon.Extract_Val(Request.QueryString("PassThruRewardID"))
    DeliverableType = MyCommon.Extract_Val(Request.QueryString("action"))
    GiftCardID = MyCommon.Extract_Val(Request.QueryString("GiftCardID"))
    ProximityMessageID = MyCommon.Extract_Val(Request.QueryString("ProximityMessageID"))
    PreferenceRewardID = MyCommon.Extract_Val(Request.QueryString("PreferenceRewardID"))
    If (OfferID = 0) Then
        Response.Status = "301 Moved Permanently"
        Response.AddHeader("Location", "UEoffer-gen.aspx")
    End If
    restrictRewardforRPOS = (SystemCacheData.GetSystemOption_UE_ByOptionId(234) = "1")
    'This would redirect to summary page if offer is awaiting deployment or if offer is awaiting recommendations.
    CheckIfValidOffer(MyCommon, OfferID)

    Dim isTranslatedOffer As Boolean = MyCommon.IsTranslatedUEOffer(OfferID, MyCommon)
    Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249) = "1", True, False)

    Dim bEnableAdditionalLockoutRestrictionsOnOffers As Boolean = IIf(MyCommon.Fetch_SystemOption(260) = "1", True, False)
    Dim bOfferEditable As Boolean = MyCommon.IsOfferEditablePastLockOutPeriod(Logix.UserRoles.EditOfferPastLockoutPeriod, MyCommon, OfferID)

    EPMInstalled = MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER, IntegrationVals)
    MyCommon.QueryStr = "select IncentiveName, IsTemplate, FromTemplate, EngineSubTypeID, DeferCalcToTotal,buy.ExternalBuyerId as BuyerID from CPE_Incentives cpe with (NoLock) left outer join Buyers as buy with (nolock) on buy.BuyerId= cpe.BuyerId  where IncentiveID=" & OfferID
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
        If (MyCommon.Fetch_UE_SystemOption(168) = "1" AndAlso MyCommon.NZ(rst.Rows(0).Item("BuyerID"), "") <> "") Then
            Name = "Buyer " + rst.Rows(0).Item("BuyerID").ToString() + " - " + MyCommon.NZ(rst.Rows(0).Item("IncentiveName"), "").ToString()
        Else
            Name = MyCommon.NZ(rst.Rows(0).Item("IncentiveName"), Copient.PhraseLib.Lookup("term.unknown", LanguageID))
        End If
        ' Name = MyCommon.NZ(rst.Rows(0).Item("IncentiveName"), "")
        IsTemplate = MyCommon.NZ(rst.Rows(0).Item("IsTemplate"), False)
        FromTemplate = MyCommon.NZ(rst.Rows(0).Item("FromTemplate"), False)
        EngineSubTypeID = MyCommon.NZ(rst.Rows(0).Item("EngineSubTypeID"), 0)
        DeferCalcToTotal = MyCommon.NZ(rst.Rows(0).Item("DeferCalcToTotal"), False)
    End If

    MyCommon.QueryStr = "select RewardOptionID, TierLevels from CPE_RewardOptions with (NoLock) " &
                        "where IncentiveID=" & OfferID & " and TouchResponse=0 and Deleted=0;"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
        ParentROID = MyCommon.NZ(rst.Rows(0).Item("RewardOptionID"), 0)
        If (RewardID = 0) Then
            RewardID = ParentROID
        End If
        TierLevels = MyCommon.NZ(rst.Rows(0).Item("TierLevels"), 1)
    End If

    MyCommon.QueryStr = "select PassThruRewardID from PassThruRewards with (NoLock);"
    rst = MyCommon.LRT_Select
    AddOptionArray = New BitArray(8 + rst.Rows.Count, True)

    IsFooterOffer = MyCpe.IsFooterOffer(OfferID)
    If IsFooterOffer Then
        AddOptionArray = New BitArray(8 + rst.Rows.Count, False)
        AddOptionArray.Set(1, True)
    End If

    MyCommon.QueryStr = "select CG.CustomerGroupID, Name, ExcludedUsers from CPE_IncentiveCustomerGroups as ICG with (NoLock) " &
                        "left join CustomerGroups as CG with (NoLock) on CG.CustomerGroupID=ICG.CustomerGroupID " &
                        "where RewardOptionID=" & RewardID & ";"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
        IsCustomerAssigned = True
    End If
    MyCommon.QueryStr = "select IncentiveAttributeID from CPE_IncentiveAttributes as IA with (NoLock) " &
                        "where RewardOptionID=" & RewardID & " and Deleted=0;"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
        IsCustomerAssigned = True
    End If

    If ExistsProximityMessageReward(MyCommon, ParentROID) AndAlso Not String.IsNullOrEmpty(Request.QueryString("mode")) Then
        If Request.QueryString("mode") <> "DeleteProximityMessage" And Request.QueryString("mode").ToLower().Contains("delete") Then
            Dim RewardCount As Int32
            If AnyOtherRewardTypeExists(MyCommon, ParentROID, DELIVERABLE_TYPES.PROXIMITY_MESSAGE, RewardCount) AndAlso RewardCount = 1 Then
                infoMessage = Copient.PhraseLib.Lookup("error.rewarddeletepmrconflict", LanguageID)
                DeleteAllowed = False
            End If
        End If
    End If
    If DeleteAllowed Then
        If (Request.QueryString("mode") = "DeleteGraphic") Then
            RemoveGraphic(OfferID, DeliverableID)
        ElseIf (Request.QueryString("mode") = "DeleteCashierMsg") Then
            If (DeliverableID > 0 AndAlso MessageID > 0) Then
                MyCommon.QueryStr = "delete from CPE_CashierMessageTiers with (RowLock) where MessageID=" & MessageID & ";"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "delete from CPE_CashierMessages with (RowLock) where MessageID=" & MessageID & ";"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "delete from CPE_Deliverables with (RowLock) where DeliverableID=" & DeliverableID & ";"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1, DeferCalcToTotal=0 where IncentiveID=" & OfferID & ";"
                MyCommon.LRT_Execute()
                ResetOfferApprovalStatus(OfferID)
                DeferCalcToTotal = False
                MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Reward.deletecmsg", LanguageID))
            End If
        ElseIf (Request.QueryString("mode") = "DeleteMembership") Then
            If (DeliverableID > 0 AndAlso DeliverableType > 0 AndAlso OfferID > 0) Then
                MyCommon.QueryStr = "delete from CPE_DeliverableCustomerGroupTiers with (RowLock) where DeliverableID=" & DeliverableID & ";"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "delete from CPE_Deliverables with (RowLock) where RewardOptionID=" & RewardID & " and RewardOptionPhase=3 and DeliverableTypeID=" & DeliverableType & ";"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
                MyCommon.LRT_Execute()
                ResetOfferApprovalStatus(OfferID)
                MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Reward.deletemembership", LanguageID))
            End If
        ElseIf (Request.QueryString("mode") = "DeletePoints") Then
            If (DeliverableID > 0 AndAlso OfferID > 0) Then
                MyCommon.QueryStr = "delete from CPE_DeliverablePointTiers where DPPKID in " &
                                    "(select PKID from CPE_DeliverablePoints where DeliverableID=" & DeliverableID & ");"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "delete from CPE_DeliverablePoints where DeliverableID=" & DeliverableID & ";"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "delete from CPE_Deliverables with (RowLock) where DeliverableID=" & DeliverableID & ";"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
                MyCommon.LRT_Execute()
                ResetOfferApprovalStatus(OfferID)
                'Check if accumulation message needs to be removed
                MyCommon.QueryStr = "dbo.pa_CPE_AccumMsgEligible"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@ROID", SqlDbType.Int, 4).Value = RewardID
                MyCommon.LRTsp.Parameters.Add("@AccumEligible", SqlDbType.Bit, 1).Direction = ParameterDirection.Output
                MyCommon.LRTsp.ExecuteNonQuery()
                AccumEligible = MyCommon.LRTsp.Parameters("@AccumEligible").Value
                MyCommon.Close_LRTsp()
                If Not (AccumEligible) Then
                    'Delete any accumulation messages
                    MyCommon.QueryStr = "delete from CPE_Deliverables with (RowLock) where DeliverableID in " &
                                        "(select D.DeliverableID from CPE_RewardOptions as RO " &
                                        " inner join CPE_Deliverables as D on RO.RewardOptionID=D.RewardOptionID " &
                                        " where RO.Deleted=0 and RO.IncentiveID=" & OfferID & " and RewardOptionPhase=2 and DeliverableTypeID=4);"
                    MyCommon.LRT_Execute()
                End If
                MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Reward.deletepoints", LanguageID))
            End If
        ElseIf (Request.QueryString("mode") = "DeleteTrackableCouponReward") Then
            If (DeliverableID > 0 AndAlso OfferID > 0) Then
                Dim Result As Boolean = objCouponService.DeleteTrackableCouponReward(DeliverableID, OfferID)
                If Result Then
                    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("term.deleted", LanguageID) + " " + Copient.PhraseLib.Lookup("term.trackablecoupon", LanguageID).ToLower() + " " + Copient.PhraseLib.Lookup("term.reward", LanguageID).ToLower())
                End If
            End If
        ElseIf (Request.QueryString("mode") = "DeleteStoredValue") Then
            If (DeliverableID > 0 AndAlso OfferID > 0) Then
                MyCommon.QueryStr = "delete from CPE_DeliverableStoredValueTiers where DSVPKID in " &
                                    "(select PKID from CPE_DeliverableStoredValue where DeliverableID=" & DeliverableID & ");"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "delete from CPE_DeliverableStoredValue where DeliverableID=" & DeliverableID & ";"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "delete from CPE_Deliverables with (RowLock) where DeliverableID=" & DeliverableID & ";"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
                MyCommon.LRT_Execute()
                ResetOfferApprovalStatus(OfferID)
            End If
            MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Reward.deletesv", LanguageID))
        ElseIf (Request.QueryString("mode") = "DeleteMonStoredValue") Then
            If (DeliverableID > 0 AndAlso OfferID > 0) Then
                MyCommon.QueryStr = "delete from CPE_DeliverableMonSVTranslations where  SVProgramID in (select SVProgramID from CPE_DeliverableMonStoredValue where DeliverableID=" & DeliverableID & ");"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "delete from CPE_DeliverableMonStoredValue where DeliverableID=" & DeliverableID & ";"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "delete from CPE_Deliverables with (RowLock) where DeliverableID=" & DeliverableID & ";"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
                MyCommon.LRT_Execute()
                ResetOfferApprovalStatus(OfferID)
            End If
            MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("cpe_reward.deletemsv", LanguageID))
        ElseIf (Request.QueryString("mode") = "DeletePrintedMsg") Then
            If (DeliverableID > 0 AndAlso OfferID > 0) Then
                MyCommon.QueryStr = "delete from PrintedMessageTiers with (RowLock) where MessageID=" & MessageID & ";"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "delete from PrintedMessages with (RowLock) where MessageID=" & MessageID & ";"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "delete from CPE_Deliverables with (RowLock) where DeliverableID=" & DeliverableID & " and RewardOptionPhase=3 and DeliverableTypeID=4;"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
                MyCommon.LRT_Execute()
                ResetOfferApprovalStatus(OfferID)
                MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Reward.deletepmsg", LanguageID))
            End If
        ElseIf (Request.QueryString("mode") = "DeleteDiscount") Then
            If (DeliverableID > 0 AndAlso OfferID > 0) Then
                MyCommon.QueryStr = "delete from CPE_SpecialPricing with (RowLock) where DiscountID=" & DiscountID & ";"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "delete from CPE_DiscountTiers with (RowLock) where DiscountID=" & DiscountID & ";"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "delete from CPE_Discounts with (RowLock) where DiscountID=" & DiscountID & ";"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "delete from CPE_Deliverables with (RowLock) where DeliverableID=" & DeliverableID & " and RewardOptionPhase=3 and DeliverableTypeID=2;"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
                MyCommon.LRT_Execute()
                ResetOfferApprovalStatus(OfferID)
                MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Reward.deletediscount", LanguageID))
            End If
        ElseIf (Request.QueryString("mode") = "DeleteFrankingMsg") Then
            If (DeliverableID > 0 AndAlso FrankID > 0) Then
                MyCommon.QueryStr = "delete from CPE_FrankingMessageTiers with (RowLock) where FrankID=" & FrankID & ";"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "delete from CPE_FrankingMessages with (RowLock) where FrankID=" & FrankID & ";"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "delete from CPE_Deliverables with (RowLock) where DeliverableID=" & DeliverableID & ";"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
                MyCommon.LRT_Execute()
                ResetOfferApprovalStatus(OfferID)
                MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Reward.deletefmsg", LanguageID))
            End If
        ElseIf (Request.QueryString("mode") = "DeletePassThru") Then
            If (DeliverableID > 0 AndAlso OfferID > 0) Then
                MyCommon.QueryStr = "delete from PassThruTierValues where PTPKID in (select PKID from PassThrus where DeliverableID=" & DeliverableID & ");"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "delete from PassThruTiers where PTPKID in " &
                                    "(select PKID from PassThrus where DeliverableID=" & DeliverableID & ");"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "delete from PassThrus where DeliverableID=" & DeliverableID & ";"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "delete from CPE_Deliverables with (RowLock) where DeliverableID=" & DeliverableID & ";"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
                MyCommon.LRT_Execute()
                ResetOfferApprovalStatus(OfferID)
            End If
            MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Reward.deletepassthru", LanguageID))

        ElseIf (Request.QueryString("mode") = "DeleteGiftCard") Then
            If (DeliverableID > 0 AndAlso OfferID > 0) Then
                MyCommon.QueryStr = "select ID from GiftCardTier where GiftCardID=" & GiftCardID & ";"
                rst3 = MyCommon.LRT_Select()
                'Delete all the tier translations of GiftCard
                If rst3.Rows.Count > 0 Then
                    For Each row In rst3.Rows
                        MyCommon.QueryStr = "DELETE FROM GIFTCARDTIERTRANSLATION with (RowLock) where GiftCardTierId=" & MyCommon.NZ(row.Item("ID"), 0) & ";"
                        MyCommon.LRT_Execute()
                    Next
                End If
                'Delete permissions 
                MyCommon.QueryStr = "delete from TemplateFieldPermissions with (RowLock) where OfferID=" & OfferID & " " &
                                "and FieldID in (select FieldID from UIFields where PageName='UEoffer-rew-giftcard.aspx');"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "DELETE FROM GIFTCARDTIER with (RowLock) WHERE GIFTCARDID=" & GiftCardID & ";"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "DELETE FROM GIFTCARD with (RowLock) WHERE ID=" & GiftCardID & ";"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "delete from CPE_Deliverables with (RowLock) where DeliverableID=" & DeliverableID & " and RewardOptionPhase=3 and DeliverableTypeID=13;"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
                MyCommon.LRT_Execute()
                ResetOfferApprovalStatus(OfferID)
                MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("Reward.deleteGiftCard", LanguageID))
            End If
        ElseIf (Request.QueryString("mode") = "DeletePreferenceReward") Then
            If (DeliverableID > 0 AndAlso OfferID > 0) Then
                objPreferenceService.DeletePreferenceReward(DeliverableID, PreferenceRewardID, AdminUserID, OfferID)

                MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("Reward.deletePreferenceReward", LanguageID))
            End If
        ElseIf (Request.QueryString("mode") = "DeleteProximityMessage") Then
            If (DeliverableID > 0 AndAlso OfferID > 0) Then
                MyCommon.QueryStr = "select ID from ProximityMessageTier where ProximityMessageID=" & ProximityMessageID & ";"
                rst3 = MyCommon.LRT_Select()
                'Delete all the tier translations of ProximityMessageTier
                If rst3.Rows.Count > 0 Then
                    For Each row In rst3.Rows
                        MyCommon.QueryStr = "DELETE FROM ProximityMessageTierTranslation with (RowLock) where ProximityMessageTierId=" & MyCommon.NZ(row.Item("ID"), 0) & ";"
                        MyCommon.LRT_Execute()
                    Next
                End If
                'Delete permissions 
                MyCommon.QueryStr = "delete from TemplateFieldPermissions with (RowLock) where OfferID=" & OfferID & " " &
                                "and Deliverableid= " & DeliverableID & " and FieldID in (select FieldID from UIFields where PageName='UEoffer-rew-proximitymsg.aspx');"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "DELETE FROM ProximityMessageTier with (RowLock) WHERE ProximityMessageID=" & ProximityMessageID & ";"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "DELETE FROM ProximityMessage with (RowLock) WHERE ID=" & ProximityMessageID & ";"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "delete from CPE_Deliverables with (RowLock) where DeliverableID=" & DeliverableID & " and RewardOptionPhase=3 and DeliverableTypeID=14;"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
                MyCommon.LRT_Execute()
                ResetOfferApprovalStatus(OfferID)
                MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.ueofferrew-pmr-deleted", LanguageID))
            End If

        End If
    End If
    'update the template permission for Conditions
    If (Request.QueryString("save") <> "") Then
        If (Request.QueryString("IsTemplate") = "IsTemplate") Then
            ' time to update the status bits for the templates
            Dim form_Disallow_Rewards As Integer = 0
            If (Request.QueryString("Disallow_Rewards") = "on") Then
                form_Disallow_Rewards = 1
            End If
            MyCommon.QueryStr = "update TemplatePermissions with (RowLock) set Disallow_Rewards=" & form_Disallow_Rewards &
                                " where OfferID=" & OfferID & ";"
            MyCommon.LRT_Execute()
        End If

        'Update the lock status for each condition
        Rewards = Request.QueryString.GetValues("rew")
        LockedStatus = Request.QueryString.GetValues("locked")
        If (Not Rewards Is Nothing AndAlso Not LockedStatus Is Nothing AndAlso Rewards.Length = LockedStatus.Length) Then
            For LoopCtr = 0 To Rewards.GetUpperBound(0)
                MyCommon.QueryStr = "update CPE_Deliverables with (RowLock) set DisallowEdit=" & LockedStatus(LoopCtr) & " " &
                                    "where DeliverableID=" & Rewards(LoopCtr) & ";"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "select DeliverableTypeID from CPE_Deliverables with (RowLock) where DeliverableID=" & Rewards(LoopCtr)
                rst = MyCommon.LRT_Select
                If (Convert.ToInt32(rst.Rows(0)(0)) = 14) AndAlso LockedStatus(LoopCtr) = 1 Then
                    MyCommon.QueryStr = "update TemplateFieldPermissions with (RowLock) set Editable=0 where DeliverableID=" & Rewards(LoopCtr) & ";"
                    MyCommon.LRT_Execute()
                End If

            Next
        End If

    End If

    If (IsTemplate Or FromTemplate) Then
        ' lets dig the permissions if its a template
        MyCommon.QueryStr = "select * from TemplatePermissions with (NoLock) where OfferID=" & OfferID & ";"
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
            For Each row In rst.Rows
                ' ok there are some rows for the template
                Disallow_Rewards = MyCommon.NZ(row.Item("Disallow_Rewards"), True)
            Next
        End If
    End If

    If (IsTemplate) Then
        ActiveSubTab = 25
        IsTemplateVal = "IsTemplate"
    Else
        ActiveSubTab = 24
        IsTemplateVal = "Not"
    End If
    m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
    If Not IsTemplate Then
        DeleteBtnDisabled = IIf(Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And Disallow_Rewards) And Not IsOfferWaitingForApproval(OfferID), "", " disabled=""disabled""")
    Else
        DeleteBtnDisabled = IIf((Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer), "", " disabled=""disabled""")
    End If

    SetDeleteBtnDisabled(DeleteBtnDisabled) 'method found in included file GraphicReward.aspx

    Send_HeadBegin("term.offer", "term.rewards", OfferID)
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    If (IsTemplate) Then
        Send("<style type=""text/css"">")
        Send("  * html #tblReward {")
        Send("    position: relative;")
        Send("    top: -20px;")
        Send("  }")
        Send("</style>")
    End If
    Send_Scripts()
%>
<script type="text/javascript">
    function LoadDocument(url) {
        location = url;
    }

    function openTouchptReward(index, roid) {
        var pageName = "";
        var qryStr = "";
        var tpElem = document.getElementById("newrewtouchpt" + index);

        if (tpElem != null) {
            var rewType = tpElem.options[tpElem.options.selectedIndex].value

            if (rewType >= 1 && rewType <= 12) {
                qryStr = "?RewardID=<%Sendb(RewardID)%>&OfferID=<%Sendb(OfferID)%>&tp=1&roid=" + roid;
                if (rewType == 1) {
                    pageName = "UEoffer-rew-graphic.aspx";
                } else if (rewType == 2) {
                    pageName = "UEoffer-rew-discount.aspx";
                } else if (rewType == 4) {
                    pageName = "UEoffer-rew-pmsg.aspx";
                } else if (rewType == 5) {
                    qryStr += "&action=5"
                    pageName = "UEoffer-rew-membership.aspx";
                } else if (rewType == 6) {
                    /* Remove membership rewards are currently disabled */
                } else if (rewType == 7) {
                    /* Silent deliverable not supported */
                } else if (rewType == 8) {
                    pageName = "UEoffer-rew-point.aspx";
                } else if (rewType == 9) {
                    pageName = "UEoffer-rew-cmsg.aspx";
                } else if (rewType == 10) {
                    /* Franking not supported */
                } else if (rewType == 11) {
                    pageName = "UEoffer-rew-sv.aspx";
                } else if (rewType == 12) {
                    /* Pass-thrus not supported */
                }
                openPopup(pageName + qryStr);
            }
        }
    }

    function updateLocked(elemName, bChecked) {
        var elem = document.getElementById(elemName);

        if (elem != null) {
            elem.value = (bChecked) ? "1" : "0";
        }
    }
    $(document).ready(function () {
        var savedTimeVal = document.getElementById('savedTime');
        var offerIDVal = document.getElementById('OfferID');
        if (savedTimeVal != null && offerIDVal != null) {
            var savedTime = new Date(savedTimeVal.value).getTime();
            var presentTime = new Date().getTime();
            var seconds = (presentTime - savedTime) / 1000;
            if (seconds > 2) {
                $.support.cors = true;
                $.ajax({
                    type: "POST",
                    url: "/Connectors/AjaxProcessingFunctions.asmx/GetLockedSystemOptions",
                    data: JSON.stringify({ offerID: offerIDVal.value }),
                    contentType: "application/json; charset=utf-8",
                    dataType: "json"
                })
                    .done(function (data) {
                        if (data.d == "true") {
                            window.location.href = window.location.href.replace("UEOffer-rew.aspx", "UEOffer-sum.aspx");
                        }
                    })

            }
        }
    });
</script>
<script runat="server">
    Dim DiscountexclusionPGList As List(Of DiscountProductGroup)
    Dim resultDiscountExcludedPGList As AMSResult(Of List(Of DiscountProductGroup))
    'Dim sb As New StringBuilder()
    Function DisplayDiscountExclusionGroups(ByVal DiscountPGService As IDiscountRewardService, ByVal Discountid As Integer, ByRef infoMessage As String, ByRef MyCommon As Copient.CommonInc, ByVal languageId As Integer) As String
        Dim sb As New StringBuilder()
        Dim extBuyerId As string

        resultDiscountExcludedPGList = DiscountPGService.GetAllExclusionGroups(Discountid)
        If (resultDiscountExcludedPGList.ResultType = AMSResultType.Success) Then
            DiscountexclusionPGList = resultDiscountExcludedPGList.Result
            If DiscountexclusionPGList.Count > 0 Then
                sb.Append(" ")
                sb.Append(Copient.PhraseLib.Lookup("term.excluding", languageId))
                sb.Append(" ")
                For Each dpcpg As DiscountProductGroup In DiscountexclusionPGList
                    If (MyCommon.IsEngineInstalled(9) And MyCommon.Fetch_UE_SystemOption(168) = "1" And dpcpg.BuyerId > 0) Then
                        extBuyerId = MyCommon.GetExternalBuyerId(dpcpg.BuyerId)
                        sb.Append("<a href=""/logix/pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(dpcpg.ProductGroupId, 0) & """>" & "Buyer " & extBuyerId & " - " & MyCommon.NZ(dpcpg.ProductGroupName, "") & "</a>, ")
                    Else
                        sb.Append("<a href=""/logix/pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(dpcpg.ProductGroupId, 0) & """>" & MyCommon.NZ(dpcpg.ProductGroupName, "") & "</a>, ")
                    End If
                    'Send("<a href=""/logix/pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(pcpg.ProductGroupId, 0) & """>Prod:" & MyCommon.NZ(pcpg.ProductGroupId, "") & "</a>,")
                Next
                'Remove the last comma 
                sb.Remove(sb.Length - 2, 2)
                ' Send(sb.ToString())
                'Send("</li>")
            End If
        Else
            infoMessage = resultDiscountExcludedPGList.PhraseString
        End If
        Return sb.ToString()
    End Function
    Function ExistsProximityMessageReward(ByRef MyCommon As Copient.CommonInc, ByVal ROID As Long) As Boolean
        Dim existsPMR As Boolean = False
        Dim rst As DataTable
        MyCommon.QueryStr = "select Count(DeliverableId) from CPE_Deliverables where DeliverableTypeId = @DeliverableTypeID And RewardOptionId=@ROID AND Deleted=0"
        MyCommon.DBParameters.Add("@ROID", SqlDbType.BigInt).Value = ROID
        MyCommon.DBParameters.Add("@DeliverableTypeID", SqlDbType.Int).Value = DELIVERABLE_TYPES.PROXIMITY_MESSAGE
        rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)

        If Not rst Is Nothing AndAlso CInt(rst.Rows(0)(0)) > 0 Then
            existsPMR = True
        End If
        Return existsPMR
    End Function
    Function AnyOtherRewardTypeExists(ByRef MyCommon As Copient.CommonInc, ByVal ROID As Long, ByVal DeliverableTypeID As Int32, Optional ByRef RewardCount As Int32 = 0) As Boolean
        Dim existsAnyOtherReward As Boolean = True
        Dim rst As DataTable
        MyCommon.QueryStr = "select Count(DeliverableId) from CPE_Deliverables where DeliverableTypeId <> @DeliverableTypeID And RewardOptionId=@ROID AND Deleted=0"
        MyCommon.DBParameters.Add("@ROID", SqlDbType.BigInt).Value = ROID
        MyCommon.DBParameters.Add("@DeliverableTypeID", SqlDbType.Int).Value = DeliverableTypeID
        rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)

        If Not rst Is Nothing Then
            If CInt(rst.Rows(0)(0)) = 0 Then
                existsAnyOtherReward = False
            Else
                RewardCount = CInt(rst.Rows(0)(0))
            End If
        End If
        Return existsAnyOtherReward

    End Function
    Sub GetGcrQueryFlags(ByRef MyCommon As Copient.CommonInc, ByVal ParentROID As Long, ByRef ProductCondition As Boolean, ByRef GCRPercentOffAllowed As Boolean, ByVal OfferId As Long, ByRef UnitTypeId As Int32)
        Dim rst2 As DataTable
        MyCommon.QueryStr = "select QtyUnitType from dbo.CPE_IncentiveProductGroups WHERE RewardOptionID=@ParentROID AND Deleted=0"
        MyCommon.DBParameters.Add("@ParentROID", SqlDbType.BigInt).Value = ParentROID
        rst2 = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)

        Dim productConditionPriceTypeExist As Boolean = False
        Dim productConditionOthersExist As Boolean = False
        If rst2.Rows.Count > 0 Then
            ProductCondition = True

            For Each row1 In rst2.Rows
                If CInt(row1("QtyUnitType")) = CPEUnitTypes.Dollars Then
                    productConditionPriceTypeExist = True
                    UnitTypeId = CInt(row1("QtyUnitType"))
                End If
                If CInt(row1("QtyUnitType")) <> CPEUnitTypes.Dollars Then
                    productConditionOthersExist = True
                End If
            Next

            If productConditionPriceTypeExist AndAlso productConditionOthersExist Then
                GCRPercentOffAllowed = False
            ElseIf productConditionPriceTypeExist AndAlso GetOfferLimit(MyCommon, OfferId) <> 1 Then 'The offer limit must be one when product condition of US Dollars exists to enable percentoff GCR. See AL-5916
                GCRPercentOffAllowed = False
            End If
        End If
    End Sub
    Function GetPreferenceNameByID(ByRef MyCommon As Copient.CommonInc, ByVal PreferenceID As Int64) As String
        Dim dst As New DataTable
        Dim PreferenceName As String = ""
        MyCommon.QueryStr = "SELECT Name As PreferenceName FROM Preferences Where PreferenceID = @PreferenceID"
        MyCommon.DBParameters.Add("@PreferenceID", SqlDbType.BigInt).Value = PreferenceID
        dst = MyCommon.ExecuteQuery(Copient.DataBases.PrefManRT)
        Return PreferenceName
    End Function
    Function GetOfferLimit(ByRef MyCommon As Copient.CommonInc, ByVal OfferID As Long) As Int32
        Dim offerLimit As Int32
        Dim rst2 As DataTable
        MyCommon.QueryStr = "select P3DistQtyLimit from dbo.CPE_Incentives WHERE IncentiveId = @OfferId AND Deleted=0"
        MyCommon.DBParameters.Add("@OfferId", SqlDbType.BigInt).Value = OfferID
        rst2 = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)

        If rst2.Rows.Count > 0 AndAlso rst2.Rows(0)(0) <> Nothing Then
            offerLimit = CInt(rst2.Rows(0)(0))
        End If
        Return offerLimit
    End Function
</script>
<%
    Send_HeadEnd()
    If (IsTemplate) Then
        Send_BodyBegin(11)
    Else
        Send_BodyBegin(1)
    End If
    Send_Bar(Handheld)
    Send_Help(CopientFileName)
    Send_Logos()
    Send_Tabs(Logix, 2)
    Send_Subtabs(Logix, ActiveSubTab, 6, , OfferID)

    If (Logix.UserRoles.AccessOffers = False AndAlso Not IsTemplate) Then
        Send_Denied(1, "perm.offers-access")
        GoTo done
    End If
    If (Logix.UserRoles.AccessTemplates = False AndAlso IsTemplate) Then
        Send_Denied(1, "perm.offers-access-templates")
        GoTo done
    End If
    If (Logix.UserRoles.AccessInstantWinOffers = False AndAlso EngineID = 2 AndAlso EngineSubTypeID = 1) Then
        Send_Denied(1, "perm.offers-accessinstantwin")
        GoTo done
    End If
    If (BannersEnabled AndAlso Not Logix.IsAccessibleOffer(AdminUserID, OfferID)) Then
        Send("<script type=""text/javascript"" language=""javascript"">")
        Send("  function updateCookie() { return true; } ")
        Send("</script>")
        Send_Denied(1, "banners.access-denied-offer")
        Send_BodyEnd()
        GoTo done
    End If

    If (Request.QueryString("addGlobal") <> "") Then
        Dim rewChoice As Integer = MyCommon.Extract_Val(Request.QueryString("newrewglobal"))
        If (rewChoice > 0) Then
            Select Case rewChoice
                Case 1
                    Send("<script type=""text/javascript"">openPopup('UEoffer-rew-graphic.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "&Phase=3')</script>")
                    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-graphic", LanguageID))
                Case 2
                    Dim m_Offer As IOffer = CMS.AMS.CurrentRequest.Resolver.Resolve(Of IOffer)()
                    Dim listProdCondition As AMSResult(Of List(Of RegularProductCondition)) = m_Offer.GetRegularProductConditionsByOfferId(OfferID)
                    If (listProdCondition.Result.Count > 0 AndAlso listProdCondition.Result.Item(0).IncentiveProductGroupId > 0) Then
                        Send("<script type=""text/javascript"">openPopup('UEoffer-rew-discount.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "&Phase=3')</script>")
                        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-discount", LanguageID))
                    Else
                        infoMessage = Copient.PhraseLib.Lookup("CPEoffer-rew.DiscountDisallowed", LanguageID)
                    End If
                Case 3
                    'Not used
                Case 4
                    Send("<script type=""text/javascript"">openPopup('UEoffer-rew-pmsg.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "&Phase=3')</script>")
                    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-pmsg", LanguageID))
                Case 5
                    Send("<script type=""text/javascript"">openPopup('UEoffer-rew-membership.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "&action=5&Phase=3')</script>")
                    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-member", LanguageID))
                Case 6
                    'Revoke membership -- not used
                Case 7
                    'Silent deliverable -- not used
                Case 8
                    Send("<script type=""text/javascript"">openPopup('UEoffer-rew-point.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "&Phase=3&New=1')</script>")
                    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-points", LanguageID))
                Case 9
                    Send("<script type=""text/javascript"">openPopup('UEoffer-rew-cmsg.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "&Phase=3')</script>")
                    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-cmsg", LanguageID))
                Case 10
                    Send("<script type=""text/javascript"">openPopup('UEoffer-rew-franking.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "&Phase=3')</script>")
                    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-franking", LanguageID))
                Case 11
                    Send("<script type=""text/javascript"">openPopup('UEoffer-rew-sv.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "&Phase=3')</script>")
                    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-sv", LanguageID))
                Case 12
                    Send("<script type=""text/javascript"">openPopup('/logix/Offer-Rew-XMLPassThru.aspx?OfferID=" & OfferID & "&PassThruRewardID=" & (rewChoice - 12) & "&Phase=3&DeliverableID=0')</script>")
                    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-passthru", LanguageID))

                    'Giftcard
                Case 13
                    Dim GCRPercentOffAllowed As Boolean = True
                    Dim ProductCondition As Boolean = False
                    Dim UnitTypeId As Int32
                    GetGcrQueryFlags(MyCommon, ParentROID, ProductCondition, GCRPercentOffAllowed, OfferID, UnitTypeId)
                    Dim RestrictProrationTypeToAllConditional As Boolean = False
                    If GCRPercentOffAllowed AndAlso UnitTypeId = CPEUnitTypes.Dollars Then
                        RestrictProrationTypeToAllConditional = True
                    End If
                    'If product Condition is not set, dont show percent off in valuetype ddl.
                    Send("<script type=""text/javascript"">openPopup('UEoffer-rew-giftCard.aspx?OfferID=" & OfferID & "&RewardID=" & RewardID & "&productCondition=" & ProductCondition & "&PercentOffAllowed=" & GCRPercentOffAllowed & "&RestrictProrationTypeToAllConditional=" & RestrictProrationTypeToAllConditional & "&Phase=3&DeliverableID=0')</script>")
                    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.ueofferrew-giftcard", LanguageID))

                Case 14
                    If AnyOtherRewardTypeExists(MyCommon, ParentROID, DELIVERABLE_TYPES.PROXIMITY_MESSAGE) Then
                        Send("<script type=""text/javascript"">openPopup('UEoffer-rew-proximitymsg.aspx?OfferID=" & OfferID & "&RewardID=" & RewardID & "&Phase=3')</script>")
                        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.ueofferrew-pmr", LanguageID))
                    Else
                        infoMessage = Copient.PhraseLib.Lookup("error.nopmrwithoutotherreward", LanguageID)
                    End If
                Case 15
                    Send("<script type=""text/javascript"">openPopup('UEoffer-rew-pref.aspx?OfferID=" & OfferID & "&RewardID=" & RewardID & "&Phase=3&DeliverableID=0')</script>")
                    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.ueofferrew-preference", LanguageID))
                Case 16
                    Dim m_Offer As IOffer = CMS.AMS.CurrentRequest.Resolver.Resolve(Of IOffer)()
                    Dim listProdCondition As AMSResult(Of List(Of RegularProductCondition)) = m_Offer.GetRegularProductConditionsByOfferId(OfferID)
                    If (listProdCondition.Result.Count > 0 AndAlso listProdCondition.Result.Item(0).IncentiveProductGroupId > 0) Then
                        Send("<script type=""text/javascript"">openPopup('UEoffer-rew-msv.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "&Phase=3')</script>")
                        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-msv", LanguageID))
                    Else
                        infoMessage = Copient.PhraseLib.Lookup("term.ProductConditionMandatory", LanguageID)
                    End If
                Case 17
                    Send("<script type=""text/javascript"">openPopup('UEoffer-rew-tc.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "&Phase=3')</script>")
                    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.ueofferrew-tc", LanguageID))
                Case Else
                    MyCommon.QueryStr = "select PassThruRewardID, Name, PhraseID from PassThruRewards with (NoLock) order by PassThruRewardID;"
                    rstPT = MyCommon.LRT_Select
                    If rstPT.Rows.Count > 0 Then
                        Send("<script type=""text/javascript"">openPopup('UEoffer-rew-passthru.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "&PassThruRewardID=" & (rewChoice - 12) & "&Phase=3')</script>")
                        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-passthru", LanguageID))
                    End If
            End Select
        End If
    End If
%>
<form action="UEoffer-rew.aspx" id="mainform" name="mainform">
    <input type="hidden" id="OfferID" name="OfferID" value="<% Sendb(OfferID) %>" />
    <input type="hidden" id="IsTemplate" name="IsTemplate" value="<% Sendb(IsTemplateVal) %>" />
    <input type="hidden" id="savedTime" name="savedTime" value="<%=DateTime.Now()%>" />
    <div id="intro">
        <%
            If (IsTemplate) Then
                Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & ": " & MyCommon.TruncateString(Name, 50) & "</h1>")
            Else
                Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & ": " & MyCommon.TruncateString(Name, 50) & "</h1>")
            End If
        %>
        <div id="controls">
            <%
                m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
                If (Logix.UserRoles.EditTemplates And IsTemplate And m_EditOfferRegardlessOfBuyer AndAlso (Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
                    Send_Save()
                End If
                If MyCommon.Fetch_SystemOption(75) Then
                    If (OfferID > 0 And Logix.UserRoles.AccessNotes AndAlso (Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
                        Send_NotesButton(3, OfferID, AdminUserID)
                    End If
                End If
            %>
        </div>
    </div>
    <div id="main">
        <%
            MyCommon.QueryStr = "select StatusFlag from CPE_Incentives where IncentiveID=" & OfferID & ";"
            rst = MyCommon.LRT_Select
            StatusText = Logix.GetOfferStatus(OfferID, LanguageID, StatusCode)
            If Not IsTemplate Then
                If (MyCommon.NZ(rst.Rows(0).Item("StatusFlag"), 0) <> 2) Then
                    If (StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_ACTIVE) AndAlso (MyCommon.NZ(rst.Rows(0).Item("StatusFlag"), 0) > 0) Then
                        modMessage = Copient.PhraseLib.Lookup("alert.modpostdeploy", LanguageID)
                        Send("<div id=""modbar"">" & modMessage & "</div>")
                    End If
                End If
            End If
            If (infoMessage <> "") Then
                Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
            End If

            ' Send the status bar, but not if it's a new offer or a template, or if there's already a modMessage being shown.
            If (Not IsTemplate AndAlso modMessage = "") Then
                MyCommon.QueryStr = "select IncentiveID from CPE_Incentives with (NoLock) where CreatedDate=LastUpdate and IncentiveID=" & OfferID & ";"
                rst3 = MyCommon.LRT_Select
                If (rst3.Rows.Count = 0) Then
                    Send_Status(OfferID, 2)
                End If
            End If
        %>
        <div id="column">
            <div class="box" id="rewards">
                <h2>
                    <span>
                        <% Sendb(Copient.PhraseLib.Lookup("term.rewards", LanguageID))%></span>
                </h2>
                <br />
                <table class="list" id="tblReward" summary="<% Sendb(Copient.PhraseLib.Lookup("term.rewards", LanguageID))%>">
                    <thead>
                        <tr>
                            <th align="left" scope="col" class="th-del">
                                <% Sendb(Copient.PhraseLib.Lookup("term.delete", LanguageID))%>
                            </th>
                            <th align="left" scope="col" class="th-type">
                                <% Sendb(Copient.PhraseLib.Lookup("term.type", LanguageID))%>
                            </th>
                            <th align="left" scope="col" class="th-subtype">
                                <% Sendb(Copient.PhraseLib.Lookup("term.subtype", LanguageID))%>
                            </th>
                            <th align="left" scope="col" class="th-details" colspan="<% Sendb(TierLevels)%>">
                                <% Sendb(Copient.PhraseLib.Lookup("term.details", LanguageID))%>
                            </th>
                            <% If (IsTemplate OrElse FromTemplate) Then%>
                            <th align="left" scope="col" class="th-locked">
                                <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
                            </th>
                            <% End If%>
                        </tr>
                    </thead>
                    <tbody>
                        <%
                            ' Discount rewards
                            t = 1
                            MyCommon.QueryStr = "select DISC.DiscountID, DISC.Name, DISC.DiscountTypeId, DISC.DiscountAmount, DISC.DiscountedProductGroupID as SelectedPG, " &
                                                "DISC.ItemLimit, DISC.WeightLimit, DISC.DollarLimit, DISC.ExcludedProductGroupID as ExcludedPG, DISC.DiscountAmount, " &
                                                "DISC.ChargebackDeptID, DISC.AmountTypeID, DISC.L1Cap, DISC.L2DiscountAmt, DISC.L2AmountTypeID, DISC.L2Cap, DISC.L3DiscountAmt, DISC.L3AmountTypeID, " &
                                                "DISC.DecliningBalance, DISC.RetroactiveDiscount, DISC.UserGroupID, DISC.BestDeal, DISC.AllowNegative, DISC.ComputeDiscount, " &
                                                "D.DeliverableID, D.DisallowEdit, AT.AmountTypeID, AT.PhraseID as AmountPhraseID, DT.PhraseID as DiscountPhraseID " &
                                                "from CPE_Deliverables D with (NoLock) " &
                                                "inner join CPE_Discounts DISC with (NoLock) on D.OutputID=DISC.DiscountID " &
                                                "left join CPE_AmountTypes AT with (NoLock) on AT.AmountTypeID=DISC.AmountTypeID " &
                                                "left join CPE_DiscountTypes DT with (NoLock) on DT.DiscountTypeID=DISC.DiscountTypeID " &
                                                "where D.RewardOptionPhase=3 and D.RewardOptionID=" & RewardID & " and D.DeliverableTypeID=2;"
                            rst = MyCommon.LRT_Select()
                            If (rst.Rows.Count > 0) Then
                                RewardsCount += rst.Rows.Count
                                AddOptionArray.Set(0, False) 'Changing this to True will allow multiple discount rewards. False will allow only one discount
                                Send("<tr class=""shadeddark"">")
                                Send("  <td colspan=""" & 3 & """>")
                                Send("    <h3>" & Copient.PhraseLib.Lookup("CPE-rew-disc.header", LanguageID) & "</h3>")
                                Send("  </td>")
                                If TierLevels = 1 Then
                                    Send("  <td></td>")
                                Else
                                    For t = 1 To TierLevels
                                        Send("  <td>")
                                        Send("    <b>" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</b>")
                                        Send("  </td>")
                                    Next
                                End If
                                t = 1
                                If (IsTemplate Or FromTemplate) Then
                                    Send("  <td></td>")
                                End If
                                Send("</tr>")
                                m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
                                For Each row In rst.Rows
                                    If Not IsTemplate Then
                                        RewardDisabled = IIf((Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False)) And Not IsOfferWaitingForApproval(OfferID)), "", " disabled=""disabled""")
                                    Else
                                        RewardDisabled = IIf((Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer), "", " disabled=""disabled""")
                                    End If
                                    Send("<tr class=""shaded"">")
                                    Send("  <td>")
                                    If ((Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
                                        Send("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & RewardDisabled & " value=""X"" ")
                                        Send("onClick=""if(confirm('" & Copient.PhraseLib.Lookup("reward.confirmdelete", LanguageID) & "')) LoadDocument('UEoffer-rew.aspx?mode=DeleteDiscount&OfferID=" & OfferID & "&DiscountID=" & MyCommon.NZ(row.Item("DiscountID"), 0) & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "')"" />")
                                    End If
                                    Send("</td>")
                                    Send("  <td><a href=""javascript:openPopup('UEoffer-rew-discount.aspx?OfferID=" & OfferID & "&RewardID=" & RewardID & "&DiscountID=" & MyCommon.NZ(row.Item("DiscountID"), 0) & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "')"">" & Copient.PhraseLib.Lookup("term.discount", LanguageID) & "</a></td>")
                                    Send("  <td>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("AmountPhraseID"), 0), LanguageID) & "</td>")

                                    ' Find the per-tier details, and build up the details string:
                                    MyCommon.QueryStr = "select DT.PKID, DT.TierLevel, DT.DiscountAmount, DT.ReceiptDescription, DT.ItemLimit, DT.WeightLimit, DT.DollarLimit " &
                                                        "from CPE_DiscountTiers as DT with (NoLock) " &
                                                        "where DT.DiscountID=" & MyCommon.NZ(row.Item("DiscountID"), 0) & ";"
                                    rst2 = MyCommon.LRT_Select
                                    If rst2.Rows.Count = 0 Then
                                        Send("  <td colspan=""" & TierLevels & """>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                                    Else
                                        While t <= TierLevels
                                            If t > rst2.Rows.Count Then
                                                Send("  <td>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                                            Else
                                                Details = New StringBuilder(200)
                                                AmountTypeID = (MyCommon.NZ(row.Item("AmountTypeID"), 0))
                                                Select Case AmountTypeID
                                                    Case 1, 5, 9, 10, 11, 12
                                                        Details.Append(Localizer.FormatCurrency_ForOffer(CDec(MyCommon.NZ(rst2.Rows(t - 1).Item("DiscountAmount"), 0)), RewardID).ToString(MyCommon.GetAdminUser.Culture) & " " & StrConv(Copient.PhraseLib.Lookup("term.off", LanguageID), VbStrConv.Lowercase) & "&nbsp;")
                                                    Case 3
                                                        Details.Append(Math.Round(CDec(MyCommon.NZ(rst2.Rows(t - 1).Item("DiscountAmount"), 0)), 2).ToString(MyCommon.GetAdminUser.Culture) & "% " & " " & StrConv(Copient.PhraseLib.Lookup("term.off", LanguageID), VbStrConv.Lowercase) & "&nbsp;")
                                                    Case 4
                                                        Details.Append(Copient.PhraseLib.Lookup("term.free", LanguageID) & "&nbsp;")
                                                    Case 2, 6, 13, 14, 15, 16
                                                        Details.Append(Localizer.GetCached_Currency_Symbol(RewardID) & Math.Round(CDec(MyCommon.NZ(rst2.Rows(t - 1).Item("DiscountAmount"), 0)), Localizer.GetCached_Currency_Precision(RewardID)).ToString(MyCommon.GetAdminUser.Culture) & "&nbsp;")
                                                    Case 7
                                                        'do nothing
                                                    Case 8
                                                        'Special pricing
                                                        Details = New StringBuilder(200)
                                                        i = 0
                                                        MyCommon.QueryStr = "select Value, LevelID from CPE_SpecialPricing as SP with (NoLock) where DiscountTierID=" & MyCommon.NZ(rst2.Rows(t - 1).Item("PKID"), "") & ";"
                                                        rst3 = MyCommon.LRT_Select
                                                        If rst3.Rows.Count > 0 Then
                                                            For Each row3 In rst3.Rows
                                                                Value = Math.Round(MyCommon.NZ(rst3.Rows(i).Item("Value"), 0), Localizer.GetCached_Currency_Precision(RewardID))
                                                                Details.Append(Localizer.GetCached_Currency_Symbol(RewardID) & Value.ToString())
                                                                If i < rst3.Rows.Count Then
                                                                    Details.Append(", ")
                                                                Else
                                                                    Details.Append(" ")
                                                                End If
                                                                i += 1
                                                            Next
                                                        Else
                                                            Details.Append(Copient.PhraseLib.Lookup("term.undefined", LanguageID) & " ")
                                                        End If
                                                    Case Else
                                                        Details.Append(MyCommon.NZ(rst2.Rows(t - 1).Item("DiscountAmount"), "") & "&nbsp;")
                                                End Select

                                                If (MyCommon.NZ(row.Item("DiscountTypeID"), 0) = 4 OrElse MyCommon.NZ(row.Item("DiscountTypeID"), 0) = 5) AndAlso MyCommon.NZ(row.Item("SelectedPG"), 0) = 0 Then
                                                    Details.Append(StrConv(Copient.PhraseLib.Lookup("term.conditionalproducts", LanguageID), VbStrConv.Lowercase))
                                                ElseIf MyCommon.NZ(row.Item("SelectedPG"), 0) = 0 Then
                                                    Details.Append(StrConv(Copient.PhraseLib.Lookup("term.nothing", LanguageID), VbStrConv.Lowercase))
                                                Else
                                                    MyCommon.QueryStr = "select Name,buyerid from ProductGroups with (NoLock) where ProductGroupID=" & row.Item("SelectedPG")
                                                    rst3 = MyCommon.LRT_Select()
                                                    For Each row3 In rst3.Rows
                                                        If row.Item("SelectedPG") = 1 Then

                                                            Details.Append(StrConv(MyCommon.NZ(row3.Item("Name"), ""), VbStrConv.Lowercase))
                                                        Else
                                                            If (MyCommon.IsEngineInstalled(9) And MyCommon.Fetch_UE_SystemOption(168) = "1" And Not IsDBNull(row3.Item("Buyerid"))) Then
                                                                Dim buyerid As Integer = row3.Item("Buyerid")
                                                                Dim externalBuyerid = MyCommon.GetExternalBuyerId(buyerid)
                                                                Details.Append("<a href=""../pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("SelectedPG"), "") & """>" & "Buyer " & externalBuyerid & " - " & MyCommon.SplitNonSpacedString(MyCommon.NZ(row3.Item("Name"), ""), 25) & "</a>")
                                                            Else
                                                                Details.Append("<a href=""../pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("SelectedPG"), "") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row3.Item("Name"), ""), 25) & "</a>")
                                                            End If

                                                            Details.Append(DisplayDiscountExclusionGroups(m_DiscountPGService, row.Item("DiscountID"), infoMessage, MyCommon, LanguageID))
                                                        End If
                                                    Next

                                                End If
                                                If MyCommon.NZ(row.Item("DiscountTypeID"), 0) = 2 Then
                                                    Details.Append(" (" & StrConv(Copient.PhraseLib.Lookup("term.department", LanguageID), VbStrConv.Lowercase) & ")")
                                                ElseIf MyCommon.NZ(row.Item("DiscountTypeID"), 0) = 6 Then
                                                    Details.Append(" (" & StrConv(Copient.PhraseLib.Lookup("term.grouplevel", LanguageID), VbStrConv.Lowercase) & ")")
                                                ElseIf MyCommon.NZ(row.Item("DiscountTypeID"), 0) = 3 Then
                                                    Details.Append(" (" & StrConv(Copient.PhraseLib.Lookup("term.basket", LanguageID), VbStrConv.Lowercase) & ")")
                                                    Details.Append(DisplayDiscountExclusionGroups(m_DiscountPGService, row.Item("DiscountID"), infoMessage, MyCommon, LanguageID))
                                                End If
                                                'AMS-685 Show multiple exclusion groups for discount reward
                                                'If MyCommon.NZ(row.Item("ExcludedPG"), 0) = 0 Then
                                                'Else
                                                '  MyCommon.QueryStr = "select Name from ProductGroups with (NoLock) where ProductGroupID=" & row.Item("ExcludedPG")
                                                '  rst3 = MyCommon.LRT_Select()
                                                '  For Each row3 In rst3.Rows
                                                '    Details.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.excluding", LanguageID), VbStrConv.Lowercase) & " ")
                                                '    If row.Item("ExcludedPG") = 1 Then
                                                '      Details.Append(StrConv(MyCommon.NZ(row3.Item("Name"), ""), VbStrConv.Lowercase))
                                                '    Else
                                                '      Details.Append("<a href=""../pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("ExcludedPG"), "") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row3.Item("Name"), ""), 25) & "</a>")
                                                '    End If
                                                '  Next
                                                'End If

                                                If MyCommon.NZ(row.Item("L1Cap"), 0) > 0 Then
                                                    Details.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.upto", LanguageID), VbStrConv.Lowercase) & " " & Localizer.GetCached_Currency_Symbol(RewardID) & Math.Round(CDec(row.Item("L1Cap")), Localizer.GetCached_Currency_Precision(RewardID)).ToString(MyCommon.GetAdminUser.Culture))
                                                End If

                                                If MyCommon.NZ(rst2.Rows(t - 1).Item("ItemLimit"), 0) = 0 And MyCommon.NZ(rst2.Rows(t - 1).Item("WeightLimit"), 0) = 0 And MyCommon.NZ(rst2.Rows(t - 1).Item("DollarLimit"), 0) = 0 Then
                                                    Details.Append(",&nbsp;" & StrConv(Copient.PhraseLib.Lookup("term.unlimited", LanguageID), VbStrConv.Lowercase))
                                                Else
                                                    If MyCommon.NZ(row.Item("DiscountTypeID"), 0) <> 3 Then
                                                        Details.Append(",&nbsp;" & StrConv(Copient.PhraseLib.Lookup("term.limit", LanguageID), VbStrConv.Lowercase) & "&nbsp;")
                                                        If rst2.Rows(t - 1).Item("ItemLimit") > 0 Then
                                                            Details.Append(rst2.Rows(t - 1).Item("ItemLimit") & " " & StrConv(Copient.PhraseLib.Lookup("term.items", LanguageID), VbStrConv.Lowercase))
                                                            If rst2.Rows(t - 1).Item("WeightLimit") > 0 Or rst2.Rows(t - 1).Item("DollarLimit") > 0 Then
                                                                Details.Append(" / ")
                                                            End If
                                                        End If
                                                        If rst2.Rows(t - 1).Item("DollarLimit") > 0 Then
                                                            Details.Append(Localizer.GetCached_Currency_Symbol(RewardID) & Math.Round(CDec(rst2.Rows(t - 1).Item("DollarLimit")), Localizer.GetCached_Currency_Precision(RewardID)).ToString(MyCommon.GetAdminUser.Culture))
                                                            If rst2.Rows(t - 1).Item("WeightLimit") > 0 Then
                                                                Details.Append(" / ")
                                                            End If
                                                        End If
                                                        If rst2.Rows(t - 1).Item("WeightLimit") > 0 Then
                                                            Details.Append(Localizer.Round_Quantity(CDec(rst2.Rows(t - 1).Item("WeightLimit")), RewardID, 5).ToString(MyCommon.GetAdminUser.Culture))
                                                        End If
                                                    End If
                                                End If
                                                ' If there are multiple levels, this will display their details on a second line.
                                                If (MyCommon.NZ(row.Item("L2DiscountAmt"), 0) > 0) And MyCommon.NZ(row.Item("L2AmountTypeID"), 0) = 3 Then
                                                    Details.Append("<br />(" & Copient.PhraseLib.Lookup("term.over", LanguageID) & " " & Localizer.GetCached_Currency_Symbol(RewardID) & Math.Round(CDec(MyCommon.NZ(row.Item("L1Cap"), "0")), Localizer.GetCached_Currency_Precision(RewardID)).ToString(MyCommon.GetAdminUser.Culture) & ", ")
                                                    Details.Append(row.Item("L2DiscountAmt") & "% " & StrConv(Copient.PhraseLib.Lookup("term.off", LanguageID), VbStrConv.Lowercase))
                                                    If (MyCommon.NZ(row.Item("L2Cap"), 0) > 0) Then
                                                        Details.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.upto", LanguageID), VbStrConv.Lowercase) & " " & Localizer.GetCached_Currency_Symbol(RewardID) & Math.Round(CDec(row.Item("L2Cap")), Localizer.GetCached_Currency_Precision(RewardID)).ToString(MyCommon.GetAdminUser.Culture))
                                                    End If
                                                    Details.Append(")")
                                                    If (MyCommon.NZ(row.Item("L3DiscountAmt"), 0) > 0) And MyCommon.NZ(row.Item("L3AmountTypeID"), 0) = 3 Then
                                                        Details.Append("<br />(" & Copient.PhraseLib.Lookup("term.over", LanguageID) & " " & Localizer.GetCached_Currency_Symbol(RewardID) & Math.Round(CDec(MyCommon.NZ(row.Item("L2Cap"), 0)), Localizer.GetCached_Currency_Precision(RewardID)).ToString(MyCommon.GetAdminUser.Culture) & ", ")
                                                        Details.Append(row.Item("L3DiscountAmt") & "% " & StrConv(Copient.PhraseLib.Lookup("term.off", LanguageID), VbStrConv.Lowercase) & ")")
                                                    End If
                                                End If
                                                Send("  <td>" & Details.ToString & "</td>")
                                            End If
                                            t += 1
                                        End While
                                        t = 1
                                    End If
                                    If (IsTemplate) Then
                                        Send("  <td class=""templine"">")
                                        Send("    <input type=""checkbox"" id=""chkLocked1"" name=""chkLocked"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, " checked=""checked""", "") & " onclick=""javascript:updateLocked('lockDisc" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "', this.checked);"" />")
                                        Send("    <input type=""hidden"" id=""rewDisc" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""rew"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ />")
                                        Send("    <input type=""hidden"" id=""lockDisc" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""locked"" value=""" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "1", "0") & """ />")
                                        Send("  </td>")
                                    ElseIf (FromTemplate) Then
                                        Send("  <td class=""templine"">" & Copient.PhraseLib.Lookup("term." & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "yes", "no"), LanguageID) & "</td>")
                                    End If
                                    Send("</tr>")
                                Next
                            End If

                            ' Printed message rewards
                            t = 1
                            MyCommon.QueryStr = "select PM.MessageID, PM.MessageTypeID, D.DeliverableID, D.DisallowEdit " &
                                                "from CPE_Deliverables D with (NoLock) " &
                                                "inner join PrintedMessages PM with (NoLock) on D.OutputID=PM.MessageID " &
                                                "where D.RewardOptionPhase=3 and D.RewardOptionID=" & ParentROID & " and D.DeliverableTypeID=4;"
                            rst = MyCommon.LRT_Select()
                            If (rst.Rows.Count > 0) Then
                                RewardsCount += rst.Rows.Count
                                AddOptionArray.Set(1, False)
                                Send("<tr class=""shadeddark"">")
                                Send("  <td colspan=""" & 3 & """>")
                                Send("    <h3>" & Copient.PhraseLib.Lookup("CPE-rew-pmsg.header", LanguageID) & "</h3>")
                                Send("  </td>")
                                If TierLevels = 1 Then
                                    Send("  <td></td>")
                                Else
                                    For t = 1 To TierLevels
                                        Send("  <td>")
                                        Send("    <b>" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</b>")
                                        Send("  </td>")
                                    Next
                                End If
                                t = 1
                                If (IsTemplate Or FromTemplate) Then
                                    Send("  <td></td>")
                                End If
                                Send("</tr>")
                                m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
                                For Each row In rst.Rows
                                    If Not IsTemplate Then
                                        RewardDisabled = IIf((Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False)) And Not IsOfferWaitingForApproval(OfferID)), "", " disabled=""disabled""")
                                    Else
                                        RewardDisabled = IIf((Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer), "", " disabled=""disabled""")
                                    End If
                                    Send("<tr class=""shaded"">")
                                    Send("<td>")
                                    If ((Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
                                        Sendb("  <input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & RewardDisabled & " value=""X"" ")
                                        Send("onClick=""if(confirm('" & Copient.PhraseLib.Lookup("reward.confirmdelete", LanguageID) & "')) LoadDocument('UEoffer-rew.aspx?mode=DeletePrintedMsg&OfferID=" & OfferID & "&MessageID=" & MyCommon.NZ(row.Item("MessageID"), 0) & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "')"" />")
                                    End If
                                    Send("</td>")
                                    Send("  <td><a href=""javascript:openPopup('UEoffer-rew-pmsg.aspx?OfferID=" & OfferID & "&RewardID=" & RewardID & "&MessageID=" & MyCommon.NZ(row.Item("MessageID"), 0) & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "')"">" & Copient.PhraseLib.Lookup("term.printedmessage", LanguageID) & "</a></td>")
                                    Send("  <td>" & GetMessageTypeName(MyCommon.NZ(row.Item("MessageTypeID"), 0)) & "</td>")

                                    ' Find the per-tier values and build up the details string:
                                    If (MyCommon.Fetch_SystemOption(124) = "1") Then
                                        MyCommon.QueryStr = "select PM.MessageID, PMT.TierLevel, PT.BodyText " &
                                                            "from PrintedMessages as PM with (NoLock) " &
                                                            "left join PrintedMessageTiers as PMT with (NoLock) on PM.MessageID=PMT.MessageID " &
                                                            "inner join PMTranslations as PT with (NoLock) on PT.PMTiersID=PMT.PKID and PT.LanguageID = @LanguageID " &
                                                            "where PM.MessageID = @MessageID"
                                        MyCommon.DBParameters.Add("@LanguageID", SqlDbType.Int).Value = DefaultLanguageID
                                    Else
                                        MyCommon.QueryStr = "select PM.MessageID, PMT.TierLevel, PMT.BodyText " &
                                                            "from PrintedMessages as PM with (NoLock) " &
                                                            "inner join PrintedMessageTiers as PMT with (NoLock) on PM.MessageID=PMT.MessageID " &
                                                            "where PM.MessageID = @MessageID"
                                    End If
                                    MyCommon.DBParameters.Add("@MessageID", SqlDbType.BigInt).Value = MyCommon.NZ(row.Item("MessageID"), 0)
                                    rst2 = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                                    If rst2.Rows.Count = 0 Then
                                        Send("  <td colspan=""" & TierLevels & """>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                                    Else
                                        While t <= TierLevels
                                            If t > rst2.Rows.Count Then
                                                Send("  <td>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                                            Else
                                                Details = New StringBuilder(200)
                                                Details.Append(ReplaceTags(MyCommon.NZ(rst2.Rows(t - 1).Item("BodyText"), "")))
                                                If (Details.ToString().Length > 80) Then
                                                    Details = Details.Remove(77, (Details.Length - 77))
                                                    Details.Append("...")
                                                End If
                                                'Overriding String Split
                                                Send("  <td>""" & Server.HtmlEncode(MyCommon.SplitNonSpacedString(Details.ToString, Details.ToString.Length)) & """</td>")
                                            End If
                                            t += 1
                                        End While
                                    End If
                                    t = 1
                                    If (IsTemplate) Then
                                        Send("  <td class=""templine"">")
                                        Send("    <input type=""checkbox"" id=""chkLocked2"" name=""chkLocked"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, " checked=""checked""", "") & " onclick=""javascript:updateLocked('lockPmsg" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "', this.checked);"" />")
                                        Send("    <input type=""hidden"" id=""rewPmsg" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""rew"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ />")
                                        Send("    <input type=""hidden"" id=""lockPmsg" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""locked"" value=""" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "1", "0") & """ />")
                                        Send("  </td>")
                                    ElseIf (FromTemplate) Then
                                        Send("  <td class=""templine"">" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "Yes", "No") & "</td>")
                                    End If
                                    Send("</tr>")
                                Next
                            End If

                            ' Cashier message rewards
                            t = 1
                            MyCommon.QueryStr = "select D.DeliverableID, D.DisallowEdit, CM.MessageID " &
                                                "from CPE_Deliverables D with (NoLock) " &
                                                "inner join CPE_CashierMessages CM with (NoLock) on D.OutputID=CM.MessageID " &
                                                "where D.RewardOptionID=" & ParentROID & " and DeliverableTypeID=9 and D.RewardOptionPhase=3;"
                            rst = MyCommon.LRT_Select()
                            If (rst.Rows.Count > 0) Then
                                RewardsCount += rst.Rows.Count
                                AddOptionArray.Set(2, False)
                                Send("<tr class=""shadeddark"">")
                                Send("  <td colspan=""" & 3 & """>")
                                Send("    <h3>" & Copient.PhraseLib.Lookup("CPE-rew-cmsg.header", LanguageID) & "</h3>")
                                Send("  </td>")
                                If TierLevels = 1 Then
                                    Send("  <td></td>")
                                Else
                                    For t = 1 To TierLevels
                                        Send("  <td>")
                                        Send("    <b>" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</b>")
                                        Send("  </td>")
                                    Next
                                End If
                                t = 1
                                If (IsTemplate Or FromTemplate) Then
                                    Send("  <td></td>")
                                End If
                                Send("</tr>")
                                m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
                                For Each row In rst.Rows
                                    If Not IsTemplate Then
                                        RewardDisabled = IIf((Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False)) And Not IsOfferWaitingForApproval(OfferID)), "", " disabled=""disabled""")
                                    Else
                                        RewardDisabled = IIf((Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer), "", " disabled=""disabled""")
                                    End If
                                    Send("<tr class=""shaded"">")
                                    Send("<td>")
                                    If ((Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
                                        Sendb(" <input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & RewardDisabled & " value=""X"" ")
                                        Send("onClick=""if(confirm('" & Copient.PhraseLib.Lookup("reward.confirmdelete", LanguageID) & "')) LoadDocument('UEoffer-rew.aspx?mode=DeleteCashierMsg&OfferID=" & OfferID & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "&MessageID=" & MyCommon.NZ(row.Item("MessageID"), 0) & "')"" />")
                                    End If
                                    Send("</td>")
                                    Send("  <td><a href=""javascript:openPopup('UEoffer-rew-cmsg.aspx?OfferID=" & OfferID & "&RewardID=" & RewardID & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "&MessageID=" & MyCommon.NZ(row.Item("MessageID"), 0) & "')"">" & Copient.PhraseLib.Lookup("term.cashiermessage", LanguageID) & "</a></td>")
                                    Send("  <td></td>")
                                    ' Find the per-tier values:
                                    MyCommon.QueryStr = "select CM.MessageID, CMT.Line1, CMT.Line2, CMT.Line3, CMT.Line4, CMT.Line5, CMT.Line6, CMT.Line7, CMT.Line8, CMT.Line9, CMT.Line10 " &
                                                        "from CPE_CashierMessages as CM with (NoLock) " &
                                                        "left join CPE_CashierMessageTiers as CMT with (NoLock) on CM.MessageID=CMT.MessageID " &
                                                        "where CM.MessageID=" & MyCommon.NZ(row.Item("MessageID"), 0) & ";"
                                    rst2 = MyCommon.LRT_Select
                                    If rst2.Rows.Count = 0 Then
                                        Send("  <td colspan=""" & TierLevels & """>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                                    Else
                                        Dim numrows As Integer = MyCommon.Fetch_UE_SystemOption(158)
                                        While t <= TierLevels
                                            If t > rst2.Rows.Count Then
                                                Send("  <td>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                                            Else
                                                Send("  <td>""")
                                                Dim lines As Integer = 0
                                                For lines = 1 To numrows
                                                    Send(MyCommon.NZ(rst2.Rows(t - 1).Item("Line" & lines), "") & "<br />")
                                                Next
                                                Send("""</td>")
                                            End If
                                            t += 1
                                        End While
                                    End If
                                    t = 1
                                    If (IsTemplate) Then
                                        Send("  <td class=""templine"">")
                                        Send("    <input type=""checkbox"" id=""chkLocked3"" name=""chkLocked"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, " checked=""checked""", "") & " onclick=""javascript:updateLocked('lockCmsg" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "', this.checked);"" />")
                                        Send("    <input type=""hidden"" id=""rewCmsg" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""rew"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ />")
                                        Send("    <input type=""hidden"" id=""lockCmsg" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""locked"" value=""" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "1", "0") & """ />")
                                        Send("  </td>")
                                    ElseIf (FromTemplate) Then
                                        Send("  <td class=""templine"">" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "Yes", "No") & "</td>")
                                    End If
                                    Send("</tr>")
                                Next
                            End If

                            ' Franking message rewards
                            t = 1
                            MyCommon.QueryStr = "select D.DeliverableID, D.DisallowEdit, FM.FrankID " &
                                                "from CPE_Deliverables D with (NoLock) " &
                                                "inner join CPE_FrankingMessages FM with (NoLock) on D.OutputID=FM.FrankID " &
                                                "where D.RewardOptionID=" & ParentROID & " and DeliverableTypeID=10 and D.RewardOptionPhase=3;"
                            rst = MyCommon.LRT_Select()
                            If (rst.Rows.Count > 0) Then
                                RewardsCount += rst.Rows.Count
                                AddOptionArray.Set(6, False)
                                Send("<tr class=""shadeddark"">")
                                Send("  <td colspan=""" & 3 & """>")
                                Send("    <h3>" & Copient.PhraseLib.Lookup("CPE-rew-fmsg.header", LanguageID) & "</h3>")
                                Send("  </td>")
                                If TierLevels = 1 Then
                                    Send("  <td></td>")
                                Else
                                    For t = 1 To TierLevels
                                        Send("  <td>")
                                        Send("    <b>" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</b>")
                                        Send("  </td>")
                                    Next
                                End If
                                t = 1
                                If (IsTemplate Or FromTemplate) Then
                                    Send("  <td></td>")
                                End If
                                Send("</tr>")
                                m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
                                For Each row In rst.Rows
                                    If Not IsTemplate Then
                                        RewardDisabled = IIf((Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False)) And Not IsOfferWaitingForApproval(OfferID)), "", " disabled=""disabled""")
                                    Else
                                        RewardDisabled = IIf((Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer), "", " disabled=""disabled""")
                                    End If
                                    Send("<tr class=""shaded"">")
                                    Send("<td>")
                                    If ((Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
                                        Sendb(" <input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & RewardDisabled & " value=""X"" ")
                                        Send("onClick=""if(confirm('" & Copient.PhraseLib.Lookup("reward.confirmdelete", LanguageID) & "')) LoadDocument('UEoffer-rew.aspx?mode=DeleteFrankingMsg&OfferID=" & OfferID & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "&FrankID=" & MyCommon.NZ(row.Item("FrankID"), 0) & "')"" />")
                                    End If
                                    Send("</td>")
                                    Send("  <td><a href=""javascript:openPopup('UEoffer-rew-franking.aspx?OfferID=" & OfferID & "&RewardID=" & RewardID & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "&FrankID=" & MyCommon.NZ(row.Item("FrankID"), 0) & "')"">" & Copient.PhraseLib.Lookup("term.frankingmessage", LanguageID) & "</a></td>")
                                    Send("  <td></td>")
                                    ' Find the per-tier values:
                                    MyCommon.QueryStr = "select FMT.OpenDrawer, FMT.ManagerOverride, FMT.FrankFlag, FMT.FrankingText, " &
                                                        "FMT.Line1, FMT.Line2, FMT.Beep, FMT.BeepDuration " &
                                                        "from CPE_FrankingMessageTiers as FMT with (NoLock) " &
                                                        "where FMT.FrankID=" & MyCommon.NZ(row.Item("FrankID"), 0) & ";"
                                    rst2 = MyCommon.LRT_Select
                                    If rst2.Rows.Count = 0 Then
                                        Send("  <td colspan=""" & TierLevels & """>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                                    Else
                                        While t <= TierLevels
                                            If t > rst2.Rows.Count Then
                                                Send("  <td>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                                            Else
                                                If MyCommon.NZ(rst2.Rows(t - 1).Item("FrankingText"), "") = "" Then
                                                    Send("  <td>")
                                                Else
                                                    Send("  <td>""" & MyCommon.SplitNonSpacedString(rst2.Rows(t - 1).Item("FrankingText"), 25) & """<br />")
                                                End If
                                                Sendb(IIf(MyCommon.NZ(rst2.Rows(t - 1).Item("OpenDrawer"), False) = True, Copient.PhraseLib.Lookup("term.opendrawer", LanguageID) & ",&nbsp;", Copient.PhraseLib.Lookup("term.closeddrawer", LanguageID) & ",&nbsp;"))
                                                Sendb(IIf(MyCommon.NZ(rst2.Rows(t - 1).Item("ManagerOverride"), False) = True, StrConv(Copient.PhraseLib.Lookup("term.override", LanguageID), VbStrConv.Lowercase) & " " & StrConv(Copient.PhraseLib.Lookup("term.required", LanguageID), VbStrConv.Lowercase) & ",&nbsp;", StrConv(Copient.PhraseLib.Lookup("term.override", LanguageID), VbStrConv.Lowercase) & " " & StrConv(Copient.PhraseLib.Lookup("term.notrequired", LanguageID), VbStrConv.Lowercase) & ",&nbsp;"))
                                                If (MyCommon.NZ(rst2.Rows(t - 1).Item("FrankFlag"), 0) = 0) Then
                                                    Sendb(Copient.PhraseLib.Lookup("term.posdata", LanguageID) & " ")
                                                    Sendb(StrConv(Copient.PhraseLib.Lookup("term.only", LanguageID), VbStrConv.Lowercase))
                                                ElseIf (MyCommon.NZ(rst2.Rows(t - 1).Item("FrankFlag"), 0) = 1) Then
                                                    Sendb(StrConv(Copient.PhraseLib.Lookup("term.frankingmessage", LanguageID), VbStrConv.Lowercase) & " ")
                                                    Sendb(StrConv(Copient.PhraseLib.Lookup("term.only", LanguageID), VbStrConv.Lowercase))
                                                ElseIf (MyCommon.NZ(rst2.Rows(t - 1).Item("FrankFlag"), 0) = 2) Then
                                                    Sendb(Copient.PhraseLib.Lookup("term.posdata", LanguageID) & " ")
                                                    Sendb(StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
                                                    Sendb(StrConv(Copient.PhraseLib.Lookup("term.frankingmessage", LanguageID), VbStrConv.Lowercase) & " ")
                                                End If
                                                Send("  </td>")
                                            End If
                                            t += 1
                                        End While
                                    End If
                                    t = 1
                                    If (IsTemplate) Then
                                        Send("  <td class=""templine"">")
                                        Send("    <input type=""checkbox"" id=""chkLocked4"" name=""chkLocked"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, " checked=""checked""", "") & " onclick=""javascript:updateLocked('lockFmsg" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "', this.checked);"" />")
                                        Send("    <input type=""hidden"" id=""rewFmsg" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""rew"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ />")
                                        Send("    <input type=""hidden"" id=""lockFmsg" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""locked"" value=""" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "1", "0") & """ />")
                                        Send("  </td>")
                                    ElseIf (FromTemplate) Then
                                        Send("  <td class=""templine"">" & Copient.PhraseLib.Lookup("term." & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "yes", "no"), LanguageID) & "</td>")
                                    End If
                                    Send("</tr>")
                                Next
                            End If

                            ' Points reward
                            t = 1
                            MyCommon.QueryStr = "select D.DeliverableID, D.DisallowEdit, DP.Quantity, DP.Multiplier, DP.MaxAdjustment, DP.ScorecardBold, DP.ProgramID, PP.ProgramName " &
                                                "from CPE_Deliverables as D with (NoLock) " &
                                                "inner join CPE_DeliverablePoints as DP with (NoLock) on DP.DeliverableID=D.DeliverableID " &
                                                "inner join PointsPrograms as PP with (NoLock) on PP.ProgramID=DP.ProgramID " &
                                                "where D.RewardOptionID=" & ParentROID & " and D.DeliverableTypeID=8 and D.Deleted=0 and DP.Deleted=0 order by ProgramName;"
                            rst = MyCommon.LRT_Select()
                            If (rst.Rows.Count > 0) Then
                                RewardsCount += rst.Rows.Count
                                AddOptionArray.Set(3, False)
                                Send("<tr class=""shadeddark"">")
                                Send("  <td colspan=""" & 3 & """>")
                                Send("    <h3>" & Copient.PhraseLib.Lookup("term.pointsrewards", LanguageID) & "</h3>")
                                Send("  </td>")
                                If TierLevels = 1 Then
                                    Send("  <td></td>")
                                Else
                                    For t = 1 To TierLevels
                                        Send("  <td>")
                                        Send("    <b>" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</b>")
                                        Send("  </td>")
                                    Next
                                End If
                                t = 1
                                If (IsTemplate Or FromTemplate) Then
                                    Send("  <td></td>")
                                End If
                                Send("</tr>")
                                m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
                                For Each row In rst.Rows
                                    If Not IsTemplate Then
                                        RewardDisabled = IIf((Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False)) And Not IsOfferWaitingForApproval(OfferID)), "", " disabled=""disabled""")
                                    Else
                                        RewardDisabled = IIf((Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer), "", " disabled=""disabled""")
                                    End If
                                    Send("<tr class=""shaded"">")
                                    Send("<td>")
                                    If ((Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
                                        Sendb("  <input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & RewardDisabled & " value=""X"" ")
                                        Send("onClick=""if(confirm('" & Copient.PhraseLib.Lookup("reward.confirmdelete", LanguageID) & "')) LoadDocument('UEoffer-rew.aspx?mode=DeletePoints&OfferID=" & OfferID & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "&ProgramID=" & MyCommon.NZ(row.Item("ProgramID"), 0) & "')"" />")
                                    End If
                                    Send("</td>")
                                    Send("  <td><a href=""javascript:openPopup('UEoffer-rew-point.aspx?OfferID=" & OfferID & "&RewardID=" & RewardID & "&ProgramID=" & MyCommon.NZ(row.Item("ProgramID"), 0) & "&quantity=" & MyCommon.NZ(row.Item("Quantity"), 0) & "&maxadjustment=" & MyCommon.NZ(row.Item("MaxAdjustment"), 0) & IIf(MyCommon.NZ(row.Item("ScorecardBold"), 0) = 0, "", "&ScorecardBold=on") & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "')"">" & Copient.PhraseLib.Lookup("term.points", LanguageID) & "</a></td>")
                                    Send("  <td></td>")
                                    ' Find the per-tier values:
                                    MyCommon.QueryStr = "select DP.PKID, DPT.TierLevel, DPT.Quantity, DPT.Multiplier " &
                                                        "from CPE_DeliverablePoints as DP with (NoLock) " &
                                                        "left join CPE_DeliverablePointTiers as DPT with (NoLock) on DP.PKID=DPT.DPPKID " &
                                                        "where DP.DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & ";"
                                    rst2 = MyCommon.LRT_Select
                                    If rst2.Rows.Count = 0 Then
                                        Send("  <td colspan=""" & TierLevels & """>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                                    Else
                                        While t <= TierLevels
                                            If t > rst2.Rows.Count Then
                                                Send("  <td>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                                            Else
                                                Send("  <td>" & MyCommon.NZ(rst2.Rows(t - 1).Item("Quantity"), 0) * MyCommon.NZ(rst2.Rows(t - 1).Item("Multiplier"), 1) & " " & "<a href=""../point-edit.aspx?ProgramGroupID=" & MyCommon.NZ(row.Item("ProgramID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("ProgramName"), ""), 25) & "</a></td>")
                                            End If
                                            t += 1
                                        End While
                                    End If
                                    t = 1
                                    If (IsTemplate) Then
                                        Send("  <td class=""templine"">")
                                        Send("    <input type=""checkbox"" id=""chkLocked5"" name=""chkLocked"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, " checked=""checked""", "") & " onclick=""javascript:updateLocked('lockPts" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "', this.checked);"" />")
                                        Send("    <input type=""hidden"" id=""rewPts" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""rew"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ />")
                                        Send("    <input type=""hidden"" id=""lockPts" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""locked"" value=""" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "1", "0") & """ />")
                                        Send("  </td>")
                                    ElseIf (FromTemplate) Then
                                        Send("  <td class=""templine"">" & Copient.PhraseLib.Lookup("term." & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "yes", "no"), LanguageID) & "</td>")
                                    End If
                                    Send("</tr>")
                                Next
                            End If

                            'Stored value rewards
                            t = 1
                            MyCommon.QueryStr = "select SVP.Name, SVP.SVProgramID, SVP.SVTypeID, SVT.ValuePrecision, D.DeliverableID as PKID, D.DisallowEdit, DSV.Quantity, DSV.Multiplier " &
                                                "from StoredValuePrograms as SVP with (NoLock) " &
                                                "inner join CPE_DeliverableStoredValue as DSV with (NoLock) on SVP.SVProgramID=DSV.SVProgramID and SVP.Deleted=0 " &
                                                "inner join CPE_Deliverables as D with (NoLock) on D.DeliverableID=DSV.DeliverableID and D.RewardOptionPhase=3 " &
                                                "left join SVTypes as SVT with (NoLock) on SVP.SVTypeID=SVT.SVTypeID " &
                                                "where D.RewardOptionID=" & ParentROID & " and D.Deleted=0 and DSV.Deleted=0 order by SVP.Name;"
                            rst = MyCommon.LRT_Select()
                            If (rst.Rows.Count > 0) Then
                                RewardsCount += rst.Rows.Count
                                AddOptionArray.Set(7, False)
                                Send("<tr class=""shadeddark"">")
                                Send("  <td colspan=""" & 3 & """>")
                                Send("    <h3>" & Copient.PhraseLib.Lookup("term.storedvaluerewards", LanguageID) & "</h3>")
                                Send("  </td>")
                                If TierLevels = 1 Then
                                    Send("  <td></td>")
                                Else
                                    For t = 1 To TierLevels
                                        Send("  <td>")
                                        Send("    <b>" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</b>")
                                        Send("  </td>")
                                    Next
                                End If
                                t = 1
                                If (IsTemplate Or FromTemplate) Then
                                    Send("  <td></td>")
                                End If
                                Send("</tr>")
                                m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
                                For Each row In rst.Rows
                                    If Not IsTemplate Then
                                        RewardDisabled = IIf((Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False)) And Not IsOfferWaitingForApproval(OfferID)), "", " disabled=""disabled""")
                                    Else
                                        RewardDisabled = IIf((Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer), "", " disabled=""disabled""")
                                    End If
                                    Send("<tr class=""shaded"">")
                                    Send("<td>")
                                    If ((Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
                                        Sendb("  <input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & RewardDisabled & " value=""X"" ")
                                        Send("onClick=""if(confirm('" & Copient.PhraseLib.Lookup("reward.confirmdelete", LanguageID) & "')) LoadDocument('UEoffer-rew.aspx?mode=DeleteStoredValue&OfferID=" & OfferID & "&DeliverableID=" & MyCommon.NZ(row.Item("PKID"), 0) & "&ProgramID=" & MyCommon.NZ(row.Item("SVProgramID"), 0) & "')"" />")
                                    End If
                                    Send("</td>")
                                    Send("  <td><a href=""javascript:openPopup('UEoffer-rew-sv.aspx?OfferID=" & OfferID & "&RewardID=" & RewardID & "&ProgramID=" & MyCommon.NZ(row.Item("SVProgramID"), 0) & "&quantity=" & MyCommon.NZ(row.Item("Quantity"), 0) & "&DeliverableID=" & MyCommon.NZ(row.Item("PKID"), 0) & "')"">" & Copient.PhraseLib.Lookup("term.storedvalue", LanguageID) & "</a></td>")
                                    Send("  <td></td>")
                                    ' Find the per-tier values:
                                    MyCommon.QueryStr = "select DSV.PKID, DSVT.TierLevel, DSVT.Quantity, DSVT.Multiplier, SVP.Value " &
                                                        "from CPE_DeliverableStoredValue as DSV with (NoLock) " &
                                                        "left join CPE_DeliverableStoredValueTiers as DSVT with (NoLock) on DSV.PKID=DSVT.DSVPKID " &
                                                        "left join StoredValuePrograms AS SVP with (NoLock) on SVP.SVProgramID=DSV.SVProgramID " &
                                                        "where DSV.DeliverableID=" & MyCommon.NZ(row.Item("PKID"), 0) & ";"
                                    rst2 = MyCommon.LRT_Select
                                    If rst2.Rows.Count = 0 Then
                                        Send("  <td colspan=""" & TierLevels & """>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                                    Else
                                        While t <= TierLevels
                                            If t > rst2.Rows.Count Then
                                                Send("  <td>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                                            Else
                                                Send("  <td>" & (MyCommon.NZ(rst2.Rows(t - 1).Item("Quantity"), "0") * MyCommon.NZ(rst2.Rows(t - 1).Item("Multiplier"), "1")) & " ")
                                                If (MyCommon.NZ(row.Item("SVTypeID"), 0) > 1) Then
                                                    Dim temp As Decimal
                                                    temp = Math.Round(MyCommon.NZ(rst2.Rows(t - 1).Item("Value"), 0) * (MyCommon.NZ(rst2.Rows(t - 1).Item("Quantity"), 0) * MyCommon.NZ(rst2.Rows(t - 1).Item("Multiplier"), 1)), MyCommon.NZ(row.Item("ValuePrecision"), 0))
                                                    Sendb("($" & temp.ToString(MyCommon.GetAdminUser.Culture) & ") ")
                                                End If
                                                Send(Copient.PhraseLib.Lookup("term.awardedinprogram", LanguageID) & " " & "<a href=""../SV-edit.aspx?ProgramGroupID=" & MyCommon.NZ(row.Item("SVProgramID"), 0) & """>" & MyCommon.SplitNonSpacedString(row.Item("Name"), 25) & "</a></td>")
                                            End If
                                            t += 1
                                        End While
                                    End If
                                    t = 1
                                    If (IsTemplate) Then
                                        Send("  <td class=""templine"">")
                                        Send("    <input type=""checkbox"" id=""chkLocked6"" name=""chkLocked"" value=""" & MyCommon.NZ(row.Item("PKID"), 0) & """" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, " checked=""checked""", "") & " onclick=""javascript:updateLocked('lockSV" & MyCommon.NZ(row.Item("PKID"), 0) & "', this.checked);"" />")
                                        Send("    <input type=""hidden"" id=""rewSV" & MyCommon.NZ(row.Item("PKID"), 0) & """ name=""rew"" value=""" & MyCommon.NZ(row.Item("PKID"), 0) & """ />")
                                        Send("    <input type=""hidden"" id=""lockSV" & MyCommon.NZ(row.Item("PKID"), 0) & """ name=""locked"" value=""" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "1", "0") & """ />")
                                        Send("  </td>")
                                    ElseIf (FromTemplate) Then
                                        Send("  <td class=""templine"">" & Copient.PhraseLib.Lookup("term." & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "yes", "no"), LanguageID) & "</td>")
                                    End If
                                    Send("</tr>")
                                Next
                            End If

                            'Monetary Stored Value rewards
                            t = 1
                            MyCommon.QueryStr = "select SVP.Name, SVP.SVProgramID, SVP.SVTypeID, D.DeliverableID as PKID, D.DisallowEdit " &
                                                "from StoredValuePrograms as SVP with (NoLock) " &
                                                "inner join CPE_DeliverableMonStoredValue as DSV with (NoLock) on SVP.SVProgramID=DSV.SVProgramID and SVP.Deleted=0 and SVP.SVTypeId = 2 " &
                                                "inner join CPE_Deliverables as D with (NoLock) on D.DeliverableID=DSV.DeliverableID and D.RewardOptionPhase=3 " &
                                                "left join SVTypes as SVT with (NoLock) on SVP.SVTypeID=SVT.SVTypeID " &
                                                "where D.RewardOptionID=" & ParentROID & " and D.Deleted=0 and DSV.Deleted=0 order by SVP.Name;"
                            rst = MyCommon.LRT_Select()
                            If (rst.Rows.Count > 0) Then
                                RewardsCount += rst.Rows.Count
                                AddOptionArray.Set(7, False)
                                Send("<tr class=""shadeddark"">")
                                Send("  <td colspan=""" & 3 & """>")
                                Send("    <h3>" & Copient.PhraseLib.Lookup("term.monstoredvaluerewards", LanguageID) & "</h3>")
                                Send("  </td>")
                                If TierLevels = 1 Then
                                    Send("  <td></td>")
                                Else
                                    For t = 1 To TierLevels
                                        Send("  <td>")
                                        Send("    <b>" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</b>")
                                        Send("  </td>")
                                    Next
                                End If
                                t = 1
                                If ((IsTemplate Or FromTemplate) And TierLevels = 4) Then
                                    Send("  <td></td>")
                                End If
                                Send("</tr>")
                                For Each row In rst.Rows
                                    If Not IsTemplate Then
                                        RewardDisabled = IIf((Logix.UserRoles.EditOffer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False)) And Not IsOfferWaitingForApproval(OfferID)), "", " disabled=""disabled""")
                                    Else
                                        RewardDisabled = IIf((Logix.UserRoles.EditTemplates), "", " disabled=""disabled""")
                                    End If
                                    Send("<tr class=""shaded"">")
                                    Sendb("  <td> ")
                                    If ((Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
                                        Send("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & RewardDisabled & " value=""X"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("reward.confirmdelete", LanguageID) & "')) LoadDocument('UEoffer-rew.aspx?mode=DeleteMonStoredValue&OfferID=" & OfferID & "&DeliverableID=" & MyCommon.NZ(row.Item("PKID"), 0) & "&ProgramID=" & MyCommon.NZ(row.Item("SVProgramID"), 0) & "')"" /></td>")
                                    End If
                                    Send("  <td><a href=""javascript:openPopup('UEoffer-rew-msv.aspx?OfferID=" & OfferID & "&RewardID=" & RewardID & "&ProgramID=" & MyCommon.NZ(row.Item("SVProgramID"), 0) & "&DeliverableID=" & MyCommon.NZ(row.Item("PKID"), 0) & "')"">" & Copient.PhraseLib.Lookup("term.monetarystoredvalue", LanguageID) & "</a></td>")
                                    Send("<td></td>")
                                    Send("<td colspan=""" & TierLevels & """> <a href=""../SV-edit.aspx?ProgramGroupID=" & MyCommon.NZ(row.Item("SVProgramID"), 0) & """>" & MyCommon.SplitNonSpacedString(row.Item("Name"), 25) & "</a></td>")
                                    t = 1
                                    If (IsTemplate) Then
                                        Send("  <td class=""templine"">")
                                        Send("    <input type=""checkbox"" id=""chkLocked6"" name=""chkLocked"" value=""" & MyCommon.NZ(row.Item("PKID"), 0) & """" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, " checked=""checked""", "") & " onclick=""javascript:updateLocked('lockSV" & MyCommon.NZ(row.Item("PKID"), 0) & "', this.checked);"" />")
                                        Send("    <input type=""hidden"" id=""rewSV" & MyCommon.NZ(row.Item("PKID"), 0) & """ name=""rew"" value=""" & MyCommon.NZ(row.Item("PKID"), 0) & """ />")
                                        Send("    <input type=""hidden"" id=""lockSV" & MyCommon.NZ(row.Item("PKID"), 0) & """ name=""locked"" value=""" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "1", "0") & """ />")
                                        Send("  </td>")
                                    ElseIf (FromTemplate) Then
                                        Send("  <td class=""templine"">" & Copient.PhraseLib.Lookup("term." & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "yes", "no"), LanguageID) & "</td>")
                                    End If
                                    Send("</tr>")
                                Next
                            End If

                            'Group membership rewards
                            t = 1
                            MyCommon.QueryStr = "select D.DeliverableID, D.DeliverableTypeID, D.RewardOptionID as ROID, D.DisallowEdit " &
                                                "from CPE_Deliverables as D with (NoLock) " &
                                                "where D.RewardOptionID=" & ParentROID & " and D.DeliverableTypeID in (5) and D.RewardOptionPhase=3 and D.Deleted=0 " &
                                                "order by D.DeliverableTypeID;"
                            rst = MyCommon.LRT_Select()
                            If (rst.Rows.Count > 0) Then
                                RewardsCount += rst.Rows.Count
                                AddOptionArray.Set(4, False)
                                Send("<tr class=""shadeddark"">")
                                Send("  <td colspan=""" & 3 & """>")
                                Send("    <h3>" & Copient.PhraseLib.Lookup("cpeoffer-rew-grpmbr", LanguageID) & "</h3>")
                                Send("  </td>")
                                If TierLevels = 1 Then
                                    Send("  <td></td>")
                                Else
                                    For t = 1 To TierLevels
                                        Send("  <td>")
                                        Send("    <b>" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</b>")
                                        Send("  </td>")
                                    Next
                                End If
                                t = 1
                                If (IsTemplate Or FromTemplate) Then
                                    Send("  <td></td>")
                                End If
                                Send("</tr>")
                                m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
                                For Each row In rst.Rows
                                    If Not IsTemplate Then
                                        RewardDisabled = IIf((Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False)) And Not IsOfferWaitingForApproval(OfferID)), "", " disabled=""disabled""")
                                    Else
                                        RewardDisabled = IIf((Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer), "", " disabled=""disabled""")
                                    End If
                                    UrlTokens = "?OfferID=" & OfferID & "&RewardID=" & MyCommon.NZ(row.Item("ROID"), 0) & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), -1) & "&action=" & MyCommon.NZ(row.Item("DeliverableTypeID"), 0)
                                    Send("<tr class=""shaded"">")
                                    Send("<td>")
                                    If ((Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
                                        Send("  <input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & RewardDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("reward.confirmdelete", LanguageID) & "')) LoadDocument('UEoffer-rew.aspx" & UrlTokens & "&mode=DeleteMembership')"" value=""X"" />")
                                    End If
                                    Send("</td>")
                                    Send("  <td><a href=""javascript:openPopup('UEoffer-rew-membership.aspx" & UrlTokens & "')"">" & Copient.PhraseLib.Lookup("term.membership", LanguageID) & "</a></td>")
                                    Send("  <td>")
                                    'Send(IIf(MyCommon.NZ(row.Item("DeliverableTypeID"), 0) = 6, "Remove", "Add"))
                                    Send("  </td>")
                                    ' Find the per-tier values:
                                    MyCommon.QueryStr = "select DCGT.TierLevel, DCGT.CustomerGroupID, CG.Name " &
                                                        "from CPE_DeliverableCustomerGroupTiers as DCGT with (NoLock) " &
                                                        "left join CustomerGroups as CG with (NoLock) on CG.CustomerGroupID=DCGT.CustomerGroupID " &
                                                        "where DCGT.DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & ";"
                                    rst2 = MyCommon.LRT_Select
                                    If rst2.Rows.Count = 0 Then
                                        Send("  <td colspan=""" & TierLevels & """>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                                    Else
                                        While t <= TierLevels
                                            If t > rst2.Rows.Count Then
                                                Send("  <td>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                                            Else
                                                Send("  <td><a href=""../cgroup-edit.aspx?CustomerGroupID=" & MyCommon.NZ(rst2.Rows(t - 1).Item("CustomerGroupID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(rst2.Rows(t - 1).Item("Name"), ""), 25) & "</a></td>")
                                            End If
                                            t += 1
                                        End While
                                    End If
                                    t = 1
                                    If (IsTemplate) Then
                                        Send("  <td class=""templine"">")
                                        Send("    <input type=""checkbox"" id=""chkLocked7"" name=""chkLocked"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, " checked=""checked""", "") & " onclick=""javascript:updateLocked('lockGM" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "', this.checked);"" />")
                                        Send("    <input type=""hidden"" id=""rewGM" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""rew"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ />")
                                        Send("    <input type=""hidden"" id=""lockGM" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""locked"" value=""" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "1", "0") & """ />")
                                        Send("  </td>")
                                    ElseIf (FromTemplate) Then
                                        Send("  <td class=""templine"">" & Copient.PhraseLib.Lookup("term." & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "yes", "no"), LanguageID) & "</td>")
                                    End If
                                    Send("</tr>")
                                Next
                            End If

                            ' Graphics rewards
                            t = 1
                            MyCommon.QueryStr = "select OSA.Name as GraphicName, OSA.OnScreenAdID as AdID, D.DeliverableID, D.ScreenCellID as CellID, D.DisallowEdit, " &
                                                "OSA.Width, OSA.Height, OSA.ImageType, SC.Name as CellName, SL.Name as LayoutName " &
                                                "from OnScreenAds as OSA with (NoLock) " &
                                                "inner join CPE_Deliverables as D with (NoLock) on OSA.OnScreenAdID=D.OutputID " &
                                                "inner join ScreenCells as SC with (NoLock) on D.ScreenCellID=SC.CellID " &
                                                "inner join ScreenLayouts as SL with (NoLock) on SL.LayoutID=SC.LayoutID " &
                                                "where D.RewardOptionID=" & ParentROID & " and OSA.Deleted=0 and D.DeliverableTypeID=1 and D.RewardOptionPhase=3;"
                            rst = MyCommon.LRT_Select
                            If (rst.Rows.Count > 0) Then
                                RewardsCount += rst.Rows.Count
                                Send("<tr class=""shadeddark"">")
                                Send("  <td colspan=""" & 3 & """>")
                                Send("    <h3>" & Copient.PhraseLib.Lookup("CPE-rew-graphics.header", LanguageID) & "</h3>")
                                Send("  </td>")
                                Send("  <td colspan=""" & TierLevels & """>")
                                Send("  </td>")
                                If (IsTemplate Or FromTemplate) Then
                                    Send("  <td></td>")
                                End If
                                Send("</tr>")
                                m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
                                For Each row In rst.Rows
                                    If Not IsTemplate Then
                                        RewardDisabled = IIf((Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False)) And Not IsOfferWaitingForApproval(OfferID)), "", " disabled=""disabled""")
                                    Else
                                        RewardDisabled = IIf((Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer), "", " disabled=""disabled""")
                                    End If
                                    DeleteGraphicURL = "UEoffer-rew.aspx?mode=DeleteGraphic&OfferID=" & OfferID & "&deliverableid=" & MyCommon.NZ(row.Item("DeliverableID"), "")
                                    Send("<tr class=""shaded"">")
                                    Send("<td>")
                                    If ((Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
                                        Send("  <input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & RewardDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("reward.confirmdelete", LanguageID) & "')) LoadDocument('" & DeleteGraphicURL & " ');"" value=""X"" />")
                                    End If
                                    Send("</td>")
                                    Send("  <td><a href=""javascript:openPopup('UEoffer-rew-graphic.aspx?OfferID=" & OfferID & "&ad=" & MyCommon.NZ(row.Item("AdId"), "") & "&cellselect=" & MyCommon.NZ(row.Item("CellID"), "") & "&imagetype=" & MyCommon.NZ(row.Item("ImageType"), "") & "&DeliverableID=" & MyCommon.Extract_Val(row.Item("DeliverableID")) & "&preview=1')"">" & Copient.PhraseLib.Lookup("term.graphic", LanguageID) & "</a></td>")
                                    Send("  <td></td>")
                                    Sendb("  <td colspan=""" & TierLevels & """><a href=""../graphic-edit.aspx?OnScreenAdId=" & MyCommon.NZ(row.Item("AdId"), "") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("GraphicName"), ""), 25) & "</a>")
                                    Sendb("&nbsp;(" & MyCommon.NZ(row.Item("Width"), "") & " x " & MyCommon.NZ(row.Item("Height"), ""))
                                    If MyCommon.NZ(row.Item("ImageType"), "") = 1 Then
                                        Sendb("&nbsp;" & Copient.PhraseLib.Lookup("term.jpeg", LanguageID))
                                    ElseIf MyCommon.NZ(row.Item("ImageType"), "") = 2 Then
                                        Sendb("&nbsp;" & Copient.PhraseLib.Lookup("term.gif", LanguageID))
                                    End If
                                    Send(")</td>")
                                    If (IsTemplate) Then
                                        Send("  <td class=""templine"">")
                                        Send("    <input type=""checkbox"" id=""chkLocked8"" name=""chkLocked"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, " checked=""checked""", "") & " onclick=""javascript:updateLocked('lockGraphics" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "', this.checked);"" />")
                                        Send("    <input type=""hidden"" id=""rewGrapics" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""rew"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ />")
                                        Send("    <input type=""hidden"" id=""lockGraphics" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""locked"" value=""" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "1", "0") & """ />")
                                        Send("  </td>")
                                    ElseIf (FromTemplate) Then
                                        Send("  <td class=""templine"">" & Copient.PhraseLib.Lookup("term." & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "yes", "no"), LanguageID) & "</td>")
                                    End If
                                    Send("</tr>")
                                Next
                            End If

                            ' Touchpoint rewards
                            t = 1
                            MyCommon.QueryStr = "select RO.Name, RO.RewardOptionID, TA.OnScreenAdID as ParentAdID, D.DisallowEdit " &
                                                "from CPE_RewardOptions RO with (NoLock) " &
                                                "inner join CPE_DeliverableROIDs DR with (NoLock) on RO.RewardOptionID=DR.RewardOptionID " &
                                                "inner join CPE_Deliverables D with (NoLock) on D.DeliverableID=DR.DeliverableID " &
                                                "inner join TouchAreas TA with (NoLock) on DR.AreaID=TA.AreaID " &
                                                "where RO.Deleted=0 and DR.Deleted=0 and TA.Deleted=0 and RO.IncentiveID=" & OfferID & " and RO.TouchResponse=1 and D.RewardOptionPhase=3 order by RO.RewardOptionID;"
                            rst = MyCommon.LRT_Select
                            If (rst.Rows.Count > 0) Then
                                RewardsCount += rst.Rows.Count
                                Send("<tr class=""shadeddark"">")
                                Send("  <td colspan=""" & 3 & """>")
                                Send("    <h3>" & Copient.PhraseLib.Lookup("term.touchpointrewards", LanguageID) & "</h3>")
                                Send("  </td>")
                                Send("  <td colspan=""" & TierLevels & """>")
                                Send("  </td>")
                                If (IsTemplate Or FromTemplate) Then
                                    Send("  <td></td>")
                                End If
                                Send("</tr>")
                                index = 0
                                For Each row In rst.Rows
                                    ROID = MyCommon.NZ(row.Item("RewardOptionID"), 0)
                                    'AddTouchPtURL = "UEoffer-rew-deliverables.aspx?OfferID=" & OfferID & "&incentiveid=" & OfferID & "&roid=" & ROID & "&phase=3"
                                    Send("<tr class=""shadedmid"">")
                                    Send("  <td></td>")
                                    Send("  <td colspan=""2"">")
                                    Send("    <a href=""../graphic-edit.aspx?OnScreenAdId=" & MyCommon.NZ(row.Item("ParentAdID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)), 25) & "</a>")
                                    Send("  </td>")
                                    Send("  <td colspan=""" & TierLevels & """>")
                                    Send("    <label for=""newrewtouchpt" & index & """>" & Copient.PhraseLib.Lookup("CPE-rew.addtouchpoint", LanguageID) & "</label><br />")
                                    Send("    <select name=""newrewtouchpt" & index & """ id=""newrewtouchpt" & index & """>")
                                    Send_TPRewardOptions(OfferID, ROID)
                                    Send("    </select>")
                                    If ((Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable) And Not IsOfferWaitingForApproval(OfferID)) Then
                                        Send("    <input type=""button"" class=""regular"" id=""addTouchpoint"" name=""addTouchpoint"" value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """ " & DeleteBtnDisabled & " onclick=""javascript:openTouchptReward(" & index & ", " & ROID & ");"" />")
                                    End If
                                    Send("  </td>")
                                    If (IsTemplate Or FromTemplate) Then
                                        Send("  <td></td>")
                                    End If
                                    Send("</tr>")
                                    m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
                                    If Not IsTemplate Then
                                        SetEditableByUser(Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer)
                                    Else
                                        SetEditableByUser(Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer)
                                    End If

                                    Send_TouchpointRewards(OfferID, ROID, 3, TierLevels, IsTemplate, FromTemplate)
                                    index = index + 1
                                Next
                            End If

                            ' Pass-thru reward(s)
                            MyCommon.QueryStr = "select PassThruRewardID, Name, PhraseID from PassThruRewards with (NoLock) order by Name;"
                            rstPT = MyCommon.LRT_Select
                            If rstPT.Rows.Count > 0 Then
                                i = 1
                                For Each rowPT In rstPT.Rows
                                    t = 1
                                    MyCommon.QueryStr = "select D.DeliverableID, D.DisallowEdit, DPT.PassThruRewardID, PTR.Name, PTR.PhraseID, PTR.LSInterfaceID, PTR.ActionTypeID " &
                                                        "from CPE_Deliverables as D with (NoLock) " &
                                                        "inner join PassThrus as DPT with (NoLock) on DPT.PKID=D.OutputID " &
                                                        "inner join PassThruRewards as PTR with (NoLock) on PTR.PassThruRewardID=DPT.PassThruRewardID " &
                                                        "where D.RewardOptionID=" & ParentROID & " and DPT.PassThruRewardID=" & MyCommon.NZ(rowPT.Item("PassThruRewardID"), 0) & " and D.Deleted=0 and DeliverableTypeID=12 and RewardOptionPhase=3 " &
                                                        "order by Name;"
                                    rst = MyCommon.LRT_Select()
                                    If (rst.Rows.Count > 0) Then
                                        RewardsCount += rst.Rows.Count
                                        AddOptionArray.Set(8 + (i - 1), False)
                                        Send("<tr class=""shadeddark"">")
                                        Send("  <td colspan=""" & 3 & """>")
                                        If IsDBNull(rowPT.Item("PhraseID")) Then
                                            Send("    <h3>" & MyCommon.NZ(rowPT.Item("Name"), Copient.PhraseLib.Lookup("term.passthrureward", LanguageID)) & "</h3>")
                                        Else
                                            Send("    <h3>" & Copient.PhraseLib.Lookup(rowPT.Item("PhraseID"), LanguageID) & "</h3>")
                                        End If
                                        Send("  </td>")
                                        If TierLevels = 1 Then
                                            Send("  <td></td>")
                                        Else
                                            For t = 1 To TierLevels
                                                Send("  <td>")
                                                Send("    <b>" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</b>")
                                                Send("  </td>")
                                            Next
                                        End If
                                        If (IsTemplate Or FromTemplate) Then
                                            Send("  <td></td>")
                                        End If
                                        Send("</tr>")
                                        t = 1
                                        m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
                                        For Each row In rst.Rows
                                            If Not IsTemplate Then
                                                RewardDisabled = IIf((Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False)) And Not IsOfferWaitingForApproval(OfferID)), "", " disabled=""disabled""")
                                            Else
                                                RewardDisabled = IIf((Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer), "", " disabled=""disabled""")
                                            End If


                                            Send("<tr class=""shaded"">")
                                            Send("<td>")
                                            If ((Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
                                                Sendb("  <input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & RewardDisabled & " value=""X"" ")
                                                Send("onClick=""if(confirm('" & Copient.PhraseLib.Lookup("reward.confirmdelete", LanguageID) & "')) LoadDocument('UEoffer-rew.aspx?mode=DeletePassThru&OfferID=" & OfferID & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "&PassThruRewardID=" & MyCommon.NZ(row.Item("PassThruRewardID"), 0) & "')"" />")
                                            End If
                                            Send("</td>")
                                            If MyCommon.NZ(row.Item("PassThruRewardID"), 0) = 0 Then
                                                Sendb("  <td><a href=""javascript:openPopup('/logix/Offer-Rew-XMLPassThru.aspx?OfferID=" & OfferID & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "&PassThruRewardID=" & MyCommon.NZ(row.Item("PassThruRewardID"), "") & "')"">")
                                            Else
                                                Sendb("  <td><a href=""javascript:openPopup('UEoffer-rew-passthru.aspx?OfferID=" & OfferID & "&RewardID=" & RewardID & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "&PassThruRewardID=" & MyCommon.NZ(row.Item("PassThruRewardID"), "") & "')"">")
                                            End If

                                            If IsDBNull(rowPT.Item("PhraseID")) Then

                                                Send(MyCommon.NZ(rowPT.Item("Name"), Copient.PhraseLib.Lookup("term.passthrureward", LanguageID)))
                                            Else
                                                Send(Copient.PhraseLib.Lookup(rowPT.Item("PhraseID"), LanguageID))
                                            End If
                                            Send("</a></td>")
                                            Send("  <td></td>")
                                            ' Find the per-tier values:
                                            MyCommon.QueryStr = "select DPT.PKID, DPTT.TierLevel, (SUBSTRING(DPTT.Data, 0, 100) + '...') as Data " &
                                                                "from PassThrus as DPT with (NoLock) " &
                                                                "inner join PassThruTiers as DPTT with (NoLock) on DPT.PKID=DPTT.PTPKID " &
                                                                "where DPT.DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & " " &
                                                                "and LanguageID =" & DefaultLanguageID
                                            rst2 = MyCommon.LRT_Select
                                            If rst2.Rows.Count = 0 Then
                                                Send("  <td colspan=""" & TierLevels & """>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                                            Else
                                                While t <= TierLevels
                                                    If t > rst2.Rows.Count Then
                                                        Send("  <td>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                                                    Else
                                                        Send("  <td style=""word-break: break-all;""><p style=""font-size:11px;white-space: pre-wrap;"">" & MyCommon.NZ(rst2.Rows(t - 1).Item("Data"), "").ToString.Replace("<", "&lt;") & "</p></td>")
                                                    End If
                                                    t += 1
                                                End While
                                            End If
                                            t = 1
                                            If (IsTemplate) Then
                                                Send("  <td class=""templine"">")
                                                Send("    <input type=""checkbox"" id=""chkLocked8"" name=""chkLocked"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, " checked=""checked""", "") & " onclick=""javascript:updateLocked('lockPassthru" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "', this.checked);"" />")
                                                Send("    <input type=""hidden"" id=""rewPassthru" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""rew"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ />")
                                                Send("    <input type=""hidden"" id=""lockPassthru" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""locked"" value=""" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "1", "0") & """ />")
                                                Send("  </td>")
                                            ElseIf (FromTemplate) Then
                                                Send("  <td class=""templine"">" & Copient.PhraseLib.Lookup("term." & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "yes", "no"), LanguageID) & "</td>")
                                            End If
                                            Send("</tr>")
                                        Next
                                    End If
                                    i += 1
                                Next
                            End If


                            ' GiftCard reward(s)
                            t = 1
                            MyCommon.QueryStr = "SELECT G.ID, G.LASTUPDATE, D.DELIVERABLEID, D.DISALLOWEDIT  FROM CPE_DELIVERABLES D INNER JOIN GIFTCARD G ON D.OutputID=G.ID WHERE G.REWARDOPTIONID=" & RewardID & " and D.DeliverableTypeID=13;"
                            rst = MyCommon.LRT_Select()
                            If (rst.Rows.Count > 0) Then
                                RewardsCount += rst.Rows.Count
                                AddOptionArray.Set(0, False) 'Changing this to True will allow multiple giftcard rewards. False will allow only one Giftcard
                                Send("<tr class=""shadeddark"">")
                                Send("  <td colspan=""" & 3 & """>")
                                Send("    <h3>" & Copient.PhraseLib.Lookup("term.GiftCard", LanguageID) & "</h3>")
                                Send("  </td>")
                                If TierLevels = 1 Then
                                    Send("  <td></td>")
                                Else
                                    For t = 1 To TierLevels
                                        Send("  <td>")
                                        Send("    <b>" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</b>")
                                        Send("  </td>")
                                    Next
                                End If
                                t = 1
                                If (IsTemplate Or FromTemplate) Then
                                    Send("  <td></td>")
                                End If
                                Send("</tr>")
                                m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
                                For Each row In rst.Rows
                                    If Not IsTemplate Then
                                        RewardDisabled = IIf((Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False)) And Not IsOfferWaitingForApproval(OfferID)), "", " disabled=""disabled""")
                                    Else
                                        RewardDisabled = IIf((Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer), "", " disabled=""disabled""")
                                    End If
                                    Send("<tr class=""shaded"">")
                                    Send("<td>")
                                    If ((Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
                                        Sendb("  <input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & RewardDisabled & " value=""X"" ")
                                        Send("onClick=""if(confirm('" & Copient.PhraseLib.Lookup("reward.confirmdelete", LanguageID) & "')) LoadDocument('UEoffer-rew.aspx?mode=DeleteGiftCard&OfferID=" & OfferID & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "&GiftCardID=" & MyCommon.NZ(row.Item("ID"), 0) & "')"" />")
                                    End If
                                    Send("</td>")

                                    Dim GCRPercentOffAllowed As Boolean = True
                                    Dim ProductCondition As Boolean = False
                                    Dim UnitTypeId As Int32
                                    GetGcrQueryFlags(MyCommon, ParentROID, ProductCondition, GCRPercentOffAllowed, OfferID, UnitTypeId)
                                    Dim RestrictProrationTypeToAllConditional As Boolean = False
                                    If GCRPercentOffAllowed AndAlso UnitTypeId = CPEUnitTypes.Dollars Then
                                        RestrictProrationTypeToAllConditional = True
                                    End If
                                    Send("  <td><a href=""javascript:openPopup('UEoffer-rew-giftCard.aspx?OfferID=" & OfferID & "&Phase=3&RewardID=" & RewardID & "&productCondition=" & ProductCondition & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "&GiftCardID=" & MyCommon.NZ(row.Item("ID"), "") & "&PercentOffAllowed=" & GCRPercentOffAllowed & "&RestrictProrationTypeToAllConditional=" & RestrictProrationTypeToAllConditional & "')"">" & Copient.PhraseLib.Lookup("term.GiftCard", LanguageID) & "</a></td>")

                                    If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                                        MyCommon.Open_LogixRT()
                                    End If
                                    MyCommon.QueryStr = "select distinct at.PhraseId from GiftCardTier gct join CPE_AmountTypes at on (gct.AmountTypeId=at.AmountTypeId) where gct.GiftCardID=" & row("Id")
                                    rst2 = MyCommon.LRT_Select
                                    If rst2.Rows.Count = 1 Then
                                        Send("  <td>" & Copient.PhraseLib.Lookup(rst2.Rows(0)("PhraseId"), LanguageID) & "</td>")
                                    Else
                                        Send("  <td></td>")
                                    End If
                                    ' Find the per-tier details, and build up the details string:
                                    MyCommon.QueryStr = "select gt.Id,gt.CardIdentifier,gt.Name, gt.TierLevel,gt.AmountTypeID ,gt.Amount,gt.BuyDescription,gt.ChargebackDeptID,pt.PhraseID as ProrationTypePhrase, at.PhraseID as AmountTypePhrase from GiftCardTier gt with (nolock) " &
                                                        "inner join Relation_RewardProration RRP on RRP.ProrationTypeID=gt.ProrationTypeID " &
                                                        "inner join UE_ProrationTypes PT with (nolock) on PT.ProrationTypeID=rrp.ProrationTypeID " &
                                                        "inner join CPE_AmountTypes AT with (nolock) on at.AmountTypeID=gt.AmountTypeID " &
                                                        "where gt.GiftCardID=" & MyCommon.NZ(row.Item("id"), 0) & " Order by gt.TierLevel;"
                                    rst2 = MyCommon.LRT_Select
                                    If rst2.Rows.Count = 0 Then
                                        Send("  <td colspan=""" & TierLevels & """>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                                    Else
                                        While t <= TierLevels
                                            If t > rst2.Rows.Count Then
                                                Send("  <td>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                                            Else
                                                Details = New StringBuilder(200)
                                                Details.Append(Copient.PhraseLib.Lookup(MyCommon.NZ(rst2.Rows(t - 1).Item("AmountTypePhrase"), 0), LanguageID) & "&nbsp;")
                                                AmountTypeID = (MyCommon.NZ(rst2.Rows(t - 1).Item("AmountTypeID"), 0))
                                                Select Case AmountTypeID
                                                    Case 1
                                                        Details.Append(Localizer.FormatCurrency_ForOffer(MyCommon.NZ(rst2.Rows(t - 1).Item("Amount"), 0), RewardID).ToString(MyCommon.GetAdminUser.Culture) & " " & StrConv(Copient.PhraseLib.Lookup("term.on", LanguageID), VbStrConv.Lowercase) & "&nbsp;")
                                                    Case 3
                                                        Details.Append(Math.Round(CDec(MyCommon.NZ(rst2.Rows(t - 1).Item("Amount"), 0)), 2).ToString(MyCommon.GetAdminUser.Culture) & "% " & " " & StrConv(Copient.PhraseLib.Lookup("term.on", LanguageID), VbStrConv.Lowercase) & "&nbsp;")
                                                    Case Else
                                                        Details.Append(MyCommon.NZ(rst2.Rows(t - 1).Item("Amount"), "") & "&nbsp;")
                                                End Select
                                                Details.Append(Copient.PhraseLib.Lookup(MyCommon.NZ(rst2.Rows(t - 1).Item("ProrationTypePhrase"), 0), LanguageID) & "&nbsp;")
                                                Send("  <td>" & Details.ToString & "</td>")
                                            End If
                                            t += 1
                                        End While
                                        t = 1
                                    End If
                                    If (IsTemplate) Then
                                        Send("  <td class=""templine"">")
                                        Send("    <input type=""checkbox"" id=""chkLocked1"" name=""chkLocked"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, " checked=""checked""", "") & " onclick=""javascript:updateLocked('lockDisc" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "', this.checked);"" />")
                                        Send("    <input type=""hidden"" id=""rewDisc" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""rew"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ />")
                                        Send("    <input type=""hidden"" id=""lockDisc" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""locked"" value=""" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "1", "0") & """ />")
                                        Send("  </td>")
                                    ElseIf (FromTemplate) Then
                                        Send("  <td class=""templine"">" & Copient.PhraseLib.Lookup("term." & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "yes", "no"), LanguageID) & "</td>")
                                    End If
                                    Send("</tr>")
                                Next
                            End If
                            ''Preference as reward -- Start
                            If EPMInstalled = True Then
                                ' Preference reward                        
                                t = 1
                                Dim objResult As AMSResult(Of List(Of PreferenceReward)) = objPreferenceService.GetAllPreferenceRewardByROID(ParentROID)
                                Dim lstPreferenceReward As List(Of PreferenceReward) = Nothing
                                If objResult.ResultType = AMSResultType.Success AndAlso objResult.Result.Count > 0 Then
                                    lstPreferenceReward = objResult.Result
                                    RewardsCount += lstPreferenceReward.Count
                                    AddOptionArray.Set(3, False)
                                    Send("<tr class=""shadeddark"">")
                                    Send("  <td colspan=""" & 3 & """>")
                                    Send("    <h3>" & Copient.PhraseLib.Lookup("term.preferencerewards", LanguageID) & "</h3>")
                                    Send("  </td>")
                                    If TierLevels = 1 Then
                                        Send("  <td></td>")
                                    Else
                                        For t = 1 To TierLevels
                                            Send("  <td>")
                                            Send("    <b>" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</b>")
                                            Send("  </td>")
                                        Next
                                    End If
                                    t = 1
                                    If (IsTemplate Or FromTemplate) Then
                                        Send("  <td></td>")
                                    End If
                                    Send("</tr>")

                                    For Each item As PreferenceReward In lstPreferenceReward
                                        Dim preference As Preference = m_PreferenceService.GetPreferenceByID(item.PreferenceID, LanguageID).Result
                                        If preference Is Nothing Then
                                            Continue For
                                        End If
                                        preference.PreferenceValues = m_PreferenceService.GetPreferenceItemsbyPreferenceID(preference.DataTypeID, preference.PreferenceID, LanguageID).Result
                                        Dim PrefPageName As String = IIf(preference.UserCreated, "prefscustom-edit.aspx", "prefsstd-edit.aspx")
                                        Dim RootURI As String = IntegrationVals.HTTP_RootURI
                                        If RootURI IsNot Nothing AndAlso RootURI.Length > 0 AndAlso Right(RootURI, 1) <> "/" Then
                                            RootURI &= "/"
                                        End If
                                        If Not IsTemplate Then
                                            RewardDisabled = IIf((Not (FromTemplate And MyCommon.NZ(item.DisallowEdit, False)) And Not IsOfferWaitingForApproval(OfferID)), "", " disabled=""disabled""")
                                        Else
                                            RewardDisabled = IIf((Logix.UserRoles.EditTemplates), "", " disabled=""disabled""")
                                        End If
                                        Send("<tr class=""shaded"">")
                                        Send("<td>")
                                        If ((Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
                                            Sendb("  <input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & RewardDisabled & " value=""X"" ")
                                            Send("onClick=""if(confirm('" & Copient.PhraseLib.Lookup("reward.confirmdelete", LanguageID) & "')) LoadDocument('UEoffer-rew.aspx?mode=DeletePreferenceReward&OfferID=" & OfferID & "&DeliverableID=" & MyCommon.NZ(item.RewardID, 0) & "&PreferenceRewardID=" & MyCommon.NZ(item.PreferenceRewardID, 0) & "')"" />")
                                        End If
                                        Send("</td>")
                                        Send("  <td><a href=""javascript:openPopup('UEoffer-rew-pref.aspx?OfferID=" & OfferID & "&RewardID=" & MyCommon.NZ(item.RewardOptionId, 0) & "&PreferenceRewardID=" & MyCommon.NZ(item.PreferenceRewardID, 0) & "&DeliverableID=" & MyCommon.NZ(item.RewardID, 0) & "')"">" & Copient.PhraseLib.Lookup("term.preference", LanguageID) & "</a></td>")
                                        Send("  <td></td>")
                                        If item.PreferenceRewardTiers.Count = 0 Then
                                            Send("  <td colspan=""" & TierLevels & """>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                                        Else
                                            While t <= TierLevels
                                                If t > item.PreferenceRewardTiers.Count Then
                                                    Send("  <td>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                                                Else
                                                    Dim tier As PreferenceRewardTier = (From p In item.PreferenceRewardTiers Where p.TierLevel = t).ToList.First
                                                    Dim strTierValues As String = String.Join(",", preference.PreferenceValues.Where(Function(p) tier.PreferenceRewardTierValues.Select(Function(p2) p2.PreferenceValue).Contains(p.Value)).Select(Function(p3) p3.PhraseText))
                                                    Sendb("<td>")
                                                    Sendb(Copient.PhraseLib.Lookup("term.preference", LanguageID) & "  <a href=""../authtransfer.aspx?SendToURI=" & RootURI & "UI/" & PrefPageName & "?prefid=" & preference.PreferenceID & """>" & preference.PhraseText & "</a> " & Copient.PhraseLib.Lookup("term.rewarding", LanguageID).ToLower() & " " & Copient.PhraseLib.Lookup("term.values", LanguageID) & ": ")
                                                    Send(strTierValues.Trim(",").ToString & "</td>")
                                                End If
                                                t += 1
                                            End While
                                            t = 1
                                        End If
                                        If (IsTemplate) Then
                                            Send("  <td class=""templine"">")
                                            Send("    <input type=""checkbox"" id=""chkPref"" name=""chkLocked"" value=""" & MyCommon.NZ(item.RewardID, 0) & """" & IIf(MyCommon.NZ(item.DisallowEdit, False) = True, " checked=""checked""", "") & " onclick=""javascript:updateLocked('lockPref" & MyCommon.NZ(item.RewardID, 0) & "', this.checked);"" />")
                                            Send("    <input type=""hidden"" id=""rewPref" & MyCommon.NZ(item.RewardID, 0) & """ name=""rew"" value=""" & MyCommon.NZ(item.RewardID, 0) & """ />")
                                            Send("    <input type=""hidden"" id=""lockPref" & MyCommon.NZ(item.RewardID, 0) & """ name=""locked"" value=""" & IIf(MyCommon.NZ(item.DisallowEdit, False) = True, "1", "0") & """ />")
                                            Send("  </td>")
                                        ElseIf (FromTemplate) Then
                                            Send("  <td class=""templine"">" & Copient.PhraseLib.Lookup("term." & IIf(MyCommon.NZ(item.DisallowEdit, False) = True, "yes", "no"), LanguageID) & "</td>")
                                        End If
                                        Send("</tr>")
                                    Next
                                End If
                            End If
                            'Trackable Coupon Reward
                            t = 1
                            Dim objCouponReward As List(Of CouponReward) = objCouponService.GetAllCouponRewardbyROID(ParentROID)
                            If (objCouponReward.Count > 0) Then
                                RewardsCount += objCouponReward.Count
                                AddOptionArray.Set(3, False)
                                Send("<tr class=""shadeddark"">")
                                Send("  <td colspan=""" & 3 & """>")
                                Send("    <h3>" & Copient.PhraseLib.Lookup("term.trackablecoupon", LanguageID) + " " + Copient.PhraseLib.Lookup("term.rewards", LanguageID).ToLower() & "</h3>")
                                Send("  </td>")
                                If TierLevels = 1 Then
                                    Send("  <td></td>")
                                Else
                                    For t = 1 To TierLevels
                                        Send("  <td>")
                                        Send("    <b>" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</b>")
                                        Send("  </td>")
                                    Next
                                End If
                                t = 1
                                If (IsTemplate Or FromTemplate) Then
                                    Send("  <td></td>")
                                End If
                                Send("</tr>")
                                m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
                                For Each couponObj In objCouponReward
                                    If Not IsTemplate Then
                                        RewardDisabled = IIf((Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And couponObj.DisallowEdit) And Not IsOfferWaitingForApproval(OfferID)), "", " disabled=""disabled""")
                                    Else
                                        RewardDisabled = IIf((Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer), "", " disabled=""disabled""")
                                    End If
                                    Send("<tr class=""shaded"">")
                                    Send("<td>")
                                    If ((Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
                                        Sendb("  <input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & RewardDisabled & " value=""X"" ")
                                        Send("onClick=""if(confirm('" & Copient.PhraseLib.Lookup("reward.confirmdelete", LanguageID) & "')) LoadDocument('UEoffer-rew.aspx?mode=DeleteTrackableCouponReward&OfferID=" & OfferID & "&DeliverableID=" & couponObj.DeliverableID & "')"" />")
                                    End If
                                    Send("</td>")
                                    Send("  <td><a href=""javascript:openPopup('UEoffer-rew-tc.aspx?OfferID=" & OfferID & "&RewardID=" & RewardID & "&TCDeliverableID=" & couponObj.CouponRewardID & "&DeliverableID=" & couponObj.DeliverableID & "')"">" & Copient.PhraseLib.Lookup("term.trackablecoupon", LanguageID) & "</a></td>")
                                    Send("  <td></td>")
                                    ' Find the per-tier values:
                                    ' For Each tierObj In couponObj.CouponTiers
                                    t = 1
                                    While t <= TierLevels
                                        If couponObj.CouponTiers IsNot Nothing Then
                                            If couponObj.CouponTiers.Count <= (t - 1) Then
                                                Send("  <td >" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                                            Else
                                                MyCommon.QueryStr = "select Name  " &
                                                     "from TrackableCouponProgram with (NoLock) " &
                                                     "where ProgramID=" & couponObj.CouponTiers(t - 1).ProgramID & ";"
                                                rst2 = MyCommon.LRT_Select
                                                If rst2.Rows.Count = 0 Then
                                                    Send("  <td >" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                                                    'ElseIf t > rst2.Rows.Count Then
                                                    '        Send("  <td>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                                                Else
                                                    Send("  <td style='word-wrap:break-word'><a href=""../tcp-edit.aspx?tcprogramid=" & couponObj.CouponTiers(t - 1).ProgramID & """>" & (MyCommon.NZ(rst2.Rows(0).Item("Name"), "")) & "</a></td>")
                                                End If
                                            End If
                                        End If


                                        t += 1
                                    End While

                                    '  End If
                                    'Next



                                    t = 1
                                    If (IsTemplate) Then
                                        Send("  <td class=""templine"">")
                                        Send("    <input type=""checkbox"" id=""chkLocked5"" name=""chkLocked"" value=""" & couponObj.DeliverableID & """" & IIf(couponObj.DisallowEdit = True, " checked=""checked""", "") & " onclick=""javascript:updateLocked('lockPts" & couponObj.DeliverableID & "', this.checked);"" />")
                                        Send("    <input type=""hidden"" id=""rewPts" & couponObj.DeliverableID & """ name=""rew"" value=""" & couponObj.DeliverableID & """ />")
                                        Send("    <input type=""hidden"" id=""lockPts" & couponObj.DeliverableID & """ name=""locked"" value=""" & IIf(couponObj.DisallowEdit = True, "1", "0") & """ />")
                                        Send("  </td>")
                                    ElseIf (FromTemplate) Then
                                        Send("  <td class=""templine"">" & Copient.PhraseLib.Lookup("term." & IIf(couponObj.DisallowEdit = True, "yes", "no"), LanguageID) & "</td>")
                                    End If
                                    Send("</tr>")
                                Next
                            End If
                            ' Proximity Message reward(s)
                            t = 1
                            MyCommon.QueryStr = "SELECT P.ID, D.LASTUPDATE, D.DELIVERABLEID, D.DISALLOWEDIT FROM CPE_DELIVERABLES D INNER JOIN PROXIMITYMESSAGE P ON D.OutputID=P.ID WHERE D.REWARDOPTIONID=" & RewardID & " and D.DeliverableTypeID=14;"
                            rst = MyCommon.LRT_Select()
                            If (rst.Rows.Count > 0) Then
                                RewardsCount += rst.Rows.Count
                                AddOptionArray.Set(0, True) 'Changing this to True will allow multiple Proximity Message rewards. False will allow only one Proximity Message
                                Send("<tr class=""shadeddark"">")
                                Send("  <td colspan=""" & 3 & """>")
                                Send("    <h3>" & Copient.PhraseLib.Lookup("term.proximitymessagereward", LanguageID) & "</h3>")
                                Send("  </td>")
                                If TierLevels = 1 Then
                                    Send("  <td></td>")
                                Else
                                    For t = 1 To TierLevels
                                        Send("  <td>")
                                        Send("    <b>" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</b>")
                                        Send("  </td>")
                                    Next
                                End If
                                t = 1
                                If (IsTemplate Or FromTemplate) Then
                                    Send("  <td></td>")
                                End If
                                Send("</tr>")
                                m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
                                For Each row In rst.Rows
                                    If Not IsTemplate Then
                                        RewardDisabled = IIf((Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False)) And Not IsOfferWaitingForApproval(OfferID)), "", " disabled=""disabled""")
                                    Else
                                        RewardDisabled = IIf((Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer), "", " disabled=""disabled""")
                                    End If
                                    Send("<tr class=""shaded"">")
                                    Send("<td>")
                                    If ((Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
                                        Sendb("  <input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & RewardDisabled & " value=""X"" ")
                                        Send("onClick=""if(confirm('" & Copient.PhraseLib.Lookup("reward.confirmdelete", LanguageID) & "')) LoadDocument('UEoffer-rew.aspx?mode=DeleteProximityMessage&OfferID=" & OfferID & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "&ProximityMessageID=" & MyCommon.NZ(row.Item("ID"), 0) & "')"" />")
                                    End If
                                    Send("</td>")
                                    Send("<td><a href=""javascript:openPopup('UEoffer-rew-proximitymsg.aspx?OfferID=" & OfferID & "&RewardID=" & row("DeliverableID") & "&Phase=3" & "&PMID=" & row("ID") & "')"">" & Copient.PhraseLib.Lookup("term.proximitymessage", LanguageID) & "</a></td>")
                                    MyCommon.QueryStr = "select distinct ut.PhraseId from ProximityMessage pm join CPE_UnitTypes ut on (pm.ThresholdTypeID=ut.UnitTypeId) where pm.ID =" & row("Id")

                                    ' Find the per-tier details, and build up the details string:
                                    MyCommon.QueryStr = "select PM.ThresholdTypeID, PMT.TierLevel ,PMT.TriggerValue from ProximityMessage as PM " &
                                                        "left join ProximityMessageTier as PMT " &
                                                        "on PM.ID = PMT.ProximityMessageId " &
                                                        "left join CPE_Deliverables as CPED " &
                                                        "on CPED.OutputID = PM.ID " &
                                                        "where CPED.DeliverableTypeID = 14 and Deleted = 0 and CPED.RewardOptionID = " & ParentROID & " and PM.ID =" & row("Id")
                                    rst2 = MyCommon.LRT_Select
                                    If rst2.Rows.Count = 0 Then
                                        Send(" <td></td> <td colspan=""" & TierLevels & """>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                                    Else
                                        Dim valueType As Integer = Integer.Parse(rst2.Rows(0)(0))
                                        Dim valueSymbol As String = ""
                                        Dim valueAbbr As String = ""
                                        Dim valueLabel As String = ""
                                        Dim valueName As String = ""
                                        Dim valuePrecision As String = ""
                                        Dim tempPrecision As Integer = 0
                                        Select Case valueType
                                            Case 1
                                                valueSymbol = ""
                                                valueAbbr = ""
                                                valueName = Copient.PhraseLib.Lookup("term.items", LanguageID)
                                                valueLabel = Copient.PhraseLib.Lookup("term.quantityaway", LanguageID) & ":<br/> "
                                                tempPrecision = 0
                                            Case 2
                                                valueSymbol = Localizer.Get_Currency_Symbol(ParentROID)
                                                valueAbbr = ""
                                                valueName = Localizer.Get_Currency_Name(ParentROID)
                                                valueLabel = Copient.PhraseLib.Lookup("term.amountaway", LanguageID) & ":<br/> "
                                                tempPrecision = Localizer.Get_Currency_Precision(ParentROID)
                                            Case 3
                                                valueSymbol = ""
                                                valueAbbr = " lbs/gals"
                                                valueName = Copient.PhraseLib.Lookup("term.lbsgals", LanguageID)
                                                valueLabel = Copient.PhraseLib.Lookup("term.quantityaway", LanguageID) & ":<br/> "
                                                tempPrecision = 3
                                            Case CPEUnitTypes.Weight
                                                Dim quantityInfo As QuantityUnitTypeInfoRec = Localizer.Load_QuantityInfo_ForOffer(ParentROID, CPEUnitTypes.Weight)
                                                valueSymbol = ""
                                                valueAbbr = quantityInfo.Abbrevation
                                                valueName = quantityInfo.Name
                                                valueLabel = Copient.PhraseLib.Lookup("term.quantityaway", LanguageID) & ":<br/> "
                                                tempPrecision = quantityInfo.Precision
                                            Case CPEUnitTypes.Volume
                                                Dim quantityInfo As QuantityUnitTypeInfoRec = Localizer.Load_QuantityInfo_ForOffer(ParentROID, CPEUnitTypes.Volume)
                                                valueSymbol = ""
                                                valueAbbr = quantityInfo.Abbrevation
                                                valueName = quantityInfo.Name
                                                valueLabel = Copient.PhraseLib.Lookup("term.quantityaway", LanguageID) & ":<br/> "
                                                tempPrecision = quantityInfo.Precision
                                            Case CPEUnitTypes.Length
                                                Dim quantityInfo As QuantityUnitTypeInfoRec = Localizer.Load_QuantityInfo_ForOffer(ParentROID, CPEUnitTypes.Length)
                                                valueSymbol = ""
                                                valueAbbr = quantityInfo.Abbrevation
                                                valueName = quantityInfo.Name
                                                valueLabel = Copient.PhraseLib.Lookup("term.quantityaway", LanguageID) & ":<br/> "
                                                tempPrecision = quantityInfo.Precision
                                            Case CPEUnitTypes.SurfaceArea
                                                Dim quantityInfo As QuantityUnitTypeInfoRec = Localizer.Load_QuantityInfo_ForOffer(ParentROID, CPEUnitTypes.SurfaceArea)
                                                valueSymbol = ""
                                                valueAbbr = quantityInfo.Abbrevation
                                                valueName = quantityInfo.Name
                                                valueLabel = Copient.PhraseLib.Lookup("term.quantityaway", LanguageID) & ":<br/> "
                                                tempPrecision = quantityInfo.Precision
                                            Case 9
                                                valueSymbol = ""
                                                valueAbbr = " points"
                                                valueName = Copient.PhraseLib.Lookup("term.points", LanguageID)
                                                valueLabel = Copient.PhraseLib.Lookup("term.quantityaway", LanguageID) & ":<br/> "
                                                tempPrecision = 0
                                        End Select

                                        Select Case tempPrecision
                                            Case 0
                                                valuePrecision = "0"
                                            Case 1
                                                valuePrecision = "0.0"
                                            Case 2
                                                valuePrecision = "0.00"
                                        End Select

                                        Send("  <td>" & valueName & "</td>")
                                        While t <= TierLevels
                                            If t > rst2.Rows.Count Then
                                                Send("<td>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                                            Else
                                                Send("<td>" & valueLabel & valueSymbol & Math.Round(CDec(rst2.Rows(t - 1)(2)), tempPrecision).ToString(MyCommon.GetAdminUser.Culture) & " " & valueAbbr & "</td>")
                                            End If
                                            t += 1
                                        End While
                                        t = 1
                                    End If
                                    If (IsTemplate) Then
                                        Send("  <td class=""templine"">")
                                        Send("    <input type=""checkbox"" id=""chkLocked1"" name=""chkLocked"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, " checked=""checked""", "") & " onclick=""javascript:updateLocked('lockDisc" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "', this.checked);"" />")
                                        Send("    <input type=""hidden"" id=""rewDisc" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""rew"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ />")
                                        Send("    <input type=""hidden"" id=""lockDisc" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""locked"" value=""" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "1", "0") & """ />")
                                        Send("  </td>")
                                    ElseIf (FromTemplate) Then
                                        Send("  <td class=""templine"">" & Copient.PhraseLib.Lookup("term." & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "yes", "no"), LanguageID) & "</td>")
                                    End If
                                    Send("</tr>")
                                Next
                            End If

                        %>
                    </tbody>
                </table>
                <hr class="hidden" />
            </div>
            <div class="box" id="newreward">
                <h2>
                    <span>
                        <% Sendb(Copient.PhraseLib.Lookup("offer-rew.addreward", LanguageID))%>
                    </span>
                </h2>
                <%
                    DiscountWorthy = (MyCommon.Fetch_UE_SystemOption(126) = "1")
                    If Not DiscountWorthy Then
                        'First set the DiscountWorthy variable, which determines if the offer is eligible to use discount rewards
                        MyCommon.QueryStr = "select RO.IncentiveID from CPE_IncentiveTenderTypes as ITT with (NoLock) " &
                                            "left join CPE_RewardOptions as RO with (NoLock) on RO.RewardOptionID=ITT.RewardOptionID " &
                                            "where IncentiveID=" & OfferID & " and RO.Deleted=0;"
                        rst = MyCommon.LRT_Select
                        DiscountWorthy = (rst.Rows.Count = 0)
                    End If

                    If IsFooterOffer AndAlso Not AddOptionArray.Get(1) Then
                        Send(Copient.PhraseLib.Lookup("ueoffer-rew.FooterMessage", LanguageID))
                    ElseIf DeferCalcToTotal = True Then
                        'do nothing?
                    Else
                        'Check if Customer Approval is present for this Offer, disable gift card
                        custAppResult = m_CustCondService.GetCustomerApprovalByROID(ParentROID)
                        If custAppResult.Result IsNot Nothing AndAlso custAppResult.Result.CustomerApprovalID > 0 Then disableGiftCard = True

                        If IsTemplate Then
                            Send("<span class=""temp"">")
                            Send("  <input type=""checkbox"" class=""tempcheck"" id=""Disallow_Rewards"" name=""Disallow_Rewards""" & IIf(Disallow_Rewards, " checked=""checked""", "") & " />")
                            Send("  <label for=""Disallow_Rewards"">" & Copient.PhraseLib.Lookup("term.locked", LanguageID) & "</label>")
                            Send("</span>")
                        End If
                        TempQuerystr = "SELECT PECT.EngineID, PECT.EngineSubTypeID, PECT.ComponentTypeID, DT.DeliverableTypeID, DT.Description, DT.PhraseID, PECT.Singular, " &
                                            "  CASE DeliverableTypeID " &
                                            "    WHEN 1 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=1) " &
                                            "    WHEN 2 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=2) " &
                                            "    WHEN 4 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=4) " &
                                            "    WHEN 5 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=5) " &
                                            "    WHEN 6 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=6) " &
                                            "    WHEN 7 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=7) " &
                                            "    WHEN 8 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=8) " &
                                            "    WHEN 9 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=9) " &
                                            "    WHEN 10 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=10) " &
                                            "    WHEN 11 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=11) " &
                                            "    WHEN 12 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=12) " &
                                            "    WHEN 13 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=13) " &
                                            "    WHEN 14 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=14) " &
                                            "    WHEN 16 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=16) " &
                                             "    WHEN 17 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=17) " &
                                            "    ELSE 0 " &
                                            "  END as ItemCount " &
                                            "FROM PromoEngineComponentTypes AS PECT " &
                                            "INNER JOIN CPE_DeliverableTypes AS DT ON DT.DeliverableTypeID=PECT.LinkID " &
                                            "WHERE EngineID=9 AND EngineSubTypeID=" & EngineSubTypeID & " AND PECT.ComponentTypeID=2 AND Enabled=1"
                        'Impose a few special limits on the query based on various factors:

                        MyCommon.QueryStr = "SELECT ReturnedItemGroup FROM CPE_IncentiveProductGroups WHERE RewardOptionID=" & ParentROID & " AND Deleted=0"
                        rst = MyCommon.LRT_Select
                        If rst.Rows.Count > 0 Then
                            For Each row In rst.Rows
                                If (row.Item("ReturnedItemGroup") = True) Then
                                    If MyCommon.Fetch_UE_SystemOption(182) = "1" Then
                                        TempQuerystr &= " AND DeliverableTypeID=9"
                                        Exit For
                                    End If
                                End If
                            Next
                        End If
                        If disableGiftCard Then
                            'The offer has customer approval condition so disable gift card
                            TempQuerystr &= " AND DeliverableTypeID<>13 "
                        End If
                        If (Not IsCustomerAssigned) Then
                            'The offer has no customer condition, so the only available rewards is graphics
                            TempQuerystr &= " AND DeliverableTypeID=1"
                        End If
                        If (IsFooterOffer) Then
                            'Based on previous logic, points can't be included in footer offers
                            TempQuerystr &= " AND DeliverableTypeID<>8"
                        End If
                        If Not DiscountWorthy Then
                            TempQuerystr &= " AND DeliverableTypeID<>2"
                        End If
                        If UEOffer_Has_AnyCustomer(MyCommon, OfferID) Then
                            'Offer has AnyCustomer selected as the customer group condition.  Disallow reward types
                            'that require knowledge of who the customer is (points, stored value, etc.)
                            TempQuerystr &= " AND DeliverableTypeID not in (5, 10"
                            'verify whether point program with alloanycustomer setting
                            MyCommon.QueryStr = "Select Count(PP.ProgramID) as ProgramCount from PointsPrograms PP inner join PointsProgramsPromoEngineSettings PPPES with (NoLock) on PP.ProgramID =PPPES.ProgramID  where Deleted=0 and PP.ProgramID is not null And AllowAnyCustomer = 1"
                            rst = MyCommon.LRT_Select
                            If rst(0)("ProgramCount") <= 0 Then
                                TempQuerystr &= ",8"
                            End If
                            'verify whether store value program with alloanycustomer setting
                            MyCommon.QueryStr = "Select Count(SVP.SVProgramID) as ProgramCount from StoredValuePrograms  SVP inner join SVProgramsPromoEngineSettings  SVPPES with (NoLock) on SVP.SVProgramID =SVPPES.SVProgramID  where Deleted=0 and SVP.SVProgramID is not null And AllowAnyCustomer = 1"
                            rst = MyCommon.LRT_Select
                            If rst(0)("ProgramCount") <= 0 Then
                                TempQuerystr &= ",11,16"
                            End If
                            TempQuerystr &= ")"
                        End If
                        If restrictRewardforRPOS Then 'do not show these rewards when ue option id  234 is enabled (AMS-14478)
                            TempQuerystr &= " AND DeliverableTypeID not in (12, 13, 17) "
                        End If
                        TempQuerystr &= " ORDER BY DisplayOrder;"
                        MyCommon.QueryStr = TempQuerystr
                        rst = MyCommon.LRT_Select

                        Dim giftcardUsed As Boolean = False
                        'AMS-2257: discount and Monetory stored value are mutually exclusive          
                        Dim monetorySVUsed As Boolean = False
                        Dim discountUsed As Boolean = False
                        'cashier message,discount,group membership
                        Dim hideGiftCardOption As Boolean = False
                        Dim deliverableTypeID As Integer
                        If (rst.Rows.Count > 0) Then
                            For Each row In rst.Rows
                                deliverableTypeID = row.Item("DeliverableTypeID")
                                If (deliverableTypeID = 13 And row.Item("ItemCount") = 1) Then
                                    giftcardUsed = True
                                ElseIf ((deliverableTypeID = 2 OrElse deliverableTypeID = 5 OrElse deliverableTypeID = 9 OrElse deliverableTypeID = 12) And row.Item("ItemCount") >= 1) Then
                                    hideGiftCardOption = True
                                End If
                                If (deliverableTypeID = 2 And row.Item("ItemCount") >= 1) Then
                                    discountUsed = True
                                End If
                                If (deliverableTypeID = 16 And row.Item("ItemCount") >= 1) Then
                                    monetorySVUsed = True
                                End If
                            Next
                        End If

                        Dim lstPassThruRewards As List(Of Models.PassThrough) = m_PassThruReward.GetOfferPassThroughRewards(OfferID, 9).Result
                        Dim PassThruCount As Integer = 0
                        Dim IsEligibleReward As Boolean = False
                        If rst.Rows.Count > 0 Then
                            Send("<label for=""newrewglobal"">" & Copient.PhraseLib.Lookup("offer-rew.addglobal", LanguageID) & ":</label><br />")
                            Send("<select id=""newrewglobal"" name=""newrewglobal""" & DeleteBtnDisabled & ">")
                            For Each row In rst.Rows
                                deliverableTypeID = row.Item("DeliverableTypeID")
                                If (row.Item("Singular") = False) OrElse (row.Item("Singular") = True AndAlso row.Item("ItemCount") = 0) Then
                                    If deliverableTypeID = 12 Then
                                        'Type 12 is passthrus -- a special case, since each passthru must be shown as its own reward type
                                        MyCommon.QueryStr = "select PTR.PassThruRewardID, PTR.Name, PTR.PhraseID, PEPT.Singular from PassThruRewards PTR with (NoLock) " &
                                                            "INNER JOIN PromoEnginePassThrus PEPT with (NoLock) ON PEPT.PassThruRewardID = PTR.PassThruRewardID AND PEPT.Enabled = 1 AND PEPT.EngineID = @EngineID " &
                                                            "order by PassThruRewardID"
                                        MyCommon.DBParameters.Add("@EngineID", SqlDbType.Int).Value = 9
                                        rst2 = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                                        If rst2.Rows.Count > 0 Then
                                            For Each row2 In rst2.Rows
                                                If (row2.Item("Singular") = True AndAlso row.Item("ItemCount") > 0) Then
                                                    PassThruCount = (From passthru In lstPassThruRewards
                                                                     Where passthru.PassThroughRewardID = row2.Item("PassThruRewardID")
                                                                     Select passthru.PassThroughID).Count
                                                End If
                                                If (row2.Item("Singular") = False OrElse (row2.Item("Singular") = True AndAlso PassThruCount = 0)) Then
                                                    'If If a gift card reward is added to an offer, the user will NOT be allowed to add "cashier message,GWP,PWP,discount,group membership" into the same offer and vice versa.
                                                    If (giftcardUsed = False) Then
                                                        Sendb("<option value=""" & (row2.Item("PassThruRewardID") + 12) & """>")
                                                        If IsDBNull(row2.Item("PhraseID")) Then
                                                            Sendb(MyCommon.NZ(row2.Item("Name"), Copient.PhraseLib.Lookup("term.passthru", LanguageID)))
                                                        Else
                                                            Sendb(Copient.PhraseLib.Lookup(row2.Item("PhraseID"), LanguageID))
                                                        End If
                                                        Send("</option>")
                                                        IsEligibleReward = True
                                                    End If
                                                End If
                                            Next
                                        End If
                                        'If If a gift card reward is added to an offer, the user will NOT be allowed to add "cashier message,GWP,PWP,discount,group membership" into the same offer and vice versa.
                                    ElseIf ((deliverableTypeID = 2 And (giftcardUsed = False AndAlso monetorySVUsed = False) And Not DeferCalcToTotal) OrElse ((deliverableTypeID = 5 OrElse deliverableTypeID = 9) And giftcardUsed = False And Not DeferCalcToTotal) Or (deliverableTypeID = 13 And hideGiftCardOption = False And Not DeferCalcToTotal)) Then
                                        Send("<option value=""" & row.Item("DeliverableTypeID") & """>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "</option>")
                                        IsEligibleReward = True
                                    ElseIf ((deliverableTypeID = 4 Or deliverableTypeID = 8 Or deliverableTypeID = 11 Or deliverableTypeID = 17) AndAlso Not DeferCalcToTotal) Then
                                        'All the other types
                                        Send("<option value=""" & row.Item("DeliverableTypeID") & """>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "</option>")
                                        IsEligibleReward = True
                                    ElseIf (deliverableTypeID = 14 AndAlso Not DeferCalcToTotal) Then
                                        Dim ProductConditionExistsFlag As Boolean = True
                                        Dim PointsConditionExistsFlag As Boolean = True

                                        MyCommon.QueryStr = "SELECT IPG.QtyUnitType, IPG.IncentiveProductGroupID,IPG.QtyForIncentive FROM dbo.CPE_IncentiveProductGroups AS IPG " &
                                                            "LEFT JOIN CPE_RewardOptions AS RO " &
                                                            "ON RO.RewardOptionID = IPG.RewardOptionID " &
                                                            "WHERE RO.IncentiveID = " & OfferID & " and IPG.Deleted = 0 and IPG.ExcludedProducts=0"
                                        Dim rst4 As Data.DataTable
                                        rst4 = MyCommon.LRT_Select()
                                        MyCommon.QueryStr = "SELECT IPG.IncentivePointsID,IPG.QtyForIncentive FROM dbo.CPE_IncentivePointsGroups AS IPG " &
                                                            "LEFT JOIN CPE_RewardOptions AS RO " &
                                                            "ON RO.RewardOptionID = IPG.RewardOptionID " &
                                                            "WHERE RO.IncentiveID = " & OfferID & " and IPG.Deleted = 0"
                                        Dim rst45 As Data.DataTable
                                        rst45 = MyCommon.LRT_Select()

                                        MyCommon.QueryStr = "select * from CPE_Deliverables CPE LEFT JOIN CPE_RewardOptions AS RO " &
                                                            "ON RO.RewardOptionID = CPE.RewardOptionID " &
                                                            "WHERE RO.IncentiveID = " & OfferID & "  and CPE.DeliverableTypeID=14 and CPE.OutputID<>0"
                                        Dim rst46 As Data.DataTable
                                        rst46 = MyCommon.LRT_Select()
                                        If (Not rst46.Rows.Count >= 2) Then
                                            If (Not (rst4.Rows.Count = 0 And rst45.Rows.Count = 0)) Then
                                                'If (Not (rst4.Rows.Count > 1 And rst45.Rows.Count = 1) Or (rst4.Rows.Count = 1 And rst45.Rows.Count > 1)) Then
                                                If (Not (rst4.Rows.Count > 1 And rst45.Rows.Count > 1)) Then
                                                    If (Not (rst4.Rows.Count > 1 And rst45.Rows.Count = 0)) Then
                                                        If (rst46.Rows.Count = 1) Then
                                                            If (rst4.Rows.Count = 1 And rst45.Rows.Count = 1) Then
                                                                If (Not (rst4.Rows.Count = 1 AndAlso (Convert.ToInt32(rst4.Rows(0)("QtyUnitType")) = 4))) Then
                                                                    If (Not (rst4.Rows.Count = 1 AndAlso ((Convert.ToInt32(rst4.Rows(0)("QtyUnitType")) = 1 AndAlso Convert.ToInt32(rst4.Rows(0)("QtyForIncentive")) = 1) Or Convert.ToDouble(rst4.Rows(0)("QtyForIncentive")) = 0.1 Or Convert.ToDouble(rst4.Rows(0)("QtyForIncentive")) = 0.01))) Then
                                                                        If (Not (rst4.Rows.Count = 0 AndAlso rst45.Rows.Count = 1 AndAlso Convert.ToInt32(rst45.Rows(0)("QtyForIncentive")) = "1")) Then
                                                                            Send("<option value=""" & row.Item("DeliverableTypeID") & """>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "</option>")
                                                                            IsEligibleReward = True
                                                                        End If
                                                                    Else
                                                                        If ((rst45.Rows.Count = 1 AndAlso Convert.ToInt32(rst45.Rows(0)("QtyForIncentive")) <> 1) AndAlso Convert.ToDouble(rst4.Rows(0)("QtyForIncentive")) = 0.1 AndAlso Convert.ToDouble(rst4.Rows(0)("QtyForIncentive")) = 0.01) Then
                                                                            Send("<option value=""" & row.Item("DeliverableTypeID") & """>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "</option>")
                                                                            IsEligibleReward = True
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        Else
                                                            If (Not (rst4.Rows.Count = 1 AndAlso (Convert.ToInt32(rst4.Rows(0)("QtyUnitType")) = 4))) Then
                                                                If (Not (rst4.Rows.Count = 1 AndAlso ((Convert.ToInt32(rst4.Rows(0)("QtyUnitType")) = 1 AndAlso Convert.ToInt32(rst4.Rows(0)("QtyForIncentive")) = 1) Or (Convert.ToDouble(rst4.Rows(0)("QtyForIncentive")) = 0.1 Or Convert.ToDouble(rst4.Rows(0)("QtyForIncentive")) = 0.01) Or (Convert.ToInt32(rst4.Rows(0)("QtyUnitType")) = 2 AndAlso Convert.ToDouble(rst4.Rows(0)("QtyForIncentive")) = 0.01) Or (Convert.ToInt32(rst4.Rows(0)("QtyUnitType")) = 3 AndAlso Convert.ToDouble(rst4.Rows(0)("QtyForIncentive")) = 0.01)))) Then
                                                                    If (Not (rst4.Rows.Count = 0 AndAlso rst45.Rows.Count = 1 AndAlso Convert.ToInt32(rst45.Rows(0)("QtyForIncentive")) = 1)) Then
                                                                        Send("<option value=""" & row.Item("DeliverableTypeID") & """>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "</option>")
                                                                        IsEligibleReward = True
                                                                    End If
                                                                Else
                                                                    If ((rst45.Rows.Count = 1 AndAlso Convert.ToInt32(rst45.Rows(0)("QtyForIncentive")) <> "1")) Then
                                                                        Send("<option value=""" & row.Item("DeliverableTypeID") & """>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "</option>")
                                                                        IsEligibleReward = True
                                                                    End If
                                                                End If
                                                            End If

                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    ElseIf (deliverableTypeID = 15 AndAlso UEOffer_Has_AnyCustomer(MyCommon, OfferID) = False AndAlso EPMInstalled AndAlso Not DeferCalcToTotal) Then
                                        Send("<option value=""" & row.Item("DeliverableTypeID") & """>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "</option>")
                                        IsEligibleReward = True
                                    ElseIf (deliverableTypeID = 16 AndAlso UEOffer_Has_AnyCustomer(MyCommon, OfferID) = False AndAlso IsCustomerAssigned = True AndAlso discountUsed = False) Then
                                        Dim rstMSv As DataTable
                                        MyCommon.QueryStr = "select 1 from StoredValuePrograms with (NoLock) where Deleted=0 and SVTypeID=2;"
                                        rstMSv = MyCommon.LRT_Select()
                                        If (rstMSv.Rows.Count > 0) Then
                                            Send("<option value=""" & row.Item("DeliverableTypeID") & """>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "</option>")
                                            IsEligibleReward = True
                                        End If
                                    End If
                                End If
                            Next
                            If (Not IsEligibleReward) Then
                                Send("<option value=""-1"" >" & Copient.PhraseLib.Lookup("reward.noeligible", LanguageID) & "</option>")
                            End If
                            Send("</select>")
                            If ((Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable) AndAlso Not IsOfferWaitingForApproval(OfferID)) Then
                                Send("<input class=""regular"" id=""addglobal"" name=""addglobal"" type=""submit"" value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """" & DeleteBtnDisabled & IIf(IsEligibleReward, "", "disabled=""disabled"" ") & " /><br />")
                            End If
                        End If
                    End If
                %>
            </div>
        </div>
        <br clear="all" />
    </div>
</form>
<!-- #Include virtual="/include/graphic-reward.inc" -->
<%
    If (RewardsCount = 0 AndAlso IsTemplate) Then
        Send("<script type=""text/javascript"">")
        Send("document.getElementById(""rewards"").style.display = 'none';")
        Send("</script>")
    ElseIf (RewardsCount = 0) Then
        Send("<script type=""text/javascript"">")
        Send("document.getElementById(""rewards"").style.display = 'none';")
        Send("</script>")
    End If
%>
<%
    If MyCommon.Fetch_SystemOption(75) Then
        If (OfferID > 0 And Logix.UserRoles.AccessNotes) Then
            Send_Notes(3, OfferID, AdminUserID)
        End If
    End If
done:
    MyCommon.Close_LogixRT()
    Send_BodyEnd()
    MyCommon = Nothing
    Logix = Nothing
%>
