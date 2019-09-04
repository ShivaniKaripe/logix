<%@ Page Language="vb" Debug="true" CodeFile="ue-cb.vb" Inherits="UECB" %>

<%@ Import Namespace="Copient.Localization" %>

<%  ' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="CMS.AMS" %>
<%@ Import Namespace="CMS.AMS.Contract" %>
<%@ Import Namespace="CMS.AMS.Models" %>
<%@ Import Namespace="CMS.Contract" %>
<%
    ' *****************************************************************************
    ' * FILENAME: UEoffer-sum.aspx 
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
    Dim isStartEndTimeEnabled As Boolean
    Dim MyCommon As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    Dim rst As DataTable
    Dim row As DataRow
    Dim rst2 As DataTable
    Dim row2 As DataRow
    Dim rst3 As DataTable
    Dim row3 As DataRow
    Dim rst4 As DataTable
    Dim row4 As DataRow
    Dim rstTiers As DataTable
    Dim OfferID As Long = Request.QueryString("OfferID")
    Dim IsTemplate As Boolean
    Dim IsTemplateVal As String = "Not"
    Dim FromTemplate As Boolean = False
    Dim ActiveSubTab As Integer = 24
    Dim roid As Integer = 0
    Dim i As Integer = 0
    Dim Days As String = ""
    Dim Times As String = ""
    Dim Tenders As String = ""
    Dim Cookie As HttpCookie = Nothing
    Dim BoxesValue As String = ""
    Dim ShowCRM As Boolean = True
    Dim ValidateIncentiveColor As String = "green"
    Dim ComponentColor As String = "green"
    Dim IsDeployable As Boolean = False
    Dim LinksDisabled As Boolean = False
    Dim Expired As Boolean = False
    Dim TempDateTime As New DateTime
    Dim ShowActionButton As Boolean = False
    Dim SourceOfferID As Integer
    Dim DeployBtnIEOffset As Integer = 0
    Dim DeployBtnFFOffset As Integer = 0
    Dim NewUpdateLevel As Integer = 0
    Dim GCount As Integer = 0
    Dim AnyProductUsed As Boolean = False
    Dim AnyCustomerUsed As Boolean = False
    Dim AnyStoreUsed As Boolean = False
    Dim LongDate As New DateTime
    Dim LongDateZeroTime As New DateTime
    Dim TodayDateZeroTime As New DateTime
    Dim DaysDiff As Integer = 0
    Dim rowCount As Integer = 0
    Dim counter As Integer = 0
    Dim ErrorMsg As String = ""
    Dim infoMessage As String = ""
    Dim warnMessage As String = String.Empty
    Dim modMessage As String = ""
    Dim Handheld As Boolean = False
    Dim BannerNames As String() = Nothing
    Dim BannerIDs As Integer() = Nothing
    Dim BannerCt As Integer = 0
    Dim BannersEnabled As Boolean = False
    Dim StatusMessage As String = ""
    Dim StatusCode As Copient.LogixInc.STATUS_FLAGS = Copient.LogixInc.STATUS_FLAGS.STATUS_UNKNOWN
    Dim objCouponService As ICouponRewardService
    Dim StatusText As String = ""
    Dim TenderList As String = ""
    Dim TenderValue As String = ""
    Dim TenderRequired As Boolean
    Dim TenderExcluded As Boolean
    Dim TenderExcludedAmt As Object = Nothing
    Dim Popup As Boolean = False
    Dim TierLevels As Integer = 1
    Dim t As Integer = 1
    Dim OfferImported As Boolean = False
    Dim ImportMessage As String = ""
    Dim ProdGroupList As String = "-1"
    Dim EngineID As Integer = 9
    Dim EngineSubTypeID As Integer = 0
    Dim FolderNames As String = ""
    Dim x As Integer = 0
    Dim Localizer As Copient.Localization
    Dim AmountTypeID As Integer
    Dim ActivityLogMsg As String
    Dim bUseDisplayDates As Boolean = False
    Dim Offer As Models.Offer
    Dim offerValidationLogFilePrefix As String = "OfferValidationLog"
    Dim IsCollisionDetectionEnabled As Boolean = False
    Dim CollisionDetectionEnabledResp As AMSResult(Of Boolean)
    Dim RunCollision As Integer = -1
    Dim bUseMultipleProductExclusionGroups As Boolean = True
    Dim Currency As String = ""
    Dim m_ProductConditionPGService As IProductConditionService
    'Dim exclusionPGList As List(Of ProductConditionProductGroup)
    Dim m_DiscountPGService As IDiscountRewardService
    'Dim discountexclusionPGList As List(Of IDiscountRewardService)
    Const POS_CHANNEL_ID As Integer = 1
    Dim offerStatus As Int16 = 0
    Dim rejectMessage As String = ""
    CurrentRequest.Resolver.AppName = "UEoffer-sum.aspx"
    m_ProductConditionPGService = CurrentRequest.Resolver.Resolve(Of IProductConditionService)()
    m_DiscountPGService = CurrentRequest.Resolver.Resolve(Of IDiscountRewardService)()
    Dim m_Offer As IOffer = CurrentRequest.Resolver.Resolve(Of IOffer)()
    objCouponService = CurrentRequest.Resolver.Resolve(Of ICouponRewardService)()
    Dim m_CollisionDetectionService As ICollisionDetectionService = CurrentRequest.Resolver.Resolve(Of ICollisionDetectionService)()
    Dim m_CustomerConditionService As ICustomerGroupCondition = CurrentRequest.Resolver.Resolve(Of ICustomerGroupCondition)()
    Dim m_customerGroup As ICustomerGroups = CurrentRequest.Resolver.Resolve(Of ICustomerGroups)()
    Dim m_TrackableCouponCondition As ITrackableCouponConditionService = CurrentRequest.Resolver.Resolve(Of ITrackableCouponConditionService)()
    Dim m_TCProgram As ITrackableCouponProgramService = CurrentRequest.Resolver.Resolve(Of ITrackableCouponProgramService)()
    Dim m_Logger As ILogger = CurrentRequest.Resolver.Resolve(Of ILogger)()
    Dim m_PreferenceService As IPreferenceService = CurrentRequest.Resolver.Resolve(Of IPreferenceService)()
    Dim m_PreferenceRewardService As IPreferenceRewardService = CurrentRequest.Resolver.Resolve(Of IPreferenceRewardService)()
    Dim m_AnalyticsCustomerGroups As IAnalyticsCustomerGroups = CurrentRequest.Resolver.Resolve(Of IAnalyticsCustomerGroups)()
    Dim m_OAWService As IOfferApprovalWorkflowService = CurrentRequest.Resolver.Resolve(Of IOfferApprovalWorkflowService)()
    Dim isTranslatedOffer As Boolean = MyCommon.IsTranslatedUEOffer(OfferID, MyCommon)
    Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249) = "1", True, False)
    Dim bEnableAdditionalLockoutRestrictionsOnOffers As Boolean = IIf(MyCommon.Fetch_SystemOption(260) = "1", True, False)
    Dim bOfferEditable As Boolean = True
    Dim isBuyer As Boolean = False
    Dim bUpdateDatesForExpiredOfferCopy As Boolean=False

    Dim m_isOAWEnabled As Boolean = False
    Dim m_isExternalOffer As Boolean = False
    Dim m_hasOAWPermissions As Boolean = False
    Dim m_hasOfferModifiedAfterApproval As Boolean = False
    Dim m_offerApprovalStatus As Int32 = -1
    Dim m_requiresOfferApproval As Boolean = False

    'AMSPS-1402
    Dim bRetainLockedFieldsInCopiedOffers As Boolean = IIf(MyCommon.Fetch_SystemOption(283)="1",True,False)
    'AMSPS-1402 above

    Dim bEnableCopyOffer As Boolean = IIf(MyCommon.Fetch_SystemOption(286) = "1", True, False)
    Dim isCopyOfferEnabledForGeneralUser As Boolean = False

    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If

    Response.Expires = 0

    MyCommon.AppName = "UEoffer-sum.aspx"
    MyCommon.Open_LogixRT()
    MyCommon.Open_LogixXS()
    If MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER) Then
        MyCommon.Open_PrefManRT()
    End If
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
    Localizer = New Copient.Localization(MyCommon)


    MyCommon.QueryStr = "SELECT AU.AdminUserID FROM AdminUsers AU (NoLock) inner join BuyerRoleUsers BU (NoLock) ON BU.AdminUserID = AU.AdminUserID WHERE AU.AdminUserID = " & AdminUserID & ";"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
        isBuyer = MyCommon.NZ(rst.Rows(0).Item("AdminUserID"), False)
    End If


    If (isBuyer And Not Logix.UserRoles.EditOfferPastLockoutPeriod) Then
        bOfferEditable = True
    Else
        bOfferEditable = MyCommon.IsOfferEditablePastLockOutPeriod(Logix.UserRoles.EditOfferPastLockoutPeriod, MyCommon, OfferID)
    End If

    'Check if general user has separate copy offer permission
    If bEnableCopyOffer Then
        If (FromTemplate AndAlso Logix.UserRoles.CopyOfferCreatedFromTemplate) Then
            isCopyOfferEnabledForGeneralUser = True
        ElseIf (Not FromTemplate AndAlso Logix.UserRoles.CopyOfferCreatedFromBlank) Then
            isCopyOfferEnabledForGeneralUser = True
        End If
    End If

    'If ((Logix.UserRoles.ViewOffersRegardlessBuyer = False OrElse Logix.UserRoles.EditOffersRegardlessBuyer = False) AndAlso MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID) = False) Then
    '  Response.Redirect("/logix/offer-list.aspx")
    'End If
    isStartEndTimeEnabled = (MyCommon.Fetch_UE_SystemOption(200) = "1")
    If MyCommon.Fetch_UE_SystemOption(143) = "1" Then
        bUseDisplayDates = True
    Else
        bUseDisplayDates = False
    End If

    BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
    Popup = IIf((Request.QueryString("Popup") <> "") AndAlso (Request.QueryString("Popup") <> "0"), True, False)

    MyCommon.QueryStr = "select RewardOptionID, TierLevels from CPE_RewardOptions with (NoLock) where TouchResponse=0 and IncentiveID=" & OfferID & ";"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
        roid = MyCommon.NZ(rst.Rows(0).Item("RewardOptionID"), 0)
        TierLevels = MyCommon.NZ(rst.Rows(0).Item("TierLevels"), 1)
    End If

    MyCommon.QueryStr = "select C.Abbreviation As Abbreviation from CPE_RewardOptions CRO with (NoLock) " & _
                        " Inner join Currencies C with  (NoLock) on C.CurrencyID = CRO.currencyID  where IncentiveID= " & OfferID & ";"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
        Currency = MyCommon.NZ(rst.Rows(0).Item("Abbreviation"), 0)
    End If

    MyCommon.QueryStr = "select EngineSubTypeID from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID & ";"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
        EngineSubTypeID = MyCommon.NZ(rst.Rows(0).Item("EngineSubTypeID"), 0)
    End If

    ' load up all the folder names to which this offer is assigned.
    MyCommon.QueryStr = "select distinct FI.FolderID, F.FolderName from FolderItems as FI with (NoLock) " & _
                        "inner join Folders as F with (NoLock) on F.FolderID = FI.FolderID " & _
                        "where LinkID=" & OfferID & " and LinkTypeID=1;"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
        For Each row In rst.Rows
            If FolderNames <> "" Then FolderNames &= " <br />"
            If (bEnableAdditionalLockoutRestrictionsOnOffers AndAlso Not bOfferEditable) OrElse m_OAWService.CheckIfOfferIsAwaitingApproval(OfferID).Result Then
                FolderNames &= MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("FolderName"), ""), 25)
            Else
                FolderNames &= "<a href=""javascript:openPopup('/logix/folder-browse.aspx?Action=NavigateToFolder&OfferID=" & OfferID & _
                               "&FolderID=" & MyCommon.NZ(row.Item("FolderID"), "0") & "');"">" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("FolderName"), ""), 25) & "</a>"
            End If

        Next
    Else
        If (bEnableAdditionalLockoutRestrictionsOnOffers AndAlso Not bOfferEditable) OrElse m_OAWService.CheckIfOfferIsAwaitingApproval(OfferID).Result Then
            FolderNames &= Copient.PhraseLib.Lookup("term.none", LanguageID)
        Else
            FolderNames = "<a href=""javascript:openPopup('/logix/folder-browse.aspx?OfferID=" & OfferID & "');"">" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "</a>"
        End If
    End If
    If (Request.QueryString("new") <> "") Then
        If Request.QueryString("IsTemplate") <> "" Then
            IsTemplate = (Request.QueryString("IsTemplate") = "IsTemplate")
            If IsTemplate Then
                Response.Redirect("../offer-new.aspx?NewTemplate=Yes&new=New")
            Else
                Response.Redirect("../offer-new.aspx")
            End If
        End If
    ElseIf (Request.QueryString("OfferFromTemp") <> "") Then
        Try
            ' dbo.pc_CreateOfferFromTemplate @TemplateID bigint, @OfferID bigint OUTPUT
            MyCommon.QueryStr = "dbo.pc_Create_CPE_OfferFromTemplate"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@SourceOfferID", SqlDbType.NVarChar, 200).Value = OfferID
            MyCommon.LRTsp.Parameters.Add("@CreatedByAdminId", SqlDbType.Int).Value = AdminUserID
            MyCommon.LRTsp.Parameters.Add("@NewOfferID", SqlDbType.BigInt).Direction = ParameterDirection.Output
            MyCommon.LRTsp.ExecuteNonQuery()
            SourceOfferID = OfferID
            OfferID = MyCommon.LRTsp.Parameters("@NewOfferID").Value
            MyCommon.Close_LRTsp()
            Dim bIds As Integer() = Logix.GetBannersForOffer(OfferID)
            If OfferID > 0 And IsOAWEnabled(OfferID, bIds) Then
                m_OAWService.InsertUpdateOfferApprovalRecord(OfferID, AdminUserID)
            End If
            SetPromotionDisplay(MyCommon, OfferID)
            If bUseDisplayDates Then
                'Updating TemplatePermission table with the Disallow_DisplayDates based on the UE SystemOption #143
                UpdateTemplatePermissions(MyCommon, SourceOfferID, OfferID, 143)
                SaveOfferDisplayDates(MyCommon, SourceOfferID, OfferID, AdminUserID, EngineID)
            End If
            If (OfferID > 0 AndAlso bUseMultipleProductExclusionGroups) Then
                UpdateInclusionIncentiveProductGroupSet(MyCommon, OfferID)
            End If
            MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("offer.createdfromtemplate", LanguageID) & ": " & SourceOfferID)
            Response.Status = "301 Moved Permanently"
            Response.AddHeader("Location", "UEoffer-gen.aspx?OfferID=" & OfferID)
            GoTo done
        Catch ex As Exception
            If ex.Message = "error.couldnot-processoffers" Then
                infoMessage = Copient.PhraseLib.Lookup(ex.Message, LanguageID)
            Else
                infoMessage = ex.Message
            End If
        End Try
    ElseIf (Request.QueryString("saveastemp") <> "") Then
        Try
            MyCommon.QueryStr = "dbo.pc_Create_CPE_TemplateFromOffer"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@SourceOfferID", SqlDbType.NVarChar, 200).Value = OfferID
            MyCommon.LRTsp.Parameters.Add("@CreatedByAdminId", SqlDbType.Int).Value = AdminUserID
            MyCommon.LRTsp.Parameters.Add("@NewOfferID", SqlDbType.BigInt).Direction = ParameterDirection.Output
            MyCommon.LRTsp.ExecuteNonQuery()
            SourceOfferID = OfferID
            OfferID = MyCommon.LRTsp.Parameters("@NewOfferID").Value
            MyCommon.Close_LRTsp()
            If bUseDisplayDates Then
                'Updating TemplatePermission table with the Disallow_DisplayDates based on the UE SystemOption #143
                UpdateTemplatePermissions(MyCommon, SourceOfferID, OfferID, 143)
                SaveOfferDisplayDates(MyCommon, SourceOfferID, OfferID, AdminUserID, EngineID)
            End If
            SetPromotionDisplay(MyCommon, OfferID)
            MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("templates.createdfromoffer", LanguageID) & ": " & SourceOfferID)
            Response.Status = "301 Moved Permanently"
            Response.AddHeader("Location", "UEoffer-gen.aspx?OfferID=" & OfferID)
            GoTo done
        Catch ex As Exception
            If ex.Message = "error.couldnot-processoffers" Then
                infoMessage = Copient.PhraseLib.Lookup(ex.Message, LanguageID)
            Else
                infoMessage = ex.Message
            End If
        End Try
    ElseIf (Request.QueryString("deploy") <> "") Then
        IsDeployable = IsDeployableOffer(MyCommon, OfferID, roid, ErrorMsg, Logix)
        If (IsDeployable) Then
            MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=2,UpdateLevel=UpdateLevel+1, LastDeployClick=getdate(), DeployDeferred=0, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where IncentiveID=" & OfferID
            MyCommon.LRT_Execute()
            ActivityLogMsg = "[history.offer-deploy]"
            If GetCgiValue("deploytransreqskip") = "1" Then
                ActivityLogMsg = ActivityLogMsg & vbCrLf & "<BR>" & "- [history.offer-deploy.trasnrequired]" & " " & CheckForTranslationDeployError(MyCommon, roid)
            End If
            'Prevent the ActivityLogMessage from being longer than ActivityLog.Description can hold
            ActivityLogMsg = Copient.PhraseLib.DecodeEmbededTokens(ActivityLogMsg, LanguageID)
            If ActivityLogMsg.Length > 1000 Then ActivityLogMsg = Left(ActivityLogMsg, 995) & " ..."
            'Write the message to the ActivityLog
            MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.DecodeEmbededTokens(ActivityLogMsg, LanguageID))
            Response.Status = "301 Moved Permanently"
            Response.AddHeader("Location", "UEoffer-sum.aspx?OfferID=" & OfferID)
            SetLastDeployValidationMessage(MyCommon, OfferID, "term.validationsuccessful")
            m_Logger.WriteInfo(String.Format("Engine:{0}, OfferId:{1}, Message:{2}", Engines.UE, OfferID, Copient.PhraseLib.Lookup("term.validationsuccessful", LanguageID)), offerValidationLogFilePrefix)
            GoTo done
        Else
            infoMessage = ErrorMsg
            If EngineSubTypeID = 1 Then
                infoMessage &= "  An instant win condition with randomized triggers must also be present."
            End If
            SetLastDeployValidationMessage(MyCommon, OfferID, "<font color=""red"">" & infoMessage & "</font>")
            m_Logger.WriteInfo(String.Format("Engine:{0}, OfferId:{1}, Message:{2}", Engines.UE, OfferID, infoMessage), offerValidationLogFilePrefix)
        End If
    ElseIf (Request.QueryString("sendoutbound") <> "") Then
        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastCRMSendDate=getdate(), CRMEngineUpdateLevel=CRMEngineUpdateLevel+1, CRMSendToExport=1 where IncentiveID=" & OfferID
        MyCommon.LRT_Execute()
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-sendtocrm", LanguageID))
    ElseIf (Request.QueryString("delete") <> "") Then
        Dim optInGroup As CustomerGroup = m_Offer.GetOfferDefaultCustomerGroup(OfferID, EngineID)

        'Delete analytics cg for this offer if it isn't linked to any other offer.
        m_AnalyticsCustomerGroups.DeleteDefaultAnalyticsCustomerGroupForOffer(OfferID)

        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=2, Deleted=1, LastUpdate=getdate(), UpdateLevel=UpdateLevel+1 where IncentiveID=" & OfferID
        MyCommon.LRT_Execute()
        'Mark the shadow table offer as deleted as well.
        MyCommon.QueryStr = "update CPE_ST_Incentives with (RowLock) set Deleted=1, LastUpdate=getdate(), UpdateLevel = (select UpdateLevel from CPE_Incentives with (NoLock)where IncentiveID=" & OfferID & ") where IncentiveID=" & OfferID
        MyCommon.LRT_Execute()
        'Mark Client ID deleted if this is enternal offer.
        MyCommon.QueryStr = "dbo.pt_ExtOfferID_Delete"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.NVarChar, 20).Value = OfferID
        MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = EngineID
        MyCommon.LRTsp.ExecuteNonQuery()
        MyCommon.Close_LRTsp()
        'Remove any Offer Eligibility Conditions associated with the Offer.
        m_Offer.DeleteOfferEligibleConditions(OfferID, EngineID)
        If (optInGroup IsNot Nothing) Then
            m_customerGroup.DeleteCustomerGroup(optInGroup.CustomerGroupID)
        End If

        MyCommon.QueryStr = "DELETE FROM CPE_IncentivePLUs where RewardOptionID=" & OfferID
        MyCommon.LRT_Execute()

        'Remove the banners assigned to this offer
        If (BannersEnabled) Then
            MyCommon.QueryStr = "delete from BannerOffers with (RowLock) where OfferID = " & OfferID
            MyCommon.LRT_Execute()
        End If
        'Also remove/update the triggers for any associated EIW conditions
        MyCommon.QueryStr = "update CPE_EIWTriggers with (RowLock) set Removed=1, LastUpdate=getdate() where RewardOptionID=" & roid & " and Removed=0;"
        MyCommon.LRT_Execute()
        'Record activity
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-delete", LanguageID))
        Response.Status = "301 Moved Permanently"
        MyCommon.QueryStr = "select IsTemplate from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID
        rst = MyCommon.LRT_Select()
        IsTemplate = MyCommon.NZ(rst.Rows(0).Item("IsTemplate"), False)
        If (IsTemplate) Then
            Response.AddHeader("Location", "/logix/temp-list.aspx")
        Else
            Response.AddHeader("Location", "/logix/offer-list.aspx")
        End If
        GoTo done
    ElseIf (Request.QueryString("export") <> "") Then
        Dim MyExport As Copient.ExportXmlUE = CurrentRequest.Resolver.Resolve(Of Copient.ExportXmlUE)()
        Dim bStatus As Boolean
        Dim bProduction As Boolean
        Dim sFileFullPathName As String
        bProduction = True ' uses production start/end date
        sFileFullPathName = MyCommon.Fetch_SystemOption(29) & "\Offer" & Request.QueryString("OfferID") & ".gz"
        bStatus = MyExport.GenerateOfferXML(Request.QueryString("OfferID"), sFileFullPathName, bProduction)
        If Not bStatus Then
            If (MyExport.ErrorMessage.Contains("term.")) Then
                infoMessage = Copient.PhraseLib.Lookup(MyExport.ErrorMessage, LanguageID)
            Else
                infoMessage = MyExport.ErrorMessage
            End If


        Else
            If (MyExport.GetFileType = Copient.ExportXmlUE.FileTypeEnum.XML_FORMAT) Then
                Dim oRead As System.IO.StreamReader
                Dim LineIn As String
                Dim Bom As String = ChrW(65279)
                oRead = System.IO.File.OpenText(sFileFullPathName)
                Response.ContentEncoding = Encoding.UTF8
                Response.ContentType = "application/octet-stream"
                Response.AddHeader("Content-Disposition", "attachment; filename=" & "Offer" & Request.QueryString("OfferID").ToString & ".xml")

                'force little endian fffe bytes at front, why?  i dont know but is required.
                Sendb(Bom)
                While oRead.Peek <> -1
                    LineIn = oRead.ReadLine()
                    Send(LineIn)
                End While
                oRead.Close()
            ElseIf (MyExport.GetFileType = Copient.ExportXmlUE.FileTypeEnum.GZ_FORMAT) Then
                Dim fs As System.IO.FileStream
                Dim br As System.IO.BinaryReader
                Response.ContentType = "application/octet-stream"
                Response.AddHeader("Content-Disposition", "attachment; filename=" & "Offer" & Request.QueryString("OfferID").ToString & ".gz")
                fs = System.IO.File.Open(sFileFullPathName, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read)
                br = New System.IO.BinaryReader(fs)
                Response.BinaryWrite(br.ReadBytes(fs.Length))
                fs.Flush()
                br.Close()
                fs.Close()
                fs = Nothing
            End If
            System.IO.File.Delete(sFileFullPathName)
            MyCommon.Activity_Log(3, Request.QueryString("OfferID"), AdminUserID, Copient.PhraseLib.Lookup("offer.exported", LanguageID))
            GoTo done
        End If
    ElseIf (Request.QueryString("exportedw") <> "") Then
        Dim MyExport As Copient.ExportXmlUE = CurrentRequest.Resolver.Resolve(Of Copient.ExportXmlUE)()
        Dim bStatus As Boolean
        Dim bProduction As Boolean
        Dim EDWFilePath As String = ""
        Dim EDWFileName As String = ""

        bProduction = True ' uses production start/end date
        EDWFilePath = MyCommon.Fetch_SystemOption(73).Trim
        If (Right(MyCommon.Fetch_SystemOption(73), 1) <> "\") Then EDWFilePath &= "\"
        EDWFileName = "Offer" & Request.QueryString("OfferID") & "_" & Now.ToString("yyyy-MM-dd_HHmmss")

        MyExport.SetFileType(Copient.ExportXmlUE.FileTypeEnum.XML_FORMAT)
        MyExport.SetTableType(Copient.ExportXmlUE.TableTypeEnum.DEPLOYED)
        bStatus = MyExport.GenerateOfferXML(Request.QueryString("OfferID"), EDWFilePath & EDWFileName & ".xml", bProduction)
        If Not bStatus Then
            infoMessage = Copient.PhraseLib.Lookup("cpeoffer-sum.exportedwalert", LanguageID)
        Else
            ' write out a check file to state that the exported file is ready.
            System.IO.File.WriteAllText(EDWFilePath & EDWFileName & ".ok", EDWFilePath & EDWFileName & ".xml" & ControlChars.Tab & "OK")

            MyCommon.QueryStr = "dbo.pa_UpdateOfferFeeds"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = Request.QueryString("OfferID")
            MyCommon.LRTsp.Parameters.Add("@LastFeed", SqlDbType.DateTime).Value = Now()
            MyCommon.LRTsp.ExecuteNonQuery()
            MyCommon.Close_LRTsp()

            StatusMessage = Copient.PhraseLib.Lookup("cpeoffer-sum.exportedwok", LanguageID)
        End If
    ElseIf (Request.QueryString("deferdeploy") <> "") Then
        IsDeployable = IsDeployableOffer(MyCommon, OfferID, roid, ErrorMsg, Logix)
        If (IsDeployable) Then
            MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set DeployDeferred=1, UpdateLevel=UpdateLevel+1, LastDeployClick=getdate(), LastUpdate=getdate() where IncentiveID=" & OfferID
            MyCommon.LRT_Execute()
            ActivityLogMsg = "[history.offer-deferdeploy]"
            If GetCgiValue("deploytransreqskip") = "1" Then
                ActivityLogMsg = ActivityLogMsg & vbCrLf & "<BR>" & "- [history.offer-deploy.trasnrequired]" & " " & CheckForTranslationDeployError(MyCommon, roid)
            End If
            'Prevent the ActivityLogMessage from being longer than ActivityLog.Description can hold
            ActivityLogMsg = Copient.PhraseLib.DecodeEmbededTokens(ActivityLogMsg, LanguageID)
            If ActivityLogMsg.Length > 1000 Then ActivityLogMsg = Left(ActivityLogMsg, 995) & " ..."
            'Write the message to the ActivityLog
            MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.DecodeEmbededTokens(ActivityLogMsg, LanguageID))
            Response.Status = "301 Moved Permanently"
            Response.AddHeader("Location", "UEoffer-sum.aspx?OfferID=" & OfferID)
            GoTo done
        Else
            infoMessage = ErrorMsg
        End If
    ElseIf (Request.QueryString("canceldeploy") <> "") Then
        ' check if the offer is still in awaiting deployment status
        MyCommon.QueryStr = "select StatusFlag, DeployDeferred from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
            ' update status to modified (1) if offer is still awaiting deployment, otherwise alert user that offer was already deployed.
            If (MyCommon.NZ(rst.Rows(0).Item("StatusFlag"), -1) = 2) Or (MyCommon.NZ(rst.Rows(0).Item("DeployDeferred"), False) = True) Then
                MyCommon.QueryStr = "select LastUpdateLevel from PromoEngineUpdateLevels with (NoLock) " &
                                    "where LinkID=" & OfferID & " and EngineID=9 and ItemType=1;"
                rst = MyCommon.LRT_Select
                If (rst.Rows.Count > 0) Then
                    NewUpdateLevel = MyCommon.NZ(rst.Rows(0).Item("LastUpdateLevel"), 0)
                End If
                MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=1, DeployDeferred=0, UpdateLevel=" & NewUpdateLevel & " where IncentiveID=" & OfferID
                MyCommon.LRT_Execute()
                Dim offerStatusResult As AMSResult(Of Integer) = m_OAWService.GetOfferApprovalStatus(OfferID)
                If m_isOAWEnabled AndAlso offerStatusResult.Result <> 2 Then
                    ResetOfferApprovalStatus(OfferID)
                End If
                MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.canceldeploy", LanguageID))
                infoMessage = Copient.PhraseLib.Lookup("term.deploymentcanceled", LanguageID)
            ElseIf (MyCommon.NZ(rst.Rows(0).Item("StatusFlag"), -1) = 0) Then
                infoMessage = Copient.PhraseLib.Lookup("message.alreadydeployed", LanguageID)
            Else
                infoMessage = Copient.PhraseLib.Lookup("message.unablecanceldeployment", LanguageID)
            End If
        End If
    ElseIf (Request.QueryString("copyoffer") <> "") Then
        Try
            'AMSPS-1402
            MyCommon.QueryStr = "select IsTemplate, FromTemplate from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID & ";"
            rst = MyCommon.LRT_Select()
            IsTemplate = MyCommon.NZ(rst.Rows(0).Item("IsTemplate"), False)
            FromTemplate = MyCommon.NZ(rst.Rows(0).Item("FromTemplate"), False)
            'AMSPS-1402 above
            If IsTemplate Then
                MyCommon.QueryStr = "dbo.pc_Create_CPE_TemplateFromOffer"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@SourceOfferID", SqlDbType.NVarChar, 200).Value = OfferID
                MyCommon.LRTsp.Parameters.Add("@CreatedByAdminId", SqlDbType.Int).Value = AdminUserID
                MyCommon.LRTsp.Parameters.Add("@NewOfferID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                MyCommon.LRTsp.ExecuteNonQuery()
                SourceOfferID = OfferID
                OfferID = MyCommon.LRTsp.Parameters("@NewOfferID").Value
                MyCommon.Close_LRTsp()
                If bUseDisplayDates Then
                    'Updating TemplatePermission table with the Disallow_DisplayDates based on the UE SystemOption #143
                    UpdateTemplatePermissions(MyCommon, Request.QueryString("OfferID"), OfferID, 143)
                    SaveOfferDisplayDates(MyCommon, Request.QueryString("OfferID"), OfferID, AdminUserID, EngineID)
                End If
                SetPromotionDisplay(MyCommon, OfferID)
                'AMSPS-1402
            ElseIf (bRetainLockedFieldsInCopiedOffers And FromTemplate) Then
                MyCommon.QueryStr = "dbo.pc_Create_CPE_OfferFromTemplate"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@SourceOfferID", SqlDbType.NVarChar, 200).Value = OfferID
                MyCommon.LRTsp.Parameters.Add("@CreatedByAdminId", SqlDbType.Int).Value = AdminUserID
                MyCommon.LRTsp.Parameters.Add("@NewOfferID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                MyCommon.LRTsp.ExecuteNonQuery()
                SourceOfferID = OfferID
                OfferID = MyCommon.LRTsp.Parameters("@NewOfferID").Value
                MyCommon.Close_LRTsp()
                Dim bIds As Integer() = Logix.GetBannersForOffer(OfferID)
                If OfferID > 0 And IsOAWEnabled(OfferID, bIds) Then
                    m_OAWService.InsertUpdateOfferApprovalRecord(OfferID, AdminUserID)
                End If
                SetPromotionDisplay(MyCommon, OfferID)
                If bUseDisplayDates Then
                    'Updating TemplatePermission table with the Disallow_DisplayDates based on the UE SystemOption #143
                    UpdateTemplatePermissions(MyCommon, SourceOfferID, OfferID, 143)
                    SaveOfferDisplayDates(MyCommon, SourceOfferID, OfferID, AdminUserID, EngineID)
                End If

                'AMSPS-1402 above
            Else
                MyCommon.QueryStr = "dbo.pc_Copy_CPE_Offer"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@SourceOfferID", SqlDbType.BigInt).Value = Request.QueryString("OfferID")
                MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = 9
                MyCommon.LRTsp.Parameters.Add("@EngineSubTypeID", SqlDbType.Int).Value = EngineSubTypeID
                MyCommon.LRTsp.Parameters.Add("@AdminUserID", SqlDbType.BigInt).Value = AdminUserID
                MyCommon.LRTsp.Parameters.Add("@NewOfferID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                MyCommon.LRTsp.ExecuteNonQuery()
                OfferID = MyCommon.LRTsp.Parameters("@NewOfferID").Value
                'MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set FromTemplate=0 where IncentiveID=" & OfferID & ";"
                MyCommon.Close_LRTsp()
                Dim bIds As Integer() = Logix.GetBannersForOffer(OfferID)
                If OfferID > 0 And IsOAWEnabled(OfferID, bIds) Then
                    m_OAWService.InsertUpdateOfferApprovalRecord(OfferID, AdminUserID)
                End If
                If bUseDisplayDates Then
                    'Updating TemplatePermission table with the Disallow_DisplayDates based on the UE SystemOption #143
                    UpdateTemplatePermissions(MyCommon, Request.QueryString("OfferID"), OfferID, 143)
                    SaveOfferDisplayDates(MyCommon, Request.QueryString("OfferID"), OfferID, AdminUserID, EngineID)
                End If
                SetPromotionDisplay(MyCommon, OfferID)
            End If
            If (OfferID > 0 AndAlso isCopyOfferEnabledForGeneralUser AndAlso Not IsTemplate) Then
                StatusText = Logix.GetOfferStatus(OfferID, LanguageID, StatusCode)
                If (StatusText.Contains("Expired")) Then
                    MyCommon.QueryStr = "dbo.pt_OfferDates_Default"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = EngineID
                    MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
                    MyCommon.LRTsp.ExecuteNonQuery()
                    MyCommon.Close_LRTsp()
                End If
            End If
            If (OfferID > 0) Then
                If (bUseMultipleProductExclusionGroups) Then
                    UpdateInclusionIncentiveProductGroupSet(MyCommon, OfferID)
                End If
                MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-copy", LanguageID))
                Response.Status = "301 Moved Permanently"
                'Verify Trackable Coupon conditions
                Dim objResult1 As AMSResult(Of List(Of TCProgramCondition)) = m_TrackableCouponCondition.GetTCProgramConditions(Request.QueryString("OfferID"), EngineID)
                If (objResult1.ResultType = AMSResultType.Success AndAlso objResult1.Result.Count > 0) Then
                    Response.AddHeader("Location", "UEoffer-sum.aspx?OfferID=" & OfferID & "&copyfrom=" & Request.QueryString("OfferID"))
                Else
                    Response.AddHeader("Location", "UEoffer-sum.aspx?OfferID=" & OfferID)
                End If
                GoTo done
            Else
                OfferID = MyCommon.Extract_Val(Request.QueryString("OfferID"))
            End If
        Catch ex As Exception
            If ex.Message = "error.couldnot-processoffers" Then
                infoMessage = Copient.PhraseLib.Lookup(ex.Message, LanguageID)
            Else
                infoMessage = ex.Message
            End If
        End Try

    ElseIf (Request.QueryString("copyfrom") <> String.Empty) Then
        Dim ii As Int64 = 0
        If (Int64.TryParse(Request.QueryString("copyfrom"), ii) AndAlso ii > 0) Then
            Dim objResult1 As AMSResult(Of List(Of TCProgramCondition)) = m_TrackableCouponCondition.GetTCProgramConditions(ii, EngineID)
            If (objResult1.ResultType = AMSResultType.Success AndAlso objResult1.Result.Count > 0) Then
                warnMessage = String.Format(Copient.PhraseLib.Lookup("ueoffersum.copyofferwarning", LanguageID), MyCommon.TruncateString(objResult1.Result.Item(0).TCProgram.Name, 25))
            End If
        End If
    ElseIf (Request.QueryString("viewreport") <> String.Empty) Then
        Response.Redirect("..\CollidingOffers-Report.aspx?ID=" & OfferID.ToString())
    ElseIf (Request.QueryString("cancelcollisiondetection") <> String.Empty) Then
        Dim IsCollideOfferUpdated As AMSResult(Of Boolean) = m_CollisionDetectionService.UpdateCollideOfferStatus(OfferID, Models.OCD.QueueStatus.Cancel)
        If (IsCollideOfferUpdated.ResultType = AMSResultType.Success) Then
            'warnMessage = "Collision detection cancelled successfully."
        End If
    ElseIf (Request.QueryString("getRecommendations") <> String.Empty) Then
        MyCommon.QueryStr = "update ExtSegmentMap set ExtSegmentID = -1 where IncentiveID = " & OfferID & "and ExtSegmentID = 0"
        MyCommon.LRT_Execute()
        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=11, DeployDeferred=0, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where IncentiveID=" & OfferID
        MyCommon.LRT_Execute()
        ActivityLogMsg = "[alert.awaitingrecommendation]"
        ActivityLogMsg = Copient.PhraseLib.DecodeEmbededTokens(ActivityLogMsg, LanguageID)
        If ActivityLogMsg.Length > 1000 Then ActivityLogMsg = Left(ActivityLogMsg, 995) & " ..."
        'Write the message to the ActivityLog
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.DecodeEmbededTokens(ActivityLogMsg, LanguageID))
        Response.Status = "301 Moved Permanently"
        Response.AddHeader("Location", "UEoffer-sum.aspx?OfferID=" & OfferID)
        m_Logger.WriteInfo(String.Format("Engine:{0}, OfferId:{1}, Message:{2}", Engines.UE, OfferID, "[alert.awaitingrecommendation]"), offerValidationLogFilePrefix)
        GoTo done
    ElseIf (Request.QueryString("getRecommendationsAndDeploy") <> String.Empty) Then
        IsDeployable = IsDeployableOffer(MyCommon, OfferID, roid, ErrorMsg, Logix)
        If (IsDeployable) Then
            MyCommon.QueryStr = "update ExtSegmentMap set ExtSegmentID = -1 where IncentiveID = " & OfferID & " and ExtSegmentID = 0"
            MyCommon.LRT_Execute()
            MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=12, DeployDeferred=0, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where IncentiveID=" & OfferID
            MyCommon.LRT_Execute()
            ActivityLogMsg = "[alert.awaitingrecommendation] and later it will be deployed"
            'Prevent the ActivityLogMessage from being longer than ActivityLog.Description can hold
            ActivityLogMsg = Copient.PhraseLib.DecodeEmbededTokens(ActivityLogMsg, LanguageID)
            If ActivityLogMsg.Length > 1000 Then ActivityLogMsg = Left(ActivityLogMsg, 995) & " ..."
            'Write the message to the ActivityLog
            MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.DecodeEmbededTokens(ActivityLogMsg, LanguageID))
            Response.Status = "301 Moved Permanently"
            Response.AddHeader("Location", "UEoffer-sum.aspx?OfferID=" & OfferID)
            SetLastDeployValidationMessage(MyCommon, OfferID, "term.validationsuccessful")
            m_Logger.WriteInfo(String.Format("Engine:{0}, OfferId:{1}, Message:{2}", Engines.UE, OfferID, "[alert.awaitingrecommendation] and later it will be deployed"), offerValidationLogFilePrefix)
            GoTo done
        Else
            infoMessage = ErrorMsg
            If EngineSubTypeID = 1 Then
                infoMessage &= "  An instant win condition with randomized triggers must also be present."
            End If
            SetLastDeployValidationMessage(MyCommon, OfferID, "<font color=""red"">" & infoMessage & "</font>")
            m_Logger.WriteInfo(String.Format("Engine:{0}, OfferId:{1}, Message:{2}", Engines.UE, OfferID, infoMessage), offerValidationLogFilePrefix)
        End If
    ElseIf (Request.QueryString("reqApproval") <> String.Empty Or Request.QueryString("reqApprovalWithDeployment") <> String.Empty Or Request.QueryString("reqApprovalWithDeferDeployment") <> String.Empty) Then
        'Offer is validated and can be deployed. Now based on req query string, perform one of the operation.
        IsDeployable = IsDeployableOffer(MyCommon, OfferID, roid, ErrorMsg, Logix)
        If (IsDeployable) Then
            Dim tempStatusFlag As Int16 = 13
            ActivityLogMsg = ""

            If (Request.QueryString("reqApproval") <> String.Empty) Then
                ActivityLogMsg = Copient.PhraseLib.Lookup("history.reqapproval", LanguageID)
            ElseIf (Request.QueryString("reqApprovalWithDeployment") <> String.Empty) Then
                tempStatusFlag = 14
                ActivityLogMsg = Copient.PhraseLib.Lookup("history.reqapprovaldeploy", LanguageID)
            ElseIf (Request.QueryString("reqApprovalWithDeferDeployment") <> String.Empty) Then
                tempStatusFlag = 15
                ActivityLogMsg = Copient.PhraseLib.Lookup("history.reqapprovaldeferdeploy", LanguageID)
            End If


            MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=" & tempStatusFlag & ", DeployDeferred=0, LastUpdate=getdate(), " &
                                "LastUpdatedByAdminID =" & AdminUserID & " where IncentiveID=" & OfferID
            MyCommon.LRT_Execute()
            m_OAWService.InsertUpdateOfferApprovalRecord(OfferID, AdminUserID)

            'Prevent the ActivityLogMessage from being longer than ActivityLog.Description can hold
            ActivityLogMsg = Copient.PhraseLib.DecodeEmbededTokens(ActivityLogMsg, LanguageID)
            If ActivityLogMsg.Length > 1000 Then ActivityLogMsg = Left(ActivityLogMsg, 995) & " ..."
            'Write the message to the ActivityLog
            MyCommon.Activity_Log(3, OfferID, AdminUserID, ActivityLogMsg)
            Response.Status = "301 Moved Permanently"
            Response.AddHeader("Location", "UEoffer-sum.aspx?OfferID=" & OfferID)
            SetLastDeployValidationMessage(MyCommon, OfferID, "term.validationsuccessful")
            m_Logger.WriteInfo(String.Format("Engine:{0}, OfferId:{1}, Message:{2}", Engines.UE, OfferID, ActivityLogMsg), offerValidationLogFilePrefix)
            GoTo done
        Else
            infoMessage = ErrorMsg
            If EngineSubTypeID = 1 Then
                infoMessage &= "  An instant win condition with randomized triggers must also be present."
            End If
            SetLastDeployValidationMessage(MyCommon, OfferID, "<font color=""red"">" & infoMessage & "</font>")
            m_Logger.WriteInfo(String.Format("Engine:{0}, OfferId:{1}, Message:{2}", Engines.UE, OfferID, infoMessage), offerValidationLogFilePrefix)
        End If
    ElseIf (Request.QueryString("cancelgetapproval") <> "") Then
        ' check if the offer is still in waiting for approval.
        MyCommon.QueryStr = "select StatusFlag from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
            If (MyCommon.NZ(rst.Rows(0).Item("StatusFlag"), -1) > 12) Then
                MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=1 where IncentiveID=" & OfferID
                MyCommon.LRT_Execute()
                ResetOfferApprovalStatus(OfferID)
                MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.canceloaw", LanguageID))
                infoMessage = Copient.PhraseLib.Lookup("term.oawcancelled", LanguageID)
                m_OAWService.SendNotificationEmail(4, OfferID, AdminUserID)
            Else
                infoMessage = Copient.PhraseLib.Lookup("message.unablecanceloaw", LanguageID)
            End If
        End If
    ElseIf (Request.QueryString("rejectOffer") <> "") Then
        Dim message As String = Logix.TrimAll(Request.QueryString("rejectText"))
        Dim result As Boolean = m_OAWService.RejectOffer(OfferID, AdminUserID, message).Result
        Dim logMessage As String = Copient.PhraseLib.Lookup("term.offer-rejected", LanguageID)
        If message <> "" Then logMessage = logMessage & " " & Copient.PhraseLib.Lookup("term.rejectionreason", LanguageID) & ": " & message
        MyCommon.Activity_Log(3, OfferID, AdminUserID, logMessage)
        Response.Status = "301 Moved Permanently"
        Response.AddHeader("Location", "UEoffer-sum.aspx?OfferID=" & OfferID)
    ElseIf (Request.QueryString("clearMessage") <> "") Then
        Dim result As Boolean = m_OAWService.ResetOfferApprovalStatus(OfferID).Result
        Response.Status = "301 Moved Permanently"
        Response.AddHeader("Location", "UEoffer-sum.aspx?OfferID=" & OfferID)
    End If

    If (Request.QueryString("OfferID") <> "") Then
        MyCommon.QueryStr = "select IncentiveID, B.ExternalBuyerId, OID.EngineID, PE.PhraseID as EnginePhraseID, PEST.PhraseID as EngineSubTypePhraseID, " &
                            " IsTemplate, ClientOfferID, IncentiveName, CPE.Description, FromTemplate, " &
                            " PromoClassID, CRMEngineID, Priority, StartDate, EndDate, EveryDOW, VendorCouponCode, EligibilityStartDate, " &
                            " EligibilityEndDate, TestingStartDate, TestingEndDate, P1DistQtyLimit, P1DistTimeType, " &
                            " P1DistPeriod, P2DistQtyLimit, P2DistTimeType, P2DistPeriod, P3DistQtyLimit, " &
                            " P3DistTimeType, P3DistPeriod, EnableImpressRpt, EnableRedeemRpt, CPE.CreatedDate, CPE.LastUpdate, CPE.Deleted, ExportToEDW, " &
                            " CPEOARptDate, CPEOADeploySuccessDate, CPEOADeployRpt, CRMRestricted, StatusFlag, DeployDeferred, " &
                            " OC.OfferCategoryID, OC.Description as CategoryName, " &
                            " AU1.FirstName + ' ' + AU1.LastName as CreatedBy, AU2.FirstName + ' ' + AU2.LastName as LastUpdatedBy, OID.Imported, " &
                            " CPE.EngineSubTypeID, CPE.InboundCRMEngineID " &
                            " from CPE_Incentives as CPE with (NoLock) " &
                            " left join OfferIDs as OID with (NoLock) on OID.OfferID=CPE.IncentiveID " &
                            " left join PromoEngines as PE with (NoLock) on PE.EngineID=OID.EngineID " &
                            " left join PromoEngineSubTypes as PEST with (NoLock) on PEST.PromoEngineID=OID.EngineID and PEST.SubTypeID=OID.EngineSubTypeID " &
                            " left join OfferCategories as OC with (NoLock) on CPE.PromoClassID=OfferCategoryID " &
                            " left join AdminUsers as AU1 with (NoLock) on AU1.AdminUserID = CPE.CreatedByAdminID " &
                            " left join AdminUsers as AU2 with (NoLock) on AU2.AdminUserID = CPE.LastUpdatedByAdminID " &
                            " left join Buyers as B with (NoLock) on B.BuyerId = CPE.BuyerId " &
                            " where IncentiveID=" & Request.QueryString("OfferID") & " and CPE.Deleted=0;"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count < 1 Then
            infoMessage = Copient.PhraseLib.Lookup("error.deleted", LanguageID)
        Else
            IsTemplate = MyCommon.NZ(rst.Rows(0).Item("IsTemplate"), False)
            FromTemplate = MyCommon.NZ(rst.Rows(0).Item("FromTemplate"), False)
            LinksDisabled = IIf(MyCommon.NZ(rst.Rows(0).Item("StatusFlag"), -1) = 2 Or MyCommon.NZ(rst.Rows(0).Item("StatusFlag"), -1) > 10, True, False)
            offerStatus = MyCommon.NZ(rst.Rows(0).Item("StatusFlag"), -1)
            If (Not LinksDisabled) Then
                LinksDisabled = (MyCommon.NZ(rst.Rows(0).Item("DeployDeferred"), False) = True)
            End If
            If Popup Then
                LinksDisabled = True
            End If
            TempDateTime = MyCommon.NZ(rst.Rows(0).Item("EndDate"), New Date(1900, 1, 1))
            'TempDateTime = TempDateTime.AddDays(1)
            If TempDateTime < Today() Then
                Expired = True
            End If
            OfferImported = MyCommon.NZ(rst.Rows(0).Item("Imported"), False)
            EngineSubTypeID = MyCommon.NZ(rst.Rows(0).Item("EngineSubTypeID"), 0)
            m_isExternalOffer = MyCommon.NZ(rst.Rows(0).Item("InboundCRMEngineID"), 0)
        End If

        m_offerApprovalStatus = m_OAWService.GetOfferApprovalStatus(OfferID).Result
        If m_offerApprovalStatus = 2 OrElse offerStatus = 0 Then
            m_hasOfferModifiedAfterApproval = False
        Else
            m_hasOfferModifiedAfterApproval = True
        End If
        ' get the banner assigned to this offer
        If (BannersEnabled) Then
            MyCommon.QueryStr = "select BAN.BannerID, BAN.Name from BannerOffers BO with (NoLock) " &
                                "inner join Banners BAN with (NoLock) on BAN.BannerID = BO.BannerID " &
                                "where BO.OfferID = " & Request.QueryString("OfferID") & " order by BAN.Name;"
            rst2 = MyCommon.LRT_Select
            BannerCt = rst2.Rows.Count
            If (BannerCt > 0) Then
                ReDim BannerNames(BannerCt - 1)
                ReDim BannerIDs(BannerCt - 1)
                For i = 0 To BannerNames.GetUpperBound(0)
                    BannerNames(i) = MyCommon.SplitNonSpacedString(MyCommon.NZ(rst2.Rows(i).Item("Name"), ""), 25)
                    BannerIDs(i) = MyCommon.NZ(rst2.Rows(i).Item("BannerID"), -1)
                Next
            Else
                ReDim BannerNames(0)
                ReDim BannerIDs(0)
                BannerNames(0) = Copient.PhraseLib.Lookup("term.all", LanguageID)
                BannerIDs(i) = -1
            End If
            m_isOAWEnabled = m_OAWService.CheckIfOfferApprovalIsEnabledForBanners(BannerIDs).Result
        Else
            m_isOAWEnabled = m_OAWService.CheckIfOfferApprovalIsEnabled().Result
        End If

        If (m_isOAWEnabled) Then
            m_hasOAWPermissions = m_OAWService.CheckIfUserCanApproveOffer(OfferID, AdminUserID).Result
            m_requiresOfferApproval = m_OAWService.CheckIfUserRequiresOfferApproval(OfferID, AdminUserID).Result
        End If

        ' check import status
        If OfferImported Then
            ProdGroupList = GetAssociatedProductGroupIDs(roid, MyCommon)

            ' are there any pending records for a newly-imported offer's product group
            MyCommon.QueryStr = "select PG.ProductGroupID, PG.Name, PIQ.StatusFlag from ProductGroups as PG with (NoLock) " &
                                "inner join ProdInsertQueue as PIQ with (NoLock) on PIQ.ProductGroupID = PG.ProductGroupID " &
                                "where(PIQ.ProductGroupID in (" & ProdGroupList & ") And PG.CreatedDate = PG.LastUpdate) " &
                                "and PG.UpdateLevel=0 and PG.Deleted=0;"
            rst2 = MyCommon.LRT_Select
            If rst2.Rows.Count > 0 Then
                ImportMessage &= Copient.PhraseLib.Lookup("offer-import.pending-import", LanguageID)
                For Each row2 In rst2.Rows
                    ImportMessage &= MyCommon.NZ(row2.Item("ProductGroupID"), "") & " - " & MyCommon.NZ(row2.Item("Name"), "") & "<br />"
                Next
                ImportMessage &= "<a href=""UEOffer-sum.aspx?OfferID=" & OfferID & """>[" & Copient.PhraseLib.Lookup("term.refresh", LanguageID) & "]</a>"
            End If

            ' are there any failed product group imports for a newly-imported offer's product group
            MyCommon.QueryStr = "select ProductGroupID, Name, LastLoadMsg from ProductGroups " &
                                "where(ProductGroupID in (" & ProdGroupList & ")  And CreatedDate = LastUpdate And UpdateLevel = 0 And Deleted = 0) " &
                                "and LastLoadMsg like '%uploaded file%';"
            rst2 = MyCommon.LRT_Select
            If rst2.Rows.Count > 0 Then
                ImportMessage &= Copient.PhraseLib.Lookup("offer-import.failed-group-import", LanguageID)
                For Each row2 In rst2.Rows
                    ImportMessage &= "<a href=""../pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row2.Item("ProductGroupID"), "") & """>" &
                                     MyCommon.NZ(row2.Item("ProductGroupID"), "") & " - " & MyCommon.NZ(row2.Item("Name"), "") & "</a><br />"
                Next
            End If

        End If

    End If

    CollisionDetectionEnabledResp = m_Offer.IsCollisionDetectionEnabled(Engines.UE, OfferID)
    If (CollisionDetectionEnabledResp.ResultType = AMSResultType.Success AndAlso CollisionDetectionEnabledResp.Result = True) Then IsCollisionDetectionEnabled = True

    If IsCollisionDetectionEnabled Then
        Dim AwaitingDetectionResp As AMSResult(Of Models.OCD.QueueStatus) = m_CollisionDetectionService.GetOfferQueueStatus(OfferID)
        If (AwaitingDetectionResp.ResultType = AMSResultType.Success AndAlso (AwaitingDetectionResp.Result = OCD.QueueStatus.NotStarted OrElse AwaitingDetectionResp.Result = OCD.QueueStatus.InProgress)) Then
            OfferLockedforCollisionDetection = True
            LinksDisabled = True
        End If
    End If

    Dim _RunCollision As AMSResult(Of Integer) = m_Offer.IsAnyProductGroupAssociatedWithOffer(OfferID, 9)
    If _RunCollision.ResultType = AMSResultType.Success Then
        RunCollision = _RunCollision.Result
    End If

    ShowCRM = (MyCommon.Fetch_SystemOption(25) <> "0")
    MyCommon.QueryStr = "select CRMEngineID from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID & ";"
    rst4 = MyCommon.LRT_Select
    If rst4.Rows.Count > 0 Then
        If MyCommon.Extract_Val(MyCommon.NZ(rst4.Rows(0).Item("CRMEngineID"), 0)) = 0 Then
            ShowCRM = False
        End If
    End If

    If (IsTemplate) Then
        ActiveSubTab = 25
        IsTemplateVal = "IsTemplate"
    Else
        ActiveSubTab = 24
        IsTemplateVal = "Not"
    End If

    Send_HeadBegin("term.offer", "term.summary", OfferID)
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts()
    Send_HeadEnd()
    If Popup Then
        If (IsTemplate) Then
            Send_BodyBegin(13)
        Else
            Send_BodyBegin(3)
        End If
    Else
        If (IsTemplate) Then
            Send_BodyBegin(14)
        Else
            Send_BodyBegin(4)
        End If
        Send_Bar(Handheld)
        Send_Help(CopientFileName)
        Send_Logos()
        Send_Tabs(Logix, 2)
        If (rst.Rows.Count < 1) Then
            Send_Subtabs(Logix, 23, 3, , OfferID)
        Else
            Send_Subtabs(Logix, ActiveSubTab, 3, , OfferID)
        End If
    End If

    If (rst.Rows.Count < 1) Then
        Send("")
        Send("<div id=""intro"">")
        Send("  <h1 id=""title"">UE " & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & Request.QueryString("OfferID") & "</h1>")
        Send("</div>")
        Send("<div id=""main"">")
        Send("  <div id=""infobar"" class=""red-background"">")
        Send("    " & infoMessage)
        Send("  </div>")
        Send("</div>")
        Send("</div>")
        Send("</body>")
        Send("<script type=""text/javascript"" language=""javascript"">")
        Send("  function updateCookie() { }")
        Send("</script>")
        Send("</html>")
        GoTo done
    End If

    ' Get the user's preference from the cookie for collapsing/showing the boxes
    Cookie = Request.Cookies("BoxesCollapsed")
    If Not (Cookie Is Nothing) Then
        BoxesValue = Cookie.Value
        If (BoxesValue Is Nothing OrElse BoxesValue.Trim = "") Then
            BoxesValue = "0"
        End If
    Else
        BoxesValue = "0"
    End If
    If ((Logix.UserRoles.ViewOffersRegardlessBuyer = False OrElse Logix.UserRoles.EditOffersRegardlessBuyer = False) AndAlso MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID) = False) Then
        Send("<script type=""text/javascript"" language=""javascript"">")
        Send("  function updateCookie() { return true; } ")
        Send("</script>")
        If (Logix.UserRoles.ViewOffersRegardlessBuyer = False) Then
            Send_Denied(1, "perm.viewOffersRegardlessBuyer")
        ElseIf (Logix.UserRoles.EditOffersRegardlessBuyer = False) Then
            Send_Denied(1, "perm.editOffersRegardlessBuyer")
        Else
            Send_Denied(1, "perm.offers-access")
        End If
        Send_BodyEnd()
        GoTo done
    End If
    If (Logix.UserRoles.AccessOffers = False AndAlso Not IsTemplate) Then
        Send("<script type=""text/javascript"" language=""javascript"">")
        Send("  function updateCookie() { return true; } ")
        Send("</script>")
        Send_Denied(1, "perm.offers-access")
        Send_BodyEnd()
        GoTo done
    End If
    If (Logix.UserRoles.AccessTemplates = False AndAlso IsTemplate) Then
        Send("<script type=""text/javascript"" language=""javascript"">")
        Send("  function updateCookie() { return true; } ")
        Send("</script>")
        Send_Denied(1, "perm.offers-access-templates")
        Send_BodyEnd()
        GoTo done
    End If
    If (Logix.UserRoles.AccessInstantWinOffers = False AndAlso EngineID = 2 AndAlso EngineSubTypeID = 1) Then
        Send("<script type=""text/javascript"" language=""javascript"">")
        Send("  function updateCookie() { return true; } ")
        Send("</script>")
        Send_Denied(1, "perm.offers-accessinstantwin")
        Send_BodyEnd()
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
%>
<script type="text/javascript">
    var divElems = new Array("generalbody", "periodbody", "limitsbody","deploymentbody", 
                             "locationbody", "notificationbody", "conditionbody", "rewardbody", "validationbody","displaybody");
    var divVals  = new Array(1, 2, 4, 8, 16, 32, 64, 128, 256);
    var divImages = new Array("imgGeneral", "imgPeriod", "imgLimits", "imgDeployment",  
                              "imgLocations", "imgNotifications", "imgOptInConditions", "imgConditions", "imgRewards", "imgValidation","imgDisplay");
    var boxesValue = <% Sendb(BoxesValue)%>;
    var isCollisionDetectionEnabled = "<%=IsCollisionDetectionEnabled And RunCollision = 1 And Not Expired%>";
  
    function updateCookie() {
        updateBoxesCookie(divElems, divVals);
    }
  
    function collapseBoxes() {
        updatePageBoxes(divElems, divVals, divImages, boxesValue);
    }
  
    function showDiv(elemName) {
        var elem = document.getElementById(elemName);
    
        if (elem != null) {
            elem.style.display = (elem.style.display == "none") ? "block" : "none";
        }
    }
  
    function setComponentsColor(color) {
        var elem = document.getElementById("linkComponent");
    
        if (elem != null) {
            elem.style.color = color;
        }
    }
  
    function toggleDropdown() {
        if (document.getElementById("actionsmenu") != null) {
            bOpen = (document.getElementById("actionsmenu").style.visibility != 'visible')
            if (bOpen) {
                document.getElementById("actionsmenu").style.visibility = 'visible';
                document.mainform.actions.value = '<% Sendb(Copient.PhraseLib.Lookup("term.actions", LanguageID))%>▲';
          } else {
              document.getElementById("actionsmenu").style.visibility = 'hidden';
              document.mainform.actions.value = '<% Sendb(Copient.PhraseLib.Lookup("term.actions", LanguageID))%>▼';
    }
}
}

function confirmDeployCollision(deferred) {
    // Runs when the button (deploy or deferdeploy) in the actions menu is clicked.
    var confirmString = '';
    if (deferred) {
        confirmString = '<%Sendb(Copient.PhraseLib.Lookup("confirm.deferdeploy", LanguageID))%>';
    } else {
        confirmString = '<%Sendb(Copient.PhraseLib.Lookup("confirm.deploy", LanguageID))%>';
        confirmString = confirmString.replace('&#39;', "'");
    }

    if (confirm(confirmString)) {
        if (deferred) {
            DeferDeploy = true;
        } else {
            DeferDeploy = false;
        }
        ApprovalType = -1;
        var xml = xmlhttpPost('../OfferFeeds.aspx?Mode=IsDeployableOffer&OfferID=<%Sendb(OfferID)%>&EngineID=9', 'IsDeployableOffer');
        //showConfirmationBox();
    } else {
        return false
    }
}

  function CopyExpiredOffer(msg)
  {
    if(confirm(msg))
    {
         openPopup('/logix/folder-browse.aspx?fromCopyAction=1&OfferID=<%Sendb(OfferID)%>&EngineID=9');
    }
    else
    {
        return false;
    }
  }



  
function runDetectionFromSummaryPage()
{   
    // CallingLocation = 1 indicates Run Collision Detection Button is clicked from Action Drop Down.
    document.getElementById('CallingLocation').value = 1; 
    ApprovalType = -1;
    var xml = xmlhttpPost('../OfferFeeds.aspx?Mode=IsDeployableOffer&OfferID=<%Sendb(OfferID)%>&EngineID=9&RunCollisionDetection=1', 'IsDeployableOffer');
}
    function requestApproval(statusFlag) {
        ApprovalType = statusFlag;
        DeferDeploy = false;
        var xml = xmlhttpPost('../OfferFeeds.aspx?Mode=IsDeployableOffer&OfferID=<%Sendb(OfferID)%>&EngineID=9', 'IsDeployableOffer');
    }

    function approveOfferBackground() {
        var ocdEnabled = 0;
        if (isCollisionDetectionEnabled == "True")
            ocdEnabled = 1;
        var xml = xmlhttpPost('../OfferFeeds.aspx?Mode=ApproveOffer&OfferID=<%Sendb(OfferID)%>&ApprovalType=<%Sendb(offerStatus)%>&OCDEnabled=' + ocdEnabled, 'ProductCollisionsBackgroundUEDetection');
}
function showConfirmationBoxDetection(enableDetection) {
    // Reveals (and shows the appropriate buttons in) a box asking if the user wants to run collision detection
    var confirmationBox = document.getElementById("confirmingDetection");
    var collisionsBox = document.getElementById("loadingDetection");

    if (confirmationBox != null && collisionsBox != null) {
        toggleDialog('loadingDetection', false);
        if(enableDetection)
        {
            toggleDialog('confirmingDetection', true);      
        }
        else
        {
            toggleDialog('confirmingDisabledDetection', true);      
        }
    }
}

function loadCollisionsDetection() {
    showCollisionsBoxDetection();
    DeferDeploy = false;
    var CallingLocation = 1;
    xmlhttpPost('../OfferFeeds.aspx?Mode=GetProductCollisionsUEDetection&OfferID=<%Sendb(OfferID)%>&DeferDeploy='+DeferDeploy+'&CallingLocation='+ CallingLocation ,'GetProductCollisionsUEDetection');
}

function showCollisionsBoxDetection() {
    var collisionsBox = document.getElementById("loadingDetection");
    var confirmationBox = document.getElementById("confirmingDetection");

    if (collisionsBox != null && confirmationBox != null) {      
        toggleDialog('confirmingDetection', false);
        toggleDialog('loadingDetection', true);
        var loadingboxtext =  "<p><% Sendb(Copient.PhraseLib.Lookup("UEoffer-gen.FindingCollisions", LanguageID))%></p>" +
        "<p style=\"text-align:center;padding-top:10px;\"><img src=\"../../images/loadingAnimation.gif\" height=\"80px;\" alt=\"<% Sendb(Copient.PhraseLib.Lookup("term.loading", LanguageID))%>\" title=\"<% Sendb(Copient.PhraseLib.Lookup("term.loading", LanguageID))%>\" /></p>";
       
        updateCollisionsBoxDetection(loadingboxtext);
      
    }
}

function updateCollisionsBoxDetection(responseMsg) {
    var collisionsContent = document.getElementById("collisionsContentDetection");

    if (responseMsg.replace(/(\r\n|\n|\r)/gm,"") != '') {
        collisionsContent.innerHTML = responseMsg;
    } else {
        if (DeferDeploy) {
            window.location = 'UEoffer-sum.aspx?OfferID=<%Sendb(OfferID)%>&deferdeploy=<%Sendb(Copient.PhraseLib.Lookup("term.deferdeployment", LanguageID))%>';
        } else {
            window.location = 'UEoffer-sum.aspx?OfferID=<%Sendb(OfferID)%>&deploy=<%Sendb(Copient.PhraseLib.Lookup("term.deploy", LanguageID))%>';
        }
    }
}

function processcollisionbackgroundDetection() {
    DeferDeploy=false;
    var callingLocation=1;
    xmlhttpPost('../OfferFeeds.aspx?Mode=ProductCollisionsBackgroundUEDetection&OfferID=<%Sendb(OfferID)%>&DeferDeploy='+DeferDeploy+'&CallingLocation='+callingLocation,'ProductCollisionsBackgroundUEDetection');
}

    function showConfirmationBox() {
        // Reveals (and shows the appropriate buttons in) a box asking if the user wants to run collision detection
        var confirmationBox = document.getElementById("confirming");
        var collisionsBox = document.getElementById("loading");
        var deployButton = document.getElementById("confirmingDeploy");
        var deferDeployButton = document.getElementById("confirmingDeferDeploy");
        var approvalButton = document.getElementById("confirmingApproval");
        var deployApprovalButton = document.getElementById("confirmingDeployApproval");
        var deferDeployApprovalButton = document.getElementById("confirmingDeployDeferApproval");

        if (ApprovalType == -1) {
            deployButton.style.display = 'none';
            deferDeployButton.style.display = 'none';
            approvalButton.style.display = 'none';
            deployApprovalButton.style.display = 'none';
            deferDeployApprovalButton.style.display = 'none';

            if (DeferDeploy) {
                deferDeployButton.style.display = 'inline';
            } else {
                deployButton.style.display = 'inline';
            }
        } else {
            deployButton.style.display = 'none';
            deferDeployButton.style.display = 'none';
            approvalButton.style.display = 'none';
            deployApprovalButton.style.display = 'none';
            deferDeployApprovalButton.style.display = 'none';

            if (ApprovalType == 13)
                approvalButton.style.display = 'inline';
            else if (ApprovalType == 14)
                deployApprovalButton.style.display = 'inline';
            else if (ApprovalType == 15)
                deferDeployApprovalButton.style.display = 'inline';
        }

        if (confirmationBox != null && collisionsBox != null) {
            toggleDialog('loading', false);
            toggleDialog('confirming', true);
        }
    }
  
function loadCollisions() {
    showCollisionsBox();
    xmlhttpPost('../OfferFeeds.aspx?Mode=GetProductCollisionsUE&OfferID=<%Sendb(OfferID)%>&DeferDeploy=' + DeferDeploy + '&ApprovalType=' + ApprovalType, 'GetProductCollisionsUE');
   }

   function processcollisionbackground() {
       xmlhttpPost('../OfferFeeds.aspx?Mode=ProductCollisionsBackgroundUE&OfferID=<%Sendb(OfferID)%>&DeferDeploy=' + DeferDeploy + '&ApprovalType=' + ApprovalType, 'ProductCollisionsBackgroundUE');
}

function showCollisionsBox() {
    var collisionsBox = document.getElementById("loading");
    var confirmationBox = document.getElementById("confirming");

    if (collisionsBox != null && confirmationBox != null) {      
        toggleDialog('confirming', false);
        toggleDialog('loading', true);
        var loadingboxtext =  "<p><% Sendb(Copient.PhraseLib.Lookup("UEoffer-gen.FindingCollisions", LanguageID))%></p>" +
        "<p style=\"text-align:center;padding-top:10px;\"><img src=\"../../images/loadingAnimation.gif\" height=\"80px;\" alt=\"<% Sendb(Copient.PhraseLib.Lookup("term.loading", LanguageID))%>\" title=\"<% Sendb(Copient.PhraseLib.Lookup("term.loading", LanguageID))%>\" /></p>";
       
          updateCollisionsBox(loadingboxtext);
      }
  }

  function xmlhttpPost(strURL, action) {
      var xmlHttpReq = false;
      var self = this;
      var tokens = new Array();
      var isDeployOffer = "";
      var runbackground ="";
      if (window.XMLHttpRequest) { // Mozilla/Safari
          self.xmlHttpReq = new XMLHttpRequest();
      } else if (window.ActiveXObject) { // IE
          self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
      }
      self.xmlHttpReq.open('POST', strURL, true);
      self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
      self.xmlHttpReq.onreadystatechange = function() {
          if (self.xmlHttpReq != null && self.xmlHttpReq.readyState == 4) {
              if (action == "GetProductCollisionsUE") {
                  updateCollisionsBox(self.xmlHttpReq.responseText);
              }
              if (action == "IsDeployableOffer") {
                  isDeployOffer = self.xmlHttpReq.responseText.toString();
                  if(document.getElementById('CallingLocation').value == 1)
                  {
                      document.getElementById('CallingLocation').value = 0;
                      if(isDeployOffer.replace("\r\n","") == "True" && document.getElementById('hdnRunCollision').value == 1){
                          showConfirmationBoxDetection(true);
                      }
                      else {
                          showConfirmationBoxDetection(false);
                      }
                  }
                  else{
                      if(isDeployOffer.replace("\r\n","") == "True"){              
                          showConfirmationBox();
                      }
                      else {
                          var maindiv = document.getElementById('main'); 
                          var testdiv = "<div id='infobar' class='red-background'>" +isDeployOffer+ "</div>"; 
                          var infodiv = document.getElementById('infobar'); 
                          if(infodiv == null) 
                          { 
                              main.innerHTML = testdiv + main.innerHTML; 
                          }
                      }
                  }
              }
              if (action == "ProductCollisionsBackgroundUE") {
                  runbackground = self.xmlHttpReq.responseText.toString();
                  if(runbackground.replace("\r\n","") == "True"){
                      window.location("UEoffer-sum.aspx?OfferID=<%Sendb(OfferID)%>");
                }
                else {
                    var maindiv = document.getElementById('main'); 
                    var testdiv = "<div id='infobar' class='red-background'>" +runbackground+ "</div>"; 
                    var infodiv = document.getElementById('infobar'); 
                    if(infodiv == null) 
                    { 
                        main.innerHTML = testdiv + main.innerHTML; 
                    }
                }
            }
            if (action == "ProductCollisionsBackgroundUEDetection") {
                runbackground = self.xmlHttpReq.responseText.toString();
                if(runbackground.replace("\r\n","") == "True"){
                    window.location="UEoffer-sum.aspx?OfferID=<%Sendb(OfferID)%>";
          }
          else {
              var maindiv = document.getElementById('main'); 
              var testdiv = "<div id='infobar' class='red-background'>" +runbackground+ "</div>"; 
              var infodiv = document.getElementById('infobar'); 
              if(infodiv == null) 
              { 
                  main.innerHTML = testdiv + main.innerHTML; 
              }
          }
      }
      if (action == "GetProductCollisionsUEDetection") {
          updateCollisionsBoxDetection(self.xmlHttpReq.responseText);
      }
  }
    }
    self.xmlHttpReq.send();
}

function updateCollisionsBox(responseMsg) {
    var collisionsContent = document.getElementById("collisionsContent");

    if (responseMsg.replace(/(\r\n|\n|\r)/gm,"") != '') {
        collisionsContent.innerHTML = responseMsg;
    } else {
        if (DeferDeploy) {
            window.location = 'UEoffer-sum.aspx?OfferID=<%Sendb(OfferID)%>&deferdeploy=<%Sendb(Copient.PhraseLib.Lookup("term.deferdeployment", LanguageID))%>';
        } else {
            window.location = 'UEoffer-sum.aspx?OfferID=<%Sendb(OfferID)%>&deploy=<%Sendb(Copient.PhraseLib.Lookup("term.deploy", LanguageID))%>';
        }
    }
}

function showNoOfDuplicateOfferserror(content){
    var duplicateofferElem = document.getElementById("DuplicateOffererror");
    duplicateofferElem.style.display = 'block';
    duplicateofferElem.innerHTML = content;
}
  
function ClearNoOfDuplicateOfferserror(){
    var duplicateofferElem = document.getElementById("DuplicateOffererror");
    if (duplicateofferElem != null) {
        duplicateofferElem.style.display = 'none';
    }
}
  
function toggleDialog(elemName, shown) {
    var elem = document.getElementById(elemName);
    var fadeElem = document.getElementById('fadeDiv');
    if (elem != null) {
        elem.style.display = (shown) ? 'block' : 'none';
    }
    if (fadeElem != null) {
        fadeElem.style.display = (shown) ? 'block' : 'none';
    }
}
    function hideRejectConfirmation() {
        document.getElementById('rejectText').value = '';
        toggleDialog('oawreject', false);
    }
function redirtToOfferReportPage(OfferID)
{
    window.location.href='../CollidingOffers-Report.aspx?ID='+OfferID;
}


function assignNoofDuplicateOffers(shown) {
    var elem = document.getElementById('DuplicateNoofOffer');
    var fadeElem = document.getElementById('fadeDiv');
    if (elem != null) {
        elem.style.display = (shown) ? 'block' : 'none';
    }
    if (fadeElem != null) {
        fadeElem.style.display = (shown) ? 'block' : 'none';
    }
    if (shown)  {
        document.getElementById('txtDuplicateOffersCnt').value='1';
        document.getElementById('txtDuplicateOffersCnt').focus();
        ClearNoOfDuplicateOfferserror();
        return false;
    } else {
        return true;
    }
}    

function addDuplicateOfferscount() {
    var maxOffersperfolderduplicate = <% Sendb(MyCommon.NZ(MyCommon.Fetch_SystemOption(184), 0))%>;
    if (maxOffersperfolderduplicate == 0 ) {
        maxOffersperfolderduplicate = 99;
    }
    var dupOffersCntvalue =  document.getElementById("txtDuplicateOffersCnt").value;
    if (dupOffersCntvalue != null && dupOffersCntvalue != "") {
        if (!isNaN(dupOffersCntvalue)) {
            ClearNoOfDuplicateOfferserror();
            if (dupOffersCntvalue > maxOffersperfolderduplicate) {
                showNoOfDuplicateOfferserror('<%Sendb(Copient.PhraseLib.Lookup("folders.maxlimitreached", LanguageID))%>' + maxOffersperfolderduplicate + '. <%Sendb(Copient.PhraseLib.Lookup("foders.ActionCannotPerform", LanguageID))%>');
              }
              else {
                  xmlhttpPost_OfferDuplicateString('../OfferFeeds.aspx', 'Mode=NoOfDuplicateOffers&OfferID=' + document.mainform.OfferID.value + '&EngineID=9&DuplicateCnt=' + dupOffersCntvalue,'NoOfDuplicateOffers');
                  // hide the new popup
                 // toggleDialog('DuplicateNoofOffer',false);
              }
          } else {
              showNoOfDuplicateOfferserror('<%Sendb(Copient.PhraseLib.Lookup("term.invalidDuplicateOfferCount", LanguageID))%>');
          }
      }
      else {
          showNoOfDuplicateOfferserror('<%Sendb(Copient.PhraseLib.Lookup("term.enterDuplicateOfferCount", LanguageID))%>');
      }
  }
  
  function xmlhttpPost_OfferDuplicateString(strURL, qryStr, action) {
      var xmlHttpReq = false;
      var self = this;
      // Mozilla/Safari
      if (window.XMLHttpRequest) {
          self.xmlHttpReq = new XMLHttpRequest();
      }
          // IE
      else if (window.ActiveXObject) {
          self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
      }
      self.xmlHttpReq.open('POST', strURL + '?' + qryStr, true);
      self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
      self.xmlHttpReq.setRequestHeader("Content-length", qryStr.length);
      self.xmlHttpReq.setRequestHeader("Connection", "close");
      self.xmlHttpReq.onreadystatechange = function() {
          if (self.xmlHttpReq !=null && self.xmlHttpReq.readyState == 4) {
              if (action == 'NoOfDuplicateOffers') {
                  handleresponseDuplicateOffers(self.xmlHttpReq.responseText);
              }
          }
      }
      self.xmlHttpReq.send(qryStr);
      return false;
  }
  function handleresponseDuplicateOffers(responseText) {
      if (responseText.substring(0, 2) == 'OK') {
          window.location = '../offer-list.aspx';
      } else if(responseText.substring(0, 17) == 'Could not process'){
          // toggleDialog("DuplicateNoofOffer", true);
          var errMsg = responseText;
          showNoOfDuplicateOfferserror(errMsg);
      }
      else {
          window.location = 'UEOffer-gen.aspx?' + responseText;
      }
  }
    function showRejectConfirmation() {
        toggleDialog('oawreject', true);
    }
  $(document).ready(function() {
      var hasAnalytics = <%=m_AnalyticsCustomerGroups.CheckIfOfferHasAnalyticsCustomerGroup(OfferID).ToString().ToLower()%>;
      var canDeploy = <%= m_AnalyticsCustomerGroups.CanDeploy(OfferID).ToString().ToLower() %>;
      
      if(canDeploy) {     
          $('#deploy').removeAttr('disabled');
          $('#deferdeploy').removeAttr('disabled');
          $('#reqApproval').removeAttr('disabled');
          $('#reqApprovalWithDeployment').removeAttr('disabled');
          $('#reqApprovalWithDeferDeployment').removeAttr('disabled');
      }
      else {              
          $('#deploy').attr('disabled', 'disabled');
          $('#deferdeploy').attr('disabled', 'disabled');
          $('#reqApproval').attr('disabled', 'disabled');
          $('#reqApprovalWithDeployment').attr('disabled', 'disabled');
          $('#reqApprovalWithDeferDeployment').attr('disabled', 'disabled');
      }

      if(hasAnalytics) {          
          $('#getRecommendations').removeAttr('disabled');
          $('#getRecommendationsAndDeploy').removeAttr('disabled');
      } else {       
          $('#getRecommendations').attr('disabled', 'disabled');
          $('#getRecommendationsAndDeploy').attr('disabled', 'disabled');
      }
        
  });
    function checkLength(){
        elem = document.getElementById("rejectText");
        elem.onkeydown = function() {
            var key = event.keyCode || event.charCode;
            if( key == 8 || key == 46 ){
            }
            else{
                return check(elem);
            }
        };
        elem.onkeyup = function() {
            var key = event.keyCode || event.charCode;
            if( key == 8 || key == 46 ){
            }
            else{
                return check(elem);
            }
        };
        
    }
    function check(elem){
        if(elem.value.length >= 500)
                {
                    document.getElementById("rejectText").value = elem.value.substring(0, 500);
                    return false;
                }
                else
                    return true;

    }
</script>
<form id="mainform" name="mainform" action="#">
    <input type="hidden" name="OfferID" id="OfferID" value="<% Sendb(OfferID)%>" />
    <input type="hidden" name="IsTemplate" id="IsTemplate" value="<% Sendb(IsTemplateVal)%>" />
    <input type="hidden" name="CallingLocation" id="CallingLocation" value="0" />
    <div id="intro">
        <%
            Dim oName As String = ""
            If (MyCommon.Fetch_UE_SystemOption(168) = "1" AndAlso MyCommon.NZ(rst.Rows(0)("ExternalBuyerId"), "") <> "") Then
                oName = "Buyer " + rst.Rows(0).Item("ExternalBuyerId").ToString() + " - " + MyCommon.NZ(rst.Rows(0).Item("IncentiveName"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)).ToString()
            Else
                oName = MyCommon.NZ(rst.Rows(0).Item("IncentiveName"), Copient.PhraseLib.Lookup("term.unknown", LanguageID))
            End If
            If (IsTemplate) Then
                Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & ": " & MyCommon.TruncateString(oName, 43) & "</h1>")
            Else
                Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & ": " & MyCommon.TruncateString(oName, 43) & "</h1>")
            End If
            Send(vbTab & "<div id=""controls""" & IIf(Popup, " style=""display:none;""", "") & ">")
            
	  
            'If the user has permission for one of the action buttons, display the action button. 
            ShowActionButton = (Logix.UserRoles.CreateTemplate And Not IsTemplate) OrElse (Logix.UserRoles.CRUDOfferFromTemplate And IsTemplate) _
            OrElse ((Logix.UserRoles.DeleteOffer AndAlso Not IsTemplate) Or (Logix.UserRoles.DeleteTemplates AndAlso IsTemplate)) _
            OrElse (Logix.UserRoles.SendOffersToCRM And Not IsTemplate And ShowCRM) OrElse (MyCommon.Fetch_SystemOption(73) <> "") _
            OrElse ((Logix.UserRoles.DeployNonTemplateOffers OrElse Logix.UserRoles.DeployTemplateOffers) And Not IsTemplate) OrElse (Logix.UserRoles.CreateOfferFromBlank) OrElse (isCopyOfferEnabledForGeneralUser) _
            OrElse ((Logix.UserRoles.DeferDeployNonTemplateOffers OrElse Logix.UserRoles.DeferDeployTemplateOffers) And Not IsTemplate) _
            OrElse (Logix.UserRoles.EditFolders) _
            OrElse (((Logix.UserRoles.DeployNonTemplateOffers AndAlso Not FromTemplate) OrElse (Logix.UserRoles.DeployTemplateOffers AndAlso FromTemplate)) And Not IsTemplate) _
            OrElse (Logix.UserRoles.ExportOffer)

            If (Not LinksDisabled OrElse IsTemplate) Then
                If (ShowActionButton = True AndAlso (Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable OrElse isCopyOfferEnabledForGeneralUser)) Then
                    Logix.GetOfferStatus(OfferID, LanguageID, StatusCode)
                    If offerStatus > 10 Then
                        Send("<input type=""button"" class=""regular"" id=""actions"" name=""actions"" value=""" & Copient.PhraseLib.Lookup("term.actions", LanguageID) & " &#9660;"" onclick=""toggleDropdown();"" disabled=""disabled""/>")
                    Else
                        Send("<input type=""button"" class=""regular"" id=""actions"" name=""actions"" value=""" & Copient.PhraseLib.Lookup("term.actions", LanguageID) & " &#9660;"" onclick=""toggleDropdown();"" />")
				    End If
                    Send("<div class=""actionsmenu"" id=""actionsmenu"">")
                    StatusText = Logix.GetOfferStatus(OfferID, LanguageID, StatusCode)
                    
                    Dim isFolderActive As Boolean = True
                    If (isCopyOfferEnabledForGeneralUser) Then
                        Dim dtFolderDates As DataTable
                        Dim FolderEndDate As Date
                        
                    
                        dtFolderDates = GetFolderDetails(OfferID, MyCommon)
                        If dtFolderDates.Rows.Count > 0 Then
                            FolderEndDate = dtFolderDates.Rows(0).Item("EndDate")
                            If (FolderEndDate < Date.Now) Then isFolderActive = False
                        End If
                    End If
					If (bEnableAdditionalLockoutRestrictionsOnOffers AndAlso Not bOfferEditable AndAlso isCopyOfferEnabledForGeneralUser) Then
                        If Not bOfferEditable Then
                            If ((StatusText.Contains("Expired") AndAlso Not IsTemplate) OrElse (Not bOfferEditable AndAlso Not IsTemplate)) Then				
                                If (Not bOfferEditable AndAlso Not isFolderActive) Then
                                    Send_CopyExpiredOffer("onclick=""CopyExpiredOffer('Assigned folder is not Active and also in Lockout period.Please select another folder')""")
                                ElseIf (Not bOfferEditable) Then
                                    Send_CopyExpiredOffer("onclick=""CopyExpiredOffer('Assigned folder is in Lockout period.Please select another folder')""")
                                ElseIf (Not isFolderActive) Then
                                    Send_CopyExpiredOffer("onclick=""CopyExpiredOffer('Assigned folder is not active.Please select another folder')""")
                                Else
                                    Send_CopyOffer(IsTemplate)
                                End If
                            ElseIf (Logix.UserRoles.CopyTemplate And IsTemplate) Then
                                Send_CopyOffer(IsTemplate)
                                Send_OfferFromTemp
                            ElseIf (Not Logix.UserRoles.CopyTemplate And IsTemplate) Then
                                Send_OfferFromTemp()
                            ElseIf (Not Logix.UserRoles.CopyTemplate And Not IsTemplate) Then
                                Send_CopyOffer(IsTemplate)
                            Else
                                Send_CopyOffer(IsTemplate)
                            End If
                        ElseIf (Logix.UserRoles.CopyTemplate And IsTemplate) Then
                            Send_CopyOffer(IsTemplate)
                        End If
                     Else
						If (Logix.UserRoles.EditFolders) Then
							Send_AssignFolders(OfferID)
						End If
                        If (isCopyOfferEnabledForGeneralUser) Then
                            If (StatusText.Contains("Expired") AndAlso Not IsTemplate) Then
                                If (Not isFolderActive) Then
                                    Send_CopyExpiredOffer("onclick=""CopyExpiredOffer('Assigned folder is not active.Please select another folder')""")
                                Else
                                    Send_CopyOffer(IsTemplate)
                                End If
                            ElseIf (Logix.UserRoles.CopyTemplate And IsTemplate) Then
                                Send_CopyOffer(IsTemplate)
							ElseIf (Logix.UserRoles.CopyTemplate And Not IsTemplate) Then
                                Send_CopyOffer(IsTemplate)
                            ElseIf (Not Logix.UserRoles.CopyTemplate And Not IsTemplate) Then
                                Send_CopyOffer(IsTemplate)
                            End If
                        ElseIf (Logix.UserRoles.CreateOfferFromBlank) Then
							Send_CopyOffer(IsTemplate)
						End If
						If (m_isOAWEnabled And m_hasOfferModifiedAfterApproval And Not IsTemplate And m_requiresOfferApproval And ((Logix.UserRoles.DeployNonTemplateOffers AndAlso Not FromTemplate) OrElse (Logix.UserRoles.DeployTemplateOffers AndAlso FromTemplate))) Then
                            Dim isCollisionDeploy As Boolean = IsCollisionDetectionEnabled And RunCollision = 1 And Not Expired

                            If (Expired And MyCommon.Fetch_UE_SystemOption(80) = "1") Then
                            Else
                                If isCollisionDeploy Then
                                    Send_RequestApprovalCollision()
                                Else
                                    Send_RequestApproval()
                                End If
                            End If
                        End If
						If (MyCommon.Fetch_SystemOption(262) = "1") Then
                            If (((Logix.UserRoles.DeferDeployNonTemplateOffers AndAlso Not FromTemplate) OrElse (Logix.UserRoles.DeferDeployTemplateOffers AndAlso FromTemplate)) And Not IsTemplate) Then
                                Dim isCollisionDeferDeploy As Boolean = IsCollisionDetectionEnabled And RunCollision = 1 And Not Expired

                                If (Expired And MyCommon.Fetch_UE_SystemOption(80) = "1") Then
                                ElseIf (isCollisionDeferDeploy) Then
                                    If (m_isOAWEnabled And m_hasOfferModifiedAfterApproval And m_requiresOfferApproval) Then
                                        Send_RequestApprovalWithDeferDeploymentCollision()
                                    Else
                                        Send_DeferDeployCollision()
                                    End If
                                Else
                                    If (m_isOAWEnabled And m_hasOfferModifiedAfterApproval And m_requiresOfferApproval) Then
                                        Send_RequestApprovalWithDeferDeployment()
                                    Else
                                        Send_DeferDeploy()
                                    End If
                                End If
                            End If
                        Else
							If (((Logix.UserRoles.DeployNonTemplateOffers AndAlso Not FromTemplate) OrElse (Logix.UserRoles.DeployTemplateOffers AndAlso FromTemplate)) And Not IsTemplate) Then
                                Dim isCollisionDeferDeploy As Boolean = IsCollisionDetectionEnabled And RunCollision = 1 And Not Expired

                                If (Expired And MyCommon.Fetch_UE_SystemOption(80) = "1") Then
                                ElseIf (isCollisionDeferDeploy) Then
                                    If (m_isOAWEnabled And m_hasOfferModifiedAfterApproval And m_requiresOfferApproval) Then
                                        Send_RequestApprovalWithDeferDeploymentCollision()
                                    Else
                                        Send_DeferDeployCollision()
                                    End If
                                Else
                                    If (m_isOAWEnabled And m_hasOfferModifiedAfterApproval And m_requiresOfferApproval) Then
                                        Send_RequestApprovalWithDeferDeployment()
                                    Else
                                        Send_DeferDeploy()
                                    End If
                                End If
                            End If
                        End If
						If (((Logix.UserRoles.DeployNonTemplateOffers AndAlso Not FromTemplate) OrElse (Logix.UserRoles.DeployTemplateOffers AndAlso FromTemplate)) And Not IsTemplate) Then
                            Dim isCollisionDeploy As Boolean = IsCollisionDetectionEnabled And RunCollision = 1 And Not Expired

                            If (Expired And MyCommon.Fetch_UE_SystemOption(80) = "1") Then
                            ElseIf (isCollisionDeploy) Then
                                If (m_isOAWEnabled And m_hasOfferModifiedAfterApproval And m_requiresOfferApproval) Then
                                    Send_RequestApprovalWithDeploymentCollision()
                                Else
                                    Send_DeployCollision()
                                End If
                            Else
                                If (m_isOAWEnabled And m_hasOfferModifiedAfterApproval And m_requiresOfferApproval) Then
                                    Send_RequestApprovalWithDeployment()
                                Else
                                    Send_Deploy()
                                End If
                            End If
                        End If
						If (Logix.UserRoles.ExportOffer) Then
							Send_Export()
						End If
						If (MyCommon.NZ(rst.Rows(0).Item("ExportToEDW"), False) AndAlso MyCommon.Fetch_SystemOption(73).Trim <> "") Then
							Send("<input type=""submit"" id=""exportedw"" name=""exportedw"" value=""" & Copient.PhraseLib.Lookup("term.exporttoarchive", LanguageID) & """ />")
						End If
						If (Logix.UserRoles.CreateOfferFromBlank) Then
							Send_New()
						End If
						If (Logix.UserRoles.CRUDOfferFromTemplate And IsTemplate) Then
							Send_OfferFromTemp()
						End If
						If (Logix.UserRoles.CreateTemplate And Not IsTemplate) Then
							Send_Saveastemp()
						End If
						If (Logix.UserRoles.SendOffersToCRM And Not IsTemplate And ShowCRM) Then
							If (Expired And MyCommon.Fetch_UE_SystemOption(80) = "1") Then
							Else
								Send_SendOutbound()
							End If
						End If
						If (Logix.UserRoles.DeleteOffer AndAlso Not IsTemplate) Or (Logix.UserRoles.DeleteTemplates AndAlso IsTemplate) Then
							Send_Delete()
						End If

                        If m_AnalyticsCustomerGroups.CheckIfOfferHasAnalyticsCustomerGroup(OfferID) And Not IsTemplate Then
                            Send_GetRecommendations()
                            If Not m_isOAWEnabled And m_hasOfferModifiedAfterApproval Then
                                Send_GetRecommendationsAndDeploy()
                            End If
                        End If

						'adding condition for Run Collision
						If (((Logix.UserRoles.DeployNonTemplateOffers AndAlso Not FromTemplate) OrElse (Logix.UserRoles.DeployTemplateOffers AndAlso FromTemplate)) And Not IsTemplate) Then
							If (Expired And MyCommon.Fetch_UE_SystemOption(80) = "1") Then
							ElseIf (IsCollisionDetectionEnabled) Then
								Send_RunCollisionDetection()
							Else
								Send_DisabledRunCollisionDetection()
							End If
						End If
						'Adding Link to Report Page
						Dim reportListCount As Integer = -1
						Dim _reportListCount As AMSResult(Of Integer) = m_Offer.GetCollidingOfferCount(OfferID, AdminUserID, 9)
						If _reportListCount.Result = AMSResultType.Success Then
							reportListCount = _reportListCount.Result
						End If

						If (((Logix.UserRoles.DeployNonTemplateOffers AndAlso Not FromTemplate) OrElse (Logix.UserRoles.DeployTemplateOffers AndAlso FromTemplate)) And Not IsTemplate) Then
							If (Expired And MyCommon.Fetch_UE_SystemOption(80) = "1") Then
							ElseIf (IsCollisionDetectionEnabled And reportListCount > 0) Then
								Send_enableviewreport(OfferID)
							Else
								Send_disableviewreport(OfferID)
							End If
						End If
                   End If
                    Send("</div>")
                End If
            Else
                If (OfferLockedforCollisionDetection = True) Then
                    Send_CancelCollisionDetection()
                ElseIf (offerStatus > 12) Then
                    If m_hasOAWPermissions Then
                        Send("<input type=""button"" class=""regular"" id=""actions"" name=""actions"" value=""" & Copient.PhraseLib.Lookup("term.actions", LanguageID) & " &#9660;"" onclick=""toggleDropdown();"" />")
                        Send("<div class=""actionsmenu"" id=""actionsmenu"">")
                        Send_ApproveOffer()
                        Send_RejectOffer()
                        Send("</div>")
                    Else
                        Send_CancelGetApproval()
                    End If
                ElseIf (((Logix.UserRoles.DeployNonTemplateOffers OrElse Logix.UserRoles.DeployTemplateOffers) OrElse (Logix.UserRoles.DeferDeployNonTemplateOffers OrElse Logix.UserRoles.DeferDeployTemplateOffers)) And Not IsTemplate) Then
                    If offerStatus = 11 Or offerStatus = 12 Then
                        Send("<input type=""button"" class=""regular"" id=""actions"" name=""actions"" value=""" & Copient.PhraseLib.Lookup("term.actions", LanguageID) & " &#9660;"" onclick=""toggleDropdown();"" disabled=""disabled""/>")
                    Else
                        Send_CancelDeploy()
                    End If
                End If
            End If
            If MyCommon.Fetch_SystemOption(75) Then
				If (OfferID > 0 And Logix.UserRoles.AccessNotes AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
                    Send_NotesButton(3, OfferID, AdminUserID)
                End If
            End If
            Send(vbTab & "</div>")
        %>
    </div>
    <div id="main">
        <style>
            .horizontal {
             width: 280px;
             border-bottom: 1.5px solid cadetblue;
             padding-top: 6px;
            }
            .horizontal.green {
             border-bottom: 1.5px solid chartreuse;
            }
            .horizontal.greyed {
             border-bottom: 1.5px solid darkgrey;
            }

            .circle {
              width: 13px;
              height: 13px;
              border-radius: 50%;
              border:1.5px solid cadetblue;
            }
            .circle.filled {
              background: chartreuse;
            }
            .circle.greyed {
              border:1.5px solid darkgrey;
            }
            .circle.filled-greyed  {
                background: darkgrey;
                border:1.5px solid darkgrey;
            }
           
        </style>
        <% If m_isOAWEnabled And Not IsTemplate Then %>
        <div id="status-bar">
           
                <%
                    Dim isOfferAwaitingApproval As Boolean = m_OAWService.CheckIfOfferIsAwaitingApproval(OfferID).Result
                    Dim isOfferApproved As Boolean = (m_OAWService.GetOfferApprovalStatus(OfferID).Result = 2)
                    Dim disable As Boolean = False
                    Dim recExists As Boolean = (Not m_OAWService.GetOfferApprovalStatus(OfferID).Result = -1)
                    Dim reqApproval As Boolean = (m_OAWService.CheckIfUserRequiresOfferApproval(OfferID, AdminUserID).Result)
                    If Not reqApproval OrElse (reqApproval And Not recExists) Then disable = True
                    GenerateApprovalStatusBar(Logix, OfferID, isOfferAwaitingApproval, isOfferApproved, disable)
                %>
            <div id="text">
                <span><%=Copient.PhraseLib.Lookup("term.editoffer", LanguageID) %></span>
                <span style="margin-left:220px;"><%=Copient.PhraseLib.Lookup("term.waiting-approval", LanguageID) %></span>
                <span style="margin-left:200px;"><%=Copient.PhraseLib.Lookup("term.offer", LanguageID) & " " & Copient.PhraseLib.Lookup("term.approved", LanguageID) %></span>
            </div >
        </div> <br />
        <% End If %>
        <%
            If (Expired And MyCommon.Fetch_UE_SystemOption(80) = "1" And Not IsTemplate) Then
                LinksDisabled = True
            End If
            StatusText = Logix.GetOfferStatus(OfferID, LanguageID, StatusCode)
            If Not IsTemplate Then
                If (MyCommon.NZ(rst.Rows(0).Item("StatusFlag"), 0) <> 2) Then
                    If (StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_ACTIVE) AndAlso (MyCommon.NZ(rst.Rows(0).Item("StatusFlag"), 0) > 0) AndAlso (OfferLockedforCollisionDetection = False) AndAlso offerStatus <= 10 Then
                        If (MyCommon.NZ(rst.Rows(0).Item("DeployDeferred"), False) = False) Then
                            modMessage = Copient.PhraseLib.Lookup("alert.modpostdeploy", LanguageID)
                            Send("<div id=""modbar"">" & modMessage & "</div>")
                        End If
                    End If
                End If
            RejectMessage = m_OAWService.GetOfferRejectionMessage(OfferID).Result
            End If
            If (warnMessage <> String.Empty) Then
                Send("<div id=""infobar"" class=""orange-background"">" & warnMessage & "</div>")
            End If
            If (infoMessage <> "") Then
                Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
            End If
            If (RejectMessage <> String.Empty) Then
                Send("<div id=""infobar"" class=""red-background"" style="" word-wrap: break-word;"">" & rejectMessage & "<a href=""/logix/UE/UEoffer-sum.aspx?OfferID=" & OfferID & "&clearMessage=true"" style='color: white; text-decoration: none; float: right; padding-right: 5px;'>X</a></div>")
            End If
            ' Send the status bar, but not if it's a new offer or a template, or if there's already a modMessage being shown.
            If rst.Rows.Count < 1 Then
                GoTo done
            End If
            If (Not IsTemplate AndAlso modMessage = "") Then
                MyCommon.QueryStr = "select incentiveId from CPE_Incentives with (NoLock) where CreatedDate=LastUpdate and UpdateLevel=0 and IncentiveID=" & OfferID & ";"
                rst3 = MyCommon.LRT_Select
                If (rst3.Rows.Count = 0 OrElse OfferLockedforCollisionDetection = True) Then
                    Send_Status(OfferID, 2)
                End If
            End If

            ' send a message bar if this is an imported offer with a message set for feedback from the import process
            If OfferImported AndAlso ImportMessage <> "" Then
                Send("<div id=""modbar"" style=""background-color:#cc6600;"">" & ImportMessage & "</div>")
            End If

        %>
        <div id="column1">
            <div class="box" id="general">
                <%  
                    If (LinksDisabled) Then
                        Send("<h2 style=""float:left;""><span>" & Copient.PhraseLib.Lookup("term.general", LanguageID) & "</span></h2>")
                    Else
                        Send("<h2 style=""float:left;""><span><a href=""UEoffer-gen.aspx?OfferID=" & OfferID & """>" & Copient.PhraseLib.Lookup("term.general", LanguageID) & "</a></span></h2>")
                    End If
                    Send_BoxResizer("generalbody", "imgGeneral", Copient.PhraseLib.Lookup("term.general", LanguageID), True)
                    Send("<div id=""generalbody"">")
                    Send("  <table summary=""" & Copient.PhraseLib.Lookup("term.general", LanguageID) & """ cellpadding=""0"" cellspacing=""0"">")
                    Send("    <tr>")
                    Send("      <td style=""width:130px;""><b>" & Copient.PhraseLib.Lookup("term.id", LanguageID) & ":</b></td>")
                    Send("      <td>" & MyCommon.NZ(rst.Rows(0).Item("IncentiveID"), 0) & "</td>")
                    Send("    </tr>")
                    Send("    <tr>")
                    Send("      <td><b>" & Copient.PhraseLib.Lookup("term.externalid", LanguageID) & ":</b></td>")
                    Send("      <td>" & MyCommon.NZ(rst.Rows(0).Item("ClientOfferID"), "") & "</td>")
                    Send("    </tr>")
                    Send("    <tr>")
                    Send("      <td><b>" & Copient.PhraseLib.Lookup("term.roid", LanguageID) & ":</b></td>")
                    Send("      <td>" & roid & "</td>")
                    Send("    </tr>")
                    Send("    <tr>")
                    Send("      <td><b>" & Copient.PhraseLib.Lookup("term.engine", LanguageID) & ":</b></td>")
                    Sendb("      <td>")
                    Sendb(Copient.PhraseLib.Lookup(MyCommon.NZ(rst.Rows(0).Item("EnginePhraseID"), 0), LanguageID))
                    If MyCommon.NZ(rst.Rows(0).Item("EngineSubTypePhraseID"), 0) > 0 Then
                        Sendb(" " & Copient.PhraseLib.Lookup(MyCommon.NZ(rst.Rows(0).Item("EngineSubTypePhraseID"), 0), LanguageID))
                    End If
                    Send("</td>")
                    Send("    </tr>")
                    Send("    <tr>")
                    Send("      <td><b>" & Copient.PhraseLib.Lookup("term.vendor-coupon-code", LanguageID) & ":</b></td>")
                    Send("      <td>" & MyCommon.NZ(rst.Rows(0).Item("VendorCouponCode"), "") & "</td>")
                    Send("    </tr>")
                    If (MyCommon.Fetch_SystemOption(25) <> "0") Then
                        Send("<tr>")
                        Send("  <td>")
                        Send("    <b>" & Copient.PhraseLib.Lookup("term.outbound", LanguageID) & ":</b>")
                        Send("  </td>")
                        Send("  <td>")
                        MyCommon.QueryStr = "select ExtInterfaceID, Name, PhraseID from ExtCRMInterfaces with (NoLock) " &
                                            "where ExtInterfaceID=" & MyCommon.NZ(rst.Rows(0).Item("CRMEngineID"), 0)
                        rst4 = MyCommon.LRT_Select()
                        If rst4.Rows.Count > 0 Then
                            If Not IsDBNull(rst4.Rows(0).Item("PhraseID")) Then
                                Sendb(Copient.PhraseLib.Lookup(rst4.Rows(0).Item("PhraseID"), LanguageID))
                            Else
                                Sendb(MyCommon.NZ(rst4.Rows(0).Item("Name"), ""))
                            End If
                        End If
                        Send("  </td>")
                        Send("</tr>")
                    End If
                    Send("    <tr>")
                    Send("      <td><b>" & Copient.PhraseLib.Lookup("term.status", LanguageID) & ":</b></td>")
                    Send("      <td>" & Logix.GetOfferStatus(OfferID, LanguageID) & "</td>")
                    Send("    </tr>")
                    Send("    <tr>")
                    Send("      <td><b>" & Copient.PhraseLib.Lookup("term.name", LanguageID) & ":</b></td>")
                    Send("      <td>" & MyCommon.SplitNonSpacedString(oName, 25) & "</td>")
                    Send("    </tr>")
                    Send("    <tr>")
                    Send("      <td><b>" & Copient.PhraseLib.Lookup("term.description", LanguageID) & ":</b></td>")
                    Send("      <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(rst.Rows(0).Item("Description"), ""), 25) & "</td>")
                    Send("    </tr>")
                    Send("    <tr>")
                    Send("      <td><b>" & Copient.PhraseLib.Lookup("term.folders", LanguageID) & ":</b></td>")
                    Send("      <td id=""folderNames"">" & FolderNames & "</td>")
                    Send("    </tr>")
                    If MyCommon.Fetch_UE_SystemOption(206) = "1" Then
                        Send("    <tr>")
                        Send("      <td><b>" & Copient.PhraseLib.Lookup("term.category", LanguageID) & ":</b></td>")
                        'Send("      <td><a href=""javascript:openPopup('../offer-timeline.aspx?Category=" & MyCommon.NZ(rst.Rows(0).Item("OfferCategoryID"), "") & "')"">" & MyCommon.SplitNonSpacedString(MyCommon.NZ(rst.Rows(0).Item("CategoryName"), ""), 25) & "</a></td>")
                        If MyCommon.NZ(rst.Rows(0).Item("OfferCategoryID"), 0) > 0 Then
                            Send("      <td><a href=""../category-edit.aspx?OfferCategoryID=" & MyCommon.NZ(rst.Rows(0).Item("OfferCategoryID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(rst.Rows(0).Item("CategoryName"), ""), 25) & "</a></td>")
                        Else
                            Send("      <td>" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "</td>")
                        End If
                        Send("    </tr>")
                    End If
                    Send("    <tr>")
                    Send("      <td><b>" & Copient.PhraseLib.Lookup("term.priority", LanguageID) & ":</b></td>")
                    Sendb("      <td>")
                    MyCommon.QueryStr = "select PhraseID from UE_Priorities with (NoLock) where PriorityID=" & MyCommon.NZ(rst.Rows(0).Item("Priority"), 0) & ";"
                    rst4 = MyCommon.LRT_Select()
                    If rst4.Rows.Count > 0 Then
                        Sendb(Copient.PhraseLib.Lookup(rst4.Rows(0).Item("PhraseID"), LanguageID, MyCommon.NZ(rst.Rows(0).Item("Priority"), 0).ToString))
                    Else
                        Sendb(MyCommon.NZ(rst.Rows(0).Item("Priority"), 0))
                    End If
                    Send("</td>")
                    Send("    </tr>")
                    Send("    <tr>")
                    Send("      <td><b>" & Copient.PhraseLib.Lookup("term.currency", LanguageID) & ":</b></td>")
                    Send("      <td id=""currencyname"">" & Currency & "</td>")
                    Send("    </tr>")
                    Send("    <tr>")
                    Send("      <td><b>" & Copient.PhraseLib.Lookup("term.tiers", LanguageID) & ":</b></td>")
                    Send("      <td>" & TierLevels & "</td>")
                    Send("    </tr>")
                    Send("    <tr>")
                    Send("      <td><b>" & Copient.PhraseLib.Lookup("term.impression", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.reporting", LanguageID), VbStrConv.Lowercase) & ":</b></td>")
                    If (MyCommon.NZ(rst.Rows(0).Item("EnableImpressRpt"), False) = True) Then
                        If (Logix.UserRoles.AccessReports = True) AndAlso (Popup = False) Then
                            Send("      <td><a href=""../reports-detail.aspx?OfferID=" & MyCommon.NZ(rst.Rows(0).Item("IncentiveID"), 0) & "&amp;Start=" & MyCommon.NZ(rst.Rows(0).Item("Startdate"), "1/1/1900") & "&amp;End=" & MyCommon.NZ(rst.Rows(0).Item("Enddate"), "1/1/1900") & "&amp;Status=" & Copient.PhraseLib.Lookup("offer.status" & MyCommon.NZ(rst.Rows(0).Item("StatusFlag"), 0), LanguageID) & """>" & Copient.PhraseLib.Lookup("term.available", LanguageID) & "</a></td>")
                        Else
                            Send("      <td>" & Copient.PhraseLib.Lookup("term.enabled", LanguageID) & "</td>")
                        End If
                    Else
                        Send("      <td>" & Copient.PhraseLib.Lookup("term.disabled", LanguageID) & "</td>")
                    End If
                    Send("    </tr>")
                    Send("    <tr>")
                    Send("      <td><b>" & Copient.PhraseLib.Lookup("term.redemption", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.reporting", LanguageID), VbStrConv.Lowercase) & ":</b></td>")
                    If (MyCommon.NZ(rst.Rows(0).Item("EnableRedeemRpt"), False) = True) Then
                        If (Logix.UserRoles.AccessReports = True) AndAlso (Popup = False) Then
                            Send("      <td><a href=""../reports-detail.aspx?OfferID=" & MyCommon.NZ(rst.Rows(0).Item("IncentiveID"), 0) & "&amp;Start=" & MyCommon.NZ(rst.Rows(0).Item("Startdate"), "1/1/1900") & "&amp;End=" & MyCommon.NZ(rst.Rows(0).Item("Enddate"), "1/1/1900") & "&amp;Status=" & Copient.PhraseLib.Lookup("offer.status" & MyCommon.NZ(rst.Rows(0).Item("StatusFlag"), 0), LanguageID) & """>" & Copient.PhraseLib.Lookup("term.available", LanguageID) & "</a></td>")
                        Else
                            Send("      <td>" & Copient.PhraseLib.Lookup("term.enabled", LanguageID) & "</td>")
                        End If
                    Else
                        Send("      <td>" & Copient.PhraseLib.Lookup("term.disabled", LanguageID) & "</td>")
                    End If
                    Send("    </tr>")
                    Send("    <tr>")
                    Send("      <td><b>" & Copient.PhraseLib.Lookup("term.createdby", LanguageID) & ":</b></td>")
                    Send("      <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(rst.Rows(0).Item("CreatedBy"), ""), 25) & "</td>")
                    Send("    </tr>")
                    Send("    <tr>")
                    Send("      <td><b>" & Copient.PhraseLib.Lookup("term.buyer", LanguageID) & ":</b></td>")
                    Send("      <td>" & MyCommon.NZ(rst.Rows(0).Item("ExternalBuyerId"), "") & "</td>")
                    Send("    </tr>")
                    Send("    <tr>")
                    Send("      <td><b>" & Copient.PhraseLib.Lookup("term.lastupdatedby", LanguageID) & ":</b></td>")
                    Send("      <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(rst.Rows(0).Item("LastUpdatedBy"), ""), 25) & "</td>")
                    Send("    </tr>")
                    Send("  </table>")
                    Send("</div>")
                %>
            </div>
            <div class="box" id="MEG">
                <h2 style="float: left;">
                    <span>
                        <% Sendb(Copient.PhraseLib.Lookup("term.MutualExclusionGroups", LanguageID))%>
                    </span>
                </h2>
                <% Send_BoxResizer("MEGbody", "imgMEG", Copient.PhraseLib.Lookup("term.MutualExclusionGroups", LanguageID), True)%>
                <div id="MEGbody">
                    <%
                        MyCommon.QueryStr = "select MutualExclusionGroupID, Name from MutualExclusionGroups with (NoLock) where MutualExclusionGroupID in " &
                                            "  (select MutualExclusionGroupID from MutualExclusionGroupoffers where OfferID=" & OfferID & ") " &
                                            "order by Name;"
                        rst4 = MyCommon.LRT_Select()
                        If rst4.Rows.Count > 0 Then
                            Send("        <ul>")
                            For Each row In rst4.Rows
                                Send("          <li>" & MyCommon.NZ(row.Item("Name"), "") & "</li>")
                            Next
                            Send("        </ul>")
                        Else
                            Send("        " & Copient.PhraseLib.Lookup("term.none", LanguageID))
                        End If
                    %>
                </div>
            </div>
            <div class="box" id="period">
                <h2 style="float: left;">
                    <span>
                        <% Sendb(Copient.PhraseLib.Lookup("term.period", LanguageID))%>
                    </span>
                </h2>
                <% Send_BoxResizer("periodbody", "imgPeriod", Copient.PhraseLib.Lookup("term.period", LanguageID), True)%>
                <div id="periodbody">
                    <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.period", LanguageID))%>"
                        cellpadding="0" cellspacing="0">
                        <tr>
                            <td style="width: 85px;">
                                <b>
                                    <% Sendb(Copient.PhraseLib.Lookup("term.production", LanguageID))%>
                :</b>
                            </td>
                            <td>
                                <%
                                    Dim startTime As DateTime
                                    Dim endTime As DateTime
                                    startTime = MyCommon.NZ(rst.Rows(0).Item("StartDate"), "1/1/1900")
                                    If startTime > "1/1/1900" Then Sendb(IIf(isStartEndTimeEnabled, startTime.ToString(), Logix.ToShortDateString(startTime, MyCommon))) Else Sendb("?")
                                    Sendb(" - ")
                                    endTime = MyCommon.NZ(rst.Rows(0).Item("EndDate"), "1/1/1900")
                                    If endTime > "1/1/1900" Then Sendb(IIf(isStartEndTimeEnabled, endTime.ToString(), Logix.ToShortDateString(endTime, MyCommon))) Else Sendb("?")
                                    Dim span As TimeSpan = endTime.Subtract(startTime)
                                    Sendb(" (")
                                    If Not isStartEndTimeEnabled Then
                                        Sendb(span.Days + 1 & " " & Copient.PhraseLib.Lookup("term.day(s)", LanguageID))
                                    Else
                                        If (span.Days = 0 AndAlso span.Hours = 0 AndAlso span.Minutes = 0 AndAlso span.Seconds = 0) Then
                                            Sendb("0 " & Copient.PhraseLib.Lookup("term.hour(s)", LanguageID))
                                        Else
                                            If (span.Days > 0) Then Sendb(span.Days & " " & Copient.PhraseLib.Lookup("term.day(s)", LanguageID) & IIf(span.Hours > 0, " ", ""))
                                        If (span.Hours > 0) Then Sendb(span.Hours & " " & Copient.PhraseLib.Lookup("term.hour(s)", LanguageID) & IIf(span.Minutes > 0, " ", ""))
                                        If (span.Minutes > 0) Then Sendb(span.Minutes & " " & Copient.PhraseLib.Lookup("term.minute(s)", LanguageID) & IIf(span.Seconds > 0, " ", ""))
                                        If (span.Seconds > 0) Then Sendb(span.Seconds & " " & Copient.PhraseLib.Lookup("term.second(s)", LanguageID))
                                        End If
                                    End If
                                    Sendb(")")
                                %>
                            </td>
                        </tr>
                        <% 
                            'BZ2079: UE-feature-removal - removing the eligibility start/enddates
                            'To undo this change, remove the style="display: none;" from the following line
                        %>
                       <%-- <tr style="display: none;">
                            <td>
                                <b>
                                    <% Sendb(Copient.PhraseLib.Lookup("term.eligibility", LanguageID))%>
                :</b>
                            </td>
                            <td>
                                <%
                                    LongDate = MyCommon.NZ(rst.Rows(0).Item("EligibilityStartDate"), "1/1/1900")
                                    If LongDate > "1/1/1900" Then Sendb(Logix.ToShortDateString(LongDate, MyCommon)) Else Sendb("?")
                                    Sendb(" - ")
                                    LongDate = MyCommon.NZ(rst.Rows(0).Item("EligibilityEndDate"), "1/1/1900")
                                    If LongDate > "1/1/1900" Then Sendb(Logix.ToShortDateString(LongDate, MyCommon)) Else Sendb("?")
                                    Sendb(" (" & DateDiff(DateInterval.Day, MyCommon.NZ(rst.Rows(0).Item("EligibilityStartDate"), "1/1/1900"), MyCommon.NZ(rst.Rows(0).Item("EligibilityEndDate"), "1/1/1900")) + 1 & " ")
                                    If DateDiff(DateInterval.Day, MyCommon.NZ(rst.Rows(0).Item("EligibilityStartDate"), "1/1/1900"), MyCommon.NZ(rst.Rows(0).Item("EligibilityEndDate"), "1/1/1900")) + 1 = 1 Then
                                        Sendb(StrConv(Copient.PhraseLib.Lookup("term.day", LanguageID), VbStrConv.Lowercase) & ")")
                                    Else
                                        Sendb(StrConv(Copient.PhraseLib.Lookup("term.days", LanguageID), VbStrConv.Lowercase) & ")")
                                    End If
                                %>
                            </td>
                        </tr>--%>
                        <tr>
                            <td>
                                <b>
                                    <% Sendb(Copient.PhraseLib.Lookup("term.testing", LanguageID))%>
                :</b>
                            </td>
                            <td>
                                <%
                                    Dim testingStartTime As DateTime
                                    Dim testingEndTime As DateTime
                                    testingStartTime = MyCommon.NZ(rst.Rows(0).Item("TestingStartDate"), "1/1/1900")
                                    If testingStartTime > "1/1/1900" Then Sendb(IIf(isStartEndTimeEnabled, testingStartTime.ToString(), Logix.ToShortDateString(testingStartTime, MyCommon))) Else Sendb("?")
                                    Sendb(" - ")
                                    testingEndTime = MyCommon.NZ(rst.Rows(0).Item("TestingEndDate"), "1/1/1900")
                                    If testingEndTime > "1/1/1900" Then Sendb(IIf(isStartEndTimeEnabled, testingEndTime.ToString(), Logix.ToShortDateString(testingEndTime, MyCommon))) Else Sendb("?")
                                    span = testingEndTime.Subtract(testingStartTime)
                                    Sendb(" (")
                                    If Not isStartEndTimeEnabled Then
                                        Sendb(span.Days + 1 & " " & Copient.PhraseLib.Lookup("term.day(s)", LanguageID))
                                    Else
                                        If (span.Days = 0 AndAlso span.Hours = 0 AndAlso span.Minutes = 0 AndAlso span.Seconds = 0) Then
                                            Sendb("0 " & Copient.PhraseLib.Lookup("term.hour(s)", LanguageID))
                                        Else
                                            If (span.Days > 0) Then Sendb(span.Days & " " & Copient.PhraseLib.Lookup("term.day(s)", LanguageID) & IIf(span.Hours > 0, " ", ""))
                                        If (span.Hours > 0) Then Sendb(span.Hours & " " & Copient.PhraseLib.Lookup("term.hour(s)", LanguageID) & IIf(span.Minutes > 0, " ", ""))
                                        If (span.Minutes > 0) Then Sendb(span.Minutes & " " & Copient.PhraseLib.Lookup("term.minute(s)", LanguageID) & IIf(span.Seconds > 0, " ", ""))
                                        If (span.Seconds > 0) Then Sendb(span.Seconds & " " & Copient.PhraseLib.Lookup("term.second(s)", LanguageID))
                                        End If
                                    End If
                                    Sendb(")")

                                %>
                            </td>
                        </tr>
                    </table>
                </div>
            </div>
            <% If bUseDisplayDates Then%>
            <div class="box" id="displaydates">
                <h2 style="float: left;">
                    <span>
                        <% Sendb(Copient.PhraseLib.Lookup("term.displaydates", LanguageID))%>
                    </span>
                </h2>
                <% Send_BoxResizer("displaybody", "imgDisplay", Copient.PhraseLib.Lookup("term.display", LanguageID), True)%>
                <div id="displaybody">
                    <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.display", LanguageID))%>"
                        cellpadding="0" cellspacing="0">
                        <%  MyCommon.QueryStr = "SELECT DisplayStartDate,DisplayEndDate FROM OfferAccessoryFields with (NoLock) WHERE OfferID = " & OfferID & " "
                            Dim dtODisp As New DataTable
                            dtODisp = MyCommon.LRT_Select()
                            If dtODisp.Rows.Count > 0 Then
                                If Not IsDBNull(dtODisp.Rows(0).Item("DisplayStartDate")) And Not IsDBNull(dtODisp.Rows(0).Item("DisplayEndDate")) Then%>
                        <tr>
                            <td style="width: 85px; vertical-align: top;">
                                <b>
                                    <% Sendb(Copient.PhraseLib.Lookup("term.display", LanguageID))%>
                :</b>
                            </td>
                            <td>
                                <%
                                    LongDate = MyCommon.NZ(dtODisp.Rows(0).Item("DisplayStartDate"), "1/1/1900")
                                    If LongDate > "1/1/1900" Then Sendb(Logix.ToShortDateString(LongDate, MyCommon)) Else Sendb("?")
                                    Sendb(" - ")
                                    LongDate = MyCommon.NZ(dtODisp.Rows(0).Item("DisplayEndDate"), "1/1/1900")
                                    If LongDate > "1/1/1900" Then Sendb(Logix.ToShortDateString(LongDate, MyCommon)) Else Sendb("?")
                                    Sendb(" (" & DateDiff(DateInterval.Day, MyCommon.NZ(dtODisp.Rows(0).Item("DisplayStartDate"), "1/1/1900"), MyCommon.NZ(dtODisp.Rows(0).Item("DisplayEndDate"), "1/1/1900")) + 1 & " ")
                                    If DateDiff(DateInterval.Day, MyCommon.NZ(dtODisp.Rows(0).Item("DisplayStartDate"), "1/1/1900"), MyCommon.NZ(dtODisp.Rows(0).Item("DisplayEndDate"), "1/1/1900")) + 1 = 1 Then
                                        Sendb(StrConv(Copient.PhraseLib.Lookup("term.day", LanguageID), VbStrConv.Lowercase) & ")")
                                    Else
                                        Sendb(StrConv(Copient.PhraseLib.Lookup("term.days", LanguageID), VbStrConv.Lowercase) & ")")
                                    End If
                                    Sendb("<br />")
                                    LongDate = MyCommon.NZ(dtODisp.Rows(0).Item("DisplayStartDate"), "1/1/1900")
                                    If LongDate > "1/1/1900" Then Sendb(LongDate.ToString("HH:mm:ss")) Else Sendb("?")
                                    Sendb(" - ")
                                    LongDate = MyCommon.NZ(dtODisp.Rows(0).Item("DisplayEndDate"), "1/1/1900")
                                    If LongDate > "1/1/1900" Then Sendb(LongDate.ToString("HH:mm:ss")) Else Sendb("?")
                                    Sendb("<br />")

                                    Dim DispStartDateDiff As Integer = DateDiff(DateInterval.Day, Today, MyCommon.NZ(dtODisp.Rows(0).Item("DisplayStartDate"), "1/1/1900"))
                                    Dim DispEndDateDiff As Integer = DateDiff(DateInterval.Day, Today, MyCommon.NZ(dtODisp.Rows(0).Item("DisplayEndDate"), "1/1/1900"))
                                    If DispStartDateDiff > 1 Then
                                        Sendb(Copient.PhraseLib.Detokenize("offer-sum.BeginsIn", LanguageID, DispStartDateDiff) & " ")
                                    ElseIf DispStartDateDiff = 1 Then
                                        Sendb(Copient.PhraseLib.Lookup("offer-sum.BeginsTomorrow", LanguageID) & " ")
                                    ElseIf DispStartDateDiff = 0 Then
                                        Sendb(Copient.PhraseLib.Lookup("offer-sum.BeganToday", LanguageID) & " ")
                                    ElseIf DispStartDateDiff = -1 Then
                                        Sendb(Copient.PhraseLib.Lookup("offer-sum.BeganYesterday", LanguageID) & " ")
                                    Else
                                        Sendb(Copient.PhraseLib.Detokenize("offer-sum.BeganDaysAgo", LanguageID, Math.Abs(DispStartDateDiff)) & " ")
                                    End If
                                    If DispEndDateDiff > 1 Then
                                        Sendb(Copient.PhraseLib.Detokenize("offer-sum.EndsIn", LanguageID, DispEndDateDiff))
                                    ElseIf DispEndDateDiff = 1 Then
                                        Sendb(Copient.PhraseLib.Lookup("offer-sum.EndsTomorrow", LanguageID))
                                    ElseIf DispEndDateDiff = 0 Then
                                        Sendb(Copient.PhraseLib.Lookup("offer-sum.EndsToday", LanguageID))
                                    ElseIf DispEndDateDiff = -1 Then
                                        Sendb(Copient.PhraseLib.Lookup("offer-sum.EndedYesterday", LanguageID))
                                    Else
                                        Sendb(Copient.PhraseLib.Detokenize("offer-sum.EndedDaysAgo", LanguageID, Math.Abs(DispEndDateDiff)))
                                    End If

                                %>
                            </td>
                        </tr>
                        <% Else%>
                        <tr>
                            <td style="width: 85px; vertical-align: top;">
                                <b>
                                    <% Sendb(Copient.PhraseLib.Lookup("term.none", LanguageID))%>
                                </b>
                            </td>
                            <td>&nbsp;
                            </td>
                        </tr>
                        <%  End If
                            Else%>
                        <tr>
                            <td style="width: 85px; vertical-align: top;">
                                <b>
                                    <% Sendb(Copient.PhraseLib.Lookup("term.none", LanguageID))%>
                                </b>
                            </td>
                            <td>&nbsp;
                            </td>
                        </tr>
                        <%  End If%>
                    </table>
                </div>
            </div>
            <% End If%>
            <div class="box" id="limits">
                <h2 style="float: left;">
                    <span>
                        <% Sendb(Copient.PhraseLib.Lookup("term.limits", LanguageID))%>
                    </span>
                </h2>
                <% Send_BoxResizer("limitsbody", "imgLimits", Copient.PhraseLib.Lookup("term.limits", LanguageID), True)%>
                <div id="limitsbody">
                    <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.period", LanguageID))%>"
                        cellpadding="0" cellspacing="0">
                        <%
'BZ2079: UE-feature-removal - removing the eligibility limits
'To undo this change, remove the style="display: none;" from the following line
                        %>
                        <tr style="display: none;">
                            <td style="width: 85px;">
                                <b>
                                    <% Sendb(Copient.PhraseLib.Lookup("term.eligibility", LanguageID))%>
                :</b>
                            </td>
                            <td>
                                <%
                                    If (MyCommon.NZ(rst.Rows(0).Item("P1DistPeriod"), 0) = 0) Then
                                        Sendb(Copient.PhraseLib.Lookup("term.unlimited", LanguageID))
                                    ElseIf (MyCommon.NZ(rst.Rows(0).Item("P1DistQtyLimit"), 0) = 1 AndAlso MyCommon.NZ(rst.Rows(0).Item("P1DistPeriod"), 0) = 3650) Then
                                        Sendb(Copient.PhraseLib.Lookup("term.onceperoffer", LanguageID))
                                    ElseIf (MyCommon.NZ(rst.Rows(0).Item("P1DistQtyLimit"), 0) = 1 AndAlso MyCommon.NZ(rst.Rows(0).Item("P1DistPeriod"), 0) = 1 AndAlso (MyCommon.NZ(rst.Rows(0).Item("P1DistTimeType"), 0) = 2)) Then
                                        Sendb(Copient.PhraseLib.Lookup("term.oncepertransaction", LanguageID))
                                    ElseIf (MyCommon.NZ(rst.Rows(0).Item("P1DistPeriod"), 0) = 1 AndAlso MyCommon.NZ(rst.Rows(0).Item("P1DistTimeType"), 0) = 2) Then
                                        Sendb(MyCommon.NZ(rst.Rows(0).Item("P1DistQtyLimit"), 0))
                                        Sendb(" " & Copient.PhraseLib.Lookup("term.per", LanguageID) & " ")
                                        Sendb(StrConv(Copient.PhraseLib.Lookup("term.transaction", LanguageID), VbStrConv.Lowercase))
                                    Else
                                        Sendb(MyCommon.NZ(rst.Rows(0).Item("P1DistQtyLimit"), 0))
                                        Sendb(" " & Copient.PhraseLib.Lookup("term.per", LanguageID) & " ")
                                        Send(MyCommon.NZ(rst.Rows(0).Item("P1DistPeriod"), 0))
                                        MyCommon.QueryStr = "Select Description, PhraseID from CPE_DistributionTimeTypes with (NoLock) where TimeTypeId = " & MyCommon.NZ(rst.Rows(0).Item("P1DistTimeType"), 1)
                                        rst2 = MyCommon.LRT_Select
                                        If rst2.Rows.Count > 0 Then Sendb(" " & StrConv(Copient.PhraseLib.Lookup(rst2.Rows(0).Item("PhraseID"), LanguageID), VbStrConv.Lowercase))
                                    End If
                                %>
                            </td>
                        </tr>
                        <%-- Commented out 8/6 per Vince's advice. ALM
            <tr>
              <td><b><% Sendb(Copient.PhraseLib.Lookup("term.accumulation", LanguageID))%>:</b></td>
              <td><%
              If (MyCommon.NZ(rst.Rows(0).Item("P2DistPeriod"), 0) = 3650) Then
                Send(Copient.PhraseLib.Lookup("term.onceperoffer", LanguageID))
              Else
                Send(MyCommon.NZ(rst.Rows(0).Item("P2DistQtyLimit"), 0) & " " & Copient.PhraseLib.Lookup("term.per", LanguageID) & " ")
                Send(MyCommon.NZ(rst.Rows(0).Item("P2DistPeriod"), 0) & " ")
                MyCommon.QueryStr = "Select Description from CPE_DistributionTimeTypes with (NoLock) where TimeTypeId = " & MyCommon.NZ(rst.Rows(0).Item("P2DistTimeType"), 1)
                rst2 = MyCommon.LRT_Select
                Send(MyCommon.NZ(rst2.Rows(0).Item("Description"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)))
              End If
              %></td>
            </tr>
                        --%>
                        <tr>
                            <td>
                                <b>
                                    <% Sendb(Copient.PhraseLib.Lookup("term.reward", LanguageID))%>
                :</b>
                            </td>
                            <td>
                                <%
                                    If (MyCommon.NZ(rst.Rows(0).Item("P3DistQtyLimit"), 0) = 0 AndAlso MyCommon.NZ(rst.Rows(0).Item("P3DistTimeType"), 0) = 2 AndAlso MyCommon.NZ(rst.Rows(0).Item("P3DistPeriod"), 0) = 0) Then
                                        Send(Copient.PhraseLib.Lookup("term.unlimited", LanguageID))
                                    ElseIf (MyCommon.NZ(rst.Rows(0).Item("P3DistTimeType"), 0) = 1 AndAlso MyCommon.NZ(rst.Rows(0).Item("P3DistQtyLimit"), 0) = 1 AndAlso MyCommon.NZ(rst.Rows(0).Item("P3DistPeriod"), 0) = 3650) Then
                                        Send(Copient.PhraseLib.Lookup("term.onceperoffer", LanguageID))
                                    ElseIf (MyCommon.NZ(rst.Rows(0).Item("P3DistQtyLimit"), 0) = 1 AndAlso (MyCommon.NZ(rst.Rows(0).Item("P3DistTimeType"), 0) = 2)) Then
                                        Send(Copient.PhraseLib.Lookup("term.oncepertransaction", LanguageID))
                                    ElseIf (MyCommon.NZ(rst.Rows(0).Item("P3DistTimeType"), 0) = 2) Then
                                        Sendb(MyCommon.NZ(rst.Rows(0).Item("P3DistQtyLimit"), 0))
                                        Sendb(" " & Copient.PhraseLib.Lookup("term.per", LanguageID) & " ")
                                        Sendb(StrConv(Copient.PhraseLib.Lookup("term.transaction", LanguageID), VbStrConv.Lowercase))
                                    Else
                                        Sendb(MyCommon.NZ(rst.Rows(0).Item("P3DistQtyLimit"), 0))
                                        Sendb(" " & Copient.PhraseLib.Lookup("term.per", LanguageID) & " ")
                                        Send(MyCommon.NZ(rst.Rows(0).Item("P3DistPeriod"), 0))
                                        MyCommon.QueryStr = "Select Description, PhraseID from CPE_DistributionTimeTypes with (NoLock) where TimeTypeId = " & MyCommon.NZ(rst.Rows(0).Item("P3DistTimeType"), 1)
                                        rst2 = MyCommon.LRT_Select
                                        Sendb(" " & StrConv(Copient.PhraseLib.Lookup(rst2.Rows(0).Item("PhraseID"), LanguageID), VbStrConv.Lowercase))
                                    End If
                                %>
                            </td>
                        </tr>
                    </table>
                </div>
            </div>
            <% If (Not IsTemplate) Then%>
            <%--      <div class="box" id="validation">
        <h2 style="float: left;">
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.validationreport", LanguageID))%>
          </span>
        </h2>
        <%  Send_BoxResizer("validationbody", "imgValidation", Copient.PhraseLib.Lookup("term.validationreport", LanguageID), True)%>
        <div id="validationbody">
          <%
            Dim dtValid, dtComponents As DataTable
            Dim rowOK(), rowWatches(), rowWarnings() As DataRow
            Dim rowComp As DataRow
            Dim objTemp As Object
            Dim GraceHours As Integer
            Dim GraceCount As Double
            Dim AllComponentsValid As Boolean = True
            
            objTemp = MyCommon.Fetch_UE_SystemOption(41)
            If Not (Integer.TryParse(objTemp.ToString, GraceHours)) Then
              GraceHours = 4
            End If
                    
            objTemp = MyCommon.Fetch_UE_SystemOption(42)
            If Not (Double.TryParse(objTemp.ToString, GraceCount)) Then
              GraceCount = 0.1D
            End If
            
            MyCommon.QueryStr = "dbo.pa_ValidationReport_Incentive"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@IncentiveID", SqlDbType.Int).Value = OfferID
            MyCommon.LRTsp.Parameters.Add("@GraceHours", SqlDbType.Int).Value = GraceHours
            MyCommon.LRTsp.Parameters.Add("@GraceCount", SqlDbType.Decimal, 2).Value = GraceCount
            
            dtValid = MyCommon.LRTsp_select()
            
            RemoveInactiveLocations(MyCommon, dtValid, OfferID)
            
            rowOK = dtValid.Select("Status=0", "LocationName")
            rowWatches = dtValid.Select("Status=1", "LocationName")
            rowWarnings = dtValid.Select("Status=2", "LocationName")
            MyCommon.Close_LRTsp()
            
            ValidateIncentiveColor = IIf(rowWarnings.Length > 0, "red", "green")
            
            Send("<a href=""javascript:showDiv('divOffer');"" style=""color:" & ValidateIncentiveColor & ";""><b>+ " & Copient.PhraseLib.Lookup("term.offer", LanguageID) & "</b></a><br />")
            Send("<div id=""divOffer"" style=""margin-left:10px;display:none;"">")
            If Popup Then
              Send(Copient.PhraseLib.Lookup("cgroup-edit.validlocations", LanguageID) & " (" & rowOK.Length & ")<br />")
              Send(Copient.PhraseLib.Lookup("cgroup-edit.watchlocations", LanguageID) & " (" & rowWatches.Length & ")<br />")
              Send(Copient.PhraseLib.Lookup("cgroup-edit.warninglocations", LanguageID) & " (" & rowWarnings.Length & ")<br />")
            Else
              Send("<a id=""validLink" & OfferID & """ href=""javascript:openPopup('../validation-report.aspx?type=in&id=" & OfferID & "&level=0');"">")
              Send("  " & Copient.PhraseLib.Lookup("cgroup-edit.validlocations", LanguageID) & " (" & rowOK.Length & ")")
              Send("</a><br />")
              Send("<a id=""watchLink" & OfferID & """ href=""javascript:openPopup('../validation-report.aspx?type=in&id=" & OfferID & "&level=1');"">")
              Send("  " & Copient.PhraseLib.Lookup("cgroup-edit.watchlocations", LanguageID) & " (" & rowWatches.Length & ")")
              Send("</a><br />")
              Send("<a id=""warningLink" & OfferID & """ href=""javascript:openPopup('../validation-report.aspx?type=in&id=" & OfferID & "&level=2');"">")
              Send("  " & Copient.PhraseLib.Lookup("cgroup-edit.warninglocations", LanguageID) & " (" & rowWarnings.Length & ")")
              Send("</a><br />")
            End If
            Send("</div>")
            
            MyCommon.QueryStr = "dbo.pa_ValidationReport_OfferComponents"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.Int).Value = OfferID
            dtComponents = MyCommon.LRTsp_select
            
            Send("<a id=""linkComponent"" href=""javascript:showDiv('divComponents');"" style=""color:green;""><br class=""half"" /><b>+ " & Copient.PhraseLib.Lookup("term.components", LanguageID) & "</b></a><br />")
            Send("<div id=""divComponents"" style=""display:none;"">")
            For Each rowComp In dtComponents.Rows
              Send("<div style=""margin-left:10px;"">")
              WriteComponent(MyCommon, rowComp, ComponentColor, Popup)
              AllComponentsValid = AllComponentsValid AndAlso (ComponentColor.ToUpper = "GREEN")
              Send("</div>")
            Next
            Send("</div>")
            
            ' Update the Offer Validation Summary table with the most current validation information
            MyCommon.QueryStr = "dbo.pa_UpdateValidationSummary"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
            MyCommon.LRTsp.Parameters.Add("@ValidLocations", SqlDbType.Int).Value = rowOK.Length
            MyCommon.LRTsp.Parameters.Add("@WatchLocations", SqlDbType.Int).Value = rowWatches.Length
            MyCommon.LRTsp.Parameters.Add("@WarningLocations", SqlDbType.Int).Value = rowWarnings.Length
            MyCommon.LRTsp.Parameters.Add("@ComponentsValid", SqlDbType.Bit).Value = IIf(AllComponentsValid, 1, 0)
            MyCommon.LRTsp.ExecuteNonQuery()
          %>
        </div>
        <hr class="hidden" />
      </div>
            --%>
            <% End If%>
            <% If (Not IsTemplate) Then%>
            <div class="box" id="deployment">
                <h2 style="float: left;">
                    <span>
                        <% Sendb(Copient.PhraseLib.Lookup("term.deployment", LanguageID))%>
                    </span>
                </h2>
                <% Send_BoxResizer("deploymentbody", "imgDeployment", Copient.PhraseLib.Lookup("term.deployment", LanguageID), True)%>
                <div id="deploymentbody">
                    <h3>
                        <%Sendb(Copient.PhraseLib.Lookup("term.lastattempted", LanguageID))%>
          :
                    </h3>
                    <%
                        TodayDateZeroTime = New Date(DateTime.Today.Year, DateTime.Today.Month, DateTime.Today.Day, 0, 0, 0)
                        LongDate = MyCommon.NZ(rst.Rows(0).Item("CPEOARptDate"), "1/1/1900")
                        LongDateZeroTime = New Date(LongDate.Year, LongDate.Month, LongDate.Day, 0, 0, 0)
                        If LongDate > "1/1/1900" Then
                            DaysDiff = DateDiff(DateInterval.Day, LongDateZeroTime, TodayDateZeroTime)
                            Sendb("&nbsp;" & Logix.ToShortDateTimeString(LongDate, MyCommon))
                            If DaysDiff < 0 Then
                            ElseIf DaysDiff = 0 Then
                                Sendb(" (" & StrConv(Copient.PhraseLib.Lookup("term.today", LanguageID), VbStrConv.Lowercase) & ")")
                            ElseIf DaysDiff = 1 Then
                                Sendb(" (" & StrConv(Copient.PhraseLib.Lookup("term.yesterday", LanguageID), VbStrConv.Lowercase) & ")")
                            Else
                                Sendb(" (" & DaysDiff & " " & StrConv(Copient.PhraseLib.Lookup("term.daysago", LanguageID), VbStrConv.Lowercase) & ")")
                            End If
                        Else
                            Sendb(Copient.PhraseLib.Lookup("term.never", LanguageID))
                        End If
                        Sendb("<br />")
                    %>
                    <br class="half" />
                    <h3>
                        <%Sendb(Copient.PhraseLib.Lookup("term.lastsuccessful", LanguageID))%>
          :
                    </h3>
                    <%
                        LongDate = MyCommon.NZ(rst.Rows(0).Item("CPEOADeploySuccessDate"), "1/1/1900")
                        LongDateZeroTime = New Date(LongDate.Year, LongDate.Month, LongDate.Day, 0, 0, 0)
                        If LongDate > "1/1/1900" Then
                            DaysDiff = DateDiff(DateInterval.Day, LongDateZeroTime, TodayDateZeroTime)
                            Sendb("&nbsp;" & Logix.ToShortDateTimeString(LongDate, MyCommon))
                            If DaysDiff < 0 Then
                            ElseIf DaysDiff = 0 Then
                                Sendb(" (" & StrConv(Copient.PhraseLib.Lookup("term.today", LanguageID), VbStrConv.Lowercase) & ")")
                            ElseIf DaysDiff = 1 Then
                                Sendb(" (" & StrConv(Copient.PhraseLib.Lookup("term.yesterday", LanguageID), VbStrConv.Lowercase) & ")")
                            Else
                                Sendb(" (" & DaysDiff & " " & StrConv(Copient.PhraseLib.Lookup("term.daysago", LanguageID), VbStrConv.Lowercase) & ")")
                            End If
                        Else
                            Sendb(Copient.PhraseLib.Lookup("term.never", LanguageID))
                        End If
                        Sendb("<br />")
                    %>
                    <br class="half" />
                    <h3>
                        <%Sendb(Copient.PhraseLib.Lookup("term.lastvalidationmessage", LanguageID))%>
          :
                    </h3>
                    <%
                        Dim validationMessage As String = GetLastDeployValidationMessage(MyCommon, OfferID)
                        If (validationMessage <> "") Then
                            If (validationMessage <> "term.validationsuccessful" AndAlso validationMessage <> "alert.awaitingrecommendation") Then
                                'Checking ErrMgrTerms applicable in Offer Summary, because validationMessage is getting terms sometimes and actual term related text sometimes --AL-8825
                                Dim ErrMsgTerms As String = "term.validationsuccessful, term.ReqTransFailed, deploy, deferdeploy, " &
                                    "offer-sum.uomprecisionallowed, offer-sum.currhasbeenexceeded, offer-sum.currprecisionallowed, offer-sum.currhasbeenexceeded, " &
                                    "offer-sum.unsupporteduom, offer-sum.undefineduom, offer-sum.unsupportedcurrency, offer-sum.nonselectedcurrency, " &
                                    "cpeoffer-sum.deployalertforlockout, UEoffer-sum.tier-setup-invalid, offer-sum.required-incomplete," &
                                    "cpeoffer-sum.deployalertforexpire, UEoffer-sum.deployalert, UEoffer-deploy-error.invalidanalyticscg, UEoffer-deploy-error.awaitingrecommendations"
                                If ErrMsgTerms.Contains(validationMessage) Then
                                    Sendb("<font color=""red"">" & Copient.PhraseLib.Lookup(validationMessage, LanguageID) & "</font>")
                                Else
                                    Sendb("<font color=""red"">" & validationMessage & "</font>")
                                End If
                            Else
                                Sendb(Copient.PhraseLib.Lookup(validationMessage, LanguageID))
                            End If
                        End If
                        Sendb("<br />")
                    %>
                    <br class="half" />
                    <h3>
                        <%Sendb(Copient.PhraseLib.Lookup("offer-sum.crmlastsent", LanguageID))%>
          :
                    </h3>
                    <%
                        MyCommon.QueryStr = "select LastCRMSendDate from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID & ";"
                        rst4 = MyCommon.LRT_Select
                        If rst4.Rows.Count > 0 Then
                            LongDate = MyCommon.NZ(rst4.Rows(0).Item("LastCRMSendDate"), "1/1/1900")
                        Else
                            LongDate = "1/1/1900"
                        End If
                        LongDateZeroTime = New Date(LongDate.Year, LongDate.Month, LongDate.Day, 0, 0, 0)
                        If LongDate > "1/1/1900" Then
                            DaysDiff = DateDiff(DateInterval.Day, LongDateZeroTime, TodayDateZeroTime)
                            Sendb("&nbsp;" & Logix.ToShortDateTimeString(LongDate, MyCommon))
                            If DaysDiff < 0 Then
                            ElseIf DaysDiff = 0 Then
                                Sendb(" (" & StrConv(Copient.PhraseLib.Lookup("term.today", LanguageID), VbStrConv.Lowercase) & ")")
                            ElseIf DaysDiff = 1 Then
                                Sendb(" (" & StrConv(Copient.PhraseLib.Lookup("term.yesterday", LanguageID), VbStrConv.Lowercase) & ")")
                            Else
                                Sendb(" (" & DaysDiff & " " & StrConv(Copient.PhraseLib.Lookup("term.daysago", LanguageID), VbStrConv.Lowercase) & ")")
                            End If
                        Else
                            Sendb(Copient.PhraseLib.Lookup("term.never", LanguageID))
                        End If
                        Sendb("<br />")
                    %>
                    <br class="half" />
                    <h3>
                        <%Sendb(Copient.PhraseLib.Lookup("offer-sum.crmlastreceived", LanguageID))%>
          :
                    </h3>
                    <%
                        MyCommon.QueryStr = "select LastUpdate from CRMImportQueue with (NoLock) where OfferID=" & OfferID & " order by LastUpdate desc;"
                        rst4 = MyCommon.LRT_Select
                        If rst4.Rows.Count > 0 Then
                            LongDate = MyCommon.NZ(rst4.Rows(0).Item("LastUpdate"), "1/1/1900")
                        Else
                            LongDate = "1/1/1900"
                        End If
                        LongDateZeroTime = New Date(LongDate.Year, LongDate.Month, LongDate.Day, 0, 0, 0)
                        If LongDate > "1/1/1900" Then
                            DaysDiff = DateDiff(DateInterval.Day, LongDateZeroTime, TodayDateZeroTime)
                            Sendb("&nbsp;" & Logix.ToShortDateTimeString(LongDate, MyCommon))
                            If DaysDiff < 0 Then
                            ElseIf DaysDiff = 0 Then
                                Sendb(" (" & StrConv(Copient.PhraseLib.Lookup("term.today", LanguageID), VbStrConv.Lowercase) & ")")
                            ElseIf DaysDiff = 1 Then
                                Sendb(" (" & StrConv(Copient.PhraseLib.Lookup("term.yesterday", LanguageID), VbStrConv.Lowercase) & ")")
                            Else
                                Sendb(" (" & DaysDiff & " " & StrConv(Copient.PhraseLib.Lookup("term.daysago", LanguageID), VbStrConv.Lowercase) & ")")
                            End If
                        Else
                            Sendb(Copient.PhraseLib.Lookup("term.never", LanguageID))
                        End If
                        Sendb("<br />")
                    %>
                    <br class="half" />
                    <hr class="hidden" />
                </div>
            </div>
            <% End If%>
        </div>
        <div id="gutter">
        </div>
        <div id="column2">
            <div class="box" id="locationSum">
                <%  
                    If (LinksDisabled) Then
                        Send("<h2 style=""float:left;""><span>" & Copient.PhraseLib.Lookup("term.locations", LanguageID) & "</span></h2>")
                    Else
                        Send("<h2 style=""float:left;""><span><a href=""/logix/UE/UEoffer-loc.aspx?OfferID=" & OfferID & """>" & Copient.PhraseLib.Lookup("term.locations", LanguageID) & "</a></span></h2>")
                    End If
                    Send_BoxResizer("locationbody", "imgLocations", Copient.PhraseLib.Lookup("term.locations", LanguageID), True)
                %>
                <div id="locationbody">
                    <%
                        If (BannersEnabled) Then
                            Sendb("<h3>" & Copient.PhraseLib.Lookup(IIf(BannerCt > 1, "term.banners", "term.banner"), LanguageID) & "</h3>")
                            Send("<ul class=""condensed"">")
                            For i = 0 To BannerNames.GetUpperBound(0)
                                If (BannerIDs(i) > -1) Then
                                    If Popup Then
                                        Sendb("<li>" & BannerNames(i) & "</li>")
                                    Else
                                        Sendb("<li><a href=""../banner-edit.aspx?BannerID=" & BannerIDs(i) & """>" & BannerNames(i) & "</a></li>")
                                    End If
                                Else
                                    Sendb("<li>" & BannerNames(i) & "</li>")
                                End If
                            Next
                            Send("</ul>")
                        End If
                    %>
                    <h3>
                        <% Sendb(Copient.PhraseLib.Lookup("term.storegroups", LanguageID))%>
                    </h3>
                    <%
                        GCount = 0
                        MyCommon.QueryStr = "select OL.LocationGroupID,LG.Name,LG.PhraseID from OfferLocations as OL with (NoLock) " & _
                                            "left join LocationGroups as LG with (NoLock) on LG.LocationGroupID=OL.LocationGroupID " & _
                                            "where Excluded=0 and OL.Deleted=0 and OL.OfferID=" & OfferID & " order by LG.Name"
                        rst = MyCommon.LRT_Select
                        For Each row In rst.Rows
                            If (MyCommon.NZ(row.Item("LocationGroupID"), 0) = 1) Then
                                AnyStoreUsed = True
                            End If
                        Next
                        If rst.Rows.Count > 0 And Not AnyStoreUsed Then
                            For Each row In rst.Rows
                                MyCommon.QueryStr = "select count(*) as GCount from LocGroupItems with (NoLock) where LocationGroupID = " & MyCommon.NZ(row.Item("LocationGroupID"), 0) & " And Deleted = 0"
                                rst2 = MyCommon.LRT_Select()
                                For Each row2 In rst2.Rows
                                    GCount = GCount + (row2.Item("GCount"))
                                Next
                            Next
                            Sendb("&nbsp;" & GCount & " ")
                            If GCount <> 1 Then
                                Sendb(StrConv(Copient.PhraseLib.Lookup("term.stores", LanguageID), VbStrConv.Lowercase))
                            Else
                                Sendb(StrConv(Copient.PhraseLib.Lookup("term.store", LanguageID), VbStrConv.Lowercase))
                            End If
                            Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.from", LanguageID), VbStrConv.Lowercase) & " ")
                            Sendb(rst.Rows.Count & " ")
                            If rst.Rows.Count <> 1 Then
                                Send(StrConv(Copient.PhraseLib.Lookup("term.storegroups", LanguageID), VbStrConv.Lowercase) & ":<br />")
                            Else
                                Send(StrConv(Copient.PhraseLib.Lookup("term.storegroup", LanguageID), VbStrConv.Lowercase) & ":<br />")
                            End If
                        End If
                        Send("<ul class=""condensed"">")
                        For Each row In rst.Rows
                            If IsDBNull(row.Item("PhraseID")) Then
                                If Popup Then
                                    Sendb("<li>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25))
                                Else
                                    Sendb("<li><a href=""../lgroup-edit.aspx?LocationGroupId=" & MyCommon.NZ(row.Item("LocationGroupID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
                                End If
                            Else
                                If (row.Item("PhraseID") = 0) Then
                                    If Popup Then
                                        Sendb("<li>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25))
                                    Else
                                        Sendb("<li><a href=""../lgroup-edit.aspx?LocationGroupId=" & MyCommon.NZ(row.Item("LocationGroupID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
                                    End If
                                Else
                                    Sendb("<li>" & MyCommon.SplitNonSpacedString(Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID), 25))
                                End If
                            End If
                            MyCommon.QueryStr = "select count(*) as GCount from LocGroupItems with (NoLock) where LocationGroupID = " & MyCommon.NZ(row.Item("LocationGroupID"), 0) & " And Deleted = 0"
                            rst2 = MyCommon.LRT_Select()
                            For Each row2 In rst2.Rows
                                If (MyCommon.NZ(row.Item("LocationGroupID"), 0) > 1) Then
                                    Sendb(" (" & row2.Item("GCount") & ")")
                                End If
                            Next
                        Next
                        If (rst.Rows.Count = 0) Then
                            Sendb("<li>" & Copient.PhraseLib.Lookup("term.none", LanguageID))
                        End If
                        ' Check for and display any excluded store groups
                        MyCommon.QueryStr = "select OL.LocationGroupID,LG.Name,LG.PhraseID from OfferLocations as OL with (NoLock) inner join LocationGroups as LG with (NoLock) on " & _
                                            "LG.LocationGroupID=OL.LocationGroupID where Excluded=1 and OL.deleted=0 and OL.OfferID=" & OfferID
                        rst = MyCommon.LRT_Select
                        rowCount = rst.Rows.Count
                        If rowCount > 0 Then
                            Sendb("&nbsp;" & StrConv(Copient.PhraseLib.Lookup("term.excluding", LanguageID), VbStrConv.Lowercase) & "&nbsp;")
                            For Each row In rst.Rows
                                MyCommon.QueryStr = "select count(*) as GCount from LocGroupItems with (NoLock) where LocationGroupID = " & MyCommon.NZ(row.Item("LocationGroupID"), 0) & " And Deleted=0"
                                rst2 = MyCommon.LRT_Select()
                                For Each row2 In rst2.Rows
                                    GCount = GCount + (row2.Item("GCount"))
                                Next
                                If Not Popup Then
                                    Sendb("<a href=""../lgroup-edit.aspx?LocationGroupId=" & MyCommon.NZ(row.Item("LocationGroupID"), 0) & """>")
                                End If
                                If IsDBNull(row.Item("PhraseID")) Then
                                    Sendb(MyCommon.NZ(row.Item("Name"), ""))
                                Else
                                    If (row.Item("PhraseID") = 0) Then
                                        Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25))
                                    Else
                                        Sendb(MyCommon.SplitNonSpacedString(Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID), 25))
                                    End If
                                End If
                                If Not Popup Then
                                    Sendb("</a> ")
                                End If
                                Sendb("(" & GCount & ") ")
                            Next
                        End If
                        Send("</li>")
                        Send("</ul>")
                    %>
                    <h3>
                        <% Sendb(Copient.PhraseLib.Lookup("term.terminals", LanguageID))%>
                    </h3>
                    <% 
                        MyCommon.QueryStr = "select OT.TerminalTypeID as TID,T.Name,T.PhraseID from OfferTerminals as OT with (NoLock) left join TerminalTypes as T with (NoLock) on OT.TerminalTypeID=T.TerminalTypeID " & _
                                            "where OfferID=" & OfferID & " and Excluded=0 order by T.NAme"
                        rst = MyCommon.LRT_Select
                        Send("<ul class=""condensed"">")
                        For Each row In rst.Rows
                            If (MyCommon.NZ(row.Item("Name"), "").ToString.Trim = "All CPE Terminals" OrElse MyCommon.NZ(row.Item("Name"), "").ToString.Trim = "All CM Terminals" OrElse MyCommon.NZ(row.Item("Name"), "").ToString.Trim = "All DP Terminals") Then
                                If IsDBNull(row.Item("PhraseID")) Then
                                    Sendb("<li>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25))
                                Else
                                    Sendb("<li>" & MyCommon.SplitNonSpacedString(Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID), 25))
                                End If
                            Else
                                If Popup Then
                                    Sendb("<li>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25))
                                Else
                                    Sendb("<li><a href=""../terminal-edit.aspx?TerminalID=" & MyCommon.NZ(row.Item("TID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
                                End If
                            End If
                        Next
                        If (rst.Rows.Count = 0) Then
                            Sendb("<li>" & Copient.PhraseLib.Lookup("term.none", LanguageID))
                        End If
                        ' Check for and display any excluded terminals
                        MyCommon.QueryStr = "select OT.TerminalTypeID as TID,T.Name,T.PhraseID from OfferTerminals as OT with (NoLock) left join TerminalTypes as T with (NoLock) on OT.TerminalTypeID=T.TerminalTypeID " & _
                                            "where OfferID=" & OfferID & " and Excluded=1 order by T.NAme"
                        rst = MyCommon.LRT_Select
                        rowCount = rst.Rows.Count
                        If rowCount > 0 Then
                            Sendb("&nbsp;" & StrConv(Copient.PhraseLib.Lookup("term.excluding", LanguageID), VbStrConv.Lowercase) & "&nbsp;")
                            For Each row In rst.Rows
                                x = x + 1
                                If (MyCommon.NZ(row.Item("Name"), "").ToString.Trim = "All CPE Terminals" OrElse MyCommon.NZ(row.Item("Name"), "").ToString.Trim = "All CM Terminals" OrElse MyCommon.NZ(row.Item("Name"), "").ToString.Trim = "All DP Terminals") Then
                                    If IsDBNull(row.Item("PhraseID")) Then
                                        Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), 0), 25))
                                    Else
                                        Sendb(MyCommon.SplitNonSpacedString(Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID), 25))
                                    End If
                                Else
                                    If Popup Then
                                        Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25))
                                    Else
                                        Sendb("<a href=""../terminal-edit.aspx?TerminalID=" & MyCommon.NZ(row.Item("TID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a> ")
                                    End If
                                End If
                                If x = (rst.Rows.Count - 1) Then
                                    Send(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
                                ElseIf x < rst.Rows.Count Then
                                    Send(", ")
                                Else
                                    Send("")
                                End If
                            Next
                        End If
                        Send("</li>")
                        Send("</ul>")
                    %>
                    <hr class="hidden" />
                </div>
            </div>
            <div class="box" id="offernotifications">
                <%  
                    If (LinksDisabled) Then
                        Send("<h2 style=""float:left;""><span>" & Copient.PhraseLib.Lookup("term.channels", LanguageID) & "</span></h2>")
                    Else
                        Send("<h2 style=""float:left;""><span><a href=""/logix/offer-channels.aspx?OfferID=" & OfferID & """>" & Copient.PhraseLib.Lookup("term.channels", LanguageID) & "</a></span></h2>")
                    End If
                    Send_BoxResizer("notificationbody", "imgNotifications", "Notifications", True)
                %>
                <div id="notificationbody">
                    <%
                        MyCommon.QueryStr = "select distinct CO.ChannelID, CH.Name, CH.PhraseTerm, CO.StartDate, CO.EndDate  " & _
                                            "from ChannelOffers as CO with (NoLock) " & _
                                            "inner join Channels as CH with (NoLock) on CH.ChannelID = CO.ChannelID " & _
                                            "where CO.OfferID = " & OfferID & "and StartDate is not null and EndDate is not null " & _
                                            "union " & _
                                            "select distinct CO.ChannelID, CH.Name, CH.PhraseTerm, CO.StartDate, CO.EndDate  " & _
                                            "from ChannelOfferAssets as COA with (NoLock) " & _
                                            "inner join ChannelOffers as CO with (NoLock) on CO.ChannelID = COA.ChannelID " & _
                                            "inner join Channels as CH with (NoLock) on CH.ChannelID = CO.ChannelID " & _
                                            "where CO.OfferID = " & OfferID & " and ISNULL(MediaData, '') <> '' " & _
                                            "order by ChannelID; "
                        rst2 = MyCommon.LRT_Select()
                        If rst2.Rows.Count > 0 Then
                            For Each row2 In rst2.Rows
                                Send("<h3>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row2.Item("PhraseTerm"), "").ToString, LanguageID, MyCommon.NZ(row2.Item("Name"), "")) & "</h3>")
                                Sendb("<span style=""margin-left: 10px;"">")
                                LongDate = MyCommon.NZ(row2.Item("StartDate"), "1/1/1900")
                                If LongDate > "1/1/1900" Then Sendb(Logix.ToShortDateString(LongDate, MyCommon)) Else Sendb("?")
                                Sendb(" - ")
                                LongDate = MyCommon.NZ(row2.Item("EndDate"), "1/1/1900")
                                If LongDate > "1/1/1900" Then Sendb(Logix.ToShortDateString(LongDate, MyCommon)) Else Sendb("?")
                                Sendb(" (" & DateDiff(DateInterval.Day, MyCommon.NZ(row2.Item("StartDate"), "1/1/1900"), MyCommon.NZ(row2.Item("EndDate"), "1/1/1900")) + 1 & " ")
                                If DateDiff(DateInterval.Day, MyCommon.NZ(row2.Item("StartDate"), "1/1/1900"), MyCommon.NZ(row2.Item("EndDate"), "1/1/1900")) + 1 = 1 Then
                                    Sendb(StrConv(Copient.PhraseLib.Lookup("term.day", LanguageID), VbStrConv.Lowercase) & ")")
                                Else
                                    Sendb(StrConv(Copient.PhraseLib.Lookup("term.days", LanguageID), VbStrConv.Lowercase) & ")")
                                End If
                                Send("</span>")
                                Send("<br />")
                                Send("<br class=""half"" />")

                                Select Case MyCommon.NZ(row2.Item("ChannelID"), 0)
                                    Case POS_CHANNEL_ID
                                        Send("<div style=""margin-left: 10px;"">")
                                        ' Printed message notifications
                                        MyCommon.QueryStr = "select PM.MessageID, PM.MessageTypeID, PMT.BodyText, D.DeliverableID " & _
                                                            "from CPE_Deliverables as D with (NoLock) " & _
                                                            "inner join PrintedMessages as PM with (NoLock) on D.OutputID=PM.MessageID " & _
                                                            "inner join PrintedMessageTiers as PMT with (NoLock) on PM.MessageID=PMT.MessageID " & _
                                                            "where D.Deleted=0 and D.RewardOptionPhase=1 and D.RewardOptionID=" & roid & " and D.DeliverableTypeID=4 and PMT.TierLevel=1;"
                                        rst = MyCommon.LRT_Select()
                                        If (rst.Rows.Count > 0) Then
                                            counter = counter + 1
                                            Dim Details As StringBuilder
                                            Send("<h3 style=""color: #606060;"">" & Copient.PhraseLib.Lookup("term.printedmessages", LanguageID) & "</h3>")
                                            Send("<ul class=""condensed"">")
                                            For Each row In rst.Rows
                                                Details = New StringBuilder(200)
                                                Details.Append(ReplaceTags(MyCommon.NZ(row.Item("BodyText"), "")))
                                                If (Details.ToString().Length > 80) Then
                                                    Details = Details.Remove(77, (Details.Length - 77))
                                                    Details.Append("...")
                                                End If
                                                Details.Replace(vbCrLf, "<br />")
                                                Send("<li>""" & HttpUtility.HtmlEncode(MyCommon.SplitNonSpacedString(Details.ToString, 25)) & """</li>")
                                            Next
                                            Send("</ul>")
                                            Send("<br class=""half"" />")
                                        End If

                                        ' Image URL notifications
                                        MyCommon.QueryStr = "select PT.PKID, PT.PassThruRewardID, PTT.Data, D.DeliverableID " &
                                                            "from CPE_Deliverables as D with (NoLock) " &
                                                            "inner join PassThrus as PT with (NoLock) on D.OutputID=PT.PKID " &
                                                            "inner join PassThruTiers as PTT with (NoLock) on PT.PKID=PTT.PTPKID " &
                                                            "where D.Deleted=0 and D.RewardOptionPhase=1 and D.RewardOptionID=" & roid & " and D.DeliverableTypeID=12 and PTT.TierLevel=1;"
                                        rst = MyCommon.LRT_Select()
                                        If (rst.Rows.Count > 0) Then
                                            counter = counter + 1
                                            Dim Details As StringBuilder
                                            Send("<h3 style=""color: #606060;"">" & Copient.PhraseLib.Lookup("mediatype.imgurl", LanguageID) & "</h3>")
                                            Send("<ul class=""condensed"">")
                                            For Each row In rst.Rows
                                                Details = New StringBuilder(200)
                                                Details.Append(ReplaceTags(MyCommon.NZ(row.Item("Data"), "")))
                                                If (Details.ToString().Length > 80) Then
                                                    Details = Details.Remove(77, (Details.Length - 77))
                                                    Details.Append("...")
                                                End If
                                                Details.Replace(vbCrLf, "<br />")
                                                Send("<li>""" & HttpUtility.HtmlEncode(MyCommon.SplitNonSpacedString(Details.ToString, 25)) & """</li>")
                                            Next
                                            Send("</ul>")
                                            Send("<br class=""half"" />")
                                        End If
                                        
                                        ' Cashier message notifications
                                        MyCommon.QueryStr = "select D.DeliverableID, CM.MessageID, CMT.Line1, CMT.Line2 " & _
                                                                                                "from CPE_Deliverables as D with (NoLock) " & _
                                                                                                "inner join CPE_CashierMessages as CM with (NoLock) on D.OutputID=CM.MessageID " & _
                                                                                                "left join CPE_CashierMessageTiers as CMT with (NoLock) on CMT.MessageID=CM.MessageID " & _
                                                                                                "where D.RewardOptionID=" & roid & " and D.Deleted=0 and DeliverableTypeID=9 and D.RewardOptionPhase=1;"
                                        rst = MyCommon.LRT_Select
                                        If (rst.Rows.Count > 0) Then
                                            counter = counter + 1
                                            Send("<h3 style=""color: #606060;"">" & Copient.PhraseLib.Lookup("term.cashiermessages", LanguageID) & "</h3>")
                                            Send("<ul class=""condensed"">")
                                            For Each row In rst.Rows
                                                Send("<li>""" & MyCommon.NZ(row.Item("Line1"), "") & "<br />" & MyCommon.NZ(row.Item("Line2"), "") & """</li>")
                                            Next
                                            Send("</ul>")
                                            Send("<br class=""half"" />")
                                        End If

                                        ' Graphics notifications
                                        MyCommon.QueryStr = "select OSA.Name as GraphicName, OSA.OnScreenAdID as AdID, OSA.Width as Width, OSA.Height as Height, OSA.GraphicSize as Size, OSA.ImageType as Type, " & _
                                                                                                            "D.DeliverableID, D.ScreenCellID as CellID, OSA.Width, OSA.Height, OSA.ImageType, SC.Name as CellName, SL.Name as LayoutName " & _
                                                                                                            "from OnScreenAds as OSA with (NoLock) Inner Join CPE_Deliverables as D with (NoLock) on OSA.OnScreenAdID=D.OutputID and D.RewardOptionID=" & roid & " " & _
                                                                                                            "and OSA.Deleted=0 and D.Deleted=0 and D.DeliverableTypeID=1 and D.RewardOptionPhase=1 " & _
                                                                                                            "Inner Join ScreenCells as SC with (NoLock) on D.ScreenCellID=SC.CellID " & _
                                                                                                            "Inner Join ScreenLayouts as SL with (NoLock) on SL.LayoutID=SC.LayoutID;"
                                        rst = MyCommon.LRT_Select
                                        If (rst.Rows.Count > 0) Then
                                            counter = counter + 1
                                            Send("<h3 style=""color: #606060;"">" & Copient.PhraseLib.Lookup("term.graphics", LanguageID) & "</h3>")
                                            Send("<ul class=""condensed"">")
                                            For Each row In rst.Rows
                                                If Popup Then
                                                    Sendb("<li>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("GraphicName"), ""), 25) & "&nbsp;")
                                                Else
                                                    Sendb("<li><a href=""graphic-edit.aspx?OnScreenAdId=" & MyCommon.NZ(row.Item("AdId"), "") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("GraphicName"), ""), 25) & "</a>&nbsp;")
                                                End If
                                                Sendb("(" & MyCommon.NZ(row.Item("Width"), "") & " x " & MyCommon.NZ(row.Item("Height"), ""))
                                                If MyCommon.NZ(row.Item("Type"), 0) = 1 Then
                                                    Sendb("&nbsp;" & Copient.PhraseLib.Lookup("term.jpeg", LanguageID))
                                                ElseIf MyCommon.NZ(row.Item("Type"), 0) = 2 Then
                                                    Sendb("&nbsp;" & Copient.PhraseLib.Lookup("term.gif", LanguageID))
                                                End If
                                                Send(")</li>")
                                            Next
                                            Send("</ul>")
                                            Send("<br class=""half"" />")
                                        End If

                                        ' Touchpoint notifications
                                        MyCommon.QueryStr = "select RO.Name, RO.RewardOptionID, TA.OnScreenAdID as ParentAdID " & _
                                                            "from CPE_RewardOptions RO with (NoLock) inner join CPE_DeliverableROIDs DR with (NoLock) on RO.RewardOptionID = DR.RewardOptionID " & _
                                                            "inner join CPE_Deliverables D with (NoLock) on D.DeliverableID = DR.DeliverableID inner join TouchAreas TA with (NoLock) on DR.AreaID = TA.AreaID " & _
                                                            "where RO.Deleted=0 and DR.Deleted=0 and TA.Deleted=0 and D.Deleted = 0 and RO.IncentiveID=" & OfferID & " and RO.TouchResponse=1 and D.RewardOptionPhase=1 order by RO.rewardoptionid;"
                                        rst = MyCommon.LRT_Select
                                        If (rst.Rows.Count > 0) Then
                                            counter = counter + 1
                                            Dim tpROID As Integer = 0
                                            Send("<h3 style=""color: #606060;"">" & Copient.PhraseLib.Lookup("term.touchpoints", LanguageID) & "</h3>")
                                            Send("<ul class=""condensed"">")
                                            For Each row In rst.Rows
                                                tpROID = MyCommon.NZ(row.Item("RewardOptionID"), 0)
                                                Send("<li>")
                                                If Popup Then
                                                    Send(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)), 25) & "<br />")
                                                Else
                                                    Send("<a href=""graphic-edit.aspx?OnScreenAdId=" & MyCommon.NZ(row.Item("ParentAdID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)), 25) & "</a><br />")
                                                End If
                                                SetSummaryOnly(True)
                                                Send_TouchpointRewards(OfferID, tpROID, 1, TierLevels)
                                                Send("</li>")
                                            Next
                                            Send("</ul>")
                                            Send("<br class=""half"" />")
                                        End If

                                        'Accumulation notifications
                                        MyCommon.QueryStr = "select PM.MessageID, PM.MessageTypeID, PMT.BodyText, D.DeliverableID " & _
                                                            "from CPE_deliverables D with (NoLock) inner join PrintedMessages PM with (NoLock) on D.OutputID = PM.MessageID " & _
                                                            "inner join PrintedMessageTiers PMT with (NoLock) on PM.MessageID = PMT.MessageID " & _
                                                            "where D.Deleted = 0 and D.RewardOptionPhase=2 and D.RewardOptionID=" & roid & " and D.DeliverableTypeID=4 and PMT.TierLevel = 1;"
                                        rst = MyCommon.LRT_Select()
                                        If (rst.Rows.Count > 0) Then
                                            counter = counter + 1
                                            Dim Details As StringBuilder
                                            Send("<h3 style=""color: #606060;"">" & Copient.PhraseLib.Lookup("term.accumulationmessage", LanguageID) & "</h3>")
                                            Send("<ul class=""condensed"">")
                                            For Each row In rst.Rows
                                                Details = New StringBuilder(200)
                                                Details.Append(ReplaceTags(MyCommon.NZ(row.Item("BodyText"), "")))
                                                If (Details.ToString().Length > 80) Then
                                                    Details = Details.Remove(77, (Details.Length - 77))
                                                    Details.Append("...")
                                                End If
                                                Details.Replace(vbCrLf, "<br />")
                                                Send("<li>""" & HttpUtility.HtmlEncode(MyCommon.SplitNonSpacedString(Details.ToString, 25)) & """</li>")
                                            Next
                                            Send("</ul>")
                                        End If

                                        If (counter = 0) Then
                                            Send("<h3 style=""color: #606060;"">" & Copient.PhraseLib.Lookup("term.Noprintedorgraphicsmessage", LanguageID) & "</h3>")
                                        End If
                                        counter = 0

                                        Send("</div>")
                                End Select
                            Next
                        Else
                            Send(Copient.PhraseLib.Lookup("CPEoffer-sum.noChannelsConfigured", LanguageID))
                        End If

                    %>
                    <hr class="hidden" />
                </div>
            </div>
            <%        
                Send("<div class=""box"" id=""OptInConditions"">")
                If (LinksDisabled) Then
                    Send("<h2 style=""float:left;""><span>" & Copient.PhraseLib.Lookup("term.optinconditions", LanguageID) & "</span></h2>")
                Else
                    Send("<h2 style=""float:left;""><span><a href=""UEoffer-con.aspx?OfferID=" & OfferID & """>" & Copient.PhraseLib.Lookup("term.optinconditions", LanguageID) & "</a></span></h2>")
                End If
                Send_BoxResizer("optinconditionbody", "imgOptInConditions", Copient.PhraseLib.Lookup("term.optinconditions", LanguageID), True)
                Send("<div id=""optinconditionbody"">")
                Dim conditionlength As Integer
                Dim customers As List(Of CMS.AMS.Models.Customer)
                Offer = m_Offer.GetOffer(OfferID, LoadOfferOptions.AllEligibilityConditions)
                ' Customer Eligibility Conditions
                If Offer.EligibleCustomerGroupConditions IsNot Nothing Then
                    Send("<h3>" & Copient.PhraseLib.Lookup("term.customerconditions", LanguageID) & "</h3>")
                    Send("<ul class=""condensed"">")
                    conditionlength = Offer.EligibleCustomerGroupConditions.IncludeCondition.Count
                    For Each includeCondition As CustomerConditionDetails In Offer.EligibleCustomerGroupConditions.IncludeCondition
                        Send("<li>")
                        If includeCondition.CustomerGroupID = 1 OrElse includeCondition.CustomerGroupID = 2 OrElse includeCondition.CustomerGroupID = 3 OrElse includeCondition.CustomerGroupID = 4 OrElse Not m_AnalyticsCustomerGroups.HasValidExternalSegmentId(includeCondition.CustomerGroupID) Then
                            Sendb(MyCommon.SplitNonSpacedString(includeCondition.CustomerGroup.Name, 25))
                            If conditionlength > 1 Then
                                Sendb(" and ")
                            Else
                                Sendb(" ")
                            End If
                        Else
                            Sendb("<a href=""../cgroup-edit.aspx?CustomerGroupID=" & includeCondition.CustomerGroupID & """>" & MyCommon.SplitNonSpacedString(includeCondition.CustomerGroup.Name, 25) & "</a>")
                            customers = m_customerGroup.GetCustomersByGroupID(includeCondition.CustomerGroupID)
                            If customers.Count > 0 Then
                                Sendb(" (" & customers.Count & ") " & IIf(conditionlength > 1, " or ", "") & "")
                            ElseIf conditionlength > 1 Then
                                Sendb(" or ")
                            Else
                                Sendb(" ")
                            End If
                        End If
                        conditionlength = conditionlength - 1
                        Send("</li>")
                    Next

                    conditionlength = Offer.EligibleCustomerGroupConditions.ExcludeCondition.Count
                    If conditionlength > 0 Then
                        Sendb(Copient.PhraseLib.Lookup("term.excluding", LanguageID) & " ")
                    End If
                    For Each excludeCondition As CustomerConditionDetails In Offer.EligibleCustomerGroupConditions.ExcludeCondition
                        Send("<li>")
                        If excludeCondition.CustomerGroupID = 1 OrElse excludeCondition.CustomerGroupID = 2 OrElse excludeCondition.CustomerGroupID = 3 OrElse excludeCondition.CustomerGroupID = 4 OrElse Not m_AnalyticsCustomerGroups.HasValidExternalSegmentId(excludeCondition.CustomerGroupID) Then
                            Sendb(MyCommon.SplitNonSpacedString(excludeCondition.CustomerGroup.Name, 25))
                            If conditionlength > 1 Then
                                Sendb(" and ")
                            Else
                                Sendb(" ")
                            End If
                        Else
                            Sendb("<a href=""../cgroup-edit.aspx?CustomerGroupID=" & excludeCondition.CustomerGroupID & """>" & MyCommon.SplitNonSpacedString(excludeCondition.CustomerGroup.Name, 25) & "</a>")
                            customers = m_customerGroup.GetCustomersByGroupID(excludeCondition.CustomerGroupID)
                            If customers.Count > 0 Then
                                Sendb(" (" & customers.Count & ") " & IIf(conditionlength > 1, " and ", "") & "")
                            ElseIf conditionlength > 1 Then
                                Sendb(" and ")
                            End If
                        End If
                        conditionlength = conditionlength - 1
                        Send("</li>")
                    Next
                    Send("</ul>")
                    Send("<br class=""half"" />")

                    ' Points Eligibility Conditions
                    If (Offer.EligiblePointsProgramConditions.Count > 0) Then
                        Send("<h3>" & Copient.PhraseLib.Lookup("term.pointsconditions", LanguageID) & "</h3>")
                        Send("<ul class=""condensed"">")
                        For Each pointscondition As Models.PointsCondition In Offer.EligiblePointsProgramConditions
                            Sendb("<li>")
                            Sendb(CMS.Utilities.NZ(pointscondition.Quantity, 0))
                            If (pointscondition.ProgramID > 0) Then
                                If Popup Then
                                    Sendb(" " & MyCommon.SplitNonSpacedString(CMS.Utilities.NZ(pointscondition.PointsProgram.ProgramName, ""), 25))
                                Else
                                    Sendb(" <a href=""../point-edit.aspx?ProgramGroupID=" & CMS.Utilities.NZ(pointscondition.PointsProgram.ProgramID, "") & """>" & MyCommon.SplitNonSpacedString(CMS.Utilities.NZ(pointscondition.PointsProgram.ProgramName, ""), 25) & "</a>")
                                End If
                            ElseIf (pointscondition.ProgramID = 0 AndAlso pointscondition.RequiredFromTemplate = False) Then
                                Sendb(" <span class=""red"">")
                                Sendb("<small>(" & Copient.PhraseLib.Lookup("term.requiredbytemplate", LanguageID) & ")</small>")
                                Sendb("</span>")
                            Else
                                Sendb(MyCommon.SplitNonSpacedString(CMS.Utilities.NZ(pointscondition.PointsProgram.ProgramName, ""), 25))
                            End If
                            Send("</li>")
                        Next
                        Send("</ul>")
                        Send("<br class=""half"" />")
                    End If

                    ' Stored Value Eligibility Conditions
                    If (Offer.EligibleSVProgramConditions.Count > 0) Then
                        Send("<h3>" & Copient.PhraseLib.Lookup("term.storedvalueconditions", LanguageID) & "</h3>")
                        Send("<ul class=""condensed"">")
                        For Each storedvaluecondition As Models.SVCondition In Offer.EligibleSVProgramConditions
                            Sendb("<li>")
                            Sendb(CMS.Utilities.NZ(storedvaluecondition.Quantity, 0))

                            If storedvaluecondition.SVProgram.SVType.SVTypeID > 1 Then
                                Sendb(" ($" & Math.Round(storedvaluecondition.SVProgram.Value * storedvaluecondition.Quantity, storedvaluecondition.SVProgram.SVType.ValuePrecision).ToString(MyCommon.GetAdminUser.Culture) & ")")
                            End If
                            If (storedvaluecondition.SVProgramID > 0) Then
                                If Popup Then
                                    Sendb(" " & MyCommon.SplitNonSpacedString(CMS.Utilities.NZ(storedvaluecondition.SVProgram.ProgramName, ""), 25))
                                Else
                                    Sendb(" <a href=""../SV-edit.aspx?ProgramGroupID=" & CMS.Utilities.NZ(storedvaluecondition.SVProgram.SVProgramID, "") & """>" & MyCommon.SplitNonSpacedString(CMS.Utilities.NZ(storedvaluecondition.SVProgram.ProgramName, ""), 25) & "</a>")
                                End If
                            ElseIf (storedvaluecondition.SVProgramID = 0 AndAlso storedvaluecondition.RequiredFromTemplate = False) Then
                                Sendb(" <span class=""red"">")
                                Sendb("<small>(" & Copient.PhraseLib.Lookup("term.requiredbytemplate", LanguageID) & ")</small>")
                                Sendb("</span>")
                            Else
                                Sendb(MyCommon.SplitNonSpacedString(CMS.Utilities.NZ(storedvaluecondition.SVProgram.ProgramName, ""), 25))
                            End If
                            Send("</li>")
                        Next
                        Send("</ul>")
                        Send("<br class=""half"" />")
                    End If

                Else
                    Send("<h3>" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "</h3>")
                End If
                Send("<hr class=""hidden"" />")
                Send("</div>")
                Send("</div>")
            %>
            <div class="box" id="offerconditions">
                <%  
                    If (LinksDisabled) Then
                        Send("<h2 style=""float:left;""><span>" & Copient.PhraseLib.Lookup("term.conditions", LanguageID) & "</span></h2>")
                    Else
                        Send("<h2 style=""float:left;""><span><a href=""UEoffer-con.aspx?OfferID=" & OfferID & """>" & Copient.PhraseLib.Lookup("term.conditions", LanguageID) & "</a></span></h2>")
                    End If
                    Send_BoxResizer("conditionbody", "imgConditions", Copient.PhraseLib.Lookup("term.conditions", LanguageID), True)
                %>
                <div id="conditionbody">
                    <%
                        Dim icounter As Integer 'inclusion counter
                        Dim xcounter As Integer 'exclusion counter
                        ' Customer conditions
                        MyCommon.QueryStr = "select CG.CustomerGroupID, Name, ExcludedUsers, PhraseID, ICG.RequiredFromTemplate, " & _
                                            "CG.AnyCardholder, CG.AnyCustomer, CG.NewCardholders, CG.AnyCAMCardholder from CPE_IncentiveCustomerGroups as ICG with (NoLock) " & _
                                            "left join CustomerGroups as CG with (NoLock) on CG.CustomerGroupID=ICG.CustomerGroupID " & _
                                            "where RewardOptionID=" & roid & " and ICG.Deleted=0 and ExcludedUsers=0"
                        rst = MyCommon.LRT_Select
                        i = 1
                        GCount = 0
                        For Each row In rst.Rows
                            If (MyCommon.NZ(row.Item("CustomerGroupID"), 99999) <= 2) Then
                                AnyCustomerUsed = True
                            End If
                        Next
                        If (rst.Rows.Count > 0) Then
                            counter = counter + 1
                            Send("<h3>" & Copient.PhraseLib.Lookup("term.customerconditions", LanguageID) & "</h3>")
                            Send("<ul class=""condensed"">")
                            For Each row In rst.Rows
                                Sendb("<li>")
                                If IsDBNull(row.Item("CustomerGroupID")) Then
                                    Sendb(Copient.PhraseLib.Lookup("term.undefined", LanguageID))
                                Else
                                    If IsDBNull(row.Item("PhraseID")) Then
                                        If Popup Or Not m_AnalyticsCustomerGroups.HasValidExternalSegmentId(MyCommon.NZ(row.Item("CustomerGroupID"), "-1")) Then
                                            Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), "&nbsp;"), 25))
                                        Else
                                            Sendb("<a href=""../cgroup-edit.aspx?CustomerGroupID=" & MyCommon.NZ(row.Item("CustomerGroupID"), "-1") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), "&nbsp;"), 25) & "</a>")
                                        End If
                                    Else
                                        If (row.Item("PhraseID") = 0) Then
                                            If Popup Then
                                                Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(MyCommon.NZ(row.Item("Name"), ""), "&nbsp;"), 25))
                                            Else
                                                Sendb("<a href=""../cgroup-edit.aspx?CustomerGroupID=" & MyCommon.NZ(row.Item("CustomerGroupID"), "-1") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), "&nbsp;"), 25) & "</a>")
                                            End If
                                        Else
                                            Sendb(MyCommon.SplitNonSpacedString(Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID), 25))
                                        End If
                                    End If
                                End If
                                i = i + 1
                                MyCommon.QueryStr = "select count(*) as GCount from GroupMembership with (NoLock) where CustomerGroupID=" & MyCommon.NZ(row.Item("CustomerGroupID"), "-1") & " and Deleted = 0"
                                rst2 = MyCommon.LXS_Select()


                                For Each row2 In rst2.Rows
                                    If Not (row2.Item("GCount") = 0) AndAlso Not MyCommon.NZ(row.Item("AnyCardholder"), False) AndAlso Not MyCommon.NZ(row.Item("AnyCustomer"), False) AndAlso Not MyCommon.NZ(row.Item("NewCardholders"), False) AndAlso Not MyCommon.NZ(row.Item("AnyCAMCardholder"), False) Then
                                        Sendb(" (" & row2.Item("GCount") & ") ")
                                        If (icounter <= rst.Rows.Count - 2) Then
                                            Sendb(" or ")
                                        Else
                                            Sendb(" ")
                                        End If
                                    ElseIf (IsDBNull(row.Item("CustomerGroupID")) AndAlso MyCommon.NZ(row.Item("RequiredFromTemplate"), False)) Then
                                        Sendb(" <span class=""red"">")
                                        Sendb("<small>(" & Copient.PhraseLib.Lookup("term.requiredbytemplate", LanguageID) & ")</small>")
                                        Sendb("</span>")
                                    Else
                                        Sendb(" ")
                                    End If
                                    icounter = icounter + 1
                                Next
                                Sendb("</li>")
                            Next

                            ' Check for any display any excluded customer groups
                            MyCommon.QueryStr = "select CG.CustomerGroupID,Name,PhraseID,ExcludedUsers from CPE_IncentiveCustomerGroups as ICG with (NoLock) " & _
                                                "left join CustomerGroups as CG with (NoLock) on CG.CustomerGroupID=ICG.CustomerGroupID " & _
                                                "where RewardOptionID=" & roid & " and ICG.Deleted=0 and ExcludedUsers=1"
                            rst = MyCommon.LRT_Select
                            If rst.Rows.Count > 0 Then
                                Sendb(Copient.PhraseLib.Lookup("term.excluding", LanguageID) & " ")
                                For Each row In rst.Rows
                                    Sendb("<li>")
                                    If IsDBNull(row.Item("PhraseID")) Then
                                        If Popup Then
                                            Send(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25))
                                        Else
                                            Send("<a href=""../cgroup-edit.aspx?CustomerGroupID=" & MyCommon.NZ(row.Item("CustomerGroupID"), "-1") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
                                        End If
                                    Else
                                        If (row.Item("PhraseID") = 0) Then
                                            If Popup Then
                                                Send(MyCommon.SplitNonSpacedString(row.Item("Name"), 25))
                                            Else
                                                Send("<a href=""../cgroup-edit.aspx?CustomerGroupID=" & MyCommon.NZ(row.Item("CustomerGroupID"), "-1") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
                                            End If
                                        Else
                                            Send(MyCommon.SplitNonSpacedString(Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID), 25))
                                        End If
                                    End If
                                    MyCommon.QueryStr = "select count(*) as GCount from GroupMembership with (NoLock) where CustomerGroupID = " & MyCommon.NZ(MyCommon.NZ(row.Item("CustomerGroupID"), 0), "-1") & " And Deleted = 0"
                                    rst2 = MyCommon.LXS_Select()

                                    For Each row2 In rst2.Rows
                                        If (MyCommon.NZ(row.Item("CustomerGroupID"), -1) > 2) Then
                                            Sendb(" (" & row2.Item("GCount") & ") ")
                                            If (xcounter <= rst.Rows.Count - 2) Then
                                                Sendb(" and ")
                                            Else
                                                Sendb(" ")
                                            End If
                                        End If
                                        xcounter = xcounter + 1
                                    Next
                                    Sendb("</li>")
                                Next
                            End If
                            Dim CardTypeStr As String = ""
                            MyCommon.QueryStr = "select CardTypeID from CustomerConditionCardTypes where RewardOptionID=@roid"
                            MyCommon.DBParameters.Add("@roid", SqlDbType.BigInt).Value = roid
                            Dim rst1 As DataTable = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                            If (rst1.Rows.Count > 0) Then
                                For Each row1 As DataRow In rst1.Rows
                                    MyCommon.QueryStr = "select Description,PhraseTerm from CardTypes where CardTypeID=@CardTypeID"
                                    MyCommon.DBParameters.Add("@CardTypeID", SqlDbType.Int).Value = row1.Item("CardTypeID")
                                    rst2 = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)
                                    If (rst2.Rows.Count > 0) Then
                                        If CardTypeStr = "" Then
                                            CardTypeStr = Copient.PhraseLib.Lookup("term.cardtype", LanguageID) & " : " & Copient.PhraseLib.Lookup(MyCommon.NZ(rst2.Rows(0).Item("PhraseTerm"), ""), LanguageID)
                                        Else
                                            CardTypeStr &= " " & Copient.PhraseLib.Lookup("term.or", LanguageID) & " " & Copient.PhraseLib.Lookup(MyCommon.NZ(rst2.Rows(0).Item("PhraseTerm"), ""), LanguageID)
                                        End If
                                    End If
                                Next
                            End If
                            If CardTypeStr <> "" Then
                                Send("<li>")
                                Send(CardTypeStr)
                                Send("</li> ")
                            End If
                            Dim custApprovalResult As AMSResult(Of CustomerApproval) = m_CustomerConditionService.GetCustomerApprovalByROID(roid)
                            If custApprovalResult.ResultType = AMSResultType.Success AndAlso custApprovalResult.Result IsNot Nothing Then
                                Dim custApproval As CustomerApproval = custApprovalResult.Result
                                Dim custApprovalStr As String = ""
                                If custApproval.CustomerApprovalID > 0 Then
                                    Dim dtALTypes As DataTable = m_CustomerConditionService.GetCustomerApprovalLimitTypes(LanguageID)
                                    custApprovalStr = Copient.PhraseLib.Lookup("term.customerapproval", LanguageID) & " - " & Copient.PhraseLib.Lookup("term.approvallimit", LanguageID) & " : "
                                    If dtALTypes IsNot Nothing AndAlso dtALTypes.Rows.Count > 0 Then
                                        For Each row In dtALTypes.Rows
                                            If custApproval.ApprovalType = row.Item("ApprovalLimitTypeID") Then
                                                custApprovalStr &= row.Item("Phrase").ToString()
                                                Exit For
                                            End If
                                        Next
                                    End If
                                    'If custApproval.Approval = 1 Then
                                    '    custApprovalStr &= "Once for Offer"
                                    'ElseIf custApproval.Approval = 2 Then
                                    '    custApprovalStr &= "Each offer redemption"
                                    'End If
                                End If
                                If custApprovalStr <> "" Then
                                    Send("<li>")
                                    Send(custApprovalStr)
                                    Send("</li> ")
                                End If
                            End If
                            Send("</ul>")
                            Send("<br class=""half"" />")
                        End If

                        ' Attribute conditions
                        MyCommon.QueryStr = "select IA.IncentiveAttributeID, DisallowEdit, RequiredFromTemplate, RO.AttributeComboID, " & _
                                            "IAT.AttributeTypeID, AT.Description, IAT.AttributeValues " & _
                                            "from CPE_IncentiveAttributes as IA with (NoLock) " & _
                                            "left join CPE_IncentiveAttributeTiers as IAT with (NoLock) on IAT.IncentiveAttributeID=IA.IncentiveAttributeID " & _
                                            "left join CPE_RewardOptions as RO with (NoLock) on IA.RewardOptionID=RO.RewardOptionID " & _
                                            "left join AttributeTypes as AT with (NoLock) on AT.AttributeTypeID=IAT.AttributeTypeID " & _
                                            "where IA.RewardOptionID=" & roid & " and IA.Deleted=0;"
                        rst = MyCommon.LRT_Select
                        i = 1
                        GCount = 0
                        If (rst.Rows.Count > 0) Then
                            counter = counter + 1
                            Send("<h3>" & Copient.PhraseLib.Lookup("term.attributeconditions", LanguageID) & "</h3>")
                            Send("<ul class=""condensed"">")
                            For Each row In rst.Rows
                                ' spit out the AttributeComboID
                                If (i > 1) Then
                                    If (MyCommon.NZ(row.Item("AttributeComboID"), 0) = 0) Then
                                        ' single
                                    ElseIf (row.Item("AttributeComboID") = 1) Then
                                        ' and
                                        Send("&nbsp;<i>" & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & "</i>")
                                    Else
                                        ' or
                                        Send("&nbsp;<i>" & StrConv(Copient.PhraseLib.Lookup("term.or", LanguageID), VbStrConv.Lowercase) & "</i>")
                                    End If
                                End If
                                Sendb("<li>")
                                If IsDBNull(row.Item("AttributeTypeID")) Then
                                    Sendb(Copient.PhraseLib.Lookup("term.undefined", LanguageID))
                                Else
                                    Sendb("<a href=""../attribute-edit.aspx?AttributeTypeID=" & MyCommon.NZ(row.Item("AttributeTypeID"), "-1") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Description"), ""), 25) & "</a>")
                                End If
                                If (IsDBNull(row.Item("AttributeTypeID")) AndAlso MyCommon.NZ(row.Item("RequiredFromTemplate"), False)) Then
                                    Sendb(" <span class=""red"">")
                                    Sendb("<small>(" & Copient.PhraseLib.Lookup("term.requiredbytemplate", LanguageID) & ")</small>")
                                    Sendb("</span>")
                                End If

                                Send("</li>")
                                i = i + 1
                            Next
                            Send("</ul>")
                            Send("<br class=""half"" />")
                        End If


                        ' Product conditions
                        'If(bUseMultipleProductExclusionGroups) Then
                        ' 'Retreiving all included product groups.
                        '    MyCommon.QueryStr = "select PG.ProductGroupID,PG.buyerid, PG.Name, PG.PhraseID, UT.PhraseID, ExcludedProducts, ProductComboID, " & _
                        '                  " QtyForIncentive,QtyUnitType,AccumMin,AccumLimit,AccumPeriod,RequiredFromTemplate, IPG.IncentiveProductGroupID " & _
                        '                  " from CPE_IncentiveProductGroups as IPG with (NoLock) " & _
                        '                  " left join ProductGroups as PG with (NoLock) on PG.ProductGroupID=IPG.ProductGroupID " & _
                        '                  " left join CPE_RewardOptions as RO with (NoLock) on IPG.RewardOptionID=RO.RewardOptionId " & _
                        '                  " left join CPE_UnitTypes as UT with (NoLock)on UT.UnitTypeID=IPG.QtyUnitType " & _
                        '                  " where IPG.RewardOptionID=" & roid & " and IPG.Deleted=0 and Disqualifier=0 and ExcludedProducts=0;"
                        '    Dim dt1 As DataTable = MyCommon.LRT_Select
                        '      rst=dt1.Clone
                        '    If(dt1.Rows.Count>0)
                        '        For Each row In dt1.Rows
                        '             rst.ImportRow(row)
                        '            'Retrieving Excluding product group fro the corresponding included group
                        '             MyCommon.QueryStr = "select PG.ProductGroupID,PG.buyerid, PG.Name, PG.PhraseID, UT.PhraseID, ExcludedProducts, ProductComboID, " & _
                        '                  " QtyForIncentive,QtyUnitType,AccumMin,AccumLimit,AccumPeriod,RequiredFromTemplate, IPG.IncentiveProductGroupID " & _
                        '                  " from CPE_IncentiveProductGroups as IPG with (NoLock) " & _
                        '                  " left join ProductGroups as PG with (NoLock) on PG.ProductGroupID=IPG.ProductGroupID " & _
                        '                  " left join CPE_RewardOptions as RO with (NoLock) on IPG.RewardOptionID=RO.RewardOptionId " & _
                        '                  " left join CPE_UnitTypes as UT with (NoLock)on UT.UnitTypeID=IPG.QtyUnitType " & _
                        '                  " where IPG.RewardOptionID=" & roid & " and IPG.Deleted=0 and Disqualifier=0 and ExcludedProducts=1 and  IPG.InclusionIncentiveProductGroupSet=" & MyCommon.NZ(row.Item("IncentiveProductGroupID"), 0) & ";"                        
                        '               Dim dt2 As DataTable = MyCommon.LRT_Select
                        '            If(dt2.Rows.Count>0)
                        '                For Each row1 In dt2.Rows
                        '                    rst.ImportRow(row1) 'There exists an excluded group .Adding excluded group details of the corresponding included group  
                        '                 Next
                        '            End If
                        '        Next
                        '    End If
                        'Else
                        MyCommon.QueryStr = "select PG.ProductGroupID,PG.buyerid, PG.Name, PG.PhraseID, UT.PhraseID, ExcludedProducts, ProductComboID, " & _
                                            " QtyForIncentive,QtyUnitType,AccumMin,AccumLimit,AccumPeriod,RequiredFromTemplate, IPG.IncentiveProductGroupID " & _
                                            " from CPE_IncentiveProductGroups as IPG with (NoLock) " & _
                                            " left join ProductGroups as PG with (NoLock) on PG.ProductGroupID=IPG.ProductGroupID " & _
                                            " left join CPE_RewardOptions as RO with (NoLock) on IPG.RewardOptionID=RO.RewardOptionId " & _
                                            " left join CPE_UnitTypes as UT with (NoLock)on UT.UnitTypeID=IPG.QtyUnitType " & _
                                            " where IPG.RewardOptionID=" & roid & " and IPG.Deleted=0 and Disqualifier=0 " & _
                                            " Order by PG.Name;"
                        rst = MyCommon.LRT_Select
                        'End If

                        Dim isExcludedCond As Boolean = False
                        i = 1
                        GCount = 0
                        For Each row In rst.Rows
                            If (MyCommon.NZ(row.Item("ProductGroupID"), -1) = 1) Then
                                AnyProductUsed = True
                            End If
                        Next
                        If (rst.Rows.Count > 0) Then
                            counter = counter + 1
                            Send("<h3>" & Copient.PhraseLib.Lookup("term.productconditions", LanguageID) & "</h3>")
                            Send("<ul class=""condensed"">")
                            For Each row In rst.Rows

                                ' spit out the ProductComboID
                                If (i > 1 And MyCommon.NZ(row.Item("ExcludedProducts"), False) = False AndAlso isExcludedCond = False) Then
                                    If (MyCommon.NZ(row.Item("ProductComboID"), 0) = 0) Then
                                        ' single
                                    ElseIf (row.Item("ProductComboID") = 1) Then
                                        ' and
                                        Send("&nbsp;<i>" & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & "</i>")
                                    Else
                                        ' or
                                        Send("&nbsp;<i>" & StrConv(Copient.PhraseLib.Lookup("term.or", LanguageID), VbStrConv.Lowercase) & "</i>")
                                    End If
                                End If
                                'AMS-684      
                                'If exclusionPGList.Count > 0 Then
                                '    Sendb("<li>" & Copient.PhraseLib.Lookup("term.excluded", LanguageID) & ": ")
                                'Else
                                If MyCommon.NZ(row.Item("ExcludedProducts"), False) = False Then
                                    Sendb("<li>")
                                End If
                                'End If
                                If IsDBNull(row.Item("ProductGroupID")) Then
                                    Sendb(Copient.PhraseLib.Lookup("term.undefined", LanguageID))
                                Else
                                    If IsDBNull(row.Item("PhraseID")) Then
                                        If Popup Then
                                            Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25))
                                        Else
                                            If MyCommon.NZ(row.Item("ExcludedProducts"), False) = False Then
                                                If (MyCommon.IsEngineInstalled(9) And MyCommon.Fetch_UE_SystemOption(168) = "1" And Not IsDBNull(row.Item("Buyerid"))) Then
                                                    Dim buyerid As Integer = row.Item("Buyerid")
                                                    Dim externalBuyerid = MyCommon.GetExternalBuyerId(buyerid)
                                                    Sendb("<a href=""../pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("ProductGroupID"), "-1") & """>" & MyCommon.SplitNonSpacedString("Buyer " & externalBuyerid & " - " & MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
                                                Else
                                                    Sendb("<a href=""../pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("ProductGroupID"), "-1") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
                                                End If
                                            End If
                                        End If
                                    Else
                                        Sendb(MyCommon.SplitNonSpacedString(Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID), 25))
                                    End If
                                End If

                                If (IsDBNull(row.Item("ProductGroupID")) AndAlso MyCommon.NZ(row.Item("RequiredFromTemplate"), False)) Then
                                    Sendb(" <span class=""red"">")
                                    Sendb("<small>(" & Copient.PhraseLib.Lookup("term.requiredbytemplate", LanguageID) & ")</small>")
                                    Sendb("</span>")
                                End If

                                MyCommon.QueryStr = "select count(*) as GCount from ProdGroupItems with (NoLock) where ProductGroupID = " & MyCommon.NZ(row.Item("ProductGroupID"), "-1") & " And Deleted = 0"
                                rst2 = MyCommon.LRT_Select()
                                For Each row2 In rst2.Rows
                                    If (MyCommon.NZ(row.Item("ProductGroupID"), -1) > 1 And MyCommon.NZ(row.Item("ExcludedProducts"), False) = False) Then
                                        Send(" (" & row2.Item("GCount") & ")")
                                    End If
                                Next

                                ' QtyForIncentive,QtyUnitType,AccumMin,AccumLimit,AccumPeriod,UnitDescription
                                If (MyCommon.NZ(row.Item("QtyForIncentive"), 0) > 0 And Not MyCommon.NZ(row.Item("ExcludedProducts"), False)) Then
                                    Send("<br />")
                                    MyCommon.QueryStr = "select Quantity from CPE_IncentiveProductGroupTiers where IncentiveProductGroupID=" & MyCommon.NZ(row.Item("IncentiveProductGroupID"), 0) & " and RewardOptionID=" & roid & " order by TierLevel;"
                                    rstTiers = MyCommon.LRT_Select
                                    If rstTiers.Rows.Count > 0 Then
                                        t = 1
                                        For Each row4 In rstTiers.Rows
                                            Dim QuantityValue As Decimal
                                            QuantityValue = MyCommon.NZ(row4.Item("Quantity"), 0)
                                            Send(Localizer.Format_Qunatity(QuantityValue, roid, MyCommon.NZ(row.Item("QtyUnitType"), 0)))

                                            If t < TierLevels Then
                                                Sendb(" / ")
                                            End If
                                            t = t + 1
                                        Next
                                        If MyCommon.NZ(row.Item("QtyUnitType"), 0) <> 4 Then
                                            Send(" " & StrConv(Copient.PhraseLib.Lookup("term.required", LanguageID), VbStrConv.Lowercase))
                                        End If
                                    End If
                                    'AMS-684 Display if there are any associated excluded product groups
                                    DisplayExclusionGroups(m_ProductConditionPGService, row.Item("IncentiveProductGroupID"), infoMessage, MyCommon, LanguageID)
                                    If MyCommon.NZ(row.Item("AccumLimit"), 0) <> 0 OrElse MyCommon.NZ(row.Item("AccumPeriod"), 0) <> 0 OrElse MyCommon.NZ(row.Item("AccumMin"), 0) <> 0 Then
                                        ' There's at least some accumulation data set, so display it:
                                        ' Limit value
                                        If row.Item("AccumLimit") > 0 Then
                                            Sendb(Copient.PhraseLib.Lookup("term.limit", LanguageID) & " ")
                                            If MyCommon.NZ(row.Item("QtyUnitType"), 0) = 1 Then
                                                Sendb(row.Item("AccumLimit"))
                                            ElseIf MyCommon.NZ(row.Item("QtyUnitType"), 0) = 2 Then
                                                Sendb(FormatCurrency(row.Item("AccumLimit")))
                                            ElseIf (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 3) Then
                                                Sendb(row.Item("AccumLimit") & " " & Copient.PhraseLib.Lookup("term.lbsgals", LanguageID))
                                            End If
                                        Else
                                            Sendb(Copient.PhraseLib.Lookup("term.nolimit", LanguageID))
                                        End If
                                        ' Period value
                                        If row.Item("AccumPeriod") > 0 Then
                                            Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.every", LanguageID), VbStrConv.Lowercase) & " ")
                                            If row.Item("AccumPeriod") <= 1 Then
                                                Sendb(StrConv(Copient.PhraseLib.Lookup("term.day", LanguageID), VbStrConv.Lowercase))
                                            Else
                                                Sendb(row.Item("AccumPeriod") & " " & StrConv(Copient.PhraseLib.Lookup("term.days", LanguageID), VbStrConv.Lowercase))
                                            End If
                                        End If
                                        ' Minimum value
                                        If row.Item("AccumMin") > 0 Then
                                            Sendb(", " & StrConv(Copient.PhraseLib.Lookup("term.minimum", LanguageID), VbStrConv.Lowercase) & " ")
                                            If MyCommon.NZ(row.Item("QtyUnitType"), 0) = 1 Then
                                                Send(row.Item("AccumMin"))
                                            ElseIf MyCommon.NZ(row.Item("QtyUnitType"), 0) = 2 Then
                                                Send(FormatCurrency(row.Item("AccumMin")))
                                            ElseIf (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 3) Then
                                                Send(row.Item("AccumMin") & " " & Copient.PhraseLib.Lookup("term.lbsgals", LanguageID))
                                            End If
                                        Else
                                            Send(", " & StrConv(Copient.PhraseLib.Lookup("term.nominimum", LanguageID), VbStrConv.Lowercase))
                                        End If
                                    End If
                                End If
                                If MyCommon.NZ(row.Item("ExcludedProducts"), False) = False Then
                                    Send("</li>")
                                End If
                                If (i = 1 And MyCommon.NZ(row.Item("ExcludedProducts"), False) = True) Then
                                    'During Offer translation from CM to UE ,when excluded product condition is retreieved as first row in datatable ,extra "and" is dispalyed in product condition section of the Summary page
                                    isExcludedCond = True
                                Else
                                    isExcludedCond = False
                                End If
                                i = i + 1
                            Next
                            Send("</ul>")
                            Send("<br class=""half"" />")
                        End If

                        ' Product disqualifiers
                        MyCommon.QueryStr = "select PG.ProductGroupID, PG.Name, PG.PhraseID, UT.PhraseID, ExcludedProducts, ProductComboID, " & _
                                            " QtyForIncentive, QtyUnitType, AccumMin, AccumLimit, AccumPeriod, RequiredFromTemplate, IPG.IncentiveProductGroupID " & _
                                            " from CPE_IncentiveProductGroups as IPG with (NoLock) " & _
                                            " left join ProductGroups as PG with (NoLock) on PG.ProductGroupID=IPG.ProductGroupID " & _
                                            " left join CPE_RewardOptions as RO with (NoLock) on IPG.RewardOptionID=RO.RewardOptionId " & _
                                            " left join CPE_UnitTypes as UT with (NoLock)on UT.UnitTypeID=IPG.QtyUnitType " & _
                                            " where IPG.RewardOptionID=" & roid & " and IPG.Deleted=0 and Disqualifier=1;"
                        rst = MyCommon.LRT_Select
                        i = 1
                        GCount = 0
                        For Each row In rst.Rows
                            If (MyCommon.NZ(row.Item("ProductGroupID"), -1) = 1) Then
                                AnyProductUsed = True
                            End If
                        Next
                        If (rst.Rows.Count > 0) Then
                            counter = counter + 1
                            Send("<h3>" & Copient.PhraseLib.Lookup("term.productdisqualifiers", LanguageID) & "</h3>")
                            Send("<ul class=""condensed"">")
                            For Each row In rst.Rows
                                ' spit out the ProductComboID
                                If (i > 1) Then
                                    Send("&nbsp;<i>" & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & "</i>")
                                End If
                                Sendb("<li>")
                                If IsDBNull(row.Item("PhraseID")) Then
                                    If Popup Then
                                        Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25))
                                    Else
                                        Sendb("<a href=""../pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("ProductGroupID"), "-1") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
                                    End If
                                ElseIf (IsDBNull(row.Item("ProductGroupID")) AndAlso MyCommon.NZ(row.Item("RequiredFromTemplate"), False)) Then
                                    Sendb(" <span class=""red"">")
                                    Sendb("<small>(" & Copient.PhraseLib.Lookup("term.requiredbytemplate", LanguageID) & ")</small>")
                                    Sendb("</span>")
                                Else
                                    Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), "&nbsp;"), 25))
                                End If

                                MyCommon.QueryStr = "select count(*) as GCount from ProdGroupItems with (NoLock) where ProductGroupID = " & MyCommon.NZ(row.Item("ProductGroupID"), "-1") & " And Deleted = 0"
                                rst2 = MyCommon.LRT_Select()
                                For Each row2 In rst2.Rows
                                    If (MyCommon.NZ(row.Item("ProductGroupID"), -1) > 1) Then
                                        Send(" (" & row2.Item("GCount") & ")")
                                    End If
                                Next

                                ' QtyForIncentive,QtyUnitType,AccumMin,AccumLimit,AccumPeriod,UnitDescription
                                If (MyCommon.NZ(row.Item("QtyForIncentive"), 0) > 0 And Not MyCommon.NZ(row.Item("ExcludedProducts"), False)) Then
                                    Send("<br />")
                                    MyCommon.QueryStr = "select Quantity from CPE_IncentiveProductGroupTiers where IncentiveProductGroupID=" & MyCommon.NZ(row.Item("IncentiveProductGroupID"), 0) & " and RewardOptionID=" & roid & " order by TierLevel;"
                                    rstTiers = MyCommon.LRT_Select
                                    If rstTiers.Rows.Count > 0 Then
                                        t = 1
                                        For Each row4 In rstTiers.Rows
                                            Localizer.Format_Qunatity(row4.Item("Quantity"), roid, MyCommon.NZ(row.Item("QtyUnitType"), 0))
                                            t = t + 1
                                        Next
                                        Send(" " & StrConv(Copient.PhraseLib.Lookup("term.required", LanguageID), VbStrConv.Lowercase))
                                    End If
                                End If
                                Send("</li>")
                                i = i + 1
                            Next
                            Send("</ul>")
                            Send("<br class=""half"" />")
                        End If

                        'Points conditions
                        MyCommon.QueryStr = "Select IPG.ProgramID, ProgramName, QtyForIncentive, RequiredFromTemplate, IncentivePointsID " & _
                                            "from CPE_IncentivePointsGroups as IPG with (NoLock) " & _
                                            "left join PointsPrograms as PP with (NoLock) on PP.ProgramID=IPG.ProgramID " & _
                                            "where IPG.Deleted=0 and RewardOptionID=" & roid & ";"
                        rst = MyCommon.LRT_Select
                        i = 1
                        If (rst.Rows.Count > 0) Then
                            counter = counter + 1
                            Send("<h3>" & Copient.PhraseLib.Lookup("term.pointsconditions", LanguageID) & "</h3>")
                            Send("<ul class=""condensed"">")
                            For Each row In rst.Rows
                                Sendb("<li>")
                                MyCommon.QueryStr = "select Quantity from CPE_IncentivePointsGroupTiers " & _
                                                    "where IncentivePointsID=" & MyCommon.NZ(row.Item("IncentivePointsID"), 0) & " order by TierLevel;"
                                rstTiers = MyCommon.LRT_Select
                                If rstTiers.Rows.Count > 0 Then
                                    t = 1
                                    For Each row4 In rstTiers.Rows
                                        Sendb(MyCommon.NZ(row4.Item("Quantity"), 0))
                                        If t < TierLevels Then
                                            Sendb(" / ")
                                        End If
                                        t = t + 1
                                    Next
                                End If
                                If (MyCommon.NZ(row.Item("ProgramID"), -1) > 0) Then
                                    If Popup Then
                                        Sendb(" " & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("ProgramName"), ""), 25))
                                    Else
                                        Sendb(" <a href=""../point-edit.aspx?ProgramGroupID=" & MyCommon.NZ(row.Item("ProgramID"), "") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("ProgramName"), ""), 25) & "</a>")
                                    End If
                                ElseIf (IsDBNull(row.Item("ProgramID")) AndAlso MyCommon.NZ(row.Item("RequiredFromTemplate"), False)) Then
                                    Sendb(" <span class=""red"">")
                                    Sendb("<small>(" & Copient.PhraseLib.Lookup("term.requiredbytemplate", LanguageID) & ")</small>")
                                    Sendb("</span>")
                                Else
                                    Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("ProgramName"), ""), 25))
                                End If
                                Send("</li>")
                            Next
                            Send("</ul>")
                            Send("<br class=""half"" />")
                        End If

                        'Store value conditions
                        MyCommon.QueryStr = "select ISVP.SVProgramID, SVP.Name, SVP.Value, SVP.SVTypeID, SVT.ValuePrecision, QtyForIncentive, RequiredFromTemplate,IncentiveStoredValueID from CPE_IncentiveStoredValuePrograms as ISVP with (NoLock) " & _
                                          "left join StoredValuePrograms as SVP with (NoLock) on SVP.SVProgramID=ISVP.SVProgramID " & _
                                          "left join SVTypes as SVT with (NoLock) on SVP.SVTypeID=SVT.SVTypeID " & _
                                          "where ISVP.Deleted=0 and RewardOptionID=" & roid
                        rst = MyCommon.LRT_Select
                        i = 1
                        If (rst.Rows.Count > 0) Then
                            counter = counter + 1
                            Send("<h3>" & Copient.PhraseLib.Lookup("term.storedvalueconditions", LanguageID) & "</h3>")
                            Send("<ul class=""condensed"">")
                            For Each row In rst.Rows
                                Sendb("<li>")
                                MyCommon.QueryStr = "select Quantity from CPE_IncentiveStoredValueProgramTiers where RewardOptionID=" & roid & "and IncentiveStoredValueID=" & MyCommon.NZ(row.Item("IncentiveStoredValueID"), 0) & " order by TierLevel;"
                                rstTiers = MyCommon.LRT_Select
                                If rstTiers.Rows.Count > 0 Then
                                    t = 1
                                    For Each row4 In rstTiers.Rows
                                        Sendb(CInt(MyCommon.NZ(row4.Item("Quantity"), "0")))
                                        If MyCommon.NZ(row.Item("SVTypeID"), 0) > 1 Then
                                            Sendb(" ($" & Math.Round(CDec(MyCommon.NZ(row.Item("Value"), 0) * MyCommon.NZ(row4.Item("Quantity"), 0)), CInt(MyCommon.NZ(row.Item("ValuePrecision"), 0))).ToString(MyCommon.GetAdminUser.Culture) & ")")
                                        End If
                                        If t < TierLevels Then
                                            Sendb(" / ")
                                        End If
                                        t = t + 1
                                    Next
                                End If
                                If (MyCommon.NZ(row.Item("SVProgramID"), -1) > 0) Then
                                    If Popup Then
                                        Sendb(" " & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25))
                                    Else
                                        Sendb(" <a href=""../SV-edit.aspx?ProgramGroupID=" & MyCommon.NZ(row.Item("SVProgramID"), "") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
                                    End If
                                ElseIf (IsDBNull(row.Item("SVProgramID")) AndAlso MyCommon.NZ(row.Item("RequiredFromTemplate"), False)) Then
                                    Sendb(" <span class=""red"">")
                                    Sendb("<small>(" & Copient.PhraseLib.Lookup("term.requiredbytemplate", LanguageID) & ")</small>")
                                    Sendb("</span>")
                                Else
                                    Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25))
                                End If
                                Send("</li>")
                            Next
                            Send("</ul>")
                            Send("<br class=""half"" />")
                        End If

                        'Trackable Coupon conditions
                        Dim objResult As AMSResult(Of List(Of TCProgramCondition)) = m_TrackableCouponCondition.GetTCProgramConditions(OfferID, EngineID)
                        If (objResult.ResultType <> AMSResultType.Success) Then
                            infoMessage = objResult.GetLocalizedMessage(LanguageID)
                        ElseIf (objResult.Result.Count > 0) Then
                            Send("<h3>" & Copient.PhraseLib.Lookup("term.trackablecouponcondition", LanguageID) & "</h3>")
                            Send("<ul class=""condensed"">")
                            For Each tcpcondition As TCProgramCondition In objResult.Result
                                Sendb("<li>")
                                If (tcpcondition.TCProgram IsNot Nothing) Then
                                    If Popup Then
                                        Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(tcpcondition.TCProgram.Name, ""), 25))
                                    Else
                                        Sendb("<a href=""../tcp-edit.aspx?tcprogramid=" & MyCommon.NZ(tcpcondition.TCProgram.ProgramID, "") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(tcpcondition.TCProgram.Name, ""), 25) & "</a>")
                                    End If
                                End If
                                Send("</li>")
                            Next
                            Send("</ul>")
                            Send("<br class=""half"" />")
                        End If


                        'Tender conditions
                        MyCommon.QueryStr = "select ITT.IncentiveTenderID, ITT.TenderTypeID, ITT.Value, TT.Name, RequiredFromTemplate, RO.ExcludedTender, RO.ExcludedTenderAmtRequired " & _
                                            "from CPE_IncentiveTenderTypes as ITT with (NoLock) " & _
                                            "left join CPE_TenderTypes as TT with (NoLock) on TT.TenderTypeID=ITT.TenderTypeID " & _
                                            "left join CPE_RewardOptions as RO with (NoLock) on RO.RewardOptionID=ITT.RewardOptionID " & _
                                            "where ITT.Deleted=0 and ITT.RewardOptionID=" & roid & ";"
                        rst = MyCommon.LRT_Select
                        i = 1
                        If (rst.Rows.Count > 0) Then
                            counter = counter + 1
                            Send("<h3>" & Copient.PhraseLib.Lookup("term.tenderconditions", LanguageID) & "</h3>")
                            Send("<ul class=""condensed"">")
                            For Each row In rst.Rows
                                If TenderList <> "" Then
                                    TenderList &= ", "
                                End If
                                If TenderValue <> "" Then
                                    TenderValue &= ", "
                                End If
                                TenderList &= MyCommon.NZ(row.Item("Name"), "")
                                TenderValue &= Localizer.FormatCurrency_ForOffer(CDec(MyCommon.NZ(row.Item("Value"), 0)), roid)
                                TenderRequired = MyCommon.NZ(row.Item("RequiredFromTemplate"), False)
                                TenderExcluded = MyCommon.NZ(row.Item("ExcludedTender"), False)
                                TenderExcludedAmt = Localizer.Round_Currency(MyCommon.NZ(row.Item("ExcludedTenderAmtRequired"), 0), roid)
                            Next
                            If TenderExcluded Then
                                Sendb("<li>")
                                Sendb(Localizer.FormatCurrency_ForOffer(CDec(TenderExcludedAmt), roid) & " " & StrConv(Copient.PhraseLib.Lookup("term.from", LanguageID), VbStrConv.Lowercase) & " " & StrConv(Copient.PhraseLib.Lookup("term.allbut", LanguageID), VbStrConv.Lowercase) & " ")
                                Sendb("<a href=""../tender-engines.aspx"">" & MyCommon.SplitNonSpacedString(TenderList, 25) & "</a>")
                                Send("</li>")
                            Else
                                For Each row In rst.Rows
                                    Sendb("<li>")
                                    MyCommon.QueryStr = "select Value from CPE_IncentiveTenderTypeTiers where RewardOptionID=" & roid & " and IncentiveTenderID=" & MyCommon.NZ(row.Item("IncentiveTenderID"), 0) & ";"
                                    rstTiers = MyCommon.LRT_Select
                                    If rstTiers.Rows.Count > 0 Then
                                        t = 1
                                        For Each row4 In rstTiers.Rows
                                            Sendb(Localizer.FormatCurrency_ForOffer(CDec(MyCommon.NZ(row4.Item("Value"), "0")), roid))
                                            If t < TierLevels Then
                                                Sendb(" / ")
                                            End If
                                            t = t + 1
                                        Next
                                    End If
                                    Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.required", LanguageID), VbStrConv.Lowercase) & " " & StrConv(Copient.PhraseLib.Lookup("term.from", LanguageID), VbStrConv.Lowercase) & " ")
                                    If (MyCommon.NZ(row.Item("ExcludedTender"), False) = True) Then
                                        Sendb(StrConv(Copient.PhraseLib.Lookup("term.allbut", LanguageID), VbStrConv.Lowercase) & " ")
                                    End If
                                    Send("<a href=""../tender-engines.aspx"">" & MyCommon.NZ(row.Item("Name"), "") & "</a></li>")
                                Next
                            End If
                            Send("</ul>")
                            Send("<br class=""half"" />")
                        End If

                        ' Days conditions
                        MyCommon.QueryStr = "select DOWID, PhraseID from CPE_DaysOfWeek DW with (NoLock)"
                        rst = MyCommon.LRT_Select
                        MyCommon.QueryStr = "select DOWID from CPE_IncentiveDOW with (NoLock) where Deleted=0 and IncentiveID=" & OfferID
                        rst2 = MyCommon.LRT_Select
                        For Each row In rst.Rows
                            If rst2.Rows.Count >= 7 Then
                                Days = Copient.PhraseLib.Lookup("term.everyday", LanguageID)
                            Else
                                For Each row2 In rst2.Rows
                                    If (MyCommon.NZ(row2.Item("DOWID"), 0) = MyCommon.NZ(row.Item("DOWID"), 0)) Then
                                        If (Days = "") Then
                                            Days = Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID)
                                        Else
                                            Days = Days & ", " & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID)
                                        End If
                                    End If
                                Next
                            End If
                        Next
                        If (Days.Trim.Length > 0) Then
                            counter = counter + 1
                            Send("<h3>" & Copient.PhraseLib.Lookup("term.dayconditions", LanguageID) & "</h3>")
                            Send("<ul class=""condensed"">")
                            Send("<li>" & Days & "</li>")
                            Send("</ul>")
                        End If

                        ' Time conditions
                        MyCommon.QueryStr = "select StartTime, EndTime from CPE_IncentiveTOD with (NoLock) where IncentiveID=" & OfferID
                        rst = MyCommon.LRT_Select
                        If (rst.Rows.Count > 0) Then
                            For i = 0 To rst.Rows.Count - 1
                                If (i > 0) Then Times &= "; "
                                Times &= MyCommon.NZ(rst.Rows(i).Item("StartTime"), "") & " - " & MyCommon.NZ(rst.Rows(i).Item("EndTime"), "")
                            Next
                        End If
                        If (Times.Trim.Length > 0) Then
                            counter = counter + 1
                            Send("<h3>" & Copient.PhraseLib.Lookup("term.timeconditions", LanguageID) & "</h3>")
                            Send("<ul class=""condensed"">")
                            Send("<li>" & Times & "</li>")
                            Send("</ul>")
                        End If

                        'Instant win conditions
                        MyCommon.QueryStr = "select IncentiveInstantWinID,OddsOfWinning,NumPrizesAllowed,RandomWinners,RequiredFromTemplate,Unlimited " & _
                                            "from CPE_IncentiveInstantWin with (NoLock) " & _
                                            "where Deleted=0 and RewardOptionID=" & roid & ";"
                        rst = MyCommon.LRT_Select
                        i = 1
                        If (rst.Rows.Count > 0) Then
                            counter = counter + 1
                            Send("<h3>" & Copient.PhraseLib.Lookup("term.instantwinconditions", LanguageID) & "</h3>")
                            Send("<ul class=""condensed"">")
                            For Each row In rst.Rows
                                Sendb("<li>")
                                Sendb("1:" & MyCommon.NZ(row.Item("OddsOfWinning"), "0") & " ")
                                If MyCommon.NZ(row.Item("RandomWinners"), False) Then
                                    Sendb(StrConv(Copient.PhraseLib.Lookup("term.random", LanguageID), VbStrConv.Lowercase) & " ")
                                Else
                                    Sendb(StrConv(Copient.PhraseLib.Lookup("term.fixed", LanguageID), VbStrConv.Lowercase) & " ")
                                End If
                                Sendb(StrConv(Copient.PhraseLib.Lookup("term.odds", LanguageID), VbStrConv.Lowercase) & " " & StrConv(Copient.PhraseLib.Lookup("term.on", LanguageID), VbStrConv.Lowercase) & " ")
                                Sendb(IIf(MyCommon.NZ(row.Item("Unlimited"), False), StrConv(Copient.PhraseLib.Lookup("term.unlimited", LanguageID), VbStrConv.Lowercase), MyCommon.NZ(row.Item("NumPrizesAllowed"), "?")) & " ")
                                Sendb(StrConv(Copient.PhraseLib.Lookup("term.prizes", LanguageID), VbStrConv.Lowercase))
                                If (MyCommon.NZ(row.Item("RequiredFromTemplate"), False)) Then
                                    Sendb(" <span class=""red"">")
                                    Sendb("<small>(" & Copient.PhraseLib.Lookup("term.requiredbytemplate", LanguageID) & ")</small>")
                                    Sendb("</span>")
                                End If
                                Send("</li>")
                            Next
                            Send("</ul>")
                            Send("<br class=""half"" />")
                        End If

                        'Trigger code (aka PLU) conditions
                        MyCommon.QueryStr = "select IncentivePLUID, PLU, PerRedemption, CashierMessage, RequiredFromTemplate, PLUQuantity " &
                                            "from CPE_IncentivePLUs with (NoLock) " &
                                            "where RewardOptionID=" & roid & " order by PLU;"
                        rst = MyCommon.LRT_Select
                        i = 1
                        If (rst.Rows.Count > 0) Then
                            counter = counter + 1
                            Send("<h3>" & Copient.PhraseLib.Lookup("term.triggercodeconditions", LanguageID) & "</h3>")
                            Send("<ul class=""condensed"">")
                            For Each row In rst.Rows
                                Sendb("<li>")
                                If MyCommon.NZ(row.Item("PLU"), "") <> "" Then
                                    Sendb(MyCommon.NZ(row.Item("PLU"), Copient.PhraseLib.Lookup("term.undefined", LanguageID)))
                                Else
                                    Sendb(Copient.PhraseLib.Lookup("term.undefined", LanguageID))
                                End If
                                Sendb("<br/>")
                                Sendb(MyCommon.NZ(IIf(row.Item("PLUQuantity").ToString() <> "", row.Item("PLUQuantity"), 1).ToString() + " " + Copient.PhraseLib.Lookup("term.required", LanguageID), Copient.PhraseLib.Lookup("term.undefined", LanguageID)))
                                If (MyCommon.NZ(row.Item("RequiredFromTemplate"), False)) Then
                                    Sendb(" <span class=""red"">")
                                    Sendb("<small>(" & Copient.PhraseLib.Lookup("term.requiredbytemplate", LanguageID) & ")</small>")
                                    Sendb("</span>")
                                End If
                                Send("</li>")
                            Next
                            Send("</ul>")
                            Send("<br class=""half"" />")
                        End If

                        'Enterprise instant win conditions
                        MyCommon.QueryStr = "select IncentiveEIWID, NumberOfPrizes, EIW.FrequencyID, EIWF.Description, DisallowEdit, RequiredFromTemplate " & _
                                            "from CPE_IncentiveEIW as EIW with (NoLock) " & _
                                            "inner join CPE_IncentiveEIWFrequency as EIWF on EIWF.FrequencyID=EIW.FrequencyID " & _
                                            "where RewardOptionID=" & roid & ";"
                        rst = MyCommon.LRT_Select
                        i = 1
                        If (rst.Rows.Count > 0) Then
                            counter = counter + 1
                            Send("<h3>" & Copient.PhraseLib.Lookup("term.epriseinstantwinconditions", LanguageID) & "</h3>")
                            Send("<ul class=""condensed"">")
                            For Each row In rst.Rows
                                Send("  <li>")
                                MyCommon.QueryStr = "select count(*) as TriggerCount from CPE_EIWTriggers where IncentiveEIWID=" & MyCommon.NZ(row.Item("IncentiveEIWID"), 0) & " and RewardOptionID=" & roid & " and Removed=0;"
                                rst2 = MyCommon.LRT_Select
                                If rst2.Rows(0).Item("TriggerCount") = 0 Then
                                    Sendb("    " & Copient.PhraseLib.Lookup("term.no", LanguageID))
                                Else
                                    Sendb("    " & MyCommon.NZ(rst2.Rows(0).Item("TriggerCount"), 0))
                                End If
                                If rst2.Rows(0).Item("TriggerCount") = 1 Then
                                    Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.trigger", LanguageID), VbStrConv.Lowercase))
                                Else
                                    Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.triggers", LanguageID), VbStrConv.Lowercase))
                                End If
                                Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.for", LanguageID), VbStrConv.Lowercase))
                                Sendb(" " & MyCommon.NZ(row.Item("NumberOfPrizes"), 0))
                                If MyCommon.NZ(row.Item("NumberOfPrizes"), 0) = 1 Then
                                    Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.prize", LanguageID), VbStrConv.Lowercase))
                                Else
                                    Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.prizes", LanguageID), VbStrConv.Lowercase))
                                End If
                                Send(" " & StrConv(MyCommon.NZ(row.Item("Description"), ""), VbStrConv.Lowercase))
                                Send("  </li>")
                            Next
                            Send("</ul>")
                            Send("<br class=""half"" />")
                        End If


                        'Preference conditions
                        If MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER) Then
                            MyCommon.QueryStr = "select CIP.PreferenceID, RO.PreferenceComboID, CIP.IncentivePrefsID from CPE_IncentivePrefs as CIP with (NoLock) " & _
                                                "inner join CPE_RewardOptions as RO with (NoLock) on RO.RewardOptionID = CIP.RewardOptionID " & _
                                                "where CIP.RewardOptionID = " & roid
                            rst = MyCommon.LRT_Select
                            i = 1
                            If (rst.Rows.Count > 0) Then
                                counter = counter + 1
                                Send("<h3>" & Copient.PhraseLib.Lookup("term.preferenceconditions", LanguageID) & "</h3>")
                                Send("<ul class=""condensed"">")
                                For Each row In rst.Rows
                                    Send("  <li>")
                                    Send_Preference_Details(MyCommon, MyCommon.NZ(row.Item("PreferenceID"), 0))
                                    Send_Preference_Info(MyCommon, MyCommon.NZ(row.Item("IncentivePrefsID"), 0), roid, TierLevels)

                                    If i < rst.Rows.Count Then
                                        Send(" <br /><i>" & Copient.PhraseLib.Lookup("term." & IIf(MyCommon.NZ(row.Item("PreferenceComboID"), 1) = 1, "and", "or"), LanguageID).ToLower & "</i> ")
                                    End If

                                    Send("  </li>")
                                    i += 1
                                Next
                                Send("</ul>")
                                Send("<br class=""half"" />")
                            End If
                        End If


                        If (counter = 0) Then
                            Send("<h3>" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "</h3>")
                        End If
                        counter = 0
                    %>
                    <hr class="hidden" />
                </div>
            </div>
            <div class="box" id="offerrewards">
                <%  
                    If (LinksDisabled) Then
                        Send("<h2 style=""float:left;""><span>" & Copient.PhraseLib.Lookup("term.rewards", LanguageID) & "</span></h2>")
                    Else
                        Send("<h2 style=""float:left;""><span><a href=""UEoffer-rew.aspx?OfferID=" & OfferID & """>" & Copient.PhraseLib.Lookup("term.rewards", LanguageID) & "</a></span></h2>")
                    End If
                    Send_BoxResizer("rewardbody", "imgRewards", Copient.PhraseLib.Lookup("term.rewards", LanguageID), True)
                %>
                <div id="rewardbody">
                    <%
                        ' Discount reward
                        MyCommon.QueryStr = "select DISC.DiscountID, DISC.DiscountTypeID, DISC.Name, DISC.DiscountTypeId, DISC.ReceiptDescription, DISC.DiscountBarcode, DISC.VoidBarcode, " & _
                                            "DISC.DiscountAmount, DISC.DiscountedProductGroupID as SelectedPG, DISC.ItemLimit, DISC.WeightLimit, DISC.DollarLimit, DISC.ExcludedProductGroupID as ExcludedPG, " & _
                                            "DISC.DiscountAmount, DISC.ChargebackDeptID, DISC.AmountTypeID, DISC.L1Cap, DISC.L2DiscountAmt, DISC.L2AmountTypeID, DISC.L2Cap, DISC.L3DiscountAmt, DISC.L3AmountTypeID, " & _
                                            "DISC.DecliningBalance, DISC.RetroactiveDiscount, DISC.UserGroupID, DISC.BestDeal, DISC.AllowNegative, DISC.ComputeDiscount, DISC.SVProgramID, " & _
                                            "D.DeliverableID, AT.AmountTypeID, AT.PhraseID as AmountPhraseID, DT.PhraseID as DiscountPhraseID " & _
                                            "from CPE_Deliverables D with (NoLock) inner join CPE_Discounts DISC with (NoLock) on D.OutputID = DISC.DiscountID " & _
                                            "left join CPE_AmountTypes AT with (NoLock) on AT.AmountTypeID = DISC.AmountTypeID " & _
                                            "left join CPE_DiscountTypes DT with (NoLock) on DT.DiscountTypeID = DISC.DiscountTypeID " & _
                                            "where D.Deleted = 0 and DISC.Deleted = 0 and D.RewardOptionPhase=3 and D.RewardOptionID=" & roid & " and D.DeliverableTypeID=2;"
                        rst = MyCommon.LRT_Select()
                        If (rst.Rows.Count > 0) Then
                            counter = counter + 1
                            Dim Details As StringBuilder
                            Send("<h3>" & Copient.PhraseLib.Lookup("term.discounts", LanguageID) & "</h3>")
                            Send("<ul class=""condensed"">")
                            For Each row In rst.Rows
                                Details = New StringBuilder(200)
                                MyCommon.QueryStr = "select DiscountAmount, ItemLimit, WeightLimit, DollarLimit from CPE_DiscountTiers where DiscountID=" & MyCommon.NZ(row.Item("DiscountID"), 0) & ";"
                                rstTiers = MyCommon.LRT_Select
                                t = 1
                                AmountTypeID = (MyCommon.NZ(row.Item("AmountTypeID"), 0))
                                Select Case AmountTypeID
                                    Case 1, 5, 9, 10, 11, 12
                                        If rstTiers.Rows.Count > 0 Then
                                            For Each row4 In rstTiers.Rows
                                                Details.Append(Localizer.GetCached_Currency_Symbol(roid) & Math.Round(CDec(MyCommon.NZ(row4.Item("DiscountAmount"), 0)), Localizer.GetCached_Currency_Precision(roid)).ToString(MyCommon.GetAdminUser.Culture))
                                                If t < rstTiers.Rows.Count Then
                                                    Details.Append(" / ")
                                                End If
                                                t += 1
                                            Next
                                            Details.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.off", LanguageID), VbStrConv.Lowercase) & "&nbsp;")
                                        End If
                                    Case 3
                                        If rstTiers.Rows.Count > 0 Then
                                            For Each row4 In rstTiers.Rows
                                                Details.Append(Math.Round(CDec(MyCommon.NZ(row4.Item("DiscountAmount"), 0)), 2).ToString(MyCommon.GetAdminUser.Culture) & "% ")
                                                If t < rstTiers.Rows.Count Then
                                                    Details.Append(" / ")
                                                End If
                                                t += 1
                                            Next
                                            Details.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.off", LanguageID), VbStrConv.Lowercase) & "&nbsp;")
                                        End If
                                    Case 4
                                        Details.Append(Copient.PhraseLib.Lookup("term.free", LanguageID) & "&nbsp;")
                                    Case 2, 6, 13, 14, 15, 16
                                        If rstTiers.Rows.Count > 0 Then
                                            For Each row4 In rstTiers.Rows
                                                Details.Append(Localizer.GetCached_Currency_Symbol(roid) & Math.Round(CDec(MyCommon.NZ(row4.Item("DiscountAmount"), 0)), Localizer.GetCached_Currency_Precision(roid)).ToString(MyCommon.GetAdminUser.Culture) & "&nbsp;")
                                                If t < rstTiers.Rows.Count Then
                                                    Details.Append(" / ")
                                                End If
                                                t += 1
                                            Next
                                        End If
                                    Case 7
                                        MyCommon.QueryStr = "select SVProgramID, Name from StoredValuePrograms with (NoLock) where SVProgramID=" & MyCommon.NZ(row.Item("SVProgramID"), 0) & ";"
                                        rst3 = MyCommon.LRT_Select
                                        If (rst3.Rows.Count > 0) Then
                                            Details.Append(Copient.PhraseLib.Lookup("term.storedvalue", LanguageID) & " (<a href=""../sv-edit.aspx?ProgramGroupID=" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("SVProgramID"), 0), 25) & """>" & MyCommon.NZ(rst3.Rows(0).Item("Name"), "") & "</a>) " & StrConv(Copient.PhraseLib.Lookup("term.for", LanguageID), VbStrConv.Lowercase) & " ")
                                        End If
                                    Case 8
                                        Details.Append(Copient.PhraseLib.Lookup("term.specialpricing", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.for", LanguageID), VbStrConv.Lowercase) & " ")
                                    Case Else
                                        If rstTiers.Rows.Count > 0 Then
                                            For Each row4 In rstTiers.Rows
                                                Details.Append(Math.Round(CDec(MyCommon.NZ(row4.Item("DiscountAmount"), 0)), Localizer.GetCached_Currency_Precision(roid)).ToString(MyCommon.GetAdminUser.Culture) & "&nbsp;")
                                                If t < rstTiers.Rows.Count Then
                                                    Details.Append(" / ")
                                                End If
                                                t += 1
                                            Next
                                        End If
                                End Select
                
                                If (MyCommon.NZ(row.Item("DiscountTypeID"), 0) = 4 OrElse MyCommon.NZ(row.Item("DiscountTypeID"), 0) = 5) AndAlso MyCommon.NZ(row.Item("SelectedPG"), 0) = 0 Then
                                    Details.Append(StrConv(Copient.PhraseLib.Lookup("term.conditionalproducts", LanguageID), VbStrConv.Lowercase))
                                ElseIf MyCommon.NZ(row.Item("SelectedPG"), 0) = 0 Then
                                    Details.Append(StrConv(Copient.PhraseLib.Lookup("term.nothing", LanguageID), VbStrConv.Lowercase))
                                Else
                                    MyCommon.QueryStr = "select Name,buyerid from ProductGroups with (NoLock) where ProductGroupID = " & row.Item("SelectedPG")
                                    rst3 = MyCommon.LRT_Select()
                                    For Each row3 In rst3.Rows
                                        If MyCommon.NZ(row.Item("SelectedPG"), 0) = 1 Then
                                            Details.Append(StrConv(MyCommon.NZ(row3.Item("Name"), ""), VbStrConv.Lowercase))
                                        Else
                                            If Popup Then
                                                Details.Append(MyCommon.SplitNonSpacedString(MyCommon.NZ(row3.Item("Name"), ""), 25))
                                            Else
                                                If (MyCommon.IsEngineInstalled(9) And MyCommon.Fetch_UE_SystemOption(168) = "1" And Not IsDBNull(row3.Item("Buyerid"))) Then
                                                    Dim buyerid As Integer = row3.Item("Buyerid")
                                                    Dim externalBuyerid = MyCommon.GetExternalBuyerId(buyerid)
                                                    Details.Append("<a href=""../pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("SelectedPG"), "") & """>" & "Buyer " & externalBuyerid & " - " & MyCommon.SplitNonSpacedString(MyCommon.NZ(row3.Item("Name"), ""), 25) & "</a>")
                                                Else
                                                    Details.Append("<a href=""../pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("SelectedPG"), "") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row3.Item("Name"), ""), 25) & "</a>")
                                                End If
                                    
                                            End If
                                        End If
                                    Next
                                End If
                                If MyCommon.NZ(row.Item("DiscountTypeID"), 0) = 2 Then
                                    Details.Append(" (" & StrConv(Copient.PhraseLib.Lookup("term.department", LanguageID), VbStrConv.Lowercase) & ")")
                                ElseIf MyCommon.NZ(row.Item("DiscountTypeID"), 0) = 6 Then
                                    Details.Append(" (" & StrConv(Copient.PhraseLib.Lookup("term.grouplevel", LanguageID), VbStrConv.Lowercase) & ")")
                                ElseIf MyCommon.NZ(row.Item("DiscountTypeID"), 0) = 3 Then
                                    Details.Append(" (" & StrConv(Copient.PhraseLib.Lookup("term.basket", LanguageID), VbStrConv.Lowercase) & ")")
                                End If
                                'AMS-685 Show multiple exclusion groups for discount reward
                             
                                ' DisplayDiscountExclusionGroups(m_DiscountPGService, row.Item("DiscountID"), infoMessage, MyCommon, LanguageID)
                                'If MyCommon.NZ(row.Item("ExcludedPG"), 0) = 0 Then
                                'Else
                                '    MyCommon.QueryStr = "select Name from ProductGroups with (NoLock) where ProductGroupID = " & row.Item("ExcludedPG")
                                '    rst3 = MyCommon.LRT_Select()
                                '    For Each row3 In rst3.Rows
                                '        Details.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.excluding", LanguageID), VbStrConv.Lowercase) & " ")
                                '        If MyCommon.NZ(row.Item("ExcludedPG"), 0) = 1 Then
                                '            Details.Append(StrConv(MyCommon.NZ(row3.Item("Name"), ""), VbStrConv.Lowercase))
                                '        Else
                                '            If Popup Then
                                '                Details.Append(MyCommon.SplitNonSpacedString(MyCommon.NZ(row3.Item("Name"), ""), 25))
                                '            Else
                                '                Details.Append("<a href=""../pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("ExcludedPG"), "") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row3.Item("Name"), ""), 25) & "</a>")
                                '            End If
                                '        End If
                                '    Next
                                'End If
                
                                If MyCommon.NZ(row.Item("L1Cap"), 0) > 0 Then
                                    Details.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.upto", LanguageID), VbStrConv.Lowercase) & " " & Localizer.GetCached_Currency_Symbol(roid) & Math.Round(CDec(row.Item("L1Cap")), Localizer.GetCached_Currency_Precision(roid)).ToString(MyCommon.GetAdminUser.Culture))
                                End If
                
                                If MyCommon.NZ(row.Item("ItemLimit"), 0) = 0 And MyCommon.NZ(row.Item("WeightLimit"), 0) = 0 And MyCommon.NZ(row.Item("DollarLimit"), 0) = 0 Then
                                    Details.Append(",&nbsp;" & StrConv(Copient.PhraseLib.Lookup("term.unlimited", LanguageID), VbStrConv.Lowercase))
                                Else
                                    If MyCommon.NZ(row.Item("DiscountTypeID"), 0) <> 3 Then
                                        Details.Append(",&nbsp;" & StrConv(Copient.PhraseLib.Lookup("term.limit", LanguageID), VbStrConv.Lowercase) & "&nbsp;")
                                        If rstTiers.Rows.Count > 0 Then
                                            If MyCommon.NZ(rstTiers.Rows(0).Item("ItemLimit"), 0) > 0 Then
                                                t = 1
                                                For Each row4 In rstTiers.Rows
                                                    Details.Append(MyCommon.NZ(row4.Item("ItemLimit"), ""))
                                                    If t < rstTiers.Rows.Count Then
                                                        Details.Append("/")
                                                    End If
                                                    t += 1
                                                Next
                                                Details.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.items", LanguageID), VbStrConv.Lowercase))
                                            End If
                                            If MyCommon.NZ(rstTiers.Rows(0).Item("DollarLimit"), 0) > 0 Then
                                                If MyCommon.NZ(rstTiers.Rows(0).Item("ItemLimit"), 0) > 0 Then
                                                    Details.Append(" (")
                                                End If
                                                t = 1
                                                For Each row4 In rstTiers.Rows
                                                    Details.Append(Localizer.GetCached_Currency_Symbol(roid) & Math.Round(CDec(MyCommon.NZ(row4.Item("DollarLimit"), 0)), Localizer.GetCached_Currency_Precision(roid)).ToString(MyCommon.GetAdminUser.Culture))
                                                    If t < rstTiers.Rows.Count Then
                                                        Details.Append("/")
                                                    End If
                                                    t += 1
                                                Next
                                                If MyCommon.NZ(rstTiers.Rows(0).Item("ItemLimit"), 0) > 0 Then
                                                    Details.Append(")")
                                                End If
                                            End If
                                            If MyCommon.NZ(rstTiers.Rows(0).Item("WeightLimit"), 0) > 0 Then
                                                If MyCommon.NZ(rstTiers.Rows(0).Item("ItemLimit"), 0) > 0 OrElse MyCommon.NZ(rstTiers.Rows(0).Item("DollarLimit"), 0) > 0 Then
                                                    Details.Append(" (")
                                                End If
                                                t = 1
                                                For Each row4 In rstTiers.Rows
                                                    Details.Append(Math.Round(CDec(MyCommon.NZ(row4.Item("WeightLimit"), 0)), Localizer.GetCached_UOM_Precision(roid, AmountTypeID)).ToString(MyCommon.GetAdminUser.Culture))
                                                    If t < rstTiers.Rows.Count Then
                                                        Details.Append("/")
                                                    End If
                                                    t += 1
                                                Next
                                                Details.Append(" " & Localizer.GetCached_UOM_Abbreviation(roid, AmountTypeID))
                                                If MyCommon.NZ(rstTiers.Rows(0).Item("ItemLimit"), 0) > 0 OrElse MyCommon.NZ(rstTiers.Rows(0).Item("DollarLimit"), 0) > 0 Then
                                                    Details.Append(")")
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                                ' If there are multiple levels, this will display their details on a second line.
                                If (MyCommon.NZ(row.Item("L2DiscountAmt"), 0) > 0) And MyCommon.NZ(row.Item("L2AmountTypeID"), 0) = 3 Then
                                    Details.Append("<br />(" & Copient.PhraseLib.Lookup("term.over", LanguageID) & " " & Localizer.GetCached_Currency_Symbol(roid) & Math.Round(CDec(MyCommon.NZ(row.Item("L1Cap"), 0)), Localizer.GetCached_Currency_Precision(roid)).ToString(MyCommon.GetAdminUser.Culture) & ", ")
                                    Details.Append(row.Item("L2DiscountAmt") & "% " & StrConv(Copient.PhraseLib.Lookup("term.off", LanguageID), VbStrConv.Lowercase))
                                    If (MyCommon.NZ(row.Item("L2Cap"), 0) > 0) Then
                                        Details.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.upto", LanguageID), VbStrConv.Lowercase) & " " & Localizer.GetCached_Currency_Symbol(roid) & Math.Round(CDec(MyCommon.NZ(row.Item("L2Cap"), 0)), Localizer.GetCached_Currency_Precision(roid)).ToString(MyCommon.GetAdminUser.Culture) & ", ")
                                    End If
                                    Details.Append(")")
                                    If (MyCommon.NZ(row.Item("L3DiscountAmt"), 0) > 0) And MyCommon.NZ(row.Item("L3AmountTypeID"), 0) = 3 Then
                                        Details.Append("<br />(" & Copient.PhraseLib.Lookup("term.over", LanguageID) & " " & Localizer.GetCached_Currency_Symbol(roid) & Math.Round(CDec(MyCommon.NZ(row.Item("L2Cap"), 0)), Localizer.GetCached_Currency_Precision(roid)).ToString(MyCommon.GetAdminUser.Culture) & ", ")
                                        Details.Append(row.Item("L3DiscountAmt") & "% " & StrConv(Copient.PhraseLib.Lookup("term.off", LanguageID), VbStrConv.Lowercase) & ")")
                                    End If
                                End If
                                Send("<li>" & Details.ToString & "</li>")
                            Next
                            If Not ((MyCommon.NZ(row.Item("DiscountTypeID"), 0) = 4 OrElse MyCommon.NZ(row.Item("DiscountTypeID"), 0) = 5) AndAlso MyCommon.NZ(row.Item("SelectedPG"), 0) = 0) Then
                                DisplayDiscountExclusionGroups(m_DiscountPGService, row.Item("DiscountID"), infoMessage, MyCommon, LanguageID)
                            End If
                            Send("</ul>")
                            Send("<br class=""half"" />")
                        End If
                        ' Printed message rewards
                        MyCommon.QueryStr = "select PM.MessageID, PM.MessageTypeID, D.DeliverableID " & _
                                            "from CPE_Deliverables D with (NoLock) inner join PrintedMessages PM with (NoLock) on D.OutputID=PM.MessageID " & _
                                            "where D.Deleted=0 and D.RewardOptionPhase=3 and D.RewardOptionID=" & roid & " and D.DeliverableTypeID=4;"
                        rst = MyCommon.LRT_Select()
                        If (rst.Rows.Count > 0) Then
                            counter = counter + 1
                            Dim Details As StringBuilder
                            Send("<h3>" & Copient.PhraseLib.Lookup("term.printedmessages", LanguageID) & "</h3>")
                            Send("<ul class=""condensed"">")
                            For Each row In rst.Rows
                                MyCommon.QueryStr = "select PMT.BodyText from PrintedMessageTiers PMT with (NoLock) " & _
                                                    "where MessageID=" & MyCommon.NZ(row.Item("MessageID"), 0) & " order by TierLevel;"
                                rstTiers = MyCommon.LRT_Select()
                                If rstTiers.Rows.Count > 0 Then
                                    For Each row4 In rstTiers.Rows
                                        Details = New StringBuilder(200)
                                        Details.Append(ReplaceTags(MyCommon.NZ(row4.Item("BodyText"), "")))
                                        If (Details.ToString().Length > 80) Then
                                            Details = Details.Remove(77, (Details.Length - 77))
                                            Details.Append("...")
                                        End If
                                        'Details.Replace(vbCrLf, vbCrLf & "<br/>")
                                        'Overriding String Split
                                        Send("<li>""" & Server.HtmlEncode(MyCommon.SplitNonSpacedString(Details.ToString, Details.ToString.Length)) & """</li>")
                                    Next
                                End If
                            Next
                            Send("</ul>")
                            Send("<br class=""half"" />")
                        End If
            
                        ' Cashier message rewards
                        MyCommon.QueryStr = "select D.DeliverableID, CM.MessageID from CPE_Deliverables D with (NoLock) " & _
                                            "inner join CPE_CashierMessages CM with (NoLock) on D.OutputID=CM.MessageID " & _
                                            "where D.RewardOptionID=" & roid & " and D.Deleted=0 and DeliverableTypeID=9 and D.RewardOptionPhase=3;"
                        rst = MyCommon.LRT_Select
                        If (rst.Rows.Count > 0) Then
                            Dim numrows As Integer = MyCommon.Fetch_UE_SystemOption(158)
                            counter = counter + 1
                            Send("<h3>" & Copient.PhraseLib.Lookup("term.cashiermessages", LanguageID) & "</h3>")
                            Send("<ul class=""condensed"">")
                            For Each row In rst.Rows
                                MyCommon.QueryStr = "select Line1, Line2, Line3, Line4, Line5, Line6, Line7, Line8, Line9, Line10, Beep, BeepDuration from CPE_CashierMessageTiers as CMT with (NoLock) " & _
                                                    "where MessageID=" & MyCommon.NZ(row.Item("MessageID"), 0) & " order by TierLevel;"
                                rstTiers = MyCommon.LRT_Select()
                                If rstTiers.Rows.Count > 0 Then
                                    Dim rc As Integer = 1
                                    For Each row4 In rstTiers.Rows
                                        'Send("<i>Tier" & rc & "</i><br />")
                                        'Send("<li>""" & MyCommon.NZ(row4.Item("Line1"), "") & "<br />" & MyCommon.NZ(row4.Item("Line2"), "") & """</li>")
                                        'Send("""")
                                        Dim lines As Integer = 0
                                        For lines = 1 To numrows
                                            If MyCommon.NZ(row4.Item("Line" & lines), "") <> "" Then
                                                Send("<li>" & MyCommon.NZ(row4.Item("Line" & lines), "") & "<br />")
                                            End If
                                        Next
                                        rc += 1
                                    Next
                                End If
                            Next
                            Send("</ul>")
                            Send("<br class=""half"" />")
                        End If
            
                        ' Franking rewards
                        MyCommon.QueryStr = "select D.DeliverableID, FM.FrankID from CPE_Deliverables D with (NoLock) " & _
                                            "inner join CPE_FrankingMessages FM with (NoLock) on D.OutputID=FM.FrankID " & _
                                            "where D.RewardOptionID=" & roid & " and D.Deleted=0 and DeliverableTypeID=10 and D.RewardOptionPhase=3;"
                        rst = MyCommon.LRT_Select
                        If (rst.Rows.Count > 0) Then
                            counter = counter + 1
                            Send("<h3>" & Copient.PhraseLib.Lookup("term.frankingmessages", LanguageID) & "</h3>")
                            Send("<ul class=""condensed"">")
                            For Each row In rst.Rows
                                MyCommon.QueryStr = "select FrankID, OpenDrawer, ManagerOverride, FrankFlag, FrankingText, Line1, Line2, Beep, BeepDuration " & _
                                                    "from CPE_FrankingMessageTiers as FMT with (NoLock) " & _
                                                    "where FrankID=" & MyCommon.NZ(row.Item("FrankID"), 0) & " order by TierLevel;"
                                rstTiers = MyCommon.LRT_Select()
                                If rstTiers.Rows.Count > 0 Then
                                    For Each row4 In rstTiers.Rows
                                        If MyCommon.NZ(row4.Item("FrankingText"), "") = "" Then
                                            Send("<li>")
                                        Else
                                            Send("<li>""" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row4.Item("FrankingText"), ""), 25) & """<br />")
                                        End If
                                        Sendb(IIf(MyCommon.NZ(row4.Item("OpenDrawer"), False) = True, Copient.PhraseLib.Lookup("term.opendrawer", LanguageID) & ",&nbsp;", Copient.PhraseLib.Lookup("term.closeddrawer", LanguageID) & ",&nbsp;"))
                                        Sendb(IIf(MyCommon.NZ(row4.Item("ManagerOverride"), False) = True, StrConv(Copient.PhraseLib.Lookup("term.override", LanguageID), VbStrConv.Lowercase) & " " & StrConv(Copient.PhraseLib.Lookup("term.required", LanguageID), VbStrConv.Lowercase) & ",&nbsp;", StrConv(Copient.PhraseLib.Lookup("term.override", LanguageID), VbStrConv.Lowercase) & " " & StrConv(Copient.PhraseLib.Lookup("term.notrequired", LanguageID), VbStrConv.Lowercase) & ",&nbsp;"))
                                        If (MyCommon.NZ(row4.Item("FrankFlag"), 0) = 0) Then
                                            Sendb(Copient.PhraseLib.Lookup("term.posdata", LanguageID) & " ")
                                            Sendb(StrConv(Copient.PhraseLib.Lookup("term.only", LanguageID), VbStrConv.Lowercase))
                                        ElseIf (MyCommon.NZ(row4.Item("FrankFlag"), 0) = 1) Then
                                            Sendb(StrConv(Copient.PhraseLib.Lookup("term.frankingmessage", LanguageID), VbStrConv.Lowercase) & " ")
                                            Sendb(StrConv(Copient.PhraseLib.Lookup("term.only", LanguageID), VbStrConv.Lowercase))
                                        ElseIf (MyCommon.NZ(row4.Item("FrankFlag"), 0) = 2) Then
                                            Sendb(Copient.PhraseLib.Lookup("term.posdata", LanguageID) & " ")
                                            Sendb(StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
                                            Sendb(StrConv(Copient.PhraseLib.Lookup("term.frankingmessage", LanguageID), VbStrConv.Lowercase) & " ")
                                        End If
                                        Send("</li>")
                                    Next
                                End If
                            Next
                            Send("</ul>")
                            Send("<br class=""half"" />")
                        End If
            
                        ' Points rewards
                        MyCommon.QueryStr = "select PP.ProgramName as Name, PP.ProgramID, D.DeliverableID as PKID, DP.PKID as DPPKID " & _
                                            "from PointsPrograms as PP with (NoLock) " & _
                                            "inner join CPE_DeliverablePoints as DP with (NoLock) on PP.ProgramID=DP.ProgramID and DP.Deleted=0 and PP.Deleted=0 " & _
                                            "inner join CPE_Deliverables as D with (NoLock) on D.DeliverableID=DP.DeliverableID and D.DeliverableTypeID=8 and D.Deleted=0 and D.RewardOptionPhase=3 " & _
                                            "where D.RewardOptionID=" & roid & " order by PP.ProgramName;"
                        rst = MyCommon.LRT_Select()
                        If (rst.Rows.Count > 0) Then
                            counter = counter + 1
                            Send("<h3>" & Copient.PhraseLib.Lookup("term.points", LanguageID) & "</h3>")
                            Send("<ul class=""condensed"">")
                            For Each row In rst.Rows
								MyCommon.QueryStr = "select Quantity, Multiplier from CPE_DeliverablePointTiers as DPT with (NoLock) " & _
                                                    "where DPPKID=" & MyCommon.NZ(row.Item("DPPKID"), 0) & " order by TierLevel;"
                                rstTiers = MyCommon.LRT_Select()
                                If rstTiers.Rows.Count > 0 Then
                                    Sendb("<li>")
                                    t = 1
                                    For Each row4 In rstTiers.Rows
										Sendb( (MyCommon.NZ(row4.Item("Quantity"), 0) * (MyCommon.NZ(row4.Item("Multiplier"), 1)) & " ")) 'vtopol
                                        If t < TierLevels Then
                                            Sendb(" / ")
                                        End If
                                        t = t + 1
                                    Next
                                End If
                                If Popup Then
                                    Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25))
                                Else
                                    Sendb("<a href=""../point-edit.aspx?ProgramGroupID=" & MyCommon.NZ(row.Item("ProgramID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
                                End If
                            Next
                            Send("</ul>")
                            Send("<br class=""half"" />")
                        End If
                        'Trackable Coupon rewards
                        Dim objCouponReward As List(Of CouponReward) = objCouponService.GetAllCouponRewardbyROID(roid)
                        If (objCouponReward.Count > 0) Then
                            counter = counter + 1
                            Send("<h3>" & Copient.PhraseLib.Lookup("term.trackablecoupon", LanguageID) & "</h3>")
                            Send("<ul class=""condensed"">")
               
                            For Each couponReward In objCouponReward
                                If couponReward.CouponTiers.Count > 0 Then
                                    Sendb("<li style='word-wrap:break-word'>")
                                    t = 1
                                    While t <= TierLevels
                                        If couponReward.CouponTiers.Count <= (t - 1) Then
                                            Send(Copient.PhraseLib.Lookup("term.undefined", LanguageID))
                                        Else
                                            MyCommon.QueryStr = "select Name  " & _
                                                 "from TrackableCouponProgram with (NoLock) " & _
                                                 "where ProgramID=" & couponReward.CouponTiers(t - 1).ProgramID & ";"
                                            rst2 = MyCommon.LRT_Select
                                            If rst2.Rows.Count > 0 Then
                                                If Popup Then
                                                    Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(rst2.Rows(0).Item("Name"), ""), 25))
                                                Else
                                                    Sendb("<a href=""../tcp-edit.aspx?tcprogramid=" & couponReward.CouponTiers(t - 1).ProgramID & """>" & (MyCommon.NZ(rst2.Rows(0).Item("Name"), "")) & "</a>")
                                                End If
                                            Else
                                                Send(Copient.PhraseLib.Lookup("term.undefined", LanguageID))
                                            End If
                                        End If
                                        If t < TierLevels Then
                                            Sendb(" / ")
                                        End If
                                        t = t + 1
                                    End While
                        
                                End If
                   
                            Next
                            Send("</ul>")
                            Send("<br class=""half"" />")
                        End If
                        ' Stored value rewards
                        MyCommon.QueryStr = "select distinct DSV.PKID, DSV.DeliverableID, SVP.Name, SVP.Value, SVP.SVTypeID, SVT.ValuePrecision, SVP.SVProgramID from CPE_Deliverables as D with (NoLock) " & _
                                                 "Inner Join CPE_DeliverableStoredValue as DSV with (NoLock) on DSV.PKID=D.OutputID and IsNull(DSV.Deleted,0)=0 " & _
                                                 "Inner Join StoredValuePrograms as SVP with (NoLock) on SVP.SVProgramID=DSV.SVProgramID and SVP.Deleted=0 " & _
                                                 "Inner Join SVTypes as SVT with (NoLock) on SVP.SVTypeID=SVT.SVTypeID " & _
                                                 "where D.RewardOptionPhase = 3 And D.DeliverableTypeID = 11 And D.RewardOptionID =" & roid
                        rst = MyCommon.LRT_Select()
                        If (rst.Rows.Count > 0) Then
                            counter = counter + 1
                            Send("<h3>" & Copient.PhraseLib.Lookup("term.storedvalue", LanguageID) & "</h3>")
                            Send("<ul class=""condensed"">")
                            For Each row In rst.Rows
                                Sendb("<li>")
								 MyCommon.QueryStr = "select Quantity, Multiplier from CPE_DeliverableStoredValueTiers with (NoLock) where DSVPKID=" & MyCommon.NZ(row.Item("PKID"), 0) & ";"
                                rstTiers = MyCommon.LRT_Select
                                If rstTiers.Rows.Count > 0 Then
                                    t = 1
                                    For Each row4 In rstTiers.Rows
										Sendb(MyCommon.NZ(row4.Item("Quantity"), "0") * MyCommon.NZ(row4.Item("Multiplier"), "1") & " ")
                                        If MyCommon.NZ(row.Item("SVTypeID"), 0) > 1 Then
											Sendb(" ($" & Math.Round(CDec(MyCommon.NZ(row.Item("Value"), 0) * (MyCommon.NZ(row4.Item("Quantity"), 0) * MyCommon.NZ(row4.Item("Multiplier"), 1))), CInt(MyCommon.NZ(row.Item("ValuePrecision"), 0))).ToString(MyCommon.GetAdminUser.Culture) & ")")
                                        End If
                                        If t < TierLevels Then
                                            Sendb(" / ")
                                        End If
                                        t = t + 1
                                    Next
                                End If
                                If Popup Then
                                    Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25))
                                Else
                                    Sendb("<a href=""../SV-edit.aspx?ProgramGroupID=" & MyCommon.NZ(row.Item("SVProgramID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
                                End If
                            Next
                            Send("</ul>")
                            Send("<br class=""half"" />")
                        End If
          
            
                        'Group membership rewards
                        MyCommon.QueryStr = "select D.DeliverableID, D.DeliverableTypeID, D.RewardOptionID as ROID " & _
                                            "from CPE_Deliverables as D with (NoLock) " & _
                                            "where D.Deleted=0 and D.RewardOptionID=" & roid & " and D.DeliverableTypeID IN (5) and D.RewardOptionPhase=3;"
                        rst = MyCommon.LRT_Select()
                        If (rst.Rows.Count > 0) Then
                            counter = counter + 1
                            Send("<h3>" & Copient.PhraseLib.Lookup("term.groupmembership", LanguageID) & "</h3>")
                            Send("<ul class=""condensed"">")
                            For Each row In rst.Rows
                                MyCommon.QueryStr = "select DCGT.CustomerGroupID, CG.Name from CPE_DeliverableCustomerGroupTiers as DCGT with (NoLock) " & _
                                                    "inner join CustomerGroups as CG with (NoLock) on CG.CustomerGroupID=DCGT.CustomerGroupID " & _
                                                    "where DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & ";"
                                rstTiers = MyCommon.LRT_Select
                                If rstTiers.Rows.Count > 0 Then
                                    For Each row4 In rstTiers.Rows
                                        If (MyCommon.NZ(row4.Item("CustomerGroupID"), 0) > 2) Then
                                            If Popup Then
                                                Sendb("<li>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row4.Item("Name"), ""), 25) & "&nbsp;")
                                            Else
                                                Sendb("<li><a href=""../cgroup-edit.aspx?CustomerGroupID=" & MyCommon.NZ(row4.Item("CustomerGroupID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row4.Item("Name"), ""), 25) & "</a>&nbsp;")
                                            End If
                                        Else
                                            Sendb("<li>" & MyCommon.SplitNonSpacedString(row.Item("Name"), 25) & "&nbsp;")
                                        End If
                                        MyCommon.QueryStr = "select count(*) as GCount from GroupMembership with (NoLock) where CustomerGroupID=" & MyCommon.NZ(row4.Item("CustomerGroupID"), -1) & " And Deleted=0;"
                                        rst2 = MyCommon.LXS_Select()
                                        For Each row2 In rst2.Rows
                                            If (MyCommon.NZ(row4.Item("CustomerGroupID"), -1) > 2) Then
                                                Send("(" & row2.Item("GCount") & ")")
                                            End If
                                        Next
                                        Send("</li>")
                                    Next
                                End If
                            Next
                            Send("</ul>")
                            Send("<br class=""half"" />")
                        End If
            
                        ' Graphics rewards
                        MyCommon.QueryStr = "select OSA.Name as GraphicName, OSA.OnScreenAdID as AdID, D.DeliverableID, D.ScreenCellID as CellID, OSA.Width, OSA.Height, OSA.ImageType, SC.Name as CellName, SL.Name as LayoutName " & _
                                            "from OnScreenAds as OSA with (NoLock) Inner Join CPE_Deliverables as D with (NoLock) on OSA.OnScreenAdID=D.OutputID and D.RewardOptionID=" & roid & " and OSA.Deleted=0 and D.Deleted=0 and D.DeliverableTypeID=1 and D.RewardOptionPhase=3 " & _
                                            "Inner Join ScreenCells as SC with (NoLock) on D.ScreenCellID=SC.CellID " & _
                                            "Inner Join ScreenLayouts as SL with (NoLock) on SL.LayoutID=SC.LayoutID;"
                        rst = MyCommon.LRT_Select
                        If (rst.Rows.Count > 0) Then
                            counter = counter + 1
                            Send("<h3>" & Copient.PhraseLib.Lookup("term.graphics", LanguageID) & "</h3>")
                            Send("<ul class=""condensed"">")
                            For Each row In rst.Rows
                                If Popup Then
                                    Sendb("<li>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("GraphicName"), ""), 25) & "&nbsp;")
                                Else
                                    Sendb("<li><a href=""../graphic-edit.aspx?OnScreenAdId=" & MyCommon.NZ(row.Item("AdId"), "") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("GraphicName"), ""), 25) & "</a>&nbsp;")
                                End If
                                Sendb("(" & MyCommon.NZ(row.Item("Width"), "") & " x " & MyCommon.NZ(row.Item("Height"), ""))
                                If MyCommon.NZ(row.Item("ImageType"), "") = 1 Then
                                    Sendb("&nbsp;" & Copient.PhraseLib.Lookup("term.jpeg", LanguageID))
                                ElseIf MyCommon.NZ(row.Item("ImageType"), "") = 2 Then
                                    Sendb("&nbsp;" & Copient.PhraseLib.Lookup("term.gif", LanguageID))
                                End If
                                Send(")</li>")
                            Next
                            Send("</ul>")
                            Send("<br class=""half"" />")
                        End If
            
                        ' Touchpoint rewards
                        MyCommon.QueryStr = "select RO.Name, RO.RewardOptionID, TA.OnScreenAdID as ParentAdID " & _
                                            "from CPE_RewardOptions RO with (NoLock) inner join CPE_DeliverableROIDs DR with (NoLock) on RO.RewardOptionID = DR.RewardOptionID " & _
                                            "inner join CPE_Deliverables D with (NoLock) on D.DeliverableID = DR.DeliverableID inner join TouchAreas TA with (NoLock) on DR.AreaID = TA.AreaID " & _
                                            "where RO.Deleted=0 and DR.Deleted=0 and TA.Deleted=0 and D.Deleted = 0 and RO.IncentiveID=" & OfferID & " and " & _
                                            "RO.TouchResponse=1 and D.RewardOptionPhase=3 order by RO.rewardoptionid;"
                        rst = MyCommon.LRT_Select
                        If (rst.Rows.Count > 0) Then
                            counter = counter + 1
                            Dim tpROID As Integer = 0
                            Send("<h3>" & Copient.PhraseLib.Lookup("term.touchpoints", LanguageID) & "</h3>")
                            Send("<ul class=""condensed"">")
                            For Each row In rst.Rows
                                tpROID = MyCommon.NZ(row.Item("RewardOptionID"), 0)
                                Send("<li>")
                                If Popup Then
                                    Send(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)), 25) & "<br />")
                                Else
                                    Send("<a href=""../graphic-edit.aspx?OnScreenAdId=" & MyCommon.NZ(row.Item("ParentAdID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)), 25) & "</a><br />")
                                End If
                                SetSummaryOnly(True)
                                Send_TouchpointRewards(OfferID, tpROID, 3, TierLevels)
                                Send("</li>")
                            Next
                            Send("</ul>")
                        End If
            
                        ' Pass-thru rewards
                        MyCommon.QueryStr = "select PassThruRewardID, Name, PhraseID from PassThruRewards with (NoLock) order by Name;"
                        rst2 = MyCommon.LRT_Select
                        If rst2.Rows.Count > 0 Then
                            For Each row2 In rst2.Rows
                
                                MyCommon.QueryStr = "select D.DeliverableID, D.DisallowEdit, DPT.PKID, DPT.PassThruRewardID, PTR.Name, PTR.PhraseID, PTR.LSInterfaceID, PTR.ActionTypeID " & _
                                                    "from CPE_Deliverables as D with (NoLock) " & _
                                                    "inner join PassThrus as DPT with (NoLock) on DPT.PKID=D.OutputID " & _
                                                    "inner join PassThruRewards as PTR with (NoLock) on PTR.PassThruRewardID=DPT.PassThruRewardID " & _
                                                    "where D.RewardOptionID=" & roid & " and DPT.PassThruRewardID=" & MyCommon.NZ(row2.Item("PassThruRewardID"), 0) & " and D.Deleted=0 and D.DeliverableTypeID=12 and RewardOptionPhase=3 " & _
                                                    "order by Name;"
                                rst = MyCommon.LRT_Select()
                                If (rst.Rows.Count > 0) Then
                                    counter = counter + 1
                                    If IsDBNull(row2.Item("PhraseID")) Then
                                        Send("<h3>" & MyCommon.NZ(row2.Item("Name"), Copient.PhraseLib.Lookup("term.passthrureward", LanguageID)) & "</h3>")
                                    Else
                                        Send("<h3>" & Copient.PhraseLib.Lookup(row2.Item("PhraseID"), LanguageID) & "</h3>")
                                    End If
                                    Send("<ul class=""condensed"">")
                                    For Each row In rst.Rows
                                        If IsDBNull(row.Item("PhraseID")) Then
                                            Send("  <li>" & MyCommon.NZ(row2.Item("Name"), "") & "</li>")
                                        Else
                                            Send("  <li>" & Copient.PhraseLib.Lookup(row2.Item("PhraseID"), LanguageID) & "</li>")
                                        End If
                                    Next
                                    Send("</ul>")
                                    Send("<br class=""half"" />")
                                End If
                
                            Next
                        End If

                        ' GiftCard rewards
                        MyCommon.QueryStr = "SELECT G.ID, G.LASTUPDATE, D.DELIVERABLEID, D.DISALLOWEDIT  FROM CPE_DELIVERABLES D INNER JOIN GIFTCARD G ON D.OutputID=G.ID WHERE G.REWARDOPTIONID=" & roid & " and D.DeliverableTypeID=13;"
                        rst2 = MyCommon.LRT_Select
                        If rst2.Rows.Count > 0 Then
                            counter = counter + 1
                            Send("<h3>" & Copient.PhraseLib.Lookup("term.Giftcard", LanguageID) & "</h3>")
                            Send("<ul class=""condensed"">")
                            Dim Details As StringBuilder = New StringBuilder(200)
                            Dim Details2 As StringBuilder = New StringBuilder(200)
                        
                            For Each row2 In rst2.Rows
                                MyCommon.QueryStr = "select gt.Id,gt.CardIdentifier,gt.Name, gt.TierLevel,gt.AmountTypeID ,gt.Amount,gt.BuyDescription,gt.ChargebackDeptID,pt.PhraseID as ProrationTypePhrase, at.PhraseID as AmountTypePhrase from GiftCardTier gt with (nolock) " & _
                                                    "inner join Relation_RewardProration RRP on RRP.ProrationTypeID=gt.ProrationTypeID " & _
                                                    "inner join UE_ProrationTypes PT with (nolock) on PT.ProrationTypeID=rrp.ProrationTypeID " & _
                                                    "inner join CPE_AmountTypes AT with (nolock) on at.AmountTypeID=gt.AmountTypeID " & _
                                                    "where gt.GiftCardID=" & MyCommon.NZ(row2.Item("id"), 0) & ";"
                                rstTiers = MyCommon.LRT_Select()
                                t = 1
                                For Each row In rstTiers.Rows
                                    If (t > 1) Then
                                        Details.Append(" / ")
                                        Details2.Append(" / ")
                                    End If
                                    AmountTypeID = (MyCommon.NZ(row.Item("AmountTypeID"), 0))
                                    Select Case AmountTypeID
                                        Case 1
                                            Details.Append(Localizer.GetCached_Currency_Symbol(roid) & Math.Round(CDec(MyCommon.NZ(rstTiers.Rows(t - 1).Item("Amount"), 0)), Localizer.GetCached_Currency_Precision(roid)).ToString(MyCommon.GetAdminUser.Culture))
                                        Case 3
                                            Details.Append(Math.Round(CDec(MyCommon.NZ(rstTiers.Rows(t - 1).Item("Amount"), 0)), 2).ToString(MyCommon.GetAdminUser.Culture) & "%")
                                        Case Else
                                            Details.Append(MyCommon.NZ(rstTiers.Rows(t - 1).Item("Amount"), ""))
                                    End Select
                                    Details2.Append(Copient.PhraseLib.Lookup(MyCommon.NZ(rstTiers.Rows(t - 1).Item("ProrationTypePhrase"), 0), LanguageID))
                                    t += 1
                                Next
                                Send("<li>" & Details.ToString & " " & StrConv(Copient.PhraseLib.Lookup("term.on", LanguageID), VbStrConv.Lowercase) & " " & StrConv(Copient.PhraseLib.Lookup("term.proration", LanguageID), VbStrConv.Lowercase) & " (" & Details2.ToString & ") " & "</li>")
                            Next
                            Send("</ul>")
                            Send("<br class=""half"" />")
                        End If
         
            
                        ' Proximity Message reward(s)
                        t = 1
                        MyCommon.QueryStr = "SELECT P.ID, D.LASTUPDATE, D.DELIVERABLEID, D.DISALLOWEDIT FROM CPE_DELIVERABLES D INNER JOIN PROXIMITYMESSAGE P ON D.OutputID=P.ID WHERE D.REWARDOPTIONID=" & roid & " and D.DeliverableTypeID=14;"
                        rst = MyCommon.LRT_Select()
                        If (rst.Rows.Count > 0) Then
                            counter = counter + 1
                            Send("<h3>" & Copient.PhraseLib.Lookup("term.proximitymessage", LanguageID) & "</h3>")
                            Send("<ul class=""condensed"">")
                            Dim Details As StringBuilder = New StringBuilder(200)
              
                            For Each row In rst.Rows
                                ' Find the per-tier details, and build up the details string:
                                MyCommon.QueryStr = "select PM.ThresholdTypeID, PMT.TierLevel ,PMT.TriggerValue from ProximityMessage as PM " & _
                                                    "left join ProximityMessageTier as PMT " & _
                                                    "on PM.ID = PMT.ProximityMessageId " & _
                                                    "left join CPE_Deliverables as CPED " & _
                                                    "on CPED.OutputID = PM.ID " & _
                                                    "where CPED.DeliverableTypeID = 14 and Deleted = 0 and CPED.DeliverableID = " & row("DELIVERABLEID").ToString()
                                rst2 = MyCommon.LRT_Select
                                If rst2.Rows.Count > 0 Then
                                    counter = counter + 1
                                    Dim valueType As Integer = Integer.Parse(rst2.Rows(0)(0))
                                    Dim valueSymbol As String = ""
                                    Dim valueAbbr As String = ""
                                    Dim valueLabel As String = ""
                                    Dim valuePrecision As String = ""
                                    Dim tempPrecision As Integer = 0
                                    Select Case valueType
                                        Case 1
                                            valueSymbol = ""
                                            valueAbbr = ""
                                            valueLabel = Copient.PhraseLib.Lookup("term.quantityaway", LanguageID) & ": "
                                            tempPrecision = 0
                                        Case 2
                                            valueSymbol = Localizer.Get_Currency_Symbol(roid)
                                            valueAbbr = ""
                                            valueLabel = Copient.PhraseLib.Lookup("term.amountaway", LanguageID) & ": "
                                            tempPrecision = Localizer.Get_Currency_Precision(roid)
                                        Case 3
                                            valueSymbol = ""
                                            valueAbbr = " lbs/gals"
                                            valueLabel = Copient.PhraseLib.Lookup("term.quantityaway", LanguageID) & ": "
                                            tempPrecision = 3
                                        Case CPEUnitTypes.Weight
                                            Dim quantityInfo As QuantityUnitTypeInfoRec = Localizer.Load_QuantityInfo_ForOffer(roid, CPEUnitTypes.Weight)
                                            valueSymbol = ""
                                            valueAbbr = quantityInfo.Abbrevation
                                            valueLabel = Copient.PhraseLib.Lookup("term.quantityaway", LanguageID) & ": "
                                            tempPrecision = quantityInfo.Precision
                                        Case CPEUnitTypes.Volume
                                            Dim quantityInfo As QuantityUnitTypeInfoRec = Localizer.Load_QuantityInfo_ForOffer(roid, CPEUnitTypes.Volume)
                                            valueSymbol = ""
                                            valueAbbr = quantityInfo.Abbrevation
                                            valueLabel = Copient.PhraseLib.Lookup("term.quantityaway", LanguageID) & ": "
                                            tempPrecision = quantityInfo.Precision
        												
                                        Case CPEUnitTypes.Length
                                            Dim quantityInfo As QuantityUnitTypeInfoRec = Localizer.Load_QuantityInfo_ForOffer(roid, CPEUnitTypes.Length)
                                            valueSymbol = ""
                                            valueAbbr = quantityInfo.Abbrevation
                                            valueLabel = Copient.PhraseLib.Lookup("term.quantityaway", LanguageID) & ": "
                                            tempPrecision = quantityInfo.Precision
                                        Case CPEUnitTypes.SurfaceArea
                                            Dim quantityInfo As QuantityUnitTypeInfoRec = Localizer.Load_QuantityInfo_ForOffer(roid, CPEUnitTypes.SurfaceArea)
                                            valueSymbol = ""
                                            valueAbbr = quantityInfo.Abbrevation
                                            valueLabel = Copient.PhraseLib.Lookup("term.quantityaway", LanguageID) & ": "
                                            tempPrecision = quantityInfo.Precision
                                        Case 9
                                            valueSymbol = ""
                                            valueAbbr = " points"
                                            valueLabel = Copient.PhraseLib.Lookup("term.quantityaway", LanguageID) & ": "
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
                      
                                    Details.Append("<li>" & valueLabel & "<br/>")
                        
                                    While t <= TierLevels
                                        If t > 1 Then
                                            Details.Append(" / ")
                                        End If
                                        If t > rst2.Rows.Count Then
                                            Send(Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</li>")
                                        Else
                                            Details.Append(valueSymbol & Math.Round(CDec(rst2.Rows(t - 1)(2)), tempPrecision).ToString(MyCommon.GetAdminUser.Culture))
                                        End If
                                        t += 1
                                    End While
                                    Details.Append(" " & valueAbbr)
                                    Send(Details.ToString)
                                End If
                                Details.Clear()
                                t = 1
                            Next
                            Send("</ul>")
                            Send("<br class=""half"" />")
                        End If
            
                        'Preference rewards
                        Dim IntegrationVals As New Copient.CommonInc.IntegrationValues
                        If MyCommon.IsEngineInstalled(9) AndAlso MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER, IntegrationVals) Then
                            Dim preferencerewards As AMSResult(Of List(Of PreferenceReward)) = m_PreferenceRewardService.GetAllPreferenceRewardByROID(roid)
                            If preferencerewards.ResultType = AMSResultType.Success AndAlso preferencerewards.Result.Count > 0 Then
                                counter = counter + 1
                                Send("<h3>" & Copient.PhraseLib.Lookup("term.preference", LanguageID) & " " & Copient.PhraseLib.Lookup("term.rewards", LanguageID).ToLower() & "</h3>")
                                Send("<ul class=""condensed"">")
                                For Each preferencereward As PreferenceReward In preferencerewards.Result
                                    t = 1
                                    Dim preference As Preference = m_PreferenceService.GetPreferenceByID(preferencereward.PreferenceID, LanguageID).Result
                                    If preference Is Nothing Then
                                        Continue For
                                    End If
                                    preference.PreferenceValues = m_PreferenceService.GetPreferenceItemsbyPreferenceID(preference.DataTypeID, preference.PreferenceID, LanguageID).Result
                                    Dim PrefPageName As String = IIf(preference.UserCreated, "prefscustom-edit.aspx", "prefsstd-edit.aspx")
                                    Dim RootURI As String = IntegrationVals.HTTP_RootURI
                                    If RootURI IsNot Nothing AndAlso RootURI.Length > 0 AndAlso Right(RootURI, 1) <> "/" Then
                                        RootURI &= "/"
                                    End If
                                    Send("  <li>")
                                    Sendb("  <a href=""../authtransfer.aspx?SendToURI=" & RootURI & "UI/" & PrefPageName & "?prefid=" & preference.PreferenceID & """>")
                                    Send(preference.PhraseText & "</a>")
                                    Send("<br />")
                                    While t <= TierLevels
                                        If TierLevels > 1 Then
                                            Send(Copient.PhraseLib.Lookup("term.tier", LanguageID) & t & ": ")
                                        End If
                                        If t > preferencereward.PreferenceRewardTiers.Count Then
                                            Send(Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "<br />")
                                        Else
                                            Dim tier As PreferenceRewardTier = (From p In preferencereward.PreferenceRewardTiers Where p.TierLevel = t).ToList.First
                                            If tier.PreferenceRewardTierValues.Count > 1 Then
                                                Send(Copient.PhraseLib.Lookup("term.multiple", LanguageID) & "<br />")
                                            ElseIf tier.PreferenceRewardTierValues.Count = 1 Then
                                                Dim tiervalue As String = preference.PreferenceValues.Where(Function(p) tier.PreferenceRewardTierValues.Select(Function(p2) p2.PreferenceValue).Contains(p.Value)).Select(Function(p3) p3.PhraseText).First
                                                Send(tiervalue & "<br />")
                                            End If
                                        End If
                                        t += 1
                                    End While
                                    Send("  </li>")
                                Next
                                Send("</ul>")
                                Send("<br class=""half"" />")
                            End If
                        End If
          
                        ' Monetary Stored value rewards
                        MyCommon.QueryStr = "select distinct DSV.PKID, DSV.DeliverableID, SVP.Name, SVP.Value, SVP.SVTypeID, SVP.SVProgramID from CPE_Deliverables as D with (NoLock) " & _
                                                 "Inner Join CPE_DeliverableMonStoredValue as DSV with (NoLock) on DSV.PKID=D.OutputID and DSV.RewardOptionID = D.RewardOptionID and IsNull(DSV.Deleted,0)=0 " & _
                                                 "Inner Join StoredValuePrograms as SVP with (NoLock) on SVP.SVProgramID=DSV.SVProgramID and SVP.Deleted=0 " & _
                                                 "Inner Join SVTypes as SVT with (NoLock) on SVP.SVTypeID=SVT.SVTypeID " & _
                                                 "where D.RewardOptionPhase = 3 And D.DeliverableTypeID = 16 And D.RewardOptionID =" & roid
                        rst = MyCommon.LRT_Select()
                        If (rst.Rows.Count > 0) Then
                            counter = counter + 1
                            Send("<h3>" & Copient.PhraseLib.Lookup("term.monetarystoredvalue", LanguageID) & "</h3>")
                            Send("<ul class=""condensed"">")
                            For Each row In rst.Rows
                                Sendb("<li>")
                                If Popup Then
                                    Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25))
                                Else
                                    Sendb("<a href=""../SV-edit.aspx?ProgramGroupID=" & MyCommon.NZ(row.Item("SVProgramID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
                                End If
                            Next
                            Send("</li></ul>")
                            Send("<br class=""half"" />")
                        End If
                        If (counter = 0) Then
                            Send("<h3>" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "</h3>")
                        End If
                        counter = 0
                    %>
                    <hr class="hidden" />
                </div>
            </div>
        </div>
        <br clear="all" />
    </div>
</form>
<div id="fadeDiv">
</div>
<div id="DuplicateNoofOffer" class="folderdialog" style="position: absolute; top: 200px; left: 400px; width: 400px; height: 150px">
    <div class="foldertitlebar">
        <span class="dialogtitle">
            <% Sendb(Copient.PhraseLib.Lookup("term.newfromtemp", LanguageID))%></span> <span
                class="dialogclose" onclick="toggleDialog('DuplicateNoofOffer', false);">X</span>
    </div>
    <div class="dialogcontents">
        <div id="DuplicateOffererror" style="display: none; color: red;">
        </div>
        <table style="width: 90%">
            </tr><td>&nbsp;
            </td>
            </tr>
      <tr>
          <td>
              <label for="infoStart">
                  <% Sendb(Copient.PhraseLib.Lookup("term.duplicateOfferstoCreate", LanguageID).Replace("99", MyCommon.NZ(MyCommon.Fetch_SystemOption(184), 0).ToString()))%></label>
              <input type="text" style="width: 20px" id="txtDuplicateOffersCnt" name="txtDuplicateOffersCnt"
                  maxlength="2" value="" />
          </td>
      </tr>
            </tr><td>&nbsp;
            </td>
            </tr>
      <tr align="right">
          <td>
              <input type="button" name="btnOk" id="btnOk" value="<% Sendb(Copient.PhraseLib.Lookup("term.ok", LanguageID))%>"
                  onclick="addDuplicateOfferscount();" />
              <input type="button" name="btnCancel" id="btnCancel" value="<% Sendb(Copient.PhraseLib.Lookup("term.cancel", LanguageID))%>"
                  onclick="toggleDialog('DuplicateNoofOffer', false);" />
          </td>
      </tr>
        </table>
    </div>
</div>
<!-- #Include virtual="/include/graphic-reward.inc" -->
<script runat="server">
  
    Sub Send_Preference_Reward_Info(ByRef Common As Copient.CommonInc, ByVal IncentivePrefsID As Integer, ByVal ROID As Long, ByVal TierLevels As Integer)
        Dim dt, dt2 As DataTable
        Dim PreferenceID As Long = 0
        Dim i As Integer = 0
        Dim CellCount As Integer = 0

    
        ' find all the tiers in this preference reward
        Common.QueryStr = "select CIPT.PreferenceRewardTierID, CIPT.TierLevel, CIP.PreferenceID " & _
                          "from PreferenceRewardTiers as CIPT with (NoLock) " & _
                          "inner join PreferenceRewards as CIP with (NoLock) on CIP.PreferenceRewardID = CIPT.PreferenceRewardID " & _
                          "where CIPT.PreferenceRewardID=" & IncentivePrefsID & " and CIP.RewardOptionID=" & ROID & " " & _
                          "order by CIPT.TierLevel;"
        dt = Common.LRT_Select
    
        If dt.Rows.Count > 0 Then
            For Each row As DataRow In dt.Rows
                PreferenceID = Common.NZ(row.Item("PreferenceID"), 0)
      
                CellCount += 1
                If CellCount > TierLevels Then Exit For

                Send("<br />")
                If TierLevels > 1 Then
                    Send(Copient.PhraseLib.Lookup("term.tier", LanguageID) & CellCount & ": ")
                End If
        
                ' find all the tier values
                Common.QueryStr = "select IPTV.PreferenceRewardTierValueID, IPTV.PreferenceValue  " & _
                                  "from PreferenceRewardTierValues as IPTV with (NoLock) " & _
                                  "where IPTV.PreferenceRewardTierID=" & Common.NZ(row.Item("PreferenceRewardTierID"), 0)
                dt2 = Common.LRT_Select
                If dt2.Rows.Count > 1 Then
                    Send(Copient.PhraseLib.Lookup("term.multiple", LanguageID))
                Else
                    For i = 0 To dt2.Rows.Count - 1
                        Send(Get_Preference_Value(Common, PreferenceID, Common.NZ(dt2.Rows(i).Item("PreferenceValue"), "")))
                    Next
                End If
        
            Next
      
            ' account for any tiers that don't have saved information due to increasing the tiers on an existing offer
            For i = CellCount To (TierLevels - 1)
                If i > 0 AndAlso i <= (TierLevels - 1) Then Send("<br />")
                Send(Copient.PhraseLib.Lookup("term.tier", LanguageID) & (i + 1) & ": ")
                Send(Copient.PhraseLib.Lookup("term.undefined", LanguageID))
            Next
      
        End If
      
    End Sub

    Sub WriteComponent(ByRef MyCommon As Copient.CommonInc, ByVal rowComp As DataRow, ByRef ComponentColor As String, ByVal Popup As Boolean)
        Dim RecordType As String = ""
        Dim ID As Integer
        Dim StoredProcName As String = ""
        Dim IDParmName As String = ""
        Dim TypeCode As String = ""
        Dim PageName As String = ""
        Dim dtValid As DataTable
        Dim rowOK(), rowWatches(), rowWarnings() As DataRow
        Dim objTemp As Object
        Dim GraceHours As Integer
        Dim GraceCount As Double
        Dim ShowSubReport As Boolean = True
        Dim ShowLinks As Boolean
    
        If Popup Then
            ShowLinks = False
        Else
            ShowLinks = True
        End If
    
        objTemp = MyCommon.Fetch_UE_SystemOption(41)
        If Not (Integer.TryParse(objTemp.ToString, GraceHours)) Then
            GraceHours = 4
        End If
    
        objTemp = MyCommon.Fetch_UE_SystemOption(42)
        If Not (Double.TryParse(objTemp.ToString, GraceCount)) Then
            GraceCount = 0.1D
        End If
    
        RecordType = MyCommon.NZ(rowComp.Item("RecordType"), "")
        ID = MyCommon.NZ(rowComp.Item("ID"), -1)
    
        Select Case RecordType
            Case "term.customergroup"
                StoredProcName = "dbo.pa_ValidationReport_CustGroup"
                IDParmName = "@CustomerGroupID"
                TypeCode = "cg"
                PageName = "../cgroup-edit.aspx?CustomerGroupID="
                ShowSubReport = IIf(ID = 1 OrElse ID = 2, False, True)
            Case "term.productgroup"
                StoredProcName = "dbo.pa_ValidationReport_ProdGroup"
                IDParmName = "@ProductGroupID"
                TypeCode = "pg"
                PageName = "../pgroup-edit.aspx?ProductGroupID="
                ShowSubReport = IIf(ID = 1, False, True)
            Case "term.graphics"
                StoredProcName = "dbo.pa_ValidationReport_Graphic"
                IDParmName = "@OnScreenAdID"
                TypeCode = "gr"
                PageName = "../graphic-edit.aspx?OnScreenAdID="
        End Select
    
        MyCommon.QueryStr = StoredProcName
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add(IDParmName, SqlDbType.Int).Value = ID
        MyCommon.LRTsp.Parameters.Add("@GraceHours", SqlDbType.Int).Value = GraceHours
        MyCommon.LRTsp.Parameters.Add("@GraceCount", SqlDbType.Decimal, 2).Value = GraceCount
    
        dtValid = MyCommon.LRTsp_select()
    
        rowOK = dtValid.Select("Status=0", "LocationName")
        rowWatches = dtValid.Select("Status=1", "LocationName")
        rowWarnings = dtValid.Select("Status=2", "LocationName")
    
        If (ShowSubReport AndAlso ComponentColor <> "red") Then
            ComponentColor = IIf(rowWarnings.Length > 0, "red", "green")
        End If
    
        Send("<div style=""margin-left:10px;"">")
        Sendb(Copient.PhraseLib.Lookup(MyCommon.NZ(rowComp.Item("RecordType"), ""), LanguageID) & " #" & ID & ": ")
        If (ShowSubReport) Then
            If (ShowLinks) Then
                Send("<a href=""" & PageName & ID & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(rowComp.Item("Name"), "&nbsp;"), 20) & "</a>")
            Else
                Send(MyCommon.SplitNonSpacedString(MyCommon.NZ(rowComp.Item("Name"), "&nbsp;"), 20))
            End If
        Else
            Send(MyCommon.SplitNonSpacedString(MyCommon.NZ(rowComp.Item("Name"), "&nbsp;"), 20))
        End If
    
        If (ShowSubReport) Then
            Send("<div style=""margin-left:20px;"">")
            If ShowLinks Then
                Send("<a id=""validLink" & ID & """ href=""javascript:openPopup('../validation-report.aspx?type=" & TypeCode & "&id=" & ID & "&level=0');"">")
                Send("  " & Copient.PhraseLib.Lookup("cgroup-edit.validlocations", LanguageID) & " (" & rowOK.Length & ")")
                Send("</a><br />")
                Send("<a id=""watchLink" & ID & """ href=""javascript:openPopup('../validation-report.aspx?type=" & TypeCode & "&id=" & ID & "&level=1');"">")
                Send("  " & Copient.PhraseLib.Lookup("cgroup-edit.watchlocations", LanguageID) & " (" & rowWatches.Length & ")")
                Send("</a><br />")
                Send("<a id=""warningLink" & ID & """ href=""javascript:openPopup('../validation-report.aspx?type=" & TypeCode & "&id=" & ID & "&level=2');"">")
                Send("  " & Copient.PhraseLib.Lookup("cgroup-edit.warninglocations", LanguageID) & " (" & rowWarnings.Length & ")")
                Send("</a><br />")
            Else
                Send(Copient.PhraseLib.Lookup("cgroup-edit.validlocations", LanguageID) & " (" & rowOK.Length & ")<br />")
                Send(Copient.PhraseLib.Lookup("cgroup-edit.watchlocations", LanguageID) & " (" & rowWatches.Length & ")<br />")
                Send(Copient.PhraseLib.Lookup("cgroup-edit.warninglocations", LanguageID) & " (" & rowWarnings.Length & ")<br />")
            End If

            Send("</div>")
        End If
    
        Send("</div>")
        MyCommon.Close_LogixRT()
    End Sub
  
    Sub RemoveInactiveLocations(ByRef MyCommon As Copient.CommonInc, ByRef dt As DataTable, ByVal IncentiveID As Integer)
        Dim dtLoc As DataTable
        Dim row As DataRow
        Dim LocationID As Integer
        Dim LocGroupList As String = ""
        Dim IsAllLocs As Boolean
        Dim LocTable As New Hashtable
        If (Not dt Is Nothing And dt.Rows.Count > 0) Then
            ' first find all locations is selected
            MyCommon.QueryStr = "select LocationGroupID from OfferLocations with (NoLock) where OfferID=" & IncentiveID & " and Deleted=0;"
            dtLoc = MyCommon.LRT_Select
            For Each row In dtLoc.Rows
                If (LocGroupList <> "") Then LocGroupList += ","
                LocGroupList = LocGroupList + MyCommon.NZ(row.Item("LocationGroupID"), "-1").ToString
                IsAllLocs = MyCommon.NZ(row.Item("LocationGroupID"), -1) = 1
            Next
            If (Not IsAllLocs AndAlso LocGroupList <> "") Then
                ' find all the locations for the given location groups
                MyCommon.QueryStr = "select LocationID from LocGroupItems with (NoLock) where Deleted = 0 " & _
                                    "and LocationGroupID in (" & LocGroupList & ");"
                dtLoc = MyCommon.LRT_Select
                For Each row In dtLoc.Rows
                    LocationID = MyCommon.NZ(row.Item("LocationID"), "-1")
                    If (Not LocTable.ContainsKey(LocationID.ToString)) Then
                        LocTable.Add(MyCommon.NZ(row.Item("LocationID"), "-1").ToString, MyCommon.NZ(row.Item("LocationID"), "-1").ToString)
                    End If
                Next
                ' remove the location if it doesn't currently exist for the incentive
                For Each row In dt.Rows
                    LocationID = MyCommon.NZ(row.Item("LocationID"), "-1")
                    If (Not LocTable.ContainsKey(LocationID.ToString)) Then
                        row.Delete()
                    End If
                Next
            End If
        End If
    End Sub
  
</script>
<script type="text/javascript">
    collapseBoxes();
    setComponentsColor('<% Sendb(ComponentColor)%>');
</script>
<script type="text/javascript">
    if (window.captureEvents){
        window.captureEvents(Event.CLICK);
        window.onclick=handlePageClick;
    } else {
        document.onclick=handlePageClick;
    }
    <%
    If (StatusMessage <> "") Then
        Send("alert('" & StatusMessage & "');")
    End If
  %>
</script>
<script runat="server">
    Dim exclusionPGList As List(Of ProductConditionProductGroup)
    Dim resultExcludedPGList As AMSResult(Of List(Of ProductConditionProductGroup))

    Sub DisplayExclusionGroups(ByVal productConditionPGService As IProductConditionService, ByVal productConditionId As Integer, ByRef infoMessage As String, ByRef MyCommon As Copient.CommonInc, ByVal languageId As Integer)
        Dim sb As New StringBuilder()
        Dim extBuyerId As string
        resultExcludedPGList = productConditionPGService.GetExclusionProductGroups(productConditionId)
        If (resultExcludedPGList.ResultType = AMSResultType.Success) Then
            exclusionPGList = resultExcludedPGList.Result
            If exclusionPGList.Count > 0 Then
                Sendb("<li>" & Copient.PhraseLib.Lookup("term.excluded", languageId) & ": ")
                For Each pcpg As ProductConditionProductGroup In exclusionPGList
                    If (MyCommon.IsEngineInstalled(9) And MyCommon.Fetch_UE_SystemOption(168) = "1" And pcpg.BuyerId > 0) Then
                        extBuyerId = MyCommon.GetExternalBuyerId(pcpg.BuyerId)
                        sb.Append("<a href=""/logix/pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(pcpg.ProductGroupId, 0) & """>" & "Buyer " & extBuyerId & " - " & MyCommon.NZ(pcpg.ProductGroupName, "") & "</a>, ")
                    Else
                        sb.Append("<a href=""/logix/pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(pcpg.ProductGroupId, 0) & """>" & MyCommon.NZ(pcpg.ProductGroupName, "") & "</a>, ")
                    End If
                Next
                'Remove the last comma 
                sb.Remove(sb.Length - 2, 2)
                Send(sb.ToString())
                Send("</li>")
            End If
        Else
            infoMessage = resultExcludedPGList.PhraseString
        End If
    End Sub
    Dim DiscountexclusionPGList As List(Of DiscountProductGroup)
    Dim resultDiscountExcludedPGList As AMSResult(Of List(Of DiscountProductGroup))

    Sub DisplayDiscountExclusionGroups(ByVal DiscountPGService As IDiscountRewardService, ByVal Discountid As Integer, ByRef infoMessage As String, ByRef MyCommon As Copient.CommonInc, ByVal languageId As Integer)
        Dim sb As New StringBuilder()
        Dim extBuyerId As string
        resultDiscountExcludedPGList = DiscountPGService.GetAllExclusionGroups(Discountid)
        If (resultDiscountExcludedPGList.ResultType = AMSResultType.Success) Then
            DiscountexclusionPGList = resultDiscountExcludedPGList.Result
            If DiscountexclusionPGList.Count > 0 Then
                Sendb("<li>" & Copient.PhraseLib.Lookup("term.excluded", languageId) & ": ")
                For Each dpcpg As DiscountProductGroup In DiscountexclusionPGList
                    If (MyCommon.IsEngineInstalled(9) And MyCommon.Fetch_UE_SystemOption(168) = "1" And dpcpg.BuyerId > 0) Then
                        extBuyerId = MyCommon.GetExternalBuyerId(dpcpg.BuyerId)
                        sb.Append("<a href=""/logix/pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(dpcpg.ProductGroupId, 0) & """>" & "Buyer " & extBuyerId & " - " & MyCommon.NZ(dpcpg.ProductGroupName, "") & "</a>, ")
                    Else
                        sb.Append("<a href=""/logix/pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(dpcpg.ProductGroupId, 0) & """>" & MyCommon.NZ(dpcpg.ProductGroupName, "") & "</a>, ")
                    End If
                Next
                'Remove the last comma 
                sb.Remove(sb.Length - 2, 2)
                Send(sb.ToString())
                Send("</li>")
            End If
        Else
            infoMessage = resultDiscountExcludedPGList.PhraseString
        End If
    End Sub
    Function GetLastDeployValidationMessage(ByRef MyCommon As Copient.CommonInc, ByVal OfferID As Long) As String
        Dim lastDeployValidationMsg As String = String.Empty
        Dim dt As DataTable
        MyCommon.QueryStr = "Select LastDeployValidationMessage From CPE_Incentives " & _
                              " Where IncentiveId=@OfferId"
        MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        If Not dt Is Nothing And dt.Rows.Count = 1 Then
            lastDeployValidationMsg = dt.Rows(0)(0).ToString()
        End If
        Return lastDeployValidationMsg
    End Function
    '-------------------------------------------------------------------------------------------------------------------------------
    Function IsOAWEnabled(ByVal OfferID As Long, ByVal BannerIds As Integer()) As Boolean
        Dim enabled As Boolean = False
        Dim m_OAWService As IOfferApprovalWorkflowService = CurrentRequest.Resolver.Resolve(Of IOfferApprovalWorkflowService)()
        Dim MyCommon As New Copient.CommonInc
        Dim Logix As New Copient.LogixInc
        If (MyCommon.Fetch_SystemOption(66) = "1") Then
            enabled = m_OAWService.CheckIfOfferApprovalIsEnabledForBanners(BannerIds).Result
        Else
            enabled = m_OAWService.CheckIfOfferApprovalIsEnabled().Result
        End If
        Return enabled
    End Function
    '------------------------------------------------------------------------------------------------------------------------------
    Sub SetLastDeployValidationMessage(ByRef MyCommon As Copient.CommonInc, ByVal OfferID As Long, ByVal Message As String)
        MyCommon.QueryStr = "Update CPE_Incentives " & _
                           "  Set LastDeployValidationMessage=@Message " & _
                           "  where IncentiveId=@OfferId"
        MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
        MyCommon.DBParameters.Add("@Message", SqlDbType.NVarChar).Value = Message
        MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
    End Sub

    Function IsDeployableOffer(ByRef MyCommon As Copient.CommonInc, ByVal OfferID As Integer, ByVal ROID As Integer, ByRef ErrorMsg As String, ByRef Logix As Copient.LogixInc) As Boolean
        Dim Deployable As Boolean = False

        ErrorMsg = ""
        Deployable = MeetsDeploymentReqs(MyCommon, OfferID, ErrorMsg)
        If Deployable Then
            Dim m_Offer As IOffer = CMS.AMS.CurrentRequest.Resolver.Resolve(Of IOffer)()
            If Not (m_Offer.IsProdConAvailableIfLimitAsNoLimit(OfferID)) Then
                Deployable = False
                ErrorMsg = Copient.PhraseLib.Lookup("UEoffer-deploy-error.productConditionForNolimit", LanguageID)
            End If
            If Deployable Then
                Dim listProdCondition As AMSResult(Of List(Of RegularProductCondition)) = m_Offer.GetRegularProductConditionsByOfferId(OfferID)
                If (listProdCondition.Result.Count = 0) Then
                    If MyCommon.CheckRewardExistsForOffer(OfferID, 13) Then
                        If Not MyCommon.CheckChargeBackDeptIDExistsOrNot(OfferID) Then
                            Deployable = False
                            ErrorMsg = Copient.PhraseLib.Lookup("UEoffer-deploy-error.invalidchargebackdeptid", LanguageID)
                        End If
                    End If
                End If
                        If Deployable Then
                    Dim listPointsCondition As AMSResult(Of List(Of RegularPointCondition)) = m_Offer.GetRegularPointConditionsByOfferId(OfferID)
                    Dim listSVCondition As AMSResult(Of List(Of RegularSVCondition)) = m_Offer.GetRegularSVConditionsByOfferId(OfferID)
                    If (listProdCondition.Result.Count = 0 AndAlso listPointsCondition.Result.Count = 0 AndAlso listSVCondition.Result.Count = 0) Then
                            If MyCommon.IsGCRMultiTiered(OfferID) = False Then
                                Deployable = False
                                ErrorMsg = Copient.PhraseLib.Lookup("cpeoffer-sum.tier-setup-invalid", LanguageID)
                            End If
                        End If
                    End If

            End If
        End If

        If Deployable Then
            Deployable = MeetsTemplateRequirements(MyCommon, ROID)
            If Deployable Then
                Deployable = MeetsTieredReqs(MyCommon, OfferID)
                If Not Deployable Then
                    ErrorMsg = Copient.PhraseLib.Lookup("UEoffer-sum.tier-setup-invalid", LanguageID)
                End If
            Else
                ErrorMsg = Copient.PhraseLib.Lookup("offer-sum.required-incomplete", LanguageID)
            End If
            If GetCgiValue("deferdeploy") <> "" Then
                If (Deployable AndAlso MyCommon.Fetch_SystemOption(260) = "1" AndAlso Not Logix.UserRoles.DeferDeployOffersPastLockoutPeriod AndAlso MyCommon.IsLockOutPeriod(MyCommon, OfferID) = True) Then
                    Deployable = False
                    ErrorMsg = Copient.PhraseLib.Lookup("cpeoffer-sum.deferdeployalertforlockout", LanguageID)
                End If
            Else
                If (Deployable AndAlso MyCommon.Fetch_SystemOption(131) = "1" AndAlso MyCommon.Fetch_SystemOption(67) = "0") Then
                    Deployable = MeetsLockOutRequirement(Logix, MyCommon, OfferID)
                    If (Not Deployable) Then
                        ErrorMsg = Copient.PhraseLib.Lookup("cpeoffer-sum.deployalertforlockout", LanguageID)
                    End If
                End If
            End If
        End If

        If Deployable Then Deployable = MeetsLocalizationReqs(MyCommon, ROID, ErrorMsg)
        If GetCgiValue("deploy") <> "Yes" Then
            If GetCgiValue("deferdeploy") <> "Yes" Then
                If Deployable Then Deployable = MeetsTranslationRequirements(MyCommon, OfferID, ROID, ErrorMsg)
            End If
        End If
        Return Deployable
    End Function

    '-----------------------------------------------------------------------------------------------------------------------------

    Function MeetsDeploymentReqs(ByRef MyCommon As Copient.CommonInc, ByVal OfferID As Integer, ByRef ErrorMsg As String) As Boolean
        Dim bMeetsReqs As Boolean = False

        ' The user wants to deploy, so do a quick check for at least one assigned offer location and terminal,
        ' and ensure that there are no unassigned tier values
        MyCommon.QueryStr = "dbo.pa_UE_IsOfferDeployable"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
        MyCommon.LRTsp.Parameters.Add("@IsDeployable", SqlDbType.Bit).Direction = ParameterDirection.Output
        MyCommon.LRTsp.Parameters.Add("@ReturnPhraseName", SqlDbType.NVarChar, 255).Direction = ParameterDirection.Output
        MyCommon.LRTsp.Parameters.Add("@DoAnalyticsValidation", SqlDbType.Bit).Value = 0
        MyCommon.LRTsp.ExecuteNonQuery()
        bMeetsReqs = MyCommon.LRTsp.Parameters("@IsDeployable").Value
        Dim ReturnPhraseName As String = MyCommon.LRTsp.Parameters("@ReturnPhraseName").Value
        If ReturnPhraseName.Trim() <> String.Empty Then
            ErrorMsg = Copient.PhraseLib.Lookup(ReturnPhraseName, LanguageID)
        End If

        Return bMeetsReqs
    End Function

    '-----------------------------------------------------------------------------------------------------------------------------
    Function MeetsLockOutRequirement(ByRef Logix As Copient.LogixInc, ByRef MyCommon As Copient.CommonInc, ByVal OfferId As Integer) As Boolean

        Dim isDeployable As Boolean = True
        Dim bInLockoutPeriod As Boolean = False

        If Logix.UserRoles.DeployOffersPastLockoutDate = False Then
            isDeployable = IIf(MyCommon.IsLockOutPeriod(MyCommon, OfferId), False, True)
        End If

        Return isDeployable

    End Function

    Function MeetsTemplateRequirements(ByRef MyCommon As Copient.CommonInc, ByVal ROID As Integer) As Boolean
        Dim dt As DataTable

        MyCommon.QueryStr = "select 'CG' as GroupType, CustomerGroupID as GroupID from CPE_IncentiveCustomerGroups with (NoLock) " & _
                            "where RewardOptionID = " & ROID & " and Deleted=0 and RequiredFromTemplate=1 and CustomerGroupID is null " & _
                            "union " & _
                            "select 'PG' as GroupType, ISNULL(ProductGroupID, -1) as GroupID from CPE_IncentiveProductGroups with (NoLock) " & _
                            "where RewardOptionID = " & ROID & " and Deleted=0 and RequiredFromTemplate=1 and (ProductGroupID =-1 OR ProductGroupID IS NULL) " & _
                            "union " & _
                            "select 'PP' as GroupType,ProgramID as GroupID from CPE_IncentivePointsGroups with (NoLock) " & _
                            "where RewardOptionID = " & ROID & " and Deleted=0 and RequiredFromTemplate=1 and ProgramId is null; "
        dt = MyCommon.LRT_Select

        Return (dt.Rows.Count = 0)
    End Function

    '-----------------------------------------------------------------------------------------------------------------------------

    Function MeetsTieredReqs(ByRef MyCommon As Copient.CommonInc, ByVal OfferID As Long) As Boolean
        Dim bMeetsReqs As Boolean = False

        ' ensure that multi-tiered offers have a tierable condition if there is a tierable reward assigned to the offer.
        MyCommon.QueryStr = "dbo.pa_ValidateOfferTiers"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
        MyCommon.LRTsp.Parameters.Add("@IsDeployable", SqlDbType.Bit).Direction = ParameterDirection.Output
        MyCommon.LRTsp.ExecuteNonQuery()
        bMeetsReqs = MyCommon.LRTsp.Parameters("@IsDeployable").Value

        Return bMeetsReqs
    End Function

    '----------------------------------------------------------------------------------------------------------------------------------------------

    Function MeetsLocalizationReqs(ByRef Common As Copient.CommonInc, ByVal ROID As Integer, ByRef ErrorMsg As String) As Boolean

        Dim ReturnVal As Boolean = True
        Dim dst As DataTable
        Dim row As DataRow
        Dim ErrorTextList As String = ""
        Dim LegacyOffer As Boolean = False

        ErrorMsg = ""
        'See if this is a legacy offer, or one that is using localzation.  
        'A legacy offer is one where RewardOptions.CurrencyID=0 and there are either no CPE_RewardOptionUOMs records for the RewardOptionID, 
        'or all of the related CPE_RewardOptionUOMs records have a value of zero for thier UOMSubTypeID

        'If this query returns zero rows, then this is a legacy offer.  If it returns >0 rows, then something (currency or UOM) has been 
        'set for this offer and it must have compliant localization data to be deployed
        Common.QueryStr = "select RewardOptionID " & _
                          "  from CPE_RewardOptions  " & _
                          "  where CurrencyID>0 And RewardOptionID=" & ROID & " " & _
                          "Union " & _
                          "Select RewardOptionID " & _
                          "  from CPE_RewardOptionUOMs " & _
                          "  where UOMSubTypeID>0 and RewardOptionID=" & ROID & ";"
        dst = Common.LRT_Select
        If Not (dst.Rows.Count > 0) Then LegacyOffer = True
        dst = Nothing

        If Not (LegacyOffer) Then

            If Common.Fetch_UE_SystemOption(136) = "1" Then  'see if multi-currency is enabled
                'query to see if the Currency for the offer has been set on the offer general page
                Common.QueryStr = "select CurrencyID from CPE_RewardOptions where RewardOptionID=" & ROID & " and CurrencyID>0;"
                dst = Common.LRT_Select
                If Not (dst.Rows.Count > 0) Then
                    ReturnVal = False
                    ErrorMsg = Copient.PhraseLib.Lookup("offer-sum.nonselectedcurrency", LanguageID) 'The offer can not be deployed until a currency has been selected
                End If
                dst = Nothing
            End If 'mulit-currency is enabled

            If ReturnVal = True And Common.Fetch_UE_SystemOption(135) = "1" Then  'is multi-uom enabled?  no need to run this query if there has already been a violation of the requirements
                'query to see that all of the UOMTypes have been specifed for this offer - if not, they need to be selected 
                'on the offer general page before the offer can be deployed.
                Common.QueryStr = "select UOMT.UOMTypeID, UOMT.PhraseTerm " & _
                                  "from UOMTypes as UOMT left Join CPE_RewardOptionUOMs as ROUOM with (NoLock) on UOMT.UOMTypeID=ROUOM.UOMTypeID and ROUOM.RewardOptionID=" & ROID & " " & _
                                  "where isnull(ROUOM.UOMSubTypeID, 0)=0;"
                dst = Common.LRT_Select
                If dst.Rows.Count > 0 Then
                    ReturnVal = False
                    ErrorMsg = Copient.PhraseLib.Lookup("offer-sum.undefineduom", LanguageID) 'The offer can not be deployed until the unit of measure has been selected for the following type(s): 
                    ErrorTextList = ""
                    For Each row In dst.Rows
                        If Not (ErrorTextList = "") Then ErrorTextList = ErrorTextList & ", "
                        ErrorTextList = ErrorTextList & Copient.PhraseLib.Lookup(row.Item("PhraseTerm"), LanguageID)
                    Next
                    If Len(ErrorTextList) > 500 Then ErrorTextList = Left(ErrorTextList, 500) & " ..."
                    ErrorMsg = ErrorMsg & ErrorTextList
                End If
                dst = Nothing
            End If

            If ReturnVal = True And Common.Fetch_UE_SystemOption(136) = "1" Then 'is multi-currency enabled?  no need to run this query if there has already been a violation of the requirements
                'query to see if there are any locations joined to the offer that use a currencyID that is different than RewardOptions.CurrencyID 
                Common.QueryStr = "select L.LocationID, isnull(L.LocationName, '') as LocationName " & _
                                  "from Locations as L with (NoLock) Inner Join LocGroupItems as LGI with (NoLock) on L.LocationID=LGI.LocationID and LGI.Deleted=0 and L.Deleted=0 " & _
                                  "Inner Join OfferLocations as OL with (NoLock) on OL.LocationGroupID=LGI.LocationGroupID and OL.Deleted=0 " & _
                                  "Inner Join CPE_RewardOptions as RO with (NoLock) on RO.IncentiveID=OL.OfferID and RO.Deleted=0 and RO.TouchResponse=0 " & _
                                  "Where RO.RewardOptionID=" & ROID & " and L.CurrencyID<>RO.CurrencyID and RO.CurrencyID>0;"
                dst = Common.LRT_Select
                If dst.Rows.Count > 0 Then
                    ReturnVal = False
                    ErrorMsg = Copient.PhraseLib.Lookup("offer-sum.unsupportedcurrency", LanguageID) 'The following targeted locations do not support the currency selected for this offer: 
                    ErrorTextList = ""
                    For Each row In dst.Rows
                        If Not (ErrorTextList = "") Then ErrorTextList = ErrorTextList & ", "
                        ErrorTextList = ErrorTextList & row.Item("LocationName")
                    Next
                    If Len(ErrorTextList) > 500 Then ErrorTextList = Left(ErrorTextList, 500) & " ..."
                    ErrorMsg = ErrorMsg & ErrorTextList
                End If
                dst = Nothing
            End If

            If ReturnVal = True And Common.Fetch_UE_SystemOption(135) = "1" Then 'is multi-uom enabled?  no need to run this query if there has already been a violation of the requirements
                'query to see if there are any locations that do not support the units of measure that are used by this offer
                Common.QueryStr = "dbo.pa_InvalidLocationUOMsForOffer"
                Common.Open_LRTsp()
                Common.LRTsp.Parameters.Add("@RewardOptionID", SqlDbType.BigInt).Value = ROID
                dst = Common.LRTsp_select
                If Not (dst Is Nothing) AndAlso dst.Rows.Count > 0 Then
                    ReturnVal = False
                    ErrorMsg = Copient.PhraseLib.Lookup("offer-sum.unsupporteduom", LanguageID) 'The following targeted locations do not support the units of measure selected for this offer: 
                    ErrorTextList = ""
                    For Each row In dst.Rows
                        If Not (ErrorTextList = "") Then ErrorTextList = ErrorTextList & ", "
                        ErrorTextList = ErrorTextList & row.Item("LocationName")
                    Next
                    If Len(ErrorTextList) > 500 Then ErrorTextList = Left(ErrorTextList, 500) & " ..."
                    ErrorMsg = ErrorMsg & ErrorTextList
                End If
            End If


            If ReturnVal = True Then
                ReturnVal = MeetsCurrencyPrecision(Common, ROID, ErrorMsg)
            End If

            If ReturnVal = True Then
                ReturnVal = MeetsUOMPrecision(Common, ROID, ErrorMsg)
            End If



        End If  'not(LegacyOffer)
        Return ReturnVal

    End Function

    '-----------------------------------------------------------------------------------------------------------------------------

    Function MeetsCurrencyPrecision(ByRef Common As Copient.CommonInc, ByVal ROID As Long, ByRef ErrorMsg As String) As Boolean
        Dim ReturnVal As Boolean = True
        Dim ErrorTerm As String = ""
        Dim CurrencyName As String = ""
        Dim dt As DataTable

        ErrorMsg = ""
        If Common.Fetch_UE_SystemOption(136) = "1" Then  'see if multi-currency is enabled
            Common.QueryStr = "dbo.pa_CurrencyPrecisionCheck"
            Common.Open_LRTsp()
            Common.LRTsp.Parameters.Add("@ROID", SqlDbType.BigInt).Value = ROID
            Common.LRTsp.Parameters.Add("@ErrorTerm", SqlDbType.VarChar, 200).Direction = ParameterDirection.Output
            Common.LRTsp.ExecuteNonQuery()
            ErrorTerm = Common.LRTsp.Parameters("@ErrorTerm").Value
            Common.Close_LRTsp()
            If Not (ErrorTerm = "") Then
                Common.QueryStr = "Select C.NamePhraseTerm " & _
                                  "from Currencies as C Inner Join CPE_RewardOptions as RO with (NoLock) on C.CurrencyID=RO.CurrencyID " & _
                                  "where RO.RewardOptionID=" & ROID & ";"
                dt = Common.LRT_Select
                If dt.Rows.Count > 0 Then
                    CurrencyName = Copient.PhraseLib.Lookup(dt.Rows(0).Item("NamePhraseTerm"), LanguageID)
                End If
                ErrorMsg = Copient.PhraseLib.Lookup("offer-sum.currprecisionallowed", LanguageID) & " " & CurrencyName & " " & Copient.PhraseLib.Lookup("offer-sum.currhasbeenexceeded", LanguageID) & " " & Copient.PhraseLib.Lookup(ErrorTerm, LanguageID)  'The currency precision allowed for CURRENCY NAME has been exceeded by one or more fields in: 
                ReturnVal = False
            End If
        End If 'mulit-currency is enabled

        Return ReturnVal

    End Function

    '-----------------------------------------------------------------------------------------------------------------------------

    Function MeetsUOMPrecision(ByRef Common As Copient.CommonInc, ByVal ROID As Long, ByRef ErrorMsg As String) As Boolean
        Dim ReturnVal As Boolean = True
        Dim ErrorTerm As String = ""
        Dim ErrorUOMSubTypeID As Integer = 0
        Dim MeasureName As String = ""
        Dim dt As DataTable

        ErrorMsg = ""
        If Common.Fetch_UE_SystemOption(135) = "1" Then  'see if multi-UOM is enabled
            Common.QueryStr = "dbo.pa_UOMPrecisionCheck"
            Common.Open_LRTsp()
            Common.LRTsp.Parameters.Add("@ROID", SqlDbType.BigInt).Value = ROID
            Common.LRTsp.Parameters.Add("@ErrorTerm", SqlDbType.VarChar, 200).Direction = ParameterDirection.Output
            Common.LRTsp.Parameters.Add("@ErrorUOMSubTypeID", SqlDbType.Int).Direction = ParameterDirection.Output
            Common.LRTsp.ExecuteNonQuery()
            ErrorTerm = Common.LRTsp.Parameters("@ErrorTerm").Value
            ErrorUOMSubTypeID = Common.LRTsp.Parameters("@ErrorUOMSubTypeID").Value
            Common.Close_LRTsp()
            If Not (ErrorTerm = "") Then
                Common.QueryStr = "Select UOMST.NamePhraseTerm " & _
                                  "from UOMSubTypes as UOMST " & _
                                  "where UOMST.UomSubTypeID=" & ErrorUOMSubTypeID & ";"
                dt = Common.LRT_Select
                If dt.Rows.Count > 0 Then
                    MeasureName = Copient.PhraseLib.Lookup(dt.Rows(0).Item("NamePhraseTerm"), LanguageID)  'pounds, kilograms, gallons, liters, etc.
                End If
                ErrorMsg = Copient.PhraseLib.Lookup("offer-sum.uomprecisionallowed", LanguageID) & " " & MeasureName & " " & Copient.PhraseLib.Lookup("offer-sum.uomhasbeenexceeded", LanguageID) & " " & Copient.PhraseLib.Lookup(ErrorTerm, LanguageID)  'The precision allowed for MEASURE NAME has been exceeded by one or more fields in: 
                ReturnVal = False
            End If
        End If 'multi-UOM enabled

        Return ReturnVal

    End Function

    '-----------------------------------------------------------------------------------------------------------------------------

    Function MeetsTranslationRequirements(ByRef Common As Copient.CommonInc, ByVal OfferID As Long, ByVal ROID As Long, ByRef ErrorMsg As String) As Boolean
        Dim ReturnVal As Boolean = True
        Dim ErrorTerm As String = ""
        Dim TranslatedErrorTerm As String = ""
        Dim DeployType As String = ""

        ErrorMsg = ""
        If Common.Fetch_SystemOption(124) = "1" Then  'see if multi-language is enabled
            ErrorTerm = CheckForTranslationDeployError(Common, ROID)
            If Not (ErrorTerm = "") Then
                'if this section weren't in place, then you could get around the warning message by clicking defer deploy and then clicking deploy, without specifically addressing the warning
                If GetCgiValue("deferdeploy") <> "" Then
                    DeployType = "deferdeploy"
                Else
                    DeployType = "deploy"
                End If
                TranslatedErrorTerm = Copient.PhraseLib.DecodeEmbededTokens(ErrorTerm, LanguageID)
                ErrorMsg = Copient.PhraseLib.Detokenize("term.ReqTransFailed", LanguageID, TranslatedErrorTerm) & _
                  " <input type=""submit"" class=""regular"" id=""" & DeployType & """ name=""" & DeployType & """ value=""" & Copient.PhraseLib.Lookup("term.yes", LanguageID) & """ onclick=""document.getElementById('deploytransreqskip').value='1';"" />"

                ReturnVal = False
            End If
        End If  'multi-language enabled

        Return ReturnVal

    End Function

    '-----------------------------------------------------------------------------------------------------------------------------

    Function CheckForTranslationDeployError(ByRef Common As Copient.CommonInc, ByVal ROID As Long) As String
        Dim ReturnVal As String
        Common.QueryStr = "dbo.pa_ReqTranslationCheck"
        Common.Open_LRTsp()
        Common.LRTsp.Parameters.Add("@ROID", SqlDbType.BigInt).Value = ROID
        Common.LRTsp.Parameters.Add("@ErrorTerm", SqlDbType.VarChar, 2000).Direction = ParameterDirection.Output
        Common.LRTsp.ExecuteNonQuery()
        ReturnVal = Common.LRTsp.Parameters("@ErrorTerm").Value
        Common.Close_LRTsp()
        Return ReturnVal
    End Function

    '-----------------------------------------------------------------------------------------------------------------------------

    Sub Send_Preference_Details(ByRef Common As Copient.CommonInc, ByVal PreferenceID As Long)
        Dim dt As DataTable
        Dim IntegrationVals As New Copient.CommonInc.IntegrationValues
        Dim PrefPageName As String = ""
        Dim Tokens As String = ""
        Dim RootURI As String = ""

        Common.QueryStr = "select UserCreated, Name as PrefName " & _
                          "from Preferences as PREF with (NoLock) " & _
                          "where PREF.PreferenceID=" & PreferenceID & " and PREF.Deleted=0"
        dt = Common.PMRT_Select
        If dt.Rows.Count > 0 Then
            If (Common.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER, IntegrationVals)) Then
                PrefPageName = IIf(Common.NZ(dt.Rows(0).Item("UserCreated"), False), "prefscustom-edit.aspx", "prefsstd-edit.aspx")

                RootURI = IntegrationVals.HTTP_RootURI
                If RootURI IsNot Nothing AndAlso RootURI.Length > 0 AndAlso Right(RootURI, 1) <> "/" Then
                    RootURI &= "/"
                End If

                Tokens = "SendToURI="
                Sendb("  <a href=""../authtransfer.aspx?SendToURI=" & RootURI & "UI/" & PrefPageName & "?prefid=" & PreferenceID & """>")
                Send(Common.NZ(dt.Rows(0).Item("PrefName"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)) & "</a>")
            End If
        End If

    End Sub

    '-----------------------------------------------------------------------------------------------------------------------------

    Sub Send_Preference_Info(ByRef Common As Copient.CommonInc, ByVal IncentivePrefsID As Integer, ByVal ROID As Long, ByVal TierLevels As Integer)
        Dim dt, dt2 As DataTable
        Dim PreferenceID As Long = 0
        Dim ComboText As String = ""
        Dim i As Integer = 0
        Dim CellCount As Integer = 0


        ' find all the tiers in this preference condition
        Common.QueryStr = "select CIPT.IncentivePrefTiersID, CIPT.TierLevel, CIPT.ValueComboTypeID, CIP.PreferenceID " & _
                          "from CPE_IncentivePrefTiers as CIPT with (NoLock) " & _
                          "inner join CPE_IncentivePrefs as CIP with (NoLock) on CIP.IncentivePrefsID = CIPT.IncentivePrefsID " & _
                          "where CIPT.IncentivePrefsID=" & IncentivePrefsID & " and CIP.RewardOptionID=" & ROID & " " & _
                          "order by CIPT.TierLevel;"
        dt = Common.LRT_Select

        If dt.Rows.Count > 0 Then
            For Each row As DataRow In dt.Rows
                PreferenceID = Common.NZ(row.Item("PreferenceID"), 0)
                ComboText = IIf(Common.NZ(row.Item("ValueComboTypeID"), 1) = 1, "term.and", "term.or")
                ComboText = Copient.PhraseLib.Lookup(ComboText, LanguageID)

                CellCount += 1
                If CellCount > TierLevels Then Exit For

                Send("<br />")
                If TierLevels > 1 Then
                    Send(Copient.PhraseLib.Lookup("term.tier", LanguageID) & CellCount & ": ")
                End If

                ' find all the tier values
                Common.QueryStr = "select IPTV.PKID, IPTV.Value, IPTV.DateOperatorTypeID, " & _
                                  "  case when POT.PhraseID is null then POT.Description" & _
                                  "  else Convert(nvarchar(200), PT.Phrase) end as OperatorText " & _
                                  "from CPE_IncentivePrefTierValues as IPTV with (NoLock) " & _
                                  "inner join CPE_PrefOperatorTypes as POT with (NoLock) on POT.PrefOperatorTypeID = IPTV.OperatorTypeID " & _
                                  "left join PhraseText as PT with (NoLock) on PT.PhraseID = POT.PhraseID and LanguageID=" & LanguageID & " " & _
                                  "where IPTV.IncentivePrefTiersID=" & Common.NZ(row.Item("IncentivePrefTiersID"), 0)
                dt2 = Common.LRT_Select
                For i = 0 To dt2.Rows.Count - 1
                    If Common.NZ(dt2.Rows(i).Item("DateOperatorTypeID"), 0) > 0 Then
                        Send(Get_Date_Display_Text(Common, dt2.Rows(i).Item("PKID")))
                    Else
                        Send(Common.NZ(dt2.Rows(i).Item("OperatorText"), "") & " " & Get_Preference_Value(Common, PreferenceID, Common.NZ(dt2.Rows(i).Item("Value"), "")))
                    End If

                    If i < dt2.Rows.Count - 1 Then
                        Send(" <i>" & ComboText.ToLower & "</i> ")
                    End If
                Next
            Next

            ' account for any tiers that don't have saved information due to increasing the tiers on an existing offer
            For i = CellCount To (TierLevels - 1)
                If i > 0 AndAlso i <= (TierLevels - 1) Then Send("<br />")
                Send(Copient.PhraseLib.Lookup("term.tier", LanguageID) & (i + 1) & ": ")
                Send(Copient.PhraseLib.Lookup("term.undefined", LanguageID))
            Next

        End If

    End Sub


    '-----------------------------------------------------------------------------------------------------------------------------


    Function Get_Date_Display_Text(ByRef Common As Copient.CommonInc, ByVal TierValuePKID As Integer) As String
        Dim DisplayText As String = ""
        Dim dt As DataTable
        Dim ValueModifier As String = ""
        Dim Offset As Integer

        Common.QueryStr = "select IPTV.Value, IPTV.ValueModifier, IPTV.ValueTypeID, POT.PhraseID as OperatorPhraseID, " & _
                          "PDOT.PhraseID as DateOpPhraseID " & _
                          "from CPE_IncentivePrefTierValues as IPTV with (NoLock) " & _
                          "inner join CPE_PrefOperatorTypes as POT with (NoLock) on POT.PrefOperatorTypeID = IPTV.OperatorTypeID " & _
                          "inner join CPE_PrefDateOperatorTypes as PDOT with (NoLock) on PDOT.PrefDateOperatorTypeID = IPTV.DateOperatorTypeID " & _
                          "where PKID=" & TierValuePKID & ";"
        dt = Common.LRT_Select
        If dt.Rows.Count > 0 Then
            DisplayText = Copient.PhraseLib.Lookup(Common.NZ(dt.Rows(0).Item("DateOpPhraseID"), ""), LanguageID) & " "
            DisplayText &= Copient.PhraseLib.Lookup(Common.NZ(dt.Rows(0).Item("OperatorPhraseID"), ""), LanguageID) & " "
            If Common.NZ(dt.Rows(0).Item("ValueTypeID"), 0) = 1 Then
                DisplayText &= "[" & Copient.PhraseLib.Lookup("term.currentdate", LanguageID).ToLower & "]"
                ValueModifier = Common.NZ(dt.Rows(0).Item("ValueModifier"), "")
                If ValueModifier <> "" AndAlso Integer.TryParse(ValueModifier, Offset) Then
                    ValueModifier = " " & IIf(Offset < 0, " - ", " + ") & Math.Abs(Offset)
                End If
                DisplayText &= ValueModifier
            Else
                DisplayText &= " " & Common.NZ(dt.Rows(0).Item("Value"), "")
            End If
        End If

        Return DisplayText
    End Function

    'Generates the first collision detection screen showing 4 options when deploy/deferdeploy/getapproval/getapprovalwithdeploy/getapprovalwithdeferdeploy is clicked - runinteractive/runbackground/skip/cancel
    Sub GenerateConfirmationBox(ByVal OfferID As Long, ByVal CollisionThreshold As Integer, ByVal IsOAWEnabled As Boolean, ByVal HasOfferBeenModifiedAfterApproval As Boolean, ByVal RequiresOfferApproval As Boolean)
        Send("<div id=""confirming"" style=""display:none;"">")
        Send("  <div id=""confirmingwrap"" style=""width:420px;"">")
        Send("    <div class=""box"" id=""confirmingbox"" style=""height:auto;"">")
        Send("      <h2><span>" & Copient.PhraseLib.Lookup("term.ProductCollisions", LanguageID) & "</span></h2>")
        'Send("      <input type=""button"" class=""ex"" id=""uploaderclose"" name=""uploaderclose"" value=""X"" style=""float:right;position:relative;top:-27px;"" onclick=""javascript:document.getElementById('confirming').style.display='none';"" />")
        If (IsOAWEnabled And HasOfferBeenModifiedAfterApproval And RequiresOfferApproval) Then
            Send("      <p>" & Copient.PhraseLib.Lookup("UEoffer-gen.CollisionConfirm.Approval", LanguageID) & "<p>")
        Else
            Send("      <p>" & Copient.PhraseLib.Lookup("UEoffer-gen.CollisionConfirm", LanguageID) & "<p>")
        End If
        Send("      <form id=""confirmationform"" name=""confirmationform"" action=""#"">")
        Send("        <p style=""text-align:center;padding:1px"">")
        Send("          <input type=""hidden"" name=""OfferID"" value=""" & OfferID & """  />")
        'Send("          <input type=""button"" class=""regular"" value=""" & Copient.PhraseLib.Lookup("CPEoffer-gen.RunDetection", LanguageID) & """ onclick=""javascript:loadCollisions();"" />")
        Send("          <input type=""button"" class=""large"" title=""" & Copient.PhraseLib.Lookup("term.runinteractive", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.runinteractive", LanguageID) & """ onclick=""javascript:loadCollisions();"" />")
        Send("          <span style=""padding:15px"">")
        Send("          </span>")
        Send("          <input type=""submit"" class=""large"" id=""runBackground"" name=""runBackground""  title=""" & Copient.PhraseLib.Lookup("term.runbackground", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.runbackground", LanguageID) & """ onclick=""javascript:processcollisionbackground();"" />")
        Send("        </p>")
        Send("        <p style=""text-align:center;padding:1px"">")
        Send("          <input type=""submit"" class=""large"" id=""confirmingDeferDeploy"" name=""deferdeploy"" title=""" & Copient.PhraseLib.Lookup("CPEoffer-gen.SkipDetection", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("CPEoffer-gen.SkipDetection", LanguageID) & """ />")
        Send("          <input type=""submit"" class=""large"" id=""confirmingDeploy"" name=""deploy"" title=""" & Copient.PhraseLib.Lookup("CPEoffer-gen.SkipDetection", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("CPEoffer-gen.SkipDetection", LanguageID) & """ />")
        Send("          <input type=""submit"" class=""large"" id=""confirmingApproval"" name=""reqApproval"" title=""" & Copient.PhraseLib.Lookup("CPEoffer-gen.SkipDetection", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("CPEoffer-gen.SkipDetection", LanguageID) & """ />")
        Send("          <input type=""submit"" class=""large"" id=""confirmingDeployDeferApproval"" name=""reqApprovalWithDeferDeployment"" title=""" & Copient.PhraseLib.Lookup("CPEoffer-gen.SkipDetection", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("CPEoffer-gen.SkipDetection", LanguageID) & """ />")
        Send("          <input type=""submit"" class=""large"" id=""confirmingDeployApproval"" name=""reqApprovalWithDeployment"" title=""" & Copient.PhraseLib.Lookup("CPEoffer-gen.SkipDetection", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("CPEoffer-gen.SkipDetection", LanguageID) & """ />")
        Send("          <span style=""padding:15px"">")
        Send("          </span>")
        Send("          <input type=""button"" class=""large"" title=""" & Copient.PhraseLib.Lookup("term.cancel", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.cancel", LanguageID) & """ onclick=""javascript:toggleDialog('confirming', false);"" />")
        Send("        </p>")
        Send("      </form>")
        Send("    </div>")
        Send("  </div>")
        Send("</div>")
    End Sub

    '-----------------------------------------------------------------------------------------------------------------------------

    Sub GenerateRunDetectionBox(ByVal OfferID As Long, ByVal CollisionThreshold As Integer, Optional ByVal IsRunCollision As Integer = -1)
        Send("<div id=""confirmingDetection"" style=""display:none;"">")
        Send("  <div id=""confirmingwrapDetection"" style=""width:420px;"">")
        Send("    <div class=""box"" id=""confirmingboxDetection"" style=""height:auto;"">")
        Send("      <h2><span>" & Copient.PhraseLib.Lookup("term.ProductCollisions", LanguageID) & "</span></h2>")
        Send("      <p>" & Copient.PhraseLib.Lookup("UEoffer-gen.CDConfirm", LanguageID) & "<p>")
        Send("      <form id=""confirmationform"" name=""confirmationform"" action=""#"">")
        Send("        <p style=""text-align:left;padding:1px"">")
        Send("          <input type=""hidden"" name=""OfferID"" value=""" & OfferID & """  />")
        Send("          <input type=""hidden"" name=""hdnRunCollision"" id=""hdnRunCollision"" value=""" & IsRunCollision & """  />")
        Send("          <input type=""button"" class=""mediumshort"" title=""" & Copient.PhraseLib.Lookup("term.runinteractive", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.runinteractive", LanguageID) & """ onclick=""javascript:loadCollisionsDetection();"" />")
        Send("          <input type=""submit"" class=""mediumshort"" id=""runBackground"" name=""runBackground""  title=""" & Copient.PhraseLib.Lookup("term.runbackground", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.runbackground", LanguageID) & """ onclick=""javascript:processcollisionbackgroundDetection();"" />")
        Send("          <input type=""button"" class=""mediumshort"" title=""" & Copient.PhraseLib.Lookup("term.cancel", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.cancel", LanguageID) & """ onclick=""javascript:toggleDialog('confirmingDetection', false);"" />")
        Send("        </p>")
        Send("      </form>")
        Send("    </div>")
        Send("  </div>")
        Send("</div>")
        ' div with disabled options
        Send("<div id=""confirmingDisabledDetection"" style=""display:none;"">")
        Send("  <div id=""confirmingwrapDisabledDetection"" style=""width:420px;"">")
        Send("    <div class=""box"" id=""confirmingboxDisabledDetection"" style=""height:auto;"">")
        Send("      <h2><span>" & Copient.PhraseLib.Lookup("term.ProductCollisions", LanguageID) & "</span></h2>")
        Send("      <p>" & Copient.PhraseLib.Lookup("UEoffer-gen.CollisionDisabled", LanguageID) & "<p>")
        Send("      <form id=""confirmationform"" name=""confirmationform"" action=""#"">")
        Send("        <p style=""text-align:left;padding:1px"">")
        Send("          <input type=""hidden"" name=""OfferID"" value=""" & OfferID & """  />")
        Send("          <input type=""button"" class=""mediumshort"" disabled title=""" & Copient.PhraseLib.Lookup("term.runinteractive", LanguageID) & """  value=""" & Copient.PhraseLib.Lookup("term.runinteractive", LanguageID) & """ onclick=""javascript:loadCollisionsDetection();"" />")
        Send("          <input type=""submit"" class=""mediumshort"" disabled id=""runBackground"" name=""runBackground""  title=""" & Copient.PhraseLib.Lookup("term.runbackground", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.runbackground", LanguageID) & """ onclick=""javascript:processcollisionbackgroundDetection();"" />")
        Send("          <input type=""button"" class=""mediumshort"" title=""" & Copient.PhraseLib.Lookup("term.ok", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.ok", LanguageID) & """ onclick=""javascript:toggleDialog('confirmingDisabledDetection', false);"" />")
        Send("        </p>")
        Send("      </form>")
        Send("    </div>")
        Send("  </div>")
        Send("</div>")
    End Sub

    '---------------------------------------------------------------------------------------------------------------------------
    Sub GenerateCollisionsBox(ByVal OfferID As Long, ByVal CollisionThreshold As Integer)
        Send("<div id=""loading"" style=""display:none;"">")
        Send("  <div id=""loadingwrap"" style=""width:420px;"">")
        Send("    <div class=""box"" id=""loadingbox"" style=""height:auto;"" >")
        Send("      <h2><span>" & Copient.PhraseLib.Lookup("term.ProductCollisions", LanguageID) & "</span></h2>")
        'Send("      <input type=""button"" class=""ex"" id=""uploaderclose"" name=""uploaderclose"" value=""X"" style=""float:right;position:relative;top:-27px;"" onclick=""javascript:document.getElementById('loading').style.display='none';"" />")
        Send("      <form id=""collisionform"" name=""collisionform"" action=""#"" style=""height: 85%;"" >")
        Send("      <div id=""collisionsContent"" style=""width:100%; height:100%"">")
        Send("        <p>" & Copient.PhraseLib.Lookup("UEoffer-gen.FindingCollisions", LanguageID) & "</p>")
        Send("        <p style=""text-align:center;padding-top:10px;""><img src=""../../images/loadingAnimation.gif"" height=""80px;"" alt=""" & Copient.PhraseLib.Lookup("term.loading", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.loading", LanguageID) & """ /></p>")
        Send("      </div>")
        Send("      </form>")
        Send("    </div>")
        Send("  </div>")
        Send("</div>")
    End Sub

    '-----------------------------------------------------------------------------------------------------------------------------


    Sub GenerateCollisionsDetectionBox(ByVal OfferID As Long, ByVal CollisionThreshold As Integer)
        Send("<div id=""loadingDetection"" style=""display:none;"">")
        Send("  <div id=""loadingwrapDetection"" style=""width:420px;"">")
        Send("    <div class=""box"" id=""loadingboxDetection"" style=""height:auto;"" >")
        Send("      <h2><span>" & Copient.PhraseLib.Lookup("term.ProductCollisions", LanguageID) & "</span></h2>")
        Send("      <form id=""collisionform"" name=""collisionform"" action=""#"" style=""height: 85%;"" >")
        Send("      <div id=""collisionsContentDetection"" style=""width:100%; height:100%"">")
        Send("        <p>" & Copient.PhraseLib.Lookup("UEoffer-gen.FindingCollisions", LanguageID) & "</p>")
        Send("        <p style=""text-align:center;padding-top:10px;""><img src=""../../images/loadingAnimation.gif"" height=""80px;"" alt=""" & Copient.PhraseLib.Lookup("term.loading", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.loading", LanguageID) & """ /></p>")
        Send("      </div>")
        Send("      </form>")
        Send("    </div>")
        Send("  </div>")
        Send("</div>")
    End Sub

    '-----------------------------------------------------------------------------------------------------------------------------
    Sub GenerateOAWBox(ByVal OfferID As Long)
        Send("<div id=""oawreject"" style=""display: none;"">")
        Send("  <div id=""oawrejectwrap"" style=""width:420px;"">")
        Send("    <div class=""box"" id=""oawrejectbox"" style=""height:auto;"">")
        Send("      <h2><span>" & Copient.PhraseLib.Lookup("term.offerrejection", LanguageID) & "</span></h2>")
        Send("      <p><br>" & Copient.PhraseLib.Lookup("term.rejectionreason", LanguageID) & " :<p>")
        Send("      <form id=""oawform"" name=""oawform"" action=""#"">")
        Send("        <p style=""text-align:center;padding:1px"">")
        Send("          <input type=""hidden"" name=""OfferID"" value=""" & OfferID & """  />")
        Send("          <textarea rows=""4"" id=""rejectText"" name=""rejectText"" class=""boxsizingBorder"" style=""resize: none"" maxlength=""500"" onchange=""return checkLength();"" onkeyup=""return checkLength();"" onkeydown=""return checkLength();""></textarea>")
        Send("          <br Class=""half"" /><br />")
        Send("          <small>")
        Send("          (" & Copient.PhraseLib.Lookup("offerrejection.rejectiontext", LanguageID) & ")")
        Send("          </small>")
        Send("          <br />")
        Send("          <span style=""padding:15px"">")
        Send("          </span>")
        Send("          <br><br>")
        Send("          <input type=""submit"" class=""large"" id=""rejectOffer"" name=""rejectOffer""  title=""" & Copient.PhraseLib.Lookup("term.reject", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.reject", LanguageID) & """/>")
        Send("          <input type=""button"" class=""large"" id=""cancelOAW"" name=""cancelOAW""  title=""" & Copient.PhraseLib.Lookup("term.cancel", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.cancel", LanguageID) & """ onclick=""hideRejectConfirmation();""/>")
        Send("        </p>")
        Send("      </form>")
        Send("    </div>")
        Send("  </div>")
        Send("</div>")
    End Sub

    '-----------------------------------------------------------------------------------------------------------------------------

    Sub GenerateApprovalStatusBar(ByVal Logix As Copient.LogixInc, ByVal OfferID As Long, ByVal isOfferAwaitingApproval As Boolean, ByVal isOfferApproved As Boolean, ByVal disable As Boolean)
        Send("<div style=""margin-left: 20px;"">  ")
        Send("<div  " & IIf(disable, " class=""circle filled-greyed""", "class=""circle filled""") & " style=""float:left;"" ></div>")
        If disable Then
            Send("<div class=""horizontal greyed"" style=""float:left;""></div>")
            Send("<div  " & IIf(isOfferApproved OrElse isOfferAwaitingApproval, " class=""circle filled-greyed""", "class=""circle greyed""") & " style=""float:left;""></div>")
            Send("<div class=""horizontal greyed"" style=""float:left;""></div>")
            Send("<div  " & IIf(isOfferApproved, " class=""circle filled-greyed""", "class=""circle greyed""") & " style=""float:left;""></div>")
        Else
            Send("<div  " & IIf(isOfferApproved OrElse isOfferAwaitingApproval, " class=""horizontal green""", "class=""horizontal""") & " style=""float:left;""></div>")
            Send("<div  " & IIf(isOfferApproved OrElse isOfferAwaitingApproval, " class=""circle filled""", "class=""circle""") & " style=""float:left;""></div>")
            Send("<div  " & IIf(isOfferApproved, " class=""horizontal green""", "class=""horizontal""") & " style=""float:left;""></div>")
            Send("<div  " & IIf(isOfferApproved, " class=""circle filled""", "class=""circle""") & " style=""float:left;""></div>")
        End If
        Send("<div class=""circle""></div>")
        Send("</div>")
    End Sub

    '-----------------------------------------------------------------------------------------------------------------------------
    Function Get_Preference_Value(ByRef Common As Copient.CommonInc, ByVal PreferenceID As Long, ByVal Value As String) As String
        Dim TempLong As Long = 0
        Dim dt As DataTable

        Common.QueryStr = "select DataTypeID from Preferences with (NoLock) where PreferenceID=" & PreferenceID & " and Deleted=0;"
        dt = Common.PMRT_Select
        If dt.Rows.Count > 0 Then
            Select Case Common.NZ(dt.Rows(0).Item("DataTypeID"), 0)
                Case 1 ' list
                    ' lookup to see if this is a preference with list items, if so get the list item name
                    Common.QueryStr = "select case when UPT.PhraseID is null then PLI.Name " & _
                                      "       else CONVERT(nvarchar(200), UPT.Phrase) end as PhraseText " & _
                                      "from Preferences as PREF with (NoLock) " & _
                                      "inner join PreferenceListItems as PLI with (NoLock) on PLI.PreferenceID = PREF.PreferenceID " & _
                                      "left join UserPhraseText as UPT with (NoLock) on UPT.PhraseID = PLI.NamePhraseID " & _
                                      "where PREF.Deleted=0 and PREF.DataTypeID=1 and PREF.PreferenceID=" & PreferenceID & _
                                      "  and PLI.Value=N'" & Value & "';"
                    dt = Common.PMRT_Select
                    If dt.Rows.Count > 0 Then
                        Value = Common.NZ(dt.Rows(0).Item("PhraseText"), Value)
                    End If
                Case 5 ' boolean
                    Value = Copient.PhraseLib.Lookup(IIf(Value = "1", "term.true", "term.false"), LanguageID)
            End Select

        End If

        Return Value
    End Function


    '-----------------------------------------------------------------------------------------------------------------------------

    Function GetAssociatedProductGroupIDs(ByVal ROID As Long, ByRef MyCommon As Copient.CommonInc) As String
        Dim PGList As String = "-1"
        Dim dt As DataTable
        Dim PGIDs(-1) As String
        Dim i As Integer

        MyCommon.QueryStr = "select ProductGroupID from CPE_IncentiveProductGroups with (NoLock) where RewardOptionID = " & ROID & _
                            " union " & _
                            "select DiscountedProductGroupID as ProductGroupID from CPE_Discounts as DISC with (NoLock) " & _
                            "inner join CPE_Deliverables as DEL with (NoLock) on DEL.OutputID = DISC.DiscountID and DEL.DeliverableTypeID=2 " & _
                            "where(DISC.Deleted = 0 And DISC.DiscountedProductGroupID > 0 And DEL.Deleted = 0 And DEL.RewardOptionID = " & ROID & ") " & _
                            " union " & _
                            "select ExcludedProductGroupID as ProductGroupID from CPE_Discounts as DISC with (NoLock) " & _
                            "inner join CPE_Deliverables as DEL with (NoLock) on DEL.OutputID = DISC.DiscountID and DEL.DeliverableTypeID=2 " & _
                            "where(DISC.Deleted = 0 And ExcludedProductGroupID > 0 And DEL.Deleted = 0 And DEL.RewardOptionID = " & ROID & ") " & _
                            "order by ProductGroupID;"
        dt = MyCommon.LRT_Select
        If dt.Rows.Count > 0 Then
            ReDim PGIDs(dt.Rows.Count - 1)
            For i = 0 To PGIDs.GetUpperBound(0)
                PGIDs(i) = MyCommon.NZ(dt.Rows(0).Item("ProductGroupID"), "0")
            Next
            PGList = String.Join(",", PGIDs)
        End If

        Return PGList
    End Function

    Function GetFolderDetails(ByVal OfferID As Long, ByRef MyCommon As Copient.CommonInc) As DataTable

        Dim dt As DataTable = Nothing


        MyCommon.QueryStr = "select FI.FolderID, F.EndDate from FolderItems as FI with (NoLock) " & _
                                        "inner join Folders as F with (NoLock) on F.FolderID = FI.FolderID " & _
                                        "where LinkID=" & OfferID & " and LinkTypeID=1;"
        dt = MyCommon.LRT_Select

        Return dt
    End Function

    '-----------------------------------------------------------------------------------------------------------------------------
    'Updating Promotion Display flag and Prorate on display flag when system options are off.      
    Sub SetPromotionDisplay(ByRef Common As Copient.CommonInc, ByVal OfferID As Long)
        Dim isEnabled As Boolean
        isEnabled = IIf(Common.Fetch_UE_SystemOption(145) = "1", True, False)
        If isEnabled = False Then
            Common.QueryStr = "update CPE_Incentives set PromotionDisplay=0 where IncentiveID=" & OfferID & ";"
            Common.LRT_Execute()
        End If
        'Updating Prorate on Display flag when system option 154 is off.   
        isEnabled = IIf(Common.Fetch_UE_SystemOption(154) = "1", True, False)
        If isEnabled = False Then
            Common.QueryStr = "update CPE_Incentives set ProrateonDisplay=0 where IncentiveID=" & OfferID & ";"
            Common.LRT_Execute()
        End If
    End Sub

    Sub UpdateTemplatePermissions(ByRef Common As Copient.CommonInc, ByVal SourceOfferID As Long, ByVal OfferID As Long, ByVal systemOption As Integer)

        Dim dtTempPermission As New DataTable
        Dim Disallow_DisplayDates As Integer = Integer.MinValue

        If (systemOption = 143) Then
            Common.QueryStr = "SELECT Disallow_DisplayDates from TemplatePermissions with (NoLock) WHERE OfferID = " & SourceOfferID & ";"
            dtTempPermission = Common.LRT_Select()
            If dtTempPermission.Rows.Count > 0 Then
                Disallow_DisplayDates = Convert.ToInt16(dtTempPermission.Rows(0).Item("Disallow_DisplayDates"))
            End If
            Common.QueryStr = "UPDATE TemplatePermissions with (RowLock) Set Disallow_DisplayDates=" & Disallow_DisplayDates & " WHERE OfferID = " & OfferID
            Common.LRT_Execute()
        End If
    End Sub

    Sub SaveOfferDisplayDates(ByRef Common As Copient.CommonInc, ByVal SourceOfferID As Long, ByVal OfferID As Long, ByVal AdminUserID As Integer, ByVal EngineId As Integer)

        Common.QueryStr = "SELECT DisplayStartDate,DisplayEndDate FROM OfferAccessoryFields with (NoLock) WHERE OfferID = " & SourceOfferID & ";"
        Dim dtODisp As New DataTable
        Dim startDate As String = ""
        Dim endDate As String = ""
        dtODisp = Common.LRT_Select()
        If dtODisp.Rows.Count > 0 Then
            startDate = Common.NZ(dtODisp.Rows(0).Item("DisplayStartDate"), Nothing)
            endDate = Common.NZ(dtODisp.Rows(0).Item("DisplayEndDate"), Nothing)
        End If
        Common.QueryStr = "dbo.pa_UpdateOfferAccessoryFields"
        Common.Open_LRTsp()
        Common.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
        Common.LRTsp.Parameters.Add("@DisplayStartTime", SqlDbType.DateTime).Value = IIf(String.IsNullOrEmpty(startDate), DBNull.Value, startDate)
        Common.LRTsp.Parameters.Add("@DisplayEndTime", SqlDbType.DateTime).Value = IIf(String.IsNullOrEmpty(endDate), DBNull.Value, endDate)

        Common.LRTsp.Parameters.Add("@RedemThresholdPerHour", SqlDbType.BigInt).Value = DBNull.Value
        Common.LRTsp.Parameters.Add("@AdminUserId", SqlDbType.BigInt).Value = AdminUserID
        Common.LRTsp.Parameters.Add("@EngineID", SqlDbType.BigInt).Value = EngineId

        Common.LRTsp.Parameters.Add("@OptionID", SqlDbType.BigInt).Value = 85
        Common.LRTsp.ExecuteNonQuery()
        Common.Close_LRTsp()

    End Sub

    Sub UpdateInclusionIncentiveProductGroupSet(ByRef Common As Copient.CommonInc, ByVal OfferID As Long)
        Dim TargetROID As Long = 0
        Common.QueryStr = "select RewardOptionID from CPE_RewardOptions with (NoLock) where IncentiveID  = " & OfferID & ";"
        Dim dt As New DataTable

        dt = Common.LRT_Select()
        If dt.Rows.Count > 0 Then
            TargetROID = Common.NZ(dt.Rows(0).Item("RewardOptionID"), 0)
        End If

        If (TargetROID > 0) Then
            Common.QueryStr = "dbo.pc_Duplicate_CPE_ProductCondition_MulitpleExcludedProductCondition"
            Common.Open_LRTsp()
            Common.LRTsp.Parameters.Add("@TargetROID", SqlDbType.BigInt).Value = TargetROID
            Common.LRTsp.ExecuteNonQuery()
            Common.Close_LRTsp()
        End If
    End Sub
</script>
<%
    If MyCommon.Fetch_SystemOption(75) Then
        If (OfferID > 0 And Logix.UserRoles.AccessNotes) Then
            Send_Notes(3, OfferID, AdminUserID)
        End If
    End If
    GenerateConfirmationBox(OfferID, 1, m_isOAWEnabled, m_hasOfferModifiedAfterApproval, m_requiresOfferApproval)
    GenerateCollisionsBox(OfferID, 1)
    GenerateRunDetectionBox(OfferID, 1, RunCollision)
    GenerateCollisionsDetectionBox(OfferID, 1)
    GenerateOAWBox(OfferID)
    Send_BodyEnd()
done:
    MyCommon.Close_LogixRT()
    Logix = Nothing
    MyCommon = Nothing
%>