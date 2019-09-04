<%@ Page Language="vb" Debug="true" CodeFile="ue-cb.vb" Inherits="UECB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="CMS.AMS.Models" %>
<%@ Import Namespace="CMS.AMS.Contract" %>
<%@ Import Namespace="CMS.AMS" %>
<%
    ' *****************************************************************************
    ' * FILENAME: UEoffer-con-customer.aspx 
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
    Dim MLI As New Copient.Localization.MultiLanguageRec
    Dim MyCommon As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    CMS.AMS.CurrentRequest.Resolver.AppName = "UEoffer-con-customer.aspx"
    Dim rst As DataTable
    Dim rstBannerCgs As DataTable = Nothing
    Dim dst As DataTable
    Dim row As DataRow
    Dim OfferID As Long
    Dim Name As String = ""
    Dim ConditionID As String
    Dim IsTemplate As Boolean = False
    Dim FromTemplate As Boolean = False
    Dim Disallow_Edit As Boolean = True
    Dim Household As Boolean = False
    Dim MetWhenOffline As Boolean = False
    Dim DisabledAttribute As String = ""
    Dim i As Integer
    Dim roid As Long
    Dim historyString As String
    Dim CloseAfterSave As Boolean = False
    Dim Ids() As String
    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False
    Dim RequireCG As Boolean = False
    Dim HasRequiredCG As Boolean = False
    Dim EngineID As Integer = 2
    Dim BannersEnabled As Boolean = True
    Dim MultiBanneredOffer As Boolean = False
    Dim NewCardholdersID As Integer = 0
    Dim FullListSelect As New StringBuilder()
    Dim AllCAM As Integer = 0
    Dim EligibleIncludedcustomergroups As String = String.Empty
    Dim EligibleExcludedcustomergroups As String = String.Empty
    Dim IsEligibilityConditionExistForOffer As String = "False"
    Dim IsAnyCustomer As Boolean = False
    Dim selectedCards As DataTable = Nothing
    Dim cardIDs As String = ""
    Dim disableCustApproval As Boolean = False
    Dim analyticsCGService As CMS.AMS.Contract.IAnalyticsCustomerGroups = CMS.AMS.CurrentRequest.Resolver.Resolve(Of CMS.AMS.Contract.IAnalyticsCustomerGroups)()
    Dim showAnalyticsCG As Boolean = False
    Dim restrictConditionForRPOS As Boolean = False
    Dim SystemCacheData As ICacheData = CurrentRequest.Resolver.Resolve(Of ICacheData)()
    Dim AnalyticsCG As CustomerGroup
    Dim custApproval As CustomerApproval = New CustomerApproval()
    Dim Localization As Copient.Localization
    Dim gcResult As AMSResult(Of GiftCard) = New AMSResult(Of GiftCard)()
    Dim m_CustomerConditionService As CMS.AMS.Contract.ICustomerGroupCondition = CMS.AMS.CurrentRequest.Resolver.Resolve(Of CMS.AMS.Contract.ICustomerGroupCondition)()
    Dim m_GCRewardService As CMS.AMS.Contract.IGiftCardRewardService = CMS.AMS.CurrentRequest.Resolver.Resolve(Of CMS.AMS.Contract.IGiftCardRewardService)()
    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If

    Response.Expires = 0
    MyCommon.AppName = "UEoffer-con-customer.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
    Localization = New Copient.Localization(MyCommon)
    OfferID = Request.QueryString("OfferID")

    'This would redirect to summary page if offer is awaiting deployment or if offer is awaiting recommendations.
    CheckIfValidOffer(MyCommon, OfferID)

    showAnalyticsCG = (MyCommon.Fetch_UE_SystemOption(208) = "1" And Not analyticsCGService.CheckIfOfferHasAnalyticsCustomerGroup(OfferID))

    AnalyticsCG = analyticsCGService.GetDefaultAnalyticsCustomerGroupForOffer(OfferID)
    ConditionID = Request.QueryString("ConditionID")
    If (Request.QueryString("EngineID") <> "") Then
        EngineID = MyCommon.Extract_Val(Request.QueryString("EngineID"))
    Else
        MyCommon.QueryStr = "select EngineID from OfferIDs where OfferID=" & OfferID & ";"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
            EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), 0)
        End If
    End If

    If (String.IsNullOrEmpty(ConditionID)) Then
        MyCommon.QueryStr = "select IncentiveCustomerID from CPE_IncentiveCustomerGroups where Deleted=0 and ExcludedUsers=0 and RewardOptionID in(select RewardOptionID from  CPE_RewardOptions where IncentiveID =" & OfferID & ");"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
            ConditionID = MyCommon.NZ(rst.Rows(0).Item("IncentiveCustomerID"), 0)
        End If
    End If

    Dim isTranslatedOffer As Boolean = MyCommon.IsTranslatedUEOffer(OfferID, MyCommon)
    Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249) = "1", True, False)

    Dim bEnableAdditionalLockoutRestrictionsOnOffers As Boolean = IIf(MyCommon.Fetch_SystemOption(260) = "1", True, False)
    Dim bOfferEditable As Boolean = MyCommon.IsOfferEditablePastLockOutPeriod(Logix.UserRoles.EditOfferPastLockoutPeriod, MyCommon, OfferID)
    restrictConditionForRPOS = (SystemCacheData.GetSystemOption_UE_ByOptionId(234) = "1")
    BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
    ' get all the any banner cardholder customer groups for this offer 
    'If (BannersEnabled) Then
    '  MyCommon.QueryStr = "select CustomerGroupID, Name from CustomerGroups with (NoLock) " & _
    '                      "where BannerID in (select BannerID from BannerOffers with (NoLock) where OfferID=" & OfferID & ") and deleted =0;"
    '  rstBannerCgs = MyCommon.LRT_Select
    'End If

    MyCommon.QueryStr = "select RewardOptionID from CPE_RewardOptions with (NoLock) where IncentiveID=" & OfferID & " and touchresponse=0 and deleted=0;"
    rst = MyCommon.LRT_Select

    If rst.Rows.Count > 0 Then
        roid = rst.Rows(0).Item("RewardOptionID")
    End If

    'get customer approval record
    Dim custAppResult As AMSResult(Of CustomerApproval) = New AMSResult(Of CustomerApproval)()
    custAppResult = m_CustomerConditionService.GetCustomerApprovalByROID(roid)
    If custAppResult.ResultType = AMSResultType.Success AndAlso custAppResult.Result IsNot Nothing Then
        custApproval = custAppResult.Result
    End If

    'Disable Customer Approval if Gift Card reward is present for this offer
    Dim gcRewardID As Integer = 0
    MyCommon.QueryStr = "select DeliverableID from CPE_Deliverables with (NoLock) where RewardOptionID=" & roid & " and DeliverableTypeID=13 and deleted=0;"
    rst = MyCommon.LRT_Select

    If rst IsNot Nothing AndAlso rst.Rows.Count > 0 Then
        gcRewardID = rst.Rows(0).Item("DeliverableID")
    End If
    gcResult = m_GCRewardService.GetGiftCardReward(gcRewardID, EngineID)
    If custApproval.CustomerApprovalID = 0 Then
        If gcResult.Result IsNot Nothing AndAlso gcResult.Result.Id > 0 Then disableCustApproval = True
    End If

    Dim m_offer1 As CMS.AMS.Contract.IOffer = CMS.AMS.CurrentRequest.Resolver.Resolve(Of CMS.AMS.Contract.IOffer)()
    Dim objOffer As CMS.AMS.Models.Offer = m_offer1.GetOffer(OfferID, CMS.AMS.Models.LoadOfferOptions.EligibilityCustomerCondition)
    If (objOffer Is Nothing = False) Then
        IsEligibilityConditionExistForOffer = IIf(objOffer.IsOptable, "TRUE", "FALSE")
        If objOffer.IsOptable Then
            For Each item In (From p In objOffer.EligibleCustomerGroupConditions.IncludeCondition
                              Where p.Deleted = False
                              Select p.CustomerGroup).ToList()
                If (EligibleIncludedcustomergroups <> "") Then
                    EligibleIncludedcustomergroups = EligibleIncludedcustomergroups & ","
                End If
                EligibleIncludedcustomergroups = EligibleIncludedcustomergroups & item.CustomerGroupID
            Next
            For Each item In (From p In objOffer.EligibleCustomerGroupConditions.ExcludeCondition
                              Where p.Deleted = False
                              Select p.CustomerGroup).ToList()
                If (EligibleExcludedcustomergroups <> "") Then
                    EligibleExcludedcustomergroups = EligibleExcludedcustomergroups & ","
                End If
                EligibleExcludedcustomergroups = EligibleExcludedcustomergroups & item.CustomerGroupID
            Next
        End If
    End If

    ' Check to see if someone is saving the condition
    If (Request.QueryString("save") <> "") Then
        ' Find out the roid
        If roid > 0 Then
            If (Request.QueryString("selGroups") = "") AndAlso (Request.QueryString("require_cg") = "") Then
                infoMessage = Copient.PhraseLib.Lookup("term.customergroupselect", LanguageID)
            Else
                ' Need to save the hhenable choice
                Dim OptOutAllowed As Boolean = True
                Dim form_Household As Integer = IIf(Request.QueryString("household") = "on", 1, 0)
                Dim form_Offline As Integer = IIf(Request.QueryString("offline") = "on", 1, 0)

                MyCommon.QueryStr = "update CPE_RewardOptions with (RowLock) set HHEnable=" & form_Household & ", OfflineCustCondition=" & form_Offline &
                                    " where TouchResponse=0 and IncentiveID=" & OfferID
                MyCommon.LRT_Execute()

                ' Check to see if a customer condition is required by the template, if applicable
                MyCommon.QueryStr = "select CustomerGroupID from CPE_IncentiveCustomerGroups with (NoLock) where RewardOptionID=" & roid &
                                    " and RequiredFromTemplate=1 and deleted=0 and ExcludedUsers=0;"
                rst = MyCommon.LRT_Select
                HasRequiredCG = (rst.Rows.Count > 0)

                'Analytics related changes. Check for selGroups values and if > 0 and see if analyticsCg has been deselected.
                'If yes, then call the library to delete the analytics cg for this offer.
                If (Request.QueryString("selGroups") <> "") Then
                    Ids = Request.QueryString("selGroups").Split(",")

                    'Delete analytics cg if it is deselected
                    If AnalyticsCG.CustomerGroupID <> -1 And Not Ids.ToList().Contains(AnalyticsCG.CustomerGroupID) Then
                        analyticsCGService.DeleteDefaultAnalyticsCustomerGroupForOffer(OfferID)
                    End If

                    'if analytics cg exists with -1 exists in the list, create a new analytics cg and replace it
                    If Ids.ToList().Contains("-1") Then
                        AnalyticsCG.CustomerGroupID = analyticsCGService.CreateAnalyticsCustomerGroup(AnalyticsCG.Name, OfferID)
                        ' log history for this group
                        MyCommon.Activity_Log(4, AnalyticsCG.CustomerGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.cgroup-create", LanguageID))
                    End If
                End If

                ' We got some selected groups so let's blow out all the existing ones
                MyCommon.QueryStr = "update CPE_IncentiveCustomerGroups with (RowLock) set deleted=1, TCRMAStatusFlag=3 where RewardOptionID=" & roid &
" and deleted=0 and ExcludedUsers=0"
                MyCommon.LRT_Execute()

                ' Let's handle the selected groups first
                If (Request.QueryString("selGroups") <> "") Then
                    historyString = Copient.PhraseLib.Lookup("history.con-customer-edit", LanguageID) & ": "

                    For i = 0 To Ids.Length - 1
                        Dim id As Long = MyCommon.Extract_Val(Ids(i))
                        'replace the analytics cg id with the id that we created in previous step.
                        If id = -1 Then
                            id = AnalyticsCG.CustomerGroupID
                        End If
                        historyString &= id & ","
                        MyCommon.QueryStr = "insert into CPE_IncentiveCustomerGroups with (RowLock) (RewardOptionID,CustomerGroupID,ExcludedUsers,Deleted,LastUpdate,RequiredFromTemplate,TCRMAStatusFlag) " &
"values(" & roid & "," & id & ",0,0,getdate()," & IIf(HasRequiredCG, "1", "0") & ",3)"
                        MyCommon.LRT_Execute()
                    Next
                ElseIf HasRequiredCG Then
                    MyCommon.QueryStr = "insert into CPE_IncentiveCustomerGroups with (RowLock) (RewardOptionID,ExcludedUsers,Deleted,LastUpdate,RequiredFromTemplate,TCRMAStatusFlag) " &
                                        "values(" & roid & ",0,0,getdate(),1,3)"
                    MyCommon.LRT_Execute()
                End If

                ' Check to see if a customer condition is required by the template, if applicable
                MyCommon.QueryStr = "select CustomerGroupID from CPE_IncentiveCustomerGroups with (NoLock) where RewardOptionID=" & roid &
                                    " and RequiredFromTemplate=1 and deleted=0 and ExcludedUsers=1;"
                rst = MyCommon.LRT_Select
                HasRequiredCG = (rst.Rows.Count > 0)

                ' We got some selected groups so let's blow out all the existing ones
                MyCommon.QueryStr = "update CPE_IncentiveCustomerGroups with (RowLock) set deleted=1, TCRMAStatusFlag=3 where RewardOptionID=" & roid &
                                    " and deleted=0 and ExcludedUsers=1"
                MyCommon.LRT_Execute()

                ' Now let's handle the excluded groups
                If (Request.QueryString("exGroups") <> "") Then
                    historyString = historyString & " " & Copient.PhraseLib.Lookup("term.excluding", LanguageID) & ": " & Request.QueryString("exGroups")
                    Ids = Request.QueryString("exGroups").Split(",")
                    For i = 0 To Ids.Length - 1
                        MyCommon.QueryStr = "insert into CPE_IncentiveCustomerGroups with (RowLock) (RewardOptionID,CustomerGroupID,ExcludedUsers,Deleted,LastUpdate,RequiredFromTemplate, TCRMAStatusFlag) " &
                                            "values(" & roid & "," & MyCommon.Extract_Val(Ids(i)) & ",1,0,getdate()," & IIf(HasRequiredCG, "1", "0") & ",3)"
                        'Send(MyCommon.QueryStr)
                        MyCommon.LRT_Execute()
                    Next
                ElseIf HasRequiredCG Then
                    MyCommon.QueryStr = "insert into CPE_IncentiveCustomerGroups with (RowLock) (RewardOptionID,ExcludedUsers,Deleted,LastUpdate,RequiredFromTemplate, TCRMAStatusFlag) " &
                                        "values(" & roid & ",1,0,getdate(),1,3)"
                    MyCommon.LRT_Execute()
                End If
                ' Now let's handle the excluded eligilble groups
                If (objOffer Is Nothing = False And objOffer.IsOptable And Request.QueryString("mode") <> "optout") Then
                    If (Request.QueryString("tomoveeligibleincludedgroups") <> "") Then
                        MyCommon.QueryStr = "UPDATE [CustomerConditionDetails] SET [Excluded]=1 WHERE [ConditionID]= " & objOffer.EligibleCustomerGroupConditions.ConditionID & " AND [CustomerGroupID] IN (" & Request.QueryString("tomoveeligibleincludedgroups") & ")"
                        MyCommon.LRT_Execute()
                    End If
                    If (Request.QueryString("toaddeligibleexcludedgroups") <> "") Then
                        Ids = Request.QueryString("toaddeligibleexcludedgroups").Split(",")
                        For i = 0 To Ids.Length - 1
                            MyCommon.QueryStr = "INSERT INTO [CustomerConditionDetails] ([ConditionID],[CustomerGroupID],[Excluded]) VALUES(" & objOffer.EligibleCustomerGroupConditions.ConditionID & "," & MyCommon.Extract_Val(Ids(i)) & ",1)"
                            'Send(MyCommon.QueryStr)
                            MyCommon.LRT_Execute()
                        Next
                    End If
                End If

                ' Finally, set the "AllowOptOut" bit on the offer to 0 (False) if it's Any Customer or Any Cardholder with no exclusions
                If (Request.QueryString("exGroups") = "") Then
                    If (Request.QueryString("selGroups") = "1") Or (Request.QueryString("selGroups") = "2") Then
                        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set AllowOptOut=0 where IncentiveID=" & OfferID
                        MyCommon.LRT_Execute()
                    End If
                End If

                MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID
                MyCommon.LRT_Execute()
                ResetOfferApprovalStatus(OfferID)
                If historyString IsNot Nothing Then
                    MyCommon.Activity_Log(3, OfferID, AdminUserID, historyString.Substring(0, historyString.Length - 1))
                End If


                Dim selectedCardTypeIDs As String = IIf(Request.QueryString("radioCardTypes") = "1", Request.QueryString("cardTypeSelected"), "")
                m_CustomerConditionService.SaveCustomerConditionCardTypes(selectedCardTypeIDs, roid)

                'Save Customer Approval record
                Dim resultSuccessful As AMSResult(Of Boolean) = New AMSResult(Of Boolean)
                Dim approvalRequired As Boolean = MyCommon.NZ(Request.QueryString("approvalRadio"), 0)

                If custApproval IsNot Nothing Then
                    custApproval.RewardOptionID = roid
                    custApproval.MessageDescription = MyCommon.NZ(Request.QueryString("mesgDescText"), "")
                    custApproval.ApprovalType = MyCommon.NZ(Request.QueryString("approvalLimit"), 0)
                    If approvalRequired Then
                        If custApproval.MessageDescription.Trim = "" Then
                            infoMessage = Copient.PhraseLib.Lookup("error.mesgdescription", LanguageID)
                        Else
                            If custApproval.CustomerApprovalID = 0 Then
                                'save
                                resultSuccessful = m_CustomerConditionService.SaveCustomerApprovalCondition(custApproval)
                            Else
                                'update
                                resultSuccessful = m_CustomerConditionService.UpdateCustomerApprovalCondition(custApproval)
                            End If
                            If Not resultSuccessful.Result Then infoMessage = Copient.PhraseLib.Lookup("error.customerapproval-update", LanguageID)
                        End If
                    Else
                        If custApproval.CustomerApprovalID > 0 Then
                            'delete
                            Dim deletedSuccessfully As Boolean = m_CustomerConditionService.DeleteCustomerApprovalRecord(roid)
                            If Not deletedSuccessfully Then
                                infoMessage = Copient.PhraseLib.Lookup("error.customerapproval-delete", LanguageID)
                            End If
                        End If
                    End If

                    'save customer approval translation record
                    If resultSuccessful.Result Then
                        MLI.ItemID = custApproval.CustomerApprovalID
                        MLI.MLTableName = "CPE_IncentiveCustomerApprovalTranslation"
                        MLI.MLColumnName = "MessageDescription"
                        MLI.MLIdentifierName = "CustomerApprovalID"
                        MLI.StandardTableName = "CPE_IncentiveCustomerApproval"
                        MLI.StandardColumnName = "MessageDescription"
                        MLI.StandardIdentifierName = "CustomerApprovalID"
                        MLI.InputName = "mesgDescText"
                        Localization.SaveTranslationInputs(MyCommon, MLI, Request.QueryString, 9)
                    End If
                End If

            End If
        Else
            infoMessage = Copient.PhraseLib.Lookup("ueoffer-con-customer.ErrorNoROID", LanguageID)
        End If

        If (infoMessage = "") Then
            CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
        Else
            CloseAfterSave = False
        End If
        '!
        'If CloseAfterSave = False Then
        '  Response.Redirect("UEoffer-con-customer.aspx?OfferID=" & OfferID)
        'End If

    End If

    ' Dig the offer info out of the database
    ' No one clicked anything
    MyCommon.QueryStr = "Select IncentiveID,IsTemplate,ClientOfferID,IncentiveName as Name,CPE.Description,PromoClassID,CRMEngineID,Priority," &
                        "StartDate,EndDate,EveryDOW,EligibilityStartDate,EligibilityEndDate,TestingStartDate,TestingEndDate,P1DistQtyLimit,P1DistTimeType,P1DistPeriod," &
                        "P2DistQtyLimit,P2DistTimeType,P2DistPeriod,P3DistQtyLimit,P3DistTimeType,P3DistPeriod,EnableImpressRpt,EnableRedeemRpt,CreatedDate,CPE.LastUpdate,CPE.Deleted, CPEOARptDate, CPEOADeploySuccessDate, CPEOADeployRpt," &
                        "CRMRestricted,StatusFlag,OC.Description as CategoryName,IsTemplate,FromTemplate from CPE_Incentives as CPE with (NoLock) left join OfferCategories as OC with (NoLock) on CPE.PromoClassID=OfferCategoryID where IncentiveID=" & Request.QueryString("OfferID")
    rst = MyCommon.LRT_Select
    For Each row In rst.Rows
        Name = MyCommon.NZ(row.Item("Name"), "")
        IsTemplate = MyCommon.NZ(row.Item("IsTemplate"), False)
        FromTemplate = MyCommon.NZ(row.Item("FromTemplate"), False)
    Next

    If IsTemplate Then
        Dim tempCG As CMS.AMS.Models.CustomerGroup
        tempCG = analyticsCGService.GetGenericAnalyticsCustomerGroup()
        AnalyticsCG.CustomerGroupID = tempCG.CustomerGroupID
        AnalyticsCG.Name = Copient.PhraseLib.Lookup("term.getsegmentfromanalytics", LanguageID)
    End If

    ' Update the templates permission if necessary
    If (Request.QueryString("save") <> "" AndAlso Request.QueryString("IsTemplate") = "IsTemplate" AndAlso infoMessage = "") Then
        ' Update the status bits for the templates
        Dim form_Disallow_Edit As Integer = 0
        Dim form_Require_CG As Integer = 0

        If (Request.QueryString("Disallow_Edit") = "on") Then
            form_Disallow_Edit = 1
        End If

        If (Request.QueryString("require_cg") <> "") Then
            form_Require_CG = 1
        End If

        ' Both requiring and locking the customer group is not permitted 
        If (form_Disallow_Edit = 1 AndAlso form_Require_CG = 1) Then
            infoMessage = Copient.PhraseLib.Lookup("offer-con.lockeddenied", LanguageID)
            MyCommon.QueryStr = "update CPE_IncentiveCustomerGroups with (RowLock) set DisallowEdit=1, RequiredFromTemplate=0 " &
                                " where RewardOptionID=" & roid & " and deleted = 0;"
        Else
            MyCommon.QueryStr = "update CPE_IncentiveCustomerGroups with (RowLock) set DisallowEdit=" & form_Disallow_Edit &
                                ", RequiredFromTemplate=" & form_Require_CG & " " &
                                " where RewardOptionID=" & roid & " and deleted = 0;"
        End If
        MyCommon.LRT_Execute()

        ' If necessary, create an empty customer condition
        If (form_Require_CG = 1) Then
            MyCommon.QueryStr = "select CustomerGroupID from CPE_IncentiveCustomerGroups with (NoLock) where RewardOptionID=" & roid & " and deleted = 0;"
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count = 0) Then
                MyCommon.QueryStr = "insert into CPE_IncentiveCustomerGroups with (RowLock) (RewardOptionID,ExcludedUsers,Deleted,LastUpdate, RequiredFromTemplate, TCRMAStatusFlag) " &
                    " values(" & roid & ",0,0,getdate(),1,3)"
                MyCommon.LRT_Execute()
            End If
        End If

        If (infoMessage = "") Then
            CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
        Else
            CloseAfterSave = False
        End If
    End If

    If (IsTemplate Or FromTemplate) Then
        ' Dig the permissions if it's a template
        MyCommon.QueryStr = "select DisallowEdit, RequiredFromTemplate from CPE_IncentiveCustomerGroups with (NoLock) where RewardOptionID=" & roid & " and deleted = 0;"
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
            Disallow_Edit = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit"), False)
            RequireCG = MyCommon.NZ(rst.Rows(0).Item("RequiredFromTemplate"), False)
        Else
            Disallow_Edit = False
        End If
    End If

    MyCommon.QueryStr = "select HHEnable, OfflineCustCondition from CPE_RewardOptions with (NoLock) where IncentiveID=" & OfferID & " and touchresponse=0 and deleted=0;"
    rst = MyCommon.LRT_Select

    If rst.Rows.Count > 0 Then
        Household = MyCommon.NZ(rst.Rows(0).Item("HHEnable"), False)
        MetWhenOffline = (MyCommon.NZ(rst.Rows(0).Item("OfflineCustCondition"), 0) = 1)
    End If

    Dim m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
    If Not IsTemplate Then
        DisabledAttribute = IIf(Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And Disallow_Edit), "", " disabled=""disabled""")
    Else
        DisabledAttribute = IIf(Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer, "", " disabled=""disabled""")
    End If

    Send_HeadBegin("term.offer", "term.customercondition", OfferID)
    If (Request.QueryString("mode") = "optout") Then
        Send("<base target='_self'/>")
    End If


    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts()
%>


<%
    Send("<link rel=""stylesheet"" href=""/javascript/jquery-ui-1.10.3/style/jquery-ui.min.css"" />")
    Send(" <script type=""text/javascript"" src=""/javascript/jquery.min.js""></script>")
    Send(" <script type=""text/javascript"" src=""/javascript/jquery-ui-1.10.3/jquery-ui-min.js""></script>")
    Send("<script type=""text/javascript"">")



    Send("// This is the javascript array holding the function list")
    Send("// The PrintJavascriptArray ASP function can be used to print this array.")


    FullListSelect.Append("<select class=""longer"" id=""functionselect"" name=""functionselect"" multiple=""multiple"" size=""12"">")
    MyCommon.QueryStr = "select CustomerGroupID from CustomerGroups with (NoLock) where NewCardholders=1 and Deleted=0;"
    rst = MyCommon.LRT_Select

    If rst.Rows.Count > 0 Then
        NewCardholdersID = MyCommon.NZ(rst.Rows(0).Item("CustomerGroupID"), -1)
    Else
        NewCardholdersID = -1
    End If

    MyCommon.QueryStr = "select CustomerGroupID from CustomerGroups with (NoLock) where AnyCAMCardholder=1 and Deleted=0;"
    rst = MyCommon.LRT_Select

    If rst.Rows.Count > 0 Then
        AllCAM = MyCommon.NZ(rst.Rows(0).Item("CustomerGroupID"), -1)
    Else
        AllCAM = -1
    End If

    Dim blnAllowAnyCustomer As Boolean = False
    'Populate the Javascript array that holds the list of selectable customer groups
    'MyCommon.QueryStr = "Select CustomerGroupID,Name from CustomerGroups with (NoLock) where Deleted=0 and AnyCustomer<>1 and CustomerGroupID<>2 and CustomerGroupID is not null and BannerID is null and NewCardholders=0 and CAMCustomerGroup<>1 "
    MyCommon.QueryStr = "SELECT DISTINCT CG.CustomerGroupID, CG.Name " &
                        "FROM CustomerGroups CG With (NOLOCK) " &
                        "LEFT JOIN ExtSegmentMap ER on CG.CustomerGroupID = ER.InternalId " &
                        "WHERE CG.Deleted = 0 And AnyCustomer <> 1 And CustomerGroupID <> 2 " &
                        "AND CustomerGroupID IS NOT NULL AND BannerID IS NULL " &
                        "And NewCardholders = 0 And CAMCustomerGroup <> 1 AND CG.Deleted = 0 " &
                        "AND ((ER.SegmentTypeID IS NULL OR ER.SegmentTypeID = 1 OR ER.SegmentTypeID = 3) AND " &
                        "(ER.IncentiveId = " & OfferID & " OR ER.ExtSegmentID > 0 OR ER.ExtSegmentID IS NULL)) " &
                        " And CG.isOptInGroup = 0 ORDER BY CG.Name"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
        Sendb("var functionlist = Array(")
        If (EngineID <> 6) Then
            'add "special" customer groups
            'see if the offer conditions/rewards allow us to display AnyCustomer group.  (This is not allowed if conditions/rewards require a known customer ex: Points, Stored Value, etc.)
            MyCommon.QueryStr = "dbo.pa_Check_AnyCustomer_Violation"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@IncentiveID", SqlDbType.BigInt).Value = OfferID
            dst = MyCommon.LRTsp_select
            MyCommon.Close_LRTsp()
            If dst.Rows.Count = 0 Then
                blnAllowAnyCustomer = True
                Sendb("""" & StrConv(Copient.PhraseLib.Lookup("term.anycustomer", LanguageID), VbStrConv.ProperCase) & """,")
                FullListSelect.Append("<option value=""1"" style=""color:brown;font-weight:bold;"">" & StrConv(Copient.PhraseLib.Lookup("term.anycustomer", LanguageID), VbStrConv.ProperCase) & "<\/option>")
            End If
            dst = Nothing
            Sendb("""" & Copient.PhraseLib.Lookup("term.anycardholder", LanguageID) & """,")
            Sendb("""" & Copient.PhraseLib.Lookup("term.newcardholders", LanguageID) & """,")

            FullListSelect.Append("<option value=""2"" style=""color:brown;font-weight:bold;"">" & Copient.PhraseLib.Lookup("term.anycardholder", LanguageID) & "<\/option>")
            FullListSelect.Append("<option value=""" & NewCardholdersID & """ style=""color:brown;font-weight:bold;"">" & Copient.PhraseLib.Lookup("term.newcardholders", LanguageID) & "<\/option>")

            If (showAnalyticsCG) Then
                Sendb("""" & AnalyticsCG.Name & """,")
                FullListSelect.Append("<option value=""" & AnalyticsCG.CustomerGroupID & """  style=""color:brown;font-weight:bold;"">" & AnalyticsCG.Name & "<\/option>")
            End If
        Else
            Sendb("""" & Copient.PhraseLib.Lookup("term.allcam", LanguageID) & """,")
            FullListSelect.Append("<option value=""" & AllCAM & """ style=""color:brown;font-weight:bold;"">" & Copient.PhraseLib.Lookup("term.allcam", LanguageID) & "<\/option>")
        End If

        For Each row In rst.Rows
            Sendb("""" & MyCommon.NZ(row.Item("Name"), "").ToString().Replace("""", "\""") & """,")
        Next
        Send(""""");")

        Sendb("var vallist = Array(")
        If (EngineID <> 6) Then
            If blnAllowAnyCustomer Then
                Sendb("""1"",")
            End If
            Sendb("""2"",")
            Sendb("""" & NewCardholdersID & """,")
        End If
        For Each row In rst.Rows
            Sendb("""" & MyCommon.NZ(row.Item("CustomerGroupID"), 0) & """,")
            If AnalyticsCG.CustomerGroupID = MyCommon.NZ(row.Item("CustomerGroupID"), 0) Then
                FullListSelect.Append("<option value=""" & MyCommon.NZ(row.Item("CustomerGroupID"), 0) & """  style=""color:brown;font-weight:bold;"">" & MyCommon.NZ(row.Item("Name"), "").ToString().Replace("'", "\'") & "<\/option>")
            Else
                FullListSelect.Append("<option value=""" & MyCommon.NZ(row.Item("CustomerGroupID"), 0) & """>" & MyCommon.NZ(row.Item("Name"), "").ToString().Replace("'", "\'") & "<\/option>")
            End If
        Next
        Send(""""");")

        Sendb("var exceptlist = new Array(")
        If (EngineID <> 6) Then
            MyCommon.QueryStr = "select CustomerGroupID from CustomerGroups with (NoLock) where (AnyCardholder=1 or AnyCustomer=1) and Deleted=0 order by CustomerGroupID;"
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count > 0) Then
                i = 1
                For Each row In rst.Rows
                    If (i > 1) Then Sendb(",")
                    Sendb(MyCommon.NZ(row.Item("CustomerGroupID"), 0))
                    i += 1
                Next
            End If
            Send(");")
        Else
            MyCommon.QueryStr = "select CustomerGroupID from CustomerGroups with (NoLock) where AnyCAMCardholder=1 and Deleted=0 order by CustomerGroupID;"
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count > 0) Then
                Sendb("-1,")
                For Each row In rst.Rows
                    Sendb(MyCommon.NZ(row.Item("CustomerGroupID"), 0))
                Next
            Else
                Sendb("-1,-2")
            End If
            Send(");")
        End If
    Else
        Sendb("var functionlist = Array(")
        If (EngineID <> 6) Then
            Sendb("""" & StrConv(Copient.PhraseLib.Lookup("term.anycustomer", LanguageID), VbStrConv.ProperCase) & """,")
            Sendb("""" & Copient.PhraseLib.Lookup("term.anycardholder", LanguageID) & """,")
            Sendb("""" & Copient.PhraseLib.Lookup("term.newcardholders", LanguageID) & """,")
            FullListSelect.Append("<option value=""1"" style=""color:brown;font-weight:bold;"">" & StrConv(Copient.PhraseLib.Lookup("term.anycustomer", LanguageID), VbStrConv.ProperCase) & "<\/option>")
            FullListSelect.Append("<option value=""2"" style=""color:brown;font-weight:bold;"">" & Copient.PhraseLib.Lookup("term.anycardholder", LanguageID) & "<\/option>")
            FullListSelect.Append("<option value=""" & NewCardholdersID & """ style=""color:brown;font-weight:bold;"">" & Copient.PhraseLib.Lookup("term.newcardholders", LanguageID) & "<\/option>")

            If (showAnalyticsCG) Then
                Send("""" & AnalyticsCG.Name & """);")
                FullListSelect.Append("<option value=""" & AnalyticsCG.CustomerGroupID & """ style=""color:brown;font-weight:bold;"">" & AnalyticsCG.Name & "<\/option>")
            Else
                Send(""""");")
            End If

            Sendb("var vallist = Array(")
            Sendb("""" & "1" & """,")
            Sendb("""" & "2" & """,")
            Sendb("""" & NewCardholdersID & """,")
            If (showAnalyticsCG) Then
                Send("""" & AnalyticsCG.CustomerGroupID & """);")
            Else
                Send(""""");")
            End If
        Else
            Send("""" & Copient.PhraseLib.Lookup("term.allcam", LanguageID) & """);")
            FullListSelect.Append("<option value=""" & AllCAM & """ style=""color:brown;font-weight:bold;"">" & Copient.PhraseLib.Lookup("term.allcam", LanguageID) & "<\/option>")
            Sendb("var vallist = Array(")
            Send("""" & AllCAM & """);")
        End If


        Sendb("var exceptlist = new Array(")
        If (EngineID <> 6) Then
            MyCommon.QueryStr = "select CustomerGroupID from CustomerGroups with (NoLock) where (AnyCardholder=1 or AnyCustomer=1) and Deleted=0 order by CustomerGroupID;"
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count > 0) Then
                i = 1
                For Each row In rst.Rows
                    If (i > 1) Then Sendb(",")
                    Sendb(MyCommon.NZ(row.Item("CustomerGroupID"), 0))
                    i += 1
                Next
                Send(");")
            Else
                Send("""" & "-99" & """);")
            End If
        Else
            MyCommon.QueryStr = "select CustomerGroupID from CustomerGroups with (NoLock) where AnyCAMCardholder=1 and Deleted=0 order by CustomerGroupID;"
            If (rst.Rows.Count > 0) Then
                rst = MyCommon.LRT_Select
                For Each row In rst.Rows
                    Sendb(MyCommon.NZ(row.Item("CustomerGroupID"), 0))
                Next
                Send(");")
            Else
                Send("""" & "-99" & """);")
            End If
        End If
    End If
    FullListSelect.Append("<\/select>")
    Send("var fullList = '" & FullListSelect.ToString() & "';")
%>

var analyticsOfferID = <%=AnalyticsCG.CustomerGroupID %>;

// This is the function that refreshes the list after a keypress.
// The maximum number to show can be limited to improve performance with
// huge lists (1000s of entries).
// The function clears the list, and then does a linear search through the
// globally defined array and adds the matches back to the list.
function handleKeyUp(maxNumToShow) {
  var selectObj, textObj, functionListLength;
  var i,  numShown;
  var searchPattern;
  
  document.getElementById("functionselect").size = "12";
  
  // Set references to the form elements
  selectObj = document.forms[0].functionselect;
  textObj = document.forms[0].functioninput;
  
  // Remember the function list length for loop speedup
  functionListLength = functionlist.length;
  // Set the search pattern depending
searchPattern = cleanSpecialChar(textObj.value);
  if (document.forms[0].functionradio[0].checked == true){
      searchPattern = "^" + searchPattern;
  }
  
  // Create a regular expression
  re = new RegExp(searchPattern,"gi");
  
  // Loop through the array and re-add matching options
  numShown = 0;
  
  if (textObj.value == '') {
    document.getElementById("cgList").innerHTML = fullList;
  } else {
    var newList = '<select class="longer" id="functionselect" name="functionselect" size="12" multiple="multiple">';
    for(i = 0; i < functionListLength; i++) {
      if(functionlist[i].search(re) != -1) {
        if (vallist[i<% Sendb(IIf(EngineID = 6, "-1", ""))%>] != "") {        
          if (vallist[i<% Sendb(IIf(EngineID = 6, "-1", ""))%>] == 1 || vallist[i<% Sendb(IIf(EngineID = 6, "-1", ""))%>] == 2 || vallist[i<% Sendb(IIf(EngineID = 6, "-1", ""))%>] == <%Sendb(NewCardholdersID)%> || vallist[i] == <%Sendb(AnalyticsCG.CustomerGroupID) %>) {
            if(selectObj[numShown] != null ) {
            selectObj[numShown].style.fontWeight = 'bold';
            selectObj[numShown].style.color = 'brown';
            }
            newList += '<option value="' + vallist[i<% Sendb(IIf(EngineID = 6, "-1", ""))%>] + '" style="color:brown;font-weight:bold;"> ' + functionlist[i] + '<\/option>';
          } else {
            newList += '<option value="' + vallist[i<% Sendb(IIf(EngineID = 6, "-1", ""))%>] + '"> ' + functionlist[i] + '<\/option>';
          }
          numShown++;
        }
      }
      // Stop when the number to show is reached
      if(numShown == maxNumToShow) {
        break;
      }
    }
    newList += '<\/select>'
    document.getElementById("cgList").innerHTML = newList;
  }
  
  removeUsed(true);
  
  // When options list whittled to one, select that entry
  if(selectObj.length == 1) {
    selectObj.options[0].selected = true;
  }
}
function removeUsed(bSkipKeyUp) {
  if (!bSkipKeyUp) handleKeyUp(99999);
  // this function will remove items from the functionselect box that are used in 
  // selected and excluded boxes
  
  var funcSel = document.getElementById('functionselect');
  var exSel = document.getElementById('selected');
  var elSel = document.getElementById('excluded');
  var i,j;
  
  for (i = elSel.length - 1; i>=0; i--) {
    for(j=funcSel.length-1;j>=0; j--) {
      if(funcSel.options[j].value == elSel.options[i].value){
        funcSel.options[j] = null;
      }
    }
  }
  
  for (i = exSel.length - 1; i>=0; i--) {
    for(j=funcSel.length-1;j>=0; j--) {
      if(funcSel.options[j].value == exSel.options[i].value){
        funcSel.options[j] = null;
      }
    }
  }

  //Below code is for removing analytics cg when we have duplicates.
  var usedList = {};
  var hasDuplicates = false;
  $("#functionselect > option").each(function () {
      if(usedList[this.text]) {
          hasDuplicates = true;
          if(this.value != -1)
            analyticsOfferID = this.value;
      } else {
          usedList[this.text] = this.value;
      }
  });

  if (hasDuplicates) {
        $("#functionselect option[value='-1']").each(function() {
          $(this).remove();
        });
  }
  console.log(analyticsOfferID);
}
function IsSelectedForGrantMembership(selectedValue){

    <%
        Dim dt As DataTable
        Dim strID As String = ""
        MyCommon.QueryStr = "select CGT.CustomerGroupID from CPE_DeliverableCustomerGroupTiers CGT " &
               " inner Join CPE_Deliverables CD On CGT.DeliverableID = CD.DeliverableID " &
                "inner Join CPE_RewardOptions CRO on CRO.RewardOptionID = CD.RewardOptionID " &
                "inner Join CPE_Incentives CI On CI.IncentiveID = CRO.IncentiveID " &
                "where CD.Deleted = 0 And CI.EngineID = 9 And CI.IncentiveID =" & OfferID
        dt = MyCommon.LRT_Select()
        If dt IsNot Nothing And dt.Rows.Count > 0 Then
            For Each row In dt.Rows
                strID &= ", " & MyCommon.NZ(row.Item("CustomerGroupID"), 0)
            Next
        End If

     %>
    var idsstr = '<%= strID %>'
    var ids = idsstr.split(',');
    for (var i = 0; i < ids.length; i++) {
        if (ids[i].trim() == selectedValue) {
            return true;
        }
    }
    return false;
}
// this function gets the selected value and loads the appropriate
// php reference page in the display frame
// it can be modified to perform whatever action is needed, or nothing
function handleSelectClick(itemSelected) {
 <%
     CMS.AMS.CurrentRequest.Resolver.AppName = "UEoffer-con-customer.aspx"
     Dim m_offer As CMS.AMS.Contract.IOffer = CMS.AMS.CurrentRequest.Resolver.Resolve(Of CMS.AMS.Contract.IOffer)()
     Dim m_defaultcust As CMS.AMS.Models.CustomerGroup = m_offer.GetOfferDefaultCustomerGroup(OfferID, EngineID)
     Dim CustGroupId As String = ""
     If (m_defaultcust IsNot Nothing) Then
         CustGroupId = m_defaultcust.CustomerGroupID
     End If

   %>
 var IsEligibleConditionExist = '<%= m_offer.IsOfferOptable(OfferID)%>';
 var defaultcusgroupid= '<%=CustGroupId%>'
 var IsOptout = '<%= Request.QueryString("mode")%>'
  var funcSel = document.getElementById('functionselect');
  var exSel = document.getElementById('selected');
  var elSel = document.getElementById('excluded');
  var i,j; 
  textObj = document.forms[0].functioninput;
  
  selectObj = document.forms[0].functionselect;
  selectedValue = document.getElementById("functionselect").value;
  if(selectedValue != ""){ selectedText = selectObj[document.getElementById("functionselect").selectedIndex].text; }
  selectboxObj = document.forms[0].selected;
  selectedboxValue = document.getElementById("selected").value;
  if(selectedboxValue != ""){ selectedboxText = selectboxObj[document.getElementById("selected").selectedIndex].text; }
  
  excludedbox = document.forms[0].excluded;
  excludedboxValue = document.getElementById("excluded").value;
  if(excludedboxValue != ""){ excludeboxText = excludedbox[document.getElementById("excluded").selectedIndex].text; }
  
  if (itemSelected == "select1") {
    if (selectedValue != "") {
     if(IsSelectedForGrantMembership(selectedValue))
     {
        alert('<% Sendb(Copient.PhraseLib.Lookup("error.includedcustomer", LanguageID))%>');
        return;
      }
    if(IsEligibleConditionExist == "True" && IsOptout != 'optout')
    {       
      if(selectedValue == 1 || selectedValue == 2)
      {
          alert('<% Sendb(Copient.PhraseLib.Lookup("term.invalidgroupforoptin", LanguageID))%>');
          return;
      }
    }

    if (selectedValue == 1) {      
        if(document.getElementById('functionNone')!=null)
        {
          document.getElementById('functionNone').checked = 'checked';
        }
        if(document.getElementById('cardTypes')!=null)
        {
          document.getElementById('cardTypes').style.display= 'none';
        }
        if(document.getElementById('functionNone')!=null)
        {
          document.getElementById('functionNone').disabled=true;
        }
        if(document.getElementById('functionYes')!=null)
        {
          document.getElementById('functionYes').disabled=true;
        }
    } 
    else {
          if(document.getElementById('cardFilter')!=null)
          {
            document.getElementById('cardFilter').disabled=false;
          }
          if(document.getElementById('functionNone')!=null)
          {
           document.getElementById('functionNone').disabled=false;
          }
          if(document.getElementById('functionYes')!=null)
          {
            document.getElementById('functionYes').disabled=false;
          }
    }


      // add items to selected box
      document.getElementById('deselect1').disabled=false;
      document.getElementById('select2').disabled=false;
      document.getElementById('save').disabled=false;
      
      while (selectObj.selectedIndex != -1) {
        selectedText = selectObj.options[selectObj.selectedIndex].text;
        selectedValue = selectObj.options[selectObj.selectedIndex].value;
        if(selectedValue == 2 || selectedValue == <%Sendb(AllCAM)%> || selectedValue == 1){
          document.getElementById('select1').disabled=true;  
          document.getElementById('select2').disabled=false;
          // someone's adding all customers we need to empty the select box
          for (i = selectboxObj.length - 1; i>=0; i--){
            selectboxObj.options[i] = null;
          }
          selectboxObj[selectboxObj.length] = new Option(selectedText,selectedValue);
          selectboxObj[selectboxObj.length-1].style.color = 'brown';
          selectboxObj[selectboxObj.length-1].style.fontWeight = 'bold';
          selectObj[selectObj.selectedIndex].selected = false;
          break;
        } else {
          selectboxObj[selectboxObj.length] = new Option(selectedText,selectedValue);
          if (selectedValue== "<%Sendb(NewCardholdersID)%>" || selectedValue == <%Sendb(AnalyticsCG.CustomerGroupID) %>) {
            selectboxObj[selectboxObj.length-1].style.color = 'brown';
            selectboxObj[selectboxObj.length-1].style.fontWeight = 'bold';
          }
          selectObj[selectObj.selectedIndex].selected = false;
        }
      }
    }
  }
  
  if (itemSelected == "deselect1") {
 
    if (selectedboxValue != "") {
        
        //If all items are selected for remove and exclude item exist
        if ($("#selected").children().length == $("#selected :selected").length && $("#excluded").children().length>0)
        {
          alert('<% Sendb(Copient.PhraseLib.Lookup("term-ValidationOnDeleteForAllSelected", LanguageID))%>');
          return;
        }
      if (selectedValue != ""  && selectedValue == 1) {
         if(document.getElementById('functionNone')!=null)
         {
           document.getElementById('functionNone').disabled=true;
         }
         if(document.getElementById('functionYes')!=null)
         {
           document.getElementById('functionYes').disabled=true;
         }
      } 
      else{
          if(document.getElementById('functionNone')!=null)
          {
            document.getElementById('functionNone').disabled=false;
          }
          if(document.getElementById('functionYes')!=null)
          {
            document.getElementById('functionYes').disabled=false;
          }
      }
        //if Eligibility condition exist then verify that default group should not be in selected list
        if (IsEligibleConditionExist == "True" && IsOptout != 'optout')
        {
          var defaultfound=0;
         $("#selected :selected").each
         (
            function()
            {
              if (this.value == defaultcusgroupid)
              {
                defaultfound=1;
                return;
              }
            }
          );
        
          if(defaultfound == 1)
          {
              alert('<% Sendb(Copient.PhraseLib.Lookup("term.deselectdefaultgroup", LanguageID))%>');
              return;
          }
        }
       
    
      // remove items from selected box
      while (document.getElementById("selected").selectedIndex != -1) {
          document.getElementById("selected").remove(document.getElementById("selected").selectedIndex);
      }
      
      if (selectedboxValue == 2 || selectedboxValue == <%Sendb(AllCAM)%> || selectedboxValue == 1) {
        document.getElementById('select1').disabled=true;
        document.getElementById('select2').disabled=false;
        document.getElementById('deselect2').disabled= true;
      }
      
      if (selectboxObj.length == 0) {
        // nothing in the select box so disable deselect
        document.getElementById('deselect1').disabled=true;
        document.getElementById('select2').disabled=true;
        document.getElementById('deselect2').disabled=true;
        // this being the case, let's also disable the page's save button
        // (but not if it's a template with "required" checked)
        if (document.getElementById("require_cg") != null) {
          if (document.getElementById("require_cg").checked == false) {
            document.getElementById('save').disabled=true;
          }
        } else {
          document.getElementById('save').disabled=true;
        }
      }
      
      if (selectedboxValue != "") {
        if (excludedbox.length == 0) {
          document.getElementById('select1').disabled=false;
        }
        document.getElementById('select2').disabled=false;
      }
      
      if (selectboxObj.length == 0 || selectedboxValue == "") {
        document.getElementById('select1').disabled=false; 
        document.getElementById('select2').disabled=true;
      } else if (selectedboxValue != 0 || selectboxObj.length != 0) {
        document.getElementById('select2').disabled=false;
      } else {
        document.getElementById('select1').disabled=false; 
        document.getElementById('select2').disabled=false;
      }
    }
  }
  
  if (itemSelected == "select2") {
    if (selectedValue != "") {
      var AnyCardholder = ("<% Sendb(Copient.PhraseLib.Lookup("term.anycardholder", LanguageID))%>").DecodeSingleQuotes();
      var excludedCt = 0;
      if (selectedValue == analyticsOfferID) {
          alert('<% Sendb(Copient.PhraseLib.Lookup("offer-con.analyticscgexcluded", LanguageID)) %>');
      } else if(isAnyCustomerSelected()) {
          // add items to excluded box
        while(document.getElementById("functionselect").selectedIndex != -1){
          selectedText = selectObj.options[selectObj.selectedIndex].text;
          selectedValue = selectObj.options[selectObj.selectedIndex].value;
          excludedbox[excludedbox.length] = new Option(selectedText,selectedValue);
          selectObj[selectObj.selectedIndex].selected = false;
        }

        // ensure that if AnyCardholder is excluded that no other group is excluded with it
        for (var i=0; i<excludedbox.length; i++) {
          if (excludedbox.options[i].value ==2) {
            excludedCt = excludedbox.options.length
            excludedbox.options.length = 0;
            excludedbox[excludedbox.length] = new Option(AnyCardholder,'2');

            if (excludedCt > 1) {
              alert('<% Sendb(Copient.PhraseLib.Lookup("ueoffer-con-customer.TargetedGroupError", LanguageID))%>');
            }
            break;
          }
        }
      } else {
        if (selectedValue == 1) {
          alert('<% Sendb(Copient.PhraseLib.Lookup("offer-con.anycustomerexcluded", LanguageID))%>');
        }else if (selectedValue == 2) {
          alert('<% Sendb(Copient.PhraseLib.Lookup("offer-con.anycardholderexcluded", LanguageID))%>');
        }else if (selectedValue != <%Sendb(NewCardholdersID)%>) {
          // add items to excluded box
          while(document.getElementById("functionselect").selectedIndex != -1){
            selectedText = selectObj.options[selectObj.selectedIndex].text;
            selectedValue = selectObj.options[selectObj.selectedIndex].value;
            excludedbox[excludedbox.length] = new Option(selectedText,selectedValue);
            selectObj[selectObj.selectedIndex].selected = false;
          }
          if (excludedbox.length == 1) {
            document.getElementById('select2').disabled=false;
            // need to disable deselection on the selected box also since we added an excluded group
            document.getElementById('deselect1').disabled=false;
            document.getElementById('deselect2').disabled=false;
          }
        } else {
          alert("<% Sendb(Copient.PhraseLib.Lookup("offer-con.newcardholderexcluded", LanguageID))%>");
        }
      }
    } 
  }
  
  if (itemSelected == "deselect2") {
    if (excludedboxValue != "") {
      // remove items from excluded box
      while (document.getElementById("excluded").selectedIndex != -1) {
        document.getElementById("excluded").remove(document.getElementById("excluded").selectedIndex);
      }
      if (excludedbox.length == 0) {
        document.getElementById('select2').disabled=false;
        document.getElementById('deselect1').disabled=false;
        document.getElementById('deselect2').disabled=true;
      }
    }
  }
  
  updateExceptionButtons();
  
  // remove items from large list that are in the other lists
  removeUsed(false);
  return true;
}

function isAnyCustomerSelected() {
  var exSel = document.getElementById('selected');
  var anyCustSel = false;

  if (exSel != null && exSel.options.length > 0) {
   anyCustSel = (exSel.options[0].value == 1);
  } 
  
  return anyCustSel; 
}

// ensures that if AnyCustomer is selected that the correct values are in place for both selected and excluded
function handleAnyCustomer() {
  var exSel = document.getElementById('selected');
  var anyCustSel = false;

  if (exSel != null && exSel.options.length > 0) {
   for (var i=0; i < exSel.options.length; i++) {
     anyCustSel = (exSel.options[i].value == 1);
   }
  
   if (anyCustSel) {
     // remove any other selected group as "Any Customer" subsumes all other groups
     if (exSel.options.length > 1) {
       for (var i=0; i < exSel.options.length; i++) {
         if (exSel.options[i].value != 1) {
           exSel.options[i] = null;
           i--;
         }
       }  
     }
     
     // check to ensure that the exclusions include    
   }

  } 
  
}

function saveForm() {
  var funcSel = document.getElementById('functionselect');
  var exSel = document.getElementById('excluded');
  var elSel = document.getElementById('selected');
  var objinfobar = document.getElementById('infobar');
  var isOfferOptable="FALSE";
  var eligibleIncluded="";
  var eligibleExcluded="";
  var notFoundInEligibleExcludedList="";
  var needstoremovefromeligibleincludedlist="";
  var needstoremovefromeligibleincludedlistText="";
  var i,j;
  var selectList = "";
  var excludededList = "";
  var htmlContents = "";
  var isFoundInIncludedEligibleCondition="";
  objinfobar.innerHTML= '';
  objinfobar.style.display='none';
  var pagemode = '<%= Request.QueryString("mode")%>'
  
  isOfferOptable = document.getElementById('HdnIsOfferOptable').value;
  eligibleIncluded = document.getElementById('Hdnincludedeligiblegroup').value;
  eligibleExcluded = document.getElementById('Hdnexcludedeligiblegroup').value;
  
  // assemble the list of values from the selected box
  for (i = elSel.length - 1; i>=0; i--) {
    if(elSel.options[i].value != ""){
      if(selectList != "") { selectList = selectList + ","; }
      selectList = selectList + elSel.options[i].value;
    }
  }
  for (i = exSel.length - 1; i>=0; i--) {
    if(exSel.options[i].value != ""){
      
      if(excludededList != "") { excludededList = excludededList + ","; }
      excludededList = excludededList + exSel.options[i].value;
     
      //validation for eligiblegroups if offer is optable
      if(isOfferOptable=="TRUE" && pagemode != 'optout' )
      {
        
        isFoundInIncludedEligibleCondition="";
        if(eligibleIncluded.indexOf(exSel.options[i].value) != -1)
        {
          isFoundInIncludedEligibleCondition="TRUE";

          if(needstoremovefromeligibleincludedlist != "")
          { needstoremovefromeligibleincludedlist = needstoremovefromeligibleincludedlist + ","; }
          
          if(needstoremovefromeligibleincludedlistText !="")
          {needstoremovefromeligibleincludedlistText = needstoremovefromeligibleincludedlistText + ",";}

          needstoremovefromeligibleincludedlist = needstoremovefromeligibleincludedlist + exSel.options[i].value;
          needstoremovefromeligibleincludedlistText = needstoremovefromeligibleincludedlistText + exSel.options[i].text;
        }

        if(isFoundInIncludedEligibleCondition == "" && eligibleExcluded.indexOf(exSel.options[i].value) == -1 )
        {
              if(notFoundInEligibleExcludedList != "")
              { notFoundInEligibleExcludedList = notFoundInEligibleExcludedList + ","; }
              notFoundInEligibleExcludedList = notFoundInEligibleExcludedList + exSel.options[i].value;
        }
      }
    }
  }

  if(isOfferOptable=="TRUE" && pagemode != 'optout')
  {
    if(needstoremovefromeligibleincludedlist!="")
    {
       var confmsg="";
        
       if(needstoremovefromeligibleincludedlist.split(",").length >= eligibleIncluded.split(",").length)
       {
         confmsg= "<% Sendb(Copient.PhraseLib.Lookup("term.msgsavediscarded", LanguageID))%>";
         objinfobar.innerHTML=confmsg.format([needstoremovefromeligibleincludedlistText]);
         objinfobar.style.display='block';
         return false;
       }
       else
       {
         confmsg = "<% Sendb(Copient.PhraseLib.Lookup("term.confirmtodeleligiblegroup", LanguageID))%>";
         confmsg= confmsg.format([needstoremovefromeligibleincludedlistText])
         if(!confirm(confmsg))
         {
         return false;
         }
       }
    }
  }
  
    if(document.getElementById("functionYes")!=null)
    {
		if (document.getElementById("functionYes").checked == true){
		   if (document.getElementById("cardTypeSelected") == undefined || document.getElementById("cardTypeSelected") == null || document.getElementById("cardTypeSelected").value.length <= 0)
			  {
				  alert("<% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID) & " " & Copient.PhraseLib.Lookup("term.cardtype", LanguageID))%>");
				  return false;
			  }
		 }
	 }

  // time to build up the hidden variables to pass for saving
  htmlContents = "<input type=\"hidden\" name=\"selGroups\" value=" + selectList + "> ";
  htmlContents = htmlContents + "<input type=\"hidden\" name=\"exGroups\" value=" + excludededList + ">";
  htmlContents = htmlContents + "<input type=\"hidden\" name=\"tomoveeligibleincludedgroups\" value=" + needstoremovefromeligibleincludedlist + ">";
  htmlContents = htmlContents + "<input type=\"hidden\" name=\"toaddeligibleexcludedgroups\" value=" + notFoundInEligibleExcludedList + ">";
  document.getElementById("hiddenVals").innerHTML = htmlContents;
  
  //alert(htmlContents);
  
  return true;
}

function updateButtons(){
  var selVal;
  if(document.getElementById('selected').length > 0){
    selVal = document.getElementById('selected').options[0].value;
    if(selVal == 2 || selVal == 1 || selVal == <%Sendb(AllCAM)%>) {
      // all customers is in the selected box so disable adding another
      document.getElementById('select1').disabled=true; 
      if(document.forms[0].selected.length == 1) {
      //since one or more cg can be excluded, do not disable selection
        document.getElementById('select2').disabled=false; 
        if(document.forms[0].excluded.length == 0) { 
          // nothing is excluded so allow excluding one
          document.getElementById('select2').disabled=false; 
          document.getElementById('deselect2').disabled=true; 
          document.getElementById('deselect1').disabled=false; 
        } else {
          document.getElementById('select2').disabled=false; 
          document.getElementById('deselect2').disabled=false; 
          document.getElementById('deselect1').disabled=true; 
        }
      }
    } else {
      // something is selected but its not all customers
      document.getElementById('select1').disabled=false; 
      document.getElementById('deselect1').disabled=false;
      document.getElementById('select2').disabled=false; 
      document.getElementById('deselect2').disabled=true;
    }
  }
  <%
      m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
      If Not IsTemplate Then
          If Not (Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And Disallow_Edit)) Then
              Send("  disableAll();")
          End If
      Else
          If Not (Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer) Then
              Send("  disableAll();")
          End If
      End If
  %>
}

function updateExceptionButtons() {
  var exSel = document.getElementById('excluded');
  var elSel = document.getElementById('selected');
  var bEligible = false;
  
  // check if there already is an excluded group, if so disable select and enable deselect
  if (exSel != null && elSel != null && exSel.options.length == 0) {
    // check if a exception-qualify customer group is in the selected list
    for (var i=0; i < elSel.options.length && !bEligible; i++) {
      bEligible = isExceptionGroup(elSel.options[i].value)
    }
  } else if (exSel != null && exSel.options.length > 0) {
    document.getElementById('select2').disabled=false; 
    document.getElementById('deselect2').disabled=false; 
  } else {
    document.getElementById('select2').disabled=true; 
    document.getElementById('deselect2').disabled=true; 
  }
}

function isExceptionGroup(groupID) {
  var bRetVal = false;
  
  for (var i=0; i < exceptlist.length && !bRetVal; i++) {
    bRetVal = (exceptlist[i] == groupID) 
  }
  
  return bRetVal;
}

function handleRequiredToggle() {
  if(document.forms[0].selected.length == 0) {
    if (document.getElementById("require_cg").checked == false) {
      document.getElementById('save').disabled=true;
    } else {
      document.getElementById('save').disabled=false;
    }
  }
  if (document.getElementById("require_cg").checked == true) {
    document.getElementById("Disallow_Edit").checked=false;
  }
}

function disableAll() {
  document.getElementById('select1').disabled=true;
  document.getElementById('select2').disabled=true;
  document.getElementById('deselect1').disabled=true;
  document.getElementById('deselect2').disabled=true;
  document.getElementById('functionselect').disabled=true;
  document.getElementById('selected').disabled=true;
  document.getElementById('excluded').disabled=true;
}

function ShowCardFilter(){
    if(document.getElementById("functionYes").checked == true){
    document.getElementById("cardTypes").style.display = 'block';
    var selectedCardTypes=document.getElementById("cardTypeSelected").value;
    xmlhttpPost_GetCardTypesData('../OfferFeeds.aspx','Mode=GetCardTypeForCustomerCondition&conditionID=' + document.mainform.ConditionID.value , selectedCardTypes );
    }
  }

function HideCardFilter(){
    if (document.getElementById("functionNone").checked == true){
    document.getElementById("cardTypes").style.display = 'none';
    }
  }
function showCustomerApproval(){
    if (document.getElementById("yes").checked == true){
            <% If disableCustApproval Then %>
                alert('<% Sendb(Copient.PhraseLib.Lookup("info.enableCustomerApproval", LanguageID))%>');
                document.getElementById("details").style.display = 'none';
                document.getElementById("no").checked = 'checked';
            <% Else %>
                document.getElementById("details").style.display = 'block';
            <% End If %>
    
    }
}
function hideCustomerApproval(){
    if (document.getElementById("no").checked == true){
    document.getElementById("details").style.display = 'none';
    }
}
function addRowDOM(tblId,cardTypeId,cardTypeDesc)
{
           var lastRow =0
           var tblBody = document.getElementById(tblId);
           lastRow=tblBody.rows.length - 1;
           var newRow = tblBody.insertRow(lastRow);
              newRow.id = 'tr_'+ cardTypeId ;

            //Create Delete button
              var newCell0 = newRow.insertCell(0);
                            
              var newInput = document.createElement('input');
              newInput.type = 'button';
              newInput.id = 'td_'+ cardTypeId ;
              newInput.title = '<% Sendb(Copient.PhraseLib.Lookup("term.delete", LanguageID))%>' ;
              newInput.value = 'X';
              newInput.class = 'ex';
              newInput.setAttribute("onclick","deleteCardType("+cardTypeId+")");
              newInput.setAttribute("class", "ex");
              
              newCell0.appendChild(newInput);
            //end delete button

          var newCell1 = newRow.insertCell(1);
          newCell1.appendChild(document.createTextNode(cardTypeDesc));
          newCell1.setAttribute("class", "shaded"); 
            
            
}

            
    function addCardType(){
    
    var cardTypeId=$('#ddlCardType').val();
    var cardTypeDesc=$('#ddlCardType :selected').text();

    addRowDOM("cardtypeDisplay",cardTypeId,cardTypeDesc);
    
    //Persist added id in a hidden field
    document.getElementById("cardTypeSelected").value= filterCardTypeIds("add",cardTypeId);
    var selectedCardTypes=document.getElementById("cardTypeSelected").value;
   
    xmlhttpPost_GetCardTypesData('../OfferFeeds.aspx','Mode=GetCardTypeForCustomerCondition&conditionID=' + document.mainform.ConditionID.value , selectedCardTypes);
    }

function deleteCardType(cardTypeId){
            
           $("#tr_"+cardTypeId).remove();
            //Persist deleted id in a hidden field
            document.getElementById("cardTypeSelected").value=filterCardTypeIds("remove",cardTypeId);
            selectedCardTypes=document.getElementById("cardTypeSelected").value;
            xmlhttpPost_GetCardTypesData('../OfferFeeds.aspx','Mode=GetCardTypeForCustomerCondition&conditionID=' + document.mainform.ConditionID.value , selectedCardTypes);
            }

            function filterCardTypeIds(op,cardTypeId)
            {

                var finalCardTypeIds="";
                var TempArr=new Array();
                var selectedValue="";
                if(op == "add")
                {
                        TempArr=document.getElementById("cardTypeSelected").value.split(',');
                        if ($.inArray(cardTypeId.toString() , TempArr) == -1)
                           TempArr.push(cardTypeId);
                }
                else if(op == "remove")
                {
                    TempArr=document.getElementById("cardTypeSelected").value.split(',');
                    if ($.inArray(cardTypeId.toString() , TempArr) > -1)
                           TempArr.splice($.inArray(cardTypeId.toString(), TempArr),1);
                    
                }
            if(TempArr.length>0)
                finalCardTypeIds= TempArr.toString();

            return finalCardTypeIds;
            }
           
             //javascript page load


      function xmlhttpPost_GetCardTypesData(strURL, qryStr,selectedIds) {
            
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
      document.getElementById("btnAddCardType").disabled=true;
      self.xmlHttpReq.onreadystatechange = function() {
          if (self.xmlHttpReq !=null && self.xmlHttpReq.readyState == 4) {
                  handleresponseGetCardTypesData(self.xmlHttpReq.responseText,selectedIds);
          }
      }
      self.xmlHttpReq.send(qryStr);
      return false;
  }
  function handleresponseGetCardTypesData(responseText,selectedIds) {    
     if(responseText.length>0)
     {
         var savedCardTypesArr= selectedIds.split(',');
         var ajaxResponse = responseText.split('~');
         if(ajaxResponse[0] =="success" && ajaxResponse[1] != -1)
         {
             document.getElementById("ddlCardType").options.length = 0;
             
             var data = JSON.parse(ajaxResponse[1]);
        var i=0;
        for (var key in data)
        {
                if(savedCardTypesArr==null || $.inArray(key, savedCardTypesArr) == -1) 
            {
                document.getElementById("ddlCardType").options[i] = new Option(data[key],key);
                i++;
            }
        }
            
            if(document.getElementById("hdnIsDisabled").value=="0")
            {
                document.getElementById("btnAddCardType").disabled=false;
                document.getElementById("ddlCardType").disabled=false;
            }
            else
            {
                document.getElementById("btnAddCardType").disabled=true;
                document.getElementById("ddlCardType").disabled=true;
            }
            //All cards were selected in the condition ,no cardtypes were available to further select
            var hideAddRow = $('.trhideclass1');
            if( i == 0)
                hideAddRow.hide();
            else
                hideAddRow.show();
        }
       
       
     }
  }

       
$( document ).ready(function() {
   
   if( $('#cardTypes').css('display') == 'block')
   {
        selectedCardTypes=document.getElementById("cardTypeSelected").value;
        xmlhttpPost_GetCardTypesData('../OfferFeeds.aspx','Mode=GetCardTypeForCustomerCondition&conditionID=' + document.mainform.ConditionID.value ,selectedCardTypes);
   }
     
});  

      
</script>
<%
    Send("<script type=""text/javascript"">")
    Send("function ChangeParentDocument() { ")


    If (EngineID = 3) Then
        Send("  if (opener != null) {")
        Send("    var newlocation = '/logix/web-offer-con.aspx?OfferID=" & OfferID & "'; ")
        Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
        Send("  opener.location = '/logix/web-offer-con.aspx?OfferID=" & OfferID & "'; ")
    ElseIf (EngineID = 5) Then
        Send("  if (opener != null) {")
        Send("    var newlocation = '/logix/email-offer-con.aspx?OfferID=" & OfferID & "'; ")
        Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
        Send("  opener.location = '/logix/email-offer-con.aspx?OfferID=" & OfferID & "'; ")
    ElseIf (EngineID = 6) Then
        Send("  if (opener != null) {")
        Send("    var newlocation = '/logix/CAM/CAM-offer-con.aspx?OfferID=" & OfferID & "'; ")
        Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
        Send("  opener.location = '/logix/CAM/CAM-offer-con.aspx?OfferID=" & OfferID & "'; ")
    ElseIf (EngineID = 9) Then
        Send("  if (opener != null) {")
        Send("    var newlocation = '/logix/UE/UEoffer-con.aspx?OfferID=" & OfferID & "'; ")
        Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
        Send("  opener.location = '/logix/UE/UEoffer-con.aspx?OfferID=" & OfferID & "'; ")
    Else
        Send("  if (opener != null) {")
        Send("    var newlocation = '/logix/CPEoffer-con.aspx?OfferID=" & OfferID & "'; ")
        Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
        Send("  opener.location = '/logix/CPEoffer-con.aspx?OfferID=" & OfferID & "'; ")
    End If
    Send("} ")
    Send("} ")
    Send("} ")
    Send("</script>")
    Send_HeadEnd()
    If Request.QueryString("mode") = "optout" Then
        If (IsTemplate) Then
            Send_BodyBegin(13)
        Else
            Send_BodyBegin(3)
        End If
    Else
        If (IsTemplate) Then
            Send_BodyBegin(12)
        Else
            Send_BodyBegin(2)
        End If
    End If

    If (Logix.UserRoles.AccessOffers = False AndAlso Not IsTemplate) Then
        Send_Denied(2, "perm.offers-access")
        GoTo done
    End If
    If (Logix.UserRoles.AccessTemplates = False AndAlso IsTemplate) Then
        Send_Denied(2, "perm.offers-access-templates")
        GoTo done
    End If
    If (BannersEnabled AndAlso Not Logix.IsAccessibleOffer(AdminUserID, OfferID)) Then
        Send("<script type=""text/javascript"" language=""javascript"">")
        Send("  function ChangeParentDocument() { return true; } ")
        Send("</script>")
        Send_Denied(1, "banners.access-denied-offer")
        Send_BodyEnd()
        GoTo done
    End If
%>
<form action="#" name="mainform" id="mainform" onsubmit="return saveForm();">
  <span id="hiddenVals"></span>
  <input type="hidden" id="OfferID" name="OfferID" value="<% Sendb(OfferID)%>" />
  <input type="hidden" id="Name" name="Name" value="<% Sendb(Name)%>" />
  <input type="hidden" id="ConditionID" name="ConditionID" value="<% Sendb(ConditionID)%>" />
  <input type="hidden" id="EngineID" name="EngineID" value="<% Sendb(EngineID)%>" />
  <input type="hidden" id="mode" name="mode" value="<%=Request.QueryString("mode") %>" />
   <input type="hidden" id="Hdnincludedeligiblegroup" value="<% Sendb(EligibleIncludedcustomergroups)%>" />
   <input type="hidden" id="Hdnexcludedeligiblegroup" value="<% Sendb(EligibleExcludedcustomergroups)%>" />
   <input type="hidden" id="HdnIsOfferOptable" value="<% Sendb(IsEligibilityConditionExistForOffer)%>" />
   <input type="hidden" id="hdnIsDisabled" value="<% 
       If String.IsNullOrEmpty(DisabledAttribute) Then
           Sendb("0")
       Else
           Sendb("1")
       End If
       %>" />
  <input type="hidden" id="IsTemplate" name="IsTemplate" value="<% 
      If (IsTemplate) Then
          Sendb("IsTemplate")
      Else
          Sendb("Not")
      End If
        %>" />
  <div id="intro">
    <%If (IsTemplate) Then
            Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.customercondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
        Else
            Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.customercondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
        End If
    %>
    <div id="controls">
      <% If (IsTemplate) Then%>
      <span class="temp">
        <input type="checkbox" class="tempcheck" id="Disallow_Edit" name="Disallow_Edit"<% If (Disallow_Edit) Then Sendb(" checked=""checked""")%> />
        <label for="Disallow_Edit"><% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
      </span>
      <% End If%>
      <% 
          m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
          If ((Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
              If Not IsTemplate Then
                  If (Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And Disallow_Edit) And Not IsOfferWaitingForApproval(OfferID)) Then Send_Save()
              Else
                  If (Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer) Then Send_Save()
              End If
          End If
      %>
    </div>
  </div>
  <div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
     <% If (infoMessage = "") Then Send("<div id=""infobar"" class=""red-background"" style='display:none'></div>")%>
    <div id="column1">
      <div class="box" id="selector">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.customergroup", LanguageID))%>
          </span>
          <% If (IsTemplate) Then%>
          <span class="tempRequire">
            <input type="checkbox" class="tempcheck" id="require_cg" name="require_cg" onclick="handleRequiredToggle();"<% If (RequireCG) Then Sendb(" checked=""checked""")%> />
            <label for="require_cg"><% Sendb(Copient.PhraseLib.Lookup("term.required", LanguageID))%></label>
          </span>
          <% ElseIf (FromTemplate And RequireCG) Then%>
          <span class="tempRequire">*
            <%Sendb(Copient.PhraseLib.Lookup("term.required", LanguageID))%>
          </span>
          <% End If%>
        </h2>
        <input type="radio" id="functionradio1" name="functionradio" <% If (MyCommon.Fetch_SystemOption(175) = "1") Then Sendb(" checked=""checked""")%> <% Sendb(DisabledAttribute)%> /><label for="functionradio1"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
        <input type="radio" id="functionradio2" name="functionradio" <% If (MyCommon.Fetch_SystemOption(175) = "2") Then Sendb(" checked=""checked""")%> <% Sendb(DisabledAttribute)%> /><label for="functionradio2"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
        <input type="text" class="medium" id="functioninput" name="functioninput" maxlength="100" onkeyup="handleKeyUp(99999);" value=""<% Sendb(DisabledAttribute)%> /><br />
        <div id="cgList">
          <select class="longer" id="functionselect" name="functionselect" size="12"<% Sendb(DisabledAttribute)%>>
            <%
'MyCommon.QueryStr = "Select CustomerGroupID,Name from CustomerGroups with (NoLock) where deleted=0 and AnyCustomer<>1 and CustomerGroupID <> 2 and CustomerGroupID is not null and BannerID is null and NewCardholders=0 order by Name"
'rst = MyCommon.LRT_Select
'If (rst.Rows.Count > 0) Then
'  For Each row In rst.Rows
'    Send("<option value=""" & MyCommon.NZ(row.Item("CustomerGroupID"), -1) & """>" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
'  Next
'End If
            %>        
          </select>
        </div>
        <br />
        <br class="half" />
        <b><% Sendb(Copient.PhraseLib.Lookup("term.selectedcustomers", LanguageID))%>:</b>
        <br />
        <input type="button" class="regular select" id="select1" name="select1" value="&#9660; <% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID))%>" onclick="handleSelectClick('select1');" />&nbsp;
        <input type="button" class="regular deselect" id="deselect1" name="deselect1" value="<% Sendb(Copient.PhraseLib.Lookup("term.deselect", LanguageID))%> &#9650;" disabled="disabled" onclick="handleSelectClick('deselect1');" /><br />
        <br class="half" />
        <select class="longer" id="selected" name="selected" multiple="multiple" size="2"<% Sendb(DisabledAttribute)%>>
          <%
              ' alright lets find the currently selected groups on page load
              MyCommon.QueryStr = "select CG.CustomerGroupID,Name from CPE_IncentiveCustomerGroups as ICG with (NoLock) left join CustomerGroups as CG with (NoLock) " &
                  " on CG.CustomerGroupID=ICG.CustomerGroupID where RewardOptionID=" & roid &
                  " and ICG.deleted=0 and ExcludedUsers=0 and ICG.CustomerGroupID is not null"
              If Request.QueryString("mode") = "optout" Then
                  MyCommon.QueryStr = MyCommon.QueryStr & " and IsOptInGroup=0"
              End If
              rst = MyCommon.LRT_Select
              For Each row In rst.Rows
                  If MyCommon.NZ(row.Item("CustomerGroupID"), 0) = 1 Then
                      Send("<option value=""1"" style=""color:brown;font-weight:bold;"">" & Copient.PhraseLib.Lookup("term.anycustomer", LanguageID) & "</option>")
                      IsAnyCustomer = True
                  ElseIf MyCommon.NZ(row.Item("CustomerGroupID"), 0) = 2 Then
                      Send("<option value=""2"" style=""color:brown;font-weight:bold;"">" & Copient.PhraseLib.Lookup("term.anycardholder", LanguageID) & "</option>")
                  ElseIf MyCommon.NZ(row.Item("CustomerGroupID"), 0) = NewCardholdersID Then
                      Send("<option value=""" & NewCardholdersID & """ style=""color:brown;font-weight:bold;"">" & Copient.PhraseLib.Lookup("term.newcardholders", LanguageID) & "</option>")
                  ElseIf MyCommon.NZ(row.Item("CustomerGroupID"), 0) = AnalyticsCG.CustomerGroupID Then
                      If MyCommon.NZ(row.Item("Name"), "") = analyticsCGService.GetGenericAnalyticsCustomerGroup().Name Then
                          Send("<option value=""" & MyCommon.NZ(row.Item("CustomerGroupID"), 0) & """ style=""color:brown;font-weight:bold;"">" & Copient.PhraseLib.Lookup("term.getsegmentfromanalytics", LanguageID) & "</option>")
                      Else
                          Send("<option value=""" & MyCommon.NZ(row.Item("CustomerGroupID"), 0) & """ style=""color:brown;font-weight:bold;"">" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
                      End If
                  Else
                      Send("<option value=""" & MyCommon.NZ(row.Item("CustomerGroupID"), 0) & """ >" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
                  End If
              Next
          %>
        </select>
        <br />
        <br class="half" />
        <b><% Sendb(Copient.PhraseLib.Lookup("term.excludedcustomers", LanguageID))%>:</b>
        <br />
        <input type="button" class="regular select" id="select2" name="select2" value="&#9660; <% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID))%>" disabled="disabled" onclick="handleSelectClick('select2');" />&nbsp;
        <input type="button" class="regular deselect" id="deselect2" name="deselect2" value="<% Sendb(Copient.PhraseLib.Lookup("term.deselect", LanguageID))%> &#9650;" disabled="disabled" onclick="handleSelectClick('deselect2');" /><br />
        <br class="half" />
        <select class="longer" id="excluded" name="excluded" multiple="multiple" size="2"<% Sendb(DisabledAttribute)%>>
          <%
              ' alright lets find the currently selected groups on page load
              MyCommon.QueryStr = "select CG.CustomerGroupID,Name from CPE_IncentiveCustomerGroups as ICG with (NoLock) left join CustomerGroups as CG with (NoLock) " &
                                  " on CG.CustomerGroupID=ICG.CustomerGroupID where RewardOptionID=" & roid &
                                  " and ICG.deleted=0 and ExcludedUsers=1 and ICG.CustomerGroupID is not null"
              If Request.QueryString("mode") = "optout" Then
                  MyCommon.QueryStr = MyCommon.QueryStr & " and IsOptInGroup=0"
              End If
              rst = MyCommon.LRT_Select
              For Each row In rst.Rows
                  If MyCommon.NZ(row.Item("CustomerGroupID"), 0) = 1 Then
                      Send("<option value=""1"" >" & Copient.PhraseLib.Lookup("term.anycustomer", LanguageID) & "</option>")
                  ElseIf MyCommon.NZ(row.Item("CustomerGroupID"), 0) = 2 Then
                      Send("<option value=""2"" >" & Copient.PhraseLib.Lookup("term.anycardholder", LanguageID) & "</option>")
                  ElseIf MyCommon.NZ(row.Item("CustomerGroupID"), 0) = NewCardholdersID Then
                      Send("<option value=""" & NewCardholdersID & """ >" & Copient.PhraseLib.Lookup("term.newcardholders", LanguageID) & "</option>")
                  Else
                      Send("<option value=""" & MyCommon.NZ(row.Item("CustomerGroupID"), 0) & """>" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
                  End If
              Next
          %>
        </select>
        <hr class="hidden" />
      </div>
    </div>
    
    <div id="gutter">
    </div>
    <% If (EngineID <> 6) Then%>
    <div id="column2">
      <div class="box" id="hhoptions">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.options", LanguageID))%>
          </span>
        </h2>

          <div id="cardFilter">
          <span >
            <% Sendb(Copient.PhraseLib.Lookup("term.cardtype", LanguageID) & " " & Copient.PhraseLib.Lookup("term.filter", LanguageID) & ": ")%>
                        <% 
                            If (Not String.IsNullOrEmpty(ConditionID)) Then
                                rst = m_CustomerConditionService.GetCardTypesForCustomerCondition(roid)
                                If rst.Rows.Count > 0 Then
                                    For Each row In rst.Rows
                                        cardIDs = cardIDs & MyCommon.NZ(row.Item("CardTypeID"), 0) & ","
                                    Next
                                    cardIDs = cardIDs.Trim().Substring(0, cardIDs.Length - 1)
                                    selectedCards = m_CustomerConditionService.GetSpecificCardTypes(cardIDs)
                                End If
                            End If

            %>
            <input type="radio" id="functionNone" value="0" name="radioCardTypes" <% If (cardIDs = "") Then Sendb("checked=""checked""")%> <% Sendb(DisabledAttribute)%> <% If (IsAnyCustomer) Then Sendb("disabled=""disabled""")%> onclick="HideCardFilter();"/><label for="functionNone"><% Sendb(Copient.PhraseLib.Lookup("term.none", LanguageID))%></label>
            &nbsp;
           <input type="radio" id="functionYes" value="1" name="radioCardTypes"  <% If (cardIDs <> "") Then Sendb("checked=""checked""")%> <% Sendb(DisabledAttribute)%> <% If (IsAnyCustomer) Then Sendb("disabled=""disabled""")%> onclick="ShowCardFilter();"/><label for="functionYes"><% Sendb(Copient.PhraseLib.Lookup("term.yes", LanguageID))%></label><br/>
           <br/>
          </span>
    
        <div id="cardTypes" <% If (cardIDs = "") Then Sendb("style=""display:none;""") %> <% Sendb(DisabledAttribute)%>>
        
         <table id="cardtypeDisplay">
               <tr>
                <th><% Sendb(Copient.PhraseLib.Lookup("term.remove", LanguageID))%></th>
                <th><% Sendb(Copient.PhraseLib.Lookup("term.cardtype", LanguageID))%></th>
               </tr>
              <%  
                  If (selectedCards IsNot Nothing AndAlso selectedCards.Rows.Count > 0) Then
                      cardIDs = ""
                      For Each row2 In selectedCards.Rows
                          cardIDs = cardIDs & MyCommon.NZ(row2.Item("CardTypeID"), 0) & ","
                %>
                         
                        <tr id="tr_<% Sendb(MyCommon.NZ(row2.Item("CardTypeID"), 0))%>" ><td>
                        <input type="button" value="X" title="<% Sendb(Copient.PhraseLib.Lookup("term.delete", LanguageID))%>" name="td_<% Sendb(MyCommon.NZ(row2.Item("CardTypeID"), 0))%>"  id="td_<% Sendb(MyCommon.NZ(row2.Item("CardTypeID"), 0))%> " class="ex" onclick="javascript:deleteCardType(<% Sendb(MyCommon.NZ(row2.Item("CardTypeID"), 0))%>);" <% Sendb(DisabledAttribute)%>/></td>
                        <td class="shaded"> <span><% Sendb(Copient.PhraseLib.Lookup(MyCommon.NZ(row2.Item("PhraseTerm"), ""), LanguageID))%> </span></td></tr>
                         
                  <%  
                          Next
                          cardIDs = cardIDs.Trim().Substring(0, cardIDs.Length - 1)
                      End If
                %>
               <tr class="trhideclass1">
            <td><input type="button" class="add" id="btnAddCardType" value="<% Sendb(Copient.PhraseLib.Lookup("term.add", LanguageID))%>" onclick="javascript: addCardType();" /></td>
            <td><select class="medium" id="ddlCardType" name="ddlCardType"></select></td>
            </tr>
            
          </table>
             
             <input type="hidden" id="cardTypeSelected" name="cardTypeSelected" value="<% Sendb(cardIDs) %>" />
          </div>
              </div><br/>
          <input type="checkbox" class="tempcheck" id="offline" name="offline"<% If (MetWhenOffline) Then Sendb(" checked=""checked""")%><% Sendb(DisabledAttribute)%> />&nbsp;
         <label for="offline"><% Sendb(Copient.PhraseLib.Lookup("ueoffer-con-customer.metwhenoffline", LanguageID))%></label>
      </div>
        <div class="box" id="customerApproval" <% If restrictConditionForRPOS Then Sendb("style=""display:none;""") %> >
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.customerapproval", LanguageID))%>
                </span>
            </h2>
            <div id="approvalDetails">
                <span><% Sendb(Copient.PhraseLib.Lookup("term.customerapproval", LanguageID) & " " & Copient.PhraseLib.Lookup("term.required", LanguageID) & ":") %> </span>&nbsp;&nbsp;
                
                <input type="radio" id="no" value="0" name="approvalRadio" <% If (custApproval.CustomerApprovalID = 0) Then Sendb("checked=""checked""")%> <% Sendb(DisabledAttribute)%> onclick="hideCustomerApproval();" /><label for="no"><%Sendb(Copient.PhraseLib.Lookup("term.no", LanguageID)) %></label>
                <input type="radio" id="yes" value="1" name="approvalRadio" <% If (custApproval.CustomerApprovalID > 0) Then Sendb("checked=""checked""")%> <% Sendb(DisabledAttribute)%> onclick="showCustomerApproval();" /><label for="yes"><%Sendb(Copient.PhraseLib.Lookup("term.yes", LanguageID)) %></label><br /><br />

                <table id="details" <% If (custApproval.CustomerApprovalID = 0) Then Sendb("style=""display:none;""")  %><% Sendb(DisabledAttribute)%>>
                    <tbody>
                        <tr>
                            <td><label for="mesgDesc"><% Sendb(Copient.PhraseLib.Lookup("term.message", LanguageID) & " " & Copient.PhraseLib.Lookup("term.description", LanguageID) & ":") %></label></td>
                            <td>
                                <% If (custApproval.MessageDescription Is Nothing) Then custApproval.MessageDescription = ""%>
                                <%
                                    MLI.ItemID = custApproval.CustomerApprovalID
                                    MLI.MLTableName = "CPE_IncentiveCustomerApprovalTranslation"
                                    MLI.MLColumnName = "MessageDescription"
                                    MLI.MLIdentifierName = "CustomerApprovalID"
                                    MLI.StandardTableName = "CPE_IncentiveCustomerApproval"
                                    MLI.StandardColumnName = "MessageDescription"
                                    MLI.StandardIdentifierName = "CustomerApprovalID"
                                    MLI.StandardValue = custApproval.MessageDescription
                                    MLI.InputName = "mesgDescText"
                                    MLI.InputID = "mesgDescText"
                                    MLI.InputType = "text"
                                    MLI.MaxLength = 256
                                    MLI.Disabled = IIf(DisabledAttribute <> "", True, False)
                                    ' MLI.CSSStyle = "width:350px;"
                                    Send(Localization.SendTranslationInputs(MyCommon, MLI, Request.QueryString, 9))
                                %>
                            </td>
                        </tr>
                        <tr>
                            <td><label for="approval"><% Sendb(Copient.PhraseLib.Lookup("term.approvallimit", LanguageID) & ": ") %></label></td>
                            <td>
                                <select id="approvalLimit" name="approvalLimit"<% Sendb(DisabledAttribute)%>>
                                    <%
                                        Dim dtALTypes As DataTable = m_CustomerConditionService.GetCustomerApprovalLimitTypes(LanguageID)
                                        If dtALTypes IsNot Nothing AndAlso dtALTypes.Rows.Count > 0 Then
                                            For Each row In dtALTypes.Rows
                                                Send("  <option value=""" & row.Item("ApprovalLimitTypeID") & """" & IIf(custApproval.ApprovalType = row.Item("ApprovalLimitTypeID"), " selected=""selected""", "") & ">" & row.Item("Phrase") & "</option>")
                                            Next
                                        End If
                                    %>
                                </select>
                            </td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    <% End If%>
  </div>
</form>

<script type="text/javascript">
    $(document).ready(function () {
        updateButtons();
        updateExceptionButtons();
    });
<% If (CloseAfterSave) Then%>
    window.close();
<% End If%>
    removeUsed(false);
    
    
</script>

<%
done:
    MyCommon.Close_LogixRT()
    Send_BodyEnd("mainform", "functioninput")
    MyCommon = Nothing
    Logix = Nothing
%>
