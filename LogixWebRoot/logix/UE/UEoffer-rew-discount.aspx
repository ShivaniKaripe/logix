﻿<%@ Page Language="vb" Debug="true" CodeFile="ue-cb.vb" Inherits="UECB" EnableEventValidation="false" %>

<%@ Import
    Namespace="Copient.CommonIncConfigurable" %>
<%@ Import
    Namespace="System.ServiceModel.Web" %>
<%@ Import
    Namespace="System.Web.Services" %>

<%@ Register Src="~/logix/UserControls/ProductAttributeFilter.ascx" TagName="ProductAttributeFilter"
    TagPrefix="uc1" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="CMS" %>
<%@ Import Namespace="System.Collections" %>
<%@ Import Namespace="CMS.AMS.Contract" %>
<%@ Import Namespace="CMS.AMS" %>
<%@ Import Namespace="System" %>
<%@ Import Namespace="System.Web" %>
<%@ Import Namespace="System.Web.UI" %>
<%@ Import Namespace="System.Web.UI.WebControls" %>
<%@ Import Namespace="CMS.AMS.Models" %>
<%@ Import Namespace="System.Globalization" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="ProductGroup" %>
<%
    ' *****************************************************************************
    ' * FILENAME: UEoffer-rew-discount.aspx 
    ' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ' * Copyright ? 2002 - 2009.  All rights reserved by:
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
%>
<script runat="server">
    ' declaration of page level variables used by inline code and procedures found inside server script tags.
    Dim Localizer As Copient.Localization
    Dim CopientFileName As String
    Dim CopientFileVersion As String = "7.3.1.138972"
    Dim CopientProject As String = "Copient Logix"
    Dim CopientNotes As String = ""
    Dim MyCommon As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    Dim Hierarchy As New Copient.HierarchyResync(MyCommon, "UI", "HierarchyFeeds.txt")
    Dim rst As DataTable
    Dim rst2 As DataTable
    Dim row As DataRow
    Dim OfferID As Long
    Dim DeliverableID As Long
    Dim Name As String = ""
    Dim RewardID As String
    Dim BestDealChecked As String = ""
    Dim AllowNegChecked As String = ""
    Dim DiscountedProductGroupID As Long
    'Dim ExcludedProductGroupID As Long
    Dim AmountTypeID As Long
    Dim L1Cap As New Decimal
    Dim L2DiscountAmt As New Decimal
    Dim L2AmountTypeID As Long
    Dim L2Cap As New Decimal
    Dim L3DiscountAmt As New Decimal
    Dim L3AmountTypeID As Long
    Dim DiscountAmount As New Decimal
    Dim DecliningBalance As Boolean
    Dim RetroactiveDiscount As Boolean
    Dim ChargebackDeptID As Integer = -1
    Dim BestDeal As Integer
    Dim AllowNegative As Integer
    Dim ComputeDiscount As Integer
    Dim ItemLimit As Integer
    Dim WeightLimit As Double
    Dim DollarLimit As Double
    Dim AnyProduct As Boolean
    Dim DiscountBarcode As String
    Dim VoidBarcode As String
    Dim RDesc As String = "" 'Receipt description
    Dim BuyDesc As String = ""
    Dim eDiscountType As Integer  '1=Marsh style, 2=Specified PLU, 3=IBM serial integration, 4=IBM TCP/IP style
    Dim UserGroupID As Long
    Dim DeptLevel As Integer
    Dim DiscountID As Long
    Dim ErrorMsg As String = ""
    Dim Phase As Integer
    Dim CustomerGroupSelOpt As String = ""
    Dim ProductGroupSelOpt As String = ""
    Dim ExcludedPGSelOpt As New StringBuilder()
    Dim DiscountType As Integer
    Dim DiscTypeSel As String = ""
    Dim PgDisabled As String = ""
    Dim PgDblClick As String = ""
    Dim DeslctPgDblClick As String = ""
    Dim TouchPoint As Integer = 0
    Dim TpROID As Integer = 0
    Dim bCreated As Boolean = False
    Dim IsEditable As Boolean = False
    Dim Disallow_Edit As Boolean = True
    Dim FromTemplate As Boolean = False
    Dim IsTemplate As Boolean = False
    Dim IsTemplateVal As String = "Not"
    Dim DisabledAttribute As String = ""
    Dim DisabledAttributeAO As String = ""
    Dim LockFieldsList As String()
    Dim i As Integer
    Dim t As Integer
    Dim OverrideFields As New Hashtable()
    Dim OverrideDiv As New Hashtable()
    Dim OvrdFldEditable As Boolean = False
    Dim OvrdFldDisabled As String = "False"
    Dim OvrdFldClass As String = ""
    Dim CloseAfterSave As Boolean = False
    Dim DeferCalcToEOSChanged As Boolean = False
    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False
    Dim SVProgramID As Integer
    Dim bSVProgram As Boolean
    Dim SelectedStr As String = ""
    Dim EngineID As Integer = 2
    Dim BannersEnabled As Boolean = False
    Dim sQuery As String = ""
    Dim DefaultChrgBack As Integer = -1
    Dim BestDealDefaulted As Boolean = False
    Dim AllBanners As Boolean = False
    Dim TCRMAStatusFlag As Integer = 2
    Dim IsMfgCoupon As Boolean = False
    Dim FlexNeg As Boolean
    Dim flexOption As Integer = 0
    Dim grossPrice As Boolean = False
    Dim FlexNegSO As Integer = 0
    Dim PriceChange As Boolean = False
    Dim GlobalDepts As String = ""
    Dim TierLevels As Integer = 0
    Dim t1, t2 As Decimal
    Dim ValidTier As Boolean = False
    Dim ValidReceiptDesc As Boolean = True
    Dim ValidScorecardDesc As Boolean = True
    Dim WriteTier As Boolean = False
    Dim AllowNegSO As Boolean = False
    Dim ScorecardID As Integer = 0
    Dim ScorecardDesc As String = ""
    Dim DefaultExists As Boolean = False
    Dim DiscountTierID As Integer = 0
    Dim SPRepeatLevel As Integer = 0
    Dim SPLevels As Integer = 0
    Dim SPHighestLevel As Integer = 0
    Dim SPItemLimit As Integer = 0
    Dim LevelID As Integer = 0
    Dim Value As Decimal = 0
    Dim ValueString As String = ""
    Dim ValidLevels As Boolean = False
    Dim ValidLimits As Boolean = False
    Dim ValidPercent As Boolean = True
    Dim ValidAmount As Boolean = False
    Dim ValidDollerLimit As Boolean = False
    Dim ValidDiscPriceLevel As Boolean = True
    Dim ValidSVProgram As Boolean = True
    Dim l As Integer = 0
    Dim x As Integer = 0
    Dim ChargebackSet As Boolean = False
    Dim LoadDefaultChargeback As Boolean = True
    Dim ShowDiscPriceLevel As Boolean = False
    Dim DiscAtOrigPrice As Integer = 0
    Dim ProrationTypeID As Integer = 0
    Dim RewardRequired As Boolean = True
    Dim MLI As New Copient.Localization.MultiLanguageRec
    Dim PhraseLib As IPhraseLib
    Dim SVLib As IStoredValueProgramService
    Dim sbSVOptions As New StringBuilder
    Dim sbSVPointsOptions As New StringBuilder
    Dim rstTemp As DataTable
    Dim HasAnyCustomer As Boolean = False
    Dim bAllowDollarTransLimit As Boolean = False
    Dim excludeBoxSize As Integer = 0
    Dim AttributePGEnabled As Boolean = False
    Dim ProductGroupID As Long = 0
    Dim CreatedDate As String
    Dim LastUpdate As String
    Dim LastUpload As String = Nothing
    Dim LastUploadMsg As String = ""
    Dim XID As String = ""
    Dim IsSpecialGroup As Boolean = False
    Dim ProductGroupTypeID As Byte = 1
    Dim m_ProductGroupService As IProductGroupService
    Dim m_ProductService As IProductService
    Dim BuyerID As Int32 = -1
    Dim m_OfferService As IOffer
    Dim lstAttributePGIDs As List(Of Int64)
    Dim NodeID As String = ""
    Dim ProductGroupName As String = ""
    Dim PABStage As Int16 = 1
    Dim hierarchyHTML As String
    Dim AttributeProductGroupID As Int64
    'Dim locateHierarchyURL As String = ""
    Dim AttributeSwitchType As String = String.Empty
    'Dim bUseMultipleProductExclusionGroups As Boolean = True
    Dim isTranslatedOffer As Boolean = False
    Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = False
    Dim PriceFilterID As Integer = 100
    Dim m_DiscountRewardService As IDiscountRewardService
    Dim excludedProductGroups As List(Of DiscountProductGroup)
    Dim bEnableAdditionalLockoutRestrictionsOnOffers As Boolean = False
    Dim bOfferEditable As Boolean = False
    Dim grossPriceForTender As Boolean
    Dim grossPriceForDiscount As Boolean
    Dim DefaultPageSize As Int32 = 5000
    Dim ValidNodeId As Boolean = True
    'AMS-2223: SystemCacheData to get the cached system options. It is Shared because of use in shared methods for getting cached system options, otherwise like a instance variable loaded in Page_load everytime.     
    Dim SystemCacheData As ICacheData
    Dim SupportGlobalAndTieredConditions As Integer
    Dim UseSameTierValue As Integer = 0
    Dim shouldFetchPGAsync As Boolean
    Dim restrictRewardforRPOS As Boolean = False
    Protected Sub Page_Load(ByVal obj As Object, ByVal e As EventArgs)
        CurrentRequest.Resolver.AppName = "UEoffer-rew-discount.aspx"
        SystemCacheData = CurrentRequest.Resolver.Resolve(Of ICacheData)()
        SupportGlobalAndTieredConditions = SystemCacheData.GetSystemOption_UE_ByOptionId(197) 'SystemCacheData.GetSystemOption_UE_ByOptionId(197)
        shouldFetchPGAsync = SystemCacheData.GetSystemOption_General_ByOptionId(289) 'SystemCacheData.GetSystemOption_General_ByOptionId(289)
        CopientFileName = Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
        If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
            Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
        End If
        Response.Expires = 0
        MyCommon.AppName = "UEoffer-rew-discount.aspx"
        CurrentRequest.Resolver.AppName = MyCommon.AppName
        m_ProductGroupService = CurrentRequest.Resolver.Resolve(Of IProductGroupService)()
        m_ProductService = CurrentRequest.Resolver.Resolve(Of IProductService)()
        m_OfferService = CurrentRequest.Resolver.Resolve(Of IOffer)()
        MyCommon.Open_LogixRT()
        AdminUserID = Verify_AdminUser(MyCommon, Logix)
        Localizer = New Copient.Localization(MyCommon)
        
        '******************BEGIN AJAX CALLS******************************
        If GetCgiValue("mi") <> Nothing Then
            Dim ReturnString As String = ""
            If GetCgiValue("mi") = "PRODUCT_GROUPS" Then
			
			    Dim selected_PGID As Integer = 0
                If GetCgiValue("SelectedPG") <> "" Then
                    Integer.TryParse(GetCgiValue("SelectedPG"), selected_PGID)
                End If
                MyCommon.QueryStr = "SELECT ProductGroups.ProductGroupID,ProductGroups.Buyerid,ProductGroups.ExternalBuyerId, ProductGroups.Name FROM " +
                                            "(SELECT row_number() OVER (ORDER BY Name) AS row_num, " +
                                                "pg.ProductGroupID, pg.Buyerid, b.ExternalBuyerId, pg.Name, pg.deleted, pg.NonDiscountableGroup, pg.PointsNotApplyGroup " +
                                                "FROM ProductGroups pg LEFT JOIN buyers b ON pg.BuyerID = b.BuyerId) AS ProductGroups " +
                                            "WHERE ProductGroups.deleted=0 AND ProductGroups.ProductGroupID <> 1 AND ProductGroups.NonDiscountableGroup=0 AND ProductGroups.PointsNotApplyGroup=0"
                                
                If selected_PGID <> 0 Then
                        MyCommon.QueryStr &= " AND ProductGroups.ProductGroupID <> " & selected_PGID & " "
                End If
				If (MyCommon.IsUserAssociatedWithAnyBuyer(AdminUserID) AndAlso Not Logix.UserRoles.ViewProductgroupRegardlessBuyer) Then
                    MyCommon.QueryStr &= "AND (ProductGroups.BuyerId IN (SELECT BuyerId FROM BuyerRoleUsers WHERE AdminUserID=" & AdminUserID & ") OR BuyerId IS NULL) "
                End If
                                
                MyCommon.QueryStr &= "AND ProductGroups.ProductGroupID BETWEEN " + GetCgiValue("pagestart") + " AND " + GetCgiValue("pageend") + " ORDER BY Name"
                rst = MyCommon.LRT_Select
                
                ConvertDataTabletoJson(ReturnString , rst)
                ReturnString = "T_AMS_SPLITTER_AMS_" + ReturnString
                
            ElseIf GetCgiValue("mi") = "SEARCH_PRODUCT_GROUPS" Then
                Dim selected_PGID As Integer = 0
                If GetCgiValue("SelectedPG") <> "" Then
                    Integer.TryParse(GetCgiValue("SelectedPG"), selected_PGID)
                End If
                If GetCgiValue("SearchType") = "1" Then
                    
                    MyCommon.QueryStr = "SELECT top " + DefaultPageSize.ToString() + " pg.ProductGroupID,pg.Buyerid,b.ExternalBuyerId, pg.Name " +
                                                          "FROM ProductGroups pg LEFT JOIN Buyers b ON pg.BuyerID = b.BuyerId " +
                                                          "WHERE pg.deleted=0 AND pg.ProductGroupID <> 1 AND pg.NonDiscountableGroup=0 AND pg.PointsNotApplyGroup=0 " +
                                                          " AND pg.Name like @Name"
                    
					If selected_PGID <> 0 Then
                        MyCommon.QueryStr &= " AND pg.ProductGroupID <> " & selected_PGID & " "
                    End If
                    If (MyCommon.IsUserAssociatedWithAnyBuyer(AdminUserID) AndAlso Not Logix.UserRoles.ViewProductgroupRegardlessBuyer) Then
                        MyCommon.QueryStr &= "AND (ProductGroups.BuyerId IN (SELECT BuyerId FROM BuyerRoleUsers WHERE AdminUserID=" & AdminUserID & ") OR BuyerId IS NULL) "
                    End If
                                
                    MyCommon.QueryStr &= " ORDER BY Name"
                    MyCommon.DBParameters.Add("@Name", SqlDbType.NVarChar).Value = GetCgiValue("SearchText") + "%"
                    rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
             
                Else
                    
                    MyCommon.QueryStr = "SELECT top " + DefaultPageSize.ToString() + " pg.ProductGroupID,pg.Buyerid,b.ExternalBuyerId, pg.Name " +
                                                          "FROM ProductGroups pg LEFT JOIN Buyers b ON pg.BuyerID = b.BuyerId " +
                                                          "WHERE pg.deleted=0 AND pg.ProductGroupID <> 1 AND pg.NonDiscountableGroup=0 AND pg.PointsNotApplyGroup=0 " +
                                                          " AND pg.Name like @Name"
                    If selected_PGID <> 0 Then
                        MyCommon.QueryStr &= " AND pg.ProductGroupID <> " & selected_PGID & " "
                    End If
                    If (MyCommon.IsUserAssociatedWithAnyBuyer(AdminUserID) AndAlso Not Logix.UserRoles.ViewProductgroupRegardlessBuyer) Then
                        MyCommon.QueryStr &= "AND (ProductGroups.BuyerId IN (SELECT BuyerId FROM BuyerRoleUsers WHERE AdminUserID=" & AdminUserID & ") OR BuyerId IS NULL) "
                    End If
                                
                    MyCommon.QueryStr &= " ORDER BY Name"
                    MyCommon.DBParameters.Add("@Name", SqlDbType.NVarChar).Value = "%" + GetCgiValue("SearchText") + "%"
                    rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)

                End If
                
                
                ConvertDataTabletoJson(ReturnString, rst)
                ReturnString = "T_AMS_SPLITTER_AMS_" + ReturnString
            
            ElseIf GetCgiValue("mi") = "PRODUCT_GROUPS_INQUIRY" Then
                bEnableRestrictedAccessToUEOfferBuilder = IIf(MyCommon.Fetch_SystemOption(249) = "1", True, False)
                MyCommon.QueryStr = "dbo.pt_ProductGroupInquiryList"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@checkPGlistBasedOnBuyerID", SqlDbType.Bit).Value = (Logix.UserRoles.ViewProductgroupRegardlessBuyer = False)
                MyCommon.LRTsp.Parameters.Add("@adminuserid", SqlDbType.Int).Value = AdminUserID
                MyCommon.LRTsp.Parameters.Add("@bEnableRestrictedAccessToUEOfferBuilder", SqlDbType.Bit).Value = bEnableRestrictedAccessToUEOfferBuilder
                MyCommon.LRTsp.Parameters.Add("@pageStart", SqlDbType.Int).Value = GetCgiValue("pagestart")
                MyCommon.LRTsp.Parameters.Add("@pageEnd", SqlDbType.Int).Value = GetCgiValue("pageend")
                MyCommon.LRTsp.Parameters.Add("@searchText", SqlDbType.NVarChar, 50).Value = GetCgiValue("SearchText")
                MyCommon.LRTsp.Parameters.Add("@searchType", SqlDbType.Int).Value = IIf(GetCgiValue("SearchType") = "", 0, GetCgiValue("SearchType"))
                
                rst = MyCommon.LRTsp_select()
                MyCommon.Close_LRTsp()
                                
                ConvertDataTabletoJson(ReturnString, rst)
                ReturnString = "T_AMS_SPLITTER_AMS_" + ReturnString
            End If
            
            Response.Write(ReturnString)
            Response.End()
            Return
        End If
        '**********************END AJAX CALLS******************************
        
        PhraseLib = CurrentRequest.Resolver.Resolve(Of IPhraseLib)()
        SVLib = CurrentRequest.Resolver.Resolve(Of IStoredValueProgramService)()
        m_DiscountRewardService = CurrentRequest.Resolver.Resolve(Of IDiscountRewardService)()

        NodeID = GetCgiValue("NodeListID")
        Dim objResult As AMSResult(Of List(Of Int64)) = m_ProductGroupService.GetAllAttributeBasedProductGroupIDs()
        If (objResult.ResultType <> AMSResultType.Success) Then
            infoMessage = objResult.MessageString
        Else
            lstAttributePGIDs = objResult.Result
        End If
        BannersEnabled = (SystemCacheData.GetSystemOption_General_ByOptionId(66) = "1")
        BestDealDefaulted = (SystemCacheData.GetSystemOption_UE_ByOptionId(17) = "1")
        'Get the flex negative system option to use as default
        FlexNegSO = SystemCacheData.GetSystemOption_UE_ByOptionId(15)
        AllowNegSO = (SystemCacheData.GetSystemOption_UE_ByOptionId(18) = "1")
        ShowDiscPriceLevel = (SystemCacheData.GetSystemOption_UE_ByOptionId(124) = "1")
        bAllowDollarTransLimit = (SystemCacheData.GetSystemOption_UE_ByOptionId(187) = "1")
        grossPriceForDiscount = (SystemCacheData.GetSystemOption_UE_ByOptionId(209) = "1")
        grossPriceForTender = (SystemCacheData.GetSystemOption_UE_ByOptionId(230) = "1")
        If GetCgiValue("discountType") <> "" Then
            DiscountType = MyCommon.Extract_Decimal(GetCgiValue("discountType"), MyCommon.GetAdminUser.Culture)
        Else
            DiscountType = 0
        End If
        OfferID = GetCgiValue("OfferID")
        restrictRewardforRPOS = (SystemCacheData.GetSystemOption_UE_ByOptionId(234) = "1")
        'This would redirect to summary page if offer is awaiting deployment or if offer is awaiting recommendations.
        CheckIfValidOffer(MyCommon, OfferID)

        RewardID = GetCgiValue("RewardID")
        DiscountID = GetCgiValue("DiscountID")
        DeliverableID = MyCommon.Extract_Decimal(GetCgiValue("DeliverableID"), MyCommon.GetAdminUser.Culture)
        AmountTypeID = MyCommon.Extract_Decimal(GetCgiValue("l1amounttypeid"), MyCommon.GetAdminUser.Culture)

        isTranslatedOffer = MyCommon.IsTranslatedUEOffer(OfferID, MyCommon)
        bEnableRestrictedAccessToUEOfferBuilder = IIf(SystemCacheData.GetSystemOption_General_ByOptionId(249) = "1", True, False)

        bEnableAdditionalLockoutRestrictionsOnOffers = IIf(SystemCacheData.GetSystemOption_General_ByOptionId(260) = "1", True, False)
        bOfferEditable = MyCommon.IsOfferEditablePastLockOutPeriod(Logix.UserRoles.EditOfferPastLockoutPeriod, MyCommon, OfferID)

        If GetCgiValue("loadDefaultChargeback") <> "" Then
            LoadDefaultChargeback = (GetCgiValue("loadDefaultChargeback") = "1")
        End If
        If Not String.IsNullOrWhiteSpace(GetCgiValue("AttributeSwitchType")) Then AttributeSwitchType = GetCgiValue("AttributeSwitchType")
        If (GetCgiValue("EngineID") <> "") Then
            EngineID = MyCommon.Extract_Decimal(GetCgiValue("EngineID"), MyCommon.GetAdminUser.Culture)
        Else
            MyCommon.QueryStr = "select EngineID from OfferIDs where OfferID=" & OfferID & ";"
            rst = MyCommon.LRT_Select
            If rst.Rows.Count > 0 Then
                EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), 0)
            End If
        End If
        If (GetCgiValue("flexnegative") <> Nothing AndAlso GetCgiValue("flexnegative") <> "") Then
            FlexNeg = IIf((GetCgiValue("flexnegative") = "1" OrElse GetCgiValue("flexnegative") = "2"), True, False)
            flexOption = GetCgiValue("flexnegative")
        Else
            FlexNeg = IIf(FlexNegSO = 1, True, False)
            flexOption = FlexNegSO
        End If
        'AMS-685 get the excluded groups
        If DiscountID <> 0 Then
            Dim amsResult As AMSResult(Of List(Of DiscountProductGroup))
            amsResult = m_DiscountRewardService.GetAllExclusionGroups(DiscountID)
            If amsResult.ResultType = AMSResultType.Success Then
                excludedProductGroups = amsResult.Result
            Else
                infoMessage = amsResult.PhraseString
            End If
        End If
        If IsPostBack Then
            UpdateCurrentListOfProductGroups(excludedProductGroups, GetCgiValue("excludedpgid"))
        End If
        AttributePGEnabled = (SystemCacheData.GetSystemOption_UE_ByOptionId(157) = "1") AndAlso (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE))
        If (AttributePGEnabled AndAlso ProductGroupID = 0) Then
            Dim DefaultProductGroupID As String = SystemCacheData.GetSystemOption_UE_ByOptionId(156)
            ProductGroupTypeID = DefaultProductGroupID
            If Not Page.IsPostBack Then
                Dim ProductGroupTypes As AMSResult(Of List(Of ProductGroupTypes)) = m_ProductGroupService.GetProductGroupTypes()
                If (ProductGroupTypes.ResultType <> AMSResultType.Success) Then
                    infoMessage = ProductGroupTypes.MessageString
                Else
                    RadioButtonList1.DataSource = ProductGroupTypes.Result
                    RadioButtonList1.DataTextField = "Name"
                    RadioButtonList1.DataValueField = "ProductGroupTypeID"
                    RadioButtonList1.DataBind()
                    RadioButtonList1.Attributes.Add("onclick", "javascript:ProductGroupDivSelection();ShoworHideDivsWithoutPostback();")
                    If (RadioButtonList1.Items.Count > 0) Then
                        If RadioButtonList1.Items(0) IsNot Nothing Then
                            RadioButtonList1.Items(0).Text = Copient.PhraseLib.Lookup("term.selectexistingprodgroups", LanguageID)
                        End If
                        If RadioButtonList1.Items(1) IsNot Nothing Then
                            RadioButtonList1.Items(1).Text = Copient.PhraseLib.Lookup("term.useattributeprodgroups", LanguageID)
                        End If
                        RadioButtonList1.ClearSelection()
                        Dim radioBtn As ListItem = RadioButtonList1.Items.FindByValue(DefaultProductGroupID)
                        If radioBtn Is Nothing Then
                            RadioButtonList1.Items(0).Selected = True
                        Else
                            radioBtn.Selected = True
                        End If
                    End If
                End If
            Else
                ProductGroupTypeID = IIf(AttributePGEnabled = True, RadioButtonList1.SelectedItem.Value.ConvertToByte(), 0)
            End If
        Else
            RadioButtonList1.Visible = False
        End If
        'If Not string.IsNullOrWhiteSpace(Request.QueryString("LocateHierarchyURL")) Then
        '   locateHierarchyURL = HttpUtility.UrlDecode(Request.QueryString("LocateHierarchyURL"))             
        'End If                            
        'Get product group name to display
        If AttributePGEnabled Then
            ProductGroupName = GetProductGroupName()
            txtProductGroupName.Value = ProductGroupName
        End If

        'Get the tier levels
        MyCommon.QueryStr = "select TierLevels from CPE_RewardOptions where RewardOptionID=" & RewardID & ";"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
            TierLevels = MyCommon.NZ(rst.Rows(0).Item("TierLevels"), 0)
        End If

        'Get UseSameTierValue
        If TierLevels > 1 Then
            MyCommon.QueryStr = "select TierLevel, DiscountAmount from CPE_DiscountTiers where DiscountID=" & DiscountID & ";"
            rst = MyCommon.LRT_Select
            If rst.Rows.Count > 0 Then
                Dim lastAmount As Decimal = 0.0
                For Each row In rst.Rows
                    If (MyCommon.NZ(row.Item("TierLevel"), 0) > 1) And MyCommon.NZ(row.Item("DiscountAmount"), 0.0) <> lastAmount Then
                        UseSameTierValue = 0
                        Exit For
                    Else
                        lastAmount = MyCommon.NZ(row.Item("DiscountAmount"), 0.0)
                    End If
                Next
                If MyCommon.NZ(row.Item("TierLevel"), 0) = TierLevels Then
                    UseSameTierValue = 1
                End If
            End If
        Else
            UseSameTierValue = 0
        End If

        Phase = MyCommon.Extract_Decimal(GetCgiValue("phase"), MyCommon.GetAdminUser.Culture)
        If (Phase = 0) Then Phase = MyCommon.Extract_Decimal(Request.Form("Phase"), MyCommon.GetAdminUser.Culture)
        If (Phase = 0) Then Phase = 3

        TouchPoint = MyCommon.Extract_Decimal(GetCgiValue("tp"), MyCommon.GetAdminUser.Culture)
        If (TouchPoint > 0) Then TpROID = MyCommon.Extract_Decimal(GetCgiValue("roid"), MyCommon.GetAdminUser.Culture)

        ' Fetch the name and other details
        MyCommon.QueryStr = "Select IncentiveName, IsTemplate, FromTemplate, ManufacturerCoupon from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
            Name = MyCommon.NZ(rst.Rows(0).Item("IncentiveName"), "")
            IsTemplate = MyCommon.NZ(rst.Rows(0).Item("IsTemplate"), False)
            FromTemplate = MyCommon.NZ(rst.Rows(0).Item("FromTemplate"), False)
            IsMfgCoupon = MyCommon.NZ(rst.Rows(0).Item("ManufacturerCoupon"), False)
        End If
        IsTemplateVal = IIf(IsTemplate, "IsTemplate", "Not")
        Dim m_EditOfferRegardlessOfBuyer = Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID)
        If Not IsTemplate Then
            IsEditable = Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer
        Else
            IsEditable = Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer
        End If

        If (OfferID > 0 AndAlso (GetCgiValue("save") <> "")) Then
            'Tier level validation code
            If TierLevels > 1 AndAlso AmountTypeID <> 4 AndAlso AmountTypeID <> 8 Then
                'Run validation for everything except "free" and "special pricing"
                If Copient.commonShared.Contains(AmountTypeID, 2, 6, 13, 14, 15, 16) Then
                    ValidTier = True
                    WriteTier = True
                Else
                    'Fixed amount off, fixed percentage off, or stored value
                    For t = 2 To TierLevels
                        t2 = MyCommon.Extract_Decimal(GetCgiValue("tier" & t & "_l1discountamt"), MyCommon.GetAdminUser.Culture)
                        t1 = MyCommon.Extract_Decimal(GetCgiValue("tier" & t - 1 & "_l1discountamt"), MyCommon.GetAdminUser.Culture)
                        If (t2 > t1) OrElse (t1 = 0 AndAlso t2 = 0) Then
                            ValidTier = True
                            WriteTier = True
                        Else
                            ValidTier = False
                            WriteTier = False
                            Exit For
                        End If
                    Next
                End If
            Else
                ValidTier = True
                WriteTier = True
            End If

            ' validate that the discount amounts don't exceed the database maximum
            ValidAmount = True
            If AmountTypeID <> 4 AndAlso AmountTypeID <> 8 Then
                For t = 1 To TierLevels
                    t1 = MyCommon.Extract_Decimal(GetCgiValue("tier" & t & "_l1discountamt"), MyCommon.GetAdminUser.Culture)
                    If t1 <= 0 OrElse t1 >= 1000000000 Then
                        ValidAmount = False
                        Exit For
                    End If
                Next
            End If
            'validate that the doller limit don't exceed the database maximum
            ValidDollerLimit = True
            If (Copient.commonShared.Contains(DiscountType, 1, 4) AndAlso AmountTypeID <> 8) Then
                For t = 1 To TierLevels
                    t1 = MyCommon.Extract_Decimal(GetCgiValue("tier" & t & "_dollarlimit"), MyCommon.GetAdminUser.Culture)
                    If t1 < 0 OrElse t1 >= 1000000000 Then
                        ValidDollerLimit = False
                        Exit For
                    End If
                Next
            End If
            ValidSVProgram = True
            If CMS.Utilities.Extract_Val(GetCgiValue("discountsv")) = 1 AndAlso CMS.Utilities.Extract_Val(GetCgiValue("ddldiscountsv")) = 0 Then
                ValidSVProgram = False
                infoMessage = PhraseLib.Lookup("reward.discount-svprogram", LanguageID)
            End If

            ' validate that percentage off does not exceed 100%
            ValidPercent = True
            If AmountTypeID = 3 Then
                For t = 1 To TierLevels
                    t1 = MyCommon.Extract_Decimal(GetCgiValue("tier" & t & "_l1discountamt"), MyCommon.GetAdminUser.Culture)
                    If t1 <= 0 OrElse t1 > 100 Then
                        ValidPercent = False
                        Exit For
                    End If
                Next
            End If

            'Added to Check Product Node Selected.
            ValidNodeId = True
            If NodeID = "" Then
                If GetButtonListSelectedValue() = 2 Then
                    If GetCgiValue("discountType") <> "" AndAlso hdnSwitchPGID.Value.ConvertToLong() = 0 Then
                        Dim Type = MyCommon.Extract_Decimal(GetCgiValue("discountType"), MyCommon.GetAdminUser.Culture)
                        If Copient.commonShared.Contains(Type, 1, 2) Then
                            If Not ucProductAttributeFilter.PrevSelectedNodeIDs Is Nothing Then
                                If ucProductAttributeFilter.GetSelectedNodeIDs() = "" Then
                                    ValidNodeId = False
                                End If
                            Else
                                ValidNodeId = False
                            End If
                        End If
                    End If
                End If
            End If
            If MyCommon.Extract_Decimal(GetCgiValue("l1amounttypeid"), MyCommon.GetAdminUser.Culture) = 8 Then
                ' Special pricing validation: all levels must be numeric, and the number of levels must be <= the item limit
                For t = 1 To TierLevels
                    SPLevels = MyCommon.Extract_Decimal(GetCgiValue("tier" & t & "_levels"), MyCommon.GetAdminUser.Culture)
                    SPHighestLevel = MyCommon.Extract_Decimal(GetCgiValue("tier" & t & "_highestlevel"), MyCommon.GetAdminUser.Culture)
                    SPItemLimit = MyCommon.Extract_Decimal(GetCgiValue("tier" & t & "_itemlimit"), MyCommon.GetAdminUser.Culture)
                    If (SPLevels > SPItemLimit) AndAlso (SPItemLimit > 0) Then
                        ValidLevels = False
                        infoMessage = Copient.PhraseLib.Lookup("error.splevelqty", LanguageID)
                        Exit For
                    Else
                        If (SPHighestLevel = 0 OrElse SPLevels = 0) Then
                            ValidLevels = False
                            Exit For
                        End If
                        For l = 1 To SPHighestLevel
                            If GetCgiValue("tier" & t & "_level" & l) <> Nothing Then
                                If IsNumeric(GetCgiValue("tier" & t & "_level" & l)) Then
                                    ValidLevels = True
                                Else
                                    ValidLevels = False
                                    infoMessage = Copient.PhraseLib.Lookup("error.splevelvalues", LanguageID)
                                    Exit For
                                End If
                            Else
                                ValidLevels = True
                            End If
                        Next
                        If ValidLevels = False Then
                            Exit For
                        End If
                    End If
                Next
            Else
                ValidLevels = True
            End If

            If Copient.commonShared.Contains(DiscountType, 1, 5) AndAlso Not Copient.commonShared.Contains(AmountTypeID, 5, 6, 9, 10, 11, 12, 13, 14, 15, 16) Then
                For t = 1 To TierLevels
                    If (GetCgiValue("tier" & t & "_itemlimit") <> "") Then
                        hdnTemp.Value = GetCgiValue("tier" & t & "_itemlimit")
                        If Integer.TryParse(GetCgiValue("tier" & t & "_itemlimit"), Nothing) AndAlso (GetCgiValue("tier" & t & "_itemlimit") >= 0) Then
                            ValidLimits = True
                        Else
                            ValidLimits = False
                            infoMessage = Copient.PhraseLib.Detokenize("ueoffer-rew-discount.InvalidItemLimit", LanguageID, GetCgiValue("tier" & t & "_itemlimit"))
                            Exit For
                        End If
                    Else
                        If hdnTemp.Value <> Nothing Then
                            ValidLimits = True
                        End If
                    End If
                Next
            Else
                ValidLimits = True
            End If

            If GetCgiValue("discountedpgid") <> "" Then
                DiscountedProductGroupID = MyCommon.Extract_Decimal(GetCgiValue("discountedpgid"), MyCommon.GetAdminUser.Culture)
            Else
                DiscountedProductGroupID = 0
            End If

            If GetCgiValue("save") <> "" Then
                If DiscountedProductGroupID = 0 Then
                    DiscountedProductGroupID = -1
                End If
            End If

            'Validate the ScoreCardCard Desc
            ValidScorecardDesc = True
            If (GetCgiValue("ScorecardID") <> 0) Then
                If GetCgiValue("ScorecardDesc") Is Nothing OrElse GetCgiValue("ScorecardDesc").Trim() = "" Then
                    ValidScorecardDesc = False
                    infoMessage = Copient.PhraseLib.Lookup("reward.discount-scorecarddescription", LanguageID)
                End If
            End If

            ValidReceiptDesc = True
            If AmountTypeID <> 7 Then
                For t = 1 To TierLevels
                    If GetCgiValue("tier" & t & "_rdesc") Is Nothing OrElse GetCgiValue("tier" & t & "_rdesc") = "" Then
                        ValidReceiptDesc = False
                        infoMessage = Copient.PhraseLib.Lookup("reward.discount-receiptdescription", LanguageID)
                    End If
                Next
            End If

            RewardRequired = IIf(MyCommon.Extract_Val(GetCgiValue("requiredToDeliver")) = 1, True, False)

            ValidDiscPriceLevel = IsValidDiscPriceLevel(MyCommon, infoMessage)

            If ValidTier AndAlso ValidPercent AndAlso ValidLevels AndAlso ValidLimits AndAlso ValidSVProgram AndAlso ValidReceiptDesc AndAlso ValidDiscPriceLevel AndAlso ValidAmount AndAlso ValidDollerLimit AndAlso (DiscountedProductGroupID > -1 OrElse DiscountType = 4 OrElse DiscountType = 5) AndAlso ValidNodeId AndAlso ValidScorecardDesc Then
                If (DiscountID <= 0) Then
                    DiscountID = Create_Discount(OfferID, TpROID, Phase, DeliverableID, RewardRequired)
                    bCreated = True
                    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Reward.creatediscount", LanguageID))
                End If
                'AMS-685 commented ExcludedProductGroupID
                ' Store the product group and excluded product group for later possible TCRM comparison
                MyCommon.QueryStr = "select DiscountedProductGroupID, ExcludedProductGroupID from CPE_Discounts with (NoLock) where DiscountID=" & DiscountID
                rst = MyCommon.LRT_Select
                If (rst.Rows.Count > 0) Then
                    DiscountedProductGroupID = MyCommon.NZ(rst.Rows(0).Item("DiscountedProductGroupID"), -1)
                    'ExcludedProductGroupID = MyCommon.NZ(rst.Rows(0).Item("ExcludedProductGroupID"), -1)
                End If

                ' Save the contents of the discount
                If ValidLevels Then
                    ' DISCOUNT SAVES HERE *~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*
                    Save_Discount(OfferID, DiscountID, TierLevels, WriteTier, bCreated, RewardID, bAllowDollarTransLimit)
                    ' *~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*
                    If (Not bCreated) Then
                        ' Update for TCRM
                        ' Determine if the product group has changed; if so, flag s/b 3
                        'AMS-685 replaced comparison of single exclusion group with multiple exclusion groups
                        If (DiscountedProductGroupID <> MyCommon.Extract_Decimal(GetCgiValue("discountedpgid"), MyCommon.GetAdminUser.Culture) _
                           OrElse ExcludedPGChanged(excludedProductGroups, GetCgiValue("excludedpgid"))) Then
                            TCRMAStatusFlag = 3
                        Else
                            TCRMAStatusFlag = 2
                        End If
                        ' If TCRMAStatusFlag is already 3 then don't change it to 2.
                        MyCommon.QueryStr = "update CPE_Deliverables with (RowLock) set TCRMAStatusFlag=" & TCRMAStatusFlag & " " & _
                                            "where TCRMAStatusFlag <> 3 and DeliverableID=" & DeliverableID
                        MyCommon.LRT_Execute()

                        ' update the require reward flag
                        MyCommon.QueryStr = "update CPE_Deliverables with (RowLock) set Required= " & IIf(RewardRequired, 1, 0) & " " & _
                                            "where DeliverableID=" & DeliverableID
                        MyCommon.LRT_Execute()

                        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Reward.editdiscount", LanguageID))
                    End If

                    If (GetCgiValue("save") <> "") Then
                        CloseAfterSave = (SystemCacheData.GetSystemOption_General_ByOptionId(48) = "1")
                    End If
                End If
            Else
                If GetButtonListSelectedValue() <> 2 Then
                    If infoMessage <> "" Then
                        'infoMessage already set
                    ElseIf DiscountedProductGroupID = -1 AndAlso DiscountType <> 4 AndAlso DiscountType <> 5 Then
                        infoMessage = Copient.PhraseLib.Lookup("reward.groupselect", LanguageID)
                    ElseIf Not ValidTier Then
                        infoMessage = Copient.PhraseLib.Lookup("error.tiervalues", LanguageID)
                    ElseIf Not ValidPercent Then
                        infoMessage = Copient.PhraseLib.Lookup("error.discount-percent-over", LanguageID)
                    ElseIf Not ValidAmount Then
                        infoMessage = Copient.PhraseLib.Lookup("error.invalid-amount-billion", LanguageID)
                    ElseIf Not ValidDollerLimit Then
                        infoMessage = Copient.PhraseLib.Lookup("error.invalid-amount-billion", LanguageID)
                    ElseIf Copient.commonShared.Contains(AmountTypeID, 2, 6, 13, 14, 15, 16) Then
                        infoMessage = Copient.PhraseLib.Lookup("error.tiervaluesdecrease", LanguageID)
                    ElseIf AmountTypeID = 8 AndAlso Not ValidLevels Then
                        If TierLevels = 1 Then
                            infoMessage = Copient.PhraseLib.Lookup("error.nopricepoint", LanguageID)
                        ElseIf TierLevels > 1 Then
                            infoMessage = Copient.PhraseLib.Lookup("error.tiersnopricepoint", LanguageID)
                        End If
                    Else
                        infoMessage = Copient.PhraseLib.Lookup("error.general", LanguageID)
                    End If
                Else
                    If Not ValidTier Then
                        infoMessage = Copient.PhraseLib.Lookup("error.tiervalues", LanguageID)
                    ElseIf Not ValidPercent Then
                        infoMessage = Copient.PhraseLib.Lookup("error.discount-percent-over", LanguageID)
                    ElseIf Not ValidAmount Then
                        infoMessage = Copient.PhraseLib.Lookup("error.invalid-amount-billion", LanguageID)
                    ElseIf Not ValidDollerLimit Then
                        infoMessage = Copient.PhraseLib.Lookup("error.invalid-amount-billion", LanguageID)
                    ElseIf Not ValidNodeId Then
                        infoMessage = Copient.PhraseLib.Lookup("error.noselectednode", LanguageID)
                    ElseIf Copient.commonShared.Contains(AmountTypeID, 6, 13, 14, 15, 16) Then
                        infoMessage = Copient.PhraseLib.Lookup("error.tiervaluesdecrease", LanguageID)
                    ElseIf AmountTypeID = 8 AndAlso Not ValidLevels Then
                        If TierLevels = 1 Then
                            infoMessage = Copient.PhraseLib.Lookup("error.nopricepoint", LanguageID)
                        ElseIf TierLevels > 1 Then
                            infoMessage = Copient.PhraseLib.Lookup("error.tiersnopricepoint", LanguageID)
                        End If
                    End If
                End If
                If (IsTemplate) Then
                    ' time to update the status bits for the templates
                    'clear the template field exception permissions
                    Dim form_Disallow_Edit As Integer = 0
                    If (GetCgiValue("Disallow_Edit") = "on") Then
                        form_Disallow_Edit = 1
                    End If
                    MyCommon.QueryStr = "delete from TemplateFieldPermissions with (RowLock) where OfferID=" & OfferID & " " & _
                                        "and FieldID in (select FieldID from UIFields where PageName='" & MyCommon.AppName & "');"
                    MyCommon.LRT_Execute()
                    If (GetCgiValue("chkTempField") <> "") Then
                        Dim tmpFldLen As Integer = GetCgiValue("chkTempField").Length
                        If (tmpFldLen > 0) Then
                            ReDim LockFieldsList(tmpFldLen)
                            LockFieldsList = Request.Form.GetValues("chkTempField")
                            For i = 0 To LockFieldsList.Length - 1
                                MyCommon.QueryStr = "insert into TemplateFieldPermissions with (RowLock) (OfferID, FieldID,DeliverableID, Editable) " & _
                                                    "values (" & OfferID & ", " & LockFieldsList(i) & "," & DeliverableID & "," & form_Disallow_Edit & ");"
                                MyCommon.LRT_Execute()
                            Next
                        End If
                    End If
                End If
            End If
        End If
        ' Discount saving ended
        ' Loading of product group
        AnyProduct = False
        UserGroupID = 0
        If DiscountID = 0 Then
            DiscountID = MyCommon.Extract_Decimal(GetCgiValue("DiscountID"), MyCommon.GetAdminUser.Culture)
        End If

        eDiscountType = 4
        'Loading from database        
        MyCommon.QueryStr = "select DiscountTypeID from CPE_Discounts where DiscountID =" & DiscountID
        rst = MyCommon.LRT_Select
        If Not (rst.Rows.Count = 0) Then
            DiscountType = MyCommon.NZ(rst.Rows(0).Item("DiscountTypeID"), 0)
        End If
        If (DiscountType = 4 Or DiscountType = 5) Then
            MyCommon.QueryStr = "select Name, DiscountTypeID, ReceiptDescription, SpecifyBarcode, DiscountBarcode, VoidBarcode, DiscountedProductGroupID, " &
                              "ExcludedProductGroupID, BestDeal, AllowNegative, ComputeDiscount, DiscountAmount, AmountTypeID, " &
                              "L1Cap, L2DiscountAmt, L2AmountTypeID, L2Cap, L3DiscountAmt, L3AmountTypeID, ItemLimit, WeightLimit, DollarLimit, ChargebackDeptID, " &
                              "DecliningBalance, RetroactiveDiscount, UserGroupID, LastUpdate, SVProgramID, FlexNegative, ScorecardID, ScorecardDesc, DiscountAtOrigPrice, " &
                              "ProrationTypeID, PriceChange, PriceFilter, FlexOptions, GrossPrice " &
                              "from CPE_Discounts with (NoLock) where Deleted=0 and DiscountID=" & DiscountID & ";"
        Else
            MyCommon.QueryStr = "SELECT CPED.Name, DiscountTypeID, ReceiptDescription, SpecifyBarcode, DiscountBarcode, VoidBarcode, DiscountedProductGroupID," &
                                "ExcludedProductGroupID, BestDeal, AllowNegative, ComputeDiscount, DiscountAmount, AmountTypeID, L1Cap, L2DiscountAmt, " &
                                "L2AmountTypeID, L2Cap, L3DiscountAmt, L3AmountTypeID, ItemLimit, WeightLimit, DollarLimit, ChargebackDeptID, DecliningBalance," &
                                "RetroactiveDiscount, UserGroupID, CPED.LastUpdate, SVProgramID, FlexNegative, ScorecardID, ScorecardDesc, DiscountAtOrigPrice," &
                                "ProrationTypeID, PriceChange, PG.ProductGroupTypeID,CPED.PriceFilter, FlexOptions, GrossPrice FROM CPE_Discounts CPED WITH (NOLOCK) INNER JOIN " &
                                "ProductGroups PG WITH (NOLOCK) ON CPED.DiscountedProductGroupID = PG.ProductGroupID WHERE  CPED.Deleted=0 and DiscountID=" & DiscountID
        End If
        rst = MyCommon.LRT_Select
        If Not (rst.Rows.Count = 0) Then
            AmountTypeID = MyCommon.NZ(rst.Rows(0).Item("AmountTypeID"), 0)

            Name = MyCommon.NZ(rst.Rows(0).Item("Name"), "")
            DiscountType = MyCommon.NZ(rst.Rows(0).Item("DiscountTypeID"), 0)
            DeptLevel = IIf(DiscountType = 2, 1, 0)
            RDesc = MyCommon.NZ(rst.Rows(0).Item("ReceiptDescription"), "")
            DiscountBarcode = MyCommon.NZ(rst.Rows(0).Item("DiscountBarcode"), "")
            VoidBarcode = MyCommon.NZ(rst.Rows(0).Item("VoidBarcode"), "")
            If (hidDiscountType.Value = "3") Then
                DiscountedProductGroupID = 1
                discountedpgid.Value = DiscountedProductGroupID
            Else
                DiscountedProductGroupID = MyCommon.NZ(rst.Rows(0).Item("DiscountedProductGroupID"), 0)
                discountedpgid.Value = DiscountedProductGroupID
            End If
            'AMS-685 commented out 
            'ExcludedProductGroupID = MyCommon.NZ(rst.Rows(0).Item("ExcludedProductGroupID"), 0)
            If (DiscountType = 1 Or DiscountType = 2) Then
                ProductGroupTypeID = MyCommon.NZ(rst.Rows(0).Item("ProductGroupTypeID"), 1)
                If Not IsPostBack Then
                    If ProductGroupTypeID = 1 AndAlso AttributePGEnabled Then
                        RadioButtonList1.ClearSelection()
                        RadioButtonList1.Items(0).Selected = True
                    ElseIf ProductGroupTypeID = 2 AndAlso AttributePGEnabled Then
                        RadioButtonList1.ClearSelection()
                        RadioButtonList1.Items(1).Selected = True
                        hdnIsAttributeSwitch.Value = 1
                        hdnSwitchPGID.Value = discountedpgid.Value.ConvertToLong()
                    End If
                End If
            End If
            BestDeal = MyCommon.NZ(rst.Rows(0).Item("BestDeal"), 0)
            If BestDeal <> 0 Then
                BestDeal = 1
            End If
            AllowNegative = MyCommon.NZ(rst.Rows(0).Item("AllowNegative"), 0)
            If AllowNegative <> 0 Then
                AllowNegative = 1
            End If
            ComputeDiscount = MyCommon.NZ(rst.Rows(0).Item("ComputeDiscount"), 0)
            If ComputeDiscount <> 0 Then
                ComputeDiscount = 1
            End If
            DiscAtOrigPrice = MyCommon.Extract_Decimal(MyCommon.NZ(rst.Rows(0).Item("DiscountAtOrigPrice"), 0), MyCommon.GetAdminUser.Culture)
            DiscountAmount = Localizer.Round_Currency(MyCommon.NZ(rst.Rows(0).Item("DiscountAmount"), 0), RewardID, (AmountTypeID = 3))
            L1Cap = MyCommon.NZ(rst.Rows(0).Item("L1Cap"), 0)
            L2DiscountAmt = Localizer.Round_Currency(MyCommon.NZ(rst.Rows(0).Item("L2DiscountAmt"), 0), RewardID)
            L2AmountTypeID = MyCommon.NZ(rst.Rows(0).Item("L2AmountTypeID"), 0)
            L2Cap = MyCommon.NZ(rst.Rows(0).Item("L2Cap"), 0)
            L3DiscountAmt = Localizer.Round_Currency(MyCommon.NZ(rst.Rows(0).Item("L3DiscountAmt"), 0), RewardID)
            L3AmountTypeID = MyCommon.NZ(rst.Rows(0).Item("L3AmountTypeID"), 0)
            ItemLimit = MyCommon.NZ(rst.Rows(0).Item("ItemLimit"), 1)
            WeightLimit = Localizer.Round_Quantity(MyCommon.NZ(rst.Rows(0).Item("WeightLimit"), 0), RewardID, 5)
            DollarLimit = Localizer.Round_Currency(MyCommon.NZ(rst.Rows(0).Item("DollarLimit"), 1), RewardID)
            ChargebackDeptID = MyCommon.NZ(rst.Rows(0).Item("ChargebackDeptID"), 0)
            ChargebackSet = True
            LoadDefaultChargeback = False

            DecliningBalance = MyCommon.NZ(rst.Rows(0).Item("DecliningBalance"), 0)
            RetroactiveDiscount = MyCommon.NZ(rst.Rows(0).Item("RetroactiveDiscount"), False)
            UserGroupID = MyCommon.NZ(rst.Rows(0).Item("UserGroupID"), 0)
            SVProgramID = MyCommon.NZ(rst.Rows(0).Item("SVProgramID"), 0)
            PriceFilterID = MyCommon.NZ(rst.Rows(0).Item("PriceFilter"), 100)
            '' Check whether this SVProgramID exist in StoredValuePrograms table
            MyCommon.QueryStr = "select SVProgramID from StoredValuePrograms with (NoLock) where SVProgramID =" & SVProgramID
            rstTemp = MyCommon.LRT_Select
            If (rstTemp.Rows.Count <= 0) Then
                SVProgramID = 0
            End If
            If SVProgramID > 0 Then
                bSVProgram = True
            Else
                bSVProgram = False
            End If

            FlexNeg = MyCommon.NZ(rst.Rows(0).Item("FlexNegative"), False)
            grossPrice = MyCommon.NZ(rst.Rows(0).Item("GrossPrice"), False)
            flexOption = MyCommon.NZ(rst.Rows(0).Item("FlexOptions"), 0)
            PriceChange = MyCommon.NZ(rst.Rows(0).Item("PriceChange"), 0)
            ScorecardID = MyCommon.NZ(rst.Rows(0).Item("ScorecardID"), 0)
            ScorecardDesc = MyCommon.NZ(rst.Rows(0).Item("ScorecardDesc"), "")
            ProrationTypeID = MyCommon.NZ(rst.Rows(0).Item("ProrationTypeID"), 0)

        Else
            'delete template permissions if any existing 
            If Not (GetCgiValue("save") <> "") Then
                MyCommon.QueryStr = "delete from TemplateFieldPermissions with (RowLock) where OfferID=" & OfferID & " " & _
                       "and FieldID in (select FieldID from UIFields where PageName='" & MyCommon.AppName & "');"
                MyCommon.LRT_Execute()
            End If

        End If

        Dim Offer As Offer = m_OfferService.GetOffer(OfferID, LoadOfferOptions.None)
        If (Offer IsNot Nothing) Then
            If (AttributePGEnabled AndAlso Offer.BuyerID IsNot Nothing) Then
                ucProductAttributeFilter.BuyerID = Offer.BuyerID
            End If
            Name = Offer.OfferName
            IsTemplate = Offer.IsTemplate
            FromTemplate = Offer.FromTemplate
        End If

        If (AttributePGEnabled) Then
            ucProductAttributeFilter.IsAttributeSwitch = False
            If AttributeSwitchType = "SelectedAttributeGroup" AndAlso String.IsNullOrWhiteSpace(GetCgiValue("save")) AndAlso hdnIsAttributeSwitch.Value = 1 Then
                If hdnSwitchPGID.Value.ConvertToLong() > 0 Then
                    ucProductAttributeFilter.IsAttributeSwitch = True
                    AttributeProductGroupID = hdnSwitchPGID.Value.ConvertToLong()
                    RadioButtonList1.Enabled = True
                    RadioButtonList1.ClearSelection()
                    RadioButtonList1.Items(1).Selected = True
                Else
                    ucProductAttributeFilter.IsAttributeSwitch = True
                    AttributeProductGroupID = 0
                End If
                ucProductAttributeFilter.HierarchyTreeURL = "/logix/phierarchytree.aspx?ProductGroupID=" & AttributeProductGroupID & "&CloseAfterSave=" & CloseAfterSave & "PAB=1&OfferID=" & OfferID & "&AttributeProductGroupID=" & AttributeProductGroupID '& locateHierarchyURL
            ElseIf AttributeSwitchType = "DeSelectedAttributeGroup" AndAlso String.IsNullOrWhiteSpace(GetCgiValue("save")) AndAlso hdnIsAttributeSwitch.Value = 1 Then
                ucProductAttributeFilter.IsAttributeSwitch = True
                AttributeProductGroupID = 0
                ucProductAttributeFilter.HierarchyTreeURL = "/logix/phierarchytree.aspx?ProductGroupID=-1&PAB=1&OfferID=" & OfferID & "&CloseAfterSave=" & CloseAfterSave & "&AttributeProductGroupID=" & AttributeProductGroupID '& locateHierarchyURL                
            ElseIf AttributeSwitchType = String.Empty AndAlso IsPostBack AndAlso ucProductAttributeFilter.PABStage <> 2 AndAlso ucProductAttributeFilter.ReloadHierarchyTreeClicked <> True Then  'Like postback due to change in discount type
                ucProductAttributeFilter.IsAttributeSwitch = True
                AttributeProductGroupID = hdnSwitchPGID.Value.ConvertToLong()
                ucProductAttributeFilter.HierarchyTreeURL = "/logix/phierarchytree.aspx?ProductGroupID=" & AttributeProductGroupID & "&CloseAfterSave=" & CloseAfterSave & "&PAB=1&OfferID=" & OfferID & "&AttributeProductGroupID=" & AttributeProductGroupID '& locateHierarchyURL   
                ucProductAttributeFilter.ReloadHierarchyTreeClicked = False
            Else
                ucProductAttributeFilter.IsPGAttributeType = True
                AttributeProductGroupID = discountedpgid.Value.ConvertToLong()
            End If
            ucProductAttributeFilter.ProductGroupID = AttributeProductGroupID
            ucProductAttributeFilter.IsPGAttributeType = True
            ucProductAttributeFilter.LanguageID = LanguageID
            m_EditOfferRegardlessOfBuyer = Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID)
            If Not IsTemplate Then
                ucProductAttributeFilter.IsEditPermitted = (Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And Disallow_Edit))
                RadioButtonList1.Enabled = (Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And Disallow_Edit))
            Else
                ucProductAttributeFilter.IsEditPermitted = Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer
                RadioButtonList1.Enabled = Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer
            End If
        End If

        discountedpgid.Value = DiscountedProductGroupID.ConvertToString()
        'Save Attribute Based Product Group
        If (GetCgiValue("save") <> "" And infoMessage = "") Then
            If (ProductGroupID = 0) Then
                ' get the group created and find out what its ID is.                        
                If (String.IsNullOrWhiteSpace(ProductGroupName)) Then
                    infoMessage = Copient.PhraseLib.Lookup("pgroup-edit.noname", LanguageID)
                Else
                    MyCommon.QueryStr = "select ProductGroupID, Name from ProductGroups with (NoLock) where Name = @Name and Deleted=0 " & _
                                        "union " & _
                                        "select ProductGroupID, Name from ProductGroups with (NoLock) where Name like (@Name + '%') and Deleted=0"
                    MyCommon.DBParameters.Add("@Name", SqlDbType.NVarChar).Value = ProductGroupName
                    Dim pgdt As DataTable = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)

                    If GetButtonListSelectedValue() = 2 Then
                        If Offer.BuyerID Is Nothing Then Offer.BuyerID = 0
                        If discountedpgid.Value > 0 AndAlso ProductGroupTypeID = 1 Then
                            CreateProductGroup(OfferID, Offer.BuyerID, ProductGroupName)
                        ElseIf pgdt.Rows.Count = 0 Then
                            CreateProductGroup(OfferID, Offer.BuyerID, ProductGroupName)
                        ElseIf discountedpgid.Value > 0 AndAlso RadioButtonList1.SelectedItem.Value = "2" Then
                            If ProductGroupName <> m_ProductGroupService.GetProductGroupName(discountedpgid.Value.ConvertToLong()).Result Then
                                m_ProductGroupService.UpdateProductGroupName(discountedpgid.Value.ConvertToLong(), ProductGroupName)
                            End If
                        ElseIf discountedpgid.Value <= 0 Then
                            CreateProductGroup(OfferID, Offer.BuyerID, ProductGroupName)
                        End If
                    End If
                End If
            End If
            MyCommon.Close_LRTsp()

            If GetButtonListSelectedValue() = 2 Then
                'ucProductAttributeFilter.IsAttributeSwitch = True
                ucProductAttributeFilter.ProductGroupID = discountedpgid.Value.ConvertToLong()
                ucProductAttributeFilter.SelectedNodeIDs = NodeID
                ucProductAttributeFilter.SaveData = True
            End If
            AttributeProductGroupID = ucProductAttributeFilter.ProductGroupID
            If (GetCgiValue("save") <> "") Then
                If (DiscountID <= 0) Then
                    DiscountID = Create_Discount(OfferID, TpROID, Phase, DeliverableID, RewardRequired)
                    bCreated = True
                    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Reward.creatediscount", LanguageID))
                End If
                'AMS-685 commented out
                ' Store the product group and excluded product group for later possible TCRM comparison
                MyCommon.QueryStr = "select DiscountedProductGroupID, ExcludedProductGroupID from CPE_Discounts with (NoLock) where DiscountID=" & DiscountID
                rst = MyCommon.LRT_Select
                If (rst.Rows.Count > 0) Then
                    DiscountedProductGroupID = MyCommon.NZ(rst.Rows(0).Item("DiscountedProductGroupID"), -1)
                    'ExcludedProductGroupID = MyCommon.NZ(rst.Rows(0).Item("ExcludedProductGroupID"), -1)
                End If

                ' Save the contents of the discount
                If ValidLevels Then
                    ' DISCOUNT SAVES HERE *~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*
                    Save_Discount(OfferID, DiscountID, TierLevels, WriteTier, bCreated, RewardID, bAllowDollarTransLimit)
                    ' *~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*
                    If (Not bCreated) Then
                        ' Update for TCRM
                        ' Determine if the product group has changed; if so, flag s/b 3
                        If (DiscountedProductGroupID <> Convert.ToInt64(MyCommon.Extract_Decimal(GetCgiValue("discountedpgid"), MyCommon.GetAdminUser.Culture)) _
                           OrElse ExcludedPGChanged(excludedProductGroups, GetCgiValue("excludedpgid"))) Then
                            TCRMAStatusFlag = 3
                        Else
                            TCRMAStatusFlag = 2
                        End If
                        ' If TCRMAStatusFlag is already 3 then don't change it to 2.
                        MyCommon.QueryStr = "update CPE_Deliverables with (RowLock) set TCRMAStatusFlag=" & TCRMAStatusFlag & " " & _
                                            "where TCRMAStatusFlag <> 3 and DeliverableID=" & DeliverableID
                        MyCommon.LRT_Execute()

                        ' update the require reward flag
                        MyCommon.QueryStr = "update CPE_Deliverables with (RowLock) set Required= " & IIf(RewardRequired, 1, 0) & " " & _
                                            "where DeliverableID=" & DeliverableID
                        MyCommon.LRT_Execute()

                        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Reward.editdiscount", LanguageID))
                    End If

                    If (GetCgiValue("save") <> "") Then
                        CloseAfterSave = (SystemCacheData.GetSystemOption_General_ByOptionId(48) = "1")
                    End If
                End If
            End If
        End If
        ' Ensure that a manufacturer coupon does not have best deal selected
        If BestDeal = 1 AndAlso IsMfgCoupon Then
            MyCommon.QueryStr = "update CPE_Discounts with (RowLock) set BestDeal = 0 where DiscountID=" & DiscountID
            MyCommon.LRT_Execute()
        End If
        If Not (ErrorMsg = "") OrElse (GetCgiValue("mode") = "savediscount") Then
            DiscountedProductGroupID = discountedpgid.Value
            If DiscountedProductGroupID > 0 Then
                ComputeDiscount = 1
            End If
            'ExcludedProductGroupID = MyCommon.Extract_Decimal(GetCgiValue("excludedpgid"), MyCommon.GetAdminUser.Culture)
            AmountTypeID = MyCommon.Extract_Decimal(GetCgiValue("l1amounttypeid"), MyCommon.GetAdminUser.Culture)
            ItemLimit = MyCommon.Extract_Decimal(GetCgiValue("itemlimit"), MyCommon.GetAdminUser.Culture)
            WeightLimit = MyCommon.Extract_Decimal(GetCgiValue("weightlimit"), MyCommon.GetAdminUser.Culture)
            DollarLimit = MyCommon.Extract_Decimal(GetCgiValue("dollarlimit"), MyCommon.GetAdminUser.Culture)
            DiscountAmount = GetCgiValue("discountamount")
            L1Cap = MyCommon.Extract_Decimal(GetCgiValue("l1cap"), MyCommon.GetAdminUser.Culture)
            L2DiscountAmt = MyCommon.Extract_Decimal(GetCgiValue("tier1_l2discountamt"), MyCommon.GetAdminUser.Culture)
            L2AmountTypeID = MyCommon.Extract_Decimal(GetCgiValue("l2amounttypeid"), MyCommon.GetAdminUser.Culture)
            L2Cap = MyCommon.Extract_Decimal(GetCgiValue("l2cap"), MyCommon.GetAdminUser.Culture)
            L3DiscountAmt = MyCommon.Extract_Decimal(GetCgiValue("tier1_l3discountamt"), MyCommon.GetAdminUser.Culture)
            L3AmountTypeID = MyCommon.Extract_Decimal(GetCgiValue("l3amounttypeid"), MyCommon.GetAdminUser.Culture)
            ChargebackDeptID = MyCommon.Extract_Decimal(GetCgiValue("chargeback"), MyCommon.GetAdminUser.Culture)
            DiscountBarcode = MyCommon.Extract_Val(GetCgiValue("discountbarcode"))
            VoidBarcode = MyCommon.Extract_Val(GetCgiValue("voidbarcode"))
            RDesc = GetCgiValue("rdesc")
            BuyDesc = GetCgiValue("tier1_buydesc")
            UserGroupID = MyCommon.Extract_Decimal(GetCgiValue("usergroupid"), MyCommon.GetAdminUser.Culture)
            BestDeal = MyCommon.Extract_Decimal(GetCgiValue("bestdeal"), MyCommon.GetAdminUser.Culture)

            DiscAtOrigPrice = MyCommon.Extract_Decimal(GetCgiValue("discpricelevel"), MyCommon.GetAdminUser.Culture)
            If (DiscountType = 4 OrElse DiscountType = 5) Then
                PriceFilterID = 100
            Else
                PriceFilterID = MyCommon.Extract_Val(GetCgiValue("priceFilter"), MyCommon.GetAdminUser.Culture)
            End If

            If GetCgiValue("decliningbalance") = "true" Then
                DecliningBalance = True
            Else
                DecliningBalance = False
            End If
            If GetCgiValue("retrodiscount") = "true" Then
                RetroactiveDiscount = True
            Else
                RetroactiveDiscount = False
            End If
            PriceChange = (GetCgiValue("pricechange") = "true")

            DiscountType = MyCommon.Extract_Decimal(GetCgiValue("discountType"), MyCommon.GetAdminUser.Culture)
            DeptLevel = IIf(DiscountType = 2, 1, 0)
            'If DiscountType = 3 Then
            '    DeptLevel = 3 ' IIf(DiscountType = 3, 1, 0)
            'ElseIf DiscountType = 2 Then
            '    DeptLevel = 2 'IIf(DiscountType = 2, 1, 0)
            'Else
            '    DeptLevel = 1
            'End If
            If CMS.Utilities.Extract_Val(GetCgiValue("discountsv")) = 1 Then
                bSVProgram = True
                SVProgramID = CMS.Utilities.Extract_Val(GetCgiValue("ddldiscountsv"))
            Else
                bSVProgram = False
                SVProgramID = 0
            End If

            ProrationTypeID = MyCommon.Extract_Decimal(GetCgiValue("prorationtypeid"), MyCommon.GetAdminUser.Culture)
        End If

        AnyProduct = (DiscountedProductGroupID = 1)



        If (IsTemplate Or FromTemplate) Then
            ' Get the permissions if it's a template
            MyCommon.QueryStr = "select DisallowEdit from CPE_Deliverables with (NoLock) where DeliverableID=" & DeliverableID
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count > 0) Then
                Disallow_Edit = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit"), False)
            Else
                Disallow_Edit = False
            End If
            Try
                ucProductAttributeFilter.IsEditPermitted = (Logix.UserRoles.EditOffer And (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID)) And Not (FromTemplate And Disallow_Edit))
                RadioButtonList1.Enabled = (Logix.UserRoles.EditOffer And (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID)) And Not (FromTemplate And Disallow_Edit))
            Catch ex As Exception

            End Try
            ' Check field-level permissions
            MyCommon.QueryStr = "select UI.FieldID, ISNull(TFP.Editable, 0) as Editable, UI.ControlName, UI.FieldTypeID, UI.Tiered from UIFields UI with (NoLock) " & _
                                "left join TemplateFieldPermissions TFP with (NoLock) on UI.FieldID=TFP.FieldID " & _
                                "where OfferID=" & OfferID & " and UI.PageName='" & MyCommon.AppName & "';"
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count > 0) Then
                Dim fieldID As Integer
                For Each row In rst.Rows
                    If (Not OverrideFields.ContainsKey(MyCommon.NZ(row.Item("ControlName"), ""))) Then
                        fieldID = MyCommon.NZ(row.Item("FieldID"), 0)
                        'BZ 4287/4310: If the field is one that can repeat per tier, prepend the tier level to the control name
                        If MyCommon.NZ(row.Item("Tiered"), False) Then
                            For i = 1 To TierLevels
                                OverrideFields.Add("tier" & i & "_" & MyCommon.NZ(row.Item("ControlName"), ""), MyCommon.NZ(row.Item("Editable"), False))
                                If MyCommon.NZ(row.Item("FieldTypeID"), 0) = 2 Then
                                    OverrideDiv.Add("tier" & i & "_" & MyCommon.NZ(row.Item("ControlName"), ""), True)
                                Else
                                    OverrideDiv.Add("tier" & i & "_" & MyCommon.NZ(row.Item("ControlName"), ""), False)
                                End If
                            Next
                        Else
                            OverrideFields.Add(MyCommon.NZ(row.Item("ControlName"), ""), MyCommon.NZ(row.Item("Editable"), False))
                            If MyCommon.NZ(row.Item("FieldTypeID"), 0) = 2 Then
                                OverrideDiv.Add(MyCommon.NZ(row.Item("ControlName"), ""), True)
                            Else
                                OverrideDiv.Add(MyCommon.NZ(row.Item("ControlName"), ""), False)
                            End If
                        End If
                        If (MyCommon.NZ(row.Item("Editable"), False) = True) Then
                            OvrdFldEditable = True
                        End If
                    End If
                Next
            End If
        End If
        If hdnSwitchPGID.Value <> 0 Then
            AttributeProductGroupID = hdnSwitchPGID.Value
        End If
        DisabledAttribute = IIf(Logix.UserRoles.EditOffer And (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID)) And Not (FromTemplate And Disallow_Edit), "", " disabled=""disabled""")
        SetDisabledAttr(DisabledAttribute)
        If (Not String.IsNullOrEmpty(DisabledAttribute)) Then
            txtProductGroupName.Disabled = True
        End If


        Send_HeadBegin("term.offer", "term.discountreward", OfferID)
        Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
        Send_Metas()
        Send_Links(Handheld)
        Send_Scripts(New String() {"jquery.min.js"})

    End Sub
	
	Private Function ConvertDataTabletoJson(ByRef StrJson As String, ByVal ObjDataTable As DataTable) As Boolean
        Dim IsSuccess As Boolean = True
        Try
            StrJson = Newtonsoft.Json.JsonConvert.SerializeObject(ObjDataTable)
            IsSuccess = True
        Catch ex As Exception
            StrJson = String.Empty
        End Try
        
        Return IsSuccess
    End Function

	
    Function ExcludedPGChanged(ByVal listExcludedPG As List(Of DiscountProductGroup), ByVal currentlySelectedPGs As String) As Boolean
        Dim exPGListChanged As Boolean = False
        If Not String.IsNullOrWhiteSpace(currentlySelectedPGs) Then
            For Each pgId As String In currentlySelectedPGs.Split(",")
                If listExcludedPG Is Nothing AndAlso Not listExcludedPG.Exists(Function(x) x.ProductGroupId.ToString = Trim(pgId)) Then
                    exPGListChanged = True
                    Exit For
                End If
            Next
        End If
        Return exPGListChanged
    End Function
    Sub UpdateCurrentListOfProductGroups(ByRef listExcludedPG As List(Of DiscountProductGroup), ByVal currentlySelectedPGs As String)
        Dim dpg As DiscountProductGroup
        Dim tempList As New List(Of DiscountProductGroup)

        If Not String.IsNullOrWhiteSpace(currentlySelectedPGs) Then
            For Each pgId As String In currentlySelectedPGs.Split(",")
                'If Not listExcludedPG.Exists(Function(x) x.ProductGroupId.ToString = Trim(pgId)) Then
                dpg = New DiscountProductGroup()
                dpg.ProductGroupId = pgId
                'dpg.ProductGroupName = pgName  
                dpg.Excluded = 1

                tempList.Add(dpg)
                'End If
            Next
        End If
        If listExcludedPG Is Nothing Then
            listExcludedPG = New List(Of DiscountProductGroup)()
        Else
            listExcludedPG.Clear()
        End If
        listExcludedPG.AddRange(tempList)
    End Sub
    Public Sub SetScoreCardMLI(ByVal DiscountID As Long, ByRef MLI As Copient.Localization.MultiLanguageRec)
        MLI.ItemID = DiscountID
        MLI.MLTableName = "CPE_DiscountTranslations"
        MLI.MLIdentifierName = "DiscountID"
        MLI.MLColumnName = "ScorecardDesc"
        MLI.StandardTableName = "CPE_Discounts"
        MLI.StandardIdentifierName = "DiscountID"
        MLI.StandardColumnName = "ScorecardDesc"
        MLI.StandardValue = MyCommon.NZ(ScorecardDesc.Replace("""", "&quot;"), "")
        MLI.InputName = "ScorecardDesc"
        MLI.InputID = "ScorecardDesc"
        MLI.InputType = "text"
        MLI.LabelPhrase = ""
        MLI.MaxLength = 30
        MLI.CSSClass = ""
    End Sub

</script>
<script type="text/javascript" language="javascript">
    var fetchPGAsync = <%= IIf(shouldFetchPGAsync, 1, 0)%>;
    var lblNotifierText = "";
    var loadInProgress = false;
    var totalPGInPage;

    $(document).ready(function () {
        RegisterPGAsyncLoadHandler();
        totalPGInPage = GetTotalPGInPage();
    });
    // create string buffer class to efficiently concatenate strings
    window.onload = function () {
        if ($("#RadioButtonList1 input[type=radio]:checked").val() == 2) {
            ShoworHideDivs();
        }
        ShowHideIncludeBox();
        updateButtons();
    }
    function GetTotalPGInPage()
    {
        var totalProductGroupsInPage = $('#functionselect option').length + $('#selected option').length + $('#excluded option').length;
        return totalProductGroupsInPage;    
    }

    function RegisterPGAsyncLoadHandler()
    {
        if (fetchPGAsync == 1) {
            $('#functionselect').on("scroll", ProductGroupLoader);
        }
    }
    function ProductGroupLoader(event) {
        if ($('#functioninput').val() == "") {
            var data = JSON.stringify({ lastPGName: GetLastPGName(), AdminUserID: <%= AdminUserID %>, viewProductgroupRegardlessBuyer: <%= Logix.UserRoles.ViewProductgroupRegardlessBuyer.ToString.ToLower %>});
            LoadItemsOnScroll("functionselect", "<%=Request.Url.AbsolutePath%>/GetProductGroupListJSON", data);
        }
    }

    function OnLoadError(response, status, error) {
        lblNotifierText += "<%=Copient.PhraseLib.Lookup("term.erroronpgload", LanguageID)%>"; //"Error occurred while loading product groups";
        if (error)
            lblNotifierText += error;
        $('#lblAjaxNotification').text(lblNotifierText);
    }
    function BeforeSendSetup(request) {
        loadInProgress = true;

        $('#divNotification').show();
        lblNotifierText += "<%=Copient.PhraseLib.Lookup("term.loadingpg", LanguageID)%>";//"Loading product groups...";
    }
    function OnLoadSuccess(response) {
        if (response.d)
            lblNotifierText += "<%=Copient.PhraseLib.Lookup("term.done", LanguageID)%>";//"Done";
        else
            lblNotifierText = "<%=Copient.PhraseLib.Lookup("term.nomorepg", LanguageID)%>";//"No more product groups to load";
        $('#functionselect').append($.parseHTML(response.d));

        selectBoxItemLengthBeforeSearch = $('#functionselect option').length;

        $('#lblAjaxNotification').text(lblNotifierText);
        loadInProgress = false;
        $('#divNotification').delay(500).fadeOut();
        lblNotifierText = "";
    }
    //Get the last PG name in the select list box
    function GetLastPGName() {
        var const_last_pgName= "zzzz";
        var len = $('select#functionselect option').length - 1;   //Get length
        if (len > 0 && $("#functionselect option")[len].text != "")
            return $("#functionselect option")[len].text;   //Get the PG name

        return const_last_pgName;
    }

    function ShoworHideDivsWithoutPostback() {
        if ($("#RadioButtonList1 input[type=radio]:checked").val() == 2) {
            ShoworHideDivs();
        }
        ShowHideIncludeBox();
    }

    function StringBuffer() {
        this.buffer = [];
    }

    StringBuffer.prototype.append = function append(string) {
        this.buffer.push(string);
        return this;
    };

    StringBuffer.prototype.toString = function toString() {
        return this.buffer.join("");
    };

    function toggleScorecardText() {
        if (document.getElementById("ScorecardID") != null) {
            if (document.getElementById("ScorecardID").value == 0) {
                document.getElementById("ScorecardDescLine").style.display = 'none';
                document.getElementById("ScorecardDesc").value = '';
            } else {
                document.getElementById("ScorecardDescLine").style.display = '';
            }
        }
    }

    // This is the javascript array holding the function list
    // The PrintJavascriptArray ASP function can be used to print this array.
      <%
    Dim buyerid As Integer
    Dim externalBuyerid As string
    MyCommon.QueryStr = "select ProductGroupID, BuyerId, Name from ProductGroups with (NoLock) " & _
                        "where Deleted=0 and ProductGroupID<>1 and NonDiscountableGroup=0 and PointsNotApplyGroup=0 "
    If (MyCommon.IsUserAssociatedWithAnyBuyer(AdminUserID) AndAlso Not Logix.UserRoles.ViewProductgroupRegardlessBuyer) Then
        MyCommon.QueryStr &= " and (BuyerId in(select BuyerId from BuyerRoleUsers where AdminUserID=@UserId) or BuyerId is null) "
        MyCommon.DBParameters.Add("@UserId", SqlDbType.Int).Value = AdminUserID
    End If
    MyCommon.QueryStr &= " order by Name;"
    
    rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
    If (rst.Rows.Count > 0) Then
        Sendb("var functionlist = Array(")
        For Each row In rst.Rows
            If (MyCommon.IsEngineInstalled(9) And SystemCacheData.GetSystemOption_UE_ByOptionId(168) = "1" And Not IsDBNull(row.Item("Buyerid"))) Then
                buyerid = row.Item("Buyerid")
                externalBuyerid = MyCommon.GetExternalBuyerId(buyerid)
                Sendb("""" & "Buyer " & externalBuyerid & " - " & MyCommon.NZ(row.Item("Name"), "").ToString().Replace("""", "\""") & """,")
            Else 
                Sendb("""" & MyCommon.NZ(row.Item("Name"), "").ToString().Replace("""", "\""") & """,")
            End If
        Next
        Sendb(""""");")
        Sendb("var vallist = Array(")
        For Each row In rst.Rows
            Sendb("""" & MyCommon.NZ(row.Item("ProductGroupID"), 0) & """,")
        Next
        Sendb(""""");")
    Else
        Sendb("var functionlist = Array(")
        Send("""" & "" & """);")
        Sendb("var vallist = Array(")
        Send("""" & "" & """);")
    End If
    Sendb("var arrattributePGIds = [")
    Sendb(String.Join(",", lstAttributePGIDs))
    Send("];")
      %>
    function ShowHideIncludeBox()
    {
        var selectBoxDiv = document.getElementById("selectBox");
        var excludeBox = document.getElementById("excluded");
        var selectButton = document.getElementById("pselect");
        var deSelectButton = document.getElementById("pdeselect");
        var discType = document.getElementById("discountType");
        if (GetSelectedRadioButton() == 1 || (discType != null && discType.value == "3")) {
            //if(selectBoxDiv != null)
                selectBoxDiv.style.display = "block";
            //if(excludeBoxnull)
                excludeBox.size = 4;
            //if(selectButton !=null)
                selectButton.style.display = "inline";
            deSelectButton.style.display = "inline";
        }
        else 
        {
            selectBoxDiv.style.display = "none";
            excludeBox.size = 10;
            selectButton.style.display = "none";
            deSelectButton.style.display = "none";
        }
    }
    //Returns 1 for standard product group, 2 for Attribute based product group
    function GetSelectedRadioButton()
    {
        var radiobtn = $("#RadioButtonList1 input[type=radio]:checked").val();
        if (typeof (radiobtn) == "undefined") {
            radiobtn = "1";
        }

        return radiobtn;
    }
    function UpdateProductGroupTitle()
    {

    }
    // This is the function that refreshes the list after a keypress.
    // The maximum number to show can be limited to improve performance with
    // huge lists (1000s of entries).
    // The function clears the list, and then does a linear search through the
    // globally defined array and adds the matches back to the list.
    function handleKeyUp(maxNumToShow) {
        var selectObj, textObj, functionListLength;
        var i, numShown;
        var searchPattern;

        document.getElementById("functionselect").size = "10";

        // Set references to the form elements
        selectObj = document.forms[0].functionselect;
        textObj = document.forms[0].functioninput;

        // Remember the function list length for loop speedup
        functionListLength = functionlist.length;

        // Set the search pattern depending
        if (document.forms[0].functionradio[0].checked == true) {
            searchPattern = "^" + textObj.value;
        } else {
            searchPattern = textObj.value;
        }
        searchPattern = cleanRegExpString(searchPattern);

        // Create a regulare expression
        re = new RegExp(searchPattern, "gi");

        // Loop through the array and re-add matching options
        var buf = new StringBuffer();
        numShown = 0;
       
        // Clear the options list 
        selectObj = clearOptionsFast(selectObj);
        for (i = 0; i < functionListLength; i++) {
            if (textObj.value != "") {              //When there is text in search textbox
                if (functionlist[i].search(re) != -1) {
                    AddOptionSelectBoxOption(buf, i);
                numShown++;
            }
            }
            else
            {   //When there is no search text                 
                if (i < totalPGInPage) {  //totalPGInPage set during page load, Add options only for the number of items in the select box instead of full list as in functionlistlength
                    AddOptionSelectBoxOption(buf, i);
                    numShown++;
                }
            }
            // Stop when the number to show is reached
            if (numShown == maxNumToShow) {
                break;
            }
        }
        select_innerHTML(selectObj, buf.toString());
        selectObj = document.forms[0].functionselect;

        removeUsed(true);
        // When options list whittled to one, select that entry
        if (selectObj.length == 1) {
            selectObj.options[0].selected = true;
        }
    }
    function AddOptionSelectBoxOption(buf, index)
    {
        if ($.inArray(parseInt(vallist[index]), arrattributePGIds) > -1) {
            buf.append('<option title="' + functionlist[index] + '"  value="' + vallist[index] + '" style="color:blue;">' + functionlist[index] + '<\/option>')
        }
        else {
            buf.append('<option title="' + functionlist[index] + '" value="' + vallist[index] + '">' + functionlist[index] + '<\/option>')
        }
    }

    function clearOptionsFast(selectObj) {
        var selectParentNode = selectObj.parentNode;
        var newSelectObj = selectObj.cloneNode(false);
        selectParentNode.replaceChild(newSelectObj, selectObj);
        RegisterPGAsyncLoadHandler();  //Attach the scroll event handler after replacement
        return newSelectObj;
    }

    function select_innerHTML(select, inner) {
        // Firefox does not support the outerHTML property
        if (navigator.appName.indexOf("Netscape") == -1) {
            select.outerHTML = select.outerHTML.substring(0, select.outerHTML.indexOf('>', 0) + 1) + inner + '<\/select>';
        } else {
            select.innerHTML = inner;
        }
    }

    // this function gets the selected value and loads the appropriate
    // php reference page in the display frame
    // it can be modified to perform whatever action is needed, or nothing
    function handleSelectClick() {
        selectObj = document.forms[0].functionselect;
        textObj = document.forms[0].functioninput;
        //selectedValue = selectObj.options[selectObj.selectedIndex].text;
        //selectedValue = selectedValue.replace(/_/g, '-') ;
        selectedValue = document.getElementById("functionselect").value;
        if (selectedValue != "") {
        }
    }

      <% If (IsEditable) Then%>
    function submitForm() {
        assignValues();
        document.mainform.submit();
    }

    function assignValues() {
        if (typeof spin === 'function')
            spin('divAttributeBuilder');//PAB
        var elemPG = document.getElementById("discountedpgid");
        var elemExPG = document.getElementById("excludedpgid");
        var elemCG = document.getElementById("usergroupid");
        var flexnegative = document.getElementById("flexnegative");
        var hdnflexnegative = document.getElementById("hdnflexnegative");
        // Discounted Product Group
        elemPG.value = GetDiscountedPGID();
        //AMS-685
        //setLastDiscountedPGID(elemPG.value);
        // Excluded Product Group
        if (document.mainform.excluded != null && document.mainform.excluded.options.length > 0) {
            //AMS-685 multiple exclusion groups
            //elemExPG.value = document.mainform.excluded.options[0].value;
            var optionsArray = [];
            for(var i=0; i<document.mainform.excluded.options.length; i++)
            {
                optionsArray.push(document.mainform.excluded.options[i].value);
            }
            elemExPG.value = optionsArray.join(",");
        }
        else if (document.mainform.excluded != null && document.mainform.excluded.options.length == 0)
        {
            elemExPG.value = "";
        }

        hdnflexnegative.value = "";

        if (flexnegative != null) {

            hdnflexnegative.value = (flexnegative.selectedIndex);
        }

        if (document.mainform.mode != null) {
            document.mainform.mode.value = "savediscount";
        }
        <% If (FromTemplate) Then%>
        enableFormFields(document.mainform);
        <% End If%>
    }
      <% Else%>
    function submitForm() {
        sendNotEditable();
        return false;
    }

    function assignValues() {
        sendNotEditable();
        return false;
    }

    function sendNotEditable() {
        alert('<% Sendb(Copient.PhraseLib.Lookup("CPE-rew-noedit", LanguageID))%>');
    }
      <% End If%>

    function GetDiscountedPGID() {
        var elemSel = document.getElementById("selected");
        var pgID;
        if (elemSel.options.length > 0) {
            pgID = elemSel.options[0].value;
        } else {
            pgID = "";
        }
        return pgID;
    }
    function setLastDiscountedPGID(discountedPGID)
    {        
        var hdnLastDiscountedPG = document.getElementById("hdnLastDiscountedPG");
        var discType = document.getElementById('hidDiscountType'); 
        if(discType != null && discType.value != 3)
            hdnLastDiscountedPG.value = discountedPGID;
    }
    function IsSelectedValuePABGroup(selVal)
    {
        var objresult = $.inArray(parseInt(selVal), arrattributePGIds);
        return objresult;
    }
    function DoPGSwitch(objresult, selVal)
    {
        if (objresult >= 0) {
            $("#hdnIsAttributeSwitch").val("1");
            $("#hdnSwitchPGID").val(selVal);
            $("#AttributeSwitchType").val("SelectedAttributeGroup");
            submitForm();
        }
        else {
            $("#hdnSwitchPGID").val(0);
            $("#hdnIsAttributeSwitch").val("1");
            $("#AttributeSwitchType").val("DeSelectedAttributeGroup");
            submitForm();
        }
    }
    function selectItem(source, dest) {
        //AMS-4253 disable buttons so that while loading user cant select\deselect and they will be enabled when page is loaded
        disableButtons();
        var selVal = moveItem(source, dest, isPABPG);
        var isPABPG = IsSelectedValuePABGroup(selVal);
        DoPGSwitch(isPABPG, selVal);
        //updateButtons();
    }
    function disableButtons()
    {
       // $('#pselect').attr("disabled", "disabled");
        $('#pdeselect').attr("disabled", "disabled");
        $('#Button1').attr("disabled", "disabled");
        $('#Button2').attr("disabled", "disabled");
    }
    function updateButtons()
    {
        var selBox = document.getElementById('selected');
        var exBox = document.getElementById('excluded');
        var functionSelBox = document.getElementById('functionselect');

        var selectButton = document.getElementById('pselect');
        var deselectButton = document.getElementById('pdeselect');

        var exSelectButton = document.getElementById('Button1');
        var exDeselectButton = document.getElementById('Button2');
        var discType = document.getElementById("discountType");        

        if (selBox != null) {
            if (selBox.length == 1){
            selectButton.disabled = true;
            deselectButton.disabled = false;
            }
        else if (selBox.length == 0){
            deselectButton.disabled = true;
            selectButton.disabled = false;
        }
        }
        if(exBox != null)
        {
            if(exBox.length == 0)
            {
                exDeselectButton.disabled = true;
                exSelectButton.disabled = false;
            }
            else if(exBox.length > 0)
            {
                exDeselectButton.disabled = false;
                exSelectButton.disabled = false;
            }
        }
        if(functionSelBox != null)
        {
            if(functionSelBox.length == 0)                  //FOR ONE CASE LENGTH=1 EVEN IF THERE ARE NO ITEMS IN LIST
            {
                selectButton.disabled = true;
                exSelectButton.disabled = true;
            }
        }
        if (discType!=null && discType.value == "3")
        {
            selectButton.disabled = true;
            deselectButton.disabled = true;
        }
    }
    function moveItem(source, dest, isPABPG, selected) {
        var elemSource = document.getElementById(source);
        var elemDest = document.getElementById(dest);
        var selOption = null;
        var selText = "", selVal = "";
        var selIndex = -1;

        if(elemSource != null && elemDest != null)
        {
            selIndex = elemSource.options.selectedIndex;
            if (selIndex != -1) {
                selOption = elemSource.options[selIndex];
                selText = selOption.text;
                selVal = selOption.value;
                elemDest.options[elemDest.length] = new Option(selText, selVal);
            
                if (isPABPG >= 0)
                    elemDest.options[elemDest.length-1].style.color = 'blue';

                elemSource.options[selIndex] = null;
                removeUsed(selected);   //selected=true means we do not need to sort the functionlist but on deselection we need to sort
            }
                
        }
        return selVal;
    }
    function moveExcludedItem(source, dest, selected)
    {
        var elemSource = document.getElementById(source);
        var selVal = "";
        if (elemSource != null)
        {
            selVal = elemSource.options[elemSource.options.selectedIndex].value;
        }
        
        var isPABPG = IsSelectedValuePABGroup(selVal);
        moveItem(source, dest, isPABPG, selected);
    }
    function deselectItem(source,dest) {
        //removeInvalidParamsFromAction();
        var selVal = moveItem(source, dest)
<%--        var elemSource = document.getElementById(source);
        var selIndex = -1;
        var selVal = null;
        selIndex = elemSource.options.selectedIndex;
        selOption = elemSource.options[selIndex];
        selVal = selOption.value;
        if (elemSource != null && elemSource.options.selectedIndex == -1) {
            alert('<% Sendb(Copient.PhraseLib.Lookup("CPE-rew-discounts.selectproducts", LanguageID))%>');
            elemSource.focus();
        } else {
            elemSource.options[0] = null;
            removeUsed(false);
        }--%>
        var objresult = $.inArray(parseInt(selVal), arrattributePGIds);
        if (objresult >= 0) {
            $("#hdnIsAttributeSwitch").val("1");
            $("#hdnSwitchPGID").val(0);
            $("#AttributeSwitchType").val("DeSelectedAttributeGroup");
            submitForm();
        }
        //updateButtons();
    }
    function removeUsed(bSkipKeyUp) {
        if (!bSkipKeyUp) handleKeyUp(99999);
        // this function will remove items from the functionselect box that are used in 
        // selected and excluded boxes

        var funcSel = document.getElementById('functionselect');
        var exSel = document.getElementById('selected');
        var elSel = document.getElementById('excluded');


        var i, j;
        if (elSel != null) {
            for (i = elSel.length - 1; i >= 0; i--) {
                for (j = funcSel.length - 1; j >= 0; j--) {
                    if (funcSel.options[j].value == elSel.options[i].value) {
                        funcSel.options[j] = null;
                    }
                }
            }
        }
        if (exSel != null) {
            for (i = exSel.length - 1; i >= 0; i--) {
                for (j = funcSel.length - 1; j >= 0; j--) {
                    if (funcSel.options[j].value == exSel.options[i].value) {
                        funcSel.options[j] = null;
                    }
                }
            }
        }
    }


    function enableFormFields(theForm) {
        var elems = theForm.elements;
        var elem = null;

        for (var x = 0; x < elems.length; x++) {
            elem = elems[x];
            if (elem != null) {
                if (elem.disabled == true)
                {
                    elem.readOnly = true;
                    elem.disabled = false;
                }
            }
        }
    }

    function updateTableLockStatus() {
        var elemTbl = document.getElementById("tblTempFields");
        var elemTr = null, elemTd = null, elemChk = null;

        if (elemTbl != null) {
            elemTrs = elemTbl.getElementsByTagName("TR");
            for (var i = 1; i < elemTrs.length; i++) {
                elemTd = elemTrs[i].firstChild.nextSibling;
                if (elemTd != null) {
                    updateLockStatus(elemTd.firstChild);
                }
            }
        }
    }

    function updateLockStatus(elem) {
        var pageElem = document.mainform.Disallow_Edit;
        if (pageElem != null) {
            var pageLockChecked = pageElem.checked;
            if (elem != null) {
                var td1 = elem.parentNode;
                if (td1 != null) {
                    var tr = td1.parentNode;
                    if (tr != null) {
                        //var td3 = tr.lastChild;
                        var td3 = $(tr).children('td:last');
                        if (td3 != null) {
                            if (pageLockChecked) {
                                td3.html((td3.html() == '<% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>') ? '<% Sendb(Copient.PhraseLib.Lookup("term.unlocked", LanguageID))%>' : '<% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>');
                            } else {
                                td3.html((elem.checked) ? '<% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>' : '<% Sendb(Copient.PhraseLib.Lookup("term.unlocked", LanguageID))%>');
                            }
                        }
                    }
                }
            }
        }
    }

    function handleDiscType_Change(selValue) {
        var elemSel = document.getElementById("selected");
        var elemEx = document.getElementById("excluded");
        var elemList = document.getElementById("functionselect");
        var hdnLastDiscountedPG = document.getElementById("hdnLastDiscountedPG");
        document.getElementById('hidDiscountType').value = selValue;
        if (parseInt(selValue) != 0) {
            document.getElementById('l1amounttypeid').options[0].selected = true;
        }

        if (parseInt(selValue) == 3) {
            hdnLastDiscountedPG.value = GetDiscountedPGID();
            elemSel.options[0] = new Option('<% Sendb(Copient.PhraseLib.Lookup("term.anyproduct", LanguageID))%>', '1');
    } else /*if (parseInt(selValue) == 4 || parseInt(selValue) == 5  */ // clears all groups when switching discount types.
        {
        //AMS-685
        // AL-4233
        //if (elemSel.options.length > 0 && elemSel.options[0].value == 1) {
        //    // empty out all selected and excluded product groups
        //    emptySelectedOptions(elemSel);
        //    emptySelectedOptions(elemEx);
        //}
    }
    submitForm();
	document.getElementById('isDropdownChanged').value = selValue;
    //ShowHideIncludeBox(selValue);
}

function emptySelectedOptions(selector) {
    if (selector != null) {
        while (selector.length > 0) {
            selector.options[0] = null;
        }
    }
}


function handleChargebackDept(defaultValue) {
    var elemChrg = document.getElementById("chargeback");

    if (elemChrg != null) {
        for (var i = 0; i < elemChrg.options.length; i++) {
            if (elemChrg.options[i].value == defaultValue) {
                elemChrg.selectedIndex = i;
            }
        }
    }

}

function assignDiscPriceLevel(grossprice){
    var discpricelevel = document.getElementById("discpricelevel");
    if(discpricelevel != null){
        if(grossprice.checked && discpricelevel.options[1] != null){
            discpricelevel.removeChild( discpricelevel.options[1] ); 
        }
        else if(discpricelevel.options[1] == null){
            discpricelevel.options.add(new Option('<% Sendb(Copient.PhraseLib.Lookup("reward.discount-originalprice", LanguageID)) %>', '1'), discpricelevel.options[1]);
        }
    }
}

function assignGrossPriceDefault(){
        var selectedChargeBack = $("#chargeback option:selected").val();
        var grossprice = document.getElementById("grossprice");        
    
        if('<% Sendb(DiscountID)%>' <= '0'){  
            if(selectedChargeBack == "0"){
                grossprice.checked = ('<%Sendb(grossPriceForDiscount)%>' == 'True')
            }
            else if(selectedChargeBack == "14"){
                grossprice.checked = ('<%Sendb(grossPriceForTender)%>' == 'True')
            }
        }
        assignDiscPriceLevel(grossprice);
        
    }
function handleChargebackSubmit() {
    var elem = document.getElementById("loadDefaultChargeback");

    if (elem != null) {
        elem.value = "0";
    }

    submitForm();
}

function ShowFieldList() {
    var elemList = document.getElementById("templatefields");

    if (elemList != null) {
        if (elemList.style.display == 'block') {
            elemList.style.display = 'none';
        } else {
            elemList.style.display = 'block';
        }
    }
}

function xmlhttpPost(strURL) {
    var xmlHttpReq = false;
    var self = this;

    document.getElementById("results").innerHTML = "<div class=\"loading\"><img id=\"clock\" src=\"../images/clock22.png\" \/><br \/>" + '<% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID))%><\/div>';

        // Mozilla/Safari
        if (window.XMLHttpRequest) {
            self.xmlHttpReq = new XMLHttpRequest();
        }
            // IE
        else if (window.ActiveXObject) {
            self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
        }
        self.xmlHttpReq.open('POST', strURL, true);
        self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
        self.xmlHttpReq.onreadystatechange = function () {
            if (self.xmlHttpReq.readyState == 4) {
                updatepage(self.xmlHttpReq.responseText);
            }
        }
        self.xmlHttpReq.send('<%Sendb("LanguageID=" & LanguageID)%>');
        return false;
    }
    function xmlhttpPost1(strURL) {
        var xmlHttpReq = false;
        var self = this;
        document.getElementById("results").innerHTML = "<div class=\"loading\"><img id=\"clock\" src=\"../images/clock22.png\" \/><br \/>" + '<% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID))%><\/div>';
        // Mozilla/Safari
        if (window.XMLHttpRequest) {
            self.xmlHttpReq1 = new XMLHttpRequest();
        }
            // IE
        else if (window.ActiveXObject) {
            self.xmlHttpReq1 = new ActiveXObject("Microsoft.XMLHTTP");
        }
        self.xmlHttpReq1.open('POST', strURL, true);
        self.xmlHttpReq1.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
        self.xmlHttpReq1.onreadystatechange = function () {
            if (self.xmlHttpReq1.readyState == 4) {
                updatepage(self.xmlHttpReq1.responseText);
            }
        }
        self.xmlHttpReq1.send('<%Sendb("LanguageID=" & LanguageID)%>');
        return false;
    }

    function updatepage(str) {
        document.getElementById("results").innerHTML = str;
    }

    function zeroLevel(level) {
        var levelName = 'l' + level + 'cap';
        var elem = document.getElementById(levelName);
        if (elem != null) elem.value = "0";

        var levelType = 'l' + level + 'amounttypeid'
        elem = document.getElementById(levelType)

        if (elem != null) {
            elem.selectedIndex = 2; // set it to percent off
        }

        submitForm();
    }

    function handleUpToEntry(val, Level) {
        var elemArrow = document.getElementById("btnDown" + Level);
        var elemEx = document.getElementById("btnEx" + Level);

        if (!isNaN(val) && parseFloat(val) > 0) {
            if (elemArrow != null) elemArrow.disabled = false;
            if (elemEx != null) elemEx.disabled = false;
        } else {
            if (elemArrow != null) elemArrow.disabled = true;
            if (elemEx != null) elemEx.disabled = true;
        }
    }

    function addLevel(tier) {
        var levelsElem = document.getElementById('tier' + tier + '_levels');
        var levels = parseInt(levelsElem.value);

        var highestLevelElem = document.getElementById('tier' + tier + '_highestlevel');
        var highestLevel = parseInt(highestLevelElem.value);

        var itemLimitElem = document.getElementById('tier' + tier + '_itemlimit');
        var itemLimit = parseInt(itemLimitElem.value);

        var radioElem = document.getElementById('tier' + tier + '_sprepeatlevel' + (highestLevel + 1));
        var inputElem = document.getElementById('tier' + tier + '_level' + (highestLevel + 1));
        var buttonElem = document.getElementById('tier' + tier + '_deletelevel' + (highestLevel + 1));
        var rowElem = document.getElementById('tier' + tier + '_level' + (highestLevel + 1) + 'row');
        var newLevelInput = document.getElementById('tier' + tier + '_newlevel');

        if ((parseInt(itemLimitElem.value) > 0) && (levels >= itemLimit)) {
            alert('<%Sendb(Copient.PhraseLib.Lookup("ueoffer-rew-discount.ExceedsItemLimit", LanguageID).Replace("'", "\'"))%>');
        } else if (newLevelInput.value == '') {
            alert('<%Sendb(Copient.PhraseLib.Lookup("ueoffer-rew-discount.SpecifyValue", LanguageID))%>');
        } else {
            if (document.getElementById('tier' + tier + '_level' + (highestLevel + 1) + 'row') != null) {
                radioElem.disabled = false;
                inputElem.disabled = false;
                buttonElem.disabled = false;
                inputElem.value = newLevelInput.value; // Put the "new level" value into the revealed input
                rowElem.style.display = ''; //Show the row
                newLevelInput.value = ''; //Blank out the "new level" value
                levelsElem.value = (levels + 1); //Increment the hidden "levels" field
                highestLevelElem.value = (highestLevel + 1); //Increment the hidden "highest levels" field
            } else {
                alert('<%Sendb(Copient.PhraseLib.Lookup("ueoffer-rew-discount.SaveBeforeAdding", LanguageID))%>');
            }
        }
}

function deleteLevel(tier, level) {
    var levelsElem = document.getElementById('tier' + tier + '_levels');
    var levels = parseInt(levelsElem.value);

    var highestLevelElem = document.getElementById('tier' + tier + '_highestlevel');
    var highestLevel = parseInt(highestLevelElem.value);

    var radioElem = document.getElementById('tier' + tier + '_sprepeatlevel' + level);
    var inputElem = document.getElementById('tier' + tier + '_level' + level);
    var buttonElem = document.getElementById('tier' + tier + '_deletelevel' + level);
    var rowElem = document.getElementById('tier' + tier + '_level' + level + 'row');

    if (document.getElementById('tier' + tier + '_level' + level + 'row') != null) {
        radioElem.disabled = true;
        inputElem.disabled = true;
        buttonElem.disabled = true;
        rowElem.style.display = 'none'; //Hide the row
        levelsElem.value = (levels - 1); //Put the new total into the hidden "levels" field
    }
}

function toggleflexneg() {
    var allowElem = document.getElementById('allownegative');
    var flexnegElem = document.getElementById('flexnegative');

    if (allowElem != null) {
        if (allowElem.checked == false) {
            flexnegElem.disabled = false;
        } else {
            flexnegElem.selectedIndex = 0;
            flexnegElem.disabled = true;
        }
    }
}
function getIDList() {
    if (typeof Groupgrid !== "undefined")
        UpdateProductChanges();//For groupgrid update the product and level exclude details
    if (typeof idList !== "undefined") {
        var nodelist = document.getElementById('NodeListID');
        nodelist.value = idList;
    }
}

function ProductGroupDivSelection() {
    if ($("#RadioButtonList1") == null) {
        return;
    }
    var radiobtn = $("#RadioButtonList1 input[type=radio]:checked").val();
    if (typeof (radiobtn) == "undefined") {
        radiobtn = "1";
    }
    //AMS-685 enabling the box with PAB as well
    //document.getElementById('selector').style.display = (radiobtn == 1 ? 'block' : 'none');
    if (document.getElementById('divAttributeBuilder') != null) document.getElementById('divAttributeBuilder').style.display = (radiobtn == 1 ? 'none' : '');
}
function FlipUI() {
    /* OptionValue = 1 --> Represents "Basket level type Discount". 
                           In this case hide both radio buttons, hide Product Attribute builder 
                           and Show ProductGroup list with Exclude and Include lists. 
       OptionValue = 2 --> Means DiscountType is either "Group level Conditional" or "Item Level Conditional". 
                           In this case Hide both the radio buttons, hide Product Attribute builder 
                           and hide ProductGroup builder.
       ELSE            --> Show both the radio buttons 
    */
    var OptionValue = document.getElementById('discountType');
    if (OptionValue == null) {
        return;
    }
    //OptionValue.selectedItem.value;
    if (OptionValue.value == '3') {
        var PGEnabled = '<%= AttributePGEnabled %>';
        if (PGEnabled == 'True') {
            document.getElementById('RadioButtonList1').style.display = 'none';
            $("#RadioButtonList1 input[type=radio]:checked").val(1);
        }
        //AMS-685 commented as it is not required
        //document.getElementById('selector').style.display = 'block';
        if (document.getElementById('divAttributeBuilder') != null) document.getElementById('divAttributeBuilder').style.display = 'none';
    }
    else if (OptionValue.value == '4' || OptionValue.value == '5') {
        var PGEnabled = '<%= AttributePGEnabled %>';
              if (PGEnabled == 'True') {
                  document.getElementById('RadioButtonList1').style.display = 'none';
                  $("#RadioButtonList1 input[type=radio]:checked").val(1);
              }
              $("#selector").hide();
              $("#divAttributeBuilder").hide();
          }
          else {
              var PGEnabled = '<%= AttributePGEnabled %>';
              if (PGEnabled == 'True') {
                  document.getElementById('RadioButtonList1').style.display = 'block';
                  ProductGroupDivSelection();
              }
              else {
                    //AMS-685 
                  //document.getElementById('selector').style.display = 'block';
                  if (document.getElementById('divAttributeBuilder') != null) document.getElementById('divAttributeBuilder').style.display = 'none';
              }
          }
  }
  
  	    function HideMultiLangBox()
        {
           //This method is used for hiding multi language box when user doesnot explicitly click on 'X' button of multi langauage control and click on Save button directly.
             var MultiLanguageEnabled = '<%=IIf(SystemCacheData.GetSystemOption_General_ByOptionId(124) = "1", True, False)%>'
             if(MultiLanguageEnabled == 'True')
             {
                 var tier ='<%=TierLevels %>'
          
                  var levels = parseInt(tier);

                  for (var i=1;i<=levels;i++)
                  {
                    hideMultiLanguageInput('tier'+i+'_rdesc',event);
                  }
            }
        }	
</script>
<%
    Send_HeadEnd()
    If (IsTemplate) Then
        Send_BodyBegin(12)
    Else
        Send_BodyBegin(2)
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
        
		If (OfferID > 0 AndAlso (GetCgiValue("save") <> "" OrElse GetCgiValue("mode") = "savediscount")) Then
				' discounts disable the DeferCalcToEOS option of an offer
				MyCommon.QueryStr = "Select DeferCalcToEOS from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID & " and Deleted=0 and DeferCalcToEOS=1;"
				rst = MyCommon.LRT_Select
				If (rst.Rows.Count > 0) Then
						MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set DeferCalcToEOS=0 where IncentiveID=" & OfferID & ";"
						MyCommon.LRT_Execute()
						DeferCalcToEOSChanged = True
				End If
		End If
        
		Send("<script type=""text/javascript"">")
		Send("function ChangeParentDocument() { ")
		If (DeferCalcToEOSChanged) Then
				Send("  alert('" & Copient.PhraseLib.Lookup("cpeoffer-rew-disc.deferchanged", LanguageID) & "');")
		End If
		If (Phase = 3) Then
				Send("  if (opener != null) {")
				Send("    var newlocation = 'UEoffer-rew.aspx?OfferID=" & OfferID & "'; ")
				Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
				Send("  opener.location = 'UEoffer-rew.aspx?OfferID=" & OfferID & "'; ")
				Send("  }")
				Send("  }")
		ElseIf (Phase = 1) Then
				Send("  if (opener != null) {")
				Send("    var newlocation = 'UEoffer-not.aspx?OfferID=" & OfferID & "'; ")
				Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
				Send("  opener.location = 'UEoffer-not.aspx?OfferID=" & OfferID & "'; ")
				Send("  }")
				Send("  }")
		End If
    Send("} ")
    Send("</script>")
%>
<form action="#" id="mainform" name="mainform" autocomplete="off" onsubmit="return assignValues();"
    runat="server">
    <div id="divNotification" style="display:none"><label id="lblAjaxNotification" ></label></div>
    <div id="results" style="position: absolute; z-index: 99; top: 31px; right: 21px;">
    </div>
    <div id="intro">
        <input type="hidden" id="OfferID" name="OfferID" value="<% Sendb(OfferID)%>" />
        <input type="hidden" id="Name" name="Name" value="<% Sendb(Name)%>" />
        <input type="hidden" id="DeliverableID" name="DeliverableID" value="<% Sendb(DeliverableID)%>" />
        <input type="hidden" id="RewardID" name="RewardID" value="<% Sendb(RewardID)%>" />
        <input type="hidden" id="DiscountID" name="DiscountID" value="<% Sendb(DiscountID)%>" />
        <input type="hidden" id="Phase" name="Phase" value="<% Sendb(Phase)%>" />
        <input type="hidden" id="discountedpgid" name="discountedpgid" value="<% Sendb(DiscountedProductGroupID)%>"
            runat="server" />
        <input type="hidden" id="excludedpgid" name="excludedpgid" value="<% If excludedProductGroups IsNot Nothing AndAlso excludedProductGroups.Count > 0 Then Sendb(String.Join(",", excludedProductGroups.Select(Function(x) x.ProductGroupId).ToArray()))%>" />
        <input type="hidden" id="usergroupid" name="usergroupid" value="<% Sendb(UserGroupID)%>" />
        <input type="hidden" id="mode" name="mode" value="" />
        <input type="hidden" id="roid" name="roid" value="<%Sendb(TpROID)%>" />
        <input type="hidden" id="tp" name="tp" value="<%Sendb(TouchPoint)%>" />
        <input type="hidden" id="IsTemplate" name="IsTemplate" value="<% Sendb(IsTemplateVal)%>" />
        <input type="hidden" id="hdnflexnegative" name="hdnflexnegative" value="<%=flexOption%>" />
        <input type="hidden" id="hidDiscountType" name="hidDiscountType" value="0" runat="server" />
        <input type="hidden" id="hdnIsAttributeSwitch" name="hdnIsAttributeSwitch" value="0"
            runat="server" />
        <input type="hidden" id="hdnSwitchPGID" name="hdnSwitchPGID" value="0" runat="server" />
        <input type="hidden" id="hdnTemp" name="hdnTemp" value="0" runat="server" />
        <input id="NodeListID" type="hidden" name="NodeListID" value="" runat="server" />
        <input type="hidden" id="AttributeSwitchType" name="AttributeSwitchType" value="" />
        <input id="hdnLastDiscountedPG" type="hidden" runat="server" value="" />
		<input type="hidden" id="isDropdownChanged" name="isDropdownChanged" value="false" runat="server" />
        <%
            Dim lstSVPrograms As List(Of CMS.AMS.Models.SVProgram)
            HasAnyCustomer = UEOffer_Has_AnyCustomer(MyCommon, OfferID)
            If (HasAnyCustomer = True) Then
                lstSVPrograms = SVLib.GetStoredValueAllowAnyCustomerPrograms(True)
                For Each svprogram As CMS.AMS.Models.SVProgram In lstSVPrograms
                    sbSVOptions.AppendLine("<option value='" & svprogram.SVProgramID & "'>" & svprogram.ProgramName & "</option>")
                Next
            Else
                lstSVPrograms = SVLib.GetStoredValueMonetaryPrograms()
                For Each svprogram As CMS.AMS.Models.SVProgram In lstSVPrograms
                    sbSVOptions.AppendLine("<option value='" & svprogram.SVProgramID & "'>" & svprogram.ProgramName & "</option>")
                Next
            End If

            'lstSVPrograms=SVLib.GetStoredValueNonMonetaryPrograms()
            'For Each svprogram As CMS.AMS.Models.SVProgram In lstSVPrograms
            '  sbSVPointsOptions.AppendLine("<option value='" & svprogram.SVProgramID & "'>" & svprogram.ProgramName & "</option>")
            'Next
        %>
        <%
            If (IsTemplate) Then
                Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.discountreward", LanguageID), VbStrConv.Lowercase) & "</h1>")
            Else
                Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.discountreward", LanguageID), VbStrConv.Lowercase) & "</h1>")
            End If
        %>
        <div id="controls">
            <%
                If (IsTemplate) Then
                    Send("<span class=""temp"">")
                    Send("  <input type=""checkbox"" class=""tempcheck"" id=""Disallow_Edit"" onclick=""updateTableLockStatus();"" name=""Disallow_Edit""" & IIf(Disallow_Edit, " checked=""checked""", "") & " />")
                    Send("  <label for=""Disallow_Edit"">" & Copient.PhraseLib.Lookup("term.locked", LanguageID) & "</label>")
                    Send("  <a href=""javascript:ShowFieldList();"" title=""" & Copient.PhraseLib.Lookup("cpeoffer-rew-disc-clicktoview", LanguageID) & """> &#9660;</a>")
                    Send("</span>")
                End If
                If((Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable))
                    If Not IsTemplate Then
                        If (Logix.UserRoles.EditOffer And Not IsOfferWaitingForApproval(OfferID) And (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID)) And Not (FromTemplate And Disallow_Edit)) Or (Logix.UserRoles.EditRoles And OvrdFldEditable) Then
                          If (TierLevels > 1 And SupportGlobalAndTieredConditions = 1) Then   
                            Send_Save("onclick=""enableTiers(" & TierLevels & ");getIDList()""")
                          Else
                            Send_Save("onclick=""getIDList()""")
                          End If
                        End If
                    Else
                        If (Logix.UserRoles.EditTemplates) Or (Logix.UserRoles.EditRoles And OvrdFldEditable) Then
                          If (TierLevels > 1 And SupportGlobalAndTieredConditions = 1) Then  
                            Send_Save("onclick=""enableTiers(" & TierLevels & ");getIDList()""")
                          Else
                            Send_Save("onclick=""getIDList()""")
                          End If
                        End If
                    End If
                End If
            %>
        </div>
    </div>
    <div id="main">
        <%
            If (infoMessage <> "") Then
                Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
            End If
        %>
        <div>
            <br />
            <table>
                <tr>
                    <td width="30px">
                        <label for="discountType">
                            <% Sendb(Copient.PhraseLib.Lookup("term.discounttype", LanguageID))%>:</label>
                        <select id="discountType" name="discountType" size="1" onchange="handleDiscType_Change(this.value);"
                            <% Sendb(DisabledAttribute)%>>
                            <%
                                MyCommon.QueryStr = "select DiscountTypeID, Name, PhraseID from UE_DiscountTypes DT with (NoLock) order by DiscountTypeID;"
                                rst = MyCommon.LRT_Select()
                                If (rst.Rows.Count > 0) Then
'                                    If MyCommon.Fetch_UE_SystemOption(196) = 0 Then
                                    If SystemCacheData.GetSystemOption_UE_ByOptionId(196) = 0 Then
                                      rst.Rows(5).Delete() 'Remove Group Level if feature is disabled
                                      rst.AcceptChanges()
                                    End If
                                    For Each row In rst.Rows
                                        DiscTypeSel = IIf((DiscountType = MyCommon.NZ(row.Item("DiscountTypeID"), 0)), " selected=""selected""", "")
                                        Sendb("<option value=""" & MyCommon.NZ(row.Item("DiscountTypeID"), 0) & """" & DiscTypeSel & ">")
                                        If IsDBNull(row.Item("PhraseID")) Then
                                            Sendb(row.Item("Name") & Copient.PhraseLib.Lookup("term.unknown", LanguageID))
                                        Else
                                            If (row.Item("PhraseID") = 0) Then
                                                Sendb(row.Item("Name") & Copient.PhraseLib.Lookup("term.unknown", LanguageID))
                                            Else
                                                Sendb(Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID))
                                            End If
                                        End If
                                        Send("</option>")
                                    Next
                                End If
                                PgDisabled = IIf(DiscountType = 3, " disabled=""disabled""", "")
                                PgDblClick = IIf(DiscountType <> 3, "selectItem('functionselect', 'selected');", "")
                                DeslctPgDblClick = IIf(DiscountType <> 3, "deselectItem('selected');", "")
                            %>
                        </select>
                    </td>
                    <td>
                        <asp:radiobuttonlist repeatdirection="Horizontal" id="RadioButtonList1" runat="server"
                            clientidmode="static">
                        </asp:radiobuttonlist>
                    </td>
                </tr>
            </table>
        </div>
           <%           
                If Not IsPostBack OrElse (isDropdownChanged.Value = "1" AndAlso ucProductAttributeFilter.SelectedNodeIDs = "") Then
                    If AttributeProductGroupID <> 0 Then
                        ucProductAttributeFilter.HierarchyTreeURL = "/logix/phierarchytree.aspx?ProductGroupID=" & IIf(AttributeProductGroupID > 0, AttributeProductGroupID, 0) & "&PAB=1&OfferID=" & OfferID '& locateHierarchyURL
                    Else
                        ucProductAttributeFilter.HierarchyTreeURL = "/logix/phierarchytree.aspx?ProductGroupID=-1&PAB=1&OfferID=" & OfferID '& locateHierarchyURL
                    End If
                End If
            %>

        <div style="clear: both;">
        </div>
        <% If AttributePGEnabled Then%>
        <div class="box customcolumnfull" id="divAttributeBuilder" style="float: left; position: relative; width: 100%; height: auto;">
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.includedproducts", LanguageID))%>
                </span>
            </h2>
            <div style="float: left; position: relative; width: 100%; height: 8%; padding-left: 12px">
                <label>
                    Product Group Name:
                </label>
                <input id="txtProductGroupName" runat="server" type="text" style="width: 60%" value="" />
            </div>
            <div id="attributeSelector" style="float: left; position: relative; width: 100%; height: auto; min-height: 250px;">
                <uc1:ProductAttributeFilter ID="ucProductAttributeFilter" runat="server" AppName="UEoffer-rew-discount.aspx" />
            </div>
            <div id="divHierarchyContent" runat="server" style="float: left; position: relative; width: 100%; height: 93%;">
            </div>
        </div>
        <% End If%>

        <div style="clear: both;">
        </div>
        <div class="box customcolumnfull" id="selector">
            <div id="hidemydiv">
                <h2>
                    <span <% If AttributePGEnabled AndAlso RadioButtonList1.Items(1) IsNot Nothing AndAlso RadioButtonList1.Items(1).Selected = True Then Sendb(" style=""display:none""")%>>
                        <% Sendb(Copient.PhraseLib.Lookup("term.productgroup", LanguageID)) %>
                    </span>
                    <span <% If Not AttributePGEnabled  Or (AttributePGEnabled AndAlso RadioButtonList1.Items(0) IsNot Nothing AndAlso RadioButtonList1.Items(0).Selected = True) Then Sendb(" style=""display:none""") %>>
                        <% Sendb(Copient.PhraseLib.Lookup("term.excludedproductgroups", LanguageID)) %>
                    </span>
                </h2>
                <div id="prodgroupselector" <%Sendb(IIf(DiscountType = 4 OrElse DiscountType = 5, " style=""display:none;""", ""))%>>
                    <br />
                    <input type="radio" id="functionradio1" name="functionradio" <% If (SystemCacheData.GetSystemOption_General_ByOptionId(175) = "1") Then Sendb(" checked=""checked""")%>
                        <% Sendb(DisabledAttribute)%> /><label for="functionradio1"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
                    <input type="radio" id="functionradio2" name="functionradio" <% If (SystemCacheData.GetSystemOption_General_ByOptionId(175) = "2") Then Sendb(" checked=""checked""")%>
                        <% Sendb(DisabledAttribute)%> /><label for="functionradio2"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
                    <input type="text" class="medium" id="functioninput" name="functioninput" maxlength="100"
                        onkeyup="handleKeyUp(200);" value="" <% Sendb(DisabledAttribute)%> /><br />
                    <div class="customcolumn1">
                        <br />
                        <select style="width: 250px" id="functionselect" name="functionselect"
                            ondblclick="<% Sendb(PgDblClick)%>" size="10" <% Sendb(DisabledAttribute)%>>
                            <%              
                                If DiscountedProductGroupID = 1 AndAlso DiscountType = 3 Then
                                    ProductGroupSelOpt = "<option value=""1"">" & Copient.PhraseLib.Lookup("term.anyproduct", LanguageID) & "</option>"
                                    AnyProduct = True
                                ElseIf DiscountType = 3 Then
                                    ''Send("<option value=""1"">" & Copient.PhraseLib.Lookup("term.anyproduct", LanguageID) & "</option>")
                                    ProductGroupSelOpt = "<option value=""1"">" & Copient.PhraseLib.Lookup("term.anyproduct", LanguageID) & "</option>"
                                    AnyProduct = True
                                ElseIf DiscountType <> 3 AndAlso DiscountID = 0 Then
                                    AnyProduct = False
                                End If
                  
                                If IsPostBack Then
                                    If AnyProduct = False AndAlso Not String.IsNullOrWhiteSpace(hdnLastDiscountedPG.Value) Then
                                        DiscountedProductGroupID = hdnLastDiscountedPG.Value
                                    Else
                                        DiscountedProductGroupID = MyCommon.Extract_Decimal(GetCgiValue("discountedpgid"), MyCommon.GetAdminUser.Culture)
                                    End If                                                                        
                                End If
                                
                                rst = GetProductGroupListDataTable(MyCommon, "", AdminUserID, Logix.UserRoles.ViewProductgroupRegardlessBuyer, DiscountedProductGroupID, shouldFetchPGAsync)
                                
                                For Each row As DataRow In rst.Rows
                                    If MyCommon.NZ(row.Item("ProductGroupID"), 0) = DiscountedProductGroupID Then
                                        AnyProduct = False
                                        If (MyCommon.NZ(row.Item("Name"), "") <> "") Then
                                            If (MyCommon.IsEngineInstalled(9) And SystemCacheData.GetSystemOption_UE_ByOptionId(168) = "1" And Not IsDBNull(row.Item("Buyerid"))) Then
                                                Dim buyerid As Integer = row.Item("Buyerid")
                                                Dim externalBuyerid = MyCommon.GetExternalBuyerId(buyerid)
                                                ProductGroupSelOpt = "<option title=""" & MyCommon.NZ(row.Item("Name"), "") & """ value=""" & row.Item("ProductGroupID") & """ " & IIf(lstAttributePGIDs.Contains(row.Item("ProductGroupID")), "style=""color: blue;""", "") & ">" & "Buyer " & externalBuyerid & " - " & MyCommon.NZ(row.Item("Name"), "") & "</option>"
                                            Else
                                                ProductGroupSelOpt = "<option title=""" & MyCommon.NZ(row.Item("Name"), "") & """ value=""" & row.Item("ProductGroupID") & """ " & IIf(lstAttributePGIDs.Contains(row.Item("ProductGroupID")), "style=""color: blue;""", "") & ">" & MyCommon.NZ(row.Item("Name"), "") & "</option>"
                                            End If                                                        
                                        End If
                                    ElseIf excludedProductGroups IsNot Nothing AndAlso excludedProductGroups.Exists(Function(x) x.ProductGroupId = MyCommon.NZ(row.Item("ProductGroupID"), 0)) Then
                                        If (MyCommon.IsEngineInstalled(9) And SystemCacheData.GetSystemOption_UE_ByOptionId(168) = "1" And Not IsDBNull(row.Item("Buyerid"))) Then
                                            Dim buyerid As Integer = row.Item("Buyerid")
                                            Dim externalBuyerid = MyCommon.GetExternalBuyerId(buyerid)
                                            ExcludedPGSelOpt.AppendLine("<option title=""" & MyCommon.NZ(row.Item("Name"), "") & """ value=""" & row.Item("ProductGroupID") & """ " & IIf(lstAttributePGIDs.Contains(row.Item("ProductGroupID")), "style=""color: blue;""", "") & ">" & "Buyer " & externalBuyerid & " - " & MyCommon.NZ(row.Item("Name"), "") & "</option>")
                                        Else
                                            ExcludedPGSelOpt.AppendLine("<option title=""" & MyCommon.NZ(row.Item("Name"), "") & """ value=""" & row.Item("ProductGroupID") & """ " & IIf(lstAttributePGIDs.Contains(row.Item("ProductGroupID")), "style=""color: blue;""", "") & ">" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
                                        End If                                              
                                    Else
                                        If (MyCommon.NZ(row.Item("Name"), "") <> "") Then
                                            If (MyCommon.IsEngineInstalled(9) And SystemCacheData.GetSystemOption_UE_ByOptionId(168) = "1" And Not IsDBNull(row.Item("Buyerid"))) Then
                                                Dim buyerid As Integer = row.Item("Buyerid")
                                                Dim externalBuyerid = MyCommon.GetExternalBuyerId(buyerid)
                                                Send("<option  title=""" & MyCommon.NZ(row.Item("Name"), "") & """ value=""" & row.Item("ProductGroupID") & """ " & IIf(lstAttributePGIDs.Contains(row.Item("ProductGroupID")), "style=""color: blue;""", "") & ">" & "Buyer " & externalBuyerid & " - " & MyCommon.NZ(row.Item("Name"), "") & "</option>")
                                            Else
                                                Send("<option title=""" & MyCommon.NZ(row.Item("Name"), "") & """ value=""" & row.Item("ProductGroupID") & """ " & IIf(lstAttributePGIDs.Contains(row.Item("ProductGroupID")), "style=""color: blue;""", "") & ">" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
                                            End If
                           
                                        End If
                                    End If
                                Next
                            %>
                        </select>
                    </div>
                    <div class="customcolumn2">
                        <br />
                        <input type="button" class="regular select" id="pselect" name="pselect" value="<% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID))%> &#9658;"
                            onclick="selectItem('functionselect', 'selected'); " <% Sendb(DisabledAttribute)%><%Sendb(PgDisabled)%> <% If AttributePGEnabled AndAlso RadioButtonList1.Items(1) IsNot Nothing AndAlso RadioButtonList1.Items(1).Selected = True Then Sendb(" style=""display:none""")%>/>
                        <br />
                        <br />
                        <input type="button" class="regular deselect" id="pdeselect" name="pdeselect" value="&#9668; <% Sendb(Copient.PhraseLib.Lookup("term.deselect", LanguageID))%>"
                            onclick="deselectItem('selected', 'functionselect'); updateButtons();" <% Sendb(DisabledAttribute)%><%Sendb(PgDisabled)%> <% If AttributePGEnabled AndAlso RadioButtonList1.Items(1) IsNot Nothing AndAlso RadioButtonList1.Items(1).Selected = True Then Sendb(" style=""display:none""")%>/>
                        <br />
                        <br />
                        <br />
                        <% If ((AnyProduct AndAlso DiscountType = 3) Or (DiscountType = 0 Or DiscountType = 1 Or DiscountType = 2)) AndAlso (eDiscountType = 1 Or eDiscountType = 3 Or eDiscountType = 4) OrElse (DiscountType = 6) Then%>
                        <input type="button" class="regular select" name="select3" id="Button1" value="<% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID))%> &#9658;"
                            onclick="moveExcludedItem('functionselect', 'excluded', true); updateButtons();" <% Sendb(DisabledAttribute)%> />
                        <br />
                        <br />
                        <input type="button" class="regular deselect" name="deselect2" id="Button2" value="&#9668; <% Sendb(Copient.PhraseLib.Lookup("term.deselect", LanguageID))%>"
                            onclick="moveExcludedItem('excluded', 'functionselect', false); updateButtons();" <% Sendb(DisabledAttribute)%> />
                        <% End If%>
                    </div>
                    
                    <div class="customcolumn3">
                        <% 
                            If AttributePGEnabled AndAlso RadioButtonList1.Items(1) IsNot Nothing AndAlso RadioButtonList1.Items(1).Selected = True Then
                                excludeBoxSize = 10
                            Else
                                excludeBoxSize = 4
                            End If
                        %>
                        <div id="selectBox" <% If AttributePGEnabled AndAlso RadioButtonList1.Items(1) IsNot Nothing AndAlso RadioButtonList1.Items(1).Selected = True Then Sendb(" style=""display:none""")%>>
                        <b>
                        <% Sendb(Copient.PhraseLib.Lookup("term.includedgroups", LanguageID)) %>:
                        </b>
                        <select id="selected" name="selected" size="4" style="width:250px"
                            ondblclick="<% Sendb(PgDblClick)%>" <% Sendb(DisabledAttribute)%> >
                            <% Send(ProductGroupSelOpt)%>
                        </select>
                        <br />
                        <br />
                        </div>

                        <%--If discount Type is Basket Level--%>
                        <% If ((AnyProduct AndAlso DiscountType = 3) Or (DiscountType = 0 Or DiscountType = 1 Or DiscountType = 2)) AndAlso (eDiscountType = 1 Or eDiscountType = 3 Or eDiscountType = 4) Then%>
                        
                        <b><% Sendb(Copient.PhraseLib.Lookup("term.excludedgroups", LanguageID))%>:</b>
                        <select style="width: 245px" id="excluded" name="excluded" size="<%= excludeBoxSize %>" ondblclick="moveItem('excluded', 'functionselect');"
                            <% Sendb(DisabledAttribute)%>>
                            <% Send(ExcludedPGSelOpt.ToString())%>
                        </select>
                        <% End If%>

                    </div>
                    <hr class="hidden" />
                </div>
                <% Send_Proration_Box(MyCommon, ProrationTypeID, DiscountType, DisabledAttribute)%>
            </div>
            <div id="gutter">
            </div>
        </div>
        <div style="clear: both;">
        </div>

        <%
            Send("<div style=""clear: both;""></div>")
            Send("<div class=""box customcolumnfull"" id=""distribution"" style=""z-index:50;"">")
            Send("  <h2>")
            Send("    <span>")
            Send("      " & Copient.PhraseLib.Lookup("term.distribution", LanguageID))
            Send("    </span>")
            Send("  </h2>")
            'Output eDiscountType and AmountType
            Send("  <table summary=""" & Copient.PhraseLib.Lookup("term.distribution", LanguageID) & """>")
            'If eDiscountType = 2 Then
            '  Send("<tr>")
            '  Send("  <td>")
            '  Send("    <label for=""tier1_l1discountamt"">" & Copient.PhraseLib.Lookup("term.discountamount", LanguageID) & ":</label>")
            '  Send("  </td>")
            '  Send("  <td>")
            '  Send("    $<input type=""text"" id=""tier1_l1discountamt"" name=""tier1_l1discountamt"" value=""" & Math.Round(DiscountAmount, GetCurrencyPrecision(RewardID)) & """ size=""16"" maxlength=""16""" & DisabledAttribute & " />")
            '  Send("  </td>")
            '  Send("</tr>")
            '  Send("<tr>")
            '  Send("  <td>")
            '  Send("    <label for=""discountbarcode"">" & Copient.PhraseLib.Lookup("term.discountbarcode", LanguageID) & ":</label>")
            '  Send("  </td>")
            '  Send("  <td>")
            '  Send("    <input type=""text"" id=""discountbarcode"" name=""discountbarcode"" value=""" & DiscountBarcode & """ size=""30"" maxlength=""255""" & DisabledAttribute & " />")
            '  Send("  </td>")
            '  Send("</tr>")
            '  Send("<tr>")
            '  Send("  <td>")
            '  Send("    <label for=""voidbarcode"">" & Copient.PhraseLib.Lookup("term.voidbarcode", LanguageID) & ":</label>")
            '  Send("  </td>")
            '  Send("  <td>")
            '  Send("    <input type=""text"" id=""voidbarcode"" name=""voidbarcode"" value=""" & VoidBarcode & """ size=""30"" maxlength=""255""" & DisabledAttribute & " />")
            '  Send("  </td>")
            '  Send("</tr>")
            'Else
            If (eDiscountType = 4) Then
                Send_Amount_Type(AnyProduct, DiscountType, AmountTypeID, "1", TierLevels, IsMfgCoupon)
                If Copient.commonShared.Contains(AmountTypeID, 2, 6, 8, 13, 14, 15, 16) Then
                    Send_Allow_Markup(DiscountID)
                End If
            End If

            For t = 1 To TierLevels

                If GetCgiValue("tier" & t & "_l1discountamt") <> "" Then
                    DiscountAmount = MyCommon.Extract_Decimal(GetCgiValue("tier" & t & "_l1discountamt"), MyCommon.GetAdminUser.Culture)
                End If

                If TierLevels > 1 AndAlso AmountTypeID <> 7 Then
                    Send("<tr>")
                    Send("  <td><h3>" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</h3></td>")
                    If (SupportGlobalAndTieredConditions = 1 and t = 1) Then
                        Send("  <td><input type=""checkbox"" name=""useSameTierValue"" id =""useSameTierValue"" value=""1""" & IIf(UseSameTierValue = 1, " checked=""checked""", "") & " style=""margin-left:7px;"" onclick=""setSameTierValue(" & TierLevels & ")""/>")
                        Sendb("  <label for=""useThisValueForAllTiers"">" & Copient.PhraseLib.Lookup("term.UseThisValueForAllTiers", LanguageID) & "</label></td>")
                    End If

                    If (SupportGlobalAndTieredConditions = 1 and UseSameTierValue = 1) Then
                        DisabledAttribute=" disabled=""disabled"""
                        SetDisabledAttr(DisabledAttribute)
                    End If
                    Send("</tr>")
                End If
                If AmountTypeID <> 7 Then
                    If (eDiscountType = 1) Or (eDiscountType = 3) Or (eDiscountType = 4) Then
                        Send_Amount_Detail(DiscountID, AmountTypeID, DiscountAmount, "1", t, RewardID, MyCommon)
                        If TierLevels = 1 Then
                            Send_Amount_DetailLevels(AmountTypeID, DiscountAmount, L1Cap, False, "1", t, RewardID, MyCommon)
                        End If
                        If AmountTypeID = 3 And Math.Round(L1Cap, 2) > 0 Then
                            If TierLevels = 1 Then
                                Send_Amount_DetailLevels(L2AmountTypeID, L2DiscountAmt, L2Cap, False, "2", t, RewardID, MyCommon)
                            End If
                            If L2AmountTypeID = 3 And Math.Round(L2Cap, 2) > 0 Then
                                If TierLevels = 1 Then
                                    Send_Amount_DetailLevels(L3AmountTypeID, L3DiscountAmt, 0, True, "3", t, RewardID, MyCommon)
                                End If
                            End If
                        End If
                    End If
                End If

                If AmountTypeID = 8 Then
                    'Special pricing discount

                    'First find the SPRepeatLevel and ReceiptDescription for this tier.
                    If (infoMessage <> "") OrElse (GetCgiValue("mode") = "savediscount") OrElse (DiscountID = 0) Then
                        DiscountTierID = 0
                        SPRepeatLevel = MyCommon.Extract_Decimal(GetCgiValue("tier" & t & "_sprepeatlevel"), MyCommon.GetAdminUser.Culture)
                        If GetCgiValue("tier" & t & "_rdesc") <> "" Then
                            RDesc = GetCgiValue("tier" & t & "_rdesc")
                        Else
                            RDesc = ""
                        End If
                        If GetCgiValue("tier" & t & "_buydesc") <> "" Then
                            BuyDesc = GetCgiValue("tier" & t & "_buydesc")
                        Else
                            BuyDesc = ""
                        End If

                        ItemLimit = MyCommon.Extract_Decimal(GetCgiValue("tier" & t & "_itemlimit"), MyCommon.GetAdminUser.Culture)
                    Else
                        MyCommon.QueryStr = "select PKID as DiscountTierID, DT.SPRepeatLevel, DT.ReceiptDescription, DT.ItemLimit, DT.BuyDescription from CPE_DiscountTiers as DT with (NoLock) " & _
                                            "inner join CPE_Discounts as DISC with (NoLock) on DISC.DiscountID = DT.DiscountID " & _
                                            "where DT.DiscountID=" & DiscountID & " and DT.TierLevel=" & t & ";"
                        rst = MyCommon.LRT_Select
                        If rst.Rows.Count > 0 Then
                            DiscountTierID = MyCommon.NZ(rst.Rows(0).Item("DiscountTierID"), 0)
                            SPRepeatLevel = MyCommon.NZ(rst.Rows(0).Item("SPRepeatLevel"), 0)
                            RDesc = MyCommon.NZ(rst.Rows(0).Item("ReceiptDescription"), "")
                            BuyDesc = MyCommon.NZ(rst.Rows(0).Item("BuyDescription"), "")
                            ItemLimit = MyCommon.NZ(rst.Rows(0).Item("ItemLimit"), 0)
                        Else
                            DiscountTierID = 0
                            SPRepeatLevel = 0
                            RDesc = ""
                            BuyDesc = ""
                            ItemLimit = 0
                        End If
                    End If
                    If SPRepeatLevel = 0 Then
                        SPRepeatLevel = 1
                    End If
                    SPLevels = 0

                    Send("<tr>")
                    Send("  <th><small>" & Copient.PhraseLib.Lookup("term.restartpoint", LanguageID) & "</small></th>")
                    Send("  <th><small>" & Copient.PhraseLib.Lookup("term.price", LanguageID) & "</small></th>")
                    Send("</tr>")

                    'Next load up the special price details for the tier and display them.  This is done in one of two ways:
                    'by checking the query string (which we do if the offer hasn't been saved and needs to be redrawn), or
                    'by looking up the data from the table (which we do if the discount is saved and is being loaded afresh).
                    If (infoMessage <> "") OrElse (GetCgiValue("mode") = "savediscount") OrElse (DiscountID = 0) Then
                        SPLevels = MyCommon.Extract_Decimal(GetCgiValue("tier" & t & "_levels"), MyCommon.GetAdminUser.Culture)
                        SPHighestLevel = MyCommon.Extract_Decimal(GetCgiValue("tier" & t & "_highestlevel"), MyCommon.GetAdminUser.Culture)
                        For i = 1 To SPHighestLevel
                            If GetCgiValue("tier" & t & "_level" & i) <> Nothing Then
                                ValueString = GetCgiValue("tier" & t & "_level" & i)
                                If IsNumeric(ValueString) Then
                                    ValueString = Math.Round(CDec(ValueString), GetCurrencyPrecision(RewardID))
                                End If
                                LevelID = i
                                Send("<tr id=""tier" & t & "_level" & LevelID & "row"">")
                                Send("  <td>")
                                Send("    <input type=""radio"" id=""tier" & t & "_sprepeatlevel" & LevelID & """ name=""tier" & t & "_sprepeatlevel"" value=""" & LevelID & """" & IIf(LevelID = SPRepeatLevel, " checked=""checked""", "") & DisabledAttribute & " />")
                                Send("  </td>")
                                Send("  <td>" & GetCurrencySymbol(RewardID))
                                Send("    <input type=""text"" style=""width:183px;"" id=""tier" & t & "_level" & LevelID & """ name=""tier" & t & "_level" & LevelID & """ value=""" & ValueString & """" & DisabledAttribute & " />")
                                Send("    <input type=""button"" class=""ex"" name=""tier" & t & "_deletelevel" & LevelID & """ id=""tier" & t & "_deletelevel" & LevelID & """ value=""X"" onclick=""javascript:deleteLevel(" & t & ", " & LevelID & ");""" & DisabledAttribute & " />")
                                Send("  </td>")
                                Send("</tr>")
                            End If
                        Next
                    Else
                        MyCommon.QueryStr = "select Value, LevelID from CPE_SpecialPricing as SP with (NoLock) " & _
                                            "where SP.DiscountTierID=" & DiscountTierID & ";"
                        rst = MyCommon.LRT_Select
                        If (rst.Rows.Count > 0) Then
                            SPLevels = rst.Rows.Count
                            SPHighestLevel = rst.Rows.Count
                            For i = 0 To (rst.Rows.Count - 1)
                                Value = Math.Round(MyCommon.NZ(rst.Rows(i).Item("Value"), 0), GetCurrencyPrecision(RewardID))
                                LevelID = MyCommon.NZ(rst.Rows(i).Item("LevelID"), 0)
                                Send("<tr id=""tier" & t & "_level" & LevelID & "row"">")
                                Send("  <td>")
                                Send("    <input type=""radio"" id=""tier" & t & "_sprepeatlevel" & LevelID & """ name=""tier" & t & "_sprepeatlevel"" value=""" & LevelID & """" & IIf(LevelID = SPRepeatLevel, " checked=""checked""", "") & DisabledAttribute & " />")
                                Send("  </td>")
                                Send("  <td>" & GetCurrencySymbol(RewardID))
                                Send("    <input type=""text"" style=""width:183px;"" id=""tier" & t & "_level" & LevelID & """ name=""tier" & t & "_level" & LevelID & """ value=""" & Value.ToString(MyCommon.GetAdminUser.Culture) & """" & DisabledAttribute & " />")
                                Send("    <input type=""button"" class=""ex"" name=""tier" & t & "_deletelevel" & LevelID & """ id=""tier" & t & "_deletelevel" & LevelID & """ value=""X"" onclick=""javascript:deleteLevel(" & t & ", " & LevelID & ");""" & DisabledAttribute & " />")
                                Send("  </td>")
                                Send("</tr>")
                            Next
                        End If
                    End If
                    'Next draw 20 empty, hidden, disabled rows, which we'll reveal one by one if the user adds more levels
                    For i = 1 To 20
                        Send("<tr id=""tier" & t & "_level" & (SPLevels + i) & "row"" style=""display:none;"">")
                        Send("  <td>")
                        Send("    <input type=""radio"" id=""tier" & t & "_sprepeatlevel" & (SPLevels + i) & """ name=""tier" & t & "_sprepeatlevel"" value=""" & (SPLevels + i) & """" & IIf(i = 1 AndAlso DiscountID = 0, " checked=""checked""", "") & " disabled=""disabled"" />")
                        Send("  </td>")
                        Send("  <td>" & GetCurrencySymbol(RewardID))
                        Send("    <input type=""text"" style=""width:183px;"" id=""tier" & t & "_level" & (SPLevels + i) & """ name=""tier" & t & "_level" & (SPLevels + i) & """ value=""""" & DisabledAttribute & " disabled=""disabled"" /> " & GetCurrencyAbbr(RewardID))
                        Send("    <input type=""button"" class=""ex"" name=""tier" & t & "_deletelevel" & (SPLevels + i) & """ id=""tier" & t & "_deletelevel" & (SPLevels + i) & """ value=""X"" onclick=""javascript:deleteLevel(" & t & ", " & (SPLevels + i) & ");"" disabled=""disabled"" />")
                        Send("  </td>")
                        Send("</tr>")
                    Next
                    'New level line
                    Send("<tr id=""tier" & t & "_newlevelrow"">")
                    Send("  <td>")
                    Send("  </td>")
                    Send("  <td>" & GetCurrencySymbol(RewardID))
                    Send("    <input type=""text"" style=""width:183px;"" id=""tier" & t & "_newlevel"" name=""tier" & t & "_newlevel"" value="""" " & DisabledAttribute & " /> " & GetCurrencyAbbr(RewardID))
                    Send("    <input type=""button"" id=""tier" & t & "_addlevel"" name=""tier" & t & "_addlevel"" value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """  onclick=""javascript:addLevel(" & t & ");"" " & DisabledAttribute & " />")
                    Send("    <input type=""hidden"" id=""tier" & t & "_levels"" name=""tier" & t & "_levels"" value=""" & SPLevels & """ />")
                    Send("    <input type=""hidden"" id=""tier" & t & "_highestlevel"" name=""tier" & t & "_highestlevel"" value=""" & SPHighestLevel & """ />")
                    Send("   </td>")
                    Send("</tr>")
                    'Item limit line
                    'Send("<tr id=""tier" & t & "_itemlimitrow"">")
                    Send("  <td>")
                    Send("  </td>")
                    Send("  <td style=""padding-left:12px;"">")
                    Send("    <label for=""tier" & t & "_itemlimit""><small><b>" & Copient.PhraseLib.Lookup("term.itemlimit", LanguageID) & ":</b></small></label><br />")
                    Send("    <input type=""text"" id=""tier" & t & "_itemlimit"" name=""tier" & t & "_itemlimit"" value=""" & ItemLimit & """ size=""25"" maxlength=""4"" title=""" & Copient.PhraseLib.Lookup("term.itemlimitmsg", LanguageID) & """" & DisabledAttribute & " style=""width:183px""/>")
                    Send("  </td>")
                    Send("</tr>")
                    'Receipt text line
                    Send("<tr id=""tier" & t & "_receiptrow"">")
                    Send("  <td>")
                    Send("  </td>")
                    Send("  <td style=""padding-left:12px;"">")
                    Send("    <label for=""tier" & t & "_rdesc""><small><b>" & Copient.PhraseLib.Lookup("term.receipttext", LanguageID) & ":</b></small></label><br />")
                    'Send("    <input type=""text"" class=""medium"" id=""tier" & t & "_rdesc"" name=""tier" & t & "_rdesc"" maxlength=""18"" value=""" & MyCommon.NZ(RDesc, "").Replace("""", "&quot;") & """" & DisabledAttribute & " />")
                    MLI.ItemID = DiscountTierID
                    MLI.MLTableName = "CPE_DiscountTiersTranslations"
                    MLI.MLIdentifierName = "DiscountTiersID"
                    MLI.StandardTableName = "CPE_DiscountTiers"
                    MLI.StandardIdentifierName = "PKID"
                    MLI.MLColumnName = "ReceiptDesc"
                    MLI.StandardValue = MyCommon.NZ(RDesc, "").Replace("""", "&quot;")
                    MLI.InputName = "tier" & t & "_rdesc"
                    MLI.InputID = "tier" & t & "_rdesc"
                    MLI.InputType = "text"
                    MLI.LabelPhrase = ""
                    MLI.MaxLength = 18
                    MLI.CSSClass = ""
                    MLI.CSSStyle = "width:183px;"
                    MLI.Disabled = IIf(DisabledAttribute <> "", True, False)
                    Send(Localizer.SendTranslationInputs(MyCommon, MLI, Request.Form, 9))
                    Send("  </td>")
                    Send("</tr>")
                    ' buy description
                    If SystemCacheData.GetSystemOption_UE_ByOptionId(125) = "1" Then
                        Send("<tr id=""tier" & t & "_buyrow"">")
                        Send("  <td>")
                        Send("  </td>")
                        Send("  <td style=""padding-left:12px;"">")
                        Send("    <label for=""tier" & t & "_buydesc""><small><b>" & Copient.PhraseLib.Lookup("term.buydescription", LanguageID) & ":</b></small></label><br />")
                        'Send("    <input type=""text"" class=""medium"" id=""tier" & t & "_buydesc"" name=""tier" & t & "_buydesc"" maxlength=""18"" value=""" & MyCommon.NZ(BuyDesc, "").Replace("""", "&quot;") & """" & DisabledAttribute & " />")
                        MLI.ItemID = DiscountTierID
                        MLI.MLTableName = "CPE_DiscountTiersTranslations"
                        MLI.MLIdentifierName = "DiscountTiersID"
                        MLI.StandardTableName = "CPE_DiscountTiers"
                        MLI.StandardIdentifierName = "PKID"
                        MLI.MLColumnName = "BuyDesc"
                        MLI.StandardValue = BuyDesc.Replace("""", "&quot;")
                        MLI.InputName = "tier" & t & "_buydesc"
                        MLI.InputID = "tier" & t & "_buydesc"
                        MLI.InputType = "text"
                        MLI.LabelPhrase = ""
                        MLI.MaxLength = 18
                        MLI.CSSClass = ""
                        MLI.CSSStyle = "width:183px;"
                        MLI.Disabled = IIf(DisabledAttribute <> "", True, False)
                        Send(Localizer.SendTranslationInputs(MyCommon, MLI, Request.Form, 9))
                        Send("  </td>")
                        Send("</tr>")
                    End If

                ElseIf Not Copient.commonShared.Contains(AmountTypeID, 7, 8) Then
                    'All other kinds of discounts
                    If (Not (AnyProduct) And DeptLevel = 0 AndAlso DiscountType <> 6)  Then
                        Send_LimitsAndText(MyCommon, DiscountID, AmountTypeID, eDiscountType, "1", t, DisabledAttribute, RewardID, bAllowDollarTransLimit)
                    Else
                        Send_Text(MyCommon, DiscountID, AmountTypeID, eDiscountType, "1", t, DisabledAttribute)
                    End If

                End If
            Next

            Send_Reward_Required(MyCommon, DeliverableID, DisabledAttribute)

            Send("  </table>")
            Send("  <hr class=""hidden"" />")
            Send("</div>")

            ' Show the options box only if there are available options.
            If (Copient.commonShared.Contains(AmountTypeID, 5, 9, 10, 11, 12) OrElse DeptLevel = 0 OrElse DeptLevel = 1 OrElse ComputeDiscount = 1 OrElse ShowDiscPriceLevel) Then
                Send("<div style=""clear: both;""></div>")
                Send("<div class=""box customcolumnfull"" id=""options"">")
                Send("  <h2>")
                Send("    <span>")
                Send("      " & Copient.PhraseLib.Lookup("term.advancedoptions", LanguageID))
                Send("    </span>")
                Send("  </h2>")
                Send("  <table summary=""" & Copient.PhraseLib.Lookup("term.options", LanguageID) & """>")

                Send("<tr><td style=""width:170px;"">")
                '1.1 Chargeback department
                Sendb(Copient.PhraseLib.Lookup("term.chargebackdepartment", LanguageID) + ":</td>")
                If (eDiscountType = 1) Or (eDiscountType = 3) Or (eDiscountType = 4) Then
                    If AnyProduct Then
                        ' show item's department for basket-level discounts only when the proration type is set to item level (i.e. ID=1)
                        sQuery = "select ChargeBackDeptID, ExternalID, Name, PhraseID from ChargebackDepts with (NoLock) where ChargeBackDeptID not in (" & IIf(ProrationTypeID <> 1, "0,", "") & "10) "
                        GlobalDepts = IIf(ProrationTypeID <> 1, "0,", "-1")
                    ElseIf DeptLevel > 0 Then
                        ' show item's department for deparment-level discounts only when the proration type is set to item level (i.e. ID=1)
                        sQuery = "select ChargeBackDeptID, ExternalID, Name, PhraseID from ChargebackDepts with (NoLock) where ChargeBackDeptID not in (" & IIf(ProrationTypeID <> 1, "0,", "") & "10) "
                        GlobalDepts = "10"
                    Else 'item level
                        sQuery = "select ChargeBackDeptID, ExternalID, Name, PhraseID from ChargebackDepts with (NoLock) where ChargeBackDeptID not in (10) "
                        GlobalDepts = "0"
                    End If
                    If (BannersEnabled) Then
                        MyCommon.QueryStr = "select BO.BannerID, BAN.AllBanners from BannerOffers BO with (NoLock) " & _
                                            "inner join  Banners BAN with (NoLock) on BAN.BannerID = BO.BannerID and BAN.Deleted=0 " & _
                                            "where OfferID = " & OfferID
                        rst = MyCommon.LRT_Select
                        AllBanners = (rst.Rows.Count = 1 AndAlso MyCommon.NZ(rst.Rows(0).Item("AllBanners"), False))
                        If (rst.Rows.Count = 1 AndAlso Not AllBanners) Then
                            sQuery &= " and BannerID = " & MyCommon.NZ(rst.Rows(0).Item("BannerID"), -1) & " or ChargebackDeptID in (" & GlobalDepts & ")"

                        Else
                            sQuery &= " and (BannerID = 0 or BannerID IS NULL) " & " or ChargebackDeptID in (" & GlobalDepts & ")"
                        End If
                    End If
                    MyCommon.QueryStr = sQuery & " order by ExternalID;"
                    rst = MyCommon.LRT_Select

                    Send("<input type=""hidden"" name=""loadDefaultChargeback"" id=""loadDefaultChargeback"" value=""" & IIf(LoadDefaultChargeback, "1", "0") & """ />")
                    Send("<td><select id=""chargeback"" name=""chargeback"" style=""width:175px;"" onchange=""handleChargebackSubmit();""" & DisabledAttributeAO & ">")
                    If Not (rst.Rows.Count = 0) Then
                        If Not BannersEnabled And DiscountID = 0 Then
                            Select Case DiscountType
                                Case 0, 1 ' unset or item level
                                    DefaultChrgBack = MyCommon.Extract_Decimal(SystemCacheData.GetSystemOption_UE_ByOptionId(116), MyCommon.GetAdminUser.Culture)
                                Case 2 ' dept level
                                    DefaultChrgBack = MyCommon.Extract_Decimal(SystemCacheData.GetSystemOption_UE_ByOptionId(117), MyCommon.GetAdminUser.Culture)
                                Case 3 ' basket level
                                    DefaultChrgBack = MyCommon.Extract_Decimal(SystemCacheData.GetSystemOption_UE_ByOptionId(118), MyCommon.GetAdminUser.Culture)
                            End Select
                            Dim OfferType As Integer = m_OfferService.GetOfferType(OfferID)
                            If OfferType <> -1 Then
                                If OfferType = 1 Then
                                    DefaultChrgBack = 14 'Default ChargeBackDept should be Tender for Offer applied as Manufacturer Coupon
                                ElseIf OfferType = 2 Then
                                    DefaultChrgBack = 0 'Default ChargeBackDept should be Item's Department for Offer applied as Store Coupon
                                End If
                            End If
                        End If
                        Dim OptionText As String
                        For Each row In rst.Rows
                            OptionText = ""
                            If ((row.Item("ExternalID") = "") Or (row.Item("ExternalID") = "0")) Then
                            Else
                                OptionText = (row.Item("ExternalID") & " - ")
                            End If
                            If (IsDBNull(row.Item("PhraseID"))) Then
                                OptionText &= (MyCommon.NZ(row.Item("Name"), ""))
                            Else
                                If (row.Item("PhraseID") = 0) Then
                                    OptionText &= (MyCommon.NZ(row.Item("Name"), ""))
                                Else
                                    OptionText &= (Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID))
                                End If
                            End If
                            Sendb("<option title=""" & OptionText & """ value=""" & MyCommon.NZ(row.Item("ChargeBackDeptID"), 0) & """")
                            If MyCommon.NZ(row.Item("ChargeBackDeptID"), -1) = ChargebackDeptID Then
                                Sendb(" selected=""selected""")
                            End If
                            Sendb(">")

                            Sendb(TruncateWordAppendEllipsis(OptionText, 40))
                            Send("</option>")
                        Next
                    End If
                    Send("</select></td>")
                Else
                    Send("<input type=""hidden"" id=""chargeback"" name=""chargeback"" value=""1"" />")
                End If

                Send("<td>")
                Send("  <input type=""checkbox"" id=""grossprice"" name=""grossprice""  " & IIf(grossPrice = True, " checked=""checked""", "") & "  " & DisabledAttribute & " onclick=""assignDiscPriceLevel(this);""  />")
                Send("  <label for=""grossprice"">")
                Send(Copient.PhraseLib.Lookup("term.EnableGrossPrice", LanguageID))
                Send("  </label>")
                Send("</td>")

                Send("</tr>")

                If (restrictRewardforRPOS OrElse DiscountType = 4 OrElse DiscountType = 5) Then
                    '3.1 Discount price level
                    If ShowDiscPriceLevel Then
                        Send("<tr><td>")
                        Send("    <label for=""discpricelevel"">" & Copient.PhraseLib.Lookup("term.discountpricelevel", LanguageID) & ":</label></td>")
                        Send("<td>")
                        Send("    <select id=""discpricelevel"" name=""discpricelevel"" style=""width:175px;"">")
                        Send("      <option value=""0""" & IIf(DiscAtOrigPrice = 0, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("reward.discount-salesprice", LanguageID) & "</option>")
                        Send("      <option value=""1""" & IIf(DiscAtOrigPrice = 1, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("reward.discount-originalprice", LanguageID) & "</option>")
                        Send("    </select>")
                        Send("  </td>")
                    Else
                        ' default to discount at sales price (value of 0)
                        Send("  <tr><td colspan=""2""><input type=""hidden"" id=""discpricelevel"" name=""discpricelevel"" value=""0"" /></td>")
                    End If
                    '2.2 Stored Value            
                    If ((Not HasAnyCustomer) Or (HasAnyCustomer = True AndAlso lstSVPrograms.Count > 0)) Then
                        Send("<td>")
                        StoredValueDiv()
                        Send("</td>")
                    End If
                    Sendb("  </tr>")

                Else
                    '2.1 Price filter
                    Send("  <tr>")
                    Sendb("    <td>")
                    Send("<label for=""priceFilter"">" & Copient.PhraseLib.Lookup("pricefilter.selectedproducts", LanguageID) & ":</label><br />")
                    Send("     </td>")
                    Sendb("    <td>")
                    Send("<select id=""priceFilter"" name=""priceFilter"" style=""width:175px;"" " & DisabledAttribute & " >")
                    MyCommon.QueryStr = "SELECT  PriceFilterID, Description, PhraseID, DisplayOrder   FROM dbo.UE_PriceFilter upf  with (NoLock) UNION " &
                                         " SELECT ClearanceLevelValue as PriceFilterID, Description, PhraseID ,999 AS  DisplayOrder FROM dbo.UE_ClearanceLevels ucl with (NoLock) ORDER BY   DisplayOrder"
                    rst2 = MyCommon.LRT_Select
                    For Each row2 As DataRow In rst2.Rows
                        Send("  <option value=""" & MyCommon.NZ(row2.Item("PriceFilterID"), 0).ToString & """" & If(PriceFilterID = MyCommon.NZ(row2.Item("PriceFilterID"), 100), " selected=""selected""", "") & ">")
                        Send(Copient.PhraseLib.Lookup(MyCommon.NZ(row2.Item("PhraseID"), 0), LanguageID, MyCommon.NZ(row2.Item("Description"), "Unknown")))
                        Send("  </option>")
                    Next
                    Send("</select>")
                    Sendb("    </td>")

                    '2.2 Stored Value            
                    If ((Not HasAnyCustomer) Or (HasAnyCustomer = True AndAlso lstSVPrograms.Count > 0)) Then
                        Send("<td>")
                        StoredValueDiv()
                        Send("</td>")
                    End If
                    Sendb("  </tr>")

                    '3.1 Discount price level
                    If ShowDiscPriceLevel Then
                        Send("<tr><td>")
                        Send("    <label for=""discpricelevel"">" & Copient.PhraseLib.Lookup("term.discountpricelevel", LanguageID) & ":</label></td>")
                        Send("<td>")
                        Send("    <select id=""discpricelevel"" name=""discpricelevel"" style=""width:175px;"">")
                        Send("      <option value=""0""" & IIf(DiscAtOrigPrice = 0, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("reward.discount-salesprice", LanguageID) & "</option>")
                        Send("      <option value=""1""" & IIf(DiscAtOrigPrice = 1, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("reward.discount-originalprice", LanguageID) & "</option>")
                        Send("    </select>")
                        Send("  </td></tr>")
                    Else
                        ' default to discount at sales price (value of 0)
                        Send("  <tr><td colspan=""2""><input type=""hidden"" id=""discpricelevel"" name=""discpricelevel"" value=""0"" /></td></tr>")
                    End If
                End If
                Send("<tr>")
                '1.2 Flex Negative drop-down
                Sendb("<td><label for=""flexnegative"">")
                Sendb(Copient.PhraseLib.Lookup("reward.discount-flexnegative", LanguageID) & " " & Copient.PhraseLib.Lookup("term.options", LanguageID) & ":")
                Send("</label></td>")
                Sendb("    <td >")
                Sendb("			<select id=""flexnegative"" name=""flexnegative"" style=""width:175px;""" & DisabledAttribute & " >")

                If flexOption = 0 AndAlso FlexNeg Then
                    flexOption = 1
                End If

                rst2 = m_DiscountRewardService.GetFlexOptions(LanguageID)
                For Each row2 As DataRow In rst2.Rows
                    Send("  <option value=""" & MyCommon.NZ(row2.Item("FlexOptionID"), -1).ToString & """" & If(flexOption = MyCommon.NZ(row2.Item("FlexOptionID"), -1), " selected=""selected""", "") & ">")
                    Send(MyCommon.NZ(row2.Item("Phrase"), "Unknown"))
                    Send("  </option>")
                Next
                Send("</select>")
                Send("</td></tr>")
                Send("  </table>")
                Send("  <hr class=""hidden"" />")
                Send("</div>")
            End If
        %>
        <div style="clear: both;">
        </div>
        <% If DiscountedProductGroupID <> 0 OrElse (DiscountType = 4 OrElse DiscountType = 5) Then%>
        <div class="box customcolumnfull" id="scorecards">
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.scorecard", LanguageID))%>
                </span>
            </h2>
            <div id="scorecardinputs">
                <%
                  Send("<table summary=""" & Copient.PhraseLib.Lookup("term.scorecard", LanguageID) & """>")
                  Send("  <tr>")
                  Send("    <td>")
                  Send("      <label for=""ScorecardID"">" & Copient.PhraseLib.Lookup("CPEoffer-rew-discount.IncludeOnScorecard", LanguageID) & ":</label>")
                  Send("    </td>")
                  Send("    <td>")
                  MyCommon.QueryStr = "select ScorecardID, Description, EngineID, DefaultForEngine from Scorecards " & _
                            "where ScorecardTypeID=3 and Deleted=0 and EngineID=" & EngineID & ";"
                  rst = MyCommon.LRT_Select
                  Send("      <select class=""medium"" id=""ScorecardID"" name=""ScorecardID"" onchange=""toggleScorecardText();"">")
                  Send("        <option value=""0"">" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "</option>")
                  If rst.Rows.Count > 0 Then
                    For Each row In rst.Rows
                      If MyCommon.NZ(row.Item("DefaultForEngine"), False) = True Then 'Show ScorecardDesc if there is a default scorecard
                        DefaultExists = True
                      End If
                      If (ScorecardID = 0 AndAlso DiscountID <> 0) Then
                        Send("        <option value=""" & MyCommon.NZ(row.Item("ScorecardID"), 0) & """>" & MyCommon.NZ(row.Item("Description"), "") & "</option>")
                      ElseIf (MyCommon.NZ(row.Item("ScorecardID"), 0) = ScorecardID) Then
                        Send("        <option value=""" & MyCommon.NZ(row.Item("ScorecardID"), 0) & """ selected=""selected"">" & MyCommon.NZ(row.Item("Description"), "") & "</option>")
                      ElseIf (MyCommon.NZ(row.Item("DefaultForEngine"), False) = True) AndAlso (MyCommon.NZ(row.Item("EngineID"), -1) = EngineID AndAlso DiscountID = 0) Then
                        Send("        <option value=""" & MyCommon.NZ(row.Item("ScorecardID"), 0) & """ selected=""selected"">" & MyCommon.NZ(row.Item("Description"), "") & "</option>")
                      Else
                        Send("        <option value=""" & MyCommon.NZ(row.Item("ScorecardID"), 0) & """>" & MyCommon.NZ(row.Item("Description"), "") & "</option>")
                      End If
                    Next
                  End If
                  Send("      </select>")
                  Send("    </td>")
                  Send("  </tr>")
                  Dim blnShowRemarks As Boolean = True
                  If DefaultExists = True Then
                    If (ScorecardID = 0 AndAlso DiscountID = 0) Then
                      blnShowRemarks = True
                    ElseIf (ScorecardID = 0 AndAlso DiscountID <> 0) Then
                      blnShowRemarks = False
                    Else
                      blnShowRemarks = True
                    End If
                  Else
                    If (ScorecardID = 0 AndAlso DiscountID = 0) Then
                      blnShowRemarks = False
                    ElseIf (ScorecardID <> 0 AndAlso DiscountID <> 0) Then
                      blnShowRemarks = True
                    Else
                      blnShowRemarks = False
                    End If
                  End If
                  Send("  <tr id=""ScorecardDescLine"" " & IIf((blnShowRemarks = False), " style=""display:none;""", "") & " >")
                  Send("    <td>")
                  Send("      <label for=""ScorecardDesc"">" & Copient.PhraseLib.Lookup("term.scorecardtext", LanguageID) & ":</label>")
                  Send("    </td>")
                  Send("    <td>")
                  SetScoreCardMLI(DiscountID, MLI) 'Set the MLI properties for scorecard 
                  Send(Localizer.SendTranslationInputs(MyCommon, MLI, Request.Form, 9))
                  Send("    </td>")
                  Send("  </tr>")
                  Send("</table>")
                %>
            </div>
            <hr class="hidden" />
        </div>
        <% End If%>
    </div>
</form>
<script runat="server">
    Dim DisabledAttr As String = ""
    Const DISC_ORIGINAL_PRICE As Integer = 1
    Function GetProductGroupName() As String
        Dim ProductGroupName As String = String.Empty
        MyCommon.QueryStr = "select DiscountedProductGroupID from CPE_Discounts with (NoLock) where DiscountID=" & DiscountID
        rst = MyCommon.LRT_Select

        If discountedpgid.Value.ConvertToLong() <> 0 And m_ProductGroupService.GetProductGroupType(discountedpgid.Value.ConvertToLong()).Result = 2 And hdnSwitchPGID.Value.ConvertToLong() = 0 Then
            ProductGroupName = ""
        Else
            If hdnSwitchPGID.Value.ConvertToLong() <> 0 And m_ProductGroupService.GetProductGroupType(hdnSwitchPGID.Value.ConvertToLong()).Result = 2 Then
                If GetCgiValue("txtProductGroupName") = m_ProductGroupService.GetProductGroupName(hdnSwitchPGID.Value.ConvertToLong()).Result And hdnIsAttributeSwitch.Value = 1 Then
                    ProductGroupName = GetCgiValue("txtProductGroupName")
                ElseIf GetCgiValue("txtProductGroupName") <> "" AndAlso GetCgiValue("txtProductGroupName") <> m_ProductGroupService.GetProductGroupName(hdnSwitchPGID.Value.ConvertToLong()).Result AndAlso (hdnIsAttributeSwitch.Value <> 1 Or GetCgiValue("save") <> "") Then
                    ProductGroupName = GetCgiValue("txtProductGroupName")
                Else
                    ProductGroupName = m_ProductGroupService.GetProductGroupName(hdnSwitchPGID.Value.ConvertToLong()).Result
                End If
            Else
                If GetCgiValue("txtProductGroupName") <> ""  Then
                    ProductGroupName = GetCgiValue("txtProductGroupName")
                ElseIf hdnIsAttributeSwitch.Value <> 1 Then
                    If (rst.Rows.Count > 0) Then
                        If m_ProductGroupService.GetProductGroupType(Convert.ToInt64(IIf(IsDBNull(rst.Rows(0)(0)), 0, rst.Rows(0)(0)))).Result = 2 Then
                            ProductGroupName = m_ProductGroupService.GetProductGroupName(Convert.ToInt64(rst.Rows(0)(0))).Result
                        End If
                    Else
                        If m_ProductGroupService.GetProductGroupType(IIf(hdnSwitchPGID.Value.ConvertToLong() <> 0, hdnSwitchPGID.Value.ConvertToLong(), discountedpgid.Value.ConvertToLong())).Result = 2 Then
                            ProductGroupName = m_ProductGroupService.GetProductGroupName(IIf(hdnSwitchPGID.Value.ConvertToLong() <> 0, hdnSwitchPGID.Value.ConvertToLong(), discountedpgid.Value.ConvertToLong())).Result
                        End If
                    End If

                End If

            End If
        End If
        If ProductGroupName = "" Then
            ProductGroupName = String.Concat(Copient.PhraseLib.Lookup("term.offer", LanguageID), " ", OfferID.ToString(), " ", Copient.PhraseLib.Lookup("term.discountreward", LanguageID).ToLower())
        End If
        Return ProductGroupName
    End Function

    Public Shared Function PrepareProductGroupHTML(ByRef MyCommon As Copient.CommonInc, ByRef dtPG As DataTable) As String
        CurrentRequest.Resolver.AppName = "UEoffer-rew-discount.aspx"
        Dim SystemCacheData As ICacheData = CurrentRequest.Resolver.Resolve(Of ICacheData)()

        Dim sbOptions As New StringBuilder()
        For Each pgRow In dtPG.Rows
            If (MyCommon.NZ(pgRow.Item("Name"), "") <> "") Then
                If (MyCommon.IsEngineInstalled(9) And SystemCacheData.GetSystemOption_UE_ByOptionId(168) = "1" And Not IsDBNull(pgRow.Item("Buyerid"))) Then
                    Dim buyerid As Integer = pgRow.Item("Buyerid")
                    Dim externalBuyerid = MyCommon.GetExternalBuyerId(buyerid)
                    sbOptions.AppendLine("<option  title=""" & MyCommon.NZ(pgRow.Item("Name"), "") & """ value=""" & pgRow.Item("ProductGroupID") & """ " & IIf(pgRow.Item("ProductGroupTypeId") = 2, "style=""color: blue;""", "") & ">" & "Buyer " & externalBuyerid & " - " & MyCommon.NZ(pgRow.Item("Name"), "") & "</option>")
                Else
                    sbOptions.AppendLine("<option title=""" & MyCommon.NZ(pgRow.Item("Name"), "") & """ value=""" & pgRow.Item("ProductGroupID") & """ " & IIf(pgRow.Item("ProductGroupTypeId") = 2, "style=""color: blue;""", "") & ">" & MyCommon.NZ(pgRow.Item("Name"), "") & "</option>")
                End If

            End If
        Next
        Return sbOptions.ToString
    End Function
    <WebMethod>
    Public Shared Function GetProductGroupListJSON(ByVal lastPGName As String, ByVal AdminUserID As Integer, ByVal viewProductgroupRegardlessBuyer As Boolean) As String
        Dim MyCommon As New Copient.CommonInc()
        Return GetProductGroupListHTML(MyCommon, lastPGName, AdminUserID, viewProductgroupRegardlessBuyer)
    End Function

    Public Shared Function GetProductGroupListHTML(ByRef MyCommon As Copient.CommonInc, ByVal lastPGName As String, ByVal AdminUserID As Integer, ByVal viewProductgroupRegardlessBuyer As Boolean) As String
        Dim dt As DataTable = GetProductGroupListDataTable(MyCommon, lastPGName, AdminUserID, viewProductgroupRegardlessBuyer,-1, true)
        Return PrepareProductGroupHTML(MyCommon, dt)
    End Function

    Public Shared Function GetProductGroupListDataTable(ByRef MyCommon As Copient.CommonInc, ByVal lastPGName As String, ByVal AdminUserID As Integer, ByVal viewProductgroupRegardlessBuyer As Boolean, ByVal discountedPGId As Integer, ByVal fetchPGAsync As Boolean) As DataTable
        CurrentRequest.Resolver.AppName = "UEoffer-rew-discount.aspx"
        Dim SystemCacheData As ICacheData = CurrentRequest.Resolver.Resolve(Of ICacheData)()

        Dim listSize As Integer = SystemCacheData.GetSystemOption_General_ByOptionId(290)
        If (fetchPGAsync) Then
            MyCommon.QueryStr = "Select ProductGroupID,Buyerid, Name, PhraseID, ProductGroupTypeId from " & _
                                " (Select top " & listSize & " ProductGroupID,Buyerid, Name, PhraseID, ProductGroupTypeId from ProductGroups with (NoLock) where deleted=0 and ProductGroupID <> 1 and NonDiscountableGroup=0 and PointsNotApplyGroup=0 "

            If (MyCommon.IsUserAssociatedWithAnyBuyer(AdminUserID) AndAlso Not viewProductgroupRegardlessBuyer) Then
                MyCommon.QueryStr &= " and (BuyerId in(select BuyerId from BuyerRoleUsers where AdminUserID=" & AdminUserID & ") or BuyerId is null)"
            End If
            MyCommon.QueryStr &= " and Name > '" & lastPGName & "'"
            MyCommon.QueryStr &= " order by Name) as Temp"

            If fetchPGAsync AndAlso discountedPGId > 0 Then
                MyCommon.QueryStr &= " Union Select ProductGroupID,Buyerid, Name, PhraseID, ProductGroupTypeId from ProductGroups with (NoLock) where deleted=0 and ProductGroupId = " & discountedPGId
            End If
            MyCommon.QueryStr &= " order by Name"
        Else
            MyCommon.QueryStr = "Select ProductGroupID,Buyerid, Name, PhraseID, ProductGroupTypeId from ProductGroups with (NoLock) where deleted=0 and ProductGroupID <> 1 and NonDiscountableGroup=0 and PointsNotApplyGroup=0 "
            If (MyCommon.IsUserAssociatedWithAnyBuyer(AdminUserID) AndAlso Not viewProductgroupRegardlessBuyer) Then
                MyCommon.QueryStr &= " and (BuyerId in(select BuyerId from BuyerRoleUsers where AdminUserID=" & AdminUserID & ") or BuyerId is null)"
            End If
            MyCommon.QueryStr &= " order by Name"
        End If

        Return MyCommon.LRT_Select
    End Function
    Sub SetDisabledAttr(ByVal str As String)
        DisabledAttr = str
    End Sub

    Sub StoredValueDiv()
        Send("  <div id=""discountsvDiv"" name=""discountsvDiv"">")
        Send("  <input type=""checkbox"" id=""discountsv"" name=""discountsv"" value=""1"" " & IIf(SVProgramID > 0 OrElse bSVProgram, "checked", "") & " " & DisabledAttribute & " onchange=""ChangeDiscountSV(this);"" />")
        Send("  <label for=""discountsv"">")
        Send(PhraseLib.Lookup("term.storedvalue", LanguageID))
        Send("  </label>")
        Send("  <select id=""ddldiscountsv"" name=""ddldiscountsv"" size=""1"" style=""display: " & IIf(SVProgramID = 0 AndAlso bSVProgram = False, "none", "") & "; width: 150px;"" " & DisabledAttribute & ">")
        Dim str As String = sbSVOptions.ToString().Replace("value='" & SVProgramID & "'", "value='" & SVProgramID & "'  selected")
        Send(str)
        Send("  </select>")
        Send("  </div>")
    End Sub

    Sub Send_Amount_Type(ByVal AnyProduct As Boolean, ByVal DiscountType As Integer, ByVal AmountTypeID As Integer, _
                         ByVal Level As String, ByVal TierLevels As Integer, ByVal IsMfgCoupon As Boolean)
        Dim MyCommon As New Copient.CommonInc
        Dim rst As DataTable
        Dim row As DataRow
        Dim UOMCriteria As String = ""
        Const UOM_ALWAYS As Integer = -1
        Const UOM_DISABLED_ONLY As Integer = 0
        Const UOM_ENABLED_ONLY As Integer = 1
        Const UOM_OPTION_ID As Integer = 135

        MyCommon.Open_LogixRT()

        Send("<tr>")
        Sendb("  <td style=""width:70px;""><label for=""l" & Level & "amounttypeid"">" & Copient.PhraseLib.Lookup("term.type", LanguageID))
        If Level > 1 Then
            Sendb("&nbsp;" & Level)
        End If
        Send(":</label></td>")
        Send("  <td>")
        Send("    <select id=""l" & Level & "amounttypeid"" name=""l" & Level & "amounttypeid"" onchange=""submitForm();""" & DisabledAttr & " style=""width:187px; margin-left:10px;"">")

        ' determines which group of amount types to display depending on whether or not UOM is enabled.
        UOMCriteria = "(MultiUOMState = " & UOM_ALWAYS & " or MultiUOMState = " & IIf(SystemCacheData.GetSystemOption_UE_ByOptionId(UOM_OPTION_ID) = "1", UOM_ENABLED_ONLY, UOM_DISABLED_ONLY) & ")"

        ' If AnyProduct Or DeptLevel > 0 Then
        If AnyProduct And DiscountType = 3 Then
            MyCommon.QueryStr = "select AmountTypeID, PhraseID from CPE_AmountTypes AT with (NoLock) where " & UOMCriteria & " and AmountTypeID IN (1, 7" & IIf(IsMfgCoupon, "", ",3") & ")"
        ElseIf DiscountType = 2 Then
            MyCommon.QueryStr = "select AmountTypeID, PhraseID from CPE_AmountTypes AT with (NoLock) where " & UOMCriteria & " and AmountTypeID IN (1" & IIf(IsMfgCoupon, "", ",3") & ")"
        ElseIf DiscountType = 4 Then
            MyCommon.QueryStr = "select AmountTypeID, PhraseID from CPE_AmountTypes AT with (NoLock) where " & UOMCriteria & " and AmountTypeID Not IN (4,5,6,7,8,9,10,11,12,13,14,15,16)"
        ElseIf DiscountType = 6 Then
            MyCommon.QueryStr = "select AmountTypeID, PhraseID from CPE_AmountTypes AT with (NoLock) where " & UOMCriteria & " and AmountTypeID = 1"
        Else
            MyCommon.QueryStr = "select AmountTypeID, PhraseID from CPE_AmountTypes AT with (NoLock) where " & UOMCriteria & IIf(TierLevels > 1, " and AmountTypeID <> 7", "")
        End If

        ' ensure that if an AmountTypeID already saved for this discount is not returned, that it is included.
        MyCommon.QueryStr &= " union select AmountTypeID, PhraseID from CPE_AmountTypes with (NoLock) where AmountTypeID=" & AmountTypeID

        rst = MyCommon.LRT_Select
        If Not (rst.Rows.Count = 0) Then
            If AmountTypeID = 0 Then
                AmountTypeID = 1
            End If
            For Each row In rst.Rows
                'BZ2079: UE-feature-removal #: remove stored value from the amount type for UE.
                'To restore prior functionality:  remove the lines below commented with BZ2079A and BZ2079B
                If MyCommon.NZ(row.Item("AmountTypeID"), 0) <> 7 Then ' BZ2079A
                    Sendb("      <option value=""" & MyCommon.NZ(row.Item("AmountTypeID"), 0) & """")
                    If MyCommon.NZ(row.Item("AmountTypeID"), 0) = AmountTypeID Then
                        Sendb(" selected=""selected""")
                    End If
                    Sendb(">")
                    If IsDBNull(row.Item("PhraseID")) Then
                        Sendb(MyCommon.NZ(row.Item("Name"), ""))
                    Else
                        If (row.Item("PhraseID") = 0) Then
                            Sendb(MyCommon.NZ(row.Item("Name"), ""))
                        Else
                            Sendb(Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID))
                        End If
                    End If
                    Send("</option>")
                End If ' BZ2079B
            Next
        End If
        Send("    </select>")
        Send("  </td>")
        Send("</tr>")

        MyCommon.Close_LogixRT()
    End Sub


    Sub Send_Allow_Markup(ByVal DiscountID As Long)
        Dim MyCommon As New Copient.CommonInc
        Dim rst As DataTable
        Dim AllowMarkup As Integer = 0

        MyCommon.Open_LogixRT()

        MyCommon.QueryStr = "select AllowMarkup from CPE_Discounts with (NoLock) where DiscountID=" & DiscountID & ";"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
            AllowMarkup = MyCommon.NZ(rst.Rows(0).Item("AllowMarkup"), 0)
        End If

        Send("<tr>")
        Sendb("  <td><label for=""allowMarkup"">" & Copient.PhraseLib.Lookup("term.allowmarkup", LanguageID) & ":</label></td>")
        Send("  <td><div id =""allowMarkupDiv""style=""width:20px;""><input type=""checkbox"" name=""allowMarkup"" id =""allowMarkup"" value=""1""" & IIf(AllowMarkup = 1, " checked=""checked""", "") & " style=""margin-left:7px;""/></div></td>")
        Send("</tr>")

        MyCommon.Close_LogixRT()
    End Sub


    Sub Send_Amount_Detail(ByVal DiscountID As Long, ByVal AmountTypeID As Integer, ByVal DiscountAmount As Decimal, ByVal Level As String, _
                           ByVal TierLevel As Integer, ByVal RewardID As Long, ByRef MyCommon As Copient.CommonInc)
        Dim dt As DataTable
        Dim CurrentDiscAmt As New Decimal(0)
        Dim EntryMade As Boolean = False
        Dim bCloseConnection As Boolean = False

        If (MyCommon.LRTadoConn.State <> ConnectionState.Open) Then
            MyCommon.Open_LogixRT()
            bCloseConnection = True
        End If

        ' determine if we need to redisplay the current (unsaved) discount amount entered on the page.
        If GetCgiValue("tier" & TierLevel & "_l" & Level & "discountamt") <> "" Then
            EntryMade = True
            CurrentDiscAmt = Localizer.Round_Currency(MyCommon.Extract_Decimal(GetCgiValue("tier" & TierLevel & "_l" & Level & "discountamt"), MyCommon.GetAdminUser.Culture), RewardID, (AmountTypeID = 3))
        End If

        If TierLevel = 1 AndAlso Integer.Parse(Level) > 1 Then
            MyCommon.QueryStr = "select L" & Level & "DiscountAmt as DiscountAmount from CPE_Discounts " & _
                                "where DiscountID=" & DiscountID & ";"
        Else
            MyCommon.QueryStr = "select DiscountAmount from CPE_DiscountTiers " & _
                                "where DiscountID=" & DiscountID & " and TierLevel=" & TierLevel & ";"
        End If
        dt = MyCommon.LRT_Select()
        If dt.Rows.Count > 0 Then

            If Not EntryMade Then
                CurrentDiscAmt = Localizer.Round_Currency(MyCommon.NZ(dt.Rows(0).Item("DiscountAmount"), 0), RewardID, (AmountTypeID = 3))
            End If

            If (AmountTypeID = 4) Then
                Send("<tr>")
                Send("  <td colspan=""2"" style=""display:none;""><input type=""hidden"" id=""tier" & TierLevel & "_l" & Level & "discountamt"" name=""tier" & TierLevel & "_l" & Level & "discountamt"" value=""0"" /></td>")
                'Send("</tr>")
            ElseIf (Copient.commonShared.Contains(AmountTypeID, 1, 5, 9, 10, 11, 12)) Then
                Send("<tr>")
                Send("  <td><label for=""tier" & TierLevel & "_l" & Level & "discountamt"">" & Copient.PhraseLib.Lookup("term.amount", LanguageID) & ":</label></td>")
                Send("  <td>" & GetCurrencySymbol(RewardID) & "&nbsp;<input type=""text"" id=""tier" & TierLevel & "_l" & Level & "discountamt"" name=""tier" & TierLevel & "_l" & Level & "discountamt"" value=""" & Math.Round(CurrentDiscAmt, GetCurrencyPrecision(RewardID)).ToString(MyCommon.GetAdminUser.Culture) & """ size=""25"" maxlength=""16""" & DisabledAttr & " style=""width:183px""/> " & GetCurrencyAbbr(RewardID) & "</td>")
                'Send("</tr>")
            ElseIf (Copient.commonShared.Contains(AmountTypeID, 2, 6, 13, 14, 15, 16)) Then
                Send("<tr>")
                Send("  <td><label for=""tier" & TierLevel & "_l" & Level & "discountamt"">" & Copient.PhraseLib.Lookup("term.saleprice", LanguageID) & ":</label></td>")
                Send("  <td>" & GetCurrencySymbol(RewardID) & "&nbsp;<input type=""text"" id=""tier" & TierLevel & "_l" & Level & "discountamt"" name=""tier" & TierLevel & "_l" & Level & "discountamt"" value=""" & Math.Round(CurrentDiscAmt, GetCurrencyPrecision(RewardID)).ToString(MyCommon.GetAdminUser.Culture) & """ size=""25"" maxlength=""16""" & DisabledAttr & " style=""width:183px""/> " & GetCurrencyAbbr(RewardID) & "</td>")
                'Send("</tr>")
            ElseIf (AmountTypeID = 3) Then
                Send("<tr>")
                Send("  <td><label for=""tier" & TierLevel & "_l" & Level & "discountamt"">" & Copient.PhraseLib.Lookup("term.amount", LanguageID) & ":</label></td>")
                Send("  <td style=""padding-left:12px;""><input type=""text"" id=""tier" & TierLevel & "_l" & Level & "discountamt"" name=""tier" & TierLevel & "_l" & Level & "discountamt"" value=""" & Math.Round(CurrentDiscAmt, GetCurrencyPrecision(RewardID)).ToString(MyCommon.GetAdminUser.Culture) & """ size=""25"" maxlength=""16""" & DisabledAttr & " style=""width:183px""/> %</td>")
                'Send("</tr>")
            ElseIf (AmountTypeID = 7 Or AmountTypeID = 8) Then
                Send("<tr>")
                Send("  <td style=""display:none;""><label for=""tier" & TierLevel & "_l" & Level & "discountamt"">" & Copient.PhraseLib.Lookup("term.amount", LanguageID) & ":</label></td>")
                Send("  <td style=""display:none;"">" & GetCurrencySymbol(RewardID) & "<input type=""hidden"" id=""tier" & TierLevel & "_l" & Level & "discountamt"" name=""tier" & TierLevel & "_l" & Level & "discountamt"" value=""0"" /> " & GetCurrencyAbbr(RewardID).ToString(MyCommon.GetAdminUser.Culture) & "</td>")
                'Send("</tr>")
            End If
        Else
            If (AmountTypeID <> 4 And AmountTypeID <> 8) Then
                If DiscountAmount > 0 Then
                    Send("<tr>")
                    Send("  <td><label for=""tier" & TierLevel & "_l" & Level & "discountamt"">" & Copient.PhraseLib.Lookup("term.amount", LanguageID) & ":</label></td>")
                    If AmountTypeID = 3 Then
                        Send("  <td>&nbsp;&nbsp;&nbsp;<input type=""text"" id=""tier" & TierLevel & "_l" & Level & "discountamt"" name=""tier" & TierLevel & "_l" & Level & "discountamt"" value=""" & Math.Round(DiscountAmount, GetCurrencyPrecision(RewardID)).ToString(MyCommon.GetAdminUser.Culture) & """ size=""25"" maxlength=""16""" & DisabledAttr & " style=""width:183px"" />%</td>")
                    Else
                        Send("  <td>" & GetCurrencySymbol(RewardID) & "&nbsp;<input type=""text"" id=""tier" & TierLevel & "_l" & Level & "discountamt"" name=""tier" & TierLevel & "_l" & Level & "discountamt"" value=""" & Math.Round(DiscountAmount, GetCurrencyPrecision(RewardID)).ToString(MyCommon.GetAdminUser.Culture) & """ size=""16"" maxlength=""16""" & DisabledAttr & " style=""width:183px""/> " & GetCurrencyAbbr(RewardID) & "</td>")
                    End If
                    'Send("</tr>")
                Else
                    ' It's a new discount, so show the discountamt field left blank
                    Send("<tr>")
                    Send("  <td><label for=""tier" & TierLevel & "_l" & Level & "discountamt"">" & Copient.PhraseLib.Lookup("term.amount", LanguageID) & ":</label></td>")
                    If AmountTypeID = 3 Then
                        Send("  <td>&nbsp;&nbsp;&nbsp;<input type=""text"" id=""tier" & TierLevel & "_l" & Level & "discountamt"" name=""tier" & TierLevel & "_l" & Level & "discountamt"" value="""" size=""25"" maxlength=""16""" & DisabledAttr & " style=""width:183px""/> %</td>")
                    Else
                        Send("  <td>" & GetCurrencySymbol(RewardID) & "&nbsp;<input type=""text"" id=""tier" & TierLevel & "_l" & Level & "discountamt"" name=""tier" & TierLevel & "_l" & Level & "discountamt"" value="""" size=""25"" maxlength=""16""" & DisabledAttr & " style=""width:183px""/> " & GetCurrencyAbbr(RewardID) & "</td>")
                    End If
                    'Send("</tr>")
                End If
            End If
        End If
        If bCloseConnection Then MyCommon.Close_LogixRT()
    End Sub

    Sub Send_Amount_DetailLevels(ByVal AmountTypeID As Integer, ByVal DiscountAmount As Object, ByVal Cap As Decimal, ByVal NoCap As Boolean, _
                                 ByVal Level As String, ByVal TierLevel As Integer, ByVal RewardID As Long, ByVal MyCommon As Copient.CommonInc)
        Dim bCloseConnection As Boolean = False

        If (MyCommon.LRTadoConn.State <> ConnectionState.Open) Then
            MyCommon.Open_LogixRT()
            bCloseConnection = True
        End If

        If AmountTypeID = 3 And Not (NoCap) Then
            Send("<tr>")
            Send("  <td><label for=""l" & Level & "cap"">" & Copient.PhraseLib.Lookup("term.upto", LanguageID) & ":</label></td>")
            Send("  <td>" & GetCurrencySymbol(RewardID) & "&nbsp;<input type=""text"" id=""l" & Level & "cap"" name=""l" & Level & "cap"" value=""" & Math.Round(Cap, GetCurrencyPrecision(RewardID)).ToString(MyCommon.GetAdminUser.Culture) & """ size=""25"" maxlength=""16""" & DisabledAttr & " onkeyup=""handleUpToEntry(this.value, " & Level & ");"" style=""width:183px""/>")
            Send(GetCurrencyAbbr(RewardID))
            Send("  </td>")
            Send("</tr>")
        End If

        If bCloseConnection Then MyCommon.Close_LogixRT()
    End Sub

    Sub Send_LimitsAndText(ByRef MyCommon As Copient.CommonInc, ByVal DiscountID As Long, ByVal AmountTypeID As Integer, ByVal eDiscountType As Integer, ByVal Level As String, _
                           ByVal TierLevel As Integer, ByVal DisabledAttribute As String, ByVal RewardID As Long, ByVal bAllowDollarTransLimit As Boolean)
        Dim dt As DataTable
        Dim ItemLimit As Long = 0
        Dim WeightLimit As Double = 0
        Dim DollarLimit As Decimal = 0
        Dim bTransLimit As Boolean = False
        Dim RDesc As String = ""
        Dim BuyDesc As String = ""
        Dim DiscountTypeID As Integer = 0
        Dim trStyle As String = ""
        Dim EntryMade As Boolean = False
        Dim currentRDesc As String = ""
        Dim currentBuyDesc As String = ""
        Dim MLI As New Copient.Localization.MultiLanguageRec
        Const GROUP_LEVEL_CONDITIONAL As Integer = 4

        If GetCgiValue("tier" & TierLevel & "_itemlimit") <> "" Then
            ItemLimit = MyCommon.Extract_Decimal(GetCgiValue("tier" & TierLevel & "_itemlimit"), MyCommon.GetAdminUser.Culture)
            EntryMade = True
        End If
        If GetCgiValue("tier" & TierLevel & "_dollarlimit") <> "" Then
            DollarLimit = MyCommon.Extract_Decimal(GetCgiValue("tier" & TierLevel & "_dollarlimit"), MyCommon.GetAdminUser.Culture)
            EntryMade = True
        End If
        If GetCgiValue("tier" & TierLevel & "_weightlimit") <> "" Then WeightLimit = MyCommon.Extract_Decimal(GetCgiValue("tier" & TierLevel & "_weightlimit"), MyCommon.GetAdminUser.Culture)
        If GetCgiValue("tier" & TierLevel & "_rdesc") <> "" Then currentRDesc = GetCgiValue("tier" & TierLevel & "_rdesc")
        If GetCgiValue("tier" & TierLevel & "_buydesc") <> "" Then currentBuyDesc = GetCgiValue("tier" & TierLevel & "_buydesc")
        If GetCgiValue("discounttype") <> "" Then DiscountTypeID = MyCommon.Extract_Decimal(GetCgiValue("discounttype"), MyCommon.GetAdminUser.Culture)

        If bAllowDollarTransLimit Then
            If Request.QueryString("tier" & TierLevel & "_istranslimit") <> "" Then bTransLimit = (GetCgiValue("tier" & TierLevel & "_istranslimit") = "true")
        End If

        MyCommon.QueryStr = "select PKID, IsNull(DT.ReceiptDescription,'') as ReceiptDescription, DT.ItemLimit, DT.WeightLimit, DISC.DiscountTypeID, " & _
                            "  DT.DollarLimit, IsNull(DT.BuyDescription, '') as BuyDescription, DT.RewardLimitTypeID from CPE_DiscountTiers as DT with (NoLock) " & _
                            "inner join CPE_Discounts as DISC with (NoLock) on DISC.DiscountID = DT.DiscountID " & _
                            "where DT.DiscountID=" & DiscountID & " and DT.TierLevel=" & TierLevel & ";"
        dt = MyCommon.LRT_Select()
        If dt.Rows.Count > 0 Then
            RDesc = GetReceiptDesc(currentRDesc, MyCommon.NZ(dt.Rows(0).Item("ReceiptDescription").Replace("""", "&quot;"), ""))
            If DiscountTypeID = 0 Then DiscountTypeID = MyCommon.NZ(dt.Rows(0).Item("DiscountTypeID"), 0)
            trStyle = IIf(DiscountTypeID = GROUP_LEVEL_CONDITIONAL, " style=""display:none;""", "")

            If bAllowDollarTransLimit Then
                If MyCommon.NZ(dt.Rows(0).Item("RewardLimitTypeID"), "0") = "1" Then
                    bTransLimit = True
                End If
            End If

            If Not EntryMade Then
                DollarLimit = Math.Round(CDec(MyCommon.NZ(dt.Rows(0).Item("DollarLimit"), 0)), GetCurrencyPrecision(RewardID))
                ItemLimit = MyCommon.NZ(dt.Rows(0).Item("ItemLimit"), 0)
            End If
            If Not Copient.commonShared.Contains(AmountTypeID, 5, 6, 9, 10, 11, 12, 13, 14, 15, 16) Then
                'Send("<tr" & trStyle & ">")
                Send("  <td" & trStyle & "><label for=""tier" & TierLevel & "_itemlimit"">" & Copient.PhraseLib.Lookup("term.itemlimit", LanguageID) & ":</label></td>")
                Send("  <td" & trStyle & " style=""padding-left:12px;"">")
                Send("    <input type=""text"" id=""tier" & TierLevel & "_itemlimit"" name=""tier" & TierLevel & "_itemlimit"" value=""" & MyCommon.NZ(ItemLimit, 0) & """ size=""25"" maxlength=""4"" title=""" & Copient.PhraseLib.Lookup("term.itemlimitmsg", LanguageID) & """" & DisabledAttribute & " style=""width:183px""/>")
                Send("    <input type=""hidden"" id=""tier" & TierLevel & "_weightlimit"" name=""tier" & TierLevel & "_weightlimit"" value=""0"" /> " & GetUOMAbbr(RewardID, AmountTypeID))
                Send("  </td>")
                Send("</tr>")
            Else
                'Send("<tr" & trStyle & ">")
                Send("  <td" & trStyle & "><label for=""tier" & TierLevel & "_weightlimit"">" & Copient.PhraseLib.Lookup("term.quantity-limit", LanguageID) & ":</label></td>")
                Send("  <td" & trStyle & " style=""padding-left:12px;"">")
                Send("    <input type=""text"" id=""tier" & TierLevel & "_weightlimit"" name=""tier" & TierLevel & "_weightlimit"" value=""" & Localizer.Round_Quantity(CDec(MyCommon.NZ(dt.Rows(0).Item("WeightLimit"), 0)), RewardID, 5).ToString(MyCommon.GetAdminUser.Culture) & """ size=""25"" maxlength=""16"" title=""" & Copient.PhraseLib.Lookup("term.wgt-gal-limit-msg", LanguageID) & """" & DisabledAttribute & " style=""width:183px""/> " & GetUOMAbbr(RewardID, AmountTypeID))
                Send("    <input type=""hidden"" id=""tier" & TierLevel & "_itemlimit"" name=""tier" & TierLevel & "_itemlimit"" value=""0"" />")
                Send("  </td>")
                Send("</tr>")
            End If
            'If AmountTypeID <> 2 And AmountTypeID <> 6 And AmountTypeID <> 1 And AmountTypeID <> 5 Then
            Send("<tr>")
            Send("  <td" & trStyle & "><label for=""tier" & TierLevel & "_dollarlimit"">" & Copient.PhraseLib.Lookup("term.dollarlimit", LanguageID) & ":</label></td>")
            Send("  <td" & trStyle & ">" & GetCurrencySymbol(RewardID))
            Send("    <input type=""text"" id=""tier" & TierLevel & "_dollarlimit"" name=""tier" & TierLevel & "_dollarlimit"" value=""" & Math.Round(CDec(MyCommon.NZ(DollarLimit, 0)), GetCurrencyPrecision(RewardID)).ToString(MyCommon.GetAdminUser.Culture) & """ size=""25"" maxlength=""16"" title=""" & Copient.PhraseLib.Lookup("term.dollar-limit-msg", LanguageID) & """" & DisabledAttribute & " style=""width:183px""/> " & GetCurrencyAbbr(RewardID))

            If bAllowDollarTransLimit Then
                Sendb("&nbsp;&nbsp;")
                Sendb("	  <input type=""checkbox"" id=""tier" & TierLevel & "_istranslimit"" name=""tier" & TierLevel & "_istranslimit"" value=""true""")
                If bTransLimit Then Sendb(" checked=""checked""")
                Send(DisabledAttribute & " />")
                Sendb("<label for=""istranslimit"">")
                Sendb("Trans Limit")
                Send("</label>")
            End If

            Send("</td>")
            'Send("</tr>")
            'End If
            If (eDiscountType = 3) Or (eDiscountType = 4) Then
                'Send("<tr>")
                Send("  <td><label for=""tier" & TierLevel & "_rdesc"">" & Copient.PhraseLib.Lookup("term.receipttext", LanguageID) & ":</label></td>")
                Send("  <td style=""padding-left:12px;"">")
                'Send("    <input type=""text"" class=""medium"" id=""tier" & TierLevel & "_rdesc"" name=""tier" & TierLevel & "_rdesc"" maxlength=""18"" value=""" & MyCommon.NZ(dt.Rows(0).Item("ReceiptDescription").Replace("""", "&quot;"), "") & """" & DisabledAttribute & " /><br />")
                MLI.ItemID = MyCommon.NZ(dt.Rows(0).Item("PKID"), 0)
                MLI.MLTableName = "CPE_DiscountTiersTranslations"
                MLI.MLIdentifierName = "DiscountTiersID"
                MLI.StandardTableName = "CPE_DiscountTiers"
                MLI.StandardIdentifierName = "PKID"
                MLI.MLColumnName = "ReceiptDesc"
                MLI.StandardValue = RDesc 'MyCommon.NZ(dt.Rows(0).Item("ReceiptDescription").Replace("""", "&quot;"), "")
                MLI.InputName = "tier" & TierLevel & "_rdesc"
                MLI.InputID = "tier" & TierLevel & "_rdesc"
                MLI.InputType = "text"
                MLI.LabelPhrase = ""
                MLI.MaxLength = 18
                MLI.CSSClass = ""
                MLI.CSSStyle = "width:183px;"
                MLI.Disabled = IIf(DisabledAttribute <> "", True, False)
                Send(Localizer.SendTranslationInputs(MyCommon, MLI, Request.QueryString, 9))
                Send("  </td>")
                Send("</tr>")

                ' buy description
                If SystemCacheData.GetSystemOption_UE_ByOptionId(125) = "1" Then
                    BuyDesc = GetBuyDesc(currentBuyDesc, MyCommon.NZ(dt.Rows(0).Item("BuyDescription"), ""))
                    Send("<tr>")
                    Send("  <td><label for=""tier" & TierLevel & "_rdesc"">" & Copient.PhraseLib.Lookup("term.buydescription", LanguageID) & ":</label></td>")
                    Send("  <td>")
                    'Send("    <input type=""text"" class=""medium"" id=""tier" & TierLevel & "_buydesc"" name=""tier" & TierLevel & "_buydesc"" maxlength=""18"" value=""" & BuyDesc.Replace("""", "&quot;") & """" & DisabledAttribute & " /><br />")
                    MLI.ItemID = MyCommon.NZ(dt.Rows(0).Item("PKID"), 0)
                    MLI.MLTableName = "CPE_DiscountTiersTranslations"
                    MLI.MLIdentifierName = "DiscountTiersID"
                    MLI.StandardTableName = "CPE_DiscountTiers"
                    MLI.StandardIdentifierName = "PKID"
                    MLI.MLColumnName = "BuyDesc"
                    MLI.StandardValue = BuyDesc 'BuyDesc.Replace("""", "&quot;")
                    MLI.InputName = "tier" & TierLevel & "_buydesc"
                    MLI.InputID = "tier" & TierLevel & "_buydesc"
                    MLI.InputType = "text"
                    MLI.LabelPhrase = ""
                    MLI.MaxLength = 18
                    MLI.CSSClass = ""
                    MLI.CSSStyle = "width:183px;margin-left:10px;"
                    MLI.Disabled = IIf(DisabledAttribute <> "", True, False)
                    Send(Localizer.SendTranslationInputs(MyCommon, MLI, Request.QueryString, 9))
                    Send("  </td>")
                    Send("</tr>")
                End If
            End If

        Else
            If DiscountTypeID = 0 AndAlso DiscountID > 0 Then
                MyCommon.QueryStr = "select  DiscountTypeID from CPE_Discounts where DiscountID= " & DiscountID
                dt = MyCommon.LRT_Select()
                If dt.Rows.Count > 0 Then
                    DiscountTypeID = MyCommon.NZ(dt.Rows(0).Item("DiscountTypeID"), 0)
                End If
            End If
            trStyle = IIf(DiscountTypeID = GROUP_LEVEL_CONDITIONAL, " style=""display:none;""", "")

            ' It's a new discount, so present the default itemlimit and rdesc inputs, left blank.
            'Send("<tr" & trStyle & ">")
            If Not Copient.commonShared.Contains(AmountTypeID, 5, 6, 9, 10, 11, 12, 13, 14, 15, 16) Then
                Send("  <td" & trStyle & "><label for=""tier" & TierLevel & "_itemlimit"">" & Copient.PhraseLib.Lookup("term.itemlimit", LanguageID) & ":</label></td>")
                Send("  <td" & trStyle & " style=""padding-left:12px;"">")
                Send("    <input type=""text"" id=""tier" & TierLevel & "_itemlimit"" name=""tier" & TierLevel & "_itemlimit"" value=""" & IIf(ItemLimit = 0, "0", ItemLimit) & """ size=""25"" maxlength=""4"" title=""" & Copient.PhraseLib.Lookup("term.itemlimitmsg", LanguageID) & """" & DisabledAttribute & " style=""width:183px""/>")
                Send("    <input type=""hidden"" id=""tier" & TierLevel & "_weightlimit"" name=""tier" & TierLevel & "_weightlimit"" value=""" & IIf(WeightLimit = 0, "0", WeightLimit.ToString(MyCommon.GetAdminUser.Culture)) & """ /> " & GetUOMAbbr(RewardID, AmountTypeID))
                Send("  </td>")
            Else
                Send("  <td" & trStyle & "><label for=""tier" & TierLevel & "_weightlimit"">" & Copient.PhraseLib.Lookup("term.quantity-limit", LanguageID) & ":</label></td>")
                Send("  <td" & trStyle & " style=""padding-left:12px;"">")
                Send("    <input type=""hidden"" id=""tier" & TierLevel & "_itemlimit"" name=""tier" & TierLevel & "_itemlimit"" value=""" & IIf(ItemLimit = 0, "0", ItemLimit) & """ />")
                Send("    <input type=""text"" id=""tier" & TierLevel & "_weightlimit"" name=""tier" & TierLevel & "_weightlimit"" value=""" & IIf(WeightLimit = 0, "0", WeightLimit.ToString(MyCommon.GetAdminUser.Culture)) & """ size=""25"" maxlength=""8"" style=""width:183px""/> " & GetUOMAbbr(RewardID, AmountTypeID))
                Send("  </td>")
            End If
            Send("</tr>")
            Send("<tr>")
            Send("  <td" & trStyle & "><label for=""tier" & TierLevel & "_dollarlimit"">" & Copient.PhraseLib.Lookup("term.dollarlimit", LanguageID) & ":</label></td>")
            Send("  <td" & trStyle & ">" & GetCurrencySymbol(RewardID))
            Send("    <input type=""text"" id=""tier" & TierLevel & "_dollarlimit"" name=""tier" & TierLevel & "_dollarlimit"" value=""" & IIf(DollarLimit = 0, "0", DollarLimit.ToString(MyCommon.GetAdminUser.Culture)) & """ size=""25"" maxlength=""16"" title=""" & Copient.PhraseLib.Lookup("term.dollarlimit", LanguageID) & """" & DisabledAttribute & " style=""width:183px""/> " & GetCurrencyAbbr(RewardID))
            If bAllowDollarTransLimit Then
                Sendb("&nbsp;&nbsp;")
                Sendb("	  <input type=""checkbox"" id=""tier" & TierLevel & "_istranslimit"" name=""tier" & TierLevel & "_istranslimit"" value=""true""")
                If bTransLimit Then Sendb(" checked=""checked""")
                Send(DisabledAttribute & " />")
                Sendb("<label for=""istranslimit"">")
                Sendb("Trans Limit")
                Send("</label>")
            End If
            Send("</td>")
            'Send("</tr>")
            'Send("<tr>")
            Send("  <td><label for=""tier" & TierLevel & "_rdesc"">" & Copient.PhraseLib.Lookup("term.receipttext", LanguageID) & ":</label></td>")
            Send("  <td style=""padding-left:12px;"">")
            'Send("    <input type=""text"" class=""medium"" id=""tier" & TierLevel & "_rdesc"" name=""tier" & TierLevel & "_rdesc"" maxlength=""18"" value=""" & MyCommon.NZ(RDesc, "").Replace("""", "&quot;") & """ " & DisabledAttribute & " /><br />")
            MLI.ItemID = 0
            MLI.MLTableName = "CPE_DiscountTiersTranslations"
            MLI.MLIdentifierName = "DiscountTiersID"
            MLI.StandardTableName = "CPE_DiscountTiers"
            MLI.StandardIdentifierName = "PKID"
            MLI.MLColumnName = "ReceiptDesc"
            MLI.StandardValue = currentRDesc.Replace("""", "&quot;") 'MyCommon.NZ(RDesc, "").Replace("""", "&quot;")
            MLI.InputName = "tier" & TierLevel & "_rdesc"
            MLI.InputID = "tier" & TierLevel & "_rdesc"
            MLI.InputType = "text"
            MLI.LabelPhrase = ""
            MLI.MaxLength = 18
            MLI.CSSClass = ""
            MLI.CSSStyle = "width:183px;"
            MLI.Disabled = IIf(DisabledAttribute <> "", True, False)
            Send(Localizer.SendTranslationInputs(MyCommon, MLI, Request.QueryString, 9))
            Send("  </td>")
            Send("</tr>")

            If SystemCacheData.GetSystemOption_UE_ByOptionId(125) = "1" Then
                Send("<tr>")
                Send("  <td><label for=""tier" & TierLevel & "_buydesc"">" & Copient.PhraseLib.Lookup("term.buydescription", LanguageID) & ":</label></td>")
                Send("  <td style=""padding-left:12px;"">")
                'Send("    <input type=""text"" class=""medium"" id=""tier" & TierLevel & "_buydesc"" name=""tier" & TierLevel & "_buydesc"" maxlength=""18"" value=""" & MyCommon.NZ(BuyDesc, "").Replace("""", "&quot;") & """ " & DisabledAttribute & " /><br />")
                MLI.ItemID = 0
                MLI.MLTableName = "CPE_DiscountTiersTranslations"
                MLI.MLIdentifierName = "DiscountTiersID"
                MLI.StandardTableName = "CPE_DiscountTiers"
                MLI.StandardIdentifierName = "PKID"
                MLI.MLColumnName = "BuyDesc"
                MLI.StandardValue = currentBuyDesc.Replace("""", "&quot;")
                MLI.InputName = "tier" & TierLevel & "_buydesc"
                MLI.InputID = "tier" & TierLevel & "_buydesc"
                MLI.InputType = "text"
                MLI.LabelPhrase = ""
                MLI.MaxLength = 18
                MLI.CSSClass = ""
                MLI.CSSStyle = "width:183px;"
                MLI.Disabled = IIf(DisabledAttribute <> "", True, False)
                Send(Localizer.SendTranslationInputs(MyCommon, MLI, Request.QueryString, 9))
                Send("  </td>")
                Send("</tr>")
            End If

        End If
    End Sub

    'function to check unsaved receipt description with saved receiptd description and get current receipt description

    Function GetReceiptDesc(ByVal currentRdesc As String, ByVal savedRdesc As String) As String
        Dim CReceiptDesc As String = currentRdesc
        Dim SReceiptDesc As String = savedRdesc
        If String.Equals(CReceiptDesc, SReceiptDesc) Then
            Return SReceiptDesc
        Else
            If (CReceiptDesc = "" And SReceiptDesc <> "") Then 'during first page load currentRdesc will be empty so returning the savedRdesc if present
                Return SReceiptDesc
            Else
                Return CReceiptDesc
            End If
        End If

    End Function

    'function to check unsaved Buy description with saved Buy description and get current buy description

    Function GetBuyDesc(ByVal currentBuydesc As String, ByVal savedBuydesc As String) As String
        Dim CBuyDesc As String = currentBuydesc
        Dim SBuyDesc As String = savedBuydesc

        If String.Equals(CBuyDesc, SBuyDesc) Then
            Return SBuyDesc
        Else
            If (CBuyDesc = "" And SBuyDesc <> "") Then 'during first page load currentBuydesc will be empty so returning the savedBuyDesc if present
                Return SBuyDesc
            Else
                Return CBuyDesc
            End If

        End If

    End Function


    Sub Send_Text(ByRef MyCommon As Copient.CommonInc, ByVal DiscountID As Long, ByVal AmountTypeID As Integer, ByVal eDiscountType As Integer, ByVal Level As String, ByVal TierLevel As Integer, ByVal DisabledAttribute As String)
        Dim dt As DataTable
        Dim MLI As New Copient.Localization.MultiLanguageRec
        Dim currentRDesc As String = ""
        Dim currentBuyDesc As String = ""
        Dim RDesc As String = ""
        Dim BuyDesc As String = ""

        ' determine if we need to redisplay the current (unsaved) Receipt & Buy Description text entered on the page.
        If GetCgiValue("tier" & TierLevel & "_rdesc") <> "" Then currentRDesc = GetCgiValue("tier" & TierLevel & "_rdesc")
        If GetCgiValue("tier" & TierLevel & "_buydesc") <> "" Then currentBuyDesc = GetCgiValue("tier" & TierLevel & "_buydesc")

        MyCommon.QueryStr = "select PKID, IsNull(DT.ReceiptDescription, '') as ReceiptDescription, DT.ItemLimit, DT.WeightLimit, " & _
                            "  DT.DollarLimit, IsNull(DT.BuyDescription, '') as BuyDescription from CPE_DiscountTiers as DT with (NoLock) " & _
                            "inner join CPE_Discounts as DISC with (NoLock) on DISC.DiscountID = DT.DiscountID " & _
                            "where DT.DiscountID=" & DiscountID & " and DT.TierLevel=" & TierLevel & ";"
        dt = MyCommon.LRT_Select()
        If dt.Rows.Count > 0 Then
            RDesc = GetReceiptDesc(currentRDesc, MyCommon.NZ(dt.Rows(0).Item("ReceiptDescription"), "").Replace("""", "&quot;"))
            If (eDiscountType = 3) Or (eDiscountType = 4) Then
                Send("<tr>")
                Send("  <td><label for=""tier" & TierLevel & "_rdesc"">" & Copient.PhraseLib.Lookup("term.receipttext", LanguageID) & ":</label></td>")
                Send("  <td style=""padding-left:12px;"">")
                Send("    <!-- ksjdfhksjdhfksjdhf -->")
                MLI.ItemID = MyCommon.NZ(dt.Rows(0).Item("PKID"), 0)
                MLI.MLTableName = "CPE_DiscountTiersTranslations"
                MLI.MLIdentifierName = "DiscountTiersID"
                MLI.StandardTableName = "CPE_DiscountTiers"
                MLI.StandardIdentifierName = "PKID"
                MLI.MLColumnName = "ReceiptDesc"
                MLI.StandardValue = RDesc 'MyCommon.NZ(dt.Rows(0).Item("ReceiptDescription"), "").Replace("""", "&quot;")
                MLI.InputName = "tier" & TierLevel & "_rdesc"
                MLI.InputID = "tier" & TierLevel & "_rdesc"
                MLI.InputType = "text"
                MLI.LabelPhrase = ""
                MLI.MaxLength = 18
                MLI.CSSClass = ""
                MLI.CSSStyle = "width:183px;"
                MLI.Disabled = IIf(DisabledAttribute <> "", True, False)
                Send(Localizer.SendTranslationInputs(MyCommon, MLI, Request.QueryString, 9))
                Send("  </td>")
                Send("</tr>")
                ' buy description
                If SystemCacheData.GetSystemOption_UE_ByOptionId(125) = "1" Then
                    BuyDesc = GetBuyDesc(currentBuyDesc, MyCommon.NZ(dt.Rows(0).Item("BuyDescription"), "").Replace("""", "&quot;"))

                    Send("<tr>")
                    Send("  <td><label for=""tier" & TierLevel & "_buydesc"">" & Copient.PhraseLib.Lookup("term.buydescription", LanguageID) & ":</label></td>")
                    Send("  <td style=""padding-left:12px;"">")
                    MLI.ItemID = MyCommon.NZ(dt.Rows(0).Item("PKID"), 0)
                    MLI.MLTableName = "CPE_DiscountTiersTranslations"
                    MLI.MLIdentifierName = "DiscountTiersID"
                    MLI.StandardTableName = "CPE_DiscountTiers"
                    MLI.StandardIdentifierName = "PKID"
                    MLI.MLColumnName = "BuyDesc"
                    MLI.StandardValue = BuyDesc 'MyCommon.NZ(dt.Rows(0).Item("BuyDescription"), "").Replace("""", "&quot;")
                    MLI.InputName = "tier" & TierLevel & "_buydesc"
                    MLI.InputID = "tier" & TierLevel & "_buydesc"
                    MLI.InputType = "text"
                    MLI.LabelPhrase = ""
                    MLI.MaxLength = 18
                    MLI.CSSClass = ""
                    MLI.CSSStyle = "width:183px;"
                    MLI.Disabled = IIf(DisabledAttribute <> "", True, False)
                    Send(Localizer.SendTranslationInputs(MyCommon, MLI, Request.QueryString, 9))
                    Send("  </td>")
                    Send("</tr>")
                End If
            End If
        Else
            Send("<tr>")
            Send("  <td><label for=""tier" & TierLevel & "_rdesc"">" & Copient.PhraseLib.Lookup("term.receipttext", LanguageID) & ":</label></td>")
            Send("  <td style=""padding-left:12px;"">")
            MLI.ItemID = 0
            MLI.MLTableName = "CPE_DiscountTiersTranslations"
            MLI.MLIdentifierName = "DiscountTiersID"
            MLI.StandardTableName = "CPE_DiscountTiers"
            MLI.StandardIdentifierName = "PKID"
            MLI.MLColumnName = "ReceiptDesc"
            MLI.StandardValue = currentRDesc '""
            MLI.InputName = "tier" & TierLevel & "_rdesc"
            MLI.InputID = "tier" & TierLevel & "_rdesc"
            MLI.InputType = "text"
            MLI.LabelPhrase = ""
            MLI.MaxLength = 18
            MLI.CSSClass = ""
            MLI.CSSStyle = "width:183px;"
            MLI.Disabled = IIf(DisabledAttribute <> "", True, False)
            Send(Localizer.SendTranslationInputs(MyCommon, MLI, Request.QueryString, 9))
            Send("  </td>")
            Send("</tr>")
            ' buy description
            If SystemCacheData.GetSystemOption_UE_ByOptionId(125) = "1" Then
                Send("<tr>")
                Send("  <td><label for=""tier" & TierLevel & "_buydesc"">" & Copient.PhraseLib.Lookup("term.buydescription", LanguageID) & ":</label></td>")
                Send("  <td style=""padding-left:12px;"">")
                MLI.ItemID = 0
                MLI.MLTableName = "CPE_DiscountTiersTranslations"
                MLI.MLIdentifierName = "DiscountTiersID"
                MLI.StandardTableName = "CPE_DiscountTiers"
                MLI.StandardIdentifierName = "PKID"
                MLI.MLColumnName = "BuyDesc"
                MLI.StandardValue = currentBuyDesc '""
                MLI.InputName = "tier" & TierLevel & "_buydesc"
                MLI.InputID = "tier" & TierLevel & "_buydesc"
                MLI.InputType = "text"
                MLI.LabelPhrase = ""
                MLI.MaxLength = 18
                MLI.CSSClass = ""
                MLI.CSSStyle = "width:183px;"
                MLI.Disabled = IIf(DisabledAttribute <> "", True, False)
                Send(Localizer.SendTranslationInputs(MyCommon, MLI, Request.QueryString, 9))
                Send("  </td>")
                Send("</tr>")
            End If
        End If
    End Sub

    Sub Send_Reward_Required(ByRef MyCommon As Copient.CommonInc, ByVal DeliverableID As Long, ByVal DisableAttribute As String)

        'BZ2079: UE-feature-removal #: remove required successful delivery and assume checked.
        '        To restore previous functionality: remove send line directly below with the hidden form field requiredToDeliver and 
        '                                           uncomment the lines starting with BZ2079 Start through BZ20769 End
        Send("<tr><td colspan =""2""><input type=""hidden"" id=""requiredToDeliver"" name=""requiredToDeliver"" value=""1"" /></td></tr>")

        'Dim dt As DataTable ' BZ2079 Start
        'Dim RewardRequired As Boolean = True

        '' check a form field that always posts to determine if this is the intital page load or not. If not, then use the database value
        'If GetCgiValue("l1amounttypeid") <> "" Then
        '  RewardRequired = (MyCommon.Extract_Val(GetCgiValue("requiredToDeliver")) = 1)
        'ElseIf DeliverableID > 0 Then
        '  MyCommon.QueryStr = "select Required from CPE_Deliverables with (NoLock) where DeliverableID = " & DeliverableID
        '  dt = MyCommon.LRT_Select
        '  If dt.Rows.Count > 0 Then
        '    RewardRequired = MyCommon.NZ(dt.Rows(0).Item("Required"), True)
        '  End If
        'End If

        '' reward required
        'Send("<tr><td colspan =""2""><hr /></td></tr>")
        'Send("<tr>")
        'Send("  <td colspan=""2"">")
        'Send("    <input type=""checkbox"" id=""requiredToDeliver"" name=""requiredToDeliver"" value=""1""" & IIf(RewardRequired, " checked=""checked""", "") & DisableAttribute & " />")
        'Send("    <label for=""requiredToDeliver"">" & Copient.PhraseLib.Lookup("ue-reward.reward-required", LanguageID) & "</label>")
        'Send("  </td>")
        'Send("</tr>") ' BZ2079 End

    End Sub


    Sub Send_Proration_Box(ByRef MyCommon As Copient.CommonInc, ByVal ProrationTypeID As Integer, _
                           ByVal DiscountTypeID As Integer, ByVal DisableAttribute As String)

        'BZ2079: UE-feature-removal #: hard-coded proration type id to item level. 
        '        To restore prior functionality: remove hidden field for prorationtypeid directly below this line and uncomment lines from BZ2079 Start to BZ2079 End.
        Send("  <input type=""hidden"" name=""prorationtypeid"" id =""prorationtypeid"" value=""1"" />")

        'Dim dt As DataTable ' BZ2079 Start
        'Dim SelectedStr As String = ""
        'Dim OptionText As String = ""
        'Const ITEM_LEVEL_DISC_TYPE_ID As Integer = 1
        'Const GROUP_LEVEL_CONDITIONAL As Integer = 4
        'Const ITEM_LEVEL_CONDITIONAL As Integer = 5

        'MyCommon.QueryStr = "select ProrationTypeID, Description, PhraseID from UE_ProrationTypes with (NoLock);"
        'dt = MyCommon.LRT_Select

        'Send("<div class=""box"" id=""proration"">")
        'Send("  <h2>")
        'Send("    <span>" & Copient.PhraseLib.Lookup("term.proration", LanguageID) & "</span>")
        'Send("  </h2>")

        'If DiscountTypeID = ITEM_LEVEL_DISC_TYPE_ID Then
        '  Send(Copient.PhraseLib.Lookup("proration.discount-type-item-msg", LanguageID))
        '  Send("  <input type=""hidden"" name=""prorationtypeid"" id =""prorationtypeid"" value=""0"" />")
        'ElseIf DiscountTypeID = GROUP_LEVEL_CONDITIONAL Then
        '  Send(Copient.PhraseLib.Lookup("proration.discount-type-conprod-msg", LanguageID))
        '  Send("  <input type=""hidden"" name=""prorationtypeid"" id =""prorationtypeid"" value=""1"" />")
        'ElseIf DiscountTypeID = ITEM_LEVEL_CONDITIONAL Then
        '  Send(Copient.PhraseLib.Lookup("proration.discount-type-itemcon-msg", LanguageID))
        '  Send("  <input type=""hidden"" name=""prorationtypeid"" id =""prorationtypeid"" value=""1"" />")
        'Else
        '  Send("  <select id=""prorationtypeid"" name=""prorationtypeid"" class=""longer"" onchange=""submitForm();""" & DisableAttribute & ">")
        '  For Each row As DataRow In dt.Rows
        '    SelectedStr = IIf(ProrationTypeID = MyCommon.NZ(row.Item("ProrationTypeID"), -1), " selected=""selected""", "")
        '    OptionText = Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID, MyCommon.NZ(row.Item("Description"), "Unknown"))

        '    Send("    <option value=""" & MyCommon.NZ(row.Item("ProrationTypeID"), 0) & """" & SelectedStr & ">" & OptionText & "</option>")
        '  Next
        '  Send("  </select>")
        'End If

        'Send("  <hr class=""hidden"" />")
        'Send("</div>") ' BZ2079 End

    End Sub

    Function Get_UOM_SubType_Label(ByVal AmountTypeID As Integer, ByVal ROID As Long, ByRef Common As Copient.CommonInc) As String
        Dim Label As String = ""
        Dim dt As DataTable

        Select Case AmountTypeID
            Case 9 To 16
                Common.QueryStr = "select UST.AbbreviationPhraseTerm " & _
                                  "from UOMAmountTypes as UAT with (NoLock) " & _
                                  "inner join CPE_RewardOptionUOMs as ROU with (NoLock) " & _
                                  "  on ROU.RewardOptionID=" & ROID & " and ROU.UOMTypeID = UAT.UOMTypeID " & _
                                  "inner join UOMSubTypes as UST with (NoLock) " & _
                                  "  on UST.UOMSubTypeID = ROU.UOMSubTypeID " & _
                                  "where UAT.AmountTypeID = " & AmountTypeID
                dt = Common.LRT_Select
                If dt.Rows.Count > 0 Then
                    Label = " " & Copient.PhraseLib.Lookup(Common.NZ(dt.Rows(0).Item("AbbreviationPhraseTerm"), ""), LanguageID)
                End If
        End Select

        Return Label
    End Function
    Function GetButtonListSelectedValue() As Integer
        Dim selectedValue As Integer = 0
        Try
            If RadioButtonList1.Items.Count > 0 Then
                selectedValue = RadioButtonList1.SelectedItem.Value
            End If
        Catch ex As Exception
            selectedValue = 0
        End Try
        Return selectedValue
    End Function
    Public Sub CreateProductGroup(ByVal OfferID As Long, ByVal BuyerID As Int32, ByVal ProductGroupName As String)
        Dim productgroup As New ProductGroup
        ' productgroup.ProductGroupName = String.Format(Copient.PhraseLib.Lookup("term.offer", LanguageID) &" {0} " & Copient.PhraseLib.Lookup("term.discountedproducts", LanguageID), OfferID)
        productgroup.ProductGroupName = ProductGroupName
        productgroup.AnyProduct = False
        productgroup.ProductGroupTypeID = 2
        productgroup.BuyerID = BuyerID
        Dim ProductID As AMSResult(Of Int64) = m_ProductGroupService.CreateProductGroup(productgroup)
        If ProductID.ResultType <> AMSResultType.Success Then
            infoMessage = ProductID.MessageString
        Else
            DiscountedProductGroupID = ProductID.Result
            discountedpgid.Value = DiscountedProductGroupID
            MyCommon.Activity_Log(5, DiscountedProductGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.pgroup-create", LanguageID))
        End If
    End Sub
    Function Create_Discount(ByVal OfferID As String, ByVal TpROID As Long, ByVal Phase As Long, ByRef DeliverableID As Long, ByVal Required As Boolean) As Long
        Dim MyCommon As New Copient.CommonInc
        Dim DiscountID As Long = 0

        Try
            MyCommon.QueryStr = "dbo.pa_CPE_AddDiscount"
            MyCommon.Open_LogixRT()
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt, 4).Value = OfferID
            MyCommon.LRTsp.Parameters.Add("@TpROID", SqlDbType.Int, 4).Value = TpROID
            MyCommon.LRTsp.Parameters.Add("@Phase", SqlDbType.Int, 4).Value = Phase
            MyCommon.LRTsp.Parameters.Add("@Required", SqlDbType.Bit).Value = IIf(Required, 1, 0)
            MyCommon.LRTsp.Parameters.Add("@DiscountID", SqlDbType.Int, 4).Direction = ParameterDirection.Output
            MyCommon.LRTsp.Parameters.Add("@DeliverableID", SqlDbType.Int, 4).Direction = ParameterDirection.Output
            MyCommon.LRTsp.ExecuteNonQuery()
            DiscountID = MyCommon.LRTsp.Parameters("@DiscountID").Value
            DeliverableID = MyCommon.LRTsp.Parameters("@DeliverableID").Value
            MyCommon.Close_LRTsp()
        Catch ex As Exception
            DiscountID = -1
        Finally
            MyCommon.Close_LogixRT()
            MyCommon = Nothing
        End Try

        Return DiscountID
    End Function

    Function GetExclusionPGDataTable(ByVal excludedPGList As String) As DataTable
        Dim dt As New DataTable
        Dim dc As New DataColumn("ProductGroupId")
        Dim dr As DataRow
        Dim testInt As Integer
        dt.Columns.Add(dc)
        For Each exPGID In excludedPGList.Split(",")
            If (Integer.TryParse(exPGID.Trim, testInt)) Then    'For preventing XSS attack
                dr = dt.NewRow
                dr("ProductGroupId") = exPGID.Trim
                dt.Rows.Add(dr)
            End If
        Next

        Return dt
    End Function
    Sub UpdateExPGForCompatibility(ByVal excludedPGList As String, ByVal discountId As String)
        Dim exPG As String() = excludedPGList.Split(",")
        If (excludedPGList IsNot Nothing AndAlso excludedPGList <> "" AndAlso exPG.Length = 1) Then
            MyCommon.QueryStr = "Update CPE_Discounts Set ExcludedProductGroupId=@ExcludedPGId Where DiscountId=@DiscountId"
            MyCommon.DBParameters.Add("@DiscountId", SqlDbType.Int).Value = discountId
            MyCommon.DBParameters.Add("@ExcludedPGId", SqlDbType.Int).Value = exPG(0)
            MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
        ElseIf excludedPGList = "" AndAlso exPG.Length = 1 Then
            MyCommon.QueryStr = "Update CPE_Discounts Set ExcludedProductGroupId=@ExcludedPGId Where DiscountId=@DiscountId"
            MyCommon.DBParameters.Add("@DiscountId", SqlDbType.Int).Value = discountId
            MyCommon.DBParameters.Add("@ExcludedPGId", SqlDbType.Int).Value = 0
            MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
        End If
    End Sub
    Sub Save_Discount(ByVal OfferID As Long, ByVal DiscountID As Long, ByVal TierLevels As Long, ByVal WriteTier As Boolean, ByVal JustCreated As Boolean, _
                      ByVal RewardID As Long, ByVal bAllowDollarTransLimit As Boolean)
        Dim MyCommon As New Copient.CommonInc
        Dim Logix As New Copient.LogixInc
        Dim DiscountTypeID As Integer
        Dim eDiscountID As Long
        Dim NumRecs As String
        Dim rst As DataTable
        Dim rst2 As DataTable
        Dim row As DataRow
        Dim i As Integer = 0
        Dim AnyProduct As Boolean
        Dim IncludedProductGroupID As String
        'Dim ExcludedProductGroupID As String
        Dim AmountTypeID As Long
        Dim DecliningBalance As String
        Dim RetroactiveDiscount As String
        Dim ChargebackDeptID As Long
        Dim BestDeal As Integer
        Dim AllowNegative As Integer
        Dim ComputeDiscount As Integer
        Dim DeptLevel As Integer
        Dim ItemError As Boolean
        Dim L1Cap As Decimal
        Dim L2DiscountAmt As Decimal
        Dim L2AmountTypeID As Long
        Dim L2Cap As Decimal
        Dim L3DiscountAmt As Decimal
        Dim L3AmountTypeID As Long
        Dim DiscountBarcode As String
        Dim VoidBarcode As String
        Dim UserGroupID As String
        Dim AdminUserID As Integer
        Dim SVProgramID As Integer
        Dim FlexNeg As Boolean = False
        Dim flexOption As Integer = 0
        Dim grossPrice As Boolean = False
        Dim DiscountAmount As Decimal
        Dim ItemLimit As Int32
        Dim WeightLimit As Decimal
        Dim DollarLimit As Decimal
        Dim bTransLimit As Boolean = False
        Dim RDesc As String = ""
        Dim SaveError As String = ""
        Dim TierCount As Integer = 0
        Dim ScorecardID As Integer = 0
        Dim ScorecardDesc As String = ""
        Dim SPLevels As Integer = 0
        Dim SPHighestLevel As Integer = 0
        Dim l As Integer = 0
        Dim SPValue As Decimal = 0
        Dim DiscountTierID As Integer = 0
        Dim saveLevel As Integer = 0
        Dim AllowMarkup As Integer = 0
        Dim DiscAtOrigPrice As Integer = 0
        Dim BuyDesc As String = ""
        Dim ProrationTypeID As Integer = 0
        Dim PriceChange As Integer = 0
        Dim UOMPrecision As Integer = 3
        Dim CurrencyPrecision As Integer = 3
        Dim TierPKID As Integer = 0
        'Dim bUseMultipleProductExclusionGroups As Boolean = True
        Dim MLI As New Copient.Localization.MultiLanguageRec
        Dim amsResult As AMSResult(Of Boolean)

        MyCommon.AppName = "UEoffer-rew-discount.aspx"
        MyCommon.Open_LogixRT()
        AdminUserID = Verify_AdminUser(MyCommon, Logix)

        ItemError = False

        ' Non-tiered discount data:
        DiscountBarcode = MyCommon.Strip_Quotes(Trim(GetCgiValue("discountbarcode")))
        VoidBarcode = MyCommon.Strip_Quotes(Trim(GetCgiValue("voidbarcode")))
        DiscountType = MyCommon.Extract_Decimal(GetCgiValue("discountType"), MyCommon.GetAdminUser.Culture)
        If Not (DiscountType = 4 OrElse DiscountType = 5) Then
            IncludedProductGroupID = discountedpgid.Value
            'AMS-685 
            amsResult = m_DiscountRewardService.SaveExclusionGroups(GetExclusionPGDataTable(GetCgiValue("excludedpgid")), DiscountID)
            UpdateExPGForCompatibility(GetCgiValue("excludedpgid"), DiscountId)
            'If amsResult.ResultType = AMSResultType.Exception Then End If
            'ExcludedProductGroupID = Convert.ToInt64(MyCommon.Extract_Decimal(GetCgiValue("excludedpgid"), MyCommon.GetAdminUser.Culture))
        End If
        ChargebackDeptID = MyCommon.Extract_Decimal(GetCgiValue("chargeback"), MyCommon.GetAdminUser.Culture)
        UserGroupID = Convert.ToInt64(MyCommon.Extract_Decimal(GetCgiValue("usergroupid"), MyCommon.GetAdminUser.Culture))
        BestDeal = MyCommon.Extract_Decimal(GetCgiValue("bestdeal"), MyCommon.GetAdminUser.Culture)
        DiscAtOrigPrice = MyCommon.Extract_Decimal(GetCgiValue("discpricelevel"), MyCommon.GetAdminUser.Culture)

        If (DiscountType = 4 OrElse DiscountType = 5) Then
            PriceFilterID = 100
        Else
            PriceFilterID = MyCommon.Extract_Val(GetCgiValue("priceFilter"), MyCommon.GetAdminUser.Culture)
        End If

        ComputeDiscount = MyCommon.Extract_Decimal(GetCgiValue("computediscount"), MyCommon.GetAdminUser.Culture)
        DiscountTypeID = MyCommon.Extract_Decimal(GetCgiValue("discountType"), MyCommon.GetAdminUser.Culture)
        'SVProgramID = MyCommon.Extract_Val(GetCgiValue("svprogramid"))
        If GetCgiValue("decliningbalance") = "true" Then
            DecliningBalance = "1"
        Else
            DecliningBalance = "0"
        End If
        RetroactiveDiscount = "1"
        If GetCgiValue("flexnegative") = "1" OrElse GetCgiValue("flexnegative") = "2" Then
            FlexNeg = True
        End If
        flexOption = GetCgiValue("flexnegative")
        grossPrice = ((GetCgiValue("grossprice")) = "on" )
        AmountTypeID = MyCommon.Extract_Decimal(GetCgiValue("l1amounttypeid"), MyCommon.GetAdminUser.Culture)

        If CMS.Utilities.Extract_Val(GetCgiValue("discountsv")) = 1 Then
            SVProgramID = CMS.Utilities.Extract_Val(GetCgiValue("ddldiscountsv"))
        Else
            SVProgramID = 0
        End If

        If Not Copient.commonShared.Contains(AmountTypeID, 5, 9, 10, 11, 12) Then ComputeDiscount = 1
        'If ComputeDiscount = 0 Then
        '  AllowNegative = 1
        'End If

        Select Case AmountTypeID
            Case 2, 3, 6, 8, 13, 14, 15, 16 '
                PriceChange = IIf(GetCgiValue("pricechange") = "true", 1, 0)
            Case Else
                PriceChange = 0
        End Select

        L1Cap = MyCommon.Extract_Decimal(GetCgiValue("l1cap"), MyCommon.GetAdminUser.Culture)
        L2DiscountAmt = MyCommon.Extract_Decimal(GetCgiValue("tier1_l2discountamt"), MyCommon.GetAdminUser.Culture)
        L2AmountTypeID = MyCommon.Extract_Decimal(GetCgiValue("l2amounttypeid"), MyCommon.GetAdminUser.Culture)
        L2Cap = MyCommon.Extract_Decimal(GetCgiValue("l2cap"), MyCommon.GetAdminUser.Culture)
        L3DiscountAmt = MyCommon.Extract_Decimal(GetCgiValue("tier1_l3discountamt"), MyCommon.GetAdminUser.Culture)
        L3AmountTypeID = MyCommon.Extract_Decimal(GetCgiValue("l3amounttypeid"), MyCommon.GetAdminUser.Culture)
        ScorecardID = MyCommon.Extract_Decimal(GetCgiValue("ScorecardID"), MyCommon.GetAdminUser.Culture)
        ScorecardDesc = IIf(GetCgiValue("ScorecardDesc") <> "", GetCgiValue("ScorecardDesc"), "")
        AllowMarkup = IIf(GetCgiValue("AllowMarkup") = "1", 1, 0)
        ProrationTypeID = MyCommon.Extract_Decimal(GetCgiValue("prorationtypeid"), MyCommon.GetAdminUser.Culture)

        ' Tiered discount data; for now, let's just set these to the first-tier values:
        RDesc = MyCommon.Strip_Quotes(GetCgiValue("tier1_rdesc"))
        BuyDesc = MyCommon.Strip_Quotes(Trim(GetCgiValue("tier1_buydesc")))
        ItemLimit = MyCommon.Extract_Decimal(GetCgiValue("tier1_itemlimit"), MyCommon.GetAdminUser.Culture)
        WeightLimit = MyCommon.Extract_Decimal(GetCgiValue("tier1_weightlimit"), MyCommon.GetAdminUser.Culture)
        DollarLimit = MyCommon.Extract_Decimal(GetCgiValue("tier1_dollarlimit"), MyCommon.GetAdminUser.Culture)
        DiscountAmount = MyCommon.Extract_Decimal(GetCgiValue("tier1_l1discountamt"), MyCommon.GetAdminUser.Culture)

        If AmountTypeID = 4 Then DiscountAmount = 0
        If L2AmountTypeID = 4 Then L2DiscountAmt = 0
        If L3AmountTypeID = 4 Then L3DiscountAmt = 0

        ' If the discount is percent off, set a default for the L2AmountTypeID
        If ((AmountTypeID = 3) And (DiscountAmount > 0)) Then
            If L2AmountTypeID = 0 Then L2AmountTypeID = 1
        End If
        If ((L2AmountTypeID = 3) And (L2DiscountAmt > 0)) Then
            If L3AmountTypeID = 0 Then L3AmountTypeID = 1
        End If

        ' See if the discounted product group is the AnyProduct group
        AnyProduct = False
        If (String.IsNullOrEmpty(IncludedProductGroupID)) Then
            IncludedProductGroupID = 0
        End If
        MyCommon.QueryStr = "select isnull(AnyProduct, 0) as AnyProduct from ProductGroups with (NoLock) where ProductGroupID=" & IncludedProductGroupID & ";"
        rst = MyCommon.LRT_Select
        If Not (rst.Rows.Count = 0) Then
            AnyProduct = MyCommon.NZ(rst.Rows(0).Item("AnyProduct"), False)
            DeptLevel = IIf(DiscountTypeID = 2, 1, 0)
        End If
        If AnyProduct Then
            ' Stored value (AmountTypeID 7) should not set the item limit to 1 because the local server expects a 0.
            If (AmountTypeID = 7) Then
                ItemLimit = 0
            Else
                ItemLimit = 1
                DollarLimit = 0
            End If
            RetroactiveDiscount = "0"
            DecliningBalance = "0"
            'AMS-685 this else if is not required as dept discount also has exclusion groups
            'ElseIf DeptLevel > 0 Then
            'If Not bUseMultipleProductExclusionGroups Then
            '    ExcludedProductGroupID = 0
            'End If
            'If (ChargebackDeptID = 0 Or ChargebackDeptID = 14) Then ChargebackDeptID = 10
        Else
            'If Not bUseMultipleProductExclusionGroups Then
            '    ExcludedProductGroupID = 0
            'End If
            If (ChargebackDeptID = 10) Then ChargebackDeptID = "0"
        End If

        If (DiscountID > 0) Then
            eDiscountID = DiscountID
        Else
            eDiscountID = MyCommon.Extract_Decimal(GetCgiValue("DiscountID"), MyCommon.GetAdminUser.Culture)
        End If

        If IncludedProductGroupID = 0 Then
            IncludedProductGroupID = "Null"
            'ExcludedProductGroupID = "Null"
        End If

        ' Force the discount type to basket level if the discounted product grouo is AnyProduct
        If (AnyProduct AndAlso DiscountTypeID <> 3) Then
            IncludedProductGroupID = "Null"
        ElseIf (AnyProduct) Then
            DiscountTypeID = 3
        End If

        If (AmountTypeID = 8) Then
            SavePricePointLevels(OfferID, DiscountID, WriteTier)
        End If

        NumRecs = 0

        ScorecardDesc = Replace(ScorecardDesc, "'", "''")

        UOMPrecision = GetUOMPrecision(RewardID, AmountTypeID)
        CurrencyPrecision = GetCurrencyPrecision(RewardID)

        'AMS-685 removed "ExcludedProductGroupID=" & ExcludedProductGroupID & ", " & _
        ' Update the CPE_Discounts table
        MyCommon.QueryStr = "update CPE_Discounts with (RowLock) set " &
                                                "DiscountTypeID='" & DiscountTypeID & "', " &
                                                "ReceiptDescription='" & RDesc & "', " &
                                                "DiscountBarcode='" & DiscountBarcode & "', " &
                                                "VoidBarcode='" & VoidBarcode & "', " &
                                                "DiscountedProductGroupID=" & IncludedProductGroupID & ", " &
                                                "BestDeal=" & BestDeal & ", " &
                                                "AllowNegative=" & AllowNegative & ", " &
                                                "ComputeDiscount=" & ComputeDiscount & ", " &
                                                "DiscountAmount=" & Localizer.Round_Currency(DiscountAmount, TpROID, (AmountTypeID = 3)).ToString() & ", " &
                                                "AmountTypeID=" & AmountTypeID & ", " &
                                                "L1Cap=" & Localizer.Round_Currency(L1Cap, TpROID).ToString(CultureInfo.InvariantCulture) & ", " &
                                                "L2DiscountAmt=" & Localizer.Round_Currency(L2DiscountAmt, TpROID).ToString(CultureInfo.InvariantCulture) & ", " &
                                                "L2AmountTypeID=" & L2AmountTypeID & ", L2Cap=" & Math.Round(L2Cap, TpROID).ToString(CultureInfo.InvariantCulture) & ", " &
                                                "L3DiscountAmt=" & Localizer.Round_Currency(L3DiscountAmt, TpROID).ToString(CultureInfo.InvariantCulture) & ", L3AmountTypeID=" & L3AmountTypeID & ", " &
                                                "ItemLimit=" & ItemLimit & ", " &
                                                "WeightLimit=" & Localizer.Round_Quantity(WeightLimit, TpROID, 5).ToString(CultureInfo.InvariantCulture) & ", " &
                                                "DollarLimit=" & Localizer.Round_Currency(DollarLimit, TpROID).ToString(CultureInfo.InvariantCulture) & ", " &
                                                "ChargebackDeptID=" & ChargebackDeptID & ", " &
                                                "DecliningBalance=" & DecliningBalance & ", " &
                                                "RetroactiveDiscount=" & RetroactiveDiscount & ", " &
                                                "UserGroupID=" & UserGroupID & ", " &
                                                "LastUpdate=getdate(), " &
                                                "FlexNegative='" & FlexNeg & "', " &
                                                "AllowMarkup=" & AllowMarkup & ", " &
                                                "DiscountAtOrigPrice=" & IIf(DiscAtOrigPrice = 1, 1, 0) & ", " &
                                                "ProrationTypeID=" & ProrationTypeID & ", " &
                                                "PriceChange=" & PriceChange & ", " &
                                                "FlexOptions=" & flexOption & ", " &
                                                "PriceFilter=" & PriceFilterID & ", " &
                                                "GrossPrice='" & grossPrice & "', "
        MyCommon.QueryStr += "SVProgramID=" & IIf(SVProgramID > 0, SVProgramID.ToString, "Null")
        MyCommon.QueryStr += ", ScorecardID=" & IIf(ScorecardID > 0, ScorecardID, "NULL") & ", ScorecardDesc=" & IIf(ScorecardDesc = "", "NULL", "'" & ScorecardDesc & "'") & ""
        MyCommon.QueryStr += " where DiscountID=" & eDiscountID & ";"
        MyCommon.LRT_Execute()

        'Save multilanguage values:
        'ScorecardDesc
        If (ScorecardID > 0) Then
            SetScoreCardMLI(eDiscountID, MLI) 'Set the MLI properties for scorecard 
            Localizer.SaveTranslationInputs(MyCommon, MLI, Request.Form, 9)
        Else
            MyCommon.QueryStr = "select * from CPE_DiscountTranslations where DiscountID = @DiscountID;"
            MyCommon.DBParameters.Add("@DiscountID", SqlDbType.Int, 4).Value = eDiscountID
            Dim MLIScorecard As DataTable = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            If (MLIScorecard.Rows.Count > 0) Then
                MyCommon.QueryStr = "delete from CPE_DiscountTranslations where DiscountID = @DiscountID;"
                MyCommon.DBParameters.Add("@DiscountID", SqlDbType.Int, 4).Value = eDiscountID
                MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
            End If
        End If

        'Create the CPE_DiscountTiers 
        If WriteTier Then
            'Delete the tiers
            MyCommon.QueryStr = "delete from CPE_DiscountTiers with (RowLock) where DiscountID=" & eDiscountID & ";"
            MyCommon.LRT_Execute()
            For i = 1 To TierLevels
                MyCommon.QueryStr = "dbo.pa_CPE_AddDiscountTiers"
                MyCommon.Open_LogixRT()
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@DiscountID", SqlDbType.Int, 4).Value = eDiscountID
                MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int, 4).Value = i
                MyCommon.LRTsp.Parameters.Add("@ReceiptDescription", SqlDbType.NVarChar, 100).Value = MyCommon.Strip_Quotes(GetCgiValue("tier" & i & "_rdesc"))
                MyCommon.LRTsp.Parameters.Add("@DiscountAmount", SqlDbType.Decimal, 15).Value = Localizer.Round_Currency(MyCommon.Extract_Decimal(GetCgiValue("tier" & i & "_l1discountamt"), MyCommon.GetAdminUser.Culture), TpROID, (AmountTypeID = 3))
                MyCommon.LRTsp.Parameters.Add("@ItemLimit", SqlDbType.Int, 4).Value = MyCommon.Extract_Decimal(GetCgiValue("tier" & i & "_itemlimit"), MyCommon.GetAdminUser.Culture)
                MyCommon.LRTsp.Parameters.Add("@WeightLimit", SqlDbType.Decimal, 15).Value = Localizer.Round_Quantity(MyCommon.Extract_Decimal(GetCgiValue("tier" & i & "_weightlimit"), MyCommon.GetAdminUser.Culture), RewardID, 5)
                MyCommon.LRTsp.Parameters.Add("@DollarLimit", SqlDbType.Decimal, 15).Value = Localizer.Round_Currency(MyCommon.Extract_Decimal(GetCgiValue("tier" & i & "_dollarlimit"), MyCommon.GetAdminUser.Culture), RewardID, False)
                MyCommon.LRTsp.Parameters.Add("@SPRepeatLevel", SqlDbType.TinyInt, 4).Value = MyCommon.Extract_Decimal(GetCgiValue("tier" & i & "_sprepeatlevel"), MyCommon.GetAdminUser.Culture)
                MyCommon.LRTsp.Parameters.Add("@BuyDescription", SqlDbType.NVarChar, 255).Value = MyCommon.Strip_Quotes(Trim(GetCgiValue("tier" & i & "_buydesc")))
                MyCommon.LRTsp.Parameters.Add("@PKID", SqlDbType.Int, 4).Direction = ParameterDirection.Output

                If bAllowDollarTransLimit Then
                    bTransLimit = (GetCgiValue("tier" & i & "_istranslimit") = "true")
                End If
                MyCommon.LRTsp.Parameters.Add("@RewardLimitTypeID", SqlDbType.Int).Value = IIf(bTransLimit, 1, 0)

                MyCommon.LRTsp.ExecuteNonQuery()
                TierPKID = MyCommon.LRTsp.Parameters("@PKID").Value
                MyCommon.Close_LRTsp()
                'Save multilanguage values:
                'ReceiptDesc
                MLI.ItemID = TierPKID
                MLI.MLTableName = "CPE_DiscountTiersTranslations"
                MLI.MLColumnName = "ReceiptDesc"
                MLI.MLIdentifierName = "DiscountTiersID"
                MLI.StandardTableName = "CPE_DiscountTiers"
                MLI.StandardColumnName = "ReceiptDescription"
                MLI.StandardIdentifierName = "PKID"
                MLI.StandardValue = MyCommon.Strip_Quotes(GetCgiValue("tier" & i & "_rdesc"))
                MLI.InputName = "tier" & i & "_rdesc"
                Localizer.SaveTranslationInputs(MyCommon, MLI, Request.Form, 9)
                'BuyDesc
                MLI.MLColumnName = "BuyDesc"
                MLI.StandardColumnName = "BuyDescription"
                MLI.StandardValue = MyCommon.Strip_Quotes(Trim(GetCgiValue("tier" & i & "_buydesc")))
                MLI.InputName = "tier" & i & "_buydesc"
                Localizer.SaveTranslationInputs(MyCommon, MLI, Request.Form, 9)

            Next
        End If

        'If special pricing, create the CPE_SpecialPricing entries
        If (AmountTypeID = 8) Then
            'Delete any existing special pricing levels
            MyCommon.QueryStr = "delete from CPE_SpecialPricing with (RowLock) where DiscountID=" & eDiscountID & ";"
            MyCommon.LRT_Execute()
            For i = 1 To TierLevels
                'Find the PKID (DiscountTierID) of each tier
                SPLevels = MyCommon.Extract_Decimal(GetCgiValue("tier" & i & "_levels"), MyCommon.GetAdminUser.Culture)
                SPHighestLevel = MyCommon.Extract_Decimal(GetCgiValue("tier" & i & "_highestlevel"), MyCommon.GetAdminUser.Culture)
                MyCommon.QueryStr = "select PKID as DiscountTierID from CPE_DiscountTiers with (NoLock) " & _
                                    "where DiscountID=" & eDiscountID & " and TierLevel=" & i & ";"
                rst = MyCommon.LRT_Select
                If rst.Rows.Count > 0 Then
                    DiscountTierID = MyCommon.NZ(rst.Rows(0).Item("DiscountTierID"), 0)
                End If
                saveLevel = 0
                For l = 1 To SPHighestLevel
                    'Save each level
                    If GetCgiValue("tier" & i & "_level" & l) <> Nothing Then
                        saveLevel = saveLevel + 1
                        SPValue = MyCommon.Extract_Decimal(GetCgiValue("tier" & i & "_level" & l), MyCommon.GetAdminUser.Culture)
                        MyCommon.QueryStr = "insert into CPE_SpecialPricing (DiscountID, DiscountTierID, Value, LevelID) values " & _
                                            "(" & eDiscountID & ", " & DiscountTierID & ", " & Localizer.Round_Currency(SPValue, TpROID, False).ToString(CultureInfo.InvariantCulture) & ", " & saveLevel & ");"
                        MyCommon.LRT_Execute()
                    End If
                Next
            Next
        End If
        ' Update the templates permission if necessary
        If (IsTemplate) Then
            ' time to update the status bits for the templates
            Dim form_Disallow_Edit As Integer = 0
            If (GetCgiValue("Disallow_Edit") = "on") Then
                form_Disallow_Edit = 1
            End If
            MyCommon.QueryStr = "update CPE_Deliverables with (RowLock) set DisallowEdit=" & form_Disallow_Edit & " " & _
                                "where DeliverableID=" & DeliverableID & ";"
            MyCommon.LRT_Execute()
            'clear the template field exception permissions
            MyCommon.QueryStr = "delete from TemplateFieldPermissions with (RowLock) where OfferID=" & OfferID & " " & _
                                "and FieldID in (select FieldID from UIFields where PageName='" & MyCommon.AppName & "');"
            MyCommon.LRT_Execute()
            If (GetCgiValue("chkTempField") <> "") Then
                Dim tmpFldLen As Integer = GetCgiValue("chkTempField").Length
                If (tmpFldLen > 0) Then
                    ReDim LockFieldsList(tmpFldLen)
                    LockFieldsList = Request.Form.GetValues("chkTempField")
                    For i = 0 To LockFieldsList.Length - 1
                        MyCommon.QueryStr = "insert into TemplateFieldPermissions with (RowLock) (OfferID, FieldID,DeliverableID, Editable) " & _
                                            "values (" & OfferID & ", " & LockFieldsList(i) & "," & DeliverableID & "," & form_Disallow_Edit & ");"
                        MyCommon.LRT_Execute()
                    Next
                End If
            End If
        End If
        ' Update the CPE_Incentives table
        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
        MyCommon.LRT_Execute()
        ResetOfferApprovalStatus(OfferID)

breakout:
        MyCommon.Close_LogixRT()
    End Sub

    Sub SavePricePointLevels(ByVal OfferID As Long, ByVal DiscountID As Long, ByVal WriteTier As Boolean)
        Dim MyCommon As New Copient.CommonInc
        Dim Logix As New Copient.LogixInc
        Dim AdminUserID As Integer
        Dim LevelValues As String()
        Dim LastSPLevel As Integer = 0
        Dim i As Integer = 0

        If (MyCommon.LRTadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixRT()
        AdminUserID = Verify_AdminUser(MyCommon, Logix)

        ' Update price point levels
        If (GetCgiValue("level") <> "") Then
            LevelValues = Request.QueryString.GetValues("level")

            If WriteTier Then
                ' Next, save values to CPE_DiscountTiers.  We'll first delete out the existing special pricing values...
                MyCommon.QueryStr = "delete from CPE_DiscountTiers where DiscountID=" & DiscountID & ";"
                MyCommon.LRT_Execute()
                ' ...then insert the new ones.
                i = 0
                For i = 0 To (LevelValues.Length - 1)
                    MyCommon.QueryStr = "insert into CPE_DiscountTiers " & _
                                        "(DiscountID, TierLevel, ReceiptDescription, DiscountAmount, ItemLimit, WeightLimit, DollarLimit, LastUpdate) " & _
                                        "values (" & DiscountID & ", 1, '" & MyCommon.Strip_Quotes(GetCgiValue("tier1_rdesc")) & "', " & MyCommon.Extract_Decimal(LevelValues(i), MyCommon.GetAdminUser.Culture).ToString(CultureInfo.InvariantCulture) & ", " & _
                                        "0, 0, 0, " & (i + 1) & ", getdate());"
                    MyCommon.LRT_Execute()
                Next

            End If
        End If
    End Sub


    ' this function ensures that the conditions for allowing discount at original price are met
    Function IsValidDiscPriceLevel(ByRef MyCommon As Copient.CommonInc, ByRef infoMessage As String) As Boolean
        Dim Valid As Boolean = True
        Dim DiscPriceLevel As Integer = 0
        Dim AmountType As Integer = 0
        Dim DiscountType As Integer = 0

        DiscPriceLevel = MyCommon.Extract_Decimal(GetCgiValue("discpricelevel"), MyCommon.GetAdminUser.Culture)
        DiscountType = MyCommon.Extract_Decimal(GetCgiValue("discounttype"), MyCommon.GetAdminUser.Culture)
        AmountType = MyCommon.Extract_Decimal(GetCgiValue("l1amounttypeid"), MyCommon.GetAdminUser.Culture)

        If DiscPriceLevel = DISC_ORIGINAL_PRICE Then
            ' a discount-type of item-level and distribution type of Fixed Amount Off or Percent Off is required
            If DiscountType <> 1 Then
                Valid = False
                infoMessage = Copient.PhraseLib.Lookup("UE-discount.pricelevel-disc-type", LanguageID)
            ElseIf AmountType <> 1 AndAlso AmountType <> 3 Then
                Valid = False
                infoMessage = Copient.PhraseLib.Lookup("UE-discount.pricelevel-distribution-type", LanguageID)
            End If

        End If

        Return Valid
    End Function

    Function GetCurrencyPrecision(ByVal RewardID As Long) As Integer
        Return Localizer.GetCached_Currency_Precision(RewardID)
    End Function

    Function GetCurrencyAbbr(ByVal RewardID As Long) As String
        Return Localizer.GetCached_Currency_Abbreviation(RewardID)
    End Function

    Function GetCurrencySymbol(ByVal RewardID As Long) As String
        Return Localizer.GetCached_Currency_Symbol(RewardID)
    End Function

    Function GetUOMAbbr(ByVal RewardID As Long, ByVal AmountTypeID As Integer) As String
        Return Localizer.GetCached_UOM_Abbreviation(RewardID, AmountTypeID, Copient.Localization.UOMUsageEnum.AmountType)
    End Function

    Function GetUOMPrecision(ByVal RewardID As Long, ByVal AmountTypeID As Integer) As Integer
        Return Localizer.GetCached_UOM_Precision(RewardID, AmountTypeID, Copient.Localization.UOMUsageEnum.AmountType)
    End Function

    Protected Sub RadioButtonList1_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButtonList1.DataBound
        Dim PhraseID As Integer
        For i As Integer = 0 To sender.Items.Count - 1
            If (Not String.IsNullOrWhiteSpace(sender.Items(i).Text) AndAlso Int32.TryParse(sender.Items(i).Text, PhraseID)) Then
                sender.Items(i).Text = Copient.PhraseLib.Lookup(PhraseID, LanguageID)
            End If
        Next
    End Sub
    'Function to be called asynchronously by javascript for fetching paged attribute values
    <WebMethod>
    <WebInvoke(Method:="POST")>
    Public Shared Function GetAttributes(term As String, attributetype As Int16, pageindex As Int32, excludeattr As String, keyValue As String) As List(Of Attributes)
        CurrentRequest.Resolver.AppName = "UEoffer-rew-discount.aspx"
        Dim attributeService As IAttributeService = CurrentRequest.Resolver.Resolve(Of IAttributeService)()
        Dim nodeIdsDT As DataTable = HttpContext.Current.Session("AllChildNodes")
        If HttpContext.Current.Session("AllChildNodes") IsNot Nothing Then
            nodeIdsDT = CType(HttpContext.Current.Session("AllChildNodes"), DataTable)
        End If
        Dim amsResultAttributes As AMSResult(Of List(Of Attributes)) = attributeService.GetAttributesInChunks(term, attributetype, pageindex, nodeIdsDT, excludeattr, keyValue)
        If amsResultAttributes.ResultType <> AMSResultType.Success AndAlso amsResultAttributes.MessageString <> String.Empty Then
            Dim activityFields As ActivityLogFields = New ActivityLogFields
            Dim myCommon As New Copient.CommonInc
            myCommon.Open_LogixRT()
            activityFields.LinkID = attributetype
            activityFields.Description = amsResultAttributes.MessageString
            myCommon.Activity_Log3(activityFields)
            myCommon.Close_LogixRT()
        End If
        Return amsResultAttributes.Result
    End Function

</script>
<%
    If (IsTemplate) Then
        Send("<script type=""text/javascript"" language=""javascript"">")
        Send("  xmlhttpPost1(""UEtemplateFeeds.aspx?OfferID=" & OfferID & "&CloseAfterSave=" & CloseAfterSave & "&DeliverableID=" & DeliverableID & "&PageName=" & Server.UrlEncode(MyCommon.AppName) & "&PageEditable=" & (Not Disallow_Edit) & """)")
        If (Not OverrideFields Is Nothing AndAlso OverrideFields.Count > 0) Then
            Send("  var elem = null;")

            For Each de As DictionaryEntry In OverrideFields
                If OverrideDiv(de.Key.ToString).ToString.ToUpper = "TRUE" Then
                    Send("  elem = document.getElementById(""" & de.Key.ToString & "Div"");")
                Else
                    Send("  elem = document.mainform." & de.Key.ToString & ";")
                End If
		
                OvrdFldClass = IIf(de.Value.ToString.ToUpper = "TRUE", "enabledTemplateField", "disabledTemplateField")
                Send("  if (elem != null) { elem.setAttribute('className', '" & OvrdFldClass & "'); }")
                Send("  if (elem != null) { elem.className = '" & OvrdFldClass & "'; }")
            Next de
        End If
        Send("</script>")
    ElseIf (FromTemplate) Then
        If (Not OverrideFields Is Nothing AndAlso OverrideFields.Count > 0) Then
            Send("<script type=""text/javascript"" language=""javascript"">")
            Send("var elem = null;")
            For Each de As DictionaryEntry In OverrideFields
                Send("elem = document.mainform." & de.Key.ToString & ";")
                OvrdFldDisabled = IIf(de.Value.ToString.ToUpper = "TRUE", "false", "true")
                Send("if (elem != null) { elem.disabled = " & OvrdFldDisabled & "; }")
            Next de
            Send("</script>")
        End If
    End If
%>
<script type="text/javascript">
    
    $(document).ready(function () {
        assignGrossPriceDefault();
    });
    
    function ChangeDiscountSV(obj) {
        if (obj.checked) {
            $('#ddldiscountsv')[0].style.display = '';

        }
        else {
            $('#ddldiscountsv')[0].style.display = 'none';
            $('#ddldiscountsv')[0].options.selectedIndex = 0;

        }
    }

    <% If (CloseAfterSave) Then %>
    opener.location = "/logix/UE/UEoffer-rew.aspx?OfferID=<%Sendb(OfferID)%>";
    window.close();
      <%Else
        PABStage = 2%>
    ProductGroupDivSelection();
    FlipUI();
      <% End If%>
    toggleflexneg()
      <% If Not ChargebackSet And LoadDefaultChargeback Then%>
    handleChargebackDept(<%Sendb(DefaultChrgBack)%>);
    <% End If%>
      
    function setSameTierValue(tierLevels){
      var box = document.getElementById("useSameTierValue");
      var text;
      if(box.checked){
        for (i=1; i < (tierLevels + 1); i++){
          text = "tier" + i.toString() + "_l1discountamt";
          //alert(document.getElementById("tier1_l1discountamt").value.toString());
          document.getElementById(text).value = document.getElementById("tier1_l1discountamt").value;
          document.getElementById(text).setAttribute('disabled', 'disabled');
          if(document.getElementById("tier1_itemlimit")!=null)
          {
            text = "tier" + i.toString() + "_itemlimit";
            document.getElementById(text).value = document.getElementById("tier1_itemlimit").value;
            document.getElementById(text).setAttribute('disabled', 'disabled');
          }
          if(document.getElementById("tier1_weightlimit")!=null)
          {
            text = "tier" + i.toString() + "_weightlimit";
            document.getElementById(text).value = document.getElementById("tier1_weightlimit").value;
            document.getElementById(text).setAttribute('disabled', 'disabled');
          }
          if(document.getElementById("tier1_dollarlimit")!=null)
          {
            text = "tier" + i.toString() + "_dollarlimit";
            document.getElementById(text).value = document.getElementById("tier1_dollarlimit").value;
            document.getElementById(text).setAttribute('disabled', 'disabled');
          }
          if(document.getElementById("tier1_istranslimit")!=null)
          {
            text = "tier" + i.toString() + "_istranslimit";
            document.getElementById(text).checked = document.getElementById("tier1_istranslimit").checked;
            document.getElementById(text).setAttribute('disabled', 'disabled');
          }
          text = "tier" + i.toString() + "_rdesc";
          document.getElementById(text).value = document.getElementById("tier1_rdesc").value;
          document.getElementById(text).setAttribute('disabled', 'disabled');
          if(document.getElementById("tier1_buydesc")!=null)
          {
            text = "tier" + i.toString() + "_buydesc";
            document.getElementById(text).value = document.getElementById("tier1_buydesc").value;
            document.getElementById(text).setAttribute('disabled', 'disabled');
          }
        } 
      }
      else{
        for (i=1; i < (tierLevels + 1); i++){
          text = "tier" + i.toString() + "_l1discountamt";
          document.getElementById(text).disabled = false;
          if(document.getElementById("tier1_itemlimit")!=null)
          {
            text = "tier" + i.toString() + "_itemlimit";
            document.getElementById(text).disabled = false;
          }
          if(document.getElementById("tier1_weightlimit")!=null)
          {
            text = "tier" + i.toString() + "_weightlimit";
            document.getElementById(text).disabled = false;
          }
          if(document.getElementById("tier1_dollarlimit")!=null)
          {
            text = "tier" + i.toString() + "_dollarlimit";
            document.getElementById(text).disabled = false;
          }
          if(document.getElementById("tier1_istranslimit")!=null)
          {
            text = "tier" + i.toString() + "_istranslimit";
            document.getElementById(text).disabled = false;
          }
          text = "tier" + i.toString() + "_rdesc";
          document.getElementById(text).disabled = false;
          if(document.getElementById("tier1_buydesc")!=null)
          {
            text = "tier" + i.toString() + "_buydesc";
            document.getElementById(text).disabled = false;
          }
        } 
      }
    }

      
    function enableTiers(tierLevels){
     var box = document.getElementById("useSameTierValue");
      var text;
      if(box.checked && tierLevels > 1){
        for (i=1; i < (tierLevels + 1); i++){
          text = "tier" + i.toString() + "_l1discountamt";
          document.getElementById(text).disabled = false;
          if(document.getElementById("tier1_itemlimit")!=null)
          {
            text = "tier" + i.toString() + "_itemlimit";
            document.getElementById(text).disabled = false;
          }
          if(document.getElementById("tier1_weightlimit")!=null)
          {
            text = "tier" + i.toString() + "_weightlimit";
            document.getElementById(text).disabled = false;
          }
          if(document.getElementById("tier1_dollarlimit")!=null)
          {
            text = "tier" + i.toString() + "_dollarlimit";
            document.getElementById(text).disabled = false;
          }
          if(document.getElementById("tier1_istranslimit")!=null)
          {
            text = "tier" + i.toString() + "_istranslimit";
            document.getElementById(text).disabled = false;
          }
          text = "tier" + i.toString() + "_rdesc";
          document.getElementById(text).disabled = false;
          if(document.getElementById("tier1_buydesc")!=null)
          {
            text = "tier" + i.toString() + "_buydesc";
            document.getElementById(text).disabled = false;
          }
        } 
      }
    }
</script>
<%
done:
    MyCommon.Close_LogixRT()
    Send_BodyEnd("mainform", "functioninput")
    Logix = Nothing
    MyCommon = Nothing
%>