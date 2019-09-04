<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Collections" %>
<%
    ' *****************************************************************************
    ' * FILENAME: CPEoffer-rew-discount.aspx 
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
    Dim AdminUserID As Long
    Dim MyCommon As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    Dim Localization As Copient.Localization
    Dim rst As DataTable
    Dim rst2 As DataTable
    Dim row As DataRow
    Dim CondDT As DataTable
    Dim OfferID As Long
    Dim DeliverableID As Long
    Dim Name As String = ""
    Dim RewardID As String
    Dim BestDealChecked As String = ""
    Dim AllowNegChecked As String = ""
    Dim DiscountedProductGroupID As Long
    Dim ExcludedProductGroupID As Long
    Dim AmountTypeID As AmountType_t
    Dim L1Cap As New Decimal
    Dim L2DiscountAmt As New Decimal
    Dim L2AmountTypeID As AmountType_t
    Dim L2Cap As New Decimal
    Dim L3DiscountAmt As New Decimal
    Dim L3AmountTypeID As AmountType_t
    Dim DiscountAmount As New Decimal
    Dim DecliningBalance As Boolean
    Dim RetroactiveDiscount As Boolean
    Dim ChargebackDeptID As Integer = -1
    Dim BestDeal As Integer
    Dim AllowNegative As Integer
    Dim ComputeDiscount As Integer = 0
    Dim ItemLimit As Integer
    Dim WeightLimit As Double
    Dim DollarLimit As Double
    Dim AnyProduct As Boolean
    Dim DiscountBarcode As String
    Dim VoidBarcode As String
    Dim RDesc As String = "" 'Receipt description
    Dim eDiscountType As Integer  '1=Marsh style, 2=Specified PLU, 3=IBM serial integration, 4=IBM TCP/IP style
    Dim UserGroupID As Long
    Dim DeptLevel As Integer
    Dim DiscountID As Long
    Dim ErrorMsg As String = ""
    Dim Phase As Integer
    Dim CustomerGroupSelOpt As String = ""
    Dim ProductGroupSelOpt As String = ""
    Dim ExcludedPGSelOpt As String = ""
    Dim DiscountType As Integer
    Dim DiscTypeSel As String = ""
    Dim PgDisabled As String = ""
    Dim PgDblClick As String = ""
    Dim DeslctPgDblClick As String = ""
    Dim TouchPoint As Integer = 0
    Dim TpROID As Integer = 0
    Dim CreateROID As Integer = 0
    Dim bCreated As Boolean = False
    Dim IsEditable As Boolean = False
    Dim Disallow_Edit As Boolean = True
    Dim FromTemplate As Boolean = False
    Dim IsTemplate As Boolean = False
    Dim IsTemplateVal As String = "Not"
    Dim DisabledAttribute As String = ""
    Dim LockFieldsList As String()
     Dim BuyDesc As String = ""
    Dim i As Integer
    Dim t As Integer
    Dim OverrideFields As New Hashtable()
    Dim OvrdFldEditable As Boolean = False
    Dim OvrdFldDisabled As String = "False"
    Dim OvrdFldClass As String = ""
    Dim CloseAfterSave As Boolean = False
    Dim DeferCalcToEOSChanged As Boolean = False
    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False
    Dim SVProgramID As Integer
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
    Dim FlexNegSO As Boolean = False
    Dim GlobalDepts As String = ""
    Dim TierLevels As Integer = 0
    Dim t1, t2 As Decimal
    Dim ValidTier As Boolean = False
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
    Dim l As Integer = 0
    Dim x As Integer = 0
    Dim ChargebackSet As Boolean = False
    Dim LoadDefaultChargeback As Boolean = True
    Dim RECORD_LIMIT As Integer = GroupRecordLimit '500
    Dim BundleWorthy As Boolean = False
    Dim PercentFixedRounding As Decimal = 0
          
    Dim MLI As New Copient.Localization.MultiLanguageRec
    Dim rstMsgDetails As DataTable
    Dim rowMsgDetails As DataRow
  
    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If
  
    Response.Expires = 0
    MyCommon.AppName = "CPEoffer-rew-discount.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
    Localization = New Copient.Localization(MyCommon)
          
    BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
    BestDealDefaulted = (MyCommon.Fetch_CPE_SystemOption(17) = "1")
    'Get the flex negative system option to use as default
    FlexNegSO = (MyCommon.Fetch_CPE_SystemOption(15) = "1")
    AllowNegSO = (MyCommon.Fetch_CPE_SystemOption(18) = "1")

    OfferID = Request.QueryString("OfferID")
    RewardID = Request.QueryString("RewardID")
    DiscountID = Request.QueryString("DiscountID")
    DeliverableID = MyCommon.Extract_Val(Request.QueryString("DeliverableID"))
    AmountTypeID = MyCommon.Extract_Val(Request.QueryString("l1amounttypeid"))
    If Request.QueryString("loadDefaultChargeback") <> "" Then
        LoadDefaultChargeback = (Request.QueryString("loadDefaultChargeback") = "1")
    End If

    If (Request.QueryString("EngineID") <> "") Then
        EngineID = MyCommon.Extract_Val(Request.QueryString("EngineID"))
    Else
        MyCommon.QueryStr = "select EngineID from OfferIDs where OfferID=" & OfferID & ";"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
            EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), 0)
        End If
    End If
          
    ' Get the tier levels
    MyCommon.QueryStr = "select TierLevels from CPE_RewardOptions where RewardOptionID=" & RewardID & ";"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
        TierLevels = MyCommon.NZ(rst.Rows(0).Item("TierLevels"), 0)
    End If
          
    Phase = MyCommon.Extract_Val(Request.QueryString("phase"))
    If (Phase = 0) Then Phase = MyCommon.Extract_Val(Request.Form("Phase"))
    If (Phase = 0) Then Phase = 3
          
    TouchPoint = MyCommon.Extract_Val(Request.QueryString("tp"))
    If (TouchPoint > 0) Then TpROID = MyCommon.Extract_Val(Request.QueryString("roid"))
          
 
          
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
          
    If IsTemplate Then
        IsEditable = Logix.UserRoles.EditTemplates
    Else
        IsEditable = Logix.UserRoles.EditOffer
    End If
   
   
    '----------------------------------------------------------------
    ' Begin "save"
    '----------------------------------------------------------------

    If (OfferID > 0 AndAlso (Request.QueryString("save") <> "")) Then
        'discount amount validation
        For t = 1 To TierLevels
            If AmountTypeID <> AmountType_t.AMT_TYPE_FREE AndAlso AmountTypeID <> AmountType_t.AMT_TYPE_STORED_VALUE AndAlso AmountTypeID <> AmountType_t.AMT_TYPE_SPECIAL_PRICING Then
                If (Not IsNumeric(Request.QueryString("tier" & t & "_l1discountamt"))) Then
                    infoMessage = Copient.PhraseLib.Lookup("CPEoffer-rew-discount.InvalidAmt", LanguageID)
                    Exit For
                ElseIf (Request.QueryString("tier" & t & "_l1discountamt") = 0) Then
                    infoMessage = Copient.PhraseLib.Lookup("term.IWIntegerError", LanguageID) & " " & Copient.PhraseLib.Lookup("term.discountamount", LanguageID).ToLower()
                    Exit For
                End If
            End If
        Next
        'Tier level validation code
        If TierLevels > 1 AndAlso AmountTypeID <> AmountType_t.AMT_TYPE_FREE AndAlso AmountTypeID <> AmountType_t.AMT_TYPE_SPECIAL_PRICING Then
            'Run validation for everything except "free" and "special pricing"
            If AmountTypeID = AmountType_t.AMT_TYPE_PRICE_POINT OrElse AmountTypeID = AmountType_t.AMT_TYPE_PRICE_POINT_WV Then
                'Price points
                For t = 2 To TierLevels
                    t2 = MyCommon.Extract_Val(Request.QueryString("tier" & t & "_l1discountamt"))
                    t1 = MyCommon.Extract_Val(Request.QueryString("tier" & t - 1 & "_l1discountamt"))
                    If (t2 <= t1) OrElse (t1 = 0 AndAlso t2 = 0) Then
                        ValidTier = True
                        WriteTier = True
                    Else
                        ValidTier = False
                        WriteTier = False
                        Exit For
                    End If
                Next
            Else
                'Fixed amount off, fixed percentage off, or stored value
                For t = 2 To TierLevels
                    t2 = MyCommon.Extract_Val(Request.QueryString("tier" & t & "_l1discountamt"))
                    t1 = MyCommon.Extract_Val(Request.QueryString("tier" & t - 1 & "_l1discountamt"))
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
        If MyCommon.Extract_Val(Request.QueryString("l1amounttypeid")) = AmountType_t.AMT_TYPE_SPECIAL_PRICING Then
            ' Special pricing validation: all levels must be numeric, and the number of levels must be <= the item limit
            For t = 1 To TierLevels
                SPLevels = MyCommon.Extract_Val(Request.QueryString("tier" & t & "_levels"))
                SPHighestLevel = MyCommon.Extract_Val(Request.QueryString("tier" & t & "_highestlevel"))
                SPItemLimit = MyCommon.Extract_Val(Request.QueryString("tier" & t & "_itemlimit"))
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
                        If Request.QueryString("tier" & t & "_level" & l) <> Nothing Then
                            If IsNumeric(Request.QueryString("tier" & t & "_level" & l)) Then
                                Dim amt As Decimal = 0
                                If Decimal.TryParse(Request.QueryString("tier" & t & "_level" & l), amt) Then
                                    ValidLevels = True
                                Else
                                    ValidLevels = False
                                    infoMessage = Copient.PhraseLib.Lookup("error.splevelvalues", LanguageID)
                                    Exit For
                                End If
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
        If Request.QueryString("discountType") <> "" Then
            DiscountType = MyCommon.Extract_Val(Request.QueryString("discountType"))
        Else
            DiscountType = 0
        End If
        If AmountTypeID = AmountType_t.AMT_TYPE_PERCENT_OFF Then
            If (Request.QueryString("percentfixedrounding") <> "") Then
                If IsNumeric(Request.QueryString("percentfixedrounding")) Then
                    If (MyCommon.Extract_Val(Request.QueryString("percentfixedrounding")) >= 0) AndAlso (MyCommon.Extract_Val(Request.QueryString("percentfixedrounding")) < 1) Then
                        ValidLevels = True
                    Else
                        ValidLevels = False
                        infoMessage = Copient.PhraseLib.Lookup("CPEoffer-rew-discount.InvalidRounding", LanguageID)
                    End If
                Else
                    ValidLevels = False
                    infoMessage = Copient.PhraseLib.Lookup("CPEoffer-rew-discount.InvalidRounding", LanguageID)
                End If
            End If
        End If
        If DiscountType = 1 Then
            For t = 1 To TierLevels
                If (Request.QueryString("tier" & t & "_itemlimit") <> "") Then
                    If Integer.TryParse(Request.QueryString("tier" & t & "_itemlimit"), Nothing) AndAlso (Request.QueryString("tier" & t & "_itemlimit") >= 0) Then
                        ValidLimits = True
                    Else
                        ValidLimits = False
                        infoMessage = Copient.PhraseLib.Detokenize("ueoffer-rew-discount.InvalidItemLimit", LanguageID, Request.QueryString("tier" & t & "_itemlimit"))
                        Exit For
                    End If
                Else
                    ValidLimits = False
                    infoMessage = Copient.PhraseLib.Detokenize("CPEoffer-rew-discount.BlankItemlimit", LanguageID, Request.QueryString("tier" & t & "_itemlimit"))
                    Exit For
                End If
            Next
        Else
            ValidLimits = True
        End If
                
        If Request.QueryString("discountedpgid") <> "" Then
            DiscountedProductGroupID = MyCommon.Extract_Val(Request.QueryString("discountedpgid"))
        Else
            DiscountedProductGroupID = 0
        End If
            
        If Request.QueryString("save") <> "" Then
            If DiscountedProductGroupID = 0 And DiscountType <> 4 Then
                DiscountedProductGroupID = -1
            End If
        End If

        If ValidTier AndAlso ValidLevels AndAlso ValidLimits AndAlso DiscountedProductGroupID > -1 AndAlso infoMessage = "" Then
            If (DiscountID <= 0) Then
                DiscountID = Create_Discount(OfferID, TpROID, Phase, DeliverableID)
                bCreated = True
                MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Reward.creatediscount", LanguageID))
            End If
              
            ' Store the product group and excluded product group for later possible TCRM comparison
            MyCommon.QueryStr = "select DiscountedProductGroupID, ExcludedProductGroupID from CPE_Discounts with (NoLock) where DiscountID=" & DiscountID
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count > 0) Then
                DiscountedProductGroupID = MyCommon.NZ(rst.Rows(0).Item("DiscountedProductGroupID"), -1)
                ExcludedProductGroupID = MyCommon.NZ(rst.Rows(0).Item("ExcludedProductGroupID"), -1)
            End If
              
            ' Save the contents of the discount
            If ValidLevels Then
        
                Try
                    ' DISCOUNT SAVES HERE *~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*
                    Save_Discount(OfferID, DiscountID, TierLevels, WriteTier, bCreated)
                    ' *~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*
                    If (Not bCreated) Then
                        ' Update for TCRM
                        ' Determine if the product group has changed; if so, flag s/b 3
                        If (DiscountedProductGroupID <> MyCommon.Extract_Val(Request.QueryString("discountedpgid")) _
                           OrElse ExcludedProductGroupID <> MyCommon.Extract_Val(Request.QueryString("excludedpgid"))) Then
                            TCRMAStatusFlag = 3
                        Else
                            TCRMAStatusFlag = 2
                        End If
                        ' If TCRMAStatusFlag is already 3 then don't change it to 2.
                        MyCommon.QueryStr = "update CPE_Deliverables with (RowLock) set TCRMAStatusFlag=" & TCRMAStatusFlag & " " & _
                                            "where TCRMAStatusFlag <> 3 and DeliverableID=" & DeliverableID
                        MyCommon.LRT_Execute()
                  
                        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Reward.editdiscount", LanguageID))
                    End If
                    If (Request.QueryString("save") <> "") Then
                        CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
                    End If
                Catch ex As Exception
                    infoMessage = ex.ToString()
                End Try
            End If
              
        Else
            If infoMessage <> "" Then
                'infoMessage already set
            ElseIf DiscountedProductGroupID = -1 Then
                infoMessage = Copient.PhraseLib.Lookup("reward.groupselect", LanguageID)
            ElseIf AmountTypeID = AmountType_t.AMT_TYPE_PRICE_POINT OrElse AmountTypeID = AmountType_t.AMT_TYPE_PRICE_POINT_WV Then
                infoMessage = Copient.PhraseLib.Lookup("error.tiervaluesdecrease", LanguageID)
            ElseIf Not ValidTier Then
                infoMessage = Copient.PhraseLib.Lookup("error.tiervalues", LanguageID)
            ElseIf AmountTypeID = AmountType_t.AMT_TYPE_SPECIAL_PRICING AndAlso Not ValidLevels Then
                If TierLevels = 1 Then
                    infoMessage = Copient.PhraseLib.Lookup("error.nopricepoint", LanguageID)
                ElseIf TierLevels > 1 Then
                    infoMessage = Copient.PhraseLib.Lookup("error.tiersnopricepoint", LanguageID)
                End If
            Else
                infoMessage = Copient.PhraseLib.Lookup("error.general", LanguageID)
            End If
        End If
    End If
  
    '----------------------------------------------------------------
    ' End "save"
    '----------------------------------------------------------------
          
    AnyProduct = False
    UserGroupID = 0
    If DiscountID = 0 Then
        DiscountID = MyCommon.Extract_Val(Request.QueryString("DiscountID"))
    End If
          
    eDiscountType = 4
          
    MyCommon.QueryStr = "select Name, DiscountTypeID, ReceiptDescription, SpecifyBarcode, DiscountBarcode, VoidBarcode, DiscountedProductGroupID, " & _
                        "ExcludedProductGroupID, BestDeal, AllowNegative, ComputeDiscount, DiscountAmount, AmountTypeID, " & _
                        "L1Cap, L2DiscountAmt, L2AmountTypeID, L2Cap, L3DiscountAmt, L3AmountTypeID, ItemLimit, WeightLimit, DollarLimit, ChargebackDeptID, " & _
                        "DecliningBalance, RetroactiveDiscount, UserGroupID, LastUpdate, SVProgramID, FlexNegative, ScorecardID, ScorecardDesc, PercentFixedRounding " & _
                        "from CPE_Discounts with (NoLock) where Deleted=0 and DiscountID=" & DiscountID & ";"
    rst = MyCommon.LRT_Select
    If Not (rst.Rows.Count = 0) Then
        Name = MyCommon.NZ(rst.Rows(0).Item("Name"), "")
        DiscountType = MyCommon.NZ(rst.Rows(0).Item("DiscountTypeID"), 0)
        DeptLevel = IIf(DiscountType = 2, 1, 0)
        RDesc = MyCommon.NZ(rst.Rows(0).Item("ReceiptDescription"), "")
        DiscountBarcode = MyCommon.NZ(rst.Rows(0).Item("DiscountBarcode"), "")
        VoidBarcode = MyCommon.NZ(rst.Rows(0).Item("VoidBarcode"), "")
        DiscountedProductGroupID = MyCommon.NZ(rst.Rows(0).Item("DiscountedProductGroupID"), 0)
        ExcludedProductGroupID = MyCommon.NZ(rst.Rows(0).Item("ExcludedProductGroupID"), 0)
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
        DiscountAmount = Math.Round(MyCommon.NZ(rst.Rows(0).Item("DiscountAmount"), 0), 3)
        AmountTypeID = MyCommon.NZ(rst.Rows(0).Item("AmountTypeID"), 0)
        L1Cap = Math.Round(MyCommon.NZ(rst.Rows(0).Item("L1Cap"), 0), 3)
        L2DiscountAmt = Math.Round(MyCommon.NZ(rst.Rows(0).Item("L2DiscountAmt"), 0), 3)
        L2AmountTypeID = MyCommon.NZ(rst.Rows(0).Item("L2AmountTypeID"), 0)
        L2Cap = Math.Round(MyCommon.NZ(rst.Rows(0).Item("L2Cap"), 0), 3)
        L3DiscountAmt = Math.Round(MyCommon.NZ(rst.Rows(0).Item("L3DiscountAmt"), 0), 3)
        L3AmountTypeID = MyCommon.NZ(rst.Rows(0).Item("L3AmountTypeID"), 0)
        ItemLimit = MyCommon.NZ(rst.Rows(0).Item("ItemLimit"), 1)
        WeightLimit = MyCommon.NZ(rst.Rows(0).Item("WeightLimit"), 0)
        DollarLimit = MyCommon.NZ(rst.Rows(0).Item("DollarLimit"), 1)
        ChargebackDeptID = MyCommon.NZ(rst.Rows(0).Item("ChargebackDeptID"), 0)
        ChargebackSet = True
        LoadDefaultChargeback = False
        DecliningBalance = MyCommon.NZ(rst.Rows(0).Item("DecliningBalance"), 0)
        RetroactiveDiscount = MyCommon.NZ(rst.Rows(0).Item("RetroactiveDiscount"), False)
        UserGroupID = MyCommon.NZ(rst.Rows(0).Item("UserGroupID"), 0)
        SVProgramID = MyCommon.NZ(rst.Rows(0).Item("SVProgramID"), 0)
        FlexNeg = MyCommon.NZ(rst.Rows(0).Item("FlexNegative"), False)
        ScorecardID = MyCommon.NZ(rst.Rows(0).Item("ScorecardID"), 0)
        ScorecardDesc = MyCommon.NZ(rst.Rows(0).Item("ScorecardDesc"), "")
        PercentFixedRounding = MyCommon.NZ(rst.Rows(0).Item("PercentFixedRounding"), 0)
    End If
    If (Request.QueryString("hdnallownegative") <> Nothing) Then
        If (Request.QueryString("hdnallownegative") <> "") Then
            AllowNegative = CMS.ExtentionMethods.ConvertToInt32(Request.QueryString("hdnallownegative"))
        Else
            AllowNegative = IIf(AllowNegSO = True, 1, 0)
        End If
    End If
    If (Request.QueryString("hdnflexnegative") <> Nothing) Then
        If (Request.QueryString("hdnflexnegative") <> "") Then
            FlexNeg = IIf(Request.QueryString("hdnflexnegative") = "1", True, False)
        Else
            FlexNeg = FlexNegSO
        End If
    End If
  
    If (AllowNegative = 1) Then
        FlexNeg = False
    End If
    ' Ensure that a manufacturer coupon does not have best deal selected
    If BestDeal = 1 AndAlso IsMfgCoupon Then
        MyCommon.QueryStr = "update CPE_Discounts with (RowLock) set BestDeal = 0 where DiscountID=" & DiscountID
        MyCommon.LRT_Execute()
    End If
  
    If Not (ErrorMsg = "") OrElse (Request.QueryString("mode") = "savediscount") Then
        DiscountedProductGroupID = MyCommon.Extract_Val(Request.QueryString("discountedpgid"))
        If DiscountedProductGroupID > 0 Then
            ComputeDiscount = 1
        End If
        ExcludedProductGroupID = MyCommon.Extract_Val(Request.QueryString("excludedpgid"))
        AmountTypeID = MyCommon.Extract_Val(Request.QueryString("l1amounttypeid"))
        ItemLimit = MyCommon.Extract_Val(Request.QueryString("itemlimit"))
        WeightLimit = MyCommon.Extract_Val(Request.QueryString("weightlimit"))
        DollarLimit = MyCommon.Extract_Val(Request.QueryString("dollarlimit"))
        DiscountAmount = Request.QueryString("discountamount")
        L1Cap = MyCommon.Extract_Val(Request.QueryString("l1cap"))
        L2DiscountAmt = MyCommon.Extract_Val(Request.QueryString("tier1_l2discountamt"))
        L2AmountTypeID = MyCommon.Extract_Val(Request.QueryString("l2amounttypeid"))
        L2Cap = MyCommon.Extract_Val(Request.QueryString("l2cap"))
        L3DiscountAmt = MyCommon.Extract_Val(Request.QueryString("tier1_l3discountamt"))
        L3AmountTypeID = MyCommon.Extract_Val(Request.QueryString("l3amounttypeid"))
        ChargebackDeptID = MyCommon.Extract_Val(Request.QueryString("chargeback"))
        DiscountBarcode = MyCommon.Extract_Val(Request.QueryString("discountbarcode"))
        VoidBarcode = MyCommon.Extract_Val(Request.QueryString("voidbarcode"))
        RDesc = Request.QueryString("rdesc")
        UserGroupID = MyCommon.Extract_Val(Request.QueryString("usergroupid"))
        BestDeal = MyCommon.Extract_Val(Request.QueryString("bestdeal"))
 
        If Request.QueryString("decliningbalance") = "true" Then
            DecliningBalance = True
        Else
            DecliningBalance = False
        End If
        If Request.QueryString("retrodiscount") = "true" Then
            RetroactiveDiscount = True
        Else
            RetroactiveDiscount = False
        End If
        DiscountType = MyCommon.Extract_Val(Request.QueryString("discountType"))
        DeptLevel = IIf(DiscountType = 2, 1, 0)
        'If DiscountType = 3 Then
        '    DeptLevel = 3 ' IIf(DiscountType = 3, 1, 0)
        'ElseIf DiscountType = 2 Then
        '    DeptLevel = 2 'IIf(DiscountType = 2, 1, 0)
        'Else
        '    DeptLevel = 1
        'End If
        SVProgramID = MyCommon.Extract_Val(Request.QueryString("svprogramid"))
        PercentFixedRounding = MyCommon.Extract_Val(Request.QueryString("percentfixedrounding"))
    End If
  
    AnyProduct = (DiscountedProductGroupID = 1)
  
    ' Update the templates permission if necessary
    If (Request.QueryString("save") <> "" AndAlso IsTemplate) Then
        ' time to update the status bits for the templates
        Dim form_Disallow_Edit As Integer = 0
        If (Request.QueryString("Disallow_Edit") = "on") Then
            form_Disallow_Edit = 1
        End If
        MyCommon.QueryStr = "update CPE_Deliverables with (RowLock) set DisallowEdit=" & form_Disallow_Edit & " " & _
                            "where DeliverableID=" & DeliverableID & ";"
        MyCommon.LRT_Execute()
        'clear the template field exception permissions
        MyCommon.QueryStr = "delete from TemplateFieldPermissions with (RowLock) where OfferID=" & OfferID & " " & _
                            "and FieldID in (select FieldID from UIFields where PageName='" & MyCommon.AppName & "');"
        MyCommon.LRT_Execute()
        If (Request.QueryString("chkTempField") <> "") Then
            Dim tmpFldLen As Integer = Request.QueryString.GetValues("chkTempField").Length
            If (tmpFldLen > 0) Then
                ReDim LockFieldsList(tmpFldLen)
                LockFieldsList = Request.QueryString.GetValues("chkTempField")
                For i = 0 To LockFieldsList.Length - 1
                    MyCommon.QueryStr = "insert into TemplateFieldPermissions with (RowLock) (OfferID, FieldID, Editable) " & _
                                        "values (" & OfferID & ", " & LockFieldsList(i) & "," & form_Disallow_Edit & ");"
                    MyCommon.LRT_Execute()
                Next
            End If
        End If
    End If
          
    If (IsTemplate Or FromTemplate) Then
        ' Get the permissions if it's a template
        MyCommon.QueryStr = "select DisallowEdit from CPE_Deliverables with (NoLock) where DeliverableID=" & DeliverableID
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
            Disallow_Edit = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit"), False)
        Else
            Disallow_Edit = False
        End If
        ' Check field-level permissions
        MyCommon.QueryStr = "select UI.FieldID, ISNull(TFP.Editable, 0) as Editable, UI.ControlName from UIFields UI with (NoLock) left join TemplateFieldPermissions TFP with (NoLock) on UI.FieldID = TFP.FieldID " & _
                            "where OfferID = " & OfferID & " and UI.PageName = '" & MyCommon.AppName & "';"
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
            For Each row In rst.Rows
                If (Not OverrideFields.ContainsKey(MyCommon.NZ(row.Item("ControlName"), ""))) Then
                    OverrideFields.Add(MyCommon.NZ(row.Item("ControlName"), ""), MyCommon.NZ(row.Item("Editable"), False))
                    If (MyCommon.NZ(row.Item("Editable"), False) = True) Then OvrdFldEditable = True
                End If
            Next
        End If
    End If
          
    If Not IsTemplate Then
        DisabledAttribute = IIf(Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Edit), "", " disabled=""disabled""")
    Else
        DisabledAttribute = IIf(Logix.UserRoles.EditTemplates, "", " disabled=""disabled""")
    End If
    SetDisabledAttr(DisabledAttribute)
          
    Send_HeadBegin("term.offer", "term.discountreward", OfferID)
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    If MyCommon.Fetch_SystemOption(135) = 1 And MyCommon.Fetch_SystemOption(124) = 0 Then
        Send_Scripts(New String() {"datePicker.js", "popup.js"})
    Else
        Send_Scripts()
    End If
%>
<style type="text/css">
    #foldercreate
    {
        left: 325px;
        top: 152px;
    }
    #folderstatusbar
    {
        left: 16px;
        top: 478px;
        width: 300px;
    }
    * html #folderstatusbar
    {
        left: 16px;
        top: 474px;
        width: 300px;
    }
    #searchResults
    {
        width: 99%;
    }
</style>
<script type="text/javascript" language="javascript">

  var selectedrecipttextbox = '';

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
	function assigntextboxcontrol(cntrl, elemName, shown) {
	  selectedrecipttextbox = cntrl.name + '_rdesc';
      var descontrol = document.getElementById(selectedrecipttextbox);
	  var receiptmsgElem = document.getElementById("receiptmsgerror");
      receiptmsgElem.style.display = 'none';
      var elem = document.getElementById(elemName);
      var fadeElem = document.getElementById('fadeDiv');
    
      if (elem != null) {
        elem.style.display = (shown) ? 'block' : 'none';
      }
      if (fadeElem != null) {
        fadeElem.style.display = (shown) ? 'block' : 'none';
      }
    }
    function placerecmessage(mlClickedID, event) {
      var defaultInput = document.getElementById('ml_' + mlClickedID + '_default').firstChild;
      // var stndInput = document.getElementById(mlClickedID);
      // stndInput.value = defaultInput.options[defaultInput.selectedIndex].text;
      xmlhttpPost_recText('OfferFeeds.aspx', 'Mode=PredefinedReceiptText&baselangtext=' + defaultInput.value + '&mlClickedID=' + mlClickedID,'PredefinedReceiptText');
    }
 
    function xmlhttpPost_recText(strURL, qryStr, action) {
      //alert(qryStr);
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
      self.xmlHttpReq.onreadystatechange = function() {
      if (self.xmlHttpReq !=null && self.xmlHttpReq.readyState == 4) {
        if (action == 'PredefinedReceiptText') {
          placepredefinedrecmsginML(self.xmlHttpReq.responseText);
        }
       }
	  }
      self.xmlHttpReq.send('<%sendb("LanguageID=" & LanguageID) %>');
      return false;
    }

    function placepredefinedrecmsginML(response) {
      //alert(response);
	 
      var sLanguageText, sLangRecMsgs, mlClickedID,cntrl,cntrlname;
      var trailerPos = -1;
      trailerPos = response.indexOf('<ML>');
      mlClickedID = response.substring(trailerPos + 4, response.indexOf("<\/ML>", trailerPos));

      trailerPos = -1;
      trailerPos = response.indexOf('<LangIds>');
      sLanguageIDs = response.substring(trailerPos + 9, response.indexOf("<\/LangIds>", trailerPos));
       
      trailerPos = -1;
      trailerPos = response.indexOf('<LangRecMsgs>');
      sLangRecMsgs = response.substring(trailerPos + 13, response.indexOf("<\/LangRecMsgs>", trailerPos));
	   
      var langIdArry = new Array();
      var langRecMsgArry = new Array();
       
      langIdArry = sLanguageIDs.split(String.fromCharCode(20));
      langRecMsgArry = sLangRecMsgs.split(String.fromCharCode(20));
       
      cntrl = document.getElementById(mlClickedID);
      for (i=0; i<langIdArry.length; i++)  {
        var cntrlname = cntrl.name + '_'+ langIdArry[i];
		if (sLangRecMsgs != '')
		  document.getElementById(cntrlname).value = langRecMsgArry[i];
		else
          document.getElementById(cntrlname).value = '';
       }
    }
	
  function placerecmessageLocal(elemName)
  {
    var elem = document.getElementById(elemName);
    var selectedrectext = document.getElementById('BaseMessages');
    var descontrol = document.getElementById(selectedrecipttextbox);
       
   	//if (selectedrectext.value != '')
	//{
	  descontrol.value = selectedrectext.value;
	//  // hide the receipt messages popup
	  toggleDialog('foldercreate',false);
	//}
    //else
    //{
    //  showreceiptmsgerror('Invalid receipt text');
    //} 
  }	

  function showreceiptmsgerror(content){
    var receiptmsgElem = document.getElementById("receiptmsgerror");
    receiptmsgElem.style.display = 'block';
    receiptmsgElem.innerHTML = content;
  }
      // create string buffer class to efficiently concatenate strings
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


      // This is the function that refreshes the list after a keypress.
      // The maximum number to show can be limited to improve performance with
      // huge lists (1000s of entries).
      // The function clears the list, and then does a linear search through the
      // globally defined array and adds the matches back to the list.
      function handleKeyUp(maxNumToShow) {
        var selectObj, textObj, functionListLength;
        var i,  numShown;
        var searchPattern;
        
        document.getElementById("functionselect").size = "10";
        
        // Set references to the form elements
        selectObj = document.forms[0].functionselect;
        textObj = document.forms[0].functioninput;
        
        // Remember the function list length for loop speedup
        functionListLength = functionlist.length;
        
        // Set the search pattern depending
        if(document.forms[0].functionradio[0].checked == true) {
          searchPattern = "^"+textObj.value;
        } else {
          searchPattern = textObj.value;
        }
        searchPattern = cleanRegExpString(searchPattern);
        
        // Create a regulare expression
        re = new RegExp(searchPattern,"gi");
        
        // Clear the options list
        selectObj = clearOptionsFast(selectObj);
        
        // Loop through the array and re-add matching options
        var buf = new StringBuffer();
        numShown = 0;
        for(i = 0; i < functionListLength; i++) {
          if(functionlist[i].search(re) != -1) {
            buf.append('<option value="' + vallist[i] + '">' + functionlist[i] + '<\/option>')
            numShown++;
          }
          // Stop when the number to show is reached
          if(numShown == maxNumToShow) {
            break;
          }
        }
        select_innerHTML(selectObj, buf.toString());
        selectObj = document.forms[0].functionselect;
        
        // When options list whittled to one, select that entry
        if(selectObj.length == 1) {
          selectObj.options[0].selected = true;
        }
      }

      function clearOptionsFast(selectObj) {
        var selectParentNode = selectObj.parentNode;
        var newSelectObj = selectObj.cloneNode(false);
        selectParentNode.replaceChild(newSelectObj, selectObj);
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

      <% If (IsEditable) Then %>
      function submitForm() {
        assignValues();
        document.mainform.submit();
      }

      function assignValues() {
        var elemPG = document.getElementById("discountedpgid");
        var elemExPG = document.getElementById("excludedpgid");
        var elemCG = document.getElementById("usergroupid");
        var elemSel = document.getElementById("selected");
          var allownegative= document.getElementById("allownegative");
        var flexnegative= document.getElementById("flexnegative");
        var hdnallownegative= document.getElementById("hdnallownegative");
        var hdnflexnegative= document.getElementById("hdnflexnegative");

        // Discounted Product Group
        if (elemSel.options.length > 0) {
          elemPG.value = elemSel.options[0].value;
        } else {
          elemPG.value = "";
        }
        
        // Excluded Product Group
        if (document.mainform.excluded != null && document.mainform.excluded.options.length > 0) {
          elemExPG.value = document.mainform.excluded.options[0].value;
        } else {
          elemExPG.value = "";
        }
        hdnallownegative.value="";
        hdnflexnegative.value="";
        if(allownegative!=null)
        {
            hdnallownegative.value=(allownegative.checked==true?1:0);
        }
        

        if(flexnegative!=null)
        {

          hdnflexnegative.value=(flexnegative.checked==true?1:0);
        }

        
        if (document.mainform.mode != null) {
          document.mainform.mode.value = "savediscount";
        }
        <% If (FromTemplate) Then %>
          enableFormFields(document.mainform);
        <% End If %>
      }
      <% Else %>
      function submitForm() { 
        sendNotEditable();
        return false;
      }

      function assignValues() {
        sendNotEditable();
        return false;
      }

      function sendNotEditable() {
        var saveElem = document.getElementById("save");
        alert('<% Sendb(Copient.PhraseLib.Lookup("CPE-rew-noedit", LanguageID))%>');
        if (saveElem != null) {
          if (saveElem.style.visibility=='hidden') {
            saveElem.style.visibility='visible';
          }
        }
      }
      <% End If %>

      function selectItem(source, dest) {
        var elemSource = document.getElementById(source);
        var elemDest = document.getElementById(dest);   
        var selOption = null;
        var selText ="", selVal = "";
        var selIndex = -1;
        
        if (elemSource != null && elemSource.options.selectedIndex == -1) {
          alert('<% Sendb(Copient.PhraseLib.Lookup("CPE-rew-discounts.selectproducts", LanguageID)) %>');
          elemSource.focus();
        } else {
          selIndex = elemSource.options.selectedIndex;
          selOption = elemSource.options[selIndex];
          selText = selOption.text;
          selVal = selOption.value;
          elemDest.options[0] = new Option(selText, selVal);
          xmlhttpProduct('OfferFeeds.aspx', 'DiscountProductGroups');
          //submitForm();
        }
      }

      function deselectItem(source) {
        var elemSource = document.getElementById(source);
        
        if (elemSource != null && elemSource.options.selectedIndex == -1) {
          alert('<% Sendb(Copient.PhraseLib.Lookup("CPE-rew-discounts.selectproducts", LanguageID)) %>');
          elemSource.focus();
        } else {
          elemSource.options[0] = null;
          xmlhttpProduct('OfferFeeds.aspx', 'DiscountProductGroups');
          //submitForm();
        }
      }

      function enableFormFields (theForm) {
        var elems = theForm.elements;
        var elem = null;

        for (var x = 0; x < elems.length; x++) {
            elem = elems[x];
            if (elem != null) {
                if (elem.disabled == true) {
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
          elemTrs  = elemTbl.getElementsByTagName("TR");
          for (var i=1; i < elemTrs.length; i++) {
            elemTd = elemTrs[i].firstChild;
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
              if (tr !=null) {
                var td3 = tr.lastChild;
                if (td3 != null) {
                  if (pageLockChecked) {
                    td3.innerHTML = (td3.innerHTML=='<% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID)) %>') ? '<% Sendb(Copient.PhraseLib.Lookup("term.unlocked", LanguageID)) %>' : '<% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID)) %>';
                  } else  {
                    td3.innerHTML = (elem.checked) ? '<% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID)) %>' : '<% Sendb(Copient.PhraseLib.Lookup("term.unlocked", LanguageID)) %>';
                  }
                }
              }
            }
          }
        }
      }

      function handleBasketLevel(selValue) {
        var elemSel = document.getElementById("selected");
        var elemList = document.getElementById("functionselect");
        var selectDiv = document.getElementById("selectDiv");
        var bundleDiv = document.getElementById("bundleDiv");
        
        if (parseInt(selValue) != 0) {
          document.getElementById('l1amounttypeid').options[0].selected = true;
        }
        if (parseInt(selValue) == 3) { 
          elemSel.options[0] = new Option('<% Sendb(Copient.PhraseLib.Lookup("term.anyproduct", LanguageID)) %>', '1');
        } else if (parseInt(selValue) == 4) {
          elemSel.options.length=0;
        } else {
          //if the discount type is changed, remove the selected product group
          if (elemSel.options.length > 0) {
            if (elemSel.options[0].value == 1) {
              elemSel.options.length=0;
            }
          }
        }
        if (parseInt(selValue) == 4) {
          selectDiv.style.display = 'none';
          bundleDiv.style.display = 'block';
        } else {
          selectDiv.style.display = 'block';
          bundleDiv.style.display = 'none';
        }
        submitForm();
      }

      function handleChargebackDept(defaultValue) {
        var elemChrg = document.getElementById("chargeback");
        
        if (elemChrg != null) {
          for (var i=0; i < elemChrg.options.length; i++) {
            if (elemChrg.options[i].value == defaultValue) {
              elemChrg.selectedIndex = i;
            }
          }
        }
        
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
        var elemDistType = document.getElementById("l1amounttypeid");
        
        if (elemList != null) {
          if (elemList.style.display == 'block') {
            elemList.style.display = 'none';
            if (elemDistType != null) { elemDistType.style.visibility = 'visible'; }
          } else {
            elemList.style.display = 'block';
            if (elemDistType != null) { elemDistType.style.visibility = 'hidden'; }
          }
        }
      }
      
      var isFireFox = (navigator.appName.indexOf('Mozilla') !=-1) ? true : false;
      var timer;
      function xmlPostTimer(strURL,mode)
      {
        clearTimeout(timer);
        timer=setTimeout("xmlhttpProduct('" + strURL + "','" + mode + "')", 250);
      }

      function xmlhttpProduct(strURL,mode) {
        var xmlHttpReq = false;
        var self = this;
        
        document.getElementById("searchLoadDiv").innerHTML = '<% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID)) %>';
        
        // Mozilla/Safari
        if (window.XMLHttpRequest) {
          self.xmlHttpReq = new XMLHttpRequest();
        }
        // IE
        else if (window.ActiveXObject) {
          self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
        }
        var qryStr = getproductquery(mode);
        self.xmlHttpReq.open('POST', strURL + "?" + qryStr, true);
        self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
        self.xmlHttpReq.onreadystatechange = function() {
          if (self.xmlHttpReq !=null && self.xmlHttpReq.readyState == 4) {
            updateproducts(self.xmlHttpReq.responseText);
          }
        }

        self.xmlHttpReq.send(qryStr);
        //self.xmlHttpReq.send(getquerystring());
      }

      function getproductquery(mode) {
        var radioString;
        if(document.getElementById('functionradio2').checked) {
          radioString = 'functionradio2';
        }
        else {
          radioString = 'functionradio1';
        }
        var selected = document.getElementById('selected');
        var selectedGroup = 0;
        if(selected.options[0] != null){ //Can only have one group selected at a time
          selectedGroup = selected.options[0].value;
        }
        return "Mode=" + mode + "&ProductSearch=" + document.getElementById('functioninput').value + "&SelectedGroup=" + selectedGroup + "&SearchRadio=" + radioString;
       
      }

      function updateproducts(str) {
      <%
        PgDblClick = IIf(DiscountType <> 3, "selectItem(\'functionselect\', \'selected\');submitForm();", "")
       %>
        if(str.length > 0){
          if(!isFireFox){
            document.getElementById("pgList").innerHTML = '<select class="longer" id="functionselect" name="functionselect" onclick="handleSelectClick();" ondblclick="<% Sendb(PgDblClick) %>" size="10"<% sendb(DisabledAttribute) %>>' + str + '<\/select>';
          }
          else{
            document.getElementById("functionselect").innerHTML = str;
          }
          document.getElementById("searchLoadDiv").innerHTML = '&nbsp;';
        }
        else if(str.length == 0){
          if(!isFireFox){
            document.getElementById("pgList").innerHTML = '<select class="longer" id="functionselect" name="functionselect" onclick="handleSelectClick();" ondblclick="<% Sendb(PgDblClick) %>" size="10"<% sendb(DisabledAttribute) %>><\/select>';
          }
          else{
            document.getElementById("functionselect").innerHTML = '';
          }
          document.getElementById("searchLoadDiv").innerHTML = '&nbsp;';
        }
      }

      function xmlhttpPost(strURL) {
        var xmlHttpReq = false;
        var self = this;
        
        document.getElementById("results").innerHTML = "<div class=\"loading\"><img id=\"clock\" src=\"../images/clock22.png\" \/><br \/>" + '<% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID)) %><\/div>';
        
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
        self.xmlHttpReq.onreadystatechange = function() {
          if (self.xmlHttpReq.readyState == 4) {
            updatepage(self.xmlHttpReq.responseText);
          }
        }
        self.xmlHttpReq.send('<%sendb("LanguageID=" & LanguageID) %>');
        return false;
      }

      function updatepage(str){
        document.getElementById("results").innerHTML = str;
      }

      function zeroLevel(level) {
        var levelName = 'l' + level + 'cap';
        var elem = document.getElementById(levelName);
        if (elem != null) elem.value = "0";
        
        var levelType = 'l' + level + 'amounttypeid'
        elem = document.getElementById(levelType)
        
        if (elem != null) {
          if (document.getElementById("discountType").value=="3") {
           elem.selectedIndex = 1; // set to percent off for Basket level
          } else {
            elem.selectedIndex = 2; // set it to percent off all other discountTypes
          }
        }
        
        submitForm();
      }

      function handleUpToEntry(val, Level) {
        var elemArrow = document.getElementById("btnDown" + Level);
        var elemEx = document.getElementById("btnEx" + Level);
        
        if (!isNaN(val) && parseFloat(val) > 0) {
          if (elemArrow != null) elemArrow.disabled = false;
          if (elemEx != null) elemEx.disabled = false;
        }  else {
          if (elemArrow != null) elemArrow.disabled = true;
          if (elemEx != null) elemEx.disabled = true;
        }
      }

      function addLevel(tier) {
        var levelsElem = document.getElementById('tier'+tier+'_levels');
        var levels = parseInt(levelsElem.value);
        
        var highestLevelElem = document.getElementById('tier'+tier+'_highestlevel');
        var highestLevel = parseInt(highestLevelElem.value);
        
        var itemLimitElem = document.getElementById('tier'+tier+'_itemlimit');
        var itemLimit = parseInt(itemLimitElem.value);
        
        var radioElem = document.getElementById('tier'+tier+'_sprepeatlevel'+(highestLevel + 1));
        var inputElem = document.getElementById('tier'+tier+'_level'+(highestLevel + 1));
        var buttonElem = document.getElementById('tier'+tier+'_deletelevel'+(highestLevel + 1));
        var rowElem = document.getElementById('tier'+tier+'_level'+(highestLevel + 1)+'row');
        var newLevelInput = document.getElementById('tier'+tier+'_newlevel');

        var saveElem = document.getElementById("save");
        
        if ((parseInt(itemLimitElem.value) > 0) && (levels >= itemLimit)) {
          alert('<%Sendb(Copient.PhraseLib.Lookup("ueoffer-rew-discount.ExceedsItemLimit", LanguageID).Replace("'", "\'")) %>');
        } else if (newLevelInput.value == '') {
          alert('<%Sendb(Copient.PhraseLib.Lookup("ueoffer-rew-discount.SpecifyValue", LanguageID))%>');
        } else if (newLevelInput.value == 0) {
            alert('<%Sendb(Copient.PhraseLib.Lookup("error.requires_valid_positive_integer", LanguageID))%>');
        }else {
          if (document.getElementById('tier'+tier+'_level'+(highestLevel + 1)+'row') != null) {
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

        if (saveElem != null) {
          if (saveElem.style.visibility=='hidden') {
            saveElem.style.visibility='visible';
          }
        }
      }

      function deleteLevel(tier, level) {
        var levelsElem = document.getElementById('tier'+tier+'_levels');
        var levels = parseInt(levelsElem.value);
        
        var highestLevelElem = document.getElementById('tier'+tier+'_highestlevel');
        var highestLevel = parseInt(highestLevelElem.value);
        
        var radioElem = document.getElementById('tier'+tier+'_sprepeatlevel'+level);
        var inputElem = document.getElementById('tier'+tier+'_level'+level);
        var buttonElem = document.getElementById('tier'+tier+'_deletelevel'+level);
        var rowElem = document.getElementById('tier'+tier+'_level'+level+'row');
        
        if (document.getElementById('tier'+tier+'_level'+level+'row') != null) {
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
            flexnegElem.checked = false;
            flexnegElem.disabled = true;
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
        
    If (OfferID > 0 AndAlso (Request.QueryString("save") <> "" OrElse Request.QueryString("mode") = "savediscount")) Then
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
        Send("  opener.location = 'CPEoffer-rew.aspx?OfferID=" & OfferID & "'; ")
    ElseIf (Phase = 1) Then
        Send("  opener.location = 'CPEoffer-not.aspx?OfferID=" & OfferID & "'; ")
    End If
    Send("} ")
    Send("</script>")
%>
<form action="CPEoffer-rew-discount.aspx" id="mainform" name="mainform" autocomplete="off"
onsubmit="return assignValues();">
<div id="results" style="position: absolute; z-index: 99; top: 31px; right: 21px;">
</div>
<div id="intro">
    <input type="hidden" id="OfferID" name="OfferID" value="<% sendb(OfferID) %>" />
    <input type="hidden" id="Name" name="Name" value="<% sendb(Name) %>" />
    <input type="hidden" id="DeliverableID" name="DeliverableID" value="<% sendb(DeliverableID) %>" />
    <input type="hidden" id="RewardID" name="RewardID" value="<% sendb(RewardID) %>" />
    <input type="hidden" id="DiscountID" name="DiscountID" value="<% sendb(DiscountID) %>" />
    <input type="hidden" id="Phase" name="Phase" value="<% sendb(Phase) %>" />
    <input type="hidden" id="discountedpgid" name="discountedpgid" value="<% Sendb(DiscountedProductGroupID)%>" />
    <input type="hidden" id="excludedpgid" name="excludedpgid" value="<% Sendb(ExcludedProductGroupID)%>" />
    <input type="hidden" id="usergroupid" name="usergroupid" value="<% Sendb (UserGroupID)%>" />
    <input type="hidden" id="mode" name="mode" value="" />
    <input type="hidden" id="roid" name="roid" value="<%Sendb(TpROID) %>" />
    <input type="hidden" id="tp" name="tp" value="<%Sendb(TouchPoint) %>" />
    <input type="hidden" id="IsTemplate" name="IsTemplate" value="<% Sendb(IsTemplateVal)%>" />
    <input type="hidden" id="hdnallownegative" name="hdnallownegative" value="<%=AllowNegative%>" />
    <input type="hidden" id="hdnflexnegative" name="hdnflexnegative" value="<%=FlexNeg%>" />
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
            If Not IsTemplate Then
                If (Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Edit)) Or (Logix.UserRoles.EditRoles And OvrdFldEditable) Then
                    If DiscountID = 0 Then
                        Send_Save(" onclick=""this.style.visibility='hidden';""")
                    Else
                        Send_Save()
                    End If
                End If
            Else
                If (Logix.UserRoles.EditTemplates) Or (Logix.UserRoles.EditRoles And OvrdFldEditable) Then
                    If DiscountID = 0 Then
                        Send_Save(" onclick=""this.style.visibility='hidden';""")
                    Else
                        Send_Save()
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
    <div id="column1">
        <div class="box" id="selector">
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.productgroup", LanguageID))%>
                </span>
            </h2>
            <label for="discountType">
                <% Sendb(Copient.PhraseLib.Lookup("term.discounttype", LanguageID))%>:</label>
            <select id="discountType" name="discountType" size="1" onchange="handleBasketLevel(this.value);"
                <% sendb(DisabledAttribute) %>>
                <%
                    'First determine if bundle price discounts are valid (by seeing if there are more than two and'ed conditional product groups)
                    MyCommon.QueryStr = "select distinct PG.ProductGroupID, PG.Name " & _
                                        " from CPE_IncentiveProductGroups as IPG with (NoLock) " & _
                                        " left join ProductGroups as PG with (NoLock) on PG.ProductGroupID=IPG.ProductGroupID " & _
                                        " left join CPE_RewardOptions as RO with (NoLock) on IPG.RewardOptionID=RO.RewardOptionID " & _
                                        " left join CPE_UnitTypes as UT with (NoLock) on UT.UnitTypeID=IPG.QtyUnitType " & _
                                        " where RO.IncentiveID=" & OfferID & " and IPG.Deleted=0 and Disqualifier=0 " & _
                                        " and ProductComboID=1 and QtyUnitType=1 and QtyForIncentive=1;"
                    CondDT = MyCommon.LRT_Select
                    If CondDT.Rows.Count > 1 And TierLevels = 1 Then
                        BundleWorthy = True
                    End If
                    'Build the discount types dropdown
                    MyCommon.QueryStr = "select DiscountTypeID, Name, PhraseID from CPE_DiscountTypes DT with (NoLock) " & _
                                        IIf(BundleWorthy, "", " where DiscountTypeID<>4") & _
                                        " order by DiscountTypeID;"
                    rst = MyCommon.LRT_Select()
                    If (rst.Rows.Count > 0) Then
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
                    PgDblClick = IIf(DiscountType <> 3, "selectItem('functionselect', 'selected');submitForm();", "")
                    DeslctPgDblClick = IIf(DiscountType <> 3, "deselectItem('selected');", "")
                %>
            </select>
            <br />
            <br />
            <div id="selectDiv" <%Sendb(IIf(DiscountType = 4, " style=""display:none;""", ""))%>>
                <input type="radio" id="functionradio1" name="functionradio" checked="checked" <% sendb(DisabledAttribute) %> /><label
                    for="functionradio1"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
                <input type="radio" id="functionradio2" name="functionradio" <% sendb(DisabledAttribute) %> /><label
                    for="functionradio2"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
                <%--<input type="text" class="medium" id="functioninput" name="functioninput" maxlength="100" onkeyup="handleKeyUp(200);" value=""<% sendb(DisabledAttribute) %> /><br />--%>
                <input type="text" class="medium" id="functioninput" name="functioninput" maxlength="100"
                    onkeyup="javascript:xmlPostTimer('OfferFeeds.aspx','DiscountProductGroups');"
                    value="" <% sendb(DisabledAttribute) %> /><br />
                <div id="searchLoadDiv" style="display: block;">
                    &nbsp;</div>
                <div id="pgList">
                    <select class="longer" id="functionselect" name="functionselect" onclick="handleSelectClick();"
                        ondblclick="<% Sendb(PgDblClick) %>" size="10" <% sendb(DisabledAttribute) %>>
                        <%
                            If DiscountedProductGroupID = 1 AndAlso DiscountType = 3 Then
                                AnyProduct = True
                            ElseIf DiscountType = 3 Then
                                Send("<option value=""1"">" & Copient.PhraseLib.Lookup("term.anyproduct", LanguageID) & "</option>")
                                AnyProduct = True
                            ElseIf DiscountType <> 3 AndAlso DiscountID = 0 Then
                                AnyProduct = False
                            End If
            
                            Dim topString As String = ""
                            If (RECORD_LIMIT > 0) Then topString = " top " & RECORD_LIMIT & " "
                    
                            MyCommon.QueryStr = "Select " & topString & " ProductGroupID, Name, PhraseID from ProductGroups with (NoLock) where deleted=0 and ProductGroupID <> 1 and NonDiscountableGroup=0 and PointsNotApplyGroup=0 order by ProductGroupID desc ,Name"
                            rst = MyCommon.LRT_Select
                            For Each row In rst.Rows
                                If MyCommon.NZ(row.Item("ProductGroupID"), 0) = DiscountedProductGroupID Then
                                    AnyProduct = False
                                Else
                                    Send("<option value=""" & row.Item("ProductGroupID") & """ title=""" & row.Item("Name") & """>" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
                                End If
                            Next
                        %>
                    </select>
                </div>
                <%If (RECORD_LIMIT > 0) Then
                        Send(Copient.PhraseLib.Lookup("groups.display", LanguageID) & ": " & RECORD_LIMIT.ToString() & "<br />")
                    End If
                %>
                <br class="half" />
                <input type="button" class="regular select" id="pselect" name="pselect" value="&#9660; <% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID))%>"
                    onclick="selectItem('functionselect', 'selected');submitForm();" <% sendb(DisabledAttribute) %><%sendb(pgdisabled) %> />&nbsp;
                <input type="button" class="regular deselect" name="pdeselect" id="pdeselect" value="<% Sendb(Copient.PhraseLib.Lookup("term.deselect", LanguageID))%> &#9650;"
                    onclick="deselectItem('selected');" <% sendb(DisabledAttribute) %><%sendb(pgdisabled) %> /><br />
                <br class="half" />
                <select class="longer" id="selected" name="selected" size="2" ondblclick="<% Sendb(DeslctPgDblClick) %>"
                    <% sendb(DisabledAttribute) %>>
                    <% 
                        If DiscountedProductGroupID = 1 Then
                            If DiscountType = 3 Then Send("<option value=""1"">" & Copient.PhraseLib.Lookup("term.anyproduct", LanguageID) & "</option>")
                        ElseIf DiscountedProductGroupID > 0 Then
                            MyCommon.QueryStr = "Select Name from ProductGroups where Deleted=0 and NonDiscountableGroup=0 and PointsNotApplyGroup=0 and ProductGroupID=" & DiscountedProductGroupID & ";"
                            Dim selectDT As DataTable = MyCommon.LRT_Select()
                            If selectDT.Rows.Count > 0 Then Send("<option value=""" & DiscountedProductGroupID & """>" & MyCommon.NZ(selectDT.Rows(0).Item("Name"), "") & "</option>")
                        End If
          
                        'Send(ProductGroupSelOpt)
                    %>
                </select>
                <% If AnyProduct And (eDiscountType = 1 Or eDiscountType = 3 Or eDiscountType = 4) Then%>
                <br />
                <br class="half" />
                <b>
                    <% Sendb(Copient.PhraseLib.Lookup("term.excludedproducts", LanguageID))%>:</b><br />
                <input type="button" class="select" name="select3" id="select3" value="&#9660; <% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID))%>"
                    onclick="selectItem('functionselect', 'excluded');" <% sendb(DisabledAttribute) %> />&nbsp;
                <input type="button" class="deselect" name="deselect2" id="deselect2" value="<% Sendb(Copient.PhraseLib.Lookup("term.deselect", LanguageID))%> &#9650;"
                    onclick="deselectItem('excluded');" <% sendb(DisabledAttribute) %> /><br />
                <br class="half" />
                <select class="longer" id="excluded" name="excluded" size="2" ondblclick="deselectItem('excluded');"
                    <% sendb(DisabledAttribute) %>>
                    <% 
                        If ExcludedProductGroupID > 0 Then
                            MyCommon.QueryStr = "Select Name from ProductGroups where Deleted=0 and NonDiscountableGroup=0 and PointsNotApplyGroup=0 and ProductGroupID=" & ExcludedProductGroupID & ";"
                            Dim excludeDT As DataTable = MyCommon.LRT_Select()
                            If excludeDT.Rows.Count > 0 Then Send("<option value=""" & ExcludedProductGroupID & """>" & MyCommon.NZ(excludeDT.Rows(0).Item("Name"), "") & "</option>")
                        End If
          
                        'Send(ExcludedPGSelOpt)
                    %>
                </select>
                <% End If%>
            </div>
            <div id="bundleDiv" <%Sendb(IIf(DiscountType = 4, "", " style=""display:none;"""))%>>
                <%
                    If CondDT.Rows.Count > 0 Then
                        Send(Copient.PhraseLib.Lookup("term.ConditionalProductGroups", LanguageID) & ":")
                        Send("<ul>")
                        For Each row In CondDT.Rows
                            Sendb("  <li>" & MyCommon.NZ(row.Item("Name"), "") & "</li>")
                        Next
                        Send("</ul>")
                    End If
                %>
            </div>
            <hr class="hidden" />
        </div>
        <div class="box" id="department">
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.chargebackdepartment", LanguageID))%>
                </span>
            </h2>
            <%
                If (eDiscountType = 1) Or (eDiscountType = 3) Or (eDiscountType = 4) Then
                    If DiscountType = 4 Then
                        sQuery = "select ChargeBackDeptID, ExternalID, Name, PhraseID from ChargebackDepts with (NoLock) where ChargeBackDeptID=10 "
                        GlobalDepts = "10"
                    ElseIf AnyProduct Then
                        sQuery = "select ChargeBackDeptID, ExternalID, Name, PhraseID from ChargebackDepts with (NoLock) where ChargeBackDeptID<>0 "
                        GlobalDepts = "10"
                    ElseIf DeptLevel > 0 Then
                        sQuery = "select ChargeBackDeptID, ExternalID, Name, PhraseID from ChargebackDepts with (NoLock) where ChargeBackDeptID<>0 and ChargeBackDeptID<>14 "
                        GlobalDepts = "10"
                    Else 'item level
                        sQuery = "select ChargeBackDeptID, ExternalID, Name, PhraseID from ChargebackDepts with (NoLock) where ChargeBackDeptID<>10 "
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
                            MyCommon.QueryStr = "select DefaultChargebackDeptID from Banners where BannerID = " & MyCommon.NZ(rst.Rows(0).Item("BannerID"), -1)
                            rst = MyCommon.LRT_Select
                            If (rst.Rows.Count = 1) Then
                                DefaultChrgBack = MyCommon.NZ(rst.Rows(0).Item("DefaultChargebackDeptID"), -1)
                            End If

                        Else
                            sQuery &= " and (BannerID = 0 or BannerID IS NULL) " & " or ChargebackDeptID in (" & GlobalDepts & ")"
                        End If
                    End If
                    MyCommon.QueryStr = sQuery & " order by ExternalID;"
                    rst = MyCommon.LRT_Select
          
                    Send("<input type=""hidden"" name=""loadDefaultChargeback"" id=""loadDefaultChargeback"" value=""" & IIf(LoadDefaultChargeback, "1", "0") & """ />")
                    Send("<select class=""longer"" id=""chargeback"" name=""chargeback"" onchange=""handleChargebackSubmit();""" & DisabledAttribute & ">")
                    If Not (rst.Rows.Count = 0) Then
                        If DefaultChrgBack = -1 And DiscountID = 0 Then
                            Select Case DiscountType
                                Case 0, 1 ' unset or item level
                                    DefaultChrgBack = MyCommon.Extract_Val(MyCommon.Fetch_CPE_SystemOption(116))
                                Case 2 ' dept level
                                    DefaultChrgBack = MyCommon.Extract_Val(MyCommon.Fetch_CPE_SystemOption(117))
                                Case 3 ' basket level
                                    DefaultChrgBack = MyCommon.Extract_Val(MyCommon.Fetch_CPE_SystemOption(118))
                            End Select
                        End If
                        For Each row In rst.Rows
                            Sendb("<option value=""" & MyCommon.NZ(row.Item("ChargeBackDeptID"), 0) & """")
                            If MyCommon.NZ(row.Item("ChargeBackDeptID"), -1) = ChargebackDeptID Then
                                Sendb(" selected=""selected""")
                            End If
                            Sendb(">")
                            If ((row.Item("ExternalID") = "") Or (row.Item("ExternalID") = "0")) Then
                            Else
                                Sendb(row.Item("ExternalID") & " - ")
                            End If
                            If (IsDBNull(row.Item("PhraseID"))) Then
                                Send(MyCommon.NZ(row.Item("Name"), ""))
                            Else
                                If (row.Item("PhraseID") = 0) Then
                                    Send(MyCommon.NZ(row.Item("Name"), ""))
                                Else
                                    Sendb(Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID))
                                End If
                            End If
                            Send("</option>")
                        Next
                        Send("</select>")
                    End If
                Else
                    Send("<input type=""hidden"" id=""chargeback"" name=""chargeback"" value=""1"" />")
                End If
      %>
      <br />
      <hr class="hidden" />
    </div>
  </div>
  <div id="gutter">
  </div>
  <div id="column2">
    <%
      Send("<div class=""box"" id=""distribution"" style=""z-index:51;"">")
      Send("  <h2>")
      Send("    <span>")
      Send("      " & Copient.PhraseLib.Lookup("term.distribution", LanguageID))
      Send("    </span>")
      Send("  </h2>")
      'Output eDiscountType and AmountType
      Send("  <table summary=""" & Copient.PhraseLib.Lookup("term.distribution", LanguageID) & """>")
      If eDiscountType = 2 Then
        Send("<tr>")
        Send("  <td>")
        Send("    <label for=""tier1_l1discountamt"">" & Copient.PhraseLib.Lookup("term.discountamount", LanguageID) & ":</label>")
        Send("  </td>")
        Send("  <td>")
        Send("    $<input type=""text"" id=""tier1_l1discountamt"" name=""tier1_l1discountamt"" value=""" & Format(DiscountAmount, "#####0.00") & """ size=""6"" maxlength=""6""" & DisabledAttribute & " />")
        Send("  </td>")
        Send("</tr>")
        Send("<tr>")
        Send("  <td>")
        Send("    <label for=""discountbarcode"">" & Copient.PhraseLib.Lookup("term.discountbarcode", LanguageID) & ":</label>")
        Send("  </td>")
        Send("  <td>")
        Send("    <input type=""text"" id=""discountbarcode"" name=""discountbarcode"" value=""" & DiscountBarcode & """ size=""30"" maxlength=""255""" & DisabledAttribute & " />")
        Send("  </td>")
        Send("</tr>")
        Send("<tr>")
        Send("  <td>")
        Send("    <label for=""voidbarcode"">" & Copient.PhraseLib.Lookup("term.voidbarcode", LanguageID) & ":</label>")
        Send("  </td>")
        Send("  <td>")
        Send("    <input type=""text"" id=""voidbarcode"" name=""voidbarcode"" value=""" & VoidBarcode & """ size=""30"" maxlength=""255""" & DisabledAttribute & " />")
        Send("  </td>")
        Send("</tr>")
      ElseIf (eDiscountType = 4) Then
        Send_Amount_Type(AnyProduct, DiscountType, AmountTypeID, "1", TierLevels, IsMfgCoupon)
      End If
      
      If AmountTypeID = AmountType_t.AMT_TYPE_PERCENT_OFF And MyCommon.Fetch_CPE_SystemOption(141) = 1 Then
        Send("<tr>")
        Send("  <td>")
        Send("    <label for=""percentfixedrounding"">" & Copient.PhraseLib.Lookup("CPEoffer-rew-discount.PercentFixedRounding", LanguageID) & ":</label>")
        Send("  </td>")
        Send("  <td>")
        Send("    <input type=""text"" class=""shorter"" id=""percentfixedrounding"" name=""percentfixedrounding"" maxlength=""4"" value=""" & Format(PercentFixedRounding, "0.00") & """ />")
        Send("  </td>")
        Send("</tr>")
      End If
      
      For t = 1 To TierLevels
              
        If Request.QueryString("tier" & t & "_l1discountamt") <> "" Then
          DiscountAmount = MyCommon.Extract_Val(Request.QueryString("tier" & t & "_l1discountamt"))
        End If
              
        If TierLevels > 1 AndAlso AmountTypeID <> AmountType_t.AMT_TYPE_STORED_VALUE Then
          Send("<tr>")
          Send("  <td><h3>" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</h3></td>")
          Send("</tr>")
        End If
        If AmountTypeID <> AmountType_t.AMT_TYPE_STORED_VALUE Then
          If (eDiscountType = 1) Or (eDiscountType = 3) Or (eDiscountType = 4) Then
            Send_Amount_Detail(DiscountID, AmountTypeID, DiscountAmount, "1", t)
            If TierLevels = 1 Then
              Send_Amount_DetailLevels(AmountTypeID, DiscountAmount, L1Cap, False, "1", t)
            End If
            If AmountTypeID = AmountType_t.AMT_TYPE_PERCENT_OFF And Math.Round(L1Cap, 2) > 0 Then
              Send_Amount_Type(AnyProduct, DiscountType, L2AmountTypeID, "2", TierLevels, IsMfgCoupon)
              Send_Amount_Detail(DiscountID, L2AmountTypeID, L2DiscountAmt, "2", t)
              If TierLevels = 1 Then
                Send_Amount_DetailLevels(L2AmountTypeID, L2DiscountAmt, L2Cap, False, "2", t)
              End If
              If L2AmountTypeID = AmountType_t.AMT_TYPE_PERCENT_OFF And Math.Round(L2Cap, 2) > 0 Then
                Send_Amount_Type(AnyProduct, DiscountType, L3AmountTypeID, "3", TierLevels, IsMfgCoupon)
                Send_Amount_Detail(DiscountID, L3AmountTypeID, L3DiscountAmt, "3", t)
                If TierLevels = 1 Then
                  Send_Amount_DetailLevels(L3AmountTypeID, L3DiscountAmt, 0, True, "3", t)
                End If
              End If
            End If
          End If
        End If
              
        If (AmountTypeID = AmountType_t.AMT_TYPE_STORED_VALUE AndAlso t = 1) Then
          'Stored value discount
          MyCommon.QueryStr = "select SVProgramID, Name, Value from StoredValuePrograms with (NoLock) where SVTypeID in (2,4) and Deleted=0;"
          rst = MyCommon.LRT_Select
          Send("<tr style=""height:10px;"">")
          Send("  <td colspan=""2""></td>")
          Send("</tr>")
          Send("<tr>")
          Send("  <td><label for=""svprogramid"">" & Copient.PhraseLib.Lookup("term.program", LanguageID) & ":</label></td>")
          Send("  <td><select class=""mediumlong"" name=""svprogramid"" id=""svprogramid"" onchange=""submitForm();"">")
          For Each row In rst.Rows
            SelectedStr = IIf(MyCommon.NZ(row.Item("SVProgramID"), -1) = SVProgramID, " selected=""selected""", "")
            Send("    <option value=""" & MyCommon.NZ(row.Item("SVProgramID"), -1) & """" & SelectedStr & " >" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
          Next
          Send("    </select>")
          For Each row In rst.Rows
            If (MyCommon.NZ(row.Item("SVProgramID"), -1) = SVProgramID) Then
              Send(" ($" & row.Item("Value") & "/unit)")
            End If
          Next
          Send("  </td>")
          Send("</tr>")
          Send("<tr>")
          Send("  <td><label for=""dollarlimit"">" & Copient.PhraseLib.Lookup("term.dollarlimit", LanguageID) & ":</label></td>")
          Send("  <td>")
          Send("    <input type=""text"" id=""tier" & t & "_dollarlimit"" name=""tier" & t & "_dollarlimit"" value=""" & Format(DollarLimit, "#####0.00") & """ size=""8"" maxlength=""8"" title=""" & Copient.PhraseLib.Lookup("term.dollar-limit-msg", LanguageID) & """" & DisabledAttribute & " />")
          Send("    <input type=""hidden"" id=""tier" & t & "_weightlimit"" name=""tier" & t & "_weightlimit"" value=""0"" />")
          Send("    <input type=""hidden"" id=""tier" & t & "_itemlimit"" name=""tier" & t & "_itemlimit"" value=""0"" />")
          Send("  </td>")
		  Send("</tr>")
          
          '===================================================================...CLOUDSOL-1271
		  If Request.QueryString("tier" & t & "_rdesc") <> "" Then 
            RDesc = Request.QueryString("tier" & t & "_rdesc")
          End If

          Send("<tr>")
          Send("  <td><label for=""dollarlimit"">" & Copient.PhraseLib.Lookup("term.receipttext", LanguageID) & ":</label></td>")
          Send("  <td>")
		  Send("    <input type=""text"" class=""medium"" id=""tier" & t & "_rdesc"" name=""tier" & t & "_rdesc"" maxlength=""18"" value=""" & RDesc & """" & DisabledAttribute & " />")
          Send("  </td>")		  
		  Send("</tr>")
          '===================================================================...CLOUDSOL-1271
		  
        ElseIf AmountTypeID = AmountType_t.AMT_TYPE_SPECIAL_PRICING Then
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
          If (infoMessage <> "") OrElse (Request.QueryString("mode") = "savediscount") OrElse (DiscountID = 0) Then
            SPLevels = MyCommon.Extract_Val(Request.QueryString("tier" & t & "_levels"))
            SPHighestLevel = MyCommon.Extract_Val(Request.QueryString("tier" & t & "_highestlevel"))
            For i = 1 To SPHighestLevel
              If Request.QueryString("tier" & t & "_level" & i) <> Nothing Then
                ValueString = Request.QueryString("tier" & t & "_level" & i)
                LevelID = i
                Send("<tr id=""tier" & t & "_level" & LevelID & "row"">")
                Send("  <td>")
                Send("    <input type=""radio"" id=""tier" & t & "_sprepeatlevel" & LevelID & """ name=""tier" & t & "_sprepeatlevel"" value=""" & LevelID & """" & IIf(LevelID = SPRepeatLevel, " checked=""checked""", "") & DisabledAttribute & " />")
                Send("  </td>")
                Send("  <td>")
                Send("    <input type=""text"" class=""short"" id=""tier" & t & "_level" & LevelID & """ name=""tier" & t & "_level" & LevelID & """ maxlength=""8"" value=""" & ValueString & """" & DisabledAttribute & " />")
                Send("    <input type=""button"" class=""ex"" name=""tier" & t & "_deletelevel" & LevelID & """ id=""tier" & t & "_deletelevel" & LevelID & """ value=""X"" onclick=""javascript:deleteLevel(" & t & ", " & LevelID & ");""" & DisabledAttribute & " />")
                Send("  </td>")
                Send("</tr>")
              End If
            Next
          Else
            MyCommon.QueryStr = "select PKID as DiscountTierID from CPE_DiscountTiers with (NoLock) " & _
                    "where DiscountID=" & DiscountID & ";"
            rst = MyCommon.LRT_Select
            If rst.Rows.Count > 0 Then
              DiscountTierID = MyCommon.NZ(rst.Rows(0).Item("DiscountTierID"), 0)
            End If
            MyCommon.QueryStr = "select Value, LevelID from CPE_SpecialPricing as SP with (NoLock) " & _
              "where SP.DiscountTierID=" & DiscountTierID & ";"
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count > 0) Then
              SPLevels = rst.Rows.Count
              SPHighestLevel = rst.Rows.Count
              For i = 0 To (rst.Rows.Count - 1)
                Value = Math.Round(MyCommon.NZ(rst.Rows(i).Item("Value"), 0), 2)
                LevelID = MyCommon.NZ(rst.Rows(i).Item("LevelID"), 0)
                Send("<tr id=""tier" & t & "_level" & LevelID & "row"">")
                Send("  <td>")
                Send("    <input type=""radio"" id=""tier" & t & "_sprepeatlevel" & LevelID & """ name=""tier" & t & "_sprepeatlevel"" value=""" & LevelID & """" & IIf(LevelID = SPRepeatLevel, " checked=""checked""", "") & DisabledAttribute & " />")
                Send("  </td>")
                Send("  <td>")
                Send("    <input type=""text"" class=""short"" id=""tier" & t & "_level" & LevelID & """ name=""tier" & t & "_level" & LevelID & """ maxlength=""8"" value=""" & Value & """" & DisabledAttribute & " />")
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
            Send("  <td>")
            Send("    <input type=""text"" class=""short"" id=""tier" & t & "_level" & (SPLevels + i) & """ name=""tier" & t & "_level" & (SPLevels + i) & """ maxlength=""8"" value=""""" & DisabledAttribute & " disabled=""disabled"" />")
            Send("    <input type=""button"" class=""ex"" name=""tier" & t & "_deletelevel" & (SPLevels + i) & """ id=""tier" & t & "_deletelevel" & (SPLevels + i) & """ value=""X"" onclick=""javascript:deleteLevel(" & t & ", " & (SPLevels + i) & ");"" disabled=""disabled"" />")
            Send("  </td>")
            Send("</tr>")
          Next
          'New level line
          Send("<tr id=""tier" & t & "_newlevelrow"">")
          Send("  <td>")
          Send("  </td>")
          Send("  <td>")
          Send("    <input type=""text"" class=""short"" id=""tier" & t & "_newlevel"" name=""tier" & t & "_newlevel"" maxlength=""8"" value="""" " & DisabledAttribute & " />")
          Send("    <input type=""button"" id=""tier" & t & "_addlevel"" name=""tier" & t & "_addlevel"" value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """  onclick=""javascript:addLevel(" & t & ");"" " & DisabledAttribute & " />")
          Send("    <input type=""hidden"" id=""tier" & t & "_levels"" name=""tier" & t & "_levels"" value=""" & SPLevels & """ />")
          Send("    <input type=""hidden"" id=""tier" & t & "_highestlevel"" name=""tier" & t & "_highestlevel"" value=""" & SPHighestLevel & """ />")
          Send("   </td>")
          Send("</tr>")
          'Item limit line
          Send("<tr id=""tier" & t & "_itemlimitrow"">")
          Send("  <td>")
          Send("  </td>")
          Send("  <td>")
          Send("    <label for=""tier" & t & "_itemlimit""><small><b>" & Copient.PhraseLib.Lookup("term.itemlimit", LanguageID) & ":</b></small></label><br />")
          Send("    <input type=""text"" id=""tier" & t & "_itemlimit"" name=""tier" & t & "_itemlimit"" value=""" & ItemLimit & """ size=""4"" maxlength=""4"" title=""" & Copient.PhraseLib.Lookup("term.itemlimitmsg", LanguageID) & """" & DisabledAttribute & " />")
          Send("  </td>")
          Send("</tr>")
          'Receipt text line
          Send("<tr id=""tier" & t & "_receiptrow"">")
          Send("  <td>")
          Send("  </td>")
          'Send("  <td>")
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
          MLI.CSSStyle = "width:204px;"
          MLI.Disabled = IIf(DisabledAttribute <> "", True, False)
          If MyCommon.Fetch_SystemOption(135) = 1 And MyCommon.Fetch_SystemOption(124) = 0 Then
            Send("  <td>")
            Send("  <table style=""width:100%"" cellspacing=""0"" cellpadding=""0""><tr><td style=""width:90%"">")
            Send("    <label for=""tier" & t & "_rdesc""><small><b>" & Copient.PhraseLib.Lookup("term.receipttext", LanguageID) & ":</b></small></label><br />")
            Send(Localization.SendTranslationInputs(MyCommon, MLI, Request.QueryString, 2))
            Send("  </td><td style=""width:10%"" valign=""bottom"">")
            Dim rstMsgDetail As DataTable
            MyCommon.QueryStr = "Select ReceiptTextID, ReceiptTextMsg from CPE_ReceiptTextMessages where isnull(BaseReceiptTextID,0)=0 order by ReceiptTextID"
            rstMsgDetail = MyCommon.LRT_Select
            If rstMsgDetail.Rows.Count > 0 Then
              Send("<input type=""button"" class=""regular"" name=""tier" & t & """ id=""tier" & t & """ value=""..."" title=""Select pre-defined Receipt text messages"" style=""FONT-SIZE: 8pt; WIDTH: 20px; HEIGHT: 24px"" onclick=""javascript:assigntextboxcontrol(this, 'foldercreate', true);"" />")
            Else
              Send("<input type=""button"" class=""regular"" name=""tier" & t & """ id=""tier" & t & """ value=""..."" title=""Select pre-defined Receipt text messages"" style=""FONT-SIZE: 8pt; WIDTH: 20px; HEIGHT: 24px"" disabled=""disabled"" onclick=""javascript:assigntextboxcontrol(this, 'foldercreate', true);"" />")
            End If
            Send("  </td></tr></table></td>")
          Else
            If MyCommon.Fetch_SystemOption(135) = 1 Then
              MLI.SupportDropdownforBaseLanguage = True
              MLI.PopulateDropdownQueryStr = "Select '' 'ReceiptTextMsg',-1 as 'ReceiptTextID' Union Select ReceiptTextMsg,ReceiptTextID from CPE_ReceiptTextMessages where isnull(BaseReceiptTextID,0)=0 order by ReceiptTextID"
              MLI.SupportDropdownCSSStyle = "width:180px;"
            End If
            Send("  <td>")
            Send("    <label for=""tier" & t & "_rdesc""><small><b>" & Copient.PhraseLib.Lookup("term.receipttext", LanguageID) & ":</b></small></label><br />")
            Send(Localization.SendTranslationInputs(MyCommon, MLI, Request.QueryString, 2))
            Send("  </td>")
          End If
          Send("</tr>")
                
        ElseIf AmountTypeID <> AmountType_t.AMT_TYPE_STORED_VALUE AndAlso AmountTypeID <> AmountType_t.AMT_TYPE_SPECIAL_PRICING Then
          'All other kinds of discouts
          If Not (AnyProduct) AndAlso DeptLevel = 0 Then
            Send_LimitsAndText(MyCommon, DiscountID, AmountTypeID, eDiscountType, "1", t, DisabledAttribute, DiscountType)
          Else
            Send_Text(MyCommon, DiscountID, AmountTypeID, eDiscountType, "1", t, DisabledAttribute)
          End If
          
        End If
        
      Next
            
      Send("  </table>")
      Send("  <hr class=""hidden"" />")
      Send("</div>")
    %>
        <div class="box" id="scorecards" style="z-index: 50;">
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.scorecard", LanguageID))%>
                </span>
            </h2>
            <div id="scorecardinputs">
                <%
                    If ComputeDiscount = 1 Then
                        Send("<table summary=""" & Copient.PhraseLib.Lookup("term.scorecard", LanguageID) & """>")
                        Send("  <tr>")
                        Send("    <td style=""width:70px;"">")
                        Send("      <label for=""ScorecardID"">" & Copient.PhraseLib.Lookup("CPEoffer-rew-discount.IncludeOnScorecard", LanguageID) & ":</label>")
                        Send("    </td>")
                        Send("    <td>")
                        MyCommon.QueryStr = "select ScorecardID, Description, EngineID, DefaultForEngine from Scorecards " & _
                                            "where ScorecardTypeID=3 and Deleted=0 and EngineID=" & EngineID & ";"
                        rst = MyCommon.LRT_Select
                        Send("      <select class=""medium"" id=""ScorecardID"" name=""ScorecardID"" onchange=""toggleScorecardText();"" style=""width:225px;"">")
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
                                ElseIf (MyCommon.NZ(row.Item("DefaultForEngine"), False) = True) AndAlso (MyCommon.NZ(row.Item("EngineID"), -1) = EngineID) Then
                                    Send("        <option value=""" & MyCommon.NZ(row.Item("ScorecardID"), 0) & """ selected=""selected"">" & MyCommon.NZ(row.Item("Description"), "") & "</option>")
                                Else
                                    Send("        <option value=""" & MyCommon.NZ(row.Item("ScorecardID"), 0) & """>" & MyCommon.NZ(row.Item("Description"), "") & "</option>")
                                End If
                            Next
                        End If
                        Send("      </select>")
                        Send("    </td>")
                        Send("  </tr>")
                        Send("  <tr id=""ScorecardDescLine"" " & IIf(ScorecardID = 0 AndAlso DiscountID <> 0, " style=""display:none;""", "") & " >")
                        Send("    <td>")
                        Send("      <label for=""ScorecardDesc"">" & Copient.PhraseLib.Lookup("term.scorecardtext", LanguageID) & ":</label>")
                        Send("    </td>")
                        Send("    <td>")
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
                        MLI.CSSStyle = "width:204px;"
                        MLI.Disabled = False
                        Send(Localization.SendTranslationInputs(MyCommon, MLI, Request.QueryString, 2))
                        Send("    </td>")
                        Send("  </tr>")
                        Send("</table>")
                    End If
                %>
            </div>
            <hr class="hidden" />
        </div>
        <%
            ' Show the options box only if there are available options.
            If (AmountTypeID = AmountType_t.AMT_TYPE_FIXED_AMOUNT_OFF_WV OrElse DeptLevel = 0 OrElse DeptLevel = 1 OrElse ComputeDiscount = 1) Then
                Send("<div class=""box"" id=""options"">")
                Send("  <h2>")
                Send("    <span>")
                Send("      " & Copient.PhraseLib.Lookup("term.advancedoptions", LanguageID))
                Send("    </span>")
                Send("  </h2>")
                Send("  <table summary=""" & Copient.PhraseLib.Lookup("term.options", LanguageID) & """>")
                If (AmountTypeID = AmountType_t.AMT_TYPE_FIXED_AMOUNT_OFF_WV) Then
                    Send("  <tr>")
                    Sendb("    <td style=""width:20px;""><input type=""checkbox"" id=""computediscount"" name=""computediscount"" value=""1""")
                    If ComputeDiscount = 1 Then Sendb(" checked=""checked""")
                    Send(DisabledAttribute & " /></td>")
                    Send("    <td nowrap><label for=""computediscount"">" & Copient.PhraseLib.Lookup("term.computediscount", LanguageID) & "</label></td>")
                    Send("  </tr>")
                Else
                    Send("  <tr>")
                    Send("    <td><input type=""hidden"" id=""computediscount"" name=""computediscount"" value=""1"" /></td>")
                    Send("  </tr>")
                End If
                If DeptLevel <> 0 Or DiscountType = 4 Then
                    Send("  <tr>")
                    Send("    <td><input type=""hidden"" id=""bestdeal"" name=""bestdeal"" value=""0"" /></td>")
                    Send("  </tr>")
                Else
                    Send("  <tr>")
                    If IsMfgCoupon Then
                        Send("    <td style=""width:20px;""><input type=""checkbox"" id=""bestdeal"" name=""bestdeal"" value=""0"" disabled=""disabled"" ")
                    Else
                        Send("    <td style=""width:20px;""><input type=""checkbox"" id=""bestdeal"" name=""bestdeal"" value=""1""")
                    End If
                    If BestDeal = 1 OrElse (DiscountID = 0 AndAlso BestDealDefaulted AndAlso Not IsMfgCoupon) Then Sendb(" checked=""checked""")
                    Send(DisabledAttribute & " /></td>")
                    Send("    <td nowrap><label for=""bestdeal"">" & Copient.PhraseLib.Lookup("term.bestdealitem", LanguageID) & "</label></td>")
                    Send("  </tr>")
                End If
        
                If (ComputeDiscount = 1) Then
                    Send("  <tr>")
                    Sendb("    <td style=""width:20px;""><input type=""checkbox"" id=""allownegative"" name=""allownegative"" value=""1""")
          
                    If AllowNegative = 1 Then Sendb(" checked=""checked""")
        
                    Send(DisabledAttribute & " onclick=""toggleflexneg()"" /></td>")
                    Sendb("    <td nowrap><label for=""allownegative"">")
                    If AnyProduct Then
                        Sendb(Copient.PhraseLib.Lookup("reward.discount-negativebasket", LanguageID))
                    ElseIf DeptLevel > 0 Then
                        Sendb(Copient.PhraseLib.Lookup("reward.discount-negativedepartment", LanguageID))
                    Else
                        Sendb(Copient.PhraseLib.Lookup("reward.discount-negativeitem", LanguageID))
                    End If
                    Send("</label></td>")
                    Send("  </tr>")
                End If
                'Flex Negative checkbox
                Send("  <tr>")
                Sendb("    <td style=""width:20px;""><input type=""checkbox"" id=""flexnegative"" name=""flexnegative"" value=""true""")
                'If it is not a new discount then use the database Flex Negative. if it is then use the system option
      
                If FlexNeg Then Sendb(" checked=""checked""")
      
                Send(DisabledAttribute & " /></td>")
                Sendb("    <td nowrap><label for=""flexnegative"">")
                Sendb(Copient.PhraseLib.Lookup("reward.discount-flexnegative", LanguageID))
                Send("</label></td>")
                Send("  </tr>")
                
                Send("  </table>")
                Send("  <hr class=""hidden"" />")
                Send("</div>")
            End If
        %>
    </div>
</div>
</form>
<script runat="server">

    Enum AmountType_t
        AMT_TYPE_FIXED_AMOUNT_OFF = 1
        AMT_TYPE_PRICE_POINT = 2
        AMT_TYPE_PERCENT_OFF = 3
        AMT_TYPE_FREE = 4
        AMT_TYPE_FIXED_AMOUNT_OFF_WV = 5
        AMT_TYPE_PRICE_POINT_WV = 6
        AMT_TYPE_STORED_VALUE = 7
        AMT_TYPE_SPECIAL_PRICING = 8
    End Enum

    Const DisabledString As String = "disabled = ""disabled"""

    Dim DisabledAttr As String = ""
        
    Sub SetDisabledAttr(ByVal str As String)
        DisabledAttr = str
    End Sub
        
    Sub Send_Amount_Type(ByVal AnyProduct As Boolean, ByVal DiscountType As Integer, ByVal AmountTypeID As Integer, _
                         ByVal Level As String, ByVal TierLevels As Integer, ByVal IsMfgCoupon As Boolean)
        Dim MyCommon As New Copient.CommonInc
        Dim rst As DataTable
        Dim row As DataRow
        Dim UOMCriteria As String = ""
        Const UOM_ALWAYS As Integer = -1
        Const UOM_DISABLED_ONLY As Integer = 0
    
        MyCommon.Open_LogixRT()
    
        Send("<tr>")
        Sendb("  <td style=""width:70px;""><label for=""l" & Level & "amounttypeid"">" & Copient.PhraseLib.Lookup("term.type", LanguageID))
        If Level > 1 Then
            Sendb("&nbsp;" & Level)
        End If
        Send(":</label></td>")
        Send("  <td>")
        Send("    <select id=""l" & Level & "amounttypeid"" name=""l" & Level & "amounttypeid"" onchange=""submitForm();""" & DisabledAttr & " style=""width:225px;"">")
    
        ' filter-out those amount types that are exclusive to UE dealing with UOM
        UOMCriteria = "(MultiUOMState = " & UOM_ALWAYS & " or MultiUOMState = " & UOM_DISABLED_ONLY & ")"

        ' If AnyProduct Or DeptLevel > 0 Then
        If AnyProduct And DiscountType = 3 Then
            MyCommon.QueryStr = "select AmountTypeID, PhraseID from CPE_AmountTypes AT with (NoLock) where " & UOMCriteria & " and AmountTypeID IN (" & AmountType_t.AMT_TYPE_FIXED_AMOUNT_OFF & ", " & AmountType_t.AMT_TYPE_STORED_VALUE & IIf(IsMfgCoupon, "", "," & AmountType_t.AMT_TYPE_PERCENT_OFF) & ");"
        ElseIf DiscountType = 2 Then
            MyCommon.QueryStr = "select AmountTypeID, PhraseID from CPE_AmountTypes AT with (NoLock) where " & UOMCriteria & " and AmountTypeID IN (" & AmountType_t.AMT_TYPE_FIXED_AMOUNT_OFF & IIf(IsMfgCoupon, "", "," & AmountType_t.AMT_TYPE_PERCENT_OFF) & ");"
        ElseIf DiscountType = 4 Then
            MyCommon.QueryStr = "select AmountTypeID, PhraseID from CPE_AmountTypes AT with (NoLock) where AmountTypeID IN (" & AmountType_t.AMT_TYPE_PRICE_POINT & ");"
        ElseIf DiscountType = 4 Then
            MyCommon.QueryStr = "select AmountTypeID, PhraseID from CPE_AmountTypes AT with (NoLock) where AmountTypeID IN (" & AmountType_t.AMT_TYPE_PRICE_POINT & ");"
        Else
            MyCommon.QueryStr = "select AmountTypeID, PhraseID from CPE_AmountTypes AT with (NoLock) where " & UOMCriteria & IIf(TierLevels > 1, " and AmountTypeID <> " & AmountType_t.AMT_TYPE_STORED_VALUE, "") & ";"
        End If
    
        rst = MyCommon.LRT_Select
        If Not (rst.Rows.Count = 0) Then
            If AmountTypeID = 0 Then
                AmountTypeID = AmountType_t.AMT_TYPE_FIXED_AMOUNT_OFF
            End If
            For Each row In rst.Rows
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
            Next
        End If
        Send("    </select>")
        Send("  </td>")
        Send("</tr>")
    
        MyCommon.Close_LogixRT()
    End Sub
        
    Sub Send_Amount_Detail(ByVal DiscountID As Long, ByVal AmountTypeID As Integer, ByVal DiscountAmount As Object, ByVal Level As String, ByVal TierLevel As Integer)
        Dim MyCommon As New Copient.CommonInc
        Dim dt As DataTable
          
        MyCommon.Open_LogixRT()
          
        If TierLevel = 1 AndAlso Integer.Parse(Level) > 1 Then
            MyCommon.QueryStr = "select L" & Level & "DiscountAmt as DiscountAmount from CPE_Discounts " & _
                                "where DiscountID=" & DiscountID & ";"
        Else
            MyCommon.QueryStr = "select DiscountAmount from CPE_DiscountTiers " & _
                                "where DiscountID=" & DiscountID & " and TierLevel=" & TierLevel & ";"
        End If
        dt = MyCommon.LRT_Select()
        If dt.Rows.Count > 0  And AmountTypeID > 0Then
            If (AmountTypeID = AmountType_t.AMT_TYPE_FREE) Then
                Send("<tr style=""display:none;"">")
                Send("  <td colspan=""2""><input type=""hidden"" id=""tier" & TierLevel & "_l" & Level & "discountamt"" name=""tier" & TierLevel & "_l" & Level & "discountamt"" value=""0"" /></td>")
                Send("</tr>")
            ElseIf (AmountTypeID = AmountType_t.AMT_TYPE_FIXED_AMOUNT_OFF Or AmountTypeID = AmountType_t.AMT_TYPE_FIXED_AMOUNT_OFF_WV) Then
                Send("<tr>")
                Send("  <td><label for=""tier" & TierLevel & "_l" & Level & "discountamt"">" & Copient.PhraseLib.Lookup("term.amount", LanguageID) & ":</label></td>")
                Send("  <td><input type=""text"" id=""tier" & TierLevel & "_l" & Level & "discountamt"" name=""tier" & TierLevel & "_l" & Level & "discountamt"" value=""" & Format(MyCommon.NZ(dt.Rows(0).Item("DiscountAmount"), 0), "#####0.00#") & """ size=""6"" maxlength=""6""" & DisabledAttr & " /></td>")
                Send("</tr>")
            ElseIf (AmountTypeID = AmountType_t.AMT_TYPE_PRICE_POINT Or AmountTypeID = AmountType_t.AMT_TYPE_PRICE_POINT_WV) Then
                Send("<tr>")
                Send("  <td><label for=""tier" & TierLevel & "_l" & Level & "discountamt"">" & Copient.PhraseLib.Lookup("term.saleprice", LanguageID) & ":</label></td>")
                Send("  <td><input type=""text"" id=""tier" & TierLevel & "_l" & Level & "discountamt"" name=""tier" & TierLevel & "_l" & Level & "discountamt"" value=""" & Math.Round(MyCommon.NZ(dt.Rows(0).Item("DiscountAmount"), 0), 2) & """ size=""6"" maxlength=""6""" & DisabledAttr & " /></td>")
                Send("</tr>")
            ElseIf (AmountTypeID = AmountType_t.AMT_TYPE_PERCENT_OFF) Then
                Send("<tr>")
                Send("  <td><label for=""tier" & TierLevel & "_l" & Level & "discountamt"">" & Copient.PhraseLib.Lookup("term.amount", LanguageID) & ":</label></td>")
                Send("  <td><input type=""text"" id=""tier" & TierLevel & "_l" & Level & "discountamt"" name=""tier" & TierLevel & "_l" & Level & "discountamt"" value=""" & Format(MyCommon.NZ(dt.Rows(0).Item("DiscountAmount"), 0), "#####0.00#") & """ size=""6"" maxlength=""6""" & DisabledAttr & " />%</td>")
                Send("</tr>")
            ElseIf (AmountTypeID = AmountType_t.AMT_TYPE_STORED_VALUE Or AmountTypeID = AmountType_t.AMT_TYPE_SPECIAL_PRICING) Then
                Send("<tr style=""display:none;"">")
                Send("  <td><label for=""tier" & TierLevel & "_l" & Level & "discountamt"">" & Copient.PhraseLib.Lookup("term.amount", LanguageID) & ":</label></td>")
                Send("  <td><input type=""hidden"" id=""tier" & TierLevel & "_l" & Level & "discountamt"" name=""tier" & TierLevel & "_l" & Level & "discountamt"" value=""0"" /></td>")
                Send("</tr>")
            End If
        Else
            If (AmountTypeID <> AmountType_t.AMT_TYPE_FREE And AmountTypeID <> AmountType_t.AMT_TYPE_SPECIAL_PRICING) Then
                If DiscountAmount > 0 Then
                    Send("<tr>")
                    Send("  <td><label for=""tier" & TierLevel & "_l" & Level & "discountamt"">" & Copient.PhraseLib.Lookup("term.amount", LanguageID) & ":</label></td>")
                    Send("  <td><input type=""text"" id=""tier" & TierLevel & "_l" & Level & "discountamt"" name=""tier" & TierLevel & "_l" & Level & "discountamt"" value=""" & DiscountAmount & """ size=""6"" maxlength=""6""" & DisabledAttr & " /></td>")
                    Send("</tr>")
                Else
                    ' It's a new discount, so show the discountamt field left blank
                    Send("<tr>")
                    Send("  <td><label for=""tier" & TierLevel & "_l" & Level & "discountamt"">" & Copient.PhraseLib.Lookup("term.amount", LanguageID) & ":</label></td>")
                    Send("  <td><input type=""text"" id=""tier" & TierLevel & "_l" & Level & "discountamt"" name=""tier" & TierLevel & "_l" & Level & "discountamt"" value="""" size=""6"" maxlength=""6""" & DisabledAttr & " /></td>")
                    Send("</tr>")
                End If
            End If
        End If
          
        MyCommon.Close_LogixRT()
    End Sub
        
    Sub Send_Amount_DetailLevels(ByVal AmountTypeID As Integer, ByVal DiscountAmount As Object, ByVal Cap As Object, ByVal NoCap As Boolean, ByVal Level As String, ByVal TierLevel As Integer)
        Dim MyCommon As New Copient.CommonInc
          
        MyCommon.Open_LogixRT()
          
        If AmountTypeID = AmountType_t.AMT_TYPE_PERCENT_OFF And Not (NoCap) Then
            Send("<tr>")
            Send("  <td><label for=""l" & Level & "cap"">" & Copient.PhraseLib.Lookup("term.upto", LanguageID) & ":</label></td>")
            Send("  <td><input type=""text"" id=""l" & Level & "cap"" name=""l" & Level & "cap"" value=""" & Format(Cap, "#####0.00#") & """ size=""6"" maxlength=""6""" & DisabledAttr & " onkeyup=""handleUpToEntry(this.value, " & Level & ");"" />")
            If (Cap = 0) Then
                Send("     <input type=""button"" id=""btnDown" & Level & """ name=""btnDown" & Level & """ class=""down"" value=""&#9660;"" onclick=""submitForm();""  disabled=""disabled"" />")
            End If
            Send("     <input type=""button"" id=""btnEx" & Level & """ name=""btnEx" & Level & """ class=""ex"" value=""x"" onclick=""zeroLevel('" & Level & "');""   />")
            Send("  </td>")
            Send("</tr>")
        End If
          
        MyCommon.Close_LogixRT()
    End Sub
        
    Sub Send_LimitsAndText(ByRef MyCommon As Copient.CommonInc, ByVal DiscountID As Long, ByVal AmountTypeID As Integer, ByVal eDiscountType As Integer, ByVal Level As String, ByVal TierLevel As Integer, ByVal DisabledAttribute As String, ByVal DiscountType As Integer)
        Dim ItemLimit As Long = 0
        Dim WeightLimit As Double = 0
        Dim IsWeightTotal As Boolean = True
        Dim DollarLimit As Double = 0
        Dim RDesc As String = ""
        Dim PKID As Integer = 0

        Dim Disabled As Boolean = DisabledAttribute <> ""

        MyCommon.QueryStr = "select PKID, ReceiptDescription, ItemLimit, WeightLimit, DollarLimit, IsWeightTotal from CPE_DiscountTiers " & _
                            "where DiscountID=" & DiscountID & " and TierLevel=" & TierLevel & ";"
        Dim dt As DataTable = MyCommon.LRT_Select()
        If dt.Rows.Count > 0 Then
            ' This is not a new discount, so use the stored values.
            ItemLimit = MyCommon.NZ(dt.Rows(0).Item("ItemLimit"), 0)
            WeightLimit =  Math.Round(MyCommon.NZ(dt.Rows(0).Item("WeightLimit"), 0),3)
            IsWeightTotal = MyCommon.NZ(dt.Rows(0).Item("IsWeightTotal"), 1)
            DollarLimit = MyCommon.NZ(dt.Rows(0).Item("DollarLimit"), 0)
            RDesc = MyCommon.NZ(dt.Rows(0).Item("ReceiptDescription"), "")
            PKID = MyCommon.NZ(dt.Rows(0).Item("PKID"), 0)
        Else
            ' This is a new discount, so use new or page values.
            If Request.QueryString("tier" & TierLevel & "_itemlimit") <> "" Then ItemLimit = MyCommon.Extract_Val(Request.QueryString("tier" & TierLevel & "_itemlimit"))
            If Request.QueryString("tier" & TierLevel & "_weightlimit") <> "" Then WeightLimit = MyCommon.Extract_Val(Request.QueryString("tier" & TierLevel & "_weightlimit"))
            If Request.QueryString("tier" & TierLevel & "_dollarlimit") <> "" Then DollarLimit = MyCommon.Extract_Val(Request.QueryString("tier" & TierLevel & "_dollarlimit"))
            If Request.QueryString("tier" & TierLevel & "_isweighttotal") <> "" Then IsWeightTotal = MyCommon.Extract_Val(Request.QueryString("tier" & TierLevel & "_isweighttotal"))
            RDesc = Request.QueryString("tier" & TierLevel & "_rdesc")
            PKID = 0
        End If

        Send_ITEM_and_WEIGHT_LIMIT_Fields2(TierLevel, AmountTypeID <> AmountType_t.AMT_TYPE_FIXED_AMOUNT_OFF_WV And AmountTypeID <> AmountType_t.AMT_TYPE_PRICE_POINT_WV, ItemLimit, WeightLimit, IsWeightTotal, Disabled, DiscountType)
        Send_DOLLARLIMIT_Field(TierLevel, DollarLimit, Disabled, DiscountType)
        If ((eDiscountType = 3) Or (eDiscountType = 4)) Then
            Send_RDESC_Field(MyCommon, PKID, TierLevel, RDesc, Disabled)
        End If
          
    End Sub

    ' Send_Text
    Sub Send_Text(ByRef MyCommon As Copient.CommonInc, ByVal DiscountID As Long, ByVal AmountTypeID As Integer, ByVal eDiscountType As Integer, ByVal Level As String, ByVal TierLevel As Integer, ByVal DisabledAttribute As String)
        Dim dt As DataTable
          
        MyCommon.QueryStr = "select PKID, ReceiptDescription, ItemLimit, WeightLimit, DollarLimit from CPE_DiscountTiers " & _
                            "where DiscountID=" & DiscountID & " and TierLevel=" & TierLevel & ";"
        dt = MyCommon.LRT_Select()
        If dt.Rows.Count > 0 Then
            If (eDiscountType = 3) Or (eDiscountType = 4) Then
                Send_RDESC_Field(MyCommon, dt.Rows(0).Item("PKID"), TierLevel, MyCommon.NZ(dt.Rows(0).Item("ReceiptDescription"), ""), DisabledAttribute <> "")
            End If
        Else
            Send_RDESC_Field(MyCommon, 0, TierLevel, "", DisabledAttribute <> "")
        End If
          
    End Sub

    '================================================================
    ' Field-sending functions. Define the code to send a field _once_.
    '================================================================

    ' Send_ITEM_and_WEIGHT_LIMIT_Fields2
    ' This replaced Send_ITEM_and_WEIGHT_LIMIT_Fields
    Sub Send_ITEM_and_WEIGHT_LIMIT_Fields2(ByVal TierLevel As Integer, ByVal ItemOnly As Boolean, ByVal ItemLimit As Integer, ByVal WeightLimit As Double, _
                                           ByVal IsTotal As Boolean, ByVal Disabled As Boolean, ByVal DiscountType As Integer)
        Dim DisabledAttribute As String = IIf(Disabled, DisabledString, "")
        Dim ItemName As String = """tier" & TierLevel & "_itemlimit"""
        Dim WeightName As String = """tier" & TierLevel & "_weightlimit"""
        Dim SwitchName As String = """tier" & TierLevel & "_isweighttotal"""
    
        Send("<tr>")
        If DiscountType = 4 Then
            Send("  <td><input type=""hidden"" id=" & ItemName & " name=" & ItemName & " value=""0"" /></td>")
        Else
            Send("  <td><label for=" & ItemName & ">" & Copient.PhraseLib.Lookup("term.itemlimit", LanguageID) & ":</label></td>")
            Send("  <td>")
            Send("    <input type=""text"" id=" & ItemName & " name=" & ItemName & " value=""" & ItemLimit & """ size=""8"" maxlength=""4"" title=""" & Copient.PhraseLib.Lookup("term.itemlimitmsg", LanguageID) & """" & DisabledAttribute & " />")
            If ItemOnly Then
                Send("    <input type=""hidden"" id=" & WeightName & " name=" & WeightName & " value=""0"" />")
            End If
            Send("  </td>")
        End If
        Send("</tr>")
        If Not ItemOnly Then
            Send("<tr>")
            Send("  <td><label for=" & WeightName & ">" & Copient.PhraseLib.Lookup("cpeoffer-rew-disc-wgtgallimit", LanguageID) & ":</label></td>")
            Send("  <td>")
            Send("    <input type=""text"" id=" & WeightName & " name=" & WeightName & " value=""" & WeightLimit & """ size=""8"" maxlength=""8"" title=""" & Copient.PhraseLib.Lookup("term.wgt-gal-limit-msg", LanguageID) & """" & DisabledAttribute & " />")
            Send("    <select id=" & SwitchName & " name=" & SwitchName & ">")
            Send("      <option " & IIf(IsTotal, "", "selected=""selected"" ") & "value=""0"">" & Copient.PhraseLib.Lookup("term.peritem", LanguageID).ToLower() & "</option>")
            Send("      <option " & IIf(IsTotal, "selected=""selected"" ", "") & "value=""1"">" & Copient.PhraseLib.Lookup("term.total", LanguageID).ToLower() & "</option>")
            Send("    </select>")
            Send("  </td>")
            Send("</tr>")
        End If
    End Sub

    ' Send_DOLLARLIMIT_Field
    Sub Send_DOLLARLIMIT_Field(ByVal TierLevel As Integer, ByVal DollarLimit As Double, ByVal Disabled As Boolean, ByVal DiscountType As Integer)
        Send("<tr>")
        If DiscountType = 4 Then
            Send("  <td><input type=""hidden"" id=""tier" & TierLevel & "_dollarlimit"" name=""tier" & TierLevel & "_dollarlimit"" value=""0"" /></td>")
        Else
            Send("  <td><label for=""tier" & TierLevel & "_dollarlimit"">" & Copient.PhraseLib.Lookup("term.dollarlimit", LanguageID) & ":</label></td>")
            Send("  <td>")
            Send("    <input type=""text"" id=""tier" & TierLevel & "_dollarlimit"" name=""tier" & TierLevel & "_dollarlimit"" value=""" & Format(DollarLimit, "#####0.00") & """ size=""8"" maxlength=""8"" title=""" & Copient.PhraseLib.Lookup("term.dollarlimit", LanguageID) & """" & IIf(Disabled, DisabledString, "") & " />")
            Send("  </td>")
        End If
        Send("</tr>")
    End Sub

    ' Send_RDESC_Field
    Sub Send_RDESC_Field(ByRef MyCommon As Copient.CommonInc, ByVal DiscountTiersID As Integer, ByVal TierLevel As Integer, ByVal Dirty_Value As String, ByVal Disabled As Boolean)
        Dim Clean_Value As String = HttpUtility.HtmlEncode(Dirty_Value)
        Send("<tr>")
        Send("  <td><label for=""tier" & TierLevel & "_rdesc"">" & Copient.PhraseLib.Lookup("term.receipttext", LanguageID) & ":</label></td>")
        'Send("    <input type=""text"" class=""medium"" id=""tier" & TierLevel & "_rdesc"" name=""tier" & TierLevel & "_rdesc"" maxlength=""18"" value=""" & Clean_Value & """" & IIf(Disabled, DisabledString, "") & " /><br />")
        Dim Localization As Copient.Localization
        Localization = New Copient.Localization(MyCommon)
        Dim MLI As New Copient.Localization.MultiLanguageRec
        MLI.ItemID = DiscountTiersID
        MLI.MLTableName = "CPE_DiscountTiersTranslations"
        MLI.MLIdentifierName = "DiscountTiersID"
        MLI.StandardTableName = "CPE_DiscountTiers"
        MLI.StandardIdentifierName = "PKID"
        MLI.MLColumnName = "ReceiptDesc"
        MLI.StandardValue = Clean_Value
        MLI.InputName = "tier" & TierLevel & "_rdesc"
        MLI.InputID = "tier" & TierLevel & "_rdesc"
        MLI.InputType = "text"
        MLI.LabelPhrase = ""
        MLI.MaxLength = 18
        MLI.CSSClass = ""
        MLI.CSSStyle = "width:204px;"
        MLI.Disabled = Disabled
        If MyCommon.Fetch_SystemOption(135) = 1 And MyCommon.Fetch_SystemOption(124) = 0 Then
            Send("  <td>")
            Send("  <table style=""width:100%"" cellspacing=""0"" cellpadding=""0""><tr><td style=""width:90%"">")
            Send(Localization.SendTranslationInputs(MyCommon, MLI, Request.QueryString, 2))
            Send("  </td><td style=""width:10%"">")
            Dim rstMsgDetail As DataTable
            MyCommon.QueryStr = "Select ReceiptTextID, ReceiptTextMsg from CPE_ReceiptTextMessages where isnull(BaseReceiptTextID,0)=0 order by ReceiptTextID"
            rstMsgDetail = MyCommon.LRT_Select
            If rstMsgDetail.Rows.Count > 0 Then
                Send("<input type=""button"" class=""regular"" name=""tier" & TierLevel & """ id=""tier" & TierLevel & """  value=""..."" title=""Select pre-defined Receipt text messages""  style=""FONT-SIZE: 8pt; WIDTH: 20px; HEIGHT: 24px"" onclick=""javascript:assigntextboxcontrol(this, 'foldercreate', true);"" />")
            Else
                Send("<input type=""button"" class=""regular"" name=""tier" & TierLevel & """ id=""tier" & TierLevel & """  value=""..."" title=""Select pre-defined Receipt text messages""  style=""FONT-SIZE: 8pt; WIDTH: 20px; HEIGHT: 24px"" disabled=""disabled"" onclick=""javascript:assigntextboxcontrol(this, 'foldercreate', true);"" />")
            End If
            Send("  </td></tr></table></td>")
        Else
            If MyCommon.Fetch_SystemOption(135) = 1 Then
                MLI.SupportDropdownforBaseLanguage = True
                MLI.PopulateDropdownQueryStr = "Select '' 'ReceiptTextMsg',-1 as 'ReceiptTextID' Union Select ReceiptTextMsg,ReceiptTextID from CPE_ReceiptTextMessages where isnull(BaseReceiptTextID,0)=0 order by ReceiptTextID"
                MLI.SupportDropdownCSSStyle = "width:180px;"
            End If
            Send("  <td>")
            Send(Localization.SendTranslationInputs(MyCommon, MLI, Request.QueryString, 2))
            Send("  </td>")
        End If
        Send("</tr>")
    End Sub

  
    '================================================================
    ' End field-sending functions. Define the code to send a field 
    ' _once_.
    '================================================================


    Function Create_Discount(ByVal OfferID As String, ByVal TpROID As Long, ByVal Phase As Long, ByRef DeliverableID As Long) As Long
        Dim MyCommon As New Copient.CommonInc
        Dim DiscountID As Long = 0
          
        Try
            MyCommon.QueryStr = "dbo.pa_CPE_AddDiscount"
            MyCommon.Open_LogixRT()
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt, 4).Value = OfferID
            MyCommon.LRTsp.Parameters.Add("@TpROID", SqlDbType.Int, 4).Value = TpROID
            MyCommon.LRTsp.Parameters.Add("@Phase", SqlDbType.Int, 4).Value = Phase
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

    ' Save Discount
    Sub Save_Discount(ByVal OfferID As Long, ByVal DiscountID As Long, ByVal TierLevels As Long, ByVal WriteTier As Boolean, ByVal JustCreated As Boolean)
        Dim MyCommon As New Copient.CommonInc
        Dim Logix As New Copient.LogixInc
        Dim Localization As Copient.Localization
        Dim DiscountTypeID As Integer
        Dim eDiscountID As Long
        Dim NumRecs As String
        Dim rst As DataTable
        Dim rst2 As DataTable
        Dim row As DataRow
        Dim i As Integer = 0
        Dim AnyProduct As Boolean
        Dim DiscountedProductGroupID As String
        Dim ExcludedProductGroupID As String
        Dim AmountTypeID As Long
        Dim DecliningBalance As String
        Dim RetroactiveDiscount As String
        Dim ChargebackDeptID As Long
        Dim BestDeal As Integer
        Dim AllowNegative As Integer
        Dim ComputeDiscount As Integer
        Dim DeptLevel As Integer
        Dim ItemError As Boolean
        Dim L1Cap As Object
        Dim L2DiscountAmt As Object
        Dim L2AmountTypeID As Long
        Dim L2Cap As Object
        Dim L3DiscountAmt As Object
        Dim L3AmountTypeID As Long
        Dim DiscountBarcode As String
        Dim VoidBarcode As String
        Dim UserGroupID As String
        Dim AdminUserID As Integer
        Dim SVProgramID As Integer
        Dim FlexNeg As Boolean = False
        Dim DiscountAmount As Object
        Dim ItemLimit As String
        Dim WeightLimit As Object
        Dim IsWeightTotal As Boolean
        Dim DollarLimit As Object
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
        Dim PercentFixedRounding As Decimal = 0

        Dim TierPKID As Integer = 0
  
        Dim MLI As New Copient.Localization.MultiLanguageRec
    
        MyCommon.Open_LogixRT()
        AdminUserID = Verify_AdminUser(MyCommon, Logix)
        Localization = New Copient.Localization(MyCommon)
    
        ItemError = False
    
        ' Non-tiered discount data:
        DiscountBarcode = MyCommon.Strip_Quotes(Trim(Request.QueryString("discountbarcode")))
        VoidBarcode = MyCommon.Strip_Quotes(Trim(Request.QueryString("voidbarcode")))
        DiscountedProductGroupID = MyCommon.Extract_Val(Request.QueryString("discountedpgid"))
        ExcludedProductGroupID = MyCommon.Extract_Val(Request.QueryString("excludedpgid"))
        ChargebackDeptID = MyCommon.Extract_Val(Request.QueryString("chargeback"))
        UserGroupID = MyCommon.Extract_Val(Request.QueryString("usergroupid"))
        BestDeal = MyCommon.Extract_Val(Request.QueryString("bestdeal"))
        'If JustCreated Then
        'AllowNegative = IIf((MyCommon.Fetch_CPE_SystemOption(18) = "1"), 1, 0)
        'Else
        AllowNegative = MyCommon.Extract_Val(Request.QueryString("hdnallownegative"))
        'End If
        ComputeDiscount = MyCommon.Extract_Val(Request.QueryString("computediscount"))
        DiscountTypeID = MyCommon.Extract_Val(Request.QueryString("discountType"))
        SVProgramID = MyCommon.Extract_Val(Request.QueryString("svprogramid"))
        If Request.QueryString("decliningbalance") = "true" Then
            DecliningBalance = "1"
        Else
            DecliningBalance = "0"
        End If
        RetroactiveDiscount = "1"
        If Request.QueryString("hdnflexnegative") = "1" Then
            FlexNeg = True
        End If
        AmountTypeID = MyCommon.Extract_Val(Request.QueryString("l1amounttypeid"))
        If AmountTypeID <> AmountType_t.AMT_TYPE_FIXED_AMOUNT_OFF_WV Then ComputeDiscount = 1
        If ComputeDiscount = 0 Then
            AllowNegative = 1
        End If
        L1Cap = MyCommon.Extract_Val(Request.QueryString("l1cap"))
        L2DiscountAmt = MyCommon.Extract_Val(Request.QueryString("tier1_l2discountamt"))
        L2AmountTypeID = MyCommon.Extract_Val(Request.QueryString("l2amounttypeid"))
        L2Cap = MyCommon.Extract_Val(Request.QueryString("l2cap"))
        L3DiscountAmt = MyCommon.Extract_Val(Request.QueryString("tier1_l3discountamt"))
        L3AmountTypeID = MyCommon.Extract_Val(Request.QueryString("l3amounttypeid"))
        ScorecardID = MyCommon.Extract_Val(Request.QueryString("ScorecardID"))
        ScorecardDesc = IIf(Request.QueryString("ScorecardDesc") <> "", Request.QueryString("ScorecardDesc"), "")
        PercentFixedRounding = MyCommon.Extract_Val(Request.QueryString("percentfixedrounding"))
          
        ' Tiered discount data; for now, let's just set these to the first-tier values:
        RDesc = MyCommon.Strip_Quotes(Trim(Request.QueryString("tier1_rdesc")))
        ItemLimit = MyCommon.Extract_Val(Request.QueryString("tier1_itemlimit"))
        WeightLimit = MyCommon.Extract_Val(Request.QueryString("tier1_weightlimit"))
        IsWeightTotal = IIf(Request.QueryString("tier1_isweighttotal") <> "0", True, False)
        DollarLimit = MyCommon.Extract_Val(Request.QueryString("tier1_dollarlimit"))
        DiscountAmount = MyCommon.Extract_Val(Request.QueryString("tier1_l1discountamt"))

        If AmountTypeID = AmountType_t.AMT_TYPE_FREE Then DiscountAmount = 0
        If L2AmountTypeID = AmountType_t.AMT_TYPE_FREE Then L2DiscountAmt = 0
        If L3AmountTypeID = AmountType_t.AMT_TYPE_FREE Then L3DiscountAmt = 0

        ' If the discount is percent off, set a default for the L2AmountTypeID
        If ((AmountTypeID = AmountType_t.AMT_TYPE_PERCENT_OFF) And (DiscountAmount > 0)) Then
            If L2AmountTypeID = 0 Then L2AmountTypeID = AmountType_t.AMT_TYPE_FIXED_AMOUNT_OFF
        End If
        If ((L2AmountTypeID = AmountType_t.AMT_TYPE_PERCENT_OFF) And (L2DiscountAmt > 0)) Then
            If L3AmountTypeID = 0 Then L3AmountTypeID = AmountType_t.AMT_TYPE_FIXED_AMOUNT_OFF
        End If

        ' See if the discounted product group is the AnyProduct group
        AnyProduct = False
        MyCommon.QueryStr = "select isnull(AnyProduct, 0) as AnyProduct from ProductGroups with (NoLock) where ProductGroupID=" & DiscountedProductGroupID & ";"
        rst = MyCommon.LRT_Select
        If Not (rst.Rows.Count = 0) Then
            AnyProduct = MyCommon.NZ(rst.Rows(0).Item("AnyProduct"), False)
            DeptLevel = IIf(DiscountTypeID = 2, 1, 0)
        End If
        If AnyProduct Then
            ' Stored value (AmountTypeID 7) should not set the item limit to 1 because the local server expects a 0.
            If (AmountTypeID = AmountType_t.AMT_TYPE_STORED_VALUE) Then
                ItemLimit = 0
            Else
                ItemLimit = 1
                DollarLimit = 0
            End If
            RetroactiveDiscount = "0"
            DecliningBalance = "0"
            If (ChargebackDeptID = 0) Then ChargebackDeptID = 14
        ElseIf DeptLevel > 0 Then
            ExcludedProductGroupID = 0
            If (ChargebackDeptID = 0 Or ChargebackDeptID = 14) Then ChargebackDeptID = 10
        Else
            ExcludedProductGroupID = 0
            If (ChargebackDeptID = 10) Then ChargebackDeptID = "0"
        End If
          
        If (DiscountID > 0) Then
            eDiscountID = DiscountID
        Else
            eDiscountID = MyCommon.Extract_Val(Request.QueryString("DiscountID"))
        End If
          
        If DiscountedProductGroupID = 0 Then
            If DiscountTypeID <> 4 Then
                DiscountedProductGroupID = "Null"
            End If
            ExcludedProductGroupID = "Null"
        End If
    
        ' Force the discount type to basket level if the discounted product grouo is AnyProduct
        If (AnyProduct AndAlso DiscountTypeID <> 3) Then
            DiscountedProductGroupID = "Null"
        ElseIf (AnyProduct) Then
            DiscountTypeID = 3
        End If
          
        If (AmountTypeID = AmountType_t.AMT_TYPE_SPECIAL_PRICING) Then
            SavePricePointLevels(OfferID, DiscountID, WriteTier)
        End If
          
        NumRecs = 0
          
        ScorecardDesc = Replace(ScorecardDesc, "'", "''")
          
        ' Update the CPE_Discounts table
        MyCommon.QueryStr = "update CPE_Discounts with (RowLock) set " & _
                            "DiscountTypeID='" & DiscountTypeID & "', " & _
                            "ReceiptDescription='" & RDesc & "', " & _
                            "DiscountBarcode='" & DiscountBarcode & "', " & _
                            "VoidBarcode='" & VoidBarcode & "', " & _
                            "DiscountedProductGroupID=" & DiscountedProductGroupID & ", " & _
                            "ExcludedProductGroupID=" & ExcludedProductGroupID & ", " & _
                            "BestDeal=" & BestDeal & ", " & _
                            "AllowNegative=" & AllowNegative & ", " & _
                            "ComputeDiscount=" & ComputeDiscount & ", " & _
                            "DiscountAmount=" & DiscountAmount & ", " & _
                            "AmountTypeID=" & AmountTypeID & ", " & _
                            "L1Cap=" & L1Cap & ", L2DiscountAmt=" & L2DiscountAmt & ", L2AmountTypeID=" & L2AmountTypeID & ", L2Cap=" & L2Cap & ", L3DiscountAmt=" & L3DiscountAmt & ", L3AmountTypeID=" & L3AmountTypeID & ", " & _
                            "ItemLimit=" & ItemLimit & ", " & _
                            "WeightLimit=" & Math.Round(CDec(WeightLimit), 3) & ", " & _
                            "IsWeightTotal=" & IIf(IsWeightTotal, "1", "0") & ", " & _
                            "DollarLimit=" & DollarLimit & ", " & _
                            "ChargebackDeptID=" & ChargebackDeptID & ", " & _
                            "DecliningBalance=" & DecliningBalance & ", " & _
                            "RetroactiveDiscount=" & RetroactiveDiscount & ", " & _
                            "UserGroupID=" & UserGroupID & ", " & _
                            "LastUpdate=getdate(), " & _
                            "FlexNegative='" & FlexNeg & "', " & _
                            "PercentFixedRounding=" & PercentFixedRounding & ", "
        MyCommon.QueryStr += "SVProgramID=" & IIf(SVProgramID > 0, SVProgramID.ToString, "Null")
        MyCommon.QueryStr += ", ScorecardID=" & IIf(ScorecardID > 0, ScorecardID, "NULL") & ", ScorecardDesc=" & IIf(ScorecardDesc = "", "NULL", "'" & ScorecardDesc & "'") & ""
        MyCommon.QueryStr += " where DiscountID=" & eDiscountID & ";"
        MyCommon.LRT_Execute()
		
        'Save multilanguage values:
        'ScorecardDesc
        MLI.ItemID = eDiscountID
        MLI.MLTableName = "CPE_DiscountTranslations"
        MLI.MLColumnName = "ScorecardDesc"
        MLI.MLIdentifierName = "DiscountID"
        MLI.StandardTableName = "CPE_Discounts"
        MLI.StandardColumnName = "ScorecardDesc"
        MLI.StandardIdentifierName = "DiscountID"
        MLI.StandardValue = ScorecardDesc
        MLI.InputName = "ScorecardDesc"
        Localization.SaveTranslationInputs(MyCommon, MLI, Request.QueryString, 2)
          
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
                MyCommon.LRTsp.Parameters.Add("@ReceiptDescription", SqlDbType.NVarChar, 100).Value = MyCommon.Strip_Quotes(Trim(Request.QueryString("tier" & i & "_rdesc")))
                MyCommon.LRTsp.Parameters.Add("@DiscountAmount", SqlDbType.Decimal, 12).Value = MyCommon.Extract_Val(Request.QueryString("tier" & i & "_l1discountamt"))
                MyCommon.LRTsp.Parameters.Add("@ItemLimit", SqlDbType.Int, 4).Value = MyCommon.Extract_Val(Request.QueryString("tier" & i & "_itemlimit"))
                MyCommon.LRTsp.Parameters.Add("@WeightLimit", SqlDbType.Decimal, 12).Value = MyCommon.Extract_Val(Request.QueryString("tier" & i & "_weightlimit"))
                MyCommon.LRTsp.Parameters.Add("@IsWeightTotal", SqlDbType.Bit, 1).Value = IIf(Request.QueryString("tier" & i & "_isweighttotal") <> "0", 1, 0)
                MyCommon.LRTsp.Parameters.Add("@DollarLimit", SqlDbType.Decimal, 12).Value = MyCommon.Extract_Val(Request.QueryString("tier" & i & "_dollarlimit"))
                MyCommon.LRTsp.Parameters.Add("@SPRepeatLevel", SqlDbType.TinyInt, 4).Value = MyCommon.Extract_Val(Request.QueryString("tier" & i & "_sprepeatlevel"))
                MyCommon.LRTsp.Parameters.Add("@PKID", SqlDbType.Int, 4).Direction = ParameterDirection.Output
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
                MLI.StandardValue = MyCommon.Strip_Quotes(Trim(Request.QueryString("tier" & i & "_rdesc")))
                MLI.InputName = "tier" & i & "_rdesc"
                Localization.SaveTranslationInputs(MyCommon, MLI, Request.QueryString, 2)
            Next
        End If
          
        'If special pricing, create the CPE_SpecialPricing entries
        If (AmountTypeID = AmountType_t.AMT_TYPE_SPECIAL_PRICING) Then
            'Delete any existing special pricing levels
            MyCommon.QueryStr = "delete from CPE_SpecialPricing with (RowLock) where DiscountID=" & eDiscountID & ";"
            MyCommon.LRT_Execute()
            For i = 1 To TierLevels
                'Find the PKID (DiscountTierID) of each tier
                SPLevels = MyCommon.Extract_Val(Request.QueryString("tier" & i & "_levels"))
                SPHighestLevel = MyCommon.Extract_Val(Request.QueryString("tier" & i & "_highestlevel"))
                MyCommon.QueryStr = "select PKID as DiscountTierID from CPE_DiscountTiers with (NoLock) " & _
                                    "where DiscountID=" & eDiscountID & " and TierLevel=" & i & ";"
                rst = MyCommon.LRT_Select
                If rst.Rows.Count > 0 Then
                    DiscountTierID = MyCommon.NZ(rst.Rows(0).Item("DiscountTierID"), 0)
                End If
                saveLevel = 0
                For l = 1 To SPHighestLevel
                    'Save each level
                    If Request.QueryString("tier" & i & "_level" & l) <> Nothing Then
                        saveLevel = saveLevel + 1

                        If Not (Decimal.TryParse(Request.QueryString("tier" & i & "_level" & l), SPValue)) Then
                            Throw New Exception(Copient.PhraseLib.Lookup("CPEoffer-reward-discount.invalidpricepoint", LanguageID))
                        Else
                            'SPValue = MyCommon.Extract_Val(Request.QueryString("tier" & i & "_level" & l))
                            MyCommon.QueryStr = "insert into CPE_SpecialPricing (DiscountID, DiscountTierID, Value, LevelID) values " & _
                                                "(" & eDiscountID & ", " & DiscountTierID & ", " & SPValue & ", " & saveLevel & ");"
                            MyCommon.LRT_Execute()
                        End If
                    End If
                Next
            Next
        End If
          
        ' Update the CPE_Incentives table
        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
        MyCommon.LRT_Execute()
          
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
        If (Request.QueryString("level") <> "") Then
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
                                        "values (" & DiscountID & ", 1, '" & MyCommon.Strip_Quotes(Trim(Request.QueryString("tier1_rdesc"))) & "', " & MyCommon.Extract_Val(LevelValues(i)) & ", " & _
                                        "0, 0, 0, " & (i + 1) & ", getdate());"
                    MyCommon.LRT_Execute()
                Next
            
            End If
        End If
    End Sub
</script>
<%
    If (IsTemplate) Then
        Send("<script type=""text/javascript"" language=""javascript"">")
        Send("  xmlhttpPost(""TemplateFeeds.aspx?OfferID=" & OfferID & "&PageName=" & Server.UrlEncode(MyCommon.AppName) & "&PageEditable=" & (Not Disallow_Edit) & """)")
        If (Not OverrideFields Is Nothing AndAlso OverrideFields.Count > 0) Then
            Send("  var elem = null;")
            For Each de As DictionaryEntry In OverrideFields
                Send("  elem = document.mainform." & de.Key.ToString & ";")
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
      <% If (CloseAfterSave) Then %>
        window.close();
      <% End If %>
      toggleflexneg()
      <% If Not ChargebackSet and LoadDefaultChargeback Then %>
        handleChargebackDept(<%sendb(DefaultChrgBack) %>);
      <% End If %>
</script>
<div id="fadeDiv">
</div>
<div id="foldercreate" class="folderdialog" style="position: relative; top: 0px;
    left: 40px; width: 350px; height: 90px">
    <div class="foldertitlebar">
        <span class="dialogtitle">select receipt text message</span> <span class="dialogclose"
            onclick="toggleDialog('foldercreate', false);">X</span>
    </div>
    <div class="dialogcontents">
        <div id="receiptmsgerror" style="display: none; color: red;">
        </div>
        <table>
            <tr>
                <td>
                    <select id="BaseMessages" style="width: 280px">
                        <%
                            MyCommon.QueryStr = "Select ReceiptTextID, ReceiptTextMsg from CPE_ReceiptTextMessages where isnull(BaseReceiptTextID,0)=0 order by ReceiptTextID"
                            rstMsgDetails = MyCommon.LRT_Select
                            Send("<option value=""""></option>")
                            For Each rowMsgDetails In rstMsgDetails.Rows
                                Send("            <option value=""" & rowMsgDetails.Item("ReceiptTextMsg") & """>" & rowMsgDetails.Item("ReceiptTextMsg") & " </option>")
                            Next
                        %>
                    </select>
                </td>
                <td>
                    <input type="button" name="btnpicrecmsg" id="btnpicrecmsg" value="Add" onclick="javascript:placerecmessageLocal('foldercreate');" />
                </td>
            </tr>
        </table>
    </div>
</div>
<%
done:
    MyCommon.Close_LogixRT()
    Send_BodyEnd()
    Logix = Nothing
    MyCommon = Nothing
%>
