<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Register Src="~/logix/UserControls/UDFListControl.ascx" TagPrefix="udf" TagName="udflist" %>
<%@ Register Src="~/logix/UserControls/UDFJavaScript.ascx" TagPrefix="udf" TagName="udfjavascript" %>
<%@ Register Src="~/logix/UserControls/UDFSaveControl.ascx" TagPrefix="udf" TagName="udfsave" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="CMS" %>
<% 
    
    ' *****************************************************************************
    ' * FILENAME: CPEoffer-gen.aspx 
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
    Dim MyCPEOffer As New Copient.EIW
    Dim Logix As New Copient.LogixInc
    Dim Localization As Copient.Localization
    Dim rst As DataTable
    Dim rstTemplates As DataTable
    Dim rst3, rst4 As DataTable
    Dim row As DataRow
    Dim rowTemplates As DataRow
    Dim OfferID As Long = Request.QueryString("OfferID")
    Dim OfferName As String = ""
    Dim UpdateLevel As Integer = 0
    Dim rst2 As DataTable
    Dim crmdt As DataTable
    Dim row2 As DataRow
    Dim IsTemplate As Boolean
    Dim IsTemplateVal As String = ""
    Dim ActiveSubTab As Integer = 91
    Dim IntroID As String = "intro"
    Dim Disallow_AdvancedOption As Boolean = True
    Dim Disallow_EmployeeFiltering As Boolean = True
    Dim Disallow_ProductionDates As Boolean = True
    Dim Disallow_Limits As Boolean = True
    Dim Disallow_Tiers As Boolean = True
    Dim Disallow_Priority As Boolean = True
    Dim Disallow_Sweepstakes As Boolean = True
    Dim Disallow_Conditions As Boolean = True
    Dim Disallow_Rewards As Boolean = True
    Dim Disallow_ExecutionEngine As Boolean = True
    Dim Disallow_UserDefinedFields As Boolean = True
    Dim Disallow_CRMEngine As Boolean = True
    Dim FromTemplate As Boolean
    Dim EmployeesExcluded As Boolean
    Dim EmployeesOnly As Boolean
    Dim ReportingImp As Boolean = False
    Dim ReportingRed As Boolean = False
    Dim EmployeeFiltered As Boolean
    Dim ExtOfferID As String = ""
    Dim ShowInboundOutboundBox As Boolean = True
    Dim ProdStartDate As Date
    Dim ProdEndDate As Date
    Dim StartDateParsed, EndDateParsed As Boolean
    Dim TestStartDate, TestEndDate As Date
    Dim roid As Integer
    Dim DuplicateName As Boolean = False
    Dim infoMessage As String = ""
    Dim modMessage As String = ""
    Dim Handheld As Boolean = False
    Dim sqlBuf As New StringBuilder()
    Dim DeferToEOS As Boolean = False
    Dim DeferToEOSDisabled As Boolean = False
    Dim i As Integer = 0
    Dim SelectedBanners, EditableBanners As ArrayList
    Dim IsEditableBanner As Boolean = False
    Dim AllowMultipleBanners As Boolean = False
    Dim BannersEnabled As Boolean = False
    Dim ExportToEDW As Boolean
    Dim Favorite As Boolean
    Dim TempInt As Integer = 0
    Dim StatusCode As Copient.LogixInc.STATUS_FLAGS = Copient.LogixInc.STATUS_FLAGS.STATUS_UNKNOWN
    Dim StatusText As String = ""
    Dim Popup As Boolean = False
    Dim Issuance As Boolean = False
    Dim IssuanceDetails As String = ""
    Dim InboundCRMEngineID As Integer
    Dim ChargebackVendorID As Integer = 0
    Dim IsMfgCoupon As Boolean = False
    Dim ExtName As String = ""
    Dim SelectedList As String = ""
    Dim TierLevels As Integer = 1
    Dim MaxTiers As Integer = 1
    Dim IsUniqueProd As Boolean = False
    Dim AccumEnabled As Boolean = False
    Dim FuelEnabled As Boolean = False
    Dim ShortStartDate, ShortEndDate As String
    Dim ShortTestStartDate, ShortTestEndDate As String
    Dim StartDT, EndDT As Date
    Dim TestStartDT, TestEndDT As Date
    Dim VendorCouponCode As String = ""
    Dim ChangeExtID As Boolean = False
    Dim DescriptLength As Boolean = False
    Dim Description As String = ""
    Dim HeaderExists As Boolean = False
    Dim DisplayTierLevel As String = ""
    Dim IsFootworthy As Boolean = False
    Dim IsDeployable As Boolean = False
    Dim HasEIW As Boolean = False
    Dim IsEIWDateLocked As Boolean = False
    Dim ErrorPhrase As String = ""
    Dim EngineID As Integer = 2
    Dim EngineSubTypeID As Integer = 0
    Dim EnginePhraseID As Integer = 0
    Dim EngineSubTypePhraseID As Integer = 0
    Dim AutoTransferable As Boolean = False
    Dim PromptForReward As Boolean = False
    Dim ScorecardID As Integer = 0
    Dim ScorecardDesc As String = ""
   Dim OldProductionStartDate As Date
   Dim OldProductionEndDate As Date
   Dim OldTestingStartDate As Date
   Dim OldTestingEndDate As Date
    Dim HasAnyCustomer As Boolean 'indicates that the offer customer group condition us using the AnyCustomer group
    Dim TempQueryStr As String
    Dim FolderStartDate As String = ""
    Dim FolderEndDate As String = ""
    Dim DisplayOfferAd As Boolean = False
    Dim OfferAdFields As Boolean = False
    Dim PageOfferAdValue As String = ""
    Dim BlockOfferAdValue As String = ""
    Dim CopyTextOfferAdValue As String = ""
    Dim CoverageAdValue As String = ""
    Dim SaleEventAdValue As String = ""
    Dim rst_AdFields As DataTable
    Dim rst_AdFieldsByOffer As DataTable
    Dim rstAdFieldDetails As DataTable
    Dim rowAdFieldDetails As DataRow
    Dim DefaultSaleEvent As String = ""
    Dim BlankTierLevelSent As Boolean = False
    Dim selectDatePicker As Integer = MyCommon.Extract_Val(MyCommon.NZ(MyCommon.Fetch_SystemOption(161), 0))
    Dim Status As Integer
    Dim m_Offer As CMS.AMS.Contract.IOffer
    Dim MLI As New Copient.Localization.MultiLanguageRec
    Dim NumberofFolders As Integer = 0
    Dim bDiscount As Boolean = False
    
    
    Dim udfSaverst As DataTable
    Dim udfrst As DataTable
    Dim UDFHistory As String
    Dim AllowSpecialCharacters As String
    AllowSpecialCharacters = MyCommon.Fetch_SystemOption(171)
    
    ''' 
    
    
    MLI.ItemID = OfferID
    MLI.MLTableName = "OfferTranslations"
    MLI.MLIdentifierName = "OfferID"
    MLI.StandardTableName = "CPE_Incentives"
    MLI.StandardIdentifierName = "IncentiveID"
    Dim rstTemp As DataTable
    Dim DefaultAsLogixID As Boolean
  
    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If
  
    Response.Expires = 0
    MyCommon.AppName = "CPEoffer-gen.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
    Localization = New Copient.Localization(MyCommon)
    CMS.AMS.CurrentRequest.Resolver.AppName = "CPEoffer-gen.aspx"
    m_Offer = CMS.AMS.CurrentRequest.Resolver.Resolve(Of CMS.AMS.Contract.IOffer)()
    Try
        If MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(129)) = 0 Then
            DisplayOfferAd = False
        Else
            DisplayOfferAd = True
        End If
    Catch ex As Exception
        DisplayOfferAd = False
    End Try
  
    If DisplayOfferAd = True Then
        MyCommon.QueryStr = "select Id from SysObjects where xtype='U' and name='OfferAdFields'"
        rst_AdFields = MyCommon.LRT_Select()
        If rst_AdFields.Rows.Count > 0 Then
            OfferAdFields = True
        End If
    End If
  
    BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
    AllowMultipleBanners = (MyCommon.Fetch_SystemOption(67) = "1")
    
    HasAnyCustomer = CPEOffer_Has_AnyCustomer(MyCommon, OfferID)
  
    'Set default Impression and Redemption defaults
    ReportingImp = (MyCommon.Fetch_CPE_SystemOption(84) = "1")
    ReportingRed = (MyCommon.Fetch_CPE_SystemOption(85) = "1")
    If Request.QueryString("new") <> "" Then
        MyCommon.QueryStr = "update CPE_Incentives set EnableImpressRpt=" & IIf(ReportingImp, "1", "0") & ", EnableRedeemRpt=" & IIf(ReportingRed, "1", "0") & " where IncentiveID=" & OfferID
        MyCommon.LRT_Execute()
    End If
  
    Popup = IIf((Request.QueryString("Popup") <> "") AndAlso (Request.QueryString("Popup") <> "0"), True, False)
  
    MyCommon.QueryStr = "select RewardOptionID, TierLevels from CPE_RewardOptions with (NoLock) where TouchResponse=0 and IncentiveID=" & OfferID
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
        roid = rst.Rows(0).Item("RewardOptionID")
        TierLevels = rst.Rows(0).Item("TierLevels")
    End If
    If (Request.QueryString.ToString().IndexOf("tierlevels") > -1 AndAlso Request.QueryString("tierlevels") = "") Then
        ' the tier levels was passed in as a blank value so convert it to the invalid value of zero so an error message is sent back to the user.
        BlankTierLevelSent = True
        DisplayTierLevel = ""
    ElseIf Request.QueryString("tierlevels") <> "" Then
        DisplayTierLevel = MyCommon.Extract_Val(Request.QueryString("tierlevels"))
    Else
        ' used for page load value for the textbox
        DisplayTierLevel = TierLevels
    End If
    MaxTiers = MyCommon.Fetch_SystemOption(89)
  
    'If the offer already has stuff that isn't compatible with tiers, force the MaxTiers to 1
    MyCommon.QueryStr = "select IncentiveInstantWinID from CPE_IncentiveInstantWin with (NoLock) where RewardOptionID=" & roid & " and Deleted=0;"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
        MaxTiers = 1
    End If
  
    'Determine if the offer has an enterprise instant win condition
    MyCommon.QueryStr = "select IncentiveEIWID from CPE_IncentiveEIW with (NoLock) where RewardOptionID=" & roid & ";"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
        HasEIW = True
    End If
    MyCommon.QueryStr = "select * from FolderItems where LinkID=" & OfferID
    rst = MyCommon.LRT_Select
    NumberofFolders = rst.Rows.Count
    'Set the favorite boolean and the updatelevel
    If OfferID > 0 Then
        MyCommon.QueryStr = "Select Favorite, UpdateLevel from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID & ";"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
            Favorite = MyCommon.NZ(rst.Rows(0).Item("Favorite"), False)
            UpdateLevel = MyCommon.NZ(rst.Rows(0).Item("UpdateLevel"), 0)
        End If
    End If
  
    If (Request.QueryString("mode") = "ChangeExtID") Then
        ExtOfferID = MyCommon.Extract_Val(Request.QueryString("ExtOffer"))
        If ExtOfferID = 0 Then
            MyCommon.QueryStr = "update CPE_Incentives set ClientOfferID=NULL where IncentiveID=" & Request.QueryString("OfferID")
            MyCommon.LRT_Execute()
        ElseIf ExtOfferID > 0 Then
            MyCommon.QueryStr = "update CPE_Incentives set ClientOfferID=" & ExtOfferID & " where IncentiveID=" & Request.QueryString("OfferID")
            MyCommon.LRT_Execute()
        End If
    End If
  
    'Find if there are any unique product flags for this roid
    MyCommon.QueryStr = "select UniqueProduct from CPE_IncentiveProductGroups where RewardOptionID=" & roid & " and UniqueProduct=1 and Deleted=0"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then IsUniqueProd = True
  
    'Check to see if this offer has accumulation; if so, the tier level input will be disabled, since accumulation and multi-tiers are incompatible
    MyCommon.QueryStr = "select ProductGroupID, AccumMin from CPE_IncentiveProductGroups with (NoLock) where RewardOptionID=" & roid & " and ExcludedProducts=0;"
    rst = MyCommon.LRT_Select()
    If rst.Rows.Count > 0 Then
        For Each row In rst.Rows
            If MyCommon.NZ(rst.Rows(0).Item("AccumMin"), 0) > 0 Then
                AccumEnabled = True
            End If
        Next
    Else
        MyCommon.QueryStr = "select ProductGroupID, AccumMin from CPE_IncentiveProductGroups with (NoLock) where RewardOptionID=" & roid & " and ExcludedProducts=1;"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
            For Each row In rst.Rows
                If MyCommon.NZ(rst.Rows(0).Item("AccumMin"), 0) > 0 Then
                    AccumEnabled = True
                End If
            Next
        End If
    End If
  
  
    'See if this offer is footworthy -- ie, can be allowed to have a footer priority
    MyCommon.QueryStr = "dbo.pa_CPE_IsFootworthy"
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
    MyCommon.LRTsp.Parameters.Add("@ROID", SqlDbType.BigInt).Value = roid
    MyCommon.LRTsp.Parameters.Add("@IsFootworthy", SqlDbType.Bit).Direction = ParameterDirection.Output
    MyCommon.LRTsp.ExecuteNonQuery()
    IsFootworthy = MyCommon.LRTsp.Parameters("@IsFootworthy").Value
    MyCommon.Close_LRTsp()
  
    'Check for an Airmiles offer
    If (Request.QueryString("OfferID") <> "") Then
        MyCommon.QueryStr = "select O.EngineID, CPE.EngineSubTypeID from CPE_Incentives as CPE " & _
                            "inner join OfferIDs O on O.OfferID=CPE.IncentiveID " & _
                            "where CPE.IncentiveID =" & Request.QueryString("OfferID")
        rst = MyCommon.LRT_Select()
        If rst.Rows.Count > 0 Then
            EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), -1)
            EngineSubTypeID = MyCommon.NZ(rst.Rows(0).Item("EngineSubTypeID"), -1)
        End If
    End If
  
    ' Check for Discount on the offer; if present, then we need to disable Defer to EOS
    MyCommon.QueryStr = "select DeliverableID from CPE_Deliverables with (NoLock) " & _
                        "where DeliverableTypeID = 2 and Deleted=0 " & _
                        "  and RewardOptionID = " & roid
    rst = MyCommon.LRT_Select()
    If rst.Rows.Count > 0 Then
        DeferToEOSDisabled = True
        bDiscount = True
    End If

  
    'Save
    If (Request.QueryString("save") <> "") AndAlso Not (MyCommon.Fetch_SystemOption(191) AndAlso NumberofFolders > 1) Then
    
        'Check the Description length
        If Request.QueryString("form_description") <> "" Then
            Description = MyCommon.Parse_Quotes(Request.QueryString("form_description"))
            If Description.Length <= 1000 Then
                DescriptLength = True
            End If
        Else
            DescriptLength = True
        End If
        If DisplayOfferAd = True Then
            If OfferAdFields = True Then
                PageOfferAdValue = Request.QueryString("PageAd")
                BlockOfferAdValue = Request.QueryString("BlockAd")
                CopyTextOfferAdValue = Request.QueryString("CTextAd")
                CopyTextOfferAdValue = Replace(CopyTextOfferAdValue, "'", "`")
                CopyTextOfferAdValue = Replace(CopyTextOfferAdValue, Chr(34), "``")
                CoverageAdValue = Request.QueryString("CoverageAd")
                SaleEventAdValue = Request.QueryString("SaleEventAd")
                If SaleEventAdValue = "--" Then
                    MyCommon.QueryStr = "Select top 1 isnull(AdFieldValue,'') from AdDetails where Methodtype='S' and DefaultValue=1 order by AdFieldID"
                    rst_AdFieldsByOffer = MyCommon.LRT_Select
                    If rst_AdFieldsByOffer.Rows.Count > 0 Then
                        If rst_AdFieldsByOffer.Rows(0)(0) <> "" Then
                            DefaultSaleEvent = rst_AdFieldsByOffer.Rows(0)(0)
                        End If
                    End If
                    If DefaultSaleEvent <> "" Then
                        SaleEventAdValue = DefaultSaleEvent
                    End If
                End If
            End If
        End If
        'Get Folder Start/End dates for the validation that user can not change the offer date which is not falling in range of folder dates.
        '192 (Offer dates should be within folder dates)
        If (MyCommon.Fetch_SystemOption(192) = "1") Then
            MyCommon.QueryStr = "Select Startdate,Enddate from Folders Fs inner join FolderItems FI on FI.folderid=Fs.folderid " & _
                                " where FI.linkid=" & OfferID
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count > 0) Then
                If (Not IsDBNull(rst.Rows(0).Item("Startdate")) OrElse Not IsDBNull(rst.Rows(0).Item("EndDate"))) Then
                    FolderStartDate = rst.Rows(0).Item("Startdate")
                    FolderEndDate = rst.Rows(0).Item("Enddate")
                End If
            End If
        End If
        'Get the current production start and end dates (prior to saving the new ones).
        'We'll use these below if there are any EIW conditions that need to be rerandomized.
      MyCommon.QueryStr = "select StartDate, EndDate, TestingStartDate, TestingEndDate from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID & ";"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
         OldProductionStartDate = MyCommon.NZ(rst.Rows(0).Item("StartDate"), "1/1/1900")
         OldProductionEndDate = MyCommon.NZ(rst.Rows(0).Item("EndDate"), "1/1/1900")
         OldTestingStartDate = MyCommon.NZ(rst.Rows(0).Item("TestingStartDate"), "1/1/1900")
         OldTestingEndDate = MyCommon.NZ(rst.Rows(0).Item("TestingEndDate"), "1/1/1900")
        End If
    
        ' Run query to check for mgfcoupon/discount compatibility
        MyCommon.QueryStr = "select DI.DiscountID, DI.DiscountTypeID, RO.RewardOptionID, RO.IncentiveID, I.ManufacturerCoupon, DI.AmountTypeID " & _
                            "from CPE_Discounts as DI with (NoLock) " & _
                            "inner join CPE_Deliverables as DE with (NoLock) on DE.OutputID=DI.DiscountID " & _
                            "inner join CPE_RewardOptions as RO with (NoLock) on DE.RewardOptionID=RO.RewardOptionID " & _
                            "inner join CPE_Incentives as I with (NoLock) on I.IncentiveID=RO.IncentiveID " & _
                            "where I.IncentiveID=" & OfferID & " and DI.Deleted=0 and DE.Deleted=0 and DE.DeliverableTypeID=2;"
        rst = MyCommon.LRT_Select
        ' Also, run a query to see if there's a category that has this offer as its base offer
        MyCommon.QueryStr = "select OfferCategoryID from OfferCategories as OC with (NoLock) where OC.Deleted=0 and BaseOfferID=" & OfferID & " and OfferCategoryID=(" & _
                            "  select IsNull(PromoClassID, 0) from CPE_Incentives where IncentiveID=" & OfferID & ");"
        rst2 = MyCommon.LRT_Select
        If (rst.Rows.Count > 0 AndAlso MyCommon.NZ(rst.Rows(0).Item("DiscountTypeID"), 0) = 3 AndAlso MyCommon.NZ(rst.Rows(0).Item("AmountTypeID"), 0) <> 1 AndAlso Request.QueryString("mfgCoupon") = "on") Then
            infoMessage = Copient.PhraseLib.Lookup("ueoffer-gen.InvalidMfgCoupon", LanguageID)
        ElseIf (Request.QueryString("productionstart") = "" Or Request.QueryString("productionend") = "") Then
            infoMessage = Copient.PhraseLib.Lookup("CPEoffer_gen.noproductiondates", LanguageID)
      
        ElseIf (MyCommon.Extract_Val(Request.QueryString("priority")) = 900) AndAlso (MyCommon.Extract_Val(Request.QueryString("footerpriority")) < 0 OrElse MyCommon.Extract_Val(Request.QueryString("footerpriority")) > 99) Then
            infoMessage = Copient.PhraseLib.Lookup("CPEoffer_gen.badpriority", LanguageID)
        ElseIf (MyCommon.Extract_Val(Request.QueryString("priority")) = 900) AndAlso (MyCommon.Extract_Val(Request.QueryString("tierlevels")) > 1) Then
            infoMessage = Copient.PhraseLib.Lookup("CPEoffer_gen.badfootertiers", LanguageID)
        ElseIf DescriptLength = False Then
            infoMessage = Copient.PhraseLib.Lookup("error.description", LanguageID)
        ElseIf (HasEIW) And Not (Request.QueryString("issuance") = "on") Then
            infoMessage = Copient.PhraseLib.Lookup("ueoffer-gen.InvalidDisableIssuance", LanguageID)
        ElseIf ((rst2.Rows.Count > 0) AndAlso (MyCommon.Extract_Val(Request.QueryString("form_Category")) <> MyCommon.NZ(rst2.Rows(0).Item("OfferCategoryID"), 0))) Then
            infoMessage = Copient.PhraseLib.Lookup("ueoffer-gen.InvalidCategoryChange", LanguageID)
        ElseIf DisplayOfferAd = True AndAlso (Request.QueryString("PageAd") <> "" AndAlso (IsNumeric(Request.QueryString("PageAd")) = False Or MyCommon.Extract_Val(Request.QueryString("PageAd")) < 0)) Then
            infoMessage = "Invalid Page advertising field value."
        ElseIf DisplayOfferAd = True AndAlso (Request.QueryString("BlockAd") <> "" AndAlso (IsNumeric(Request.QueryString("BlockAd")) = False Or MyCommon.Extract_Val(Request.QueryString("BlockAd")) < 0)) Then
            infoMessage = "Invalid Block advertising field value."
        ElseIf DisplayOfferAd = True AndAlso (SaleEventAdValue <> "" AndAlso SaleEventAdValue = "--") Then
            infoMessage = "Invalid Sale event type advertising field value."
        Else
            StartDateParsed = Date.TryParse(Request.QueryString("productionstart"), MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, ProdStartDate)
            EndDateParsed = Date.TryParse(Request.QueryString("productionend"), MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, ProdEndDate)
            If (StartDateParsed AndAlso EndDateParsed) Then
                StartDateParsed = Date.TryParse(Request.QueryString("testingstart"), MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, TestStartDate)
                If (Not StartDateParsed) Then TestStartDate = ProdStartDate
                EndDateParsed = Date.TryParse(Request.QueryString("testingend"), MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, TestEndDate)
                If (Not EndDateParsed) Then TestEndDate = ProdEndDate
        
                ' check for an incentive already with that name
                MyCommon.QueryStr = "select IncentiveName from CPE_Incentives with (NoLock) " & _
                                    "where Deleted=0 and IncentiveName = @Name and IncentiveID <> @OfferID " & _
                                    " union all " & _
                                    "select Name from Offers with (NoLock) " & _
                                    "where Deleted=0 and Name = @Name and OfferID <> @OfferID "
                MyCommon.DBParameters.Add("@Name", SqlDbType.NVarChar, 200).Value = GetCgiValue("form_name")
                MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = GetCgiValue("OfferID")
                rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                DuplicateName = (rst.Rows.Count > 0)
        
                Dim SQLParamList As New List(Of SqlClient.SqlParameter)()
                sqlBuf.Append("Update CPE_Incentives with (RowLock) set ")
                sqlBuf.Append("IncentiveName=@IncentiveName,")
                sqlBuf.Append("Description=@Description,")
                If (GetCgiValue("form_Category") <> "") Then
                    sqlBuf.Append("PromoClassID=@PromoClassID,")
                    SQLParamList.Add(New SqlClient.SqlParameter("@PromoClassID", SqlDbType.Int) With {.Value = MyCommon.Extract_Val(GetCgiValue("form_Category"))})
                End If
                If (GetCgiValue("priority") <> "") Then
                    If (MyCommon.Extract_Val(GetCgiValue("priority")) = 900) Then
                        sqlBuf.Append("Priority=@Priority,")
                        SQLParamList.Add(New SqlClient.SqlParameter("@Priority", SqlDbType.Int) With {.Value = MyCommon.Extract_Val(GetCgiValue("footerpriority")) + 900})
                    Else
                        sqlBuf.Append("Priority=@Priority,")
                        SQLParamList.Add(New SqlClient.SqlParameter("@Priority", SqlDbType.Int) With {.Value = MyCommon.Extract_Val(GetCgiValue("priority"))})
                    End If
                End If
                If (GetCgiValue("crmengine") <> "") Then
                    sqlBuf.Append("CRMEngineID=@CRMEngineID,")
                    SQLParamList.Add(New SqlClient.SqlParameter("@CRMEngineID", SqlDbType.Int) With {.Value = MyCommon.Extract_Val(GetCgiValue("crmengine"))})
                End If
                If (IsUniqueProd) Then
                    sqlBuf.Append("P3DistQtyLimit=1,")
                ElseIf (GetCgiValue("limit3") <> "") Then
                    sqlBuf.Append("P3DistQtyLimit=@P3DistQtyLimit,")
                    SQLParamList.Add(New SqlClient.SqlParameter("@P3DistQtyLimit", SqlDbType.Int) With {.Value = MyCommon.Extract_Val(GetCgiValue("limit3"))})
                End If
                If (IsUniqueProd) Then
                    sqlBuf.Append("P3DistPeriod=1,")
                ElseIf (GetCgiValue("limit3period") <> "") Then
                    sqlBuf.Append("P3DistPeriod=@P3DistPeriod,")
                    SQLParamList.Add(New SqlClient.SqlParameter("@P3DistPeriod", SqlDbType.Int) With {.Value = MyCommon.Extract_Val(GetCgiValue("limit3period"))})
                End If
                If (IsUniqueProd) Then
                    sqlBuf.Append("P3DistTimeType=2,")
                ElseIf (GetCgiValue("P3DistTimeType") <> "") Then
                    sqlBuf.Append("P3DistTimeType=@P3DistTimeType,")
                    SQLParamList.Add(New SqlClient.SqlParameter("@P3DistTimeType", SqlDbType.Int) With {.Value = MyCommon.Extract_Val(GetCgiValue("P3DistTimeType"))})
                End If
                sqlBuf.Append("StartDate=@StartDate,")
                sqlBuf.Append("EndDate=@EndDate,")
                sqlBuf.Append("TestingStartDate=@TestingStartDate,")
                sqlBuf.Append("TestingEndDate=@TestingEndDate,")
                sqlBuf.Append("EmployeesOnly=@EmployeesOnly,")
                sqlBuf.Append("EnableImpressRpt=@EnableImpressRpt,")
                sqlBuf.Append("EnableRedeemRpt=@EnableRedeemRpt,")
                sqlBuf.Append("EmployeesExcluded=@EmployeesExcluded,")
                If (MyCommon.Extract_Val(GetCgiValue("priority")) >= 900) Then
                    sqlBuf.Append("DeferCalcToEOS=1,")
                Else
                    sqlBuf.Append("DeferCalcToEOS=@DeferCalcToEOS,")
                    SQLParamList.Add(New SqlClient.SqlParameter("@DeferCalcToEOS", SqlDbType.Bit) With {.Value = IIf(GetCgiValue("deferToEOS") = "on", 1, 0)})
                End If
                sqlBuf.Append("ExportToEDW=@ExportToEDW,")
                sqlBuf.Append("Favorite=@Favorite,")
                If GetCgiValue("InboundCRMEngineID") <> "" Then
                    sqlBuf.Append("InboundCRMEngineID=@InboundCRMEngineID,")
                    SQLParamList.Add(New SqlClient.SqlParameter("@InboundCRMEngineID", SqlDbType.Int) With {.Value = MyCommon.Extract_Val(GetCgiValue("InboundCRMEngineID"))})
                End If
                sqlBuf.Append("SendIssuance=@SendIssuance,")
                sqlBuf.Append("ChargebackVendorID=@ChargebackVendorID,")
                sqlBuf.Append("ManufacturerCoupon=@ManufacturerCoupon,")
                sqlBuf.Append("RestrictedRedemption=@RestrictedRedemption,")
                sqlBuf.Append("PromptForReward=" & IIf(Request.QueryString("PromptForReward") = "on", 1, 0) & ",")
                sqlBuf.Append("VendorCouponCode=@VendorCouponCode, ")
                sqlBuf.Append("AutoTransferable=@AutoTransferable,")
                sqlBuf.Append("ScorecardID=@ScorecardID,")
                sqlBuf.Append("ScorecardDesc=@ScorecardDesc,")
                sqlBuf.Append("LastUpdate=getdate(), ")
                sqlBuf.Append("LastUpdatedByAdminID=@LastUpdatedByAdminID, ")
                'If the offer is Airmiles allow for the changeing of the external ID
                If EngineID = 2 AndAlso EngineSubTypeID = 2 Then
                    If (GetCgiValue("ExtID") <> "") Then
                        sqlBuf.Append("ClientOfferID = @ClientOfferID,")
                        SQLParamList.Add(New SqlClient.SqlParameter("@ClientOfferID", SqlDbType.NVarChar, 20) With {.Value = Logix.TrimAll(Left(GetCgiValue("ExtID"), 20))})
                    Else
                        sqlBuf.Append("ClientOfferID=NULL,")
                    End If
                End If
                sqlBuf.Append("StatusFlag=1 ")
                sqlBuf.Append("where IncentiveID=@IncentiveID ")
                MyCommon.QueryStr = sqlBuf.ToString()
                SQLParamList.Add(New SqlClient.SqlParameter("@IncentiveName", SqlDbType.NVarChar, 100) With {.Value = Logix.TrimAll(GetCgiValue("form_name"))})
                SQLParamList.Add(New SqlClient.SqlParameter("@Description", SqlDbType.NVarChar, 1000) With {.Value = GetCgiValue("form_description")})
                SQLParamList.Add(New SqlClient.SqlParameter("@StartDate", SqlDbType.DateTime) With {.Value = ProdStartDate})
                SQLParamList.Add(New SqlClient.SqlParameter("@EndDate", SqlDbType.DateTime) With {.Value = ProdEndDate})
                SQLParamList.Add(New SqlClient.SqlParameter("@TestingStartDate", SqlDbType.DateTime) With {.Value = TestStartDate})
                SQLParamList.Add(New SqlClient.SqlParameter("@TestingEndDate", SqlDbType.DateTime) With {.Value = TestEndDate})
                SQLParamList.Add(New SqlClient.SqlParameter("@EmployeesOnly", SqlDbType.Bit) With {.Value = IIf(GetCgiValue("employeesonly") = "on", 1, 0)})
                SQLParamList.Add(New SqlClient.SqlParameter("@EnableImpressRpt", SqlDbType.Bit) With {.Value = IIf(GetCgiValue("reportingimp") = "on", 1, 0)})
                SQLParamList.Add(New SqlClient.SqlParameter("@EnableRedeemRpt", SqlDbType.Bit) With {.Value = IIf(GetCgiValue("reportingred") = "on", 1, 0)})
                SQLParamList.Add(New SqlClient.SqlParameter("@EmployeesExcluded", SqlDbType.Bit) With {.Value = IIf(GetCgiValue("employeesexcluded") = "on", 1, 0)})
                SQLParamList.Add(New SqlClient.SqlParameter("@ExportToEDW", SqlDbType.Bit) With {.Value = IIf(GetCgiValue("exporttoedw") = "on", 1, 0)})
                SQLParamList.Add(New SqlClient.SqlParameter("@Favorite", SqlDbType.Bit) With {.Value = IIf(GetCgiValue("favorite") = "on", 1, 0)})
                SQLParamList.Add(New SqlClient.SqlParameter("@SendIssuance", SqlDbType.Int) With {.Value = IIf(GetCgiValue("issuance") = "on", 1, 0)})
                SQLParamList.Add(New SqlClient.SqlParameter("@ChargebackVendorID", SqlDbType.Int) With {.Value = MyCommon.Extract_Val(GetCgiValue("vendor"))})
                SQLParamList.Add(New SqlClient.SqlParameter("@ManufacturerCoupon", SqlDbType.Int) With {.Value = IIf(GetCgiValue("mfgCoupon") = "on", 1, 0)})
                SQLParamList.Add(New SqlClient.SqlParameter("@RestrictedRedemption", SqlDbType.Bit) With {.Value = IIf(GetCgiValue("RestrictedRdmpt") = "on", 1, 0)})
                SQLParamList.Add(New SqlClient.SqlParameter("@VendorCouponCode", SqlDbType.NVarChar, 20) With {.Value = GetCgiValue("vendorCouponCode")})
                SQLParamList.Add(New SqlClient.SqlParameter("@AutoTransferable", SqlDbType.Bit) With {.Value = IIf(GetCgiValue("autotransferable") = "on", 1, 0)})
                SQLParamList.Add(New SqlClient.SqlParameter("@ScorecardID", SqlDbType.Int) With {.Value = IIf(GetCgiValue("ScorecardID") <> "", MyCommon.Extract_Val(GetCgiValue("ScorecardID")), 0)})
                SQLParamList.Add(New SqlClient.SqlParameter("@ScorecardDesc", SqlDbType.NVarChar, 100) With {.Value = GetCgiValue("ScorecardDesc")})
                SQLParamList.Add(New SqlClient.SqlParameter("@LastUpdatedByAdminID", SqlDbType.Int) With {.Value = AdminUserID})
                SQLParamList.Add(New SqlClient.SqlParameter("@IncentiveID", SqlDbType.BigInt) With {.Value = MyCommon.Extract_Val(GetCgiValue("OfferID"))})
                If (ProdEndDate < ProdStartDate) Then
                    infoMessage = Copient.PhraseLib.Lookup("offer-gen.baddate", LanguageID)
            ElseIf (ProdStartDate < CDate("1/1/2000")) Then  'the date value is earlier than we can allow
               infoMessage = Copient.PhraseLib.Lookup("term.datetooearly", LanguageID)
                ElseIf (ProdEndDate >= CDate("1/1/9999")) Then  'the date value is larger that what we can allow
                    infoMessage = Copient.PhraseLib.Lookup("term.datetoolarge", LanguageID)
                ElseIf (TestEndDate < TestStartDate) Then
                    infoMessage = Copient.PhraseLib.Lookup("offer-gen.baddate", LanguageID)
            ElseIf (TestStartDate < CDate("1/1/2000")) Then  'the date value is earlier than we can allow
               infoMessage = Copient.PhraseLib.Lookup("term.datetooearly", LanguageID)
                ElseIf (TestEndDate >= CDate("1/1/9999")) Then  'the date value is larger that what we can allow
                    infoMessage = Copient.PhraseLib.Lookup("term.datetoolarge", LanguageID)
                ElseIf (TestStartDate < CDate(IIf(FolderStartDate = "", TestStartDate, FolderStartDate)) OrElse TestEndDate > CDate(IIf(FolderEndDate = "", TestEndDate, FolderEndDate)) OrElse ProdStartDate < CDate(IIf(FolderStartDate = "", ProdStartDate, FolderStartDate)) OrElse ProdEndDate > CDate(IIf(FolderEndDate = "", ProdEndDate, FolderEndDate))) AndAlso (MyCommon.Fetch_SystemOption(192) = "1") Then
                    infoMessage = Copient.PhraseLib.Lookup("folders.OfferNotInFolderDateRange", LanguageID) & "(" & FolderStartDate & " - " & FolderEndDate & ")"
                ElseIf DuplicateName Then
                    infoMessage = Copient.PhraseLib.Lookup("offer-gen.nameused", LanguageID)
                ElseIf Logix.TrimAll(Request.QueryString("form_name")) = "" Then
                    infoMessage = Copient.PhraseLib.Lookup("offer-gen.noname", LanguageID)
                ElseIf (Request.QueryString("limit3") <> "" AndAlso (Not Integer.TryParse(Request.QueryString("limit3"), TempInt) OrElse (TempInt < 0 AndAlso MyCommon.Extract_Val(Request.QueryString("p3type")) > 0) OrElse (TempInt <= 0 AndAlso MyCommon.Extract_Val(Request.QueryString("p3type")) = 0))) Then
                    infoMessage = Copient.PhraseLib.Lookup("offer-gen.badlimit", LanguageID)
                ElseIf (Request.QueryString("limit3period") <> "" AndAlso (Not Integer.TryParse(Request.QueryString("limit3period"), TempInt) OrElse (TempInt < 0 AndAlso MyCommon.Extract_Val(Request.QueryString("p3type")) > 0) OrElse (TempInt <= 0 AndAlso MyCommon.Extract_Val(Request.QueryString("p3type")) = 0))) Then
                    infoMessage = Copient.PhraseLib.Lookup("offer-gen.badlimit", LanguageID)
                ElseIf BlankTierLevelSent OrElse (Request.QueryString("tierlevels") <> "" AndAlso (Not Integer.TryParse(Request.QueryString("tierlevels"), TempInt) OrElse TempInt < 1 OrElse TempInt > MaxTiers)) Then
                    infoMessage = Copient.PhraseLib.Lookup("offer-gen.invalidtierscpe", LanguageID) & MaxTiers & "."
                ElseIf (MyCommon.Extract_Val(Request.QueryString("tierlevels")) > 1 AndAlso (Request.QueryString("crmengine") = 1 OrElse Request.QueryString("crmengine") = 2)) Then
                    infoMessage = Copient.PhraseLib.Lookup("offer-gen.invalidoutbound", LanguageID)
                Else
                    MyCommon.DBParameters.AddRange(SQLParamList.ToArray)
                    MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
                    m_Offer.UpdateOfferDefaultGroupName(MyCommon.Extract_Val(Request.QueryString("OfferID")), 2, MyCommon.Parse_Quotes(Logix.TrimAll(Request.QueryString("form_name"))))
                    ' if TierLevels has changed, update the value
                    If MyCommon.Extract_Val(Request.QueryString("tierlevels")) <> TierLevels Then
                        MyCommon.QueryStr = "update CPE_RewardOptions set TierLevels=" & IIf(MyCommon.Extract_Val(Request.QueryString("tierlevels")) <= 0, 1, MyCommon.Extract_Val(Request.QueryString("tierlevels"))) & " " & _
                                            "where RewardOptionID=" & roid & ";"
                        MyCommon.LRT_Execute()
                        If MyCommon.Extract_Val(Request.QueryString("tierlevels")) < TierLevels Then
                            ' TierLevels value has been lowered, so run the procedure to delete the now-orphaned tier records
                            MyCommon.QueryStr = "dbo.pa_CPE_PurgeDecrementedTiers"
                            MyCommon.Open_LRTsp()
                            MyCommon.LRTsp.Parameters.Add("@ROID", SqlDbType.BigInt).Value = roid
                            MyCommon.LRTsp.Parameters.Add("@NewTierLevel", SqlDbType.BigInt).Value = MyCommon.Extract_Val(Request.QueryString("tierlevels"))
                            MyCommon.LRTsp.ExecuteNonQuery()
                            MyCommon.Close_LRTsp()
                        End If
                    End If
          
                    'Storing in OfferAdvertisingFields
                    If DisplayOfferAd = True Then
                        If OfferAdFields = True Then
                            Dim QueryStr As String = ""
                            MyCommon.QueryStr = "select OfferID from OfferAdFields with (NoLock) where OfferID=" & OfferID & " and Deleted=0"
                            rst_AdFieldsByOffer = MyCommon.LRT_Select()
                            If rst_AdFieldsByOffer.Rows.Count > 0 Then
                                QueryStr = "Update OfferAdFields with (RowLock) set "
                                If PageOfferAdValue = "" Then
                                    QueryStr = QueryStr & "Page=NULL"
                                Else
                                    QueryStr = QueryStr & "Page='" & PageOfferAdValue & "'"
                                End If
                                If BlockOfferAdValue = "" Then
                                    QueryStr = QueryStr & ",Block=NULL"
                                Else
                                    QueryStr = QueryStr & ",Block='" & BlockOfferAdValue & "'"
                                End If
                                If CoverageAdValue = "--" Then
                                    QueryStr = QueryStr & ",CoverageMethodID=NULL"
                                Else
                                    QueryStr = QueryStr & ",CoverageMethodID='" & CoverageAdValue & "'"
                                End If
                                MyCommon.QueryStr = QueryStr & ",CopyText='" & CopyTextOfferAdValue & "',SaleEventTypeID='" & SaleEventAdValue & "' where OfferID=" & OfferID
                            Else
                                QueryStr = "Insert into OfferAdFields(OfferID,Page,Block,CoverageMethodID,CopyText,SaleEventTypeID,Deleted) values("
                                QueryStr = QueryStr & OfferID
                                If PageOfferAdValue = "" Then
                                    QueryStr = QueryStr & ",NULL"
                                Else
                                    QueryStr = QueryStr & ",'" & PageOfferAdValue & "'"
                                End If
                                If BlockOfferAdValue = "" Then
                                    QueryStr = QueryStr & ",NULL"
                                Else
                                    QueryStr = QueryStr & ",'" & BlockOfferAdValue & "'"
                                End If
                                If CoverageAdValue = "--" Then
                                    QueryStr = QueryStr & ",NULL"
                                Else
                                    QueryStr = QueryStr & ",'" & CoverageAdValue & "'"
                                End If
                                MyCommon.QueryStr = QueryStr & ",'" & CopyTextOfferAdValue & "','" & SaleEventAdValue & "',0)"
                            End If
                            MyCommon.LRT_Execute()
                        End If
                    End If
		  
                    ' when an offer is flagged as a manufacturer coupon offer, best deal should be disabled.
                    If Request.QueryString("mfgCoupon") = "on" Then
                        MyCommon.QueryStr = "select DISC.DiscountID from CPE_Discounts DISC with (NoLock) " & _
                                            "inner join CPE_Deliverables DEL with (NoLock) on DEL.OutputID=DISC.DiscountID and DEL.DeliverableTypeID=2 " & _
                                            "  and DEL.RewardOptionPhase=3 and DEL.Deleted=0 " & _
                                            "inner join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID=DEL.RewardOptionID and RO.Deleted=0 " & _
                                            "where DISC.Deleted=0 and RO.IncentiveID=" & MyCommon.Extract_Val(Request.QueryString("OfferID"))
                        rst = MyCommon.LRT_Select
                        If rst.Rows.Count Then
                            For Each row In rst.Rows
                                MyCommon.QueryStr = "update CPE_Discounts with (RowLock) set BestDeal=0 where DiscountID=" & MyCommon.NZ(row.Item("DiscountID"), 0)
                                MyCommon.LRT_Execute()
                            Next
                        End If
                    End If
          
                End If
                
                
                
                
                '''BEGIN User Defined Fields saving block
                If MyCommon.Fetch_SystemOption(156) = "1" Then
%>
<udf:udfsave id="udfsavecontrol" runat="server" />
<%                   
                    
    infoMessage = udfsavecontrol.infoMessage
    UDFHistory = udfsavecontrol.UDFHistory
End If ' MyCommon.Fetch_SystemOption(156) = "1" 
'''END User Defined Fields saving block
                

''populate UDF string values from adhoc table to UserDefinedFieldsValues
'    MyCommon.QueryStr = "dbo.pt_PopulateUDFStringValues"
'       MyCommon.Open_LRTsp()
'       MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
'       MyCommon.LRTsp.ExecuteNonQuery()
'       MyCommon.Close_LRTsp()
	  
IsTemplate = (Request.QueryString("IsTemplate") = "IsTemplate")
If (IsTemplate) Then
    'Update template permissions
    Dim form_Disallow_ExecutionEngine As Integer = IIf(Request.QueryString("Disallow_ExecutionEngine") = "on", 1, 0)
    Dim form_Disallow_CRMEngine As Integer = IIf(Request.QueryString("Disallow_CRMEngine") = "on", 1, 0)
    Dim form_Disallow_EmployeeFiltering As Integer = IIf(Request.QueryString("Disallow_EmployeeFiltering") = "on", 1, 0)
    Dim form_Disallow_ProductionDates As Integer = IIf(Request.QueryString("Disallow_ProductionDates") = "on", 1, 0)
    Dim form_Disallow_Limits As Integer = IIf(Request.QueryString("Disallow_Limits") = "on", 1, 0)
    Dim form_Disallow_Tiers As Integer = IIf(Request.QueryString("Disallow_Tiers") = "on", 1, 0)
    Dim form_Disallow_Priority As Integer = IIf(Request.QueryString("Disallow_Priority") = "on", 1, 0)
    Dim form_Disallow_Sweepstakes As Integer = IIf(Request.QueryString("Disallow_Sweepstakes") = "on", 1, 0)
    Dim form_Disallow_UserDefinedFields As Integer = IIf(Request.QueryString("Disallow_UserDefinedFields") = "on", 1, 0)
    Dim form_Disallow_AdvancedOption As Integer = IIf(Request.QueryString("Disallow_AdvancedOption") = "on", 1, 0)
    MyCommon.QueryStr = "update TemplatePermissions with (RowLock) set Disallow_EmployeeFiltering=" & form_Disallow_EmployeeFiltering & _
                        " , Disallow_ProductionDates=" & form_Disallow_ProductionDates & _
                        " , Disallow_limits=" & form_Disallow_Limits & _
                        " , Disallow_Tiers=" & form_Disallow_Tiers & _
                        " , Disallow_Priority=" & form_Disallow_Priority & _
                        " , Disallow_CRMEngine=" & form_Disallow_CRMEngine & _
                        " , Disallow_ExecutionEngine=" & form_Disallow_ExecutionEngine & _
                        " , Disallow_Sweepstakes=" & form_Disallow_Sweepstakes & _
                        " , Disallow_AdvancedOption=" & form_Disallow_AdvancedOption & _
   " , Disallow_UserDefinedFields=" & form_Disallow_UserDefinedFields & " where OfferID=" & OfferID

    MyCommon.LRT_Execute()
End If
        
'Update the banner engine (if necessary)
If (BannersEnabled AndAlso AllowMultipleBanners AndAlso Request.QueryString("bannerschanged") = "true") Then
    ' first clear out the existing banners
    MyCommon.QueryStr = "delete from BannerOffers with (RowLock) where OfferID =" & OfferID & ";"
    MyCommon.LRT_Execute()
    ' add the selected banners
    If (Request.QueryString("bannerids") <> "") Then
        For i = 0 To Request.QueryString.GetValues("bannerids").GetUpperBound(0)
            MyCommon.QueryStr = "insert into BannerOffers with (RowLock) (BannerID, OfferID) values (" & MyCommon.Extract_Val(Request.QueryString.GetValues("bannerids")(i)) & "," & OfferID & ");"
            MyCommon.LRT_Execute()
        Next i
    End If
End If
        
MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-edit", LanguageID))
        
'If the offer has an EIW condition, and the dates have changed, rerandomize the triggers
If HasEIW Then
               If (OldProductionStartDate <> ProdStartDate) OrElse (OldProductionEndDate <> ProdEndDate) Then
        MyCPEOffer.RandomizeTriggersByOffer(OfferID)
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-instantwin-randomize", LanguageID))
    End If
End If
Else
infoMessage = Copient.PhraseLib.Lookup("CPEoffer_gen.noproductiondates", LanguageID)
End If
End If
'Update multi-language:
'Name
Dim ErrorMessage As String = ""
MLI.MLColumnName = "OfferName"
MLI.StandardColumnName = "IncentiveName"
MLI.StandardValue = OfferName
MLI.InputName = "form_Name"
Localization.SaveTranslationInputs(MyCommon, MLI, Request.QueryString, 2, ErrorMessage)
If (ErrorMessage <> "") Then
infoMessage = ErrorMessage
End If
'Description
MLI.MLColumnName = "OfferDescription"
MLI.StandardColumnName = "Description"
MLI.StandardValue = Description
MLI.InputName = "form_Description"
Localization.SaveTranslationInputs(MyCommon, MLI, Request.QueryString, 2)
'Limit scorecard description
MLI.MLColumnName = "LimitScorecardDesc"
MLI.StandardColumnName = "ScorecardDesc"
MLI.StandardValue = MyCommon.Parse_Quotes(Request.QueryString("ScorecardDesc"))
MLI.InputName = "ScorecardDesc"
Localization.SaveTranslationInputs(MyCommon, MLI, Request.QueryString, 2)
        
      'Don't deploy if there is an error.  If there is an infoMessage, there is something wrong.
      If infoMessage = "" Then
'Check for "Expired Due To Back Dating" of previously deployed offer and handle expiration due to back-dating and the need to send the offer to the Store for cleanup to prevent Sanity Check error as last step in SAVE
      If ((OldProductionEndDate >= Today And ProdEndDate < Today) OrElse (OldTestingEndDate >= Today And TestEndDate < Today) OrElse (OldProductionStartDate >= Today And ProdStartDate < Today) OrElse (OldTestingStartDate >= Today And TestStartDate < Today)) Then
'if offer already deployed To Store then deploy delete to Store even though Deleted Offer may be Locked to deactivate and clean-up offer, so Sanity Check does not Fail
MyCommon.QueryStr = "select LastUpdateLevel from PromoEngineUpdateLevels with (NoLock) " & _
     "where LinkID=" & OfferID & " and EngineID=2 and ItemType=1;"
rst = MyCommon.LRT_Select
If (rst.Rows.Count > 0) Then
If (MyCommon.NZ(rst.Rows(0).Item("LastUpdateLevel"), 0) > 0) Then
    IsDeployable = IsDeployableOffer(MyCommon, OfferID, roid, ErrorPhrase)
    If (IsDeployable) Then
        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=2, UpdateLevel=UpdateLevel+1, LastDeployClick=getdate(), DeployDeferred=0, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", ExpireLocked=1 where IncentiveID=" & OfferID & ";"
        MyCommon.LRT_Execute()
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-deploy", LanguageID))
        Response.Status = "301 Moved Permanently"
        Response.AddHeader("Location", "CPEoffer-sum.aspx?OfferID=" & OfferID)
    End If
End If
End If
End If
End If
End If ''end save block
                
                
If MyCommon.Fetch_SystemOption(156) = "1" Then
MyCommon.QueryStr = "delete from OfferUDFStringValues where OfferID = " & OfferID
MyCommon.LRT_Execute()
MyCommon.QueryStr = "update UserDefinedFieldsValues set deleted = 0 where deleted= 1 and OfferID = " & OfferID ' any UDF marked for deletion should be deleted when saved
MyCommon.LRT_Execute()
End If
  
If (Request.QueryString("Deploy") <> "") Then
IsDeployable = IsDeployableOffer(MyCommon, OfferID, roid, ErrorPhrase)
If (IsDeployable) Then
MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=2, UpdateLevel=UpdateLevel+1, LastDeployClick=getdate(), DeployDeferred=0, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", ExpireLocked=1 where IncentiveID=" & OfferID & ";"
MyCommon.LRT_Execute()
MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-deploy", LanguageID))
Response.Status = "301 Moved Permanently"
Response.AddHeader("Location", "CPEoffer-sum.aspx?OfferID=" & OfferID)
GoTo done
Else
infoMessage = Copient.PhraseLib.Lookup(ErrorPhrase, LanguageID)
End If
End If
  
Dim IsRestrictedRedemption As Boolean = False
  
If (Request.QueryString("OfferID") <> "") Then
MyCommon.QueryStr = "select IncentiveID, OID.EngineID, PE.PhraseID as EnginePhraseID, PEST.PhraseID as EngineSubTypePhraseID," & _
                    "IsTemplate, FromTemplate, ClientOfferID, IncentiveName, CPE.Description, PromoClassID, CRMEngineID, Priority, " & _
                    " StartDate, EndDate, EveryDOW, TestingStartDate, TestingEndDate," & _
                    " P2DistQtyLimit, P2DistTimeType, P2DistPeriod, P3DistQtyLimit, P3DistTimeType, P3DistPeriod," & _
                    " EnableImpressRpt, EnableRedeemRpt, CPE.CreatedDate, CPE.LastUpdate, CPE.Deleted, CPEOARptDate, CPEOADeploySuccessDate, CPEOADeployRpt," & _
                    " CRMRestricted, StatusFlag, EmployeesOnly, EmployeesExcluded, DeferCalcToEOS, ExportToEDW," & _
                    " Favorite, OC.Description as CategoryName, SendIssuance, InboundCRMEngineID, ChargebackVendorID, ManufacturerCoupon, VendorCouponCode," & _
                    " AutoTransferable, CPE.EngineSubTypeID, CPE.RestrictedRedemption, CPE.ScorecardID, CPE.ScorecardDesc, CPE.PromptForReward " & _
                    "from CPE_Incentives as CPE with (NoLock) " & _
                    "left join OfferCategories as OC with (NoLock) on CPE.PromoClassID=OfferCategoryID " & _
                    "left join OfferIDs as OID with (NoLock) on OID.OfferID=CPE.IncentiveID " & _
                    "left join PromoEngines as PE with (NoLock) on PE.EngineID=OID.EngineID " & _
                    "left join PromoEngineSubTypes as PEST with (NoLock) on PEST.PromoEngineID=OID.EngineID and PEST.SubTypeID=OID.EngineSubTypeID " & _
                    "where IncentiveID=" & Request.QueryString("OfferID") & ";"
rst = MyCommon.LRT_Select
If rst.Rows.Count < 1 Then
infoMessage = Copient.PhraseLib.Lookup("term.notfound", LanguageID)
Else
IsTemplate = MyCommon.NZ(rst.Rows(0).Item("IsTemplate"), False)
FromTemplate = MyCommon.NZ(rst.Rows(0).Item("FromTemplate"), False)
OfferName = MyCommon.NZ(rst.Rows(0).Item("IncentiveName"), "")
EmployeesOnly = MyCommon.NZ(rst.Rows(0).Item("EmployeesOnly"), False)
ReportingImp = MyCommon.NZ(rst.Rows(0).Item("EnableImpressRpt"), False)
ReportingRed = MyCommon.NZ(rst.Rows(0).Item("EnableRedeemRpt"), False)
EmployeesExcluded = MyCommon.NZ(rst.Rows(0).Item("EmployeesExcluded"), False)
EmployeeFiltered = EmployeesOnly Or EmployeesExcluded
ExtOfferID = MyCommon.NZ(rst.Rows(0).Item("ClientOfferID"), "")
DeferToEOS = MyCommon.NZ(rst.Rows(0).Item("DeferCalcToEOS"), False)
ExportToEDW = MyCommon.NZ(rst.Rows(0).Item("ExportToEDW"), False)
Favorite = MyCommon.NZ(rst.Rows(0).Item("Favorite"), False)
Issuance = (MyCommon.NZ(rst.Rows(0).Item("SendIssuance"), 0) = 1)
InboundCRMEngineID = MyCommon.NZ(rst.Rows(0).Item("InboundCRMEngineID"), 0)
ChargebackVendorID = MyCommon.NZ(rst.Rows(0).Item("ChargebackVendorID"), 0)
IsMfgCoupon = (MyCommon.NZ(rst.Rows(0).Item("ManufacturerCoupon"), 0) = 1)
AutoTransferable = MyCommon.NZ(rst.Rows(0).Item("AutoTransferable"), False)
EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), 0)
EnginePhraseID = MyCommon.NZ(rst.Rows(0).Item("EnginePhraseID"), 0)
EngineSubTypeID = MyCommon.NZ(rst.Rows(0).Item("EngineSubTypeID"), 0)
EngineSubTypePhraseID = MyCommon.NZ(rst.Rows(0).Item("EngineSubTypePhraseID"), 0)
VendorCouponCode = MyCommon.NZ(rst.Rows(0).Item("VendorCouponCode"), "")
EngineSubTypeID = MyCommon.NZ(rst.Rows(0).Item("EngineSubTypeID"), 0)
IsRestrictedRedemption = MyCommon.NZ(rst.Rows(0).Item("RestrictedRedemption"), False)
PromptForReward = MyCommon.NZ(rst.Rows(0).Item("PromptForReward"), False)
ScorecardID = MyCommon.NZ(rst.Rows(0).Item("ScorecardID"), 0)
ScorecardDesc = MyCommon.NZ(rst.Rows(0).Item("ScorecardDesc"), "")
End If
    
' If this is a footer priority offer, turn on the DeferCalcToEOS and disable the control
If MyCommon.NZ(rst.Rows(0).Item("Priority"), 0) >= 900 Then
DeferToEOSDisabled = True
DeferToEOS = True
Else If (MyCommon.Fetch_CPE_SystemOption(192) AndAlso Not bDiscount)  Then
DeferToEOS = True
End If
    
If (IsTemplate Or FromTemplate) Then
' lets dig the permissions if its a template
MyCommon.QueryStr = "select * from templatepermissions with (NoLock) where OfferID=" & OfferID
rstTemplates = MyCommon.LRT_Select
If (rstTemplates.Rows.Count > 0) Then
For Each rowTemplates In rstTemplates.Rows
    ' ok there are some rows for the template
    Disallow_EmployeeFiltering = MyCommon.NZ(rowTemplates.Item("Disallow_EmployeeFiltering"), True)
    Disallow_ProductionDates = MyCommon.NZ(rowTemplates.Item("Disallow_ProductionDates"), True)
    Disallow_Limits = MyCommon.NZ(rowTemplates.Item("Disallow_Limits"), True)
    Disallow_Tiers = MyCommon.NZ(rowTemplates.Item("Disallow_Tiers"), True)
    Disallow_Priority = MyCommon.NZ(rowTemplates.Item("Disallow_Priority"), True)
    Disallow_Sweepstakes = MyCommon.NZ(rowTemplates.Item("Disallow_Sweepstakes"), True)
    Disallow_Conditions = MyCommon.NZ(rowTemplates.Item("Disallow_Conditions"), True)
    Disallow_Rewards = MyCommon.NZ(rowTemplates.Item("Disallow_Rewards"), True)
    Disallow_ExecutionEngine = MyCommon.NZ(rowTemplates.Item("Disallow_ExecutionEngine"), True)
    Disallow_CRMEngine = MyCommon.NZ(rowTemplates.Item("Disallow_CRMEngine"), True)
    Disallow_UserDefinedFields = MyCommon.NZ(rowTemplates.Item("Disallow_UserDefinedFields"), True)
    Disallow_AdvancedOption = MyCommon.NZ(rowTemplates.Item("Disallow_AdvancedOption"), True)
Next
Else
Disallow_EmployeeFiltering = False
Disallow_ProductionDates = False
Disallow_Limits = False
Disallow_Tiers = False
Disallow_Priority = False
Disallow_Sweepstakes = False
Disallow_Conditions = False
Disallow_Rewards = False
Disallow_ExecutionEngine = False
Disallow_CRMEngine = False
Disallow_UserDefinedFields = False
Disallow_AdvancedOption = False
End If
End If
End If
  
'Check Advertisement Fields based on the OfferID in OfferAdvertisingFields
If DisplayOfferAd = True Then
If OfferAdFields = True Then
MyCommon.QueryStr = "select isnull(Convert(varchar,Page),'') 'Page',isnull(Convert(varchar,Block),'') 'Block',CopyText," & _
                    " isnull(CoverageMethodID,'--') 'CoverageMethodID',isnull(SaleEventTypeID,'') 'SaleEventTypeID' " & _
                    " from OfferAdFields with (NoLock) where OfferID=" & OfferID & " and Deleted=0"
rst_AdFieldsByOffer = MyCommon.LRT_Select()
If rst_AdFieldsByOffer.Rows.Count > 0 Then
PageOfferAdValue = rst_AdFieldsByOffer.Rows(0)(0)
BlockOfferAdValue = rst_AdFieldsByOffer.Rows(0)(1)
CopyTextOfferAdValue = rst_AdFieldsByOffer.Rows(0)(2)
CoverageAdValue = rst_AdFieldsByOffer.Rows(0)(3)
SaleEventAdValue = rst_AdFieldsByOffer.Rows(0)(4)
End If
End If
End If
	
'Check that the External OfferID can be changed
If IsMfgCoupon Or InboundCRMEngineID = 1 Or InboundCRMEngineID = 2 Then
ChangeExtID = False
Else
If InboundCRMEngineID = 0 Then
ChangeExtID = True
Else
MyCommon.QueryStr = "select AllowExtOfferIDChange from ExtCRMInterfaces where ExtInterfaceID=" & InboundCRMEngineID
rst4 = MyCommon.LRT_Select()
If rst4.Rows.Count > 0 Then
If MyCommon.NZ(rst4.Rows(0).Item("AllowExtOfferIDChange"), False) = True Then
    ChangeExtID = True
End If
End If
End If
End If
  
StatusText = Logix.GetOfferStatus(OfferID, LanguageID, StatusCode)
ShowInboundOutboundBox = (MyCommon.Fetch_SystemOption(25) <> "0")
  
If (IsTemplate) Then
ActiveSubTab = 25
IntroID = "intro"
IsTemplateVal = "IsTemplate"
Else
ActiveSubTab = 24
IntroID = "intro"
IsTemplateVal = "Not"
End If
  
Send_HeadBegin("term.offer", "term.general", OfferID)
Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
Send_Metas()
Send_Links(Handheld)
Send_Scripts(New String() {"datePicker.js", "popup.js"})
Send_HeadEnd()
If (IsTemplate) Then
Send_BodyBegin(IIf(Popup, 13, 11))
Else
Send_BodyBegin(IIf(Popup, 3, 1))
End If
%>
<style type="text/css">
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
<script type="text/javascript">
window.name = "CPEofferGen"
var datePickerDivID = "datepicker";

if (window.captureEvents){
  window.captureEvents(Event.CLICK);
  window.onclick=handlePageClick;
} else {
  document.onclick=handlePageClick;
}

<% Send_Calendar_Overrides(MyCommon) %>

function handlePageClick(e) {
  var calFrame = document.getElementById('calendariframe');
  var el=(typeof event!=='undefined')? event.srcElement : e.target        

  if (el != null) {
    var pickerDiv = document.getElementById(datePickerDivID);
    if (pickerDiv != null && pickerDiv.style.visibility == "visible") {
      if (el.id!="production-start-picker" && el.id!="production-end-picker"
      && el.id!="testing-start-picker" && el.id!="testing-end-picker" && el.id!="udf-datevalue-picker") {
        if (!isDatePickerControl(el.className)) {
          pickerDiv.style.visibility = "hidden";
          pickerDiv.style.display = "none"; 
          if (calFrame != null) {
            calFrame.style.visibility = 'hidden';
            calFrame.style.display = "none";
          }
        }
      } else  {
        pickerDiv.style.visibility = "visible";            
        pickerDiv.style.display = "block";     
        if (calFrame != null) {
          calFrame.style.visibility = 'visible';
          calFrame.style.display = "block";
        }
      }
    }
  }
}

function isDatePickerControl(ctrlClass) {
  var retVal = false;
  
  if (ctrlClass != null && ctrlClass.length >= 2) {
    if (ctrlClass.substring(0,2) == "dp") {
      retVal = true;
    }
  }
  return retVal;
}
 
function elmName(){
  for(i=0; i<document.mainform.elements.length; i++) {
    document.mainform.elements[i].disabled=false;
    //alert(document.mainform.elements[i].name)
  }
  return true;
}
function setPlimits(){
  // start out custom
  document.getElementById("p3type").value = 0;
  
  if( document.getElementById("limit3").value == 0 && document.getElementById("limit3period").value == 0
  && ( document.getElementById("P3DistTimeType").value == 2 || document.getElementById("BeginP3TimeTypeID").value == -1) ) {
    // no limit
    document.getElementById("p3type").value = 1;
    }
  else if (document.getElementById("limit3").value == 1) {
    if ( document.getElementById("limit3period").value == 1 ) {
      if ( document.getElementById("P3DistTimeType").value == 2 ) {
        //Once Per Transaction
        document.getElementById("p3type").value = 2;
        }
      else if (document.getElementById("P3DistTimeType").value == 1) {
        //Once Per Day
        document.getElementById("p3type").value = 3;
        }
      }
    else if (document.getElementById("limit3period").value == 7 && document.getElementById("P3DistTimeType").value == 1 ){
      //Once Per Week
      document.getElementById("p3type").value = 4;
      }
    else if (document.getElementById("limit3period").value == 3650 && document.getElementById("P3DistTimeType").value == 1) {
      //Once Per Offer
      document.getElementById("p3type").value = 5;
    }
  }
  // call update to invis the proper stuff if needed
  updateP3limit();
}


function updateP3limit(){
  if(document.getElementById("p3type").value == 0 ){
    document.getElementById("p3row2").style.display =''
    document.getElementById("p3row3").style.display =''
    document.getElementById("p3row4").style.display =''
<% 
  if HasAnyCustomer then 
    send("    document.getElementById(""limit3period"").value=1 ")
    send("    document.getElementById(""P3DistTimeType"").value=2 ")
  end if
%>    
  } else {
    document.getElementById("p3row2").style.display ='none'
    document.getElementById("p3row3").style.display ='none'
    document.getElementById("p3row4").style.display ='none'
    if(document.getElementById("p3type").value == 1){
    // no limit 0 0 2
      document.getElementById("limit3").value=0
      document.getElementById("limit3period").value=0
      document.getElementById("P3DistTimeType").value=2
    }
    else if(document.getElementById("p3type").value == 2){
    // no limit 0 0 2
      document.getElementById("limit3").value=1
      document.getElementById("limit3period").value=1
      document.getElementById("P3DistTimeType").value=2
    }
    else if(document.getElementById("p3type").value == 3){
    // no limit 0 0 2
      document.getElementById("limit3").value=1
      document.getElementById("limit3period").value=1
      document.getElementById("P3DistTimeType").value=1
    }
    else if(document.getElementById("p3type").value == 4){
    // no limit 0 0 2
      document.getElementById("limit3").value=1
      document.getElementById("limit3period").value=7
      document.getElementById("P3DistTimeType").value=1
    }
    else if(document.getElementById("p3type").value == 5){
    // no limit 0 0 2
      document.getElementById("limit3").value=1
      document.getElementById("limit3period").value=3650
      document.getElementById("P3DistTimeType").value=1
    }
  }
  if (document.getElementById("inoutHeader") != null) {
    document.getElementById("inoutHeader").innerHTML = '<% Sendb(Copient.PhraseLib.Lookup("term.inbound/outbound", LanguageID))%>';
  }
}

function handleEmployeeFiltering() {
  var elemFilter = document.getElementById("EmployeeFiltering");
  var elemOnly = document.getElementById("employeesonly");
  var elemExcluded = document.getElementById("employeesexcluded");
  
  if (elemFilter != null && !elemFilter.checked) {
    if (elemOnly != null) {
      elemOnly.checked = false;
    }
    if (elemExcluded != null) {
      elemExcluded.checked = false;
    }
  }
  
  if ( (elemOnly!=null && elemOnly.checked) || (elemExcluded!=null && elemExcluded.checked) ) {
    if (elemFilter != null) {
      elemFilter.checked = true;
    }
  }
}

function toggleEmployee(elemName) {
  var elemFilter = document.getElementById("EmployeeFiltering");
  var elemOnly = document.getElementById("employeesonly");
  var elemExcluded = document.getElementById("employeesexcluded");
  
  if( document.getElementById(elemName).checked==true){
    document.getElementById(elemName).checked=false;
  }
  if ( (elemOnly!=null && elemOnly.checked) || (elemExcluded!=null && elemExcluded.checked) ) {
    if (elemFilter != null) {
      elemFilter.checked = true;
    }
  }
}

function handleAllBanners(elemAll) {
  var elem = null;
  var i = 0;
  
  if (elemAll.checked) {
    elem = document.getElementById('bannerid' + i);
    while (elem != null) {
      elem.checked = false;      
      i++;
      elem = document.getElementById('bannerid' + i);
    }
  }
  document.getElementById('bannerschanged').value='true';
}

function handleBanners(elem) {
  var elemAll = null;
  var i = 0;
  
  if (elem.checked) {
    elem = document.getElementById('allbannerid' + i);
    while (elem != null) {
      elem.checked = false;      
      i++;
      elem = document.getElementById('allbannerid' + i);
    }
  }
  document.getElementById('bannerschanged').value='true';
}

function promptForDeploy() {
  var elem = document.getElementById("IsActive");
  var retVal = true;
  var elemEnd = document.getElementById("productionend");
  var dtNow = new Date();
  var dtEnd = new Date();
  
  if (elem != null && elem.value == "true" && elemEnd != null) {
    dtEnd = getDateFromFormat(elemEnd.value, '<%Sendb(MyCommon.GetAdminUser.Culture.DateTimeFormat.ShortDatePattern) %>');
    retVal = (dtEnd != null);
    if (retVal) {
      dtEnd.setDate(dtEnd.getDate() + 1);
      if (dtEnd < dtNow) {
        retVal = confirm('<%Sendb(Copient.PhraseLib.Lookup("term.expire-confirm", LanguageID)) %>');
        <%
          If MyCommon.Fetch_CPE_SystemOption(80) = 1 Then
            Send("if (retVal == true) {")
            Send("  document.getElementById(""Deploy"").value = 1;")
            Send("}")
          End If
        %>
      }
    } else {
      alert('<%Sendb(Copient.PhraseLib.Lookup("offer-gen.invalidenddate", LanguageID)) %>');
    }
  }
  return retVal;
}

function xmlhttpPost(strURL, mode) {
  var xmlHttpReq = false;
  var self = this;
  
  //document.getElementById("tools").innerHTML = "<div class=\"loading\"><img id=\"clock\" src=\"../images/clock22.png\" \/><br \/>" + '<% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID)) %><\/div>';
  handleWait(true);
  
  // Mozilla/Safari
  if (window.XMLHttpRequest) {
    self.xmlHttpReq = new XMLHttpRequest();
  }
  // IE
  else if (window.ActiveXObject) {
    self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
  }
  strURL += "?" + getQueryString(mode);
  self.xmlHttpReq.open('POST', strURL, true);
  self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
  self.xmlHttpReq.onreadystatechange = function() {
    if (self.xmlHttpReq !=null && self.xmlHttpReq.readyState == 4) {
      updatePage(self.xmlHttpReq.responseText);
    }
  }
  self.xmlHttpReq.send(getQueryString(mode));
}

function getQueryString(mode) {  
  return "Mode=" + mode + "&OfferID=<%Sendb(OfferID)%>&AdminUserID=<%Sendb(AdminUserID)%>";
}

function updatePage(responseMsg) {
  var favImg = document.getElementById("favImg");
  var allUserMsg = '<%Sendb(Copient.PhraseLib.Lookup("offer.allusersfavorited", LanguageID)) %>';
  
  if (responseMsg == 'OK') {
    alert(allUserMsg);
    if (favImg != null) {
      favImg.setAttribute("alt", allUserMsg);
      favImg.setAttribute("title", allUserMsg);
    }
  } else {
    alert(responseMsg);
  }
  handleWait(false);
}

function handleWait(bShow) {
  var elem = document.getElementById("disabledBkgrd");
  
  if (elem != null) {
    elem.style.display = (bShow)  ? 'block' : 'none';
  }
}

function toggleFooterPriority() {
  if (document.getElementById("priority").value < 900) {
    document.getElementById("footerpriority").style.display = 'none';
    document.getElementById("footerpriority").value = '';
    document.getElementById("footerprioritylabel").style.display = 'none';
    document.getElementById("footerprioritynote").style.display = 'none';
  } else {
    document.getElementById("footerpriority").style.display = 'inline';
    document.getElementById("footerpriority").value = '';
    document.getElementById("footerprioritylabel").style.display = 'inline';
    document.getElementById("footerprioritynote").style.display = 'inline';
  }
}

function toggleScorecardText() {
  if (document.getElementById("ScorecardID").value == 0) {
    document.getElementById("scdesc").style.display = 'none';
    document.getElementById("ScorecardDesc").value = '';
  } else {
    document.getElementById("scdesc").style.display = '';
  }
}

function toggleScorecard() {
  var scoreCardExists = document.getElementById("scorecard");
  if(scoreCardExists)
  {
      if (document.getElementById("p3type").value == 1 || document.getElementById("p3type").value == 2) {
        document.getElementById("scorecard").style.display = 'none';
        document.getElementById("ScorecardID").value = 0;
        document.getElementById("ScorecardDesc").value = '';
      } else {
        document.getElementById("scorecard").style.display = '';
      }
  }
}


</script>
<udf:udfjavascript id="udfjavascriptcontrol" runat="server" />
<%
    If (Not Popup) Then
        Send_Bar(Handheld)
        Send_Help(CopientFileName)
        Send_Logos()
        Send_Tabs(Logix, 2)
        Send_Subtabs(Logix, ActiveSubTab, 4, , OfferID)
    End If
  
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
%>
<form id="mainform" name="mainform" action="CPEoffer-gen.aspx" method="get" onsubmit="handleFormElements(this, false);return promptForDeploy();">
<input type="hidden" name="OfferID" id="OfferID" value="<%Sendb(OfferID)%>" />
<input type="hidden" id="form_OfferID" name="form_OfferID" value="<%sendb(OfferID) %>" />
<input type="hidden" name="IsActive" id="IsActive" value="<%Sendb(IIf(StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_ACTIVE OrElse (StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_SCHEDULED AndAlso Logix.IsOfferDeployedToStore(OfferID)), "true", "false"))%>" />
<input type="hidden" name="IsTemplate" id="IsTemplate" value="<% Sendb(IsTemplateVal)%>" />
<input type="hidden" name="Popup" id="Popup" value="<%Sendb(IIf(Popup, 1, 0)) %>" />
<input type="hidden" name="Deploy" id="Deploy" value="" />
<input type="hidden" name="SelectedUDF" id="SelectedUDF" value="" />
<div id="<% Sendb(IntroID)%>">
    <% 
        If rst.Rows.Count > 0 Then
            If (IsTemplate) Then
                Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & ": " & MyCommon.TruncateString(OfferName, 43) & "</h1>")
            Else
                Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & ": " & MyCommon.TruncateString(OfferName, 43) & "</h1>")
            End If
        End If
    %>
    <div id="controls">
        <%
            If Not IsTemplate Then
                If (Logix.UserRoles.EditInstantWinOffers AndAlso EngineID = 2 AndAlso EngineSubTypeID = 1) Then
                    Send_Save()
                ElseIf (Logix.UserRoles.EditOffer AndAlso Not (EngineID = 2 AndAlso EngineSubTypeID = 1)) Then
                    Send_Save()
                End If
            Else
                If (Logix.UserRoles.EditInstantWinOffers AndAlso EngineID = 2 AndAlso EngineSubTypeID = 1) Then
                    Send_Save()
                ElseIf (Logix.UserRoles.EditTemplates AndAlso Not (EngineID = 2 AndAlso EngineSubTypeID = 1)) Then
                    Send_Save()
                End If
            End If
        
            If MyCommon.Fetch_SystemOption(75) Then
                If (OfferID > 0 AndAlso Logix.UserRoles.AccessNotes AndAlso Not Popup) Then
                    Send_NotesButton(3, OfferID, AdminUserID)
                End If
            End If
        %>
    </div>
</div>
<div id="main">
    <%
        If Not IsTemplate Then
            If (MyCommon.NZ(rst.Rows(0).Item("StatusFlag"), 0) <> 2) Then
                If (StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_ACTIVE) AndAlso (MyCommon.NZ(rst.Rows(0).Item("StatusFlag"), 0) > 0) Then
                    modMessage = Copient.PhraseLib.Lookup("alert.modpostdeploy", LanguageID)
                    Send("<div id=""modbar"">" & modMessage & "</div>")
                End If
            End If
        End If
        If MyCommon.Fetch_SystemOption(191) AndAlso NumberofFolders > 1 Then
            infoMessage = infoMessage & " " & "Offer cannot have more than one Folder associated"
        End If
        If (infoMessage <> "") Then
            Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
        End If
        If (IsTemplate) Then
            Send("<div id=""infobar"" class=""red-background"">" & Copient.PhraseLib.Lookup("temp.note", LanguageID) & "</div>")
        End If
      
        ' Send the status bar, but not if it's a new offer or a template, or if there's already a modMessage being shown.
        If rst.Rows.Count < 1 Then
            GoTo done
        End If
        If (Not IsTemplate AndAlso modMessage = "") Then
            MyCommon.QueryStr = "select incentiveId from CPE_Incentives with (NoLock) where CreatedDate = LastUpdate and IncentiveID=" & OfferID
            rst3 = MyCommon.LRT_Select
            If (rst3.Rows.Count = 0) Then
                Send_Status(OfferID, 2)
            End If
        End If
    %>
    <div id="column1">
        <div class="box" id="identification" style="z-index: 50;">
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.identification", LanguageID))%>
                </span>
            </h2>
            <% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID) & ": " & OfferID)%>
            <br />
            <%
                'Allow for the change of the external id to the offer id
                Sendb(Copient.PhraseLib.Lookup("term.externalid", LanguageID) & ": ")
                'If offer is US AirMiles then allow setting the ExtID
                If EngineID = 2 AndAlso EngineSubTypeID = 2 Then
                    Send("<input size=""20"" id=""ExtID"" name=""ExtID"" maxlength=""20"" type=""text"" value=""" & ExtOfferID.Replace("""", "&quot;") & """/><br />")
                Else
                    If ChangeExtID Then
                        If ExtOfferID = "" Then
                            Send(Copient.PhraseLib.Lookup("term.none", LanguageID) & " <small><a href=""CPEoffer-gen.aspx?OfferID=" & OfferID & "&amp;ExtOffer=" & OfferID & "&amp;mode=ChangeExtID"">(" & Copient.PhraseLib.Lookup("CPEoffer-gen.xid-add", LanguageID) & ")</a></small>")
                        Else
                            Send(ExtOfferID & " <small><a href=""CPEoffer-gen.aspx?OfferID=" & OfferID & "&amp;ExtOffer=0&amp;mode=ChangeExtID"">(" & Copient.PhraseLib.Lookup("CPEoffer-gen.xid-rem", LanguageID) & ")</a></small>")
                        End If
                    Else
                        MyCommon.QueryStr = "select DefaultAsLogixID from ExtCRMInterfaces where ExtInterfaceID = " & InboundCRMEngineID
                        rstTemp = MyCommon.LRT_Select
                        DefaultAsLogixID = MyCommon.NZ(rstTemp.Rows(0).Item("DefaultAsLogixID"), False)
                        If (DefaultAsLogixID = True) Then
                            ExtOfferID = OfferID
                            MyCommon.QueryStr = "update CPE_Incentives set ClientOfferID=" & ExtOfferID & " where IncentiveID=" & Request.QueryString("OfferID")
                            MyCommon.LRT_Execute()
                            Send(ExtOfferID)
                        Else
                            Send(ExtOfferID)
                        End If
                    End If
                    Send("<br />")
                End If
                Send(Copient.PhraseLib.Lookup("term.roid", LanguageID) & ": " & roid & "<br />")
                Send(Copient.PhraseLib.Lookup("term.engine", LanguageID) & ": " & Copient.PhraseLib.Lookup(EnginePhraseID, LanguageID) & IIf(EngineSubTypePhraseID > 0, " " & Copient.PhraseLib.Lookup(EngineSubTypePhraseID, LanguageID), "") & "<br />")
                Send(Copient.PhraseLib.Lookup("term.status", LanguageID) & ": " & StatusText & "<br />")
            %>
            <br class="half" />
            <%'Name input
                MLI.MLColumnName = "OfferName"
                MLI.StandardValue = OfferName
                MLI.InputName = "form_Name"
                MLI.InputID = "name"
                MLI.InputType = "text"
                MLI.LabelPhrase = "term.name"
                MLI.MaxLength = 100
                MLI.CSSClass = "longest"
                MLI.CSSStyle = "width:92%;"
                Send(Localization.SendTranslationInputs(MyCommon, MLI, Request.QueryString, 2))
            %>
            <%'Translation input
                MLI.MLColumnName = "OfferDescription"
                MLI.StandardValue = MyCommon.NZ(rst.Rows(0).Item("Description"), "")
                MLI.InputName = "form_Description"
                MLI.InputID = "desc"
                MLI.InputType = "textarea"
                MLI.LabelPhrase = "term.description"
                MLI.MaxLength = 1000
                MLI.CSSClass = "longest"
                MLI.CSSStyle = "width:92%;"
                Send(Localization.SendTranslationInputs(MyCommon, MLI, Request.QueryString, 2))
            %>
            <br class="half" />
            <small>
                <%Send("(" & Copient.PhraseLib.Lookup("CPEoffergen.description", LanguageID) & ")")%></small><br
                    class="half" />
            <br class="half" />
            <label for="category">
                <% Sendb(Copient.PhraseLib.Lookup("term.category", LanguageID))%>:</label><br />
            <select class="medium" id="category" name="form_Category">
                <%
                    ' get the category list from database
                    MyCommon.QueryStr = "select OfferCategoryID, Description from OfferCategories with (NoLock) where Deleted=0 order by Description"
                    rst2 = MyCommon.LRT_Select()
                    For Each row2 In rst2.Rows
                        If (MyCommon.NZ(rst.Rows(0).Item("PromoClassID"), 0) = MyCommon.NZ(row2.Item("OfferCategoryID"), 0)) Then
                            Sendb("<option value=""" & MyCommon.NZ(row2.Item("OfferCategoryID"), 0) & """ selected=""selected"">" & MyCommon.NZ(row2.Item("Description"), "") & "</option>")
                        Else
                            Sendb("<option value=""" & MyCommon.NZ(row2.Item("OfferCategoryID"), 0) & """>" & MyCommon.NZ(row2.Item("Description"), "") & "</option>")
                        End If
                    Next
                %>
            </select>
            <br />
            <br class="half" />
            <label for="vendorCouponCode">
                <% Sendb(Copient.PhraseLib.Lookup("term.vendor-coupon-code", LanguageID))%>:</label><br />
            <input class="medium" id="vendorCouponCode" name="vendorCouponCode" maxlength="20"
                type="text" value="<% sendb(VendorCouponCode.Replace("""", "&quot;")) %>" />&nbsp;
            <br class="half" />
            <hr class="hidden" />
        </div>
        <div class="box" id="priorities">
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.priority", LanguageID))%>
                </span>
            </h2>
            <% If (IsTemplate) Then%>
            <span class="temp">
                <input type="checkbox" class="tempcheck" id="Disallow_Priority" name="Disallow_Priority"
                    <% if(disallow_priority)then send(" checked=""checked""") %> />
                <label for="Disallow_Priority">
                    <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
            </span>
            <br class="printonly" />
            <% End If%>
            <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.priority", LanguageID)) %>">
                <tr>
                    <td style="width: 120px;">
                        <label for="priority">
                            <% Sendb(Copient.PhraseLib.Lookup("CPEoffer-gen.priority", LanguageID) & ":")%></label>
                    </td>
                    <td>
                        <select id="priority" name="priority" onchange="toggleFooterPriority();" <% if(FromTemplate and disallow_priority)then sendb(" disabled=""disabled""") %>>
                            <%
                                MyCommon.QueryStr = "select IncentiveID from CPE_Incentives with (NoLock) where Priority=-1 and Deleted=0;"
                                rst2 = MyCommon.LRT_Select
                                If rst2.Rows.Count > 0 Then
                                    If MyCommon.NZ(MyCommon.Extract_Val(rst2.Rows(0).Item("IncentiveID")), 0) = OfferID Then
                                        'If the current offer is offer with header priority then show the option to select it
                                        HeaderExists = False
                                    Else
                                        HeaderExists = True
                                    End If
                                End If
                                MyCommon.QueryStr = "select PriorityID, Description, PhraseID from CPE_IncentivePriorities with (NoLock) "
                                If HeaderExists Then
                                    If IsFootworthy Then
                                        MyCommon.QueryStr &= "where PriorityID>-1;"
                                    Else
                                        MyCommon.QueryStr &= "where PriorityID>-1 and PriorityID<900;"
                                    End If
                                Else
                                    If IsFootworthy Then
                                        MyCommon.QueryStr &= ";"
                                    Else
                                        MyCommon.QueryStr &= "where PriorityID<900;"
                                    End If
                                End If
                                rst2 = MyCommon.LRT_Select
                                For Each row2 In rst2.Rows
                                    If (MyCommon.NZ(rst.Rows(0).Item("Priority"), 2) = MyCommon.NZ(row2.Item("PriorityID"), 0)) OrElse (MyCommon.NZ(rst.Rows(0).Item("Priority"), 2) >= 900 AndAlso MyCommon.NZ(row2.Item("PriorityID"), 0) >= 900) Then
                                        Send("<option value=""" & MyCommon.NZ(row2.Item("PriorityID"), 0) & """ selected=""selected"">" & Copient.PhraseLib.Lookup(MyCommon.NZ(row2.Item("PhraseID"), 0), LanguageID) & " </option>")
                                    Else
                                        Send("<option value=""" & MyCommon.NZ(row2.Item("PriorityID"), 0) & """>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row2.Item("PhraseID"), 0), LanguageID) & " </option>")
                                    End If
                                Next
                            %>
                        </select>
                    </td>
                </tr>
                <tr>
                    <td>
                        <label id="footerprioritylabel" for="footerpriority" <% Sendb(IIf(MyCommon.NZ(rst.Rows(0).Item("Priority"), 2) < 900, " style=""display:none;""", "")) %>>
                            <% Sendb(Copient.PhraseLib.Lookup("term.footer", LanguageID) & " " & Copient.PhraseLib.Lookup("term.priority", LanguageID) & ":")%></label>
                    </td>
                    <td>
                        <input type="text" class="shortest" id="footerpriority" name="footerpriority" maxlength="2"
                            <% Sendb(IIf(MyCommon.NZ(rst.Rows(0).Item("Priority"), 2) < 900, " style=""display:none;""", "")) %>
                            value="<% Sendb(MyCommon.NZ(rst.Rows(0).Item("Priority"), 0) - 900) %>" />
                        <span id="footerprioritynote" <% Sendb(IIf(MyCommon.NZ(rst.Rows(0).Item("Priority"), 2) < 900, " style=""display:none;""", "")) %>>
                            <% Sendb(Copient.PhraseLib.Lookup("CPEoffer-gen.footernote", LanguageID))%></span>
                    </td>
                </tr>
            </table>
            &nbsp;
            <br class="half" />
            <hr class="hidden" />
        </div>
        <div class="box" id="dates">
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.dates", LanguageID))%>
                </span>
            </h2>
            <% If (IsTemplate) Then%>
            <span class="temp">
                <input type="checkbox" class="tempcheck" id="Disallow_ProductionDates" name="Disallow_ProductionDates"
                    <% if(disallow_productiondates)then send(" checked=""checked""") %> />
                <label for="Disallow_ProductionDates">
                    <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
            </span>
            <br class="printonly" />
            <% End If%>
            <%
                If HasEIW Then
                    If (DateTime.Parse(MyCommon.NZ(rst.Rows(0).Item("StartDate"), "1/1/1900")) < Now) AndAlso (UpdateLevel > 0) Then
                        IsEIWDateLocked = True
                        Send("<p>" & Copient.PhraseLib.Lookup("ueoffer-gen.HasEIW", LanguageID) & "</p>")
                    Else
                        IsEIWDateLocked = False
                        Send("<p>" & Copient.PhraseLib.Lookup("ueoffer-gen.Re-randomizeEIW", LanguageID) & "</p>")
                    End If
                End If
            %>
            <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.dates", LanguageID))%>">
                <tr>
                    <td>
                        <%
                            If rst.Rows.Count > 0 Then
                                ShortStartDate = ""
                                ShortEndDate = ""
                                StartDT = MyCommon.NZ(rst.Rows(0).Item("StartDate"), "1/1/1900")
                                EndDT = MyCommon.NZ(rst.Rows(0).Item("EndDate"), "1/1/1900")
                                If StartDT <> "1/1/1900" Then ShortStartDate = Logix.ToShortDateString(StartDT, MyCommon)
                                If EndDT <> "1/1/1900" Then ShortEndDate = Logix.ToShortDateString(EndDT, MyCommon)
                            Else
                                ShortStartDate = ""
                                ShortEndDate = ""
                            End If
                        %>
                        <label for="productionstart">
                            <% Sendb(Copient.PhraseLib.Lookup("term.production", LanguageID))%>:</label><br />
                        <input type="text" class="short" id="productionstart" name="productionstart" maxlength="10"
                            value="<% sendb(ShortStartDate) %>" <% if((FromTemplate and disallow_productiondates) Or IsEIWDateLocked) then sendb(" disabled=""disabled""") %> />
                        <img src="../images/calendar.png" class="calendar" id="production-start-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
                            title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
                            onclick="displayDatePicker('productionstart', event);" />
                        <% Sendb(Copient.PhraseLib.Lookup("term.to", LanguageID))%>
                        <input type="text" class="short" id="productionend" name="productionend" maxlength="10"
                            value="<% sendb(ShortEndDate) %>" <% if((FromTemplate and disallow_productiondates) Or IsEIWDateLocked) then sendb(" disabled=""disabled""") %> />
                        <img src="../images/calendar.png" class="calendar" id="production-end-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
                            title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
                            onclick="autoDatePicker('productionend', event, <% sendb(selectDatePicker) %>);" />
                        (<% Sendb(MyCommon.GetAdminUser.Culture.DateTimeFormat.ShortDatePattern)%>)<br />
                        <br class="half" />
                    </td>
                </tr>
                <tr>
                    <td>
                        <%
                            If rst.Rows.Count > 0 Then
                                ShortTestStartDate = ""
                                ShortTestEndDate = ""
                                TestStartDT = MyCommon.NZ(rst.Rows(0).Item("TestingStartDate"), "1/1/1900")
                                TestEndDT = MyCommon.NZ(rst.Rows(0).Item("TestingEndDate"), "1/1/1900")
                                If TestStartDT <> "1/1/1900" Then ShortTestStartDate = Logix.ToShortDateString(TestStartDT, MyCommon)
                                If TestEndDT <> "1/1/1900" Then ShortTestEndDate = Logix.ToShortDateString(TestEndDT, MyCommon)
                            Else
                                ShortTestStartDate = ""
                                ShortTestEndDate = ""
                            End If
                        %>
                        <label for="testingstart">
                            <% Sendb(Copient.PhraseLib.Lookup("term.testing", LanguageID))%>:</label><br />
                        <input type="text" class="short" id="testingstart" name="testingstart" maxlength="10"
                            value="<% sendb(ShortTestStartDate) %>" <% if((FromTemplate and disallow_productiondates) Or IsEIWDateLocked) then sendb(" disabled=""disabled""") %> />
                        <img src="../images/calendar.png" class="calendar" id="testing-start-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
                            title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
                            onclick="displayDatePicker('testingstart', event);" />
                        <% Sendb(Copient.PhraseLib.Lookup("term.to", LanguageID))%>
                        <input type="text" class="short" id="testingend" name="testingend" maxlength="10"
                            value="<% sendb(ShortTestEndDate) %>" <% if((FromTemplate and disallow_productiondates) Or IsEIWDateLocked) then sendb(" disabled=""disabled""") %> />
                        <img src="../images/calendar.png" class="calendar" id="testing-end-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
                            title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
                            onclick="displayDatePicker('testingend', event);" />
                        (<% Sendb(MyCommon.GetAdminUser.Culture.DateTimeFormat.ShortDatePattern)%>)<br />
                    </td>
                </tr>
            </table>
            <hr class="hidden" />
        </div>
        <div id="datepicker" class="dpDiv">
        </div>
        <%
            If Request.Browser.Type = "IE6" Then
                Send("<iframe src=""javascript:'';"" id=""calendariframe"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no"" style=""visibility:hidden;display:none;""></iframe>")
            End If
        %>
        <% If DisplayOfferAd = True Then%>
        <% If OfferAdFields = True Then%>
        <div class="box" id="AdvertisingFields">
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.advertisementfields", LanguageID))%></span>
            </h2>
            <label for="lblCopytoTextAd">
                <% Sendb(Copient.PhraseLib.Lookup("term.copytext", LanguageID))%>:</label><br />
            <input class="longest" style="width: 95%;" id="CopyTextofferAd" name="CTextAd" maxlength="125"
                type="text" value="<% sendb(CopyTextOfferAdValue) %>" /><br />
            <label for="lblCoverageMethod">
                <% Sendb(Copient.PhraseLib.Lookup("term.coveragemethod", LanguageID))%>:</label><br />
            <select id="CoverageMethod" name="CoverageAd">
                <%
                    Send("<option value=""--"">Select</option>")
                    MyCommon.QueryStr = "Select AdFieldValue,AdFieldDescription from AdDetails with (NoLock) where MethodType='C' order by AdFieldID"
                    rstAdFieldDetails = MyCommon.LRT_Select
                    For Each rowAdFieldDetails In rstAdFieldDetails.Rows
                        If CoverageAdValue = rowAdFieldDetails.Item("AdFieldValue") Then
                            Send("<option value=""" & rowAdFieldDetails.Item("AdFieldValue") & """ selected=""selected"">" & rowAdFieldDetails.Item("AdFieldDescription") & " </option>")
                        Else
                            Send("<option value=""" & rowAdFieldDetails.Item("AdFieldValue") & """>" & rowAdFieldDetails.Item("AdFieldDescription") & " </option>")
                        End If
                    Next
                %>
            </select>
            <br />
            <label for="lblPageAd">
                <% Sendb(Copient.PhraseLib.Lookup("term.page", LanguageID))%>:</label><br />
            <input class="longest" style="width: 50px;" id="pageofferAd" name="PageAd" maxlength="3"
                type="text" value="<% sendb(PageOfferAdValue) %>" /><br />
            <label for="lblBlockAd">
                <% Sendb(Copient.PhraseLib.Lookup("term.block", LanguageID))%>:</label><br />
            <input class="longest" style="width: 50px;" id="blockofferAd" name="BlockAd" maxlength="3"
                type="text" value="<% sendb(BlockOfferAdValue) %>" /><br />
            <label for="lblSaleEventType">
                <% Sendb(Copient.PhraseLib.Lookup("term.saleeventtype", LanguageID))%>:</label><br />
            <select id="SaleEventType" name="SaleEventAd">
                <%
                    Send("<option value=""--"">Select</option>")
                    MyCommon.QueryStr = "Select AdFieldValue,AdFieldDescription,isnull(DefaultValue,0) 'DefaultValue' from AdDetails with (NoLock) where MethodType='S' order by AdFieldID"
                    rstAdFieldDetails = MyCommon.LRT_Select
                    For Each rowAdFieldDetails In rstAdFieldDetails.Rows
                        If rowAdFieldDetails.Item("DefaultValue") = 1 Or rowAdFieldDetails.Item("DefaultValue") = True Then
                            DefaultSaleEvent = rowAdFieldDetails.Item("AdFieldValue")
                            If SaleEventAdValue = "" Or SaleEventAdValue = "--" Then
                                SaleEventAdValue = DefaultSaleEvent
                            End If
                        End If
                        If SaleEventAdValue = rowAdFieldDetails.Item("AdFieldValue") Then
                            Send("<option value=""" & rowAdFieldDetails.Item("AdFieldValue") & """ selected=""selected"">" & rowAdFieldDetails.Item("AdFieldDescription") & " </option>")
                        Else
                            Send("<option value=""" & rowAdFieldDetails.Item("AdFieldValue") & """>" & rowAdFieldDetails.Item("AdFieldDescription") & " </option>")
                        End If
                    Next
                %>
            </select>
            <br />
            <br class="half" />
            <hr class="hidden" />
        </div>
        <% End If%>
        <% End If%>
    </div>
    <div id="gutter">
    </div>
    <div id="column2">
        <div class="box" id="limits">
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.limits", LanguageID))%>
                </span>
            </h2>
            <% If (IsTemplate) Then%>
            <span class="temp">
                <input type="checkbox" class="tempcheck" id="Disallow_Limits" name="Disallow_Limits"
                    <% if(disallow_limits)then send(" checked=""checked""") %> />
                <label for="Disallow_Limits">
                    <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
            </span>
            <br class="printonly" />
            <% End If%>
            <label for="limit3" style="position: relative;">
                <b>
                    <% Sendb(Copient.PhraseLib.Lookup("term.reward", LanguageID))%>:</b></label>
            <br />
            <%        
                Send("<table summary=""" & Copient.PhraseLib.Lookup("term.rewards", LanguageID) & """>")
                Send("<tr>")
                Send("  <td>")
                Send("    <label for=""p3type"">" & Copient.PhraseLib.Lookup("term.frequency", LanguageID) & ":</label>")
                Send("  </td>")
                Send("  <td>")
                Send("    <select name=""p3type"" id=""p3type"" onchange=""updateP3limit();toggleScorecard();""" & IIf((FromTemplate And Disallow_Limits Or IsUniqueProd), " disabled=""disabled""", "") & ">")
                Send("      <option value=""1"">" & Copient.PhraseLib.Lookup("term.nolimit", LanguageID) & "</option>")
                Send("      <option value=""2"">" & Copient.PhraseLib.Lookup("term.oncepertransaction", LanguageID) & "</option>")
                If Not (HasAnyCustomer) Then
                    Send("      <option value=""3"">" & Copient.PhraseLib.Lookup("term.onceperday", LanguageID) & "</option>")
                    Send("      <option value=""4"">" & Copient.PhraseLib.Lookup("term.onceperweek", LanguageID) & "</option>")
                    Send("      <option value=""5"">" & Copient.PhraseLib.Lookup("term.onceperoffer", LanguageID) & "</option>")
                End If
                Send("      <option value=""0"">" & Copient.PhraseLib.Lookup("term.custom", LanguageID) & "</option>")
                Send("    </select>")
                Send("  </td>")
                Send("</tr>")
                If (IsUniqueProd) Then Sendb("<tr><td colspan=""2""><small style=""margin-left:100px;"">(" & Copient.PhraseLib.Lookup("term.disabledunique", LanguageID) & ")</small></td><td></td></tr>")
                Send("<tr id=""p3row2"">")
                Send("  <td>")
                Send("    <label for=""limit3"">" & Copient.PhraseLib.Lookup("term.limit", LanguageID) & ":</label>")
                Send("  </td>")
                Send("  <td>")
                Send("    <input type=""text"" class=""shorter"" id=""limit3"" name=""limit3"" maxlength=""6"" value=""" & MyCommon.NZ(rst.Rows(0).Item("P3DistQtyLimit"), 0) & """" & IIf((FromTemplate And Disallow_Limits), " disabled=""disabled""", "") & " />")
                Send("  </td>")
                Send("</tr>")
                Send("<tr id=""p3row3"">")
                Send("  <td>")
                Send("    <label for=""limit3period"">" & Copient.PhraseLib.Lookup("term.period", LanguageID) & ":</label>")
                Send("  </td>")
                Send("  <td>")
                Sendb("    <input type=""text"" class=""shorter"" id=""limit3period"" name=""limit3period"" maxlength=""6"" value=""" & MyCommon.NZ(rst.Rows(0).Item("P3DistPeriod"), 0))
                Send("""" & IIf(((FromTemplate And Disallow_Limits) Or (HasAnyCustomer)), " disabled=""disabled""", "") & " />")
                Send("  </td>")
                Send("</tr>")
                Send("<tr id=""p3row4"">")
                Send("  <td>")
                Send("    <label for=""P3DistTimeType"">" & Copient.PhraseLib.Lookup("term.type", LanguageID) & ":</label>")
                Send("  </td>")
                Send("  <td>")
                Send("    <select id=""P3DistTimeType"" name=""P3DistTimeType""" & IIf((FromTemplate And Disallow_Limits), " disabled=""disabled""", "") & ">")
                TempQueryStr = "select TimeTypeID,PhraseID from CPE_DistributionTimeTypes with (NoLock)"
                If HasAnyCustomer Then
                    'Restrict time type to hours if the offer has an AnyCustomer condition since we can't carry the limits beyond a single transaction
                    TempQueryStr &= " where TimeTypeID=2"
                End If
                MyCommon.QueryStr = TempQueryStr
                rst2 = MyCommon.LRT_Select
                For Each row2 In rst2.Rows
                    If (MyCommon.NZ(rst.Rows(0).Item("P3DistTimeType"), 0) = MyCommon.NZ(row2.Item("TimeTypeID"), 0)) Then
                        Send("      <option value=""" & MyCommon.NZ(row2.Item("TimeTypeID"), 0) & """ selected=""selected"">" & Copient.PhraseLib.Lookup(MyCommon.NZ(row2.Item("PhraseID"), 0), LanguageID) & "</option>")
                    Else
                        Send("      <option value=""" & MyCommon.NZ(row2.Item("TimeTypeID"), 0) & """>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row2.Item("PhraseID"), 0), LanguageID) & "</option>")
                    End If
                Next
                Send("    </select>")
                Send("    <input type=""hidden"" id=""BeginP3TimeTypeID"" name=""BeginP3TimeTypeID"" value=""" & MyCommon.NZ(rst.Rows(0).Item("P3DistTimeType"), -1) & """ />")
                Send("  </td>")
                Send("</tr>")
                Send("</table>")
            %>
            <hr class="hidden" />
        </div>
        <%
            MyCommon.QueryStr = "select ScorecardID, Description from Scorecards with (NoLock) " & _
                                "where EngineID=2 and ScorecardTypeID=4 and Deleted=0 " & _
                                "order by Description;"
            rst2 = MyCommon.LRT_Select
            If rst2.Rows.Count > 0 Then
                Send("<div class=""box"" id=""scorecard"">")
                Send("  <h2>")
                Send("    <span>")
                Send("      " & Copient.PhraseLib.Lookup("term.scorecard", LanguageID))
                Send("    </span>")
                Send("  </h2>")
                Send("  <table summary=""" & Copient.PhraseLib.Lookup("term.scorecard", LanguageID) & """>")
                Send("    <tr>")
                Send("      <td>")
                Send("        <label for=""ScorecardID"">" & Copient.PhraseLib.Lookup("term.scorecard", LanguageID) & ":</label>")
                Send("      </td>")
                Send("      <td>")
                Send("        <select class=""medium"" id=""ScorecardID"" name=""ScorecardID"" onchange=""toggleScorecardText();"">")
                Send("          <option value=""0"">" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "</option>")
                For Each row In rst2.Rows
                    Send("          <option value=""" & MyCommon.NZ(row.Item("ScorecardID"), 0) & """" & IIf(MyCommon.NZ(row.Item("ScorecardID"), 0) = ScorecardID, " selected=""selected""", "") & ">" & MyCommon.NZ(row.Item("Description"), "") & "</option>")
                Next
                Send("        </select>")
                Send("      </td>")
                Send("    </tr>")
                Send("    <tr id=""scdesc""" & IIf(ScorecardID = 0, " style=""display:none;""", "") & ">")
                Send("      <td>")
                Send("        <label for=""ScorecardDesc"">" & Copient.PhraseLib.Lookup("term.scorecardtext", LanguageID) & ":</label>")
                Send("      </td>")
                Send("      <td>")
                'Send("        <input type=""text"" class=""medium"" id=""ScorecardDesc"" name=""ScorecardDesc"" maxlength=""31"" value=""" & ScorecardDesc & """ />")
                MLI.MLColumnName = "LimitScorecardDesc"
                MLI.StandardValue = ScorecardDesc
                MLI.InputName = "ScorecardDesc"
                MLI.InputID = "ScorecardDesc"
                MLI.InputType = "text"
                MLI.LabelPhrase = ""
                MLI.MaxLength = 31
                MLI.CSSClass = ""
                MLI.CSSStyle = "width:233px;"
                Send(Localization.SendTranslationInputs(MyCommon, MLI, Request.QueryString, 2))
                Send("      </td>")
                Send("    </tr>")
                Send("  </table>")
                Send("</div>")
            End If
        %>
        <%
            If MaxTiers > 1 AndAlso Not (EngineID = 2 And EngineSubTypeID = 1) Then
                Send("<div class=""box"" id=""tiering"" style=""position:relative;"">")
                Send("  <h2>")
                Send("    <span>")
                Send("      " & Copient.PhraseLib.Lookup("term.tiers", LanguageID))
                Send("    </span>")
                Send("  </h2>")
                If IsTemplate Then
                    Send("  <span class=""temp"">")
                    Send("    <input type=""checkbox"" class=""tempcheck"" id=""Disallow_Tiers"" name=""Disallow_Tiers""" & IIf(Disallow_Tiers, " checked=""checked""", "") & " />")
                    Send("    <label for=""Disallow_Tiers"">" & Copient.PhraseLib.Lookup("term.locked", LanguageID) & "</label>")
                    Send("  </span>")
                    Send("  <br class=""printonly"" />")
                End If
                Send("  <label for=""tierlevels"">" & Copient.PhraseLib.Lookup("offer-gen.tiers", LanguageID) & " (1 " & StrConv(Copient.PhraseLib.Lookup("term.to", LanguageID), VbStrConv.Lowercase) & " " & MaxTiers & "):</label>")
                'Send("  <input type=""text"" class=""shortest"" id=""tierlevels"" name=""tierlevels"" maxlength=""2"" value=""" & TierLevels & """" & IIf((FromTemplate And Disallow_Tiers) OrElse AccumEnabled OrElse FuelEnabled, " disabled=""disabled""", "") & " /><br />")
                Send("  <input type=""text"" class=""shortest"" id=""tierlevels"" name=""tierlevels"" maxlength=""2"" value=""" & DisplayTierLevel & """" & IIf((FromTemplate And Disallow_Tiers) OrElse AccumEnabled, " disabled=""disabled""", "") & " /><br />")
                Send("  <hr class=""hidden"" />")
                Send("</div>")
            Else
                Send("<input type=""hidden"" id=""tierlevels"" name=""tierlevels"" value=""1"" />")
            End If
        %>
        <% If (ShowInboundOutboundBox) Then%>
        <div class="box" id="inboundoutbound">
            <h2>
                <span id="inoutHeader">
                    <% Sendb(Copient.PhraseLib.Lookup("term.inbound/outbound", LanguageID))%>
                </span>
            </h2>
            <% If (IsTemplate) Then%>
            <span class="temp">
                <input type="checkbox" class="tempcheck" id="Disallow_CRMEngine" name="Disallow_CRMEngine"
                    <% if(disallow_crmengine)then send(" checked=""checked""") %> />
                <label for="Disallow_CRMEngine">
                    <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
            </span>
            <br class="printonly" />
            <% End If%>
            <label for="InboundCRMEngineID" style="position: relative;">
                <% Sendb(Copient.PhraseLib.Lookup("term.creationsource", LanguageID))%>:</label>
            <br />
            <%
                If Logix.UserRoles.EditOfferSource Then
                    MyCommon.QueryStr = "select ExtInterfaceID, PhraseID, Name from ExtCRMInterfaces with (NoLock) where Deleted=0 and Active=1;"
                    rst2 = MyCommon.LRT_Select
                    If rst2.Rows.Count > 0 Then
                        Send("<select id=""InboundCRMEngineID"" name=""InboundCRMEngineID""" & IIf(FromTemplate And Disallow_CRMEngine, " disabled=""disabled""", "") & ">")
                        For Each row In rst2.Rows
                            If MyCommon.NZ(row.Item("ExtInterfaceID"), 0) = 0 Then
                                Sendb("  <option value=""0""" & IIf(InboundCRMEngineID = 0, " selected=""selected""", "") & ">")
                                Sendb(Copient.PhraseLib.Lookup("term.logix", LanguageID))
                                Send("</option>")
                            Else
                                Sendb("  <option value=""" & MyCommon.NZ(row.Item("ExtInterfaceID"), 0) & """" & IIf(MyCommon.NZ(row.Item("ExtInterfaceID"), 0) = InboundCRMEngineID, " selected=""selected""", "") & ">")
                                Sendb(Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID, MyCommon.NZ(row.Item("Name"), "")))
                                Send("</option>")
                            End If
                        Next
                        Send("</select>")
                    End If
                Else
                    If (InboundCRMEngineID > 0) Then
                        MyCommon.QueryStr = "select ExtInterfaceID, PhraseID, Name from ExtCRMInterfaces with (NoLock) where ExtInterfaceID=" & InboundCRMEngineID
                        rst2 = MyCommon.LRT_Select
                        If (rst2.Rows.Count > 0) Then
                            If MyCommon.NZ(rst2.Rows(0).Item("ExtInterfaceID"), 0) = 0 Then
                                Sendb(Copient.PhraseLib.Lookup("term.logix", LanguageID))
                            Else
                                Sendb(Copient.PhraseLib.Lookup(MyCommon.NZ(rst2.Rows(0).Item("PhraseID"), 0), LanguageID, MyCommon.NZ(rst2.Rows(0).Item("Name"), "")))
                            End If
                        Else
                            Sendb(Copient.PhraseLib.Lookup("term.unknown", LanguageID))
                        End If
                    Else
                        Sendb(Copient.PhraseLib.Lookup("term.logix", LanguageID))
                    End If
                End If
            %>
            <br />
            <br class="half" />
            <label for="crmengine" style="position: relative;">
                <% Sendb(Copient.PhraseLib.Lookup("offer-gen.sendoutbound", LanguageID))%>:</label>
            <br />
            <%
                Dim CRMEngineID As Integer = MyCommon.NZ(rst.Rows(0).Item("CRMEngineID"), -1)
                MyCommon.QueryStr = "select ExtInterfaceID, Name, PhraseID from ExtCRMInterfaces with (NoLock) " & _
                                    "where Deleted=0 and Active=1 and OutboundEnabled=1" & _
                                    IIf(TierLevels > 1, " and ExtInterfaceID not in (1, 2)", "") & _
                                    ";"
                rst2 = MyCommon.LRT_Select()
                MyCommon.QueryStr = "Select CRMEngineID from CPE_Incentives with (NoLock) where Deleted=0 and IncentiveID=" & OfferID & ";"
                crmdt = MyCommon.LRT_Select()
                If crmdt.Rows.Count > 0 Then
                    CRMEngineID = MyCommon.NZ(crmdt.Rows(0).Item("CRMEngineID"), 0)
                End If
            %>
            <select id="crmengine" name="crmengine" class="longer" <% if(FromTemplate and disallow_crmengine)then sendb(" disabled=""disabled""") %>>
                <%
                    For Each row2 In rst2.Rows
                        If IsDBNull(row2.Item("PhraseID")) Then
                            ExtName = MyCommon.NZ(row2.Item("Name"), "")
                        Else
                            ExtName = Copient.PhraseLib.Lookup(row2.Item("PhraseID"), LanguageID, MyCommon.NZ(row2.Item("Name"), ""))
                        End If
                        If (CRMEngineID = MyCommon.NZ(row2.Item("ExtInterfaceID"), 0)) Then
                            Sendb("<option value=""" & MyCommon.NZ(row2.Item("ExtInterfaceID"), 0) & """ selected=""selected"">" & ExtName & "</option>")
                        Else
                            Sendb("<option value=""" & MyCommon.NZ(row2.Item("ExtInterfaceID"), 0) & """>" & ExtName & "</option>")
                        End If
                    Next
                %>
            </select>
            <br />
            <br class="half" />
            <label for="vendor" style="position: relative;">
                <% Sendb(Copient.PhraseLib.Lookup("term.chargebackvendor", LanguageID))%>:</label>
            <br />
            <select id="vendor" name="vendor" class="longer" <% if(FromTemplate and disallow_crmengine)then sendb(" disabled=""disabled""") %>>
                <option value="0">
                    <% Sendb(Copient.PhraseLib.Lookup("term.none", LanguageID))%>
                </option>
                <%
                    MyCommon.QueryStr = "select VendorID, ExtVendorID, Name from Vendors with (NoLock) where Chargeable=1 and AnyVendor <> 1 order by ExtVendorID;"
                    rst2 = MyCommon.LRT_Select
                    For Each row2 In rst2.Rows
                        If (ChargebackVendorID = MyCommon.NZ(row2.Item("VendorID"), 0)) Then
                            Sendb("<option value=""" & MyCommon.NZ(row2.Item("VendorID"), 0) & """ selected=""selected"">" & MyCommon.NZ(row2.Item("ExtVendorID"), "") & " - " & MyCommon.NZ(row2.Item("Name"), "") & "</option>")
                        Else
                            Sendb("<option value=""" & MyCommon.NZ(row2.Item("VendorID"), 0) & """>" & MyCommon.NZ(row2.Item("ExtVendorID"), "") & " - " & MyCommon.NZ(row2.Item("Name"), "") & "</option>")
                        End If
                    Next
                %>
            </select>
            <hr class="hidden" />
        </div>
        <% End If%>
        <% If Not (HasAnyCustomer) AndAlso Logix.UserRoles.EditEmployeeFiltering Then%>
        <div class="box" id="employees">
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.employees", LanguageID))%>
                </span>
            </h2>
            <%If (IsTemplate) Then%>
            <span class="temp">
                <input type="checkbox" class="tempcheck" id="Disallow_EmployeeFiltering" name="Disallow_EmployeeFiltering"
                    <% if(disallow_employeefiltering)then send(" checked=""checked""") %> />
                <label for="Disallow_EmployeeFiltering">
                    <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
            </span>
            <br class="printonly" />
            <% End If%>
            <input type="checkbox" id="EmployeeFiltering" name="EmployeeFiltering" onclick="handleEmployeeFiltering();"
                <% if (employeefiltered)then sendb(" checked=""checked""") %><% if(FromTemplate and disallow_employeefiltering)then sendb(" disabled=""disabled""") %> />
            <label for="EmployeeFiltering">
                <% Sendb(Copient.PhraseLib.Lookup("offer-gen.empfilter", LanguageID))%></label>
            <br />
            &nbsp;&nbsp;
            <input type="radio" id="employeesonly" name="employeesonly" <% if (employeesonly) then sendb(" checked=""checked""") %>
                onclick="toggleEmployee('employeesexcluded');" <% if(FromTemplate and disallow_employeefiltering)then sendb(" disabled=""disabled""") %> />
            <label for="employeesonly">
                <% Sendb(Copient.PhraseLib.Lookup("term.employeesonly", LanguageID))%></label>
            <br />
            &nbsp;&nbsp;
            <input type="radio" id="employeesexcluded" name="employeesexcluded" style="padding-left: 5px;"
                <% if (employeesexcluded) then sendb(" checked=""checked""") %> onclick="toggleEmployee('employeesonly');"
                <% if(FromTemplate and disallow_employeefiltering)then sendb(" disabled=""disabled""") %> />
            <label for="employeesexcluded">
                <% Sendb(Copient.PhraseLib.Lookup("term.excludeemployees", LanguageID))%></label>
            <br />
            <hr class="hidden" />
        </div>
        <% End If%>
        <%		
            If MyCommon.Fetch_SystemOption(156) = "1" Then
                udflistcontrol.IsTemplate = IsTemplate
                udflistcontrol.bUseTemplateLocks = FromTemplate
                udflistcontrol.Disallow_UserDefinedFields = Disallow_UserDefinedFields
        %>
        <udf:udflist ID="udflistcontrol" runat="server" />
        <% End If%>
        <%--
      <div class="box" id="sweepstakes">
        <h2><span><% Sendb(Copient.PhraseLib.Lookup("term.instantwin", LanguageID))%></span></h2>
        <input class="checkbox" id="instantwin" name="form_InstantWin" type="checkbox" />
        <label for="instantwin"><% Sendb(Copient.PhraseLib.Lookup("term.instantwin", LanguageID))%></label>
        <br />
        <br class="half" />
        <label for="prizes"><% Sendb(Copient.PhraseLib.Lookup("offer-gen.prizesawarded", LanguageID))%></label><br />
        &nbsp;&nbsp;
        <input class="short" id="prizes" maxlength="12" name="form_NumPrizesAllowed" type="text" value="" /><br />
        <br class="half" />
        <label for="odds"><% Sendb(Copient.PhraseLib.Lookup("offer-gen.oddsofwinning", LanguageID))%></label><br />
        1:<input class="short" id="odds" name="form_OddsOfWinning" maxlength="12" type="text" value="" />
        <input id="fixed" name="form_RandomWinners" type="radio" value="0" /><label for="fixed"><% Sendb(Copient.PhraseLib.Lookup("term.fixed", LanguageID))%></label>
        <input id="random" name="form_RandomWinners" type="radio" value="1" /><label for="random"><% Sendb(Copient.PhraseLib.Lookup("term.random", LanguageID))%></label>
        <br />
        <br class="half" />
        <% Sendb(Copient.PhraseLib.Lookup("offer-gen.oddscalculation", LanguageID))%><br />
        <input id="odds-calconce" name="form_IWTransLevel" type="radio" value="0" /><label for="odds-calconce"><% Sendb(Copient.PhraseLib.Lookup("offer-gen.odds1", LanguageID))%></label><br />
        <input id="odds-calceach" name="form_IWTransLevel" type="radio" value="1" /><label for="odds-calceach"><% Sendb(Copient.PhraseLib.Lookup("offer-gen.odds2", LanguageID))%></label><br />
      </div>
        --%>
        <div class="box" id="options">
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.advancedoptions", LanguageID))%>
                </span>
            </h2>
            <%If (IsTemplate) Then%>
            <span class="temp">
                <input type="checkbox" class="tempcheck" id="Disallow_AdvancedOption" name="Disallow_AdvancedOption"
                    <% if(Disallow_AdvancedOption)then send(" checked=""checked""") %> />
                <label for="Disallow_AdvancedOption">
                    <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
                </label>
            </span>
            <br class="printonly" />
            <% End If%>
            <input type="checkbox" name="autotransferable" id="autotransferable" <%Sendb(IIf(AutoTransferable, " checked=""checked""", ""))%>
                <%Sendb(IIF((FromTemplate And Disallow_AdvancedOption)," disabled=""disabled""","")) %> />
            <label for="autotransferable">
                <% Sendb(Copient.PhraseLib.Lookup("term.autotransferable", LanguageID))%></label>
            <br />
            <% 
            %>
            <input type="checkbox" name="deferToEOS" id="deferToEOS" <% if (deferToEOS) then sendb(" checked=""checked""") %><%if (DeferToEOSDisabled or (FromTemplate and Disallow_AdvancedOption)) then sendb(" disabled=""disabled""") %> />
            <label for="deferToEOS">
                <% Sendb(Copient.PhraseLib.Lookup("term.defercalc", LanguageID))%></label>
            <br />
            <% If (MyCommon.Fetch_CPE_SystemOption(26).Trim = "1") Then%>
            <% If Not (HasAnyCustomer) Then%>
            <input type="checkbox" name="reportingimp" id="reportingimp" <% if (reportingimp) then sendb(" checked=""checked""") %>
                <%Sendb(IIF((FromTemplate And Disallow_AdvancedOption)," disabled=""disabled""","")) %> />
            <label for="reportingimp">
                <% Sendb(Copient.PhraseLib.Lookup("offer-gen.enablereporting-imp", LanguageID))%></label>
            <br />
            <% End If%>
            <input type="checkbox" name="reportingred" id="reportingred" <% if (reportingred) then sendb(" checked=""checked""") %>
                <%Sendb(IIF((FromTemplate And Disallow_AdvancedOption)," disabled=""disabled""","")) %> />
            <label for="reportingred">
                <% Sendb(Copient.PhraseLib.Lookup("offer-gen.enablereporting-red", LanguageID))%></label>
            <br />
            <% End If%>
            <% If (MyCommon.Fetch_SystemOption(73).Trim <> "") Then%>
            <input type="checkbox" name="exporttoedw" id="exporttoedw" <% if (ExportToEDW) then sendb(" checked=""checked""") %>
                <%Sendb(IIF((FromTemplate And Disallow_AdvancedOption)," disabled=""disabled""","")) %> />
            <label for="exporttoedw">
                <% Sendb(Copient.PhraseLib.Lookup("term.exporttoarchive", LanguageID))%></label>
            <br />
            <% End If%>
            <%
                If (MyCommon.Extract_Val(MyCommon.Fetch_CPE_SystemOption(70)) = 1) Then
                    'MyCommon.QueryStr = "Select Description, PhraseID from CPE_DeliverableTypes where IssuanceEnabled=1;"
                    MyCommon.QueryStr = "select RDO.PKID, RDS.Description, RDO.Enabled from RemoteDataOptions as RDO " & _
                                        "inner join RemoteDataStyles as RDS on RDO.StyleID=RDS.StyleID and RDO.RemoteDataTypeID=RDS.RemoteDataTypeID " & _
                                        "where RDO.Enabled=1 and RDO.RemoteDataTypeID=1;"
                    rst2 = MyCommon.LRT_Select()
                    If (rst2.Rows.Count > 0) Then
                        IssuanceDetails &= "<b>" & Copient.PhraseLib.Lookup("ueoffer-gen.IssuanceSent", LanguageID) & "</b><br /><ul>"
                        For Each row2 In rst2.Rows
                            IssuanceDetails &= "<li>"
                            IssuanceDetails &= MyCommon.NZ(row2.Item("Description"), "")
                            IssuanceDetails &= "</li>"
                        Next
                        IssuanceDetails &= "</ul>"
                    Else
                        IssuanceDetails &= "<br/><br/><br/><center><b>" & Copient.PhraseLib.Lookup("ueoffer-gen.IssuanceNotSent", LanguageID) & "</b></center><br/>"
                    End If
            %>
            <input type="checkbox" name="issuance" id="issuance" <% if (Issuance) then sendb(" checked=""checked""") %>
                <%Sendb(IIF((FromTemplate And Disallow_AdvancedOption)," disabled=""disabled""","")) %> />
            <label for="issuance">
                <% Sendb(Copient.PhraseLib.Lookup("offer-gen.enableissuance", LanguageID))%></label>
            <a href="#" onclick="javascript:showGrowPopup(event, '<%Sendb(IssuanceDetails)%>', 300, 200);">
                <img src="../images/info.png" alt="<% Sendb(Copient.PhraseLib.Lookup("term.details", LanguageID)) %>"
                    title="<% Sendb(Copient.PhraseLib.Lookup("term.details", LanguageID)) %>" style="position: relative;
                    top: 2px;" /></a>
            <div id="divIssuance" style="display: none;">
                Test
            </div>
            <br />
            <% End If%>
            <input type="checkbox" name="mfgCoupon" id="mfgCoupon" <% if (IsMfgCoupon) then sendb(" checked=""checked""") %>
                <%Sendb(IIF((FromTemplate And Disallow_AdvancedOption)," disabled=""disabled""","")) %> />
            <label for="mfgCoupon">
                <% Sendb(Copient.PhraseLib.Lookup("CPEoffer-gen.mfgcoupons", LanguageID))%></label>
            <br />
            <% If (MyCommon.Fetch_CPE_SystemOption(129).Trim = "1") Then%>
            <input type="checkbox" name="RestrictedRdmpt" id="RestrictedRdmpt" <% If (IsRestrictedRedemption) Then Sendb(" checked=""checked""") %>
                <%Sendb(IIF((FromTemplate And Disallow_AdvancedOption)," disabled=""disabled""","")) %> />
            <label for="RestrictedRdmpt" id="lblRestRdmpt">
                <% Sendb(Copient.PhraseLib.Lookup("term.restrictedredemption", LanguageID))%></label>
            <br />
            <% End If%>
            <% If (MyCommon.Fetch_CPE_SystemOption(163).Trim = "1") Then%>
            <input type="checkbox" name="PromptForReward" id="PromptForReward" <% If (PromptForReward) Then Sendb(" checked=""checked""") %> />
            <label for="PromptForReward" id="lblPromptForReward">
                <% Sendb(Copient.PhraseLib.Lookup("term.promptforreward", LanguageID))%></label>
            <br />
            <% End If%>
            <%
                If (Logix.UserRoles.FavoriteOffersForOthers AndAlso Not IsTemplate) Then
                    Send("<br class=""half"" />")
                    MyCommon.QueryStr = "select AdminUserID from AdminUserOffers where OfferID=" & OfferID & ";"
                    rst2 = MyCommon.LRT_Select
                    MyCommon.QueryStr = "select AdminUserID from AdminUsers;"
                    rst3 = MyCommon.LRT_Select
                    Send("<button type=""button"" id=""favorite"" name=""favorite"" value=""favorite""" & IIf((FromTemplate And Disallow_AdvancedOption), " disabled=""disabled""", "") & "onclick=""javascript:xmlhttpPost('OfferFeeds.aspx', 'FavoriteForAll');"">" & Copient.PhraseLib.Lookup("offer-gen.favoriteall", LanguageID) & "</button>")
                    Sendb("<a href=""javascript:openPopup('offer-favorite.aspx?OfferID=" & OfferID & "&bUseTemplateLocks=" & FromTemplate & "&Disallow_AdvancedOption=" & Disallow_AdvancedOption & "')""><img id=""favImg"" src=""../images/user.png"" ")
                    Sendb("alt=""" & rst2.Rows.Count & " " & StrConv(Copient.PhraseLib.Lookup("term.of", LanguageID), VbStrConv.Lowercase) & " " & rst3.Rows.Count & " " & StrConv(Copient.PhraseLib.Lookup("term.users", LanguageID), VbStrConv.Lowercase) & " " & Copient.PhraseLib.Lookup("offer-gen.havefavorite", LanguageID) & """ ")
                    Sendb("title=""" & rst2.Rows.Count & " " & StrConv(Copient.PhraseLib.Lookup("term.of", LanguageID), VbStrConv.Lowercase) & " " & rst3.Rows.Count & " " & StrConv(Copient.PhraseLib.Lookup("term.users", LanguageID), VbStrConv.Lowercase) & " " & Copient.PhraseLib.Lookup("offer-gen.havefavorite", LanguageID) & """ ")
                    Send("/></a><br />")
                End If
            %>
        </div>
        <% If (BannersEnabled AndAlso AllowMultipleBanners) Then%>
        <div class="box" id="banners">
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.banners", LanguageID))%>
                </span>
            </h2>
            <% 
                ' get the selected banners and store for later lookup
                MyCommon.QueryStr = "select BAN.BannerID from BannerOffers BO with (NoLock) " & _
                                    "inner join Banners BAN with (NoLock) on BAN.BannerID = BO.BannerID " & _
                                    "where BAN.Deleted=0 and BO.OfferID = " & OfferID
                rst2 = MyCommon.LRT_Select
                SelectedBanners = New ArrayList(rst2.Rows.Count)
                For Each row2 In rst2.Rows
                    SelectedBanners.Add(MyCommon.NZ(row2.Item("BannerID"), -1))
                    If (SelectedList <> "") Then SelectedList &= ","
                    SelectedList &= MyCommon.NZ(row2.Item("BannerID"), -1)
                Next
          
                'Send("<input type=""hidden"" name=""existingbanners"" id=""existingbanners"" value=""" & SelectedList & """ />")
                'Send("<input type=""hidden"" name=""newbanners"" id=""newbanners"" value=""" & "" & """ />")
                Send("<input type=""hidden"" name=""bannerschanged"" id=""bannerschanged"" value=""false"" />")
          
                ' get the banners for which this user is permitted to edit
                MyCommon.QueryStr = "select BAN.BannerID, BAN.Name from Banners BAN with (NoLock)" & _
                                    "inner join BannerEngines BE with (NoLock) on BE.BannerID = BAN.BannerID " & _
                                    "inner join AdminUserBanners AUB with (NoLock) on AUB.BannerID = BAN.BannerID " & _
                                    "where BE.EngineID=2 and AUB.AdminUserID =" & AdminUserID & ";"
                rst2 = MyCommon.LRT_Select
                EditableBanners = New ArrayList(rst2.Rows.Count)
                For Each row2 In rst2.Rows
                    EditableBanners.Add(MyCommon.NZ(row2.Item("BannerID"), -1))
                Next
          
                ' get all the assigned banners for CPE
                i = 0
                MyCommon.QueryStr = "select BAN.BannerID, BAN.Name from Banners BAN with (NoLock) " & _
                                    "inner join BannerEngines BE with (NoLock) on BE.BannerID = BAN.BannerID " & _
                                    "where BE.EngineID=2 and BAN.AllBanners=0;"
                rst2 = MyCommon.LRT_Select()
                For Each row2 In rst2.Rows
                    IsEditableBanner = EditableBanners.Contains(MyCommon.NZ(row2.Item("BannerID"), -1))
                    Sendb(Space(10) & "<input type=""checkbox"" name=""bannerids"" id=""bannerid" & i & """ value=""" & MyCommon.NZ(row2.Item("BannerID"), -1) & """")
                    Sendb(IIf(SelectedBanners.Contains(MyCommon.NZ(row2.Item("BannerID"), -1)), " checked=""checked""", " "))
                    Sendb(IIf(IsEditableBanner, " ", " disabled = ""disabled"""))
                    Sendb(" onClick=""handleBanners(this);""")
                    Sendb(" />")
                    Sendb("<label for=""bannerid" & i & """ title=""" & Copient.PhraseLib.Lookup(IIf(IsEditableBanner, "banners.add-to-offer-note", "banners.not-user-note"), LanguageID) & """")
                    Send(">" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row2.Item("Name"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)), 25) & "</label><br />")
                    i += 1
                Next
          
                ' get all the assigned ALL banners for CPE
                i = 0
                MyCommon.QueryStr = "select BAN.BannerID, BAN.Name from Banners BAN with (NoLock) " & _
                                    "inner join BannerEngines BE with (NoLock) on BE.BannerID = BAN.BannerID " & _
                                    "where BE.EngineID=2 and BAN.AllBanners=1;"
                rst2 = MyCommon.LRT_Select()
                If (rst2.Rows.Count > 0) Then
                    Send("<br />")
                    Send(Copient.PhraseLib.Lookup("term.allbanners", LanguageID) & ":<br />")
                    For Each row2 In rst2.Rows
                        IsEditableBanner = EditableBanners.Contains(MyCommon.NZ(row2.Item("BannerID"), -1))
                        Sendb(Space(10) & "<input type=""checkbox"" name=""bannerids"" id=""allbannerid" & i & """ value=""" & MyCommon.NZ(row2.Item("BannerID"), -1) & """")
                        Sendb(IIf(SelectedBanners.Contains(MyCommon.NZ(row2.Item("BannerID"), -1)), " checked=""checked""", " "))
                        Sendb(IIf(IsEditableBanner, " ", " disabled = ""disabled"""))
                        Sendb(" onClick=""handleAllBanners(this);""")
                        Sendb(" />")
                        Sendb("<label for=""allbannerid" & i & """ title=""" & Copient.PhraseLib.Lookup(IIf(IsEditableBanner, "banners.add-to-offer-note", "banners.not-user-note"), LanguageID) & """")
                        Send(">" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row2.Item("Name"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)), 25) & "</label><br />")
                        i += 1
                    Next
                End If
            %>
        </div>
        <% End If%>
    </div>
    <br clear="all" />
</div>
</form>
<script type="text/javascript">
<% Send_Date_Picker_Terms() %>

function datePickerClosed(targetDateField) {
  var elemProdStart = document.getElementById("productionstart");
  var elemProdEnd = document.getElementById("productionend");
  var elemTestStart = document.getElementById("testingstart");
  var elemTestEnd = document.getElementById("testingend");
      
  if (targetDateField.id == "productionstart") {
    // populate productionend, etc. if unpopulated
    if (elemProdEnd != null && elemProdEnd.value == "") {
      elemProdEnd.value = targetDateField.value;
    }
    if (elemTestStart != null && elemTestStart.value == "") {
      elemTestStart.value = targetDateField.value;
    }
    if (elemTestEnd != null && elemTestEnd.value == "") {
      elemTestEnd.value = targetDateField.value;
    }
  } else if (targetDateField.id == "productionend") {
    if (elemTestEnd != null) {
      elemTestEnd.value = targetDateField.value;
    }
  }
}






function toggleDialog(elemName, shown) {
      var elem = document.getElementById(elemName);
      var fadeElem = document.getElementById('UDFfadeDiv');
    
      if (elem != null) {
        elem.style.display = (shown) ? 'block' : 'none';
      }
      if (fadeElem != null) {
        fadeElem.style.display = (shown) ? 'block' : 'none';
      }
	  if (!shown)  {
	    document.getElementById("txtOfferUDFstringValue").value = "";
	  }
}



</script>
<script type="text/javascript">
    setPlimits();
    toggleScorecard();
</script>
<script runat="server">
    Function IsDeployableOffer(ByRef MyCommon As Copient.CommonInc, ByVal OfferID As Integer, ByVal ROID As Integer, ByRef ErrorPhrase As String) As Boolean
        Dim bDeployable As Boolean = False
    
        ErrorPhrase = ""
        bDeployable = MeetsDeploymentReqs(MyCommon, OfferID)
    
        If bDeployable Then
            bDeployable = MeetsTemplateRequirements(MyCommon, ROID)
            If (Not bDeployable) Then
                ErrorPhrase = "offer-sum.required-incomplete"
            End If
        Else
            ErrorPhrase = "cpeoffer-sum.deployalert"
        End If
    
        Return bDeployable
    End Function
  
    Function MeetsDeploymentReqs(ByRef MyCommon As Copient.CommonInc, ByVal OfferID As Integer) As Boolean
        Dim bMeetsReqs As Boolean = False
    
        ' The user wants to deploy, so do a quick check for at least one assigned offer location and terminal,
        ' and ensure that there are no unassigned tier values
        MyCommon.QueryStr = "dbo.pa_CPE_IsOfferDeployable"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
        MyCommon.LRTsp.Parameters.Add("@IsDeployable", SqlDbType.Bit).Direction = ParameterDirection.Output
        MyCommon.LRTsp.ExecuteNonQuery()
        bMeetsReqs = MyCommon.LRTsp.Parameters("@IsDeployable").Value
    
        Return bMeetsReqs
    End Function
  
    Function MeetsTemplateRequirements(ByRef MyCommon As Copient.CommonInc, ByVal ROID As Integer) As Boolean
        Dim dt As DataTable
    
        MyCommon.QueryStr = "select 'CG' as GroupType, CustomerGroupID as GroupID from CPE_IncentiveCustomerGroups with (NoLock) " & _
                            "where RewardOptionID = " & ROID & " and Deleted=0 and RequiredFromTemplate=1 and CustomerGroupID is null " & _
                            "union " & _
                            "select 'PG' as GroupType, ProductGroupID as GroupID from CPE_IncentiveProductGroups with (NoLock) " & _
                            "where RewardOptionID = " & ROID & " and Deleted=0 and RequiredFromTemplate=1 and ProductGroupID is null " & _
                            "union " & _
                            "select 'PP' as GroupType, ProgramID as GroupID from CPE_IncentivePointsGroups with (NoLock) " & _
                            "where RewardOptionID = " & ROID & " and Deleted=0 and RequiredFromTemplate=1 and ProgramId is null; "
        dt = MyCommon.LRT_Select
    
        Return (dt.Rows.Count = 0)
    End Function
</script>
<div id="UDFfadeDiv">
</div>
<div id="foldercreate" class="folderdialog" style="position: absolute; width: 400px;
    height: 150px;">
    <div class="foldertitlebar">
        <span class="dialogtitle">Enter text for the selected UDF</span> <span class="dialogclose"
            onclick="toggleDialog('foldercreate', false);">X</span>
    </div>
    <div class="dialogcontents">
        <div id="receiptmsgerror" style="display: none; color: red;">
        </div>
        <table>
            <tr>
                <td>
                    <textarea name="textarea" id="txtOfferUDFstringValue" style="width: 300px; height: 100px"></textarea>
                </td>
                <td>
                    <input type="button" name="btnpicrecmsg" id="btnpicrecmsg" value="Add" onclick="javascript:addUDFTextmessagetoOffer('foldercreate');" />
                </td>
            </tr>
        </table>
    </div>
</div>
<%
    If MyCommon.Fetch_SystemOption(75) Then
        If (OfferID > 0 And Logix.UserRoles.AccessNotes) Then
            Send_Notes(3, OfferID, AdminUserID)
        End If
    End If
done:
    If MyCommon.Fetch_SystemOption(124) = "0" Then
        Send_FocusScript("mainform", "name")
    End If
    Send_WrapEnd()
    Send("<div id=""disabledBkgrd"" style=""position:absolute;top:0px;left:0px;right:0px;width:100%;height:100%;background-color:Gray;display:none;z-index:99;filter:alpha(opacity=50);-moz-opacity:.50;opacity:.50;""></div>")
    Send_PageEnd()
    MyCommon.Close_LogixRT()
    MyCommon = Nothing
    Logix = Nothing
%>
