<%@ Page Language="vb" Debug="true" CodeFile="/logix/LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<% 
  ' *****************************************************************************
  ' * FILENAME: CAM-offer-gen.aspx 
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
  Dim rst As DataTable
  Dim rstTemplates As DataTable
  Dim rst3 As DataTable
  Dim row As DataRow
  Dim rowTemplates As DataRow
  Dim OfferID As Long = Request.QueryString("OfferID")
  Dim OfferName As String = ""
  Dim rst2 As DataTable
  Dim row2 As DataRow
  Dim IsTemplate As Boolean
  Dim IsTemplateVal As String = ""
  Dim ActiveSubTab As Integer = 205
  Dim IntroID As String = "intro"
  Dim Disallow_EmployeeFiltering As Boolean = True
  Dim Disallow_ProductionDates As Boolean = True
  Dim Disallow_Limits As Boolean = True
  Dim Disallow_Tiers As Boolean = True
  Dim Disallow_Priority As Boolean = True
  Dim Disallow_Sweepstakes As Boolean = True
  Dim Disallow_Conditions As Boolean = True
  Dim Disallow_Rewards As Boolean = True
  Dim Disallow_ExecutionEngine As Boolean = True
  Dim Disallow_CRMEngine As Boolean = True
  Dim FromTemplate As Boolean
  Dim DisabledOnCFW As Boolean
  Dim ReportingImp As Boolean = False
  Dim ReportingRed As Boolean = False
  Dim ExtOfferID As String = ""
  Dim ShowInboundOutboundBox As Boolean = True
  Dim ProdStartDate As Date
  Dim ProdEndDate As Date
  Dim StartDateParsed, EndDateParsed As Boolean
  Dim EligStartDate, EligEndDate As Date
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
  Dim EngineID As Integer = 0
  Dim EnginePhraseID As Integer = 0
  Dim EngineSubTypeID As Integer = 0
  Dim EngineSubTypePhraseID As Integer = 0
  Dim CheckedLMG As Boolean = False
  Dim ExtName As String = ""
  Dim SelectedList As String = ""
  Dim TierLevels As Integer = 1
  Dim MaxTiers As Integer = 1
  Dim IsUniqueProd As Boolean = False
  Dim ShortStartDate, ShortEndDate As String
  Dim ShortTestStartDate, ShortTestEndDate As String
  Dim StartDT, EndDT As Date
  Dim TestStartDT, TestEndDT As Date
  Dim DescriptLength As Boolean = False
  Dim Description As String = ""
  Dim HeaderExists As Boolean = False
  Dim DollarProductCondition As Boolean = False
  Dim DisplayTierLevel As String = ""
  Dim IsFootworthy As Boolean = False
  Dim IsDeployable As Boolean = False
  Dim ErrorPhrase As String = ""
  Dim MutuallyExclusive As Boolean = False
  Dim AutoTransferable As Boolean = False
  Dim selectDatePicker As Integer = MyCommon.Extract_Val(MyCommon.NZ(MyCommon.Fetch_SystemOption(161), 0))
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "CAM-offer-gen.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  AllowMultipleBanners = (MyCommon.Fetch_SystemOption(67) = "1")
  Popup = IIf((Request.QueryString("Popup") <> "") AndAlso (Request.QueryString("Popup") <> "0"), True, False)
  
  MyCommon.QueryStr = "select RewardOptionID, TierLevels from CPE_RewardOptions with (NoLock) where TouchResponse=0 and IncentiveID=" & OfferID
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    roid = rst.Rows(0).Item("RewardOptionID")
    TierLevels = rst.Rows(0).Item("TierLevels")
  End If
  If Request.QueryString("tierlevels") <> "" Then
    DisplayTierLevel = MyCommon.Extract_Val(Request.QueryString("tierlevels"))
  Else
    DisplayTierLevel = TierLevels
  End If
  MaxTiers = MyCommon.Fetch_SystemOption(89)
 
  'Set the favorite boolean
  If OfferID > 0 Then
    MyCommon.QueryStr = "Select Favorite from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
      Favorite = MyCommon.NZ(rst.Rows(0).Item("Favorite"), False)
    End If
  End If
  
  'Find if there are any unique product flags for this roid
  MyCommon.QueryStr = "select UniqueProduct from CPE_IncentiveProductGroups where RewardOptionID=" & roid & " and UniqueProduct=1 and Deleted=0"
  rst = MyCommon.LRT_Select
  If (rst.Rows.Count > 0) Then IsUniqueProd = True
  
  'See if this offer is footworthy -- ie, can be allowed to have a footer priority
  MyCommon.QueryStr = "dbo.pa_CPE_IsFootworthy"
  MyCommon.Open_LRTsp()
  MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
  MyCommon.LRTsp.Parameters.Add("@ROID", SqlDbType.BigInt).Value = roid
  MyCommon.LRTsp.Parameters.Add("@IsFootworthy", SqlDbType.Bit).Direction = ParameterDirection.Output
  MyCommon.LRTsp.ExecuteNonQuery()
  IsFootworthy = MyCommon.LRTsp.Parameters("@IsFootworthy").Value
  MyCommon.Close_LRTsp()
  
  'Save
  If (Request.QueryString("save") <> "") Then
    
    'Check the Description length
    If Request.QueryString("form_description") <> "" Then
      Description = MyCommon.Parse_Quotes(Request.QueryString("form_description"))
      If Description.Length <= 1000 Then
        DescriptLength = True
      End If
    Else
      DescriptLength = True
    End If
    
    'Run query to check for mgfcoupon/discount compatibility
    MyCommon.QueryStr = "select DI.DiscountID, DI.DiscountTypeID, RO.RewardOptionID, RO.IncentiveID, I.ManufacturerCoupon " & _
                        "from CPE_Discounts as DI with (NoLock) " & _
                        "inner join CPE_Deliverables as DE with (NoLock) on DE.OutputID=DI.DiscountID " & _
                        "inner join CPE_RewardOptions as RO with (NoLock) on DE.RewardOptionID=RO.RewardOptionID " & _
                        "inner join CPE_Incentives as I with (NoLock) on I.IncentiveID=RO.IncentiveID " & _
                        "where I.IncentiveID=" & OfferID & " and DI.Deleted=0 and DE.Deleted=0;"
    rst = MyCommon.LRT_Select
    ' Also, run a query to see if there's a category that has this offer as its base offer
    MyCommon.QueryStr = "select OfferCategoryID from OfferCategories where Deleted=0 and BaseOfferID=" & OfferID & " and OfferCategoryID=(" & _
                        "  select IsNull(PromoClassID, 0) from CPE_Incentives where IncentiveID=" & OfferID & ");"
    rst2 = MyCommon.LRT_Select
    If (rst.Rows.Count > 0 AndAlso MyCommon.NZ(rst.Rows(0).Item("DiscountTypeID"), 0) <> 1 AndAlso Request.QueryString("mfgCoupon") = "on") Then
      infoMessage = Copient.PhraseLib.Lookup("cam-offer-gen.InvalidMfgCoupon", LanguageID)
    ElseIf (Request.QueryString("productionstart") = "" Or Request.QueryString("productionend") = "") Then
      infoMessage = Copient.PhraseLib.Lookup("CPEoffer_gen.noproductiondates", LanguageID)
    ElseIf (MyCommon.Extract_Val(Request.QueryString("priority")) = 900) AndAlso (MyCommon.Extract_Val(Request.QueryString("footerpriority")) < 0 OrElse MyCommon.Extract_Val(Request.QueryString("footerpriority")) > 99) Then
      infoMessage = Copient.PhraseLib.Lookup("CPEoffer_gen.badpriority", LanguageID)
    ElseIf DescriptLength = False Then
      infoMessage = Copient.PhraseLib.Lookup("error.description", LanguageID)
    ElseIf ((rst2.Rows.Count > 0) AndAlso (MyCommon.Extract_Val(Request.QueryString("form_Category")) <> MyCommon.NZ(rst2.Rows(0).Item("OfferCategoryID"), 0))) Then
      infoMessage = Copient.PhraseLib.Lookup("ueoffer-gen.InvalidCategoryChange", LanguageID)
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
                            "where Deleted=0 and IncentiveName='" & MyCommon.Parse_Quotes(Request.QueryString("form_name")) & "' and IncentiveID<>" & Request.QueryString("OfferID") & _
                            " union all " & _
                            "select Name from Offers with (NoLock) " & _
                            "where Deleted=0 and Name='" & MyCommon.Parse_Quotes(Request.QueryString("form_name")) & "' and OfferID<>" & Request.QueryString("OfferID") & ";"
        rst = MyCommon.LRT_Select
        DuplicateName = (rst.Rows.Count > 0)
        
        sqlBuf.Append("Update CPE_Incentives with (RowLock) set ")
        sqlBuf.Append("IncentiveName=N'" & MyCommon.Parse_Quotes(Logix.TrimAll(Request.QueryString("form_name"))) & "',")
        sqlBuf.Append("Description=N'" & MyCommon.Parse_Quotes(Request.QueryString("form_description")) & "',")
        If (Request.QueryString("form_Category") <> "") Then
          sqlBuf.Append("PromoClassID=" & Request.QueryString("form_Category") & ",")
        End If
        If (Request.QueryString("priority") <> "") Then
          If (MyCommon.Extract_Val(Request.QueryString("priority")) = 900) Then
            sqlBuf.Append("Priority=" & MyCommon.Extract_Val(Request.QueryString("footerpriority")) + 900 & ",")
          Else
            sqlBuf.Append("Priority=" & Request.QueryString("priority") & ",")
          End If
        End If
        If (Request.QueryString("crmengine") <> "") Then
          sqlBuf.Append("CRMEngineID=" & Request.QueryString("crmengine") & ",")
        End If
        If (IsUniqueProd) Then
          sqlBuf.Append("P3DistQtyLimit=1,")
        ElseIf (Request.QueryString("limit3") <> "") Then
          sqlBuf.Append("P3DistQtyLimit=" & Request.QueryString("limit3") & ",")
        End If
        If (IsUniqueProd) Then
          sqlBuf.Append("P3DistPeriod=1,")
        ElseIf (Request.QueryString("limit3period") <> "") Then
          sqlBuf.Append("P3DistPeriod=" & Request.QueryString("limit3period") & ",")
        End If
        If (IsUniqueProd) Then
          sqlBuf.Append("P3DistTimeType=2,")
        ElseIf (Request.QueryString("P3DistTimeType") <> "") Then
          sqlBuf.Append("P3DistTimeType=" & Request.QueryString("P3DistTimeType") & ",")
        End If
        sqlBuf.Append("StartDate='" & ProdStartDate.ToShortDateString() & "',")
        sqlBuf.Append("EndDate='" & ProdEndDate.ToShortDateString() & "',")
        sqlBuf.Append("TestingStartDate='" & TestStartDate.ToShortDateString() & "',")
        sqlBuf.Append("TestingEndDate='" & TestEndDate.ToShortDateString() & "',")
        sqlBuf.Append("DisabledOnCFW=" & IIf(Request.QueryString("DisabledOnCFW") = "on", 1, 0) & ",")
        sqlBuf.Append("EnableImpressRpt=" & IIf(Request.QueryString("reportingimp") = "on", 1, 0) & ",")
        sqlBuf.Append("EnableRedeemRpt=" & IIf(Request.QueryString("reportingred") = "on", 1, 0) & ",")
        If (MyCommon.Extract_Val(Request.QueryString("priority")) >= 900) Then
          sqlBuf.Append("DeferCalcToEOS=1,")
        Else
          sqlBuf.Append("DeferCalcToEOS=" & IIf(Request.QueryString("deferToEOS") = "on", 1, 0) & ",")
        End If
        sqlBuf.Append("ExportToEDW=" & IIf(Request.QueryString("exporttoedw") = "on", 1, 0) & ",")
        sqlBuf.Append("Favorite=" & IIf(Request.QueryString("favorite") = "on", 1, 0) & ",")
        If Request.QueryString("InboundCRMEngineID") <> "" Then
          sqlBuf.Append("InboundCRMEngineID=" & Request.QueryString("InboundCRMEngineID") & ",")
        End If
        sqlBuf.Append("SendIssuance=" & IIf(Request.QueryString("issuance") = "on", 1, 0) & ",")
        sqlBuf.Append("ChargebackVendorID=" & MyCommon.Extract_Val(Request.QueryString("vendor")) & ",")
        sqlBuf.Append("ManufacturerCoupon=" & IIf(Request.QueryString("mfgCoupon") = "on", 1, 0) & ",")
        sqlBuf.Append("AutoTransferable=" & IIf(Request.QueryString("autotransferable") = "on", 1, 0) & ",")
        sqlBuf.Append("LMGRegistered=" & IIf(Request.QueryString("LMGReg") = "on", 1, 0) & ", ")
        sqlBuf.Append("MutuallyExclusive=" & IIf(Request.QueryString("MutEx") = "on", 1, 0) & ", ")
        sqlBuf.Append("LastUpdate=getdate(), ")
        'Employee Filtering - Disabled and set to 0 for both
        'sqlBuf.Append("EmployeesOnly=" & IIf(Request.QueryString("employeesonly") = "on", 1, 0) & ", ")
        'sqlBuf.Append("EmployeesExcluded=" & IIf(Request.QueryString("employeesexcluded") = "on", 1, 0) & ", ")
        sqlBuf.Append("LastUpdatedByAdminID=" & AdminUserID & ", ")
        sqlBuf.Append("StatusFlag=1 ")
        sqlBuf.Append("where IncentiveID=" & Request.QueryString("OfferID"))
        MyCommon.QueryStr = sqlBuf.ToString
        
        'Send(MyCommon.QueryStr)
        If (ProdEndDate < ProdStartDate) Then
          infoMessage = Copient.PhraseLib.Lookup("offer-gen.baddate", LanguageID)
        ElseIf (EligEndDate < EligStartDate) Then
          infoMessage = Copient.PhraseLib.Lookup("offer-gen.baddate", LanguageID)
        ElseIf (TestEndDate < TestStartDate) Then
          infoMessage = Copient.PhraseLib.Lookup("offer-gen.baddate", LanguageID)
        ElseIf DuplicateName Then
          infoMessage = Copient.PhraseLib.Lookup("offer-gen.nameused", LanguageID)
        ElseIf Logix.TrimAll(Request.QueryString("form_name")) = "" Then
          infoMessage = Copient.PhraseLib.Lookup("offer-gen.noname", LanguageID)
        ElseIf (Request.QueryString("limit3") <> "" AndAlso (Not Integer.TryParse(Request.QueryString("limit3"), TempInt) OrElse (TempInt < 0 AndAlso MyCommon.Extract_Val(Request.QueryString("p3type")) > 0) OrElse (TempInt <= 0 AndAlso MyCommon.Extract_Val(Request.QueryString("p3type")) = 0))) Then
          infoMessage = Copient.PhraseLib.Lookup("offer-gen.badlimit", LanguageID)
        ElseIf (Request.QueryString("limit3period") <> "" AndAlso (Not Integer.TryParse(Request.QueryString("limit3period"), TempInt) OrElse (TempInt < 0 AndAlso MyCommon.Extract_Val(Request.QueryString("p3type")) > 0) OrElse (TempInt <= 0 AndAlso MyCommon.Extract_Val(Request.QueryString("p3type")) = 0))) Then
          infoMessage = Copient.PhraseLib.Lookup("offer-gen.badlimit", LanguageID)
        ElseIf (Request.QueryString("tierlevels") <> "" AndAlso (Not Integer.TryParse(Request.QueryString("tierlevels"), TempInt) OrElse TempInt < 1 OrElse TempInt > MaxTiers)) Then
          infoMessage = Copient.PhraseLib.Detokenize("offer-gen.invalidtierscpe", LanguageID, MaxTiers)
        Else
          MyCommon.LRT_Execute()
          
          ' if TierLevels has changed, update the value
          If MyCommon.Extract_Val(Request.QueryString("tierlevels")) <> TierLevels Then
            MyCommon.QueryStr = "update CPE_RewardOptions set TierLevels=" & MyCommon.Extract_Val(Request.QueryString("tierlevels")) & " " & _
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
          
          ' when an offer is flagged as a manufacturer coupon offer, best deal should be disabled.
          If Request.QueryString("mfgCoupon") = "on" Then
            MyCommon.QueryStr = "select DISC.DiscountID from CPE_Discounts DISC with (NoLock) " & _
                                "inner join CPE_Deliverables DEL with (NoLock) on DEL.OutputID = DISC.DiscountID and DEL.DeliverableTypeID=2 " & _
                                "   and DEL.RewardOptionPhase=3 and DEL.Deleted=0 " & _
                                "inner join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID = DEL.RewardOptionId and RO.Deleted=0 " & _
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
        
        IsTemplate = (Request.QueryString("IsTemplate") = "IsTemplate")
        If (IsTemplate) Then
          'Update template permissions
          Dim form_Disallow_ExecutionEngine As Integer = IIf(Request.QueryString("Disallow_ExecutionEngine") = "on", 1, 0)
          Dim form_Disallow_CRMEngine As Integer = IIf(Request.QueryString("Disallow_CRMEngine") = "on", 1, 0)
          'Employee filtering not allowed in CAM offers
          Dim form_Disallow_EmployeeFiltering As Integer = 1
          Dim form_Disallow_ProductionDates As Integer = IIf(Request.QueryString("Disallow_ProductionDates") = "on", 1, 0)
          Dim form_Disallow_Limits As Integer = IIf(Request.QueryString("Disallow_Limits") = "on", 1, 0)
          Dim form_Disallow_Tiers As Integer = IIf(Request.QueryString("Disallow_Tiers") = "on", 1, 0)
          Dim form_Disallow_Priority As Integer = IIf(Request.QueryString("Disallow_Priority") = "on", 1, 0)
          Dim form_Disallow_Sweepstakes As Integer = IIf(Request.QueryString("Disallow_Sweepstakes") = "on", 1, 0)
          MyCommon.QueryStr = "update TemplatePermissions with (RowLock) set Disallow_EmployeeFiltering=" & form_Disallow_EmployeeFiltering & _
                              " , Disallow_ProductionDates=" & form_Disallow_ProductionDates & _
                              " , Disallow_limits=" & form_Disallow_Limits & _
                              " , Disallow_Tiers=" & form_Disallow_Tiers & _
                              " , Disallow_Priority=" & form_Disallow_Priority & _
                              " , Disallow_CRMEngine=" & form_Disallow_CRMEngine & _
                              " , Disallow_ExecutionEngine=" & form_Disallow_ExecutionEngine & _
                              " , Disallow_Sweepstakes=" & form_Disallow_Sweepstakes & " where OfferID=" & OfferID
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
      Else
        infoMessage = Copient.PhraseLib.Lookup("CPEoffer_gen.noproductiondates", LanguageID)
      End If
    End If
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
  
  If (Request.QueryString("OfferID") <> "") Then
    MyCommon.QueryStr = "Select IncentiveID, OID.EngineID, PE.PhraseID as EnginePhraseID, PEST.PhraseID as EngineSubTypePhraseID, " & _
                        "IsTemplate, FromTemplate, ClientOfferID, IncentiveName, CPE.Description, PromoClassID, CRMEngineID, Priority, " & _
                        "StartDate, EndDate, EveryDOW, EligibilityStartDate, EligibilityEndDate, TestingStartDate, TestingEndDate, P1DistQtyLimit, P1DistTimeType, P1DistPeriod, " & _
                        "P2DistQtyLimit, P2DistTimeType, P2DistPeriod, P3DistQtyLimit, P3DistTimeType, P3DistPeriod, EnableImpressRpt, EnableRedeemRpt, CPE.CreatedDate, CPE.LastUpdate, CPE.Deleted, CPEOARptDate, " & _
                        "CPEOADeploySuccessDate, CPEOADeployRpt, CRMRestricted, StatusFlag, DisabledOnCFW, DisplayOnWebKiosk, EmployeesOnly, EmployeesExcluded, DeferCalcToEOS, ExportToEDW, " & _
                        "Favorite, OC.Description as CategoryName, SendIssuance, InboundCRMEngineID, ChargebackVendorID, ManufacturerCoupon, AutoTransferable, LMGRegistered, CPE.MutuallyExclusive, CPE.EngineSubTypeID " & _
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
      DisabledOnCFW = MyCommon.NZ(rst.Rows(0).Item("DisabledOnCFW"), False)
      ReportingImp = MyCommon.NZ(rst.Rows(0).Item("EnableImpressRpt"), False)
      ReportingRed = MyCommon.NZ(rst.Rows(0).Item("EnableRedeemRpt"), False)
      ExtOfferID = MyCommon.NZ(rst.Rows(0).Item("ClientOfferID"), "")
      DeferToEOS = MyCommon.NZ(rst.Rows(0).Item("DeferCalcToEOS"), False)
      ExportToEDW = MyCommon.NZ(rst.Rows(0).Item("ExportToEDW"), False)
      Favorite = MyCommon.NZ(rst.Rows(0).Item("Favorite"), False)
      Issuance = (MyCommon.NZ(rst.Rows(0).Item("SendIssuance"), 0) = 1)
      InboundCRMEngineID = MyCommon.NZ(rst.Rows(0).Item("InboundCRMEngineID"), 0)
      ChargebackVendorID = MyCommon.NZ(rst.Rows(0).Item("ChargebackVendorID"), 0)
      IsMfgCoupon = (MyCommon.NZ(rst.Rows(0).Item("ManufacturerCoupon"), 0) = 1)
      EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), 0)
      EnginePhraseID = MyCommon.NZ(rst.Rows(0).Item("EnginePhraseID"), 0)
      EngineSubTypeID = MyCommon.NZ(rst.Rows(0).Item("EngineSubTypeID"), 0)
      EngineSubTypePhraseID = MyCommon.NZ(rst.Rows(0).Item("EngineSubTypePhraseID"), 0)
      CheckedLMG = MyCommon.NZ(rst.Rows(0).Item("LMGRegistered"), 0)
      MutuallyExclusive = MyCommon.NZ(rst.Rows(0).Item("MutuallyExclusive"), 0)
      AutoTransferable = MyCommon.NZ(rst.Rows(0).Item("AutoTransferable"), False)
      'Employees filtering not allowed for CAM offers
      'EmployeesExcluded = MyCommon.NZ(rst.Rows(0).Item("EmployeesExcluded"), False) 
      'EmployeesOnly = MyCommon.NZ(rst.Rows(0).Item("EmployeesOnly"), False)
      'EmployeeFiltered = EmployeesOnly Or EmployeesExcluded 
    End If
    
    ' Ensure that defer to end of sale is still valid
    MyCommon.QueryStr = "select DeliverableID, I.DeferCalcToEOS from CPE_Deliverables D with (NoLock) " & _
                        "inner join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID=D.RewardOptionID and RO.Deleted=0 and D.Deleted=0 " & _
                        "inner join CPE_Incentives I with (NoLock) on I.IncentiveID=RO.IncentiveID and I.Deleted=0 " & _
                        "where I.IncentiveID = " & OfferID & " and DeliverableTypeID = 2;"
    rst2 = MyCommon.LRT_Select
    If (rst2.Rows.Count > 0) Then
      If (MyCommon.NZ(rst2.Rows(0).Item("DeferCalcToEOS"), False)) Then
        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set DeferCalcToEOS=0 where IncentiveID=" & OfferID
        MyCommon.LRT_Execute()
      End If
      DeferToEOSDisabled = True
      DeferToEOS = True
    End If
    ' If this is a footer priority offer, turn on the DeferCalcToEOS and disable the control
    If MyCommon.NZ(rst.Rows(0).Item("Priority"), 0) >= 900 Then
      DeferToEOSDisabled = True
      DeferToEOS = True
    End If
    
    If (IsTemplate Or FromTemplate) Then
      ' lets dig the permissions if its a template
      MyCommon.QueryStr = "select * from templatepermissions with (NoLock) where OfferID=" & OfferID
      rstTemplates = MyCommon.LRT_Select
      If (rstTemplates.Rows.Count > 0) Then
        For Each rowTemplates In rstTemplates.Rows
          ' ok there are some rows for the template
          'Employees filtering not allowed for CAM offers
          Disallow_EmployeeFiltering = False
          Disallow_ProductionDates = MyCommon.NZ(rowTemplates.Item("Disallow_ProductionDates"), True)
          Disallow_Limits = MyCommon.NZ(rowTemplates.Item("Disallow_Limits"), True)
          Disallow_Tiers = MyCommon.NZ(rowTemplates.Item("Disallow_Tiers"), True)
          Disallow_Priority = MyCommon.NZ(rowTemplates.Item("Disallow_Priority"), True)
          Disallow_Sweepstakes = MyCommon.NZ(rowTemplates.Item("Disallow_Sweepstakes"), True)
          Disallow_Conditions = MyCommon.NZ(rowTemplates.Item("Disallow_Conditions"), True)
          Disallow_Rewards = MyCommon.NZ(rowTemplates.Item("Disallow_Rewards"), True)
          Disallow_ExecutionEngine = MyCommon.NZ(rowTemplates.Item("Disallow_ExecutionEngine"), True)
          Disallow_CRMEngine = MyCommon.NZ(rowTemplates.Item("Disallow_CRMEngine"), True)
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
      End If
    End If
  End If
  
  StatusText = Logix.GetOfferStatus(OfferID, LanguageID, StatusCode)
  ShowInboundOutboundBox = (MyCommon.Fetch_SystemOption(25) <> "0")
  
  If (IsTemplate) Then
    ActiveSubTab = 206
    IntroID = "intro"
    IsTemplateVal = "IsTemplate"
  Else
    ActiveSubTab = 205
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
<script type="text/javascript">
window.name = "CAMofferGen"
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
      && el.id!="testing-start-picker" && el.id!="testing-end-picker") {
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

//function handleEmployeeFiltering() {
//  var elemFilter = document.getElementById("EmployeeFiltering");
//  var elemOnly = document.getElementById("employeesonly");
//  var elemExcluded = document.getElementById("employeesexcluded");
//  
//  if (elemFilter != null && !elemFilter.checked) {
//    if (elemOnly != null) {
//      elemOnly.checked = false;
//    }
//    if (elemExcluded != null) {
//      elemExcluded.checked = false;
//    }
//  }
//  
//  if ( (elemOnly!=null && elemOnly.checked) || (elemExcluded!=null && elemExcluded.checked) ) {
//    if (elemFilter != null) {
//      elemFilter.checked = true;
//    }
//  }
//}

//function toggleEmployee(elemName) {
//  var elemFilter = document.getElementById("EmployeeFiltering");
//  var elemOnly = document.getElementById("employeesonly");
//  var elemExcluded = document.getElementById("employeesexcluded");
//  
//  if( document.getElementById(elemName).checked==true){
//    document.getElementById(elemName).checked=false;
//  }
//  if ( (elemOnly!=null && elemOnly.checked) || (elemExcluded!=null && elemExcluded.checked) ) {
//    if (elemFilter != null) {
//      elemFilter.checked = true;
//    }
//  }
//}

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
    retVal = isDate(elemEnd.value);
    if (retVal) {
      dtEnd = new Date(Date.parse(elemEnd.value));
      dtEnd.setDate(dtEnd.getDate() + 1);
      if (dtEnd < dtNow) {
        retVal = confirm('<%Sendb(Copient.PhraseLib.Lookup("term.expire-confirm", LanguageID)) %>');
        <%
          If MyCommon.Fetch_CPE_SystemOption(80) = 1 Then
            Send("if (retVal = true) {")
            Send("  document.getElementById(""Deploy"").value = 1;")
            Send("}")
          End If
        %>
      }
    }
  }
  return retVal;
}

function xmlhttpPost(strURL, mode) {
  var self = this;
  
  //document.getElementById("tools").innerHTML = "<div class=\"loading\"><img id=\"clock\" src=\"../../images/clock22.png\" \/><br \/>" + '<% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID)) %><\/div>';
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
</script>
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
  
  If (BannersEnabled AndAlso Not Logix.IsAccessibleOffer(AdminUserID, OfferID)) Then
    Send("<script type=""text/javascript"" language=""javascript"">")
    Send("  function updateCookie() { return true; } ")
    Send("</script>")
    Send_Denied(1, "banners.access-denied-offer")
    Send_BodyEnd()
    GoTo done
  End If
%>
<form id="mainform" name="mainform" action="CAM-offer-gen.aspx" method="get" onsubmit="handleFormElements(this, false);return promptForDeploy();">
  <input type="hidden" name="OfferID" id="OfferID" value="<%Sendb(OfferID)%>" />
  <input type="hidden" name="IsActive" id="IsActive" value="<%Sendb(IIf(StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_ACTIVE, "true", "false")) %>" />
  <input type="hidden" name="IsTemplate" id="IsTemplate" value="<% Sendb(IsTemplateVal)%>" />
  <input type="hidden" name="Popup" id="Popup" value="<%Sendb(IIf(Popup, 1, 0)) %>" />
  <input type="hidden" name="Deploy" id="Deploy" value="" />
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
        If Not (IsTemplate) Then
          If (Logix.UserRoles.EditOffer) Then
            Send_Save()
          End If
        Else
          If (Logix.UserRoles.EditTemplates) Then
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
      <div class="box" id="identification">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.identification", LanguageID))%>
          </span>
        </h2>
        <%
          Send(Copient.PhraseLib.Lookup("term.id", LanguageID) & ": " & OfferID & "<br />")
          Send(Copient.PhraseLib.Lookup("term.externalid", LanguageID) & ": " & ExtOfferID & "<br />")
          Send(Copient.PhraseLib.Lookup("term.roid", LanguageID) & ": " & roid & "<br />")
          Send(Copient.PhraseLib.Lookup("term.engine", LanguageID) & ": " & Copient.PhraseLib.Lookup(EnginePhraseID, LanguageID) & IIf(EngineSubTypePhraseID > 0, " " & Copient.PhraseLib.Lookup(EngineSubTypePhraseID, LanguageID), "") & "<br />")
          Send(Copient.PhraseLib.Lookup("term.status", LanguageID) & ": " & StatusText & "<br />")
        %>
        <br class="half" />
        <label for="name"><% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>:</label><br />
        <input class="longest" style="width: 95%;" id="name" name="form_Name" maxlength="100" type="text" value="<% sendb(OfferName.Replace("""", "&quot;")) %>" /><br />
        <br class="half" />
        <label for="desc"><% Sendb(Copient.PhraseLib.Lookup("term.description", LanguageID))%>:</label><br />
        <textarea class="longest" style="width: 95%;" cols="48" rows="3" id="desc" name="form_Description" maxlength="1000"><% Sendb(MyCommon.NZ(rst.Rows(0).Item("Description"), ""))%></textarea><br />
        <small><%Send("(" & Copient.PhraseLib.Lookup("CPEoffergen.description", LanguageID) & ")")%></small><br class="half" /><br class="half" />
        <br class="half" />
        <label for="category"><% Sendb(Copient.PhraseLib.Lookup("term.category", LanguageID))%>:</label><br />
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
          <input type="checkbox" class="tempcheck" id="Disallow_Priority" name="Disallow_Priority"<% if(disallow_priority)then send(" checked=""checked""") %> />
          <label for="Disallow_Priority"><% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
        </span>
        <br class="printonly" />
        <% End If%>
        <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.priority", LanguageID)) %>">
          <tr>
            <td style="width:120px;">
              <label for="priority"><% Sendb(Copient.PhraseLib.Lookup("CPEoffer-gen.priority", LanguageID) & ":")%></label>
            </td>
            <td>
              <select id="priority" name="priority" onchange="toggleFooterPriority();"<% if(FromTemplate and disallow_priority)then sendb(" disabled=""disabled""") %>>
                <%
                  MyCommon.QueryStr = "select * from CPE_Incentives with (NoLock) where Priority=-1;"
                  rst2 = MyCommon.LRT_Select
                  If rst2.Rows.Count > 0 Then
                    HeaderExists = True
                  End If
                  MyCommon.QueryStr = "select PriorityID, Description, PhraseID from CPE_IncentivePriorities with (NoLock) "
                  If HeaderExists Then
                    MyCommon.QueryStr &= "where PriorityID>-1"
                  End If
                  If IsFootworthy Then
                    MyCommon.QueryStr &= ";"
                  Else
                    MyCommon.QueryStr &= IIf(HeaderExists, " and ", " where ") & "  PriorityID<900;"
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
              <label id="footerprioritylabel" for="footerpriority"<% Sendb(IIf(MyCommon.NZ(rst.Rows(0).Item("Priority"), 2) < 900, " style=""display:none;""", "")) %>><% Sendb(Copient.PhraseLib.Lookup("term.footer", LanguageID) & " " & Copient.PhraseLib.Lookup("term.priority", LanguageID) & ":")%></label>
            </td>
            <td>
              <input type="text" class="shortest" id="footerpriority" name="footerpriority" maxlength="2"<% Sendb(IIf(MyCommon.NZ(rst.Rows(0).Item("Priority"), 2) < 900, " style=""display:none;""", "")) %> value="<% Sendb(MyCommon.NZ(rst.Rows(0).Item("Priority"), 0) - 900) %>" /> <span id="footerprioritynote"<% Sendb(IIf(MyCommon.NZ(rst.Rows(0).Item("Priority"), 2) < 900, " style=""display:none;""", "")) %>><% Sendb(Copient.PhraseLib.Lookup("CPEoffer-gen.footernote", LanguageID))%></span>
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
          <input type="checkbox" class="tempcheck" id="Disallow_ProductionDates" name="Disallow_ProductionDates"<% if(disallow_productiondates)then send(" checked=""checked""") %> />
          <label for="Disallow_ProductionDates"><% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
        </span>
        <br class="printonly" />
        <% End If%>
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
              <label for="productionstart"><% Sendb(Copient.PhraseLib.Lookup("term.production", LanguageID))%>:</label><br />
              <input type="text" class="short" id="productionstart" name="productionstart" maxlength="10" value="<% sendb(ShortStartDate) %>"<% if(FromTemplate and disallow_productiondates)then sendb(" disabled=""disabled""") %> />
              <img src="../../images/calendar.png" class="calendar" id="production-start-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('productionstart', event);" />
              <% Sendb(Copient.PhraseLib.Lookup("term.to", LanguageID))%>
              <input type="text" class="short" id="productionend" name="productionend" maxlength="10" value="<% sendb(ShortEndDate) %>"<% if(FromTemplate and disallow_productiondates)then sendb(" disabled=""disabled""") %> />
              <img src="../../images/calendar.png" class="calendar" id="production-end-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="autoDatePicker('productionend', event, <% sendb(selectDatePicker) %>);" />
              (<% Sendb(Copient.PhraseLib.Lookup("term.mmddyyyy", LanguageID))%>)<br />
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
              <label for="testingstart"><% Sendb(Copient.PhraseLib.Lookup("term.testing", LanguageID))%>:</label><br />
              <input type="text" class="short" id="testingstart" name="testingstart" maxlength="10" value="<% sendb(ShortTestStartDate) %>"<% if(FromTemplate and disallow_productiondates)then sendb(" disabled=""disabled""") %> />
              <img src="../../images/calendar.png" class="calendar" id="testing-start-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('testingstart', event);" />
              <% Sendb(Copient.PhraseLib.Lookup("term.to", LanguageID))%>
              <input type="text" class="short" id="testingend" name="testingend" maxlength="10" value="<% sendb(ShortTestEndDate) %>"<% if(FromTemplate and disallow_productiondates)then sendb(" disabled=""disabled""") %> />
              <img src="../../images/calendar.png" class="calendar" id="testing-end-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('testingend', event);" />
              (<% Sendb(Copient.PhraseLib.Lookup("term.mmddyyyy", LanguageID))%>)<br />
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
          <input type="checkbox" class="tempcheck" id="Disallow_Limits" name="Disallow_Limits"<% if(disallow_limits)then send(" checked=""checked""") %> />
          <label for="Disallow_Limits"><% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
        </span>
        <br class="printonly" />
        <% End If%>
        <label for="limit3" style="position: relative;"><b><% Sendb(Copient.PhraseLib.Lookup("term.reward", LanguageID))%>:</b></label>
        <br />
        <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.rewards", LanguageID))%>">
          <tr>
            <td>
              <label for="p3type"><% Sendb(Copient.PhraseLib.Lookup("term.frequency", LanguageID))%>:</label>
            </td>
            <td>
              <select name="p3type" id="p3type" onchange="updateP3limit();"<% if(FromTemplate and disallow_limits or IsUniqueProd)then sendb(" disabled=""disabled""") %>>
                <option value="1">
                  <% Sendb(Copient.PhraseLib.Lookup("term.nolimit", LanguageID))%>
                </option>
                <option value="2">
                  <% Sendb(Copient.PhraseLib.Lookup("term.oncepertransaction", LanguageID))%>
                </option>
                <option value="3">
                  <% Sendb(Copient.PhraseLib.Lookup("term.onceperday", LanguageID))%>
                </option>
                <option value="4">
                  <% Sendb(Copient.PhraseLib.Lookup("term.onceperweek", LanguageID))%>
                </option>
                <option value="5">
                  <% Sendb(Copient.PhraseLib.Lookup("term.onceperoffer", LanguageID))%>
                </option>
                <option value="0">
                  <% Sendb(Copient.PhraseLib.Lookup("term.custom", LanguageID))%>
                </option>
              </select>
            </td> 
          </tr>
          <% If (IsUniqueProd) Then Sendb("<tr><td colspan=""2""><small style=""margin-left:100px;"">(" & Copient.PhraseLib.Lookup("term.disabledunique", LanguageID) & ")</small></td><td></td></tr>")%>
          <tr id="p3row2">
            <td>
              <label for="limit3"><% Sendb(Copient.PhraseLib.Lookup("term.limit", LanguageID))%>:</label>
            </td>
            <td>
              <input type="text" class="shorter" id="limit3" name="limit3" maxlength="6" value="<% sendb(mycommon.nz(rst.rows(0).item("P3DistQtyLimit"),0)) %>"<% if(FromTemplate and disallow_limits)then sendb(" disabled=""disabled""") %> />
            </td>
          </tr>
          <tr id="p3row3">
            <td>
              <label for="limit3period"><% Sendb(Copient.PhraseLib.Lookup("term.period", LanguageID))%>:</label>
            </td>
            <td>
              <input type="text" class="shorter" id="limit3period" name="limit3period" maxlength="6" value="<% sendb(mycommon.nz(rst.rows(0).item("P3DistPeriod"),0)) %>"<% if(FromTemplate and disallow_limits)then sendb(" disabled=""disabled""") %> />
            </td>
          </tr>
          <tr id="p3row4">
            <td>
              <label for="P3DistTimeType"><% Sendb(Copient.PhraseLib.Lookup("term.type", LanguageID))%>:</label>
            </td>
            <td>
              <select id="P3DistTimeType" name="P3DistTimeType"<% if(FromTemplate and disallow_limits)then sendb(" disabled=""disabled""") %>>
                <%
                  MyCommon.QueryStr = "select TimeTypeID,PhraseID from CPE_DistributionTimeTypes with (NoLock)"
                  rst2 = MyCommon.LRT_Select
                  For Each row2 In rst2.Rows
                    If (MyCommon.NZ(rst.Rows(0).Item("P3DistTimeType"), 0) = MyCommon.NZ(row2.Item("TimeTypeID"), 0)) Then
                      Send("<option value=""" & MyCommon.NZ(row2.Item("TimeTypeID"), 0) & """ selected=""selected"">" & Copient.PhraseLib.Lookup(MyCommon.NZ(row2.Item("PhraseID"), 0), LanguageID) & "</option>")
                    Else
                      Send("<option value=""" & MyCommon.NZ(row2.Item("TimeTypeID"), 0) & """>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row2.Item("PhraseID"), 0), LanguageID) & "</option>")
                    End If
                  Next
                %>
              </select>
              <input type="hidden" id="BeginP3TimeTypeID" name="BeginP3TimeTypeID" value="<% Sendb(MyCommon.NZ(rst.Rows(0).Item("P3DistTimeType"), -1)) %>" />
            </td>
          </tr>
        </table>
        <hr class="hidden" />
      </div>
      
      <%
        If MaxTiers > 1 Then
          Send("<div class=""box"" id=""tiering"">")
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
          Send("  <input type=""text"" class=""shortest"" id=""tierlevels"" name=""tierlevels"" maxlength=""2"" value=""" & DisplayTierLevel & """" & IIf(FromTemplate And Disallow_Tiers, " disabled=""disabled""", "") & " /><br />")
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
          <input type="checkbox" class="tempcheck" id="Disallow_CRMEngine" name="Disallow_CRMEngine"<% if(disallow_crmengine)then send(" checked=""checked""") %> />
          <label for="Disallow_CRMEngine"><% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
        </span>
        <br class="printonly" />
        <% End If%>
        <label for="InboundCRMEngineID" style="position: relative;"><% Sendb(Copient.PhraseLib.Lookup("term.creationsource", LanguageID))%>:</label>
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
        <label for="crmengine" style="position: relative;"><% Sendb(Copient.PhraseLib.Lookup("offer-gen.sendoutbound", LanguageID))%>:</label>
        <br />
        <%
          ' get the EngineID list from database
          Dim CRMEngineID As Integer = MyCommon.NZ(rst.Rows(0).Item("CRMEngineID"), -1)
          MyCommon.QueryStr = "select ExtInterfaceID, Name, PhraseID from ExtCRMInterfaces with (NoLock) where deleted=0 and active=1 and OutboundEnabled=1;"
          rst2 = MyCommon.LRT_Select()
          If (Request.QueryString("new") <> "") Then
            If (CRMEngineID = -1) Then
              CRMEngineID = Int(MyCommon.Fetch_SystemOption(39))
              'Send("CRMEngineID=" & CRMEngineID)
            End If
          End If
        %>
        <select id="crmengine" name="crmengine" class="longer"<% if(FromTemplate and disallow_crmengine)then sendb(" disabled=""disabled""") %>>
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
        <label for="vendor" style="position: relative;"><% Sendb(Copient.PhraseLib.Lookup("term.chargebackvendor", LanguageID))%>:</label>
        <br />
        <select id="vendor" name="vendor" class="longer"<% if(FromTemplate and disallow_crmengine)then sendb(" disabled=""disabled""") %>>
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
      <%--
      <div class="box" id="employees">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.employees", LanguageID))%>
          </span>
        </h2>
        <%If (IsTemplate) Then%>
        <span class="temp">
          <input type="checkbox" class="tempcheck" id="Disallow_EmployeeFiltering" name="Disallow_EmployeeFiltering"<% if(disallow_employeefiltering)then send(" checked=""checked""") %> />
          <label for="Disallow_EmployeeFiltering"><% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
        </span>
        <br class="printonly" />
        <% End If%>
        <input type="checkbox" id="EmployeeFiltering" name="EmployeeFiltering" onclick="handleEmployeeFiltering();"<% if (employeefiltered)then sendb(" checked=""checked""") %><% if(FromTemplate and disallow_employeefiltering)then sendb(" disabled=""disabled""") %> />
        <label for="EmployeeFiltering"><% Sendb(Copient.PhraseLib.Lookup("offer-gen.empfilter", LanguageID))%></label>
        <br />
        &nbsp;&nbsp;
        <input type="radio" id="employeesonly" name="employeesonly"<% if (employeesonly) then sendb(" checked=""checked""") %> onclick="toggleEmployee('employeesexcluded');"<% if(FromTemplate and disallow_employeefiltering)then sendb(" disabled=""disabled""") %> />
        <label for="employeesonly"><% Sendb(Copient.PhraseLib.Lookup("term.employeesonly", LanguageID))%></label>
        <br />
        &nbsp;&nbsp;
        <input type="radio" id="employeesexcluded" name="employeesexcluded" style="padding-left: 5px;"<% if (employeesexcluded) then sendb(" checked=""checked""") %> onclick="toggleEmployee('employeesonly');"<% if(FromTemplate and disallow_employeefiltering)then sendb(" disabled=""disabled""") %> />
        <label for="employeesexcluded"><% Sendb(Copient.PhraseLib.Lookup("term.excludeemployees", LanguageID))%></label>
        <br />
        <hr class="hidden" />
      </div>
      
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
        <input type="checkbox" name="autotransferable" id="autotransferable"<%Sendb(IIf(AutoTransferable, " checked=""checked""", ""))%> />
        <label for="autotransferable"><% Sendb(Copient.PhraseLib.Lookup("term.autotransferable", LanguageID))%></label>
        <br />
        <input type="checkbox" name="deferToEOS" id="deferToEOS"<% if (deferToEOS OrElse DollarProductCondition) then sendb(" checked=""checked""") %><%if (DeferToEOSDisabled OrElse DollarProductCondition) then sendb(" disabled=""disabled""") %> />
        <label for="deferToEOS"><% Sendb(Copient.PhraseLib.Lookup("term.defercalc", LanguageID))%></label>
        <br />
        <%If (MyCommon.Fetch_CPE_SystemOption(26).Trim = "1") Then
            If (Request.QueryString("new") <> "") Then
              'reporting system options
              If (Request.QueryString("reportingimp") = "") Then
                If (MyCommon.Fetch_CPE_SystemOption(84).Trim = "1") Then
                  ReportingImp = True
                Else
                  ReportingImp = False
                End If
              End If
              If (Request.QueryString("reportingred") = "") Then
                If (MyCommon.Fetch_CPE_SystemOption(85).Trim = "1") Then
                  ReportingRed = True
                Else
                  ReportingRed = False
                End If
              End If
            End If
        %>
        <input type="checkbox" name="reportingimp" id="reportingimp"<% if (ReportingImp) then sendb(" checked=""checked""") %> />
        <label for="reportingimp"><% Sendb(Copient.PhraseLib.Lookup("offer-gen.enablereporting-imp", LanguageID))%></label>
        <br />
        <input type="checkbox" name="reportingred" id="reportingred"<% if (ReportingRed) then sendb(" checked=""checked""") %> />
        <label for="reportingred"><% Sendb(Copient.PhraseLib.Lookup("offer-gen.enablereporting-red", LanguageID))%></label>
        <br />
        <% End If%>
        <% If (MyCommon.Fetch_SystemOption(73).Trim <> "") Then%>
        <input type="checkbox" name="exporttoedw" id="exporttoedw"<% if (ExportToEDW) then sendb(" checked=""checked""") %> />
        <label for="exporttoedw"><% Sendb(Copient.PhraseLib.Lookup("term.exporttoarchive", LanguageID))%></label>
        <br />
        <% End If%>
        <%
          If (MyCommon.Extract_Val(MyCommon.Fetch_CPE_SystemOption(70)) = 1) Then
            'MyCommon.QueryStr = "Select Description, PhraseID from CPE_DeliverableTypes where IssuanceEnabled=1;"
            MyCommon.QueryStr = "select RDO.PKID, RDS.Description, RDO.Enabled from RemoteDataOptions as RDO " & _
                                "inner join RemoteDataStyles as RDS on RDO.StyleID=RDS.StyleID and RDO.RemoteDataTypeID=RDS.RemoteDataTypeID " & _
                                "where RDO.Enabled=1 and RDO.RemoteDataTypeID=1;"
            rst2 = MyCommon.LRT_Select
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
        <input type="checkbox" name="issuance" id="issuance" <% if (Issuance) then sendb(" checked=""checked""") %> />
        <label for="issuance"><% Sendb(Copient.PhraseLib.Lookup("offer-gen.enableissuance", LanguageID))%></label>
        <a href="#" onclick="javascript:showGrowPopup(event, '<%Sendb(IssuanceDetails)%>', 300, 200);"><img src="../../images/info.png" alt="<% Sendb(Copient.PhraseLib.Lookup("term.details", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.details", LanguageID)) %>" style="position: relative; top: 2px;" /></a>
        <div id="divIssuance" style="display: none;">
          Test
        </div>
        <br />
        <% End If%>
<%-- Commenting out the mfgcoupon option for CAM, 6/22 -hw
        <input type="checkbox" name="mfgCoupon" id="mfgCoupon"<% if (IsMfgCoupon) then sendb(" checked=""checked""") %> />
        <label for="mfgCoupon"><% Sendb(Copient.PhraseLib.Lookup("CPEoffer-gen.mfgcoupons", LanguageID))%></label>
        <br />
--%>
        <%If EngineID = 6 Then %>
          <input type="checkbox" name="LMGReg" id="LMGReg"<% if (CheckedLMG) then sendb(" checked=""checked""") %> />
          <label for="LMGReg"><% Sendb(Copient.PhraseLib.Lookup("term.LMGRegistrationComplete", LanguageID))%></label> 
          <br />
        <%End If %>
        <%--Mutually Exclusive check--%>
        <%
          MyCommon.QueryStr = "select IncentivePLUID from CPE_IncentivePLUs with (NoLock) where RewardOptionID=" & roid
          rst2 = MyCommon.LRT_Select()
          If rst2.Rows.Count > 0 Then
            Send("<input type=""checkbox"" name=""MutEx"" id=""MutEx"" " & IIf(MutuallyExclusive, " checked=""checked"" ", "") & " />")
            Send("<label for=""MutEx"">" & Copient.PhraseLib.Lookup("term.mutuallyexclusive", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.triggercodes", LanguageID), VbStrConv.Lowercase) & "</label>")
          End If
          
        %>
        <%
          If (Logix.UserRoles.FavoriteOffersForOthers AndAlso Not IsTemplate) Then
            Send("<br class=""half"" />")
            MyCommon.QueryStr = "select AdminUserID from AdminUserOffers where OfferID=" & OfferID & ";"
            rst2 = MyCommon.LRT_Select
            MyCommon.QueryStr = "select AdminUserID from AdminUsers;"
            rst3 = MyCommon.LRT_Select
            Send("<button type=""button"" id=""favorite"" name=""favorite"" value=""" & Copient.PhraseLib.Lookup("term.favorite", LanguageID).ToLower & """ onclick=""javascript:xmlhttpPost('/logix/OfferFeeds.aspx', 'FavoriteForAll');"">" & Copient.PhraseLib.Lookup("offer-gen.favoriteall", LanguageID) & "</button>")
            Sendb("<a href=""javascript:openPopup('/logix/offer-favorite.aspx?OfferID=" & OfferID & "')""><img id=""favImg"" src=""../../images/user.png"" ")
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
  var elemEligStart = document.getElementById("eligibilitystart");
  var elemEligEnd = document.getElementById("eligibilityend");
  var elemTestStart = document.getElementById("testingstart");
  var elemTestEnd = document.getElementById("testingend");
      
  if (targetDateField.id == "productionstart") {
    // populate productionend, etc. if unpopulated
    if (elemProdEnd != null && elemProdEnd.value == "") {
      elemProdEnd.value = targetDateField.value;
    }
    if (elemEligStart != null && elemEligStart.value == "") {
      elemEligStart.value = targetDateField.value;
    }
    if (elemEligEnd != null && elemEligEnd.value == "") {
      elemEligEnd.value = targetDateField.value;
    }
    if (elemTestStart != null && elemTestStart.value == "") {
      elemTestStart.value = targetDateField.value;
    }
    if (elemTestEnd != null && elemTestEnd.value == "") {
      elemTestEnd.value = targetDateField.value;
    }
  } else if (targetDateField.id == "productionend") {
    if (elemEligEnd != null) {
      elemEligEnd.value = targetDateField.value;
    }
    if (elemTestEnd != null) {
      elemTestEnd.value = targetDateField.value;
    }
  }
}
</script>
<script type="text/javascript">
  setPlimits();
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
<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (OfferID > 0 And Logix.UserRoles.AccessNotes) Then
      Send_Notes(3, OfferID, AdminUserID)
    End If
  End If
done:
  Send_FocusScript("mainform", "name")
  Send_WrapEnd()
  Send("<div id=""disabledBkgrd"" style=""position:absolute;top:0px;left:0px;right:0px;width:100%;height:100%;background-color:Gray;display:none;z-index:99;filter:alpha(opacity=50);-moz-opacity:.50;opacity:.50;""></div>")
  Send_PageEnd()
  MyCommon.Close_LogixRT()
  MyCommon = Nothing
  Logix = Nothing
%>
