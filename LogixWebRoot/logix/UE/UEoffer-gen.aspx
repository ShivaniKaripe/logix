<%@ Page Language="vb" Debug="true" CodeFile="ue-cb.vb" Inherits="UECB" %>

<%@ Register Src="~/logix/UserControls/UDFListControl.ascx" TagPrefix="udf" TagName="udflist" %>
<%@ Register Src="~/logix/UserControls/UDFJavaScript.ascx" TagPrefix="udf" TagName="udfjavascript" %>
<%@ Register Src="~/logix/UserControls/UDFSaveControl.ascx" TagPrefix="udf" TagName="udfsave" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="CMS.AMS" %>
<%@ Import Namespace="CMS.AMS.Contract" %>
<%@ Import Namespace="CMS.AMS.Models" %>
<script type="text/javascript">
  var nVer = navigator.appVersion;
  var nAgt = navigator.userAgent;
  var browserName = navigator.appName;
  var nameOffset, verOffset, ix;
  var didUdfChange = false;
  var didViewImage = false;

  var browser = navigator.appName;

  // In Opera, the true version is after "Opera" or after "Version"
  if ((verOffset = nAgt.indexOf("Opera")) != -1) {
    browserName = "Opera";
  }
  // In MSIE, the true version is after "MSIE" in userAgent
  else if ((verOffset = nAgt.indexOf("MSIE")) != -1) {
    browserName = "IE";
  }
  // In Chrome, the true version is after "Chrome" 
  else if ((verOffset = nAgt.indexOf("Chrome")) != -1) {
    browserName = "Chrome";
  }
  // In Safari, the true version is after "Safari" or after "Version" 
  else if ((verOffset = nAgt.indexOf("Safari")) != -1) {
    browserName = "Safari";
  }
  // In Firefox, the true version is after "Firefox" 
  else if ((verOffset = nAgt.indexOf("Firefox")) != -1) {
    browserName = "Firefox";
  }
  // In most other browsers, "name/version" is at the end of userAgent 
  else if ((nameOffset = nAgt.lastIndexOf(' ') + 1) <
          (verOffset = nAgt.lastIndexOf('/'))) {
    browserName = nAgt.substring(nameOffset, verOffset);
    fullVersion = nAgt.substring(verOffset + 1);
    if (browserName.toLowerCase() == browserName.toUpperCase()) {
      browserName = navigator.appName;
    }
  }


  if (browserName == "IE") {
    document.attachEvent("onclick", PageClick);
  }
  else {
    document.onclick = function (evt)
    {
      var target = document.all ? event.srcElement : evt.target;
      if (target.href && target.className != "calendar") {
        //didUdfChange - set to true when UDF is added/deleted, 
        //when the offer is saved, the page will get reloaded and the variables will be set to their default value of false
        //also, didUdfChange handles a very specific case that IsFormChanged does not. IsFormChanged does not capture the case where
        //the list box is added with nothing selected and then a selection is made
        //didViewImage is set when image is shown full size and cleared when closed
        if (IsFormChanged(document.mainform) || didUdfChange) {
          if (didViewImage) {
            var id = target.id.toString();
            if (id.indexOf('Image') < 0) {
              didViewImage = false;
            }
          } else if ('<%= Not IsOfferWaitingForApproval(IIf((Request.QueryString("OfferID") IsNot Nothing), Request.QueryString("OfferID"), 0)) %>' == 'True'){
            var bConfirm = confirm('Warning:  Unsaved data will be lost.  Are you sure you wish to continue?');
            return bConfirm;
          }
        }
      }
    };
  }

  function PageClick(evt)
  {
    var target = document.all ? event.srcElement : evt.target;

    if (target.href && target.className != "calendar") {
      //didUdfChange - set to true when UDF is added/deleted, 
      //when the offer is saved, the page will get reloaded and the variables will be set to their default value of false
      //also, didUdfChange handles a very specific case that IsFormChanged does not. IsFormChanged does not capture the case where
      //the list box is added with nothing selected and then a selection is made
      //didViewImage is set when image is shown full size and cleared when closed
      if (IsFormChanged(document.mainform) || didUdfChange) {
        if (didViewImage) {
          var id = target.id.toString();
          if (id.indexOf('Image') < 0) {
            didViewImage = false;
          }
        } else {
          var bConfirm = confirm('Warning:  Unsaved data will be lost.  Are you sure you wish to continue?');
          return bConfirm;
        }
      }
    }
  }

  function openPagePopup(url)
  {
    var bConfirm = true;
    if (IsFormChanged(document.mainform)) {
      bConfirm = confirm('Warning:  Unsaved data will be lost.  Are you sure you wish to continue?');
    }
    if (bConfirm) {
      popW = 700;
      popH = 570;
      siteWindow = window.open(url, "Popup", "width=" + popW + ", height=" + popH + ", top=" + calcTop(popH) + ", left=" + calcLeft(popW) + ", toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=yes");
      siteWindow.focus();
    }
  }
</script>
<% 
    ' *****************************************************************************
    ' * FILENAME: UEoffer-gen.aspx 
    ' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ' * Copyright © 2002 - 2011.  All rights reserved by:
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
    Dim infoMsg As String = ""
    Dim MyCommon As New Copient.CommonInc
    Dim MyCPEOffer As New Copient.EIW
    Dim Logix As New Copient.LogixInc
    Dim Localization As Copient.Localization
    Dim rst As DataTable
    Dim rstTemplates As DataTable
    Dim rst3, rst4 As DataTable
    Dim row As DataRow
    Dim row3 As DataRow
    Dim OfferType As Integer = -1
    Dim rowTemplates As DataRow
    Dim OfferID As Long = Request.QueryString("OfferID")
    Dim OfferName As String = ""
    Dim UpdateLevel As Integer = 0
    Dim StatusFlag As Integer
    Dim NumberofFolders As Integer = 0
    Dim rst2 As DataTable
    Dim crmdt As DataTable
    Dim row2 As DataRow
    Dim IsTemplate As Boolean
    Dim IsTemplateVal As String = ""
    Dim ActiveSubTab As Integer = 91
    Dim IntroID As String = "intro"
    Dim Disallow_EmployeeFiltering As Boolean = True
    Dim Disallow_ProductionDates As Boolean = True
    Dim Disallow_Limits As Boolean = True
    Dim Disallow_Tiers As Boolean = True
    Dim Disallow_Priority As Boolean = True
    Dim Disallow_Sweepstakes As Boolean = True
    Dim Disallow_Conditions As Boolean = True
    Dim Disallow_Rewards As Boolean = True
    Dim Disallow_RewardEvaluation As Boolean = True
    Dim Disallow_AdvancedOption As Boolean = True
    Dim Disallow_PreOrder As Boolean = True
    Dim Disallow_ExecutionEngine As Boolean = True
    Dim Disallow_CRMEngine As Boolean = True
    Dim Disallow_UserDefinedFields As Boolean = True
    Dim Disallow_MutualExclusionGroups As Boolean = True
    Dim Disallow_OfferType As Boolean = True
    Dim FromTemplate As Boolean
    Dim EmployeesExcluded As Boolean
    Dim EmployeesOnly As Boolean
    Dim ReportingImp As Boolean = False
    Dim ReportingRed As Boolean = False
    Dim EmployeeFiltered As Boolean
    Dim ExtOfferID As String = ""
    Dim ChangedVal As String = String.Empty
    Dim ArrChangedVal() As String
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
    Dim DeferCalcToTotal As Boolean = False
    Dim DeferCalcToTotalDisabled As Boolean = False
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
    Dim IsMfgCoupon As Integer = 0
    Dim ExtName As String = ""
    Dim SelectedList As String = ""
    Dim TierLevels As Integer = 1
    Dim MaxTiers As Integer = 1
    Dim IsUniqueProd As Boolean = False
    Dim AccumEnabled As Boolean = False
    Dim FuelEnabled As Boolean = False
    Dim ShortStartDate, ShortEndDate As String
    Dim ShortEligStartDate, ShortEligEndDate As String
    Dim ShortTestStartDate, ShortTestEndDate As String
    Dim StartDT, EndDT As Date
    Dim EligStartDT, EligEndDT As Date
    Dim TestStartDT, TestEndDT As Date
    Dim VendorCouponCode As String = ""
    Dim ChangeExtID As Boolean = False
    Dim DescriptLength As Boolean = False
    Dim Description As String = ""
    Dim HeaderExists As Boolean = False
    Dim DisplayTierLevel As String = ""
    Dim IsDeployable As Boolean = False
    Dim HasEIW As Boolean = False
    Dim IsEIWDateLocked As Boolean = False
    Dim ErrorMsg As String = ""
    Dim EngineID As Integer = 2
    Dim EngineSubTypeID As Integer = 0
    Dim EnginePhraseID As Integer = 0
    Dim EngineSubTypePhraseID As Integer = 0
    Dim AutoTransferable As Boolean = False
    Dim EnableCollisionDetection As Boolean = False
    Dim PreOrderEligibility As Boolean = False  'CR8
    Dim OldStartDate As Date
    Dim OldEndDate As Date
    Dim FolderStartDate As Date
    Dim FolderEndDate As Date
    Dim Priority As Integer = 0
    Dim HasAnyCustomer As Boolean 'indicates that the offer customer group condition us using the AnyCustomer group
    Dim TempQueryStr As String
    Dim DiscEvalTypeID As Integer = 0
    Dim NoLimit As Boolean = False
    Dim HourLimit As Boolean = False
    Dim OncePerOfferLimit As Boolean = False
    Dim UOMSubTypes As New Dictionary(Of Integer, Integer)
    Dim SubTypeID As Integer
    Dim CurrencyID As Integer
    Dim BlankTierLevelSent As Boolean = False
    Dim IsCustomerAssigned As Boolean = False ' indicates that the offer has not customer condition
    Dim Status As Integer
    Dim m_Offer As CMS.AMS.Contract.IOffer
    Dim m_TrackableCouponCondition As ITrackableCouponConditionService
    Dim m_TCProgram As ITrackableCouponProgramService
    Dim MLI As New Copient.Localization.MultiLanguageRec
    MLI.ItemID = OfferID
    MLI.MLTableName = "OfferTranslations"
    MLI.MLIdentifierName = "OfferID"
    MLI.StandardTableName = "CPE_Incentives"
    MLI.StandardIdentifierName = "IncentiveID"
    Dim IsPromotionDisplay As Boolean
    Dim bUseDisplayFlag As Boolean = False
    Dim DispStartDateParsed, DispEndDateParsed As Boolean
    Dim DStartDate, DEndDate As Date
    Dim DispStartDate, DispEndDate As String
    Dim form_DispStartdate As String = ""
    Dim form_DispEnddate As String = ""
    Dim form_DispStartHr As String = ""
    Dim form_DispStartMin As String = ""
    Dim form_DispEndHr As String = ""
    Dim form_DispEndMin As String = ""
    Dim DispStartHr As String = ""
    Dim DispStartMin As String = ""
    Dim DispEndHr As String = ""
    Dim DispEndMin As String = ""
    Dim Disallow_DisplayDates As Boolean = True
    Dim bUseDisplayDates As Boolean = False
    Dim dtODisp As New DataTable
    Dim tempDateTime As Date
    Dim sDateOnlyFormat As String = "MM/dd/yyyy"
    Dim sHourOnlyFormat As String = "HH"
    Dim sMinutesOnlyFormat As String = "mm"
    Const LIMIT_NONE As Integer = -1
    Const LIMIT_ONCE_PER_OFFER As Integer = -2
    Const LIMIT_HOURS_LAST_AWARDED As Integer = -3
    Dim IsProrateonDisplay As Boolean = False
    Dim bUseProrateFlag As Boolean = False
    Dim ComposedHist As String = ""
    Dim UDFHistory As String = ""
    Dim bConflictingGCRExists As Boolean
    Dim SelectID As Integer
    Dim SelectName As String = ""
    Dim dt As DataTable
    Dim dt2 As DataTable
    Dim SelectChecked As String = ""
    Dim SelectValue As Integer = 2
    Dim  testcount As Integer = 0
    Dim SavedSelectID As Integer = -1
    Dim PointsProgramWatch as Integer = Request.QueryString("PointsProgramWatch")
    Dim PointsProgramWatchDisabled as Boolean = False
    Dim rstTemp As DataTable
    Dim DefaultAsLogixID As Boolean
    Dim isEligVisible As Boolean
    Dim IsUserModifiedOffer As Boolean = False

    Dim selectDatePicker As Integer
    Dim isTranslatedOffer As Boolean
    Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean

    Dim bAllowTimeWithStartEndDates As Boolean
    Dim ProdStartHr As String = ""
    Dim ProdStartMin As String = ""
    Dim ProdEndHr As String = ""
    Dim ProdEndMin As String = ""
    Dim sProdStartDate, sProdEndDate As String
    Dim TestStartHr As String = ""
    Dim TestStartMin As String = ""
    Dim TestEndHr As String = ""
    Dim TestEndMin As String = ""
    Dim sTestStartDate, sTestEndDate As String
    Dim StoreCoupon As Integer = 0
    Dim SelectedOfferType As Integer = 0

    Const TRACKABLE_COUPON_EXPIRE_DATE_SYSOPTION_ID As Integer = 325

    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If

    'This would redirect to summary page if offer is awaiting deployment or if offer is awaiting recommendations.
    CheckIfValidOffer(MyCommon, OfferID)

    Response.Expires = 0
    MyCommon.AppName = "UEoffer-gen.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
    sDateOnlyFormat = MyCommon.GetAdminUser.Culture.DateTimeFormat.ShortDatePattern 'Set Display Date Format to Match with Current User Culture
    Localization = New Copient.Localization(MyCommon)
    CurrentRequest.Resolver.AppName = "UEoffer-gen.aspx"
    m_Offer = CurrentRequest.Resolver.Resolve(Of CMS.AMS.Contract.IOffer)()
    m_TrackableCouponCondition = CurrentRequest.Resolver.Resolve(Of ITrackableCouponConditionService)()
    m_TCProgram = CurrentRequest.Resolver.Resolve(Of ITrackableCouponProgramService)()
    BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
    AllowMultipleBanners = (MyCommon.Fetch_SystemOption(67) = "1")
    HasAnyCustomer = UEOffer_Has_AnyCustomer(MyCommon, OfferID)
    selectDatePicker = MyCommon.Extract_Val(MyCommon.NZ(MyCommon.Fetch_SystemOption(161), 0))
    isTranslatedOffer = MyCommon.IsTranslatedUEOffer(OfferID, MyCommon)
    bEnableRestrictedAccessToUEOfferBuilder = IIf(MyCommon.Fetch_SystemOption(249) = "1", True, False)

    bAllowTimeWithStartEndDates = (MyCommon.Fetch_UE_SystemOption(200) = "1")

    Dim bEnableAdditionalLockoutRestrictionsOnOffers As Boolean = IIf(MyCommon.Fetch_SystemOption(260) = "1", True, False)
    Dim bOfferEditable As Boolean = MyCommon.IsOfferEditablePastLockOutPeriod(Logix.UserRoles.EditOfferPastLockoutPeriod, MyCommon, OfferID)

    If MyCommon.Fetch_UE_SystemOption(145) = "1" Then
        bUseDisplayFlag = True
    Else
        bUseDisplayFlag = False
    End If

    If MyCommon.Fetch_UE_SystemOption(143) = "1" Then
        bUseDisplayDates = True
    Else
        bUseDisplayDates = False
    End If

    If MyCommon.Fetch_UE_SystemOption(154) = "1" Then
        bUseProrateFlag = True
    Else
        bUseProrateFlag = False
    End If

    'Set default Impression and Redemption defaults
    ReportingImp = (MyCommon.Fetch_UE_SystemOption(84) = "1")
    ReportingRed = (MyCommon.Fetch_UE_SystemOption(85) = "1")
    If Request.QueryString("new") <> "" Then
        MyCommon.QueryStr = "update CPE_Incentives set EnableImpressRpt=" & IIf(ReportingImp, "1", "0") & ", EnableRedeemRpt=" & IIf(ReportingRed, "1", "0") & " where IncentiveID=" & OfferID
        MyCommon.LRT_Execute()
    End If

    Popup = IIf((Request.QueryString("Popup") <> "") AndAlso (Request.QueryString("Popup") <> "0"), True, False)

    MyCommon.QueryStr = "select RewardOptionID, TierLevels, CurrencyID from CPE_RewardOptions with (NoLock) where TouchResponse=0 and IncentiveID=" & OfferID
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
        roid = rst.Rows(0).Item("RewardOptionID")
        TierLevels = rst.Rows(0).Item("TierLevels")
        CurrencyID = rst.Rows(0).Item("CurrencyID")
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

    MyCommon.QueryStr = "select * from FolderItems where LinkID=" & OfferID
    rst = MyCommon.LRT_Select
    NumberofFolders = rst.Rows.Count
    'Determine if the offer has an enterprise instant win condition
    MyCommon.QueryStr = "select IncentiveEIWID from CPE_IncentiveEIW with (NoLock) where RewardOptionID=" & roid & ";"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
        HasEIW = True
    End If

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
            MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offerextid-remove", LanguageID))
        ElseIf ExtOfferID > 0 Then
            MyCommon.QueryStr = "update CPE_Incentives set ClientOfferID=" & ExtOfferID & " where IncentiveID=" & Request.QueryString("OfferID")
            MyCommon.LRT_Execute()
            MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offerextid-create", LanguageID))
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

    'Check to see if the offer is associated with a fuel terminal; if so, the tier input will be disabled, since fuel and multi-tiers are incompatible
    MyCommon.QueryStr = "select OT.TerminalTypeID, TT.FuelProcessing from OfferTerminals as OT with (RowLock) " & _
                        "inner join TerminalTypes as TT with (NoLock) on TT.TerminalTypeID=OT.TerminalTypeID " & _
                        "where OfferID=" & OfferID & ";"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
        For Each row In rst.Rows
            If MyCommon.NZ(row.Item("FuelProcessing"), False) = True Then
                FuelEnabled = True
            End If
        Next
    End If

    ' Check if the offer has a customer group
    MyCommon.QueryStr = "SELECT CG.CustomerGroupID, CG.Name, ICG.ExcludedUsers FROM CPE_IncentiveCustomerGroups AS ICG WITH (NoLock) " & _
                "LEFT JOIN CustomerGroups AS CG WITH (NoLock) ON CG.CustomerGroupID=ICG.CustomerGroupID " & _
                "LEFT JOIN CPE_RewardOptions AS RO WITH (NoLock) ON RO.RewardOptionID=ICG.RewardOptionID " & _
                "WHERE ICG.ExcludedUsers=0 AND ICG.Deleted=0 AND CG.Deleted=0 AND RO.Deleted=0 AND RO.IncentiveID=" & OfferID & ";"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
        IsCustomerAssigned = True
    End If

    MyCommon.QueryStr = "SELECT IA.IncentiveAttributeID FROM CPE_IncentiveAttributes AS IA WITH (NoLock) " & _
                "LEFT JOIN CPE_RewardOptions AS RO WITH (NoLock) ON IA.RewardOptionID=RO.RewardOptionID " & _
                "WHERE IA.Deleted=0 And RO.IncentiveID=" & OfferID & ";"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
        IsCustomerAssigned = True
    End If

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

    ' Check for Discount on the offer; if present, then we need to disable Defer to EOS, if anything but Cashier Msg or Pass Thru then disable Defer to Total
    MyCommon.QueryStr = "select DeliverableTypeID from CPE_Deliverables with (NoLock) " & _
                      " where Deleted=0 " & _
                        "  and RewardOptionID = " & roid
    rst = MyCommon.LRT_Select()
    If rst.Rows.Count > 0 Then
        For each row in rst.Rows
            If row.Item("DeliverableTypeID") = 2 Then
                DeferToEOSDisabled = True
                DeferCalcToTotalDisabled = True
                PointsProgramWatchDisabled = True
                PointsProgramWatch = 0
            ElseIf row.Item("DeliverableTypeID") = 16 Then
                DeferToEOSDisabled = True
                DeferCalcToTotalDisabled = True
            ElseIf row.Item("DeliverableTypeID") <> 9 AndAlso row.Item("DeliverableTypeID") <> 12 Then
                DeferCalcToTotalDisabled = True
                PointsProgramWatchDisabled = True
                PointsProgramWatch = 0
            End If
        Next
    End If

    ' Check if the cashier message has to be deferred to total
    MyCommon.QueryStr = "select DeferCalcToTotal from CPE_Incentives where IncentiveID = " & Request.QueryString("OfferID") +" And DeferCalcToTotal = 1"
    rst = MyCommon.LRT_Select()
    If rst.Rows.Count > 0 Then
        DeferToEOSDisabled = True
    End If

    If MyCommon.Fetch_SystemOption(156) = "1" Then
                %>
                    <udf:udfsave id="udfsavecontrol" runat="server" />
                <%

                        infoMessage = udfsavecontrol.infoMessage
                        UDFHistory = udfsavecontrol.UDFHistory
                    End If ' MyCommon.Fetch_SystemOption(156) = "1" 
                    'Save
                    If (Request.QueryString("save") <> "" AndAlso (MyCommon.Fetch_SystemOption(191) AndAlso NumberofFolders > 1)) Then
                        infoMessage = Copient.PhraseLib.Lookup("term.cannotsavemultiplefolder", LanguageID)
                    End If
                    If (Request.QueryString("save") <> "" AndAlso Not (MyCommon.Fetch_SystemOption(191) AndAlso NumberofFolders > 1)) Then

                        'Check the Description length
                        Dim Descfromspecial As String = ""
                        Dim dtDesc As DataTable
                        MyCommon.QueryStr = "Select Description from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID
                        dtDesc = MyCommon.LRT_Select
                        If dtDesc.Rows.Count > 0 Then
                            Descfromspecial = MyCommon.NZ(dtDesc.Rows(0)(0), "")
                            MLI.StandardValue = Descfromspecial
                        End If
                        If Descfromspecial <> "" Then
                            Description = Descfromspecial.Replace("&quot;", Chr(34))
                            If Description.Length <= 1000 Then
                                DescriptLength = True
                            End If
                        Else
                            DescriptLength = True
                        End If
                        If MyCommon.Fetch_UE_SystemOption(175) = "1" Then
                            If Request.QueryString("PromoGridCategory") <> "" Then
                                '	 sqlBuf.Append("PromoOfferCategories=" & MyCommon.Extract_Val(Request.QueryString("PromoOfferCategories")) & ",")
                                If MyCommon.Extract_Val(Request.QueryString("PromoGridCategory")) >-1 Then
                                    SavedSelectID=Request.QueryString("PromoGridCategory")
                                    'MyCommon.QueryStr = "MERGE INTO PromoGridOffers AS target USING (SELECT "& OfferID &") AS source (OfferID) ON target.IncentiveID = source.OfferID WHEN MATCHED THEN UPDATE SET target.PromoCategoryID=" & SavedSelectID &
                                    '"WHEN NOT MATCHED BY TARGET THEN INSERT(PromoCategoryID, IncentiveID) VALUES(" & SavedSelectID &","& OfferID &");"
                                    'MyCommon.LRT_Execute()
                                    MyCommon.QueryStr = "dbo.pt_SavePromoCategory"
                                    MyCommon.Open_LRTsp()
                                    MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
                                    MyCommon.LRTsp.Parameters.Add("@PromoCategoryID", SqlDbType.Int).Value = SavedSelectID
                                    MyCommon.LRTsp.Parameters.Add("@Priority", SqlDbType.BigInt).Direction = ParameterDirection.Output
                                    MyCommon.LRTsp.ExecuteNonQuery()
                                    Priority = MyCommon.LRTsp.Parameters("@Priority").Value
                                    MyCommon.Close_LRTsp()
                                Else
                                    MyCommon.QueryStr = "Update PromoGridOffers set Deleted = 1 where IncentiveID = " & OfferID
                                    MyCommon.LRT_Execute()
                                End If
                            End If
                        End If
                        'Get the current production start and end dates (prior to saving the new ones).
                        'We'll use these below if there are any EIW conditions that need to be rerandomized.
                        MyCommon.QueryStr = "select StartDate, EndDate from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID & ";"
                        rst = MyCommon.LRT_Select
                        If rst.Rows.Count > 0 Then
                            OldStartDate = MyCommon.NZ(rst.Rows(0).Item("StartDate"), "1/1/1900")
                            OldEndDate = MyCommon.NZ(rst.Rows(0).Item("EndDate"), "1/1/1900")
                        End If
                        MyCommon.QueryStr = "select fold.StartDate, fold.EndDate from FolderItems fi join Folders fold on fi.FolderID=fold.FolderID where fi.LinkID=" & OfferID & ";"
                        rst = MyCommon.LRT_Select
                        If rst.Rows.Count > 0 Then
                            FolderStartDate = MyCommon.NZ(rst.Rows(0).Item("StartDate"), "1/1/1900")
                            FolderEndDate = MyCommon.NZ(rst.Rows(0).Item("EndDate"), "1/1/1900")
                        Else
                            FolderStartDate = "1/1/1900"
                            FolderEndDate = "1/1/1900"
                        End If


                        If bUseDisplayDates Then
                            If Request.QueryString("displaystart") <> "" Then
                                DispStartDateParsed = Date.TryParse(Request.QueryString("displaystart"), MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, DStartDate)
                            End If
                            If Request.QueryString("displayend") <> "" Then
                                DispEndDateParsed = Date.TryParse(Request.QueryString("displayend"), MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, DEndDate)
                            End If
                            If DispStartDateParsed And (Not DispEndDateParsed) Then
                                DEndDate = DStartDate
                            End If
                            form_DispStartdate = Request.QueryString("displaystart")
                            form_DispEnddate = Request.QueryString("displayend")

                            form_DispStartHr = Request.QueryString("form_DispStartHr")
                            form_DispStartMin = Request.QueryString("form_DispStartMin")

                            If form_DispStartHr = "" Then
                                form_DispStartHr = "00"
                            Else
                                form_DispStartHr = form_DispStartHr
                            End If
                            If form_DispStartMin = "" Then
                                form_DispStartMin = "00"
                            Else
                                form_DispStartMin = form_DispStartMin
                            End If

                            If Not String.IsNullOrEmpty(form_DispStartdate) Then
                                form_DispStartdate = form_DispStartdate & " " & form_DispStartHr & ":" & form_DispStartMin & ":00"
                            End If

                            form_DispEndHr = Request.QueryString("form_DispEndHr")
                            form_DispEndMin = Request.QueryString("form_DispEndMin")

                            If form_DispEndHr = "" Then
                                form_DispEndHr = "23"
                            Else
                                form_DispEndHr = form_DispEndHr
                            End If
                            If form_DispEndMin = "" Then
                                form_DispEndMin = "59"
                            Else
                                form_DispEndMin = form_DispEndMin
                            End If

                            If Not String.IsNullOrEmpty(form_DispEnddate) Then
                                form_DispEnddate = form_DispEnddate & " " & form_DispEndHr & ":" & form_DispEndMin & ":00"
                            End If

                            If form_DispStartdate <> "" And form_DispEnddate = "" Then
                                form_DispEnddate = Request.QueryString("displaystart") & " " & form_DispEndHr & ":" & form_DispEndMin & ":00"
                            End If

                            If form_DispStartdate <> "" And form_DispEnddate = "" Then
                                form_DispEnddate = Request.QueryString("displaystart")
                            End If

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
                        If (rst.Rows.Count > 0 AndAlso MyCommon.NZ(rst.Rows(0).Item("DiscountTypeID"), 0) = 3 AndAlso MyCommon.NZ(rst.Rows(0).Item("AmountTypeID"), 0) <> 1 AndAlso Request.QueryString("offerType") = "1") Then
                            infoMessage = Copient.PhraseLib.Lookup("ueoffer-gen.InvalidMfgCoupon", LanguageID)
                        ElseIf (Request.QueryString("productionstart") = "" Or Request.QueryString("productionend") = "") Then
                            infoMessage = Copient.PhraseLib.Lookup("CPEoffer_gen.noproductiondates", LanguageID)
                        ElseIf bUseDisplayDates And (Request.QueryString("displaystart") = "" AndAlso Request.QueryString("displayend") <> "") Then
                            infoMessage = Copient.PhraseLib.Lookup("offer-gen.invaliddispstartdate", LanguageID)
                        ElseIf DescriptLength = False Then
                            infoMessage = Copient.PhraseLib.Lookup("error.description", LanguageID)
                        ElseIf (HasEIW) And Not (Request.QueryString("issuance") = "on") Then
                            infoMessage = Copient.PhraseLib.Lookup("ueoffer-gen.InvalidDisableIssuance", LanguageID)
                        ElseIf ((rst2.Rows.Count > 0) AndAlso (MyCommon.Extract_Val(Request.QueryString("form_Category")) <> MyCommon.NZ(rst2.Rows(0).Item("OfferCategoryID"), 0))) Then
                            infoMessage = Copient.PhraseLib.Lookup("ueoffer-gen.InvalidCategoryChange", LanguageID)
                        ElseIf (Request.QueryString("deferToEOS") = "on" AndAlso Request.QueryString("DeferCalcToTotal") = "on") Then
                            infoMessage = Copient.PhraseLib.Lookup("error.invaliddefers", LanguageID)
                        ElseIf (Request.QueryString("PointsProgramWatch") > 0 AndAlso Request.QueryString("DeferCalcToTotal") = "") Then
                            infoMessage = Copient.PhraseLib.Lookup("error.nodefer", LanguageID)
                        ElseIf (Request.QueryString("PointsProgramWatch") = 0 AndAlso Request.QueryString("DeferCalcToTotal") = "on") Then
                            infoMessage = Copient.PhraseLib.Lookup("error.noptprg", LanguageID)
                        Else
                            sProdStartDate = Request.QueryString("productionstart")
                            sProdEndDate = Request.QueryString("productionend")
                            StartDateParsed = Date.TryParse(sProdStartDate, MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, ProdStartDate)
                            EndDateParsed = Date.TryParse(sProdEndDate, MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, ProdEndDate)
                            If (StartDateParsed AndAlso EndDateParsed) Then
                                StartDateParsed = Date.TryParse(Request.QueryString("eligibilitystart"), MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, EligStartDate)
                                If (Not StartDateParsed) Then EligStartDate = ProdStartDate
                                EndDateParsed = Date.TryParse(Request.QueryString("eligibilityend"), MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, EligEndDate)
                                If (Not EndDateParsed) Then EligEndDate = ProdEndDate

                                sTestStartDate = Request.QueryString("testingstart")
                                sTestEndDate = Request.QueryString("testingend")
                                StartDateParsed = Date.TryParse(sTestStartDate, MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, TestStartDate)
                                If (Not StartDateParsed) Then
                                    TestStartDate = ProdStartDate
                                    sTestStartDate = sProdStartDate
                                End If
                                EndDateParsed = Date.TryParse(sTestEndDate, MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, TestEndDate)
                                If (Not EndDateParsed) Then
                                    TestEndDate = ProdEndDate
                                    sTestEndDate = sProdEndDate
                                End If

                                If bAllowTimeWithStartEndDates Then
                                    TestStartHr = MyCommon.NZ(Request.QueryString("form_TestStartHr"), "00")
                                    TestStartMin = MyCommon.NZ(Request.QueryString("form_TestStartMin"), "00")
                                    TestStartDate = TestStartDate & " " & TestStartHr & ":" & TestStartMin & ":00"

                                    TestEndHr = MyCommon.NZ(Request.QueryString("form_TestEndHr"), "23")
                                    TestEndMin = MyCommon.NZ(Request.QueryString("form_TestEndMin"), "59")
                                    TestEndDate = TestEndDate & " " & TestEndHr & ":" & TestEndMin & ":00"

                                    ProdStartHr = MyCommon.NZ(Request.QueryString("form_ProdStartHr"), "00")
                                    ProdStartMin = MyCommon.NZ(Request.QueryString("form_ProdStartMin"), "00")
                                    ProdStartDate = ProdStartDate & " " & ProdStartHr & ":" & ProdStartMin & ":00"

                                    ProdEndHr = MyCommon.NZ(Request.QueryString("form_ProdEndHr"), "23")
                                    ProdEndMin = MyCommon.NZ(Request.QueryString("form_ProdEndMin"), "59")
                                    ProdEndDate = ProdEndDate & " " & ProdEndHr & ":" & ProdEndMin & ":00"
                                End If

                                Dim strQuerySupportOfferName As String = Logix.TrimAll(Request.QueryString("form_name"))
                                Dim strQuerySupportDescription As String = Description
                                ' check for an incentive already with that name
                                MyCommon.QueryStr = "select IncentiveName from CPE_Incentives with (NoLock) " & _
                                                    "where Deleted=0 and IncentiveName='" & MyCommon.Parse_Quotes(strQuerySupportOfferName) & "' and IncentiveID<>" & Request.QueryString("OfferID") & _
                                                    " union all " & _
                                                    "select Name from Offers with (NoLock) " & _
                                                    "where Deleted=0 and Name='" & MyCommon.Parse_Quotes(strQuerySupportOfferName) & "' and OfferID<>" & Request.QueryString("OfferID") & ";"
                                rst = MyCommon.LRT_Select
                                DuplicateName = (rst.Rows.Count > 0)

                                'Check OfferType before updating it
                                OfferType = m_Offer.GetOfferType(OfferID)

                                sqlBuf.Append("Update CPE_Incentives with (RowLock) set ")
                                sqlBuf.Append("IncentiveName=N'" & MyCommon.Parse_Quotes(strQuerySupportOfferName) & "',")
                                sqlBuf.Append("Description=N'" & MyCommon.Parse_Quotes(strQuerySupportDescription) & "',")
                                If (Request.QueryString("form_Category") <> "") Then
                                    sqlBuf.Append("PromoClassID=" & MyCommon.Extract_Val(Request.QueryString("form_Category")) & ",")
                                End If
                                If (Request.QueryString("priority") <> "" and Priority =0) Then
                                    sqlBuf.Append("Priority=" & MyCommon.Extract_Val(Request.QueryString("priority")) & ",")
                                End If
                                If (Request.QueryString("crmengine") <> "") Then
                                    sqlBuf.Append("CRMEngineID=" & MyCommon.Extract_Val(Request.QueryString("crmengine")) & ",")
                                End If

                                Select Case Request.QueryString("P3DistTimeType")
                                    Case LIMIT_NONE.ToString
                                        sqlBuf.Append("P3DistTimeType=2,")
                                        sqlBuf.Append("P3DistQtyLimit=0,") 'limit is set to 0
                                        sqlBuf.Append("P3DistPeriod=0,")
                                        bConflictingGCRExists = (ExistGCRPercentOff(MyCommon, roid) AndAlso ExistProductPriceCondition(MyCommon, roid))
                                    Case LIMIT_ONCE_PER_OFFER.ToString
                                        sqlBuf.Append("P3DistTimeType=1,")
                                        sqlBuf.Append("P3DistQtyLimit=1,")
                                        sqlBuf.Append("P3DistPeriod=3650,")
                                    Case LIMIT_HOURS_LAST_AWARDED.ToString
                                        sqlBuf.Append("P3DistTimeType=1,")
                                        sqlBuf.Append("P3DistQtyLimit=" & MyCommon.Extract_Val(Request.QueryString("limit3")) & ",")
                                        ' negative limit3period will signify hourly offer limit and a positive limit3period will signify Days
                                        sqlBuf.Append("P3DistPeriod=" & MyCommon.Extract_Val(Request.QueryString("limit3period")) * (-1) & ",")
                                    Case Else
                                        If (IsUniqueProd) Then
                                            sqlBuf.Append("P3DistTimeType=2,")
                                        ElseIf (Request.QueryString("P3DistTimeType") <> "") Then
                                            sqlBuf.Append("P3DistTimeType=" & MyCommon.Extract_Val(Request.QueryString("P3DistTimeType")) & ",")
                                        End If

                                        If (IsUniqueProd) Then
                                            sqlBuf.Append("P3DistQtyLimit=1,")
                                        ElseIf (Request.QueryString("limit3") <> "") Then
                                            Dim limit As String = MyCommon.Extract_Val(Request.QueryString("limit3"))
                                            sqlBuf.Append("P3DistQtyLimit=" & MyCommon.Extract_Val(Request.QueryString("limit3")) & ",")
                                            'Al-5916
                                            bConflictingGCRExists = (ExistGCRPercentOff(MyCommon, roid) AndAlso ExistProductPriceCondition(MyCommon, roid) AndAlso limit <> 1)
                                        End If

                                        If (IsUniqueProd) Then
                                            sqlBuf.Append("P3DistPeriod=1,")
                                        ElseIf (Request.QueryString("limit3period") <> "") Then
                                            sqlBuf.Append("P3DistPeriod=" & MyCommon.Extract_Val(Request.QueryString("limit3period")) & ",")
                                        ElseIf (Request.QueryString("limit3period") = "") Then
                                            sqlBuf.Append("P3DistPeriod=null,")
                                        End If
                                End Select

                                If bAllowTimeWithStartEndDates Then
                                    sqlBuf.Append("StartDate='" & ProdStartDate & "',")
                                    sqlBuf.Append("EndDate='" & ProdEndDate & "',")
                                Else
                                    sqlBuf.Append("StartDate='" & ProdStartDate.ToShortDateString & "',")
                                    sqlBuf.Append("EndDate='" & ProdEndDate.ToShortDateString & "',")
                                End If

                                sqlBuf.Append("EligibilityStartDate='" & EligStartDate.ToShortDateString & "',")
                                sqlBuf.Append("EligibilityEndDate='" & EligEndDate.ToShortDateString & "',")

                                If bAllowTimeWithStartEndDates Then
                                    sqlBuf.Append("TestingStartDate='" & TestStartDate & "',")
                                    sqlBuf.Append("TestingEndDate='" & TestEndDate & "',")
                                Else
                                    sqlBuf.Append("TestingStartDate='" & TestStartDate.ToShortDateString & "',")
                                    sqlBuf.Append("TestingEndDate='" & TestEndDate.ToShortDateString & "',")
                                End If

                                sqlBuf.Append("EmployeesOnly=" & IIf(Request.QueryString("employeesonly") = "on", 1, 0) & ",")
                                sqlBuf.Append("EnableImpressRpt=" & IIf(Request.QueryString("reportingimp") = "on", 1, 0) & ",")
                                sqlBuf.Append("EnableRedeemRpt=" & IIf(Request.QueryString("reportingred") = "on", 1, 0) & ",")
                                sqlBuf.Append("EmployeesExcluded=" & IIf(Request.QueryString("employeesexcluded") = "on", 1, 0) & ",")
                                sqlBuf.Append("DeferCalcToEOS=" & IIf(Request.QueryString("deferToEOS") = "on", 1, 0) & ",")
                                sqlBuf.Append("ExportToEDW=" & IIf(Request.QueryString("exporttoedw") = "on", 1, 0) & ",")
                                sqlBuf.Append("Favorite=" & IIf(Request.QueryString("favorite") = "on", 1, 0) & ",")
                                If Request.QueryString("InboundCRMEngineID") <> "" Then
                                    sqlBuf.Append("InboundCRMEngineID=" & MyCommon.Extract_Val(Request.QueryString("InboundCRMEngineID")) & ",")
                                End If
                                sqlBuf.Append("SendIssuance=" & IIf(Request.QueryString("issuance") = "on", 1, 0) & ",")
                                sqlBuf.Append("ChargebackVendorID=" & MyCommon.Extract_Val(Request.QueryString("vendor")) & ",")
                                If Request.QueryString("offerType") = "1" Then
                                    sqlBuf.Append("ManufacturerCoupon=1 ,")
                                Else
                                    sqlBuf.Append("ManufacturerCoupon=0 ,")
                                End If
                                'sqlBuf.Append("ManufacturerCoupon=" & IIf(Request.QueryString("mfgCoupon") = "on", 1, 0) & ",")
                                sqlBuf.Append("VendorCouponCode=N'" & MyCommon.Parse_Quotes(Request.QueryString("vendorCouponCode")) & "', ")
                                sqlBuf.Append("AutoTransferable=" & IIf(Request.QueryString("autotransferable") = "on", 1, 0) & ",")
                                sqlBuf.Append("PreOrderEligibility=" & IIf(Request.QueryString("preordereligibility") = "on", 1, 0) & ",")  'CR8
                                sqlBuf.Append("LastUpdate=getdate(), ")
                                sqlBuf.Append("LastUpdatedByAdminID=" & AdminUserID & ", ")
                                'If the offer is Airmiles allow for the changing of the external ID
                                If EngineID = 2 AndAlso EngineSubTypeID = 2 Then
                                    If (Request.QueryString("ExtID") <> "") Then
                                        sqlBuf.Append("ClientOfferID='" & MyCommon.Parse_Quotes(Logix.TrimAll(Left(Request.QueryString("ExtID"), 20))) & "', ")
                                    Else
                                        sqlBuf.Append("ClientOfferID=NULL,")
                                    End If
                                End If
                                sqlBuf.Append("DiscountEvalTypeID=" & MyCommon.Extract_Val(Request.QueryString("discEvalType")) & ", ")
                                sqlBuf.Append("PromotionDisplay=" & IIf(Request.QueryString("promotiondisplay") = "on", 1, 0) & ",")
                                sqlBuf.Append("ProrateonDisplay=" & IIf(Request.QueryString("prorateondisplay") = "on", 1, 0) & ",")
                                sqlBuf.Append("DeferCalcToTotal=" & IIf(Request.QueryString("DeferCalcToTotal") = "on", 1, 0) & ",")
                                sqlBuf.Append("PointsProgramWatch=" & MyCommon.Extract_Val(Request.QueryString("PointsProgramWatch")) & ",")
                                sqlBuf.Append("EnableCollisionDetection=" & IIf(Request.QueryString("enableCollisionDetection") = "on", 1, 0) & ",")
                                If Request.QueryString("offerType") = "2" Then
                                    sqlBuf.Append("StoreCoupon=1 ,")
                                Else
                                    sqlBuf.Append("StoreCoupon=0 ,")
                                End If
                                sqlBuf.Append("StatusFlag=1 ")
                                sqlBuf.Append("where IncentiveID=" & MyCommon.Extract_Val(Request.QueryString("OfferID")))

                                If (Not String.IsNullOrEmpty(Request.QueryString("hdnEmployeeRow"))) Then
                                    isEligVisible = Convert.ToBoolean(Request.QueryString("hdnEmployeeRow"))
                                End If
                                'Send(MyCommon.QueryStr)
                                If (ProdEndDate < ProdStartDate) Then
                                    infoMessage = Copient.PhraseLib.Lookup("offer-gen.baddate", LanguageID)
                                ElseIf (ProdEndDate >= CDate("1/1/9999")) Then  'the date value is larger that what we can allow
                                    infoMessage = Copient.PhraseLib.Lookup("term.datetoolarge", LanguageID)
                                ElseIf (isEligVisible And EligEndDate < EligStartDate) Then
                                    infoMessage = Copient.PhraseLib.Lookup("offer-gen.baddate", LanguageID)
                                ElseIf (EligEndDate >= CDate("1/1/9999")) Then  'the date value is larger that what we can allow
                                    infoMessage = Copient.PhraseLib.Lookup("term.datetoolarge", LanguageID)
                                ElseIf (TestEndDate < TestStartDate) Then
                                    infoMessage = Copient.PhraseLib.Lookup("offer-gen.baddate", LanguageID)
                                ElseIf (TestEndDate >= CDate("1/1/9999")) Then  'the date value is larger that what we can allow
                                    infoMessage = Copient.PhraseLib.Lookup("term.datetoolarge", LanguageID)
                                ElseIf (bUseDisplayDates AndAlso DispStartDateParsed AndAlso DispEndDateParsed AndAlso (CDate(DEndDate) < CDate(DStartDate))) Then
                                    infoMessage = Copient.PhraseLib.Lookup("offer-gen.dispbaddate", LanguageID)
                                ElseIf (MyCommon.Fetch_SystemOption(192) = "1" AndAlso CDate(FolderStartDate) = CDate("1/1/1900") AndAlso CDate(FolderEndDate) = CDate("1/1/1900")) AndAlso NumberofFolders > 0 Then
                                    infoMessage = Copient.PhraseLib.Lookup("folders.FolderDatesEmpty", LanguageID)
                                ElseIf (MyCommon.Fetch_SystemOption(192) = "1" AndAlso CDate(FolderStartDate) <> CDate("1/1/1900") AndAlso CDate(FolderEndDate) <> CDate("1/1/1900") AndAlso (ProdStartDate < FolderStartDate Or ProdEndDate < FolderStartDate Or ProdStartDate > FolderEndDate Or ProdEndDate > FolderEndDate) OrElse (TestStartDate < FolderStartDate Or TestEndDate < FolderStartDate Or TestStartDate > FolderEndDate Or TestEndDate > FolderEndDate)) AndAlso NumberofFolders > 0 Then
                                    infoMessage = Copient.PhraseLib.Lookup("folders.OfferNotInFolderDateRange", LanguageID) & "(" & FolderStartDate & " - " & FolderEndDate & ")"
                                ElseIf DuplicateName Then
                                    infoMessage = Copient.PhraseLib.Lookup("offer-gen.nameused", LanguageID)
                                ElseIf Logix.TrimAll(Request.QueryString("form_name")) = "" Then
                                    infoMessage = Copient.PhraseLib.Lookup("offer-gen.noname", LanguageID)
                                ElseIf ((MyCommon.Extract_Val(Request.QueryString("P3DistTimeType")) = 1) _
                                    OrElse ( _
                                        MyCommon.Extract_Val(Request.QueryString("P3DistTimeType")) = 2)) _
                                    AndAlso ( _
                                      Not Integer.TryParse(Request.QueryString("limit3"), TempInt) _
                                      OrElse ( _
                                        TempInt < 0 AndAlso MyCommon.Extract_Val(Request.QueryString("P3DistTimeType")) > 0 _
                                      ) OrElse ( _
                                        TempInt <= 0 AndAlso (MyCommon.Extract_Val(Request.QueryString("P3DistTimeType")) = 0 Or MyCommon.Extract_Val(Request.QueryString("P3DistTimeType")) = 1) _
                                      ) _
                                  ) Then
                                    infoMessage = Copient.PhraseLib.Lookup("offer-gen.badlimit", LanguageID)
                                ElseIf TempInt <> 1 And bConflictingGCRExists Then
                                    infoMessage = Copient.PhraseLib.Lookup("offer-gen.badlimitduetoGCRPercentOff", LanguageID)
                                ElseIf MyCommon.Extract_Val(Request.QueryString("P3DistTimeType")) = 1 _
                                    AndAlso ( _
                                      Not Integer.TryParse(Request.QueryString("limit3period"), TempInt) _
                                      OrElse ( _
                                        TempInt < 0 AndAlso MyCommon.Extract_Val(Request.QueryString("P3DistTimeType")) > 0 _
                                      ) OrElse ( _
                                        TempInt <= 0 AndAlso (MyCommon.Extract_Val(Request.QueryString("P3DistTimeType")) = 0 Or MyCommon.Extract_Val(Request.QueryString("P3DistTimeType")) = 1) _
                                      ) _
                                  ) Then
                                    infoMessage = Copient.PhraseLib.Lookup("offer-gen.badlimit", LanguageID)
                                ElseIf BlankTierLevelSent OrElse (Request.QueryString("tierlevels") <> "" AndAlso (Not Integer.TryParse(Request.QueryString("tierlevels"), TempInt) OrElse TempInt < 1 OrElse TempInt > MaxTiers)) Then
                                    infoMessage = Copient.PhraseLib.Lookup("offer-gen.invalidtierscpe", LanguageID) & MaxTiers & "."
                                ElseIf (MyCommon.Extract_Val(Request.QueryString("tierlevels")) > 1 AndAlso (Request.QueryString("crmengine") = 1 OrElse Request.QueryString("crmengine") = 2)) Then
                                    infoMessage = Copient.PhraseLib.Lookup("offer-gen.invalidoutbound", LanguageID)
                                Else
                                    MyCommon.QueryStr = sqlBuf.ToString
                                    MyCommon.LRT_Execute()
                                    ResetOfferApprovalStatus(OfferID)

                                    'Set ChargeBacDept for discount reward based on offer type if the offer type is updated
                                    If (OfferType <> -1 AndAlso OfferType <> Request.QueryString("offerType")) Then
                                        Dim ChargeBackDeptName As String = m_Offer.SetChargeBackDept(OfferID, Request.QueryString("offerType"), LanguageID)
                                        If (ChargeBackDeptName <> "") Then
                                            Dim strOfferType As String = ""
                                            If (Request.QueryString("offerType") = 0) Then
                                                strOfferType = Copient.PhraseLib.Lookup("term.standardoffer", LanguageID)
                                            ElseIf (Request.QueryString("offerType") = 1) Then
                                                strOfferType = Copient.PhraseLib.Lookup("term.manufacturercoupon", LanguageID)
                                            Else
                                                strOfferType = Copient.PhraseLib.Lookup("term.storecoupon", LanguageID)
                                            End If
                                            infoMsg = Copient.PhraseLib.Detokenize("setchargebackdept-UEOffer", LanguageID, ChargeBackDeptName, strOfferType)
                                        End If
                                    End If

                                    'Update ExpireDate of any associated Trackable Coupon Programs
                                    Dim lstTCProgramCondition As New AMSResult(Of List(Of TCProgramCondition))
                                    Dim bTCPExpireDateEnabled As Boolean = IIf(MyCommon.Fetch_SystemOption(TRACKABLE_COUPON_EXPIRE_DATE_SYSOPTION_ID) = "1", True, False)

                                    lstTCProgramCondition = m_TrackableCouponCondition.GetTCProgramConditions(OfferID, Engines.UE)
                                    If (lstTCProgramCondition.ResultType <> AMSResultType.Success) Then
                                        infoMessage = lstTCProgramCondition.GetLocalizedMessage(LanguageID)
                                    Else
                                        For Each tcprogramcondition As TCProgramCondition In lstTCProgramCondition.Result
                                            If (tcprogramcondition.TCProgram IsNot Nothing) Then
                                                ' Only update the program expiration date if trackable coupon expiration feature 
                                                ' is disabled or the expire type is legacy Offer End Date
                                                If ((Not bTCPExpireDateEnabled) or (tcprogramcondition.TCProgram.TCExpireType = 1))
                                                   m_TCProgram.UpdateTCProgramExpiryDate(tcprogramcondition.TCProgram.ProgramID, ProdEndDate)
                                                End If
                                            End If
                                        Next
                                    End If

                                    m_Offer.UpdateOfferDefaultGroupName(MyCommon.Extract_Val(Request.QueryString("OfferID")), 2, strQuerySupportOfferName)
                                    ' if TierLevels has changed, update the value
                                    If MyCommon.Extract_Val(Request.QueryString("tierlevels")) <> TierLevels Then
                                        MyCommon.QueryStr = "update CPE_RewardOptions set TierLevels=" & IIf(MyCommon.Extract_Val(Request.QueryString("tierlevels")) <= 0, 1, MyCommon.Extract_Val(Request.QueryString("tierlevels"))) & " " &
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
                                    If Request.QueryString("offerType") = "1" Then
                                        MyCommon.QueryStr = "select DISC.DiscountID from CPE_Discounts DISC with (NoLock) " &
                                                            "inner join CPE_Deliverables DEL with (NoLock) on DEL.OutputID=DISC.DiscountID and DEL.DeliverableTypeID=2 " &
                                                            "  and DEL.RewardOptionPhase=3 and DEL.Deleted=0 " &
                                                            "inner join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID=DEL.RewardOptionID and RO.Deleted=0 " &
                                                            "where DISC.Deleted=0 and RO.IncentiveID=" & MyCommon.Extract_Val(Request.QueryString("OfferID"))
                                        rst = MyCommon.LRT_Select
                                        If rst.Rows.Count Then
                                            For Each row In rst.Rows
                                                MyCommon.QueryStr = "update CPE_Discounts with (RowLock) set BestDeal=0 where DiscountID=" & MyCommon.NZ(row.Item("DiscountID"), 0)
                                                MyCommon.LRT_Execute()
                                            Next
                                        End If
                                    End If

                                    'Update the Currency used with this offer (rewardoption)
                                    If Not (Request.QueryString("CurrencyID") Is Nothing) Then
                                        MyCommon.QueryStr = "Update CPE_RewardOptions set CurrencyID=" & MyCommon.Extract_Val(Request.QueryString("CurrencyID")) & " where RewardOptionID=" & roid & ";"
                                        MyCommon.LRT_Execute()
                                        CurrencyID = MyCommon.Extract_Val(Request.QueryString("CurrencyID"))
                                    End If

                                    'Update the UOM's used for the offer if multi-UOM is enabled
                                    If MyCommon.Extract_Val(MyCommon.Fetch_UE_SystemOption(135)) = 1 Then
                                        ' update the UOM for each type
                                        MyCommon.QueryStr = "select UOMTypeID from UOMTypes with (NoLock);"
                                        rst = MyCommon.LRT_Select
                                        For Each row In rst.Rows
                                            If GetCgiValue("uomtype" & MyCommon.NZ(row.Item("UOMTypeID"), 0)) <> "" Then
                                                MyCommon.QueryStr = "dbo.pt_CPE_RewardOptionUOMs_Update"
                                                MyCommon.Open_LRTsp()
                                                MyCommon.LRTsp.Parameters.Add("@RewardOptionID", SqlDbType.BigInt).Value = roid
                                                MyCommon.LRTsp.Parameters.Add("@UOMTypeID", SqlDbType.Int).Value = row.Item("UOMTypeID")
                                                MyCommon.LRTsp.Parameters.Add("@UOMSubTypeID", SqlDbType.Int).Value = GetCgiValue("uomtype" & row.Item("UOMTypeID"))
                                                MyCommon.LRTsp.ExecuteNonQuery()
                                            End If
                                        Next
                                    End If
                                    If bUseDisplayDates Then
                                        MyCommon.QueryStr = "dbo.pa_UpdateOfferAccessoryFields"
                                        MyCommon.Open_LRTsp()
                                        MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = MyCommon.Extract_Val(Request.QueryString("OfferID"))
                                        MyCommon.LRTsp.Parameters.Add("@DisplayStartTime", SqlDbType.DateTime).Value = IIf(String.IsNullOrEmpty(form_DispStartdate), DBNull.Value, form_DispStartdate)
                                        MyCommon.LRTsp.Parameters.Add("@DisplayEndTime", SqlDbType.DateTime).Value = IIf(String.IsNullOrEmpty(form_DispEnddate), DBNull.Value, form_DispEnddate)

                                        MyCommon.LRTsp.Parameters.Add("@RedemThresholdPerHour", SqlDbType.BigInt).Value = DBNull.Value
                                        MyCommon.LRTsp.Parameters.Add("@AdminUserId", SqlDbType.BigInt).Value = AdminUserID
                                        MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.BigInt).Value = EngineID

                                        MyCommon.LRTsp.Parameters.Add("@OptionID", SqlDbType.BigInt).Value = 143
                                        MyCommon.LRTsp.ExecuteNonQuery()
                                        MyCommon.Close_LRTsp()
                                    End If
                                End If

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
                                    Dim form_Disallow_MutualExclusionGroups As Integer = IIf(Request.QueryString("Disallow_MutualExclusionGroups") = "on", 1, 0)
                                    Dim form_Disallow_UserDefinedFields As Integer = IIf(Request.QueryString("Disallow_UserDefinedFields") = "on", 1, 0)
                                    Dim form_Disallow_DisplayDates As Integer = 0
                                    If bUseDisplayDates Then
                                        If (Request.QueryString("Disallow_DisplayDates") = "on") Then
                                            form_Disallow_DisplayDates = 1
                                        End If
                                    End If
                                    Dim form_Disallow_RewardEvaluation As Integer = IIf(Request.QueryString("Disallow_RewardEvaluation") = "on", 1, 0)
                                    Dim form_Disallow_AdvancedOption As Integer = IIf(Request.QueryString("Disallow_AdvancedOption") = "on", 1, 0)
                                    Dim form_Disallow_PreOrder As Integer = IIf(Request.QueryString("Disallow_PreOrder") = "on", 1, 0)
                                    Dim form_Disallow_OfferType As Integer = IIf(Request.QueryString("Disallow_OfferType") = "on", 1, 0)
                                    MyCommon.QueryStr = "update TemplatePermissions with (RowLock) set Disallow_EmployeeFiltering=" & form_Disallow_EmployeeFiltering &
                                                        " , Disallow_ProductionDates=" & form_Disallow_ProductionDates &
                                                        " , Disallow_limits=" & form_Disallow_Limits &
                                                        " , Disallow_Tiers=" & form_Disallow_Tiers &
                                                        " , Disallow_Priority=" & form_Disallow_Priority &
                                                        " , Disallow_CRMEngine=" & form_Disallow_CRMEngine &
                                                        " , Disallow_ExecutionEngine=" & form_Disallow_ExecutionEngine &
                                                        " , Disallow_Sweepstakes=" & form_Disallow_Sweepstakes &
                                                        " , Disallow_MutualExclusionGroups=" & form_Disallow_MutualExclusionGroups &
                                                                  " , Disallow_UserDefinedFields=" & form_Disallow_UserDefinedFields &
                                                                  " , Disallow_RewardEvaluation=" & form_Disallow_RewardEvaluation &
                                                                  " ,Disallow_AdvancedOption=" & form_Disallow_AdvancedOption &
                                                                              " ,Disallow_PreOrder=" & form_Disallow_PreOrder &
                                                                              " ,Disallow_OfferType=" & form_Disallow_OfferType

                                    If bUseDisplayDates Then
                                        MyCommon.QueryStr = MyCommon.QueryStr & " , Disallow_DisplayDates=" & form_Disallow_DisplayDates & " where OfferID=" & OfferID
                                    Else
                                        MyCommon.QueryStr = MyCommon.QueryStr & " where OfferID=" & OfferID
                                    End If
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

                                'MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-edit", LanguageID))

                                ArrChangedVal = Request.QueryString("hdnChangedVal").Split("|")

                                If infoMessage = "" AndAlso ArrChangedVal.Length > 0 Then
                                    ComposedHist = Copient.PhraseLib.Lookup("history.offer-editgen", LanguageID)
                                    If Array.IndexOf(ArrChangedVal, "name") >= 0 OrElse Array.IndexOf(ArrChangedVal, "name_en-US") >= 0 Then
                                        ComposedHist += Copient.PhraseLib.Lookup("term.name", LanguageID) & ", "
                                    End If
                                    If Array.IndexOf(ArrChangedVal, "desc") >= 0 OrElse Array.IndexOf(ArrChangedVal, "desc_en-US") >= 0 Then
                                        ComposedHist += Copient.PhraseLib.Lookup("term.description", LanguageID) & ", "
                                    End If
                                    If Array.IndexOf(ArrChangedVal, "category") >= 0 Then
                                        ComposedHist += Copient.PhraseLib.Lookup("term.category", LanguageID) & ", "
                                    End If
                                    If Array.IndexOf(ArrChangedVal, "vendorCouponCode") >= 0 Then
                                        ComposedHist += Copient.PhraseLib.Lookup("term.vendor-coupon-code", LanguageID) & ", "
                                    End If
                                    If Array.IndexOf(ArrChangedVal, "productionstart") >= 0 Then
                                        ComposedHist += Copient.PhraseLib.Lookup("history.offerstartdate-edit", LanguageID) & ", "
                                    End If
                                    If Array.IndexOf(ArrChangedVal, "productionend") >= 0 Then
                                        ComposedHist += Copient.PhraseLib.Lookup("history.offerenddate-edit", LanguageID) & ", "
                                    End If
                                    If Array.IndexOf(ArrChangedVal, "testingstart") >= 0 Then
                                        ComposedHist += Copient.PhraseLib.Lookup("history.testingstartdate-edit", LanguageID) & ", "
                                    End If
                                    If Array.IndexOf(ArrChangedVal, "testingend") >= 0 Then
                                        ComposedHist += Copient.PhraseLib.Lookup("history.testingenddate-edit", LanguageID) & ", "
                                    End If
                                    If Array.IndexOf(ArrChangedVal, "displaystart") >= 0 Then
                                        ComposedHist += Copient.PhraseLib.Lookup("history.offerdisplaystartdate-edit", LanguageID) & ", "
                                    End If
                                    If Array.IndexOf(ArrChangedVal, "displayend") >= 0 Then
                                        ComposedHist += Copient.PhraseLib.Lookup("history.offerdisplayenddate-edit", LanguageID) & ", "
                                    End If
                                    If Array.IndexOf(ArrChangedVal, "disp-start-hr") >= 0 OrElse Array.IndexOf(ArrChangedVal, "disp-start-min") >= 0 Then
                                        ComposedHist += Copient.PhraseLib.Lookup("history.offerdisplaystarttime-edit", LanguageID) & ", "
                                    End If
                                    If Array.IndexOf(ArrChangedVal, "disp-end-hr") >= 0 OrElse Array.IndexOf(ArrChangedVal, "disp-end-min") >= 0 Then
                                        ComposedHist += Copient.PhraseLib.Lookup("history.offerdisplayendtime-edit", LanguageID) & ", "
                                    End If
                                    If Array.IndexOf(ArrChangedVal, "priority") >= 0 Then
                                        ComposedHist += Copient.PhraseLib.Lookup("term.priority", LanguageID) & ", "
                                    End If
                                    If Array.IndexOf(ArrChangedVal, "EmployeeFiltering") >= 0 OrElse Array.IndexOf(ArrChangedVal, "employeesonly") >= 0 _
                                       OrElse Array.IndexOf(ArrChangedVal, "employeesexcluded") >= 0 Then
                                        ComposedHist += Copient.PhraseLib.Lookup("history.offeremployee-edit", LanguageID) & ", "
                                    End If
                                    If Array.IndexOf(ArrChangedVal, "P3DistTimeType") >= 0 OrElse Array.IndexOf(ArrChangedVal, "limit3period") >= 0 OrElse Array.IndexOf(ArrChangedVal, "limit3") >= 0 Then
                                        ComposedHist += Copient.PhraseLib.Lookup("term.limits", LanguageID) & ", "
                                    End If
                                    If Array.IndexOf(ArrChangedVal, "tierlevels") >= 0 Then
                                        ComposedHist += Copient.PhraseLib.Lookup("term.tiers", LanguageID) & ", "
                                    End If
                                    If Array.IndexOf(ArrChangedVal, "crmengine") >= 0 OrElse Array.IndexOf(ArrChangedVal, "InboundCRMEngineID") >= 0 _
                                      OrElse Array.IndexOf(ArrChangedVal, "vendor") >= 0 Then
                                        ComposedHist += Copient.PhraseLib.Lookup("term.inbound/outbound", LanguageID) & ", "
                                    End If
                                    If Array.IndexOf(ArrChangedVal, "discEvalType0") >= 0 OrElse Array.IndexOf(ArrChangedVal, "discEvalType1") >= 0 OrElse Array.IndexOf(ArrChangedVal, "discEvalType2") >= 0 Then
                                        ComposedHist += Copient.PhraseLib.Lookup("term.discountevaluation", LanguageID) & ", "
                                    End If
                                    If Array.IndexOf(ArrChangedVal, "prorateondisplay") >= 0 OrElse Array.IndexOf(ArrChangedVal, "promotiondisplay") >= 0 _
                                       OrElse Array.IndexOf(ArrChangedVal, "autotransferable") >= 0 OrElse Array.IndexOf(ArrChangedVal, "deferToEOS") >= 0 _
                                       OrElse Array.IndexOf(ArrChangedVal, "reportingimp") >= 0 OrElse Array.IndexOf(ArrChangedVal, "reportingred") >= 0 _
                                       OrElse Array.IndexOf(ArrChangedVal, "exporttoedw") >= 0 OrElse Array.IndexOf(ArrChangedVal, "preordereligibility") >= 0 _
                                           OrElse Array.IndexOf(ArrChangedVal, "issuance") >= 0 _
                                           OrElse Array.IndexOf(ArrChangedVal, "DeferCalcToTotal") >= 0 _
                                           OrElse Array.IndexOf(ArrChangedVal, "PointsProgramWatch") >= 0 Then
                                        ComposedHist += Copient.PhraseLib.Lookup("term.advancedoptions", LanguageID) & ", "
                                    End If
                                    If UDFHistory <> "" Then
                                        ComposedHist += UDFHistory
                                    End If
                                    If ComposedHist.Contains(",") Then
                                        ComposedHist = ComposedHist.Remove(ComposedHist.LastIndexOf(","))
                                        MyCommon.Activity_Log(3, OfferID, AdminUserID, ComposedHist)
                                    End If
                                End If



                                'If the offer has an EIW condition, and the dates have changed, rerandomize the triggers
                                If HasEIW Then
                                    If (OldStartDate <> ProdStartDate) OrElse (OldEndDate <> ProdEndDate) Then
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
                        Localization.SaveTranslationInputs(MyCommon, MLI, Request.QueryString, 9, ErrorMessage)
                        If (ErrorMessage <> "") Then
                            infoMessage = ErrorMessage
                        End If
                        'Description
                        MLI.MLColumnName = "OfferDescription"
                        MLI.StandardColumnName = "Description"
                        MLI.StandardValue = Description
                        MLI.InputName = "form_Description"
                        Localization.SaveTranslationInputs(MyCommon, MLI, Request.QueryString, 9)

                        If MyCommon.Fetch_SystemOption(284) = "1" AndAlso ProdEndDate <> FolderEndDate AndAlso Not IsTemplate Then
                            MyCommon.QueryStr = "Update CPE_Incentives with (rowlock) set UserModifiedOffer = 1 where incentiveid =" & OfferID
                            MyCommon.LRT_Execute()
                        Else
                            MyCommon.QueryStr = "Update CPE_Incentives with (rowlock) set UserModifiedOffer = 0 where incentiveid =" & OfferID
                            MyCommon.LRT_Execute()
                        End If

                    End If

                    If (Request.QueryString("Deploy") <> "") AndAlso (infoMessage = "") Then
                        Send("<!-- The deploy button was clicked -->")
                        IsDeployable = IsDeployableOffer(MyCommon, OfferID, roid, ErrorMsg)
                        If (IsDeployable) Then
                            MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=2, UpdateLevel=UpdateLevel+1, LastDeployClick=getdate(), DeployDeferred=0, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", ExpireLocked=1 where IncentiveID=" & OfferID & ";"
                            MyCommon.LRT_Execute()
                            MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-deploy", LanguageID))
                            Response.Status = "301 Moved Permanently"
                            Response.AddHeader("Location", "UEoffer-sum.aspx?OfferID=" & OfferID)
                            GoTo done
                        Else
                            infoMessage = Copient.PhraseLib.Lookup(ErrorMsg, LanguageID)
                        End If
                    End If

                    If MyCommon.Fetch_SystemOption(156) = "1" Then
                        ' clear any edits that were not saved
                        MyCommon.QueryStr = "delete from OfferUDFStringValues where OfferID = " & OfferID & ";"
                        MyCommon.LRT_Execute()
                        ' clear any new adds that were not saved
                        'MyCommon.QueryStr = "delete from UserDefinedFieldsValues where StringValue is null and BooleanValue is null and DateValue is null and IntValue is null and OfferID = " & OfferID & ";"
                        'MyCommon.LRT_Execute()
                        ' if UDF marked for deletion but still exists then it was not saved, so unmark the deletion
                        MyCommon.QueryStr = "update UserDefinedFieldsValues set deleted = 0 where deleted= 1 and OfferID = " & OfferID & ";"
                        MyCommon.LRT_Execute()
                    End If

                    If (Request.QueryString("OfferID") <> "") Then
                        MyCommon.QueryStr = "Select IncentiveID, OID.EngineID, PE.PhraseID as EnginePhraseID, PEST.PhraseID as EngineSubTypePhraseID, " &
                                            "IsTemplate, FromTemplate, ClientOfferID, IncentiveName, CPE.Description, PromoClassID, CRMEngineID, Priority, " &
                                            "StartDate, EndDate, EveryDOW, EligibilityStartDate, EligibilityEndDate, TestingStartDate, TestingEndDate, P1DistQtyLimit, P1DistTimeType, P1DistPeriod, " &
                                            "P2DistQtyLimit, P2DistTimeType, P2DistPeriod, P3DistQtyLimit, P3DistTimeType, P3DistPeriod, EnableImpressRpt, EnableRedeemRpt, CPE.CreatedDate, CPE.LastUpdate, CPE.Deleted, CPEOARptDate, " &
                                            "CPEOADeploySuccessDate, CPEOADeployRpt, CRMRestricted, StatusFlag, EmployeesOnly, EmployeesExcluded, DeferCalcToEOS, ExportToEDW, " &
                                            "Favorite, OC.Description as CategoryName, SendIssuance, InboundCRMEngineID, ChargebackVendorID, ManufacturerCoupon, VendorCouponCode, AutoTransferable,PromotionDisplay,ProrateonDisplay,EnableCollisionDetection,PreOrderEligibility, CPE.EngineSubTypeID, " &
                                            "CPE.DiscountEvalTypeID,buy.ExternalBuyerId as BuyerID, DeferCalcToTotal, PointsProgramWatch, CPE.UserModifiedOffer, CPE.StoreCoupon " &
                                            "from CPE_Incentives as CPE with (NoLock) " &
                                            "left join OfferCategories as OC with (NoLock) on CPE.PromoClassID=OfferCategoryID " &
                                            "left join OfferIDs as OID with (NoLock) on OID.OfferID=CPE.IncentiveID " &
                                            "left join PromoEngines as PE with (NoLock) on PE.EngineID=OID.EngineID " &
                                            "left join PromoEngineSubTypes as PEST with (NoLock) on PEST.PromoEngineID=OID.EngineID and PEST.SubTypeID=OID.EngineSubTypeID " &
                                            "left outer join Buyers as buy with (nolock) on buy.BuyerId= cpe.BuyerId " &
                                            "where IncentiveID=" & Request.QueryString("OfferID") & ";"
                        rst = MyCommon.LRT_Select
                        If rst.Rows.Count < 1 Then
                            infoMessage = Copient.PhraseLib.Lookup("term.notfound", LanguageID)
                        Else
                            IsTemplate = MyCommon.NZ(rst.Rows(0).Item("IsTemplate"), False)
                            FromTemplate = MyCommon.NZ(rst.Rows(0).Item("FromTemplate"), False)
                            'If (MyCommon.Fetch_UE_SystemOption(168) = "1" AndAlso MyCommon.NZ(rst.Rows(0)("BuyerID"), 0) <> 0) Then
                            '    OfferName = "Buyer " + rst.Rows(0).Item("BuyerID").ToString() + " - " + MyCommon.NZ(rst.Rows(0).Item("IncentiveName"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)).ToString()
                            'Else
                            '    OfferName = MyCommon.NZ(rst.Rows(0).Item("IncentiveName"), Copient.PhraseLib.Lookup("term.unknown", LanguageID))
                            'End If
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
                            IsMfgCoupon = MyCommon.NZ(rst.Rows(0).Item("ManufacturerCoupon"), 0)
                            AutoTransferable = MyCommon.NZ(rst.Rows(0).Item("AutoTransferable"), False)
                            PreOrderEligibility = MyCommon.NZ(rst.Rows(0).Item("PreOrderEligibility"), False)   'CR8
                            EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), 0)
                            EnginePhraseID = MyCommon.NZ(rst.Rows(0).Item("EnginePhraseID"), 0)
                            EngineSubTypeID = MyCommon.NZ(rst.Rows(0).Item("EngineSubTypeID"), 0)
                            EngineSubTypePhraseID = MyCommon.NZ(rst.Rows(0).Item("EngineSubTypePhraseID"), 0)
                            VendorCouponCode = MyCommon.NZ(rst.Rows(0).Item("VendorCouponCode"), "")
                            EngineSubTypeID = MyCommon.NZ(rst.Rows(0).Item("EngineSubTypeID"), 0)
                            Priority = IIf(MyCommon.Fetch_UE_SystemOption(146) = 1, MyCommon.NZ(rst.Rows(0).Item("Priority"), 50), MyCommon.NZ(rst.Rows(0).Item("Priority"), 0))
                            DiscEvalTypeID = MyCommon.NZ(rst.Rows(0).Item("DiscountEvalTypeID"), 0)
                            IsPromotionDisplay = MyCommon.NZ(rst.Rows(0).Item("PromotionDisplay"), False)
                            IsProrateonDisplay = MyCommon.NZ(rst.Rows(0).Item("ProrateonDisplay"), False)
                            DeferCalcToTotal = MyCommon.NZ(rst.Rows(0).Item("DeferCalcToTotal"), False)
                            PointsProgramWatch = MyCommon.NZ(rst.Rows(0).Item("PointsProgramWatch"), 0)
                            IsUserModifiedOffer = MyCommon.NZ(rst.Rows(0).Item("UserModifiedOffer"), False)
                            EnableCollisionDetection = MyCommon.NZ(rst.Rows(0).Item("EnableCollisionDetection"), False)
                            StoreCoupon = MyCommon.NZ(rst.Rows(0).Item("StoreCoupon"), 0)

                            ' determine if these are custom limits
                            OncePerOfferLimit = MyCommon.NZ(rst.Rows(0).Item("P3DistTimeType"), 0) = 1 _
                                                AndAlso MyCommon.NZ(rst.Rows(0).Item("P3DistQtyLimit"), 0) = 1 _
                                                AndAlso MyCommon.NZ(rst.Rows(0).Item("P3DistPeriod"), 0) = 3650

                            NoLimit = MyCommon.NZ(rst.Rows(0).Item("P3DistTimeType"), 0) = 2 _
                                      AndAlso MyCommon.NZ(rst.Rows(0).Item("P3DistQtyLimit"), 0) = 0 _
                                      AndAlso MyCommon.NZ(rst.Rows(0).Item("P3DistPeriod"), 0) = 0

                            HourLimit = MyCommon.NZ(rst.Rows(0).Item("P3DistTimeType"), 0) = 1 _
                                        AndAlso MyCommon.NZ(rst.Rows(0).Item("P3DistPeriod"), 0) < 1
                            'load up any selected saved or defaults for the UOM subtype
                            MyCommon.QueryStr = "select UOMT.UOMTypeID, UOMT.DefaultUOMSubTypeID, UOMSubTypeID as SelectedSubTypeID " &
                                                "from UOMTypes as UOMT with (NoLock) " &
                                                "left join CPE_RewardOptionUOMs as ROU with (NoLock) " &
                                                "  on ROU.UOMTypeID = UOMT.UOMTypeID " &
                                                "  and RewardOptionID =" & roid
                            rst2 = MyCommon.LRT_Select
                            For Each row2 In rst2.Rows
                                If Not UOMSubTypes.ContainsKey(row2.Item("UOMTypeID")) Then
                                    ' if no sub type is assigned to the offer, then set it to unspecified (0).
                                    If IsDBNull(row2.Item("SelectedSubTypeID")) Then
                                        UOMSubTypes.Add(row2.Item("UOMTypeID"), 0)
                                    Else
                                        UOMSubTypes.Add(row2.Item("UOMTypeID"), MyCommon.NZ(row2.Item("SelectedSubTypeID"), 0))
                                    End If
                                End If
                            Next
                            If bUseDisplayDates Then
                                MyCommon.QueryStr = "SELECT DisplayStartDate,DisplayEndDate FROM OfferAccessoryFields with (NoLock) WHERE OfferID = " & Request.QueryString("OfferID") & ";"
                                dtODisp = MyCommon.LRT_Select()
                            End If
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
                                    Disallow_MutualExclusionGroups = MyCommon.NZ(rowTemplates.Item("Disallow_MutualExclusionGroups"), True)
                                    Disallow_UserDefinedFields = MyCommon.NZ(rowTemplates.Item("Disallow_UserDefinedFields"), True)
                                    If bUseDisplayDates Then
                                        Disallow_DisplayDates = MyCommon.NZ(rowTemplates.Item("Disallow_DisplayDates"), True)
                                    End If
                                    Disallow_RewardEvaluation = MyCommon.NZ(rowTemplates.Item("Disallow_RewardEvaluation"), True)
                                    Disallow_AdvancedOption = MyCommon.NZ(rowTemplates.Item("Disallow_AdvancedOption"), True)
                                    Disallow_PreOrder = MyCommon.NZ(rowTemplates.Item("Disallow_PreOrder"), True)
                                    Disallow_OfferType = MyCommon.NZ(rowTemplates.Item("Disallow_OfferType"), True)

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
                                Disallow_MutualExclusionGroups = False
                                Disallow_DisplayDates = False
                                Disallow_UserDefinedFields = False
                                Disallow_RewardEvaluation = False
                                Disallow_AdvancedOption = False
                                Disallow_PreOrder = False
                                Disallow_OfferType = False
                            End If
                        End If
                    End If

                    'Check that the External OfferID can be changed
                    If IsMfgCoupon = 1 Or InboundCRMEngineID = 1 Or InboundCRMEngineID = 2 Then
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
                    Send_Scripts(New String() {"datePicker.js", "popup.js", "jquery.min.js"})
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
  #UDFfadeDiv
  {
    background-color: #e0e0e0;
    position: absolute;
    top: 0px;
    left: 0px;
    width: 5000px;
    height: 5000px;
    z-index: 99;
    display: none;
    opacity: .4;
    filter: alpha(opacity=40);
  }
</style>
<script type="text/javascript">
window.name = "UEofferGen"
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
      && el.id!="eligibility-start-picker" && el.id!="eligibility-end-picker"
      && el.id!="testing-start-picker" && el.id!="testing-end-picker"
      && el.id!="display-start-picker" && el.id!="display-end-picker" && el.id!="udf-datevalue-picker") {
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

function handleOnSubmit() {
  var retVal = false;
  
  retVal = checkdesclengthdata() && ValidateDispDates() &&  ValidateTimes() && promptForDeploy();
  
  return retVal;
}

function ValidateDispDates() {
 var elemDispStart = document.getElementById("displaystart");
 var elemDispEnd = document.getElementById("displayend");
 var retVal = true;
   <% If bUseDisplayDates Then%>
       if (retVal == true && elemDispStart != null && elemDispStart.value != "") {
          retVal = IsValidLocalizedDate(elemDispStart.value, '<% Sendb(MyCommon.GetAdminUser.Culture.DateTimeFormat.ShortDatePattern) %>');
        if (!retVal){
           elemDispStart.focus();
           retVal = false;
          }
        }
        if (retVal == true && elemDispEnd != null && elemDispEnd.value != "") {
         retVal = IsValidLocalizedDate(elemDispEnd.value, '<% Sendb(MyCommon.GetAdminUser.Culture.DateTimeFormat.ShortDatePattern) %>');
        if (!retVal){
             elemDispEnd.focus();
             retVal = false;
        }
        }
    <% End If %>
  return retVal;
}

function ValidateTimes() {
  var elemTestStartHr = document.getElementById("test-start-hr");
  var elemTestStartMin = document.getElementById("test-start-min");
  var elemTestEndHr = document.getElementById("test-end-hr");
  var elemTestEndMin = document.getElementById("test-end-min");
  var elemProdStartHr = document.getElementById("prod-start-hr");
  var elemProdStartMin = document.getElementById("prod-start-min");
  var elemProdEndHr = document.getElementById("prod-end-hr");
  var elemProdEndMin = document.getElementById("prod-end-min");
  var elemDispStartHr = document.getElementById("disp-start-hr");
  var elemDispStartMin = document.getElementById("disp-start-min");
  var elemDispEndHr = document.getElementById("disp-end-hr");
  var elemDispEndMin = document.getElementById("disp-end-min");
  var retVal = true;

  <% If bAllowTimeWithStartEndDates Then%>
  if (retVal == true && elemTestStartHr != null && (!isInteger(elemTestStartHr.value) || (parseInt(elemTestStartHr.value) < 0) || (parseInt(elemTestStartHr.value) > 23))) {
		alert('<%Sendb(Copient.PhraseLib.Lookup("offer-gen.InvalidTestStart", LanguageID)) %>');
    retVal = false;
  }
  if (retVal == true && elemTestStartMin != null && (!isInteger(elemTestStartMin.value) || (parseInt(elemTestStartMin.value) < 0) || (parseInt(elemTestStartMin.value) > 59))) {
		alert('<%Sendb(Copient.PhraseLib.Lookup("offer-gen.InvalidStartMinute", LanguageID)) %>');
    retVal = false;
  }
  if (retVal == true && elemTestEndHr != null && (!isInteger(elemTestEndHr.value) || (parseInt(elemTestEndHr.value) < 0) || (parseInt(elemTestEndHr.value) > 23))) {
		alert('<%Sendb(Copient.PhraseLib.Lookup("offer-gen.InvalidEndHour", LanguageID)) %>');
    retVal = false;
  }
  if (retVal == true && elemTestEndMin != null && (!isInteger(elemTestEndMin.value) || (parseInt(elemTestEndMin.value) < 0) || (parseInt(elemTestEndMin.value) > 59))) {
		alert('<%Sendb(Copient.PhraseLib.Lookup("offer-gen.InvalidEndMinute", LanguageID)) %>');
    retVal = false;
  }
  if (retVal == true && elemProdStartHr != null && (!isInteger(elemProdStartHr.value) || (parseInt(elemProdStartHr.value) < 0) || (parseInt(elemProdStartHr.value) > 23))) {
		alert('<%Sendb(Copient.PhraseLib.Lookup("offer-gen.InvalidProdStartHour", LanguageID)) %>');
    retVal = false;
  }
  if (retVal == true && elemProdStartMin != null && (!isInteger(elemProdStartMin.value) || (parseInt(elemProdStartMin.value) < 0) || (parseInt(elemProdStartMin.value) > 59))) {
		alert('<%Sendb(Copient.PhraseLib.Lookup("offer-gen.InvalidProdStartMinute", LanguageID)) %>');
    retVal = false;
  }
  if (retVal == true && elemProdEndHr != null && (!isInteger(elemProdEndHr.value) || (parseInt(elemProdEndHr.value) < 0) || (parseInt(elemProdEndHr.value) > 23))) {
		alert('<%Sendb(Copient.PhraseLib.Lookup("offer-gen.InvalidProdEndHour", LanguageID)) %>');
    retVal = false;
  }
  if (retVal == true && elemProdEndMin != null && (!isInteger(elemProdEndMin.value) || (parseInt(elemProdEndMin.value) < 0) || (parseInt(elemProdEndMin.value) > 59))) {
		alert('<%Sendb(Copient.PhraseLib.Lookup("offer-gen.InvalidProdEndMinute", LanguageID)) %>');
    retVal = false;
  }
  <% End If %>

  <% If bUseDisplayDates Then%>
    if (retVal == true && elemDispStartHr != null && (!isInteger(elemDispStartHr.value) || (parseInt(elemDispStartHr.value) < 0) || (parseInt(elemDispStartHr.value) > 23))) {
		alert('<%Sendb(Copient.PhraseLib.Lookup("offer-gen.InvalidDispStartHour", LanguageID)) %>');
    retVal = false;
  }
  if (retVal == true && elemDispStartMin != null && (!isInteger(elemDispStartMin.value) || (parseInt(elemDispStartMin.value) < 0) || (parseInt(elemDispStartMin.value) > 59))) {
		alert('<%Sendb(Copient.PhraseLib.Lookup("offer-gen.InvalidDispStartMinute", LanguageID)) %>');
    retVal = false;
  }
  if (retVal == true && elemDispEndHr != null && (!isInteger(elemDispEndHr.value) || (parseInt(elemDispEndHr.value) < 0) || (parseInt(elemDispEndHr.value) > 23))) {
		alert('<%Sendb(Copient.PhraseLib.Lookup("offer-gen.InvalidDispEndHour", LanguageID)) %>');
    retVal = false;
  }
  if (retVal == true && elemDispEndMin != null && (!isInteger(elemDispEndMin.value) || (parseInt(elemDispEndMin.value) < 0) || (parseInt(elemDispEndMin.value) > 59))) {
		alert('<%Sendb(Copient.PhraseLib.Lookup("offer-gen.InvalidDispEndMinute", LanguageID)) %>');
    retVal = false;
  }
  <% End If %>
  return retVal;
}

function setLimits(){
  var elem = document.getElementById('P3DistTimeType');
  var periodRow = document.getElementById('p3row2');
  var limitRow = document.getElementById('p3row3');
  var elemPeriod = document.getElementById('limit3period');
  
  
  if (limitRow != null) {
    if (elem != null && elem.value=='2') {
      periodRow.style.display = '';
      limitRow.style.display = 'none';
      if (elemPeriod!=null) elemPeriod.value = '';
    } else if (elem != null && (elem.value=='<%Sendb(LIMIT_ONCE_PER_OFFER) %>' || elem.value=='<%Sendb(LIMIT_NONE) %>')) {
      limitRow.style.display = 'none';
      periodRow.style.display = 'none';
  } else {
      limitRow.style.display = '';    
      periodRow.style.display = '';
    }
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

function updatehistory() {
    
	var control, controls = document.mainform.elements;
    var tagName, type;
   	var changeVal = "";
	
    for (var i=0, iLen=controls.length; i<iLen; i++) {
      control = controls[i];
      tagName = control.tagName.toLowerCase();
      type = control.type;
        
      // textarea
      if (tagName == 'textarea') {

        if (control.value != control.defaultValue) {
          changeVal = changeVal + control.id + "|";
		 
        }

      // input
      } else if (tagName == 'input') {
            
        // text
        if (type == 'text') {

          if (control.value != control.defaultValue) {

			changeVal = changeVal + control.id + "|";
          }

        // radio and checkbox
        } else if (type == 'radio' || type == 'checkbox') {
          if (control.checked != control.defaultChecked) {
            changeVal = changeVal + control.id + "|";
          }
        }

      // select multiple and single
      } else if (tagName == 'select') {
	  
        var c = false;
		  var options = control.options;
		  var def = 0, o, ol, opt;
          
          for (o = 0, ol = options.length; o < ol; o++) {
	        opt = options[o];
			
	        c = c || (opt.selected != opt.defaultSelected);
	        if (opt.defaultSelected) def = o;
          }
          if (c && !control.multiple) c = (def != control.selectedIndex);

          if (c) changeVal = changeVal + control.id + "|";		 
      }
    }
	
    document.getElementById("hdnChangedVal").value = changeVal;
	  var obj = document.getElementById("eligibleRow").currentStyle['display'];
    if(obj == 'none')
    {
	    document.getElementById("hdnEmployeeRow").value = 'false';
    }
    else
    {
        document.getElementById("hdnEmployeeRow").value = 'true';
    }
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
          If MyCommon.Fetch_UE_SystemOption(80) = 1 Then
            Send("if (retVal = true) {")
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
  
  //document.getElementById("tools").innerHTML = "<div class=\"loading\"><img id=\"clock\" src=\"/images/clock22.png\" \/><br \/>" + '<% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID)) %><\/div>';
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
function replaceAll(find, replace, str) 
    {
      while( str.indexOf(find) > -1)
      {
        str = str.replace(find, replace);
      }
      return str;
    }
	
function controllengthstring(pvalue) {
var encodestrring=pvalue;
	encodestrring = encodeURI(encodestrring); 
	encodestrring = replaceAll("&","%26",encodestrring);
	encodestrring = replaceAll("+","%2B",encodestrring);
	encodestrring = replaceAll("#","%23",encodestrring);
return 	encodestrring;
}	

function checkdesclength(cntrl){
var elemvalue = document.getElementById("desc").value;
if (elemvalue.length <= 1000) {
xmlhttpPost_OfferDescription('../OfferFeeds.aspx', 'AllowSpecialCharactersUE');
}
}

function getQueryStringOfferDesc(mode) {
  var elemvalue = document.getElementById("desc").value;
  elemvalue = controllengthstring(elemvalue);
  return  "Mode=" + mode + "&OfferID=<%Sendb(OfferID)%>" + "&OfferDescription="+ elemvalue;
}

function xmlhttpPost_OfferDescription(strURL, mode) {
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
  self.xmlHttpReq.open('POST', strURL, true);
  self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
  self.xmlHttpReq.send(getQueryStringOfferDesc(mode));
  self.xmlHttpReq.onreadystatechange = function() {
  if (self.xmlHttpReq != null && self.xmlHttpReq.readyState == 4) {
  //updatePage(self.xmlHttpReq.responseText);
  }
  }
}

function checkdesclengthdata(){
var elemvalue = document.getElementById("desc").value;
if (elemvalue.length > 1000) {
alert('<%Sendb(Copient.PhraseLib.Lookup("error.description", LanguageID)) %>');
return false;
}else
{
document.getElementById("desc").value;
//try
//{
// document.getElementById("desc_en-US").value = "";
//}
//catch(err)
// {
//}
        return true;
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
<form id="mainform" name="mainform" action="UEoffer-gen.aspx" method="get" onsubmit="handleFormElements(this, false);return handleOnSubmit();">
  <input type="hidden" name="OfferID" id="OfferID" value="<%Sendb(OfferID)%>" />
  <input type="hidden" id="form_OfferID" name="form_OfferID" value="<%Sendb(OfferID) %>" />
  <input type="hidden" name="IsActive" id="IsActive" value="<%Sendb(IIf(StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_ACTIVE, "true", "false")) %>" />
  <input type="hidden" name="IsTemplate" id="IsTemplate" value="<% Sendb(IsTemplateVal)%>" />
  <input type="hidden" name="Popup" id="Popup" value="<%Sendb(IIf(Popup, 1, 0)) %>" />
  <input type="hidden" name="Deploy" id="Deploy" value="" />
  <input type="hidden" name="SelectedUDF" id="SelectedUDF" value="" />   
  <input type="hidden" id="hdnChangedVal" name="hdnChangedVal"/>
  <input type="hidden" id="hdnEmployeeRow" name="hdnEmployeeRow" />
  <input type="hidden" id="savedTime" name="savedTime" value="<%=DateTime.Now()%>" />
  <div id="<% Sendb(IntroID)%>">
    <% 
    Dim Name As String = ""
    If rst.Rows.Count > 0 Then
      If (MyCommon.Fetch_UE_SystemOption(168) = "1" AndAlso MyCommon.NZ(rst.Rows(0).Item("BuyerID"), "") <> "") Then
        Name = "Buyer " + rst.Rows(0).Item("BuyerID").ToString() + " - " + MyCommon.NZ(rst.Rows(0).Item("IncentiveName"), "").ToString()
      Else
        Name = MyCommon.NZ(rst.Rows(0).Item("IncentiveName"), Copient.PhraseLib.Lookup("term.unknown", LanguageID))
      End If
      If (IsTemplate) Then
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & ": " & MyCommon.TruncateString(Name, 43) & "</h1>")
      Else
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & ": " & MyCommon.TruncateString(Name, 43) & "</h1>")
      End If
    End If
  %>
  <div id="controls">
    <%
      Dim m_EditOfferRegardlessOfBuyer = Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID)
      If Not IsTemplate Then
            If (Logix.UserRoles.EditOffer And Not IsOfferWaitingForApproval(OfferID) And m_EditOfferRegardlessOfBuyer AndAlso (Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
          Send_Save("onclick=""updatehistory()""")
        End If
      Else
        If (Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer AndAlso (Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
            Send_Save("onclick=""updatehistory()""")
          End If
        End If
        If MyCommon.Fetch_SystemOption(75) Then
          If (OfferID > 0 AndAlso Logix.UserRoles.AccessNotes AndAlso Not Popup AndAlso (Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse  bOfferEditable)) Then
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
        
    If MyCommon.Fetch_SystemOption(191) AndAlso NumberofFolders > 1 And infoMessage = "" Then
      infoMessage = "Offer cannot have more than one Folder associated"
    End If
            If (infoMessage <> "") Then
                Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
            End If
            If (infoMsg <> "") Then
                Send("<div id=""infobar"" class=""orange-background"">" & infoMsg & "</div>")
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
      <div class="box" id="identification" style="z-index:50;">
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
                Send(Copient.PhraseLib.Lookup("term.none", LanguageID) & " <small><a href=""UEoffer-gen.aspx?OfferID=" & OfferID & "&amp;ExtOffer=" & OfferID & "&amp;mode=ChangeExtID"">(" & Copient.PhraseLib.Lookup("CPEoffer-gen.xid-add", LanguageID) & ")</a></small>")
              Else
                Send(ExtOfferID & " <small><a href=""UEoffer-gen.aspx?OfferID=" & OfferID & "&amp;ExtOffer=0&amp;mode=ChangeExtID"">(" & Copient.PhraseLib.Lookup("CPEoffer-gen.xid-rem", LanguageID) & ")</a></small>")
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
            MLI.Disabled = IsOfferWaitingForApproval(OfferID)
          Send(Localization.SendTranslationInputs(MyCommon, MLI, Request.QueryString, 9))
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
            MLI.Disabled = IsOfferWaitingForApproval(OfferID)
          If (bEnableRestrictedAccessToUEOfferBuilder And isTranslatedOffer) Then
            MLI.Disabled = True
          End If
          Send(Localization.SendTranslationInputs(MyCommon, MLI, Request.QueryString, 9))
        %>
        <br class="half" />
      <small>
        <%Send("(" & Copient.PhraseLib.Lookup("CPEoffergen.description", LanguageID) & ")")%></small>
      <br />
      <% If MyCommon.Fetch_UE_SystemOption(206) = "1" Then%>
      <br />
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
        <% End If %>
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
            <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
          </label>
        </span>
        <br class="printonly" />
        <% End If%>
        <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.priority", LanguageID)) %>">
          <tr>
            <td style="width: 120px;">
              <label for="priority">
                <% Sendb(Copient.PhraseLib.Lookup("CPEoffer-gen.priority", LanguageID) & ":")%>
              </label>
            </td>
            <td>
              <select id="priority" name="priority" <% if(FromTemplate and disallow_priority)then sendb(" disabled=""disabled""") %>>
                <%
                  MyCommon.QueryStr = "select PriorityID, Name, PhraseID from UE_Priorities with (NoLock);"
                  rst2 = MyCommon.LRT_Select
                  For Each row2 In rst2.Rows
                    Send("<option value=""" & MyCommon.NZ(row2.Item("PriorityID"), 0) & """" & IIf(Priority = MyCommon.NZ(row2.Item("PriorityID"), 0), " selected=""selected""", "") & ">" & _
                         "" & Copient.PhraseLib.Lookup(MyCommon.NZ(row2.Item("PhraseID"), 0), LanguageID, MyCommon.NZ(row2.Item("Name"), "")) & " </option>")
                  Next
                %>
              </select>
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
            <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
          </label>
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
                StartDT = MyCommon.NZ(rst.Rows(0).Item("StartDate"), "1/1/1900")
                EndDT = MyCommon.NZ(rst.Rows(0).Item("EndDate"), "1/1/1900")
                If StartDT = "1/1/1900" Then
                  ShortStartDate = ""
                Else
                  ShortStartDate = Logix.ToShortDateString(StartDT, MyCommon)
                End If
                If EndDT = "1/1/1900" Then
                  ShortEndDate = ""
                Else
                  ShortEndDate = Logix.ToShortDateString(EndDT, MyCommon)
                End If
                If bAllowTimeWithStartEndDates Then
                  ProdStartHr = StartDT.ToString(sHourOnlyFormat)
                  ProdStartMin = StartDT.ToString(sMinutesOnlyFormat)
                  ProdEndHr = EndDT.ToString(sHourOnlyFormat)
                  ProdEndMin = EndDT.ToString(sMinutesOnlyFormat)
                  'If (ProdEndHr = 0 And ProdEndMin = 0) Then
                   ' ProdEndHr = 23
                   ' ProdEndMin = 59
                  'End If
                End If
              Else
                ShortStartDate = ""
                ShortEndDate = ""
                If bAllowTimeWithStartEndDates Then
                  ProdStartHr = ""
                  ProdEndHr = ""
                  ProdStartMin = ""
                  ProdEndMin = ""
                End If
              End If
            %>
            <span>
            <%
              If bAllowTimeWithStartEndDates Then
                Sendb(Copient.PhraseLib.Lookup("term.enter-datetime", LanguageID))
                Sendb("<br class=""half"" />")
                Sendb("<br class=""half"" />")
              End If
            %>
            </span>
            <label for="productionstart">
              <% Sendb(Copient.PhraseLib.Lookup("term.production", LanguageID))%>
              :</label><br />
            <input type="text" class="short" id="productionstart" name="productionstart" maxlength="10"
              value="<% sendb(ShortStartDate) %>" <% if((FromTemplate and disallow_productiondates) Or IsEIWDateLocked) then sendb(" disabled=""disabled""") %> />
            <img src="/images/calendar.png" class="calendar" id="production-start-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
              title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('productionstart', event);" />
            <%If bAllowTimeWithStartEndDates Then%>
              <input class="shortest" id="prod-start-hr" maxlength="2" name="form_ProdStartHr"
               type="text" value="<% sendb(ProdStartHr)%>" <% if((FromTemplate and disallow_productiondates) Or IsEIWDateLocked) then sendb(" disabled=""disabled""") %> />:<input
              class="shortest" id="prod-start-min" maxlength="2" name="form_ProdStartMin" type="text"
               value="<% sendb(ProdStartMin)%>" <% if((FromTemplate and disallow_productiondates) Or IsEIWDateLocked) then sendb(" disabled=""disabled""") %> />
            <% End If%>
            <% Sendb(Copient.PhraseLib.Lookup("term.to", LanguageID))%>
            <input type="text" class="short" id="productionend" name="productionend" maxlength="10"
              value="<% sendb(ShortEndDate) %>" <% if((FromTemplate and disallow_productiondates) Or IsEIWDateLocked) then sendb(" disabled=""disabled""") %> />
            <img src="/images/calendar.png" class="calendar" id="production-end-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
              title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="autoDatePicker('productionend', event, <% sendb(selectDatePicker) %>);" />
            <%If bAllowTimeWithStartEndDates Then%>
              <input class="shortest" id="prod-end-hr" maxlength="2" name="form_ProdEndHr"
               type="text" value="<% sendb(ProdEndHr)%>" <% if((FromTemplate and disallow_productiondates) Or IsEIWDateLocked) then sendb(" disabled=""disabled""") %> />:<input
              class="shortest" id="prod-end-min" maxlength="2" name="form_ProdEndMin" type="text"
               value="<% sendb(ProdEndMin)%>" <% if((FromTemplate and disallow_productiondates) Or IsEIWDateLocked) then sendb(" disabled=""disabled""") %> />
            <% Else%>
            <% Sendb("(" & MyCommon.GetAdminUser.Culture.DateTimeFormat.ShortDatePattern & ")")%>
            <% End If%>
            <br />
            <br class="half" />
          </td>
        </tr>
          <tr id="eligibleRow" style="display: none;">
            <!-- BZ2079: UE-feature-removal #22: Hiding eligibility and testing dates.  To restore, remove the style attribute from this TR -->
            <td>
              <%
                If rst.Rows.Count > 0 Then
                  ShortEligStartDate = ""
                  ShortEligEndDate = ""
                  EligStartDT = MyCommon.NZ(rst.Rows(0).Item("EligibilityStartDate"), "1/1/1900")
                  EligEndDT = MyCommon.NZ(rst.Rows(0).Item("EligibilityEndDate"), "1/1/1900")
                  If EligStartDT <> "1/1/1900" Then ShortEligStartDate = Logix.ToShortDateString(EligStartDT, MyCommon)
                  If EligEndDT <> "1/1/1900" Then ShortEligEndDate = Logix.ToShortDateString(EligEndDT, MyCommon)
                Else
                  ShortEligStartDate = ""
                  ShortEligEndDate = ""
                End If
              %>
              <label for="eligibilitystart">
                <% Sendb(Copient.PhraseLib.Lookup("term.eligibility", LanguageID))%>
                :</label><br />
              <input type="text" class="short" id="eligibilitystart" name="eligibilitystart" maxlength="10"
                value="<% sendb(ShortEligStartDate) %>" <% if((FromTemplate and disallow_productiondates) Or IsEIWDateLocked) then sendb(" disabled=""disabled""") %> />
              <img src="/images/calendar.png" class="calendar" id="eligibility-start-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
                title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('eligibilitystart', event);" />
              <% Sendb(Copient.PhraseLib.Lookup("term.to", LanguageID))%>
              <input type="text" class="short" id="eligibilityend" name="eligibilityend" maxlength="10"
                value="<% sendb(ShortEligEndDate) %>" <% if((FromTemplate and disallow_productiondates) Or IsEIWDateLocked) then sendb(" disabled=""disabled""") %> />
              <img src="/images/calendar.png" class="calendar" id="eligibility-end-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
                title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('eligibilityend', event);" />
              (<% Sendb(Copient.PhraseLib.Lookup("term.mmddyyyy", LanguageID))%>)<br />
              <br class="half" />
            </td>
          </tr>
          <tr>
            <td>
              <%
                  If rst.Rows.Count > 0 Then
                      TestStartDT = MyCommon.NZ(rst.Rows(0).Item("TestingStartDate"), "1/1/1900")
                      TestEndDT = MyCommon.NZ(rst.Rows(0).Item("TestingEndDate"), "1/1/1900")
                      If TestStartDT = "1/1/1900" Then
                          ShortTestStartDate = ""
                      Else
                          ShortTestStartDate = Logix.ToShortDateString(TestStartDT, MyCommon)
                      End If
                      If TestEndDT = "1/1/1900" Then
                          ShortTestEndDate = ""
                      Else
                          ShortTestEndDate = Logix.ToShortDateString(TestEndDT, MyCommon)
                      End If
                      If bAllowTimeWithStartEndDates Then
                          TestStartHr = TestStartDT.ToString(sHourOnlyFormat)
                          TestStartMin = TestStartDT.ToString(sMinutesOnlyFormat)
                          TestEndHr = TestEndDT.ToString(sHourOnlyFormat)
                          TestEndMin = TestEndDT.ToString(sMinutesOnlyFormat)
                          'If (TestEndHr = 0 And TestEndMin = 0) Then
                          '  TestEndHr = 23
                          '  TestEndMin = 59
                          'End If
                      End If
                  Else
                      ShortTestStartDate = ""
                      ShortTestEndDate = ""
                      If bAllowTimeWithStartEndDates Then
                          TestStartHr = ""
                          TestEndHr = ""
                          TestStartMin = ""
                          TestEndMin = ""
                      End If
                  End If
            %>
            <label for="testingstart">
              <% Sendb(Copient.PhraseLib.Lookup("term.testing", LanguageID))%>
              :</label><br />
            <input type="text" class="short" id="testingstart" name="testingstart" maxlength="10"
              value="<% sendb(ShortTestStartDate) %>" <% if((FromTemplate and disallow_productiondates) Or IsEIWDateLocked) then sendb(" disabled=""disabled""") %> />
            <img src="/images/calendar.png" class="calendar" id="testing-start-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
              title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('testingstart', event);" />
            <%If bAllowTimeWithStartEndDates Then%>
              <input class="shortest" id="test-start-hr" maxlength="2" name="form_TestStartHr"
               type="text" value="<% sendb(TestStartHr)%>" <% if((FromTemplate and disallow_productiondates) Or IsEIWDateLocked) then sendb(" disabled=""disabled""") %> />:<input
              class="shortest" id="test-start-min" maxlength="2" name="form_TestStartMin" type="text"
               value="<% sendb(TestStartMin)%>" <% if((FromTemplate and disallow_productiondates) Or IsEIWDateLocked) then sendb(" disabled=""disabled""") %> />
            <% End If%>

            <% Sendb(Copient.PhraseLib.Lookup("term.to", LanguageID))%>
            <input type="text" class="short" id="testingend" name="testingend" maxlength="10"
              value="<% sendb(ShortTestEndDate) %>" <% if((FromTemplate and disallow_productiondates) Or IsEIWDateLocked) then sendb(" disabled=""disabled""") %> />
            <img src="/images/calendar.png" class="calendar" id="testing-end-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
              title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('testingend', event);" />
            <%If bAllowTimeWithStartEndDates Then%>
              <input class="shortest" id="test-end-hr" maxlength="2" name="form_TestEndHr"
               type="text" value="<% sendb(TestEndHr)%>" <% if((FromTemplate and disallow_productiondates) Or IsEIWDateLocked) then sendb(" disabled=""disabled""") %> />:<input
              class="shortest" id="test-end-min" maxlength="2" name="form_TestEndMin" type="text"
               value="<% sendb(TestEndMin)%>" <% if((FromTemplate and disallow_productiondates) Or IsEIWDateLocked) then sendb(" disabled=""disabled""") %> />
            <% Else%>
            <% Sendb("(" & MyCommon.GetAdminUser.Culture.DateTimeFormat.ShortDatePattern & ")")%>
            <% End If%>
            <br />
          </td>
          </tr>
      <% If (Not IsTemplate) Then%>
		  <tr>
        <td>
        <input type="checkbox" name="usermodoffer" id="usermodoffer" <% if (IsUserModifiedOffer) then sendb(" checked=""checked""") %>
          <%Sendb(IIF((MyCommon.Fetch_SystemOption(284) = "1")," style=""visibility:visible""","style=""visibility:hidden""")) %> />
             <%  If (MyCommon.Fetch_SystemOption(284) = "1") Then%>
                  <label for="usermodoffer">
                  <% Sendb(Copient.PhraseLib.Lookup("term.usermodoffer", LanguageID))%>
                </label>
             <% End If%>
        </td>
        </tr>
       <% End If%>
        </table>
        <hr class="hidden" />
      </div>
       <% If bUseDisplayDates Then%>
        <div class="box" id="displaydates">
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.displaydates", LanguageID))%>
                </span>
            </h2>
            <% If (IsTemplate) Then%>
            <span class="temp">
                <input type="checkbox" class="tempcheck" id="Disallow_DisplayDates" name="Disallow_DisplayDates"
                    <% if(disallow_displayDates)then send(" checked=""checked""") %> />
                <label for="Disallow_DisplayDates">
                    <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
                </label>
            </span>
            <br class="printonly" />
            <% End If%>
            <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.displaydates", LanguageID))%>">
                <tr>
                    <td>
                        <%  
                 
                            If dtODisp.Rows.Count > 0 Then
                                tempDateTime = MyCommon.NZ(dtODisp.Rows(0).Item("DisplayStartDate"), Nothing)
                                If tempDateTime <> Nothing Then
                                    DispStartDate = tempDateTime.ToString(sDateOnlyFormat)
                                    DispStartHr = tempDateTime.ToString(sHourOnlyFormat)
                                    DispStartMin = tempDateTime.ToString(sMinutesOnlyFormat)
                                End If
                        
                                tempDateTime = MyCommon.NZ(dtODisp.Rows(0).Item("DisplayEndDate"), Nothing)
                                If tempDateTime <> Nothing Then
                                    DispEndDate = tempDateTime.ToString(sDateOnlyFormat)
                                    DispEndHr = tempDateTime.ToString(sHourOnlyFormat)
                                    DispEndMin = tempDateTime.ToString(sMinutesOnlyFormat)
                                    If (DispEndHr = 0 And DispEndMin = 0) Then
                                        DispEndHr = 23
                                        DispEndMin = 59
                                    End If
                                End If
                      
                            End If
                        %>
                        <span>
                            <%
                                Sendb(Copient.PhraseLib.Lookup("term.enter-datetime", LanguageID))
                            %>
                        </span>
                        <br class="half" />
                        <br class="half" />
                        <label for="displaystart">
                            <% Sendb(Copient.PhraseLib.Lookup("term.display", LanguageID))%>
                            :</label><br />
                        <input type="text" class="short" id="displaystart" name="displaystart" maxlength="10"
                            value="<% sendb(DispStartDate) %>" <% if(FromTemplate and disallow_displayDates) then sendb(" disabled=""disabled""") %> />
            <img src="/images/calendar.png" class="calendar" id="display-start-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
              title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('displaystart', event);" />
            <input class="datesshortest" id="disp-start-hr" maxlength="2" name="form_DispStartHr"
              type="text" value="<% sendb(DispStartHr)%>" <% if(FromTemplate and Disallow_DisplayDates)then sendb(" disabled=""disabled""") %> />:<input
                class="datesshortest" id="disp-start-min" maxlength="2" name="form_DispStartMin"
                type="text" value="<% sendb(DispStartMin)%>" <% if(FromTemplate and Disallow_DisplayDates)then sendb(" disabled=""disabled""") %> />
            <% Sendb(Copient.PhraseLib.Lookup("term.to", LanguageID))%>
            <input type="text" class="short" id="displayend" name="displayend" maxlength="10"
              value="<% sendb(DispEndDate) %>" <% if(FromTemplate and disallow_displayDates) then sendb(" disabled=""disabled""") %> />
            <img src="/images/calendar.png" class="calendar" id="display-end-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
              title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('displayend', event);" />
            <input class="datesshortest" id="disp-end-hr" maxlength="2" name="form_DispEndHr"
              type="text" value="<% sendb(DispEndHr)%>" <% if(FromTemplate and Disallow_DisplayDates)then sendb(" disabled=""disabled""") %> />:<input
                                class="datesshortest" id="disp-end-min" maxlength="2" name="form_DispEndMin" type="text"
                                value="<% sendb(DispEndMin)%>" <% if(FromTemplate and Disallow_DisplayDates)then sendb(" disabled=""disabled""") %> />
                        <br class="half" />
                    </td>
                </tr>
            </table>
            <hr class="hidden" />
        </div>
        <% End If%>
      <div id="datepicker" class="dpDiv">
      </div>
      <%
        If Request.Browser.Type = "IE6" Then
          Send("<iframe src=""javascript:'';"" id=""calendariframe"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no"" style=""visibility:hidden;display:none;""></iframe>")
        End If
      %>
      <div class="box" id="uom">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.localization", LanguageID))%>
          </span>
        </h2>
        <%
          'see if multi-currency usage is enabled
          If MyCommon.Fetch_UE_SystemOption(136) = "0" Then
            'multi-currency is not enabled - display the system currency
            MyCommon.QueryStr = "select C.CurrencyID, C.NamePhraseTerm, C.AbbreviationPhraseTerm, C.Symbol " & _
                                "from Currencies as C with (NoLock) where CurrencyID=" & CurrencyID & ";"
            rst2 = MyCommon.LRT_Select
            If rst2.Rows.Count > 0 Then
              Send(Copient.PhraseLib.Lookup("term.currency", LanguageID) & ":<br />")
              Send("&nbsp;&nbsp;" & Copient.PhraseLib.Lookup(rst2.Rows(0).Item("NamePhraseTerm"), LanguageID) & " (" & Copient.PhraseLib.Lookup(rst2.Rows(0).Item("AbbreviationPhraseTerm"), LanguageID) & ")")
              Send("        <br />")
              Send("        <br class=""half"" />")
            End If
          Else
            Send(Copient.PhraseLib.Lookup("term.currency", LanguageID) & ":<br />")

            'multi-currency is enabled - display the currency selector
            MyCommon.QueryStr = "select C.CurrencyID, C.NamePhraseTerm, C.AbbreviationPhraseTerm, C.Symbol " & _
                                "from Currencies as C with (NoLock) " & _
                                "where exists(select 1 from Locations with (NoLock) where LocationTypeID=1 and Deleted=0 and CurrencyID=C.CurrencyID);"
            rst2 =MyCommon.LRT_Select
            If rst2.Rows.Count > 0 Then
              Send("        <select name=""CurrencyID"" class=""long"">")
              If CurrencyID <= 0 Then
                Send("             <option value=""0"" selected=""selected"">" & Copient.PhraseLib.Lookup("term.unspecified", LanguageID) & "</option>")
              End If              
              For Each row2 In rst2.Rows
                Sendb("          <option value=""" & row2.Item("CurrencyID") & """" & IIf(CurrencyID = row2.Item("CurrencyID"), " selected=""selected""", "") & ">")
                Sendb(Copient.PhraseLib.Lookup(row2.Item("NamePhraseTerm"), LanguageID) & " (")
                Sendb(Copient.PhraseLib.Lookup(row2.Item("AbbreviationPhraseTerm"), LanguageID) & ")")
                Send("</option>")
              Next
              Send("        </select>")
            Else
              Sendb(Copient.PhraseLib.Lookup("term.nolocationcurrencies", LanguageID))
            End If
            rst2 = Nothing
            Send("        <br />")
            Send("        <br class=""half"" />")
          End If  'UE_SystemOption(136) (multi-currency enabled)
          
          'send the UOM selectors
          If MyCommon.Extract_Val(MyCommon.Fetch_UE_SystemOption(135)) = 0 Then
            'multi-UOM is not enabled.  If some units of measure were previous selected, display them
            MyCommon.QueryStr = "select UT.PhraseTerm as UOMTypePhrase, UST.NamePhraseTerm as UOMSubTypePhrase " & _
                                "from CPE_RewardOptionUOMs as ROU with (NoLock) Inner Join UOMTypes as UT on UT.UOMTypeID=ROU.UOMTypeID " & _
                                "Inner Join UOMSubTypes as UST on UST.UOMSubTypeID=ROU.UOMSubTypeID " & _
                                "where RewardOptionID=" & roid & ";"
            rst2 = MyCommon.LRT_Select
            If rst2.Rows.Count > 0 Then
              For Each row2 In rst2.Rows
                Send(Copient.PhraseLib.Lookup(MyCommon.NZ(row2.Item("UOMTypePhrase"), ""), LanguageID) & ":<br />")
                Send("&nbsp;&nbsp;" & Copient.PhraseLib.Lookup(MyCommon.NZ(row2.Item("UOMSubTypePhrase"), ""), LanguageID) & "<br />")
                Send("        <br class=""half"" />")
              Next
            End If
          Else
            MyCommon.QueryStr = "select distinct UOMT.UOMTypeID, UOMT.PhraseTerm " & _
                                "from UOMTypes as UOMT with (NoLock) " & _
                                "inner join UOMSubTypes as UOMST with (NoLock) " & _
                                "  on UOMST.UOMTypeID = UOMT.UOMTypeID;"
            rst2 = MyCommon.LRT_Select
            If rst2.Rows.Count > 0 Then
              For Each row2 In rst2.Rows
                ' lookup the selected subtype for this UOM Type
                If UOMSubTypes.ContainsKey(MyCommon.NZ(row2.Item("UOMTypeID"), 0)) Then
                  SubTypeID = UOMSubTypes.Item(MyCommon.NZ(row2.Item("UOMTypeID"), 0))
                Else
                  SubTypeID = 0
                End If
              
                Send(Copient.PhraseLib.Lookup(MyCommon.NZ(row2.Item("PhraseTerm"), ""), LanguageID) & ":<br />")
                Send("        <select name=""uomtype" & MyCommon.NZ(row2.Item("UOMTypeID"), 0) & """ class=""long"">")
                
                MyCommon.QueryStr = "select distinct UOMST.UOMTypeID, UOMST.UOMSubTypeID, UOMST.NamePhraseTerm, " & _
                                    "  UOMST.AbbreviationPhraseTerm " & _
                                    "from UOMSetItems as UOMSI with (NoLock) " & _
                                    "inner join Locations as LOC with (NoLock) on LOC.UOMSetID = UOMSI.UOMSetID " & _
                                    "inner join UOMTypes as UOMT with (NoLock) on UOMT.UOMTypeID = UOMSI.UOMTypeID " & _
                                    "inner join UOMSubTypes as UOMST with(NoLock) on UOMST.UOMSubTypeID = UOMSI.UOMSubTypeID " & _
                                    "where UOMT.UOMTypeID = @UOMTypeID and LOC.Deleted=0;"
                        MyCommon.DBParameters.Add("@UOMTypeID",SqlDbType.Int).Value = MyCommon.NZ(row2.Item("UOMTypeID"), 0)  
                rst3 = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                If SubTypeID <= 0 Then
                  Send("             <option value=""0"">" & Copient.PhraseLib.Lookup("term.unspecified", LanguageID) & "</option>")
                End If

                For Each row3 In rst3.Rows
                  Sendb("          <option value=""" & MyCommon.NZ(row3.Item("UOMSubTypeID"), 0) & """" & IIf(SubTypeID = MyCommon.NZ(row3.Item("UOMSubTypeID"), 0), " selected=""selected""", "") & ">")
                  Sendb(Copient.PhraseLib.Lookup(MyCommon.NZ(row3.Item("NamePhraseTerm"), ""), LanguageID) & "(")
                  Sendb(Copient.PhraseLib.Lookup(MyCommon.NZ(row3.Item("AbbreviationPhraseTerm"), ""), LanguageID) & ")")
                  Send("</option>")
                Next
                Send("        </select>")
                Send("        <br />")
                Send("        <br class=""half"" />")
              Next
            End If
          End If  'UE_SystemOption(135) (multi-uom enabled)
        %>
      </div>
      <div class="box" id="MEGselector">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.MutualExclusionGroups", LanguageID))%>
          </span>
        </h2>
        <% If (IsTemplate) Then%>
        <span class="temp">
        <input type="checkbox" class="tempcheck" id="Disallow_MutualExclusionGroups" name="Disallow_MutualExclusionGroups"
          <% if (Disallow_MutualExclusionGroups) then send(" checked=""checked""") %> />
        <label for="Disallow_MutualExclusionGroups">
          <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
        </span>
        <br class="printonly" />
        <% End If%>
        <%
          Send("<br class=""half"" />")
            Send("<input type=""button"" onclick=""javascript:openPagePopup('/logix/UE/UEoffer-meg.aspx?OfferID=" & OfferID & "');"" value=""" & Copient.PhraseLib.Lookup("ueoffer-gen.ManageMEGs", LanguageID) & """ name=""meg"" id=""meg""" & IIf(((FromTemplate And Disallow_MutualExclusionGroups) OrElse IsOfferWaitingForApproval(OfferID)), " disabled=""disabled""", "") & IIf((bEnableRestrictedAccessToUEOfferBuilder And isTranslatedOffer), " disabled=""disabled""", "") & IIf((bEnableAdditionalLockoutRestrictionsOnOffers And Not bOfferEditable), " disabled=""disabled""", "") & " />")
        %>
      </div>
	<%	    
	    If MyCommon.Fetch_SystemOption(156) = "1" Then
	        udflistcontrol.IsTemplate = IsTemplate
	        udflistcontrol.bUseTemplateLocks = FromTemplate
	        udflistcontrol.Disallow_UserDefinedFields = Disallow_UserDefinedFields
	        %>               
       <udf:udflist ID="udflistcontrol" runat="server"/>
 <% End If %>
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
            <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
          </label>
        </span>
        <br class="printonly" />
        <% End If%>
        <label for="limit3" style="position: relative;">
          <b>
            <% Sendb(Copient.PhraseLib.Lookup("term.reward", LanguageID))%>
            :</b></label>
        <br />
        <%        
  Send("<table summary=""" & Copient.PhraseLib.Lookup("term.rewards", LanguageID) & """>")
  Send("<tr id=""p3row1"">")
          Send("  <td>")
  Send("    <label for=""P3DistTimeType"">" & Copient.PhraseLib.Lookup("term.type", LanguageID) & ":</label>")
          Send("  </td>")
          Send("  <td>")
  Send("    <select id=""P3DistTimeType"" name=""P3DistTimeType"" class=""long""" & IIf((FromTemplate And Disallow_Limits), " disabled=""disabled""", "") & " onchange=""setLimits();"">")
  TempQueryStr = "select TimeTypeID,PhraseID from UE_DistributionTimeTypes with (NoLock)"
          If HasAnyCustomer Then
            TempQueryStr &= " where TimeTypeID=2"
          End If
          MyCommon.QueryStr = TempQueryStr
          rst2 = MyCommon.LRT_Select
          For Each row2 In rst2.Rows
    If (MyCommon.NZ(rst.Rows(0).Item("P3DistTimeType"), 0) = MyCommon.NZ(row2.Item("TimeTypeID"), 0) AndAlso Not NoLimit AndAlso Not OncePerOfferLimit andalso Not HourLimit) Then
              Send("      <option value=""" & MyCommon.NZ(row2.Item("TimeTypeID"), 0) & """ selected=""selected"">" & Copient.PhraseLib.Lookup(MyCommon.NZ(row2.Item("PhraseID"), 0), LanguageID) & "</option>")
            Else
              Send("      <option value=""" & MyCommon.NZ(row2.Item("TimeTypeID"), 0) & """>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row2.Item("PhraseID"), 0), LanguageID) & "</option>")
            End If
          Next
          Send("      <option value=""" & LIMIT_NONE & """" & IIf(NoLimit, "selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.nolimit", LanguageID) & "</option>")
  
          ' Do not allow "Once Per Offer" if any customer is selected or if there is no customer condition is specified since it requires that we generate reward distribution records to track the earning of the offer
          If (IsCustomerAssigned) AndAlso (Not HasAnyCustomer) Then
            Send("      <option value=""" & LIMIT_ONCE_PER_OFFER & """" & IIf(OncePerOfferLimit, "selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.onceperoffer", LanguageID) & "</option>")
            If MyCommon.Fetch_UE_SystemOption(205) = "1" Then
              Send("      <option value=""" & LIMIT_HOURS_LAST_AWARDED & """" & IIf(HourLimit, "selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.hourssincelastawarded", LanguageID) & "</option>")
            End If
          End If
          Send("    </select>")
          Send("    <input type=""hidden"" id=""BeginP3TimeTypeID"" name=""BeginP3TimeTypeID"" value=""" & MyCommon.NZ(rst.Rows(0).Item("P3DistTimeType"), -1) & """ />")
          Send("  </td>")
          Send("</tr>")
          If (IsUniqueProd) Then Sendb("<tr><td colspan=""2""><small style=""margin-left:100px;"">(" & Copient.PhraseLib.Lookup("term.disabledunique", LanguageID) & ")</small></td><td></td></tr>")
          Send("<tr id=""p3row2""" & IIf(NoLimit OrElse OncePerOfferLimit, " style=""display:none;""", "") & ">")
          Send("  <td>")
          Send("    <label for=""limit3"">" & Copient.PhraseLib.Lookup("term.limit", LanguageID) & ":</label>")
          Send("  </td>")
          Send("  <td>")
          Send("    <input type=""text"" class=""shorter"" id=""limit3"" name=""limit3"" maxlength=""6"" value=""" & MyCommon.NZ(rst.Rows(0).Item("P3DistQtyLimit"), 0) & """" & IIf((FromTemplate And Disallow_Limits), " disabled=""disabled""", "") & " />")
          Send("  </td>")
          Send("</tr>")
          Send("<tr id=""p3row3""" & IIf(MyCommon.NZ(rst.Rows(0).Item("P3DistTimeType"), 0) = 2 OrElse NoLimit OrElse OncePerOfferLimit, "style=""display:none;""", "") & ">")
          Send("  <td>")
          Send("    <label for=""limit3period"">" & Copient.PhraseLib.Lookup("term.period", LanguageID) & ":</label>")
          Send("  </td>")
          Send("  <td>")
          Sendb("    <input type=""text"" class=""shorter"" id=""limit3period"" name=""limit3period"" maxlength=""6"" value=""")
          Sendb(IIf(MyCommon.NZ(rst.Rows(0).Item("P3DistPeriod"), 0) > 0, MyCommon.NZ(rst.Rows(0).Item("P3DistPeriod"), 0), MyCommon.NZ(rst.Rows(0).Item("P3DistPeriod"), 0) * -1))
          Send("""" & IIf(((FromTemplate And Disallow_Limits) Or (HasAnyCustomer)), " disabled=""disabled""", "") & " />")
          Send("  </td>")
          Send("</tr>")
          Send("</table>")
%>
        <hr class="hidden" />
      </div>
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
            <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
          </label>
        </span>
        <br class="printonly" />
        <% End If%>
        <label for="InboundCRMEngineID" style="position: relative;">
          <% Sendb(Copient.PhraseLib.Lookup("term.creationsource", LanguageID))%>
          :</label>
        <br />
        <%
          If Logix.UserRoles.EditOfferSource Then
            MyCommon.QueryStr = "select ExtInterfaceID, PhraseID, Name from ExtCRMInterfaces with (NoLock) where Deleted=0 and Active=1 and ExtInterfaceTypeID in (0,1);"
            rst2 = MyCommon.LRT_Select
            If rst2.Rows.Count > 0 Then
              Send("<select id=""InboundCRMEngineID"" style=""width:295px;"" name=""InboundCRMEngineID""" & IIf(FromTemplate And Disallow_CRMEngine, " disabled=""disabled""", "") & ">")
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
          <% Sendb(Copient.PhraseLib.Lookup("offer-gen.sendoutbound", LanguageID))%>
          :</label>
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
          <% Sendb(Copient.PhraseLib.Lookup("term.chargebackvendor", LanguageID))%>
          :</label>
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
      <% If Not (HasAnyCustomer) Then%>
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
              <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
            </label>
          </span>
          <br class="printonly" />
          <% End If%>
          <input type="checkbox" id="EmployeeFiltering" name="EmployeeFiltering" onclick="handleEmployeeFiltering();"
            <% if (employeefiltered)then sendb(" checked=""checked""") %><% If ((FromTemplate And Disallow_EmployeeFiltering) Or Not Logix.UserRoles.EditEmployeeFiltering) Then Sendb(" disabled=""disabled""")%> />
          <label for="EmployeeFiltering">
            <% Sendb(Copient.PhraseLib.Lookup("offer-gen.empfilter", LanguageID))%>
          </label>
          <br />
          &nbsp;&nbsp;
          <input type="radio" id="employeesonly" name="employeesonly" <% if (employeesonly) then sendb(" checked=""checked""") %>
            onclick="toggleEmployee('employeesexcluded');" <% if((FromTemplate and disallow_employeefiltering) Or Not Logix.UserRoles.EditEmployeeFiltering )then sendb(" disabled=""disabled""") %> />
          <label for="employeesonly">
            <% Sendb(Copient.PhraseLib.Lookup("term.employeesonly", LanguageID))%>
          </label>
          <br />
          &nbsp;&nbsp;
          <input type="radio" id="employeesexcluded" name="employeesexcluded" style="padding-left: 5px;"
            <% if (employeesexcluded) then sendb(" checked=""checked""") %> onclick="toggleEmployee('employeesonly');"
            <% if((FromTemplate and disallow_employeefiltering) Or Not Logix.UserRoles.EditEmployeeFiltering)then sendb(" disabled=""disabled""") %> />
          <label for="employeesexcluded">
            <% Sendb(Copient.PhraseLib.Lookup("term.excludeemployees", LanguageID))%>
          </label>
          <br />
          <hr class="hidden" />
        </div>
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
    <% Send_Disc_Eval_Box(MyCommon, OfferID, DiscEvalTypeID, IsTemplate, FromTemplate, Disallow_RewardEvaluation)%>
    <div class ="box" id="offerType">
         <h2>
        <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.offertype", LanguageID))%>
        </span>
        </h2>
          <%  If (IsTemplate) Then
        Send("<span class=""temp"">")
                  Send("<input type=""checkbox"" class=""tempcheck"" id=""Disallow_OfferType"" name=""Disallow_OfferType""")
        If (Disallow_OfferType) Then
          Send(" checked=""checked""")
        End If
        Send(" />")
        Send("<label for=""Disallow_OfferType"">")
        Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))
        Send("</label>")
        Send("</span>")
      End If %>

         <label for="OfferType"> <% Sendb(Copient.PhraseLib.Lookup("term.applyoffertotransaction", LanguageID))%>:</label> &nbsp;    
         <% MyCommon.QueryStr = "select CouponTypeID,Description,PhraseID from CouponTypes with (NoLock);"
             If IsMfgCoupon = 1 Then
                 SelectedOfferType = 1
             End If
             If StoreCoupon = 1 Then
                 SelectedOfferType = 2
             End If
             rst2 = MyCommon.LRT_Select
             If rst2.Rows.Count > 0 Then
                 Send("<select id=""offerType"" name=""offerType""" & IIF((FromTemplate And Disallow_OfferType)," disabled=""disabled""","") & ">")
                 For Each row In rst2.Rows
                     If MyCommon.NZ(row.Item("CouponTypeID"), 0) = SelectedOfferType Then
                         Sendb("  <option value=""" & MyCommon.NZ(row.Item("CouponTypeID"), 0) & """ selected=""selected"">")
                         Sendb(Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID, MyCommon.NZ(row.Item("Description"), "")))
                         Send("</option>")
                     Else
                         Sendb("  <option value=""" & MyCommon.NZ(row.Item("CouponTypeID"), 0) & """>")
                         Sendb(Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID, MyCommon.NZ(row.Item("Description"), "")))
                         Send("</option>")
                     End If
                 Next
                 Send("</select>")
             End If%>
    </div>
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
          <% Sendb(Copient.PhraseLib.Lookup("term.autotransferable", LanguageID))%>
        </label>
        <br />
      <input type="checkbox" name="deferToEOS" id="deferToEOS" <% if (deferToEOS) then sendb(" checked=""checked""") %><%if (DeferToEOSDisabled or (FromTemplate and Disallow_AdvancedOption)) then sendb(" disabled=""disabled""") %> />
      <label for="deferToEOS">
        <% Sendb(Copient.PhraseLib.Lookup("term.defercalc", LanguageID))%>
      </label>
      <br />
      <% If MyCommon.Fetch_UE_SystemOption(211) = "1" Then %>
        <input type="checkbox" name="deferCalcToTotal" id="deferCalcToTotal" <% if (DeferCalcToTotal) then sendb(" checked=""checked""") %><%if (DeferCalcToTotalDisabled or deferToEOS or(FromTemplate and Disallow_AdvancedOption)) then sendb(" disabled=""disabled""") %> />
        <label for="deferCalcToTotal">
        <% Sendb(Copient.PhraseLib.Lookup("term.deferCalcToTotal", LanguageID))%>
        </label>
        <br />
        &nbsp;
        <label for="PointsProgramWatch"><% Sendb(Copient.PhraseLib.Lookup("term.pointsprogramwatch", LanguageID))%></label>
        <% 
        MyCommon.QueryStr = "select ProgramID, ProgramName from pointsprograms where Deleted = 0"
        dt = MyCommon.LRT_Select
		  'PointsProgramWatch = 1
        Send("<select id =""PointsProgramWatch"" name = ""PointsProgramWatch""  >")
        If(PointsProgramWatchDisabled or DeferCalcToTotalDisabled or deferToEOS or(FromTemplate and Disallow_AdvancedOption)) then 
         Send("<option value=0 />")
        Else
          Send("<option value=0 />")
          For each row in dt.Rows
            If PointsProgramWatch = row.Item("ProgramID") Then
              Send("<option value=" & row.Item("ProgramID") & " selected= true>" &  row.item("ProgramName")  & " </option>")
            Else
              Send("<option value=" & row.Item("ProgramID") & ">" &  row.item("ProgramName")  & " </option>")
            End If
          Next
        End If
        Send("</select>")
        %>
        <br />
      <% End If %>
      <input type="checkbox" name="enableCollisionDetection" id="enableCollisionDetection" <%Sendb(IIf(EnableCollisionDetection, " checked=""checked""", ""))%>
        <%Sendb(IIF((FromTemplate And Disallow_AdvancedOption)," disabled=""disabled""","")) %> />
      <label for="enableCollisionDetection">
        <% Sendb(Copient.PhraseLib.Lookup("term.enableCD", LanguageID))%>
      </label>
      <br />
      <% If bUseProrateFlag Then%>
      <input type="checkbox" name="prorateondisplay" id="prorateondisplay" <%Sendb(IIf(IsProrateonDisplay, " checked=""checked""", ""))%>
        <%Sendb(IIF((FromTemplate And Disallow_AdvancedOption)," disabled=""disabled""","")) %> />
      <label for="prorateondisplay">
        <% Sendb(Copient.PhraseLib.Lookup("offer-gen.EnableProrateonDisplay", LanguageID))%>
      </label>
      <br />
      <%End If%>
      <% If bUseDisplayFlag Then%>
      <input type="checkbox" name="promotiondisplay" id="promotiondisplay" <% if (IsPromotionDisplay) then sendb(" checked=""checked""") %>
        <%Sendb(IIF((FromTemplate And Disallow_AdvancedOption)," disabled=""disabled""",""))%> />
      <label for="promotiondisplay">
        <% Sendb(Copient.PhraseLib.Lookup("offer-gen.EnableProrationDisplay", LanguageID))%>
      </label>
      <br />
      <%End If%>
	   
      <% If (MyCommon.Fetch_UE_SystemOption(26).Trim = "1") Then%>
      <% If Not (HasAnyCustomer) Then%>
      <input type="checkbox" name="reportingimp" id="reportingimp" <% if (reportingimp) then sendb(" checked=""checked""") %>
        <%Sendb(IIF((FromTemplate And Disallow_AdvancedOption)," disabled=""disabled""","")) %> />
      <label for="reportingimp">
        <% Sendb(Copient.PhraseLib.Lookup("offer-gen.enablereporting-imp", LanguageID))%>
      </label>
      <br />
      <%End If%>
      <input type="checkbox" name="reportingred" id="reportingred" <% if (reportingred) then sendb(" checked=""checked""") %>
        <%Sendb(IIF((FromTemplate And Disallow_AdvancedOption)," disabled=""disabled""","")) %> />
      <label for="reportingred">
        <% Sendb(Copient.PhraseLib.Lookup("offer-gen.enablereporting-red", LanguageID))%>
      </label>
      <br />
      <% End If%>
      <% If (MyCommon.Fetch_SystemOption(73).Trim <> "") Then%>
      <input type="checkbox" name="exporttoedw" id="exporttoedw" <% if (ExportToEDW) then sendb(" checked=""checked""") %>
        <%Sendb(IIF((FromTemplate And Disallow_AdvancedOption)," disabled=""disabled""","")) %> />
      <label for="exporttoedw">
          <% Sendb(Copient.PhraseLib.Lookup("term.exporttoarchive", LanguageID))%>
        </label>
        <br />
        <% End If%>
        <%
          If (MyCommon.Extract_Val(MyCommon.Fetch_UE_SystemOption(70)) = 1) Then
            'MyCommon.QueryStr = "Select Description, PhraseID from CPE_DeliverableTypes where IssuanceEnabled=1;"
                MyCommon.QueryStr = "select RDO.PKID, RDS.Description, RDO.Enabled from RemoteDataOptions as RDO " & _
                                    "inner join RemoteDataStyles as RDS on RDO.StyleID=RDS.StyleID and RDO.RemoteDataTypeID=RDS.RemoteDataTypeID " & _
                                    "where RDO.Enabled=1 and "
                If EngineID = 9 Then
                    MyCommon.QueryStr &= "RDO.RemoteDataTypeID in (1,3);"
                Else
                    MyCommon.QueryStr &= "RDO.RemoteDataTypeID=1;"
                End If
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
        <% Sendb(Copient.PhraseLib.Lookup("offer-gen.enableissuance", LanguageID))%>
      </label>
      <a href="#" onclick="javascript:showGrowPopup(event, '<%Sendb(IssuanceDetails)%>', 330, 200);">
        <img src="/images/info.png" alt="<% Sendb(Copient.PhraseLib.Lookup("term.details", LanguageID)) %>"
          title="<% Sendb(Copient.PhraseLib.Lookup("term.details", LanguageID)) %>" style="position: relative;
          top: 2px;" /></a>
      <div id="divIssuance" style="display: none;">
        Test
      </div>
      <br />
      <% End If%>
	  <% If MyCommon.Fetch_UE_SystemOption(175) = "1" Then
		 ' load the list items
		 
                 MyCommon.QueryStr = "select PromoCategoryID from PromoGridOffers where Deleted = 0 and IncentiveID="& OfferID
		 dt=MyCommon.LRT_Select
		 For Each rowS As DataRow In dt.Rows
			SavedSelectID=MyCommon.NZ(rowS.Item("PromoCategoryID"), 0)
                 Next	
                 MyCommon.QueryStr = "select PromoCategoryID, PhraseID, Name, Visible from PromoOfferCategories" 
                 dt = MyCommon.LRT_Select
                 Send("     <label for=""PromoGridCategory"" >"& Copient.PhraseLib.Lookup("promogrid.label", LanguageID) &"</label>   ")
                 Send("	<select id=""PromoGridCategory"" name=""PromoGridCategory"" " & IIF((FromTemplate And Disallow_AdvancedOption)," disabled=""disabled""","") & "> ")
                 Send("<option value=""-1"" " & IIf(SavedSelectID = -1, "selected", "") & "/>")
			
                 For Each rowH As DataRow In dt.Rows
				
                   SelectName=Copient.PhraseLib.Lookup(MyCommon.NZ(rowH.Item("PhraseID"), MyCommon.NZ(rowH.Item("Name"), "Unknown")), LanguageID,)
                   SelectID=MyCommon.NZ(rowH.Item("PromoCategoryID"), 0)
                   If(MyCommon.NZ(rowH.Item("Visible"),0))Then
                     If (SelectID = SavedSelectID) Then
                       Send("<option value="& SelectID & " selected>"& SelectName &"</option>" )
                     Else 
                       Send("<option value="""& SelectID & """>"& SelectName &"</option>" ) 
                     End If
                   End If	
                Next
                Send("</select>")
        End If %>
        <% If MyCommon.Fetch_UE_SystemOption(207) = "1" Then%>
        <!-- BZ2079: UE-feature-removal #23: Hide mfg coupon box.  To restore, remove this <div> (and its corresponding </div>) -->
        <input type="checkbox" name="mfgCoupon" id="mfgCoupon" <% if (IsMfgCoupon) then sendb(" checked=""checked""") %>
          <%Sendb(IIF((FromTemplate And Disallow_AdvancedOption)," disabled=""disabled""","")) %> />
          <label for="mfgCoupon">
            <% Sendb(Copient.PhraseLib.Lookup("CPEoffer-gen.mfgcoupons", LanguageID))%>
          </label>
          <br />
        <% End If %>
        <%
          If (Logix.UserRoles.FavoriteOffersForOthers AndAlso Not IsTemplate) Then
            Send("<br class=""half"" />")
            MyCommon.QueryStr = "select AdminUserID from AdminUserOffers where OfferID=" & OfferID & ";"
            rst2 = MyCommon.LRT_Select
            MyCommon.QueryStr = "select AdminUserID from AdminUsers;"
            rst3 = MyCommon.LRT_Select
          Send("<button type=""button"" id=""favorite"" name=""favorite"" value=""favorite""" & IIf((FromTemplate And isTranslatedOffer), " disabled=""disabled""", "") &IIf((bEnableRestrictedAccessToUEOfferBuilder And Disallow_AdvancedOption), " disabled=""disabled""", "") &IIf((bEnableAdditionalLockoutRestrictionsOnOffers And Not bOfferEditable), " disabled=""disabled""", "") & "onclick=""javascript:xmlhttpPost('/logix/OfferFeeds.aspx', 'FavoriteForAll');"">" & Copient.PhraseLib.Lookup("offer-gen.favoriteall", LanguageID) & "</button>")
          Sendb("<a href=""javascript:openPopup('/logix/offer-favorite.aspx?OfferID=" & OfferID & "&bUseTemplateLocks=" & FromTemplate & "&Disallow_AdvancedOption=" & Disallow_AdvancedOption & "')""><img id=""favImg"" src=""/images/user.png"" ")
            Sendb("alt=""" & rst2.Rows.Count & " " & StrConv(Copient.PhraseLib.Lookup("term.of", LanguageID), VbStrConv.Lowercase) & " " & rst3.Rows.Count & " " & StrConv(Copient.PhraseLib.Lookup("term.users", LanguageID), VbStrConv.Lowercase) & " " & Copient.PhraseLib.Lookup("offer-gen.havefavorite", LanguageID) & """ ")
            Sendb("title=""" & rst2.Rows.Count & " " & StrConv(Copient.PhraseLib.Lookup("term.of", LanguageID), VbStrConv.Lowercase) & " " & rst3.Rows.Count & " " & StrConv(Copient.PhraseLib.Lookup("term.users", LanguageID), VbStrConv.Lowercase) & " " & Copient.PhraseLib.Lookup("offer-gen.havefavorite", LanguageID) & """ ")
            Send("/></a><br />")
          End If
        %>
		
      </div>
      <% If MyCommon.Fetch_UE_SystemOption(180) = "1" Then%>
	  	<div class="box" id="preorder">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.preordereligibility", LanguageID))%>
        </span>
      </h2>
	   <%If (IsTemplate) Then%>
      <span class="temp">
        <input type="checkbox" class="tempcheck" id="Disallow_PreOrder" name="Disallow_PreOrder"
          <% if(Disallow_PreOrder)then send(" checked=""checked""") %> />
        <label for="Disallow_PreOrder">
          <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
        </label>
      </span>
      <br class="printonly" />
      <% End If%>
      <input type="checkbox" name="preorderEligibility" id="preorderEligibility" <%Sendb(IIf(PreOrderEligibility, " checked=""checked""", ""))%><%Sendb(IIF((FromTemplate And Disallow_PreOrder)," disabled=""disabled""","")) %>/>
      <label for="preorderEligibility">
		<% Sendb(Copient.PhraseLib.Lookup("term.preordercheckbox", LanguageID))%>
      </label>
	  </div>
	  <%End If %>
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
                              "where BE.EngineID=9 and AUB.AdminUserID =" & AdminUserID & ";"
          rst2 = MyCommon.LRT_Select
          EditableBanners = New ArrayList(rst2.Rows.Count)
          For Each row2 In rst2.Rows
            EditableBanners.Add(MyCommon.NZ(row2.Item("BannerID"), -1))
          Next
          
          ' get all the assigned banners for CPE
          i = 0
          MyCommon.QueryStr = "select BAN.BannerID, BAN.Name from Banners BAN with (NoLock) " & _
                              "inner join BannerEngines BE with (NoLock) on BE.BannerID = BAN.BannerID " & _
                              "where BE.EngineID=9 and BAN.AllBanners=0;"
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
                              "where BE.EngineID=9 and BAN.AllBanners=1;"
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

$(document).ready(function() {
        var savedTimeVal = document.getElementById('savedTime');
        var offerIDVal = document.getElementById('OfferID');
        if(savedTimeVal != null && offerIDVal != null) {
            var savedTime = new Date(savedTimeVal.value).getTime();
            var presentTime = new Date().getTime();
            var seconds = (presentTime - savedTime) / 1000;
            if(seconds > 2){
                $.support.cors = true;
                $.ajax({
                    type: "POST",
                    url: "/Connectors/AjaxProcessingFunctions.asmx/GetLockedSystemOptions",
                    data: JSON.stringify({ offerID : offerIDVal.value }),
                    contentType: "application/json; charset=utf-8",
                    dataType: "json"
                })
                .done(function (data) {
                    if(data.d == "true"){
                        window.location.href = window.location.href.replace("UEOffer-gen.aspx", "UEOffer-sum.aspx");
                    }
                });
            }
        }
    });
</script>
<script runat="server">
  
  Function ExistGCRPercentOff(ByRef MyCommon As Copient.CommonInc, ByVal ROID As Long) As Boolean
    Dim percentOffGCRLinked As Boolean = False
    Dim dt As DataTable
    MyCommon.QueryStr = "Select distinct gc.Id From CPE_Deliverables cd join GiftCard gc on (cd.OutputId=gc.Id) join GiftCardTier gct on (gc.id=gct.GiftCardId) " & _
                        " Where gct.AmountTypeId=@AmountTypeId And cd.RewardOptionid=@ROID And DeliverableTypeId=@DeliverableTypeId And cd.Deleted=0;"
    MyCommon.DBParameters.Add("@ROID", SqlDbType.BigInt).Value = ROID
    MyCommon.DBParameters.Add("@DeliverableTypeId", SqlDbType.BigInt).Value = DELIVERABLE_TYPES.GIFTCARD
    MyCommon.DBParameters.Add("@AmountTypeId", SqlDbType.BigInt).Value = CPEAmountTypes.PercentageOff
    dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
    If Not dt Is Nothing AndAlso dt.Rows.Count = 1 AndAlso CInt(dt.Rows(0)(0)) > 0 Then
      percentOffGCRLinked = True
    End If
    Return percentOffGCRLinked
  End Function

  Function ExistProductPriceCondition(ByRef MyCommon As Copient.CommonInc, ByVal ROID As Long) As Boolean
    Dim priceConditionLinked As Boolean = False
    Dim dt As DataTable
    MyCommon.QueryStr = "select count(IncentiveProductGroupID) from CPE_IncentiveProductGroups Where QtyUnitType=@UnitTypeID And RewardOptionId=@ROID And Deleted=0;"
    MyCommon.DBParameters.Add("@ROID", SqlDbType.BigInt).Value = ROID
    MyCommon.DBParameters.Add("@UnitTypeID", SqlDbType.Int).Value = CPEUnitTypes.Dollars
    dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
    If Not dt Is Nothing AndAlso dt.Rows.Count = 1 AndAlso CInt(dt.Rows(0)(0)) > 0 Then
      priceConditionLinked = True
    End If
    Return priceConditionLinked
    End Function

  '----------------------------------------------------------------------------------------------------------------------------------------------
  
  Function IsDeployableOffer(ByRef MyCommon As Copient.CommonInc, ByVal OfferID As Integer, ByVal ROID As Integer, ByRef ErrorMsg As String) As Boolean
    Dim bDeployable As Boolean = False
    
    ErrorMsg = ""
    bDeployable = MeetsDeploymentReqs(MyCommon, OfferID)
    
    If bDeployable Then
      bDeployable = MeetsTemplateRequirements(MyCommon, ROID)
      If (Not bDeployable) Then
        ErrorMsg = Copient.PhraseLib.Lookup("offer-sum.required-incomplete", LanguageID)
      End If
    Else
      ErrorMsg = Copient.PhraseLib.Lookup("cpeoffer-sum.deployalert", LanguageID)
    End If
    
    If bDeployable Then
      bDeployable = MeetsLocalizationReqs(MyCommon, ROID, ErrorMsg)
    End If
    
    Return bDeployable
  End Function
  
  '----------------------------------------------------------------------------------------------------------------------------------------------
  
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
  
  '----------------------------------------------------------------------------------------------------------------------------------------------
  
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
  
  '----------------------------------------------------------------------------------------------------------------------------------------------
  
  Function MeetsLocalizationReqs(ByRef Common As Copient.CommonInc, ByVal ROID As Integer, ByRef ErrorMsg As String) As Boolean
    
    Dim ReturnVal As Boolean = True
    Dim dst As DataTable
    Dim row As DataRow
    Dim StoreList As String = ""
    
    ErrorMsg = ""
    'query to see if there are any locations joined to the offer that use a currencyID that is different than RewardOptions.CurrencyID 
    Common.QueryStr = "select 1" & _
                      "from Locations as L with (NoLock) Inner Join LocGroupItems as LGI with (NoLock) on L.LocationID=LGI.LocationID and LGI.Deleted=0 and L.Deleted=0 " & _
                      "Inner Join OfferLocations as OL with (NoLock) on OL.LocationGroupID=LGI.LocationGroupID and OL.Deleted=0 " & _
                      "Inner Join CPE_RewardOptions as RO with (NoLock) on RO.IncentiveID=OL.OfferID and RO.Deleted=0 and RO.TouchResponse=0 " & _
                      "Where RO.RewardOptionID=" & ROID & " and L.CurrencyID<>RO.CurrencyID;"
    dst = Common.LRT_Select
    If dst.Rows.Count > 0 Then
      ReturnVal = False
      ErrorMsg = Copient.PhraseLib.Lookup("offer-sum.unsupportedcurrency", LanguageID) 'Not all targeted locations support the currency selected for this offer
    End If
    
    If ReturnVal = True Then  'no need to run this query if there has already been a violation of the requirements
      'query to see if there are any locations that do not support the units of measure that are used by this offer
      Common.QueryStr = "dbo.pa_invalidlocationuomsforoffer"
      Common.Open_LRTsp()
      Common.LRTsp.Parameters.Add("@RewardOptionID", SqlDbType.BigInt).Value = ROID
      dst = Common.LRTsp_select
      If dst IsNot Nothing AndAlso dst.Rows.Count > 0 Then
        ReturnVal = False
        ErrorMsg = Copient.PhraseLib.Lookup("offer-sum.unsupporteduom", LanguageID) 'The following targeted locations do not support the units of measure selected for this offer: 
        For Each row In dst.Rows
          If Not (StoreList = "") Then StoreList = StoreList & ", "
          StoreList = StoreList & row.Item("LocationName")
        Next
        If Len(StoreList) > 500 Then StoreList = Left(StoreList, 500) & " ..."
        ErrorMsg = ErrorMsg & StoreList
      End If
    End If

    Return ReturnVal
    
  End Function
  
  '----------------------------------------------------------------------------------------------------------------------------------------------
  
  Sub Send_Disc_Eval_Box(ByRef MyCommon As Copient.CommonInc, ByVal OfferID As Integer, ByVal SelectedValue As Integer, ByVal IsTemplate As Boolean, ByVal FromTemplate As Boolean, ByVal Disallow_RewardEvaluation As Boolean)
    Dim dt As DataTable
    Dim Name As String = ""
    Dim EvalID As Integer = 0
    Dim CheckedStr As String = ""
    Dim DisabledStr As String = ""
    ' load the list items
    MyCommon.QueryStr = "select DiscountEvalTypeID, Name, PhraseID from UE_DiscountEvalTypes with (NoLock);"
    dt = MyCommon.LRT_Select
    
    If dt.Rows.Count > 0 Then
      Send("<div class=""box"" id=""discounteval"">")
      Send("  <h2>")
      Send("    <span>")
      Send("      " & Copient.PhraseLib.Lookup("term.discountevaluation", LanguageID))
      Send("    </span>")
      Send("    </h2>")
      
      If (IsTemplate) Then
        Send("<span class=""temp"">")
        Send("<input type=""checkbox"" class=""tempcheck"" id=""Disallow_RewardEvaluation"" name=""Disallow_RewardEvaluation""")
        If (Disallow_RewardEvaluation) Then
          Send(" checked=""checked""")
        End If
        Send(" />")
        Send("<label for=""Disallow_RewardEvaluation"">")
        Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))
        Send("</label>")
        Send("</span>")
      End If
      If (FromTemplate And Disallow_RewardEvaluation) Then
        DisabledStr = " disabled=""disabled"" "
      End If
      For Each row As DataRow In dt.Rows
        EvalID = MyCommon.NZ(row.Item("DiscountEvalTypeID"), 0)
        Name = Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID, MyCommon.NZ(row.Item("Name"), "Unknown"))

        CheckedStr = IIf(SelectedValue = EvalID, " checked=""checked""", "")

        Send("    <input type=""radio"" id=""discEvalType" & EvalID & """ name=""discEvalType"" value=""" & EvalID & """" & CheckedStr & DisabledStr & " />")
        Send("    <label for=""discEvalType" & EvalID & """>" & Name & "</label>")
      Next

      Send("    <br />")
      Send("    <hr class=""hidden"" />")
      Send("</div>")
    End If
      
  End Sub
  
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
  If MyCommon.Fetch_SystemOption(156) = "1" Then
    Send("<div id=""imagepopup"" style=""display:none;"">")
    Send("  <div style=""float:right;"">")
    Send("    <a href=""#"" onclick=""javascript:closeImage();"">" & Copient.PhraseLib.Lookup("term.close", LanguageID, "Close") & "</a>")
    Send("  </div>")
    Send("  <div id=""imagebody"">")
    Send("    <table id=""centertable""><tr><td><img id=""fullSizedImage"" src="""" onerror=""this.src='/images/notfound.png'"" /></td></tr></table>")
    Send("  </div>")
    Send("</div>")
  End If
%>
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