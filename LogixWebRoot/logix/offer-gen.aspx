<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Register Src="~/logix/UserControls/UDFListControl.ascx" TagPrefix="udf" TagName="udflist" %>
<%@ Register Src="~/logix/UserControls/UDFJavaScript.ascx" TagPrefix="udf" TagName="udfjavascript" %>
<%@ Register Src="~/logix/UserControls/UDFSaveControl.ascx" TagPrefix="udf" TagName="udfsave" %>
<%@ Import Namespace="System.Data" %>
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
          } else {
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
</script>
<%
  ' *****************************************************************************
  ' * FILENAME: offer-gen.aspx 
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
  Dim OfferID As Long
  Dim Name As String = ""
  Dim NumberofFolders As Integer = 0
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim rst As DataTable
    
  Dim udfSaverst As DataTable
  Dim udfrst As DataTable
    
  Dim row As DataRow
  Dim rst2 As DataTable
  Dim crmdt As DataTable
  Dim row2 As DataRow
  Dim rst3 As DataTable
  Dim InstallPath As String
  Dim TierTypeID As Integer
  Dim EmployeeFiltering As Boolean
  Dim NonEmployeesOnly As Boolean
  Dim OfferCategoryID As Integer
  Dim NumTiers As Integer
  Dim PriorityLevel As Integer
  Dim EngineID As Integer = 0
  Dim EngineSubTypeID As Integer = 0
  Dim CRMEngineID As Integer
  Dim CMOADeployStatus As Integer
  Dim ExtOfferID As String = String.Empty
  Dim ChangedVal As String = String.Empty
  Dim ArrChangedVal() As String
  Dim ExtOfferID2 As String
  Dim Description As String
  Dim DistPeriodLimit As Integer
  Dim DistPeriod As Integer
  Dim InstantWin As Integer
  Dim NumPrizesAllowed As Integer
  Dim OddsOfWinning As Integer
  Dim RandomWinners As Boolean
  Dim IWTransLevel As Boolean
  Dim SharedLimitID As Integer
  Dim IsTemplate As Boolean
  Dim VarID As Long
  Dim DisabledOnCFW As Boolean
  Dim DisplayOnWebKiosk As Boolean
  Dim ExportToEDW As Boolean
  Dim Favorite As Boolean
  Dim AutoTransferable As Boolean = False
  Dim HighValue As Boolean = False
  Dim bFobEligible As Boolean
  Dim bAutoTranslate As Boolean = False
  Dim DisableDisconnectedOffer As Boolean = False
  Dim form_OfferID As String
  Dim form_Name As String
  Dim form_Description As String
  Dim form_Category As Integer
  Dim form_EmployeeFiltering As String
  Dim form_NonEmployeesOnly As String
  Dim form_TestStartdate As String
  Dim form_TestEnddate As String
  Dim form_ProdStartdate As String
  Dim form_ProdEnddate As String
  Dim form_DispStartdate As String
  Dim form_DispEnddate As String
  Dim form_DistPeriod As Integer
  Dim form_DistPeriodLimit As Integer
  Dim form_TierTypeID As Integer
  Dim form_NumTiers As Integer
  Dim form_EngineID As Integer
  Dim form_CRMEngineID As Integer
  Dim form_Priority As Integer
  Dim form_InstantWin As String
  Dim form_NumPrizesAllowed As Integer
  Dim form_OddsOfWinning As Integer
  Dim form_ExtOfferID2 As String
  Dim form_RandomWinners As String
  Dim form_IWTransLevel As String
  Dim form_DisabledOnCFW As String
  Dim form_DisplayOnWebKiosk As String
  Dim form_ExportToEDW As String
  Dim form_Favorite As String
  Dim form_AutoTransferable As String
  Dim form_HighValue As String
  Dim form_Fob As String
  Dim form_AutoTranslate As String
  Dim form_DisableDisconnectedOffer As String
  Dim form_ProdStartHr As String
  Dim form_ProdStartMin As String
  Dim form_ProdEndHr As String
  Dim form_ProdEndMin As String
  Dim form_TestStartHr As String
  Dim form_TestStartMin As String
  Dim form_TestEndHr As String
  Dim form_TestEndMin As String
  Dim form_DispStartHr As String
  Dim form_DispStartMin As String
  Dim form_DispEndHr As String
  Dim form_DispEndMin As String

  Dim FolderStartDate As String = ""
  Dim FolderEndDate As String = ""
  Dim CreatedDate As String
  Dim LastUpdate As String
  Dim StatusFlag As Integer
  Dim iterate As Integer
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
  Dim Disallow_DisplayDates As Boolean = True
  Dim Disallow_UserDefinedFields As Boolean = True
  Dim Disallow_AdvancedOption As Boolean = True
  Dim ShowInboundOutboundBox As Boolean = True
  Dim FormValid As Boolean = True
  Dim DuplicateName As Boolean = False
  Dim infoMessage As String = ""
  Dim modMessage As String = ""
  Dim Handheld As Boolean = False
  Dim BannerName As String = ""
  Dim rstBanners As DataTable = Nothing
  Dim i As Integer = 0
  Dim SelectedBanners, EditableBanners As ArrayList
  Dim IsEditableBanner As Boolean = False
  Dim BannersEnabled As Boolean = False
  Dim StatusCode As Copient.LogixInc.STATUS_FLAGS = Copient.LogixInc.STATUS_FLAGS.STATUS_UNKNOWN
  Dim StatusText As String = ""
  Dim bUseTestDates As Boolean
  Dim DescriptLength As Boolean = False
  Dim Descript As String = ""
  Dim bUseTemplateLocks As Boolean
  Dim AdvancedLimitID As Long
  Dim ValidDayLimit As Boolean = True
  Dim InboundCRMEngineID As Integer
  Dim ChangeExtID As Boolean = False
  Dim rst4 As DataTable

  Dim bAllowTimeWithStartEndDates As Boolean
  Dim bUseDisplayDates As Boolean = False
  Dim ProdStartDate As String
  Dim ProdStartHr As String
  Dim ProdStartMin As String
  Dim ProdEnddate As String
  Dim ProdEndHr As String
  Dim ProdEndMin As String
  Dim TestStartdate As String
  Dim TestStartHr As String
  Dim TestStartMin As String
  Dim TestEndDate As String
  Dim TestEndHr As String
  Dim TestEndMin As String
  Dim DispStartDate As String
  Dim DispStartHr As String
  Dim DispStartMin As String
  Dim DispEnddate As String
  Dim DispEndHr As String
  Dim DispEndMin As String
  Dim sDateOnlyFormat As String = "MM/dd/yyyy"
  Dim sHourOnlyFormat As String = "HH"
  Dim sMinutesOnlyFormat As String = "mm"
  Dim tempDateTime As Date
  Dim HasPrefCondition As Boolean = False
  Dim IsExtIdInboundIdExist As Boolean = False
  Dim Status As Integer
  Dim UseLegacyflag As Boolean
  Dim AllowSpecialCharacters As String
  Dim ComposedHist As String
  Dim UDFHistory As String
  Const gintEcouponOfferTypeID As Integer = 5 'E-Coupon will be OfferTypeID 5
  Dim defaultPriority As Integer
  Dim defaultECPriority As Integer
  Dim CashPromptRew_Offer As Boolean
  Dim permission_disallow_priority As Boolean
  
  Dim OfferRedemptionThresholdperHour As Integer = 0
  Dim bUseOfferRedemptionThreshold As Boolean = IIf(MyCommon.Fetch_CM_SystemOption(83) = "1", True, False)
   
   
  Dim Disallow_OfferRedempThreshold As Boolean = True
  Dim bCopyInboundCrmEngineID As Boolean = True
  
  Dim rstTemp As DataTable
  Dim DefaultAsLogixID As Boolean
  Dim bFobEligibilityEnabled As Boolean
  
  Dim bDisplayBuckTierType As Boolean = False
  Dim bDisableEditTierType As Boolean = False
  Dim bDisableEditName As Boolean = False
  Dim lBuckPromoVarId As Long = 0
  Dim MyImport As Copient.ImportXml
    
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "offer-gen.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  ' fill in if it was a get method
  OfferID = Request.QueryString("OfferID")
  form_OfferID = Request.QueryString("form_OfferID")
  InstallPath = MyCommon.Get_Install_Path(Request.PhysicalPath)
  
  AllowSpecialCharacters = MyCommon.Fetch_SystemOption(171)
  Dim bUseAdvertisementText As Boolean = IIf(MyCommon.Fetch_CM_SystemOption(99) = "1", True, False)
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  bUseTestDates = (MyCommon.Fetch_SystemOption(88) = "1")
  bAllowTimeWithStartEndDates = (MyCommon.Fetch_CM_SystemOption(26) = "1")
  PriorityLevel = 50
  EmployeeFiltering = False
  RandomWinners = False
  IWTransLevel = False
  NonEmployeesOnly = False
  
  bFobEligibilityEnabled = IIf(MyCommon.Fetch_CM_SystemOption(142) = "1", True, False)
  
  If MyCommon.Fetch_CM_SystemOption(85) = "1" Then
    bUseDisplayDates = True
  Else
    bUseDisplayDates = False
  End If

  If MyCommon.Fetch_CM_SystemOption(107) = "1" Then
    bCopyInboundCrmEngineID = True
  Else
    bCopyInboundCrmEngineID = False
  End If
    
  If (Request.QueryString("NewTemplate") <> "") Then
    IsTemplate = True
  End If
  
  ' handle form submissions here
  If (Request.QueryString("infoMessage") <> "") Then
    infoMessage = Request.QueryString("infomessage")
  End If
  
  'Set the favorite boolean
  If OfferID > 0 Then
    MyCommon.QueryStr = "Select Favorite from Offers with (NoLock) where OfferID=" & OfferID
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
      Favorite = MyCommon.NZ(rst.Rows(0).Item("Favorite"), False)
    End If
  End If
  
  Dim bEnableBuckOffers As Boolean = (MyCommon.Fetch_CM_SystemOption(137) = "1")
  'CR 102
  If bEnableBuckOffers Then
    MyImport = New Copient.ImportXml(MyCommon)
    If OfferID > 0 Then
      Dim oBuckStatus As Copient.ImportXml.BuckOfferStatus

      oBuckStatus = MyImport.BuckOfferGetStatus(OfferID)
      Select Case oBuckStatus
        Case Copient.ImportXml.BuckOfferStatus.BuckParentNoChildren,
          Copient.ImportXml.BuckOfferStatus.BuckParentBothChildren,
          Copient.ImportXml.BuckOfferStatus.BuckParentPaperChildrenOnly,
          Copient.ImportXml.BuckOfferStatus.BuckParentDigitalChildrenOnly
          bDisableEditName = True
          bDisableEditTierType = True
          bDisplayBuckTierType = True
        Case Copient.ImportXml.BuckOfferStatus.BuckChildPaper
          bDisableEditName = True
          bDisableEditTierType = True
          bDisplayBuckTierType = False
        Case Copient.ImportXml.BuckOfferStatus.BuckChildDigital
          bDisableEditName = False
          bDisableEditTierType = True
          bDisplayBuckTierType = False
        Case Copient.ImportXml.BuckOfferStatus.BuckTiered
          bDisableEditName = False
          bDisableEditTierType = False
          bDisplayBuckTierType = True
        Case Else
          bDisableEditName = False
          bDisableEditTierType = False
          bDisplayBuckTierType = True
          If oBuckStatus = Copient.ImportXml.BuckOfferStatus.ErrorOccurred Then
            infoMessage = MyImport.GetErrorMsg()
          End If
      End Select

      ' Get 1st buck PromoVarID for display  
      MyCommon.QueryStr = "Select PromoVarID from PointsPrograms with (NoLock) where BuckTierNumber=1 and BuckOfferID=" & OfferID & ";"
      rst = MyCommon.LRT_Select
      If rst.Rows.Count > 0 Then
        lBuckPromoVarId = MyCommon.NZ(rst.Rows(0).Item("PromoVarID"), 0)
      Else
        lBuckPromoVarId = 0
      End If
    End If
  End If
  
  If (Request.QueryString("save") <> "") AndAlso Not (MyCommon.Fetch_SystemOption(191) AndAlso NumberofFolders > 1) Then
    ' user wants to save something, so run stored procedure for saving group
    form_OfferID = Request.QueryString("form_OfferID")
	
    Dim Descfromspecial As String = ""
    Dim dtDesc As DataTable
    MyCommon.QueryStr = "Select Description from Offers with (NoLock) where OfferID=" & form_OfferID
    dtDesc = MyCommon.LRT_Select
    If dtDesc.Rows.Count > 0 Then
      Descfromspecial = MyCommon.NZ(dtDesc.Rows(0)(0), "")
    End If
    If Descfromspecial <> "" Then
      Descript = Descfromspecial.Replace("&quot;", Chr(34))
      If Descript.Length <= 1000 Then
        DescriptLength = True
      End If
    Else
      DescriptLength = True
    End If
    
    If AllowSpecialCharacters <> "" Then
      form_Name = Logix.TrimAll(Request.QueryString("form_Name"))
      form_Description = Descfromspecial
    Else
      form_Name = Logix.TrimAll(Request.QueryString("form_Name"))
      form_Description = Descfromspecial
    End If
    
    form_Category = Request.QueryString("form_Category")
    form_EmployeeFiltering = Request.QueryString("EmployeeFiltering")    ' EmployeeFiltering
    form_NonEmployeesOnly = Request.QueryString("eligradio")     ' eligradio button
    
    form_TestStartdate = MyCommon.NZ(Request.QueryString("form_TestStartDate"), Date.Now.ToShortDateString)
    form_TestEnddate = MyCommon.NZ(Request.QueryString("form_TestEnddate"), Date.Now.ToShortDateString)
    form_ProdStartdate = Request.QueryString("form_ProdStartdate")
    form_ProdEnddate = Request.QueryString("form_ProdEnddate")
    InboundCRMEngineID = MyCommon.Extract_Val(Request.QueryString("InboundCRMEngineID"))
    
    If bUseDisplayDates Then
      form_DispStartdate = Request.QueryString("form_DispStartdate")
      form_DispEnddate = Request.QueryString("form_DispEnddate")
                
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
        form_DispEnddate = Request.QueryString("form_DispStartdate") & " " & form_DispEndHr & ":" & form_DispEndMin & ":00"
      End If

      If form_DispStartdate <> "" And form_DispEnddate = "" Then
        form_DispEnddate = Request.QueryString("form_DispStartdate")
      End If
             
    End If
    If bAllowTimeWithStartEndDates Then
      form_TestStartHr = MyCommon.NZ(Request.QueryString("form_TestStartHr"), "00")
      form_TestStartMin = MyCommon.NZ(Request.QueryString("form_TestStartMin"), "00")
      form_TestStartdate = form_TestStartdate & " " & form_TestStartHr & ":" & form_TestStartMin & ":00"
      
      form_TestEndHr = MyCommon.NZ(Request.QueryString("form_TestEndHr"), "23")
      form_TestEndMin = MyCommon.NZ(Request.QueryString("form_TestEndMin"), "59")
      form_TestEnddate = form_TestEnddate & " " & form_TestEndHr & ":" & form_TestEndMin & ":00"

      form_ProdStartHr = MyCommon.NZ(Request.QueryString("form_ProdStartHr"), "00")
      form_ProdStartMin = MyCommon.NZ(Request.QueryString("form_ProdStartMin"), "00")
      form_ProdStartdate = form_ProdStartdate & " " & form_ProdStartHr & ":" & form_ProdStartMin & ":00"
      
      form_ProdEndHr = MyCommon.NZ(Request.QueryString("form_ProdEndHr"), "23")
      form_ProdEndMin = MyCommon.NZ(Request.QueryString("form_ProdEndMin"), "59")
      form_ProdEnddate = form_ProdEnddate & " " & form_ProdEndHr & ":" & form_ProdEndMin & ":00"
    End If
    
    form_TierTypeID = Request.QueryString("form_TierTypeID")
    form_NumTiers = MyCommon.Extract_Val(Request.QueryString("form_NumTiers"))
    form_EngineID = Request.QueryString("form_EngineID")
    form_ExtOfferID2 = Request.QueryString("form_ExtOfferID2")
    form_CRMEngineID = Request.QueryString("form_CRMEngineID")
    form_Priority = MyCommon.Extract_Val(Request.QueryString("form_Priority"))
    If (form_Priority = 0) Then
      form_Priority = 1
    End If
    form_InstantWin = Request.QueryString("form_InstantWin")
    form_NumPrizesAllowed = MyCommon.Extract_Val(Request.QueryString("form_NumPrizesAllowed"))
    form_OddsOfWinning = MyCommon.Extract_Val(Request.QueryString("form_OddsOfWinning"))
    form_RandomWinners = Request.QueryString("form_RandomWinners")
    form_IWTransLevel = Request.QueryString("form_IWTransLevel")
    form_DisabledOnCFW = Request.QueryString("form_disabledonCFW")
    form_DisplayOnWebKiosk = Request.QueryString("form_displayOnWebKiosk")
    form_ExportToEDW = Request.QueryString("form_exportEDW")
    form_Favorite = Request.QueryString("form_favorite")
    form_AutoTransferable = Request.QueryString("form_autotransferable")
    form_HighValue = Request.QueryString("form_highvalue")
    form_DisableDisconnectedOffer = Request.QueryString("form_disabledisconnectedoffer")
    form_Fob = Request.QueryString("form_fob")
    form_AutoTranslate = Request.QueryString("form_autotranslate")
    Dim form_eCoupon As String = Request.QueryString("form_eCoupon")
    
    form_DistPeriod = MyCommon.Extract_Val(Request.QueryString("limitperiod"))
    'Get Folder Start/End dates for the validation that user can not change the offer date which is not falling in range of folder dates.
    '192 (Offer dates should be within folder dates)
    If (MyCommon.Fetch_SystemOption(192) = "1") Then
      MyCommon.QueryStr = "Select Startdate,Enddate from Folders Fs inner join FolderItems FI on FI.folderid=Fs.folderid " & _
                          " where FI.linkid=" & form_OfferID
      rst = MyCommon.LRT_Select
      If (rst.Rows.Count > 0) Then
        If (Not IsDBNull(rst.Rows(0).Item("Startdate")) OrElse Not IsDBNull(rst.Rows(0).Item("EndDate"))) Then
          FolderStartDate = rst.Rows(0).Item("Startdate")
          FolderEndDate = rst.Rows(0).Item("Enddate")
        End If
      End If
    End If
    'If Request.QueryString("limitvalue") <> "" AndAlso (Not Integer.TryParse(Request.QueryString("limitvalue"), form_DistPeriodLimit) OrElse form_DistPeriodLimit < 0) Then
    ' infoMessage = Copient.PhraseLib.Lookup("offer-gen.badlimit", LanguageID)
    'End If

    If Request.QueryString("limitvalue") <> "" Then
      If IsNumeric(Request.QueryString("limitvalue")) = False Then
        infoMessage = Copient.PhraseLib.Lookup("offer-gen.badlimit", LanguageID)
      ElseIf CInt(Request.QueryString("limitvalue")) < 0 Then
        infoMessage = Copient.PhraseLib.Lookup("offer-gen.badlimit", LanguageID)
      ElseIf Convert.ToDecimal(Request.QueryString("limitvalue")) <> CInt(Request.QueryString("limitvalue")) Then
        infoMessage = Copient.PhraseLib.Lookup("offer-gen.badlimit", LanguageID)
      Else
        form_DistPeriodLimit = CInt(Request.QueryString("limitvalue"))
      End If
    End If

    If (Request.QueryString("selectadv") <> "") Then
      AdvancedLimitID = MyCommon.NZ(Request.QueryString("selectadv"), 0)
      If AdvancedLimitID > 0 Then
        MyCommon.QueryStr = "select AL.PromoVarID,AL.LimitTypeID, AL.LimitValue, AL.LimitPeriod " & _
                            "from CM_AdvancedLimits as AL with (NoLock) where Deleted=0 and LimitID='" & AdvancedLimitID & "';"
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
          VarID = MyCommon.NZ(rst.Rows(0).Item("PromoVarID"), 0)
          form_DistPeriodLimit = MyCommon.NZ(rst.Rows(0).Item("LimitValue"), 0)
          form_DistPeriod = MyCommon.NZ(rst.Rows(0).Item("LimitPeriod"), 0)
        End If
      End If
    End If
    
    'If the limit is saved as days, do not allow the period to be 0
    If (MyCommon.Extract_Val(Request.QueryString("selectday")) = 1) Then
      If Not form_DistPeriod > 0 Then ValidDayLimit = False
    End If
    
    ' someone saves, lets do the special case and set a promo variable if the distribution is
    ' not equal to zero and the promo variable doesnt already exist
    If (AdvancedLimitID = 0 And form_DistPeriodLimit <> 0) Then
      MyCommon.Open_LogixXS()
      MyCommon.QueryStr = "select PromoVarID, VarTypeID, LinkID " & _
                          "from PromoVariables with (NoLock) where Deleted=0 and VarTypeID=1 and LinkID=" & form_OfferID & ";"
      rst = MyCommon.LXS_Select
      If (rst.Rows.Count > 0) Then
        VarID = MyCommon.NZ(rst.Rows(0).Item("PromoVarID"), 0)
      Else
        MyCommon.QueryStr = "dbo.pc_DistributionVar_Create"
        MyCommon.Open_LXSsp()
        MyCommon.LXSsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = form_OfferID
        MyCommon.LXSsp.Parameters.Add("@VarID", SqlDbType.BigInt).Direction = ParameterDirection.Output
        MyCommon.LXSsp.ExecuteNonQuery()
        VarID = MyCommon.LXSsp.Parameters("@VarID").Value
        MyCommon.Close_LXSsp()
      End If
      MyCommon.Close_LogixXS()
    End If

    MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where offerid=" & form_OfferID
    MyCommon.LRT_Execute()
    
    ' get the existing number of tiers
    MyCommon.QueryStr = "select NumTiers from offers with (NoLock) where offerID=" & form_OfferID
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      NumTiers = rst.Rows(0).Item("NumTiers")
    End If
    MyCommon.QueryStr = "select * from FolderItems where LinkID=" & OfferID
    rst = MyCommon.LRT_Select
    NumberofFolders = rst.Rows.Count
    ' check for an incentive already with that name
    MyCommon.QueryStr = "SELECT Name FROM Offers with (NoLock) WHERE Name = '" & MyCommon.Parse_Quotes(form_Name) & "' AND Deleted=0 and OfferID <> " & form_OfferID & ";"
    rst = MyCommon.LRT_Select
	
    DuplicateName = (rst.Rows.Count > 0)
    
    ' check if the offer already has at least one preference condition 
    MyCommon.QueryStr = "select ConditionID from OfferConditions with (NoLock) where OfferID = " & form_OfferID & " and ConditionTypeID=100 and Deleted=0;"
    rst = MyCommon.LRT_Select
    HasPrefCondition = (rst.Rows.Count > 0)

    ' Also, run a query to see if there's a category that has this offer as its base offer
    MyCommon.QueryStr = "select OfferCategoryID from OfferCategories where Deleted=0 and BaseOfferID=" & form_OfferID & " and OfferCategoryID=(" & _
                        "  select IsNull(OfferCategoryID, 0) from Offers where OfferID=" & form_OfferID & ");"
    rst2 = MyCommon.LRT_Select
    
    'Check for an external id for this creation source already exists
    MyCommon.QueryStr = "SELECT OfferID FROM Offers with (NoLock) WHERE Deleted=0 and OfferID <> " & form_OfferID & " and ExtOfferID = '" & Request.QueryString("hdnExtOfferId") & "' AND InboundCRMEngineID= " & InboundCRMEngineID
    rst = MyCommon.LRT_Select
    IsExtIdInboundIdExist = (rst.Rows.Count > 0)
    
    If (form_Name = "") Then
      infoMessage = Copient.PhraseLib.Lookup("offer-gen.noname", LanguageID)
      FormValid = False
    ElseIf (DuplicateName) Then
      infoMessage = Copient.PhraseLib.Lookup("offer-gen.nameused", LanguageID)
      FormValid = False
    ElseIf (Not IsDate(form_ProdStartdate)) Then
      infoMessage = Copient.PhraseLib.Lookup("offer-gen.invalidstartdate", LanguageID)
      FormValid = False
    ElseIf (Not IsDate(form_ProdEnddate)) Then
      infoMessage = Copient.PhraseLib.Lookup("offer-gen.invalidenddate", LanguageID)
      FormValid = False
    ElseIf (bUseTestDates AndAlso (Not IsDate(form_TestStartdate))) Then
      infoMessage = Copient.PhraseLib.Lookup("offer-gen.invalidteststartdate", LanguageID)
      FormValid = False
    ElseIf (bUseTestDates AndAlso (Not IsDate(form_TestEnddate))) Then
      infoMessage = Copient.PhraseLib.Lookup("offer-gen.invalidtestenddate", LanguageID)
      FormValid = False
    ElseIf bUseDisplayDates And (Not IsDate(form_DispStartdate) AndAlso IsDate(form_DispEnddate)) Then

      infoMessage = Copient.PhraseLib.Lookup("offer-gen.invaliddispstartdate", LanguageID)
      FormValid = False

            'ElseIf (form_NumTiers < NumTiers) Then
            '    infoMessage = Copient.PhraseLib.Lookup("offer-gen.invalidtiers", LanguageID)
            '    FormValid = False
        ElseIf DescriptLength = False Then
            infoMessage = Copient.PhraseLib.Lookup("error.description", LanguageID)
            FormValid = False
        ElseIf (form_NumTiers > 1 AndAlso (form_CRMEngineID = 1 OrElse form_CRMEngineID = 2)) Then
            infoMessage = Copient.PhraseLib.Lookup("offer-gen.invalidoutbound", LanguageID)
            FormValid = False
        ElseIf (rst2.Rows.Count > 0) AndAlso (form_Category <> MyCommon.NZ(rst2.Rows(0).Item("OfferCategoryID"), 0)) Then
            infoMessage = Copient.PhraseLib.Lookup("ueoffer-gen.InvalidCategoryChange", LanguageID)
            FormValid = False
        ElseIf Not ValidDayLimit Then
            infoMessage = Copient.PhraseLib.Lookup("offer-gen.invalidlimitperiod", LanguageID)
            FormValid = False
        ElseIf form_OddsOfWinning > Int16.MaxValue Or form_OddsOfWinning < Int16.MinValue Then
            infoMessage = Copient.PhraseLib.Lookup("offer-gen.InvalidOdds", LanguageID)
            FormValid = False
        ElseIf form_NumTiers > 0 AndAlso HasPrefCondition Then
            infoMessage = Copient.PhraseLib.Lookup("offer-gen.PrefConTierInvalid", LanguageID)
            FormValid = False
        ElseIf IsExtIdInboundIdExist Then
            infoMessage = Copient.PhraseLib.Lookup("offer-gen.extidinboundidexist", LanguageID)
            FormValid = False
        ElseIf ValidateTiers(form_OfferID, form_TierTypeID, form_NumTiers) Then
            infoMessage = Copient.PhraseLib.Lookup("term.mustDltExistingConAndRew", LanguageID)
      FormValid = False
    Else
      If (form_DistPeriodLimit < 0) Then
        form_DistPeriodLimit = 0
      End If
      If (form_Priority < 0) Then
        form_Priority = 1
      End If
      If (form_NumTiers < 0) Then
        form_NumTiers = 0
      End If
      If (form_TierTypeID > 0) And (form_NumTiers < 1) Then
        form_NumTiers = 0
        form_TierTypeID = 0
      End If
      If (form_NumPrizesAllowed < 0) Then
        form_NumPrizesAllowed = 0
      End If
      If (form_OddsOfWinning < 0) Then
        form_OddsOfWinning = 0
      End If
      If (bUseTestDates AndAlso (CDate(form_TestEnddate) < CDate(form_TestStartdate))) Then
        infoMessage = Copient.PhraseLib.Lookup("offer-gen.baddate", LanguageID)
        FormValid = False
      ElseIf (CDate(form_ProdEnddate) < CDate(form_ProdStartdate)) Then
        infoMessage = Copient.PhraseLib.Lookup("offer-gen.baddate", LanguageID)
        FormValid = False
      ElseIf bUseOfferRedemptionThreshold AndAlso (Not Integer.TryParse(Request.QueryString("OfferRedemptionThresholdperHour"), 0)) Then
        infoMessage = Copient.PhraseLib.Lookup("offer-gen.badredemptionthresholdperhour", LanguageID)
        FormValid = False
      ElseIf bUseOfferRedemptionThreshold AndAlso (CInt(Request.QueryString("OfferRedemptionThresholdperHour")) < 0) Then
        infoMessage = Copient.PhraseLib.Lookup("offer-gen.badredemptionthresholdperhour", LanguageID)
        FormValid = False
      ElseIf bUseDisplayDates AndAlso (form_DispStartdate <> "" AndAlso form_DispEnddate <> "") AndAlso (CDate(form_DispEnddate) < CDate(form_DispStartdate)) Then
        infoMessage = Copient.PhraseLib.Lookup("offer-gen.dispbaddate", LanguageID)
        FormValid = False
      ElseIf (CDate(form_TestStartdate) < CDate(IIf(FolderStartDate = "", CDate(form_TestStartdate), FolderStartDate)) OrElse CDate(form_TestEnddate) > CDate(IIf(FolderEndDate = "", CDate(form_TestEnddate), FolderEndDate)) OrElse CDate(form_ProdStartdate) < CDate(IIf(FolderStartDate = "", CDate(form_ProdStartdate), FolderStartDate)) OrElse CDate(form_ProdEnddate) > CDate(IIf(FolderEndDate = "", CDate(form_ProdEnddate), FolderEndDate))) AndAlso (MyCommon.Fetch_SystemOption(192) = "1") Then
        infoMessage = Copient.PhraseLib.Lookup("folders.OfferNotInFolderDateRange", LanguageID) & "(" & FolderStartDate & " - " & FolderEndDate & ")"
      Else
        MyCommon.LRT_Execute()
      End If
    End If
    
    If (form_OfferID = 0) Then
      ' first things first if they entered tiers then we need to set the tiertype for them
      If (form_NumTiers > 1 And form_TierTypeID = 0) Then
        form_TierTypeID = 1
      End If
      ' get the group created and find out what its ID is.
      MyCommon.QueryStr = "dbo.pt_Offers_Insert"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 200).Value = form_Name
      MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = form_EngineID
      MyCommon.LRTsp.Parameters.Add("@EngineSubTypeID", SqlDbType.Int).Value = EngineSubTypeID
      MyCommon.LRTsp.Parameters.Add("@IsTemplate", SqlDbType.Bit).Value = IIf(Request.QueryString("IsTemplate") = "IsTemplate", 1, 0)
      MyCommon.LRTsp.Parameters.Add("@AdminID", SqlDbType.Int).Value = AdminUserID
      MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Direction = ParameterDirection.Output

      If (form_Name = "") Then
        infoMessage = Copient.PhraseLib.Lookup("offer-gen.noname", LanguageID)
      Else
        MyCommon.LRTsp.ExecuteNonQuery()
        OfferID = MyCommon.LRTsp.Parameters("@OfferID").Value
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-create", LanguageID))
      End If
      MyCommon.Close_LRTsp()
      If (Request.QueryString("IsTemplate") = "IsTemplate") Then
        ' update the template bit on the created offer
        MyCommon.QueryStr = "update Offers with (RowLock) set IsTemplate=1 where OfferID=" & OfferID
        MyCommon.LRT_Execute()
        MyCommon.QueryStr = "insert into TemplatePermissions with (RowLock) (offerid) values(" & OfferID & ")"
        MyCommon.LRT_Execute()
        
        ' time to update the status bits for the templates
        Dim form_Disallow_EmployeeFiltering As Integer = 0
        Dim form_Disallow_ProductionDates As Integer = 0
        Dim form_Disallow_Limits As Integer = 0
        Dim form_Disallow_Tiers As Integer = 0
        Dim form_Disallow_Priority As Integer = 0
        Dim form_Disallow_Sweepstakes As Integer = 0
        Dim form_Disallow_ExecutionEngine As Integer = 0
        Dim form_Disallow_CRMEngine As Integer = 0
        Dim form_Disallow_DisplayDates As Integer = 0
        Dim form_Disallow_UserDefinedFields As Integer = 0
        Dim form_Disallow_AdvancedOption As Integer = 0
        Dim form_Disallow_OfferRedempThreshold As Integer = 0
                
        If (Request.QueryString("Disallow_ExecutionEngine") = "on") Then
          form_Disallow_ExecutionEngine = 1
        End If
        If (Request.QueryString("Disallow_CRMEngine") = "on") Then
          form_Disallow_CRMEngine = 1
        End If
        If (Request.QueryString("Disallow_EmployeeFiltering") = "on") Then
          form_Disallow_EmployeeFiltering = 1
        End If
        If (Request.QueryString("Disallow_ProductionDates") = "on") Then
          form_Disallow_ProductionDates = 1
        End If
        If (Request.QueryString("Disallow_Limits") = "on") Then
          form_Disallow_Limits = 1
        End If
        If (Request.QueryString("Disallow_Tiers") = "on") Then
          form_Disallow_Tiers = 1
        End If
        If (Request.QueryString("Disallow_Priority") = "on") Then
          form_Disallow_Priority = 1
        End If
        If (Request.QueryString("Disallow_Sweepstakes") = "on") Then
          form_Disallow_Sweepstakes = 1
        End If
        If (Request.QueryString("Disallow_UserDefinedFields") = "on") Then
          form_Disallow_UserDefinedFields = 1
        End If
        If bUseDisplayDates Then
          If (Request.QueryString("Disallow_DisplayDates") = "on") Then
            form_Disallow_DisplayDates = 1
          End If
        End If
                
        If (bUseOfferRedemptionThreshold AndAlso Request.QueryString("Disallow_OfferRedempThreshold") = "on") Then
          form_Disallow_OfferRedempThreshold = 1
        End If
                
        If (Request.QueryString("Disallow_AdvancedOption") = "on") Then
          form_Disallow_AdvancedOption = 1
        End If
                
        MyCommon.QueryStr = "update TemplatePermissions with (RowLock) set Disallow_EmployeeFiltering=" & form_Disallow_EmployeeFiltering & _
                            " , Disallow_ProductionDates=" & form_Disallow_ProductionDates & _
                            " , Disallow_limits=" & form_Disallow_Limits & _
                            " , Disallow_Tiers=" & form_Disallow_Tiers & _
                            " , Disallow_Priority=" & form_Disallow_Priority & _
                            " , Disallow_CRMEngine=" & form_Disallow_CRMEngine & _
                            " , Disallow_ExecutionEngine=" & form_Disallow_ExecutionEngine & _
                            " , Disallow_Sweepstakes=" & form_Disallow_Sweepstakes & _
       " , Disallow_UserDefinedFields=" & form_Disallow_UserDefinedFields & _
       " , Disallow_AdvancedOption=" & form_Disallow_AdvancedOption
        
        If bUseDisplayDates Then
          MyCommon.QueryStr = MyCommon.QueryStr & " , Disallow_DisplayDates=" & form_Disallow_DisplayDates
        End If
                  
        If bUseOfferRedemptionThreshold Then
          MyCommon.QueryStr = MyCommon.QueryStr & " , Disallow_OfferRedempThreshold=" & form_Disallow_OfferRedempThreshold
        End If
               
                
        MyCommon.QueryStr = MyCommon.QueryStr & " where OfferID=" & OfferID
                
        MyCommon.LRT_Execute()
      End If
    ElseIf (Not FormValid) Then
      OfferID = form_OfferID
    Else
      ' ok this is an update to the offer.  We need to check if the tiers is changing from the already set value
      ' lets get the current tiers from the database
      ' first things first if they entered tiers then we need to set the tiertype for them
      If (form_NumTiers > 1 And form_TierTypeID = 0) Then
        form_TierTypeID = 1
      End If
      MyCommon.QueryStr = "select NumTiers from offers with (NoLock) where offerID=" & form_OfferID
      rst = MyCommon.LRT_Select
      If (rst.Rows.Count > 0) Then
        Dim lastTierNum = rst.Rows(0).Item("NumTiers")
        If (lastTierNum > form_NumTiers) Then
          ' The new number of tiers is lower than the old, so we need to update the existing rewards and conditions
          ' First, conditions
          ' Find out which are still in place and are tiered
          MyCommon.QueryStr = "select ConditionID from offerconditions with (NoLock) where tiered=1 and deleted=0 and offerid=" & form_OfferID
          rst = MyCommon.LRT_Select
          For Each row In rst.Rows
            ' Here are the conditions we need to update by reducing the number of tiers
            MyCommon.QueryStr = "delete from ConditionTiers with (RowLock) where ConditionID =" & row.Item("ConditionID") & " and TierLevel > " & form_NumTiers
            MyCommon.LRT_Execute()
          Next
          ' Next, rewards
          ' Find out which are still in place and are tiered
          MyCommon.QueryStr = "select RewardID,linkID from offerrewards with (NoLock) where  tiered=1 and deleted=0 and offerid=" & form_OfferID
          rst = MyCommon.LRT_Select
          For Each row In rst.Rows
            ' Here are the rewards we need to update by reducing the number of tiers
            MyCommon.QueryStr = "delete from RewardTiers with (RowLock) where RewardID =" & row.Item("RewardID") & " and TierLevel > " & form_NumTiers
            MyCommon.LRT_Execute()
            MyCommon.QueryStr = "delete from RewardCustomerGroupTiers with (RowLock) where RewardID =" & row.Item("RewardID") & " and TierLevel > " & form_NumTiers
            MyCommon.LRT_Execute()
            MyCommon.QueryStr = "delete from CashierMessageTiers with (RowLock) where MessageID =" & row.Item("LinkID") & " and TierLevel > " & form_NumTiers
            MyCommon.LRT_Execute()
            MyCommon.QueryStr = "delete from PrintedMessageTiers with (RowLock) where MessageID =" & row.Item("LinkID") & " and TierLevel > " & form_NumTiers
            MyCommon.LRT_Execute()
          Next
        ElseIf (lastTierNum < form_NumTiers) Then
          ' The user wants to increase the offer's tiers, so we have to add a blank one to each tier condition
          ' First, conditions
          ' Find out which are still in place and are tiered
          MyCommon.QueryStr = "select ConditionID from offerconditions with (NoLock) where tiered=1 and deleted=0 and offerid=" & form_OfferID
          rst = MyCommon.LRT_Select
          For Each row In rst.Rows
            ' Here are the conditions we need to update by adding tiers
            For iterate = lastTierNum + 1 To form_NumTiers
              MyCommon.QueryStr = "insert into ConditionTiers with (RowLock) (ConditionID,TierLevel,AmtRequired) values (" & row.Item("ConditionID") & "," & iterate & ",0)"
              MyCommon.LRT_Execute()
            Next
          Next
          ' Next, rewards
          ' Find out which are still in place and are tiered
          MyCommon.QueryStr = "select RewardID,linkID from offerrewards with (NoLock) where  tiered=1 and deleted=0 and offerid=" & form_OfferID
          rst = MyCommon.LRT_Select
          For Each row In rst.Rows
            For iterate = lastTierNum + 1 To form_NumTiers
              ' Here are the rewards we need to update by adding tiers
              MyCommon.QueryStr = "insert into RewardTiers with (RowLock) (RewardID,TierLevel,RewardAmount) values (" & row.Item("RewardID") & "," & iterate & ",0)"
              MyCommon.LRT_Execute()
              MyCommon.QueryStr = "insert into RewardCustomerGroupTiers with (RowLock) (RewardID,TierLevel,CustomerGroupID) values (" & row.Item("RewardID") & "," & iterate & ",0)"
              MyCommon.LRT_Execute()
              MyCommon.QueryStr = "insert into CashierMessageTiers with (RowLock) (MessageID,TierLevel,Line1Text) values (" & row.Item("linkID") & "," & iterate & ",'')"
              MyCommon.LRT_Execute()
              MyCommon.QueryStr = "insert into PrintedMessageTiers with (RowLock) (MessageID,TierLevel,BodyText) values (" & row.Item("linkID") & "," & iterate & ",'')"
              MyCommon.LRT_Execute()
            Next
          Next
        End If
      End If
      
      MyCommon.QueryStr = "dbo.pt_Offers_Update"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 200).Value = form_Name
      MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 1000).Value = form_Description
      MyCommon.LRTsp.Parameters.Add("@OfferCategoryID", SqlDbType.Int).Value = form_Category
      
      If (form_EmployeeFiltering = "on") Then
        MyCommon.LRTsp.Parameters.Add("@EmployeeFiltering", SqlDbType.Bit).Value = 1
      Else
        MyCommon.LRTsp.Parameters.Add("@EmployeeFiltering", SqlDbType.Bit).Value = 0
      End If
      If (form_NonEmployeesOnly = "1") Then
        MyCommon.LRTsp.Parameters.Add("@NonEmployeesOnly", SqlDbType.Bit).Value = 1
      Else
        MyCommon.LRTsp.Parameters.Add("@NonEmployeesOnly", SqlDbType.Bit).Value = 0
      End If
      
      MyCommon.LRTsp.Parameters.Add("@TestStartDate", SqlDbType.DateTime).Value = form_TestStartdate
      MyCommon.LRTsp.Parameters.Add("@TestEndDate", SqlDbType.DateTime).Value = form_TestEnddate
      MyCommon.LRTsp.Parameters.Add("@ProdStartDate", SqlDbType.DateTime).Value = form_ProdStartdate
      MyCommon.LRTsp.Parameters.Add("@ProdEndDate", SqlDbType.DateTime).Value = form_ProdEnddate
      MyCommon.LRTsp.Parameters.Add("@DistPeriod", SqlDbType.Int).Value = form_DistPeriod
      MyCommon.LRTsp.Parameters.Add("@DistPeriodLimit", SqlDbType.Int).Value = form_DistPeriodLimit
      MyCommon.LRTsp.Parameters.Add("@TierTypeID", SqlDbType.Int).Value = form_TierTypeID
      MyCommon.LRTsp.Parameters.Add("@NumTiers", SqlDbType.Int).Value = form_NumTiers
      MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = form_EngineID
      MyCommon.LRTsp.Parameters.Add("@CRMEngineID", SqlDbType.Int).Value = form_CRMEngineID
      MyCommon.LRTsp.Parameters.Add("@ExtOfferID2", SqlDbType.NVarChar, 40).Value = form_ExtOfferID2
      
      If (form_InstantWin = "on") Then
        MyCommon.LRTsp.Parameters.Add("@InstantWin", SqlDbType.Bit).Value = 1
      Else
        MyCommon.LRTsp.Parameters.Add("@InstantWin", SqlDbType.Bit).Value = 0
      End If
      MyCommon.LRTsp.Parameters.Add("@NumPrizesAllowed", SqlDbType.Int).Value = MyCommon.Extract_Val(form_NumPrizesAllowed)
      MyCommon.LRTsp.Parameters.Add("@OddsOfWinning", SqlDbType.Int).Value = MyCommon.Extract_Val(form_OddsOfWinning)
      If (form_RandomWinners = "1") Then
        MyCommon.LRTsp.Parameters.Add("@RandomWinners", SqlDbType.Bit).Value = 1
      Else
        MyCommon.LRTsp.Parameters.Add("@RandomWinners", SqlDbType.Bit).Value = 0
      End If
      If (form_IWTransLevel = "1") Then
        MyCommon.LRTsp.Parameters.Add("@IWTransLevel", SqlDbType.Bit).Value = 1
      Else
        MyCommon.LRTsp.Parameters.Add("@IWTransLevel", SqlDbType.Bit).Value = 0
      End If
      If (form_DisabledOnCFW = "on") Then
        MyCommon.LRTsp.Parameters.Add("@DisabledOnCFW", SqlDbType.Bit).Value = 1
        UseLegacyflag = True
      Else
        MyCommon.LRTsp.Parameters.Add("@DisabledOnCFW", SqlDbType.Bit).Value = 0
      End If
      If (form_DisplayOnWebKiosk = "on") Then
        MyCommon.LRTsp.Parameters.Add("@DisplayOnWebKiosk", SqlDbType.Bit).Value = 1
        UseLegacyflag = True
      Else
        MyCommon.LRTsp.Parameters.Add("@DisplayOnWebKiosk", SqlDbType.Bit).Value = 0
      End If
      If UseLegacyflag Then
        MyCommon.LRTsp.Parameters.Add("@UseLegacyWebKiosk", SqlDbType.Bit).Value = 1
      Else
        MyCommon.LRTsp.Parameters.Add("@UseLegacyWebKiosk", SqlDbType.Bit).Value = 0
      End If
      If (form_ExportToEDW = "on") Then
        MyCommon.LRTsp.Parameters.Add("@ExportToEDW", SqlDbType.Bit).Value = 1
      Else
        MyCommon.LRTsp.Parameters.Add("@ExportTOEDW", SqlDbType.Bit).Value = 0
      End If
      If (form_Favorite = "on") Then
        MyCommon.LRTsp.Parameters.Add("@Favorite", SqlDbType.Bit).Value = 1
      Else
        MyCommon.LRTsp.Parameters.Add("@Favorite", SqlDbType.Bit).Value = 0
      End If
      If (form_AutoTransferable = "on") Then
        MyCommon.LRTsp.Parameters.Add("@AutoTransferable", SqlDbType.Bit).Value = 1
      Else
        MyCommon.LRTsp.Parameters.Add("@AutoTransferable", SqlDbType.Bit).Value = 0
      End If
      If (form_HighValue = "on") Then
        MyCommon.LRTsp.Parameters.Add("@HighValue", SqlDbType.Bit).Value = 1
      Else
        MyCommon.LRTsp.Parameters.Add("@HighValue", SqlDbType.Bit).Value = 0
      End If
      If (form_DisableDisconnectedOffer = "on") Then
        MyCommon.LRTsp.Parameters.Add("@DisableDisconnectedOffer", SqlDbType.Bit).Value = 1
      Else
        MyCommon.LRTsp.Parameters.Add("@DisableDisconnectedOffer", SqlDbType.Bit).Value = 0
      End If

      
      If (form_eCoupon = "on") Then
        MyCommon.LRTsp.Parameters.Add("@OfferTypeID", SqlDbType.Int).Value = gintEcouponOfferTypeID
        'stored procedure will default to 1 for Standard offer is not sent
      End If
	  
      MyCommon.LRTsp.Parameters.Add("@CashierApprovalMessage", SqlDbType.Bit).Value = IIf(Request.QueryString("cash_prompt_rew") = "on", True, False)
     
      MyCommon.LRTsp.Parameters.Add("@PriorityLevel", SqlDbType.Int).Value = form_Priority
      MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = form_OfferID
      If Request.QueryString("InboundCRMEngineID") <> "" Then
        MyCommon.LRTsp.Parameters.Add("@InboundCRMEngineId", SqlDbType.Int).Value = InboundCRMEngineID
      End If
      
      If bFobEligibilityEnabled AndAlso (InboundCRMEngineID > 0) Then
        If (form_Fob = "on") Then
          MyCommon.LRTsp.Parameters.Add("@FobEligible", SqlDbType.Bit).Value = 1
        Else
          MyCommon.LRTsp.Parameters.Add("@FobEligible", SqlDbType.Bit).Value = 0
        End If
      End If

      If (form_Name = "") Then
      Else
        MyCommon.LRTsp.ExecuteNonQuery()
      End If
      
      MyCommon.Close_LRTsp()
      
      If bUseDisplayDates Then
        MyCommon.QueryStr = "dbo.pa_UpdateOfferAccessoryFields"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = form_OfferID
        MyCommon.LRTsp.Parameters.Add("@DisplayStartTime", SqlDbType.DateTime).Value = IIf(String.IsNullOrEmpty(form_DispStartdate), DBNull.Value, form_DispStartdate)
        MyCommon.LRTsp.Parameters.Add("@DisplayEndTime", SqlDbType.DateTime).Value = IIf(String.IsNullOrEmpty(form_DispEnddate), DBNull.Value, form_DispEnddate)
                            
        MyCommon.LRTsp.Parameters.Add("@RedemThresholdPerHour", SqlDbType.BigInt).Value = DBNull.Value
        MyCommon.LRTsp.Parameters.Add("@AdminUserId", SqlDbType.BigInt).Value = AdminUserID
        MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.BigInt).Value = EngineID
                
        MyCommon.LRTsp.Parameters.Add("@OptionID", SqlDbType.BigInt).Value = 85
        MyCommon.LRTsp.ExecuteNonQuery()
        MyCommon.Close_LRTsp()
      End If

      If (bUseOfferRedemptionThreshold) Then
        OfferRedemptionThresholdperHour = Convert.ToInt64(Request.QueryString("OfferRedemptionThresholdperHour"))
        MyCommon.QueryStr = "dbo.pa_UpdateOfferAccessoryFields"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = form_OfferID
        MyCommon.LRTsp.Parameters.Add("@DisplayStartTime", SqlDbType.DateTime).Value = DBNull.Value
        MyCommon.LRTsp.Parameters.Add("@DisplayEndTime", SqlDbType.DateTime).Value = DBNull.Value
        'Updating only OfferRedemptionThresholdperHour because it depends on CM SystemOption 'pa_UpdateOfferAccessoryFields' contains logic to insert/update based on the engineID and systemoption passed     
        MyCommon.LRTsp.Parameters.Add("@RedemThresholdPerHour", SqlDbType.BigInt).Value = OfferRedemptionThresholdperHour
        MyCommon.LRTsp.Parameters.Add("@AdminUserId", SqlDbType.BigInt).Value = AdminUserID
        MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.BigInt).Value = EngineID
        MyCommon.LRTsp.Parameters.Add("@OptionID", SqlDbType.BigInt).Value = 83
        MyCommon.LRTsp.ExecuteNonQuery()
        MyCommon.Close_LRTsp()
                
      End If
      OfferID = form_OfferID
	  
      If MyCommon.Fetch_SystemOption(156) = "1" Then
%>
<udf:udfsave ID="udfsavecontrol" runat="server" />
<%
  infoMessage = udfsavecontrol.infoMessage
  UDFHistory = udfsavecontrol.UDFHistory
End If ' MyCommon.Fetch_SystemOption(156) = "1" 
               
If (Request.QueryString("IsTemplate") = "IsTemplate") Then
  ' time to update the status bits for the templates
  Dim form_Disallow_EmployeeFiltering As Integer = 0
  Dim form_Disallow_ProductionDates As Integer = 0
  Dim form_Disallow_Limits As Integer = 0
  Dim form_Disallow_Tiers As Integer = 0
  Dim form_Disallow_Priority As Integer = 0
  Dim form_Disallow_Sweepstakes As Integer = 0
  Dim form_Disallow_ExecutionEngine As Integer = 0
  Dim form_Disallow_CRMEngine As Integer = 0
  Dim form_Disallow_DisplayDates As Integer = 0
  Dim form_Disallow_OfferRedempThreshold As Integer = 0
  Dim form_Disallow_UserDefinedFields As Integer = 0
  Dim form_Disallow_AdvancedOption As Integer = 0
  If (Request.QueryString("Disallow_ExecutionEngine") = "on") Then
    form_Disallow_ExecutionEngine = 1
  End If
  If (Request.QueryString("Disallow_CRMEngine") = "on") Then
    form_Disallow_CRMEngine = 1
  End If
  If (Request.QueryString("Disallow_EmployeeFiltering") = "on") Then
    form_Disallow_EmployeeFiltering = 1
  End If
  If (Request.QueryString("Disallow_ProductionDates") = "on") Then
    form_Disallow_ProductionDates = 1
  End If
  If (Request.QueryString("Disallow_Limits") = "on") Then
    form_Disallow_Limits = 1
  End If
  If (Request.QueryString("Disallow_Tiers") = "on") Then
    form_Disallow_Tiers = 1
  End If
  If (Request.QueryString("Disallow_Priority") = "on") Then
    form_Disallow_Priority = 1
  End If
  If (Request.QueryString("Disallow_Sweepstakes") = "on") Then
    form_Disallow_Sweepstakes = 1
  End If
  If (Request.QueryString("Disallow_UserDefinedFields") = "on") Then
    form_Disallow_UserDefinedFields = 1
  End If
  If bUseDisplayDates Then
    If (Request.QueryString("Disallow_DisplayDates") = "on") Then
      form_Disallow_DisplayDates = 1
    End If
  End If
                
  If bUseOfferRedemptionThreshold Then
    If (Request.QueryString("Disallow_OfferRedempThreshold") = "on") Then
      form_Disallow_OfferRedempThreshold = 1
    End If
  End If
               
  If (Request.QueryString("Disallow_AdvancedOption") = "on") Then
    form_Disallow_AdvancedOption = 1
  End If
  MyCommon.QueryStr = "update TemplatePermissions with (RowLock) set Disallow_EmployeeFiltering=" & form_Disallow_EmployeeFiltering & _
                      " , Disallow_ProductionDates=" & form_Disallow_ProductionDates & _
                      " , Disallow_limits=" & form_Disallow_Limits & _
                      " , Disallow_Tiers=" & form_Disallow_Tiers & _
                      " , Disallow_Priority=" & form_Disallow_Priority & _
                      " , Disallow_CRMEngine=" & form_Disallow_CRMEngine & _
                      " , Disallow_ExecutionEngine=" & form_Disallow_ExecutionEngine & _
                      " , Disallow_Sweepstakes=" & form_Disallow_Sweepstakes & _
 " , Disallow_UserDefinedFields=" & form_Disallow_UserDefinedFields & _
 " , Disallow_AdvancedOption=" & form_Disallow_AdvancedOption
							
  If bUseDisplayDates Then
    MyCommon.QueryStr = MyCommon.QueryStr & " , Disallow_DisplayDates=" & form_Disallow_DisplayDates
  End If
              
  If bUseOfferRedemptionThreshold Then
    MyCommon.QueryStr = MyCommon.QueryStr & " , Disallow_OfferRedempThreshold=" & form_Disallow_OfferRedempThreshold
  End If
               
  MyCommon.QueryStr = MyCommon.QueryStr & " where OfferID=" & OfferID
                
  MyCommon.LRT_Execute()
End If
' someone update an offer lets set flags on it
' add 1 to the udpatelevel and the crmupdatelevel
' MyCommon.QueryStr = "update Offers set UpdateLevel=UpdateLevel+1 where OfferID=" & OfferID
' MyCommon.LRT_Execute()
'MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-edit", LanguageID))
	  
ArrChangedVal = Request.QueryString("hdnChangedVal").Split("|")
	  
If ArrChangedVal.Length > 0 Then
  ComposedHist = Copient.PhraseLib.Lookup("history.offer-editgen", LanguageID)
  If Array.IndexOf(ArrChangedVal, "form_Name") >= 0 Then
    ComposedHist += Copient.PhraseLib.Lookup("term.name", LanguageID) & ", "
  End If
  If Array.IndexOf(ArrChangedVal, "form_Description") >= 0 Then
    ComposedHist += Copient.PhraseLib.Lookup("term.description", LanguageID) & ", "
  End If
  If Array.IndexOf(ArrChangedVal, "form_Category") >= 0 Then
    ComposedHist += Copient.PhraseLib.Lookup("term.category", LanguageID) & ", "
  End If
  If Array.IndexOf(ArrChangedVal, "prod-start-date") >= 0 Then
    ComposedHist += Copient.PhraseLib.Lookup("history.offerstartdate-edit", LanguageID) & ", "
  End If
  If Array.IndexOf(ArrChangedVal, "prod-end-date") >= 0 Then
    ComposedHist += Copient.PhraseLib.Lookup("history.offerenddate-edit", LanguageID) & ", "
  End If
  If Array.IndexOf(ArrChangedVal, "prod-start-hr") >= 0 OrElse Array.IndexOf(ArrChangedVal, "prod-start-min") >= 0 Then
    ComposedHist += Copient.PhraseLib.Lookup("history.offerstarttime-edit", LanguageID) & ", "
  End If
  If Array.IndexOf(ArrChangedVal, "prod-end-hr") >= 0 OrElse Array.IndexOf(ArrChangedVal, "prod-end-min") >= 0 Then
    ComposedHist += Copient.PhraseLib.Lookup("history.offerendtime-edit", LanguageID) & ", "
  End If
  If Array.IndexOf(ArrChangedVal, "form_DispStartDate") >= 0 Then
    ComposedHist += Copient.PhraseLib.Lookup("history.offerdisplaystartdate-edit", LanguageID) & ", "
  End If
  If Array.IndexOf(ArrChangedVal, "form_DispEndDate") >= 0 Then
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
  If Array.IndexOf(ArrChangedVal, "EmployeeFiltering") >= 0 OrElse Array.IndexOf(ArrChangedVal, "OnlyEmployees") >= 0 _
      OrElse Array.IndexOf(ArrChangedVal, "NonOnlyEmployeess") >= 0 Then
    ComposedHist += Copient.PhraseLib.Lookup("history.offeremployee-edit", LanguageID) & ", "
  End If
  If Array.IndexOf(ArrChangedVal, "OfferRedemptionThresholdperHour") >= 0 Then
    ComposedHist += Copient.PhraseLib.Lookup("history.offerredemption-edit", LanguageID) & ", "
  End If
  If Array.IndexOf(ArrChangedVal, "limitvalue") >= 0 OrElse Array.IndexOf(ArrChangedVal, "limitperiod") >= 0 _
 OrElse Array.IndexOf(ArrChangedVal, "selectadv") >= 0 Then
    ComposedHist += Copient.PhraseLib.Lookup("term.limits", LanguageID) & ", "
  End If
  If Array.IndexOf(ArrChangedVal, "tiertype0") >= 0 OrElse Array.IndexOf(ArrChangedVal, "tiertype1") >= 0 _
  OrElse Array.IndexOf(ArrChangedVal, "tiertype2") >= 0 OrElse Array.IndexOf(ArrChangedVal, "form_NumTiers") >= 0 Then
    ComposedHist += Copient.PhraseLib.Lookup("term.tiers", LanguageID) & ", "
  End If
  If Array.IndexOf(ArrChangedVal, "crmengine") >= 0 OrElse Array.IndexOf(ArrChangedVal, "InboundCRMEngineID") >= 0 Then
    ComposedHist += Copient.PhraseLib.Lookup("term.inbound/outbound", LanguageID) & ", "
  End If
  If Array.IndexOf(ArrChangedVal, "instantwin") >= 0 OrElse Array.IndexOf(ArrChangedVal, "prizes") >= 0 _
     OrElse Array.IndexOf(ArrChangedVal, "odds") >= 0 OrElse Array.IndexOf(ArrChangedVal, "random") >= 0 _
  OrElse Array.IndexOf(ArrChangedVal, "fixed") >= 0 OrElse Array.IndexOf(ArrChangedVal, "odds-calconce") >= 0 _
  OrElse Array.IndexOf(ArrChangedVal, "odds-calceach") >= 0 Then
    ComposedHist += Copient.PhraseLib.Lookup("term.instantwin", LanguageID) & ", "
  End If
  If Array.IndexOf(ArrChangedVal, "form_DisableDisconnectedOffer") >= 0 OrElse Array.IndexOf(ArrChangedVal, "form_autotransferable") >= 0 _
      OrElse Array.IndexOf(ArrChangedVal, "form_highvalue") >= 0 OrElse Array.IndexOf(ArrChangedVal, "form_exportEDW") >= 0 _
  OrElse Array.IndexOf(ArrChangedVal, "form_eCoupon") >= 0  OrElse Array.IndexOf(ArrChangedVal, "form_fob") >= 0 OrElse Array.IndexOf(ArrChangedVal, "form_autotranslate") >= 0 Then
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
End If
    
' Update distribution Limit info
If ValidDayLimit Then 'Don't update the dis variables unless the day limit is valid. This will always be true if the limit is not set to be days
MyCommon.QueryStr = "update Offers with (RowLock) set" & _
                    " AdvancedLimitID=" & AdvancedLimitID & _
                    ",DistPeriodVarID=" & VarID & _
                    ",DistPeriodLimit=" & form_DistPeriodLimit & _
                    ",DistPeriod=" & form_DistPeriod & _
                    " where OfferID=" & OfferID & ";"
MyCommon.LRT_Execute()
End If
  
' update the banner engine (if necessary)
If (MyCommon.Fetch_SystemOption(66) = "1" AndAlso Request.QueryString("bannerschanged") = "true") Then
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

    If bEnableBuckOffers Then
      If OfferID > 0 Then
        Dim oBuckStatus As Copient.ImportXml.BuckOfferStatus

        oBuckStatus = MyImport.BuckOfferGetStatus(OfferID)
        If oBuckStatus = Copient.ImportXml.BuckOfferStatus.BuckTiered Or oBuckStatus = Copient.ImportXml.BuckOfferStatus.BuckParentNoChildren Then
          MyCommon.QueryStr = "select PKID from OfferLocations with (NoLock) where Deleted=0 and OfferID=" & OfferID
          rst = MyCommon.LRT_Select
          If rst.Rows.Count = 0 Then
            Dim lAllStores As Long
            MyCommon.QueryStr = "select LocationGroupID from LocationGroups with (NoLock) where Deleted=0 and isnull(EngineID,0)=0 and AllLocations=1;"
            rst = MyCommon.LRT_Select
            If rst.Rows.Count > 0 Then
              ' assign "All Stores"
              lAllStores = rst.Rows(0).Item(0)
              MyCommon.QueryStr = "dbo.pt_OfferLocations_Insert"
              MyCommon.Open_LRTsp()
              MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
              MyCommon.LRTsp.Parameters.Add("@LocationGroupID", SqlDbType.BigInt).Value = lAllStores
              MyCommon.LRTsp.Parameters.Add("@Excluded", SqlDbType.Bit).Value = 0
              MyCommon.LRTsp.Parameters.Add("@TCRMAStatusFlag", SqlDbType.Int).Value = 3
              MyCommon.LRTsp.Parameters.Add("@PKID", SqlDbType.BigInt).Direction = ParameterDirection.Output
              MyCommon.LRTsp.ExecuteNonQuery()
              MyCommon.Close_LRTsp()
      
              MyCommon.QueryStr = "update OfferLocations with (RowLock) set StatusFlag=2,TCRMAStatusFlag=3 where OfferID=" & OfferID & " and LocationGroupID=" & lAllStores & ";"
              MyCommon.LRT_Execute()

              MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where OfferID=" & OfferID & ";"
              MyCommon.LRT_Execute()
              MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-addstore", LanguageID))
            End If
          End If

          MyCommon.QueryStr = "select TerminalTypeId from OfferTerminals with (NoLock) where OfferID=" & OfferID & ";"
          rst = MyCommon.LRT_Select
          If rst.Rows.Count = 0 Then
            Dim iTerminalTypeId As Integer
            MyCommon.QueryStr = "select TerminalTypeId from TerminalTypes with (NoLock) where EngineId=0 and AnyTerminal=1;"
            rst = MyCommon.LRT_Select
            If rst.Rows.Count > 0 Then
              ' assign "All CM Terminals"
              iTerminalTypeId = rst.Rows(0).Item(0)
              MyCommon.QueryStr = "insert into OfferTerminals with (RowLock) (OfferID, TerminalTypeID, LastUpdate, Excluded) values(" & OfferID & ", " & iTerminalTypeId & ", getdate(), 0)"
              MyCommon.LRT_Execute()
              MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-addterminal", LanguageID))
            
              MyCommon.QueryStr = "select TerminalTypeId from TerminalTypes with (NoLock) where EngineId=0 and ExtTerminalCode='15';"
              rst = MyCommon.LRT_Select
              If rst.Rows.Count > 0 Then
                ' exclude "15 - Gas Stations "
                iTerminalTypeId = rst.Rows(0).Item(0)
                MyCommon.QueryStr = "insert into OfferTerminals with (RowLock) (OfferID, TerminalTypeID, LastUpdate, Excluded) values(" & OfferID & ", " & iTerminalTypeId & ", getdate(), 1)"
                MyCommon.LRT_Execute()
                MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-excludeterminal", LanguageID))
              End If
            End If
            MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where OfferID=" & OfferID & ";"
            MyCommon.LRT_Execute()
          End If
        End If
      End If
    End If

    If (MyCommon.Fetch_CM_SystemOption(135) = "1") Then
      If (form_AutoTranslate = "on") Then
        MyCommon.QueryStr = "update Offers with (RowLock) set" & _
                            " AutoTranslateToUE=1" & _
                            " where OfferID=" & OfferID & ";"
      Else
        MyCommon.QueryStr = "update Offers with (RowLock) set" & _
                            " AutoTranslateToUE=0" & _
                            " where OfferID=" & OfferID & ";"
      End If
      MyCommon.LRT_Execute()
    End If

End If ' end save block
  
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
  
If (Request.QueryString("mode") = "addShared") Then
MyCommon.QueryStr = "update Offers with (RowLock) set SharedLimitID=" & Request.QueryString("ID") & " where OfferID=" & OfferID
MyCommon.LRT_Execute()
MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-addshared", LanguageID))
End If
If (Request.QueryString("mode") = "remShared") Then
MyCommon.QueryStr = "update Offers with (RowLock) set SharedLimitID=0 where OfferID=" & Request.QueryString("ID")
MyCommon.LRT_Execute()
MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-removeshared", LanguageID))
End If
  
' end handle form submissions
Dim eCoupon_Offer As Boolean = False
  
' dig the offer info out of the database
If (Request.QueryString("mode") <> "Create") Then
' no one clicked anything
    
MyCommon.QueryStr = "select OfferID,IsTemplate,FromTemplate,StatusFlag,ExtOfferID,ExtOfferID2,Name,Description,OfferCategoryID,OfferTypeID," & _
                    "ProdStartDate,ProdEndDate,TestStartDate,TestEndDate,TierTypeID,NumTiers,DistPeriod,DistPeriodLimit,DistPeriodVarID," & _
                    "EmployeeFiltering,NonEmployeesOnly,CRMRestricted,LastUpdate,PriorityLevel,EngineID,SharedLimitID,InstantWin," & _
                    "NumPrizesAllowed,CMOADeployStatus,OddsOfWinning,RandomWinners,IWTransLevel,CRMEngineID,DisabledOnCFW,DisplayOnWebKiosk," & _
                    "ExportToEDW,Favorite,AutoTransferable,HighValue,DisableDisconnectedOffer,AdvancedLimitID,OfferTypeID,InboundCRMEngineID,AutoTranslateToUE " & _
                    ", FobEligible " & _
                    "from Offers with (NoLock) where OfferID=" & OfferID & " and Deleted=0 and Visible=1;"

rst = MyCommon.LRT_Select()
For Each row In rst.Rows
If AllowSpecialCharacters <> "" Then
  Name = MyCommon.NZ(row.Item("Name"), "").Replace(Chr(34), "&quot;")
Else
  Name = MyCommon.NZ(row.Item("Name"), "")
End If
	  
'   CreatedDate = row.Item("CreatedDate")
If (row.Item("EmployeeFiltering") = True) Then EmployeeFiltering = True
If (row.Item("NonEmployeesOnly") = True) Then
  NonEmployeesOnly = True
Else
  NonEmployeesOnly = False
End If
      
IsTemplate = MyCommon.NZ(row.Item("IsTemplate"), False)
If IsTemplate Then
  bUseTemplateLocks = False
Else
  bUseTemplateLocks = MyCommon.NZ(row.Item("FromTemplate"), False)
End If
TierTypeID = MyCommon.NZ(row.Item("TierTypeID"), 0)
NumTiers = MyCommon.NZ(row.Item("NumTiers"), 0)
LastUpdate = row.Item("LastUpdate")
OfferCategoryID = MyCommon.NZ(row.Item("OfferCategoryID"), 0)
EngineID = MyCommon.NZ(row.Item("EngineID"), 0)
ExtOfferID2 = MyCommon.NZ(row.Item("ExtOfferID2"), "")
CRMEngineID = MyCommon.NZ(row.Item("CRMEngineID"), 0)
PriorityLevel = MyCommon.NZ(row.Item("PriorityLevel"), 50)
ExtOfferID = MyCommon.NZ(row.Item("ExtOfferID"), "")
      
tempDateTime = MyCommon.NZ(row.Item("TestStartdate"), Now())
TestStartdate = tempDateTime.ToString(sDateOnlyFormat)
TestStartHr = tempDateTime.ToString(sHourOnlyFormat)
TestStartMin = tempDateTime.ToString(sMinutesOnlyFormat)

tempDateTime = MyCommon.NZ(row.Item("TestEnddate"), Now())
TestEndDate = tempDateTime.ToString(sDateOnlyFormat)
TestEndHr = tempDateTime.ToString(sHourOnlyFormat)
TestEndMin = tempDateTime.ToString(sMinutesOnlyFormat)
If Integer.Parse(TestEndHr) = 0 And Integer.Parse(TestEndMin) = 0 Then
  TestEndHr = "23"
  TestEndMin = "59"
End If
      
tempDateTime = MyCommon.NZ(row.Item("ProdStartdate"), Now())
ProdStartDate = tempDateTime.ToString(MyCommon.GetAdminUser.Culture)
ProdStartHr = tempDateTime.ToString(sHourOnlyFormat)
ProdStartMin = tempDateTime.ToString(sMinutesOnlyFormat)
      
tempDateTime = MyCommon.NZ(row.Item("ProdEnddate"), Now())
ProdEnddate = tempDateTime.ToString(MyCommon.GetAdminUser.Culture)
ProdEndHr = tempDateTime.ToString(sHourOnlyFormat)
ProdEndMin = tempDateTime.ToString(sMinutesOnlyFormat)
If (ProdEndHr = 0 And ProdEndMin = 0) Then
  ProdEndHr = 23
  ProdEndMin = 59
End If

DistPeriod = MyCommon.NZ(row.Item("DistPeriod"), 0)
DistPeriodLimit = MyCommon.NZ(row.Item("DistPeriodLimit"), 0)
SharedLimitID = MyCommon.NZ(row.Item("SharedLimitID"), 0)
Description = MyCommon.NZ(row.Item("Description"), "")
If Description <> "" Then
  If AllowSpecialCharacters <> "" Then
    Description = Description.Replace(Chr(34), "&quot;")
  End If
End If
InstantWin = MyCommon.NZ(row.Item("InstantWin"), 0)
NumPrizesAllowed = MyCommon.NZ(row.Item("NumPrizesAllowed"), 0)
OddsOfWinning = MyCommon.NZ(row.Item("OddsOfWinning"), 0)
StatusFlag = MyCommon.NZ(row.Item("StatusFlag"), 0)
CMOADeployStatus = MyCommon.NZ(row.Item("CMOADeployStatus"), 0)
DisabledOnCFW = MyCommon.NZ(row.Item("DisabledOnCFW"), False)
DisplayOnWebKiosk = MyCommon.NZ(row.Item("DisplayOnWebKiosk"), False)
ExportToEDW = MyCommon.NZ(row.Item("ExportToEDW"), False)
Favorite = MyCommon.NZ(row.Item("Favorite"), False)
AutoTransferable = MyCommon.NZ(row.Item("AutoTransferable"), False)
HighValue = MyCommon.NZ(row.Item("HighValue"), False)
DisableDisconnectedOffer = MyCommon.NZ(row.Item("DisableDisconnectedOffer"), False)
AdvancedLimitID = MyCommon.NZ(row.Item("AdvancedLimitID"), 0)
eCoupon_Offer = IIf(MyCommon.NZ(row.Item("OfferTypeID"), 1) = gintEcouponOfferTypeID, True, False)
bFobEligible = MyCommon.NZ(row.Item("FobEligible"), False)      
bAutoTranslate = MyCommon.NZ(row.Item("AutoTranslateToUE"), False)
      
If (MyCommon.NZ(row.Item("RandomWinners"), False) = True) Then
  RandomWinners = True
Else
  RandomWinners = False
End If
If (MyCommon.NZ(row.Item("IWTransLevel"), 0) = True) Then
  IWTransLevel = True
Else
  IWTransLevel = False
End If
InboundCRMEngineID = MyCommon.NZ(row.Item("InboundCRMEngineID"), 0)
' Response.Write(RandomWinners & IWTransLevel)
If bUseDisplayDates Then
  MyCommon.QueryStr = "SELECT DisplayStartDate,DisplayEndDate FROM OfferAccessoryFields with (NoLock) WHERE OfferID = " & OfferID & ";"
  Dim dtODisp As New DataTable
  dtODisp = MyCommon.LRT_Select()
  If dtODisp.Rows.Count > 0 Then
    tempDateTime = MyCommon.NZ(dtODisp.Rows(0).Item("DisplayStartDate"), Nothing)
    If tempDateTime <> Nothing Then
      DispStartDate = tempDateTime.ToString(sDateOnlyFormat)
      DispStartHr = tempDateTime.ToString(sHourOnlyFormat)
      DispStartMin = tempDateTime.ToString(sMinutesOnlyFormat)
    End If
                        
    tempDateTime = MyCommon.NZ(dtODisp.Rows(0).Item("DisplayEndDate"), Nothing)
    If tempDateTime <> Nothing Then
      DispEnddate = tempDateTime.ToString(sDateOnlyFormat)
      DispEndHr = tempDateTime.ToString(sHourOnlyFormat)
      DispEndMin = tempDateTime.ToString(sMinutesOnlyFormat)
      If (DispEndHr = 0 And DispEndMin = 0) Then
        DispEndHr = 23
        DispEndMin = 59
      End If
    End If
                      
  End If
End If
            
If (bUseOfferRedemptionThreshold) Then
  MyCommon.QueryStr = "SELECT RedemThresholdPerHour FROM offerAccessoryFields with (NoLock) WHERE OfferID = " & OfferID & ";"
  Dim dtRedemption As New DataTable
  dtRedemption = MyCommon.LRT_Select()
  If dtRedemption.Rows.Count > 0 Then
    OfferRedemptionThresholdperHour = MyCommon.NZ(dtRedemption.Rows(0).Item("RedemThresholdPerHour"), 0)
  End If
End If
Next
End If 'end <> create block
  
' If (Not ExtOfferID) Then ExtOfferID = 0
  
If (Request.QueryString("mode") = "ChangeExtID") Then
'Check for an external id for this creation source already exists
MyCommon.QueryStr = "SELECT OfferID FROM Offers with (NoLock) WHERE Deleted=0 and ExtOfferID = '" & MyCommon.Extract_Val(Request.QueryString("ExtOffer")) & "' AND InboundCRMEngineID= " & InboundCRMEngineID
rst = MyCommon.LRT_Select
If rst.Rows.Count > 0 Then
infoMessage = Copient.PhraseLib.Lookup("offer-gen.extidinboundidexist", LanguageID)
FormValid = False
Else
ExtOfferID = MyCommon.Extract_Val(Request.QueryString("ExtOffer"))
If ExtOfferID = 0 Then
  MyCommon.QueryStr = "update Offers set ExtOfferID=NULL where OfferID=" & Request.QueryString("OfferID")
  MyCommon.LRT_Execute()
  ExtOfferID = String.Empty
  MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offerextid-remove", LanguageID))
ElseIf ExtOfferID > 0 Then
  MyCommon.QueryStr = "update Offers set ExtOfferID=" & ExtOfferID & " where OfferID=" & Request.QueryString("OfferID")
  MyCommon.LRT_Execute()
  MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offerextid-create", LanguageID))
End If
End If
End If
  
If (IsTemplate Or bUseTemplateLocks) Then
' lets dig the permissions if its a template
MyCommon.QueryStr = "select * from templatepermissions with (NoLock) where OfferID=" & OfferID
rst = MyCommon.LRT_Select
If (rst.Rows.Count > 0) Then
For Each row In rst.Rows
  ' ok there are some rows for the template
  Disallow_EmployeeFiltering = MyCommon.NZ(row.Item("Disallow_EmployeeFiltering"), True)
  Disallow_ProductionDates = MyCommon.NZ(row.Item("Disallow_ProductionDates"), True)
  Disallow_Limits = MyCommon.NZ(row.Item("Disallow_Limits"), True)
  Disallow_Tiers = MyCommon.NZ(row.Item("Disallow_Tiers"), True)
  Disallow_Priority = MyCommon.NZ(row.Item("Disallow_Priority"), True)
  Disallow_Sweepstakes = MyCommon.NZ(row.Item("Disallow_Sweepstakes"), True)
  Disallow_Conditions = MyCommon.NZ(row.Item("Disallow_Conditions"), True)
  Disallow_Rewards = MyCommon.NZ(row.Item("Disallow_Rewards"), True)
  Disallow_ExecutionEngine = MyCommon.NZ(row.Item("Disallow_ExecutionEngine"), True)
  Disallow_CRMEngine = MyCommon.NZ(row.Item("Disallow_CRMEngine"), True)
  Disallow_UserDefinedFields = MyCommon.NZ(row.Item("Disallow_UserDefinedFields"), True)
  If bUseDisplayDates Then
    Disallow_DisplayDates = MyCommon.NZ(row.Item("Disallow_DisplayDates"), True)
  End If
             
  If bUseOfferRedemptionThreshold Then
    Disallow_OfferRedempThreshold = MyCommon.NZ(row.Item("Disallow_OfferRedempThreshold"), True)
  End If
  Disallow_AdvancedOption = MyCommon.NZ(row.Item("Disallow_AdvancedOption"), True)
               
Next
End If
End If
  
'Check that the External OfferID can be changed
If InboundCRMEngineID = 1 Or InboundCRMEngineID = 2 Then
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
 
  
ShowInboundOutboundBox = (MyCommon.Fetch_SystemOption(25) <> "0")
StatusText = Logix.GetOfferStatus(OfferID, LanguageID, StatusCode)
  
Response.Expires = 0
'AdminUserID = Verify_AdminUser(MyCommon, )
Send_HeadBegin("term.offer", "term.general", OfferID)
Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
Send_Metas()
Send_Links(Handheld)
Send_Scripts(New String() {"datePicker.js"})
Send_HeadEnd()
If (IsTemplate) Then
Send_BodyBegin(11)
Else
Send_BodyBegin(1)
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
  .tempofferredeem
  {
    background-color: #ffdddd;
    border: 1px solid #ff0000;
    display: block;
    float: right;
    font-size: 12px;
    font-weight: normal;
    margin: 0 3px 3px 0;
    padding: 1px 3px 1px 2px;
    position: relative;
    right: -5px;
  }
  * html .tempofferredeem
  {
    padding: 0 3px 0 0;
    top: -18px;
  }
</style>
<script type="text/javascript">
    window.name="offerGen"
    var datePickerDivID = "datepicker";
    
    <% Send_Calendar_Overrides(MyCommon) %>
    
    function disableUnload() {
        window.onunload = null;
    }
        
    function elmName(){
        window.onunload = null;
        for(i=0; i<document.mainform.elements.length; i++)
        {
            document.mainform.elements[i].disabled=false;
            //alert(document.mainform.elements[i].name)
        }
        return true;
    }
    
    if (window.captureEvents){
        window.captureEvents(Event.CLICK);
        window.onclick=handlePageClick;
    }
    else {
        document.onclick=handlePageClick;
    }
    
    function handlePageClick(e) {
      var calFrame = document.getElementById('calendariframe');
      var el=(typeof event!=='undefined')? event.srcElement : e.target        
      
      if (el != null) {
        var pickerDiv = document.getElementById(datePickerDivID);
        if (pickerDiv != null && pickerDiv.style.visibility == "visible") {
          if (el.id!="prod-start-picker" && el.id!="prod-end-picker" && el.id!="test-start-picker" && el.id!="test-end-picker" && el.id!="disp-start-picker" && el.id!="disp-end-picker" && el.id!="udf-datevalue-picker") {
            if (!isDatePickerControl(el.className)) {
              pickerDiv.style.visibility = "hidden";
              pickerDiv.style.display = "none";  
              if (calFrame != null) {
                calFrame.style.visibility = 'hidden';
                calFrame.style.display = 'none';
              }
            }
          } else  {
              pickerDiv.style.visibility = "visible";            
              pickerDiv.style.display = "block";            
              if (calFrame != null) {
                calFrame.style.visibility = 'visible';
                calFrame.style.display = 'block';
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
	
  }

function ValidateDispDates() { 
  var elemDispStart = document.getElementById("form_DispStartDate"); 
  var elemDispEnd = document.getElementById("form_DispEndDate"); 
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


function promptForDeploy() {
  var elem = document.getElementById("IsActive");
  var retVal = true;
  var elemEnd = document.getElementById("prod-end-date");
  var dtNow = new Date();
  var dtEnd = new Date();

  if (elem != null && elem.value == "true" && elemEnd != null) {
     retVal = IsValidLocalizedDate(elemEnd.value, '<% Sendb(MyCommon.GetAdminUser.Culture.DateTimeFormat.ShortDatePattern) %>');
    if (retVal) {
      dtEnd = new Date(Date.parse(elemEnd.value));
      dtEnd.setDate(dtEnd.getDate() + 1);
      if (dtEnd < dtNow) {       
        retVal = confirm('<%Sendb(Copient.PhraseLib.Lookup("term.expire-confirm", LanguageID)) %>');
      }
    }
  }
  return retVal;
}

function handleOnSubmit() {
  var retVal = false;
   retVal = checkdesclengthdata() && ValidateOfferForm('<% Sendb(MyCommon.GetAdminUser.Culture.DateTimeFormat.ShortDatePattern) %>') && ValidateDispDates() && ValidateTimes() && promptForDeploy();
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


// This is the javascript array holding the function list
// The PrintJavascriptArray ASP function can be used to print this array.
<%
    MyCommon.QueryStr = "select 0 as LimitID, '" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "' as Name, 1 as LimitTypeID, DistPeriodLimit as LimitValue, DistPeriod as LimitPeriod " & _
                        "from Offers with (NoLock) where OfferID=" & OfferID & " " & _
                        "union " & _
                        "select LimitID, Name, LimitTypeID, LimitValue, LimitPeriod " & _
                        "from CM_AdvancedLimits with (NoLock) where Deleted=0 and LimitTypeID=1 order By LimitID;"
    rst2 = MyCommon.LRT_Select
    If rst2.Rows.Count > 0 Then
        Sendb("var ALfunctionlist = Array(")
        For Each row2 In rst2.Rows
            Sendb("""" & MyCommon.NZ(row2.item("Name"), "").ToString().Replace("""", "\""") & """,")
        Next
        Send(""""");")
        
        Sendb("var ALvallist1 = Array(")
        For Each row2 In rst2.Rows
            Sendb("""" & row2.item("LimitID") & """,")
        Next
        Send(""""");")
        
        Sendb("var ALvallist2 = Array(")
        For Each row2 In rst2.Rows
            Sendb("""" & row2.item("LimitPeriod") & """,")
        Next
        Send(""""");")
        
        Sendb("var ALvallist3 = Array(")
        For Each row2 In rst2.Rows
            Sendb("""" & row2.item("LimitValue") & """,")
        Next
        Send(""""");")
    End If
%>

function setlimitsection(bSelect) {
  var elemSelectAdv = document.getElementById("selectadv");
  var elemSelectDay=document.getElementById("selectday");
  var elemPeriod=document.getElementById("limitperiod");
  var elemValue=document.getElementById("limitvalue");
  var elemDisabled=document.getElementById("LimitsDisabled");
  
  if ((bSelect == true) || (elemSelectAdv != null)) {
    if ((elemDisabled == null) || (elemDisabled != null && elemDisabled.value == 'False')) {
      if (elemSelectAdv != null && elemSelectAdv.value == '0') {
        if (elemSelectDay != null) {
          elemSelectDay.disabled = false;
        }
        if (elemPeriod != null) {
          elemPeriod.disabled = false;
        }
        if (elemValue != null) {
          elemValue.disabled = false;
        }
      }
      else {
        if (elemSelectDay != null) {
          elemSelectDay.disabled = true;
        }
        if (elemPeriod != null) {
          elemPeriod.disabled = true;
        }
        if (elemValue != null) {
          elemValue.disabled = true;
        }
      }
    }
 
    for(i = 0; i < ALfunctionlist.length; i++)
    {
      if(elemSelectAdv.value == ALvallist1[i])
      {
        elemPeriod.value = ALvallist2[i];
        elemValue.value = ALvallist3[i];
        if (elemPeriod.value == -1) {
          elemSelectDay.value = '3';
          elemPeriod.style.visibility = 'hidden';
        }
        else if (elemPeriod.value == 0) {
          elemSelectDay.value = '2';
          elemPeriod.style.visibility = 'hidden';
        }
        else
        {
          elemSelectDay.value = '1';
          elemPeriod.style.visibility = 'visible';
        }
        break;
      }
    }
  }
  if (!bSelect) {
    elemPeriod.defaultValue = elemPeriod.value;
    elemValue.defaultValue = elemValue.value;
    elemSelectDay.defaultValue = elemSelectDay.value;
  }
}

function setperiodsection(bSelect) {
  var elemSelectDay = document.getElementById("selectday");
  var elemPeriod=document.getElementById("limitperiod");
  var elemOriginalPeriod=document.getElementById("OriginalPeriod");
  var elemImpliedPeriod=document.getElementById("ImpliedPeriod");

  if (elemSelectDay != null && (elemSelectDay.value == '2') || (elemSelectDay.value == '3')) {
    if (elemPeriod != null) {
      elemPeriod.style.visibility = 'hidden';
    }
    if (elemSelectDay.value == '2') {
      elemImpliedPeriod.value = '0';
      elemPeriod.value = '0';
    }
    else {
      elemImpliedPeriod.value = '-1';
      elemPeriod.value = '-1';
    }
  }
  else {
    if (elemPeriod != null) {
      if (bSelect && elemOriginalPeriod != null) {
        if ((elemOriginalPeriod.value == '-1') || (elemOriginalPeriod.value == '0')) {
          elemPeriod.value = '0';
        }
        else {
          elemPeriod.value = elemOriginalPeriod.value;
          elemImpliedPeriod.value = elemOriginalPeriod.value;
        }
      }
      elemPeriod.style.visibility = 'visible';
    }
  }
  if (!bSelect) {
    elemPeriod.defaultValue = elemPeriod.value;
    elemImpliedPeriod.defaultValue = elemImpliedPeriod.value;
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
	
function checkdesclength(){
var elemvalue = document.getElementById("form_Description").value;
if (elemvalue.length <= 1000) {
xmlhttpPost_OfferDescription('OfferFeeds.aspx', 'AllowSpecialCharactersCM');
}
}

function getQueryStringOfferDesc(mode) {
  var elemvalue = document.getElementById("form_Description").value;
  elemvalue = controllengthstring(elemvalue);
  //alert(elemvalue);
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
var elemvalue = document.getElementById("form_Description").value;
if (elemvalue.length > 1000) {
alert('<%Sendb(Copient.PhraseLib.Lookup("error.description", LanguageID)) %>');
return false;
}else
{
document.getElementById("form_Description").value = "";
return true;
}
}	

</script>
<udf:udfjavascript ID="udfjavascriptcontrol" runat="server" />
<%
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  Send_Tabs(Logix, 2)
  If (IsTemplate) Then
    Send_Subtabs(Logix, 22, 4, , OfferID)
  Else
    Send_Subtabs(Logix, 21, 4, , OfferID)
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
<form id="mainform" name="mainform" action="offer-gen.aspx" method="get" onsubmit="elmName(); return handleOnSubmit();">
<input type="hidden" id="form_OfferID" name="form_OfferID" value="<%sendb(OfferID) %>" />
<input type="hidden" id="OfferID" name="OfferID" value="<%sendb(OfferID) %>" />
<input type="hidden" name="IsActive" id="IsActive" value="<%Sendb(IIf(StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_ACTIVE, "true", "false")) %>" />
<input type="hidden" id="OriginalPeriod" name="OriginalPeriod" value="<% sendb(DistPeriod) %>" />
<input type="hidden" id="ImpliedPeriod" name="ImpliedPeriod" value="<% sendb(DistPeriod) %>" />
<input type="hidden" id="LimitsDisabled" name="LimitsDisabled" value="<% sendb(bUseTemplateLocks and Disallow_Limits) %>" />
<input type="hidden" name="SelectedUDF" id="SelectedUDF" value="" />
<input type="hidden" id="IsTemplate" name="IsTemplate" value="<% 
    if(istemplate)then 
    sendb("IsTemplate")
    else 
    sendb("Not") 
    end if
     %>" />
<input type="hidden" id="hdnExtOfferId" name="hdnExtOfferId" value="<% sendb(ExtOfferID) %>" />
<input type="hidden" id="hdnChangedVal" name="hdnChangedVal" />
<div id="intro">
  <%
    If (IsTemplate) Then
      If (OfferID = 0) Then
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & ": " & Copient.PhraseLib.Lookup("term.newtemplate", LanguageID) & "</h1>")
      Else
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & ": " & MyCommon.TruncateString(Name, 50) & "</h1>")
      End If
    Else
      If (OfferID = 0) Then
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & ": " & Copient.PhraseLib.Lookup("term.newoffer", LanguageID) & "</h1>")
      Else
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & ": " & MyCommon.TruncateString(Name, 50) & "</h1>")
      End If
    End If
  %>
  <div id="controls">
    <%
      If Not (IsTemplate) Then
        If (Logix.UserRoles.EditOffer) Then
          Send_Save("onclick=""updatehistory()""")
        End If
      Else
        If (Logix.UserRoles.EditTemplates) Then
          Send_Save("onclick=""updatehistory()""")
        End If
      End If
      If MyCommon.Fetch_SystemOption(75) Then
        If (OfferID > 0 And Logix.UserRoles.AccessNotes) Then
          Send_NotesButton(3, OfferID, AdminUserID)
        End If
      End If
    %>
  </div>
</div>
<div id="main">
  <%
    If Not IsTemplate Then
      If (StatusFlag <> 2) Then
        If (StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_ACTIVE) AndAlso (StatusFlag > 0) Then
          modMessage = Copient.PhraseLib.Lookup("alert.modpostdeploy", LanguageID)
          Send("<div id=""modbar"">" & modMessage & "</div>")
        End If
      End If
    End If
    If MyCommon.Fetch_SystemOption(191) AndAlso NumberofFolders > 1 Then
      infoMessage = infoMessage & " " & "Offer cannot have more than one Folder associated"
    End If
    If (IsTemplate) Then
      Send(" <div id=""infobar"" class=""red-background"">" & Copient.PhraseLib.Lookup("temp.note", LanguageID) & "</div>")
    End If
    If (infoMessage <> "") Then
      Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
    End If
      
    ' Send the status bar, but not if it's a new offer or a template, or if there's already a modMessage being shown.
    If (Not IsTemplate AndAlso modMessage = "") Then
      MyCommon.QueryStr = "select OfferID from Offers with (NoLock) where CreatedDate = LastUpdate and OfferID=" & OfferID
      rst3 = MyCommon.LRT_Select
      If (rst3.Rows.Count = 0) Then
        Send_Status(OfferID)
      End If
    End If
  %>
  <div id="column1">
    <div class="box" id="identification" style="z-index: 999">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.identification", LanguageID))%>
        </span>
      </h2>
      <% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID) & ": " & OfferID)%>
      <br />
      <%
        If (Not IsTemplate) Then
          'Allow for the change of the external id to the offer id
          Sendb(Copient.PhraseLib.Lookup("term.externalid", LanguageID) & ": ")
          If ChangeExtID Then
            If ExtOfferID = "" Then
              Send(Copient.PhraseLib.Lookup("term.none", LanguageID) & " <small><a href=""offer-gen.aspx?OfferID=" & OfferID & "&amp;ExtOffer=" & OfferID & "&amp;mode=ChangeExtID"">(" & Copient.PhraseLib.Lookup("CPEoffer-gen.xid-add", LanguageID) & ")</a></small>")
            Else
              Send(ExtOfferID & " <small><a href=""offer-gen.aspx?OfferID=" & OfferID & "&amp;ExtOffer=0&amp;mode=ChangeExtID"">(" & Copient.PhraseLib.Lookup("CPEoffer-gen.xid-rem", LanguageID) & ")</a></small>")
            End If
          Else
                                   
            MyCommon.QueryStr = "select DefaultAsLogixID from ExtCRMInterfaces where ExtInterfaceID = " & InboundCRMEngineID
            rstTemp = MyCommon.LRT_Select
            DefaultAsLogixID = MyCommon.NZ(rstTemp.Rows(0).Item("DefaultAsLogixID"), False)
            If (DefaultAsLogixID = True) Then
              ExtOfferID = OfferID
              MyCommon.QueryStr = "update Offers set ExtOfferID=" & ExtOfferID & " where OfferID=" & Request.QueryString("OfferID")
              MyCommon.LRT_Execute()
              Send(ExtOfferID)
            Else
              Send(ExtOfferID)
            End If
          End If
          Send("<br />")
        Else
          Sendb(Copient.PhraseLib.Lookup("term.externalid", LanguageID) & ": " & ExtOfferID)
          Send("<br />")
        End If
      %>
      <% Sendb(Copient.PhraseLib.Lookup("term.engine", LanguageID) & ": " & Copient.PhraseLib.Lookup("term.cm", LanguageID))%>
      <br />
      <% Sendb(Copient.PhraseLib.Lookup("term.status", LanguageID) & ": " & StatusText)%>
      <br />
      <br class="half" />
      <label for="form_Name">
        <% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>:</label><br />
      <%  If bUseAdvertisementText Then%>
      <input class="longest" id="form_Name" name="form_Name" maxlength="156" type="text"
        value="<% Sendb(Name) %>" />
      <%Else%>
      <input class="longest" id="form_Name" name="form_Name" maxlength="100" type="text"
        value="<% Sendb(Name) %>" />
      <% End If%>
      <br />
      <label for="form_Description">
        <% Sendb(Copient.PhraseLib.Lookup("term.description", LanguageID))%>:</label><br />
      <textarea class="longest" cols="48" rows="3" id="form_Description" name="form_Description"
        maxlength="1000" onchange="javascript:checkdesclength();"><% Sendb(Description)%></textarea><br />
      <br class="half" />
      <small>
        <%Send("(" & Copient.PhraseLib.Lookup("CPEoffergen.description", LanguageID) & ")")%></small><br
          class="half" />
      <br class="half" />
      <%' get the category list from database
        MyCommon.QueryStr = "select OfferCategoryID, Description from OfferCategories with (NoLock) where Deleted=0 order by Description"
        rst2 = MyCommon.LRT_Select()
      %>
      <label for="form_Category">
        <% Sendb(Copient.PhraseLib.Lookup("term.category", LanguageID))%>:</label><br />
      <select class="medium" id="form_Category" name="form_Category">
        <%
          For Each row2 In rst2.Rows
            If (OfferCategoryID = row2.Item("OfferCategoryID")) Then
              Sendb("<option value=""" & row2.Item("OfferCategoryID") & """ selected=""selected"">" & row2.Item("Description") & "</option>")
            Else
              Sendb("<option value=""" & row2.Item("OfferCategoryID") & """>" & row2.Item("Description") & "</option>")
            End If
          Next
        %>
      </select>
      <br />
      <br class="half" />
      <hr class="hidden" />
    </div>
    <br />
    <div class="box" id="period">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.period", LanguageID))%>
        </span>
      </h2>
      <% If (IsTemplate) Then%>
      <span class="temp">
        <input type="checkbox" class="tempcheck" id="temp-production" name="Disallow_ProductionDates"
          <% if(disallow_productiondates)then send(" checked=""checked""") %> />
        <label for="temp-production">
          <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
      </span>
      <br class="half" clear="all" />
      <% End If%>
      <span>
        <%
          If bAllowTimeWithStartEndDates Then
            Sendb(Copient.PhraseLib.Lookup("term.enter-datetime", LanguageID))
          Else
            Sendb(Copient.PhraseLib.Lookup("term.enter-date", LanguageID))
          End If
        %>
      </span>
      <br class="half" />
      <br class="half" />
      <% If bUseTestDates Then%>
      <label for="test-start-date">
        <% Sendb(Copient.PhraseLib.Lookup("term.test", LanguageID))%>:</label><br />
      <input class="short" id="test-start-date" name="form_TestStartDate" maxlength="10"
        type="text" value="<% sendb(Logix.ToShortDateString(TestStartDate,MyCommon)) %>"
        <% if(bUseTemplateLocks and disallow_productiondates)then sendb(" disabled=""disabled""") %> />
      <img src="../images/calendar.png" class="calendar" id="test-start-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
        title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('form_TestStartDate', event);" />
      <%If bAllowTimeWithStartEndDates Then%>
      <input class="shortest" id="test-start-hr" maxlength="2" name="form_TestStartHr"
        type="text" value="<% sendb(TestStartHr)%>" <% if(bUseTemplateLocks and disallow_productiondates)then sendb(" disabled=""disabled""") %> />:<input
          class="shortest" id="test-start-min" maxlength="2" name="form_TestStartMin" type="text"
          value="<% sendb(TestStartMin)%>" <% if(bUseTemplateLocks and disallow_productiondates)then sendb(" disabled=""disabled""") %> />
      <% End If%>
      <% Sendb(Copient.PhraseLib.Lookup("term.to", LanguageID))%>
      <input class="short" id="test-end-date" name="form_TestEndDate" maxlength="10" type="text"
        value="<% sendb(Logix.ToShortDateString(TestEndDate,MyCommon)) %>" <% if(bUseTemplateLocks and disallow_productiondates)then sendb(" disabled=""disabled""") %> />
      <img src="../images/calendar.png" class="calendar" id="test-end-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
        title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('form_TestEndDate', event);" />
      <%If bAllowTimeWithStartEndDates Then%>
      <input class="shortest" id="test-end-hr" maxlength="2" name="form_TestEndHr" type="text"
        value="<% sendb(TestEndHr)%>" <% if(bUseTemplateLocks and disallow_productiondates)then sendb(" disabled=""disabled""") %> />:<input
          class="shortest" id="test-end-min" maxlength="2" name="form_TestEndMin" type="text"
          value="<% sendb(TestEndMin)%>" <% if(bUseTemplateLocks and disallow_productiondates)then sendb(" disabled=""disabled""") %> />
      <% End If%>
	  <br />
      <br class="half" />
      <% End If%>
	  <label for="prod-start-date">
        <% Sendb(Copient.PhraseLib.Lookup("term.production", LanguageID))%>:</label><br />
      <input class="short" id="prod-start-date" name="form_ProdStartDate" maxlength="10"
        type="text" value="<% sendb(Logix.ToShortDateString(ProdStartDate,MyCommon)) %>"
        <% if(bUseTemplateLocks and disallow_productiondates)then sendb(" disabled=""disabled""") %> />
      <img src="../images/calendar.png" class="calendar" id="prod-start-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
        title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('form_ProdStartDate', event);" />
      <%If bAllowTimeWithStartEndDates Then%>
      <input class="shortest" id="prod-start-hr" maxlength="2" name="form_ProdStartHr"
        type="text" value="<% sendb(ProdStartHr)%>" <% if(bUseTemplateLocks and disallow_productiondates)then sendb(" disabled=""disabled""") %> />:<input
          class="shortest" id="prod-start-min" maxlength="2" name="form_ProdStartMin" type="text"
          value="<% sendb(ProdStartMin)%>" <% if(bUseTemplateLocks and disallow_productiondates)then sendb(" disabled=""disabled""") %> />
      <% End If%>
      <% Sendb(Copient.PhraseLib.Lookup("term.to", LanguageID))%>
      <input class="short" id="prod-end-date" name="form_ProdEndDate" maxlength="10" type="text"
        value="<% sendb(Logix.ToShortDateString(ProdEndDate,MyCommon)) %>" <% if(bUseTemplateLocks and disallow_productiondates)then sendb(" disabled=""disabled""") %> />
      <img src="../images/calendar.png" class="calendar" id="prod-end-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
        title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('form_ProdEndDate', event);" />
      <%If bAllowTimeWithStartEndDates Then%>
      <input class="shortest" id="prod-end-hr" maxlength="2" name="form_ProdEndHr" type="text"
        value="<% sendb(ProdEndHr)%>" <% if(bUseTemplateLocks and disallow_productiondates)then sendb(" disabled=""disabled""") %> />:<input
          class="shortest" id="prod-end-min" maxlength="2" name="form_ProdEndMin" type="text"
          value="<% sendb(ProdEndMin)%>" <% if(bUseTemplateLocks and disallow_productiondates)then sendb(" disabled=""disabled""") %> />
      <% End If%>
      <br />
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
          <% if(Disallow_DisplayDates)then send(" checked=""checked""") %> />
        <label for="temp-priority">
          <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
      </span>
      <br class="printonly" />
      <% End If%>
      <span>
        <%
          Sendb(Copient.PhraseLib.Lookup("term.enter-datetime", LanguageID))
        %>
      </span>
      <br class="half" />
      <br class="half" />
      <label for="disp-start-date">
        <% Sendb(Copient.PhraseLib.Lookup("term.display", LanguageID))%>:</label><br />
      <input class="short" id="form_DispStartDate" name="form_DispStartDate" maxlength="10"
        type="text" value="<% If Not String.IsNullOrEmpty(DispStartDate) Then sendb(Logix.ToShortDateString(DispStartDate, MyCommon)) %>"
        <% if(bUseTemplateLocks and Disallow_DisplayDates)then sendb(" disabled=""disabled""") %> />
      <img src="../images/calendar.png" class="calendar" id="disp-start-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
        title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('form_DispStartDate', event);" />
      <input class="shortest" id="disp-start-hr" maxlength="2" name="form_DispStartHr"
        type="text" value="<% sendb(DispStartHr)%>" <% if(bUseTemplateLocks and Disallow_DisplayDates)then sendb(" disabled=""disabled""") %> />:<input
          class="shortest" id="disp-start-min" maxlength="2" name="form_DispStartMin" type="text"
          value="<% sendb(DispStartMin)%>" <% if(bUseTemplateLocks and Disallow_DisplayDates)then sendb(" disabled=""disabled""") %> />
      <% Sendb(Copient.PhraseLib.Lookup("term.to", LanguageID))%>
      <input class="short" id="form_DispEndDate" name="form_DispEndDate" maxlength="10"
        type="text" value="<%If Not String.IsNullOrEmpty(DispEndDate) Then sendb(Logix.ToShortDateString(DispEndDate,MyCommon))%>"
        <% if(bUseTemplateLocks and Disallow_DisplayDates)then sendb(" disabled=""disabled""") %> />
      <img src="../images/calendar.png" class="calendar" id="disp-end-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
        title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('form_DispEndDate', event);" />
      <input class="shortest" id="disp-end-hr" maxlength="2" name="form_DispEndHr" type="text"
        value="<% sendb(DispEndHr)%>" <% if(bUseTemplateLocks and Disallow_DisplayDates)then sendb(" disabled=""disabled""") %> />:<input
          class="shortest" id="disp-end-min" maxlength="2" name="form_DispEndMin" type="text"
          value="<% sendb(DispEndMin)%>" <% if(bUseTemplateLocks and Disallow_DisplayDates)then sendb(" disabled=""disabled""") %> />
      <br />
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
        <label for="temp-priority">
          <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
      </span>
      <br class="printonly" />
      <% End If%>
      <label for="priority">
        <% Sendb(Copient.PhraseLib.Lookup("offer-gen.priority", LanguageID))%></label>
      <br />
      <%If (MyCommon.Fetch_CM_SystemOption(117) = "1") Then
          If (Logix.UserRoles.EditOfferPriority) Then
            permission_disallow_priority = False
          Else
            permission_disallow_priority = True
          End If
        End If%>
      <input class="shortest" id="priority" maxlength="2" name="form_Priority" type="text"
        value="<% sendb(PriorityLevel)%>" <% if((bUseTemplateLocks and disallow_priority) or permission_disallow_priority)then sendb(" disabled=""disabled""") %> /><br />
      <hr class="hidden" />
    </div>
    <div class="box" id="employees">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.employees", LanguageID))%>
        </span>
      </h2>
      <%If (IsTemplate) Then%>
      <span class="temp">
        <input type="checkbox" class="tempcheck" id="temp-employees" name="Disallow_EmployeeFiltering"
          <% if(disallow_employeefiltering)then sendb(" checked=""checked""") %> />
        <label for="temp-employees">
          <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
      </span>
      <br class="printonly" />
      <% End If%>
      <input class="checkbox" id="EmployeeFiltering" name="EmployeeFiltering" type="checkbox"
        onclick="javascript:if(this.checked){document.mainform.eligradio[0].checked=true;}else{document.mainform.eligradio[0].checked=false;document.mainform.eligradio[1].checked=false;}"
        <% if(bUseTemplateLocks and disallow_employeefiltering)then sendb(" disabled=""disabled""") %><% if(employeefiltering)then sendb(" checked=""checked""") %> />
      <label for="EmployeeFiltering">
        <% Sendb(Copient.PhraseLib.Lookup("offer-gen.empfilter", LanguageID))%></label>
      <br />
      &nbsp;
      <input class="radio" id="OnlyEmployees" name="eligradio" value="0" onclick="document.getElementById('EmployeeFiltering').checked=true;"
        type="radio" <% if(bUseTemplateLocks and disallow_employeefiltering)then sendb(" disabled=""disabled""") %><% if(employeefiltering and not nonemployeesonly)then sendb(" checked=""checked""") %> />
      <label for="OnlyEmployees">
        <% Sendb(Copient.PhraseLib.Lookup("offer-gen.emp", LanguageID))%></label>
      <br />
      &nbsp;
      <input class="radio" id="NonOnlyEmployees" name="eligradio" value="1" onclick="document.getElementById('EmployeeFiltering').checked=true;"
        type="radio" <% if(bUseTemplateLocks and disallow_employeefiltering)then sendb(" disabled=""disabled""") %><% if(employeefiltering and nonemployeesonly)then sendb(" checked=""checked""") %> />
      <label for="NonOnlyEmployees">
        <% Sendb(Copient.PhraseLib.Lookup("offer-gen.nonemp", LanguageID))%></label>
      <br />
      <hr class="hidden" />
    </div>
    <%If MyCommon.Fetch_SystemOption(156) = "1" Then
        udflistcontrol.IsTemplate = IsTemplate
        udflistcontrol.bUseTemplateLocks = bUseTemplateLocks
        udflistcontrol.Disallow_UserDefinedFields = Disallow_UserDefinedFields
    %>
    <udf:udflist ID="udflistcontrol" runat="server" />
    <% End If%>
    <%
      If (bUseOfferRedemptionThreshold) Then
    %>
    <div class="box" id="OfferRedemptionThreshold">
      <h2>
        <span class="redemption">
          <% Sendb(Copient.PhraseLib.Lookup("term.offerredemptionthreshold", LanguageID))%>
        </span>
      </h2>
      <%-- Br28Starts--%>
      <% If (IsTemplate) Then%>
      <span class="tempofferredeem">
        <input type="checkbox" class="tempcheck" id="Disallow_OfferRedempThreshold" name="Disallow_OfferRedempThreshold"
          <% if(Disallow_OfferRedempThreshold)then send(" checked=""checked""") %> />
        <label for="temp-priority">
          <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
      </span>
      <br class="printonly" />
      <% End If%>
      <%-- Br28Ends--%>
      </br></br>
      <label for="OfferRedemptionThresholdperHour">
        <% Sendb(Copient.PhraseLib.Lookup("term.offerredemptionthresholdperhour", LanguageID))%>
        :</label>
      <br />
      <input class="short" id="OfferRedemptionThresholdperHour" name="OfferRedemptionThresholdperHour"
        maxlength="5" type="text" value="<% sendb(OfferRedemptionThresholdperHour)%>" <% If(bUseTemplateLocks and Disallow_OfferRedempThreshold) Then sendb(" disabled=""disabled""") %> />
    </div>
    <% End If%>
    <% If (MyCommon.Fetch_SystemOption(66) = "1") Then%>
    <div class="box" id="banners">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.banners", LanguageID))%>
        </span>
      </h2>
      <% 
        Dim SelectedList As String = ""

            
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
                            "where BE.EngineID=0 and AUB.AdminUserID =" & AdminUserID & ";"
        rst2 = MyCommon.LRT_Select
        EditableBanners = New ArrayList(rst2.Rows.Count)
        For Each row2 In rst2.Rows
          EditableBanners.Add(MyCommon.NZ(row2.Item("BannerID"), -1))
        Next
           
        ' get all the assigned banners for CPE
        i = 0
        MyCommon.QueryStr = "select BAN.BannerID, BAN.Name from Banners BAN with (NoLock) " & _
                            "inner join BannerEngines BE with (NoLock) on BE.BannerID = BAN.BannerID " & _
                            "where BE.EngineID=0 and BAN.AllBanners=0;"
        rst2 = MyCommon.LRT_Select()
        For Each row2 In rst2.Rows
          IsEditableBanner = EditableBanners.Contains(MyCommon.NZ(row2.Item("BannerID"), -1))
          Sendb(Space(10) & "<input type=""checkbox"" name=""bannerids"" id=""bannerid" & i & """ value=""" & MyCommon.NZ(row2.Item("BannerID"), -1) & """")
          Sendb(IIf(SelectedBanners.Contains(MyCommon.NZ(row2.Item("BannerID"), -1)), " checked=""checked""", " "))
          Sendb(IIf(IsEditableBanner, " ", " disabled=""disabled"""))
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
                            "where BE.EngineID=0 and BAN.AllBanners=1;"
        rst2 = MyCommon.LRT_Select()
        If (rst2.Rows.Count > 0) Then
          Send("<br />")
          Send(Copient.PhraseLib.Lookup("term.allbanners", LanguageID) & ":<br />")
          For Each row2 In rst2.Rows
            IsEditableBanner = EditableBanners.Contains(MyCommon.NZ(row2.Item("BannerID"), -1))
            Sendb(Space(10) & "<input type=""checkbox"" name=""bannerids"" id=""allbannerid" & i & """ value=""" & MyCommon.NZ(row2.Item("BannerID"), -1) & """")
            Sendb(IIf(SelectedBanners.Contains(MyCommon.NZ(row2.Item("BannerID"), -1)), " checked=""checked""", " "))
            Sendb(IIf(IsEditableBanner, " ", " disabled=""disabled"""))
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
        <label for="temp-limits">
          <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
      </span>
      <br class="printonly" />
      <% End If%>
      <%
        MyCommon.QueryStr = "Select LimitId, Name from CM_AdvancedLimits with (NoLock) where Deleted=0 and LimitTypeID=1 order By Name;"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
      %>
      <label for="selectadv">
        <% Sendb(Copient.PhraseLib.Lookup("term.advlimits", LanguageID))%>:</label>
      <select id="selectadv" name="selectadv" class="longer" onchange="setlimitsection(true);"
        <% If(bUseTemplateLocks and Disallow_Limits) Then sendb(" disabled=""disabled""") %>>
        <%
          Sendb("<option value=""0""")
          If (AdvancedLimitID = 0) Then
            Sendb(" selected=""selected""")
          End If
          Sendb(">" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "</option>")
          For Each row In rst.Rows
            Sendb("<option value=""" & row.Item("LimitID") & """")
            If (AdvancedLimitID = MyCommon.Extract_Val(MyCommon.NZ(row.Item("LimitID"), 0))) Then
              Sendb(" selected=""selected""")
            End If
            Sendb(">")
            Sendb(row.Item("Name"))
            Sendb("</option>")
          Next
        %>
      </select>
      <br class="half" />
      <% End If%>
      <br class="half" />
      <input class="shorter" id="limitvalue" name="limitvalue" maxlength="9" type="text"
        value="<% sendb(DistPeriodLimit) %>" <% If(bUseTemplateLocks and Disallow_Limits) Then sendb(" disabled=""disabled""") %> />
      &nbsp;<% Sendb(Copient.PhraseLib.Lookup("term.per", LanguageID))%>
      <input class="shorter" id="limitperiod" name="limitperiod" maxlength="4" type="text"
        value="<% sendb(DistPeriod) %>" <% If(bUseTemplateLocks and Disallow_Limits) Then sendb(" disabled=""disabled""") %> />
      <select id="selectday" name="selectday" onchange="setperiodsection(true);" <% If(bUseTemplateLocks and Disallow_Limits) Then sendb(" disabled=""disabled""") %>>
        <option value="1" <% if(distperiod>0)then sendb(" selected=""selected""") %>>
          <% Sendb(Copient.PhraseLib.Lookup("term.days", LanguageID))%>
        </option>
        <option value="2" <% if(distperiod=0)then sendb(" selected=""selected""") %>>
          <% Sendb(Copient.PhraseLib.Lookup("term.transaction", LanguageID))%>
        </option>
        <option value="3" <% if(distperiod=-1)then sendb(" selected=""selected""") %>>
          <% Sendb(Copient.PhraseLib.Lookup("term.customer", LanguageID))%>
        </option>
      </select>
      <hr class="hidden" />
    </div>
    <div class="box" id="tiers">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.tiers", LanguageID))%>
        </span>
      </h2>
      <% If (IsTemplate) Then%>
      <span class="temp">
        <input type="checkbox" class="tempcheck" id="Disallow_Tiers" name="Disallow_Tiers"
          <% if(disallow_tiers)then send(" checked=""checked""") %> />
        <label for="temp-Tiers">
          <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
      </span>
      <br class="printonly" />
      <% End If%>
      <%
        MyCommon.QueryStr = "select TierTypeID, PhraseID from TierTypes with (NoLock);"
        rst2 = MyCommon.LRT_Select()
        For Each row2 In rst2.Rows
          If row2.Item("TierTypeID") = 3 Then
            If IsTemplate Then
              ' skip tier Type 3
              Continue For
            End If
            If Not bEnableBuckOffers Then
              ' skip tier Type 3
              Continue For
            End If
            If (Not bDisplayBuckTierType) Then
              ' skip tier Type 3
              Continue For
            End If
          End If
          Sendb("<input class=""radio"" type=""radio"" id=""tiertype" & row2.Item("TierTypeID") & """ name=""form_TierTypeID"" value=""" & row2.Item("TierTypeID") & """")
          If (bUseTemplateLocks And Disallow_Tiers) Or ((TierTypeID > 0) AndAlso (row2.Item("TierTypeID") = 0)) Then
            Sendb(" disabled=""disabled""")
          ElseIf (bEnableBuckOffers) AndAlso bDisableEditTierType Then
            Sendb(" disabled=""disabled""")
          End If
          If (TierTypeID = row2.Item("TierTypeID")) Then
            Sendb(" checked=""checked""")
          End If
          Send(" />")
          If (bEnableBuckOffers) And row2.Item("TierTypeID") = 3 And lBuckPromoVarId > 0 Then
            Send("<label for=""tiertype" & row2.Item("TierTypeID") & """>" & Copient.PhraseLib.Lookup(row2.Item("PhraseID"), LanguageID) & " (" & Copient.PhraseLib.Lookup("term.var", LanguageID) & ": " & lBuckPromoVarId & ")</label>")
          Else
            Send("<label for=""tiertype" & row2.Item("TierTypeID") & """>" & Copient.PhraseLib.Lookup(row2.Item("PhraseID"), LanguageID) & "</label>")
          End If
          Send("<br />")
        Next
      %>
      <label for="form_NumTiers">
        <% Sendb(Copient.PhraseLib.Lookup("offer-gen.tiers", LanguageID) & " (" & NumTiers & "-99):")%></label>
      <input class="shortest" id="form_NumTiers" maxlength="2" name="form_NumTiers" type="text"
                value="<% sendb(NumTiers) %>" <% if(bUseTemplateLocks and disallow_tiers)then sendb(" disabled=""disabled""") %> /><br />
      <hr class="hidden" />
    </div>
    <div class="box" id="engines" style="display: none;">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.executionengine", LanguageID))%>
        </span>
      </h2>
      <% If (IsTemplate) Then%>
      <span class="temp">
        <input type="checkbox" class="tempcheck" id="Disallow_ExecutionEngine" name="Disallow_ExecutionEngine"
          <% if(disallow_executionengine)then send(" checked=""checked""") %> />
        <label for="temp-Tiers">
          <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
      </span>
      <br class="printonly" />
      <% End If%>
      <label for="engine">
        <% Sendb(Copient.PhraseLib.Lookup("offer-gen.engine", LanguageID))%></label>
      <br />
      <%' get the EngineID list from database
        MyCommon.QueryStr = "select EngineID, Description, PhraseID from PromoEngines with (NoLock) where Installed=1 or EngineID=" & EngineID & ";"
        rst2 = MyCommon.LRT_Select()
      %>
      <select id="engine" name="form_EngineID" <% if(cmoadeploystatus <> 0 or (bUseTemplateLocks and disallow_executionengine))then sendb(" disabled=""disabled""") %>
        onchange="javascript:if(this.value==1){document.mainform.form_ExtOfferID2.style.visibility='visible';document.getElementById('mclu').style.visibility='visible';document.mainform.form_ExtOfferID2.value=''}else{document.mainform.form_ExtOfferID2.style.visibility='hidden';;document.getElementById('mclu').style.visibility='hidden';document.mainform.form_ExtOfferID2.value='';}">
        <%
          For Each row2 In rst2.Rows
            Sendb("<option value=""" & row2.Item("EngineID") & """" & IIf(EngineID = row2.Item("EngineID"), " selected=""selected""", "") & ">")
            If MyCommon.NZ(row2.Item("PhraseID"), 0) > 0 Then
              Sendb(Copient.PhraseLib.Lookup(row2.Item("PhraseID"), LanguageID))
            Else
              Sendb(MyCommon.NZ(row2.Item("Description"), ""))
            End If
            Send("</option>")
          Next
        %>
      </select>
      <span id="mclu" <%if(engineid=0)then sendb(" style=""visibility:hidden;""") %>>&nbsp;
        <label for="form_ExtOfferID2">
          <% Sendb(Copient.PhraseLib.Lookup("term.mclu", LanguageID))%>:<input class="short"
            id="form_ExtOfferID2" name="form_ExtOfferID2" type="text" value="<% Sendb(extofferid2) %>" /></label><br />
      </span>
      <hr class="hidden" />
    </div>
    <% If (ShowInboundOutboundBox) Then%>
    <div class="box" id="inboundoutbound">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.inbound/outbound", LanguageID))%>
        </span>
      </h2>
      <% If (IsTemplate) Then%>
      <span class="temp">
        <input type="checkbox" class="tempcheck" id="Disallow_CRMEngine" name="Disallow_CRMEngine"
          <% if(disallow_crmengine)then send(" checked=""checked""") %> />
        <label for="temp-crmengine">
          <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
      </span>
      <br class="printonly" />
      <% End If%>
      <% If (Not IsTemplate) Or bCopyInboundCrmEngineID Then%>
      <label for="InboundCRMEngineID" style="position: relative;">
        <% Sendb(Copient.PhraseLib.Lookup("term.creationsource", LanguageID))%>
        :</label>
      <%
        If Logix.UserRoles.EditOfferSource Then
          MyCommon.QueryStr = "select ExtInterfaceID, PhraseID, Name from ExtCRMInterfaces with (NoLock) where Deleted=0 and Active=1 and ExtInterfaceTypeID in (0,1);"
          rst2 = MyCommon.LRT_Select
          If rst2.Rows.Count > 0 Then
            Send("<select id=""InboundCRMEngineID"" name=""InboundCRMEngineID""" & IIf(bUseTemplateLocks And Disallow_CRMEngine, " disabled=""disabled""", "") & ">")
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
      <% End If%>
      <label for="crmengine">
        <% Sendb(Copient.PhraseLib.Lookup("offer-gen.sendoutbound", LanguageID))%>:</label>
      <br />
      <%' get the EngineID list from database
        Dim iCrmIntegration As Integer = 0
        Dim sQuery As String
        If Not Integer.TryParse(MyCommon.Fetch_SystemOption(25), iCrmIntegration) Then iCrmIntegration = 0
        Select Case iCrmIntegration
          Case 0
            sQuery = "-1"
          Case 1
            ' TCRM
            If NumTiers > 1 Then
              sQuery = "-1"
            Else
              sQuery = "1"
            End If
          Case 2
            ' CRM
            sQuery = "0"
          Case Else
            ' All
            If NumTiers > 1 Then
              sQuery = "0"
            Else
              sQuery = "0,1"
            End If
        End Select
        MyCommon.QueryStr = "select ExtInterfaceID, Name, PhraseID from ExtCRMInterfaces with (NoLock) " & _
                            "where Deleted=0 and Active=1 and OutboundEnabled=1 and (ExtInterfaceId=0 or ExtInterfaceTypeID in (" & sQuery & "));"
        rst2 = MyCommon.LRT_Select
        MyCommon.QueryStr = "Select CRMEngineID from Offers with (NoLock) where Deleted=0 and OfferID=" & OfferID & ";"
        crmdt = MyCommon.LRT_Select()
        If crmdt.Rows.Count > 0 Then
          CRMEngineID = MyCommon.NZ(crmdt.Rows(0).Item("CRMEngineID"), 0)
        End If
      %>
      <select id="crmengine" name="form_CRMEngineID" <% if(bUseTemplateLocks and disallow_crmengine)then sendb(" disabled=""disabled""") %>>
        <%
          Dim ExtName As String = ""
            
          For Each row2 In rst2.Rows
            If IsDBNull(row2.Item("PhraseID")) Then
              ExtName = MyCommon.NZ(row2.Item("Name"), "")
            Else
              ExtName = Copient.PhraseLib.Lookup(row2.Item("PhraseID"), LanguageID, MyCommon.NZ(row2.Item("Name"), ""))
            End If
              
            If (CRMEngineID = row2.Item("ExtInterfaceID")) Then
              Sendb("<option value=""" & row2.Item("ExtInterfaceID") & """ selected=""selected"">" & ExtName & "</option>")
            Else
              Sendb("<option value=""" & row2.Item("ExtInterfaceID") & """>" & ExtName & "</option>")
            End If
          Next
        %>
      </select>
      <br />
      <hr class="hidden" />
    </div>
    <% End If%>
    <div class="box" id="sweepstakes">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.instantwin", LanguageID))%>
        </span>
      </h2>
      <% If (IsTemplate) Then%>
      <span class="temp">
        <input type="checkbox" class="tempcheck" id="Disallow_Sweepstakes" name="Disallow_Sweepstakes"
          <% if(disallow_sweepstakes)then send(" checked=""checked""") %> />
        <label for="temp-Sweepstakes">
          <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
      </span>
      <br class="printonly" />
      <% End If%>
      <input class="checkbox" id="instantwin" name="form_InstantWin" type="checkbox" <% if(instantwin)then send(" checked=""checked""")%><% if(bUseTemplateLocks and disallow_sweepstakes)then sendb(" disabled=""disabled""") %> />
      <label for="instantwin">
        <% Sendb(Copient.PhraseLib.Lookup("term.instantwin", LanguageID))%></label>
      <br />
      <br class="half" />
      <label for="prizes">
        <% Sendb(Copient.PhraseLib.Lookup("offer-gen.prizesawarded", LanguageID))%></label>
      <br />
      &nbsp;&nbsp;<input class="short" id="prizes" maxlength="9" name="form_NumPrizesAllowed"
        type="text" value="<%sendb(NumPrizesAllowed) %>" <% if(bUseTemplateLocks and disallow_sweepstakes)then sendb(" disabled=""disabled""") %> /><br />
      <br class="half" />
      <label for="odds">
        <% Sendb(Copient.PhraseLib.Lookup("offer-gen.oddsofwinning", LanguageID))%></label>
      <br />
      1:<input class="short" id="odds" name="form_OddsOfWinning" maxlength="5" type="text"
        value="<%sendb(OddsOfWinning) %>" <% if(bUseTemplateLocks and disallow_sweepstakes)then sendb(" disabled=""disabled""") %> />
      <input id="fixed" name="form_RandomWinners" type="radio" <% if(bUseTemplateLocks and disallow_sweepstakes)then sendb(" disabled=""disabled""") %>
        value="0" <% if(randomwinners=false)then sendb(" checked=""checked""") %> />
      <label for="fixed">
        <% Sendb(Copient.PhraseLib.Lookup("term.fixed", LanguageID))%></label>
      <input id="random" name="form_RandomWinners" type="radio" <% if(bUseTemplateLocks and disallow_sweepstakes)then sendb(" disabled=""disabled""") %>
        value="1" <% if(randomwinners)then sendb(" checked=""checked""") %> />
      <label for="random">
        <% Sendb(Copient.PhraseLib.Lookup("term.random", LanguageID))%></label>
      <br />
      <br class="half" />
      <% Sendb(Copient.PhraseLib.Lookup("offer-gen.oddscalculation", LanguageID))%>
      <br />
      <input id="odds-calconce" name="form_IWTransLevel" type="radio" <% if(bUseTemplateLocks and disallow_sweepstakes)then sendb(" disabled=""disabled""") %>
        value="0" <% if(iwtranslevel=false)then sendb(" checked=""checked""") %> />
      <label for="odds-calconce">
        <% Sendb(Copient.PhraseLib.Lookup("offer-gen.odds1", LanguageID))%></label>
      <br />
      <input id="odds-calceach" name="form_IWTransLevel" type="radio" value="1" <% if(iwtranslevel)then sendb(" checked=""checked""") %><% if(bUseTemplateLocks and disallow_sweepstakes)then sendb(" disabled=""disabled""") %> />
      <label for="odds-calceach">
        <% Sendb(Copient.PhraseLib.Lookup("offer-gen.odds2", LanguageID))%></label>
      <br />
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
      <%
        Send("<input type=""checkbox"" id=""form_DisableDisconnectedOffer"" name=""form_DisableDisconnectedOffer""" & IIf(DisableDisconnectedOffer, " checked=""checked""", "") & IIf((bUseTemplateLocks And Disallow_AdvancedOption), " disabled=""disabled""", "") & " />")
        Send("<label for=""form_DisableDisconnectedOffer"">" & Copient.PhraseLib.Lookup("term.offlinedisable", LanguageID) & "</label>")
        'Send("<label for=""form_DisableDisconnectedOffer"">Disable Disconnected Offer</label>")
        Send("<br />")

        Send("<input type=""checkbox"" id=""form_autotransferable"" name=""form_autotransferable""" & IIf(AutoTransferable, " checked=""checked""", "") & IIf((bUseTemplateLocks And Disallow_AdvancedOption), " disabled=""disabled""", "") & " />")
        Send("<label for=""form_autotransferable"">" & Copient.PhraseLib.Lookup("term.autotransferable", LanguageID) & "</label>")
        Send("<br />")
  
        If (MyCommon.Fetch_SystemOption(73) <> "") Then
          Send("<input type=""checkbox"" id=""form_exportEDW"" name=""form_exportEDW""" & IIf(ExportToEDW, " checked=""checked""", "") & IIf((bUseTemplateLocks And Disallow_AdvancedOption), " disabled=""disabled""", "") & " />")
          Send("<label for=""form_exportEDW"">" & Copient.PhraseLib.Lookup("term.exporttoarchive", LanguageID) & "</label>")
          Send("<br />")
        End If
        'Cashier Prompt for Reward'
        If (MyCommon.Fetch_CM_SystemOption(133) = "1") Then
          MyCommon.QueryStr = "Select CashierApprovalMessage from Offers where OfferID=" & OfferID
		
          rst2 = MyCommon.LRT_Select()
          CashPromptRew_Offer = MyCommon.NZ(rst2.Rows(0).Item("CashierApprovalMessage"), False)
		
		
          Send("<input type=""checkbox"" id=""cash_prompt_rew"" name=""cash_prompt_rew"" " & IIf(CashPromptRew_Offer, " checked=""checked""", "") & " />")
          Send("<label for=""cash_prompt_rew"">" & Copient.PhraseLib.Lookup("term.cashpromptrew", LanguageID) & "</label>")
          Send("<br />")
        End If
        'E-coupon Promotion
        If (MyCommon.Fetch_CM_SystemOption(39) = "1") Then 'System option to allow making offers e-coupon promotions
          Dim ecoupon_options as String = IIf((bUseTemplateLocks And Disallow_AdvancedOption), " disabled=""disabled""", "")

          If (MyCommon.Fetch_CM_SystemOption(117) = "1") Then
            Dim AmountType As Integer
            MyCommon.QueryStr = "Select RewardAmountTypeID from OfferRewards where OfferID=" & OfferID & " and RewardTypeID=1 and Deleted=0"
            rst2 = MyCommon.LRT_Select()
            If rst2.Rows.Count > 0 Then
              AmountType = MyCommon.NZ(rst2.Rows(rst2.Rows.Count - 1).Item("RewardAmountTypeID"), 0)
            End If
					
            If (AmountType = 1 Or AmountType = 3 Or AmountType = 4 Or AmountType = 8) Then
              defaultPriority = MyCommon.Fetch_CM_SystemOption(122)
              defaultECPriority = MyCommon.Fetch_CM_SystemOption(123)
            ElseIf (AmountType = 2 Or AmountType = 9) Then
              defaultPriority = MyCommon.Fetch_CM_SystemOption(124)
              defaultECPriority = MyCommon.Fetch_CM_SystemOption(125)
            ElseIf (AmountType = 5 Or AmountType = 6) Then
              defaultPriority = MyCommon.Fetch_CM_SystemOption(120)
              defaultECPriority = MyCommon.Fetch_CM_SystemOption(121)
            ElseIf (AmountType = 7) Then
              defaultPriority = MyCommon.Fetch_CM_SystemOption(118)
              defaultECPriority = MyCommon.Fetch_CM_SystemOption(119)
					
            End If
            ecoupon_options = "onclick=""resetPriority("& defaultPriority &","& defaultECPriority &");"" "

          End If
          Send("<input type=""checkbox"" id=""form_eCoupon"" name=""form_eCoupon"" " & IIf(eCoupon_Offer, " checked=""checked""", "") & ecoupon_options & " />")
          Send("<label for=""form_eCoupon"">" & Copient.PhraseLib.Lookup("term.ecouponpromotion", LanguageID) & "</label>")
          Send("<br />")
        End If
        'High Value Promotion
        If (MyCommon.Fetch_CM_SystemOption(72) = "1") Then 'System option to allow making offers e-coupon promotions
          Send("<input type=""checkbox"" id=""form_highvalue"" name=""form_highvalue"" " & IIf(HighValue, " checked=""checked""", "") & IIf((bUseTemplateLocks And Disallow_AdvancedOption), " disabled=""disabled""", "") & " />")
          Send("<label for=""form_highvalue"">" & Copient.PhraseLib.Lookup("term.highvalue", LanguageID) & "</label>")
          Send("<br />")
        End If
        'auto translate offer
        If (MyCommon.Fetch_CM_SystemOption(135) = "1") Then 'Trnslate CM offers to UE is enabled
          Send("<input type=""checkbox"" id=""form_autotranslate"" name=""form_autotranslate"" " & IIf(bAutoTranslate, " checked=""checked""", "") & IIf((bUseTemplateLocks And Disallow_AdvancedOption), " disabled=""disabled""", "") & " />")
          Send("<label for=""form_autotranslate"">" & Copient.PhraseLib.Lookup("term.autotranslate", LanguageID) & "</label>")
          Send("<br />")
		End If
        'FOB Eligible
        If bFobEligibilityEnabled And (InboundCRMEngineID > 0) Then 'FOB is enabled
          Send("<input type=""checkbox"" id=""form_fob"" name=""form_fob"" " & IIf(bFobEligible, " checked=""checked""", "") & IIf((bUseTemplateLocks And Disallow_AdvancedOption), " disabled=""disabled""", "") & " />")
          Send("<label for=""form_fob"">" & Copient.PhraseLib.Lookup("term.fobeligible", LanguageID) & "</label>")
          Send("<br />")
        End If
        If (Logix.UserRoles.FavoriteOffersForOthers AndAlso Not IsTemplate) Then
          Send("<br class=""half"" />")
          MyCommon.QueryStr = "select AdminUserID from AdminUserOffers where OfferID=" & OfferID & ";"
          rst2 = MyCommon.LRT_Select
          MyCommon.QueryStr = "select AdminUserID from AdminUsers;"
          rst3 = MyCommon.LRT_Select
          Send("<button type=""button"" id=""favorite"" name=""favorite"" value=""favorite""" & IIf((bUseTemplateLocks And Disallow_AdvancedOption), " disabled=""disabled""", "") & "onclick=""javascript:xmlhttpPost('OfferFeeds.aspx', 'FavoriteForAll');"">" & Copient.PhraseLib.Lookup("offer-gen.favoriteall", LanguageID) & "</button>")
          Sendb("<a href=""javascript:openPopup('offer-favorite.aspx?OfferID=" & OfferID & "&bUseTemplateLocks=" & bUseTemplateLocks & "&Disallow_AdvancedOption=" & Disallow_AdvancedOption & "')""><img id=""favImg"" src=""../images/user.png"" ")
          Sendb("alt=""" & rst2.Rows.Count & " " & StrConv(Copient.PhraseLib.Lookup("term.of", LanguageID), VbStrConv.Lowercase) & " " & rst3.Rows.Count & " " & StrConv(Copient.PhraseLib.Lookup("term.users", LanguageID), VbStrConv.Lowercase) & " " & Copient.PhraseLib.Lookup("offer-gen.havefavorite", LanguageID) & """ ")
          Sendb("title=""" & rst2.Rows.Count & " " & StrConv(Copient.PhraseLib.Lookup("term.of", LanguageID), VbStrConv.Lowercase) & " " & rst3.Rows.Count & " " & StrConv(Copient.PhraseLib.Lookup("term.users", LanguageID), VbStrConv.Lowercase) & " " & Copient.PhraseLib.Lookup("offer-gen.havefavorite", LanguageID) & """ ")
          Send("/></a><br />")
        End If
      %>
    </div>
    &nbsp;&nbsp;
  </div>
  <br clear="all" />
</div>
</form>
<script type="text/javascript">
<% Send_Date_Picker_Terms() %>

  setlimitsection(false);
  setperiodsection(false);

</script>
<script>
  function resetPriority(priority_default, eCoupon_default) {
    if (document.getElementById("form_eCoupon").checked) {
      document.getElementById("priority").value = eCoupon_default;
    }
    else { document.getElementById("priority").value = priority_default; }
  }
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
  Send_FocusScript("mainform", "form_Name")
  Send_WrapEnd()
  Send("<div id=""disabledBkgrd"" style=""position:absolute;top:0px;left:0px;right:0px;width:100%;height:100%;background-color:Gray;display:none;z-index:99;filter:alpha(opacity=50);-moz-opacity:.50;opacity:.50;""></div>")
  Send_PageEnd()
  MyCommon.Close_LogixRT()
  MyCommon = Nothing
  Logix = Nothing
%>