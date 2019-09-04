<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<% 
  ' *****************************************************************************
  ' * FILENAME: web-offer-gen.aspx 
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
  Dim rstTemplates As DataTable
  Dim rst3 As DataTable
  Dim row As DataRow
  Dim rowTemplates As DataRow
  Dim OfferID As Long = Request.QueryString("OfferID")
  Dim OfferName As String = ""
  Dim StatusFlag As Integer
  Dim rst2 As DataTable
  Dim row2 As DataRow
  Dim IsTemplate As Boolean
  Dim IsTemplateVal As String = ""
  Dim ActiveSubTab As Integer = 91
  Dim IntroID As String = "intro"
  Dim OptOutChecked As String = ""
  Dim AllowOptOut As Boolean = False
  Dim Disallow_EmployeeFiltering As Boolean = True
  Dim Disallow_ProductionDates As Boolean = True
  Dim Disallow_Tiers As Boolean = True
  Dim Disallow_Conditions As Boolean = True
  Dim Disallow_Rewards As Boolean = True
  Dim FromTemplate As Boolean
  Dim DisabledOnCFW As Boolean
  Dim EmployeesExcluded As Boolean
  Dim EmployeesOnly As Boolean
  Dim ReportingImp As Boolean = False
  Dim ReportingRed As Boolean = False
  Dim EmployeeFiltered As Boolean
  Dim ExtOfferID As String = ""
  Dim ProdStartDate As Date
  Dim ProdEndDate As Date
  Dim StartDateParsed, EndDateParsed As Boolean
  Dim roid As Integer
  Dim DuplicateName As Boolean = False
  Dim infoMessage As String = ""
  Dim modMessage As String = ""
  Dim Handheld As Boolean = False
  Dim sqlBuf As New StringBuilder()
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
  Dim ChargebackVendorID As Integer = 0
  Dim EngineID As Integer = 3
  Dim EngineSubTypeID As Integer = 0
  Dim EnginePhraseID As Integer = 0
  Dim EngineSubTypePhraseID As Integer = 0
  Dim SelectedList As String = ""
  Dim TierLevels As Integer = 1
  Dim MaxTiers As Integer = 1
  Dim ShortStartDate, ShortEndDate As String
  Dim StartDT, EndDT As Date
  Dim DescriptLength As Boolean = False
  Dim Description As String = ""
  Dim DisplayTierLevel As String = ""
  Dim AutoTransferable As Boolean = False
  Dim MLI As New Copient.Localization.MultiLanguageRec
  Dim selectDatePicker As Integer = MyCommon.Extract_Val(MyCommon.NZ(MyCommon.Fetch_SystemOption(161), 0))
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "web-offer-gen.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  Localization = New Copient.Localization(MyCommon)
  
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
    
    If (Request.QueryString("productionstart") = "" Or Request.QueryString("productionend") = "") Then
      infoMessage = Copient.PhraseLib.Lookup("CPEoffer_gen.noproductiondates", LanguageID)
    ElseIf DescriptLength = False Then
      infoMessage = Copient.PhraseLib.Lookup("error.description", LanguageID)
    Else
      StartDateParsed = Date.TryParse(Request.QueryString("productionstart"), ProdStartDate)
      EndDateParsed = Date.TryParse(Request.QueryString("productionend"), ProdEndDate)
      If (StartDateParsed AndAlso EndDateParsed) Then
        ' Check for an incentive already with that name
        MyCommon.QueryStr = "select IncentiveName from CPE_Incentives with (NoLock) " & _
                            "where Deleted=0 and IncentiveName='" & MyCommon.Parse_Quotes(Request.QueryString("form_name")) & "' and IncentiveID<>" & Request.QueryString("OfferID") & _
                            " union all " & _
                            "select Name from Offers with (NoLock) " & _
                            "where Deleted=0 and Name='" & MyCommon.Parse_Quotes(Request.QueryString("form_name")) & "' and OfferID<>" & Request.QueryString("OfferID") & ";"
        rst = MyCommon.LRT_Select
        DuplicateName = (rst.Rows.Count > 0)
        ' Also, run a query to see if there's a category that has this offer as its base offer
        MyCommon.QueryStr = "select OfferCategoryID from OfferCategories where Deleted=0 and BaseOfferID=" & OfferID & " and OfferCategoryID=(" & _
                            "  select IsNull(PromoClassID, 0) from CPE_Incentives where IncentiveID=" & OfferID & ");"
        rst2 = MyCommon.LRT_Select
        sqlBuf.Append("Update CPE_Incentives with (RowLock) set ")
        sqlBuf.Append("IncentiveName=N'" & MyCommon.Parse_Quotes(Logix.TrimAll(Request.QueryString("form_name"))) & "',")
        sqlBuf.Append("Description=N'" & MyCommon.Parse_Quotes(Request.QueryString("form_description")) & "',")
        If (Request.QueryString("form_Category") <> "") Then
          sqlBuf.Append("PromoClassID=" & Request.QueryString("form_Category") & ",")
        End If
        sqlBuf.Append("StartDate='" & ProdStartDate.ToShortDateString() & "',")
        sqlBuf.Append("EndDate='" & ProdEndDate.ToShortDateString() & "',")
        sqlBuf.Append("DisabledOnCFW=" & IIf(Request.QueryString("DisabledOnCFW") = "on", 1, 0) & ",")
        sqlBuf.Append("EmployeesOnly=" & IIf(Request.QueryString("employeesonly") = "on", 1, 0) & ",")
        sqlBuf.Append("EnableImpressRpt=" & IIf(Request.QueryString("reportingimp") = "on", 1, 0) & ",")
        sqlBuf.Append("EnableRedeemRpt=" & IIf(Request.QueryString("reportingred") = "on", 1, 0) & ",")
        sqlBuf.Append("EmployeesExcluded=" & IIf(Request.QueryString("employeesexcluded") = "on", 1, 0) & ",")
        sqlBuf.Append("ExportToEDW=" & IIf(Request.QueryString("exporttoedw") = "on", 1, 0) & ",")
        sqlBuf.Append("Favorite=" & IIf(Request.QueryString("favorite") = "on", 1, 0) & ",")
        sqlBuf.Append("AllowOptOut=" & IIf(Request.QueryString("allowoptout") = "on", 1, 0) & ",")
        sqlBuf.Append("SendIssuance=" & IIf(Request.QueryString("issuance") = "on", 1, 0) & ",")
        sqlBuf.Append("AutoTransferable=" & IIf(Request.QueryString("autotransferable") = "on", 1, 0) & ",")
        sqlBuf.Append("LastUpdate=getdate(), ")
        sqlBuf.Append("LastUpdatedByAdminID=" & AdminUserID & ", ")
        sqlBuf.Append("StatusFlag=1 ")
        sqlBuf.Append("where IncentiveID=" & Request.QueryString("OfferID"))
        MyCommon.QueryStr = sqlBuf.ToString
        
        'Send(MyCommon.QueryStr)
        If (ProdEndDate < ProdStartDate) Then
          infoMessage = Copient.PhraseLib.Lookup("offer-gen.baddate", LanguageID)
        ElseIf DuplicateName Then
          infoMessage = Copient.PhraseLib.Lookup("offer-gen.nameused", LanguageID)
        ElseIf Logix.TrimAll(Request.QueryString("form_name")) = "" Then
          infoMessage = Copient.PhraseLib.Lookup("offer-gen.noname", LanguageID)
        ElseIf (Request.QueryString("tierlevels") <> "" AndAlso (Not Integer.TryParse(Request.QueryString("tierlevels"), TempInt) OrElse TempInt < 1 OrElse TempInt > MaxTiers)) Then
          infoMessage = Copient.PhraseLib.Lookup("offer-gen.invalidtierscpe", LanguageID) & MaxTiers & "."
        ElseIf ((rst2.Rows.Count > 0) AndAlso (MyCommon.Extract_Val(Request.QueryString("form_Category")) <> MyCommon.NZ(rst2.Rows(0).Item("OfferCategoryID"), 0))) Then
          infoMessage = Copient.PhraseLib.Lookup("ueoffer-gen.InvalidCategoryChange", LanguageID)
        Else
          MyCommon.LRT_Execute()
          MyCommon.QueryStr = "update CPE_RewardOptions set TierLevels=" & MyCommon.Extract_Val(Request.QueryString("tierlevels")) & " " & _
                              "where RewardOptionID=" & roid & ";"
          MyCommon.LRT_Execute()
        End If
        
        IsTemplate = (Request.QueryString("IsTemplate") = "IsTemplate")
        If (IsTemplate) Then
          'Update template permissions
          Dim form_Disallow_ProductionDates As Integer = IIf(Request.QueryString("Disallow_ProductionDates") = "on", 1, 0)
          Dim form_Disallow_Tiers As Integer = IIf(Request.QueryString("Disallow_Tiers") = "on", 1, 0)
          MyCommon.QueryStr = "update TemplatePermissions with (RowLock) set " & _
                              " Disallow_ProductionDates=" & form_Disallow_ProductionDates & ", " & _
                              " Disallow_Tiers=" & form_Disallow_Tiers & _
                              " where OfferID=" & OfferID
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
    'Update multi-language:
    MLI.ItemID = OfferID
    MLI.MLTableName = "OfferTranslations"
    MLI.MLIdentifierName = "OfferID"
    MLI.StandardTableName = "CPE_Incentives"
    MLI.StandardIdentifierName = "IncentiveID"
    'Name
    MLI.MLColumnName = "OfferName"
    MLI.StandardColumnName = "IncentiveName"
    MLI.StandardValue = OfferName
    MLI.InputName = "form_Name"
    Localization.SaveTranslationInputs(MyCommon, MLI, Request.QueryString, 2)
    'Description
    MLI.MLColumnName = "OfferDescription"
    MLI.StandardColumnName = "Description"
    MLI.StandardValue = Description
    MLI.InputName = "form_Description"
    Localization.SaveTranslationInputs(MyCommon, MLI, Request.QueryString, 2)
  End If
  
  If (Request.QueryString("OfferID") <> "") Then
    MyCommon.QueryStr = "Select IncentiveID, OID.EngineID, PE.PhraseID as EnginePhraseID, PEST.PhraseID as EngineSubTypePhraseID, " & _
                        "IsTemplate, FromTemplate, ClientOfferID, IncentiveName, CPE.Description, PromoClassID, CRMEngineID, Priority, " & _
                        "StartDate, EndDate, EveryDOW, EligibilityStartDate, EligibilityEndDate, TestingStartDate, TestingEndDate, P1DistQtyLimit, P1DistTimeType, P1DistPeriod, " & _
                        "P2DistQtyLimit, P2DistTimeType, P2DistPeriod, P3DistQtyLimit, P3DistTimeType, P3DistPeriod, EnableImpressRpt, EnableRedeemRpt, CPE.CreatedDate, CPE.LastUpdate, CPE.Deleted, CPEOARptDate, " & _
                        "CPEOADeploySuccessDate, CPEOADeployRpt, CRMRestricted, StatusFlag, DisabledOnCFW, EmployeesOnly, EmployeesExcluded, DeferCalcToEOS, ExportToEDW, " & _
                        "Favorite, OC.Description as CategoryName, AllowOptOut, SendIssuance, InboundCRMEngineID, ChargebackVendorID, ManufacturerCoupon, AutoTransferable, CPE.EngineSubTypeID " & _
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
      EmployeesOnly = MyCommon.NZ(rst.Rows(0).Item("EmployeesOnly"), False)
      ReportingImp = MyCommon.NZ(rst.Rows(0).Item("EnableImpressRpt"), False)
      ReportingRed = MyCommon.NZ(rst.Rows(0).Item("EnableRedeemRpt"), False)
      EmployeesExcluded = MyCommon.NZ(rst.Rows(0).Item("EmployeesExcluded"), False)
      EmployeeFiltered = EmployeesOnly Or EmployeesExcluded
      ExtOfferID = MyCommon.NZ(rst.Rows(0).Item("ClientOfferID"), "")
      ExportToEDW = MyCommon.NZ(rst.Rows(0).Item("ExportToEDW"), False)
      Favorite = MyCommon.NZ(rst.Rows(0).Item("Favorite"), False)
      AllowOptOut = MyCommon.NZ(rst.Rows(0).Item("AllowOptOut"), False)
      OptOutChecked = IIf(AllowOptOut, " checked=""checked""", "")
      Issuance = (MyCommon.NZ(rst.Rows(0).Item("SendIssuance"), 0) = 1)
      ChargebackVendorID = MyCommon.NZ(rst.Rows(0).Item("ChargebackVendorID"), 0)
      AutoTransferable = MyCommon.NZ(rst.Rows(0).Item("AutoTransferable"), False)
      EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), 3)
      EnginePhraseID = MyCommon.NZ(rst.Rows(0).Item("EnginePhraseID"), 0)
      EngineSubTypeID = MyCommon.NZ(rst.Rows(0).Item("EngineSubTypeID"), 0)
      EngineSubTypePhraseID = MyCommon.NZ(rst.Rows(0).Item("EngineSubTypePhraseID"), 0)
    End If
    
    If (IsTemplate Or FromTemplate) Then
      ' lets dig the permissions if its a template
      MyCommon.QueryStr = "select * from TemplatePermissions with (NoLock) where OfferID=" & OfferID
      rstTemplates = MyCommon.LRT_Select
      If (rstTemplates.Rows.Count > 0) Then
        For Each rowTemplates In rstTemplates.Rows
          ' ok there are some rows for the template
          Disallow_EmployeeFiltering = MyCommon.NZ(rowTemplates.Item("Disallow_EmployeeFiltering"), True)
          Disallow_ProductionDates = MyCommon.NZ(rowTemplates.Item("Disallow_ProductionDates"), True)
          Disallow_Tiers = MyCommon.NZ(rowTemplates.Item("Disallow_Tiers"), True)
          Disallow_Conditions = MyCommon.NZ(rowTemplates.Item("Disallow_Conditions"), True)
          Disallow_Rewards = MyCommon.NZ(rowTemplates.Item("Disallow_Rewards"), True)
        Next
      Else
        Disallow_EmployeeFiltering = False
        Disallow_ProductionDates = False
        Disallow_Tiers = False
        Disallow_Conditions = False
        Disallow_Rewards = False
      End If
    End If
  End If
  
  StatusText = Logix.GetOfferStatus(OfferID, LanguageID, StatusCode)
  
  If (IsTemplate) Then
    ActiveSubTab = 27
    IntroID = "intro"
    IsTemplateVal = "IsTemplate"
  Else
    ActiveSubTab = 27
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
window.name = "WebOfferGen"
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
    retVal = isDate(elemEnd.value);
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
<form id="mainform" name="mainform" action="web-offer-gen.aspx" method="get" onsubmit="handleFormElements(this, false);return promptForDeploy();">
  <input type="hidden" name="OfferID" id="OfferID" value="<%Sendb(OfferID)%>" />
  <input type="hidden" name="IsActive" id="IsActive" value="<%Sendb(IIf(StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_ACTIVE, "true", "false")) %>" />
  <input type="hidden" name="IsTemplate" id="IsTemplate" value="<% Sendb(IsTemplateVal)%>" />
  <input type="hidden" name="Popup" id="Popup" value="<%Sendb(IIf(Popup, 1, 0)) %>" />
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
        If (Logix.UserRoles.EditOffer) Then
          Send_Save()
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
          Send(Copient.PhraseLib.Lookup("term.roid", LanguageID) & ": " & roid & "<br />")
          Send(Copient.PhraseLib.Lookup("term.engine", LanguageID) & ": " & Copient.PhraseLib.Lookup(EnginePhraseID, LanguageID) & IIf(EngineSubTypePhraseID > 0, " " & Copient.PhraseLib.Lookup(EngineSubTypePhraseID, LanguageID), "") & "<br />")
          Send(Copient.PhraseLib.Lookup("term.status", LanguageID) & ": " & StatusText & "<br />")
        %>
        <br class="half" />
        <%
          MLI.MLTableName = "OfferTranslations"
          MLI.MLIdentifierName = "OfferID"
          MLI.StandardTableName = "CPE_Incentives"
          MLI.StandardIdentifierName = "IncentiveID"
          'Name input
          MLI.MLColumnName = "OfferName"
          MLI.StandardValue = OfferName.Replace("""", "&quot;")
          MLI.InputName = "form_Name"
          MLI.InputID = "name"
          MLI.InputType = "text"
          MLI.LabelPhrase = "term.name"
          MLI.MaxLength = 100
          MLI.CSSClass = "longest"
          MLI.CSSStyle = "width:92%;"
          Send(Localization.SendTranslationInputs(MyCommon, MLI, Request.QueryString, 2))
          'Description input
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
        <small><%Send("(" & Copient.PhraseLib.Lookup("CPEoffergen.description", LanguageID) & ")")%></small><br class="half" /><br class="half" />
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
        <br />
        <br class="half" />
        <%
          Sendb("<input type=""checkbox"" id=""allowoptout"" name=""allowoptout""" & OptOutChecked & " />")
          Send("<label for=""allowoptout"">" & Copient.PhraseLib.Lookup("web-offer.allowoptout", LanguageID) & "</label><br />")
        %>
        <br />
        <br class="half" />
        <hr class="hidden" />
      </div>
    </div>
    
    <div id="gutter">
    </div>
    
    <div id="column2">
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
              <img src="../images/calendar.png" class="calendar" id="production-start-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('productionstart', event);" />
              <% Sendb(Copient.PhraseLib.Lookup("term.to", LanguageID))%>
              <input type="text" class="short" id="productionend" name="productionend" maxlength="10" value="<% sendb(ShortEndDate) %>"<% if(FromTemplate and disallow_productiondates)then sendb(" disabled=""disabled""") %> />
              <img src="../images/calendar.png" class="calendar" id="production-end-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="autoDatePicker('productionend', event, <% sendb(selectDatePicker) %>);" />
              (<% Sendb(Copient.PhraseLib.Lookup("term.mmddyyyy", LanguageID))%>)<br />
              <br class="half" />
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
      
      <%
        If (MaxTiers > 1) AndAlso (EngineSubTypeID <> 1) Then
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
          Send("  <label for=""tierlevels"">" & Copient.PhraseLib.Lookup("offer-gen.tiers", LanguageID) & " (" & StrConv(Copient.PhraseLib.Lookup("term.upto", LanguageID), VbStrConv.Lowercase) & " " & MaxTiers & "):</label>")
          Send("  <input type=""text"" class=""shortest"" id=""tierlevels"" name=""tierlevels"" maxlength=""2"" value=""" & DisplayTierLevel & """" & IIf(FromTemplate And Disallow_Tiers, " disabled=""disabled""", "") & " /><br />")
          Send("  <hr class=""hidden"" />")
          Send("</div>")
        Else
          Send("<input type=""hidden"" id=""tierlevels"" name=""tierlevels"" value=""1"" />")
        End If
      %>
      
      <div class="box" id="options">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.advancedoptions", LanguageID))%>
          </span>
        </h2>
        <input type="checkbox" name="autotransferable" id="autotransferable"<%Sendb(IIf(AutoTransferable, " checked=""checked""", ""))%> />
        <label for="autotransferable"><% Sendb(Copient.PhraseLib.Lookup("term.autotransferable", LanguageID))%></label>
        <br />
        <%
          If (Logix.UserRoles.FavoriteOffersForOthers AndAlso Not isTemplate) Then
            Send("<br class=""half"" />")
            MyCommon.QueryStr = "select AdminUserID from AdminUserOffers where OfferID=" & OfferID & ";"
            rst2 = MyCommon.LRT_Select
            MyCommon.QueryStr = "select AdminUserID from AdminUsers;"
            rst3 = MyCommon.LRT_Select
            Send("<button type=""button"" id=""favorite"" name=""favorite"" value=""favorite"" onclick=""javascript:xmlhttpPost('OfferFeeds.aspx', 'FavoriteForAll');"">" & Copient.PhraseLib.Lookup("offer-gen.favoriteall", LanguageID) & "</button>")
            Sendb("<a href=""javascript:openPopup('offer-favorite.aspx?OfferID=" & OfferID & "')""><img id=""favImg"" src=""../images/user.png"" ")
            Sendb("alt=""" & rst2.Rows.Count & " " & StrConv(Copient.PhraseLib.Lookup("term.of", LanguageID), VbStrConv.Lowercase) & " " & rst3.Rows.Count & " " & StrConv(Copient.PhraseLib.Lookup("term.users", LanguageID), VbStrConv.Lowercase) & " " & Copient.PhraseLib.Lookup("offer-gen.havefavorite", LanguageID) & """ ")
            Sendb("title=""" & rst2.Rows.Count & " " & StrConv(Copient.PhraseLib.Lookup("term.of", LanguageID), VbStrConv.Lowercase) & " " & rst3.Rows.Count & " " & StrConv(Copient.PhraseLib.Lookup("term.users", LanguageID), VbStrConv.Lowercase) & " " & Copient.PhraseLib.Lookup("offer-gen.havefavorite", LanguageID) & """ ")
            Send("/></a><br />")
          End If
        %>
        <hr class="hidden" />
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
          
          ' get all the assigned banners for web
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
          
          ' get all the assigned ALL banners for web
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
  
  if (targetDateField.id == "productionstart") {
    // populate productionend, etc. if unpopulated
    if (elemProdEnd != null && elemProdEnd.value == "") {
      elemProdEnd.value = targetDateField.value;
    }
  }
}
</script>
<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (OfferID > 0 And Logix.UserRoles.AccessNotes) Then
      Send_Notes(3, OfferID, AdminUserID)
    End If
  End If
done:
  Send_FocusScript()
  Send_WrapEnd()
  Send("<div id=""disabledBkgrd"" style=""position:absolute;top:0px;left:0px;right:0px;width:100%;height:100%;background-color:Gray;display:none;z-index:99;filter:alpha(opacity=50);-moz-opacity:.50;opacity:.50;""></div>")
  Send_PageEnd()
  MyCommon.Close_LogixRT()
  MyCommon = Nothing
  Logix = Nothing
%>
