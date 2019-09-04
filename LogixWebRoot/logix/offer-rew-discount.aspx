<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Text.RegularExpressions" %>
<%
  ' *****************************************************************************
  ' * FILENAME: offer-rew-discount.aspx 
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
  
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim MyImport As New Copient.ImportXml(MyCommon)

  Dim AdminUserID As Long
  Dim rst As DataTable
  Dim rst1 As DataTable
  Dim rst2 As DataTable
  Dim rst3 As DataTable
  Dim row As DataRow
  Dim row2 As DataRow
  Dim OfferID As Long
  Dim Name As String = ""
  Dim RewardID As String
  Dim ExcludedItem As Integer
  Dim SelectedItem As Integer
  Dim NumTiers As Integer
  Dim LinkID As Long
  Dim VarID As String
  Dim RewardAmountTypeID As Integer
  Dim TriggerQty As Integer
  Dim ApplyToLimit As Integer
  Dim DoNotItemDistribute As Boolean
  Dim TransactionLevelPossible As Boolean
  Dim TransactionLevelSelected As Boolean
  Dim StoredValueSelected As Boolean
  Dim FreeItemSelected As Boolean
  Dim DistPeriod As Integer
  Dim AllowNegative As Boolean
  Dim ChargeBackDeptID As Integer
  Dim BestDeal As Boolean
  Dim StaticFuel As Boolean
  Dim DiscountableItemsOnly As Boolean
  Dim UseSpecialPricing As Boolean
  Dim SPRepeatAtOccur As Integer
  Dim ValueRadio As Integer
  Dim PrintLineText As String = ""
  Dim WebText As String = ""
  Dim q As Integer
  Dim x As Integer
  Dim Tiered As Integer
  Dim SponsorID As Integer
  Dim PromoteToTransLevel As Boolean
  Dim SameItem As Boolean
  Dim EffectMinOrder As Boolean
  Dim VirtualLink As Boolean
  Dim RewardLimit As Decimal
  Dim RewardLimitTypeID As Integer
  Dim InvalidRewardAmt As Boolean
  Dim FlagTier As Integer
  Dim Disallow_Edit As Boolean = True
  Dim IsTemplate As Boolean = False
  Dim CloseAfterSave As Boolean = False
  Dim DisabledAttribute As String = ""
  Dim OfferEngineID As Long = 0
  Dim RequirePG As Boolean = False
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim ProductGroupID As Integer = 0
  Dim ExcludedID As Integer = 0
  Dim ProgramID As Long = 0
  Dim SVLinkID As Long = 0
  Dim SVType As Integer = 0
  Dim BannersEnabled As Boolean = False
  Dim WeightUom As Integer
  Dim sQuery As String = ""
  Dim AllBanners As Boolean = False

  Dim bUseTemplateLocks As Boolean
  Dim bDisallowEditPg As Boolean = False
  Dim bDisallowEditDept As Boolean = False
  Dim bDisallowEditSpon As Boolean = False
  Dim bDisallowEditMsg As Boolean = False
  Dim bDisallowEditPp As Boolean = False
  Dim bDisallowEditDist As Boolean = False
  Dim bDisallowEditLimit As Boolean = False
  Dim bDisallowEditSpc As Boolean = False
  Dim bDisallowEditAdv As Boolean = False
  Dim sDisabled As String
  Dim AdvancedLimitID As Long
  Dim RECORD_LIMIT As Integer = GroupRecordLimit '500
  Dim topString As String = ""
  Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249) = "1", True, False)
  Dim GroupSize As Integer
  Dim ProductTypeID As Integer = 0
  Dim rstItems As DataTable = Nothing
  Dim AmountType As Integer
  Dim ProductsWithoutDesc As Integer
  Dim ShowAllItems As Boolean
  Dim ListBoxSize As Integer
  Dim descriptionItem As String = String.Empty
  Dim ExtProductID As String = ""
  Dim prodDT As DataTable
  Dim Description As String = ""
  Dim outputStatus As Integer
  Dim OfferStartDate As Date
  Dim GName As String = ""
  Dim PagePostBack As Boolean = True
  Dim ByExistingPGSelector As Boolean
  Dim ByAddSingleProduct As Boolean = True
  Dim tempProducts As String = ""
  Dim tempProductsList() As String = Nothing
  Dim validItemList As List(Of String) = New List(Of String)
  Dim invalidItemList As List(Of String) = New List(Of String)
  Dim DuplicateItemCount As Integer = 0
  Dim OutputString As String = ""
  Dim tempTableInsertStatement As StringBuilder = New StringBuilder()
  Dim maxLimit As Integer = 0
  Dim upc As String = ""
  Dim PriorityFlag As Boolean
  Dim bStoreUser As Boolean = False
  Dim sValidLocIDs As String = ""
  Dim sValidSU As String = ""
  Dim wherestr As String = ""
  Dim sJoin As String = ""
  Dim iLen As Integer = 0
  Dim i As Integer = 0
  Dim IDLength As Integer = 0
  Dim bCreateGroupOrProgramFromOffer As Boolean
  Dim bStaticPG As Boolean
  
  If RECORD_LIMIT > 0 Then topString = "top " & RECORD_LIMIT
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "offer-rew-discount.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  'Store User
  If (MyCommon.Fetch_CM_SystemOption(131) = "1") Then
    MyCommon.QueryStr = "select LocationID from StoreUsers where UserID=" & AdminUserID & ";"
    rst = MyCommon.LRT_Select
    iLen = rst.Rows.Count
    If iLen > 0 Then
      bStoreUser = True
      sValidSU = AdminUserID
      For i = 0 To (iLen - 1)
        If i = 0 Then
          sValidLocIDs = rst.Rows(0).Item("LocationID")
        Else
          sValidLocIDs &= "," & rst.Rows(i).Item("LocationID")
        End If
      Next
    
      MyCommon.QueryStr = "select UserID from StoreUsers where LocationID in (" & sValidLocIDs & ") and NOT UserID=" & AdminUserID & ";"
      rst = MyCommon.LRT_Select
      iLen = rst.Rows.Count
      If iLen > 0 Then
        For i = 0 To (iLen - 1)
          sValidSU &= "," & rst.Rows(i).Item("UserID")
        Next
      End If
    End If
  End If
  
  TransactionLevelPossible = False
  TransactionLevelSelected = False
  StoredValueSelected = False
  FreeItemSelected = False
  OfferID = Request.QueryString("OfferID")
  RewardID = Request.QueryString("RewardID")
  NumTiers = Request.QueryString("NumTiers")
  ProductGroupID = MyCommon.Extract_Val(Request.QueryString("ProductGroupID"))
  ExcludedID = MyCommon.Extract_Val(Request.QueryString("ExcludedID"))
  
  
  ByExistingPGSelector = IIf(MyCommon.Fetch_SystemOption(222) = "0", True, False)
  bCreateGroupOrProgramFromOffer = IIf(MyCommon.Fetch_CM_SystemOption(134) = "1", True, False)
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  
  If (MyCommon.Fetch_CM_SystemOption(117)) Then
    MyCommon.QueryStr = "Select RewardAmountTypeID from OfferRewards where RewardID=" & RewardID
    rst = MyCommon.LRT_Select
	
    If (MyCommon.NZ(rst.Rows(0).Item("RewardAmountTypeID"), True) = True) Then
      PriorityFlag = True
    Else
      PriorityFlag = False
    End If
	
  End If
	
  
  If (Request.QueryString("save") <> "") Then
    If Request.QueryString("InvalidRewardAmt") = "true" Then
      infoMessage = Copient.PhraseLib.Lookup("CPEoffer-rew-discount.InvalidAmt", LanguageID)
    ElseIf (Request.QueryString("tier0") <> "") And (Not IsNumeric(Request.QueryString("tier0"))) Then
      infoMessage = Copient.PhraseLib.Lookup("CPEoffer-rew-discount.InvalidAmt", LanguageID)
    Else
      CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
    End If
  End If
    
  If (Request.QueryString("addvalue") <> "") Then
    If Request.QueryString("InvalidRewardAmt") = "true" Then
      infoMessage = Copient.PhraseLib.Lookup("CPEoffer-rew-discount.InvalidAmt", LanguageID)
    End If
  End If
    
  ' fetch the name
  MyCommon.QueryStr = "Select Name,IsTemplate,FromTemplate,ProdStartDate from Offers with (NoLock) where OfferID=" & OfferID
  rst = MyCommon.LRT_Select
  If (rst.Rows.Count > 0) Then
    OfferStartDate = rst.Rows(0).Item("ProdStartDate")
    Name = rst.Rows(0).Item("Name")
    IsTemplate = MyCommon.NZ(rst.Rows(0).Item("IsTemplate"), False)
    If IsTemplate Then
      bUseTemplateLocks = False
    Else
      bUseTemplateLocks = MyCommon.NZ(rst.Rows(0).Item("FromTemplate"), False)
    End If
  End If

  If (IsTemplate Or bUseTemplateLocks) Then
    MyCommon.QueryStr = "select Disallow_Edit,DisallowEdit1,DisallowEdit2,DisallowEdit3,DisallowEdit4," & _
                        "DisallowEdit5,DisallowEdit6,DisallowEdit7,DisallowEdit8,DisallowEdit9 " & _
                        "from OfferRewards with (NoLock) where RewardID=" & RewardID
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      Disallow_Edit = MyCommon.NZ(rst.Rows(0).Item("Disallow_Edit"), True)
      bDisallowEditPg = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit1"), False)
      bDisallowEditDept = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit2"), False)
      bDisallowEditSpon = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit3"), False)
      bDisallowEditMsg = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit4"), False)
      bDisallowEditPp = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit5"), False)
      bDisallowEditDist = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit6"), False)
      bDisallowEditLimit = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit7"), False)
      bDisallowEditSpc = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit8"), False)
      bDisallowEditAdv = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit9"), False)
      If bUseTemplateLocks Then
        If Disallow_Edit Then
          bDisallowEditPg = True
          bDisallowEditDept = True
          bDisallowEditSpon = True
          bDisallowEditMsg = True
          bDisallowEditPp = True
          bDisallowEditDist = True
          bDisallowEditLimit = True
          bDisallowEditSpc = True
          bDisallowEditAdv = True
        Else
          If bDisallowEditDist Then
            bDisallowEditSpc = True
          End If
          Disallow_Edit = bDisallowEditPg And bDisallowEditDept And bDisallowEditSpon And bDisallowEditMsg And _
                          bDisallowEditPp And bDisallowEditDist And bDisallowEditLimit And bDisallowEditSpc And bDisallowEditAdv
        End If
      End If
    End If
  End If

  Dim bEnableBuckOffers As Boolean
  Dim bBuckChildOffer As Boolean
  Dim bBuckParentOffer As Boolean
  Dim oBuckStatus As Copient.ImportXml.BuckOfferStatus
  bEnableBuckOffers = (MyCommon.Fetch_CM_SystemOption(137) = "1")
  If bEnableBuckOffers Then
    oBuckStatus = MyImport.BuckOfferGetStatus(OfferID)
    Select Case oBuckStatus
      Case Copient.ImportXml.BuckOfferStatus.BuckParentNoChildren,
       Copient.ImportXml.BuckOfferStatus.BuckParentBothChildren,
       Copient.ImportXml.BuckOfferStatus.BuckParentPaperChildrenOnly,
       Copient.ImportXml.BuckOfferStatus.BuckParentDigitalChildrenOnly
        bBuckParentOffer = True
        bBuckChildOffer = False
      Case Copient.ImportXml.BuckOfferStatus.BuckChildPaper,
       Copient.ImportXml.BuckOfferStatus.BuckChildDigital
        bBuckParentOffer = False
        bBuckChildOffer = True
      Case Copient.ImportXml.BuckOfferStatus.BuckTiered
        bBuckParentOffer = True
        bBuckChildOffer = False
      Case Else
        bBuckParentOffer = False
        bBuckChildOffer = False
        If oBuckStatus = Copient.ImportXml.BuckOfferStatus.ErrorOccurred Then
          infoMessage = MyImport.GetErrorMsg()
        End If
    End Select
    If bBuckParentOffer Then
      bUseTemplateLocks = True
      Disallow_Edit = False
      bDisallowEditSpc = True
    End If
  Else
    bBuckParentOffer = False
    bBuckChildOffer = False
    oBuckStatus = Copient.ImportXml.BuckOfferStatus.NotBuckOffer
  End If

  If Request.QueryString("pgselectortype") Is Nothing OrElse Request.QueryString("pgselectortype") = "" Then
    PagePostBack = False
  Else
    If Request.QueryString("pgselectortype") = "directadd" Then
      ByExistingPGSelector = False
    End If
  End If
    
  If Request.QueryString("prodaddselector") = "prodlistadd" Then
    ByAddSingleProduct = False
  Else
    ByAddSingleProduct = True
  End If
  
  MyCommon.QueryStr = "select PG.Name, isnull(PG.IsStatic,0) as IsStatic from OfferRewards ORWD with (NoLock) Inner Join ProductGroups PG with (nolock) on ORWD.productgroupid=PG.productgroupid where ORWD.RewardID=" & RewardID & " and ORWD.deleted=0;"
  rst = MyCommon.LRT_Select
  
  If rst.Rows.Count > 0 Then
    GName = MyCommon.NZ(rst.Rows(0).Item("Name"), "")
    bStaticPG = MyCommon.NZ(rst.Rows(0).Item("IsStatic"), False)
  Else
    MyCommon.QueryStr = "select rewardid from offerrewards with (NoLock) where offerid=" & OfferID & " and rewardtypeid=1"
    rst3 = MyCommon.LRT_Select
    If rst3.Rows.Count = 1 Then
      GName = OfferStartDate.ToString("yyyyMMdd") & Name & "PGR"
    ElseIf rst3.Rows.Count > 1 Then
      GName = OfferStartDate.ToString("yyyyMMdd") & Name & "PGR(" & rst3.Rows.Count - 1 & ")"
    End If
    bStaticPG = False
  End If
  
  Send_HeadBegin("term.offer", "term.discountreward", OfferID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
%>
<script type="text/javascript">
var fullSelect = null;
var isFireFox = (navigator.appName.indexOf('Mozilla') !=-1) ? true : false;

/***************************************************************************************************************************/
//Script to call server method through JavaScript
//to load product based on search criteria.
var timer;

function xmlPostTimer(strURL,mode)
{
  clearTimeout(timer);
  timer=setTimeout("xmlhttpPost('" + strURL + "','" + mode + "')", 250);
}

function xmlhttpPost(strURL,mode) {
  var xmlHttpReq = false;
  var self = this;
  document.getElementById("searchLoadDiv").innerHTML = '<% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID)) %>';
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
      updatepage(self.xmlHttpReq.responseText);
    }
  }

  self.xmlHttpReq.send(qryStr);
  //self.xmlHttpReq.send(getquerystring());
}

function xmlhttpPostModifyPGroup(strURL, mode) {

        var xmlHttpReq = false;
        var self = this;

        handleWait(true);

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
        self.xmlHttpReq.send(getQueryString(mode));
        self.xmlHttpReq.onreadystatechange = function() {
            if (self.xmlHttpReq != null && self.xmlHttpReq.readyState == 4) {
                updatePageModPG(self.xmlHttpReq.responseText);
            }
        }

    }


    function updatePageModPG(responseMsg) {



        var elempasteproducts = document.getElementById("pasteproducts");

	var responseArray = new Array();
	responseArray=responseMsg.split('|');

        

        if (responseArray[0] != ""){      
           var div = document.createElement('div');
           div.innerHTML = responseArray[0];



           div.setAttribute('class', 'red-background'); // and make sure myclass has some styles in css
           document.getElementById('infobar1').appendChild(div);     
           
        }  
        else if (responseArray[1] != ""){
          document.getElementById("ProductGroupID").value = responseArray[1];
          elempasteproducts.value = elempasteproducts.defaultvalue;
          document.mainform.submit();           

        }
    handleWait(false);

    }

	 function handleWait(bShow) {
         var elem = document.getElementById("disabledBkgrd");

        if (elem != null) {
            elem.style.display = (bShow) ? 'block' : 'none';
        }
    }

   function getQueryString(mode) {

        var products = document.getElementById("pasteproducts").value;
        var GName = document.getElementById("modprodgroupname").value;
        var operationType = get_radio_value();
        var productType = 1;
        var RewardID = <%= RewardID%>;

        var bAllowHyphen  = '<% Sendb(MyCommon.Fetch_SystemOption(208))%>';
        if(bAllowHyphen == 1) {
            products = (products.toString().trim().replace(/\s/g, ', ')).replace(/-/g, '');
        } else {
            products = products.replace(/\r?\n/g, ', ');
        }
        return "Mode=" + mode + "&Products="+ products + "&OperationType="+ operationType + "&ProductType="+ productType + "&GName="+ GName + "&RewardID="+ RewardID + "&IsCondition="+ false;

    }

	function get_radio_value() {
            var inputs = document.getElementsByName("modifyoperation");
            for (var i = 0; i < inputs.length; i++) {
              if (inputs[i].checked) {
                return inputs[i].value;
              }
            }
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
  if(selected.options[0] != null){
    selectedGroup = selected.options[0].value;
  }
  var excluded = document.getElementById('excluded');
  var excludedGroup = 0;
  if(excluded.options[0] != null){
    excludedGroup = excluded.options[0].value;
  }
  return "Mode=" + mode + "&ProductSearch=" + document.getElementById('functioninput').value + "&OfferID=" + document.getElementById('OfferID').value + "&SelectedGroup=" + selectedGroup + "&ExcludedGroup=" + excludedGroup + "&SearchRadio=" + radioString;
 
}

function updatepage(str) {
  if(str.length > 0){
    if(!isFireFox){
      document.getElementById("pgList").innerHTML = '<select class="long" id="functionselect" name="functionselect" size="20"<% sendb(disabledattribute) %>>' + str + '</select>';
    }
    else{
      document.getElementById("functionselect").innerHTML = str;
    }
    document.getElementById("searchLoadDiv").innerHTML = '&nbsp;';
    if (document.getElementById("functionselect").options.length > 0) {
      document.getElementById("functionselect").options[0].selected = true;
    }
  }
  else if(str.length == 0){
    if(!isFireFox){
      document.getElementById("pgList").innerHTML = '<select class="long" id="functionselect" name="functionselect" size="20"<% sendb(disabledattribute) %>>&nbsp;</select>';
    }
    else{
      document.getElementById("functionselect").innerHTML = '&nbsp;';
    }
    document.getElementById("searchLoadDiv").innerHTML = '&nbsp;';
  }
}

function removeUsed(bSkipKeyUp)
{
    if (!bSkipKeyUp) { xmlhttpPost('OfferFeeds.aspx','ProductGroupsCM'); }
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
}

function handleSelectClick(itemSelected)
{
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
    if (selectedValue != ""){
      // empty the select box
      for (i = selectboxObj.length - 1; i>=0; i--) {
        selectboxObj.options[i] = null;
      }
      // add items to selected box
      selectboxObj[selectboxObj.length] = new Option(selectedText,selectedValue);
    }
  }
  
  if (itemSelected == "deselect1") {
    if (selectedboxValue != "") {
      // remove items from selected box
      document.getElementById("selected").remove(document.getElementById("selected").selectedIndex)
    }
  }
  
  if (itemSelected == "select2") {
    if (selectedValue != "" && selectedValue != "1") {
      for (i = excludedbox.length - 1; i>=0; i--) {
        excludedbox.options[i] = null;
      }
      // add item to excluded box
      excludedbox[excludedbox.length] = new Option(selectedText,selectedValue);
    } else if (selectedValue == "1") {
      alert('<%Sendb(Copient.PhraseLib.Lookup("term.anyproduct-not-excluded", LanguageID)) %>');
    }
  }
    
  if (itemSelected == "deselect2") {
    if (excludedboxValue != "") {
      // remove items from excluded box    
      document.getElementById("excluded").remove(document.getElementById("excluded").selectedIndex)
    }
  }
  
  updateButtons();
  
  // remove items from large list that are in the other lists
  removeUsed();
  return true;
}

function updateButtons(){
  var elemDisallowEditPgOpt = document.getElementById("DisallowEditPgOpt");
  var elemDisallowEditDistOpt = document.getElementById("DisallowEditDistOpt");
  var selectObj = document.getElementById('selected');
  var excludedObj = document.getElementById('excluded');
  
  if (elemDisallowEditPgOpt != null && elemDisallowEditPgOpt.value == '1') {
      document.getElementById('select1').disabled=true;
      document.getElementById('deselect1').disabled=true;
      document.getElementById('select2').disailed=true;
      document.getElementById('deselect2').disabled=true;
  } else {
    if (selectObj.length == 0) {
      if (excludedObj.length == 0) { 
        document.getElementById('select1').disabled=false;
        document.getElementById('deselect1').disabled=true;
        document.getElementById('select2').disabled=true;
        document.getElementById('deselect2').disabled=true;
      } else {
        // this state should not be allowed, but just in case
        document.getElementById('select1').disabled=true;
        document.getElementById('deselect1').disabled=true;
        document.getElementById('select2').disabled=true;
        document.getElementById('deselect2').disabled=false;
      }
    } else {
      if (excludedObj.length == 0) { 
        document.getElementById('select1').disabled=false;
        document.getElementById('deselect1').disabled=false;
        document.getElementById('select2').disabled=false;
        document.getElementById('deselect2').disabled=true;
      } else {
        document.getElementById('select1').disabled=false;
        document.getElementById('deselect1').disabled=true;
        document.getElementById('select2').disabled=false;
        document.getElementById('deselect2').disabled=false;
      }
    }
    if (elemDisallowEditDistOpt != null && elemDisallowEditDistOpt.value == '1') {
      if (selectObj.length == 0) {
        document.getElementById('select1').disabled=true;
        document.getElementById('deselect1').disabled=true;
        document.getElementById('select2').disabled=true;
        document.getElementById('deselect2').disabled=true;
      } else {
        if (excludedObj.length == 0) { 
          document.getElementById('select1').disabled=false;
          document.getElementById('deselect1').disabled=true;
          document.getElementById('select2').disabled=false;
          document.getElementById('deselect2').disabled=true;
        } else {
          document.getElementById('select1').disabled=false;
          document.getElementById('deselect1').disabled=true;
          document.getElementById('select2').disabled=false;
          document.getElementById('deselect2').disabled=false;
        }
      }
    }
  }
}

/***************************************************************************************************************************/

function checkDiscountTypeState() {
  var elemWebText=document.getElementById("divWebText");
  var elemExcluded=document.getElementById("excluded");
  var elemSelected=document.getElementById("selected");
  var elemDivAmount = document.getElementById("DivAmount");
  var elemNonTransOptions = document.getElementById("nontransoptions");
  var elemPrograms = document.getElementById("programs");
  var elemSpecialPrice = document.getElementById("specialprice");
  var elemDiscountType = document.getElementById("discounttype");
  var elemSpecialPricing = document.getElementById("specialpricing");
  var elemLimits = document.getElementById("limits");
  var elemUom = document.getElementById("WeightUom");
  var elemPselected = document.getElementById("Pselected");
  var elemStaticFuel = document.getElementById("staticfuel");
  var elemAdded=document.getElementById("PKID");
  var isTransactionLevel = false;
  var isStoredValue = false;
  var isAmountOffPerVolume = false;
  var isFreeItem = false;
  var isUom = false;
  var id, i;
  
  isTransactionLevel = ((elemExcluded != null && elemExcluded.options.length == 0) && (elemSelected != null && (elemSelected.options.length==0)) && (elemAdded != null && (elemAdded.options.length==0)));
  if (isTransactionLevel == true)
  {
    isStoredValue = false;
    isFreeItem = false;
    isUom = false;
    isAmountOffPerVolume = false;
  }
  else
  {
    isStoredValue = (elemDiscountType != null && (elemDiscountType.options.length > 7 && elemDiscountType.options[7].selected == true));
    isFreeItem = (elemDiscountType != null && (elemDiscountType.options.length > 6 && elemDiscountType.options[6].selected == true));
    isAmountOffPerVolume = (elemDiscountType != null && (elemDiscountType.options.length > 3 && elemDiscountType.options[3].selected == true));
    isUom = false;
    if (isStoredValue == true)
    {
      if ((elemPselected != null) && (elemPselected.options.length > 0))
      {
        id = elemPselected.options[0].value;
        for (i=0; i<Pvallist.length; i ++)
        {
          if (Pvallist[i] == id)
          {
            if (Ptypelist[i] == 3)
            {
              isUom = true;
            }
            break;
          }
        }
      }
    }
  }
  if (elemStaticFuel != null)
  {
    if (isAmountOffPerVolume == true)
    {
      elemStaticFuel.disabled=false;
    }
    else
    {
      elemStaticFuel.disabled=true;
    }
  }
  if (elemDivAmount != null)
  {
    if ((isStoredValue == true) || (isFreeItem == true))
    {
      elemDivAmount.style.display = "none";
    }
    else
    {
      elemDivAmount.style.display = "block";
    }
  }
  if (elemNonTransOptions != null)
  {
    if ((isStoredValue == true) || (isTransactionLevel == true))
    {
      elemNonTransOptions.style.display = "none";
      enableDistribution(false)
    }
    else
    {
      elemNonTransOptions.style.display = "block";
      enableDistribution(true)
    }
  }
  if (elemPrograms != null)
  {
    if (isStoredValue == true)
    {
      elemPrograms.style.display = "block";
    }
    else
    {
      elemPrograms.style.display = "none";
    }
  }
  if (elemSpecialPrice != null)
  {
    if (isStoredValue == true)
    {
      elemSpecialPrice.style.display = "none";
    }
    else
    {
      elemSpecialPrice.style.display = "block";
    }
  }
  if (elemSpecialPricing != null)
  {
    if ((isStoredValue == true) || (isTransactionLevel == true))
    {
      elemSpecialPricing.value = "off";
      elemSpecialPricing.checked = false;
    }
  }
  if (elemLimits != null)
  {
    if (isStoredValue == true)
    {
      elemLimits.style.display = "none";
    }
    else
    {
      elemLimits.style.display = "block";
    }
  }
  if (elemUom != null)
  {
    if (isUom == true)
    {
      elemUom.style.display = "block";
    }
    else
    {
      elemUom.style.display = "none";
    }
  }
  if (elemWebText != null)
  {
    if (isStoredValue == true)
    {
      elemWebText.style.display = "block";
    }
    else
    {
      elemWebText.style.display = "none";
    }
  }
  setLimitValueTypes(isTransactionLevel);
}

function setLimitValueTypes(isTransactionLevel) {
  var elemLimitValueType = document.getElementById("limitvaluetype");
  
  if (isTransactionLevel) {
    elemLimitValueType.options[0].selected = true;
    elemLimitValueType.options[1].style.display = "none";
    elemLimitValueType.options[2].style.display = "none";
  } else {
    elemLimitValueType.options[1].style.display = "block";
    elemLimitValueType.options[2].style.display = "block";
  }
}

function checkConditionState() {
  var elemExcluded=document.getElementById("excluded");
  var elemSelected=document.getElementById("selected");
  var elemAdded=document.getElementById("PKID");
  
  isTransactionLevel = ((elemExcluded != null && elemExcluded.options.length == 0) && (elemSelected != null && (elemSelected.options.length==0)) && (elemAdded != null && (elemAdded.options.length==0)));
  populateDiscountType(isTransactionLevel);
  showAdvancedOptions(isTransactionLevel);
  checkDiscountTypeState();
}

function enableDistribution(isEnabled) {
  var elemDisallowEditDistOpt = document.getElementById("DisallowEditDistOpt");
  var elemTriggerbogo=document.getElementById("triggerbogo");
  var elemXbox=document.getElementById("Xbox");
  var elemTriggerbxgy=document.getElementById("triggerbxgy");
  var elemBxgy1=document.getElementById("bxgy1");
  var elemBxgy2=document.getElementById("bxgy2");
  var elemTriggerprorate=document.getElementById("triggerprorate");
  var elemProrate=document.getElementById("prorate");
  var elemBuckParent=document.getElementById("BuckParent");
  var elemBestDeal=document.getElementById("bestdeal");
  var disableProrate=false;
  
  if (elemBuckParent != null && elemBuckParent.value == 'True'){
    if (elemTriggerbogo != null) { elemTriggerbogo.disabled = true; }
    if (elemXbox != null) { elemXbox.disabled = true; }
    if (elemTriggerbxgy != null) { elemTriggerbxgy.disabled = true; }
    if (elemBxgy1 != null) { elemBxgy1.disabled = true; }
    if (elemBxgy2 != null) { elemBxgy2.disabled = true; }
    if (elemTriggerprorate != null) { elemTriggerprorate.disabled = true; }
    if (elemProrate != null) { elemProrate.disabled = true; }
  } else if (elemDisallowEditDistOpt != null && elemDisallowEditDistOpt.value == '1') {
    if (elemTriggerbogo != null) { elemTriggerbogo.disabled = true; }
    if (elemXbox != null) { elemXbox.disabled = true; }
    if (elemTriggerbxgy != null) { elemTriggerbxgy.disabled = true; }
    if (elemBxgy1 != null) { elemBxgy1.disabled = true; }
    if (elemBxgy2 != null) { elemBxgy2.disabled = true; }
    if (elemTriggerprorate != null) { elemTriggerprorate.disabled = true; }
    if (elemProrate != null) { elemProrate.disabled = true; }
  } else {
    if (elemTriggerbogo != null) { elemTriggerbogo.disabled = (isEnabled) ? false : true; }
    if (elemXbox != null) { elemXbox.disabled = (isEnabled) ? false : true; }
    if (elemTriggerbxgy != null) { elemTriggerbxgy.disabled = (isEnabled) ? false : true; }
    if (elemBxgy1 != null) { elemBxgy1.disabled = (isEnabled) ? false : true; }
    if (elemBxgy2 != null) { elemBxgy2.disabled = (isEnabled) ? false : true; }
    if (elemBestDeal != null)
    {
      if (elemBestDeal.checked) { disableProrate = true; }
    }
    if (isEnabled == false) { disableProrate = true; }
    
    if (elemTriggerprorate != null) { elemTriggerprorate.disabled = disableProrate; }
    if (elemProrate != null) { elemProrate.disabled = disableProrate; }
  }
}
 
var prevSelectedVal = -1
var prevIsTransactionLevel = -1
 
function populateDiscountType(isTransactionLevel) {
  var elemDiscountType=document.getElementById("discounttype");
  var elemTiered=document.getElementById("tiered");
  var elemBuckParent=document.getElementById("BuckParent");
  var currSelectedVal = -1;
  var sValue = "";
  
  if (elemDiscountType != null) {
    currSelectedVal = elemDiscountType.options[elemDiscountType.selectedIndex].value;
    if (prevSelectedVal == -1) {
      prevSelectedVal = currSelectedVal;
    }

    // remove all the options from the discount type select box
    while (elemDiscountType.options.length > 0) {
      elemDiscountType.options[0] = null;
    }

    if (isTransactionLevel) {
      elemDiscountType.options[0] = new Option("<% Sendb(Copient.PhraseLib.Lookup("reward.discount-amountofftrans", LanguageID))%>", "8", false, false);  
      elemDiscountType.options[1] = new Option("<% Sendb(Copient.PhraseLib.Lookup("reward.discount-percentofftrans", LanguageID))%>", "9", false, false);  
    } else {
      elemDiscountType.options[0] = new Option("<% Sendb(Copient.PhraseLib.Lookup("reward.discount-amountoffitem", LanguageID))%>", "1", false, false);  
      elemDiscountType.options[1] = new Option("<% Sendb(Copient.PhraseLib.Lookup("reward.discount-percentoffitem", LanguageID))%>", "2", false, false);  
      if ((elemBuckParent == null) || (elemBuckParent != null && elemBuckParent.value == 'False')){
        elemDiscountType.options[2] = new Option("<% Sendb(Copient.PhraseLib.Lookup("reward.discount-amountoffweight", LanguageID))%>", "3", false, false);  
        elemDiscountType.options[3] = new Option("<% Sendb(Copient.PhraseLib.Lookup("reward.discount-amountoffvolume", LanguageID))%>", "4", false, false);  
        elemDiscountType.options[4] = new Option("<% Sendb(Copient.PhraseLib.Lookup("reward.discount-pricepoint", LanguageID))%>", "5", false, false);  
        elemDiscountType.options[5] = new Option("<% Sendb(Copient.PhraseLib.Lookup("reward.discount-pricepointweight", LanguageID))%>", "6", false, false);  
        elemDiscountType.options[6] = new Option("<% Sendb(Copient.PhraseLib.Lookup("reward.discount-freeitem", LanguageID))%>", "7", false, false);
        if ((elemTiered != null) && (elemTiered.value == "0")){
          elemDiscountType.options[7] = new Option("<% Sendb(Copient.PhraseLib.Lookup("term.storedvalue", LanguageID))%>", "10", false, false);
        }
      }
    }
    
    // select the previously selected discount type
    if (isTransactionLevel != prevIsTransactionLevel) {
      // select the previously selected discount type
      for (var i=0; i<elemDiscountType.options.length; i ++) {
        if (elemDiscountType.options[i].value == prevSelectedVal) {
          elemDiscountType.options[i].selected = true;
          break;
        }
      }
      prevSelectedVal = currSelectedVal;
      prevIsTransactionLevel = isTransactionLevel;
    } else {
      for (var i=0; i<elemDiscountType.options.length; i ++) {
        if (elemDiscountType.options[i].value == currSelectedVal) {
          elemDiscountType.options[i].selected = true;
          break;
        }
      }
    }
  }
}

function IsValidRegularExpression() 
	{    
	  if (isValidProductList() == true) {
	    var re = new RegExp("[^A-Za-z0-9/\r/\n,]");
        var bAllowSpacesTab  = '<% Sendb(MyCommon.Fetch_SystemOption(207))%>';
		var bAllowHyphen  = '<% Sendb(MyCommon.Fetch_SystemOption(208))%>';
		if(bAllowSpacesTab == 1) {
            re = new RegExp("[^A-Za-z0-9 /\r/\n/\t,]");
		} 
		if(bAllowHyphen == 1) {
            re = new RegExp("[^-A-Za-z0-9/\r/\n,]");
		}
		if(bAllowSpacesTab == 1 && bAllowHyphen == 1) {
            re = new RegExp("[^-A-Za-z0-9 /\r/\n/\t,]");
		}		
		if (document.getElementById("pasteproducts").value.match(re)) {
			alert('<% Sendb(Copient.PhraseLib.Lookup("pgroup-edit.invaliddata", LanguageID))%>');
			return false;
		} 
		else {    
		return true;
		}
	  }
     else {
	   return false;
	 }	  
		
		
	}

function showAdvancedOptions(isTransactionLevel) {
  var elemDisallowEditAdvOpt = document.getElementById("DisallowEditAdvOpt");
  var elemPromote = document.getElementById("promote");
  var elemDiscOnly = document.getElementById("disconly");
  var elemSameItem = document.getElementById("SameItem");
  var elemExclude = document.getElementById("exclude");
  var elemStaticFuel = document.getElementById("staticfuel");
  var elemBest = document.getElementById("bestdeal");
  var elemDecrement = document.getElementById("decrement");
  var elemVirtualLink = document.getElementById("virtuallink");
  
  if (elemDisallowEditAdvOpt != null && elemDisallowEditAdvOpt.value == '1') {
    if (elemPromote != null) {
      elemPromote.disabled = true;
    }
    if (elemDiscOnly != null) {
      elemDiscOnly.disabled = true;
    }
    if (elemVirtualLink != null) {
      elemVirtualLink.disabled = true;
    }
    if (elemExclude != null) {
      elemExclude.disabled = true;
    }
    if (elemStaticFuel != null) {
      elemStaticFuel.disabled = true;
    }
    if (elemBest != null) {
      elemBest.disabled = true;
    }
    if (elemDecrement != null) {
      elemDecrement.disabled = true;
    }
	if (elemSameItem != null) {
	  elemSameItem.disabled = true;
	}
  } else {
    if (elemPromote != null) {
      elemPromote.disabled = isTransactionLevel;
    }
    if (elemDiscOnly != null) {
      elemDiscOnly.disabled = isTransactionLevel;
      if (isTransactionLevel) {
        elemDiscOnly.checked = true;
      }
     }
	if (elemSameItem != null) {
      elemSameItem.disabled = isTransactionLevel;
    }
  }
 }
 
function validateNumeric(q) { 
var elemInvalidRewardAmt = document.getElementById("InvalidRewardAmt");
elemInvalidRewardAmt.value = false;
var elemFlagTier = document.getElementById("FlagTier");

  //for loop
  for (var i = 1; i <= q; i++) {
      var elemTier = document.getElementById("tier" + i);
      
      var value = elemTier.value;
      // isNaN() - true if the value is NaN (Not a Number), and false if not.
      // isFinite() - returns false if the value is +infinity, -infinity, or NaN.
      var y = !isNaN(parseFloat(value)) && isFinite(value); 
  
      if (y != true) { 
        elemTier.value = 0.000;
        elemInvalidRewardAmt.value = true;
        elemFlagTier.value = i;
      }
  }
}
// This is the javascript array holding the function list
// The PrintJavascriptArray ASP function can be used to print this array.
<%
    MyCommon.QueryStr = "select " & topString & " ProductGroupID,Name from ProductGroups with (NoLock) where Deleted=0 "
    If(bEnableRestrictedAccessToUEOfferBuilder) Then MyCommon.QueryStr &= " and isnull(TranslatedFromOfferID,0) = 0 "  
    MyCommon.QueryStr &= " order by AnyProduct desc, Name "
    rst = MyCommon.LRT_Select
    If (rst.rows.count>0)
        Sendb("var functionlist = Array(")
        For Each row In rst.Rows
            Sendb("""" & MyCommon.NZ(row.item("Name"), "").ToString().Replace("""", "\""") & """,")
        Next
        Send(""""");")
        Sendb("var vallist = Array(")
        For Each row In rst.Rows
            Sendb("""" & row.item("ProductGroupID") & """,")
        Next
        Send(""""");")
    Else
        Sendb("var functionlist = Array(")
        Send("""" & Copient.PhraseLib.Lookup("term.anyproduct", LanguageID) & """);")
        Sendb("var vallist = Array(")
        Send("""" & "1" & """);")
    End If
%>

// This is the function that refreshes the list after a keypress.
// The maximum number to show can be limited to improve performance with
// huge lists (1000s of entries).
// The function clears the list, and then does a linear search through the
// globally defined array and adds the matches back to the list.
function handleKeyUp(maxNumToShow)
{
  var selectObj, textObj, functionListLength;
  var i,  numShown;
  var searchPattern;
  
  document.getElementById("functionselect").size = "15";
  
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
  
  // Create a regular expression
  re = new RegExp(searchPattern,"gi");
  
  // Loop through the array and re-add matching options
  numShown = 0;
  if (textObj.value == '' && fullSelect != null) {
    var newSelectBox = fullSelect.cloneNode(true);
    document.getElementById('pgList').replaceChild(newSelectBox, selectObj);
  } else {
    var newSelectBox = selectObj.cloneNode(false);
    document.getElementById('pgList').replaceChild(newSelectBox, selectObj);
    selectObj = document.getElementById("functionselect");
    for(i = 0; i < functionListLength; i++) {
      if(functionlist[i].search(re) != -1) {
        if (vallist[i] != "") {
          selectObj[numShown] = new Option(functionlist[i], vallist[i]);
          if (vallist[i] == 1) {
            selectObj[numShown].style.fontWeight = 'bold';
            selectObj[numShown].style.color = 'brown';
          }
          numShown++;
        }
      }
      // Stop when the number to show is reached
      if(numShown == maxNumToShow) {
        break;
      }
    }
  }
  removeUsed(true);
  // When options list whittled to one, select that entry
  if(selectObj.length == 1) {
    selectObj.options[0].selected = true;
  }
}

// this function gets the selected value and loads the appropriate
// php reference page in the display frame
// it can be modified to perform whatever action is needed, or nothing
function handleSelectClick(itemSelected)
{
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
    
  var virtualLink=document.getElementById("virtualON").value;
  if (itemSelected == "select1") {
    if (selectboxObj.length == 0) {
      var discountcheck = document.getElementById("disconly");
      discountcheck.checked = false;
      if (virtualLink==1){
        var virtualcheck = document.getElementById("virtuallink");
        virtualcheck.checked = false;
        document.getElementById("virtuallink").disabled = false;
      }
    }
    if (selectedValue != ""){
      // empty the select box
	  if (selectboxObj.length == 0) {
		priorityReset();
		}
      for (i = selectboxObj.length - 1; i>=0; i--) {
        selectboxObj.options[i] = null;
      }
      // add items to selected box
      selectboxObj[selectboxObj.length] = new Option(selectedText,selectedValue);
    } 
  }
  
  if (itemSelected == "deselect1") {
    if (selectedboxValue != "") {
      // remove items from selected box
      document.getElementById("selected").remove(document.getElementById("selected").selectedIndex);
	  document.getElementById("PKID").options.length=0 ;
      if (virtualLink==1){
        var virtualcheck = document.getElementById("virtuallink");
        virtualcheck.checked = false;
        virtualcheck.disabled = true;
      }
      priorityReset();
    }
  }
  
  if (itemSelected == "select2") {
    if (selectedValue != "" && selectedValue != "1") {
      for (i = excludedbox.length - 1; i>=0; i--) {
        excludedbox.options[i] = null;
      }
      // add item to excluded box
      excludedbox[excludedbox.length] = new Option(selectedText,selectedValue);
    } else if (selectedValue == "1") {
      alert('<%Sendb(Copient.PhraseLib.Lookup("term.anyproduct-not-excluded", LanguageID)) %>');
    }
  }
    
  if (itemSelected == "deselect2") {
    if (excludedboxValue != "") {
      // remove items from excluded box    
      document.getElementById("excluded").remove(document.getElementById("excluded").selectedIndex)
    }
  }
  
  updateButtons();
  removeUsed();
  checkConditionState();

  return true;
}

function handleClickSent() {
  saveForm();
  document.mainform.clicksent.value='true';
  document.getElementById("specialpricing").value = document.mainform.usespecialpricing.checked? 'on':'off';
  document.mainform.submit();
}

function handleBestDealClickSent() {
  var elemTriggerProrate=document.getElementById("triggerprorate");
  var elemProrate=document.getElementById("prorate");
  var elemBestDeal=document.getElementById("bestdeal");

  if ((elemTriggerProrate != null) && (elemProrate != null)) {
    if (elemBestDeal != null) {
      if (elemBestDeal.checked) {
        elemTriggerProrate.disabled=true;
        elemProrate.disabled=true;
      }
      else {
        elemTriggerProrate.disabled=false;
        elemProrate.disabled=false;
      }
    }
  }
}

function handleProrateClicked() {
  var elemBestDeal=document.getElementById("bestdeal");
  if (elemBestDeal != null) {
    elemBestDeal.disabled=true;
  }
}

function handleProrateUnchecked() {
  var elemBestDeal=document.getElementById("bestdeal");
  if (elemBestDeal != null) {
    elemBestDeal.disabled=false;
  }
}

function ModifyGroup()
	{
	    var prods=document.getElementById("pasteproducts").value;

		if (IsValidRegularExpression())
		{
		if(prods!=null && prods!="" ){
			
				    var bCreateProducts  = '<% Sendb(MyCommon.Fetch_SystemOption(150))%>';
					if(document.getElementsByName("modifyoperation")[0].checked == true || document.getElementsByName("modifyoperation")[1].checked  == true)
					{
						if(bCreateProducts == 0)
						{
							alert('<%Sendb(Copient.PhraseLib.Lookup("pgroup-edit.productnotexist", LanguageID)) %>');
							return false;
						}
					}

					xmlhttpPostModifyPGroup('XMLFeeds.aspx', 'ModifyProducts');
				}
			
			else{
					alert('<% Sendb(Copient.PhraseLib.Lookup("pgroup-edit.invalidproducts", LanguageID))%>');
					return false;
				}
	}
		else{
					alert('<% Sendb(Copient.PhraseLib.Lookup("pgroup-edit.invaliddata", LanguageID))%>');
					return false;
			}
	}

function saveForm(){
    var funcSel = document.getElementById('functionselect');
    var exSel = document.getElementById('excluded');
    var elSel = document.getElementById('selected');
    var Pselected = document.getElementById('Pselected');
	var elenewprodgrpid = document.getElementById('NewCreatedProdGroupID');
    var i,j;
    var selectList = "";
    var excludededList = "";
    var htmlContents = "";

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
        }
    }
    
    if (Pselected != null && Pselected.options.length > 0) {
        document.getElementById("ProgramID").value = Pselected.options[0].value;
    }
    if (selectList != '') { 
  document.getElementById("ProductGroupID").value = selectList;     
  }

  else {    
    
     if (elenewprodgrpid != null){
	    document.getElementById("ProductGroupID").value = elenewprodgrpid.value;
	  }	  
  }
    document.getElementById("ExcludedID").value = excludededList;
        
    // alert(htmlContents);
    return true;
}

function handleCreateClick(createbtn)
    {
        var alertMessage="";
        if(document.getElementById(createbtn)!= undefined && document.getElementById(createbtn) != null)
        {
            var searchText= document.getElementById('functioninput').value;
            if(searchText != null && searchText!="")
            {
                if (searchText.toLowerCase() == '<% Sendb(Copient.PhraseLib.Lookup("term.anyproduct", LanguageID).ToLower())%>')
                {
                    alertMessage='<% Sendb(Copient.PhraseLib.Lookup("term.enter", LanguageID))%>' +' ' + '<% Sendb(Copient.PhraseLib.Lookup("term.valid", LanguageID).ToLower())%>'+' ' + '<% Sendb(Copient.PhraseLib.Lookup("term.productgroup", LanguageID).ToLower())%>'+' ' + '<% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID).ToLower())%>';
                    alert(alertMessage);
                    return false;
                }
                else
                {
                xmlhttpPost_CreateGroupOrProgramFromOffer('OfferFeeds.aspx', 'CreateGroupOrProgramFromOffer');
                }
            }
            else
            {
                alertMessage='<% Sendb(Copient.PhraseLib.Lookup("term.enter", LanguageID))%>' +' ' + '<% Sendb(Copient.PhraseLib.Lookup("term.valid", LanguageID).ToLower())%>'+' ' + '<% Sendb(Copient.PhraseLib.Lookup("term.productgroup", LanguageID).ToLower())%>'+' ' + '<% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID).ToLower())%>';
                alert(alertMessage);
                document.getElementById('functioninput').focus();
                return false;
            }
        }
        return true;
    }

 function xmlhttpPost_CreateGroupOrProgramFromOffer(strURL,mode) {
      var xmlHttpReq = false;
      var self = this;
      
      if (window.XMLHttpRequest) {
        self.xmlHttpReq = new XMLHttpRequest();
      }
      // IE
      else if (window.ActiveXObject) {
        self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
      }
      var qryStr = getcreatequery(mode);
      self.xmlHttpReq.open('POST', strURL + "?" + qryStr, true);
      self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
      self.xmlHttpReq.onreadystatechange = function() {
        if (self.xmlHttpReq !=null && self.xmlHttpReq.readyState == 4) {
          updatepage_creategroupOrprogramfromoffer(self.xmlHttpReq.responseText);
        }
      }

      self.xmlHttpReq.send(qryStr);
    }
    function getcreatequery(mode)
    {
        return "Mode=" + mode + "&CreateType=ProductGroup&Name=" + document.getElementById('functioninput').value
    }
    function updatepage_creategroupOrprogramfromoffer(str) 
    {
       if(str.length > 0)
      {
          var status ="";
          var responseArr = str.split('~');
          if(responseArr.length >0)
          {
            status=responseArr[0];
            if(status=="Ok")
            {
                var resultArr=responseArr[1].split('|');
                 if(document.getElementById('selected').options[0] !=undefined && document.getElementById('selected').options[0]!=null){
                    document.getElementById('selected').options[0].selected = 'selected';
                }
                handleSelectClick('deselect1');
                addNewGrouptoSelectbox(resultArr[0],resultArr[1]);
                document.getElementById('functionselect').value=resultArr[1];
                handleSelectClick('select1') ;
            }
            else if(status =="Fail")
            {
                var resultArr=responseArr[1].split('|');
                var selectedGroupValue= -1;
                if(document.getElementById('selected').options[0] !=undefined && document.getElementById('selected').options[0]!=null)
                    selectedGroupValue= document.getElementById('selected').options[0].value
                if(parseInt(selectedGroupValue) != parseInt(resultArr[1] ))
                {
                    alert(responseArr[2]);
                    if(selectedGroupValue != -1){
                        document.getElementById('selected').options[0].selected = 'selected';
                    }
                    handleSelectClick('deselect1');
                    document.getElementById('functionselect').value=resultArr[1];
                    handleSelectClick('select1') ;
                }
                else
                {
                    alert('<% Sendb(Copient.PhraseLib.Lookup("term.productgroup", LanguageID))%>' + ": '"+ resultArr[0] + "' " + '<% Sendb(Copient.PhraseLib.Lookup("term.is", LanguageID).ToLower())%>'+'  '+'<% Sendb(Copient.PhraseLib.Lookup("term.already", LanguageID).ToLower())%>'+'  '+'<% Sendb(Copient.PhraseLib.Lookup("term.selected", LanguageID).ToLower())%>');
                }
            }
            else if(status =="Error")
            {
                alert(responseArr[1]);
                return false;
            }
         }
      }
       
   }

   function addNewGrouptoSelectbox(text,val)
   {
        var sel = document.getElementById('functionselect');
        var opt = document.createElement('option'); // create new option element
        // create text node to add to option element (opt)
        opt.appendChild( document.createTextNode(text) );
        opt.value = val; // set value property of opt
        sel.appendChild(opt); // add opt to end of select box (sel)
        functionlist.push(text);
        vallist.push(val);
   }
</script>
<script type="text/javascript">
// This is the javascript array holding the function list
// The PrintJavascriptArray ASP function can be used to print this array.
<%
    MyCommon.QueryStr = "select SVProgramID as ProgramID,Name as ProgramName, SVTypeId from StoredValuePrograms with (NoLock)" & _
                        " where deleted=0 and Visible=1 and SVProgramID is not null and SVTypeId > 1 order by ProgramName"
    rst2 = MyCommon.LRT_Select
    
    If (rst2.rows.count>0)
        Sendb("var Pfunctionlist = Array(")
        For Each row2 In rst2.Rows
            Sendb("""" & MyCommon.NZ(row2.item("ProgramName"), "").ToString().Replace("""", "\""") & """,")
        Next
        Send(""""");")
        Sendb("var Pvallist = Array(")
        For Each row2 In rst2.Rows
            Sendb("""" & row2.item("ProgramID") & """,")
        Next
        Send(""""");")
        Sendb("var Ptypelist = Array(")
        For Each row2 In rst2.Rows
            Sendb("""" & row2.item("SVTypeId") & """,")
        Next
        Send(""""");")
    Else
        Send("var Pfunctionlist = Array();")
        Send("var Pvallist = Array();")
        Send("var Ptypelist = Array();")
    End If
%>

// This is the function that refreshes the list after a keypress.
// The maximum number to show can be limited to improve performance with
// huge lists (1000s of entries).
// The function clears the list, and then does a linear search through the
// globally defined array and adds the matches back to the list.
function PhandleKeyUp(maxNumToShow)
{
    var selectObj, textObj, PfunctionListLength;
    var i, numShown;
    var searchPattern;
    var selectedList;
    
    document.getElementById("Pfunctionselect").size = "10";
    
    // Set references to the form elements
    selectObj = document.forms[0].Pfunctionselect;
    textObj = document.forms[0].Pfunctioninput;
    selectedList = document.getElementById("Pselected");

    // Remember the function list length for loop speedup
    PfunctionListLength = Pfunctionlist.length;
    
    // Set the search pattern depending
    if(document.forms[0].Pfunctionradio[0].checked == true)
    {
        searchPattern = "^"+textObj.value;
    }
    else
    {
        searchPattern = textObj.value;
    }
    searchPattern = cleanRegExpString(searchPattern);

    // Create a regulare expression

    re = new RegExp(searchPattern,"gi");
    // Clear the options list
    selectObj.length = 0;

    // Loop through the array and re-add matching options
    numShown = 0;
    for(i = 0; i < PfunctionListLength; i++)
    {
        if(Pfunctionlist[i].search(re) != -1)
        {
            if (Pvallist[i] != "" && (selectedList.options.length < 1 || Pvallist[i] != selectedList.options[0].value) ) {
                selectObj[numShown] = new Option(Pfunctionlist[i],Pvallist[i]);
                numShown++;
            }
        }
        // Stop when the number to show is reached
        if(numShown == maxNumToShow)
        {
            break;
        }
    }
    // When options list whittled to one, select that entry
    if(selectObj.length == 1)
    {
        selectObj.options[0].selected = true;
    }
}

function PremoveUsed()
{
    PhandleKeyUp(99999);
    // this function will remove items from the functionselect box that are used in 
    // selected and excluded boxes

    var funcSel = document.getElementById('Pfunctionselect');
    var elSel = document.getElementById('Pselected');
    var i,j;
  
    for (i = elSel.length - 1; i>=0; i--) {
        for(j=funcSel.length-1;j>=0; j--) {
            if(funcSel.options[j].value == elSel.options[i].value){
                funcSel.options[j] = null;
            }
        }
    }
}


// this function gets the selected value and loads the appropriate
// php reference page in the display frame
// it can be modified to perform whatever action is needed, or nothing
function PhandleSelectClick(itemSelected)
{
  textObj = document.forms[0].Pfunctioninput;
     
  selectObj = document.forms[0].Pfunctionselect;
  selectedValue = document.getElementById("Pfunctionselect").value;
  if (selectedValue != ""){ selectedText = selectObj[document.getElementById("Pfunctionselect").selectedIndex].text; }
    
  selectboxObj = document.forms[0].Pselected;
  selectedboxValue = document.getElementById("Pselected").value;
  if (selectedboxValue != ""){ selectedboxText = selectboxObj[document.getElementById("Pselected").selectedIndex].text; }
    
  if (itemSelected == "Pselect1") {
    if (selectedValue != ""){
      // empty the select box
      for (i = selectboxObj.length - 1; i>=0; i--) {
        selectboxObj.options[i] = null;
      }
      // add items to selected box
      selectboxObj[selectboxObj.length] = new Option(selectedText,selectedValue);
    }
  }
    
  if (itemSelected == "Pdeselect1") {
    if (selectedboxValue != ""){
      // remove items from selected box
      document.getElementById("Pselected").remove(document.getElementById("Pselected").selectedIndex)
    }
  }
    
  PupdateButtons();
  // remove items from large list that are in the other lists
  PremoveUsed();
  checkConditionState();
  return true;
}

function PupdateButtons(){
  var elemDisallowEditPpOpt = document.getElementById("DisallowEditPpOpt");
  var selectObj = document.getElementById('Pselected');
  
  if (elemDisallowEditPpOpt != null && elemDisallowEditPpOpt.value == '1') {
      document.getElementById('Pselect1').disabled=true;
      document.getElementById('Pdeselect1').disabled=true;
  } else {
    if (selectObj.length == 0) {
      document.getElementById('Pselect1').disabled=false;
      document.getElementById('Pdeselect1').disabled=true;
    } else {
      document.getElementById('Pselect1').disabled=false;
      document.getElementById('Pdeselect1').disabled=false;
    }
  }
}

// This is the javascript array holding the function list
// The PrintJavascriptArray ASP function can be used to print this array.
<%
    MyCommon.QueryStr = "select 0 as LimitID, 'None' as Name, RewardLimitTypeID as LimitTypeID, RewardLimit as LimitValue, RewardDistPeriod as LimitPeriod " & _
                        "from OfferRewards with (NoLock) where RewardID=" & RewardID & " " & _
                        "union " & _
                        "select LimitID, Name, LimitTypeID, LimitValue, LimitPeriod " & _
                        "from CM_AdvancedLimits with (NoLock) where Deleted=0 and LimitTypeID in (2,3,4) order By LimitID;"
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

        Sendb("var ALvallist4 = Array(")
        For Each row2 In rst2.Rows
            Sendb("""" & row2.item("LimitTypeID") & """,")
        Next
        Send(""""");")
    End If
%>

function setlimitsection(bSelect) {
  var elemSelectAdv = document.getElementById("selectadv");
  var elemSelectDay=document.getElementById("selectday");
  var elemPeriod=document.getElementById("limitperiod");
  var elemValue=document.getElementById("limitvalue");
  var elemValueType=document.getElementById("limitvaluetype");
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
        if (elemValueType != null) {
          elemValueType.disabled = false;
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
        if (elemValueType != null) {
          elemValueType.disabled = true;
        }
      }
    }
    
    for(i = 0; i < ALfunctionlist.length; i++)
    {
      if(elemSelectAdv.value == ALvallist1[i])
      {
        elemPeriod.value = ALvallist2[i];
        elemValue.value = ALvallist3[i];
        elemValueType.value = ALvallist4[i];
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
}
</script>
<%
  Send_Scripts()
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
  Send("<div id=""disabledBkgrd"" style=""position:absolute;top:0px;left:0px;right:0px;width:100%;height:100%;background-color:Gray;display:none;z-index:99;filter:alpha(opacity=50);-moz-opacity:.50;opacity:.50;""></div>")
  ' we need to determine our linkid for updates and tiered
  MyCommon.QueryStr = "select LinkID,Tiered,SponsorID,RewardDistPeriod,PromoteToTransLevel,RewardLimit,RewardLimitTypeID,TriggerQty,RewardAmountTypeID, " & _
                      "UseSpecialPricing, SPRepeatAtOccur,ApplyToLimit,DoNotItemDistribute,AdvancedLimitID, SameItemReward, VirtualLink from OfferRewards with (NoLock) where RewardID=" & RewardID
  rst = MyCommon.LRT_Select
  For Each row In rst.Rows
    DistPeriod = MyCommon.NZ(row.Item("RewardDistPeriod"), 0)
    LinkID = row.Item("LinkID")
    Tiered = MyCommon.NZ(row.Item("Tiered"), 0)
    SponsorID = MyCommon.NZ(row.Item("SponsorID"), 0)
    PromoteToTransLevel = MyCommon.NZ(row.Item("PromoteToTransLevel"), 0)
    SameItem = MyCommon.NZ(row.Item("SameItemReward"), 0)
    RewardLimit = MyCommon.NZ(row.Item("RewardLimit"), 0)
    RewardLimitTypeID = MyCommon.NZ(row.Item("RewardLimitTypeID"), 2)
    RewardAmountTypeID = MyCommon.NZ(row.Item("RewardAmountTypeID"), 1)
    AdvancedLimitID = MyCommon.NZ(row.Item("AdvancedLimitID"), 0)
    TriggerQty = MyCommon.NZ(row.Item("TriggerQty"), 1)
    ApplyToLimit = MyCommon.NZ(row.Item("ApplyToLimit"), 1)
    UseSpecialPricing = MyCommon.NZ(row.Item("UseSpecialPricing"), 0)
    SPRepeatAtOccur = MyCommon.NZ(row.Item("SPRepeatAtOccur"), 1)
    DoNotItemDistribute = MyCommon.NZ(row.Item("DoNotItemDistribute"), 0)
    VirtualLink=MyCommon.NZ(row.Item("VirtualLink"),0)
  Next
  
  If (ApplyToLimit = 0) Then
    ValueRadio = 3
  ElseIf (TriggerQty = ApplyToLimit And TriggerQty <> 0) Then
    ValueRadio = 1
  Else
    ValueRadio = 2
  End If
  
  'If (DoNotItemDistribute = True And TriggerQty = ApplyToLimit) Then
  '  ValueRadio = 1
  'ElseIf (DoNotItemDistribute = True And TriggerQty <> ApplyToLimit And ApplyToLimit <> 0) Then
  '  ValueRadio = 2
  'ElseIf (DoNotItemDistribute = False And ApplyToLimit = 0) Then
  '  ValueRadio = 3
  'End If
  
  'Response.Write(DoNotItemDistribute)
  ' items from Discouts table we need to fill into forms
  MyCommon.QueryStr = "select AllowNegative,ChargeBackDeptID,EffectMinOrder,BestDeal,StaticFuel,DiscountableITemsOnly,PrintLineText,SVLinkID,Webtext from Discounts with (NoLock) where DiscountID=" & LinkID
  rst = MyCommon.LRT_Select()
  For Each row In rst.Rows
    AllowNegative = MyCommon.NZ(row.Item("AllowNegative"), 0)
    ChargeBackDeptID = MyCommon.NZ(row.Item("ChargeBackDeptID"), 0)
    BestDeal = MyCommon.NZ(row.Item("BestDeal"), 0)
    StaticFuel = MyCommon.NZ(row.Item("StaticFuel"), 0)
    DiscountableItemsOnly = MyCommon.NZ(row.Item("DiscountableItemsOnly"), 0)
    PrintLineText = MyCommon.NZ(row.Item("PrintLineText"), "")
    EffectMinOrder = MyCommon.NZ(row.Item("EffectMinOrder"), 0)
    SVLinkID = MyCommon.NZ(row.Item("SVLinkID"), 0)
    WebText = MyCommon.NZ(row.Item("WebText"), "")
  Next
  
  ProgramID = 0
  If RewardAmountTypeID = 10 Then
    MyCommon.QueryStr = "select RewardStoredValuesID,ProgramID, SVTypeID, WeightUOM from CM_RewardStoredValues where RewardStoredValuesID=" & SVLinkID
    rst = MyCommon.LRT_Select()
    If rst.Rows.Count > 0 Then
      ProgramID = MyCommon.NZ(rst.Rows(0).Item("ProgramID"), 0)
      SVType = MyCommon.NZ(rst.Rows(0).Item("SVTypeID"), 0)
      WeightUom = MyCommon.NZ(rst.Rows(0).Item("WeightUOM"), 0)
    End If
  Else
    SVType = 0
  End If
  
 
  If (Request.QueryString("pgroup-add1") <> "" And Request.QueryString("pgroup-avail") <> "") Then
    MyCommon.QueryStr = "update OfferRewards with (RowLock) set ProductGroupID=0 where RewardID=" & RewardID & ";"
    MyCommon.LRT_Execute()
    MyCommon.QueryStr = "update OfferRewards with (RowLock) set ProductGroupID=" & Request.QueryString("pgroup-avail") & " where RewardID=" & RewardID & ";"
    MyCommon.LRT_Execute()
  ElseIf (Request.QueryString("pgroup-rem1") <> "") Then
    MyCommon.QueryStr = "update OfferRewards with (RowLock) set ProductGroupID=0 where RewardID=" & RewardID & ";"
    MyCommon.LRT_Execute()
    MyCommon.QueryStr = "update offerrewards with (RowLock) set rewardamounttypeid=1 where rewardid=" & RewardID & ";"
    MyCommon.LRT_Execute()
    RewardAmountTypeID = 1
  ElseIf (Request.QueryString("pgroup-add2") <> "" And Request.QueryString("pgroup-avail") <> "" And Request.QueryString("pgroup-avail") <> "1") Then
    MyCommon.QueryStr = "update OfferRewards with (RowLock) set ExcludedProdGroupID=0 where RewardID=" & RewardID & ";"
    MyCommon.LRT_Execute()
    MyCommon.QueryStr = "update OfferRewards with (RowLock) set ExcludedProdGroupID=" & Request.QueryString("pgroup-avail") & " where RewardID=" & RewardID & ";"
    MyCommon.LRT_Execute()
  ElseIf (Request.QueryString("pgroup-rem2") <> "") Then
    MyCommon.QueryStr = "update OfferRewards with (RowLock) set ExcludedProdGroupID=0 where RewardID=" & RewardID & ";"
    MyCommon.LRT_Execute()
  End If
  
  If (Request.QueryString("addvalue") <> "") Then
    MyCommon.QueryStr = "select max(TierLevel) as maxtier from RewardTiers with (NoLock) where RewardID=" & RewardID
    rst = MyCommon.LRT_Select
    ' ok we know the highest tier now so we need to add one
    'dbo.pt_ConditionTiers_Update @ConditionID bigint, @TierLevel int,@AmtRequired decimal(12,3)
    MyCommon.QueryStr = "dbo.pt_RewardTiers_Update"
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add("@RewardID", SqlDbType.BigInt).Value = RewardID
    MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int).Value = MyCommon.Extract_Val(MyCommon.NZ(rst.Rows(0).Item("maxtier"), 0)) + 1
    MyCommon.LRTsp.Parameters.Add("@RewardAmount", SqlDbType.Decimal).Value = 0
    MyCommon.LRTsp.ExecuteNonQuery()
    MyCommon.Close_LRTsp()
    End If
    If (Request.QueryString("deletespecial") <> "") Then
    ' we need to delete the last special pricing level
    MyCommon.QueryStr = "select max(TierLevel) as maxtier from RewardTiers with (NoLock) where RewardID=" & RewardID
    rst = MyCommon.LRT_Select
    ' ok we know the max so lets get rid of it
    MyCommon.QueryStr = "delete from RewardTiers with (RowLock) where TierLevel = " & rst.Rows(0).Item("maxtier") & " and RewardID=" & RewardID
    MyCommon.LRT_Execute()
  ElseIf (Request.QueryString("save") <> "" Or Request.QueryString("clicksent") = "true" Or _
    Request.QueryString("pgroup-add1") <> "" Or _
    Request.QueryString("pgroup-rem1") <> "" Or _
    Request.QueryString("pgroup-add2") <> "" Or _
      Request.QueryString("addvalue") <> "" Or _
    Request.QueryString("pgroup-rem2") <> "") Then
    ' minorderitems nodistribute
    If (Request.QueryString("IsTemplate") = "IsTemplate") Then
      ' time to update the status bits for the templates
      Dim form_Disallow_Edit As Integer = 0
      Dim iDisallowEditPg As Integer = 0
      Dim iDisallowEditDept As Integer = 0
      Dim iDisallowEditSpon As Integer = 0
      Dim iDisallowEditMsg As Integer = 0
      Dim iDisallowEditPp As Integer = 0
      Dim iDisallowEditDist As Integer = 0
      Dim iDisallowEditLimit As Integer = 0
      Dim iDisallowEditSpc As Integer = 0
      Dim iDisallowEditAdv As Integer = 0
      
      Disallow_Edit = False
      bDisallowEditPg = False
      bDisallowEditDept = False
      bDisallowEditSpon = False
      bDisallowEditMsg = False
      bDisallowEditPp = False
      bDisallowEditDist = False
      bDisallowEditLimit = False
      bDisallowEditSpc = False
      bDisallowEditAdv = False
      
      If (Request.QueryString("Disallow_Edit") = "on") Then
        form_Disallow_Edit = 1
        Disallow_Edit = True
      End If
      If (Request.QueryString("DisallowEditPg") = "on") Then
        iDisallowEditPg = 1
        bDisallowEditPg = True
      End If
      If (Request.QueryString("DisallowEditDept") = "on") Then
        iDisallowEditDept = 1
        bDisallowEditDept = True
      End If
      If (Request.QueryString("DisallowEditSpon") = "on") Then
        iDisallowEditSpon = 1
        bDisallowEditSpon = True
      End If
      If (Request.QueryString("DisallowEditMsg") = "on") Then
        iDisallowEditMsg = 1
        bDisallowEditMsg = True
      End If
      If (Request.QueryString("DisallowEditPp") = "on") Then
        iDisallowEditPp = 1
        bDisallowEditPp = True
      End If
      If (Request.QueryString("DisallowEditDist") = "on") Then
        iDisallowEditDist = 1
        bDisallowEditDist = True
      End If
      If (Request.QueryString("DisallowEditLimit") = "on") Then
        iDisallowEditLimit = 1
        bDisallowEditLimit = True
      End If
      If (Request.QueryString("DisallowEditSpc") = "on") Then
        iDisallowEditSpc = 1
        bDisallowEditSpc = True
      End If
      If (Request.QueryString("DisallowEditAdv") = "on") Then
        iDisallowEditAdv = 1
        bDisallowEditAdv = True
      End If
      
      ' auto lock hidden frames
      If (Request.QueryString("discounttype") <> "") Then
        RewardAmountTypeID = MyCommon.Extract_Val(Request.QueryString("discounttype"))
        If RewardAmountTypeID = 10 Then
          ' stored value type
          iDisallowEditLimit = 1
          bDisallowEditLimit = True
          iDisallowEditSpc = 1
          bDisallowEditSpc = True
        Else
          iDisallowEditPp = 1
          bDisallowEditPp = True
        End If
      End If
      
      MyCommon.QueryStr = "update OfferRewards with (RowLock) set Disallow_Edit=" & form_Disallow_Edit & _
        ",DisallowEdit1=" & iDisallowEditPg & _
        ",DisallowEdit2=" & iDisallowEditDept & _
        ",DisallowEdit3=" & iDisallowEditSpon & _
        ",DisallowEdit4=" & iDisallowEditMsg & _
        ",DisallowEdit5=" & iDisallowEditPp & _
        ",DisallowEdit6=" & iDisallowEditDist & _
        ",DisallowEdit7=" & iDisallowEditLimit & _
        ",DisallowEdit8=" & iDisallowEditSpc & _
        ",DisallowEdit9=" & iDisallowEditAdv & _
        " where RewardID=" & RewardID
      MyCommon.LRT_Execute()
    End If
    
    
    If Not (bUseTemplateLocks And bDisallowEditPg) Then
      Select Case ProductGroupID
        Case 0
          MyCommon.QueryStr = "update OfferRewards with (RowLock) set ProductGroupID=0, ExcludedProdGroupID=0 where RewardID=" & RewardID & " and deleted=0;"
          MyCommon.LRT_Execute()
        Case 1
          MyCommon.QueryStr = "update OfferRewards with (RowLock) set ProductGroupID=" & ProductGroupID & ", ExcludedProdGroupID=" & ExcludedID & " where RewardID=" & RewardID & " and deleted=0;"
          MyCommon.LRT_Execute()
        Case Else
          If (ExcludedID > 0) Then
            MyCommon.QueryStr = "update OfferRewards with (RowLock) set ProductGroupID=" & ProductGroupID & ", ExcludedProdGroupID=" & ExcludedID & " where RewardID=" & RewardID & " and deleted=0;"
          Else
            MyCommon.QueryStr = "update OfferRewards with (RowLock) set ProductGroupID=" & ProductGroupID & ", ExcludedProdGroupID=0 where RewardID=" & RewardID & " and deleted=0;"
          End If
          MyCommon.LRT_Execute()
      End Select
    End If

    If Not (bUseTemplateLocks And bDisallowEditAdv) Then
      If (Request.QueryString("promote") = "on") Then
        MyCommon.QueryStr = "update OfferRewards with (RowLock) set PromoteToTransLevel=1 where RewardID=" & RewardID
        MyCommon.LRT_Execute()
        PromoteToTransLevel = True
      Else
        MyCommon.QueryStr = "update OfferRewards with (RowLock) set PromoteToTransLevel=0 where RewardID=" & RewardID
        MyCommon.LRT_Execute()
        PromoteToTransLevel = False
      End If
      If (Request.QueryString("exclude") = "on") Then
        MyCommon.QueryStr = "update OfferRewards with (RowLock) set DoNotItemDistribute=1 where RewardID=" & RewardID
        MyCommon.LRT_Execute()
        DoNotItemDistribute = True
      Else
        MyCommon.QueryStr = "update OfferRewards with (RowLock) set DoNotItemDistribute=0 where RewardID=" & RewardID
        MyCommon.LRT_Execute()
        DoNotItemDistribute = False
      End If
      If (Request.QueryString("disc-items-only") = "on") Then
        MyCommon.QueryStr = "update discounts with (RowLock) set DiscountableItemsOnly=1 where DiscountID=" & LinkID
        MyCommon.LRT_Execute()
        DiscountableItemsOnly = True
      Else
        MyCommon.QueryStr = "update discounts with (RowLock) set DiscountableItemsOnly=0 where DiscountID=" & LinkID
        MyCommon.LRT_Execute()
        DiscountableItemsOnly = False
      End If
      If (Request.QueryString("bestdeal") = "on") Then
        MyCommon.QueryStr = "update discounts with (RowLock) set BestDeal=1 where DiscountID=" & LinkID
        MyCommon.LRT_Execute()
        BestDeal = True
      Else
        MyCommon.QueryStr = "update discounts with (RowLock) set BestDeal=0 where DiscountID=" & LinkID
        MyCommon.LRT_Execute()
        BestDeal = False
      End If
      If (Request.QueryString("staticfuel") = "on") Then
        MyCommon.QueryStr = "update discounts with (RowLock) set StaticFuel=1 where DiscountID=" & LinkID
        MyCommon.LRT_Execute()
        StaticFuel = True
      Else
        MyCommon.QueryStr = "update discounts with (RowLock) set StaticFuel=0 where DiscountID=" & LinkID
        MyCommon.LRT_Execute()
        StaticFuel = False
      End If
      If (Request.QueryString("decrement") = "on") Then
        MyCommon.QueryStr = "update discounts with (RowLock) set EffectMinOrder=1 where DiscountID=" & LinkID
        MyCommon.LRT_Execute()
        EffectMinOrder = True
      Else
        MyCommon.QueryStr = "update discounts with (RowLock) set EffectMinOrder=0 where DiscountID=" & LinkID
        MyCommon.LRT_Execute()
        EffectMinOrder = False
      End If
      If (Request.QueryString("SameItem") = "on") Then
        MyCommon.QueryStr = "update OfferRewards with (RowLock) set SameItemReward=1 where RewardID=" & RewardID
        MyCommon.LRT_Execute()
        SameItem = True
      Else
        MyCommon.QueryStr = "update OfferRewards with (RowLock) set SameItemReward=0 where RewardID=" & RewardID
        MyCommon.LRT_Execute()
        SameItem = False
      End If
      If MyCommon.Fetch_CM_SystemOption(136) Then
        If (Request.QueryString("virtuallink") = "on") Then
          MyCommon.QueryStr = "update OfferRewards with (RowLock) set VirtualLink=1 where RewardID=" & RewardID
          MyCommon.LRT_Execute()
          EffectMinOrder = True
        Else
          MyCommon.QueryStr = "update OfferRewards with (RowLock) set VirtualLink=0 where RewardID=" & RewardID
          MyCommon.LRT_Execute()
          EffectMinOrder = False
        End If
      End If
    End If

    
    If Not (bUseTemplateLocks And bDisallowEditDept) Then
      If (Request.QueryString("dept") <> "") Then
        ChargeBackDeptID = Request.QueryString("dept")
        MyCommon.QueryStr = "update Discounts with (RowLock) set ChargeBackDeptID=" & ChargeBackDeptID & " where DiscountID=" & LinkID
        MyCommon.LRT_Execute()
      End If
    End If
    
    If Not (bUseTemplateLocks And bDisallowEditSpon) Then
      If (Request.QueryString("sponsor") <> "") Then
        SponsorID = MyCommon.Extract_Val(Request.QueryString("sponsor"))
        MyCommon.QueryStr = "update OfferRewards with (RowLock) set SponsorID=" & SponsorID & " where RewardID=" & RewardID
        MyCommon.LRT_Execute()
      End If
    End If

    If Not (bUseTemplateLocks And bDisallowEditMsg) Then
      If (GetCgiValue("PrintLineText") <> "") Then
        PrintLineText = GetCgiValue("PrintLineText")
        MyCommon.QueryStr = "UPDATE discounts with (RowLock) set PrintLineText=@PrintLineText where DiscountID=@DiscountID"
        MyCommon.DBParameters.Add("@PrintLineText", SqlDbType.NVarChar, 1000).Value = PrintLineText
        MyCommon.DBParameters.Add("@DiscountID", SqlDbType.BigInt).Value = LinkID
        MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
      Else
        MyCommon.QueryStr = "update discounts with (RowLock) set PrintLineText=null where DiscountID=" & LinkID
        MyCommon.LRT_Execute()
      End If
      If (GetCgiValue("WebText") <> "") Then
        WebText = GetCgiValue("WebText")
        MyCommon.QueryStr = "update discounts with (RowLock) set WebText=@WebText where DiscountID=@DiscountID"
        MyCommon.DBParameters.Add("@WebText", SqlDbType.NVarChar, 1000).Value = WebText
        MyCommon.DBParameters.Add("@DiscountID", SqlDbType.BigInt).Value = LinkID
        MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
      Else
        MyCommon.QueryStr = "update discounts with (RowLock) set WebText=null where DiscountID=" & LinkID
        MyCommon.LRT_Execute()
      End If
    End If

    If Not (bUseTemplateLocks And bDisallowEditDist) Then
      If (Request.QueryString("discounttype") <> "") Then
        RewardAmountTypeID = MyCommon.Extract_Val(Request.QueryString("discounttype"))
        MyCommon.QueryStr = "update offerRewards with (RowLock) set RewardAmountTypeID=" & RewardAmountTypeID & " where RewardID=" & RewardID
        MyCommon.LRT_Execute()
      End If
      If (MyCommon.Fetch_CM_SystemOption(117) = "1") Then
        If (Request.QueryString("PriorityReset") = "True") Then
					
          Dim defaultPriority As Integer
          Dim defaultECPriority As Integer
          MyCommon.QueryStr = "Select RewardAmountTypeID from OfferRewards where OfferID=" & OfferID & " and RewardTypeID=1 and RewardID=" & RewardID
          rst2 = MyCommon.LRT_Select()
          If rst2.Rows.Count > 0 Then
            AmountType = MyCommon.NZ(rst2.Rows(0).Item("RewardAmountTypeID"), 0)
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
					
          MyCommon.QueryStr = "Select OfferTypeID from Offers where OfferID=" & OfferID & " and OfferTypeID=5"
          rst2 = MyCommon.LRT_Select()
          If rst2.Rows.Count > 0 Then
            MyCommon.QueryStr = "Update Offers with (RowLock) set PriorityLevel=" & defaultECPriority & " where OfferID=" & OfferID
          Else
            MyCommon.QueryStr = "Update Offers with (RowLock) set PriorityLevel=" & defaultPriority & " where OfferID=" & OfferID
          End If
          MyCommon.LRT_Execute()
        End If
			
      End If
      If (Request.QueryString("trigger") <> "") Then
        If (Request.QueryString("trigger") = "1") Then
          ' set  TriggerQty=Xbox
          'If (TriggerQty = ApplyToLimit) Then ValueRadio = 1
          TriggerQty = MyCommon.Extract_Val(Request.QueryString("Xbox"))
          If (TriggerQty = 0) Then
            TriggerQty = 1
          End If
          MyCommon.QueryStr = "update OfferRewards with (RowLock) set TriggerQty=" & TriggerQty & "," & _
                              "ApplyToLimit=" & TriggerQty & " where RewardID=" & RewardID
          MyCommon.LRT_Execute()
          ValueRadio = 1
        ElseIf (Request.QueryString("trigger") = "2") Then
          ' Set  and TriggerQty=Xbox2+Ybox2 and ApplyToLimit=Ybox2
          TriggerQty = Int(MyCommon.Extract_Val(Request.QueryString("Xbox2"))) + Int(MyCommon.Extract_Val(Request.QueryString("Ybox2")))
          ApplyToLimit = MyCommon.Extract_Val(Request.QueryString("Ybox2"))
          MyCommon.QueryStr = "update OfferRewards with (RowLock) set TriggerQty=" & TriggerQty & "," & _
                              "ApplyToLimit=" & ApplyToLimit & " where RewardID=" & RewardID
          MyCommon.LRT_Execute()
          ValueRadio = 2
        ElseIf (Request.QueryString("trigger") = "3") Then
          ' Set  and TriggerQty=Xbox3 and ApplyToLimit=0
          TriggerQty = MyCommon.Extract_Val(Request.QueryString("Xbox3"))
          MyCommon.QueryStr = "update OfferRewards with (RowLock) set TriggerQty=" & TriggerQty & "," & _
                              "ApplyToLimit=0 where RewardID=" & RewardID
          MyCommon.LRT_Execute()
          ' they picked pro rate make sure best deal is set to off
          MyCommon.QueryStr = "update discounts with (RowLock) set BestDeal=0,StaticFuel=0 where DiscountID=" & LinkID
          MyCommon.LRT_Execute()
          BestDeal = False
          StaticFuel = False
          ValueRadio = 3
        End If
      End If
    End If
    
    If Not (bUseTemplateLocks And bDisallowEditPp) Then
      If (Request.QueryString("ProgramID") <> "") Then
        ProgramID = Request.QueryString("ProgramID")
        If ProgramID > 0 Then
          If SVLinkID > 0 Then
            MyCommon.QueryStr = "dbo.pa_CM_UpdateRewardStoredValues"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@ProgramID", SqlDbType.BigInt).Value = ProgramID
            MyCommon.LRTsp.Parameters.Add("@WeightUOM", SqlDbType.Bit).Value = 0
            MyCommon.LRTsp.Parameters.Add("@Linkid", SqlDbType.BigInt).Value = SVLinkID
            MyCommon.LRTsp.ExecuteNonQuery()
            MyCommon.Close_LRTsp()
          Else
            MyCommon.QueryStr = "dbo.pa_CM_InsertDiscountStoredValues"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@ProgramID", SqlDbType.BigInt).Value = ProgramID
            MyCommon.LRTsp.Parameters.Add("@WeightUOM", SqlDbType.Bit).Value = 0
            MyCommon.LRTsp.Parameters.Add("@LinkID", SqlDbType.BigInt).Direction = ParameterDirection.Output
            MyCommon.LRTsp.ExecuteNonQuery()
            SVLinkID = MyCommon.LRTsp.Parameters("@LinkID").Value
            MyCommon.Close_LRTsp()
            MyCommon.QueryStr = "update discounts with (RowLock) set SVLinkID=" & SVLinkID & " where DiscountID=" & LinkID
            MyCommon.LRT_Execute()
          End If
        Else
          If SVLinkID > 0 Then
            MyCommon.QueryStr = "update CM_RewardStoredValues with (RowLock) set ProgramID=Null where RewardStoredValuesID=" & SVLinkID
            MyCommon.LRT_Execute()
          End If
          If RewardAmountTypeID = 10 Then
            infoMessage = Copient.PhraseLib.Lookup("cpe-discount-selectprogram", LanguageID)
          End If
        End If
        
        If (RewardLimit <> "0") Then
          RewardLimit = "0"
          MyCommon.QueryStr = "update offerRewards with (RowLock) set RewardLimit=" & RewardLimit & " where RewardID=" & RewardID
          MyCommon.LRT_Execute()
        End If
      Else
        If SVLinkID > 0 Then
          MyCommon.QueryStr = "update CM_RewardStoredValues with (RowLock) set ProgramID=Null where RewardStoredValuesID=" & SVLinkID
          MyCommon.LRT_Execute()
        End If
        If RewardAmountTypeID = 10 Then
          infoMessage = Copient.PhraseLib.Lookup("cpe-discount-selectprogram", LanguageID)
        End If
      End If

      If (Request.QueryString("WeightUom") <> "") Then
        WeightUom = MyCommon.Extract_Val(Request.QueryString("WeightUom"))
        If SVLinkID > 0 Then
          MyCommon.QueryStr = "update CM_RewardStoredValues with (RowLock) set WeightUOM=" & WeightUom & " where RewardStoredValuesID=" & SVLinkID
          MyCommon.LRT_Execute()
        End If
      Else
        If SVLinkID > 0 Then
          MyCommon.QueryStr = "update CM_RewardStoredValues with (RowLock) set WeightUOM=0 where RewardStoredValuesID=" & SVLinkID
          MyCommon.LRT_Execute()
        End If
      End If
      
    End If

    If Not (bUseTemplateLocks And bDisallowEditLimit) Then
      If (Request.QueryString("selectadv") <> "") Then
        AdvancedLimitID = Request.QueryString("selectadv")
        If AdvancedLimitID > 0 Then
          MyCommon.QueryStr = "select AL.PromoVarID,AL.LimitTypeID, AL.LimitValue, AL.LimitPeriod " & _
                              "from CM_AdvancedLimits as AL with (NoLock) where Deleted=0 and LimitID='" & AdvancedLimitID & "';"
          rst = MyCommon.LRT_Select
          If (rst.Rows.Count > 0) Then
            VarID = MyCommon.NZ(rst.Rows(0).Item("PromoVarID"), 0)
            RewardLimitTypeID = MyCommon.NZ(rst.Rows(0).Item("LimitTypeID"), 5)
            RewardLimit = MyCommon.NZ(rst.Rows(0).Item("LimitValue"), 0)
            DistPeriod = MyCommon.NZ(rst.Rows(0).Item("LimitPeriod"), 0)
            MyCommon.QueryStr = "update OfferRewards with (RowLock) set AdvancedLimitID=" & AdvancedLimitID & _
                                ",RewardDistLimitVarID=" & VarID & _
                                ",RewardLimitTypeID=" & RewardLimitTypeID & _
                                ",RewardLimit=" & RewardLimit & _
                                ",RewardDistPeriod=" & DistPeriod & _
                                " where RewardID=" & RewardID & ";"
            MyCommon.LRT_Execute()
          Else
            MyCommon.QueryStr = "update OfferRewards with (RowLock) set AdvancedLimitID=0" & _
                                ",RewardLimitTypeID=2" & _
                                ",RewardDistPeriod=0" & _
                                ",RewardLimit=0.0" & _
                                " where RewardID=" & RewardID & ";"
            MyCommon.LRT_Execute()
          End If
        Else
          MyCommon.Open_LogixXS()
          MyCommon.QueryStr = "select PromoVarID, VarTypeID, LinkID " & _
                              "from PromoVariables with (NoLock) where Deleted=0 and VarTypeID=4 and LinkID=" & RewardID & ";"
          rst = MyCommon.LXS_Select
          If (rst.Rows.Count > 0) Then
            VarID = MyCommon.NZ(rst.Rows(0).Item("PromoVarID"), 0)
          Else
            VarID = 0
          End If
          MyCommon.QueryStr = "update OfferRewards with (RowLock) set AdvancedLimitID=" & AdvancedLimitID & _
                              ",RewardDistLimitVarID=" & VarID & _
                              ",RewardDistPeriod=0" & _
                              " where RewardID=" & RewardID & ";"
          MyCommon.LRT_Execute()
          MyCommon.Close_LogixXS()
        End If
      End If
      If AdvancedLimitID = 0 Then
        'RewardLimitTypeID
        If (Request.QueryString("limitvaluetype") <> "") Then
          RewardLimitTypeID = MyCommon.Extract_Val(Request.QueryString("limitvaluetype"))
          MyCommon.QueryStr = "update OfferRewards with (RowLock) set RewardLimitTypeID=" & RewardLimitTypeID & " where RewardID=" & RewardID
          MyCommon.LRT_Execute()
          If (RewardLimitTypeID < 0) Then
            infoMessage = Copient.PhraseLib.Lookup("reward.badlimit", LanguageID)
          End If
        End If
        If (Request.QueryString("limitvalue") <> "") Then
          RewardLimit = MyCommon.Extract_Val(Request.QueryString("limitvalue"))
          MyCommon.QueryStr = "update OfferRewards with (RowLock) set RewardLimit=" & RewardLimit & " where RewardID=" & RewardID
          MyCommon.LRT_Execute()
        End If
        If (Request.QueryString("limitexceed") = "on") Then
          MyCommon.QueryStr = "update discounts with (RowLock) set AllowNegative=1 where DiscountID=" & LinkID & ";"
          MyCommon.LRT_Execute()
          AllowNegative = True
        Else
          MyCommon.QueryStr = "update discounts with (RowLock) set AllowNegative=0 where DiscountID=" & LinkID & ";"
          MyCommon.LRT_Execute()
          AllowNegative = False
        End If
        If (Request.QueryString("form_DistPeriod") <> "") Then
          DistPeriod = Int(MyCommon.Extract_Val(Request.QueryString("form_DistPeriod")))
          If DistPeriod = 0 Then
            MyCommon.QueryStr = "update OfferRewards with (RowLock) set RewardDistPeriod =0 where RewardID=" & RewardID
          ElseIf DistPeriod = -1 Then
            MyCommon.QueryStr = "update OfferRewards with (RowLock) set RewardDistPeriod =-1 where RewardID=" & RewardID
          Else
            MyCommon.QueryStr = "update OfferRewards with (RowLock) set RewardDistPeriod =" & DistPeriod & " where RewardID=" & RewardID
          End If
          MyCommon.LRT_Execute()
          ' someone saves - let's do the special case and set a promo variable if the
          ' distribution's greater than zero and the promo variable doesn't already exist
          If DistPeriod <> 0 Then
            MyCommon.QueryStr = "select RewardDistLimitVarID from OfferRewards with (NoLock) where RewardID=" & RewardID
            rst = MyCommon.LRT_Select
            For Each row In rst.Rows
              If (MyCommon.NZ(row.Item("RewardDistLimitVarID"), 0) = 0) Then
                'dbo.pa_DistributionVar_Create @OfferID bigint, @VarID bigint OUTPUT
                MyCommon.Open_LogixXS()
                MyCommon.QueryStr = "dbo.pc_RewardLimitVar_Create"
                MyCommon.Open_LXSsp()
                MyCommon.LXSsp.Parameters.Add("@RewardID", SqlDbType.BigInt).Value = RewardID
                MyCommon.LXSsp.Parameters.Add("@VarID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                MyCommon.LXSsp.ExecuteNonQuery()
                VarID = MyCommon.LXSsp.Parameters("@VarID").Value
                MyCommon.Close_LXSsp()
                MyCommon.Close_LogixXS()
                MyCommon.QueryStr = "update OfferRewards with (RowLock) set RewardDistLimitVarID=" & VarID & " where RewardID=" & RewardID
                MyCommon.LRT_Execute()
              End If
            Next
          End If
        End If
      End If
    End If

    
    If Not (bUseTemplateLocks And (bDisallowEditSpc And bDisallowEditDist)) Then
      If (Request.QueryString("usespecialpricing") = "on") Then
        Dim RepeatAtFormVal As Integer
        RepeatAtFormVal = 0
        If (Request.QueryString("repeat") > 0) Then
          RepeatAtFormVal = Request.QueryString("repeat")
        End If
        If (Not UseSpecialPricing) Then
          ' were making the switch to special pricing lets empty the table
          MyCommon.QueryStr = "delete from RewardTiers with (RowLock) where RewardID=" & RewardID
          MyCommon.LRT_Execute()
        End If
        MyCommon.QueryStr = "update OfferRewards with (RowLock) set UseSpecialPricing=1,SPRepeatAtOccur=" & RepeatAtFormVal & " where RewardID=" & RewardID
        MyCommon.LRT_Execute()
        MyCommon.QueryStr = "select max(TierLevel) as maxtier from RewardTiers with (NoLock) where RewardID=" & RewardID
        rst = MyCommon.LRT_Select
        Dim tierHolder As Integer
        tierHolder = MyCommon.NZ(rst.Rows(0).Item("maxtier"), 1)
        For x = 1 To tierHolder
          'Response.Write("->@ConditionID: " & ConditionID & " @TierLevel: " & x & " @AmtRequired: " & Request.QueryString("tier" & x) & "<br />")
          MyCommon.QueryStr = "dbo.pt_RewardTiers_Update"
          MyCommon.Open_LRTsp()
          MyCommon.LRTsp.Parameters.Add("@RewardID", SqlDbType.BigInt).Value = MyCommon.Extract_Val(RewardID)
          MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int).Value = x
          MyCommon.LRTsp.Parameters.Add("@RewardAmount", SqlDbType.Decimal).Value = MyCommon.Extract_Val(Request.QueryString("tier" & x))
          MyCommon.LRTsp.ExecuteNonQuery()
          MyCommon.Close_LRTsp()
        Next
        UseSpecialPricing = True
      Else
        If Not (bUseTemplateLocks And bDisallowEditSpc) Then
          ' ok we need to find out if special pricing was on before or not
          ' if it was then we need to clear out the reward tiers table
          If (UseSpecialPricing) Then
            MyCommon.QueryStr = "delete from RewardTiers with (RowLock) where RewardID=" & RewardID
            MyCommon.LRT_Execute()
          End If
          MyCommon.QueryStr = "update OfferRewards with (RowLock) set UseSpecialPricing=0 where RewardID=" & RewardID
          MyCommon.LRT_Execute()
          UseSpecialPricing = False
        End If
        If Not (bUseTemplateLocks And bDisallowEditDist) Then
          If (Tiered = 0) Then
            'dbo.pt_ConditionTiers_Update @ConditionID bigint, @TierLevel int,@AmtRequired decimal(12,3)
            MyCommon.QueryStr = "dbo.pt_RewardTiers_Update"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@RewardID", SqlDbType.BigInt).Value = RewardID
            MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int).Value = 0
            If (Not IsNumeric(Request.QueryString("tier0"))) Then
                MyCommon.LRTsp.Parameters.Add("@RewardAmount", SqlDbType.Decimal).Value = 0
            Else  
            MyCommon.LRTsp.Parameters.Add("@RewardAmount", SqlDbType.Decimal).Value = MyCommon.Extract_Val(Request.QueryString("tier0"))
            End If                
            MyCommon.LRTsp.ExecuteNonQuery()
            MyCommon.Close_LRTsp()
          Else
            ' delete the current tier ammounts
            MyCommon.QueryStr = "delete from RewardTiers with (RowLock) where RewardID=" & RewardID
            MyCommon.LRT_Execute()
            For x = 1 To NumTiers
              'Response.Write("->@ConditionID: " & ConditionID & " @TierLevel: " & x & " @AmtRequired: " & Request.QueryString("tier" & x) & "<br />")
              MyCommon.QueryStr = "dbo.pt_RewardTiers_Update"
              MyCommon.Open_LRTsp()
              MyCommon.LRTsp.Parameters.Add("@RewardID", SqlDbType.BigInt).Value = MyCommon.Extract_Val(RewardID)
              MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int).Value = x
              MyCommon.LRTsp.Parameters.Add("@RewardAmount", SqlDbType.Decimal).Value = MyCommon.Extract_Val(Request.QueryString("tier" & x))
              MyCommon.LRTsp.ExecuteNonQuery()
              MyCommon.Close_LRTsp()
            Next
          End If
        End If
      End If
    End If

    MyCommon.QueryStr = "update OfferRewards with (RowLock) set TCRMAStatusFlag=2,CMOAStatusFlag=2 where RewardID=" & RewardID
    MyCommon.LRT_Execute()
    MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where offerid=" & OfferID
    MyCommon.LRT_Execute()
    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.rew-discount", LanguageID))
  
    If (Request.QueryString("pgroup-add1") <> "" Or Request.QueryString("pgroup-rem1") <> "" Or Request.QueryString("pgroup-add2") <> "" Or Request.QueryString("pgroup-rem2") <> "") Then
      MyCommon.QueryStr = "update OfferRewards with (RowLock) set TCRMAStatusFlag=3,CMOAStatusFlag=2 where RewardID=" & RewardID
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where offerid=" & OfferID
      MyCommon.LRT_Execute()
    End If
  ElseIf (GetCgiValue("add") <> "") Then
        
    MyCommon.QueryStr = "select ProductGroupID, ExcludedProdGroupID from OfferRewards with (NoLock) where RewardID=" & RewardID & " and deleted=0;"
    rst = MyCommon.LRT_Select
	
    GName = GetCgiValue("modprodgroupname")

    If (rst.Rows(0).Item("ProductGroupID") = 0) Then 'Create new product group
                 
      MyCommon.QueryStr = "SELECT ProductGroupID FROM ProductGroups with (NoLock) WHERE Name = '" & IIf(GName.Contains("'"), GName.Replace("'", "''"), GName) & "' AND Deleted=0"
      rst1 = MyCommon.LRT_Select
      If (rst1.Rows.Count > 0) Then
        ProductGroupID = rst1.Rows(0).Item("ProductGroupID")
      Else
        MyCommon.QueryStr = "dbo.pt_ProductGroups_Insert"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 200).Value = GName
        GName = MyCommon.Parse_Quotes(Logix.TrimAll(GName))
        MyCommon.LRTsp.Parameters.Add("@AnyProduct", SqlDbType.Bit).Value = 0
        MyCommon.LRTsp.Parameters.Add("@AdminID", SqlDbType.Int).Value = AdminUserID
        MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt).Direction = ParameterDirection.Output
        MyCommon.LRTsp.ExecuteNonQuery()
        ProductGroupID = MyCommon.LRTsp.Parameters("@ProductGroupID").Value
        Send("<input type=""hidden"" id=""NewCreatedProdGroupID"" name=""NewCreatedProdGroupID"" value=""" & ProductGroupID & """ />")
        MyCommon.Activity_Log(5, ProductGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.pgroup-create", LanguageID))
        MyCommon.Close_LRTsp()
      End If
    Else
      ProductGroupID = rst.Rows(0).Item("ProductGroupID")
    End If
        
    If (Trim(GetCgiValue("ExtProductID")) = "") Then
      infoMessage = Copient.PhraseLib.Lookup("pgroup-edit.badid", LanguageID)
    Else
      '4722 Above 
      Dim bGoodItemCode As Boolean = True
      Dim bProductExist As Boolean = True
      Dim bCreateProducts As Boolean = MyCommon.Fetch_SystemOption(150)
      Dim bAddProduct As Boolean = True
      ' desired product add to group   
      ' dbo.pt_ProdGroupItems_Insert  @ExtProductID nvarchar(20), @ProductGroupID bigint, @ProductTypeID int, @Status int OUTPU
      'Send("Inserting product type : " & GetCgiValue("producttype"))
      If (Int(GetCgiValue("producttype")) = 1) Then
        Integer.TryParse(MyCommon.Fetch_SystemOption(52), IDLength)
      ElseIf (Int(GetCgiValue("producttype")) = 2) Then
        Integer.TryParse(MyCommon.Fetch_SystemOption(54), IDLength)
      Else
        IDLength = 0
      End If
      If (IDLength > 0) Then
		If (Int(GetCgiValue("producttype")) = 2) Then
			ExtProductID = MyCommon.Parse_Quotes(Left(Trim(GetCgiValue("ExtProductID")), 120)).PadLeft(IDLength, "0")
		Else
			ExtProductID = MyCommon.Parse_Quotes(Left(Trim(GetCgiValue("ExtProductID")), 26)).PadLeft(IDLength, "0")
		End If
      Else
        If (Int(GetCgiValue("producttype")) = 2) Then
			ExtProductID = MyCommon.Parse_Quotes(Left(Trim(GetCgiValue("ExtProductID")), 120))
		Else
			ExtProductID = MyCommon.Parse_Quotes(Left(Trim(GetCgiValue("ExtProductID")), 26))
		End If
      End If
      'Don't change the product description if it is saved as blank
      MyCommon.QueryStr = "Select Description from Products where ExtProductID='" & ExtProductID & "' and ProductTypeID=" & Int(GetCgiValue("producttype")) & ";"
      prodDT = MyCommon.LRT_Select()
      If prodDT.Rows.Count > 0 Then
        Description = MyCommon.NZ(prodDT.Rows(0).Item("Description"), "")
      Else
        bProductExist = False
      End If
      If GetCgiValue("productdesc") <> "" Then
        Description = GetCgiValue("productdesc")
      End If
        
      If (MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(97)) = 1 AndAlso CleanUPC(GetCgiValue("ExtProductID")) = False) Then
        infoMessage = Copient.PhraseLib.Lookup("pgroup-edit.badid", LanguageID)
        bGoodItemCode = False
      ElseIf bProductExist = False AndAlso bCreateProducts = False Then
        infoMessage = Copient.PhraseLib.Lookup("pgroup-edit.productnotexist", LanguageID)
        bAddProduct = False
      End If
    
      Const Product_NotChanged_Status As Integer = 0
      Const Not_Changed_Status As Integer = 0
      Const Product_Add_Status As Integer = 1
      Const Add_Status As Integer = 1
      Const Product_Update_Status As Integer = 2
      Const Update_Status As Integer = 2
      Dim productOutputStatus As Integer = 0

      bGoodItemCode = True
      If (MyCommon.Extract_Val(GetCgiValue("ExtProductID")) < 1) Or (Int(MyCommon.Extract_Val(GetCgiValue("ExtProductID"))) <> MyCommon.Extract_Val(GetCgiValue("ExtProductID"))) Then
        infoMessage = Copient.PhraseLib.Lookup("pgroup-edit.badid", LanguageID)
        bGoodItemCode = False
      ElseIf (MyCommon.Fetch_CM_SystemOption(82) = "1" AndAlso MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM) = True) Then
        Dim sItemCode As String = GetCgiValue("ExtProductID").ToString
        Dim productType As Integer = Int(GetCgiValue("producttype"))
        If (productType = 1) Then
          If (CheckItemCode(sItemCode, infoMessage) = False) Then
            bGoodItemCode = False
          End If
        End If
      ElseIf (CleanUPC(GetCgiValue("ExtProductID")) = False) Then
        infoMessage = Copient.PhraseLib.Lookup("pgroup-edit.badid", LanguageID)
        bGoodItemCode = False
      End If

      If bGoodItemCode = True AndAlso bAddProduct = True Then
        MyCommon.QueryStr = "dbo.pa_ProdGroupItems_ManualInsert"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@ExtProductID", SqlDbType.NVarChar, 120).Value = ExtProductID
        MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt).Value = ProductGroupID
        MyCommon.LRTsp.Parameters.Add("@ProductTypeID", SqlDbType.Int).Value = Int(GetCgiValue("producttype"))
        MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 200).Value = Description
        MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
        MyCommon.LRTsp.Parameters.Add("@ProductStatus", SqlDbType.Int).Direction = ParameterDirection.Output
        MyCommon.LRTsp.ExecuteNonQuery()
        outputStatus = MyCommon.LRTsp.Parameters("@Status").Value
        productOutputStatus = MyCommon.LRTsp.Parameters("@ProductStatus").Value
      End If
      MyCommon.Close_LRTsp()
      MyCommon.QueryStr = "Select PhraseID,Name from ProductTypes where ProductTypeID=" & Int(GetCgiValue("producttype")) & ";"
      Dim productTypeTable As DataTable = MyCommon.LRT_Select()
      Dim typePhrase As Integer = 0
      If (productTypeTable.Rows.Count > 0) Then
        typePhrase = MyCommon.NZ(productTypeTable.Rows(0).Item("PhraseID"), 0)
      End If
      If (productOutputStatus > Product_NotChanged_Status) Then
        If (productOutputStatus = Product_Add_Status) Then
          MyCommon.Activity_Log(5, ProductGroupID, AdminUserID, Copient.PhraseLib.Lookup("term.created", LanguageID) & " " & Copient.PhraseLib.Lookup("term.product", LanguageID).ToLower() & " " & Copient.PhraseLib.Lookup("term.id", LanguageID) & " " & ExtProductID & _
                                IIf(typePhrase > 0, "(" & Copient.PhraseLib.Lookup(typePhrase, LanguageID) & ")", ""))
        ElseIf (productOutputStatus = Product_Update_Status) Then
          MyCommon.Activity_Log(5, ProductGroupID, AdminUserID, Copient.PhraseLib.Lookup("term.updated", LanguageID) & " " & Copient.PhraseLib.Lookup("term.product", LanguageID).ToLower() & " " & Copient.PhraseLib.Lookup("term.id", LanguageID) & " " & ExtProductID & _
                                IIf(typePhrase > 0, "(" & Copient.PhraseLib.Lookup(typePhrase, LanguageID) & ")", ""))
        End If
      End If
      If (outputStatus > Not_Changed_Status) Then
        If (outputStatus = Add_Status) Then
          MyCommon.Activity_Log(5, ProductGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.pgroup-add", LanguageID) & " " & ExtProductID & IIf(typePhrase > 0, "(" & Copient.PhraseLib.Lookup(typePhrase, LanguageID) & ")", ""))
        ElseIf (outputStatus = Update_Status) Then
          'Product was updated to be a manual product entry from a linked product.
          'MyCommon.Activity_Log(5, ProductGroupID, AdminUserID, "Updated product ID " & ExtProductID)
        End If
      End If
	
      If (outputStatus <> 0 OrElse productOutputStatus <> 0) Then
        MyCommon.QueryStr = "update productgroups with (RowLock) set updatelevel=updatelevel+1, LastUpdate=getdate() where ProductGroupID=" & ProductGroupID
        MyCommon.LRT_Execute()
      End If
      'If infoMessage = "" Then
      '    Response.Status = "301 Moved Permanently"
      '    Response.AddHeader("Location", "pgroup-edit.aspx?ProductGroupID=" & ProductGroupID)
      'End If
    End If
  
  ElseIf (GetCgiValue("mremove") <> "") Then
    ' desired product remove from group  dbo.pt_GroupMembership_Delete_ByID  @MembershipID bigint
    ' dbo.pt_ProdGroupItems_Delete  @ExtProductID nvarchar(20), @ProductGroupID bigint, @ProductTypeID int, @Status int OUTPUT
    MyCommon.QueryStr = "select ProductGroupID, ExcludedProdGroupID from OfferRewards with (NoLock) where RewardID=" & RewardID & " and deleted=0;"
    rst = MyCommon.LRT_Select
    ' Build product group name 
    
    If (rst.Rows(0).Item("ProductGroupID") = 0) Then
      infoMessage = "Can not remove products. No productgroup associated with the offer"
    Else
      If (GetCgiValue("ExtProductID") <> "") Then
        If (Int(GetCgiValue("producttype")) = 1) Then
          Integer.TryParse(MyCommon.Fetch_SystemOption(52), IDLength)
        ElseIf (Int(GetCgiValue("producttype")) = 2) Then
          Integer.TryParse(MyCommon.Fetch_SystemOption(54), IDLength)
        Else
          IDLength = 0
        End If
        If (IDLength > 0) Then
			If (Int(GetCgiValue("producttype")) = 2) Then
				ExtProductID = MyCommon.Parse_Quotes(Left(Trim(GetCgiValue("ExtProductID")), 120)).PadLeft(IDLength, "0")
			Else
				ExtProductID = MyCommon.Parse_Quotes(Left(Trim(GetCgiValue("ExtProductID")), 26)).PadLeft(IDLength, "0")
			End If          
        Else
			If (Int(GetCgiValue("producttype")) = 2) Then
				ExtProductID = MyCommon.Parse_Quotes(Left(Trim(GetCgiValue("ExtProductID")), 120))
			Else
				ExtProductID = MyCommon.Parse_Quotes(Left(Trim(GetCgiValue("ExtProductID")), 26))
			End If
        End If
      
        ' check if the product is linked to the group, if so then exclude the product from the group.
        MyCommon.QueryStr = "select PROD.ProductID, PGI.ExtHierarchyID from Products as PROD with (NoLock) " & _
                            "inner join ProdGroupItems as PGI with (NoLock) on PGI.ProductID = PROD.ProductID " & _
                            "where PGI.Deleted=0 and PGI.ProductGroupID=" & ProductGroupID & " and IsNull(PGI.ExtHierarchyID, '') <> ''" & _
                            "   and PROD.ExtProductID='" & MyCommon.Parse_Quotes(ExtProductID) & "' and PROD.ProductTypeID=" & Int(GetCgiValue("producttype"))
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
          If rst.Rows(0).Item("ProductID") > 0 AndAlso MyCommon.NZ(rst.Rows(0).Item("ExtHierarchyID"), "") <> "" Then
            MyCommon.QueryStr = "insert into ProdGroupHierarchyExclusions (ExtHierarchyID,ProductGroupID, HierarchyLevel, LevelID) " & _
                                "      values ('" & MyCommon.Parse_Quotes(MyCommon.NZ(rst.Rows(0).Item("ExtHierarchyID"), "")) & "', " & ProductGroupID & ", 2, '" & rst.Rows(0).Item("ProductID") & "')"
            MyCommon.LRT_Execute()
          End If
        End If
      
        MyCommon.Open_LogixRT()
        MyCommon.QueryStr = "dbo.[pt_ProdGroupItems_DeleteItem]"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@ExtProductID", SqlDbType.NVarChar, 120).Value = ExtProductID
        MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt).Value = ProductGroupID
        MyCommon.LRTsp.Parameters.Add("@ProductTypeID", SqlDbType.Int).Value = Int(GetCgiValue("producttype"))
        MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
        MyCommon.LRTsp.ExecuteNonQuery()
        outputStatus = MyCommon.LRTsp.Parameters("@Status").Value
        MyCommon.Close_LRTsp()
        MyCommon.Activity_Log(5, ProductGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.pgroup-remove", LanguageID) & " " & ExtProductID)
        If (outputStatus <> 0) Then
          infoMessage = Copient.PhraseLib.Lookup("pgroup-edit.notmember", LanguageID)
        Else
          MyCommon.QueryStr = "update productgroups with (RowLock) set updatelevel=updatelevel+1, LastUpdate=getdate() where ProductGroupID=" & ProductGroupID
          MyCommon.LRT_Execute()
        End If
      Else
        infoMessage = Copient.PhraseLib.Lookup("pgroup-edit.badid", LanguageID)
      End If
    End If
	
    
  ElseIf (GetCgiValue("remove") <> "" AndAlso GetCgiValue("PKID") <> "") Then
    MyCommon.Open_LogixRT()
    For i = 0 To Request.QueryString.GetValues("PKID").GetUpperBound(0)
      MyCommon.QueryStr = "select P.ExtProductID from Products as P with (NoLock) Inner Join ProdGroupItems as PGI " & _
                          "with (NoLock) on P.ProductID=PGI.ProductID where PGI.PKID=" & Request.QueryString.GetValues("PKID")(i)
      rst = MyCommon.LRT_Select()
      If rst.Rows.Count > 0 Then
        upc = rst.Rows(0).Item("ExtProductID")
      End If
      rst = Nothing
      MyCommon.QueryStr = "dbo.pt_ProdGroupItems_Delete_ByID"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@PKID", SqlDbType.BigInt).Value = Request.QueryString.GetValues("PKID")(i)
      MyCommon.LRTsp.ExecuteNonQuery()
      MyCommon.Close_LRTsp()
      MyCommon.Activity_Log(5, ProductGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.pgroup-remove", LanguageID) & " " & upc)
    Next
    MyCommon.QueryStr = "update productgroups with (RowLock) set updatelevel=updatelevel+1, LastUpdate=getdate() where ProductGroupID=" & ProductGroupID
    MyCommon.LRT_Execute()
  End If
    
  'ProgramID = MyCommon.Extract_Val(Request.QueryString("Pselected"))
  'WeightUom = MyCommon.Extract_Val(Request.QueryString("WeightUom"))
    
  
  ' ok lets find out if were possibly supposed to show the transaction level choices and if we already have one selected
  MyCommon.QueryStr = "select LinkID,RewardAmountTypeID,ProductGroupID,ExcludedProdGroupID from OfferRewards with (NoLock) where RewardID=" & RewardID
  rst = MyCommon.LRT_Select
  If (rst.Rows.Count > 0) Then
    For Each row In rst.Rows
      If (MyCommon.NZ(row.Item("ProductGroupID"), 0) = 0) Then
        TransactionLevelPossible = True
        TransactionLevelSelected = True
      End If
      If (MyCommon.NZ(row.Item("RewardAmountTypeID"), 0) = 7) Then
        TransactionLevelSelected = False
        TransactionLevelPossible = True
        FreeItemSelected = True
      ElseIf (MyCommon.NZ(row.Item("RewardAmountTypeID"), 0) = 10) Then
        TransactionLevelSelected = False
        TransactionLevelPossible = True
        StoredValueSelected = True
      ElseIf (MyCommon.NZ(row.Item("RewardAmountTypeID"), 0) > 7) Then
        If (MyCommon.NZ(row.Item("ProductGroupID"), 0) = 0) Then
          TransactionLevelSelected = True
        Else
          TransactionLevelSelected = False
        End If
      End If
    Next
  Else
    TransactionLevelPossible = True
    TransactionLevelSelected = True
  End If

  If ProductGroupID > 0 Then
    MyCommon.QueryStr = "select count(*) as GCount from ProdGroupItems with (NoLock) where ProductGroupID = " & ProductGroupID & " And Deleted = 0"
    rst = MyCommon.LRT_Select()
    For Each row In rst.Rows
      GroupSize = row.Item("GCount")
    Next
    MyCommon.QueryStr = "select count(*) as PCount from ProdGroupItems PGI with (NoLock) inner join products PRD on PGI.productid = PRD.productid " & _
                  "where PGI.ProductGroupID = " & ProductGroupID & " And PGI.Deleted = 0 And (PRD.Description IS NULL OR  PRD.Description = '')"
    rst = MyCommon.LRT_Select()
    If rst.Rows.Count > 0 Then
      ProductsWithoutDesc = rst.Rows(0).Item("PCount")
    End If
  Else
    MyCommon.QueryStr = "select count(*) as GCount from ProdGroupItems PGI with (NoLock) inner join OfferRewards ORe on ORe.ProductGroupID = PGI.ProductGroupID " & _
                        "where ORe.RewardID = " & RewardID & " And ORe.Deleted = 0 And PGI.Deleted = 0"
    rst = MyCommon.LRT_Select()
    For Each row In rst.Rows
      GroupSize = row.Item("GCount")
    Next
        
    MyCommon.QueryStr = "select count(*) as PCount from ProdGroupItems PGI with (NoLock) inner join products PRD on PGI.productid = PRD.productid " & _
                        "inner join OfferRewards ORe with (NoLock) on ORe.ProductGroupID = PGI.ProductGroupID " & _
                        "where ORe.RewardID = " & RewardID & " And PGI.Deleted = 0 And (PRD.Description IS NULL OR PRD.Description = '') And ORe.Deleted = 0"
    rst = MyCommon.LRT_Select()
    If rst.Rows.Count > 0 Then
      ProductsWithoutDesc = rst.Rows(0).Item("PCount")
    End If
  End If
		
  ShowAllItems = (GetCgiValue("showall") = "true")
  Dim bBlankDescProd As Boolean = (MyCommon.Fetch_SystemOption(206) = "1")
  Dim sBlankDescOrderByStr As String = " order by case when Description is null or Description = '' then nullif(Description, '') else CAST(ExtProductID as bigint) end, CAST(ExtProductID as bigint), ExtProductID DESC;"
  If ProductGroupID > 0 Then
    MyCommon.QueryStr = "select" & If(ShowAllItems, "", " top 100") & " GM.ProductID, PKID, CID.ProductTypeID, ExtProductID, Description, PT.Name as ProductType, PT.PhraseID " & _
                      "from Products as CID with (NoLock) " & _
                      "inner join ProdGroupItems as GM with (NoLock) on CID.ProductID=GM.ProductID " & _
                      "left join ProductTypes as PT with (NoLock) on PT.ProductTypeID=CID.ProductTypeID " & _
                      "where GM.ProductGroupID=" & ProductGroupID & " and GM.Deleted=0 and IsNull(GM.ExtHierarchyID, '')='' " & _
                      "and IsNull(GM.ExtNodeID, '')='' " & If(bBlankDescProd, sBlankDescOrderByStr, " order by ExtProductID;")
  Else
    MyCommon.QueryStr = "select" & If(ShowAllItems, "", " top 100") & " GM.ProductID, PKID, CID.ProductTypeID, ExtProductID, Description, PT.Name as ProductType, PT.PhraseID " & _
                      "from Products as CID with (NoLock) " & _
                      "inner join ProdGroupItems as GM with (NoLock) on CID.ProductID=GM.ProductID " & _
                      "left join ProductTypes as PT with (NoLock) on PT.ProductTypeID=CID.ProductTypeID " & _
                      "inner join OfferRewards as ORe with (NoLock) on ORe.ProductGroupID = gm.ProductGroupID " & _
                      "where ORe.RewardID = " & RewardID & " and GM.Deleted=0 and ORe.Deleted =0 and IsNull(GM.ExtHierarchyID, '')='' " & _
                      "and IsNull(GM.ExtNodeID, '')='' " & If(bBlankDescProd, sBlankDescOrderByStr, " order by ExtProductID;")
    
  End If
    
    
    
  rstItems = MyCommon.LRT_Select()
  ListBoxSize = rstItems.Rows.Count
  Send("<script type=""text/javascript"">")
  Send("function ChangeParentDocument() { ")
  Send("  if (opener != null) {")
  Send("    var newlocation = 'offer-rew.aspx?OfferID=" & OfferID & "'; ")
  Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
  Send("  opener.location = 'offer-rew.aspx?OfferID=" & OfferID & "'; ")
  Send("} ")
  Send("  }")
  Send("}")
  Send("</script>")
%>
<script type="text/javascript">
// JS QuickTags version 1.2
//
// Copyright (c) 2002-2005 Alex King
// http://www.alexking.org/
//
// Licensed under the LGPL license
// http://www.gnu.org/copyleft/lesser.html
//
// This JavaScript will insert the tags below at the cursor position in IE and 
// Gecko-based browsers (Mozilla, Camino, Firefox, Netscape). For browsers that 
// do not support inserting at the cursor position (Safari, OmniWeb) it appends
// the tags to the end of the content.

var edButtons = new Array();
var edOpenTags = new Array();

//
//
// Functions

function edShowButton(button) {
	if (button.access) {
		var accesskey = ' accesskey = "' + button.access + '"'
	}
	else {
		var accesskey = '';
	}
}

function edAddTag(button) {
	if (edButtons[button].tagEnd != '') {
		edOpenTags[edOpenTags.length] = button;
		document.getElementById(edButtons[button].id).value = '/' + document.getElementById(edButtons[button].id).value;
	}
}

function edRemoveTag(button) {
	for (i = 0; i < edOpenTags.length; i++) {
		if (edOpenTags[i] == button) {
			edOpenTags.splice(i, 1);
			document.getElementById(edButtons[button].id).value = document.getElementById(edButtons[button].id).value.replace('/', '');
		}
	}
}

function isValidID() {
        var retVal = true;
        var elemNumericOnly = document.getElementById("NumericOnly");
        var elemID = document.getElementById("productid");
		var elSel = document.getElementById('selected');
        var selectList = "";
		
        if((elemID != null) && (elemID.value.length == 0)) {
            retVal = false;
            alert('<%Sendb(Copient.PhraseLib.Lookup("term.invalid", LanguageID) + " " + Copient.PhraseLib.Lookup("term.productid", LanguageID)) %>');
        }
        if ((elemNumericOnly != null) && (elemNumericOnly.value != "")) {
           if ((elemID != null) && (isNaN(elemID.value))) {
              retVal = false;
              alert('<%Sendb(Copient.PhraseLib.Lookup("product.mustbenumeric", LanguageID)) %>');
           }
        }
		// assemble the list of values from the selected box
        for (i = elSel.length - 1; i>=0; i--) {
          if(elSel.options[i].value != ""){
          if(selectList != "") { selectList = selectList + ","; }
            selectList = selectList + elSel.options[i].value;
          }
		}  
		if (selectList == 1 && retVal == true){
		  retVal = false;
          alert('Product Group should not be "Any Product"');		
		}
  	
        return retVal;
    }

	function isValidProductList() {
  
  
  var retVal = true;
  var elemID = document.getElementById("pasteproducts");
  
  if((elemID != null) && (elemID.value.length == 0)) {
            retVal = false;
            alert('No products were provided');
        }
		
		return retVal;
}
	
function ProductAddSelection() {

  var elemaddsingleprod = document.getElementById("addsingleproduct");
  var elemaddprodlist = document.getElementById("addproductlist");
  var elemradioprodaddselector1 = document.getElementById("prodaddselector1");
  var elemradioprodaddselector2 = document.getElementById("prodaddselector2");
  
  if (elemradioprodaddselector1.checked) {
      //alert('checked'); 
      elemaddprodlist.style.display = "none";
      elemaddsingleprod.style.display = "block";
    } else {
      //alert('unchecked');
      elemaddprodlist.style.display = "block";
      elemaddsingleprod.style.display = "none";
    }

}

function ProductGroupTypeSelection() {
  
  var elemexistingpgroup = document.getElementById("selector");
  var elemaddpeoducts = document.getElementById("directprodaddselector");
  var elemradiopgselect = document.getElementById("pgselectortype1");
  var elemradioaddprod = document.getElementById("pgselectortype2");
  
  if (elemradiopgselect.checked) {
      //alert('checked'); 
      elemaddpeoducts.style.display = "none";
      elemexistingpgroup.style.display = "block";
    } else {
      //alert('unchecked');
      elemaddpeoducts.style.display = "block";
      elemexistingpgroup.style.display = "none";
    }
}

function edCheckOpenTags(button) {
	var tag = 0;
	for (i = 0; i < edOpenTags.length; i++) {
		if (edOpenTags[i] == button) {
			tag++;
		}
	}
	if (tag > 0) {
		return true; // tag found
	}
	else {
		return false; // tag not found
	}
}

//
//
// Insertion code

function edInsertTag(myField, i) {
	//IE support
	if (document.selection) {
		myField.focus();
		sel = document.selection.createRange();
		if (sel.text.length > 0) {
			sel.text = edButtons[i].tagStart + sel.text + edButtons[i].tagEnd;
		}
		else {
			if (!edCheckOpenTags(i) || edButtons[i].tagEnd == '') {
				sel.text = edButtons[i].tagStart;
				edAddTag(i);
			}
			else {
				sel.text = edButtons[i].tagEnd;
				edRemoveTag(i);
			}
		}
		myField.focus();
	}
	//MOZILLA/NETSCAPE support
	else if (myField.selectionStart || myField.selectionStart == '0') {
		var startPos = myField.selectionStart;
		var endPos = myField.selectionEnd;
		var cursorPos = endPos;
		var scrollTop = myField.scrollTop;
		if (startPos != endPos) {
			myField.value = myField.value.substring(0, startPos)
			              + edButtons[i].tagStart
			              + myField.value.substring(startPos, endPos) 
			              + edButtons[i].tagEnd
			              + myField.value.substring(endPos, myField.value.length);
			cursorPos += edButtons[i].tagStart.length + edButtons[i].tagEnd.length;
		}
		else {
			if (!edCheckOpenTags(i) || edButtons[i].tagEnd == '') {
				myField.value = myField.value.substring(0, startPos) 
				              + edButtons[i].tagStart
				              + myField.value.substring(endPos, myField.value.length);
				edAddTag(i);
				cursorPos = startPos + edButtons[i].tagStart.length;
			}
			else {
				myField.value = myField.value.substring(0, startPos) 
				              + edButtons[i].tagEnd
				              + myField.value.substring(endPos, myField.value.length);
				edRemoveTag(i);
				cursorPos = startPos + edButtons[i].tagEnd.length;
			}
		}
		myField.focus();
		myField.selectionStart = cursorPos;
		myField.selectionEnd = cursorPos;
		myField.scrollTop = scrollTop;
	}
	else {
		if (!edCheckOpenTags(i) || edButtons[i].tagEnd == '') {
			myField.value += edButtons[i].tagStart;
			edAddTag(i);
		}
		else {
			myField.value += edButtons[i].tagEnd;
			edRemoveTag(i);
		}
		myField.focus();
	}
}

function edInsertContent(myField, myValue) {
	//IE support
	if (document.selection) {
		myField.focus();
		sel = document.selection.createRange();
		sel.text = myValue;
		myField.focus();
	}
	//MOZILLA/NETSCAPE support
	else if (myField.selectionStart || myField.selectionStart == '0') {
		var startPos = myField.selectionStart;
		var endPos = myField.selectionEnd;
		var scrollTop = myField.scrollTop;
		myField.value = myField.value.substring(0, startPos)
		              + myValue 
                      + myField.value.substring(endPos, myField.value.length);
		myField.focus();
		myField.selectionStart = startPos + myValue.length;
		myField.selectionEnd = startPos + myValue.length;
		myField.scrollTop = scrollTop;
	} else {
		myField.value += myValue;
		myField.focus();
	}
}

<%
  MyCommon.QueryStr = "Select Distinct MT.MarkupID,MT.Tag,MT.Description,MT.PhraseID,MT.NumParams,MT.Param1Name,MT.Param1PhraseID,MT.Param2Name,MT.Param2PhraseID,MT.Param3Name,MT.Param3PhraseID,MT.DisplayOrder,MTU.RewardTypeID,MTU.EngineID,PhT.PhraseID,PhT.LanguageID,Convert(nvarchar(50),Phrase) as Phrase,PTy.Installed from MarkupTags as MT with (NoLock) Inner Join PhraseText as PhT with (NoLock) on MT.PhraseID=PhT.PhraseID Inner Join PrinterTranslation as PTr  with (NoLock) on MT.MarkupID=PTr.MarkupID Inner Join PrinterTypes as PTy with (NoLock) on PTr.PrinterTypeID=PTy.PrinterTypeID Inner Join MarkupTagUsage as MTU with (NoLock) on MT.MarkupID=MTU.MarkupID where MTU.RewardTypeID=1 and MTU.EngineID=" & OfferEngineID & " and PTy.Installed=1 order by MT.DisplayOrder"
  rst = MyCommon.LRT_Select
  Dim funcname As String
  For Each row In rst.Rows
    funcname = row.Item("Tag")
    funcname = funcname.Replace("#", "Amt")
    funcname = funcname.Replace("$", "Dol")
    funcname = funcname.Replace("/", "Off")
    If (row.Item("NumParams") = 0) Then
      Send("function edInsert" & (StrConv(funcname, VbStrConv.ProperCase)) & "(myField) {")
      Send("    myValue = '|" & row.Item("Tag") & "|';")
      Send("    edInsertContent(myField, myValue);")
      Send("}")
    Else
      If (row.Item("Tag") = "UPCA") or (row.Item("Tag") = "EAN13") or (row.Item("Tag") = "CODE39") Then
        Send("function edInsert" & (StrConv(funcname, VbStrConv.ProperCase)) & "(myField) {")
        Send("    var myValue = prompt('" & row.Item("Param1PhraseID") & "', '');")
        Send("    if (myValue) {")
        Send("        myValue = '|" & row.Item("Tag") & "[' + myValue + ']|';")
        Send("        edInsertContent(myField, myValue);")
        Send("    }")
        Send("}")
      Else
        Send("function edInsert" & (StrConv(funcname, VbStrConv.ProperCase)) & "(myField, myValue) {")
        If (UCase(Right(funcname, 3)) = "AMT") then
          Send("    var n = document.getElementById(""functionselect2"").value;")
        Else
          Send("    var n = document.getElementById(""functionselect"").value;")
        End If
        Send("    var myValue = n;")
        Send("    if (myValue) {")
        Send("        myValue = '|" & row.Item("Tag") & "[' + myValue + ']|';")
        Send("        edInsertContent(myField, myValue);")
        Send("    }")
        Send("}")
      End If
    End If
  Next
%>
</script>
<form action="offer-rew-discount.aspx" id="mainform" name="mainform" onsubmit="return saveForm();">
<div id="intro">
  <input type="hidden" id="OfferID" name="OfferID" value="<% sendb(OfferID) %>" />
  <input type="hidden" id="Name" name="Name" value="<% sendb(Name) %>" />
  <input type="hidden" id="RewardID" name="RewardID" value="<% sendb(RewardID) %>" />
  <input type="hidden" id="InvalidRewardAmt" name="InvalidRewardAmt" value="false" />
  <input type="hidden" id="FlagTier" name="FlagTier" value="<% sendb(FlagTier) %>" />
  <input type="hidden" id="ProductGroupID" name="ProductGroupID" value="<% sendb(ProductGroupID) %>" />
  <input type="hidden" id="ExcludedID" name="ExcludedID" value="<% sendb(ExcludedID) %>" />
  <input type="hidden" id="tiered" name="tiered" value="<% sendb(Tiered) %>" />
  <input type="hidden" id="ProgramID" name="ProgramID" value="" />
  <input type="hidden" id="SVType" name="SVType" value="" />
  <input type="hidden" id="OriginalPeriod" name="OriginalPeriod" value="<% sendb(DistPeriod) %>" />
  <input type="hidden" id="ImpliedPeriod" name="ImpliedPeriod" value="<% sendb(DistPeriod) %>" />
  <input type="hidden" id="LimitsDisabled" name="LimitsDisabled" value="<% sendb(bUseTemplateLocks and bDisallowEditLimit) %>" />
  <%Send("<input type=""hidden"" id=""PriorityReset"" name=""PriorityReset"" value=" & PriorityFlag & " />")%>
  <input type="hidden" id="BuckParent" name="BuckParent" value="<% sendb(bBuckParentOffer) %>" />
  <input type="hidden" id="IsTemplate" name="IsTemplate" value="<% 
            if(istemplate)then 
            sendb("IsTemplate")
            else 
            sendb("Not") 
            end if
        %>" />
  <%
    If (IsTemplate) Then
      Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.discountreward", LanguageID), VbStrConv.Lowercase) & "</h1>")
    Else
        If bBuckParentOffer Then
          Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.discountreward", LanguageID), VbStrConv.Lowercase) & " (" & StrConv(Copient.PhraseLib.Lookup("term.bucks", LanguageID), VbStrConv.Lowercase) & ")</h1>")
        Else
          Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.discountreward", LanguageID), VbStrConv.Lowercase) & "</h1>")
        End If
    End If
    If (bUseTemplateLocks And bDisallowEditPg) Then
      Send("<input type=""hidden"" id=""DisallowEditPgOpt"" name=""DisallowEditPgOpt"" value=""1"" />")
    Else
      Send("<input type=""hidden"" id=""DisallowEditPgOpt"" name=""DisallowEditPgOpt"" value=""0"" />")
    End If
    If (bUseTemplateLocks And bDisallowEditDist) Then
      Send("<input type=""hidden"" id=""DisallowEditDistOpt"" name=""DisallowEditDistOpt"" value=""1"" />")
    Else
      Send("<input type=""hidden"" id=""DisallowEditDistOpt"" name=""DisallowEditDistOpt"" value=""0"" />")
    End If
    If (bUseTemplateLocks And bDisallowEditPp) Then
      Send("<input type=""hidden"" id=""DisallowEditPpOpt"" name=""DisallowEditPpOpt"" value=""1"" />")
    Else
      Send("<input type=""hidden"" id=""DisallowEditPpOpt"" name=""DisallowEditPpOpt"" value=""0"" />")
    End If
    If (bUseTemplateLocks And bDisallowEditAdv) Then
      Send("<input type=""hidden"" id=""DisallowEditAdvOpt"" name=""DisallowEditAdvOpt"" value=""1"" />")
    Else
      Send("<input type=""hidden"" id=""DisallowEditAdvOpt"" name=""DisallowEditAdvOpt"" value=""0"" />")
    End If
    If (bUseTemplateLocks And bDisallowEditLimit) Then
      Send("<input type=""hidden"" id=""DisallowEditLimitOpt"" name=""DisallowEditLimitOpt"" value=""1"" />")
    Else
      Send("<input type=""hidden"" id=""DisallowEditLimitOpt"" name=""DisallowEditLimitOpt"" value=""0"" />")
    End If
    If MyCommon.Fetch_SystemOption(97) = "1" Then
      Send("<input type=""hidden"" id=""NumericOnly"" name=""NumericOnly"" value=""true"" />")
    End If
  %>
  <div id="controls">
    <% If (IsTemplate) Then%>
    <span class="temp">
      <input type="checkbox" class="tempcheck" id="temp-employees" name="Disallow_Edit"
        <% if(disallow_edit)then sendb(" checked=""checked""") %> />
      <label for="temp-employees">
        <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
      </label>
    </span>
    <% End If%>
    <%If Not (IsTemplate) Then
        If (Logix.UserRoles.EditOffer And Not (bUseTemplateLocks And Disallow_Edit)) Then Send_Save()
      Else
        If (Logix.UserRoles.EditTemplates) Then Send_Save()
      End If
    %>
  </div>
</div>
<div id="main">
    <%Send("<div id=""infobar1"" >")
    Send("</div>")  %>
  <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <% If MyCommon.Fetch_SystemOption(222) = "0" Or bStaticPG Then%>
      <% If bStaticPG Then ByExistingPGSelector = True%>
  <input type="radio" id="pgselectortype1" name="pgselectortype" onclick="javascript:ProductGroupTypeSelection();"
    <% if(ByExistingPGSelector) then sendb(" checked=""checked""") %> value="existingadd" /><label
      for="pgselectortype1"><% Sendb(Copient.PhraseLib.Lookup("term.selectexistingprodgroups", LanguageID))%></label>
  <input type="radio" id="pgselectortype2" name="pgselectortype" onclick="javascript:ProductGroupTypeSelection();"
    <%
     if(Not ByExistingPGSelector) then sendb(" checked=""checked""")
     if (bStaticPG) then sendb(" disabled=""disabled""")
    %>
     value="directadd" /><label for="pgselectortype2">Add products to reward</label>
  <%Else%>
  <input type="radio" id="pgselectortype2" name="pgselectortype" onclick="javascript:ProductGroupTypeSelection();"
    <% if(Not ByExistingPGSelector) then sendb(" checked=""checked""") %> value="directadd" /><label
      for="pgselectortype2">Add products to reward</label>
  <input type="radio" id="pgselectortype1" name="pgselectortype" onclick="javascript:ProductGroupTypeSelection();"
    <% if(ByExistingPGSelector) then sendb(" checked=""checked""") %> value="existingadd" /><label
      for="pgselectortype1"><% Sendb(Copient.PhraseLib.Lookup("term.selectexistingprodgroups", LanguageID))%></label>
  <%End If%>
  <div id="columnfull">
    <div class="box" id="directprodaddselector" <% Sendb(IIf(ByExistingPGSelector = True, " style=""display: none; overflow:auto;""", " style=""overflow:auto;""")) %>>
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("pgroup-edit.addremove", LanguageID))%>
        </span>
      </h2>
      <span>
        <label for="modprodgrpname">
          Product group name:</label><br />
        <input type="text" id="modprodgroupname" style="width: 347px;" name="modprodgroupname"
          maxlength="200" value="<% Sendb(GName) %>" /><br />
      </span>
      <br class="half" />
      <div style="float: left; width: 310px;">
        <span style="position: relative">
          <%
            If (ShowAllItems OrElse GroupSize <= 100) AndAlso rstItems IsNot Nothing Then
              Sendb(Copient.PhraseLib.Lookup("pgroup-edit.all-items-note", LanguageID) & " (" & rstItems.Rows.Count & " ")
              If (rstItems.Rows.Count = 1) Then
                Sendb(Copient.PhraseLib.Lookup("term.product", LanguageID).ToString.ToLower & ")<br />")
              Else
                Sendb(Copient.PhraseLib.Lookup("term.products", LanguageID).ToString.ToLower & ")<br />")
              End If
            Else
              Sendb(Copient.PhraseLib.Lookup("pgroup-edit.listnote", LanguageID) & "<br />")
            End If
          %>
          <%
            Sendb(Copient.PhraseLib.Lookup("term.contains", LanguageID) & " ")
            If (ProductsWithoutDesc = 1) Then
              Response.Write(ProductsWithoutDesc & " ")
              Sendb(StrConv(Copient.PhraseLib.Lookup("term.product", LanguageID) & " ", VbStrConv.Lowercase))
              Sendb(StrConv(Copient.PhraseLib.Lookup("term.withoutdesc", LanguageID), VbStrConv.Lowercase))
            Else
              Response.Write(ProductsWithoutDesc & " ")
              Sendb(StrConv(Copient.PhraseLib.Lookup("term.products", LanguageID) & " ", VbStrConv.Lowercase))
              Sendb(StrConv(Copient.PhraseLib.Lookup("term.withoutdesc", LanguageID), VbStrConv.Lowercase))
            End If
		  
          %>
        </span>
        <select name="PKID" id="PKID" size="15" multiple="multiple" onscroll="handlePageClick(this);"
          class="longer" style="width: 290px;">
          <%
            descriptionItem = String.Empty
            If (GroupSize > 0) Then
              For Each row4 As DataRow In rstItems.Rows
                descriptionItem = MyCommon.NZ(row4.Item("ExtProductID"), " ") & " " & MyCommon.NZ(row4.Item("Description"), " ") & "-"
                If MyCommon.NZ(row4.Item("PhraseID"), 0) > 0 Then
                  descriptionItem &= Copient.PhraseLib.Lookup(MyCommon.NZ(row4.Item("PhraseID"), 0), LanguageID)
                Else
                  If MyCommon.NZ(row4.Item("ProductType"), "") <> "" Then
                    descriptionItem &= row4.Item("ProductType")
                  Else
                    descriptionItem &= Copient.PhraseLib.Lookup("term.unknown", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.type", LanguageID), VbStrConv.Lowercase) & " " & MyCommon.NZ(row4.Item("ProductTypeID"), 0)
                  End If
                End If
                Send("<option value=""" & row4.Item("PKID") & """>" & descriptionItem & "</option>")
              Next
            End If
          %>
        </select>
        <br />
        <%
          'If (Not IsSpecialGroup OrElse (IsSpecialGroup And CanEditSpecialGroup)) Then
          If (Logix.UserRoles.EditProductGroups) Then
            Send("    <br class=""half"" /><input type=""submit"" class=""large"" id=""remove"" name=""remove"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("confirm.remove", LanguageID) & "')){}else{return false}"" style=""width:150px;"" value=""" & Copient.PhraseLib.Lookup("term.removefromlist", LanguageID) & """ />")
          End If
          'End If
          'If (Not ShowAllItems AndAlso GroupSize > 100) Then
          '    Send("<input class=""regular"" id=""btnShowAll"" name=""btnShowAll"" type=""button"" value=""" & Copient.PhraseLib.Lookup("list.showall", LanguageID) & """ onclick=""submitShowAll();"" />")
          'End If
        %>
        <br />
      </div>
      <div style="margin-left: 310px;">
        <table cellpadding="1" cellspacing="1">
          <tr>
            <td>
              <% Sendb(Copient.PhraseLib.Lookup("term.type", LanguageID))%>:<br />
              <select id="producttype" name="producttype">
                <%
                  'BZ2079: UE-feature-removal #: Remove unsupported product types for UE (Mix/Match Code, Manufacturer Family code, Pool Code)
                  '        To restore previous functionality: remove the all code in the If statement checking engines except the query without a where clause.
                  
                  MyCommon.QueryStr = "select ProductTypeID,PhraseID from ProductTypes with (NoLock)"
                  
                  rst2 = MyCommon.LRT_Select
                  For Each row3 As DataRow In rst2.Rows
                    ProductTypeID = row3.Item("ProductTypeID")
                    Send("     <option value=""" & row3.Item("ProductTypeID") & """>" & Copient.PhraseLib.Lookup(row3.Item("PhraseID"), LanguageID) & "</option>")
                  Next
                %>
              </select>
            </td>
          </tr>
        </table>
        <br />
        <input type="radio" id="prodaddselector1" name="prodaddselector" onclick="javascript:ProductAddSelection();"
          <% if(ByAddSingleProduct) then sendb(" checked=""checked""") %> value="prodadd" /><label
            for="prodaddselector1"><% Sendb(Copient.PhraseLib.Lookup("gen.addsingleproduct", LanguageID))%></label>
        <input type="radio" id="prodaddselector2" name="prodaddselector" onclick="javascript:ProductAddSelection();"
          <% if(Not ByAddSingleProduct) then sendb(" checked=""checked""") %> value="prodlistadd" /><label
            for="prodaddselector2"><% Sendb(Copient.PhraseLib.Lookup("gen.addproductlist", LanguageID))%></label>
        <br />
        <br />
        <div id="addsingleproduct" <%Sendb(IIf(ByAddSingleProduct, "", " style=""display:none;"""))%>>
          <% Integer.TryParse(MyCommon.Fetch_SystemOption(52), IDLength)%>
          <% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>:<br />
          <input type="text" id="productid" maxlength="19" name="ExtProductID" <% Sendb(If(IDLength > 0, " maxlength=""" & IDLength & """", ""))%>
            style="width: 137px;" value="" />
          <br />
          <br />
          <label for="productdesc">
            <% Sendb(Copient.PhraseLib.Lookup("term.description", LanguageID))%>:</label><br />
          <input type="text" id="productdesc" style="width: 347px;" name="productdesc" maxlength="200"
            value="" /><br />
          <br class="half" />
          <% If (Logix.UserRoles.EditProductGroups) Then%>
          <div style="float: left;">
            <input type="submit" class="large" id="add" name="add" value="<% Sendb(Copient.PhraseLib.Lookup("term.add", LanguageID)) %>"
              onclick="return isValidID();" /></div>
          <div style="float: right; margin-right: 20px">
            <input type="submit" class="large" id="mremove" name="mremove" onclick="if(confirm('<% Sendb(Copient.PhraseLib.Lookup("confirm.remove", LanguageID)) %>')){}else{return false}"
              style="width: 150px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.removemanually", LanguageID)) %>" /></div>
          <br />
          <% End If%>
        </div>
        <div id="addproductlist" <%Sendb(IIf(Not ByAddSingleProduct, "", " style=""display:none;"""))%>>
          <%
            Sendb("<textarea name=""pasteproducts"" id=""pasteproducts"" style=""width: 290px; height: 150px"">")
            Sendb("</textarea>")
            Send("<br />")
            Send("<br />")
          
            Sendb("<input type=""radio"" name=""modifyoperation"" value=""0"" checked=""checked"" />")
            Send("<label for=""operation4"">" & Copient.PhraseLib.Lookup("term.FullReplace", LanguageID) & "</label>&nbsp;&nbsp;")
            Sendb("<input type=""radio"" name=""modifyoperation""  value=""1""  />")
            Send("<label for=""operation5"">" & Copient.PhraseLib.Lookup("term.AddToGroup", LanguageID) & "</label>&nbsp;&nbsp;")
            Sendb("<input type=""radio"" name=""modifyoperation"" value=""2""  />")
            Send("<label for=""operation6"">" & Copient.PhraseLib.Lookup("term.removefromgroup", LanguageID) & "</label>")
            Send("<br />")
          %>
          <br />
          <%
            If (Logix.UserRoles.EditProductGroups) Then
              Send("     <input type=""button"" class=""regular"" id=""modifyprodgroup"" name=""modifyprodgroup"" value=""Apply Changes"" onclick=""javascript:ModifyGroup();"" />")
              Send("     <br />")
            End If
          %>
        </div>
      </div>
      <hr class="hidden" />
    </div>
    <div class="box" id="selector" <% Sendb(IIf(ByExistingPGSelector = False, " style=""display: none; overflow:auto;""", " style=""overflow:auto;""")) %>>
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.productcondition", LanguageID))%>
        </span>
        <% If (IsTemplate) Then%>
        <span class="tempRequire">
          <input type="checkbox" class="tempcheck" id="DisallowEditPg1" name="DisallowEditPg"
            <% if(bDisallowEditPg)then send(" checked=""checked""") %> />
          <label for="temp-Tiers">
            <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
        </span>
        <% ElseIf (bUseTemplateLocks And bDisallowEditPg) Then%>
        <span class="tempRequire">
          <input type="checkbox" class="tempcheck" id="DisallowEditPg2" name="DisallowEditPg"
            disabled="disabled" checked="checked" />
          <label for="temp-Tiers">
            <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
        </span>
        <% End If%>
      </h2>
      <%
        If Not (Logix.UserRoles.EditOffer And Not (bUseTemplateLocks And Disallow_Edit)) Then
          DisabledAttribute = " disabled=""disabled"""
        End If
      %>
      <input type="radio" id="functionradio1" name="functionradio" <% if(MyCommon.Fetch_SystemOption(175)= "1" ) then sendb(" checked=""checked""") %>
        <% sendb(disabledattribute) %> />
      <label for="functionradio1">
        <% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
      <input type="radio" id="functionradio2" name="functionradio" <% if(MyCommon.Fetch_SystemOption(175)= "2" ) then sendb(" checked=""checked""") %>
        <% sendb(disabledattribute) %> />
      <label for="functionradio2">
        <% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
      <input type="text" class="medium" id="functioninput" name="functioninput" maxlength="100"
        onkeyup="javascript:xmlPostTimer('OfferFeeds.aspx','ProductGroupsCM');" value=""
        <% sendb(disabledattribute) %> />
      <% If (bCreateGroupOrProgramFromOffer AndAlso Logix.UserRoles.CreateProductGroups) Then%>
      <input class="regular" name="btncreate" id="btncreate" type="button" value="<% Sendb(Copient.PhraseLib.Lookup("term.create", LanguageID))%>"
        onclick="javascript:handleCreateClick('btncreate');" />
      <% End If%>
      <br />
      <div id="searchLoadDiv" style="display: block;">
        &nbsp;</div>
      <div id="pgList" class="column3x1">
        <select class="long" id="functionselect" name="functionselect" size="20" <% sendb(disabledattribute) %>>
          <%
            Dim orderBy As String = ""
            If (MyCommon.Fetch_SystemOption(235) = "1") Then
              orderBy = " order by AnyProduct desc, Name"
            Else
              orderBy = " order by AnyProduct desc, ProductGroupID desc, Name asc"
            End If
                        
            Dim Limiter As String = String.Empty
            If (ExcludedItem) Then Limiter = "and pg.ProductGroupID <> " & ExcludedItem
            If (SelectedItem) Then Limiter = Limiter & " and pg.ProductGroupID <> " & SelectedItem

            If bStoreUser Then
              sJoin = "Full Outer Join ProductGroupLocUpdate pglu with (NoLock) on pg.ProductGroupID=pglu.ProductGroupID "
              wherestr = " and (LocationID in (" & sValidLocIDs & ") or (CreatedByAdminID in (" & sValidSU & ") and Isnull(LocationID,0)=0)) and AnyProduct=0"
            End If
            MyCommon.QueryStr = "select " & topString & " pg.ProductGroupID,CreatedDate,Name,PhraseId,LastUpdate,AnyProduct from ProductGroups pg with (NoLock) " & sJoin & " where Deleted=0" & Limiter & wherestr
            If (bEnableRestrictedAccessToUEOfferBuilder) Then MyCommon.QueryStr &= " and isnull(TranslatedFromOfferID,0) = 0 "

            MyCommon.QueryStr &= orderBy
            rst = MyCommon.LRT_Select
                        
            For Each row In rst.Rows
              If MyCommon.NZ(row.Item("ProductGroupID"), 0) = 1 AndAlso (ExcludedItem <> 1 AndAlso SelectedItem <> 1) Then
                Send("<option value=""1"" style=""color: brown; font-weight: bold;"">" & Copient.PhraseLib.Lookup("term.anyproduct", LanguageID) & "</option>")
              Else
                Sendb("<option value=""" & row.Item("ProductGroupID") & """ title=""" & row.Item("Name") & """>")
                If (MyCommon.NZ(row.Item("PhraseID"), 0) > 0) Then
                  Sendb(Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID))
                Else
                  Sendb(MyCommon.NZ(row.Item("Name"), ""))
                End If
                Send("</option>")
              End If
            Next
          %>
        </select>
      </div>
      <%If (RECORD_LIMIT > 0) Then
          If (MyCommon.Fetch_SystemOption(235) = "1") Then
            Send(Copient.PhraseLib.Lookup("groups.displayname", LanguageID) & ": " & GroupRecordLimit.ToString() & "<br />")
          Else
            Send(Copient.PhraseLib.Lookup("groups.display", LanguageID) & ": " & GroupRecordLimit.ToString() & "<br />")
          End If
        End If
      %>
      <div class="column3x2">
        <center>
          <br />
          <br />
          <input type="button" class="regular select" id="select1" name="select1" value="<% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID))%> &#9658;"
            onclick="select1_onclick();" <% sendb(disabledattribute) %> />
          <br />
          <br />
          <input type="button" class="regular deselect" id="deselect1" name="deselect1" value="&#9668; <% Sendb(Copient.PhraseLib.Lookup("term.deselect", LanguageID))%>"
            disabled="disabled" onclick="deselect1_onclick();" /><br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <input type="button" class="regular select" name="select2" id="select2" value="<% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID))%> &#9658;"
            disabled="disabled" onclick="select2_onclick();" <% sendb(disabledattribute) %> />
          <br />
          <br />
          <input type="button" class="regular deselect" name="deselect2" id="deselect2" value="&#9668; <% Sendb(Copient.PhraseLib.Lookup("term.deselect", LanguageID))%>"
            disabled="disabled" onclick="deselect2_onclick();" <% sendb(disabledattribute) %> /><br />
        </center>
      </div>
      <br />
      <div class="column3x3">
        <div class="graybox">
          <h3>
            <span>
              <% Sendb(Copient.PhraseLib.Lookup("term.selectedproducts", LanguageID))%>
            </span>
          </h3>
          <select class="long" id="selected" name="selected" size="7" <% sendb(disabledattribute) %>>
            <%
              MyCommon.QueryStr = "select OFR.ProductGroupID,ExcludedProdGroupID,C.Name from OfferRewards as OFR with (NoLock) join ProductGroups as C with (NoLock) on OFR.ProductGroupID=C.ProductGroupID and RewardID=" & RewardID & ";"
              rst = MyCommon.LRT_Select
              If (rst.Rows.Count = 0) Then
                TransactionLevelSelected = True
              End If
              For Each row In rst.Rows
                If MyCommon.NZ(row.Item("ProductGroupID"), 0) = 1 Then
                  Send("  <option style=""font-weight:bold;color:brown;"" value=""" & row.Item("ProductGroupID") & """>" & row.Item("Name") & "</option>")
                Else
                  Send("  <option value=""" & row.Item("ProductGroupID") & """>" & row.Item("Name") & "</option>")
                End If
                SelectedItem = row.Item("ProductGroupID")
              Next
            %>
          </select>
        </div>
        <br />
        <br class="half" />
        <div class="graybox">
          <h3>
            <span>
              <% Sendb(Copient.PhraseLib.Lookup("term.excludedproducts", LanguageID))%>
            </span>
          </h3>
          <select class="long" id="excluded" name="excluded" size="7" <% sendb(disabledattribute) %>>
            <%
              MyCommon.QueryStr = "select OFR.ProductGroupID,OFR.ExcludedProdGroupID,C.Name from OfferRewards as OFR with (NoLock) join ProductGroups as C with (NoLock) on OFR.ExcludedProdGroupID=C.ProductGroupID  where not(ExcludedProdGroupID=0) and RewardID=" & RewardID & ";"
              rst = MyCommon.LRT_Select
              For Each row In rst.Rows
                Send("<option value=""" & row.Item("ExcludedProdGroupID") & """>" & row.Item("Name") & "</option>")
                ExcludedItem = row.Item("ExcludedProdGroupID")
              Next
            %>
          </select>
        </div>
        <hr class="hidden" />
      </div>
    </div>
    <%        
      TransactionLevelPossible = TransactionLevelSelected
    %>
  </div>
  <div id="column1">
    <div class="box" id="distribution">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.distribution", LanguageID))%>
        </span>
        <% If (IsTemplate) Then%>
        <span class="tempRequire">
          <input type="checkbox" class="tempcheck" id="DisallowEditDist1" name="DisallowEditDist"
            <% if(bDisallowEditDist)then send(" checked=""checked""") %> />
          <label for="temp-Tiers">
            <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
          </label>
        </span>
        <% ElseIf (bUseTemplateLocks And bDisallowEditDist) Then%>
        <span class="tempRequire">
          <input type="checkbox" class="tempcheck" id="DisallowEditDist2" name="DisallowEditDist"
            disabled="disabled" checked="checked" />
          <label for="temp-Tiers">
            <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
          </label>
        </span>
        <% End If%>
      </h2>
      <label for="discounttype" style="position: relative;">
        <% Sendb(Copient.PhraseLib.Lookup("term.discount", LanguageID))%>
        :</label>
      <select id="discounttype" name="discounttype" onchange="checkConditionState();priorityReset();"
        style="position: relative;" <% If(bUseTemplateLocks and bDisallowEditDist) Then sendb(" disabled=""disabled""") %>>
        <% If (StoredValueSelected Or (Not TransactionLevelPossible)) Then%>
        <option value="1" <%if(rewardamounttypeid=1)then sendb(" selected=""selected""")%>>
          <% Sendb(Copient.PhraseLib.Lookup("reward.discount-amountoffitem", LanguageID))%>
        </option>
        <option value="2" <%if(rewardamounttypeid=2)then sendb(" selected=""selected""")%>>
          <% Sendb(Copient.PhraseLib.Lookup("reward.discount-percentoffitem", LanguageID))%>
        </option>
          <% If (Not bBuckParentOffer) Then%>
        <option value="3" <%if(rewardamounttypeid=3)then sendb(" selected=""selected""")%>>
          <% Sendb(Copient.PhraseLib.Lookup("reward.discount-amountoffweight", LanguageID))%>
        </option>
        <option value="4" <%if(rewardamounttypeid=4)then sendb(" selected=""selected""")%>>
          <% Sendb(Copient.PhraseLib.Lookup("reward.discount-amountoffvolume", LanguageID))%>
        </option>
        <option value="5" <%if(rewardamounttypeid=5)then sendb(" selected=""selected""")%>>
          <% Sendb(Copient.PhraseLib.Lookup("reward.discount-pricepoint", LanguageID))%>
        </option>
        <option value="6" <%if(rewardamounttypeid=6)then sendb(" selected=""selected""")%>>
          <% Sendb(Copient.PhraseLib.Lookup("reward.discount-pricepointweight", LanguageID))%>
        </option>
        <option value="7" <%if(rewardamounttypeid=7)then sendb(" selected=""selected""")%>>
          <% Sendb(Copient.PhraseLib.Lookup("reward.discount-freeitem", LanguageID))%>
        </option>
        <% If (StoredValueSelected Or ((Not TransactionLevelPossible) And (Not Tiered))) Then%>
        <option value="10" <%if(rewardamounttypeid=10)then sendb(" selected=""selected""")%>>
          <% Sendb(Copient.PhraseLib.Lookup("term.storedvalue", LanguageID))%>
        </option>
          <% End If%>
        <% End If%>
        <% Else%>
        <option value="8" <%if(rewardamounttypeid=8)then sendb(" selected=""selected""")%>>
          <% Sendb(Copient.PhraseLib.Lookup("reward.discount-amountofftrans", LanguageID))%>
        </option>
        <option value="9" <%if(rewardamounttypeid=9)then sendb(" selected=""selected""")%>>
          <% Sendb(Copient.PhraseLib.Lookup("reward.discount-percentofftrans", LanguageID))%>
        </option>
        <% End If%>
      </select>
      <br />
      <div id="DivAmount" style="display: <% Sendb(IIF((StoredValueSelected or FreeItemSelected), "none", "block"))%>;">
        <%
          If (bUseTemplateLocks And bDisallowEditDist) Then
            sDisabled = " disabled=""disabled"""
          Else
            sDisabled = ""
          End If

          MyCommon.QueryStr = "select CT.RewardID,Tiered,O.Numtiers,O.OfferID,CT.TierLevel,CT.RewardAmount,UseSpecialPricing from Offerrewards as OC with (NoLock) " & _
                              "left join Offers as O with (NoLock) on O.OfferID=OC.OfferID " & _
                              "left join RewardTiers as CT with (NoLock) on OC.RewardID=CT.RewardID " & _
                              "where OC.RewardID=" & RewardID
          rst = MyCommon.LRT_Select()
          NumTiers = MyCommon.NZ(rst.Rows(0).Item("Numtiers"), 0)
          q = 1
          For Each row In rst.Rows
            If (row.Item("UseSpecialPricing") <> True) Then
              If (row.Item("Tiered") = False) Then
                  Send("<label for=""tier0"" style=""position:relative;"">" & Copient.PhraseLib.Lookup("term.amount", LanguageID) & ":</label> &nbsp; <input class=""shorter"" id=""tier0"" name=""tier0"" type=""text"" maxlength=""9"" value=""" & MyCommon.NZ(row.Item("RewardAmount"), "0.000") & """" & sDisabled & " /><br />")
              Else
                  Send("<label for=""tier" & q & """ style=""position:relative;"">" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & q & ":</label> <input class=""shorter"" id=""tier" & q & """ name=""tier" & q & """ type=""text"" onchange=""validateNumeric("& rst.Rows.Count & ")"" maxlength=""9"" value=""" & MyCommon.NZ(row.Item("RewardAmount"), "0.000") & """" & sDisabled & " />")
              End If
            End If
            q = q + 1
          Next
          Send("<input type=""hidden"" name=""NumTiers"" value=""" & NumTiers & """ />")
        %>
      </div>
      <%
        If TransactionLevelPossible Then
          If (bUseTemplateLocks And bDisallowEditDist) Then
            sDisabled = " disabled=""disabled"""
          Else
            sDisabled = ""
          End If
        End If
      %>
      <div id="nontransoptions" style="display: <% Sendb(IIF(TransactionLevelPossible, "none", "block"))%>;
        position: relative;">
        <input id="triggerbogo" name="trigger" type="radio" <%if(valueradio=1)then sendb(" checked=""checked""") %>
            value="1" onclick="handleProrateUnchecked();" <% Sendb(sDisabled)%> />
        <label for="triggerbogo">
          <% Sendb(Copient.PhraseLib.Lookup("reward.discountevery", LanguageID))%>
        </label>
        <br />
        &nbsp; &nbsp; &nbsp; &nbsp;
        <label for="Xbox">
          <% Sendb(Copient.PhraseLib.Lookup("term.mustpurchase", LanguageID))%>
        </label>
        <input class="shortest" id="Xbox" name="Xbox" maxlength="9" type="text" <%if(valueradio=1)then sendb(" value=""" & triggerqty & """ ") %><% Sendb(sDisabled)%> />
        <% Sendb(Copient.PhraseLib.Lookup("term.item(s)", LanguageID))%>
        <br />
        <br class="half" />
        <input id="triggerbxgy" name="trigger" type="radio" value="2" onclick="handleProrateUnchecked();" <%if(valueradio=2)then sendb(" checked=""checked""") %><% Sendb(sDisabled)%> />
        <label for="triggerbxgy">
          <% Sendb(Copient.PhraseLib.Lookup("term.buy", LanguageID))%>
        </label>
        <input class="shortest" id="bxgy1" name="Xbox2" maxlength="9" type="text" <%if(valueradio=2)then sendb(" value=""" & triggerqty-applytolimit & """ ") %><% Sendb(sDisabled)%> />
        <% Sendb(Copient.PhraseLib.Lookup("term.item(s)", LanguageID))%>
        ,
        <% Sendb(Copient.PhraseLib.Lookup("term.get", LanguageID))%>
        <input class="shortest" id="bxgy2" name="Ybox2" maxlength="9" type="text" <%if(valueradio=2)then sendb(" value=""" & applytolimit & """ ") %><% Sendb(sDisabled)%> />
        <% Sendb(Copient.PhraseLib.Lookup("term.discount(s)", LanguageID))%>
        <br />
        <br class="half" />
          <%
            If BestDeal Then
              sDisabled = " disabled=""disabled"""
            End If
          %>
          <input id="triggerprorate" name="trigger" type="radio" value="3" onclick="handleProrateClicked();" <%if(valueradio=3)then sendb(" checked=""checked""") %><% Sendb(sDisabled)%> />
        <label for="triggerprorate">
          <% Sendb(Copient.PhraseLib.Lookup("reward.discountprorate", LanguageID))%>
        </label>
        <input class="shortest" id="prorate" name="Xbox3" maxlength="9" type="text" <%if(valueradio=3)then sendb(" value=""" & triggerqty & """ ") %><% Sendb(sDisabled)%> />
        <% Sendb(Copient.PhraseLib.Lookup("term.item(s)", LanguageID))%>
      </div>
      <hr class="hidden" />
    </div>
    <div class="box" id="department">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.chargebackdepartment", LanguageID))%>
        </span>
        <% If (IsTemplate) Then%>
        <span class="tempRequire">
          <input type="checkbox" class="tempcheck" id="DisallowEditDept1" name="DisallowEditDept"
            <% if(bDisallowEditDept)then send(" checked=""checked""") %> />
          <label for="temp-Tiers">
            <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
          </label>
        </span>
        <% ElseIf (bUseTemplateLocks And bDisallowEditDept) Then%>
        <span class="tempRequire">
          <input type="checkbox" class="tempcheck" id="DisallowEditDept2" name="DisallowEditDept"
            disabled="disabled" checked="checked" />
          <label for="temp-Tiers">
            <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
          </label>
        </span>
        <% End If%>
      </h2>
      <select id="dept" name="dept" class="longer" <% If(bUseTemplateLocks and bDisallowEditDept) Then sendb(" disabled=""disabled""") %>>
        <%
          If (BannersEnabled) Then
            MyCommon.QueryStr = "select BO.BannerID, BAN.AllBanners from BannerOffers BO with (NoLock) " & _
                                "inner join  Banners BAN with (NoLock) on BAN.BannerID = BO.BannerID and BAN.Deleted=0 " & _
                                "where OfferID = " & OfferID
            rst = MyCommon.LRT_Select
            AllBanners = (rst.Rows.Count = 1 AndAlso MyCommon.NZ(rst.Rows(0).Item("AllBanners"), False))
              
            sQuery = "Select * from ChargeBackDepts with (NoLock) "
            If (rst.Rows.Count = 1 And Not AllBanners) Then
              sQuery &= " where ChargebackDeptID in (0) and ChargebackDeptID <> 10 or BannerID = " & MyCommon.NZ(rst.Rows(0).Item("BannerID"), -1) & " "
            Else
              sQuery &= " where ChargebackDeptID in (0) and ChargebackDeptID <> 10 or ((BannerID = 0 or BannerID IS NULL) and ChargebackDeptID<>10) "
            End If
            MyCommon.QueryStr = sQuery
          Else
            MyCommon.QueryStr = "Select * from ChargeBackDepts with (NoLock) where ChargebackDeptID <> 10"
          End If
          MyCommon.QueryStr &= " Order By ExternalID "
          rst = MyCommon.LRT_Select

          For Each row In rst.Rows
            Sendb("<option value=""" & row.Item("ChargeBackDeptID") & """")
            If (ChargeBackDeptID = MyCommon.Extract_Val(MyCommon.NZ(row.Item("ChargeBackDeptID"), 0))) Then
              Sendb(" selected=""selected""")
            End If
            Sendb(">")
            If ((row.Item("ExternalID") = "") Or (row.Item("ExternalID") = "0")) Then
            Else
              Sendb(row.Item("ExternalID") & " - ")
            End If
            If (IsDBNull(row.Item("PhraseID"))) Then
              Sendb(row.Item("Name"))
            Else
              If (row.Item("PhraseID") = 0) Then
                Sendb(row.Item("Name"))
              Else
                Sendb(Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID))
              End If
            End If
            Sendb("</option>")
          Next
        %>
      </select>
      <br />
      <hr class="hidden" />
    </div>
    <div class="box" id="specialprice" style="display: <% Sendb(IIF(StoredValueSelected, "none", "block"))%>;">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.specialpricing", LanguageID))%>
        </span>
        <% If (IsTemplate Or (bUseTemplateLocks And bDisallowEditSpc)) Then%>
        <span class="tempRequire">
          <input type="checkbox" class="tempcheck" id="DisallowEditSpc" name="DisallowEditSpc"
            <% if(bDisallowEditSpc)then send(" checked=""checked""") %><% If(bUseTemplateLocks) Then sendb(" disabled=""disabled""") %> />
          <label for="temp-Tiers">
            <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
          </label>
        </span>
        <br class="printonly" />
        <% End If%>
      </h2>
      <%
        If (Tiered) Then
          sDisabled = " disabled=""disabled"""
        Else
          If (bUseTemplateLocks And bDisallowEditSpc) Then
            sDisabled = " disabled=""disabled"""
          Else
            sDisabled = ""
          End If
        End If
      %>
      <input type="hidden" name="clicksent" value="false" />
      <input class="checkbox" id="specialpricing" name="usespecialpricing" type="checkbox"
        onclick="handleClickSent();" <% sendb(sDisabled) %><%if (usespecialpricing=true) then sendb(" checked=""checked""")%> />
      <label for="specialpricing">
        <% Sendb(Copient.PhraseLib.Lookup("reward.spenable", LanguageID))%>
      </label>
      <br />
      <br class="half" />
      <%
        If (bUseTemplateLocks And bDisallowEditSpc) Then
          sDisabled = " disabled=""disabled"""
        Else
          sDisabled = ""
        End If
        MyCommon.QueryStr = "select CT.RewardID,Tiered,OC.SPRepeatAtOccur,O.Numtiers,O.OfferID,CT.TierLevel,CT.RewardAmount,UseSpecialPricing from Offerrewards as OC with (NoLock) " & _
                            "left join Offers as O with (NoLock) on O.OfferID=OC.OfferID left  join RewardTiers as CT with (NoLock) on OC.RewardID=CT.RewardID where OC.RewardID=" & RewardID
        rst = MyCommon.LRT_Select()
        q = 1
        If rst.Rows.Count > 0 Then
          Send("<table cellpadding=""0"" cellspacing=""0"" summary=""" & Copient.PhraseLib.Lookup("term.specialpricing", LanguageID) & """>")
          For Each row In rst.Rows
            If (row.Item("UseSpecialPricing") = True) Then
              If (row.Item("SPRepeatAtOccur") = q Or row.Item("SPRepeatAtOccur") = 0) Then
                Send("<tr>")
                Send("    <td>" & q & "</td>")
                Send("    <td><input type=""radio"" id=""repeat" & q & """ name=""repeat"" style=""margin: 0px;"" value=""" & q & """ checked=""checked""" & sDisabled & " /></td>")
              Else
                Send("<tr>")
                Send("    <td>" & q & "</td>")
                Send("    <td><input type=""radio"" id=""repeat" & q & """ name=""repeat"" style=""margin: 0px;"" value=""" & q & """" & sDisabled & " /></td>")
              End If
              Dim Key As String = "tier" & q
              If Request.QueryString(Key) <> "" Then
                row.Item("RewardAmount") = Request.QueryString(Key)
              End If
                Send("    <td><input class=""shorter"" maxlength=""9"" id=""tier"  & q & """ name=""tier" & q & """ type=""text"" onchange=""validateNumeric("& rst.Rows.Count & ")"" value=""" & MyCommon.NZ(row.Item("RewardAmount"), "0.000") & """" & sDisabled & " />")
              If (q = rst.Rows.Count And q > 1) Then
                Send("    <input type=""submit"" class=""ex"" name=""deletespecial"" value=""X"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & sDisabled & " />")
              End If
              If (q = Request.QueryString("FlagTier")) Then
                Send("    <td style=""font-weight:bold;color:red;font-size:8pt;"">" & q & "" & " INVALID" & "</td>")
              End If
              Send("    </td>")
              Sendb("</tr>")
            End If
            q = q + 1
          Next
          Send("</table>")
        End If
            
        If (UseSpecialPricing = False) Then
          sDisabled = " disabled=""disabled"""
        Else
          If (bUseTemplateLocks And bDisallowEditSpc) Then
            sDisabled = " disabled=""disabled"""
          Else
            sDisabled = ""
          End If
        End If
      %>
      <input id="addvalue" name="addvalue" type="submit" value="<% Sendb(Copient.PhraseLib.Lookup("term.addvalue", LanguageID)) %>"
        <% sendb(sDisabled) %> /><br />
      <p>
        <% Sendb(Copient.PhraseLib.Lookup("reward.spnote", LanguageID))%>
      </p>
      <hr class="hidden" />
    </div>
  </div>
  <div id="gutter">
  </div>
  <div id="column2">
    <div class="box" id="programs" style="display: <% Sendb(IIF(StoredValueSelected, "block", "none"))%>;">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.program", LanguageID))%>
        </span>
        <% If (IsTemplate) Then%>
        <span class="tempRequire">
          <input type="checkbox" class="tempcheck" id="DisallowEditPp1" name="DisallowEditPp"
            <% if(bDisallowEditPp)then send(" checked=""checked""") %> />
          <label for="temp-Tiers">
            <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
          </label>
        </span>
        <% ElseIf (bUseTemplateLocks And bDisallowEditPp) Then%>
        <span class="tempRequire">
          <input type="checkbox" class="tempcheck" id="DisallowEditPp2" name="DisallowEditPp"
            disabled="disabled" checked="checked" />
          <label for="temp-Tiers">
            <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
          </label>
        </span>
        <% End If%>
      </h2>
      <input type="radio" id="Pfunctionradio1" name="Pfunctionradio" <% if(MyCommon.Fetch_SystemOption(175)= "1" ) then sendb(" checked=""checked""") %>
        <% sendb(disabledattribute) %> /><label for="Pfunctionradio1"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
      <input type="radio" id="Pfunctionradio2" name="Pfunctionradio" <% if(MyCommon.Fetch_SystemOption(175)= "2" ) then sendb(" checked=""checked""") %>
        <% sendb(disabledattribute) %> /><label for="Pfunctionradio2"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
      <input class="medium" onkeyup="PhandleKeyUp(200);" id="Pfunctioninput" name="Pfunctioninput"
        type="text" maxlength="100" value="" <% sendb(disabledattribute) %> /><br />
      <select class="longer" id="Pfunctionselect" name="Pfunctionselect" size="10" <% sendb(disabledattribute) %>>
        <%
          MyCommon.QueryStr = "select SVProgramID as ProgramID,Name as ProgramName, SVTypeId from StoredValuePrograms with (NoLock)" & _
                              " where deleted=0 and Visible=1 and SVProgramID is not null and SVTypeId > 1 order by ProgramName"
          rst2 = MyCommon.LRT_Select
          Dim RowSelected As Integer
          If (rst2.Rows.Count > 0) Then
            RowSelected = rst2.Rows(0).Item("ProgramID")
          Else
            RowSelected = 0
          End If
          For Each row2 In rst2.Rows
            Send("<option value=" & row2.Item("ProgramID") & ">" & row2.Item("ProgramName") & "</option>")
          Next
        %>
      </select>
      <br />
      <br class="half" />
      <%
        MyCommon.QueryStr = "select SVP.SVProgramID as ProgramId,SVP.Name as ProgramName, SVP.SVTypeId from StoredValuePrograms as SVP with (NoLock) " & _
                  "inner join CM_RewardStoredValues as RSV with (NoLock) on RSV.ProgramID=SVP.SVProgramID " & _
                  "inner join Discounts as DSC with (NoLock) on DSC.SVLinkID=RSV.RewardStoredValuesID " & _
                  "inner join OfferRewards as OFR with (NoLock) on OFR.LinkID=DSC.DiscountID " & _
                  "where RewardID=" & RewardID & " and SVP.Deleted=0 and OFR.Deleted=0;"
        rst2 = MyCommon.LRT_Select
        Send("<label for=""Pselected""><b>" & Copient.PhraseLib.Lookup("term.selectedprogram", LanguageID) & "</b></label><br />")
        Send("<input class=""regular"" id=""Pselect1"" name=""Pselect1"" type=""button"" value=""&#9660; " & Copient.PhraseLib.Lookup("term.select", LanguageID) & """ onclick=""PhandleSelectClick('Pselect1');""" & IIf(rst2.Rows.Count > 0, " disabled=""disabled""", "") & " />&nbsp;")
        Send("<input class=""regular"" id=""Pdeselect1"" name=""Pdeselect1"" type=""button"" value=""" & Copient.PhraseLib.Lookup("term.deselect", LanguageID) & " &#9650;"" onclick=""PhandleSelectClick('Pdeselect1');""" & IIf(rst2.Rows.Count = 0, " disabled=""disabled""", "") & " /><br />")
        Send("<br class=""half"" />")
        Send("<select class=""longer"" id=""Pselected"" name=""Pselected"" size=""2""" & DisabledAttribute & ">")
        For Each row2 In rst2.Rows
          Send("<option value=""" & row2.Item("ProgramID") & """>" & row2.Item("ProgramName") & "</option>")
        Next
        Send("</select>")
      %>
      <hr class="hidden" />
    </div>
    <div class="box" id="limits" style="display: <% Sendb(IIF(StoredValueSelected, "none", "block"))%>;">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.limits", LanguageID))%>
        </span>
        <% If (IsTemplate Or (bUseTemplateLocks And bDisallowEditLimit)) Then%>
        <span class="tempRequire">
          <input type="checkbox" class="tempcheck" id="DisallowEditLimit" name="DisallowEditLimit"
            <% if(bDisallowEditLimit)then send(" checked=""checked""") %><% If(bUseTemplateLocks) Then sendb(" disabled=""disabled""") %> />
          <label for="temp-Tiers">
            <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
          </label>
        </span>
        <br class="printonly" />
        <% End If%>
      </h2>
      <%
        MyCommon.QueryStr = "Select LimitId, Name from CM_AdvancedLimits with (NoLock) where Deleted=0 and LimitTypeID in (2,3,4) order By Name;"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
      %>
      <label for="selectadv">
        <% Sendb(Copient.PhraseLib.Lookup("term.advlimits", LanguageID))%>:</label>
      <select id="selectadv" name="selectadv" class="longer" onchange="setlimitsection(true);"
        <% If(bUseTemplateLocks and bDisallowEditLimit) Then sendb(" disabled=""disabled""") %>>
        <%
          Sendb("<option value=""0""")
          If (AdvancedLimitID = 0) Then
            Sendb(" selected=""selected""")
          End If
          Sendb(">None</option>")
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
        value="<% sendb(RewardLimit) %>" <% If(bUseTemplateLocks and bDisallowEditLimit) Then sendb(" disabled=""disabled""") %> />
      <select id="limitvaluetype" name="limitvaluetype" <% If(bUseTemplateLocks and bDisallowEditLimit) Then sendb(" disabled=""disabled""") %>>
        <option value="2" <% if(rewardlimittypeid=2)then sendb(" selected=""selected""")%>>
          <% Sendb(Copient.PhraseLib.Lookup("term.amount", LanguageID))%>
        </option>
        <option value="3" <% if (transactionlevelpossible) then sendb(" style=""display:none;""") %><%if(rewardlimittypeid=3)then sendb(" selected=""selected""")%>>
          <% Sendb(Copient.PhraseLib.Lookup("term.weight", LanguageID))%>
        </option>
        <option value="4" <% if (transactionlevelpossible) then sendb(" style=""display:none;""") %><%if(rewardlimittypeid=4)then sendb(" selected=""selected""")%>>
          <% Sendb(Copient.PhraseLib.Lookup("term.volume", LanguageID))%>
        </option>
      </select><% Sendb(Copient.PhraseLib.Lookup("term.per", LanguageID))%><input class="shorter"
        id="limitperiod" name="form_DistPeriod" maxlength="4" type="text" value="<% sendb(DistPeriod) %>"
        <% If(bUseTemplateLocks and bDisallowEditLimit) Then sendb(" disabled=""disabled""") %> />
      <select id="selectday" name="selectday" onchange="setperiodsection(true);" <% If(bUseTemplateLocks and bDisallowEditLimit) Then sendb(" disabled=""disabled""") %>>
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
      <br />
      <% If (Not TransactionLevelPossible) Then%>
      <input class="checkbox" id="limitexceed" name="limitexceed" type="checkbox" onchange="rewardsavecheck();"
        <%if(allownegative=true)then send(" checked=""checked""")%> />
      <label for="limitexceed">
        <% Sendb(Copient.PhraseLib.Lookup("reward.exceeditem", LanguageID))%>
      </label>
      <br />
      <% Else%>
      <input class="checkbox" id="limitexceed" name="limitexceed" type="checkbox" onchange="rewardsavecheck();"
        <%if(allownegative=true)then send(" checked=""checked""")%> />
      <label for="limitexceed">
        <% Sendb(Copient.PhraseLib.Lookup("reward.exceedtrans", LanguageID))%>
      </label>
      <br />
      <% End If%>
      <hr class="hidden" />
    </div>
    <div class="box" id="sponsor">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.sponsor", LanguageID))%>
        </span>
        <% If (IsTemplate) Then%>
        <span class="tempRequire">
          <input type="checkbox" class="tempcheck" id="DisallowEditSpon1" name="DisallowEditSpon"
            <% if(bDisallowEditSpon)then send(" checked=""checked""") %> />
          <label for="temp-Tiers">
            <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
          </label>
        </span>
        <% ElseIf (bUseTemplateLocks And bDisallowEditSpon) Then%>
        <span class="tempRequire">
          <input type="checkbox" class="tempcheck" id="DisallowEditSpon2" name="DisallowEditSpon"
            disabled="disabled" checked="checked" />
          <label for="temp-Tiers">
            <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
          </label>
        </span>
        <% End If%>
      </h2>
      <%
        MyCommon.QueryStr = "select SponsorID, Description, PhraseID from Sponsors with (NoLock)"
        rst = MyCommon.LRT_Select()
        For Each row In rst.Rows
          Sendb("<input class=""radio"" id=""" & StrConv(row.Item("Description"), VbStrConv.Lowercase) & """ name=""sponsor"" type=""radio"" value=""" & row.Item("SponsorID") & """")
          If SponsorID = row.Item("SponsorID") Then
            Sendb(" checked=""checked""")
          End If
          If (bUseTemplateLocks And bDisallowEditSpon) Then
            Sendb(" disabled=""disabled""")
          End If
          Send(" /><label for=""" & StrConv(row.Item("Description"), VbStrConv.Lowercase) & """>" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID) & "</label>")
        Next
      %>
      <hr class="hidden" />
    </div>
    <div class="box" id="message">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.message", LanguageID))%>
        </span>
        <% If (IsTemplate) Then%>
        <span class="tempRequire">
          <input type="checkbox" class="tempcheck" id="DisallowEditMsg1" name="DisallowEditMsg"
            <% if(bDisallowEditMsg)then send(" checked=""checked""") %> />
          <label for="temp-Tiers">
            <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
          </label>
        </span>
        <% ElseIf (bUseTemplateLocks And bDisallowEditMsg) Then%>
        <span class="tempRequire">
          <input type="checkbox" class="tempcheck" id="DisallowEditMsg2" name="DisallowEditMsg"
            disabled="disabled" checked="checked" />
          <label for="temp-Tiers">
            <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
          </label>
        </span>
        <% End If%>
      </h2>
      <input class="longer" id="PrintLineText" name="PrintLineText" maxlength="100" type="text"
        value="<%sendb(PrintLineText) %>" <% If(bUseTemplateLocks and bDisallowEditMsg) Then sendb(" disabled=""disabled""") %> /><br />
      <%
        If Not (bUseTemplateLocks And bDisallowEditMsg) Then
          ' TOOLBAR
          Send("<div id=""ed_toolbar"" style=""background-color:#d0d0d0;"">")
          Send("    <span id=""tools"">")
          Sendb("      ")
          MyCommon.QueryStr = "select Distinct MT.MarkupID,MT.Tag,MT.Description,MT.PhraseID,MT.NumParams," & _
                              "MT.Param1Name,MT.Param1PhraseID,MT.Param2Name,MT.Param2PhraseID,MT.Param3Name,MT.Param3PhraseID," & _
                              "MT.DisplayOrder,MTU.RewardTypeID,MTU.EngineID,PhT.PhraseID,PhT.LanguageID," & _
                              "convert(nvarchar(50),Phrase) as Phrase,PTy.Installed from MarkupTags as MT with (NoLock) " & _
                              "inner join PhraseText as PhT with (NoLock) on MT.PhraseID=PhT.PhraseID " & _
                              "inner join PrinterTranslation as PTr with (NoLock) on MT.MarkupID=PTr.MarkupID " & _
                              "inner join PrinterTypes as PTy with (NoLock) on PTr.PrinterTypeID=PTy.PrinterTypeID " & _
                              "inner join MarkupTagUsage as MTU with (NoLock) on MT.MarkupID=MTU.MarkupID " & _
                              "where MTU.RewardTypeID=1 and MTU.EngineID=" & OfferEngineID & " and PTy.Installed=1 and LanguageID=" & LanguageID & " " & _
                              "order by MT.DisplayOrder;"
          rst = MyCommon.LRT_Select
          Dim cleanid As String
          For Each row In rst.Rows
            cleanid = row.Item("Tag")
            cleanid = cleanid.Replace("#", "Amt")
            cleanid = cleanid.Replace("$", "Dol")
            cleanid = cleanid.Replace("/", "Off")
            If (cleanid = "NETDol") Or (cleanid = "INITIALDol") Or (cleanid = "EARNEDDol") Or (cleanid = "REDEEMEDDol") Then
              Sendb("<input type=""button"" id=""ed_" & (StrConv(cleanid, VbStrConv.Lowercase)) & """ class=""ed_button"" " & DisabledAttribute & " title=""" & row.Item("Phrase") & """ onclick=""javascript:showDialogSpan(true, 1, this.value);"" value=""" & (StrConv(row.Item("Tag"), VbStrConv.ProperCase)) & """ />")
            ElseIf (cleanid = "NETAmt") Or (cleanid = "INITIALAmt") Or (cleanid = "EARNEDAmt") Or (cleanid = "REDEEMEDAmt") Then
              Sendb("<input type=""button"" id=""ed_" & (StrConv(cleanid, VbStrConv.Lowercase)) & """ class=""ed_button"" " & DisabledAttribute & " title=""" & row.Item("Phrase") & """ onclick=""javascript:showDialogSpan(true, 2, this.value);"" value=""" & (StrConv(row.Item("Tag"), VbStrConv.ProperCase)) & """ />")
            Else
              Sendb("<input type=""button"" id=""ed_" & (StrConv(cleanid, VbStrConv.Lowercase)) & """ class=""ed_button"" " & DisabledAttribute & " title=""" & row.Item("Phrase") & """ onclick=""edInsert" & (StrConv(cleanid, VbStrConv.ProperCase)) & "(PrintLineText);"" value=""" & (StrConv(row.Item("Tag"), VbStrConv.ProperCase)) & """ />")
            End If
          Next
          Send("      <br />")
          Send("    </span>")
          Send("    </div>")
        End If
      %>
      <hr class="hidden" />
      <% If MyCommon.Fetch_CM_SystemOption(23) Then%>
      <div id="divWebText">
        <br />
        <label for="WebText" style="position: relative;">
          <% Sendb(Copient.PhraseLib.Lookup("term.webmessage", LanguageID))%>:</label>
        <textarea class="longer" cols="48" rows="3" id="WebText" name="WebText"><% Sendb(WebText)%><% If (bUseTemplateLocks And bDisallowEditMsg) Then Sendb(" disabled=""disabled""")%></textarea><br />
        <br />
      </div>
      <% End If%>
    </div>
    <div class="box" id="options">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.advancedoptions", LanguageID))%>
        </span>
        <% If (IsTemplate Or (bUseTemplateLocks And bDisallowEditAdv)) Then%>
        <span class="tempRequire">
          <input type="checkbox" class="tempcheck" id="DisallowEditAdv" name="DisallowEditAdv"
            <% if(bDisallowEditAdv)then send(" checked=""checked""") %><% If(bUseTemplateLocks) Then sendb(" disabled=""disabled""") %> />
          <label for="temp-Tiers">
            <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
          </label>
        </span>
        <br class="printonly" />
        <% End If%>
      </h2>
      <% If (MyCommon.Fetch_CM_SystemOption(13) = "1") Then%>
      <%
        If (RewardAmountTypeID <> 4) Then
          sDisabled = " disabled=""disabled"""
        Else
          If (bUseTemplateLocks And bDisallowEditAdv) Then
            sDisabled = " disabled=""disabled"""
          Else
            sDisabled = ""
          End If
        End If
      %>
      <input class="checkbox" id="staticfuel" name="staticfuel" type="checkbox" <% sendb(sDisabled) %><% if(staticfuel=true)then sendb(" checked=""checked""") %> /><label
        for="staticfuel"><% Sendb(Copient.PhraseLib.Lookup("term.staticfuel", LanguageID))%></label><br />
      <% End If%>
      <% If (MyCommon.Fetch_CM_SystemOption(8) = "1") Then%>
      <%
        If (ValueRadio = 3) Then
          sDisabled = " disabled=""disabled"""
        Else
          If (bUseTemplateLocks And bDisallowEditAdv) Then
            sDisabled = " disabled=""disabled"""
          Else
            sDisabled = ""
          End If
        End If
      %>
      <input class="checkbox" id="bestdeal" name="bestdeal" type="checkbox" onclick="handleBestDealClickSent();" <% sendb(sDisabled) %><% if(bestdeal=true)then sendb(" checked=""checked""") %> /><label
        for="bestdeal"><% Sendb(Copient.PhraseLib.Lookup("term.bestdealitem", LanguageID))%></label><br />
      <% End If%>
      <%
        If (TransactionLevelPossible) Then
          sDisabled = " disabled=""disabled"""
        Else
          If (bUseTemplateLocks And bDisallowEditAdv) Then
            sDisabled = " disabled=""disabled"""
          Else
            sDisabled = ""
          End If
        End If
      %>
      <input class="checkbox" id="promote" name="promote" type="checkbox" <% sendb(sDisabled) %><% if(promotetotranslevel=true)then sendb(" checked=""checked""") %> /><label
        for="promote"><% Sendb(Copient.PhraseLib.Lookup("reward.promote", LanguageID))%></label><br />
      <% If MyCommon.Fetch_CM_SystemOption(3) Then%>
      <input class="checkbox" id="exclude" name="exclude" type="checkbox" <% sendb(sDisabled) %><% if(donotitemdistribute=true)then sendb(" checked=""checked""") %> /><label
        for="exclude"><% Sendb(Copient.PhraseLib.Lookup("reward.exclude", LanguageID))%></label><br />
      <% End If%>
      <input class="checkbox" id="disconly" name="disc-items-only" type="checkbox" <% sendb(sDisabled) %><% if(discountableitemsonly=true OrElse SDisabled<>"")then sendb(" checked=""checked""") %> /><label
        for="disconly"><% Sendb(Copient.PhraseLib.Lookup("reward.discountableonly", LanguageID))%></label><br />
		 <% If MyCommon.Fetch_CM_SystemOption(130) Then%>
            <input class="checkbox" id="SameItem" name="SameItem" type="checkbox" <% sendb(sDisabled) %><% if(sameitem=true)then sendb(" checked=""checked""") %> /><label
                for="SameItem"><% Sendb(Copient.PhraseLib.Lookup("term.sameitemreward", LanguageID))%></label><br />
         <% End If%>
      <% If MyCommon.Fetch_CM_SystemOption(2) Then%>
      <%
        If (bUseTemplateLocks And bDisallowEditAdv) Then
          sDisabled = " disabled=""disabled"""
        Else
          sDisabled = ""
        End If
      %>
      <input class="checkbox" id="decrement" name="decrement" type="checkbox" <% sendb(sDisabled) %><% if(effectminorder=true)then sendb(" checked=""checked""") %> /><label
        for="decrement"><% Sendb(Copient.PhraseLib.Lookup("reward.decrement", LanguageID))%></label><br />
      <% End If%>
      <% If MyCommon.Fetch_CM_SystemOption(136) Then%>
        <input type="hidden" name="virtualON" id="virtualON" value=1>
		  
		
        <%
          If (TransactionLevelPossible) Then
            sDisabled = " disabled=""disabled"""
          Else
            If (bUseTemplateLocks And bDisallowEditAdv) Then
              sDisabled = " disabled=""disabled"""
            Else
              sDisabled = ""
            End If
          End If
        %>
		
        <input class="checkbox" id="virtuallink" name="virtuallink" type="checkbox"<% sendb(sDisabled) %><% if(virtualLink=true)then sendb(" checked=""checked""") %> /><label for ="virtuallink"> Virtual Link Only </label><br />  

        <%Else%>
          <input type="hidden" name="virtualON" id="virtualON" value=0>
        <%End If%>
    </div>
  </div>
</div>
</form>
<script type="text/javascript">
<% If ((CloseAfterSave) and (infoMessage = "")) Then %>
    window.close();
<% End If %>

updateButtons();
removeUsed(false);
PupdateButtons();
PremoveUsed();
checkConditionState();
setlimitsection(false);
setperiodsection(false);

if (document.getElementById("functionselect") != null) {
  fullSelect = document.getElementById("functionselect").cloneNode(true);
}
  
document.getElementById("select1").onclick=select1_onclick;
document.getElementById("deselect1").onclick=deselect1_onclick;
document.getElementById("select2").onclick=select2_onclick;
document.getElementById("deselect2").onclick=deselect2_onclick;
function select1_onclick() {
  handleSelectClick('select1')
}
function deselect1_onclick() {
  handleSelectClick('deselect1')
}
function select2_onclick() {
  handleSelectClick('select2')
}
function deselect2_onclick() {
  handleSelectClick('deselect2')
}

function rewardsavecheck(){
  var elemSaveReward = document.getElementById("limitexceed");
 
  if(elemSaveReward.checked == true){
    elemSaveReward.checked = true;
  }else {
    elemSaveReward.checked = false;
  }
}
function priorityReset(){
	
			
		document.getElementById("PriorityReset").value="True";
		
		}	

</script>
<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd("mainform", "functioninput")
  Logix = Nothing
  MyCommon = Nothing
%>
