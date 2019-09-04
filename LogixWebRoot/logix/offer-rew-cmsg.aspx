﻿<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Text.RegularExpressions" %>
<%
  ' *****************************************************************************
  ' * FILENAME: offer-rew-cmsg.aspx 
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
  Dim row As DataRow
    Dim rst1 As DataTable
    Dim rst2 As DataTable
    Dim rst3 As DataTable
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
  Dim TransactionLevelSelected As Boolean = False
  Dim infoMessage As String = ""
  Dim DistPeriod As Integer
  Dim DisplayNow As Boolean
  Dim AllowNegative As Boolean
  Dim ChargeBackDeptID As Integer
  Dim ExItemLevelDist As Boolean
  Dim ProgramID As Long
  Dim DiscountableItemsOnly As Boolean
  Dim UseSpecialPricing As Boolean
  Dim SPRepeatAtOccur As Integer
  Dim ValueRadio As Integer
  Dim PrintLineText As String
  Dim q As Integer
  Dim x As Integer
  Dim Tiered As Integer
  Dim SponsorID As Integer
  Dim PromoteToTransLevel As Boolean
  Dim EffectMinOrder As Boolean
  Dim RewardLimit As Integer
  Dim RewardLimitTypeID As Integer
  Dim Disallow_Edit As Boolean = True
  Dim IsTemplate As Boolean = False
  Dim CloseAfterSave As Boolean = False
  Dim RequirePG As Boolean = False
  Dim DisabledAttribute As String = ""
  Dim Handheld As Boolean = False
  Dim ProductGroupID As Integer = 0
  Dim ExcludedID As Integer = 0
  Dim BannersEnabled As Boolean = False
  
  Dim bUseTemplateLocks As Boolean
  Dim bDisallowEditPg As Boolean = False
  Dim bDisallowEditSpon As Boolean = False
  Dim bDisallowEditMsg As Boolean = False
  Dim bDisallowEditDist As Boolean = False
  Dim bDisallowEditLimit As Boolean = False
  Dim sDisabled As String
  Dim AdvancedLimitID As Long
  Dim bStoreUser As Boolean = False
  Dim sValidLocIDs As String = ""
  Dim sValidSU As String = ""
  Dim wherestr As String = "" 
  Dim sJoin As String = ""
  Dim iLen As Integer = 0
  Dim i As Integer
  Dim RECORD_LIMIT As Integer = GroupRecordLimit '500
  Dim topString As String = ""
  Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249)="1",True,False)
  Dim ByExistingPGSelector As Boolean = IIf(MyCommon.Fetch_SystemOption(222) = "0", True, False)
  Dim PagePostBack As Boolean = True
  Dim ByAddSingleProduct As Boolean = True
  Dim ShowAllItems As Boolean
  Dim GroupSize As Integer
  Dim rstItems As DataTable = Nothing
  Dim descriptionItem As String = String.Empty
  Dim ProductTypeID As Integer = 0
  Dim IDLength As Integer = 0
  Dim GName As String = ""
  Dim OfferStartDate As Date
  Dim ExtProductID As String = ""
  Dim prodDT As DataTable
  Dim Description As String = ""
  Dim outputStatus As Integer
  Dim tempProducts As String = ""
  Dim tempProductsList() As String = Nothing
  Dim maxLimit As Integer = 0
  Dim validItemList As List(Of String) = New List(Of String)
  Dim invalidItemList As List(Of String) = New List(Of String)
  Dim tempTableInsertStatement As StringBuilder = New StringBuilder()
  Dim upc As String = ""
  Dim ProductsWithoutDesc As Integer
  Dim ListBoxSize As Integer
  If RECORD_LIMIT > 0 Then topString = "top " & RECORD_LIMIT
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "offer-rew-cmsg.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  'Store User
  If(MyCommon.Fetch_CM_SystemOption(131) = "1") Then
    MyCommon.QueryStr = "select LocationID from StoreUsers where UserID=" & AdminUserID & ";"
    rst = MyCommon.LRT_Select
    iLen = rst.Rows.Count
    If iLen > 0 Then
      bStoreUser = True
      sValidSU = AdminUserID
      For i=0 to (iLen-1)
        If i=0 Then 
          sValidLocIDs = rst.Rows(0).Item("LocationID")
        Else 
          sValidLocIDs &= "," & rst.Rows(i).Item("LocationID")
        End If
      Next
    
      MyCommon.QueryStr = "select UserID from StoreUsers where LocationID in (" & sValidLocIDs & ") and NOT UserID=" & AdminUserID & ";"
      rst = MyCommon.LRT_Select
      iLen = rst.Rows.Count
      If iLen > 0 Then
        For i=0 to (iLen-1)
          sValidSU &= "," & rst.Rows(i).Item("UserID") 
        Next
      End If
    End If
  End If
  
  OfferID = Request.QueryString("OfferID")
  RewardID = Request.QueryString("RewardID")
  NumTiers = Request.QueryString("NumTiers")
  ProductGroupID = MyCommon.Extract_Val(Request.QueryString("ProductGroupID"))
  ExcludedID = MyCommon.Extract_Val(Request.QueryString("ExcludedID"))
  
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  
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
      bDisallowEditSpon = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit3"), False)
      bDisallowEditMsg = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit4"), False)
      bDisallowEditDist = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit6"), False)
      bDisallowEditLimit = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit7"), False)
      If bUseTemplateLocks Then
        If Disallow_Edit Then
          bDisallowEditPg = True
          bDisallowEditSpon = True
          bDisallowEditMsg = True
          bDisallowEditDist = True
          bDisallowEditLimit = True
        Else
          Disallow_Edit = bDisallowEditPg And bDisallowEditSpon And bDisallowEditMsg And _
                          bDisallowEditDist And bDisallowEditLimit
        End If
      End If
    End If
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
	
  MyCommon.QueryStr = "select PG.Name from OfferRewards ORWD with (NoLock) Inner Join ProductGroups PG with (nolock) on ORWD.productgroupid=PG.productgroupid where ORWD.RewardID=" & RewardID & " and ORWD.deleted=0;"
  rst = MyCommon.LRT_Select
  
  If rst.Rows.Count > 0 Then
    GName = MyCommon.NZ(rst.Rows(0).Item("Name"), "")
  Else
    MyCommon.QueryStr = "select rewardid from offerrewards with (NoLock) where offerid=" & OfferID & " and rewardtypeid=4"
    rst3 = MyCommon.LRT_Select
    If rst3.Rows.Count = 1 Then
      GName = OfferStartDate.ToString("yyyyMMdd") & Name & "PGRCMSG"
    ElseIf rst3.Rows.Count > 1 Then
      GName = OfferStartDate.ToString("yyyyMMdd") & Name & "PGRCMSG(" & rst3.Rows.Count - 1 & ")"
    End If
  End If
  
  Send_HeadBegin("term.offer", "term.cmsgreward", OfferID)
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

function isValidProductList() {
  
  
  var retVal = true;
  var elemID = document.getElementById("pasteproducts");
  
  if((elemID != null) && (elemID.value.length == 0)) {
            retVal = false;
            alert('No products were provided');
        }
		
		return retVal;
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

// This is the javascript array holding the function list
// The PrintJavascriptArray ASP function can be used to print this array.
<%
    If bStoreUser Then
      sJoin = "Full Outer Join ProductGroupLocUpdate pglu with (NoLock) on pg.ProductGroupID=pglu.ProductGroupID " 
      wherestr = " and (LocationID in (" & sValidLocIDs & ") or (CreatedByAdminID in (" & sValidSU & ") and Isnull(LocationID,0)=0)) and AnyProduct=0" 
    End If
    
    MyCommon.QueryStr = "select " & topString & " pg.ProductGroupID,Name from ProductGroups pg with (NoLock) " & sJoin & " where Deleted=0 " & wherestr
    If(bEnableRestrictedAccessToUEOfferBuilder) Then MyCommon.QueryStr &= " and isnull(pg.TranslatedFromOfferID,0) = 0 " 
    
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

function checkConditionState() {
  var elemExcluded=document.getElementById("excluded");
  var elemSelected=document.getElementById("selected");
  var elemDistribution = document.getElementById("distribution");
  var currSelectedVal = -1;
  var isTransactionLevel = 0;
  var isAnyProduct = false;
  var hasNoExcluded = false;
  var hasNoSelected = false;
  
  hasNoSelected = (elemSelected != null && elemSelected.options.length==0);
  hasNoExcluded = (elemExcluded != null  &&  elemExcluded.options.length == 0);
  isAnyProduct = ((elemSelected != null) && (elemSelected.options.length==1 && elemSelected.options[0].value=='1'));

  if (hasNoSelected) {
    isTransactionLevel = true
  } else {
    isTransactionLevel = false
  }
  
  if (isTransactionLevel == true) {
    if (elemDistribution != null) {
      elemDistribution.style.display = "none";
      enableDistribution(false)
    }
  } else {
    if (elemDistribution != null) {
      elemDistribution.style.display = "block";
      enableDistribution(true)
    }
  }
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
  
  if (elemDisallowEditDistOpt != null && elemDisallowEditDistOpt.value == '1') {
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
    if (elemTriggerprorate != null) { elemTriggerprorate.disabled = (isEnabled) ? false : true; }
    if (elemProrate != null) { elemProrate.disabled = (isEnabled) ? false : true; }
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
    
   if (selectList != '') { 
  document.getElementById("ProductGroupID").value = selectList;     
  }

  else {    
    
    document.getElementById("ProductGroupID").value = document.getElementById("NewCreatedProdGroupID").value
  }
    document.getElementById("ExcludedID").value = excludededList;
    // alert(htmlContents);
    return true;
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

// This is the javascript array holding the function list
// The PrintJavascriptArray ASP function can be used to print this array.
<%
    MyCommon.QueryStr = "select 0 as LimitID, 'None' as Name, RewardLimitTypeID as LimitTypeID, RewardLimit as LimitValue, RewardDistPeriod as LimitPeriod " & _
                        "from OfferRewards with (NoLock) where RewardID=" & RewardID & " " & _
                        "union " & _
                        "select LimitID, Name, LimitTypeID, LimitValue, LimitPeriod " & _
                        "from CM_AdvancedLimits with (NoLock) where Deleted=0 and LimitTypeID=6 order By LimitID;"
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
  var elemPeriod=document.getElementById("limitPeriod");
  var elemValue=document.getElementById("limitValue");
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
  MyCommon.QueryStr = "select LinkID,Tiered,SponsorID,PromoteToTransLevel,RewardDistPeriod,RewardLimit,RewardLimitTypeID,TriggerQty,RewardAmountTypeID," & _
                      "UseSpecialPricing,SPRepeatAtOccur,ApplyToLimit,DoNotItemDistribute,AdvancedLimitID from OfferRewards with (NoLock) where RewardID=" & RewardID
  rst = MyCommon.LRT_Select
  For Each row In rst.Rows
    DistPeriod = MyCommon.NZ(row.Item("RewardDistPeriod"), 0)
    LinkID = row.Item("LinkID")
    Tiered = MyCommon.NZ(row.Item("Tiered"), 0)
    SponsorID = MyCommon.NZ(row.Item("SponsorID"), 0)
    PromoteToTransLevel = MyCommon.NZ(row.Item("PromoteToTransLevel"), 0)
    RewardLimit = MyCommon.NZ(row.Item("RewardLimit"), 0)
    RewardLimitTypeID = MyCommon.NZ(row.Item("RewardLimitTypeID"), 2)
    RewardAmountTypeID = MyCommon.NZ(row.Item("RewardAmountTypeID"), 1)
    AdvancedLimitID = MyCommon.NZ(row.Item("AdvancedLimitID"), 0)
    'ExItemLevelDist = MyCommon.NZ(row.Item("ExItemLevelDist"), 0)
    TriggerQty = MyCommon.NZ(row.Item("TriggerQty"), 1)
    ApplyToLimit = MyCommon.NZ(row.Item("ApplyToLimit"), 1)
    UseSpecialPricing = MyCommon.NZ(row.Item("UseSpecialPricing"), 0)
    SPRepeatAtOccur = MyCommon.NZ(row.Item("SPRepeatAtOccur"), 1)
    DoNotItemDistribute = row.Item("DoNotItemDistribute")
  Next
  
  If (TriggerQty = ApplyToLimit And TriggerQty <> 0) Then
    ValueRadio = 1
  Else
    ValueRadio = 2
  End If
  
  ' find out if were delivering this now or end of sale
  MyCommon.QueryStr = "select DisplayNow from CashierMessages with (NoLock) where MessageID=" & LinkID
  rst = MyCommon.LRT_Select
  For Each row In rst.Rows
    DisplayNow = MyCommon.NZ(row.Item("DisplayNow"), 0)
  Next
  
  If (Request.QueryString("save") <> "") Then

    If (Request.QueryString("IsTemplate") = "IsTemplate") Then
      ' time to update the status bits for the templates
      Dim form_Disallow_Edit As Integer = 0
      Dim iDisallowEditPg As Integer = 0
      Dim iDisallowEditSpon As Integer = 0
      Dim iDisallowEditMsg As Integer = 0
      Dim iDisallowEditDist As Integer = 0
      Dim iDisallowEditLimit As Integer = 0
      
      Disallow_Edit = False
      bDisallowEditPg = False
      bDisallowEditSpon = False
      bDisallowEditMsg = False
      bDisallowEditDist = False
      bDisallowEditLimit = False
      
      If (Request.QueryString("Disallow_Edit") = "on") Then
        form_Disallow_Edit = 1
        Disallow_Edit = True
      End If
      If (Request.QueryString("DisallowEditPg") = "on") Then
        iDisallowEditPg = 1
        bDisallowEditPg = True
      End If
      If (Request.QueryString("DisallowEditSpon") = "on") Then
        iDisallowEditSpon = 1
        bDisallowEditSpon = True
      End If
      If (Request.QueryString("DisallowEditMsg") = "on") Then
        iDisallowEditMsg = 1
        bDisallowEditMsg = True
      End If
      If (Request.QueryString("DisallowEditDist") = "on") Then
        iDisallowEditDist = 1
        bDisallowEditDist = True
      End If
      If (Request.QueryString("DisallowEditLimit") = "on") Then
        iDisallowEditLimit = 1
        bDisallowEditLimit = True
      End If
      MyCommon.QueryStr = "update OfferRewards with (RowLock) set Disallow_Edit=" & form_Disallow_Edit & _
        ",DisallowEdit1=" & iDisallowEditPg & _
        ",DisallowEdit3=" & iDisallowEditSpon & _
        ",DisallowEdit4=" & iDisallowEditMsg & _
        ",DisallowEdit6=" & iDisallowEditDist & _
        ",DisallowEdit7=" & iDisallowEditLimit & _
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

    If Not (bUseTemplateLocks And bDisallowEditMsg) Then
      If (Request.QueryString("displaynow") = "on") Then
        MyCommon.QueryStr = "update CashierMessages with (RowLock) set DisplayNow=1 where MessageID=" & LinkID
        MyCommon.LRT_Execute()
        DisplayNow = True
      Else
        MyCommon.QueryStr = "update CashierMessages with (RowLock) set DisplayNow=0 where MessageID=" & LinkID
        MyCommon.LRT_Execute()
        DisplayNow = False
      End If
      ' ok here we need to handle the tiering stuffs
      If (Tiered = 0) Then
        'dbo.pt_ConditionTiers_Update @ConditionID bigint, @TierLevel int,@AmtRequired decimal(12,3)
        MyCommon.QueryStr = "dbo.pt_CashierTiers_Update"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@MessageID", SqlDbType.BigInt).Value = LinkID
        MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int).Value = 0
        MyCommon.LRTsp.Parameters.Add("@Line1Text", SqlDbType.NVarChar, 100).Value = Left(Request.QueryString("tier0-1"), 100)
        'MyCommon.LRTsp.Parameters.Add("@Line2Text", SqlDbType.NVarChar, 25).Value = MyCommon.Parse_Quotes(Request.QueryString("tier0-2"))
        'MyCommon.LRTsp.Parameters.Add("@Line3Text", SqlDbType.NVarChar, 25).Value = MyCommon.Parse_Quotes(Request.QueryString("tier0-3"))
        MyCommon.LRTsp.ExecuteNonQuery()
        MyCommon.Close_LRTsp()
      Else
        ' delete the current tier ammounts
        MyCommon.QueryStr = "delete from CashierMessageTiers with (RowLock) where MessageID=" & LinkID
        MyCommon.LRT_Execute()
        For x = 1 To NumTiers
          'Response.Write("->@ConditionID: " & ConditionID & " @TierLevel: " & x & " @AmtRequired: " & Request.QueryString("tier" & x) & "<br />")
          MyCommon.QueryStr = "dbo.pt_CashierTiers_Update"
          MyCommon.Open_LRTsp()
          MyCommon.LRTsp.Parameters.Add("@MessageID", SqlDbType.BigInt).Value = LinkID
          MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int).Value = x
          MyCommon.LRTsp.Parameters.Add("@Line1Text", SqlDbType.NVarChar, 100).Value = Left(Request.QueryString("tier" & x & "-1"), 100)
          'MyCommon.LRTsp.Parameters.Add("@Line2Text", SqlDbType.NVarChar, 25).Value = MyCommon.Parse_Quotes(Request.QueryString("tier" & x & "-2"))
          'MyCommon.LRTsp.Parameters.Add("@Line3Text", SqlDbType.NVarChar, 25).Value = MyCommon.Parse_Quotes(Request.QueryString("tier" & x & "-3"))
          MyCommon.LRTsp.ExecuteNonQuery()
          MyCommon.Close_LRTsp()
        Next
      End If
    End If

    If Not (bUseTemplateLocks And bDisallowEditDist) Then
      If (Request.QueryString("trigger") <> "") Then
        If (Request.QueryString("trigger") = "1") Then
          ' set  TriggerQty=Xbox
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
          'If (TriggerQty = ApplyToLimit) Then ValueRadio = 1
        End If
      Else
        MyCommon.QueryStr = "update OfferRewards with (RowLock) set TriggerQty=0," & _
        "ApplyToLimit=1 where RewardID=" & RewardID
        MyCommon.LRT_Execute()
      End If
    End If
    
    If Not (bUseTemplateLocks And bDisallowEditLimit) Then
      If (Request.QueryString("selectadv") <> "") Then
        AdvancedLimitID = MyCommon.NZ(Request.QueryString("selectadv"), 0)
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
        If (Request.QueryString("RewardLimitTypeID") <> "") Then
          RewardLimitTypeID = MyCommon.Extract_Val(Request.QueryString("RewardLimitTypeID"))
          MyCommon.QueryStr = "update OfferRewards with (RowLock) set RewardLimitTypeID=" & RewardLimitTypeID & " where RewardID=" & RewardID
          MyCommon.LRT_Execute()
        End If
        If (Request.QueryString("limitvalue") <> "") Then
          RewardLimit = MyCommon.Extract_Val(Request.QueryString("limitvalue"))
          MyCommon.QueryStr = "update OfferRewards with (RowLock) set RewardLimit=" & RewardLimit & " where RewardID=" & RewardID
          MyCommon.LRT_Execute()
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
    
    If Not (bUseTemplateLocks And bDisallowEditSpon) Then
      If (Request.QueryString("sponsor") <> "") Then
        SponsorID = MyCommon.Extract_Val(Request.QueryString("sponsor"))
        MyCommon.QueryStr = "update OfferRewards with (RowLock) set SponsorID=" & SponsorID & " where RewardID=" & RewardID
        MyCommon.LRT_Execute()
      End If
    End If

    MyCommon.QueryStr = "update OfferRewards with (RowLock) set TCRMAStatusFlag=2,CMOAStatusFlag=2 where RewardID=" & RewardID
    MyCommon.LRT_Execute()
    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.rew-cmsg", LanguageID))

    MyCommon.QueryStr = "update OfferRewards with (RowLock) set TCRMAStatusFlag=3,CMOAStatusFlag=2 where RewardID=" & RewardID
    MyCommon.LRT_Execute()
    MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where offerid=" & OfferID
    MyCommon.LRT_Execute()

        CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
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
  Send("    opener.location = 'offer-rew.aspx?OfferID=" & OfferID & "'; ")
    Send("    } ")
    Send("  }")
    Send("}")
  Send("</script>")
%>
<form action="offer-rew-cmsg.aspx" id="mainform" name="mainform" onsubmit="return saveForm();">
  <div id="intro">
    <input type="hidden" id="OfferID" name="OfferID" value="<% sendb(OfferID) %>" />
    <input type="hidden" id="Name" name="Name" value="<% sendb(Name) %>" />
    <input type="hidden" id="RewardID" name="RewardID" value="<% sendb(RewardID) %>" />
    <input type="hidden" id="ProductGroupID" name="ProductGroupID" value="<% sendb(ProductGroupID) %>" />
    <input type="hidden" id="ExcludedID" name="ExcludedID" value="<% sendb(ExcludedID) %>" />
    <input type="hidden" id="OriginalPeriod" name="OriginalPeriod" value="<% sendb(DistPeriod) %>" />
    <input type="hidden" id="ImpliedPeriod" name="ImpliedPeriod" value="<% sendb(DistPeriod) %>" />
    <input type="hidden" id="LimitsDisabled" name="LimitsDisabled" value="<% sendb(bUseTemplateLocks and bDisallowEditLimit) %>" />
    <input type="hidden" id="IsTemplate" name="IsTemplate" value="<% 
            if(istemplate)then 
            sendb("IsTemplate")
            else 
            sendb("Not") 
            end if
        %>" />
    <%If (IsTemplate) Then
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.cmsgreward", LanguageID), VbStrConv.Lowercase) & "</h1>")
      Else
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.cmsgreward", LanguageID), VbStrConv.Lowercase) & "</h1>")
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
      <% If Not (IsTemplate) Then
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
    <% If MyCommon.Fetch_SystemOption(222) = "0" Then %>
	  <input type="radio" id="pgselectortype1" name="pgselectortype" onclick="javascript:ProductGroupTypeSelection();"
      <% if(ByExistingPGSelector) then sendb(" checked=""checked""") %> value="existingadd" /><label
      for="pgselectortype1"><% Sendb(Copient.PhraseLib.Lookup("term.selectexistingprodgroups", LanguageID))%></label>
      <input type="radio" id="pgselectortype2" name="pgselectortype" onclick="javascript:ProductGroupTypeSelection();"
      <% if(Not ByExistingPGSelector) then sendb(" checked=""checked""") %> value="directadd" /><label
      for="pgselectortype2">Add products to reward</label>
	<%Else%>
	  <input type="radio" id="pgselectortype2" name="pgselectortype" onclick="javascript:ProductGroupTypeSelection();"
      <% if(Not ByExistingPGSelector) then sendb(" checked=""checked""") %> value="directadd" /><label
      for="pgselectortype2">Add products to reward</label>
	  <input type="radio" id="pgselectortype1" name="pgselectortype" onclick="javascript:ProductGroupTypeSelection();"
      <% if(ByExistingPGSelector) then sendb(" checked=""checked""") %> value="existingadd" /><label
      for="pgselectortype1"><% Sendb(Copient.PhraseLib.Lookup("term.selectexistingprodgroups", LanguageID))%></label>
	<%End If%>
    <div id="columnfull">
     <%Send_DirectProductAddSelector(Logix, ByExistingPGSelector, ShowAllItems , GroupSize , rstItems , ProductsWithoutDesc , descriptionItem , ByAddSingleProduct , IDLength, GName) %>
      <% Send_ProductConditionSelectorAdv(Logix, TransactionLevelSelected, bUseTemplateLocks, Disallow_Edit, SelectedItem, ExcludedItem, RewardID, Copient.CommonInc.InstalledEngines.CM, IsTemplate, bDisallowEditPg, ByExistingPGSelector, bStoreUser, sValidLocIDs, sValidSU)%>
      
    </div>
    
    <div id="column1">
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
        <%
          If (bUseTemplateLocks And bDisallowEditMsg) Then
            sDisabled = " disabled=""disabled"""
          Else
            sDisabled = ""
          End If
          MyCommon.QueryStr = "select CT.MessageID,Tiered,O.Numtiers,O.OfferID,CT.TierLevel,CT.Line1Text from Offerrewards as OC with (NoLock) " & _
                              "left join Offers as O with (NoLock) on O.OfferID=OC.OfferID left join CashierMessageTiers as CT with (NoLock) on OC.LinkID=CT.MessageID where OC.RewardID=" & RewardID
          rst = MyCommon.LRT_Select()
          q = 1
          For Each row In rst.Rows
            If q = 1 Then
              Send("<input type=""hidden"" name=""NumTiers"" value=""" & row.Item("NumTiers") & """ />")
            End If
            If (row.Item("Tiered") = False) Then
              Send("        <label for=""tier0-1""><b>" & Copient.PhraseLib.Lookup("term.line", LanguageID) & ":</b></label><br />")
              Send("        <input class=""longer"" id=""tier0-1"" name=""tier0-1"" type=""text"" maxlength=""99"" value=""" & row.Item("Line1Text") & """" & sDisabled & " /><br />")
            Else
              Send("        <label for=""tier" & q & """><b>" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & q & ":</b></label><br />")
              Send("        <input class=""longer"" id=""tier" & q & """ name=""tier" & q & "-1"" type=""text"" maxlength=""99"" value=""" & row.Item("Line1Text") & """" & sDisabled & " /><br />")
            End If
            q = q + 1
          Next
        %>
        <br class="half" />
        <input id="displaynow" name="displaynow" type="checkbox"<% Send(sDisabled) %><%if(displaynow)then send(" checked=""checked""") %> /><label
          for="displaynow"><% Sendb(Copient.PhraseLib.Lookup("reward.displayimmediately", LanguageID))%></label><br />
      </div>

      <div class="box" id="distribution" style="display: <% Sendb(IIF(TransactionLevelSelected, "block", "none"))%>;">
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
        <% 
          If (bUseTemplateLocks And bDisallowEditDist) Then
            sDisabled = " disabled=""disabled"""
          Else
            sDisabled = ""
          End If
        %>
        <input id="triggerbogo" name="trigger" type="radio" <% if(valueradio=1)then sendb(" checked=""checked""") %>
          value="1"<% Send(sDisabled) %> />
        <label for="triggerbogo">
          <% Sendb(Copient.PhraseLib.Lookup("reward.messageevery", LanguageID))%>
        </label>
        <br />
        &nbsp; &nbsp; &nbsp; &nbsp;
        <label for="Xbox">
          <% Sendb(Copient.PhraseLib.Lookup("term.mustpurchase", LanguageID))%>
        </label>
        <input class="shortest" id="Xbox" name="Xbox" maxlength="9" type="text"<% Send(sDisabled) %><% if(valueradio=1)then sendb(" value=""" & triggerqty & """ ") %> />
        <% Sendb(Copient.PhraseLib.Lookup("term.item(s)", LanguageID))%>
        <br />
        <input id="triggerbxgy" name="trigger" type="radio" value="2"<% Send(sDisabled) %><% if(valueradio=2)then sendb(" checked=""checked""") %> />
        <label for="triggerbxgy">
          <% Sendb(Copient.PhraseLib.Lookup("term.buy", LanguageID))%>
        </label>
        <input class="shortest" id="bxgy1" name="Xbox2" maxlength="9" type="text"<% Send(sDisabled) %><% if(valueradio=2)then sendb(" value=""" & triggerqty-applytolimit & """ ") %> />
        <% Sendb(Copient.PhraseLib.Lookup("term.item(s)", LanguageID))%>
        ,
        <% Sendb(Copient.PhraseLib.Lookup("reward.givemessageto", LanguageID))%>
        <input class="shortest" id="bxgy2" name="Ybox2" maxlength="9" type="text"<% Send(sDisabled) %><% if(valueradio=2)then sendb(" value=""" & applytolimit & """ ") %> /><br />
        &nbsp;<br />
        <hr class="hidden" />
      </div>
      </div>
      <div id="gutter">
    </div>
    <div id="column2">
      <div class="box" id="limits">
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
          MyCommon.QueryStr = "Select LimitId, Name from CM_AdvancedLimits with (NoLock) where Deleted=0 and LimitTypeID=6 order By Name;"
          rst = MyCommon.LRT_Select
          If rst.Rows.Count > 0 Then
        %>
        <label for="selectadv"><% Sendb(Copient.PhraseLib.Lookup("term.advlimits", LanguageID))%>:</label>
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
        &nbsp;<% Sendb(Copient.PhraseLib.Lookup("term.per", LanguageID))%>
        <input class="shorter" id="limitperiod" name="form_DistPeriod" maxlength="4" type="text" value="<% sendb(DistPeriod) %>"
          <% If(bUseTemplateLocks and bDisallowEditLimit) Then sendb(" disabled=""disabled""") %> />
        <select id="selectday" name="selectday" onchange="setperiodsection(true);"
          <% If(bUseTemplateLocks and bDisallowEditLimit) Then sendb(" disabled=""disabled""") %>>
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
      <div class="box" id="sponsor">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.sponsor", LanguageID))%>
          </span>
          <% If (IsTemplate) Then%>
          <span class="tempRequire">
            <input type="checkbox" class="tempcheck" id="DisallowEditSpon1" name="DisallowEditSpon"
              <% if (bDisallowEditSpon) then send(" checked=""checked""") %> />
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
    </div>
  </div>
</form>

<script type="text/javascript">
<% If (CloseAfterSave) Then %>
    window.close();
<% End If %>

updateButtons();
removeUsed(false);
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
  checkConditionState();
}
function deselect1_onclick() {
  handleSelectClick('deselect1')
  checkConditionState();
}
function select2_onclick() {
  handleSelectClick('select2')
  checkConditionState();
}
function deselect2_onclick() {
  handleSelectClick('deselect2')
  checkConditionState();
}

</script>

<%
done:
  MyCommon.Close_LogixRT()
  If Tiered = 0 Then
    Send_BodyEnd("mainform", "tier0-1")
  Else
    Send_BodyEnd("mainform", "tier1")
  End If
  Logix = Nothing
  MyCommon = Nothing
%>
