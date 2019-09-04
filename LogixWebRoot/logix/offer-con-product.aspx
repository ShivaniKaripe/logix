<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Text.RegularExpressions" %>
<%
    ' *****************************************************************************
    ' * FILENAME: offer-con-product.aspx 
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
    Dim rst1 As DataTable
    Dim row As DataRow
    Dim rst2 As DataTable
    Dim rst3 As DataTable
    Dim row2 As DataRow
    Dim OfferID As Long
    Dim Name As String = ""
    Dim ConditionID As String
    Dim ExcludedItem As Integer
    Dim SelectedItem As Integer
    Dim NumTiers As Integer
    Dim Tiered As Boolean
    Dim IsTransactionLevel As Boolean = False
    Dim bUseTemplateLocks As Boolean
    Dim IsTemplate As Boolean = False
    Dim CloseAfterSave As Boolean = False
    Dim DisabledAttribute As String = ""
    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False
    Dim ProductGroupID As Integer = 0
    Dim ExcludedID As Integer = 0
    Dim BannersEnabled As Boolean = False
    Dim ProdGroupChanged As Boolean = False
    Dim DiscountableItemsOnly As String = ""
    Dim UseNetForItems As String = ""
    Dim Disallow_Edit As Boolean = True
    Dim RequirePG As Boolean = False
    Dim bDisallowEditValue As Boolean = False
    Dim bDisallowEditRewards As Boolean = False
    Dim bDisallowEditAdvanced As Boolean = False
    Dim bDisallowEditPg As Boolean = False
    Dim sDisabled As String
    Dim sChecked As String
    Dim RECORD_LIMIT As Integer = GroupRecordLimit '500
    Dim IDLength As Integer = 0
    Dim ProductTypeID As Integer = 0
    Dim rstItems As DataTable = Nothing
    Dim GroupSize As Integer
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
    Dim ByExistingPGSelector As Boolean = IIf(MyCommon.Fetch_SystemOption(222) = "0", True, False)
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
    Dim bStoreUser As Boolean = False
    Dim sValidLocIDs As String = ""
    Dim sValidSU As String = ""
    Dim wherestr As String = "" 
    Dim sJoin As String = ""
    Dim iLen As Integer = 0
    Dim i As Integer = 0
    Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249)="1",True,False)
    Dim bCreateGroupOrProgramFromOffer As Boolean = IIf(MyCommon.Fetch_CM_SystemOption(134) ="1",True,False)
    Dim bStaticPG As Boolean
    
    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If
  
    Response.Expires = 0
    MyCommon.AppName = "offer-con-product.aspx"
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
  
    Tiered = False
    OfferID = Request.QueryString("OfferID")
    ConditionID = Request.QueryString("ConditionID")
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
  
    If (IsTemplate Or bUseTemplateLocks) Then
        ' lets dig the permissions if its a template
        MyCommon.QueryStr = "select Tiered,Disallow_Edit,RequiredFromTemplate,DisallowEdit1,DisallowEdit2,DisallowEdit3,DisallowEdit4 from OfferConditions with (NoLock) where ConditionID=" & ConditionID & ";"
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
            Disallow_Edit = MyCommon.NZ(rst.Rows(0).Item("Disallow_Edit"), True)
            RequirePG = MyCommon.NZ(rst.Rows(0).Item("RequiredFromTemplate"), False)
            bDisallowEditPg = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit1"), False)
            bDisallowEditValue = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit2"), False)
            bDisallowEditRewards = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit3"), False)
            bDisallowEditAdvanced = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit4"), False)
            Tiered = MyCommon.NZ(rst.Rows(0).Item("Tiered"), False)
            If Tiered Then
                bDisallowEditRewards = True
            End If
            If bUseTemplateLocks Then
                If Disallow_Edit Then
                    bDisallowEditPg = True
                    bDisallowEditValue = True
                    bDisallowEditRewards = True
                    bDisallowEditAdvanced = True
                Else
                    Disallow_Edit = bDisallowEditPg And bDisallowEditValue And bDisallowEditRewards And bDisallowEditAdvanced
                End If
            End If
        End If
    End If
  
  MyCommon.QueryStr = "select PG.Name, isnull(PG.IsStatic,0) as IsStatic from offerconditions OCD with (NoLock) Inner Join ProductGroups PG with (nolock) on OCD.LinkID=PG.productgroupid where OCD.ConditionID=" & ConditionID & " and OCD.deleted=0;"
    rst = MyCommon.LRT_Select
  
    If rst.Rows.Count > 0 Then
      GName = MyCommon.NZ(rst.Rows(0).Item("Name"), "")
      bStaticPG = MyCommon.NZ(rst.Rows(0).Item("IsStatic"), False)
    Else
      MyCommon.QueryStr = "select conditionid from offerconditions with (NoLock) where offerid=" & OfferID & " and conditiontypeid=2"
      rst3 = MyCommon.LRT_Select
      If rst3.Rows.Count = 1 Then
        GName = OfferStartDate.ToString("yyyyMMdd") & Name & "PGC"
      ElseIf rst3.Rows.Count > 1 Then
        GName = OfferStartDate.ToString("yyyyMMdd") & Name & "PGC(" & rst3.Rows.Count - 1 & ")"
      End If
        bStaticPG = False
    End If
  
    Send_HeadBegin("term.offer", "term.productcondition", OfferID)
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
%>
<script type="text/javascript" language="javascript">
var fullSelect = null;
/***************************************************************************************************************************/
//Script to call server method through JavaScript
//to load product based on search criteria.
var isFireFox = (navigator.appName.indexOf('Mozilla') !=-1) ? true : false;
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


        //alert(responseMsg);
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
        var ConditionID = <%= ConditionID%>;

        var bAllowHyphen  = '<% Sendb(MyCommon.Fetch_SystemOption(208))%>';
        if(bAllowHyphen == 1) {
            products = (products.toString().trim().replace(/\s/g, ', ')).replace(/-/g, '');
        } else {
            products = products.replace(/\r?\n/g, ', ');
        }
        return "Mode=" + mode + "&Products="+ products + "&OperationType="+ operationType + "&ProductType="+ productType + "&GName="+ GName + "&RewardID="+ ConditionID + "&IsCondition="+ true;

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
  return "Mode=" + mode + "&ProductSearch=" + document.getElementById('functioninput').value + "&OfferID=" + document.getElementById('OfferID').value + "&SelectedGroup=" + selectedGroup + "&ExcludedGroup=" + excludedGroup + "&SearchRadio=" + radioString + "&CallingPage=conProduct";
 
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
    //handleSaveButton(false);
  }
  else if(str.length == 0){
    if(!isFireFox){
      document.getElementById("pgList").innerHTML = '<select class="long" id="functionselect" name="functionselect" size="20"<% sendb(disabledattribute) %>>&nbsp;</select>';
    }
    else{
      document.getElementById("functionselect").innerHTML = '&nbsp;';
    }
    document.getElementById("searchLoadDiv").innerHTML = '&nbsp;';
    
   // handleSaveButton(false);
  }
}

/***************************************************************************************************************************/

var prevSelectedVal = -1
var prevIsTransactionLevel = -1 

function populateConditionType() {
  var elemConditionType=document.getElementById("valuetype");
  var elemSelected=document.getElementById("selected");
  var elemExcluded=document.getElementById("excluded");
  var elemAdded=document.getElementById("PKID");
    var elemradioaddprod = document.getElementById("pgselectortype2");

  var currSelectedVal = -1;
  var isTransactionLevel = 0;
  var isAnyProduct = false;
  var hasNoExcluded = false;
  var hasNoSelected = false;
  var hasNoAdded = false;
  
  hasNoSelected = (elemSelected != null && elemSelected.options.length==0);
  hasNoAdded = (elemAdded != null && elemAdded.options.length==0)
  hasNoExcluded = (elemExcluded != null  &&  elemExcluded.options.length == 0);
  isAnyProduct = ((elemSelected != null) && (elemSelected.options.length==1 && elemSelected.options[0].value=='1'));

  if(elemradioaddprod.checked){
    if (hasNoAdded)
        isTransactionLevel = 1
       else 
        isTransactionLevel = 0    
  }else{
      if (hasNoSelected)
        isTransactionLevel = 1
       else 
        isTransactionLevel = 0     
  }

  if (elemConditionType != null) {
    
    currSelectedVal = elemConditionType.options[elemConditionType.selectedIndex].value;
    if (prevSelectedVal == -1) {
      prevSelectedVal = currSelectedVal;
    }

    // remove all the options from the discount type select box
    while (elemConditionType.options.length > 0) {
      elemConditionType.options[0] = null;
    }

    if (isTransactionLevel != 0) {
      elemConditionType.options[0] = new Option("<% Sendb(Copient.PhraseLib.Lookup("term.minimumorder", LanguageID))%>", "7", false, false);
      elemConditionType.options[1] = new Option("<% Sendb(Copient.PhraseLib.Lookup("term.transactiontotal", LanguageID))%>", "6", false, false);
      elemConditionType.options[2] = new Option("<% Sendb(Copient.PhraseLib.Lookup("term.trxdisctotals", LanguageID))%>", "10", false, false);
    } else {
      elemConditionType.options[0] = new Option("<% Sendb(Copient.PhraseLib.Lookup("term.quantity", LanguageID) & " (" & Copient.PhraseLib.Lookup("term.count", LanguageID) & ")")%>", "1", false, false);
      elemConditionType.options[1] = new Option("<% Sendb(Copient.PhraseLib.Lookup("term.amount", LanguageID) & " (" & Copient.PhraseLib.Lookup("term.dollars", LanguageID) & ")")%>", "2", false, false);
      elemConditionType.options[2] = new Option("<% Sendb(Copient.PhraseLib.Lookup("term.weight", LanguageID) & " (" & Copient.PhraseLib.Lookup("term.pounds", LanguageID) & ")")%>", "3", false, false);
      elemConditionType.options[3] = new Option("<% Sendb(Copient.PhraseLib.Lookup("term.volume", LanguageID) & " (" & Copient.PhraseLib.Lookup("term.gallons", LanguageID) & ")")%>", "4", false, false);
    }

    if (isTransactionLevel != prevIsTransactionLevel) {
      // select the previously selected discount type
      for (var i=0; i<elemConditionType.options.length; i ++) {
        if (elemConditionType.options[i].value == prevSelectedVal) {
          elemConditionType.options[i].selected = true;
          break;
        }
      }
      prevSelectedVal = currSelectedVal;
      prevIsTransactionLevel = isTransactionLevel;
    } else {
      for (var i=0; i<elemConditionType.options.length; i ++) {
        if (elemConditionType.options[i].value == currSelectedVal) {
          elemConditionType.options[i].selected = true;
          break;
        }
      }
    }
  }
}

function showAdvancedOptions() {
  var elemSelected = document.getElementById("selected");
  var elemMinOrderItems = document.getElementById("minorderitems");
  var elemDiscItemsOnly = document.getElementById("discitemsonly");
  var elemUseNetForItems = document.getElementById("UseNetForItems");
  var elemConditionType = document.getElementById("valuetype");
  var elemDisallowEditAdvancedOpt = document.getElementById("DisallowEditAdvancedOpt");
  var isTransactionLevel = false;
  
  if (elemDisallowEditAdvancedOpt != null && elemDisallowEditAdvancedOpt.value == '1') {
    elemMinOrderItems.disabled = true;
    elemDiscItemsOnly.disabled = true;
    elemUseNetForItems.disabled = true;
  } else {
    isTransactionLevel = (elemSelected != null && elemSelected.options.length==0 );

    if (elemMinOrderItems != null) { 
      if (isTransactionLevel) {
        elemMinOrderItems.disabled = true;
      } else {
        elemMinOrderItems.disabled = false;
      }
    }

    if (elemDiscItemsOnly != null) { 
      if (isTransactionLevel) {
        elemDiscItemsOnly.disabled = true;
      } else {
        elemDiscItemsOnly.disabled = false;
      }
    } 

    if (elemUseNetForItems != null) { 
      if (!isTransactionLevel && ((elemConditionType.options[1].selected == true) || (elemConditionType.options[0].selected == true))) {
        elemUseNetForItems.disabled = false;
      } else {
        elemUseNetForItems.disabled = true;
      }
    } 
  }
} 

function removeUsed(bSkipKeyUp)
{
  //if (!bSkipKeyUp) handleKeyUp(99999);
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
  var bValidEntry = false;
  
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
    
  // update the hidden fields
  if (selectList != '') { 
  document.getElementById("ProductGroupID").value = selectList;     
  }

  else {        
  if(document.getElementById("NewCreatedProdGroupID") != null)
    document.getElementById("ProductGroupID").value = document.getElementById("NewCreatedProdGroupID").value
  }
  document.getElementById("ExcludedID").value = excludededList;   
    return true;
}

function updateButtons(){
  var elemDisallowEditPgOpt = document.getElementById("DisallowEditPgOpt");
  var elemDisallowEditValueOpt = document.getElementById("DisallowEditValueOpt");
  var selectObj = document.getElementById('selected');
  var excludedObj = document.getElementById('excluded');
  
  if (elemDisallowEditPgOpt != null && elemDisallowEditPgOpt.value == '1') {
      document.getElementById('select1').disabled=true;
      document.getElementById('deselect1').disabled=true;
      document.getElementById('select2').disabled=true;
      document.getElementById('deselect2').disabled=true;
       if(document.getElementById("btncreate")!=undefined && document.getElementById("btncreate")!=null){
        document.getElementById("btncreate").disabled=true;
      }
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
       if(document.getElementById("btncreate")!=undefined && document.getElementById("btncreate")!=null){
        document.getElementById("btncreate").disabled=false;
      }
    }
    if (elemDisallowEditValueOpt != null && elemDisallowEditValueOpt.value == '1') {
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

function isValidEntry(elemID) {
  var bValid = true;
  var elem = document.getElementById(elemID);
  
  if (elem != null) {
    bValid = !isNaN(elem.value);
  }
  
  return bValid;
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
function ValidateEntry(evt, src) {
    var elem = document.getElementById("valuetype");

    if (elem != null) {
        if (elem.value == "1") {
            return NumberCheck(evt, src, false);
        } else {
            return NumberCheck(evt, src, true);
        }            
    }
}

function ReformatForQty(isTiered) {
    var elem = document.getElementById("valuetype");
    var txtElem = null;
    var maxCt = 999, decPt = -1;
    var txtVal = "";
            
    if (elem != null) {
        if (elem.value == "1") {
            var i = (isTiered) ? 1 : 0;
            while (i < maxCt) {
                txtElem = document.getElementById("tier" + i);
                if (txtElem != null) {
                    txtVal = txtElem.value;
                    decPt = txtVal.indexOf(".");
                    if (decPt >= 0) {
                        txtElem.value = txtVal.substring(0, decPt);
                    }
                } else {
                    break;
                }
                i++;
            }
        }
    }
    showAdvancedOptions();
}
function submitShowAll() {
      var elem = document.getElementById("showall");

      if (elem != null) {
          elem.value = "true";
      }
      document.mainform.submit();
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
    populateConditionType();
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
      document.getElementById("searchLoadDiv").innerHTML = '<% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID)) %>';
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
                var selectedGroupValue=-1;
                if(document.getElementById('selected').options[0] !=undefined && document.getElementById('selected').options[0]!=null)
                    selectedGroupValue= document.getElementById('selected').options[0].value
                if(parseInt(selectedGroupValue) != parseInt(resultArr[1]))
                {
                    alert(responseArr[2]);
                    if(selectedGroupValue != -1 ){
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
       document.getElementById("searchLoadDiv").innerHTML = '&nbsp;';
   }

   function addNewGrouptoSelectbox(text,val)
   {
        var sel = document.getElementById('functionselect');
        var opt = document.createElement('option'); // create new option element
        // create text node to add to option element (opt)
        opt.appendChild( document.createTextNode(text) );
        opt.value = val; // set value property of opt
        sel.appendChild(opt); // add opt to end of select box (sel)
   
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
    If (Request.QueryString("save") <> "") Then

        ' determine if the previous product group values were changed for use in determining TCRMAStatusFlag value.
        ProdGroupChanged = False
        If Not (bUseTemplateLocks And bDisallowEditPg) Then
            MyCommon.QueryStr = "select LinkID, ExcludedID from OfferConditions with (NoLock) where ConditionID=" & ConditionID & " and deleted=0;"
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count > 0) Then
                ProdGroupChanged = ProductGroupID <> MyCommon.NZ(rst.Rows(0).Item("LinkID"), 0)
                ProdGroupChanged = ProdGroupChanged OrElse (ExcludedID <> MyCommon.NZ(rst.Rows(0).Item("ExcludedID"), 0))
            End If
        
            Select Case ProductGroupID
                Case 0
                    MyCommon.QueryStr = "update OfferConditions with (RowLock) set LinkID=0, ExcludedID=0 where ConditionID=" & ConditionID & " and deleted=0;"
                    MyCommon.LRT_Execute()
                Case 1
                    MyCommon.QueryStr = "update OfferConditions with (RowLock) set LinkID=" & ProductGroupID & ", ExcludedId=" & ExcludedID & " where ConditionID=" & ConditionID & " and deleted=0;"
                    MyCommon.LRT_Execute()
                Case Else
                    If (ExcludedID > 0) Then
                        MyCommon.QueryStr = "update OfferConditions with (RowLock) set LinkID=" & ProductGroupID & ", ExcludedId=" & ExcludedID & " where ConditionID=" & ConditionID & " and deleted=0;"
                    Else
                        MyCommon.QueryStr = "update OfferConditions with (RowLock) set LinkID=" & ProductGroupID & ", ExcludedId=0 where ConditionID=" & ConditionID & " and deleted=0;"
                    End If
                    MyCommon.LRT_Execute()
            End Select
        End If
    
        If Not (bUseTemplateLocks And bDisallowEditAdvanced) Then
            If (Request.QueryString("nodistribute") = "on") Then
                MyCommon.QueryStr = "update OfferConditions with (RowLock) set DoNotItemDistribute=1 where ConditionID=" & ConditionID & ";"
                MyCommon.LRT_Execute()
            Else
                MyCommon.QueryStr = "update OfferConditions with (RowLock) set DoNotItemDistribute=0 where ConditionID=" & ConditionID & ";"
                MyCommon.LRT_Execute()
            End If

            If (Request.QueryString("minorderitems") = "on") Then
                MyCommon.QueryStr = "update OfferConditions with (RowLock) set MinOrderItemsOnly=1 where ConditionID=" & ConditionID & ";"
                MyCommon.LRT_Execute()
            Else
                MyCommon.QueryStr = "update OfferConditions with (RowLock) set MinOrderItemsOnly=0 where ConditionID=" & ConditionID & ";"
                MyCommon.LRT_Execute()
            End If
    
            MyCommon.QueryStr = "update OfferConditions with (RowLock) set DiscountableItemsOnly =" & IIf(Request.QueryString("discitemsonly") = "1", "1", "0") & " where ConditionID=" & ConditionID & ";"
            MyCommon.LRT_Execute()
    
            MyCommon.QueryStr = "update OfferConditions with (RowLock) set PointsUseNetValue =" & IIf(Request.QueryString("UseNetForItems") = "1", "1", "0") & " where ConditionID=" & ConditionID & ";"
            MyCommon.LRT_Execute()
        End If

        If Not (bUseTemplateLocks And bDisallowEditRewards) Then
            If (Request.QueryString("granted") <> "") Then
                MyCommon.QueryStr = "update OfferConditions with (RowLock) set GrantTypeID=" & Request.QueryString("granted") & " where ConditionID=" & ConditionID & ";"
                MyCommon.LRT_Execute()
            End If
        End If

        If Not (bUseTemplateLocks And bDisallowEditValue) Then
            If (Request.QueryString("valuetype") <> "") Then
                MyCommon.QueryStr = "update OfferConditions with (RowLock) set QtyUnitType=" & Request.QueryString("valuetype") & " where ConditionID=" & ConditionID & ";"
                MyCommon.LRT_Execute()
            End If

            If (Request.QueryString("tier0") <> "" And Request.QueryString("Tiered") = "False") Then
                'dbo.pt_ConditionTiers_Update @ConditionID bigint, @TierLevel int,@AmtRequired decimal(12,3)
                MyCommon.QueryStr = "dbo.pt_ConditionTiers_Update"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@ConditionID", SqlDbType.BigInt).Value = ConditionID
                MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int).Value = 0
                MyCommon.LRTsp.Parameters.Add("@AmtRequired", SqlDbType.Decimal).Value = MyCommon.Extract_Val(Request.QueryString("tier0"))
                If (MyCommon.Extract_Val(Request.QueryString("tier0")) < 0) Then
                    infoMessage = Copient.PhraseLib.Lookup("condition.badvalue", LanguageID)
                Else
                    MyCommon.LRTsp.ExecuteNonQuery()
                End If
                MyCommon.Close_LRTsp()
            ElseIf (Request.QueryString("Tiered") = "True") Then
                ' delete the current tier ammounts
                MyCommon.QueryStr = "delete from ConditionTiers with (RowLock) where ConditionID=" & ConditionID
                MyCommon.LRT_Execute()
                Dim x As Integer
                For x = 1 To NumTiers
                    'Response.Write("->@ConditionID: " & ConditionID & " @TierLevel: " & x & " @AmtRequired: " & Request.QueryString("tier" & x) & "<br />")
                    MyCommon.QueryStr = "dbo.pt_ConditionTiers_Update"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@ConditionID", SqlDbType.BigInt).Value = MyCommon.Extract_Val(ConditionID)
                    MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int).Value = x
                    MyCommon.LRTsp.Parameters.Add("@AmtRequired", SqlDbType.Decimal).Value = MyCommon.Extract_Val(Request.QueryString("tier" & x))
                    If (x > 1) And (Int(MyCommon.Extract_Val(Request.QueryString("tier" & x))) < Int(MyCommon.Extract_Val(Request.QueryString("tier" & (x - 1))))) Then
                        infoMessage = Copient.PhraseLib.Lookup("condition.tiervalues", LanguageID)
                    ElseIf (MyCommon.Extract_Val(Request.QueryString("tier" & x)) < 0) Then
                        infoMessage = Copient.PhraseLib.Lookup("condition.badvalue", LanguageID)
                    End If
                    MyCommon.LRTsp.ExecuteNonQuery()
                    MyCommon.Close_LRTsp()
                Next
            End If
        End If

        If (Request.QueryString("IsTemplate") = "IsTemplate") Then
            ' time to update the status bits for the templates
            Dim form_Disallow_Edit As Integer = 0
            Dim form_require_pg As Integer = 0
            Dim iDisallowEditPg As Integer = 0
            Dim iDisallowEditValue As Integer = 0
            Dim iDisallowEditRewards As Integer = 0
            Dim iDisallowEditAdvanced As Integer = 0
      
            Disallow_Edit = False
            RequirePG = False
            bDisallowEditValue = False
            bDisallowEditRewards = False
            bDisallowEditAdvanced = False
            bDisallowEditPg = False
      
            If (Request.QueryString("Disallow_Edit") = "on") Then
                form_Disallow_Edit = 1
                Disallow_Edit = True
            End If
            If (Request.QueryString("require_pg") <> "") Then
                form_require_pg = 1
                RequirePG = True
            End If
            If (Request.QueryString("DisallowEditPg") = "on") Then
                iDisallowEditPg = 1
                bDisallowEditPg = True
            End If
            If (Request.QueryString("DisallowEditValue") = "on") Then
                iDisallowEditValue = 1
                bDisallowEditValue = True
            End If
            If (Request.QueryString("DisallowEditRewards") = "on") Then
                iDisallowEditRewards = 1
                bDisallowEditRewards = True
            End If
            If (Request.QueryString("DisallowEditAdvanced") = "on") Then
                iDisallowEditAdvanced = 1
                bDisallowEditAdvanced = True
            End If
            ' both requiring and locking the product group is not permitted 
            If (form_require_pg = 1 AndAlso (form_Disallow_Edit = 1 Or iDisallowEditPg = 1)) Then
                infoMessage = Copient.PhraseLib.Lookup("offer-con.lockeddenied", LanguageID)
            Else
                MyCommon.QueryStr = "update OfferConditions with (RowLock) set Disallow_Edit=" & form_Disallow_Edit & _
                ",RequiredFromTemplate=" & form_require_pg & _
                ",DisallowEdit1=" & iDisallowEditPg & _
                ",DisallowEdit2=" & iDisallowEditValue & _
                ",DisallowEdit3=" & iDisallowEditRewards & _
                ",DisallowEdit4=" & iDisallowEditAdvanced & _
                " where ConditionID=" & ConditionID
                MyCommon.LRT_Execute()
            End If
        End If
    
        ' update the flags
        ' determine if we need to set TCRMAStatusFlag to 2 or leave it 3
        Dim TempFlagHolder As Integer = 2
    
        If (ProdGroupChanged) Then
            TempFlagHolder = 3
        Else
            MyCommon.QueryStr = "select TCRMAStatusFlag from offerConditions with (NoLock) where ConditionID=" & ConditionID
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count > 0) Then
                If rst.Rows(0).Item("TCRMAStatusFlag") = 3 Then
                    TempFlagHolder = 3
                End If
            End If
        End If
    
    MyCommon.QueryStr = "update OfferConditions with (RowLock) set TCRMAStatusFlag=" & TempFlagHolder & ",CRMAStatusFlag=2 where ConditionID=" & ConditionID
    MyCommon.LRT_Execute()
    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-product", LanguageID))
    If (infoMessage = "") Then
      CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
    End If
  ElseIf (GetCgiValue("add") <> "") Then
        
    MyCommon.QueryStr = "select LinkID, ExcludedID from OfferConditions with (NoLock) where ConditionID=" & ConditionID & " and deleted=0;"
    rst = MyCommon.LRT_Select   
    
	GName = GetCgiValue("modprodgroupname")
        
    If (rst.Rows(0).Item("LinkID") = 0) Then 'Create new product group
                  
      MyCommon.QueryStr = "SELECT ProductGroupID FROM ProductGroups with (NoLock) WHERE Name = '" & IIf(GName.Contains("'"), GName.Replace("'", "''"), GName) & "' AND Deleted=0"
      rst1 = MyCommon.LRT_Select
      If (rst1.Rows.Count > 0) Then
        ProductGroupID = rst1.Rows(0).Item("ProductGroupID")
      Else
	    MyCommon.QueryStr = "dbo.pt_ProductGroups_Insert"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 200).Value =  GName
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
      ProductGroupID = rst.Rows(0).Item("LinkID")
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
    End If  
  ElseIf (GetCgiValue("mremove") <> "") Then
    ' desired product remove from group  dbo.pt_GroupMembership_Delete_ByID  @MembershipID bigint
    ' dbo.pt_ProdGroupItems_Delete  @ExtProductID nvarchar(20), @ProductGroupID bigint, @ProductTypeID int, @Status int OUTPUT
	MyCommon.QueryStr = "select LinkID, ExcludedID from OfferConditions with (NoLock) where ConditionID=" & ConditionID & " and deleted=0;"
    rst = MyCommon.LRT_Select
    
    If (rst.Rows(0).Item("LinkID") = 0) Then
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
    
  If (Request.QueryString("save") <> "" Or Request.QueryString("select1") <> "" Or Request.QueryString("deselect1") <> "" Or Request.QueryString("select2") <> "" Or Request.QueryString("deselect2") <> "") Then
    MyCommon.QueryStr = "update OfferConditions with (RowLock) set TCRMAStatusFlag=3, CRMAStatusFlag=2 where ConditionID=" & ConditionID & ";"
    MyCommon.LRT_Execute()
    MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where OfferID=" & OfferID & ";"
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
    MyCommon.QueryStr = "select count(*) as GCount from ProdGroupItems PGI with (NoLock) inner join OfferConditions OC on OC.LinkID = PGI.ProductGroupID " & _
                        "where OC.ConditionID = " & ConditionID & " And OC.Deleted = 0 And PGI.Deleted = 0"
    rst = MyCommon.LRT_Select()
    For Each row In rst.Rows
      GroupSize = row.Item("GCount")
    Next
        
    MyCommon.QueryStr = "select count(*) as PCount from ProdGroupItems PGI with (NoLock) inner join products PRD on PGI.productid = PRD.productid " & _
                        "inner join OfferConditions OC with (NoLock) on OC.LinkID = PGI.ProductGroupID " & _
                        "where OC.ConditionID = " & ConditionID & " And PGI.Deleted = 0 And (PRD.Description IS NULL OR  PRD.Description = '') And OC.Deleted = 0"
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
                      "inner join OfferConditions as OC with (NoLock) on OC.LinkID = gm.ProductGroupID " & _
                      "where OC.ConditionID = " & ConditionID & " and GM.Deleted=0 and OC.Deleted =0 and IsNull(GM.ExtHierarchyID, '')='' " & _
                      "and IsNull(GM.ExtNodeID, '')='' " & If(bBlankDescProd, sBlankDescOrderByStr, " order by ExtProductID;")
    
  End If
    
    
    
  rstItems = MyCommon.LRT_Select()
  ListBoxSize = rstItems.Rows.Count
  
    Send("<script type=""text/javascript"">")
    Send("function ChangeParentDocument() { ")
    Send("  if (opener != null) {")
    Send("    var newlocation = 'offer-con.aspx?OfferID=" & OfferID & "'; ")
    Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
    Send("     opener.location = 'offer-con.aspx?OfferID=" & OfferID & "'; ")
    Send("  }")
    Send("  }")
    Send("} ")
    Send("</script>")
%>
<form action="offer-con-product.aspx" id="mainform" name="mainform"  onsubmit="saveForm()">
<div id="intro">
    <input type="hidden" id="OfferID" name="OfferID" value="<% sendb(OfferID) %>" />
    <input type="hidden" id="Name" name="Name" value="<% sendb(Name) %>" />
    <input type="hidden" id="ConditionID" name="ConditionID" value="<% sendb(ConditionID) %>" />
    <input type="hidden" id="ProductGroupID" name="ProductGroupID" value="<% sendb(ProductGroupID) %>" />
    <input type="hidden" id="ExcludedID" name="ExcludedID" value="<% sendb(ExcludedID) %>" />
    <input type="hidden" id="IsTemplate" name="IsTemplate" value="<% 
            if(istemplate)then 
            sendb("IsTemplate")
            else 
            sendb("Not") 
            end if
        %>" />
  <input type="hidden" name="showall" id="showall" value="<% sendb( ShowAllItems.ToString.ToLower)%>" />
  <%If MyCommon.Fetch_SystemOption(97) = "1" Then%>
  <input type="hidden" id="NumericOnly" name="NumericOnly" value="true" />
  <%End If%>
    <%
        If (IsTemplate) Then
            Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.productcondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
        Else
            Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.productcondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
        End If
        If (bUseTemplateLocks And bDisallowEditAdvanced) Then
            Send("<input type=""hidden"" id=""DisallowEditAdvancedOpt"" name=""DisallowEditAdvancedOpt"" value=""1"" />")
        Else
            Send("<input type=""hidden"" id=""DisallowEditAdvancedOpt"" name=""DisallowEditAdvancedOpt"" value=""0"" />")
        End If
        If (bUseTemplateLocks And bDisallowEditPg) Then
            Send("<input type=""hidden"" id=""DisallowEditPgOpt"" name=""DisallowEditPgOpt"" value=""1"" />")
        Else
            Send("<input type=""hidden"" id=""DisallowEditPgOpt"" name=""DisallowEditPgOpt"" value=""0"" />")
        End If
        If (bUseTemplateLocks And bDisallowEditValue) Then
            Send("<input type=""hidden"" id=""DisallowEditValueOpt"" name=""DisallowEditValueOpt"" value=""1"" />")
        Else
            Send("<input type=""hidden"" id=""DisallowEditValueOpt"" name=""DisallowEditValueOpt"" value=""0"" />")
        End If
    %>
    <div id="controls">
        <% If (IsTemplate) Then%>
        <span class="temp">
            <input type="checkbox" class="tempcheck" id="temp-employees" name="Disallow_Edit"
                <% if(disallow_edit)then sendb(" checked=""checked""") %> />
            <label for="temp-employees">
                <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
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
      value="directadd" /><label for="pgselectortype2">Add products to condition</label>
  <%Else%>  
    <input type="radio" id="pgselectortype2" name="pgselectortype" onclick="javascript:ProductGroupTypeSelection();"
      <% if(Not ByExistingPGSelector) then sendb(" checked=""checked""") %> value="directadd" /><label
        for="pgselectortype2">Add products to condition</label>
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
          <input type="text" id="modprodgroupname" style="width: 347px;" name="modprodgroupname" maxlength="200"
            value="<% Sendb(GName) %>"/><br />
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
            for="prodaddselector2"> <% Sendb(Copient.PhraseLib.Lookup("gen.addproductlist", LanguageID))%></label>
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
             Sendb("<textarea name=""pasteproducts"" id=""pasteproducts""  autofocus=""autofocus"" style=""width: 290px; height: 150px"">")
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
          <% Sendb(Copient.PhraseLib.Lookup("term.productgroup", LanguageID))%>
        </span>
        <% If (IsTemplate) Then%>
        <span class="tempRequire">
          <input type="checkbox" class="tempcheck" id="require_pg1" name="require_pg" <% if(requirepg)then sendb(" checked=""checked""") %> />
          <label for="require_pg">
            <% Sendb(Copient.PhraseLib.Lookup("term.required", LanguageID))%></label>
        </span><span class="tempLocked">
          <input type="checkbox" class="tempcheck" id="DisallowEditPg1" name="DisallowEditPg"
            <% if(bDisallowEditPg)then send(" checked=""checked""") %> />
          <label for="temp-Tiers">
            <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
        </span>
        <% ElseIf (bUseTemplateLocks) Then%>
        <% If (RequirePG) Then%>
        <span class="tempRequire">
          <input type="checkbox" class="tempcheck" id="require_pg2" name="require_pg" disabled="disabled"
            checked="checked" />
          <label for="require_pg">
            <% Sendb(Copient.PhraseLib.Lookup("term.required", LanguageID))%></label>
        </span>
        <% ElseIf (bDisallowEditPg) Then%>
        <span class="tempRequire">
          <input type="checkbox" class="tempcheck" id="DisallowEditPg2" name="DisallowEditPg"
            disabled="disabled" checked="checked" />
          <label for="temp-Tiers">
            <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
        </span>
        <% End If%>
        <% End If%>
      </h2>
      <input type="radio" id="functionradio1" name="functionradio" <% if(MyCommon.Fetch_SystemOption(175)= "1" ) then sendb(" checked=""checked""") %>
        <% sendb(disabledattribute) %> /><label for="functionradio1"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
      <input type="radio" id="functionradio2" name="functionradio" <% if(MyCommon.Fetch_SystemOption(175)= "2" ) then sendb(" checked=""checked""") %>
        <% sendb(disabledattribute) %> /><label for="functionradio2"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
      <input type="text" class="medium" id="functioninput" name="functioninput" maxlength="100"
        onkeyup="javascript:xmlPostTimer('OfferFeeds.aspx','ProductGroupsCM');" value=""
        <% sendb(disabledattribute) %> />
      <% If (bCreateGroupOrProgramFromOffer AndAlso Logix.UserRoles.CreateProductGroups) Then%>
           <input class="regular" name="btncreate" id="btncreate" type="button" value="<% Sendb(Copient.PhraseLib.Lookup("term.create", LanguageID))%>" onclick="javascript:handleCreateClick('btncreate');" />
      <% End If%>   
      <br />
      <div id="searchLoadDiv" style="display:block;">&nbsp;</div>
	  <div id="pgList" class="column3x1">
        <select class="long" id="functionselect" name="functionselect" size="20" <% sendb(disabledattribute) %>>
          <%
            Dim topString As String = ""
            If RECORD_LIMIT > 0 Then topString = "top " & RECORD_LIMIT
                        
            If bStoreUser Then
              sJoin = "Full Outer Join ProductGroupLocUpdate pglu with (NoLock) on pg.ProductGroupID=pglu.ProductGroupID " 
              wherestr = " and (LocationID in (" & sValidLocIDs & ") or (CreatedByAdminID in (" & sValidSU & ") and Isnull(LocationID,0)=0)) " 
            End If
            
                        Dim orderBy = ""
                        If(MyCommon.Fetch_SystemOption(235) = "1") Then
                          orderBy = " order by AnyProduct desc, Name"
                        Else
                          orderBy = " order by AnyProduct desc, ProductGroupID desc, Name asc"
                        End If
                      
                        Dim Limiter As String
                        If (ExcludedItem) Then Limiter = "and ProductGroupID <> " & ExcludedItem
                        If (SelectedItem) Then Limiter = Limiter & " and ProductGroupID <> " & SelectedItem
            		         MyCommon.QueryStr = "select " & topString & " pg.ProductGroupID,CreatedDate,Name,PhraseID,LastUpdate,AnyProduct from ProductGroups pg with (NoLock) " & sJoin & " where Deleted=0" & Limiter & wherestr
                         If(bEnableRestrictedAccessToUEOfferBuilder) Then MyCommon.QueryStr &= " and isnull(pg.TranslatedFromOfferID,0) = 0 " 
                        
                        MyCommon.QueryStr = MyCommon.QueryStr & orderBy
                        
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
          If(MyCommon.Fetch_SystemOption(235) = "1") Then
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
              MyCommon.QueryStr = "select LinkID,ExcludedID,C.Name,C.PhraseID from OfferConditions with (NoLock) join ProductGroups as C with (NoLock) on LinkID=ProductGroupID and ConditionID=" & ConditionID & ";"
              rst = MyCommon.LRT_Select
              MyCommon.QueryStr = "select LinkID,ExcludedID,C.Name,C.PhraseID from OfferConditions with (NoLock) join ProductGroups as C with (NoLock) on ProductGroupID=ExcludedID where not(ExcludedID=0) and ConditionID=" & ConditionID & ";"
              rst2 = MyCommon.LRT_Select
              If (rst.Rows.Count = 0 And rst2.Rows.Count = 0) Then
                IsTransactionLevel = True
              Else
                For Each row In rst.Rows
                  Sendb("<option value=""" & row.Item("LinkID") & """>")
                  If (MyCommon.NZ(row.Item("PhraseID"), 0) > 0) Then
                    Sendb(Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID))
                  Else
                    Sendb(row.Item("Name"))
                  End If
                  Send("</option>")
                  SelectedItem = row.Item("LinkID")
                Next
              End If
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
              MyCommon.QueryStr = "select LinkID,ExcludedID,C.Name,C.PhraseID from OfferConditions with (NoLock) join ProductGroups as C with (NoLock) on ProductGroupID=ExcludedID  where not(ExcludedID=0) and ConditionID=" & ConditionID & ";"
              rst2 = MyCommon.LRT_Select
              For Each row2 In rst2.Rows
                Sendb("<option value=""" & row2.Item("ExcludedID") & """>")
                If (MyCommon.NZ(row2.Item("PhraseID"), 0) > 0) Then
                  Sendb(Copient.PhraseLib.Lookup(row2.Item("PhraseID"), LanguageID))
                Else
                  Sendb(row2.Item("Name"))
                End If
                Send("</option>")
                ExcludedItem = row2.Item("ExcludedID")
              Next
            %>
            <%
              If rst.Rows.Count = 0 Then
                IsTransactionLevel = True
              End If
            %>
          </select>
        </div>
        <hr class="hidden" />
      </div>
    </div>
    <div id="column1">
      <div class="box" id="value">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.value", LanguageID))%>
          </span>
        </h2>
        <% If (IsTemplate Or (bUseTemplateLocks And bDisallowEditValue)) Then%>
        <span class="temp">
          <input type="checkbox" class="tempcheck" id="DisallowEditValue" name="DisallowEditValue"
            <% if(bDisallowEditValue)then send(" checked=""checked""") %> <% If(bUseTemplateLocks) Then sendb(" disabled=""disabled""") %> />
          <label for="temp-Tiers">
            <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
        </span>
        <br class="printonly" />
        <% End If%>
        <label for="valuetype">
          <% Sendb(Copient.PhraseLib.Lookup("condition.valuetype", LanguageID))%></label>
        <br />
        <%
          Dim firstPass As Boolean
          Dim tieredParm As String
          firstPass = True
          MyCommon.QueryStr = "select LinkID,Tiered,O.Numtiers,QtyUnitType,O.OfferID,CT.TierLevel,CT.AmtRequired from OfferConditions as OC with (NoLock) left  join Offers as O with (NoLock) on O.OfferID=OC.OfferID left  join ConditionTiers as CT with (NoLock) on OC.ConditionID=CT.ConditionID where OC.ConditionID=" & ConditionID & ";"
          rst = MyCommon.LRT_Select()
          For Each row In rst.Rows
            If (firstPass) Then
              If (row.Item("Tiered") = 0) Then
                tieredParm = "false"
              Else
                tieredParm = "true"
              End If
        %>
        <select id="valuetype" name="valuetype" onchange="ReformatForQty(<% Sendb(tieredParm) %>);"
          <% If(bUseTemplateLocks and bDisallowEditValue) Then sendb(" disabled=""disabled""") %>>
          <% If Not IsTransactionLevel Then%>
          <option value="1" <%if(row.item("qtyunittype") = 1)then sendb(" selected=""selected""") %>>
            <% Sendb(Copient.PhraseLib.Lookup("term.quantity", LanguageID) & " (" & Copient.PhraseLib.Lookup("term.count", LanguageID) & ")")%>
          </option>
          <option value="2" <%if(row.item("qtyunittype") = 2)then sendb(" selected=""selected""") %>>
            <% Sendb(Copient.PhraseLib.Lookup("term.amount", LanguageID) & " (" & Copient.PhraseLib.Lookup("term.dollars", LanguageID) & ")")%>
          </option>
          <option value="3" <%if(row.item("qtyunittype") = 3)then sendb(" selected=""selected""") %>>
            <% Sendb(Copient.PhraseLib.Lookup("term.weight", LanguageID) & " (" & Copient.PhraseLib.Lookup("term.pounds", LanguageID) & ")")%>
          </option>
          <option value="4" <%if(row.item("qtyunittype") = 4)then sendb(" selected=""selected""") %>>
            <% Sendb(Copient.PhraseLib.Lookup("term.volume", LanguageID) & " (" & Copient.PhraseLib.Lookup("term.gallons", LanguageID) & ")")%>
          </option>
          <% Else%>
          <option value="7" <%if(row.item("qtyunittype") = 7)then sendb(" selected=""selected""") %>>
            <% Sendb(Copient.PhraseLib.Lookup("term.minimumorder", LanguageID))%>
          </option>
          <option value="6" <%if(row.item("qtyunittype") = 6)then sendb(" selected=""selected""") %>>
            <% Sendb(Copient.PhraseLib.Lookup("term.transactiontotal", LanguageID))%>
          </option>
          <option value="10" <%if(row.item("qtyunittype") = 10)then sendb(" selected=""selected""") %>>
            <% Sendb(Copient.PhraseLib.Lookup("term.trxdisctotals", LanguageID))%>
          </option>
          <% End If%>
        </select>
        <br />
        <% 
          firstPass = False
        End If
      Next
        %>
        <br class="half" />
        <label for="tier0">
          <% Sendb(Copient.PhraseLib.Lookup("condition.valueneeded", LanguageID))%></label>
        <br />
        <%
          MyCommon.QueryStr = "select LinkID,Tiered,O.Numtiers,QtyUnitType,O.OfferID,CT.TierLevel,CT.AmtRequired from OfferConditions as OC with (NoLock) left  join Offers as O with (NoLock) on O.OfferID=OC.OfferID left  join ConditionTiers as CT with (NoLock) on OC.ConditionID=CT.ConditionID where OC.ConditionID=" & ConditionID & ";"
          rst = MyCommon.LRT_Select()
          Dim q As Integer
          Dim amtValue As String
          q = 1
          If (bUseTemplateLocks And bDisallowEditValue) Then
            sDisabled = " disabled=""disabled"""
          Else
            sDisabled = ""
          End If
          For Each row In rst.Rows
            'Convert value to integer if it is a quantity unit type
            If (row.Item("QtyUnitType") = 1) Then
              amtValue = CStr(CInt(row.Item("AmtRequired")))
            Else
              amtValue = row.Item("AmtRequired")
            End If
            If (row.Item("Tiered") = 0) Then
              Send("<input class=""shorter"" id=""tier0"" name=""tier0"" type=""text"" maxlength=""9"" value=""" & amtValue & """ onkeydown=""return ValidateEntry(event, this);""" & sDisabled & " /><br />")
            Else
              Send("<label for=""tier" & q & """><b>" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & q & ":</b></label> <input class=""shorter"" id=""tier" & q & """ name=""tier" & q & """ type=""text"" value=""" & amtValue & """ onkeydown=""return ValidateEntry(event, this);""" & sDisabled & " /><br />")
              Tiered = True
            End If
            q = q + 1
          Next
          Send("<input type=""hidden"" name=""NumTiers"" value=""" & row.Item("NumTiers") & """ />")
          Send("<input type=""hidden"" name=""Tiered"" value=""" & row.Item("Tiered") & """ />")
        %>
        &nbsp;<br />
        <%
          MyCommon.QueryStr = "select LinkID,ExcludedID,MinOrderItemsOnly,GrantTypeID,DoNotItemDistribute from OfferConditions with (NoLock) where ConditionID=" & ConditionID & ";"
          rst = MyCommon.LRT_Select
          For Each row In rst.Rows
          Next
        %>
        <hr class="hidden" />
      </div>
    </div>
    <div id="gutter">
    </div>
    <div id="column2">
      <div class="box" id="grants" <%if(tiered)then sendb("style=""display: none; visibility: hidden;""") %>>
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.rewards", LanguageID))%>
          </span>
        </h2>
        <% If (IsTemplate Or IsTemplate Or (bUseTemplateLocks And bDisallowEditRewards)) Then%>
        <span class="temp">
          <input type="checkbox" class="tempcheck" id="DisallowEditRewards" name="DisallowEditRewards"
            <% if(bDisallowEditRewards)then send(" checked=""checked""") %><% If(bUseTemplateLocks) Then sendb(" disabled=""disabled""") %> />
          <label for="temp-Tiers">
            <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
        </span>
        <br class="printonly" />
        <% End If%>
        <% Sendb(Copient.PhraseLib.Lookup("condition.rewardsgranted", LanguageID))%>
        <br />
        <input class="radio" id="eachtime" name="granted" value="3" type="radio" <% if(row.item("granttypeid")=3)then sendb(" checked=""checked""") %><% If(bUseTemplateLocks and bDisallowEditRewards) Then sendb(" disabled=""disabled""") %> />
        <label for="eachtime">
          <% Sendb(Copient.PhraseLib.Lookup("condition.eachtime", LanguageID))%></label>
        <br />
        <input class="radio" id="equalto" name="granted" value="1" type="radio" <% if(row.item("granttypeid")=1)then sendb(" checked=""checked""") %><% If(bUseTemplateLocks and bDisallowEditRewards) Then sendb(" disabled=""disabled""") %> />
        <label for="equalto">
          <% Sendb(Copient.PhraseLib.Lookup("condition.equalto", LanguageID))%></label>
        <br />
        <input class="radio" id="greaterthan" name="granted" value="2" type="radio" <% if(row.item("granttypeid")=2)then sendb(" checked=""checked""") %><% If(bUseTemplateLocks and bDisallowEditRewards) Then sendb(" disabled=""disabled""") %> />
        <label for="greaterthan">
          <% Sendb(Copient.PhraseLib.Lookup("condition.greaterthan", LanguageID))%></label>
        <br />
        <hr class="hidden" />
      </div>
    </div>
    <div id="column6">
      <div class="box" id="options">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.advancedoptions", LanguageID))%>
          </span>
        </h2>
        <% If (IsTemplate Or (bUseTemplateLocks And bDisallowEditAdvanced)) Then%>
        <span class="temp">
          <input type="checkbox" class="tempcheck" id="DisallowEditAdvanced" name="DisallowEditAdvanced"
            <% if(bDisallowEditAdvanced)then send(" checked=""checked""") %><% If(bUseTemplateLocks) Then sendb(" disabled=""disabled""") %> />
          <label for="temp-Tiers">
            <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
        </span>
        <br class="printonly" />
        <% End If%>
        <%
          MyCommon.QueryStr = "select LinkID,QtyUnitType,ExcludedID,MinOrderItemsOnly,GrantTypeID,DoNotItemDistribute,DiscountableItemsOnly,PointsUseNetValue from OfferConditions with (NoLock) where ConditionID=" & ConditionID & ";"
          rst = MyCommon.LRT_Select
          For Each row In rst.Rows
            If MyCommon.Fetch_CM_SystemOption(3) Then
              If (row.Item("DoNotItemDistribute")) Then
                sChecked = " checked=""checked"""
              Else
                sChecked = ""
              End If
              If (bUseTemplateLocks And bDisallowEditAdvanced) Then
                sDisabled = " disabled=""disabled"""
              Else
                sDisabled = ""
              End If
              Send("<input class=""checkbox""" & sDisabled & sChecked & " id=""nodistribute"" name=""nodistribute"" type=""checkbox"" />")
              Send("<label for=""nodistribute"">" & Copient.PhraseLib.Lookup("condition.itemdistribute", LanguageID) & "</label><br />")
            End If
            
                    If ((row.Item("LinkID") = "1" AndAlso MyCommon.NZ(row.Item("ExcludedID"), 0) = 0) Or IsTransactionLevel) Then
                        sDisabled = " disabled=""disabled"""
                    Else
                        If (bUseTemplateLocks And bDisallowEditAdvanced) Then
                            sDisabled = " disabled=""disabled"""
                        Else
                            sDisabled = ""
                        End If
                    End If
                    If (row.Item("MinOrderItemsOnly")) Then
                        sChecked = " checked=""checked"""
                    Else
                        sChecked = ""
                    End If
                    Send("<input class=""checkbox""" & sDisabled & sChecked & " id=""minorderitems"" name=""minorderitems"" type=""checkbox"" />")
                    Send("<label for=""minorderitems"">" & Copient.PhraseLib.Lookup("condition.minimumorder", LanguageID) & "</label><br />")
            
                    If (IsTransactionLevel) Then
                        sDisabled = " disabled=""disabled"""
                    Else
                        If (bUseTemplateLocks And bDisallowEditAdvanced) Then
                            sDisabled = " disabled=""disabled"""
                        Else
                            sDisabled = ""
                        End If
                    End If
                    If (MyCommon.NZ(row.Item("DiscountableItemsOnly"), False)) Then
                        sChecked = " checked=""checked"""
                    Else
                        sChecked = ""
                    End If
                    Send("<input class=""checkbox""" & sDisabled & sChecked & " id=""discitemsonly"" name=""discitemsonly"" type=""checkbox"" value=""1"" />")
                    Send("<label for=""discitemsonly"">" & Copient.PhraseLib.Lookup("term.discitemsonly", LanguageID) & "</label><br />")
            
                    If (IsTransactionLevel Or Not (row.Item("qtyunittype") = 2 Or row.Item("qtyunittype") = 1)) Then
                        sDisabled = " disabled=""disabled"""
                    Else
                        If (bUseTemplateLocks And bDisallowEditAdvanced) Then
                            sDisabled = " disabled=""disabled"""
                        Else
                            sDisabled = ""
                        End If
                    End If
                    If (MyCommon.NZ(row.Item("PointsUseNetValue"), False)) Then
                        sChecked = " checked=""checked"""
                    Else
                        sChecked = ""
                    End If
                    Send("<input class=""checkbox""" & sDisabled & sChecked & " id=""UseNetForItems"" name=""UseNetForItems"" type=""checkbox"" value=""1"" />")
                    Send("<label for=""UseNetForItems"">" & Copient.PhraseLib.Lookup("term.usenet", LanguageID) & "</label><br />")

                    Send("</div>")
                Next
            %>
        </div>
    </div>
  </div>
</div>
</form>
<script type="text/javascript">
<% If (CloseAfterSave) Then %>
    window.close();
<% End If %>

 document.getElementById("pasteproducts").focus();
 preventEnterSubmit;
 function preventEnterSubmit(e) {
    if (e.which == 13) {
        var $targ = $(e.target);

        if (!$targ.is("textarea") && !$targ.is(":button,:submit")) {
            var focusNext = false;
            $(this).find(":input:visible:not([disabled],[readonly]), a").each(function(){
                if (this === e.target) {
                    focusNext = true;
                }
                else if (focusNext){
                    $(this).focus();
                    return false;
                }
            });

            return false;
        }
    }
}

updateButtons();
removeUsed(false);
populateConditionType();
  
if (document.getElementById("functionselect") != null) {
  fullSelect = document.getElementById("functionselect").cloneNode(true);
}

document.getElementById("select1").onclick=select1_onclick;
document.getElementById("deselect1").onclick=deselect1_onclick;
document.getElementById("select2").onclick=select2_onclick;
document.getElementById("deselect2").onclick=deselect2_onclick;
function select1_onclick() {
  handleSelectClick('select1')
  populateConditionType();
  showAdvancedOptions();
}
function deselect1_onclick() {
  handleSelectClick('deselect1')
  populateConditionType();
  showAdvancedOptions();
}
function select2_onclick() {
  handleSelectClick('select2')
  populateConditionType();
  showAdvancedOptions();
}
function deselect2_onclick() {
  handleSelectClick('deselect2')
  populateConditionType();
  showAdvancedOptions();
}  
</script>
<%
done:
    MyCommon.Close_LogixRT()
    Send_BodyEnd("mainform", "functioninput")
    MyCommon = Nothing
    Logix = Nothing
%>
