<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: offer-rew-autogiftreceipt.aspx 
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
  Dim rst2 As DataTable
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
  Dim TransactionLevelSelected As Boolean = False
  Dim infoMessage As String = ""
  Dim DistPeriod As Integer
  Dim UseSpecialPricing As Boolean
  Dim SPRepeatAtOccur As Integer
  Dim ValueRadio As Integer
  Dim q As Integer
  Dim x As Integer
  Dim Tiered As Integer
  Dim SponsorID As Integer
  Dim PromoteToTransLevel As Boolean
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
  Dim bDisallowEditDist As Boolean = False
  Dim bDisallowEditLimit As Boolean = False
  Dim sDisabled As String
  Dim sXmlText As String
  Dim AdvancedLimitID As Long
  Dim RECORD_LIMIT As Integer = GroupRecordLimit '500
  Dim topString As String = ""
  
  If RECORD_LIMIT > 0 Then topString = "top " & RECORD_LIMIT
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "offer-rew-autogiftreceipt.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  OfferID = Request.QueryString("OfferID")
  RewardID = Request.QueryString("RewardID")
  NumTiers = Request.QueryString("NumTiers")
  ProductGroupID = MyCommon.Extract_Val(Request.QueryString("ProductGroupID"))
  ExcludedID = MyCommon.Extract_Val(Request.QueryString("ExcludedID"))
  
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  
  ' fetch the name
  MyCommon.QueryStr = "Select Name,IsTemplate,FromTemplate from Offers with (NoLock) where OfferID=" & OfferID
  rst = MyCommon.LRT_Select
  If (rst.Rows.Count > 0) Then
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
      bDisallowEditDist = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit6"), False)
      bDisallowEditLimit = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit7"), False)
      If bUseTemplateLocks Then
        If Disallow_Edit Then
          bDisallowEditPg = True
          bDisallowEditSpon = True
          bDisallowEditDist = True
          bDisallowEditLimit = True
        Else
          Disallow_Edit = bDisallowEditPg And bDisallowEditSpon And _
                          bDisallowEditDist And bDisallowEditLimit
        End If
      End If
    End If
  End If
  
  Send_HeadBegin("term.offer", "term.autogiftreceiptreward", OfferID)
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
      document.getElementById("pgList").innerHTML = '<select class="longer" id="functionselect" name="functionselect" size="12"<% sendb(disabledattribute) %>>' + str + '</select>';
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
      document.getElementById("pgList").innerHTML = '<select class="longer" id="functionselect" name="functionselect" size="12"<% sendb(disabledattribute) %>>&nbsp;</select>';
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

<%
    MyCommon.QueryStr = "select " & topString & " ProductGroupID,Name from ProductGroups with (NoLock) where Deleted=0 order by AnyProduct desc, Name"
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
    var selectedList, excludedList;

    document.getElementById("functionselect").size = "10";
    
    // Set references to the form elements
    selectObj = document.forms[0].functionselect;
    textObj = document.forms[0].functioninput;
    selectedList = document.getElementById("selected");
    excludedList = document.getElementById("excluded");

    // Remember the function list length for loop speedup
    functionListLength = functionlist.length;

    // Set the search pattern depending
    if(document.forms[0].functionradio[0].checked == true)
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
    for(i = 0; i < functionListLength; i++)
    {
        if(functionlist[i].search(re) != -1)
        {
            if (vallist[i] != "" && (selectedList.options.length < 1 || vallist[i] != selectedList.options[0].value)
              && (excludedList.options.length < 1 || vallist[i] != excludedList.options[0].value) ) {
                selectObj[numShown] = new Option(functionlist[i],vallist[i]);
                if (vallist[i] == 1) {
                    selectObj[numShown].style.fontWeight = 'bold';
                    selectObj[numShown].style.color = 'brown';
                }
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
    
    document.getElementById("ProductGroupID").value = selectList;
    document.getElementById("ExcludedID").value = excludededList;
    // alert(htmlContents);
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
  
  If infoMessage = "" Then
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
              MyCommon.QueryStr = "update OfferRewards with (RowLock) set AdvancedLimitID=" & AdvancedLimitID & _
                                  ",RewardDistLimitVarID=" & VarID & _
                                  ",RewardDistPeriod=0" & _
                                  " where RewardID=" & RewardID & ";"
              MyCommon.LRT_Execute()
            Else
              MyCommon.QueryStr = "update OfferRewards with (RowLock) set AdvancedLimitID=" & AdvancedLimitID & _
                                  ",RewardDistPeriod=0" & _
                                  " where RewardID=" & RewardID & ";"
              MyCommon.LRT_Execute()
            End If
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
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.autogiftreceipt", LanguageID))

      MyCommon.QueryStr = "update OfferRewards with (RowLock) set TCRMAStatusFlag=3,CMOAStatusFlag=2 where RewardID=" & RewardID
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where offerid=" & OfferID
      MyCommon.LRT_Execute()

      CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
    End If
  End If
  
  Send("<script type=""text/javascript"">")
    Send("function ChangeParentDocument() { ")
    Send("  if (opener != null) {")
    Send("    var newlocation = 'offer-rew.aspx?OfferID=" & OfferID & "'; ")
    Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
  Send("    opener.location = 'offer-rew.aspx?OfferID=" & OfferID & "'; ")
    Send("    } ")
    Send("  }")
    Send("  }")
  Send("</script>")
%>
<form action="offer-rew-autogiftreceipt.aspx" id="mainform" name="mainform" onsubmit="return saveForm();">
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
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.autogiftreceiptreward", LanguageID), VbStrConv.Lowercase) & "</h1>")
      Else
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.autogiftreceiptreward", LanguageID), VbStrConv.Lowercase) & "</h1>")
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
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <div id="column1">
      <% Send_ProductConditionSelector(Logix, TransactionLevelSelected, bUseTemplateLocks, Disallow_Edit, SelectedItem, ExcludedItem, RewardID, Copient.CommonInc.InstalledEngines.CM, IsTemplate, bDisallowEditPg)%>
    </div>
    <div id="gutter">
    </div>
    <div id="column2">
      <div class="box" id="message">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.message", LanguageID))%>
          </span>
        </h2>
        <%
          sDisabled = " disabled=""disabled"""
          MyCommon.QueryStr = "select OFR.RewardID,Tiered,O.Numtiers,O.OfferID,XT.TierLevel,XT.XmlText from OfferRewards as OFR with (NoLock) " & _
                              "left join Offers as O with (NoLock) on O.OfferID=OFR.OfferID left join RewardXmlTiers as XT with (NoLock) on OFR.RewardID=XT.RewardID where OFR.RewardID=" & RewardID
          rst = MyCommon.LRT_Select()
          q = 1
          For Each row In rst.Rows
            sXmlText = Copient.PhraseLib.Lookup("term.autogiftreceipt", LanguageID)
            If (row.Item("Tiered") = False) Then
              Send("<input class=""longer"" id=""tier0"" name=""tier0"" type=""text"" value=""" & sXmlText & """ maxlength=""14""" & sDisabled & " /><br />")
            Else
              Send("<label for=""tier" & q & """><b>" & "MCLU Tier " & q & ":</b></label><br />")
              Send("<input class=""longer"" id=""tier" & q & """ name=""tier" & q & """ type=""text"" value=""" & sXmlText & """ maxlength=""14""" & sDisabled & " /><br />")
            End If
            q = q + 1
          Next
          Send("<input type=""hidden"" name=""NumTiers"" value=""" & row.Item("NumTiers") & """ />")
        %>
        <br class="half" />
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
    </div>
  </div>
</form>

<script type="text/javascript">
<% If (CloseAfterSave) Then %>
    <% If (infoMessage = "") Then %>
        window.close();
    <% End If %>
<% End If %>

updateButtons();
removeUsed(false);
checkConditionState();

setlimitsection(false);
setperiodsection(false);

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
  Send_BodyEnd()
  Logix = Nothing
  MyCommon = Nothing
%>
