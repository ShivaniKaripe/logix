<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: CM-offer-rew-cents-off-fuel.aspx 
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
  Dim TransactionLevelSelected As Boolean
  Dim DistPeriod As Integer
  Dim RewardLimit As Decimal
  Dim RewardLimitTypeID As Integer
  Dim ChargeBackDeptID As Integer
  Dim ProgramID As Long
  Dim DiscountableItemsOnly As Boolean
  Dim x As Integer
  Dim Tiered As Integer
  Dim SponsorID As Integer
  Dim Disallow_Edit As Boolean = True
  Dim IsTemplate As Boolean = False
  Dim CloseAfterSave As Boolean = False
  Dim RequirePG As Boolean = False
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim disabledattribute As String = ""
  Dim ProductGroupID As Integer = 0
  Dim ExcludedID As Integer = 0
  Dim BannersEnabled As Boolean = False
  Dim OfferEngineID As Long = 0

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
  Dim objTemp As Object
  Dim decTemp As Decimal
  Dim intNumDecimalPlaces As Integer
  Dim decFactor As Decimal
  Dim sTemp As String
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  Response.Expires = 0
  MyCommon.AppName = "CM-offer-rew-cents-off-fuel.aspx"
  
  TransactionLevelSelected = False
  
  OfferID = Request.QueryString("OfferID")
  RewardID = Request.QueryString("RewardID")
  NumTiers = Request.QueryString("NumTiers")
  ProductGroupID = MyCommon.Extract_Val(Request.QueryString("ProductGroupID"))
  ExcludedID = MyCommon.Extract_Val(Request.QueryString("ExcludedID"))
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  
  If (Request.QueryString("save") <> "") Then
    CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
  End If
  
  objTemp = MyCommon.Fetch_CM_SystemOption(41)
  If Not (Integer.TryParse(objTemp.ToString, intNumDecimalPlaces)) Then
    intNumDecimalPlaces = 0
  End If
  decFactor = (10 ^ intNumDecimalPlaces)
  
  ' ok lets find out if were possibly supposed to show the transaction level choices and if we already have one selected
  MyCommon.QueryStr = "select LinkID,RewardAmountTypeID,ProductGroupID,ExcludedProdGroupID from OfferRewards with (NoLock) where RewardID=" & RewardID
  rst = MyCommon.LRT_Select
  If (rst.Rows.Count > 0) Then
    For Each row In rst.Rows
      If (MyCommon.NZ(row.Item("ProductGroupID"), 0) = 0) Then
        TransactionLevelSelected = True
      End If
      If (MyCommon.NZ(row.Item("RewardAmountTypeID"), 0) > 7) Then
        If (MyCommon.NZ(row.Item("ProductGroupID"), 0) = 0) Then
          TransactionLevelSelected = True
        Else
          TransactionLevelSelected = False
        End If
      End If
    Next
  Else
    TransactionLevelSelected = True
  End If
  
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

  Send_HeadBegin("term.offer", "term.centsofffuelreward", OfferID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
%>
<script type="text/javascript">
var fullSelect = null;

function checkConditionState() {
  var elemExcluded=document.getElementById("excluded");
  var elemSelected=document.getElementById("selected");
  var elemNonTransOptions = document.getElementById("nontransoptions");
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
  if (elemNonTransOptions != null)
  {
    if (isTransactionLevel == true)
    {
      elemNonTransOptions.style.display = "none";
    }
    else
    {
      elemNonTransOptions.style.display = "block";
    }
  }
}

 
var prevSelectedVal = -1
var prevIsTransactionLevel = -1 
 

// This is the javascript array holding the function list
// The PrintJavascriptArray ASP function can be used to print this array.
<%
    MyCommon.QueryStr = "select ProductGroupID,Name from ProductGroups with (NoLock) where Deleted=0 order by AnyProduct desc, Name"
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

function removeUsed(bSkipKeyUp)
{
    if (!bSkipKeyUp) handleKeyUp(99999);
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
  return true;
}

function handleClickSent() {
  saveForm();
  document.mainform.clicksent.value='true';
  document.mainform.submit();
}
function saveForm(){
    var funcSel = document.getElementById('functionselect');
    var exSel = document.getElementById('excluded');
    var elSel = document.getElementById('selected');
    var Pselected = document.getElementById('Pselected');
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
    document.getElementById("ProductGroupID").value = selectList;

    for (i = exSel.length - 1; i>=0; i--) {
        if(exSel.options[i].value != ""){
            if(excludededList != "") { excludededList = excludededList + ","; }
            excludededList = excludededList + exSel.options[i].value;
        }
    }
    document.getElementById("ExcludedID").value = excludededList;

    if (Pselected != null && Pselected.options.length > 0) {
        document.getElementById("ProgramID").value = Pselected.options[0].value;
    }
        
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
    MyCommon.QueryStr = "select SVProgramID as ProgramID,Name as ProgramName from StoredValuePrograms with (NoLock) " & _
                        "where Deleted=0 and SVProgramID is not null and SVTypeID=1 " & _
                        "and SVExpireType=5 and SVExpirePeriodType=3 order by ProgramName;"
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
    End If
%>

// This is the function that refreshes the list after a keypress.
// The maximum number to show can be limited to improve performance with
// huge lists (1000s of entries).
// The function clears the list, and then does a linear search through the
// globally defined array and adds the matches back to the list.
function PhandleKeyUp(maxNumToShow) {
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
                        "from CM_AdvancedLimits with (NoLock) where Deleted=0 and LimitTypeID=2 order By LimitID;"
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
  MyCommon.QueryStr = "select LinkID,Tiered,SponsorID,RewardDistPeriod,PromoteToTransLevel,RewardLimit,RewardLimitTypeID,TriggerQty,RewardAmountTypeID, " & _
                      "ApplyToLimit,DoNotItemDistribute,AdvancedLimitID from OfferRewards with (NoLock) where RewardID=" & RewardID
  rst = MyCommon.LRT_Select
  For Each row In rst.Rows
    DistPeriod = MyCommon.NZ(row.Item("RewardDistPeriod"), 0)
    LinkID = row.Item("LinkID")
    Tiered = MyCommon.NZ(row.Item("Tiered"), 0)
    SponsorID = MyCommon.NZ(row.Item("SponsorID"), 0)
    ' ExItemLevelDist = MyCommon.NZ(row.Item("ExItemLevelDist"), 0)
    RewardLimit = MyCommon.NZ(row.Item("RewardLimit"), 0)
    RewardLimitTypeID = MyCommon.NZ(row.Item("RewardLimitTypeID"), 2)
    RewardAmountTypeID = MyCommon.NZ(row.Item("RewardAmountTypeID"), 1)
    AdvancedLimitID = MyCommon.NZ(row.Item("AdvancedLimitID"), 0)
    TriggerQty = MyCommon.NZ(row.Item("TriggerQty"), 1)
    ApplyToLimit = MyCommon.NZ(row.Item("ApplyToLimit"), 1)
    DoNotItemDistribute = row.Item("DoNotItemDistribute")
  Next
  
  MyCommon.QueryStr = "select RewardStoredValuesID,ProgramID,ChargeBackDeptID,DiscountableItemsOnly from CM_RewardStoredValues with (NoLock) where RewardStoredValuesID=" & LinkID
  rst = MyCommon.LRT_Select()
  For Each row In rst.Rows
    ChargeBackDeptID = MyCommon.NZ(row.Item("ChargeBackDeptID"), 0)
    ProgramID = MyCommon.NZ(row.Item("ProgramID"), 0)
    DiscountableItemsOnly = MyCommon.NZ(row.Item("DiscountableItemsOnly"), 0)
  Next
     
  If (Request.QueryString("pgroup-add1") <> "" And Request.QueryString("pgroup-avail") <> "") Then
    MyCommon.QueryStr = "update OfferRewards with (RowLock) set ProductGroupID=0 where RewardID=" & RewardID & ";"
    MyCommon.LRT_Execute()
    MyCommon.QueryStr = "update OfferRewards with (RowLock) set ProductGroupID=" & Request.QueryString("pgroup-avail") & " where RewardID=" & RewardID & ";"
    MyCommon.LRT_Execute()
  ElseIf (Request.QueryString("pgroup-rem1") <> "") Then
    MyCommon.QueryStr = "update OfferRewards with (RowLock) set ProductGroupID=0 where RewardID=" & RewardID & ";"
    MyCommon.LRT_Execute()
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
  ElseIf (Request.QueryString("save") <> "" Or Request.QueryString("clicksent") = "true" Or _
    Request.QueryString("pgroup-add1") <> "" Or _
    Request.QueryString("pgroup-rem1") <> "" Or _
    Request.QueryString("pgroup-add2") <> "" Or _
    Request.QueryString("pgroup-rem2") <> "") Then
    
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
          infoMessage = Copient.PhraseLib.Lookup("reward.groupselect", LanguageID)
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

    If Not (bUseTemplateLocks And bDisallowEditPp) Then
      If (Request.QueryString("ProgramID") <> "") Then
        ProgramID = Request.QueryString("ProgramID")
        MyCommon.QueryStr = "dbo.pa_CM_UpdateRewardStoredValues"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@ProgramID", SqlDbType.BigInt).Value = ProgramID
        MyCommon.LRTsp.Parameters.Add("@WeightUOM", SqlDbType.Bit).Value = 0
        MyCommon.LRTsp.Parameters.Add("@Linkid", SqlDbType.BigInt).Value = LinkID
        MyCommon.LRTsp.ExecuteNonQuery()
        MyCommon.Close_LRTsp()
      Else
        MyCommon.QueryStr = "update CM_RewardStoredValues with (RowLock) set ProgramID=Null where RewardStoredValuesID=" & LinkID
        MyCommon.LRT_Execute()
        infoMessage = Copient.PhraseLib.Lookup("cpe-discount-selectprogram", LanguageID)
      End If
    End If

    If Not (bUseTemplateLocks And bDisallowEditDist) Then
      If (Request.QueryString("Xbox2") <> "") Then
        decTemp = MyCommon.Extract_Val(Request.QueryString("Xbox2")) * decFactor
        TriggerQty = Int(decTemp + 0.5)
        MyCommon.QueryStr = "update OfferRewards with (RowLock) set TriggerQty=" & TriggerQty & _
                            " where RewardID=" & RewardID & ";"
        MyCommon.LRT_Execute()
      End If
      If (Request.QueryString("Ybox2") <> "") Then
        decTemp = MyCommon.Extract_Val(Request.QueryString("Ybox2")) * 100.0
        ApplyToLimit = Int(decTemp + 0.5)
        MyCommon.QueryStr = "update OfferRewards with (RowLock) set ApplyToLimit=" & ApplyToLimit & _
                            " where RewardID=" & RewardID & ";"
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
            RewardLimitTypeID = MyCommon.NZ(rst.Rows(0).Item("LimitTypeID"), 2)
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

    If Not (bUseTemplateLocks And bDisallowEditDept) Then
      If (Request.QueryString("dept") <> "") Then
        ChargeBackDeptID = Request.QueryString("dept")
        MyCommon.QueryStr = "update CM_RewardStoredValues with (RowLock) set ChargeBackDeptID=" & ChargeBackDeptID & " where RewardStoredValuesID=" & LinkID
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
    
    If Not (bUseTemplateLocks And (bDisallowEditSpc And bDisallowEditDist)) Then
      If Not (bUseTemplateLocks And bDisallowEditSpc) Then
        MyCommon.QueryStr = "update OfferRewards with (RowLock) set UseSpecialPricing=0 where RewardID=" & RewardID
        MyCommon.LRT_Execute()
      End If
      If Not (bUseTemplateLocks And bDisallowEditDist) Then
        If (Tiered = 0) Then
          'dbo.pt_ConditionTiers_Update @ConditionID bigint, @TierLevel int,@AmtRequired decimal(12,3)
          MyCommon.QueryStr = "dbo.pt_RewardTiers_Update"
          MyCommon.Open_LRTsp()
          MyCommon.LRTsp.Parameters.Add("@RewardID", SqlDbType.BigInt).Value = RewardID
          MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int).Value = 0
          MyCommon.LRTsp.Parameters.Add("@RewardAmount", SqlDbType.Decimal).Value = MyCommon.Extract_Val(Request.QueryString("tier0"))
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
    
    
    MyCommon.QueryStr = "update OfferRewards with (RowLock) set TCRMAStatusFlag=2,CMOAStatusFlag=2 where RewardID=" & RewardID
    MyCommon.LRT_Execute()
    MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where offerid=" & OfferID
    MyCommon.LRT_Execute()
    
    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.editcentsofffuel", LanguageID))
  End If
  
  If (Request.QueryString("pgroup-add1") <> "" Or Request.QueryString("pgroup-rem1") <> "" Or Request.QueryString("pgroup-add2") <> "" Or Request.QueryString("pgroup-rem2") <> "") Then
    MyCommon.QueryStr = "update OfferRewards with (RowLock) set TCRMAStatusFlag=3,CMOAStatusFlag=2 where RewardID=" & RewardID
    MyCommon.LRT_Execute()
    MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where offerid=" & OfferID
    MyCommon.LRT_Execute()
  End If
  
%>
<script type="text/javascript">
function ChangeParentDocument() {
    var newlocation = 'offer-rew.aspx?OfferID=<% sendb(OfferID) %>';
    if (opener != null) {
        if (opener.location.href.indexOf(newlocation) > -1) {
            opener.location = 'offer-rew.aspx?OfferID=<% sendb(OfferID) %>';
        }
    }
} 
 
function submitenter(e) {
	var key = window.event ? e.keyCode : e.which;
	var keychar = String.fromCharCode(key);
	// if key is not decimal or numeric , then return false
	// 46 = '.', 48 = '0' and 57 = '9'
	return((key > 47 && key < 58) || key == 46);
}
    
</script>


<form action="CM-offer-rew-cents-off-fuel.aspx" id="mainform" name="mainform" onsubmit="return saveForm();">
  <div id="intro">
    <input type="hidden" id="OfferID" name="OfferID" value="<% sendb(OfferID) %>" />
    <input type="hidden" id="Name" name="Name" value="<% sendb(Name) %>" />
    <input type="hidden" id="RewardID" name="RewardID" value="<% sendb(RewardID) %>" />
    <input type="hidden" id="ProductGroupID" name="ProductGroupID" value="<% sendb(ProductGroupID) %>" />
    <input type="hidden" id="ExcludedID" name="ExcludedID" value="<% sendb(ExcludedID) %>" />
    <input type="hidden" id="ProgramID" name="ProgramID" value="" />
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
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.centsofffuelreward", LanguageID), VbStrConv.Lowercase) & "</h1>")
      Else
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.centsofffuelreward", LanguageID), VbStrConv.Lowercase) & "</h1>")
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
      <%
        Send_ProductConditionSelector(Logix, TransactionLevelSelected, bUseTemplateLocks, Disallow_Edit, SelectedItem, ExcludedItem, RewardID, Copient.CommonInc.InstalledEngines.CM, IsTemplate, bDisallowEditPg)
      %>
      <div class="box" id="department">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.department", LanguageID))%>
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
            MyCommon.QueryStr = "Select * from ChargeBackDepts with (NoLock) Order By ExternalID"
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
    <div id="gutter">
    </div>
    <div id="column2">
      <div class="box" id="programs">
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
        <input type="radio" id="Pfunctionradio1" name="Pfunctionradio" checked="checked"
          <% sendb(disabledattribute) %> /><label for="Pfunctionradio1"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
        <input type="radio" id="Pfunctionradio2" name="Pfunctionradio" <% sendb(disabledattribute) %> /><label
          for="Pfunctionradio2"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
        <input class="medium" onkeyup="PhandleKeyUp(200);" id="Pfunctioninput" name="Pfunctioninput"
          type="text" maxlength="100" value="" <% sendb(disabledattribute) %> /><br />
        <select class="longer" id="Pfunctionselect" name="Pfunctionselect" size="10" <% sendb(disabledattribute) %>>
          <%
            MyCommon.QueryStr = "select SVProgramId as ProgramID,Name as ProgramName from StoredValuePrograms with (NoLock) " & _
                                "where Deleted=0 and SVProgramID is not null and SVTypeID=1 " & _
                                "and SVExpireType=5 and SVExpirePeriodType=3 order by ProgramName;"
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
          MyCommon.QueryStr = "select SVP.SVProgramID as ProgramID, SVP.Name as ProgramName from StoredValuePrograms as SVP with (NoLock) " & _
                    "inner join CM_RewardStoredValues as RSV with (NoLock) on RSV.ProgramID=SVP.SVProgramID " & _
                    "inner join OfferRewards as OFR with (NoLock) on OFR.LinkID=RSV.RewardStoredValuesID " & _
                    "where RewardID=" & RewardID & " and SVP.Deleted=0 and OFR.Deleted=0;"
          rst2 = MyCommon.LRT_Select
          Send("<label for=""Pselected""><b>" & Copient.PhraseLib.Lookup("term.selectedprogram", LanguageID) & "</b></label><br />")
          Send("<input class=""regular"" id=""Pselect1"" name=""Pselect1"" type=""button"" value=""&#9660; " & Copient.PhraseLib.Lookup("term.select", LanguageID) & """ onclick=""PhandleSelectClick('Pselect1');""" & IIf(rst2.Rows.Count > 0, " disabled=""disabled""", "") & " />&nbsp;")
          Send("<input class=""regular"" id=""Pdeselect1"" name=""Pdeselect1"" type=""button"" value=""" & Copient.PhraseLib.Lookup("term.deselect", LanguageID) & " &#9650;"" onclick=""PhandleSelectClick('Pdeselect1');""" & IIf(rst2.Rows.Count = 0, " disabled=""disabled""", "") & " /><br />")
          Send("<br class=""half"" />")
          Send("<select class=""longer"" id=""Pselected"" name=""Pselected"" size=""2""" & disabledattribute & ">")
          For Each row2 In rst2.Rows
            Send("<option value=""" & row2.Item("ProgramID") & """>" & row2.Item("ProgramName") & "</option>")
          Next
          Send("</select>")
        %>
        <hr class="hidden" />
      </div>
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
        <%
          If (bUseTemplateLocks And bDisallowEditDist) Then
            sDisabled = " disabled=""disabled"""
          Else
            sDisabled = ""
          End If

          MyCommon.QueryStr = "select ORw.RewardID,ORw.Tiered,O.Numtiers,O.OfferID,RT.TierLevel,RT.RewardAmount from OfferRewards as ORw with (NoLock) " & _
                              "inner join Offers as O with (NoLock) on O.OfferID=ORw.OfferID " & _
                              "inner join RewardTiers as RT with (NoLock) on RT.RewardID=ORw.RewardID " & _
                              "where ORw.RewardID=" & RewardID
          rst = MyCommon.LRT_Select()
          Send("</select>")
          Send("</td>")
          Send("</tr>")
          Send("</table>")
        %>
        <div id="nontransoptions" style="display:<% Sendb(IIF(TransactionLevelSelected, "none", "block"))%>; position:relative;">
          <%
            decTemp = (ApplyToLimit * 1.0) / 100.0
            sTemp = decTemp.ToString("0.00")
            Sendb(Copient.PhraseLib.Lookup("term.get", LanguageID) & " $")
          %>
          &nbsp;
          <input class="shorter" id="bxgy2" name="Ybox2" maxlength="9" type="text" onkeypress="return submitenter(event)" <% Send(sDisabled) %><%sendb(" value=""" & sTemp & """ ") %> />
          &nbsp;
          <% Sendb(Copient.PhraseLib.Lookup("term.OffPerGallon", LanguageID))%>
          <br /><br />
          <% 
            If intNumDecimalPlaces > 0 Then
              decTemp = (TriggerQty * 1.0) / decFactor
              sTemp = FormatNumber(decTemp, intNumDecimalPlaces)
            Else
              sTemp = TriggerQty.ToString()
            End If
            Sendb(Copient.PhraseLib.Lookup("term.foreach", LanguageID) & " ")
          %>
          &nbsp;
          <input class="shorter" id="bxgy1" name="Xbox2" maxlength="9" type="text" onkeypress="return submitenter(event)" <% Send(sDisabled) %><%send(" value=""" & sTemp & """ ") %> />
          &nbsp;
          <% Sendb(Copient.PhraseLib.Lookup("term.PointsEarned", LanguageID))%>
        </div>
        &nbsp;<br />
        <hr class="hidden" />
      </div>
      <div class="box" id="limits" style="position:relative;">
        <h2 style="position:relative;">
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
          MyCommon.QueryStr = "Select LimitId, Name from CM_AdvancedLimits with (NoLock) where Deleted=0 and LimitTypeID=2 Order By Name;"
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
          value="<% sendb(RewardLimit) %>" <% If(bUseTemplateLocks and bDisallowEditLimit) Then sendb(" disabled=""disabled""") %> />
        <% Sendb(Copient.PhraseLib.Lookup("term.amount", LanguageID))%>
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
<% If ((CloseAfterSave) and (infoMessage = "")) Then %>
    window.close();
<% End If %>
  updateButtons();
  removeUsed(true);
  PupdateButtons();
  PremoveUsed();
  checkConditionState();
  
  PhandleKeyUp(99999);

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
  Send_BodyEnd()
  Logix = Nothing
  MyCommon = Nothing
%>
