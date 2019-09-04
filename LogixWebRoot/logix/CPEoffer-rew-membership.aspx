﻿<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
    ' *****************************************************************************
    ' * FILENAME: CPEoffer-rew-membership.aspx 
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
    Dim rstTiers As DataTable
    Dim OfferID As Long
    Dim DeliverableID As Long
    Dim Phase As Integer
    Dim Name As String = ""
    Dim RewardID As String
    Dim RewardTypeID As Integer = 5
    Dim GrantSelected As String = ""
    Dim RemoveSelected As String = ""
    Dim bIsErrorMsg As Boolean = False
    Dim TouchPoint As Integer = 0
    Dim TpROID As Integer = 0
    Dim CreateROID As Integer = 0
    Dim CloseAfterSave As Boolean = False
    Dim Disallow_Edit As Boolean = True
    Dim FromTemplate As Boolean = False
    Dim IsTemplate As Boolean = False
    Dim IsTemplateVal As String = "Not"
    Dim DisabledAttribute As String = ""
    Const GRANT_MEMBERSHIP As Integer = 5
    Const REMOVE_MEMBERSHIP As Integer = 6
    Dim Action As Integer = GRANT_MEMBERSHIP
    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False
    Dim EngineID As Integer = 2
    Dim BannersEnabled As Boolean = True
    Dim TierLevels As Integer = 1
    Dim t As Integer = 1
    Dim ValidTiers As Boolean = True
    Dim RECORD_LIMIT As Integer = GroupRecordLimit '500

    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If

    Response.Expires = 0
    MyCommon.AppName = "CPEoffer-rew-membership.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)

    BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")

    OfferID = Request.QueryString("OfferID")
    RewardID = Request.QueryString("RewardID")
    DeliverableID = MyCommon.Extract_Val(Request.QueryString("DeliverableID"))
    Phase = MyCommon.Extract_Val(Request.QueryString("phase"))
    If (Phase = 0) Then Phase = MyCommon.Extract_Val(Request.Form("Phase"))
    If (Phase = 0) Then Phase = 3
    If (Request.QueryString("EngineID") <> "") Then
        EngineID = MyCommon.Extract_Val(Request.QueryString("EngineID"))
    Else
        MyCommon.QueryStr = "select EngineID from OfferIDs where OfferID=" & OfferID & ";"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
            EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), 0)
        End If
    End If

    MyCommon.QueryStr = "select TierLevels from CPE_RewardOptions where RewardOptionID=" & RewardID & ";"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
        TierLevels = MyCommon.NZ(rst.Rows(0).Item("TierLevels"), 1)
    End If

    Action = MyCommon.Extract_Val(Request.QueryString("action"))
    TouchPoint = MyCommon.Extract_Val(Request.QueryString("tp"))
    If (TouchPoint > 0) Then TpROID = MyCommon.Extract_Val(Request.QueryString("roid"))

    GrantSelected = IIf(Action = GRANT_MEMBERSHIP, " selected=""selected""", "")
    RemoveSelected = IIf(Action = REMOVE_MEMBERSHIP, " selected=""selected""", "")

    ' Fetch the name
    MyCommon.QueryStr = "Select IncentiveName, IsTemplate, FromTemplate from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
        Name = MyCommon.NZ(rst.Rows(0).Item("IncentiveName"), "")
        IsTemplate = MyCommon.NZ(rst.Rows(0).Item("IsTemplate"), False)
        FromTemplate = MyCommon.NZ(rst.Rows(0).Item("FromTemplate"), False)
    End If
    IsTemplateVal = IIf(IsTemplate, "IsTemplate", "Not")

    ' Get the per-tier customer group IDs
    MyCommon.QueryStr = "select TierLevel, CustomerGroupID from CPE_DeliverableCustomerGroupTiers with (NoLock) where DeliverableID=" & DeliverableID & ";"
    rstTiers = MyCommon.LRT_Select

    ' Save logic
    If (Request.QueryString("save") <> "") Then
        Action = MyCommon.Extract_Val(Request.QueryString("action"))
        If (Action > 0) Then
            CreateROID = IIf(TpROID > 0, TpROID, RewardID)
            ' Delete existing tier records
            MyCommon.QueryStr = "delete from CPE_DeliverableCustomerGroupTiers with (RowLock) where DeliverableID in (0, " & DeliverableID & ");"
            MyCommon.LRT_Execute()
            ' Add a new record to CPE_Deliverables
            MyCommon.QueryStr = "dbo.pa_CPE_AddMembership"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.Int, 4).Value = OfferID
            MyCommon.LRTsp.Parameters.Add("@ROID", SqlDbType.Int, 4).Value = CreateROID
            MyCommon.LRTsp.Parameters.Add("@DeliverableTypeID", SqlDbType.Int, 4).Value = Action
            MyCommon.LRTsp.Parameters.Add("@Phase", SqlDbType.Int, 4).Value = Phase
            MyCommon.LRTsp.Parameters.Add("@OutputID", SqlDbType.Int, 4).Value = MyCommon.Extract_Val(Request.QueryString("t" & t & "_CustomerGroupID"))
            MyCommon.LRTsp.Parameters.Add("@DeliverableID", SqlDbType.Int, 4).Direction = ParameterDirection.Output
            MyCommon.LRTsp.ExecuteNonQuery()
            DeliverableID = MyCommon.LRTsp.Parameters("@DeliverableID").Value
            MyCommon.Close_LRTsp()
            ' Add the tier records to CPE_DeliverableCustomerGroupTiers
            For t = 1 To TierLevels
                MyCommon.QueryStr = "dbo.pa_CPE_AddMembershipTiers"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@DeliverableID", SqlDbType.Int, 4).Value = DeliverableID
                MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int, 4).Value = t
                MyCommon.LRTsp.Parameters.Add("@CustomerGroupID", SqlDbType.Int, 4).Value = MyCommon.Extract_Val(Request.QueryString("t" & t & "_CustomerGroupID"))
                MyCommon.LRTsp.ExecuteNonQuery()
                MyCommon.Close_LRTsp()
            Next
            MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
            MyCommon.LRT_Execute()
            infoMessage = IIf(Action = GRANT_MEMBERSHIP, Copient.PhraseLib.Lookup("term.grant", LanguageID), Copient.PhraseLib.Lookup("term.remove", LanguageID))
            infoMessage += Copient.PhraseLib.Lookup("CPE-rew-membership.edit", LanguageID) & OfferID
            MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Reward.createmembership", LanguageID))
        End If
        CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
    End If

    'update the templates permission if necessary
    If (Request.QueryString("save") <> "" AndAlso Request.QueryString("IsTemplate") = "IsTemplate") Then
        ' time to update the status bits for the templates
        Dim form_Disallow_Edit As Integer = 0
        If (Request.QueryString("Disallow_Edit") = "on") Then
            form_Disallow_Edit = 1
        End If
        MyCommon.QueryStr = "update CPE_Deliverables with (RowLock) set DisallowEdit=" & form_Disallow_Edit & _
                            " where DeliverableID=" & DeliverableID
        MyCommon.LRT_Execute()
    End If

    If (IsTemplate Or FromTemplate) Then
        ' lets dig the permissions if its a template
        MyCommon.QueryStr = "select DisallowEdit from CPE_Deliverables with (NoLock) where DeliverableID=" & DeliverableID
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
            Disallow_Edit = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit"), False)
        Else
            Disallow_Edit = False
        End If
    End If

    If Not isTemplate Then
        DisabledAttribute = IIf(Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Edit), "", " disabled=""disabled""")
    Else
        DisabledAttribute = IIf(Logix.UserRoles.EditTemplates, "", " disabled=""disabled""")
    End If

    Send_HeadBegin("term.offer", "term.membershipreward", OfferID)
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts()
%>
<script type="text/javascript">
// This is the javascript array holding the function list
// The PrintJavascriptArray ASP function can be used to print this array.
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
  
  //document.getElementById("functionselect").style.display = "none";
  document.getElementById("searchLoadDiv").innerHTML = '<% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID)) %>';
  
  // Mozilla/Safari
  if (window.XMLHttpRequest) {
    self.xmlHttpReq = new XMLHttpRequest();
  }
  // IE
  else if (window.ActiveXObject) {
    self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
  }
  var qryStr = getgroupquery(mode);
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

function getgroupquery(mode) {
  var radioString;
  if(document.getElementById('functionradio2').checked) {
    radioString = 'functionradio2';
  }
  else {
    radioString = 'functionradio1';
  }

return "Mode=" + mode + "&OfferID=" + document.getElementById('OfferID').value + "&RewardID=" + document.getElementById('RewardID').value + "&Search=" + document.getElementById('functioninput').value + "&EngineID=" + <% Sendb(EngineID) %> + "&SearchRadio=" + radioString + "" + GetQueryString();

}

function updatepage(str) {
  if(str.length > 0){
    if(!isFireFox){
      document.getElementById("cgList").innerHTML = '<select class="longer" id="functionselect" name="functionselect" onkeydown="return handleSlctKeyDown(event);" onclick="handleSelectClick();" ondblclick="addGroupToSelect();" size="12"<% Sendb(DisabledAttribute) %>>' + str + '</select>';
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
      document.getElementById("cgList").innerHTML = '<select class="longer" id="functionselect" name="functionselect" onkeydown="return handleSlctKeyDown(event);" onclick="handleSelectClick();" ondblclick="addGroupToSelect();" size="12"<% Sendb(DisabledAttribute) %>></select>';
    }
    else{
      document.getElementById("functionselect").innerHTML = '';
      }

      document.getElementById("searchLoadDiv").innerHTML = '&nbsp;';
    }
  }

  function GetQueryString() {
    var iTierLevels = <% Sendb(TierLevels) %>;
    var selectedGroups = "";
    var qString = "";
    var SelectCount = 0;
    for (i=1; i <= iTierLevels; i++) {
        var selElem = document.getElementById("t"+ i +"_selected");
        if (i != iTierLevels ) {
            if (selElem.options.length > 0) {
                SelectCount++;
                qString = qString + "Group" + SelectCount + "=" + selElem.options[0].value + "&";
            }
        }
        else {
            if (selElem.options.length > 0) {
                SelectCount++
                qString = qString + "Group" + SelectCount + "=" + selElem.options[0].value;
            }
        }
    }
    if (qString != "") {
        qString = "&GroupCount=" + SelectCount + "&" + qString;
    }
    return qString;
  }

// This is the function that refreshes the list after a keypress.
// The maximum number to show can be limited to improve performance with
// huge lists (1000s of entries).
// The function clears the list, and then does a linear search through the
// globally defined array and adds the matches back to the list.
function handleKeyUp(maxNumToShow, TierLevels) {
  var selectObj, textObj, functionListLength;
  var i;
  var t;
  var numShown;
  var searchPattern;
  var selectedList;
  
  document.getElementById("functionselect").size = "12";
  
  // Set references to the form elements
  selectObj = document.forms[0].functionselect;
  textObj = document.forms[0].functioninput;
  selectedList = document.getElementById("t1_selected");
  
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
  
  // Clear the options list
  selectObj.length = 0;
  
  // Loop through the array and re-add matching options
  numShown = 0;
  for(i = 0; i < functionListLength; i++) {
    if(functionlist[i].search(re) != -1) {
//      if (vallist[i] != "" && (selectedList.options.length < 1 || vallist[i] != selectedList.options[0].value)) {      if (vallist[i] != "" && (selectedList.options.length < 1 || vallist[i] != selectedList.options[0].value)) {
      if (vallist[i] != "" ) {
        selectObj[numShown] = new Option(functionlist[i],vallist[i]);
        if (vallist[i] == 2) {
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
  // When options list whittled to one, select that entry
  if(selectObj.length == 1) {
    selectObj.options[0].selected = true;
  }
}

// this function gets the selected value and loads the appropriate
// php reference page in the display frame
// it can be modified to perform whatever action is needed, or nothing
function handleSelectClick() {
  selectObj = document.forms[0].functionselect;
  textObj = document.forms[0].functioninput;
  
  selectedValue = document.getElementById("functionselect").value;
  
  if(selectedValue != "") {
  }
}

function addGroupToSelect(t) {
  var elemSource = document.getElementById("functionselect");
  var elemDest = document.getElementById("t" + t + "_selected");   
  var elemDeselect = document.getElementById("t" + t + "_deselect");  
  var selOption = null;
  var selText ="", selVal = "";
  var selIndex = -1;
  
  if (elemSource != null && elemSource.options.selectedIndex == -1) {
    alert('<% Sendb(Copient.PhraseLib.Lookup("CPE-rew-membership.selectgroup", LanguageID)) %>');
    elemSource.focus();
  } else {
    selIndex = elemSource.options.selectedIndex;
    selOption = elemSource.options[selIndex];
    selText = selOption.text;
    selVal = selOption.value;

    elemDeselect.disabled=false;

    if (elemDest != null && elemDest.options.length > 0) {
      removeGroupFromSelect(t);
    }
    elemDest.options[0] = new Option(selText, selVal);
    elemDest.options[0].style.fontweight = 'bold';
    if (selVal ==2 ) {
      elemDest.options[0].style.color = 'brown';
      elemSource.options[selIndex] = null;
    }
    //handleKeyUp(99999); 
    removeUsed(t);
  }
}

function removeGroupFromSelect(t) {
  var elem = document.getElementById("t" + t + "_selected");   
  var elemList = document.getElementById("functionselect");
  
  if (elem != null && elem.options.length > 0) {
    elemList.options[elemList.options.length] = new Option(elem.options[0].text, elem.options[0].value);
    elem.options[0] = null;
    //handleKeyUp(99999);
  }
  removeUsed(t);
}

function validateEntry(tierLevels) {
  var retVal = true;
  var i;
  var elem = document.getElementById("t1_selected");
  var elemGroup = document.getElementById("t1_CustomerGroupID");
  
  // Loop through the tiers
  for(i = 1; i <= tierLevels; i++) {
    elem = document.getElementById("t" + i + "_selected");
    elemGroup = document.getElementById("t" + i + "_CustomerGroupID");
    if (elem != null && elemGroup != null) {
      if (elem.options.length == 0) {
        retVal = false;
        alert('<% Sendb(Copient.PhraseLib.Lookup("CPE-rew-membership.selectgroup", LanguageID)) %>');
        elem.focus();
      } else {
        elemGroup.value = elem.options[0].value;
      }
    }
  }
  return retVal;
}

// this function will remove items from the functionselect box that are used in 
// selected and excluded boxes
//function removeUsed() {
//  handleKeyUp(99999);
//  
//  var funcSel = document.getElementById('functionselect');
//  var exSel = document.getElementById('selected');
//  var i;
//  
//  for(j=funcSel.length-1;j>=0; j--) {
//    if(funcSel.options[j].value == elSel.options[i].value){
//      funcSel.options[j] = null;
//    }
//  }
//}

  function removeUsed(t) {

  var iTierLevels = <% Sendb(TierLevels) %>;

   for (i=1; i <= iTierLevels; i++) {
        var elem = document.getElementById("t" + i + "_selected");
         var elemDeselect = document.getElementById("t" + i + "_deselect");  
        if (elem != null && elem.options.length > 0) {
        
            elemDeselect.disabled=false;
        }
        else
        {
            elemDeselect.disabled=true;
        }
    }
      xmlPostTimer('OfferFeeds.aspx', 'GrantMembership');
  }

function selectByGroupID(groupID) {
  var elemSelect = document.getElementById("functionselect");
  var selectLength = elemSelect.length;
  
  for(i = 0; i < selectLength; i++) {
    if (elemSelect != null && elemSelect.options[i] != null && elemSelect.options[i].value == groupID) {
      elemSelect.options[i].selected = true;
      addGroupToSelect();
    }
  }
}

function handleKeyDown(e) {
  var key = e.which ? e.which : e.keyCode;
  
  if (key == 40) {
    var elemSlct = document.getElementById("functionselect");
    if (elemSlct != null) { elemSlct.focus(); }
  }
}

function handleSlctKeyDown(e) {
  var key = e.which ? e.which : e.keyCode;
  
  if (key == 13) {
    var elemSlct = document.getElementById("functionselect");
    if (elemSlct != null && elemSlct.disabled == false) { 
      addGroupToSelect();
      clearEntry();
    }
    e.returnValue=false;
    return false;
  }
}

function clearEntry() {
  var elemInput = document.getElementById("functioninput");
  
  if (elemInput != null) {
    elemInput.value = "";
    //handleKeyUp(200);
    elemInput.focus();
  }
}
</script>
<%
  Send("<script type=""text/javascript"">")
  Send("function ChangeParentDocument() { ")
  If (EngineID = 3) Then
    Send("  opener.location = '/logix/web-offer-rew.aspx?OfferID=" & OfferID & "'; ")
  ElseIf (EngineID = 5) Then
    Send("  opener.location = '/logix/email-offer-rew.aspx?OfferID=" & OfferID & "'; ")
  ElseIf (EngineID = 6) Then
    Send("  opener.location = '/logix/CAM/CAM-offer-rew.aspx?OfferID=" & OfferID & "'; ")
  ElseIf (EngineID = 9) Then
    Send("  opener.location = 'UEoffer-rew.aspx?OfferID=" & OfferID & "'; ")
  Else
    Send("  opener.location = '/logix/CPEoffer-rew.aspx?OfferID=" & OfferID & "'; ")
  End If
  Send("} ")
  Sendb("</")
  Send("script>")
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
  
  MyCommon.QueryStr = "select DCGT.TierLevel, DCGT.CustomerGroupID, CG.Name, CG.PhraseID from CPE_DeliverableCustomerGroupTiers as DCGT with (NoLock) " & _
                      "left join CustomerGroups as CG with (NoLock) on CG.CustomerGroupID=DCGT.CustomerGroupID " & _
                      "where DeliverableID=" & DeliverableID & ";"
  rstTiers = MyCommon.LRT_Select
%>
<form action="CPEoffer-rew-membership.aspx" id="mainform" name="mainform" onsubmit="return validateEntry(<% Sendb(TierLevels) %>);">
  <div id="intro">
    <input type="hidden" id="OfferID" name="OfferID" value="<% Sendb(OfferID)%>" />
    <input type="hidden" id="Name" name="Name" value="<% Sendb(Name)%>" />
    <input type="hidden" id="RewardID" name="RewardID" value="<% Sendb(RewardID)%>" />
    <input type="hidden" id="DeliverableID" name="DeliverableID" value="<% Sendb(DeliverableID) %>" />
    <input type="hidden" id="Phase" name="Phase" value="<% Sendb(Phase) %>" />
    <%
      For t = 1 To TierLevels
        Sendb("<input type=""hidden"" id=""t" & t & "_CustomerGroupID"" name=""t" & t & "_CustomerGroupID"" ")
        If rstTiers.Rows.Count > 0 Then
          If t <= rstTiers.Rows.Count Then
            Send("value=""" & MyCommon.NZ(rstTiers.Rows(t - 1).Item("CustomerGroupID"), 0) & """ />")
          Else
            Send("value="""" />")
          End If
        Else
          Send("value="""" />")
        End If
      Next
    %>
    <input type="hidden" id="roid" name="roid" value="<%Sendb(TpROID) %>" />
    <input type="hidden" id="tp" name="tp" value="<%Sendb(TouchPoint) %>" />
    <input type="hidden" id="action" name="action" value="<%Sendb(action) %>" />
    <input type="hidden" id="IsTemplate" name="IsTemplate" value="<% Sendb(IsTemplateVal)%>" />
    <input type="hidden" id="EngineID" name="EngineID" value="<%Sendb(EngineID) %>" />
    <%
      If (IsTemplate) Then
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.membershipreward", LanguageID), VbStrConv.Lowercase) & "</h1>")
      Else
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.membershipreward", LanguageID), VbStrConv.Lowercase) & "</h1>")
      End If
    %>
    <div id="controls">
      <% If (IsTemplate) Then%>
      <span class="temp">
        <input type="checkbox" class="tempcheck" id="Disallow_Edit" name="Disallow_Edit"<% if(Disallow_Edit)then Sendb(" checked=""checked""") %> />
        <label for="Disallow_Edit">
          <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
        </label>
      </span>
      <% End If%>
      <%
        If Not Istemplate Then
          If (Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Edit)) Then
            If DeliverableID = 0 Then
              Send_Save(" onclick=""this.style.visibility='hidden';""")
            Else
              Send_Save()
            End If
          End If
        Else
          If (Logix.UserRoles.EditTemplates) Then
            If DeliverableID = 0 Then
              Send_Save(" onclick=""this.style.visibility='hidden';""")
            Else
              Send_Save()
            End If
          End If
        End If
      %>
    </div>
  </div>
  <div id="main">
    <%
      If (infoMessage <> "" And bIsErrorMsg) Then
        Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
      ElseIf (infoMessage <> "") Then
        Send("<div id=""infobar"" class=""green-background"">" & infoMessage & "</div>")
      End If
    %>
    <div id="column1">
      <input type="hidden" name="action" value="<% Sendb(Action)%>" />
      <!-- Disabled Remove Membership to comply with current Logix 3.9.2 functionality
      <div class="box" id="distribution">
        <h2><span><% 'Sendb(Copient.PhraseLib.Lookup("term.action", LanguageID))%></span></h2>
        <select id="action" name="action">
          <option value="<% Sendb(GRANT_MEMBERSHIP) %>"<% Sendb(GrantSelected) %>><% Sendb(Copient.PhraseLib.Lookup("term.grantmembership", LanguageID))%></option>
          <option value="<% Sendb(REMOVE_MEMBERSHIP) %>"<% Sendb(RemoveSelected) %>><% Sendb(Copient.PhraseLib.Lookup("term.removemembership", LanguageID))%></option>
        </select>
      </div>
      -->
      <div class="box" id="selector">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.customergroup", LanguageID))%>
          </span>
        </h2>
        <input type="radio" id="functionradio1" name="functionradio" <% if(MyCommon.Fetch_SystemOption(175)= "1" ) then sendb(" checked=""checked""") %> <% Sendb(DisabledAttribute) %> /><label for="functionradio1"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
        <input type="radio" id="functionradio2" name="functionradio" <% if(MyCommon.Fetch_SystemOption(175)= "2" ) then sendb(" checked=""checked""") %> /><label for="functionradio2"<% Sendb(DisabledAttribute) %>><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
        <input type="text" class="medium" id="functioninput" name="functioninput" onkeyup="javascript:xmlPostTimer('OfferFeeds.aspx','GrantMembership');" 
            maxlength="100" value=""<% Sendb(DisabledAttribute) %> /><br />
        <div id="searchLoadDiv" style="display:block;">
            &nbsp;</div>
        <div id="cgList">
           <select class="longer" id="Select1" name="functionselect" onkeydown="return handleSlctKeyDown(event);" onclick="handleSelectClick();" ondblclick="addGroupToSelect();" size="12"<% Sendb(DisabledAttribute)%>>
           <%
               Dim topString As String = ""
               If (RECORD_LIMIT > 0) Then topString = " top " & RECORD_LIMIT
               MyCommon.QueryStr = "select" & topString & " CG.CustomerGroupID, CG.Name From CustomerGroups as CG " &
                   "Left Outer Join (Select CGI.CustomerGroupID From CPE_IncentiveCustomerGroups as CGI " &
                   "Inner Join CPE_RewardOptions as RO on CGI.RewardOptionID = RO.RewardOptionID " &
                   "Where RO.IncentiveID = " & OfferID & " and CGI.ExcludedUsers = 0 and CGI.Deleted = 0) as EX on EX.CustomerGroupID = CG.CustomerGroupID " &
                   "Left join ExtSegmentMap exs on exs.InternalId = cg.CustomerGroupID " &
                    "Where EX.CustomerGroupID is null and CG.AnyCardholder <> 1 and CG.AnyCustomer <> 1 and CG.NewCardholders <> 1 " &
                    "and (exs.ExtSegmentID is null or exs.ExtSegmentID > 0) and (exs.SegmentTypeID is null or exs.SegmentTypeID = 1) " &
                    "and CG.Deleted = 0 and CG.CustomerGroupID <> 1 "

               If EngineID = 6 Then
                   MyCommon.QueryStr &= "and CG.CAMCustomerGroup=1 "
               Else
                   MyCommon.QueryStr &= "and CG.CAMCustomerGroup=0 "
               End If

               MyCommon.QueryStr &= "order by CG.CustomerGroupID desc, CG.Name;"
               rst = MyCommon.LRT_Select
               For Each row In rst.Rows
                   Send("<option value=""" & MyCommon.NZ(row.Item("CustomerGroupID"), 0) & """>" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
               Next
            %>
          </select>
        </div>
        <%If (RECORD_LIMIT > 0) Then
            Send(Copient.PhraseLib.Lookup("groups.display", LanguageID) & ": " & RECORD_LIMIT.ToString() & "<br />")
          End If
        %>
        
        <br class="half" />
        <%
          For t = 1 To TierLevels
            If TierLevels > 1 Then
              Send("<br />")
              Send("<label for=""t" & t & "_selected""><b>" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</b></label>&nbsp;")
            End If
            Send("<input type=""button"" class=""regular select"" id=""t" & t & "_select"" name=""t" & t & "_select"" value=""&#9660; " & Copient.PhraseLib.Lookup("term.select", LanguageID) & """ onclick=""addGroupToSelect(" & t & ");""" & DisabledAttribute & " />&nbsp;")
            Send("<input type=""button"" class=""regular select"" id=""t" & t & "_deselect"" name=""t" & t & "_deselect"" value=""" & Copient.PhraseLib.Lookup("term.deselect", LanguageID) & " &#9650;"" onclick=""removeGroupFromSelect(" & t & ");""" & DisabledAttribute & " /><br />")
            Send("<select class=""longer"" id=""t" & t & "_selected"" name=""t" & t & "_selected"" ondblclick=""removeGroupFromSelect(" & t & ");"" size=""2""" & DisabledAttribute & ">")
            If rstTiers.Rows.Count > 0 Then
              If t <= rstTiers.Rows.Count Then
                Send("  <option value=""" & MyCommon.NZ(rstTiers.Rows(t - 1).Item("CustomerGroupID"), 0) & """>" & MyCommon.NZ(rstTiers.Rows(t - 1).Item("Name"), "&nbsp;") & "</option>")
              End If
            End If
            Send("</select>")
            Send("<br />")
          Next
        %>
        <hr class="hidden" />
      </div>
    </div>
    
    <div id="gutter">
    </div>
    
    <div id="column2">
    </div>
  </div>
</form>

<script type="text/javascript">
<% If (CloseAfterSave) Then %>
  window.close();
<% End If %>
removeUsed();
//handleKeyUp(9999, 4);
</script>
<%
  'If (GroupID > 0) Then
  '  Send("<script type=""text/javascript"">")
  '  Send("  selectByGroupID(" & GroupID & ");")
  '  Send("</script>")
  'End If
%>
<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd("mainform", "functioninput")
  Logix = Nothing
  MyCommon = Nothing
%>
