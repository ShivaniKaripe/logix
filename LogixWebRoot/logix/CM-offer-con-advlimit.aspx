<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: CM-offer-con-advlimit.aspx 
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
  Dim OfferID As Long
  Dim Name As String = ""
  Dim ConditionID As String
  Dim Tiered As Boolean
  Dim NumTiers As Integer
  Dim Disallow_Edit As Boolean = True
  Dim bUseTemplateLocks As Boolean
  Dim IsTemplate As Boolean = False
  Dim CloseAfterSave As Boolean = False
  Dim DisabledAttribute As String = ""
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim RequirePP As Boolean = False
  Dim LimitID As Integer = 0
  Dim BannersEnabled As Boolean = False
  Dim bDisallowEditValue as Boolean = False
  Dim bDisallowEditRewards as Boolean = False
  Dim bDisallowEditPp As Boolean = False
  Dim iRadioValue As Int16
  Dim sDisabled As String
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  
  Response.Expires = 0
  MyCommon.AppName = "CM-offer-con-advlimit.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  Tiered = False
  OfferID = Request.QueryString("OfferID")
  ConditionID = Request.QueryString("ConditionID")
  NumTiers = Request.QueryString("NumTiers")
  LimitID = MyCommon.Extract_Val(Request.QueryString("LimitID"))
  
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  
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
    ' lets dig the permissions if its a template
    MyCommon.QueryStr = "select Tiered,Disallow_Edit,RequiredFromTemplate,DisallowEdit1,DisallowEdit2,DisallowEdit3 from OfferConditions with (NoLock) where ConditionID=" & ConditionID & ";"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      Disallow_Edit = MyCommon.NZ(rst.Rows(0).Item("Disallow_Edit"), True)
      RequirePP = MyCommon.NZ(rst.Rows(0).Item("RequiredFromTemplate"), False)
      bDisallowEditPp = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit1"), False)
      bDisallowEditValue = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit2"), False)
      bDisallowEditRewards = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit3"), False)
      Tiered = MyCommon.NZ(rst.Rows(0).Item("Tiered"), False)
      If Tiered Then
        bDisallowEditRewards = True
      End If
      If bUseTemplateLocks Then
        If Disallow_Edit Then
          bDisallowEditPp = True
          bDisallowEditValue = True
          bDisallowEditRewards = True
        Else
          Disallow_Edit = bDisallowEditPp And bDisallowEditValue And bDisallowEditRewards
        End If
      End If
    End If
  End If
  
  Send_HeadBegin("term.offer", "term.advlimitcondition", OfferID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
%>
<script type="text/javascript" language="javascript">
// This is the javascript array holding the function list
// The PrintJavascriptArray ASP function can be used to print this array.

<%
  MyCommon.QueryStr = "Select LimitID, Name, LimitTypeID from CM_AdvancedLimits with (NoLock) where deleted=0 and LimitID is not null order by Name"
  rst = MyCommon.LRT_Select
  If (rst.rows.count>0)
    Sendb("var functionlist = Array(")
    For Each row In rst.Rows
        Sendb("""" & MyCommon.NZ(row.item("Name"), "").ToString().Replace("""", "\""") & """,")
    Next
    Sendb(""""");")
    Sendb("var vallist = Array(")
    For Each row In rst.Rows
        Sendb("""" & MyCommon.NZ(row.item("LimitID"), -1) & """,")
    Next
    Sendb(""""");")
  Else
    Sendb("var functionlist = Array(")
    Send("""" & "" & """);")
    Sendb("var vallist = Array(")
    Send("""" & "" & """);")
  End If
%>

// This is the function that refreshes the list after a keypress.
// The maximum number to show can be limited to improve performance with
// huge lists (1000s of entries).
// The function clears the list, and then does a linear search through the
// globally defined array and adds the matches back to the list.
function handleKeyUp(maxNumToShow) {
    var selectObj, textObj, functionListLength;
    var i,  numShown;
    var searchPattern;
    var selectedList;
    
    document.getElementById("functionselect").size = "16";
    
    // Set references to the form elements
    selectObj = document.forms[0].functionselect;
    textObj = document.forms[0].functioninput;
    selectedList = document.getElementById("selected");
    
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
            if (vallist[i] != "" && (selectedList.options.length < 1 || vallist[i] != selectedList.options[0].value)) {
                selectObj[numShown] = new Option(functionlist[i],vallist[i]);
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

function removeUsed() {
  handleKeyUp(99999);
  // this function will remove items from the functionselect box that are used in 
  // selected and excluded boxes

  var funcSel = document.getElementById('functionselect');
  var elSel = document.getElementById('selected');
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
function handleSelectClick(itemSelected) {
    textObj = document.forms[0].functioninput;
     
    selectObj = document.forms[0].functionselect;
    selectedValue = document.getElementById("functionselect").value;
    if(selectedValue != ""){ selectedText = selectObj[document.getElementById("functionselect").selectedIndex].text; }
    
    selectboxObj = document.forms[0].selected;
    selectedboxValue = document.getElementById("selected").value;
    if(selectedboxValue != ""){ selectedboxText = selectboxObj[document.getElementById("selected").selectedIndex].text; }
    
    if(itemSelected == "select1") {
      if (selectedValue != "") {
        // empty the select box
        for (i = selectboxObj.length - 1; i>=0; i--) {
          selectboxObj.options[i] = null;
        }
        // add items to selected box
        selectboxObj[selectboxObj.length] = new Option(selectedText,selectedValue);
      }
    }
    
    if (itemSelected == "deselect1") {
      if(selectedboxValue != "") {
        // remove items from selected box
        document.getElementById("selected").remove(document.getElementById("selected").selectedIndex)
      }
    }
    
    updateButtons();
    // remove items from large list that are in the other lists
    removeUsed();
    return true;
}

function saveForm() {
  var funcSel = document.getElementById('functionselect');
  var elSel = document.getElementById('selected');
  var i,j;
  var selectList = "";
  var excludededList = "";
  var htmlContents = "";
  
  if (!validateEntry()) {
    return false;
  }
  
  // assemble the list of values from the selected box
  for (i = elSel.length - 1; i>=0; i--) {
      if(elSel.options[i].value != ""){
            if(selectList != "") { selectList = selectList + ","; }
            selectList = selectList + elSel.options[i].value;
      }
    }
    // ok time to build up the hidden variables to pass for saving
    htmlContents = "<input type=\"hidden\" name=\"selGroups\" value=" + selectList + "> ";
    document.getElementById("hiddenVals").innerHTML = htmlContents;
    // alert(htmlContents);
    return true;
}

function validateEntry() {
    var retVal = true;
    var elemPP = document.getElementById("require_pp");
    var elem = document.getElementById("selected");   
    var qtyElem = document.getElementById("QtyForIncentive");
    var elemProgram = document.getElementById("LimitID");
    var msg = '';
    
    if (elemPP == null || !elemPP.checked) {
      if (elem != null && elem.options.length == 0) {
          retVal = false;
          msg += '* <% Sendb(Copient.PhraseLib.Lookup("CPE-rew-membership.selectpoints", LanguageID)) %>'
          elem.focus();
      } else if (elem !=null && elemProgram != null) {
          elemProgram.value = elem.options[0].value;
      }
      if (qtyElem != null) {
          // trim the string
          var qtyVal = qtyElem.value.replace(/^\s+|\s+$/g, ''); 
          if (qtyVal == "" || isNaN(qtyVal)) {
              retVal = false;
              if (msg != '') { msg += '\n\r\n\r'; }
              msg += '* <% Sendb(Copient.PhraseLib.Lookup("CPE-rew-membership.enterquantity", LanguageID)) %>';
              qtyElem.focus();
              qtyElem.select();
            }
        }
    }
    if (msg != '') {
        alert(msg);
    }
    return retVal;
}

function updateButtons(){
  var elemDisallowEditPpOpt = document.getElementById("DisallowEditPpOpt");
  var selectObj = document.getElementById('selected');
  
  if (elemDisallowEditPpOpt != null && elemDisallowEditPpOpt.value == '1') {
      document.getElementById('select1').disabled=true;
      document.getElementById('deselect1').disabled=true;
  } else {
    if (selectObj.length == 0) {
      document.getElementById('select1').disabled=false;
      document.getElementById('deselect1').disabled=true;
    } else {
      document.getElementById('select1').disabled=false;
      document.getElementById('deselect1').disabled=false;
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
  
  If (Request.QueryString("save") <> "") Then
    
    If Not (bUseTemplateLocks And bDisallowEditPp) Then
      If (LimitID <> 0) Then
        MyCommon.QueryStr = "update OfferConditions with (RowLock) set LinkID=" & LimitID & " where ConditionID=" & ConditionID & ";"
        MyCommon.LRT_Execute()
      End If
    End If
    
    If Not (bUseTemplateLocks And bDisallowEditRewards) Then
      If (Request.QueryString("granted") <> "") Then
        MyCommon.QueryStr = "update OfferConditions with (RowLock) set GrantTypeID=" & Request.QueryString("granted") & " where ConditionID=" & ConditionID & ";"
        MyCommon.LRT_Execute()
      End If
    End If

    If Not (bUseTemplateLocks And bDisallowEditValue) Then
      If (Request.QueryString("radioValue") <> "") Then
        iRadioValue = Int16.Parse(Request.QueryString("radioValue"))
        If iRadioValue = 4 Then
          MyCommon.QueryStr = "update OfferConditions with (RowLock) set PointsRedeemInstant=1,PointsUseEarnedValue=0,PointsUseNetValue=0 where ConditionID=" & ConditionID & ";"
        ElseIf iRadioValue = 2 Then
          MyCommon.QueryStr = "update OfferConditions with (RowLock) set PointsRedeemInstant=0,PointsUseEarnedValue=1,PointsUseNetValue=0 where ConditionID=" & ConditionID & ";"
        ElseIf iRadioValue = 1 Then
          MyCommon.QueryStr = "update OfferConditions with (RowLock) set PointsRedeemInstant=0,PointsUseEarnedValue=0,PointsUseNetValue=1 where ConditionID=" & ConditionID & ";"
        Else
          MyCommon.QueryStr = "update OfferConditions with (RowLock) set PointsRedeemInstant=0,PointsUseEarnedValue=0,PointsUseNetValue=0 where ConditionID=" & ConditionID & ";"
        End If
        MyCommon.LRT_Execute()
      End If
    
      If (Request.QueryString("valuetype") <> "") Then
        MyCommon.QueryStr = "update OfferConditions with (RowLock) set QtyUnitType=5 where ConditionID=" & ConditionID & ";"
        MyCommon.LRT_Execute()
      End If
      
      If (Request.QueryString("tier0") <> "" And (Request.QueryString("Tiered") = "False")) Then
        'dbo.pt_ConditionTiers_Update @ConditionID bigint, @TierLevel int,@AmtRequired decimal(12,3)
        MyCommon.QueryStr = "dbo.pt_ConditionTiers_Update"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@ConditionID", SqlDbType.BigInt).Value = ConditionID
        MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int).Value = 0
        MyCommon.LRTsp.Parameters.Add("@AmtRequired", SqlDbType.Decimal).Value = MyCommon.Extract_Val(Request.QueryString("tier0"))
        MyCommon.LRTsp.ExecuteNonQuery()
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
          End If
          MyCommon.LRTsp.ExecuteNonQuery()
          MyCommon.Close_LRTsp()
        Next
      End If
    End If

    If (Request.QueryString("IsTemplate") = "IsTemplate") Then
      ' time to update the status bits for the templates
      Dim form_Disallow_Edit As Integer = 0
      Dim form_require_al As Integer = 0
      Dim iDisallowEditAl As Integer = 0
      Dim iDisallowEditValue As Integer = 0
      Dim iDisallowEditRewards As Integer = 0
      
      Disallow_Edit = False
      RequirePP = False
      bDisallowEditValue = False
      bDisallowEditRewards = False
      bDisallowEditPp = False
      
      If (Request.QueryString("Disallow_Edit") = "on") Then
        form_Disallow_Edit = 1
        Disallow_Edit = True
      End If
      If (Request.QueryString("require_pp") <> "") Then
        form_require_al = 1
        RequirePP = True
      End If
      If (Request.QueryString("DisallowEditPp") = "on") Then
        iDisallowEditAl = 1
        bDisallowEditPp = True
      End If
      If (Request.QueryString("DisallowEditValue") = "on") Then
        iDisallowEditValue = 1
        bDisallowEditValue = True
      End If
      If (Request.QueryString("DisallowEditRewards") = "on") Then
        iDisallowEditRewards = 1
        bDisallowEditRewards = True
      End If
      
      ' both requiring and locking the advanced limit is not permitted 
      If (form_require_al = 1 AndAlso (form_Disallow_Edit = 1 Or iDisallowEditAl = 1)) Then
        infoMessage = Copient.PhraseLib.Lookup("offer-con.lockeddenied", LanguageID)
      Else
        MyCommon.QueryStr = "update OfferConditions with (RowLock) set Disallow_Edit=" & form_Disallow_Edit & _
        ",RequiredFromTemplate=" & form_require_al & _
        ",DisallowEdit1=" & iDisallowEditAl & _
        ",DisallowEdit2=" & iDisallowEditValue & _
        ",DisallowEdit3=" & iDisallowEditRewards & _
        " where ConditionID=" & ConditionID
        MyCommon.LRT_Execute()
        ' update the advanced limit requirement
        MyCommon.QueryStr = "update OfferConditions with (RowLock) set RequiredFromTemplate = " & _
                            IIf(Request.QueryString("require_pp") <> "", 1, 0) & "where ConditionID = " & ConditionID
        MyCommon.LRT_Execute()
      End If
    End If
    ' udpate the flags for this condition
    ' update the flags
    MyCommon.QueryStr = "update OfferConditions with (RowLock) set TCRMAStatusFlag=2,CRMAStatusFlag=2 where ConditionID=" & ConditionID
    MyCommon.LRT_Execute()
    MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where offerid=" & OfferID
    MyCommon.LRT_Execute()
    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-advlimit", LanguageID))
    If (infoMessage = "") Then
      CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
    End If
  End If
  
  Send("<script type=""text/javascript"">")
    Send("function ChangeParentDocument() { ")
    Send("  if (opener != null) {")
    Send("    var newlocation = 'offer-con.aspx?OfferID=" & OfferID & "'; ")
    Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
  Send("    opener.location = 'offer-con.aspx?OfferID=" & OfferID & "'; ")
    Send("    } ")
    Send("  }")
    Send("}")
  Send("</script>")
%>
<form action="#" id="mainform" name="mainform" onsubmit="return validateEntry();">
  <div id="intro">
    <input type="hidden" id="OfferID" name="OfferID" value="<% sendb(OfferID) %>" />
    <input type="hidden" id="Name" name="Name" value="<% sendb(Name) %>" />
    <input type="hidden" id="ConditionID" name="ConditionID" value="<% sendb(ConditionID) %>" />
    <input type="hidden" name="LimitID" id="LimitID" value="<% Sendb(LimitID) %>" />    
    <input type="hidden" id="IsTemplate" name="IsTemplate" value="<% 
            if(istemplate)then 
            sendb("IsTemplate")
            else 
            sendb("Not") 
            end if
        %>" />
    <%If (isTemplate) Then
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.advlimitcondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
      Else
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.advlimitcondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
      End If
      If (bUseTemplateLocks And bDisallowEditPp) Then
        Send("<input type=""hidden"" id=""DisallowEditPpOpt"" name=""DisallowEditPpOpt"" value=""1"" />")
      Else
        Send("<input type=""hidden"" id=""DisallowEditPpOpt"" name=""DisallowEditPpOpt"" value=""0"" />")
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
        <input type="checkbox" class="tempcheck" id="temp-employees" name="Disallow_Edit"<% if(disallow_edit)then sendb(" checked=""checked""") %> />
        <label for="temp-employees"><% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
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
      <div class="box" id="selector">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.advlimits", LanguageID))%>
          </span>
          <% If (IsTemplate) Then%>
          <span class="tempRequire">
            <input type="checkbox" class="tempcheck" id="require_pp1" name="require_pp"<% if(requirepp)then sendb(" checked=""checked""") %> />
            <label for="require_pp"><% Sendb(Copient.PhraseLib.Lookup("term.required", LanguageID))%></label>
          </span>
          <span class="tempLocked">
            <input type="checkbox" class="tempcheck" id="DisallowEditPp1" name="DisallowEditPp" <% if(bDisallowEditPp)then send(" checked=""checked""") %> />
            <label for="temp-Tiers"><% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
          </span>
          <% ElseIf (bUseTemplateLocks) Then%>
            <% If (RequirePP) Then%>
            <span class="tempRequire">
              <input type="checkbox" class="tempcheck" id="require_pp2" name="require_pp" disabled="disabled" checked="checked" />
              <label for="require_pp"><% Sendb(Copient.PhraseLib.Lookup("term.required", LanguageID))%></label>
            </span>
            <% Else If (bDisallowEditPp) Then%>
            <span class="tempRequire">
              <input type="checkbox" class="tempcheck" id="DisallowEditPp2" name="DisallowEditPp" disabled="disabled" checked="checked" />
              <label for="temp-Tiers"><% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
            </span>
            <% End If%>
          <% End If%>
        </h2>
        <input type="radio" id="functionradio1" name="functionradio" <% if(MyCommon.Fetch_SystemOption(175)= "1" ) then sendb(" checked=""checked""") %> <% sendb(disabledattribute) %> /><label for="functionradio1"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
        <input type="radio" id="functionradio2" name="functionradio" <% if(MyCommon.Fetch_SystemOption(175)= "2" ) then sendb(" checked=""checked""") %> <% sendb(disabledattribute) %> /><label for="functionradio2"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
        <input class="medium" onkeyup="handleKeyUp(200);" id="functioninput" name="functioninput" type="text" maxlength="100" value=""<% sendb(disabledattribute) %> /><br />
        <select class="longer" id="functionselect" name="functionselect" size="16"<% sendb(disabledattribute) %>>
          <%
            MyCommon.QueryStr = "Select AL.LimitID,AL.Name from CM_AdvancedLimits as AL with (NoLock) where deleted=0 and LimitID is not null order by Name"
            rst = MyCommon.LRT_Select
            For Each row In rst.Rows
              Send("<option value=" & MyCommon.NZ(row.Item("LimitID"), -1) & ">" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
            Next
          %>
        </select>
        <br />
        <br class="half" />
        <input class="regular select" id="select1" name="select1" type="button" value="&#9660; <% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID)) %>" onclick="handleSelectClick('select1');"<% sendb(disabledattribute) %> />&nbsp;
        <input class="regular deselect" id="deselect1" name="deselect1" type="button" value="<% Sendb(Copient.PhraseLib.Lookup("term.deselect", LanguageID)) %> &#9650;" onclick="handleSelectClick('deselect1');"<% sendb(disabledattribute) %> /><br />
        <br class="half" />
        <select class="longer" id="selected" name="selected" size="2"<% sendb(disabledattribute) %>>
          <%
            MyCommon.QueryStr = "select AL.LimitID,AL.Name,OC.LinkID from CM_AdvancedLimits as AL with (NoLock) " & _
                                "left join OfferConditions as OC with (NoLock) on OC.LinkID=AL.LimitID " & _
                                "where OC.ConditionID=" & ConditionID
            rst = MyCommon.LRT_Select
            For Each row In rst.Rows
              Send("<option value=""" & MyCommon.NZ(row.Item("LimitID"), -1) & """>" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
            Next
          %>
        </select>
        <hr class="hidden" />
      </div>
    </div>

    <div id="gutter">
    </div>
    
    <div id="column2">
      <div class="box" id="value">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.value", LanguageID))%>
          </span>
        </h2>
        <% If (IsTemplate or (bUseTemplateLocks and bDisallowEditValue)) Then%>
        <span class="temp">
          <input type="checkbox" class="tempcheck" id="DisallowEditValue" name="DisallowEditValue"<% if(bDisallowEditValue)then send(" checked=""checked""") %> <% If(bUseTemplateLocks) Then sendb(" disabled=""disabled""") %>/>
          <label for="temp-Tiers"><% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
        </span>
        <br class="printonly" />
        <% End If%>
        <label for="tier0"><% Sendb(Copient.PhraseLib.Lookup("condition.valueneeded", LanguageID))%></label>
        <br />
        <%
          MyCommon.QueryStr = "select LinkID,Tiered,O.Numtiers,QtyUnitType,O.OfferID,CT.TierLevel,CT.AmtRequired from OfferConditions as OC with (NoLock) left  join Offers as O with (NoLock) on O.OfferID=OC.OfferID left  join ConditionTiers as CT with (NoLock) on OC.ConditionID=CT.ConditionID where OC.ConditionID=" & ConditionID
          rst = MyCommon.LRT_Select()
          Dim q As Integer
          q = 1
          If (bUseTemplateLocks And bDisallowEditValue) Then
            sDisabled = " disabled=""disabled"""
          Else
            sDisabled = ""
          End If
          For Each row In rst.Rows
            If (MyCommon.NZ(row.Item("Tiered"), 0) = 0) Then
              Send("<input class=""shorter"" id=""tier0"" name=""tier0"" type=""text"" maxlength=""9"" value=""" & Int(MyCommon.NZ(row.Item("AmtRequired"), 0)) & """" & sDisabled & " /><br />")
            Else
              Tiered = True
              Send("<label for=""tier" & q & """><b>" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & q & ":</b></label> <input class=""shorter"" id=""tier" & q & """ name=""tier" & q & """ type=""text"" value=""" & Int(MyCommon.NZ(row.Item("AmtRequired"), 0)) & """" & sDisabled & " /><br />")
            End If
            q = q + 1
          Next
          Send("<input type=""hidden"" name=""NumTiers"" value=""" & MyCommon.NZ(row.Item("NumTiers"), "") & """ />")
          Send("<input type=""hidden"" name=""Tiered"" value=""" & MyCommon.NZ(row.Item("Tiered"), "") & """ />")
          MyCommon.QueryStr = "select LinkID,ExcludedID,PointsRedeemInstant,MinOrderItemsOnly,GrantTypeID,DoNotItemDistribute,PointsUseEarnedValue,PointsUseNetValue from OfferConditions with (NoLock) where ConditionID=" & ConditionID & ";"
          rst = MyCommon.LRT_Select
          For Each row In rst.Rows
          Next
        %>
        <br class="half" />
        <%
          Sendb(Copient.PhraseLib.Lookup("condition.satisfy", LanguageID))
          If (MyCommon.NZ(row.Item("pointsredeeminstant"), False)) Then
            iRadioValue = 4
          ElseIf (MyCommon.NZ(row.Item("pointsuseearnedvalue"), False)) Then
            iRadioValue = 2
          ElseIf (MyCommon.NZ(row.Item("pointsusenetvalue"), False)) Then
            iRadioValue = 1
          Else
            iRadioValue = 0
          End If
        %>
        <br />
        <input class="radioValue" id="RadioValue1" name="radioValue" value="0" type="radio"<% if(iRadioValue=0)then sendb(" checked=""checked""") %><% If(bUseTemplateLocks and bDisallowEditValue) Then sendb(" disabled=""disabled""") %> />
        <label for="RadioValue1"><% Sendb(Copient.PhraseLib.Lookup("condition.alPreviousValue", LanguageID))%></label>
        <br />
        <input class="radioValue" id="RadioValue2" name="radioValue" value="2" type="radio"<% if(iRadioValue=2)then sendb(" checked=""checked""") %><% If(bUseTemplateLocks and bDisallowEditValue) Then sendb(" disabled=""disabled""") %> />
        <label for="RadioValue2"><% Sendb(Copient.PhraseLib.Lookup("condition.alCurrentValue", LanguageID))%></label>
        <br />
        <input class="radioValue" id="RadioValue3" name="radioValue" value="4" type="radio"<% if(iRadioValue=4)then sendb(" checked=""checked""") %><% If(bUseTemplateLocks and bDisallowEditValue) Then sendb(" disabled=""disabled""") %> />
        <label for="RadioValue3"><% Sendb(Copient.PhraseLib.Lookup("condition.alTotalValue", LanguageID))%></label>
        <br />
        <input class="radioValue" id="RadioValue4" name="radioValue" value="1" type="radio"<% if(iRadioValue=1)then sendb(" checked=""checked""") %><% If(bUseTemplateLocks and bDisallowEditValue) Then sendb(" disabled=""disabled""") %> />
        <label for="RadioValue4"><% Sendb(Copient.PhraseLib.Lookup("condition.alNetValue", LanguageID))%></label>
        <br />
        <hr class="hidden" />
      </div>
      <div class="box" id="grants"<%if(tiered)then send(" style=""visibility: hidden;""") %>>
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.rewards", LanguageID))%>
          </span>
        </h2>
        <% If (IsTemplate or IsTemplate or (bUseTemplateLocks and bDisallowEditRewards)) Then%>
        <span class="temp">
          <input type="checkbox" class="tempcheck" id="DisallowEditRewards" name="DisallowEditRewards"<% if(bDisallowEditRewards)then send(" checked=""checked""") %><% If(bUseTemplateLocks) Then sendb(" disabled=""disabled""") %> />
          <label for="temp-Tiers"><% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
        </span>
        <br class="printonly" />
        <% End If%>
        <% Sendb(Copient.PhraseLib.Lookup("condition.rewardsgranted", LanguageID))%>
        <br />
        <input class="radio" id="eachtime" name="granted" value="3" type="radio"<% if(MyCommon.NZ(row.item("granttypeid"), 0)=3)then sendb(" checked=""checked""") %><% If(bUseTemplateLocks and bDisallowEditRewards) Then sendb(" disabled=""disabled""") %> />
        <label for="eachtime"><% Sendb(Copient.PhraseLib.Lookup("condition.eachtime", LanguageID))%></label>
        <br />
        <input class="radio" id="equalto" name="granted" value="1" type="radio"<% if(MyCommon.NZ(row.item("granttypeid"),0)=1)then sendb(" checked=""checked""") %><% If(bUseTemplateLocks and bDisallowEditRewards) Then sendb(" disabled=""disabled""") %> />
        <label for="equalto"><% Sendb(Copient.PhraseLib.Lookup("condition.equalto", LanguageID))%></label>
        <br />
        <input class="radio" id="greaterthan" name="granted" value="2" type="radio"<% if(MyCommon.NZ(row.item("granttypeid"),0)=2)then sendb(" checked=""checked""") %><% If(bUseTemplateLocks and bDisallowEditRewards) Then sendb(" disabled=""disabled""") %> />
        <label for="greaterthan">
          <% Sendb(Copient.PhraseLib.Lookup("condition.greaterthan", LanguageID))%>
        </label>
        <br />
      </div>
    </div>
  </div>
</form>

<script type="text/javascript">
<% If (CloseAfterSave) Then %>
    window.close();
<% End If %>
handleKeyUp(9999);
updateButtons();

</script>

<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd("mainform", "functioninput")
  MyCommon = Nothing
  Logix = Nothing
%>
