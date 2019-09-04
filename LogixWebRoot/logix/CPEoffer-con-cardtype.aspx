<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: CPEoffer-con-cardtype.aspx 
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
  Dim rst2 As DataTable
  Dim row As DataRow
  Dim OfferID As Long
  Dim Name As String = ""
  Dim ConditionID As String
  Dim IsTemplate As Boolean = False
  Dim FromTemplate As Boolean = False
  Dim Disallow_Edit As Boolean = True
  Dim Household As Boolean = False
  Dim DisabledAttribute As String = ""
  Dim i As Integer
  Dim roid As Integer
  Dim historyString As String
  Dim CloseAfterSave As Boolean = False
  Dim Ids() As String
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim RequireCT As Boolean = False
  Dim HasRequiredCT As Boolean = False
  Dim EngineID As Integer = 2
  Dim BannersEnabled As Boolean = True
  Dim FullListSelect As New StringBuilder()
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "CPEoffer-con-cardtype.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  OfferID = Request.QueryString("OfferID")
  ConditionID = Request.QueryString("ConditionID")
  If (Request.QueryString("EngineID") <> "") Then
    EngineID = MyCommon.Extract_Val(Request.QueryString("EngineID"))
  Else
    MyCommon.QueryStr = "select EngineID from OfferIDs where OfferID=" & OfferID & ";"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
      EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), 0)
    End If
  End If
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")

  MyCommon.QueryStr = "select RewardOptionID from CPE_RewardOptions with (NoLock) where IncentiveID=" & OfferID & " and TouchResponse=0 and Deleted=0;"
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    roid = rst.Rows(0).Item("RewardOptionID")
  End If
  
  If (Request.QueryString("save") <> "") Then
    If roid > 0 Then
      If (Request.QueryString("selTypes") = "") AndAlso (Request.QueryString("require_ct") = "") Then
        infoMessage = Copient.PhraseLib.Lookup("CPEoffer-con-cardtype.SelectCardType", LanguageID)
      Else
        ' Check to see if a card type condition is required by the template, if applicable
        MyCommon.QueryStr = "select CardTypeID from CPE_IncentiveCardTypes with (NoLock) " & _
                            "where RewardOptionID=" & roid & " and RequiredFromTemplate=1 and Deleted=0;"
        rst = MyCommon.LRT_Select
        HasRequiredCT = (rst.Rows.Count > 0)
        
        ' We got some selected types so let's blow out all the existing ones
        MyCommon.QueryStr = "update CPE_IncentiveCardTypes with (RowLock) set Deleted=1 " & _
                            "where RewardOptionID=" & roid & " and Deleted=0;"
        MyCommon.LRT_Execute()
        
        ' Handle the selected card types
        If (Request.QueryString("selTypes") <> "") Then
          historyString = Copient.PhraseLib.Lookup("history.con-cardtype-edit", LanguageID) & ": " & Request.QueryString("selTypes")
          Ids = Request.QueryString("selTypes").Split(",")
          For i = 0 To Ids.Length - 1
            MyCommon.QueryStr = "insert into CPE_IncentiveCardTypes with (RowLock) (RewardOptionID, CardTypeID, Deleted, LastUpdate, RequiredFromTemplate) " & _
                                "values (" & roid & ", " & Ids(i) & ", 0, getdate(), " & IIf(HasRequiredCT, "1", "0") & ")"
            MyCommon.LRT_Execute()
          Next
        ElseIf HasRequiredCT Then
          MyCommon.QueryStr = "insert into CPE_IncentiveCardTypes with (RowLock) (RewardOptionID, Deleted, LastUpdate, RequiredFromTemplate) " & _
                              "values (" & roid & ", 0, getdate(), 1)"
          MyCommon.LRT_Execute()
        End If
        
        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
        MyCommon.LRT_Execute()
        MyCommon.Activity_Log(3, OfferID, AdminUserID, historyString)
      End If
    Else
      infoMessage = Copient.PhraseLib.Lookup("CPEoffer-con-cardtype.ErrorSaving", LanguageID)
    End If
    
    If (infoMessage = "") Then
      CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
    Else
      CloseAfterSave = False
    End If
    
  End If
  
  ' Dig the offer info out of the database
  MyCommon.QueryStr = "Select IncentiveName as Name, IsTemplate, FromTemplate " & _
                      "from CPE_Incentives as CPE with (NoLock) where IncentiveID=" & Request.QueryString("OfferID") & ";"
  rst = MyCommon.LRT_Select
  For Each row In rst.Rows
    Name = MyCommon.NZ(row.Item("Name"), "")
    IsTemplate = MyCommon.NZ(row.Item("IsTemplate"), False)
    FromTemplate = MyCommon.NZ(row.Item("FromTemplate"), False)
  Next
  
  ' Update the templates permission if necessary
  If (Request.QueryString("save") <> "" AndAlso Request.QueryString("IsTemplate") = "IsTemplate" AndAlso infoMessage = "") Then
    ' Update the status bits for the templates
    Dim form_Disallow_Edit As Integer = 0
    Dim form_Require_CT As Integer = 0
    
    If (Request.QueryString("Disallow_Edit") = "on") Then
      form_Disallow_Edit = 1
    End If
    
    If (Request.QueryString("require_ct") <> "") Then
      form_Require_CT = 1
    End If
    
    ' Both requiring and locking the card type condition is not permitted 
    If (form_Disallow_Edit = 1 AndAlso form_Require_CT = 1) Then
      infoMessage = Copient.PhraseLib.Lookup("offer-con.lockeddenied", LanguageID)
      MyCommon.QueryStr = "update CPE_IncentiveCardTypes with (RowLock) set DisallowEdit=1, RequiredFromTemplate=0 " & _
                          "where RewardOptionID=" & roid & " and Deleted=0;"
    Else
      MyCommon.QueryStr = "update CPE_IncentiveCardTypes with (RowLock) set DisallowEdit=" & form_Disallow_Edit & ", RequiredFromTemplate=" & form_Require_CT & " " & _
                          "where RewardOptionID=" & roid & " and Deleted=0;"
    End If
    MyCommon.LRT_Execute()
    
    ' If necessary, create an empty card type condition
    If (form_Require_CT = 1) Then
      MyCommon.QueryStr = "select CardTypeID from CPE_IncentiveCardTypes with (NoLock) where RewardOptionID=" & roid & " and Deleted=0;"
      rst = MyCommon.LRT_Select
      If (rst.Rows.Count = 0) Then
        MyCommon.QueryStr = "insert into CPE_IncentiveCardTypes with (RowLock) (RewardOptionID, Deleted, LastUpdate, RequiredFromTemplate) " & _
                            "values (" & roid & ", 0, getdate(), 1)"
        MyCommon.LRT_Execute()
      End If
    End If
    
    If (infoMessage = "") Then
      CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
    Else
      CloseAfterSave = False
    End If
  End If
  
  If (IsTemplate Or FromTemplate) Then
    ' Dig the permissions if it's a template
    MyCommon.QueryStr = "select DisallowEdit, RequiredFromTemplate from CPE_IncentiveCardTypes with (NoLock) where RewardOptionID=" & roid & " and Deleted=0;"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      Disallow_Edit = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit"), False)
      RequireCT = MyCommon.NZ(rst.Rows(0).Item("RequiredFromTemplate"), False)
    Else
      Disallow_Edit = False
    End If
  End If
  
  MyCommon.QueryStr = "select HHEnable from CPE_RewardOptions with (NoLock) where IncentiveID=" & OfferID & " and TouchResponse=0 and Deleted=0;"
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    Household = MyCommon.NZ(rst.Rows(0).Item("HHEnable"), False)
  End If
  
  If Not isTemplate Then
    DisabledAttribute = IIf(Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Edit), "", " disabled=""disabled""")
  Else
    DisabledAttribute = IIf(Logix.UserRoles.EditTemplates, "", " disabled=""disabled""")
  End If
  
  Send_HeadBegin("term.offer", "term.cardtypecondition", OfferID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
%>

<%
  Send("<script type=""text/javascript"">")

  Send("// This is the javascript array holding the function list")
  Send("// The PrintJavascriptArray ASP function can be used to print this array.")

  FullListSelect.Append("<select class=""longer"" id=""functionselect"" name=""functionselect"" multiple=""multiple"" size=""12"">")
  
  'Populate the JavaScript array that holds the list of selectable card types
    MyCommon.QueryStr = "select CardTypeID, Description, PhraseID from CardTypes with (NoLock) where CardTypeID<>8 order by Description;"
  rst = MyCommon.LXS_Select
  If (rst.Rows.Count > 0) Then
    Sendb("var functionlist = Array(")
    For Each row In rst.Rows
      Sendb("""" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID, Convert.ToString(row.Item("Description"))).Replace("""", "\""") & """,")
    Next
    Send(""""");")
    Sendb("var vallist = Array(")
    For Each row In rst.Rows
      Sendb("""" & MyCommon.NZ(row.Item("CardTypeID"), -1) & """,")
      FullListSelect.Append("<option value=""" & MyCommon.NZ(row.Item("CardTypeID"), -1) & """>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID, Convert.ToString(row.Item("Description"))).Replace("'", "\'") & "<\/option>")
    Next
    Send(""""");")
  Else
    Sendb("var functionlist = Array();")
    Sendb("var vallist = Array();")
  End If
  
  FullListSelect.Append("<\/select>")
  Send("var fullList = '" & FullListSelect.ToString() & "';")
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
  var optVal;
  
  document.getElementById("functionselect").size = "12";
  
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
  
  if (textObj.value == '') {
    document.getElementById("typeList").innerHTML = fullList;
  } else {
    var newList = '<select class="longer" id="functionselect" name="functionselect" size="12" multiple="multiple">';
    for(i = 0; i < functionListLength; i++) {
      if(functionlist[i].search(re) != -1) {
        optVal = vallist[i<% Sendb(IIf(EngineID=6, "-1", "")) %>];
        if (optVal != "") {
          newList += '<option value="' + optVal + '"> ' + functionlist[i] + '<\/option>';
          numShown++;
        }
      }
      // Stop when the number to show is reached
      if(numShown == maxNumToShow) {
        break;
      }
    }
    newList += '<\/select>'
    document.getElementById("typeList").innerHTML = newList;
  }
  
  removeUsed(true);
  
  // When options list whittled to one, select that entry
  if(selectObj.length == 1) {
    selectObj.options[0].selected = true;
  }
}

function removeUsed(bSkipKeyUp) {
  if (!bSkipKeyUp) handleKeyUp(99999);
  // this function will remove items from the functionselect box
  // that are used in the selected box
  
  var funcSel = document.getElementById('functionselect');
  var exSel = document.getElementById('selected');
  var i,j;
  
  for (i = exSel.length - 1; i>=0; i--) {
    for (j=funcSel.length-1;j>=0; j--) {
      if (funcSel.options[j].value == exSel.options[i].value) {
        funcSel.options[j] = null;
      }
    }
  }
}

// this function gets the selected value and loads the appropriate
// php reference page in the display frame
// it can be modified to perform whatever action is needed, or nothing
function handleSelectClick(itemSelected) {
  var funcSel = document.getElementById('functionselect');
  var exSel = document.getElementById('selected');
  var i,j;
  
  textObj = document.forms[0].functioninput;
  selectObj = document.forms[0].functionselect;
  selectedValue = document.getElementById("functionselect").value;
  if (selectedValue != "") {
    selectedText = selectObj[document.getElementById("functionselect").selectedIndex].text;
  }
  selectboxObj = document.forms[0].selected;
  selectedboxValue = document.getElementById("selected").value;
  if (selectedboxValue != "") {
    selectedboxText = selectboxObj[document.getElementById("selected").selectedIndex].text;
  }
  
  if (itemSelected == "select1") {
    if (selectedValue != "") {
      // add items to selected box
      document.getElementById('deselect1').disabled=false;
      document.getElementById('save').disabled=false;
      
      while (selectObj.selectedIndex != -1) {
        selectedText = selectObj.options[selectObj.selectedIndex].text;
        selectedValue = selectObj.options[selectObj.selectedIndex].value;
        selectboxObj[selectboxObj.length] = new Option(selectedText,selectedValue);
        selectObj[selectObj.selectedIndex].selected = false;
      }
    }
  }
  
  if (itemSelected == "deselect1") {
    if (selectedboxValue != "") {
      // remove items from selected box
      while (document.getElementById("selected").selectedIndex != -1) {
        document.getElementById("selected").remove(document.getElementById("selected").selectedIndex);
      }
      if (selectboxObj.length == 0) {
        // nothing in the select box so disable deselect
        document.getElementById('deselect1').disabled=true;
        // this being the case, let's also disable the page's save button
        // (but not if it's a template with "required" checked)
        if (document.getElementById("require_ct") != null) {
          if (document.getElementById("require_ct").checked == false) {
            document.getElementById('save').disabled=true;
          }
        } else {
          document.getElementById('save').disabled=true;
        }
      }
      if (selectboxObj.length == 0 || selectedboxValue == "") {
        document.getElementById('select1').disabled=false;
      }
    }
  }
  
  // remove items from large list that are in the other lists
  removeUsed(false);
  return true;
}

function saveForm() {
  var funcSel = document.getElementById('functionselect');
  var exSel = document.getElementById('selected');
  var i,j;
  var selectList = "";
  var htmlContents = "";
  
  // assemble the list of values from the selected box
  for (i = exSel.length - 1; i>=0; i--) {
    if(exSel.options[i].value != ""){
      if(selectList != "") { selectList = selectList + ","; }
      selectList = selectList + exSel.options[i].value;
    }
  }
  
  // time to build up the hidden variables to pass for saving
  htmlContents = "<input type=\"hidden\" name=\"selTypes\" value=" + selectList + "> ";
  document.getElementById("hiddenVals").innerHTML = htmlContents;
  
  // alert(htmlContents);
  return true;
}

function updateButtons() {
  var selectedValue = document.getElementById('selected').options[0].value;
  
  if (document.getElementById('selected').length > 0){
    document.getElementById('select1').disabled=false; 
    document.getElementById('deselect1').disabled=false;
  }
  <%
   If Not isTemplate Then   
     If Not (Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Edit)) Then
      Send("  disableAll();")
     End If 
   Else
     If Not (Logix.UserRoles.EditTemplates) Then
      Send("  disableAll();")
     End If
   End If        
  %>
}

function handleRequiredToggle() {
  if(document.forms[0].selected.length == 0) {
    if (document.getElementById("require_ct").checked == false) {
      document.getElementById('save').disabled=true;
    } else {
      document.getElementById('save').disabled=false;
    }
  }
  if (document.getElementById("require_ct").checked == true) {
    document.getElementById("Disallow_Edit").checked=false;
  }
}

function disableAll() {
  document.getElementById('select1').disabled=true;
  document.getElementById('deselect1').disabled=true;
  document.getElementById('functionselect').disabled=true;
  document.getElementById('selected').disabled=true;
}
</script>
<%
  Send("<script type=""text/javascript"">")
  Send("function ChangeParentDocument() { ")
  If (EngineID = 3) Then
    Send("  opener.location = 'web-offer-con.aspx?OfferID=" & OfferID & "'; ")
  ElseIf (EngineID = 5) Then
    Send("  opener.location = 'email-offer-con.aspx?OfferID=" & OfferID & "'; ")
  ElseIf (EngineID = 6) Then
    Send("  opener.location = 'CAM/CAM-offer-con.aspx?OfferID=" & OfferID & "'; ")
  Else
    Send("  opener.location = 'CPEoffer-con.aspx?OfferID=" & OfferID & "'; ")
  End If
  Send("} ")
  Send("</script>")
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
%>
<form action="#" name="mainform" id="mainform" onsubmit="saveForm()">
  <span id="hiddenVals"></span>
  <input type="hidden" id="OfferID" name="OfferID" value="<% Sendb(OfferID) %>" />
  <input type="hidden" id="Name" name="Name" value="<% Sendb(Name) %>" />
  <input type="hidden" id="ConditionID" name="ConditionID" value="<% Sendb(ConditionID) %>" />
  <input type="hidden" id="EngineID" name="EngineID" value="<% Sendb(EngineID) %>" />
  <input type="hidden" id="IsTemplate" name="IsTemplate" value="<% Sendb(IIf(IsTemplate, "IsTemplate", "Not"))%> " />
  <div id="intro">
    <%
      Send("<h1 id=""title"">" & IIf(IsTemplate, Copient.PhraseLib.Lookup("term.template", LanguageID), Copient.PhraseLib.Lookup("term.offer", LanguageID)) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.cardtypecondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
      Send("<div id=""controls"">")
      If IsTemplate Then
        Send("  <span class=""temp"">")
        Send("    <input type=""checkbox"" class=""tempcheck"" id=""Disallow_Edit"" name=""Disallow_Edit""" & IIf(Disallow_Edit, " checked=""checked""", "") & " />")
        Send("    <label for=""Disallow_Edit"">" & Copient.PhraseLib.Lookup("term.locked", LanguageID) & "</label>")
        Send("  </span>")
      End If
      If Not IsTemplate Then
        If (Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Edit)) Then
          Send_Save()
        End If
      Else
        If (Logix.UserRoles.EditTemplates) Then
          Send_Save()
        End If
      End If
      Send("</div>")
    %>
  </div>
  <div id="main">
    <%
      If (infoMessage <> "") Then
        Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
      End If
    %>
    <div id="column1">
      <div class="box" id="selector">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.cardtypes", LanguageID))%>
          </span>
          <%
            If (IsTemplate) Then
              Send("<span class=""tempRequire"">")
              Send("  <input type=""checkbox"" class=""tempcheck"" id=""require_ct"" name=""require_ct"" onclick=""handleRequiredToggle();""" & IIf(RequireCT, " checked=""checked""", "") & " />")
              Send("  <label for=""require_ct"">" & Copient.PhraseLib.Lookup("term.required", LanguageID) & "</label>")
              Send("</span>")
            ElseIf (FromTemplate And RequireCT) Then
              Send("<span class=""tempRequire"">* " & Copient.PhraseLib.Lookup("term.required", LanguageID) & "</span>")
            End If
          %>
        </h2>
        <input type="radio" id="functionradio1" name="functionradio" checked="checked"<% sendb(disabledattribute) %> /><label for="functionradio1"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
        <input type="radio" id="functionradio2" name="functionradio"<% sendb(disabledattribute) %> /><label for="functionradio2"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
        <input type="text" class="medium" id="functioninput" name="functioninput" maxlength="100" onkeyup="handleKeyUp(99999);" value=""<% sendb(disabledattribute) %> /><br />
        <div id="typeList">
          <select class="longer" id="functionselect" name="functionselect" size="12"<% sendb(disabledattribute) %>>
          </select>
        </div>
        
        <br class="half" />
        
        <b><% Sendb(Copient.PhraseLib.Lookup("term.selectedcardtypes", LanguageID))%>:</b><br />
        <input type="button" class="regular select" id="select1" name="select1" value="&#9660; <% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID))%>" onclick="handleSelectClick('select1');" />&nbsp;
        <input type="button" class="regular deselect" id="deselect1" name="deselect1" value="<% Sendb(Copient.PhraseLib.Lookup("term.deselect", LanguageID))%> &#9650;" disabled="disabled" onclick="handleSelectClick('deselect1');" /><br />
        <br class="half" />
        <select class="longer" id="selected" name="selected" multiple="multiple" size="4"<% sendb(disabledattribute) %>>
          <%
            ' Find the currently selected card types on page load
            MyCommon.QueryStr = "select CardTypeID from CPE_IncentiveCardTypes with (NoLock) " & _
                                "where RewardOptionID=" & roid & " and Deleted=0 and CardTypeID is not null;"
            rst = MyCommon.LRT_Select
            For Each row In rst.Rows
              MyCommon.QueryStr = "select Description, PhraseID from CardTypes with (NoLock) " & _
                                  "where CardTypeID=" & MyCommon.NZ(row.Item("CardTypeID"), -1) & ";"
              rst2 = MyCommon.LXS_Select
              Sendb("<option value=""" & MyCommon.NZ(row.Item("CardTypeID"), 0) & """>")
              If rst2.Rows.Count > 0 Then
                Sendb(Copient.PhraseLib.Lookup(MyCommon.NZ(rst2.Rows(0).Item("PhraseID"), 0), LanguageID, Convert.ToString(rst2.Rows(0).Item("Description"))))
              Else
                Sendb(Copient.PhraseLib.Lookup("term.unknown", LanguageID))
              End If
              Send("</option>")
            Next
          %>
        </select>
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
removeUsed(false);
updateButtons();
</script>

<%
done:
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  Send_BodyEnd("mainform", "functioninput")
  MyCommon = Nothing
  Logix = Nothing
%>
