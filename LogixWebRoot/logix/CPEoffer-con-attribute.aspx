<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: CPEoffer-con-attribute.aspx 
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
  Dim isTemplate As Boolean
  Dim Disallow_Edit As Boolean = True
  Dim DisabledAttribute As String = ""
  Dim FromTemplate As Boolean = False
  Dim RequiredFromTemplate As Boolean = False
  Dim roid As Integer = 0
  Dim IncentiveAttributeID As Long = 0
  Dim AttributeTypeID As Long = 0
  Dim AttributeType As String = ""
  Dim AttributeValues As String = ""
  Dim Ids() As String
  Dim i As Integer
  Dim historyString As String = ""
  Dim CloseAfterSave As Boolean = False
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim RequirePP As Boolean = False
  Dim HasRequiredPP As Boolean = False
  Dim EngineID As Integer = 2
  Dim EngineSubTypeID As Integer = 0
  Dim BannersEnabled As Boolean = True
  Dim TierLevels As Integer = 1
  Dim t As Integer = 1
  Dim ProgramID As Integer = 0
  Dim PointsID As Integer = 0
  Dim NumArray(99) As Decimal
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "CPEoffer-con-attribute.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  
  OfferID = Request.QueryString("OfferID")
  IncentiveAttributeID = Request.QueryString("IncentiveAttributeID")
  If (Request.QueryString("EngineID") <> "") Then
    EngineID = MyCommon.Extract_Val(Request.QueryString("EngineID"))
    EngineSubTypeID = MyCommon.Extract_Val(Request.QueryString("EngineSubTypeID"))
  Else
    MyCommon.QueryStr = "select EngineID, EngineSubTypeID from OfferIDs with (NoLock) where OfferID=" & OfferID & ";"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
      EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), 0)
      EngineSubTypeID = MyCommon.NZ(rst.Rows(0).Item("EngineSubTypeID"), 0)
    End If
  End If
  If (Request.QueryString("AttributeTypeID") <> "") Then
    AttributeTypeID = MyCommon.Extract_Val(Request.QueryString("AttributeTypeID"))
  Else
    If IncentiveAttributeID > 0 Then
      MyCommon.QueryStr = "select AttributeTypeID from CPE_IncentiveAttributeTiers with (NoLock) " & _
                          "where IncentiveAttributeID=" & IncentiveAttributeID & ";"
      rst = MyCommon.LRT_Select
      If rst.Rows.Count > 0 Then
        AttributeTypeID = MyCommon.NZ(rst.Rows(0).Item("AttributeTypeID"), 0)
      End If
    Else
      AttributeTypeID = 0
    End If
  End If
  
  'Get attribute type name (if a type is specified)
  If AttributeTypeID > 0 Then
    MyCommon.QueryStr = "select Description from AttributeTypes with (NoLock) " & _
                        "where AttributeTypeID=" & AttributeTypeID & " and Deleted=0;"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
      AttributeType = MyCommon.NZ(rst.Rows(0).Item("Description"), "")
    End If
  End If
  
  'Get attribute values (if condition exists)
  If IncentiveAttributeID > 0 Then
    MyCommon.QueryStr = "select AttributeValues from CPE_IncentiveAttributeTiers where IncentiveAttributeID=" & IncentiveAttributeID & ";"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
      AttributeValues = MyCommon.NZ(rst.Rows(0).Item("AttributeValues"), "0")
    Else
      Response.Redirect("CPEoffer-con-attribute.aspx?OfferID=" & OfferID)
    End If
  Else
    AttributeValues = "0"
  End If
  
  MyCommon.QueryStr = "select RewardOptionID, TierLevels from CPE_RewardOptions with (NoLock) " & _
                      "where IncentiveID=" & OfferID & " and TouchResponse=0 and Deleted=0;"
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    roid = rst.Rows(0).Item("RewardOptionID")
    TierLevels = rst.Rows(0).Item("TierLevels")
  End If
  
  'Save routine
  If (Request.QueryString("save") <> "") Then
    'Store the existing locking value for use in newly-created records
    MyCommon.QueryStr = "select DisallowEdit from CPE_IncentiveDOW with (NoLock) where Deleted=0 and IncentiveID=" & OfferID & ";"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      Disallow_Edit = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit"), False)
    Else
      Disallow_Edit = False
    End If
    If IncentiveAttributeID = 0 Then
      'Create new condition
      MyCommon.QueryStr = "dbo.pa_CPE_AddAttribute"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@ROID", SqlDbType.Int, 4).Value = roid
      MyCommon.LRTsp.Parameters.Add("@DisallowEdit", SqlDbType.Bit, 1).Value = IIf(Disallow_Edit, 1, 0)
      MyCommon.LRTsp.Parameters.Add("@RequiredFromTemplate", SqlDbType.Bit, 1).Value = IIf(RequiredFromTemplate, 1, 0)
      MyCommon.LRTsp.Parameters.Add("@IncentiveAttributeID", SqlDbType.Int, 4).Direction = ParameterDirection.Output
      MyCommon.LRTsp.ExecuteNonQuery()
      IncentiveAttributeID = MyCommon.LRTsp.Parameters("@IncentiveAttributeID").Value
      MyCommon.Close_LRTsp()
      AttributeTypeID = MyCommon.Extract_Val(Request.QueryString("AttributeTypeID"))
      AttributeValues = Request.QueryString("selGroups")
      MyCommon.QueryStr = "dbo.pa_CPE_AddAttributeTiers"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@IncentiveAttributeID", SqlDbType.Int, 4).Value = IncentiveAttributeID
      MyCommon.LRTsp.Parameters.Add("@ROID", SqlDbType.Int, 4).Value = roid
      MyCommon.LRTsp.Parameters.Add("@AttributeTypeID", SqlDbType.Int, 4).Value = AttributeTypeID
      MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int, 4).Value = 1
      MyCommon.LRTsp.Parameters.Add("@AttributeValues", SqlDbType.NVarChar, 255).Value = AttributeValues
      MyCommon.LRTsp.ExecuteNonQuery()
      MyCommon.Close_LRTsp()
      historyString = "Added attribute condition"
    Else
      'Update existing condition
      AttributeValues = Request.QueryString("selGroups")
      MyCommon.QueryStr = "update CPE_IncentiveAttributeTiers set AttributeValues='" & AttributeValues & "' where IncentiveAttributeID=" & IncentiveAttributeID & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "update CPE_IncentiveAttributes set LastUpdate=getdate() where IncentiveAttributeID=" & IncentiveAttributeID & ";"
      MyCommon.LRT_Execute()
      historyString = "Edited attribute condition"
    End If
    MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
    MyCommon.LRT_Execute()
    MyCommon.Activity_Log(3, OfferID, AdminUserID, historyString)
    CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1") AndAlso infoMessage = ""
  End If
  
  'Load offer data
  MyCommon.QueryStr = "Select IsTemplate, FromTemplate from CPE_Incentives as CPE with (NoLock) " & _
                      "where IncentiveID=" & Request.QueryString("OfferID") & ";"
  rst = MyCommon.LRT_Select
  For Each row In rst.Rows
    isTemplate = MyCommon.NZ(row.Item("IsTemplate"), False)
    FromTemplate = MyCommon.NZ(row.Item("FromTemplate"), False)
  Next
  
  If (Request.QueryString("save") <> "" AndAlso Request.QueryString("IsTemplate") = "IsTemplate") Then
    'Update the status bits for the templates
    Dim form_Disallow_Edit As Integer = 0
    If (Request.QueryString("Disallow_Edit") = "on") Then
      form_Disallow_Edit = 1
    End If
    MyCommon.QueryStr = "update CPE_IncentiveAttributes with (RowLock) set DisallowEdit=" & form_Disallow_Edit & " " & _
                        "where RewardOptionID=" & roid & " and Deleted=0;"
    MyCommon.LRT_Execute()
  End If
  
  If (isTemplate Or FromTemplate) Then
    'Determine the permissions if it's a template
    MyCommon.QueryStr = "select DisallowEdit from CPE_IncentiveDOW with (NoLock) " & _
                        "where IncentiveID=" & OfferID & " and Deleted=0;"
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
  
  
  Send_HeadBegin("term.offer", "term.attributecondition", OfferID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
%>
<script type="text/javascript">
// This is the javascript array holding the function list
// The PrintJavascriptArray ASP function can be used to print this array.
<%
  MyCommon.QueryStr = "select AV.AttributeValueID, AV.ExtID, AV.Description from AttributeValues as AV with (NoLock) " & _
                      "where AV.AttributeTypeID=" & AttributeTypeID & " and Deleted=0 order by Description;"
  rst = MyCommon.LRT_Select
  If (rst.rows.count>0)
    Sendb("var functionlist = Array(")
    For Each row In rst.Rows
      Sendb("""" & MyCommon.NZ(row.item("Description"), "").ToString().Replace("""", "\""") & """,")
    Next
    Sendb(""""");")
    Sendb("var vallist = Array(")
    For Each row In rst.Rows
      Sendb("""" & MyCommon.NZ(row.item("AttributeValueID"), 0) & """,")
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
  
  // Clear the options list
  selectObj.length = 0;
  
  // Loop through the array and re-add matching options
  numShown = 0;
  for(i = 0; i < functionListLength; i++) {
    if(functionlist[i].search(re) != -1) {
      if (vallist[i] != "") {
        selectObj[numShown] = new Option(functionlist[i],vallist[i]);
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
  if (selectedValue != "") {
    selectedText = selectObj[document.getElementById("functionselect").selectedIndex].text;
  }
  
  selectboxObj = document.forms[0].selected;
  selectedboxValue = document.getElementById("selected").value;
  if (selectedboxValue != ""){
    selectedboxText = selectboxObj[document.getElementById("selected").selectedIndex].text;
  }
  
  if (itemSelected == "select1") {
    if (selectedValue != "") {
      // add items to selected box
      document.getElementById('deselect1').disabled=false;
      selectboxObj[selectboxObj.length] = new Option(selectedText,selectedValue);
      if (selectboxObj.length == 0) {
        // nothing in the select box so disable deselect
        document.getElementById('deselect1').disabled=true;
      }
    }
  }
  
  if (itemSelected == "deselect1") {
    if (selectedboxValue != "") {
      // remove items from selected box
      document.getElementById("selected").remove(document.getElementById("selected").selectedIndex)
      if (selectboxObj.length == 0) {
        // nothing in the select box so disable deselect
        document.getElementById('deselect1').disabled=true;
        document.getElementById('select1').disabled=false;      
      }
    }
  }
  // remove items from large list that are in the other lists
  removeUsed();
  updateButtons();
  return true;
}

function saveForm() {
  var funcSel = document.getElementById('functionselect');
  var elSel = document.getElementById('selected');
  var i,j;
  var selectList = "";
  var htmlContents = "";
  
  // assemble the list of values from the selected box
  for (i = elSel.length - 1; i>=0; i--) {
    if (elSel.options[i].value != "") {
      if (selectList != "") {
        selectList = selectList + ",";
      }
      selectList = selectList + elSel.options[i].value;
    }
  }
  
  // ok time to build up the hidden variables to pass for saving
  htmlContents = "<input type=\"hidden\" name=\"selGroups\" value=" + selectList + "> ";
  document.getElementById("hiddenVals").innerHTML = htmlContents;
  return true;
}

function validateEntry() {
  var retVal = true;
  var elemPP = document.getElementById("require_pp");
  var elem = document.getElementById("selected");   
  var qtyElem = document.getElementById("QtyForIncentive");
  var elemProgram = document.getElementById("ProgramID");
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
      if (qtyVal == "" || isNaN(qtyVal) || !isInteger(qtyVal)) {
        retVal = false;
        if (msg != '') { msg += '\n\r\n\r'; }
        msg += '* <% Sendb(Copient.PhraseLib.Lookup("CPEoffer-con-attribute.positiveinteger", LanguageID)) %>';
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

function updateButtons() {
  var elemSelect1 = document.getElementById('select1');
  var elemDeselect1 = document.getElementById('deselect1');
  var elemSave = document.getElementById('save');
  var elemSelected = document.forms[0].selected;
  
  if (elemSelected != null) {
    if (elemSelected.length == 0) {
      elemSelect1.disabled = false;
      elemDeselect1.disabled = true;
      if (document.getElementById('require_pp') != null) {
        if (document.getElementById('require_pp').checked == true) {
          elemSave.disabled = false;
        } else {
          elemSave.disabled = true;
        }
      } else {
        elemSave.disabled = true;
      }
    } else {
      if (document.getElementById('functionselect').length == 0) {
        elemSelect1.disabled = true;
      } else {
        elemSelect1.disabled = false;
      }
      elemDeselect1.disabled = false;
      elemSave.disabled = false;
    }
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
    if (document.getElementById("require_pp").checked == false) {
      document.getElementById('save').disabled=true;
    } else {
      document.getElementById('save').disabled=false;
    }
  }
  if (document.getElementById("require_pp").checked == true) {
    document.getElementById("Disallow_Edit").checked=false;
  }
}

function disableAll() {
  document.getElementById('select1').disabled=true;
  document.getElementById('deselect1').disabled=true;
  document.getElementById('functionselect').disabled=true;
  document.getElementById('selected').disabled=true;
}

function typeChange(AttributeTypeID) {
  var frm = document.mainform;
  frm.action = "CPEoffer-con-attribute.aspx?OfferID=<%Sendb(OfferID)%>&AttributeTypeID=" + AttributeTypeID;
  frm.submit();
}

function handleFilterRegEx(newFilter) {
  var frm = document.frmIter;
  var elemAdv = frm.advSql;
  var currentURL = window.location.href;
  var newURL = "";
  
  if (elemAdv != null && elemAdv.value !="") {
    if (currentURL.indexOf('filterOffer=') > -1) {
      newURL = currentURL.replace(/filterOffer=[0-9]?/g, 'filterOffer=' + newFilter);
      newURL = newURL.replace(/pagenum=[0-9]+/g, '');
    } else {
      if (currentURL.indexOf("&") > -1) {
        newURL = currentURL + "&amp;filterOffer=" + newFilter;
      } else {
        newURL = currentURL + "?filterOffer=" + newFilter;
      }
    }
    frm.action = newURL;
    frm.submit();
  } else { 
    if (document.getElementById("searchform") != null) { document.getElementById("searchform").submit(); }
  }    
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
  
  If (isTemplate) Then
    Send_BodyBegin(12)
  Else
    Send_BodyBegin(2)
  End If
  If (Logix.UserRoles.AccessOffers = False AndAlso Not isTemplate) Then
    Send_Denied(2, "perm.offers-access")
    GoTo done
  End If
  If (Logix.UserRoles.AccessTemplates = False AndAlso isTemplate) Then
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
<form action="#" id="mainform" name="mainform" onsubmit="return saveForm();">
  <span id="hiddenVals"></span>
  <input type="hidden" id="OfferID" name="OfferID" value="<% Sendb(OfferID) %>" />
  <input type="hidden" id="IncentiveAttributeID" name="IncentiveAttributeID" value="<% Sendb(IncentiveAttributeID) %>" />
  <input type="hidden" id="ROID" name="ROID" value="<% Sendb(roid) %>" />
  <input type="hidden" id="EngineID" name="EngineID" value="<% Sendb(EngineID) %>" />
  <input type="hidden" id="EngineSubTypeID" name="EngineSubTypeID" value="<% Sendb(EngineSubTypeID) %>" />
  <input type="hidden" id="IsTemplate" name="IsTemplate" value="<% Sendb(IIf(IsTemplate, "IsTemplate", "Not")) %>" />
  <div id="intro">
    <%
      If (isTemplate) Then
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.attributecondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
      Else
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.attributecondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
      End If
      Send("<div id=""controls"">")
      If IsTemplate Then
        Send("<span class=""temp"">")
        Send("  <input type=""checkbox"" class=""tempcheck"" id=""Disallow_Edit"" name=""Disallow_Edit""" & IIf(Disallow_Edit, " checked=""checked""", "") & " />")
        Send("  <label for=""Disallow_Edit"">" & Copient.PhraseLib.Lookup("term.locked", LanguageID) & "</label>")
        Send("</span>")
      End If
      If Not isTemplate Then
        If (Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Edit)) Then
          If IncentiveAttributeID = 0 Then
            Send_Save(" onclick=""this.style.visibility='hidden';""")
          Else
            Send_Save()
          End If
        End If
      Else
        If (Logix.UserRoles.EditTemplates) Then
          If IncentiveAttributeID = 0 Then
            Send_Save(" onclick=""this.style.visibility='hidden';""")
          Else
            Send_Save()
          End If
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
      <div class="box" id="type">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.attribute", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.type", LanguageID), VbStrConv.Lowercase))%>
          </span>
        </h2>
        <%
          If IncentiveAttributeID > 0 AndAlso AttributeTypeID > 0 Then
            Send("<h3>" & AttributeType & "</h3>")
            Send("<input type=""hidden"" id=""AttributeTypeID"" name=""AttributeTypeID"" value=""" & AttributeTypeID & """ />")
          Else
            Send("<select id=""AttributeTypeID"" name=""AttributeTypeID"" class=""longer"" onchange=""javascript:typeChange(this.value);"">")
            Send("  <option value=""0"">(Select an attribute type)</option>")
            MyCommon.QueryStr = "select AT.AttributeTypeID, AT.Description from AttributeTypes as AT with (NoLock) " & _
                                "inner join AttributeTypeEngines as ATE on ATE.AttributeTypeID=AT.AttributeTypeID " & _
                                "where AT.Deleted=0 and EngineID=" & EngineID & " and EngineSubTypeID=" & EngineSubTypeID & " " & _
                                "order by Description;"
            rst = MyCommon.LRT_Select
            If rst.Rows.Count > 0 Then
              For Each row In rst.Rows
                Send("  <option value=""" & MyCommon.NZ(row.Item("AttributeTypeID"), 0) & """" & IIf(MyCommon.NZ(row.Item("AttributeTypeID"), 0) = AttributeTypeID, " selected=""selected""", "") & ">" & MyCommon.NZ(row.Item("Description"), "") & "</option>")
              Next
            End If
            Send("</select>")
            Send("<br />")
          End If
          Send("<br class=""half"" />")
        %>
      </div>
      
      <div class="box" id="value"<% Sendb(IIf(AttributeTypeID=0, " style=""display:none;""", "")) %>>
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.attribute", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.values", LanguageID), VbStrConv.Lowercase))%>
          </span>
        </h2>
        <input type="radio" id="functionradio1" name="functionradio" checked="checked"<% Sendb(DisabledAttribute) %> /><label for="functionradio1"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
        <input type="radio" id="functionradio2" name="functionradio"<% Sendb(DisabledAttribute) %> /><label for="functionradio2"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
        <input type="text" class="medium" id="functioninput" name="functioninput" maxlength="100" onkeyup="handleKeyUp(200);" value=""<% Sendb(DisabledAttribute) %> /><br />
        <select class="longer" id="functionselect" name="functionselect" size="12"<% Sendb(DisabledAttribute) %>>
          <%
            MyCommon.QueryStr = "select AttributeValueID, ExtID, Description from AttributeValues with (NoLock) " & _
                                "where AttributeTypeID=" & AttributeTypeID & " and AttributeValueID not in (" & AttributeValues & ") and Deleted=0;"
            rst = MyCommon.LRT_Select
            If rst.Rows.Count > 0 Then
              For Each row In rst.Rows
                Send("<option value=" & MyCommon.NZ(row.Item("AttributeValueID"), 0) & ">" & MyCommon.NZ(row.Item("Description"), "") & "</option>")
              Next
            End If
          %>
        </select>
        <br />
        <br class="half" />
        <input type="button" class="regular select" id="select1" name="select1" value="&#9660; <% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID)) %>" onclick="handleSelectClick('select1');"<% Sendb(DisabledAttribute) %> />&nbsp;
        <input type="button" class="regular deselect" id="deselect1" name="deselect1" value="<% Sendb(Copient.PhraseLib.Lookup("term.deselect", LanguageID)) %> &#9650;" disabled="disabled" onclick="handleSelectClick('deselect1');"<% Sendb(DisabledAttribute) %> /><br />
        <br class="half" />
        <select class="longer" id="selected" name="selected" size="6"<% Sendb(DisabledAttribute) %>>
          <%
            If AttributeTypeID > 0 Then
              MyCommon.QueryStr = "select AttributeValueID, ExtID, Description from AttributeValues with (NoLock) " & _
                                  "where AttributeTypeID=" & AttributeTypeID & " and AttributeValueID in (" & AttributeValues & ") and Deleted=0;"
              rst = MyCommon.LRT_Select
              If rst.Rows.Count > 0 Then
                For Each row In rst.Rows
                  Send("<option value=""" & MyCommon.NZ(row.Item("AttributeValueID"), 0) & """>" & MyCommon.NZ(row.Item("Description"), "") & "</option>")
                Next
              End If
            End If
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
removeUsed();
updateButtons();
</script>
<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd("mainform", "functioninput")
  MyCommon = Nothing
  Logix = Nothing
%>
