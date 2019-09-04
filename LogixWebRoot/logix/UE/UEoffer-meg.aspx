<%@ Page Language="vb" Debug="true" CodeFile="ue-cb.vb" Inherits="UECB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
    ' *****************************************************************************
    ' * FILENAME: UEoffer-meg.aspx 
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
    Dim dt As DataTable
    Dim row As DataRow
    Dim i As Integer = 0
    Dim OfferID As Long
    Dim OfferName As String
    Dim ClientOfferID As String
    Dim MEGSelectList As String = ""
    Dim IsTemplate As Boolean = False
    Dim FromTemplate As Boolean = False
    Dim CloseAfterSave As Boolean = False
    Dim Disallow_Edit As Boolean = True
    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False

    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If

    Response.Expires = 0
    MyCommon.AppName = "UEoffer-meg.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)

    'Fetch offer details
    OfferID = Request.QueryString("OfferID")

    'This would redirect to summary page if offer is awaiting deployment or if offer is awaiting recommendations.
    CheckIfValidOffer(MyCommon, OfferID)

    MyCommon.QueryStr = "Select IncentiveName, IsTemplate, FromTemplate, ClientOfferID from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID & ";"
    dt = MyCommon.LRT_Select
    If (dt.Rows.Count > 0) Then
        OfferName = MyCommon.NZ(dt.Rows(0).Item("IncentiveName"), "")
        IsTemplate = MyCommon.NZ(dt.Rows(0).Item("IsTemplate"), False)
        FromTemplate = MyCommon.NZ(dt.Rows(0).Item("FromTemplate"), False)
        ClientOfferID = MyCommon.NZ(dt.Rows(0).Item("ClientOfferID"), "")
    End If

    'Save
    If (Request.QueryString("meg-add") <> "") OrElse (Request.QueryString("meg-rem") <> "") Then
        If (Request.QueryString("meg-add") <> "") Then
            'Create MEGOffer records for selected items in "meg-available"
            If (Request.QueryString("meg-available") <> "") Then
                For i = 0 To Request.QueryString.GetValues("meg-available").GetUpperBound(0)
                    MyCommon.QueryStr = "insert into MutualExclusionGroupOffers with (RowLock) (MutualExclusionGroupID, OfferID, ClientOfferID) values (" & MyCommon.Extract_Val(Request.QueryString.GetValues("meg-available")(i)) & "," & OfferID & ",'" & ClientOfferID & "');"
                    MyCommon.LRT_Execute()
                Next i
            End If
        ElseIf (Request.QueryString("meg-rem") <> "") Then
            'Delete MEGOffer records for selected items in "meg-select"
            If (Request.QueryString("meg-select") <> "") Then
                For i = 0 To Request.QueryString.GetValues("meg-select").GetUpperBound(0)
                    MyCommon.QueryStr = "delete from MutualExclusionGroupOffers where OfferID=" & OfferID & " and MutualExclusionGroupID=" & MyCommon.Extract_Val(Request.QueryString.GetValues("meg-select")(i)) & ";"
                    MyCommon.LRT_Execute()
                Next i
            End If
        End If
        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1, " & _
                            "MutuallyExclusive=(select case when COUNT(MutualExclusionGroupID)>0 then 1 else 0 end from MutualExclusionGroupOffers where OfferID=" & OfferID & ")  " & _
                            " where IncentiveID=" & OfferID & ";"
        MyCommon.LRT_Execute()
        ResetOfferApprovalStatus(OfferID)
        If (Request.QueryString("meg-add") <> "") Then
            MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-megadd", LanguageID))
        ElseIf (Request.QueryString("meg-rem") <> "") Then
            MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-megremove", LanguageID))
        End If
        CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
    End If

    Send_HeadBegin("term.offer", "term.MutualExclusionGroups", OfferID)
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts()
%>
<script type="text/javascript">
  var isIE =(navigator.appName.indexOf("Microsoft")!=-1) ? true : false;
  var isOpera = (navigator.appName.indexOf("Opera")!=-1) ? true : false;
  var fullAvailList = null;
  
  // This is the javascript array holding the function list
  // The PrintJavascriptArray ASP function can be used to print this array.
  // Locations array
  <% 
    MyCommon.QueryStr = "select MutualExclusionGroupID, Name from MutualExclusionGroups with (NoLock) " & _
                        "where Deleted=0 and MutualExclusionGroupID not in " & _
                        "  (select MutualExclusionGroupID from MutualExclusionGroupOffers with (NoLock) where OfferID=" & OfferID & ") " & _
                        "order by Name;"
    dt = MyCommon.LRT_Select
    If (dt.rows.count > 0)
      Sendb("var functionlist = Array(")
      For Each row In dt.Rows
        Sendb("""" & MyCommon.NZ(row.item("Name"), "").ToString().Replace("""", "\""") & """,")
      Next
      Send(""""");")
      Sendb("  var vallist = Array(")
      For Each row In dt.Rows
        Sendb("""" & row.item("MutualExclusionGroupID") & """,")
      Next
      Send(""""");")
    Else
      Sendb("var functionlist = Array(")
      Send("""" & "" & """);")
      Sendb("  var vallist = Array(")
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
    var newOpt, optGp;
    
    document.getElementById("meg-available").size = "12";
    
    // Set references to the form elements
    selectObj = document.getElementById("meg-available");
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
    
    // Clear the options list (for IE, cloning the select box without its option and then replacing the 
    // existing one, is significantly faster than removing each option.
    if (textObj.value == '' && fullAvailList != null)  {
      document.getElementById('sgrouplist').replaceChild(fullAvailList, selectObj);
    } else {
      //var newSelectBox = selectObj.cloneNode(false);
      if (isIE) {
        var newSelectBox = selectObj.cloneNode(false);
      } else {
        var newSelectBox = document.createElement('select');
        newSelectBox.id = 'meg-available';
        newSelectBox.name = 'meg-available';
        newSelectBox.className = 'longest';
        newSelectBox.size = '12';
        newSelectBox.multiple = true;
      }
      
      document.getElementById('sgrouplist').replaceChild(newSelectBox, selectObj);
      selectObj = document.getElementById("meg-available");
     
      // Loop through the array and re-add matching options
      numShown = 0;
      for(i = 0; i < functionListLength; i++) {
        if(functionlist[i].search(re) != -1) {
          if (vallist[i] != "") {
            var newOpt = document.createElement('OPTION');
            newOpt.value = vallist[i];
            if (isIE) { newOpt.innerText = functionlist[i]}; 
            newOpt.text =  functionlist[i];
            selectObj[numShown] = new Option(newOpt.text, newOpt.value);
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
        try {
          selectObj.options[0].selected = true;
          enableEditButton(1);
        } catch (ex) {
          // ignore if unable to select (workaround for problem in IE 6)
        }
      }
    }
  }
  
  function handleKeyDown(e, slctName) {
    var key = e.which ? e.which : e.keyCode;
    
    if (key == 40) {
      var elemSlct = document.getElementById(slctName);
      if (elemSlct != null) { elemSlct.focus(); }
    }
  }
</script>
<%
    Send("<script type=""text/javascript"">")
    Send("function ChangeParentDocument() { ")
    Send("  if (opener != null) {")
    Send("    var newlocation = 'UEoffer-gen.aspx?OfferID=" & OfferID & "'; ")
    Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
    Send("  opener.location = 'UEoffer-gen.aspx?OfferID=" & OfferID & "'; ")
    Send("  }")
    Send("  }")
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
%>
<form action="UEoffer-meg.aspx" id="mainform" name="mainform">
  <div id="intro">
    <input type="hidden" id="OfferID" name="OfferID" value="<% Sendb(OfferID)%>" />
    <%
      Sendb("<h1 id=""title"">")
      Sendb(Copient.PhraseLib.Lookup(IIf(IsTemplate, "term.template", "term.offer"), LanguageID))
      Sendb(" #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.MutualExclusionGroups", LanguageID), VbStrConv.Lowercase))
      Send("</h1>")
    %>
    <div id="controls">
      <%
        'If Logix.UserRoles.EditOffer Then
        '  Send_Save()
        'End If
      %>
    </div>
  </div>
  <div id="main">
    <%
      If (infoMessage <> "") Then
        Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
      End If
    %>
    <div id="column">
      <div class="box" id="MEGselector">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.groups", LanguageID))%>
          </span>
        </h2>
        <label for="meg-available"><b><% Sendb(Copient.PhraseLib.Lookup("term.available", LanguageID) & ":")%></b></label>
        <br clear="all" />
        <input type="radio" id="functionradio1" name="functionradio" checked="checked" /><label for="functionradio1"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
        <input type="radio" id="functionradio2" name="functionradio" /><label for="functionradio2"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
        <input class="longer" onkeydown="handleKeyDown(event, 'meg-available');" onkeyup="handleKeyUp(200);" id="functioninput" name="functioninput" maxlength="100" type="text" value="" /><br />
        <br class="half" />
        <%
          'AVAILABLE GROUPS
          Send("<span id=""sgrouplist"">")
          Send("<select class=""longest"" multiple=""multiple"" id=""meg-available"" name=""meg-available"" size=""12"">")
          MyCommon.QueryStr = "select MutualExclusionGroupID, Name from MutualExclusionGroups with (NoLock) " & _
                              "where Deleted=0 and MutualExclusionGroupID not in " & _
                              "  (select MutualExclusionGroupID from MutualExclusionGroupOffers with (NoLock) where OfferID=" & OfferID & ") " & _
                              "order by Name;"
          dt = MyCommon.LRT_Select
          If (dt.Rows.Count > 0) Then
            For Each row In dt.Rows
              Send("<option value=""" & row.Item("MutualExclusionGroupID") & """>" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
            Next
          End If
          Send("</select>")
          Send("</span>")
          Send("<br />")
          Send("<br class=""half"" />")
          
          'BUTTONS
          Sendb("<input type=""submit"" class=""regular select"" id=""meg-add"" name=""meg-add"" title=""" & Copient.PhraseLib.Lookup("term.select", LanguageID) & """ value=""&#9660; " & Copient.PhraseLib.Lookup("term.select", LanguageID) & """ />")
          Sendb("<input type=""submit"" class=""regular deselect"" id=""meg-rem"" name=""meg-rem"" title=""" & Copient.PhraseLib.Lookup("term.unselect", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.deselect", LanguageID) & " &#9650;"" />")
          Send("<br />")
          
          'SELECTED GROUPS
          MyCommon.QueryStr = "select MutualExclusionGroupID, Name from MutualExclusionGroups with (NoLock) " & _
                              "where Deleted=0 and MutualExclusionGroupID in " & _
                              "  (select MutualExclusionGroupID from MutualExclusionGroupOffers with (NoLock) where OfferID=" & OfferID & ") " & _
                              "order by Name;"
          dt = MyCommon.LRT_Select
          Send("<label for=""meg-select""><b>" & Copient.PhraseLib.Lookup("term.selected", LanguageID) & ":</b></label>")
          Send("<br />")
          Send("<select class=""longest"" multiple=""multiple"" id=""meg-select"" name=""meg-select"" size=""4"">")
          If (dt.Rows.Count > 0) Then
            For Each row In dt.Rows
              Send("<option value=""" & row.Item("MutualExclusionGroupID") & """>" & row.Item("Name") & "</option>")
            Next
          End If
          Send("</select>")
        %>
        <hr class="hidden" />
      </div>
    </div>
 
  </div>
</form>

<script type="text/javascript">
handleKeyUp(9999, 4);
</script>
<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd("mainform", "functioninput")
  Logix = Nothing
  MyCommon = Nothing
%>