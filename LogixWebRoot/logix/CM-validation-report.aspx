<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: CM-validation-report.aspx 
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
  Dim rows() As DataRow = Nothing
  Dim type As String
  Dim id As Integer
  Dim level As Integer
  Dim iGraceHours As Integer
  Dim iGraceHoursWarn As Integer
  Dim OptionText As String
  Dim IntroText As String = ""
  Dim LocationName As String = ""
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "CM-validation-report.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  type = Request.QueryString("type")
  id = MyCommon.Extract_Val(Request.QueryString("id"))
  level = MyCommon.Extract_Val(Request.QueryString("level"))
  iGraceHours = MyCommon.Extract_Val(Request.QueryString("gh"))
  iGraceHoursWarn = MyCommon.Extract_Val(Request.QueryString("ghw"))
  
  Select Case type
    Case "pg"
      MyCommon.QueryStr = "pa_CM_ValidationReport_ProdGroup"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.Int).Value = id
      MyCommon.LRTsp.Parameters.Add("@GraceHours", SqlDbType.Int).Value = iGraceHours
      MyCommon.LRTsp.Parameters.Add("@GraceHoursWarn", SqlDbType.Int).Value = iGraceHoursWarn
      rst = MyCommon.LRTsp_select()
      MyCommon.Close_LRTsp()
      rows = rst.Select("Status=" & level, "ExtLocationCode")
      IntroText = Copient.PhraseLib.Lookup("term.productgroup", LanguageID) & " #" & id & ": "
    Case "cg"
      MyCommon.QueryStr = "pa_CM_ValidationReport_CustGroup"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@CustomerGroupID", SqlDbType.Int).Value = id
      MyCommon.LRTsp.Parameters.Add("@GraceHours", SqlDbType.Int).Value = iGraceHours
      MyCommon.LRTsp.Parameters.Add("@GraceHoursWarn", SqlDbType.Int).Value = iGraceHoursWarn
      rst = MyCommon.LRTsp_select()
      MyCommon.Close_LRTsp()
      rows = rst.Select("Status=" & level, "ExtLocationCode")
      IntroText = Copient.PhraseLib.Lookup("term.customergroup", LanguageID) & " #" & id & ": "
    Case "in"
      MyCommon.QueryStr = "pa_CM_ValidationReport_Offer"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.Int).Value = id
      MyCommon.LRTsp.Parameters.Add("@GraceHours", SqlDbType.Int).Value = iGraceHours
      MyCommon.LRTsp.Parameters.Add("@GraceHoursWarn", SqlDbType.Int).Value = iGraceHoursWarn
      rst = MyCommon.LRTsp_select()
      MyCommon.Close_LRTsp()
      rows = rst.Select("Status=" & level, "ExtLocationCode")
      IntroText = Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & id & ": "
  End Select
  
  If (level = 0) Then
    IntroText += " " & Copient.PhraseLib.Lookup("term.validatedlocations", LanguageID)
  ElseIf (level = 1) Then
    IntroText += " " & Copient.PhraseLib.Lookup("cgroup-edit.waitlocations", LanguageID)
  ElseIf (level = 2) Then
    IntroText += " " & Copient.PhraseLib.Lookup("term.watchlocations", LanguageID)
  ElseIf (level = 3) Then
    IntroText += " " & Copient.PhraseLib.Lookup("term.warninglocations", LanguageID)
  End If
  
  Send_HeadBegin("term.validationreport")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
%>

<script type="text/javascript" language="javascript">
// This is the javascript array holding the function list
// The PrintJavascriptArray ASP function can be used to print this array.
<%
  If (Not rows is Nothing AndAlso rows.length > 0)
    Sendb("var functionlist = Array(")
    For Each row In rows
      LocationName = MyCommon.NZ(row.Item("ExtLocationCode").ToString(), "*") & " - " & MyCommon.NZ(row.Item("LocationName").ToString(), "")
      If (LocationName.Length < 37) Then
        LocationName = LocationName.PadRight(37, " ")
      Else
        LocationName = LocationName.Substring(0, 33) & "... "
      End If
      OptionText =  row.Item("StatusMessage")
      If (OptionText = "") Then
        OptionText = LocationName & GetReasonText(row.Item("ReturnCode"))
      Else
        OptionText = LocationName & OptionText
      End If
      OptionText = OptionText.Replace("""", "\""")
      Sendb("""" & OptionText & """,")
    Next
    Send(""""");")
    Sendb("var vallist = Array(")
    For Each row In rows
      Sendb("""" & MyCommon.NZ(row.item("LocationID"), 0) & """,")
    Next
    Send(""""");")
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
    
    document.getElementById("functionselect").size = "20";
    
    // Set references to the form elements
    selectObj = document.forms[0].functionselect;
    textObj = document.forms[0].functioninput;
    
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
            if (vallist[i] != "") {
                selectObj[numShown] = new Option(functionlist[i],vallist[i]);
                selectObj[numShown].style.whiteSpace = 'pre';
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

function handleKeyDown(e) {
    var key = e.which ? e.which : e.keyCode;
    
    if (key == 40) {
        var elemSlct = document.getElementById("functionselect");
        if (elemSlct != null) { elemSlct.focus(); }
    }
}

function clearEntry() {
    var elemInput = document.getElementById("functioninput");
    
    if (elemInput != null) {
        elemInput.value = "";
        handleKeyUp(200);
        elemInput.focus();
    }
}
</script>

<%
  Send_Scripts()
  Send_HeadEnd()
  Send_BodyBegin(3)
  
  If (Logix.UserRoles.AccessOffers = False) Then
    Send_Denied(2, "perm.offers-access")
    GoTo done
  End If
%>
<form action="CM-validation-report.aspx" id="mainform" name="mainform">
  <div id="intro">
    <h1 id="title">
      <% Sendb(IntroText)%>
    </h1>
  </div>
  <div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <div id="column">
      <div class="box" id="selector">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.stores", LanguageID))%>
          </span>
        </h2>
        <input type="radio" id="functionradio1" name="functionradio" checked="checked" /><label for="functionradio1"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
        <input type="radio" id="functionradio2" name="functionradio" /><label for="functionradio2"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
        <input type="text" class="medium" onkeydown="handleKeyDown(event);" onkeyup="handleKeyUp(200);" id="functioninput" name="functioninput" value="" /><br />
        <br class="half" />
        <div style="overflow:auto;width:650px;">
          <select id="functionselect" name="functionselect" size="25" style="font-family:courier;min-width:640px;">
            <%
              If (Not rows Is Nothing) Then
                For Each row In rows
                  LocationName = MyCommon.NZ(row.Item("ExtLocationCode").ToString(), "*") & " - " & MyCommon.NZ(row.Item("LocationName").ToString(), "")
                  If (LocationName.Length < 37) Then
                    LocationName = LocationName.PadRight(37, " ")
                  Else
                    LocationName = LocationName.Substring(0, 33) & "... "
                  End If
                  OptionText = LocationName.Replace(" ", "&nbsp;")
                  Sendb("<option value=""" & MyCommon.NZ(row.Item("LocationID"), 0) & """>" & OptionText)
                  OptionText = row.Item("StatusMessage")
                  If OptionText = "" Then
                    Send(GetReasonText(row.Item("ReturnCode")) & "</option>")
                  Else
                    Send(OptionText & "</option>")
                  End If
                Next
              End If
            %>
          </select>
        </div>
        <br />
        <br class="half" />
        <hr class="hidden" />
      </div>
    </div>
  </div>
</form>

<script runat="server">
  Function GetReasonText(ByVal reasonCode As Integer) As String
    Dim reasonText As String = ""
    Select Case reasonCode
      Case 0
        reasonText = Copient.PhraseLib.Lookup("term.ok", LanguageID)
      Case 1
        reasonText = Copient.PhraseLib.Lookup("validationreport-waiting", LanguageID)
      Case 2
        reasonText = Copient.PhraseLib.Lookup("validationreport-notification", LanguageID)
      Case 3
        reasonText = Copient.PhraseLib.Lookup("validationreport-processed", LanguageID)
      Case 4
        reasonText = Copient.PhraseLib.Lookup("validationreport-recoverable", LanguageID)
      Case 5
        reasonText = Copient.PhraseLib.Lookup("validationreport-nonrecoverable", LanguageID)
      Case Else
        reasonText = Copient.PhraseLib.Lookup("term.unknown", LanguageID)
    End Select
    Return reasonText
  End Function
</script>

<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd("mainform", "functioninput")
  Logix = Nothing
  MyCommon = Nothing
%>
