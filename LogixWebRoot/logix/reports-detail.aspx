<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: reports-detail.aspx 
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
%>
<script runat="server">
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc

</script>
<%  
  Dim CopientFileName As String = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
  Dim CopientFileVersion As String = "7.3.1.138972" 
  Dim CopientProject As String = "Copient Logix"
  Dim CopientNotes As String = ""
  
  Dim dt As DataTable
  Dim AdminUserID As Long
  Dim ReportStartDate As String = ""
  Dim ReportEndDate As String = ""
  Dim ProdStartDate As String = ""
  Dim ProdEndDate As String = ""
  Dim OfferID As Long = -1
  Dim EngineID As Long = -1
  Dim OfferName As String = ""
  Dim Status As String = ""
  Dim ShowReport As Boolean = False
  Dim bParsed As Boolean = False
  Dim dtStart As Date
  Dim dtEnd As Date
  Dim OfferLink As String = "#"
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim TempDate As Date
  Dim DefaultToEnhancedCustomReport As Integer = 0
  Dim HeightOffset = 0
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "reports-detail.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  If MyCommon.Fetch_SystemOption(274) = "1" Then
    DefaultToEnhancedCustomReport = 1
  End If
  
    OfferID = MyCommon.Extract_Val(Server.HtmlEncode(Request.QueryString("OfferID")))
  
    ReportStartDate = Server.HtmlEncode(Request.QueryString("Start"))
  If Date.TryParse(ReportStartDate, MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, TempDate) Then
    ReportStartDate = Logix.ToShortDateString(TempDate, MyCommon)
  End If
    ReportEndDate = Server.HtmlEncode(Request.QueryString("End"))
  If Date.TryParse(ReportEndDate, MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, TempDate) Then
    ReportEndDate = Logix.ToShortDateString(TempDate, MyCommon)
  End If

  Status = Server.HtmlEncode(Request.QueryString("Status"))
  
  If OfferID > 0 Then
    MyCommon.QueryStr = "select IncentiveName as Name from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID & _
                        " union " & _
                        "select Name from Offers with (NoLock) where OfferID=" & OfferID & ";"
    dt = MyCommon.LRT_Select
    If dt.Rows.Count > 0 Then
      OfferName = MyCommon.NZ(dt.Rows(0).Item("Name"), "")
    End If
  End If
  
  bParsed = DateTime.TryParse(ReportStartDate, MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, dtStart)
  If (bParsed) Then
    ProdStartDate = Logix.ToShortDateString(dtStart, MyCommon)
  Else
    ProdStartDate = StrConv(Copient.PhraseLib.Lookup("term.never", LanguageID), VbStrConv.Lowercase)
  End If
  
  bParsed = DateTime.TryParse(ReportEndDate, MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, dtEnd)
  If (bParsed) Then
    ProdEndDate = Logix.ToShortDateString(dtEnd, MyCommon)
  Else
    ProdEndDate = StrConv(Copient.PhraseLib.Lookup("term.never", LanguageID), VbStrConv.Lowercase)
  End If
  
  If (Request.Form("exportRpt") = "1") Then
    GenerateReport(Request.Form("Reports"))
    Response.End()
  End If
  
  Send_HeadBegin("term.offer", "term.report", OfferID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts(New String() {"datePicker.js"})
%>
<script type="text/javascript">
  var datePickerDivID = "datepicker";
  
  var highlightedRow = -1;
  var highlightedCol = -1;
  
  if (window.captureEvents){
    window.captureEvents(Event.CLICK);
    window.onclick=handlePageClick;
  } else {
    document.onclick=handlePageClick;
  }
  
  <% Send_Calendar_Overrides(MyCommon) %>

  function handlePageClick(e) {
    var calFrame = document.getElementById('calendariframe');
    var el=(typeof event!=='undefined')? event.srcElement : e.target        
    
    if (el != null) {
      var pickerDiv = document.getElementById(datePickerDivID);
      if (pickerDiv != null && pickerDiv.style.visibility == "visible") {
        if (el.id!="prod-start-picker" && el.id!="prod-end-picker") {
          if (!isDatePickerControl(el.className)) {
            pickerDiv.style.visibility = "hidden";
            pickerDiv.style.display = "none";
            if (calFrame != null) {
              calFrame.style.visibility = 'hidden';
              calFrame.style.display = "none";
            }
          }
        } else  {
          pickerDiv.style.visibility = "visible";            
          pickerDiv.style.display = "block";            
          if (calFrame != null) {
            calFrame.style.visibility = 'visible';
            calFrame.style.display = "block";
          }
        }
      }
    }
  }
  
  function isDatePickerControl(ctrlClass) {
    var retVal = false;
    
    if (ctrlClass != null && ctrlClass.length >= 2) {
      if (ctrlClass.substring(0,2) == "dp") {
        retVal = true;
      }
    }
    
    return retVal;
  }

  function handleSearch() {
    var rptStart = document.getElementById("reportstart");
    var rptEnd = document.getElementById("reportend");
    var elemfrequency = document.getElementById("frequency");
   
    if(elemfrequency.value == '1'){
      document.getElementById("row1").innerHTML = '(<%Sendb(Copient.PhraseLib.Lookup("term.weekly", LanguageID)) %>)';
      document.getElementById("row2").innerHTML = '(<%Sendb(Copient.PhraseLib.Lookup("term.weekly", LanguageID)) %>)';
      document.getElementById("row3").innerHTML = '(<%Sendb(Copient.PhraseLib.Lookup("term.weekly", LanguageID)) %>)';
      document.getElementById("row4").innerHTML = '(<%Sendb(Copient.PhraseLib.Lookup("term.weekly", LanguageID)) %>)';
      
    }else {
      document.getElementById("row1").innerHTML = '(<%Sendb(Copient.PhraseLib.Lookup("term.daily", LanguageID)) %>)';
      document.getElementById("row2").innerHTML = '(<%Sendb(Copient.PhraseLib.Lookup("term.daily", LanguageID)) %>)';
      document.getElementById("row3").innerHTML = '(<%Sendb(Copient.PhraseLib.Lookup("term.daily", LanguageID)) %>)';
      document.getElementById("row4").innerHTML = '(<%Sendb(Copient.PhraseLib.Lookup("term.daily", LanguageID)) %>)';
    }
             
    
    generateReport();
  }
    
  function generateReport() {
    xmlhttpPost("XMLFeeds.aspx")
  }
  
  function xmlhttpPost(strURL) {
    var xmlHttpReq = false;
    var self = this;
    
    document.getElementById("generate").disabled = true;
    document.getElementById("frequency").disabled = true;
    document.getElementById("reportstart").disabled = true;
    document.getElementById("reportend").disabled = true;
    document.getElementById("download").disabled = true;
    document.getElementById("nota").style.display = "none";
    document.getElementById("rptHeader").style.visibility = 'hidden';
    document.getElementById("reportTitle").style.visibility = "hidden";
    document.getElementById("report").style.visibility = "hidden";
    
    resetRow(highlightedRow);
    highlightedRow = -1;
    
    document.getElementById("wait").style.visibility = "visible";
    document.getElementById("wait").innerHTML = "<div class=\"loading\"><br \/><img id=\"clock\" src=\"../images/clock22.png\" \/><br \/>" + '<% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID)) %><\/div>';
    
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
    self.xmlHttpReq.onreadystatechange = function() {
      if (self.xmlHttpReq.readyState == 4) {
        updatepage(self.xmlHttpReq.responseText);
      }
    }
    self.xmlHttpReq.send(getPostData());
  }
  
  function getPostData() {
    var form = document.forms[1];
    var freq = form.frequency.value;
    var startDate = form.reportstart.value;
    var endDate = form.reportend.value;
    var offerId = form.offerID.value;
    
    qstr = 'Reports=1&frequency=' + escape(freq); // NOTE: no '?' before querystring
    qstr += "&reportstart=" + escape(startDate);
    qstr += "&reportend=" + escape(endDate);  
    qstr += "&lang=<% Sendb(LanguageID) %>";
    qstr += "&offerID=" + escape(offerId);
    
    return qstr;
  }
  
  function updatepage(str){
    if (str != null && str.indexOf("No Data Found") == -1) {
      document.getElementById("report").innerHTML = str;
      document.getElementById("nota").style.display = "none";
      document.getElementById("download").style.visibility = "visible";
      document.getElementById("reportTitle").style.visibility = "visible";
      document.getElementById("report").style.visibility = "visible";
      document.getElementById("rptHeader").style.visibility = "visible";
    } else {
      document.getElementById("nota").innerHTML = str; 
      document.getElementById("nota").style.display = "inline"
    }
    document.getElementById("generate").disabled = false;
    document.getElementById("frequency").disabled = false;
    document.getElementById("reportstart").disabled = false;
    document.getElementById("reportend").disabled = false;
    document.getElementById("download").disabled = false;
    document.getElementById("wait").style.visibility = "hidden";
  }
  
  function handleDownload() {
    var form = document.forms['generator'];
    
    form.exportRpt.value = "1";
    //form.action = "XMLFeeds.aspx";
    form.action = "reports-detail.aspx";
    form.method = "POST";
    form.submit();
  }
  
  function launchGraph(type) {
    var strURL = "graph-display.aspx?type=" + type;
    var form = document.forms[1];
    var freq = form.frequency.value;
    var startDate = form.reportstart.value;
    var endDate = form.reportend.value;
    var offerId = form.offerID.value;
    
    strURL += "&offerId=" + offerId + "&freq=" + freq + "&start=" + startDate + "&end=" + endDate;
    openReports(strURL);
  }
  
  function toggleHighlight(row) {
    var rowHdrElem = document.getElementById("rowHdr"+row);
    var rowBodyElem = document.getElementById("rowBody"+row);
    var bHighlight = false;
    var bSameRow = false;
    
    if (highlightedCol > -1) {
      toggleHighlight(highlightedCol);
      highlightedCol = -1;    
    }
    
    bSameRow = (row == highlightedRow);
    resetRow(highlightedRow);
    highlightedRow = -1;
    
    if (rowHdrElem != null && rowBodyElem != null && !bSameRow) {
      bHighlight = (rowHdrElem.className != "reportRowHeader rowHighlighted");
      if (bHighlight) {
        rowHdrElem.className = 'reportRowHeader rowHighlighted';
        rowBodyElem.className = 'rowHighlighted';
        highlightedRow = row;
      }
    }
  }
  
  function resetRow(row) {
    var rowHdrElem = document.getElementById("rowHdr"+row);
    var rowBodyElem = document.getElementById("rowBody"+row);
    
    if (rowHdrElem != null && rowBodyElem != null) {
      if (row % 2 == 1) {
        rowHdrElem.className = "reportRowHeader"
        rowBodyElem.className = "noclass"
      } else {
        rowHdrElem.className = "reportRowHeader shaded"
        rowBodyElem.className = "shaded"
      }
    }
  }
</script>
<%
  Send_HeadEnd()
  Send_BodyBegin(1)
  
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  Send_Tabs(Logix, 8)
  Send_Subtabs(Logix, 8, 8)
  
  If (Logix.UserRoles.AccessReports = False) Then
    Send_Denied(1, "perm.admin-reports")
    GoTo done
  End If
  
  'find out the which promotion engine for this offer 
  MyCommon.QueryStr = "select EngineID from OfferIDs with (NoLock) where OfferID=" & OfferID
  dt = MyCommon.LRT_Select
  If (dt.Rows.Count > 0) Then
    EngineID = MyCommon.NZ(dt.Rows(0).Item("EngineID"), -1)
  End If
  Select Case EngineID
    Case 0, 1
      OfferLink = "offer-sum.aspx"
    Case 2
      OfferLink = "CPEoffer-sum.aspx"
    Case 3
      OfferLink = "web-offer-sum.aspx"
    Case 5
      OfferLink = "email-offer-sum.aspx"
    Case 6
      OfferLink = "CAM/CAM-offer-sum.aspx"
    Case 9
      OfferLink = "UE/UEoffer-sum.aspx"
    Case Else
      OfferLink = "offer-sum.aspx"
  End Select
  
  HeightOffset = 0
%>
<div id="intro">
  <h1 id="title">
    <% Sendb(Copient.PhraseLib.Lookup("term.report", LanguageID) & ": " & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID)%>
  </h1>
  <div id="controls">
    <form action="#" id="controlsform" name="controlsform">
      <input type="button" class="regular" id="download" name="download" value="<% Sendb(Copient.PhraseLib.Lookup("term.download", LanguageID)) %>" style="visibility: hidden;" title="<% Sendb(Copient.PhraseLib.Lookup("reports.downloadnote", LanguageID)) %>" onclick="handleDownload();" />
      <%
        If MyCommon.Fetch_SystemOption(75) Then
          If (OfferID > 0 And Logix.UserRoles.AccessNotes) Then
            Send_NotesButton(18, OfferID, AdminUserID)
          End If
        End If
      %>
    </form>
  </div>
</div>
<div id="main">
  <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
  <div id="column">
   <div style="height:75px;">
    <b>
      <a href="<% Sendb(OfferLink) %>?OfferID=<% Sendb(OfferID) %>"><% Sendb(MyCommon.SplitNonSpacedString(OfferName, 50))%></a>
    </b>
    <br />
    <%Sendb(Copient.PhraseLib.Lookup("term.runsfrom", LanguageID) & " " & ProdStartDate & " " & Copient.PhraseLib.Lookup("term.to", LanguageID) & " " & ProdEndDate)%>
    <br />
    <%Sendb(Copient.PhraseLib.Lookup("term.status", LanguageID) & ": ")
      If (dtEnd < Today) Then
        Sendb(Copient.PhraseLib.Lookup("term.expired", LanguageID))
      Else
        Sendb(Status)
      End If
    %>
    <br />
    <br />
    </div>
    <form action="#" id="generator" name="generator">
      <label for="reportstart"><b><% Sendb(Copient.PhraseLib.Lookup("reports.generatereport", LanguageID))%>:</b></label><br />
      <br class="half" />
      <input type="text" class="short" id="reportstart" name="reportstart" maxlength="10" value="<% Sendb(ReportStartDate) %>" />
      <img src="../images/calendar.png" class="calendar" id="prod-start-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('reportstart', event);" />
      <% Sendb(Copient.PhraseLib.Lookup("term.to", LanguageID))%>
      <input type="text" class="short" id="reportend" name="reportend" maxlength="10" value="<% Sendb(ReportEndDate) %>" />
      <img src="../images/calendar.png" class="calendar" id="prod-end-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('reportend', event);" />
      <% Sendb(Copient.PhraseLib.Lookup("term.by", LanguageID))%>
      <select id="frequency" name="frequency" onchange="handleSearch();">
        <option value="1">
          <% Sendb(Copient.PhraseLib.Lookup("term.week", LanguageID))%>
        </option>
        <option value="2">
          <% Sendb(Copient.PhraseLib.Lookup("term.day", LanguageID))%>
        </option>
      </select>
      <input type="button" class="regular" id="generate" name="generate" value="<% Sendb(Copient.PhraseLib.Lookup("term.generate", LanguageID)) %>" onclick="handleSearch();" />
      <div id="datepicker" class="dpDiv">
      </div>
      <%
        If Request.Browser.Type = "IE6" Then
          Send("<iframe src=""javascript:'';"" id=""calendariframe"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no"" style=""visibility:hidden;display:none;""></iframe>")
        End If
      %>
      <input type="hidden" id="offerID" name="offerID" value="<% Sendb(OfferID) %>" />
      <input type="hidden" id="exportRpt" name="exportRpt" value="0" />
      <input type="hidden" id="Reports" name="Reports" value="1" />
      <input type="hidden" id="lang" name="lang" value="<% Sendb(LanguageID) %>" />
        

    </form>
    <hr class="hidden" />
    <br />
    <div id="nota" style="display: none;">
    </div>
    <div class="box" id="wait" style="visibility: hidden;">
    </div>
   
    <div class="box reportHeader" id="rptHeader" style="margin-top:30px;" >
      <h2>
        <span id="reportTitle" style="visibility: hidden;">
          <% Sendb(Copient.PhraseLib.Lookup("term.report", LanguageID))%>
        </span>
      </h2>
    
      <span id="rowHdr1" style="top: 24px;" class="reportRowHeader" ondblclick="javascript:toggleHighlight(1);">
       <a href="javascript:launchGraph(1);">
          <img src="../images/graph.png" alt="<% Sendb(Copient.PhraseLib.Lookup("reports.viewgraph", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("reports.viewgraph", LanguageID)) %>" onmouseover="this.src='../images/graph-on.png';" onmouseout="this.src='../images/graph.png';" />
        </a>&nbsp;&nbsp;&nbsp;<% Sendb(Copient.PhraseLib.Lookup("term.impressions", LanguageID))%>&nbsp;&nbsp;<span id="row1" ><% Sendb(Copient.PhraseLib.Lookup("term.daily", LanguageID))%>"/></span>
      </span>
      
      <span id="rowHdr2" style="top: 44px;" class="reportRowHeader shaded" ondblclick="javascript:toggleHighlight(2);">
        <a href="javascript:launchGraph(2);">
          <img src="../images/graph.png" alt="<% Sendb(Copient.PhraseLib.Lookup("reports.viewgraph", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("reports.viewgraph", LanguageID)) %>" onmouseover="this.src='../images/graph-on.png';" onmouseout="this.src='../images/graph.png';" />
        </a>&nbsp;&nbsp;&nbsp;<% Sendb(Copient.PhraseLib.Lookup("term.impressions", LanguageID) & " (" & Copient.PhraseLib.Lookup("term.cumulative", LanguageID) & ")")%>
      </span>
      <span id="rowHdr3" style="top: 64px;" class="reportRowHeader" ondblclick="javascript:toggleHighlight(3);">
        <a href="javascript:launchGraph(3);">
          <img src="../images/graph.png" alt="<% Sendb(Copient.PhraseLib.Lookup("reports.viewgraph", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("reports.viewgraph", LanguageID)) %>" onmouseover="this.src='../images/graph-on.png';" onmouseout="this.src='../images/graph.png';" />
        </a>&nbsp;&nbsp;&nbsp;<% Sendb(Copient.PhraseLib.Lookup("term.redemption", LanguageID))%>&nbsp;&nbsp;<span id="row2" ><% Sendb(Copient.PhraseLib.Lookup("term.daily", LanguageID))%>"/></span>
      </span>
      <span id="rowHdr4" style="top: 82px;" class="reportRowHeader shaded" ondblclick="javascript:toggleHighlight(4);">
        <a href="javascript:launchGraph(4);">
          <img src="../images/graph.png" alt="<% Sendb(Copient.PhraseLib.Lookup("reports.viewgraph", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("reports.viewgraph", LanguageID)) %>" onmouseover="this.src='../images/graph-on.png';" onmouseout="this.src='../images/graph.png';" />
        </a>&nbsp;&nbsp;&nbsp;<% Sendb(Copient.PhraseLib.Lookup("term.redemption", LanguageID) & " (" & Copient.PhraseLib.Lookup("term.cumulative", LanguageID) & ")")%>
      </span>
      
      <span id="rowHdr5" style="<% If (DefaultToEnhancedCustomReport=1) Then Send("top: 102px;") Else Send("display:none") %>" class="reportRowHeader" ondblclick="javascript:toggleHighlight(5);">
        <a href="javascript:launchGraph(5);">
          <img src="../images/graph.png" alt="<% Sendb(Copient.PhraseLib.Lookup("reports.viewgraph", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("reports.viewgraph", LanguageID)) %>" onmouseover="this.src='../images/graph-on.png';" onmouseout="this.src='../images/graph.png';" />
        </a>&nbsp;&nbsp;&nbsp;<% Sendb(Copient.PhraseLib.Lookup("term.transaction", LanguageID))%>&nbsp;&nbsp;<span id="row3" ><% Sendb(Copient.PhraseLib.Lookup("term.daily", LanguageID))%>"/></span>
      </span>
      <span id="rowHdr6" style="<% If (DefaultToEnhancedCustomReport=1) Then Send("top: 120px;") Else Send("display:none") %>" class="reportRowHeader shaded" ondblclick="javascript:toggleHighlight(6);">
        <a href="javascript:launchGraph(6);">
          <img src="../images/graph.png" alt="<% Sendb(Copient.PhraseLib.Lookup("reports.viewgraph", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("reports.viewgraph", LanguageID)) %>" onmouseover="this.src='../images/graph-on.png';" onmouseout="this.src='../images/graph.png';" />
        </a>&nbsp;&nbsp;&nbsp;<% Sendb(Copient.PhraseLib.Lookup("term.transaction", LanguageID) & " (" & Copient.PhraseLib.Lookup("term.cumulative", LanguageID) & ")")%>
      </span>

      <span id="rowHdr7" style="<% If (DefaultToEnhancedCustomReport=1) Then Send("top: 140px;") Else Send("top: 102px;") %>" class="reportRowHeader" ondblclick="javascript:toggleHighlight(7);">
        <a href="javascript:launchGraph(7);">
          <img src="../images/graph.png" alt="<% Sendb(Copient.PhraseLib.Lookup("reports.viewgraph", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("reports.viewgraph", LanguageID)) %>" onmouseover="this.src='../images/graph-on.png';" onmouseout="this.src='../images/graph.png';" />
        </a>&nbsp;&nbsp;&nbsp;<% Sendb(Copient.PhraseLib.Lookup("term.markdowns", LanguageID))%>&nbsp;&nbsp;<span id="row4" ><% Sendb(Copient.PhraseLib.Lookup("term.daily", LanguageID))%>"/></span>
      </span>
      <span id="rowHdr8" style="<% If (DefaultToEnhancedCustomReport=1) Then Send("top: 158px;") Else Send("top: 120px;") %>" class="reportRowHeader shaded" ondblclick="javascript:toggleHighlight(8);">
        <a href="javascript:launchGraph(8);">
          <img src="../images/graph.png" alt="<% Sendb(Copient.PhraseLib.Lookup("reports.viewgraph", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("reports.viewgraph", LanguageID)) %>" onmouseover="this.src='../images/graph-on.png';" onmouseout="this.src='../images/graph.png';" />
        </a>&nbsp;&nbsp;&nbsp;<% Sendb(Copient.PhraseLib.Lookup("term.markdowns", LanguageID) & " (" & Copient.PhraseLib.Lookup("term.cumulative", LanguageID) & ")")%>
      </span>
      <span id="rowHdr9" style="<% If (DefaultToEnhancedCustomReport=1) Then Send("top: 178px;") Else Send("top: 140px;") %>" class="reportRowHeader" ondblclick="javascript:toggleHighlight(9);">
        <a href="javascript:launchGraph(9);">
          <img src="../images/graph.png" alt="<% Sendb(Copient.PhraseLib.Lookup("reports.viewgraph", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("reports.viewgraph", LanguageID)) %>" onmouseover="this.src='../images/graph-on.png';" onmouseout="this.src='../images/graph.png';" />
        </a>&nbsp;&nbsp;&nbsp;<% Sendb(Copient.PhraseLib.Lookup("term.redemptionrate", LanguageID))%>
      </span>
    </div>
    <div class="box reportBody" id="report" style="margin-top:30px;margin-left:18px;width:520px;">
    </div>
    
  </div>
  <br clear="all" />
  <!-- <img src="reports-graph.aspx" alt="graph" /> -->
</div>

<script type="text/javascript">
<% Send_Date_Picker_Terms() %>
</script>

<script runat="server">
  Public DefaultLanguageID
  
  Sub GenerateReport(ByVal ReportID As String)
    Dim builder As StringBuilder = New StringBuilder()
    Dim bParsed As Boolean = False
    Dim LanguageID As Integer = 1
    Dim Frequency As String = 1
    
    If (Request.Form("lang") <> "") Then
      bParsed = Integer.TryParse(Request.Form("lang"), LanguageID)
      If (Not bParsed) Then LanguageID = 1
    End If
    
    If (Request.Form("frequency") <> "") Then
      Frequency = Request.Form("frequency")
    End If
    
    builder.Append(CreateReport(LanguageID))
    Response.Write(builder)
  End Sub
  
  Function CreateReport(ByVal LanguageID As String) As String
    Dim ReportStartDate As Date
    Dim ReportEndDate As Date
    Dim ReportWeeks As Integer
    Dim RowCount As Integer
    Dim CumulativeImpress As Integer
    Dim CumulativeRedeem As Integer
    Dim CumulativeAmtRedeem As Double
    Dim RedemptionRate As Double
    Dim AmtRedeem As Double
    Dim Redemptions As Integer
    Dim Impressions As Integer
    Dim i As Integer
    Dim OfferID As String = ""
    Dim dst As System.Data.DataTable
    Dim bParsed As Boolean
    Dim ExportRequested As String
    Dim builder As StringBuilder = New StringBuilder()
	Dim frequency AS String = ""
        
        
    
    bParsed = DateTime.TryParse(Request.Form("reportstart"), MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, ReportStartDate)
    If (Not bParsed) Then ReportStartDate = Now()
    bParsed = DateTime.TryParse(Request.Form("reportend"), MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, ReportEndDate)
    If (Not bParsed) Then ReportEndDate = Now
    ReportWeeks = DateTime.Compare(ReportStartDate, ReportEndDate) / 7
    OfferID = Request.Form("offerId")
    ExportRequested = Request.Form("exportRpt")
    
    MyCommon.Open_LogixWH()
    
    MyCommon.QueryStr = "select NumImpressions, NumRedemptions, NumTransactions, AmountRedeemed, ReportingDate from OfferReporting with (nolock) " & _
                        "where OfferID = " & OfferID & " " & _
                        "and ReportingDate between '" & ReportStartDate.ToString & "' and '" & ReportEndDate.ToString & "' " & _
                        "order by ReportingDate"
    dst = MyCommon.LWH_Select
        
    If (Request.Form("frequency") = "1") Then
		dst = RollupReportWeek(dst, ReportStartDate, ReportEndDate)
		frequency = "weekly"
    ElseIf (Request.Form("frequency") = "2") Then
		dst = FillInDays(dst, ReportStartDate, ReportEndDate)
		frequency = "daily"
    End If
    
    If (ExportRequested <> "") Then
      Response.AddHeader("Content-Disposition", "attachment; filename=Offer" & OfferID & "_Rpt.csv")
      Response.ContentType = "application/octet-stream"
      Return ExportReport(dst, frequency)
    End If
    
    Return builder.ToString
  End Function
  
  '----------------------------------------------------------------------------------
  
  Function RollupReportWeek(ByVal dst As DataTable, ByVal ReportStartDate As Date, ByVal ReportEndDate As Date) As DataTable
    Dim dstWeek As New DataTable
    Dim i, j As Integer
    Dim numRedeem As Integer
    Dim numImpression As Integer
    Dim amtRedeem As Double
    Dim numTransactions as Integer
    
    If (dst.Rows.Count > 0) Then
      Dim CurrentStart As Date
      Dim CurrentEnd As Date
      Dim ReportWeeks As Integer
      Dim row As DataRow
      Dim rowCt As Integer
      
      dstWeek = dst.Copy()
      dstWeek.Clear()
      
      CurrentStart = ReportStartDate
      CurrentEnd = ReportStartDate.AddDays(6)
      ReportWeeks = DateDiff(DateInterval.Day, ReportStartDate, ReportEndDate) / 7
      
      For i = 0 To ReportWeeks
        If (DateTime.Compare(ReportEndDate, CurrentStart) >= 0) Then
          dst.DefaultView.RowFilter = "ReportingDate >= '" & CurrentStart.ToString() & "' and ReportingDate <= '" & CurrentEnd.ToString() & "'"
          rowCt = dst.DefaultView.Count
          If (rowCt > 0) Then
            For j = 0 To rowCt - 1
              numRedeem += dst.DefaultView(j).Item("NumRedemptions")
              amtRedeem += dst.DefaultView(j).Item("AmountRedeemed")
              numImpression += dst.DefaultView(j).Item("NumImpressions")
              numTransactions += dst.DefaultView(j).Item("NumTransactions")
              If (j = dst.DefaultView.Count - 1) Then
                row = dst.DefaultView(j).Row
                row.Item("ReportingDate") = CurrentStart
                row.Item("NumRedemptions") = numRedeem
                row.Item("AmountRedeemed") = amtRedeem
                row.Item("NumImpressions") = numImpression
                row.Item("NumTransactions") = numTransactions
                dstWeek.ImportRow(row)
              End If
            Next
          Else
            row = dstWeek.NewRow()
            row.Item("ReportingDate") = CurrentStart
            row.Item("NumRedemptions") = 0
            row.Item("AmountRedeemed") = 0.0
            row.Item("NumImpressions") = 0
            row.Item("NumTransactions") = 0
            dstWeek.Rows.Add(row)
          End If
          numRedeem = 0
          amtRedeem = 0.0
          numImpression = 0
          numTransactions = 0
          CurrentStart = CurrentEnd.AddDays(1)
          CurrentEnd = CurrentStart.AddDays(6)
        End If
      Next
    End If
    
    Return dstWeek
  End Function
  
  '----------------------------------------------------------------------------------
  
  Function FillInDays(ByVal dst As DataTable, ByVal ReportStartDate As Date, ByVal ReportEndDate As Date) As DataTable
    Dim dstDay As New DataTable
    Dim CurrentDate As Date
    Dim RptDate As Date
    Dim row As DataRow
    
    dstDay = dst.Copy()
    dstDay.Clear()
    
    CurrentDate = ReportStartDate
    RptDate = ReportStartDate
    
    For Each row In dst.Rows
      RptDate = row.Item("ReportingDate")
      If (CurrentDate < RptDate) Then
        AddEmptyDays(dstDay, CurrentDate, RptDate)
        dstDay.ImportRow(row)
      Else
        dstDay.ImportRow(row)
      End If
      CurrentDate = RptDate.AddDays(1)
    Next
    
    If (ReportEndDate > RptDate) Then
      If (RptDate = ReportStartDate) Then
        AddEmptyDays(dstDay, RptDate, ReportEndDate.AddDays(1))
      Else
        AddEmptyDays(dstDay, RptDate.AddDays(1), ReportEndDate.AddDays(1))
      End If
    End If
    
    Return dstDay
  End Function
  
  '----------------------------------------------------------------------------------
  
  Sub AddEmptyDays(ByRef dst As DataTable, ByVal StartDate As Date, ByVal EndDate As Date)
    Dim CurrentDate As Date
    Dim row As DataRow
    CurrentDate = StartDate
    While (CurrentDate < EndDate)
      row = dst.NewRow()
      row.Item("ReportingDate") = CurrentDate
      row.Item("NumRedemptions") = 0
      row.Item("NumImpressions") = 0
      row.Item("NumTransactions") = 0
      dst.Rows.Add(row)
      CurrentDate = CurrentDate.AddDays(1)
    End While
  End Sub
  
  '----------------------------------------------------------------------------------
  
    Function ExportReport(ByVal dst As DataTable, ByVal frequency As String) As String
        Dim builder As StringBuilder = New StringBuilder()
    
        If (Not dst Is Nothing) Then
			builder.Append(",")
            builder.Append(WriteExportRow(dst, "ReportingDate", False))
            If (frequency.Contains("weekly")) Then
                builder.Append("Impressions (weekly),")
			Else
			    builder.Append("Impressions (daily),")
			End If
            builder.Append(WriteExportRow(dst, "NumImpressions", False))
            builder.Append("Impressions (cumulative),")
            builder.Append(WriteExportRow(dst, "NumImpressions", True))
			If (frequency.Contains("weekly")) Then
                builder.Append("Redemptions (weekly),")
			Else
			    builder.Append("Redemptions (daily),")
			End If
            builder.Append(WriteExportRow(dst, "NumRedemptions", False))
            builder.Append("Redemptions (cumulative),")
            builder.Append(WriteExportRow(dst, "NumRedemptions", True))

         If MyCommon.Fetch_SystemOption(274) = "1" Then
            If (frequency.Contains("weekly")) Then
               builder.Append("Transactions (weekly),")
            Else
               builder.Append("Transactions (daily),")
            End If
            builder.Append(WriteExportRow(dst, "NumTransactions", False))
            builder.Append("Transactions (cumulative),")
            builder.Append(WriteExportRow(dst, "NumTransactions", True))
         End If
            
			If (frequency.Contains("weekly")) Then
                builder.Append("Mark Downs ($) (weekly),")
			Else
			    builder.Append("Mark Downs ($) (daily),")
			End If
			builder.Append(WriteExportRow(dst, "AmountRedeemed", False))
            builder.Append("Mark Downs ($) (cumulative),")
            builder.Append(WriteExportRow(dst, "AmountRedeemed", True))
            builder.Append("Redemption Rate,")
            builder.Append(WriteRedemptionRow(dst))
        End If
    
        Return builder.ToString
    End Function
  
  '----------------------------------------------------------------------------------
  
  Function WriteExportRow(ByVal dst As DataTable, ByVal field As String, ByVal bCumulative As Boolean) As String
    Dim builder As StringBuilder = New StringBuilder()
    Dim RowCount, i As Integer
    Dim cumulative As Double
    Dim dt As Date
        
    RowCount = dst.Rows.Count
    
    For i = 0 To (RowCount - 1)
      If (field = "" OrElse IsDBNull(dst.Rows(i).Item(field))) Then
        builder.Append("0")
      Else
        If (bCumulative) Then
          cumulative += dst.Rows(i).Item(field)
          builder.Append(cumulative)
        Else
          If (IsDate(dst.Rows(i).Item(field))) Then
            dt = dst.Rows(i).Item(field)
            builder.Append(dt.ToString("M/dd/yyyy"))
          Else
            builder.Append(dst.Rows(i).Item(field))
          End If
        End If
      End If
      If (i = (RowCount - 1)) Then
        builder.Append(vbNewLine)
      Else
        builder.Append(",")
      End If
    Next
    Return builder.ToString()
  End Function
  
  '----------------------------------------------------------------------------------
  
  Function WriteRedemptionRow(ByVal dst As DataTable) As String
    Dim builder As StringBuilder = New StringBuilder()
    Dim RowCount, i As Integer
    Dim Impressions, Redemptions As Integer
    Dim RedemptionRate As Double
    
    RowCount = dst.Rows.Count
    
    For i = 0 To (RowCount - 1)
      Impressions = MyCommon.NZ(dst.Rows(i).Item("NumImpressions"), 0)
      Redemptions = MyCommon.NZ(dst.Rows(i).Item("NumRedemptions"), 0)
      If (Impressions > 0) Then
        RedemptionRate = Redemptions / Impressions
      Else
        RedemptionRate = 0.0
      End If
      builder.Append(RedemptionRate.ToString("0.####"))
      If (i = (RowCount - 1)) Then
        builder.Append(vbNewLine)
      Else
        builder.Append(",")
      End If
    Next
    
    Return builder.ToString()
  End Function
</script>

<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (OfferID > 0 And Logix.UserRoles.AccessNotes) Then
      Send_Notes(18, OfferID, AdminUserID)
    End If
  End If
done:
  Send_BodyEnd("generator", "reportstart")
  MyCommon.Close_LogixRT()
  Logix = Nothing
  MyCommon = Nothing
%>
