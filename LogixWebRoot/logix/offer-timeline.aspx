<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%    
  ' *****************************************************************************
  ' * FILENAME: offer-timeline.aspx 
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
  Dim Category As String = "0"
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim MonthNames(-1) As String
  Dim MonthList As String = ""
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "offer-timeline.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)

  ' Localize the month abbreviation to the users language
  If MyCommon.GetAdminUser.Culture IsNot Nothing Then
    MonthNames = MyCommon.GetAdminUser.Culture.DateTimeFormat.AbbreviatedMonthNames
    If MonthNames IsNot Nothing AndAlso MonthNames.Length > 0 Then
      For Each m As String In MonthNames
        If m <> "" Then
          If MonthList <> "" Then MonthList &= ","
          MonthList &= """" & m & """"
        End If
      Next
    End If
  End If
  If MonthList Is Nothing OrElse MonthList.Trim.Length = 0 Then
    MonthList = """Jan"",""Feb"",""Mar"",""Apr"",""May"",""Jun"",""Jul"",""Aug"",""Sep"",""Oct"",""Nov"",""Dec"""
  End If

  If (Request.QueryString("Category") <> "") Then
    Category = Request.QueryString("Category")
  End If
  
  Send_HeadBegin("term.offer", "term.timeline", , 1)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
%>
<script src="../javascript/api/timeline-api.js" type="text/javascript"></script>
<script type="text/javascript">
var tl;
function onLoad() {
    var eventSource = new Timeline.DefaultEventSource();
    var myMonths=new Array(<%Sendb(MonthList)%>);
    var currentTime = new Date()
    var month = currentTime.getMonth()
    var day = currentTime.getDate()
    var year = currentTime.getFullYear()
    var bandInfos = [
    Timeline.createBandInfo({
        trackHeight:    1.1,
        eventSource:    eventSource,
        date:           myMonths[month] + " " + day + " " + year + " 00:00:00 GMT",
        width:          "70%", 
        intervalUnit:   Timeline.DateTime.MONTH, 
        intervalPixels: 100
    }),
    Timeline.createBandInfo({
        showEventText:  false,
        trackHeight:    0.5,
        trackGap:       0.2,
        eventSource:    eventSource,
        date:           myMonths[month] + " " + day + " " + year + " 00:00:00 GMT",
        width:          "30%", 
        intervalUnit:   Timeline.DateTime.YEAR, 
        intervalPixels: 200
    })
  ];
  bandInfos[1].syncWith = 0;
  bandInfos[1].highlight = true;
  tl = Timeline.create(document.getElementById("timeline"), bandInfos);
  Timeline.loadXML("XMLFeeds.aspx?Category=<%sendb(Category) %>", function(xml, url) { eventSource.loadXML(xml, url); });
}

var resizeTimerID = null;
function onResize() {
    if (resizeTimerID == null) {
        resizeTimerID = window.setTimeout(function() {
            resizeTimerID = null;
            tl.layout();
        }, 500);
    }
}

function ChangeParentDocument(OfferID) { 
    opener.location = 'offer-sum.aspx?OfferID=' + OfferID;
    close();
   }
</script>
<%
  Send_Scripts()
  Send_HeadEnd()
%>
<body class="popup" onload="onLoad();" onresize="onResize();">
  <div id="main">
    <div id="timeline">
    </div>
<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd()
  Logix = Nothing
  MyCommon = Nothing
%>
