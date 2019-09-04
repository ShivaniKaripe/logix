<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: graph-display.aspx 
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
  Dim bParsed As Boolean
  Dim ReportStartDate As Date
  Dim ReportEndDate As Date
  Dim GraphType As Integer
  Dim OfferId As String
  Dim strURL As String
  Dim altText As String
  Dim Frequency As Integer
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "graph-display.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  bParsed = DateTime.TryParse(Request.QueryString("start"), ReportStartDate)
  If (Not bParsed) Then ReportStartDate = Now()
  bParsed = DateTime.TryParse(Request.QueryString("end"), ReportEndDate)
  If (Not bParsed) Then ReportEndDate = Now
    OfferId = Server.HtmlEncode(Request.QueryString("offerId"))
    bParsed = Integer.TryParse(Server.HtmlEncode(Request.QueryString("type")), GraphType)
    bParsed = Integer.TryParse(Server.HtmlEncode(Request.QueryString("freq")), Frequency)
  strURL = "reports-graph.aspx?offerId=" & OfferId & "&start=" & Logix.ToShortDateString(ReportStartDate, MyCommon) & _
           "&end=" & Logix.ToShortDateString(ReportEndDate, MyCommon) & "&type=" & GraphType & "&freq=" & Frequency
  altText = Copient.PhraseLib.Lookup("term.offer", LanguageID) & " " & OfferId & ", " & Logix.ToShortDateString(ReportStartDate, MyCommon) & _
            "-" & ReportEndDate.ToString("M/dd/yyyy")
  
  Send_HeadBegin("term.graph")
%>
<script type="text/javascript">
    function imageLoadFailed() {
        alert('<% Sendb(Copient.PhraseLib.Lookup("reports.errorcreating", LanguageID)) %>');
    }
</script>
<%
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
  Send_HeadEnd()
  Send_BodyBegin(3)
  
  If (Logix.UserRoles.AccessReports = False) Then
    Send_Denied(1, "perm.admin-reports")
    GoTo done
  End If
%>
<img src="<% Sendb(strURL) %>" alt="<% Sendb(altText) %>" title="<% Sendb(altText) %>" onerror="javascript:imageLoadFailed();" />
<%
done:
  Send_BodyEnd()
  MyCommon.Close_LogixRT()
  Logix = Nothing
  MyCommon = Nothing
%>
