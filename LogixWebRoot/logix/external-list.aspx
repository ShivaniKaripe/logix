<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: external.aspx 
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
  Dim AdminUserID As Long
  Dim dt As System.Data.DataTable
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "external.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  Send_HeadBegin("term.external")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
  Send_HeadEnd()
  Send_BodyBegin(1)
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  Send_Tabs(Logix, 8)
  Send_Subtabs(Logix, 8, 10)
  
  
  If (Logix.UserRoles.AccessReports = False) Then
    Send_Denied(1, "perm.admin-reports")
    GoTo done
  End If
%>
<script type="text/javascript">
  function launchCacheSettings() {
    var popW = 700;
    var popH = 570;
    var url = 'cachesettings.aspx';

    cacheWindow = window.open(url, "cacheTree", "width=" + popW + ", height=" + popH + ", top=" + calcTop(popH) + ", left=" + calcLeft(popW) + ", toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=yes");
    cacheWindow.focus();
  }
  function launchProdHierarchy() {
    var popW = 900;
    var popH = 570;
    var url = 'phierarchytree.aspx?ProductGroupID=-1';
    
    phierWindow = window.open(url,"pHierTree", "width="+popW+", height="+popH+", top="+calcTop(popH)+", left="+calcLeft(popW)+", toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=yes");
    phierWindow.focus();
  }
  function launchLocHierarchy() {
    var popW = 700;
    var popH = 570;
    var url = 'lhierarchytree.aspx?LocationGroupID=-1';
    
    lhierWindow = window.open(url,"lHierTree", "width="+popW+", height="+popH+", top="+calcTop(popH)+", left="+calcLeft(popW)+", toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=yes");
    lhierWindow.focus();
  }
</script>

<div id="intro">
  <h1 id="title">
    <% Sendb(Copient.PhraseLib.Lookup("term.external", LanguageID))%>
  </h1>
  <div id="controls">
  </div>
</div>
<div id="main">
  <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
  <p>
    
  </p>
  
  <div style="float: left; padding-right: 10px;">
   
    <form action="externalrewardsreports.aspx" id="externalrewardsreportsform" name="externalrewardsreportsform">
      <input type="submit" id="externalrewarsreport" name="externalrewarsreport" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.externalrewardsreports", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.externalrewardsreports", LanguageID)) %>"/>
    </form>
    &nbsp;
   
  </div>
  
  <div style="float: left; padding-right: 10px;">
    
  </div>
  
  <div style="float: left; padding-right: 10px;">
   
  </div>
  <div style="float: left; padding-right: 10px;">
    
  </div>
</div>
<%
done:
  Send_BodyEnd()
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  MyCommon = Nothing
  Logix = Nothing
%>
