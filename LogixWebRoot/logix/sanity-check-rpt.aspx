<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.CommonInc" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%
  ' *****************************************************************************
  ' * FILENAME: sanity-check-rpt.aspx.aspx 
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
  Dim Handheld As Boolean = False

  Dim rst As System.Data.DataTable
  Dim LocationID As Integer
  Dim SanityCheckPassed As Boolean = False
  Dim SanityCheckReport As String = ""
  Dim LastRptDate As Date = Nothing
  Dim LastRptDateDisplay As String = ""
  
  Response.Expires = 0
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  MyCommon.AppName = "sanity-check-rpt.aspx"
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If

  LocationID = MyCommon.Extract_Val(Request.QueryString("loc"))
  
  
  MyCommon.QueryStr = "select DBOK, LastReport, LastReportDate from SanityCheckStatus with (NoLock) where LocationID=" & LocationID
  rst = MyCommon.LRT_Select
  If (rst.Rows.Count > 0) Then
    SanityCheckPassed = MyCommon.NZ(rst.Rows(0).Item("DBOK"), False)
    SanityCheckReport = MyCommon.NZ(rst.Rows(0).Item("LastReport"), "")
    If (Not IsDBNull(rst.Rows(0).Item("LastReportDate"))) Then
      LastRptDate = rst.Rows(0).Item("LastReportDate")
      LastRptDateDisplay = Logix.ToShortDateTimeString(LastRptDate, MyCommon)
    Else
      LastRptDateDisplay = Copient.PhraseLib.Lookup("term.never", LanguageID)
    End If
  End If
  
  Send_HeadBegin("term.storehealth", , LocationID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
  Send_HeadEnd()
  Send_BodyBegin(3)

  
  %>
  <div id="intro">
    <h1 id="title">
      <% Send(Copient.PhraseLib.Lookup("term.sanitycheck", LanguageID))%>
    </h1>
    <div id="controls">
    </div>
  </div>
  <div id="main">  
    <div id="column">
      <div class="box" id="status">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.general", LanguageID))%>
          </span>
        </h2>
        <%
          Sendb(Copient.PhraseLib.Lookup("term.status", LanguageID) & ": ")
          Send(Copient.PhraseLib.Lookup(IIf(SanityCheckPassed, "term.passed", "term.failed"), LanguageID))
          Send("<br />")
          Send(Copient.PhraseLib.Lookup("term.lastsanitycheck", LanguageID) & ": " & LastRptDateDisplay)
        %>      
      </div>

      <div class="box" id="report" style="height:400px;">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.report", LanguageID))%>
          </span>
        </h2>
        <br class="half" />
        <div class="boxscroll" style="height:350px;width:97%;margin-left:7px;">
        <%
          If (SanityCheckReport <> "") Then
            SanityCheckReport = SanityCheckReport.Replace(vbCrLf, "<br />")
          End If
          Send(SanityCheckReport)
        %>
        </div>
      </div>
    </div>
  </div>
  <%
done:
  Send_BodyEnd()
  MyCommon.Close_LogixRT()
  MyCommon = Nothing
  Logix = Nothing
%>
