<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.Commoninc" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%
  ' *****************************************************************************
  ' * FILENAME: CRMoffer-list.aspx 
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
  Dim dt As System.Data.DataTable
  Dim idNumber As Integer
  Dim idSearch As String
  Dim idSearchText As String
  Dim sSearchQuery As String
  Dim PageNum As Integer
  Dim MorePages As Boolean
  Dim linesPerPage As Integer = 20
  Dim sizeOfData As Integer
  Dim i As Integer
  Dim PrctSignPos As Integer
  Dim Shaded As String = "shaded"
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim BannersEnabled As Boolean = False
  Dim FilterHealth As Integer = 0
  Dim DurationString As String = ""
  Dim DurationDate As DateTime = "1/1/1900 00:00:00"
  Dim DurationBase As DateTime = "1/1/1900 00:00:00"
  
  Dim SortText As String = "OfferID"
  Dim SortDirection As String
  Dim WhereClauseCpe As String = ""
  Dim WhereClauseCm As String = ""
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "CRMoffer-list.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  PageNum = Request.QueryString("pagenum")
  If PageNum < 0 Then
    PageNum = 0
  End If
  MorePages = False
  
  Try
    Send_HeadBegin("term.offerhealth")
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
    Send_Subtabs(Logix, 8, 6)
    
    If (Logix.UserRoles.EditSystemConfiguration = False) Then
      Send_Denied(1, "perm.admin-configuration")
      GoTo done
    End If
    
    
    
    BannersEnabled = (MyCommon.Fetch_SystemOption("66") = "1")
    FilterHealth = MyCommon.Extract_Val(Request.QueryString("filterhealth"))
  
    If FilterHealth <> 3 Then
      Response.Status = "301 Moved Permanently"
      Response.AddHeader("Location", "offer-health.aspx?&filterhealth=" & FilterHealth)
    End If
        
    If (Request.QueryString("SortText") <> "") Then
      SortText = Request.QueryString("SortText")
    End If
    
    If (Request.QueryString("pagenum") = "") Then
      If (Request.QueryString("SortDirection") = "ASC") Then
        SortDirection = "DESC"
      ElseIf (Request.QueryString("SortDirection") = "DESC") Then
        SortDirection = "ASC"
      Else
        SortDirection = "ASC"
      End If
    Else
      SortDirection = Request.QueryString("SortDirection")
    End If
    
    idSearchText = Request.QueryString("searchterms")
    If (idSearchText <> "") Then
      If (Integer.TryParse(Request.QueryString("searchterms"), idNumber)) Then
        idSearch = idNumber.ToString
      Else
        idSearch = "-1"
      End If
      PrctSignPos = idSearchText.IndexOf("%")
      If (PrctSignPos > -1) Then
        idSearch = "-1"
        idSearchText = idSearchText.Replace("%", "[%]")
      End If
      If (idSearchText.IndexOf("_") > -1) Then
        idSearchText = idSearchText.Replace("_", "[_]")
      End If
      idSearchText = MyCommon.Parse_Quotes(idSearchText)
      WhereClauseCpe = " AND (IncentiveID=" & MyCommon.Extract_Val(idSearchText) & " OR IncentiveName LIKE N'%" & idSearchText & "%') "
      WhereClauseCm = " AND (OfferID=" & MyCommon.Extract_Val(idSearchText) & " OR Name LIKE N'%" & idSearchText & "%') "
    End If
    sSearchQuery = "SELECT 0 AS PKID, IncentiveID AS OfferID, IncentiveName, '" & Copient.PhraseLib.Lookup("term.OutboundProcessing", LanguageID) & "' AS CRMStatus, LastCRMSendDate, (GETDATE()-LastCRMSendDate) AS Duration, Deleted " & _
                   "FROM CPE_Incentives AS I WITH (NoLock) " & _
                   "WHERE I.Deleted=0 And I.CRMSendToExport>0 " & WhereClauseCpe & _
                   " UNION " & _
                       "SELECT 0 AS PKID, OFR.OfferID, Name as IncentiveName, '" & Copient.PhraseLib.Lookup("term.OutboundProcessing", LanguageID) & "' AS CRMStatus, LastCRMSendDate, (GETDATE()-LastCRMSendDate) AS Duration, Deleted " & _
                   "FROM Offers AS OFR WITH (NoLock) " & _
                   "WHERE OFR.Deleted=0 And OFR.CRMSendToExport>0 " & WhereClauseCm & _
                   " UNION " & _
                   "SELECT EQ.PKID, EQ.OfferID, I.IncentiveName, '" & Copient.PhraseLib.Lookup("term.OutboundPrepared", LanguageID) & "' AS CRMStatus, EQ.LastUpdate, (GETDATE()-EQ.LastUpdate) AS Duration, EQ.Deleted " & _
                   "FROM CRMExportQueue AS EQ WITH (NoLock) " & _
                   "INNER JOIN CPE_Incentives AS I ON I.IncentiveID=EQ.OfferID " & _
                   "WHERE EQ.Deleted=0 " & WhereClauseCpe & _
                   " UNION " & _
                   "SELECT EQ.PKID, EQ.OfferID, OFR.Name as IncentiveName, '" & Copient.PhraseLib.Lookup("term.OutboundPrepared", LanguageID) & "' AS CRMStatus, EQ.LastUpdate, (GETDATE()-EQ.LastUpdate) AS Duration, EQ.Deleted " & _
                   "FROM CRMExportQueue AS EQ WITH (NoLock) " & _
                   "INNER JOIN Offers AS OFR ON OFR.OfferID=EQ.OfferID " & _
                       "WHERE EQ.Deleted=0 AND (EQ.OfferID=" & MyCommon.Extract_Val(idSearchText) & " OR Name LIKE N'%" & idSearchText & "%') " & _
                   " UNION " & _
                   "SELECT IQ.PKID, IQ.OfferID, I.IncentiveName, '" & Copient.PhraseLib.Lookup("term.InboundPending", LanguageID) & "' AS CRMStatus, IQ.LastUpdate, (GETDATE()-IQ.LastUpdate) AS Duration, IQ.Deleted " & _
                   "FROM CRMImportQueue AS IQ WITH (NoLock) " & _
                   "INNER JOIN CPE_Incentives AS I ON I.IncentiveID=IQ.OfferID " & _
                   "WHERE IQ.Deleted=0" & WhereClauseCpe & _
                   " UNION " & _
                   "SELECT IQ.PKID, IQ.OfferID, OFR.Name as IncentiveName, '" & Copient.PhraseLib.Lookup("term.InboundPending", LanguageID) & "' AS CRMStatus, IQ.LastUpdate, (GETDATE()-IQ.LastUpdate) AS Duration, IQ.Deleted " & _
                   "FROM CRMImportQueue AS IQ WITH (NoLock) " & _
                   "INNER JOIN Offers AS OFR ON OFR.OfferID=IQ.OfferID " & _
                       "WHERE IQ.Deleted=0 AND (IQ.OfferID=" & MyCommon.Extract_Val(idSearchText) & " OR Name LIKE N'%" & idSearchText & "%') "
    MyCommon.QueryStr = sSearchQuery & " ORDER BY " & SortText & " " & SortDirection
    dt = MyCommon.LRT_Select
    sizeOfData = dt.Rows.Count
    i = linesPerPage * PageNum
%>
<div id="intro">
  <h1 id="title">
    <% Sendb(Copient.PhraseLib.Lookup("term.offerhealth", LanguageID))%>
  </h1>
  <div id="controls">
    <form action="CRMoffer-list.aspx" id="controlsform" name="controlsform">
    </form>
  </div>
</div>
<div id="main">
  <%
    If (infoMessage <> "") Then
      Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
    End If
    Send_Listbar(linesPerPage, sizeOfData, PageNum, Request.QueryString("searchterms"), "&amp;SortText=" & SortText & "&amp;SortDirection=" & SortDirection)
  %>
  <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.crmoffers", LanguageID)) %>">
    <thead>
      <tr>
        <th align="left" class="th-id" scope="col">
          <a href="CRMoffer-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=OfferID&amp;SortDirection=<% Sendb(SortDirection) %>&amp;filterhealth=<% Sendb(FilterHealth) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>
          </a>
          <%
            If SortText = "OfferID" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" class="th-name" scope="col">
          <a href="CRMoffer-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=IncentiveName&amp;SortDirection=<% Sendb(SortDirection) %>&amp;filterhealth=<% Sendb(FilterHealth) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>
          </a>
          <%
            If SortText = "IncentiveName" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" class="th-name" scope="col">
          <a href="CRMoffer-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=CRMStatus&amp;SortDirection=<% Sendb(SortDirection) %>&amp;filterhealth=<% Sendb(FilterHealth) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.status", LanguageID))%>
          </a>
          <%
            If SortText = "CRMStatus" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" class="th-name" scope="col">
          <a href="CRMoffer-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=Duration&amp;SortDirection=<% Sendb(SortDirection) %>&amp;filterhealth=<% Sendb(FilterHealth) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.duration", LanguageID))%>
          </a>
          <%
            If SortText = "Duration" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
      </tr>
    </thead>
    <tbody>
      <%
        Shaded = "shaded"
        While (i < sizeOfData And i < linesPerPage + linesPerPage * PageNum)
          DurationString = ""
          Send("      <tr class=""" & Shaded & """>")
          Send("        <td>" & MyCommon.NZ(dt.Rows(i).Item("OfferID"), 0) & "</td>")
          If MyCommon.NZ(dt.Rows(i).Item("IncentiveName"), "") <> "" Then
            Send("        <td><a href=""offer-redirect.aspx?OfferID=" & MyCommon.NZ(dt.Rows(i).Item("OfferID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(dt.Rows(i).Item("IncentiveName"), "&nbsp;"), 30) & "</a></td>")
          Else
            Send("        <td>&nbsp;</td>")
          End If
          If Not IsDBNull(dt.Rows(i).Item("Duration")) Then
            DurationDate = dt.Rows(i).Item("Duration")
            DurationString &= DateDiff(DateInterval.Day, DurationBase, DurationDate) & " " & StrConv(Copient.PhraseLib.Lookup("term.days", LanguageID), VbStrConv.Lowercase) & ", "
            DurationDate = DurationDate.AddDays(-DateDiff(DateInterval.Day, DurationBase, DurationDate))
            DurationString &= DateDiff(DateInterval.Hour, DurationBase, DurationDate) & " " & StrConv(Copient.PhraseLib.Lookup("term.hours", LanguageID), VbStrConv.Lowercase) & ", "
            DurationDate = DurationDate.AddHours(-DateDiff(DateInterval.Hour, DurationBase, DurationDate))
            DurationString &= DateDiff(DateInterval.Minute, DurationBase, DurationDate) & " " & StrConv(Copient.PhraseLib.Lookup("term.minutes", LanguageID), VbStrConv.Lowercase)
          Else
            DurationString &= Copient.PhraseLib.Lookup("term.unknown", LanguageID)
          End If
          Send("        <td>" & MyCommon.NZ(dt.Rows(i).Item("CRMStatus"), "") & "</td>")
          Send("        <td>" & DurationString & "</td>")
          Send("      </tr>")
          If Shaded = "shaded" Then
            Shaded = ""
          Else
            Shaded = "shaded"
          End If
          i = i + 1
        End While
      %>
    </tbody>
  </table>
</div>
<%
done:
Finally
  MyCommon.Close_LogixRT()
  Logix = Nothing
  MyCommon = Nothing
End Try
Send_BodyEnd("searchform", "searchterms")
%>
