<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.Commoninc" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%
  ' *****************************************************************************
  ' * FILENAME: instantwin-list.aspx 
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
  Dim dst As System.Data.DataTable
  Dim dst2 As System.Data.DataTable
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
  Dim StartDate As New DateTime
  Dim EndDate As New DateTime
  Dim IsExpired As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "instantwin-list.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  
  PageNum = Request.QueryString("pagenum")
  If PageNum < 0 Then
    PageNum = 0
  End If
  MorePages = False
  
  Try
    Send_HeadBegin("term.instantwin", "term.report")
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
    Send_Subtabs(Logix, 8, 8)
    
    If (Logix.UserRoles.AccessInstantWinReports = False) Then
      Send_Denied(1, "perm.admin-instantwin")
      GoTo done
    End If
    
    
    BannersEnabled = (MyCommon.Fetch_SystemOption("66") = "1")
    
    Dim SortText As String = "IncentiveID"
    Dim SortDirection As String
    
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
    
    sSearchQuery = "select I.IncentiveID, I.IncentiveName, I.IsTemplate, I.StartDate, I.EndDate, RO.RewardOptionID, EIW.IncentiveEIWID, EIW.LastUpdate, " & _
                   "(select COUNT(TriggerID) from CPE_EIWTriggers as EIWT where EIWT.IncentiveEIWID=EIW.IncentiveEIWID and Original=1) as OriginalTriggers, " & _
                   "(select COUNT(TriggerID) from CPE_EIWTriggers as EIWT where EIWT.IncentiveEIWID=EIW.IncentiveEIWID and Original=0) as AddedTriggers, " & _
                   "(select COUNT(TriggerID) from CPE_EIWTriggers as EIWT where EIWT.IncentiveEIWID=EIW.IncentiveEIWID and Removed=1) as RemovedTriggers, " & _
                   "(select COUNT(TriggerID) from CPE_EIWTriggers as EIWT where EIWT.IncentiveEIWID=EIW.IncentiveEIWID) - (select COUNT(TriggerID) from CPE_EIWTriggers as EIWT where EIWT.IncentiveEIWID=EIW.IncentiveEIWID and Removed=1) as TotalTriggers, " & _
                   "(select COUNT(TriggerID) from CPE_EIWTriggersUsed as EIWTU where EIWTU.TriggerID in (select TriggerID from CPE_EIWTriggers as EIWT where EIWT.IncentiveEIWID=EIW.IncentiveEIWID)) as UsedTriggers, " & _
                   "(select COUNT(TriggerID) from CPE_EIWTriggers as EIWT where EIWT.IncentiveEIWID=EIW.IncentiveEIWID and Removed=0) - (select COUNT(TriggerID) from CPE_EIWTriggersUsed as EIWTU where EIWTU.TriggerID in (select TriggerID from CPE_EIWTriggers as EIWT where EIWT.IncentiveEIWID=EIW.IncentiveEIWID)) as AvailableTriggers " & _
                   "from CPE_Incentives as I with (NoLock) " & _
                   "inner join CPE_RewardOptions as RO on RO.IncentiveID=I.IncentiveID " & _
                   "inner join CPE_IncentiveEIW as EIW on EIW.RewardOptionID=RO.RewardOptionID " & _
                   "where I.IsTemplate=0 and I.Deleted=0 and RO.Deleted=0"
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
      sSearchQuery = sSearchQuery & " and I.IncentiveID=" & MyCommon.Extract_Val(idSearchText)
      sSearchQuery = sSearchQuery & " or I.IncentiveName like N'%" & idSearchText & "%'"
    End If
    MyCommon.QueryStr = sSearchQuery & " order by " & SortText & " " & SortDirection
    dst = MyCommon.LRT_Select
    sizeOfData = dst.Rows.Count
    i = linesPerPage * PageNum
    
    'If (sizeOfData = 1 AndAlso Request.QueryString("searchterms") <> "") Then
    '  Response.Status = "301 Moved Permanently"
    '  Response.AddHeader("Location", "instantwin-detail.aspx?IncentiveEIWID=" & dst.Rows(i).Item("IncentiveEIWID"))
    'End If
%>
<div id="intro">
  <h1 id="title">
    <% Sendb(Copient.PhraseLib.Lookup("term.instantwin", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.report", LanguageID), VbStrConv.Lowercase))%>
  </h1>
  <div id="controls">
    <form action="instantwin-list.aspx" id="controlsform" name="controlsform">
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
  <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.instantwin", LanguageID)) %>">
    <thead>
      <tr>
        <th align="left" class="th-id" scope="col" valign="bottom">
          <a href="instantwin-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=IncentiveID&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.offerid", LanguageID))%>
          </a>
          <%
            If SortText = "IncentiveID" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" class="th-name" scope="col" valign="bottom">
          <a href="instantwin-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=IncentiveName&amp;SortDirection=<% Sendb(SortDirection) %>">
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
        <th align="left" scope="col" style="width:70px;">
          <a href="instantwin-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=OriginalTriggers&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.original", LanguageID))%>
          </a>
          <%
            If SortText = "OriginalTriggers" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" scope="col" style="width:70px;">
          <a href="instantwin-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=AddedTriggers&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.added", LanguageID))%>
          </a>
          <%
            If SortText = "AddedTriggers" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" scope="col" style="width:70px;">
          <a href="instantwin-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=RemovedTriggers&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.removed", LanguageID))%>
          </a>
          <%
            If SortText = "RemovedTriggers" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" scope="col" style="width:70px;">
          <a href="instantwin-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=TotalTriggers&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.total", LanguageID))%>
          </a>
          <%
            If SortText = "TotalTriggers" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" scope="col" style="width:70px;">
          <a href="instantwin-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=UsedTriggers&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.used", LanguageID))%>
          </a>
          <%
            If SortText = "UsedTriggers" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" scope="col" style="width:70px;">
          <a href="instantwin-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=AvailableTriggers&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.available", LanguageID))%>
          </a>
          <%
            If SortText = "AvailableTriggers" Then
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
          StartDate = MyCommon.NZ(dst.Rows(i).Item("StartDate"), "1/1/1900 00:00:00")
          EndDate = MyCommon.NZ(dst.Rows(i).Item("EndDate"), "1/1/1900 00:00:00")
          EndDate = New Date(EndDate.Year, EndDate.Month, EndDate.Day, 23, 59, 59)
          If EndDate < Now.Date Then
            IsExpired = True
          Else
            IsExpired = False
          End If
          Send("      <tr class=""" & Shaded & """>")
          Send("        <td>" & MyCommon.NZ(dst.Rows(i).Item("IncentiveID"), 0) & "</td>")
          Send("        <td><a href=""offer-redirect.aspx?OfferID=" & MyCommon.NZ(dst.Rows(i).Item("IncentiveID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(dst.Rows(i).Item("IncentiveName"), "&nbsp;"), 30) & "</a>" & IIf(IsExpired, " <small>(" & Copient.PhraseLib.Lookup("term.expired", LanguageID) & ")</small>", "") & "</td>")
          Send("        <td>" & MyCommon.NZ(dst.Rows(i).Item("OriginalTriggers"), 0) & "</td>")
          Send("        <td>" & MyCommon.NZ(dst.Rows(i).Item("AddedTriggers"), 0) & "</td>")
          Send("        <td>" & MyCommon.NZ(dst.Rows(i).Item("RemovedTriggers"), 0) & "</td>")
          Send("        <td>" & MyCommon.NZ(dst.Rows(i).Item("TotalTriggers"), 0) & "</td>")
          Send("        <td>" & MyCommon.NZ(dst.Rows(i).Item("UsedTriggers"), 0) & "</td>")
          Send("        <td>" & MyCommon.NZ(dst.Rows(i).Item("AvailableTriggers"), 0) & "</td>")
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
  'If MyCommon.Fetch_SystemOption(75) Then
  '  If (Logix.UserRoles.AccessNotes) Then
  '    Send_Notes(22, 0, AdminUserID)
  '  End If
  'End If
done:
Finally
  MyCommon.Close_LogixRT()
  Logix = Nothing
  MyCommon = Nothing
End Try
Send_BodyEnd("searchform", "searchterms")
%>
