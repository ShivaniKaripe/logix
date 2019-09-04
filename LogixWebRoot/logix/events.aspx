<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: events.aspx 
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
  Dim dst As System.Data.DataTable
  Dim row As System.Data.DataRow
  Dim idNumber As Integer
  Dim idSearch As String
  Dim idSearchText As String
  Dim sSearchQuery As String
  Dim DayDifference As Integer = 0
  Dim YearFromNow As New DateTime
  Dim PageNum As Integer
  Dim MorePages As Boolean
  Dim linesPerPage As Integer = 20
  Dim sizeOfData As Integer
  Dim i As Integer = 0
  Dim PrctSignPos As Integer
  Dim Shaded As String = "shaded"
  Dim SortText As String = "EventDate"
  Dim SortDirection As String = "ASC"
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "events.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)

  PageNum = Request.QueryString("pagenum")
  If PageNum < 0 Then
    PageNum = 0
  End If
  MorePages = False
  
  If (Request.QueryString("SortText") <> "") Then
    SortText = Request.QueryString("SortText")
  End If
    
  If (Request.QueryString("pagenum") = "") Then
    If (Request.QueryString("SortDirection") = "DESC") Then
      SortDirection = "ASC"
    ElseIf (Request.QueryString("SortDirection") = "ASC") Then
      SortDirection = "DESC"
    Else
      SortDirection = "ASC"
    End If
  Else
    SortDirection = Request.QueryString("SortDirection")
  End If
  
  If (Request.QueryString("new") <> "") Then
    Send("<script type=""text/javascript"">openPopup('event-edit.aspx')</script>")
  End If
  
  YearFromNow = DateAdd(DateInterval.Year, 1, DateTime.Now)
  
  sSearchQuery = "select ItemType, LinkID, Name, Recurrence, FixedDate, EventDate from ( " & _
                 "select 'offerstart' as ItemType, OfferID as LinkID, Name, 0 as Recurrence, 1 as FixedDate, ProdStartDate as EventDate from Offers as O with (NoLock) where Deleted=0 and isnull(isTemplate,0)=0 and ProdStartDate>=getdate() " & _
                 "union " & _
                 "select 'offerend' as ItemType, OfferID as LinkID, Name, 0 as Recurrence, 1 as FixedDate, ProdEndDate as EventDate from Offers as O with (NoLock) where Deleted=0 and isnull(isTemplate,0)=0 and ProdEndDate>=getdate() " & _
                 "union " & _
                 "select 'offerstart' as ItemType, IncentiveID as LinkID, IncentiveName as Name, 0 as Recurrence, 1 as FixedDate, StartDate as EventDate from CPE_Incentives as I with (NoLock) where Deleted=0 and isnull(isTemplate,0)=0 and StartDate>=getdate() " & _
                 "union " & _
                 "select 'offerend' as ItemType, IncentiveID as LinkID, IncentiveName as Name, 0 as Recurrence, 1 as FixedDate, EndDate as EventDate from CPE_Incentives as I with (NoLock) where Deleted=0 and isnull(isTemplate,0)=0 and EndDate>=getdate() " & _
                 "union " & _
                 "select 'event' as ItemType, EventID as LinkID, Description as Name, Recurrence, FixedDate, EventDate from Events as E with (NoLock) where FixedDate=1 and Recurrence=0 and Deleted=0 and EventDate>=getdate() " & _
                 "union " & _
                 "select 'event' as ItemType, EventID as LinkID, Description as Name, Recurrence, FixedDate, dbo.NthDayInMonth(Ordinal,Day,Month,Year) as EventDate from Events as E with (NoLock) where FixedDate=0 and Recurrence=0 and Deleted=0 " & _
                 ") as EventsTable"
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
    If (idSearchText.IndexOf("_") > -1) Then idSearchText = idSearchText.Replace("_", "[_]")
    idSearchText = MyCommon.Parse_Quotes(idSearchText)
    sSearchQuery = sSearchQuery & " where Name like N'%" & idSearchText & "%'"
  End If
  MyCommon.QueryStr = sSearchQuery & " order by " & SortText & " " & SortDirection
  dst = MyCommon.LRT_Select
  sizeOfData = dst.Rows.Count
  i = linesPerPage * PageNum
  
  Send_HeadBegin("term.events")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
  Send_HeadEnd()
  Send_BodyBegin(1)
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  Send_Tabs(Logix, 1)
  Send_Subtabs(Logix, 1, 1)
%>
<div id="intro">
  <h1 id="title">
    <%Send(Copient.PhraseLib.Lookup("term.events", LanguageID)) %>
  </h1>
  <div id="controls">
  <!--
    <form action="#" id="controlsform" name="controlsform">
      <button id="new" name="new" class="regular" onclick="javascript:openPopup('event-edit.aspx')"><%Send(Copient.PhraseLib.Lookup("term.new", LanguageID)) %></button>
    </form>
  -->
  </div>
</div>
<div id="main">
  <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
  <% Send_Listbar(linesPerPage, sizeOfData, PageNum, Request.QueryString("searchterms"), "&SortText=" & SortText & "&SortDirection=" & SortDirection)%>
  <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.events", LanguageID)) %>">
    <thead>
      <tr>
        <th align="left" scope="col" style="width:130px;">
          <% Sendb(Copient.PhraseLib.Lookup("term.date", LanguageID))%>
        </th>
        <th align="left" scope="col">
          <% Sendb(Copient.PhraseLib.Lookup("term.event", LanguageID))%>
        </th>
        <th align="left" scope="col" style="width:70px;">
          Days away
        </th>
        <th align="left" scope="col" style="width:35px;">
          <% Sendb(Copient.PhraseLib.Lookup("term.day", LanguageID))%>
        </th>
        <th align="left" scope="col" style="width:35px;">
          W
        </th>
        <th align="left" scope="col" style="width:35px;">
          Q
        </th>
      </tr>
    </thead>
    <tbody>
      <%
        Shaded = "shaded"
        While (i < sizeOfData And i < linesPerPage + linesPerPage * PageNum)
          DayDifference = DateDiff(DateInterval.Day, DateTime.Today, dst.Rows(i).Item("EventDate"))
          Send("<tr class=""" & Shaded & """>")
          
          ' Date
          If (Not IsDBNull(dst.Rows(i).Item("EventDate"))) Then
            Send("  <td>" & Logix.ToShortDateString(dst.Rows(i).Item("EventDate"), MyCommon) & "</td>")
          Else
            Send("  <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
          End If
          
          ' Description
          Sendb("  <td>")
          If (dst.Rows(i).Item("ItemType") = "offerstart") Then
            Sendb("<span class=""greenlight"">&#9679;</span>")
          ElseIf (dst.Rows(i).Item("ItemType") = "offerend") Then
            Sendb("<span class=""redlight"">&#9679;</span>")
          End If
          
          If (MyCommon.NZ(dst.Rows(i).Item("ItemType"), "") = "offerstart") Or (MyCommon.NZ(dst.Rows(i).Item("ItemType"), "") = "offerend") Then
            Sendb(Copient.PhraseLib.Lookup("term.offer", LanguageID) & " ")
            Sendb("<a href=""offer-redirect.aspx?OfferID=" & MyCommon.NZ(dst.Rows(i).Item("LinkID"), 0) & """>")
            If (MyCommon.NZ(dst.Rows(i).Item("Name"), "") <> "") Then
              Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(dst.Rows(i).Item("Name"), ""), 25) & "</a>")
            Else
              Sendb(MyCommon.NZ(dst.Rows(i).Item("LinkID"), "") & "</a>")
            End If
          ElseIf (MyCommon.NZ(dst.Rows(i).Item("ItemType"), "") = "event") Then
            Sendb("<b><a href=""javascript:openPopup('event-edit.aspx?EventID=" & MyCommon.NZ(dst.Rows(i).Item("LinkID"), "") & "')"">")
            If (MyCommon.NZ(dst.Rows(i).Item("Name"), "") <> "") Then
              Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(dst.Rows(i).Item("Name"), ""), 25) & "</a></b>")
            Else
              Sendb(MyCommon.NZ(dst.Rows(i).Item("LinkID"), "") & "</a></b>")
            End If
          End If
          If (dst.Rows(i).Item("ItemType") = "offerstart") Then
            Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.starts", LanguageID), VbStrConv.Lowercase))
          ElseIf (dst.Rows(i).Item("ItemType") = "offerend") Then
            Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.ends", LanguageID), VbStrConv.Lowercase))
          Else
          End If
          Send("</td>")
          
          'Other rows
          Send("  <td>" & DayDifference & "</td>")
          Send("  <td>" & Format(dst.Rows(i).Item("EventDate"), "ddd") & "</td>")
          Send("  <td>" & DatePart(DateInterval.WeekOfYear, dst.Rows(i).Item("EventDate")) & "</td>")
          Send("  <td>" & DatePart(DateInterval.Quarter, dst.Rows(i).Item("EventDate")) & "</td>")
          Send("</tr>")
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
  Send_BodyEnd()
  MyCommon.Close_LogixRT()
  Logix = Nothing
  MyCommon = Nothing
%>
