<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: banner-list.aspx 
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
  Dim row As System.Data.DataRow
  Dim idNumber As Integer
  Dim idSearch As String
  Dim idSearchText As String
  Dim PageNum As Integer = 0
  Dim MorePages As Boolean
  Dim linesPerPage As Integer = 20
  Dim sizeOfData As Integer
  Dim i As Integer = 0
  Dim Shaded As String = "shaded"
  Dim SortText As String = "BannerID"
  Dim SortDirection As String = "DESC"
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "banner-list.aspx"
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
    If (Request.QueryString("SortDirection") = "ASC") Then
      SortDirection = "DESC"
    ElseIf (Request.QueryString("SortDirection") = "DESC") Then
      SortDirection = "ASC"
    Else
      SortDirection = "DESC"
    End If
  Else
    SortDirection = Request.QueryString("SortDirection")
  End If
  
  Send_HeadBegin("term.banners")
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
  Send_Subtabs(Logix, 8, 2)
  

  If (Logix.UserRoles.AccessBanners = False) Then
    Send_Denied(1, "perm.access-banners")
    GoTo done
  ElseIf (MyCommon.Fetch_SystemOption(66) <> "1") Then
    Send_Denied(1, "banners.disabled-note")
    GoTo done
  End If

  If (Request.QueryString("searchterms") <> "") Then
    If (Integer.TryParse(Request.QueryString("searchterms"), idNumber)) Then
      idSearch = idNumber.ToString
    Else
      idSearch = "-1"
    End If
    idSearchText = MyCommon.Parse_Quotes(Request.QueryString("searchterms"))
    MyCommon.QueryStr = "select BannerID, Name, CreatedDate, LastUpdate, AllBanners, DefaultBanner " & _
                        "from Banners with (NoLock) WHERE deleted=0 and (BannerID=" & idSearch & " or Name like '%" & idSearchText & "%')"
  Else
    MyCommon.QueryStr = "select BannerID, Name, CreatedDate, LastUpdate, AllBanners, DefaultBanner " & _
                        "from Banners with (NoLock) where deleted=0"
  End If
  MyCommon.QueryStr += " ORDER BY " & SortText & " " & SortDirection
  
  dst = MyCommon.LRT_Select
  sizeOfData = dst.Rows.Count
  
  ' set page variable(s)
  i = linesPerPage * PageNum
    
  If (sizeOfData = 1 AndAlso Request.QueryString("searchterms") <> "") Then
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "banner-edit.aspx?BannerID=" & MyCommon.NZ(dst.Rows(i).Item("BannerID"), 0))
  End If
%>
<div id="intro">
  <h1 id="title">
    <% Sendb(Copient.PhraseLib.Lookup("term.banners", LanguageID))%>
  </h1>
  <div id="controls">
    <form action="banner-edit.aspx" id="controlsform" name="controlsform">
      <%
        If (Logix.UserRoles.CreateBanners = True) Then
          Send_New()
        End If
        'If MyCommon.Fetch_SystemOption(75) Then
        '  If (Logix.UserRoles.AccessNotes) Then
        '    Send_NotesButton(20, 0, AdminUserID)
        '  End If
        'End If
      %>
    </form>
  </div>
</div>
<div id="main">
  <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
  <% Send_Listbar(linesPerPage, sizeOfData, PageNum, Request.QueryString("searchterms"), "&amp;SortText=" & SortText & "&amp;SortDirection=" & SortDirection)%>
  <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.users", LanguageID)) %>">
    <thead>
      <tr>
        <th align="left" class="th-id" scope="col">
          <a href="banner-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=BannerID&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>
          </a>
          <%If SortText = "BannerID" Then
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
          <a href="banner-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=Name&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>
          </a>
          <%If SortText = "Name" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" class="th-datetime" scope="col">
          <a href="banner-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=CreatedDate&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.created", LanguageID))%>
          </a>
          <%If SortText = "CreatedDate" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" class="th-datetime" scope="col">
          <a href="banner-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=LastUpdate&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.edited", LanguageID))%>
          </a>
          <%If SortText = "LastUpdate" Then
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
        While (i < sizeOfData And i < linesPerPage + linesPerPage * PageNum)
          row = dst.Rows(i)
          Send("<tr class=""" & Shaded & """>")
          Send("  <td>" & MyCommon.NZ(row.Item("BannerID"), 0) & "</td>")
          Send("  <td><a href=""banner-edit.aspx?BannerID=" & MyCommon.NZ(row.Item("BannerID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), "&nbsp;"), 25) & _
                  IIf(MyCommon.NZ(row.Item("AllBanners"), False), " (" & Copient.PhraseLib.Lookup("term.allbanners", LanguageID) & ")", "") & _
                  IIf(MyCommon.NZ(row.Item("DefaultBanner"), 0) = 1, " (" & Copient.PhraseLib.Lookup("term.defaultbanner", LanguageID) & ")", "") & "</a></td>")
          If (Not IsDBNull(row.Item("CreatedDate"))) Then
            Send("  <td>" & Logix.ToShortDateTimeString(row.Item("CreatedDate"), MyCommon) & "</td>")
          Else
            Send("  <td>" & Copient.PhraseLib.Lookup("term.never", LanguageID) & "</td>")
          End If
          If (Not IsDBNull(row.Item("LastUpdate"))) Then
            Send("  <td>" & Logix.ToShortDateTimeString(row.Item("LastUpdate"), MyCommon) & "</td>")
          Else
            Send("  <td>" & Copient.PhraseLib.Lookup("term.never", LanguageID) & "</td>")
          End If
          Send("</tr>")
          If Shaded = "shaded" Then
            Shaded = ""
          Else
            Shaded = "shaded"
          End If
          i = i + 1
        End While
        MyCommon.Close_LogixRT()
      %>
    </tbody>
  </table>
</div>
<%
  'If MyCommon.Fetch_SystemOption(75) Then
  '  If (Logix.UserRoles.AccessNotes) Then
  '    Send_Notes(20, 0, AdminUserID)
  '  End If
  'End If
done:
  Send_BodyEnd("searchform", "searchterms")
  MyCommon.Close_LogixRT()
  Logix = Nothing
  MyCommon = Nothing
%>
