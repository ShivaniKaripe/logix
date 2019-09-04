<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: graphic-list.aspx 
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
  Dim PrctSignPos As Integer
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "graphic-list.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  PageNum = Request.QueryString("pagenum")
  If PageNum < 0 Then PageNum = 0
  MorePages = False
  
  Send_HeadBegin("term.graphics")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
  Send_HeadEnd()
  Send_BodyBegin(1)
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  Send_Tabs(Logix, 6)
  Send_Subtabs(Logix, 60, 1)
  
  If (Logix.UserRoles.AccessGraphics = False) Then
    Send_Denied(1, "perm.graphics-access")
    GoTo done
  End If
  
  Dim SortText As String = "OnScreenAdID"
  Dim SortDirection As String = ""
  
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
  
  If (Request.QueryString("searchterms") <> "") Then
    If (Integer.TryParse(Request.QueryString("searchterms"), idNumber)) Then
      idSearch = idNumber.ToString
    Else
      idSearch = "-1"
    End If
    idSearchText = MyCommon.Parse_Quotes(Request.QueryString("searchterms"))
    PrctSignPos = idSearchText.IndexOf("%")
    If (PrctSignPos > -1) Then
      idSearch = -1
      idSearchText = idSearchText.Replace("%", "[%]")
    End If
    If (idSearchText.IndexOf("_") > -1) Then idSearchText = idSearchText.Replace("_", "[_]")
    MyCommon.QueryStr = "select Name, OnScreenAdID, LastUpload, Width, Height from OnScreenAds with (nolock) where Deleted=0 and (OnScreenAdID = " & idSearch & " or Name like N'%" & idSearchText & "%')  order by ISNULL(" & SortText & ", 0) " & SortDirection
  Else
    MyCommon.QueryStr = "select Name, OnScreenAdID, LastUpload, Width, Height from OnScreenAds with (nolock) where Deleted=0 order by ISNULL(" & SortText & ", 0) " & SortDirection
  End If
  
  dst = MyCommon.LRT_Select
  sizeOfData = dst.Rows.Count
  i = linesPerPage * PageNum
  
  If (sizeOfData = 1 AndAlso Request.QueryString("searchterms") <> "") Then
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "graphic-edit.aspx?OnScreenAdID=" & dst.Rows(0).Item("OnScreenAdID"))
  End If
%>
<div id="intro">
  <h1 id="title">
    <% Sendb(Copient.PhraseLib.Lookup("term.graphics", LanguageID))%>
  </h1>
  <div id="controls">
    <form action="graphic-edit.aspx" id="controlsform" name="controlsform">
      <%
        If (Logix.UserRoles.CreateGraphics) Then
          Send_New()
        End If
        'If MyCommon.Fetch_SystemOption(75) Then
        '  If (Logix.UserRoles.AccessNotes) Then
        '    Send_NotesButton(11, 0, AdminUserID)
        '  End If
        'End If
      %>
    </form>
  </div>
</div>
<div id="main">
  <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
  <% Send_Listbar(linesPerPage, sizeOfData, PageNum, Request.QueryString("searchterms"), "&SortText=" & SortText & "&SortDirection=" & SortDirection)%>
  <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.graphics", LanguageID)) %>">
    <thead>
      <tr>
        <th align="left" class="th-id" scope="col">
          <a href="graphic-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&SortText=OnScreenAdID&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>
          </a>
          <%If SortText = "OnScreenAdID" Then
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
          <a href="graphic-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&SortText=Name&SortDirection=<% Sendb(SortDirection) %>">
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
        <th align="left" class="th-width" scope="col">
          <a href="graphic-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&SortText=Width&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.width", LanguageID))%>
          </a>
          <%If SortText = "Width" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" class="th-height" scope="col">
          <a href="graphic-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&SortText=Height&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.height", LanguageID))%>
          </a>
          <%If SortText = "Height" Then
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
          <a href="graphic-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&SortText=LastUpload&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.lastupload", LanguageID))%>
          </a>
          <%If SortText = "LastUpload" Then
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
          Send("<tr class=""" & Shaded & """>")
          Send("  <td>" & dst.Rows(i).Item("OnScreenAdID") & "</td>")
          If (dst.Rows(i).Item("OnScreenAdID") >= 1) Then
            Send("  <td><a href=""graphic-edit.aspx?OnScreenAdID=" & dst.Rows(i).Item("OnScreenAdID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Name"), 40) & "</a></td>")
          Else
            Send("  <td>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Name"), 40) & "</td>")
          End If
          Send("  <td>" & MyCommon.NZ(dst.Rows(i).Item("Width"), "0") & "</td>")
          Send("  <td>" & MyCommon.NZ(dst.Rows(i).Item("Height"), "0") & "</td>")
          If (Not IsDBNull(dst.Rows(i).Item("LastUpload"))) Then
            Send("  <td>" & Logix.ToShortDateTimeString(dst.Rows(i).Item("LastUpload"), MyCommon) & "</td>")
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
      %>
    </tbody>
  </table>
</div>
<%
  'If MyCommon.Fetch_SystemOption(75) Then
  '  If (Logix.UserRoles.AccessNotes) Then
  '    Send_Notes(11, 0, AdminUserID)
  '  End If
  'End If
done:
  Send_BodyEnd("searchform", "searchterms")
  MyCommon.Close_LogixRT()
  Logix = Nothing
  MyCommon = Nothing
%>
