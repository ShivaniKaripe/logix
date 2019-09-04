<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.Commoninc" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%
  ' *****************************************************************************
  ' * FILENAME: sources-list.aspx 
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
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "sources-list.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  PageNum = Request.QueryString("pagenum")
  If PageNum < 0 Then
    PageNum = 0
  End If
  MorePages = False
  
  Try
    Send_HeadBegin("term.externalsources")
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
    Send_Subtabs(Logix, 8, 4)
    
    If (Logix.UserRoles.AccessExternalSources = False) Then
      Send_Denied(1, "perm.access-external")
      GoTo done
    End If
    
    

    Dim SortText As String = "ExtCode"
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
        SortDirection = "DESC"
      End If
    Else
      SortDirection = Request.QueryString("SortDirection")
    End If
    
    sSearchQuery = "select ExtInterfaceID, ExtCode, Name, LastUpdate from ExtCRMInterfaces where editable = 1 and active=1 and Deleted=0 "
                  
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
      sSearchQuery = sSearchQuery & " and (ExtCode like N'%" & idSearchText & "%'"
      sSearchQuery = sSearchQuery & " or Name like N'%" & idSearchText & "%'"
      sSearchQuery = sSearchQuery & " or ExtInterfaceID like N'%" & idSearchText & "%')"
    End If
    MyCommon.QueryStr = sSearchQuery & " order by " & SortText & " " & SortDirection
    dst = MyCommon.LRT_Select
    sizeOfData = dst.Rows.Count
    i = linesPerPage * PageNum
    
    If (sizeOfData = 1 AndAlso Request.QueryString("searchterms") <> "") Then
      Response.Status = "301 Moved Permanently"
      Response.AddHeader("Location", "sources-edit.aspx?SourceID=" & dst.Rows(i).Item("ExtInterfaceID"))
    End If
%>
<div id="intro">
  <h1 id="title">
    <% Sendb(Copient.PhraseLib.Lookup("term.externalsources", LanguageID))%>
  </h1>
  <div id="controls">
    <form action="sources-edit.aspx" id="controlsform" name="controlsform">
    <%
      If (Logix.UserRoles.CreateExternalSource = True) Then
        Send_New()
      End If
      'If MyCommon.Fetch_SystemOption(75) Then
      '  If (Logix.UserRoles.AccessNotes) Then
      '    Send_NotesButton(33, 0, AdminUserID)
      '  End If
      'End If
    %>
    </form>
  </div>
</div>
<div id="main">
  <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
  <% Send_Listbar(linesPerPage, sizeOfData, PageNum, Request.QueryString("searchterms"), "&amp;SortText=" & SortText & "&amp;SortDirection=" & SortDirection)%>
  <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.externalsources", LanguageID)) %>">
    <thead>
      <tr>
        <th align="left" class="th-id" scope="col">
          <a href="sources-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=ExtInterfaceID&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>
          </a>
          <%
            If SortText = "ExtInterfaceID" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" class="th-code" scope="col">
          <a href="sources-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=ExtCode&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.code", LanguageID))%>
          </a>
          <%
            If SortText = "ExtCode" Then
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
          <a href="sources-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=Name&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>
          </a>
          <%
            If SortText = "Name" Then
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
          <a href="sources-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=LastUpdate&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.edited", LanguageID))%>
          </a>
          <%
            If SortText = "LastUpdate" Then
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
          Send("      <tr class=""" & Shaded & """>")
          Send("        <td>" & dst.Rows(i).Item("ExtInterfaceID") & "</td>")
          Send("        <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(dst.Rows(i).Item("ExtCode"), "&nbsp;"), 25) & "</td>")
          Send("        <td><a href=""sources-edit.aspx?SourceID=" & dst.Rows(i).Item("ExtInterfaceID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Name"), 30) & "</a></td>")
          If (Not IsDBNull(dst.Rows(i).Item("LastUpdate"))) Then
            Send("        <td>" & Logix.ToShortDateTimeString(dst.Rows(i).Item("LastUpdate"), MyCommon) & "</td>")
          Else
            Send("        <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
          End If
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
  '    Send_Notes(33, 0, AdminUserID)
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
