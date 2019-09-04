<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:5.99.1.75910.Unstable Build - WUSBV185038-9SQ %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.CommonInc" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%
  ' *****************************************************************************
  ' * FILENAME: reason-list.aspx 
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
  ' * Version : 5.99.1.75910 
  ' *
  ' *****************************************************************************
  
  Dim CopientFileName As String = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
  Dim CopientFileVersion As String = "5.99.1.75910" 
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
  Dim BannersEnabled As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "reasons-list.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  PageNum = Request.QueryString("pagenum")
  If PageNum < 0 Then
    PageNum = 0
  End If
  MorePages = False
  
  Try
    Send_HeadBegin("term.reasoncodes")
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
    
    If (Logix.UserRoles.EditSystemConfiguration = False) Then
      Send_Denied(1, "perm.admin-configuration")
      GoTo done
    End If
    

    BannersEnabled = (MyCommon.Fetch_SystemOption("66") = "1")

    Dim SortText As String = "ReasonID"
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
    
    sSearchQuery = "select ReasonID, Description, LastUpdate from AdjustmentReasons with (NoLock) where UserDefined=1"
    idSearchText = Server.HtmlEncode(Request.QueryString("searchterms"))
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
      sSearchQuery = sSearchQuery & " and (ReasonID = " & idSearch & " "
      sSearchQuery = sSearchQuery & " or Description like N'%" & Server.HtmlDecode(idSearchText) & "%')"
    End If
    MyCommon.QueryStr = sSearchQuery & " order by " & SortText & " " & SortDirection
    dst = MyCommon.LXS_Select
    sizeOfData = dst.Rows.Count
    i = linesPerPage * PageNum
    
    If (sizeOfData = 1 AndAlso Request.QueryString("searchterms") <> "") Then
      Response.Status = "301 Moved Permanently"
      Response.AddHeader("Location", "reason-edit.aspx?ReasonID=" & dst.Rows(i).Item("ReasonID"))
    End If
%>
<div id="intro">
  <h1 id="title">
    <% Sendb(Copient.PhraseLib.Lookup("term.reasoncodes", LanguageID))%>
  </h1>
  <div id="controls">
    <form action="reason-edit.aspx" id="controlsform" name="controlsform">
      <%
        If (Logix.UserRoles.EditReasonCodes = True) Then
          Send_New()
        End If
        'If MyCommon.Fetch_SystemOption(75) Then
        '  If (Logix.UserRoles.AccessNotes) Then
        '    Send_NotesButton(26, 0, AdminUserID)
        '  End If
        'End If
      %>
    </form>
  </div>
</div>
<div id="main">
  <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
  <% Send_Listbar(linesPerPage, sizeOfData, PageNum, Server.HtmlEncode(Request.QueryString("searchterms")), "&SortText=" & SortText & "&SortDirection=" & SortDirection)%>
  <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.reason", LanguageID)) %>">
    <thead>
      <tr>
        <th align="left" class="th-id" scope="col">
          <a href="reasons-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&SortText=ReasonID&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>
          </a>
          <%If SortText = "ReasonID" Then
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
          <a href="reasons-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&SortText=Description&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.description", LanguageID))%>
          </a>
          <%If SortText = "Description" Then
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
          <a href="reasons-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&SortText=LastUpdate&SortDirection=<% Sendb(SortDirection) %>">
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
        Shaded = "shaded"
        While (i < sizeOfData AndAlso i < linesPerPage + linesPerPage * PageNum)
          Send("      <tr class=""" & Shaded & """>")
          Send("        <td>" & dst.Rows(i).Item("ReasonId") & "</td>")
              Send("        <td><a href=""reason-edit.aspx?ReasonID=" & dst.Rows(i).Item("ReasonID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 25) & "</a></td>")
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
  '    Send_Notes(26, 0, AdminUserID)
  '  End If
  'End If
done:
Finally
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  Logix = Nothing
  MyCommon = Nothing
End Try
Send_BodyEnd("searchform", "searchterms")
%>
