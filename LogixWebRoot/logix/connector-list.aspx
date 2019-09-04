<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: connector-list.aspx 
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
  Dim row As System.Data.DataRow
  Dim ShowAll As Boolean = False
  Dim idSearchText As String
  Dim sizeOfData As Integer
  Dim i As Integer = 0
  Dim SortText As String = "Name"
  Dim SortDirection As String = "ASC"
  Dim shaded As Boolean = True
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "connector-list.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  Send_HeadBegin("term.connectors")
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
  Send_Subtabs(Logix, 8, 3)
  
  If (Logix.UserRoles.AccessConnectors = False) Then
    Send_Denied(1, "perm.connectors-access")
    GoTo done
  End If
  
    If (Server.HtmlEncode(Request.QueryString("ShowAll")) <> "") Then
        ShowAll = True
    End If
  
    If (Server.HtmlEncode(Request.QueryString("SortText")) <> "") Then
        SortText = Request.QueryString("SortText")
    End If
  
    If (Server.HtmlEncode(Request.QueryString("SortDirection")) <> "") Then
        If (Server.HtmlEncode(Request.QueryString("SortDirection")) = "ASC") Then
            SortDirection = "DESC"
        ElseIf (Server.HtmlEncode(Request.QueryString("SortDirection")) = "DESC") Then
            SortDirection = "ASC"
        End If
    End If
  
    If (Server.HtmlEncode(Request.QueryString("action")) <> "") Then
        If (Server.HtmlEncode(Request.QueryString("action")) = "mute") Then
      
        End If
    End If
  
    If (Server.HtmlEncode(Request.QueryString("searchterms")) <> "") Then
        idSearchText = MyCommon.Parse_Quotes(Server.HtmlEncode(Request.QueryString("searchterms")))
        If (idSearchText.IndexOf("_") > -1) Then idSearchText = idSearchText.Replace("_", "[_]")
        MyCommon.QueryStr = "select C.ConnectorID, C.Name,C.NamePhraseID, C.DescriptionPhraseID, C.Path, C.UsesGUIDs, C.Installed, C.Visible, " & _
                        "(select count(GUID) from ConnectorGUIDs as CG where CG.ConnectorID=C.ConnectorID) as GUIDCount " & _
                        "from Connectors as C with (NoLock) " & _
                        "where C.ConnectorID in " & _
                        " (select ConnectorID from ConnectorEngines where EngineID in (select EngineID from PromoEngines where Installed=1)) " & _
                        "and C.Installed=1 and Name like N'%" & idSearchText & "%' "
    Else
        MyCommon.QueryStr = "select C.ConnectorID, C.Name,C.NamePhraseID, C.DescriptionPhraseID, C.Path, C.UsesGUIDs, C.Installed, C.Visible, " & _
                        "(select count(GUID) from ConnectorGUIDs as CG where CG.ConnectorID=C.ConnectorID) as GUIDCount " & _
                        "from Connectors as C with (NoLock) " & _
                        "where C.ConnectorID in " & _
                        " (select ConnectorID from ConnectorEngines where EngineID in (select EngineID from PromoEngines where Installed=1)) " & _
                        "and C.Installed=1 "
    End If
  If Not ShowAll Then
    MyCommon.QueryStr &= "and Visible=1 "
  End If
  MyCommon.QueryStr &= "order by " & SortText & " " & SortDirection & ";"
  dt = MyCommon.LRT_Select
  sizeOfData = dt.Rows.Count
%>
<div id="intro">
  <h1 id="title">
    <% Sendb(Copient.PhraseLib.Lookup("term.connectors", LanguageID))%>
  </h1>
  <div id="controls">
    <form action="<%Sendb("log-view.aspx?filetype=-1&amp;fileyear=" & Year(Today) & "&amp;filemonth=" & Month(Today) & "&amp;fileday=" & Day(Today)) %>" target="_blank" id="controlsform" name="controlsform" style="float: right;">
      <%
        If (Logix.UserRoles.AccessLogs = True) Then
          Send("<input type=""submit"" class=""regular"" id=""logs"" name=""logs"" value=""" & Copient.PhraseLib.Lookup("term.logs", LanguageID) & "..."" />")
        End If
      %>
    </form>
  </div>
</div>
<div id="main">
  <%
    If (infoMessage <> "") Then
      Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
    End If
  %>
  <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.connectors", LanguageID)) %>">
    <thead>
      <tr>
        <th align="left" class="th-id" scope="col">
          <a href="connector-list.aspx?searchterms=<%Sendb(Server.HtmlEncode(Request.QueryString("searchterms")))%>&amp;SortText=ConnectorID&amp;SortDirection=<% Sendb(SortDirection) %><% Sendb(IIf(ShowAll, "&amp;ShowAll=1", "")) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>
          </a>
          <%
            If SortText = "ConnectorID" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            End If
          %>
        </th>
        <th align="left" class="th-name" scope="col">
          <a href="connector-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=Name&amp;SortDirection=<% Sendb(SortDirection) %><% Sendb(IIf(ShowAll, "&amp;ShowAll=1", "")) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.connector", LanguageID))%>
          </a>
          <%
            If SortText = "Name" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            End If
          %>
        </th>
        <th align="left" class="th-name" scope="col">
          <a href="connector-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&amp;SortText=GUIDCount&amp;SortDirection=<% Sendb(SortDirection) %><% Sendb(IIf(ShowAll, "&amp;ShowAll=1", "")) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.guids", LanguageID))%>
          </a>
          <%
            If SortText = "GUIDCount" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            End If
          %>
        </th>
      </tr>
    </thead>
    <tbody>
      <%
        If (dt.Rows.Count > 0) Then
          For Each row In dt.Rows
            If (shaded) Then
              Send("<tr class=""shaded"">")
              shaded = False
            Else
              Send("<tr>")
              shaded = True
            End If
            Send("  <td" & IIf(MyCommon.NZ(row.Item("Visible"), False), "", " class=""grey""") & ">")
            Send("    " & MyCommon.NZ(row.Item("ConnectorID"), 0))
            Send("  </td>")
            Send("  <td>")
            Send("    <a href=""connector-detail.aspx?ConnectorID=" & MyCommon.NZ(row.Item("ConnectorID"), 0) & """" & IIf(Visible, "", " class=""disabled""") & ">" & Copient.phraseLib.Lookup(row.Item("NamePhraseID"),LanguageID) & "</a>")
            Send("  </td>")
            Send("  <td>")
            If MyCommon.NZ(row.Item("UsesGUIDs"), False) Then
              Send("    " & MyCommon.NZ(row.Item("GUIDCount"), 0))
            Else
              Send("    ")
            End If
            Send("  </td>")
            Send("</tr>")
          Next
        Else
          Send("<tr>")
          Send("  <td></td>")
          Send("  <td></td>")
          Send("  <td></td>")
          Send("</tr>")
        End If
      %>
    </tbody>
  </table>
</div>

<script runat="server">
</script>
<%
done:
  Send_BodyEnd()
  MyCommon.Close_LogixRT()
  MyCommon = Nothing
  Logix = Nothing
%>
