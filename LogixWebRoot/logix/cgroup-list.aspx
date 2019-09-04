<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.CommonInc" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%
    ' *****************************************************************************
    ' * FILENAME: cgroup-list.aspx 
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
    CMS.AMS.CurrentRequest.Resolver.AppName = "cgroup-list.aspx"
    Dim MyCommon As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    Dim dst As System.Data.DataTable
    Dim Shaded As String = "shaded"
    Dim idNumber As Integer
    Dim idSearch As String = "-1"
    Dim idSearchText As String = ""
    Dim PageNum As Integer = 0
    Dim MorePages As Boolean
    Dim linesPerPage As Integer = 20
    Dim sizeOfData As Integer
    Dim i As Integer = 0
    Dim PrctSignPos As Integer
    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False
    Dim analyticsCGService As CMS.AMS.Contract.IAnalyticsCustomerGroups = CMS.AMS.CurrentRequest.Resolver.Resolve(Of CMS.AMS.Contract.IAnalyticsCustomerGroups)()

    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If

    Response.Expires = 0
    MyCommon.AppName = "cgroup-list.aspx"
    MyCommon.Open_LogixRT()


    AdminUserID = Verify_AdminUser(MyCommon, Logix)

    PageNum = Request.QueryString("pagenum")
    If PageNum < 0 Then PageNum = 0
    MorePages = False

    Send_HeadBegin("term.customergroups")
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts()
    Send_HeadEnd()
    Send_BodyBegin(1)
    Send_Bar(Handheld)
    Send_Help(CopientFileName)
    Send_Logos()
    Send_Tabs(Logix, 3)
    Send_Subtabs(Logix, 30, 1)

    If (Logix.UserRoles.AccessCustomerGroups = False) Then
        Send_Denied(1, "perm.cgroup-access")
        GoTo done
    End If



    Dim SortText As String = "CustomerGroupID"
    Dim SortDirection As String = "DESC"

    If (Request.QueryString("SortText") <> "" AndAlso Not CMS.ExtentionMethods.IsSqlInjectioned(SortText)) Then
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
    ElseIf Not CMS.ExtentionMethods.IsSqlInjectioned(SortText) Then
        SortDirection = Request.QueryString("SortDirection")
    End If

    If (Request.QueryString("searchterms") <> "") Then
        If (Integer.TryParse(Request.QueryString("searchterms"), idNumber)) Then
            idSearch = idNumber.ToString
        Else
            idSearch = "-1"
        End If
        idSearchText = MyCommon.Parse_Quotes(Server.HtmlEncode(Request.QueryString("searchterms")))
        PrctSignPos = idSearchText.IndexOf("%")
        If (PrctSignPos > -1) Then
            idSearch = "-1"
            idSearchText = idSearchText.Replace("%", "[%]")
        End If
        If (idSearchText.IndexOf("_") > -1) Then idSearchText = idSearchText.Replace("_", "[_]")
        If (idSearchText.IndexOf("&amp;") > -1) Then idSearchText = idSearchText.Replace("&amp;", "&")
    End If

    dst = analyticsCGService.GetValidCustomerGroups(SortText, SortDirection, Convert.ToInt64(idSearch), MyCommon.Parse_Quotes(Server.HtmlDecode(idSearchText)))
    sizeOfData = dst.Rows.Count
    i = linesPerPage * PageNum

    If (sizeOfData = 1 AndAlso Request.QueryString("searchterms") <> "") Then
        Response.Status = "301 Moved Permanently"
        Response.AddHeader("Location", "cgroup-edit.aspx?CustomerGroupID=" & MyCommon.NZ(dst.Rows(i).Item("CustomerGroupID"), 0))
    End If
%>
<div id="intro">
  <h1 id="title">
    <% Sendb(Copient.PhraseLib.Lookup("term.customergroups", LanguageID))%>
  </h1>
  <div id="controls">
    <form action="cgroup-edit.aspx" id="controlsform" name="controlsform">
      <%
        If (Logix.UserRoles.CreateCustomerGroups) Then
          Send_New()
        End If
        'If MyCommon.Fetch_SystemOption(75) Then
        '  If (Logix.UserRoles.AccessNotes) Then
        '    Send_NotesButton(5, 0, AdminUserID)
        '  End If
        'End If
      %>
    </form>
  </div>
</div>
<div id="main">
  <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
  <% Send_Listbar(linesPerPage, sizeOfData, PageNum, Server.HtmlEncode(Request.QueryString("searchterms")), "&amp;SortText=" & SortText & "&amp;SortDirection=" & SortDirection)%>
  <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.customergroups", LanguageID)) %>">
    <thead>
      <tr>
        <th align="left" class="th-xid" scope="col">
          <a href="cgroup-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&amp;SortText=ExtGroupID&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.xid", LanguageID))%>
          </a>
          <%
            If SortText = "ExtGroupID" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" class="th-id" scope="col">
          <a href="cgroup-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&amp;SortText=CustomerGroupID&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>
          </a>
          <%
            If SortText = "CustomerGroupID" Then
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
          <a href="cgroup-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&amp;SortText=Name&amp;SortDirection=<% Sendb(SortDirection) %>">
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
        <% If MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CAM) Then%>
        <th align="left" class="th-engine" scope="col">
          <a href="cgroup-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&SortText=CAMCustomerGroup&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.cam", LanguageID))%>
          </a>
          <%
            If SortText = "CAMCustomerGroup" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <% End If %>
        <th align="left" class="th-datetime" scope="col">
          <a href="cgroup-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&amp;SortText=CreatedDate&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.created", LanguageID))%>
          </a>
          <%
            If SortText = "CreatedDate" Then
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
          <a href="cgroup-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&amp;SortText=LastUpdate&amp;SortDirection=<% Sendb(SortDirection) %>">
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
        While (i < sizeOfData AndAlso i < linesPerPage + linesPerPage * PageNum)
          Send("<tr class=""" & Shaded & """>")
          Send("  <td>" & IIf(MyCommon.NZ(dst.Rows(i).Item("ExtGroupID"), "") = "0", "", MyCommon.NZ(dst.Rows(i).Item("ExtGroupID"), "")) & "</td>")
          Send("  <td>" & MyCommon.NZ(dst.Rows(i).Item("CustomerGroupID"), 0) & "</td>")
          If (Not MyCommon.NZ(dst.Rows(i).Item("AnyCardholder"), False) And Not MyCommon.NZ(dst.Rows(i).Item("AnyCustomer"), False)) Then
                  ' Send("  <td><a  href=""cgroup-edit.aspx?CustomerGroupID=" & MyCommon.NZ(dst.Rows(i).Item("CustomerGroupID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(dst.Rows(i).Item("Name"), ""), 45) & "</a></td>")
                  Send(" <td><div style=""width:330px;word-wrap:break-word;""><a  href=""cgroup-edit.aspx?CustomerGroupID=" & MyCommon.NZ(dst.Rows(i).Item("CustomerGroupID"), 0) & """>" & MyCommon.NZ(dst.Rows(i).Item("Name"), "") & "</a></div></td>")
          Else
                  Send("  <td><div style=""width:330px;word-wrap:break-word;"">" & MyCommon.NZ(dst.Rows(i).Item("Name"), "") & "</div></td>")
          End If
          If MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CAM) Then
            Send("  <td>" & IIf(dst.Rows(i).Item("CAMCustomerGroup") = True, Copient.PhraseLib.Lookup("term.cam", LanguageID), "") & "</td>")
          End If
          If (Not IsDBNull(dst.Rows(i).Item("CreatedDate"))) Then
            Send("  <td>" & Logix.ToShortDateTimeString(MyCommon.NZ(dst.Rows(i).Item("CreatedDate"), "1/1/1900"), MyCommon) & "</td>")
          Else
            Send("  <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
          End If
          If (Not IsDBNull(dst.Rows(i).Item("LastUpdate"))) Then
            Send("  <td>" & Logix.ToShortDateTimeString(MyCommon.NZ(dst.Rows(i).Item("LastUpdate"), "1/1/1900"), MyCommon) & "</td>")
          Else
            Send("  <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
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
  '    Send_Notes(5, 0, AdminUserID)
  '  End If
  'End If
done:
  Send_BodyEnd("searchform", "searchterms")
  MyCommon.Close_LogixRT()
  MyCommon = Nothing
  Logix = Nothing
%>
