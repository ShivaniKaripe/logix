<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: point-list.aspx 
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
  
  Const LINES_PER_PAGE As Integer = 20
  Dim CopientFileName As String = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
  Dim CopientFileVersion As String = "7.3.1.138972" 
  Dim CopientProject As String = "Copient Logix"
  Dim CopientNotes As String = ""
    Dim SearchTerms As String = ""
  Dim AdminUserID As Long
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim MyPoints As New Copient.Points
  Dim Mapper As New Copient.ListMapper(LINES_PER_PAGE, "ProgramID")
  Dim SearchCriteria As New Copient.ListCriteria
  Dim SortDirection As Copient.ListCriteria.SORT_DIRECTIONS = Copient.ListCriteria.SORT_DIRECTIONS.ASC
  Dim dst As System.Data.DataTable
  Dim row As DataRow
  Dim Shaded As String = "shaded"
  Dim sizeOfData As Integer
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "point-list.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
    
  Send_HeadBegin("term.pointsprograms")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
  Send_HeadEnd()
  Send_BodyBegin(1)
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  Send_Tabs(Logix, 5)
  Send_Subtabs(Logix, 50, 1)
  
  If (Logix.UserRoles.AccessPointsPrograms = False) Then
    Send_Denied(1, "perm.points-access")
    GoTo done
  End If
  
  SearchCriteria = Mapper.MapFields(Request.QueryString)
  dst = MyPoints.GetPointsProgramList(SearchCriteria, sizeOfData)

  If (sizeOfData = 1 AndAlso Request.QueryString("searchterms") <> "") Then
    Response.Status = "301 Moved Permanently"
        Response.AddHeader("Location", "point-edit.aspx?ProgramGroupID=" & dst.Rows(0).Item("ProgramID"))
    End If
    If Request.QueryString("searchterms") <> "" Then
        SearchTerms = Request.QueryString("searchterms")
    End If
%>
<div id="intro">
  <h1 id="title">
    <% Sendb(Copient.PhraseLib.Lookup("term.pointsprograms", LanguageID))%>
  </h1>
  <div id="controls">
    <form action="point-edit.aspx" id="controlsform" name="controlsform">
      <%
        If (Logix.UserRoles.CreatePointsPrograms) Then
          Send_New()
        End If
      %>
    </form>
  </div>
</div>
<div id="main">
  <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
  <%  Send_Listbar(LINES_PER_PAGE, sizeOfData, SearchCriteria.GetCurrentPage, Server.HtmlEncode(SearchTerms), "&SortText=" & SearchCriteria.GetSortField & "&SortDirection=" & SearchCriteria.GetSortDirection)%>
  <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.points", LanguageID)) %>">
    <thead>
      <tr>
        <th align="left" class="th-id" scope="col">
          <a href="point-list.aspx?searchterms=<%sendb(Server.HtmlEncode(SearchTerms)) %>&SortText=ProgramID&SortDirection=<% Sendb(SearchCriteria.GetSortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>
          </a>
          <%
            If SearchCriteria.GetSortField = "ProgramID" Then
              If SearchCriteria.GetSortDirection = Copient.ListCriteria.SORT_DIRECTIONS.ASC Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" class="th-name" scope="col">
          <a href="point-list.aspx?searchterms=<%sendb(Server.HtmlEncode(SearchTerms)) %>&SortText=ProgramName&SortDirection=<% Sendb(SearchCriteria.GetSortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>
          </a>
          <%
            If SearchCriteria.GetSortField = "ProgramName" Then
              If SearchCriteria.GetSortDirection = Copient.ListCriteria.SORT_DIRECTIONS.ASC Then
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
          <a href="point-list.aspx?searchterms=<%sendb(Server.HtmlEncode(SearchTerms)) %>&SortText=CAMProgram&SortDirection=<% Sendb(SearchCriteria.GetSortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.cam", LanguageID))%>
          </a>
          <%
            If SearchCriteria.GetSortField = "CAMProgram" Then
              If SearchCriteria.GetSortDirection = Copient.ListCriteria.SORT_DIRECTIONS.ASC Then
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
          <a href="point-list.aspx?searchterms=<%sendb(Server.HtmlEncode(SearchTerms)) %>&SortText=CreatedDate&SortDirection=<% Sendb(SearchCriteria.GetSortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.created", LanguageID))%>
          </a>
          <%
            If SearchCriteria.GetSortField = "CreatedDate" Then
              If SearchCriteria.GetSortDirection = Copient.ListCriteria.SORT_DIRECTIONS.ASC Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" class="th-datetime" scope="col">
          <a href="point-list.aspx?searchterms=<%sendb(Server.HtmlEncode(SearchTerms)) %>&SortText=LastUpdate&SortDirection=<% Sendb(SearchCriteria.GetSortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.edited", LanguageID))%>
          </a>
          <%
            If SearchCriteria.GetSortField = "LastUpdate" Then
              If SearchCriteria.GetSortDirection = Copient.ListCriteria.SORT_DIRECTIONS.ASC Then
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
        For Each row In dst.Rows
          Send("<tr class=""" & Shaded & """>")
          Send("  <td>" & MyCommon.NZ(row.Item("ProgramID"), 0) & "</td>")
          Send("  <td><a href=""point-edit.aspx?ProgramGroupID=" & MyCommon.NZ(row.Item("ProgramID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("ProgramName"), ""), 30) & "</a></td>")
          If MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CAM) Then
            Send("  <td>" & IIf(MyCommon.NZ(row.Item("CAMProgram"), False) = True, Copient.PhraseLib.Lookup("term.cam", LanguageID), "") & "</td>")
          End If
          If (Not IsDBNull(row.Item("CreatedDate"))) Then
            Send("  <td>" & Logix.ToShortDateTimeString(row.Item("CreatedDate"), MyCommon) & "</td>")
          Else
            Send("  <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
          End If
          If (Not IsDBNull(row.Item("LastUpdate"))) Then
            Send("  <td>" & Logix.ToShortDateTimeString(row.Item("LastUpdate"), MyCommon) & "</td>")
          Else
            Send("  <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
          End If
          Send("</tr>")
          Shaded = IIf(Shaded = "shaded", "", "shaded")
        Next
      %>
    </tbody>
  </table>
</div>
<%
done:
  Send_BodyEnd("searchform", "searchterms")
  MyCommon.Close_LogixRT()
  Logix = Nothing
  MyCommon = Nothing
%>
