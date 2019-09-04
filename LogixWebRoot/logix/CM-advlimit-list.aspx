<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: CM-advlimit-list.aspx 
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
  
  Dim AdminUserID As Long
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim Mapper As New Copient.ListMapper(LINES_PER_PAGE, "LimitID")
  Dim SearchCriteria As New Copient.ListCriteria
  Dim row As DataRow
  Dim Shaded As String = "shaded"
  Dim sizeOfData As Integer
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False

  Dim dt As DataTable = Nothing
  Dim dst As DataTable = Nothing
  Dim SearchText As String = ""
  Dim SearchID As Long
  Dim SortText As String = ""
  Dim SortDirection As Copient.ListCriteria.SORT_DIRECTIONS = Copient.ListCriteria.SORT_DIRECTIONS.ASC
  Dim PrctSignPos, Index As Integer
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "CM-advlimit-list.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
    
  Send_HeadBegin("term.advlimit")
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
  Send_Subtabs(Logix, 50, 4)
  
  
  If (Logix.UserRoles.AccessPointsPrograms = False) Then
    Send_Denied(1, "perm.points-access")
    GoTo done
  End If
  
  SearchCriteria = Mapper.MapFields(Request.QueryString)
  If SearchCriteria IsNot Nothing Then
    If SearchCriteria.GetSearchText IsNot Nothing Then SearchText = MyCommon.Parse_Quotes(SearchCriteria.GetSearchText.Trim)
    If SearchCriteria.GetSortField IsNot Nothing Then SortText = SearchCriteria.GetSortField.Trim
    SortDirection = SearchCriteria.GetSortDirection

    ' handle the searching and sorting
    If SearchText = "" Then
      ' no search strings, just query
      If SortText <> "" Then
        MyCommon.QueryStr = "select AL.LimitID, AL.Name, ALT.PhraseID as TypePhraseID, AL.CreatedDate, AL.LastUpdate from CM_AdvancedLimits as AL with (NoLock) " & _
                            "inner join CM_AdvancedLimitTypes as ALT with (NoLock) on ALT.TypeId=AL.LimitTypeID " & _
                            "where AL.Deleted=0  order by " & SortText & " " & IIf(SortDirection = Copient.ListCriteria.SORT_DIRECTIONS.ASC, "ASC;", "DESC;")
      Else
        MyCommon.QueryStr = "select AL.LimitID, AL.Name, ALT.PhraseID as TypePhraseID, AL.CreatedDate, AL.LastUpdate from CM_AdvancedLimits as AL with (NoLock) " & _
                            "inner join CM_AdvancedLimitTypes as ALT with (NoLock) on ALT.TypeId=AL.LimitTypeID " & _
                            "where AL.Deleted=0;"
      End If
    Else
      If Not Integer.TryParse(SearchText, SearchID) Then
        SearchID = -1
      End If
      PrctSignPos = SearchText.IndexOf("%")
      If (PrctSignPos > -1) Then
        SearchID = -1
        SearchText = SearchText.Replace("%", "[%]")
      End If
      If (SearchText.IndexOf("_") > -1) Then SearchText = SearchText.Replace("_", "[_]")

      If SortText <> "" Then
                MyCommon.QueryStr = "select AL.LimitID, AL.Name, ALT.PhraseID as TypePhraseID, AL.CreatedDate, AL.LastUpdate from CM_AdvancedLimits as AL with (NoLock) " & _
                                    "inner join CM_AdvancedLimitTypes as ALT with (NoLock) on ALT.TypeId=AL.LimitTypeID " & _
                                    "where Deleted=0 and (LimitID =" & SearchID & " or Name like N'%" & SearchText & "%')  order by " & SortText & _
                                    " " & IIf(SortDirection = Copient.ListCriteria.SORT_DIRECTIONS.ASC, "ASC;", "DESC;")
      Else
        MyCommon.QueryStr = "select AL.LimitID, AL.Name, ALT.PhraseID as TypePhraseID, AL.CreatedDate, AL.LastUpdate from CM_AdvancedLimits as AL with (NoLock) " & _
                            "inner join CM_AdvancedLimitTypes as ALT with (NoLock) on ALT.TypeId=AL.LimitTypeID " & _
                            "where Deleted=0 and (LimitID =" & SearchID & " or Name like N'%" & SearchText & "%')"
      End If
    End If

    dt = MyCommon.LRT_Select
    sizeOfData = dt.Rows.Count

    ' handle the paging
    If SearchCriteria.GetPageSize > 0 Then
      dst = dt.Clone
      Index = SearchCriteria.GetPageSize * SearchCriteria.GetCurrentPage
      While (Index < sizeOfData And Index < SearchCriteria.GetPageSize + SearchCriteria.GetPageSize * SearchCriteria.GetCurrentPage)
        dst.ImportRow(dt.Rows(Index))
        Index += 1
      End While
    Else
            dst = dt.Copy
    End If
  End If

  If (sizeOfData = 1 AndAlso Request.QueryString("searchterms") <> "") Then
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "CM-advlimit-edit.aspx?LimitID=" & dst.Rows(0).Item("LimitID"))
  End If
%>
<div id="intro">
  <h1 id="title">
    <% Sendb(Copient.PhraseLib.Lookup("term.advlimit", LanguageID))%>
  </h1>
  <div id="controls">
    <form action="CM-advlimit-edit.aspx" id="controlsform" name="controlsform">
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
  <%  Send_Listbar(LINES_PER_PAGE, sizeOfData, SearchCriteria.GetCurrentPage, Request.QueryString("searchterms"), "&SortText=" & SearchCriteria.GetSortField & "&SortDirection=" & SearchCriteria.GetSortDirection)%>
  <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.advlimit", LanguageID)) %>">
    <thead>
      <tr>
        <th align="left" class="th-id" scope="col">
          <a href="CM-advlimit-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&SortText=LimitID&SortDirection=<% Sendb(SearchCriteria.GetSortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>
          </a>
          <%
            If SearchCriteria.GetSortField = "LimitID" Then
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
          <a href="CM-advlimit-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&SortText=Name&SortDirection=<% Sendb(SearchCriteria.GetSortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>
          </a>
          <%
            If SearchCriteria.GetSortField = "Name" Then
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
          <a href="CM-advlimit-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&SortText=TypeID&SortDirection=<% Sendb(SearchCriteria.GetSortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.type", LanguageID))%>
          </a>
          <%
            If SearchCriteria.GetSortField = "TypeID" Then
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
          <a href="CM-advlimit-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&SortText=CreatedDate&SortDirection=<% Sendb(SearchCriteria.GetSortDirection) %>">
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
          <a href="CM-advlimit-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&SortText=LastUpdate&SortDirection=<% Sendb(SearchCriteria.GetSortDirection) %>">
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
          Send("  <td>" & MyCommon.NZ(row.Item("LimitID"), 0) & "</td>")
          Send("  <td><a href=""CM-advlimit-edit.aspx?LimitID=" & MyCommon.NZ(row.Item("LimitID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 30) & "</a></td>")
          Send("  <td>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("TypePhraseID"), 0), LanguageID) & "</td>")
          If (Not IsDBNull(row.Item("CreatedDate"))) Then
            Send("  <td>" & Format(row.Item("CreatedDate"), "dd MMM yyyy, HH:mm:ss") & "</td>")
          Else
            Send("  <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
          End If
          If (Not IsDBNull(row.Item("LastUpdate"))) Then
            Send("  <td>" & Format(row.Item("LastUpdate"), "dd MMM yyyy, HH:mm:ss") & "</td>")
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
