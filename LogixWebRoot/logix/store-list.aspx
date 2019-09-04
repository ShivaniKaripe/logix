<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.CommonInc" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%
  ' *****************************************************************************
  ' * FILENAME: store-list.aspx 
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
  Dim Shaded As String = "shaded"
  Dim PrctSignPos As Integer
  Dim SubPhraseID As Integer
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim BannersEnabled As Boolean = False
  Dim LocationTypeID As Integer = 1
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "store-list.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  
  PageNum = Request.QueryString("pagenum")
  If PageNum < 0 Then
    PageNum = 0
  End If
  MorePages = False
  
  Try
    If (Request.QueryString("LocationTypeID") <> "") Then
      LocationTypeID = MyCommon.Extract_Val(Request.QueryString("LocationTypeID"))
    End If
    If LocationTypeID = 2 Then
      Send_HeadBegin("term.servers")
    Else
      Send_HeadBegin("term.stores")
    End If
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts()
    Send_HeadEnd()
    Send_BodyBegin(1)
    Send_Bar(Handheld)
    Send_Help(CopientFileName)
    Send_Logos()
    Send_Tabs(Logix, 7)
    If LocationTypeID = 2 Then
      Send_Subtabs(Logix, 70, 3)
    Else
      Send_Subtabs(Logix, 70, 2)
    End If
    
    If (Logix.UserRoles.AccessStores = False) Then
      Send_Denied(1, "perm.store-access")
      GoTo done
    End If
    
    
    BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
    
    Dim SortText As String = "LocationID"
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
    
    MyCommon.QueryStr = "select SubPhraseID from Countries with (NoLock) where CountryID=" & MyCommon.Fetch_SystemOption(65)
    dst = MyCommon.LRT_Select
    SubPhraseID = MyCommon.NZ(dst.Rows(0).Item(0), "0")
    
    sSearchQuery = "select L.LocationID, L.ExtLocationCode, L.LocationName, L.City, L.State, L.LastUpdate, L.TestingLocation, L.EngineID, PE.Description, PE.PhraseID as EnginePhraseID " & _
                   "from Locations as L with (NoLock) " & _
                   "inner join PromoEngines as PE with (NoLock) on L.EngineID=PE.EngineID " & _
                   "where Deleted=0"
    If LocationTypeID = 1 Then
      sSearchQuery &= " and LocationTypeID=1"
    ElseIf LocationTypeID = 2 Then
      sSearchQuery &= " and LocationTypeID=2"
    End If
    idSearchText = MyCommon.Parse_Quotes(HttpUtility.UrlDecode(Request.QueryString("searchterms")))
    If (idSearchText <> "") Then
      If (Integer.TryParse(Request.QueryString("searchterms"), idNumber)) Then
        idSearch = idNumber.ToString
      Else
        idSearch = "-1"
      End If
      PrctSignPos = idSearchText.IndexOf("%")
      If (PrctSignPos > -1) Then
        idSearch = -1
        idSearchText = idSearchText.Replace("%", "[%]")
      End If
      If (idSearchText.IndexOf("_") > -1) Then idSearchText = idSearchText.Replace("_", "[_]")
      sSearchQuery &= " and (LocationID=" & idSearch & " "
      sSearchQuery &= " or ExtLocationCode like '%" & Server.HtmlDecode(idSearchText) & "%'"
      sSearchQuery &= " or LocationName like N'%" & Server.HtmlDecode(idSearchText) & "%'"
      sSearchQuery &= " or City like '%" & Server.HtmlDecode(idSearchText) & "%'"
      sSearchQuery &= " or State like '%" & Server.HtmlDecode(idSearchText) & "%'"
      sSearchQuery &= " or PE.Description like '%" & Server.HtmlDecode(idSearchText) & "%')"
    End If
    
    If (BannersEnabled) Then
      sSearchQuery &= " and (BannerID is Null or BannerID =0 or BannerID in (select BannerID from AdminUserBanners where AdminUserID=" & AdminUserID & ") " & _
                      "      or EXISTS(select BAN.BannerID from Banners BAN with (NoLock) " & _
                      "                inner join BannerEngines BE with (NoLock) on BE.BannerID = BAN.BannerID " & _
                      "                inner join AdminUserBanners AUB with (NoLock) on AUB.BannerID = BAN.BannerID " & _
                      "                where BE.EngineID = L.EngineID and AUB.AdminUserID=" & AdminUserID & " and BAN.AllBanners=1 and BAN.Deleted=0) " & _
                      ")"
    End If
    If LocationTypeID = 1 Then
      sSearchQuery &= " and LocationTypeID=1"
    ElseIf LocationTypeID = 2 Then
      sSearchQuery &= " and LocationTypeID=2"
    End If
    sSearchQuery = sSearchQuery & " order by " & SortText & " " & SortDirection
    MyCommon.QueryStr = sSearchQuery
    dst = MyCommon.LRT_Select
    sizeOfData = dst.Rows.Count
    i = linesPerPage * PageNum
    If (sizeOfData = 1 AndAlso Server.HtmlEncode(Request.QueryString("searchterms")) <> "") Then
      Response.Status = "301 Moved Permanently"
      Response.AddHeader("Location", "store-edit.aspx?LocationID=" & dst.Rows(i).Item("LocationID"))
    End If
%>
<div id="intro">
  <h1 id="title">
    <%
      If LocationTypeID = 2 Then
        Sendb(Copient.PhraseLib.Lookup("term.servers", LanguageID))
      Else
        Sendb(Copient.PhraseLib.Lookup("term.stores", LanguageID))
      End If
    %>
  </h1>
  <div id="controls">
    <form action="store-edit.aspx" id="controlsform" name="controlsform">
      <%
        If Logix.UserRoles.CreateStores Then
          Send_New()
        End If
        If LocationTypeID > 0 Then
          Send("<input type=""hidden"" id=""LocationTypeID"" name=""LocationTypeID"" value=""" & LocationTypeID & """ />")
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
    Send_Listbar(linesPerPage, sizeOfData, PageNum, Server.HtmlEncode(Request.QueryString("searchterms")), "&amp;SortText=" & SortText & "&amp;SortDirection=" & SortDirection)
  %>
  <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.stores", LanguageID)) %>">
    <thead>
      <tr>
        <th align="left" class="th-id" scope="col">
          <a href="store-list.aspx?<%Sendb(IIf(LocationTypeID>0, "LocationTypeID=" & LocationTypeID & "&amp;", ""))%>searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&amp;SortText=LocationID&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>
          </a>
          <%
            If SortText = "LocationID" Then
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
          <a href="store-list.aspx?<%Sendb(IIf(LocationTypeID>0, "LocationTypeID=" & LocationTypeID & "&amp;", ""))%>searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&amp;SortText=ExtLocationCode&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.code", LanguageID))%>
          </a>
          <%
            If SortText = "ExtLocationCode" Then
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
          <a href="store-list.aspx?<%Sendb(IIf(LocationTypeID>0, "LocationTypeID=" & LocationTypeID & "&amp;", ""))%>searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&amp;SortText=LocationName&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>
          </a>
          <%
            If SortText = "LocationName" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <%
          'Show city and state columns only for true stores, not servers
          If LocationTypeID = 1 Then
        %>
        <th align="left" class="th-city" scope="col">
          <a href="store-list.aspx?<%Sendb(IIf(LocationTypeID>0, "LocationTypeID=" & LocationTypeID & "&amp;", ""))%>searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&amp;SortText=City&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.city", LanguageID))%>
          </a>
          <%
            If SortText = "City" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" class="th-state" scope="col">
          <a href="store-list.aspx?<%Sendb(IIf(LocationTypeID>0, "LocationTypeID=" & LocationTypeID & "&amp;", ""))%>searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&amp;SortText=State&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup(SubPhraseID, LanguageID))%>
          </a>
          <%
            If SortText = "State" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <% End If%>
        <th align="left" class="th-engine" scope="col">
          <a href="store-list.aspx?<%Sendb(IIf(LocationTypeID>0, "LocationTypeID=" & LocationTypeID & "&amp;", ""))%>searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&amp;SortText=Description&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.engine", LanguageID))%>
          </a>
          <%
            If SortText = "Description" Then
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
          <a href="store-list.aspx?<%Sendb(IIf(LocationTypeID>0, "LocationTypeID=" & LocationTypeID & "&amp;", ""))%>searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&amp;SortText=LastUpdate&amp;SortDirection=<% Sendb(SortDirection) %>">
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
          Send("<tr class=""" & Shaded & """>")
          Send("  <td>" & dst.Rows(i).Item("LocationID") & "</td>")
          Send("  <td>" & MyCommon.NZ(dst.Rows(i).Item("ExtLocationCode"), "&nbsp;") & "</td>")
          Sendb("  <td><a href=""store-edit.aspx?LocationID=" & dst.Rows(i).Item("LocationID") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(dst.Rows(i).Item("LocationName"), "&nbsp;&nbsp;"), 20) & "</a>")
          If (dst.Rows(i).Item("TestingLocation") = True) Then
            Sendb("<span> (" & Copient.PhraseLib.Lookup("term.testlab", LanguageID) & ")</span>")
          End If
          Send("</td>")
          If LocationTypeID = 1 Then
            Send("  <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(dst.Rows(i).Item("City"), "&nbsp;"), 20) & "</td>")
            Send("  <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(dst.Rows(i).Item("State"), "&nbsp;"), 20) & "</td>")
          End If
          Send("  <td>" & MyCommon.SplitNonSpacedString(Copient.PhraseLib.Lookup(CInt(MyCommon.NZ(dst.Rows(i).Item("EnginePhraseID"), 0)), LanguageID, MyCommon.NZ(dst.Rows(i).Item("Description"), "&nbsp;")), 20) & "</td>")
          If (Not IsDBNull(dst.Rows(i).Item("LastUpdate"))) Then
            Send("  <td>" & Logix.ToShortDateTimeString(dst.Rows(i).Item("LastUpdate"), MyCommon) & "</td>")
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
done:
Finally
  MyCommon.Close_LogixRT()
End Try
Send_BodyEnd("searchform", "searchterms")
Logix = Nothing
MyCommon = Nothing
%>
