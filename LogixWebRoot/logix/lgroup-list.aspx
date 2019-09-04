<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.CommonInc" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%
  ' *****************************************************************************
  ' * FILENAME: lgroup-list.aspx 
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
  Dim Shaded As String = "shaded"
  Dim idNumber As Integer
  Dim idSearch As String
  Dim idSearchText As String
  Dim sSearchQuery1 As String
  Dim sSearchQuery2 As String
  Dim PageNum As Integer = 0
  Dim MorePages As Boolean
  Dim linesPerPage As Integer = 20
  Dim sizeOfData As Integer
  Dim i As Integer = 0
  Dim PrctSignPos As Integer
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim BannersEnabled As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "lgroup-list.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  PageNum = Request.QueryString("pagenum")
  If PageNum < 0 Then PageNum = 0
  MorePages = False
  Try
    Send_HeadBegin("term.storegroups")
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
    Send_Subtabs(Logix, 70, 1)
    
    If (Logix.UserRoles.AccessStoreGroups = False) Then
      Send_Denied(1, "perm.lgroup-access")
      GoTo done
    End If
    
    
    
    BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
    
    Dim SortText As String = "LocationGroupID"
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
      SortDirection = Server.HtmlEncode(Request.QueryString("SortDirection"))
    End If
    
    sSearchQuery1 = "select BannerID, a.LocationGroupID,a.Name,(select count(*) from Locations with (NoLock) where Deleted = 0) as Locations,a.CreatedDate,a.LastUpdate,a.AllLocations, pe.description as EngineName, pe.phraseid as EnginePhraseID from LocationGroups a with (NoLock) inner join PromoEngines pe on a.engineid = pe.engineid where a.AllLocations = 1 and a.Deleted = 0"
    sSearchQuery2 = "union select BannerID, a.LocationGroupID,a.Name,(select count(*) from LocGroupItems b with (NoLock) where b.Deleted = 0 and b.LocationGroupId = a.LocationGroupId) as Locations,a.CreatedDate,a.LastUpdate,a.AllLocations, pe.description as EngineName, pe.phraseid as EnginePhraseID from LocationGroups a with (NoLock) inner join PromoEngines pe with (NoLock) on a.engineid = pe.engineid where a.AllLocations = 0 and a.Deleted = 0"
    
    If (Request.QueryString("searchterms") <> "") Then
      If (Integer.TryParse(Request.QueryString("searchterms"), idNumber)) Then
        idSearch = idNumber.ToString
      Else
        idSearch = "-1"
      End If
      idSearchText = MyCommon.Parse_Quotes(HttpUtility.UrlDecode(Request.QueryString("searchterms")))
      PrctSignPos = idSearchText.IndexOf("%")
      If (PrctSignPos > -1) Then
        idSearch = -1
        idSearchText = idSearchText.Replace("%", "[%]")
      End If
      If (idSearchText.IndexOf("_") > -1) Then idSearchText = idSearchText.Replace("_", "[_]")
      sSearchQuery1 = sSearchQuery1 & " and (a.LocationGroupID=" & idSearch & " or a.Name like N'%" & idSearchText & "%')"
      sSearchQuery2 = sSearchQuery2 & " and (a.LocationGroupID=" & idSearch & " or a.Name like N'%" & idSearchText & "%')"
    End If
        
    MyCommon.QueryStr = "select * from (" & sSearchQuery1 & sSearchQuery2 & ") LocGroups"
    
    If (BannersEnabled) Then
      MyCommon.QueryStr &= " where (BannerID is Null or BannerID =0 or BannerID in (select BannerID from AdminUserBanners where AdminUserID=" & AdminUserID & "))"
    End If
    
    If (SortText <> "") Then
      MyCommon.QueryStr &= " order by " & SortText & " " & SortDirection
    End If
    
    dst = MyCommon.LRT_Select
    sizeOfData = dst.Rows.Count
    i = linesPerPage * PageNum
    
    If (sizeOfData = 1 AndAlso Request.QueryString("searchterms") <> "") Then
      Response.Status = "301 Moved Permanently"
      Response.AddHeader("Location", "lgroup-edit.aspx?LocationGroupID=" & dst.Rows(i).Item("LocationGroupID"))
    End If
%>
<div id="intro">
  <h1 id="title">
    <% Sendb(Copient.PhraseLib.Lookup("term.storegroups", LanguageID))%>
  </h1>
  <div id="controls">
    <form action="lgroup-edit.aspx" id="controlsform" name="controlsform">
      <%
        If (Logix.UserRoles.CreateStoreGroups) Then
          Send_New()
        End If
        'If MyCommon.Fetch_SystemOption(75) Then
        '  If (Logix.UserRoles.AccessNotes) Then
        '    Send_NotesButton(14, 0, AdminUserID)
        '  End If
        'End If
      %>
    </form>
  </div>
</div>
<div id="main">
  <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
  <% Send_Listbar(linesPerPage, sizeOfData, PageNum, Server.HtmlEncode(Request.QueryString("searchterms")), "&SortText=" & SortText & "&SortDirection=" & SortDirection)%>
  <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.storegroups", LanguageID)) %>">
    <thead>
      <tr>
        <th align="left" class="th-id" scope="col">
          <a href="lgroup-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&SortText=LocationGroupID&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>
          </a>
          <%If SortText = "LocationGroupID" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If%>
        </th>
        <th align="left" class="th-engine" scope="col">
          <a href="lgroup-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&SortText=EngineName&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.engine", LanguageID))%>
          </a>
          <%If SortText = "EngineName" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If%>
        </th>
        <th align="left" class="th-name" scope="col">
          <a href="lgroup-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&SortText=Name&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>
          </a>
          <%If SortText = "Name" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If%>
        </th>
        <th align="left" class="th-locations" scope="col" style="text-align: center;">
          <% Sendb(Copient.PhraseLib.Lookup("term.locations", LanguageID))%>
        </th>
        <th align="left" class="th-datetime" scope="col">
          <a href="lgroup-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&SortText=CreatedDate&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.created", LanguageID))%>
          </a>
          <%If SortText = "CreatedDate" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If%>
        </th>
        <th align="left" class="th-datetime" scope="col">
          <a href="lgroup-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&SortText=LastUpdate&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.edited", LanguageID))%>
          </a>
          <%If SortText = "LastUpdate" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If%>
        </th>
      </tr>
    </thead>
    <tbody>
      <% 
        Shaded = "shaded"
        'Response.Write( "Testing query -> " & MyCommon.QueryStr & " <- against database <br />" )	
        While (i < sizeOfData AndAlso i < linesPerPage + linesPerPage * PageNum)
          ' For Each row In dst.Rows
          Send("<tr class=""" & Shaded & """>")
          Send("  <td>" & dst.Rows(i).Item("LocationGroupID") & "</td>")
          Send("  <td>" & Copient.PhraseLib.Lookup(CInt(MyCommon.NZ(dst.Rows(i).Item("EnginePhraseID"), 0)), LanguageID, MyCommon.NZ(dst.Rows(i).Item("EngineName"), "")) & "</td>")
          If (dst.Rows(i).Item("AllLocations").Equals(System.DBNull.Value) OrElse Not dst.Rows(i).Item("AllLocations")) Then
            Send("  <td><a href=""lgroup-edit.aspx?LocationGroupID=" & dst.Rows(i).Item("LocationGroupID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Name"), 30) & "</a></td>")
          Else
            Send("  <td>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Name"), 30) & "</td>")
          End If
          Send("  <td align=""center"">" & dst.Rows(i).Item("Locations") & "</td>")
          If (Not IsDBNull(dst.Rows(i).Item("CreatedDate"))) Then
            Send("  <td>" & Logix.ToShortDateTimeString(dst.Rows(i).Item("CreatedDate"), MyCommon) & "</td>")
          Else
            Send("  <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
          End If
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
          ' Next
        End While
      %>
    </tbody>
  </table>
</div>
<%
  'If MyCommon.Fetch_SystemOption(75) Then
  '  If (Logix.UserRoles.AccessNotes) Then
  '    Send_Notes(14, 0, AdminUserID)
  '  End If
  'End If
done:
  ' Catch ex As Exception
  ' MyCommon.Error_Processor("Catch", ex.Message, "lgroup-list.aspx", "Locations")
  ' Throw ex
Finally
  MyCommon.Close_LogixRT()
End Try
Send_BodyEnd("searchform", "searchterms")
MyCommon = Nothing
Logix = Nothing
%>
