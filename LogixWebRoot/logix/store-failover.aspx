<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: store-failover.aspx 
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
  Dim dstStoreList As System.Data.DataTable
  Dim row As System.Data.DataRow
  Dim shaded As Boolean
  Dim SortText As String = "L.LocationID"
  Dim SortDirection As String = ""
  Dim idNumber As Integer = 0
  Dim idSearch As String
  Dim idSearchText As String
  Dim PageNum As Integer = 0
  Dim MorePages As Boolean
  Dim linesPerPage As Integer = 20
  Dim sizeOfData As Integer
  Dim i As Integer = 0
  Dim SearchClause As String = ""
  Dim lastHeardAlertMins As Integer = 20
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim IsCPEInstalled As Boolean = False
  Dim BannersEnabled As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "store-failover.aspx"
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  MyCommon.Open_LogixRT()
    
  PageNum = Request.QueryString("pagenum")
  If PageNum < 0 Then PageNum = 0
  MorePages = False
  shaded = True
  
  Send_HeadBegin("term.failover")
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
  Send_Subtabs(Logix, 8, 7)
  
  If (Logix.UserRoles.AccessStoreHealth = False) Then
    Send_Denied(1, "perm.admin-store-health")
    GoTo done
  End If
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  IsCPEInstalled = MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE)
  
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
    idSearchText = MyCommon.Parse_Quotes(Request.QueryString("searchterms"))
    If (Integer.TryParse(idSearchText, idNumber)) Then
      idSearch = idNumber
    Else
      idSearch = -1
    End If
    If (idSearchText.IndexOf("_") > -1) Then idSearchText = idSearchText.Replace("_", "[_]")
    SearchClause = " and (LS.LocalServerID=" & idSearch & " or L.LocationID=" & idSearch & " or L.LocationName like '%" & idSearchText & "%' " & " or L.ExtLocationCode like '%" & idSearchText & "%') "
  End If
  
  MyCommon.QueryStr = "select L.LocationID, L.LocationName, L.ExtLocationCode, LS.FailoverServer, " & _
                      "LS.LastIP, LS.MacAddress, LS.LastLocationID, LS.LocalServerID, " & _
                      "LS.TransactionLastHeard, LS.TransDownloadLastHeard, LS.IncentiveLastHeard, LS.LastHeard " & _
                      "from LocalServers as Ls " & _
                      "Left Join Locations as L on L.LocationID = LS.LocationID and L.Deleted = 0 " & _
                      "where LS.FailoverServer=1 " & _
                      "union " & _
                      "select 0 as LocationID, '' as LocationName, '' as ExtLocationCode, LS.FailoverServer, " & _
                      "LS.LastIP, LS.MacAddress, LS.LastLocationID, LS.LocalServerID, " & _
                      "LS.TransactionLastHeard, LS.TransDownloadLastHeard, LS.IncentiveLastHeard, LS.LastHeard " & _
                      "from LocalServers as Ls where LocationID = 0 "
  ' check if banners are enabled
  MyCommon.QueryStr &= " order by " & SortText & " " & SortDirection
  
  'Send(MyCommon.QueryStr)
  'GoTo done
  dstStoreList = MyCommon.LRT_Select
  sizeOfData = MyCommon.NZ(dstStoreList.Rows.Count, 0)
  ' set i
  i = linesPerPage * PageNum
%>
<div id="intro">
  <h1 id="title" style="display:inline;">
    <% Sendb(Copient.PhraseLib.Lookup("term.failoverservers", LanguageID))%>
  </h1>
    <form action="<%Sendb("log-view.aspx?filetype=-1&amp;fileyear=" & Year(Today) & "&amp;filemonth=" & Month(Today) & "&amp;fileday=" & Day(Today)) %>" id="controlsform" name="controlsform" target="_blank" style="float: right;">
      <div id="controls">
        <input type="submit" class="regular" id="logs" name="logs" value="<% Sendb(Copient.PhraseLib.Lookup("term.logs", LanguageID)) %>..." />
        <%
          'If MyCommon.Fetch_SystemOption(75) Then
          '  If (Logix.UserRoles.AccessNotes) Then
          '    Send_NotesButton(15, 0, AdminUserID)
          '  End If
          'End If
        %>
      </div>
    </form>
</div>
<div id="main">
  <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
  <% Send_Listbar(linesPerPage, sizeOfData, PageNum, Request.QueryString("searchterms"), "&SortText=" & SortText & "&SortDirection=" & SortDirection)%>
  <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.health", LanguageID)) %>">
    <thead>
      <tr>
        <th align="left" class="th-id" scope="col" valign="bottom">
          <a href="store-failover.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&action=Sort&SortText=L.LocationID&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>
          </a>
          <%If SortText = "L.LocationID" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" class="th-id" scope="col" valign="bottom">
          <a href="store-failover.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&action=Sort&SortText=LS.LastLocationID&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.lastlocationserviced", LanguageID))%>
          </a>
          <%If SortText = "LS.LastLocationID" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>        
        <th align="left" class="th-id" scope="col" valign="bottom">
          <a href="store-failover.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&action=Sort&SortText=LS.LocalServerID&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.lsid", LanguageID))%>
          </a>
          <%If SortText = "LS.LocalServerID" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>        
        <th align="left" class="th-code" scope="col" valign="bottom">
          <a href="store-failover.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&action=Sort&SortText=L.ExtLocationCode&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.code", LanguageID))%>
          </a>
          <%If SortText = "L.ExtLocationCode" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" class="th-location" scope="col" valign="bottom">
          <a href="store-failover.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&action=Sort&SortText=L.LocationName&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.location", LanguageID))%>
          </a>
          <%If SortText = "L.LocationName" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" scope="col" valign="bottom">
          <a href="store-failover.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&action=Sort&SortText=LS.LastIP&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.lastip", LanguageID))%>
          </a>
          <%If SortText = "LS.LastIP" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" scope="col" valign="bottom">
          <a href="store-failover.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&action=Sort&SortText=LS.MacAddress&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.macaddress", LanguageID))%>
          </a>
          <%If SortText = "LS.MacAddress" Then
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
        Dim LID As Integer
        
        MyCommon.QueryStr = "select OptionValue from SystemOptions with (nolock) where OptionID = 41;"
        dst = MyCommon.LRT_Select
        If (dst.Rows.Count > 0) Then
          row = dst.Rows(0)
          If (IsNumeric(row.Item("OptionValue"))) Then
            lastHeardAlertMins = CInt(row.Item("OptionValue"))
          End If
        End If
        If (dstStoreList.Rows.Count > 0) Then
          While (i < sizeOfData And i < linesPerPage + linesPerPage * PageNum)
            row = dstStoreList.Rows(i)
            LID = MyCommon.NZ(row.Item("LocationID"), 0)

            If (LID = 0) Then
              Send("<tr style=""background-color:#eeeeee;"">")
            Else
              If (shaded) Then
                Send("<tr class=""shaded"">")
                shaded = False
              Else
                Send("<tr>")
                shaded = True
              End If
            End If
            
            ' Location ID, External Location Code and location name columns
            Send("     <td>" & MyCommon.NZ(row.Item("LocationID"), "&nbsp;") & "</td>")
            Send("     <td>" & " <a href=""store-detail.aspx?LocationID=" & row.Item("LastLocationID") & """>" & MyCommon.NZ(row.Item("LastLocationID"), "&nbsp;") & "</a></td>")
            Send("     <td>" & MyCommon.NZ(row.Item("LocalServerID"), "&nbsp;") & "</td>")
              
            Send("<td>" & MyCommon.NZ(row.Item("ExtLocationCode"), "&nbsp;") & "</td>")
            If (Logix.UserRoles.CRUDStoresAndTerminals = True) Then
              Send("<td>" & " <a href=""store-detail.aspx?LocationID=" & row.Item("LocationID") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("LocationName"), "&nbsp;"), 25) & "</a></td>")
            Else
              Send("<td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("LocationName"), "&nbsp;"), 25) & "</td>")
            End If
            
            ' Last IP Address column
            Send("     <td>" & MyCommon.NZ(row.Item("LastIP"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)) & "</td>")
            
            ' MAC Address column
            Send("     <td>" & MyCommon.NZ(row.Item("MacAddress"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)) & "</td>")
            
            Send("</tr>")
            i += 1
          End While
          
        End If
      %>
    </tbody>
  </table>
</div>

<%
  'If MyCommon.Fetch_SystemOption(75) Then
  '  If (Logix.UserRoles.AccessNotes) Then
  '    Send_Notes(15, 0, AdminUserID)
  '  End If
  'End If
done:
  Send_BodyEnd("searchform", "searchterms")
  MyCommon.Close_LogixRT()
  MyCommon = Nothing
  Logix = Nothing
%>
