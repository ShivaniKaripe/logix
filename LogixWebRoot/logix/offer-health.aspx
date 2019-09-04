<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: offer-health.aspx 
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
  Dim dst2 As System.Data.DataTable
  Dim row2 As System.Data.DataRow
  Dim shaded As String = ""
  Dim tdTag As String
  Dim rowAlert As String
  Dim lastHeardAlertMins As Integer = 20
  Dim SortText As String = "IncentiveID"
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
  Dim SanityCheckResult As String = ""
  Dim FilterHealth As Integer = 0
  Dim FilterClause As String = ""
  Dim IncentiveFetchResult As String = ""
  Dim IncentiveFetchFilter As String = ""
  Dim ValidLocations As Integer = 0
  Dim WatchLocations As Integer = 0
  Dim WarningLocations As Integer = 0
  Dim ValidOfferComponents As Boolean = False
  Dim RowColor = ""
  Dim dtHealth As DataTable = Nothing
  Dim rowHealth As DataRow = Nothing
  Dim rows As DataRow() = Nothing
  Dim rowCt As Integer = 0
  Dim pageEnd As Integer = 0
  Dim bailOut As Integer = 0
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim BannersEnabled As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "offer-health.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  Dim IsEngineInstalled As Boolean = MyCommon.GetInstalledEngines().Length > 1

  PageNum = Request.QueryString("pagenum")
  If PageNum < 0 Then PageNum = 0
  MorePages = False
  
  Send_HeadBegin("term.offerhealth")
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
  Send_Subtabs(Logix, 8, 6)
  
  If (Logix.UserRoles.ViewOfferHealth = False) Then
    Send_Denied(1, "perm.admin-offer-health")
    GoTo done
  End If
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  FilterHealth = MyCommon.Extract_Val(Request.QueryString("filterhealth"))
  
  If FilterHealth = 3 Then
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "CRMoffer-list.aspx?&filterhealth=" & FilterHealth)
  End If
  
  ' check if the user clicked to mark all offers to revalidate
  If (Request.QueryString("revalidateall") <> "") Then
    MyCommon.QueryStr = "update ValidSummary with (RowLock) set Pending=1;"
    MyCommon.LRT_Execute()
    MyCommon.QueryStr = "insert into ValidSummary with (RowLock)" & _
                        "  select IncentiveID, null, null, null, null, null, 1 " & _
                        "  from OfferValidationView with (NoLock) " & _
                        "  where(Pending Is Null Or Pending <> 1) and IncentiveID Not IN (select OfferID from ValidSummary with (NoLock));"
    MyCommon.LRT_Execute()
  ElseIf (Request.QueryString("offer") <> "") Then
    MyCommon.QueryStr = "update ValidSummary with (RowLock) set Pending=1 where OfferID=" & MyCommon.Extract_Val(Request.QueryString("offer")) & ";"
    MyCommon.LRT_Execute()
  End If
  
  If (Request.QueryString("SortText") <> "") Then
    SortText = Request.QueryString("SortText")
  End If
  
  If (Request.QueryString("pagenum") = "") Then
        If (Server.HtmlEncode(Request.QueryString("SortDirection")) = "ASC") Then
      SortDirection = "DESC"
        ElseIf (Server.HtmlEncode(Request.QueryString("SortDirection")) = "DESC") Then
      SortDirection = "ASC"
    Else
      SortDirection = "DESC"
    End If
  Else
        SortDirection = Server.HtmlEncode(Request.QueryString("SortDirection"))
  End If
  
    If (Server.HtmlEncode(Request.QueryString("searchterms")) <> "") Then
        idSearchText = MyCommon.Parse_Quotes(Server.HtmlEncode(Request.QueryString("searchterms")))
    If (Integer.TryParse(idSearchText, idNumber)) Then
      idSearch = idNumber
    Else
      idSearch = -1
    End If
    If (idSearchText.IndexOf("_") > -1) Then idSearchText = idSearchText.Replace("_", "[_]")
    SearchClause = " IncentiveID=" & idSearch & " or IncentiveName like '%" & idSearchText & "%' "
  End If
  
  Select Case FilterHealth
    Case 1
      FilterClause = "WarningLocations=0 and ValidOfferComponents=1"
    Case 2
      FilterClause = "WarningLocations>0 or ValidOfferComponents=0"
    Case Else
      FilterClause = "1=1"
  End Select
  
  ' check if banners are enabled
  If (BannersEnabled) Then
    SearchClause &= (IIf(SearchClause.Trim() <> "", " and ", ""))
    SearchClause &= "  (IncentiveID not in (select OfferID from BannerOffers BO with (NoLock)) " & _
                         "   or IncentiveID in (select BO.OfferID from AdminUserBanners AUB with (NoLock) " & _
                         "                  inner join BannerOffers BO with (NoLock) on BO.BannerID = AUB.BannerID" & _
                         "                  where AUB.AdminUserID = " & AdminUserID & ")) "
  End If
  
  MyCommon.QueryStr = "select * from OfferValidationView"
  If (SearchClause.Trim <> "") Then
    MyCommon.QueryStr += " where " & SearchClause
  End If
  dtHealth = MyCommon.LRT_Select
  If (Not dtHealth Is Nothing) Then
    rows = dtHealth.Select(FilterClause, SortText & " " & SortDirection)
    sizeOfData = rows.Length
  End If
  
  i = linesPerPage * PageNum
%>

<script type="text/javascript" language="javascript">
  function setFilter(index) {
    var elem = document.getElementById("filterhealth");
    
    if (elem != null && index < elem.options.length) {
      elem.options[index].selected = true
    }
  }
  function addToPending(offerID) {
    var searchString = '<%Sendb(Server.UrlEncode(Server.HtmlEncode(Request.QueryString("searchterms"))))%>';
    var filterString = '<%Sendb(Server.UrlEncode(Server.HtmlEncode(Request.QueryString("filterhealth"))))%>';
    var sortDirString = '<%Sendb(Server.UrlEncode(Server.HtmlEncode(Request.QueryString("SortDirection"))))%>';
    
    if (sortDirString == '') { sortDirString = 'DESC'; }
    var submitURL = "?pagenum=<%Sendb(PageNum)%>&searchterms=" + searchString + "&SortText=<%Sendb(SortText)%>&SortDirection=" + sortDirString + "&filterhealth=" + filterString + "&offer=" + offerID;
    document.location.href = submitURL;
  }
</script>

<div id="intro">
  <h1 id="title">
    <% Sendb(Copient.PhraseLib.Lookup("term.offerhealth", LanguageID))%>
  </h1>
  <%If MyCommon.IsEngineInstalled(0) OrElse MyCommon.IsEngineInstalled(2) Then %>
  <div id="controls">
    <form name="frmRevalidate" id="frmRevalidate" method="get" action="#">
      <% Send_RevalidateAll()%>
    </form>
  </div>
   <%End If %>
</div>
<div id="main">
  <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
  <% Send_Listbar(linesPerPage, sizeOfData, PageNum, Server.HtmlEncode(Request.QueryString("searchterms")), "&SortText=" & SortText & "&SortDirection=" & SortDirection)%>
  <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.health", LanguageID)) %>">
    <thead>
      <tr>
        <th align="left" class="th-bigid" scope="col" valign="bottom">
          <a href="offer-health.aspx?action=Sort&SortText=IncentiveID&SortDirection=<% Sendb(SortDirection) %>&searchterms=<% Sendb(idSearchText)%>&filterhealth=<%Sendb(Server.UrlEncode(Request.QueryString("filterhealth")))%>">
            <% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>
          </a>
          <%If SortText = "IncentiveID" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" class="th-name" scope="col" valign="bottom">
          <a href="offer-health.aspx?action=Sort&SortText=IncentiveName&SortDirection=<% Sendb(SortDirection) %>&searchterms=<% Sendb(idSearchText)%>&filterhealth=<%Sendb(Server.UrlEncode(Request.QueryString("filterhealth")))%>">
            <% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>
          </a>
          <%If SortText = "IncentiveName" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th class="th-amount" scope="col" style="text-align: center;" valign="bottom">
          <a href="offer-health.aspx?action=Sort&SortText=ValidLocations&SortDirection=<% Sendb(SortDirection) %>&searchterms=<% Sendb(idSearchText)%>&filterhealth=<%Sendb(Server.UrlEncode(Request.QueryString("filterhealth")))%>">
            <% Sendb(Copient.PhraseLib.Lookup("term.valid", LanguageID))%>
          </a>
          <%If SortText = "ValidLocations" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th class="th-amount" scope="col" style="text-align: center;" valign="bottom">
          <a href="offer-health.aspx?action=Sort&SortText=WatchLocations&SortDirection=<% Sendb(SortDirection) %>&searchterms=<% Sendb(idSearchText)%>&filterhealth=<%Sendb(Server.UrlEncode(Request.QueryString("filterhealth")))%>">
            <% Sendb(Copient.PhraseLib.Lookup("term.watch", LanguageID))%>
          </a>
          <%If SortText = "WatchLocations" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th class="th-total" scope="col" style="text-align: center;" valign="bottom">
          <a href="offer-health.aspx?action=Sort&SortText=WarningLocations&SortDirection=<% Sendb(SortDirection) %>&searchterms=<% Sendb(idSearchText)%>&filterhealth=<%Sendb(Server.UrlEncode(Request.QueryString("filterhealth")))%>">
            <% Sendb(Copient.PhraseLib.Lookup("term.warning", LanguageID))%>
          </a>
          <%If SortText = "WarningLocations" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th class="th-components" scope="col" style="text-align: center;" valign="bottom">
          <a href="offer-health.aspx?action=Sort&SortText=ValidOfferComponents&SortDirection=<% Sendb(SortDirection) %>&searchterms=<% Sendb(idSearchText)%>&filterhealth=<%Sendb(Server.UrlEncode(Request.QueryString("filterhealth")))%>">
            <% Sendb(Copient.PhraseLib.Lookup("term.components", LanguageID))%>
            <br />
            <% Sendb(Copient.PhraseLib.Lookup("term.valid", LanguageID))%>
          </a>
          <%If SortText = "ValidOfferComponents" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th class="th-datetime" scope="col" style="text-align: center;" valign="bottom">
          <a href="offer-health.aspx?action=Sort&SortText=LastValidated&SortDirection=<% Sendb(SortDirection) %>&searchterms=<% Sendb(idSearchText)%>&filterhealth=<%Sendb(Server.UrlEncode(Request.QueryString("filterhealth")))%>">
            <% Sendb(Copient.PhraseLib.Lookup("term.lastvalidated", LanguageID))%>
          </a>
          <%If SortText = "LastValidated" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th class="th-total" scope="col" style="text-align: center;" valign="bottom">
          <a href="offer-health.aspx?action=Sort&SortText=Pending&SortDirection=<% Sendb(SortDirection) %>&searchterms=<% Sendb(idSearchText)%>&filterhealth=<%Sendb(Server.UrlEncode(Request.QueryString("filterhealth")))%>">
            <% Sendb(Copient.PhraseLib.Lookup("term.Pending", LanguageID))%>
          </a>
          <%If SortText = "Pending" Then
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
          If (MyCommon.NZ(rows(i).Item("WarningLocations"), 0) > 0 OrElse Not MyCommon.NZ(rows(i).Item("ValidOfferComponents"), False)) Then
            RowColor = "color:red;"
          Else
            RowColor = ""
          End If
                Dim Name As String=""
            If (MyCommon.Fetch_UE_SystemOption(168) = "1" AndAlso MyCommon.NZ(rows(i).Item("BuyerID"), "") <> "") Then
            Name = "Buyer " + rows(i).Item("BuyerID").ToString() + " - " + MyCommon.SplitNonSpacedString(rows(i).Item("IncentiveName"), 25).ToString()
            Else
            Name = MyCommon.NZ(MyCommon.SplitNonSpacedString(rows(i).Item("IncentiveName"), 25), Copient.PhraseLib.Lookup("term.unknown", LanguageID))
            End If
          Send("<tr class=""" & shaded & """>")
          Send("    <td>" & MyCommon.NZ(rows(i).Item("IncentiveID"), 0) & "</td>")
          Send("    <td><a href=""offer-redirect.aspx?OfferID=" & MyCommon.NZ(rows(i).Item("IncentiveID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(Name, ""), 25) & "</a></td>")
          Send("    <td align=""center"">" & MyCommon.NZ(rows(i).Item("ValidLocations"), 0) & "</td>")
          Send("    <td align=""center"">" & MyCommon.NZ(rows(i).Item("WatchLocations"), 0) & "</td>")
          Send("    <td align=""center"" style=""" & IIf(MyCommon.NZ(rows(i).Item("WarningLocations"), 0) = 0, "", RowColor) & """>" & MyCommon.NZ(rows(i).Item("WarningLocations"), 0) & "</td>")
          Send("    <td align=""center"" style=""" & IIf(MyCommon.NZ(rows(i).Item("ValidOfferComponents"), 0), "", RowColor) & """>" & IIf(MyCommon.NZ(rows(i).Item("ValidOfferComponents"), False), Copient.PhraseLib.Lookup("term.yes", LanguageID), Copient.PhraseLib.Lookup("term.no", LanguageID)) & "</td>")
          If (Not IsDBNull(rows(i).Item("LastValidated"))) Then
            Send("    <td>" & Logix.ToShortDateTimeString(rows(i).Item("LastValidated"), MyCommon) & "</td>")
          Else
            Send("    <td>" & Copient.PhraseLib.Lookup("term.never", LanguageID) & "</td>")
          End If
          Send("    <td align=""center"">" & IIf(MyCommon.NZ(rows(i).Item("Pending"), 0), Copient.PhraseLib.Lookup("term.yes", LanguageID), Copient.PhraseLib.Lookup("term.no", LanguageID) & "&nbsp;&nbsp;[<a href=""javascript:addToPending(" & MyCommon.NZ(rows(i).Item("IncentiveID"), 0) & ");"">" & Copient.PhraseLib.Lookup("term.add", LanguageID) & "</a>]") & "</td>")
          Send("</tr>")
          If (shaded = "shaded") Then
            shaded = ""
          Else
            shaded = "shaded"
          End If
          i = i + 1
        End While
      %>
    </tbody>
  </table>
</div>
<%
done:
  Send_BodyEnd("searchform", "searchterms")
  MyCommon.Close_LogixRT()
  MyCommon = Nothing
  Logix = Nothing
%>
