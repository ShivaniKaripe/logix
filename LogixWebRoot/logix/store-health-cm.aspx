<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%@ Import Namespace="CMS.AMS.Contract" %>

<%@ Import Namespace="CMS.AMS" %>

<%@ Import Namespace="CMS.DB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: store-health-cm.aspx 
  ' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' * Copyright (c) 2002, 2003, 2004, 2005, 2006, 2007  - All rights reserved by:
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
  Dim dstStoreList As System.Data.DataTable
  Dim shaded As Boolean
  Dim tdTag As String
  Dim rowAlert As String
  Dim lastHeardAlertMins As Integer = 20
  Dim SortText As String = "loc.LocationID"
  Dim SortDirection As String = ""
  Dim idNumber As Integer = 0
		Dim idSearch As String = ""
		Dim idSearchText As String = ""
  Dim PageNum As Integer = 0
  Dim MorePages As Boolean
  Dim linesPerPage As Integer = 20
  Dim sizeOfData As Integer
  Dim i As Integer = 0
  Dim SearchClause As String = ""
  Dim SanityCheckResult As String = ""
  Dim FilterHealth As Integer = 0
  Dim IncentiveFetchResult As String = ""
  Dim IncentiveFetchFilter As String = ""
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim IsCPEInstalled As Boolean = False
  Dim CommsFilter As String = ""
  Dim BannersEnabled As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "store-health-cm.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  PageNum = Request.QueryString("pagenum")
  If PageNum < 0 Then PageNum = 0
  MorePages = False
  shaded = True
  
  Send_HeadBegin("term.storehealth")
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
  
  If (Request.QueryString("action") <> "" And Request.QueryString("locid") <> "") Then
    If (Request.QueryString("action") = "mute") Then
      MyCommon.QueryStr = "update Locations with (RowLock) Set SendAlert = 0 where LocationID = " & Request.QueryString("locid")
      MyCommon.LRT_Execute()
    ElseIf (Request.QueryString("action") = "unmute") Then
      MyCommon.QueryStr = "update Locations with (RowLock) Set SendAlert = 1 where LocationID = " & Request.QueryString("locid")
      MyCommon.LRT_Execute()
    ElseIf (Request.QueryString("action") = "disable") Then
      MyCommon.QueryStr = "update Locations with (RowLock) Set HealthReported = 0 where LocationID = " & Request.QueryString("locid")
      MyCommon.LRT_Execute()
    ElseIf (Request.QueryString("action") = "enable") Then
      MyCommon.QueryStr = "update Locations with (RowLock) Set HealthReported = 1 where LocationID = " & Request.QueryString("locid")
      MyCommon.LRT_Execute()
    End If
  End If
  
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
				SearchClause = " and (ls.LocalServerID=@idSearch or loc.LocationID=@idSearch or loc.LocationName like '%' + @idSearchText +'%' or loc.ExtLocationCode like '%'+@idSearchText+'%') "
				MyCommon.DBParameters.Add("@idSearch", SqlDbType.NVarChar).Value = idSearch
				MyCommon.DBParameters.Add("@idSearchText", SqlDbType.NVarChar).Value = idSearchText
  End If
  
  CommsFilter = "   (CASE WHEN DATEADD(n, 90, LS.CMLastHeard) >= getDate() THEN  1 ELSE 0 END ) "

  
  If (Request.QueryString("filterhealth") <> "") Then
    FilterHealth = MyCommon.Extract_Val(Request.QueryString("filterhealth"))
    Select Case FilterHealth
      Case 1 ' communications ok
        SearchClause += " and isnull(" & CommsFilter & ",1) = 1  "
      Case 2 ' all exceptions
        SearchClause += " and isnull(" & CommsFilter & ",0) = 0  and HealthReported = 1 "
      Case 5 ' failover yes
        SearchClause += " and FailoverServer = 1 "
      Case 6 ' IPL Needed
        SearchClause += " and MustIPL = 1 "
      Case 7 ' active locations only
        SearchClause += " and loc.locationID = ls.LocationID "
    End Select
  End If
  
  MyCommon.QueryStr = "select loc.LocationID, ls.LocalServerID, loc.LocationName, loc.EngineID, loc.ExtLocationCode, loc.HealthReported, PE.Description as EngineName, ls.CMLastHeard, ls.LastHeard, ls.FailoverServer, ls.MustIPL, ls.SanityCheckLastHeard, " & _
                      "IsNull(loc.SendAlert,0) SendAlert, " & CommsFilter & " as Comms " & _
                      "from Locations as loc with (nolock) " & _
                      "left join PromoEngines PE with (NoLock) on PE.EngineID=loc.EngineID " & _
                      "left join LocalServers as ls with (NoLock) on loc.LocationID = ls.LocationID " & _
                      "left join SanityCheckStatus scs with (NoLock) on loc.LocationID = scs.LocationID " & _
                      "where loc.EngineID = 0 and loc.Deleted = 0 " & SearchClause
  
  ' check if banners are enabled
  If (BannersEnabled) Then
    MyCommon.QueryStr &= " and (BannerID is Null or BannerID=0 or BannerID in (select BannerID from AdminUserBanners AUB with (NoLock) where AdminUserID = " & AdminUserID & ")) "
  End If
		MyCommon.QueryStr &= " order by " & SortText & " " & SortDirection
	
		'Send(MyCommon.QueryStr)
  'GoTo done
		dstStoreList = MyCommon.ExecuteQuery(DataBases.LogixRT)
  sizeOfData = MyCommon.NZ(dstStoreList.Rows.Count, 0)
  ' set i
  i = linesPerPage * PageNum
%>
<script runat="server">
    Protected Overrides Sub OnPreInit(e As EventArgs)
        MyBase.OnPreInit(e)
        CurrentRequest.Resolver.AppName = "offer-list.aspx"
        Dim xssEncoding As IXssEncoding = CMS.AMS.CurrentRequest.Resolver.Resolve(Of IXssEncoding)()
        Dim requestRef As HttpRequest = Request
        xssEncoding.EncodeInputParams(requestRef)
    End Sub
</script>
<script type="text/javascript" language="javascript">
    function setFilter(index) {
        var elem = document.getElementById("filterhealth");
        
        if (elem != null && index < elem.options.length) {
            elem.options[index].selected = true
        }
    }
    function loadStoreHealth() {
      var elem = document.getElementById("engine");
      var engine = 0;
      var pageUrl = 'store-health-cm.aspx';
      
      if (elem != null) {
        engine = parseInt(elem.options[elem.options.selectedIndex].value);
      }
      
      switch (engine) {
        case 0:
          pageUrl = 'store-health-cm.aspx?filterhealth=2'; 
          break;
        case 2:
          pageUrl = 'store-health-cpe.aspx?filterhealth=2';
          break;
        case 9:
            pageUrl = 'UE/store-health-UE.aspx?filterhealth=2';
          break;
        default:
          pageUrl = 'store-health-cm.aspx?filterhealth=2';
          break;
      }
        
      document.location = pageUrl;
    }
</script>

  <div id="intro">
    <%
      Send("<h1 id=""title"" style=""display:inline;"">")
      Sendb("  " & Copient.PhraseLib.Lookup("term.health", LanguageID) & " " & Copient.PhraseLib.Lookup("term.for", LanguageID).ToLower & " ")
      MyCommon.QueryStr = "select EngineID, PhraseID, DefaultEngine from PromoEngines where Installed=1 and EngineID in (0,2,9);"
      dst = MyCommon.LRT_Select
      If (dst.Rows.Count = 1) Then
        Sendb(Copient.PhraseLib.Lookup(MyCommon.NZ(dst.Rows(0).Item("PhraseID"), 0), LanguageID))
        Send("</h1>")
      ElseIf (dst.Rows.Count > 1) Then
        Send("</h1>")
        Send("<select name=""engine"" id=""engine"" onchange=""loadStoreHealth();"">")
        For Each row In dst.Rows
          Send("  <option value=""" & MyCommon.NZ(row.Item("EngineID"), -1) & """ " & IIf(MyCommon.NZ(row.Item("EngineID"), -1) = 0, "selected", "") & ">" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "</option>")
        Next
        Send("</select>")
        End If

    %>
    <form action="<%Sendb("log-view.aspx?filetype=-1&amp;fileyear=" & Year(Today) & "&amp;filemonth=" & Month(Today) & "&amp;fileday=" & Day(Today)) %>" id="controlsform" name="controlsform" target="_blank" style="float: right;">
      <div id="controls">
      <% If (Logix.UserRoles.AccessLogs) Then%>
        <input type="submit" class="regular" id="logs" name="logs" value="<% Sendb(Copient.PhraseLib.Lookup("term.logs", LanguageID)) %>..." />
      <% End If%>
      </div>
    </form>
  </div>
<div id="main">
  <br />
  <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
  <% Send_Listbar(linesPerPage, sizeOfData, PageNum, Request.QueryString("searchterms"), "&SortText=" & SortText & "&SortDirection=" & SortDirection)%>
  <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.health", LanguageID)) %>">
    <thead>
      <tr>
        <th align="left" class="th-alert" scope="col" valign="bottom">
          <a href="store-health-cm.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&action=Sort&SortText=SendAlert&SortDirection=<% Sendb(SortDirection) %>&filterhealth=<% Sendb(FilterHealth) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.alert", LanguageID))%>
          </a>
          <%If SortText = "SendAlert" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" class="th-alert" scope="col" valign="bottom">
          <a href="store-health-cm.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&action=Sort&SortText=HealthReported&SortDirection=<% Sendb(SortDirection) %>&filterhealth=<% Sendb(FilterHealth) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.report", LanguageID))%>
          </a>
          <%If SortText = "HealthReported" Then
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
          <a href="store-health-cm.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&action=Sort&SortText=loc.LocationID&SortDirection=<% Sendb(SortDirection) %>&filterhealth=<% Sendb(FilterHealth) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>
          </a>
          <%If SortText = "loc.LocationID" Then
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
          <a href="store-health-cm.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&action=Sort&SortText=ls.LocalServerID&SortDirection=<% Sendb(SortDirection) %>&filterhealth=<% Sendb(FilterHealth) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.lsid", LanguageID))%>
          </a>
          <%If SortText = "ls.LocalServerID" Then
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
          <a href="store-health-cm.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&action=Sort&SortText=loc.ExtLocationCode&SortDirection=<% Sendb(SortDirection) %>&filterhealth=<% Sendb(FilterHealth) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.code", LanguageID))%>
          </a>
          <%If SortText = "loc.ExtLocationCode" Then
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
          <a href="store-health-cm.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&action=Sort&SortText=loc.LocationName&SortDirection=<% Sendb(SortDirection) %>&filterhealth=<% Sendb(FilterHealth) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.location", LanguageID))%>
          </a>
          <%If SortText = "loc.LocationName" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th class="th-failover" scope="col" style="text-align: center;" valign="bottom">
          <a href="store-health-cm.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&action=Sort&SortText=ls.FailoverServer&SortDirection=<% Sendb(SortDirection) %>&filterhealth=<% Sendb(FilterHealth) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.failover", LanguageID))%>
          </a>
          <%If SortText = "ls.FailoverServer" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th class="th-failover" scope="col" style="text-align: center;" valign="bottom">
          <a href="store-health-cm.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&action=Sort&SortText=Comms&SortDirection=<% Sendb(SortDirection) %>&filterhealth=<% Sendb(FilterHealth) %>">
            Comm
          </a>
          <%If SortText = "Comms" Then
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
            Dim LID As Integer = row.Item("LocationID")
            
            If (MyCommon.NZ(row.Item("HealthReported"), False) AndAlso (MyCommon.NZ(row.Item("MustIPL"), True) = True OrElse MyCommon.NZ(row.Item("Comms"), 0) = 0)) Then
              'Send("<!-- MustIPL: " & MyCommon.NZ(row.Item("MustIPL"), True) & " : COMMS: " & MyCommon.NZ(row.Item("Comms"), 0) & " -->")
              Send("<tr class=""shadedred"">")
            Else
              If (shaded) Then
                Send("<tr class=""shaded"">")
                shaded = False
              Else
                Send("<tr>")
                shaded = True
              End If
            End If
                                     
            ' Send Alert column
            If (row.Item("SendAlert") = True) Then
              Send("     <td><a href=""?action=mute&locid=" & row.Item("LocationID") & "&filterhealth=" & FilterHealth & """><img src=""../images/unmute.png"" border=""0"" alt=""" & Copient.PhraseLib.Lookup("health.mute", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("health.mute", LanguageID) & """ /></a></td>")
            Else
              Send("     <td><a href=""?action=unmute&locid=" & row.Item("LocationID") & "&filterhealth=" & FilterHealth & """><img src=""../images/mute.png"" border=""0"" alt=""" & Copient.PhraseLib.Lookup("health.unmute", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("health.unmute", LanguageID) & """ /></a></td>")
            End If
            
            ' Report Health column
            If (row.Item("HealthReported") = True) Then
              Send("     <td><a href=""?action=disable&locid=" & row.Item("LocationID") & "&filterhealth=" & FilterHealth & """><img src=""../images/report-on.png"" border=""0"" alt=""" & Copient.PhraseLib.Lookup("store-health-ue.DisableReporting", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("store-health-ue.DisableReporting", LanguageID) & """ /></a></td>")
            Else
              Send("     <td><a href=""?action=enable&locid=" & row.Item("LocationID") & "&filterhealth=" & FilterHealth & """><img src=""../images/report-off.png"" border=""0"" alt=""" & Copient.PhraseLib.Lookup("store-health-ue.EnableReporting", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("store-health-ue.EnableReporting", LanguageID) & """ /></a></td>")
            End If
            
            ' Location ID, External Location Code and location name columns
            Send("     <td>" & MyCommon.NZ(row.Item("LocationID"), "&nbsp;") & "</td>")
            Send("     <td>" & MyCommon.NZ(row.Item("LocalServerID"), "&nbsp;") & "</td>")

            Send("<td>" & MyCommon.NZ(row.Item("ExtLocationCode"), "&nbsp;") & "</td>")
            If (Logix.UserRoles.AccessStoreHealth = True) Then
              Send("<td>" & " <a href=""store-detail.aspx?LocationID=" & row.Item("LocationID") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("LocationName"), "&nbsp;"), 25) & "</a></td>")
            Else
              Send("<td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("LocationName"), "&nbsp;"), 25) & "</td>")
            End If
            
            ' Failover Server column
            If (MyCommon.NZ(row.Item("FailoverServer"), False) = True) Then
              Send("     <td align=""center"">" & (Copient.PhraseLib.Lookup("term.yes", LanguageID)) & "</td>")
            Else
              Send("     <td align=""center""></td>")
            End If
            
            ' Comms column
            If (MyCommon.NZ(row.Item("Comms"), 0) = 1) Then
              Send("     <td align=""center"" class=""green"">" & (Copient.PhraseLib.Lookup("term.ok", LanguageID)) & "</td>")
            Else
              Send("     <td align=""center"" class=""red"">" & (Copient.PhraseLib.Lookup("term.failed", LanguageID)) & "</td>")
            End If
            
            Send("</tr>")
            i += 1
          End While
        Else
          Send("<tr>")
          Send("<td></td>")
          Send("</tr>")
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
