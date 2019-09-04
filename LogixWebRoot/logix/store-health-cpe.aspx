<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: store-health-cpe.aspx 
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
  Dim dst3 As System.Data.DataTable
  Dim dstStoreList As System.Data.DataTable
  Dim shaded As Boolean
  Dim tdTag As String
  Dim rowAlert As String
  Dim lastHeardAlertMins As Integer = 20
  Dim SortText As String = "SeveritySort"
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
  Dim IncentiveFetchResult As String = ""
  Dim IncentiveFetchFilter As String = ""
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim IsCPEInstalled As Boolean = False
  Dim CommsFilter As String = ""
  Dim EnterpriseCommsFilter As String = ""
  Dim BannersEnabled As Boolean = False
  Dim rows() As DataRow
  Dim SummaryText As String = ""
  Dim SeverityTypes As New Hashtable(5)
  Dim Severity As SeverityEntry
  Dim SeverityDesc As String = ""
  Dim Sev1Errs, Sev5Errs As Integer
  Dim SevPhraseID As String
  Dim CentralErrors() As Integer
  Dim CentralHighValue As Integer = 180
  Dim CentralMediumValue As Integer = 180
  Dim CentralLowValue As Integer = 90
  Dim MinutesInError As Integer = 0
  Dim QryStr As String = ""
  Dim ErrorText As String = ""
  Dim LocalServerID As Integer = -1
  Dim RowCt, Counter As Integer
  Dim ExtLocCode As String = ""
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "store-health-cpe.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixWH()
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
      SortDirection = "ASC"
    End If
  Else
    If Request.QueryString("SortDirection") <> "" Then
      SortDirection = Request.QueryString("SortDirection")
    End If
  End If
  
  ' load up the severity types
  MyCommon.QueryStr = "select HealthSeverityID, Description, PhraseID from LS_HealthSeverityTypes with (NoLock)"
  dst = MyCommon.LWH_Select
  For Each row In dst.Rows
    Severity = New SeverityEntry(MyCommon.NZ(row.Item("Description"), ""), MyCommon.NZ(row.Item("PhraseID"), 0))
    SeverityTypes.Add("Sev" & MyCommon.NZ(row.Item("HealthSeverityID"), "-1").ToString, Severity)
  Next
  
  ' load up the system options for the severity threshold levels
  If (Not Integer.TryParse(MyCommon.Fetch_CPE_SystemOption(56), CentralHighValue)) Then CentralHighValue = 270
  If (Not Integer.TryParse(MyCommon.Fetch_CPE_SystemOption(57), CentralMediumValue)) Then CentralMediumValue = 180
  If (Not Integer.TryParse(MyCommon.Fetch_CPE_SystemOption(58), CentralLowValue)) Then CentralLowValue = 90
  
  If (Request.QueryString("searchterms") <> "") Then
    idSearchText = MyCommon.Parse_Quotes(Request.QueryString("searchterms"))
    If (Integer.TryParse(idSearchText, idNumber)) Then
      idSearch = idNumber
    Else
      idSearch = -1
    End If
    If (idSearchText.IndexOf("_") > -1) Then idSearchText = idSearchText.Replace("_", "[_]")
    SearchClause = " and (ls.LocalServerID=" & idSearch & " or loc.LocationID=" & idSearch & " or loc.LocationName like '%" & idSearchText & "%' " & " or loc.ExtLocationCode like '%" & idSearchText & "%') "
  End If
  
  Dim atEnterpriseStr as String = MyCommon.Fetch_CPE_SystemOption(91).Trim
  Dim atEnterprise as Boolean = ( atEnterpriseStr.Length > 0 AndAlso atEnterpriseStr <> "0"  )

  CommsFilter = " (CASE WHEN DATEADD(n, " & CentralLowValue & ", LS.LastHeard) >= getDate() and DATEADD(n," & CentralLowValue & ", LS.IncentiveLastHeard) >=getDate() and DATEADD(n," & CentralLowValue & ",LS.TransactionLastHeard) >=getDate() and DATEADD(n," & CentralLowValue & ", LS.TransDownLoadLastHeard) >=getDate() and DBOK=1 THEN  1 ELSE 0 END) "
  
  EnterpriseCommsFilter = " (CASE WHEN DATEADD(n, " & CentralLowValue & ", LS.LastHeard) >= getDate() and DATEADD(n," & CentralLowValue & ", LS.IncentiveLastHeard) >=getDate() and DATEADD(n," & CentralLowValue & ",LS.TransactionLastHeard) >=getDate() and DBOK=1 THEN  1 ELSE 0 END) "
  
  If (Request.QueryString("filterhealth") <> "") Then
    FilterHealth = MyCommon.Extract_Val(Request.QueryString("filterhealth"))
    Select Case FilterHealth
      Case 1 ' communications ok
        SearchClause &= " and DBOK = 1 and isnull(" & IIf(atEnterprise, EnterpriseCommsFilter, CommsFilter) & ",1) = 1 "
      Case 2 ' all exceptions
        SearchClause &= " and (DBOK = 0 or isnull(" & IIf(atEnterprise, EnterpriseCommsFilter, CommsFilter) & ",0) = 0 or (ls.Sev1Errors>0 or ls.Sev10Errors>0)) and HealthReported = 1 and ls.LocalServerID is not null "
      Case 3 ' central exceptions only
        SearchClause &= "  and (DBOK = 0 or isnull(" & IIf(atEnterprise, EnterpriseCommsFilter, CommsFilter) & ",0) = 0) and HealthReported = 1 and ls.LocalServerID is not null "
      Case 4 ' local exceptions only
        SearchClause &= " and (ls.Sev1Errors > 0 or ls.Sev10Errors > 0) and HealthReported = 1 and ls.LocalServerID is not null "
      Case 5 ' failover yes
        SearchClause &= " and FailoverServer = 1 "
      Case 6
        CommsFilter = " (CASE WHEN DATEADD(n, " & CentralLowValue & ", LS.LastHeard) >= getDate() and DATEADD(n," & CentralLowValue & ", LS.IncentiveLastHeard) >=getDate() and DATEADD(n," & CentralLowValue & ",LS.TransactionLastHeard) >=getDate() and DATEADD(n," & CentralLowValue & ", LS.TransDownLoadLastHeard) >=getDate() THEN  1 ELSE 0 END) "
        EnterpriseCommsFilter = " (CASE WHEN DATEADD(n, " & CentralLowValue & ", LS.LastHeard) >= getDate() and DATEADD(n," & CentralLowValue & ", LS.IncentiveLastHeard) >=getDate() and DATEADD(n," & CentralLowValue & ",LS.TransactionLastHeard) >=getDate() THEN  1 ELSE 0 END) "
      Case 7 ' IPL Needed
        SearchClause &= " and MustIPL = 1 "
      Case 8 ' active locations only
        SearchClause &= " and loc.locationID = ls.LocationID "
    End Select
  End If
  
  If (FilterHealth = 6) Then
    MyCommon.QueryStr = "select ls.LocationID, ls.LocalServerID, '' as LocationName, 2 as EngineID, loc.ExtLocationCode, 0 as HealthReported, 'CPE' as EngineName, ls.CMLastHeard, ls.LastHeard, ls.FailoverServer, ls.MustIPL, ls.SanityCheckLastHeard," & _
                        "0 as SendAlert, 1 as SanityCheckResult," & IIf(atEnterprise, EnterpriseCommsFilter, CommsFilter) & " as Comms, " & _
                        "ls.LastRunID, ls.Sev1Errors, ls.Sev10Errors, ls.LastHeard, ls.IncentiveLastHeard, ls.TransactionLastHeard, ls.TransDownloadLastHeard, 1 as DBOK, getdate() as LastReportDate, ls.LastIP, " & _
                        "case " & _
                        "  when ls.LocalServerID > -1 and (ls.Sev1Errors > 0 or DATEADD(n, " & CentralHighValue & ", LS.LastHeard) <= getDate() or DATEADD(n," & CentralHighValue & ", LS.IncentiveLastHeard) <=getDate() or DATEADD(n," & CentralHighValue & ",LS.TransactionLastHeard) <=getDate() " & IIf(atEnterprise, "", "or DATEADD(n," & CentralHighValue & ", LS.TransDownLoadLastHeard) <=getDate()") & ") then 1 " & _
                        "  when ls.LocalServerID > -1 and (DATEADD(n," & CentralMediumValue & ", LS.LastHeard) <= getDate() or DATEADD(n," & CentralMediumValue & ", LS.IncentiveLastHeard) <=getDate() or DATEADD(n," & CentralMediumValue & ",LS.TransactionLastHeard) <=getDate() " & IIf(atEnterprise, "", "or DATEADD(n," & CentralMediumValue & ", LS.TransDownLoadLastHeard) <=getDate() ") & ") then 5 " & _
                        "  when ls.LocalServerID > -1 and (ls.Sev10Errors > 0 or DATEADD(n," & CentralLowValue & ", LS.LastHeard) <= getDate() or DATEADD(n," & CentralLowValue & ", LS.IncentiveLastHeard) <=getDate() or DATEADD(n," & CentralLowValue & ",LS.TransactionLastHeard) <=getDate()) then 10 " & _
                        "  else 0 " & _
                        "end as Severity, " & _
                        "case " & _
                        "  when ls.LocalServerID > -1 and (ls.Sev1Errors > 0 or DATEADD(n, " & CentralHighValue & ", LS.LastHeard) <= getDate() or DATEADD(n," & CentralHighValue & ", LS.IncentiveLastHeard) <=getDate() or DATEADD(n," & CentralHighValue & ",LS.TransactionLastHeard) <=getDate() " & IIf(atEnterprise, "", "or DATEADD(n," & CentralHighValue & ", LS.TransDownLoadLastHeard) <=getDate()") & ") then 1 " & _
                        "  when ls.LocalServerID > -1 and (DATEADD(n," & CentralMediumValue & ", LS.LastHeard) <= getDate() or DATEADD(n," & CentralMediumValue & ", LS.IncentiveLastHeard) <=getDate() or DATEADD(n," & CentralMediumValue & ",LS.TransactionLastHeard) <=getDate() " & IIf(atEnterprise, "", "or DATEADD(n," & CentralMediumValue & ", LS.TransDownLoadLastHeard) <=getDate() ") & ") then 5 " & _
                        "  when ls.LocalServerID > -1 and (ls.Sev10Errors > 0 or DATEADD(n," & CentralLowValue & ", LS.LastHeard) <= getDate() or DATEADD(n," & CentralLowValue & ", LS.IncentiveLastHeard) <=getDate() or DATEADD(n," & CentralLowValue & ",LS.TransactionLastHeard) <=getDate()) then 10 " & _
                        "  else 99 " & _
                        "end as SeveritySort " & _
                        "from LocalServers as ls with (NoLock) " & _
                        "left join Locations as loc with (NoLock) on loc.LocationID = ls.LocationID " & _
                        "where FailoverServer = 1 "
  Else
    MyCommon.QueryStr = "select loc.LocationID, ls.LocalServerID, loc.LocationName, loc.EngineID, loc.ExtLocationCode, loc.HealthReported, PE.Description as EngineName, ls.CMLastHeard, ls.LastHeard, ls.FailoverServer, ls.MustIPL, ls.SanityCheckLastHeard, " & _
                        "IsNull(loc.SendAlert,0) SendAlert, scs.DBOK as SanityCheckResult, " & IIf(atEnterprise, EnterpriseCommsFilter, CommsFilter) & " as Comms, " & _
                        "ls.LastRunID, ls.Sev1Errors, ls.Sev10Errors, ls.LastHeard, ls.IncentiveLastHeard, ls.TransactionLastHeard, ls.TransDownloadLastHeard, DBOK, scs.LastReportDate, ls.LastIP, " & _
                        "case " & _
                        "  when ls.LocalServerID > -1 and (ls.Sev1Errors > 0 or DATEADD(n, " & CentralHighValue & ", LS.LastHeard) <= getDate() or DATEADD(n," & CentralHighValue & ", LS.IncentiveLastHeard) <=getDate() or DATEADD(n," & CentralHighValue & ",LS.TransactionLastHeard) <=getDate() " & IIf(atEnterprise, "", "or DATEADD(n," & CentralHighValue & ", LS.TransDownLoadLastHeard) <=getDate()") & ") then 1 " & _
                        "  when ls.LocalServerID > -1 and (DATEADD(n," & CentralMediumValue & ", LS.LastHeard) <= getDate() or DATEADD(n," & CentralMediumValue & ", LS.IncentiveLastHeard) <=getDate() or DATEADD(n," & CentralMediumValue & ",LS.TransactionLastHeard) <=getDate() " & IIf(atEnterprise, "", "or DATEADD(n," & CentralMediumValue & ", LS.TransDownLoadLastHeard) <=getDate() ") & "or (DBOK=0 and DATEADD(n," & CentralMediumValue & ", scs.LastReportDate) <= getDate())) then 5 " & _
                        "  when ls.LocalServerID > -1 and (ls.Sev10Errors > 0 or DATEADD(n," & CentralLowValue & ", LS.LastHeard) <= getDate() or DATEADD(n," & CentralLowValue & ", LS.IncentiveLastHeard) <=getDate() or DATEADD(n," & CentralLowValue & ",LS.TransactionLastHeard) <=getDate() " & IIf(atEnterprise, "", "or DATEADD(n," & CentralLowValue & ", LS.TransDownLoadLastHeard) <=getDate() ") & "or (DBOK=0)) then 10 " & _
                        "  else 0 " & _
                        "end as Severity, " & _
                        "case " & _
                        "  when ls.LocalServerID > -1 and (ls.Sev1Errors > 0 or DATEADD(n, " & CentralHighValue & ", LS.LastHeard) <= getDate() or DATEADD(n," & CentralHighValue & ", LS.IncentiveLastHeard) <=getDate() or DATEADD(n," & CentralHighValue & ",LS.TransactionLastHeard) <=getDate() " & IIf(atEnterprise, "", "or DATEADD(n," & CentralHighValue & ", LS.TransDownLoadLastHeard) <=getDate()") & ") then 1 " & _
                        "  when ls.LocalServerID > -1 and (DATEADD(n," & CentralMediumValue & ", LS.LastHeard) <= getDate() or DATEADD(n," & CentralMediumValue & ", LS.IncentiveLastHeard) <=getDate() or DATEADD(n," & CentralMediumValue & ",LS.TransactionLastHeard) <=getDate() " & IIf(atEnterprise, "", "or DATEADD(n," & CentralMediumValue & ", LS.TransDownLoadLastHeard) <=getDate() ") & "or (DBOK=0 and DATEADD(n," & CentralMediumValue & ", scs.LastReportDate) <= getDate())) then 5 " & _
                        "  when ls.LocalServerID > -1 and (ls.Sev10Errors > 0 or DATEADD(n," & CentralLowValue & ", LS.LastHeard) <= getDate() or DATEADD(n," & CentralLowValue & ", LS.IncentiveLastHeard) <=getDate() or DATEADD(n," & CentralLowValue & ",LS.TransactionLastHeard) <=getDate()" & IIf(atEnterprise, "", "or DATEADD(n," & CentralLowValue & ", LS.TransDownLoadLastHeard) <=getDate() ") & "or (DBOK=0)) then 10 " & _
                        "  else 99 " & _
                        "end as SeveritySort " & _
                        "from Locations as loc with (nolock) " & _
                        "left join PromoEngines PE with (NoLock) on PE.EngineID=loc.EngineID " & _
                        "left join LocalServers as ls with (NoLock) on loc.LocationID = ls.LocationID " & _
                        "left join SanityCheckStatus scs with (NoLock) on loc.LocationID = scs.LocationID " & _
                        "where loc.EngineID = 2 and loc.Deleted=0 " & SearchClause & _
                        IIf(atEnterprise, " and LocationTypeID=2", " and LocationTypeID=1")
  End If
  If (BannersEnabled) Then
    MyCommon.QueryStr &= " and (BannerID is Null or BannerID=0 or BannerID in (select BannerID from AdminUserBanners AUB with (NoLock) where AdminUserID = " & AdminUserID & ")) "
  End If
  MyCommon.QueryStr &= " order by " & SortText & " " & SortDirection
  
  dstStoreList = MyCommon.LRT_Select
  sizeOfData = MyCommon.NZ(dstStoreList.Rows.Count, 0)
  i = linesPerPage * PageNum
%>

<script type="text/javascript" language="javascript">
    function setFilter(index) {
        var elem = document.getElementById("filterhealth");
        
        if (elem != null && index < elem.options.length) {
            elem.options[index].selected = true
        }
    }
    function loadStoreHealth() {
      var elem = document.getElementById("engine");
      var engine = 2;
      var pageUrl = 'store-health-cpe.aspx';
      
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
          pageUrl = 'store-health-cpe.aspx?filterhealth=2';
          break;
      }
        
      document.location = pageUrl;
    }
    
    function toggleRow(row) {
      var elemRow = document.getElementById('details' + row);
      var elemImg = document.getElementById('img' + row);
      
      if (elemRow != null) {
        elemRow.style.display = (elemRow.style.display != '') ? '' : 'none';
        if (elemImg != null) {
          elemImg.src = (elemRow.style.display == '') ? '../images/minus2.png' : '../images/plus2.png';
        }
      }

    }
</script>

<div id="intro">
  <%
    Send("<h1 id=""title"" style=""display:inline;"">")
    Sendb("  " & Copient.PhraseLib.Lookup("term.health", LanguageID) & " " & Copient.PhraseLib.Lookup("term.for", LanguageID).ToLower & " ")
    MyCommon.QueryStr = "select EngineID, PhraseID, DefaultEngine from PromoEngines where Installed=1 and EngineID in (0,2,9);"
    dst = MyCommon.LRT_Select
    If (dst.Rows.Count = 1) Then
      Send(Copient.PhraseLib.Lookup(MyCommon.NZ(dst.Rows(0).Item("PhraseID"), 0), LanguageID))
      Send("</h1>")
    ElseIf (dst.Rows.Count > 1) Then
      Send("</h1>")
      Send("<select name=""engine"" id=""engine"" onchange=""loadStoreHealth();"">")
      For Each row In dst.Rows
        Send("  <option value=""" & MyCommon.NZ(row.Item("EngineID"), -1) & """ " & IIf(MyCommon.NZ(row.Item("EngineID"), -1) = 2, "selected", "") & ">" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "</option>")
      Next
      Send("</select>")
    End If
  %>
  <div id="controls" style="width:250px;">
    <% If (Logix.UserRoles.AccessLogs) Then%>
    <form action="<%Sendb("log-view.aspx?filetype=-1&amp;fileyear=" & Year(Today) & "&amp;filemonth=" & Month(Today) & "&amp;fileday=" & Day(Today)) %>" id="controlsform" name="controlsform" target="_blank" style="float: right;">
      <input type="submit" class="regular" id="logs" name="logs" value="<% Sendb(Copient.PhraseLib.Lookup("term.logs", LanguageID)) %>..." />
    </form>
    <% End If %>
  </div>
</div>

<div id="main">
  <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
  <br class="half" />
  <% Send_Listbar(linesPerPage, sizeOfData, PageNum, Request.QueryString("searchterms"), "&SortText=" & SortText & "&SortDirection=" & SortDirection)%>
  <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.health", LanguageID)) %>">
    <thead>
      <tr>
        <th align="center" style="width:20px;" scope="col" valign="bottom">&nbsp;
        </th>
        <th align="left" class="th-code" scope="col" valign="bottom">
          <a href="store-health-cpe.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&action=Sort&SortText=loc.ExtLocationCode&SortDirection=<% Sendb(SortDirection) %>&filterhealth=<% Sendb(FilterHealth) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.code", LanguageID))%>
          </a>
          <%
            If SortText = "loc.ExtLocationfCode" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <% If (FilterHealth <> 6) Then%>
        <th align="left" class="th-id" scope="col" valign="bottom">
          <a href="store-health-cpe.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&action=Sort&SortText=loc.LocationID&SortDirection=<% Sendb(SortDirection) %>&filterhealth=<% Sendb(FilterHealth) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>
          </a>
          <%
            If SortText = "loc.LocationID" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <% Else %>
        <th align="left" class="th-lastip" scope="col" valign="bottom">
          <a href="store-health-cpe.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&action=Sort&SortText=LastIP&SortDirection=<% Sendb(SortDirection) %>&filterhealth=<% Sendb(FilterHealth) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.lastip", LanguageID))%>
          </a>
          <%
            If SortText = "LastIP" Then
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
        <th align="left" class="th-id" scope="col" valign="bottom">
          <a href="store-health-cpe.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&action=Sort&SortText=ls.LocalServerID&SortDirection=<% Sendb(SortDirection) %>&filterhealth=<% Sendb(FilterHealth) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.lsid", LanguageID))%>
          </a>
          <%
            If SortText = "ls.LocalServerID" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" class="th-severity" scope="col" valign="bottom">
          <a href="store-health-cpe.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&action=Sort&SortText=SeveritySort&SortDirection=<% Sendb(SortDirection) %>&filterhealth=<% Sendb(FilterHealth) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.severity", LanguageID))%>
          </a>
          <%
            If SortText = "SeveritySort" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" class="th-errors" scope="col" valign="bottom">
          <a href="store-health-cpe.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&action=Sort&SortText=Sev1Errors&SortDirection=<% Sendb(SortDirection) %>&filterhealth=<% Sendb(FilterHealth) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.errors", LanguageID))%>
          </a>
          <%
            If SortText = "Sev1Errors" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th class="th-failover" scope="col" style="text-align: center; display: none;" valign="bottom">
          <a href="store-health-cpe.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&action=Sort&SortText=ls.FailoverServer&SortDirection=<% Sendb(SortDirection) %>&filterhealth=<% Sendb(FilterHealth) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.failover", LanguageID))%>
          </a>
          <%
            If SortText = "ls.FailoverServer" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th class="th-failover" scope="col" style="text-align: center; display: none;" valign="bottom">
          <a href="store-health-cpe.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&action=Sort&SortText=Comms&SortDirection=<% Sendb(SortDirection) %>&filterhealth=<% Sendb(FilterHealth) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.communications", LanguageID))%>
          </a>
          <%
            If SortText = "Comms" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <% If (FilterHealth <> 6) then %>
        <th align="center" class="th-id" scope="col" valign="bottom">
          <a href="store-health-cpe.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&action=Sort&SortText=HealthReported&SortDirection=<% Sendb(SortDirection) %>&filterhealth=<% Sendb(FilterHealth) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.report", LanguageID))%>
          </a>
          <%
            If SortText = "HealthReported" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="center" class="th-alert" scope="col" valign="bottom">
          <a href="store-health-cpe.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&action=Sort&SortText=SendAlert&SortDirection=<% Sendb(SortDirection) %>&filterhealth=<% Sendb(FilterHealth) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.alert", LanguageID))%>
          </a>
          <%
            If SortText = "SendAlert" Then
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
            Dim LID As Integer = MyCommon.NZ(row.Item("LocationID"), -1)
            LocalServerID = MyCommon.NZ(row.Item("LocalServerID"), -1)
            
            If (MyCommon.NZ(row.Item("HealthReported"), False) AndAlso (LocalServerID > -1) AndAlso (MyCommon.NZ(row.Item("MustIPL"), False) = True OrElse MyCommon.NZ(row.Item("Comms"), 0) = 0 OrElse MyCommon.NZ(row.Item("Severity"), -1) > 0)) Then
              If (MyCommon.NZ(row.Item("Severity"), -1) = 1) Then
                Send("<tr class=""shadeddarkred"">")
              ElseIf (MyCommon.NZ(row.Item("Severity"), -1) = 5) Then
                Send("<tr class=""shadedred"">")
              Else
                Send("<tr class=""shadedlightred"">")
              End If
            Else
              If (shaded) Then
                Send("<tr class=""shaded"">")
                shaded = False
              Else
                Send("<tr>")
                shaded = True
              End If
            End If
            
            ' Details button
            If (MyCommon.NZ(row.Item("Severity"), -1) > 0) Then
              Send("<td align=""center"">")
              Send("<a href=""javascript:toggleRow(" & i & ");""><img id=""img" & i & """ src=""../images/plus2.png"" border=""0"" alt=""" & Copient.PhraseLib.Lookup("store-health-ue.ViewHideErrorDetails", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("store-health-ue.ViewHideErrorDetails", LanguageID) & """ /></a>")
              Send("</td>")
            Else
              Send("<td align=""center"">")
              Send("<img src=""../images/plus2-disabled.png"" border=""0"" alt="""" />")
              Send("</td>")
            End If
            
            
            ' Location ID, External Location Code and location name columns
            ExtLocCode = MyCommon.NZ(row.Item("ExtLocationCode"), "")
            If (Logix.UserRoles.AccessStoreHealth = True) Then
              Sendb("<td>")
              If (ExtLocCode = "") Then
                Sendb(Copient.PhraseLib.Lookup("term.inactive", LanguageID))
              Else
                Sendb(" <a href=""store-detail.aspx?LocationID=" & row.Item("LocationID") & """>" & ExtLocCode & "</a>")
              End If
              Send("</td>")
            Else
              Send("<td>" & IIf(ExtLocCode <> "", ExtLocCode, Copient.PhraseLib.Lookup("term.inactive", LanguageID)) & "</td>")
            End If
            
            If (FilterHealth <> 6) Then
              Send("     <td>" & MyCommon.NZ(row.Item("LocationID"), "&nbsp;") & "</td>")
            Else
              Send("     <td>" & MyCommon.NZ(row.Item("LastIP"), "&nbsp;") & "</td>")
            End If
            
            Send("     <td>" & MyCommon.NZ(row.Item("LocalServerID"), "&nbsp;") & "</td>")
            
            ' Severity            
            If (LocalServerID > -1 AndAlso MyCommon.NZ(row.Item("Severity"), -1) > 0) Then
              If (SeverityTypes.Contains("Sev" & MyCommon.NZ(row.Item("Severity"), "-1").ToString)) Then
                Severity = SeverityTypes.Item("Sev" & MyCommon.NZ(row.Item("Severity"), "-1"))
                If (Severity IsNot Nothing) Then
                  Send("     <td>" & Copient.PhraseLib.Lookup(Severity.PhraseID, LanguageID, Severity.Description) & "</td>")
                Else
                  Send("     <td>" & MyCommon.NZ(row.Item("Severity"), "-1") & "</td>")
                End If
              Else
                Send("     <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
              End If
            Else
              Send("     <td>&nbsp;</td>")
            End If
            
            ' Error text
            If (LocalServerID > -1 AndAlso MyCommon.NZ(row.Item("Severity"), -1) > 0) Then
              SeverityDesc = ""
              'local server errors
              Sev1Errs = MyCommon.NZ(row.Item("Sev1Errors"), 0)
              Sev5Errs = MyCommon.NZ(row.Item("Sev10Errors"), 0)
              SeverityDesc = GetSeverityDesc(Sev1Errs, 0, Sev5Errs, "term.local", SeverityTypes)
              
              ' central server errors
              CentralErrors = GetCentralErrors(row, CentralHighValue, CentralMediumValue, CentralLowValue)
              If (CentralErrors(0) > 0 OrElse CentralErrors(1) > 0 OrElse CentralErrors(2) > 0) Then
                If (SeverityDesc <> "") Then SeverityDesc &= "<br />"
                SeverityDesc &= GetSeverityDesc(CentralErrors(0), CentralErrors(1), CentralErrors(2), "term.central", SeverityTypes)
              End If
              
              Send("     <td>" & SeverityDesc & "</td>")
            Else
              Send("     <td>&nbsp;</td>")
            End If
            
            ' Failover Server column
            If (MyCommon.NZ(row.Item("FailoverServer"), False) = True) Then
              Send("     <td align=""center"" style=""display:none;"">" & (Copient.PhraseLib.Lookup("term.yes", LanguageID)) & "</td>")
            Else
              Send("     <td align=""center"" style=""display:none;""></td>")
            End If
            
            ' Comms column
            If (MyCommon.NZ(row.Item("Comms"), 0) = 1) Then
              Send("     <td align=""center"" class=""green"" style=""display:none;"">" & (Copient.PhraseLib.Lookup("term.ok", LanguageID)) & "</td>")
            Else
              Send("     <td align=""center"" class=""red"" style=""display:none;"">" & (Copient.PhraseLib.Lookup("term.failed", LanguageID)) & "</td>")
            End If

            If (FilterHealth <> 6) Then
              ' Report Health column
              If (row.Item("HealthReported") = True) Then
                Send("     <td align=""center""><a href=""?action=disable&locid=" & row.Item("LocationID") & "&filterhealth=" & FilterHealth & """><img src=""../images/report-on.png"" border=""0"" alt=""" & Copient.PhraseLib.Lookup("store-health-ue.DisableReporting", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("store-health-ue.DisableReporting", LanguageID) & """ /></a></td>")
              Else
                Send("     <td align=""center""><a href=""?action=enable&locid=" & row.Item("LocationID") & "&filterhealth=" & FilterHealth & """><img src=""../images/report-off.png"" border=""0"" alt=""" & Copient.PhraseLib.Lookup("store-health-ue.EnableReporting", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("store-health-ue.EnableReporting", LanguageID) & """ /></a></td>")
              End If
            
              ' Send Alert column
              If (row.Item("SendAlert") = True) Then
                Send("     <td align=""center""><a href=""?action=mute&locid=" & row.Item("LocationID") & """><img src=""../images/unmute.png"" border=""0"" alt=""" & Copient.PhraseLib.Lookup("health.mute", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("health.mute", LanguageID) & """ /></a></td>")
              Else
                Send("     <td align=""center""><a href=""?action=unmute&locid=" & row.Item("LocationID") & """><img src=""../images/mute.png"" border=""0"" alt=""" & Copient.PhraseLib.Lookup("health.unmute", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("health.unmute", LanguageID) & """ /></a></td>")
              End If
            End If
            
            Send("</tr>")
            
            If (LocalServerID > -1 AndAlso MyCommon.NZ(row.Item("Severity"), -1) > 0) Then
              Send("<tr id=""details" & i & """ style=""display:none;position:relative;top:-2px;"">")
              Send("<td class=""errordetails"" colspan=""8"">")
              MyCommon.QueryStr = "select distinct 'term.local' as ServerType, 1 as ServerTypeID, HE.HealthSeverityID, 'LS' + CONVERT(nvarchar(6),HE.ErrorID) as ErrorCode, " & _
                                  "HE.ErrorID, 0 as MinutesInError, HE.ErrorText, HT.TagName, HS.SectionName from LS_HealthErrors HE with (NoLock) " & _
                                  "left join HealthTags HT with (NoLock) on HT.TagID = HE.TagID and HT.ServerTypeID=1 " & _
                                  "left join HealthSections HS with (NoLock) on HS.SectionID = HE.SectionID " & _
                                  "where LocalServerID=" & MyCommon.NZ(row.Item("LocalServerID"), -1) & _
                                  " and RunID=" & MyCommon.NZ(row.Item("LastRunID"), -1) & " order by HE.HealthSeverityID;"
              dst2 = MyCommon.LWH_Select
              
              MyCommon.QueryStr = "dbo.pa_StoreHealth_CentralCommErrors"
              MyCommon.Open_LRTsp()
              MyCommon.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = MyCommon.NZ(row.Item("LocalServerID"), -1)
              MyCommon.LRTsp.Parameters.Add("@High", SqlDbType.Int).Value = CentralHighValue
              MyCommon.LRTsp.Parameters.Add("@Medium", SqlDbType.Int).Value = CentralMediumValue
              MyCommon.LRTsp.Parameters.Add("@Low", SqlDbType.Int).Value = CentralLowValue
              dst3 = MyCommon.LRTsp_select
              MyCommon.Close_LRTsp()
              dst2.Merge(dst3)
              
              Send("  <table style=""width:100%"" summary=""" & Copient.PhraseLib.Lookup("term.errors", LanguageID) & """>")
              Send("    <tr>")
              Send("      <th>" & Copient.PhraseLib.Lookup("term.severity", LanguageID) & "</th>")
              Send("      <th>" & Copient.PhraseLib.Lookup("term.from", LanguageID) & "</th>")
              Send("      <th>" & Copient.PhraseLib.Lookup("term.id", LanguageID) & "</th>")
              Send("      <th>" & Copient.PhraseLib.Lookup("term.description", LanguageID) & "</th>")
              Send("      <th style=""width:180px;"">" & Copient.PhraseLib.Lookup("term.duration", LanguageID) & "</th>")
              Send("    </tr>")
              For Each row2 In dst2.Rows
                If (MyCommon.NZ(row2.Item("ServerType"), "term.central") = "term.local") Then
                  ' this is a local server error so we need to find out the duration of the error
                  MyCommon.QueryStr = "dbo.pa_StoreHealth_ErrorDuration"
                  MyCommon.Open_LWHsp()
                  MyCommon.LWHsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = MyCommon.NZ(row.Item("LocalServerID"), -1)
                  MyCommon.LWHsp.Parameters.Add("@ErrorID", SqlDbType.Int).Value = MyCommon.NZ(row2.Item("ErrorID"), -1)
                  MyCommon.LWHsp.Parameters.Add("@MinutesInError", SqlDbType.Int).Direction = ParameterDirection.Output
                  MyCommon.LWHsp.ExecuteNonQuery()
                  MinutesInError = MyCommon.LWHsp.Parameters("@MinutesInError").Value
                  MyCommon.Close_LWHsp()
                  row2.Item("MinutesInError") = MinutesInError
                End If
              Next
              
              rows = dst2.Select("", "HealthSeverityID asc, MinutesInError desc")
              RowCt = rows.Length
              Counter = 1
              For Each row2 In rows
                MinutesInError = MyCommon.NZ(row2.Item("MinutesInError"), 0)
                Send("    <tr>")
                Send("      <td>" & GetSeverityText(MyCommon.NZ(row2.Item("HealthSeverityID"), -1), SeverityTypes) & "</td>")
                Send("      <td>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row2.Item("ServerType"), ""), LanguageID) & "</td>")
                
                QryStr = "?SrvType=" & MyCommon.NZ(row2.Item("ServerTypeID"), "2") & "&Err=" & MyCommon.NZ(row2.Item("ErrorID"), 0)
                Sendb("      <td><a href=""javascript:openPopup('health-resolutions.aspx" & QryStr & "');"" ")
                Sendb(" title=""" & Copient.PhraseLib.Lookup("store-health.resolution-note", LanguageID) & """")
                
                Send(">" & MyCommon.NZ(row2.Item("ErrorCode"), "&nbsp;") & "</a></td>")

                ErrorText = ""
                If (MyCommon.NZ(row2.Item("TagName"), "") <> "") Then
                  ErrorText = "[" & MyCommon.NZ(row2.Item("TagName"), "") & ":" & MyCommon.NZ(row2.Item("SectionName"), "") & "] "
                End If
                ErrorText &= MyCommon.NZ(row2.Item("ErrorText"), "&nbsp;")
                
                Send("      <td>" & ErrorText & "</td>")
                Send("      <td>" & GetDurationText(MinutesInError) & "</td>")
                Send("    </tr>")

                If (Counter < RowCt) Then
                  Send("    <tr>")
                  Send("      <td colspan=""5"" style=""background-color: #cccccc;height: 1px;padding: 0;margin: 0;""></td>")
                  Send("    </tr>")
                End If
                Counter += 1
              Next
              Send("  </table>")

              Send("</td>")
              Send("</tr>")
            End If
            
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

<script type="text/javascript" language="javascript">
    setFilter(<%Sendb(FilterHealth) %>);
</script>

<script runat="server">
  Public Class SeverityEntry
    Public Description As String
    Public PhraseID As Integer
    Sub New(ByVal SevDesc As String, ByVal SevPhrase As Integer)
      Description = SevDesc
      PhraseID = SevPhrase
    End Sub
  End Class
  
  Function GetDurationText(ByVal MinutesInError As Integer) As String
    Dim DurationText As New StringBuilder()
    Dim DurSpan As New TimeSpan(MinutesInError * TimeSpan.TicksPerMinute)
    Dim ErrDays, ErrHours, ErrMinutes As Integer

    ErrDays = DurSpan.Days
    ErrHours = DurSpan.Hours
    ErrMinutes = DurSpan.Minutes

    If (ErrDays > 0) Then DurationText.Append(ErrDays & " " & Copient.PhraseLib.Lookup("term.days", LanguageID) & ", ")
    If (ErrHours > 0 OrElse ErrDays > 0) Then DurationText.Append(ErrHours & " " & Copient.PhraseLib.Lookup("term.hours", LanguageID) & ", ")
    If (ErrMinutes > 0 OrElse ErrHours > 0 OrElse ErrDays > 0) Then DurationText.Append(ErrMinutes & " " & Copient.PhraseLib.Lookup("term.minutes", LanguageID))
    
    Return DurationText.ToString
  End Function
    
    
  Function GetSeverityText(ByVal SeverityID As Integer, ByVal SeverityTypes As Hashtable) As String
    Dim SeverityText As String = ""
    Dim Severity As SeverityEntry
    
    Severity = SeverityTypes.Item("Sev" & SeverityID)
    If (Severity IsNot Nothing) Then
      SeverityText = Copient.PhraseLib.Lookup(Severity.PhraseID, LanguageID, Severity.Description)
    Else
      SeverityText = Copient.PhraseLib.Lookup("term.unknown", LanguageID)
    End If
    
    Return SeverityText
  End Function
  

  ' NOTE: Sev1Errs = High Severity; Sev5Errs = Medium Severity; Sev10Errs = Low Severity
  Function GetSeverityDesc(ByVal Sev1Errs As Integer, ByVal Sev5Errs As Integer, ByVal Sev10Errs As Integer, ByVal PhraseName As String, _
                           ByVal SeverityTypes As Hashtable) As String
    Dim SeverityDesc As String = ""
    Dim SevPhraseName As String = ""
    Dim Severity As SeverityEntry
    
    If (Sev1Errs > 0 OrElse Sev5Errs > 0 OrElse Sev10Errs > 0) Then
      SeverityDesc &= "[" & Copient.PhraseLib.Lookup(PhraseName, LanguageID) & "] "
    End If
              
    If (Sev1Errs > 0) Then ' high severity
      SevPhraseName = IIf(Sev1Errs = 1, "store-health.error-of", "store-health.errors-of")
      SeverityDesc &= Sev1Errs & " " & Copient.PhraseLib.Lookup(SevPhraseName, LanguageID) & " "
      Severity = SeverityTypes.Item("Sev1")
      If (Severity IsNot Nothing) Then
        SeverityDesc &= Copient.PhraseLib.Lookup(Severity.PhraseID, LanguageID, Severity.Description) & " "
      End If
      SeverityDesc &= Copient.PhraseLib.Lookup("term.severity", LanguageID).ToLower
      SeverityDesc &= IIf(Sev5Errs > 0 Or Sev10Errs > 0, "; ", "")
    End If
              
    If (Sev5Errs > 0) Then ' medium severity
      SevPhraseName = IIf(Sev5Errs = 1, "store-health.error-of", "store-health.errors-of") & " "
      SeverityDesc &= Sev5Errs & " " & Copient.PhraseLib.Lookup(SevPhraseName, LanguageID) & " "
      Severity = SeverityTypes.Item("Sev5")
      If (Severity IsNot Nothing) Then
        SeverityDesc &= Copient.PhraseLib.Lookup(Severity.PhraseID, LanguageID, Severity.Description) & " "
      End If
      SeverityDesc &= Copient.PhraseLib.Lookup("term.severity", LanguageID).ToLower
      SeverityDesc &= IIf(Sev10Errs > 0, "; ", "")
    End If
    
    If (Sev10Errs > 0) Then ' low severity
      SevPhraseName = IIf(Sev10Errs = 1, "store-health.error-of", "store-health.errors-of") & " "
      SeverityDesc &= Sev10Errs & " " & Copient.PhraseLib.Lookup(SevPhraseName, LanguageID) & " "
      Severity = SeverityTypes.Item("Sev10")
      If (Severity IsNot Nothing) Then
        SeverityDesc &= Copient.PhraseLib.Lookup(Severity.PhraseID, LanguageID, Severity.Description) & " "
      End If
      SeverityDesc &= Copient.PhraseLib.Lookup("term.severity", LanguageID).ToLower
    End If
    
    Return SeverityDesc
  End Function
  
  Function GetCentralErrors(ByVal row As DataRow, ByVal HighValue As Integer, ByVal MediumValue As Integer, ByVal LowValue As Integer) As Integer()
    Dim HighCount, MediumCount, LowCount As Integer
    Dim Counts(3) As Integer
    Dim ErrorCount(3) As Integer
    Dim i As Integer
    Dim ColNames() As String = {"LastHeard", "IncentiveLastHeard", "TransactionLastHeard", "TransDownloadLastHeard"}
    
    For i = 0 To ColNames.GetUpperBound(0)
      Counts = GetFieldSevCount(row, ColNames(i), HighValue, MediumValue, LowValue)
      HighCount += Counts(0)
      MediumCount += Counts(1)
      LowCount += Counts(2)
    Next
    
    ' problems with sanity check elevate based on how long since it was reported
    If (Not IsDBNull(row.Item("DBOK")) AndAlso Not row.Item("DBOK")) Then
      Counts = GetFieldSevCount(row, "LastReportDate", HighValue, MediumValue, LowValue)
      ' only elevate Sanity Check to Medium Severity
      If (Counts(0) > 0 OrElse Counts(1) > 0) Then
        MediumCount += Counts(0) + Counts(1)
      End If
      LowCount += Counts(2)
    End If

    ErrorCount(0) = HighCount
    ErrorCount(1) = MediumCount
    ErrorCount(2) = LowCount

    Return ErrorCount
  End Function
  
  Function GetFieldSevCount(ByVal row As DataRow, ByVal FieldName As String, ByVal HighValue As Integer, ByVal MediumValue As Integer, ByVal LowValue As Integer) As Integer()
    Dim Counts(3) As Integer
    Dim TempDate As Date
    Dim span As TimeSpan
    Dim TotalMinutes As Double
    
    If (Not IsDBNull(row.Item(FieldName)) AndAlso Date.TryParse(row.Item(FieldName).ToString, TempDate)) Then
      span = Now.Subtract(TempDate)
      TotalMinutes = span.TotalMinutes
      Select Case TotalMinutes
        Case Is >= HighValue
          Counts(0) = 1
        Case Is >= MediumValue
          Counts(1) = 1
        Case Is >= LowValue
          Counts(2) = 1
      End Select
    Else
      Counts(0) = 1
    End If
    Return Counts
  End Function
</script>
<%
  'If MyCommon.Fetch_SystemOption(75) Then
  '  If (Logix.UserRoles.AccessNotes) Then
  '    Send_Notes(15, 0, AdminUserID)
  '  End If
  'End If
done:
  Send_BodyEnd("searchform", "searchterms")
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixWH()
  MyCommon = Nothing
  Logix = Nothing
%>
