<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: user-hist.aspx 
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
  Dim rst As DataTable
  Dim dst As System.Data.DataTable
  Dim dst2 As DataTable
  Dim row As System.Data.DataRow
  Dim sizeOfData As Integer
  Dim i As Integer = 0
  Dim maxEntries As Integer = 500
  Dim Username As String
  Dim UserID As Long
  Dim ActivityDate As New DateTime
  Dim Shaded As String = "shaded"
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim TempURL As String = ""
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "user-hist.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  If MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER) Then
    MyCommon.Open_PrefManRT()
  End If
  
  UserID = Request.QueryString("UserID")
  ' Check in case it was a POST instead of get
  If (AdminUserID = 0 And Not Request.QueryString("save") <> "") Then
    UserID = Request.Form("UserID")
  End If
  
  
  Send_HeadBegin("term.user", "term.history", UserID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
%>

<script type="text/javascript" language="javascript">
      function launchHierarchy(url) {
          var popW = 700;
          var popH = 570;
          
          lhierWindow = window.open(url,"hierTree", "width="+popW+", height="+popH+", top="+calcTop(popH)+", left="+calcLeft(popW)+", toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=yes");
          lhierWindow.focus();
      }
</script>

<%  
  Send_HeadEnd()
  Send_BodyBegin(3)
  
  If (Logix.UserRoles.ViewOthersInfo = False AndAlso AdminUserID <> UserID) Then
    Send_Denied(2, "perm.admin-users-seeothers")
    GoTo done
  ElseIf (Logix.UserRoles.ViewHistory = False) Then
    Send_Denied(2, "perm.admin-history")
    GoTo done
  End If
  
  MyCommon.QueryStr = "select FirstName, LastName from AdminUsers with (NoLock) Where AdminUserID='" & UserID & "';"
  dst = MyCommon.LRT_Select
  If (dst.Rows.Count > 0) Then
    Send("<div id=""intro"">")
    Send("  <h1 id=""title"">" & Copient.PhraseLib.Lookup("term.activity", LanguageID) & ": " & dst.Rows(i).Item("FirstName") & " " & dst.Rows(i).Item("LastName") & "</h1>")
    Send("</div>")
    Send("")
    Send("<div id=""main"">")
    If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
    Send("<div id=""column"">")
    Send("")
    Send("<table class=""list"" style=""width: 100%;"" summary=""" & Copient.PhraseLib.Lookup("term.history", LanguageID) & """>")
    Send("<thead>")
    Send("  <tr>")
    Send("    <th align=""left"" scope=""col"" class=""th-timedate"">" & Copient.PhraseLib.Lookup("term.timedate", LanguageID) & "</th>")
    Send("    <th align=""left"" scope=""col"" class=""th-action"">" & Copient.PhraseLib.Lookup("term.action", LanguageID) & "</th>")
    Send("    <th align=""left"" scope=""col"" class=""th-object"">" & Copient.PhraseLib.Lookup("term.object", LanguageID) & "</th>")
    Send("  </tr>")
    Send("</thead>")
    Send("<tbody>")
  Else
    Send("<div id=""intro"">")
    Send("  <h1 id=""title"">" & Copient.PhraseLib.Lookup("term.user", LanguageID) & " #" & UserID & "</h1>")
    Send("</div>")
    Send("")
    Send("<div id=""main"">")
    Send("  <div id=""infobar"" class=""red-background"">")
    Send("    " & Copient.PhraseLib.Lookup("error.deleted", LanguageID))
    Send("  </div>")
    Send("</div>")
    GoTo done
  End If
  
  MyCommon.QueryStr = "select isnull(AL.ActivityTypeID, 0) as ActivityTypeID, AU.FirstName, AU.LastName, AL.ActivityDate, AL.Description, ActT.Name as ActivityTypeName, AL.LinkID " & _
                      "from ActivityLog as AL with (NoLock) " & _
                      "left join AdminUsers as AU with (NoLock) on AL.AdminID=AU.AdminUserID " & _
                      "left join ActivityTypes as ActT with (NoLock) on ActT.ActivityTypeID=AL.ActivityTypeID " & _
                      "where AdminUserID=" & UserID & " and Name Is Not NULL order by ActivityDate desc;"
  dst = MyCommon.LRT_Select
  sizeOfData = dst.Rows.Count
  While (i < sizeOfData And i < maxEntries)
    Send("  <tr class=""" & Shaded & """>")
    If (Not IsDBNull(dst.Rows(i).Item("ActivityDate"))) Then
      Send("    <td>" & Logix.ToShortDateTimeString(dst.Rows(i).Item("ActivityDate"), MyCommon) & "</td>")
    Else
      Send("    <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
    End If
    Send("    <td>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</td>")
    
    If MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER) AndAlso dst.Rows(i).Item("ActivityTypeID") = 100003 Then 'EPM Connectors
      Send("    <td><a href=""" & MyCommon.Build_EPM_URI("UI/connector-detail.aspx?ConnectorID=" & dst.Rows(i).Item("LinkID")) & """ onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.connector", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</a></td>")
    ElseIf MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER) AndAlso dst.Rows(i).Item("ActivityTypeID") = 100004 Then 'EPM Users (Admin Users)
      Send("    <td><a href=""" & MyCommon.Build_EPM_URI("UI/user-edit.aspx?UserID=" & dst.Rows(i).Item("LinkID")) & """ onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.user", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</a></td>")
    ElseIf MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER) AndAlso dst.Rows(i).Item("ActivityTypeID") = 100005 Then 'EPM Agents
      Send("    <td><a href=""" & MyCommon.Build_EPM_URI("UI/agent-detail.aspx?appid=" & dst.Rows(i).Item("LinkID")) & """ onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.agent", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</a></td>")
    ElseIf MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER) AndAlso dst.Rows(i).Item("ActivityTypeID") = 100006 Then 'EPM Settings
      Send("    <td><a href=""" & MyCommon.Build_EPM_URI("UI/settings.aspx") & """ onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.settings", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</a></td>")
    ElseIf MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER) AndAlso dst.Rows(i).Item("ActivityTypeID") = 100007 Then 'EPM Roles
      Send("    <td><a href=""" & MyCommon.Build_EPM_URI("UI/roles.aspx") & """ onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.roles", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</a></td>")
    ElseIf MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER) AndAlso dst.Rows(i).Item("ActivityTypeID") = 100008 Then 'EPM Themes
      Send("    <td><a href=""" & MyCommon.Build_EPM_URI("UI/prefstheme-edit.aspx?themeid=" & dst.Rows(i).Item("LinkID")) & """ onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.preference", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</a></td>")
    ElseIf MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER) AndAlso dst.Rows(i).Item("ActivityTypeID") = 100009 Then 'EPM Preferences
      TempURL = "UI/prefsstd-list.aspx"
      'we need to find out if this is a system or custom preference
      MyCommon.QueryStr = "select UserCreated from Preferences where PreferenceID=" & MyCommon.Extract_Val(dst.Rows(i).Item("LinkID")) & ";"
      dst2 = MyCommon.PMRT_Select
      If dst2.Rows.Count > 0 Then
        If dst2.Rows(0).Item("UserCreated") Then
          'this is a custom preference
          TempURL = "UI/prefscustom-edit.aspx?prefid=" & dst.Rows(i).Item("LinkID")
        Else
          'this is a system preference
          TempURL = "UI/prefsstd-edit.aspx?prefid=" & dst.Rows(i).Item("LinkID")
        End If
      End If
      dst2 = Nothing
      Send("    <td><a href=""" & MyCommon.Build_EPM_URI(TempURL) & """ onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.preference", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Logged on") Then
      Send("    <td>&nbsp;</td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Logged out") Then
      Send("    <td>&nbsp;</td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Offer") Then
      Send("    <td><a href=""offer-redirect.aspx?OfferID=" & dst.Rows(i).Item("LinkID") & """ onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Customer Group") Then
      Send("    <td><a href=""cgroup-edit.aspx?CustomerGroupID=" & dst.Rows(i).Item("LinkID") & """ onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.customergroup", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Product Group") Then
      Send("    <td><a href=""pgroup-edit.aspx?ProductGroupID=" & dst.Rows(i).Item("LinkID") & """ onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.productgroup", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Points Program") Then
      Send("    <td><a href=""point-edit.aspx?ProgramGroupID=" & dst.Rows(i).Item("LinkID") & """ onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.pointsprogram", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Location") Then
      Send("    <td><a href=""store-edit.aspx?LocationID=" & dst.Rows(i).Item("LinkID") & """ onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.store", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Location Group") Then
      Send("    <td><a href=""lgroup-edit.aspx?LocationGroupID=" & dst.Rows(i).Item("LinkID") & """ onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.storegroup", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Graphic") Then
      Send("    <td><a href=""graphic-edit.aspx?OnScreenAdID=" & dst.Rows(i).Item("LinkID") & """ onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.graphic", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Screen Layout") Then
      Send("    <td><a href=""layout-edit.aspx?LayoutID=" & dst.Rows(i).Item("LinkID") & """ onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.layout", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Categories") Then
      Send("    <td><a href=""categories.aspx"" onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.categories", LanguageID) & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Departments") Then
      Send("    <td><a href=""department-edit.aspx?DeptID=" & dst.Rows(i).Item("LinkID") & """ onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.department", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Product Hierarchy") Then
      Send("    <td><a href=""javascript:launchHierarchy('phierarchytree.aspx');"" onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.producthierarchies", LanguageID) & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Store Hierarchy") Then
      Send("    <td><a href=""javascript:launchHierarchy('lhierarchytree.aspx');"" onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.storehierarchies", LanguageID) & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Roles") Then
      Send("    <td><a href=""role-list.aspx"" onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.role", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Terminals") Then
      Send("    <td><a href=""terminal-edit.aspx?TerminalID=" & dst.Rows(i).Item("LinkID") & """ onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.terminal", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Tenders") Then
      If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) AndAlso MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) Then
        Send("    <td>" & Copient.PhraseLib.Lookup("term.tendertypes", LanguageID) & "</td>")
      ElseIf MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) Then
        Send("    <td><a href=""tender-engines.aspx"" onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.tendertypes", LanguageID) & "</a></td>")
      ElseIf MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM) Then
        Send("    <td><a href=""tender.aspx"" onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.tendertypes", LanguageID) & "</a></td>")
      Else
        Send("    <td>" & Copient.PhraseLib.Lookup("term.tendertypes", LanguageID) & "</td>")
      End If
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Admin Users") Then
      Send("    <td><a href=""user-edit.aspx?UserID=" & dst.Rows(i).Item("LinkID") & """ onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.user", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Settings") OrElse (dst.Rows(i).Item("ActivityTypeName") = "settings") Then
      Send("    <td><a href=""settings.aspx"" onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.settings", LanguageID) & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Customer Inquiry") Then
      Send("    <td><a href=""customer-inquiry.aspx?CustPK=" & dst.Rows(i).Item("LinkID") & """ onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.customer", LanguageID) & " " & GetCustomerExtID(MyCommon, dst.Rows(i).Item("LinkID")) & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Stored Value") Then
      Send("    <td><a href=""sv-edit.aspx?ProgramGroupID=" & dst.Rows(i).Item("LinkID") & """ onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.storedvalueprogram", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Promotion Variables") Then
      Send("    <td><a href=""promovar-edit.aspx?PromoVarID=" & dst.Rows(i).Item("LinkID") & """ onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.promovar", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Banner") Then
      Send("    <td><a href=""banner-edit.aspx?BannerID=" & dst.Rows(i).Item("LinkID") & """ onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.banner", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Reports") Then
      Send("    <td><a href=""reports-detail.aspx?OfferID=" & dst.Rows(i).Item("LinkID") & """ onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.report", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Agents") Then
      Send("    <td><a href=""system-detail.aspx?appid=" & dst.Rows(i).Item("LinkID") & """ onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.agent", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "CM Settings") OrElse (dst.Rows(i).Item("ActivityTypeName") = "CM settings") Then
      Send("    <td><a href=""CM-settings.aspx"" onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.cmsettings", LanguageID) & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "CPE Settings") OrElse (dst.Rows(i).Item("ActivityTypeName") = "CPE settings") Then
      Send("    <td><a href=""CPEsettings.aspx"" onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.cpesettings", LanguageID) & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Web Settings") OrElse (dst.Rows(i).Item("ActivityTypeName") = "Web settings") Then
      Send("    <td><a href=""websettings.aspx"" onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.websettings", LanguageID) & "</a></td>")
      'ElseIf (dst.Rows(i).Item("ActivityTypeName") = "DP Settings") OrElse (dst.Rows(i).Item("ActivityTypeName") = "DP settings") Then
      '  Send("    <td><a href=""DP-settings.aspx"" onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.dpsettings", LanguageID) & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "External Sources") Then
      Send("    <td><a href=""sources-edit.aspx?SourceID=" & dst.Rows(i).Item("LinkID") & """ onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.source", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Issuance") Then
      Send("    <td><a href=""issuance.aspx"" onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.issuance", LanguageID) & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Vendor") Then
      Send("    <td><a href=""vendor-edit.aspx?VendorID=" & dst.Rows(i).Item("LinkID") & """ onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.vendor", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Campaign") Then
      Send("    <td>" & Copient.PhraseLib.Lookup("term.campaign", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Event") Then
      Send("    <td>" & Copient.PhraseLib.Lookup("term.event", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Scorecard") Then
      Send("    <td><a href=""scorecard-edit.aspx?ScorecardID=" & dst.Rows(i).Item("LinkID") & """ onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.scorecard", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "TerminalLockingGroup") Then
      Send("    <td><a href=""terminal-lockgroup-edit.aspx?TerminalLockingGroupID=" & dst.Rows(i).Item("LinkID") & """ onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.terminallockgroup", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Connectors") Then
      Send("    <td><a href=""connector-detail.aspx?ConnectorID=" & dst.Rows(i).Item("LinkID") & """ onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.connector", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Attributes") Then
      Send("    <td><a href=""attribute-edit.aspx?AttributeID=" & dst.Rows(i).Item("LinkID") & """ onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.attribute", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Advanced Limits") Then
      Send("    <td><a href=""CM-advlimit-edit.aspx?LimitID=" & dst.Rows(i).Item("LinkID") & """ onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.advlimit", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Folders") Then
      Send("    <td><a href=""folders.aspx?FolderID=" & dst.Rows(i).Item("LinkID") & """ onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.folder", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Health Settings") OrElse (dst.Rows(i).Item("ActivityTypeName") = "Health settings") Then
      Send("    <td><a href=""health-settings.aspx"" onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.healthsettings", LanguageID) & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Customer supplemental field") Then
      Send("    <td><a href=""customer-supplemental-edit.aspx?FieldID=" & dst.Rows(i).Item("LinkID") & """ onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.customersupplementalfield", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</a></td>")
    ElseIf (dst.Rows(i).Item("ActivityTypeName") = "Mutual exclusion groups") Then
      Send("    <td><a href=""MEG-edit.aspx?MutualExclusionGroupID=" & dst.Rows(i).Item("LinkID") & """ onClick=""return targetOpener(this)"">" & Copient.PhraseLib.Lookup("term.MutualExclusionGroup", LanguageID) & " " & dst.Rows(i).Item("LinkID") & "</a></td>")
    End If
    Send("  </tr>")
    If Shaded = "shaded" Then
      Shaded = ""
    Else
      Shaded = "shaded"
    End If
    i = i + 1
  End While
  If (sizeOfData = 0) Then
    Send("<tr>")
    Send("  <td colspan=""3""></td>")
    Send("</tr>")
  End If
  Send("</tbody>")
  Send("</table>")
  Send("")
  Send("</div>")
  Send("</div>")
%>

<script runat="server">
    Dim MyCryptLib As New Copient.CryptLib
  Function GetCustomerExtID(ByRef MyCommon As Copient.CommonInc, ByVal PKID As Integer) As String
    Dim dt As DataTable
    Dim ExtCardID As String = ""
    If (Not MyCommon.LXSadoConn.State = ConnectionState.Open) Then MyCommon.Open_LogixXS()
    MyCommon.QueryStr = "select top 1 ExtCardID from CardIDs with (NoLock) where CustomerPK=" & PKID & ";"
    dt = MyCommon.LXS_Select
    If (dt.Rows.Count > 0) Then
            ExtCardID = MyCryptLib.SQL_StringDecrypt(dt.Rows(0).Item("ExtCardID").ToString())
    End If
    If (Not MyCommon.LXSadoConn.State = ConnectionState.Closed) Then MyCommon.Close_LogixXS()
    Return ExtCardID
  End Function
</script>

<%
done:
  Send_BodyEnd()
  If MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER) Then
    MyCommon.Close_PrefManRT()
  End If
  MyCommon.Close_LogixRT()
  Logix = Nothing
  MyCommon = Nothing
%>
