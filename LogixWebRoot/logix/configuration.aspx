<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: configuration.aspx 
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
  
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim AdminUserID As Long
  Dim dt As System.Data.DataTable
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "configuration.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  Send_HeadBegin("term.configuration")
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
  Send_Subtabs(Logix, 8, 4)
  
  If (Logix.UserRoles.EditSystemConfiguration = False) Then
    Send_Denied(1, "perm.admin-configuration")
    GoTo done
  End If
%>
<script type="text/javascript">
  function launchCacheSettings() {
    var popW = 700;
    var popH = 570;
    var url = 'cachesettings.aspx';

    cacheWindow = window.open(url, "cacheTree", "width=" + popW + ", height=" + popH + ", top=" + calcTop(popH) + ", left=" + calcLeft(popW) + ", toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=yes");
    cacheWindow.focus();
  }
  function launchProdHierarchy() {
    var popW = 900;
    var popH = 570;
    var url = 'phierarchytree.aspx?ProductGroupID=-1';
    
    phierWindow = window.open(url,"pHierTree", "width="+popW+", height="+popH+", top="+calcTop(popH)+", left="+calcLeft(popW)+", toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=yes");
    phierWindow.focus();
  }
  function launchLocHierarchy() {
    var popW = 700;
    var popH = 570;
    var url = 'lhierarchytree.aspx?LocationGroupID=-1';
    
    lhierWindow = window.open(url,"lHierTree", "width="+popW+", height="+popH+", top="+calcTop(popH)+", left="+calcLeft(popW)+", toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=yes");
    lhierWindow.focus();
  }
</script>

<div id="intro">
  <h1 id="title">
    <% Sendb(Copient.PhraseLib.Lookup("term.configuration", LanguageID))%>
  </h1>
  <div id="controls">
  </div>
</div>
<div id="main">
  <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
  <p>
    <% Sendb(Copient.PhraseLib.Lookup("configuration.main", LanguageID))%>
  </p>
  
  <div style="float: left; padding-right: 10px;">
    <% If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE)) Then%>
    <form action="attribute-list.aspx" id="attributesform" name="attributesform">
      <input type="submit" id="attributes" name="attributes" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.attributes", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.attributes", LanguageID)) %>"/>
    </form>
    &nbsp;
    <% End If%>
    <!-- BZ2079: UE-feature-removal #31: Hiding "Categories" button -->
    <% If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) Or MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) Then%>
    <form action="category-list.aspx" id="categoriesform" name="categoriesform">
      <input type="submit" id="categories" name="categories" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.categories", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.categories", LanguageID)) %>" />
    </form>
    &nbsp;
    <% End If %>
    <%
      If (MyCommon.Fetch_SystemOption(110) = "1") Then
        MyCommon.QueryStr = "select FieldID from CustomerSupplementalFields with (NoLock);"
        dt = MyCommon.LXS_Select
        If dt.Rows.Count > 0 Then
          Send("    <form action=""customer-supplemental-list.aspx"" id=""customersupplementalform"" name=""customersupplementalform"">")
          Send("      <input type=""submit"" id=""customersupplemental"" name=""customersupplemental"" style=""width:175px;"" value=""" & Copient.PhraseLib.Lookup("term.customersupplementals", LanguageID) & """ & title=""" & Copient.PhraseLib.Lookup("term.customersupplementals", LanguageID) & """ />")
          Send("    </form>")
          Send("    &nbsp;")
        End If
      End If
    %>
    <form action="dataexports.aspx" id="dataexportsform" name="dataexportsform">
      <input type="submit" id="dataexport" name="dataexport" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.dataexports", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.dataexports", LanguageID)) %>"/>
    </form>
    &nbsp;
    <form action="department-list.aspx" id="departmentsform" name="departmentsform">
      <input type="submit" id="departments" name="departments" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.departments", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.departments", LanguageID)) %>"/>
    </form>
    &nbsp;
    <% If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE)) Then%>
    <form action="buyer-list.aspx" id="buyersimportform" name="buyerimportform">
    <input type="submit" id="roles2" name="roles" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.buyers", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.buyers", LanguageID)) %>" />
    </form>
    &nbsp;
    <% End If %>
    <form action="sources-list.aspx" id="sourcesform" name="sourcesform">
      <input type="submit" id="sources" name="roles" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.externalsources", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.externalsources", LanguageID)) %>" />
    </form>
    &nbsp;
    <!-- BZ2079: UE-feature-removal #33: Hiding "Manual Adjustment UPCs" button -->
    <% If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) Or MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CAM)) Then%>
    <form action="adjustmentUPC-list.aspx" id="adjustmentupcform" name="adjustmentupcform">
      <input type="submit" id="adjustmentupc" name="adjustmentupc" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.adjustmentupcs", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.adjustmentupcs", LanguageID)) %>"/>
    </form>
    &nbsp;
    <% End If %>
    <% If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE) AndAlso MyCommon.Fetch_UE_SystemOption(157) = "1") Then%>
    <form action="Attribute-PGBuilderConfig.aspx" id="pgattributeconfigurationform" name="pgattributeconfigurationform">
      <input type="submit" id="pgattributeconfig" name="pgattributeconfig" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.attributepg-config", LanguageID)) %>"  title="<% Sendb(Copient.PhraseLib.Lookup("term.attributepg-config", LanguageID)) %>" />
    </form>
    &nbsp;
    <% End If %>
    <% If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE) AndAlso MyCommon.Fetch_UE_SystemOption(135) = "1") Then%>
    <form action="units-of-measure.aspx" id="Form2" name="uomform">
      <input type="submit" id="uom" name="uom" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.unitsofmeasure", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.unitsofmeasure", LanguageID)) %>" />
    </form>
    &nbsp;
    <% End If %>
      <% If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE)) Then%>
    <form action="UE/OfferApproval.aspx" id="offerapproval" name="offerapproval">
      <input type="submit" id="oaw" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.offerapprovalworkflow", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.offerapprovalworkflow", LanguageID)) %>" />
    </form>
    &nbsp;
    <% End If %>
	<% If (MyCommon.Fetch_SystemOption(156) = "1" AndAlso Logix.UserRoles.AccessUserDefinedFields = True)  Then%>
    <form action="UserDefinedFields-list.aspx" id="UserDefinedFields" name="UserDefinedFields">
      <input type="submit" id="UserDefinedFields" name="UserDefinedFields" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.userdefinedfields", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.userdefinedfields", LanguageID)) %>"/>
    </form>
    &nbsp;
    <% End If %>
  </div>
  
  <div style="float: left; padding-right: 10px;">
    <% If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE)) Then%>
    <form action="MEG-list.aspx" id="MEGform" name="MEGform">
      <input type="submit" id="MEG" name="MEG" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.MutualExclusionGroups", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.MutualExclusionGroups", LanguageID)) %>" />
    </form>
    &nbsp;
    <% End If %>
    <form action="" id="phierarchyform" name="phierarchyform">
      <input type="button" id="phierarchy" name="phierarchy" onclick="javascript:launchProdHierarchy();" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.producthierarchies", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.producthierarchies", LanguageID)) %>" />
    </form>
    &nbsp;
    <form action="role-list.aspx" id="rolesform" name="rolesform">
      <input type="submit" id="roles" name="roles" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.roles", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.roles", LanguageID)) %>"  />
    </form>
    &nbsp;
    <% If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE)) Or (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CAM)) Or (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE)) Then%>
    <form action="scorecard-list.aspx" id="scorecardformd" name="scorecardformd">
      <input type="hidden" id="ScorecardTypeIDd" name="ScorecardTypeID" value="3" />
      <input type="submit" id="scorecardd" name="scorecardd" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.scorecards", LanguageID) & ": " & Copient.PhraseLib.Lookup("term.discount", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.scorecards", LanguageID) & ": " & Copient.PhraseLib.Lookup("term.discount", LanguageID)) %>"/>
    </form>
    &nbsp;
    <% End If%>

    <% If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE)) Or (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CAM)) Then%>
    <form action="scorecard-list.aspx" id="scorecardformpp" name="scorecardformpp">
      <input type="hidden" id="ScorecardTypeIDpp" name="ScorecardTypeID" value="1" />
      <input type="submit" id="scorecardpp" name="scorecardpp" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.scorecards", LanguageID) & ": " & Copient.PhraseLib.Lookup("term.points", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.scorecards", LanguageID) & ": " & Copient.PhraseLib.Lookup("term.points", LanguageID)) %>" />
    </form>
    &nbsp;
    <form action="scorecard-list.aspx" id="scorecardformsv" name="scorecardformsv">
      <input type="hidden" id="ScorecardTypeIDsv" name="ScorecardTypeID" value="2" />
      <input type="submit" id="scorecardsv" name="scorecardsv" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.scorecards", LanguageID) & ": " & Copient.PhraseLib.Lookup("term.sv", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.scorecards", LanguageID) & ": " & Copient.PhraseLib.Lookup("term.sv", LanguageID)) %>"/>
    </form>
    &nbsp;
    <form action="scorecard-list.aspx" id="scorecardforml" name="scorecardforml">
      <input type="hidden" id="ScorecardTypeIDl" name="ScorecardTypeID" value="4" />
      <input type="submit" id="scorecardl" name="scorecardl" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.scorecards", LanguageID) & ": " & Copient.PhraseLib.Lookup("term.limits", LanguageID)) %>" title= "<% Sendb(Copient.PhraseLib.Lookup("term.scorecards", LanguageID) & ": " & Copient.PhraseLib.Lookup("term.limits", LanguageID)) %>"/>
    </form>
    &nbsp;
    <% End If%>
    <% If (MyCommon.Fetch_UE_SystemOption(221) = "1") Then%>
    <form action="CardRangeConfig.aspx" id="CardRangeConfig" name="CardRangeConfig">
       <input type="submit" id="RangeConfig" name="CardRangeConfig" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.card", LanguageID) & ": " & Copient.PhraseLib.Lookup("term.range", LanguageID)) %>" title= "<% Sendb(Copient.PhraseLib.Lookup("term.card", LanguageID) & ": " & Copient.PhraseLib.Lookup("term.range", LanguageID))%>"/>
    </form>
    &nbsp;
    <% End If%>
    <% If MyCommon.Fetch_SystemOption("318") = "1" Then%>
    <form action="reports-enhanced-configuration.aspx" id="reportconfigurationform" name="reportconfigurationform">
      <input type="hidden" id="reportconfigurationID" name="reportconfigurationID" value="4" />
      <input type="submit" id="reportConfig" name="reportConfig" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.reportconfiguration", LanguageID)) %>" title= "<% Sendb(Copient.PhraseLib.Lookup("term.reportconfiguration", LanguageID)) %>"/>
    </form>
    &nbsp;
   <% End If%>
  </div>
  
  <div style="float: left; padding-right: 10px;">
    <form action="" id="lhierarchyform" name="lhierarchyform">
      <input type="button" id="lhierarchy" name="lhierarchy" onclick="javascript:launchLocHierarchy();" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.storehierarchies", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.storehierarchies", LanguageID)) %>" />
    </form>
    &nbsp;
    <% If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) OrElse MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE)) AndAlso (MyCommon.Fetch_SystemOption(198) <> MyCommon.Fetch_SystemOption(199)) Then%>
    <form action="PLU-list.aspx" id="PLUsform" name="PLUsform">
      <input type="submit" id="PLUs" name="PLUs" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.triggercodes", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.triggercodes", LanguageID)) %>"/>
    </form>
    &nbsp;
    <% End If %>
    <% If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) Then%>
    <form action="tender.aspx" id="tendersform" name="tendersform">
      <input type="submit" id="tenders" name="tenders" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.tendertypes", LanguageID) & ": " & Copient.PhraseLib.Lookup("term.cm", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.tendertypes", LanguageID) & ": " & Copient.PhraseLib.Lookup("term.cm", LanguageID)) %>" />
    </form>
    &nbsp;
    <% End If%>
    <% If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) OrElse MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE)) Then%>
    <form action="tender-engines.aspx" id="CPEtendersform" name="CPEtendersform">
      <input type="submit" id="CPEtenders" name="CPEtenders" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.tendertypes", LanguageID) & ": " & Copient.PhraseLib.Lookup("term.cpe", LanguageID) & "/" & Copient.PhraseLib.Lookup("term.ue", LanguageID)) %>"  title="<% Sendb(Copient.PhraseLib.Lookup("term.tendertypes", LanguageID) & ": " & Copient.PhraseLib.Lookup("term.cpe", LanguageID) & "/" & Copient.PhraseLib.Lookup("term.ue", LanguageID)) %>" />
    </form>
    &nbsp;
    <% End If%>
    <form action="terminal-list.aspx" id="terminalsform" name="terminalsform">
      <input type="submit" id="terminals" name="terminals" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.terminals", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.terminals", LanguageID)) %>"/>
    </form>
    &nbsp;
    <% If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE)) AndAlso (MyCommon.Fetch_CPE_SystemOption(86) = "2") Then%>
    <form action="terminal-lockgroup-list.aspx" id="terminallockgroupsform" name="terminallockgroupsform">
      <input type="submit" id="terminallockgroups" name="terminallockgroups" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.terminallockgroups", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.terminallockgroups", LanguageID)) %>" />
    </form>
    &nbsp;
    <% End If%>
    <form action="vendor-list.aspx" id="vendorform" name="vendorform">
      <input type="submit" id="vendor" name="vendor" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.vendors", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.vendors", LanguageID)) %>"/>
    </form>
    &nbsp;
    <% If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CAM) Or MyCommon.IsEngineSubTypeInstalled(Copient.CommonInc.InstalledEngineSubTypes.CPEUSAirMiles)) Then%>
    <form action="rejectiontypes-edit.aspx" id="rejectiontypes" name="rejectiontypes">
      <input type="submit" id="rejectiontypes" name="rejectiontypes" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.rejectiontypes", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.rejectiontypes", LanguageID)) %>" />
    </form>
    &nbsp;
    <% End If%>
		<%If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE)) Then		%>
    <form action="triggercodemessages-list.aspx" id="Form1" name="triggercodemessages">
      <input type="submit" id="predefinedtriggercode" name="predefinedtriggercode" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.triggercodemessage", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.triggercodemessage", LanguageID)) %>" />
    </form>
    &nbsp;
    <%End If%>
	<% 
  If MyCommon.Fetch_SystemOption(124) = "1" And MyCommon.IsIntegrationInstalled(Copient.CommonIncConfigurable.Integrations.PREFERENCE_MANAGER) = True Then
    'only display the Customer Languages button if multi-language is enabled and EPM is installed
    Send("<form action=""customer-languages.aspx"" id=""custLangForm"" name=""custLangForm"">")
    Send("    <input type=""submit"" id=""customerLanguages"" name=""customerLanguages"" style=""width:175px;"" value=""" & Copient.PhraseLib.Lookup("term.customerlanguages", LanguageID) & """ & title=""" & Copient.PhraseLib.Lookup("term.customerlanguages", LanguageID) & """ />")
    Send("</form>")
  End If
  %>
  &nbsp;
  </div>
  <div style="float: left; padding-right: 10px;">
    <form action="settings.aspx" id="settingsform" name="settingsform">
      <input type="submit" id="settings" name="settings" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.settings", LanguageID)) %>"  title="<% Sendb(Copient.PhraseLib.Lookup("term.settings", LanguageID)) %>" />
    </form>
    &nbsp;
    <% If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) Then%>
    <form action="CM-settings.aspx" id="CMsettingsform" name="CMsettingsform">
      <input type="submit" id="cmsettings" name="cmsettings" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.settings", LanguageID)) %>: <% Sendb(Copient.PhraseLib.Lookup("term.cm", LanguageID)) %>"  title="<% Sendb(Copient.PhraseLib.Lookup("term.settings", LanguageID)) %>: <% Sendb(Copient.PhraseLib.Lookup("term.cm", LanguageID)) %>"/>
    </form>
    &nbsp;
    <% End If%>
    <% If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) AndAlso (MyCommon.Fetch_CM_SystemOption(35) = "1") Then%>
    <form action="CM-cashier-inquiry-options.aspx" id="CMcashierform" name="CMcashierform">
      <input type="submit" id="cmcashier" name="cmcashier" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.settings", LanguageID)) %>: <% Sendb(Copient.PhraseLib.Lookup("term.cm", LanguageID) & " " & Copient.PhraseLib.Lookup("term.cashier", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.settings", LanguageID)) %>: <% Sendb(Copient.PhraseLib.Lookup("term.cm", LanguageID) & " " & Copient.PhraseLib.Lookup("term.cashier", LanguageID)) %>" />
    </form>
    &nbsp;
    <% End If%>
    <% If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE)) Then%>
    <form action="CPEsettings.aspx" id="CPEsettingsform" name="CPEsettingsform">
      <input type="submit" id="cpesettings" name="cpesettings" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.settings", LanguageID)) %>: <% Sendb(Copient.PhraseLib.Lookup("term.cpe", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.settings", LanguageID)) %>: <% Sendb(Copient.PhraseLib.Lookup("term.cpe", LanguageID)) %>" />
    </form>
    &nbsp;
    <% End If%>

      <!-- BZ2079: UE-feature-removal #41: Hiding "Settings: Health" button -->
    <% If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) Or MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CAM)) Then%>
    <form action="health-settings.aspx" id="healthsettingsform" name="healthsettingsform">
      <input type="submit" id="healthsettings" name="healthsettings" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.settings", LanguageID)) %>: <% Sendb(Copient.PhraseLib.Lookup("term.cpe", LanguageID) +" " + Copient.PhraseLib.Lookup("term.health", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.settings", LanguageID)) %>: <% Sendb(Copient.PhraseLib.Lookup("term.cpe", LanguageID) +" " + Copient.PhraseLib.Lookup("term.health", LanguageID)) %>" />
    </form>
    &nbsp;
    <% End If%>

    <% If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE)) Then%>
    <form action="UE/UEsettings.aspx" id="Form1" name="UEsettingsform">
      <input type="submit" id="Submit1" name="uesettings" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.settings", LanguageID)) %>: <% Sendb(Copient.PhraseLib.Lookup("term.ue", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.settings", LanguageID)) %>: <% Sendb(Copient.PhraseLib.Lookup("term.ue", LanguageID)) %>" />
    </form>
    &nbsp;
    <%--
    <form action="CPEsettings_schedule.aspx" id="CPEsettingsform_schedule" name="CPEsettingsform_schedule">
      <input type="submit" id="cpesettings_schedule" name="cpesettings_schedule" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.settings", LanguageID)) %>: <% Sendb(Copient.PhraseLib.Lookup("term.settings-schedule", LanguageID)) %>" />
    </form>
    &nbsp;
    --%>
    <% End If%>
    <% If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) AndAlso (MyCommon.Fetch_CM_SystemOption(60) = "1") Then%>
    <form action="CM-extracts.aspx" id="CmExtractsform" name="CmExtractsform">
       <input type="submit" id="cmextracts" name="cmextracts" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.settings", LanguageID)) %>: <% Sendb(Copient.PhraseLib.Lookup("term.cmextracts", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.settings", LanguageID)) %>: <% Sendb(Copient.PhraseLib.Lookup("term.cmextracts", LanguageID)) %>" />
    </form>
    &nbsp;
    <% End If%>
     <% If (MyCommon.Fetch_SystemOption(190) = "1") Then%>
    <form action="Agent-Schedulingoptions.aspx" id="schedulingoptionsform" name="schedulingoptionsform">
       <input type="submit" id="agentschedulingoptions" name="agentschedulingoptions" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.settings", LanguageID)) %>: <% Sendb(Copient.PhraseLib.Lookup("term.schedulingoptions", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.settings", LanguageID)) %>: <% Sendb(Copient.PhraseLib.Lookup("term.schedulingoptions", LanguageID)) %>"/>
    </form>
    &nbsp;
    <% End If%>
  
    <% If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.Website)) Then%>
    <form action="websettings.aspx" id="websettingsform" name="websettingsform">
      <input type="submit" id="websettings" name="websettings" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.settings", LanguageID)) %>: <% Sendb(Copient.PhraseLib.Lookup("term.web", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.settings", LanguageID)) %>: <% Sendb(Copient.PhraseLib.Lookup("term.web", LanguageID)) %>" />
    </form>
    &nbsp;
    <% End If%>
    <form action="" id="cachesettingsform" name="cachesettingsform">
      <input type="button" id="cachesettings" name="cachesettings" onclick="javascript:launchCacheSettings();" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.settings", LanguageID)) %>: <% Sendb(Copient.PhraseLib.Lookup("term.cache", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.settings", LanguageID)) %>: <% Sendb(Copient.PhraseLib.Lookup("term.cache", LanguageID)) %>"/>
    </form>
    &nbsp;
    <% If (Logix.IsPermitted(193, AdminUserID, MyCommon) AndAlso MyCommon.use_development_feature("BZ2222")) Then%>
    <form action="configurator/management.aspx" id="settingsconfigform" name="settingsconfigform">
      <input type="submit" id="settingsconfig" name="settingsconfig" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.settingsmanagement", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.settingsmanagement", LanguageID)) %>" />
    </form>
    <% End If%>
    <!-- AMSPS435 Add Reason Code button -->
    <% If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) Or MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE) Or MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) Then%>
    <form action="reasons-list.aspx" id="reasoncodes" name="reasoncodes">
      <input type="submit" id="reasons" name="reasons" style="width:175px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.reasoncodes", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.reasoncodes", LanguageID)) %>" />
    </form>
    <% End If%>
  </div>
</div>
<%
done:
  Send_BodyEnd()
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  MyCommon = Nothing
  Logix = Nothing
%>
