﻿<%@ Page Language="vb" Debug="true" CodeFile="ue-cb.vb" Inherits="UECB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
    ' *****************************************************************************
    ' * FILENAME: UEoffer-con-instantwin.aspx 
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
    Dim OfferID As Long
    Dim Name As String
    Dim ConditionID As String
    Dim IsTemplate As Boolean = False
    Dim Disallow_Edit As Boolean = False
    Dim DisabledAttribute As String = ""
    Dim FromTemplate As Boolean = False
    Dim Unlimited As Boolean = False
    Dim i As Integer
    Dim row As DataRow
    Dim rst As DataTable
    Dim CloseAfterSave As Boolean = False
    Dim historyString As String = ""
    Dim roid As Integer
    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False
    Dim EngineID As Integer = 2
    Dim BannersEnabled As Boolean = True
    Dim TempInt As Integer
    Dim CalcMethod As Integer = 0

    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If

    Response.Expires = 0
    MyCommon.AppName = "UEoffer-con-instantwin.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)

    BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
    OfferID = Request.QueryString("OfferID")

    'This would redirect to summary page if offer is awaiting deployment or if offer is awaiting recommendations.
    CheckIfValidOffer(MyCommon, OfferID)

    ConditionID = Request.QueryString("ConditionID")
    If (Request.QueryString("EngineID") <> "") Then
        EngineID = MyCommon.Extract_Val(Request.QueryString("EngineID"))
    Else
        MyCommon.QueryStr = "select EngineID from OfferIDs where OfferID=" & OfferID & ";"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
            EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), 0)
        End If
    End If

    MyCommon.QueryStr = "select RewardOptionID from CPE_RewardOptions with (NoLock) where IncentiveID=" & OfferID & " and TouchResponse=0 and Deleted=0;"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
        roid = rst.Rows(0).Item("RewardOptionID")
    End If

    ' User clicked save
    If (Request.QueryString("save") <> "") Then
        ' validate entries
        Unlimited = IIf(Request.QueryString("Unlimited") = "on", True, False)
        If (Not Integer.TryParse(Request.QueryString("NumPrizesAllowed"), TempInt) OrElse TempInt <= 0) AndAlso Not Unlimited Then
            infoMessage = Copient.PhraseLib.Lookup("cpe-offer-con-instantwin.invalidlimit", LanguageID)
        ElseIf Not Integer.TryParse(Request.QueryString("OddsOfWinning"), TempInt) OrElse TempInt <= 0 Then
            infoMessage = Copient.PhraseLib.Lookup("cpe-offer-con-instantwin.oddsofwinning", LanguageID)
        ElseIf Not Integer.TryParse(Request.QueryString("RandomWinners"), TempInt) OrElse TempInt < 0 OrElse TempInt > 1 Then
            infoMessage = Copient.PhraseLib.Lookup("cpe-offer-con-instantwin.randomwinner", LanguageID)
        Else
            ' Store the existing locking value for use in newly-created records
            MyCommon.QueryStr = "select DisallowEdit from CPE_IncentiveInstantWin with (NoLock) where Deleted=0 and RewardOptionID=" & OfferID & ";"
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count > 0) Then
                Disallow_Edit = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit"), False)
            Else
                Disallow_Edit = False
            End If
            ' First, delete all instant win records for this offer, then insert new
            MyCommon.QueryStr = "update CPE_IncentiveInstantWin with (RowLock) set Deleted=1 where RewardOptionID=" & roid & ";"
            MyCommon.LRT_Execute()
            MyCommon.QueryStr = "insert into CPE_IncentiveInstantWin with (RowLock) (RewardOptionID, NumPrizesAllowed, OddsOfWinning, RandomWinners, Deleted, LastUpdate, DisallowEdit, RequiredFromTemplate, Unlimited) " & _
                                "values (" & roid & ", " & IIf(Request.QueryString("NumPrizesAllowed") <> "", Request.QueryString("NumPrizesAllowed"), "0") & ", " & Request.QueryString("OddsOfWinning") & ", " & Request.QueryString("RandomWinners") & ", 0, getdate(), " & IIf(Disallow_Edit, "1", "0") & ", 0, " & IIf(Unlimited, "1", "0") & ");"
            MyCommon.LRT_Execute()
            MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1, EnableImpressRpt=1, EnableRedeemRpt=1 where IncentiveID=" & OfferID & ";"
            MyCommon.LRT_Execute()
            ResetOfferApprovalStatus(OfferID)
            MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-instantwin", LanguageID))
            CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1") AndAlso infoMessage = ""
        End If
    End If

    ' Load offer data
    MyCommon.QueryStr = "Select IncentiveID,IsTemplate,ClientOfferID,IncentiveName as Name,CPE.Description,PromoClassID,CRMEngineID,Priority," & _
                        "StartDate,EndDate,EveryDOW,EligibilityStartDate,EligibilityEndDate,TestingStartDate,TestingEndDate,P1DistQtyLimit,P1DistTimeType,P1DistPeriod," & _
                        "P2DistQtyLimit,P2DistTimeType,P2DistPeriod,P3DistQtyLimit,P3DistTimeType,P3DistPeriod,EnableImpressRpt,EnableRedeemRpt,CreatedDate,CPE.LastUpdate, CPE.Deleted, CPEOARptDate, CPEOADeploySuccessDate, CPEOADeployRpt," & _
                        "CRMRestricted,StatusFlag,OC.Description as CategoryName,IsTemplate,FromTemplate from CPE_Incentives as CPE with (NoLock) " & _
                        "left join OfferCategories as OC with (NoLock) on CPE.PromoClassID=OfferCategoryID where IncentiveID=" & Request.QueryString("OfferID")
    rst = MyCommon.LRT_Select
    For Each row In rst.Rows
        Name = MyCommon.NZ(row.Item("Name"), "")
        IsTemplate = MyCommon.NZ(row.Item("IsTemplate"), False)
        FromTemplate = MyCommon.NZ(row.Item("FromTemplate"), False)
    Next

    ' Update templates permission if necessary
    If (Request.QueryString("save") <> "" AndAlso Request.QueryString("IsTemplate") = "IsTemplate") Then
        ' update the status bits for the templates
        Dim form_Disallow_Edit As Integer = 0
        If (Request.QueryString("Disallow_Edit") = "on") Then
            form_Disallow_Edit = 1
        End If
        MyCommon.QueryStr = "update CPE_IncentiveInstantWin with (RowLock) set DisallowEdit=" & form_Disallow_Edit & _
                            " where RewardOptionID=" & roid & " and Deleted=0;"
        MyCommon.LRT_Execute()
    End If

    If (IsTemplate Or FromTemplate) Then
        ' Load permissions if it's a template
        MyCommon.QueryStr = "select DisallowEdit from CPE_IncentiveInstantWin with (NoLock) where RewardOptionID=" & roid & " and Deleted=0;"
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
            Disallow_Edit = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit"), False)
        Else
            Disallow_Edit = False
        End If
    End If
    Dim m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
    If Not IsTemplate Then
        DisabledAttribute = IIf(Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And Disallow_Edit), "", "disabled=""disabled""")
    Else
        DisabledAttribute = IIf(Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer, "", "disabled=""disabled""")
    End If

    Send_HeadBegin("term.offer", "term.instantwincondition", OfferID)
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts()
    Send("<script type=""text/javascript"">")
    Send("function ChangeParentDocument() { ")
    If (EngineID = 3) Then
        Send("  if (opener != null) {")
        Send("    var newlocation = '/logix/web-offer-con.aspx?OfferID=" & OfferID & "'; ")
        Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
        Send("  opener.location = '/logix/web-offer-con.aspx?OfferID=" & OfferID & "'; ")
    ElseIf (EngineID = 5) Then
        Send("  if (opener != null) {")
        Send("    var newlocation = '/logix/email-offer-con.aspx?OfferID=" & OfferID & "'; ")
        Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
        Send("  opener.location = '/logix/email-offer-con.aspx?OfferID=" & OfferID & "'; ")
    ElseIf (EngineID = 6) Then
        Send("  if (opener != null) {")
        Send("    var newlocation = '/logix/CAM/CAM-offer-con.aspx?OfferID=" & OfferID & "'; ")
        Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
        Send("  opener.location = '/logix/CAM/CAM-offer-con.aspx?OfferID=" & OfferID & "'; ")
    ElseIf (EngineID = 9) Then
        Send("  if (opener != null) {")
        Send("    var newlocation = '/logix/UE/UEoffer-con.aspx?OfferID=" & OfferID & "'; ")
        Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
        Send("  opener.location = '/logix/UE/UEoffer-con.aspx?OfferID=" & OfferID & "'; ")
    Else
        Send("  if (opener != null) {")
        Send("    var newlocation = '/logix/CPEoffer-con.aspx?OfferID=" & OfferID & "'; ")
        Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
        Send("  opener.location = '/logix/CPEoffer-con.aspx?OfferID=" & OfferID & "'; ")
    End If
    Send("} ")
    Send("} ")
    Send("} ")
%>
  function toggleLimit() {
    var unlimitedCheckbox = document.getElementById("Unlimited");
    var limitsInput = document.getElementById("NumPrizesAllowed");

    if ((unlimitedCheckbox != null) && (limitsInput != null)) {
      if (unlimitedCheckbox.checked == true) {
        limitsInput.style.display = 'none';
        limitsInput.value = '';
      } else {
        limitsInput.style.display = 'inline';
        limitsInput.focus();
      }
    }
  }

  // Fix for IE8 where onchange works after the element loses focus (AL-1561)
  function radioClick() {
    this.blur();  
    this.focus();  
  }

<%
  Send("</script>")
  Send_HeadEnd()
  
  If (IsTemplate) Then
    Send_BodyBegin(12)
  Else
    Send_BodyBegin(2)
  End If
  If (Logix.UserRoles.AccessOffers = False AndAlso Not IsTemplate) Then
    Send_Denied(2, "perm.offers-access")
    GoTo done
  End If
  If (Logix.UserRoles.AccessTemplates = False AndAlso IsTemplate) Then
    Send_Denied(2, "perm.offers-access-templates")
    GoTo done
  End If
  If (BannersEnabled AndAlso Not Logix.IsAccessibleOffer(AdminUserID, OfferID)) Then
    Send("<script type=""text/javascript"" language=""javascript"">")
    Send("  function ChangeParentDocument() { return true; } ")
    Send("</script>")
    Send_Denied(1, "banners.access-denied-offer")
    Send_BodyEnd()
    GoTo done
  End If
%>
<form action="#" id="mainform" name="mainform">
  <div id="intro">
    <input type="hidden" id="OfferID" name="OfferID" value="<% sendb(OfferID) %>" />
    <input type="hidden" id="Name" name="Name" value="<% sendb(Name) %>" />
    <input type="hidden" id="ConditionID" name="ConditionID" value="<% sendb(ConditionID) %>" />
    <input type="hidden" id="IsTemplate" name="IsTemplate" value="<% sendb(IIf(IsTemplate, "IsTemplate", "Not")) %>" />
    <%
      If (IsTemplate) Then
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.instantwincondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
      Else
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.instantwincondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
      End If
    %>
    <div id="controls">
      <% If (IsTemplate) Then%>
      <span class="temp">
        <input type="checkbox" class="tempcheck" id="Disallow_Edit" name="Disallow_Edit"<% if(disallow_edit)then sendb(" checked=""checked""") %> />
        <label for="Disallow_Edit"><% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
      </span>
      <% End If%>     
      <% 
          m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
      If Not IsTemplate Then
              If (Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And Disallow_Edit) And Not IsOfferWaitingForApproval(OfferID)) Then Send_Save()
      Else
              If (Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer) Then Send_Save()
      End If
      %>
    </div>
  </div>
  <div id="main">
    <%
      If (infoMessage <> "") Then
        Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
      End If
      
      MyCommon.QueryStr = "select NumPrizesAllowed, OddsOfWinning, RandomWinners, DisallowEdit, Unlimited " & _
                          "from CPE_IncentiveInstantWin as IIW with (NoLock) " & _
                          "where IIW.Deleted=0 and RewardOptionID=" & roid & ";"
      rst = MyCommon.LRT_Select
      If infoMessage <> "" Then
        Unlimited = IIf(Request.QueryString("Unlimited") = "on", True, False)
      Else
        If rst.Rows.Count > 0 Then
          Unlimited = MyCommon.NZ(rst.Rows(0).Item("Unlimited"), False)
        Else
          Unlimited = IIf(Request.QueryString("Unlimited") = "on", True, False)
        End If
      End If
    %>
    <div id="column1">
      <div class="box" id="limit">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.limit", LanguageID))%>
          </span>
        </h2>
        <%
          Send("<table summary=""" & Copient.PhraseLib.Lookup("term.limits", LanguageID) & """>")
          Send("  <tr>")
          Send("    <td style=""width:130px;"">")
          Send("      <label for=""NumPrizesAllowed"">" & Copient.PhraseLib.Lookup("term.awardlimit", LanguageID) & ":</label>")
          Send("    </td>")
          Send("    <td>")
          Send("      <input type=""checkbox"" id=""Unlimited"" name=""Unlimited"" onclick=""javascript:radioClick();"" onchange=""javascript:toggleLimit();""" & IIf(Unlimited, " checked=""checked""", "") & " /><label for=""Unlimited"">" & Copient.PhraseLib.Lookup("term.unlimited", LanguageID) & "</label><br />")
          If rst.Rows.Count > 0 Then
            Send("      <input class=""short"" id=""NumPrizesAllowed"" name=""NumPrizesAllowed"" maxlength=""9"" type=""text"" value=""" & IIf(infoMessage = "", MyCommon.NZ(rst.Rows(0).Item("NumPrizesAllowed"), ""), Request.QueryString("NumPrizesAllowed")) & """" & DisabledAttribute & IIf(Unlimited, " style=""display:none;""", "") & " />")
          Else
            Send("      <input class=""short"" id=""NumPrizesAllowed"" name=""NumPrizesAllowed"" maxlength=""9"" type=""text"" value=""" & Request.QueryString("NumPrizesAllowed") & """ " & DisabledAttribute & IIf(Unlimited, " style=""display:none;""", "") & " /><br />")
          End If
          Send("      <br class=""half"" />")
          Send("    </td>")
          Send("  </tr>")
          Send("</table>")
        %>
      </div>
      <div class="box" id="odds">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.odds", LanguageID))%>
          </span>
        </h2>
        <%
          Send("<table summary=""" & Copient.PhraseLib.Lookup("term.odds", LanguageID) & """>")
          Send("  <tr>")
          Send("    <td style=""width:130px;"">")
          Send("      <label for=""OddsOfWinning"">" & Copient.PhraseLib.Lookup("offer-gen.oddsofwinning", LanguageID) & "</label>")
          Send("    </td>")
          Send("    <td>")
          If rst.Rows.Count > 0 Then
            Send("      1:<input class=""short"" id=""OddsOfWinning"" name=""OddsOfWinning"" maxlength=""9"" type=""text"" value=""" & IIf(infoMessage = "", MyCommon.NZ(rst.Rows(0).Item("OddsOfWinning"), ""), Request.QueryString("OddsOfWinning")) & """" & DisabledAttribute & " />")
          Else
            Send("      1:<input class=""short"" id=""OddsOfWinning"" name=""OddsOfWinning"" maxlength=""9"" type=""text"" value=""" & Request.QueryString("OddsOfWinning") & """ " & DisabledAttribute & " />")
          End If
          Send("    </td>")
          Send("  </tr>")
          Send("  <tr>")
          Send("    <td>")
          Send("      " & Copient.PhraseLib.Lookup("term.calculation", LanguageID) & ":")
          Send("    </td>")
          Send("    <td>")
          If rst.Rows.Count > 0 Then
            If infoMessage <> "" AndAlso Request.QueryString("RandomWinners") <> "" Then
              CalcMethod = MyCommon.Extract_Val(Request.QueryString("RandomWinners"))
            Else
              If MyCommon.NZ(rst.Rows(0).Item("RandomWinners"), False) Then
                CalcMethod = 1
              Else
                CalcMethod = 0
              End If
            End If
            Send("      <input id=""fixed"" name=""RandomWinners"" type=""radio"" value=""0""" & IIf(CalcMethod = 0, " checked=""checked""", "") & DisabledAttribute & " /><label for=""fixed"">" & Copient.PhraseLib.Lookup("term.fixed", LanguageID) & "</label>")
            Send("      <br />")
            Send("      <input id=""random"" name=""RandomWinners"" type=""radio"" value=""1""" & IIf(CalcMethod = 1, " checked=""checked""", "") & DisabledAttribute & " /><label for=""random"">" & Copient.PhraseLib.Lookup("term.random", LanguageID) & "</label>")
          Else
            If infoMessage <> "" AndAlso Request.QueryString("RandomWinners") <> "" Then
              CalcMethod = MyCommon.Extract_Val(Request.QueryString("RandomWinners"))
            Else
              CalcMethod = -1
            End If
            Send("      <input id=""fixed"" name=""RandomWinners"" type=""radio"" value=""0""" & DisabledAttribute & IIf(CalcMethod = 0, " checked=""checked""", "") & " /><label for=""fixed"">" & Copient.PhraseLib.Lookup("term.fixed", LanguageID) & "</label>")
            Send("      <br />")
            Send("      <input id=""random"" name=""RandomWinners"" type=""radio"" value=""1""" & DisabledAttribute & IIf(CalcMethod = 1, " checked=""checked""", "") & " /><label for=""random"">" & Copient.PhraseLib.Lookup("term.random", LanguageID) & "</label>")
          End If
          Send("    </td>")
          Send("  </tr>")
          Send("</table>")
        %>
        <hr class="hidden" />
      </div>
    </div>
    
    <div id="gutter">
    </div>
    
    <div id="column2">
    </div>
  </div>
</form>

<script type="text/javascript">
<% If (CloseAfterSave) Then %>
    window.close();
<% End If %>
</script>

<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd("mainform", "NumPrizesAllowed")
  MyCommon = Nothing
  Logix = Nothing
%>