<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Text.RegularExpressions" %>
<%
  ' *****************************************************************************
  ' * FILENAME: web-offer-rew.aspx 
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
  Dim row As DataRow
  Dim rst2 As DataTable
  Dim row2 As DataRow
  Dim rst3 As DataTable
  Dim rstPT As DataTable
  Dim rowPT As DataRow
  Dim OfferID As Long
  Dim RewardID As Long
  Dim DeliverableID As Long
  Dim PKID As Long
  Dim MessageID As Long = 0
  Dim ParentROID As Long = 0
  Dim ROID As Long = 0
  Dim Name As String = ""
  Dim UrlTokens As String = ""
  Dim DeliverableType As Integer
  Dim AddOptionArray As New BitArray(8, True)
  Dim MessageTypeLabel As String = ""
  Dim index As Integer = 0
  Dim i As Integer = 0
  Dim DeleteBtnDisabled As String = ""
  Dim IsTemplate As Boolean = False
  Dim IsTemplateVal As String = "Not"
  Dim ActiveSubTab As Integer = 24
  Dim FromTemplate As Boolean = False
  Dim Disallow_Rewards As Boolean = False
  Dim RewardsCount As Integer = 0
  Dim IsCustomerAssigned As Boolean = False
  Dim AccumEligible As Boolean = False
  Dim infoMessage As String = ""
  Dim modMessage As String = ""
  Dim Handheld As Boolean = False
  Dim Rewards As String() = Nothing
  Dim LockedStatus As String() = Nothing
  Dim LoopCtr As Integer = 0
  Dim RewardDisabled As String = ""
  Dim BannersEnabled As Boolean = True
  Dim StatusCode As Copient.LogixInc.STATUS_FLAGS = Copient.LogixInc.STATUS_FLAGS.STATUS_UNKNOWN
  Dim StatusText As String = ""
  Dim TierLevels As Integer = 1
  Dim t As Integer = 1
  Dim DiscountWorthy As Boolean = False
  Dim EngineID As Integer = 3
  Dim EngineSubTypeID As Integer = 0
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "web-offer-rew.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  
  OfferID = Request.QueryString("OfferID")
  RewardID = Request.QueryString("RewardID")
  DeliverableID = MyCommon.Extract_Val(Request.QueryString("DeliverableID"))
  PKID = MyCommon.Extract_Val(Request.QueryString("PKID"))
  MessageID = MyCommon.Extract_Val(Request.QueryString("MessageID"))
  DeliverableType = MyCommon.Extract_Val(Request.QueryString("action"))
  
  If (OfferID = 0) Then
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "web-offer-gen.aspx")
  End If
  
  MyCommon.QueryStr = "select PassThruRewardID from PassThruRewards with (NoLock);"
  rst = MyCommon.LRT_Select
  AddOptionArray = New BitArray(8 + rst.Rows.Count, True)
  
  MyCommon.QueryStr = "select IncentiveName, IsTemplate, FromTemplate, EngineID, EngineSubTypeID " & _
                      "from CPE_Incentives with (NoLock) " & _
                      "where IncentiveID=" & OfferID & ";"
  rst = MyCommon.LRT_Select
  If (rst.Rows.Count > 0) Then
    Name = MyCommon.NZ(rst.Rows(0).Item("IncentiveName"), "")
    IsTemplate = MyCommon.NZ(rst.Rows(0).Item("IsTemplate"), False)
    FromTemplate = MyCommon.NZ(rst.Rows(0).Item("FromTemplate"), False)
    EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), 3)
    EngineSubTypeID = MyCommon.NZ(rst.Rows(0).Item("EngineSubTypeID"), 0)
  End If
  
  MyCommon.QueryStr = "select RewardOptionID, TierLevels from CPE_RewardOptions with (NoLock) " & _
                      "where IncentiveID=" & OfferID & " and TouchResponse=0 and Deleted=0;"
  rst = MyCommon.LRT_Select
  If (rst.Rows.Count > 0) Then
    ParentROID = MyCommon.NZ(rst.Rows(0).Item("RewardOptionID"), 0)
    If (RewardID = 0) Then
      RewardID = ParentROID
    End If
    TierLevels = MyCommon.NZ(rst.Rows(0).Item("TierLevels"), 1)
  End If
    
  MyCommon.QueryStr = "select CG.CustomerGroupID, Name, ExcludedUsers from CPE_IncentiveCustomerGroups as ICG with (NoLock) " & _
                      "left join CustomerGroups as CG with (NoLock) on CG.CustomerGroupID=ICG.CustomerGroupID " & _
                      "where RewardOptionID=" & RewardID & ";"
  rst = MyCommon.LRT_Select
  If (rst.Rows.Count > 0) Then
    IsCustomerAssigned = True
  End If
  
  If (Request.QueryString("mode") = "DeleteMembership") Then
    If (DeliverableID > 0 AndAlso DeliverableType > 0 AndAlso OfferID > 0) Then
      MyCommon.QueryStr = "delete from CPE_DeliverableCustomerGroupTiers with (RowLock) where DeliverableID=" & DeliverableID & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "delete from CPE_Deliverables with (RowLock) where RewardOptionID=" & RewardID & " and RewardOptionPhase=3 and DeliverableTypeID=" & DeliverableType & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
      MyCommon.LRT_Execute()
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Reward.deletemembership", LanguageID))
    End If
  ElseIf (Request.QueryString("mode") = "DeletePassThru") Then
    If (DeliverableID > 0 AndAlso OfferID > 0) Then
      MyCommon.QueryStr = "delete from PassThruTierValues where PTPKID in (select PKID from PassThrus where DeliverableID=" & DeliverableID & ");"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "delete from PassThruTiers where PTPKID in " & _
                          "(select PKID from PassThrus where DeliverableID=" & DeliverableID & ");"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "delete from PassThrus where DeliverableID=" & DeliverableID & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "delete from CPE_Deliverables with (RowLock) where DeliverableID=" & DeliverableID & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
      MyCommon.LRT_Execute()
    End If
    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Reward.deletepassthru", LanguageID))
  End If
  
  'update the template permission for Conditions
  If (Request.QueryString("save") <> "") Then
    If (Request.QueryString("IsTemplate") = "IsTemplate") Then
      ' time to update the status bits for the templates
      Dim form_Disallow_Rewards As Integer = 0
      If (Request.QueryString("Disallow_Rewards") = "on") Then
        form_Disallow_Rewards = 1
      End If
      MyCommon.QueryStr = "update TemplatePermissions with (RowLock) set Disallow_Rewards=" & form_Disallow_Rewards & _
                          " where OfferID=" & OfferID & ";"
      MyCommon.LRT_Execute()
    End If
    
    'Update the lock status for each condition
    Rewards = Request.QueryString.GetValues("rew")
    LockedStatus = Request.QueryString.GetValues("locked")
    If (Not Rewards Is Nothing AndAlso Not LockedStatus Is Nothing AndAlso Rewards.Length = LockedStatus.Length) Then
      For LoopCtr = 0 To Rewards.GetUpperBound(0)
        MyCommon.QueryStr = "update CPE_Deliverables with (RowLock) set DisallowEdit=" & LockedStatus(LoopCtr) & " " & _
                            "where DeliverableID=" & Rewards(LoopCtr) & ";"
        MyCommon.LRT_Execute()
      Next
    End If
    
  End If
  
  If (IsTemplate Or FromTemplate) Then
    ' lets dig the permissions if its a template
    MyCommon.QueryStr = "select * from TemplatePermissions with (NoLock) where OfferID=" & OfferID & ";"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      For Each row In rst.Rows
        ' ok there are some rows for the template
        Disallow_Rewards = MyCommon.NZ(row.Item("Disallow_Rewards"), True)
      Next
    End If
  End If
  
  If (IsTemplate) Then
    ActiveSubTab = 27
    IsTemplateVal = "IsTemplate"
  Else
    ActiveSubTab = 27
    IsTemplateVal = "Not"
  End If
  
  DeleteBtnDisabled = IIf(Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Rewards), "", " disabled=""disabled""")
  SetDeleteBtnDisabled(DeleteBtnDisabled) 'method found in included file GraphicReward.aspx
  
  Send_HeadBegin("term.offer", "term.rewards", OfferID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  If (IsTemplate) Then
    Send("<style type=""text/css"">")
    Send("  * html #tblReward {")
    Send("    position: relative;")
    Send("    top: -20px;")
    Send("  }")
    Send("</style>")
  End If
  Send_Scripts()
%>
<script type="text/javascript">
function LoadDocument(url) { 
  location = url;
}

function openTouchptReward(index, roid) {
  var pageName = "";
  var qryStr = "";
  var tpElem = document.getElementById("newrewtouchpt" + index);
  
  if (tpElem != null) {
    var rewType = tpElem.options[tpElem.options.selectedIndex].value
    
    if (rewType >= 1 && rewType <= 8) {
      qryStr = "?RewardID=<%Sendb(RewardID)%>&OfferID=<%Sendb(OfferID)%>&tp=1&roid=" + roid;
      if (rewType == 5) {
        qryStr += "&action=5"
        pageName = "CPEoffer-rew-membership.aspx";
      /* Disabled Revoke Membership to comply with current Logix 3.9.2 functionality
      } else if (rewType == 6) {
        pageName = "CPEoffer-rew-membership.aspx";
      */
      }
      openPopup(pageName + qryStr);
    }
  }
}

function updateLocked(elemName, bChecked) {
  var elem = document.getElementById(elemName);
  
  if (elem != null) {
    elem.value = (bChecked) ? "1" : "0";
  }
}
</script>
<%
  Send_HeadEnd()
  If (IsTemplate) Then
    Send_BodyBegin(11)
  Else
    Send_BodyBegin(1)
  End If
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  Send_Tabs(Logix, 2)
  Send_Subtabs(Logix, ActiveSubTab, 6, , OfferID)
  
  If (Logix.UserRoles.AccessOffers = False AndAlso Not IsTemplate) Then
    Send_Denied(1, "perm.offers-access")
    GoTo done
  End If
  If (Logix.UserRoles.AccessTemplates = False AndAlso IsTemplate) Then
    Send_Denied(1, "perm.offers-access-templates")
    GoTo done
  End If
  If (BannersEnabled AndAlso Not Logix.IsAccessibleOffer(AdminUserID, OfferID)) Then
    Send("<script type=""text/javascript"" language=""javascript"">")
    Send("  function updateCookie() { return true; } ")
    Send("</script>")
    Send_Denied(1, "banners.access-denied-offer")
    Send_BodyEnd()
    GoTo done
  End If
  
  If (Request.QueryString("addGlobal") <> "") Then
    Dim rewChoice As Integer = MyCommon.Extract_Val(Request.QueryString("newrewglobal"))
    If (rewChoice > 0) Then
      Select Case rewChoice
        Case 5
          Send("<script type=""text/javascript"">openPopup('CPEoffer-rew-membership.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "&action=5&Phase=3')</script>")
          MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-member", LanguageID))
        Case 6
          'Disabled Revoke Membership to comply with current Logix 3.9.2 functionality
          'Send("<script type=""text/javascript"">openPopup('CPEoffer-rew-membership.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "&action=6')</script>")
          'MyCommon.Activity_Log(3, OfferID, AdminUserID, "Added global membership reward")
      End Select
    End If
  End If
  
  If (Request.QueryString("addGlobal") <> "") Then
    Dim rewChoice As Integer = MyCommon.Extract_Val(Request.QueryString("newrewglobal"))
    If (rewChoice > 0) Then
      Select Case rewChoice
        Case 1
          Send("<script type=""text/javascript"">openPopup('CPEoffer-rew-graphic.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "&Phase=3')</script>")
          MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-graphic", LanguageID))
        Case 2
          Send("<script type=""text/javascript"">openPopup('CPEoffer-rew-discount.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "&Phase=3')</script>")
          MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-discount", LanguageID))
        Case 3
          'Not used
        Case 4
          Send("<script type=""text/javascript"">openPopup('CPEoffer-rew-pmsg.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "&Phase=3')</script>")
          MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-pmsg", LanguageID))
        Case 5
          Send("<script type=""text/javascript"">openPopup('CPEoffer-rew-membership.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "&action=5&Phase=3')</script>")
          MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-member", LanguageID))
        Case 6
          'Revoke membership -- not used
        Case 7
          'Silent deliverable -- not used
        Case 8
          Send("<script type=""text/javascript"">openPopup('CPEoffer-rew-point.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "&Phase=3&New=1')</script>")
          MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-points", LanguageID))
        Case 9
          Send("<script type=""text/javascript"">openPopup('CPEoffer-rew-cmsg.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "&Phase=3')</script>")
          MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-cmsg", LanguageID))
        Case 10
          Send("<script type=""text/javascript"">openPopup('CPEoffer-rew-franking.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "&Phase=3')</script>")
          MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-franking", LanguageID))
        Case 11
          Send("<script type=""text/javascript"">openPopup('CPEoffer-rew-sv.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "&Phase=3')</script>")
          MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-sv", LanguageID))
        Case Else
          MyCommon.QueryStr = "select PassThruRewardID, Name, PhraseID from PassThruRewards with (NoLock) order by PassThruRewardID;"
          rstPT = MyCommon.LRT_Select
          If rstPT.Rows.Count > 0 Then
            Send("<script type=""text/javascript"">openPopup('CPEoffer-rew-passthru.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "&PassThruRewardID=" & (rewChoice - 12) & "&Phase=3')</script>")
            MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-passthru", LanguageID))
          End If
      End Select
    End If
  End If
%>
<form action="web-offer-rew.aspx" id="mainform" name="mainform">
  <input type="hidden" id="OfferID" name="OfferID" value="<% Sendb(OfferID) %>" />
  <input type="hidden" id="IsTemplate" name="IsTemplate" value="<% Sendb(IsTemplateVal) %>" />
  <div id="intro">
    <%
      If (IsTemplate) Then
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & ": " & MyCommon.TruncateString(Name, 50) & "</h1>")
      Else
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & ": " & MyCommon.TruncateString(Name, 50) & "</h1>")
      End If
    %>
    <div id="controls">
      <%
        If (Logix.UserRoles.EditOffer And IsTemplate) Then
          Send_Save()
        End If
        If MyCommon.Fetch_SystemOption(75) Then
          If (OfferID > 0 And Logix.UserRoles.AccessNotes) Then
            Send_NotesButton(3, OfferID, AdminUserID)
          End If
        End If
      %>
    </div>
  </div>
  <div id="main">
    <%
      MyCommon.QueryStr = "select StatusFlag from CPE_Incentives where IncentiveID=" & OfferID & ";"
      rst = MyCommon.LRT_Select
      StatusText = Logix.GetOfferStatus(OfferID, LanguageID, StatusCode)
      If Not IsTemplate Then
        If (MyCommon.NZ(rst.Rows(0).Item("StatusFlag"), 0) <> 2) Then
          If (StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_ACTIVE) AndAlso (MyCommon.NZ(rst.Rows(0).Item("StatusFlag"), 0) > 0) Then
            modMessage = Copient.PhraseLib.Lookup("alert.modpostdeploy", LanguageID)
            Send("<div id=""modbar"">" & modMessage & "</div>")
          End If
        End If
      End If
      If (infoMessage <> "") Then
        Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
      End If
      
      ' Send the status bar, but not if it's a new offer or a template, or if there's already a modMessage being shown.
      If (Not IsTemplate AndAlso modMessage = "") Then
        MyCommon.QueryStr = "select IncentiveID from CPE_Incentives with (NoLock) where CreatedDate=LastUpdate and IncentiveID=" & OfferID & ";"
        rst3 = MyCommon.LRT_Select
        If (rst3.Rows.Count = 0) Then
          Send_Status(OfferID, 2)
        End If
      End If
    %>
    <div id="column">
      <div class="box" id="rewards">
        <h2>
          <span><% Sendb(Copient.PhraseLib.Lookup("term.rewards", LanguageID))%></span>
        </h2>
        <br />
        <table class="list" id="tblReward" summary="<% Sendb(Copient.PhraseLib.Lookup("term.rewards", LanguageID)) %>">
          <thead>
            <tr>
              <th align="left" scope="col" class="th-del">
                <% Sendb(Copient.PhraseLib.Lookup("term.delete", LanguageID))%>
              </th>
              <th align="left" scope="col" class="th-type">
                <% Sendb(Copient.PhraseLib.Lookup("term.type", LanguageID))%>
              </th>
              <th align="left" scope="col" class="th-subtype">
                <% Sendb(Copient.PhraseLib.Lookup("term.subtype", LanguageID))%>
              </th>
              <th align="left" scope="col" class="th-details" colspan="<% Sendb(TierLevels) %>">
                <% Sendb(Copient.PhraseLib.Lookup("term.details", LanguageID))%>
              </th>
              <% If (isTemplate OrElse FromTemplate) Then%>
              <th align="left" scope="col" class="th-locked">
                <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
              </th>
              <% End If%>               
            </tr>
          </thead>
          <tbody>
            <%
              'Group membership rewards
              t = 1
              MyCommon.QueryStr = "select D.DeliverableID, D.DeliverableTypeID, D.RewardOptionID as ROID, D.DisallowEdit " & _
                                  "from CPE_Deliverables as D with (NoLock) " & _
                                  "where D.RewardOptionID=" & ParentROID & " and D.DeliverableTypeID in (5) and D.RewardOptionPhase=3 " & _
                                  "order by D.DeliverableTypeID;"
              rst = MyCommon.LRT_Select()
              If (rst.Rows.Count > 0) Then
                RewardsCount += rst.Rows.Count
                AddOptionArray.Set(4, False)
                Send("<tr class=""shadeddark"">")
                Send("  <td colspan=""" & 3 & """>")
                Send("    <h3>" & Copient.PhraseLib.Lookup("cpeoffer-rew-grpmbr", LanguageID) & "</h3>")
                Send("  </td>")
                If TierLevels = 1 Then
                  Send("  <td></td>")
                Else
                  For t = 1 To TierLevels
                    Send("  <td>")
                    Send("    <b>" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</b>")
                    Send("  </td>")
                  Next
                End If
                t = 1
                If (IsTemplate Or FromTemplate) Then
                  Send("  <td></td>")
                End If
                Send("</tr>")
                For Each row In rst.Rows
                  RewardDisabled = IIf((Logix.UserRoles.EditOffer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False))), "", " disabled=""disabled""")
                  UrlTokens = "?OfferID=" & OfferID & "&RewardID=" & MyCommon.NZ(row.Item("ROID"), 0) & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), -1) & "&action=" & MyCommon.NZ(row.Item("DeliverableTypeID"), 0)
                  Send("<tr class=""shaded"">")
                  Send("  <td><input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & RewardDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("reward.confirmdelete", LanguageID) & "')) LoadDocument('web-offer-rew.aspx" & UrlTokens & "&mode=DeleteMembership')"" value=""X"" /></td>")
                  Send("  <td><a href=""javascript:openPopup('CPEoffer-rew-membership.aspx" & UrlTokens & "')"">" & Copient.PhraseLib.Lookup("term.membership", LanguageID) & "</a></td>")
                  Send("  <td>")
                  'Send(IIf(MyCommon.NZ(row.Item("DeliverableTypeID"), 0) = 6, "Remove", "Add"))
                  Send("  </td>")
                  ' Find the per-tier values:
                  MyCommon.QueryStr = "select DCGT.TierLevel, DCGT.CustomerGroupID, CG.Name " & _
                                      "from CPE_DeliverableCustomerGroupTiers as DCGT with (NoLock) " & _
                                      "left join CustomerGroups as CG with (NoLock) on CG.CustomerGroupID=DCGT.CustomerGroupID " & _
                                      "where DCGT.DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & ";"
                  rst2 = MyCommon.LRT_Select
                  If rst2.Rows.Count = 0 Then
                    Send("  <td colspan=""" & TierLevels & """>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                  Else
                    While t <= TierLevels
                      If t > rst2.Rows.Count Then
                        Send("  <td>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                      Else
                        Send("  <td><a href=""cgroup-edit.aspx?CustomerGroupID=" & MyCommon.NZ(rst2.Rows(t - 1).Item("CustomerGroupID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(rst2.Rows(t - 1).Item("Name"), ""), 25) & "</a></td>")
                      End If
                      t += 1
                    End While
                  End If
                  t = 1
                  If (IsTemplate) Then
                    Send("  <td class=""templine"">")
                    Send("    <input type=""checkbox"" id=""chkLocked7"" name=""chkLocked"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, " checked=""checked""", "") & " onclick=""javascript:updateLocked('lockGM" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "', this.checked);"" />")
                    Send("    <input type=""hidden"" id=""rewGM" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""rew"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ />")
                    Send("    <input type=""hidden"" id=""lockGM" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""locked"" value=""" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "1", "0") & """ />")
                    Send("  </td>")
                  ElseIf (FromTemplate) Then
                    Send("  <td class=""templine"">" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "Yes", "No") & "</td>")
                  End If
                  Send("</tr>")
                Next
              End If
              
              ' Pass-thru reward(s)
              MyCommon.QueryStr = "select PassThruRewardID, Name, PhraseID from PassThruRewards with (NoLock) order by Name;"
              rstPT = MyCommon.LRT_Select
              If rstPT.Rows.Count > 0 Then
                i = 1
                For Each rowPT In rstPT.Rows
                  t = 1
                  MyCommon.QueryStr = "select D.DeliverableID, D.DisallowEdit, DPT.PassThruRewardID, PTR.Name, PTR.PhraseID, PTR.LSInterfaceID, PTR.ActionTypeID " & _
                                      "from CPE_Deliverables as D with (NoLock) " & _
                                      "inner join PassThrus as DPT with (NoLock) on DPT.PKID=D.OutputID " & _
                                      "inner join PassThruRewards as PTR with (NoLock) on PTR.PassThruRewardID=DPT.PassThruRewardID " & _
                                      "where D.RewardOptionID=" & ParentROID & " and DPT.PassThruRewardID=" & MyCommon.NZ(rowPT.Item("PassThruRewardID"), 0) & " and D.Deleted=0 and DeliverableTypeID=12 " & _
                                      "order by Name;"
                  rst = MyCommon.LRT_Select()
                  If (rst.Rows.Count > 0) Then
                    RewardsCount += rst.Rows.Count
                    
                    
                    
                    AddOptionArray.Set(8 + (i - 1), False)
                    Send("<tr class=""shadeddark"">")
                    Send("  <td colspan=""" & 3 & """>")
                    If IsDBNull(rowPT.Item("PhraseID")) Then
                      Send("    <h3>" & MyCommon.NZ(rowPT.Item("Name"), Copient.PhraseLib.Lookup("term.passthrureward", LanguageID)) & "</h3>")
                    Else
                      Send("    <h3>" & Copient.PhraseLib.Lookup(rowPT.Item("PhraseID"), LanguageID) & "</h3>")
                    End If
                    Send("  </td>")
                    If TierLevels = 1 Then
                      Send("  <td></td>")
                    Else
                      For t = 1 To TierLevels
                        Send("  <td>")
                        Send("    <b>" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</b>")
                        Send("  </td>")
                      Next
                    End If
                    If (IsTemplate Or FromTemplate) Then
                      Send("  <td></td>")
                    End If
                    Send("</tr>")
                    t = 1
                    For Each row In rst.Rows
                      RewardDisabled = IIf((Logix.UserRoles.EditOffer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False))), "", " disabled=""disabled""")
                      Send("<tr class=""shaded"">")
                      Sendb("  <td><input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & RewardDisabled & " value=""X"" ")
                      Send("onClick=""if(confirm('" & Copient.PhraseLib.Lookup("reward.confirmdelete", LanguageID) & "')) LoadDocument('web-offer-rew.aspx?mode=DeletePassThru&OfferID=" & OfferID & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "&PassThruRewardID=" & MyCommon.NZ(row.Item("PassThruRewardID"), 0) & "')"" /></td>")
                      Sendb("  <td><a href=""javascript:openPopup('CPEoffer-rew-passthru.aspx?OfferID=" & OfferID & "&RewardID=" & RewardID & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "&PassThruRewardID=" & MyCommon.NZ(row.Item("PassThruRewardID"), "") & "')"">")
                      If IsDBNull(rowPT.Item("PhraseID")) Then
                        Send(MyCommon.NZ(rowPT.Item("Name"), Copient.PhraseLib.Lookup("term.passthrureward", LanguageID)))
                      Else
                        Send(Copient.PhraseLib.Lookup(rowPT.Item("PhraseID"), LanguageID))
                      End If
                      Send("</a></td>")
                      Send("  <td></td>")
                      ' Find the per-tier values:
                      MyCommon.QueryStr = "select DPT.PKID, DPTT.TierLevel, DPTT.Data " & _
                                          "from PassThrus as DPT with (NoLock) " & _
                                          "inner join PassThruTiers as DPTT with (NoLock) on DPT.PKID=DPTT.PTPKID " & _
                                          "where DPT.DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & ";"
                      rst2 = MyCommon.LRT_Select
                      If rst2.Rows.Count = 0 Then
                        Send("  <td colspan=""" & TierLevels & """>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                      Else
                        While t <= TierLevels
                          If t > rst2.Rows.Count Then
                            Send("  <td>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                          Else
                            Send("  <td><pre style=""font-size:11px;"">" & MyCommon.NZ(rst2.Rows(t - 1).Item("Data"), "").ToString.Replace("<", "&lt;") & "</pre></td>")
                          End If
                          t += 1
                        End While
                      End If
                      t = 1
                      If (IsTemplate) Then
                        Send("  <td class=""templine"">")
                        Send("    <input type=""checkbox"" id=""chkLocked8"" name=""chkLocked"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, " checked=""checked""", "") & " onclick=""javascript:updateLocked('lockPts" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "', this.checked);"" />")
                        Send("    <input type=""hidden"" id=""rewPassthru" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""rew"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ />")
                        Send("    <input type=""hidden"" id=""lockPassthru" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""locked"" value=""" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "1", "0") & """ />")
                        Send("  </td>")
                      ElseIf (FromTemplate) Then
                        Send("  <td class=""templine"">" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "Yes", "No") & "</td>")
                      End If
                      Send("</tr>")
                    Next
                  End If
                  i += 1
                Next
              End If
            %>
          </tbody>
        </table>
        <hr class="hidden" />
      </div>
      
      <div class="box" id="newreward"<% Sendb(IIf(IsCustomerAssigned AndAlso AddOptionArray.Get(4), "", " style=""display:none;""")) %>>
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("offer-rew.addreward", LanguageID))%>
          </span>
        </h2>
        <%
          If IsTemplate Then
            Send("<span class=""temp"">")
            Send("  <input type=""checkbox"" class=""tempcheck"" id=""Disallow_Rewards"" name=""Disallow_Rewards""" & IIf(Disallow_Rewards, " checked=""checked""", "") & " />")
            Send("  <label for=""Disallow_Rewards"">" & Copient.PhraseLib.Lookup("term.locked", LanguageID) & "</label>")
            Send("</span>")
          End If
            MyCommon.QueryStr = "SELECT PECT.EngineID, PECT.EngineSubTypeID, PECT.ComponentTypeID, DT.DeliverableTypeID, DT.Description, DT.PhraseID, PECT.Singular, " & _
                                "  CASE DeliverableTypeID " & _
                                "    WHEN 1 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=1) " & _
                                "    WHEN 2 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=2) " & _
                                "    WHEN 4 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=4) " & _
                                "    WHEN 5 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=5) " & _
                                "    WHEN 6 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=6) " & _
                                "    WHEN 7 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=7) " & _
                                "    WHEN 8 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=8) " & _
                                "    WHEN 9 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=9) " & _
                                "    WHEN 10 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=10) " & _
                                "    WHEN 11 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=11) " & _
                                "    WHEN 12 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=12) " & _
                                "    ELSE 0 " & _
                                "  END as ItemCount " & _
                                "FROM PromoEngineComponentTypes AS PECT " & _
                                "INNER JOIN CPE_DeliverableTypes AS DT ON DT.DeliverableTypeID=PECT.LinkID " & _
                                "WHERE EngineID=3 AND EngineSubTypeID=" & EngineSubTypeID & " AND PECT.ComponentTypeID=2 AND Enabled=1"
          'Impose a few special limits on the query based on various factors:
          If (Not IsCustomerAssigned) Then
            'The offer has no customer condition, so the only available reward is graphics
            MyCommon.QueryStr &= " AND DeliverableTypeID=1"
          End If
          MyCommon.QueryStr &= " ORDER BY DisplayOrder;"
          rst = MyCommon.LRT_Select
          If rst.Rows.Count > 0 Then
            Send("<label for=""newrewglobal"">" & Copient.PhraseLib.Lookup("offer-rew.addglobal", LanguageID) & ":</label><br />")
            Send("<select id=""newrewglobal"" name=""newrewglobal""" & DeleteBtnDisabled & ">")
            For Each row In rst.Rows
              If (row.Item("Singular") = False) OrElse (row.Item("Singular") = True AndAlso row.Item("ItemCount") = 0) Then
                If (row.Item("DeliverableTypeID") = 2 AndAlso Not DiscountWorthy) Then
                  'Discount disallowed (due to the presence of a tender condition or some other non-discountworthy factor).
                ElseIf row.Item("DeliverableTypeID") = 12 Then
                  'Type 12 is passthrus -- a special case, since each passthru must be shown as its own reward type
                  MyCommon.QueryStr = "select PassThruRewardID, Name, PhraseID from PassThruRewards with (NoLock) order by PassThruRewardID;"
                  rst2 = MyCommon.LRT_Select
                  If rst2.Rows.Count > 0 Then
                    For Each row2 In rst2.Rows
                      Sendb("<option value=""" & (row2.Item("PassThruRewardID") + 12) & """>")
                      If IsDBNull(row2.Item("PhraseID")) Then
                        Sendb(MyCommon.NZ(row2.Item("Name"), Copient.PhraseLib.Lookup("term.passthru", LanguageID)))
                      Else
                        Sendb(Copient.PhraseLib.Lookup(row2.Item("PhraseID"), LanguageID))
                      End If
                      Send("</option>")
                    Next
                  End If
                Else
                  'All the other types
                  Send("<option value=""" & row.Item("DeliverableTypeID") & """>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "</option>")
                End If
              End If
            Next
            Send("</select>")
            Send("<input class=""regular"" id=""addglobal"" name=""addglobal"" type=""submit"" value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """" & DeleteBtnDisabled & " /><br />")
          End If
        %>
      </div>
    </div>
    <br clear="all" />
  </div>
</form>
<!-- #Include virtual="/include/graphic-reward.inc" -->
<%
  If (RewardsCount = 0 AndAlso IsTemplate) Then
    Send("<script type=""text/javascript"">")
    Send("document.getElementById(""tblReward"").style.display = 'none';")
    Send("</script>")
  ElseIf (RewardsCount = 0) Then
    Send("<script type=""text/javascript"">")
    Send("document.getElementById(""rewards"").style.display = 'none';")
    Send("</script>")
  End If
%>
<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (OfferID > 0 And Logix.UserRoles.AccessNotes) Then
      Send_Notes(3, OfferID, AdminUserID)
    End If
  End If
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd()
  MyCommon = Nothing
  Logix = Nothing
%>
