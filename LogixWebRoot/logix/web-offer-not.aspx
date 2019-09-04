<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: web-offer-not.aspx 
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
  Dim rst3 As DataTable
  Dim OfferID As Long
  Dim RewardID As Long
  Dim Name As String = ""
  Dim ConditionID As Integer
  Dim isTemplate As Boolean = False
  Dim FromTemplate As Boolean = False
  Dim roid As Integer
  Dim DeleteGraphicURL As String = ""
  Dim AddTouchPtURL As String = ""
  Dim UrlTokens As String = ""
  Dim DeliverableType As Integer
  Dim AddOptionArray As New BitArray(3, True)
  Dim MessageTypeLabel As String = ""
  Dim DeliverableID As Long
  Dim PKID As Long
  Dim MessageID As Long
  Dim index As Integer = 0
  Dim tpROID As Integer
  Dim Phase As Integer
  Dim Details As StringBuilder
  Dim DeleteBtnDisabled As String = ""
  Dim Disallow_Notifications As Boolean = False
  Dim ActiveSubTab As Integer = 91
  Dim NotificationCount As Integer = 0
  Dim AccumCount As Integer = 0
  Dim IsCustomerAssigned As Boolean = False
  Dim infoMessage As String = ""
  Dim modMessage As String = ""
  Dim Handheld As Boolean = False
  Dim Notifications As String() = Nothing
  Dim LockedStatus As String() = Nothing
  Dim LoopCtr As Integer = 0
  Dim NotificationDisabled As String = ""
  Dim BannersEnabled As Boolean = True
  Dim HasPointsReward As Boolean = False
  Dim StatusCode As Copient.LogixInc.STATUS_FLAGS = Copient.LogixInc.STATUS_FLAGS.STATUS_UNKNOWN
  Dim StatusText As String = ""
  Dim TierLevels As Integer = 1
  Dim t As Integer = 1
  Dim EngineID As Integer = 3
  Dim EngineSubTypeID As Integer = 0
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "web-offer-not.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  
  OfferID = Request.QueryString("OfferID")
  RewardID = Request.QueryString("RewardID")
  DeliverableID = MyCommon.Extract_Val(Request.QueryString("DeliverableID"))
  PKID = MyCommon.Extract_Val(Request.QueryString("PKID"))
  MessageID = MyCommon.Extract_Val(Request.QueryString("MessageID"))
  Phase = MyCommon.Extract_Val(Request.QueryString("Phase"))
  DeliverableType = MyCommon.Extract_Val(Request.QueryString("action"))
  
  If (OfferID = 0) Then
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "web-offer-gen.aspx")
  End If
  
  MyCommon.QueryStr = "select RewardOptionID, TierLevels from CPE_RewardOptions with (NoLock) " & _
                      "where IncentiveID=" & OfferID & " and TouchResponse=0 and Deleted=0;"
  rst = MyCommon.LRT_Select
  If (rst.Rows.Count > 0) Then
    If (RewardID = 0) Then
      RewardID = MyCommon.NZ(rst.Rows(0).Item("RewardOptionID"), 0)
      roid = MyCommon.NZ(rst.Rows(0).Item("RewardOptionID"), 0)
      'TierLevels = MyCommon.NZ(rst.Rows(0).Item("TierLevels"), 1)
    End If
  End If
  TierLevels = 1 ' -- hardcoding TierLevels to 1, since notifications can't be tiered. -hjw
  
  ' determine if customer condition is set for the offer
  MyCommon.QueryStr = "select CG.CustomerGroupID,Name,ExcludedUsers from CPE_IncentiveCustomerGroups as ICG with (NoLock) " & _
                      "left join CustomerGroups as CG with (NoLock) on CG.CustomerGroupID=ICG.CustomerGroupID " & _
                      "where RewardOptionID=" & roid & " and ICG.Deleted=0;"
  rst = MyCommon.LRT_Select
  If (rst.Rows.Count > 0) Then
    IsCustomerAssigned = True
  End If
  
  Send_HeadBegin("term.offer", "term.notifications", OfferID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
  Send_HeadEnd()
  
  'update the template permission for Conditions
  If (Request.QueryString("Save") <> "") Then
    If (Request.QueryString("IsTemplate") = "IsTemplate") Then
      ' time to update the status bits for the templates
      Dim form_Disallow_Notifications As Integer = 0
      If (Request.QueryString("Disallow_Notifications") = "on") Then
        form_Disallow_Notifications = 1
      End If
      MyCommon.QueryStr = "update TemplatePermissions with (RowLock) set Disallow_Notifications=" & form_Disallow_Notifications & _
                          " where OfferID=" & OfferID & ";"
      MyCommon.LRT_Execute()
      'Update the lock status for each condition
      Notifications = Request.QueryString.GetValues("rew")
      LockedStatus = Request.QueryString.GetValues("locked")
      If (Not Notifications Is Nothing AndAlso Not LockedStatus Is Nothing AndAlso Notifications.Length = LockedStatus.Length) Then
        For LoopCtr = 0 To Notifications.GetUpperBound(0)
          MyCommon.QueryStr = "update CPE_Deliverables with (RowLock) set DisallowEdit=" & LockedStatus(LoopCtr) & " " & _
                              "where DeliverableID=" & Notifications(LoopCtr) & ";"
          MyCommon.LRT_Execute()
        Next
      End If
    End If
  Else
    ' handle adding stuff on
    If (Request.QueryString("neweliglobal") <> "") Then
      If (Request.QueryString("neweliglobal") = 1) Then
        ' time to add a new printed to the CPE_IncentiveCustomerGroups table
        Send("<script type=""text/javascript"">openPopup('CPEoffer-rew-graphic.aspx?Phase=1&OfferID=" & OfferID & "&RewardId=" & roid & "')</script>")
      ElseIf (Request.QueryString("neweliglobal") = 9) Then
        Send("<script type=""text/javascript"">openPopup('CPEoffer-rew-cmsg.aspx?Phase=1&OfferID=" & OfferID & "&RewardId=" & roid & "')</script>")
      'ElseIf (Request.QueryString("neweliglobal") = 3) Then
      '  Send("<script type=""text/javascript"">openPopup('CPEoffer-rew-graphic.aspx?Phase=1&OfferID=" & OfferID & "&RewardId=" & roid & "')</script>")
      ElseIf (Request.QueryString("neweliglobal") = 4) Then
        Send("<script type=""text/javascript"">openPopup('CPEoffer-rew-pmsg.aspx?Phase=1&OfferID=" & OfferID & "&RewardId=" & roid & "')</script>")
      End If
    End If
  End If
  
  If (Request.QueryString("mode") = "DeleteGraphic") Then
    RemoveGraphic(OfferID, DeliverableID)
    MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where IncentiveID=" & OfferID & ";"
    MyCommon.LRT_Execute()
    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Notification.deletegraphic", LanguageID))
  ElseIf (Request.QueryString("mode") = "DeletePrintedMsg") Then
    If (DeliverableID > 0 AndAlso OfferID > 0) Then
      MyCommon.QueryStr = "delete from PrintedMessageTiers with (RowLock) where MessageID=" & MessageID & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "delete from PrintedMessages with (RowLock) where MessageID=" & MessageID & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "delete from CPE_Deliverables with (RowLock) where DeliverableID=" & DeliverableID & " and RewardOptionPhase=" & Phase & " and DeliverableTypeID=4;"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
      MyCommon.LRT_Execute()
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Notification.deletepmsg", LanguageID))
    End If
  End If
  
  ConditionID = Request.QueryString("ConditionID")
  ' dig the offer info out of the database
  ' no one clicked anything
  MyCommon.QueryStr = "select IncentiveID, ClientOfferID, IncentiveName as Name, CPE.Description, PromoClassID, CRMEngineID, Priority, " & _
                      "StartDate, EndDate, EveryDOW, EligibilityStartDate, EligibilityEndDate, TestingStartDate, TestingEndDate, " & _
                      "P1DistQtyLimit, P1DistTimeType, P1DistPeriod, P2DistQtyLimit, P2DistTimeType, P2DistPeriod, P3DistQtyLimit, P3DistTimeType, P3DistPeriod, " & _
                      "EnableImpressRpt, EnableRedeemRpt, CreatedDate, CPE.LastUpdate, CPE.Deleted, CPEOARptDate, CPEOADeploySuccessDate, CPEOADeployRpt, " & _
                      "CRMRestricted, StatusFlag, OC.Description as CategoryName, IsTemplate, FromTemplate, EngineID, EngineSubTypeID " & _
                      "from CPE_Incentives as CPE with (NoLock) " & _
                      "left join OfferCategories as OC with (NoLock) on CPE.PromoClassID=OfferCategoryID " & _
                      "where IncentiveID=" & OfferID & ";"
  rst = MyCommon.LRT_Select
  For Each row In rst.Rows
    Name = MyCommon.NZ(row.Item("Name"), "")
    isTemplate = MyCommon.NZ(row.Item("IsTemplate"), False)
    FromTemplate = MyCommon.NZ(row.Item("FromTemplate"), False)
    EngineID = MyCommon.NZ(row.Item("EngineID"), 3)
    EngineSubTypeID = MyCommon.NZ(row.Item("EngineSubTypeID"), 0)
  Next
  
  If (isTemplate Or FromTemplate) Then
    ' lets dig the permissions if its a template
    MyCommon.QueryStr = "select Disallow_Notifications from TemplatePermissions with (NoLock) where OfferID=" & OfferID & ";"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      For Each row In rst.Rows
        ' ok there are some rows for the template
        Disallow_Notifications = MyCommon.NZ(row.Item("Disallow_Notifications"), True)
      Next
    End If
  End If
  
  DeleteBtnDisabled = IIf(Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Notifications), "", " disabled=""disabled""")
  SetDeleteBtnDisabled(DeleteBtnDisabled) 'method found in included file GraphicReward.aspx
  
  If (isTemplate) Then
    ActiveSubTab = 27
  Else
    ActiveSubTab = 27
  End If
  
  If (isTemplate) Then
    Send_BodyBegin(11)
  Else
    Send_BodyBegin(1)
  End If
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  Send_Tabs(Logix, 2)
  Send_Subtabs(Logix, ActiveSubTab, 7, , OfferID)
  
  If (Logix.UserRoles.AccessOffers = False AndAlso Not isTemplate) Then
    Send_Denied(1, "perm.offers-access")
    GoTo done
  End If
  If (Logix.UserRoles.AccessTemplates = False AndAlso isTemplate) Then
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
    
    if (rewType >= 1 && rewType <= 7) {
      qryStr = "?RewardID=<%Sendb(RewardID)%>&OfferID=<%Sendb(OfferID)%>&Phase=1&tp=1&roid=" + roid;
      if (rewType == 1) {
        pageName = "CPEoffer-rew-discount.aspx";
      } else if (rewType == 2) {
        pageName = "CPEoffer-rew-point.aspx";
      } else if (rewType == 3) {
        pageName = "CPEoffer-rew-pmsg.aspx";
      } else if (rewType == 4) {
        pageName = "CPEoffer-rew-cmsg.aspx";
      } else if (rewType == 5) {
        qryStr += "&action=5"
        pageName = "CPEoffer-rew-membership.aspx";
      /* Disabled Revoke Membership to comply with current Logix 3.9.2 functionality
      } else if (rewType == 6) {
        pageName = "CPEoffer-rew-membership.aspx";
      */
      } else if (rewType ==7) {
        pageName = "CPEoffer-rew-graphic.aspx";
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

<form action="web-offer-not.aspx" id="mainform" name="mainform">
  <input type="hidden" id="OfferID" name="OfferID" value="<% Sendb(OfferID) %>" />
  <input type="hidden" id="IsTemplate" name="IsTemplate" value="<% 
        If (IsTemplate) Then
        Sendb("IsTemplate")
        Else
        Sendb("Not")
        End If
        %>" />
  <div id="intro">
    <%
      If (isTemplate) Then
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & ": " & MyCommon.TruncateString(Name, 50) & "</h1>")
      Else
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & ": " & MyCommon.TruncateString(Name, 50) & "</h1>")
      End If
    %>
    <div id="controls">
      <%
        If (Logix.UserRoles.EditOffer And isTemplate) Then
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
      If Not isTemplate Then
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
      If (Not isTemplate AndAlso modMessage = "") Then
        MyCommon.QueryStr = "select IncentiveID from CPE_Incentives with (NoLock) where CreatedDate=LastUpdate and IncentiveID=" & OfferID
        rst3 = MyCommon.LRT_Select
        If (rst3.Rows.Count = 0) Then
          Send_Status(OfferID, 2)
        End If
      End If
    %>
    <div id="column">
      <div class="box" id="eligibility">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.eligibilitynotifications", LanguageID))%>
          </span>
        </h2>
        <table class="list" id="tblNotify" summary="<% Sendb(Copient.PhraseLib.Lookup("term.eligibilitynotifications", LanguageID)) %>">
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
              ' Printed message notifications
              t = 1
              MyCommon.QueryStr = "select PM.MessageID, PM.MessageTypeID, D.DeliverableID, D.DisallowEdit " & _
                                  "from CPE_Deliverables D with (NoLock) " & _
                                  "inner join PrintedMessages PM with (NoLock) on D.OutputID=PM.MessageID " & _
                                  "where D.RewardOptionPhase=1 and D.RewardOptionID=" & roid & " and D.DeliverableTypeID=4;"
              rst = MyCommon.LRT_Select()
              If (rst.Rows.Count > 0) Then
                NotificationCount += rst.Rows.Count
                AddOptionArray.Set(0, False)
                Send("<tr class=""shadeddark"">")
                Send("  <td colspan=""" & 3 & """>")
                Send("    <h3>" & Copient.PhraseLib.Lookup("term.printedmessages", LanguageID) & "</h3>")
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
                If (isTemplate Or FromTemplate) Then
                  Send("  <td></td>")
                End If
                Send("</tr>")
                For Each row In rst.Rows
                  NotificationDisabled = IIf((Logix.UserRoles.EditOffer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False))), "", " disabled=""disabled""")
                  Send("<tr class=""shaded"">")
                  Sendb("  <td><input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & NotificationDisabled & " value=""X"" ")
                  Send("onClick=""if(confirm('" & Copient.PhraseLib.Lookup("notification.confirmdelete", LanguageID) & "')) LoadDocument('web-offer-not.aspx?mode=DeletePrintedMsg&Phase=1&OfferID=" & OfferID & "&MessageID=" & MyCommon.NZ(row.Item("MessageID"), 0) & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "')"" /></td>")
                  Send("  <td><a href=""javascript:openPopup('CPEoffer-rew-pmsg.aspx?OfferID=" & OfferID & "&Phase=1&RewardID=" & roid & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "&MessageID=" & MyCommon.NZ(row.Item("MessageID"), 0) & "')"">" & Copient.PhraseLib.Lookup("term.printedmessage", LanguageID) & "</a></td>")
                  Send("  <td>" & GetMessageTypeName(MyCommon.NZ(row.Item("MessageTypeID"), 0)) & "</td>")
                  ' Find the per-tier values and build up the details string:
                  MyCommon.QueryStr = "select PM.MessageID, PMT.TierLevel, PMT.BodyText " & _
                                      "from PrintedMessages as PM with (NoLock) " & _
                                      "left join PrintedMessageTiers as PMT with (NoLock) on PM.MessageID=PMT.MessageID " & _
                                      "where PM.MessageID=" & MyCommon.NZ(row.Item("MessageID"), 0) & ";"
                  rst2 = MyCommon.LRT_Select
                  If rst2.Rows.Count = 0 Then
                    Send("  <td colspan=""" & TierLevels & """>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                  Else
                    While t <= TierLevels
                      If t > rst2.Rows.Count Then
                        Send("  <td>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                      Else
                        Details = New StringBuilder(200)
                        Details.Append(ReplaceTags(MyCommon.NZ(rst2.Rows(t - 1).Item("BodyText"), "")))
                        If (Details.ToString().Length > 80) Then
                          Details = Details.Remove(77, (Details.Length - 77))
                          Details.Append("...")
                        End If
                        Details.Replace(vbCrLf, "<br />")
                        Send("  <td>""" & MyCommon.SplitNonSpacedString(Details.ToString, 25) & """</td>")
                      End If
                      t += 1
                    End While
                  End If
                  t = 1
                  If (isTemplate) Then
                    Send("  <td class=""templine"">")
                    Send("    <input type=""checkbox"" id=""chkLocked1"" name=""chkLocked"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, " checked=""checked""", "") & " onclick=""javascript:updateLocked('lockPmsg" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "', this.checked);"" />")
                    Send("    <input type=""hidden"" id=""rewPmsg" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""rew"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ />")
                    Send("    <input type=""hidden"" id=""lockPmsg" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""locked"" value=""" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "1", "0") & """ />")
                    Send("  </td>")
                  ElseIf (FromTemplate) Then
                    Send("  <td class=""templine"">" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "Yes", "No") & "</td>")
                  End If
                  Send("</tr>")
                Next
              End If
              
              ' Graphics notifications
              t = 1
              MyCommon.QueryStr = "select OSA.Name as GraphicName, OSA.OnScreenAdID as AdID, D.DeliverableID, D.ScreenCellID as CellID, D.DisallowEdit, " & _
                                  "OSA.Width, OSA.Height, OSA.ImageType, SC.Name as CellName, SL.Name as LayoutName " & _
                                  "from OnScreenAds as OSA with (NoLock) " & _
                                  "inner join CPE_Deliverables as D with (NoLock) on OSA.OnScreenAdID=D.OutputID " & _
                                  "inner join ScreenCells as SC with (NoLock) on D.ScreenCellID=SC.CellID " & _
                                  "inner join ScreenLayouts as SL with (NoLock) on SL.LayoutID=SC.LayoutID " & _
                                  "where D.RewardOptionID=" & roid & " and OSA.Deleted=0 and D.DeliverableTypeID=1 and D.RewardOptionPhase=1;"
              rst = MyCommon.LRT_Select
              If (rst.Rows.Count > 0) Then
                NotificationCount += rst.Rows.Count
                Send("<tr class=""shadeddark"">")
                Send("  <td colspan=""" & 3 & """>")
                Send("    <h3>" & Copient.PhraseLib.Lookup("term.graphics", LanguageID) & "</h3>")
                Send("  </td>")
                Send("  <td colspan=""" & TierLevels & """>")
                Send("  </td>")
                If (isTemplate Or FromTemplate) Then
                  Send("  <td></td>")
                End If
                Send("</tr>")
                For Each row In rst.Rows
                  RewardDisabled = IIf((Logix.UserRoles.EditOffer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False))), "", " disabled=""disabled""")
                  DeleteGraphicURL = "web-offer-not.aspx?mode=DeleteGraphic&OfferID=" & OfferID & "&deliverableid=" & MyCommon.NZ(row.Item("DeliverableID"), "")
                  Send("<tr class=""shaded"">")
                  Send("  <td><input type=""button"" class=""ex"" name=""ex"" " & RewardDisabled & " title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ onClick=""if(confirm('" & Copient.PhraseLib.Lookup("notification.confirmdelete", LanguageID) & "')) LoadDocument('" & DeleteGraphicURL & " ');"" value=""X"" /></td>")
                  Send("  <td><a href=""javascript:openPopup('CPEoffer-rew-graphic.aspx?OfferID=" & OfferID & "&ad=" & MyCommon.NZ(row.Item("AdId"), "") & "&cellselect=" & MyCommon.NZ(row.Item("CellID"), "") & "&imagetype=" & MyCommon.NZ(row.Item("ImageType"), "") & "&preview=1&Phase=1')"">" & Copient.PhraseLib.Lookup("term.graphic", LanguageID) & "</a></td>")
                  Send("  <td></td>")
                  Sendb("  <td colspan=""" & TierLevels & """><a href=""graphic-edit.aspx?OnScreenAdId=" & MyCommon.NZ(row.Item("AdId"), "") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("GraphicName"), ""), 25) & "</a>")
                  Sendb("&nbsp;(" & MyCommon.NZ(row.Item("Width"), "") & " x " & MyCommon.NZ(row.Item("Height"), ""))
                  If MyCommon.NZ(row.Item("ImageType"), "") = 1 Then
                    Sendb("&nbsp;" & Copient.PhraseLib.Lookup("term.jpeg", LanguageID))
                  ElseIf MyCommon.NZ(row.Item("ImageType"), "") = 2 Then
                    Sendb("&nbsp;" & Copient.PhraseLib.Lookup("term.gif", LanguageID))
                  End If
                  Send(")</td>")
                  If (isTemplate) Then
                    Send("  <td class=""templine"">")
                    Send("    <input type=""checkbox"" id=""chkLocked3"" name=""chkLocked"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, " checked=""checked""", "") & " onclick=""javascript:updateLocked('lockGraphics" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "', this.checked);"" />")
                    Send("    <input type=""hidden"" id=""rewGrapics" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""rew"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ />")
                    Send("    <input type=""hidden"" id=""lockGraphics" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""locked"" value=""" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "1", "0") & """ />")
                    Send("  </td>")
                  ElseIf (FromTemplate) Then
                    Send("  <td class=""templine"">" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "Yes", "No") & "</td>")
                  End If
                  Send("</tr>")
                Next
              End If
            %>
          </tbody>
        </table>
        <hr class="hidden" />
      </div>
      <div class="box" id="accumulation">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.accumnotifications", LanguageID))%>
          </span>
        </h2>
        <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.accumnotifications", LanguageID)) %>">
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
              <th align="left" scope="col" class="th-details">
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
            <tr class="shadeddark">
              <td colspan="<%Sendb(3)%>">
                <h3>
                  <% Sendb(Copient.PhraseLib.Lookup("term.printedmessages", LanguageID))%>
                </h3>
              </td>
            </tr>
            <%
              ' Accumulation printed message notifications
              MyCommon.QueryStr = "select PM.MessageID, PM.MessageTypeID, PMT.BodyText, D.DeliverableID, D.DisallowEdit " & _
                                  "from CPE_Deliverables D with (NoLock) " & _
                                  "inner join PrintedMessages PM with (NoLock) on D.OutputID=PM.MessageID " & _
                                  "inner join PrintedMessageTiers PMT with (NoLock) on PM.MessageID=PMT.MessageID " & _
                                  "where D.RewardOptionPhase=2 and D.RewardOptionID=" & roid & " and D.DeliverableTypeID=4 and PMT.TierLevel=0;"
              rst = MyCommon.LRT_Select()
              If (rst.Rows.Count > 0) Then
                AddOptionArray.Set(1, False)
                AccumCount += rst.Rows.Count
                For Each row In rst.Rows
                  RewardDisabled = IIf((Logix.UserRoles.EditOffer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False))), "", " disabled=""disabled""")
                  Send("<tr class=""shaded"">")
                  Sendb(" <td><input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & RewardDisabled & " value=""X"" ")
                  Send("onClick=""if(confirm('" & Copient.PhraseLib.Lookup("accumulation.confirmdelete", LanguageID) & "')) LoadDocument('web-offer-not.aspx?mode=DeletePrintedMsg&Phase=2&OfferID=" & OfferID & "&MessageID=" & MyCommon.NZ(row.Item("MessageID"), 0) & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "')"" /></td>")
                  Send(" <td><a class=""hidden"" href=""CPEoffer-rew-pmsg.aspx?OfferID=" & OfferID & "&Phase=2&RewardID=" & roid & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "&MessageID=" & MyCommon.NZ(row.Item("MessageID"), 0) & """>►</a>")
                  Send(" <a href=""javascript:openPopup('CPEoffer-rew-pmsg.aspx?OfferID=" & OfferID & "&Phase=2&RewardID=" & roid & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "&MessageID=" & MyCommon.NZ(row.Item("MessageID"), 0) & "')"">" & Copient.PhraseLib.Lookup("term.printedmessage", LanguageID) & "</a></td>")
                  Send(" <td>" & GetMessageTypeName(MyCommon.NZ(row.Item("MessageTypeID"), 0)) & "</td>")
                  Details = New StringBuilder(200)
                  Details.Append(ReplaceTags(MyCommon.NZ(row.Item("BodyText"), "")))
                  If (Details.ToString().Length > 80) Then
                    Details = Details.Remove(77, (Details.Length - 77))
                    Details.Append("...")
                  End If
                  Details.Replace(vbCrLf, "<br />")
                  Send("<td>""" & Details.ToString & """</td>")
                  If (isTemplate) Then
                    Send("<td class=""templine"">")
                    Send("  <input type=""checkbox"" id=""chkLocked4"" name=""chkLocked"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, " checked=""checked""", "") & " onclick=""javascript:updateLocked('lockAccum" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "', this.checked);"" />")
                    Send("  <input type=""hidden"" id=""rewAccum" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""rew"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ />")
                    Send("  <input type=""hidden"" id=""lockAccum" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""locked"" value=""" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "1", "0") & """ />")
                    Send("</td>")
                  ElseIf (FromTemplate) Then
                    Send("<td class=""templine"">" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "Yes", "No") & "</td>")
                  End If
                  Send("</tr>")
                Next
              Else
                Send("<tr><td colspan=""" & 3 & """ class=""red"">*" & Copient.PhraseLib.Lookup("offer-accum.pmsgnotavailable", LanguageID) & "</td></tr>")
              End If
            %>
          </tbody>
        </table>
        <hr class="hidden" />
      </div>
      <div class="box" id="neweligibility">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("CPE-offer-not.addeligibility", LanguageID))%>
          </span>
        </h2>
        <%
          If isTemplate Then
            Send("<span class=""temp"">")
            Send("  <input type=""checkbox"" class=""tempcheck"" id=""Disallow_Notifications"" name=""Disallow_Notifications""" & IIf(Disallow_Notifications, " checked=""checked""", "") & " />")
            Send("  <label for=""Disallow_Notifications"">" & Copient.PhraseLib.Lookup("term.locked", LanguageID) & "</label>")
            Send("</span>")
          End If
          MyCommon.QueryStr = "SELECT PECT.EngineID, PECT.EngineSubTypeID, PECT.ComponentTypeID, DT.DeliverableTypeID, DT.Description, DT.PhraseID, PECT.Singular, " & _
                              "  CASE DeliverableTypeID " & _
                              "    WHEN 1 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & roid & " and Deleted=0 and RewardOptionPhase=1 and DeliverableTypeID=1) " & _
                              "    WHEN 2 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & roid & " and Deleted=0 and RewardOptionPhase=1 and DeliverableTypeID=2) " & _
                              "    WHEN 4 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & roid & " and Deleted=0 and RewardOptionPhase=1 and DeliverableTypeID=4) " & _
                              "    WHEN 5 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & roid & " and Deleted=0 and RewardOptionPhase=1 and DeliverableTypeID=5) " & _
                              "    WHEN 6 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & roid & " and Deleted=0 and RewardOptionPhase=1 and DeliverableTypeID=6) " & _
                              "    WHEN 7 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & roid & " and Deleted=0 and RewardOptionPhase=1 and DeliverableTypeID=7) " & _
                              "    WHEN 8 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & roid & " and Deleted=0 and RewardOptionPhase=1 and DeliverableTypeID=8) " & _
                              "    WHEN 9 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & roid & " and Deleted=0 and RewardOptionPhase=1 and DeliverableTypeID=9) " & _
                              "    WHEN 10 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & roid & " and Deleted=0 and RewardOptionPhase=1 and DeliverableTypeID=10) " & _
                              "    WHEN 11 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & roid & " and Deleted=0 and RewardOptionPhase=1 and DeliverableTypeID=11) " & _
                              "    WHEN 12 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & roid & " and Deleted=0 and RewardOptionPhase=1 and DeliverableTypeID=12) " & _
                              "    ELSE 0 " & _
                              "  END as ItemCount " & _
                              "FROM PromoEngineComponentTypes AS PECT " & _
                              "INNER JOIN CPE_DeliverableTypes AS DT ON DT.DeliverableTypeID=PECT.LinkID " & _
                              "WHERE EngineID=3 AND EngineSubTypeID=" & EngineSubTypeID & " AND PECT.ComponentTypeID=3 AND Enabled=1"
          'Impose a few special limits on the query based on various factors:
          If (Not IsCustomerAssigned) Then
            'If no customer condition exists, then no notifications are allowed
            MyCommon.QueryStr &= " AND DeliverableTypeID=0"
          End If
          MyCommon.QueryStr &= " ORDER BY DisplayOrder;"
          rst = MyCommon.LRT_Select
          If rst.Rows.Count > 0 Then
            Send("<select id=""neweliglobal"" name=""neweliglobal""" & DeleteBtnDisabled & ">")
            For Each row In rst.Rows
              If (row.Item("Singular") = False) OrElse (row.Item("Singular") = True AndAlso row.Item("ItemCount") = 0) Then
                Send("<option value=""" & row.Item("DeliverableTypeID") & """>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "</option>")
              End If
            Next
            Send("</select>")
            Send("<input type=""submit"" class=""regular"" id=""addglobal"" name=""addglobal"" value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """" & DeleteBtnDisabled & " />")
            Send("<br />")
          Else
            'Nothing's available.  Tell the user why.
            Send("<br class=""half"" />")
            If (Not IsCustomerAssigned) Then
              Send(Copient.PhraseLib.Detokenize("web-offer-not.NotAvailable", LanguageID, OfferID))
            Else
              Send(Copient.PhraseLib.Lookup("ueoffer-not.NoneAvailable", LanguageID))
            End If
            Send("<br class=""half"" />")
          End If
        %>
      </div>
    </div>
    <br clear="all" />
  </div>
</form>
<!-- #Include virtual="/include/graphic-reward.inc" -->
<%
  If (NotificationCount = 0 Or AccumCount = 0) Then
    Send("<script type=""text/javascript"">")
    If (NotificationCount = 0 AndAlso isTemplate) Then
      Send("document.getElementById(""tblNotify"").style.display = 'none';")
    ElseIf (NotificationCount = 0) Then
      Send("document.getElementById(""eligibility"").style.display = 'none';")
    End If
    If (AccumCount = 0) Then
      Send("document.getElementById(""accumulation"").style.display = 'none';")
    End If
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
