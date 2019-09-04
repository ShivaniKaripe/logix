<%@ Page Language="vb" Debug="true" CodeFile="/logix/LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Text.RegularExpressions" %>
<%
  ' *****************************************************************************
  ' * FILENAME: CAM-offer-rew.aspx 
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
  Dim MyCam As New Copient.CAM
  Dim rst As DataTable
  Dim row As DataRow
  Dim rst2 As DataTable
  Dim row2 As DataRow
  Dim rst3 As DataTable
  Dim rstPT As DataTable
  Dim OfferID As Long
  Dim RewardID As Long
  Dim DeliverableID As Long
  Dim PKID As Long
  Dim MessageID As Long = 0
  Dim FrankID As Long = 0
  Dim DiscountID As Long = 0
  Dim ParentROID As Long = 0
  Dim ROID As Long = 0
  Dim Name As String = ""
  Dim DeleteGraphicURL As String = ""
  Dim AddTouchPtURL As String = ""
  Dim UrlTokens As String = ""
  Dim DeliverableType As Integer
  Dim AddOptionArray As New BitArray(8, True)
  Dim MessageTypeLabel As String = ""
  Dim index As Integer = 0
  Dim i As Integer = 0
  Dim DeleteBtnDisabled As String = ""
  Dim Details As StringBuilder
  Dim IsTemplate As Boolean = False
  Dim IsTemplateVal As String = "Not"
  Dim ActiveSubTab As Integer = 205
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
  Dim IsFooterOffer As Boolean = False
  Dim DiscountWorthy As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "CAM-offer-rew.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  
  OfferID = Request.QueryString("OfferID")
  RewardID = Request.QueryString("RewardID")
  DeliverableID = MyCommon.Extract_Val(Request.QueryString("DeliverableID"))
  PKID = MyCommon.Extract_Val(Request.QueryString("PKID"))
  MessageID = MyCommon.Extract_Val(Request.QueryString("MessageID"))
  FrankID = MyCommon.Extract_Val(Request.QueryString("FrankID"))
  DiscountID = MyCommon.Extract_Val(Request.QueryString("DiscountID"))
  DeliverableType = MyCommon.Extract_Val(Request.QueryString("action"))
  
  If (OfferID = 0) Then
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "/logix/CAM/CAM-offer-gen.aspx")
  End If
  
  MyCommon.QueryStr = "select IncentiveName, IsTemplate, FromTemplate from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID
  rst = MyCommon.LRT_Select
  If (rst.Rows.Count > 0) Then
    Name = MyCommon.NZ(rst.Rows(0).Item("IncentiveName"), "")
    IsTemplate = MyCommon.NZ(rst.Rows(0).Item("IsTemplate"), False)
    FromTemplate = MyCommon.NZ(rst.Rows(0).Item("FromTemplate"), False)
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
  
  IsFooterOffer = MyCam.IsFooterOffer(OfferID)
  If IsFooterOffer Then
    AddOptionArray = New BitArray(8, False)
    AddOptionArray.Set(1, True)
  End If
  
  MyCommon.QueryStr = "select CG.CustomerGroupID, Name, ExcludedUsers from CPE_IncentiveCustomerGroups as ICG with (NoLock) " & _
                      "left join CustomerGroups as CG with (NoLock) on CG.CustomerGroupID=ICG.CustomerGroupID " & _
                      "where RewardOptionID=" & RewardID & ";"
  rst = MyCommon.LRT_Select
  If (rst.Rows.Count > 0) Then
    IsCustomerAssigned = True
  End If
  
  If (Request.QueryString("mode") = "DeleteGraphic") Then
    RemoveGraphic(OfferID, DeliverableID)
  ElseIf (Request.QueryString("mode") = "DeleteCashierMsg") Then
    If (DeliverableID > 0 AndAlso MessageID > 0) Then
      MyCommon.QueryStr = "delete from CPE_CashierMessageTiers with (RowLock) where MessageID=" & MessageID & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "delete from CPE_CashierMessages with (RowLock) where MessageID=" & MessageID & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "delete from CPE_Deliverables with (RowLock) where DeliverableID=" & DeliverableID & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
      MyCommon.LRT_Execute()
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Reward.deletecmsg", LanguageID))
    End If
  ElseIf (Request.QueryString("mode") = "DeleteMembership") Then
    If (DeliverableID > 0 AndAlso DeliverableType > 0 AndAlso OfferID > 0) Then
      MyCommon.QueryStr = "delete from CPE_DeliverableCustomerGroupTiers with (RowLock) where DeliverableID=" & DeliverableID & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "delete from CPE_Deliverables with (RowLock) where RewardOptionID=" & RewardID & " and RewardOptionPhase=3 and DeliverableTypeID=" & DeliverableType & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
      MyCommon.LRT_Execute()
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Reward.deletemembership", LanguageID))
    End If
  ElseIf (Request.QueryString("mode") = "DeletePoints") Then
    If (DeliverableID > 0 AndAlso OfferID > 0) Then
      MyCommon.QueryStr = "delete from CPE_DeliverablePointTiers where DPPKID in " & _
                          "(select PKID from CPE_DeliverablePoints where DeliverableID=" & DeliverableID & ");"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "delete from CPE_DeliverablePoints where DeliverableID=" & DeliverableID & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "delete from CPE_Deliverables with (RowLock) where DeliverableID=" & DeliverableID & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
      MyCommon.LRT_Execute()
      'Check if accumulation message needs to be removed
      MyCommon.QueryStr = "dbo.pa_CPE_AccumMsgEligible"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@ROID", SqlDbType.Int, 4).Value = RewardID
      MyCommon.LRTsp.Parameters.Add("@AccumEligible", SqlDbType.Bit, 1).Direction = ParameterDirection.Output
      MyCommon.LRTsp.ExecuteNonQuery()
      AccumEligible = MyCommon.LRTsp.Parameters("@AccumEligible").Value
      MyCommon.Close_LRTsp()
      If Not (AccumEligible) Then
        'Delete any accumulation messages
        MyCommon.QueryStr = "delete from CPE_Deliverables with (RowLock) where DeliverableID in " & _
                            "(select D.DeliverableID from CPE_RewardOptions as RO " & _
                            " inner join CPE_Deliverables as D on RO.RewardOptionID=D.RewardOptionID " & _
                            " where RO.Deleted=0 and RO.IncentiveID=" & OfferID & " and RewardOptionPhase=2 and DeliverableTypeID=4);"
        MyCommon.LRT_Execute()
      End If
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Reward.deletepoints", LanguageID))
    End If
  ElseIf (Request.QueryString("mode") = "DeleteStoredValue") Then
    If (DeliverableID > 0 AndAlso OfferID > 0) Then
      MyCommon.QueryStr = "delete from CPE_DeliverableStoredValueTiers where DSVPKID in " & _
                          "(select PKID from CPE_DeliverableStoredValue where DeliverableID=" & DeliverableID & ");"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "delete from CPE_DeliverableStoredValue where DeliverableID=" & DeliverableID & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "delete from CPE_Deliverables with (RowLock) where DeliverableID=" & DeliverableID & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
      MyCommon.LRT_Execute()
    End If
    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Reward.deletesv", LanguageID))
  ElseIf (Request.QueryString("mode") = "DeletePrintedMsg") Then
    If (DeliverableID > 0 AndAlso OfferID > 0) Then
      MyCommon.QueryStr = "delete from PrintedMessageTiers with (RowLock) where MessageID=" & MessageID & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "delete from PrintedMessages with (RowLock) where MessageID=" & MessageID & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "delete from CPE_Deliverables with (RowLock) where DeliverableID=" & DeliverableID & " and RewardOptionPhase=3 and DeliverableTypeID=4;"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
      MyCommon.LRT_Execute()
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Reward.deletepmsg", LanguageID))
    End If
  ElseIf (Request.QueryString("mode") = "DeleteDiscount") Then
    If (DeliverableID > 0 AndAlso OfferID > 0) Then
      MyCommon.QueryStr = "delete from CPE_DiscountTiers with (RowLock) where DiscountID=" & DiscountID & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "delete from CPE_Discounts with (RowLock) where DiscountID=" & DiscountID & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "delete from CPE_Deliverables with (RowLock) where DeliverableID=" & DeliverableID & " and RewardOptionPhase=3 and DeliverableTypeID=2;"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
      MyCommon.LRT_Execute()
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Reward.deletediscount", LanguageID))
    End If
  ElseIf (Request.QueryString("mode") = "DeleteFrankingMsg") Then
    If (DeliverableID > 0 AndAlso FrankID > 0) Then
      MyCommon.QueryStr = "delete from CPE_FrankingMessageTiers with (RowLock) where FrankID=" & FrankID & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "delete from CPE_FrankingMessages with (RowLock) where FrankID=" & FrankID & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "delete from CPE_Deliverables with (RowLock) where DeliverableID=" & DeliverableID & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
      MyCommon.LRT_Execute()
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Reward.deletefmsg", LanguageID))
    End If
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
    ActiveSubTab = 206
    IsTemplateVal = "IsTemplate"
  Else
    ActiveSubTab = 205
    IsTemplateVal = "Not"
  End If
  
  If Not IsTemplate Then
    DeleteBtnDisabled = IIf(Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Rewards), "", " disabled=""disabled""")
  Else
    DeleteBtnDisabled = IIf(Logix.UserRoles.EditTemplates, "", " disabled=""disabled""")
  End If
  
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

    if (rewType >= 1 && rewType <= 12) {
      qryStr = "?RewardID=<%Sendb(RewardID)%>&OfferID=<%Sendb(OfferID)%>&tp=1&roid=" + roid;
      if (rewType == 1) {
        pageName = "/logix/CPEoffer-rew-graphic.aspx";
      } else if (rewType == 2) {
        pageName = "/logix/CPEoffer-rew-discount.aspx";
      } else if (rewType == 4) {
        pageName = "/logix/CPEoffer-rew-pmsg.aspx";
      } else if (rewType == 5) {
        qryStr += "&action=5"
        pageName = "/logix/CPEoffer-rew-membership.aspx";
      } else if (rewType == 6) {
        /* Remove membership rewards are currently disabled */
      } else if (rewType == 7) {
        /* Silent deliverable not supported */
      } else if (rewType == 8) {
        pageName = "/logix/CPEoffer-rew-point.aspx";
      } else if (rewType == 9) {
        pageName = "/logix/CPEoffer-rew-cmsg.aspx";
      } else if (rewType == 10) {
        /* Franking not supported */
      } else if (rewType == 11) {
        pageName = "/logix/CPEoffer-rew-sv.aspx";
      } else if (rewType == 12) {
        /* Pass-thrus not supported */
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
        Case 1
          Send("<script type=""text/javascript"">openPopup('/logix/CPEoffer-rew-graphic.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "&Phase=3')</script>")
          MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-graphic", LanguageID))
        Case 2
          Send("<script type=""text/javascript"">openPopup('/logix/CPEoffer-rew-discount.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "&Phase=3')</script>")
          MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-discount", LanguageID))
        Case 3
          'Not used
        Case 4
          Send("<script type=""text/javascript"">openPopup('/logix/CPEoffer-rew-pmsg.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "&Phase=3')</script>")
          MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-pmsg", LanguageID))
        Case 5
          Send("<script type=""text/javascript"">openPopup('/logix/CPEoffer-rew-membership.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "&action=5&Phase=3')</script>")
          MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-member", LanguageID))
        Case 6
          'Revoke membership -- not used
        Case 7
          'Silent deliverable -- not used
        Case 8
          Send("<script type=""text/javascript"">openPopup('/logix/CPEoffer-rew-point.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "&Phase=3&New=1')</script>")
          MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-points", LanguageID))
        Case 9
          Send("<script type=""text/javascript"">openPopup('/logix/CPEoffer-rew-cmsg.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "&Phase=3')</script>")
          MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-cmsg", LanguageID))
        Case 10
          Send("<script type=""text/javascript"">openPopup('/logix/CPEoffer-rew-franking.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "&Phase=3')</script>")
          MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-franking", LanguageID))
        Case 11
          Send("<script type=""text/javascript"">openPopup('/logix/CPEoffer-rew-sv.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "&Phase=3')</script>")
          MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-sv", LanguageID))
        Case Else
          MyCommon.QueryStr = "select PassThruRewardID, Name, PhraseID from PassThruRewards with (NoLock) order by PassThruRewardID;"
          rstPT = MyCommon.LRT_Select
          If rstPT.Rows.Count > 0 Then
            Send("<script type=""text/javascript"">openPopup('/logix/CPEoffer-rew-passthru.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "&PassThruRewardID=" & (rewChoice - 12) & "&Phase=3')</script>")
            MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-passthru", LanguageID))
          End If
      End Select
    End If
  End If
%>
<form action="CAM-offer-rew.aspx" id="mainform" name="mainform">
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
        If (Logix.UserRoles.EditTemplates) Then
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
              ' Printed message rewards
              t = 1
              MyCommon.QueryStr = "select PM.MessageID, PM.MessageTypeID, D.DeliverableID, D.DisallowEdit " & _
                                  "from CPE_Deliverables D with (NoLock) " & _
                                  "inner join PrintedMessages PM with (NoLock) on D.OutputID=PM.MessageID " & _
                                  "where D.RewardOptionPhase=3 and D.RewardOptionID=" & ParentROID & " and D.DeliverableTypeID=4;"
              rst = MyCommon.LRT_Select()
              If (rst.Rows.Count > 0) Then
                RewardsCount += rst.Rows.Count
                AddOptionArray.Set(1, False)
                Send("<tr class=""shadeddark"">")
                Send("  <td colspan=""" & 3 & """>")
                Send("    <h3>" & Copient.PhraseLib.Lookup("CPE-rew-pmsg.header", LanguageID) & "</h3>")
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
                  If Not isTemplate Then
                    RewardDisabled = IIf((Logix.UserRoles.EditOffer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False))), "", " disabled=""disabled""")
                  Else
                    RewardDisabled = IIf(Logix.UserRoles.EditTemplates, "", " disabled=""disabled""")
                  End If
                  
                  Send("<tr class=""shaded"">")
                  Sendb("  <td><input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & RewardDisabled & " value=""X"" ")
                  Send("onClick=""if(confirm('" & Copient.PhraseLib.Lookup("reward.confirmdelete", LanguageID) & "')) LoadDocument('/logix/CAM/CAM-offer-rew.aspx?mode=DeletePrintedMsg&OfferID=" & OfferID & "&MessageID=" & MyCommon.NZ(row.Item("MessageID"), 0) & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "')"" /></td>")
                  Send("  <td><a href=""javascript:openPopup('/logix/CPEoffer-rew-pmsg.aspx?OfferID=" & OfferID & "&RewardID=" & RewardID & "&MessageID=" & MyCommon.NZ(row.Item("MessageID"), 0) & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "')"">" & Copient.PhraseLib.Lookup("term.printedmessage", LanguageID) & "</a></td>")
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
                  If (IsTemplate) Then
                    Send("  <td class=""templine"">")
                    Send("    <input type=""checkbox"" id=""chkLocked2"" name=""chkLocked"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, " checked=""checked""", "") & " onclick=""javascript:updateLocked('lockPmsg" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "', this.checked);"" />")
                    Send("    <input type=""hidden"" id=""rewPmsg" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""rew"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ />")
                    Send("    <input type=""hidden"" id=""lockPmsg" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""locked"" value=""" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "1", "0") & """ />")
                    Send("  </td>")
                  ElseIf (FromTemplate) Then
                    Send("  <td class=""templine"">" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "Yes", "No") & "</td>")
                  End If
                  Send("</tr>")
                Next
              End If
              
              ' Cashier message rewards
              t = 1
              MyCommon.QueryStr = "select D.DeliverableID, D.DisallowEdit, CM.MessageID " & _
                                  "from CPE_Deliverables D with (NoLock) " & _
                                  "inner join CPE_CashierMessages CM with (NoLock) on D.OutputID=CM.MessageID " & _
                                  "where D.RewardOptionID=" & ParentROID & " and DeliverableTypeID=9 and D.RewardOptionPhase=3;"
              rst = MyCommon.LRT_Select()
              If (rst.Rows.Count > 0) Then
                RewardsCount += rst.Rows.Count
                AddOptionArray.Set(2, False)
                Send("<tr class=""shadeddark"">")
                Send("  <td colspan=""" & 3 & """>")
                Send("    <h3>" & Copient.PhraseLib.Lookup("CPE-rew-cmsg.header", LanguageID) & "</h3>")
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
                  If Not IsTemplate Then
                    RewardDisabled = IIf((Logix.UserRoles.EditOffer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False))), "", " disabled=""disabled""")
                  Else
                    RewardDisabled = IIf(Logix.UserRoles.EditTemplates, "", " disabled=""disabled""")
                  End If
                  Send("<tr class=""shaded"">")
                  Sendb("  <td><input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & RewardDisabled & " value=""X"" ")
                  Send("onClick=""if(confirm('" & Copient.PhraseLib.Lookup("reward.confirmdelete", LanguageID) & "')) LoadDocument('/logix/CAM/CAM-offer-rew.aspx?mode=DeleteCashierMsg&OfferID=" & OfferID & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "&MessageID=" & MyCommon.NZ(row.Item("MessageID"), 0) & "')"" /></td>")
                  Send("  <td><a href=""javascript:openPopup('/logix/CPEoffer-rew-cmsg.aspx?OfferID=" & OfferID & "&RewardID=" & RewardID & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "&MessageID=" & MyCommon.NZ(row.Item("MessageID"), 0) & "')"">" & Copient.PhraseLib.Lookup("term.cashiermessage", LanguageID) & "</a></td>")
                  Send("  <td></td>")
                  ' Find the per-tier values:
                  MyCommon.QueryStr = "select CM.MessageID, CMT.Line1, CMT.Line2 " & _
                                      "from CPE_CashierMessages as CM with (NoLock) " & _
                                      "left join CPE_CashierMessageTiers as CMT with (NoLock) on CM.MessageID=CMT.MessageID " & _
                                      "where CM.MessageID=" & MyCommon.NZ(row.Item("MessageID"), 0) & ";"
                  rst2 = MyCommon.LRT_Select
                  If rst2.Rows.Count = 0 Then
                    Send("  <td colspan=""" & TierLevels & """>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                  Else
                    While t <= TierLevels
                      If t > rst2.Rows.Count Then
                        Send("  <td>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                      Else
                        Send("  <td>""" & MyCommon.NZ(rst2.Rows(t - 1).Item("Line1"), "") & "<br />" & MyCommon.NZ(rst2.Rows(t - 1).Item("Line2"), "") & """</td>")
                      End If
                      t += 1
                    End While
                  End If
                  t = 1
                  If (IsTemplate) Then
                    Send("  <td class=""templine"">")
                    Send("    <input type=""checkbox"" id=""chkLocked3"" name=""chkLocked"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, " checked=""checked""", "") & " onclick=""javascript:updateLocked('lockCmsg" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "', this.checked);"" />")
                    Send("    <input type=""hidden"" id=""rewCmsg" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""rew"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ />")
                    Send("    <input type=""hidden"" id=""lockCmsg" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""locked"" value=""" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "1", "0") & """ />")
                    Send("  </td>")
                  ElseIf (FromTemplate) Then
                    Send("  <td class=""templine"">" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "Yes", "No") & "</td>")
                  End If
                  Send("</tr>")
                Next
              End If
              
              ' Franking message rewards
              t = 1
              MyCommon.QueryStr = "select D.DeliverableID, D.DisallowEdit, FM.FrankID " & _
                                  "from CPE_Deliverables D with (NoLock) " & _
                                  "inner join CPE_FrankingMessages FM with (NoLock) on D.OutputID=FM.FrankID " & _
                                  "where D.RewardOptionID=" & ParentROID & " and DeliverableTypeID=10 and D.RewardOptionPhase=3;"
              rst = MyCommon.LRT_Select()
              If (rst.Rows.Count > 0) Then
                RewardsCount += rst.Rows.Count
                AddOptionArray.Set(6, False)
                Send("<tr class=""shadeddark"">")
                Send("  <td colspan=""" & 3 & """>")
                Send("    <h3>" & Copient.PhraseLib.Lookup("CPE-rew-fmsg.header", LanguageID) & "</h3>")
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
                  If Not IsTemplate Then
                    RewardDisabled = IIf((Logix.UserRoles.EditOffer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False))), "", " disabled=""disabled""")
                  Else
                    RewardDisabled = IIf(Logix.UserRoles.EditTemplates, "", " disabled=""disabled""")
                  End If
                  Send("<tr class=""shaded"">")
                  Sendb("  <td><input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & RewardDisabled & " value=""X"" ")
                  Send("onClick=""if(confirm('" & Copient.PhraseLib.Lookup("reward.confirmdelete", LanguageID) & "')) LoadDocument('/logix/CAM/CAM-offer-rew.aspx?mode=DeleteFrankingMsg&OfferID=" & OfferID & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "&FrankID=" & MyCommon.NZ(row.Item("FrankID"), 0) & "')"" /></td>")
                  Send("  <td><a href=""javascript:openPopup('/logix/CPEoffer-rew-franking.aspx?OfferID=" & OfferID & "&RewardID=" & RewardID & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "&FrankID=" & MyCommon.NZ(row.Item("FrankID"), 0) & "')"">" & Copient.PhraseLib.Lookup("term.frankingmessage", LanguageID) & "</a></td>")
                  Send("  <td></td>")
                  ' Find the per-tier values:
                  MyCommon.QueryStr = "select FMT.OpenDrawer, FMT.ManagerOverride, FMT.FrankFlag, FMT.FrankingText, " & _
                                      "FMT.Line1, FMT.Line2, FMT.Beep, FMT.BeepDuration " & _
                                      "from CPE_FrankingMessageTiers as FMT with (NoLock) " & _
                                      "where FMT.FrankID=" & MyCommon.NZ(row.Item("FrankID"), 0) & ";"
                  rst2 = MyCommon.LRT_Select
                  If rst2.Rows.Count = 0 Then
                    Send("  <td colspan=""" & TierLevels & """>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                  Else
                    While t <= TierLevels
                      If t > rst2.Rows.Count Then
                        Send("  <td>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                      Else
                        If MyCommon.NZ(rst2.Rows(t - 1).Item("FrankingText"), "") = "" Then
                          Send("  <td>")
                        Else
                          Send("  <td>""" & MyCommon.SplitNonSpacedString(rst2.Rows(t - 1).Item("FrankingText"), 25) & """<br />")
                        End If
                        Sendb(IIf(MyCommon.NZ(rst2.Rows(t - 1).Item("OpenDrawer"), False) = True, Copient.PhraseLib.Lookup("term.opendrawer", LanguageID) & ",&nbsp;", Copient.PhraseLib.Lookup("term.closeddrawer", LanguageID) & ",&nbsp;"))
                        Sendb(IIf(MyCommon.NZ(rst2.Rows(t - 1).Item("ManagerOverride"), False) = True, StrConv(Copient.PhraseLib.Lookup("term.override", LanguageID), VbStrConv.Lowercase) & " " & StrConv(Copient.PhraseLib.Lookup("term.required", LanguageID), VbStrConv.Lowercase) & ",&nbsp;", StrConv(Copient.PhraseLib.Lookup("term.override", LanguageID), VbStrConv.Lowercase) & " " & StrConv(Copient.PhraseLib.Lookup("term.notrequired", LanguageID), VbStrConv.Lowercase) & ",&nbsp;"))
                        If (MyCommon.NZ(rst2.Rows(t - 1).Item("FrankFlag"), 0) = 0) Then
                          Sendb(Copient.PhraseLib.Lookup("term.posdata", LanguageID) & " ")
                          Sendb(StrConv(Copient.PhraseLib.Lookup("term.only", LanguageID), VbStrConv.Lowercase))
                        ElseIf (MyCommon.NZ(rst2.Rows(t - 1).Item("FrankFlag"), 0) = 1) Then
                          Sendb(StrConv(Copient.PhraseLib.Lookup("term.frankingmessage", LanguageID), VbStrConv.Lowercase) & " ")
                          Sendb(StrConv(Copient.PhraseLib.Lookup("term.only", LanguageID), VbStrConv.Lowercase))
                        ElseIf (MyCommon.NZ(rst2.Rows(t - 1).Item("FrankFlag"), 0) = 2) Then
                          Sendb(Copient.PhraseLib.Lookup("term.posdata", LanguageID) & " ")
                          Sendb(StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
                          Sendb(StrConv(Copient.PhraseLib.Lookup("term.frankingmessage", LanguageID), VbStrConv.Lowercase) & " ")
                        End If
                        Send("  </td>")
                      End If
                      t += 1
                    End While
                  End If
                  t = 1
                  If (IsTemplate) Then
                    Send("  <td class=""templine"">")
                    Send("    <input type=""checkbox"" id=""chkLocked4"" name=""chkLocked"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, " checked=""checked""", "") & " onclick=""javascript:updateLocked('lockFmsg" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "', this.checked);"" />")
                    Send("    <input type=""hidden"" id=""rewFmsg" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""rew"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ />")
                    Send("    <input type=""hidden"" id=""lockFmsg" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""locked"" value=""" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "1", "0") & """ />")
                    Send("  </td>")
                  ElseIf (FromTemplate) Then
                    Send("  <td class=""templine"">" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "Yes", "No") & "</td>")
                  End If
                  Send("</tr>")
                Next
              End If
              
              ' Points reward
              t = 1
              MyCommon.QueryStr = "select D.DeliverableID, D.DisallowEdit, DP.Quantity, DP.MaxAdjustment, DP.ScorecardBold, DP.ProgramID, PP.ProgramName " & _
                                  "from CPE_Deliverables as D with (NoLock) " & _
                                  "inner join CPE_DeliverablePoints as DP with (NoLock) on DP.DeliverableID=D.DeliverableID " & _
                                  "inner join PointsPrograms as PP with (NoLock) on PP.ProgramID=DP.ProgramID " & _
                                  "where D.RewardOptionID=" & ParentROID & " order by ProgramName;"
              rst = MyCommon.LRT_Select()
              If (rst.Rows.Count > 0) Then
                RewardsCount += rst.Rows.Count
                AddOptionArray.Set(3, False)
                Send("<tr class=""shadeddark"">")
                Send("  <td colspan=""" & 3 & """>")
                Send("    <h3>" & Copient.PhraseLib.Lookup("term.pointsrewards", LanguageID) & "</h3>")
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
                  If Not isTemplate Then
                    RewardDisabled = IIf((Logix.UserRoles.EditOffer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False))), "", " disabled=""disabled""")
                  Else
                    RewardDisabled = IIf(Logix.UserRoles.EditTemplates, "", " disabled=""disabled""")
                  End If
                  Send("<tr class=""shaded"">")
                  Sendb("  <td><input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & RewardDisabled & " value=""X"" ")
                  Send("onClick=""if(confirm('" & Copient.PhraseLib.Lookup("reward.confirmdelete", LanguageID) & "')) LoadDocument('/logix/CAM/CAM-offer-rew.aspx?mode=DeletePoints&OfferID=" & OfferID & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "&ProgramID=" & MyCommon.NZ(row.Item("ProgramID"), 0) & "')"" /></td>")
                  Send("  <td><a href=""javascript:openPopup('/logix/CPEoffer-rew-point.aspx?OfferID=" & OfferID & "&RewardID=" & RewardID & "&ProgramID=" & MyCommon.NZ(row.Item("ProgramID"), 0) & "&quantity=" & MyCommon.NZ(row.Item("Quantity"), 0) & "&maxadjustment=" & MyCommon.NZ(row.Item("MaxAdjustment"), 0) & IIf(MyCommon.NZ(row.Item("ScorecardBold"), 0) = 0, "", "&ScorecardBold=on") & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "')"">" & Copient.PhraseLib.Lookup("term.points", LanguageID) & "</a></td>")
                  Send("  <td></td>")
                  ' Find the per-tier values:
                  MyCommon.QueryStr = "select DP.PKID, DPT.TierLevel, DPT.Quantity " & _
                                      "from CPE_DeliverablePoints as DP with (NoLock) " & _
                                      "left join CPE_DeliverablePointTiers as DPT with (NoLock) on DP.PKID=DPT.DPPKID " & _
                                      "where DP.DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & ";"
                  rst2 = MyCommon.LRT_Select
                  If rst2.Rows.Count = 0 Then
                    Send("  <td colspan=""" & TierLevels & """>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                  Else
                    While t <= TierLevels
                      If t > rst2.Rows.Count Then
                        Send("  <td>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                      Else
                        Send("  <td>" & MyCommon.NZ(rst2.Rows(t - 1).Item("Quantity"), 0) & " " & "<a href=""/logix/point-edit.aspx?ProgramGroupID=" & MyCommon.NZ(row.Item("ProgramID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("ProgramName"), ""), 25) & "</a></td>")
                      End If
                      t += 1
                    End While
                  End If
                  t = 1
                  If (IsTemplate) Then
                    Send("  <td class=""templine"">")
                    Send("    <input type=""checkbox"" id=""chkLocked5"" name=""chkLocked"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, " checked=""checked""", "") & " onclick=""javascript:updateLocked('lockPts" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "', this.checked);"" />")
                    Send("    <input type=""hidden"" id=""rewPts" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""rew"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ />")
                    Send("    <input type=""hidden"" id=""lockPts" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""locked"" value=""" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "1", "0") & """ />")
                    Send("  </td>")
                  ElseIf (FromTemplate) Then
                    Send("  <td class=""templine"">" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "Yes", "No") & "</td>")
                  End If
                  Send("</tr>")
                Next
              End If
              
              'Stored value rewards
              t = 1
              MyCommon.QueryStr = "select SVP.Name, SVP.SVProgramID, D.DeliverableID as PKID, D.DisallowEdit, DSV.Quantity " & _
                                  "from StoredValuePrograms as SVP with (NoLock) " & _
                                  "inner join CPE_DeliverableStoredValue as DSV with (NoLock) on SVP.SVProgramID=DSV.SVProgramID and SVP.Deleted=0 " & _
                                  "inner join CPE_Deliverables as D with (NoLock) on D.DeliverableID=DSV.DeliverableID and D.RewardOptionPhase=3 " & _
                                  "where D.RewardOptionID=" & ParentROID & " order by SVP.Name;"
              rst = MyCommon.LRT_Select()
              If (rst.Rows.Count > 0) Then
                RewardsCount += rst.Rows.Count
                AddOptionArray.Set(7, False)
                Send("<tr class=""shadeddark"">")
                Send("  <td colspan=""" & 3 & """>")
                Send("    <h3>" & Copient.PhraseLib.Lookup("term.storedvaluerewards", LanguageID) & "</h3>")
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
                  If Not IsTemplate Then
                    RewardDisabled = IIf((Logix.UserRoles.EditOffer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False))), "", " disabled=""disabled""")
                  Else
                    RewardDisabled = IIf(Logix.UserRoles.EditTemplates, "", " disabled=""disabled""")
                  End If
                  Send("<tr class=""shaded"">")
                  Sendb("  <td><input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & RewardDisabled & " value=""X"" ")
                  Send("onClick=""if(confirm('" & Copient.PhraseLib.Lookup("reward.confirmdelete", LanguageID) & "')) LoadDocument('/logix/CAM/CAM-offer-rew.aspx?mode=DeleteStoredValue&OfferID=" & OfferID & "&DeliverableID=" & MyCommon.NZ(row.Item("PKID"), 0) & "&ProgramID=" & MyCommon.NZ(row.Item("SVProgramID"), 0) & "')"" /></td>")
                  Send("  <td><a href=""javascript:openPopup('/logix/CPEoffer-rew-sv.aspx?OfferID=" & OfferID & "&RewardID=" & RewardID & "&ProgramID=" & MyCommon.NZ(row.Item("SVProgramID"), 0) & "&quantity=" & MyCommon.NZ(row.Item("Quantity"), 0) & "&DeliverableID=" & MyCommon.NZ(row.Item("PKID"), 0) & "')"">" & Copient.PhraseLib.Lookup("term.storedvalue", LanguageID) & "</a></td>")
                  Send("  <td></td>")
                  ' Find the per-tier values:
                  MyCommon.QueryStr = "select DSV.PKID, DSVT.TierLevel, DSVT.Quantity " & _
                                      "from CPE_DeliverableStoredValue as DSV with (NoLock) " & _
                                      "left join CPE_DeliverableStoredValueTiers as DSVT with (NoLock) on DSV.PKID=DSVT.DSVPKID " & _
                                      "where DSV.DeliverableID=" & MyCommon.NZ(row.Item("PKID"), 0) & ";"
                  rst2 = MyCommon.LRT_Select
                  If rst2.Rows.Count = 0 Then
                    Send("  <td colspan=""" & TierLevels & """>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                  Else
                    While t <= TierLevels
                      If t > rst2.Rows.Count Then
                        Send("  <td>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                      Else
                        Send("  <td>" & CInt(MyCommon.NZ(rst2.Rows(t - 1).Item("Quantity"), "0")) & " " & Copient.PhraseLib.Lookup("term.awardedinprogram", LanguageID) & " " & "<a href=""/logix/SV-edit.aspx?ProgramGroupID=" & MyCommon.NZ(row.Item("SVProgramID"), 0) & """>" & MyCommon.SplitNonSpacedString(row.Item("Name"), 25) & "</a></td>")
                      End If
                      t += 1
                    End While
                  End If
                  t = 1
                  If (IsTemplate) Then
                    Send("  <td class=""templine"">")
                    Send("    <input type=""checkbox"" id=""chkLocked6"" name=""chkLocked"" value=""" & MyCommon.NZ(row.Item("PKID"), 0) & """" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, " checked=""checked""", "") & " onclick=""javascript:updateLocked('lockSV" & MyCommon.NZ(row.Item("PKID"), 0) & "', this.checked);"" />")
                    Send("    <input type=""hidden"" id=""rewSV" & MyCommon.NZ(row.Item("PKID"), 0) & """ name=""rew"" value=""" & MyCommon.NZ(row.Item("PKID"), 0) & """ />")
                    Send("    <input type=""hidden"" id=""lockSV" & MyCommon.NZ(row.Item("PKID"), 0) & """ name=""locked"" value=""" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "1", "0") & """ />")
                    Send("  </td>")
                  ElseIf (FromTemplate) Then
                    Send("  <td class=""templine"">" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "Yes", "No") & "</td>")
                  End If
                  Send("</tr>")
                Next
              End If
              
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
                  If Not IsTemplate Then
                    RewardDisabled = IIf((Logix.UserRoles.EditOffer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False))), "", " disabled=""disabled""")
                  Else
                    RewardDisabled = IIf(Logix.UserRoles.EditTemplates, "", " disabled=""disabled""")
                  End If
                  UrlTokens = "?OfferID=" & OfferID & "&RewardID=" & MyCommon.NZ(row.Item("ROID"), 0) & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), -1) & "&action=" & MyCommon.NZ(row.Item("DeliverableTypeID"), 0)
                  Send("<tr class=""shaded"">")
                  Send("  <td><input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & RewardDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("reward.confirmdelete", LanguageID) & "')) LoadDocument('/logix/CAM/CAM-offer-rew.aspx" & UrlTokens & "&mode=DeleteMembership')"" value=""X"" /></td>")
                  Send("  <td><a href=""javascript:openPopup('/logix/CPEoffer-rew-membership.aspx" & UrlTokens & "')"">" & Copient.PhraseLib.Lookup("term.membership", LanguageID) & "</a></td>")
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
                        Send("  <td><a href=""/logix/cgroup-edit.aspx?CustomerGroupID=" & MyCommon.NZ(rst2.Rows(t - 1).Item("CustomerGroupID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(rst2.Rows(t - 1).Item("Name"), ""), 25) & "</a></td>")
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
              
              ' Graphics rewards
              t = 1
              MyCommon.QueryStr = "select OSA.Name as GraphicName, OSA.OnScreenAdID as AdID, D.DeliverableID, D.ScreenCellID as CellID, D.DisallowEdit, " & _
                                  "OSA.Width, OSA.Height, OSA.ImageType, SC.Name as CellName, SL.Name as LayoutName " & _
                                  "from OnScreenAds as OSA with (NoLock) " & _
                                  "inner join CPE_Deliverables as D with (NoLock) on OSA.OnScreenAdID=D.OutputID " & _
                                  "inner join ScreenCells as SC with (NoLock) on D.ScreenCellID=SC.CellID " & _
                                  "inner join ScreenLayouts as SL with (NoLock) on SL.LayoutID=SC.LayoutID " & _
                                  "where D.RewardOptionID=" & ParentROID & " and OSA.Deleted=0 and D.DeliverableTypeID=1 and D.RewardOptionPhase=3;"
              rst = MyCommon.LRT_Select
              If (rst.Rows.Count > 0) Then
                RewardsCount += rst.Rows.Count
                Send("<tr class=""shadeddark"">")
                Send("  <td colspan=""" & 3 & """>")
                Send("    <h3>" & Copient.PhraseLib.Lookup("CPE-rew-graphics.header", LanguageID) & "</h3>")
                Send("  </td>")
                Send("  <td colspan=""" & TierLevels & """>")
                Send("  </td>")
                If (IsTemplate Or FromTemplate) Then
                  Send("  <td></td>")
                End If
                Send("</tr>")
                For Each row In rst.Rows
                  If Not IsTemplate Then
                    RewardDisabled = IIf((Logix.UserRoles.EditOffer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False))), "", " disabled=""disabled""")
                  Else
                    RewardDisabled = IIf(Logix.UserRoles.EditTemplates, "", " disabled=""disabled""")
                  End If
                  DeleteGraphicURL = "/logix/CAM/CAM-offer-rew.aspx?mode=DeleteGraphic&OfferID=" & OfferID & "&deliverableid=" & MyCommon.NZ(row.Item("DeliverableID"), "")
                  Send("<tr class=""shaded"">")
                  Send("  <td><input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & RewardDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("reward.confirmdelete", LanguageID) & "')) LoadDocument('" & DeleteGraphicURL & " ');"" value=""X"" /></td>")
                  Send("  <td><a href=""javascript:openPopup('/logix/CPEoffer-rew-graphic.aspx?OfferID=" & OfferID & "&ad=" & MyCommon.NZ(row.Item("AdId"), "") & "&cellselect=" & MyCommon.NZ(row.Item("CellID"), "") & "&imagetype=" & MyCommon.NZ(row.Item("ImageType"), "") & "&DeliverableID=" & MyCommon.Extract_Val(row.Item("DeliverableID")) & "&preview=1')"">" & Copient.PhraseLib.Lookup("term.graphic", LanguageID) & "</a></td>")
                  Send("  <td></td>")
                  Sendb("  <td colspan=""" & TierLevels & """><a href=""/logix/graphic-edit.aspx?OnScreenAdId=" & MyCommon.NZ(row.Item("AdId"), "") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("GraphicName"), ""), 25) & "</a>")
                  Sendb("&nbsp;(" & MyCommon.NZ(row.Item("Width"), "") & " x " & MyCommon.NZ(row.Item("Height"), ""))
                  If MyCommon.NZ(row.Item("ImageType"), "") = 1 Then
                    Sendb("&nbsp;" & Copient.PhraseLib.Lookup("term.jpeg", LanguageID))
                  ElseIf MyCommon.NZ(row.Item("ImageType"), "") = 2 Then
                    Sendb("&nbsp;" & Copient.PhraseLib.Lookup("term.gif", LanguageID))
                  End If
                  Send(")</td>")
                  If (IsTemplate) Then
                    Send("  <td class=""templine"">")
                    Send("    <input type=""checkbox"" id=""chkLocked8"" name=""chkLocked"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, " checked=""checked""", "") & " onclick=""javascript:updateLocked('lockGraphics" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "', this.checked);"" />")
                    Send("    <input type=""hidden"" id=""rewGrapics" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""rew"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ />")
                    Send("    <input type=""hidden"" id=""lockGraphics" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""locked"" value=""" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "1", "0") & """ />")
                    Send("  </td>")
                  ElseIf (FromTemplate) Then
                    Send("  <td class=""templine"">" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "Yes", "No") & "</td>")
                  End If
                  Send("</tr>")
                Next
              End If
              
              ' Touchpoint rewards
              t = 1
              MyCommon.QueryStr = "select RO.Name, RO.RewardOptionID, TA.OnScreenAdID as ParentAdID, D.DisallowEdit " & _
                                  "from CPE_RewardOptions RO with (NoLock) " & _
                                  "inner join CPE_DeliverableROIDs DR with (NoLock) on RO.RewardOptionID=DR.RewardOptionID " & _
                                  "inner join CPE_Deliverables D with (NoLock) on D.DeliverableID=DR.DeliverableID " & _
                                  "inner join TouchAreas TA with (NoLock) on DR.AreaID=TA.AreaID " & _
                                  "where RO.Deleted=0 and DR.Deleted=0 and TA.Deleted=0 and RO.IncentiveID=" & OfferID & " and RO.TouchResponse=1 and D.RewardOptionPhase=3 order by RO.RewardOptionID;"
              rst = MyCommon.LRT_Select
              If (rst.Rows.Count > 0) Then
                RewardsCount += rst.Rows.Count
                Send("<tr class=""shadeddark"">")
                Send("  <td colspan=""" & 3 & """>")
                Send("    <h3>" & Copient.PhraseLib.Lookup("term.touchpointrewards", LanguageID) & "</h3>")
                Send("  </td>")
                Send("  <td colspan=""" & TierLevels & """>")
                Send("  </td>")
                If (IsTemplate Or FromTemplate) Then
                  Send("  <td></td>")
                End If
                Send("</tr>")
                index = 0
                For Each row In rst.Rows
                  ROID = MyCommon.NZ(row.Item("RewardOptionID"), 0)
                  'AddTouchPtURL = "/logix/CPEoffer-rew-deliverables.aspx?OfferID=" & OfferID & "&incentiveid=" & OfferID & "&roid=" & ROID & "&phase=3"
                  Send("<tr class=""shadedmid"">")
                  Send("  <td></td>")
                  Send("  <td colspan=""2"">")
                  Send("    <a href=""/logix/graphic-edit.aspx?OnScreenAdId=" & MyCommon.NZ(row.Item("ParentAdID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)), 25) & "</a>")
                  Send("  </td>")
                  Send("  <td colspan=""" & TierLevels & """>")
                  Send("    <label for=""newrewtouchpt" & index & """>" & Copient.PhraseLib.Lookup("CPE-rew.addtouchpoint", LanguageID) & "</label><br />")
                  Send("    <select name=""newrewtouchpt" & index & """ id=""newrewtouchpt" & index & """>")
                  Send_TPRewardOptions(OfferID, ROID)
                  Send("    </select>")
                  Send("    <input type=""button"" class=""regular"" id=""addTouchpoint"" name=""addTouchpoint"" value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """ " & DeleteBtnDisabled & " onclick=""javascript:openTouchptReward(" & index & ", " & ROID & ");"" />")
                  Send("  </td>")
                  If (IsTemplate Or FromTemplate) Then
                    Send("  <td></td>")
                  End If
                  Send("</tr>")
                  If Not IsTemplate Then
                    SetEditableByUser(Logix.UserRoles.EditOffer)
                  Else
                    SetEditableByUser(Logix.UserRoles.EditTemplates)
                  End If 
                  
                  Send_TouchpointRewards(OfferID, ROID, 3, TierLevels, IsTemplate, FromTemplate)
                  index = index + 1
                Next
              End If
            %>
          </tbody>
        </table>
        <hr class="hidden" />
      </div>
      
      <div class="box" id="newreward">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("offer-rew.addreward", LanguageID))%>
          </span>
        </h2>
        <%
          'First set the DiscountWorthy variable, which determines if the offer is eligible to use discount rewards
          MyCommon.QueryStr = "select RO.IncentiveID from CPE_IncentiveTenderTypes as ITT with (NoLock) " & _
                              "left join CPE_RewardOptions as RO with (NoLock) on RO.RewardOptionID=ITT.RewardOptionID " & _
                              "where IncentiveID=" & OfferID & " and RO.Deleted=0;"
          rst = MyCommon.LRT_Select
          If rst.Rows.Count = 0 Then
            DiscountWorthy = True
          End If
          
          If IsFooterOffer AndAlso Not AddOptionArray.Get(1) Then
            Send(Copient.PhraseLib.Lookup("ueoffer-rew.FooterMessage", LanguageID))
          Else
            If IsTemplate Then
              Send("<span class=""temp"">")
              Send("  <input type=""checkbox"" class=""tempcheck"" id=""Disallow_Rewards"" name=""Disallow_Rewards""" & IIf(Disallow_Rewards, " checked=""checked""", "") & " />")
              Send("  <label for=""Disallow_Rewards"">" & Copient.PhraseLib.Lookup("term.locked", LanguageID) & "</label>")
              Send("</span>")
            End If
            MyCommon.QueryStr = "SELECT PECT.EngineID, PECT.ComponentTypeID, DT.DeliverableTypeID, DT.Description, DT.PhraseID, PECT.Singular, " & _
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
                                "WHERE EngineID=6 AND PECT.ComponentTypeID=2 AND Enabled=1"
            'Impose a few special limits on the query based on various factors:
            If (Not IsCustomerAssigned) Then
              'The offer has no customer condition, so the only available rewards is graphics
              MyCommon.QueryStr &= " AND DeliverableTypeID=1"
            End If
            If (IsFooterOffer) Then
              'Based on previous logic, points can't be included in footer offers
              MyCommon.QueryStr &= " AND DeliverableTypeID<>8"
            End If
            MyCommon.QueryStr &= " ORDER BY DisplayOrder;"
            rst = MyCommon.LRT_Select
            If rst.Rows.Count > 0 Then
              Send("<label for=""newrewglobal"">" & Copient.PhraseLib.Lookup("offer-rew.addglobal", LanguageID) & ":</label><br />")
              Send("<select id=""newrewglobal"" name=""newrewglobal""" & DeleteBtnDisabled & ">")
              For Each row In rst.Rows
                If (row.Item("Singular") = False) OrElse (row.Item("Singular") = True AndAlso row.Item("ItemCount") = 0) Then
                  If row.Item("DeliverableTypeID") = 12 Then
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
