<%@ Page Language="vb" Debug="true" CodeFile="ue-cb.vb" Inherits="UECB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
    ' *****************************************************************************
    ' * FILENAME: UEoffer-not.aspx 
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
    Dim MyCpe As New Copient.CPEOffer
    Dim rst As DataTable
    Dim row As DataRow
    Dim rst2 As DataTable
    Dim row2 As DataRow
    Dim rst3 As DataTable
    Dim OfferID As Long
    Dim RewardID As Long
    Dim Name As String = ""
    Dim Description As String
    Dim ConditionID As Integer
    Dim isTemplate As Boolean = False
    Dim FromTemplate As Boolean = False
    Dim roid As Integer
    Dim i As Integer
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
    Dim AccumEnabled As Boolean = False
    Dim OfferHasSVDiscount As Boolean = False
    Dim DeleteBtnDisabled As String = ""
    Dim Disallow_Notifications As Boolean = False
    Dim ActiveSubTab As Integer = 91
    Dim NotificationCount As Integer = 0
    Dim AccumCount As Integer = 0
    Dim IsCustomerAssigned As Boolean = False
    Dim IsProductAssigned As Boolean = False
    Dim IsRewardPointsAssigned As Boolean = False
    Dim infoMessage As String = ""
    Dim modMessage As String = ""
    Dim Handheld As Boolean = False
    Dim Notifications As String() = Nothing
    Dim LockedStatus As String() = Nothing
    Dim LoopCtr As Integer = 0
    Dim NotificationDisabled As String = ""
    Dim BannersEnabled As Boolean = True
    Dim HasPointsReward As Boolean = False
    Dim HasSVReward As Boolean = False
    Dim StatusCode As Copient.LogixInc.STATUS_FLAGS = Copient.LogixInc.STATUS_FLAGS.STATUS_UNKNOWN
    Dim StatusText As String = ""
    Dim TierLevels As Integer = 1
    Dim t As Integer = 1
    Dim IsFooterOffer As Boolean = False
    Dim EngineID As Integer = 2
    Dim EngineSubTypeID As Integer = 0

    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If

    Response.Expires = 0
    MyCommon.AppName = "UEoffer-not.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)

    BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")

    OfferID = Request.QueryString("OfferID")
    'This would redirect to summary page if offer is awaiting deployment or if offer is awaiting recommendations.
    CheckIfValidOffer(MyCommon, OfferID)
    RewardID = Request.QueryString("RewardID")
    DeliverableID = MyCommon.Extract_Val(Request.QueryString("DeliverableID"))
    PKID = MyCommon.Extract_Val(Request.QueryString("PKID"))
    MessageID = MyCommon.Extract_Val(Request.QueryString("MessageID"))
    Phase = MyCommon.Extract_Val(Request.QueryString("Phase"))
    DeliverableType = MyCommon.Extract_Val(Request.QueryString("action"))

    If (OfferID = 0) Then
        Response.Status = "301 Moved Permanently"
        Response.AddHeader("Location", "UEoffer-gen.aspx")
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

    IsFooterOffer = MyCpe.IsFooterOffer(OfferID)
    If IsFooterOffer Then AddOptionArray = New BitArray(3, False)

    ' Determine if a customer condition is set for the offer
    MyCommon.QueryStr = "select CG.CustomerGroupID,Name,ExcludedUsers from CPE_IncentiveCustomerGroups as ICG with (NoLock) " & _
                        "left join CustomerGroups as CG with (NoLock) on CG.CustomerGroupID=ICG.CustomerGroupID " & _
                        "where RewardOptionID=" & roid & " and ICG.Deleted=0;"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
        IsCustomerAssigned = True
    End If

    ' Determine if a product condition is set for the offer
    MyCommon.QueryStr = "select count(*) NumRecs from CPE_IncentiveProductGroups with (NoLock) where Deleted=0 and RewardOptionID=" & roid & ";"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
        IsProductAssigned = (MyCommon.NZ(rst.Rows(0).Item("NumRecs"), 0) > 0)
    End If

    ' Determine if the offer has a stored value discount -- this in part determines if accumulation notifications are made available.
    MyCommon.QueryStr = "select count(*) NumRecs from CPE_Discounts as DI with (NoLock) " & _
                        "inner join CPE_Deliverables as DE on DE.OutputID=DI.DiscountID " & _
                        "where DI.AmountTypeID=7 and DE.RewardOptionID=" & roid & ";"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
        OfferHasSVDiscount = (MyCommon.NZ(rst.Rows(0).Item("NumRecs"), 0) > 0)
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
                Send("<script type=""text/javascript"">openPopup('UEoffer-rew-graphic.aspx?Phase=1&OfferID=" & OfferID & "&RewardID=" & roid & "')</script>")
            ElseIf (Request.QueryString("neweliglobal") = 2) Then
                Send("<script type=""text/javascript"">openPopup('UEoffer-rew-discount.aspx?Phase=1&OfferID=" & OfferID & "&RewardID=" & roid & "')</script>")
            ElseIf (Request.QueryString("neweliglobal") = 3) Then
                'Not used
            ElseIf (Request.QueryString("neweliglobal") = 4) Then
                Send("<script type=""text/javascript"">openPopup('UEoffer-rew-pmsg.aspx?Phase=1&OfferID=" & OfferID & "&RewardID=" & roid & "')</script>")
            ElseIf (Request.QueryString("neweliglobal") = 5) Then
                'Group membership
            ElseIf (Request.QueryString("neweliglobal") = 6) Then
                'Revoke group membership
            ElseIf (Request.QueryString("neweliglobal") = 7) Then
                'Silent deliverable
            ElseIf (Request.QueryString("neweliglobal") = 8) Then
                Send("<script type=""text/javascript"">openPopup('UEoffer-rew-point.aspx?Phase=1&OfferID=" & OfferID & "&RewardID=" & roid & "&New=1')</script>")
            ElseIf (Request.QueryString("neweliglobal") = 9) Then
                Send("<script type=""text/javascript"">openPopup('UEoffer-rew-cmsg.aspx?Phase=1&OfferID=" & OfferID & "&RewardID=" & roid & "')</script>")
            ElseIf (Request.QueryString("neweliglobal") = 10) Then
                Send("<script type=""text/javascript"">openPopup('UEoffer-rew-franking.aspx?Phase=1&OfferID=" & OfferID & "&RewardID=" & roid & "')</script>")
            ElseIf (Request.QueryString("neweliglobal") = 11) Then
                Send("<script type=""text/javascript"">openPopup('UEoffer-rew-sv.aspx?Phase=1&OfferID=" & OfferID & "&RewardID=" & roid & "')</script>")
            ElseIf (Request.QueryString("neweliglobal") = 99) Then
                Send("<script type=""text/javascript"">openPopup('UEoffer-rew-pmsg.aspx?Phase=2&OfferID=" & OfferID & "&RewardID=" & roid & "')</script>")
            Else
                MyCommon.QueryStr = "select PassThruRewardID, Name, PhraseID from PassThruRewards with (NoLock) order by PassThruRewardID;"
                rst = MyCommon.LRT_Select
                If rst.Rows.Count > 0 Then
                    Send("<script type=""text/javascript"">openPopup('UEoffer-rew-passthru.aspx?Phase=1&OfferID=" & OfferID & "&RewardID=" & roid & "&PassThruRewardID=" & (Request.QueryString("neweliglobal") - 12) & "')</script>")
                End If
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
            ResetOfferApprovalStatus(OfferID)
            MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Notification.deletepmsg", LanguageID))
        End If
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
            ResetOfferApprovalStatus(OfferID)
            MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Notification.deletecashier", LanguageID))
        End If
    End If

    ConditionID = Request.QueryString("ConditionID")
    ' dig the offer info out of the database
    ' no one clicked anything
    MyCommon.QueryStr = "select IncentiveID, ClientOfferID, IncentiveName as Name, CPE.Description, PromoClassID, CRMEngineID, Priority, " & _
                        "StartDate, EndDate, EveryDOW, EligibilityStartDate, EligibilityEndDate, TestingStartDate, TestingEndDate, " & _
                        "P1DistQtyLimit, P1DistTimeType, P1DistPeriod, P2DistQtyLimit, P2DistTimeType, P2DistPeriod, P3DistQtyLimit, P3DistTimeType, P3DistPeriod, " & _
                        "EnableImpressRpt, EnableRedeemRpt, CreatedDate, CPE.LastUpdate, CPE.Deleted, CPEOARptDate, CPEOADeploySuccessDate, CPEOADeployRpt, " & _
                        "CRMRestricted, StatusFlag, OC.Description as CategoryName, IsTemplate, FromTemplate, EngineSubTypeID " & _
                        "from CPE_Incentives as CPE with (NoLock) " & _
                        "left join OfferCategories as OC with (NoLock) on CPE.PromoClassID=OfferCategoryID " & _
                        "where IncentiveID=" & OfferID & ";"
    rst = MyCommon.LRT_Select
    For Each row In rst.Rows
        Name = MyCommon.NZ(row.Item("Name"), "")
        isTemplate = MyCommon.NZ(row.Item("IsTemplate"), False)
        FromTemplate = MyCommon.NZ(row.Item("FromTemplate"), False)
        EngineSubTypeID = MyCommon.NZ(row.Item("EngineSubTypeID"), 0)
    Next

    ' select out to determine if we should show the option for a reward printed message which is phase2
    MyCommon.QueryStr = "select PG.ProductGroupID, PG.Name, PT.Phrase as UnitDescription, ExcludedProducts, ProductComboID, " & _
                        "QtyForIncentive, QtyUnitType, AccumMin, AccumLimit, AccumPeriod from CPE_IncentiveProductGroups as IPG with (NoLock) " & _
                        "left join ProductGroups as PG with (NoLock) on PG.ProductGroupID=IPG.ProductGroupID " & _
                        "left join CPE_RewardOptions as RO with (NoLock) on IPG.RewardOptionID=RO.RewardOptionID " & _
                        "left join CPE_UnitTypes as UT with (NoLock) on UT.UnitTypeID=IPG.QtyUnitType " & _
                        "inner join PhraseText PT with (NoLock) on PT.PhraseID=UT.PhraseID " & _
                        "where IPG.RewardOptionID=" & roid & " and IPG.Deleted=0;"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0 AndAlso MyCommon.NZ(rst.Rows(0).Item("AccumMin"), 0) > 0) Then
        AccumEnabled = True
    End If

    ' if it's not already enabled, then check for a customer and points reward being set.
    ' this causes the accumulation to be enabled as well.
    If (Not (AccumEnabled) AndAlso IsCustomerAssigned) Then
        MyCommon.QueryStr = "select count(*) as NumRecs from CPE_Deliverables with (NoLock) where DeliverableTypeID in (8, 11) " & _
                            "and RewardOptionPhase=3 and RewardOptionID=" & roid & " and Deleted=0;"
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0 AndAlso MyCommon.NZ(rst.Rows(0).Item("NumRecs"), 0) > 0) Then
            AccumEnabled = True
            IsRewardPointsAssigned = True
        End If
    End If

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
    Dim m_EditOfferRegardlessOfBuyer = Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID)
    If Not isTemplate Then
        DeleteBtnDisabled = IIf(Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And Disallow_Notifications), "", " disabled=""disabled""")
    Else
        DeleteBtnDisabled = IIf(Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer, "", " disabled=""disabled""")
    End If

    SetDeleteBtnDisabled(DeleteBtnDisabled) 'method found in included file GraphicReward.aspx
    ActiveSubTab = IIf(isTemplate, 25, 24)

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
    If (Logix.UserRoles.AccessInstantWinOffers = False AndAlso EngineID = 2 AndAlso EngineSubTypeID = 1) Then
        Send_Denied(1, "perm.offers-accessinstantwin")
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

            if (rewType >= 1) {
                qryStr = "?RewardID=<%Sendb(RewardID)%>&OfferID=<%Sendb(OfferID)%>&Phase=1&tp=1&roid=" + roid;
                if (rewType == 1) {
                    pageName = "UEoffer-rew-graphic.aspx";
                } else if (rewType == 2) {
                    pageName = "UEoffer-rew-discount.aspx";
                } else if (rewType == 4) {
                    pageName = "UEoffer-rew-pmsg.aspx";
                } else if (rewType == 5) {
                    qryStr += "&action=5"
                    pageName = "UEoffer-rew-membership.aspx";
                } else if (rewType == 8) {
                    pageName = "UEoffer-rew-point.aspx";
                } else if (rewType == 9) {
                    pageName = "UEoffer-rew-cmsg.aspx";
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
<form action="UEoffer-not.aspx" id="mainform" name="mainform">
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
            m_EditOfferRegardlessOfBuyer = Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID)
            If (Logix.UserRoles.EditTemplates And isTemplate And m_EditOfferRegardlessOfBuyer) Then
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
                                            "where D.Deleted=0 and D.RewardOptionPhase=1 and D.RewardOptionID=" & roid & " and D.DeliverableTypeID=4;"
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
                                m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
                                If Not isTemplate Then
                                    NotificationDisabled = IIf((Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False))), "", " disabled=""disabled""")
                                Else
                                    NotificationDisabled = IIf((Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer), "", " disabled=""disabled""")
                                End If

                                Send("<tr class=""shaded"">")
                                Sendb("  <td><input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & NotificationDisabled & " value=""X"" ")
                                Send("onClick=""if(confirm('" & Copient.PhraseLib.Lookup("notification.confirmdelete", LanguageID) & "')) LoadDocument('UEoffer-not.aspx?mode=DeletePrintedMsg&Phase=1&OfferID=" & OfferID & "&MessageID=" & MyCommon.NZ(row.Item("MessageID"), 0) & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "')"" /></td>")
                                Send("  <td><a href=""javascript:openPopup('UEoffer-rew-pmsg.aspx?OfferID=" & OfferID & "&Phase=1&RewardID=" & roid & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "&MessageID=" & MyCommon.NZ(row.Item("MessageID"), 0) & "')"">" & Copient.PhraseLib.Lookup("term.printedmessage", LanguageID) & "</a></td>")
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
                                            Details.Replace("<", "&lt;")
                                            Details.Replace(vbCrLf, vbCrLf & "<br/>")
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
              
                        ' Cashier message notifications
                        t = 1
                        MyCommon.QueryStr = "select D.DeliverableID, D.DisallowEdit, CM.MessageID " & _
                                            "from CPE_Deliverables D with (NoLock) " & _
                                            "inner join CPE_CashierMessages CM with (NoLock) on D.OutputID=CM.MessageID " & _
                                            "where D.RewardOptionID=" & roid & " and DeliverableTypeID=9 and D.RewardOptionPhase=1;"
                        rst = MyCommon.LRT_Select()
                        If (rst.Rows.Count > 0) Then
                            NotificationCount += rst.Rows.Count
                            AddOptionArray.Set(2, False)
                            Send("<tr class=""shadeddark"">")
                            Send("  <td colspan=""" & 3 & """>")
                            Send("    <h3>" & Copient.PhraseLib.Lookup("term.cashiermessages", LanguageID) & "</h3>")
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
                            m_EditOfferRegardlessOfBuyer = Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID)
                            For Each row In rst.Rows
                                If Not isTemplate Then
                                    RewardDisabled = IIf((Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False))), "", " disabled=""disabled""")
                                Else
                                    RewardDisabled = IIf((Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer), "", " disabled=""disabled""")
                                End If
                  
                                Send("<tr class=""shaded"">")
                                Sendb("  <td><input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & RewardDisabled & " value=""X"" ")
                                Send("onClick=""if(confirm('" & Copient.PhraseLib.Lookup("notification.confirmdelete", LanguageID) & "')) LoadDocument('UEoffer-not.aspx?mode=DeleteCashierMsg&OfferID=" & OfferID & "&MessageID=" & MyCommon.NZ(row.Item("MessageID"), 0) & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "')"" /></td>")
                                Send("  <td><a href=""javascript:openPopup('UEoffer-rew-cmsg.aspx?OfferID=" & OfferID & "&Phase=1&RewardID=" & roid & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "')"">" & Copient.PhraseLib.Lookup("term.cashiermessage", LanguageID) & "</a></td>")
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
                                If (isTemplate) Then
                                    Send("  <td class=""templine"">")
                                    Send("    <input type=""checkbox"" id=""chkLocked2"" name=""chkLocked"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, " checked=""checked""", "") & " onclick=""javascript:updateLocked('lockCmsg" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "', this.checked);"" />")
                                    Send("    <input type=""hidden"" id=""rewCmsg" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""rew"" value=""" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ />")
                                    Send("    <input type=""hidden"" id=""lockCmsg" & MyCommon.NZ(row.Item("DeliverableID"), 0) & """ name=""locked"" value=""" & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "1", "0") & """ />")
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
                            m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
                            For Each row In rst.Rows
                                If Not isTemplate Then
                                    RewardDisabled = IIf((Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False))), "", " disabled=""disabled""")
                                Else
                                    RewardDisabled = IIf((Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer), "", " disabled=""disabled""")
                                End If
                                DeleteGraphicURL = "UEoffer-not.aspx?mode=DeleteGraphic&OfferID=" & OfferID & "&deliverableid=" & MyCommon.NZ(row.Item("DeliverableID"), "")
                                Send("<tr class=""shaded"">")
                                Send("  <td><input type=""button"" class=""ex"" name=""ex"" " & RewardDisabled & " title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ onClick=""if(confirm('" & Copient.PhraseLib.Lookup("notification.confirmdelete", LanguageID) & "')) LoadDocument('" & DeleteGraphicURL & " ');"" value=""X"" /></td>")
                                Send("  <td><a href=""javascript:openPopup('UEoffer-rew-graphic.aspx?OfferID=" & OfferID & "&ad=" & MyCommon.NZ(row.Item("AdId"), "") & "&cellselect=" & MyCommon.NZ(row.Item("CellID"), "") & "&imagetype=" & MyCommon.NZ(row.Item("ImageType"), "") & "&preview=1&Phase=1')"">" & Copient.PhraseLib.Lookup("term.graphic", LanguageID) & "</a></td>")
                                Send("  <td></td>")
                                Sendb("  <td colspan=""" & TierLevels & """><a href=""../graphic-edit.aspx?OnScreenAdId=" & MyCommon.NZ(row.Item("AdId"), "") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("GraphicName"), ""), 25) & "</a>")
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
              
                        ' Touchpoint notifications
                        t = 1
                        MyCommon.QueryStr = "select RO.Name, RO.RewardOptionID, TA.OnScreenAdID as ParentAdID, D.DisallowEdit " & _
                                            "from CPE_RewardOptions RO with (NoLock) " & _
                                            "inner join CPE_DeliverableROIDs DR with (NoLock) on RO.RewardOptionID=DR.RewardOptionID " & _
                                            "inner join CPE_Deliverables D with (NoLock) on D.DeliverableID=DR.DeliverableID " & _
                                            "inner join TouchAreas TA with (NoLock) on DR.AreaID=TA.AreaID " & _
                                            "where RO.Deleted=0 and DR.Deleted=0 and TA.Deleted=0 and RO.IncentiveID=" & OfferID & " and RO.TouchResponse=1 and D.RewardOptionPhase=1 order by RO.RewardOptionID;"
                        rst = MyCommon.LRT_Select
                        If (rst.Rows.Count > 0) Then
                            Send("<tr class=""shadeddark"">")
                            Send("  <td colspan=""" & 3 & """>")
                            Send("    <h3>" & Copient.PhraseLib.Lookup("term.touchpointrewards", LanguageID) & "</h3>")
                            Send("  </td>")
                            Send("  <td colspan=""" & TierLevels & """>")
                            Send("  </td>")
                            If (isTemplate Or FromTemplate) Then
                                Send("  <td></td>")
                            End If
                            Send("</tr>")
                            index = 0
                            NotificationCount += rst.Rows.Count
                            For Each row In rst.Rows
                                tpROID = MyCommon.NZ(row.Item("RewardOptionID"), 0)
                                'AddTouchPtURL = "UEoffer-rew-deliverables.aspx?OfferID=" & OfferID & "&incentiveid=" & OfferID & "&roid=" & ROID & "&phase=3"
                                Send("<tr class=""shadedmid"">")
                                Send("  <td></td>")
                                Send("  <td colspan=""2"">")
                                Send("    <a href=""../graphic-edit.aspx?OnScreenAdId=" & MyCommon.NZ(row.Item("ParentAdID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)), 25) & "</a>")
                                Send("  </td>")
                                Send("  <td colspan=""" & TierLevels & """>")
                                Send("    <label for=""newrewtouchpt" & index & """>" & Copient.PhraseLib.Lookup("CPE-rew.addtouchpoint", LanguageID) & "</label><br />")
                                Send("    <select name=""newrewtouchpt" & index & """ id=""newrewtouchpt" & index & """>")
                                Send("      <option value=""1"">" & Copient.PhraseLib.Lookup("term.graphic", LanguageID) & "</option>")
                                Send("      <option value=""4"">" & Copient.PhraseLib.Lookup("term.printedmessage", LanguageID) & "</option>")
                                Send("      <option value=""9"">" & Copient.PhraseLib.Lookup("term.cashiermessage", LanguageID) & "</option>")
                                Send("    </select>")
                                Send("    <input type=""button"" class=""regular"" id=""addTouchpoint"" name=""addTouchpoint"" value=""Add"" " & DeleteBtnDisabled & " onclick=""javascript:openTouchptReward(" & index & ", " & tpROID & ");"" />")
                                Send("  </td>")
                                If (isTemplate Or FromTemplate) Then
                                    Send("<td></td>")
                                End If
                                Send("</tr>")
                                m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
                                If Not isTemplate Then
                                    SetEditableByUser(Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer)
                                Else
                                    SetEditableByUser(Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer)
                                End If
                  
                                Send_TouchpointRewards(OfferID, tpROID, 1, TierLevels, isTemplate, FromTemplate)
                                index = index + 1
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
                        <td colspan="<% Sendb(TierLevels) %>">
                        </td>
                    </tr>
                    <%
                        ' Accumulation printed message notifications
                        MyCommon.QueryStr = "select PM.MessageID, PM.MessageTypeID, PMT.BodyText, D.DeliverableID, D.DisallowEdit " & _
                                            "from CPE_Deliverables D with (NoLock) " & _
                                            "inner join PrintedMessages PM with (NoLock) on D.OutputID=PM.MessageID " & _
                                            "inner join PrintedMessageTiers PMT with (NoLock) on PM.MessageID=PMT.MessageID " & _
                                            "where D.RewardOptionPhase=2 and D.RewardOptionID=" & roid & " and D.DeliverableTypeID=4 and PMT.TierLevel=1;"
                        rst = MyCommon.LRT_Select()
                        If (rst.Rows.Count > 0) Then
                            AddOptionArray.Set(1, False)
                            AccumCount += rst.Rows.Count
                            m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
                            For Each row In rst.Rows
                                If Not isTemplate Then
                                    RewardDisabled = IIf((Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False))), "", " disabled=""disabled""")
                                Else
                                    RewardDisabled = IIf((Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer), "", " disabled=""disabled""")
                                End If
                                Send("<tr class=""shaded"">")
                                Sendb(" <td><input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & RewardDisabled & " value=""X"" ")
                                Send("onClick=""if(confirm('" & Copient.PhraseLib.Lookup("accumulation.confirmdelete", LanguageID) & "')) LoadDocument('UEoffer-not.aspx?mode=DeletePrintedMsg&Phase=2&OfferID=" & OfferID & "&MessageID=" & MyCommon.NZ(row.Item("MessageID"), 0) & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "')"" /></td>")
                                Send(" <td><a class=""hidden"" href=""UEoffer-rew-pmsg.aspx?OfferID=" & OfferID & "&Phase=2&RewardID=" & roid & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "&MessageID=" & MyCommon.NZ(row.Item("MessageID"), 0) & """>►</a>")
                                Send(" <a href=""javascript:openPopup('UEoffer-rew-pmsg.aspx?OfferID=" & OfferID & "&Phase=2&RewardID=" & roid & "&DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & "&MessageID=" & MyCommon.NZ(row.Item("MessageID"), 0) & "')"">" & Copient.PhraseLib.Lookup("term.printedmessage", LanguageID) & "</a></td>")
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
                            If (Not AccumEnabled) Then
                                Send("<tr><td colspan=""" & 3 & """ class=""red"">*" & Copient.PhraseLib.Lookup("offer-accum.pmsgnotavailable", LanguageID) & "</td></tr>")
                            End If
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
                                    "WHERE EngineID=9 AND EngineSubTypeID=" & EngineSubTypeID & " AND PECT.ComponentTypeID=3 AND Enabled=1"
                'Impose a few special limits on the query based on various factors:
                If (Not IsCustomerAssigned OrElse Not IsProductAssigned) OrElse (IsFooterOffer) Then
                    'If no customer and product conditions exist, then no notifications are allowed
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
                    If ((IsCustomerAssigned And IsProductAssigned) AndAlso ((AccumEnabled OrElse OfferHasSVDiscount) AndAlso AddOptionArray.Get(1))) _
                    OrElse ((IsRewardPointsAssigned) AndAlso (AccumEnabled AndAlso AddOptionArray.Get(1))) Then
                        Send("<option value=""99"">" & Copient.PhraseLib.Lookup("term.rewardprintedmessage", LanguageID) & "</option>")
                    End If
                    Send("</select>")
                    Send("<input type=""submit"" class=""regular"" id=""addglobal"" name=""addglobal"" value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """" & DeleteBtnDisabled & " />")
                    Send("<br />")
                Else
                    'Nothing's available.  Tell the user why.
                    Send("<br class=""half"" />")
                    If IsFooterOffer Then
                        Send(Copient.PhraseLib.Lookup("ueoffer-not.NotAvailableFooter", LanguageID) & "<br />")
                    ElseIf (Not IsCustomerAssigned OrElse Not IsProductAssigned) Then
                        Send(Copient.PhraseLib.Detokenize("ueoffer-not.ProdCustConRequired", LanguageID, OfferID) & "<br />")
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
