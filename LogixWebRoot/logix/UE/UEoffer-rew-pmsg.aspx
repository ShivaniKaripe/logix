<%@ Page Language="vb" Debug="true" CodeFile="ue-cb.vb" Inherits="UECB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%
    ' *****************************************************************************
    ' * FILENAME: UEoffer-rew-pmsg.aspx 
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
    Dim Localization As Copient.Localization
    Dim rst As DataTable
    Dim row As DataRow
    Dim rst2 As DataTable
    Dim OfferID As Long
    Dim DeliverableID As Long
    Dim Name As String = ""
    Dim RewardID As Long
    Dim MessageID As Long
    Dim MessageTypeID As Long
    Dim BodyText As String = ""
    Dim Phase As Integer = 0
    Dim TouchPoint As Integer = 0
    Dim TpROID As Integer = 0
    Dim CreateROID As Integer = 0
    Dim PhaseTitle As String = ""
    Dim Disallow_Edit As Boolean = True
    Dim FromTemplate As Boolean = False
    Dim IsTemplate As Boolean = False
    Dim IsTemplateVal As String = "Not"
    Dim DisabledAttribute As String = ""
    Dim CloseAfterSave As Boolean = False
    Dim OpenTagEscape As String = Chr(1)
    Dim TextAreaBodyText As String = ""
    Dim TierText As String = ""
    Dim PrinterWidthBuf As New StringBuilder()
    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False
    Dim EngineID As Integer = 0
    Dim BannersEnabled As Boolean = True
    Dim Priority As Integer = 0
    Dim TierLevels As Integer = 1
    Dim t As Integer = 1
    Dim l = 1
    Dim ValidTiers As Boolean = True
    Dim ExcentusInstalled As Boolean = False
    Dim SuppressZeroBal As Boolean = False
    Dim DefaultSuppressed As Boolean = False
    Dim ShowSuppress As Boolean = False
    Dim PrinterPhrase As String = ""
    Dim Tag As String = ""
    Dim FinalTag As String = ""
    Dim CentralRendered As Boolean = False
    Dim LanguagesDT As DataTable
    Dim MLI As New Copient.Localization.MultiLanguageRec
    Dim MultiLanguageEnabled As Boolean = False
    Dim DefaultLanguageID As Integer = 0
    Dim PKID As Integer = 0
    Dim SortID As Integer = 0
    Dim RewardValue As Decimal
    Dim RewardRequired As Boolean = True
    Dim isAnyCustomer = False

    Dim SupportGlobalAndTieredConditions As Integer = MyCommon.Fetch_UE_SystemOption(197)
    Dim UseSameTierValue As Integer = 0
    Dim MLLanguageID As Integer = 1
    Dim isTranslatedOffer As Boolean = False
    Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean =False
    Dim bEnableAdditionalLockoutRestrictionsOnOffers As Boolean = False
    Dim bOfferEditable As Boolean =False

    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If

    Response.Expires = 0
    MyCommon.AppName = "UEoffer-rew-pmsg.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
    Localization = New Copient.Localization(MyCommon)

    MultiLanguageEnabled = IIf(MyCommon.Fetch_SystemOption(124) = "1", True, False)
    Integer.TryParse(MyCommon.Fetch_SystemOption(1), DefaultLanguageID)
    If DefaultLanguageID = 0 Then DefaultLanguageID = 1

    BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")

    Phase = MyCommon.Extract_Val(Request.QueryString("Phase"))
    If (Phase = 0) Then Phase = MyCommon.Extract_Val(Request.Form("Phase"))
    If (Phase = 0) Then Phase = 3
    Select Case Phase
        Case 1 ' Notification
            PhaseTitle = "term.pmsgnotification"
        Case 2 ' Accumulation
            PhaseTitle = "term.pmsgaccumulation"
        Case 3 ' Reward
            PhaseTitle = "term.pmsgreward"
        Case Else
            PhaseTitle = "term.pmsgreward"
    End Select

    OfferID = MyCommon.Extract_Val(Request.QueryString("OfferID"))
    If (OfferID = 0) Then OfferID = MyCommon.Extract_Val(Request.Form("OfferID"))
    'This would redirect to summary page if offer is awaiting deployment or if offer is awaiting recommendations.
    CheckIfValidOffer(MyCommon, OfferID)
    DeliverableID = MyCommon.Extract_Val(Request.QueryString("DeliverableID"))
    If (DeliverableID = 0) Then DeliverableID = MyCommon.Extract_Val(Request.Form("DeliverableID"))
    RewardID = MyCommon.Extract_Val(Request.QueryString("RewardID"))
    If (RewardID = 0) Then RewardID = MyCommon.Extract_Val(Request.Form("RewardID"))
    MessageID = MyCommon.Extract_Val(Request.QueryString("MessageID"))
    If (MessageID = 0) Then MessageID = MyCommon.Extract_Val(Request.Form("MessageID"))
    MessageTypeID = MyCommon.Extract_Val(Request.QueryString("type"))
    If (MessageTypeID = 0) Then MessageTypeID = MyCommon.Extract_Val(Request.Form("type"))

    isTranslatedOffer  =MyCommon.IsTranslatedUEOffer(OfferID,  MyCommon)
    bEnableRestrictedAccessToUEOfferBuilder= IIf(MyCommon.Fetch_SystemOption(249)="1",True,False)

    bEnableAdditionalLockoutRestrictionsOnOffers = IIf(MyCommon.Fetch_SystemOption(260) = "1", True, False)
    bOfferEditable = MyCommon.IsOfferEditablePastLockOutPeriod(Logix.UserRoles.EditOfferPastLockoutPeriod, MyCommon, OfferID)

    If (Request.QueryString("EngineID") <> "" Or Request.Form("EngineID") <> "") Then
        EngineID = MyCommon.Extract_Val(Request.QueryString("EngineID"))
        If EngineID = 0 Then
            EngineID = MyCommon.Extract_Val(Request.Form("EngineID"))
        End If
    Else
        MyCommon.QueryStr = "select EngineID from OfferIDs where OfferID=" & OfferID & ";"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
            EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), 0)
        End If
    End If
    isAnyCustomer = UEOffer_Has_AnyCustomer(MyCommon, OfferID)
    TouchPoint = MyCommon.Extract_Val(Request.Form("tp"))
    If (TouchPoint = 0) Then TouchPoint = MyCommon.Extract_Val(Request.QueryString("tp"))
    If (TouchPoint > 0) Then TpROID = MyCommon.Extract_Val(Request.Form("roid"))
    If (TouchPoint > 0 AndAlso TpROID = 0) Then TpROID = MyCommon.Extract_Val(Request.QueryString("roid"))

    If Phase = 3 Then
        MyCommon.QueryStr = "select TierLevels from CPE_RewardOptions with (NoLock) where RewardOptionID=" & RewardID & ";"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
            TierLevels = MyCommon.NZ(rst.Rows(0).Item("TierLevels"), 1)
        End If
    Else
        TierLevels = 1
    End If

    ' if not already set, use the default setting for whether this message should print a zero balance for an accumulation
    DefaultSuppressed = (MyCommon.Fetch_UE_SystemOption(44) = "1")
    SuppressZeroBal = (MyCommon.Extract_Val(Request.Form("SuppressZeroBal")) = 1)
    SortID = MyCommon.Extract_Val(Request.Form("SortID"))

    ' Fetch the name
    MyCommon.QueryStr = "Select IncentiveName,IsTemplate,FromTemplate from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
        Name = MyCommon.NZ(rst.Rows(0).Item("IncentiveName"), "")
        IsTemplate = MyCommon.NZ(rst.Rows(0).Item("IsTemplate"), False)
        FromTemplate = MyCommon.NZ(rst.Rows(0).Item("FromTemplate"), False)
    End If
    IsTemplateVal = IIf(IsTemplate, "IsTemplate", "Not")

    CreateROID = IIf(TpROID > 0, TpROID, RewardID)
    If (Request.Form("save") <> "" AndAlso IsValidMessageEntry(infoMessage, DefaultLanguageID, TierLevels)) Then
        If (OfferID > 0 AndAlso CreateROID > 0) Then
            If (MyCommon.Extract_Val(Request.Form("remLines")) >= 0) Then
                If (EngineID = 5) Then
                    ' Email engine allows creation of multiple printed messages
                    If (DeliverableID > 0) Then
                        ' Store the priority so the order stays the same after the edit.
                        MyCommon.QueryStr = "select Priority from CPE_Deliverables where DeliverableID=" & DeliverableID & ";"
                        rst = MyCommon.LRT_Select
                        If (rst.Rows.Count > 0) Then
                            Priority = MyCommon.NZ(rst.Rows(0).Item("Priority"), 0)
                        End If
                        MyCommon.QueryStr = "delete from CPE_Deliverables with (RowLock) where DeliverableID=" & DeliverableID & ";"
                        MyCommon.LRT_Execute()
                    End If
                Else
                    ' Deletes the existing printed message deliverables, since only one printed message per offer is currently supported
                    MyCommon.QueryStr = "delete from CPE_Deliverables with (RowLock) where RewardOptionID=" & CreateROID & " and DeliverableTypeID=4 and RewardOptionPhase=" & Phase & ";"
                    MyCommon.LRT_Execute()
                    MyCommon.QueryStr = "delete from PrintedMessages with (RowLock) where MessageID in " & _
                                        "(select OutputID as MessageID from CPE_Deliverables with (NoLock) where DeliverableID=" & DeliverableID & ");"
                    MyCommon.LRT_Execute()
                End If
                MessageID = CreateMessage(MessageID, MessageTypeID, SuppressZeroBal, SortID)
                If (MessageID > 0) Then
                    ' Write the tier values
                    t = 1
                    For t = 1 To TierLevels
                        CreateMessageTiers(MyCommon, MessageID, t, OpenTagEscape, DefaultLanguageID, RewardValue)
                    Next
                    RewardRequired = MyCommon.Extract_Val(GetCgiValue("requiredToDeliver"))

                    DeliverableID = AddMessageDeliverable(OfferID, CreateROID, MessageID, Phase, RewardRequired)
                    ' Is this is a printed message for the email engine, set the priority
                    If (EngineID = 5) Then
                        If (Priority > 0) Then
                            MyCommon.QueryStr = "update CPE_Deliverables with (RowLock) set Priority=" & Priority & " where DeliverableID=" & DeliverableID & ";"
                        Else
                            MyCommon.QueryStr = "update CPE_Deliverables with (RowLock) set Priority=" & _
                                                "(select Max(IsNull(Priority,0)) + 1 from CPE_Deliverables where RewardOptionID=" & CreateROID & ") " & _
                                                "where DeliverableID=" & DeliverableID & ";"
                        End If
                        MyCommon.LRT_Execute()
                    End If
                End If
                MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
                MyCommon.LRT_Execute()
                ResetOfferApprovalStatus(OfferID)
                MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup(IIf(Phase = 1, "CPE_not.createpmsg", "CPE_Reward.createpmsg"), LanguageID))
            Else
                infoMessage = Copient.PhraseLib.Lookup("reward.pmsgtoolong", LanguageID)
            End If
        End If
        If infoMessage = "" Then
            CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
        End If
    End If

    If (MessageID > 0 AndAlso RewardID > 0) Then
        MyCommon.QueryStr = "select PM.MessageTypeID, PMT.BodyText, PM.SuppressZeroBalance, PM.SortID, D.Required " & _
                            "from CPE_deliverables D with (NoLock) inner join PrintedMessages PM with (NoLock) on D.OutputID = PM.MessageID " & _
                            "inner join PrintedMessageTiers PMT with (NoLock) on PM.MessageID = PMT.MessageID " & _
                            "where D.Deleted = 0 and D.RewardOptionPhase=" & Phase & " and D.OutputID =" & MessageID & " and D.RewardOptionID=" & CreateROID & "and D.DeliverableTypeID=4 and PMT.TierLevel = 1;"
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
            MessageTypeID = MyCommon.NZ(rst.Rows(0).Item("MessageTypeID"), 0)
            BodyText = MyCommon.NZ(rst.Rows(0).Item("BodyText"), "")
            SuppressZeroBal = MyCommon.NZ(rst.Rows(0).Item("SuppressZeroBalance"), False)
            SortID = MyCommon.NZ(rst.Rows(0).Item("SortID"), 0)
            RewardRequired = MyCommon.NZ(rst.Rows(0).Item("Required"), True)
        End If

        'Get UseSameTierValue
        If TierLevels > 1 And SupportGlobalAndTieredConditions = 1 Then
            MyCommon.QueryStr = "select PM.MessageTypeID, PMT.TierLevel, PMT.BodyText " & _
                              "from CPE_deliverables D with (NoLock) inner join PrintedMessages PM with (NoLock) on D.OutputID = PM.MessageID " & _
                              "inner join PrintedMessageTiers PMT with (NoLock) on PM.MessageID = PMT.MessageID " & _
                              "where D.Deleted = 0 and D.RewardOptionPhase=" & Phase & " and D.OutputID =" & MessageID & " and D.RewardOptionID=" & CreateROID & "and D.DeliverableTypeID=4;"
            rst = MyCommon.LRT_Select
            If rst.Rows.Count > 0 Then
                Dim lastMessage As String = ""
                For Each row In rst.Rows
                    If (MyCommon.NZ(row.Item("TierLevel"), 0) > 1) and MyCommon.NZ(row.Item("BodyText"), "") <> lastMessage Then
                        UseSameTierValue = 0
                        Exit For
                    Else
                        lastMessage = MyCommon.NZ(row.Item("BodyText"), "")
                    End If
                Next
                If MyCommon.NZ(row.Item("TierLevel"), 0) = TierLevels Then
                    UseSameTierValue = 1
                End If
            End If
        Else
            UseSameTierValue = 0
        End If

    End If

    'update the templates permission if necessary
    If (Request.Form("save") <> "" AndAlso Request.Form("IsTemplate") = "IsTemplate" AndAlso infoMessage = "") Then
        ' time to update the status bits for the templates
        Dim form_Disallow_Edit As Integer = 0
        If (Request.Form("Disallow_Edit") = "on") Then
            form_Disallow_Edit = 1
        End If
        MyCommon.QueryStr = "update CPE_Deliverables with (RowLock) set DisallowEdit=" & form_Disallow_Edit & _
                            " where DeliverableID=" & DeliverableID
        MyCommon.LRT_Execute()
    End If

    If (IsTemplate Or FromTemplate) Then
        ' lets dig the permissions if its a template
        MyCommon.QueryStr = "select DisallowEdit from CPE_Deliverables with (NoLock) where DeliverableID=" & DeliverableID
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

    Send_HeadBegin("term.offer", PhaseTitle, OfferID)
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts()
    Send("<script type=""text/javascript"">")
    Send("function ChangeParentDocument() { ")
    If (Phase = 3) Then
        If (EngineID = 6) Then
            Send("  if (opener != null) {")
            Send("    var newlocation = '/logix/CAM/CAM-offer-rew.aspx?OfferID=" & OfferID & "'; ")
            Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
            Send("  opener.location = '/logix/CAM/CAM-offer-rew.aspx?OfferID=" & OfferID & "'; ")
        Else
            Send("  if (opener != null) {")
            Send("    var newlocation = 'UEoffer-rew.aspx?OfferID=" & OfferID & "'; ")
            Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
            Send("  opener.location = 'UEoffer-rew.aspx?OfferID=" & OfferID & "'; ")
        End If
    ElseIf (Phase = 1 Or Phase = 2) Then
        If (EngineID = 3) Then
            Send("  if (opener != null) {")
            Send("    var newlocation = '/logix/web-offer-not.aspx?OfferID=" & OfferID & "'; ")
            Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
            Send("  opener.location = '/logix/web-offer-not.aspx?OfferID=" & OfferID & "'; ")
        ElseIf (EngineID = 5) Then
            Send("  if (opener != null) {")
            Send("    var newlocation = '/logix/email-offer-not.aspx?OfferID=" & OfferID & "'; ")
            Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
            Send("  opener.location = '/logix/email-offer-not.aspx?OfferID=" & OfferID & "'; ")
        ElseIf (EngineID = 6) Then
            Send("  if (opener != null) {")
            Send("    var newlocation = '/logix/offer-channels.aspx?OfferID=" & OfferID & "'; ")
            Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
            Send("  opener.location = '/logix/offer-channels.aspx?OfferID=" & OfferID & "'; ")
        ElseIf (EngineID = 9) Then
            Send("  if (opener != null) {")
            Send("    var newlocation = '/logix/offer-channels.aspx?OfferID=" & OfferID & "'; ")
            Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
            Send("  opener.location = '/logix/offer-channels.aspx?OfferID=" & OfferID & "'; ")
        Else
            Send("  if (opener != null) {")
            Send("    var newlocation = '/logix/CPEoffer-not.aspx?OfferID=" & OfferID & "'; ")
            Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
            Send("  opener.location = '/logix/CPEoffer-not.aspx?OfferID=" & OfferID & "'; ")
        End If
    End If
    Send("} ")
    Send("} ")
    Send("} ")
    Send("</script>")
%>
<script type="text/javascript">
// JS QuickTags version 1.2
//
// Copyright (c) 2002-2005 Alex King
// http://www.alexking.org/
//
// Licensed under the LGPL license
// http://www.gnu.org/copyleft/lesser.html
//
// This JavaScript will insert the tags below at the cursor position in IE and 
// Gecko-based browsers (Mozilla, Camino, Firefox, Netscape). For browsers that 
// do not support inserting at the cursor position (Safari, OmniWeb) it appends
// the tags to the end of the content.

var edButtons = new Array();
var edOpenTags = new Array();
var caretPos = 0;
var caretTID = 0;
var bCapture = true;

caretID = setInterval("captureCursorPosition();", 200);

//
// Functions

function isIE() {
  return /msie/i.test(navigator.userAgent) && !/opera/i.test(navigator.userAgent);
}

function edShowButton(button) {
	if (button.access) {
		var accesskey = ' accesskey = "' + button.access + '"'
	} else {
		var accesskey = '';
	}
}

function edAddTag(button) {
	if (edButtons[button].tagEnd != '') {
		edOpenTags[edOpenTags.length] = button;
		document.getElementById(edButtons[button].id).value = '/' + document.getElementById(edButtons[button].id).value;
	}
}

function edRemoveTag(button) {
	for (i = 0; i < edOpenTags.length; i++) {
		if (edOpenTags[i] == button) {
			edOpenTags.splice(i, 1);
			document.getElementById(edButtons[button].id).value = document.getElementById(edButtons[button].id).value.replace('/', '');
		}
	}
}

function edCheckOpenTags(button) {
	var tag = 0;
	for (i = 0; i < edOpenTags.length; i++) {
		if (edOpenTags[i] == button) {
			tag++;
		}
	}
	if (tag > 0) {
		return true; // tag found
	} else {
		return false; // tag not found
	}
}

//
//
// Insertion code
function edInsertTag(myField, i) {
	//IE support
	if (document.selection) {
		myField.focus();
		sel = document.selection.createRange();
		if (sel.text.length > 0) {
			sel.text = edButtons[i].tagStart + sel.text + edButtons[i].tagEnd;
		} else {
			if (!edCheckOpenTags(i) || edButtons[i].tagEnd == '') {
				sel.text = edButtons[i].tagStart;
				edAddTag(i);
			}
			else {
				sel.text = edButtons[i].tagEnd;
				edRemoveTag(i);
			}
		}
		myField.focus();
	}
	//MOZILLA/NETSCAPE support
	else if (myField.selectionStart || myField.selectionStart == '0') {
		var startPos = myField.selectionStart;
		var endPos = myField.selectionEnd;
		var cursorPos = endPos;
		var scrollTop = myField.scrollTop;
		if (startPos != endPos) {
			myField.value = myField.value.substring(0, startPos)
			              + edButtons[i].tagStart
			              + myField.value.substring(startPos, endPos) 
			              + edButtons[i].tagEnd
			              + myField.value.substring(endPos, myField.value.length);
			cursorPos += edButtons[i].tagStart.length + edButtons[i].tagEnd.length;
		} else {
			if (!edCheckOpenTags(i) || edButtons[i].tagEnd == '') {
				myField.value = myField.value.substring(0, startPos) 
				              + edButtons[i].tagStart
				              + myField.value.substring(endPos, myField.value.length);
				edAddTag(i);
				cursorPos = startPos + edButtons[i].tagStart.length;
			} else {
				myField.value = myField.value.substring(0, startPos) 
				              + edButtons[i].tagEnd
				              + myField.value.substring(endPos, myField.value.length);
				edRemoveTag(i);
				cursorPos = startPos + edButtons[i].tagEnd.length;
			}
		}
		myField.focus();
		myField.selectionStart = cursorPos;
		myField.selectionEnd = cursorPos;
		myField.scrollTop = scrollTop;
	} else {
		if (!edCheckOpenTags(i) || edButtons[i].tagEnd == '') {
			myField.value += edButtons[i].tagStart;
			edAddTag(i);
		} else {
			myField.value += edButtons[i].tagEnd;
			edRemoveTag(i);
		}
		myField.focus();
	}
}

// text field to last have focus
var srcElement = document.getElementById('t1_text_<%Sendb(DefaultLanguageID)%>');

function edInsertContent(myValue) {
  var myField = document.getElementById(srcElement.name);
  var suppressSpan = document.getElementById('suppressBal');
  var s1 = myField.value.slice(0,caretPos);
  var s2 = s1.replace(/(\n|[\n])/g, "");
  var lfDiff = (s1.length - s2.length);
  
  if (isIE()) {
    if (getInternetExplorerVersion() < 9) {
      caretPos = caretPos - lfDiff;
    }
  }
  setCursorPosition();
  if (isIE()) {
    caretPos = caretPos + myValue.length;
  }
	//IE support
	if (document.selection) {
		myField.focus();
		sel = document.selection.createRange();
		sel.text = myValue;
		myField.focus();
	}
	//MOZILLA/NETSCAPE support
	else if (myField.selectionStart || myField.selectionStart == '0') {
		var startPos = myField.selectionStart;
		var endPos = myField.selectionEnd;
		var scrollTop = myField.scrollTop;
		myField.value = myField.value.substring(0, startPos)
		              + myValue 
                      + myField.value.substring(endPos, myField.value.length);
		myField.focus();
		myField.selectionStart = startPos + myValue.length;
		myField.selectionEnd = startPos + myValue.length;
		myField.scrollTop = scrollTop;
	} else {
		myField.value += myValue;
		myField.focus();
	}
	
	// check if the tag should make the suppress zero balance option available
	if (suppressSpan != null) {
	  if (myValue.indexOf("|TOTALPOINTS|")>-1 || myValue.indexOf("|PTSASPEN|")>-1 
	      || myValue.indexOf("|ACCUMAMT|")>-1 || myValue.indexOf("|SVBAL")>-1
	      || myValue.indexOf("|SVREDEEM")>-1 || myValue.indexOf("|SVVAL")>-1
	      || myValue.indexOf("|PTBAL")>-1 || myValue.indexOf("|PTVAL")>-1
	      || myValue.indexOf("|NET#")>-1 || myValue.indexOf("|INITIAL#")>-1
	      || myValue.indexOf("|EARNED#")>-1 || myValue.indexOf("|REDEEMED#")>-1) {
  	  suppressSpan.style.visibility = 'visible';
  	}
	}	
}

// ~~~~~~~~~~ DYNAMICALLY-GENERATED TAG INSERT FUNCTIONS BEGIN HERE ~~~~~~~~~~
<%
  If MyCommon.IsEngineInstalled(8) Then
    ExcentusInstalled = True
  End If
  MyCommon.QueryStr = "select distinct MT.MarkupID, MT.Tag, MT.Description, MT.PhraseID, MT.NumParams, " & _
                      "MT.Param1Name, MT.Param1PhraseID, MT.Param2Name, MT.Param2PhraseID, " & _
                      "MT.Param3Name, MT.Param3PhraseID, MT.Param4Name, MT.Param4PhraseID, " & _
                      "MT.DisplayOrder, MT.CentralRendered, MT.ButtonText, " & _
                      "MTU.RewardTypeID, MTU.EngineID from MarkupTags as MT with (NoLock) " & _
                      "inner join MarkupTagUsage as MTU with (NoLock) on MT.MarkupID=MTU.MarkupID " & _
                      "where MTU.RewardTypeID=4 "
  If ExcentusInstalled Then
    MyCommon.QueryStr &= "and MTU.EngineID in (" & EngineID & ", 8) order by MT.DisplayOrder;"
  Else
    MyCommon.QueryStr &= "and MTU.EngineID=" & EngineID & " order by MT.DisplayOrder;"
  End If
  rst = MyCommon.LRT_Select
  Dim funcname As String
  For Each row In rst.Rows
    Tag = MyCommon.NZ(row.Item("ButtonText"), "")
    CentralRendered = IIf(row.Item("CentralRendered"), True, False)
    funcname = MyCommon.NZ(row.Item("ButtonText"), "")
    funcname = funcname.Replace("#", "Amt")
    funcname = funcname.Replace("$", "Dol")
    funcname = funcname.Replace("/", "Off")
    If (MyCommon.NZ(row.Item("NumParams"), 0) = 0) Then
      Send("function edInsert" & (StrConv(funcname, VbStrConv.ProperCase)) & "() {")
      Send(IIf(CentralRendered, "  myValue = '\" & row.Item("Tag") & "';", "  myValue = '|" & row.Item("Tag") & "|';"))
      Send("  edInsertContent(myValue);")
      Send("}")
    ElseIf (MyCommon.NZ(row.Item("NumParams"), 0) = 1) Then
      If (Tag = "UPCA") or (Tag = "UPCB") or (Tag = "EAN13") or (Tag = "CODE39") or (Tag="CODE128") Then
        Send("function edInsert" & (StrConv(funcname, VbStrConv.ProperCase)) & "() {")
        Send("  var myValue = prompt('" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("Param1PhraseID"), 0), LanguageID) & "', '');")
        Send("  if (myValue) {")
        Send("    myValue = '|" & row.Item("Tag") & "[' + myValue + ']|';")
        Send("    edInsertContent(myValue);")
        Send("  }")
        Send("}")
      ElseIf (Tag = "VERSION") Or (Tag = "LOCALMSG") Or (Tag = "ASCII") Or (Tag = "HEX") Or (Tag = "TAB") Or (Tag = "ADJUSTLINE") Or (Tag = "FONT") Or (Tag = "ZONE") Or (Tag = "PREMSG") Or (Tag = "ASCIILIST") Then
        Send("function edInsert" & (StrConv(funcname, VbStrConv.ProperCase)) & "() {")
        Send("  var finalTag = '\" & row.Item("Tag") & "';")
        Send("  var myValue1 = document.getElementById(""monoparameter1"").value;")
        Send("  finalTag = finalTag.replace('^1', myValue1);")
        Send("  document.getElementById(""monoparameter1"").value = '';")
        Send("  edInsertContent(finalTag);")
        Send("}")
      Else
        Send("function edInsert" & (StrConv(funcname, VbStrConv.ProperCase)) & "() {")
        If StrConv(funcname, VbStrConv.ProperCase) = "Scorecard" Then
          Send("  var myValue = document.getElementById(""functionselect4"").value;")
        ElseIf StrConv(funcname, VbStrConv.ProperCase) = "Svscorecard" Then
          Send("  var myValue = document.getElementById(""functionselect6"").value;")
        ElseIf StrConv(funcname, VbStrConv.ProperCase) = "Ptbal" OrElse StrConv(funcname, VbStrConv.ProperCase) = "Ptval" _
               OrElse StrConv(funcname, VbStrConv.ProperCase) = "Netamt" OrElse StrConv(funcname, VbStrConv.ProperCase) = "Initialamt" _
               OrElse StrConv(funcname, VbStrConv.ProperCase) = "Earnedamt" OrElse StrConv(funcname, VbStrConv.ProperCase) = "Redeemedamt" Then
          Send("  var myValue = document.getElementById(""functionselect2"").value;")
        ElseIf StrConv(funcname, VbStrConv.ProperCase) = "Dscorecard" Then
          Send("  var myValue = document.getElementById(""functionselect7"").value;")
        Else
          Send("  var myValue = document.getElementById(""functionselect"").value;")
        End If
        Send("  if (myValue) {")
        Send("    myValue = '|" & row.Item("Tag") & "[' + myValue + ']|';")
        Send("    edInsertContent(myValue);")
        Send("  }")
        Send("}")
      End If
    ElseIf (MyCommon.NZ(row.Item("NumParams"), 0) = 2) Then
      If (Tag = "UPCA") Or (Tag = "CUTGAP") Or (Tag = "OTHER") Or (Tag = "ZONEACTION") Then
        Send("function edInsert" & (StrConv(funcname, VbStrConv.ProperCase)) & "() {")
        Send("  var finalTag = '\" & row.Item("Tag") & "';")
        Send("  var myValue1 = document.getElementById(""diparameter1"").value;")
        Send("  var myValue2 = document.getElementById(""diparameter2"").value;")
        Send("  finalTag = finalTag.replace('^1', myValue1);")
        Send("  finalTag = finalTag.replace('^2', myValue2);")
        Send("  document.getElementById(""diparameter1"").value = '';")
        Send("  document.getElementById(""diparameter2"").value = '';")
        Send("  edInsertContent(finalTag);")
        Send("}")
      End If
    ElseIf (MyCommon.NZ(row.Item("NumParams"), 0) = 3) Then
      If (Tag = "SVRATIO") Then
        Send("function edInsert" & (StrConv(funcname, VbStrConv.ProperCase)) & "() {")
        Send("  var myValue1 = document.getElementById(""functionselect"").value;")
        Send("  var myValue2 = document.getElementById(""incrementqty"").value;")
        Send("  var myValue3 = document.getElementById(""incrementamt"").value;")
        Send("  if (myValue1 && myValue2 && myValue3) {")
        Send("    myValue = '|" & row.Item("Tag") & "[' + myValue1 + ',' + myValue2 + ',' + myValue3 + ']|';")
        Send("    edInsertContent(myValue);")
        Send("  }")
        Send("}")
      ElseIf (Tag = "SVSCRATIO") Then
        Send("function edInsert" & (StrConv(funcname, VbStrConv.ProperCase)) & "() {")
        Send("  var myValue1 = document.getElementById(""functionselect6"").value;")
        Send("  var myValue2 = document.getElementById(""incrementqty"").value;")
        Send("  var myValue3 = document.getElementById(""incrementamt"").value;")
        Send("  if (myValue1 && myValue2 && myValue3) {")
        Send("    myValue = '|" & row.Item("Tag") & "[' + myValue1 + ',' + myValue2 + ',' + myValue3 + ']|';")
        Send("    edInsertContent(myValue);")
        Send("  }")
        Send("}")
      ElseIf (Tag = "BOX") Or (Tag = "BARCODE128") Or (Tag = "IF") Then
        Send("function edInsert" & (StrConv(funcname, VbStrConv.ProperCase)) & "() {")
        Send("  var finalTag = '\" & row.Item("Tag") & "';")
        Send("  var myValue1 = document.getElementById(""triparameter1"").value;")
        Send("  var myValue2 = document.getElementById(""triparameter2"").value;")
        Send("  var myValue3 = document.getElementById(""triparameter3"").value;")
        Send("  finalTag = finalTag.replace('^1', myValue1);")
        Send("  finalTag = finalTag.replace('^2', myValue2);")
        Send("  finalTag = finalTag.replace('^3', myValue3);")
        Send("  document.getElementById(""triparameter1"").value = '';")
        Send("  document.getElementById(""triparameter2"").value = '';")
        Send("  document.getElementById(""triparameter3"").value = '';")
        Send("  edInsertContent(finalTag);")
        Send("}")
      End If
    ElseIf (MyCommon.NZ(row.Item("NumParams"), 0) = 4) Then
      If (Tag = "CUSTLINE") Or (Tag = "BOXTYPE") Then
        Send("function edInsert" & (StrConv(funcname, VbStrConv.ProperCase)) & "() {")
        Send("  var finalTag = '\" & row.Item("Tag") & "';")
        Send("  var myValue1 = document.getElementById(""tetraparameter1"").value;")
        Send("  var myValue2 = document.getElementById(""tetraparameter2"").value;")
        Send("  var myValue3 = document.getElementById(""tetraparameter3"").value;")
        Send("  var myValue4 = document.getElementById(""tetraparameter4"").value;")
        Send("  finalTag = finalTag.replace('^1', myValue1);")
        Send("  finalTag = finalTag.replace('^2', myValue2);")
        Send("  finalTag = finalTag.replace('^3', myValue3);")
        Send("  finalTag = finalTag.replace('^4', myValue4);")
        Send("  document.getElementById(""tetraparameter1"").value = '';")
        Send("  document.getElementById(""tetraparameter2"").value = '';")
        Send("  document.getElementById(""tetraparameter3"").value = '';")
        Send("  document.getElementById(""tetraparameter4"").value = '';")
        Send("  edInsertContent(finalTag);")
        Send("}")
      End If
    End If
  Next
%>
// ~~~~~~~~~~ DYNAMICALLY-GENERATED TAG INSERT FUNCTIONS END HERE ~~~~~~~~~~


function doPreviewPopup() {
  var pSelect = document.getElementById('printerselect').value;
  var myField = "";
  if (srcElement != null) {
    myField = document.getElementById(srcElement.name);
  } 
  else {
    myField = document.getElementById('t1_text_<%Sendb(DefaultLanguageID)%>');
  }
  var numtiers = document.getElementById('TierLevels').value;
  var text = "";
  if (srcElement != null) {
    text = "" + srcElement.name;
  }
  else {
     text = "" + document.getElementById('t1_text_<%Sendb(DefaultLanguageID)%>').name;
  }
  var message = '';

  if (numtiers == '1') {
    var text1 = ""
  }
  else {
    var text1 = '<%Sendb(Copient.PhraseLib.Lookup("term.tier", languageID))%> " + text.substring(1,2) + ": ';
  }

  message = myField.value;
  message = message.replace(/</g,"&lt;");
  message = message.replace(/>/g,"&gt;");
  message = escape(message);

  var myUrl = 'UEoffer-rew-pmsgpreview.aspx?PrinterTypeID=' + pSelect + '&Message=' + message + '&TierLevel=' + text1;
  openPreviewPopup(myUrl);
}
</script>
<script type="text/javascript">
  var msgTimer = setInterval('checkMessage()', 3000);
// This is the javascript array holding the function list
// The PrintJavascriptArray ASP function can be used to print this array.
<%
  Dim QueryList(1) As String 
 
  If(isAnyCustomer) Then
    QueryList(0)="Select SVP.SVProgramID as ID, SVP.Name from StoredValuePrograms SVP inner join dbo.SVProgramsPromoEngineSettings as PESVP  with (NoLock) on SVP.SVProgramID =PESVP.SVProgramID where  EngineID=9 and Deleted=0 And AllowAnyCustomer=1 order by Name;"
  Else
    QueryList(0)= "Select SVProgramID as ID, Name  from StoredValuePrograms with (NoLock) where Deleted=0 order by Name;"
  End If
 
  If (isAnyCustomer) Then
    QueryList(1)= "Select PP.ProgramID  as ID, PP.ProgramName  as Name from PointsPrograms PP inner join dbo.PointsProgramsPromoEngineSettings as PEPP  with (NoLock) on PP.ProgramID =PEPP.ProgramID where  EngineID=9 and  Deleted=0 And AllowAnyCustomer=1 order by ProgramName;"
  Else
    QueryList(1) = "Select ProgramID  as ID, ProgramName as Name from PointsPrograms with (NoLock) where Deleted=0 order by ProgramName;"
  End If

  
    For ct as Integer = 0 To QueryList.GetUpperBound(0)
    MyCommon.QueryStr = QueryList(ct)
    rst = MyCommon.LRT_Select
    If (rst.rows.count > 0)
      Sendb("var functionlist" & IIf(ct > 0, ct+1, "") & " = Array(")
      For Each row In rst.Rows
        Sendb("""" & MyCommon.NZ(row.item("Name"), "").ToString().Replace("""", "\""") & """,")
      Next
      Send(""""");")
      Sendb("var vallist" & IIf(ct > 0, ct+1, "") & "= Array(")
      For Each row In rst.Rows
        Sendb("""" & MyCommon.NZ(row.item("ID"), 0) & """,")
      Next
      Send(""""");")
    Else
      Send("var functionlist = Array("""");")
      Send("var vallist = Array(""-99"");")
    End If
  Next
%>

// This is the function that refreshes the list after a keypress.
// The maximum number to show can be limited to improve performance with
// huge lists (1000s of entries).
// The function clears the list, and then does a linear search through the
// globally defined array and adds the matches back to the list.
function handleKeyUp(maxNumToShow, listIndex) {
  var selectObj, textObj, functionListLength;
  var i,  numShown;
  var searchPattern;
  var flist, vlist;
  var selectorIndex = '';

  selectorIndex = (listIndex > 1) ? listIndex : '';

  document.getElementById("functionselect" + selectorIndex).size = "10";
  
  // Set references to the form elements
  if (listIndex == 2) {
    // points program list
    selectObj = document.forms[0].functionselect2;
    textObj = document.forms[0].functioninput2;
    flist = functionlist2;
    vlist = vallist2;
    searchPattern = (document.forms[0].functionradio2[0].checked) ?  "^"+textObj.value : textObj.value;
  } else {
    // stored value program list
    selectObj = document.forms[0].functionselect;
    textObj = document.forms[0].functioninput;
    flist = functionlist;
    vlist = vallist;
    searchPattern = (document.forms[0].functionradio[0].checked) ?  "^"+textObj.value : textObj.value;
  }
  
  // Remember the function list length for loop speedup
  functionListLength = flist.length;
  
  searchPattern = cleanRegExpString(searchPattern);

  // Create a regulare expression
  re = new RegExp(searchPattern,"gi");
  
  // Clear the options list
  selectObj.length = 0;
  
  // Loop through the array and re-add matching options
  numShown = 0;
  for(i = 0; i < functionListLength; i++) {
    if(flist[i].search(re) != -1) {
      selectObj[numShown] = new Option(flist[i],vlist[i]);
      numShown++;
    }
    // Stop when the number to show is reached
    if(numShown == maxNumToShow) {
      break;
    }
  }
  // When options list whittled to one, select that entry
  if(selectObj.length == 1) {
    selectObj.options[0].selected = true;
  }
}

// this function gets the selected value and loads the appropriate
// php reference page in the display frame
// it can be modified to perform whatever action is needed, or nothing
function handleSelectClick(type) {
  if (type == 2) {
    selectObj = document.forms[0].functionselect2;
    textObj = document.forms[0].functioninput2;
    selectedValue = document.getElementById("functionselect2").value;
  } else if (type == 4) {
    selectObj = document.forms[0].functionselect4;
    textObj = document.forms[0].functioninput4;
    selectedValue = document.getElementById("functionselect4").value;
  } else if (type == 6) {
    selectObj = document.forms[0].functionselect6;
    textObj = document.forms[0].functioninput6;
    selectedValue = document.getElementById("functionselect6").value;
  } else if (type == 7) {
    selectObj = document.forms[0].functionselect7;
    textObj = document.forms[0].functioninput7;
    selectedValue = document.getElementById("functionselect7").value;
  } else if (type == 11) {
    textObj = document.forms[0].monoparameter1;
    selectedValue = document.getElementById("monoparameter1").value;
  } else if (type == 21) {
    textObj = document.forms[0].diparameter1;
    selectedValue = document.getElementById("diparameter1").value;
  } else if (type == 31) {
    textObj = document.forms[0].triparameter1;
    selectedValue = document.getElementById("triparameter1").value;
  } else if (type == 41) {
    textObj = document.forms[0].tetraparameter1;
    selectedValue = document.getElementById("tetraparameter1").value;
  } else {
    selectObj = document.forms[0].functionselect;
    textObj = document.forms[0].functioninput;
    selectedValue = document.getElementById("functionselect").value;
  }
  if ((selectedValue != "") || ((selectedValue == "") && (type > 10))) {
    if (type ==2) {
      var elemTag = document.getElementById("ptTagName");
    } else if (type == 4) {
      var elemTag = document.getElementById("scTagName");
    } else if (type == 6) {
      var elemTag = document.getElementById("svscTagName");
    } else if (type == 7) {
      var elemTag = document.getElementById("dscTagName");
    } else if (type == 11) {
      var elemTag = document.getElementById("monoparametricTagName");
    } else if (type == 21) {
      var elemTag = document.getElementById("diparametricTagName");
    } else if (type == 31) {
      var elemTag = document.getElementById("triparametricTagName");
    } else if (type == 41) {
      var elemTag = document.getElementById("tetraparametricTagName");
    } else {
      var elemTag = document.getElementById("svTagName");
    }

    // Test resumes here
    if (elemTag.value == "Ptbal") {
      edInsertPtbal("t1_text", selectedValue);
    } else if (elemTag.value == "Ptval") {
      edInsertPtval("t1_text", selectedValue);
    } else if (elemTag.value == "Svbal") {
      edInsertSvbal("t1_text", selectedValue);
    } else if (elemTag.value == "Svval") {
      edInsertSvval("t1_text", selectedValue);
    } else if (elemTag.value == "Svbalexp") {
      edInsertSvbalexp("t1_text", selectedValue);
    } else if (elemTag.value == "Svvalexp") {
      edInsertSvvalexp("t1_text", selectedValue);
    } else if (elemTag.value == "Svlimit") {
      edInsertSvlimit("t1_text", selectedValue);
    } else if (elemTag.value == "Svredeem") {
      edInsertSvredeem("t1_text", selectedValue);
    } else if (elemTag.value == "Scorecard") {
      edInsertScorecard("t1_text", selectedValue);
    } else if (elemTag.value == "Svscorecard") {
      edInsertSvscorecard("t1_text", selectedValue);
    } else if (elemTag.value == "Dscorecard") {
      edInsertDscorecard("t1_text", selectedValue);
    } else if (elemTag.value == "Svratio") {
      edInsertSvratio("t1_text", selectedValue);
    } else if (elemTag.value == "Svscratio") {
      edInsertSvscratio("t1_text", selectedValue);
    // New Rite Aid tags: monoparametric
    } else if (elemTag.value == "Tab") {
      edInsertTab("t1_text", selectedValue);
    } else if (elemTag.value == "Version") {
      edInsertVersion("t1_text", selectedValue);
    } else if (elemTag.value == "Localmsg") {
      edInsertLocalmessage("t1_text", selectedValue);
    } else if (elemTag.value == "Ascii") {
      edInsertAscii("t1_text", selectedValue);
    } else if (elemTag.value == "Hex") {
      edInsertHex("t1_text", selectedValue);
    } else if (elemTag.value == "Adjustline") {
      edInsertAdjustline("t1_text", selectedValue);
    } else if (elemTag.value == "Font") {
      edInsertFont("t1_text", selectedValue);
    } else if (elemTag.value == "Zone") {
      edInsertZone("t1_text", selectedValue);
    } else if (elemTag.value == "Premsg") {
      edInsertPremsg("t1_text", selectedValue);
    } else if (elemTag.value == "Asciilist") {
      edInsertAsciilist("t1_text", selectedValue);
    // New Rite Aid tags: diparametric
    } else if (elemTag.value == "Upca") {
      edInsertUpca("t1_text", selectedValue);
    } else if (elemTag.value == "Cutgap") {
      edInsertCutgap("t1_text", selectedValue);
    } else if (elemTag.value == "Other") {
      edInsertOther("t1_text", selectedValue);
    } else if (elemTag.value == "Zoneaction") {
      edInsertZoneaction("t1_text", selectedValue);
    // New Rite Aid tags: triparametric
    } else if (elemTag.value == "Box") {
      edInsertBox("t1_text", selectedValue);
    } else if (elemTag.value == "Barcode128") {
      edInsertBarcode128("t1_text", selectedValue);
    } else if (elemTag.value == "If") {
      edInsertIf("t1_text", selectedValue);
    // New Rite Aid tags: triparametric
    } else if (elemTag.value == "Custline") {
      edInsertCustline("t1_text", selectedValue);
    } else if (elemTag.value == "Boxtype") {
      edInsertBoxtype("t1_text", selectedValue);
    } else if (elemTag.value == "Net#") {
      edInsertNetamt("t1_text", selectedValue);
    } else if (elemTag.value == "Earned#") {
      edInsertEarnedamt("t1_text", selectedValue);
    } else if (elemTag.value == "Initial#") {
      edInsertInitialamt("t1_text", selectedValue);
    } else if (elemTag.value == "Redeemed#") {
      edInsertRedeemedamt("t1_text", selectedValue);
    }
    showDialogSpan(false, 3, "")
  }
}

function showDialogSpan(bShow, type, caption, label1, label2, label3, label4) {
  var elemBox = document.getElementById("dialogbox");
  var elemPt = document.getElementById("pointselector");
  var elemPtTag = document.getElementById("ptTag");
  var elemSv = document.getElementById("svselector");
  var elemSvTag = document.getElementById("svTag");
  var elemSc = document.getElementById("scselector");
  var elemScTag = document.getElementById("scTag");
  var elemSvsc = document.getElementById("svscselector");
  var elemSvscTag = document.getElementById("svscTag");
  var elemDsc = document.getElementById("dscselector");
  var elemDscTag = document.getElementById("dscTag");
  var elemRr = document.getElementById("rrValues");
  var elemRrTag = document.getElementById("rrTag");
  var elemMonoparametric = document.getElementById("monoparametric");
  var elemMonoparametricTag = document.getElementById("monoparametricTag");
  var elemDiparametric = document.getElementById("diparametric");
  var elemDiparametricTag = document.getElementById("diparametricTag");
  var elemTriparametric = document.getElementById("triparametric");
  var elemTriparametricTag = document.getElementById("triparametricTag");
  var elemTetraparametric = document.getElementById("tetraparametric");
  var elemTetraparametricTag = document.getElementById("tetraparametricTag");
  var elemTag = null;
  
  if (bShow) {
    if (elemPt != null && type == 2) {
      elemPt.style.display = "block";
      if (caption != "" && elemPtTag != null) {
        elemPtTag.innerHTML = "Tag type: " + caption
        elemTag = document.getElementById("ptTagName");
        if (elemTag != null) {
          elemTag.value = caption;
        }
      }
    }

    if (elemSv != null && type == 3) {
      elemSv.style.display = "block";
      if (caption != "" && elemSvTag != null) {
        elemSvTag.innerHTML = "Tag type: " + caption
        elemTag = document.getElementById("svTagName");
        if (elemTag != null) {
          elemTag.value = caption;
        }
      }
    }
    if (elemSc != null && type == 4) {
      elemSc.style.display = "block";
      if (caption != "" && elemScTag != null) {
        elemScTag.innerHTML = "Tag type: " + caption
        elemTag = document.getElementById("scTagName");
        if (elemTag != null) {
          elemTag.value = caption;
        }
      }
    }
    if (elemRr != null && type == 5) {
      elemRr.style.display = "block";
      if (caption != "" && elemRrTag != null) {
        elemRrTag.innerHTML = "Tag type: " + caption
        elemTag = document.getElementById("rrTagName");
        if (elemTag != null) {
          elemTag.value = caption;
        }
      }
    }
    if (elemSvsc != null && type == 6) {
      elemSvsc.style.display = "block";
      if (caption != "" && elemSvscTag != null) {
        elemSvscTag.innerHTML = "Tag type: " + caption
        elemTag = document.getElementById("svscTagName");
        if (elemTag != null) {
          elemTag.value = caption;
        }
      }
    }
    if (elemDsc != null && type == 7) {
      elemDsc.style.display = "block";
      if (caption != "" && elemDscTag != null) {
        elemDscTag.innerHTML = "Tag Type: " + caption
        elemTag = document.getElementById("dscTagName");
        if (elemTag != null) {
          elemTag.value = caption;
        }
      }
    }
    if (elemMonoparametric != null && type == 11) {
      elemMonoparametric.style.display = "block";
      if (caption != "" && elemMonoparametricTag != null) {
        elemMonoparametricTag.innerHTML = "Tag type: " + caption
        document.getElementById("monoparametricLabel1").innerHTML = label1;
        elemTag = document.getElementById("monoparametricTagName");
        if (elemTag != null) {
          elemTag.value = caption;
        }
      }
    }
    if (elemDiparametric != null && type == 21) {
      elemDiparametric.style.display = "block";
      if (caption != "" && elemDiparametricTag != null) {
        elemDiparametricTag.innerHTML = "Tag type: " + caption
        document.getElementById("diparametricLabel1").innerHTML = label1;
        document.getElementById("diparametricLabel2").innerHTML = label2;
        elemTag = document.getElementById("diparametricTagName");
        if (elemTag != null) {
          elemTag.value = caption;
        }
      }
    }
    if (elemTriparametric != null && type == 31) {
      elemTriparametric.style.display = "block";
      if (caption != "" && elemTriparametricTag != null) {
        elemTriparametricTag.innerHTML = "Tag type: " + caption
        document.getElementById("triparametricLabel1").innerHTML = label1;
        document.getElementById("triparametricLabel2").innerHTML = label2;
        document.getElementById("triparametricLabel3").innerHTML = label3;
        elemTag = document.getElementById("triparametricTagName");
        if (elemTag != null) {
          elemTag.value = caption;
        }
      }
    }
    if (elemTetraparametric != null && type == 41) {
      elemTetraparametric.style.display = "block";
      if (caption != "" && elemTetraparametricTag != null) {
        elemTetraparametricTag.innerHTML = "Tag type: " + caption
        document.getElementById("tetraparametricLabel1").innerHTML = label1;
        document.getElementById("tetraparametricLabel2").innerHTML = label2;
        document.getElementById("tetraparametricLabel3").innerHTML = label3;
        document.getElementById("tetraparametricLabel4").innerHTML = label4;
        elemTag = document.getElementById("tetraparametricTagName");
        if (elemTag != null) {
          elemTag.value = caption;
        }
      }
    }
  }
  if (elemBox != null) {
    elemBox.style.display = (bShow) ? "block" : "none";
    elemPt.style.display = (bShow && type == 2) ? "block" : "none";
    elemSv.style.display = (bShow && type == 3) ? "block" : "none";
    elemSc.style.display = (bShow && type == 4) ? "block" : "none";
    elemRr.style.display = (bShow && type == 5) ? "block" : "none";
    elemSvsc.style.display = (bShow && type == 6) ? "block" : "none";
    elemDsc.style.display = (bShow && type == 7) ? "block" : "none";
    elemMonoparametric.style.display = (bShow && type == 11) ? "block" : "none";
    elemDiparametric.style.display = (bShow && type == 21) ? "block" : "none";
    elemTriparametric.style.display = (bShow && type == 31) ? "block" : "none";
    elemTetraparametric.style.display = (bShow && type == 41) ? "block" : "none";
  }
}

function showSvSelect() {
  showDialogSpan(false, 5);
  showDialogSpan(true, 3, "Svratio");
  document.getElementById("rrTagName").value = "Svratio";
}
function showScSelect() {
  showDialogSpan(false, 5)
  showDialogSpan(true, 6, "Svscratio");
  document.getElementById("rrTagName").value = "Svscratio";
}
function showNextSelect() {
  if (document.getElementById("rrTagName").value == "Svratio") {
    showSvSelect();
  } else if (document.getElementById("rrTagName").value == "Svscratio") {
    showScSelect();
  }
}

function xmlhttpPost(strURL) {
  var xmlHttpReq = false;
  var self = this;
  
  document.getElementById("tools").innerHTML = "<div class=\"loading\"><img id=\"clock\" src=\"../../images/clock22.png\" \/><br \/>" + '<% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID)) %><\/div>';
  
  // Mozilla/Safari
  if (window.XMLHttpRequest) {
    self.xmlHttpReq = new XMLHttpRequest();
  }
  // IE
  else if (window.ActiveXObject) {
    self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
  }
  strURL += "?" + getQueryString();
  self.xmlHttpReq.open('POST', strURL, true);
  self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
  self.xmlHttpReq.onreadystatechange = function() {
    if (self.xmlHttpReq !=null && self.xmlHttpReq.readyState == 4) {
      updatepage(self.xmlHttpReq.responseText, <% Sendb(TierLevels) %>);
    }
  }
  self.xmlHttpReq.send(getQueryString());
}

function getQueryString() {    
  var selElem = document.getElementById('printerselect');
  var qstr = "";
  var textAreaName = "t1_text_<%Sendb(DefaultLanguageID)%>";
  
  if(selElem != null) {
    qstr = "Mode=MarkupTags&OfferID=<%Sendb(OfferID)%>&EngineID=<%Sendb(EngineID)%>&Phase=<%Sendb(Phase)%>&PrinterTypeID=" + selElem.value + "&TextAreaName=" + textAreaName
  }
  return qstr;
}

function updatepage(str, t){
  var selElem = document.getElementById('printerselect');
  var ptWidthElem = null;
  var ptLinesElem = null;
  var taElem = document.getElementById("t1_text_<%Sendb(DefaultLanguageID)%>");
  var languageIDs = document.getElementById("LanguageIDs").value.split(',');
  var i = 1;
  
  document.getElementById("tools").innerHTML = str;
  
  if (selElem != null) {
    ptWidthElem = document.getElementById("PT" + selElem.value);
    ptLinesElem = document.getElementById("PT" + selElem.value + "lines");
    lineCounter();
    for (i = 1; i <= t; i++) { // TIER LOOP BEGINS
      for (var j = 0; j < languageIDs.length; j++) { // LANGUAGE LOOP BEGINS
        
        taElem = document.getElementById("t" + i + "_text_" + languageIDs[j]);
      // Set the width
      if (ptWidthElem != null) {
        if (taElem != null) {
          <%
          If (Request.Browser.Type.ToString.ToUpper.IndexOf("FIREFOX") > -1) Then
            Send("taElem.style.width = ((parseInt(ptWidthElem.value) * 8) + 19) + 'px';")
          Else
            Send("taElem.style.width = ((parseInt(ptWidthElem.value) * 8) + 23) + 'px';")
          End If
          %>
        }
      } else {
        if (taElem != null) {
          <%
          If (Request.Browser.Type.ToString.ToUpper.IndexOf("FIREFOX") > -1) Then
            Send("taElem.style.width = 434 + 'px';")
          Else
            Send("taElem.style.width = 438 + 'px';")
          End If
          %>
        }
      }
      // Set the height
      if ((ptLinesElem != null) && (ptLinesElem.value >= 1) && (ptLinesElem.value <= 20)) {
        if (taElem != null) {
          //taElem.style.overflow = 'hidden';
          <%
          If (Request.Browser.Type.ToString.ToUpper.IndexOf("FIREFOX") > -1) Then
            Send("taElem.style.height = (parseInt(ptLinesElem.value) * 16) + 'px';")
          Else
            Send("taElem.style.height = ((parseInt(ptLinesElem.value) * 16) + 6) + 'px';")
          End If
          %>
        }
      } else {
        if (taElem != null) {
          //taElem.style.overflow = 'auto';
          <%
          If (Request.Browser.Type.ToString.ToUpper.IndexOf("FIREFOX") > -1) Then
            Send("taElem.style.height = 192 + 'px';")
          Else
            Send("taElem.style.height = 192 + 'px';")
          End If
          %>
        }
      }
      // Set the wrappiness
      if (selElem.value == 9) {
        if (taElem != null) {
          setWrap(taElem, "off");
        }
      } else {
        if (taElem != null) {
          setWrap(taElem, "hard");
        }
      }

      } // LANGUAGE LOOP ENDS
    } // TIER LOOP ENDS
    }
  
}

function setWrap(area, wrap) {
  if (area.wrap) {
    area.wrap = wrap;
  } else { // wrap attribute not supported - try Mozilla workaround
    area.setAttribute('wrap', wrap);
    var newarea= area.cloneNode(true);
    newarea.value= area.value;
    area.parentNode.replaceChild(newarea, area);
  }
}


function cleanMessage() {
  var elem = document.getElementById("t1_text_<%Sendb(DefaultLanguageID)%>");
  
  if (elem != null) {
    elem.value = elem.value.replace("<", "\1");
  }
  
  return true;
}

function checkMessage() {
  var t = 1;
  var elem = document.getElementById("t" + t + "_text");
  var suppressSpan = document.getElementById("suppressBal");
  var myValue = '';
  var bTagFound = false;
  
  while (elem != null && !bTagFound) {
    myValue = elem.value;  
	  if (myValue.indexOf("|TOTALPOINTS|")>-1 || myValue.indexOf("|PTSASPEN|")>-1 
	      || myValue.indexOf("|ACCUMAMT|")>-1 || myValue.indexOf("|SVBAL")>-1
	      || myValue.indexOf("|SVREDEEM") >-1 || myValue.indexOf("|SVVAL")>-1
	      || myValue.indexOf("|PTBAL")>-1 || myValue.indexOf("|PTVAL")>-1
	      || myValue.indexOf("|NET#")>-1 || myValue.indexOf("|INITIAL#")>-1
	      || myValue.indexOf("|EARNED#")>-1 || myValue.indexOf("|REDEEMED#")>-1) {

      bTagFound = true;
    } else {
      bTagFound = false;
    }        
    t++; 
    elem = document.getElementById("t" + t + "_text");
  }
  
  if (suppressSpan != null) {
    suppressSpan.style.visibility = (bTagFound) ? 'visible' : 'hidden';
  }
}

function lineCounter() {
  var activePrinter = document.getElementById("printerselect").value;
  var maxLines = 0;
  var maxPerLine = 0;
  var strTemp = "";
  var strLineCounter = 1;
  var strCharCounter = 0;
  var counter = document.getElementById("counter");
  var theField = document.getElementById("t1_text_<%Sendb(DefaultLanguageID)%>");
  var theLineCounter = document.getElementById("remLines");
  
  if ((activePrinter > 0) && (activePrinter < 999)) {
    maxLines = document.getElementById("PT" + activePrinter + "lines").value;
    maxPerLine = document.getElementById("PT" + activePrinter).value;
  }
  if (maxLines == 0) {
    theLineCounter.value = '';
    theLineCounter.style.backgroundColor = '#ccffcc';
    counter.style.backgroundColor = '#ccffcc';
    counter.style.display = 'none';
  } else {
    for (var i = 0; i < theField.value.length; i++) {
      var strChar = theField.value.substring(i, i + 1);
      if (strChar == '\n') {
        strTemp += strChar;
        strCharCounter = 1;
        strLineCounter += 1;
      } else if (strCharCounter == maxPerLine) {
        strTemp += '\n' + strChar;
        strCharCounter = 1;
        strLineCounter += 1;
      } else {
        strTemp += strChar;
        strCharCounter ++;
      }
    }
    theLineCounter.value = maxLines - strLineCounter;
    if (theLineCounter.value < 0) {
      theLineCounter.style.backgroundColor = '#ffcccc';
      counter.style.backgroundColor = '#ffcccc';
      counter.style.display = 'block';
    } else {
      theLineCounter.style.backgroundColor = '#ccffcc';
      counter.style.backgroundColor = '#ccffcc';
      counter.style.display = 'block';
    }
  }
}

function captureCursorPosition() {
  if (bCapture && srcElement != null) {
    caretPos = getCaret(srcElement);
  }
}

function setCursorPosition() {
  if (isIE()) {
    var elem = srcElement;
    if (elem != null) {
      if (elem.createTextRange) {
        var range = elem.createTextRange();
        range.move('character', caretPos);
        range.select();
      } else {
        if (elem.selectionStart) {
          elem.focus();
          elem.setSelectionRange(caretPos, caretPos);
        }
        else
        elem.focus();
      }
    }
    bCapture = true;
  }
}

//function getCaret(el) {
//  if (document.selection) {
//    el.focus();
//    
//    var r = document.selection.createRange();
//    if (r == null) {
//      return 0;
//    }
//    
//    var re = el.createTextRange(),
//        rc = re.duplicate();
//    re.moveToBookmark(r.getBookmark());
//    rc.setEndPoint('EndToStart', re);
//    
//    return rc.text.length;
//  }
//  return 0;
//}

function getCaret(el) {
  var CaretPos = 0;
  // IE Support
  if (document.selection) {
    el.focus();
    var Sel = document.selection.createRange();
    var Sel2 = Sel.duplicate();
    Sel2.moveToElementText(el);
    var CaretPos = 0;
    var CharactersAdded = 1;
    while (Sel2.inRange(Sel)) {
      //old GetCaretPosition always counts 1 for linetermination
      if (Sel2.htmlText.substr(0, 2) == "\r\n") {
        CaretPos += 2;
        CharactersAdded = 2;
      } else {
        CaretPos++;
        CharactersAdded = 1;
      }
      Sel2.moveStart('character');
    }
    CaretPos -= CharactersAdded;
  }
  // Firefox support
  else if (el.selectionStart || el.selectionStart == '0')
    CaretPos = el.selectionStart;
  return (CaretPos);
}

function validateForm() {
  var valid = false;
  var elem = document.getElementById('SortID');
  var sortID = 0;

  if (elem != null) {
    if (trimString(elem.value)=='') {
      elem.value = '0';
    }

    if (!isNaN(elem.value) && parseInt(elem.value) >= 0 && parseInt(elem.value) <= 999) {
      valid = true;
    } else {  
      alert('<%Sendb(Copient.PhraseLib.Lookup("reward.invalid-sortid", LanguageID))%>');  
    }
  }
  
  return valid;
}

function handleOnSubmit() {
  var shouldSubmit = false;
  
  if (validateForm()) {
    shouldSubmit = cleanMessage();
  }

  return shouldSubmit;
}

function getInternetExplorerVersion()
// Returns the version of Internet Explorer or a -1
// (indicating the use of another browser).
{
  var rv = -1; // Return value assumes failure.
  if (navigator.appName == 'Microsoft Internet Explorer')
  {
    var ua = navigator.userAgent;
    var re  = new RegExp("MSIE ([0-9]{1,}[\.0-9]{0,})");
    if (re.exec(ua) != null)
      rv = parseFloat( RegExp.$1 );
  }
  return rv;
}
</script>
<%
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
    Send("<script type=""text/javascript"">")
    Send("  function ChangeParentDocument() { return true; } ")
    Send("</script>")
    Send_Denied(1, "banners.access-denied-offer")
    Send_BodyEnd()
    GoTo done
  End If
  
  TextAreaBodyText = BodyText
  If (TextAreaBodyText <> "" AndAlso TextAreaBodyText.IndexOf("<") > -1) Then
    TextAreaBodyText = TextAreaBodyText.Replace("<", "&lt;")
  End If
  If (TextAreaBodyText <> "" AndAlso TextAreaBodyText.IndexOf(">") > -1) Then
    TextAreaBodyText = TextAreaBodyText.Replace(">", "&gt;")
  End If
%>
<form action="UEoffer-rew-pmsg.aspx" id="mainform" name="mainform" method="post"
onsubmit="return handleOnSubmit();">
<input type="hidden" id="OfferID" name="OfferID" value="<% Sendb(OfferID)%>" />
<input type="hidden" id="RewardID" name="RewardID" value="<% Sendb(RewardID)%>" />
<input type="hidden" id="DeliverableID" name="DeliverableID" value="<% sendb(DeliverableID) %>" />
<input type="hidden" id="MessageID" name="MessageID" value="<% Sendb(MessageID)%>" />
<input type="hidden" id="Phase" name="Phase" value="<% Sendb(Phase)%>" />
<input type="hidden" id="roid" name="roid" value="<%Sendb(TpROID) %>" />
<input type="hidden" id="tp" name="tp" value="<%Sendb(TouchPoint) %>" />
<input type="hidden" id="IsTemplate" name="IsTemplate" value="<% Sendb(IsTemplateVal)%>" />
<input type="hidden" id="EngineID" name="EngineID" value="<%Sendb(EngineID) %>" />
<input type="hidden" id="TierLevels" name="TierLevels" value="<%Sendb(TierLevels) %>" />
<div id="intro">
  <%
    If (IsTemplate) Then
      Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup(PhaseTitle, LanguageID), VbStrConv.Lowercase) & "</h1>")
    Else
      Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup(PhaseTitle, LanguageID), VbStrConv.Lowercase) & "</h1>")
    End If
  %>
  <div id="controls">
    <% If (IsTemplate) Then%>
    <span class="temp">
      <input type="checkbox" class="tempcheck" id="Disallow_Edit" name="Disallow_Edit"
        <% if(disallow_edit)then sendb(" checked=""checked""") %> />
      <label for="Disallow_Edit">
        <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
      </label>
    </span>
    <% End If%>
    <button class="regular" id="preview" name="preview" type="button" onclick="javascript:doPreviewPopup()">
      <% Sendb(Copient.PhraseLib.Lookup("term.preview", LanguageID))%>
    </button>
    <% 
        m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
      If ((Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
        If Not IsTemplate Then
          If (TierLevels > 1 And SupportGlobalAndTieredConditions = 1) Then
                    If (Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And Disallow_Edit) And Not IsOfferWaitingForApproval(OfferID)) Then Send_Save("onclick=""enableTiers(" & TierLevels & ", " & MLLanguageID & ");""")
          Else
                    If (Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And Disallow_Edit) And Not IsOfferWaitingForApproval(OfferID)) Then Send_Save()
          End If
        Else
          If (TierLevels > 1 And SupportGlobalAndTieredConditions = 1) Then
            If (Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer) Then Send_Save("onclick=""enableTiers(" & TierLevels & ", " & MLLanguageID & ");""")
          Else
            If (Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer) Then Send_Save()
          End If
        End If
      End If
    %>
  </div>
</div>
<div id="main">
  <%
    If (infoMessage <> "") Then
      Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
    End If
  %>
  <div id="column2x">
    <div class="box" id="message">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.message", LanguageID))%>
        </span>
      </h2>
      <div style="height: 460px; overflow-y: auto;">
        <select id="type" name="type" <% sendb(disabledattribute) %>>
          <%
            MyCommon.QueryStr = "select TypeID, PhraseID from PrintedMessageTypes with (NoLock) where EngineID = 9 order by TypeID"
            rst = MyCommon.LRT_Select()
            For Each row In rst.Rows
              Sendb("<option value=""" & row.Item("TypeID") & """")
              If MessageTypeID = row.Item("TypeID") Then
                Sendb(" selected=""selected""")
              End If
              Send(">")
              Send(Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID))
              Send("</option>")
            Next
          %>
          <!--<option value="5"<% if (messagetypeid=5) then sendb(" selected=""selected""") %>>
            <% Sendb(Copient.PhraseLib.Lookup("term.generalreceipt", LanguageID))%>
          </option> -->
          <!-- Commented out per Adam Rausch's request 5/17
          (If these ever get restored, be sure to use the PhraseID value from the PrintedMessageTypes table rather than the Description)
          <option value="2"<% if (MessageTypeID=2) then Sendb(" selected=""selected""") %>><% Sendb(Copient.PhraseLib.Lookup("term.summary", LanguageID))%></option>
          <option value="3"<% if (MessageTypeID=3) then Sendb(" selected=""selected""") %>><% Sendb(Copient.PhraseLib.Lookup("term.aftertrailer", LanguageID))%></option>
          <option value="4"<% if (MessageTypeID=4) then Sendb(" selected=""selected""") %>><% Sendb(Copient.PhraseLib.Lookup("term.sweepstakes", LanguageID))%></option>
          -->
        </select>
        &nbsp;&nbsp;
        <label for="SortID">
          <%Sendb(Copient.PhraseLib.Lookup("term.sortid", LanguageID))%>:</label>
        <input type="text" class="shortest" name="SortID" id="SortID" value="<%Sendb(SortID)%>"
          maxlength="3" />
        <%
          ' check to see if the suppress zero balances should be displayed by looking for the presence of specific tags
          MyCommon.QueryStr = "Select count(*) as TagCount from PrintedMessageTiers with (NoLock) " & _
                              "where MessageID=" & MessageID & " and (BodyText like'%|TOTALPOINTS|%' " & _
                              " or BodyText like '%|PTSASPEN|%' or BodyText like'%|ACCUMAMT|%' " & _
                              " or BodyText like '%|SVBAL%' or BodyText like '%|SVREDEEM%' " & _
                              " or BodyText like '%|SVVAL%' or BodyText like '%PTBAL%' " & _
                              " or BodyText like '%PTVAL%' or BodyText like '%NET#%' " & _
                              " or BodyText like '%INITIAL#%' or BodyText like '%EARNED#%' " & _
                              " or BodyText like '%REDEEMED#%');"
          rst2 = MyCommon.LRT_Select
          If rst2.Rows.Count > 0 Then
            ShowSuppress = (MyCommon.NZ(rst2.Rows(0).Item("TagCount"), 0) > 0)
          End If
        %>
        <span id="suppressBal" style="visibility: <%Sendb(IIf(ShowSuppress, "visible", "hidden"))%>;">
          &nbsp;&nbsp;
          <input type="checkbox" name="suppressZeroBal" id="suppressZeroBal" value="1" <%Sendb(IIf(SuppressZeroBal OrElse (DefaultSuppressed AndAlso MessageID=0), " checked=""checked""", ""))%> />
          <label for="suppressZeroBal">
            <%Sendb(Copient.PhraseLib.Lookup("CPE_Reward.suppress-zero-balance", LanguageID))%></label>
        </span>
        <br />
        <br class="half" />
        <%
          Dim LanguageIDsList As String = ""
          Dim TierRecordDT As DataTable
          Dim CustomerFacingLangID As Integer = 1
          Integer.TryParse(MyCommon.Fetch_SystemOption(125), CustomerFacingLangID)
          For t = 1 To TierLevels
            MyCommon.QueryStr = "Select PKID, TierLevel, BodyText from PrintedMessageTiers with (NoLock) " & _
                                "where MessageID=" & MessageID & " and TierLevel=" & t & ";"
            TierRecordDT = MyCommon.LRT_Select
            If TierRecordDT.Rows.Count > 0 Then
              PKID = MyCommon.NZ(TierRecordDT.Rows(0).Item("PKID"), 0)
            Else
              PKID = 0
            End If
            
            l = 1
            MyCommon.QueryStr = "SELECT L.LanguageID, L.Name, L.MSNetCode, L.JavaLocaleCode, L.PhraseTerm, L.RightToLeftText, T.BodyText " & _
                                "FROM Languages AS L " & _
                                "LEFT JOIN PMTranslations AS T ON T.LanguageID=L.LanguageID AND T.PMTiersID=" & PKID & " " & _
                                "WHERE L.LanguageID in (" & IIf(MultiLanguageEnabled, "SELECT TLV.LanguageID FROM TransLanguagesCF_UE AS TLV", DefaultLanguageID) & ") " & _
                                "ORDER BY CASE WHEN L.LanguageID=" & DefaultLanguageID & " THEN 1 ELSE 2 END, L.LanguageID;"
            LanguagesDT = MyCommon.LRT_Select
            For Each row In LanguagesDT.Rows
              If t = 1 Then
                LanguageIDsList &= IIf(LanguageIDsList = "", "", ",") & row.Item("LanguageID")
              End If
              
              If (MultiLanguageEnabled = True) Or (MultiLanguageEnabled = False AndAlso MyCommon.NZ(row.Item("LanguageID"), 0) = DefaultLanguageID) Then
                Dim MLLanguageCode As String = MyCommon.NZ(row.Item("MSNetCode"), "")
                Dim MLLanguageName As String = Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseTerm"), ""), MyCommon.GetAdminUser.LanguageID)
                MLLanguageID = MyCommon.NZ(row.Item("LanguageID"), 0)
                
                Sendb("<label for=""t" & t & "_text_" & MLLanguageID & """")
                If t > 1 Then
                  Sendb(" style=""color:#" & IIf(t Mod 2 = 0, "009900", "000099") & """")
                End If
                Sendb(">")
                If MultiLanguageEnabled Then
                  Sendb(MLLanguageName & IIf(MLLanguageID = CustomerFacingLangID, " (" & Copient.PhraseLib.Lookup("term.default", MyCommon.GetAdminUser.LanguageID) & ")", ""))
                End If
                If TierLevels > 1 Then
                  Sendb("<b>" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & ":</b> ")

                If (SupportGlobalAndTieredConditions = 1 and t = 1) Then
                  Send("  <td><input type=""checkbox"" name=""useSameTierValue"" id =""useSameTierValue"" value=""1""" & IIf(UseSameTierValue = 1, " checked=""checked""", "") & " style=""margin-left:7px;"" onclick=""setSameTierValue(" & TierLevels & ", " & MLLanguageID & ")""/>")
                  Sendb("  <label for=""useThisValueForAllTiers"">" & Copient.PhraseLib.Lookup("term.UseThisValueForAllTiers", LanguageID) & "</label></td>") 
                End If
                End If
                Send("</label><br />")
                Send("<div class=""pmsgwrap"">")
                Sendb("  <textarea class=""CPEpmsg"" id=""t" & t & "_text_" & MLLanguageID & """ name=""t" & t & "_text_" & MLLanguageID & """")
                Sendb(" cols=""38"" rows=""" & IIf(TierLevels = 1, "24", "8") & """" & IIf(UseSameTierValue = 1, " disabled=""disabled""", DisabledAttribute))
                Sendb(IIf(MyCommon.NZ(row.Item("RightToLeftText"), False), " dir=""rtl""", ""))
                Sendb(IIf(MyCommon.NZ(row.Item("MSNetCode"), "") <> "", " lang=""" & row.Item("MSNetCode"), "") & """")
                Sendb("wrap=""hard"" ")
                Sendb("onfocus=""srcElement=this;setCursorPosition();"" ")
                Sendb("onKeyUp=""lineCounter()"" ")
                Sendb("onPaste=""lineCounter()"" ")
                Sendb("onCut=""lineCounter()"" ")
                Sendb("onblur=""javascript:bCapture=false;"" ")
                Sendb(">")
                        Dim TextAreaBodyText1 = GetCgiValue("t" & t & "_text_" & DefaultLanguageID)
                        If TextAreaBodyText1 IsNot Nothing Then
                            Sendb(Server.HtmlEncode(TextAreaBodyText1))
                        ElseIf TierRecordDT.Rows.Count > 0 Then
                            If l = 1 AndAlso IsDBNull(row.Item("BodyText")) Then
                                Sendb(Server.HtmlEncode(MyCommon.NZ(TierRecordDT.Rows(0).Item("BodyText"), "")))
                            Else
                                Sendb(Server.HtmlEncode(MyCommon.NZ(row.Item("BodyText"), "")))
                            End If
                        End If
                Send("</textarea>")
                Send("<br />")
                Send("<div id=""counter"" style=""display:none;width:175px;padding:3px;"">")
                Send("  <input type=""text"" id=""remLines"" name=""remLines"" value="""" style=""border-width:0px;width:25px;"" readonly /> " & Copient.PhraseLib.Lookup("term.LinesLeft", LanguageID) & "<br />")
                Send("</div>")
                Send("</div>")
              End If
              l += 1
            Next
            If MultiLanguageEnabled And TierLevels > 1 And t < TierLevels Then
              Send("<hr />")
            End If
          Next
          Send("<input type=""hidden"" id=""LanguageIDs"" name=""LanguageIDs"" value=""" & LanguageIDsList & """ />")
        %>
        <hr class="hidden" />
      </div>
      <!-- end scroll div -->
    </div>
    <!-- end box -->
    <div id="debug">
    </div>
  </div>
  <div id="gutter">
  </div>
  <div id="column1x">
    <div class="box" id="printer">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.printer", LanguageID))%>
        </span>
      </h2>
      <label for="printerselect">
        <% Sendb(Copient.PhraseLib.Lookup("reward.selectprinter", LanguageID) & ":")%>
      </label>
      <br />
      <select id="printerselect" name="printerselect" style="width:100%" onchange="xmlhttpPost('/logix/PrintedMessageFeeds.aspx');">
        <%
          MyCommon.QueryStr = "select PTy.PrinterTypeID,PTy.PageWidth,PTy.Name,PTy.PhraseID,PTy.Installed,PTy.DefaultPrinter,PTy.MaxLines,PEPT.EngineID " & _
                              "from PrinterTypes as PTy with (NoLock) " & _
                              "inner join PromoEnginePrinterTypes as PEPT with (NoLock) on PEPT.PrinterTypeID=PTy.PrinterTypeID " & _
                              "where PTy.Installed=1 and PEPT.EngineID=" & EngineID & " order by PTy.Name"
          rst2 = MyCommon.LRT_Select
          Send("<option value=""999"">" & Copient.PhraseLib.Lookup("term.allprinters", LanguageID) & "</option>")
          For Each row In rst2.Rows
            If (MyCommon.NZ(row.Item("PrinterTypeID"), 0) > 0) Then
              PrinterWidthBuf.Append("<input type=""hidden"" id=""PT" & row.Item("PrinterTypeID") & """ name=""PT" & row.Item("PrinterTypeID") & """ value=""" & IIf(row.Item("PageWidth") <= 0, 52, row.Item("PageWidth")) & """ />" & vbCrLf)
              PrinterWidthBuf.Append("<input type=""hidden"" id=""PT" & row.Item("PrinterTypeID") & "lines"" name=""PT" & row.Item("PrinterTypeID") & "lines"" value=""" & MyCommon.NZ(row.Item("MaxLines"), 0) & """ />" & vbCrLf)
              Sendb("<option value=""" & row.Item("PrinterTypeID") & """")
              If MyCommon.NZ(row.Item("DefaultPrinter"), False) Or rst2.Rows.Count = 1 Then
                Sendb(" selected=""selected""")
              End If
              If MyCommon.NZ(row.Item("PhraseID"), "0") <> "0" Then
                PrinterPhrase = Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID)
              Else
                PrinterPhrase = MyCommon.NZ(row.Item("Name"), "")
              End If
              Send(">" & PrinterPhrase & "</option>")
            Else
            End If
          Next
        %>
      </select>
      <% Send(PrinterWidthBuf.ToString())%>
      <hr class="hidden" />
    </div>
    <div class="box" id="tags">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.tags", LanguageID))%>
        </span>
      </h2>
      <br class="half" />
      <div id="ed_toolbar" style="background-color: #d0d0d0; text-align: center;">
        <div id="tools">
        </div>
      </div>
      <br />
      <% Sendb(Copient.PhraseLib.Lookup("reward.tagnotes", LanguageID))%>
      <hr class="hidden" />
    </div>
    <div class="box" id="distribution">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.distribution", LanguageID))%>
        </span>
      </h2>
      <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.distribution", LanguageID))%>">
        <tr>
          <td colspan="2">
            <input type="checkbox" id="requiredToDeliver" name="requiredToDeliver" value="1"
              <%Sendb(IIf(RewardRequired, " checked=""checked""", ""))%> />
            <label for="requiredToDeliver">
              <% Sendb(Copient.PhraseLib.Lookup("ue-reward.reward-required", LanguageID))%></label>
          </td>
        </tr>
        <%'BZ2079: UE-feature-removal #: reward nominal value textbox removed.
          '        To restore previous functionality: remove the style="display:none" attribute from the tr tag below.
        %>
        <tr style="display: none;">
          <td>
            <label for="t1_value">
              <% Sendb(Copient.PhraseLib.Lookup("term.value", LanguageID))%>:</label>
          </td>
          <td>
            <%
              If DeliverableID = 0 Then
                For t = 1 To TierLevels
                  If Request.QueryString("t" & t & "_value") <> "" Then
                    If TierLevels > 1 Then
                      Send("<label for=""t" & t & "_value"">" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & ":</label>")
                    End If
                    Sendb("<input type=""text"" class=""shorter"" id=""t" & t & "_value"" name=""t" & t & "_value"" maxlength=""9"" value=""" & MyCommon.Extract_Val(Request.QueryString("t" & t & "_value")) & """/>")
                  Else
                    If TierLevels > 1 Then
                      Send("<label for=""t" & t & "_value"">" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & ":</label>")
                    End If
                    Sendb("<input type=""text"" class=""shorter"" id=""t" & t & "_value"" name=""t" & t & "_value"" maxlength=""9"" value=""0""/>")
                  End If
                  Send("<br />")
                Next
              Else
                MyCommon.QueryStr = "Select Value from PrintedMessageTiers with (NoLock) " & _
                                    "where MessageID=" & MessageID & " order by TierLevel;"
                rst = MyCommon.LRT_Select
                For t = 1 To TierLevels
                  If TierLevels > 1 Then
                    Send("<label for=""t" & t & "_value"">" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & ":</label>")
                  End If
                  Sendb("<input type=""text"" class=""shorter"" id=""t" & t & "_value"" name=""t" & t & "_value"" maxlength=""9""")
                  If t > rst.Rows.Count Then
                    Sendb(" value=""0""" & DisabledAttribute & " />")
                  Else
                    Sendb(" value=""" & MyCommon.NZ(rst.Rows(t - 1).Item("Value"), 0) & """" & DisabledAttribute & " />")
                  End If
                  Send("<br />")
                Next
              End If
            %>
          </td>
        </tr>
      </table>
    </div>
  </div>
</div>
<div id="dialogbox" style="height: 315px;">
  <div id="pointselector" style="display: none;">
    <div id="ptTag">
    </div>
    <br />
    <input type="hidden" name="ptTagName" id="ptTagName" value="" />
    <b>
      <% Sendb(Copient.PhraseLib.Lookup("offer-rew.selectpoints", LanguageID) & ":")%>
    </b>
    <br />
    <input type="radio" id="functionradio2a" name="functionradio2" checked="checked" /><label
      for="functionradio2a"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
    <input type="radio" id="functionradio2b" name="functionradio2" /><label for="functionradio2b"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
    <input type="text" class="medium" id="functioninput2" name="functioninput2" onkeyup="handleKeyUp(200, 2);"
      value="" /><br />
    <select onclick="handleSelectClick(2);" id="functionselect2" name="functionselect2"
      size="10" style="width: 220px;">
      <%
        If (isAnyCustomer) Then
          MyCommon.QueryStr = "Select PP.ProgramID, PP.ProgramName from PointsPrograms PP inner join dbo.PointsProgramsPromoEngineSettings as PEPP  with (NoLock) on PP.ProgramID =PEPP.ProgramID where Deleted=0 And AllowAnyCustomer=1 order by ProgramName;"
        Else
          MyCommon.QueryStr = "Select ProgramID, ProgramName  from PointsPrograms  with (NoLock) where Deleted=0 order by ProgramName;"
        End If
        rst = MyCommon.LRT_Select
        For Each row In rst.Rows
          Send("<option value=""" & MyCommon.NZ(row.Item("ProgramID"), 0) & """>" & MyCommon.NZ(row.Item("ProgramName"), "") & "</option>")
        Next
      %>
    </select>
    <br />
    <br />
    <input type="button" id="close2" name="close2" value="<% Sendb(Copient.PhraseLib.Lookup("term.close", LanguageID))%>"
      onclick="javascript:showDialogSpan(false, 2);" />
  </div>
  <div id="svselector" style="display: none;">
    <div id="svTag">
    </div>
    <br />
    <input type="hidden" name="svTagName" id="svTagName" value="" />
    <b>
      <% Sendb(Copient.PhraseLib.Lookup("offer-rew.selectsv", LanguageID) & ":")%>
    </b>
    <br />
    <input type="radio" id="functionradioa" name="functionradio" checked="checked" /><label
      for="functionradio"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
    <input type="radio" id="functionradiob" name="functionradio" /><label for="functionradio"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
    <input type="text" class="medium" id="functioninput" name="functioninput" onkeyup="handleKeyUp(200, 1);"
      value="" /><br />
    <select onclick="handleSelectClick(3);" id="functionselect" name="functionselect"
      size="10" style="width: 220px;">
      <%
        If (isAnyCustomer) Then
          MyCommon.QueryStr = "Select SVP.SVProgramID, SVP.Name from StoredValuePrograms SVP inner join dbo.SVProgramsPromoEngineSettings as PESVP  with (NoLock) on SVP.SVProgramID =PESVP.SVProgramID where Deleted=0 And AllowAnyCustomer=1 order by Name;"
        Else
          MyCommon.QueryStr = "Select SVProgramID, Name  from StoredValuePrograms with (NoLock) where Deleted=0 order by Name;"
        End If
        
        rst = MyCommon.LRT_Select
        For Each row In rst.Rows
          Send("<option value=""" & MyCommon.NZ(row.Item("SVProgramID"), 0) & """>" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
        Next
      %>
    </select>
    <br />
    <br />
    <input type="button" id="close" name="close" value="<% Sendb(Copient.PhraseLib.Lookup("term.close", LanguageID))%>"
      onclick="javascript:showDialogSpan(false, 3);" />
  </div>
  <div id="scselector" style="display: none;">
    <div id="scTag">
    </div>
    <br />
    <input type="hidden" name="scTagName" id="scTagName" value="Scorecard" />
    <b>
      <% Sendb(Copient.PhraseLib.Lookup("offer-rew.selectsc", LanguageID) & ":")%>
    </b>
    <br />
    <select onclick="handleSelectClick(4);" id="functionselect4" name="functionselect4"
      size="10" style="width: 220px;">
      <%
        MyCommon.QueryStr = "Select ScorecardID, Description from Scorecards where Deleted=0 and EngineID=" & EngineID & " and ScorecardTypeID=1 order by Description;"
        rst = MyCommon.LRT_Select
        For Each row In rst.Rows
          Send("<option value=""" & MyCommon.NZ(row.Item("ScorecardID"), 0) & """>" & MyCommon.NZ(row.Item("Description"), "") & "</option>")
        Next
      %>
    </select>
    <br />
    <br />
    <input type="button" id="close4" name="close4" value="<% Sendb(Copient.PhraseLib.Lookup("term.close", LanguageID))%>"
      onclick="javascript:showDialogSpan(false, 4);" />
  </div>
  <div id="svscselector" style="display: none;">
    <div id="svscTag">
    </div>
    <br />
    <input type="hidden" name="svscTagName" id="svscTagName" value="Svscorecard" />
    <b>
      <% Sendb(Copient.PhraseLib.Lookup("offer-rew.selectsc", LanguageID) & ":")%>
    </b>
    <br />
    <select onclick="handleSelectClick(6);" id="functionselect6" name="functionselect6"
      size="10" style="width: 220px;">
      <%
        MyCommon.QueryStr = "Select ScorecardID, Description from Scorecards where Deleted=0 and EngineID=" & EngineID & " and ScorecardTypeID=2 order by Description;"
        rst = MyCommon.LRT_Select
        For Each row In rst.Rows
          Send("<option value=""" & MyCommon.NZ(row.Item("ScorecardID"), 0) & """>" & MyCommon.NZ(row.Item("Description"), "") & "</option>")
        Next
      %>
    </select>
    <br />
    <br />
    <input type="button" id="close6" name="close6" value="<% Sendb(Copient.PhraseLib.Lookup("term.close", LanguageID))%>"
      onclick="javascript:showDialogSpan(false, 6);" />
  </div>
  <div id="dscselector" style="display: none;">
    <div id="dscTag">
    </div>
    <br />
    <input type="hidden" name="dscTagName" id="dscTagName" value="Dscorecard" />
    <b>
      <% Sendb(Copient.PhraseLib.Lookup("offer-rew.selectsc", LanguageID) & ":")%>
    </b>
    <br />
    <select onclick="handleSelectClick(7);" id="functionselect7" name="functionselect7"
      size="10" style="width: 220px;">
      <%
        MyCommon.QueryStr = "Select ScorecardID, Description from Scorecards where Deleted=0 and EngineID=" & EngineID & " and ScorecardTypeID=3 order by Description;"
        rst = MyCommon.LRT_Select
        For Each row In rst.Rows
          Send("<option value=""" & MyCommon.NZ(row.Item("ScorecardID"), 0) & """>" & MyCommon.NZ(row.Item("Description"), "") & "</option>")
        Next
      %>
    </select>
    <br />
    <br />
    <input type="button" id="close7" name="close7" value="<% Sendb(Copient.PhraseLib.Lookup("term.close", LanguageID))%>"
      onclick="javascript:showDialogSpan(false, 7);" />
  </div>
  <div id="rrValues" style="display: none;">
    <div id="rrTag">
    </div>
    <br />
    <input type="hidden" name="rrTagName" id="rrTagName" value="Svratio" />
    <b>
      <% Sendb(Copient.PhraseLib.Lookup("tag.sv-ratio1", LanguageID))%>
    </b>
    <br />
    <br />
    <br />
    <label for="incrementqty">
      <% Sendb(Copient.PhraseLib.Lookup("tag.sv-ratio2", LanguageID) & ": ")%>
    </label>
    <input type="text" class="shortest" id="incrementqty" name="incrementqty" maxlength="4"
      style="width: 35px;" />
    <br />
    <br />
    <br />
    <label for="incrementamt">
      <% Sendb(Copient.PhraseLib.Lookup("tag.sv-ratio3", LanguageID) & ": ")%>
    </label>
    <input type="text" class="shortest" id="incrementamt" name="incrementamt" maxlength="5"
      style="width: 40px;" />
    <br />
    <br />
    <br />
    <input type="button" id="continue1" name="continue1" value="<% Sendb(Copient.PhraseLib.Lookup("term.continue", LanguageID))%>"
      onclick="javascript:showNextSelect();" />
    <input type="button" id="close3" name="close3" value="<% Sendb(Copient.PhraseLib.Lookup("term.close", LanguageID))%>"
      onclick="javascript:showDialogSpan(false, 5);" />
  </div>
  <div id="monoparametric" style="display: none;">
    <div id="monoparametricTag">
    </div>
    <br />
    <input type="hidden" name="monoparametricTagName" id="monoparametricTagName" value="" />
    <label id="monoparametricLabel1" for="monoparameter1">
    </label>
    <br />
    <input type="text" class="medium" id="monoparameter1" name="monoparameter1" value="" /><br />
    <br />
    <input type="button" id="submit11" name="submit11" value="<% Sendb(Copient.PhraseLib.Lookup("term.submit", LanguageID))%>"
      onclick="handleSelectClick(11);" />
    <input type="button" id="close11" name="close11" value="<% Sendb(Copient.PhraseLib.Lookup("term.close", LanguageID))%>"
      onclick="javascript:showDialogSpan(false, 11);" />
  </div>
  <div id="diparametric" style="display: none;">
    <div id="diparametricTag">
    </div>
    <br />
    <input type="hidden" name="diparametricTagName" id="diparametricTagName" value="" />
    <label id="diparametricLabel1" for="diparameter1">
    </label>
    <br />
    <input type="text" class="medium" id="diparameter1" name="diparameter1" value="" /><br />
    <br class="half" />
    <label id="diparametricLabel2" for="diparameter2">
    </label>
    <br />
    <input type="text" class="medium" id="diparameter2" name="diparameter2" value="" /><br />
    <br />
    <input type="button" id="submit21" name="submit21" value="<% Sendb(Copient.PhraseLib.Lookup("term.submit", LanguageID))%>"
      onclick="handleSelectClick(21);" />
    <input type="button" id="close21" name="close21" value="<% Sendb(Copient.PhraseLib.Lookup("term.close", LanguageID))%>"
      onclick="javascript:showDialogSpan(false, 21);" />
  </div>
  <div id="triparametric" style="display: none;">
    <div id="triparametricTag">
    </div>
    <br />
    <input type="hidden" name="triparametricTagName" id="triparametricTagName" value="" />
    <label id="triparametricLabel1" for="triparameter1">
    </label>
    <br />
    <input type="text" class="medium" id="triparameter1" name="triparameter1" value="" /><br />
    <br class="half" />
    <label id="triparametricLabel2" for="triparameter2">
    </label>
    <br />
    <input type="text" class="medium" id="triparameter2" name="triparameter2" value="" /><br />
    <br class="half" />
    <label id="triparametricLabel3" for="triparameter3">
    </label>
    <br />
    <input type="text" class="medium" id="triparameter3" name="triparameter3" value="" /><br />
    <br />
    <input type="button" id="submit31" name="submit31" value="<% Sendb(Copient.PhraseLib.Lookup("term.submit", LanguageID))%>"
      onclick="handleSelectClick(31);" />
    <input type="button" id="close31" name="close31" value="<% Sendb(Copient.PhraseLib.Lookup("term.close", LanguageID))%>"
      onclick="javascript:showDialogSpan(false, 31);" />
  </div>
  <div id="tetraparametric" style="display: none;">
    <div id="tetraparametricTag"></div>
    <br />
    <input type="hidden" name="tetraparametricTagName" id="tetraparametricTagName" value="" />
    <label id="tetraparametricLabel1" for="tetraparameter1">
    </label>
    <br />
    <input type="text" class="medium" id="tetraparameter1" name="tetraparameter1" value="" /><br />
    <br class="half" />
    <label id="tetraparametricLabel2" for="tetraparameter2">
    </label>
    <br />
    <input type="text" class="medium" id="tetraparameter2" name="tetraparameter2" value="" /><br />
    <br class="half" />
    <label id="tetraparametricLabel3" for="tetraparameter3">
    </label>
    <br />
    <input type="text" class="medium" id="tetraparameter3" name="tetraparameter3" value="" /><br />
    <br class="half" />
    <label id="tetraparametricLabel4" for="tetraparameter4">
    </label>
    <br />
    <input type="text" class="medium" id="tetraparameter4" name="tetraparameter4" value="" /><br />
    <br />
    <input type="button" id="submit41" name="submit41" value="<% Sendb(Copient.PhraseLib.Lookup("term.submit", LanguageID))%>"
      onclick="handleSelectClick(41);" />
    <input type="button" id="close41" name="close41" value="<% Sendb(Copient.PhraseLib.Lookup("term.close", LanguageID))%>"
      onclick="javascript:showDialogSpan(false, 41);" />
  </div>
</div>
</div>
</form>
<script runat="server">
  Function CreateMessage(ByVal MessageID As Long, ByVal MessageTypeID As String, ByVal SuppressZeroBal As Boolean, ByVal SortID As Integer) As Long
    Dim MyCommon As New Copient.CommonInc
    Try
      MyCommon.Open_LogixRT()
      ' Create a message if one doesn't already exist
      MyCommon.QueryStr = "dbo.pt_PrintedMessages_Insert"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@MessageTypeID", SqlDbType.Int, 4).Value = MessageTypeID
      MyCommon.LRTsp.Parameters.Add("@SuppressZeroBalance", SqlDbType.Bit).Value = IIf(SuppressZeroBal, 1, 0)
      MyCommon.LRTsp.Parameters.Add("@SortID", SqlDbType.Int, 4).Value = SortID
      MyCommon.LRTsp.Parameters.Add("@MessageID", SqlDbType.BigInt, 8).Direction = ParameterDirection.Output
      MyCommon.LRTsp.ExecuteNonQuery()
      MessageID = MyCommon.LRTsp.Parameters("@MessageID").Value
      MyCommon.Close_LRTsp()
    Catch ex As Exception
      MessageID = -1
      Send(ex.ToString)
    Finally
      MyCommon.Close_LogixRT()
      MyCommon = Nothing
    End Try
    Return MessageID
  End Function
  
  Sub CreateMessageTiers(ByRef MyCommon As Copient.CommonInc, ByVal MessageID As Long, ByVal Tier As Integer, ByVal OpenTagEscape As String, ByVal DefaultLanguageID As Integer, ByVal RewardValue As Decimal)
    Dim BodyText As String = ""
    Dim Localization As Copient.Localization
    Dim MLI As New Copient.Localization.MultiLanguageRec
    Dim PKID As Integer = 0
    
    Localization = New Copient.Localization(MyCommon)
    
    BodyText = Request.QueryString("t" & Tier & "_text_" & DefaultLanguageID)
    If (BodyText = "") Then BodyText = Request.Form("t" & Tier & "_text_" & DefaultLanguageID)
    If (BodyText <> "" AndAlso BodyText.IndexOf(OpenTagEscape) > -1) Then
      BodyText = BodyText.Replace(OpenTagEscape, "<")
    End If
    If BodyText Is Nothing Then
      BodyText = ""
    End If
    
    If (MessageID > 0) Then
      If RewardValue < 0 Then RewardValue = 0
      
      MyCommon.QueryStr = "dbo.pt_PrintedMsgTiers_Update"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@MessageID", SqlDbType.BigInt, 8).Value = MessageID
      MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int, 4).Value = Tier
      MyCommon.LRTsp.Parameters.Add("@BodyText", SqlDbType.NVarChar, 4000).Value = BodyText
      MyCommon.LRTsp.Parameters.Add("@Value", SqlDbType.Decimal, 9).Value = RewardValue
      MyCommon.LRTsp.Parameters.Add("@PKID", SqlDbType.Int).Direction = ParameterDirection.Output
      MyCommon.LRTsp.ExecuteNonQuery()
      PKID = MyCommon.LRTsp.Parameters("@PKID").Value
      MyCommon.Close_LRTsp()
      
      'Save multilanguage values
      'If (MyCommon.Fetch_SystemOption(124) = "1") Then
        Dim LanguagesDT As DataTable
        Dim row As DataRow
        MyCommon.QueryStr = "SELECT LanguageID FROM Languages AS L " & _
                            "WHERE L.LanguageID in (SELECT TLV.LanguageID FROM TransLanguagesCF_UE AS TLV) " & _
                            "ORDER BY CASE WHEN L.LanguageID=" & DefaultLanguageID & " THEN 1 ELSE 2 END, L.LanguageID;"
        LanguagesDT = MyCommon.LRT_Select
        For Each row In LanguagesDT.Rows
          BodyText = Request.QueryString("t" & Tier & "_text_" & MyCommon.NZ(row.Item("LanguageID"), 0))
          If (BodyText = "") Then BodyText = Request.Form("t" & Tier & "_text_" & MyCommon.NZ(row.Item("LanguageID"), 0))
          If (BodyText <> "" AndAlso BodyText.IndexOf(OpenTagEscape) > -1) Then
            BodyText = BodyText.Replace(OpenTagEscape, "<")
          End If
          If BodyText Is Nothing Then
            BodyText = ""
          End If
          If BodyText <> "" Then
            MyCommon.QueryStr = "INSERT INTO PMTranslations (PMTiersID, LanguageID, BodyText) " & _
                                "VALUES (" & PKID & ", " & row.Item("LanguageID") & ", N'" & MyCommon.Parse_Quotes(BodyText) & "');"
            MyCommon.LRT_Execute()
          End If
        Next
                ' End If
            End If
  End Sub
  
  Function AddMessageDeliverable(ByVal OfferID As Long, ByVal RewardID As Long, ByVal MessageID As Long, ByVal Phase As Long, _
                                 ByVal Required As Boolean) As Long
    Dim MyCommon As New Copient.CommonInc
    Dim DeliverableID As Long = 0

    Try
      'Add the printed message to the RewardOption
      MyCommon.Open_LogixRT()
      ' Create a message if one doesn't already exist
      MyCommon.QueryStr = "dbo.pa_CPE_AddPrintedMessage"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@ROID", SqlDbType.Int, 4).Value = RewardID
      MyCommon.LRTsp.Parameters.Add("@MessageID", SqlDbType.Int, 4).Value = MessageID
      MyCommon.LRTsp.Parameters.Add("@Phase", SqlDbType.Int, 4).Value = Phase
      MyCommon.LRTsp.Parameters.Add("@Required", SqlDbType.Int, 4).Value = IIf(Required, 1, 0)
      MyCommon.LRTsp.Parameters.Add("@DeliverableID", SqlDbType.Int, 4).Direction = ParameterDirection.Output
      MyCommon.LRTsp.ExecuteNonQuery()
      DeliverableID = MyCommon.LRTsp.Parameters("@DeliverableID").Value
      MyCommon.Close_LRTsp()
    Catch ex As Exception
      Send(ex.ToString)
    Finally
      MyCommon.Close_LogixRT()
      MyCommon = Nothing
    End Try
    Return DeliverableID
  End Function

  Function IsValidMessageEntry(ByRef InfoMessage As String, ByVal DefaultLanguageID As Integer, ByVal TierLevels As Integer) As Boolean
    Dim ValidEntry As Boolean = False
    Dim TempStr As String = ""
    Dim SortID As Integer
        Dim BodyText As String
        Dim IsMessageEmpty As Boolean = False
    TempStr = GetCgiValue("SortID")
    If TempStr Is Nothing OrElse TempStr.Trim = "" Then TempStr = "0"
    
    ValidEntry = Integer.TryParse(TempStr, SortID) AndAlso SortID >= 0 AndAlso SortID <= 999
    If Not ValidEntry Then
      InfoMessage = Copient.PhraseLib.Lookup("reward.invalid-sortid", LanguageID)
    End If
        If InfoMessage = "" Then
            For Tier = 1 To TierLevels
                BodyText = GetCgiValue("t" & Tier & "_text_" & DefaultLanguageID)
                If BodyText = "" Then
                    IsMessageEmpty = True
                    ValidEntry = False
                    InfoMessage = Copient.PhraseLib.Lookup("term.messagerequiredwarning", LanguageID) & " " & Tier
                    Exit For
                ElseIf BodyText.Length > 4000 Then
                    ValidEntry = False
                    InfoMessage = String.Format(Copient.PhraseLib.Lookup("CPEoffer-rew.passthru-AlphanumericLengthExceed", LanguageID), Copient.PhraseLib.Lookup("term.message", LanguageID), "4000")
                    Exit For
                End If
            Next
        End If
    Return ValidEntry
  End Function
  
  
  '----------------------------------------------------------------------------------------------
  
  Function ProcessMsgLine(ByVal PrinterTypeID As Integer, ByVal PrintLine As String, ByRef MyCommon As Copient.CommonInc) As String
    
    Dim ReturnVal As String = PrintLine
    Dim Done As Boolean = False
    Dim ptr As Integer
    Dim TagStart As Integer
    Dim TagEnd As Integer
    Dim FullTempTag As String
    Dim TempTag As String
    Dim PrinterTag As String
    Dim row As DataRow
    Dim ParamPos As Integer
    Dim TempStr As String
    Dim TagDST As DataTable
    Dim TagMatchSuccess As Boolean = False
    Dim index As Integer
    
    ' get the tags
    MyCommon.QueryStr = "select isnull(PrinterTypeID, 0) as PrinterTypeID, '|' + tag + '|' as Tag, isnull(pt.previewchars,'') as PreviewChars, AlignmentTag from MarkupTags as MT with (NoLock) " & _
                        "left join PrinterTranslation as PT with (NoLock) on MT.MarkupID=PT.MarkupID;"
    TagDST = MyCommon.LRT_Select

    ptr = 1
    'Send("<!-- PrintLine=" & PrintLine & " -->")
    While Not (Done) 'look for tags until we are finished with the line
      index = index + 1
      'get the next tag from the PrintLine
      TagStart = InStr(ptr, PrintLine, "|", CompareMethod.Binary)
      If TagStart > 0 Then
        TagEnd = InStr(TagStart + 1, PrintLine, "|", CompareMethod.Binary)
        If TagEnd > 0 Then
          TagEnd = TagEnd + 1
          FullTempTag = Mid(PrintLine, TagStart, TagEnd - TagStart)
          
          'handling for tags with parameters
          ParamPos = InStr(FullTempTag, "[", CompareMethod.Binary)
          If ParamPos > 0 Then
            TempTag = Left(FullTempTag, ParamPos - 1) & "|"
          Else
            TempTag = FullTempTag
          End If
          If TempTag.IndexOf(ControlChars.Cr) > -1 Then TempTag = TempTag.Replace(ControlChars.Cr, String.Empty)
          If TempTag.IndexOf(ControlChars.Lf) > -1 Then TempTag = TempTag.Replace(ControlChars.Lf, String.Empty)
          
          'Send("  <!-- Found possible tag: " & TempTag & " -->")
          'we have found the starting and ending points of a possible tag
          'nows search through our list and see if there is a match
          TagMatchSuccess = False
          For Each row In TagDST.Rows
            PrinterTag = row.Item("tag")
            'Send("<!-- TempTag=" & TempTag & "  PrinterTag=" & PrinterTag & " -->")
            If TempTag = PrinterTag And row.Item("PrinterTypeID") = PrinterTypeID Then
              TagMatchSuccess = True
              'we found a tag match, check if it has a line break character 
              TempStr = FullTempTag
              If TempStr.IndexOf(ControlChars.Cr) > -1 Then TempStr = TempStr.Replace(ControlChars.Cr, String.Empty)
              If TempStr.IndexOf(ControlChars.Lf) > -1 Then TempStr = TempStr.Replace(ControlChars.Lf, String.Empty)
              ReturnVal = ReturnVal.Replace(FullTempTag, TempStr)
            End If
          Next
          If TagMatchSuccess Then
            ptr = TagEnd  'starting looking for the next tag with the first charater AFTER the last tag end
          Else
            'what we thought might have been a valid tag, wasn't a match in the database - maybe it was just pipe characters in the message text?
            ptr = TagEnd - 1 'start looking for the next tag using the invalid tag's ending pipe character
          End If
        Else
          Done = True
        End If
      Else
        Done = True
      End If
      'enable for runaway prevention
      If index > 100 Then
        Done = True
        Send("<!-- loop runaway in ProcessMsgLine -->")
      End If
    End While
    'Send("<!-- ReturnVal=" & ReturnVal & " -->")
    
    Return ReturnVal
    
  End Function
  
  
</script>
<script type="text/javascript">
    function setSameTierValue(tierLevels, multiLanguageID){
      var box = document.getElementById("useSameTierValue");
      var textarea1 = "t1_text_" + multiLanguageID.toString();
      var text;

      if(box.checked){
        for (i=1; i < (tierLevels + 1); i++){
          text = "t" + i.toString() + "_text_" + multiLanguageID.toString();
          //alert(document.getElementById(text).value.toString());
          document.getElementById(text).value = document.getElementById(textarea1).value;
          document.getElementById(text).setAttribute('disabled', 'disabled');
        } 
      }
      else{
        for (i=1; i < (tierLevels + 1); i++){
          text = "t" + i.toString() + "_text_" + multiLanguageID.toString();
          document.getElementById(text).disabled = false;
        } 
      }
    }

    function enableTiers(tierLevels,  multiLanguageID){
      //alert(multiLanguageID.toString());
      var box = document.getElementById("useSameTierValue");
      var text;
      if(box.checked && tierLevels > 1){
        for (i=1; i < (tierLevels + 1); i++){
          text = "t" + i.toString() + "_text_" + multiLanguageID.toString();
          document.getElementById(text).disabled = false;
        } 
      }
    }

<%
  If (CloseAfterSave) Then
    Sendb("  window.close();")
  Else
    Sendb("  xmlhttpPost(""/logix/UE/UEpmsg-feeds.aspx"");")
  End If
  Sendb("  lineCounter();")
%>
</script>
<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd("mainform", "t1_text_" & DefaultLanguageID)
  Logix = Nothing
  MyCommon = Nothing
%>
