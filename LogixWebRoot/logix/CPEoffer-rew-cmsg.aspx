<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: CPEoffer-rew-cmsg.aspx 
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
  Dim Localization As Copient.Localization
  Dim rst As DataTable
  Dim row As DataRow
  Dim rst2 As DataTable
  Dim dt As DataTable
  Dim row2 As DataRow
  Dim OfferID As Long
    Dim Name As String = ""
    Dim Line1 As String = ""
  Dim Line2 As String = ""
  Dim TierLine1 As String = ""
  Dim TierLine2 As String = ""
  Dim TierLine2Tag As String = ""
  Dim LineLength As Integer = 0
  Dim Line2Text As String = ""
  Dim Line2Tag As String = ""
  Dim HasTag As Boolean = False
  Dim TagStart As Integer = 0
  Dim TierBeep As Integer = 0
  Dim TierBeepDuration As Integer = 0
  Dim RewardID As Long
  Dim DeliverableID As Long
  Dim MessageID As Long
  Dim MsgAdded As Boolean = False
  Dim bIsErrorMsg As Boolean = False
  Dim Phase As Integer = 0
  Dim PhaseTitle As String = ""
  Dim TouchPoint As Integer = 0
  Dim TpROID As Integer = 0
  Dim Disallow_Edit As Boolean = True
  Dim FromTemplate As Boolean = False
  Dim IsTemplate As Boolean = False
  Dim IsTemplateVal As String = "Not"
  Dim DisabledAttribute As String = ""
  Dim CloseAfterSave As Boolean = False
  Dim Beep As Integer = 0
  Dim BeepDuration As Integer = 1
  Dim BeepDurDisplay As String = "none"
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim EngineID As Integer = 2
  Dim BannersEnabled As Boolean = True
  Dim TierLevels As Integer = 1
  Dim t As Integer = 1
  Dim l = 1
  Dim ValidTiers As Boolean = True
  Dim DisplayImmediate As Integer = 0
  Dim Display As Boolean = False
  Dim LanguagesDT As DataTable
  Dim MLI As New Copient.Localization.MultiLanguageRec
  Dim MultiLanguageEnabled As Boolean = False
  Dim DefaultLanguageID As Integer = 0
  Dim PKID As Integer = 0
  Dim InstalledLanguages As String = String.Empty
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "CPEoffer-rew-cmsg.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  Localization = New Copient.Localization(MyCommon)
          
  MultiLanguageEnabled = IIf(MyCommon.Fetch_SystemOption(124) = "1", True, False)
  Integer.TryParse(MyCommon.Fetch_SystemOption(1), DefaultLanguageID)
  If DefaultLanguageID = 0 Then DefaultLanguageID = 1
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  
  OfferID = Request.QueryString("OfferID")
  RewardID = MyCommon.Extract_Val(Request.QueryString("RewardID"))
  DeliverableID = MyCommon.Extract_Val(Request.QueryString("DeliverableID"))
  MessageID = MyCommon.Extract_Val(Request.QueryString("MessageID"))
  If (Request.QueryString("EngineID") <> "") Then
    EngineID = MyCommon.Extract_Val(Request.QueryString("EngineID"))
  Else
    MyCommon.QueryStr = "select EngineID from OfferIDs where OfferID=" & OfferID & ";"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
      EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), 0)
    End If
  End If
  
  If Request.QueryString("DisplayImmediateGrp") <> "" Then
    If Request.QueryString("DisplayImmediateGrp") = "1" Then
      DisplayImmediate = 1
    Else
      DisplayImmediate = 0
    End If
  Else
    DisplayImmediate = 0
  End If
  
  TouchPoint = MyCommon.Extract_Val(Request.QueryString("tp"))
  If (TouchPoint > 0) Then TpROID = MyCommon.Extract_Val(Request.QueryString("roid"))
  
  ' Fetch the name
  MyCommon.QueryStr = "select IncentiveName, IsTemplate, FromTemplate from CPE_Incentives with (NoLock) " & _
                      "where IncentiveID=" & OfferID
  rst = MyCommon.LRT_Select
  If (rst.Rows.Count > 0) Then
    Name = MyCommon.NZ(rst.Rows(0).Item("IncentiveName"), "")
    IsTemplate = MyCommon.NZ(rst.Rows(0).Item("IsTemplate"), False)
    FromTemplate = MyCommon.NZ(rst.Rows(0).Item("FromTemplate"), False)
  End If
  IsTemplateVal = IIf(IsTemplate, "IsTemplate", "Not")
  
  Phase = MyCommon.Extract_Val(Request.QueryString("Phase"))
  If (Phase = 0) Then Phase = MyCommon.Extract_Val(Request.Form("Phase"))
  If (Phase = 0) Then Phase = 3
  Select Case Phase
    Case 1 ' Notification
      PhaseTitle = "term.cmsgnotification"
    Case 2 ' Accumulation
      PhaseTitle = ""
    Case 3 ' Reward
      PhaseTitle = "term.cmsgreward"
    Case Else
      PhaseTitle = "term.cmsgreward"
  End Select
  
  If Phase = 3 Then
    MyCommon.QueryStr = "select TierLevels from CPE_RewardOptions with (NoLock) where RewardOptionID=" & RewardID & ";"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
      TierLevels = MyCommon.NZ(rst.Rows(0).Item("TierLevels"), 1)
    End If
  Else
    TierLevels = 1
  End If
  
  Line1 = Left(Trim(Request.QueryString("t1_line1_" & DefaultLanguageID)), 20)
  Line2 = Left(Trim(Request.QueryString("t1_line2_" & DefaultLanguageID)), 20)
  Line2Tag = Request.QueryString("t1_line2tag_" & DefaultLanguageID)
  Beep = MyCommon.Extract_Val(Request.QueryString("t1_beep"))
  BeepDuration = MyCommon.Extract_Val(Request.QueryString("t1_beepDuration"))
  
  If (Request.QueryString("save") <> "" AndAlso DeliverableID = 0) Then
    ' Create a new cashier message
    MsgAdded = Create_Message(OfferID, Line1, Line2, Line2Tag, Phase, TpROID, DeliverableID)
    If (DeliverableID > 0) Then
      ' Get the new message ID
      MyCommon.QueryStr = "select OutputID from CPE_Deliverables with (NoLock) where DeliverableID=" & DeliverableID & ";"
      rst = MyCommon.LRT_Select()
      If rst.Rows.Count > 0 Then
        MessageID = MyCommon.NZ(rst.Rows(0).Item("OutputID"), 0)
      End If
      ' Delete any existing records for this message from the tiers table, then insert new values
      MyCommon.QueryStr = "delete from CPE_CashierMsgTranslations with (RowLock) where CashierMsgTierID in " & _
                          "(select PKID from CPE_CashierMessageTiers where MessageID in (0, " & MessageID & "));"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "delete from CPE_CashierMessageTiers with (RowLock) where MessageID in (0, " & MessageID & ");"
      MyCommon.LRT_Execute()
      If TierLevels > 1 Then DisplayImmediate = 0
      ' Insert tier values
      t = 1
      For t = 1 To TierLevels
        Create_MessageTiers(MessageID, DisplayImmediate, t, DefaultLanguageID)
      Next
    End If
    ' Update the CPE_Incentives table
    MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 " & _
                        "where IncentiveID=" & OfferID & ";"
    MyCommon.LRT_Execute()
    infoMessage = IIf((MsgAdded), Copient.PhraseLib.Lookup("CPE-rew-cmsg.created", LanguageID) & OfferID, Copient.PhraseLib.Lookup("CPE-rew-cmsg.error", LanguageID))
    CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup(IIf(Phase = 1, "CPE_not.createdcmsg", "CPE_Reward.createdcmsg"), LanguageID))
  ElseIf (Request.QueryString("save") <> "" AndAlso DeliverableID > 0) Then
    ' Update an existing cashier message
    MyCommon.QueryStr = "update CPE_CashierMessages with (RowLock) set LastUpdate=getDate() " & _
                        "where MessageID=" & MessageID & ";"
    MyCommon.LRT_Execute()
    ' Delete any existing records for this message from the tiers table, then insert new values
    MyCommon.QueryStr = "delete from CPE_CashierMsgTranslations with (RowLock) where CashierMsgTierID in " & _
                        "(select PKID from CPE_CashierMessageTiers where MessageID in (0, " & MessageID & "));"
    MyCommon.LRT_Execute()
    MyCommon.QueryStr = "delete from CPE_CashierMessageTiers with (RowLock) where MessageID in (0, " & MessageID & ");"
    MyCommon.LRT_Execute()
    ' Insert tier values
    t = 1
    For t = 1 To TierLevels
      Create_MessageTiers(MessageID, DisplayImmediate, t, DefaultLanguageID)
    Next
    ' Update the CPE_Incentives table
    MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 " & _
                        "where IncentiveID=" & OfferID & ";"
    MyCommon.LRT_Execute()
    infoMessage = Copient.PhraseLib.Lookup("CPE-rew-cmsg.edit", LanguageID) & OfferID
    CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup(IIf(Phase = 1, "CPE_not.editcmsg", "CPE_Reward.editcmsg"), LanguageID))
  End If
  
  MyCommon.QueryStr = "select CM.MessageID, CMT.Line1, CMT.Line2, CMT.Beep, CMT.BeepDuration " & _
                      "from CPE_Deliverables D with (NoLock) " & _
                      "inner join CPE_CashierMessages CM with (NoLock) on D.OutputID=CM.MessageID " & _
                      "inner join CPE_CashierMessageTiers CMT with (NoLock) on CMT.MessageID=CM.MessageID " & _
                      "where D.RewardOptionID=" & RewardID & " and D.DeliverableID=" & DeliverableID & " and D.Deleted=0 and DeliverableTypeID=9;"
  rst = MyCommon.LRT_Select
  If (rst.Rows.Count > 0) Then
    Line1 = MyCommon.NZ(rst.Rows(0).Item("Line1"), "")
    Line2 = MyCommon.NZ(rst.Rows(0).Item("Line2"), "")
    MessageID = MyCommon.NZ(rst.Rows(0).Item("MessageID"), 0)
    Beep = MyCommon.NZ(rst.Rows(0).Item("Beep"), 0)
    BeepDuration = MyCommon.NZ(rst.Rows(0).Item("BeepDuration"), 0)
  End If
  
  'update the templates permission if necessary
  If (Request.QueryString("save") <> "" AndAlso Request.QueryString("IsTemplate") = "IsTemplate") Then
    ' time to update the status bits for the templates
    Dim form_Disallow_Edit As Integer = 0
    If (Request.QueryString("Disallow_Edit") = "on") Then
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
  
  If Not IsTemplate Then
    DisabledAttribute = IIf(Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Edit), "", " disabled=""disabled""")
  Else
    DisabledAttribute = IIf(Logix.UserRoles.EditTemplates, "", " disabled=""disabled""")
  End If
  
  Send_HeadBegin("term.offer", PhaseTitle, OfferID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
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

//
//
// Functions

function edShowButton(button) {
	if (button.access) {
		var accesskey = ' accesskey = "' + button.access + '"'
	}	else {
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
			} else {
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
var srcElement;

function edInsertContent(myValue) {
  var srcElementName = srcElement.name;
  var tier = srcElementName.slice(1,2);
  var languageID = srcElementName.slice(-2);
  if (languageID.substring(0, 1) == "_") {
    languageID = languageID.slice(1);
  }
  
  var line2 = document.getElementById('t' + tier + '_line2_' + languageID);
  var line2tag = document.getElementById('t' + tier + '_line2tag_' + languageID);
  var line2tagdisplay = document.getElementById('t' + tier + '_line2tagdisplay_' + languageID);
  var droptagButton = document.getElementById('t' + tier + '_droptag_' + languageID);
  
  droptagButton.style.display = '';
  line2tagdisplay.style.display = '';
  line2tagdisplay.innerHTML = myValue;
  line2tag.value = myValue;
  line2.value = line2.value.slice(0,10);
  line2.style.width = '100px';
  line2.maxLength = '10';
  
//	//IE support
//	if (document.selection) {
//		myField.focus();
//		sel = document.selection.createRange();
//		sel.text = myValue;
//		myField.focus();
//	}
//	//MOZILLA/NETSCAPE support
//	else if (myField.selectionStart || myField.selectionStart == '0') {
//		var startPos = myField.selectionStart;
//		var endPos = myField.selectionEnd;
//		var scrollTop = myField.scrollTop;
//		myField.value = myField.value.substring(0, startPos)
//		              + myValue 
//                      + myField.value.substring(endPos, myField.value.length);
//		myField.focus();
//		myField.selectionStart = startPos + myValue.length;
//		myField.selectionEnd = startPos + myValue.length;
//		myField.scrollTop = scrollTop;
//	} else {
//		myField.value += myValue;
//		myField.focus();
//	}
}

<%
  MyCommon.QueryStr = "Select Distinct MT.MarkupID,MT.Tag,MT.Description,MT.PhraseID," & _
                      "MT.NumParams,MT.Param1Name,MT.Param1PhraseID,MT.Param2Name,MT.Param2PhraseID,MT.Param3Name,MT.Param3PhraseID," & _
                      "MT.DisplayOrder,MTU.RewardTypeID,MTU.EngineID,PhT.PhraseID,PhT.LanguageID," & _
                      "Convert(nvarchar(50),PhT1.Phrase) as Param1Phrase," & _
                      "Convert(nvarchar(50),PhT2.Phrase) as Param2Phrase," & _
                      "Convert(nvarchar(50),PhT3.Phrase) as Param3Phrase," & _
                      "Convert(nvarchar(50),PhT.Phrase) as Phrase from MarkupTags as MT with (NoLock) " & _
                      "Left Join PhraseText as PhT with (NoLock) on MT.PhraseID=PhT.PhraseID and PhT.LanguageId=" & LanguageId & " " & _
                      "Left Join PhraseText as PhT1 with (NoLock) on MT.Param1PhraseID=PhT1.PhraseID and PhT1.LanguageId=" & LanguageId & " " & _
                      "Left Join PhraseText as PhT2 with (NoLock) on MT.Param2PhraseID=PhT2.PhraseID and PhT2.LanguageId=" & LanguageId & " " & _
                      "Left Join PhraseText as PhT3 with (NoLock) on MT.Param3PhraseID=PhT3.PhraseID and PhT3.LanguageId=" & LanguageId & " " & _
                      "Inner Join MarkupTagUsage as MTU with (NoLock) on MT.MarkupID=MTU.MarkupID " & _
                      "where MTU.RewardTypeID=9 and MTU.EngineID=" & EngineID & " order by MT.DisplayOrder"
  rst = MyCommon.LRT_Select
  Dim funcname As String
  For Each row In rst.Rows
    funcname = MyCommon.NZ(row.Item("Tag"), "")
    funcname = funcname.Replace("#", "Amt")
    funcname = funcname.Replace("$", "Dol")
    funcname = funcname.Replace("/", "Off")
    If (MyCommon.NZ(row.Item("NumParams"), 0) = 0) Then
      Send("function edInsert" & (StrConv(funcname, VbStrConv.ProperCase)) & "() {")
      Send("  myValue = '|" & row.Item("Tag") & "|';")
      Send("  edInsertContent(myValue);")
      Send("}")
    Else
      Send("function edInsert" & (StrConv(funcname, VbStrConv.ProperCase)) & "() {")
      If StrConv(funcname, VbStrConv.ProperCase) = "Svbal" Then
        Send("  var myValue = document.getElementById(""functionselect1"").value;")
      Else If StrConv(funcname, VbStrConv.ProperCase) = "Svvalnet" Then
	    Send("  var myValue = document.getElementById(""functionselect1"").value;")
      Else
        Send("  var myValue = document.getElementById(""functionselect2"").value;")
      End If
      Send("  if (myValue) {")
      Send("    myValue = '|" & row.Item("Tag") & "[' + myValue + ']|';")
      Send("    edInsertContent(myValue);")
      Send("  }")
      Send("}")
    End If
  Next
%>

function doPreviewPopup() {
  var numtiers = document.getElementById('TierLevels').value;
  var text = "";
  
  if (srcElement != null) {
    text = "" + srcElement.name;
  } else {
    text = "" + 't1_line1_<%Sendb(DefaultLanguageID)%>';
  }
  if (numtiers == '1') {
    var text1 = ""
    var myField1 = document.getElementById("t1_line1_<%Sendb(DefaultLanguageID)%>").value;
    var myField2 = document.getElementById("t1_line2_<%Sendb(DefaultLanguageID)%>").value + document.getElementById("t1_line2tag_<%Sendb(DefaultLanguageID)%>").value;
  } else {
    var text1 = "Tier " + text.substring(1,2) + ": ";
    var myField1 = document.getElementById("t" + text.substring(1,2) + "_line1_<%Sendb(DefaultLanguageID)%>").value;
    var myField2 = document.getElementById("t" + text.substring(1,2) + "_line2_<%Sendb(DefaultLanguageID)%>").value + document.getElementById("t" + text.substring(1,2) + "_line2tag_<%Sendb(DefaultLanguageID)%>").value;
  }
  var myUrl = 'offer-rew-cmsgpreview.aspx?Line1=' + escape(myField1) + '&Line2=' + escape(myField2) + '&TierLevel=' + text1;
  openMiniPopup(myUrl);
}
</script>
<script type="text/javascript">
// This is the javascript array holding the function list
// The PrintJavascriptArray ASP function can be used to print this array.
<%
  MyCommon.QueryStr = "Select SVProgramID, Name from StoredValuePrograms where Deleted=0 order by Name;"
  rst = MyCommon.LRT_Select
  If (rst.rows.count>0)
    Sendb("var functionlist1 = Array(")
    For Each row In rst.Rows
      Sendb("""" & MyCommon.NZ(row.item("Name"), "").ToString().Replace("""", "\""") & """,")
    Next
    Send(""""");")
    Sendb("var vallist1 = Array(")
    For Each row In rst.Rows
      Sendb("""" & MyCommon.NZ(row.item("SVProgramID"), 0) & """,")
    Next
    Send(""""");")
  End If
  
  MyCommon.QueryStr = "Select ProgramID, ProgramName from PointsPrograms where Deleted=0 order by ProgramName;"
  rst = MyCommon.LRT_Select
  If (rst.rows.count>0)
    Sendb("var functionlist2 = Array(")
    For Each row In rst.Rows
      Sendb("""" & MyCommon.NZ(row.item("ProgramName"), "").ToString().Replace("""", "\""") & """,")
    Next
    Send(""""");")
    Sendb("var vallist2 = Array(")
    For Each row In rst.Rows
      Sendb("""" & MyCommon.NZ(row.item("ProgramID"), 0) & """,")
    Next
    Send(""""");")
  End If
%>

// This is the function that refreshes the list after a keypress.
// The maximum number to show can be limited to improve performance with
// huge lists (1000s of entries).
// The function clears the list, and then does a linear search through the
// globally defined array and adds the matches back to the list.
function handleKeyUp1(maxNumToShow) {
  var selectObj, textObj, functionListLength1;
  var i,  numShown;
  var searchPattern;
  
  document.getElementById("functionselect1").size = "10";
  
  // Set references to the form elements
  selectObj = document.forms[0].functionselect1;
  textObj = document.forms[0].functioninput1;
  
  // Remember the function list length for loop speedup
  functionListLength1 = functionlist1.length;
  
  // Set the search pattern depending
  if(document.forms[0].functionradio1[0].checked == true) {
    searchPattern = "^"+textObj.value;
  } else {
    searchPattern = textObj.value;
  }
  searchPattern = cleanRegExpString(searchPattern);
  
  // Create a regulare expression
  re = new RegExp(searchPattern,"gi");
  
  // Clear the options list
  selectObj.length = 0;
  
  // Loop through the array and re-add matching options
  numShown = 0;
  for(i = 0; i < functionListLength1; i++) {
    if(functionlist1[i].search(re) != -1) {
      selectObj[numShown] = new Option(functionlist1[i],vallist1[i]);
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

function handleKeyUp2(maxNumToShow) {
  var selectObj, textObj, functionListLength2;
  var i,  numShown;
  var searchPattern;
  
  document.getElementById("functionselect2").size = "10";
  
  // Set references to the form elements
  selectObj = document.forms[0].functionselect2;
  textObj = document.forms[0].functioninput2;
  
  // Remember the function list length for loop speedup
  functionListLength2 = functionlist2.length;
  
  // Set the search pattern depending
  if(document.forms[0].functionradio2[0].checked == true) {
    searchPattern = "^"+textObj.value;
  } else {
    searchPattern = textObj.value;
  }
  searchPattern = cleanRegExpString(searchPattern);
  
  // Create a regulare expression
  re = new RegExp(searchPattern,"gi");
  
  // Clear the options list
  selectObj.length = 0;
  
  // Loop through the array and re-add matching options
  numShown = 0;
  for(i = 0; i < functionListLength2; i++) {
    if(functionlist2[i].search(re) != -1) {
      selectObj[numShown] = new Option(functionlist2[i],vallist2[i]);
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
  if (type == 1) {
    selectObj = document.forms[0].functionselect1;
    textObj = document.forms[0].functioninput1;
    selectedValue = document.getElementById("functionselect1").value;
  } else {
    selectObj = document.forms[0].functionselect2;
    textObj = document.forms[0].functioninput2;
    selectedValue = document.getElementById("functionselect2").value;
  }
  
  if(selectedValue != "") {
    if (type == 1) {
      var elemTag = document.getElementById("svTagName");
    } else {
      var elemTag = document.getElementById("ptTagName");
    }
    if (elemTag.value == "Svbal") {
      edInsertSvbal(selectedValue);
	} else if (elemTag.value == "Svvalnet") {
	  edInsertSvvalnet(selectedValue);
    } else if (elemTag.value == "Ptbal") {
      edInsertPtbal(selectedValue);
    }
    showDialogSpan(false, type, "")
  }
}

function showDialogSpan(bShow, type, caption) {
  var elemBox = document.getElementById("dialogbox");
  var elemSv = document.getElementById("svselector");
  var elemSvTag = document.getElementById("svTag");
  var elemPt = document.getElementById("pointselector");
  var elemPtTag = document.getElementById("ptTag");
  var elemTag = null;
  
  if (bShow) {
    if (elemSv != null && type == 1) {
      elemSv.style.display = "block";
      if (caption != "" && elemSvTag != null) {
        elemSvTag.innerHTML = "Tag Type: " + caption
        elemTag = document.getElementById("svTagName");
        if (elemTag != null) {
          elemTag.value = caption;
        }
      }
    }
    else if (elemPt != null && type == 2) {
      elemPt.style.display = "block";
      if (caption != "" && elemPtTag != null) {
        elemPtTag.innerHTML = "Tag Type: " + caption
        elemTag = document.getElementById("ptTagName");
        if (elemTag != null) {
          elemTag.value = caption;
        }
      }
    }
  }
  
  if (elemBox != null) {
    elemBox.style.display = (bShow) ? "block" : "none";
  }
  if (elemSv != null) {
    elemSv.style.display = (bShow && type == 1) ? "block" : "none";
  }
  if (elemPt != null) {
    elemPt.style.display = (bShow && type == 2) ? "block" : "none";
  }
}

function xmlhttpPost(strURL) {
  var xmlHttpReq = false;
  var self = this;
  
  document.getElementById("tools").innerHTML = "<div class=\"loading\"><img id=\"clock\" src=\"../images/clock22.png\" \/><br \/>" + '<% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID)) %><\/div>';
  
  // Mozilla/Safari
  if (window.XMLHttpRequest) {
    self.xmlHttpReq = new XMLHttpRequest();
  }
  // IE
  else if (window.ActiveXObject) {
    self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
  }
  strURL += "?" + getRequestString();
  self.xmlHttpReq.open('POST', strURL, true);
  self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
  self.xmlHttpReq.onreadystatechange = function() {
    if (self.xmlHttpReq !=null && self.xmlHttpReq.readyState == 4) {
      <%
        If DisabledAttribute = "" Then
          Sendb("updatepage(self.xmlHttpReq.responseText, " & TierLevels & ");")
        Else
          Sendb("document.getElementById(""tools"").innerHTML = ""&nbsp;"";")
        End If
      %>
    }
  }
  self.xmlHttpReq.send(getRequestString());
}

function getRequestString() {    
  var qstr = "";
  var textAreaName = "t1_line1";
  var OfferID = document.getElementById('OfferID').value  
  qstr = "Mode=MarkupTags&EngineID=2&TextAreaName=" + textAreaName + "&OfferID=" + OfferID
  return qstr;
}

function updatepage(str, t){
  var selElem = document.getElementById('printerselect');
  var ptWidthElem = null;
  var taElem = document.getElementById("t1_line1");
  var i = 1;
  
  document.getElementById("tools").innerHTML = str;
  
  if (selElem != null) {
    ptWidthElem = document.getElementById("PT" + selElem.value);
    for(i = 1; i <= t; i++) {
      taElem = document.getElementById("t" + i + "_line1");
      if (ptWidthElem != null) {
        if (taElem != null) {
          <%
          If (Request.Browser.Type.ToString.ToUpper.IndexOf("FIREFOX") > -1) Then
            Send("taElem.style.width = ((parseInt(ptWidthElem.value) * 8) + 18) + 'px';")
          Else
            Send("taElem.style.width = ((parseInt(ptWidthElem.value) * 8) + 22) + 'px';")
          End If
          %>
        }
      } else {
        if (taElem != null) {
          <%
          If (Request.Browser.Type.ToString.ToUpper.IndexOf("FIREFOX") > -1) Then
            Send("taElem.style.width = ((parseInt(52) * 8) + 18) + 'px';")
          Else
            Send("taElem.style.width = ((parseInt(52) * 8) + 22) + 'px';")
          End If
          %>
        }    
      } 
    }
  }
}

function cleanMessage() {
  var elem = document.getElementById("t1_line1");
  
  if (elem != null) {
    elem.value = elem.value.replace("<", "\1");
  }
  
  return true;
}
</script>
<script type="text/javascript">
  function beepTypeChanged(elem, t) {
    var BEEP_DURATION_VALUE = 3;
    var elemDurationRow = document.getElementById("t" + t + "_trDuration");
    var elemDurText = document.getElementById("t" + t + "_beepDuration");

    if (elem != null && elemDurationRow != null) {
      if (elem.options[elem.selectedIndex].value == BEEP_DURATION_VALUE) {
        elemDurationRow.style.display = "";
        elemDurText.focus();
        elemDurText.select();
      } else {
        elemDurationRow.style.display = "none";
        elemDurText.value = "";
      }
    }
  }

  function validateEntry(text, t) {
    var retVal = true;
    var elemType = document.getElementById("t" + t + "_beep");
    var elemDur = document.getElementById("t" + t + "_beepDuration");
    var elemDurationRow = document.getElementById("t" + t + "_trDuration");
    var saveElem = document.getElementById("save");
    var msg = "";

    if (elemType.value == "0") {
      elemDur.value = ""
    }
    if (elemDur != "" && elemDurationRow.style.display != "none" && (isNaN(elemDur.value) || parseInt(elemDur.value) < 0)) {
      if (text != "") {
        msg += "(" + text + ": " + t + ") ";
      }
      msg += '<% Sendb(Copient.PhraseLib.Lookup("term.beep-warning", LanguageID)) %>';
      retVal = false;
    }
    if (msg != "") {
      alert(msg);
      if (saveElem != null) {
        if (saveElem.style.visibility == 'hidden') {
          saveElem.style.visibility = 'visible';
        }
      }
    }

    return retVal;
  }

  function dropTag(tier, languageID) {
    var line2 = document.getElementById('t' + tier + '_line2_' + languageID);
    var line2tag = document.getElementById('t' + tier + '_line2tag_' + languageID);
    var line2tagdisplay = document.getElementById('t' + tier + '_line2tagdisplay_' + languageID);
    var droptagButton = document.getElementById('t' + tier + '_droptag_' + languageID);

    line2.maxLength = '20';
    line2.style.width = '200px';
    line2tag.value = '';
    line2tagdisplay.innerHTML = '';
    line2tagdisplay.style.display = 'none';
    droptagButton.style.display = 'none';
  }
</script>
<%
  Dim i As Integer
  Send("<script type=""text/javascript"">")
  Send("function validateEntries() {")
  Send("  var bRetVal = true;")
  For i = 1 To TierLevels
    Send("  if (bRetVal) {")
    If TierLevels > 1 Then
      Send("    bRetVal=validateEntry(""" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & """,""" & i & """);")
    Else
      Send("    bRetVal=validateEntry("""",""" & i & """);")
    End If
    Send("  }")
  Next
  Send("  var IsValidLineTags = true; ")
  Send("  var installedlanguages = $('#InstalledLanguages').val(); ")
  Send("  var languages = installedlanguages.split(','); ")
  Send("  var numtiers = $('#TierLevels').val(); ")
  Send("  for(var tierLevel=1; tierLevel <= numtiers; tierLevel++) { ")
  Send("    for(var i=0; i< languages.length; i++) { ")
  Send("      var line1 = $('#t' + tierLevel + '_line1_' + languages[i]); ")
  Send("      var line2 = $('#t' + tierLevel + '_line2_' + languages[i]); ")
  Send("      var line2tag = $('#t' + tierLevel + '_line2tagdisplay_' + languages[i]); ")
  Send("      if( line2tag.text() != '' && ($.trim(line1.val()).length == 0 && $.trim(line2.val()).length == 0)) { ")
  Send("        bRetVal = false; ")
  Send("        IsValidLineTags = false; ")
  Send("      } ")
  Send("    }  ")
  Send("  }  ")
  Send("  if (IsValidLineTags == false) {  ")
  Send("    if ($('div#infobar').length == 0) { ")
  Send("      $('div#main').prepend('<div id=""infobar"" class=""red-background""></div>'); ")
  Send("    } ")
  Send("    $('div#infobar').html('" & Copient.PhraseLib.Lookup("term.errorsavetags", LanguageID) & "');")
  Send("  }  ")
  Send("  return bRetVal;")
  Send("}")
  Send("")
  
  Send("function ChangeParentDocument() { ")
  If (Phase = 3) Then
          If (EngineID = Copient.CommonInc.InstalledEngines.CAM) Then
                  Send("  opener.location = 'CAM/CAM-offer-rew.aspx?OfferID=" & OfferID & "'; ")
          Else
                  Send("  opener.location = 'CPEoffer-rew.aspx?OfferID=" & OfferID & "'; ")
          End If
  ElseIf (Phase = 1) Then
          Send("  opener.location = 'offer-channels.aspx?OfferID=" & OfferID & "'; ")
  End If
  Send("} ")
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
  
  Dim bProcessTags As Boolean = False
  Try
    MyCommon.QueryStr = "dbo.pa_Cashier_Message_Tags"
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = EngineID
    MyCommon.LRTsp.Parameters.Add("@RewardTypeID", SqlDbType.Int).Value = IIf(EngineID = 1, 4, 9)
    rst = MyCommon.LRTsp_select
    If (rst.Rows.Count > 0) Then
      bProcessTags = True
    End If
    MyCommon.Close_LRTsp()
  Catch ex As Exception
    bProcessTags = False
  End Try
%>
<form action="CPEoffer-rew-cmsg.aspx" id="mainform" name="mainform" onsubmit="return validateEntries();">
<div id="intro">
  <input type="hidden" id="OfferID" name="OfferID" value="<% Sendb(OfferID) %>" />
  <input type="hidden" id="Name" name="Name" value="<% Sendb(Name) %>" />
  <input type="hidden" id="RewardID" name="RewardID" value="<% Sendb(RewardID) %>" />
  <input type="hidden" id="DeliverableID" name="DeliverableID" value="<% Sendb(DeliverableID) %>" />
  <input type="hidden" id="MessageID" name="MessageID" value="<% Sendb(MessageID) %>" />
  <input type="hidden" id="Phase" name="Phase" value="<% Sendb(Phase )%>" />
  <input type="hidden" id="roid" name="roid" value="<% Sendb(TpROID) %>" />
  <input type="hidden" id="tp" name="tp" value="<% Sendb(TouchPoint) %>" />
  <input type="hidden" id="IsTemplate" name="IsTemplate" value="<% Sendb(IsTemplateVal)%>" />
  <input type="hidden" id="TierLevels" name="TierLevels" value="<%Sendb(TierLevels) %>" />
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
        <% if(disallow_edit)then Sendb(" checked=""checked""") %> />
      <label for="Disallow_Edit">
        <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
      </label>
    </span>
    <% End If%>
    <% If (bProcessTags) Then%>
    <button class="regular" id="preview" name="preview" type="button" onclick="javascript:doPreviewPopup();">
      <% Sendb(Copient.PhraseLib.Lookup("term.preview", LanguageID))%>
    </button>
    <% End If%>
    <%
      If Not IsTemplate Then
        If (Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Edit)) Then
          If MessageID = 0 Then
            Send_Save(" onclick=""this.style.visibility='hidden';""")
          Else
            Send_Save()
          End If
        End If
      Else
        If (Logix.UserRoles.EditTemplates) Then
          If MessageID = 0 Then
            Send_Save(" onclick=""this.style.visibility='hidden';""")
          Else
            Send_Save()
          End If
        End If
      End If
    %>
  </div>
</div>
<div id="main">
  <%
    If (infoMessage <> "" And bIsErrorMsg) Then
      Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
    ElseIf (infoMessage <> "") Then
      Send("<div id=""infobar"" class=""green-background"">" & infoMessage & "</div>")
    End If
  %>
  <div id="column1">
    <div class="box" id="message">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.message", LanguageID))%>
        </span>
      </h2>
      <%
        Send("<div style=""height:390px;overflow-x:auto;"">")
          
        Dim TierRecordDT As DataTable
        Dim CustomerFacingLangID As Int32 = 1
        Int32.TryParse(MyCommon.Fetch_SystemOption(125), CustomerFacingLangID)
        For t = 1 To TierLevels
          MyCommon.QueryStr = "select PKID, TierLevel, Line1, Line2, Beep, BeepDuration from CPE_CashierMessageTiers with (NoLock) " & _
                              "where MessageID=" & MessageID & " and TierLevel=" & t & ";"
          TierRecordDT = MyCommon.LRT_Select
          If TierRecordDT.Rows.Count > 0 Then
            PKID = MyCommon.NZ(TierRecordDT.Rows(0).Item("PKID"), 0)
          Else
            PKID = 0
          End If
            
          If TierLevels > 1 Then
            Send("<label for=""t" & t & "_line1"" style=""position:relative;""><b>" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & ":</b></label>")
            Send("<br />")
          End If
            
          ' Beep
          Send("<select id=""t" & t & "_beep"" name=""t" & t & "_beep"" style=""float:left;position:relative;"" onchange=""beepTypeChanged(this, " & t & ");""" & DisabledAttribute & ">")
          MyCommon.QueryStr = "select BeepTypeID, PhraseID from BeepTypes BT with (NoLock);"
          rst2 = MyCommon.LRT_Select
          For Each row2 In rst2.Rows
            Sendb("  <option value=""" & MyCommon.NZ(row2.Item("BeepTypeID"), 0) & """")
            If TierRecordDT.Rows.Count > 0 Then
              If (MyCommon.NZ(row2.Item("BeepTypeID"), 0) = MyCommon.NZ(TierRecordDT.Rows(0).Item("Beep"), 0)) Then
                Sendb(" selected=""selected""")
              End If
            End If
            Sendb(">" & Copient.PhraseLib.Lookup(row2.Item("PhraseID"), LanguageID))
            Send("</option>")
          Next
          Send("</select>")
            
          ' Beep duration
          If TierRecordDT.Rows.Count > 0 Then
            BeepDurDisplay = IIf(MyCommon.NZ(TierRecordDT.Rows(0).Item("Beep"), 0) = 3, "inline", "none")
          End If
          Send("<span id=""t" & t & "_trDuration"" style=""display:" & BeepDurDisplay & "; float:left; position:relative;"">")
          Sendb("  :&nbsp;<input type=""text"" class=""shortest"" id=""t" & t & "_beepDuration"" name=""t" & t & "_beepDuration"" maxlength=""2""")
          If TierRecordDT.Rows.Count = 0 Then
            Send(" value=""0""" & DisabledAttribute & " />")
          Else
            Send(" value=""" & MyCommon.NZ(TierRecordDT.Rows(0).Item("BeepDuration"), 0) & """" & DisabledAttribute & " />")
          End If
          Send("</span>")
          Send("<br clear=""left"" />")
          Send("<br class=""half"" />")
            
          l = 1
          MyCommon.QueryStr = "SELECT L.LanguageID, L.Name, L.MSNetCode, L.JavaLocaleCode, L.PhraseTerm, L.RightToLeftText, T.Line1, T.Line2 " & _
                              "FROM Languages AS L " & _
                              "LEFT JOIN CPE_CashierMsgTranslations AS T ON T.LanguageID=L.LanguageID AND T.CashierMsgTierID=" & PKID & " " & _
                              "WHERE L.LanguageID in (" & IIf(MultiLanguageEnabled, "SELECT TLV.LanguageID FROM TransLanguagesRF_CPE AS TLV", DefaultLanguageID) & ") " & _
                              "ORDER BY CASE WHEN L.LanguageID=" & DefaultLanguageID & " THEN 1 ELSE 2 END, L.LanguageID;"
          LanguagesDT = MyCommon.LRT_Select
          For Each row In LanguagesDT.Rows
            Dim MLLanguageCode As String = MyCommon.NZ(row.Item("MSNetCode"), "")
            Dim MLLanguageName As String = Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseTerm"), ""), MyCommon.GetAdminUser.LanguageID)
            Dim MLLanguageID As Integer = MyCommon.NZ(row.Item("LanguageID"), 0)
              
              
            If (MultiLanguageEnabled = True) Or (MultiLanguageEnabled = False AndAlso MyCommon.NZ(row.Item("LanguageID"), 0) = DefaultLanguageID) Then
              Send("<label for=""t" & t & "_line1_" & MLLanguageID & """>" & MLLanguageName & IIf(MLLanguageID = CustomerFacingLangID, " (" & Copient.PhraseLib.Lookup("term.default", MyCommon.GetAdminUser.LanguageID) & ")", "") & ":</label><br />")
              ' Line 1 input
              Dim Line1Raw As String = ""
              Sendb("&nbsp;<input type=""text"" id=""t" & t & "_line1_" & MLLanguageID & """ name=""t" & t & "_line1_" & MLLanguageID & """ onfocus=""srcElement=this"" style=""font-family:Courier; width:200px;"" maxlength=""20""")
              If TierRecordDT.Rows.Count = 0 Then
                Send(" value=""""" & DisabledAttribute & " />")
              Else
                If l = 1 Then
                  Line1Raw = MyCommon.NZ(TierRecordDT.Rows(0).Item("Line1"), "")
                Else
                  Line1Raw = MyCommon.NZ(row.Item("Line1"), "")
                End If
                Send(" value=""" & IIf(Line1Raw = "", "", Line1Raw.Replace("""", "&quot;")) & """" & DisabledAttribute & " />")
              End If
              Send("<br />")
              ' Line 2 input
              Dim Line2Raw As String = ""
              If TierRecordDT.Rows.Count > 0 Then
                ' See if the line has a tag; if so, get its position and split the line into two strings
                If l = 1 Then
                  Line2Raw = MyCommon.NZ(TierRecordDT.Rows(0).Item("Line2"), "")
                Else
                  Line2Raw = MyCommon.NZ(row.Item("Line2"), "")
                End If
                LineLength = Len(Line2Raw)
                TagStart = InStr(Line2Raw, "|")
                If TagStart > 0 Then
                  HasTag = True
                  Line2Text = Left(Line2Raw, TagStart - 1)
                  Line2Tag = Right(Line2Raw, (LineLength - TagStart) + 1)
                End If
              End If
              If HasTag Then
                Sendb("&nbsp;<input type=""text"" id=""t" & t & "_line2_" & MLLanguageID & """ name=""t" & t & "_line2_" & MLLanguageID & """ onfocus=""srcElement=this"" style=""font-family:Courier; width:100px;"" maxlength=""10""")
                If TierRecordDT.Rows.Count = 0 Then
                  Sendb(" value=""""" & DisabledAttribute & " />")
                Else
                  Sendb(" value=""" & Line2Text.Replace("""", "&quot;") & """" & DisabledAttribute & " />")
                End If
              Else
                Sendb("&nbsp;<input type=""text"" id=""t" & t & "_line2_" & MLLanguageID & """ name=""t" & t & "_line2_" & MLLanguageID & """ onfocus=""srcElement=this"" style=""font-family:Courier; width:200px;"" maxlength=""20""")
                If TierRecordDT.Rows.Count = 0 Then
                  Sendb(" value=""""" & DisabledAttribute & " />")
                Else
                  Sendb(" value=""" & IIf(Line2Raw = "", "", Line2Raw.Replace("""", "&quot;")) & """" & DisabledAttribute & " />")
                End If
              End If
              Send("&nbsp;<span id=""t" & t & "_line2tagdisplay_" & MLLanguageID & """ style=""position:relative;" & IIf(HasTag, "", "display:none;") & """>" & Line2Tag & "</span>")
              Send("<input type=""hidden"" id=""t" & t & "_line2tag_" & MLLanguageID & """ name=""t" & t & "_line2tag_" & MLLanguageID & """ value=""" & Line2Tag & """ />")
              Send("<button type=""button"" id=""t" & t & "_droptag_" & MLLanguageID & """ name=""t" & t & "_droptag_" & MLLanguageID & """ style=""color:#ff0000;font-size:8px;font-weight:bold;height:18px;width:18px;" & IIf(HasTag, "", "display:none;") & """ onclick=""javascript:dropTag(" & t & "," & MLLanguageID & ");"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """>X</button><br />")
              Send("<br class=""half"" />")
              HasTag = False
              LineLength = 0
              TagStart = 0
              Line2Text = ""
              Line2Tag = ""
            End If
            InstalledLanguages = InstalledLanguages & MyCommon.NZ(row.Item("LanguageID"), "")
            If (LanguagesDT.Rows.Count > l) Then
              InstalledLanguages = InstalledLanguages & ","
            End If
            l += 1
          Next
          Send("<input type=""hidden"" id=""InstalledLanguages"" name=""InstalledLanguages"" value=""" & InstalledLanguages & """ />")
          If MultiLanguageEnabled And TierLevels > 1 And t < TierLevels Then
            Send("<hr />")
          End If
        Next
        Send("</div>")
      %>
    </div>
  </div>
  <div id="gutter">
  </div>
  <div id="column2">
    <div class="box" id="display">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.display", LanguageID))%>
        </span>
      </h2>
      <%
        'Limit display setting to the hidden system option
        If TierLevels > 1 AndAlso MyCommon.Fetch_CPE_SystemOption(112) = "0" Then
          Send("<input type=""radio"" id=""DisplayImmediate"" name=""DisplayImmediateGrp"" value=""1"" disabled=""disabled"" " & DisabledAttribute & " /><label for=""DisplayImmediate"">" & Copient.PhraseLib.Lookup("reward.displayimmediately", LanguageID) & "</label>")
          Send("<input type=""radio"" id=""DisplayAfter"" name=""DisplayImmediateGrp"" value=""0"" checked=""checked""" & DisabledAttribute & " /><label for=""DisplayAfter"">" & Copient.PhraseLib.Lookup("reward.displayeos", LanguageID) & "</label>")
        Else
          If MessageID > 0 Then
            MyCommon.QueryStr = "select DisplayImmediate from CPE_CashierMessageTiers where TierLevel=1 and MessageID=" & MessageID
            dt = MyCommon.LRT_Select()
            If dt.Rows.Count > 0 Then
              Send("<input type=""radio"" id=""DisplayImmediate"" name=""DisplayImmediateGrp"" value=""1"" " & IIf(dt.Rows(0).Item("DisplayImmediate") = "True", " checked=""checked"" ", " ") & DisabledAttribute & " /><label for=""DisplayImmediate"">" & Copient.PhraseLib.Lookup("reward.displayimmediately", LanguageID) & "</label>")
              Send("<input type=""radio"" id=""DisplayAfter"" name=""DisplayImmediateGrp"" value=""0"" " & IIf(dt.Rows(0).Item("DisplayImmediate") = "False", " checked=""checked"" ", " ") & DisabledAttribute & " /><label for=""DisplayAfter"">" & Copient.PhraseLib.Lookup("reward.displayeos", LanguageID) & "</label>")
            Else
              Send("<input type=""radio"" id=""DisplayImmediate"" name=""DisplayImmediateGrp"" value=""1"" " & DisabledAttribute & " /><label for=""DisplayImmediate"">" & Copient.PhraseLib.Lookup("reward.displayimmediately", LanguageID) & "</label>")
              Send("<input type=""radio"" id=""DisplayAfter"" name=""DisplayImmediateGrp"" value=""0"" checked=""checked"" " & DisabledAttribute & " /><label for=""DisplayAfter"">" & Copient.PhraseLib.Lookup("reward.displayeos", LanguageID) & "</label>")
            End If
          Else
            Send("<input type=""radio"" id=""DisplayImmediate"" name=""DisplayImmediateGrp"" value=""1"" " & DisabledAttribute & " /><label for=""DisplayImmediate"">" & Copient.PhraseLib.Lookup("reward.displayimmediately", LanguageID) & "</label>")
            Send("<input type=""radio"" id=""DisplayAfter"" name=""DisplayImmediateGrp"" value=""0"" checked=""checked"" " & DisabledAttribute & " /><label for=""DisplayAfter"">" & Copient.PhraseLib.Lookup("reward.displayeos", LanguageID) & "</label>")
          End If
        End If
        '!TO DO
        'Display Selector
        'MyCommon.QueryStr = "select DisplayImmediate from CPE_CashierMessageTiers where TierLevel=1 and MessageID=" & MessageID
        'dt = MyCommon.LRT_Select()
        'If dt.Rows.Count > 0 Then
        '  Send("<input type=""radio"" id=""DisplayImmediate"" name=""DisplayImmediateGrp"" value=""1"" " & IIf(dt.Rows(0).Item("DisplayImmediate") = "True", " checked=""checked"" ", " ") & DisabledAttribute & " /><label for=""DisplayImmediate"">" & Copient.PhraseLib.Lookup("reward.displayimmediately", LanguageID) & "</label>")
        '  Send("<input type=""radio"" id=""DisplayAfter"" name=""DisplayImmediateGrp"" value=""0"" " & IIf(dt.Rows(0).Item("DisplayImmediate") = "False", " checked=""checked"" ", " ") & DisabledAttribute & " /><label for=""DisplayAfter"">" & Copient.PhraseLib.Lookup("reward.displayeos", LanguageID) & "</label>")
        'Else
        '  Send("<input type=""radio"" id=""DisplayImmediate"" name=""DisplayImmediateGrp"" value=""1""" & DisabledAttribute & " /><label for=""DisplayImmediate"">" & Copient.PhraseLib.Lookup("reward.displayimmediately", LanguageID) & "</label>")
        '  Send("<input type=""radio"" id=""DisplayAfter"" name=""DisplayImmediateGrp"" value=""0"" checked=""checked""" & DisabledAttribute & " /><label for=""DisplayAfter"">" & Copient.PhraseLib.Lookup("reward.displayeos", LanguageID) & "</label>")
        'End If
          
        'If dt.Rows.Count > 0 Then
        '  Send("<input type=""checkbox"" id=""DisplayImmediate"" name=""DisplayImmediate"" " & IIf(dt.Rows(0).Item("DisplayImmediate") = "True", " checked=""checked"" ", " ") & " /><label for=""DisplayImmediate"">" & Copient.PhraseLib.Lookup("reward.displayimmediately", LanguageID) & "</label>")
        'Else
        '  Send("<input type=""checkbox"" id=""DisplayImmediate"" name=""DisplayImmediate"" /><label for=""DisplayImmediate"">" & Copient.PhraseLib.Lookup("reward.displayimmediately", LanguageID) & "</label>")
        'End If
        Send("<br />")
      %>
    </div>
    <%
      If DisabledAttribute <> "" Then
        bProcessTags = False
      End If
        
      If bProcessTags Then
        Send("<div class=""box"" id=""tags"" >")
      Else
        Send("<div class=""box"" id=""tags"" style=""display: none;"" >")
      End If
    %>
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
</div>
<div id="dialogbox">
  <div id="svselector" style="display: none;">
    <div id="svTag">
    </div>
    <br />
    <input type="hidden" name="svTagName" id="svTagName" value="Svbal" />
    <b>
      <% Sendb(Copient.PhraseLib.Lookup("offer-rew.selectsv", LanguageID) & ":")%>
    </b>
    <br />
    <input type="radio" id="functionradio1a" name="functionradio1" checked="checked" /><label
      for="functionradio1a"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
    <input type="radio" id="functionradio1b" name="functionradio1" /><label for="functionradio1b"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
    <input type="text" class="medium" id="functioninput1" name="functioninput1" onkeyup="handleKeyUp1(200);"
      value="" /><br />
    <select onclick="handleSelectClick(1);" id="functionselect1" name="functionselect1"
      size="10" style="width: 220px;">
      <%
        MyCommon.QueryStr = "Select SVProgramID, Name from StoredValuePrograms where Deleted=0 order by Name;"
        rst = MyCommon.LRT_Select
        For Each row In rst.Rows
          Send("<option value=""" & MyCommon.NZ(row.Item("SVProgramID"), 0) & """>" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
        Next
      %>
    </select>
    <br />
    <br />
    <input type="button" id="close" name="close" value="<% Sendb(Copient.PhraseLib.Lookup("term.close", LanguageID))%>"
      onclick="javascript:showDialogSpan(false, 1);" />
  </div>
  <div id="pointselector" style="display: none;">
    <div id="ptTag">
    </div>
    <br />
    <input type="hidden" name="ptTagName" id="ptTagName" value="Ptbal" />
    <b>
      <% Sendb(Copient.PhraseLib.Lookup("offer-rew.selectpoints", LanguageID) & ":")%>
    </b>
    <br />
    <input type="radio" id="functionradio2a" name="functionradio2" checked="checked" /><label
      for="functionradio2a"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
    <input type="radio" id="functionradio2b" name="functionradio2" /><label for="functionradio2b"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
    <input type="text" class="medium" id="functioninput2" name="functioninput2" onkeyup="handleKeyUp2(200);"
      value="" /><br />
    <select onclick="handleSelectClick(2);" id="functionselect2" name="functionselect2"
      size="10" style="width: 220px;">
      <%
        MyCommon.QueryStr = "Select ProgramID, ProgramName from PointsPrograms where Deleted=0 order by ProgramName;"
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
</div>
</div>
</form>
<script runat="server">
  Function Create_Message(ByVal OfferID As String, ByVal Line1 As String, ByVal Line2 As String, ByVal Line2Tag As String, ByVal Phase As Integer, ByVal TpROID As Integer, ByRef DeliverableID As Long) As Boolean
    Dim MyCommon As New Copient.CommonInc
    Dim Status As Integer = 0
    
    Try
      MyCommon.QueryStr = "dbo.pa_CPE_AddCashierMessage"
      MyCommon.Open_LogixRT()
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt, 4).Value = OfferID
      MyCommon.LRTsp.Parameters.Add("@TpROID", SqlDbType.Int, 4).Value = TpROID
      MyCommon.LRTsp.Parameters.Add("@Phase", SqlDbType.Int, 4).Value = Phase
      MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.Int, 4).Direction = ParameterDirection.Output
      MyCommon.LRTsp.Parameters.Add("@DeliverableID", SqlDbType.Int, 4).Direction = ParameterDirection.Output
      MyCommon.LRTsp.ExecuteNonQuery()
      Status = MyCommon.LRTsp.Parameters("@Status").Value
      DeliverableID = MyCommon.LRTsp.Parameters("@DeliverableID").Value
      MyCommon.Close_LRTsp()
    Catch ex As Exception
      Status = -1
      Send(ex.ToString)
    Finally
      MyCommon.Close_LogixRT()
      MyCommon = Nothing
    End Try
    
    Return (Status = 0)
  End Function
  
  Sub Create_MessageTiers(ByVal MessageID As Long, ByVal DisplayImmediate As Integer, ByVal TierLevel As Long, Optional ByVal DefaultLanguageID As Integer = 1)
    Dim MyCommon As New Copient.CommonInc
    Dim Localization As Copient.Localization
    Dim MLI As New Copient.Localization.MultiLanguageRec
    
    Dim Line1 As String = Request.QueryString("t" & TierLevel & "_line1_" & DefaultLanguageID)
    Dim Line2 As String = Request.QueryString("t" & TierLevel & "_line2_" & DefaultLanguageID)
    Dim Line2Tag As String = Request.QueryString("t" & TierLevel & "_line2tag_" & DefaultLanguageID)
    Dim Beep As Integer = MyCommon.Extract_Val(Request.QueryString("t" & TierLevel & "_beep"))
    Dim BeepDuration As Integer = MyCommon.Extract_Val(Request.QueryString("t" & TierLevel & "_beepduration"))
    
    Dim Line1Clean As String = ""
    Dim Line2Clean As String = ""
    Dim PKID As Integer = 0
    
    Localization = New Copient.Localization(MyCommon)
    MyCommon.QueryStr = "dbo.pa_CPE_AddCashierMessageTiers"
    
    MyCommon.Open_LogixRT()
    MyCommon.Open_LRTsp()
    
    If Line1 <> "" Then Line1Clean = Replace(Line1, "|", "")
    If Line1 <> "" OrElse Line2 <> "" Then Line2Clean = Replace(Line2, "|", "") & Line2Tag
    
    MyCommon.LRTsp.Parameters.Add("@MessageID", SqlDbType.Int, 4).Value = MessageID
    MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int, 4).Value = TierLevel
    MyCommon.LRTsp.Parameters.Add("@Line1", SqlDbType.NVarChar, 30).Value = Line1Clean
    MyCommon.LRTsp.Parameters.Add("@Line2", SqlDbType.NVarChar, 30).Value = Line2Clean
    MyCommon.LRTsp.Parameters.Add("@Beep", SqlDbType.Int, 4).Value = Beep
    MyCommon.LRTsp.Parameters.Add("@BeepDuration", SqlDbType.Int, 4).Value = BeepDuration
    MyCommon.LRTsp.Parameters.Add("@DisplayImmediate", SqlDbType.Bit).Value = DisplayImmediate
    MyCommon.LRTsp.Parameters.Add("@PKID", SqlDbType.Int).Direction = ParameterDirection.Output
    MyCommon.LRTsp.ExecuteNonQuery()
    PKID = MyCommon.LRTsp.Parameters("@PKID").Value
    MyCommon.Close_LRTsp()
    MyCommon.Close_LogixRT()
    
    'Save multilanguage values
       ' If (MyCommon.Fetch_SystemOption(124) = "1") Then
            Dim LanguagesDT As DataTable
            Dim row As DataRow
            MyCommon.Open_LogixRT()
            MyCommon.QueryStr = "SELECT LanguageID FROM Languages AS L " & _
                                "WHERE L.LanguageID in (SELECT TLV.LanguageID FROM TransLanguagesRF_CPE AS TLV) " & _
                                "ORDER BY CASE WHEN L.LanguageID=" & DefaultLanguageID & " THEN 1 ELSE 2 END, L.LanguageID;"
            LanguagesDT = MyCommon.LRT_Select
            For Each row In LanguagesDT.Rows
                Line1Clean = ""
                Line2Clean = ""
                Line1 = Request.QueryString("t" & TierLevel & "_line1_" & MyCommon.NZ(row.Item("LanguageID"), 0))
                Line2 = Request.QueryString("t" & TierLevel & "_line2_" & MyCommon.NZ(row.Item("LanguageID"), 0))
                Line2Tag = Request.QueryString("t" & TierLevel & "_line2tag_" & MyCommon.NZ(row.Item("LanguageID"), 0))
                If Line1 <> "" Then Line1Clean = Replace(Line1, "|", "")
                If Line1 <> "" OrElse Line2 <> "" Then Line2Clean = Replace(Line2, "|", "") & Line2Tag
                If Line1Clean <> "" OrElse Line2Clean <> "" Then
                    MyCommon.QueryStr = "INSERT INTO CPE_CashierMsgTranslations (CashierMsgTierID, LanguageID, Line1, Line2) " & _
                                        "VALUES (" & PKID & ", " & row.Item("LanguageID") & ", N'" & MyCommon.Parse_Quotes(Line1Clean) & "', N'" & MyCommon.Parse_Quotes(Line2Clean) & "');"
                    MyCommon.LRT_Execute()
                End If
            Next
            MyCommon.Close_LogixRT()
            MyCommon = Nothing
            'End If
  End Sub
</script>
<script type="text/javascript">
<% If (CloseAfterSave) Then %>
     window.close();
<% Else %>
     xmlhttpPost("CashierMessageFeeds.aspx");
<% End If %>
</script>
<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd("mainform", "t1_line1_" & DefaultLanguageID)
  Logix = Nothing
  MyCommon = Nothing
%>
