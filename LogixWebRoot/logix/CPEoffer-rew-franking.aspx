<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: CPEoffer-rew-franking.aspx 
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
  Dim rst2 As DataTable
  Dim row As DataRow
  Dim CloseAfterSave As Boolean = False
  Dim IsTemplate As Boolean = False
  Dim FromTemplate As Boolean = False
  Dim disallow_edit As Boolean = False
  Dim bIsErrorMsg As Boolean = False
  Dim MsgAdded As Boolean = False
  Dim Line1 As String = ""
  Dim Line2 As String = ""
  Dim Beep As Integer
  Dim BeepDuration As Integer
  Dim OpenDrawer As Integer
  Dim OpenChecked As String = ""
  Dim ManagerOverride As Integer
  Dim OverrideChecked As String = ""
  Dim FrankFlag As Integer
  Dim IsTemplateVal As String = ""
  Dim DisabledAttribute As String = ""
  Dim FrankingText As String = ""
  Dim Name As String = ""
  Dim OfferID As Integer
  Dim BeepDurDisplay As String = "none"
  Dim RewardID As Integer
  Dim DeliverableID As Integer
  Dim FrankID As Integer
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim EngineID As Integer = 2
  Dim BannersEnabled As Boolean = True
  Dim ProtoBuild As Integer = 0
  Dim TierLevels As Integer = 1
  Dim t As Integer = 1
  Dim ValidTiers As Boolean = True
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "CPEoffer-rew-franking.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  ProtoBuild = MyCommon.Fetch_CPE_SystemOption(59)
  
  OfferID = Request.QueryString("OfferID")
  RewardID = MyCommon.Extract_Val(Request.QueryString("RewardID"))
  DeliverableID = MyCommon.Extract_Val(Request.QueryString("DeliverableID"))
  FrankID = MyCommon.Extract_Val(Request.QueryString("FrankID"))
  If (Request.QueryString("EngineID") <> "") Then
    EngineID = MyCommon.Extract_Val(Request.QueryString("EngineID"))
  Else
    MyCommon.QueryStr = "select EngineID from OfferIDs where OfferID=" & OfferID & ";"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
      EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), 0)
    End If
  End If
  
  MyCommon.QueryStr = "select TierLevels from CPE_RewardOptions with (NoLock) where RewardOptionID=" & RewardID & ";"
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    TierLevels = MyCommon.NZ(rst.Rows(0).Item("TierLevels"), 1)
  End If
  
  MyCommon.QueryStr = "select IncentiveName,IsTemplate,FromTemplate from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID
  rst = MyCommon.LRT_Select
  If (rst.Rows.Count > 0) Then
    Name = MyCommon.NZ(rst.Rows(0).Item("IncentiveName"), "")
    IsTemplate = MyCommon.NZ(rst.Rows(0).Item("IsTemplate"), False)
    FromTemplate = MyCommon.NZ(rst.Rows(0).Item("FromTemplate"), False)
  End If
  IsTemplateVal = IIf(IsTemplate, "IsTemplate", "Not")
  
  If (Request.QueryString("save") <> "" AndAlso DeliverableID = 0) Then
    ' Create a new franking message
    MsgAdded = Create_Message(OfferID, DeliverableID)
    If (MsgAdded AndAlso DeliverableID > 0) Then
      ' Get the new frank ID
      MyCommon.QueryStr = "select OutputID from CPE_Deliverables with (NoLock) where DeliverableID=" & DeliverableID & ";"
      rst = MyCommon.LRT_Select()
      If rst.Rows.Count > 0 Then
        FrankID = MyCommon.NZ(rst.Rows(0).Item("OutputID"), 0)
      End If
      ' Delete any existing tier values for this message from the tiers table, then insert new values
      MyCommon.QueryStr = "delete from CPE_FrankingMessageTiers with (RowLock) where FrankID in (0, " & FrankID & ");"
      MyCommon.LRT_Execute()
      ' Insert tier values
      t = 1
      For t = 1 To TierLevels
        If Request.QueryString("t" & t & "_openDrawer") <> "" Then
          OpenDrawer = MyCommon.Extract_Val(Request.QueryString("t" & t & "_openDrawer"))
        Else
          OpenDrawer = 0
        End If
        If Request.QueryString("t" & t & "_managerOverride") <> "" Then
          ManagerOverride = MyCommon.Extract_Val(Request.QueryString("t" & t & "_managerOverride"))
        Else
          ManagerOverride = 0
        End If
        FrankFlag = MyCommon.Extract_Val(Request.QueryString("t" & t & "_frankFlag"))
        FrankingText = Left(Trim(Request.QueryString("t" & t & "_frankingText")), 38)
        Line1 = Left(Trim(Request.QueryString("t" & t & "_line1")), 20)
        Line2 = Left(Trim(Request.QueryString("t" & t & "_line2")), 20)
        Beep = MyCommon.Extract_Val(Request.QueryString("t" & t & "_beepType"))
        BeepDuration = MyCommon.Extract_Val(Request.QueryString("t" & t & "_beepDuration"))
        Create_MessageTiers(FrankID, t, OpenDrawer, ManagerOverride, FrankFlag, FrankingText, Line1, Line2, Beep, BeepDuration)
      Next
      ' Update the CPE_Incentives table
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 " & _
                          "where IncentiveID=" & OfferID & ";"
      MyCommon.LRT_Execute()
    End If
    infoMessage = IIf((MsgAdded), Copient.PhraseLib.Lookup("CPE-rew-fmsg.created", LanguageID) & OfferID, Copient.PhraseLib.Lookup("CPE-rew-fmsg.error", LanguageID))
    CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Reward.createdfmsg", LanguageID))
  ElseIf (Request.QueryString("save") <> "" AndAlso DeliverableID > 0) Then
    ' Update an existing franking message
    MyCommon.QueryStr = "update CPE_FrankingMessages with (RowLock) set LastUpdate=getDate() " & _
                        "where FrankID=" & FrankID & ";"
    MyCommon.LRT_Execute()
    ' Delete any existing tier values for this message from the tiers table, then insert new values
    MyCommon.QueryStr = "delete from CPE_FrankingMessageTiers with (RowLock) where FrankID in (0, " & FrankID & ");"
    MyCommon.LRT_Execute()
    ' Insert tier values
    t = 1
    For t = 1 To TierLevels
      If Request.QueryString("t" & t & "_openDrawer") <> "" Then
        OpenDrawer = MyCommon.Extract_Val(Request.QueryString("t" & t & "_openDrawer"))
      Else
        OpenDrawer = 0
      End If
      If Request.QueryString("t" & t & "_managerOverride") <> "" Then
        ManagerOverride = MyCommon.Extract_Val(Request.QueryString("t" & t & "_managerOverride"))
      Else
        ManagerOverride = 0
      End If
      FrankFlag = MyCommon.Extract_Val(Request.QueryString("t" & t & "_frankFlag"))
      FrankingText = Left(Trim(Request.QueryString("t" & t & "_frankingText")), 38)
      Line1 = Left(Trim(Request.QueryString("t" & t & "_line1")), 20)
      Line2 = Left(Trim(Request.QueryString("t" & t & "_line2")), 20)
      Beep = MyCommon.Extract_Val(Request.QueryString("t" & t & "_beepType"))
      BeepDuration = MyCommon.Extract_Val(Request.QueryString("t" & t & "_beepDuration"))
      Create_MessageTiers(FrankID, t, OpenDrawer, ManagerOverride, FrankFlag, FrankingText, Line1, Line2, Beep, BeepDuration)
    Next
    ' Update the CPE_Incentives table
    MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 " & _
                        "where IncentiveID=" & OfferID & ";"
    MyCommon.LRT_Execute()
    infoMessage = Copient.PhraseLib.Lookup("CPE-rew-fmsg.edit", LanguageID) & " " & OfferID
    CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Reward.editfmsg", LanguageID))
  End If
  
  MyCommon.QueryStr = "select FM.FrankID, FMT.OpenDrawer, FMT.ManagerOverride, FMT.FrankFlag, FMT.FrankingText, FMT.Line1, FMT.Line2, FMT.Beep, FMT.BeepDuration " & _
                      "from CPE_Deliverables D with (NoLock) " & _
                      "inner join CPE_FrankingMessages FM with (NoLock) on D.OutputID=FM.FrankID " & _
                      "left join CPE_FrankingMessageTiers FMT with (NoLock) on FMT.FrankID=FM.FrankID " & _
                      "where D.RewardOptionID=" & RewardID & " and D.DeliverableID=" & DeliverableID & " and DeliverableTypeID=10;"
  rst = MyCommon.LRT_Select
  If (rst.Rows.Count > 0) Then
    FrankID = MyCommon.NZ(rst.Rows(0).Item("FrankID"), 0)
    OpenDrawer = MyCommon.NZ(rst.Rows(0).Item("OpenDrawer"), 0)
    Beep = MyCommon.NZ(rst.Rows(0).Item("Beep"), 0)
    BeepDuration = MyCommon.NZ(rst.Rows(0).Item("BeepDuration"), 0)
    Line1 = MyCommon.NZ(rst.Rows(0).Item("Line1"), "")
    Line2 = MyCommon.NZ(rst.Rows(0).Item("Line2"), "")
    BeepDurDisplay = IIf(Beep = 3, "block", "none")
    OpenChecked = IIf(OpenDrawer = 0, "", " checked=""checked""")
    ManagerOverride = MyCommon.NZ(rst.Rows(0).Item("ManagerOverride"), 0)
    OverrideChecked = IIf(ManagerOverride = 0, "", " checked=""checked""")
    FrankFlag = MyCommon.NZ(rst.Rows(0).Item("FrankFlag"), 2)
    FrankingText = MyCommon.NZ(rst.Rows(0).Item("FrankingText"), "")
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
      disallow_edit = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit"), False)
    Else
      disallow_edit = False
    End If
  End If
  
  If Not isTemplate Then
    DisabledAttribute = IIf(Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Edit), "", " disabled=""disabled""")
  Else
    DisabledAttribute = IIf(Logix.UserRoles.EditTemplates, "", " disabled=""disabled""")
  End If
  
  Send_HeadBegin("term.offer", "term.fmsgreward", OfferID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
  Send("<script type=""text/javascript"">")
  Send("function ChangeParentDocument() { ")
  If (EngineID = 6) Then
    Send("  opener.location = 'CAM/CAM-offer-rew.aspx?OfferID=" & OfferID & "'; ")
  Else
    Send("  opener.location = 'CPEoffer-rew.aspx?OfferID=" & OfferID & "'; ")
  End If
  Send("} ")
  Send("</script>")
%>
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
  var elemType = document.getElementById("t" + t + "_beepType");
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
  Send("  return bRetVal;")
  Send("}")
  Send("")
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
<form action="CPEoffer-rew-franking.aspx" id="mainform" name="mainform" onsubmit="return validateEntries();">
  <div id="intro">
    <input type="hidden" id="OfferID" name="OfferID" value="<% Sendb(OfferID) %>" />
    <input type="hidden" id="Name" name="Name" value="<% Sendb(Name) %>" />
    <input type="hidden" id="RewardID" name="RewardID" value="<% Sendb(RewardID) %>" />
    <input type="hidden" id="DeliverableID" name="DeliverableID" value="<% Sendb(DeliverableID) %>" />
    <input type="hidden" id="FrankID" name="FrankID" value="<% Sendb(FrankID) %>" />
    <input type="hidden" id="IsTemplate" name="IsTemplate" value="<% Sendb(IsTemplateVal)%>" />
    <%
      If (IsTemplate) Then
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.fmsgreward", LanguageID), VbStrConv.Lowercase) & "</h1>")
      Else
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.fmsgreward", LanguageID), VbStrConv.Lowercase) & "</h1>")
      End If
    %>
    <div id="controls">
      <% If (IsTemplate) Then%>
      <span class="temp">
        <input type="checkbox" class="tempcheck" id="Disallow_Edit" name="Disallow_Edit"<% if(disallow_edit)then Sendb(" checked=""checked""") %> />
        <label for="Disallow_Edit"><% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
      </span>
      <% End If%>
      <%
        If Not IsTemplate Then
          If (Logix.UserRoles.EditOffer And Not (FromTemplate And disallow_edit)) Then
            If FrankID = 0 Then
              Send_Save(" onclick=""this.style.visibility='hidden';""")
            Else
              Send_Save()
            End If
          End If
        Else
          If (Logix.UserRoles.EditTemplates) Then
            If FrankID = 0 Then
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
      <%
        MyCommon.QueryStr = "select TierLevel, OpenDrawer, ManagerOverride, FrankFlag, FrankingText, Line1, Line2, Beep, BeepDuration " & _
                            "from CPE_FrankingMessageTiers with (NoLock) " & _
                            "where FrankID=" & FrankID & " order by TierLevel;"
        rst = MyCommon.LRT_Select
        For t = 1 To TierLevels
          Send("<div class=""box"" id=""t" & t & "_fmessage"">")
          Send("  <h2>")
          Send("    <span>")
          If TierLevels > 1 Then
            Send("      " & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t)
          Else
            Send("      " & Copient.PhraseLib.Lookup("term.details", LanguageID))
          End If
          Send("    </span>")
          Send("  </h2>")
          ' Franking message
          Send("<label for=""t" & t & "_frankingText"">" & Copient.PhraseLib.Lookup("term.frankingmessage", LanguageID) & ":</label><br />")
          Sendb("<input type=""text"" class=""longer"" id=""t" & t & "_frankingText"" name=""t" & t & "_frankingText"" maxlength=""38""")
          If (t > rst.Rows.Count) Or (FrankID = 0) Then
            Send(" value=""""" & DisabledAttribute & " /><br />")
          Else
            Send(" value=""" & IIf(MyCommon.NZ(rst.Rows(t - 1).Item("FrankingText"), "") = "", "", MyCommon.NZ(rst.Rows(t - 1).Item("FrankingText"), "").Replace("""", "&quot;")) & """" & DisabledAttribute & " /><br />")
          End If
          Send("<br class=""half"" />")
          Send("<select class=""longer"" id=""t" & t & "_frankFlag"" name=""t" & t & "_frankFlag""" & DisabledAttribute & ">")
          If FrankID = 0 Then
            Send("  <option value=""0"">")
            Send("    " & Copient.PhraseLib.Lookup("CPE-rew-fmsg.posdataonly", LanguageID))
            Send("  </option>")
            Send("  <option value=""1"">")
            Send("    " & Copient.PhraseLib.Lookup("CPE-rew-fmsg.applyfmsgonly", LanguageID))
            Send("  </option>")
            Send("  <option value=""2"">")
            Send("    " & Copient.PhraseLib.Lookup("CPE-rew-fmsg.applyfmsgpos", LanguageID))
            Send("  </option>")
          Else
            If t > rst.Rows.Count Then
              Send("  <option value=""0"">")
              Send("    " & Copient.PhraseLib.Lookup("CPE-rew-fmsg.posdataonly", LanguageID))
              Send("  </option>")
              Send("  <option value=""1"">")
              Send("    " & Copient.PhraseLib.Lookup("CPE-rew-fmsg.applyfmsgonly", LanguageID))
              Send("  </option>")
              Send("  <option value=""2"">")
              Send("    " & Copient.PhraseLib.Lookup("CPE-rew-fmsg.applyfmsgpos", LanguageID))
              Send("  </option>")
            Else
              Send("  <option value=""0""" & IIf(MyCommon.NZ(rst.Rows(t - 1).Item("FrankFlag"), 2) = 0, " selected=""selected""", "") & ">")
              Send("    " & Copient.PhraseLib.Lookup("CPE-rew-fmsg.posdataonly", LanguageID))
              Send("  </option>")
              Send("  <option value=""1""" & IIf(MyCommon.NZ(rst.Rows(t - 1).Item("FrankFlag"), 2) = 1, " selected=""selected""", "") & ">")
              Send("    " & Copient.PhraseLib.Lookup("CPE-rew-fmsg.applyfmsgonly", LanguageID))
              Send("  </option>")
              Send("  <option value=""2""" & IIf(MyCommon.NZ(rst.Rows(t - 1).Item("FrankFlag"), 2) = 2, " selected=""selected""", "") & ">")
              Send("    " & Copient.PhraseLib.Lookup("CPE-rew-fmsg.applyfmsgpos", LanguageID))
              Send("  </option>")
            End If
          End If
          Send("</select>")
          Send("<br />")
          Send("<br class=""half"" />")
          If ProtoBuild = 1 Then
            ' Cashier message
            Send("<label for=""t" & t & "_line1"">" & Copient.PhraseLib.Lookup("term.cashiermessage", LanguageID) & ":</label>")
            Send("<br />")
            Sendb("<input type=""text"" class=""mediumlong"" id=""t" & t & "_line1"" name=""t" & t & "_line1"" style=""font-family:Courier;"" maxlength=""20""")
            If (t > rst.Rows.Count) Or (FrankID = 0) Then
              Send(" value=""""" & DisabledAttribute & " /><br />")
            Else
              Send(" value=""" & IIf(MyCommon.NZ(rst.Rows(t - 1).Item("Line1"), "") = "", "", MyCommon.NZ(rst.Rows(t - 1).Item("Line1"), "").Replace("""", "&quot;")) & """" & DisabledAttribute & " /><br />")
            End If
            Sendb("<input type=""text"" class=""mediumlong"" id=""t" & t & "_line2"" name=""t" & t & "_line2"" style=""font-family:Courier;"" maxlength=""20""")
            If (t > rst.Rows.Count) Or (FrankID = 0) Then
              Send(" value=""""" & DisabledAttribute & " /><br />")
            Else
              Send(" value=""" & IIf(MyCommon.NZ(rst.Rows(t - 1).Item("Line2"), "") = "", "", MyCommon.NZ(rst.Rows(t - 1).Item("Line2"), "").Replace("""", "&quot;")) & """" & DisabledAttribute & " /><br />")
            End If
            Send("<br class=""half"" />")
            ' Beep
            Send("<label for=""t" & t & "_beepType"">" & Copient.PhraseLib.Lookup("term.beep", LanguageID) & ":</label>")
            Send("<br />")
            Send("<select id=""t" & t & "_beepType"" name=""t" & t & "_beepType"" style=""float:left;position:relative;"" onchange=""beepTypeChanged(this, " & t & ");""" & DisabledAttribute & ">")
            Dim BeepTypeId As Integer = 0
            Dim BeepDesc As String = ""
            MyCommon.QueryStr = "select BeepTypeID, PhraseID from BeepTypes BT with (NoLock)"
            rst2 = MyCommon.LRT_Select
            For Each row In rst2.Rows
              Sendb("  <option value=""" & MyCommon.NZ(row.Item("BeepTypeID"), 0) & """")
              If t <= rst.Rows.Count Then
                If (MyCommon.NZ(row.Item("BeepTypeID"), 0) = MyCommon.NZ(rst.Rows(t - 1).Item("Beep"), 0)) Then
                  Sendb(" selected=""selected""")
                End If
              End If
              Sendb(">" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID))
              Send("  </option>")
            Next
            Send("</select>")
            If t <= rst.Rows.Count Then
              BeepDurDisplay = IIf(MyCommon.NZ(rst.Rows(t - 1).Item("Beep"), 0) = 3, "inline", "none")
            End If
            Send("<span id=""t" & t & "_trDuration"" style=""display:" & BeepDurDisplay & "; float:left; position:relative;"">")
            Sendb("  :&nbsp;<input type=""text"" class=""shortest"" id=""t" & t & "_beepDuration"" name=""t" & t & "_beepDuration"" maxlength=""2""")
            If t > rst.Rows.Count Then
              Send(" value=""0""" & DisabledAttribute & " />")
            Else
              Send(" value=""" & MyCommon.NZ(rst.Rows(t - 1).Item("BeepDuration"), 0) & """" & DisabledAttribute & " />")
            End If
            Send("</span>")
            Send("<br clear=""left"" />")
            Send("<br class=""half"" />")
          End If
          If FrankID > 0 Then
            If t > rst.Rows.Count Then
              OpenChecked = ""
              OverrideChecked = ""
            Else
              OpenChecked = IIf(MyCommon.NZ(rst.Rows(t - 1).Item("OpenDrawer"), False) = True, " checked=""checked""", "")
              OverrideChecked = IIf(MyCommon.NZ(rst.Rows(t - 1).Item("ManagerOverride"), False) = True, " checked=""checked""", "")
            End If
          End If
          Send("<input type=""checkbox"" id=""t" & t & "_openDrawer"" name=""t" & t & "_openDrawer"" value=""1""" & OpenChecked & DisabledAttribute & " />")
          Send("<label for=""t" & t & "_openDrawer"">" & Copient.PhraseLib.Lookup("CPE-rew-fmsg.opencashdrawer", LanguageID) & "</label>")
          Send("<br />")
          Send("<input type=""checkbox"" id=""t" & t & "_managerOverride"" name=""t" & t & "_managerOverride"" value=""1""" & OverrideChecked & DisabledAttribute & " />")
          Send("<label for=""t" & t & "_managerOverride"">" & Copient.PhraseLib.Lookup("CPE-rew-fmsg.manageroverride", LanguageID) & "</label>")
          Send("<br />")
          Send("</div>")
        Next
      %>
      
    </div>
  </div>
</form>

<script runat="server">
  Function Create_Message(ByVal OfferID As String, ByRef DeliverableID As Long) As Boolean
    Dim MyCommon As New Copient.CommonInc
    Dim Status As Integer = 0
    
    Try
      MyCommon.QueryStr = "dbo.pa_CPE_AddFrankingMessage"
      MyCommon.Open_LogixRT()
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt, 4).Value = OfferID
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
  
  Sub Create_MessageTiers(ByVal FrankID As Long, ByVal TierLevel As Integer, ByVal OpenDrawer As Integer, ByVal ManagerOverride As Integer, _
                          ByVal FrankFlag As Integer, ByVal FrankingText As String, ByVal Line1 As String, ByVal Line2 As String, _
                          ByVal Beep As Integer, ByVal BeepDuration As Integer)
    Dim MyCommon As New Copient.CommonInc
    
    MyCommon.QueryStr = "dbo.pa_CPE_AddFrankingMessageTiers"
    
    MyCommon.Open_LogixRT()
    MyCommon.Open_LRTsp()
    
    MyCommon.LRTsp.Parameters.Add("@FrankID", SqlDbType.Int, 4).Value = FrankID
    MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int, 4).Value = TierLevel
    MyCommon.LRTsp.Parameters.Add("@OpenDrawer", SqlDbType.Bit).Value = OpenDrawer
    MyCommon.LRTsp.Parameters.Add("@ManagerOverride", SqlDbType.Bit).Value = ManagerOverride
    MyCommon.LRTsp.Parameters.Add("@FrankFlag", SqlDbType.Int).Value = FrankFlag
    MyCommon.LRTsp.Parameters.Add("@FrankingText", SqlDbType.NVarChar, 38).Value = FrankingText
    MyCommon.LRTsp.Parameters.Add("@Line1", SqlDbType.NVarChar, 38).Value = Line1
    MyCommon.LRTsp.Parameters.Add("@Line2", SqlDbType.NVarChar, 38).Value = Line2
    MyCommon.LRTsp.Parameters.Add("@Beep", SqlDbType.Int).Value = Beep
    MyCommon.LRTsp.Parameters.Add("@BeepDuration", SqlDbType.Int).Value = BeepDuration
    
    MyCommon.LRTsp.ExecuteNonQuery()
    
    MyCommon.Close_LRTsp()
    MyCommon.Close_LogixRT()
    MyCommon = Nothing
    
  End Sub
</script>

<script type="text/javascript">
<% If (CloseAfterSave) Then %>
    window.close();
<% End If %>
</script>

<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd("mainform", "t1_frankingText")
  Logix = Nothing
  MyCommon = Nothing
%>
