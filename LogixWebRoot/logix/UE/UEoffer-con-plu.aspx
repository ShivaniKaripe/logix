<%@ Page Language="vb" Debug="true" CodeFile="ue-cb.vb" Inherits="UECB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
    ' *****************************************************************************
    ' * FILENAME: UEoffer-con-plu.aspx 
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
    Dim MyOffer As New Copient.CPEOffer
    Dim rst As DataTable
    Dim row As DataRow
    Dim i As Decimal
    Dim OfferID As Long
    Dim OfferName As String
    Dim IncentivePLUID As Integer = 0
    'Dim PLU As Decimal = 0
    Dim PLUString As String = ""
    Dim PLUQty As Integer = 0
    Dim PaddedPLU As String = ""
    Dim IsTemplate As Boolean = False
    Dim DisallowEdit As Boolean = False
    Dim DisabledAttribute As String = ""
    Dim FromTemplate As Boolean = False
    Dim RequiredFromTemplate As Boolean = True
    Dim CloseAfterSave As Boolean = False
    Dim roid As Integer
    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False
    Dim EngineID As Integer = 2
    Dim BannersEnabled As Boolean = True
    Dim RangeBegin As Decimal = 0
    Dim RangeBeginString As String = ""
    Dim RangeEnd As Decimal = 0
    Dim RangeEndString As String = ""
    Dim Range As Decimal = 0
    Dim RangeLocked As Boolean = True
    Dim RangeUndefined As Boolean = False
    Dim MultipleOffers As Boolean = False
    Dim PerRedemption As Boolean = False
    Dim CashierMessage As Boolean = False
    Dim CashierMessageText As String = ""
    Dim ValidPLU As Boolean = False
    Dim IDLength As Integer = 0
    Dim ErrorMessage As String = ""
    Dim counter As Integer = 1
    Dim ExtRedemAuthorization As Boolean = False
    Dim sbInsertUpdate As New StringBuilder
    Dim ExtRedemAuthorizationEnabled As Boolean = MyCommon.Fetch_UE_SystemOption(172)
    Dim sbRetrieve As New StringBuilder

    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If

    Response.Expires = 0
    MyCommon.AppName = "UEoffer-con-plu.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)

    BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")

    Integer.TryParse(MyCommon.Fetch_SystemOption(52), IDLength)
    OfferID = Request.QueryString("OfferID")

    'This would redirect to summary page if offer is awaiting deployment or if offer is awaiting recommendations.
    CheckIfValidOffer(MyCommon, OfferID)

    Dim isTranslatedOffer As Boolean =MyCommon.IsTranslatedUEOffer(OfferID,  MyCommon)
    Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249)="1",True,False)

    Dim bEnableAdditionalLockoutRestrictionsOnOffers As Boolean = IIf(MyCommon.Fetch_SystemOption(260) = "1", True, False)
    Dim bOfferEditable As Boolean = MyCommon.IsOfferEditablePastLockOutPeriod(Logix.UserRoles.EditOfferPastLockoutPeriod, MyCommon, OfferID)

    If Request.QueryString("IncentivePLUID") <> "" Then
        IncentivePLUID = MyCommon.Extract_Val(Request.QueryString("IncentivePLUID"))
    End If
    If IncentivePLUID > 0 Then
        MyCommon.QueryStr = "select PLU as PLUString, PLUQuantity as PLUQty " &
                            "from CPE_IncentivePLUs with (NoLock) where IncentivePLUID=" & IncentivePLUID & ";"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
            'PLU = MyCommon.NZ(rst.Rows(0).Item("PLU"), 0)
            PLUString = MyCommon.NZ(rst.Rows(0).Item("PLUString"), "")
            PLUQty = MyCommon.NZ(rst.Rows(0).Item("PLUQty"), "")

        End If
    End If
    RangeLocked = IIf(MyCommon.Fetch_UE_SystemOption(95) = "1", False, True)
    MultipleOffers = IIf(MyCommon.Fetch_UE_SystemOption(94) = "1", True, False)
    If (MyCommon.Fetch_SystemOption(198) = "") OrElse (MyCommon.Fetch_SystemOption(199) = "") Then
        RangeUndefined = True
    End If
    If MyCommon.Fetch_SystemOption(198) <> "" Then
        If Not Decimal.TryParse(MyCommon.Fetch_SystemOption(198), RangeBegin) And RangeLocked Then
            infoMessage = Copient.PhraseLib.Lookup("plu.InvalidRange", LanguageID)
        End If
        RangeBeginString = MyCommon.Fetch_SystemOption(198).ToString.PadLeft(IDLength, "0")
    Else
        RangeBegin = 0
        RangeBeginString = RangeBegin.ToString.PadLeft(IDLength, "0")
    End If
    If MyCommon.Fetch_SystemOption(199) <> "" Then
        If Not Decimal.TryParse(MyCommon.Fetch_SystemOption(199), RangeEnd) And RangeLocked Then
            infoMessage = Copient.PhraseLib.Lookup("plu.InvalidRange", LanguageID)
        End If
        RangeEndString = MyCommon.Fetch_SystemOption(199).ToString.PadLeft(IDLength, "0")
    Else
        RangeEnd = 9
        RangeEndString = RangeEnd.ToString.PadLeft(IDLength, "9")
        RangeEnd = CDec(RangeEndString)
    End If

    Range = (RangeEnd - RangeBegin) + 1

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

    MyCommon.QueryStr = "select IncentiveID, IncentiveName as Name, IsTemplate, FromTemplate from CPE_Incentives as CPE with (NoLock) " & _
                        "where IncentiveID=" & Request.QueryString("OfferID") & ";"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
        OfferName = MyCommon.NZ(rst.Rows(0).Item("Name"), "")
        IsTemplate = MyCommon.NZ(rst.Rows(0).Item("IsTemplate"), False)
        FromTemplate = MyCommon.NZ(rst.Rows(0).Item("FromTemplate"), False)
    End If

    'Someone is saving
    If (Request.QueryString("save") <> "") Then

        ValidPLU = MyOffer.IsValidPLU(Request.QueryString("PLU"), roid, IncentivePLUID, IsTemplate, MyCommon, ErrorMessage)

        If ValidPLU Then
            Dim qty As Integer = 0
            If Not Integer.TryParse(Request.QueryString("pluqty"), qty) OrElse qty < 1 OrElse qty > 1000 Then
                ValidPLU = False
                ErrorMessage = Copient.PhraseLib.Lookup("error.invalid-triggercodeqty", LanguageID)
            End If
        End If
        If ValidPLU Then
            'If (IDLength > 0 And MyOffer.AllDigits(Request.QueryString("PLU"))) Then
            If (IDLength > 0) Then
                PaddedPLU = Left(MyCommon.Parse_Quotes(Trim(Request.QueryString("PLU"))), 26).PadLeft(IDLength, "0")
            Else
                PaddedPLU = Left(MyCommon.Parse_Quotes(Trim(Request.QueryString("PLU"))), 26)
            End If
            If MyCommon.Extract_Val(Request.QueryString("IncentivePLUID")) = 0 Then
                'It's a new condition, so we'll insert it
                If IsTemplate AndAlso ((MyCommon.Fetch_UE_SystemOption(94) = 0) OrElse (MyCommon.Fetch_UE_SystemOption(94) = 1 AndAlso MyCommon.Fetch_SystemOption(275) = 1 )) Then
                    sbInsertUpdate.Clear()
                    sbInsertUpdate.Append("insert into CPE_IncentivePLUs with (RowLock) (RewardOptionID, PLU, PerRedemption, CashierMessage, LastUpdate, DisallowEdit, RequiredFromTemplate, PLUQuantity")
                    sbInsertUpdate.Append(IIf(ExtRedemAuthorizationEnabled, ", ExternalRedemptionAuthorization)", ")"))
                    sbInsertUpdate.Append(" values (" & roid & ",''," & IIf(Request.QueryString("PerRedemption") = "1", "1", "0") & "," & IIf(Request.QueryString("CashierMessage") <> "", "1", "0") & ",getdate()," & IIf(Request.QueryString("DisallowEdit") <> "", "1", "0") & "," & IIf(Request.QueryString("RequiredFromTemplate") <> "", "1", "0") & "," & IIf(Request.QueryString("pluqty").ToString() <> "", Convert.ToInt16(Request.QueryString("pluqty")), 1) & "")
                    sbInsertUpdate.Append(IIf(ExtRedemAuthorizationEnabled, "," & IIf(Request.QueryString("ExtRedemAuthorization") <> "", "1", "0") & ");", ");"))
                Else
                    sbInsertUpdate.Clear()
                    sbInsertUpdate.Append("insert into CPE_IncentivePLUs with (RowLock) (RewardOptionID, PLU, PerRedemption, CashierMessage, LastUpdate, DisallowEdit, RequiredFromTemplate, PLUQuantity")
                    sbInsertUpdate.Append(IIf(ExtRedemAuthorizationEnabled, ", ExternalRedemptionAuthorization)", ")"))
                    sbInsertUpdate.Append(" values (" & roid & ",'" & PaddedPLU & "'," & IIf(Request.QueryString("PerRedemption") = "1", "1", "0") & "," & IIf(Request.QueryString("CashierMessage") <> "", "1", "0") & ",getdate()," & IIf(Request.QueryString("DisallowEdit") <> "", "1", "0") & "," & IIf(Request.QueryString("RequiredFromTemplate") <> "", "1", "0") & "," & IIf(Request.QueryString("pluqty").ToString() <> "", Convert.ToInt16(Request.QueryString("pluqty")), 1) & "")
                    sbInsertUpdate.Append(IIf(ExtRedemAuthorizationEnabled, "," & IIf(Request.QueryString("ExtRedemAuthorization") <> "", "1", "0") & ");", ");"))
                End If
                MyCommon.QueryStr = sbInsertUpdate.ToString()
                MyCommon.LRT_Execute()
                MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-plu-add", LanguageID))
            Else
                'It's an existing condition, so we'll update it
                sbInsertUpdate.Clear()
                sbInsertUpdate.Append("update CPE_IncentivePLUs set " &
                                   "  PLU='" & PaddedPLU & "', " &
                                   "  PerRedemption=" & IIf(Request.QueryString("PerRedemption") = "1", "1", "0") & ", " &
                                   "  CashierMessage=" & IIf(Request.QueryString("CashierMessage") <> "", "1", "0") & ", " &
                                   "  DisallowEdit=" & IIf(Request.QueryString("DisallowEdit") <> "", "1", "0") & ", " &
                                   "  RequiredFromTemplate=" & IIf(Request.QueryString("RequiredFromTemplate") <> "", "1", "0") & ", " &
                                   "  PLUQuantity = " & IIf(Request.QueryString("pluqty").ToString() <> "", Request.QueryString("pluqty"), 1))
                sbInsertUpdate.Append(IIf(ExtRedemAuthorizationEnabled, ", " & "ExternalRedemptionAuthorization=" & IIf(Request.QueryString("ExtRedemAuthorization") <> "", "1", "0") & " ", " "))
                sbInsertUpdate.Append("where IncentivePLUID=" & MyCommon.Extract_Val(Request.QueryString("IncentivePLUID")) & ";")
                MyCommon.QueryStr = sbInsertUpdate.ToString()
                MyCommon.LRT_Execute()
                MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-plu-edit", LanguageID))
            End If
            'Synchronize the PerRedemption setting to be the same in *all* PLU conditions for this offer
            MyCommon.QueryStr = "update CPE_IncentivePLUs set PerRedemption=" & IIf(Request.QueryString("PerRedemption") = "1", "1", "0") & " " & _
                                "where RewardOptionID=" & roid & ";"
            MyCommon.LRT_Execute()
            'Update CPE_Incentives
            MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
            MyCommon.LRT_Execute()
            ResetOfferApprovalStatus(OfferID)
        Else
            infoMessage = ErrorMessage
        End If
        CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1") AndAlso infoMessage = ""
    End If

    'No one clicked anything -- load details
    If IncentivePLUID > 0 Then
        sbRetrieve.Clear()
        sbRetrieve.Append("select PerRedemption, CashierMessage")
        sbRetrieve.Append(IIf(ExtRedemAuthorizationEnabled, ", ExternalRedemptionAuthorization", ""))
        sbRetrieve.Append(" from CPE_IncentivePLUs with (NoLock) where IncentivePLUID=" & IncentivePLUID & ";")
        MyCommon.QueryStr = sbRetrieve.ToString()
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
            PerRedemption = IIf(MyCommon.NZ(rst.Rows(0).Item("PerRedemption"), False), True, False)
            CashierMessage = IIf(MyCommon.NZ(rst.Rows(0).Item("CashierMessage"), False), True, False)
            If (ExtRedemAuthorizationEnabled) Then
                ExtRedemAuthorization = IIf(MyCommon.NZ(rst.Rows(0).Item("ExternalRedemptionAuthorization"), False), True, False)
            End If
        End If
    End If
    If (IsTemplate Or FromTemplate) Then
        'If a template, find its permissions
        MyCommon.QueryStr = "select DisallowEdit, RequiredFromTemplate from CPE_IncentivePLUs with (NoLock) " & _
                            "where RewardOptionID=" & roid & " and IncentivePLUID=" & IncentivePLUID & ";"
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
            DisallowEdit = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit"), False)
            RequiredFromTemplate = MyCommon.NZ(rst.Rows(0).Item("RequiredFromTemplate"), False)
        Else
            DisallowEdit = False
            'RequiredFromTemplate = False
        End If
    End If
    MyCommon.QueryStr = "select CM.MessageID, CMT.Line1, CMT.Line2 from CPE_CashierMessages as CM " & _
                        "inner join CPE_CashierMessageTiers as CMT with (NoLock) on CMT.MessageID=CM.MessageID " & _
                        "where CM.PLU=1 and CMT.TierLevel=1;"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
        CashierMessageText = MyCommon.NZ(rst.Rows(0).Item("Line1"), "")
        If MyCommon.NZ(rst.Rows(0).Item("Line2"), "") <> "" Then
            CashierMessageText &= "<br />" & MyCommon.NZ(rst.Rows(0).Item("Line2"), "")
        End If
    End If
    Dim m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
    If Not IsTemplate Then
        DisabledAttribute = IIf(Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And DisallowEdit), "", "disabled=""disabled""")
    Else
        DisabledAttribute = IIf(Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer, "", "disabled=""disabled""")
    End If

    Send_HeadBegin("term.offer", "term.triggercodecondition", OfferID)
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts()
%>
<script type="text/javascript">
  function ChangeParentDocument() {
  <%
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
  %>
  }
  
  function handleRequiredToggle(element) {
    if (element == 'RequiredFromTemplate') {
      if (document.getElementById("RequiredFromTemplate").checked == true) {
        document.getElementById("DisallowEdit").checked=false;
      }
    } else if (element == 'DisallowEdit') {
      if (document.getElementById("DisallowEdit").checked == true) {
        document.getElementById("RequiredFromTemplate").checked=false;
      }
    }
  }
  
  function selectPLU(idLen) {
    var PLUelem = document.getElementById("plu");
    var selElem = document.getElementById("selector");
    var qtyElem = document.getElementById("pluqty");
    
    if (PLUelem != null && selElem != null) {
        PLUelem.value = padLeft(selElem.options[selElem.selectedIndex].value, idLen);
        qtyElem.value = 1; 
    }
  }
  
  function padLeft(str, totalLength) {
    var pd = '';

    str = str.toString();
    if (totalLength > str.length) {
      for (var i=0; i < (totalLength-str.length); i++) {
        pd += '0';
      }      
    }
    
    return pd + str.toString();
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
    <input type="hidden" id="OfferID" name="OfferID" value="<% Sendb(OfferID) %>" />
    <input type="hidden" id="IncentivePLUID" name="IncentivePLUID" value="<% Sendb(IncentivePLUID) %>" />
    <input type="hidden" id="IsTemplate" name="IsTemplate" value="<% Sendb(IIf(IsTemplate, "IsTemplate", "Not")) %>" />
    <%
      If (IsTemplate) Then
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.triggercodecondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
      Else
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.triggercodecondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
      End If
    %>
    <div id="controls">
      <%
        If (IsTemplate AndAlso ((MyCommon.Fetch_UE_SystemOption(94) = 0) OrElse (MyCommon.Fetch_UE_SystemOption(94) = 1 AndAlso MyCommon.Fetch_SystemOption(275) = 1 ))) Then
          Send("<span class=""temp"">")
          DisallowEdit = False 'Hard-coding this to false because RequiredFromTemplate must be hard-coded to true; see below --hjw
          Send("  <input type=""checkbox"" class=""tempcheck"" id=""DisallowEdit"" name=""DisallowEdit"" " & IIf(DisallowEdit, " checked=""checked""", "") & " disabled=""disabled"" />")
          Send("  <label for=""DisallowEdit"">" & Copient.PhraseLib.Lookup("term.locked", LanguageID) & "</label>")
          Send("</span>")
        ElseIf (IsTemplate AndAlso (MyCommon.Fetch_UE_SystemOption(94) <> 0)) Then
          Send("<span class=""temp"">")
          Send("  <input type=""checkbox"" class=""tempcheck"" id=""DisallowEdit"" name=""DisallowEdit"" " & IIf(DisallowEdit, " checked=""checked""", "") & "  />")
          Send("  <label for=""DisallowEdit"">" & Copient.PhraseLib.Lookup("term.locked", LanguageID) & "</label>")
          Send("</span>")
        End If
          m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
        If((Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
          If Not IsTemplate Then
                  If (Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And DisallowEdit) And Not IsOfferWaitingForApproval(OfferID)) Then Send_Save()
          Else
                If (Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer) Then Send_Save()
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
    <div id="column1">
      <div class="box" id="code">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.code", LanguageID))%>
          </span>
          <%
            If (IsTemplate AndAlso ((MyCommon.Fetch_UE_SystemOption(94) = 0) OrElse (MyCommon.Fetch_UE_SystemOption(94) = 1 AndAlso MyCommon.Fetch_SystemOption(275) = 1 ))) Then
              Send("<span class=""tempRequire"">")
              RequiredFromTemplate = True 'Hard-coding to true because templates cannot include a PLU and must leave the PLU input field blank; see below --hjw
              Send("  <input type=""checkbox"" class=""tempcheck"" id=""RequiredFromTemplate"" name=""RequiredFromTemplate"" " & IIf(RequiredFromTemplate, " checked=""checked""", "") & " disabled=""disabled"" />")
              Send("  <label for=""require_pp"">")
              Send("    " & Copient.PhraseLib.Lookup("term.required", LanguageID))
              Send("  </label>")
              Send("</span>")
            ElseIf (IsTemplate AndAlso (MyCommon.Fetch_UE_SystemOption(94) = 1)) Then
              Send("<span class=""tempRequire"">")
              Send("  <input type=""checkbox"" class=""tempcheck"" id=""RequiredFromTemplate"" name=""RequiredFromTemplate"" " & IIf(RequiredFromTemplate, " checked=""checked""", "") & " />")
              Send("  <label for=""require_pp"">")
              Send("    " & Copient.PhraseLib.Lookup("term.required", LanguageID))
              Send("  </label>")
              Send("</span>")
            ElseIf (FromTemplate And RequiredFromTemplate) Then
              Send("<span class=""tempRequire"">*")
              Send("  " & Copient.PhraseLib.Lookup("term.required", LanguageID))
              Send("</span>")
            End If
          %>
        </h2>
        <%  'CLOUDSOL-2074:trigger code needs to be editable all the time
            'If (IncentivePLUID > 0) AndAlso (PLUString <> "") Then
            '  Send("<input type=""hidden"" id=""plu"" name=""plu"" value=""" & IIf(IncentivePLUID > 0, PLUString, "") & """ />")
            '  Send("<input type=""text"" id=""plu-static"" name=""plu-static"" style=""width:200px;"" value=""" & IIf(IncentivePLUID > 0, PLUString, "") & """ disabled=""disabled"" /><br />")
            'Else
            ' PLU input
            If (IsTemplate AndAlso ((MyCommon.Fetch_UE_SystemOption(94) = 0) OrElse (MyCommon.Fetch_UE_SystemOption(94) = 1 AndAlso MyCommon.Fetch_SystemOption(275) = 1 ))) Then
                'Forcing this input to be blank for a template, since they cannot include a defined PLU number
                Send("<input type=""text"" id=""plu"" name=""plu"" maxlength=""" & IIf(IDLength > 0, IDLength, "") & """ style=""width:200px;"" value="""" disabled=""disabled"" /><br />")
            Else
                Send("<input type=""text"" id=""plu"" name=""plu"" maxlength=""" & IIf(IDLength > 0, IDLength, "") & """ style=""width:200px;"" value=""" & IIf(IncentivePLUID > 0 And PLUString <> "", PLUString, "") & """" & IIf(FromTemplate And DisallowEdit, " disabled=""disabled""", "") & " /><br />")
            End If
            If Not (RangeBegin = 0 AndAlso RangeEnd = 0) Then
                If RangeBegin <> RangeEnd Then
                    If RangeBegin > RangeEnd Then
                        Sendb(Copient.PhraseLib.Lookup("ueoffer-con-plu.InvalidRangeDefinition", LanguageID))
                    Else
                        Sendb(Copient.PhraseLib.Detokenize("ueoffer-con-plu.RangeBounds", LanguageID, RangeBeginString, RangeEndString))
                    End If
                Else
                    Sendb(Copient.PhraseLib.Detokenize("ueoffer-con-plu.RangeBegin", LanguageID, RangeBeginString))
                End If
            Else
                Sendb(Copient.PhraseLib.Lookup("ueoffer-con-plu.NoRange", LanguageID))
            End If
            If MyCommon.Fetch_UE_SystemOption(95) Then
                Sendb(" " & Copient.PhraseLib.Lookup("ueoffer-con-plu.OutOfRangeAccepted", LanguageID))
            Else
                Sendb(" " & Copient.PhraseLib.Lookup("ueoffer-con-plu.OutOfRangeNotAccepted", LanguageID))
            End If
            Send("<br />")
            Send("<label>  " & Copient.PhraseLib.Lookup("term.Quantity", LanguageID) & " </label>")
            Send("<input type=""text"" id=""pluqty"" name=""pluqty"" maxlength=""4""  style=""width:70px;"" value=""" & IIf(IncentivePLUID > 0 And PLUQty <> 0, PLUQty, 1) & """ onkeypress=""return AllowOnlyInt(event);"" required/>")
            Send("<br />")
            Send("<br class=""half"" />")
            Send("<hr />")
            ' PLU selector
            Send("<label for=""selector"">" & Copient.PhraseLib.Lookup("ueoffer-con-plu.TopUnusedCodes", LanguageID) & "</label><br />")
            If IsTemplate AndAlso ((MyCommon.Fetch_UE_SystemOption(94) = 0) OrElse (MyCommon.Fetch_UE_SystemOption(94) = 1 AndAlso MyCommon.Fetch_SystemOption(275) = 1 )) Then
                'Forcing this list box disabled for a template, since they cannot include a defined PLU number
                Send("<select id=""selector"" name=""selector"" size=""10"" style=""width:  220px;"" disabled=""disabled"";"">")
            Else
                Send("<select id=""selector"" name=""selector"" size=""10"" style=""width:220px;"" ondblclick=""javascript:selectPLU(" & IDLength & ");""" & IIf(FromTemplate And DisallowEdit, " disabled=""disabled""", "") &">")
            End If
            i = RangeBegin
            counter = 1
            While (counter <= 100) AndAlso (i <= RangeEnd)
                MyCommon.QueryStr = "select CAST(PLU as decimal(26,0)) as PLU from CPE_IncentivePLUs as CIP with (NoLock) " & _
                                    "left join CPE_RewardOptions as RO on RO.RewardOptionID=CIP.RewardOptionID " & _
                                    "left join CPE_Incentives as I on I.IncentiveID=RO.IncentiveID " & _
                                    "where IsNull(PLU, '') <> '' and PLU='" & i.ToString.PadLeft(IDLength, "0") & "' and I.Deleted=0;"
                rst = MyCommon.LRT_Select
            If rst.Rows.Count = 0 Then
                Send("  <option value=""" & i & """>" & i.ToString.PadLeft(IDLength, "0") & "</option>")
                counter += 1
            End If
            i += 1
            End While
            Send("</select>")
            'End If
        %>
      </div>
      <div class="box" id="requirement">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.requirement", LanguageID))%>
          </span>
        </h2>
        <input type="radio" id="PerRedemption0" name="PerRedemption"<% Sendb(IIf(PerRedemption, "", " checked=""checked""")) %><% Sendb(IIf(FromTemplate And DisallowEdit, " disabled=""disabled""", "")) %> value="0" /><label for="PerRedemption0"><% Sendb(Copient.PhraseLib.Lookup("term.oncepertransaction", LanguageID))%></label>
        <br />
        <input type="radio" id="PerRedemption1" name="PerRedemption"<% Sendb(IIf(PerRedemption, " checked=""checked""", "")) %><% Sendb(IIf(FromTemplate And DisallowEdit, " disabled=""disabled""", "")) %> value="1" /><label for="PerRedemption1"><% Sendb(Copient.PhraseLib.Lookup("term.OncePerRedemption", LanguageID))%></label>
        <hr class="hidden" />
      </div>
	  <%If ExtRedemAuthorizationEnabled Then%>
	    <div class="box" id="advanced">
          <h2>
            <span>
              <% Sendb(Copient.PhraseLib.Lookup("term.advanced", LanguageID))%>
            </span>
          </h2>
		  	<input type="checkbox" id="ExtRedemAuthorization" name="ExtRedemAuthorization"<% If (ExtRedemAuthorization) Then Sendb(" checked=""checked""") %><% Sendb(IIf(FromTemplate And DisallowEdit, " disabled=""disabled""", "")) %> />
			<label for="ExtRedemAuthorization"><% Sendb(Copient.PhraseLib.Lookup("term.extredemauthorization", LanguageID))%></label>
			<hr class="hidden" />
        </div>
	  <%End If%>
    </div>
    <div id="gutter">
    </div>
    
    <div id="column2">
    </div>
  </div>
</form>

<script runat="server">
  Public Function AllDigits(ByVal txt As String) As Boolean
    Dim ch As String
    Dim i As Integer
    
    AllDigits = True
    For i = 1 To Len(txt)
      ' See if the next character is a non-digit.
      ch = Mid$(txt, i, 1)
      If ch < "0" Or ch > "9" Then
        ' This is not a digit.
        AllDigits = False
        Exit For
      End If
    Next i
  End Function
</script>

<script type="text/javascript">
<% If (CloseAfterSave) Then %>
    opener.location = "/logix/UE/UEoffer-con.aspx?OfferID=<%Sendb(OfferID)%>";
    window.close();
<% End If %>
    function AllowOnlyInt(evt) {
    var nkeycode = (window.event) ? window.event.keyCode : evt.which;
    var exceptionKeycodes = new Array(8, 9, 13, 16);
    if ((nkeycode >= 48 && nkeycode <= 57)) {
        return true;
    } else {
        for (var i = 0; i < exceptionKeycodes.length; i++) {
            if (nkeycode == exceptionKeycodes[i]) {
                return true;
            }
        }
        return false;
    }
}
</script>

<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd("mainform", "plu")
  MyCommon = Nothing
  Logix = Nothing
%>
