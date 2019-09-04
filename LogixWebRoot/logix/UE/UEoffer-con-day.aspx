<%@ Page Language="vb" Debug="true" CodeFile="ue-cb.vb" Inherits="UECB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
    ' *****************************************************************************
    ' * FILENAME: UEoffer-con-day.aspx 
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
    Dim MyCPEOffer As New Copient.EIW
    Dim Logix As New Copient.LogixInc
    Dim OfferID As Long
    Dim Name As String = ""
    Dim ConditionID As String
    Dim IsTemplate As Boolean = False
    Dim Disallow_Edit As Boolean = True
    Dim DisabledAttribute As String = ""
    Dim FromTemplate As Boolean = False
    Dim i As Integer
    Dim row As DataRow
    Dim rst As DataTable
    Dim rst2 As DataTable
    Dim row2 As DataRow
    Dim CloseAfterSave As Boolean = False
    Dim historyString As String = ""
    Dim roid As Integer
    Dim slot As Integer = 0
    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False
    Dim EngineID As Integer = 2
    Dim BannersEnabled As Boolean = True
    Dim NoDaysChecked As Boolean
    Dim OfferHasEIW As Boolean = False

    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If


    Response.Expires = 0
    MyCommon.AppName = "UEoffer-con-day.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)

    BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")

    OfferID = Request.QueryString("OfferID")

    'This would redirect to summary page if offer is awaiting deployment or if offer is awaiting recommendations.
    CheckIfValidOffer(MyCommon, OfferID)

    ConditionID = Request.QueryString("ConditionID")

    Dim isTranslatedOffer As Boolean =MyCommon.IsTranslatedUEOffer(OfferID,  MyCommon)
    Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249)="1",True,False)

    Dim bEnableAdditionalLockoutRestrictionsOnOffers As Boolean = IIf(MyCommon.Fetch_SystemOption(260) = "1", True, False)
    Dim bOfferEditable As Boolean = MyCommon.IsOfferEditablePastLockOutPeriod(Logix.UserRoles.EditOfferPastLockoutPeriod, MyCommon, OfferID)

    If (Request.QueryString("EngineID") <> "") Then
        EngineID = MyCommon.Extract_Val(Request.QueryString("EngineID"))
    Else
        MyCommon.QueryStr = "select EngineID from OfferIDs where OfferID=" & OfferID & ";"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
            EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), 0)
        End If
    End If

    MyCommon.QueryStr = "select RewardOptionID from CPE_RewardOptions with (NoLock) where IncentiveID=" & OfferID & " and touchresponse=0 and deleted=0;"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
        roid = rst.Rows(0).Item("RewardOptionID")
    End If

    'Determine if the offer has an enterprise instant win condition
    MyCommon.QueryStr = "select IncentiveEIWID from CPE_IncentiveEIW with (NoLock) where RewardOptionID=" & roid & ";"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
        OfferHasEIW = True
    End If

    ' see if someone is saving
    If (Request.QueryString("save") <> "") Then

        'store the existing locking value for use in newly-created records
        MyCommon.QueryStr = "select DisallowEdit from CPE_IncentiveDOW with (NoLock) where deleted=0 and IncentiveID=" & OfferID
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
            Disallow_Edit = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit"), False)
        Else
            Disallow_Edit = False
        End If

        ' first delete all the ones out of the db for this offer
        MyCommon.QueryStr = "update CPE_IncentiveDOW with (RowLock) set deleted=1 where IncentiveID=" & OfferID
        MyCommon.LRT_Execute()
        MyCommon.QueryStr = "select DOWID, PhraseID from CPE_DaysOfWeek DW with (NoLock)"
        rst = MyCommon.LRT_Select
        historyString = Copient.PhraseLib.Lookup("history.con-day", LanguageID)
        NoDaysChecked = True
        For i = 0 To 6
            'Send("querystr-> val-" & i & "=" & Request.QueryString("val-" & i) & "<br />")
            If (Request.QueryString("val-" & i) = "on") Then
                NoDaysChecked = False
                MyCommon.QueryStr = "insert into CPE_IncentiveDOW with (RowLock) (IncentiveID,DOWID,Deleted,DisallowEdit) values(" & OfferID & "," & i & "," & "0," & IIf(Disallow_Edit, "1", "0") & ")"
                'Send(MyCommon.QueryStr & "<br />")
                MyCommon.LRT_Execute()
                ' add currently selected day to history string
                For Each row In rst.Rows
                    If (MyCommon.NZ(row.Item("DOWID"), 0) = i) Then
                        historyString = historyString & " " & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & " "
                    End If
                Next
            End If
        Next

        ' reset the EveryDOW column to reflect the change, if all 7 days chosen then set to 1, otherwise set to 0
        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set EveryDOW = " & _
                            "   case (select count(*) DayCount from CPE_IncentiveDOW where IncentiveID=" & OfferID & " and Deleted=0) " & _
                            "       when 0 then 1 " & _
                            "       when 7 then 1 " & _
                            "       else 0 " & _
                            "   end " & _
                            "where incentiveid = " & OfferID & " and deleted=0;"
        MyCommon.LRT_Execute()

        'Randomize any EIW triggers associated with this offer:
        If OfferHasEIW Then
            MyCPEOffer.RandomizeTriggersByOffer(OfferID)
            MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-instantwin-randomize", LanguageID))
        End If

        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID
        MyCommon.LRT_Execute()
        ResetOfferApprovalStatus(OfferID)
        MyCommon.Activity_Log(3, OfferID, AdminUserID, historyString)
        CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1") AndAlso infoMessage = ""
    End If

    ' dig the offer info out of the database
    ' no one clicked anything
    MyCommon.QueryStr = "Select IncentiveID,IsTemplate,ClientOfferID,IncentiveName as Name,CPE.Description,PromoClassID,CRMEngineID,Priority," & _
                        "StartDate,EndDate,EveryDOW,EligibilityStartDate,EligibilityEndDate,TestingStartDate,TestingEndDate,P1DistQtyLimit,P1DistTimeType,P1DistPeriod," & _
                        "P2DistQtyLimit,P2DistTimeType,P2DistPeriod,P3DistQtyLimit,P3DistTimeType,P3DistPeriod,EnableImpressRpt,EnableRedeemRpt,CreatedDate,CPE.LastUpdate,CPE.Deleted, CPEOARptDate, CPEOADeploySuccessDate, CPEOADeployRpt," & _
                        "CRMRestricted,StatusFlag,OC.Description as CategoryName,IsTemplate,FromTemplate from CPE_Incentives as CPE with (NoLock) " & _
                        "left join OfferCategories as OC with (NoLock) on CPE.PromoClassID=OfferCategoryID where IncentiveID=" & Request.QueryString("OfferID")
    rst = MyCommon.LRT_Select
    For Each row In rst.Rows
        Name = MyCommon.NZ(row.Item("Name"), "")
        IsTemplate = MyCommon.NZ(row.Item("IsTemplate"), False)
        FromTemplate = MyCommon.NZ(row.Item("FromTemplate"), False)
    Next

    'update the templates permission if necessary
    If (Request.QueryString("save") <> "" AndAlso Request.QueryString("IsTemplate") = "IsTemplate") Then
        ' update the status bits for the templates
        Dim form_Disallow_Edit As Integer = 0
        If (Request.QueryString("Disallow_Edit") = "on") Then
            form_Disallow_Edit = 1
        End If
        MyCommon.QueryStr = "update CPE_IncentiveDOW with (RowLock) set DisallowEdit=" & form_Disallow_Edit & _
                            " where IncentiveID=" & OfferID & " and deleted = 0;"
        MyCommon.LRT_Execute()
    End If

    If (IsTemplate Or FromTemplate) Then
        ' lets dig the permissions if its a template
        MyCommon.QueryStr = "select DisallowEdit from CPE_IncentiveDOW with (NoLock) where IncentiveID=" & OfferID & " and deleted = 0;"
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

    Send_HeadBegin("term.offer", "term.daycondition", OfferID)
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
    <input type="hidden" id="IsTemplate" name="IsTemplate" value="<% 
        If (IsTemplate) Then
        Sendb("IsTemplate")
        Else
        Sendb("Not")
        End If
        %>" />
    <%If (IsTemplate) Then
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.daycondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
      Else
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.daycondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
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
      If((Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
        If Not IsTemplate Then
                  If (Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And Disallow_Edit) And Not IsOfferWaitingForApproval(OfferID)) Then Send_Save()
        Else
                If (Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer) Then Send_Save()
        End If
      End If
      %>
    </div>
  </div>
  <div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <div id="column1">
      <div class="box" id="day" style="min-height:300px; height:auto !important; height:300px;">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.day", LanguageID))%>
          </span>
        </h2>
        <% Sendb(Copient.PhraseLib.Lookup("condition.day", LanguageID))%>
        <br />
        <br class="half" />
        <%  
          MyCommon.QueryStr = "select DOWID, DayName, PhraseID from CPE_DaysOfWeek DW with (NoLock)"
          rst = MyCommon.LRT_Select
          MyCommon.QueryStr = "select DOWID from CPE_IncentiveDOW with (NoLock) where deleted=0 and IncentiveID=" & OfferID
          rst2 = MyCommon.LRT_Select
          For Each row In rst.Rows
            Sendb("<input type=""checkbox"" id=""" & StrConv(MyCommon.NZ(row.Item("DayName"), ""), VbStrConv.Lowercase) & """ name=""val-" & MyCommon.NZ(row.Item("DOWID"), 0) & """")
            If (rst2.Rows.Count > 0) Then
              For Each row2 In rst2.Rows
                If (MyCommon.NZ(row2.Item("DOWID"), 0) = MyCommon.NZ(row.Item("DOWID"), 0)) Then
                  Sendb(" checked=""checked""")
                End If
              Next
            Else
              ' if there are no specific days checked then by default check them all
              Sendb(" checked=""checked""")
            End If
            Sendb(DisabledAttribute)
            Sendb(" /><label ")
            Send("for=""" & StrConv(MyCommon.NZ(row.Item("DayName"), ""), VbStrConv.Lowercase) & """>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "</label><br />")
          Next
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
  Send_BodyEnd()
  MyCommon = Nothing
  Logix = Nothing
%>
