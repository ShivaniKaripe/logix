<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME:web-offer-con.aspx 
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
  Dim dt As DataTable
  Dim OfferID As Long
  Dim ConditionID As Long
  Dim Name As String = ""
  Dim isTemplate As Boolean = False
  Dim FromTemplate As Boolean = False
  Dim Disallow_Conditions As Boolean = False
  Dim IsTemplateVal As String = "Not"
  Dim ActiveSubTab As Integer = 91
  Dim roid As Integer
  Dim i As Integer
  Dim isCustomer As Boolean = False
  Dim isPoint As Boolean = False
  Dim DeleteBtnDisabled As String = ""
  Dim infoMessage As String = ""
  Dim modMessage As String = ""
  Dim Handheld As Boolean = False
  Dim CondTypes As String() = Nothing
  Dim Conditions As String() = Nothing
  Dim LockedStatus As String() = Nothing
  Dim LoopCtr As Integer = 0
  Dim sQuery As String = ""
  Dim BannersEnabled As Boolean = True
  Dim StatusCode As Copient.LogixInc.STATUS_FLAGS = Copient.LogixInc.STATUS_FLAGS.STATUS_UNKNOWN
  Dim StatusText As String = ""
  Dim StatusFlag As Integer
  Dim TierLevels As Integer = 1
  Dim t As Integer = 1
  Dim IncentiveID As Integer = 0
  Dim IncentiveTenderID As Integer = 0
  Dim EngineID As Integer = 3
  Dim EngineSubTypeID As Integer = 0
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "web-offer-con.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)

  DeleteBtnDisabled = IIf(Logix.UserRoles.EditOffer, "", " disabled=""disabled""")
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  
  OfferID = Request.QueryString("OfferID")
  ConditionID = Request.QueryString("ConditionID")
  
  If (OfferID = 0) Then
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "web-offer-gen.aspx")
  End If
  
  MyCommon.QueryStr = "select RewardOptionID, TierLevels from CPE_RewardOptions with (NoLock) " & _
                      "where IncentiveID=" & OfferID & " and TouchResponse=0 and Deleted=0;"
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    roid = rst.Rows(0).Item("RewardOptionID")
    TierLevels = rst.Rows(0).Item("TierLevels")
  End If
  
  If Request.QueryString("IncentiveTenderID") <> "" Then
    IncentiveTenderID = MyCommon.Extract_Val(Request.QueryString("IncentiveTenderID"))
  End If
  
  Send_HeadBegin("term.offer", "term.conditions", OfferID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
%>
<script type="text/javascript">
function LoadDocument(url) { 
  location = url; 
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
  
  ' handle adding stuff on
  If (Request.QueryString("Save") = "" And Request.QueryString("newconglobal") <> "") Then
    If (Request.QueryString("newconglobal") = 1) Then
      Send("<script type=""text/javascript"">openPopup('CPEoffer-con-customer.aspx?OfferID=" & OfferID & "')</script>")
    ElseIf (Request.QueryString("newconglobal") = 3) Then
      Send("<script type=""text/javascript"">openPopup('CPEoffer-con-point.aspx?OfferID=" & OfferID & "')</script>")
    End If
  ElseIf (Request.QueryString("mode") = "Delete") Then
    If (Request.QueryString("Option") = "Customer") Then
      ' ok someone clicked the X on the customer group stuff lets ditch all the associated groups on this offer
      MyCommon.QueryStr = "update CPE_IncentiveCustomerGroups with (RowLock) set deleted=1, LastUpdate=getdate(),TCRMAStatusFlag=3 where RewardOptionID=" & roid & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1, AllowOptOut=0 where IncentiveID=" & OfferID & ";"
      MyCommon.LRT_Execute()
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-customer-delete", LanguageID))
    ElseIf (Request.QueryString("Option") = "Point") Then
      ' ok someone clicked the X on the customer group stuff lets ditch all the associated groups on this offer
      MyCommon.QueryStr = "delete from CPE_IncentivePointsGroups with (RowLock) where RewardOptionID=" & roid & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "delete from CPE_IncentivePointsGroupTiers with (RowLock) where RewardOptionID=" & roid & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
      MyCommon.LRT_Execute()
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-point-delete", LanguageID))
    End If
  End If
  
  'update the template permission for Conditions
  If (Request.QueryString("Save") <> "") Then
    If (Request.QueryString("IsTemplate") = "IsTemplate") Then
      ' time to update the status bits for the templates   
      Dim form_Disallow_Conditions As Integer = 0
      If (Request.QueryString("Disallow_Conditions") = "on") Then
        form_Disallow_Conditions = 1
      End If
      MyCommon.QueryStr = "update TemplatePermissions with (RowLock) set Disallow_Conditions=" & form_Disallow_Conditions & _
                          " where OfferID=" & OfferID
      MyCommon.LRT_Execute()
      
      'Update the lock status for each condition
      CondTypes = Request.QueryString.GetValues("conType")
      Conditions = Request.QueryString.GetValues("con")
      LockedStatus = Request.QueryString.GetValues("locked")
      If (Not CondTypes Is Nothing AndAlso Not Conditions Is Nothing AndAlso Not LockedStatus Is Nothing AndAlso Conditions.Length = LockedStatus.Length) Then
        For LoopCtr = 0 To Conditions.GetUpperBound(0)
          Select Case CondTypes(LoopCtr)
            Case "Customer"
              sQuery = "update CPE_IncentiveCustomerGroups with (RowLock) set DisallowEdit=" & LockedStatus(LoopCtr) & ", " & _
                       "RequiredFromTemplate=" & IIf(LockedStatus(LoopCtr) = "1", "0", "RequiredFromTemplate") & " " & _
                       "where IncentiveCustomerID = " & Conditions(LoopCtr) & ";"
            Case "Points"
              sQuery = "update CPE_IncentivePointsGroups with (RowLock) set DisallowEdit=" & LockedStatus(LoopCtr) & ", " & _
                       "RequiredFromTemplate=" & IIf(LockedStatus(LoopCtr) = "1", "0", "RequiredFromTemplate") & " " & _
                       "where IncentivePointsID = " & Conditions(LoopCtr) & ";"
          End Select
          MyCommon.QueryStr = sQuery
          MyCommon.LRT_Execute()
        Next
      End If
    End If
  End If
  
  ' dig the offer info out of the database
  ' no one clicked anything
  MyCommon.QueryStr = "select IncentiveID, ClientOfferID, IncentiveName as Name, CPE.Description, PromoClassID, CRMEngineID, Priority," & _
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
    StatusFlag = MyCommon.NZ(row.Item("StatusFlag"), 0)
    EngineID = MyCommon.NZ(row.Item("EngineID"), 3)
    EngineSubTypeID = MyCommon.NZ(row.Item("EngineSubTypeID"), 0)
  Next
  
  If (isTemplate Or FromTemplate) Then
    ' lets dig the permissions if its a template
    MyCommon.QueryStr = "select Disallow_Conditions from TemplatePermissions with (NoLock) where OfferID=" & OfferID & ";"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      For Each row In rst.Rows
        ' ok there are some rows for the template
        Disallow_Conditions = MyCommon.NZ(row.Item("Disallow_Conditions"), True)
      Next
    End If
  End If
  
  If (isTemplate) Then
    ActiveSubTab = 27
    IsTemplateVal = "IsTemplate"
  Else
    ActiveSubTab = 27
    IsTemplateVal = "Not"
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
  Send_Subtabs(Logix, ActiveSubTab, 5, , OfferID)
  
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
<form action="web-offer-con.aspx" id="mainform" name="mainform">
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
      StatusText = Logix.GetOfferStatus(OfferID, LanguageID, StatusCode)
      If Not isTemplate Then
        If (StatusFlag <> 2) Then
          If (StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_ACTIVE) AndAlso (StatusFlag > 0) Then
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
        MyCommon.QueryStr = "select IncentiveID from CPE_Incentives with (NoLock) where CreatedDate=LastUpdate and IncentiveID=" & OfferID & ";"
        rst3 = MyCommon.LRT_Select
        If (rst3.Rows.Count = 0) Then
          Send_Status(OfferID, 2)
        End If
      End If
    %>
    <div id="column">
      <div class="box" id="conditions">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.conditions", LanguageID))%>
          </span>
        </h2>
        <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.conditions", LanguageID)) %>">
          <thead>
            <tr>
              <th align="left" scope="col" class="th-del">
                <% Sendb(Copient.PhraseLib.Lookup("term.delete", LanguageID))%>
              </th>
              <th align="left" scope="col" class="th-andor">
                <% Sendb(Copient.PhraseLib.Lookup("term.andor", LanguageID))%>
              </th>
              <th align="left" scope="col" class="th-type">
                <% Sendb(Copient.PhraseLib.Lookup("term.type", LanguageID))%>
              </th>
              <th align="left" scope="col" class="th-details">
                <% Sendb(Copient.PhraseLib.Lookup("term.details", LanguageID))%>
              </th>
              <th align="left" scope="col" class="th-information" colspan="<% Sendb(TierLevels) %>">
                <% Sendb(Copient.PhraseLib.Lookup("term.information", LanguageID))%>
              </th>
              <% If (isTemplate OrElse FromTemplate) Then%>
              <th align="left" scope="col" class="th-locked">
                <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
              </th>
              <% End If%>
            </tr>
          </thead>
          <tbody>
            <!-- CUSTOMER CONDITIONS -->
            <%
              t = 1
              ' Find the currently selected groups on page load
              MyCommon.QueryStr = "select ICG.IncentiveCustomerID, CG.CustomerGroupID,Name,PhraseID,ExcludedUsers,DisallowEdit," & _
                                  " RequiredFromTemplate from CPE_IncentiveCustomerGroups as ICG with (NoLock) " & _
                                  " left join CustomerGroups as CG with (NoLock) " & _
                                  " on CG.CustomerGroupID=ICG.CustomerGroupID where RewardOptionID=" & roid & _
                                  " and ICG.Deleted=0;"
              rst = MyCommon.LRT_Select
              If (rst.Rows.Count > 0) Then
                Send("<tr class=""shadeddark"">")
                Send("  <td colspan=""4"">")
                Send("    <h3>")
                Send("      " & Copient.PhraseLib.Lookup("term.customerconditions", LanguageID))
                Send("    </h3>")
                Send("  </td>")
                Send("  <td colspan=""" & TierLevels & """>")
                Send("  </td>")
                If (isTemplate Or FromTemplate) Then
                  Send("  <td></td>")
                End If
                Send("</tr>")
              End If
              i = 1
              For Each row In rst.Rows
                ' We got in the loop, so there's a customer condition; set it as such.
                isCustomer = True
                Send("<tr class=""shaded"">")
                Send("  <td>")
                If (i = 1) Then
                  'If (Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Conditions)) Then
                  '    Sendb("<input type=""button"" class=""ex"" id=""customerDelete"" name=""customerDelete"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('web-offer-con.aspx?mode=Delete&Option=Customer&OfferID=" & OfferID & "')}"" value=""X"" />")
                  'Else
                  Sendb("<input type=""button"" class=""ex"" id=""customerDelete"" name=""customerDelete"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('web-offer-con.aspx?mode=Delete&Option=Customer&OfferID=" & OfferID & "')}"" value=""X"" />")
                  'End If
                End If
                Send("  </td>")
                Send("  <td>")
                If (i > 1 And MyCommon.NZ(row.Item("ExcludedUsers"), False) = False) Then
                  Send("    " & Copient.PhraseLib.Lookup("term.or", LanguageID))
                End If
                Send("  </td>")
                Send("  <td>")
                Send("    <a href=""javascript:openPopup('CPEoffer-con-customer.aspx?OfferID=" & OfferID & "')"">" & Copient.PhraseLib.Lookup("term.customer", LanguageID) & "</a>")
                Send("  </td>")
                Send("  <td>")
                If (MyCommon.NZ(row.Item("ExcludedUsers"), False) = True) Then Sendb(StrConv(Copient.PhraseLib.Lookup("term.excluding", LanguageID), VbStrConv.Lowercase) & " ")
                If (MyCommon.NZ(row.Item("CustomerGroupID"), -1) > 2) Then
                  Sendb("<a href=""cgroup-edit.aspx?CustomerGroupID=" & row.Item("CustomerGroupID") & """>")
                  If IsDBNull(row.Item("PhraseID")) Then
                    Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
                  Else
                    If (MyCommon.NZ(row.Item("PhraseID"), 0) = 0) Then
                      Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
                    Else
                      Sendb(MyCommon.SplitNonSpacedString(Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID), 25) & "</a>")
                    End If
                  End If
                ElseIf (IsDBNull(row.Item("CustomerGroupID")) AndAlso MyCommon.NZ(row.Item("RequiredFromTemplate"), False)) Then
                  Sendb("<span class=""red"">* ")
                  Sendb(Copient.PhraseLib.Lookup("term.required", LanguageID))
                  Sendb(" " & Copient.PhraseLib.Lookup("term.by", LanguageID))
                  Sendb(" " & Copient.PhraseLib.Lookup("term.template", LanguageID))
                  Sendb("</span>")
                Else
                  If IsDBNull(row.Item("PhraseID")) Then
                    Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25))
                  Else
                    If (MyCommon.NZ(row.Item("PhraseID"), 0) = 0) Then
                      Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25))
                    Else
                      Sendb(MyCommon.SplitNonSpacedString(Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID), 25))
                    End If
                  End If
                End If
                Send("  </td>")
                Send("  <td colspan=""" & TierLevels & """>")
                Send("  </td>")
                If (isTemplate) Then
                  Send("  <td class=""templine"">")
                  %>
                  <input type="checkbox" id="chkLocked1" name="chkLocked" value="<%Sendb(MyCommon.NZ(row.Item("IncentiveCustomerID"), 0))%>"<%Sendb(IIf(MyCommon.NZ(row.Item("DisallowEdit"), False)=True, " checked=""checked""", "")) %> onclick="javascript:updateLocked('lockCust<%Sendb(MyCommon.NZ(row.Item("IncentiveCustomerID"), 0))%>', this.checked);"<%Sendb(IIf(IsDBNull(row.Item("CustomerGroupID")) AndAlso MyCommon.NZ(row.Item("RequiredFromTemplate"), False), " disabled=""disabled""", "")) %> />
                  <input type="hidden" id="conType1" name="conType" value="Customer" />
                  <input type="hidden" id="conCust<%Sendb(MyCommon.NZ(row.Item("IncentiveCustomerID"), 0))%>" name="con" value="<%Sendb(MyCommon.NZ(row.Item("IncentiveCustomerID"), 0))%>" />
                  <input type="hidden" id="lockCust<%Sendb(MyCommon.NZ(row.Item("IncentiveCustomerID"), 0))%>" name="locked" value="<%Sendb(IIf(MyCommon.NZ(row.Item("DisallowEdit"), False)=True, "1", "0"))%>" />
                  <%
                  Send("  </td>")
                ElseIf (FromTemplate) Then
                  Send("  <td class=""templine"">")
                  Send("    " & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "Yes", "No"))
                  Send("  </td>")
                End If
                Send("</tr>")
                i += 1
              Next
            %>
            
            <!-- POINTS CONDITIONS -->
            <%
              t = 1
              MyCommon.QueryStr = "Select IPG.IncentivePointsID, IPG.ProgramID, ProgramName, QtyForIncentive, DisallowEdit, RequiredFromTemplate " & _
                                  "from CPE_IncentivePointsGroups as IPG with (NoLock) " & _
                                  "left join PointsPrograms as PP with (NoLock) " & _
                                  "on PP.ProgramID=IPG.ProgramID where RewardOptionID=" & roid & ";"
              rst = MyCommon.LRT_Select
              If (rst.Rows.Count > 0) Then
                Send("<tr class=""shadeddark"">")
                Send("  <td colspan=""" & 4 & """>")
                Send("    <h3>")
                Send("      " & Copient.PhraseLib.Lookup("term.pointsconditions", LanguageID))
                Send("    </h3>")
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
                If (isTemplate Or FromTemplate) Then
                  Send("  <td></td>")
                End If
                Send("</tr>")
              End If
              For Each row In rst.Rows
                isPoint = True
                Send("<tr class=""shaded"">")
                Send("  <td>")
                If (Logix.UserRoles.EditOffer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False))) Then
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('web-offer-con.aspx?mode=Delete&Option=Point&OfferID=" & OfferID & "')}"" value=""X"" />")
                Else
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('web-offer-con.aspx?mode=Delete&Option=Point&OfferID=" & OfferID & "')}"" value=""X"" />")
                End If
                Send("  </td>")
                Send("  <td>")
                Send("  </td>")
                Send("  <td>")
                Send("    <a href=""javascript:openPopup('CPEoffer-con-point.aspx?OfferID=" & OfferID & "&IncentivePointsID=" & MyCommon.NZ(row.Item("IncentivePointsID"), 0) & "')"">" & Copient.PhraseLib.Lookup("term.points", LanguageID) & "</a>")
                Send("  </td>")
                Send("  <td>")
                If (MyCommon.NZ(row.Item("ProgramID"), -1) > -1) Then
                  Sendb("    <a href=""point-edit.aspx?ProgramGroupID=" & row.Item("ProgramID") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("ProgramName"), ""), 25) & "</a>")
                ElseIf (IsDBNull(row.Item("ProgramID")) AndAlso MyCommon.NZ(row.Item("RequiredFromTemplate"), False)) Then
                  Sendb("<span class=""red"">* ")
                  Sendb(Copient.PhraseLib.Lookup("term.required", LanguageID))
                  Sendb(" " & Copient.PhraseLib.Lookup("term.by", LanguageID))
                  Sendb(" " & Copient.PhraseLib.Lookup("term.template", LanguageID))
                  Send("</span>")
                End If
                Send("  </td>")
                ' Find the per-tier values:
                t = 1
                MyCommon.QueryStr = "select IncentivePointsID, TierLevel, Quantity from CPE_IncentivePointsGroupTiers as IPGT with (NoLock) " & _
                                    "where RewardOptionID=" & roid & ";"
                rst2 = MyCommon.LRT_Select
                If rst2.Rows.Count = 0 Then
                  Send("  <td colspan=""" & TierLevels & """>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                Else
                  While t <= TierLevels
                    If t > rst2.Rows.Count Then
                      Send("  <td>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                    Else
                      Send("  <td>")
                      Send("    " & MyCommon.NZ(rst2.Rows(t - 1).Item("Quantity"), 0) & " " & StrConv(Copient.PhraseLib.Lookup("term.points", LanguageID), VbStrConv.Lowercase) & " " & StrConv(Copient.PhraseLib.Lookup("term.required", LanguageID), VbStrConv.Lowercase))
                      Send("  </td>")
                    End If
                    t += 1
                  End While
                End If
                If (isTemplate) Then
                  Send("  <td class=""templine"">")
                  %>
                  <input type="checkbox" id="chkLocked3" name="chkLocked" value="<%Sendb(MyCommon.NZ(row.Item("IncentivePointsID"), 0))%>"<%Sendb(IIf(MyCommon.NZ(row.Item("DisallowEdit"), False)=True, " checked=""checked""", "")) %> onclick="javascript:updateLocked('lockPt<%Sendb(MyCommon.NZ(row.Item("IncentivePointsID"), 0))%>', this.checked);"<%Sendb(IIf(IsDBNull(row.Item("ProgramID")) AndAlso MyCommon.NZ(row.Item("RequiredFromTemplate"), False), " disabled=""disabled""", "")) %> />
                  <input type="hidden" id="conType3" name="conType" value="Points" />
                  <input type="hidden" id="conPt<%Sendb(MyCommon.NZ(row.Item("IncentivePointsID"), 0))%>" name="con" value="<%Sendb(MyCommon.NZ(row.Item("IncentivePointsID"), 0))%>" />
                  <input type="hidden" id="lockPt<%Sendb(MyCommon.NZ(row.Item("IncentivePointsID"), 0))%>" name="locked" value="<%Sendb(IIf(MyCommon.NZ(row.Item("DisallowEdit"), False)=True, "1", "0"))%>" />
                  <%
                  Send("  </td>")
                ElseIf (FromTemplate) Then
                  Send("  <td class=""templine"">")
                  Send("    " & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "Yes", "No"))
                  Send("  </td>")
                End If
                Send("</tr>")
              Next
            %>
          </tbody>
        </table>
        <hr class="hidden" />
      </div>
      <%
        Dim HideAdd As Boolean = (isCustomer AndAlso isPoint AndAlso Not isTemplate)
      %>
      <div class="box" id="newcondition"<%Sendb(IIf(HideAdd, " style=""visibility:hidden;"" ", ""))%>>
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("offer-con.addcondition", LanguageID))%>
          </span>
        </h2>
        <%
          If (isTemplate) Then
            Send("<span class=""temp"">")
            Send("  <input type=""checkbox"" class=""tempcheck"" id=""Disallow_Conditions"" name=""Disallow_Conditions""" & IIf(Disallow_Conditions, " checked=""checked""", "") & " />")
            Send("  <label for=""Disallow_Conditions"">" & Copient.PhraseLib.Lookup("term.locked", LanguageID) & "</label>")
            Send("</span>")
          End If
          MyCommon.QueryStr = "SELECT PECT.EngineID, PECT.EngineSubTypeID, PECT.ComponentTypeID, CT.ConditionTypeID, CT.Description, CT.PhraseID, PECT.Singular, " & _
                              "  CASE ConditionTypeID " & _
                              "    WHEN 1 THEN (SELECT COUNT(*) FROM CPE_IncentiveCustomerGroups WITH (NOLOCK) where RewardOptionID=" & roid & " and Deleted=0) " & _
                              "    WHEN 2 THEN (SELECT COUNT(*) FROM CPE_IncentiveProductGroups WITH (NOLOCK) where RewardOptionID=" & roid & " and Deleted=0 and Disqualifier=0) " & _
                              "    WHEN 3 THEN (SELECT COUNT(*) FROM CPE_IncentivePointsGroups WITH (NOLOCK) where RewardOptionID=" & roid & " and Deleted=0) " & _
                              "    WHEN 4 THEN (SELECT COUNT(*) FROM CPE_IncentiveStoredValuePrograms WITH (NOLOCK) where RewardOptionID=" & roid & " and Deleted=0) " & _
                              "    WHEN 5 THEN (SELECT COUNT(*) FROM CPE_IncentiveTenderTypes WITH (NOLOCK) where RewardOptionID=" & roid & " and Deleted=0) " & _
                              "    WHEN 6 THEN (SELECT COUNT(*) FROM CPE_IncentiveDOW WITH (NOLOCK) where IncentiveID=" & OfferID & " and Deleted=0) " & _
                              "    WHEN 7 THEN (SELECT COUNT(*) FROM CPE_IncentiveTOD WITH (NOLOCK) where IncentiveID=" & OfferID & ") " & _
                              "    WHEN 8 THEN (SELECT COUNT(*) FROM CPE_IncentiveInstantWin WITH (NOLOCK) where RewardOptionID=" & roid & " and Deleted=0) " & _
                              "    WHEN 9 THEN (SELECT COUNT(*) FROM CPE_IncentivePLUs WITH (NOLOCK) where RewardOptionID=" & roid & ") " & _
                              "    WHEN 10 THEN (SELECT COUNT(*) FROM CPE_IncentiveProductGroups WITH (NOLOCK) where RewardOptionID=" & roid & " and Deleted=0 and Disqualifier=1) " & _
                              "    WHEN 11 THEN (SELECT COUNT(*) FROM CPE_IncentiveEIW WITH (NOLOCK) where RewardOptionID=" & roid & ") " & _
                              "    WHEN 12 THEN (SELECT COUNT(*) FROM CPE_IncentiveAttributes WITH (NOLOCK) where RewardOptionID=" & roid & " and Deleted=0) " & _
                              "    WHEN 13 THEN (SELECT COUNT(*) FROM CPE_IncentiveCardTypes WITH (NOLOCK) where RewardOptionID=" & roid & " and Deleted=0) " & _
                              "    ELSE 0 " & _
                              "  END as ItemCount " & _
                              "FROM PromoEngineComponentTypes AS PECT " & _
                              "INNER JOIN CPE_ConditionTypes AS CT ON CT.ConditionTypeID=PECT.LinkID " & _
                              "WHERE EngineID=3 AND EngineSubTypeID=" & EngineSubTypeID & " AND PECT.ComponentTypeID=1 AND Enabled=1"
          'Impose a few special limits on the query based on various in-page factors:
          If (Not isCustomer) Then
            'Offer has no customer condition, so limit it just to that
            MyCommon.QueryStr &= " AND CT.ConditionTypeID=1"
          End If
          MyCommon.QueryStr &= " ORDER BY DisplayOrder;"
          rst = MyCommon.LRT_Select
          If rst.Rows.Count > 0 Then
            Send("<label for=""newconglobal"">" & Copient.PhraseLib.Lookup("offer-con.addglobal", LanguageID) & ":</label><br />")
            Send("<select id=""newconglobal"" name=""newconglobal""" & IIf(isTemplate OrElse (Not Disallow_Conditions), "", " disabled=""disabled""") & ">")
            For Each row In rst.Rows
              If (row.Item("Singular") = False) OrElse (row.Item("Singular") = True AndAlso row.Item("ItemCount") = 0) Then
                Send("<option value=""" & row.Item("ConditionTypeID") & """>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "</option>")
              End If
            Next
            Send("</select>")
            Sendb("<input class=""regular"" id=""addGlobal"" name=""addGlobal"" type=""submit"" value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """")
            If isTemplate OrElse (Not (isCustomer And isPoint) And Not Disallow_Conditions) Then
            Else
              Sendb(" disabled=""disabled""")
            End If
            Sendb(" />")
          End If
        %>
      </div>
    </div>
    <br clear="all" />
  </div>
</form>
<!-- #Include virtual="/include/graphic-reward.inc" -->
<%If (isPoint) Then%>
<script type="text/javascript">
  var elemCustDelBtn = document.getElementById("customerDelete");
  
  if (elemCustDelBtn != null) {
    elemCustDelBtn.disabled = true;
  }
</script>
<%End If%>
<%If (isCustomer OrElse isPoint) Then%>
<%Else%>
<script type="text/javascript" language="javascript">
  var elemConditions = document.getElementById("conditions");
  
  if (elemConditions != null) {
    elemConditions.style.display = "none";
  }
</script>
<%End If%>
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
