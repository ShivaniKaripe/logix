<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%@ Register TagPrefix="uc" TagName="ucOptinCondition" Src="~/logix/UserControls/OfferEligibilityConditions.ascx" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="CMS" %>
<%
  ' *****************************************************************************
  ' * FILENAME: offer-con.aspx 
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
  Dim rt As DataTable
  Dim OfferID As Long
  Dim Name As String = ""
  Dim isTemplate As Boolean = False
  Dim FromTemplate As Boolean = False
  Dim Disallow_Conditions As Boolean = False
  Dim Disallow_OptIn As Boolean = False
  Dim NumTiers As String
  Dim ConditionID As Long
  Dim prog As Integer
  Dim sConId As String
  Dim sDescription As String
  Dim bEnabled As Boolean
  Dim infoMessage As String = ""
  Dim modMessage As String = ""
  Dim Handheld As Boolean = False
  Dim ColumnCount As Integer = 7
  Dim Conditions As String() = Nothing
  Dim LockedStatus As String() = Nothing
  Dim LoopCtr As Integer
  Dim BannersEnabled As Boolean = False
  Dim StatusCode As Copient.LogixInc.STATUS_FLAGS = Copient.LogixInc.STATUS_FLAGS.STATUS_UNKNOWN
  Dim StatusText As String = ""
  Dim i As Integer = 0

  Dim objTemp As Object
  Dim intNumDecimalPlaces As Integer
  Dim decFactor As Decimal
  Dim decTemp As Decimal
  Dim sTemp As String
  Dim bNeedToFormat As Boolean = False
  Dim PrefManInstalled As Boolean = False
  
  Dim bStoreUser As Boolean = False
    
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "offer-con.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  'Store User
  If(MyCommon.Fetch_CM_SystemOption(131) = "1") Then
    MyCommon.QueryStr = "select LocationID from StoreUsers where UserID=" & AdminUserID & ";"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count Then
      bStoreUser = True
    End If
  End If
  
  OfferID = Request.QueryString("OfferID")
  
  If (OfferID = 0) Then
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "offer-gen.aspx?new=New")
  End If
  
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  
  PrefManInstalled = MyCommon.IsIntegrationInstalled(Copient.CommonIncConfigurable.Integrations.PREFERENCE_MANAGER)
  If PrefManInstalled Then MyCommon.Open_PrefManRT()
  
  
  objTemp = MyCommon.Fetch_CM_SystemOption(41)
  If Not (Integer.TryParse(objTemp.ToString, intNumDecimalPlaces)) Then
    intNumDecimalPlaces = 0
  End If
  decFactor = (10 ^ intNumDecimalPlaces)
  
  ConditionID = Request.QueryString("ConditionID")
  ' dig the offer info out of the database
  ' no one clicked anything
  MyCommon.QueryStr = "select OfferID,Name,IsTemplate,FromTemplate,Description,OfferCategoryID,OfferTypeID,ProdStartDate,ProdEndDate,TestStartDate,TestEndDate,TierTypeID,NumTiers,DistPeriod,DistPeriodLimit,DistPeriodVarID,EmployeeFiltering,NonEmployeesOnly,CRMRestricted,LastUpdate,PriorityLevel,EngineID,SharedLimitID,StatusFlag from Offers with (nolock) where OfferID=" & OfferID & " and Deleted=0 and visible=1"
  rst = MyCommon.LRT_Select()
  
  For Each row In rst.Rows
    NumTiers = row.Item("NumTiers")
    Name = row.Item("Name")
    isTemplate = row.Item("IsTemplate")
    FromTemplate = row.Item("FromTemplate")
  Next
  
  ' first we need to find out if the item were acting on is tiered or not
  Dim deleteTiered As Integer
  deleteTiered = 0
  
  MyCommon.QueryStr = "select Tiered from OfferConditions with (NoLock) where ConditionID=" & ConditionID
  rst = MyCommon.LRT_Select
  For Each row In rst.Rows
    If (row.Item("Tiered") = "1") Then
      deleteTiered = 1
    End If
  Next
  
  Send_HeadBegin("term.offer", "term.conditions", OfferID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
%>
<style type="text/css">
.th-details {
  min-width: 90px;
  }
* html .th-details {
  <% Send(IIf(NumTiers > 3, "  width: 100px", "")) %>
  }
</style>
<%
  Send_Scripts()
  Send_HeadEnd()
  
  'Response.Write("Mode: " & Request.QueryString("mode") & "<br />")
  
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
      Conditions = Request.QueryString.GetValues("con")
      LockedStatus = Request.QueryString.GetValues("locked")
      If (Not Conditions Is Nothing AndAlso Not LockedStatus Is Nothing AndAlso Conditions.Length = LockedStatus.Length) Then
        For LoopCtr = 0 To Conditions.GetUpperBound(0)
          ' disallow a field to be both locked and required.
          If (LockedStatus(LoopCtr) = "1") Then
            MyCommon.QueryStr = "update OfferConditions with (RowLock) set Disallow_Edit = " & LockedStatus(LoopCtr) & ", RequiredFromTemplate=0 " & _
                                "where ConditionID=" & Conditions(LoopCtr)
          Else
            MyCommon.QueryStr = "update OfferConditions with (RowLock) set Disallow_Edit = " & LockedStatus(LoopCtr) & " " & _
                                "where ConditionID=" & Conditions(LoopCtr)
          End If
          MyCommon.LRT_Execute()
        Next
      End If
    End If
  ElseIf (Request.QueryString("mode") = "CycleJoinDescription") Then
    MyCommon.QueryStr = "select JoinTypeID from OfferConditions with (NoLock) where ConditionID=" & Request.QueryString("ConditionID")
    rst = MyCommon.LRT_Select
    For Each row In rst.Rows
      If (row.Item("JoinTypeID") < 2) Then
        ' add one
        MyCommon.QueryStr = "Update OfferConditions with (RowLock) set JoinTypeID=" & row.Item("JoinTypeID") + 1 & " where ConditionID=" & Request.QueryString("ConditionID")
      Else
        ' set to zero
        MyCommon.QueryStr = "Update OfferConditions with (RowLock) set JoinTypeID=1 where ConditionID=" & Request.QueryString("ConditionID")
      End If
      MyCommon.LRT_Execute()
    Next
  ElseIf (Request.QueryString("mode") = "Delete") Then
    If (Request.QueryString("ConditionOrder") <> "0") Then
      MyCommon.QueryStr = "update OfferConditions with(rowlock) set ConditionOrder = (ConditionOrder - 1) " & _
      "where ConditionOrder > " & Request.QueryString("ConditionOrder") & " and ConditionOrder > 1 and Tiered=" & deleteTiered & " and OfferID=" & OfferID
      MyCommon.LRT_Execute()
    End If
    MyCommon.QueryStr = "delete from ConditionTiers with(rowlock) where ConditionID=" & Request.QueryString("ConditionID")
    MyCommon.LRT_Execute()
    MyCommon.QueryStr = "update OfferConditions with (RowLock) set Deleted=1 where ConditionID=" & Request.QueryString("ConditionID") & " and OfferID=" & OfferID
    MyCommon.LRT_Execute()
    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-removecondition", LanguageID))
    'Reset And/Or if last  Condition
    MyCommon.QueryStr = "select ConditionID from OfferConditions with (NoLock) where Deleted=0 and ConditionTypeID<>1 and OfferID=" & OfferID
    rt = MyCommon.LRT_Select()
    If rt.Rows.Count = 1 Then
      MyCommon.QueryStr = "update OfferConditions with (RowLock) set JoinTypeID=1 where Deleted=0 and ConditionTypeID<>1 and ConditionOrder=1 and OfferID=" & OfferID
      MyCommon.LRT_Execute()
    End If
  ElseIf (Request.QueryString("mode") = "Up") Then
    MyCommon.QueryStr = "update OfferConditions with (RowLock) set ConditionOrder=" & Request.QueryString("ConditionOrder") & _
    " where OfferID=" & OfferID & " and Tiered=" & deleteTiered & " and ConditionOrder > 0 and ConditionOrder=" & Request.QueryString("ConditionOrder") - 1
    MyCommon.LRT_Execute()
    MyCommon.QueryStr = "update OfferConditions with (RowLock) set ConditionOrder=" & Request.QueryString("ConditionOrder") - 1 & _
    " where OfferID=" & OfferID & " and Tiered=" & deleteTiered & " and ConditionOrder > 0 and ConditionID=" & Request.QueryString("ConditionID")
    MyCommon.LRT_Execute()
  ElseIf (Request.QueryString("mode") = "Down") Then
    MyCommon.QueryStr = "update OfferConditions with (RowLock) set ConditionOrder=" & Request.QueryString("ConditionOrder") & " where OfferID=" & OfferID & " and Tiered=" & deleteTiered & " and ConditionOrder > 0 and ConditionOrder=" & Request.QueryString("ConditionOrder") + 1
    MyCommon.LRT_Execute()
    MyCommon.QueryStr = "update OfferConditions with (RowLock) set ConditionOrder=" & Request.QueryString("ConditionOrder") + 1 & " where OfferID=" & OfferID & " and Tiered=" & deleteTiered & " and ConditionOrder > 0  and ConditionID=" & Request.QueryString("ConditionID")
    MyCommon.LRT_Execute()
  ElseIf (Request.QueryString("addtier") <> "" Or Request.QueryString("addGlobal") <> "") Then
    'dbo.pt_OfferConditions_Insert @OfferID bigint, @ConditionTypeID int, @Tiered bit, @ConditionOrder int, @ConditionID bigint OUTPUT
    MyCommon.QueryStr = "dbo.pt_OfferConditions_Insert"
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
    If (Request.QueryString("addGlobal") <> "") Then
      MyCommon.LRTsp.Parameters.Add("@ConditionTypeID", SqlDbType.Int).Value = MyCommon.Extract_Val(Request.QueryString("newconglobal"))
      MyCommon.LRTsp.Parameters.Add("@Tiered", SqlDbType.Bit).Value = 0
    Else
      MyCommon.LRTsp.Parameters.Add("@ConditionTypeID", SqlDbType.Int).Value = MyCommon.Extract_Val(Request.QueryString("newcontiered"))
      MyCommon.LRTsp.Parameters.Add("@Tiered", SqlDbType.Bit).Value = 1
    End If
    MyCommon.LRTsp.Parameters.Add("@TCRMAStatusFlag", SqlDbType.Int).Value = 3
    MyCommon.LRTsp.Parameters.Add("@ConditionID", SqlDbType.BigInt).Direction = ParameterDirection.Output
    MyCommon.LRTsp.ExecuteNonQuery()
    ConditionID = MyCommon.LRTsp.Parameters("@ConditionID").Value
    MyCommon.Close_LRTsp()
    If (Request.QueryString("addtier") <> "") Then
      Dim x As Integer
      For x = 1 To NumTiers
        'Response.Write("->@ConditionID: " & ConditionID & " @TierLevel: " & x & " @AmtRequired: " & Request.QueryString("tier" & x) & "<br />")
        MyCommon.QueryStr = "dbo.pt_ConditionTiers_Update"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@ConditionID", SqlDbType.BigInt).Value = ConditionID
        MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int).Value = x
        MyCommon.LRTsp.Parameters.Add("@AmtRequired", SqlDbType.Decimal).Value = 0
        MyCommon.LRTsp.ExecuteNonQuery()
        MyCommon.Close_LRTsp()
      Next
    Else
      If (Request.QueryString("newconglobal") = "2" Or Request.QueryString("newconglobal") = "4" Or Request.QueryString("newconglobal") = "3") Then
        ' we need to set default values if its tender or product
        'dbo.pt_ConditionTiers_Update @ConditionID bigint, @TierLevel int,@AmtRequired decimal(12,3)
        MyCommon.QueryStr = "dbo.pt_ConditionTiers_Update"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@ConditionID", SqlDbType.BigInt).Value = ConditionID
        MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int).Value = 0
        If (Request.QueryString("newconglobal") = "2") Then
          MyCommon.LRTsp.Parameters.Add("@AmtRequired", SqlDbType.Decimal).Value = 1
        Else
          MyCommon.LRTsp.Parameters.Add("@AmtRequired", SqlDbType.Decimal).Value = 0
        End If
        MyCommon.LRTsp.ExecuteNonQuery()
        MyCommon.Close_LRTsp()
      End If
    End If
    If (Request.QueryString("addGlobal") <> "") Then
      If (Request.QueryString("newconglobal") = 1) Then
        Send("<script type=""text/javascript"">openPopup('offer-con-customer.aspx?ConditionID=" & ConditionID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, "Added global customer condition")
      ElseIf (Request.QueryString("newconglobal") = 2) Then
        Send("<script type=""text/javascript"">openPopup('offer-con-product.aspx?ConditionID=" & ConditionID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, "Added global product condition")
      ElseIf (Request.QueryString("newconglobal") = 3) Then
        Send("<script type=""text/javascript"">openPopup('offer-con-point.aspx?ConditionID=" & ConditionID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, "Added global points condition")
      ElseIf (Request.QueryString("newconglobal") = 4) Then
        Send("<script type=""text/javascript"">openPopup('offer-con-tender.aspx?ConditionID=" & ConditionID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, "Added global tender condition")
      ElseIf (Request.QueryString("newconglobal") = 5) Then
        Send("<script type=""text/javascript"">openPopup('offer-con-time.aspx?ConditionID=" & ConditionID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, "Added global time condition")
      ElseIf (Request.QueryString("newconglobal") = 6) Then
        Send("<script type=""text/javascript"">openPopup('offer-con-SV.aspx?ConditionID=" & ConditionID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, "Added global SV condition")
      ElseIf (Request.QueryString("newconglobal") = 7) Then
        Send("<script type=""text/javascript"">openPopup('CM-offer-con-advlimit.aspx?ConditionID=" & ConditionID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, "Added global Advanced Limit condition")
      ElseIf (Request.QueryString("newconglobal") = 100) Then
        Send("<script type=""text/javascript"">openPopup('offer-con-preference.aspx?ConditionID=" & ConditionID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, "Added global preference condition")
      End If
    Else
      If (Request.QueryString("newcontiered") = 2) Then
        Send("<script type=""text/javascript"">openPopup('offer-con-product.aspx?ConditionID=" & ConditionID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, "Added tiered product condition")
      ElseIf (Request.QueryString("newcontiered") = 3) Then
        Send("<script type=""text/javascript"">openPopup('offer-con-point.aspx?ConditionID=" & ConditionID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, "Added tiered points condition")
      ElseIf (Request.QueryString("newcontiered") = 4) Then
        Send("<script type=""text/javascript"">openPopup('offer-con-tender.aspx?ConditionID=" & ConditionID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, "Added tiered tender condition")
      ElseIf (Request.QueryString("newcontiered") = 6) Then
        Send("<script type=""text/javascript"">openPopup('offer-con-SV.aspx?ConditionID=" & ConditionID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, "Added tiered SV condition")
      ElseIf (Request.QueryString("newcontiered") = 7) Then
        Send("<script type=""text/javascript"">openPopup('CM-offer-con-advlimit.aspx?ConditionID=" & ConditionID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, "Added tiered Advanced Limit condition")
      End If
    End If
  End If
  
  ' decide if we need to update the status flags
  If (Request.QueryString("addtier") <> "" Or Request.QueryString("addGlobal") <> "" Or Request.QueryString("mode") = "Down" Or Request.QueryString("mode") = "up" Or Request.QueryString("mode") = "Delete" Or Request.QueryString("mode") = "CycleJoinDescription") Then
    MyCommon.QueryStr = "update OfferConditions with (RowLock) set CRMAStatusFlag=2 where ConditionID=" & ConditionID
    MyCommon.LRT_Execute()
    MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where offerid=" & OfferID
    MyCommon.LRT_Execute()
  End If
  
  ' update the tcrm status flag
  If (Request.QueryString("addtier") <> "" Or Request.QueryString("addGlobal") <> "" Or Request.QueryString("mode") = "Delete") Then
    MyCommon.QueryStr = "update OfferConditions with (RowLock) set TCRMAStatusFlag=3 where ConditionID=" & ConditionID
    MyCommon.LRT_Execute()
    MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where offerid=" & OfferID
    MyCommon.LRT_Execute()
  ElseIf (Request.QueryString("mode") = "CycleJoinDescription") Then
    MyCommon.QueryStr = "update OfferConditions with (RowLock) set TCRMAStatusFlag=2 where ConditionID=" & ConditionID
    MyCommon.LRT_Execute()
    MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where offerid=" & OfferID
    MyCommon.LRT_Execute()
  End If
      
  If (Request.QueryString("addGlobal") <> "" Or Request.QueryString("addtier") <> ""  Or Request.QueryString("mode") = "Delete") Then
    Send("<script type=""text/javascript"">window.location=""offer-con.aspx?OfferID=" & OfferID & """</script>")
  End If
  
  If (isTemplate Or FromTemplate) Then
    ' lets dig the permissions if its a template
    MyCommon.QueryStr = "select * from templatepermissions with (NoLock) where OfferID=" & OfferID
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      For Each row In rst.Rows
        ' ok there are some rows for the template
        Disallow_Conditions = MyCommon.NZ(row.Item("Disallow_Conditions"), True)
        Disallow_OptIn = CMS.Utilities.NZ(row.Item("Disallow_Optin"), True)
      Next
    End If
    ColumnCount = 8
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
  If (isTemplate) Then
    Send_Subtabs(Logix, 22, 5, , OfferID)
  Else
    Send_Subtabs(Logix, 21, 5, , OfferID)
  End If
  
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

function updateLocked(elemName, bChecked) {
  var elem = document.getElementById(elemName);
  
  if (elem != null) {
    elem.value = (bChecked) ? "1" : "0";
  }
}
function submitform() {
  document.getElementById('save1').click();
}

</script>
  <div id="intro">
    <%
      If (isTemplate) Then
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & ": " & MyCommon.TruncateString(Name, 43) & "</h1>")
      Else
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & ": " & MyCommon.TruncateString(Name, 43) & "</h1>")
      End If
    %>
    <div id="controls">
      <%
        If (Logix.UserRoles.EditTemplates And isTemplate) Then
          Send_Save("onclick='submitform();'")
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
      MyCommon.QueryStr = "select StatusFlag from Offers where OfferID=" & OfferID & ";"
      rst3 = MyCommon.LRT_Select
      StatusText = Logix.GetOfferStatus(OfferID, LanguageID, StatusCode)
      If Not isTemplate Then
        If (MyCommon.NZ(rst3.Rows(0).Item("StatusFlag"), 0) <> 2) Then
          If (StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_ACTIVE) AndAlso (MyCommon.NZ(rst3.Rows(0).Item("StatusFlag"), 0) > 0) Then
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
        MyCommon.QueryStr = "select OfferID from Offers with (NoLock) where CreatedDate = LastUpdate and OfferID=" & OfferID
        rst3 = MyCommon.LRT_Select
        If (rst3.Rows.Count = 0) Then
          Send_Status(OfferID)
        End If
      End If
    %>
    <div id="column">
    <form runat="server" id="form1">     
    <uc:ucOptInCondition ID="ucOfferEligibilityCondition" runat="server" AppName="Offer-con.aspx" />  
    </form>

    <form action="offer-con.aspx" id="mainform" name="mainform">
      <input type="submit" name="save" id="save1" style="display:none" />
      <input type="hidden" id="OfferID" name="OfferID" value="<% Sendb(OfferID) %>" />
      <input type="hidden" id="IsOptInPanelLocked" name="IsOptInPanelLocked"  value="<%=iif(Disallow_OptIn , 1,0)%>"/>
      <input type="hidden" id="IsTemplate" name="IsTemplate" value="<% 
        If (IsTemplate) Then
        Sendb("IsTemplate")
        Else
        Sendb("Not")
        End If
        %>" />

      <div class="box" id="conditions">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.conditions", LanguageID))%>
          </span>
        </h2>
        <% 
          Dim TiersWidth As Integer
          TiersWidth = (NumTiers * 60)
        %>
        <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.conditions", LanguageID)) %>">
          <thead>
            <tr>
              <th align="left" scope="col" class="th-reorder">
                <% Sendb(Copient.PhraseLib.Lookup("term.reorder", LanguageID))%>
              </th>
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
              <th align="left" scope="col" class="th-test">
                <% Sendb(Copient.PhraseLib.Lookup("term.test", LanguageID))%>
              </th>
              <th align="left" scope="col" class="th-value" colspan="<% If(NumTiers=0) Then Sendb("1") Else Sendb(NumTiers) %>"
                <% if(numtiers<=3) then else sendb(" style=""width:" & tierswidth & "px;""")%>>
                <% Sendb(Copient.PhraseLib.Lookup("term.value", LanguageID))%>
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
              <td colspan="<% If (NumTiers = 0) Then Sendb(ColumnCount) Else Sendb(NumTiers + 6) %>">
                <h3>
                  <%  Sendb(Copient.PhraseLib.Lookup("term.customercondition", LanguageID))%>
                </h3>
              </td>
            </tr>
            <%' lets see if there are any customer conditions to display
              MyCommon.QueryStr = "select OfferID,O.ConditionID,PTS.ProgramName,PTS.ProgramID,SVP.SVProgramID,SVP.Name as SVProgramName," & _
                                  "C.Description as ConditionDescription,O.ConditionTypeID," & _
                                  "O.GrantTypeID, O.Disallow_Edit, G.Description as GrantDescription, G.PhraseID as GrantPhraseID, O.RequiredFromTemplate, " & _
                                  "QtyUnitType,ConditionOrder,LinkID,CT.AmtRequired,CG.Name as CGName,CG.PhraseID as CGPhraseID,CG.CustomerGroupID as CGID,CG.NewCardholders as CGNewCardholders,ExcludedID,CGE.Name as EXName," & _
                                  "CGE.CustomerGroupID as EXID,PG.Name as PGName,PGE.Name as PGEName,U.Description as UnitDescription,J.PhraseID as JoinDescPhraseID,QtyUnitType, CG.IsOptInGroup from OfferConditions as O " & _
                                  "with (nolock) left join ConditionTypes as C with (nolock) on O.ConditionTypeID=C.ConditionTypeID " & _
                                  "left join GrantTypes as G with (nolock) on O.GrantTypeID=G.GrantTypeID " & _
                                  "left join JoinTypes as J with (nolock) on O.JoinTypeID=J.JoinTypeID " & _
                                  "left join UnitTypes as U with (nolock) on O.QtyUnitType=U.UnitTypeID " & _
                                  "left join CustomerGroups as CG with (nolock) on O.LinkID=CG.CustomerGroupID " & _
                                  "left join CustomerGroups as CGE with (nolock) on O.ExcludedID=CGE.CustomerGroupID " & _
                                  "left join ProductGroups as PG with (nolock) on O.LinkID=PG.ProductGroupID " & _
                                  "left join ProductGroups as PGE with (nolock) on O.ExcludedID=PGE.ProductGroupID " & _
                                  "left join PointsPrograms as PTS with (nolock) on O.LinkID=PTS.ProgramID " & _
                                  "left join StoredValuePrograms as SVP with (nolock) on O.LinkID=SVP.SVProgramID " & _
                                  "left join CM_AdvancedLimits as AL with (nolock) on O.LinkID=AL.LimitID " & _
                                  "left join ConditionTiers as CT with (nolock) on O.ConditionID=CT.ConditionID and CT.TierLevel=0 " & _
                                  "where O.OfferID=" & OfferID & " and O.Tiered=0 and O.deleted=0 and conditionorder=0 order by ConditionOrder, O.ConditionTypeID, O.ConditionID "
              rst = MyCommon.LRT_Select
              prog = 1
              For Each row In rst.Rows
            %>
            <tr class="shaded">
              <td>
                <% Sendb("<input class=""up"" type=""button"" value=""&#9650;"" name=""up"" title=""" & Copient.PhraseLib.Lookup("term.moveup", LanguageID) & """ disabled=""disabled"" />")%>
                <% Sendb("<input class=""down"" type=""button"" value=""&#9660;"" name=""down"" title=""" & Copient.PhraseLib.Lookup("term.movedown", LanguageID) & """ disabled=""disabled"" />")%>
              </td>
              <td>
                <%
                  If Not (IsTemplate) Then
                    If (Logix.UserRoles.EditOffer And Not (FromTemplate And MyCommon.NZ(row.Item("Disallow_Edit"), False))) Then
                      If Utilities.NZ(row.Item("ConditionTypeID"),0) = 1 AndAlso Utilities.NZ( row.Item("IsOptInGroup") , False)  Then
                        Sendb("<input class=""ex"" type=""button"" value=""X"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('offer-con.aspx?mode=Delete&ConditionOrder=" & row.Item("ConditionOrder") & "&ConditionID=" & row.Item("ConditionID") & "&OfferID=" & OfferID & "')}else{return false}"" />")
                      Else
                        Sendb("<input class=""ex"" type=""button"" value=""X"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('offer-con.aspx?mode=Delete&ConditionOrder=" & row.Item("ConditionOrder") & "&ConditionID=" & row.Item("ConditionID") & "&OfferID=" & OfferID & "')}else{return false}"" />")
                      End If
                        
                    Else
                        Sendb("<input class=""ex"" type=""button"" value=""X"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('offer-con.aspx?mode=Delete&ConditionOrder=" & row.Item("ConditionOrder") & "&ConditionID=" & row.Item("ConditionID") & "&OfferID=" & OfferID & "')}else{return false}"" />")
                      End If
                    Else
                    If (Logix.UserRoles.EditTemplates) Then
                      If Utilities.NZ(row.Item("ConditionTypeID"),0) = 1 AndAlso Utilities.NZ( row.Item("IsOptInGroup") , False) Then
                        Sendb("<input class=""ex"" type=""button"" value=""X"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('offer-con.aspx?mode=Delete&ConditionOrder=" & row.Item("ConditionOrder") & "&ConditionID=" & row.Item("ConditionID") & "&OfferID=" & OfferID & "')}else{return false}"" />")
                      Else
                        Sendb("<input class=""ex"" type=""button"" value=""X"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('offer-con.aspx?mode=Delete&ConditionOrder=" & row.Item("ConditionOrder") & "&ConditionID=" & row.Item("ConditionID") & "&OfferID=" & OfferID & "')}else{return false}"" />")
                      End If
                      
                      
                    Else
                      Sendb("<input class=""ex"" type=""button"" value=""X"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('offer-con.aspx?mode=Delete&ConditionOrder=" & row.Item("ConditionOrder") & "&ConditionID=" & row.Item("ConditionID") & "&OfferID=" & OfferID & "')}else{return false}"" />")
                    End If
                    End If
                %>
              </td>
              <td>
                <%
                  If prog > 1 Then
                    Sendb(Copient.PhraseLib.Lookup("term.and", LanguageID))
                  End If
                %>
              </td>
              <td>
                <% 
                  If (row.Item("ConditionTypeID") = 1) Then
                    Send("<a class=""hidden"" href=""offer-con-customer.aspx?ConditionID=" & row.Item("conditionid") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                    Send("<a href=""javascript:openPopup('offer-con-customer.aspx?ConditionID=" & row.Item("conditionid") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.customer", LanguageID) & "</a>")
                  ElseIf (row.Item("ConditionTypeID") = 100) AndAlso PrefManInstalled Then
                    Send("<a class=""hidden"" href=""offer-con-preference.aspx?ConditionID=" & row.Item("conditionid") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                    Send("<a href=""javascript:openPopup('offer-con-preference.aspx?ConditionID=" & row.Item("conditionid") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.preference", LanguageID) & "</a>")
                  End If
                %>
              </td>
              <td>
                <%
                  If (row.Item("ConditionTypeID") = 1) Then
                    If (Int(MyCommon.NZ(row.Item("CGID"), 0)) <= 2) Or (MyCommon.NZ(row.Item("CGNewCardholders"), 0)) Then
                      If (MyCommon.NZ(row.Item("CGPhraseID"), 0) > 0) Then
                        Sendb(Copient.PhraseLib.Lookup(row.Item("CGPhraseID"), LanguageID))
                      Else
                        Sendb(MyCommon.NZ(row.Item("CGName"), "&nbsp;") & "<br />")
                      End If
                      If ((isTemplate Or FromTemplate) AndAlso Int(MyCommon.NZ(row.Item("CGID"), 0)) = 0 AndAlso MyCommon.NZ(row.Item("RequiredFromTemplate"), False)) Then
                        Send("<span class=""red"">* " & Copient.PhraseLib.Lookup("term.required", LanguageID) & " " & Copient.PhraseLib.Lookup("term.by", LanguageID) & " " & Copient.PhraseLib.Lookup("term.template", LanguageID) & "</span>")
                      End If
                    Else
                      Sendb("<a href=""cgroup-edit.aspx?CustomerGroupID=" & row.Item("CGID") & """>")
                      If (MyCommon.NZ(row.Item("CGPhraseID"), 0) > 0) Then
                        Sendb(Copient.PhraseLib.Lookup(row.Item("CGPhraseID"), LanguageID))
                      Else
                        Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("CGName"), "&nbsp;"), 25))
                      End If
                      Send("</a><br />")
                    End If
                    If (MyCommon.NZ(row.Item("EXName"), "") <> "") Then
                      If (Int(MyCommon.NZ(row.Item("EXID"), 0)) <= 2) Then
                        Send("&nbsp;<i>" & Copient.PhraseLib.Lookup("term.excluding", LanguageID) & "</i> " & row.Item("EXName"))
                      Else
                        Send("&nbsp;<i>" & Copient.PhraseLib.Lookup("term.excluding", LanguageID) & "</i> <a href=""cgroup-edit.aspx?CustomerGroupID=" & row.Item("EXID") & """>" & MyCommon.SplitNonSpacedString(row.Item("EXName"), 25) & "</a>")
                      End If
                    End If
                  ElseIf (row.Item("ConditionTypeID") = 100) AndAlso PrefManInstalled Then
                    Send_Preference_Details(MyCommon, MyCommon.NZ(row.Item("LinkID"), 0))
                  End If
                %>
              </td>
              <td>
                &nbsp;</td>
              <td colspan="<% if(numtiers=0)then sendb("1") else sendb(NumTiers) %>">
              <%If (row.Item("ConditionTypeID") = 100) AndAlso PrefManInstalled Then
                  Send_Preference_Info(MyCommon, row.Item("conditionid"))
                Else
                  Send("&nbsp;")
                End If
              %>
              </td>
              <% If (isTemplate) Then%>
              <td class="templine">
                <input type="checkbox" id="chkLocked1" name="chkLocked" value="<%Sendb(row.Item("conditionid"))%>"
                  <%Sendb(IIf(MyCommon.NZ(row.Item("Disallow_Edit"), False)=True, " checked=""checked""", "")) %>
                  onclick="javascript:updateLocked('lock<%Sendb(row.Item("conditionid"))%>', this.checked);" />
                <input type="hidden" id="con<%Sendb(row.Item("conditionid"))%>" name="con" value="<%Sendb(row.Item("conditionid"))%>" />
                <input type="hidden" id="lock<%Sendb(row.Item("conditionid"))%>" name="locked" value="<%Sendb(IIf(MyCommon.NZ(row.Item("Disallow_Edit"), False)=True, "1", "0"))%>" />
              </td>
              <%ElseIf (FromTemplate) Then%>
              <td class="templine">
                <%Sendb(IIf(MyCommon.NZ(row.Item("Disallow_Edit"), False) = True, "Yes", "No"))%>
              </td>
              <% End If%>
            </tr>
            <% 
              prog += 1
            Next
            %>
            <tr class="shadeddark">
              <td colspan="<% If (NumTiers = 0) Then Sendb(ColumnCount) Else Sendb(NumTiers + 6) %>">
                <h3>
                  <% Sendb(Copient.PhraseLib.Lookup("term.globalconditions", LanguageID))%>
                </h3>
              </td>
            </tr>
            <%' lets see if there are any global conditions to display 
              MyCommon.QueryStr = "select OfferID,O.ConditionID,PTS.ProgramName,PTS.ProgramID,SVP.SVProgramID,SVP.Name as SVProgramName," & _
                                  "AL.LimitID,AL.Name as AlName," & _
                                  "C.Description as ConditionDescription,O.ConditionTypeID," & _
                                  "O.GrantTypeID, O.Disallow_Edit, G.Description as GrantDescription, G.PhraseID as GrantPhraseID, O.RequiredFromTemplate," & _
                                  "QtyUnitType,ConditionOrder,LinkID,CT.AmtRequired,CG.Name as CGName,ExcludedID,CGE.Name as EXName,PG.Name as PGName,PG.PhraseID as PGPhraseID," & _
                                  "PG.ProductGroupID as PGID,PGE.Name as PGEName,PGE.PhraseID as PGEPhraseID,PGE.ProductGroupID as PGEID,U.Description as UnitDescription," & _
                                  "J.PhraseID as JoinDescPhraseID,QtyUnitType from OfferConditions as O " & _
                                  "with (nolock) left join ConditionTypes as C with (nolock) on O.ConditionTypeID=C.ConditionTypeID " & _
                                  "left join GrantTypes as G with (nolock) on O.GrantTypeID=G.GrantTypeID " & _
                                  "left join JoinTypes as J with (nolock) on O.JoinTypeID=J.JoinTypeID " & _
                                  "left join UnitTypes as U with (nolock) on O.QtyUnitType=U.UnitTypeID " & _
                                  "left join CustomerGroups as CG with (nolock) on O.LinkID=CG.CustomerGroupID " & _
                                  "left join CustomerGroups as CGE with (nolock) on O.ExcludedID=CGE.CustomerGroupID " & _
                                  "left join ProductGroups as PG with (nolock) on O.LinkID=PG.ProductGroupID " & _
                                  "left join ProductGroups as PGE with (nolock) on O.ExcludedID=PGE.ProductGroupID " & _
                                  "left join PointsPrograms as PTS with (nolock) on O.LinkID=PTS.ProgramID " & _
                                  "left join StoredValuePrograms as SVP with (nolock) on O.LinkID=SVP.SVProgramID " & _
                                  "left join CM_AdvancedLimits as AL with (nolock) on O.LinkID=AL.LimitID " & _
                                  "left join ConditionTiers as CT with (nolock) on O.ConditionID=CT.ConditionID and CT.TierLevel=0 " & _
                                  "where O.OfferID=" & OfferID & " and O.Tiered=0 and O.deleted=0 and conditionorder > 0 order by ConditionOrder"
              rst = MyCommon.LRT_Select
              prog = 1
              For Each row In rst.Rows
            %>
            <tr class="shaded">
              <td>
                <%
                  ' if(FromTemplate and Disallow_EmployeeFiltering)then sendb("disabled=""disabled""")
                If (Not IsTemplate) Then
                  If (row.Item("ConditionOrder") > 1 And Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Conditions)) Then
                    Sendb("<input class=""up"" type=""button"" value=""&#9650;"" name=""up"" title=""" & Copient.PhraseLib.Lookup("term.moveup", LanguageID) & """ onClick=""LoadDocument('offer-con.aspx?mode=Up&ConditionOrder=" & row.Item("ConditionOrder") & "&ConditionID=" & row.Item("ConditionID") & "&OfferID=" & OfferID & "')"" />")
                  Else
                    Sendb("<input class=""up"" type=""button"" value=""&#9650;"" name=""up"" title=""" & Copient.PhraseLib.Lookup("term.moveup", LanguageID) & """ disabled=""disabled"" />")
                  End If
                  If (prog < rst.Rows.Count And Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Conditions)) Then
                    Sendb("<input class=""down"" type=""button"" value=""&#9660;"" name=""down"" title=""" & Copient.PhraseLib.Lookup("term.movedown", LanguageID) & """ onClick=""LoadDocument('offer-con.aspx?mode=Down&ConditionOrder=" & row.Item("ConditionOrder") & "&ConditionID=" & row.Item("ConditionID") & "&OfferID=" & OfferID & "')"" />")
                  Else
                    Sendb("<input class=""down"" type=""button"" value=""&#9660;"" name=""down"" title=""" & Copient.PhraseLib.Lookup("term.movedown", LanguageID) & """ disabled=""disabled"" />")
                  End If
                  prog = prog + 1
                Else
                  If (row.Item("ConditionOrder") > 1 And Logix.UserRoles.EditTemplates) Then
                    Sendb("<input class=""up"" type=""button"" value=""&#9650;"" name=""up"" title=""" & Copient.PhraseLib.Lookup("term.moveup", LanguageID) & """ onClick=""LoadDocument('offer-con.aspx?mode=Up&ConditionOrder=" & row.Item("ConditionOrder") & "&ConditionID=" & row.Item("ConditionID") & "&OfferID=" & OfferID & "')"" />")
                  Else
                    Sendb("<input class=""up"" type=""button"" value=""&#9650;"" name=""up"" title=""" & Copient.PhraseLib.Lookup("term.moveup", LanguageID) & """ disabled=""disabled"" />")
                  End If
                  If (prog < rst.Rows.Count And Logix.UserRoles.EditTemplates) Then
                    Sendb("<input class=""down"" type=""button"" value=""&#9660;"" name=""down"" title=""" & Copient.PhraseLib.Lookup("term.movedown", LanguageID) & """ onClick=""LoadDocument('offer-con.aspx?mode=Down&ConditionOrder=" & row.Item("ConditionOrder") & "&ConditionID=" & row.Item("ConditionID") & "&OfferID=" & OfferID & "')"" />")
                  Else
                    Sendb("<input class=""down"" type=""button"" value=""&#9660;"" name=""down"" title=""" & Copient.PhraseLib.Lookup("term.movedown", LanguageID) & """ disabled=""disabled"" />")
                  End If
                  prog = prog + 1
                End If 
                %>
              </td>
              <td>
                <%
                  If (Not IsTemplate) Then
                    If (Logix.UserRoles.EditOffer And Not (FromTemplate And MyCommon.NZ(row.Item("Disallow_Edit"), False))) Then
                        Sendb("<input class=""ex"" type=""button"" value=""X"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('offer-con.aspx?mode=Delete&ConditionOrder=" & row.Item("ConditionOrder") & "&ConditionID=" & row.Item("ConditionID") & "&OfferID=" & OfferID & "')}else{return false}"" />")
                                      
                          
                      
                    Else
                      Sendb("<input class=""ex"" type=""button"" value=""X"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('offer-con.aspx?mode=Delete&ConditionOrder=" & row.Item("ConditionOrder") & "&ConditionID=" & row.Item("ConditionID") & "&OfferID=" & OfferID & "')}else{return false}"" />")
                    End If
                  Else
                    If (Logix.UserRoles.EditTemplates) Then
                      Sendb("<input class=""ex"" type=""button"" value=""X"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('offer-con.aspx?mode=Delete&ConditionOrder=" & row.Item("ConditionOrder") & "&ConditionID=" & row.Item("ConditionID") & "&OfferID=" & OfferID & "')}else{return false}"" />")
                    Else
                      Sendb("<input class=""ex"" type=""button"" value=""X"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('offer-con.aspx?mode=Delete&ConditionOrder=" & row.Item("ConditionOrder") & "&ConditionID=" & row.Item("ConditionID") & "&OfferID=" & OfferID & "')}else{return false}"" />")
                    End If
                  End If
                %>
              </td>
              <td>
                <% 
                If Not(IsTemplate) Then
                  If (row.Item("ConditionOrder") > 1 And Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Conditions)) Then
                    Sendb("<a href=""offer-con.aspx?mode=CycleJoinDescription&ConditionID=" & row.Item("ConditionID") & "&OfferID=" & OfferID & """>" & Copient.PhraseLib.Lookup(row.Item("JoinDescPhraseID"), LanguageID) & "</a>")
                  ElseIf (row.Item("ConditionOrder") > 1) Then
                    Sendb(Copient.PhraseLib.Lookup(row.Item("JoinDescPhraseID"), LanguageID))
                  Else
                    Sendb("&nbsp;")
                  End If
                Else
                  If (row.Item("ConditionOrder") > 1 And Logix.UserRoles.EditTemplates) Then
                    Sendb("<a href=""offer-con.aspx?mode=CycleJoinDescription&ConditionID=" & row.Item("ConditionID") & "&OfferID=" & OfferID & """>" & Copient.PhraseLib.Lookup(row.Item("JoinDescPhraseID"), LanguageID) & "</a>")
                  ElseIf (row.Item("ConditionOrder") > 1) Then
                    Sendb(Copient.PhraseLib.Lookup(row.Item("JoinDescPhraseID"), LanguageID))
                  Else
                    Sendb("&nbsp;")
                  End If
                End If
                %>
              </td>
              <td>
                <% 
                  If (row.Item("ConditionTypeID") = 1) Then
                    Send("<a class=""hidden"" href=""offer-con-customer.aspx?ConditionID=" & row.Item("conditionid") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                    Send("<a href=""javascript:openPopup('offer-con-customer.aspx?ConditionID=" & row.Item("conditionid") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.customer", LanguageID) & "</a>")
                  ElseIf (row.Item("ConditionTypeID") = 2) Then
                    Send("<a class=""hidden"" href=""offer-con-product.aspx?ConditionID=" & row.Item("conditionid") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                    Send("<a href=""javascript:openPopup('offer-con-product.aspx?ConditionID=" & row.Item("conditionid") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.product", LanguageID) & "</a>")
                  ElseIf (row.Item("ConditionTypeID") = 3) Then
                    Send("<a class=""hidden"" href=""offer-con-point.aspx?ConditionID=" & row.Item("conditionid") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                    Send("<a href=""javascript:openPopup('offer-con-point.aspx?ConditionID=" & row.Item("conditionid") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.points", LanguageID) & "</a>")
                  ElseIf (row.Item("ConditionTypeID") = 4) Then
                    Send("<a class=""hidden"" href=""offer-con-tender.aspx?ConditionID=" & row.Item("conditionid") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                    Send("<a href=""javascript:openPopup('offer-con-tender.aspx?ConditionID=" & row.Item("conditionid") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.tender", LanguageID) & "</a>")
                  ElseIf (row.Item("ConditionTypeID") = 5) Then
                    Send("<a class=""hidden"" href=""offer-con-time.aspx?ConditionID=" & row.Item("conditionid") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                    Send("<a href=""javascript:openPopup('offer-con-time.aspx?ConditionID=" & row.Item("conditionid") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.timeday", LanguageID) & "</a>")
                  ElseIf (row.Item("ConditionTypeID") = 6) Then
                    Send("<a class=""hidden"" href=""offer-con-SV.aspx?ConditionID=" & row.Item("conditionid") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                    Send("<a href=""javascript:openPopup('offer-con-SV.aspx?ConditionID=" & row.Item("conditionid") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.storedvalue", LanguageID) & "</a>")
                  ElseIf (row.Item("ConditionTypeID") = 7) Then
                    Send("<a class=""hidden"" href=""CM-offer-con-advlimit.aspx?ConditionID=" & row.Item("conditionid") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                    Send("<a href=""javascript:openPopup('CM-offer-con-advlimit.aspx?ConditionID=" & row.Item("conditionid") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.advlimit", LanguageID) & "</a>")
                  End If
                %>
              </td>
              <td>
                <%If (row.Item("ConditionTypeID") = 1) Then
                    Send(MyCommon.NZ(row.Item("CGName"), "&nbsp;") & "<br />")
                    If (MyCommon.NZ(row.Item("EXName"), "") <> "") Then Send("&nbsp;<i>" & Copient.PhraseLib.Lookup("term.excluding", LanguageID) & "</i> " & row.Item("EXName"))
                  ElseIf (row.Item("ConditionTypeID") = 2) Then
                    If (Int(MyCommon.NZ(row.Item("PGID"), 0)) = 1) Or (Int(MyCommon.NZ(row.Item("PGID"), 0)) = 2) Then
                      If (MyCommon.NZ(row.Item("PGPhraseID"), 0) > 0) Then
                        Send(Copient.PhraseLib.Lookup(row.Item("PGPhraseID"), LanguageID) & "<br />")
                      Else
                      Send(MyCommon.NZ(row.Item("PGName"), "&nbsp;") & "<br />")
                      End If
                    ElseIf (Int(MyCommon.NZ(row.Item("PGID"), 0)) = 0) Then
                      Send(Copient.PhraseLib.Lookup("term.entiretransaction", LanguageID) & "<br />")
                      If ((isTemplate Or FromTemplate) AndAlso MyCommon.NZ(row.Item("RequiredFromTemplate"), False)) Then
                        Send("<span class=""red"">* " & Copient.PhraseLib.Lookup("term.required", LanguageID) & " " & Copient.PhraseLib.Lookup("term.by", LanguageID) & " " & Copient.PhraseLib.Lookup("term.template", LanguageID) & "</span>")
                      End If
                    Else
                      Sendb("<a href=""pgroup-edit.aspx?ProductGroupID=" & row.Item("PGID") & """>")
                      If (MyCommon.NZ(row.Item("PGPhraseID"), 0) > 0) Then
                        Sendb(Copient.PhraseLib.Lookup(row.Item("PGPhraseID"), LanguageID))
                      Else
                        Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("PGName"), "&nbsp;"), 25))
                    End If
                      Send("</a><br />")
                    End If
                    If (MyCommon.NZ(row.Item("PGEName"), "") <> "") Then
                      If (Int(MyCommon.NZ(row.Item("PGEID"), 0)) <= 1) Then
                        Send("&nbsp;<i>" & Copient.PhraseLib.Lookup("term.excluding", LanguageID) & "</i> " & row.Item("PGEName"))
                      Else
                        Send("&nbsp;<i>" & Copient.PhraseLib.Lookup("term.excluding", LanguageID) & "</i> <a href=""pgroup-edit.aspx?ProductGroupID=" & row.Item("PGEID") & """>" & MyCommon.SplitNonSpacedString(row.Item("PGEName"), 25) & "</a>")
                      End If
                    End If
                  ElseIf (row.Item("ConditionTypeID") = 3) Then
                    If (Int(MyCommon.NZ(row.Item("ProgramID"), 0)) < 1) Then
                      Send("&nbsp;")
                    Else
                      Send("<a href=""point-edit.aspx?ProgramGroupID=" & row.Item("ProgramID") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("ProgramName"), "&nbsp;"), 25) & "</a><br />")
                    End If
                  ElseIf (row.Item("ConditionTypeID") = 6) Then
                    If (Int(MyCommon.NZ(row.Item("SVProgramID"), 0)) < 1) Then
                      Send("&nbsp;")
                    Else
                      Send("<a href=""sv-edit.aspx?ProgramGroupID=" & row.Item("SVProgramID") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("SVProgramName"), "&nbsp;"), 25) & "</a><br />")
                    End If
                  ElseIf (row.Item("ConditionTypeID") = 7) Then
                    If (Int(MyCommon.NZ(row.Item("LimitID"), 0)) < 1) Then
                      Send("&nbsp;")
                    Else
                      Send("<a href=""CM-advlimit-edit.aspx?LimitID=" & row.Item("LimitID") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("AlName"), "&nbsp;"), 25) & "</a><br />")
                    End If
                  ElseIf (row.Item("ConditionTypeID") = 4) Then
                    MyCommon.QueryStr = "select CTT.TenderTypeID,Description from ConditionTenderTypes as CTT with (NoLock) left join TenderTypes as TT with (NoLock) on TT.TenderTypeID=CTT.TenderTypeID where ConditionID=" & row.Item("ConditionID") & " order by Description"
                    rst2 = MyCommon.LRT_Select
                    i = 0
                    For Each row2 In rst2.Rows
                      i = i + 1
                      Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row2.Item("Description"), ""), 25))
                      If i = (rst2.Rows.Count - 1) Then
                        Send(" " & StrConv(Copient.PhraseLib.Lookup("term.or", LanguageID), VbStrConv.Lowercase) & " ")
                      ElseIf i < rst2.Rows.Count Then
                        Send(", ")
                      Else
                        Send("")
                      End If
                    Next
                  ElseIf (row.Item("ConditionTypeID") = 5) Then
                    MyCommon.QueryStr = "select CTM.Sunday,Monday,Tuesday,Wednesday,Thursday,Friday,Saturday from ConditionTimes as CTM with (NoLock) where ConditionID=" & row.Item("ConditionID")
                    rst2 = MyCommon.LRT_Select
                    Dim DaysTrue As Integer = 0
                    If (rst2.Rows.Count = 0) Then
                      Send(Copient.PhraseLib.Lookup("term.everyday", LanguageID))
                    ElseIf (MyCommon.NZ(rst2.Rows(0).Item("Sunday"), "") AndAlso MyCommon.NZ(rst2.Rows(0).Item("Monday"), "") AndAlso MyCommon.NZ(rst2.Rows(0).Item("Tuesday"), "") AndAlso MyCommon.NZ(rst2.Rows(0).Item("Wednesday"), "") AndAlso MyCommon.NZ(rst2.Rows(0).Item("Thursday"), "") AndAlso MyCommon.NZ(rst2.Rows(0).Item("Friday"), "") AndAlso MyCommon.NZ(rst2.Rows(0).Item("Saturday"), "")) Then
                      Send(Copient.PhraseLib.Lookup("term.everyday", LanguageID))
                      DaysTrue = 7
                    Else
                      For Each row2 In rst2.Rows
                        If (MyCommon.NZ(row2.Item("Sunday"), "")) Then
                          Sendb(Left(Copient.PhraseLib.Lookup("term.sunday", LanguageID), 3) & "&nbsp;")
                          DaysTrue = DaysTrue + 1
                        End If
                        If (MyCommon.NZ(row2.Item("Monday"), "")) Then
                          Sendb(Left(Copient.PhraseLib.Lookup("term.monday", LanguageID), 3) & "&nbsp;")
                          DaysTrue = DaysTrue + 1
                        End If
                        If (MyCommon.NZ(row2.Item("Tuesday"), "")) Then
                          Sendb(Left(Copient.PhraseLib.Lookup("term.tuesday", LanguageID), 3) & "&nbsp;")
                          DaysTrue = DaysTrue + 1
                        End If
                        If (MyCommon.NZ(row2.Item("Wednesday"), "")) Then
                          Sendb(Left(Copient.PhraseLib.Lookup("term.wednesday", LanguageID), 3) & "&nbsp;")
                          DaysTrue = DaysTrue + 1
                        End If
                        If (MyCommon.NZ(row2.Item("Thursday"), "")) Then
                          Sendb(Left(Copient.PhraseLib.Lookup("term.thursday", LanguageID), 3) & "&nbsp;")
                          DaysTrue = DaysTrue + 1
                        End If
                        If (MyCommon.NZ(row2.Item("Friday"), "")) Then
                          Sendb(Left(Copient.PhraseLib.Lookup("term.friday", LanguageID), 3) & "&nbsp;")
                          DaysTrue = DaysTrue + 1
                        End If
                        If (MyCommon.NZ(row2.Item("Saturday"), "")) Then
                          Sendb(Left(Copient.PhraseLib.Lookup("term.saturday", LanguageID), 3) & "&nbsp;")
                          DaysTrue = DaysTrue + 1
                        End If
                      Next
                    End If
                    If (DaysTrue > 0) Then
                      Sendb("<br />")
                    End If
                    MyCommon.QueryStr = "select CTM.StartHour,StartMinute,EndHour,EndMinute from ConditionTimes as CTM with (NoLock) where ConditionID=" & row.Item("ConditionID")
                    rst2 = MyCommon.LRT_Select
                    If (rst2.Rows.Count = 0) Then
                    Else
                      Dim StartHour, EndHour, StartMinute, EndMinute As String
                      For Each row2 In rst2.Rows
                        StartHour = row2.Item("StartHour").ToString.PadLeft(2, "0")
                        EndHour = row2.Item("EndHour").ToString.PadLeft(2, "0")
                        StartMinute = row2.Item("StartMinute").ToString.PadLeft(2, "0")
                        EndMinute = row2.Item("EndMinute").ToString.PadLeft(2, "0")
                        If (StartHour = "00") And (StartMinute = "00") And (EndHour = "00") And (EndMinute = "00") Then
                        Else
                          Sendb(StartHour & ":" & StartMinute)
                          Sendb(" - ")
                          Sendb(EndHour & ":" & EndMinute)
                        End If
                      Next
                    End If
                  End If
                %>
              </td>
              <td>
                <%If (row.Item("ConditionTypeID") = 1) Or (row.Item("ConditionTypeID") = 5) Then
                    Send("&nbsp;")
                  Else
                    If (MyCommon.NZ(row.Item("GrantPhraseID"), 0) = 0) Then
                      Send(row.Item("GrantDescription"))
                    Else
                      Send(Copient.PhraseLib.Lookup(row.Item("GrantPhraseID"), LanguageID))
                    End If
                  End If
                %>
              </td>
              <td colspan="<% if(numtiers=0)then sendb("1") else sendb(NumTiers) %>">
                <%If (MyCommon.NZ(row.Item("ConditionTypeID"), 0) = 4) Then
                    Sendb("$")
                  End If
                  If (MyCommon.NZ(row.Item("ConditionTypeID"), 0) = 1) Or (MyCommon.NZ(row.Item("ConditionTypeID"), 0) = 5) Then
                    Send("&nbsp;")
                  ElseIf (MyCommon.NZ(row.Item("ConditionTypeID"), 0) = 3) Then
                    Send(Int(MyCommon.NZ(row.Item("AmtRequired"), 0)))
                  ElseIf (MyCommon.NZ(row.Item("ConditionTypeID"), 0) = 6) Then
                    bNeedToFormat = False
                    If intNumDecimalPlaces > 0 Then
                      MyCommon.QueryStr = "Select SVTypeID from StoredValuePrograms with (NoLock) where SVProgramID=" & row.Item("LinkID") & ";"
                      rst2 = MyCommon.LRT_Select
                      If (rst2.Rows.Count > 0) Then
                        If Int(MyCommon.NZ(rst2.Rows(0).Item("SVTypeID"), 0)) = 1 Then
                          bNeedToFormat = True
                        End If
                      End If
                    End If
                    If bNeedToFormat Then
                      decTemp = (Int(MyCommon.NZ(row.Item("AmtRequired"), 0)) * 1.0) / decFactor
                      sTemp = FormatNumber(decTemp, intNumDecimalPlaces)
                      Send(sTemp)
                    Else
                      Send(Int(MyCommon.NZ(row.Item("AmtRequired"), 0)))
                    End If
                  ElseIf (MyCommon.NZ(row.Item("ConditionTypeID"), 0) = 7) Then
                    Send(Int(MyCommon.NZ(row.Item("AmtRequired"), 0)))
                  ElseIf (MyCommon.NZ(row.Item("ConditionTypeID"), 0) = 4) Then
                    Send(MyCommon.NZ(row.Item("AmtRequired"), 0))
                  ElseIf (MyCommon.NZ(row.Item("ConditionTypeID"), 0) = 2) Then
                    If (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 2) Or (MyCommon.NZ(row.Item("QtyUnitType"), 0) >= 6) Then
                      Sendb("$")
                    End If
                                        
                    If (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 1) Or (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 5) Then
                      Sendb(Int(MyCommon.NZ(row.Item("AmtRequired"), 0)))
                    Else
                      Sendb(MyCommon.NZ(row.Item("AmtRequired"), 0))
                    End If
                                        
                    If (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 3) Then
                      If MyCommon.NZ(row.Item("AmtRequired"), 0) = 1 Then
                        Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.pound", LanguageID), VbStrConv.Lowercase))
                      Else
                        Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.pounds", LanguageID), VbStrConv.Lowercase))
                      End If
                    ElseIf (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 4) Then
                      If MyCommon.NZ(row.Item("AmtRequired"), 0) = 1 Then
                        Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.gallon", LanguageID), VbStrConv.Lowercase))
                      Else
                        Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.gallons", LanguageID), VbStrConv.Lowercase))
                      End If
                    End If
                  Else
                  End If
                %>
              </td>
              <% If (isTemplate) Then%>
              <td class="templine">
                <input type="checkbox" id="chkLocked2" name="chkLocked" value="<%Sendb(row.Item("conditionid"))%>"
                  <%Sendb(IIf(MyCommon.NZ(row.Item("Disallow_Edit"), False)=True, " checked=""checked""", "")) %>
                  onclick="javascript:updateLocked('lock<%Sendb(row.Item("conditionid"))%>', this.checked);" />
                <input type="hidden" id="con<%Sendb(row.Item("conditionid"))%>" name="con" value="<%Sendb(row.Item("conditionid"))%>" />
                <input type="hidden" id="lock<%Sendb(row.Item("conditionid"))%>" name="locked" value="<%Sendb(IIf(MyCommon.NZ(row.Item("Disallow_Edit"), False)=True, "1", "0"))%>" />
              </td>
              <%ElseIf (FromTemplate) Then%>
              <td class="templine">
                <%Sendb(IIf(MyCommon.NZ(row.Item("Disallow_Edit"), False) = True, "Yes", "No"))%>
              </td>
              <% End If%>
            </tr>
            <%
            Next
            %>
            <% If (NumTiers > 0) Then%>
            <tr class="shadeddark">
              <td colspan="6">
                <h3>
                  <% Sendb(Copient.PhraseLib.Lookup("term.tierconditions", LanguageID))%>
                </h3>
              </td>
              <%
                Dim q As Integer
                For q = 1 To NumTiers
                  Send("<td>" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & "&nbsp;" & q & "</td>")
                Next
                If (isTemplate OrElse FromTemplate) Then
                  Send("<td></td>")
                End If
              %>
            </tr>
            <%' alright, lets get all the tiered conditions and display them
              MyCommon.QueryStr = "select OfferID,O.ConditionID,PTS.ProgramName,PTS.ProgramID,SVP.SVProgramID,SVP.Name as SVProgramName," & _
                                  "AL.LimitID,AL.Name as AlName," & _
                                  "C.Description as ConditionDescription,O.ConditionTypeID," & _
                                  "O.GrantTypeID, O.Disallow_Edit, G.Description as GrantDescription, G.PhraseID as GrantPhraseID, " & _
                                  "QtyUnitType,ConditionOrder,LinkID,CT.AmtRequired,CG.Name as CGName,ExcludedID,CGE.Name as EXName,PG.Name as PGName," & _
                                  "PG.ProductGroupID as PGID,PGE.Name as PGEName,PGE.ProductGroupID as PGEID,U.Description as UnitDescription,J.PhraseID as JoinDescPhraseID,QtyUnitType from OfferConditions as O " & _
                                  "with (nolock) left join ConditionTypes as C with (nolock) on O.ConditionTypeID=C.ConditionTypeID " & _
                                  "left join GrantTypes as G with (nolock) on O.GrantTypeID=G.GrantTypeID " & _
                                  "left join JoinTypes as J with (nolock) on O.JoinTypeID=J.JoinTypeID " & _
                                  "left join UnitTypes as U with (nolock) on O.QtyUnitType=U.UnitTypeID " & _
                                  "left join CustomerGroups as CG with (nolock) on O.LinkID=CG.CustomerGroupID " & _
                                  "left join CustomerGroups as CGE with (nolock) on O.ExcludedID=CGE.CustomerGroupID " & _
                                  "left join ProductGroups as PG with (nolock) on O.LinkID=PG.ProductGroupID " & _
                                  "left join ProductGroups as PGE with (nolock) on O.ExcludedID=PGE.ProductGroupID " & _
                                  "left join PointsPrograms as PTS with (nolock) on O.LinkID=PTS.ProgramID " & _
                                  "left join StoredValuePrograms as SVP with (nolock) on O.LinkID=SVP.SVProgramID " & _
                                  "left join CM_AdvancedLimits as AL with (nolock) on O.LinkID=AL.LimitID " & _
                                  "left join ConditionTiers as CT with (nolock) on O.ConditionID=CT.ConditionID and CT.TierLevel=0 " & _
                                  "where O.OfferID=" & OfferID & " and O.Tiered=1 and O.deleted=0 and conditionorder>0 order by ConditionOrder"
              rst = MyCommon.LRT_Select
              prog = 1
              For Each row In rst.Rows
            %>
            <tr class="shaded">
              <td>
                <%
                If Not (IsTemplate)
                     If (row.Item("ConditionOrder") > 1 And Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Conditions)) Then
                    Sendb("<input  class=""up"" type=""button"" value=""&#9650;"" name=""up"" title=""" & Copient.PhraseLib.Lookup("term.moveup", LanguageID) & """ onClick=""LoadDocument('offer-con.aspx?mode=Up&ConditionOrder=" & row.Item("ConditionOrder") & "&ConditionID=" & row.Item("ConditionID") & "&OfferID=" & OfferID & "')"" />")
                  Else
                    Sendb("<input class=""up"" type=""button"" value=""&#9650;"" name=""up"" title=""" & Copient.PhraseLib.Lookup("term.moveup", LanguageID) & """disabled=""disabled"" />")
                  End If
                  If (prog < rst.Rows.Count And Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Conditions)) Then
                    Sendb("<input class=""down"" type=""button"" value=""&#9660;"" name=""down"" title=""" & Copient.PhraseLib.Lookup("term.movedown", LanguageID) & """ onClick=""LoadDocument('offer-con.aspx?mode=Down&ConditionOrder=" & row.Item("ConditionOrder") & "&ConditionID=" & row.Item("ConditionID") & "&OfferID=" & OfferID & "')"" />")
                  Else
                    Sendb("<input class=""down"" type=""button"" value=""&#9660;"" name=""down"" title=""" & Copient.PhraseLib.Lookup("term.movedown", LanguageID) & """ disabled=""disabled"" />")
                  End If
                  prog = prog + 1
                Else
                  If (row.Item("ConditionOrder") > 1 And Logix.UserRoles.EditTemplates) Then
                    Sendb("<input  class=""up"" type=""button"" value=""&#9650;"" name=""up"" title=""" & Copient.PhraseLib.Lookup("term.moveup", LanguageID) & """ onClick=""LoadDocument('offer-con.aspx?mode=Up&ConditionOrder=" & row.Item("ConditionOrder") & "&ConditionID=" & row.Item("ConditionID") & "&OfferID=" & OfferID & "')"" />")
                  Else
                    Sendb("<input class=""up"" type=""button"" value=""&#9650;"" name=""up"" title=""" & Copient.PhraseLib.Lookup("term.moveup", LanguageID) & """disabled=""disabled"" />")
                  End If
                  If (prog < rst.Rows.Count And Logix.UserRoles.EditTemplates) Then
                    Sendb("<input class=""down"" type=""button"" value=""&#9660;"" name=""down"" title=""" & Copient.PhraseLib.Lookup("term.movedown", LanguageID) & """ onClick=""LoadDocument('offer-con.aspx?mode=Down&ConditionOrder=" & row.Item("ConditionOrder") & "&ConditionID=" & row.Item("ConditionID") & "&OfferID=" & OfferID & "')"" />")
                  Else
                    Sendb("<input class=""down"" type=""button"" value=""&#9660;"" name=""down"" title=""" & Copient.PhraseLib.Lookup("term.movedown", LanguageID) & """ disabled=""disabled"" />")
                  End If
                  prog = prog + 1
                End If
                %>
              </td>
              <td>
                <%
                  If Not (IsTemplate)
                  If (Logix.UserRoles.EditOffer And Not (FromTemplate And MyCommon.NZ(row.Item("Disallow_Edit"), False))) Then
                    Sendb("<input class=""ex"" type=""button"" value=""X"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('offer-con.aspx?mode=Delete&ConditionOrder=" & row.Item("ConditionOrder") & "&ConditionID=" & row.Item("ConditionID") & "&OfferID=" & OfferID & "')}else{return false}"" />")
                  Else
                    Sendb("<input class=""ex"" type=""button"" value=""X"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('offer-con.aspx?mode=Delete&ConditionOrder=" & row.Item("ConditionOrder") & "&ConditionID=" & row.Item("ConditionID") & "&OfferID=" & OfferID & "')}else{return false}"" />")
                  End If
                Else
                  If (Logix.UserRoles.EditTemplates) Then
                    Sendb("<input class=""ex"" type=""button"" value=""X"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('offer-con.aspx?mode=Delete&ConditionOrder=" & row.Item("ConditionOrder") & "&ConditionID=" & row.Item("ConditionID") & "&OfferID=" & OfferID & "')}else{return false}"" />")
                  Else
                    Sendb("<input class=""ex"" type=""button"" value=""X"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('offer-con.aspx?mode=Delete&ConditionOrder=" & row.Item("ConditionOrder") & "&ConditionID=" & row.Item("ConditionID") & "&OfferID=" & OfferID & "')}else{return false}"" />")
                  End If
                End If
                %>
              </td>
              <td>
                <% 
                If Not (IsTemplate)
                  If (row.Item("ConditionOrder") > 1 And Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Conditions)) Then
                    Sendb("<a href=""offer-con.aspx?mode=CycleJoinDescription&ConditionID=" & row.Item("ConditionID") & "&OfferID=" & OfferID & """>" & Copient.PhraseLib.Lookup(row.Item("JoinDescPhraseID"), LanguageID) & "</a>")
                  ElseIf (row.Item("ConditionOrder") > 1) Then
                    Sendb(Copient.PhraseLib.Lookup(row.Item("JoinDescPhraseID"), LanguageID))
                  Else
                    Sendb("&nbsp;")
                  End If
                Else
                  If (row.Item("ConditionOrder") > 1 And Logix.UserRoles.EditTemplates) Then
                    Sendb("<a href=""offer-con.aspx?mode=CycleJoinDescription&ConditionID=" & row.Item("ConditionID") & "&OfferID=" & OfferID & """>" & Copient.PhraseLib.Lookup(row.Item("JoinDescPhraseID"), LanguageID) & "</a>")
                  ElseIf (row.Item("ConditionOrder") > 1) Then
                    Sendb(Copient.PhraseLib.Lookup(row.Item("JoinDescPhraseID"), LanguageID))
                  Else
                    Sendb("&nbsp;")
                  End If
                End If
                %>
              </td>
              <td>
                <% 
                  If (row.Item("ConditionTypeID") = 1) Then
                    Send("<a class=""hidden"" href=""offer-con-customer.aspx?ConditionID=" & row.Item("conditionid") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                    Send("<a href=""javascript:openPopup('offer-con-customer.aspx?ConditionID=" & row.Item("conditionid") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.customer", LanguageID) & "</a>")
                  ElseIf (row.Item("ConditionTypeID") = 2) Then
                    Send("<a class=""hidden"" href=""offer-con-product.aspx?ConditionID=" & row.Item("conditionid") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                    Send("<a href=""javascript:openPopup('offer-con-product.aspx?ConditionID=" & row.Item("conditionid") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.product", LanguageID) & "</a>")
                  ElseIf (row.Item("ConditionTypeID") = 3) Then
                    Send("<a class=""hidden"" href=""offer-con-point.aspx?ConditionID=" & row.Item("conditionid") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                    Send("<a href=""javascript:openPopup('offer-con-point.aspx?ConditionID=" & row.Item("conditionid") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.points", LanguageID) & "</a>")
                  ElseIf (row.Item("ConditionTypeID") = 4) Then
                    Send("<a class=""hidden"" href=""offer-con-tender.aspx?ConditionID=" & row.Item("conditionid") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                    Send("<a href=""javascript:openPopup('offer-con-tender.aspx?ConditionID=" & row.Item("conditionid") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.tender", LanguageID) & "</a>")
                  ElseIf (row.Item("ConditionTypeID") = 5) Then
                    Send("<a class=""hidden"" href=""offer-con-time.aspx?ConditionID=" & row.Item("conditionid") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                    Send("<a href=""javascript:openPopup('offer-con-time.aspx?ConditionID=" & row.Item("conditionid") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.timeday", LanguageID) & "</a>")
                  ElseIf (row.Item("ConditionTypeID") = 6) Then
                    Send("<a class=""hidden"" href=""offer-con-SV.aspx?ConditionID=" & row.Item("conditionid") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                    Send("<a href=""javascript:openPopup('offer-con-SV.aspx?ConditionID=" & row.Item("conditionid") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.storedvalue", LanguageID) & "</a>")
                  ElseIf (row.Item("ConditionTypeID") = 7) Then
                    Send("<a class=""hidden"" href=""CM-offer-con-advlimit.aspx?ConditionID=" & row.Item("conditionid") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                    Send("<a href=""javascript:openPopup('CM-offer-con-advlimit.aspx?ConditionID=" & row.Item("conditionid") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.advlimit", LanguageID) & "</a>")
                  End If
                %>
              </td>
              <td>
                <%If (row.Item("ConditionTypeID") = 1) Then
                    Send(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("CGName"), "&nbsp;"), 25) & "<br />")
                    If (MyCommon.NZ(row.Item("EXName"), "") <> "") Then Send("&nbsp;<i>" & Copient.PhraseLib.Lookup("term.excluding", LanguageID) & "</i> " & MyCommon.SplitNonSpacedString(row.Item("EXName"), 25))
                  ElseIf (row.Item("ConditionTypeID") = 2) Then
                    If (Int(MyCommon.NZ(row.Item("PGID"), 0)) = 1) Or (Int(MyCommon.NZ(row.Item("PGID"), 0)) = 2) Then
                      Send(MyCommon.NZ(row.Item("PGName"), "&nbsp;") & "<br />")
                    ElseIf (Int(MyCommon.NZ(row.Item("PGID"), 0)) = 0) Then
                      Send(Copient.PhraseLib.Lookup("term.entiretransaction", LanguageID) & "<br />")
                    Else
                      Send("<a href=""pgroup-edit.aspx?ProductGroupID=" & row.Item("PGID") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("PGName"), "&nbsp;"), 25) & "</a><br />")
                    End If
                    If (MyCommon.NZ(row.Item("PGEName"), "") <> "") Then
                      If (Int(MyCommon.NZ(row.Item("PGEID"), 0)) <= 1) Then
                        Send("&nbsp;<i>" & Copient.PhraseLib.Lookup("term.excluding", LanguageID) & "</i> " & row.Item("PGEName"))
                      Else
                        Send("&nbsp;<i>" & Copient.PhraseLib.Lookup("term.excluding", LanguageID) & "</i> <a href=""pgroup-edit.aspx?ProductGroupID=" & row.Item("PGEID") & """>" & MyCommon.SplitNonSpacedString(row.Item("PGEName"), 25) & "</a>")
                      End If
                    End If
                  ElseIf (row.Item("ConditionTypeID") = 3) Then
                    If (Int(MyCommon.NZ(row.Item("ProgramID"), 0)) < 1) Then
                      Send("&nbsp;")
                    Else
                      Send("<a href=""point-edit.aspx?ProgramGroupID=" & row.Item("ProgramID") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("ProgramName"), "&nbsp;"), 25) & "</a><br />")
                    End If
                  ElseIf (row.Item("ConditionTypeID") = 6) Then
                    If (Int(MyCommon.NZ(row.Item("SVProgramID"), 0)) < 1) Then
                      Send("&nbsp;")
                    Else
                      Send("<a href=""sv-edit.aspx?ProgramGroupID=" & row.Item("ProgramID") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("SVProgramName"), "&nbsp;"), 25) & "</a><br />")
                    End If
                  ElseIf (row.Item("ConditionTypeID") = 7) Then
                    If (Int(MyCommon.NZ(row.Item("LimitID"), 0)) < 1) Then
                      Send("&nbsp;")
                    Else
                      Send("<a href=""CM-advlimit-edit.aspx?LimitID=" & row.Item("LimitID") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("AlName"), "&nbsp;"), 25) & "</a><br />")
                    End If
                  ElseIf (row.Item("ConditionTypeID") = 4) Then
                    MyCommon.QueryStr = "select CTT.TenderTypeID,Description from ConditionTenderTypes as CTT with (NoLock) left join TenderTypes as TT with (NoLock) on TT.TenderTypeID=CTT.TenderTypeID where ConditionID=" & row.Item("ConditionID") & " order by Description"
                    rst2 = MyCommon.LRT_Select
                    i = 0
                    For Each row2 In rst2.Rows
                      i = i + 1
                      Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row2.Item("Description"), ""), 25))
                      If i = (rst2.Rows.Count - 1) Then
                        Send(" " & StrConv(Copient.PhraseLib.Lookup("term.or", LanguageID), VbStrConv.Lowercase) & " ")
                      ElseIf i < rst2.Rows.Count Then
                        Send(", ")
                      Else
                        Send("")
                      End If
                    Next
                  End If
                %>
              </td>
              <td>
                <%
                  If (MyCommon.NZ(row.Item("GrantPhraseID"), 0) = 0) Then
                    Send(row.Item("GrantDescription"))
                  Else
                    Send(Copient.PhraseLib.Lookup(row.Item("GrantPhraseID"), LanguageID))
                  End If
                %>
              </td>
              <%  
                Dim colSpanAmt As Integer
                If (NumTiers = 0) Then
                  colSpanAmt = 1
                Else
                  colSpanAmt = NumTiers
                End If
                If ((row.Item("ConditionTypeID") = 1 Or row.Item("ConditionTypeID") = 2) Or row.Item("ConditionTypeID") = 3 Or row.Item("ConditionTypeID") = 4 Or row.Item("ConditionTypeID") = 6 Or row.Item("ConditionTypeID") = 7) Then
                  If row.Item("ConditionTypeID") = 6 Then
                    bNeedToFormat = False
                    If intNumDecimalPlaces > 0 Then
                      MyCommon.QueryStr = "Select SVTypeID from StoredValuePrograms with (NoLock) where SVProgramID=" & row.Item("LinkID") & ";"
                      rst2 = MyCommon.LRT_Select
                      If (rst2.Rows.Count > 0) Then
                        If Int(MyCommon.NZ(rst2.Rows(0).Item("SVTypeID"), 0)) = 1 Then
                          bNeedToFormat = True
                        End If
                      End If
                    End If
                  End If

                  MyCommon.QueryStr = "Select AmtRequired from ConditionTiers where ConditionID=" & row.Item("ConditionID")
                  rst2 = MyCommon.LRT_Select
                  For Each row2 In rst2.Rows
                    If (MyCommon.NZ(row.Item("ConditionTypeID"), 0) = 4) Then
                      Send("<td>$" & row2.Item("AmtRequired") & "</td>")
                    ElseIf (MyCommon.NZ(row.Item("ConditionTypeID"), 0) = 3) Then
                      Send("<td>" & Int(row2.Item("AmtRequired")) & "</td>")
                    ElseIf (MyCommon.NZ(row.Item("ConditionTypeID"), 0) = 6) Then
                      If bNeedToFormat Then
                        decTemp = (Int(MyCommon.NZ(row2.Item("AmtRequired"), 0)) * 1.0) / decFactor
                        sTemp = FormatNumber(decTemp, intNumDecimalPlaces)
                        Send("<td>" & sTemp & "</td>")
                      Else
                        Send("<td>" & Int(row2.Item("AmtRequired")) & "</td>")
                      End If
                    ElseIf (MyCommon.NZ(row.Item("ConditionTypeID"), 0) = 7) Then
                      Send("<td>" & Int(row2.Item("AmtRequired")) & "</td>")
                    ElseIf (MyCommon.NZ(row.Item("ConditionTypeID"), 0) = 2) Then
                      Sendb("<td>")
                      If (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 2) Or (MyCommon.NZ(row.Item("QtyUnitType"), 0) >= 6) Then
                        Sendb("$")
                      End If
                      If (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 1) Or (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 5) Then
                        Sendb(Int(row2.Item("AmtRequired")))
                      Else
                        Sendb(row2.Item("AmtRequired"))
                      End If
                      If (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 3) Then
                        Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.pounds", LanguageID), VbStrConv.Lowercase))
                      ElseIf (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 4) Then
                        Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.gallons", LanguageID), VbStrConv.Lowercase))
                      End If
                      Send("</td>")
                    Else
                      Send("<td>" & row2.Item("AmtRequired") & "</td>")
                    End If
                  Next
                Else
                  Send("<td colspan=""" & colSpanAmt & """>&nbsp;</td>")
                End If
              %>
              <%--
                <td colspan="<% if(numtiers=0)then Sendb("1") else Sendb(NumTiers) %>">
                  <%
                    If (MyCommon.NZ(row.Item("AmtRequired"), 0) <> 0) Then
                      If (MyCommon.NZ(row.Item("ConditionTypeID"), 0) = 2 Or MyCommon.NZ(row.Item("ConditionTypeID"), 0) = 4) Then
                        Send("$")
                      End If
                      Send(row.Item("AmtRequired"))
                    End If
                  %>
                </td>
              --%>
              <% If (isTemplate) Then%>
              <td class="templine">
                <input type="checkbox" id="chkLocked3" name="chkLocked" value="<%Sendb(row.Item("conditionid"))%>"
                  <%Sendb(IIf(MyCommon.NZ(row.Item("Disallow_Edit"), False)=True, " checked=""checked""", "")) %>
                  onclick="javascript:updateLocked('lock<%Sendb(row.Item("conditionid"))%>', this.checked);" />
                <input type="hidden" id="con<%Sendb(row.Item("conditionid"))%>" name="con" value="<%Sendb(row.Item("conditionid"))%>" />
                <input type="hidden" id="lock<%Sendb(row.Item("conditionid"))%>" name="locked" value="<%Sendb(IIf(MyCommon.NZ(row.Item("Disallow_Edit"), False)=True, "1", "0"))%>" />
              </td>
              <%ElseIf (FromTemplate) Then%>
              <td class="templine">
                <%Sendb(IIf(MyCommon.NZ(row.Item("Disallow_Edit"), False) = True, "Yes", "No"))%>
              </td>
              <% End If%>
            </tr>
            <%  
            Next
          End If
            %>
          </tbody>
        </table>
        <br />
        <hr class="hidden" />
      </div>
      <div class="box" id="newcondition">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("offer-con.addcondition", LanguageID))%>
          </span>
        </h2>
        <%
          If isTemplate Then
            Send("<span class=""temp"">")
            Send("  <input type=""checkbox"" class=""tempcheck"" id=""temp-conditions"" name=""Disallow_Conditions""" & IIf(Disallow_Conditions, " checked=""checked""", "") & " />")
            Send("  <label for=""temp-conditions"">" & Copient.PhraseLib.Lookup("term.locked", LanguageID) & "</label>")
            Send("</span>")
          End If
          MyCommon.QueryStr = "SELECT PECT.EngineID, PECT.ComponentTypeID, CT.ConditionTypeID, CT.Description, CT.PhraseID, PECT.Singular, PECT.Tierable," & _
                              "  CASE ConditionTypeID " & _
                              "    WHEN 1 THEN (SELECT COUNT(*) FROM OfferConditions WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and ConditionTypeID=1) " & _
                              "    WHEN 2 THEN (SELECT COUNT(*) FROM OfferConditions WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and ConditionTypeID=2) " & _
                              "    WHEN 3 THEN (SELECT COUNT(*) FROM OfferConditions WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and ConditionTypeID=3) " & _
                              "    WHEN 4 THEN (SELECT COUNT(*) FROM OfferConditions WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and ConditionTypeID=4) " & _
                              "    WHEN 5 THEN (SELECT COUNT(*) FROM OfferConditions WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and ConditionTypeID=5) " & _
                              "    WHEN 6 THEN (SELECT COUNT(*) FROM OfferConditions WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and ConditionTypeID=6) " & _
                              "    WHEN 7 THEN (SELECT COUNT(*) FROM OfferConditions WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and ConditionTypeID=7) " & _
                              "    WHEN 100 THEN (SELECT COUNT(*) FROM OfferConditions WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and ConditionTypeID=100) " & _
                              "    ELSE 0 " & _
                              "  END as ItemCount " & _
                              "FROM PromoEngineComponentTypes AS PECT " & _
                              "INNER JOIN ConditionTypes AS CT ON CT.ConditionTypeID=PECT.LinkID " & _
                              "WHERE EngineID=0 AND PECT.ComponentTypeID=1 AND PECT.Enabled=1"
          
          ' preference conditions do not work with tiered offers because evaluation of preference condition is performed on central before the engine can evaluate the tier.
          ' Also - do not include preference conditions if EPM is not installed
          If ((NumTiers > 0) Or (PrefManInstalled = False)) Then
            MyCommon.QueryStr &= " and CT.ConditionTypeID <> 100"
          End If
          rst = MyCommon.LRT_Select
          If rst.Rows.Count > 0 Then
            Send("<label for=""newconglobal"">" & Copient.PhraseLib.Lookup("offer-con.addglobal", LanguageID) & ":</label><br />")
            Send("<select class=""medium"" id=""newconglobal"" name=""newconglobal""" & IIf(FromTemplate And Disallow_Conditions, " disabled=""disabled""", "") & ">")
            For Each row In rst.Rows
              If (row.Item("Singular") = False) OrElse (row.Item("Singular") = True AndAlso row.Item("ItemCount") = 0) Then
                If Not (bStoreUser AND (row.Item("ConditionTypeID") = "3" OR row.Item("ConditionTypeID") = "6")) Then
                Send("<option value=""" & row.Item("ConditionTypeID") & """>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "</option>")
                End If
              End If
            Next
            Send("</select>")
            If Not (IsTemplate)
               Send("<input type=""submit"" class=""regular"" id=""addGlobal"" name=""addGlobal""" & IIf(Not Logix.UserRoles.EditOffer, " disabled=""disabled""", "") & IIf(FromTemplate And Disallow_Conditions, " disabled=""disabled""", "") & " value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """ />")
              Send("<br />")
            Else
               Send("<input type=""submit"" class=""regular"" id=""addGlobal"" name=""addGlobal""" & IIf(Not Logix.UserRoles.EditTemplates, " disabled=""disabled""", "") & " value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """ />")
              Send("<br />")
            End If
          End If
          
          If (NumTiers > 0) Then
            MyCommon.QueryStr = "SELECT PECT.EngineID, PECT.ComponentTypeID, CT.ConditionTypeID, CT.Description, CT.PhraseID, PECT.Singular, PECT.Tierable," & _
                                "  CASE ConditionTypeID " & _
                                "    WHEN 1 THEN (SELECT COUNT(*) FROM OfferConditions WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and ConditionTypeID=1) " & _
                                "    WHEN 2 THEN (SELECT COUNT(*) FROM OfferConditions WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and ConditionTypeID=2) " & _
                                "    WHEN 3 THEN (SELECT COUNT(*) FROM OfferConditions WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and ConditionTypeID=3) " & _
                                "    WHEN 4 THEN (SELECT COUNT(*) FROM OfferConditions WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and ConditionTypeID=4) " & _
                                "    WHEN 5 THEN (SELECT COUNT(*) FROM OfferConditions WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and ConditionTypeID=5) " & _
                                "    WHEN 6 THEN (SELECT COUNT(*) FROM OfferConditions WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and ConditionTypeID=6) " & _
                                "    WHEN 7 THEN (SELECT COUNT(*) FROM OfferConditions WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and ConditionTypeID=7) " & _
                                "    ELSE 0 " & _
                                "  END as ItemCount " & _
                                "FROM PromoEngineComponentTypes AS PECT " & _
                                "INNER JOIN ConditionTypes AS CT ON CT.ConditionTypeID=PECT.LinkID " & _
                                "WHERE EngineID=0 AND PECT.ComponentTypeID=1 AND PECT.Enabled=1 and PECT.Tierable=1;"
            rst = MyCommon.LRT_Select
            If rst.Rows.Count > 0 Then
              Send("<label for=""newcontiered"">" & Copient.PhraseLib.Lookup("offer-con.addtiered", LanguageID) & ":</label><br />")
              Send("<select class=""medium"" id=""newcontiered"" name=""newcontiered""" & IIf(FromTemplate And Disallow_Conditions, " disabled=""disabled""", "") & ">")
              For Each row In rst.Rows
                If (row.Item("Singular") = False) OrElse (row.Item("Singular") = True AndAlso row.Item("ItemCount") = 0) Then
                  Send("<option value=""" & row.Item("ConditionTypeID") & """>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "</option>")
                End If
              Next
              Send("</select>")
              If Not (IsTemplate) Then
                Send("<input type=""submit"" class=""regular"" id=""addTier"" name=""addTier""" & IIf(Not Logix.UserRoles.EditOffer, " disabled=""disabled""", "") & IIf(FromTemplate And Disallow_Conditions, " disabled=""disabled""", "") & " value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """ />")
                Send("<br />")
              Else
                Send("<input type=""submit"" class=""regular"" id=""addTier"" name=""addTier""" & IIf(Not Logix.UserRoles.EditTemplates, " disabled=""disabled""", "") & " value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """ />")
                Send("<br />")
              End If
            End If
          End If
        %>
      </div>
     </form>
    </div>
    <br clear="all" />
  </div>

<script runat="server">
  Const ANNIVERSARY_DATE_OP As Integer = 2

    Protected Sub Page_Load(ByVal obj As Object, ByVal e As EventArgs)
    Dim MyCommon As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    MyCommon.AppName = "CPEoffer-con.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
   
    
    Dim uc As logix_UserControls_OfferEligibilityConditions = Page.FindControl("ucOfferEligibilityCondition")
    Dim OfferID As Long
    Dim dt As New DataTable
        uc.Disable = Logix.UserRoles.EditOffer And (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
    OfferID = Request.QueryString("OfferID")
    uc.OfferID = OfferID
   
    uc.LanguageID = LanguageID
    MyCommon.QueryStr = " SELECT LinkID as CustomerGroupID, Disallow_Edit, RequiredFromTemplate, IsOptInGroup " & _
                        "FROM OfferConditions O WITH (NOLOCK) inner join CustomerGroups CG on O.LinkID=CG.CustomerGroupID " & _
                        "WHERE OfferID = @OfferID AND O.Deleted=0 and CG.Deleted=0 and ConditionTypeID=1 "
    MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
    dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
    If dt.IsNotEmpty() Then
      If Convert.ToInt64(dt.Rows(0)("CustomerGroupID")) > 0 andalso Convert.ToBoolean(dt.Rows(0)("IsOptInGroup")) =False  Then
        uc.IsOptInDisabled = True
   
      End If
    End If
    
    uc.AdminUserID = AdminUserID
    If (Request.QueryString("Save") <> "") Then
      Dim isOptInPanelLocked As Integer = 0
      If (Request.QueryString("IsOptInPanelLocked") = "1") Then
        isOptInPanelLocked = 1
      End If
 
      MyCommon.QueryStr = "update TemplatePermissions with (RowLock) set Disallow_OptIn =@isOptInPanelLocked where OfferID=@OfferID"
      MyCommon.DBParameters.Add("isOptInPanelLocked", SqlDbType.Bit).Value = isOptInPanelLocked
      MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
      MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
      
      If Request.QueryString("IsTemplate") = "IsTemplate" Then
        Dim Condid() As String = Request.QueryString.GetValues("EligibilityCondID")
        Dim CondVal() As String = Request.QueryString.GetValues("EligibilityCondVal")
        
        Dim sQuery As String = ""
        
        
        If (Not Condid Is Nothing AndAlso Not CondVal Is Nothing AndAlso Condid.Length = CondVal.Length) Then
          For LoopCtr As Integer = 0 To Condid.Count - 1
            
            If CondVal(LoopCtr) = "1" Then
              MyCommon.QueryStr = "update conditions with (RowLock) set DisallowEdit=@DisallowEdit, " & _
                            "RequiredFromTemplate=@RequiredFromTemplate " & _
                            "where ConditionID=@ConditionID;"
              MyCommon.DBParameters.Add("@RequiredFromTemplate", SqlDbType.Bit).Value = False
            Else
              MyCommon.QueryStr = "update conditions with (RowLock) set DisallowEdit=@DisallowEdit " & _
                          "where ConditionID=@ConditionID;"
            End If
            MyCommon.DBParameters.Add("@DisallowEdit", SqlDbType.Bit).Value = IIf(CondVal(LoopCtr) = "1", True, False)
            MyCommon.DBParameters.Add("@ConditionID", SqlDbType.BigInt).Value = Condid(LoopCtr)
            MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
          Next
      
        
        End If
      
      End If
    End If
    MyCommon.QueryStr = "SELECT Disallow_Optin FROM TemplatePermissions WHERE OfferID = @OfferID"
    MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
    dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
    If dt.Rows.Count > 0 Then
      For Each row As DataRow In dt.Rows
        uc.IsOptInBlockLocked = MyCommon.NZ(row("Disallow_Optin"), False)
      Next
    End If
  End Sub
  Sub Send_Preference_Details(ByRef Common As Copient.CommonInc, ByVal PreferenceID As Long)
    Dim dt As DataTable
    Dim IntegrationVals As New Copient.CommonInc.IntegrationValues
    Dim PrefPageName As String = ""
    Dim Tokens As String = ""
    Dim RootURI As String = ""
    
    Common.QueryStr = "select UserCreated, Name as PrefName " & _
                      "from Preferences as PREF with (NoLock) " & _
                      "where PREF.PreferenceID=" & PreferenceID & " and PREF.Deleted=0;"
    dt = Common.PMRT_Select
    If dt.Rows.Count > 0 Then
      If (Common.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER, IntegrationVals)) Then
        PrefPageName = IIf(Common.NZ(dt.Rows(0).Item("UserCreated"), False), "prefscustom-edit.aspx", "prefsstd-edit.aspx")
          
        RootURI = IntegrationVals.HTTP_RootURI
        If RootURI IsNot Nothing AndAlso RootURI.Length > 0 AndAlso Right(RootURI, 1) <> "/" Then
          RootURI &= "/"
        End If
        
        Tokens = "SendToURI="
        Sendb("  <a href=""authtransfer.aspx?SendToURI=" & RootURI & "UI/" & PrefPageName & "?prefid=" & PreferenceID & """>")
        Send(Common.NZ(dt.Rows(0).Item("PrefName"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)) & "</a>")
      End If
    End If
  End Sub

  Sub Send_Preference_Info(ByRef Common As Copient.CommonInc, ByVal ConditionID As Integer)
    Dim dt As DataTable
    Dim PreferenceID As Long = 0
    Dim ComboText As String = ""
    Dim i As Integer = 0
    Dim CellCount As Integer = 0
    Dim ValueSent As Boolean = False
    Dim AndComboed As Boolean = True
    
    ' find all the tier values
    Common.QueryStr = "select CPV.PKID, CPV.PreferenceID, CPV.Value, CPV.ValueComboTypeID, CPV.DateOperatorTypeID, " & _
                      "  case when POT.PhraseID is null then POT.Description" & _
                      "  else Convert(nvarchar(200), PT.Phrase) end as OperatorText " & _
                      "from CM_ConditionPreferenceValues as CPV with (NoLock) " & _
                      "inner join CPE_PrefOperatorTypes as POT with (NoLock) on POT.PrefOperatorTypeID = CPV.OperatorTypeID " & _
                      "left join PhraseText as PT with (NoLock) on PT.PhraseID = POT.PhraseID and LanguageID=" & LanguageID & " " & _
                      "where CPV.ConditionID=" & ConditionID
    dt = Common.LRT_Select
    For i = 0 To dt.Rows.Count - 1
      AndComboed = (Common.NZ(dt.Rows(i).Item("ValueComboTypeID"), 2) = 1)
      PreferenceID = Common.NZ(dt.Rows(i).Item("PreferenceID"), 0)
      
      If ValueSent Then Send(" " & Copient.PhraseLib.Lookup(IIf(AndComboed, "term.and", "term.or"), LanguageID) & " ")

      If Common.NZ(dt.Rows(i).Item("DateOperatorTypeID"), 0) > 0 Then
        Send(Get_Date_Display_Text(Common, dt.Rows(i).Item("PKID")))
      Else
        Send(Common.NZ(dt.Rows(i).Item("OperatorText"), "") & " " & Get_Preference_Value(Common, PreferenceID, Common.NZ(dt.Rows(i).Item("Value"), "")))
      End If

      If i < dt.Rows.Count - 1 Then
        Send(" <i>" & ComboText.ToLower & "</i> ")
      End If

      ValueSent = True
    Next
  End Sub
  
  Function Get_Preference_Value(ByRef Common As Copient.CommonInc, ByVal PreferenceID As Long, ByVal Value As String) As String
    Dim TempLong As Long = 0
    Dim dt As DataTable
    
    Common.QueryStr = "select DataTypeID from Preferences with (NoLock) where PreferenceID=" & PreferenceID & " and Deleted=0;"
    dt = Common.PMRT_Select
    If dt.Rows.Count > 0 Then
      Select Case Common.NZ(dt.Rows(0).Item("DataTypeID"), 0)
        Case 1 ' list
          ' lookup to see if this is a preference with list items, if so get the list item name
          Common.QueryStr = "select case when UPT.PhraseID is null then PLI.Name " & _
                            "       else CONVERT(nvarchar(200), UPT.Phrase) end as PhraseText " & _
                            "from Preferences as PREF with (NoLock) " & _
                            "inner join PreferenceListItems as PLI with (NoLock) on PLI.PreferenceID = PREF.PreferenceID " & _
                            "left join UserPhraseText as UPT with (NoLock) on UPT.PhraseID = PLI.NamePhraseID " & _
                            "where PREF.Deleted=0 and PREF.DataTypeID=1 and PREF.PreferenceID=" & PreferenceID & _
                            "  and PLI.Value=N'" & Value & "';"
          dt = Common.PMRT_Select
          If dt.Rows.Count > 0 Then
            Value = Common.NZ(dt.Rows(0).Item("PhraseText"), Value)
          End If
        Case 5 ' boolean
          Value = Copient.PhraseLib.Lookup(IIf(Value = "1", "term.true", "term.false"), LanguageID)
      End Select
      
    End If

    Return Value
  End Function
  
  Function Get_Date_Display_Text(ByRef Common As Copient.CommonInc, ByVal ValuePKID As Integer) As String
    Dim DisplayText As String = ""
    Dim dt As DataTable
    Dim ValueModifier As String = ""
    Dim Offset, DaysBefore, DaysAfter As Integer
    
    Common.QueryStr = "select CPV.Value, CPV.ValueModifier, CPV.ValueTypeID, POT.PhraseID as OperatorPhraseID, CPV.DaysBefore, CPV.DaysAfter, " & _
                      "CPV.DateOperatorTypeID, PDOT.PhraseID as DateOpPhraseID " & _
                      "from CM_ConditionPreferenceValues as CPV with (NoLock) " & _
                      "inner join CPE_PrefOperatorTypes as POT with (NoLock) on POT.PrefOperatorTypeID = CPV.OperatorTypeID " & _
                      "inner join CPE_PrefDateOperatorTypes as PDOT with (NoLock) on PDOT.PrefDateOperatorTypeID = CPV.DateOperatorTypeID " & _
                      "where PKID=" & ValuePKID & ";"
    dt = Common.LRT_Select
    If dt.Rows.Count > 0 Then
      DisplayText = Copient.PhraseLib.Lookup(Common.NZ(dt.Rows(0).Item("DateOpPhraseID"), ""), LanguageID) & " "
      DisplayText &= Copient.PhraseLib.Lookup(Common.NZ(dt.Rows(0).Item("OperatorPhraseID"), ""), LanguageID) & " "
      If Common.NZ(dt.Rows(0).Item("ValueTypeID"), 0) = 1 Then
        DisplayText &= "[" & Copient.PhraseLib.Lookup("term.currentdate", LanguageID).ToLower & "]"
        ValueModifier = Common.NZ(dt.Rows(0).Item("ValueModifier"), "")
        If ValueModifier <> "" AndAlso Integer.TryParse(ValueModifier, Offset) Then
          ValueModifier = " " & IIf(Offset < 0, " - ", " + ") & Math.Abs(Offset)
        End If
        DisplayText &= ValueModifier
      Else
        DisplayText &= " " & Common.NZ(dt.Rows(0).Item("Value"), "")
      End If
      
      If Common.NZ(dt.Rows(0).Item("DateOperatorTypeID"), 0) = ANNIVERSARY_DATE_OP Then
        DaysBefore = Common.NZ(dt.Rows(0).Item("DaysBefore"), 0)
        DaysAfter = Common.NZ(dt.Rows(0).Item("DaysAfter"), 0)

        If DaysBefore > 0 AndAlso DaysAfter > 0 Then
          DisplayText &= " (-" & DaysBefore & " / +" & DaysAfter & ")"
        ElseIf DaysBefore > 0 AndAlso DaysAfter = 0 Then
          DisplayText &= " (-" & DaysBefore & ")"
        ElseIf DaysBefore = 0 AndAlso DaysAfter > 0 Then
          DisplayText &= " (+" & DaysAfter & ")"
        End If
      End If
    End If
    
    Return DisplayText
  End Function

</script>
<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (OfferID > 0 And Logix.UserRoles.AccessNotes) Then
      Send_Notes(3, OfferID, AdminUserID)
    End If
  End If
done:
  MyCommon.Close_LogixRT()
  If PrefManInstalled Then MyCommon.Close_PrefManRT()
  Send_BodyEnd()
  MyCommon = Nothing
  Logix = Nothing
%>
