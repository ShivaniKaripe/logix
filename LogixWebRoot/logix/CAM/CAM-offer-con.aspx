<%@ Page Language="vb" Debug="true" CodeFile="/logix/LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: CAM-offer-con.aspx 
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
  Dim dt2, dt3 As DataTable
  Dim OfferID As Long
  Dim ConditionID As Long
  Dim Name As String = ""
  Dim isTemplate As Boolean = False
  Dim FromTemplate As Boolean = False
  Dim Disallow_Conditions As Boolean = False
  Dim IsTemplateVal As String = "Not"
  Dim ActiveSubTab As Integer = 205
  Dim roid As Integer
  Dim i As Integer
  Dim Days As String = ""
  Dim Times As String = ""
  Dim isCustomer As Boolean = False
  Dim isProduct As Boolean = False
  Dim isProductDisqualifier As Boolean = False
  Dim isPoint As Boolean = False
  Dim isDay As Boolean = False
  Dim isTime As Boolean = False
  Dim isStoredValue As Boolean = False
  Dim isTender As Boolean = False
  Dim isInstantWin As Boolean = False
  Dim isPLU As Boolean = False
  Dim DeleteBtnDisabled As String = ""
  Dim AccumEligible As Boolean = False
  Dim infoMessage As String = ""
  Dim modMessage As String = ""
  Dim Handheld As Boolean = False
  Dim DaysLocked As Boolean = False
  Dim TimeLocked As Boolean = False
  Dim CondTypes As String() = Nothing
  Dim Conditions As String() = Nothing
  Dim LockedStatus As String() = Nothing
  Dim LoopCtr As Integer = 0
  Dim sQuery As String = ""
  Dim BannersEnabled As Boolean = True
  Dim StatusCode As Copient.LogixInc.STATUS_FLAGS = Copient.LogixInc.STATUS_FLAGS.STATUS_UNKNOWN
  Dim StatusText As String = ""
  Dim TenderList As String = ""
  Dim TenderValue As String = ""
  Dim TenderDisallowEdit As Boolean
  Dim TenderRequired As Boolean
  Dim TenderExcluded As Boolean
  Dim TenderExcludedAmt As Object
  Dim StatusFlag As Integer
  Dim TierLevels As Integer = 1
  Dim t As Integer = 1
  Dim ProdID As Integer = 0
  Dim IncentiveID As Integer = 0
  Dim ExcludedIncentiveID As Integer = 0
  Dim ProductConditions As Integer = 0
  Dim ProductCombo As Integer = 0
  Dim ExcludedProductGroupID As Integer = 0
  Dim ExcludedProductGroupName As String = ""
  Dim AccumEnabled As Boolean = False
  Dim IncentiveTenderID As Integer = 0
  Dim IsFooterOffer As Boolean = False
  Dim TenderWorthy As Boolean = False

  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "CAM-offer-con.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  
  OfferID = Request.QueryString("OfferID")
  ConditionID = Request.QueryString("ConditionID")
  
  If (OfferID = 0) Then
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "/logix/CAM/CAM-offer-gen.aspx")
  End If
  
  MyCommon.QueryStr = "select RewardOptionID, TierLevels, ProductComboID from CPE_RewardOptions with (NoLock) " & _
                      "where IncentiveID=" & OfferID & " and TouchResponse=0 and Deleted=0;"
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    roid = rst.Rows(0).Item("RewardOptionID")
    TierLevels = rst.Rows(0).Item("TierLevels")
    ProductCombo = rst.Rows(0).Item("ProductComboID")
  End If

  IsFooterOffer = MyCam.IsFooterOffer(OfferID)

  MyCommon.QueryStr = "select AccumMin from CPE_IncentiveProductGroups where RewardOptionID=" & roid & " and deleted=0"
  rst = MyCommon.LRT_Select()
  If rst.Rows.Count > 0 Then
    If rst.Rows.Count = 1 And MyCommon.NZ(rst.Rows(0).Item("AccumMin"), 0) > 0 Then
      AccumEnabled = True
    End If
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
      Send("<script type=""text/javascript"">openPopup('/logix/CPEoffer-con-customer.aspx?OfferID=" & OfferID & "')</script>")
    ElseIf (Request.QueryString("newconglobal") = 2) Then
      Send("<script type=""text/javascript"">openPopup('/logix/CPEoffer-con-product.aspx?OfferID=" & OfferID & "&Disqualifier=0')</script>")
    ElseIf (Request.QueryString("newconglobal") = 3) Then
      Send("<script type=""text/javascript"">openPopup('/logix/CPEoffer-con-point.aspx?OfferID=" & OfferID & "')</script>")
    ElseIf (Request.QueryString("newconglobal") = 5) Then
      Send("<script type=""text/javascript"">openPopup('/logix/CPEoffer-con-tender.aspx?OfferID=" & OfferID & "')</script>")
    ElseIf (Request.QueryString("newconglobal") = 6) Then
      Send("<script type=""text/javascript"">openPopup('/logix/CPEoffer-con-day.aspx?OfferID=" & OfferID & "')</script>")
    ElseIf (Request.QueryString("newconglobal") = 7) Then
      Send("<script type=""text/javascript"">openPopup('/logix/CPEoffer-con-time.aspx?OfferID=" & OfferID & "')</script>")
      'ElseIf (Request.QueryString("newconglobal") = 8) Then
      '  Send("<script type=""text/javascript"">openPopup('/logix/CPEoffer-con-instantwin.aspx?OfferID=" & OfferID & "')</script>")
    ElseIf (Request.QueryString("newconglobal") = 9) Then
      Send("<script type=""text/javascript"">openPopup('/logix/CPEoffer-con-plu.aspx?OfferID=" & OfferID & "')</script>")
    ElseIf (Request.QueryString("newconglobal") = 10) Then
      Send("<script type=""text/javascript"">openPopup('/logix/CPEoffer-con-product.aspx?OfferID=" & OfferID & "&Disqualifier=1')</script>")
      'ElseIf (Request.QueryString("newconglobal") = 11) Then
      '  Send("<script type=""text/javascript"">openPopup('/logix/CPEoffer-con-plu.aspx?OfferID=" & OfferID & "')</script>")
    End If
  ElseIf (Request.QueryString("mode") = "ChangeProductCombo") Then
    If (Request.QueryString("pc") <> "") Then
      ProductCombo = MyCommon.Extract_Val(Request.QueryString("pc"))
      'Set the ProductCombo for this offer
      If ProductCombo = 1 Then
        ProductCombo = 2
        MyCommon.QueryStr = "Update CPE_RewardOptions set ProductComboID=" & ProductCombo & " where RewardOptionID=" & roid
        MyCommon.LRT_Execute()
      ElseIf ProductCombo = 2 Then
        ProductCombo = 1
        MyCommon.QueryStr = "Update CPE_RewardOptions set ProductComboID=" & ProductCombo & " where RewardOptionID=" & roid
        MyCommon.LRT_Execute()
      End If
  End If
  ElseIf (Request.QueryString("mode") = "Delete") Then
    If (Request.QueryString("Option") = "Customer") Then
      ' ok someone clicked the X on the customer group stuff lets ditch all the associated groups on this offer
      MyCommon.QueryStr = "update CPE_IncentiveCustomerGroups with (RowLock) set deleted=1, LastUpdate=getdate(),TCRMAStatusFlag=3 where RewardOptionID=" & roid & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1, AllowOptOut=0 where IncentiveID=" & OfferID & ";"
      MyCommon.LRT_Execute()
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-customer-delete", LanguageID))
    ElseIf (Request.QueryString("Option") = "Product") Then
      If Request.QueryString("IncentiveProductGroupID") <> "" Then
        IncentiveID = MyCommon.Extract_Val(Request.QueryString("IncentiveProductGroupID"))
      End If
      ' If the requested deletion is for Any Product, see if there's an accompanying excluded group that should be deleted as well
      MyCommon.QueryStr = "select ProductGroupID from CPE_IncentiveProductGroups with (NoLock) where IncentiveProductGroupID=" & IncentiveID & ";"
      rst = MyCommon.LRT_Select
      If rst.Rows.Count > 0 Then
        If MyCommon.NZ(rst.Rows(0).Item("ProductGroupID"), 0) = 1 Then
          MyCommon.QueryStr = "select IncentiveProductGroupID from CPE_IncentiveProductGroups with (NoLock) " & _
                              "where RewardOptionID=" & roid & " and ExcludedProducts=1 and Deleted=0;"
          rst2 = MyCommon.LRT_Select
          If rst2.Rows.Count > 0 Then
            ExcludedIncentiveID = MyCommon.NZ(rst2.Rows(0).Item("IncentiveProductGroupID"), 0)
          End If
        End If
      End If
      MyCommon.QueryStr = "select IncentiveProductGroupID from CPE_IncentiveProductGroups where RewardOptionID=" & roid & " and deleted =0"
      rst = MyCommon.LRT_Select
      ProductConditions = rst.Rows.Count
      If ProductConditions = 2 Then
        'Change the ProductComboID to single
        MyCommon.QueryStr = "Update CPE_RewardOptions set ProductComboID=0 where RewardOptionID=" & roid
        MyCommon.LRT_Execute()
      End If
      ' Someone clicked the X on the product group condition stuff
      MyCommon.QueryStr = "update CPE_IncentiveProductGroups with (RowLock) set Deleted=1, LastUpdate=getdate(),TCRMAStatusFlag=3 " & _
                          "where RewardOptionID=" & roid & " "
      If ExcludedIncentiveID > 0 Then
        MyCommon.QueryStr &= " and IncentiveProductGroupID in (" & IncentiveID & "," & ExcludedIncentiveID & ");"
      Else
        MyCommon.QueryStr &= " and IncentiveProductGroupID=" & IncentiveID & ";"
      End If
      MyCommon.LRT_Execute()
      'Remove the tier records from the product condition being deleted
      MyCommon.QueryStr = "delete from CPE_IncentiveProductGroupTiers where IncentiveProductGroupID not in " & _
                          "(select IncentiveProductGroupID from CPE_IncentiveProductGroups where RewardOptionID=" & roid & " and deleted=0) " & _
                          "and RewardOptionID=" & roid
      MyCommon.LRT_Execute()
      
      'If it's the last product condition then remove any accumulation printed message that may have been created
      If ProductConditions = 1 Then
        ' Check if accumulation message needs to be removed
        MyCommon.QueryStr = "dbo.pa_CPE_AccumMsgEligible"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@ROID", SqlDbType.Int, 4).Value = roid
        MyCommon.LRTsp.Parameters.Add("@AccumEligible", SqlDbType.Bit, 1).Direction = ParameterDirection.Output
        MyCommon.LRTsp.ExecuteNonQuery()
        AccumEligible = MyCommon.LRTsp.Parameters("@AccumEligible").Value
        MyCommon.Close_LRTsp()
      
        If Not (AccumEligible) Then
          ' Mark any accumulation messages as deleted
          MyCommon.QueryStr = "update CPE_Deliverables with (RowLock) set Deleted = 1 where DeliverableID in " & _
                              "(select D.DeliverableID from CPE_RewardOptions RO with (NoLock) inner join CPE_Deliverables D with (NoLock) on RO.RewardOptionID = D.RewardOptionID " & _
                              "where RO.Deleted = 0 and D.Deleted = 0 and RO.IncentiveID = " & OfferID & " and RewardOptionPhase = 2 and DeliverableTypeID = 4);"
          MyCommon.LRT_Execute()
        End If     
     
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-product-delete", LanguageID))
      
        ' Since there is no product condition assigned, then remove printed messages and cashier messages
        ' notifications for this incentive
        MyCommon.QueryStr = "dbo.pa_CPE_RemoveNotifications"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.Int, 4).Value = OfferID
        MyCommon.LRTsp.Parameters.Add("@ParentROID", SqlDbType.Int, 4).Value = roid
        MyCommon.LRTsp.ExecuteNonQuery()
        MyCommon.Close_LRTsp()
      
        'now remove the graphics and clean up the touchpoints
        MyCommon.QueryStr = "select DeliverableID from CPE_Deliverables with (NoLock) where RewardOptionID= " & roid & " and deleted = 0 " & _
                            "and DeliverableTypeID=1 and RewardOptionPhase=1"
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
          For Each row In rst.Rows
            RemoveGraphic(OfferID, MyCommon.NZ(row.Item("DeliverableID"), 0))
          Next
        End If
      End If
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID
      MyCommon.LRT_Execute()
    ElseIf (Request.QueryString("Option") = "ProductDisqualifier") Then
      ' Someone clicked the X on the product disqualifier stuff
      MyCommon.QueryStr = "update CPE_IncentiveProductGroups with (RowLock) set Deleted=1, LastUpdate=getdate(),TCRMAStatusFlag=3 " & _
                          "where RewardOptionID=" & roid & " and Disqualifier=1;"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID
      MyCommon.LRT_Execute()
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-product-delete", LanguageID))
    ElseIf (Request.QueryString("Option") = "Point") Then
      ' ok someone clicked the X on the customer group stuff lets ditch all the associated groups on this offer
      MyCommon.QueryStr = "delete from CPE_IncentivePointsGroups with (RowLock) where RewardOptionID=" & roid & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "delete from CPE_IncentivePointsGroupTiers with (RowLock) where RewardOptionID=" & roid & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
      MyCommon.LRT_Execute()
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-point-delete", LanguageID))
    ElseIf (Request.QueryString("Option") = "StoredValue") Then
      ' ok someone clicked the X on the stored value lets ditch all the associated groups on this offer
      MyCommon.QueryStr = "delete from CPE_IncentiveStoredValuePrograms with (RowLock) where RewardOptionID=" & roid & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "delete from CPE_IncentiveStoredValueProgramTiers with (RowLock) where RewardOptionID=" & roid & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
      MyCommon.LRT_Execute()
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-sv-delete", LanguageID))
    ElseIf (Request.QueryString("Option") = "Tender") Then
      ' Someone clicked the X on the tender stuff lets ditch all the associated groups on this offer
      If Not IncentiveTenderID = 0 Then
        MyCommon.QueryStr = "delete from CPE_IncentiveTenderTypes with (RowLock) where RewardOptionID=" & roid & " and IncentiveTenderID=" & IncentiveTenderID & ";"
        MyCommon.LRT_Execute()
        MyCommon.QueryStr = "delete from CPE_IncentiveTenderTypeTiers with (RowLock) where RewardOptionID=" & roid & " and IncentiveTenderID=" & IncentiveTenderID & ";"
        MyCommon.LRT_Execute()
      End If
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
      MyCommon.LRT_Execute()
      'MyCommon.QueryStr = "update CPE_RewardOptions with (RowLock) set ExcludedTender=0 where IncentiveID=" & OfferID & ";"
      'MyCommon.LRT_Execute()
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-tender-delete", LanguageID))
    ElseIf (Request.QueryString("Option") = "Day") Then
      ' Someone clicked the X on a day condition
      MyCommon.QueryStr = "delete from CPE_IncentiveDOW with (RowLock) where IncentiveID=" & OfferID & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
      MyCommon.LRT_Execute()
      ' reset the EveryDOW column to reflect the change - if all 7 days chosen then set to 1, otherwise set to 0
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set EveryDOW=1 where IncentiveID=" & OfferID & " and Deleted=0;"
      MyCommon.LRT_Execute()
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-day-delete", LanguageID))
    ElseIf (Request.QueryString("Option") = "Time") Then
      ' Someone clicked the X on a time condition
      MyCommon.QueryStr = "delete from CPE_IncentiveTOD with (RowLock) where IncentiveID=" & OfferID & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
      MyCommon.LRT_Execute()
      ' reset the EveryTOD column to reflect the change
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set EveryTOD=1 where IncentiveID=" & OfferID & " and Deleted=0;"
      MyCommon.LRT_Execute()
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-time-delete", LanguageID))
    ElseIf (Request.QueryString("Option") = "InstantWin") Then
      ' Someone clicked the X on an instant win condition
      MyCommon.QueryStr = "delete from CPE_IncentiveInstantWin with (RowLock) where RewardOptionID=" & roid & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
      MyCommon.LRT_Execute()
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-instantwin-delete", LanguageID))
    ElseIf (Request.QueryString("Option") = "PLU") Then
      ' Someone clicked the X on a PLU condition
      If Request.QueryString("IncentivePLUID") <> "" Then
        IncentiveID = MyCommon.Extract_Val(Request.QueryString("IncentivePLUID"))
      End If
      MyCommon.QueryStr = "select IncentivePLUID from CPE_IncentivePLUs where RewardOptionID=" & roid
      dt3 = MyCommon.LRT_Select()
      If dt3.Rows.Count = 1 Then
        MyCommon.QueryStr = "delete from CPE_IncentivePLUs with (RowLock) where RewardOptionID=" & roid & " and IncentivePLUID=" & IncentiveID & ";"
        MyCommon.LRT_Execute()
        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1, MutuallyExclusive=0 where IncentiveID=" & OfferID & ";"
        MyCommon.LRT_Execute()
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-plu-delete", LanguageID))
      Else
        MyCommon.QueryStr = "delete from CPE_IncentivePLUs with (RowLock) where RewardOptionID=" & roid & " and IncentivePLUID=" & IncentiveID & ";"
        MyCommon.LRT_Execute()
        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
        MyCommon.LRT_Execute()
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-plu-delete", LanguageID))
      End If
    End If
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "/logix/CAM/CAM-offer-con.aspx?OfferID=" & OfferID)
    GoTo done
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
                        "where IncentiveCustomerID=" & Conditions(LoopCtr) & ";"
            Case "Product"
              sQuery = "update CPE_IncentiveProductGroups with (RowLock) set DisallowEdit=" & LockedStatus(LoopCtr) & ", " & _
                        "RequiredFromTemplate=" & IIf(LockedStatus(LoopCtr) = "1", "0", "RequiredFromTemplate") & " " & _
                        "where IncentiveProductGroupID=" & Conditions(LoopCtr) & ";"
            Case "Points"
              sQuery = "update CPE_IncentivePointsGroups with (RowLock) set DisallowEdit=" & LockedStatus(LoopCtr) & ", " & _
                        "RequiredFromTemplate=" & IIf(LockedStatus(LoopCtr) = "1", "0", "RequiredFromTemplate") & " " & _
                        "where IncentivePointsID=" & Conditions(LoopCtr) & ";"
            Case "Days"
              sQuery = "update CPE_IncentiveDOW with (RowLock) set DisallowEdit=" & LockedStatus(LoopCtr) & ", " & _
                        "RequiredFromTemplate=" & IIf(LockedStatus(LoopCtr) = "1", "0", "RequiredFromTemplate") & " " & _
                        "where IncentiveID=" & Conditions(LoopCtr) & ";"
            Case "Time"
              sQuery = "update CPE_IncentiveTOD with (RowLock) set DisallowEdit=" & LockedStatus(LoopCtr) & ", " & _
                        "RequiredFromTemplate=" & IIf(LockedStatus(LoopCtr) = "1", "0", "RequiredFromTemplate") & " " & _
                        "where IncentiveID=" & Conditions(LoopCtr) & ";"
            Case "StoredValue"
              sQuery = "update CPE_IncentiveStoredValuePrograms with (RowLock) set DisallowEdit=" & LockedStatus(LoopCtr) & ", " & _
                        "RequiredFromTemplate=" & IIf(LockedStatus(LoopCtr) = "1", "0", "RequiredFromTemplate") & " " & _
                        "where IncentiveStoredValueID=" & Conditions(LoopCtr) & ";"
            Case "Tender"
              sQuery = "update CPE_IncentiveTenderTypes with (RowLock) set DisallowEdit=" & LockedStatus(LoopCtr) & ", " & _
                        "RequiredFromTemplate=" & IIf(LockedStatus(LoopCtr) = "1", "0", "RequiredFromTemplate") & " " & _
                        "where IncentiveTenderID=" & Conditions(LoopCtr) & ";"
            Case "InstantWin"
              sQuery = "update CPE_IncentiveInstantWin with (RowLock) set DisallowEdit=" & LockedStatus(LoopCtr) & ", " & _
                        "RequiredFromTemplate=" & IIf(LockedStatus(LoopCtr) = "1", "0", "RequiredFromTemplate") & " " & _
                        "where IncentiveInstantWinID=" & Conditions(LoopCtr) & ";"
            Case "PLU"
              sQuery = "update CPE_IncentivePLUs with (RowLock) set DisallowEdit=" & LockedStatus(LoopCtr) & ", " & _
                        "RequiredFromTemplate=" & IIf(LockedStatus(LoopCtr) = "1", "0", "RequiredFromTemplate") & " " & _
                        "where IncentivePLUID=" & Conditions(LoopCtr) & ";"
          End Select
          MyCommon.QueryStr = sQuery
          MyCommon.LRT_Execute()
        Next
      End If
    End If
  End If
  
  ' dig the offer info out of the database
  ' no one clicked anything
  MyCommon.QueryStr = "Select IncentiveID, ClientOfferID, IncentiveName as Name, CPE.Description, PromoClassID, CRMEngineID, Priority," & _
                      "StartDate, EndDate, EveryDOW, EligibilityStartDate, EligibilityEndDate, TestingStartDate, TestingEndDate, " & _
                      "P1DistQtyLimit, P1DistTimeType, P1DistPeriod, P2DistQtyLimit, P2DistTimeType, P2DistPeriod, P3DistQtyLimit, P3DistTimeType, P3DistPeriod, " & _
                      "EnableImpressRpt, EnableRedeemRpt, CreatedDate, CPE.LastUpdate, CPE.Deleted, CPEOARptDate, CPEOADeploySuccessDate, CPEOADeployRpt, " & _
                      "CRMRestricted, StatusFlag, OC.Description as CategoryName, IsTemplate, FromTemplate from CPE_Incentives as CPE with (NoLock) " & _
                      "left join OfferCategories as OC with (NoLock) on CPE.PromoClassID=OfferCategoryID " & _
                      "where IncentiveID=" & OfferID & ";"
  rst = MyCommon.LRT_Select
  
  For Each row In rst.Rows
    Name = MyCommon.NZ(row.Item("Name"), "")
    isTemplate = MyCommon.NZ(row.Item("IsTemplate"), False)
    FromTemplate = MyCommon.NZ(row.Item("FromTemplate"), False)
    StatusFlag = MyCommon.NZ(row.Item("StatusFlag"), 0)
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
  
  IF Not isTemplate Then
    DeleteBtnDisabled = IIf(Logix.UserRoles.EditOffer, "", " disabled=""disabled""")
  Else
    DeleteBtnDisabled = IIf(Logix.UserRoles.EditTemplates, "", " disabled=""disabled""")
  End If
  

  ActiveSubTab = IIf(isTemplate, 206, 205)
  
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
<form action="CAM-offer-con.aspx" id="mainform" name="mainform">
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
        If (Logix.UserRoles.EditTemplates And isTemplate) Then
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
        MyCommon.QueryStr = "select incentiveId from CPE_Incentives with (NoLock) where CreatedDate = LastUpdate and IncentiveID=" & OfferID
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
              MyCommon.QueryStr = "select ICG.IncentiveCustomerID, CG.CustomerGroupID, CG.NewCardholders, CG.AnyCAMCardholder, Name, PhraseID, " & _
                                  " ExcludedUsers, DisallowEdit, RequiredFromTemplate from CPE_IncentiveCustomerGroups as ICG with (NoLock) " & _
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
                  Send("<td></td>")
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
                  '    Sendb("<input type=""button"" class=""ex"" id=""customerDelete"" name=""customerDelete"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('/logix/CAM/CAM-offer-con.aspx?mode=Delete&Option=Customer&OfferID=" & OfferID & "')}"" value=""X"" />")
                  'Else
                  Sendb("<input type=""button"" class=""ex"" id=""customerDelete"" name=""customerDelete"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('/logix/CAM/CAM-offer-con.aspx?mode=Delete&Option=Customer&OfferID=" & OfferID & "')}"" value=""X"" />")
                  'End If
                End If
                Send("  </td>")
                Send("  <td>")
                If (i > 1 And MyCommon.NZ(row.Item("ExcludedUsers"), False) = False) Then
                  Send("    " & Copient.PhraseLib.Lookup("term.or", LanguageID))
                End If
                Send("  </td>")
                Send("  <td>")
                Send("    <a href=""javascript:openPopup('/logix/CPEoffer-con-customer.aspx?OfferID=" & OfferID & "')"">" & Copient.PhraseLib.Lookup("term.customer", LanguageID) & "</a>")
                Send("  </td>")
                Send("  <td>")
                If (MyCommon.NZ(row.Item("ExcludedUsers"), False) = True) Then
                  Sendb(StrConv(Copient.PhraseLib.Lookup("term.excluding", LanguageID), VbStrConv.Lowercase) & " ")
                End If
                If (MyCommon.NZ(row.Item("CustomerGroupID"), -1) > 2) AndAlso (MyCommon.NZ(row.Item("NewCardholders"), 0) = 0) AndAlso (MyCommon.NZ(row.Item("AnyCAMCardholder"), 0) = 0) Then
                  Sendb("<a href=""/logix/cgroup-edit.aspx?CustomerGroupID=" & row.Item("CustomerGroupID") & """>")
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
                  Sendb(Copient.PhraseLib.Lookup("term.undefined", LanguageID))
                  Sendb(" <span class=""red"">")
                  Sendb("(" & Copient.PhraseLib.Lookup("term.requiredbytemplate", LanguageID) & ")")
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
            
            <!-- PRODUCT CONDITIONS -->
            <%
              t = 1
              ' Find the currently selected groups on page load:
              MyCommon.QueryStr = "select IPG.IncentiveProductGroupID, PG.ProductGroupID, PG.Name, PG.PhraseID, PG.AnyProduct, UT.PhraseID as UnitPhraseID, ExcludedProducts, ProductComboID, " & _
                                  " QtyForIncentive, QtyUnitType, AccumMin, AccumLimit, AccumPeriod, DisallowEdit, RequiredFromTemplate, Disqualifier from CPE_IncentiveProductGroups as IPG with (NoLock) " & _
                                  " left join ProductGroups as PG with (NoLock) on PG.ProductGroupID=IPG.ProductGroupID " & _
                                  " left join CPE_RewardOptions as RO with (NoLock) on IPG.RewardOptionID=RO.RewardOptionID " & _
                                  " left join CPE_UnitTypes as UT with (NoLock) on UT.UnitTypeID=IPG.QtyUnitType " & _
                                  " where IPG.RewardOptionID=" & roid & "and IPG.Deleted=0 and Disqualifier=0 " & _
                                  " order by AnyProduct DESC, Name;"
              rst = MyCommon.LRT_Select
              ' Also, go ahead and find the currently DISQUALIFIED groups on page load too:
              MyCommon.QueryStr = "select IPG.IncentiveProductGroupID, PG.ProductGroupID, PG.Name, PG.PhraseID, PG.AnyProduct, UT.PhraseID as UnitPhraseID, ExcludedProducts, ProductComboID, " & _
                                  " QtyForIncentive, QtyUnitType, AccumMin, AccumLimit, AccumPeriod, DisallowEdit, RequiredFromTemplate, Disqualifier from CPE_IncentiveProductGroups as IPG with (NoLock) " & _
                                  " left join ProductGroups as PG with (NoLock) on PG.ProductGroupID=IPG.ProductGroupID " & _
                                  " left join CPE_RewardOptions as RO with (NoLock) on IPG.RewardOptionID=RO.RewardOptionId " & _
                                  " left join CPE_UnitTypes as UT with (NoLock) on UT.UnitTypeID=IPG.QtyUnitType " & _
                                  " where IPG.RewardOptionID=" & roid & "and IPG.Deleted=0 and Disqualifier=1 " & _
                                  " order by Name;"
              rst2 = MyCommon.LRT_Select
              ' If there's a product condition, then continue:
              If (rst.Rows.Count > 0) Then
                Send("<tr class=""shadeddark"">")
                Send("  <td colspan=""" & 4 & """>")
                Send("    <h3>")
                Send("      " & Copient.PhraseLib.Lookup("term.productconditions", LanguageID))
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
                  Send("<td></td>")
                End If
                Send("</tr>")
              End If
              i = 1
              For Each row In rst.Rows
                ' Get the excluded group details, if any
                If MyCommon.NZ(row.Item("ExcludedProducts"), False) Then
                  ExcludedProductGroupID = MyCommon.NZ(row.Item("ProductGroupID"), 0)
                  MyCommon.QueryStr = "select Name from ProductGroups with (NoLock) where ProductGroupID=" & ExcludedProductGroupID & ";"
                  rst3 = MyCommon.LRT_Select
                  If rst3.Rows.Count > 0 Then
                    ExcludedProductGroupName = MyCommon.NZ(rst3.Rows(0).Item("Name"), "")
                  End If
                End If
              Next
              For Each row In rst.Rows
                IncentiveID = row.Item("IncentiveProductGroupID")
                ' we got in the loop so there is a product condition
                isProduct = True
                If Not MyCommon.NZ(row.Item("ExcludedProducts"), False) Then
                  Send("<tr class=""shaded"">")
                  Send("  <td>")
                  'If (i = 1) Then
                  'If rst2.Rows.Count = 0 Then
                  If rst.Rows.Count > 1 Then
                    If (Logix.UserRoles.EditOffer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False))) Then
                      If (MyCommon.NZ(row.Item("RequiredFromTemplate"), False) And Not isTemplate) Then
                        Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('/logix/CAM/CAM-offer-con.aspx?mode=Delete&Option=Product&OfferID=" & OfferID & "&IncentiveProductGroupID=" & row.Item("IncentiveProductGroupID") & "')}"" value=""X"" />")
                      Else
                        Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('/logix/CAM/CAM-offer-con.aspx?mode=Delete&Option=Product&OfferID=" & OfferID & "&IncentiveProductGroupID=" & row.Item("IncentiveProductGroupID") & "')}"" value=""X"" />")
                      End If
                    ElseIf (Logix.UserRoles.EditTemplates) Then
                      Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('/logix/CAM/CAM-offer-con.aspx?mode=Delete&Option=Product&OfferID=" & OfferID & "&IncentiveProductGroupID=" & row.Item("IncentiveProductGroupID") & "')}"" value=""X"" />")
                    Else
                      Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('/logix/CAM/CAM-offer-con.aspx?mode=Delete&Option=Product&OfferID=" & OfferID & "&IncentiveProductGroupID=" & row.Item("IncentiveProductGroupID") & "')}"" value=""X"" />")
                    End If
                  Else
                    If rst2.Rows.Count = 0 Then
                      Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CAM-offer-con.aspx?mode=Delete&Option=Product&OfferID=" & OfferID & "&IncentiveProductGroupID=" & row.Item("IncentiveProductGroupID") & "')}"" value=""X"" />")
                    Else
                      Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CAM-offer-con.aspx?mode=Delete&Option=Product&OfferID=" & OfferID & "&IncentiveProductGroupID=" & row.Item("IncentiveProductGroupID") & "')}"" value=""X"" />")
                    End If
                  End If
                  'End If
                  Send("  </td>")
                  Send("  <td>")
                  ' lets spit out the ProductComboID
                  If (i > 1 And MyCommon.NZ(row.Item("ExcludedProducts"), False) = False) Then
                    If (MyCommon.NZ(row.Item("ProductComboID"), 0) = 0) Then
                      ' single
                      Send("<a href=""/logix/CAM/CAM-offer-con.aspx?mode=ChangeProductCombo&pc=1&OfferID=" & OfferID & """>" & Copient.PhraseLib.Lookup("term.and", LanguageID) & "</a>")
                      MyCommon.QueryStr = "update CPE_RewardOptions set ProductComboID=1 where RewardOptionID=" & roid
                      MyCommon.LRT_Execute()
                    ElseIf (MyCommon.NZ(row.Item("ProductComboID"), 0) = 1) Then
                      ' and
                      Send("<a href=""/logix/CAM/CAM-offer-con.aspx?mode=ChangeProductCombo&pc=1&OfferID=" & OfferID & """>" & Copient.PhraseLib.Lookup("term.and", LanguageID) & "</a>")
                    Else
                      ' or
                      Send("<a href=""/logix/CAM/CAM-offer-con.aspx?mode=ChangeProductCombo&pc=2&OfferID=" & OfferID & """>" & Copient.PhraseLib.Lookup("term.or", LanguageID) & "</a>")
                    End If
                  End If
                  Send("  </td>")
                  Send("  <td>")
                  Send("    <a href=""javascript:openPopup('/logix/CPEoffer-con-product.aspx?OfferID=" & OfferID & "&Disqualifier=0&IncentiveProductGroupID=" & IncentiveID & "')"">" & Copient.PhraseLib.Lookup("term.product", LanguageID) & "</a>")
                  Send("  </td>")
                  Send("  <td>")
                  If (MyCommon.NZ(row.Item("ProductGroupID"), -1) > 1) Then
                    If (MyCommon.NZ(row.Item("ExcludedProducts"), False) = True) Then
                      Sendb(Copient.PhraseLib.Lookup("term.anyproduct", LanguageID) & " ")
                      Sendb(StrConv(Copient.PhraseLib.Lookup("term.excluding", LanguageID), VbStrConv.Lowercase) & " ")
                    End If
                    Sendb("<a href=""/logix/pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("ProductGroupID"), -1) & """>")
                    If IsDBNull(row.Item("PhraseID")) Then
                      Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
                    Else
                      If (MyCommon.NZ(row.Item("PhraseID"), 0) = 0) Then
                        Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
                      Else
                        Sendb(MyCommon.SplitNonSpacedString(Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID), 25) & "</a>")
                      End If
                    End If
                  ElseIf (IsDBNull(row.Item("ProductGroupID")) AndAlso MyCommon.NZ(row.Item("RequiredFromTemplate"), False)) Then
                    Sendb(Copient.PhraseLib.Lookup("term.undefined", LanguageID))
                    Sendb(" <span class=""red"">")
                    Sendb("(" & Copient.PhraseLib.Lookup("term.requiredbytemplate", LanguageID) & ")")
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
                    If (MyCommon.NZ(row.Item("ProductGroupID"), -1) = 1) AndAlso (ExcludedProductGroupID > 0) Then
                      Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.excluding", LanguageID), VbStrConv.Lowercase) & " ")
                      Send("<a href=""/logix/pgroup-edit.aspx?ProductGroupID=" & ExcludedProductGroupID & """>" & ExcludedProductGroupName & "</a>")
                    End If
                  End If
                  Send("  </td>")
                  ' Find the per-tier values:
                  t = 1
                  MyCommon.QueryStr = "select IPG.IncentiveProductGroupID, IPG.Disqualifier, IPGT.TierLevel, IPGT.Quantity as QtyForIncentive from CPE_IncentiveProductGroups as IPG with (NoLock) " & _
                                      "left join CPE_IncentiveProductGroupTiers as IPGT with (NoLock) on IPGT.IncentiveProductGroupID=IPG.IncentiveProductGroupID " & _
                                      "where IPG.RewardOptionID=" & roid & " and IPG.Deleted=0 and IPG.IncentiveProductGroupID=" & IncentiveID & " order by TierLevel;"
                  rst3 = MyCommon.LRT_Select
                  If rst3.Rows.Count = 0 Then
                    Send("  <td colspan=""" & TierLevels & """>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                  Else
                    While t <= TierLevels
                      If t > rst3.Rows.Count Then
                        Send("  <td>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                      Else
                        Send("  <td>")
                        ' QtyForIncentive, QtyUnitType, AccumMin, AccumLimit, AccumPeriod, UnitDescription
                        If (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 1) Then
                          Sendb(Math.Truncate(MyCommon.NZ(rst3.Rows(t - 1).Item("QtyForIncentive"), 0)) & " " & StrConv(Copient.PhraseLib.Lookup("term.required", LanguageID), VbStrConv.Lowercase))
                        ElseIf (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 2) Then
                          Sendb(FormatCurrency(MyCommon.NZ(rst3.Rows(t - 1).Item("QtyForIncentive"), 0)) & " " & StrConv(Copient.PhraseLib.Lookup("term.required", LanguageID), VbStrConv.Lowercase))
                        ElseIf (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 3) Or (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 4) Then
                          Sendb(MyCommon.NZ(rst3.Rows(t - 1).Item("QtyForIncentive"), 0) & " " & Copient.PhraseLib.Lookup("term.lbsgals", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.required", LanguageID), VbStrConv.Lowercase))
                        End If
                        Send("<br />")
                        If MyCommon.NZ(row.Item("AccumLimit"), 0) <> 0 OrElse MyCommon.NZ(row.Item("AccumPeriod"), 0) <> 0 OrElse MyCommon.NZ(row.Item("AccumMin"), 0) <> 0 Then
                          ' There's at least some accumulation data set, so display it:
                          ' Limit value
                          If MyCommon.NZ(row.Item("AccumLimit"), 0) > 0 Then
                            Sendb(Copient.PhraseLib.Lookup("term.limit", LanguageID) & " ")
                            If MyCommon.NZ(row.Item("QtyUnitType"), 0) = 1 Then
                              Sendb(Math.Truncate(MyCommon.NZ(row.Item("AccumLimit"), 0)))
                            ElseIf MyCommon.NZ(row.Item("QtyUnitType"), 0) = 2 Then
                              Sendb(FormatCurrency(MyCommon.NZ(row.Item("AccumLimit"), 0)))
                            ElseIf (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 3) Or (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 4) Then
                              Sendb(MyCommon.NZ(row.Item("AccumLimit"), 0) & " " & Copient.PhraseLib.Lookup("term.lbsgals", LanguageID))
                            End If
                          Else
                            Sendb(Copient.PhraseLib.Lookup("term.nolimit", LanguageID))
                          End If
                          ' Period value
                          If MyCommon.NZ(row.Item("AccumPeriod"), 0) > 0 Then
                            Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.every", LanguageID), VbStrConv.Lowercase) & " ")
                            If MyCommon.NZ(row.Item("AccumPeriod"), 0) <= 1 Then
                              Sendb(StrConv(Copient.PhraseLib.Lookup("term.day", LanguageID), VbStrConv.Lowercase))
                            Else
                              Sendb(MyCommon.NZ(row.Item("AccumPeriod"), 0) & " " & StrConv(Copient.PhraseLib.Lookup("term.days", LanguageID), VbStrConv.Lowercase))
                            End If
                          End If
                          ' Minimum value
                          If MyCommon.NZ(row.Item("AccumMin"), 0) > 0 Then
                            Sendb(", " & StrConv(Copient.PhraseLib.Lookup("term.minimum", LanguageID), VbStrConv.Lowercase) & " ")
                            If MyCommon.NZ(row.Item("QtyUnitType"), 0) = 1 Then
                              Send(Math.Truncate(MyCommon.NZ(row.Item("AccumMin"), 0)))
                            ElseIf MyCommon.NZ(row.Item("QtyUnitType"), 0) = 2 Then
                              Send(FormatCurrency(MyCommon.NZ(row.Item("AccumMin"), 0)))
                            ElseIf (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 3) Or (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 4) Then
                              Send(MyCommon.NZ(row.Item("AccumMin"), 0) & " " & Copient.PhraseLib.Lookup("term.lbsgals", LanguageID))
                            End If
                          Else
                            Send(", " & StrConv(Copient.PhraseLib.Lookup("term.nominimum", LanguageID), VbStrConv.Lowercase))
                          End If
                        End If
                        Send("  </td>")
                      End If
                      t += 1
                    End While
                  End If
                  If (isTemplate) Then
                    Send("  <td class=""templine"">")
                    %>
                    <input type="checkbox" id="chkLocked2" name="chkLocked" value="<%Sendb(MyCommon.NZ(row.Item("IncentiveProductGroupID"), 0))%>"<%Sendb(IIf(MyCommon.NZ(row.Item("DisallowEdit"), False)=True, " checked=""checked""", "")) %> onclick="javascript:updateLocked('lockProd<%Sendb(MyCommon.NZ(row.Item("IncentiveProductGroupID"), 0))%>', this.checked);"<%Sendb(IIf(IsDBNull(row.Item("ProductGroupID")) AndAlso MyCommon.NZ(row.Item("RequiredFromTemplate"), False), " disabled=""disabled""", "")) %> />
                    <input type="hidden" id="conType2" name="conType" value="Product" />
                    <input type="hidden" id="conProd<%Sendb(MyCommon.NZ(row.Item("IncentiveProductGroupID"), 0))%>" name="con" value="<%Sendb(MyCommon.NZ(row.Item("IncentiveProductGroupID"), 0))%>" />
                    <input type="hidden" id="lockProd<%Sendb(MyCommon.NZ(row.Item("IncentiveProductGroupID"), 0))%>" name="locked" value="<%Sendb(IIf(MyCommon.NZ(row.Item("DisallowEdit"), False)=True, "1", "0"))%>" />
                    <%
                    Send("  </td>")
                  ElseIf (FromTemplate) Then
                    Send("  <td class=""templine"">")
                    Send("    " & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "Yes", "No"))
                    Send("  </td>")
                  End If
                  Send("</tr>")
                  i += 1
                  End If
              Next
            %>
            
            <!-- PRODUCT DISQUALIFIERS -->
            <%
              t = 1
              If (rst2.Rows.Count > 0) Then
                Send("<tr class=""shadeddark"">")
                Send("  <td colspan=""" & 4 & """>")
                Send("    <h3>")
                Send("      " & Copient.PhraseLib.Lookup("term.productdisqualifiers", LanguageID))
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
                  Send("<td></td>")
                End If
                Send("</tr>")
              End If
              i = 1
              For Each row In rst2.Rows
                ' we got in the loop so there is a customer disqualifier set it as such
                isProductDisqualifier = True
                Send("<tr class=""shaded"">")
                Send("  <td>")
                If (i = 1) Then
                  If (Logix.UserRoles.EditOffer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False))) Then
                    If (MyCommon.NZ(row.Item("RequiredFromTemplate"), False) And Not isTemplate) Then
                      Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('/logix/CAM/CAM-offer-con.aspx?mode=Delete&Option=ProductDisqualifier&OfferID=" & OfferID & "&IncentiveProductGroupID=" & row.Item("IncentiveProductGroupID") & "')}"" value=""X"" />")
                    Else
                      Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('/logix/CAM/CAM-offer-con.aspx?mode=Delete&Option=ProductDisqualifier&OfferID=" & OfferID & "&IncentiveProductGroupID=" & row.Item("IncentiveProductGroupID") & "')}"" value=""X"" />")
                    End If
                  ElseIf (Logix.UserRoles.EditTemplates) Then
                    Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('/logix/CAM/CAM-offer-con.aspx?mode=Delete&Option=ProductDisqualifier&OfferID=" & OfferID & "&IncentiveProductGroupID=" & row.Item("IncentiveProductGroupID") & "')}"" value=""X"" />")
                  Else
                    Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('/logix/CAM/CAM-offer-con.aspx?mode=Delete&Option=ProductDisqualifier&OfferID=" & OfferID & "&IncentiveProductGroupID=" & row.Item("IncentiveProductGroupID") & "')}"" value=""X"" />")
                  End If
                End If
                Send("  </td>")
                Send("  <td>")
                Send("    " & Copient.PhraseLib.Lookup("term.not", LanguageID))
                Send("  </td>")
                Send("  <td>")
                Send("    <a href=""javascript:openPopup('/logix/CPEoffer-con-product.aspx?OfferID=" & OfferID & "&Disqualifier=1&IncentiveProductGroupID=" & row.Item("IncentiveProductGroupID") & "')"">" & Copient.PhraseLib.Lookup("term.product", LanguageID) & "</a>")
                Send("  </td>")
                Send("  <td>")
                If (MyCommon.NZ(row.Item("ProductGroupID"), -1) > 1) Then
                  If (MyCommon.NZ(row.Item("ExcludedProducts"), False) = True) Then
                    Sendb(StrConv(Copient.PhraseLib.Lookup("term.excluding", LanguageID), VbStrConv.Lowercase) & " ")
                  End If
                  Sendb("<a href=""/logix/pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("ProductGroupID"), -1) & """>")
                  If IsDBNull(row.Item("PhraseID")) Then
                    Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
                  Else
                    If (MyCommon.NZ(row.Item("PhraseID"), 0) = 0) Then
                      Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
                    Else
                      Sendb(MyCommon.SplitNonSpacedString(Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID), 25) & "</a>")
                    End If
                  End If
                ElseIf (IsDBNull(row.Item("ProductGroupID")) AndAlso MyCommon.NZ(row.Item("RequiredFromTemplate"), False)) Then
                  Sendb(Copient.PhraseLib.Lookup("term.undefined", LanguageID))
                  Sendb(" <span class=""red"">")
                  Sendb("(" & Copient.PhraseLib.Lookup("term.requiredbytemplate", LanguageID) & ")")
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
                ' Find the per-tier values:
                t = 1
                MyCommon.QueryStr = "select IPG.IncentiveProductGroupID, IPG.Disqualifier, IPGT.TierLevel, IPGT.Quantity as QtyForIncentive from CPE_IncentiveProductGroups as IPG with (NoLock) " & _
                                    "left join CPE_IncentiveProductGroupTiers as IPGT with (NoLock) on IPGT.IncentiveProductGroupID=IPG.IncentiveProductGroupID " & _
                                    "where IPG.RewardOptionID=" & roid & " and IPG.Deleted=0 order by TierLevel;"
                rst3 = MyCommon.LRT_Select
                If rst3.Rows.Count = 0 Then
                  Send("  <td colspan=""" & TierLevels & """>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                Else
                  While t <= TierLevels
                    If t > rst3.Rows.Count Then
                      Send("  <td>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                    Else
                      Send("  <td>")
                      ' QtyForIncentive,QtyUnitType,AccumMin,AccumLimit,AccumPeriod,UnitDescription
                      If Not MyCommon.NZ(row.Item("ExcludedProducts"), False) Then
                        If (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 1) Then
                          Sendb(Math.Truncate(MyCommon.NZ(rst3.Rows(t - 1).Item("QtyForIncentive"), 0)) & " " & StrConv(Copient.PhraseLib.Lookup("term.required", LanguageID), VbStrConv.Lowercase))
                        ElseIf (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 2) Then
                          Sendb(FormatCurrency(MyCommon.NZ(rst3.Rows(t - 1).Item("QtyForIncentive"), 0)) & " " & StrConv(Copient.PhraseLib.Lookup("term.required", LanguageID), VbStrConv.Lowercase))
                        ElseIf (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 3) Or (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 4) Then
                          Sendb(MyCommon.NZ(rst3.Rows(t - 1).Item("QtyForIncentive"), 0) & " " & Copient.PhraseLib.Lookup("term.lbsgals", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.required", LanguageID), VbStrConv.Lowercase))
                        End If
                        Send("<br />")
                        If MyCommon.NZ(row.Item("AccumLimit"), 0) <> 0 OrElse MyCommon.NZ(row.Item("AccumPeriod"), 0) <> 0 OrElse MyCommon.NZ(row.Item("AccumMin"), 0) <> 0 Then
                          ' There's at least some accumulation data set, so display it:
                          ' Limit value
                          If MyCommon.NZ(row.Item("AccumLimit"), 0) > 0 Then
                            Sendb(Copient.PhraseLib.Lookup("term.limit", LanguageID) & " ")
                            If MyCommon.NZ(row.Item("QtyUnitType"), 0) = 1 Then
                              Sendb(row.Item("AccumLimit"))
                            ElseIf MyCommon.NZ(row.Item("QtyUnitType"), 0) = 2 Then
                              Sendb(FormatCurrency(row.Item("AccumLimit")))
                            ElseIf (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 3) Or (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 4) Then
                              Sendb(row.Item("AccumLimit") & " " & Copient.PhraseLib.Lookup("term.lbsgals", LanguageID))
                            End If
                          Else
                            Sendb(Copient.PhraseLib.Lookup("term.nolimit", LanguageID))
                          End If
                          ' Period value
                          If MyCommon.NZ(row.Item("AccumPeriod"), 0) > 0 Then
                            Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.every", LanguageID), VbStrConv.Lowercase) & " ")
                            If row.Item("AccumPeriod") <= 1 Then
                              Sendb(StrConv(Copient.PhraseLib.Lookup("term.day", LanguageID), VbStrConv.Lowercase))
                            Else
                              Sendb(row.Item("AccumPeriod") & " " & StrConv(Copient.PhraseLib.Lookup("term.days", LanguageID), VbStrConv.Lowercase))
                            End If
                          End If
                          ' Minimum value
                          If MyCommon.NZ(row.Item("AccumMin"), 0) > 0 Then
                            Sendb(", " & StrConv(Copient.PhraseLib.Lookup("term.minimum", LanguageID), VbStrConv.Lowercase) & " ")
                            If MyCommon.NZ(row.Item("QtyUnitType"), 0) = 1 Then
                              Send(Math.Truncate(MyCommon.NZ(row.Item("AccumMin"), 0)))
                            ElseIf MyCommon.NZ(row.Item("QtyUnitType"), 0) = 2 Then
                              Send(FormatCurrency(MyCommon.NZ(row.Item("AccumMin"), 0)))
                            ElseIf (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 3) Or (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 4) Then
                              Send(MyCommon.NZ(row.Item("AccumMin"), 0) & " " & Copient.PhraseLib.Lookup("term.lbsgals", LanguageID))
                            End If
                          Else
                            Send(", " & StrConv(Copient.PhraseLib.Lookup("term.nominimum", LanguageID), VbStrConv.Lowercase))
                          End If
                        End If
                      End If
                      Send("  </td>")
                    End If
                    t += 1
                  End While
                End If
                If (isTemplate) Then
                  Send("  <td class=""templine"">")
                  %>
                  <input type="checkbox" id="chkLocked10" name="chkLocked" value="<%Sendb(MyCommon.NZ(row.Item("IncentiveProductGroupID"), 0))%>"<%Sendb(IIf(MyCommon.NZ(row.Item("DisallowEdit"), False)=True, " checked=""checked""", "")) %> onclick="javascript:updateLocked('lockProd<%Sendb(MyCommon.NZ(row.Item("IncentiveProductGroupID"), 0))%>', this.checked);"<%Sendb(IIf(IsDBNull(row.Item("ProductGroupID")) AndAlso MyCommon.NZ(row.Item("RequiredFromTemplate"), False), " disabled=""disabled""", "")) %> />
                  <input type="hidden" id="conType10" name="conType" value="Product" />
                  <input type="hidden" id="conProd<%Sendb(MyCommon.NZ(row.Item("IncentiveProductGroupID"), 0))%>" name="con" value="<%Sendb(MyCommon.NZ(row.Item("IncentiveProductGroupID"), 0))%>" />
                  <input type="hidden" id="lockProd<%Sendb(MyCommon.NZ(row.Item("IncentiveProductGroupID"), 0))%>" name="locked" value="<%Sendb(IIf(MyCommon.NZ(row.Item("DisallowEdit"), False)=True, "1", "0"))%>" />
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
                  Send("<td></td>")
                End If
                Send("</tr>")
              End If
              For Each row In rst.Rows
                isPoint = True
                Send("<tr class=""shaded"">")
                Send("  <td>")
                If (Logix.UserRoles.EditOffer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False))) Then
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('/logix/CAM/CAM-offer-con.aspx?mode=Delete&Option=Point&OfferID=" & OfferID & "')}"" value=""X"" />")
                ElseIf (Logix.UserRoles.EditTemplates) Then
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('/logix/CAM/CAM-offer-con.aspx?mode=Delete&Option=Point&OfferID=" & OfferID & "')}"" value=""X"" />")
                Else
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('/logix/CAM/CAM-offer-con.aspx?mode=Delete&Option=Point&OfferID=" & OfferID & "')}"" value=""X"" />")
                End If
                Send("  </td>")
                Send("  <td>")
                Send("  </td>")
                Send("  <td>")
                Send("    <a href=""javascript:openPopup('/logix/CPEoffer-con-point.aspx?OfferID=" & OfferID & "&IncentivePointsID=" & MyCommon.NZ(row.Item("IncentivePointsID"), 0) & "')"">" & Copient.PhraseLib.Lookup("term.points", LanguageID) & "</a>")
                Send("  </td>")
                Send("  <td>")
                If (MyCommon.NZ(row.Item("ProgramID"), -1) > -1) Then
                  Sendb("    <a href=""/logix/point-edit.aspx?ProgramGroupID=" & row.Item("ProgramID") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("ProgramName"), ""), 25) & "</a>")
                ElseIf (IsDBNull(row.Item("ProgramID")) AndAlso MyCommon.NZ(row.Item("RequiredFromTemplate"), False)) Then
                  Sendb(Copient.PhraseLib.Lookup("term.undefined", LanguageID))
                  Sendb(" <span class=""red"">")
                  Sendb("(" & Copient.PhraseLib.Lookup("term.requiredbytemplate", LanguageID) & ")")
                  Sendb("</span>")
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
            
            <!-- STORED VALUE CONDITIONS -->
            <%
              t = 1
              MyCommon.QueryStr = "select ISVP.IncentiveStoredValueID, ISVP.SVProgramID, SVP.Name, QtyForIncentive, DisallowEdit, RequiredFromTemplate " & _
                                  "from CPE_IncentiveStoredValuePrograms as ISVP with (NoLock) " & _
                                  "left join StoredValuePrograms as SVP with (NoLock) " & _
                                  "on SVP.SVProgramID=ISVP.SVProgramID where RewardOptionID=" & roid & ";"
              rst = MyCommon.LRT_Select
              If (rst.Rows.Count > 0) Then
                Send("<tr class=""shadeddark"">")
                Send("  <td colspan=""" & 4 & """>")
                Send("    <h3>")
                Send("      " & Copient.PhraseLib.Lookup("term.storedvalueconditions", LanguageID))
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
                  Send("<td></td>")
                End If
                Send("</tr>")
              End If
              For Each row In rst.Rows
                isStoredValue = True
                Send("<tr class=""shaded"">")
                Send("  <td>")
                If (Logix.UserRoles.EditOffer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False))) Then
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('/logix/CAM/CAM-offer-con.aspx?mode=Delete&Option=StoredValue&OfferID=" & OfferID & "')}"" value=""X"" />")
                ElseIf (Logix.UserRoles.EditTemplates) Then
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('/logix/CAM/CAM-offer-con.aspx?mode=Delete&Option=StoredValue&OfferID=" & OfferID & "')}"" value=""X"" />")
                Else
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('/logix/CAM/CAM-offer-con.aspx?mode=Delete&Option=StoredValue&OfferID=" & OfferID & "')}"" value=""X"" />")
                End If
                Send("  </td>")
                Send("  <td>")
                Send("  </td>")
                Send("  <td>")
                Send("    <a href=""javascript:openPopup('/logix/CPEoffer-con-sv.aspx?OfferID=" & OfferID & "')"">" & Copient.PhraseLib.Lookup("term.storedvalue", LanguageID) & "</a>")
                Send("  </td>")
                Send("  <td>")
                If (MyCommon.NZ(row.Item("SVProgramID"), -1) > -1) Then
                  Sendb("<a href=""/logix/SV-edit.aspx?ProgramGroupID=" & row.Item("SVProgramID") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
                End If
                Send("  </td>")
                ' Find the per-tier values:
                t = 1
                MyCommon.QueryStr = "select ISVP.IncentiveStoredValueID, ISVPT.TierLevel, ISVPT.Quantity " & _
                                    "from CPE_IncentiveStoredValuePrograms as ISVP with (NoLock) " & _
                                    "left join CPE_IncentiveStoredValueProgramTiers as ISVPT with (NoLock) on ISVPT.IncentiveStoredValueID=ISVP.IncentiveStoredValueID " & _
                                    "where ISVP.RewardOptionID=" & roid & ";"
                rst2 = MyCommon.LRT_Select
                If rst2.Rows.Count = 0 Then
                  Send("  <td colspan=""" & TierLevels & """>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                Else
                  While t <= TierLevels
                    If t > rst2.Rows.Count Then
                      Send("  <td>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                    Else
                      Send("  <td>")
                      Send("    " & CInt(MyCommon.NZ(rst2.Rows(t - 1).Item("Quantity"), "0")) & " " & StrConv(Copient.PhraseLib.Lookup("term.units", LanguageID), VbStrConv.Lowercase) & " " & StrConv(Copient.PhraseLib.Lookup("term.required", LanguageID), VbStrConv.Lowercase))
                      Send("  </td>")
                    End If
                    t += 1
                  End While
                End If
                If (isTemplate) Then
                  Send("  <td class=""templine"">")
                  %>
                  <input type="checkbox" id="chkLocked6" name="chkLocked" value="<%Sendb(MyCommon.NZ(row.Item("IncentiveStoredValueID"), 0))%>"<%Sendb(IIf(MyCommon.NZ(row.Item("DisallowEdit"), False)=True, " checked=""checked""", "")) %> onclick="javascript:updateLocked('lockSV<%Sendb(MyCommon.NZ(row.Item("IncentiveStoredValueID"), 0))%>', this.checked);"<%Sendb(IIf(IsDBNull(row.Item("SVProgramID")) AndAlso MyCommon.NZ(row.Item("RequiredFromTemplate"), False), " disabled=""disabled""", "")) %> />
                  <input type="hidden" id="conType6" name="conType" value="StoredValue" />
                  <input type="hidden" id="conSV<%Sendb(MyCommon.NZ(row.Item("IncentiveStoredValueID"), 0))%>" name="con" value="<%Sendb(MyCommon.NZ(row.Item("IncentiveStoredValueID"), 0))%>" />
                  <input type="hidden" id="lockSV<%Sendb(MyCommon.NZ(row.Item("IncentiveStoredValueID"), 0))%>" name="locked" value="<%Sendb(IIf(MyCommon.NZ(row.Item("DisallowEdit"), False)=True, "1", "0"))%>" />
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
            
            <!-- DAY CONDITIONS -->
            <%
              t = 1
              MyCommon.QueryStr = "select DOWID, DayName, PhraseID from CPE_DaysOfWeek DW with (NoLock);"
              rst = MyCommon.LRT_Select
              MyCommon.QueryStr = "select IncentiveDOWID, DOWID, DisallowEdit from CPE_IncentiveDOW with (NoLock) " & _
                                  "where IncentiveID=" & OfferID & " and Deleted=0;"
              rst2 = MyCommon.LRT_Select
              For Each row In rst.Rows
                If rst2.Rows.Count >= 7 Then
                  Days = Copient.PhraseLib.Lookup("term.everyday", LanguageID)
                  DaysLocked = MyCommon.NZ(rst2.Rows(0).Item("DisallowEdit"), False)
                Else
                  For Each row2 In rst2.Rows
                    If (MyCommon.NZ(row2.Item("DOWID"), 0) = MyCommon.NZ(row.Item("DOWID"), 0)) Then
                      If (Days = "") Then
                        Days = Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID)
                      Else
                        Days = Days & ", " & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID)
                      End If
                    End If
                    DaysLocked = MyCommon.NZ(row2.Item("DisallowEdit"), False)
                  Next
                End If
              Next
              If (Days <> "") Then
                isDay = True
                Send("<tr class=""shadeddark"">")
                Send("  <td colspan=""4"">")
                Send("    <h3>")
                Send("      " & Copient.PhraseLib.Lookup("term.dayconditions", LanguageID))
                Send("    </h3>")
                Send("  </td>")
                Send("  <td colspan=""" & TierLevels & """>")
                Send("  </td>")
                If (isTemplate Or FromTemplate) Then
                  Send("<td></td>")
                End If
                Send("</tr>")
                Send("<tr class=""shaded"">")
                Send("  <td>")
                If (Logix.UserRoles.EditOffer And Not (FromTemplate And DaysLocked)) Then
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('/logix/CAM/CAM-offer-con.aspx?mode=Delete&Option=Day&OfferID=" & OfferID & "')}"" value=""X"" />")
                ElseIf (Logix.UserRoles.EditTemplates) Then
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('/logix/CAM/CAM-offer-con.aspx?mode=Delete&Option=Day&OfferID=" & OfferID & "')}"" value=""X"" />")
                Else
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('/logix/CAM/CAM-offer-con.aspx?mode=Delete&Option=Day&OfferID=" & OfferID & "')}"" value=""X"" />")
                End If
                Send("  </td>")
                Send("  <td>")
                Send("  </td>")
                Send("  <td>")
                Send("    <a href=""javascript:openPopup('/logix/CPEoffer-con-day.aspx?OfferID=" & OfferID & "')"">" & Copient.PhraseLib.Lookup("term.day", LanguageID) & "</a>")
                Send("  </td>")
                Send("  <td>")
                Send("    " & Copient.PhraseLib.Lookup("term.valid", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.on", LanguageID), VbStrConv.Lowercase))
                Send("  </td>")
                Send("  <td colspan=""" & TierLevels & """>")
                Send("    " & Days)
                Send("  </td>")
                If (isTemplate) Then
                  Send("  <td class=""templine"">")
                  %>
                  <input type="checkbox" id="chkLocked5" name="chkLocked" value="<%Sendb(OfferID)%>"<%Sendb(IIf(DaysLocked, " checked=""checked""", "")) %> onclick="javascript:updateLocked('lockDay<%Sendb(OfferID)%>', this.checked);" />
                  <input type="hidden" id="conType5" name="conType" value="Days" />
                  <input type="hidden" id="conDay<%Sendb(OfferID)%>" name="con" value="<%Sendb(OfferID)%>" />
                  <input type="hidden" id="lockDay<%Sendb(OfferID)%>" name="locked" value="<%Sendb(IIf(DaysLocked, "1", "0"))%>" />
                  <%
                  Send("  </td>")
                ElseIf (FromTemplate) Then
                  Send("  <td class=""templine"">")
                  Send("    " & IIf(DaysLocked, "Yes", "No"))
                  Send("  </td>")
                End If
                Send("</tr>")
              End If
            %>
            
            <!-- TIME CONDITIONS -->
            <%
              t = 1
              MyCommon.QueryStr = "select StartTime, EndTime, DisallowEdit from CPE_IncentiveTOD with (NoLock) where IncentiveID=" & OfferID
              rst = MyCommon.LRT_Select
              If (rst.Rows.Count > 0) Then
                For i = 0 To rst.Rows.Count - 1
                  If (i > 0) Then Times &= "; "
                  Times &= MyCommon.NZ(rst.Rows(i).Item("StartTime"), "") & " - " & MyCommon.NZ(rst.Rows(i).Item("EndTime"), "")
                  TimeLocked = MyCommon.NZ(rst.Rows(i).Item("DisallowEdit"), False)
                Next
              End If
              If (Times <> "") Then
                isTime = True
                Send("<tr class=""shadeddark"">")
                Send("  <td colspan=""4"">")
                Send("    <h3>")
                Send("      " & Copient.PhraseLib.Lookup("term.timeconditions", LanguageID))
                Send("    </h3>")
                Send("  </td>")
                Send("  <td colspan=""" & TierLevels & """>")
                Send("  </td>")
                If (isTemplate Or FromTemplate) Then
                  Send("<td></td>")
                End If
                Send("</tr>")
                Send("<tr class=""shaded"">")
                Send("  <td>")
                If (Logix.UserRoles.EditOffer And Not (FromTemplate And TimeLocked)) Then
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('/logix/CAM/CAM-offer-con.aspx?mode=Delete&Option=Time&OfferID=" & OfferID & "')}"" value=""X"" />")
                ElseIf (Logix.UserRoles.EditTemplates) Then
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('/logix/CAM/CAM-offer-con.aspx?mode=Delete&Option=Time&OfferID=" & OfferID & "')}"" value=""X"" />")
                Else
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('/logix/CAM/CAM-offer-con.aspx?mode=Delete&Option=Time&OfferID=" & OfferID & "')}"" value=""X"" />")
                End If
                Send("  </td>")
                Send("  <td>")
                Send("  </td>")
                Send("  <td>")
                Send("    <a href=""javascript:openPopup('/logix/CPEoffer-con-time.aspx?OfferID=" & OfferID & "')"">" & Copient.PhraseLib.Lookup("term.time", LanguageID) & "</a>")
                Send("  </td>")
                Send("  <td>")
                Send("    " & Copient.PhraseLib.Lookup("term.valid", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.from", LanguageID), VbStrConv.Lowercase))
                Send("  </td>")
                Send("  <td colspan=""" & TierLevels & """>")
                Send("    " & Times)
                Send("  </td>")
                If (isTemplate) Then
                  Send("  <td class=""templine"">")
                  %>
                  <input type="checkbox" id="chkLocked8" name="chkLocked" value="<%Sendb(OfferID)%>"<%Sendb(IIf(TimeLocked, " checked=""checked""", "")) %> onclick="javascript:updateLocked('lockTime<%Sendb(OfferID)%>', this.checked);" />
                  <input type="hidden" id="conType8" name="conType" value="Time" />
                  <input type="hidden" id="conTime<%Sendb(OfferID)%>" name="con" value="<%Sendb(OfferID)%>" />
                  <input type="hidden" id="lockTime<%Sendb(OfferID)%>" name="locked" value="<%Sendb(IIf(TimeLocked, "1", "0"))%>" />
                  <%
                  Send("  </td>")
                ElseIf (FromTemplate) Then
                  Send("  <td class=""templine"">")
                  Send("    " & IIf(TimeLocked, "Yes", "No"))
                  Send("  </td>")
                End If
                Send("</tr>")
              End If
            %>

            <!-- TENDER CONDITIONS -->
            <%
              t = 1
              MyCommon.QueryStr = "select ExcludedTender from CPE_RewardOptions where RewardOptionID=" & roid
              dt2 = MyCommon.LRT_Select()
              If dt2.Rows.Count > 0 Then
                If dt2.Rows(0).Item("ExcludedTender") = 1 Then TenderExcluded = True
                If dt2.Rows(0).Item("ExcludedTender") = 0 Then TenderExcluded = False
              End If
              MyCommon.QueryStr = "Select ITT.IncentiveTenderID, ITT.TenderTypeID, TT.Name, Value, DisallowEdit, RequiredFromTemplate, ITT.RewardOptionID, RO.ExcludedTender, RO.ExcludedTenderAmtRequired " & _
                                  "from CPE_IncentiveTenderTypes as ITT with (NoLock) " & _
                                  "inner join CPE_RewardOptions as RO with (NoLock) on RO.RewardOptionID=ITT.RewardOptionID " & _
                                  "inner join CPE_TenderTypes as TT with (NoLock) on TT.TenderTypeID=ITT.TenderTypeID " & _
                                  "where ITT.RewardOptionID=" & roid & ";"
              rst = MyCommon.LRT_Select
              If (rst.Rows.Count > 0) Then
                If rst.Rows(0).Item("ExcludedTender") = True Then
                  Send("<tr class=""shadeddark"">")
                  Send("  <td colspan=""" & 4 & """>")
                  Send("    <h3>")
                  Send("      " & Copient.PhraseLib.Lookup("term.tenderconditions", LanguageID))
                  Send("    </h3>")
                  Send("  <td colspan=""" & TierLevels & """>")
                  Send("    <h3>")
                  Send("      " & Copient.PhraseLib.Lookup("term.value", LanguageID))
                  Send("    </h3>")
                  Send("  </td>")
               
                  If (isTemplate Or FromTemplate) Then
                    Send("<td></td>")
                  End If
                  Send("</tr>")
                  
                  For Each row In rst.Rows
                    isTender = True
                    'TenderList &= MyCommon.NZ(row.Item("Name"), "") & "<br />"
                    'TenderValue &= FormatCurrency(MyCommon.NZ(row.Item("Value"), 0), 3) & " " & StrConv(Copient.PhraseLib.Lookup("term.required", LanguageID), VbStrConv.Lowercase) & "<br />"
                    TenderList = MyCommon.NZ(row.Item("Name"), "") & "<br />"
                    TenderValue = FormatCurrency(MyCommon.NZ(row.Item("Value"), 0), 3) & " " & StrConv(Copient.PhraseLib.Lookup("term.required", LanguageID), VbStrConv.Lowercase) & "<br />"
                    TenderDisallowEdit = MyCommon.NZ(row.Item("DisallowEdit"), True)
                    TenderRequired = MyCommon.NZ(row.Item("RequiredFromTemplate"), False)
                    TenderExcluded = MyCommon.NZ(row.Item("ExcludedTender"), False)
                    TenderExcludedAmt = MyCommon.NZ(row.Item("ExcludedTenderAmtRequired"), 0)
                
                    Send("<tr class=""shaded"">")
                    Send("  <td>")
                    If (Logix.UserRoles.EditOffer And Not (FromTemplate And TenderDisallowEdit And TenderRequired)) Then
                      If (TenderRequired And Not isTemplate) OrElse (TenderDisallowEdit And Not isTemplate) Then
                        Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('/logix/CAM/CAM-offer-con.aspx?mode=Delete&Option=Tender&OfferID=" & OfferID & "&IncentiveTenderID=" & row.Item("IncentiveTenderID") & "')}"" value=""X"" />")
                      Else
                        Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('/logix/CAM/CAM-offer-con.aspx?mode=Delete&Option=Tender&OfferID=" & OfferID & "&IncentiveTenderID=" & row.Item("IncentiveTenderID") & "')}"" value=""X"" />")
                      End If
                    ElseIf (Logix.UserRoles.EditTemplates) Then
                      Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('/logix/CAM/CAM-offer-con.aspx?mode=Delete&Option=Tender&OfferID=" & OfferID & "&IncentiveTenderID=" & row.Item("IncentiveTenderID") & "')}"" value=""X"" />")
                    Else
                      Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('/logix/CAM/CAM-offer-con.aspx?mode=Delete&Option=Tender&OfferID=" & OfferID & "&IncentiveTenderID=" & row.Item("IncentiveTenderID") & "')}"" value=""X"" />")
                    End If
                    Send("  </td>")
                    Send("  <td>")
                    Send("  </td>")
                    Send("  <td>")
                    Send("    <a href=""javascript:openPopup('/logix/CPEoffer-con-tender.aspx?OfferID=" & OfferID & "&IncentiveTenderID=" & row.Item("IncentiveTenderID") & "')"">" & Copient.PhraseLib.Lookup("term.tender", LanguageID) & "</a>")
                    Send("  </td>")
                    Send("  <td>")
                    If (isTender) Then
                      If TenderExcluded Then
                        Sendb(Copient.PhraseLib.Lookup("term.allbut", LanguageID) & ":<br />")
                      End If
                      Sendb("<a href=""/logix/tender-engines.aspx"">" & MyCommon.SplitNonSpacedString(TenderList, 25) & "</a>")
                    ElseIf (Not isTender AndAlso TenderRequired) Then
                      Sendb(Copient.PhraseLib.Lookup("term.undefined", LanguageID))
                      Sendb(" <span class=""red"">")
                      Sendb("(" & Copient.PhraseLib.Lookup("term.requiredbytemplate", LanguageID) & ")")
                      Sendb("</span>")
                    End If
                    Send("  </td>")
                    Send(" <td colspan=""" & TierLevels & """>")
                    Send(FormatCurrency(MyCommon.NZ(row.Item("ExcludedTenderAmtRequired"), 0), 3))
                    Send(" </td>")
                    If (isTemplate) Then
                      Send("  <td class=""templine"">")
                      %>
                      <input type="checkbox" id="chkLocked7" name="chkLocked" value="<%Sendb(roid)%>"<%Sendb(IIf(MyCommon.NZ(row.Item("DisallowEdit"), False)=True, " checked=""checked""", "")) %> onclick="javascript:updateLocked('lockTT<%Sendb(MyCommon.NZ(row.Item("RewardOptionID"), 0))%>', this.checked);"<%Sendb(IIf(IsDBNull(row.Item("TenderTypeID")) AndAlso MyCommon.NZ(row.Item("RequiredFromTemplate"), False), " disabled=""disabled""", "")) %> />
                      <input type="hidden" id="conType7" name="conType" value="Tender" />
                      <input type="hidden" id="conTT<%Sendb(roid)%>" name="con" value="<%Sendb(roid)%>" />
                      <input type="hidden" id="lockTT<%Sendb(roid)%>" name="locked" value="<%Sendb(IIf(TenderDisallowEdit, "1", "0"))%>" />
                      <%
                      Send("  </td>")
                      ElseIf (FromTemplate) Then
                        Send("  <td class=""templine"">")
                        Send("    " & Copient.PhraseLib.Lookup("term." & IIf(TenderDisallowEdit, "yes", "no"), LanguageID))
                        Send("  </td>")
                      End If
                      Send("</tr>")
                    Next
                  Else
                    Send("<tr class=""shadeddark"">")
                    Send("  <td colspan=""" & 4 & """>")
                    Send("    <h3>")
                    Send("      " & Copient.PhraseLib.Lookup("term.tenderconditions", LanguageID))
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
                      Send("<td></td>")
                    End If
                    Send("</tr>")
                    For Each row In rst.Rows
                      isTender = True
                      'TenderList &= MyCommon.NZ(row.Item("Name"), "") & "<br />"
                      'TenderValue &= FormatCurrency(MyCommon.NZ(row.Item("Value"), 0), 3) & " " & StrConv(Copient.PhraseLib.Lookup("term.required", LanguageID), VbStrConv.Lowercase) & "<br />"
                      TenderList = MyCommon.NZ(row.Item("Name"), "") & "<br />"
                      TenderValue = FormatCurrency(MyCommon.NZ(row.Item("Value"), 0), 3) & " " & StrConv(Copient.PhraseLib.Lookup("term.required", LanguageID), VbStrConv.Lowercase) & "<br />"
                      TenderDisallowEdit = MyCommon.NZ(row.Item("DisallowEdit"), True)
                      TenderRequired = MyCommon.NZ(row.Item("RequiredFromTemplate"), False)
                      TenderExcluded = MyCommon.NZ(row.Item("ExcludedTender"), False)
                      TenderExcludedAmt = MyCommon.NZ(row.Item("ExcludedTenderAmtRequired"), 0)
                
                      Send("<tr class=""shaded"">")
                      Send("  <td>")
                      If (Logix.UserRoles.EditOffer And Not (FromTemplate And TenderDisallowEdit And TenderRequired)) Then
                        If (TenderRequired And Not isTemplate) OrElse (TenderDisallowEdit And Not isTemplate) Then
                          Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('/logix/CAM/CAM-offer-con.aspx?mode=Delete&Option=Tender&OfferID=" & OfferID & "&IncentiveTenderID=" & row.Item("IncentiveTenderID") & "')}"" value=""X"" />")
                        Else
                          Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('/logix/CAM/CAM-offer-con.aspx?mode=Delete&Option=Tender&OfferID=" & OfferID & "&IncentiveTenderID=" & row.Item("IncentiveTenderID") & "')}"" value=""X"" />")
                        End If
                      ElseIf (Logix.UserRoles.EditTemplates) Then
                        Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('/logix/CAM/CAM-offer-con.aspx?mode=Delete&Option=Tender&OfferID=" & OfferID & "&IncentiveTenderID=" & row.Item("IncentiveTenderID") & "')}"" value=""X"" />")
                      Else
                        Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('/logix/CAM/CAM-offer-con.aspx?mode=Delete&Option=Tender&OfferID=" & OfferID & "&IncentiveTenderID=" & row.Item("IncentiveTenderID") & "')}"" value=""X"" />")
                      End If
                      Send("  </td>")
                      Send("  <td>")
                      Send("  </td>")
                      Send("  <td>")
                      Send("    <a href=""javascript:openPopup('/logix/CPEoffer-con-tender.aspx?OfferID=" & OfferID & "&IncentiveTenderID=" & row.Item("IncentiveTenderID") & "')"">" & Copient.PhraseLib.Lookup("term.tender", LanguageID) & "</a>")
                      Send("  </td>")
                      Send("  <td>")
                      If (isTender) Then
                        If TenderExcluded Then
                          Sendb(Copient.PhraseLib.Lookup("term.allbut", LanguageID) & ":<br />")
                        End If
                        Sendb("<a href=""/logix/tender-engines.aspx"">" & MyCommon.SplitNonSpacedString(TenderList, 25) & "</a>")
                      ElseIf (Not isTender AndAlso TenderRequired) Then
                        Sendb(Copient.PhraseLib.Lookup("term.undefined", LanguageID))
                        Sendb(" <span class=""red"">")
                        Sendb("(" & Copient.PhraseLib.Lookup("term.requiredbytemplate", LanguageID) & ")")
                        Sendb("</span>")
                      End If
                      Send("  </td>")
                      ' Find the per-tier values:
                      t = 1
                      MyCommon.QueryStr = "select IncentiveTenderID, TierLevel, Value from CPE_IncentiveTenderTypeTiers as ITTT with (NoLock) " & _
                                          "where RewardOptionID=" & roid & " and IncentiveTenderID=" & row.Item("IncentiveTenderID") & ";"
                      rst2 = MyCommon.LRT_Select
                      If rst2.Rows.Count = 0 Then
                        Send("  <td colspan=""" & TierLevels & """>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                      Else
                        While t <= TierLevels
                          If t > rst2.Rows.Count Then
                            Send("  <td>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                          Else
                            TenderValue = FormatCurrency(MyCommon.Extract_Val(rst2.Rows(t - 1).Item("Value")), 3)
                            Send("  <td>")
                            If TenderExcluded Then
                              Sendb(FormatCurrency(TenderExcludedAmt, 3) & " " & StrConv(Copient.PhraseLib.Lookup("term.required", LanguageID), VbStrConv.Lowercase))
                            Else
                              Sendb(TenderValue)
                            End If
                            Send("  </td>")
                          End If
                          t += 1
                        End While
                      End If
                      If (isTemplate) Then
                        Send("  <td class=""templine"">")
                  %>
                  <input type="checkbox" id="chkLocked7" name="chkLocked" value="<%Sendb(MyCommon.NZ(row.Item("IncentiveTenderID"), 0))%>"<%Sendb(IIf(MyCommon.NZ(row.Item("DisallowEdit"), False)=True, " checked=""checked""", "")) %> onclick="javascript:updateLocked('lockTT<%Sendb(MyCommon.NZ(row.Item("IncentiveTenderID"), 0))%>', this.checked);"<%Sendb(IIf(IsDBNull(row.Item("TenderTypeID")) AndAlso MyCommon.NZ(row.Item("RequiredFromTemplate"), False), " disabled=""disabled""", "")) %> />
                  <input type="hidden" id="conType7" name="conType" value="Tender" />
                  <input type="hidden" id="conTT<%Sendb(MyCommon.NZ(row.Item("IncentiveTenderID"), 0))%>" name="con" value="<%Sendb(MyCommon.NZ(row.Item("IncentiveTenderID"), 0))%>" />
                  <input type="hidden" id="lockTT<%Sendb(MyCommon.NZ(row.Item("IncentiveTenderID"), 0))%>" name="locked" value="<%Sendb(IIf(TenderDisallowEdit, "1", "0"))%>" />
                  <%
                    Send("  </td>")
                  ElseIf (FromTemplate) Then
                    Send("  <td class=""templine"">")
                    Send("    " & Copient.PhraseLib.Lookup("term." & IIf(TenderDisallowEdit, "yes", "no"), LanguageID))
                    Send("  </td>")
                  End If
                Next
                Send("</tr>")
              End If
            End If
            %>
            
            <!-- INSTANT WIN CONDITIONS -->
            <%
              t = 1
              MyCommon.QueryStr = "select IncentiveInstantWinID, NumPrizesAllowed, OddsOfWinning, RandomWinners, DisallowEdit, RequiredFromTemplate " & _
                                  "from CPE_IncentiveInstantWin as IWW with (NoLock) " & _
                                  "where RewardOptionID=" & roid & ";"
              rst = MyCommon.LRT_Select
              If (rst.Rows.Count > 0) Then
                Send("<tr class=""shadeddark"">")
                Send("  <td colspan=""4"">")
                Send("    <h3>")
                Send("      " & Copient.PhraseLib.Lookup("term.instantwinconditions", LanguageID))
                Send("    </h3>")
                Send("  </td>")
                Send("  <td colspan=""" & TierLevels & """>")
                Send("  </td>")
                If (isTemplate Or FromTemplate) Then
                  Send("<td></td>")
                End If
                Send("</tr>")
              End If
              For Each row In rst.Rows
                isInstantWin = True
                Send("<tr class=""shaded"">")
                Send("  <td>")
                If (Logix.UserRoles.EditOffer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False))) Then
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('/logix/CAM/CAM-offer-con.aspx?mode=Delete&Option=InstantWin&OfferID=" & OfferID & "')}"" value=""X"" />")
                ElseIf (Logix.UserRoles.EditTemplates) Then
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('/logix/CAM/CAM-offer-con.aspx?mode=Delete&Option=InstantWin&OfferID=" & OfferID & "')}"" value=""X"" />")
                Else
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('/logix/CAM/CAM-offer-con.aspx?mode=Delete&Option=InstantWin&OfferID=" & OfferID & "')}"" value=""X"" />")
                End If
                Send("  </td>")
                Send("  <td>")
                Send("  </td>")
                Send("  <td>")
                Send("    <a href=""javascript:openPopup('/logix/CPEoffer-con-instantwin.aspx?OfferID=" & OfferID & "')"">" & Copient.PhraseLib.Lookup("term.instantwin", LanguageID) & "</a>")
                Send("  </td>")
                Send("  <td>")
                If MyCommon.NZ(row.Item("RandomWinners"), False) Then
                  Send(Copient.PhraseLib.Lookup("term.random", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.odds", LanguageID), VbStrConv.Lowercase))
                Else
                  Send(Copient.PhraseLib.Lookup("term.fixed", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.odds", LanguageID), VbStrConv.Lowercase))
                End If
                Send("  </td>")
                Send("  <td colspan=""" & TierLevels & """>")
                Send("    1:" & MyCommon.NZ(row.Item("OddsOfWinning"), "?") & " " & StrConv(Copient.PhraseLib.Lookup("term.on", LanguageID), VbStrConv.Lowercase) & " " & MyCommon.NZ(row.Item("NumPrizesAllowed"), "?") & " " & StrConv(Copient.PhraseLib.Lookup("term.prizes", LanguageID), VbStrConv.Lowercase))
                Send("  </td>")
                If (isTemplate) Then
                  Send("  <td class=""templine"">")
                  %>
                  <input type="checkbox" id="chkLocked9" name="chkLocked" value="<%Sendb(MyCommon.NZ(row.Item("IncentiveInstantWinID"), 0))%>"<%Sendb(IIf(MyCommon.NZ(row.Item("DisallowEdit"), False)=True, " checked=""checked""", "")) %> onclick="javascript:updateLocked('lockIW<%Sendb(MyCommon.NZ(row.Item("IncentiveInstantWinID"), 0))%>', this.checked);"<%Sendb(IIf(MyCommon.NZ(row.Item("RequiredFromTemplate"), False), " disabled=""disabled""", "")) %> />
                  <input type="hidden" id="conType9" name="conType" value="InstantWin" />
                  <input type="hidden" id="conIW<%Sendb(MyCommon.NZ(row.Item("IncentiveInstantWinID"), 0))%>" name="con" value="<%Sendb(MyCommon.NZ(row.Item("IncentiveInstantWinID"), 0))%>" />
                  <input type="hidden" id="lockIW<%Sendb(MyCommon.NZ(row.Item("IncentiveInstantWinID"), 0))%>" name="locked" value="<%Sendb(IIf(MyCommon.NZ(row.Item("DisallowEdit"), False)=True, "1", "0"))%>" />
                  <%
                  Send("  </td>")
                  ElseIf (FromTemplate) Then
                    Send("  <td class=""templine"">")
                    Send("    " & Copient.PhraseLib.Lookup("term." & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "yes", "no"), LanguageID))
                    Send("  </td>")
                End If
                Send("</tr>")
              Next
            %>
            
            <!-- TRIGGER CODE (aka PLU) CONDITIONS -->
            <%
              t = 1
              MyCommon.QueryStr = "select IncentivePLUID, PLU, PerRedemption, CashierMessage, DisallowEdit, RequiredFromTemplate " & _
                                  "from CPE_IncentivePLUs as CIP with (NoLock) " & _
                                  "where RewardOptionID=" & roid & " order by IncentivePLUID;"
              rst = MyCommon.LRT_Select
              If (rst.Rows.Count > 0) Then
                Send("<tr class=""shadeddark"">")
                Send("  <td colspan=""4"">")
                Send("    <h3>")
                Send("      " & Copient.PhraseLib.Lookup("term.triggercodeconditions", LanguageID))
                Send("    </h3>")
                Send("  </td>")
                Send("  <td colspan=""" & TierLevels & """>")
                Send("  </td>")
                If (isTemplate Or FromTemplate) Then
                  Send("<td></td>")
                End If
                Send("</tr>")
              End If
              i = 1
              For Each row In rst.Rows
                isPLU = True
                Send("<tr class=""shaded"">")
                Send("  <td>")
                If (Logix.UserRoles.EditOffer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False))) Then
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CAM-offer-con.aspx?mode=Delete&amp;Option=PLU&amp;OfferID=" & OfferID & "&amp;IncentivePLUID=" & MyCommon.NZ(row.Item("IncentivePLUID"), 0) & "')}"" value=""X"" />")
                ElseIf (Logix.UserRoles.EditTemplates) Then
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CAM-offer-con.aspx?mode=Delete&amp;Option=PLU&amp;OfferID=" & OfferID & "&amp;IncentivePLUID=" & MyCommon.NZ(row.Item("IncentivePLUID"), 0) & "')}"" value=""X"" />")
                Else
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CAM-offer-con.aspx?mode=Delete&amp;Option=PLU&amp;OfferID=" & OfferID & "&amp;IncentivePLUID=" & MyCommon.NZ(row.Item("IncentivePLUID"), 0) & "')}"" value=""X"" />")
                End If
                Send("  </td>")
                Send("  <td>")
                If i > 1 Then
                  Send("    " & Copient.PhraseLib.Lookup("term.or", LanguageID))
                End If
                Send("  </td>")
                Send("  <td>")
                Send("    <a href=""javascript:openPopup('/logix/CPEoffer-con-plu.aspx?OfferID=" & OfferID & "&amp;IncentivePLUID=" & MyCommon.NZ(row.Item("IncentivePLUID"), 0) & "')"">" & Copient.PhraseLib.Lookup("term.triggercode", LanguageID) & "</a>")
                Send("  </td>")
                Send("  <td>")
                If MyCommon.NZ(row.Item("PLU"), "") = "" Then
                  Send("    " & Copient.PhraseLib.Lookup("term.undefined", LanguageID))
                Else
                  Send("    " & MyCommon.NZ(row.Item("PLU"), ""))
                End If
                Send("  </td>")
                Send("  <td colspan=""" & TierLevels & """>")
                Send("    " & Copient.PhraseLib.Lookup(IIf(MyCommon.NZ(row.Item("PerRedemption"), False), "term.OncePerRedemption", "term.oncepertransaction"), LanguageID))
                Send("  </td>")
                If (isTemplate) Then
                  Send("  <td class=""templine"">")
                    %>
                    <input type="checkbox" id="chkLocked11" name="chkLocked" value="<%Sendb(MyCommon.NZ(row.Item("IncentivePLUID"), 0))%>"<%Sendb(IIf(MyCommon.NZ(row.Item("DisallowEdit"), False)=True, " checked=""checked""", "")) %> onclick="javascript:updateLocked('lockPLU<%Sendb(MyCommon.NZ(row.Item("IncentivePLUID"), 0))%>', this.checked);"<%Sendb(IIf(MyCommon.NZ(row.Item("RequiredFromTemplate"), False), " disabled=""disabled""", "")) %> />
                    <input type="hidden" id="conType11" name="conType" value="PLU" />
                    <input type="hidden" id="conPLU<%Sendb(MyCommon.NZ(row.Item("IncentivePLUID"), 0))%>" name="con" value="<%Sendb(MyCommon.NZ(row.Item("IncentivePLUID"), 0))%>" />
                    <input type="hidden" id="lockPLU<%Sendb(MyCommon.NZ(row.Item("IncentivePLUID"), 0))%>" name="locked" value="<%Sendb(IIf(MyCommon.NZ(row.Item("DisallowEdit"), False)=True, "1", "0"))%>" />
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
          </tbody>
        </table>
        <hr class="hidden" />
      </div>
      
      <div class="box" id="newcondition">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("offer-con.addcondition", LanguageID))%>
          </span>
        </h2>
        <%
          'First set the TenderWorthy variable, which determines if the offer is eligible to use tender conditions
          MyCommon.QueryStr = "select RO.IncentiveID from CPE_Deliverables as D with (NoLock) " & _
                              "left join CPE_RewardOptions as RO with (NoLock) on RO.RewardOptionID=D.RewardOptionID " & _
                              "where DeliverableTypeID=2 and IncentiveID=" & OfferID & " and RO.Deleted=0;"
          rst = MyCommon.LRT_Select
          If rst.Rows.Count = 0 Then
            MyCommon.QueryStr = "select TenderTypeID from CPE_TenderTypes with (NoLock) where TenderTypeID not in " & _
                                "(select TenderTypeID from CPE_IncentiveTenderTypes where RewardOptionID=" & roid & ");"
            rst = MyCommon.LRT_Select()
            If rst.Rows.Count > 0 Then
              TenderWorthy = True
            End If
          End If
          
          If IsFooterOffer AndAlso isCustomer Then
            Send(Copient.PhraseLib.Lookup("ueoffer-con.FooterPrintedMessage", LanguageID))
          Else
            If (isTemplate) Then
              Send("<span class=""temp"">")
              Send("  <input type=""checkbox"" class=""tempcheck"" id=""Disallow_Conditions"" name=""Disallow_Conditions""" & IIf(Disallow_Conditions, " checked=""checked""", "") & " />")
              Send("  <label for=""Disallow_Conditions"">" & Copient.PhraseLib.Lookup("term.locked", LanguageID) & "</label>")
              Send("</span>")
            End If
            MyCommon.QueryStr = "SELECT PECT.EngineID, PECT.ComponentTypeID, CT.ConditionTypeID, CT.Description, CT.PhraseID, PECT.Singular, " & _
                                "  CASE ConditionTypeID " & _
                                "    WHEN 1 THEN (SELECT COUNT(*) FROM CPE_IncentiveCustomerGroups WITH (NOLOCK) where RewardOptionID=" & roid & ") " & _
                                "    WHEN 2 THEN (SELECT COUNT(*) FROM CPE_IncentiveProductGroups WITH (NOLOCK) where RewardOptionID=" & roid & " and Deleted=0 and Disqualifier=0) " & _
                                "    WHEN 3 THEN (SELECT COUNT(*) FROM CPE_IncentivePointsGroups WITH (NOLOCK) where RewardOptionID=" & roid & ") " & _
                                "    WHEN 4 THEN (SELECT COUNT(*) FROM CPE_IncentiveStoredValuePrograms WITH (NOLOCK) where RewardOptionID=" & roid & ") " & _
                                "    WHEN 5 THEN (SELECT COUNT(*) FROM CPE_IncentiveTenderTypes WITH (NOLOCK) where RewardOptionID=" & roid & ") " & _
                                "    WHEN 6 THEN (SELECT COUNT(*) FROM CPE_IncentiveDOW WITH (NOLOCK) where IncentiveID=" & OfferID & " and Deleted=0) " & _
                                "    WHEN 7 THEN (SELECT COUNT(*) FROM CPE_IncentiveTOD WITH (NOLOCK) where IncentiveID=" & OfferID & ") " & _
                                "    WHEN 8 THEN (SELECT COUNT(*) FROM CPE_IncentiveInstantWin WITH (NOLOCK) where RewardOptionID=" & roid & " and Deleted=0) " & _
                                "    WHEN 9 THEN (SELECT COUNT(*) FROM CPE_IncentivePLUs WITH (NOLOCK) where RewardOptionID=" & roid & ") " & _
                                "    WHEN 10 THEN (SELECT COUNT(*) FROM CPE_IncentiveProductGroups WITH (NOLOCK) where RewardOptionID=" & roid & " and Deleted=0 and Disqualifier=1) " & _
                                "    ELSE 0 " & _
                                "  END as ItemCount " & _
                                "FROM PromoEngineComponentTypes AS PECT " & _
                                "INNER JOIN CPE_ConditionTypes AS CT ON CT.ConditionTypeID=PECT.LinkID " & _
                                "WHERE EngineID=6 AND PECT.ComponentTypeID=1 AND Enabled=1"
            'Impose a few special limits on the query based on various in-page factors:
            If (Not isCustomer) OrElse (IsFooterOffer AndAlso Not isCustomer) Then
              'Offer has no customer condition, so limit it just to that
              MyCommon.QueryStr &= " AND CT.ConditionTypeID=1"
            End If
            If (AccumEnabled) Then
              'Accumulation is on, so no more product conditions
              MyCommon.QueryStr &= " AND CT.ConditionTypeID<>2"
            End If
            If (Not TenderWorthy) Then
              MyCommon.QueryStr &= " AND CT.ConditionTypeID<>5"
            End If
            If (TierLevels > 1) Then
              'Offer is multitiered, so instant win is invalid
              MyCommon.QueryStr &= " AND CT.ConditionTypeID<>8"
            End If
            If (Not isProduct) OrElse (isProduct AndAlso AccumEnabled) Then
              'Offer has no product condition (or has one with accumulation), so disallow product disqualifiers
              MyCommon.QueryStr &= " AND CT.ConditionTypeID<>10"
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
              If isTemplate OrElse (Not (isCustomer And isProduct And isProductDisqualifier And isDay And isTime And isTender) And Not Disallow_Conditions) Then
              Else
                Sendb(" disabled=""disabled""")
              End If
              Sendb(" />")
            End If
          End If
          Send("<br />")
        %>
      </div>
    </div>
    <br clear="all" />
  </div>
</form>
<!-- #Include virtual="/include/graphic-reward.inc" -->
<%If (isProduct OrElse isPoint OrElse isDay OrElse isStoredValue) Then%>
<script type="text/javascript">
  var elemCustDelBtn = document.getElementById("customerDelete");
  
  if (elemCustDelBtn != null) {
      elemCustDelBtn.disabled = true;
  }
</script>
<%End If%>
<%If (isCustomer OrElse isProduct OrElse isPoint OrElse isDay OrElse isStoredValue) Then%>
<%Else%>
<script type="text/javascript">
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
