<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%-- version:7.3.1.138972 --%>
<%@ Register TagPrefix="uc" TagName="ucOptInCondition" Src="~/logix/UserControls/OfferEligibilityConditions.ascx" %>
<%@ Import Namespace="System.Data" %>

<%
  ' *****************************************************************************
  ' * FILENAME: CPEoffer-con.aspx 
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
  Dim MyCPEOffer As New Copient.EIW
  Dim Logix As New Copient.LogixInc
  Dim MyCpe As New Copient.CPEOffer
  Dim rst As DataTable
  Dim row As DataRow
  Dim rst2 As DataTable
  Dim row2 As DataRow
  Dim rst3 As DataTable
  Dim row3 As DataRow
  Dim dt, dt2 As DataTable
  Dim OfferID As Long
  Dim ConditionID As Long
  Dim Name As String = ""
  Dim isTemplate As Boolean = False
  Dim FromTemplate As Boolean = False
  Dim Disallow_Conditions As Boolean = False
  Dim Disallow_OptIn As Boolean = False
  Dim IsTemplateVal As String = "Not"
  Dim ActiveSubTab As Integer = 91
  Dim roid As Integer = 0
  Dim i As Integer
  Dim Days As String = ""
  Dim Times As String = ""
  Dim isCustomer As Boolean = False
  Dim isAttribute As Boolean = False
  Dim isTargeted As Boolean = False
  'isTargeted indicates if customer or attribute conditions (one or the other) are set.
  Dim isProduct As Boolean = False
  Dim isProductDisqualifier As Boolean = False
  Dim isPoint As Boolean = False
  Dim isDay As Boolean = False
  Dim isTime As Boolean = False
  Dim isStoredValue As Boolean = False
  Dim isTender As Boolean = False
  Dim isInstantWin As Boolean = False
  Dim isPLU As Boolean = False
  Dim isEIW As Boolean = False
  Dim isCardType As Boolean = False
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
  Dim DisqualifierID As Integer = 0
  Dim ExcludedIncentiveID As Integer = 0
  Dim ProductConditions As Integer = 0
  Dim ProductCombo As Integer = 0
  Dim PreferenceCombo As Integer = 1
  Dim ExcludedProductGroupID As Integer = 0
  Dim ExcludedProductGroupName As String = ""
  Dim AccumEnabled As Boolean = False
  Dim IncentiveTenderID As Integer = 0
  Dim IncentiveEIWID As Integer = 0
  Dim IncentiveAttributeID As Integer = 0
  Dim IsFooterOffer As Boolean = False
  Dim PRoductCount As Integer = 0
  Dim EngineID As Integer = 2
  Dim EngineSubTypeID As Integer = 0
  Dim CustomerConditions As Integer = 0
  Dim AttributeConditions As Integer = 0
  Dim AttributeCombo As Integer = 0
  Dim HasEIW As Boolean = False
  Dim HasDiscount As Boolean = False
  Dim HasBundleDiscount As Boolean = False
  Dim ProductConditionCount As Integer = 0
  Dim TempQueryStr As String
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "CPEoffer-con.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  
  If MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER) Then
    MyCommon.Open_PrefManRT()
  End If
  
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  
  OfferID = Request.QueryString("OfferID")
  ConditionID = Request.QueryString("ConditionID")
  
  If (OfferID = 0) Then
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "CPEoffer-gen.aspx")
  End If
  
  MyCommon.QueryStr = "select RewardOptionID, TierLevels, ProductComboID from CPE_RewardOptions with (NoLock) " & _
                      "where IncentiveID=" & OfferID & " and TouchResponse=0 and Deleted=0;"
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    roid = rst.Rows(0).Item("RewardOptionID")
    TierLevels = rst.Rows(0).Item("TierLevels")
    ProductCombo = rst.Rows(0).Item("ProductComboID")
  End If
  
  IsFooterOffer = MyCpe.IsFooterOffer(OfferID)

  'Determine if the offer has an enterprise instant win condition
  MyCommon.QueryStr = "select IncentiveEIWID from CPE_IncentiveEIW with (NoLock) where RewardOptionID=" & roid & ";"
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    HasEIW = True
  End If
  
  'Determine if the offer has a group-level conditional (bundle) discount
  MyCommon.QueryStr = "select DiscountID, DiscountTypeID from CPE_Discounts as DIS with (NoLock) " & _
                      "inner join CPE_Deliverables as DEL on DEL.OutputID=DIS.DiscountID " & _
                      "where DEL.RewardOptionID=" & roid & " and DEL.DeliverableTypeID=2 and DEL.Deleted=0 and DIS.Deleted=0;"
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    HasDiscount = True
    If MyCommon.NZ(rst.Rows(0).Item("DiscountTypeID"), 0) = 4 Then
      HasBundleDiscount = True
    End If
  End If
  
  'Determine the number of associated product conditions
  MyCommon.QueryStr = "select IPG.IncentiveProductGroupID from CPE_IncentiveProductGroups as IPG with (NoLock) " & _
                      " left join CPE_RewardOptions as RO with (NoLock) on IPG.RewardOptionID=RO.RewardOptionID " & _
                      " where IPG.RewardOptionID=" & roid & "and IPG.Deleted=0;"
  rst = MyCommon.LRT_Select
  ProductConditionCount = rst.Rows.Count
  
  MyCommon.QueryStr = "select AccumMin from CPE_IncentiveProductGroups where RewardOptionID=" & roid & " and ExcludedProducts=0 and Deleted=0;"
  rst = MyCommon.LRT_Select()
  If rst.Rows.Count > 0 Then
    If rst.Rows.Count = 1 And MyCommon.NZ(rst.Rows(0).Item("AccumMin"), 0) > 0 Then
      AccumEnabled = True
    End If
  End If
  
  If Request.QueryString("IncentiveTenderID") <> "" Then
    IncentiveTenderID = MyCommon.Extract_Val(Request.QueryString("IncentiveTenderID"))
  End If
  If Request.QueryString("IncentiveEIWID") <> "" Then
    IncentiveEIWID = MyCommon.Extract_Val(Request.QueryString("IncentiveEIWID"))
  End If
  If Request.QueryString("IncentiveAttributeID") <> "" Then
    IncentiveAttributeID = MyCommon.Extract_Val(Request.QueryString("IncentiveAttributeID"))
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
function submitform() {
  document.getElementById('save1').click();
}
</script>
<%
  Send_HeadEnd()
  
  ' handle adding stuff on
  If (Request.QueryString("Save") = "" And Request.QueryString("newconglobal") <> "") Then
    If (Request.QueryString("newconglobal") = 1) Then
      Send("<script type=""text/javascript"">openPopup('CPEoffer-con-customer.aspx?OfferID=" & OfferID & "')</script>")
    ElseIf (Request.QueryString("newconglobal") = 2) Then
      Send("<script type=""text/javascript"">openPopup('CPEoffer-con-product.aspx?OfferID=" & OfferID & "&Disqualifier=0')</script>")
    ElseIf (Request.QueryString("newconglobal") = 3) Then
      Send("<script type=""text/javascript"">openPopup('CPEoffer-con-point.aspx?OfferID=" & OfferID & "')</script>")
    ElseIf (Request.QueryString("newconglobal") = 4) Then
      Send("<script type=""text/javascript"">openPopup('CPEoffer-con-sv.aspx?OfferID=" & OfferID & "')</script>")
    ElseIf (Request.QueryString("newconglobal") = 5) Then
      Send("<script type=""text/javascript"">openPopup('CPEoffer-con-tender.aspx?OfferID=" & OfferID & "')</script>")
    ElseIf (Request.QueryString("newconglobal") = 6) Then
      Send("<script type=""text/javascript"">openPopup('CPEoffer-con-day.aspx?OfferID=" & OfferID & "')</script>")
    ElseIf (Request.QueryString("newconglobal") = 7) Then
      Send("<script type=""text/javascript"">openPopup('CPEoffer-con-time.aspx?OfferID=" & OfferID & "')</script>")
    ElseIf (Request.QueryString("newconglobal") = 8) Then
      Send("<script type=""text/javascript"">openPopup('CPEoffer-con-instantwin.aspx?OfferID=" & OfferID & "')</script>")
    ElseIf (Request.QueryString("newconglobal") = 9) Then
      Send("<script type=""text/javascript"">openPopup('CPEoffer-con-plu.aspx?OfferID=" & OfferID & "')</script>")
    ElseIf (Request.QueryString("newconglobal") = 10) Then
      Send("<script type=""text/javascript"">openPopup('CPEoffer-con-product.aspx?OfferID=" & OfferID & "&Disqualifier=1')</script>")
    ElseIf (Request.QueryString("newconglobal") = 11) Then
      Send("<script type=""text/javascript"">openPopup('CPEoffer-con-einstantwin.aspx?OfferID=" & OfferID & "')</script>")
    ElseIf (Request.QueryString("newconglobal") = 12) Then
      Send("<script type=""text/javascript"">openPopup('CPEoffer-con-attribute.aspx?OfferID=" & OfferID & "')</script>")
    ElseIf (Request.QueryString("newconglobal") = 13) Then
      Send("<script type=""text/javascript"">openPopup('CPEoffer-con-cardtype.aspx?OfferID=" & OfferID & "')</script>")
    ElseIf (Request.QueryString("newconglobal") = 14) Then
      Send("<script type=""text/javascript"">openPopup('CPEoffer-con-pref.aspx?OfferID=" & OfferID & "')</script>")
    End If
  ElseIf (Request.QueryString("mode") = "ChangeProductCombo") Then
    If HasBundleDiscount Then
      infoMessage = Copient.PhraseLib.Lookup("CPEoffer-con.GroupLevelConditionalLocked", LanguageID)
    Else
      If (Request.QueryString("pc") <> "") Then
        ProductCombo = MyCommon.Extract_Val(Request.QueryString("pc"))
        'Set the ProductCombo for this offer
        If ProductCombo = 1 Then
          ProductCombo = 2
          MyCommon.QueryStr = "update CPE_RewardOptions with (RowLock) set ProductComboID=" & ProductCombo & " where RewardOptionID=" & roid & ";"
          MyCommon.LRT_Execute()
        ElseIf ProductCombo = 2 Then
          ProductCombo = 1
          MyCommon.QueryStr = "update CPE_RewardOptions with (RowLock) set ProductComboID=" & ProductCombo & " where RewardOptionID=" & roid & ";"
          MyCommon.LRT_Execute()
        End If
        'Change the offer status
        MyCommon.QueryStr = "Update CPE_Incentives with (RowLock) set StatusFlag=1 where IncentiveID=" & OfferID
        MyCommon.LRT_Execute()
      End If
    End If
   ElseIf (Request.QueryString("mode") = "ChangeAttributeCombo") Then
    If (Request.QueryString("ac") <> "") Then
      AttributeCombo = MyCommon.Extract_Val(Request.QueryString("ac"))
      'Set the AttributeCombo for this offer
      If AttributeCombo = 1 Then
        AttributeCombo = 2
        MyCommon.QueryStr = "update CPE_RewardOptions set AttributeComboID=" & AttributeCombo & " where RewardOptionID=" & roid & ";"
        MyCommon.LRT_Execute()
      ElseIf AttributeCombo = 2 Then
        AttributeCombo = 1
        MyCommon.QueryStr = "update CPE_RewardOptions set AttributeComboID=" & AttributeCombo & " where RewardOptionID=" & roid & ";"
        MyCommon.LRT_Execute()
            End If
            'Change the offer status
            MyCommon.QueryStr = "Update CPE_Incentives with (RowLock) set StatusFlag=1 where IncentiveID=" & OfferID
            MyCommon.LRT_Execute()
    End If
  ElseIf (Request.QueryString("mode") = "ChangePreferenceCombo") Then
    If (Request.QueryString("pc") <> "") Then
      PreferenceCombo = MyCommon.Extract_Val(Request.QueryString("pc"))
      'Set the Preference Combo for this offer
      If PreferenceCombo = 1 Then
        PreferenceCombo = 2
        MyCommon.QueryStr = "update CPE_RewardOptions with (RowLock) set PreferenceComboID=" & PreferenceCombo & " where RewardOptionID=" & roid & ";"
        MyCommon.LRT_Execute()
      ElseIf PreferenceCombo = 2 Then
        PreferenceCombo = 1
        MyCommon.QueryStr = "update CPE_RewardOptions with (RowLock) set PreferenceComboID=" & PreferenceCombo & " where RewardOptionID=" & roid & ";"
        MyCommon.LRT_Execute()
      End If
      'Change the offer status
      MyCommon.QueryStr = "Update CPE_Incentives with (RowLock) set StatusFlag=1 where IncentiveID=" & OfferID
      MyCommon.LRT_Execute()
    End If
  ElseIf (Request.QueryString("mode") = "Delete") Then
    If (Request.QueryString("Option") = "Customer") Then
      ' ok someone clicked the X on the customer group stuff lets ditch all the associated groups on this offer
      MyCommon.QueryStr = "update CPE_IncentiveCustomerGroups with (RowLock) set Deleted=1, LastUpdate=getdate(),TCRMAStatusFlag=3 where RewardOptionID=" & roid & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1, AllowOptOut=0 where IncentiveID=" & OfferID & ";"
      MyCommon.LRT_Execute()
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-customer-delete", LanguageID))
    ElseIf (Request.QueryString("Option") = "Product") Then
      If HasBundleDiscount Then
        infoMessage = Copient.PhraseLib.Lookup("CPEoffer-con.GroupLevelConditionalLocked", LanguageID)
      ElseIf HasDiscount And ProductConditionCount = 1 Then
        infoMessage = Copient.PhraseLib.Lookup("CPEoffer-con.DiscountPresent", LanguageID)
      Else
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
        MyCommon.QueryStr = "select IncentiveProductGroupID from CPE_IncentiveProductGroups where RewardOptionID=" & roid & " and Deleted=0"
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
                            "(select IncentiveProductGroupID from CPE_IncentiveProductGroups where RewardOptionID=" & roid & " and Deleted=0) " & _
                            "and RewardOptionID=" & roid & ";"
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
        
          '! Taken       
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
      End If
    ElseIf (Request.QueryString("Option") = "ProductDisqualifier") Then
      ' Someone clicked the X on the product disqualifier stuff
      'Get IncentiveProductGroupID
      MyCommon.QueryStr = "select IncentiveProductGroupID from CPE_IncentiveProductGroups with (NoLock) where Deleted=0 and RewardOptionID=" & roid & " and Disqualifier=1"
      rst = MyCommon.LRT_Select()
      If rst.Rows.Count > 0 Then
        DisqualifierID = MyCommon.Extract_Val(rst.Rows(0).Item("IncentiveProductGroupID"))
      Else
        DisqualifierID = 0
      End If
      'Delete from IncentiveProductGroupTiers
      MyCommon.QueryStr = "Delete from CPE_IncentiveProductGroupTiers where RewardOptionID=" & roid & " and IncentiveProductGroupID=" & DisqualifierID
      MyCommon.LRT_Execute()
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
        MyCommon.QueryStr = "select IncentiveTenderID from CPE_IncentiveTenderTypes with (NoLock) where RewardOptionID=" & roid & ";"
        rst = MyCommon.LRT_Select()
        If rst.Rows.Count = 0 Then
          MyCommon.QueryStr = "update CPE_RewardOptions with (RowLock) set ExcludedTender=0, ExcludedTenderAmtRequired=0 where RewardOptionID=" & roid & ";"
          MyCommon.LRT_Execute()
        End If
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
      ' If the offer has an EIW condition, rerandomize its triggers
      If HasEIW Then
        MyCPEOffer.RandomizeTriggersByOffer(OfferID)
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-instantwin-randomize", LanguageID))
      End If
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
      ' If the offer has an EIW condition, rerandomize its triggers
      If HasEIW Then
        MyCPEOffer.RandomizeTriggersByOffer(OfferID)
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-instantwin-randomize", LanguageID))
      End If
    ElseIf (Request.QueryString("Option") = "InstantWin") Then
      ' Someone clicked the X on a store-level instant win condition
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
      MyCommon.QueryStr = "delete from CPE_IncentivePLUs with (RowLock) where RewardOptionID=" & roid & " and IncentivePLUID=" & IncentiveID & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
      MyCommon.LRT_Execute()
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-plu-delete", LanguageID))
    ElseIf (Request.QueryString("Option") = "EInstantWin") Then
      ' Someone clicked the X on an enterprise instant win condition
      MyCommon.QueryStr = "update CPE_EIWTriggers with (RowLock) set Removed=1, LastUpdate=getdate() where RewardOptionID=" & roid & " and Removed=0;"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "delete from CPE_IncentiveEIW with (RowLock) where RewardOptionID=" & roid & " and IncentiveEIWID=" & IncentiveEIWID & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
      MyCommon.LRT_Execute()
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-einstantwin-delete", LanguageID))
    ElseIf (Request.QueryString("Option") = "Attribute") Then
      ' Someone clicked the X on an attribute condition
      MyCommon.QueryStr = "select IncentiveAttributeID from CPE_IncentiveAttributes where RewardOptionID=" & roid & " and Deleted=0;"
      rst = MyCommon.LRT_Select
      AttributeConditions = rst.Rows.Count
      If AttributeConditions = 2 Then
        'Change the AttributeComboID to single
        MyCommon.QueryStr = "update CPE_RewardOptions set AttributeComboID=0 where RewardOptionID=" & roid & ";"
        MyCommon.LRT_Execute()
      End If
      MyCommon.QueryStr = "delete from CPE_IncentiveAttributes with (RowLock) where RewardOptionID=" & roid & " and IncentiveAttributeID=" & IncentiveAttributeID & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "delete from CPE_IncentiveAttributeTiers with (RowLock) where RewardOptionID=" & roid & " and IncentiveAttributeID=" & IncentiveAttributeID & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1, AllowOptOut=0 where IncentiveID=" & OfferID & ";"
      MyCommon.LRT_Execute()
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-attribute-delete", LanguageID))
    ElseIf (Request.QueryString("Option") = "CardType") Then
      MyCommon.QueryStr = "update CPE_IncentiveCardTypes with (RowLock) set Deleted=1, LastUpdate=getdate() where RewardOptionID=" & roid & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
      MyCommon.LRT_Execute()
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-cardtype-delete", LanguageID))
    ElseIf (Request.QueryString("Option") = "Preference") Then
      MyCommon.QueryStr = "dbo.pt_CPE_IncentivePrefs_Delete"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@IncentivePrefsID", SqlDbType.Int).Value = MyCommon.Extract_Val(Request.QueryString("IncentivePrefsID"))
      MyCommon.LRTsp.ExecuteNonQuery()
      
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
      MyCommon.LRT_Execute()
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-preference-delete", LanguageID))
    End If
    If infoMessage = "" Then
      Response.Status = "301 Moved Permanently"
      Response.AddHeader("Location", "CPEoffer-con.aspx?OfferID=" & OfferID)
      GoTo done
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
      Dim IsOptInPanelLocked = Request.QueryString("IsOptInPanelLocked")
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
              Case "EInstantWin"
                sQuery = "update CPE_IncentiveEIW with (RowLock) set DisallowEdit=" & LockedStatus(LoopCtr) & ", " & _
                          "RequiredFromTemplate=" & IIf(LockedStatus(LoopCtr) = "1", "0", "RequiredFromTemplate") & " " & _
                          "where IncentiveEIWID=" & Conditions(LoopCtr) & ";"
              Case "Attribute"
                sQuery = "update CPE_IncentiveAttributes with (RowLock) set DisallowEdit=" & LockedStatus(LoopCtr) & ", " & _
                          "RequiredFromTemplate=" & IIf(LockedStatus(LoopCtr) = "1", "0", "RequiredFromTemplate") & " " & _
                          "where IncentiveAttributeID=" & Conditions(LoopCtr) & ";"
              Case "CardType"
                sQuery = "update CPE_IncentiveCardTypes with (RowLock) set DisallowEdit=" & LockedStatus(LoopCtr) & ", " & _
                          "RequiredFromTemplate=" & IIf(LockedStatus(LoopCtr) = "1", "0", "RequiredFromTemplate") & " " & _
                          "where IncentiveCardTypeID=" & Conditions(LoopCtr) & ";"
              Case "Preference"
                sQuery = "update CPE_IncentivePrefs with (RowLock) set DisallowEdit=" & LockedStatus(LoopCtr) & ", " & _
                          "RequiredFromTemplate=" & IIf(LockedStatus(LoopCtr) = "1", "0", "RequiredFromTemplate") & " " & _
                          "where IncentivePrefsID=" & Conditions(LoopCtr) & ";"
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
                        "CRMRestricted, StatusFlag, OC.Description as CategoryName, IsTemplate, FromTemplate, EngineSubTypeID " & _
                        "from CPE_Incentives as CPE with (NoLock) " & _
                        "left join OfferCategories as OC with (NoLock) on CPE.PromoClassID=OfferCategoryID " & _
                        "where IncentiveID=" & OfferID & ";"
    rst = MyCommon.LRT_Select
    For Each row In rst.Rows
      Name = MyCommon.NZ(row.Item("Name"), "")
      isTemplate = MyCommon.NZ(row.Item("IsTemplate"), False)
      FromTemplate = MyCommon.NZ(row.Item("FromTemplate"), False)
      StatusFlag = MyCommon.NZ(row.Item("StatusFlag"), 0)
      EngineSubTypeID = MyCommon.NZ(row.Item("EngineSubTypeID"), 0)
    Next
  
    If (isTemplate Or FromTemplate) Then
      ' lets dig the permissions if its a template
      MyCommon.QueryStr = "select Disallow_Conditions,Disallow_Optin from TemplatePermissions with (NoLock) where OfferID=" & OfferID & ";"
      rst = MyCommon.LRT_Select
      If (rst.Rows.Count > 0) Then
        For Each row In rst.Rows
          ' ok there are some rows for the template
        Disallow_Conditions = MyCommon.NZ(row.Item("Disallow_Conditions"), True)
        Disallow_OptIn = CMS.Utilities.NZ(row.Item("Disallow_Optin"), True)
        Next
      End If
    End If
    If Not isTemplate Then
       DeleteBtnDisabled = IIf(Logix.UserRoles.EditOffer, "", " disabled=""disabled""")
    Else
       DeleteBtnDisabled = IIf(Logix.UserRoles.EditTemplates, "", " disabled=""disabled""")
    End If
     
     
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
    Send_Subtabs(Logix, ActiveSubTab, 5, , OfferID)
  
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
    <form runat="server" id="form1">     
    <uc:ucOptInCondition ID="ucOfferEligibilityCondition" runat="server" AppName="UEoffer-con.aspx" />  
    </form>
    <form action="CPEoffer-con.aspx" id="mainform" name="mainform">
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
            <%
              'For the purposes of determining whether to allow customer and attribute conditions to be deleted,
              'check first to see if the offer has any of these condition types.
              MyCommon.QueryStr = "select count(IncentiveCustomerID) as Count from CPE_IncentiveCustomerGroups with (NoLock) where RewardOptionID=" & roid & " and Deleted=0;"
              rst = MyCommon.LRT_Select
              CustomerConditions = MyCommon.NZ(rst.Rows(0).Item("Count"), 0)
              MyCommon.QueryStr = "select count(IncentiveAttributeID) as Count from CPE_IncentiveAttributes with (NoLock) where RewardOptionID=" & roid & " and Deleted=0;"
              rst = MyCommon.LRT_Select
              AttributeConditions = MyCommon.NZ(rst.Rows(0).Item("Count"), 0)
            %>
          
            <!-- CUSTOMER CONDITIONS -->
            <%
              t = 1
              ' Find the currently selected groups on page load
              MyCommon.QueryStr = "select ICG.IncentiveCustomerID, CG.CustomerGroupID, CG.NewCardholders, CG.AnyCAMCardholder, Name, PhraseID, " & _
                                  " ExcludedUsers, DisallowEdit, RequiredFromTemplate from CPE_IncentiveCustomerGroups as ICG with (NoLock) " & _
                                  " left join CustomerGroups as CG with (NoLock) " & _
                                  " on CG.CustomerGroupID=ICG.CustomerGroupID where RewardOptionID=" & roid & _
                                  " and ICG.Deleted=0 order by ExcludedUsers ;"
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
                  Sendb("<input type=""button"" class=""ex"" id=""customerDelete"" name=""customerDelete"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & IIf(AttributeConditions = 0, " disabled=""disabled""", "") & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=Customer&OfferID=" & OfferID & "')}"" value=""X"" />")
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
                If (MyCommon.NZ(row.Item("ExcludedUsers"), False) = True) Then
                  Sendb(StrConv(Copient.PhraseLib.Lookup("term.excluding", LanguageID), VbStrConv.Lowercase) & " ")
                End If
                If (MyCommon.NZ(row.Item("CustomerGroupID"), -1) > 2) AndAlso (MyCommon.NZ(row.Item("NewCardholders"), 0) = 0) AndAlso (MyCommon.NZ(row.Item("AnyCAMCardholder"), 0) = 0) Then
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
            
            <!-- ATTRIBUTE CONDITIONS -->
            <%
              t = 1
              ' Find the currently selected attributes on page load
              MyCommon.QueryStr = "select IA.IncentiveAttributeID, DisallowEdit, RequiredFromTemplate, RO.AttributeComboID " & _
                                  " from CPE_IncentiveAttributes as IA with (NoLock) " & _
                                  " left join CPE_RewardOptions as RO with (NoLock) on IA.RewardOptionID=RO.RewardOptionID " & _
                                  " where IA.RewardOptionID=" & roid & " and IA.Deleted=0;"
              rst = MyCommon.LRT_Select
              If (rst.Rows.Count > 0) Then
                'There's at least one attribute condition; set it as such.
                isAttribute = True
                i = 1
                Send("<tr class=""shadeddark"">")
                Send("  <td colspan=""4"">")
                Send("    <h3>")
                Send("      " & Copient.PhraseLib.Lookup("term.attributeconditions", LanguageID))
                Send("    </h3>")
                Send("  </td>")
                Send("  <td colspan=""" & TierLevels & """>")
                Send("  </td>")
                If (isTemplate Or FromTemplate) Then
                  Send("<td></td>")
                End If
                Send("</tr>")
                For Each row In rst.Rows
                  Send("<tr class=""shaded"">")
                  Send("  <td>")
                  If (Logix.UserRoles.EditOffer And ((CustomerConditions > 0) OrElse (AttributeConditions > 1)) And Not (FromTemplate And Disallow_Conditions)) Then
                    Sendb("<input type=""button"" class=""ex"" id=""attributeDelete"" name=""attributeDelete"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=Attribute&OfferID=" & OfferID & "&IncentiveAttributeID=" & MyCommon.NZ(row.Item("IncentiveAttributeID"), 0) & "')}"" value=""X"" />")
                  ElseIf (Logix.UserRoles.EditTemplates And ((CustomerConditions > 0) OrElse (AttributeConditions > 1))) Then
                    Sendb("<input type=""button"" class=""ex"" id=""attributeDelete"" name=""attributeDelete"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=Attribute&OfferID=" & OfferID & "&IncentiveAttributeID=" & MyCommon.NZ(row.Item("IncentiveAttributeID"), 0) & "')}"" value=""X"" />")
                  Else                  
                    Sendb("<input type=""button"" class=""ex"" id=""attributeDelete"" name=""attributeDelete"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=Attribute&OfferID=" & OfferID & "&IncentiveAttributeID=" & MyCommon.NZ(row.Item("IncentiveAttributeID"), 0) & "')}"" value=""X"" />")
                  End If
                  Send("  </td>")
                  Send("  <td>")
                  If (i > 1) Then
                    If (MyCommon.NZ(row.Item("AttributeComboID"), 0) = 0) Then
                      'Single
                      Send("<a href=""CPEoffer-con.aspx?mode=ChangeAttributeCombo&ac=1&OfferID=" & OfferID & """>" & Copient.PhraseLib.Lookup("term.and", LanguageID) & "</a>")
                      MyCommon.QueryStr = "update CPE_RewardOptions set AttributeComboID=1 where RewardOptionID=" & roid & ";"
                      MyCommon.LRT_Execute()
                    ElseIf (MyCommon.NZ(row.Item("AttributeComboID"), 0) = 1) Then
                      'And
                      Send("<a href=""CPEoffer-con.aspx?mode=ChangeAttributeCombo&ac=1&OfferID=" & OfferID & """>" & Copient.PhraseLib.Lookup("term.and", LanguageID) & "</a>")
                    Else
                      'Or
                      Send("<a href=""CPEoffer-con.aspx?mode=ChangeAttributeCombo&ac=2&OfferID=" & OfferID & """>" & Copient.PhraseLib.Lookup("term.or", LanguageID) & "</a>")
                    End If
                  End If
                  Send("  </td>")
                  Send("  <td>")
                  Send("    <a href=""javascript:openPopup('CPEoffer-con-attribute.aspx?OfferID=" & OfferID & "&IncentiveAttributeID=" & MyCommon.NZ(row.Item("IncentiveAttributeID"), 0) & "')"">" & Copient.PhraseLib.Lookup("term.attribute", LanguageID) & "</a>")
                  Send("  </td>")
                  Send("  <td>")
                  'Get type-specific details
                  MyCommon.QueryStr = "select IAT.TierLevel, IAT.AttributeTypeID, AT.Description as Name, IAT.AttributeValues from CPE_IncentiveAttributeTiers as IAT with (NoLock) " & _
                                      "left join AttributeTypes as AT on AT.AttributeTypeID=IAT.AttributeTypeID " & _
                                      "where IncentiveAttributeID=" & MyCommon.NZ(row.Item("IncentiveAttributeID"), 0) & ";"
                  rst3 = MyCommon.LRT_Select
                  If rst3.Rows.Count > 0 Then
                    For Each row3 In rst3.Rows
                      Send("    " & MyCommon.NZ(row3.Item("Name"), ""))
                    Next
                  End If
                  Send("  </td>")
                  Send("  <td colspan=""" & TierLevels & """>")
                  'Get value-specific details
                  If (rst3.Rows.Count > 0) Then
                    MyCommon.QueryStr = "select AttributeValueID, Description as Name from AttributeValues with (NoLock) " & _
                                        "where AttributeTypeID=" & MyCommon.NZ(rst3.Rows(0).Item("AttributeTypeID"), 0) & " " & _
                                        "and AttributeValueID in (" & MyCommon.NZ(rst3.Rows(0).Item("AttributeValues"), 0) & ");"
                    rst3 = MyCommon.LRT_Select
                    If rst3.Rows.Count > 0 Then
                      For Each row3 In rst3.Rows
                        Send("    " & MyCommon.NZ(row3.Item("Name"), "") & "<br />")
                      Next
                    End If
                  End If
                  Send("  </td>")
                  If (isTemplate) Then
                    Send("  <td class=""templine"">")
                    %>
                    <input type="checkbox" id="chkLocked12" name="chkLocked" value="<%Sendb(MyCommon.NZ(row.Item("IncentiveAttributeID"), 0))%>"<%Sendb(IIf(MyCommon.NZ(row.Item("DisallowEdit"), False)=True, " checked=""checked""", "")) %> onclick="javascript:updateLocked('lockAttribute<%Sendb(MyCommon.NZ(row.Item("IncentiveAttributeID"), 0))%>', this.checked);"<%Sendb(IIf(IsDBNull(row.Item("IncentiveAttributeID")) AndAlso MyCommon.NZ(row.Item("RequiredFromTemplate"), False), " disabled=""disabled""", "")) %> />
                    <input type="hidden" id="conType12" name="conType" value="Attribute" />
                    <input type="hidden" id="conAttribute<%Sendb(MyCommon.NZ(row.Item("IncentiveAttributeID"), 0))%>" name="con" value="<%Sendb(MyCommon.NZ(row.Item("IncentiveAttributeID"), 0))%>" />
                    <input type="hidden" id="lockAttribute<%Sendb(MyCommon.NZ(row.Item("IncentiveAttributeID"), 0))%>" name="locked" value="<%Sendb(IIf(MyCommon.NZ(row.Item("DisallowEdit"), False)=True, "1", "0"))%>" />
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
              End If
            %>
            
            <!-- PRODUCT CONDITIONS -->
            <%
              t = 1
              ' Find the currently selected groups on page load:
              MyCommon.QueryStr = "select IPG.IncentiveProductGroupID, PG.ProductGroupID, PG.Name, PG.PhraseID, PG.AnyProduct, UT.PhraseID as UnitPhraseID, ExcludedProducts, Rounding, ProductComboID, " & _
                                  " QtyForIncentive, QtyUnitType, AccumMin, AccumLimit, AccumPeriod, DisallowEdit, RequiredFromTemplate, Disqualifier, MinPurchAmt from CPE_IncentiveProductGroups as IPG with (NoLock) " & _
                                  " left join ProductGroups as PG with (NoLock) on PG.ProductGroupID=IPG.ProductGroupID " & _
                                  " left join CPE_RewardOptions as RO with (NoLock) on IPG.RewardOptionID=RO.RewardOptionID " & _
                                  " left join CPE_UnitTypes as UT with (NoLock) on UT.UnitTypeID=IPG.QtyUnitType " & _
                                  " where IPG.RewardOptionID=" & roid & "and IPG.Deleted=0 and Disqualifier=0 " & _
                                  " order by AnyProduct DESC, Name;"
              rst = MyCommon.LRT_Select
              ' Also, go ahead and find the currently DISQUALIFIED groups on page load too:
              MyCommon.QueryStr = "select IPG.IncentiveProductGroupID, PG.ProductGroupID, PG.Name, PG.PhraseID, PG.AnyProduct, UT.PhraseID as UnitPhraseID, ExcludedProducts, Rounding, ProductComboID, " & _
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
                  If rst.Rows.Count > 1 Then
                    If (Logix.UserRoles.EditOffer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False))) Then
                      If (MyCommon.NZ(row.Item("RequiredFromTemplate"), False) And Not isTemplate) Then
                        Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=Product&OfferID=" & OfferID & "&IncentiveProductGroupID=" & row.Item("IncentiveProductGroupID") & "')}"" value=""X"" />")
                      Else
                        Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=Product&OfferID=" & OfferID & "&IncentiveProductGroupID=" & row.Item("IncentiveProductGroupID") & "')}"" value=""X"" />")
                      End If
                    ElseIf (Logix.UserRoles.EditTemplates) Then
                      Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=Product&OfferID=" & OfferID & "&IncentiveProductGroupID=" & row.Item("IncentiveProductGroupID") & "')}"" value=""X"" />")
                    Else
                      Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=Product&OfferID=" & OfferID & "&IncentiveProductGroupID=" & row.Item("IncentiveProductGroupID") & "')}"" value=""X"" />")
                    End If
                  Else
                    If rst2.Rows.Count = 0 Then
                      Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & IIf((Not Logix.UserRoles.EditOffer) Or (FromTemplate AndAlso MyCommon.NZ(row.Item("DisallowEdit"), False)), " disabled=""disabled""", " ") & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=Product&OfferID=" & OfferID & "&IncentiveProductGroupID=" & row.Item("IncentiveProductGroupID") & "')}"" value=""X"" />")
                    Else
                      Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=Product&OfferID=" & OfferID & "&IncentiveProductGroupID=" & row.Item("IncentiveProductGroupID") & "')}"" value=""X"" />")
                    End If
                  End If
                  Send("  </td>")
                  Send("  <td>")
                  ' lets spit out the ProductComboID
                  If (i > 1 And MyCommon.NZ(row.Item("ExcludedProducts"), False) = False) Then
                    If (MyCommon.NZ(row.Item("ProductComboID"), 0) = 0) Then
                      ' single
                      Send("<a href=""CPEoffer-con.aspx?mode=ChangeProductCombo&pc=1&OfferID=" & OfferID & """>" & Copient.PhraseLib.Lookup("term.and", LanguageID) & "</a>")
                      MyCommon.QueryStr = "update CPE_RewardOptions set ProductComboID=1 where RewardOptionID=" & roid
                      MyCommon.LRT_Execute()
                    ElseIf (MyCommon.NZ(row.Item("ProductComboID"), 0) = 1) Then
                      ' and
                      Send("<a href=""CPEoffer-con.aspx?mode=ChangeProductCombo&pc=1&OfferID=" & OfferID & """>" & Copient.PhraseLib.Lookup("term.and", LanguageID) & "</a>")
                    Else
                      ' or
                      Send("<a href=""CPEoffer-con.aspx?mode=ChangeProductCombo&pc=2&OfferID=" & OfferID & """>" & Copient.PhraseLib.Lookup("term.or", LanguageID) & "</a>")
                    End If
                  End If
                  Send("  </td>")
                  Send("  <td>")
                  Send("    <a href=""javascript:openPopup('CPEoffer-con-product.aspx?OfferID=" & OfferID & "&Disqualifier=0&IncentiveProductGroupID=" & IncentiveID & "')"">" & Copient.PhraseLib.Lookup("term.product", LanguageID) & "</a>")
                  Send("  </td>")
                  Send("  <td>")
                  If (MyCommon.NZ(row.Item("ProductGroupID"), -1) > 1) Then
                    If (MyCommon.NZ(row.Item("ExcludedProducts"), False) = True) Then
                      Sendb(Copient.PhraseLib.Lookup("term.anyproduct", LanguageID) & " ")
                      Sendb(StrConv(Copient.PhraseLib.Lookup("term.excluding", LanguageID), VbStrConv.Lowercase) & " ")
                    End If
                    Sendb("<a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("ProductGroupID"), -1) & """>")
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
                      Send("<a href=""pgroup-edit.aspx?ProductGroupID=" & ExcludedProductGroupID & """>" & ExcludedProductGroupName & "</a>")
                    End If
                  End If
                  Send("<br />")
                  'Minimum Purchase Amount
                  If (MyCommon.NZ(row.Item("MinPurchAmt"), 0) = 0) Then
                    Sendb(Copient.PhraseLib.Lookup("term.no", LanguageID) & " " & Copient.PhraseLib.Lookup("term.minimumpurchase", LanguageID))
                  ElseIf (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 1) Then
                    Sendb(Math.Truncate(MyCommon.NZ(row.Item("MinPurchAmt"), 0)) & " " & Copient.PhraseLib.Lookup("term.minimumpurchase", LanguageID))
                  ElseIf (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 2) Then
                    Sendb(FormatCurrency(MyCommon.NZ(row.Item("MinPurchAmt"), 0), 2) & " " & Copient.PhraseLib.Lookup("term.minimumpurchase", LanguageID))
                  ElseIf (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 3) Then
                    Sendb(Math.Round(MyCommon.NZ(row.Item("MinPurchAmt"), 0), 3) & " " & Copient.PhraseLib.Lookup("term.lbsgals", LanguageID) & " " & " " & Copient.PhraseLib.Lookup("term.minimumpurchase", LanguageID))
                  ElseIf (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 10) Then
    Sendb(Math.Round(MyCommon.NZ(row.Item("MinPurchAmt"), 0), 2) & " " & Copient.PhraseLib.Lookup("term.minimumpurchase", LanguageID))
End If

                  Send("  </td>")
                  ' Find the per-tier values:
                  t = 1
                  MyCommon.QueryStr = "select IPG.IncentiveProductGroupID, IPG.Disqualifier, IPGT.TierLevel, IPGT.Quantity as QtyForIncentive, IPG.MinPurchAmt from CPE_IncentiveProductGroups as IPG with (NoLock) " & _
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
                        ElseIf (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 3) Then
                          Sendb(Math.Round(MyCommon.NZ(rst3.Rows(t - 1).Item("QtyForIncentive"), 0), 3) & " " & Copient.PhraseLib.Lookup("term.lbsgals", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.required", LanguageID), VbStrConv.Lowercase))
                        ElseIf (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 4) Then
                          Sendb(Copient.PhraseLib.Lookup("term.qty1atprice", LanguageID) & " " & FormatCurrency(MyCommon.NZ(rst3.Rows(t - 1).Item("QtyForIncentive"), 0)))
                        ElseIf (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 10) Then
    Sendb(Math.Round(MyCommon.NZ(row.Item("QtyForIncentive"), 0), 2) & " " & Copient.PhraseLib.Lookup("term.required", LanguageID))
End If
                        If MyCommon.NZ(row.Item("Rounding"), True) Then
                          Send("<span style=""color:#0000ff;cursor:pointer;"" title=""" & Copient.PhraseLib.Lookup("term.rounded", LanguageID) & """>*</span>")
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
                  End If
                  i += 1
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
                  Send("  <td colspan=""" & TierLevels & """></td>")
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
                      Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=ProductDisqualifier&OfferID=" & OfferID & "&IncentiveProductGroupID=" & row.Item("IncentiveProductGroupID") & "')}"" value=""X"" />")
                    Else
                      Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=ProductDisqualifier&OfferID=" & OfferID & "&IncentiveProductGroupID=" & row.Item("IncentiveProductGroupID") & "')}"" value=""X"" />")
                    End If
                  ElseIf (Logix.UserRoles.EditTemplates) Then
                    Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=ProductDisqualifier&OfferID=" & OfferID & "&IncentiveProductGroupID=" & row.Item("IncentiveProductGroupID") & "')}"" value=""X"" />")
                  Else
                    Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=ProductDisqualifier&OfferID=" & OfferID & "&IncentiveProductGroupID=" & row.Item("IncentiveProductGroupID") & "')}"" value=""X"" />")
                  End If
                End If
                Send("  </td>")
                Send("  <td>")
                Send("    " & Copient.PhraseLib.Lookup("term.not", LanguageID))
                Send("  </td>")
                Send("  <td>")
                Send("    <a href=""javascript:openPopup('CPEoffer-con-product.aspx?OfferID=" & OfferID & "&Disqualifier=1&IncentiveProductGroupID=" & row.Item("IncentiveProductGroupID") & "')"">" & Copient.PhraseLib.Lookup("term.product", LanguageID) & "</a>")
                Send("  </td>")
                Send("  <td>")
                If (MyCommon.NZ(row.Item("ProductGroupID"), -1) > 1) Then
                  If (MyCommon.NZ(row.Item("ExcludedProducts"), False) = True) Then
                    Sendb(StrConv(Copient.PhraseLib.Lookup("term.excluding", LanguageID), VbStrConv.Lowercase) & " ")
                  End If
                  Sendb("<a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("ProductGroupID"), -1) & """>")
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
                                    "where IPG.RewardOptionID=" & roid & " and IPG.Deleted=0 and IPG.Disqualifier=1 order by TierLevel;"
                rst3 = MyCommon.LRT_Select
                If TierLevels = 1 Then
                  Sendb("  <td>")
                Else
                  Sendb("  <td colspan=""" & TierLevels & """>")
                End If
                If rst3.Rows.Count = 0 Then
                  Sendb(Copient.PhraseLib.Lookup("term.undefined", LanguageID))
                Else
                  If t > rst3.Rows.Count Then
                    Sendb(Copient.PhraseLib.Lookup("term.undefined", LanguageID))
                  Else
                    ' QtyForIncentive,QtyUnitType,AccumMin,AccumLimit,AccumPeriod,UnitDescription
                    If Not MyCommon.NZ(row.Item("ExcludedProducts"), False) Then
                      If (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 1) Then
                        Sendb(Math.Truncate(MyCommon.NZ(rst3.Rows(t - 1).Item("QtyForIncentive"), 0)) & " " & StrConv(Copient.PhraseLib.Lookup("term.required", LanguageID), VbStrConv.Lowercase))
                      ElseIf (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 2) Then
                        Sendb(FormatCurrency(MyCommon.NZ(rst3.Rows(t - 1).Item("QtyForIncentive"), 0)) & " " & StrConv(Copient.PhraseLib.Lookup("term.required", LanguageID), VbStrConv.Lowercase))
                      ElseIf (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 3) Or (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 4) Then
                        Sendb(MyCommon.NZ(rst3.Rows(t - 1).Item("QtyForIncentive"), 0) & " " & Copient.PhraseLib.Lookup("term.lbsgals", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.required", LanguageID), VbStrConv.Lowercase))
                      End If
                      Sendb("<br />")
                    End If
                  End If
                End If
                Send("</td>")
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
                                  "left join PointsPrograms as PP with (NoLock) on PP.ProgramID=IPG.ProgramID " & _
                                  "where RewardOptionID=" & roid & " and IPG.Deleted=0;"
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
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=Point&OfferID=" & OfferID & "')}"" value=""X"" />")
                ElseIf (Logix.UserRoles.EditTemplates) Then
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=Point&OfferID=" & OfferID & "')}"" value=""X"" />")
                Else
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=Point&OfferID=" & OfferID & "')}"" value=""X"" />")
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
                  Sendb(Copient.PhraseLib.Lookup("term.undefined", LanguageID))
                  Sendb(" <span class=""red"">")
                  Sendb("(" & Copient.PhraseLib.Lookup("term.requiredbytemplate", LanguageID) & ")")
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
            
            <!-- STORED VALUE CONDITIONS -->
            <%
              t = 1
              MyCommon.QueryStr = "select ISVP.IncentiveStoredValueID, ISVP.SVProgramID, SVP.Name, SVP.SVTypeID, SVT.ValuePrecision, QtyForIncentive, DisallowEdit, RequiredFromTemplate " & _
                                  "from CPE_IncentiveStoredValuePrograms as ISVP with (NoLock) " & _
                                  "left join StoredValuePrograms as SVP with (NoLock) on SVP.SVProgramID=ISVP.SVProgramID " & _
                                  "left join SVTypes as SVT with (NoLock) on SVP.SVTypeID=SVT.SVTypeID " & _
                                  "where RewardOptionID=" & roid & " and ISVP.Deleted=0;"
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
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=StoredValue&OfferID=" & OfferID & "')}"" value=""X"" />")
                ElseIf (Logix.UserRoles.EditTemplates) Then
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=StoredValue&OfferID=" & OfferID & "')}"" value=""X"" />")
                Else                
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=StoredValue&OfferID=" & OfferID & "')}"" value=""X"" />")
                End If
                Send("  </td>")
                Send("  <td>")
                Send("  </td>")
                Send("  <td>")
                Send("    <a href=""javascript:openPopup('CPEoffer-con-sv.aspx?OfferID=" & OfferID & "&IncentiveStoredValueID=" & MyCommon.NZ(row.Item("IncentiveStoredValueID"), 0) & "')"">" & Copient.PhraseLib.Lookup("term.storedvalue", LanguageID) & "</a>")
                Send("  </td>")
                Send("  <td>")
                If (MyCommon.NZ(row.Item("SVProgramID"), -1) > -1) Then
                  Sendb("<a href=""SV-edit.aspx?ProgramGroupID=" & row.Item("SVProgramID") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
                End If
                Send("  </td>")
                ' Find the per-tier values:
                t = 1
                MyCommon.QueryStr = "select ISVP.IncentiveStoredValueID, ISVPT.TierLevel, ISVPT.Quantity, SVP.Value " & _
                                    "from CPE_IncentiveStoredValuePrograms as ISVP with (NoLock) " & _
                                    "left join CPE_IncentiveStoredValueProgramTiers as ISVPT with (NoLock) on ISVPT.IncentiveStoredValueID=ISVP.IncentiveStoredValueID " & _
                                    "left join StoredValuePrograms AS SVP with (NoLock) on SVP.SVProgramID=ISVP.SVProgramID " & _
                                    "where ISVP.Deleted=0 and ISVP.RewardOptionID=" & roid & ";"
                rst2 = MyCommon.LRT_Select
                If rst2.Rows.Count = 0 Then
                  Send("  <td colspan=""" & TierLevels & """>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                Else
                  While t <= TierLevels
                    If t > rst2.Rows.Count Then
                      Send("  <td>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
                    Else
                      Send("  <td>")
                      Sendb("    " & CInt(MyCommon.NZ(rst2.Rows(t - 1).Item("Quantity"), "0")) & " " & StrConv(Copient.PhraseLib.Lookup("term.units", LanguageID), VbStrConv.Lowercase) & " ")
                      If (MyCommon.NZ(row.Item("SVTypeID"), 0) > 1) Then
                        Sendb("($" & Math.Round(MyCommon.NZ(rst2.Rows(t - 1).Item("Value"), 0) * MyCommon.NZ(rst2.Rows(t - 1).Item("Quantity"), 0), MyCommon.NZ(row.Item("ValuePrecision"), 0)) & ") ")
                      End If
                      Send(StrConv(Copient.PhraseLib.Lookup("term.required", LanguageID), VbStrConv.Lowercase))
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
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=Day&OfferID=" & OfferID & "')}"" value=""X"" />")
                ElseIf (Logix.UserRoles.EditTemplates) Then
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=Day&OfferID=" & OfferID & "')}"" value=""X"" />")
                Else                
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=Day&OfferID=" & OfferID & "')}"" value=""X"" />")
                End If
                Send("  </td>")
                Send("  <td>")
                Send("  </td>")
                Send("  <td>")
                Send("    <a href=""javascript:openPopup('CPEoffer-con-day.aspx?OfferID=" & OfferID & "')"">" & Copient.PhraseLib.Lookup("term.day", LanguageID) & "</a>")
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
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=Time&OfferID=" & OfferID & "')}"" value=""X"" />")
                ElseIf (Logix.UserRoles.EditTemplates) Then
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=Time&OfferID=" & OfferID & "')}"" value=""X"" />")
                Else
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=Time&OfferID=" & OfferID & "')}"" value=""X"" />")
                End If
                Send("  </td>")
                Send("  <td>")
                Send("  </td>")
                Send("  <td>")
                Send("    <a href=""javascript:openPopup('CPEoffer-con-time.aspx?OfferID=" & OfferID & "')"">" & Copient.PhraseLib.Lookup("term.time", LanguageID) & "</a>")
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
                                  "where ITT.RewardOptionID=" & roid & " and ITT.Deleted=0;"
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
                  
                  i = 1
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
                        Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=Tender&OfferID=" & OfferID & "&IncentiveTenderID=" & row.Item("IncentiveTenderID") & "')}"" value=""X"" />")
                      Else
                        Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=Tender&OfferID=" & OfferID & "&IncentiveTenderID=" & row.Item("IncentiveTenderID") & "')}"" value=""X"" />")
                      End If
                    ElseIf (Logix.UserRoles.EditTemplates) Then
                      Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=Tender&OfferID=" & OfferID & "&IncentiveTenderID=" & row.Item("IncentiveTenderID") & "')}"" value=""X"" />")
                    Else
                      Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=Tender&OfferID=" & OfferID & "&IncentiveTenderID=" & row.Item("IncentiveTenderID") & "')}"" value=""X"" />")
                    End If
                    Send("  </td>")
                    Send("  <td>")
                    If i > 1 Then
                      Send(Copient.PhraseLib.Lookup("term.or", LanguageID))
                    End If
                    Send("  </td>")
                    Send("  <td>")
                    Send("    <a href=""javascript:openPopup('CPEoffer-con-tender.aspx?OfferID=" & OfferID & "&IncentiveTenderID=" & row.Item("IncentiveTenderID") & "')"">" & Copient.PhraseLib.Lookup("term.tender", LanguageID) & "</a>")
                    Send("  </td>")
                    Send("  <td>")
                    If (isTender) Then
                      If TenderExcluded Then
                        Sendb(Copient.PhraseLib.Lookup("term.allbut", LanguageID) & ":<br />")
                      End If
                      Sendb("<a href=""tender-engines.aspx"">" & MyCommon.SplitNonSpacedString(TenderList, 25) & "</a>")
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
                      Send("    " & IIf(TenderDisallowEdit, "Yes", "No"))
                      Send("  </td>")
                    End If

                    Send("</tr>")
                    i += 1
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
                  i = 1
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
                        Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=Tender&OfferID=" & OfferID & "&IncentiveTenderID=" & row.Item("IncentiveTenderID") & "')}"" value=""X"" />")
                      Else
                        Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=Tender&OfferID=" & OfferID & "&IncentiveTenderID=" & row.Item("IncentiveTenderID") & "')}"" value=""X"" />")
                      End If
                    ElseIf (Logix.UserRoles.EditTemplates) Then
                      Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=Tender&OfferID=" & OfferID & "&IncentiveTenderID=" & row.Item("IncentiveTenderID") & "')}"" value=""X"" />")
                    Else
                      Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=Tender&OfferID=" & OfferID & "&IncentiveTenderID=" & row.Item("IncentiveTenderID") & "')}"" value=""X"" />")
                    End If
                    Send("  </td>")
                    Send("  <td>")
                    If i > 1 Then
                      Send(Copient.PhraseLib.Lookup("term.or", LanguageID))
                    End If
                    Send("  </td>")
                    Send("  <td>")
                    Send("    <a href=""javascript:openPopup('CPEoffer-con-tender.aspx?OfferID=" & OfferID & "&IncentiveTenderID=" & row.Item("IncentiveTenderID") & "')"">" & Copient.PhraseLib.Lookup("term.tender", LanguageID) & "</a>")
                    Send("  </td>")
                    Send("  <td>")
                    If (isTender) Then
                      If TenderExcluded Then
                        Sendb(Copient.PhraseLib.Lookup("term.allbut", LanguageID) & ":<br />")
                      End If
                      Sendb("<a href=""tender-engines.aspx"">" & MyCommon.SplitNonSpacedString(TenderList, 25) & "</a>")
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
                    Send("    " & IIf(TenderDisallowEdit, "Yes", "No"))
                    Send("  </td>")
                  End If
                  i += 1
                Next
                Send("</tr>")
              End If
            End If
            %>
            
            <!-- STORE-LEVEL INSTANT WIN CONDITIONS -->
            <%
              t = 1
              MyCommon.QueryStr = "select IncentiveInstantWinID, NumPrizesAllowed, OddsOfWinning, RandomWinners, DisallowEdit, RequiredFromTemplate " & _
                                  "from CPE_IncentiveInstantWin as IWW with (NoLock) " & _
                                  "where Deleted=0 and RewardOptionID=" & roid & ";"
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
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=InstantWin&OfferID=" & OfferID & "')}"" value=""X"" />")
                ElseIf (Logix.UserRoles.EditTemplates) Then
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=InstantWin&OfferID=" & OfferID & "')}"" value=""X"" />")
                Else
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=InstantWin&OfferID=" & OfferID & "')}"" value=""X"" />")
                End If
                Send("  </td>")
                Send("  <td>")
                Send("  </td>")
                Send("  <td>")
                Send("    <a href=""javascript:openPopup('CPEoffer-con-instantwin.aspx?OfferID=" & OfferID & "')"">" & Copient.PhraseLib.Lookup("term.instantwin", LanguageID) & "</a>")
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
                  Send("    " & IIf(MyCommon.NZ(row.Item("DisallowEdit"), False) = True, "Yes", "No"))
                  Send("  </td>")
                End If
                Send("</tr>")
              Next
            %>
            
            <!-- ENTERPRISE INSTANT WIN CONDITIONS -->
            <%
              t = 1
              MyCommon.QueryStr = "select IncentiveEIWID, NumberOfPrizes, EIW.FrequencyID, EIWF.Description, DisallowEdit, RequiredFromTemplate " & _
                                  "from CPE_IncentiveEIW as EIW with (NoLock) " & _
                                  "inner join CPE_IncentiveEIWFrequency as EIWF on EIWF.FrequencyID=EIW.FrequencyID " & _
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
                isEIW = True
                Send("<tr class=""shaded"">")
                Send("  <td>")
                If (Logix.UserRoles.EditOffer And Not (FromTemplate And MyCommon.NZ(row.Item("DisallowEdit"), False))) Then
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=EInstantWin&OfferID=" & OfferID & "&IncentiveEIWID=" & row.Item("IncentiveEIWID") & "')}"" value=""X"" />")
                ElseIf (Logix.UserRoles.EditOffer) Then
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=EInstantWin&OfferID=" & OfferID & "&IncentiveEIWID=" & row.Item("IncentiveEIWID") & "')}"" value=""X"" />")
                Else
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&Option=EInstantWin&OfferID=" & OfferID & "&IncentiveEIWID=" & row.Item("IncentiveEIWID") & "')}"" value=""X"" />")
                End If
                Send("  </td>")
                Send("  <td>")
                Send("  </td>")
                Send("  <td>")
                Send("    <a href=""javascript:openPopup('CPEoffer-con-einstantwin.aspx?OfferID=" & OfferID & "&IncentiveEIWID=" & MyCommon.NZ(row.Item("IncentiveEIWID"), 0) & "')"">" & Copient.PhraseLib.Lookup("term.instantwin", LanguageID) & "</a>")
                Send("  </td>")
                Send("  <td>")
                MyCommon.QueryStr = "select count(*) as TriggerCount from CPE_EIWTriggers where IncentiveEIWID=" & MyCommon.NZ(row.Item("IncentiveEIWID"), 0) & " and RewardOptionID=" & roid & " and Removed=0;"
                rst2 = MyCommon.LRT_Select
                If rst2.Rows(0).Item("TriggerCount") = 0 Then
                  Sendb(Copient.PhraseLib.Lookup("term.no", LanguageID))
                Else
                  Sendb(MyCommon.NZ(rst2.Rows(0).Item("TriggerCount"), 0))
                End If
                If rst2.Rows(0).Item("TriggerCount") = 1 Then
                  Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.trigger", LanguageID), VbStrConv.Lowercase))
                Else
                  Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.triggers", LanguageID), VbStrConv.Lowercase))
                End If
                Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.for", LanguageID), VbStrConv.Lowercase))
                Sendb(" " & MyCommon.NZ(row.Item("NumberOfPrizes"), 0))
                If MyCommon.NZ(row.Item("NumberOfPrizes"), 0) = 1 Then
                  Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.prize", LanguageID), VbStrConv.Lowercase))
                Else
                  Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.prizes", LanguageID), VbStrConv.Lowercase))
                End If
                Sendb(" " & StrConv(MyCommon.NZ(row.Item("Description"), ""), VbStrConv.Lowercase))
                Send("  </td>")
                Send("  <td colspan=""" & TierLevels & """>")
                Send("  </td>")
                If (isTemplate) Then
                  Send("  <td class=""templine"">")
                  %>
                  <input type="checkbox" id="chkLocked11" name="chkLocked" value="<%Sendb(MyCommon.NZ(row.Item("IncentiveEIWID"), 0))%>"<%Sendb(IIf(MyCommon.NZ(row.Item("DisallowEdit"), False)=True, " checked=""checked""", "")) %> onclick="javascript:updateLocked('lockEIW<%Sendb(MyCommon.NZ(row.Item("IncentiveEIWID"), 0))%>', this.checked);"<%Sendb(IIf(MyCommon.NZ(row.Item("RequiredFromTemplate"), False), " disabled=""disabled""", "")) %> />
                  <input type="hidden" id="conType11" name="conType" value="EInstantWin" />
                  <input type="hidden" id="conEIW<%Sendb(MyCommon.NZ(row.Item("IncentiveEIWID"), 0))%>" name="con" value="<%Sendb(MyCommon.NZ(row.Item("IncentiveEIWID"), 0))%>" />
                  <input type="hidden" id="lockEIW<%Sendb(MyCommon.NZ(row.Item("IncentiveEIWID"), 0))%>" name="locked" value="<%Sendb(IIf(MyCommon.NZ(row.Item("DisallowEdit"), False)=True, "1", "0"))%>" />
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
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&amp;Option=PLU&amp;OfferID=" & OfferID & "&amp;IncentivePLUID=" & MyCommon.NZ(row.Item("IncentivePLUID"), 0) & "')}"" value=""X"" />")
                ElseIf (Logix.UserRoles.EditTemplates) Then
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&amp;Option=PLU&amp;OfferID=" & OfferID & "&amp;IncentivePLUID=" & MyCommon.NZ(row.Item("IncentivePLUID"), 0) & "')}"" value=""X"" />")
                Else
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&amp;Option=PLU&amp;OfferID=" & OfferID & "&amp;IncentivePLUID=" & MyCommon.NZ(row.Item("IncentivePLUID"), 0) & "')}"" value=""X"" />")
                End If
                Send("  </td>")
                Send("  <td>")
                If i > 1 Then
                  Send("    " & Copient.PhraseLib.Lookup("term.or", LanguageID))
                End If
                Send("  </td>")
                Send("  <td>")
                Send("    <a href=""javascript:openPopup('CPEoffer-con-plu.aspx?OfferID=" & OfferID & "&amp;IncentivePLUID=" & MyCommon.NZ(row.Item("IncentivePLUID"), 0) & "')"">" & Copient.PhraseLib.Lookup("term.triggercode", LanguageID) & "</a>")
                Send("  </td>")
                Send("  <td>")
                If MyCommon.NZ(row.Item("PLU"), "") = "" Then
                  Send("    " & Copient.PhraseLib.Lookup("term.undefined", LanguageID))
                Else
                  Send("    " & MyCommon.NZ(row.Item("PLU"), ""))
                End If
                Send("  </td>")
                Send("  <td colspan=""" & TierLevels & """>")
                Send("    " & IIf(MyCommon.NZ(row.Item("PerRedemption"), False), Copient.PhraseLib.Lookup("term.OncePerRedemption", LanguageID), Copient.PhraseLib.Lookup("term.oncepertransaction", LanguageID)))
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
            
            <!-- CARD TYPE CONDITIONS -->
            <%
              t = 1
              MyCommon.QueryStr = "select IncentiveCardTypeID, CardTypeID, DisallowEdit, RequiredFromTemplate " & _
                                  "from CPE_IncentiveCardTypes as CT with (NoLock) " & _
                                  "where RewardOptionID=" & roid & " and Deleted=0 " & _
                                  "order by IncentiveCardTypeID;"
              rst = MyCommon.LRT_Select
              If (rst.Rows.Count > 0) Then
                Send("<tr class=""shadeddark"">")
                Send("  <td colspan=""4"">")
                Send("    <h3>")
                Send("      " & Copient.PhraseLib.Lookup("term.cardtypeconditions", LanguageID))
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
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&amp;Option=CardType&amp;OfferID=" & OfferID & "&amp;IncentiveCardTypeID=" & MyCommon.NZ(row.Item("IncentiveCardTypeID"), 0) & "')}"" value=""X"" />")
                ElseIf (Logix.UserRoles.EditTemplates) Then
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & DeleteBtnDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&amp;Option=CardType&amp;OfferID=" & OfferID & "&amp;IncentiveCardTypeID=" & MyCommon.NZ(row.Item("IncentiveCardTypeID"), 0) & "')}"" value=""X"" />")
                Else
                  Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&amp;Option=CardType&amp;OfferID=" & OfferID & "&amp;IncentiveCardTypeID=" & MyCommon.NZ(row.Item("IncentiveCardTypeID"), 0) & "')}"" value=""X"" />")
                End If
                Send("  </td>")
                Send("  <td>")
                If i > 1 Then
                  Send("    " & Copient.PhraseLib.Lookup("term.or", LanguageID))
                End If
                Send("  </td>")
                Send("  <td>")
                Send("    <a href=""javascript:openPopup('CPEoffer-con-cardtype.aspx?OfferID=" & OfferID & "')"">" & Copient.PhraseLib.Lookup("term.cardtype", LanguageID) & "</a>")
                Send("  </td>")
                Send("  <td>")
                If MyCommon.NZ(row.Item("CardTypeID"), -1) = -1 Then
                  Send("    " & Copient.PhraseLib.Lookup("term.undefined", LanguageID))
                Else
                  MyCommon.QueryStr = "select Description, PhraseID from CardTypes with (NoLock) where CardTypeID=" & MyCommon.NZ(row.Item("CardTypeID"), -1) & ";"
                  rst2 = MyCommon.LXS_Select
                  If rst2.Rows.Count > 0 Then
                    If (MyCommon.NZ(rst2.Rows(0).Item("PhraseID"), 0) > 0) Then
                      Send("    " & Copient.PhraseLib.Lookup(rst2.Rows(0).Item("PhraseID"), LanguageID))
                    Else
                      Send("    " & MyCommon.NZ(rst2.Rows(0).Item("Description"), "&nbsp;"))
                    End If
                  Else
                    Send("    " & Copient.PhraseLib.Lookup("term.unknown", LanguageID))
                  End If
                End If
                Send("  </td>")
                Send("  <td colspan=""" & TierLevels & """>&nbsp;</td>")
                If (isTemplate) Then
                  Send("  <td class=""templine"">")
                    %>
                    <input type="checkbox" id="chkLocked13" name="chkLocked" value="<%Sendb(MyCommon.NZ(row.Item("IncentiveCardTypeID"), 0))%>"<%Sendb(IIf(MyCommon.NZ(row.Item("DisallowEdit"), False)=True, " checked=""checked""", "")) %> onclick="javascript:updateLocked('lockCardType<%Sendb(MyCommon.NZ(row.Item("IncentiveCardTypeID"), 0))%>', this.checked);"<%Sendb(IIf(MyCommon.NZ(row.Item("RequiredFromTemplate"), False), " disabled=""disabled""", "")) %> />
                    <input type="hidden" id="conType13" name="conType" value="CardType" />
                    <input type="hidden" id="conCardType<%Sendb(MyCommon.NZ(row.Item("IncentiveCardTypeID"), 0))%>" name="con" value="<%Sendb(MyCommon.NZ(row.Item("IncentiveCardTypeID"), 0))%>" />
                    <input type="hidden" id="lockCardType<%Sendb(MyCommon.NZ(row.Item("IncentiveCardTypeID"), 0))%>" name="locked" value="<%Sendb(IIf(MyCommon.NZ(row.Item("DisallowEdit"), False)=True, "1", "0"))%>" />
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

            <!-- PREFERENCE CONDITIONS -->
            <%
              If MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER) Then
              t = 1
              MyCommon.QueryStr = "select CIP.IncentivePrefsID, CIP.PreferenceID, CIP.DisallowEdit, CIP.RequiredFromTemplate, RO.PreferenceComboID " & _
                                  "from CPE_IncentivePrefs as CIP with (NoLock) " & _
                                  "inner join CPE_RewardOptions as RO with (NoLock) on RO.RewardOptionID = CIP.RewardOptionID " & _
                                  "where CIP.RewardOptionID=" & roid & " order by CIP.IncentivePrefsID;"
              rst = MyCommon.LRT_Select
              If (rst.Rows.Count > 0) Then
                Send("<tr class=""shadeddark"">")
                Send("  <td colspan=""4"">")
                Send("    <h3>")
                Send("      " & Copient.PhraseLib.Lookup("term.preferenceconditions", LanguageID))
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
                Send("<tr class=""shaded"">")
                Send("  <td>")
                If Not isTemplate Then
                   If (Logix.UserRoles.EditOffer) Then
                     Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & IIf((FromTemplate AndAlso MyCommon.NZ(row.Item("DisallowEdit"), False)), " disabled=""disabled""", " ") & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&amp;Option=Preference&amp;OfferID=" & OfferID & "&amp;IncentivePrefsID=" & MyCommon.NZ(row.Item("IncentivePrefsID"), 0) & "')}"" value=""X"" />")
                   Else
                     Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&amp;Option=Preference&amp;OfferID=" & OfferID & "&amp;IncentivePrefsID=" & MyCommon.NZ(row.Item("IncentivePrefsID"), 0) & "')}"" value=""X"" />")
                   End If
                Else                   
                     Sendb("<input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & IIf(Logix.UserRoles.EditTemplates," ", " disabled=""disabled""") & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup("condition.confirmdelete", LanguageID) & "')){LoadDocument('CPEoffer-con.aspx?mode=Delete&amp;Option=Preference&amp;OfferID=" & OfferID & "&amp;IncentivePrefsID=" & MyCommon.NZ(row.Item("IncentivePrefsID"), 0) & "')}"" value=""X"" />")                   
                End If
                Send("  </td>")
                Send("  <td>")
                If i > 1 Then
                  If MyCommon.NZ(row.Item("PreferenceComboID"), 0) = 1 Then
                    ' and
                    Send("<a href=""CPEoffer-con.aspx?mode=ChangePreferenceCombo&pc=1&OfferID=" & OfferID & """>" & Copient.PhraseLib.Lookup("term.and", LanguageID) & "</a>")
                  Else
                    ' or
                    Send("<a href=""CPEoffer-con.aspx?mode=ChangePreferenceCombo&pc=2&OfferID=" & OfferID & """>" & Copient.PhraseLib.Lookup("term.or", LanguageID) & "</a>")
                  End If
                End If
                Send("  </td>")
                Send("  <td>")
                Send("    <a href=""javascript:openPopup('CPEoffer-con-pref.aspx?OfferID=" & OfferID & "&amp;IncentivePrefsID=" & MyCommon.NZ(row.Item("IncentivePrefsID"), 0) & "')"">" & Copient.PhraseLib.Lookup("term.preference", LanguageID) & "</a>")
                Send("  </td>")
                Send("  <td>")
                Send_Preference_Details(MyCommon, MyCommon.NZ(row.Item("PreferenceID"), 0))
                Send("  </td>")
                Send_Preference_Info(MyCommon, MyCommon.NZ(row.Item("IncentivePrefsID"), 0), roid, TierLevels)
                If (isTemplate) Then
                  Send("  <td class=""templine"">")
                  %>
                  <input type="checkbox" id="chkLocked14" name="chkLocked" value="<%Sendb(MyCommon.NZ(row.Item("IncentivePrefsID"), 0))%>"<%Sendb(IIf(MyCommon.NZ(row.Item("DisallowEdit"), False)=True, " checked=""checked""", "")) %> onclick="javascript:updateLocked('lockPref<%Sendb(MyCommon.NZ(row.Item("IncentivePrefsID"), 0))%>', this.checked);"<%Sendb(IIf(MyCommon.NZ(row.Item("RequiredFromTemplate"), False), " disabled=""disabled""", "")) %> />
                  <input type="hidden" id="conType14" name="conType" value="Preference" />
                  <input type="hidden" id="conPref<%Sendb(MyCommon.NZ(row.Item("IncentivePrefsID"), 0))%>" name="con" value="<%Sendb(MyCommon.NZ(row.Item("IncentivePrefsID"), 0))%>" />
                  <input type="hidden" id="lockPref<%Sendb(MyCommon.NZ(row.Item("IncentivePrefsID"), 0))%>" name="locked" value="<%Sendb(IIf(MyCommon.NZ(row.Item("DisallowEdit"), False)=True, "1", "0"))%>" />
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
              End If 'is EPM integration installed
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
          If (isCustomer OrElse isTargeted) Then
            isTargeted = True
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
            TempQueryStr = "SELECT PECT.EngineID, PECT.EngineSubTypeID, PECT.ComponentTypeID, CT.ConditionTypeID, CT.Description, CT.PhraseID, PECT.Singular, " & _
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
                                "WHERE EngineID=2 AND EngineSubTypeID=" & EngineSubTypeID & " AND PECT.ComponentTypeID=1 AND Enabled=1"
            'Impose a few special limits on the query based on various in-page factors:
            If (Not isTargeted) Then
              'Offer has no customer or attribute condition, so limit it to just those
              If (IsFooterOffer AndAlso Not isCustomer) Then
                TempQueryStr &= " AND CT.ConditionTypeID=1"
              ElseIf (Not isCustomer AndAlso Not isAttribute) Then
                TempQueryStr &= " AND CT.ConditionTypeID in (1,12)"
              End If
            End If
            If (AccumEnabled) Then
              'Accumulation is on, so no more product conditions
              TempQueryStr &= " AND CT.ConditionTypeID<>2"
            End If
            If (TierLevels > 1) Then
              'Offer is multitiered, so instant win is invalid
              TempQueryStr &= " AND CT.ConditionTypeID<>8"
            End If
            If (Not isProduct) OrElse (isProduct AndAlso AccumEnabled) Then
              'Offer has no product condition (or has one with accumulation), so disallow product disqualifiers
              TempQueryStr &= " AND CT.ConditionTypeID<>10"
            End If
            If CPEOffer_Has_AnyCustomer(MyCommon, OfferID) Then
              'Offer has AnyCustomer selected as the customer group condition.  Disallow other condition types
              'that require knowledge of who the customer is (points, stored value, etc.)
              TempQueryStr &= " AND CT.ConditionTypeID not in (3, 4, 12, 13)"
            End If
            If Not MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER) Then
              TempQueryStr &= " AND CT.ConditionTypeID<>14"
            End If
            TempQueryStr &= " ORDER BY DisplayOrder;"
            MyCommon.QueryStr = TempQueryStr
            rst = MyCommon.LRT_Select
            If rst.Rows.Count > 0 Then
              Send("<label for=""newconglobal"">" & Copient.PhraseLib.Lookup("offer-con.addglobal", LanguageID) & ":</label><br />")
              Send("<select id=""newconglobal"" name=""newconglobal""" & IIf(isTemplate OrElse (Not Disallow_Conditions), "", " disabled=""disabled""") & DeleteBtnDisabled & ">")
              For Each row In rst.Rows
                If (row.Item("Singular") = False) OrElse (row.Item("Singular") = True AndAlso row.Item("ItemCount") = 0) Then
                  Send("<option value=""" & row.Item("ConditionTypeID") & """>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "</option>")
                End If
              Next
              Send("</select>")
              Sendb("<input class=""regular"" id=""addGlobal"" name=""addGlobal"" type=""submit"" value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """" & DeleteBtnDisabled)
              If isTemplate OrElse (Not (isCustomer And isProduct And isProductDisqualifier And isPoint And isStoredValue And isDay And isTime And isTender And isInstantWin) And Not Disallow_Conditions) Then
              Else
                Sendb(" disabled=""disabled""")
              End If
              Sendb(" />")
            End If
          End If
          Send("<br />")
        %>
      </div>
      </form>
    </div>
    <br clear="all" />
  </div>
<!-- #Include virtual="/include/graphic-reward.inc" -->

<script runat="server">
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
    MyCommon.QueryStr = " SELECT RewardOptionID, ICG.IncentiveCustomerID, CG.CustomerGroupID, ExcludedUsers, DisallowEdit, RequiredFromTemplate " & _
                        "FROM CPE_IncentiveCustomerGroups AS ICG WITH (NOLOCK)" & _
                        "LEFT JOIN CustomerGroups AS CG WITH (NOLOCK)" & _
                        "ON CG.CustomerGroupID=ICG.CustomerGroupID WHERE RewardOptionID IN (SELECT RewardOptionID  FROM CPE_RewardOptions WITH (NOLOCK)" & _
                        "WHERE IncentiveID = @OfferID AND TouchResponse=0 AND Deleted=0) AND ICG.Deleted = 0 AND ExcludedUsers = 0"
    MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
    dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
    If dt.Rows.Count > 0 Then
      For Each row As DataRow In dt.Rows
        ''uc.RewardOptionID = row("RewardOptionID")
        ''uc.IncentiveCustomerID = row("IncentiveCustomerID")
        Dim CustomerGroupID As Int32
        CustomerGroupID = row("CustomerGroupID")
        ''Any Card holders, Any Customers or All CAM Cardholders group is Included in Customer Condition then disallow to add customer condition 
        
        If (CustomerGroupID = 1 Or CustomerGroupID = 2 Or CustomerGroupID = 4) Then
          uc.IsOptInDisabled = True
          Exit for
        End If
        
      Next
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

  Sub Send_Preference_Info(ByRef Common As Copient.CommonInc, ByVal IncentivePrefsID As Integer, ByVal ROID As Long, ByVal TierLevels As Integer)
    Dim dt, dt2 As DataTable
    Dim PreferenceID As Long = 0
    Dim ComboText As String = ""
    Dim i As Integer = 0
    Dim CellCount As Integer = 0

    ' find all the tiers in this preference condition
    Common.QueryStr = "select CIPT.IncentivePrefTiersID, CIPT.TierLevel, CIPT.ValueComboTypeID, CIP.PreferenceID " & _
                      "from CPE_IncentivePrefTiers as CIPT with (NoLock) " & _
                      "inner join CPE_IncentivePrefs as CIP with (NoLock) on CIP.IncentivePrefsID = CIPT.IncentivePrefsID " & _
                      "where CIPT.IncentivePrefsID=" & IncentivePrefsID & " and CIP.RewardOptionID=" & ROID & " " & _
                      "order by CIPT.TierLevel;"
    dt = Common.LRT_Select
    If dt.Rows.Count > 0 Then
      For Each row As DataRow In dt.Rows
        PreferenceID = Common.NZ(row.Item("PreferenceID"), 0)
        ComboText = IIf(Common.NZ(row.Item("ValueComboTypeID"), 1) = 1, "term.and", "term.or")
        ComboText = Copient.PhraseLib.Lookup(ComboText, LanguageID)
      
        CellCount += 1
        If CellCount > TierLevels Then Exit For

        Send("<td>")
        ' find all the tier values
        Common.QueryStr = "select IPTV.PKID, IPTV.Value, IPTV.DateOperatorTypeID, " & _
                          "  case when POT.PhraseID is null then POT.Description" & _
                          "  else Convert(nvarchar(200), PT.Phrase) end as OperatorText " & _
                          "from CPE_IncentivePrefTierValues as IPTV with (NoLock) " & _
                          "inner join CPE_PrefOperatorTypes as POT with (NoLock) on POT.PrefOperatorTypeID = IPTV.OperatorTypeID " & _
                          "left join PhraseText as PT with (NoLock) on PT.PhraseID = POT.PhraseID and LanguageID=" & LanguageID & " " & _
                          "where IPTV.IncentivePrefTiersID=" & Common.NZ(row.Item("IncentivePrefTiersID"), 0)
        dt2 = Common.LRT_Select
        For i = 0 To dt2.Rows.Count - 1
          If Common.NZ(dt2.Rows(i).Item("DateOperatorTypeID"), 0) > 0 Then
            Send(Get_Date_Display_Text(Common, dt2.Rows(i).Item("PKID")))
          Else
            Send(Common.NZ(dt2.Rows(i).Item("OperatorText"), "") & " " & Get_Preference_Value(Common, PreferenceID, Common.NZ(dt2.Rows(i).Item("Value"), "")))
          End If

          If i < dt2.Rows.Count - 1 Then
            Send(" <i>" & ComboText.ToLower & "</i> ")
          End If
        Next
        Send("</td>")
      Next
      
      ' account for any tiers that don't have saved information due to increasing the tiers on an existing offer
      For i = CellCount To (TierLevels - 1)
        Send("  <td>" & Copient.PhraseLib.Lookup("term.undefined", LanguageID) & "</td>")
      Next
      
    Else
      For i = 1 To TierLevels
        Send("  <td>&nbps;</td>")
      Next
    End If
      
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
  
  Function Get_Date_Display_Text(ByRef Common As Copient.CommonInc, ByVal TierValuePKID As Integer) As String
    Dim DisplayText As String = ""
    Dim dt As DataTable
    Dim ValueModifier As String = ""
    Dim Offset As Integer
    
    Common.QueryStr = "select IPTV.Value, IPTV.ValueModifier, IPTV.ValueTypeID, POT.PhraseID as OperatorPhraseID, " & _
                      "PDOT.PhraseID as DateOpPhraseID " & _
                      "from CPE_IncentivePrefTierValues as IPTV with (NoLock) " & _
                      "inner join CPE_PrefOperatorTypes as POT with (NoLock) on POT.PrefOperatorTypeID = IPTV.OperatorTypeID " & _
                      "inner join CPE_PrefDateOperatorTypes as PDOT with (NoLock) on PDOT.PrefDateOperatorTypeID = IPTV.DateOperatorTypeID " & _
                      "where PKID=" & TierValuePKID & ";"
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
    End If
    
    Return DisplayText
  End Function

</script>

<%If (isCustomer OrElse isAttribute OrElse isProduct OrElse isPoint OrElse isDay OrElse isStoredValue) Then%>
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
  If MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER) Then
    If MyCommon.PMRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_PrefManRT()
  End If
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  Send_BodyEnd()
  MyCommon = Nothing
  Logix = Nothing
%>
