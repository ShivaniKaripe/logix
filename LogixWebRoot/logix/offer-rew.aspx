<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: offer-rew.aspx 
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
  Dim MyImport As New Copient.ImportXml(MyCommon)

  Dim AdminUserID As Long
  Dim OfferID As Long
  Dim rst As DataTable
  Dim row As DataRow
  Dim rst2 As DataTable
  Dim row2 As DataRow
  Dim rst3 As DataTable
  Dim Name As String = ""
  Dim NumTiers As String = "0"
  Dim iNumtiers As Integer
  Dim LinkID As Long
  Dim RewardID As Long
  Dim prog As Integer
  Dim x As Integer
  Dim isTemplate As Boolean = False
  Dim FromTemplate As Boolean = False
  Dim Disallow_Rewards As Boolean = False
  Dim bShowGenericXml As Boolean = True
  Dim bShowCatalina As Boolean = True
  Dim bShowBinRanges As Boolean = True
  Dim iRewardType As Integer
  Dim iTiered As Integer
  Dim infoMessage As String = ""
  Dim modMessage As String = ""
  Dim Handheld As Boolean = False
  Dim NumOfCols As Integer = 7
  Dim Rewards As String() = Nothing
  Dim LockedStatus As String() = Nothing
  Dim LoopCtr As Integer = 0
  Dim BannersEnabled As Boolean = False
  Dim StatusCode As Copient.LogixInc.STATUS_FLAGS = Copient.LogixInc.STATUS_FLAGS.STATUS_UNKNOWN
  Dim StatusText As String = ""
  
  Dim objTemp As Object
  Dim intNumDecimalPlaces As Integer
  Dim decFactor As Decimal
  Dim decTemp As Decimal
  Dim sTemp1 As String
  Dim sTemp2 As String
  
  Dim bStoreUser As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "offer-rew.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)

  'Store User
  If (MyCommon.Fetch_CM_SystemOption(131) = "1") Then
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
  
  Dim bEnableBuckOffers As Boolean
  Dim bBuckChildOffer As Boolean
  Dim bBuckParentOffer As Boolean
  Dim oBuckStatus As Copient.ImportXml.BuckOfferStatus
  bEnableBuckOffers = MyCommon.IsEngineInstalled(0) AndAlso (MyCommon.Fetch_CM_SystemOption(137) = "1")
  If bEnableBuckOffers Then
    oBuckStatus = MyImport.BuckOfferGetStatus(OfferID)
    Select Case oBuckStatus
      Case Copient.ImportXml.BuckOfferStatus.BuckParentNoChildren,
       Copient.ImportXml.BuckOfferStatus.BuckParentBothChildren,
       Copient.ImportXml.BuckOfferStatus.BuckParentPaperChildrenOnly,
       Copient.ImportXml.BuckOfferStatus.BuckParentDigitalChildrenOnly
        bBuckParentOffer = True
        bBuckChildOffer = False
      Case Copient.ImportXml.BuckOfferStatus.BuckChildPaper,
       Copient.ImportXml.BuckOfferStatus.BuckChildDigital
        bBuckParentOffer = False
        bBuckChildOffer = True
      Case Copient.ImportXml.BuckOfferStatus.BuckTiered
        bBuckParentOffer = True
        bBuckChildOffer = False
      Case Else
        bBuckParentOffer = False
        bBuckChildOffer = False
        If oBuckStatus = Copient.ImportXml.BuckOfferStatus.ErrorOccurred Then
          infoMessage = MyImport.GetErrorMsg()
        End If
    End Select
  Else
    bBuckParentOffer = False
    bBuckChildOffer = False
    oBuckStatus = Copient.ImportXml.BuckOfferStatus.NotBuckOffer
  End If

  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  
  objTemp = MyCommon.Fetch_CM_SystemOption(41)
  If Not (Integer.TryParse(objTemp.ToString, intNumDecimalPlaces)) Then
    intNumDecimalPlaces = 0
  End If
  decFactor = (10 ^ intNumDecimalPlaces)

  RewardID = Request.QueryString("RewardID")
  
  ' dig the offer info out of the database
  MyCommon.QueryStr = "select OfferID,IsTemplate,FromTemplate,Name,Description,OfferCategoryID,OfferTypeID,ProdStartDate,ProdEndDate,TestStartDate,TestEndDate,TierTypeID,NumTiers,DistPeriod,DistPeriodLimit,DistPeriodVarID,EmployeeFiltering,NonEmployeesOnly,CRMRestricted,LastUpdate,PriorityLevel,EngineID,SharedLimitID,StatusFlag from Offers with (nolock) where OfferID=" & OfferID & " and Deleted=0 and visible=1"
  rst = MyCommon.LRT_Select()
  For Each row In rst.Rows
    NumTiers = row.Item("NumTiers")
    Name = row.Item("Name")
    isTemplate = row.Item("IsTemplate")
    FromTemplate = row.Item("FromTemplate")
  Next
  
  Send_HeadBegin("term.offer", "term.rewards", OfferID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
  Send_HeadEnd()
  
  ' first we need to find out if the item were acting on is tiered or not
  Dim deleteTiered As Integer
  deleteTiered = 0
  
  MyCommon.QueryStr = "select Tiered from OfferRewards with (NoLock) where RewardID=" & RewardID
  rst = MyCommon.LRT_Select
  For Each row In rst.Rows
    If (row.Item("Tiered") = "1") Then
      deleteTiered = 1
    End If
  Next
  
  If (Request.QueryString("Save") <> "") Then
    If (Request.QueryString("IsTemplate") = "IsTemplate") Then
      ' time to update the status bits for the templates
      Dim form_Disallow_Rewards As Integer = 0
      If (Request.QueryString("Disallow_Rewards") = "on") Then
        form_Disallow_Rewards = 1
      End If
      MyCommon.QueryStr = "update TemplatePermissions with (RowLock) set Disallow_Rewards=" & form_Disallow_Rewards & _
      " where OfferID=" & OfferID
      MyCommon.LRT_Execute()
    
      'Update the lock status for each condition
      Rewards = Request.QueryString.GetValues("rew")
      LockedStatus = Request.QueryString.GetValues("locked")
      If (Not Rewards Is Nothing AndAlso Not LockedStatus Is Nothing AndAlso Rewards.Length = LockedStatus.Length) Then
        For LoopCtr = 0 To Rewards.GetUpperBound(0)
          MyCommon.QueryStr = "update OfferRewards with (RowLock) set Disallow_Edit = " & LockedStatus(LoopCtr) & " " & _
                                "where RewardID=" & Rewards(LoopCtr)
          MyCommon.LRT_Execute()
        Next
      End If
    
    End If
  ElseIf (Request.QueryString("mode") = "CycleJoinDescription") Then
    MyCommon.QueryStr = "select JoinTypeID from OfferRewards with (NoLock) where RewardID=" & Request.QueryString("RewardID")
    rst = MyCommon.LRT_Select
    For Each row In rst.Rows
      If (row.Item("JoinTypeID") < 2) Then
        ' add one
        MyCommon.QueryStr = "Update OfferRewards with (RowLock) set JoinTypeID=" & row.Item("JoinTypeID") + 1 & " where RewardID=" & Request.QueryString("RewardID")
      Else
        ' set to zero
        MyCommon.QueryStr = "Update OfferRewards with (RowLock) set JoinTypeID=1 where RewardID=" & Request.QueryString("RewardID")
      End If
      MyCommon.LRT_Execute()
    Next
  ElseIf (Request.QueryString("mode") = "Delete") Then
    MyCommon.QueryStr = "update OfferRewards with (rowlock) set RewardOrder = (RewardOrder - 1) where RewardOrder > " & Request.QueryString("RewardOrder") & " and Tiered=" & deleteTiered & " and OfferID=" & OfferID
    MyCommon.LRT_Execute()
    MyCommon.QueryStr = "delete from RewardTiers with (rowlock) where RewardID=" & Request.QueryString("RewardID")
    MyCommon.LRT_Execute()
    MyCommon.QueryStr = "update OfferRewards with (RowLock) set deleted=1 where RewardID=" & Request.QueryString("RewardID") & " and OfferID=" & OfferID
    'MyCommon.QueryStr = "delete from OfferRewards with(rowlock) where RewardID=" & Request.QueryString("RewardID") & " and OfferID=" & OfferID
    MyCommon.LRT_Execute()
    'Response.Write(MyCommon.QueryStr & "<br />")
    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-removereward", LanguageID))
  ElseIf (Request.QueryString("mode") = "Up") Then
    MyCommon.QueryStr = "update OfferRewards with (RowLock) set RewardOrder=" & Request.QueryString("RewardOrder") & " where OfferID=" & OfferID & " and Tiered=" & deleteTiered & " and RewardOrder=" & Request.QueryString("RewardOrder") - 1
    MyCommon.LRT_Execute()
    MyCommon.QueryStr = "update OfferRewards with (RowLock) set RewardOrder=" & Request.QueryString("RewardOrder") - 1 & " where OfferID=" & OfferID & " and Tiered=" & deleteTiered & " and RewardID=" & Request.QueryString("RewardID")
    MyCommon.LRT_Execute()
  ElseIf (Request.QueryString("mode") = "Down") Then
    MyCommon.QueryStr = "update OfferRewards with (RowLock) set RewardOrder=" & Request.QueryString("RewardOrder") & " where OfferID=" & OfferID & " and Tiered=" & deleteTiered & " and RewardOrder=" & Request.QueryString("RewardOrder") + 1
    MyCommon.LRT_Execute()
    MyCommon.QueryStr = "update OfferRewards with (RowLock) set RewardOrder=" & Request.QueryString("RewardOrder") + 1 & " where OfferID=" & OfferID & " and Tiered=" & deleteTiered & " and RewardID=" & Request.QueryString("RewardID")
    MyCommon.LRT_Execute()
  ElseIf (Request.QueryString("addtier") <> "" Or Request.QueryString("addGlobal") <> "") Then
    'dbo.pt_OfferConditions_Insert @OfferID bigint, @ConditionTypeID int, @Tiered bit, @ConditionOrder int, @ConditionID bigint OUTPUT
    MyCommon.QueryStr = "dbo.pt_OfferRewards_Insert"
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
    If (Request.QueryString("addGlobal") <> "") Then
      iRewardType = MyCommon.Extract_Val(Request.QueryString("newrewglobal"))
      iTiered = 0
    Else
      iRewardType = MyCommon.Extract_Val(Request.QueryString("newrewtiered"))
      iTiered = 1
    End If
    MyCommon.LRTsp.Parameters.Add("@RewardTypeID", SqlDbType.Int).Value = iRewardType
    MyCommon.LRTsp.Parameters.Add("@Tiered", SqlDbType.Bit).Value = iTiered
    MyCommon.LRTsp.Parameters.Add("@RewardID", SqlDbType.BigInt).Direction = ParameterDirection.Output
    MyCommon.LRTsp.ExecuteNonQuery()
    RewardID = MyCommon.LRTsp.Parameters("@RewardID").Value
    MyCommon.Close_LRTsp()
    If iRewardType = 11 Then
      MyCommon.QueryStr = "update OfferRewards with (RowLock) set TriggerQty=1, ApplyToLimit=1, RewardLimit=1.0 where RewardID=" & RewardID
    Else
      MyCommon.QueryStr = "update OfferRewards with (RowLock) set TriggerQty=1, ApplyToLimit=1 where RewardID=" & RewardID
    End If
    MyCommon.LRT_Execute()
    
    If (iRewardType = 2 Or iRewardType = 10 Or iRewardType = 12 Or iRewardType = 13) Then
      If Request.QueryString("addGlobal") <> "" Then
        MyCommon.QueryStr = "dbo.pt_RewardTiers_Update"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@RewardID", SqlDbType.BigInt).Value = RewardID
        MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int).Value = 0
        MyCommon.LRTsp.Parameters.Add("@RewardAmount", SqlDbType.Decimal).Value = 0
        MyCommon.LRTsp.ExecuteNonQuery()
        MyCommon.Close_LRTsp()
      Else
        Integer.TryParse(NumTiers, iNumtiers)
        If iNumtiers = 0 Then
          MyCommon.QueryStr = "dbo.pt_RewardTiers_Update"
          MyCommon.Open_LRTsp()
          MyCommon.LRTsp.Parameters.Add("@RewardID", SqlDbType.BigInt).Value = RewardID
          MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int).Value = 0
          MyCommon.LRTsp.Parameters.Add("@RewardAmount", SqlDbType.Decimal).Value = 0
          MyCommon.LRTsp.ExecuteNonQuery()
          MyCommon.Close_LRTsp()
        Else
          For x = 1 To iNumtiers
            MyCommon.QueryStr = "dbo.pt_RewardTiers_Update"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@RewardID", SqlDbType.BigInt).Value = RewardID
            MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int).Value = x
            MyCommon.LRTsp.Parameters.Add("@RewardAmount", SqlDbType.Decimal).Value = 0
            MyCommon.LRTsp.ExecuteNonQuery()
            MyCommon.Close_LRTsp()
          Next
        End If
      End If
    End If
    
    ' set the default chargeback id if only one banner is assigned for this offer
    If (BannersEnabled) Then
      MyCommon.QueryStr = "select BO.BannerID, BAN.DefaultChargebackDeptID from BannerOffers BO with (NoLock) " & _
                          "inner join Banners BAN with (NoLock) on BAN.BannerID = BO.BannerID " & _
                          "where BO.OfferID = " & OfferID
      rst = MyCommon.LRT_Select
      If (rst.Rows.Count = 1) Then
        MyCommon.QueryStr = "update Discounts set ChargebackDeptID = " & MyCommon.NZ(rst.Rows(0).Item("DefaultChargebackDeptID"), 0) & " where DiscountID in " & _
                            "  (select DISC.DiscountID from OfferRewards REW with (NoLock)" & _
                            "   inner join Discounts DISC with (NoLock) on DISC.DiscountID = REW.LinkID" & _
                            "   where REW.RewardID = " & RewardID & ");"
        MyCommon.LRT_Execute()
      End If
    End If
    
    If (Request.QueryString("addtier") <> "") Then
      iRewardType = MyCommon.Extract_Val(Request.QueryString("newrewtiered"))
      ' only create if were doing discounts
      If (iRewardType = 1 Or iRewardType = 2 Or iRewardType = 10 Or iRewardType = 12 Or iRewardType = 13) Then
        For x = 1 To NumTiers
          MyCommon.QueryStr = "dbo.pt_RewardTiers_Update"
          MyCommon.Open_LRTsp()
          MyCommon.LRTsp.Parameters.Add("@RewardID", SqlDbType.BigInt).Value = RewardID
          MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int).Value = x
          MyCommon.LRTsp.Parameters.Add("@RewardAmount", SqlDbType.Decimal).Value = 0
          MyCommon.LRTsp.ExecuteNonQuery()
          MyCommon.Close_LRTsp()
        Next
      ElseIf (iRewardType = 5 Or iRewardType = 6) Then
        ' ok someone wants customer groups lets mke them
        For x = 1 To NumTiers
          MyCommon.QueryStr = "insert into RewardCustomerGroupTiers with (RowLock) (RewardID,TierLevel,CustomerGroupID) values(" & RewardID & "," & x & ",0)"
          MyCommon.LRT_Execute()
        Next
      ElseIf (iRewardType = 4) Then
        ' ok someone wants cashier groups lets mke them
        MyCommon.QueryStr = "select LinkID from OfferRewards with (NoLock) where RewardID=" & RewardID
        rst = MyCommon.LRT_Select()
        LinkID = MyCommon.NZ(rst.Rows(0).Item("LinkID"), 0)
        If (LinkID > 0) Then
          For x = 1 To NumTiers
            MyCommon.QueryStr = "insert into CashierMessageTiers with (RowLock) (MessageID,TierLevel) values(" & LinkID & "," & x & ")"
            MyCommon.LRT_Execute()
          Next
        End If
      ElseIf (iRewardType = 3) Then
        ' ok someone wants printed messages lets mke them
        MyCommon.QueryStr = "select LinkID from OfferRewards where RewardID=" & RewardID
        rst = MyCommon.LRT_Select()
        LinkID = MyCommon.NZ(rst.Rows(0).Item("LinkID"), 0)
        If (LinkID > 0) Then
          For x = 1 To NumTiers
            MyCommon.QueryStr = "insert into PrintedMessageTiers with (RowLock) (MessageID,TierLevel) values(" & LinkID & "," & x & ")"
            MyCommon.LRT_Execute()
          Next
        End If
      ElseIf (iRewardType = 7 Or iRewardType = 8 Or iRewardType = 9) Then
        ' XML Pass Thru rewards
        For x = 1 To NumTiers
          MyCommon.QueryStr = "insert into RewardXmlTiers with (RowLock) (RewardID,TierLevel) values(" & RewardID & "," & x & ")"
          MyCommon.LRT_Execute()
        Next
      End If
    End If
    If (Request.QueryString("addGlobal") <> "") Then
      iRewardType = MyCommon.Extract_Val(Request.QueryString("newrewglobal"))
      If (iRewardType = 1) Then
        Send("<script type=""text/javascript"">openPopup('offer-rew-discount.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-discount", LanguageID))
      ElseIf (iRewardType = 2) Then
        Send("<script type=""text/javascript"">openPopup('offer-rew-point.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-points", LanguageID))
      ElseIf (iRewardType = 3) Then
        Send("<script type=""text/javascript"">openPopup('offer-rew-pmsg.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-pmsg", LanguageID))
      ElseIf (iRewardType = 4) Then
        Send("<script type=""text/javascript"">openPopup('offer-rew-cmsg.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-cmsg", LanguageID))
      ElseIf (iRewardType = 5) Then
        Send("<script type=""text/javascript"">openPopup('offer-rew-membership.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-member", LanguageID))
      ElseIf (iRewardType = 6) Then
        Send("<script type=""text/javascript"">openPopup('offer-rew-membership.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-member", LanguageID))
      ElseIf (iRewardType = 7) Then
        Send("<script type=""text/javascript"">openPopup('offer-rew-xml.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, "Added global Generic XML Pass Through reward")
      ElseIf (iRewardType = 8) Then
        Send("<script type=""text/javascript"">openPopup('offer-rew-catalina.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, "Added global Catalina Coupon reward")
      ElseIf (iRewardType = 9) Then
        Send("<script type=""text/javascript"">openPopup('offer-rew-bins.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, "Added global Bin Ranges reward")
      ElseIf (iRewardType = 10) Then
        Send("<script type=""text/javascript"">openPopup('offer-rew-sv.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, "Added global Stored Value reward")
      ElseIf (iRewardType = 11) Then
        Send("<script type=""text/javascript"">openPopup('offer-rew-autogiftreceipt.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.addautogiftreceipt", LanguageID))
      ElseIf (iRewardType = 12) Then
        Send("<script type=""text/javascript"">openPopup('CM-offer-rew-advlimit.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.addadvlimitreward", LanguageID))
      ElseIf (iRewardType = 13) Then
        Send("<script type=""text/javascript"">openPopup('offer-rew-tender-point.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-points", LanguageID))
      ElseIf (iRewardType = 14) Then
        Send("<script type=""text/javascript"">openPopup('CM-offer-rew-cents-off-fuel.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.addcentsofffuel", LanguageID))
      End If
    Else
      iRewardType = MyCommon.Extract_Val(Request.QueryString("newrewtiered"))
      If (iRewardType = 1) Then
        Send("<script type=""text/javascript"">openPopup('offer-rew-discount.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("reward-add-tiered-disc", LanguageID))
      ElseIf (iRewardType = 2) Then
        Send("<script type=""text/javascript"">openPopup('offer-rew-point.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("reward-add-tiered-points", LanguageID))
      ElseIf (iRewardType = 3) Then
        Send("<script type=""text/javascript"">openPopup('offer-rew-pmsg.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("reward-add-tiered-pmsg", LanguageID))
      ElseIf (iRewardType = 4) Then
        Send("<script type=""text/javascript"">openPopup('offer-rew-cmsg.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("reward-add-tiered-cmsg", LanguageID))
      ElseIf (iRewardType = 5) Then
        Send("<script type=""text/javascript"">openPopup('offer-rew-membership.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("reward-add-tiered-membership", LanguageID))
      ElseIf (iRewardType = 6) Then
        Send("<script type=""text/javascript"">openPopup('offer-rew-membership.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("reward-add-tiered-membership", LanguageID))
      ElseIf (iRewardType = 7) Then
        Send("<script type=""text/javascript"">openPopup('offer-rew-xml.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, "Added tiered Generic XML Pass Through reward")
      ElseIf (iRewardType = 8) Then
        Send("<script type=""text/javascript"">openPopup('offer-rew-catalina.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, "Added tiered Catalina Coupon reward")
      ElseIf (iRewardType = 10) Then
        Send("<script type=""text/javascript"">openPopup('offer-rew-sv.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, "history.addtieredstoredvaluereward")
      ElseIf (iRewardType = 12) Then
        Send("<script type=""text/javascript"">openPopup('CM-offer-rew-advlimit.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.addtieredadvlimitreward", LanguageID))
      ElseIf (iRewardType = 13) Then
        Send("<script type=""text/javascript"">openPopup('offer-rew-tender-point.aspx?RewardID=" & RewardID & "&OfferID=" & OfferID & "')</script>")
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.cpeofferrew-points", LanguageID))
      End If
    End If
  End If
  
  ' decide if we need to update the status flags
  If (Request.QueryString("addtier") <> "" Or Request.QueryString("addGlobal") <> "" Or Request.QueryString("mode") = "Down" Or Request.QueryString("mode") = "up" Or Request.QueryString("mode") = "Delete" Or Request.QueryString("mode") = "CycleJoinDescription") Then
    MyCommon.QueryStr = "update OfferRewards with (RowLock) set CMOAStatusFlag=2,TCRMAStatusFlag=2 where RewardID=" & RewardID
    MyCommon.LRT_Execute()
    MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where offerid=" & OfferID
    MyCommon.LRT_Execute()
  End If
  
  If (Request.QueryString("addGlobal") <> "" Or Request.QueryString("addtier") <> "") Then
    Send("<script type=""text/javascript"">window.location=""offer-rew.aspx?OfferID=" & OfferID & """</script>")
  End If
  
  If (isTemplate Or FromTemplate) Then
    ' lets dig the permissions if its a template
    MyCommon.QueryStr = "select * from templatepermissions with (NoLock) where OfferID=" & OfferID
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      For Each row In rst.Rows
        ' ok there are some rows for the template
        Disallow_Rewards = MyCommon.NZ(row.Item("Disallow_Rewards"), True)
      Next
    End If
    NumOfCols = 8
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
    Send_Subtabs(Logix, 22, 6, , OfferID)
  Else
    Send_Subtabs(Logix, 21, 6, , OfferID)
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
</script>
<form action="offer-rew.aspx" id="mainform" name="mainform">
<input type="hidden" id="OfferID" name="OfferID" value="<%sendb(OfferID) %>" />
<input type="hidden" id="IsTemplate" name="IsTemplate" value="<% If (istemplate) Then Sendb("IsTemplate") Else Sendb("Not") %>" />
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
    <div class="box" id="rewards">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.rewards", LanguageID))%>
        </span>
      </h2>
      <% 
        Dim TiersWidth As Integer
        TiersWidth = (NumTiers * 62)
      %>
      <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.rewards", LanguageID)) %>">
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
            <th align="left" scope="col" class="th-subtype">
              <% Sendb(Copient.PhraseLib.Lookup("term.subtype", LanguageID))%>
            </th>
            <th align="left" scope="col" class="th-condition">
              <% Sendb(Copient.PhraseLib.Lookup("term.productcondition", LanguageID))%>
            </th>
            <th align="left" scope="col" class="th-details" colspan="<% If(NumTiers=0)Then Sendb("1") Else Sendb(NumTiers) %>"
              <% if(numtiers<=3)then else sendb(" style=""width:" & tierswidth & "px;""") %>>
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
            <td colspan="<% Sendb(NumTiers+NumOfCols)%>">
              <h3>
                <% Sendb(Copient.PhraseLib.Lookup("term.globalrewards", LanguageID))%>
              </h3>
            </td>
          </tr>
          <%
            ' dig out the conditions
            ' lets see if there are any global rewards to display 
            MyCommon.QueryStr = "select R.RewardID,R.Disallow_Edit,OfferID,Tiered,RewardAmountTypeID,RewardOrder,R.RewardTypeID,LinkID," & _
                                "R.JoinTypeID,R.ProductGroupID,R.ExcludedProdGroupID,TriggerQty,ApplyToLimit,ALMultiTrans,SponsorID,RewardAmountTypeID," & _
                                "UseSpecialPricing,SPRepeatAtOccur,PromoteToTransLevel,DoNotItemDistribute,PG.Name as PGName,PG.ProductGroupID As PGID,EPG.Name as EPGName,EPG.ProductGroupID As EPGID," & _
                                "RP.ProgramID,PP.ProgramName," & _
                                "RSV.ProgramID as SVProgramID,SVP.Name as SVProgramName," & _
                                "AL.LimitID as AdvLimitID,AL.Name as AdvLimitName," & _
                                "JT.PhraseID as JoinDescPhraseID,RTiers.RewardAmount,RT.Description as RewardDescription,AMT.Description as AMTDESC,AMT.PhraseID as AMTPHRASEID " & _
                                "from OfferRewards as R with (NoLock) " & _
                                "left join AmountTypes as AMT with (NoLock) on R.RewardAmountTypeID=AMT.AmountTypeID " & _
                                "left join JoinTypes as JT with (NoLock) on JT.JoinTypeID=R.JoinTypeID " & _
                                "left join ProductGroups as PG with (NoLock) on PG.ProductGroupID=R.ProductGroupID " & _
                                "left join ProductGroups as EPG with (NoLock) on EPG.ProductGroupID=R.ExcludedProdGroupID " & _
                                "left join RewardPoints as RP with (NoLock) on RP.RewardPointsID=R.LinkID " & _
                                "left join PointsPrograms as PP with (NoLock) on PP.ProgramID=RP.ProgramID " & _
                                "left join CM_RewardStoredValues as RSV with (NoLock) on RSV.RewardStoredValuesID=R.LinkID " & _
                                "left join StoredValuePrograms as SVP with (NoLock) on SVP.SVProgramID=RSV.ProgramID " & _
                                "left join CM_RewardAdvancedLimits as RAL with (NoLock) on RAL.RewardAdvLimitID=R.LinkID " & _
                                "left join CM_AdvancedLimits as AL with (NoLock) on AL.LimitID=RAL.LimitID " & _
                                "left join RewardTypes as RT with (NoLock) on RT.RewardTypeID=R.RewardTypeID " & _
                                "left join RewardTiers as RTiers with (nolock) on R.RewardID=RTiers.RewardID and RTiers.TierLevel=0 " & _
                                "where OfferID=" & OfferID & " and Tiered=0 and R.deleted=0 order by RewardOrder;"
            rst = MyCommon.LRT_Select
            prog = 1
            Dim MaxCount As Integer = rst.Rows.Count
            For Each row In rst.Rows
          %>
          <tr class="shaded">
            <td>
              <%
                'Send(FromTemplate & Disallow_Rewards & prog & "<" & MaxCount)
                If Not (isTemplate) Then
                  If (row.Item("RewardOrder") > 1 And Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Rewards)) Then
                    Sendb("<input  class=""up"" type=""button"" value=""&#9650;"" name=""up"" title=""" & Copient.PhraseLib.Lookup("term.moveup", LanguageID) & """ onClick=""LoadDocument('offer-rew.aspx?mode=Up&RewardOrder=" & row.Item("RewardOrder") & "&RewardID=" & row.Item("RewardID") & "&OfferID=" & OfferID & "')"" />")
                  Else
                    Sendb("<input class=""up"" type=""button"" value=""&#9650;"" name=""up"" title=""" & Copient.PhraseLib.Lookup("term.moveup", LanguageID) & """ disabled=""disabled"" />")
                  End If
                  If (prog < MaxCount And Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Rewards)) Then
                    Sendb("<input class=""down"" type=""button"" value=""&#9660;"" name=""down"" title=""" & Copient.PhraseLib.Lookup("term.movedown", LanguageID) & """ onClick=""LoadDocument('offer-rew.aspx?mode=Down&RewardOrder=" & row.Item("RewardOrder") & "&RewardID=" & row.Item("RewardID") & "&OfferID=" & OfferID & "')"" />")
                  Else
                    Sendb("<input class=""down"" type=""button"" value=""&#9660;"" name=""down"" title=""" & Copient.PhraseLib.Lookup("term.movedown", LanguageID) & """ disabled=""disabled"" />")
                  End If
                  prog = prog + 1
                Else
                  If (row.Item("RewardOrder") > 1 And Logix.UserRoles.EditTemplates) Then
                    Sendb("<input  class=""up"" type=""button"" value=""&#9650;"" name=""up"" title=""" & Copient.PhraseLib.Lookup("term.moveup", LanguageID) & """ onClick=""LoadDocument('offer-rew.aspx?mode=Up&RewardOrder=" & row.Item("RewardOrder") & "&RewardID=" & row.Item("RewardID") & "&OfferID=" & OfferID & "')"" />")
                  Else
                    Sendb("<input class=""up"" type=""button"" value=""&#9650;"" name=""up"" title=""" & Copient.PhraseLib.Lookup("term.moveup", LanguageID) & """ disabled=""disabled"" />")
                  End If
                  If (prog < MaxCount And Logix.UserRoles.EditTemplates) Then
                    Sendb("<input class=""down"" type=""button"" value=""&#9660;"" name=""down"" title=""" & Copient.PhraseLib.Lookup("term.movedown", LanguageID) & """ onClick=""LoadDocument('offer-rew.aspx?mode=Down&RewardOrder=" & row.Item("RewardOrder") & "&RewardID=" & row.Item("RewardID") & "&OfferID=" & OfferID & "')"" />")
                  Else
                    Sendb("<input class=""down"" type=""button"" value=""&#9660;"" name=""down"" title=""" & Copient.PhraseLib.Lookup("term.movedown", LanguageID) & """ disabled=""disabled"" />")
                  End If
                  prog = prog + 1
                End If
                  
              %>
            </td>
            <td>
              <%
                If Not isTemplate Then
                  If (Logix.UserRoles.EditOffer And Not (FromTemplate And MyCommon.NZ(row.Item("Disallow_Edit"), False))) Then
                    Sendb("<input class=""ex"" type=""button"" value=""X"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ onClick=""if(confirm('" & Copient.PhraseLib.Lookup("reward.confirmdelete", LanguageID) & "')){LoadDocument('offer-rew.aspx?mode=Delete&RewardOrder=" & row.Item("RewardOrder") & "&RewardID=" & row.Item("RewardID") & "&OfferID=" & OfferID & "')}else{return false}"" />")
                  Else
                    Sendb("<input class=""ex"" type=""button"" VALUE=""X"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("reward.confirmdelete", LanguageID) & "')){LoadDocument('offer-rew.aspx?mode=Delete&RewardOrder=" & row.Item("RewardOrder") & "&RewardID=" & row.Item("RewardID") & "&OfferID=" & OfferID & "')}else{return false}"" />")
                  End If
                Else
                  If (Logix.UserRoles.EditTemplates) Then
                    Sendb("<input class=""ex"" type=""button"" value=""X"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ onClick=""if(confirm('" & Copient.PhraseLib.Lookup("reward.confirmdelete", LanguageID) & "')){LoadDocument('offer-rew.aspx?mode=Delete&RewardOrder=" & row.Item("RewardOrder") & "&RewardID=" & row.Item("RewardID") & "&OfferID=" & OfferID & "')}else{return false}"" />")
                  Else
                    Sendb("<input class=""ex"" type=""button"" VALUE=""X"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("reward.confirmdelete", LanguageID) & "')){LoadDocument('offer-rew.aspx?mode=Delete&RewardOrder=" & row.Item("RewardOrder") & "&RewardID=" & row.Item("RewardID") & "&OfferID=" & OfferID & "')}else{return false}"" />")
                  End If
                End If
              %>
            </td>
            <td>
              <% 
                Sendb(Copient.PhraseLib.Lookup(row.Item("JoinDescPhraseID"), LanguageID))
              %>
            </td>
            <td>
              <% 
                If (row.Item("RewardTypeID") = 1) Then
                  Send("<a class=""hidden"" href=""offer-rew-discount.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                  Send("<a href=""javascript:openPopup('offer-rew-discount.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.discount", LanguageID) & "</a>")
                ElseIf (row.Item("RewardTypeID") = 2) Then
                  Send("<a class=""hidden"" href=""offer-rew-point.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                  Send("<a href=""javascript:openPopup('offer-rew-point.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.points", LanguageID) & "</a>")
                ElseIf (row.Item("RewardTypeID") = 3) Then
                  Send("<a class=""hidden"" href=""offer-rew-pmsg.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                  Send("<a href=""javascript:openPopup('offer-rew-pmsg.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.printedmessage", LanguageID) & "</a>")
                ElseIf (row.Item("RewardTypeID") = 4) Then
                  Send("<a class=""hidden"" href=""offer-rew-cmsg.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                  Send("<a href=""javascript:openPopup('offer-rew-cmsg.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.cashiermessage", LanguageID) & "</a>")
                ElseIf (row.Item("RewardTypeID") = 5) Then
                  Send("<a class=""hidden"" href=""offer-rew-membership.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                  Send("<a href=""javascript:openPopup('offer-rew-membership.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.membership", LanguageID) & "</a>")
                ElseIf (row.Item("RewardTypeID") = 6) Then
                  Send("<a class=""hidden"" href=""offer-rew-membership.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                  Send("<a href=""javascript:openPopup('offer-rew-membership.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.membership", LanguageID) & "</a>")
                ElseIf (row.Item("RewardTypeID") = 7) Then
                  Send("<a class=""hidden"" href=""offer-rew-xml.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                  Send("<a href=""javascript:openPopup('offer-rew-xml.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.xmlgeneric", LanguageID) & "</a>")
                ElseIf (row.Item("RewardTypeID") = 8) Then
                  Send("<a class=""hidden"" href=""offer-rew-catalina.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                  Send("<a href=""javascript:openPopup('offer-rew-catalina.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.xmlcatalina", LanguageID) & "</a>")
                ElseIf (row.Item("RewardTypeID") = 9) Then
                  Send("<a class=""hidden"" href=""offer-rew-bins.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                  Send("<a href=""javascript:openPopup('offer-rew-bins.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.xmlbinranges", LanguageID) & "</a>")
                ElseIf (row.Item("RewardTypeID") = 10) Then
                  Send("<a class=""hidden"" href=""offer-rew-sv.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                  Send("<a href=""javascript:openPopup('offer-rew-sv.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.storedvalue", LanguageID) & "</a>")
                ElseIf (row.Item("RewardTypeID") = 11) Then
                  Send("<a class=""hidden"" href=""offer-rew-autogiftreceipt.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                  Send("<a href=""javascript:openPopup('offer-rew-autogiftreceipt.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.autogiftreceipt", LanguageID) & "</a>")
                ElseIf (row.Item("RewardTypeID") = 12) Then
                  Send("<a class=""hidden"" href=""CM-offer-rew-advlimit.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                  Send("<a href=""javascript:openPopup('CM-offer-rew-advlimit.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.advlimit", LanguageID) & "</a>")
                ElseIf (row.Item("RewardTypeID") = 13) Then
                  Send("<a class=""hidden"" href=""offer-rew-tender-point.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                  Send("<a href=""javascript:openPopup('offer-rew-tender-point.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.tenderpoints", LanguageID) & "</a>")
                ElseIf (row.Item("RewardTypeID") = 14) Then
                  Send("<a class=""hidden"" href=""CM-offer-rew-cents-off-fuel.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                  Send("<a href=""javascript:openPopup('CM-offer-rew-cents-off-fuel.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.centsofffuel", LanguageID) & "</a>")
                End If
              %>
            </td>
            <td>
              <% 
                If (row.Item("RewardTypeID") = 5) Then
                  Send(Copient.PhraseLib.Lookup("term.add", LanguageID))
                ElseIf (row.Item("RewardTypeID") = 6) Then
                  Send(Copient.PhraseLib.Lookup("term.remove", LanguageID))
                ElseIf ((row.Item("RewardTypeID") >= 7) And (row.Item("RewardTypeID") <= 9)) Then
                  Send(Copient.PhraseLib.Lookup("term.xmlpassthru", LanguageID))
                ElseIf (row.Item("RewardTypeID") = 11) Then
                  Send(Copient.PhraseLib.Lookup("term.xmlpassthru", LanguageID))
                ElseIf (row.Item("RewardTypeID") <> 2 And row.Item("RewardTypeID") <> 10 And row.Item("RewardTypeID") <> 12 And row.Item("RewardTypeID") <> 13) Then
                  Send(Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("AMTPHRASEID"), -1), LanguageID, "") & "&nbsp;")
                Else
                  Send("&nbsp;")
                End If
              %>
            </td>
            <td>
              <% 
                If (Int(MyCommon.NZ(row.Item("PGID"), 0)) = 1) Then
                  Send(MyCommon.NZ(row.Item("PGName"), "&nbsp;") & "<br />")
                ElseIf (Int(MyCommon.NZ(row.Item("PGID"), 0)) = 0) Then
                  Send(Copient.PhraseLib.Lookup("term.entiretransaction", LanguageID) & "<br />")
                Else
                  Send("<a href=""pgroup-edit.aspx?ProductGroupID=" & row.Item("PGID") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("PGName"), "&nbsp;"), 25) & "</a><br />")
                End If
                If (MyCommon.NZ(row.Item("EPGName"), "") <> "") Then
                  If (Int(MyCommon.NZ(row.Item("EPGID"), 0)) <= 1) Then
                    Send("&nbsp;<i>" & Copient.PhraseLib.Lookup("term.excluding", LanguageID) & "</i> " & row.Item("EPGName"))
                  Else
                    Send("&nbsp;<i>" & Copient.PhraseLib.Lookup("term.excluding", LanguageID) & "</i> <a href=""pgroup-edit.aspx?ProductGroupID=" & row.Item("EPGID") & """>" & MyCommon.SplitNonSpacedString(row.Item("EPGName"), 25) & "</a>")
                  End If
                End If
              %>
            </td>
            <td colspan="<% if(NumTiers=0)then sendb("1") else sendb(NumTiers) %>">
              <%' amount
                If (row.Item("UseSpecialPricing")) Then
                  Send(Copient.PhraseLib.Lookup("term.specialpricing", LanguageID))
                ElseIf (row.Item("RewardTypeID") = 1) Then
                  If (MyCommon.NZ(row.Item("RewardAmountTypeID"), 0) <> 10) Then
                    If (MyCommon.NZ(row.Item("RewardAmountTypeID"), 0) = 2) Or (MyCommon.NZ(row.Item("RewardAmountTypeID"), 0) = 7) Or (MyCommon.NZ(row.Item("RewardAmountTypeID"), 0) = 9) Then
                    Else
                      Sendb("$")
                    End If
                    If (MyCommon.NZ(row.Item("RewardAmountTypeID"), 0) = 7) Then
                      Sendb("&nbsp;")
                    Else
                      Sendb(MyCommon.NZ(row.Item("RewardAmount"), 0))
                    End If
                    If (MyCommon.NZ(row.Item("RewardAmountTypeID"), 0) = 2) Or (MyCommon.NZ(row.Item("RewardAmountTypeID"), 0) = 9) Then
                      Send("%")
                    Else
                      Send("")
                    End If
                  End If
                ElseIf (row.Item("RewardTypeID") = 2) Then
                  Send(Int(MyCommon.NZ(row.Item("RewardAmount"), 0)) & " <a href=""point-edit.aspx?ProgramGroupID=" & MyCommon.NZ(row.Item("ProgramID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("ProgramName"), ""), 25) & "</a>")
                ElseIf (row.Item("RewardTypeID") = 13) Then
                  Send(Int(MyCommon.NZ(row.Item("RewardAmount"), 0)) & " <a href=""point-edit.aspx?ProgramGroupID=" & MyCommon.NZ(row.Item("ProgramID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("ProgramName"), ""), 25) & "</a>")
                ElseIf (row.Item("RewardTypeID") = 12) Then
                  Send(Int(MyCommon.NZ(row.Item("RewardAmount"), 0)) & " <a href=""CM-advlimit-edit.aspx?LimitID=" & MyCommon.NZ(row.Item("AdvLimitID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("AdvLimitName"), ""), 25) & "</a>")
                ElseIf (row.Item("RewardTypeID") = 10) Then
                  Dim bNeedToFormat As Boolean = False
                  If intNumDecimalPlaces > 0 Then
                    MyCommon.QueryStr = "Select SVTypeID from CM_RewardStoredValues with (NoLock) where RewardStoredValuesID=" & row.Item("LinkID") & ";"
                    rst2 = MyCommon.LRT_Select
                    If (rst2.Rows.Count > 0) Then
                      If Int(MyCommon.NZ(rst2.Rows(0).Item("SVTypeID"), 0)) = 1 Then
                        bNeedToFormat = True
                      End If
                    End If
                  End If
                  If bNeedToFormat Then
                    decTemp = (Int(MyCommon.NZ(row.Item("RewardAmount"), 0)) * 1.0) / decFactor
                    sTemp1 = FormatNumber(decTemp, intNumDecimalPlaces)
                    Send(sTemp1 & " <a href=""sv-edit.aspx?ProgramGroupID=" & MyCommon.NZ(row.Item("SVProgramID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("SVProgramName"), ""), 25) & "</a>")
                  Else
                    Send(Int(MyCommon.NZ(row.Item("RewardAmount"), 0)) & " <a href=""sv-edit.aspx?ProgramGroupID=" & MyCommon.NZ(row.Item("SVProgramID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("SVProgramName"), ""), 25) & "</a>")
                  End If
                ElseIf (row.Item("RewardTypeID") = 3) Then
                  Send("&nbsp;")
                ElseIf (row.Item("RewardTypeID") = 4) Then
                  Send("&nbsp;")
                ElseIf (row.Item("RewardTypeID") = 5 Or row.Item("RewardTypeID") = 6) Then
                  MyCommon.QueryStr = "Select RewardID,RC.CustomerGroupID,CG.Name,CG.CustomerGroupID from RewardCustomerGroupTiers as RC with (NoLock) left join CustomerGroups as CG with (NoLock) on CG.CustomerGroupID=RC.CustomerGroupID where RewardID=" & row.Item("RewardID") & " order by TierLevel"
                  rst = MyCommon.LRT_Select
                  If (rst.Rows.Count > 0) Then
                    Send("<a href=""cgroup-edit.aspx?CustomerGroupID=" & rst.Rows(0).Item("CustomerGroupID") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(rst.Rows(0).Item("Name"), ""), 25) & "</a>")
                  Else
                    Send("&nbsp;&nbsp;")
                  End If
                ElseIf (row.Item("RewardTypeID") = 8) Then
                  MyCommon.QueryStr = "Select XmlText from RewardXmlTiers with (NoLock) where RewardID=" & row.Item("RewardID") & " and TierLevel=0"
                  rst2 = MyCommon.LRT_Select
                  If (rst2.Rows.Count > 0) Then
                    Send(rst2.Rows(0).Item("XmlText"))
                  Else
                    Send("&nbsp;&nbsp;")
                  End If
                ElseIf (row.Item("RewardTypeID") = 14) Then
                  If intNumDecimalPlaces > 0 Then
                    decTemp = (Int(MyCommon.NZ(row.Item("TriggerQty"), 0)) * 1.0) / decFactor
                    sTemp1 = FormatNumber(decTemp, intNumDecimalPlaces)
                  Else
                    sTemp1 = Int(MyCommon.NZ(row.Item("TriggerQty"), 0)).ToString()
                  End If
                  decTemp = (Int(MyCommon.NZ(row.Item("ApplyToLimit"), 0)) * 1.0) / 100.0
                  sTemp2 = decTemp.ToString("0.00")
                  Send("$" & sTemp2 & " " & StrConv(Copient.PhraseLib.Lookup("term.off", LanguageID), VbStrConv.Lowercase) & " " & Copient.PhraseLib.Lookup("term.per", LanguageID) & " " & sTemp1 & " " & StrConv(Copient.PhraseLib.Lookup("term.points", LanguageID), VbStrConv.Lowercase) & " <a href=""sv-edit.aspx?ProgramGroupID=" & MyCommon.NZ(row.Item("SVProgramID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("SVProgramName"), ""), 25) & "</a>")
                Else
                  Send("&nbsp;&nbsp;")
                End If
              %>
              <% If (isTemplate) Then%>
              <td class="templine">
                <input type="checkbox" id="chkLocked1" name="chkLocked" value="<%Sendb(row.Item("RewardID"))%>"
                  <%Sendb(IIf(MyCommon.NZ(row.Item("Disallow_Edit"), False)=True, " checked=""checked""", "")) %>
                  onclick="javascript:updateLocked('lock<%Sendb(row.Item("RewardID"))%>', this.checked);" />
                <input type="hidden" id="rew<%Sendb(row.Item("RewardID"))%>" name="rew" value="<%Sendb(row.Item("RewardID"))%>" />
                <input type="hidden" id="lock<%Sendb(row.Item("RewardID"))%>" name="locked" value="<%Sendb(IIf(MyCommon.NZ(row.Item("Disallow_Edit"), False)=True, "1", "0"))%>" />
              </td>
              <%ElseIf (FromTemplate) Then%>
              <td class="templine">
                <%Sendb(IIf(MyCommon.NZ(row.Item("Disallow_Edit"), False) = True, "Yes", "No"))%>
              </td>
              <% End If%>
            </td>
          </tr>
          <%
          Next
          %>
          <% If (NumTiers > 0) Then%>
          <tr class="shadeddark">
            <td colspan="6">
              <h3>
                <% Sendb(Copient.PhraseLib.Lookup("term.tierrewards", LanguageID))%>
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
          <%
            ' lets see if there are any tiered rewards to display 
            MyCommon.QueryStr = "select R.RewardID,R.Disallow_Edit,OfferID,Tiered,RewardOrder,RewardAmountTypeID,R.RewardTypeID,LinkID,R.JoinTypeID,R.ProductGroupID,R.ExcludedProdGroupID," & _
                                " TriggerQty,ApplyToLimit,ALMultiTrans,SponsorID,RewardAmountTypeID," & _
                                " UseSpecialPricing,SPRepeatAtOccur,PromoteToTransLevel,DoNotItemDistribute,PG.Name as PGName,PG.ProductGroupID As PGID,EPG.Name as EPGName,EPG.ProductGroupID As EPGID," & _
                                " JT.PhraseID as JoinDescPhraseID,RTiers.RewardAmount,RT.Description as RewardDescription, AMT.Description as AMTDESC, AMT.PhraseID as AMTPHRASEID from OfferRewards as R  with (NoLock) " & _
                                " left join AmountTypes as AMT with (NoLock) on R.RewardAmountTypeID=AMT.AmountTypeID" & _
                                " left join JoinTypes as JT with (NoLock) on JT.JoinTypeID=R.JoinTypeID" & _
                                " left join ProductGroups as PG with (NoLock) on PG.ProductGroupID=R.ProductGroupID" & _
                                " left join ProductGroups as EPG with (NoLock) on EPG.ProductGroupID=R.ExcludedProdGroupID" & _
                                " left join RewardTypes as RT with (NoLock) on RT.RewardTypeID=R.RewardTypeID" & _
                                " left join RewardTiers as RTiers with (NoLock) on R.RewardID=RTiers.RewardID and RTiers.TierLevel=0 " & _
                                " where OfferID=" & OfferID & " and Tiered=1 and R.deleted=0 order by RewardOrder"
            rst = MyCommon.LRT_Select
            prog = 1
            For Each row In rst.Rows
          %>
          <tr class="shaded">
            <td>
              <%
                If Not isTemplate Then
                  If (row.Item("RewardOrder") > 1 And Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Rewards)) Then
                    Sendb("<input class=""up"" type=""button"" value=""&#9650;"" name=""up"" title=""" & Copient.PhraseLib.Lookup("term.moveup", LanguageID) & """ onClick=""LoadDocument('offer-rew.aspx?mode=Up&RewardOrder=" & row.Item("RewardOrder") & "&RewardID=" & row.Item("RewardID") & "&OfferID=" & OfferID & "')"" />")
                  Else
                    Sendb("<input class=""up"" type=""button"" value=""&#9650;"" name=""up"" title=""" & Copient.PhraseLib.Lookup("term.moveup", LanguageID) & """ disabled=""disabled"" />")
                  End If
                  If (prog < rst.Rows.Count And Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Rewards)) Then
                    Sendb("<input class=""down"" type=""button"" value=""&#9660;"" name=""down"" title=""" & Copient.PhraseLib.Lookup("term.movedown", LanguageID) & """ onClick=""LoadDocument('offer-rew.aspx?mode=Down&RewardOrder=" & row.Item("RewardOrder") & "&RewardID=" & row.Item("RewardID") & "&OfferID=" & OfferID & "')"" />")
                  Else
                    Sendb("<input class=""down"" type=""button"" value=""&#9660;"" name=""down"" title=""" & Copient.PhraseLib.Lookup("term.movedown", LanguageID) & """ disabled=""disabled"" />")
                  End If
                  prog = prog + 1
                Else
                  If (row.Item("RewardOrder") > 1 And Logix.UserRoles.EditTemplates) Then
                    Sendb("<input class=""up"" type=""button"" value=""&#9650;"" name=""up"" title=""" & Copient.PhraseLib.Lookup("term.moveup", LanguageID) & """ onClick=""LoadDocument('offer-rew.aspx?mode=Up&RewardOrder=" & row.Item("RewardOrder") & "&RewardID=" & row.Item("RewardID") & "&OfferID=" & OfferID & "')"" />")
                  Else
                    Sendb("<input class=""up"" type=""button"" value=""&#9650;"" name=""up"" title=""" & Copient.PhraseLib.Lookup("term.moveup", LanguageID) & """ disabled=""disabled"" />")
                  End If
                  If (prog < rst.Rows.Count And Logix.UserRoles.EditTemplates) Then
                    Sendb("<input class=""down"" type=""button"" value=""&#9660;"" name=""down"" title=""" & Copient.PhraseLib.Lookup("term.movedown", LanguageID) & """ onClick=""LoadDocument('offer-rew.aspx?mode=Down&RewardOrder=" & row.Item("RewardOrder") & "&RewardID=" & row.Item("RewardID") & "&OfferID=" & OfferID & "')"" />")
                  Else
                    Sendb("<input class=""down"" type=""button"" value=""&#9660;"" name=""down"" title=""" & Copient.PhraseLib.Lookup("term.movedown", LanguageID) & """ disabled=""disabled"" />")
                  End If
                  prog = prog + 1
                End If
                  
              %>
            </td>
            <td>
              <%
                If Not isTemplate Then
                  If (Logix.UserRoles.EditOffer And Not (FromTemplate And MyCommon.NZ(row.Item("Disallow_Edit"), False))) Then
                    Sendb("<input class=""ex"" type=""button"" value=""X"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ onClick=""if(confirm('" & Copient.PhraseLib.Lookup("reward.confirmdelete", LanguageID) & "')){LoadDocument('offer-rew.aspx?mode=Delete&RewardOrder=" & row.Item("RewardOrder") & "&RewardID=" & row.Item("RewardID") & "&OfferID=" & OfferID & "')}else{return false}"" />")
                  Else
                    Sendb("<input class=""ex"" type=""button"" value=""X"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("reward.confirmdelete", LanguageID) & "')){LoadDocument('offer-rew.aspx?mode=Delete&RewardOrder=" & row.Item("RewardOrder") & "&RewardID=" & row.Item("RewardID") & "&OfferID=" & OfferID & "')}else{return false}"" />")
                  End If
                Else
                  If (Logix.UserRoles.EditTemplates) Then
                    Sendb("<input class=""ex"" type=""button"" value=""X"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ onClick=""if(confirm('" & Copient.PhraseLib.Lookup("reward.confirmdelete", LanguageID) & "')){LoadDocument('offer-rew.aspx?mode=Delete&RewardOrder=" & row.Item("RewardOrder") & "&RewardID=" & row.Item("RewardID") & "&OfferID=" & OfferID & "')}else{return false}"" />")
                  Else
                    Sendb("<input class=""ex"" type=""button"" value=""X"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ disabled=""disabled"" onClick=""if(confirm('" & Copient.PhraseLib.Lookup("reward.confirmdelete", LanguageID) & "')){LoadDocument('offer-rew.aspx?mode=Delete&RewardOrder=" & row.Item("RewardOrder") & "&RewardID=" & row.Item("RewardID") & "&OfferID=" & OfferID & "')}else{return false}"" />")
                  End If
                End If
              %>
            </td>
            <td>
              <% 
                Sendb(Copient.PhraseLib.Lookup(row.Item("JoinDescPhraseID"), LanguageID))
              %>
            </td>
            <td>
              <% 
                If (row.Item("RewardTypeID") = 1) Then
                  Send("<a class=""hidden"" href=""offer-rew-discount.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                  Send("<a href=""javascript:openPopup('offer-rew-discount.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.discount", LanguageID) & "</a>")
                ElseIf (row.Item("RewardTypeID") = 2) Then
                  Send("<a class=""hidden"" href=""offer-rew-point.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                  Send("<a href=""javascript:openPopup('offer-rew-point.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.points", LanguageID) & "</a>")
                ElseIf (row.Item("RewardTypeID") = 3) Then
                  Send("<a class=""hidden"" href=""offer-rew-pmsg.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                  Send("<a href=""javascript:openPopup('offer-rew-pmsg.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.printedmessage", LanguageID) & "</a>")
                ElseIf (row.Item("RewardTypeID") = 4) Then
                  Send("<a class=""hidden"" href=""offer-rew-cmsg.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                  Send("<a href=""javascript:openPopup('offer-rew-cmsg.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.cashiermessage", LanguageID) & "</a>")
                ElseIf (row.Item("RewardTypeID") = 5) Then
                  Send("<a class=""hidden"" href=""offer-rew-membership.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                  Send("<a href=""javascript:openPopup('offer-rew-membership.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.membership", LanguageID) & "</a>")
                ElseIf (row.Item("RewardTypeID") = 6) Then
                  Send("<a class=""hidden"" href=""offer-rew-membership.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                  Send("<a href=""javascript:openPopup('offer-rew-membership.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.membership", LanguageID) & "</a>")
                ElseIf (row.Item("RewardTypeID") = 7) Then
                  Send("<a class=""hidden"" href=""offer-rew-xml.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                  Send("<a href=""javascript:openPopup('offer-rew-xml.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.xmlgeneric", LanguageID) & "</a>")
                ElseIf (row.Item("RewardTypeID") = 8) Then
                  Send("<a class=""hidden"" href=""offer-rew-catalina.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                  Send("<a href=""javascript:openPopup('offer-rew-catalina.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.xmlcatalina", LanguageID) & "</a>")
                ElseIf (row.Item("RewardTypeID") = 10) Then
                  Send("<a class=""hidden"" href=""offer-rew-sv.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                  Send("<a href=""javascript:openPopup('offer-rew-sv.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.storedvalue", LanguageID) & "</a>")
                ElseIf (row.Item("RewardTypeID") = 11) Then
                  Send("<a class=""hidden"" href=""offer-rew-autogiftreceipt.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                  Send("<a href=""javascript:openPopup('offer-rew-autogiftreceipt.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.autogiftreceipt", LanguageID) & "</a>")
                ElseIf (row.Item("RewardTypeID") = 12) Then
                  Send("<a class=""hidden"" href=""CM-offer-rew-advlimit.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                  Send("<a href=""javascript:openPopup('CM-offer-rew-advlimit.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.advlimit", LanguageID) & "</a>")
                ElseIf (row.Item("RewardTypeID") = 13) Then
                  Send("<a class=""hidden"" href=""offer-rew-tender-point.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & """>►</a>")
                  Send("<a href=""javascript:openPopup('offer-rew-tender-point.aspx?RewardID=" & row.Item("RewardID") & "&OfferID=" & row.Item("OfferID") & "')"">" & Copient.PhraseLib.Lookup("term.tenderpoints", LanguageID) & "</a>")
                End If
              %>
            </td>
            <td>
              <% 
                If (row.Item("RewardTypeID") = 5) Then
                  Send(Copient.PhraseLib.Lookup("term.add", LanguageID))
                ElseIf (row.Item("RewardTypeID") = 6) Then
                  Send(Copient.PhraseLib.Lookup("term.remove", LanguageID))
                ElseIf ((row.Item("RewardTypeID") >= 7) And (row.Item("RewardTypeID") <= 9)) Then
                  Send(Copient.PhraseLib.Lookup("term.xmlpassthru", LanguageID))
                ElseIf (row.Item("RewardTypeID") = 11) Then
                  Send(Copient.PhraseLib.Lookup("term.xmlpassthru", LanguageID))
                ElseIf (row.Item("RewardTypeID") <> 2 And row.Item("RewardTypeID") <> 10 And row.Item("RewardTypeID") <> 13) Then
                  Send(Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("AMTPHRASEID"), -1), LanguageID, "") & "&nbsp;")
                Else
                  Send("&nbsp;")
                End If
              %>
            </td>
            <td>
              <% 
                If (Int(MyCommon.NZ(row.Item("PGID"), 0)) = 1) Then
                  Send(MyCommon.NZ(row.Item("PGName"), "&nbsp;") & "<br />")
                ElseIf (Int(MyCommon.NZ(row.Item("PGID"), 0)) = 0) Then
                  Send(Copient.PhraseLib.Lookup("term.entiretransaction", LanguageID) & "<br />")
                Else
                  Send("<a href=""pgroup-edit.aspx?ProductGroupID=" & row.Item("PGID") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("PGName"), "&nbsp;"), 25) & "</a><br />")
                End If
                If (MyCommon.NZ(row.Item("EPGName"), "") <> "") Then
                  If (Int(MyCommon.NZ(row.Item("EPGID"), 0)) <= 1) Then
                    Send("&nbsp;<i>" & Copient.PhraseLib.Lookup("term.excluding", LanguageID) & "</i> " & row.Item("EPGName"))
                  Else
                    Send("&nbsp;<i>" & Copient.PhraseLib.Lookup("term.excluding", LanguageID) & "</i> <a href=""pgroup-edit.aspx?ProductGroupID=" & row.Item("EPGID") & """>" & MyCommon.SplitNonSpacedString(row.Item("EPGName"), 25) & "</a>")
                  End If
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
              If (row.Item("RewardTypeID") = 1) Then
                MyCommon.QueryStr = "Select RewardAmount from RewardTiers with (NoLock) where RewardID=" & row.Item("RewardID")
                rst2 = MyCommon.LRT_Select
                For Each row2 In rst2.Rows
                  Sendb("<td>")
                  If (MyCommon.NZ(row.Item("RewardAmountTypeID"), 0) <> 10) Then
                    If (MyCommon.NZ(row.Item("RewardAmountTypeID"), 0) = 2) Or (MyCommon.NZ(row.Item("RewardAmountTypeID"), 0) = 7) Or (MyCommon.NZ(row.Item("RewardAmountTypeID"), 0) = 9) Then
                    Else
                      Sendb("$")
                    End If
                    If (MyCommon.NZ(row.Item("RewardAmountTypeID"), 0) = 7) Then
                      Sendb("&nbsp;")
                    Else
                      Sendb(MyCommon.NZ(row2.Item("RewardAmount"), 0))
                    End If
                    If (MyCommon.NZ(row.Item("RewardAmountTypeID"), 0) = 2) Or (MyCommon.NZ(row.Item("RewardAmountTypeID"), 0) = 9) Then
                      Send("%")
                    Else
                      Send("")
                    End If
                  End If
                  Send("</td>")
                Next
              ElseIf (row.Item("RewardTypeID") = 2 Or row.Item("RewardTypeID") = 10 Or row.Item("RewardTypeID") = 12 Or row.Item("RewardTypeID") = 13) Then
                Dim bNeedToFormat As Boolean = False
                If row.Item("RewardTypeID") = 10 Then
                  If intNumDecimalPlaces > 0 Then
                    MyCommon.QueryStr = "Select SVTypeID from CM_RewardStoredValues with (NoLock) where RewardStoredValuesID=" & row.Item("LinkID") & ";"
                    rst2 = MyCommon.LRT_Select
                    If (rst2.Rows.Count > 0) Then
                      If Int(MyCommon.NZ(rst2.Rows(0).Item("SVTypeID"), 0)) = 1 Then
                        bNeedToFormat = True
                      End If
                    End If
                  End If
                End If
                MyCommon.QueryStr = "Select RewardAmount from RewardTiers with (NoLock) where RewardID=" & row.Item("RewardID")
                rst2 = MyCommon.LRT_Select
                For Each row2 In rst2.Rows
                  If bNeedToFormat Then
                    decTemp = (Int(MyCommon.NZ(row2.Item("RewardAmount"), 0)) * 1.0) / decFactor
                    sTemp1 = FormatNumber(decTemp, intNumDecimalPlaces)
                    Send("<td>" & sTemp1 & "</td>")
                  Else
                    Send("<td>" & Int(row2.Item("RewardAmount")) & "</td>")
                  End If
                Next
              ElseIf (row.Item("RewardTypeID") = 3) Then
                Send("<td colspan=""" & colSpanAmt & """>&nbsp;</td>")
              ElseIf (row.Item("RewardTypeID") = 4) Then
                Send("<td colspan=""" & colSpanAmt & """>&nbsp;</td>")
              ElseIf (row.Item("RewardTypeID") = 5 Or row.Item("RewardTypeID") = 6) Then
                MyCommon.QueryStr = "Select RewardID,RC.CustomerGroupID,CG.Name as Name,CG.CustomerGroupID as CGID from RewardCustomerGroupTiers as RC with (NoLock) left join CustomerGroups as CG with (NoLock) on CG.CustomerGroupID=RC.CustomerGroupID where RewardID=" & row.Item("RewardID") & " order by TierLevel"
                rst2 = MyCommon.LRT_Select
                If (rst2.Rows.Count > 0) Then
                  For Each row2 In rst2.Rows
                    Send("<td><a href=""cgroup-edit.aspx?CustomerGroupID=" & MyCommon.NZ(row2.Item("CGID"), "") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row2.Item("Name"), ""), 25) & "</a></td>")
                  Next
                End If
              ElseIf (row.Item("RewardTypeID") = 8) Then
                MyCommon.QueryStr = "Select XmlText from RewardXmlTiers with (NoLock) where RewardID=" & row.Item("RewardID")
                rst2 = MyCommon.LRT_Select
                For Each row2 In rst2.Rows
                  Send("<td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row2.Item("XmlText"), ""), 25) & "</td>")
                Next
              Else
                Send("<td colspan=""" & colSpanAmt & """>&nbsp;</td>")
              End If
            %>
            <% If (isTemplate) Then%>
            <td class="templine">
              <input type="checkbox" id="chkLocked2" name="chkLocked" value="<%Sendb(row.Item("RewardID"))%>"
                <%Sendb(IIf(MyCommon.NZ(row.Item("Disallow_Edit"), False)=True, " checked=""checked""", "")) %>
                onclick="javascript:updateLocked('lock<%Sendb(row.Item("RewardID"))%>', this.checked);" />
              <input type="hidden" id="rew<%Sendb(row.Item("RewardID"))%>" name="rew" value="<%Sendb(row.Item("RewardID"))%>" />
              <input type="hidden" id="lock<%Sendb(row.Item("RewardID"))%>" name="locked" value="<%Sendb(IIf(MyCommon.NZ(row.Item("Disallow_Edit"), False)=True, "1", "0"))%>" />
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
      <hr class="hidden" />
    </div>
    <div class="box" id="newreward">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("offer-rew.addreward", LanguageID))%>
        </span>
      </h2>
      <%
        If isTemplate Then
          Send("<span class=""temp"">")
          Send("  <input type=""checkbox"" class=""tempcheck"" id=""temp-rewards"" name=""Disallow_Rewards""" & IIf(Disallow_Rewards, " checked=""checked""", "") & " />")
          Send("  <label for=""temp-rewards"">" & Copient.PhraseLib.Lookup("term.locked", LanguageID) & "</label>")
          Send("</span>")
        End If
        MyCommon.QueryStr = "SELECT PECT.EngineID, PECT.ComponentTypeID, RT.RewardTypeID, RT.Description, RT.PhraseID, PECT.Singular, PECT.Tierable, PECT.Touchpoint," & _
                            "  CASE RewardTypeID" & _
                            "    WHEN 1 THEN (SELECT COUNT(*) FROM OfferRewards WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and RewardTypeID=1) " & _
                            "    WHEN 2 THEN (SELECT COUNT(*) FROM OfferRewards WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and RewardTypeID=2) " & _
                            "    WHEN 3 THEN (SELECT COUNT(*) FROM OfferRewards WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and RewardTypeID=3) " & _
                            "    WHEN 4 THEN (SELECT COUNT(*) FROM OfferRewards WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and RewardTypeID=4) " & _
                            "    WHEN 5 THEN (SELECT COUNT(*) FROM OfferRewards WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and RewardTypeID=5) " & _
                            "    WHEN 6 THEN (SELECT COUNT(*) FROM OfferRewards WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and RewardTypeID=6) " & _
                            "    WHEN 7 THEN (SELECT COUNT(*) FROM OfferRewards WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and RewardTypeID=7) " & _
                            "    WHEN 8 THEN (SELECT COUNT(*) FROM OfferRewards WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and RewardTypeID=8) " & _
                            "    WHEN 9 THEN (SELECT COUNT(*) FROM OfferRewards WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and RewardTypeID=9) " & _
                            "    WHEN 10 THEN (SELECT COUNT(*) FROM OfferRewards WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and RewardTypeID=10) " & _
                            "    WHEN 11 THEN (SELECT COUNT(*) FROM OfferRewards WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and RewardTypeID=11) " & _
                            "    WHEN 12 THEN (SELECT COUNT(*) FROM OfferRewards WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and RewardTypeID=12) " & _
                            "    WHEN 13 THEN (SELECT COUNT(*) FROM OfferRewards WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and RewardTypeID=13) " & _
                            "    WHEN 14 THEN (SELECT COUNT(*) FROM OfferRewards WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and RewardTypeID=14) " & _
                            "    ELSE 0 " & _
                            "  END as ItemCount " & _
                            "FROM PromoEngineComponentTypes AS PECT " & _
                            "INNER JOIN RewardTypes AS RT ON RT.RewardTypeID=PECT.LinkID " & _
                            "WHERE EngineID=0 And ((RT.Enabled is null) or (RT.Enabled=1)) And PECT.ComponentTypeID=2 And PECT.Enabled=1;"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
          Send("<label for=""newrewglobal"">" & Copient.PhraseLib.Lookup("offer-rew.addglobal", LanguageID) & "</label>")
          Send("<br />")
          Send("<select class=""medium"" id=""newrewglobal"" name=""newrewglobal""" & IIf(FromTemplate And Disallow_Rewards, " disabled=""disabled""", "") & ">")
          For Each row In rst.Rows
            If (row.Item("Singular") = False) OrElse (row.Item("Singular") = True AndAlso row.Item("ItemCount") = 0) Then
                If (bBuckChildOffer AndAlso row.Item("ItemCount") > 0) Then
                  Send("<option value=""" & row.Item("RewardTypeID") & """ disabled>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "</option>")
                Else
                  If Not (bStoreUser And (row.Item("RewardTypeID") = "2" Or row.Item("RewardTypeID") = "10")) Then
                    Send("<option value=""" & row.Item("RewardTypeID") & """>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "</option>")
                  End If
                End If
            End If
          Next
          Send("</select>")
          If Not isTemplate Then
            Send("<input type=""submit"" class=""regular"" id=""addglobal"" name=""addglobal""" & IIf(Not Logix.UserRoles.EditOffer, " disabled=""disabled""", "") & IIf(FromTemplate And Disallow_Rewards, " disabled=""disabled""", "") & " value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """ /><br />")
          Else
            Send("<input type=""submit"" class=""regular"" id=""addglobal"" name=""addglobal""" & IIf(Not Logix.UserRoles.EditTemplates, " disabled=""disabled""", "") & " value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """ /><br />")
          End If
            
        End If
          
        If NumTiers > 0 Then
          MyCommon.QueryStr = "SELECT PECT.EngineID, PECT.ComponentTypeID, RT.RewardTypeID, RT.Description, RT.PhraseID, PECT.Singular, PECT.Tierable, PECT.Touchpoint," & _
                              "  CASE RewardTypeID" & _
                              "    WHEN 1 THEN (SELECT COUNT(*) FROM OfferRewards WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and RewardTypeID=1) " & _
                              "    WHEN 2 THEN (SELECT COUNT(*) FROM OfferRewards WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and RewardTypeID=2) " & _
                              "    WHEN 3 THEN (SELECT COUNT(*) FROM OfferRewards WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and RewardTypeID=3) " & _
                              "    WHEN 4 THEN (SELECT COUNT(*) FROM OfferRewards WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and RewardTypeID=4) " & _
                              "    WHEN 5 THEN (SELECT COUNT(*) FROM OfferRewards WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and RewardTypeID=5) " & _
                              "    WHEN 6 THEN (SELECT COUNT(*) FROM OfferRewards WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and RewardTypeID=6) " & _
                              "    WHEN 7 THEN (SELECT COUNT(*) FROM OfferRewards WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and RewardTypeID=7) " & _
                              "    WHEN 8 THEN (SELECT COUNT(*) FROM OfferRewards WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and RewardTypeID=8) " & _
                              "    WHEN 9 THEN (SELECT COUNT(*) FROM OfferRewards WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and RewardTypeID=9) " & _
                              "    WHEN 10 THEN (SELECT COUNT(*) FROM OfferRewards WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and RewardTypeID=10) " & _
                              "    WHEN 11 THEN (SELECT COUNT(*) FROM OfferRewards WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and RewardTypeID=11) " & _
                              "    WHEN 12 THEN (SELECT COUNT(*) FROM OfferRewards WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and RewardTypeID=12) " & _
                              "    WHEN 13 THEN (SELECT COUNT(*) FROM OfferRewards WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and RewardTypeID=13) " & _
                              "    WHEN 14 THEN (SELECT COUNT(*) FROM OfferRewards WITH (NOLOCK) where OfferID=" & OfferID & " and Deleted=0 and RewardTypeID=14) " & _
                              "    ELSE 0 " & _
                              "  END as ItemCount " & _
                              "FROM PromoEngineComponentTypes AS PECT " & _
                              "INNER JOIN RewardTypes AS RT ON RT.RewardTypeID=PECT.LinkID " & _
                              "WHERE EngineID=0 And ((RT.Enabled is null) or (RT.Enabled=1)) And PECT.ComponentTypeID=2 And PECT.Enabled=1 and PECT.Tierable=1;"
          rst = MyCommon.LRT_Select
          If rst.Rows.Count > 0 Then
            Send("<label for=""newrewtiered"">" & Copient.PhraseLib.Lookup("offer-rew.addtiered", LanguageID) & ":</label>")
            Send("<br />")
            Send("<select class=""medium"" id=""newrewtiered"" name=""newrewtiered""" & IIf(FromTemplate And Disallow_Rewards, " disabled=""disabled""", "") & ">")
            For Each row In rst.Rows
              If (row.Item("Singular") = False) OrElse (row.Item("Singular") = True AndAlso row.Item("ItemCount") = 0) Then
                  If (bBuckParentOffer AndAlso row.Item("ItemCount") > 0) Then
                    Send("<option value=""" & row.Item("RewardTypeID") & """ disabled>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "</option>")
                  Else
                    Send("<option value=""" & row.Item("RewardTypeID") & """>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "</option>")
                  End If
              End If
            Next
            Send("</select>")
            If Not isTemplate Then
              Send("<input type=""submit"" class=""regular"" id=""addtier"" name=""addtier""" & IIf(Not Logix.UserRoles.EditOffer, " disabled=""disabled""", "") & IIf(FromTemplate And Disallow_Rewards, " disabled=""disabled""", "") & " value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """ /><br />")
            Else
              Send("<input type=""submit"" class=""regular"" id=""addtier"" name=""addtier""" & IIf(Not Logix.UserRoles.EditTemplates, " disabled=""disabled""", "") & " value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """ /><br />")
            End If
              
          End If
        End If
      %>
    </div>
  </div>
  <br clear="all" />
</div>
</form>
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
