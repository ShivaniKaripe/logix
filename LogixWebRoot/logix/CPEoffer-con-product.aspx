<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:5.99.1.72284.Unstable Build - JLWVVIQQ-PC %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: CPEoffer-con-product.aspx 
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
  ' * Version : 5.99.1.72284 
  ' *
  ' *****************************************************************************
  
  Dim CopientFileName As String = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
  Dim CopientFileVersion As String = "5.99.1.72284" 
  Dim CopientProject As String = "Copient Logix"
  Dim CopientNotes As String = ""
  
  Dim AdminUserID As Long
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim rst As DataTable
  Dim dt As DataTable
  Dim row As DataRow
  Dim OfferID As Long
  Dim Name As String = ""
  Dim ConditionID As String
  Dim isTemplate As Boolean
  Dim FromTemplate As Boolean
  Dim Disallow_Edit As Boolean = True
  Dim Disqualifier As Boolean = False
  Dim DisabledAttribute As String = ""
  Dim roid As Integer
  Dim Ids() As String
  Dim i, ProductGroups As Integer
  Dim historyString
  Dim CloseAfterSave As Boolean = False
  Dim Qty As Decimal
  Dim Type As Integer
  Dim AccumMin As Decimal
  Dim AccumLimit As Decimal
  Dim AccumPeriod As Integer
  Dim AccumEligible As Boolean
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim RequirePG As Boolean = False
  Dim HasRequiredPG As Boolean = False
  Dim EngineID As Integer = 2
  Dim BannersEnabled As Boolean = True
  Dim HasTenderCondition As Boolean = False
  Dim HasDisqualifier As Boolean = False
  Dim UniqueProduct As Integer = 0
  Dim IsUniqueProd As Boolean = False
  Dim TierLevels As Integer = 1
  Dim t As Integer = 1
  Dim TierQty As Decimal
  Dim ValidTier As Boolean = False
  Dim pg As Integer
  Dim ExitFor As Boolean = False
  Dim ProdID As Integer = 0
  Dim ExProdID As Integer = 0
  Dim IncentiveProdGroupID As Integer = 0
  Dim rst2 As DataTable
  Dim rst3 As DataTable
  Dim UniqueChecked As Boolean = False
  Dim IsItem As Boolean = False
  Dim IsDollar As Boolean = False
  Dim IsQty1 As Boolean = False
  Dim CleanGroupName As String = ""
  Dim Shaded As String = " class=""shaded"""
  Dim TierDT As DataTable
  Dim row3 As DataRow
  Dim ProductComboID As Integer = 1
  Dim limits As String = ""
  Dim ShowAccum, ValidAccum As Boolean
  Dim EnableAccum As Boolean = False
  Dim t1, t2 As Decimal
  Dim AnyProductUsed As Boolean = False
  Dim Rounding As Boolean = False
  Dim ValidRounding As Boolean = True
  Dim MinPurchAmt As Decimal = 0
  Dim HasAnyCustomer As Boolean
  Dim HasBundleDiscount As Boolean = False
  Dim RECORD_LIMIT As Integer = GroupRecordLimit '500
  Dim RestrictExcluding As Boolean = False
  Dim isSingleItemPriceThreshold As Boolean = False
  Dim isWtVol As Boolean = False
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "CPEoffer-con-product.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  OfferID = Request.QueryString("OfferID")
  ConditionID = Request.QueryString("ConditionID")
  Disqualifier = IIf(Request.QueryString("Disqualifier") = "1", True, False)
  HasAnyCustomer = CPEOffer_Has_AnyCustomer(MyCommon, OfferID)
   
  If Request.QueryString("IncentiveProductGroupID") <> "" Then
    IncentiveProdGroupID = Request.QueryString("IncentiveProductGroupID")
  Else
    IncentiveProdGroupID = 0
  End If
  
  If (Request.QueryString("EngineID") <> "") Then
    EngineID = MyCommon.Extract_Val(Request.QueryString("EngineID"))
  Else
    MyCommon.QueryStr = "select EngineID from OfferIDs with (NoLock) where OfferID=" & OfferID & ";"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
      EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), 0)
    End If
  End If
  
  MyCommon.QueryStr = "select RewardOptionID,TierLevels from CPE_RewardOptions with (NoLock) where IncentiveID=" & OfferID & " and TouchResponse=0 and Deleted=0;"
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    roid = rst.Rows(0).Item("RewardOptionID")
    TierLevels = rst.Rows(0).Item("TierLevels")
  Else
    infoMessage = Copient.PhraseLib.Lookup("term.errornorewardoption", LanguageID)
  End If
  
  MyCommon.QueryStr = "select IncentiveProductGroupID from CPE_IncentiveProductGroups with (NoLock) " & _
                      "where RewardOptionID=" & roid & " and ProductGroupID=1 and Deleted=0;"
  dt = MyCommon.LRT_Select
  If dt.Rows.Count > 0 Then
    AnyProductUsed = True
  End If
  
  'Determine if the offer has a group-level conditional (bundle) discount
  MyCommon.QueryStr = "select DiscountID from CPE_Discounts as DIS with (NoLock) " & _
                      "inner join CPE_Deliverables as DEL on DEL.OutputID=DIS.DiscountID " & _
                      "where DIS.DiscountTypeID=4 and DEL.RewardOptionID=" & roid & " and DEL.Deleted=0 and DIS.Deleted=0;"
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    HasBundleDiscount = True
  End If
  
  'Find the product groups for this ROID
  MyCommon.QueryStr = "select IncentiveProductGroupID from CPE_IncentiveProductGroups with (NoLock) where Deleted=0 and ExcludedProducts=0 and RewardOptionID=" & roid & ";"
  dt = MyCommon.LRT_Select()
  If dt.Rows.Count > 0 Then
    ProductGroups = dt.Rows.Count
    If ProductGroups >= 2 Then
      'There's more than one group, so don't allow accumulation
      ShowAccum = False
    ElseIf ProductGroups = 1 Then
      'There's one group, and this is it, so allow accumulation
      If IncentiveProdGroupID = dt.Rows(0).Item("IncentiveProductGroupID") Then
        ShowAccum = True
      Else
        ShowAccum = False
      End If
    End If
  Else
    'There are no groups so allow accumulation
    ProductGroups = 0
    ShowAccum = True
  End If
  
  RestrictExcluding = IsExcludingRestricted(MyCommon, roid)
  
  If IncentiveProdGroupID = 0 Then
    ProdID = -1
    ExProdID = -1
    If Request.QueryString("EnableAccum") = "1" Then
      EnableAccum = True
    ElseIf Request.QueryString("EnableAccum") = "0" Then
      EnableAccum = False
    End If
  Else
    MyCommon.QueryStr = "select ProductGroupID,AccumMin from CPE_IncentiveProductGroups with (NoLock) where RewardOptionID=" & roid & " and ExcludedProducts=0 and Deleted=0 and IncentiveProductGroupID=" & IncentiveProdGroupID & ";"
    rst = MyCommon.LRT_Select()
    If rst.Rows.Count > 0 Then
      ProdID = MyCommon.NZ(rst.Rows(0).Item("ProductGroupID"), -1)
      ExProdID = -1
      If MyCommon.NZ(rst.Rows(0).Item("AccumMin"), 0) > 0 Then
        EnableAccum = True
      Else
        EnableAccum = False
      End If
    Else
      MyCommon.QueryStr = "select ProductGroupID, AccumMin from CPE_IncentiveProductGroups with (NoLock) where RewardOptionID=" & roid & " and ExcludedProducts=1 and IncentiveProductGroupID=" & IncentiveProdGroupID & ";"
      dt = MyCommon.LRT_Select
      If dt.Rows.Count > 0 Then
        ProdID = 1
        ExProdID = dt.Rows(0).Item("ProductGroupID")
        If MyCommon.NZ(dt.Rows(0).Item("AccumMin"), 0) > 0 Then
          EnableAccum = True
        Else
          EnableAccum = False
        End If
      Else
        EnableAccum = True
      End If
    End If
    If Request.QueryString("EnableAccum") = "1" Then
      EnableAccum = True
    ElseIf Request.QueryString("EnableAccum") = "0" Then
      EnableAccum = False
    End If
  End If
  
  If IncentiveProdGroupID > 0 Then
    MyCommon.QueryStr = "select Rounding from CPE_IncentiveProductGroups where RewardOptionID=" & roid & " and IncentiveProductGroupID=" & IncentiveProdGroupID & ";"
    dt = MyCommon.LRT_Select
    If dt.Rows.Count > 0 Then
      If dt.Rows(0).Item("Rounding") = "1" Then
        Rounding = True
      End If
    End If
  End If
  
  ' see if someone is saving
  If (Request.QueryString("save") <> "" And roid > 0) Then
    'Tier level validation code
    If TierLevels > 1 Then
      For t = 2 To TierLevels
        t2 = MyCommon.Extract_Val(Request.QueryString("t" & t & "_limit"))
        t1 = MyCommon.Extract_Val(Request.QueryString("t" & t - 1 & "_limit"))
        If t2 > t1 Then
          ValidTier = True
        Else
          ValidTier = False
          Exit For
        End If
      Next
    Else
      ValidTier = True
    End If
    If Disqualifier Then
      ValidTier = True
    End If
    'Rounding validation code
    If Request.QueryString("Rounding") = "on" Then
      Rounding = True
    Else
      Rounding = False
    End If
    If Rounding Then
      For t = 1 To TierLevels
        If CInt(MyCommon.Extract_Val(Request.QueryString("t" & t & "_limit"))) <> MyCommon.Extract_Val(Request.QueryString("t" & t & "_limit")) Then
          ValidRounding = False
        End If
      Next
    End If
    'Accumulation validation code
    If Request.QueryString("EnableAccum") = "" Then
      EnableAccum = False
    End If
    ValidAccum = False   'Set default for the valid accumulation
    If EnableAccum Then
      AccumMin = IIf(MyCommon.Extract_Val(Request.QueryString("accummin")) <> "", MyCommon.Extract_Val(Request.QueryString("accummin")), 0)
      If AccumMin > 0 Then
        ValidAccum = True
      End If
    Else
      ValidAccum = True
    End If
    
    If roid > 0 And ValidTier And ValidAccum And ValidRounding Then
      If (Not IsValidEntry(MyCommon)) Then
        infoMessage = Copient.PhraseLib.Lookup("term.invalidnumericentry", LanguageID)
      Else
        If Request.QueryString("selGroups") <> "" Then
          ProdID = MyCommon.Extract_Val(Request.QueryString("selGroups"))
        Else
          ProdID = -1
        End If
        
        ' Check to see if a product condition is required by the template, if applicable
        MyCommon.QueryStr = "select ProductGroupID from CPE_IncentiveProductGroups with (NoLock) where RewardOptionID=" & roid & _
                            " and RequiredFromTemplate=1 and Deleted=0 and ExcludedProducts=0;"
        rst = MyCommon.LRT_Select
        HasRequiredPG = (rst.Rows.Count > 0)
        
        MyCommon.QueryStr = "select IncentiveTenderID from CPE_IncentiveTenderTypes with (NoLock) where Deleted = 0 and RewardOptionID = " & roid
        rst = MyCommon.LRT_Select
        HasTenderCondition = (rst.Rows.Count > 0)
        
        MyCommon.QueryStr = "select IPG.IncentiveProductGroupID, PG.ProductGroupID, PG.Name, PG.PhraseID, PG.AnyProduct, UT.PhraseID as UnitPhraseID, " & _
                            " ExcludedProducts, ProductComboID, QtyForIncentive, QtyUnitType, AccumMin, AccumLimit, AccumPeriod, DisallowEdit, " & _
                            " RequiredFromTemplate, Disqualifier, Rounding, MinPurchAmt from CPE_IncentiveProductGroups as IPG with (NoLock) " & _
                            " left join ProductGroups as PG with (NoLock) on PG.ProductGroupID=IPG.ProductGroupID " & _
                            " left join CPE_RewardOptions as RO with (NoLock) on IPG.RewardOptionID=RO.RewardOptionID " & _
                            " left join CPE_UnitTypes as UT with (NoLock) on UT.UnitTypeID=IPG.QtyUnitType " & _
                            " where IPG.RewardOptionID=" & roid & "and IPG.Deleted=0 and Disqualifier=1 " & _
                            " order by Name;"
        rst = MyCommon.LRT_Select
        HasDisqualifier = (rst.Rows.Count > 0)
        
        If (HasTenderCondition AndAlso Request.QueryString("selGroups") = "") Then
          infoMessage = Copient.PhraseLib.Lookup("ueoffer-con-product.TenderRequirement", LanguageID)
        ElseIf (HasDisqualifier AndAlso Request.QueryString("selGroups") = "") And Not Disqualifier Then
          infoMessage = Copient.PhraseLib.Lookup("ueoffer-con-product.ProdDisqualierRequirement", LanguageID)
        Else
          'Find IncentiveProductGroupID and use it to delete the records from the tiers table
          If IncentiveProdGroupID <> 0 Then
            MyCommon.QueryStr = "delete from CPE_IncentiveProductGroupTiers where IncentiveProductGroupID=" & IncentiveProdGroupID
            MyCommon.LRT_Execute()
          End If
        End If
        
        ' we need to do some work to set the limit values if there are any, otherwise just set to 0
        ' in theory there should be one set of limit values for each selected groups and possibly an accumulation infos        
        If Not EnableAccum OrElse Disqualifier Then
          AccumMin = 0
          AccumLimit = 0
          AccumPeriod = 0
        Else
          AccumMin = IIf(MyCommon.Extract_Val(Request.QueryString("accummin")) <> "", MyCommon.Extract_Val(Request.QueryString("accummin")), 0)
          AccumLimit = IIf(MyCommon.Extract_Val(Request.QueryString("accumlimit")) <> "", MyCommon.Extract_Val(Request.QueryString("accumlimit")), 0)
          AccumPeriod = IIf(MyCommon.Extract_Val(Request.QueryString("accumperiod")) <> "", MyCommon.Extract_Val(Request.QueryString("accumperiod")), 0)
        End If
        Qty = IIf(MyCommon.Extract_Val(Request.QueryString("t1_limit")) <> "", MyCommon.Extract_Val(Request.QueryString("t1_limit")), 0)
        Type = IIf(MyCommon.Extract_Val(Request.QueryString("select")) <> "", MyCommon.Extract_Val(Request.QueryString("select")), 0)
        UniqueProduct = IIf(MyCommon.Extract_Val(Request.QueryString("Unique")) <> "", MyCommon.Extract_Val(Request.QueryString("Unique")), 0)
        If TierLevels > 1 Or EnableAccum Then
          MinPurchAmt = 0
        Else
          MinPurchAmt = IIf(MyCommon.Extract_Val(Request.QueryString("MinPurchAmt")) <> "", MyCommon.Extract_Val(Request.QueryString("MinPurchAmt")), 0)
        End If
        ' lets handle the selected first
        If (Request.QueryString("selGroups") <> "") OrElse (Request.QueryString("require_pg") <> "") Then
          historyString = Copient.PhraseLib.Lookup("term.alteredproductgroups", LanguageID) & ": " & Request.QueryString("selGroups")
          If (UniqueProduct = 1) Then
            IsUniqueProd = True
          End If
          If IncentiveProdGroupID = 0 AndAlso (Not Request.QueryString("require_pg") <> "") Then
            'If this is a product disqualifier then then Qty=0 and Type=1
            MyCommon.QueryStr = "dbo.pa_CPE_AddProductGroup"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@ROID", SqlDbType.Int, 4).Value = roid
            MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.Int, 4).Value = ProdID
            MyCommon.LRTsp.Parameters.Add("@QtyForIncentive", SqlDbType.Decimal, 12).Value = IIf(Disqualifier, 1, Qty)
            MyCommon.LRTsp.Parameters.Add("@QtyUnitType", SqlDbType.Int, 4).Value = IIf(Disqualifier, 1, Type)
            MyCommon.LRTsp.Parameters.Add("@AccumMin", SqlDbType.Decimal, 12).Value = AccumMin
            MyCommon.LRTsp.Parameters.Add("@AccumLimit", SqlDbType.Decimal, 12).Value = AccumLimit
            MyCommon.LRTsp.Parameters.Add("@AccumPeriod", SqlDbType.Decimal, 12).Value = AccumPeriod
            MyCommon.LRTsp.Parameters.Add("@UniqueProduct", SqlDbType.Bit).Value = UniqueProduct
            MyCommon.LRTsp.Parameters.Add("@RequiredFromTemplate", SqlDbType.Bit).Value = IIf(HasRequiredPG, 1, 0)
            MyCommon.LRTsp.Parameters.Add("@Disqualifier", SqlDbType.Bit).Value = IIf(Disqualifier, 1, 0)
            MyCommon.LRTsp.Parameters.Add("@Rounding", SqlDbType.Bit).Value = IIf(Rounding, 1, 0)
            MyCommon.LRTsp.Parameters.Add("@MinPurchAmt", SqlDbType.Decimal, 12).Value = MinPurchAmt
            MyCommon.LRTsp.Parameters.Add("@ReturnedItemGroup", SqlDbType.bit).Value = 0
            MyCommon.LRTsp.Parameters.Add("@IncentiveProductGroupID", SqlDbType.Int, 4).Direction = ParameterDirection.Output
            MyCommon.LRTsp.ExecuteNonQuery()
            IncentiveProdGroupID = MyCommon.LRTsp.Parameters("@IncentiveProductGroupID").Value
            MyCommon.Close_LRTsp()
          Else
            MyCommon.QueryStr = "update CPE_IncentiveProductGroups set ProductGroupID=" & IIf(ProdID = -1, "NULL", ProdID) & ", QtyForIncentive=" & IIf(Disqualifier, 1, Qty) & ", QtyUnitType=" & IIf(Disqualifier, 1, Type) & ", " & _
                                "AccumMin=" & AccumMin & ", AccumLimit=" & AccumLimit & ", AccumPeriod=" & AccumPeriod & ", ExcludedProducts=0, " & _
                                "RequiredFromTemplate=" & IIf(HasRequiredPG, "1", "0") & ", TCRMAStatusFlag=3, Disqualifier=" & IIf(Disqualifier, "1", "0") & ", " & _
                                "UniqueProduct=" & UniqueProduct & ", Rounding=" & IIf(Rounding, "1", "0") & ", MinPurchAmt=" & MinPurchAmt & " " & _
                                "where IncentiveProductGroupID=" & IncentiveProdGroupID & " and RewardOptionID=" & roid & ";"
            MyCommon.LRT_Execute()
          End If
          
          ' For a CAM offer, if the QtyUnitType=2 (dollars), then force the offer's EOS processing to turn on.
          If (EngineID = 6) Then
            If (Type = 2) Then
              MyCommon.QueryStr = "update CPE_Incentives set DeferCalcToEOS=1 where IncentiveID=" & OfferID & ";"
              MyCommon.LRT_Execute()
            End If
          End If
          
          'Saving Tiers
          If IncentiveProdGroupID <> 0 Then
            MyCommon.QueryStr = "delete from CPE_IncentiveProductGroupTiers where RewardOptionID=" & roid & " and IncentiveProductGroupID=" & IncentiveProdGroupID
            MyCommon.LRT_Execute()
          Else
            'MyCommon.QueryStr = "select IncentiveProductGroupID from CPE_IncentiveProductGroups where RewardOptionID=" & roid & " and ProductGroupID=" & ProdID & " and deleted=0"
            'TierDT = MyCommon.LRT_Select()
            'IncentiveProdGroupID = TierDT.Rows(0).Item("IncentiveProductGroupID")
          End If
          If Disqualifier Then
            'Qty is now set to 0 for a disqualifier
            TierQty = MyCommon.Extract_Val(Request.QueryString("t1_limit"))
            MyCommon.QueryStr = "dbo.pa_CPE_AddProductGroupTiers"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@ROID", SqlDbType.Int, 4).Value = roid
            MyCommon.LRTsp.Parameters.Add("@IncentiveProductGroupID", SqlDbType.Int, 4).Value = IncentiveProdGroupID
            MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int, 4).Value = 1
            MyCommon.LRTsp.Parameters.Add("@Qty", SqlDbType.Decimal, 12).Value = 1 'TierQty
            MyCommon.LRTsp.ExecuteNonQuery()
            MyCommon.Close_LRTsp()
          Else
            For t = 1 To TierLevels
              TierQty = MyCommon.Extract_Val(Request.QueryString("t" & t & "_limit"))
              MyCommon.QueryStr = "dbo.pa_CPE_AddProductGroupTiers"
              MyCommon.Open_LRTsp()
              MyCommon.LRTsp.Parameters.Add("@ROID", SqlDbType.Int, 4).Value = roid
              MyCommon.LRTsp.Parameters.Add("@IncentiveProductGroupID", SqlDbType.Int, 4).Value = IncentiveProdGroupID
              MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int, 4).Value = t
              MyCommon.LRTsp.Parameters.Add("@Qty", SqlDbType.Decimal, 12).Value = TierQty
              MyCommon.LRTsp.ExecuteNonQuery()
              MyCommon.Close_LRTsp()
            Next
          End If
          'Set offer to limit frequency to once per transaction
          If Request.QueryString("Unique") <> "" Then
            MyCommon.QueryStr = "update CPE_Incentives set P3DistQtyLimit=1, P3DistTimeType=2, P3DistPeriod=1 where IncentiveID=" & OfferID
            MyCommon.LRT_Execute()
          End If
        ElseIf HasRequiredPG Then
          If IncentiveProdGroupID = 0 Then
            MyCommon.QueryStr = "insert into CPE_IncentiveProductGroups (RewardOptionID,ProductGroupID,ExcludedProducts,Deleted,LastUpdate,QtyForIncentive,QtyUnitType,AccumMin,AccumLimit,AccumPeriod,RequiredFromTemplate,TCRMAStatusFlag,Disqualifier,UniqueProduct,Rounding,MinPurchAmt)" & _
                                " values (" & roid & "," & ProdID & ",0,0,getdate()," & IIf(Disqualifier, 1, Qty) & "," & IIf(Disqualifier, 1, Type) & "," & AccumMin & "," & AccumLimit & "," & AccumPeriod & ",1,3," & IIf(Disqualifier, "1", "0") & "," & UniqueProduct & "," & IIf(Rounding, "1", "0") & "," & MinPurchAmt & ")"
            MyCommon.LRT_Execute()
          Else
            MyCommon.QueryStr = "update CPE_IncentiveProductGroups set ProductGroupID=" & ProdID & ",QtyForIncentive=" & IIf(Disqualifier, 1, Qty) & "," & _
                                "QtyUnitType=" & IIf(Disqualifier, 1, Type) & ",AccumMin=" & AccumMin & ",AccumLimit=" & AccumLimit & ",AccumPeriod=" & AccumPeriod & "," & _
                                "RequiredFromTemplate=" & IIf(HasRequiredPG, "1", "0") & ",TCRMAStatusFlag=3,Disqualifier=" & IIf(Disqualifier, "1", "0") & "," & _
                                "UniqueProduct=" & UniqueProduct & ", Rounding=" & IIf(Rounding, "1", "0") & ", MinPurchAmt=" & MinPurchAmt & " " & _
                                "where IncentiveProductGroupID=" & IncentiveProdGroupID & " and RewardOptionID=" & roid & ";"
            MyCommon.LRT_Execute()
          End If
        End If
        
        ' check to see if a product condition is required by the template, if applicable
        MyCommon.QueryStr = "select ProductGroupID from CPE_IncentiveProductGroups where RewardOptionID=" & roid & _
                            " and RequiredFromTemplate=1 and Deleted=0 and ExcludedProducts=1;"
        rst = MyCommon.LRT_Select
        HasRequiredPG = (rst.Rows.Count > 0)
        
        ' we got some excluded groups so lets blow out all the existing ones
        If Not Disqualifier Then
          MyCommon.QueryStr = "update CPE_IncentiveProductGroups set Deleted=1,TCRMAStatusFlag=3 where RewardOptionID=" & roid & _
                              " and Deleted=0 and ExcludedProducts=1;"
          MyCommon.LRT_Execute()
        End If
        
        'Change
        ' now lets handle the excluded
        If (Request.QueryString("exGroups") <> "") Then
          historyString = historyString & " " & Copient.PhraseLib.Lookup("term.excluding", LanguageID) & ": " & Request.QueryString("exGroups")
          'Ids = Request.QueryString("exGroups").Split(",")
          ExProdID = MyCommon.Extract_Val(Request.QueryString("exGroups"))
          If ExProdID > 0 Then
            MyCommon.QueryStr = "select IncentiveProductGroupID from CPE_IncentiveProductGroups with (NoLock) " & _
                                "where RewardOptionID=" & roid & " and Deleted=0 and ExcludedProducts=1;"
            dt = MyCommon.LRT_Select
            If dt.Rows.Count > 0 Then
              MyCommon.QueryStr = "update CPE_IncentiveProductGroups set ProductGroupID=" & ExProdID & ",ExcludedProducts=1, QtyForIncentive=" & Qty & ", QtyUnitType=" & Type & ", AccumMin=0, AccumLimit=0," & _
                                  " AccumPeriod=0, RequiredFromTemplate=" & IIf(HasRequiredPG, "1", "0") & ", TCRMAStatusFlag=3, Disqualifier=" & IIf(Disqualifier, "1", "0") & ", UniqueProduct=" & UniqueProduct & ", Rounding=" & IIf(Rounding, "1", "0") & ", MinPurchAmt=" & MinPurchAmt & " " & _
                                  "where IncentiveProductGroupID=" & MyCommon.NZ(dt.Rows(0).Item("IncentiveProductGroupID"), 0) & " and RewardOptionID=" & roid
              MyCommon.LRT_Execute()
            Else
              RestrictExcluding = IsExcludingRestricted(MyCommon, roid)
              If Not RestrictExcluding Then
                MyCommon.QueryStr = "insert into CPE_IncentiveProductGroups (RewardOptionID, ProductGroupID, ExcludedProducts, QtyForIncentive, QtyUnitType, AccumMin, AccumLimit, AccumPeriod, Deleted, LastUpdate, RequiredFromTemplate, TCRMAStatusFlag, Disqualifier, UniqueProduct, Rounding, MinPurchAmt) " & _
                                    " values (" & roid & "," & ExProdID & ", 1, " & Qty & "," & Type & ",0,0,0," & _
                                    " 0, getdate()," & IIf(HasRequiredPG, "1", "0") & ",3," & IIf(Disqualifier, "1", "0") & "," & IIf(UniqueProduct, "1", "0") & "," & IIf(Rounding, "1", "0") & "," & MinPurchAmt & ")"
                MyCommon.LRT_Execute()
              End If
            End If
          End If
        ElseIf HasRequiredPG Then
          If IncentiveProdGroupID = 0 Then
            MyCommon.QueryStr = "insert into CPE_IncentiveProductGroups (RewardOptionID,ProductGroupID,ExcludedProducts,Deleted,LastUpdate,RequiredFromTemplate,TCRMAStatusFlag,Disqualifier,UniqueProduct,Rounding,MinPurchAmt) " & _
                                " values(" & roid & "," & ExProdID & ",1,0,getdate(),1,3," & IIf(Disqualifier, "1", "0") & "," & UniqueProduct & "," & IIf(Rounding, "1", "0") & "," & MinPurchAmt & ")"
            MyCommon.LRT_Execute()
          Else
            MyCommon.QueryStr = "update CPE_IncentiveProductGroups set ProductGroupID=" & ExProdID & ",ExcludedProducts=1,QtyForIncentive=" & Qty & ",QtyUnitType=" & Type & ",AccumMin=" & AccumMin & ",AccumLimit=" & AccumLimit & "," & _
                                " AccumPeriod=" & AccumPeriod & ",RequiredFromTemplate=" & IIf(HasRequiredPG, "1", "0") & ",TCRMAStatusFlag=3,Disqualifier=" & IIf(Disqualifier, "1", "0") & ",UniqueProduct=" & UniqueProduct & ",Rounding=" & IIf(Rounding, "1", "0") & ", MinPurchAmt=" & MinPurchAmt & " " & _
                                "where IncentiveProductGroupID=" & IncentiveProdGroupID & " and RewardOptionID=" & roid
            MyCommon.LRT_Execute()
          End If
        End If
        
        
        
        MyCommon.QueryStr = "update CPE_Incentives set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID
        MyCommon.LRT_Execute()
        
        'check if accumulation message needs to be removed
        MyCommon.QueryStr = "dbo.pa_CPE_AccumMsgEligible"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@ROID", SqlDbType.Int, 4).Value = roid
        MyCommon.LRTsp.Parameters.Add("@AccumEligible", SqlDbType.Bit, 1).Direction = ParameterDirection.Output
        MyCommon.LRTsp.ExecuteNonQuery()
        AccumEligible = MyCommon.LRTsp.Parameters("@AccumEligible").Value
        MyCommon.Close_LRTsp()
        
        If Not (AccumEligible) Then
          'Mark any accumulation messages as deleted
          MyCommon.QueryStr = "update CPE_Deliverables set Deleted = 1 where DeliverableID in " & _
                              "(select D.DeliverableID from CPE_RewardOptions RO inner join CPE_Deliverables D on RO.RewardOptionID = D.RewardOptionID " & _
                              "where RO.Deleted = 0 and D.Deleted = 0 and RO.IncentiveID = " & OfferID & " and RewardOptionPhase = 2 and DeliverableTypeID = 4);"
          'MyCommon.QueryStr = "update CPE_Deliverables set Deleted=1 where RewardOptionId=" & roid & " and RewardOptionPhase=2 and deleted=0;"
          MyCommon.LRT_Execute()
        End If
        
        'Finally: accumulation and disqualification are mutually incompatible within an offer, so:
        'if this condition is a disqualifier, disable accumulation on every product condition in this offer, or
        'if this condition has accumulation, delete any associated disqualifiers from this offer.
        If Disqualifier Then
          MyCommon.QueryStr = "update CPE_IncentiveProductGroups set AccumMin=0, AccumLimit=0, AccumPeriod=0 " & _
                              "where RewardOptionID=" & roid & ";"
          MyCommon.LRT_Execute()
        ElseIf (AccumMin > 0 OrElse AccumLimit > 0 OrElse AccumPeriod > 0) Then
          MyCommon.QueryStr = "delete from CPE_IncentiveProductGroupTiers where IncentiveProductGroupID in " & _
                              "(select IncentiveProductGroupID from CPE_IncentiveProductGroups where RewardOptionID=" & roid & " and Deleted=0 and Disqualifier=1) " & _
                              "and RewardOptionID=" & roid & ";"
          MyCommon.LRT_Execute()
          MyCommon.QueryStr = "update CPE_IncentiveProductGroups with (RowLock) set Deleted=1, LastUpdate=getdate(), TCRMAStatusFlag=3 " & _
                              "where RewardOptionID=" & roid & " and Disqualifier=1;"
          MyCommon.LRT_Execute()
        End If
        
        MyCommon.Activity_Log(3, OfferID, AdminUserID, historyString)
      End If
    Else
      If Not ValidTier Then
        infoMessage = Copient.PhraseLib.Lookup("error.tiervalues", LanguageID)
        'IncentiveProdGroupID = 0
      ElseIf Not ValidAccum Then
        infoMessage = Copient.PhraseLib.Lookup("error.validaccum", LanguageID)
        'IncentiveProdGroupID = 0
      ElseIf Not ValidRounding Then
        infoMessage = Copient.PhraseLib.Lookup("error.rounding", LanguageID)
        'IncentiveProdGroupID = 0
      Else
        infoMessage = Copient.PhraseLib.Lookup("term.errornorewardoption", LanguageID)
        'IncentiveProdGroupID = 0
      End If
    End If
    
    If (infoMessage = "") Then
      CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
    Else
      CloseAfterSave = False
    End If
  End If
  
  ' dig the offer info out of the database
  ' no one clicked anything
  MyCommon.QueryStr = "Select IncentiveID,IsTemplate,ClientOfferID,IncentiveName as Name,CPE.Description,PromoClassID,CRMEngineID,Priority," & _
                      "StartDate,EndDate,EveryDOW,EligibilityStartDate,EligibilityEndDate,TestingStartDate,TestingEndDate,P1DistQtyLimit,P1DistTimeType,P1DistPeriod," & _
                      "P2DistQtyLimit,P2DistTimeType,P2DistPeriod,P3DistQtyLimit,P3DistTimeType,P3DistPeriod,EnableRedeemRpt,EnableImpressRpt,CreatedDate,CPE.LastUpdate,CPE.Deleted,CPEOARptDate,CPEOADeploySuccessDate,CPEOADeployRpt," & _
                      "CRMRestricted,StatusFlag,OC.Description as CategoryName,IsTemplate,FromTemplate from CPE_Incentives as CPE " & _
                      "left join OfferCategories as OC on CPE.PromoClassID=OfferCategoryID " & _
                      "where IncentiveID=" & Request.QueryString("OfferID") & ";"
  rst = MyCommon.LRT_Select
  For Each row In rst.Rows
    Name = MyCommon.NZ(row.Item("Name"), "")
    isTemplate = MyCommon.NZ(row.Item("IsTemplate"), False)
    FromTemplate = MyCommon.NZ(row.Item("FromTemplate"), False)
  Next
  
  'update the templates permission if necessary
  If (Request.QueryString("save") <> "" AndAlso Request.QueryString("IsTemplate") = "IsTemplate") Then
    
    ' time to update the status bits for the templates
    Dim form_Disallow_Edit As Integer = 0
    Dim form_Require_PG As Integer = 0
    
    If (Request.QueryString("Disallow_Edit") = "on") Then
      form_Disallow_Edit = 1
    End If
    
    If (Request.QueryString("require_pg") <> "") Then
      form_Require_PG = 1
    End If
    
    If (form_Disallow_Edit = 1 AndAlso form_Require_PG = 1) Then
      infoMessage = Copient.PhraseLib.Lookup("offer-con.lockeddenied", LanguageID)
      MyCommon.QueryStr = "update CPE_IncentiveProductGroups set DisallowEdit=1, RequiredFromTemplate=0 " & _
                          " where RewardOptionID=" & roid & " and Deleted=0 and Disqualifier=" & IIf(Disqualifier, "1", "0") & ";"
    Else
      MyCommon.QueryStr = "update CPE_IncentiveProductGroups set DisallowEdit=" & form_Disallow_Edit & ", RequiredFromTemplate=" & form_Require_PG & _
                          " where RewardOptionID=" & roid & " and Deleted=0 and Disqualifier=" & IIf(Disqualifier, "1", "0") & ";"
    End If
    MyCommon.LRT_Execute()
    
    ' If necessary, create an empty product condition
    If (form_Require_PG = 1) Then
      MyCommon.QueryStr = "select ProductGroupID from CPE_IncentiveProductGroups with (NoLock) " & _
                          " where RewardOptionID=" & roid & " and Deleted=0;"
      rst = MyCommon.LRT_Select
      If (rst.Rows.Count = 0) Then
        ' Create the product condition record, then the product condition tier record(s)
        MyCommon.QueryStr = "insert into CPE_IncentiveProductGroups (RewardOptionID,QtyForIncentive,QtyUnitType,ExcludedProducts,Deleted,LastUpdate,RequiredFromTemplate,TCRMAStatusFlag)" & _
                            " values(" & roid & "," & MyCommon.Extract_Val(Request.QueryString("t1_limit")) & "," & MyCommon.Extract_Val(Request.QueryString("select")) & ",0,0,getdate(),1,3);"
        MyCommon.LRT_Execute()
        MyCommon.QueryStr = "select top 1 IncentiveProductGroupID from CPE_IncentiveProductGroups where RewardOptionID=" & roid & " order by LastUpdate DESC;"
        rst = MyCommon.LRT_Select()
        If rst.Rows.Count > 0 Then
          For t = 1 To TierLevels
            TierQty = MyCommon.Extract_Val(Request.QueryString("t" & t & "_limit"))
            MyCommon.QueryStr = "dbo.pa_CPE_AddProductGroupTiers"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@ROID", SqlDbType.Int, 4).Value = roid
            MyCommon.LRTsp.Parameters.Add("@IncentiveProductGroupID", SqlDbType.Int, 4).Value = rst.Rows(0).Item("IncentiveProductGroupID")
            MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int, 4).Value = t
            MyCommon.LRTsp.Parameters.Add("@Qty", SqlDbType.Decimal, 12).Value = TierQty
            MyCommon.LRTsp.ExecuteNonQuery()
            MyCommon.Close_LRTsp()
          Next
        End If
      End If
    End If
    
    If (infoMessage = "") Then
      CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
    Else
      CloseAfterSave = False
    End If
    
  End If
  
  If (isTemplate Or FromTemplate) Then
    ' lets dig the permissions if its a template
    MyCommon.QueryStr = "select DisallowEdit, RequiredFromTemplate from CPE_IncentiveProductGroups with (NoLock) " & _
                        " where RewardOptionID=" & roid & " and Deleted=0 and Disqualifier=" & IIf(Disqualifier, "1", "0") & ";"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      Disallow_Edit = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit"), False)
      RequirePG = MyCommon.NZ(rst.Rows(0).Item("RequiredFromTemplate"), False)
    Else
      Disallow_Edit = False
    End If
  End If
  
  If Not isTemplate Then
    DisabledAttribute = IIf(Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Edit), "", " disabled=""disabled""")
  Else
    DisabledAttribute = IIf(Logix.UserRoles.EditTemplates, "", " disabled=""disabled""")
  End If
  
  If HasBundleDiscount Then
    DisabledAttribute = " disabled=""disabled"""
  End If
  
  If Not Disqualifier Then
    Send_HeadBegin("term.offer", "term.productcondition", OfferID)
  Else
    Send_HeadBegin("term.offer", "term.productdisqualifier", OfferID)
  End If
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
%>
<script type="text/javascript">
  var fullSelect = null;
  var IsExcludeRestricted = false;
  
// This is the javascript array holding the function list
// The PrintJavascriptArray ASP function can be used to print this array.


// This is the function that refreshes the list after a keypress.
// The maximum number to show can be limited to improve performance with
// huge lists (1000s of entries).
// The function clears the list, and then does a linear search through the
// globally defined array and adds the matches back to the list.
function handleKeyUp(maxNumToShow) {
  var selectObj, textObj, functionListLength;
  var i,  numShown;
  var searchPattern;
  
  document.getElementById("functionselect").size = "12";
  
  // Set references to the form elements
  selectObj = document.forms[0].functionselect;
  textObj = document.forms[0].functioninput;
  
  // Remember the function list length for loop speedup
  functionListLength = functionlist.length;
  
  // Set the search pattern depending
  if(document.forms[0].functionradio[0].checked == true) {
    searchPattern = "^"+textObj.value;
  } else {
    searchPattern = textObj.value;
  }
  searchPattern = cleanRegExpString(searchPattern);
  
  // Create a regular expression
  re = new RegExp(searchPattern,"gi");
  
  // Loop through the array and re-add matching options
  numShown = 0;
  if (textObj.value == '' && fullSelect != null) {
    var newSelectBox = fullSelect.cloneNode(true);
    document.getElementById('pgList').replaceChild(newSelectBox, selectObj);
  } else {
    var newSelectBox = selectObj.cloneNode(false);
    document.getElementById('pgList').replaceChild(newSelectBox, selectObj);
    selectObj = document.getElementById("functionselect");
    for(i = 0; i < functionListLength; i++) {
      if(functionlist[i].search(re) != -1) {
        if (vallist[i] != "") {
          selectObj[numShown] = new Option(functionlist[i], vallist[i]);
          if (vallist[i] == 1) {
            selectObj[numShown].style.fontWeight = 'bold';
            selectObj[numShown].style.color = 'brown';
          }
          numShown++;
        }
      }
      // Stop when the number to show is reached
      if(numShown == maxNumToShow) {
        break;
      }
    }
  }
  removeUsed(true);
  // When options list whittled to one, select that entry
  if(selectObj.length == 1) {
    selectObj.options[0].selected = true;
  }
}

function removeUsed(bSkipKeyUp) {
  //if (!bSkipKeyUp) handleKeyUp(99999);
  if (!bSkipKeyUp) { xmlhttpPost('OfferFeeds.aspx','ProductGroups'); }
  // this function will remove items from the functionselect box that are used in 
  // selected and excluded boxes
  var funcSel = document.getElementById('functionselect');
  var exSel = document.getElementById('selected');
  var elSel = document.getElementById('excluded');
  var i,j;
  for (i = elSel.length - 1; i>=0; i--) {
    for(j=funcSel.length-1;j>=0; j--) {
      if(funcSel.options[j].value == elSel.options[i].value){
        funcSel.options[j] = null;
      }
    }
  }
  for (i = exSel.length - 1; i>=0; i--) {
    for(j=funcSel.length-1;j>=0; j--) {
      if(funcSel.options[j].value == exSel.options[i].value){
        funcSel.options[j] = null;
      }
    }
  }
}

// this function gets the selected value and loads the appropriate
// php reference page in the display frame
// it can be modified to perform whatever action is needed, or nothing
function handleSelectClick(itemSelected) {
  var isAnyProductSelected = false;
  var item, parent;
  
  textObj = document.forms[0].functioninput;
  
  selectObj = document.forms[0].functionselect;
  selectedValue = document.getElementById("functionselect").value;
  if(selectedValue != ""){ selectedText = selectObj[document.getElementById("functionselect").selectedIndex].text; }
  
  selectboxObj = document.forms[0].selected;
  selectedboxValue = document.getElementById("selected").value;
  if(selectedboxValue != ""){ selectedboxText = selectboxObj[document.getElementById("selected").selectedIndex].text; }
  
  excludedbox = document.forms[0].excluded;
  excludedboxValue = document.getElementById("excluded").value;
  if(excludedboxValue != ""){ excludeboxText = excludedbox[document.getElementById("excluded").selectedIndex].text; }
  
  if(itemSelected == "select1") {
    if(selectedValue != "") {
      // add items to selected box
//    while (selectObj.selectedIndex != -1) {
//      selectedText = selectObj.options[selectObj.selectedIndex].text;
//      selectedValue = selectObj.options[selectObj.selectedIndex].value;
//      selectboxObj[selectboxObj.length] = new Option(selectedText,selectedValue);
//      selectObj[selectObj.selectedIndex].selected = false;
//      document.getElementById("select1").disabled = true;
//    }
      selectedText = selectObj.options[selectObj.selectedIndex].text;
      selectedValue = selectObj.options[selectObj.selectedIndex].value;
      selectboxObj[selectboxObj.length] = new Option(selectedText,selectedValue);
      selectObj[selectObj.selectedIndex].selected = false;
      if (selectedValue == "1") {
        selectboxObj[selectboxObj.length-1].style.color = 'brown';
        selectboxObj[selectboxObj.length-1].style.fontWeight = 'bold';
      }
      document.getElementById("select1").disabled = true;
    }
  }
  
  if(itemSelected == "deselect1") {
    if(selectedboxValue != "") {
      // remove items from selected box
      while (document.getElementById("selected").selectedIndex != -1) {
        if(selectedboxValue == 1) {
          if (excludedbox.length > 0) { 
            if (confirm('<%Sendb(Copient.PhraseLib.Lookup("CPEoffer-con-product.anyproductconfirm", LanguageID))%>')) {
              document.getElementById("excluded").remove(0);
              document.getElementById("selected").remove(document.getElementById("selected").selectedIndex);
              document.getElementById("select1").disabled = false;
            } else {
              return;
            }
          } else {
            document.getElementById("selected").remove(document.getElementById("selected").selectedIndex);
            document.getElementById("select1").disabled = false;
          }
        } else {
          document.getElementById("selected").remove(document.getElementById("selected").selectedIndex);
          document.getElementById("select1").disabled = false;
        }
      }        
    }
    selectboxObj.options.selectedIndex = -1;
    document.getElementById("select1").disabled = false;
  }
  
  if(itemSelected == "select2") {
    if(selectedValue != ""){
      // add items to excluded box
      
      for (var i=0; i < selectboxObj.length; i++) {
        if (selectboxObj.options[i].value == '1') {
          isAnyProductSelected = true;
        }
      }        
      if(isAnyProductSelected){
        excludedbox[excludedbox.length] = new Option(selectedText,selectedValue);
      }
    }
  }
  
  if(itemSelected == "deselect2") {
    if(excludedboxValue != ""){
      // remove items from excluded box    
      //document.getElementById("excluded").remove(document.getElementById("excluded").selectedIndex)
      document.getElementById("excluded").options[document.getElementById("excluded").selectedIndex] = null;
    }
  }
  
  // remove items from large list that are in the other lists
  removeUsed(false);
  //removeUsed(true);
  updateButtons();
  
  return true;
}

function saveForm(){
  var dqElem = document.getElementById('Disqualifier');
  var funcSel = document.getElementById('functionselect');
  var exSel = document.getElementById('excluded');
  var elSel = document.getElementById('selected');
  var i,j;
  var selectList = "";
  var excludededList = "";
  var htmlContents = "";
  var bValidEntry = false;
  var isDisqualifier = false;
  
  if (dqElem!=null) { isDisqualifier = (dqElem.value=="1") }
  
  if(!ValidSave(isDisqualifier)) {
    return false;
  }
    
  // assemble the list of values from the selected box
  for (i = elSel.length - 1; i>=0; i--) {
    if(elSel.options[i].value != ""){
      if(selectList != "") { selectList = selectList + ","; }
      selectList = selectList + elSel.options[i].value;
      
      bValidEntry = checkEntries(elSel.options[i].value);
      if (!bValidEntry) { return false; }
    }
  }
  for (i = exSel.length - 1; i>=0; i--) {
    if(exSel.options[i].value != ""){
      if(excludededList != "") { excludededList = excludededList + ","; }
      excludededList = excludededList + exSel.options[i].value;
    }
  }
  // ok time to build up the hidden variables to pass for saving
  htmlContents = "<input type=\"hidden\" name=\"selGroups\" value=" + selectList + "> ";
  htmlContents = htmlContents + "<input type=\"hidden\" name=\"exGroups\" value=" + excludededList + ">";
  document.getElementById("hiddenVals").innerHTML = htmlContents;
  return true;
}

function updateButtons() {
  var elemSelect1 = document.getElementById('select1');
  var elemSelect2 = document.getElementById('select2');
  var elemDeselect1 = document.getElementById('deselect1');
  var elemDeselect2 = document.getElementById('deselect2');
  var elemSave = document.getElementById('save');
  var elemSelected = document.getElementById('selected');
  
  var selectboxObj = document.forms[0].selected;
  var selectedboxValue = selectboxObj.value;
  var excludedbox = document.forms[0].excluded;
  var isAnyProductSelected = false;
  
  if (selectboxObj != null) {
    elemDeselect1.disabled = (selectboxObj.length == 0 || elemSelected.disabled) ? true : false;
    elemSelect1.disabled = (selectboxObj.length > 0) ? true : false;
    if (selectboxObj.length == 0) {
      if (document.getElementById('require_pg') != null) {
        if (document.getElementById('require_pg').checked == true) {
          elemSave.disabled = false;
        } else {
          elemSave.disabled = true;
        }
      } else {
        elemSave.disabled = true;
      }
    } else {
      elemSave.disabled = false;
    }
    for (var i=0; i < selectboxObj.length; i++) {
      if (selectboxObj.options[i].value == '1') {
        isAnyProductSelected = true;
      }
    }
  } else {
    elemSelect1.disabled = false;
  }
  
  if (excludedbox != null) {
    elemSelect2.disabled = (excludedbox.length == 0 && isAnyProductSelected) ? false : true;
    elemDeselect2.disabled = (excludedbox.length > 0) ? false : true;     
  }
  
  IsExcludeRestricted = '<%Sendb(RestrictExcluding)%>';
  if (IsExcludeRestricted == "True") {
    document.getElementById("select2").disabled = true;
  }

  
  <%
   If Not isTemplate Then   
     If Not (Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Edit)) Then
      Send("  disableAll();")
     End If 
   Else
     If Not (Logix.UserRoles.EditTemplates) Then
      Send("  disableAll();")
     End If
   End If        
  %>
}
var isFireFox = (navigator.appName.indexOf('Mozilla') !=-1) ? true : false;
var timer;
function xmlPostTimer(strURL,mode)
{
  clearTimeout(timer);
  timer=setTimeout("xmlhttpPost('" + strURL + "','" + mode + "')", 250);
}


function xmlhttpPost(strURL,mode) {
  var xmlHttpReq = false;
  var self = this;
  
  //document.getElementById("functionselect").style.display = "none";
  document.getElementById("searchLoadDiv").innerHTML = '<% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID)) %>';
  
  //handleSaveButton(true);
  // Mozilla/Safari
  if (window.XMLHttpRequest) {
    self.xmlHttpReq = new XMLHttpRequest();
  }
  // IE
  else if (window.ActiveXObject) {
    self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
  }
  var qryStr = getproductquery(mode);
  self.xmlHttpReq.open('POST', strURL + "?" + qryStr, true);
  self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
  self.xmlHttpReq.onreadystatechange = function() {
    if (self.xmlHttpReq !=null && self.xmlHttpReq.readyState == 4) {
      updatepage(self.xmlHttpReq.responseText);
    }
  }

  self.xmlHttpReq.send(qryStr);
  //self.xmlHttpReq.send(getquerystring());
}

function getproductquery(mode) {
  var radioString;
  if(document.getElementById('functionradio2').checked) {
    radioString = 'functionradio2';
  }
  else {
    radioString = 'functionradio1';
  }
  var selected = document.getElementById('selected');
  var selectedGroup = 0;
  if(selected.options[0] != null){
    selectedGroup = selected.options[0].value;
  }
  var excluded = document.getElementById('excluded');
  var excludedGroup = 0;
  if(excluded.options[0] != null){
    excludedGroup = excluded.options[0].value;
  }
  return "Mode=" + mode + "&ProductSearch=" + document.getElementById('functioninput').value + "&ROID=" + document.getElementById('roid').value + "&Disqualifier=" + document.getElementById('Disqualifier').value + "&SelectedGroup=" + selectedGroup + "&ExcludedGroup=" + excludedGroup + "&SearchRadio=" + radioString;
 
}

function updatepage(str) {
  if(str.length > 0){
    if(!isFireFox){
      document.getElementById("pgList").innerHTML = '<select class="longer" id="functionselect" name="functionselect" size="12"<% sendb(disabledattribute) %>>' + str + '</select>';
    }
    else{
      document.getElementById("functionselect").innerHTML = str;
    }
    document.getElementById("searchLoadDiv").innerHTML = '&nbsp;';
    if (document.getElementById("functionselect").options.length > 0) {
      document.getElementById("functionselect").options[0].selected = true;
    }
    //handleSaveButton(false);
  }
  else if(str.length == 0){
    if(!isFireFox){
      document.getElementById("pgList").innerHTML = '<select class="longer" id="functionselect" name="functionselect" size="12"<% sendb(disabledattribute) %>>&nbsp;</select>';
    }
    else{
      document.getElementById("functionselect").innerHTML = '&nbsp;';
    }
    document.getElementById("searchLoadDiv").innerHTML = '&nbsp;';
    
   // handleSaveButton(false);
  }
}

function handleSaveButton(bDisabled) {
  var elemSave = document.getElementById("save");
  
  if (elemSave != null) {
    elemSave.disabled = bDisabled
  }
}

function checkEntries(pgID) {
  var bValidEntry = false;
  var invalidElemID = '';
  var elemInvalid = null;
  var saveElem = document.getElementById("save");
  
  // check for invalid entries in the textboxes
  bValidEntry = isValidEntry('limit', 'select');
  if (!bValidEntry) { invalidElemID = 'limit'; }
  bValidEntry = bValidEntry && isValidEntry('accummin', 'select');
  if (!bValidEntry && invalidElemID == '') { invalidElemID = 'accummin'; }
  bValidEntry = bValidEntry && isValidEntry('accumlimit', 'select');
  if (!bValidEntry && invalidElemID == '') { invalidElemID = 'accumlimit'; }
  bValidEntry = bValidEntry && isValidEntry('accumperiod', 'select');
  if (!bValidEntry && invalidElemID == '') { invalidElemID = 'accumperiod'; }
  
  if (!bValidEntry) {
    alert('<%Sendb(Copient.PhraseLib.Lookup("term.invalidnumericentry", LanguageID))%>');
    elemInvalid = document.getElementById(invalidElemID);
    if (elemInvalid != null) {
      elemInvalid.focus();
      elemInvalid.select();
    }
    if (saveElem != null) {
      if (saveElem.style.visibility=='hidden') {
        saveElem.style.visibility='visible';
      }
    }
  }
  
  return bValidEntry;
}

function isValidEntry(elemID, elemType) {
  var bValid = true;
  var elem = document.getElementById(elemID);
  var elemUnitType = document.getElementById(elemType);
  var unitType = 0;
  
  if (elem != null) {
    bValid = !isNaN(elem.value);
    if (bValid && elemUnitType != null) {
      unitType = elemUnitType.value;
      
      if (unitType == "1") {
        if (isInteger(elem.value) || elem.value != Math.round(elem.value)) {
          bValid = isInteger(elem.value);
        }
      } else if (unitType == "2") {
        if ((decimalPlaces(elem.value, '.') > 2 && elem.value != (Math.round(elem.value*100)/100)) || parseFloat(elem.Value) < 0) {
          bValid = false;
        }
      } else if (unitType == "3") {
        if ((decimalPlaces(elem.value, '.') > 3 && elem.value != (Math.round(elem.value*1000)/1000)) || parseFloat(elem.Value) < 0) {
          bValid = false;
        }
      }
    }
  }
  
  return bValid;
}

function ValidateEntry(evt, src) {
  var elem = document.getElementById("valuetype");
  
  if (elem != null) {
    if (elem.value == "1") {
        return NumberCheck(evt, src, false);
    } else {
        return NumberCheck(evt, src, true);
    }
  }
}

function NumberCheck(evt, src, allowDecimal) {
  var nkeycode=(window.event) ? window.event.keyCode : evt.which;
  var exceptionKeycodes = new Array(8, 9, 46, 13, 16, 36);
  
  if (nkeycode >= 48 && nkeycode <= 57) {
    return true;
  } else {
    for (var i=0; i < exceptionKeycodes.length; i++) {
      if (nkeycode == exceptionKeycodes[i]) {
        return true;
      }
    }
    if (allowDecimal == true && nkeycode == 190) {
      if (src !=null && src.value.indexOf(".") < 0) {
        return true;
      }
    }
    return false;
  }
}

function ReformatForQty(isTiered) {
  var elem = document.getElementById("valuetype");
  var txtElem = null;
  var maxCt = 999, decPt = -1;
  var txtVal = "";
  
  if (elem != null) {
    if (elem.value == "1") {
      var i = (isTiered) ? 1 : 0;
      while (i < maxCt) {
        txtElem = document.getElementById("tier" + i);
        if (txtElem != null) {
          txtVal = txtElem.value;
          decPt = txtVal.indexOf(".");
          if (decPt >= 0) {
            txtElem.value = txtVal.substring(0, decPt);
          }
        } else {
          break;
        }
        i++;
      }
    }
  }
}

function ValidSave(isDisqualifier){
  var retVal = true;
  var elem = document.getElementById("selected");
  var unitElem = document.getElementById("select");   
  var qtyElem = document.getElementById("t1_limit");
  var minElem = document.getElementById("MinPurchAmt");
  var elemProgram = document.getElementById("IncentiveProdGroupID");
  var saveElem = document.getElementById("save");
  var msg = '';
  var t = 1;
  var unitType = 1;
  
  if (elem != null && elem.options.length == 0) {
    if (document.getElementById('require_pg').checked == false) {
      retVal = false;
      msg += '* <% Sendb(Copient.PhraseLib.Lookup("CPEoffer-con-product.selectgroup", LanguageID)) %>'
      elem.focus();
    }
  } else if (elem !=null && elemProgram != null) {
    elemProgram.value = elem.options[0].value;
  }
  
  if (!isDisqualifier) {
    if (unitElem != null) {
      unitType = parseInt(unitElem.value);
    }
    
    while (qtyElem != null) {
      // trim the string
      var qtyVal = qtyElem.value.replace(/^\s+|\s+$/g, ''); 
      if (unitType==1 && (!isInt(qtyVal) || parseInt(qtyVal)<= 0)) {
        retVal = false;
        if (msg != '') { msg += '\n\r\n\r'; }
        msg += '* <% Sendb(Copient.PhraseLib.Lookup("CPEoffer-con-point.positiveinteger", LanguageID)) %>: <% Sendb(Copient.PhraseLib.Lookup("term.quantity", LanguageID)) %>';
        qtyElem.focus();
        qtyElem.select();
        break;
      } else if ((unitType!=1 && (qtyVal=="" || isNaN(qtyVal))) || (!isNaN(qtyVal) && unitType!=1 && parseFloat(qtyVal)<= 0)) {
        retVal = false;
        if (msg != '') { msg += '\n\r\n\r'; }
        if(unitType==3)
        {
          msg += '* <% Sendb(Copient.PhraseLib.Lookup("CPEoffer-con-point.positivedecimal", LanguageID)) %>: <% Sendb(Copient.PhraseLib.Lookup("term.quantity", LanguageID)) %>';
        }
        else
        {
          msg += '* <% Sendb(Copient.PhraseLib.Lookup("CPEoffer-con-point.positivedecimal", LanguageID)) %>: <% Sendb(Copient.PhraseLib.Lookup("term.price", LanguageID)) %>';
        }
        qtyElem.focus();
        qtyElem.select();
        break;
      }
      t++;
      qtyElem = document.getElementById("t" + t + "_limit");
    }
    
    //Check minimum purchase amount
      var minVal = minElem.value.replace(/^\s+|\s+$/g, '');
      if (unitType==1 && (!isInt(minVal) || parseInt(minVal)< 0)) {
        retVal = false;
        if (msg != '') { msg += '\n\r\n\r'; }
      msg += '* <% Sendb(Copient.PhraseLib.Lookup("CPEoffer-con-point.positiveinteger", LanguageID)) %>: <% Sendb(Copient.PhraseLib.Lookup("term.minimumpurchase", LanguageID)) %>';
        minElem.focus();
        minElem.select();
      } else if ((unitType!=1 && (minVal=="" || isNaN(minVal))) || (!isNaN(minVal) && unitType!=1 && parseFloat(minVal)< 0)) {
        retVal = false;
        if (msg != '') { msg += '\n\r\n\r'; }
      msg += '* <% Sendb(Copient.PhraseLib.Lookup("CPEoffer-con-point.positivedecimal", LanguageID)) %>: <% Sendb(Copient.PhraseLib.Lookup("term.minimumpurchase", LanguageID)) %>';
        minElem.focus();
        minElem.select();
      }
  }
  
  if (msg != '') {
    alert(msg);
    if (saveElem != null) {
      if (saveElem.style.visibility=='hidden') {
        saveElem.style.visibility='visible';
      }
    }
  }
  return retVal;
}

function isInt(sNum) {
  return (sNum!="" && !isNaN(sNum) && (sNum/1)==parseInt(sNum));
}

function ChangeUnit(Type) {
  var prod = document.getElementById('Unique');
  var accum = document.getElementById('EnableAccum');
  var accumlabel = document.getElementById('AccumChecklbl');
  var rounding = document.getElementById('rounding');
  var roundingspan = document.getElementById('roundingspan');
  var minpurch = document.getElementById('MinPurch');
  var minpurchamt = document.getElementById('MinPurchAmt');
  var accumulation = document.getElementById('accumulation');
  var unittypedesc = document.getElementById('UnitTypeDesc');
  
  if (Type == "1") {
    if (accum == null || accum.checked == false) {
      prod.disabled = false; 
    }
    if (roundingspan != null) { roundingspan.style.display = 'none'; }
    if (rounding != null) { rounding.checked = false; }
    if (accum != null) { accum.style.display = 'inline'; }
    if (accumlabel != null) { accumlabel.style.display = 'inline'; }
    if (minpurch != null) { minpurch.style.display = ''; }
    if (unittypedesc != null) { unittypedesc.innerHTML = '<% Sendb(Copient.PhraseLib.Lookup("term.quantity", LanguageID)) %>'; }
  } else if (Type == "4") {
    if (prod != null) { prod.checked = false; }
    if (prod != null) { prod.disabled = true; }
    if (roundingspan != null) { roundingspan.style.display = 'none'; }
    if (rounding != null) { rounding.checked = false; }
    if (accum != null) { accum.style.display = 'none'; }
    if (accumlabel != null) { accumlabel.style.display = 'none'; }
    if (minpurch != null) { minpurch.style.display = 'none'; }
    if (minpurchamt != null) { minpurchamt.value = '0'; }
    if (accum != null) { accum.checked = false; }
    if (accumulation != null) { accumulation.style.display = 'none'; }
    if (unittypedesc != null) { unittypedesc.innerHTML = '<% Sendb(Copient.PhraseLib.Lookup("term.price", LanguageID)) %>'; }
  } else {
    if (prod != null) { prod.checked = false; }
    if (prod != null) { prod.disabled = true; }
    if (roundingspan != null) { roundingspan.style.display = 'inline'; }
    if (rounding != null) { rounding.style.display = 'inline'; }
    if (accum != null) { accum.style.display = 'inline'; }
    if (accumlabel != null) { accumlabel.style.display = 'inline'; }
    if (minpurch != null) { minpurch.style.display = ''; }
    if (Type == "2") {
      if (unittypedesc != null) { unittypedesc.innerHTML = '<% Sendb(Copient.PhraseLib.Lookup("term.price", LanguageID)) %>'; }
    } else {
      if (unittypedesc != null) { unittypedesc.innerHTML = '<% Sendb(Copient.PhraseLib.Lookup("term.quantity", LanguageID)) %>'; }
    }
  }
}

function handleRequiredToggle() {
  if(document.forms[0].selected.length == 0) {
    if (document.getElementById("require_pg").checked == false) {
      document.getElementById('save').disabled=true;
    } else {
      document.getElementById('save').disabled=false;
    }
  }
  if (document.getElementById("require_pg").checked == true) {
    document.getElementById("Disallow_Edit").checked=false;
  }
}

function disableAll() {
  document.getElementById('select1').disabled=true;
  document.getElementById('select2').disabled=true;
  document.getElementById('deselect1').disabled=true;
  document.getElementById('deselect2').disabled=true;
  document.getElementById('functionselect').disabled=true;
  document.getElementById('selected').disabled=true;
  document.getElementById('excluded').disabled=true;
}
</script>
<%
  Send("<script type=""text/javascript"">")
  
  Send("function toggleAccum() {")
  Send("  var elem = document.getElementById(""EnableAccum"");")
  Send("  var elemDiv = document.getElementById(""accumulation"");")
  Send("  var elemMin = document.getElementById(""accummin"");")
  Send("  var elemLimit = document.getElementById(""accumlimit"");")
  Send("  var elemPeriod = document.getElementById(""accumperiod"");")
  Send("  var elemMinPurch = document.getElementById(""MinPurchAmt"");")
  Send("  var elemMinPurchlbl = document.getElementById(""MinPurchlbl"");")

  Send("  if (elem != null && elemDiv != null) { ")
  Send("    if (elem.checked) { ")
  Send("      elemDiv.style.display = '';")
  Send("      elemMinPurch.style.display = 'none';")
  Send("      elemMinPurchlbl.style.display = 'none';")
  Send("    } else { ")
  Send("      elemDiv.style.display = 'none';")
  Send("      elemMinPurch.style.display = '';")
  Send("      elemMinPurchlbl.style.display = '';")
  Send("      if (elemMin != null) { elemMin.value = '0'; }")
  Send("      if (elemLimit != null) { elemLimit.value = '0'; }")
  Send("      if (elemPeriod != null) { elemPeriod.value = '0'; }")
  Send("      if (elemMinPurch != null) { elemMinPurch.value = '0'; }")
  Send("    }")
  Send("  }")
  Send("} ")
  
  Send("function ChangeAccum() {")
  If Not EnableAccum Then
    Send("  document.location = 'CPEoffer-con-product.aspx?OfferID=" & OfferID & "&EnableAccum=1&Disqualifier=" & IIf(Disqualifier, 1, 0) & IIf(IncentiveProdGroupID > 0, "&IncentiveProductGroupID=" & IncentiveProdGroupID, "") & "';")
  ElseIf EnableAccum Then
    Send("  document.location = 'CPEoffer-con-product.aspx?OfferID=" & OfferID & "&EnableAccum=0&Disqualifier=" & IIf(Disqualifier, 1, 0) & IIf(IncentiveProdGroupID > 0, "&IncentiveProductGroupID=" & IncentiveProdGroupID, "") & "';")
  End If
  Send("} ")
  
  Send("function ChangeParentDocument() { ")
  If (EngineID = 3) Then
    Send("  opener.location = 'web-offer-con.aspx?OfferID=" & OfferID & "'; ")
  ElseIf (EngineID = 5) Then
    Send("  opener.location = 'email-offer-con.aspx?OfferID=" & OfferID & "'; ")
  ElseIf (EngineID = 6) Then
    Send("  opener.location = 'CAM/CAM-offer-con.aspx?OfferID=" & OfferID & "'; ")
  Else
    Send("  opener.location = 'CPEoffer-con.aspx?OfferID=" & OfferID & "'; ")
  End If
  Send("} ")
  Send("</script>")
  Send_HeadEnd()
  
  If (isTemplate) Then
    Send_BodyBegin(12)
  Else
    Send_BodyBegin(2)
  End If
  If (Logix.UserRoles.AccessOffers = False AndAlso Not isTemplate) Then
    Send_Denied(2, "perm.offers-access")
    GoTo done
  End If
  If (Logix.UserRoles.AccessTemplates = False AndAlso isTemplate) Then
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
<form action="#" id="mainform" name="mainform" onsubmit="return saveForm();">
  <div id="intro">
    <span id="hiddenVals"></span>
    <input type="hidden" id="OfferID" name="OfferID" value="<% sendb(OfferID) %>" />
    <input type="hidden" id="Name" name="Name" value="<% sendb(Name) %>" />
    <input type="hidden" id="ConditionID" name="ConditionID" value="<% Sendb(ConditionID) %>" />
    <input type="hidden" id="IncentiveProductGroupID" name="IncentiveProductGroupID" value="<% Sendb(IncentiveProdGroupID) %>" />
    <input type="hidden" id="RestrictExcluding" name="RestrictExcluding" value="<% Sendb(RestrictExcluding) %>" />
    <%
      If Not Disqualifier Then
        Sendb("<input type=""hidden"" id=""Disqualifier"" name=""Disqualifier"" value=""0"" />")
      Else
        Sendb("<input type=""hidden"" id=""Disqualifier"" name=""Disqualifier"" value=""1"" />")
      End If
    %>
    <input type="hidden" id="roid" name="roid" value="<%sendb(roid) %>" />
    <input type="hidden" id="IsTemplate" name="IsTemplate" value="<% 
          If (IsTemplate) Then
          Sendb("IsTemplate")
          Else
          Sendb("Not")
          End If
          %>" />
    <%
      If (isTemplate) Then
        Sendb("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID)
      Else
        Sendb("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID)
      End If
      If Not Disqualifier Then
        Send(" " & StrConv(Copient.PhraseLib.Lookup("term.productcondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
      Else
        Send(" " & StrConv(Copient.PhraseLib.Lookup("term.productdisqualifier", LanguageID), VbStrConv.Lowercase) & "</h1>")
      End If
    %>
    <div id="controls">
      <% If (isTemplate) Then%>
      <span class="temp">
        <input type="checkbox" class="tempcheck" id="Disallow_Edit" name="Disallow_Edit"<% if(disallow_edit)then sendb(" checked=""checked""") %> />
        <label for="Disallow_Edit">
          <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
        </label>
      </span>
      <% End If%>
      <%
        If Not IsTemplate Then
          If (Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Edit)) Then
            If Not HasBundleDiscount Then
              If IncentiveProdGroupID = 0 Then
                Send_Save(" onclick=""this.style.visibility='hidden';""")
              Else
                Send_Save()
              End If
            End If
          End If
        Else
          If (Logix.UserRoles.EditTemplates) Then
            If Not HasBundleDiscount Then
              If IncentiveProdGroupID = 0 Then
                Send_Save(" onclick=""this.style.visibility='hidden';""")
              Else
                Send_Save()
              End If
            End If
          End If
        End If
      %>
    </div>
  </div>
  <div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <div id="column1">
      <div class="box" id="selector">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.productgroup", LanguageID))%>
          </span>
          <% If (isTemplate) Then%>
          <span class="tempRequire">
            <input type="checkbox" class="tempcheck" id="require_pg" name="require_pg" onclick="handleRequiredToggle();"<% if(RequirePG)then sendb(" checked=""checked""") %> />
            <label for="require_pg">
              <% Sendb(Copient.PhraseLib.Lookup("term.required", LanguageID))%>
            </label>
          </span>
          <% ElseIf (FromTemplate And RequirePG) Then%>
          <span class="tempRequire">
            <%Sendb("*" & Copient.PhraseLib.Lookup("term.required", LanguageID))%>
          </span>
          <% End If%>
        </h2>
        <input type="radio" id="functionradio1" name="functionradio" checked="checked"<% sendb(disabledattribute) %> /><label for="functionradio1"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
        <input type="radio" id="functionradio2" name="functionradio"<% sendb(disabledattribute) %> /><label for="functionradio2"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
        <%--<input type="text" class="medium" id="functioninput" name="functioninput" maxlength="100" onkeyup="handleKeyUp(200);" value=""<% sendb(disabledattribute) %> /><br />--%>
        <input type="text" class="medium" id="functioninput" name="functioninput" maxlength="100" onkeyup="javascript:xmlPostTimer('OfferFeeds.aspx','ProductGroups');" value=""<% sendb(disabledattribute) %> /><br />
        <div id="searchLoadDiv" style="display:block;">&nbsp;</div>
        <div id="pgList">
          <select class="longer" id="functionselect" name="functionselect" size="12"<% sendb(disabledattribute) %>>
            <%
              Dim topString As String = ""
              If RECORD_LIMIT > 0 Then topString = "top " & RECORD_LIMIT
              MyCommon.QueryStr = "select " & topString & " ProductGroupID, Name from ProductGroups where ProductGroupID is not null " & _
                                   "and Deleted=0 and NonDiscountableGroup=0 and PointsNotApplyGroup=0 "
              If Disqualifier Then
                MyCommon.QueryStr &= " and ProductGroupID <> 1  and ProductGroupID not in " & _
                                     "   (select ProductGroupID from CPE_IncentiveProductGroups with (NoLock) where Deleted=0 and RewardOptionID=" & roid & " and Disqualifier=0 and ExcludedProducts=0)"
              Else
                MyCommon.QueryStr &= " and ProductGroupID not in " & _
                                     "   (select ProductGroupID from CPE_IncentiveProductGroups with (NoLock) where Deleted=0 and RewardOptionID=" & roid & " and Disqualifier=1 and ExcludedProducts=0)"
              End If
              MyCommon.QueryStr &= " order by AnyProduct desc, ProductGroupID desc, Name asc"
              
              rst = MyCommon.LRT_Select
              For Each row In rst.Rows
                If MyCommon.NZ(row.Item("ProductGroupID"), 0) = 1 Then
                  Send("<option value=""1"" style=""color: brown; font-weight: bold;"">" & Copient.PhraseLib.Lookup("term.anyproduct", LanguageID) & "</option>")
                Else
                  Send("<option value=""" & MyCommon.NZ(row.Item("ProductGroupID"), 0) & """ title=""" & MyCommon.NZ(row.Item("Name"), "") & """>" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
                End If
              Next
            %>
          </select>
        </div>
        <%If (RECORD_LIMIT > 0) Then
            Send(Copient.PhraseLib.Lookup("groups.display", LanguageID) & ": " & RECORD_LIMIT.ToString() & "<br />")
          End If
        %>
        <br class="half" />
        <b><% Sendb(Copient.PhraseLib.Lookup("term.selectedproducts", LanguageID))%>:</b><br />
        <input type="button" class="regular select" id="select1" name="select1" value="&#9660; <% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID))%>" onclick="handleSelectClick('select1');"<% sendb(disabledattribute) %> />&nbsp;
        <input type="button" class="regular deselect" id="deselect1" name="deselect1" value="<% Sendb(Copient.PhraseLib.Lookup("term.deselect", LanguageID))%> &#9650;" disabled="disabled" onclick="handleSelectClick('deselect1');" /><br />
        <br class="half" />
        <select class="longer" id="selected" name="selected" multiple="multiple" size="2"<% sendb(disabledattribute) %>>
          <%
            'If ProdID > 0 And ExProdID = -1 Then
            If IncentiveProdGroupID > 0 Then
              ' alright lets find the currently selected groups on page load
              MyCommon.QueryStr = "select Name from ProductGroups where ProductGroupID=" & ProdID
              rst = MyCommon.LRT_Select
              If rst.Rows.Count > 0 Then
                If ProdID = 1 Then
                  Send("<option value=""1"" style=""color: brown; font-weight: bold;"">" & Copient.PhraseLib.Lookup("term.anyproduct", LanguageID) & "</option>")
                Else
                  Send("<option value=""" & ProdID & """>" & MyCommon.NZ(rst.Rows(0).Item("Name"), "") & "</option>")
                End If

              End If
            Else
              If Request.QueryString("selGroups") <> "" Then
                MyCommon.QueryStr = "select Name from ProductGroups where ProductGroupID=" & MyCommon.Extract_Val(Request.QueryString("selGroups"))
                rst = MyCommon.LRT_Select
                If rst.Rows.Count > 0 Then
                  Send("<option value=""" & MyCommon.Extract_Val(Request.QueryString("selGroups")) & """>" & MyCommon.NZ(rst.Rows(0).Item("Name"), "") & "</option>")
                End If
              End If
            End If
          %>
        </select>
        <br />
        <%
          If Disqualifier Then
            Send("<div style=""display:none;"">")
          End If
        %>
        <br class="half" />
        <b><% Sendb(Copient.PhraseLib.Lookup("term.excludedproducts", LanguageID))%>:</b><br />
        <input type="button" class="regular select" name="select2" id="select2" value="&#9660; <% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID))%>" disabled="disabled" onclick="handleSelectClick('select2');"<% sendb(disabledattribute) %> />&nbsp;
        <input type="button" class="regular deselect" name="deselect2" id="deselect2" value="<% Sendb(Copient.PhraseLib.Lookup("term.deselect", LanguageID))%> &#9650;" disabled="disabled" onclick="handleSelectClick('deselect2');"<% sendb(disabledattribute) %> /><br />
        <br class="half" />
        <select class="longer" id="excluded" name="excluded" size="2"<% sendb(disabledattribute) %>>
          <%
            If IncentiveProdGroupID > 0 And ProdID = 1 Then
              ' alright lets find the currently excluded groups on page load
              MyCommon.QueryStr = "select PG.ProductGroupID,Name from CPE_IncentiveProductGroups as IPG left join ProductGroups as PG " & _
                                  " on PG.ProductGroupID=IPG.ProductGroupID where RewardOptionID=" & roid & _
                                  " and IPG.deleted=0 and ExcludedProducts=1 and IPG.ProductGroupID is not null"
              If Not Disqualifier Then
                MyCommon.QueryStr = MyCommon.QueryStr & " and Disqualifier=0;"
              Else
                MyCommon.QueryStr = MyCommon.QueryStr & " and Disqualifier=1;"
              End If
              rst = MyCommon.LRT_Select
              For Each row In rst.Rows
                      Send("<option value=""" & MyCommon.NZ(row.Item("ProductGroupID"), 0) & """ title=""" & MyCommon.NZ(row.Item("Name"), "") & """>" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
              Next
            Else
              If Request.QueryString("exGroups") <> "" Then
                MyCommon.QueryStr = "select Name from ProductGroups where ProductGroupID=" & MyCommon.Extract_Val(Request.QueryString("exGroups"))
                rst = MyCommon.LRT_Select
                If rst.Rows.Count > 0 Then
                  Send("<option value=""" & MyCommon.Extract_Val(Request.QueryString("exGroups")) & """>" & MyCommon.NZ(rst.Rows(0).Item("Name"), "") & "</option>")
                End If
              End If
            End If
          %>
        </select>
        <%
          If Disqualifier Then
            Send("</div>")
          End If
        %>
        <hr class="hidden" />
      </div>
    </div>
    
    <div id="gutter">
    </div>
    
    <%If (Disqualifier = False) Then%>
    <div id="column2">
      <div class="box" id="value">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.value", LanguageID))%>
          </span>
        </h2>
        <div id="results">
          <%
              If Request.QueryString("selGroups") <> "" Then
                  limits = Request.QueryString("selGroups")
              ElseIf ProdID = 1 AndAlso ExProdID <> -1 Then
                  limits = ExProdID
              Else
                  If ProdID <> 0 Then
                      limits = ProdID
                  Else
                      limits = -1
                  End If
              End If

              ' Preload the unit types
              MyCommon.QueryStr = "select UnitTypeID, PhraseID, Description from CPE_UnitTypes UT with (NoLock) where MultiUOMState in (-1,0) and UnitTypeID <> 10;"
              rst3 = MyCommon.LRT_Select
              If IncentiveProdGroupID <> 0 Then
                  If (limits <> -1) OrElse (limits = -1 AndAlso isTemplate AndAlso RequirePG) OrElse (limits = -1 AndAlso FromTemplate AndAlso RequirePG) Then
                      Send("<table summary=""" & Copient.PhraseLib.Lookup("term.groups", LanguageID) & """>")
                      Send("  <thead>")
                      Send("    <tr>")
                      Send("      <th class=""th-group"" scope=""col"" id=""UnitTypeDesc"">" & Copient.PhraseLib.Lookup("term.quantity", LanguageID) & "</th>")
                      Send("      <th class=""th-unit"" scope=""col"">" & Copient.PhraseLib.Lookup("term.unit", LanguageID) & "</th>")
                      Send("    </tr>")
                      Send("  </thead>")
                      Send("  <tbody>")
                      ' Determine limits
                      ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                      MyCommon.QueryStr = "select IncentiveProductGroupID,ProductGroupID,QtyForIncentive,QtyUnitType,AccumMin,AccumLimit,AccumPeriod,UniqueProduct,MinPurchAmt " & _
                                          "from CPE_IncentiveProductGroups with (NoLock) " & _
                                          "where Deleted=0 and RewardOptionID=" & roid & " and IncentiveProductGroupID=" & IncentiveProdGroupID & ";"
                      rst2 = MyCommon.LRT_Select
                      If rst2.Rows.Count > 0 Then
                          If MyCommon.NZ(rst2.Rows(0).Item("ProductGroupID"), 0) > 0 Then
                              MyCommon.QueryStr = "select Name from ProductGroups with (NoLock) where ProductGroupID=" & rst2.Rows(0).Item("ProductGroupID") & ";"
                              rst = MyCommon.LRT_Select
                              If rst.Rows.Count > 0 Then
                                  CleanGroupName = MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 15)
                              End If
                          End If
                          Qty = MyCommon.NZ(rst2.Rows(0).Item("QtyForIncentive"), 1)
                          Type = MyCommon.NZ(rst2.Rows(0).Item("QtyUnitType"), 1)
                          AccumMin = MyCommon.NZ(rst2.Rows(0).Item("AccumMin"), 0)
                          AccumLimit = MyCommon.NZ(rst2.Rows(0).Item("AccumLimit"), 0)
                          AccumPeriod = MyCommon.NZ(rst2.Rows(0).Item("AccumPeriod"), 0)
                          UniqueChecked = MyCommon.NZ(rst2.Rows(0).Item("UniqueProduct"), False)
                          MinPurchAmt = MyCommon.NZ(rst2.Rows(0).Item("MinPurchAmt"), 0)
                          If Type = 1 Then
                              IsItem = True
                          ElseIf Type = 2 Then
                              IsDollar = True
                          ElseIf Type = 4 Then
                              IsQty1 = True
                          ElseIf Type = 3 Then
                              isWtVol = True
                          ElseIf Type = 10 Then
                              isSingleItemPriceThreshold = True
                          End If
                      Else
                          Qty = 1
                          Type = 1
                          AccumMin = 0
                          AccumLimit = 0
                          AccumPeriod = 0
                      End If
                      Send("    <tr " & Shaded & ">")
                      Send("      <td>")
                      If TierLevels = 1 Or Disqualifier Then
                          If rst2.Rows.Count > 0 Then
                              MyCommon.QueryStr = "Select Quantity from CPE_IncentiveProductGroupTiers with (NoLock) where RewardOptionID=" & roid & " and TierLevel=1 and IncentiveProductGroupID=" & rst2.Rows(0).Item("IncentiveProductGroupID")
                              TierDT = MyCommon.LRT_Select()
                              If TierDT.Rows.Count > 0 Then
                                  TierQty = MyCommon.NZ(TierDT.Rows(0).Item("Quantity"), 0)
                              Else
                                  TierQty = 0
                              End If
                          Else
                              TierQty = 0
                          End If
                          If IsItem Then
                              TierQty = Math.Truncate(TierQty)
                              MinPurchAmt = Math.Truncate(MinPurchAmt)
                          ElseIf IsDollar Then
                              TierQty = Math.Round(TierQty, 2)
                              MinPurchAmt = Math.Round(MinPurchAmt, 2)
                          ElseIf IsQty1 Then
                              TierQty = Math.Round(TierQty, 2)
                          ElseIf isWtVol Then
                              TierQty = Math.Round(TierQty, 3)
                              MinPurchAmt = Math.Round(MinPurchAmt, 3)
                          ElseIf isSingleItemPriceThreshold Then
                                  TierQty = Math.Round(TierQty, 2)
                                  MinPurchAmt = Math.Round(MinPurchAmt, 2)
                              End If
                          Sendb("        <input type=""text"" class=""shorter"" maxlength=""9"" name=""t1_limit"" id=""t1_limit"" value=""" & TierQty & """ />")
                      Else
                          For t = 1 To TierLevels
                              If rst2.Rows.Count > 0 Then
                                  MyCommon.QueryStr = "Select Quantity from CPE_IncentiveProductGroupTiers with (NoLock) where RewardOptionID=" & roid & " and TierLevel=" & t & " and IncentiveProductGroupID=" & rst2.Rows(0).Item("IncentiveProductGroupID")
                                  TierDT = MyCommon.LRT_Select()
                                  If TierDT.Rows.Count > 0 Then
                                      TierQty = MyCommon.NZ(TierDT.Rows(0).Item("Quantity"), 0)
                                  Else
                                      TierQty = 0
                                  End If
                              Else
                                  TierQty = 0
                              End If
                              'Save the values that were incorrect
                              If TierQty = 0 Then
                                  If Request.QueryString("t" & t & "_limit") <> "" Then
                                      TierQty = Request.QueryString("t" & t & "_limit")
                                  End If
                              End If
                              If IsItem Then
                                  TierQty = Math.Truncate(TierQty)
                                  MinPurchAmt = Math.Truncate(MinPurchAmt)
                              ElseIf IsDollar Then
                                  TierQty = Math.Round(TierQty, 2)
                                  MinPurchAmt = Math.Round(MinPurchAmt, 2)
                              ElseIf IsQty1 Then
                                  TierQty = Math.Round(TierQty, 2)
                              ElseIf isWtVol Then
                                  TierQty = Math.Round(TierQty, 3)
                                  MinPurchAmt = Math.Round(MinPurchAmt, 3)
                              ElseIf isSingleItemPriceThreshold Then
                                  TierQty = Math.Round(TierQty, 2)
                                  MinPurchAmt = Math.Round(MinPurchAmt, 2)
                              End If
                              If TierLevels > 1 Then
                                  Sendb("        <label style=""margin-left:15px;margin-right:15px;"" for=""t" & t & "_limit"">" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</label><input type=""text"" class=""shorter"" maxlength=""9"" name=""t" & t & "_limit"" id=""t" & t & "_limit"" value=""" & TierQty & """ />" & IIf(TierLevels > 1, "<br />", ""))
                              End If
                          Next
                      End If
                      Send("</td>")
                      Send("      <td>")
                      Send("        <select name=""select"" id=""select"" onchange=""ChangeUnit(this.options[this.selectedIndex].value);"" >")
                      For Each row3 In rst3.Rows
                          Sendb("          <option")
                          If (Type = row3.Item("UnitTypeID")) Then
                              Sendb(" selected=""selected""")
                          End If
                          Sendb(" value=""" & row3.Item("UnitTypeID") & """>" & Copient.PhraseLib.Lookup(row3.Item("PhraseID"), LanguageID, MyCommon.NZ(row3.Item("Description"), "Qty 1 at Price")))
                          Send("    </option>")
                      Next
                      Send("        </select>")
                      Send("      </td>")
                      Send("    </tr>")
                      Send("    <tr " & Shaded & ">")
                      Send("      <td></td>")
                      Send("      <td colspan=""2"">")
                      Send("        <input type=""checkbox"" id=""Unique"" name=""Unique""" & IIf(UniqueChecked And Not EnableAccum, " checked=""checked""", "") & " value=""1""" & IIf(Type > 1 OrElse EnableAccum OrElse IsQty1, " disabled=""disabled""", "") & "/><label for=""Unique"" id=""SelectUnique"">" & Copient.PhraseLib.Lookup("term.uniqueproduct", LanguageID) & "</label>")
                      Send("      </td>")
                      Send("    </tr>")
                      Send("    <tr " & Shaded & " id=""MinPurch""" & IIf(IsQty1, " style=""display:none;"" ", "") & ">")
                      Send("      <td><label id=""MinPurchlbl"" for=""MinPurchAmt"" " & IIf(EnableAccum OrElse TierLevels > 1, " style=""display:none;"" ", "") & " >" & Copient.PhraseLib.Lookup("term.minimumpurchase", LanguageID) & "</label></td>")
                      Send("      <td><input type=""text"" class=""short"" id=""MinPurchAmt"" name=""MinPurchAmt"" " & IIf(EnableAccum Or TierLevels > 1, " style=""display:none;"" ", "") & " value=""" & MinPurchAmt & """ maxlength=""9"" />")
                      Send("    </tr>")
                      rst2 = Nothing
                      If Shaded = " class=""shaded""" Then
                          Shaded = ""
                      Else
                          Shaded = " class=""shaded"""
                      End If
                      ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                      Send("  </tbody>")
                      Send("</table>")
                      Send("<br />")
                  End If
              Else
                  Send("<table summary=""" & Copient.PhraseLib.Lookup("term.groups", LanguageID) & """>")
                  Send("  <thead>")
                  Send("    <tr>")
                  Send("      <th class=""th-group"" scope=""col"" id=""UnitTypeDesc"">" & Copient.PhraseLib.Lookup("term.quantity", LanguageID) & "</th>")
                  Send("      <th class=""th-unit"" scope=""col"">" & Copient.PhraseLib.Lookup("term.unit", LanguageID) & "</th>")
                  Send("    </tr>")
                  Send("  </thead>")
                  Send("  <tbody>")
                  Send("    <tr " & Shaded & ">")
                  Send("      <td>")
                  If Disqualifier Then
                      If Request.QueryString("t1_limit") <> "" Then
                          Sendb("        <input type=""text"" class=""shorter"" maxlength=""9"" name=""t1_limit"" id=""t1_limit"" value=""" & MyCommon.Extract_Val(Request.QueryString("t1_limit")) & """ />")
                      Else
                          Sendb("        <input type=""text"" class=""shorter"" maxlength=""9"" name=""t1_limit"" id=""t1_limit"" value=""0"" />")
                      End If
                  Else
                      If TierLevels = 1 Then
                          If Request.QueryString("t1_limit") <> "" Then
                              Sendb("        <input type=""text"" class=""shorter"" maxlength=""9"" name=""t1_limit"" id=""t1_limit"" value=""" & MyCommon.Extract_Val(Request.QueryString("t1_limit")) & """ />")
                          Else
                              Sendb("        <input type=""text"" class=""shorter"" maxlength=""9"" name=""t1_limit"" id=""t1_limit"" value=""0"" />")
                          End If
                      Else
                          For t = 1 To TierLevels
                              If TierLevels > 1 Then
                                  If Request.QueryString("t" & t & "_limit") <> "" Then
                                      Sendb("        <label style=""margin-left:15px;margin-right:15px;"" for=""t" & t & "_limit"">" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</label><input type=""text"" class=""shorter"" maxlength=""9"" name=""t" & t & "_limit"" id=""t" & t & "_limit"" value=""" & MyCommon.Extract_Val(Request.QueryString("t" & t & "_limit")) & """ />" & IIf(TierLevels > 1, "<br />", ""))
                                  Else
                                      Sendb("        <label style=""margin-left:15px;margin-right:15px;"" for=""t" & t & "_limit"">" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</label><input type=""text"" class=""shorter"" maxlength=""9"" name=""t" & t & "_limit"" id=""t" & t & "_limit"" value=""0"" />" & IIf(TierLevels > 1, "<br />", ""))
                                  End If
                              End If
                          Next
                      End If
                  End If
                  Send("</td>")
                  Send("      <td><select name=""select"" id=""select"" onchange=""ChangeUnit(this.options[this.selectedIndex].value);"">")
                  For Each row3 In rst3.Rows
                      Sendb("<option")
                      If Request.QueryString("select") <> "" Then
                          Type = MyCommon.Extract_Val(Request.QueryString("select"))
                      Else
                          Type = 1
                      End If
                      If (Type = row3.Item("UnitTypeID")) Then
                          Sendb(" selected=""selected""")
                      End If
                      Sendb(" value=""" & row3.Item("UnitTypeID") & """>" & Copient.PhraseLib.Lookup(row3.Item("PhraseID"), LanguageID, MyCommon.NZ(row3.Item("Description"), "Qty 1 at Price")))
                      Send("</option>")
                  Next
                  Send("</select>")
                  Send("      </td>")
                  Send("    </tr>")
                  Send("    <tr " & Shaded & ">")
                  Send("      <td></td>")
                  Send("      <td colspan=""2"">")
                  Send("        <input type=""checkbox"" id=""Unique"" name=""Unique""" & IIf(Request.QueryString("Unique") <> "" And Not EnableAccum, " checked=""checked""", "") & " value=""1"" " & IIf(Type > 1 OrElse EnableAccum, " disabled ", "") & "/><label for=""Unique"" id=""SelectUnique"">" & Copient.PhraseLib.Lookup("term.uniqueproduct", LanguageID) & "</label>")
                  Send("      </td>")
                  Send("    </tr>")
                  Send("    <tr " & Shaded & " id=""MinPurch"">")
                  Send("      <td><label id=""MinPurchlbl"" for=""MinPurchAmt"" " & IIf(EnableAccum Or TierLevels > 1, " style=""display:none;"" ", "") & " >" & Copient.PhraseLib.Lookup("term.minimumpurchase", LanguageID) & "</label></td>")
                  Send("      <td><input type=""text"" class=""short"" id=""MinPurchAmt"" name=""MinPurchAmt"" " & IIf(EnableAccum Or TierLevels > 1, " style=""display:none;"" ", "") & " value=""" & MinPurchAmt & """ maxlength=""9"" />")
                  Send("    </tr>")
                  rst2 = Nothing
                  If Shaded = " class=""shaded""" Then
                      Shaded = ""
                  Else
                      Shaded = " class=""shaded"""
                  End If
                  Send("  </tbody>")
                  Send("</table>")
                  Send("<br />")
              End If

              'Rounding
              If MyCommon.Fetch_CPE_SystemOption(89) = "1" And EngineID = 2 Then
                  Send("<span id=""roundingspan""" & IIf(Type = 1 OrElse Type = 4, " style=""display:none;""", "") & ">")
                  Send("<input type=""checkbox"" id=""rounding"" name=""rounding"" value=""on""" & IIf(Rounding, " checked=""checked""", "") & " /><label for=""rounding"">" & Copient.PhraseLib.Lookup("CPEoffer-con-product.rounding", LanguageID) & "</label><br />")
                  Send("<br class=""half"" />")
                  Send("</span>")
              End If

              'Accumulation checkbox
              'If (ShowAccum And Not Disqualifier And Not HasDisqualifier) AndAlso (TierLevels = 1) AndAlso (EngineID <> 6) Then 
              If (ShowAccum And Not Disqualifier And Not HasDisqualifier And Not hasAnyCustomer) AndAlso (TierLevels = 1) Then
                  Send("<input type=""checkbox"" id=""EnableAccum"" onclick=""toggleAccum();"" name=""EnableAccum"" value=""1"" " & IIf(EnableAccum, "checked=""checked""", "") & IIf(IsQty1, " style=""display:none;""", "") & " /><label id=""AccumChecklbl"" for=""EnableAccum""" & IIf(IsQty1, " style=""display:none;""", "") & ">" & Copient.PhraseLib.Lookup("term.enable", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.accumulation", LanguageID), VbStrConv.Lowercase) & "</label>")
                  Send("<br />")
                  ' Do the accumulation stuff here
                  Send("<div id=""accumulation"" style=""width:310px;" & IIf(EnableAccum AndAlso Not IsQty1, "", "display:none;") & """>")
                  Send("<table summary=""" & Copient.PhraseLib.Lookup("term.accumulation", LanguageID) & """>")
                  Send("  <thead>")
                  Send("    <tr>")
                  Send("      <th class=""th-minimum"" scope=""col"">" & Copient.PhraseLib.Lookup("term.minimum", LanguageID) & "</th>")
                  Send("      <th class=""th-limit"" scope=""col"">" & Copient.PhraseLib.Lookup("term.limit", LanguageID) & "</th>")
                  Send("      <th class=""th-period"" scope=""col"">" & Copient.PhraseLib.Lookup("term.period", LanguageID) & "</th>")
                  Send("    </tr>")
                  Send("  </thead>")
                  Send("  <tbody>")
                  Send("    <tr>")
                  If IsItem Then
                      AccumMin = Math.Truncate(AccumMin)
                  ElseIf IsDollar Then
                      AccumMin = Math.Round(AccumMin, 2)
                  End If
                  If Not EnableAccum Then AccumMin = 0
                  Send("      <td><input type=""text"" class=""shorter"" maxlength=""9"" name=""accummin"" id=""accummin"" value=""" & AccumMin & """ /></td>")
                  If IsItem Then
                      AccumLimit = Math.Truncate(AccumLimit)
                  ElseIf IsDollar Then
                      AccumLimit = Math.Round(AccumLimit, 2)
                  End If
                  If Not EnableAccum Then
                      AccumLimit = 0
                      AccumPeriod = 0
                  End If
                  Send("      <td><input type=""text"" class=""shorter"" maxlength=""9"" name=""accumlimit"" id=""accumlimit"" value=""" & AccumLimit & """ /> " & StrConv(Copient.PhraseLib.Lookup("term.per", LanguageID), VbStrConv.Lowercase) & "</td>")
                  Send("      <td><input type=""text"" class=""shorter"" maxlength=""9"" name=""accumperiod"" id=""accumperiod"" value=""" & AccumPeriod & """ /> " & StrConv(Copient.PhraseLib.Lookup("term.days", LanguageID), VbStrConv.Lowercase) & "</td>")
                  Send("    </tr>")
                  Send("  </tbody>")
                  Send("</table>")
                  Send("</div>")
                  If EnableAccum Then ProductComboID = 0
              End If
          %>
        </div>
        <br class="half" />
        <hr class="hidden" />
      </div>
    </div>
    <%Else%>
    <input type="hidden" name="t1_limit" id="t1_limit" value="0" />
    <%End If %>
  </div>
</form>

<script runat="server">
  Function IsValidEntry(ByRef MyCommon As Copient.CommonInc) As Boolean
    Dim Ids() As String
    Dim i As Integer
    Dim Qty As Decimal
    Dim Type As Integer
    Dim AccumMin As Decimal
    Dim AccumLimit As Decimal
    Dim AccumPeriod As Integer
    Dim MinPurchAmt As Decimal
    Dim IsValid As Boolean = True
    Dim ProdID As Integer = 0
    
    If (Request.QueryString("selGroups") <> "") Then
      'Ids = Request.QueryString("selGroups").Split(",")
      ProdID = MyCommon.Extract_Val(Request.QueryString("selGroups"))
      
      ' we need to do some work to set the limit values if there are any, otherwise just set to 0
      ' in theory there should be one set of limit values for each selected groups and possibly an accumulation infos
      Qty = IIf(MyCommon.Extract_Val(Request.QueryString("limit")) <> "", MyCommon.Extract_Val(Request.QueryString("limit")), 0)
      Type = IIf(MyCommon.Extract_Val(Request.QueryString("select")) <> "", MyCommon.Extract_Val(Request.QueryString("select")), 0)
      AccumMin = IIf(MyCommon.Extract_Val(Request.QueryString("accummin")) <> "", MyCommon.Extract_Val(Request.QueryString("accummin")), 0)
      AccumLimit = IIf(MyCommon.Extract_Val(Request.QueryString("accumlimit")) <> "", MyCommon.Extract_Val(Request.QueryString("accumlimit")), 0)
      AccumPeriod = IIf(MyCommon.Extract_Val(Request.QueryString("accumperiod")) <> "", MyCommon.Extract_Val(Request.QueryString("accumperiod")), 0)
      MinPurchAmt = IIf(MyCommon.Extract_Val(Request.QueryString("MinPurchAmt")) <> "", MyCommon.Extract_Val(Request.QueryString("MinPurchAmt")), 0)
      IsValid = IsValid AndAlso IsProperFormat(Type, Qty)
      IsValid = IsValid AndAlso IsProperFormat(Type, AccumMin)
      IsValid = IsValid AndAlso IsProperFormat(Type, AccumLimit)
      IsValid = IsValid AndAlso IsProperFormat(Type, AccumPeriod)
      IsValid = IsValid AndAlso IsProperFormat(Type, MinPurchAmt)
      'If (Not IsValid) Then Exit For
    End If
    
    Return IsValid
  End Function
  
  Function IsProperFormat(ByVal UnitType As Integer, ByVal Value As Double) As Boolean
    Dim FormatOk As Boolean = True
    Dim StrValue As String = Value.ToString()
    Dim DecPtPos, CharAfterDec As Integer
    
    DecPtPos = StrValue.IndexOf(".")
    If (DecPtPos > -1) Then
      CharAfterDec = (StrValue.Length - (DecPtPos + 1))
    End If
    
    Select Case UnitType
      Case 1 ' ###,##0
        FormatOk = (DecPtPos = -1)
      Case 2 ' ###,##0.00
        FormatOk = (DecPtPos = -1) OrElse (CharAfterDec >= 0 AndAlso CharAfterDec <= 2)
      Case 3 ' ###,##0.000
        FormatOk = (DecPtPos = -1) OrElse (CharAfterDec >= 0 AndAlso CharAfterDec <= 3)
      Case Else
        FormatOk = True
    End Select
    
    Return FormatOk
  End Function
  
  Function IsExcludingRestricted(ByRef MyCommon As Copient.CommonInc, ByVal roid As Integer) As Boolean
    'If we have two product condition with AND, and out of two one is ANYProduct, then excluded product group cannot be selected
    Dim AnyProductGrpExist As Boolean = False
     Dim ANDConditionExist As Boolean = False
     Dim Spec_ProdGroup As Boolean = False
    Dim dt As DataTable
    Dim resEx As Boolean
    MyCommon.QueryStr = "select RO.ProductComboID, * from CPE_IncentiveProductGroups as IPG with (NoLock) left join CPE_RewardOptions as RO with (NoLock) on IPG.RewardOptionID=RO.RewardOptionID where IPG.Deleted=0 and IPG.ExcludedProducts=0 and IPG.RewardOptionID=" & roid & ";"
    dt = MyCommon.LRT_Select()
    If dt.Rows.Count >= 2 Then
         For Each row In dt.Rows
                If MyCommon.NZ(row.Item("ProductGroupID"), 0) = 1 Then
                    AnyProductGrpExist = True
                End If
                
                If MyCommon.NZ(row.Item("ProductGroupID"), 0) > 1 Then
                    Spec_ProdGroup = True
                End If
                
                If MyCommon.NZ(row.Item("ProductComboID"), 0) = 1 Or MyCommon.NZ(row.Item("ProductComboID"), 0) = 0 Then
                    ANDConditionExist = True
                End If
            Next
      If AnyProductGrpExist AndAlso ANDConditionExist AndAlso Spec_ProdGroup Then
        'disable the select button to add the excluded group
        resEx = True
      End If
    End If
    Return resEx
  End Function
</script>

<script type="text/javascript">
<% If (CloseAfterSave) Then %>
  window.close();
<% Else %> 
  if (document.getElementById("functionselect") != null) {
    fullSelect = document.getElementById("functionselect").cloneNode(true);
  }
  removeUsed(true);
  updateButtons();

<% End If %>
</script>
<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd("mainform", "functioninput")
  MyCommon = Nothing
  Logix = Nothing
%>
