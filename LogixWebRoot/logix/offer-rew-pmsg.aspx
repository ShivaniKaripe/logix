﻿﻿<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Text.RegularExpressions" %>
<%
  ' *****************************************************************************
  ' * FILENAME: offer-rew-pmsg.aspx 
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
    Dim MyCryptLib As New Copient.CryptLib
  Dim rst As DataTable
    Dim row As DataRow
    Dim rst1 As DataTable
  Dim rst2 As DataTable
  Dim row2 As DataRow
  Dim rst3 As DataTable
  Dim OfferID As Long
  Dim Name As String = ""
  Dim RewardID As String
  Dim ExcludedItem As Integer
  Dim SelectedItem As Integer
  Dim NumTiers As Integer
  Dim LinkID As Long
  Dim RewardAmountTypeID As Integer
  Dim TriggerQty As Integer
  Dim ApplyToLimit As Integer
  Dim DoNotItemDistribute As Boolean
  Dim ExItemLevelDist As Boolean
  Dim UseSpecialPricing As Boolean
  Dim SPRepeatAtOccur As Integer
  Dim MessageTypeID As Integer
  Dim PrintReceipt As Integer
  Dim PrintOnBack As Integer
  Dim CheckedStatus As Boolean
  Dim i As Integer
  Dim q As Integer
  Dim x As Integer
  Dim Tiered As Integer
  Dim SponsorID As Integer
  Dim PromoteToTransLevel As Boolean
  Dim RewardLimit As Integer
  Dim RewardLimitTypeID As Integer
  Dim ValueRadio As Integer
  Dim TransactionLevelSelected As Boolean = False
  Dim DistPeriod As Integer
  Dim VarID As Integer
  Dim Disallow_Edit As Boolean = True
  Dim IsTemplate As Boolean = False
  Dim OfferEngineID As Long = 0
  Dim DisabledAttribute As String = ""
  Dim CloseAfterSave As Boolean = False
  Dim PrinterWidthBuf As New StringBuilder()
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim BannersEnabled As Boolean = False
  Dim ExcentusInstalled As Boolean = False
  
  Dim bUseTemplateLocks As Boolean
  Dim bDisallowEditPg As Boolean = False
  Dim bDisallowEditSpon As Boolean = False
  Dim bDisallowEditMsg As Boolean = False
  Dim bDisallowEditDist As Boolean = False
  Dim bDisallowEditLimit As Boolean = False
  Dim bDisallowEditAdv As Boolean = False
  Dim sDisabled As String
  Dim AdvancedLimitID As Long
  Dim sLifetimePointsId As String

  Dim bRequireExternalIds As Boolean
  Dim MyExport As Copient.ExportXml = Nothing
  Dim lPromoVarId As Long
  Dim dt As DataTable
    Dim lExternalId As Long
  Dim bStoreUser As Boolean = False
  Dim sValidLocIDs As String = ""
  Dim sValidSU As String = ""
  Dim iLen As Integer = 0
    Dim SelectedProductGroupId As Integer
    Dim ExcludedProdGroupID As Integer
    Dim ByExistingPGSelector As Boolean = IIf(MyCommon.Fetch_SystemOption(222) = "0", True, False)
    Dim PagePostBack As Boolean = True
    Dim ByAddSingleProduct As Boolean = True
    Dim ShowAllItems As Boolean
    Dim GroupSize As Integer
    Dim rstItems As DataTable = Nothing
    Dim descriptionItem As String = String.Empty
    Dim ProductTypeID As Integer = 0
    Dim IDLength As Integer = 0
    Dim GName As String = ""
    Dim OfferStartDate As Date
    Dim ExtProductID As String = ""
    Dim prodDT As DataTable
    Dim Description As String = ""
    Dim outputStatus As Integer
    Dim tempProducts As String = ""
    Dim tempProductsList() As String = Nothing
    Dim maxLimit As Integer = 0
    Dim validItemList As List(Of String) = New List(Of String)
    Dim invalidItemList As List(Of String) = New List(Of String)
    Dim tempTableInsertStatement As StringBuilder = New StringBuilder()
    Dim upc As String = ""
    Dim ProductsWithoutDesc As Integer
    Dim ListBoxSize As Integer
    Dim ProductGroupID As Integer = 0
    Dim SearchProductGrouptext As String = Nothing
    Dim RadioButtonforStart As Boolean = False
    Dim RadioButtonForContain As Boolean = false
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "offer-rew-pmsg.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  MyCommon.Open_LogixXS()

  'Store User
  If(MyCommon.Fetch_CM_SystemOption(131) = "1") Then
    MyCommon.QueryStr = "select LocationID from StoreUsers where UserID=" & AdminUserID & ";"
    rst = MyCommon.LRT_Select
    iLen = rst.Rows.Count
    If iLen > 0 Then
      bStoreUser = True
      sValidSU = AdminUserID
      For i=0 to (iLen-1)
        If i=0 Then 
          sValidLocIDs = rst.Rows(0).Item("LocationID")
        Else 
          sValidLocIDs &= "," & rst.Rows(i).Item("LocationID")
        End If
      Next
    
      MyCommon.QueryStr = "select UserID from StoreUsers where LocationID in (" & sValidLocIDs & ") and NOT UserID=" & AdminUserID & ";"
      rst = MyCommon.LRT_Select
      iLen = rst.Rows.Count
      If iLen > 0 Then
        For i=0 to (iLen-1)
          sValidSU &= "," & rst.Rows(i).Item("UserID") 
        Next
      End If
    End If
  End If
  
  If (Request.QueryString("OfferID") <> "") Then
    OfferID = MyCommon.Extract_Val(Request.QueryString("OfferID"))
  Else
    OfferID = MyCommon.Extract_Val(Request.Form("OfferID"))
  End If
  If (Request.QueryString("RewardID") <> "") Then
    RewardID = Request.QueryString("RewardID")
  Else
    RewardID = Request.Form("RewardID")
  End If
  If (Request.QueryString("NumTiers") <> "") Then
    NumTiers = MyCommon.Extract_Val(Request.QueryString("NumTiers"))
  Else
    NumTiers = MyCommon.Extract_Val(Request.Form("NumTiers"))
  End If
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  sLifetimePointsId = MyCommon.NZ(MyCommon.Fetch_CM_SystemOption(37), "")
  bRequireExternalIds = (MyCommon.Fetch_CM_SystemOption(42) = "1")

  If (Request.Form("save") <> "") Then
    CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
  End If
  
  ' fetch the name
  MyCommon.QueryStr = "Select Name,IsTemplate,FromTemplate,EngineId,ProdStartDate from Offers with (NoLock) where OfferID=" & OfferID
  rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
        OfferStartDate = rst.Rows(0).Item("ProdStartDate")
        Name = rst.Rows(0).Item("Name")
        IsTemplate = MyCommon.NZ(rst.Rows(0).Item("IsTemplate"), False)
        If IsTemplate Then
            bUseTemplateLocks = False
        Else
            bUseTemplateLocks = MyCommon.NZ(rst.Rows(0).Item("FromTemplate"), False)
        End If
        OfferEngineID = rst.Rows(0).Item("EngineId")
    End If

  If (IsTemplate Or bUseTemplateLocks) Then
    MyCommon.QueryStr = "select Disallow_Edit,DisallowEdit1,DisallowEdit2,DisallowEdit3,DisallowEdit4," & _
                        "DisallowEdit5,DisallowEdit6,DisallowEdit7,DisallowEdit8,DisallowEdit9 " & _
                        "from OfferRewards with (NoLock) where RewardID=" & RewardID
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      Disallow_Edit = MyCommon.NZ(rst.Rows(0).Item("Disallow_Edit"), True)
      bDisallowEditPg = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit1"), False)
      bDisallowEditSpon = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit3"), False)
      bDisallowEditMsg = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit4"), False)
      bDisallowEditDist = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit6"), False)
      bDisallowEditLimit = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit7"), False)
      bDisallowEditAdv = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit9"), False)
      If bUseTemplateLocks Then
        If Disallow_Edit Then
          bDisallowEditPg = True
          bDisallowEditSpon = True
          bDisallowEditMsg = True
          bDisallowEditDist = True
          bDisallowEditLimit = True
          bDisallowEditAdv = True
        Else
          Disallow_Edit = bDisallowEditPg And bDisallowEditSpon And bDisallowEditMsg And _
                          bDisallowEditDist And bDisallowEditLimit And bDisallowEditAdv
        End If
      End If
    End If
  End If
  
    If Request.Form("pgselectortype") Is Nothing OrElse Request.Form("pgselectortype") = "" Then
        PagePostBack = False
    Else
        If Request.Form("pgselectortype") = "directadd" Then
            ByExistingPGSelector = False
		Else
            ByExistingPGSelector = True 		
        End If
    End If
    
    If Request.Form("prodaddselector") = "prodlistadd" Then
        ByAddSingleProduct = False
    Else
        ByAddSingleProduct = True
    End If
    
  MyCommon.QueryStr = "select PG.Name from OfferRewards ORWD with (NoLock) Inner Join ProductGroups PG with (nolock) on ORWD.productgroupid=PG.productgroupid where ORWD.RewardID=" & RewardID & " and ORWD.deleted=0;"
  rst = MyCommon.LRT_Select
  
  If rst.Rows.Count > 0 Then
    GName = MyCommon.NZ(rst.Rows(0).Item("Name"), "")
  Else
    MyCommon.QueryStr = "select rewardid from offerrewards with (NoLock) where offerid=" & OfferID & " and rewardtypeid=3"
    rst3 = MyCommon.LRT_Select
    If rst3.Rows.Count = 1 Then
      GName = OfferStartDate.ToString("yyyyMMdd") & Name & "PGRPMSG"
    ElseIf rst3.Rows.Count > 1 Then
      GName = OfferStartDate.ToString("yyyyMMdd") & Name & "PGRPMSG(" & rst3.Rows.Count - 1 & ")"
    End If
  End If
	
  Send_HeadBegin("term.offer", "term.pmsgreward", OfferID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
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
    Send("<script type=""text/javascript"">")
    Send("  function ChangeParentDocument() { return true; } ")
    Send("</script>")
    Send_Denied(1, "banners.access-denied-offer")
    Send_BodyEnd()
    GoTo done
  End If
  
  ' we need to determine our linkid for updates and tiered
  MyCommon.QueryStr = "select LinkID,Tiered,SponsorID,RewardDistPeriod,PromoteToTransLevel,RewardLimit,RewardLimitTypeID,TriggerQty,RewardAmountTypeID, " & _
                      "UseSpecialPricing, SPRepeatAtOccur,ApplyToLimit,DoNotItemDistribute,AdvancedLimitID from OfferRewards with (NoLock) where RewardID=" & MyCommon.Extract_Val(RewardID)
  rst = MyCommon.LRT_Select
  For Each row In rst.Rows
    DistPeriod = MyCommon.NZ(row.Item("RewardDistPeriod"), 0)
    LinkID = row.Item("LinkID")
    Tiered = MyCommon.NZ(row.Item("Tiered"), 0)
    SponsorID = MyCommon.NZ(row.Item("SponsorID"), 0)
    PromoteToTransLevel = MyCommon.NZ(row.Item("PromoteToTransLevel"), 0)
    RewardLimit = MyCommon.NZ(row.Item("RewardLimit"), 0)
    RewardLimitTypeID = MyCommon.NZ(row.Item("RewardLimitTypeID"), 2)
    RewardAmountTypeID = MyCommon.NZ(row.Item("RewardAmountTypeID"), 1)
    AdvancedLimitID = MyCommon.NZ(row.Item("AdvancedLimitID"), 0)
    'ExItemLevelDist = MyCommon.NZ(row.Item("ExItemLevelDist"), 0)
    TriggerQty = MyCommon.NZ(row.Item("TriggerQty"), 1)
    ApplyToLimit = MyCommon.NZ(row.Item("ApplyToLimit"), 1)
    UseSpecialPricing = MyCommon.NZ(row.Item("UseSpecialPricing"), 0)
    SPRepeatAtOccur = MyCommon.NZ(row.Item("SPRepeatAtOccur"), 1)
    DoNotItemDistribute = row.Item("DoNotItemDistribute")
  Next
  If (TriggerQty = ApplyToLimit And TriggerQty <> 0) Then
    ValueRadio = 1
  Else
    ValueRadio = 2
  End If
  
        
  MyCommon.QueryStr = "select MessageTypeID,PrintOnBack from PrintedMessages with (NoLock) where MessageID=" & LinkID
  rst = MyCommon.LRT_Select
  For Each row In rst.Rows
    MessageTypeID = MyCommon.NZ(row.Item("MessageTypeID"), "3")
    PrintOnBack = MyCommon.NZ(row.Item("PrintOnBack"), "0")
  Next
  
  If PrintOnBack = 0 Then
    CheckedStatus = False
  Else
    CheckedStatus = True
  End If
  
  
  If Not (bUseTemplateLocks And bDisallowEditPg) Then
        If (Request.Form("save") <> "" And MyCommon.Extract_Val(Request.Form("HdSelectedProductGroup")) <> "0") Then
            MyCommon.QueryStr = "update OfferRewards with (RowLock) set ProductGroupID=0 where RewardID=" & RewardID & ";"
            MyCommon.LRT_Execute()
            MyCommon.QueryStr = "update OfferRewards with (RowLock) set ProductGroupID=" & MyCommon.Extract_Val(Request.Form("HdSelectedProductGroup")) & " where RewardID=" & RewardID & ";"
            MyCommon.LRT_Execute()
        ElseIf (Request.Form("save") <> "" And SelectedProductGroupId = 0) Then
            MyCommon.QueryStr = "update OfferRewards with (RowLock) set ProductGroupID=0 where RewardID=" & RewardID & ";"
            MyCommon.LRT_Execute()
        End If
        
        If (Request.Form("save") <> "" And MyCommon.Extract_Val(Request.Form("HdExcludedProdGroupID")) <> "0" And MyCommon.Extract_Val(Request.Form("HdExcludedProdGroupID")) <> "1") Then
            MyCommon.QueryStr = "update OfferRewards with (RowLock) set ExcludedProdGroupID=0 where RewardID=" & RewardID & ";"
            MyCommon.LRT_Execute()
            MyCommon.QueryStr = "update OfferRewards with (RowLock) set ExcludedProdGroupID=" & MyCommon.Extract_Val(Request.Form("HdExcludedProdGroupID")) & " where RewardID=" & RewardID & ";"
            MyCommon.LRT_Execute()
        ElseIf (Request.Form("save") <> "" And ExcludedProdGroupID = 0) Then
            MyCommon.QueryStr = "update OfferRewards with (RowLock) set ExcludedProdGroupID=0 where RewardID=" & RewardID & ";"
            MyCommon.LRT_Execute()
        ElseIf (GetCgiValue("add") <> "") Then
        
            MyCommon.QueryStr = "select ProductGroupID, ExcludedProdGroupID from OfferRewards with (NoLock) where RewardID=" & RewardID & " and deleted=0;"
            rst = MyCommon.LRT_Select
            
			GName = GetCgiValue("modprodgroupname")

            If (rst.Rows(0).Item("ProductGroupID") = 0) Then 'Create new product group
                 
                MyCommon.QueryStr = "SELECT ProductGroupID FROM ProductGroups with (NoLock) WHERE Name = '" & IIf(GName.Contains("'"), GName.Replace("'", "''"), GName) & "' AND Deleted=0"
                rst1 = MyCommon.LRT_Select
                If (rst1.Rows.Count > 0) Then
                    SelectedProductGroupId = rst1.Rows(0).Item("ProductGroupID")
                Else
                    MyCommon.QueryStr = "dbo.pt_ProductGroups_Insert"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 200).Value = GName
                    GName = MyCommon.Parse_Quotes(Logix.TrimAll(GName))
                    MyCommon.LRTsp.Parameters.Add("@AnyProduct", SqlDbType.Bit).Value = 0
                    MyCommon.LRTsp.Parameters.Add("@AdminID", SqlDbType.Int).Value = AdminUserID
                    MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                    MyCommon.LRTsp.ExecuteNonQuery()
                    SelectedProductGroupId = MyCommon.LRTsp.Parameters("@ProductGroupID").Value
                    Send("<input type=""hidden"" id=""NewCreatedProdGroupID"" name=""NewCreatedProdGroupID"" value=""" & SelectedProductGroupId & """ />")
                    MyCommon.Activity_Log(5, ProductGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.pgroup-create", LanguageID))
                    MyCommon.Close_LRTsp()
                End If
            Else
                SelectedProductGroupId = rst.Rows(0).Item("ProductGroupID")
            End If
        
            If (Trim(GetCgiValue("ExtProductID")) = "") Then
                infoMessage = Copient.PhraseLib.Lookup("pgroup-edit.badid", LanguageID)
            Else
                '4722 Above 
                Dim bGoodItemCode As Boolean = True
                Dim bProductExist As Boolean = True
                Dim bCreateProducts As Boolean = MyCommon.Fetch_SystemOption(150)
                Dim bAddProduct As Boolean = True
                ' desired product add to group   
                ' dbo.pt_ProdGroupItems_Insert  @ExtProductID nvarchar(20), @ProductGroupID bigint, @ProductTypeID int, @Status int OUTPU
                'Send("Inserting product type : " & GetCgiValue("producttype"))
                If (Int(GetCgiValue("producttype")) = 1) Then
                    Integer.TryParse(MyCommon.Fetch_SystemOption(52), IDLength)
                ElseIf (Int(GetCgiValue("producttype")) = 2) Then
                    Integer.TryParse(MyCommon.Fetch_SystemOption(54), IDLength)
                Else
                    IDLength = 0
                End If
                If (IDLength > 0) Then
					If (Int(GetCgiValue("producttype")) = 2) Then
						ExtProductID = MyCommon.Parse_Quotes(Left(Trim(GetCgiValue("ExtProductID")), 120)).PadLeft(IDLength, "0")
					Else
						ExtProductID = MyCommon.Parse_Quotes(Left(Trim(GetCgiValue("ExtProductID")), 26)).PadLeft(IDLength, "0")
					End If
                Else
					If (Int(GetCgiValue("producttype")) = 2) Then
						ExtProductID = MyCommon.Parse_Quotes(Left(Trim(GetCgiValue("ExtProductID")), 120))
					Else
						ExtProductID = MyCommon.Parse_Quotes(Left(Trim(GetCgiValue("ExtProductID")), 26))
					End If                    
                End If
                'Don't change the product description if it is saved as blank
                MyCommon.QueryStr = "Select Description from Products where ExtProductID='" & ExtProductID & "' and ProductTypeID=" & Int(GetCgiValue("producttype")) & ";"
                prodDT = MyCommon.LRT_Select()
                If prodDT.Rows.Count > 0 Then
                    Description = MyCommon.NZ(prodDT.Rows(0).Item("Description"), "")
                Else
                    bProductExist = False
                End If
                If GetCgiValue("productdesc") <> "" Then
                    Description = GetCgiValue("productdesc")
                End If
        
                If (MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(97)) = 1 AndAlso CleanUPC(GetCgiValue("ExtProductID")) = False) Then
                    infoMessage = Copient.PhraseLib.Lookup("pgroup-edit.badid", LanguageID)
                    bGoodItemCode = False
                ElseIf bProductExist = False AndAlso bCreateProducts = False Then
                    infoMessage = Copient.PhraseLib.Lookup("pgroup-edit.productnotexist", LanguageID)
                    bAddProduct = False
                End If
    
                Const Product_NotChanged_Status As Integer = 0
                Const Not_Changed_Status As Integer = 0
                Const Product_Add_Status As Integer = 1
                Const Add_Status As Integer = 1
                Const Product_Update_Status As Integer = 2
                Const Update_Status As Integer = 2
                Dim productOutputStatus As Integer = 0

                bGoodItemCode = True
                If (MyCommon.Extract_Val(GetCgiValue("ExtProductID")) < 1) Or (Int(MyCommon.Extract_Val(GetCgiValue("ExtProductID"))) <> MyCommon.Extract_Val(GetCgiValue("ExtProductID"))) Then
                    infoMessage = Copient.PhraseLib.Lookup("pgroup-edit.badid", LanguageID)
                    bGoodItemCode = False
                ElseIf (MyCommon.Fetch_CM_SystemOption(82) = "1" AndAlso MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM) = True) Then
                    Dim sItemCode As String = GetCgiValue("ExtProductID").ToString
                    Dim productType As Integer = Int(GetCgiValue("producttype"))
                    If (productType = 1) Then
                        If (CheckItemCode(sItemCode, infoMessage) = False) Then
                            bGoodItemCode = False
                        End If
                    End If
                ElseIf (CleanUPC(GetCgiValue("ExtProductID")) = False) Then
                    infoMessage = Copient.PhraseLib.Lookup("pgroup-edit.badid", LanguageID)
                    bGoodItemCode = False
                End If

                If bGoodItemCode = True AndAlso bAddProduct = True Then
                    MyCommon.QueryStr = "dbo.pa_ProdGroupItems_ManualInsert"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@ExtProductID", SqlDbType.NVarChar, 120).Value = ExtProductID
                    MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt).Value = SelectedProductGroupId
                    MyCommon.LRTsp.Parameters.Add("@ProductTypeID", SqlDbType.Int).Value = Int(GetCgiValue("producttype"))
                    MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 200).Value = Description
                    MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
                    MyCommon.LRTsp.Parameters.Add("@ProductStatus", SqlDbType.Int).Direction = ParameterDirection.Output
                    MyCommon.LRTsp.ExecuteNonQuery()
                    outputStatus = MyCommon.LRTsp.Parameters("@Status").Value
                    productOutputStatus = MyCommon.LRTsp.Parameters("@ProductStatus").Value
                End If
                MyCommon.Close_LRTsp()
                MyCommon.QueryStr = "Select PhraseID,Name from ProductTypes where ProductTypeID=" & Int(GetCgiValue("producttype")) & ";"
                Dim productTypeTable As DataTable = MyCommon.LRT_Select()
                Dim typePhrase As Integer = 0
                If (productTypeTable.Rows.Count > 0) Then
                    typePhrase = MyCommon.NZ(productTypeTable.Rows(0).Item("PhraseID"), 0)
                End If
                If (productOutputStatus > Product_NotChanged_Status) Then
                    If (productOutputStatus = Product_Add_Status) Then
                        MyCommon.Activity_Log(5, SelectedProductGroupId, AdminUserID, Copient.PhraseLib.Lookup("term.created", LanguageID) & " " & Copient.PhraseLib.Lookup("term.product", LanguageID).ToLower() & " " & Copient.PhraseLib.Lookup("term.id", LanguageID) & " " & ExtProductID & _
                                              IIf(typePhrase > 0, "(" & Copient.PhraseLib.Lookup(typePhrase, LanguageID) & ")", ""))
                    ElseIf (productOutputStatus = Product_Update_Status) Then
                        MyCommon.Activity_Log(5, SelectedProductGroupId, AdminUserID, Copient.PhraseLib.Lookup("term.updated", LanguageID) & " " & Copient.PhraseLib.Lookup("term.product", LanguageID).ToLower() & " " & Copient.PhraseLib.Lookup("term.id", LanguageID) & " " & ExtProductID & _
                                              IIf(typePhrase > 0, "(" & Copient.PhraseLib.Lookup(typePhrase, LanguageID) & ")", ""))
                    End If
                End If
                If (outputStatus > Not_Changed_Status) Then
                    If (outputStatus = Add_Status) Then
                        MyCommon.Activity_Log(5, SelectedProductGroupId, AdminUserID, Copient.PhraseLib.Lookup("history.pgroup-add", LanguageID) & " " & ExtProductID & IIf(typePhrase > 0, "(" & Copient.PhraseLib.Lookup(typePhrase, LanguageID) & ")", ""))
                    ElseIf (outputStatus = Update_Status) Then
                        'Product was updated to be a manual product entry from a linked product.
                        'MyCommon.Activity_Log(5, ProductGroupID, AdminUserID, "Updated product ID " & ExtProductID)
                    End If
                End If
	
                If (outputStatus <> 0 OrElse productOutputStatus <> 0) Then
                    MyCommon.QueryStr = "update productgroups with (RowLock) set updatelevel=updatelevel+1, LastUpdate=getdate() where ProductGroupID=" & SelectedProductGroupId
                    MyCommon.LRT_Execute()
                End If
                'If infoMessage = "" Then
                '    Response.Status = "301 Moved Permanently"
                '    Response.AddHeader("Location", "pgroup-edit.aspx?ProductGroupID=" & ProductGroupID)
                'End If
            End If
        ElseIf GetCgiValue("modifyprodgroup") <> "" Then
    
            Dim Products As String = GetCgiValue("pasteproducts").Trim
            Dim OperationType As Integer = Int(GetCgiValue("modifyoperation"))
            Dim ProductType As Integer = Int(GetCgiValue("producttype"))

    
            MyCommon.QueryStr = "select ProductGroupID, ExcludedProdGroupID from OfferRewards with (NoLock) where RewardID=" & RewardID & " and deleted=0;"
            rst = MyCommon.LRT_Select
           
		    GName = GetCgiValue("modprodgroupname")
    
            If (rst.Rows(0).Item("ProductGroupID") = 0) Then 'Create new product group
      
                MyCommon.QueryStr = "SELECT ProductGroupID FROM ProductGroups with (NoLock) WHERE Name = '" & IIf(GName.Contains("'"), GName.Replace("'", "''"), GName) & "' AND Deleted=0"
                rst1 = MyCommon.LRT_Select
                If (rst1.Rows.Count > 0) Then
                    SelectedProductGroupId = rst1.Rows(0).Item("ProductGroupID")
                Else
                    MyCommon.QueryStr = "dbo.pt_ProductGroups_Insert"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 200).Value = GName
                    GName = MyCommon.Parse_Quotes(Logix.TrimAll(GName))
                    MyCommon.LRTsp.Parameters.Add("@AnyProduct", SqlDbType.Bit).Value = 0
                    MyCommon.LRTsp.Parameters.Add("@AdminID", SqlDbType.Int).Value = AdminUserID
                    MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                    MyCommon.LRTsp.ExecuteNonQuery()
                    SelectedProductGroupId = MyCommon.LRTsp.Parameters("@ProductGroupID").Value
                    Send("<input type=""hidden"" id=""NewCreatedProdGroupID"" name=""NewCreatedProdGroupID"" value=""" & SelectedProductGroupId & """ />")
                    MyCommon.Activity_Log(5, SelectedProductGroupId, AdminUserID, Copient.PhraseLib.Lookup("history.pgroup-create", LanguageID))
                    MyCommon.Close_LRTsp()
                End If
            Else
                SelectedProductGroupId = rst.Rows(0).Item("ProductGroupID")
            End If
        
    
            If infoMessage = "" Then
        
                If (Not String.IsNullOrEmpty(Products)) Then
                    If MyCommon.Fetch_SystemOption(208) = 1 Then
                        Products = Regex.Replace(Products, "\s", ", ").Replace("-", "")
                    Else
                        Products = Regex.Replace(Products, "\r?\n", ", ")
                    End If
                    tempProducts = CleanProductCodes(Products)
                    tempProductsList = tempProducts.Split(",")
                    Integer.TryParse(MyCommon.Fetch_SystemOption(166), maxLimit)
			 
                    If (tempProductsList.Count > 0 AndAlso maxLimit = 0) Or (tempProductsList.Count > 0 AndAlso tempProductsList.Count <= maxLimit) Then
                        For Each item In tempProductsList
                            item = Trim(item)
                            If (Not String.IsNullOrEmpty(item)) Then
                                If (Not validItemList.Contains(item)) Then
                                    If (IsValidItemCode(SelectedProductGroupId, ProductType, item, OperationType, infoMessage)) Then
                                        validItemList.Add(item)
                                        tempTableInsertStatement.Append(SaveProduct(item, ProductType, OperationType))
                                    Else
                                        invalidItemList.Add(item)
                                    End If
                                End If
                            End If
                        Next
                
                        If (validItemList.Count > 0) Then
                            SaveProductToProductGroup(tempTableInsertStatement.ToString(), SelectedProductGroupId, OperationType, ProductType)
                        End If
                    Else
                        infoMessage = "Invalid" & "~|" & Copient.PhraseLib.Lookup("pgroup-edit.maxlimit", LanguageID) & " : " & maxLimit
                    End If
                End If
            End If
            If invalidItemList.Count > 0 AndAlso (infoMessage = "" OrElse infoMessage Is Nothing) Then
                infoMessage = "There are " & invalidItemList.Count & " invalid items"
            End If
        ElseIf (GetCgiValue("mremove") <> "") Then
            ' desired product remove from group  dbo.pt_GroupMembership_Delete_ByID  @MembershipID bigint
            ' dbo.pt_ProdGroupItems_Delete  @ExtProductID nvarchar(20), @ProductGroupID bigint, @ProductTypeID int, @Status int OUTPUT
            MyCommon.QueryStr = "select ProductGroupID, ExcludedProdGroupID from OfferRewards with (NoLock) where RewardID=" & RewardID & " and deleted=0;"
            rst = MyCommon.LRT_Select
            ' Build product group name 
    
            If (rst.Rows(0).Item("ProductGroupID") = 0) Then
                infoMessage = "Can not remove products. No productgroup associated with the offer"
            Else
                If (GetCgiValue("ExtProductID") <> "") Then
                    If (Int(GetCgiValue("producttype")) = 1) Then
                        Integer.TryParse(MyCommon.Fetch_SystemOption(52), IDLength)
                    ElseIf (Int(GetCgiValue("producttype")) = 2) Then
                        Integer.TryParse(MyCommon.Fetch_SystemOption(54), IDLength)
                    Else
                        IDLength = 0
                    End If
                    If (IDLength > 0) Then
						If (Int(GetCgiValue("producttype")) = 2) Then
							ExtProductID = MyCommon.Parse_Quotes(Left(Trim(GetCgiValue("ExtProductID")), 120)).PadLeft(IDLength, "0")
						Else
							ExtProductID = MyCommon.Parse_Quotes(Left(Trim(GetCgiValue("ExtProductID")), 26)).PadLeft(IDLength, "0")
						End If
                    Else
						If (Int(GetCgiValue("producttype")) = 2) Then
							ExtProductID = MyCommon.Parse_Quotes(Left(Trim(GetCgiValue("ExtProductID")), 120))
						Else
							ExtProductID = MyCommon.Parse_Quotes(Left(Trim(GetCgiValue("ExtProductID")), 26))
						End If
                    End If
      
                    ' check if the product is linked to the group, if so then exclude the product from the group.
                    MyCommon.QueryStr = "select PROD.ProductID, PGI.ExtHierarchyID from Products as PROD with (NoLock) " & _
                                        "inner join ProdGroupItems as PGI with (NoLock) on PGI.ProductID = PROD.ProductID " & _
                                        "where PGI.Deleted=0 and PGI.ProductGroupID=" & SelectedProductGroupId & " and IsNull(PGI.ExtHierarchyID, '') <> ''" & _
                                        "   and PROD.ExtProductID='" & MyCommon.Parse_Quotes(ExtProductID) & "' and PROD.ProductTypeID=" & Int(GetCgiValue("producttype"))
                    rst = MyCommon.LRT_Select
                    If rst.Rows.Count > 0 Then
                        If rst.Rows(0).Item("ProductID") > 0 AndAlso MyCommon.NZ(rst.Rows(0).Item("ExtHierarchyID"), "") <> "" Then
                            MyCommon.QueryStr = "insert into ProdGroupHierarchyExclusions (ExtHierarchyID,ProductGroupID, HierarchyLevel, LevelID) " & _
                                                "      values ('" & MyCommon.Parse_Quotes(MyCommon.NZ(rst.Rows(0).Item("ExtHierarchyID"), "")) & "', " & SelectedProductGroupId & ", 2, '" & rst.Rows(0).Item("ProductID") & "')"
                            MyCommon.LRT_Execute()
                        End If
                    End If
      
                    MyCommon.Open_LogixRT()
                    MyCommon.QueryStr = "dbo.[pt_ProdGroupItems_DeleteItem]"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@ExtProductID", SqlDbType.NVarChar, 120).Value = ExtProductID
                    MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt).Value = SelectedProductGroupId
                    MyCommon.LRTsp.Parameters.Add("@ProductTypeID", SqlDbType.Int).Value = Int(GetCgiValue("producttype"))
                    MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
                    MyCommon.LRTsp.ExecuteNonQuery()
                    outputStatus = MyCommon.LRTsp.Parameters("@Status").Value
                    MyCommon.Close_LRTsp()
                    MyCommon.Activity_Log(5, SelectedProductGroupId, AdminUserID, Copient.PhraseLib.Lookup("history.pgroup-remove", LanguageID) & " " & ExtProductID)
                    If (outputStatus <> 0) Then
                        infoMessage = Copient.PhraseLib.Lookup("pgroup-edit.notmember", LanguageID)
                    Else
                        MyCommon.QueryStr = "update productgroups with (RowLock) set updatelevel=updatelevel+1, LastUpdate=getdate() where ProductGroupID=" & SelectedProductGroupId
                        MyCommon.LRT_Execute()
                    End If
                Else
                    infoMessage = Copient.PhraseLib.Lookup("pgroup-edit.badid", LanguageID)
                End If
            End If
	
    
        ElseIf (GetCgiValue("remove") <> "" AndAlso GetCgiValue("PKID") <> "") Then
            MyCommon.Open_LogixRT()
            For i = 0 To Request.Form.GetValues("PKID").GetUpperBound(0)
                MyCommon.QueryStr = "select P.ExtProductID from Products as P with (NoLock) Inner Join ProdGroupItems as PGI " & _
                                    "with (NoLock) on P.ProductID=PGI.ProductID where PGI.PKID=" & Request.Form.GetValues("PKID")(i)
                rst = MyCommon.LRT_Select()
                If rst.Rows.Count > 0 Then
                    upc = rst.Rows(0).Item("ExtProductID")
                End If
                rst = Nothing
                MyCommon.QueryStr = "dbo.pt_ProdGroupItems_Delete_ByID"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@PKID", SqlDbType.BigInt).Value = Request.Form.GetValues("PKID")(i)
                MyCommon.LRTsp.ExecuteNonQuery()
                MyCommon.Close_LRTsp()
                MyCommon.Activity_Log(5, SelectedProductGroupId, AdminUserID, Copient.PhraseLib.Lookup("history.pgroup-remove", LanguageID) & " " & upc)
            Next
            MyCommon.QueryStr = "update productgroups with (RowLock) set updatelevel=updatelevel+1, LastUpdate=getdate() where ProductGroupID=" & SelectedProductGroupId
            MyCommon.LRT_Execute()
        End If
    End If
  
    If SelectedProductGroupId > 0 Then
        MyCommon.QueryStr = "select count(*) as GCount from ProdGroupItems with (NoLock) where ProductGroupID = " & SelectedProductGroupId & " And Deleted = 0"
        rst = MyCommon.LRT_Select()
        For Each row In rst.Rows
            GroupSize = row.Item("GCount")
        Next
        MyCommon.QueryStr = "select count(*) as PCount from ProdGroupItems PGI with (NoLock) inner join products PRD on PGI.productid = PRD.productid " & _
                      "where PGI.ProductGroupID = " & SelectedProductGroupId & " And PGI.Deleted = 0 And (PRD.Description IS NULL OR  PRD.Description = '')"
        rst = MyCommon.LRT_Select()
        If rst.Rows.Count > 0 Then
            ProductsWithoutDesc = rst.Rows(0).Item("PCount")
        End If
    Else
        MyCommon.QueryStr = "select count(*) as GCount from ProdGroupItems PGI with (NoLock) inner join OfferRewards ORe on ORe.ProductGroupID = PGI.ProductGroupID " & _
                            "where ORe.RewardID = " & RewardID & " And ORe.Deleted = 0 And PGI.Deleted = 0"
        rst = MyCommon.LRT_Select()
        For Each row In rst.Rows
            GroupSize = row.Item("GCount")
        Next
        
        MyCommon.QueryStr = "select count(*) as PCount from ProdGroupItems PGI with (NoLock) inner join products PRD on PGI.productid = PRD.productid " & _
                            "inner join OfferRewards ORe with (NoLock) on ORe.ProductGroupID = PGI.ProductGroupID " & _
                            "where ORe.RewardID = " & RewardID & " And PGI.Deleted = 0 And (PRD.Description IS NULL OR PRD.Description = '') And ORe.Deleted = 0"
        rst = MyCommon.LRT_Select()
        If rst.Rows.Count > 0 Then
            ProductsWithoutDesc = rst.Rows(0).Item("PCount")
        End If
    End If
		
    ShowAllItems = (GetCgiValue("showall") = "true")
    Dim bBlankDescProd As Boolean = (MyCommon.Fetch_SystemOption(206) = "1")
    Dim sBlankDescOrderByStr As String = " order by case when Description is null or Description = '' then nullif(Description, '') else CAST(ExtProductID as bigint) end, CAST(ExtProductID as bigint), ExtProductID DESC;"
    If SelectedProductGroupId > 0 Then
        MyCommon.QueryStr = "select" & If(ShowAllItems, "", " top 100") & " GM.ProductID, PKID, CID.ProductTypeID, ExtProductID, Description, PT.Name as ProductType, PT.PhraseID " & _
                          "from Products as CID with (NoLock) " & _
                          "inner join ProdGroupItems as GM with (NoLock) on CID.ProductID=GM.ProductID " & _
                          "left join ProductTypes as PT with (NoLock) on PT.ProductTypeID=CID.ProductTypeID " & _
                          "where GM.ProductGroupID=" & SelectedProductGroupId & " and GM.Deleted=0 and IsNull(GM.ExtHierarchyID, '')='' " & _
                          "and IsNull(GM.ExtNodeID, '')='' " & If(bBlankDescProd, sBlankDescOrderByStr, " order by ExtProductID;")
    Else
        MyCommon.QueryStr = "select" & If(ShowAllItems, "", " top 100") & " GM.ProductID, PKID, CID.ProductTypeID, ExtProductID, Description, PT.Name as ProductType, PT.PhraseID " & _
                          "from Products as CID with (NoLock) " & _
                          "inner join ProdGroupItems as GM with (NoLock) on CID.ProductID=GM.ProductID " & _
                          "left join ProductTypes as PT with (NoLock) on PT.ProductTypeID=CID.ProductTypeID " & _
                          "inner join OfferRewards as ORe with (NoLock) on ORe.ProductGroupID = gm.ProductGroupID " & _
                          "where ORe.RewardID = " & RewardID & " and GM.Deleted=0 and ORe.Deleted =0 and IsNull(GM.ExtHierarchyID, '')='' " & _
                          "and IsNull(GM.ExtNodeID, '')='' " & If(bBlankDescProd, sBlankDescOrderByStr, " order by ExtProductID;")
    
    End If
    
    
    
    rstItems = MyCommon.LRT_Select()
    ListBoxSize = rstItems.Rows.Count

  
        If (Request.Form("save") <> "" Or _
        Request.Form("pgroup-add1") <> "" Or _
        Request.Form("pgroup-rem1") <> "" Or _
        Request.Form("pgroup-add2") <> "" Or _
        Request.Form("pgroup-rem2") <> "") Then
    
            Dim TemplateString As String = ""
            If (Request.QueryString("IsTemplate") <> "") Then
                TemplateString = Request.QueryString("IsTemplate")
            Else
                TemplateString = Request.Form("IsTemplate")
            End If
            If (TemplateString = "IsTemplate") Then
                ' time to update the status bits for the templates
                Dim form_Disallow_Edit As Integer = 0
                Dim iDisallowEditPg As Integer = 0
                Dim iDisallowEditSpon As Integer = 0
                Dim iDisallowEditMsg As Integer = 0
                Dim iDisallowEditDist As Integer = 0
                Dim iDisallowEditLimit As Integer = 0
                Dim iDisallowEditAdv As Integer = 0
      
                Disallow_Edit = False
                bDisallowEditPg = False
                bDisallowEditSpon = False
                bDisallowEditMsg = False
                bDisallowEditDist = False
                bDisallowEditLimit = False
                bDisallowEditAdv = False
      
                If (Request.Form("Disallow_Edit") = "on") Then
                    form_Disallow_Edit = 1
                    Disallow_Edit = True
                End If

                If (Request.Form("DisallowEditPg") = "on") Then
                    iDisallowEditPg = 1
                    bDisallowEditPg = True
                End If

                If (Request.Form("DisallowEditSpon") = "on") Then
                    iDisallowEditSpon = 1
                    bDisallowEditSpon = True
                End If

                If (Request.Form("DisallowEditMsg") = "on") Then
                    iDisallowEditMsg = 1
                    bDisallowEditMsg = True
                End If

                If (Request.Form("DisallowEditDist") = "on") Then
                    iDisallowEditDist = 1
                    bDisallowEditDist = True
                End If

                If (Request.Form("DisallowEditLimit") = "on") Then
                    iDisallowEditLimit = 1
                    bDisallowEditLimit = True
                End If

                If (Request.Form("DisallowEditAdv") = "on") Then
                    iDisallowEditAdv = 1
                    bDisallowEditAdv = True
                End If
                MyCommon.QueryStr = "update OfferRewards with (RowLock) set Disallow_Edit=" & form_Disallow_Edit & _
                  ",DisallowEdit1=" & iDisallowEditPg & _
                  ",DisallowEdit3=" & iDisallowEditSpon & _
                  ",DisallowEdit4=" & iDisallowEditMsg & _
                  ",DisallowEdit6=" & iDisallowEditDist & _
                  ",DisallowEdit7=" & iDisallowEditLimit & _
                  ",DisallowEdit9=" & iDisallowEditAdv & _
                  " where RewardID=" & RewardID
                MyCommon.LRT_Execute()
            End If

            If Not (bUseTemplateLocks And bDisallowEditSpon) Then
                If (Request.Form("sponsor") <> "") Then
                    SponsorID = MyCommon.Extract_Val(Request.Form("sponsor"))
                    MyCommon.QueryStr = "update OfferRewards with (RowLock) set SponsorID=" & SponsorID & " where RewardID=" & RewardID
                    MyCommon.LRT_Execute()
                End If
            End If

            If Not (bUseTemplateLocks And bDisallowEditAdv) Then
                If (Request.Form("promote") = "on") Then
                    MyCommon.QueryStr = "update OfferRewards with (RowLock) set PromoteToTransLevel=1 where RewardID=" & RewardID
                    MyCommon.LRT_Execute()
                    PromoteToTransLevel = True
                Else
                    MyCommon.QueryStr = "update OfferRewards with (RowLock) set PromoteToTransLevel=0 where RewardID=" & RewardID
                    MyCommon.LRT_Execute()
                    PromoteToTransLevel = False
                End If
            End If

   
            If Not (bUseTemplateLocks And bDisallowEditLimit) Then
                If (Request.Form("selectadv") <> "") Then
                    AdvancedLimitID = Request.Form("selectadv")
                    If AdvancedLimitID > 0 Then
                        MyCommon.QueryStr = "select AL.PromoVarID,AL.LimitTypeID, AL.LimitValue, AL.LimitPeriod " & _
                                            "from CM_AdvancedLimits as AL with (NoLock) where Deleted=0 and LimitID='" & AdvancedLimitID & "';"
                        rst = MyCommon.LRT_Select
                        If (rst.Rows.Count > 0) Then
                            VarID = MyCommon.NZ(rst.Rows(0).Item("PromoVarID"), 0)
                            RewardLimitTypeID = MyCommon.NZ(rst.Rows(0).Item("LimitTypeID"), 5)
                            RewardLimit = MyCommon.NZ(rst.Rows(0).Item("LimitValue"), 0)
                            DistPeriod = MyCommon.NZ(rst.Rows(0).Item("LimitPeriod"), 0)
                            MyCommon.QueryStr = "update OfferRewards with (RowLock) set AdvancedLimitID=" & AdvancedLimitID & _
                                                ",RewardDistLimitVarID=" & VarID & _
                                                ",RewardLimitTypeID=" & RewardLimitTypeID & _
                                                ",RewardLimit=" & RewardLimit & _
                                                ",RewardDistPeriod=" & DistPeriod & _
                                                " where RewardID=" & RewardID & ";"
                            MyCommon.LRT_Execute()
                        Else
                            MyCommon.QueryStr = "update OfferRewards with (RowLock) set AdvancedLimitID=0" & _
                                                ",RewardDistPeriod=0" & _
                                                ",RewardLimit=0.0" & _
                                                " where RewardID=" & RewardID & ";"
                            MyCommon.LRT_Execute()
                        End If
                    Else
                        MyCommon.QueryStr = "select PromoVarID, VarTypeID, LinkID " & _
                                            "from PromoVariables with (NoLock) where Deleted=0 and VarTypeID=4 and LinkID=" & RewardID & ";"
                        rst = MyCommon.LXS_Select
                        If (rst.Rows.Count > 0) Then
                            VarID = MyCommon.NZ(rst.Rows(0).Item("PromoVarID"), 0)
                        Else
                            VarID = 0
                        End If
                        MyCommon.QueryStr = "update OfferRewards with (RowLock) set AdvancedLimitID=" & AdvancedLimitID & _
                                            ",RewardDistLimitVarID=" & VarID & _
                                            ",RewardDistPeriod=0" & _
                                            " where RewardID=" & RewardID & ";"
                        MyCommon.LRT_Execute()
                    End If
                End If
                If AdvancedLimitID = 0 Then
                    If (Request.Form("RewardLimitTypeID") <> "") Then
                        RewardLimitTypeID = MyCommon.Extract_Val(Request.Form("RewardLimitTypeID"))
                        MyCommon.QueryStr = "update OfferRewards with (RowLock) set RewardLimitTypeID=" & RewardLimitTypeID & " where RewardID=" & RewardID
                        MyCommon.LRT_Execute()
                    End If
                    If (Request.Form("limitvalue") <> "") Then
                        RewardLimit = MyCommon.Extract_Val(Request.Form("limitvalue"))
                        MyCommon.QueryStr = "update OfferRewards with (RowLock) set RewardLimit=" & RewardLimit & " where RewardID=" & RewardID
                        MyCommon.LRT_Execute()
                    End If
                    If (Request.Form("form_DistPeriod") <> "") Then
                        DistPeriod = Int(MyCommon.Extract_Val(Request.Form("form_DistPeriod")))
                        If DistPeriod = 0 Then
                            MyCommon.QueryStr = "update OfferRewards with (RowLock) set RewardDistPeriod =0 where RewardID=" & RewardID
                        ElseIf DistPeriod = -1 Then
                            MyCommon.QueryStr = "update OfferRewards with (RowLock) set RewardDistPeriod =-1 where RewardID=" & RewardID
                        Else
                            MyCommon.QueryStr = "update OfferRewards with (RowLock) set RewardDistPeriod =" & DistPeriod & " where RewardID=" & RewardID
                        End If
                        MyCommon.LRT_Execute()
                        ' someone saves - let's do the special case and set a promo variable if the
                        ' distribution's greater than zero and the promo variable doesn't already exist
                        If DistPeriod <> 0 Then
                            MyCommon.QueryStr = "select RewardDistLimitVarID from OfferRewards with (NoLock) where RewardID=" & RewardID
                            rst = MyCommon.LRT_Select
                            For Each row In rst.Rows
                                If (MyCommon.NZ(row.Item("RewardDistLimitVarID"), 0) = 0) Then
                                    'dbo.pa_DistributionVar_Create @OfferID bigint, @VarID bigint OUTPUT
                                    MyCommon.QueryStr = "dbo.pc_RewardLimitVar_Create"
                                    MyCommon.Open_LXSsp()
                                    MyCommon.LXSsp.Parameters.Add("@RewardID", SqlDbType.BigInt).Value = RewardID
                                    MyCommon.LXSsp.Parameters.Add("@VarID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                                    MyCommon.LXSsp.ExecuteNonQuery()
                                    VarID = MyCommon.LXSsp.Parameters("@VarID").Value
                                    MyCommon.Close_LXSsp()
                                    MyCommon.QueryStr = "update OfferRewards with (RowLock) set RewardDistLimitVarID=" & VarID & " where RewardID=" & RewardID
                                    MyCommon.LRT_Execute()
                                End If
                            Next
                        End If
                    End If
                End If
            End If
    
    
            If Not (bUseTemplateLocks And bDisallowEditMsg) Then
                If (Request.Form("type") <> "") Then
                    PrintReceipt = MyCommon.Extract_Val(Request.Form("PrintOnBackReceipt"))
                    MessageTypeID = MyCommon.Extract_Val(Request.Form("type"))
                    MyCommon.QueryStr = "update PrintedMessages with (RowLock) set MessageTypeID=" & MessageTypeID & ", PrintOnBack=" & PrintReceipt & "where MessageID=" & LinkID
                    MyCommon.LRT_Execute()
        
                    MyCommon.QueryStr = "update CM_ST_PrintedMessages with (RowLock) set MessageTypeID=" & MessageTypeID & ", PrintOnBack=" & PrintReceipt & "where MessageID=" & LinkID
                    MyCommon.LRT_Execute()
                End If

                ' ok here we need to handle the tiering stuffs
                If (Tiered = 0) Then
                    MyCommon.QueryStr = "delete from PrintedMessageTiers with (RowLock) where MessageID=" & LinkID
                    MyCommon.LRT_Execute()
                    MyCommon.QueryStr = "dbo.pt_PrintedMsgTiers_Update"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@MessageID", SqlDbType.BigInt).Value = LinkID
                    MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int).Value = 0
                    MyCommon.LRTsp.Parameters.Add("@BodyText", SqlDbType.NVarChar, 4000).Value = Request.Form("tier0")
                    MyCommon.LRTsp.ExecuteNonQuery()
                    MyCommon.Close_LRTsp()
                Else
                    MyCommon.QueryStr = "delete from PrintedMessageTiers with (RowLock) where MessageID=" & LinkID
                    MyCommon.LRT_Execute()
                    For x = 1 To NumTiers
                        MyCommon.QueryStr = "dbo.pt_PrintedMsgTiers_Update"
                        MyCommon.Open_LRTsp()
                        MyCommon.LRTsp.Parameters.Add("@MessageID", SqlDbType.BigInt).Value = LinkID
                        MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int).Value = x
                        MyCommon.LRTsp.Parameters.Add("@BodyText", SqlDbType.NVarChar, 4000).Value = Request.Form("tier" & x)
                        MyCommon.LRTsp.ExecuteNonQuery()
                        MyCommon.Close_LRTsp()
                    Next
                End If
            End If
    
            If Not (bUseTemplateLocks And bDisallowEditDist) Then
                If (Request.Form("trigger") <> "") Then
                    If (Request.Form("trigger") = "1") Then
                        ' set  TriggerQty=Xbox
                        TriggerQty = MyCommon.Extract_Val(Request.Form("Xbox"))
                        If (TriggerQty = 0) Then
                            TriggerQty = 1
                        End If
                        MyCommon.QueryStr = "update OfferRewards with (RowLock) set TriggerQty=" & TriggerQty & "," & _
                        "ApplyToLimit=" & TriggerQty & " where RewardID=" & RewardID
                        MyCommon.LRT_Execute()
                        ValueRadio = 1
                    ElseIf (Request.Form("trigger") = "2") Then
                        TriggerQty = Int(MyCommon.Extract_Val(Request.Form("Xbox2"))) + Int(MyCommon.Extract_Val(Request.Form("Ybox2")))
                        ApplyToLimit = MyCommon.Extract_Val(Request.Form("Ybox2"))
                        MyCommon.QueryStr = "update OfferRewards with (RowLock) set TriggerQty=" & TriggerQty & "," & _
                        "ApplyToLimit=" & ApplyToLimit & " where RewardID=" & RewardID
                        MyCommon.LRT_Execute()
                        ValueRadio = 2
                        'If (TriggerQty = ApplyToLimit) Then ValueRadio = 1
                    End If
                Else
                    MyCommon.QueryStr = "update OfferRewards with (RowLock) set TriggerQty=0," & _
                    "ApplyToLimit=1 where RewardID=" & RewardID
                    MyCommon.LRT_Execute()
                End If
            End If

            MyCommon.QueryStr = "update OfferRewards with (RowLock) set TCRMAStatusFlag=2,CMOAStatusFlag=2 where RewardID=" & RewardID
            MyCommon.LRT_Execute()
            MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where OfferID=" & OfferID
            MyCommon.LRT_Execute()
            MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.rew-pmsg", LanguageID))
    
        End If
  
        If (Request.Form("pgroup-add1") <> "" Or Request.Form("pgroup-rem1") <> "" Or Request.Form("pgroup-add2") <> "" Or Request.Form("pgroup-rem2") <> "") Then
            MyCommon.QueryStr = "update OfferRewards with (RowLock) set TCRMAStatusFlag=3,CMOAStatusFlag=2 where RewardID=" & RewardID
            MyCommon.LRT_Execute()
            MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where OfferID=" & OfferID
            MyCommon.LRT_Execute()
        End If
  
        Send("<script type=""text/javascript"">")
    Send("function ChangeParentDocument() { ")
    Send("  if (opener != null) {")
    Send("    var newlocation = 'offer-rew.aspx?OfferID=" & OfferID & "'; ")
    Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
    Send("    opener.location = 'offer-rew.aspx?OfferID=" & OfferID & "'; ")
    Send("    } ")
    Send("    } ")
    Send("    } ")
        Send("</script>")
  
        MyCommon.QueryStr = "Select MT.MarkupID, Tag, Description, PhraseID, NumParams, " & _
                            "Param1Name, Param1PhraseID, Param2Name, Param2PhraseID," & _
                            "Param3Name, Param3PhraseID, Param4Name, Param4PhraseID, DisplayOrder, CentralRendered, ButtonText " & _
                            "from MarkupTags as MT with (NoLock) " & _
                            "left join MarkupTagUsage as MTU with (NoLock) on MT.MarkupID=MTU.MarkupID " & _
                            "where MTU.RewardTypeID=1"
        rst = MyCommon.LRT_Select
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

var timer;
var isFireFox = (navigator.appName.indexOf('Mozilla') !=-1) ? true : false;
function xmlPostTimer(strURL,mode)
{
  clearTimeout(timer);
  timer=setTimeout("xmlhttpPostNew('" + strURL + "','" + mode + "')", 250);
}

function xmlhttpPostNew(strURL,mode) {
  var xmlHttpReq = false;
  var self = this;
  document.getElementById("searchLoadDiv").innerHTML = '<% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID)) %>';
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
      updatepagePost(self.xmlHttpReq.responseText);
    }
  }
  self.xmlHttpReq.send(qryStr);
  setSearchTypeRadiobutton();
}


function getproductquery(mode) {
  
  var radioString;
  if(document.getElementById('functionradio2').checked) {
    radioString = 'functionradio2';
  }
  else {
    radioString = 'functionradio1';
  }
  var selected = document.getElementById('pgroup-select');
  var selectedGroup = 0;
  if(selected.options[0] != null){
    selectedGroup = selected.options[0].value;
  }
  var excluded = document.getElementById('pgroup-exclude');
  var excludedGroup = 0;
  if(excluded.options[0] != null){
    excludedGroup = excluded.options[0].value;
  }
  return "Mode=" + mode + "&ProductSearch=" + document.getElementById('functioninputSearch').value + "&OfferID=" + document.getElementById('OfferID').value + "&SelectedGroup=" + selectedGroup + "&ExcludedGroup=" + excludedGroup + "&SearchRadio=" + radioString;
}

function setSearchTypeRadiobutton()
{
 if(document.getElementById('functionradio1').checked) 
  {
   document.getElementById('hdSearchType').value = "1"; 
  }
  if(document.getElementById('functionradio2').checked ) 
  {
   document.getElementById('hdSearchType').value = "2"; 
  }
}


function updatepagePost(str) {
  if(str.length > 0){
    if(!isFireFox){
      document.getElementById("pgavaildiv").innerHTML = '<select class="longer" id="pgroup-avail" name="pgroup-avail" size="6"<% sendb(disabledattribute) %>>' + str + '</select>';
    }
    else{
      document.getElementById("pgavaildiv").innerHTML = str;
    }
    document.getElementById("searchLoadDiv").innerHTML = '&nbsp;';
    if (document.getElementById("pgroup-avail").options.length > 0) {
      document.getElementById("pgroup-avail").options[0].selected = true;
    }
  }
  else if(str.length == 0){
    if(!isFireFox){
      document.getElementById("pgavaildiv").innerHTML = '<select class="longer" id="pgroup-avail" name="pgroup-avail" size="6"<% sendb(disabledattribute) %>>&nbsp;</select>';
    }
    else{
      document.getElementById("pgavaildiv").innerHTML = '&nbsp;';
    }
    document.getElementById("searchLoadDiv").innerHTML = '&nbsp;';
  }
}
function edShowButton(button) {
	if (button.access) {
		var accesskey = ' accesskey = "' + button.access + '"'
	}
	else {
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
  // reinitialize myField variable
  if (document.getElementById(textAreaName) != null) {
    myField = document.getElementById(textAreaName);
  }

	//IE support
	if (document.selection) {
		myField.focus();
		sel = document.selection.createRange();
		if (sel.text.length > 0) {
			sel.text = edButtons[i].tagStart + sel.text + edButtons[i].tagEnd;
		}
		else {
			if (!edCheckOpenTags(i) || edButtons[i].tagEnd == '') {
				sel.text = edButtons[i].tagStart;
				edAddTag(i);
			}
			else {
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
			}	else {
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
		}	else {
			myField.value += edButtons[i].tagEnd;
			edRemoveTag(i);
		}
		myField.focus();
	}
}

function isValidID() {
        var retVal = true;
        var elemNumericOnly = document.getElementById("NumericOnly");
        var elemID = document.getElementById("productid");
		var elSel = document.getElementById('selected');
        var selectList = "";
		
        if((elemID != null) && (elemID.value.length == 0)) {
            retVal = false;
            alert('<%Sendb(Copient.PhraseLib.Lookup("term.invalid", LanguageID) + " " + Copient.PhraseLib.Lookup("term.productid", LanguageID)) %>');
        }
        if ((elemNumericOnly != null) && (elemNumericOnly.value != "")) {
           if ((elemID != null) && (isNaN(elemID.value))) {
              retVal = false;
              alert('<%Sendb(Copient.PhraseLib.Lookup("product.mustbenumeric", LanguageID)) %>');
           }
        }
		// assemble the list of values from the selected box
        for (i = elSel.length - 1; i>=0; i--) {
          if(elSel.options[i].value != ""){
          if(selectList != "") { selectList = selectList + ","; }
            selectList = selectList + elSel.options[i].value;
          }
		}  
		if (selectList == 1 && retVal == true){
		  retVal = false;
          alert('Product Group should not be "Any Product"');		
		}
  	
        return retVal;
    }

function ProductAddSelection() {

  var elemaddsingleprod = document.getElementById("addsingleproduct");
  var elemaddprodlist = document.getElementById("addproductlist");
  var elemradioprodaddselector1 = document.getElementById("prodaddselector1");
  var elemradioprodaddselector2 = document.getElementById("prodaddselector2");
  
  if (elemradioprodaddselector1.checked) {
      //alert('checked'); 
      elemaddprodlist.style.display = "none";
      elemaddsingleprod.style.display = "block";
    } else {
      //alert('unchecked');
      elemaddprodlist.style.display = "block";
      elemaddsingleprod.style.display = "none";
    }

}

function ProductGroupTypeSelection() {
  
  var elemexistingpgroup = document.getElementById("groups");
  var elemaddpeoducts = document.getElementById("directprodaddselector");
  var elemradiopgselect = document.getElementById("pgselectortype1");
  var elemradioaddprod = document.getElementById("pgselectortype2");
  
  if (elemradiopgselect.checked) {
      //alert('checked'); 
      elemaddpeoducts.style.display = "none";
      elemexistingpgroup.style.display = "block";
    } else {
      //alert('unchecked');
      elemaddpeoducts.style.display = "block";
      elemexistingpgroup.style.display = "none";
    }
}

function edInsertContent(myFieldName, myValue) {
  // reinitialize myField variable
  var myField = document.getElementById(myFieldName);
  if (document.getElementById(textAreaName) != null) {
    myField = document.getElementById(textAreaName);
  }
	//IE support
	if (document.selection) {
		myField.focus();
		sel = document.selection.createRange();
		sel.text = myValue;
		myField.focus();
	}
	//MOZILLA/NETSCAPE support
	else if (myField.selectionStart || myField.selectionStart == '0') {
		var startPos = myField.selectionStart;
		var endPos = myField.selectionEnd;
		var scrollTop = myField.scrollTop;
		myField.value = myField.value.substring(0, startPos)
		              + myValue 
                      + myField.value.substring(endPos, myField.value.length);
		myField.focus();
		myField.selectionStart = startPos + myValue.length;
		myField.selectionEnd = startPos + myValue.length;
		myField.scrollTop = scrollTop;
	} else {
		myField.value += myValue;
		myField.focus();
	}
}

// ~~~~~~~~~~ DYNAMICALLY-GENERATED TAG INSERT FUNCTIONS BEGIN HERE ~~~~~~~~~~
<%
  If MyCommon.IsEngineInstalled(8) Then
    ExcentusInstalled = True
  End If
  MyCommon.QueryStr = "select distinct MT.MarkupID, MT.Tag, MT.Description, MT.PhraseID, MT.NumParams, " & _
                      "MT.Param1Name, MT.Param1PhraseID, MT.Param2Name, MT.Param2PhraseID, " & _
                      "MT.Param3Name, MT.Param3PhraseID, MT.Param4Name, MT.Param4PhraseID, " & _
                      "MT.DisplayOrder, MT.CentralRendered, MT.ButtonText, " & _
                      "MTU.RewardTypeID, MTU.EngineID from MarkupTags as MT with (NoLock) " & _
                      "inner join MarkupTagUsage as MTU with (NoLock) on MT.MarkupID=MTU.MarkupID " & _
                      "where MTU.RewardTypeID=3"
  If ExcentusInstalled Then
    MyCommon.QueryStr &= " and MTU.EngineID in (" & OfferEngineID & ", 8) order by MT.DisplayOrder;"
  Else
    MyCommon.QueryStr &= " and MTU.EngineID=" & OfferEngineID & " order by MT.DisplayOrder;"
  End If
  rst = MyCommon.LRT_Select
  Dim funcname As String
  For Each row In rst.Rows
    funcname = row.Item("ButtonText")
    funcname = funcname.Replace("$", "Amt")
    funcname = funcname.Replace("#", "Dol")
    funcname = funcname.Replace("/", "Off")
    If (row.Item("NumParams") = 0) Then
      Send("function edInsert" & (StrConv(funcname, VbStrConv.ProperCase)) & "(myField) {")
      Send("  myValue = '|" & row.Item("ButtonText") & "|';")
      Send("  edInsertContent(myField, myValue);")
      Send("}")
    Else
      If (row.Item("ButtonText") = "UPCA") or (row.Item("ButtonText") = "UPCB") or (row.Item("ButtonText") = "EAN13") or (row.Item("ButtonText") = "CODE39") Then
        Send("function edInsert" & (StrConv(funcname, VbStrConv.ProperCase)) & "(myField) {")
        Send("  var myValue = prompt('" & Copient.PhraseLib.Lookup(row.Item("Param1PhraseID"), LanguageID) & "', '');")
        Send("  if (myValue) {")
        Send("    myValue = '|" & row.Item("ButtonText") & "[' + myValue + ']|';")
        Send("    edInsertContent(myField, myValue);")
        Send("  }")
        Send("}")
      Else
        Send("function edInsert" & (StrConv(funcname, VbStrConv.ProperCase)) & "(myField, myValue) {")
        If (UCase(Left(funcname, 4)) = "LIFE") then
          Send("  var n = document.getElementById(""functionselect4"").value;")
        ElseIf (UCase(Left(funcname,2)) = "SV") Then
          Send("  var n = document.getElementById(""functionselect3"").value;")
        ElseIf (UCase(Right(funcname, 3)) = "AMT") then
          Send("  var n = document.getElementById(""functionselect2"").value;")
        Else
          Send("  var n = document.getElementById(""functionselect"").value;")
        End If
        Send("  var myValue = n;")
        Send("  if (myValue) {")
        Send("    myValue = '|" & row.Item("ButtonText") & "[' + myValue + ']|';")
        Send("    edInsertContent(myField, myValue);")
        Send("  }")
        Send("}")
      End If
    End If
  Next
%>
// ~~~~~~~~~~ DYNAMICALLY-GENERATED TAG INSERT FUNCTIONS END HERE ~~~~~~~~~~
</script>

<script type="text/javascript">
var textAreaName = "tier0"
// This is the javascript array holding the function list
// The PrintJavascriptArray ASP function can be used to print this array.
<%
  MyCommon.QueryStr = "Select RewardID,OfferID,RewardTypeID,Deleted from OfferRewards with (NoLock) where RewardTypeID=1 and Deleted=0 order by OfferID"
  rst = MyCommon.LRT_Select
  If (rst.rows.count>0)
    Sendb("var functionlist = Array(")
    For Each row In rst.Rows
      Sendb("""" & row.item("OfferID") & """,")
    Next
    Send(""""");")
    Sendb("var vallist = Array(")
    For Each row In rst.Rows
      Sendb("""" & row.item("RewardID") & """,")
    Next
    Send(""""");")
  End If
  
  MyCommon.QueryStr = "select distinct PP.PromoVarID, PP.ProgramName from OfferRewards ORwrd with (NoLock) " & _
                      "inner join RewardPoints RP with (NoLock) on ORwrd.LinkID = RP.RewardPointsID " & _
                      "inner join PointsPrograms PP with (NoLock) on RP.ProgramID = PP.ProgramID " & _
                      "where (ORwrd.RewardTypeID=2 or ORwrd.RewardTypeID=13) and ORwrd.Deleted=0 and PP.Deleted=0 " & _
                      "order by ProgramName;"
  rst = MyCommon.LRT_Select
  If (rst.rows.count>0)
    Sendb("var functionlist2 = Array(")
    For Each row In rst.Rows
      Sendb(""""  &  MyCommon.NZ(row.Item("ProgramName"), "").ToString().Replace("""", "\""") & """,")
    Next
    Send(""""");")
    Sendb("var vallist2 = Array(")
    If bRequireExternalIds Then
      For Each row In rst.Rows
        lPromoVarId = MyCommon.NZ(row.Item("PromoVarID"), 0)
        MyCommon.QueryStr = "select ExternalID from PromoVariables with (NoLock) where PromoVarID=" & lPromoVarId & ";"
        dt = MyCommon.LXS_Select
        If dt.Rows.Count > 0 Then
                    If Not Long.TryParse(MyCryptLib.SQL_StringDecrypt(dt.Rows(0).Item(0).ToString()), lExternalId) Then lExternalId = 0
          If lExternalId = 0 Then
            If MyExport Is Nothing Then
              MyExport = New Copient.ExportXml(MyCommon)
            End If
            lExternalId = MyExport.ExportOfferCrmGetExternalId(13, lPromoVarId)
            MyCommon.QueryStr = "update PromoVariables with (RowLock) set LastUpdate = GetDate(), ExternalID='" & lExternalId & "' where PromoVarID=" & lPromoVarId & ";"
            MyCommon.LXS_Execute()
          End If
          Sendb("""" & lExternalId & """,")
        End If
      Next
    Else
      For Each row In rst.Rows
        Sendb("""" & row.item("PromoVarID") & """,")
      Next
    End If
    Send(""""");")
  End If
  
  MyCommon.QueryStr = "Select SVProgramID,Name,ExtProgramID from StoredValuePrograms where Deleted=0 order by Name;"
  rst = MyCommon.LRT_Select
  If (rst.rows.count>0)
    Sendb("var functionlist3 = Array(")
    For Each row In rst.Rows
      Sendb("""" & MyCommon.NZ(row.item("Name"), "").ToString().Replace("""", "\""") & """,")
    Next
    Send(""""");")
    Sendb("var vallist3 = Array(")
    If bRequireExternalIds Then
      For Each row In rst.Rows
        If Not Long.TryParse(MyCommon.NZ(row.Item("ExtProgramID"), ""), lExternalId) Then lExternalId = 0
        If lExternalId = 0 Then
          If MyExport Is Nothing Then
            MyExport = New Copient.ExportXml(MyCommon)
          End If
          lExternalId = MyExport.ExportOfferCrmGetExternalId(6, lPromoVarId)
          MyCommon.QueryStr = "update StoredValuePrograms with (RowLock) set ExtProgramID='" & lExternalId & "' where SVProgramID=" & row.Item("SVProgramID") & ";"
          MyCommon.LRT_Execute()
        End If
        Sendb("""" & lExternalId & """,")
      Next
    Else
      For Each row In rst.Rows
        Sendb("""" & row.item("SVProgramID") & """,")
      Next
    End If
    Send(""""");")
  End If

  If sLifetimePointsId <> "" Then
    MyCommon.QueryStr = "select distinct PromoVarID, ProgramName from PointsPrograms " & _
                        "where Deleted=0 and ProgramID in (" & sLifetimePointsId & ") " & _
                        "order by ProgramName;"
    rst = MyCommon.LRT_Select
    If (rst.rows.count>0)
      Sendb("var functionlist4 = Array(")
      For Each row In rst.Rows
        Sendb(""""  &  MyCommon.NZ(row.Item("ProgramName"), "").ToString().Replace("""", "\""") & """,")
      Next
      Send(""""");")
      Sendb("var vallist4 = Array(")
      If bRequireExternalIds Then
        For Each row In rst.Rows
          lPromoVarId = MyCommon.NZ(row.Item("PromoVarID"), 0)
          MyCommon.QueryStr = "select ExternalID from PromoVariables with (NoLock) where PromoVarID=" & lPromoVarId & ";"
          dt = MyCommon.LXS_Select
          If dt.Rows.Count > 0 Then
            If Not Long.TryParse(MyCommon.NZ(dt.Rows(0).Item(0), ""), lExternalId) Then lExternalId = 0
            If lExternalId = 0 Then
              If MyExport Is Nothing Then
                MyExport = New Copient.ExportXml(MyCommon)
              End If
              lExternalId = MyExport.ExportOfferCrmGetExternalId(13, lPromoVarId)
              MyCommon.QueryStr = "update PromoVariables with (RowLock) set LastUpdate = GetDate(), ExternalID='" & lExternalId & "' where PromoVarID=" & lPromoVarId & ";"
              MyCommon.LXS_Execute()
            End If
            Sendb("""" & lExternalId & """,")
          End If
        Next
      Else
        For Each row In rst.Rows
          Sendb("""" & row.item("PromoVarID") & """,")
        Next
      End If
      Send(""""");")
    End If
  End If

%>

// This is the function that refreshes the list after a keypress.
// The maximum number to show can be limited to improve performance with
// huge lists (1000s of entries).
// The function clears the list, and then does a linear search through the
// globally defined array and adds the matches back to the list.
function handleKeyUp(maxNumToShow) {
  var selectObj, textObj, functionListLength;
  var i,  numShown;
  var searchPattern;
  
  document.getElementById("functionselect").size = "10";
  
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
  // Clear the options list
  selectObj.length = 0;
  
  // Loop through the array and re-add matching options
  numShown = 0;
  for(i = 0; i < functionListLength; i++) {
    if(functionlist[i].search(re) != -1) {
      selectObj[numShown] = new Option('<%Sendb(Copient.PhraseLib.Lookup("term.offer", LanguageID)) %> ' + functionlist[i] + ' - <%Sendb(Copient.PhraseLib.Lookup("term.discount", LanguageID)) %> ' + vallist[i],vallist[i]);
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
function handleSelectClick() {
  selectObj = document.forms[0].functionselect;
  textObj = document.forms[0].functioninput;
  
  //selectedValue = selectObj.options[selectObj.selectedIndex].text;
  //selectedValue = selectedValue.replace(/_/g, '-') ;
  selectedValue = document.getElementById("functionselect").value;
  
  if(selectedValue != "") {
    if (textAreaName == "tier0" && parseInt(document.getElementById("NumTiers").value)>0) {
      if (document.getElementById("tier1")){
        textAreaName = "tier1";
      }
    }
    var elemTag = document.getElementById("discTagName");
    if (elemTag.value=="Net#") {
      edInsertNetdol(document.getElementById(textAreaName), selectedValue);
    } else if (elemTag.value == "Initial#") {
      edInsertInitialdol(document.getElementById(textAreaName), selectedValue);
    } else if (elemTag.value == "Earned#") {
      edInsertEarneddol(document.getElementById(textAreaName), selectedValue);
    } else if (elemTag.value == "Redeemed#") {
      edInsertRedeemeddol(document.getElementById(textAreaName), selectedValue);
    }
    showDialogSpan(false, 1, "")
  }
}
// This is the function that refreshes the list after a keypress.
// The maximum number to show can be limited to improve performance with
// huge lists (1000s of entries).
// The function clears the list, and then does a linear search through the
// globally defined array and adds the matches back to the list.
function handleKeyUp2(maxNumToShow) {
  var selectObj, textObj, functionListLength;
  var i, numShown;
  var searchPattern;
  
  document.getElementById("functionselect2").size = "10";
  
  // Set references to the form elements
  selectObj = document.forms[0].functionselect2;
  textObj = document.forms[0].functioninput2;
  
  // Remember the function list length for loop speedup
  functionListLength = functionlist2.length;
  
  // Set the search pattern depending
  if(document.forms[0].functionradio2[0].checked == true) {
    searchPattern = "^"+textObj.value;
  } else {
    searchPattern = textObj.value;
  }
  searchPattern = cleanRegExpString(searchPattern);
  
  // Create a regular expression
  re = new RegExp(searchPattern,"gi");
  // Clear the options list
  selectObj.length = 0;
  
  // Loop through the array and re-add matching options
  numShown = 0;
  for(i = 0; i < functionListLength; i++) {
    if(functionlist2[i].search(re) != -1 && functionlist2[i] != "") {
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
function handleSelectClick2() {
  selectObj = document.forms[0].functionselect2;
  textObj = document.forms[0].functioninput2;
  //selectedValue = selectObj.options[selectObj.selectedIndex].text;
  //selectedValue = selectedValue.replace(/_/g, '-') ;
  
  selectedValue = document.getElementById("functionselect2").value;

  if(selectedValue != "") {
    if (textAreaName == "tier0" && parseInt(document.getElementById("NumTiers").value)>0) {
      if (document.getElementById("tier1")){
        textAreaName = "tier1";
      }
    }
    var elemTag = document.getElementById("ptTagName");
    if (elemTag.value=="Net$") {
      edInsertNetamt(document.getElementById(textAreaName), selectedValue);
    } else if (elemTag.value == "Initial$") {
      edInsertInitialamt(document.getElementById(textAreaName), selectedValue);
    } else if (elemTag.value == "Earned$") {
      edInsertEarnedamt(document.getElementById(textAreaName), selectedValue);
    } else if (elemTag.value == "Redeemed$") {
      edInsertRedeemedamt(document.getElementById(textAreaName), selectedValue);
    }
    showDialogSpan(false, 1, "")
  }
}
// This is the function that refreshes the list after a keypress.
// The maximum number to show can be limited to improve performance with
// huge lists (1000s of entries).
// The function clears the list, and then does a linear search through the
// globally defined array and adds the matches back to the list.
function handleKeyUp3(maxNumToShow) {
  var selectObj, textObj, functionListLength;
  var i, numShown;
  var searchPattern;
  
  document.getElementById("functionselect3").size = "10";
  
  // Set references to the form elements
  selectObj = document.forms[0].functionselect3;
  textObj = document.forms[0].functioninput3;
  
  // Remember the function list length for loop speedup
  functionListLength = functionlist3.length;
  
  // Set the search pattern depending
  if(document.forms[0].functionradio3[0].checked == true) {
    searchPattern = "^"+textObj.value;
  } else {
    searchPattern = textObj.value;
  }
  searchPattern = cleanRegExpString(searchPattern);
  
  // Create a regular expression
  re = new RegExp(searchPattern,"gi");
  // Clear the options list
  selectObj.length = 0;
  
  // Loop through the array and re-add matching options
  numShown = 0;
  for(i = 0; i < functionListLength; i++) {
    if(functionlist3[i].search(re) != -1 && functionlist3[i] != "") {
      selectObj[numShown] = new Option(functionlist3[i], vallist3[i]);
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
function handleSelectClick3() {
  selectObj = document.forms[0].functionselect3;
  textObj = document.forms[0].functioninput3;
  
  selectedValue = document.getElementById("functionselect3").value;
  
  if(selectedValue != "") {
    var elemTag = document.getElementById("svTagName");
    if (elemTag.value == "Svbal") {
      edInsertSvbal(textAreaName, selectedValue);
    } else if (elemTag.value == "Svval") {
      edInsertSvval(textAreaName, selectedValue);
    } else if (elemTag.value == "Svbalexp") {
      edInsertSvbalexp(textAreaName, selectedValue);
    } else if (elemTag.value == "Svvalexp") {
      edInsertSvvalexp(textAreaName, selectedValue);
    } else if (elemTag.value == "Svlimit") {
      edInsertSvlimit(textAreaName, selectedValue);
    } else if (elemTag.value == "Svvalnet") {
      edInsertSvvalnet(textAreaName, selectedValue);
    } else if (elemTag.value == "Svvalinitial") {
      edInsertSvvalinitial(textAreaName, selectedValue);
    } else if (elemTag.value == "Svvalearned") {
      edInsertSvvalearned(textAreaName, selectedValue);
    } else if (elemTag.value == "Svvalredeemed") {
      edInsertSvvalredeemed(textAreaName, selectedValue);
    } else if (elemTag.value == "Svexp_Eom") {
      edInsertSvexp_Eom(textAreaName, selectedValue);
    } else if (elemTag.value == "Eom_Date_Mmdd") {
      edInsertEom_Date_Mmdd(textAreaName, selectedValue);
    }
    showDialogSpan(false, 1, "")
  }
}

// This is the function that refreshes the list after a keypress.
// The maximum number to show can be limited to improve performance with
// huge lists (1000s of entries).
// The function clears the list, and then does a linear search through the
// globally defined array and adds the matches back to the list.
function handleKeyUp4(maxNumToShow) {
  var selectObj, textObj, functionListLength;
  var i, numShown;
  var searchPattern;
  
  document.getElementById("functionselect4").size = "10";
  
  // Set references to the form elements
  selectObj = document.forms[0].functionselect4;
  textObj = document.forms[0].functioninput4;
  
  // Remember the function list length for loop speedup
  functionListLength = functionlist4.length;
  
  // Set the search pattern depending
  if(document.forms[0].functionradio4[0].checked == true) {
    searchPattern = "^"+textObj.value;
  } else {
    searchPattern = textObj.value;
  }
  searchPattern = cleanRegExpString(searchPattern);
  
  // Create a regular expression
  re = new RegExp(searchPattern,"gi");
  // Clear the options list
  selectObj.length = 0;
  
  // Loop through the array and re-add matching options
  numShown = 0;
  for(i = 0; i < functionListLength; i++) {
    if(functionlist4[i].search(re) != -1 && functionlist4[i] != "") {
      selectObj[numShown] = new Option(functionlist4[i],vallist4[i]);
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
function handleSelectClick4() {
  selectObj = document.forms[0].functionselect4;
  textObj = document.forms[0].functioninput4;
  //selectedValue = selectObj.options[selectObj.selectedIndex].text;
  //selectedValue = selectedValue.replace(/_/g, '-') ;
  
  selectedValue = document.getElementById("functionselect4").value;
  
  if(selectedValue != "") {
    if (textAreaName == "tier0" && parseInt(document.getElementById("NumTiers").value)>0) {
      if (document.getElementById("tier1")){
        textAreaName = "tier1";
      }
    }
    var elemTag = document.getElementById("ltTagName");
    if (elemTag.value=="Lifetime#") {
      edInsertLifetimeamt(document.getElementById(textAreaName), selectedValue);
    }
    showDialogSpan(false, 1, "")
  }
}

function showDialogSpan(bShow, type, caption) {
  var elemBox = document.getElementById("dialogbox");
  var elemDisc = document.getElementById("discountselector");
  var elemPt = document.getElementById("pointselector");
  var elemSv = document.getElementById("svselector");
  var elemLt = document.getElementById("lifetimeselector");
  var elemDTag = document.getElementById("discTag");
  var elemPtTag = document.getElementById("ptTag");
  var elemSvTag = document.getElementById("svTag");
  var elemLtTag = document.getElementById("ltTag");
  var elemTag = null;
  
  if (bShow) {
    if (elemDisc != null && elemPt != null) { 
      if (type == 1) {
        elemDisc.style.display = "block";
        elemPt.style.display = "none";
        elemSv.style.display = "none";
        elemLt.style.display = "none";
        if (caption != "" && elemDTag != null) {
          elemDTag.innerHTML = "Tag Type: " + caption
          elemTag = document.getElementById("discTagName");
          if (elemTag != null) {
            elemTag.value = caption;
          }
        }
      } else if (type == 2) {
        elemDisc.style.display = "none";
        elemPt.style.display = "block";
        elemSv.style.display = "none";
        elemLt.style.display = "none";
        if (caption != "" && elemPtTag != null) {
          elemPtTag.innerHTML = "Tag Type: " + caption
          elemTag = document.getElementById("ptTagName");
          if (elemTag != null) {
            elemTag.value = caption;
          }
        }
      } else if (type == 3) {
        elemDisc.style.display = "none";
        elemPt.style.display = "none";
        elemSv.style.display = "block";
        elemLt.style.display = "none";
        if (caption != "" && elemSvTag != null) {
          elemSvTag.innerHTML = "Tag Type: " + caption
          elemTag = document.getElementById("svTagName");
          if (elemTag != null) {
            elemTag.value = caption;
          }
        }
      } else if (type == 8) {
        elemDisc.style.display = "none";
        elemPt.style.display = "none";
        elemSv.style.display = "none";
        elemLt.style.display = "block";
        if (caption != "" && elemLtTag != null) {
          elemLtTag.innerHTML = "Tag Type: " + caption
          elemTag = document.getElementById("ltTagName");
          if (elemTag != null) {
            elemTag.value = caption;
          }
        }
      } 
    }
  }
  if (elemBox != null) {
    elemBox.style.display = (bShow) ? "block" : "none";
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
  strURL += "?" + getQueryString();
  self.xmlHttpReq.open('POST', strURL, true);
  self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
  self.xmlHttpReq.onreadystatechange = function() {
    if (self.xmlHttpReq !=null && self.xmlHttpReq.readyState == 4) {
      updatepage(self.xmlHttpReq.responseText);
    }
  }
  self.xmlHttpReq.send(getQueryString());
}

function getQueryString() {    
  var selElem = document.getElementById('printerselect');
  var qstr = "";
  
  if(selElem != null) {
    qstr = "Mode=MarkupTags&OfferID=<%Sendb(OfferID)%>&EngineID=<%Sendb(OfferEngineID)%>&Phase=3&PrinterTypeID=" + selElem.value + "&NumTiers=" + document.getElementById('NumTiers').value;
  }
  return qstr;
}

function updatepage(str){
  var elemTools = document.getElementById("tools");

  if (elemTools != null) {
    elemTools.innerHTML = str;
  }
}

function getPreviewMsg() {
  var elemTier = document.getElementById("tierselect");
  var elemMsg = null;
  var tierNum = 0;
  var msg = '';
  
  if (elemTier != null) {
    tierNum = elemTier.value;
  }
  elemMsg = document.getElementById("tier" + tierNum);
  if (elemMsg != null) {
    msg = elemMsg.value;
  }  
  return msg;
}

// This is the javascript array holding the function list
// The PrintJavascriptArray ASP function can be used to print this array.
<%
    MyCommon.QueryStr = "select 0 as LimitID, 'None' as Name, RewardLimitTypeID as LimitTypeID, RewardLimit as LimitValue, RewardDistPeriod as LimitPeriod " & _
                        "from OfferRewards with (NoLock) where RewardID=" & RewardID & " " & _
                        "union " & _
                        "select LimitID, Name, LimitTypeID, LimitValue, LimitPeriod " & _
                        "from CM_AdvancedLimits with (NoLock) where Deleted=0 and LimitTypeID=6 order By LimitID;"
    rst2 = MyCommon.LRT_Select
    If rst2.Rows.Count > 0 Then
        Sendb("var ALfunctionlist = Array(")
        For Each row2 In rst2.Rows
            Sendb("""" & MyCommon.NZ(row2.item("Name"), "").ToString().Replace("""", "\""") & """,")
        Next
        Send(""""");")
        
        Sendb("var ALvallist1 = Array(")
        For Each row2 In rst2.Rows
            Sendb("""" & row2.item("LimitID") & """,")
        Next
        Send(""""");")
        
        Sendb("var ALvallist2 = Array(")
        For Each row2 In rst2.Rows
            Sendb("""" & row2.item("LimitPeriod") & """,")
        Next
        Send(""""");")
        
        Sendb("var ALvallist3 = Array(")
        For Each row2 In rst2.Rows
            Sendb("""" & row2.item("LimitValue") & """,")
        Next
        Send(""""");")
    End If
%>

function setlimitsection(bSelect) {
  var elemSelectAdv = document.getElementById("selectadv");
  var elemSelectDay=document.getElementById("selectday");
  var elemPeriod=document.getElementById("limitperiod");
  var elemValue=document.getElementById("limitvalue");
  var elemDisabled=document.getElementById("LimitsDisabled");
 
  if ((bSelect == true) || (elemSelectAdv != null)) {
    if ((elemDisabled == null) || (elemDisabled != null && elemDisabled.value == 'False')) {
      if (elemSelectAdv != null && elemSelectAdv.value == '0') {
        if (elemSelectDay != null) {
          elemSelectDay.disabled = false;
        }
        if (elemPeriod != null) {
          elemPeriod.disabled = false;
        }
        if (elemValue != null) {
          elemValue.disabled = false;
        }
      }
      else {
        if (elemSelectDay != null) {
          elemSelectDay.disabled = true;
        }
        if (elemPeriod != null) {
          elemPeriod.disabled = true;
        }
        if (elemValue != null) {
          elemValue.disabled = true;
        }
      }
    }
 
    for(i = 0; i < ALfunctionlist.length; i++)
    {
      if(elemSelectAdv.value == ALvallist1[i])
      {
        elemPeriod.value = ALvallist2[i];
        elemValue.value = ALvallist3[i];
        if (elemPeriod.value == -1) {
          elemSelectDay.value = '3';
          elemPeriod.style.visibility = 'hidden';
        }
        else if (elemPeriod.value == 0) {
          elemSelectDay.value = '2';
          elemPeriod.style.visibility = 'hidden';
        }
        else
        {
          elemSelectDay.value = '1';
          elemPeriod.style.visibility = 'visible';
        }
        break;
      }
    }
  }
}

function IsValidRegularExpression() 
	{    
	   if (isValidProductList() == true) {
	    var re = new RegExp("[^A-Za-z0-9/\r/\n,]");
        var bAllowSpacesTab  = '<% Sendb(MyCommon.Fetch_SystemOption(207))%>';
		var bAllowHyphen  = '<% Sendb(MyCommon.Fetch_SystemOption(208))%>';
		if(bAllowSpacesTab == 1) {
            re = new RegExp("[^A-Za-z0-9 /\r/\n/\t,]");
		} 
		if(bAllowHyphen == 1) {
            re = new RegExp("[^-A-Za-z0-9/\r/\n,]");
		}
		if(bAllowSpacesTab == 1 && bAllowHyphen == 1) {
            re = new RegExp("[^-A-Za-z0-9 /\r/\n/\t,]");
		}		
		if (document.getElementById("pasteproducts").value.match(re)) {
			alert('<% Sendb(Copient.PhraseLib.Lookup("pgroup-edit.invaliddata", LanguageID))%>');
			return false;
		} 
		else {    
		return true;
		}
	  }
     else {
	   return false;
	 }	  
		
		
}

function selecttype(bSelect){
  var elemType = document.getElementById("type");
  var elemCheckStatus = document.getElementById("optcheck");
  var elemPrintOnBack=document.getElementById("printonback");
    
  if (elemType != null && (elemType.value == '2') || (elemType.value == '3')){
    elemCheckStatus.style.visibility ='visible';
    elemPrintOnBack.style.visibility = 'visible'; 
   }else {
     elemCheckStatus.checked = false; 
    elemCheckStatus.style.visibility = 'hidden';
    elemPrintOnBack.style.visibility = 'hidden'; 
   }
}

  
function hideprintonback(bcheck){
  var elemType = document.getElementById("type");
  var elemCheckStatus = document.getElementById("optcheck");
  var elemPrintOnBack=document.getElementById("printonback");
  
  if (elemType != null && (elemType.value == '1' || elemType.value == '4')){
    elemCheckStatus.checked = false; 
    elemCheckStatus.style.visibility = 'hidden';
    elemPrintOnBack.style.visibility = 'hidden'; 
  }
   
  if ((elemType.value == '2' || elemType.value == '3') && (elemCheckStatus.checked == true)){
    document.getElementById("PrintOnBackReceipt").value = '1';
  }else {
    document.getElementById("PrintOnBackReceipt").value = '0';
  }
}

function setperiodsection(bSelect) {
  var elemSelectDay = document.getElementById("selectday");
  var elemPeriod=document.getElementById("limitperiod");
  var elemOriginalPeriod=document.getElementById("OriginalPeriod");
  var elemImpliedPeriod=document.getElementById("ImpliedPeriod");

  if (elemSelectDay != null && (elemSelectDay.value == '2') || (elemSelectDay.value == '3')) {
    if (elemPeriod != null) {
      elemPeriod.style.visibility = 'hidden';
    }
    if (elemSelectDay.value == '2') {
      elemImpliedPeriod.value = '0';
      elemPeriod.value = '0';
    }
    else {
      elemImpliedPeriod.value = '-1';
      elemPeriod.value = '-1';
    }
  }
  else {
    if (elemPeriod != null) {
      if (bSelect && elemOriginalPeriod != null) {
        if ((elemOriginalPeriod.value == '-1') || (elemOriginalPeriod.value == '0')) {
          elemPeriod.value = '0';
        }
        else {
          elemPeriod.value = elemOriginalPeriod.value;
          elemImpliedPeriod.value = elemOriginalPeriod.value;
        }
      }
      elemPeriod.style.visibility = 'visible';
    }
  }
}

</script>

<form action="offer-rew-pmsg.aspx" id="mainform" name="mainform" method="post" >
  <div id="intro">
    <input type="hidden" id="OfferID" name="OfferID" value="<% sendb(OfferID) %>" />
    <input type="hidden" id="Name" name="Name" value="<% sendb(Name) %>" />
    <input type="hidden" id="RewardID" name="RewardID" value="<% sendb(RewardID) %>" />
    <input type="hidden" id="OriginalPeriod" name="OriginalPeriod" value="<% sendb(DistPeriod) %>" />
    <input type="hidden" id="ImpliedPeriod" name="ImpliedPeriod" value="<% sendb(DistPeriod) %>" />
    <input type="hidden" id="LimitsDisabled" name="LimitsDisabled" value="<% sendb(bUseTemplateLocks and bDisallowEditLimit) %>" />
    <input type="hidden" id="IsTemplate" name="IsTemplate" value="<%Sendb(IIf(istemplate, "IsTemplate", "Not"))%>" />
    <input type="hidden" id="PrintOnBackReceipt" name="PrintOnBackReceipt" value="""" />
    <input type="hidden" id="HdSelectedProductGroup" name="HdSelectedProductGroup" value="<%sendb(SelectedProductGroupId) %>" />
    <input type="hidden" id="HdExcludedProdGroupID" name="HdExcludedProdGroupID" value="<% sendb(ExcludedProdGroupID) %>" />
    <input type="hidden" id="hdSearchType" name="hdSearchType" value="""" />
	<% If MyCommon.Fetch_SystemOption(97) = "1" Then
		Send("<input type=""hidden"" id=""NumericOnly"" name=""NumericOnly"" value=""true"" />")
	  End If %>
    <h1 id="title">
      <% Sendb(Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & Copient.PhraseLib.Lookup("term.pmsgreward", LanguageID))%>
    </h1>
    <div id="controls">
      <% If (IsTemplate) Then%>
      <span class="temp">
        <input type="checkbox" class="tempcheck" id="temp-employees" name="Disallow_Edit"<% if(disallow_edit)then sendb(" checked=""checked""") %> />
        <label for="temp-employees"><% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
      </span>
      <% End If%>
      <% If Tiered = 0 Then%>
      <button class="regular" id="preview" name="preview" type="button" onclick="javascript:openPreviewPopup('offer-rew-pmsgpreview.aspx?PrinterTypeID='+document.getElementById('printerselect').value+'&Message='+escape(document.getElementById('tier0').value))">
        <% Sendb(Copient.PhraseLib.Lookup("term.preview", LanguageID))%>
      </button>
      <% Else%>
      <button class="regular" id="preview" name="preview" type="button" onclick="javascript:openPreviewPopup('offer-rew-pmsgpreview.aspx?PrinterTypeID='+document.getElementById('printerselect').value+'&Message='+escape(getPreviewMsg()))">
        <% Sendb(Copient.PhraseLib.Lookup("term.preview", LanguageID))%>
      </button>
      <% End If%>
      <% If Not (IsTemplate) Then
           If (Logix.UserRoles.EditOffer And Not (bUseTemplateLocks And Disallow_Edit)) Then Send_Save()
         Else
           If (Logix.UserRoles.EditTemplates) Then Send_Save()
         End If    
      %>
    </div>
  </div>
  <div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <% If MyCommon.Fetch_SystemOption(222) = "0" Then %>
	  <input type="radio" id="pgselectortype1" name="pgselectortype" onclick="javascript:ProductGroupTypeSelection();"
      <% if(ByExistingPGSelector) then sendb(" checked=""checked""") %> value="existingadd" /><label
      for="pgselectortype1"><% Sendb(Copient.PhraseLib.Lookup("term.selectexistingprodgroups", LanguageID))%></label>
      <input type="radio" id="pgselectortype2" name="pgselectortype" onclick="javascript:ProductGroupTypeSelection();"
      <% if(Not ByExistingPGSelector) then sendb(" checked=""checked""") %> value="directadd" /><label
      for="pgselectortype2">Add products to reward</label>
	<%Else%>
	  <input type="radio" id="pgselectortype2" name="pgselectortype" onclick="javascript:ProductGroupTypeSelection();"
      <% if(Not ByExistingPGSelector) then sendb(" checked=""checked""") %> value="directadd" /><label
      for="pgselectortype2">Add products to reward</label>
	  <input type="radio" id="pgselectortype1" name="pgselectortype" onclick="javascript:ProductGroupTypeSelection();"
      <% if(ByExistingPGSelector) then sendb(" checked=""checked""") %> value="existingadd" /><label
      for="pgselectortype1"><% Sendb(Copient.PhraseLib.Lookup("term.selectexistingprodgroups", LanguageID))%></label>
	<%End If%>
    <div id="columnfull">
      <%  If (Request.Form("pgroup-avail") <> "" And Request.Form("pgroup-add1") <> "") Then
              SelectedProductGroupId = Request.Form("pgroup-avail")
          ElseIf (MyCommon.Extract_Val(Request.Form("HdSelectedProductGroup")) <> "0" And Request.Form("pgroup-rem1") Is Nothing) Then
              SelectedProductGroupId = MyCommon.Extract_Val(Request.Form("HdSelectedProductGroup"))
          ElseIf Not Request.Form("pgroup-rem1") Is Nothing And Request.Form("pgroup-add2") Is Nothing Then
              ' Setting -1 Value 
              SelectedProductGroupId = -1
              ExcludedProdGroupID = -1
          ElseIf Request.Form("pgroup-rem1") Is Nothing And Not Request.Form("pgroup-add2") Is Nothing Then
              SelectedProductGroupId = -1
          End If%>  
          
         <%  If (Request.Form("pgroup-avail") <> "" And Request.Form("pgroup-add2") <> "") Then
                 ExcludedProdGroupID = Request.Form("pgroup-avail")
             ElseIf (MyCommon.Extract_Val(Request.Form("HdExcludedProdGroupID")) <> "0" And Request.Form("pgroup-rem2") Is Nothing) Then
                 ExcludedProdGroupID = MyCommon.Extract_Val(Request.Form("HdExcludedProdGroupID"))
             ElseIf Not Request.Form("pgroup-rem2") Is Nothing And Request.Form("pgroup-add1") Is Nothing Then
                 ' Setting -1 Value
                 ExcludedProdGroupID = -1
                 If (MyCommon.Extract_Val(Request.Form("HdSelectedProductGroup")) = "0") Then
                     SelectedProductGroupId = -1
                 End If
                 
             ElseIf Request.Form("pgroup-rem2") Is Nothing And Not Request.Form("pgroup-add1") Is Nothing Then
                 ExcludedProdGroupID = -1
             End If %>   
                
             <%  If (Request.Form("functioninputSearch")) <> "" Then
                     SearchProductGrouptext = Request.Form("functioninputSearch")
                 End If
                 If (Request.Form("hdSearchType") <> "") Then
                     If (MyCommon.Extract_Val(Request.Form("hdSearchType"))) = "1" Then
                         RadioButtonforStart = True
                         RadioButtonForContain = False
                     ElseIf (MyCommon.Extract_Val(Request.Form("hdSearchType"))) = "2" Then
                         RadioButtonForContain = True
                         RadioButtonforStart = False
                     End If
                 End If
             %> 
      <%
          Send_DirectProductAddSelector(Logix, ByExistingPGSelector, ShowAllItems , GroupSize , rstItems , ProductsWithoutDesc , descriptionItem , ByAddSingleProduct , IDLength, GName, True)
          %>   
      <% Send_ProductGroupSelector(Logix, TransactionLevelSelected, bUseTemplateLocks, bDisallowEditPg, SelectedItem, ExcludedItem, RewardID, 0, IsTemplate, bDisallowEditPg, bStoreUser, sValidLocIDs, sValidSU, SelectedProductGroupId, ExcludedProdGroupID, ByExistingPGSelector, SearchProductGrouptext, RadioButtonforStart, RadioButtonForContain)%>
      
    </div>
    <div id="column1x">
      <%If Not TransactionLevelSelected Then%>
      <div class="box" id="distribution">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.distribution", LanguageID))%>
          </span>
          <% If (IsTemplate) Then%>
          <span class="tempRequire">
            <input type="checkbox" class="tempcheck" id="DisallowEditDist1" name="DisallowEditDist"
              <% if(bDisallowEditDist)then send(" checked=""checked""") %> />
            <label for="temp-Tiers">
              <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
            </label>
          </span>
          <% ElseIf (bUseTemplateLocks And bDisallowEditDist) Then%>
          <span class="tempRequire">
            <input type="checkbox" class="tempcheck" id="DisallowEditDist2" name="DisallowEditDist"
              disabled="disabled" checked="checked" />
            <label for="temp-Tiers">
              <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
            </label>
          </span>
          <% End If%>
        </h2>
        <% 
          If (bUseTemplateLocks And bDisallowEditDist) Then
            sDisabled = " disabled=""disabled"""
          Else
            sDisabled = ""
          End If
        %>
        <input id="triggerbogo" name="trigger" type="radio" <% if(valueradio=1)then sendb(" checked=""checked""") %>
          value="1" <% Send(sDisabled) %> />
        <label for="triggerbogo">
          <% Sendb(Copient.PhraseLib.Lookup("reward.messageevery", LanguageID))%>
        </label>
        <br />
        &nbsp; &nbsp; &nbsp; &nbsp;
        <label for="Xbox">
          <% Sendb(Copient.PhraseLib.Lookup("term.mustpurchase", LanguageID))%>
        </label>
        <input class="shortest" id="Xbox" name="Xbox" maxlength="9" type="text" <% if(valueradio=1)then sendb(" value=""" & triggerqty & """ ") %><% Send(sDisabled) %> />
        <% Sendb(Copient.PhraseLib.Lookup("term.item(s)", LanguageID))%>
        <br />
        <input id="triggerbxgy" name="trigger" type="radio" value="2" <% if(valueradio=2)then sendb(" checked=""checked""") %><% Send(sDisabled) %> />
        <label for="triggerbxgy">
          <% Sendb(Copient.PhraseLib.Lookup("term.buy", LanguageID))%>
        </label>
        <input class="shortest" id="bxgy1" name="Xbox2" maxlength="9" type="text" <% if(valueradio=2)then sendb(" value=""" & triggerqty-applytolimit & """ ") %><% Send(sDisabled) %> />,
        <% Sendb(Copient.PhraseLib.Lookup("reward.givemessageto", LanguageID))%>
        <input class="shortest" id="bxgy2" name="Ybox2" maxlength="9" type="text" <% if(valueradio=2)then sendb(" value=""" & applytolimit & """ ") %><% Send(sDisabled) %> /><br />
        <hr class="hidden" />
      </div>
      <% End If%>
      <div class="box" id="limits">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.limits", LanguageID))%>
          </span>
          <% If (IsTemplate Or (bUseTemplateLocks And bDisallowEditLimit)) Then%>
          <span class="tempRequire">
            <input type="checkbox" class="tempcheck" id="DisallowEditLimit" name="DisallowEditLimit"
              <% if(bDisallowEditLimit)then send(" checked=""checked""") %><% If(bUseTemplateLocks) Then sendb(" disabled=""disabled""") %> />
            <label for="temp-Tiers">
              <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
            </label>
          </span>
          <br class="printonly" />
          <% End If%>
        </h2>
        <%
          MyCommon.QueryStr = "Select LimitId, Name from CM_AdvancedLimits with (NoLock) where Deleted=0 and LimitTypeID=6 order By Name;"
          rst = MyCommon.LRT_Select
          If rst.Rows.Count > 0 Then
        %>
        <label for="selectadv"><% Sendb(Copient.PhraseLib.Lookup("term.advlimits", LanguageID))%>:</label>
        <select id="selectadv" name="selectadv" class="mediumplus" onchange="setlimitsection(true);"
          <% If(bUseTemplateLocks and bDisallowEditLimit) Then sendb(" disabled=""disabled""") %>>
          <%
            Sendb("<option value=""0""")
            If (AdvancedLimitID = 0) Then
              Sendb(" selected=""selected""")
            End If
            Sendb(">" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "</option>")
            For Each row In rst.Rows
              Sendb("<option value=""" & row.Item("LimitID") & """")
              If (AdvancedLimitID = MyCommon.Extract_Val(MyCommon.NZ(row.Item("LimitID"), 0))) Then
                Sendb(" selected=""selected""")
              End If
              Sendb(">")
              Sendb(row.Item("Name"))
              Sendb("</option>")
            Next
          %>
        </select>
        <br class="half" />
        <% End If%>
        <br class="half" />
        <input class="shorter" id="limitvalue" name="limitvalue" maxlength="9" type="text"
          value="<% sendb(RewardLimit) %>" <% If(bUseTemplateLocks and bDisallowEditLimit) Then sendb(" disabled=""disabled""") %> />
        &nbsp;<% Sendb(Copient.PhraseLib.Lookup("term.per", LanguageID))%>
        <input class="shortest" id="limitperiod" name="form_DistPeriod" maxlength="4" type="text" value="<% sendb(DistPeriod) %>"
          <% If(bUseTemplateLocks and bDisallowEditLimit) Then sendb(" disabled=""disabled""") %> />
        <select id="selectday" name="selectday" class="short" onchange="setperiodsection(true);"
          <% If(bUseTemplateLocks and bDisallowEditLimit) Then sendb(" disabled=""disabled""") %>>
          <option value="1" <% if(distperiod>0)then sendb(" selected=""selected""") %>>
            <% Sendb(Copient.PhraseLib.Lookup("term.days", LanguageID))%>
          </option>
          <option value="2" <% if(distperiod=0)then sendb(" selected=""selected""") %>>
            <% Sendb(Copient.PhraseLib.Lookup("term.transaction", LanguageID))%>
          </option>
          <option value="3" <% if(distperiod=-1)then sendb(" selected=""selected""") %>>
            <% Sendb(Copient.PhraseLib.Lookup("term.customer", LanguageID))%>
          </option>
        </select>
        <hr class="hidden" />
      </div>
      <div class="box" id="sponsor">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.sponsor", LanguageID))%>
          </span>
          <% If (IsTemplate) Then%>
          <span class="tempRequire">
            <input type="checkbox" class="tempcheck" id="DisallowEditSpon1" name="DisallowEditSpon"
              <% if (bDisallowEditSpon) then send(" checked=""checked""") %> />
            <label for="temp-Tiers">
              <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
            </label>
          </span>
          <% ElseIf (bUseTemplateLocks And bDisallowEditSpon) Then%>
          <span class="tempRequire">
            <input type="checkbox" class="tempcheck" id="DisallowEditSpon2" name="DisallowEditSpon"
              disabled="disabled" checked="checked" />
            <label for="temp-Tiers">
              <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
            </label>
          </span>
          <% End If%>
        </h2>
        <%
          MyCommon.QueryStr = "select SponsorID, Description, PhraseID from Sponsors with (NoLock)"
          rst = MyCommon.LRT_Select()
          For Each row In rst.Rows
            Sendb("<input class=""radio"" id=""" & StrConv(row.Item("Description"), VbStrConv.Lowercase) & """ name=""sponsor"" type=""radio"" value=""" & row.Item("SponsorID") & """")
            If SponsorID = row.Item("SponsorID") Then
              Sendb(" checked=""checked""")
            End If
            If (bUseTemplateLocks And bDisallowEditSpon) Then
              Sendb(" disabled=""disabled""")
            End If
            Send(" /><label for=""" & StrConv(row.Item("Description"), VbStrConv.Lowercase) & """>" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID) & "</label>")
          Next
        %>
        <hr class="hidden" />
      </div>
      <div class="box" id="misc">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.advancedoptions", LanguageID))%>
          </span>
          <% If (IsTemplate Or (bUseTemplateLocks And bDisallowEditAdv)) Then%>
          <span class="tempRequire">
            <input type="checkbox" class="tempcheck" id="DisallowEditAdv" name="DisallowEditAdv"
              <% if(bDisallowEditAdv)then send(" checked=""checked""") %><% If(bUseTemplateLocks) Then sendb(" disabled=""disabled""") %> />
            <label for="temp-Tiers">
              <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
            </label>
          </span>
          <br class="printonly" />
          <% End If%>
        </h2>
        <%
          If (TransactionLevelSelected) Then
            sDisabled = " disabled=""disabled"""
          Else
            If (bUseTemplateLocks And bDisallowEditAdv) Then
              sDisabled = " disabled=""disabled"""
            Else
              sDisabled = ""
            End If
          End If
        %>
        <input class="checkbox" id="promote" name="promote" type="checkbox" <% sendb(sDisabled) %><% if(promotetotranslevel=true)then sendb(" checked=""checked""") %> /><label
          for="promote"><% Sendb(Copient.PhraseLib.Lookup("reward.promote", LanguageID))%></label><br />
        <hr class="hidden" />
      </div>
      <div class="box" id="printer">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.printer", LanguageID))%>
          </span>
        </h2>
        <label for="printerselect">
          <% Sendb(Copient.PhraseLib.Lookup("reward.selectprinter", LanguageID))%>:</label><br />
        <select id="printerselect" name="printerselect" onchange="xmlhttpPost('PrintedMessageFeeds.aspx');">
          <%
            MyCommon.QueryStr = "select PTy.PrinterTypeID,PTy.PageWidth,PTy.Name,PTy.PhraseID,PTy.Installed,PTy.DefaultPrinter,PEPT.EngineID " & _
                                "from PrinterTypes as PTy with (NoLock) " & _
                                "inner join PromoEnginePrinterTypes as PEPT with (NoLock) on PEPT.PrinterTypeID=PTy.PrinterTypeID " & _
                                "where PTy.Installed=1 and PEPT.EngineID=" & OfferEngineID & " order by PTy.Name"
            rst2 = MyCommon.LRT_Select
            Send("<option value=""999"">" & Copient.PhraseLib.Lookup("term.allprinters", LanguageID) & "</option>")
            For Each row In rst2.Rows
              If (row.Item("PrinterTypeID") > 0) Then
                PrinterWidthBuf.Append("<input type=""hidden"" name=""PT" & row.Item("PrinterTypeID") & """ id=""PT" & row.Item("PrinterTypeID") & """ value=""" & MyCommon.NZ(row.Item("PageWidth"), 50) & """ />" & vbCrLf)
                Send("<option value=""" & row.Item("PrinterTypeID") & """ ")
                If row.Item("DefaultPrinter") Or rst2.Rows.Count = 1 Then
                  Sendb(" selected=""selected""")
                End If
                Send(">" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID) & "</option>")
              Else
              End If
            Next
          %>
        </select>
        <%
          If (Tiered = 0) Then
          Else
            MyCommon.QueryStr = "select Tiered,O.Numtiers,O.OfferID,CT.TierLevel from Offerrewards as OC with (NoLock) " & _
                                "left join Offers as O with (NoLock) on O.OfferID=OC.OfferID " & _
                                "left join PrintedMessageTiers as CT with (NoLock) on OC.LinkID=CT.MessageID " & _
                                "where OC.RewardID=" & RewardID
            rst3 = MyCommon.LRT_Select()
            NumTiers = rst3.Rows(0).Item("Numtiers")
            i = 1
            Send("<br />")
            Send("<br class=""half"" />")
            Send("<label for=""tierselect"">" & Copient.PhraseLib.Lookup("reward.tierpreview", LanguageID) & ":</label><br />")
            Send("<select id=""tierselect"" name=""tierselect"">")
            For i = 1 To NumTiers
              Send("<option value=""" & i & """>" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & i & "</option>")
            Next
            Send("</select>")
          End If
        %>
        <hr class="hidden" />
      </div>
    </div>
    <div id="gutter">
    </div>
    <div id="column2x">
      <div class="box" id="message">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.message", LanguageID))%>
          </span>
          <% If (IsTemplate) Then%>
          <span class="tempRequire">
            <input type="checkbox" class="tempcheck" id="DisallowEditMsg1" name="DisallowEditMsg"
              <% if(bDisallowEditMsg)then send(" checked=""checked""") %> />
            <label for="temp-Tiers">
              <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
            </label>
          </span>
          <% ElseIf (bUseTemplateLocks And bDisallowEditMsg) Then%>
          <span class="tempRequire">
            <input type="checkbox" class="tempcheck" id="DisallowEditMsg2" name="DisallowEditMsg"
              disabled="disabled" checked="checked" />
            <label for="temp-Tiers">
              <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
            </label>
          </span>
          <% End If%>
        </h2>
        <%
          If (bUseTemplateLocks And bDisallowEditMsg) Then
            sDisabled = " disabled=""disabled"""
          Else
            sDisabled = ""
          End If
        %>
        <label for="type"><% Sendb(Copient.PhraseLib.Lookup("term.location", LanguageID))%>:</label><br />
        <select id="type" name="type" <% send(sDisabled) %> onchange="selecttype(true);">
          <%
          MyCommon.QueryStr = "select TypeID, PhraseID from PrintedMessageTypes with (NoLock) where EngineID = 0 order by TypeID"
            rst = MyCommon.LRT_Select()
            For Each row In rst.Rows
              Sendb("<option value=""" & row.Item("TypeID") & """")
              If MessageTypeID = row.Item("TypeID") Then
                Sendb(" selected=""selected""")
              End If
              Send(">")
              Send(Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID))
              Send("</option>")
            Next
          %>
        </select>
        <span id ="idprintonback">
          <input type="checkbox" class="tempcheck" id="optcheck" name="optcheck" onchange="hideprintonback(true);"<%If (CheckedStatus = true) then send(" checked=""checked""") %> />&nbsp;
          <label id="printonback" for="optcheck"><% Sendb(Copient.PhraseLib.Lookup("term.printonback", LanguageID))%></label>
        </span>
        <br />
        <br class="half" />
        
        <%
          MyCommon.QueryStr = "select CT.MessageID,Tiered,O.Numtiers,O.OfferID,CT.TierLevel,CT.BodyText from Offerrewards as OC with (NoLock) " & _
                              "left join Offers as O with (NoLock) on O.OfferID=OC.OfferID " & _
                              "left join PrintedMessageTiers as CT with (NoLock) on OC.LinkID=CT.MessageID " & _
                              "where OC.RewardID=" & RewardID & " order by TierLevel asc;"
          'rst = MyCommon.LRT_Select()
          Dim tierDT As DataTable = MyCommon.LRT_Select()
          q = 1
          Dim targetTextarea As String
          For Each tierRow As DataRow In tierDT.Rows
            If (tierRow.Item("Tiered") = False) Then
              targetTextarea = "tier0"
              Send("<div class=""pmsgwrap"">")
              Send("    <label for=""tier0""><b>" & Copient.PhraseLib.Lookup("term.message", LanguageID) & "</b></label><br />")
              Send("    <textarea id=""tier0"" name=""tier0"" cols=""50"" rows=""7"" wrap=""soft"" onfocus=""textAreaName='tier0';""" & sDisabled & ">" & tierRow.Item("BodyText") & "</textarea><br />")
              Send("</div>")
            Else
              targetTextarea = "tier1"
              Send("<div class=""pmsgwrap"">")
              Send("    <label for=""tier" & q & """><b>" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & q & " " & Copient.PhraseLib.Lookup("term.message", LanguageID) & ":</b></label><br />")
              Send("    <textarea id=""tier" & q & """ name=""tier" & q & """ cols=""50"" rows=""7"" wrap=""soft"" onfocus=""textAreaName='tier" & q & "';""" & sDisabled & ">" & tierRow.Item("BodyText") & "</textarea><br />")
              Send("</div>")
              If q = 1 Then
                Send("    <script type=""text/javascript"">var edCanvas = document.getElementById('tier" & q & "');</script>")
              End If
            End If
            q = q + 1
          Next
          'Send("    <input type=""hidden"" id=""NumTiers"" name=""NumTiers"" value=""" & row.Item("NumTiers") & """ />")
          Send("    <input type=""hidden"" id=""NumTiers"" name=""NumTiers"" value=""" & MyCommon.NZ(tierDT.Rows(0).Item("NumTiers"), 0) & """ />")
                    
          If Not (bUseTemplateLocks And bDisallowEditMsg) Then
            ' --- TOOLBAR ---
            Send("     <div id=""ed_toolbar"" style=""background-color:#d0d0d0;text-align:center;"">")
            Send("     <div id=""tools"">")
            Sendb("      ")
            MyCommon.QueryStr = "select distinct MT.MarkupID, MT.Tag, MT.Description, MT.PhraseID, MT.NumParams, " & _
                            "MT.Param1Name, MT.Param1PhraseID, MT.Param2Name, MT.Param2PhraseID, " & _
                            "MT.Param3Name, MT.Param3PhraseID, MT.Param4Name, MT.Param4PhraseID, " & _
                            "MT.DisplayOrder, MT.CentralRendered, MT.ButtonText, " & _
                            "MTU.RewardTypeID, MTU.EngineID from MarkupTags as MT with (NoLock) " & _
                            "inner join MarkupTagUsage as MTU with (NoLock) on MT.MarkupID=MTU.MarkupID " & _
                            "where MTU.RewardTypeID=3"
            If ExcentusInstalled Then
              MyCommon.QueryStr &= " and MTU.EngineID in (" & OfferEngineID & ", 8) order by MT.DisplayOrder;"
            Else
              MyCommon.QueryStr &= " and MTU.EngineID=" & OfferEngineID & " order by MT.DisplayOrder;"
            End If
            rst = MyCommon.LRT_Select
            Dim cleanid As String
            For Each row In rst.Rows
              cleanid = row.Item("ButtonText")
              cleanid = cleanid.Replace("#", "Amt")
              cleanid = cleanid.Replace("$", "Dol")
              cleanid = cleanid.Replace("/", "Off")
              If (cleanid = "NETDol") Or (cleanid = "INITIALDol") Or (cleanid = "EARNEDDol") Or (cleanid = "REDEEMEDDol") Then
                        Sendb("<input type=""button"" id=""ed_" & (StrConv(cleanid, VbStrConv.Lowercase)) & """ class=""ed_button"" " & DisabledAttribute & " title=""" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & """ onclick=""javascript:showDialogSpan(true, 2, this.value);"" value=""" & (StrConv(row.Item("ButtonText"), VbStrConv.ProperCase)) & """ />")
              ElseIf (cleanid = "NETAmt") Or (cleanid = "INITIALAmt") Or (cleanid = "EARNEDAmt") Or (cleanid = "REDEEMEDAmt") Then
                        Sendb("<input type=""button"" id=""ed_" & (StrConv(cleanid, VbStrConv.Lowercase)) & """ class=""ed_button"" " & DisabledAttribute & " title=""" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & """ onclick=""javascript:showDialogSpan(true, 1, this.value);"" value=""" & (StrConv(row.Item("ButtonText"), VbStrConv.ProperCase)) & """ />")
              ElseIf (cleanid = "Svbal") Or (cleanid = "Svval") Or (cleanid = "Svbalexp") Or (cleanid = "Svvalexp") Or (cleanid = "Svlimit") Or (cleanid = "Svvalnet") Or (cleanid = "Svvalinitial") Or (cleanid = "Svvalearned") Or (cleanid = "Svvalredeemed") Or (cleanid = "Svexp_Eom") Then
                Sendb("<input type=""button"" id=""ed_" & (StrConv(cleanid, VbStrConv.Lowercase)) & """ class=""ed_button"" " & DisabledAttribute & " title=""" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & """ onclick=""javascript:showDialogSpan(true, 3, this.value);"" value=""" & (StrConv(row.Item("ButtonText"), VbStrConv.ProperCase)) & """ />")
              ElseIf (cleanid = "LIFETIMEAmt") Then
                Sendb("<input type=""button"" id=""ed_" & (StrConv(cleanid, VbStrConv.Lowercase)) & """ class=""ed_button"" " & DisabledAttribute & " title=""" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & """ onclick=""javascript:showDialogSpan(true, 8, this.value);"" value=""" & (StrConv(row.Item("ButtonText"), VbStrConv.ProperCase)) & """ />")
              Else
                Sendb("<input type=""button"" id=""ed_" & (StrConv(cleanid, VbStrConv.Lowercase)) & """ class=""ed_button"" " & DisabledAttribute & " title=""" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & """ onclick=""edInsert" & (StrConv(cleanid, VbStrConv.ProperCase)) & "(" & targetTextarea & ");"" value=""" & (StrConv(row.Item("ButtonText"), VbStrConv.ProperCase)) & """ />")
              End If
            Next
            Send("      <br />")
            Send("     </div>")
            Send("     </div>")
          End If
        %>
      </div>
    </div>
    <div id="dialogbox">
      <div id="discountselector">
        <div id="discTag">
        </div>
        <br />
        <input type="hidden" name="discTagName" id="discTagName" value="" />
      <b><% Sendb(Copient.PhraseLib.Lookup("offer-rew.selectdiscount", LanguageID))%>:</b><br />
      <input type="radio" id="functionradio1a" name="functionradio" checked="checked" /><label for="functionradio"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
      <input type="radio" id="functionradio1b" name="functionradio" /><label for="functionradio"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
      <input class="medium" onkeyup="handleKeyUp(200);" id="functioninput" name="functioninput" type="text" value="" /><br />
      <select onclick="handleSelectClick();" id="functionselect" name="functionselect" size="10" style="width: 220px;">
          <%
            MyCommon.QueryStr = "Select RewardID, OfferID, RewardTypeID, Deleted from OfferRewards with (NoLock) " & _
                                "where RewardTypeID=1 and Deleted=0 order by OfferID"
            rst = MyCommon.LRT_Select
            For Each row In rst.Rows
              Send("<option value=""" & row.Item("RewardID") & """>" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " " & row.Item("OfferID") & " - " & Copient.PhraseLib.Lookup("term.discount", LanguageID) & " " & row.Item("RewardID") & "</option>")
            Next
          %>
        </select>
        <br />
        <br />
        <input type="button" id="btnClose1" name="btnClose1" value="<% Sendb(Copient.PhraseLib.Lookup("term.close", LanguageID))%>" onclick="javascript:showDialogSpan(false, 1);" />
      </div>
      <div id="pointselector">
        <div id="ptTag">
        </div>
        <br />
        <center>
          <input type="hidden" name="ptTagName" id="ptTagName" value="" />
        <b><% Sendb(Copient.PhraseLib.Lookup("offer-rew.selectpoints", LanguageID))%>:</b><br />
        <input type="radio" id="functionradio2a" name="functionradio2" checked="checked" /><label for="functionradio2"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
        <input type="radio" id="functionradio2b" name="functionradio2" /><label for="functionradio2"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
        <input class="medium" onkeyup="handleKeyUp2(200);" id="functioninput2" name="functioninput2" type="text" value="" /><br />
        <select onclick="handleSelectClick2();" id="functionselect2" name="functionselect2" size="10" style="width: 220px;">
            <%
              MyCommon.QueryStr = "select distinct PP.PromoVarID, PP.ProgramName from PointsPrograms PP " & _
                                  "inner join RewardPoints RP on RP.ProgramID=PP.ProgramID " & _
                                  "inner join OfferRewards OREW on RP.RewardPointsID=OREW.LinkID " & _
                                  "where (OREW.RewardTypeID=2 or OREW.RewardTypeID=13) and OREW.Deleted=0 and PP.Deleted=0 " & _
                                  "order by ProgramName;"
              rst = MyCommon.LRT_Select
              If bRequireExternalIds Then
                For Each row In rst.Rows
                  lPromoVarId = MyCommon.NZ(row.Item("PromoVarID"), 0)
                  MyCommon.QueryStr = "select ExternalID from PromoVariables with (NoLock) where PromoVarID=" & lPromoVarId & ";"
                  dt = MyCommon.LXS_Select
                  If dt.Rows.Count > 0 Then
                    If Not Long.TryParse(MyCommon.NZ(dt.Rows(0).Item(0), ""), lExternalId) Then lExternalId = 0
                    If lExternalId = 0 Then
                      If MyExport Is Nothing Then
                        MyExport = New Copient.ExportXml(MyCommon)
                      End If
                      lExternalId = MyExport.ExportOfferCrmGetExternalId(13, lPromoVarId)
                      MyCommon.QueryStr = "update PromoVariables with (RowLock) set LastUpdate = GetDate(), ExternalID='" & lExternalId & "' where PromoVarID=" & lPromoVarId & ";"
                      MyCommon.LXS_Execute()
                    End If
                    Send("<option value=" & lExternalId & ">" & row.Item("ProgramName") & "</option>")
                  End If
                Next
              Else
                For Each row In rst.Rows
                  Send("<option value=" & row.Item("PromoVarID") & ">" & row.Item("ProgramName") & "</option>")
                Next
              End If

            %>
          </select>
          <br />
          <br />
          <input type="button" id="btnClose2" name="btnClose2" value="<% Sendb(Copient.PhraseLib.Lookup("term.close", LanguageID))%>" onclick="javascript:showDialogSpan(false, 2);" />
        </center>
      </div>
      <div id="svselector">
        <div id="svTag">
        </div>
        <br />
        <input type="hidden" name="svTagName" id="svTagName" value="" />
      <b><% Sendb(Copient.PhraseLib.Lookup("offer-rew.selectsv", LanguageID))%>:</b><br />
      <input type="radio" id="functionradio3a" name="functionradio3" checked="checked" /><label for="functionradio3"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
      <input type="radio" id="functionradio3b" name="functionradio3" /><label for="functionradio3"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
      <input type="text" class="medium" id="functioninput3" name="functioninput3" onkeyup="handleKeyUp3(200);" value="" /><br />
      <select onclick="handleSelectClick3();" id="functionselect3" name="functionselect3" size="10" style="width: 220px;">
        <%
          MyCommon.QueryStr = "Select SVProgramID,Name,ExtProgramID from StoredValuePrograms where Deleted=0 order by Name;"
          rst = MyCommon.LRT_Select
          If bRequireExternalIds Then
            For Each row In rst.Rows
              If Not Long.TryParse(MyCommon.NZ(row.Item("ExtProgramID"), ""), lExternalId) Then lExternalId = 0
              If lExternalId = 0 Then
                If MyExport Is Nothing Then
                  MyExport = New Copient.ExportXml(MyCommon)
                End If
                lExternalId = MyExport.ExportOfferCrmGetExternalId(6, lPromoVarId)
                MyCommon.QueryStr = "update StoredValuePrograms with (RowLock) set ExtProgramID='" & lExternalId & "' where SVProgramID=" & row.Item("SVProgramID") & ";"
                MyCommon.LRT_Execute()
              End If
              Send("<option value=""" & lExternalId & """>" & row.Item("Name") & "</option>")
            Next
          Else
            For Each row In rst.Rows
              Send("<option value=""" & row.Item("SVProgramID") & """>" & row.Item("Name") & "</option>")
            Next
          End If

        %>
      </select>
      <br />
      <br />
      <input type="button" id="btnClose3" name="btnClose3" value="<% Sendb(Copient.PhraseLib.Lookup("term.close", LanguageID))%>" onclick="javascript:showDialogSpan(false, 3);" />
      </div>
      <div id="lifetimeselector">
        <div id="ltTag">
        </div>
        <br />
        <center>
          <input type="hidden" name="ltTagName" id="ltTagName" value="" />
          <b><% Sendb(Copient.PhraseLib.Lookup("offer-rew.selectpoints", LanguageID))%>:</b><br />
          <input type="radio" id="functionradio4a" name="functionradio4" checked="checked" /><label for="functionradio4"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
          <input type="radio" id="functionradio4b" name="functionradio4" /><label for="functionradio4"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
          <input class="medium" onkeyup="handleKeyUp4(200);" id="functioninput4" name="functioninput4" type="text" value="" /><br />
          <select onclick="handleSelectClick4();" id="functionselect4" name="functionselect4" size="10" style="width: 220px;">
            <%
              If sLifetimePointsId <> "" Then
                MyCommon.QueryStr = "select distinct PromoVarID, ProgramName from PointsPrograms " & _
                                    "where Deleted=0 and ProgramID in (" & sLifetimePointsId & ") " & _
                                    "order by ProgramName;"
                rst = MyCommon.LRT_Select
                If bRequireExternalIds Then
                  For Each row In rst.Rows
                    lPromoVarId = MyCommon.NZ(row.Item("PromoVarID"), 0)
                    MyCommon.QueryStr = "select ExternalID from PromoVariables with (NoLock) where PromoVarID=" & lPromoVarId & ";"
                    dt = MyCommon.LXS_Select
                    If dt.Rows.Count > 0 Then
                      If Not Long.TryParse(MyCommon.NZ(dt.Rows(0).Item(0), ""), lExternalId) Then lExternalId = 0
                      If lExternalId = 0 Then
                        If MyExport Is Nothing Then
                          MyExport = New Copient.ExportXml(MyCommon)
                        End If
                        lExternalId = MyExport.ExportOfferCrmGetExternalId(13, lPromoVarId)
                        MyCommon.QueryStr = "update PromoVariables with (RowLock) set LastUpdate = GetDate(), ExternalID='" & lExternalId & "' where PromoVarID=" & lPromoVarId & ";"
                        MyCommon.LXS_Execute()
                      End If
                      Send("<option value=" & lExternalId & ">" & row.Item("ProgramName") & "</option>")
                    End If
                  Next
                Else
                  For Each row In rst.Rows
                    Send("<option value=" & row.Item("PromoVarID") & ">" & row.Item("ProgramName") & "</option>")
                  Next
                End If
              End If
            %>
          </select>
          <br />
          <br />
          <input type="button" id="btnClose4" name="btnClose4" value="<% Sendb(Copient.PhraseLib.Lookup("term.close", LanguageID))%>" onclick="javascript:showDialogSpan(false, 8);" />
        </center>
      </div>
    </div>
  </div>
</form>

<script type="text/javascript">
setlimitsection(false);
setperiodsection(false);
hideprintonback(true);
  
   
   if ($("#pgroup-select option").val() != null){
     document.getElementById("HdSelectedProductGroup").value =$("#pgroup-select option").val();    
   }
   else{
     document.getElementById("HdSelectedProductGroup").value =$("#pgroup-select option").val();  
   }

  document.getElementById("HdExcludedProdGroupID").value = $("#pgroup-exclude option").val();
  setSearchTypeRadiobutton();
<% If (CloseAfterSave) Then %>
    window.close();
<% Else %>
    xmlhttpPost("PrintedMessageFeeds.aspx");
<% End If %>



</script>

<%
done:
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  Send_BodyEnd("mainform", "tier0")
  Logix = Nothing
  MyCommon = Nothing
%>
