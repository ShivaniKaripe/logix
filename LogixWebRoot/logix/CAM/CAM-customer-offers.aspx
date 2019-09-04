<%@ Page Language="vb" Debug="true" CodeFile="../LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%
  ' *****************************************************************************
  ' * FILENAME: CAM-customer-offers.aspx 
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
  Dim MyCryptLib As New Copient.CryptLib
  Dim MyLookup As New Copient.CustomerLookup
  Dim Logix As New Copient.LogixInc
  Dim rstResults As DataTable = Nothing
  Dim rst As DataTable
  Dim row As DataRow
  Dim rst2 As DataTable
  Dim row2 As DataRow
  Dim rstOffers As DataTable = Nothing
  Dim rowCount As Integer
  Dim CurrentOffers As String = ""
  Dim CustomerPK As Long
  Dim CardPK As Long = 0
  Dim ExtCardID As String = ""
  Dim FullName As String = ""
  Dim ExtCustomerID As String = ""
  Dim TotalRedeemCt As Integer = 0
  Dim TotalRedeemAmt As Double = 0.0
  Dim CustExtID As String = ""
  Dim i As Integer = 0
  Dim j As Integer = 0
  Dim n As Integer = 0
  Dim r As Integer = 0
  Dim offerCt As Integer = 0
  Dim transCt As Integer = 0
  Dim ElemStyle As String = ""
  Dim OfferName As String = ""
  Dim XID As String = ""
  Dim IsPtsOffer As Boolean = False
  Dim IsSVOffer As Boolean = False
  Dim IsAccumOffer As Boolean = False
  Dim UnknownPhrase As String = ""
  Dim DisabledPtsAdj As String = ""
  Dim DisabledSVAdj As String = ""
  Dim DisabledAccumAdj As String = ""
  Dim CgGroupIDs As String = ""
  Dim SortText As String = "Name"
  Dim SortDirection As String = "ASC"
  Dim OfferCMWhere As String = ""
  Dim OfferCPEWhere As String = ""
  Dim OfferTerms As String = ""
  Dim IsHouseholdID As Boolean = False
  Dim HHPK As Integer = 0
  Dim HouseholdID As String = ""
  Dim HHCustIdList As New ArrayList(5)
  Dim PaddedExtID As String = StrDup(25, "0")
  Dim PaddedProdExtID As String = ""
  Dim CustExtIdList As String = ""
  Dim Shaded As String = " class=""shaded"""
  Dim HasSearchResults As Boolean = False
  Dim FullAddress As String = ""
  Dim CustomerTypeID As Integer = 0
  Dim Employee As Integer
  Dim TestCard As Boolean = False
  Dim OfferID As Integer = 0
  Dim PrimaryExtID As String = ""
  Dim ClientUserID1 As String = ""
  Dim IDLength As Integer = 0
  Dim CustomerGroupIDs As String() = Nothing
  Dim loopCtr As Integer = 0
  Dim searchterms As String = ""
  Dim restrictLinks As Boolean = False
  Dim PointsIDBuf As New StringBuilder()
  Dim PointsNameBuf As New StringBuilder()
  Dim OffersList As String = ""
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim HistoryText As String = ""
  Dim PrevOfferID As Long = 0
  
  Dim Favorite As Integer
  Dim CgXml As String = ""
  Dim reader As SqlDataReader = Nothing
  Dim dtAddOffers As DataTable = Nothing
  Dim dtAssigned As DataTable = Nothing
  Dim sortedRows() As DataRow = Nothing
  Dim filteredRows() as DataRow = nothing
  Dim ColValues(11) As Object
  Dim RSCount As Integer = 1
  Dim OfferStatusCode As Copient.LogixInc.STATUS_FLAGS
  Dim OfferStatus As String = ""
  Dim StatusTable As New Hashtable(200)
  Dim OfferInQuo As Integer = 0
  Dim AllCAMCardholdersID As Long = 0
  Dim OfferSearchDate As Date = Nothing
  Dim SearchDateSet As Boolean = False
  Dim InvalidSearchDate As Boolean = False
  Dim OfferTable As Hashtable
  Dim ProdExtIDLen As Integer = 0
  Dim RemovedCustomerGroups As String = ""
  Dim OfferStatusDate As Date
  Dim ExcludeExpired As Boolean = False
  Dim Filter As String = ""
  Dim Fields As New Copient.CommonInc.ActivityLogFields
  Dim AssocLinks(-1) As Copient.CommonInc.ActivityLink
  Dim SessionID As String = ""
  Dim OfferStart As Integer = 0
  Dim index As Integer = 0
  Dim PointsTable As New Hashtable()
  Dim pointRow As DataRow
  Dim pointdt As DataTable
  Dim PointsProgram As Integer = 0
  Dim HasPoints As Boolean = False
  Dim pointTemp, amountTemp As Integer
  Dim PaddedTrigger As String = ""
  Dim ShowCreditOnly As Boolean = False
  Dim BalanceOffers As New Hashtable()
  
  ' default urls for links from this page
  Dim URLCAMOfferSum As String = "CAM-offer-sum.aspx"
  Dim URLcgroupedit As String = "/logix/cgroup-edit.aspx"
  Dim URLpointedit As String = "/logix/point-edit.aspx"
  Dim URLtrackBack As String = ""
  Dim inCardNumber As String = ""
  ' tack on the customercare remote links if needed
  Dim extraLink As String = ""
  
  Dim UserRoleIDs() As Integer
  Dim RoleMatch As Boolean = False
  Dim x As Integer = 0
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "CAM-customer-offers.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  MyCommon.Open_LogixWH()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  If Session("SessionID") IsNot Nothing AndAlso Session("SessionID").ToString.Trim <> "" Then
    SessionID = Session("SessionID").ToString.Trim
    CopientNotes &= "SessionID: " & SessionID
  End If
    'get the paddign length of upc code
    MyCommon.QueryStr = "select paddingLength from producttypes where producttypeid=1"
    rst = MyCommon.LRT_Select
    If rst IsNot Nothing Then
        ProdExtIDLen = Convert.ToInt32(rst.Rows(0).Item("paddingLength"))
    End If
    rst = Nothing 
  ' lets check the logged in user and see if they are to be restricted to this page
  MyCommon.QueryStr = "select AU.StartPageID,AUSP.PageName ,AUSP.prestrict from AdminUsers as AU with (NoLock) " & _
                      "inner join AdminUserStartPages as AUSP with (NoLock) on AU.StartPageID=AUSP.StartPageID " & _
                      "where AU.AdminUserID=" & AdminUserID
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    If (rst.Rows(0).Item("prestrict") = True) Then
      ' ok we got in here then we need to restrict the user from seeing any other pages
      restrictLinks = True
    End If
  End If
  
  'Populate the UserRoleIDs array
  MyCommon.QueryStr = "select RoleID from AdminUserRoles with (NoLock) where AdminUserID=" & AdminUserID & ";"
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    UserRoleIDs = New Integer(rst.Rows.Count - 1) {}
    For Each row In rst.Rows
      UserRoleIDs(x) = MyCommon.NZ(row.Item("RoleID"), 0)
      x += 1
    Next
    x = 0
  End If
  
  CustomerPK = MyCommon.Extract_Val(Request.QueryString("CustPK"))
  If CustomerPK = 0 Then
    CustomerPK = MyCommon.Extract_Val(Request.QueryString("CustomerPK"))
  End If
  
  CardPK = MyCommon.Extract_Val(Request.QueryString("CardPK"))
  If CardPK = 0 Then
    CardPK = MyLookup.FindCardPK(CustomerPK, 2)
  End If
  
  If CardPK > 0 Then
    ExtCardID = MyLookup.FindExtCardIDFromCardPK(CardPK)
  End If
  
  If (CustomerPK > 0) Then
    MyCommon.QueryStr = "select CustomerTypeID from Customers with (NoLock) where CustomerPK=" & MyCommon.Extract_Val(Request.QueryString("CustPK"))
    rst = MyCommon.LXS_Select
    If (rst.Rows.Count > 0) Then
      CustomerTypeID = rst.Rows(0).Item("CustomerTypeID")
    End If
  End If
  
  If (Request.QueryString("Favorite") = "0" OrElse Request.QueryString("Favorite") = "FALSE") Then
    Favorite = 0
  ElseIf (Request.QueryString("Favorite") = "1" OrElse Request.QueryString("Favorite") = "TRUE") Then
    Favorite = 1
    SortText = "Priority"
    SortDirection = "ASC"
    '!REMOVE
  ElseIf (Request.QueryString("Favorite") = "2") Then
    Favorite = 0
    ExcludeExpired = True
  ElseIf Request.QueryString("Favorite") = "" Then
    MyCommon.QueryStr = "select OfferID from AdminUserOffers where AdminUserID=" & AdminUserID & ";"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count = 0 Then
      Favorite = 0
    Else
      Favorite = 1
      SortText = "Priority"
      SortDirection = "ASC"
    End If
  ElseIf (Request.QueryString("Favorite") = "5") Then
    Favorite = 0
    ExcludeExpired = False
    ShowCreditOnly = True
  End If
  
  If (CustomerPK = 0) Then
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "/logix/customer-inquiry.aspx")
  End If
  
  ' validate any entered offer date search term
  If (Request.QueryString("offerdate") <> "" AndAlso Request.QueryString("offerdate").Trim <> "") Then
    If Not (Date.TryParse(Request.QueryString("offerdate"), OfferSearchDate)) Then
      infoMessage &= Copient.PhraseLib.Lookup("customer-inquiry.invalid-date", LanguageID) & "<br />"
      InvalidSearchDate = True
    End If
  End If
  
  ' validate the entered offer store term exists as a valid store
  If (Request.QueryString("offerstore") <> "" AndAlso Request.QueryString("offerstore").Trim <> "") Then
    MyCommon.QueryStr = "select LocationID from Locations with (NoLock) where ExtLocationCode='" & Request.QueryString("offerstore").Trim & "'"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count = 0) Then
      infoMessage &= ControlChars.Cr & Copient.PhraseLib.Lookup("customer-inquiry.store-not-found", LanguageID) & "<br />"
    End If
  End If
  
  ' validate the entered offer PLU term exists as a valid product
  If (Request.QueryString("offerUPC") <> "" AndAlso Request.QueryString("offerUPC").Trim <> "") Then
    ' pad the PLU (if necessary)
    'Integer.TryParse(MyCommon.Fetch_SystemOption(52), ProdExtIDLen)
    If ProdExtIDLen > 0 Then
      PaddedProdExtID = Request.QueryString("offerUPC").Trim.PadLeft(ProdExtIDLen, "0")
    Else
      PaddedProdExtID = Request.QueryString("offerUPC").Trim
    End If
    MyCommon.QueryStr = "select ProductID from Products with (NoLock) where ProductTypeID=1 and ExtProductID = '" & PaddedProdExtID & "'"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count = 0) Then
      infoMessage &= ControlChars.Cr & Copient.PhraseLib.Lookup("customer-inquiry.product-not-found", LanguageID) & "<br />"
    End If
  End If
  
  ' validate the entered offer Trigger code term exists as a valid code
  If (Request.QueryString("offerTrigger") <> "" AndAlso Request.QueryString("offerTrigger").Trim <> "") Then
    ' pad the Trigger code (if necessary)
    Integer.TryParse(MyCommon.Fetch_SystemOption(52), ProdExtIDLen)
    If ProdExtIDLen > 0 Then
      PaddedTrigger = Request.QueryString("offerTrigger").Trim.PadLeft(ProdExtIDLen, "0")
    Else
      PaddedTrigger = Request.QueryString("offerTrigger").Trim
    End If
    MyCommon.QueryStr = "select RewardOptionID from CPE_IncentivePLUs with (NoLock) where PLU = '" & PaddedTrigger & "'"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count = 0) Then
      infoMessage &= ControlChars.Cr & Copient.PhraseLib.Lookup("customer-inquiry.product-not-found", LanguageID) & "<br />"
    End If
  End If
  
  'Use CustomerPK to find all the points groups that they have points in. Put results in a hashtable
  MyCommon.QueryStr = "select ProgramID, Amount from Points where CustomerPK=" & CustomerPK & " and Amount > 0;"
  pointdt = MyCommon.LXS_Select()
  If pointdt.Rows.Count > 0 Then
    For Each pointRow In pointdt.Rows
      pointTemp = MyCommon.Extract_Val(MyCommon.NZ(pointRow.Item("ProgramID"), 0))
      amountTemp = MyCommon.Extract_Val(MyCommon.NZ(pointRow.Item("Amount"), 0))
      PointsTable.Add(pointTemp, amountTemp)
    Next
  End If
  
  If CardPK > 0 Then
    Send_HeadBegin("term.customer", "term.offers", MyCommon.Extract_Val(ExtCardID))
  Else
    Send_HeadBegin("term.customer", "term.offers")
  End If
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  If restrictLinks Then
    Send("<link rel=""stylesheet"" href=""/css/logix-restricted.css"" type=""text/css"" media=""screen"" />")
  End If
%>
<style type="text/css">
#functionselect {
  height: 166px;
  }
* html #functionselect {
  height: 175px;
  }
</style>
<%
  Send_Scripts()
  Send_HeadEnd()
  
  ' Before anything else, check if we're supposed to remove someone from an offer
  If (Request.QueryString("RemoveFromOffer") <> "") Then
    ' Remove customer from a group; incoming will look like CustomerGroupID=4&CustomerPK=46
    
    ' Determine if customer is a household or cardholder
    If (MyCommon.Extract_Val(Request.QueryString("CustomerPK")) > 0) Then
      CustomerPK = MyCommon.Extract_Val(Request.QueryString("CustomerPK"))
      MyCommon.QueryStr = "select CustomerTypeID from Customers with (NoLock) where CustomerPK=" & CustomerPK
      rst = MyCommon.LXS_Select
      If (rst.Rows.Count > 0) Then
        CustomerTypeID = rst.Rows(0).Item("CustomerTypeID")
      End If
    End If
    
    MyCommon.QueryStr = "select distinct ICG.CustomerGroupID, CG.EditControlTypeID, CG.RoleID from CPE_IncentiveCustomerGroups as ICG with (NoLock) " & _
                        "inner join CPE_RewardOptions as RO on RO.RewardOptionID=ICG.RewardOptionID " & _
                        "inner join CustomerGroups as CG with (NoLock) on CG.CustomerGroupID=ICG.CustomerGroupID " & _
                        "where RO.IncentiveID=" & MyCommon.Extract_Val(Request.QueryString("OfferID")) & " and RO.Deleted=0 and ICG.Deleted=0;"
    rst = MyCommon.LRT_Select
    For Each row In rst.Rows
      ReDim AssocLinks(rst.Rows.Count - 1)
      i = 0
      If (MyCommon.NZ(row.Item("EditControlTypeID"), 0) = 3) Then 'removal is limited per role, so loop through the user's roles and see if there's a match to the customer group's RoleID 
        For x = 0 To UserRoleIDs.Length - 1
          If UserRoleIDs(x) = MyCommon.NZ(row.Item("RoleID"), 0) Then
            RoleMatch = True
          End If
        Next
      End If
      If ExtCardID IsNot Nothing AndAlso ExtCardID.Trim <> "" AndAlso MyCommon.NZ(row.Item("CustomerGroupID"), 0) > 0 Then
        If (MyCommon.NZ(row.Item("EditControlTypeID"), 0) = 0 AndAlso Logix.UserRoles.RemoveCustomerFromOffers) OrElse (MyCommon.NZ(row.Item("EditControlTypeID"), 0) = 1) OrElse (MyCommon.NZ(row.Item("EditControlTypeID"), 0) = 3 AndAlso RoleMatch) Then
          MyCommon.QueryStr = "dbo.pt_GroupMembership_Delete_ByPK"
          MyCommon.Open_LXSsp()
          MyCommon.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.NVarChar, 26).Value = CustomerPK
          MyCommon.LXSsp.Parameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = MyCommon.NZ(row.Item("CustomerGroupID"), 0)
          MyCommon.LXSsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
          MyCommon.LXSsp.ExecuteNonQuery()
          'outputStatus = MyCommon.LXSsp.Parameters("@Status").Value
          MyCommon.Close_LXSsp()
          If RemovedCustomerGroups <> "" Then RemovedCustomerGroups &= ","
          RemovedCustomerGroups &= MyCommon.NZ(row.Item("CustomerGroupID"), 0)
          AssocLinks(i).LinkID = MyCommon.NZ(row.Item("CustomerGroupID"), 0)
          AssocLinks(i).LinkTypeID = 2
          i += 1
        Else
          infoMessage = Copient.PhraseLib.Detokenize("customer-offers.CustomerCouldNotBeRemoved", LanguageID, Request.QueryString("OfferID"))
        End If
      End If
    Next
    
    ' Determine offers associated with the customer group to add to the history
    OffersList = ""
    MyCommon.QueryStr = "select distinct Name,O.OfferID from Offers as O with (NoLock) " & _
                        "left join OfferConditions as OC with (NoLock) on OC.OfferID=O.OfferID " & _
                        "where(O.IsTemplate = 0 And OC.ConditionTypeID = 1 And OC.LinkID = " & Request.QueryString("CustomerGroupID") & " Or OC.ExcludedID = " & Request.QueryString("CustomerGroupID") & ")" & _
                        " union " & _
                        "select distinct I.IncentiveName,I.IncentiveID as OfferID " & _
                        "from CPE_IncentiveCustomerGroups ICG with (NoLock) " & _
                        "inner join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID = ICG.RewardOptionID " & _
                        "inner join CPE_Incentives I with (NoLock) on I.IncentiveID = RO.IncentiveID " & _
                        "where(ICG.Deleted = 0 And RO.Deleted = 0 And i.Deleted = 0 And i.IsTemplate = 0 And ICG.CustomerGroupID = " & Request.QueryString("CustomerGroupID") & ") " & _
                        "order by OfferID ASC;"
    rst2 = MyCommon.LRT_Select
    If rst2.Rows.Count > 0 Then
      OfferStart = AssocLinks.Length
      ReDim Preserve AssocLinks(OfferStart + rst2.Rows.Count - 1)
      If rst2.Rows.Count = 1 Then
        OffersList = rst2.Rows(0).Item("OfferID")
        AssocLinks(OfferStart).LinkID = MyCommon.NZ(rst2.Rows(0).Item("OfferID"), 0)
        AssocLinks(OfferStart).LinkTypeID = 1 ' Offer link type
        AssocLinks(OfferStart).Selected = (AssocLinks(OfferStart).LinkID = MyCommon.Extract_Val(Request.QueryString("OfferID")))
      ElseIf rst2.Rows.Count > 1 Then
        i = 1
        For Each row In rst2.Rows
          If i = 1 Then
            OffersList = row.Item("OfferID")
          Else
            OffersList = OffersList & ", " & row.Item("OfferID")
          End If
          index = OfferStart + i - 1
          AssocLinks(index).LinkID = MyCommon.NZ(row.Item("OfferID"), 0)
          AssocLinks(index).LinkTypeID = 1 ' Offer link type
          AssocLinks(index).Selected = (AssocLinks(index).LinkID = MyCommon.Extract_Val(Request.QueryString("OfferID")))

          i = i + 1
        Next
      End If
      HistoryText = Copient.PhraseLib.Lookup("history.customer-remove-offer", LanguageID) & " #" & RemovedCustomerGroups & " (" & OffersList & ")"
      If Len(HistoryText) > 245 Then
        HistoryText = Left(HistoryText, 245)
        i = HistoryText.LastIndexOf(",")
        If i > -1 Then
          HistoryText = HistoryText.Substring(0, i) & "...)"
        End If
      End If
      If Len(OffersList) > 245 Then
        OffersList = Left(OffersList, 245)
        i = OffersList.LastIndexOf(",")
        If i > -1 Then
          OffersList = OffersList.Substring(0, i) & " ..."
        End If
      End If
      'MyCommon.Activity_Log2(25, 16, CustomerPK, AdminUserID, HistoryText, MyCommon.Extract_Val(Request.QueryString("OfferID")))
    Else
      HistoryText = Copient.PhraseLib.Lookup("history.customer-remove-offer", LanguageID) & " #" & Request.QueryString("CustomerGroupID")
      'MyCommon.Activity_Log2(25, 16, CustomerPK, AdminUserID, Copient.PhraseLib.Lookup("history.customer-remove-offer", LanguageID) & " #" & RemovedCustomerGroups, MyCommon.Extract_Val(Request.QueryString("OfferID")))
    End If
    
    'Log the addition of the offer and any associated offers
    Fields.ActivityTypeID = 25
    Fields.ActivitySubTypeID = 16
    Fields.LinkID = CustomerPK
    Fields.AdminUserID = AdminUserID
    Fields.Description = HistoryText
    Fields.LinkID2 = MyCommon.Extract_Val(Request.QueryString("CustomerGroupID"))
    Fields.AssociatedLinks = AssocLinks
    Fields.SessionID = SessionID
    MyCommon.Activity_Log3(Fields)
    
  ElseIf (Request.QueryString("AddFavorite") <> "") Then
    ' User clicked to mark an offer as a favorite...
    OfferInQuo = Request.QueryString("AddFavorite")
    MyCommon.QueryStr = "select AdminUserID, OfferID, Priority, FavoredBy, FavoredDate from AdminUserOffers " & _
                        "where AdminUserID=" & AdminUserID & " and OfferID=" & OfferInQuo & ";"
    rst2 = MyCommon.LRT_Select
    If rst2.Rows.Count = 0 Then
      MyCommon.QueryStr = "insert into AdminUserOffers with (RowLock) (AdminUserID, OfferID, Priority, FavoredBy, FavoredDate) " & _
                          "values (" & AdminUserID & ", " & OfferInQuo & ", 1, " & AdminUserID & ", '" & Now() & "');"
      MyCommon.LRT_Execute()
    End If
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "CAM-customer-offers.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&CardPK=" & CardPK, "") & "&Favorite=" & Request.QueryString("Favorite") & "&SortText=" & Request.QueryString("SortText") & "&SortDirection=" & Request.QueryString("SortDirection") & _
                       "&offerterms=" & Request.QueryString("offerterms") & "&offerdate=" & Request.QueryString("offerdate") & "&offerstore=" & Request.QueryString("offerstore") & "&offerUPC=" & Request.QueryString("offerUPC"))
  ElseIf (Request.QueryString("DeleteFavorite") <> "") Then
    ' User clicked unmark an offer as a favorite...
    OfferInQuo = Request.QueryString("DeleteFavorite")
    MyCommon.QueryStr = "select AdminUserID, OfferID, Priority, FavoredBy, FavoredDate from AdminUserOffers " & _
                        "where AdminUserID=" & AdminUserID & " and OfferID=" & OfferInQuo & ";"
    rst2 = MyCommon.LRT_Select
    If rst2.Rows.Count > 0 Then
      MyCommon.QueryStr = "delete from AdminUserOffers with (RowLock) where AdminUserID=" & AdminUserID & " and OfferID=" & OfferInQuo & ";"
      MyCommon.LRT_Execute()
    End If
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "CAM-customer-offers.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&CardPK=" & CardPK, "") & "&Favorite=" & Request.QueryString("Favorite") & "&SortText=" & Request.QueryString("SortText") & "&SortDirection=" & Request.QueryString("SortDirection") & _
                       "&offerterms=" & Request.QueryString("offerterms") & "&offerdate=" & Request.QueryString("offerdate") & "&offerstore=" & Request.QueryString("offerstore") & "&offerUPC=" & Request.QueryString("offerUPC"))
  End If
  
  ' special handling for customer inquery direct link in 
  If (restrictLinks) Then
    URLCAMOfferSum = ""
    URLcgroupedit = ""
    URLpointedit = ""
  End If
  
  ' set session to nothing just to be sure
  Session.Add("extraLink", "")
  
  If (Request.QueryString("mode") = "summary") Then
    URLtrackBack = Request.QueryString("exiturl")
    inCardNumber = Request.QueryString("cardnumber")
    extraLink = "&mode=summary&exiturl=" & URLtrackBack & "&cardnumber=" & inCardNumber
    Session.Add("extraLink", extraLink)
  End If
  
  ' hack for popups check session for extralink
  If (Session("extraLink").ToString = "") Then
    extraLink = Session("extraLink")
  End If
  
  If (MyCommon.Extract_Val(Request.QueryString("CustPK")) > 0 Or (Request.QueryString("searchterms") <> "" And _
      (Request.QueryString("Search") <> "" Or Request.QueryString("searchPressed") <> "")) Or _
      inCardNumber <> "" _
      ) Then
    ' someone wants to search for a customer.  First lets get their primary key from our database
    If (MyCommon.Extract_Val(Request.QueryString("CustPK")) > 0 Or (MyCommon.Extract_Val(Request.QueryString("CustomerPK")) > 0)) Then
      If (MyCommon.Extract_Val(Request.QueryString("CustPK")) > 0) Then
        CustomerPK = MyCommon.Extract_Val(Request.QueryString("CustPK"))
      ElseIf (MyCommon.Extract_Val(Request.QueryString("CustomerPK")) > 0) Then
        CustomerPK = MyCommon.Extract_Val(Request.QueryString("CustomerPK"))
      End If
      MyCommon.QueryStr = "select C.CustomerPK, C.CustomerTypeID, C.HHPK, C.FirstName, C.MiddleName, C.LastName, " & _
                          "C.TestCard, CE.Address, CE.City, CE.State, CE.Zip, CE.PhoneAsEntered as Phone, CE.email " & _
                          "from Customers C with (NoLock) left join CustomerExt CE with (NoLock) on CE.CustomerPK=C.CustomerPK " & _
                          "left join Customers C2 with (NoLock) on C2.CustomerPK=C.HHPK " & _
                          "where C.CustomerPK=" & CustomerPK & ";"
    Else
      ' IF the page was called from an outside application set ClientUserID1 to the outside passed in value
      If (inCardNumber <> "" And Request.QueryString("mode") = "summary") Then
        searchterms = Request.QueryString("searchterms")
        MyCommon.QueryStr = "select C.CustomerPK, C.CustomerTypeID, C.HHPK, C.FirstName, C.MiddleName, C.LastName, " & _
                            "C.TestCard, CE.Address, CE.City, CE.State, CE.Zip, CE.PhoneAsEntered as Phone, CE.email " & _
                            "from Customers C with (NoLock) left join CustomerExt CE with (NoLock) on CE.CustomerPK = C.CustomerPK " & _
                            "left join Customers C2 with (NoLock) on C2.CustomerPK=C.HHPK " & _
                            "where C.PrimaryExtID='" & MyCommon.Parse_Quotes(ClientUserID1) & "';"
      End If
    End If
    rstResults = MyCommon.LXS_Select
    
    If (rstResults.Rows.Count = 1) Then
      ' ok we found a primary key for the external id provided
      CustomerPK = rstResults.Rows(0).Item("CustomerPK")
      IsHouseholdID = MyCommon.NZ(rstResults.Rows(0).Item("CustomerTypeID"), 0) = 1
      TestCard = MyCommon.NZ(rstResults.Rows(0).Item("TestCard"), False)
    Else
      infoMessage = "" & Copient.PhraseLib.Lookup("customer.notfound", LanguageID) & ""
      infoMessage = infoMessage & " <a href=""CAM-customer-offers.aspx?mode=add&Search=Search" & extraLink & "&searchterms=" & Request.QueryString("searchterms") & """>" & Copient.PhraseLib.Lookup("term.add", LanguageID) & "</a>"
    End If
    
  End If
  
  UnknownPhrase = Copient.PhraseLib.Lookup("term.unknown", LanguageID)
  
  ' sorting for the offers 
  If (Request.QueryString("SortText") <> "") Then
    SortText = Request.QueryString("SortText")
  End If
  If (Request.QueryString("SortDirection") = "ASC") Then
    SortDirection = "DESC"
  ElseIf (Request.QueryString("SortDirection") = "DESC") Then
    SortDirection = "ASC"
  Else
    SortDirection = "ASC"
  End If
  
  Send_BodyBegin(1)
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  If (Not restrictLinks) Then
    Send_Tabs(Logix, 3)
    If CardPK > 0 Then
      Send_Subtabs(Logix, 33, 4, LanguageID, CustomerPK, , CardPK)
    Else
      Send_Subtabs(Logix, 33, 4, LanguageID, CustomerPK)
    End If
  Else
    If CardPK > 0 Then
      Send_Subtabs(Logix, 94, 4, LanguageID, CustomerPK, extraLink, CardPK)
    Else
      Send_Subtabs(Logix, 94, 4, LanguageID, CustomerPK, extraLink)
    End If
  End If
  If (Logix.UserRoles.AccessCustomerInquiry = False) Then
    Send_Denied(1, "perm.customer-access")
    GoTo done
  End If
%>
<script type="text/javascript">
<!--
var lastStartPos = -1;
var currentStart = -1;
var sizeOfData = -1;
var OFFER_ROWS_SHOWN = 10;

function submitenter(myfield,e) {
  var keycode;
  var submitThing;
  
  if (window.event) keycode = window.event.keyCode;
  else if (e) keycode = e.which;
  else return true;
  
  if (keycode == 13) {
    if (document.mainform.searchterms != null && document.mainform.searchterms.value != "") {
      submitThing = document.getElementById("searchPressed");
      submitThing.value = "Search";
      searchNew();
      myfield.form.submit();
      return false;
    } else {
      return false;
    }
  }
  else
    return true;
}

function addOffer(custPK, cardPK, offerID) {
  document.getElementById('AddPromoLink').href = '/logix/XMLFeeds.aspx?AddOffer=' + offerID + '&CustPK=' + custPK + '&CardPK=' + cardPK + '&AdminUserID=<%Sendb(AdminUserID)%>&height=300&width=300&LanguageID=<%Sendb(LanguageID)%>';
  
  var fireOnThis = document.getElementById('AddPromoLink');
  if( document.createEvent ) {
    var evObj = document.createEvent('MouseEvents');
    evObj.initEvent( 'click', true, false );
    fireOnThis.dispatchEvent(evObj);
  } else if( document.createEventObject ) {
    fireOnThis.fireEvent('onclick');
  }
}

function showOfferPage(position) {
  //var startPos = ((page-1) * OFFER_ROWS_SHOWN) + 1;
  var trElem = null;
  var transrows = getElementsByClassName("transrow");
  var imgTags = null;

  // hide all open transrow-class elements
  for (i = 0; i < transrows.length; i++) {
    transrows[i].style.display = 'none';
  }
  
  // hide last page shown
  if (lastStartPos > -1) {
    for (var i=lastStartPos; i < lastStartPos + OFFER_ROWS_SHOWN; i++) {
      trElem = document.getElementById("trOffer" + i);
            if (trElem != null) {
              trElem.style.display = 'none';
              
              // reset expand/contract icon for the run back to the plus icon
              imgTags = trElem.getElementsByTagName("IMG");
              if (imgTags != null && imgTags.length > 0) {
                imgTags[0].src = '/images/plus.png';
              }
            }
        }
    }
  
  // show current page
  for (var i=position; i < position + OFFER_ROWS_SHOWN; i++) {
    trElem = document.getElementById("trOffer" + i);
    if (trElem != null) trElem.style.display = '';
  }
  
  lastStartPos = position;
  updateOfferRecStatus();
  handleOfferButtons();    
}

function showOfferNextPage() {
  if (lastStartPos == -1) {
    showOfferPage(1);
  } else {
    showOfferPage(lastStartPos + OFFER_ROWS_SHOWN);
  }
}

function showOfferPrevPage() {
  if (lastStartPos <= OFFER_ROWS_SHOWN) {
    showOfferPage(1);
  } else {
    showOfferPage(lastStartPos - OFFER_ROWS_SHOWN);
  }
}

function showOfferFirstPage() {
  showOfferPage(1);
}

function showOfferLastPage() {
  var elemRecCt = document.getElementById("offerTableRowCt");
  var recCt = 0;
  var start = 1;
  
  if (elemRecCt != null) {
    recCt = elemRecCt.value;
    start = (Math.floor(recCt / OFFER_ROWS_SHOWN) * OFFER_ROWS_SHOWN) + 1;
    if (recCt % OFFER_ROWS_SHOWN == 0 && recCt > OFFER_ROWS_SHOWN) {
      start = start - OFFER_ROWS_SHOWN;
    }
    showOfferPage(start);
  }
}

function updateOfferRecStatus() {
  var elemStart = document.getElementById("startPos");
  var elemEnd = document.getElementById("endPos");
  var elemRecCt = document.getElementById("offerTableRowCt");
  var recCt = 0;
  
  if (elemStart != null) elemStart.innerHTML = lastStartPos;
  if (elemRecCt != null) {
    recCt = elemRecCt.value;
    if (elemEnd!=null && recCt < (lastStartPos + OFFER_ROWS_SHOWN - 1)) {
      elemEnd.innerHTML = recCt;
    } else if (elemEnd!=null) {
      elemEnd.innerHTML = (lastStartPos + OFFER_ROWS_SHOWN - 1);
    }
  }
}

function handleOfferButtons() {
  handleOfferPrevButton();
  handleOfferNextButton();
}

function handleOfferPrevButton() {
  var elemFirstOn = document.getElementById("first");
  var elemPrevOn = document.getElementById("previous");
  var elemFirstOff = document.getElementById("firstOff");
  var elemPrevOff = document.getElementById("previousOff");
  
  if (lastStartPos > 1) {
    if (elemFirstOn != null)  elemFirstOn.style.display = "";
    if (elemPrevOn != null)  elemPrevOn.style.display = "";
    if (elemFirstOff != null) elemFirstOff.style.display = "none";          
    if (elemPrevOff != null) elemPrevOff.style.display = "none";          
  } else {
    if (elemFirstOn != null)  elemFirstOn.style.display = "none";
    if (elemPrevOn != null)  elemPrevOn.style.display = "none";
    if (elemFirstOff != null) elemFirstOff.style.display = "";          
    if (elemPrevOff != null) elemPrevOff.style.display = "";          
  }
}

function handleOfferNextButton() {
  var elemLastOn = document.getElementById("last");
  var elemNextOn = document.getElementById("next");
  var elemLastOff = document.getElementById("lastOff");
  var elemNextOff = document.getElementById("nextOff");
  var elemRecCt = document.getElementById("offerTableRowCt");
  var recCt = (elemRecCt != null) ? elemRecCt.value : 0;
  
  if ((lastStartPos + OFFER_ROWS_SHOWN) <= recCt) {
    if (elemNextOn != null)  elemNextOn.style.display = "";
    if (elemLastOn != null)  elemLastOn.style.display = "";
    if (elemNextOff != null) elemNextOff.style.display = "none";          
    if (elemLastOff != null) elemLastOff.style.display = "none";  
  } else  {
    if (elemNextOn != null)  elemNextOn.style.display = "none";
    if (elemLastOn != null)  elemLastOn.style.display = "none";
    if (elemNextOff != null) elemNextOff.style.display = "";          
    if (elemLastOff != null) elemLastOff.style.display = ""; 
  }
}

// This is the function that refreshes the list after a keypress.
// The maximum number to show can be limited to improve performance with
// huge lists (1000s of entries).
// The function clears the list, and then does a linear search through the
// globally defined array and adds the matches back to the list.
function handleKeyUp(maxNumToShow) {
  var selectObj, textObj, functionListLength;
  var i,  numShown;
  var searchPattern;
  
  document.getElementById("functionselect").size = "20";
  
  // Set references to the form elements
  selectObj = document.forms['mainform'].functionselect;
  textObj = document.forms['mainform'].functioninput;
  
  // Remember the function list length for loop speedup
  functionListLength = functionlist.length;
  
  // Set the search pattern depending
  if(document.forms['mainform'].functionradio[0].checked == true) {
    searchPattern = "^"+textObj.value;
  } else {
    searchPattern = textObj.value;
  }
  searchPattern = cleanRegExpString(searchPattern);
  
  // Create a regular expression
  re = new RegExp(searchPattern,"gi");
  // Clear the options list
  selectObj.length = 0;
  
  // clear the description
  if (document.getElementById('detailsbox') != null) {
    document.getElementById('detailsbox').innerHTML = '<span class=\"grey\"><% Sendb(Copient.PhraseLib.Lookup("term.none", LanguageID))%><\/span>';
  }
  
  // Loop through the array and re-add matching options
  numShown = 0;
  for(i = 0; i < functionListLength; i++) {
    if(functionlist[i].search(re) != -1 && vallist[i] != "") {
      pointerlist[numShown] = i;
      selectObj[numShown] = new Option(functionlist[i],vallist[i]);
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
    if (document.getElementById('detailsbox') != null) {
      if (descs[pointerlist[0]] == '') {
        document.getElementById('detailsbox').innerHTML = '<span class=\"grey\"><% Sendb(Copient.PhraseLib.Lookup("term.none", LanguageID))%><\/span>';
      } else {
      document.getElementById('detailsbox').innerHTML = descs[pointerlist[0]];
      }
    }
  }
}

// this function gets the selected value and loads the appropriate
// php reference page in the display frame
// it can be modified to perform whatever action is needed, or nothing
function handleSelectClick() {
  selectObj = document.forms['mainform'].functionselect;
  textObj = document.forms['mainform'].functioninput;
  
  if (selectObj != null && selectObj.selectedIndex > -1) {
    selectedValue = selectObj.options[selectObj.selectedIndex].text;
    selectedValue = selectedValue.replace(/_/g, '-') ;
    var dashPos = selectedValue.indexOf('- ');
    
    if (dashPos > -1) {
      selectedValue = selectedValue.substring(dashPos + 2);
    }
    //selectedValue = document.getElementById("functionselect").value;
    if(selectedValue != "") {   
      textObj.value = selectedValue;
    }
  }
}

function searchOffers() {
  var elemOfferTerms = document.getElementById("offerterms");
  var elemOfferDate = document.getElementById("offerdate");
  var elemOfferStore = document.getElementById("offerstore");
  var elemofferUPC = document.getElementById("offerUPC");
  var elemofferTrigger = document.getElementById("offerTrigger");
  
  var offerTerms = '';
  var offerDate = '';
  var offerStore = '';
  var offerUPC = '';
  var offerTrigger = '';
  
  if (elemOfferTerms != null) { offerTerms = elemOfferTerms.value; }
  if (elemOfferDate != null) { offerDate = elemOfferDate.value; }
  if (elemOfferStore != null) { offerStore = elemOfferStore.value; }
  if (elemofferUPC != null) { offerUPC = elemofferUPC.value; }
  if (elemofferTrigger != null) { offerTrigger = elemofferTrigger.value; }
  
  <%
    Dim strTerms As String = Request.QueryString("searchterms")
    If (strTerms <> "") Then
        strTerms = strTerms.Replace("'", "\'")
        strTerms = strTerms.Replace("""", "\""")
    End If
   %>
  offerTerms = escape(offerTerms);
  offerTerms = offerTerms.replace('+', '%2B');
  var qryStr = 'CAM-customer-offers.aspx?search=Search&searchterms=<%Sendb(strTerms)%>&CustomerPK=<%Sendb(CustomerPK)%><%Sendb(IIf(CardPK > 0, "&CardPK=" & CardPK, ""))%>&Favorite=0&offerSearch=Search&offerterms=' + offerTerms + '&offerdate=' + offerDate + '&offerstore=' + offerStore + '&offerUPC=' + offerUPC + '&offerTrigger=' + offerTrigger + '#h01';
  document.location = qryStr;
}

function submitOfferSearch(e) {
  var key = e.which ? e.which : e.keyCode;
  
  if (key == 13) {
    if (e && e.preventDefault) {
      e.preventDefault(); // DOM style
      searchOffers();
    } else {
      e.keyCode = 9;
      searchOffers();
    }
    return false;
  }
  return true;
}

function searchNew() {
  var elem = document.getElementById("CustomerPK");
  
  if (elem != null) {
    elem.value = ""
  }
}

function handleKeyDown(e) {
  var key = e.which ? e.which : e.keyCode;
  var elem = document.getElementById('functionselect');
  
  if (elem != null && key == 13) {
    if (e && e.preventDefault) {
      e.preventDefault(); // DOM style
    } else {
      e.keyCode = 9;
    }
    if (elem.selectedIndex > -1) {
      addOffer(<% Sendb(CustomerPK) %>, <% Sendb(CardPK) %>, elem.value);
    }
    return false;
  }
  return true;
}

function toggleFavorites(option) {
  var frm = document.mainform;
  var currentURL = window.location.href;
  var newURL = "";
  
  newURL = currentURL + "&Favorite=" + option;
  
  frm.action = newURL;
  frm.submit();
}

function expandRow(offerID) {
  var trTranElem = document.getElementById("trTrans" + offerID);
  var imgElem = document.getElementById("plus" + offerID);
  var isOpen = false;
  var args = new Array(offerID, 0)
  var qryStr = 'CAMOfferTransactions=1&OfferID=' + offerID + '&CustPK=<%Sendb(CustomerPK)%>&ExtCustID=<%Sendb(ExtCustomerID)%>';
    
  if (imgElem != null) {
    isOpen = (imgElem.src.indexOf('minus.png') > -1) ? true : false;
    if (isOpen) {
      imgElem.src = '/images/plus.png';
    } else {
      imgElem.src = '/images/minus.png';
    }
  }
  
  if (trTranElem != null) {
    trTranElem.style.display = (isOpen) ? 'none' : '';
    if (!isOpen) { xmlhttpPost('CAM-Feeds.aspx?'+ qryStr, qryStr, 'ShowTrans', args); }
  }
}
  
function xmlhttpPost(strURL, qryStr, mode, args) {
  var xmlHttpReq = false;
  var self = this;
  var respTxt = '';
  var i = 0;
  
  // Mozilla/Safari
  if (window.XMLHttpRequest) {
    self.xmlHttpReq = new XMLHttpRequest();
  }
  // IE
  else if (window.ActiveXObject) {
    self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
  }
  strURL += "?" + qryStr;
  self.xmlHttpReq.open('POST', strURL, true);
  self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
  self.xmlHttpReq.onreadystatechange = function() {
    if (self.xmlHttpReq !=null && self.xmlHttpReq.readyState == 4) {
      respTxt = self.xmlHttpReq.responseText
      if (mode == 'ShowTrans') {
        updateOfferTransactions(respTxt, args[0]);
      } else if (mode == 'CreateTrans') {
        updateCreateTrans(respTxt);
      }
    }
  }
  self.xmlHttpReq.send(qryStr);
}

function updateOfferTransactions(respTxt, offerID) {
  var elem = document.getElementById('trTrans' + offerID);
  var elemTd = document.getElementById('tdTrans' + offerID);
  
  if (elem != null && elemTd != null) {
    elem.style.display = '';
    elemTd.style.display = '';
    elemTd.innerHTML = respTxt;
  }
}
//-->
</script>
<script type="text/javascript" src="/javascript/jquery.js"></script>
<script type="text/javascript" src="/javascript/thickbox.js"></script>

<form id="mainform" name="mainform" action="#">
  <div id="intro">
    <h1 id="title">
      <%
        If CardPK = 0 Then
          If (IsHouseholdID) Then
            Sendb(Copient.PhraseLib.Lookup("term.household", LanguageID))
          Else
            Sendb(Copient.PhraseLib.Lookup("term.customer", LanguageID))
          End If
        Else
          If (IsHouseholdID) Then
            Sendb(Copient.PhraseLib.Lookup("term.householdcard", LanguageID) & " #" & ExtCardID)
          Else
            Sendb(Copient.PhraseLib.Lookup("term.camcard", LanguageID) & " #" & ExtCardID)
          End If
        End If
        MyCommon.QueryStr = "select Prefix, FirstName, MiddleName, LastName, Suffix from Customers with (NoLock) where CustomerPK=" & CustomerPK & ";"
        rst2 = MyCommon.LXS_Select
        If rst2.Rows.Count > 0 Then
          FullName = IIf(MyCommon.NZ(rst2.Rows(0).Item("Prefix"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_Prefix, MyCommon.NZ(rst2.Rows(0).Item("Prefix"), "") & " ", "")
          FullName &= IIf(MyCommon.NZ(rst2.Rows(0).Item("FirstName"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_FirstName, MyCommon.NZ(rst2.Rows(0).Item("FirstName"), "") & " ", "")
          FullName &= IIf(MyCommon.NZ(rst2.Rows(0).Item("MiddleName"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_MiddleName, Left(MyCommon.NZ(rst2.Rows(0).Item("MiddleName"), ""), 1) & ". ", "")
          FullName &= IIf(MyCommon.NZ(rst2.Rows(0).Item("LastName"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_LastName, MyCommon.NZ(rst2.Rows(0).Item("LastName"), ""), "")
          FullName &= IIf(MyCommon.NZ(rst2.Rows(0).Item("Suffix"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_Suffix, " " & MyCommon.NZ(rst2.Rows(0).Item("Suffix"), ""), "")
        End If
        If FullName <> "" Then
          Sendb(": " & MyCommon.TruncateString(FullName, 30))
        End If
        If (restrictLinks AndAlso URLtrackBack <> "") Then
          Send(" - <a href=""" & URLtrackBack & """>" & Copient.PhraseLib.Lookup("customer-inquiry.return", LanguageID) & "</a>")
        End If
      %>
    </h1>
    <div id="controls"<% Sendb(IIf(restrictLinks AndAlso URLtrackBack <> "", " style=""width:230px;""", "")) %>>
      <%
        If (CustomerPK > 0 And Logix.UserRoles.ViewCustomerNotes) Then
          Send_CustomerNotes(CustomerPK, CardPK)
        End If
        If (CustomerPK > 0) Then
          Send_CAMAddOffer(CustomerPK, CardPK)
        End If
      %>
    </div>
  </div>
  <div id="main">
    <%
      If (Logix.UserRoles.ViewCustomerOffers = False) Then
        Send(Copient.PhraseLib.Lookup("error.forbidden", LanguageID))
        Send("</div>")
        Send("</form>")
        GoTo done
      End If
    %>
    <%
      If (infoMessage <> "") Then
        Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div><br class=""half"" />")
      End If
      Send("<input type=""hidden"" id=""CustomerPK"" name=""CustomerPK"" value=""" & CustomerPK & """ />")
      Send("<input type=""hidden"" id=""CustPK"" name=""CustPK"" value=""" & CustomerPK & """ />")
      If CardPK > 0 Then
        Send("<input type=""hidden"" id=""CardPK"" name=""CardPK"" value=""" & CardPK & """ />")
      End If
      If (Request.QueryString("mode") = "summary") Then
        Send("<input type=""hidden"" id=""mode"" name=""mode"" value=""summary"" />")
        Send("<input type=""hidden"" id=""exiturl"" name=""exiturl"" value=""" & URLtrackBack & """ />")
      End If
    %>
    <div id="column">
      <% If (Logix.UserRoles.ViewCustomerOffers AndAlso CustomerPK > 0) Then%>
      <a name="h01"></a>
      <div class="box" id="divSearch">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.search", LanguageID))%>
          </span>
        </h2>
        <div>
          <label for="offerterms"><b><%Sendb(Copient.PhraseLib.Lookup("term.offer", LanguageID) & " / " & Copient.PhraseLib.Lookup("term.group", LanguageID))%> :</b></label>
          <input type="text" id="offerterms" name="offerterms" style="font-family:arial;font-size:12px;width:70px;" value="<%Sendb(Request.QueryString("offerterms"))%>" onkeydown="submitOfferSearch(event);" />&nbsp;
          <label for="offerdate"><b><%Sendb(Copient.PhraseLib.Lookup("term.date", LanguageID))%>:</b></label>
          <input type="text" id="offerdate" name="offerdate" style="font-family:arial;font-size:12px;width:70px;" value="<%Sendb(Request.QueryString("offerdate"))%>" onkeydown="submitOfferSearch(event);" />&nbsp;
          <label for="offerstore"><b><%Sendb(Copient.PhraseLib.Lookup("term.store", LanguageID))%>:</b></label>
          <input type="text" id="offerstore" name="offerstore" style="font-family:arial;font-size:12px;width:70px;" value="<%Sendb(Request.QueryString("offerstore"))%>" onkeydown="submitOfferSearch(event);" />&nbsp;
          <label for="offerTrigger"><b><%Sendb(Copient.PhraseLib.Lookup("term.trigger", LanguageID))%>:</b></label>
          <input type="text" id="offerTrigger" name="offerTrigger" style="font-family:arial;font-size:12px;width:90px;" value="<%Sendb(Request.QueryString("offerTrigger"))%>" onkeydown="submitOfferSearch(event);" />&nbsp;
          <label for="offerUPC"><b><%Sendb(Copient.PhraseLib.Lookup("term.upc", LanguageID))%>:</b></label>
          <input type="text" id="offerUPC" name="offerUPC" style="font-family:arial;font-size:12px;width:90px;" value="<%Sendb(Request.QueryString("offerUPC"))%>" onkeydown="submitOfferSearch(event);" />&nbsp;
          <input type="button" style="font-family:arial;font-size:12px;" id="btnOffer" name="btnOffer" onclick="searchOffers();" value="<%Sendb(Copient.PhraseLib.Lookup("term.find", LanguageID)) %>" />        </div>
        <hr class="hidden" />
      </div>
      <br class="half" />
      <%
        Send("<div id=""listbar"">")
        Send("</div>")
      %>
      
      <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID)) %>">
        <thead>
          <tr>
            <th align="left" style="width:15px;"></th>
            <th align="left" class="th-button" scope="col" style="text-align: center;">
              <a href="CAM-customer-offers.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&offerterms=<%Sendb(Request.QueryString("offerterms"))%>&CustPK=<%Sendb(CustomerPK)%><% Sendb(IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "")) %>&search=Search&SortText=Priority&SortDirection=<% Sendb(SortDirection & extralink) %>&Favorite=<% Sendb(Request.QueryString("Favorite")) %>&offerdate=<%Sendb(Request.QueryString("offerdate"))%>&offerstore=<%Sendb(Request.QueryString("offerstore"))%>&offerTrigger=<%Sendb(Request.QueryString("offerTrigger"))%>&offerUPC=<%Sendb(Request.QueryString("offerUPC"))%>">
                <% Sendb(Copient.PhraseLib.Lookup("term.fav", LanguageID))%>
              </a>
              <%
                If SortText = "Priority" Then
                  If SortDirection = "ASC" Then
                    Sendb("<span class=""sortarrow"">&#9660;</span>")
                  Else
                    Sendb("<span class=""sortarrow"">&#9650;</span>")
                  End If
                Else
                End If
              %>
            </th>
            <th align="left" class="th-button" scope="col" style="text-align: center;">
              Rem
            </th>
            <th align="center" class="th-button" scope="col" style="text-align: center;">
              <% Sendb(Copient.PhraseLib.Lookup("term.adjust", LanguageID))%>
            </th>
            <th align="left" class="th-bigid" scope="col">
              <a href="CAM-customer-offers.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&offerterms=<%Sendb(Request.QueryString("offerterms"))%>&CustPK=<%Sendb(CustomerPK)%><% Sendb(IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "")) %>&search=Search&SortText=OfferID&SortDirection=<% Sendb(SortDirection & extralink) %>&Favorite=<% Sendb(Request.QueryString("Favorite")) %>&offerdate=<%Sendb(Request.QueryString("offerdate"))%>&offerstore=<%Sendb(Request.QueryString("offerstore"))%>&offerTrigger=<%Sendb(Request.QueryString("offerTrigger"))%>&offerUPC=<%Sendb(Request.QueryString("offerUPC"))%>">
                <% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>
              </a>
              <%
                If SortText = "OfferID" Then
                  If SortDirection = "ASC" Then
                    Sendb("<span class=""sortarrow"">&#9660;</span>")
                  Else
                    Sendb("<span class=""sortarrow"">&#9650;</span>")
                  End If
                Else
                End If
              %>
            </th>
            <th align="left" class="th-name" scope="col">
              <a href="CAM-customer-offers.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&offerterms=<%Sendb(Request.QueryString("offerterms"))%>&CustPK=<%Sendb(CustomerPK)%><% Sendb(IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "")) %>&search=Search&SortText=Name&SortDirection=<% Sendb(SortDirection & extralink) %>&Favorite=<% Sendb(Request.QueryString("Favorite")) %>&offerdate=<%Sendb(Request.QueryString("offerdate"))%>&offerstore=<%Sendb(Request.QueryString("offerstore"))%>&offerTrigger=<%Sendb(Request.QueryString("offerTrigger"))%>&offerUPC=<%Sendb(Request.QueryString("offerUPC"))%>">
                <% Sendb(Copient.PhraseLib.Lookup("term.offer", LanguageID))%>
              </a>
              <%
                If SortText = "Name" Then
                  If SortDirection = "ASC" Then
                    Sendb("<span class=""sortarrow"">&#9660;</span>")
                  Else
                    Sendb("<span class=""sortarrow"">&#9650;</span>")
                  End If
                Else
                End If
              %>
            </th>
            <th align="center" class="th-status" scope="col">
              <% 
                If (Date.TryParse(Request.QueryString("offerdate"), OfferSearchDate)) Then
                  Sendb(Copient.PhraseLib.Lookup("term.status", LanguageID) & " " & Copient.PhraseLib.Lookup("term.on", LanguageID))
                  Sendb("<br />" & OfferSearchDate.ToShortDateString)
                Else
                  Sendb(Copient.PhraseLib.Lookup("term.status", LanguageID))
                End If
                %>
              <%If SortText = "StatusFlag" Then
                  If SortDirection = "ASC" Then
                    Sendb("<span class=""sortarrow"">&#9660;</span>")
                  Else
                    Sendb("<span class=""sortarrow"">&#9650;</span>")
                  End If
                Else
                End If
              %>
            </th>
            <th align="center" class="th-group" scope="col">
              <a href="CAM-customer-offers.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&offerterms=<%Sendb(Request.QueryString("offerterms"))%>&CustPK=<%Sendb(CustomerPK)%><% Sendb(IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "")) %>&search=Search&SortText=GroupName&SortDirection=<% Sendb(SortDirection & extralink) %>&Favorite=<% Sendb(Request.QueryString("Favorite")) %>&offerdate=<%Sendb(Request.QueryString("offerdate"))%>&offerstore=<%Sendb(Request.QueryString("offerstore"))%>&offerTrigger=<%Sendb(Request.QueryString("offerTrigger"))%>&offerUPC=<%Sendb(Request.QueryString("offerUPC"))%>">
                <% Sendb(Copient.PhraseLib.Lookup("term.customergroup", LanguageID))%>
              </a>
              <%
                If SortText = "GroupName" Then
                  If SortDirection = "ASC" Then
                    Sendb("<span class=""sortarrow"">&#9660;</span>")
                  Else
                    Sendb("<span class=""sortarrow"">&#9650;</span>")
                  End If
                Else
                End If
              %>
            </th>
          </tr>
        </thead>
        <tbody>
          <%
            offerCt = 0
            
            CgXml = "<customergroups>"
            MyCommon.QueryStr = "select CustomerGroupID from CustomerGroups with (NoLock) where AnyCAMCardholder=1;"
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count > 0) Then
              AllCAMCardholdersID = MyCommon.NZ(rst.Rows(0).Item("CustomerGroupID"), 0)
              CgXml &= "<id>" & AllCAMCardholdersID & "</id>"
            End If

            MyCommon.QueryStr = "select CustomerGroupID from GroupMembership with (NoLock) where CustomerPK=" & CustomerPK & " and Deleted=0"
            rst = MyCommon.LXS_Select()

            rowCount = rst.Rows.Count
            If rowCount > 0 Then
              For Each row In rst.Rows
                CgXml &= "<id>" & MyCommon.NZ(row.Item("CustomerGroupID"), "") & "</id>"
              Next
            End If
            CgXml &= "</customergroups>"
            
            MyCommon.QueryStr = "dbo.pa_CAM_CustomerOffersCurrent"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@cgXml", SqlDbType.Xml).Value = CgXml
            MyCommon.LRTsp.Parameters.Add("@IsEmployee", SqlDbType.Bit).Value = Employee
            If (Request.QueryString("offerterms") <> "") Then
              Filter = Request.QueryString("offerterms")
              Filter = Filter.Replace("%", "[%]")
              Filter = Filter.Replace("_", "[_]")
              MyCommon.LRTsp.Parameters.Add("@Filter", SqlDbType.NVarChar, 50).Value = Filter
            End If
            MyCommon.LRTsp.Parameters.Add("@Favorite", SqlDbType.Bit).Value = Favorite
            MyCommon.LRTsp.Parameters.Add("@AdminUserID", SqlDbType.Int).Value = AdminUserID
            reader = MyCommon.LRTsp.ExecuteReader
            
            Dim ds As New DataSet()
            dtAssigned = New DataTable
            ds.Tables.Add(dtAssigned)
            dtAddOffers = New DataTable
            ds.Tables.Add(dtAddOffers)
            
            ds.Load(reader, LoadOption.OverwriteChanges, Nothing, New DataTable() {dtAssigned, dtAddOffers})
            
            MyCommon.Close_LRTsp()
            reader.Close()
            
            'If showing credits only, get the list of customer offers that
            'have balances, then drop non-matches from the main list.
            If (ShowCreditOnly) Then
              BalanceOffers = GetBalanceOffers(CustomerPK)
              'Send("<h3>BalanceOffers</h3>")
              'Dim OID As Integer
              'For Each OID In BalanceOffers.Keys
              '  Send(OID & ", ")
              'Next
              i = 0
              While i < dtAssigned.Rows.Count
                If Not BalanceOffers.ContainsKey(MyCommon.NZ(dtAssigned.Rows(i).Item("OfferID"), 0).ToString) Then
                  dtAssigned.Rows.RemoveAt(i)
                  dtAssigned.AcceptChanges()
                Else
                  i += 1
                End If
              End While
            End If
            
            'Sort the Assigned offers
            If (Date.TryParse(Request.QueryString("offerdate"), OfferSearchDate)) Then
              sortedRows = dtAssigned.Select("StartDate <= '" & OfferSearchDate.ToString("yyyy-MM-dd 00:00:00") & "' ", SortText & " " & SortDirection)
              SearchDateSet = True
            ElseIf InvalidSearchDate Then
              ' return no records if the date is invalid
              sortedRows = dtAssigned.Select("1=2", SortText & " " & SortDirection)
            Else
              sortedRows = dtAssigned.Select("", SortText & " " & SortDirection)
            End If
            
            ' filter out offers that don't match of the store criteria
            If (Request.QueryString("offerstore") <> "" AndAlso Request.QueryString("offerstore").Trim <> "") Then
              MyCommon.QueryStr = "dbo.pa_CAM_CustomerOffersByLocation"
              MyCommon.Open_LRTsp()
              MyCommon.LRTsp.Parameters.Add("@ExtLocationCode", SqlDbType.NVarChar, 20).Value = Request.QueryString("offerstore").Trim
              rst = MyCommon.LRTsp_select
              MyCommon.Close_LRTsp()
              If rst.Rows.Count > 0 Then
                ' load up a hashtable with all the offers found at this location
                OfferTable = New Hashtable(rst.Rows.Count)
                For Each row2 In rst.Rows
                  OfferTable.Add(MyCommon.NZ(row2.Item("OfferID"), "-1").ToString, MyCommon.NZ(row2.Item("OfferID"), "-1").ToString)
                Next
                ' remove the offers that don't match the store
                For Each row2 In sortedRows
                  If Not OfferTable.Contains(MyCommon.NZ(row2.Item("OfferID"), "0").ToString) Then
                    row2.Delete()
                  End If
                Next
              Else
                ' no matches for this location were found, so remove all the offers
                For Each row2 In sortedRows
                  row2.Delete()
                Next
              End If
            End If
            
            ' filter out offers that don't match the UPC criteria
            If (Request.QueryString("offerUPC") <> "" AndAlso Request.QueryString("offerUPC").Trim <> "") Then
              MyCommon.QueryStr = "dbo.pa_CAM_CustomerOffersByProduct"
              MyCommon.Open_LRTsp()
              MyCommon.LRTsp.Parameters.Add("@ExtProductID", SqlDbType.NVarChar, 120).Value = PaddedProdExtID 'Request.QueryString("offerUPC").Trim 
              MyCommon.LRTsp.Parameters.Add("@ProductTypeID", SqlDbType.Int).Value = 1 ' use UPC product type = 1
              rst = MyCommon.LRTsp_select
              MyCommon.Close_LRTsp()
              If rst.Rows.Count > 0 Then
                ' load up a hashtable with all the offers found for this product
                OfferTable = New Hashtable(rst.Rows.Count)
                For Each row2 In rst.Rows
                  If row2.RowState <> DataRowState.Deleted Then
                    OfferTable.Add(MyCommon.NZ(row2.Item("OfferID"), "-1").ToString, MyCommon.NZ(row2.Item("OfferID"), "-1").ToString)
                  End If
                Next
                ' remove the offers that don't match the store
                For Each row2 In sortedRows
                  If row2.RowState <> DataRowState.Deleted Then
                    If Not OfferTable.Contains(MyCommon.NZ(row2.Item("OfferID"), "0").ToString) Then
                      row2.Delete()
                    End If
                  End If
                Next
              Else
                ' no matches for this product were found, so remove all the offers
                For Each row2 In sortedRows
                  If row2.RowState <> DataRowState.Deleted Then
                    row2.Delete()
                  End If
                Next
              End If
            End If
            
            ' filter out offers that don't match the Trigger code criteria
            If (Request.QueryString("offerTrigger") <> "" AndAlso Request.QueryString("offerTrigger").Trim <> "") Then
              MyCommon.QueryStr = "select PLU.RewardOptionID, RO.IncentiveID as OfferID from CPE_IncentivePLUs as PLU with (NoLock) " & _
                                  "inner join CPE_RewardOptions as RO on PLU.RewardOptionID=RO.RewardOptionID " & _
                                  "where PLU='" & PaddedTrigger & "';"
              rst = MyCommon.LRT_Select()
              If rst.Rows.Count > 0 Then
                ' load up a hashtable with all the offers found with this trigger code
                OfferTable = New Hashtable(rst.Rows.Count)
                For Each row2 In rst.Rows
                  If row2.RowState <> DataRowState.Deleted Then
                    OfferTable.Add(MyCommon.NZ(row2.Item("OfferID"), "-1").ToString, MyCommon.NZ(row2.Item("OfferID"), "-1").ToString)
                  End If
                Next
                ' remove the offers that don't match the store
                For Each row2 In sortedRows
                  If row2.RowState <> DataRowState.Deleted Then
                    If Not OfferTable.Contains(MyCommon.NZ(row2.Item("OfferID"), "0").ToString) Then
                      row2.Delete()
                    End If
                  End If
                Next
              Else
                ' no matches for this trigger were found, so remove all the offers
                For Each row2 In sortedRows
                  If row2.RowState <> DataRowState.Deleted Then
                    row2.Delete()
                  End If
                Next
              End If
            End If
            
            If (SearchDateSet) Then
              Date.TryParse(OfferSearchDate.ToString("yyyy-MM-dd 23:59:59"), OfferStatusDate)
            Else
              Date.TryParse(Now.ToString("yyyy-MM-dd 23:59:59"), OfferStatusDate)
            End If
            StatusTable = LoadOfferStatuses(sortedRows, OfferStatusDate, MyCommon, Logix)
            
            i = 0
            For i = 0 To sortedRows.Length - 1
              row2 = sortedRows(i)
              If row2.RowState <> DataRowState.Deleted AndAlso PrevOfferID <> MyCommon.NZ(row2.Item("OfferID"), 0) Then
                OfferStatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_UNKNOWN
                OfferStatus = StatusTable.Item(MyCommon.NZ(row2.Item("OfferID"), "0").ToString)
                If (OfferStatus IsNot Nothing) Then
                  OfferStatusCode = OfferStatus
                End If
                'If (TestCard = False AndAlso (OfferStatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_EXPIRED OrElse OfferStatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_ACTIVE)) _
                '    OrElse (TestCard = True AndAlso OfferStatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_TESTING AndAlso MyCommon.Fetch_SystemOption(88) = "1") Then
                'Filter Offers Shown
                If (TestCard = False AndAlso IIf(ExcludeExpired, OfferStatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_ACTIVE, (OfferStatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_EXPIRED OrElse OfferStatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_ACTIVE))) _
                    OrElse (TestCard = True AndAlso OfferStatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_TESTING AndAlso MyCommon.Fetch_SystemOption(88) = "1") Then
                  offerCt += 1
                  System.Math.DivRem(offerCt, 2, r)
                  Send("<tr id=""trOffer" & offerCt & """" & IIf(r = 0, "", " class=""shaded""") & " style=""display:none;"">")
                  
                  ' Expander column
                  Send("<td><img id=""plus" & MyCommon.NZ(row2.Item("OfferID"), "-1") & """ src=""/images/plus.png"" style=""cursor:hand;"" onclick=""expandRow(" & MyCommon.NZ(row2.Item("OfferID"), "-1") & ");"" /></td>")
                  
                  ' Favorite column
                  Send("  <td align=""center"">")
                  If (MyCommon.NZ(row2.Item("AdminUserID"), 0)) = AdminUserID Then
                    If Logix.UserRoles.FavoriteOffersForSelf Then
                      Sendb("    <a href=""CAM-customer-offers.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&CardPK=" & CardPK, "") & "&DeleteFavorite=" & MyCommon.NZ(row2.Item("OfferID"), -1) & "&Favorite=" & Favorite & "&SortText=" & SortText & "&SortDirection=" & IIf(SortDirection = "ASC", "DESC", "ASC") & "&offerterms=" & Request.QueryString("offerterms") & "&offerdate=" & Request.QueryString("offerdate") & "&offerstore=" & Request.QueryString("offerstore") & "&offerUPC=" & Request.QueryString("offerUPC") & """>")
                      Send("<img src=""/images/star-on.png"" alt=""" & Copient.PhraseLib.Lookup("customer-inquiry.unfavorite", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("customer-inquiry.unfavorite", LanguageID) & """ /></a>")
                    Else
                      Send("    <img src=""/images/star-on.png"" alt=""" & Copient.PhraseLib.Lookup("term.favorite", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.favorite", LanguageID) & """ /></a>")
                    End If
                  Else
                    If Logix.UserRoles.FavoriteOffersForSelf Then
                      Sendb("    <a href=""CAM-customer-offers.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&CardPK=" & CardPK, "") & "&AddFavorite=" & MyCommon.NZ(row2.Item("OfferID"), -1) & "&Favorite=" & Favorite & "&SortText=" & SortText & "&SortDirection=" & IIf(SortDirection = "ASC", "DESC", "ASC") & "&offerterms=" & Request.QueryString("offerterms") & "&offerdate=" & Request.QueryString("offerdate") & "&offerstore=" & Request.QueryString("offerstore") & "&offerUPC=" & Request.QueryString("offerUPC") & """>")
                      Send("<img src=""/images/star-off.png"" alt=""" & Copient.PhraseLib.Lookup("customer-inquiry.favorite", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("customer-inquiry.favorite", LanguageID) & """ /></a>")
                    Else
                      Send("    <img src=""/images/star-off.png"" alt=""" & Copient.PhraseLib.Lookup("customer-inquiry.notafavorite", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("customer-inquiry.notafavorite", LanguageID) & """ /></a>")
                    End If
                  End If
                  Send("  </td>")
                  
                  ' Remove column
                  If (MyCommon.NZ(row2.Item("CustomerGroupID"), -1) <= 2) OrElse MyCommon.NZ(row2.Item("CustomerGroupID"), -1) = AllCAMCardholdersID Then
                    Send("  <td></td>")
                  Else
                    Send("  <td align=""center"">")
                    RoleMatch = False
                    If (MyCommon.NZ(row2.Item("EditControlTypeID"), 0) = 3) Then 'removal is limited per role, so loop through the user's roles and see if there's a match to the customer group's RoleID 
                      For x = 0 To UserRoleIDs.Length - 1
                        If UserRoleIDs(x) = MyCommon.NZ(row2.Item("RoleID"), 0) Then
                          RoleMatch = True
                        End If
                      Next
                    End If
                    If (MyCommon.NZ(row2.Item("EditControlTypeID"), 0) = 0 AndAlso Logix.UserRoles.RemoveCustomerFromOffers) OrElse (MyCommon.NZ(row2.Item("EditControlTypeID"), 0) = 1) OrElse (MyCommon.NZ(row2.Item("EditControlTypeID"), 0) = 3 AndAlso RoleMatch) Then
                      Sendb("    <a href=""/logix/XMLFeeds.aspx?OtherOfferscheck=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "&amp;OfferID=" & MyCommon.NZ(row2.Item("OfferID"), -1) & "&amp;CustomerGroupID=" & MyCommon.NZ(row2.Item("CustomerGroupID"), -1) & "&amp;ExtCustomerID=" & ExtCustomerID & "&amp;AdminUserID=" & AdminUserID & "&amp;height=300&amp;width=300"" title=""" & Copient.PhraseLib.Lookup("term.alert", LanguageID) & """ class=""thickbox"">")
                      Send("<input type=""button"" class=""ex"" id=""ex" & MyCommon.NZ(row2.Item("OfferID"), -1) & """ name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.remove", LanguageID) & """ " & " value=""X"" /></a>")
                    End If
                    Send("  </td>")
                  End If
                  
                  ' Adjust points
                  Send("  <td style=""width:30px;"">")
                  
                  MyCommon.QueryStr = "dbo.pa_CustomerOfferHasPointsProgram"
                  MyCommon.Open_LRTsp()
                  MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.Int).Value = MyCommon.NZ(row2.Item("OfferID"), -1)
                  MyCommon.LRTsp.Parameters.Add("@HasPointsProgram", SqlDbType.Bit).Direction = ParameterDirection.Output
                  pointdt = MyCommon.LRTsp_select()
                  IsPtsOffer = MyCommon.LRTsp.Parameters("@HasPointsProgram").Value
                  MyCommon.Close_LRTsp()
                  
                  'Find if the customer has points if there is a pointsprogram
                  If pointdt.Rows.Count > 0 Then
                    If pointdt.Rows.Count = 1 Then
                      If MyCommon.NZ(pointdt.Rows(0).Item("ProgramID"), 0) > 0 Then
                        PointsProgram = MyCommon.Extract_Val(MyCommon.NZ(pointdt.Rows(0).Item("ProgramID"), 0))
                      End If
                      If PointsProgram > 0 Then
                        HasPoints = PointsTable.ContainsKey(PointsProgram)
                      End If
                    Else
                      For Each pointRow In pointdt.Rows
                        If MyCommon.NZ(pointRow.Item("ProgramID"), 0) > 0 Then
                          PointsProgram = MyCommon.Extract_Val(MyCommon.NZ(pointRow.Item("ProgramID"), 0))
                        End If
                        If PointsProgram > 0 Then
                          HasPoints = PointsTable.ContainsKey(PointsProgram)
                        End If
                        If HasPoints Then Exit For
                      Next
                    End If
                  End If
                  
                  If (IsPtsOffer) Then
                    If (Logix.UserRoles.AccessPointsBalances = False) Then
                      DisabledPtsAdj = " disabled=""disabled"""
                    Else
                      DisabledPtsAdj = ""
                    End If
                    Sendb("  <input type=""button"" class=""adjust"" id=""ptsAdj" & MyCommon.NZ(row2.Item("OfferID"), "") & """ name=""ptsAdj"" value=""P""" & DisabledPtsAdj & " " & IIf(HasPoints, "style=""""", " style=""background-color:#CCCCCC;font-style:italic;""") & " title=""" & Copient.PhraseLib.Lookup("term.adjust", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.points", LanguageID), VbStrConv.Lowercase) & """ ")
                    Send("onClick=""javascript:openPopup('CAM-point-adjust.aspx?OfferID=" & MyCommon.NZ(row2.Item("OfferID"), "") & "&amp;OfferName=" & Server.UrlEncode(MyCommon.NZ(row2.Item("Name"), "")).ToString.Replace("'", "\'") & "&amp;CustomerPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "&amp;offerdate=" & Request.QueryString("offerdate") & "');"" />")
                  End If
                  Send("  </td>")
                  
                  ' Offer ID
                  Send("  <td>" & MyCommon.NZ(row2.Item("OfferID"), -1) & "</td>")
                  
                  ' Offer name and link
                  If (Not restrictLinks) Then
                    Send("  <td><a href=""" & URLCAMOfferSum & "?OfferID=" & MyCommon.NZ(row2.Item("OfferID"), -1) & """ title=""" & MyCommon.NZ(row2.Item("Description"), "").ToString().Replace("""", "'") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row2.Item("Name"), ""), 25) & "</a></td>")
                  Else
                    Send("  <td><a href=""javascript:openPopup('CAM-offer-sum.aspx?OfferID=" & MyCommon.NZ(row2.Item("OfferID"), -1) & "&Popup=1')"" title=""" & MyCommon.NZ(row2.Item("Description"), "").ToString().Replace("""", "'") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row2.Item("Name"), ""), 25) & "</a></td>")
                  End If
                  
                  ' Offer status
                  Send("  <td>")
                  Send("    " & Logix.GetOfferStatusHtml(Integer.Parse(OfferStatus), LanguageID))
                  Send("  </td>")
                  
                  ' Customer group
                  If (MyCommon.NZ(row2.Item("CustomerGroupID"), -1) <> 1 And MyCommon.NZ(row2.Item("CustomerGroupID"), -1) <> 2) And (MyCommon.NZ(row2.Item("NewCardholders"), False) = False) And MyCommon.NZ(row2.Item("CustomerGroupID"), -1) <> AllCAMCardholdersID Then
                    Send("  <td>")
                    ' get all the customer groups assigned to this offer.  Peek ahead after getting the current row to get any additional customer groups assigned to this offer
                    For j = i To sortedRows.Length - 1
                      If sortedRows(j).RowState <> DataRowState.Deleted AndAlso MyCommon.NZ(sortedRows(j).Item("OfferID"), 0) = MyCommon.NZ(row2.Item("OfferID"), 0) Then
                        If (Not restrictLinks) Then
                          Send("<a href=""" & URLcgroupedit & "?CustomerGroupID=" & MyCommon.NZ(sortedRows(j).Item("CustomerGroupID"), -1) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(sortedRows(j).Item("GroupName"), ""), 25) & "</a><br />")
                        Else
                          Send(MyCommon.SplitNonSpacedString(MyCommon.NZ(sortedRows(j).Item("GroupName"), ""), 25))
                        End If
                      Else
                        Exit For
                      End If
                    Next
                    Send("</td>")
                  Else
                    Send("  <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row2.Item("GroupName"), ""), 25) & "</td>")
                  End If
                  Send("</tr>")
                  
                  'create the Transactions row
                  Send("<tr id=""trTrans" & MyCommon.NZ(row2.Item("OfferID"), "-1") & """ class=""transrow"" style=""display:none;background-color:#dddddd;"" >")
                  Send("  <td id = ""tdTrans" & MyCommon.NZ(row2.Item("OfferID"), "-1") & """ colspan=""8"">")
                  Send("    <img src=""/images/loadingAnimation.gif"" />")
                  Send("  </td>")
                  Send("</tr>")
                End If
              End If
              If row2.RowState <> DataRowState.Deleted Then PrevOfferID = MyCommon.NZ(row2.Item("OfferID"), 0)
            Next
            If offerCt = 0 AndAlso Favorite = 1 Then
              Send("<tr>")
              Send("  <td colspan=""8"" align=""center"" style=""padding-top:5px;"">" & Copient.PhraseLib.Lookup("customer-inquiry.nofavorites", LanguageID) & "</td>")
              Send("</tr>")
            End If
          %>
        </tbody>
      </table>
      <%
        Send(" <div id=""pageIter"">")
        Send("<div id=""searcher"" title=""Search terms""></div>")
        If (offerCt > 0) Then
          Send("  <div id=""paginator"">")
          Send("   <span id=""first"" style=""display:none;""><a href=""javascript:showOfferFirstPage();""><b>|</b>◄" & Copient.PhraseLib.Lookup("term.first", LanguageID) & "</a>&nbsp;</span>")
          Send("   <span id=""previous"" style=""display:none;""><a href=""javascript:showOfferPrevPage();"">◄" & Copient.PhraseLib.Lookup("term.previous", LanguageID) & "</a>&nbsp;</span>")
          Send("   <span id=""firstOff""><b>|</b>◄" & Copient.PhraseLib.Lookup("term.first", LanguageID) & "&nbsp;</span>")
          Send("   <span id=""previousOff"">◄" & Copient.PhraseLib.Lookup("term.previous", LanguageID) & "&nbsp;</span>")
        Else
          Send("  <div id=""paginator"">")
          Send("   <span id=""firstOff""><b>|</b>◄" & Copient.PhraseLib.Lookup("term.first", LanguageID) & "&nbsp;</span>")
          Send("   <span id=""previousOff"">◄" & Copient.PhraseLib.Lookup("term.previous", LanguageID) & "&nbsp;</span>")
        End If
        If offerCt = 0 Then
          Send("   &nbsp;[ " & Copient.PhraseLib.Lookup("term.noresults", LanguageID) & " ]&nbsp;")
        Else
          Send("   &nbsp;[ <b><span id=""startPos""></span>-<span id=""endPos""></span></b> " & Copient.PhraseLib.Lookup("term.of", LanguageID) & " <b>" & offerCt & "</b> ]&nbsp;")
        End If
        If (offerCt > 0) Then
          Send("   <span id=""next""><a href=""javascript:showOfferNextPage();"">" & Copient.PhraseLib.Lookup("term.next", LanguageID) & "►</a>&nbsp;</span>")
          Send("   <span id=""last""><a href=""javascript:showOfferLastPage();"">" & Copient.PhraseLib.Lookup("term.last", LanguageID) & "►<b>|</b></a>&nbsp;</span>")
          Send("   <span id=""nextOff"" style=""display:none;"">" & Copient.PhraseLib.Lookup("term.next", LanguageID) & "►&nbsp;</span>")
          Send("   <span id=""lastOff"" style=""display:none;"">" & Copient.PhraseLib.Lookup("term.last", LanguageID) & "►<b>|</b>&nbsp;</span>")
        Else
          Send("   <span id=""nextOff"" style=""display:none;"">" & Copient.PhraseLib.Lookup("term.next", LanguageID) & "►&nbsp;</span>")
          Send("   <span id=""lastOff"" style=""display:none;"">" & Copient.PhraseLib.Lookup("term.last", LanguageID) & "►<b>|</b></span><br />")
        End If
        Send("  </div>")
        Send("  <div id=""filter"" title=""Filter"">")
        Send("   <select id=""Favorite"" name=""Favorite"" onchange=mainform.submit()>")
        Send("    <option value=""0""" & IIf(Favorite = 0 AndAlso Not ExcludeExpired, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("customer-inquiry.showall", LanguageID) & "</option>")
        Send("    <option value=""1""" & IIf(Favorite = 1, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("customer-inquiry.showfavorites", LanguageID) & "</option>")
        Send("    <option value=""2""" & IIf(Favorite = 0 AndAlso ExcludeExpired, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showall-noexpired", LanguageID) & "</option>")
        Send("    <option value=""5""" & IIf(Favorite = 0 AndAlso ShowCreditOnly, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showonlycredit", LanguageID) & "</option>")
        Send("   </select>")
        Send("  </div>")
        Send(" </div>")
        Send(" <input type=""hidden"" id=""offerTableRowCt"" name=""offerTableRowCt"" value=""" & offerCt & """ />")
      %>
      <!--<b><a href="#">View accumulation adjustment history</a></b><br />-->
      <hr class="hidden" />
      <% End If%>
    </div>
    <br clear="all" />
  </div>
</form>
<%Send("<script type=""text/javascript"">")
  If (Request.QueryString("refresh") = "") Then
    Send(" if (document.mainform.searchterms != null) document.mainform.searchterms.focus();")
  End If
  Send(" if (document.getElementById(""listbar"") != null && document.getElementById(""pageIter"") !=null ) { ")
  Send("   document.getElementById(""listbar"").innerHTML = document.getElementById(""pageIter"").innerHTML;")
  Send("   document.getElementById(""pageIter"").innerHTML = """ & """")
  Send(" }")
  Send("   showOfferPage(1);")
  Send("</script>")
%>

<script type="text/javascript" language="javascript">
// This is the javascript array holding the function list
// The PrintJavascriptArray ASP function can be used to print this array.
<%
  Dim elementBuf as new StringBuilder()
  If (dtAddOffers IsNot Nothing) then
    If (dtAddOffers.rows.count>0)
      Sendb("var functionlist = Array(")
      For Each row In dtAddOffers.Rows
        Sendb("""" & MyCommon.NZ(row.item("Name"), "").ToString().Replace("""", "\""") & """,")
      Next
      Sendb(""""");")
      Sendb("var vallist = Array(")
      For Each row In dtAddOffers.Rows
        Sendb("""" & MyCommon.NZ(row.item("OfferId"), -1) & """,")
      Next
      Sendb(""""");")
      Sendb("var pointerlist = Array(")
      i = 0
      For Each row In dtAddOffers.Rows
        Sendb("""" & i & """,")
        i += 1
      Next
      Sendb(""""");")
    Else
      Sendb("var functionlist = Array(")
      Send("""" & "" & """);")
      Sendb("var vallist = Array(")
      Send("""" & "" & """);")
      Sendb("var pointerlist = Array(")
      Send("""" & "" & """);")
    End If
  End If
%>
</script>

<script runat="server">
  Function GetCustomerPK(ByRef MyCommon As Copient.CommonInc, ByVal PrimaryExtID As String) As Integer
    Dim dt As DataTable
    Dim CustomerPK As Integer = 0
    
    If (Not MyCommon.LXSadoConn.State = ConnectionState.Open) Then MyCommon.Open_LogixXS()
    
    MyCommon.QueryStr = "select CustomerPK from Customers with (NoLock) where CustomerTypeID=2 and PrimaryExtID='" & PrimaryExtID & "';"
    dt = MyCommon.LXS_Select
    If (dt.Rows.Count > 0) Then
      CustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)
    End If
    
    Return CustomerPK
  End Function
  
  Function LoadOfferStatuses(ByVal rows() As DataRow, ByVal StatusDate As Date, ByRef MyCommon As Copient.CommonInc, ByRef Logix As Copient.LogixInc) As Hashtable
    Dim Statuses As New Hashtable(200)
    Dim i, ct, activeCt As Integer
    Dim OfferList As New ArrayList(10)
    Dim FilteredList(0) As String
    
    ct = rows.Length
    If (ct > 0) Then
      For i = 0 To ct - 1
        If rows(i).RowState <> DataRowState.Deleted Then
          OfferList.Add(MyCommon.NZ(rows(i).Item("OfferID"), "0"))
          activeCt += 1
        End If
      Next
      ' trim the offer list array to remove the empty elements caused by the filtered-out rows
      If (activeCt >= 1) Then
        ReDim FilteredList(OfferList.Count - 1)
        For i = 0 To FilteredList.GetUpperBound(0)
          FilteredList(i) = OfferList.Item(i)
        Next
        Statuses = Logix.GetStatusForOffers(FilteredList, LanguageID, StatusDate)
      End If
    End If
    
    Return Statuses
  End Function
  
  Function GetBalanceOffers(ByVal CustomerPK As Long) As Hashtable
    Dim MyCommon As New Copient.CommonInc
    Dim dt As DataTable
    Dim row As DataRow
    Dim BalancePointsPrograms As String = "-1"
    Dim BalanceSVPrograms As String = "-1"
    Dim BalanceAccumRoids As String = "-1"
    Dim BalanceOffers As New Hashtable
    
    MyCommon.Open_LogixRT()
    MyCommon.Open_LogixXS()
    
    'Points programs
    MyCommon.QueryStr = "select distinct ProgramID from Points with (NoLock) " & _
                        "where Amount>0 and CustomerPK=" & CustomerPK & " order by ProgramID;"
    dt = MyCommon.LXS_Select
    If dt.Rows.Count > 0 Then
      For Each row In dt.Rows
        BalancePointsPrograms &= "," & MyCommon.NZ(row.Item("ProgramID"), 0)
      Next
    End If
    
    'And now the big ol' query that gets all the offers that have a balance
    MyCommon.QueryStr = " select DISTINCT O.IncentiveID as OfferID from CPE_Incentives as O " & _
                        " INNER JOIN CPE_RewardOptions as RO on RO.IncentiveID=O.IncentiveID " & _
                        " INNER JOIN CPE_DeliverablePoints as DP on DP.RewardOptionID=RO.RewardOptionID " & _
                        " where DP.ProgramID in (" & BalancePointsPrograms & ") " & _
                        " and O.Deleted=0 and RO.Deleted=0 and DP.Deleted=0 " & _
                        "   union " & _
                        " select DISTINCT O.IncentiveID as OfferID from CPE_Incentives as O " & _
                        " INNER JOIN CPE_RewardOptions as RO on RO.IncentiveID=O.IncentiveID " & _
                        " INNER JOIN CPE_IncentivePointsGroups as IP on IP.RewardOptionID=RO.RewardOptionID " & _
                        " where IP.ProgramID in (" & BalancePointsPrograms & ") " & _
                        " and O.Deleted=0 and RO.Deleted=0 and IP.Deleted=0 " & _
                        "   union " & _
                        " select DISTINCT O.OfferID from Offers as O " & _
                        " INNER JOIN OfferConditions as C on C.OfferID=O.OfferID " & _
                        " where C.ConditionTypeID=3 and C.LinkID in (" & BalancePointsPrograms & ") " & _
                        " and O.Deleted=0 and C.Deleted=0 " & _
                        "   union " & _
                        " select DISTINCT O.OfferID from Offers as O " & _
                        " INNER JOIN OfferRewards as R on R.OfferID=O.OfferID " & _
                        " where R.RewardTypeID=2 and R.LinkID in (" & BalancePointsPrograms & ") " & _
                        " and O.Deleted=0 and R.Deleted=0 " & _
                        "ORDER BY OfferID;"
    dt = MyCommon.LRT_Select
    If dt.Rows.Count > 0 Then
      For Each row In dt.Rows
        BalanceOffers.Add(MyCommon.NZ(row.Item("OfferID"), 0).ToString, MyCommon.NZ(row.Item("OfferID"), 0).ToString)
      Next
    End If
    
    MyCommon.Close_LogixRT()
    MyCommon.Close_LogixXS()
    
    Return BalanceOffers
    
  End Function
  
  Function CleanDescription(ByVal Description As String) As String
    Dim CleanDesc As String = Description
    
    CleanDesc = Description.Replace(Chr(13) & Chr(10), "<br />")
    CleanDesc = CleanDesc.Replace(Chr(13), "<br />")
    CleanDesc = CleanDesc.Replace(Chr(10), "<br />")
    CleanDesc = CleanDesc.Replace("'", "\'")
    Return CleanDesc
  End Function
</script>

<%
done:
  Send_BodyEnd()
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  MyCommon.Close_LogixWH()
  MyCommon = Nothing
  Logix = Nothing
%>
