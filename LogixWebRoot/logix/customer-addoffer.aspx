<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%
  ' *****************************************************************************
  ' * FILENAME: customer-addoffer.aspx 
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
  Dim rst3 As DataTable
  Dim row3 As DataRow
  Dim rst4 As DataTable
  Dim rstPts As DataTable
  Dim rstSV As DataTable
  Dim rstAccum As DataTable
  Dim rstOffers As DataTable = Nothing
  Dim rowCount As Integer
  Dim OfferGroupType As Integer
  Dim CurrentOffers As String = ""
  Dim CustomerPK As Long
  Dim CardPK As Long
  Dim ExtCardID As String = ""
  Dim FirstName As String = ""
  Dim MiddleName As String = ""
  Dim LastName As String = ""
  Dim FullName As String = ""
  Dim ExtCustomerID As String = ""
  Dim ProgramID As String
  Dim ProgramName As String
  Dim Edit As Boolean
  Dim TotalRedeemCt As Integer = 0
  Dim TotalRedeemAmt As Double = 0.0
  Dim CustExtID As String = ""
  Dim i As Integer = 0
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
  Dim CustExtIdList As String = ""
  Dim Shaded As String = " class=""shaded"""
  Dim HasSearchResults As Boolean = False
  Dim FullAddress As String = ""
  Dim CustomerTypeID As Integer = 0
  Dim ExcludedFromOffer As Boolean
  Dim Employee As Integer
  Dim OfferID As Integer = 0
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
  Dim statusMessage As String = ""
  Dim Handheld As Boolean = False
  Dim HistoryText As String = ""
  Dim ExtractedCustomerGroupID As Long
  Dim TestCard As Boolean = False
  Dim OfferRowCount As Integer = 0
  Dim CgXml As String = ""
  Dim reader As SqlDataReader = Nothing
  Dim dtAddOffers As DataTable = Nothing
  Dim dtAssigned As DataTable = Nothing
  Dim sortedRows() As DataRow = Nothing
  Dim ColValues(11) As Object
  Dim RSCount As Integer = 1
  Dim OfferStatusCode As Copient.LogixInc.STATUS_FLAGS
  Dim OfferStatus As String = ""
  Dim StatusTable As New Hashtable(200)
  Dim SelectedOffer As Long = 0
  Dim SelectedOfferName As String = ""
  Dim Fields As New Copient.CommonInc.ActivityLogFields
  Dim AssocOffers(-1) As Copient.CommonInc.ActivityLink
  Dim SessionID As String = ""
  
  ' default urls for links from this page
  Dim URLOfferSum As String = "offer-sum.aspx"
  Dim URLCPEOfferSum As String = "CPEoffer-sum.aspx"
  Dim URLcgroupedit As String = "cgroup-edit.aspx"
  Dim URLpointedit As String = "point-edit.aspx"
  Dim URLtrackBack As String = ""
  Dim inCardNumber As String = ""
  ' tack on the customercare remote links if needed
  Dim extraLink As String = ""
  
  Dim UserRoleIDs() As Integer
  Dim RoleMatch As Boolean = False
  Dim x As Integer = 0
  Dim LastOfferID As Long = 0
  Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249)="1",True,False)
  Dim conditionalQuery = String.Empty
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "customer-addoffer.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  
  If Session("SessionID") IsNot Nothing AndAlso Session("SessionID").ToString.Trim <> "" Then
    SessionID = Session("SessionID").ToString.Trim
    CopientNotes &= "SessionID: " & SessionID
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
  If CardPK > 0 Then
    ExtCardID = MyLookup.FindExtCardIDFromCardPK(CardPK)
  End If
  If (CustomerPK > 0) Then
    MyCommon.QueryStr = "select CustomerTypeID, TestCard from Customers with (NoLock) where CustomerPK=" & MyCommon.Extract_Val(Request.QueryString("CustPK"))
    rst = MyCommon.LXS_Select
    If (rst.Rows.Count > 0) Then
      CustomerTypeID = MyCommon.NZ(rst.Rows(0).Item("CustomerTypeID"), 0)
      TestCard = MyCommon.NZ(rst.Rows(0).Item("TestCard"), False)
    End If
  End If
  
  If Not (CustomerPK <> 0) Then
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "customer-inquiry.aspx")
  End If
  
  Send_HeadBegin("term.customer", "term.addoffer", CustomerPK)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
%>
<style type="text/css">
  #detailsbox
  {
    float: none;
    height: 100px;
    width: 650px;
    margin-left: 0;
    margin-top: 2px;
    padding: 0;
    border: 0;
    overflow-y: scroll;
  }
  #grey
  {
    background-color: #ffffff;
  }
  #detailsbox div
  {
    background-color: #ddddff;
    padding: 8px;
  }
</style>
<%
  Send_Scripts()
  Send_HeadEnd()
  
  'Before anything else, check if we're supposed to add someone to an offer
  If (Request.QueryString("AddOffer") = Copient.PhraseLib.Lookup("term.add", LanguageID)) Then
    'Find out whether this customer is a household or cardholder
    CustomerPK = MyCommon.Extract_Val(Request.QueryString("CustomerPK"))
    If (CustomerPK > 0) Then
      MyCommon.QueryStr = "select CustomerTypeID from Customers with (NoLock) where CustomerPK=" & MyCommon.Extract_Val(Request.QueryString("CustomerPK"))
      rst = MyCommon.LXS_Select
      If (rst.Rows.Count > 0) Then
        CustomerTypeID = MyCommon.NZ(rst.Rows(0).Item("CustomerTypeID"), 0)
      End If
    End If
    CustomerGroupIDs = Request.QueryString("CustomerGroupID").Split(",")
    For loopCtr = CustomerGroupIDs.GetLowerBound(0) To CustomerGroupIDs.GetUpperBound(0)
      'For each customer group, check its EditControlTypeID to see if it's permissible to add a customer to it.
      MyCommon.QueryStr = "select EditControlTypeID, RoleID from CustomerGroups with (NoLock) where CustomerGroupID=" & CustomerGroupIDs(loopCtr) & ";"
      rst = MyCommon.LRT_Select
      If rst.Rows.Count > 0 Then
        RoleMatch = False
        If (MyCommon.NZ(rst.Rows(0).Item("EditControlTypeID"), 0) = 3) Then 'removal is limited per role, so loop through the user's roles and see if there's a match to the customer group's RoleID 
          For x = 0 To UserRoleIDs.Length - 1
            If UserRoleIDs(x) = MyCommon.NZ(rst.Rows(0).Item("RoleID"), 0) Then
              RoleMatch = True
            End If
          Next
        End If
        If (MyCommon.NZ(rst.Rows(0).Item("EditControlTypeID"), 0) = 0 AndAlso Logix.UserRoles.AddCustomerToOffers) Or (MyCommon.NZ(rst.Rows(0).Item("EditControlTypeID"), 0) = 1) OrElse (MyCommon.NZ(rst.Rows(0).Item("EditControlTypeID"), 0) = 3 AndAlso RoleMatch) Then
          MyCommon.QueryStr = "dbo.pt_GroupMembership_Insert_ByPK"
          MyCommon.Open_LXSsp()
          MyCommon.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.NVarChar, 26).Value = CustomerPK
          MyCommon.LXSsp.Parameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = CustomerGroupIDs(loopCtr)
          MyCommon.LXSsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
          MyCommon.LXSsp.ExecuteNonQuery()
          MyCommon.Close_LXSsp()
        End If
      End If
      If loopCtr = CustomerGroupIDs.GetLowerBound(0) Then
        ExtractedCustomerGroupID = CustomerGroupIDs(loopCtr)
      End If
    Next
    ' Determine offers associated with the customer group to add to the history
        
    If(bEnableRestrictedAccessToUEOfferBuilder) Then
        conditionalQuery=GetRestrictedAccessToUEBuilderQuery(MyCommon,Logix,"I")
    End If
    OffersList = ""
    SelectedOffer = MyCommon.Extract_Val(Request.QueryString("SelectedOfferID"))
    MyCommon.QueryStr = "select distinct Name,O.OfferID from Offers as O with (NoLock) " & _
                        "left join OfferConditions as OC with (NoLock) on OC.OfferID=O.OfferID " & _
                        "where O.IsTemplate=0 and OC.ConditionTypeID=1 and OC.Linkid in (" & Request.QueryString("CustomerGroupID") & ") or OC.ExcludedID in (" & Request.QueryString("CustomerGroupID") & ") " & _
                        " union " & _
                        "select distinct I.IncentiveName,I.IncentiveID as OfferID " & _
                        "from CPE_IncentiveCustomerGroups ICG with (NoLock) " & _
                        "inner join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID = ICG.RewardOptionID " & _
                        "inner join CPE_Incentives I with (NoLock) on I.IncentiveID = RO.IncentiveID " & _
                        "where ICG.Deleted=0 and RO.Deleted=0 and I.Deleted=0 and I.IsTemplate=0 and ICG.CustomerGroupID IN (" & Request.QueryString("CustomerGroupID") & ") " 
    If(bEnableRestrictedAccessToUEOfferBuilder AndAlso Not String.IsNullOrEmpty(conditionalQuery)) Then MyCommon.QueryStr &=conditionalQuery & " "
    MyCommon.QueryStr &= " order by OfferID ASC;"
        
    rst2 = MyCommon.LRT_Select
    If rst2.Rows.Count > 0 Then
      ReDim AssocOffers(rst2.Rows.Count - 1)
      If rst2.Rows.Count = 1 Then
        OffersList = MyCommon.NZ(rst2.Rows(0).Item("OfferID"), 0)
        AssocOffers(0).LinkID = MyCommon.NZ(rst2.Rows(0).Item("OfferID"), 0)
        AssocOffers(0).LinkTypeID = 1 ' Offer link type
        AssocOffers(0).Selected = (AssocOffers(0).LinkID = SelectedOffer)
      ElseIf rst2.Rows.Count > 1 Then
        i = 1
        For Each row In rst2.Rows
          If i = 1 Then
            OffersList = MyCommon.NZ(row.Item("OfferID"), 0)
          Else
            OffersList = OffersList & ", " & MyCommon.NZ(row.Item("OfferID"), 0)
          End If
          AssocOffers(i - 1).LinkID = MyCommon.NZ(row.Item("OfferID"), 0)
          AssocOffers(i - 1).LinkTypeID = 1 ' Offer link type
          AssocOffers(i - 1).Selected = (AssocOffers(i - 1).LinkID = SelectedOffer)
          i = i + 1
        Next
      End If
      HistoryText = Copient.PhraseLib.Lookup("history.customer-add-offer", LanguageID) & " #" & Request.QueryString("CustomerGroupID") & " (" & OffersList & ")"
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
    Else
      HistoryText = Copient.PhraseLib.Lookup("history.customer-add-offer", LanguageID) & " #" & Request.QueryString("CustomerGroupID")
    End If

    ' log the addition of the offer and any associated offers
    Fields.ActivityTypeID = 25
    Fields.ActivitySubTypeID = 15
    Fields.LinkID = CustomerPK
    Fields.AdminUserID = AdminUserID
    Fields.Description = HistoryText
    Fields.LinkID2 = ExtractedCustomerGroupID
    Fields.AssociatedLinks = AssocOffers
    Fields.SessionID = SessionID
	Dim CardTypeDesc As String = "" 
	MyCommon.QueryStr = "select Description from CardTypes with (NoLock) where CustTypeID=" & CustomerTypeID & ";" 
  	rst = MyCommon.LXS_Select  
       If (rst.Rows.Count > 0) Then  
        CardTypeDesc = MyCommon.NZ(rst.Rows(0).Item("Description"), "")  
      End If  


  	MyCommon.Activity_Log3(Fields) 
  	MyCommon.Activity_Log(4, ExtractedCustomerGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.cgroup-add", LanguageID) & " " & ExtCardID & IIf      (CardTypeDesc <> "", " (" & CardTypeDesc & ")", "")) 
     

    statusMessage = Copient.PhraseLib.Lookup("customer-report.addoffer", LanguageID)
    If SelectedOffer > 0 Then
      MyCommon.QueryStr = "select IncentiveName,buy.ExternalBuyerId as BuyerID from CPE_Incentives as I  with (NoLock) left join Buyers as buy with (NoLock) on buy.buyerid = I.buyerID where IncentiveID = " & SelectedOffer
      rst = MyCommon.LRT_Select
      If rst.Rows.Count > 0 Then
                  Dim Name As String=""
                                 If (MyCommon.Fetch_UE_SystemOption(168) = "1" AndAlso MyCommon.NZ(rst.Rows(0).Item("BuyerID"), "") <> "") Then
                                    Name = "Buyer " + rst.Rows(0).Item("BuyerID").ToString() + " - " + MyCommon.SplitNonSpacedString(rst.Rows(0).Item("IncentiveName"), 25).ToString()
                                Else
                                    Name = MyCommon.NZ(MyCommon.SplitNonSpacedString(rst.Rows(0).Item("IncentiveName"), 25), Copient.PhraseLib.Lookup("term.unknown", LanguageID))
                                End If
        SelectedOfferName = Name
        statusMessage &= " (" & SelectedOfferName & ")"
      End If
    End If
  ElseIf (Request.QueryString("Search") = Copient.PhraseLib.Lookup("term.search", LanguageID) OrElse Request.QueryString("searchMode") = Copient.PhraseLib.Lookup("term.search", LanguageID)) Then
    offerCt = 0
    
    MyCommon.QueryStr = "select CustomerGroupID from GroupMembership with (NoLock) where CustomerPK=" & CustomerPK & " and Deleted=0;"
    rst = MyCommon.LXS_Select()
    
    CgXml = "<customergroups><id>1</id><id>2</id>"
    rowCount = rst.Rows.Count
    If rowCount > 0 Then
      For Each row In rst.Rows
        CgXml &= "<id>" & MyCommon.NZ(row.Item("CustomerGroupID"), "") & "</id>"
      Next
    End If
    CgXml &= "</customergroups>"
    
    MyCommon.QueryStr = "dbo.pa_CustomerOffersAdd"
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add("@cgXml", SqlDbType.Xml).Value = CgXml
    MyCommon.LRTsp.Parameters.Add("@IsEmployee", SqlDbType.Bit).Value = Employee
    MyCommon.LRTsp.Parameters.Add("@Filter", SqlDbType.NVarChar).Value = IIf(Request.QueryString("functionradio") = "1", "", "%") & MyCommon.Parse_Quotes(Request.QueryString("functioninput")) & "%"
    dtAddOffers = MyCommon.LRTsp_select
    MyCommon.Close_LRTsp()
    
    'Removing transformed offers
     If (bEnableRestrictedAccessToUEOfferBuilder) Then
         RemoveTransformedOffers(dtAddOffers,MyCommon)
     End If
        
    ' sort the Assigned offers
    sortedRows = dtAddOffers.Select("", SortText & " " & SortDirection)
    StatusTable = LoadOfferStatuses(sortedRows, MyCommon, Logix)
    If sortedRows.Length > 0 Then
      For Each row2 In sortedRows
        OfferStatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_UNKNOWN
        OfferStatus = StatusTable.Item(MyCommon.NZ(row2.Item("OfferID"), "0").ToString)
        If (OfferStatus IsNot Nothing) Then
          OfferStatusCode = OfferStatus
        End If
        If (TestCard = False AndAlso OfferStatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_ACTIVE) _
            OrElse (TestCard = True AndAlso OfferStatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_TESTING AndAlso MyCommon.Fetch_SystemOption(88) = "1") Then
          ' these are the only valid rows for add offer
          OfferRowCount += 1
        Else
          row2.Delete()
        End If
      Next
    End If
    If OfferRowCount = 0 Then
      infoMessage = Copient.PhraseLib.Lookup("customer-inquiry.SearchReturnedNoResults", LanguageID)
      sortedRows = Nothing
    End If
  End If
  
  If (MyCommon.Extract_Val(Request.QueryString("custPK")) > 0 Or (MyCommon.Extract_Val(Request.QueryString("CustomerPK"))) Or _
      (Request.QueryString("searchterms") <> "" And _
      (Request.QueryString("search") <> "" Or Request.QueryString("searchPressed") <> "")) Or _
      inCardNumber <> "" _
      ) Then
    ' someone wants to search for a customer.  First lets get their primary key from our database
    If (MyCommon.Extract_Val(Request.QueryString("custPK")) > 0 Or (MyCommon.Extract_Val(Request.QueryString("CustomerPK")) > 0)) Then
      If (MyCommon.Extract_Val(Request.QueryString("custPK")) > 0) Then
        CustomerPK = MyCommon.Extract_Val(Request.QueryString("custPK"))
      ElseIf (MyCommon.Extract_Val(Request.QueryString("CustomerPK")) > 0) Then
        CustomerPK = MyCommon.Extract_Val(Request.QueryString("CustomerPK"))
      End If
      
      MyCommon.QueryStr = "select C.CustomerPK, C.CustomerTypeID, C.HHPK, C.FirstName, C.MiddleName, C.LastName, " & _
                          "CE.Address, CE.City, CE.State, CE.Zip, CE.PhoneAsEntered, CE.email " & _
                          "from Customers C with (NoLock) left join CustomerExt CE with (NoLock) on CE.CustomerPK = C.CustomerPK " & _
                          "left join Customers C2 with (NoLock) on C2.CustomerPK = C.HHPK " & _
                          "where C.CustomerPK = " & CustomerPK
    Else
      ' IF the page was called from an outside application set ClientUserID1 to the outside passed in value
      If (inCardNumber <> "" And Request.QueryString("mode") = "summary") Then
        ClientUserID1 = MyCommon.Pad_ExtCardID(inCardNumber, CustomerTypeID)
        searchterms = Request.QueryString("searchterms")
        MyCommon.QueryStr = "select C.CustomerPK, C.CustomerTypeID, C.HHPK, C.FirstName, C.MiddleName, C.LastName, " & _
                            "CE.Address, CE.City, CE.State, CE.Zip, CE.PhoneAsEntered, CE.email " & _
                            "from Customers C with (NoLock) left join CustomerExt CE with (NoLock) on CE.CustomerPK = C.CustomerPK " & _
                            "left join Customers C2 with (NoLock) on C2.CustomerPK=C.HHPK " & _
                            "where C.PrimaryExtID = '" & MyCommon.Parse_Quotes(ClientUserID1) & "';"
      End If
      If (Request.QueryString("searchterms") <> "" And ClientUserID1 = "") Then
        ClientUserID1 = MyCommon.Pad_ExtCardID(Request.QueryString("searchterms"), CustomerTypeID)
                MyCommon.QueryStr = "select C.CustomerPK, C.CustomerTypeID, C.HHPK, C.FirstName, C.MiddleName, C.LastName, " & _
                                          "CE.Address, CE.City, CE.State, CE.Zip, CE.PhoneAsEntered, CE.email " & _
                                          "from Customers C with (NoLock) left join CustomerExt CE with (NoLock) on CE.CustomerPK = C.CustomerPK " & _
                                          "left join Customers C2 with (NoLock) on C2.CustomerPK = C.HHPK " & _
                                          "where C.PrimaryExtID = '" & MyCryptLib.SQL_StringEncrypt(ClientUserID1) & "' or CE.PhoneDigitsOnly = '" & MyCryptLib.SQL_StringEncrypt(MyCommon.DigitsOnly(Request.QueryString("searchterms"))) & _
                                          "' or CE.email = '" & MyCryptLib.SQL_StringEncrypt(MyCommon.Parse_Quotes(Request.QueryString("searchterms"))) & "' or C.LastName like '%" & MyCommon.Parse_Quotes(Request.QueryString("searchterms")) & "%';"
      End If
    End If
    rstResults = MyCommon.LXS_Select
    
    If (rstResults.Rows.Count = 1) Then
      ' ok we found a primary key for the external id provided
      CustomerPK = MyCommon.NZ(rstResults.Rows(0).Item("CustomerPK"), 0)
      FirstName = MyCommon.NZ(rstResults.Rows(0).Item("FirstName"), "")
      MiddleName = MyCommon.NZ(rstResults.Rows(0).Item("MiddleName"), "")
      LastName = MyCommon.NZ(rstResults.Rows(0).Item("LastName"), "")
      IsHouseholdID = (MyCommon.NZ(rstResults.Rows(0).Item("CustomerTypeID"), 0) = 1)
    Else
      infoMessage = "" & Copient.PhraseLib.Lookup("customer.notfound", LanguageID) & ""
      infoMessage = infoMessage & " <a href=""customer-addoffer.aspx?mode=add&Search=Search" & extraLink & "&searchterms=" & Request.QueryString("searchterms") & """>" & Copient.PhraseLib.Lookup("term.add", LanguageID) & "</a>"
    End If
    
  End If
  
  UnknownPhrase = Copient.PhraseLib.Lookup("term.unknown", LanguageID)
  
  Send_BodyBegin(2)
  
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
  document.getElementById('AddPromoLink').href = 'XMLFeeds.aspx?AddOffer=' + offerID + '&CustPK=' + custPK + '&CardPK=' + cardPK + '&AdminUserID=<%Sendb(AdminUserID)%>&height=300&width=300&LanguageID=<%Sendb(LanguageID)%>';
  $('#AddPromoLink').click();
}




function showOfferPage(position) {
  //var startPos = ((page-1) * OFFER_ROWS_SHOWN) + 1;
  var trElem = null;
  
  // hide last page shown
  if (lastStartPos > -1) {
    for (var i=lastStartPos; i < lastStartPos + OFFER_ROWS_SHOWN; i++) {
      trElem = document.getElementById("trOffer" + i);
      if (trElem != null) trElem.style.display = 'none';
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
  
  if (lastStartPos + OFFER_ROWS_SHOWN >= recCt) {
    if (elemNextOn != null)  elemNextOn.style.display = "none";
    if (elemLastOn != null)  elemLastOn.style.display = "none";
    if (elemNextOff != null) elemNextOff.style.display = "";          
    if (elemLastOff != null) elemLastOff.style.display = "";          
  } else {
    if (elemNextOn != null)  elemNextOn.style.display = "";
    if (elemLastOn != null)  elemLastOn.style.display = "";
    if (elemNextOff != null) elemNextOff.style.display = "none";          
    if (elemLastOff != null) elemLastOff.style.display = "none";          
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
  
  if (document.getElementById("functionselect") == null) { return; }
  
  document.getElementById("functionselect").size = "15";
  
  // Set references to the form elements
  selectObj = document.forms['mainform'].functionselect;
  textObj = document.forms['mainform'].functioninput;
  
  // Remember the function list length for loop speedup
  if (functionlist != null) {
    functionListLength = functionlist.length;
  } else {
    functionListLength = 0;
  }
  
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
    document.getElementById('detailsbox').innerHTML = '<div class=\"grey\"><% Sendb(Copient.PhraseLib.Lookup("term.none", LanguageID))%><\/div>';
  }
  
  // Loop through the array and re-add matching options
  numShown = 0;
  if (functionlist != null && textObj.value != "") { 
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
  }
  // When options list whittled to one, select that entry
  if(selectObj.length == 1) {
    selectObj.options[0].selected = true;
    if (document.getElementById('detailsbox') != null) {
      if (descs[pointerlist[0]] == '') {
        document.getElementById('detailsbox').innerHTML = '<div class=\"grey\"><% Sendb(Copient.PhraseLib.Lookup("term.none", LanguageID))%><\/div>';
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
  var offerTerms = '';
  
  if (elemOfferTerms != null) { offerTerms = elemOfferTerms.value; }
  <%
    Dim strTerms = Request.QueryString("searchterms")
    If (strTerms <> "") Then
      strTerms = strTerms.Replace("'", "\'")
      strTerms = strTerms.Replace("""", "\""")
    End If
   %>
  var qryStr = 'customer-addoffer.aspx?search=Search&searchterms=<%Sendb(strTerms)%>&CustPK=<%Sendb(CustomerPK)%><%Sendb(IIf(CardPK > 0, "&CardPK=" & CardPK, ""))%>&offerSearch=Search&offerterms=' + offerTerms + '#h01';
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
    } else {
      submitSearch();
    }
    return false;
  } else if (elem != null && elem.options.length > 0 && key == 40) {
    elem.options[0].selected = true;
    elem.focus();
    if (e && e.preventDefault) {
      e.preventDefault(); // DOM style
    }
    e.returnValue = false;
    return false;
  } else if (key == 13) {
    submitSearch();
  }
  return true;
}

function handleSelectKeyDown(e) {
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

function setCaretPosition(elemId, caretPos) {
  var elem = document.getElementById(elemId);
  
  if(elem != null) {
    if(elem.createTextRange) {
      var range = elem.createTextRange();
      range.move('character', caretPos);
      range.select();
    } else {
      if(elem.selectionStart) {
        elem.focus();
        elem.setSelectionRange(caretPos, caretPos);
      }
      else
        elem.focus();
    }
  }
}

function submitSearch() {
  var elem = document.getElementById('searchMode');
  
  if (elem != null) {
    elem.value = '<%Sendb(Copient.PhraseLib.Lookup("term.search", LanguageID)) %>';
  }
  document.mainform.submit()
}

function ChangeParentDocument() {
  <%  If (Request.QueryString("AddOffer") = Copient.PhraseLib.Lookup("term.add", LanguageID)) Then %>
  if (opener != null && !opener.closed) {
    opener.location = 'customer-offers.aspx?CustomerPK=<%Sendb(CustomerPK)%>&CustPK=<%Sendb(CustomerPK)%><%Sendb(IIf(CardPK > 0, "&CardPK=" & CardPK, ""))%>&offerterms=<%Sendb(Server.UrlEncode(SelectedOfferName))%>&Favorite=0';  
  }
  <% End If %>
}
//-->
</script>
<script type="text/javascript" src="../javascript/jquery.min.js"></script>
<script type="text/javascript" src="../javascript/thickbox.js"></script>
<form id="mainform" name="mainform" action="customer-addoffer.aspx" autocomplete="off">
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
          Sendb(Copient.PhraseLib.Lookup("term.customercard", LanguageID) & " #" & ExtCardID)
        End If
      End If
      If (FirstName <> "" OrElse LastName <> "") Then
        FullName = IIf(FirstName <> "", FirstName & " ", "")
        FullName &= IIf(MiddleName <> "", Left(MiddleName, 1) & ". ", "")
        FullName &= IIf(LastName <> "", LastName, "")
        Sendb(": " & MyCommon.TruncateString(FullName, 30))
      End If
      If (restrictLinks AndAlso URLtrackBack <> "") Then
        Send(" - <a href=""" & URLtrackBack & """>" & Copient.PhraseLib.Lookup("customer-inquiry.return", LanguageID) & "</a>")
      End If
    %>
  </h1>
  <div id="controls" <% Sendb(IIf(restrictLinks AndAlso URLtrackBack <> "", " style=""width:115px;""", "")) %>>
    <%
      If (sortedRows IsNot Nothing AndAlso OfferRowCount > 0) Then
        Send("<input type=""button"" class=""regular"" id=""pselect"" name=""pselect"" value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """ onclick=""if (document.mainform.functionselect.selectedIndex > -1) { addOffer(" & CustomerPK & ", " & CardPK & ", document.getElementById('functionselect').value); }"" />")
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
    If (statusMessage <> "") Then
      Send("<div id=""infobar"" class=""green-background"">" & statusMessage & "</div>")
    End If
    If (infoMessage <> "") Then
      Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
    End If
    Send("<input type=""hidden"" id=""CustomerPK"" name=""CustomerPK"" value=""" & CustomerPK & """ />")
    If CardPK > 0 Then
      Send("<input type=""hidden"" id=""CardPK"" name=""CardPK"" value=""" & CardPK & """ />")
    End If
    Send("<input type=""hidden"" id=""searchMode"" name=""searchMode"" value="""" />")
    If (Request.QueryString("mode") = "summary") Then
      Send("<input type=""hidden"" id=""mode"" name=""mode"" value=""summary"" />")
      Send("<input type=""hidden"" id=""exiturl"" name=""exiturl"" value=""" & URLtrackBack & """ />")
    End If
  %>
  <div id="column">
    <%
      If (CustomerPK > 0) Then
        Send("<div class=""box"" id=""availableoffers"">")
        Send("  <h2>")
        Send("    <span>" & IIf(sortedRows Is Nothing, Copient.PhraseLib.Lookup("customer-addoffer.EnterOfferSearchTerms", LanguageID), Copient.PhraseLib.Lookup("customer-addoffer.AddCustomer", LanguageID)) & "</span>")
        Send("  </h2>")
        Send("<input type=""radio"" id=""functionradio1"" name=""functionradio""" & IIf(Request.QueryString("functionradio") = "2", "", "checked=""checked""") & " value=""1"" /><label for=""functionradio1"">" & Copient.PhraseLib.Lookup("term.startingwith", LanguageID) & "</label>")
        Send("<input type=""radio"" id=""functionradio2"" name=""functionradio""" & IIf(Request.QueryString("functionradio") = "2", "checked=""checked""", "") & " value=""2"" /><label for=""functionradio2"">" & Copient.PhraseLib.Lookup("term.containing", LanguageID) & "</label><br />")
        Send("<input type=""text"" class=""longest"" onkeyup=""handleKeyUp(200);"" onkeydown=""return handleKeyDown(event);"" id=""functioninput"" name=""functioninput"" maxlength=""100"" value=""" & Request.QueryString("functioninput") & """ />")
        Send("<input type=""button"" class=""regular"" name=""search"" id=""search"" value=""" & Copient.PhraseLib.Lookup("term.search", LanguageID) & """ onclick=""submitSearch();"" />")
        Send("<script type=""text/javascript"">")
        If sortedRows IsNot Nothing AndAlso OfferRowCount > 0 Then
          Send("    var descs = new Array(" & OfferRowCount & ");")
        Else
          Send("    var descs = new Array(" & 0 & ");")
        End If
        n = 0
        If sortedRows IsNot Nothing Then
          LastOfferID = -1
          For Each row3 In sortedRows
            If row3.RowState <> DataRowState.Deleted AndAlso  row3.RowState <> DataRowState.Detached  Then
              RoleMatch = False
              If (MyCommon.NZ(row3.Item("EditControlTypeID"), 0) = 3) Then 'removal is limited per role, so loop through the user's roles and see if there's a match to the customer group's RoleID 
                For x = 0 To UserRoleIDs.Length - 1
                  If UserRoleIDs(x) = MyCommon.NZ(row3.Item("RoleID"), 0) Then
                    RoleMatch = True
                  End If
                Next
              End If
              If (MyCommon.NZ(row3.Item("EditControlTypeID"), 0) = 0 AndAlso Logix.UserRoles.AddCustomerToOffers) Or (MyCommon.NZ(row3.Item("EditControlTypeID"), 0) = 1) OrElse (MyCommon.NZ(row3.Item("EditControlTypeID"), 0) = 3 AndAlso RoleMatch) Then
                If MyCommon.NZ(row3.Item("OfferID"), -1) <> LastOfferID Then
                  Send("    descs[" & n & "] = '" & MyCommon.SplitNonSpacedString(CleanDescription(MyCommon.NZ(row3.Item("Description"), "").ToString), 25) & "';")
                  LastOfferID = MyCommon.NZ(row3.Item("OfferID"), -1)
                  n = n + 1
                End If
              End If
            End If
          Next
        End If
        Send("function setInnerHTML (index) {")
        Send("  if (document.getElementById('detailsbox')) {")
        Send("    if(descs[pointerlist[index]] == '') {")
        Send("      document.getElementById('detailsbox').innerHTML = '<div class=""grey"">" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "<\/div>';")
        Send("    } else {")
        Send("      document.getElementById('detailsbox').innerHTML = '<div>' + descs[pointerlist[index]] + '<\/div>';")
        Send("    }")
        Send("  }")
        Send("}")
        Send("</script>")
        If sortedRows IsNot Nothing Then
          Send("<select id=""functionselect"" name=""functionselect"" size=""15"" style=""width: 650px;"" onclick=""handleSelectClick();"" onkeydown=""return handleSelectKeyDown(event);"" ondblclick=""if (this.selectedIndex > -1) { addOffer(" & CustomerPK & ", " & CardPK & ", document.getElementById('functionselect').value); }"" onchange=""setInnerHTML(selectedIndex);"">")
          If sortedRows IsNot Nothing Then
            LastOfferID = -1
            For Each row3 In sortedRows
              If row3.RowState <> DataRowState.Deleted AndAlso  row3.RowState <> DataRowState.Detached Then
                RoleMatch = False
                If (MyCommon.NZ(row3.Item("EditControlTypeID"), 0) = 3) Then 'removal is limited per role, so loop through the user's roles and see if there's a match to the customer group's RoleID 
                  For x = 0 To UserRoleIDs.Length - 1
                    If UserRoleIDs(x) = MyCommon.NZ(row3.Item("RoleID"), 0) Then
                      RoleMatch = True
                    End If
                  Next
                End If
                Dim Name As String=""
                    If (MyCommon.Fetch_UE_SystemOption(168) = "1" AndAlso MyCommon.NZ(row3.Item("BuyerID"), "") <> "") Then
                    Name = "Buyer " + row3.Item("BuyerID").ToString() + " - " + MyCommon.SplitNonSpacedString(row3.Item("Name"), 25).ToString()
                Else
                    Name = MyCommon.NZ(MyCommon.SplitNonSpacedString(row3.Item("Name"), 25), Copient.PhraseLib.Lookup("term.unknown", LanguageID))
                End If
                If (MyCommon.NZ(row3.Item("EditControlTypeID"), 0) = 0 AndAlso Logix.UserRoles.AddCustomerToOffers) Or (MyCommon.NZ(row3.Item("EditControlTypeID"), 0) = 1) OrElse (MyCommon.NZ(row3.Item("EditControlTypeID"), 0) = 3 AndAlso RoleMatch) Then
                  If MyCommon.NZ(row3.Item("OfferID"), -1) <> LastOfferID Then
                    Send("  <option value=""" & MyCommon.NZ(row3.Item("OfferID"), -1) & """>" & Name & "</option>")
                    LastOfferID = MyCommon.NZ(row3.Item("OfferID"), -1)
                  End If
                End If
              End If
            Next
          End If
          Send("</select>")
          Send("<br />")
          Send("<br class=""half"" />")
          Send("<h3>")
          Send("  " & Copient.PhraseLib.Lookup("term.description", LanguageID))
          Send("</h3>")
          Send("<div class=""detailsbox"" id=""detailsbox"">")
          Send("  <div class=""grey"">")
          Send("    " & Copient.PhraseLib.Lookup("term.none", LanguageID))
          Send("  </div>")
          Send("</div>")
          Send("<br clear=""left"" />")
          Send("<br class=""half"" />")
        End If
        Send("<a href=""#"" id=""AddPromoLink"" name=""AddPromoLink"" title=""" & Copient.PhraseLib.Lookup("term.alert", LanguageID) & """ class=""thickbox""></a>")
        Send("<hr class=""hidden"" />")
        Send("</div>")
      End If
    %>
  </div>
  <br clear="all" />
</div>
</form>
<%
  Send("<script type=""text/javascript"">")
  If (Request.QueryString("refresh") = "") Then
    Send("  if (document.mainform.searchterms != null) document.mainform.searchterms.focus();")
  End If
  Send("  if (document.getElementById(""paginator"") != null && document.getElementById(""pageIter"")!=null ) { ")
  Send("    document.getElementById(""paginator"").innerHTML = document.getElementById(""pageIter"").innerHTML;")
  Send("    document.getElementById(""pageIter"").innerHTML = """ & """")
  Send("  }")
  Send("  showOfferPage(1);")
  Send("</script>")
%>
<script type="text/javascript">
// This is the javascript array holding the function list
// The PrintJavascriptArray ASP function can be used to print this array.
<%
  Dim elementBuf as new StringBuilder()
  
  If (sortedRows IsNot Nothing AndAlso OfferRowCount >0)
    LastOfferID = -1
    Sendb("var functionlist = Array(")
    For Each row In sortedRows
      If row.RowState <> DataRowState.Deleted AndAlso  row.RowState <> DataRowState.Detached Then
        RoleMatch = False
        If (MyCommon.NZ(row.Item("EditControlTypeID"), 0) = 3) Then
          For x = 0 To UserRoleIDs.Length - 1
            If UserRoleIDs(x) = MyCommon.NZ(row.Item("RoleID"), 0) Then
              RoleMatch = True
            End If
          Next
        End If
          Dim Name As String=""
                                 If (MyCommon.Fetch_UE_SystemOption(168) = "1" AndAlso MyCommon.NZ(row.Item("BuyerID"), "") <> "") Then
                                    Name = "Buyer " + row.Item("BuyerID").ToString() + " - " + MyCommon.SplitNonSpacedString(row.Item("Name"), 25).ToString()
                                Else
                                    Name = MyCommon.NZ(MyCommon.SplitNonSpacedString(row.Item("Name"), 25), Copient.PhraseLib.Lookup("term.unknown", LanguageID))
                                End If
        If (MyCommon.NZ(row.Item("EditControlTypeID"), 0) = 0 AndAlso Logix.UserRoles.AddCustomerToOffers) Or (MyCommon.NZ(row.Item("EditControlTypeID"), 0) = 1) OrElse (MyCommon.NZ(row.Item("EditControlTypeID"), 0) = 3 AndAlso RoleMatch) Then
          If (MyCommon.NZ(row.Item("OfferID"), -1) <> LastOfferID) Then
            Sendb("""" & Name.ToString().Replace("""", "\""") & """,")
            LastOfferID = MyCommon.NZ(row.Item("OfferID"), -1)
          End If
        End If
      End If
    Next
    Sendb(""""");")
    LastOfferID = -1
    Sendb("var vallist = Array(")
    For Each row In sortedRows
      If row.RowState <> DataRowState.Deleted AndAlso  row.RowState <> DataRowState.Detached Then
        RoleMatch = False
        If (MyCommon.NZ(row.Item("EditControlTypeID"), 0) = 3) Then
          For x = 0 To UserRoleIDs.Length - 1
            If UserRoleIDs(x) = MyCommon.NZ(row.Item("RoleID"), 0) Then
              RoleMatch = True
            End If
          Next
        End If
        If (MyCommon.NZ(row.Item("EditControlTypeID"), 0) = 0 AndAlso Logix.UserRoles.AddCustomerToOffers) Or (MyCommon.NZ(row.Item("EditControlTypeID"), 0) = 1) OrElse (MyCommon.NZ(row.Item("EditControlTypeID"), 0) = 3 AndAlso RoleMatch) Then
          If (MyCommon.NZ(row.Item("OfferID"), -1) <> LastOfferID) Then
            Sendb("""" & MyCommon.NZ(row.item("OfferID"), -1) & """,")
            LastOfferID = MyCommon.NZ(row.Item("OfferID"), -1)
          End If
        End If
      End If
    Next
    Sendb(""""");")
    LastOfferID = -1
    Sendb("var pointerlist = Array(")
    i = 0
    For Each row In sortedRows
      If row.RowState <> DataRowState.Deleted AndAlso  row.RowState <> DataRowState.Detached Then
        RoleMatch = False
        If (MyCommon.NZ(row.Item("EditControlTypeID"), 0) = 3) Then
          For x = 0 To UserRoleIDs.Length - 1
            If UserRoleIDs(x) = MyCommon.NZ(row.Item("RoleID"), 0) Then
              RoleMatch = True
            End If
          Next
        End If
        If (MyCommon.NZ(row.Item("EditControlTypeID"), 0) = 0 AndAlso Logix.UserRoles.AddCustomerToOffers) Or (MyCommon.NZ(row.Item("EditControlTypeID"), 0) = 1) OrElse (MyCommon.NZ(row.Item("EditControlTypeID"), 0) = 3 AndAlso RoleMatch) Then
          If (MyCommon.NZ(row.Item("OfferID"), -1) <> LastOfferID) Then
            Sendb("""" & i & """,")
            LastOfferID = MyCommon.NZ(row.Item("OfferID"), -1)
            i += 1
          End If
        End If
      End If
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
%>
</script>
<script runat="server">
  Function LoadOfferStatuses(ByVal rows() As DataRow, ByRef MyCommon As Copient.CommonInc, ByRef Logix As Copient.LogixInc) As Hashtable
    Dim Statuses As New Hashtable(200)
    Dim i, ct As Integer
    Dim OfferList() As String = Nothing
    
    ct = rows.Length
    If (ct > 0) Then
      ReDim OfferList(ct - 1)
      For i = 0 To ct - 1
        OfferList(i) = MyCommon.NZ(rows(i).Item("OfferID"), "0")
      Next
      Statuses = Logix.GetStatusForOffers(OfferList, LanguageID)
    End If
    
    Return Statuses
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
<script type="text/javascript">
<%
  If (Request.QueryString("functioninput") <> "") Then 
     Sendb("setCaretPosition('functioninput', " & Request.QueryString("functioninput").Length & ");")
   Else
     Sendb("setCaretPosition('functioninput', 0);")
   End If
%>
</script>
<%
  'If MyCommon.Fetch_SystemOption(75) Then
  '  If (CustomerPK > 0 And Logix.UserRoles.ViewCustomerNotes) Then
  '    Send_Notes(4, CustomerPK, AdminUserID)
  '  End If
  'End If
done:
  Send_BodyEnd()
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  MyCommon = Nothing
  Logix = Nothing
%>
