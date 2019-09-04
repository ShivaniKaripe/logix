<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%
  ' *****************************************************************************
  ' * FILENAME: customer-offers.aspx 
  ' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' * Copyright © 2002 - 2013.  All rights reserved by:
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
  Dim MyLookup As New Copient.CustomerLookup
  Dim Logix As New Copient.LogixInc
    Dim rstResults As DataTable = Nothing
    Dim MyCryptLib As New Copient.CryptLib
  Dim rst As DataTable
  Dim row As DataRow
  Dim rst2 As DataTable
  Dim row2 As DataRow
  Dim rst4 As DataTable
  Dim rowCount As Integer
  Dim CustomerPK As Long
  Dim CardPK As Long = 0
  Dim ExtCardID As String = ""
  Dim FullName As String = ""
  Dim i As Integer = 0
  Dim j As Integer = 0
  Dim r As Integer = 0
  Dim offerCt As Integer = 0
  Dim IsPtsOffer As Boolean = False
  Dim IsSVOffer As Boolean = False
  Dim IsAccumOffer As Boolean = False
  Dim UnknownPhrase As String = ""
  Dim DisabledPtsAdj As String = ""
  Dim DisabledSVAdj As String = ""
  Dim DisabledAccumAdj As String = ""
  Dim SortText As String = "Name"
  Dim SortDirection As String = "ASC"
  Dim OfferTerms As String = ""
  Dim IsHouseholdID As Boolean = False
  Dim HHPK As Integer = 0
  Dim CustomerTypeID As Integer = 0
  Dim Employee As Integer = 0
  Dim TestCard As Boolean = False
  Dim ClientUserID1 As String = ""
  Dim searchterms As String = ""
  Dim restrictLinks As Boolean = False
  Dim OffersList As String = ""
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim HistoryText As String = ""
  Dim PrevOfferID As Long = 0
  Dim ExcludeExpired As Boolean = False
  Dim Targeted As Boolean = False
  Dim Favorite As Integer
  Dim CgXml As String = ""
  Dim reader As SqlDataReader = Nothing
  Dim dtAddOffers As DataTable = Nothing
  Dim dtAssigned As DataTable = Nothing
  Dim sortedRows() As DataRow = Nothing
  Dim OfferStatusCode As Copient.LogixInc.STATUS_FLAGS
  Dim OfferStatus As String = ""
  Dim StatusTable As New Hashtable(200)
  Dim OfferInQuo As Integer = 0
  Dim Fields As New Copient.CommonInc.ActivityLogFields
  Dim AssocLinks(-1) As Copient.CommonInc.ActivityLink
  Dim SessionID As String = ""
  Dim CGList As String = "-99"
  Dim OfferStart As Integer = 0
  Dim index As Integer = 0
  Dim InstalledEngines(-1) As Integer
  Dim OfferIsAirMiles As Boolean = False
  
  Dim PointsTable As New Hashtable()
  Dim pointRow As DataRow
  Dim pointdt As DataTable
  Dim PointsProgram As Integer = 0
  Dim HasPoints As Boolean = False
  Dim pointTemp, amountTemp As Integer
  
  Dim SVTable As New Hashtable()
  Dim svRow As DataRow
  Dim svdt As DataTable
  Dim SVProgram As Integer = 0
  Dim HasSV As Boolean = False
  Dim ShowCreditOnly As Boolean = False
  Dim BalanceOffers As New Hashtable()
  Dim UnavailableCardTypeOffers As New Hashtable()
  Dim svTemp As Integer = 0
  
  Dim AccumTable As New Hashtable()
  Dim RewardOptionID As Integer = 0
  Dim acRow As DataRow
  Dim acdt As DataTable
  Dim HasAccum As Boolean = False
  Dim acTemp As Integer = 0
  
  ' default urls for links from this page
  Dim URLOfferSum As String = "offer-sum.aspx"
  Dim URLCPEOfferSum As String = "CPEoffer-sum.aspx"
  Dim URLUEOfferSum As String = "UE/UEoffer-sum.aspx"
  Dim URLcgroupedit As String = "cgroup-edit.aspx"
  Dim URLpointedit As String = "point-edit.aspx"
  Dim URLWEBOfferSum As String = "web-offer-sum.aspx"
  Dim URLtrackBack As String = ""
  Dim inCardNumber As String = ""
  ' tack on the customercare remote links if needed
  Dim extraLink As String = ""
  
  Dim UserRoleIDs() As Integer
  Dim RoleMatch As Boolean = False
  Dim x As Integer = 0
  Dim CustDefaultView As Integer = 0
  Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249)="1",True,False)
  Dim conditionalQuery = String.Empty
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  

  Response.Expires = 0
  MyCommon.AppName = "customer-offers.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  If Session("SessionID") IsNot Nothing AndAlso Session("SessionID").ToString.Trim <> "" Then
    SessionID = Session("SessionID").ToString.Trim
    CopientNotes &= "SessionID: " & SessionID
  End If

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
  If CardPK > 0 Then
    ExtCardID = MyLookup.FindExtCardIDFromCardPK(CardPK)
  End If
  If (CustomerPK > 0) Then
    MyCommon.QueryStr = "select CustomerTypeID,Employee from Customers with (NoLock) where CustomerPK=" & MyCommon.Extract_Val(Request.QueryString("CustPK"))
    rst = MyCommon.LXS_Select
    If (rst.Rows.Count > 0) Then
      CustomerTypeID = rst.Rows(0).Item("CustomerTypeID")
      Employee = MyCommon.NZ(rst.Rows(0).Item("Employee"), 0)
    End If
  End If
  
  If (Request.QueryString("Favorite") = "0" OrElse Request.QueryString("Favorite") = "FALSE") Then
    Favorite = 0
  ElseIf (Request.QueryString("Favorite") = "1" OrElse Request.QueryString("Favorite") = "TRUE") Then
    Favorite = 1
    SortText = "Priority"
    SortDirection = "ASC"
  ElseIf (Request.QueryString("Favorite") = "2") Then
    Favorite = 0
    ExcludeExpired = True
  ElseIf (Request.QueryString("Favorite") = "5") Then
    Favorite = 0
    ExcludeExpired = False
    ShowCreditOnly = True
  ElseIf (Request.QueryString("Favorite") = "6") Then
    Favorite = 0
    ExcludeExpired = True
    ShowCreditOnly = False
    Targeted = True
  ElseIf Request.QueryString("Favorite") = "" Then
    'Read the option to set the default view to load  
    CustDefaultView = MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(109))
    If (CustDefaultView = 0) Then
      Favorite = 0
    ElseIf (CustDefaultView = 1) Then
      MyCommon.QueryStr = "select OfferID from AdminUserOffers where AdminUserID=" & AdminUserID & ";"
      rst = MyCommon.LRT_Select
      If rst.Rows.Count = 0 Then
        Favorite = 0
      Else
        Favorite = 1
        SortText = "Priority"
        SortDirection = "ASC"
      End If
    ElseIf (CustDefaultView = "2") Then
      Favorite = 0
      ExcludeExpired = True
    ElseIf (CustDefaultView = "5") Then
      Favorite = 0
      ExcludeExpired = False
      ShowCreditOnly = True
    ElseIf (CustDefaultView = 6) Then
      Favorite = 0
      ExcludeExpired = True
      ShowCreditOnly = False
      Targeted = True
    End If
  End If
  
  If (CustomerPK = 0) Then
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "customer-inquiry.aspx")
  End If
  
  'Use CustomerPK to find all the points groups that they have points in. Put results in a hashtable
  MyCommon.QueryStr = "select ProgramID, Amount, PromoVarID from Points where CustomerPK=" & CustomerPK & " and Amount > 0;"
  pointdt = MyCommon.LXS_Select()
  Dim ProgramDT As DataTable
  Dim PromoVarID As Double = 0
  If pointdt.Rows.Count > 0 Then
    For Each pointRow In pointdt.Rows
      pointTemp = MyCommon.Extract_Val(MyCommon.NZ(pointRow.Item("ProgramID"), 0))
      amountTemp = MyCommon.Extract_Val(MyCommon.NZ(pointRow.Item("Amount"), 0))
      PromoVarID = MyCommon.Extract_Val(MyCommon.NZ(pointRow.Item("PromoVarID"), 0))
      'Use the ProgramID and PromoVarID to make sure there is a points program associated to the record in the Points table.
      MyCommon.QueryStr = "select ProgramID from PointsPrograms where ProgramID=" & pointTemp & " and PromoVarID=" & PromoVarID & ";"
      ProgramDT = MyCommon.LRT_Select()
      If ProgramDT.Rows.Count > 0 Then
        If Not PointsTable.Contains(pointTemp) Then
          PointsTable.Add(pointTemp, amountTemp)
        End If
      End If
    Next
  End If
  'Use CustomerPK to find all the stored value programs he has a balance in. Put results in a hashtable
  MyCommon.QueryStr = "select SVProgramID, SUM(QtyEarned - QtyUsed) as Amount from StoredValue " & _
                      "where CustomerPK=" & CustomerPK & " and (QtyEarned - QtyUsed) > 0 " & _
                      "group by SVProgramID order by SVProgramID;"
  svdt = MyCommon.LXS_Select()
  If svdt.Rows.Count > 0 Then
    For Each svRow In svdt.Rows
      svTemp = MyCommon.Extract_Val(MyCommon.NZ(svRow.Item("SVProgramID"), 0))
      amountTemp = MyCommon.Extract_Val(MyCommon.NZ(svRow.Item("Amount"), 0))
      If Not SVTable.Contains(svTemp) Then
        SVTable.Add(svTemp, amountTemp)
      End If
    Next
  End If
  'Use CustomerPK to find all the accumulation offers with balances
  MyCommon.QueryStr = "select distinct RewardOptionID, SUM(QtyPurchased + TotalPrice) as Amount from CPE_RewardAccumulation with (NoLock) " & _
                      "where (TotalPrice>0 or QtyPurchased>0) and CustomerPK=" & CustomerPK & " and Deleted=0 " & _
                      "group by RewardOptionID order by RewardOptionID;"
  acdt = MyCommon.LXS_Select()
  If acdt.Rows.Count > 0 Then
    For Each acRow In acdt.Rows
      acTemp = MyCommon.Extract_Val(MyCommon.NZ(acRow.Item("RewardOptionID"), 0))
      amountTemp = MyCommon.Extract_Val(MyCommon.NZ(acRow.Item("Amount"), 0))
      If Not AccumTable.Contains(acTemp) Then
        AccumTable.Add(acTemp, amountTemp)
      End If
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
    Send("<link rel=""stylesheet"" href=""../css/logix-restricted.css"" type=""text/css"" media=""screen"" />")
  End If
%>
<style type="text/css">
  #functionselect
  {
    height: 166px;
  }
  * html #functionselect
  {
    height: 175px;
  }
</style>
<%
    Send_Scripts(New String() {"jquery.min.js"})
%>
<script type="text/javascript">
  function showDetail(row, btn) {
    var elemTr = document.getElementById("histdetail" + row);

    if (elemTr != null && btn != null) {
      elemTr.style.display = (btn.value == "+") ? "" : "none";
      btn.value = (btn.value == "+") ? "-" : "+";
    }
  }
</script>
<%
  Send_HeadEnd()
  
  ' Before anything else, check if we're supposed to remove someone from an offer
  If (Request.QueryString("RemoveFromOffer") <> "") Then
    ' Remove customer from a group; incoming data will look like CustomerGroupID=4&amp;OfferID=123&amp;CustomerPK=46
    
    ' Determine if customer is a household or cardholder
    If (MyCommon.Extract_Val(Request.QueryString("CustomerPK")) > 0) Then
      CustomerPK = MyCommon.Extract_Val(Request.QueryString("CustomerPK"))
      MyCommon.QueryStr = "select CustomerTypeID from Customers with (NoLock) where CustomerPK=" & CustomerPK & ";"
      rst = MyCommon.LXS_Select
      If (rst.Rows.Count > 0) Then
        CustomerTypeID = rst.Rows(0).Item("CustomerTypeID")
      End If
    End If
    
    MyCommon.QueryStr = "select ICG.CustomerGroupID, CG.EditControlTypeID, CG.RoleID from CPE_RewardOptions as RO with (NoLock) " & _
                        "inner join CPE_IncentiveCustomerGroups as ICG with (NoLock) on ICG.RewardOptionID=RO.RewardOptionID " & _
                        "inner join CustomerGroups as CG with (NoLock) on CG.CustomerGroupID=ICG.CustomerGroupID " & _
                        "where RO.IncentiveID=" & Request.QueryString("OfferID") & " and RO.Deleted=0 and ICG.Deleted=0"
    MyCommon.QueryStr += " union " & _
                         "select CG.CustomerGroupID, CG.EditControlTypeID,CG.RoleID " & _
                         "from Offers as O with (NoLock) " & _
                         "left join OfferConditions as OC with (NoLock) on OC.OfferID=O.OfferID " & _
                         "inner join CustomerGroups CG with (NoLock) on CG.CustomerGroupID=LinkID " & _
                         "where(O.Deleted = 0 And O.IsTemplate = 0 And OC.Deleted = 0 And CG.Deleted = 0 And OC.ConditionTypeID = 1 And o.OfferID = " & Request.QueryString("OfferID") & ")"
    
    rst4 = MyCommon.LRT_Select
    If rst4.Rows.Count > 0 Then
      ReDim AssocLinks(rst4.Rows.Count - 1)
      i = 0
      For Each row In rst4.Rows
        If (MyCommon.NZ(row.Item("EditControlTypeID"), 0) = 3) Then 'removal is limited per role, so loop through the user's roles and see if there's a match to the customer group's RoleID 
          For x = 0 To UserRoleIDs.Length - 1
            If UserRoleIDs(x) = MyCommon.NZ(row.Item("RoleID"), 0) Then
              RoleMatch = True
            End If
          Next
        End If
        If MyCommon.NZ(row.Item("CustomerGroupID"), 0) > 0 Then
          If (MyCommon.NZ(row.Item("EditControlTypeID"), 0) = 0 AndAlso Logix.UserRoles.RemoveCustomerFromOffers) OrElse (MyCommon.NZ(row.Item("EditControlTypeID"), 0) = 1) OrElse (MyCommon.NZ(row.Item("EditControlTypeID"), 0) = 3 AndAlso RoleMatch) Then
            MyCommon.QueryStr = "dbo.pt_GroupMembership_Delete_ByPK"
            MyCommon.Open_LXSsp()
            MyCommon.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.NVarChar, 26).Value = CustomerPK
            MyCommon.LXSsp.Parameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = MyCommon.NZ(row.Item("CustomerGroupID"), 0)
            MyCommon.LXSsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
            MyCommon.LXSsp.ExecuteNonQuery()
            MyCommon.Close_LXSsp()
            CGList &= "," & MyCommon.NZ(row.Item("CustomerGroupID"), 0)
            AssocLinks(i).LinkID = MyCommon.NZ(row.Item("CustomerGroupID"), 0)
            AssocLinks(i).LinkTypeID = 2
            i += 1
          Else
            infoMessage = Copient.PhraseLib.Detokenize("customer-offers.CustomerCouldNotBeRemoved", LanguageID, Request.QueryString("OfferID"))  'Customer could not be removed from all groups associated to offer {0}.
          End If
        End If
      Next
    End If
    
    ' Determine offers associated with the customer group to add to the history
    OffersList = ""
    If(bEnableRestrictedAccessToUEOfferBuilder) Then
            conditionalQuery=GetRestrictedAccessToUEBuilderQuery(MyCommon,Logix,"I")
     End If
            
    MyCommon.QueryStr = "select distinct Name,O.OfferID from Offers as O with (NoLock) " & _
                        "left join OfferConditions as OC with (NoLock) on OC.OfferID=O.OfferID " & _
                        "where(O.IsTemplate = 0 And OC.ConditionTypeID = 1 And OC.LinkID in (" & CGList & ") Or OC.ExcludedID in (" & CGList & "))" & _
                        " union " & _
                        "select distinct I.IncentiveName,I.IncentiveID as OfferID " & _
                        "from CPE_IncentiveCustomerGroups ICG with (NoLock) " & _
                        "inner join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID = ICG.RewardOptionID " & _
                        "inner join CPE_Incentives I with (NoLock) on I.IncentiveID = RO.IncentiveID " & _
                        "where(ICG.Deleted = 0 And RO.Deleted = 0 And i.Deleted = 0 And i.IsTemplate = 0 And ICG.CustomerGroupID in (" & CGList & ")) " 
    If(bEnableRestrictedAccessToUEOfferBuilder AndAlso Not String.IsNullOrEmpty(conditionalQuery)) Then MyCommon.QueryStr &=conditionalQuery & " "  
    MyCommon.QueryStr &=" order by OfferID ASC;"
    
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
      HistoryText = Copient.PhraseLib.Lookup("history.customer-remove-offer", LanguageID) & " #" & Request.QueryString("CustomerGroupID") & " (" & OffersList & ")"
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
      HistoryText = Copient.PhraseLib.Lookup("history.customer-remove-offer", LanguageID) & " #" & Request.QueryString("CustomerGroupID")
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
    Response.AddHeader("Location", "customer-offers.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&CardPK=" & CardPK, "") & "&Favorite=" & Request.QueryString("Favorite") & "&SortText=" & Request.QueryString("SortText") & "&SortDirection=" & Request.QueryString("SortDirection") & "&offerterms=" & Request.QueryString("offerterms"))
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
    Response.AddHeader("Location", "customer-offers.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&CardPK=" & CardPK, "") & "&Favorite=" & Request.QueryString("Favorite") & "&SortText=" & Request.QueryString("SortText") & "&SortDirection=" & Request.QueryString("SortDirection") & "&offerterms=" & Request.QueryString("offerterms"))
  End If
  
  ' special handling for customer inquiry direct link in 
  If (restrictLinks) Then
    URLOfferSum = ""
    URLcgroupedit = ""
    URLpointedit = ""
  End If
  
  ' set session to nothing just to be sure
  Session.Add("extraLink", "")
  
  If (Request.QueryString("mode") = "summary") Then
    URLtrackBack = Request.QueryString("exiturl")
    inCardNumber = Request.QueryString("cardnumber")
    extraLink = "&amp;mode=summary&amp;exiturl=" & URLtrackBack & "&amp;cardnumber=" & inCardNumber
    Session.Add("extraLink", extraLink)
  End If
  
  ' hack for popups check session for extra link
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
                          "C.TestCard, CE.Address, CE.City, CE.State, CE.Zip, CE.PhoneAsEntered, CE.email " & _
                          "from Customers C with (NoLock) left join CustomerExt CE with (NoLock) on CE.CustomerPK = C.CustomerPK " & _
                          "left join Customers C2 with (NoLock) on C2.CustomerPK = C.HHPK " & _
                          "where C.CustomerPK = " & CustomerPK
    Else
      ' IF the page was called from an outside application set ClientUserID1 to the outside passed in value
      If (inCardNumber <> "" And Request.QueryString("mode") = "summary") Then
        ClientUserID1 = MyCommon.Pad_ExtCardID(inCardNumber, Copient.commonShared.CardTypes.CUSTOMER)
        searchterms = Request.QueryString("searchterms")
                MyCommon.QueryStr = "select C.CustomerPK, C.CustomerTypeID, C.HHPK, C.FirstName, C.MiddleName, C.LastName, " & _
                                    "C.TestCard, CE.Address, CE.City, CE.State, CE.Zip, CE.PhoneAsEntered, CE.email " & _
                                    "from Customers C with (NoLock) left join CustomerExt CE with (NoLock) on CE.CustomerPK=C.CustomerPK " & _
                                    "left join Customers C2 with (NoLock) on C2.CustomerPK=C.HHPK " & _
                                    "where C.PrimaryExtID='" & MyCryptLib.SQL_StringEncrypt(ClientUserID1) & "';"
      End If
      If (Request.QueryString("searchterms") <> "" And ClientUserID1 = "") Then
        ClientUserID1 = MyCommon.Pad_ExtCardID(Request.QueryString("searchterms"), Copient.commonShared.CardTypes.CUSTOMER)
                MyCommon.QueryStr = "select C.CustomerPK, C.CustomerTypeID, C.HHPK, C.FirstName, C.MiddleName, C.LastName, " & _
                                    "C.TestCard, CE.Address, CE.City, CE.State, CE.Zip, CE.PhoneAsEntered, CE.email " & _
                                    "from Customers C with (NoLock) left join CustomerExt CE with (NoLock) on CE.CustomerPK = C.CustomerPK " & _
                                    "left join Customers C2 with (NoLock) on C2.CustomerPK=C.HHPK " & _
                                    "where C.PrimaryExtID='" &  MyCryptLib.SQL_StringEncrypt(ClientUserID1) & "' or CE.PhoneDigitsOnly = '" & MyCryptLib.SQL_StringEncrypt(MyCommon.DigitsOnly(Request.QueryString("searchterms"))) & _
                                    "' or CE.email = '" & MyCryptLib.SQL_StringEncrypt(MyCommon.Parse_Quotes(Request.QueryString("searchterms"))) & "' or C.LastName like '%" & MyCommon.Parse_Quotes(Request.QueryString("searchterms")) & "%';"
      End If
    End If
    rstResults = MyCommon.LXS_Select
    If (rstResults.Rows.Count = 1) Then
      CustomerPK = rstResults.Rows(0).Item("CustomerPK")
      HHPK = rstResults.Rows(0).Item("HHPK")
      IsHouseholdID = MyCommon.NZ(rstResults.Rows(0).Item("CustomerTypeID"), 0) = 1
      TestCard = MyCommon.NZ(rstResults.Rows(0).Item("TestCard"), False)
    Else
      infoMessage = "" & Copient.PhraseLib.Lookup("customer.notfound", LanguageID) & ""
      infoMessage = infoMessage & " <a href=""customer-offers.aspx?mode=add&amp;Search=Search" & extraLink & "&amp;searchterms=" & Request.QueryString("searchterms") & """>" & Copient.PhraseLib.Lookup("term.add", LanguageID) & "</a>"
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
      Send_Subtabs(Logix, 32, 4, LanguageID, CustomerPK, , CardPK)
    Else
      Send_Subtabs(Logix, 32, 4, LanguageID, CustomerPK)
    End If
  Else
    If CardPK > 0 Then
      Send_Subtabs(Logix, 91, 5, LanguageID, CustomerPK, extraLink, CardPK)
    Else
      Send_Subtabs(Logix, 91, 5, LanguageID, CustomerPK, extraLink)
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
var OFFER_ROWS_SHOWN = 20;

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
  } else {
    return true;
    }
}

function addOffer(custPK, cardPK, offerID) {
    document.getElementById('AddPromoLink').href = 'XMLFeeds.aspx?AddOffer=' + offerID + '&CustPK=' + custPK + '&CardPK=' + cardPK + '&AdminUserID=<%Sendb(AdminUserID)%>&height=300&width=300&LanguageID=<%Sendb(LanguageID)%>';
    
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

  //if (lastStartPos + OFFER_ROWS_SHOWN <= recCt) {
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

function URLReplace(OldString) {
  var NewString
  
  NewString = OldString.replace('%', '[%25]');
  NewString = NewString.replace('+', '%2B');
  NewString = NewString.replace('/', '%2F');
  NewString = NewString.replace('?', '%3F');
  //NewString = NewString.replace('#', '%23');
  //NewString = NewString.replace('&', '%26');
  NewString = NewString.replace('_', '%5F');
  //alert(OldString + " - " +  NewString);
  
  return NewString;
}

function searchOffers() {
    var elemOfferTerms = document.getElementById("offerterms");
    var offerTerms = '';
  if (elemOfferTerms != null) {
    offerTerms = elemOfferTerms.value;
  }
    <%
        Dim strTerms = Request.QueryString("searchterms")
        If (strTerms <> "") Then
            strTerms = strTerms.Replace("'", "\'")
            strTerms = strTerms.Replace("""", "\""")
        End If
     %>

     offerTerms = encodeURIComponent(offerTerms);

    var qryStr = 'customer-offers.aspx?search=Search&searchterms=<%Sendb(strTerms)%>&CustomerPK=<%Sendb(CustomerPK)%>&CustPK=<%Sendb(CustomerPK)%><%Sendb(IIf(CardPK > 0, "&CardPK=" & CardPK, ""))%>&Favorite=0&offerSearch=Search&offerterms=' + offerTerms + '#h01';
    document.location = qryStr;
}

function submitOfferSearch(e) {
    var key = e.which ? e.which : e.keyCode;
    
    if (key == 13) {
        if (e && e.preventDefault) {
            e.preventDefault(); // DOM style`
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
  
  newURL = currentURL + "&amp;Favorite=" + option;
  
  frm.action = newURL;
  frm.submit();
}

//-->
</script>
<script type="text/javascript" src="../javascript/jquery.min.js"></script>
<script type="text/javascript" src="../javascript/thickbox.js"></script>
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
          Sendb(Copient.PhraseLib.Lookup("term.customercard", LanguageID) & " #" & ExtCardID)
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
            Sendb(": " & MyCommon.TruncateString(FullName, 16))
      End If
      If (restrictLinks AndAlso URLtrackBack <> "") Then
        Send(" - <a href=""" & URLtrackBack & """>" & Copient.PhraseLib.Lookup("customer-inquiry.return", LanguageID) & "</a>")
      End If
    %>
  </h1>
  <div id="controls" <% Sendb(IIf(restrictLinks AndAlso URLtrackBack <> "", " style=""width:230px;""", "")) %>>
    <%
      If (CustomerPK > 0 And Logix.UserRoles.ViewCustomerNotes) Then
        Send_CustomerNotes(CustomerPK, CardPK)
      End If
      If (CustomerPK > 0) Then
        Send_AddOffer(CustomerPK, CardPK)
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
      
    If (infoMessage <> "") Then
      Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div><br />")
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
    <%
      MyCommon.QueryStr = "select RewardID from StoredFranking as SF with (NoLock) where CustomerPK=" & CustomerPK & " and Status in (0,1);"
      rst2 = MyCommon.LXS_Select
      If rst2.Rows.Count > 0 Then
        Send("<div style=""text-align:right;margin-bottom:4px;"">")
        Send("  <a href=""javascript:openPopup('customer-storedfranking.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&CardPK=" & CardPK, "") & "');"">" & Copient.PhraseLib.Lookup("customer-offers.CustomerHasStoredFranking", LanguageID) & "</a>")
        Send("</div>")
      End If
    %>
    <% If (Logix.UserRoles.ViewCustomerOffers AndAlso CustomerPK > 0) Then%>
    <a name="h01"></a>
    <%
      Send("<div id=""listbar"">")
      Send("</div>")
    %>
    <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID)) %>">
      <thead>
        <tr>
          <th align="left" class="th-button" scope="col" style="text-align: center;">
            <a href="customer-offers.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;offerterms=<%Sendb(Request.QueryString("offerterms"))%>&amp;CustPK=<%Sendb(CustomerPK)%><% Sendb(IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "")) %>&amp;search=Search&amp;SortText=Priority&amp;SortDirection=<% Sendb(SortDirection & extralink) %>&amp;Favorite=<% Sendb(Request.QueryString("Favorite")) %>">
              <% Sendb(Copient.PhraseLib.Lookup("term.fav", LanguageID))%>
            </a>
            <%
              If SortText = "Priority" Then
                If SortDirection = "ASC" Then
                  Sendb("<span class=""sortarrow"">&#9660;</span>")
                Else
                  Sendb("<span class=""sortarrow"">&#9650;</span>")
                End If
              End If
            %>
          </th>
          <th align="left" class="th-button" scope="col" style="text-align: center;">
            <% Sendb(Left(Copient.PhraseLib.Lookup("term.remove", LanguageID), 3))%>
          </th>
          <th align="center" class="th-button" scope="col" style="text-align: center;">
            <% Sendb(Copient.PhraseLib.Lookup("term.view", LanguageID))%>
          </th>
          <th align="center" class="th-adjustment" scope="col" style="text-align: center;"
            colspan="3">
            <% Sendb(Copient.PhraseLib.Lookup("term.adjustments", LanguageID))%>
          </th>
          <th align="left" class="th-bigid" scope="col">
            <a href="customer-offers.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;offerterms=<%Sendb(Request.QueryString("offerterms"))%>&amp;CustPK=<%Sendb(CustomerPK)%><% Sendb(IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "")) %>&amp;search=Search&amp;SortText=OfferID&amp;SortDirection=<% Sendb(SortDirection & extralink) %>&amp;Favorite=<% Sendb(Request.QueryString("Favorite")) %>">
              <% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>
            </a>
            <%
              If SortText = "OfferID" Then
                If SortDirection = "ASC" Then
                  Sendb("<span class=""sortarrow"">&#9660;</span>")
                Else
                  Sendb("<span class=""sortarrow"">&#9650;</span>")
                End If
              End If
            %>
          </th>
          <th align="left" class="th-name" scope="col">
            <a href="customer-offers.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;offerterms=<%Sendb(Request.QueryString("offerterms"))%>&amp;CustPK=<%Sendb(CustomerPK)%><% Sendb(IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "")) %>&amp;search=Search&amp;SortText=Name&amp;SortDirection=<% Sendb(SortDirection & extralink) %>&amp;Favorite=<% Sendb(Request.QueryString("Favorite")) %>">
              <% Sendb(Copient.PhraseLib.Lookup("term.offer", LanguageID))%>
            </a>
            <%
              If SortText = "Name" Then
                If SortDirection = "ASC" Then
                  Sendb("<span class=""sortarrow"">&#9660;</span>")
                Else
                  Sendb("<span class=""sortarrow"">&#9650;</span>")
                End If
              End If
            %>
          </th>
          <th align="center" class="th-status" scope="col">
            <% Sendb(Copient.PhraseLib.Lookup("term.status", LanguageID))%>
            <%
              If SortText = "StatusFlag" Then
                If SortDirection = "ASC" Then
                  Sendb("<span class=""sortarrow"">&#9660;</span>")
                Else
                  Sendb("<span class=""sortarrow"">&#9650;</span>")
                End If
              End If
            %>
          </th>
          <th align="center" class="th-group" scope="col">
            <a href="customer-offers.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;offerterms=<%Sendb(Request.QueryString("offerterms"))%>&amp;CustPK=<%Sendb(CustomerPK)%><% Sendb(IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "")) %>&amp;search=Search&amp;SortText=GroupName&amp;SortDirection=<% Sendb(SortDirection & extralink) %>&amp;Favorite=<% Sendb(Request.QueryString("Favorite")) %>">
              <% Sendb(Copient.PhraseLib.Lookup("term.customergroup", LanguageID))%>
            </a>
            <%
              If SortText = "GroupName" Then
                If SortDirection = "ASC" Then
                  Sendb("<span class=""sortarrow"">&#9660;</span>")
                Else
                  Sendb("<span class=""sortarrow"">&#9650;</span>")
                End If
              End If
            %>
          </th>
        </tr>
      </thead>
      <tbody>
        <%
          offerCt = 0
          InstalledEngines = MyCommon.GetInstalledEngines
          If MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM) And InstalledEngines.Length = 1 Then
            ' CM engine - make this offer list match that returned via GetAccount in CmConnector
            If MyCommon.Fetch_CM_SystemOption(24) Then
              ' auto household customer groups (implicit householding groups))
              MyCommon.QueryStr = "dbo.pa_LogixServ_FetchCustGroups_MemberOrHousehold"
              MyCommon.Open_LXSsp()
              MyCommon.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
              If IsHouseholdID Then
                MyCommon.LXSsp.Parameters.Add("@HHPK", SqlDbType.BigInt).Value = CustomerPK
              Else
                MyCommon.LXSsp.Parameters.Add("@HHPK", SqlDbType.BigInt).Value = HHPK
              End If
            Else
              ' household id must be a member of a group (explicit householding groups)
              MyCommon.QueryStr = "dbo.pa_LogixServ_FetchCustGroups"
              MyCommon.Open_LXSsp()
              MyCommon.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
              MyCommon.LXSsp.Parameters.Add("@HHPK", SqlDbType.BigInt).Value = HHPK
            End If
            rst = MyCommon.LXSsp_select
            MyCommon.Close_LXSsp()
          Else
            MyCommon.QueryStr = "select CustomerGroupID from GroupMembership with (NoLock) where CustomerPK=" & CustomerPK & " and Deleted=0"
            rst = MyCommon.LXS_Select()
          End If
            
          If (Targeted) Then
            CgXml = "<customergroups>"
          Else
            CgXml = "<customergroups><id>1</id><id>2</id>"
          End If
          rowCount = rst.Rows.Count
          If rowCount > 0 Then
            For Each row In rst.Rows
              CgXml &= "<id>" & MyCommon.NZ(row.Item("CustomerGroupID"), "") & "</id>"
            Next
          End If
          CgXml &= "</customergroups>"
            
          MyCommon.QueryStr = "dbo.pa_CustomerOffersCurrent"
          MyCommon.Open_LRTsp()
          MyCommon.LRTsp.Parameters.Add("@cgXml", SqlDbType.Xml).Value = CgXml
          MyCommon.LRTsp.Parameters.Add("@IsEmployee", SqlDbType.Bit).Value = Employee
          If (Request.QueryString("offerterms") <> "") Then
            OfferTerms = Request.QueryString("offerterms")
            OfferTerms = OfferTerms.Replace("%", "[%]")
            OfferTerms = OfferTerms.Replace("_", "[_]")
            MyCommon.LRTsp.Parameters.Add("@Filter", SqlDbType.NVarChar, 50).Value = OfferTerms 'Request.QueryString("offerterms")
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
            
           'Removing transformed offers
          If (bEnableRestrictedAccessToUEOfferBuilder) Then
               RemoveTransformedOffers(dtAssigned,MyCommon)
          End If
            
          'If showing credits only, get the list of customer offers that
          'have balances, then drop non-matches from the main list.
          If (ShowCreditOnly) Then
            BalanceOffers = GetBalanceOffers(CustomerPK)
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
            
          'Check for card type conditions among the offers; if found, compare with the customer's cards.
          'Offers targeted to a card type not held by the customer are dropped from the main list.
          UnavailableCardTypeOffers = GetUnavailableCardTypeOffers(CustomerPK)
          i = 0
          While i < dtAssigned.Rows.Count
            If UnavailableCardTypeOffers.ContainsKey(MyCommon.NZ(dtAssigned.Rows(i).Item("OfferID"), 0).ToString) Then
              dtAssigned.Rows.RemoveAt(i)
              dtAssigned.AcceptChanges()
            Else
              i += 1
            End If
          End While
            
          'Remove the offers that are set to restricted redemptions if the customer record is also set to have restricted redemptions
          dtAssigned = RemoveRestrictedRedemptionOffers(MyCommon, CustomerPK, dtAssigned)
            
          'Sort the Assigned offers
          sortedRows = dtAssigned.Select("", SortText & " " & SortDirection)
          StatusTable = LoadOfferStatuses(sortedRows, MyCommon, Logix)
          i = 0
          For i = 0 To sortedRows.Length - 1
            row2 = sortedRows(i)
              
            If PrevOfferID <> MyCommon.NZ(row2.Item("OfferID"), 0) Then
              OfferStatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_UNKNOWN
              OfferStatus = StatusTable.Item(MyCommon.NZ(row2.Item("OfferID"), "0").ToString)
              If (OfferStatus IsNot Nothing) Then
                OfferStatusCode = OfferStatus
              End If
              If (MyCommon.NZ(row2.Item("EngineTypeID"), 0) = 2 AndAlso MyCommon.NZ(row2.Item("EngineSubTypeID"), 0) = 2) Then
                OfferIsAirMiles = True
              End If
              If (TestCard = False AndAlso (IIf(ExcludeExpired, OfferStatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_ACTIVE, OfferStatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_EXPIRED OrElse OfferStatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_ACTIVE)) OrElse (TestCard = True AndAlso OfferStatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_TESTING AndAlso MyCommon.Fetch_SystemOption(88) = "1")) Then
                offerCt += 1
                System.Math.DivRem(offerCt, 2, r)
                Send("<tr id=""trOffer" & offerCt & """" & IIf(r = 0, "", " class=""shaded""") & " style=""display:none;"">")
                  
                ' Favorite column
                Send("  <td align=""center"">")
                If (MyCommon.NZ(row2.Item("AdminUserID"), 0)) = AdminUserID Then
                  If Logix.UserRoles.FavoriteOffersForSelf Then
                    Send("    <a href=""customer-offers.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "&amp;DeleteFavorite=" & MyCommon.NZ(row2.Item("OfferID"), -1) & "&amp;Favorite=" & Favorite & "&amp;SortText=" & SortText & "&amp;SortDirection=" & IIf(SortDirection = "ASC", "DESC", "ASC") & "&amp;offerterms=" & Request.QueryString("offerterms") & """><img src=""../images/star-on.png"" alt=""" & Copient.PhraseLib.Lookup("customer-inquiry.unfavorite", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("customer-inquiry.unfavorite", LanguageID) & """ /></a>")
                  Else
                    Send("    <img src=""../images/star-on.png"" alt=""" & Copient.PhraseLib.Lookup("term.favorite", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.favorite", LanguageID) & """ /></a>")
                  End If
                Else
                  If Logix.UserRoles.FavoriteOffersForSelf Then
                    Send("    <a href=""customer-offers.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "&amp;AddFavorite=" & MyCommon.NZ(row2.Item("OfferID"), -1) & "&amp;Favorite=" & Favorite & "&amp;SortText=" & SortText & "&amp;SortDirection=" & IIf(SortDirection = "ASC", "DESC", "ASC") & "&amp;offerterms=" & Request.QueryString("offerterms") & """><img src=""../images/star-off.png"" alt=""" & Copient.PhraseLib.Lookup("customer-inquiry.favorite", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("customer-inquiry.favorite", LanguageID) & """ /></a>")
                  Else
                    Send("    <img src=""../images/star-off.png"" alt=""" & Copient.PhraseLib.Lookup("customer-inquiry.notafavorite", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("customer-inquiry.notafavorite", LanguageID) & """ /></a>")
                  End If
                End If
                Send("  </td>")
                  
                ' Remove column
                If (MyCommon.NZ(row2.Item("CustomerGroupID"), -1) <= 2) Then
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
                    Sendb("    <a href=""XMLFeeds.aspx?OtherOfferscheck=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "&amp;OfferID=" & MyCommon.NZ(row2.Item("OfferID"), "") & "&amp;CustomerGroupID=" & MyCommon.NZ(row2.Item("CustomerGroupID"), -1) & "&amp;AdminUserID=" & AdminUserID & "&amp;height=300&amp;width=300"" title=""" & Copient.PhraseLib.Lookup("term.alert", LanguageID) & """ class=""thickbox"">")
                    Send("<input type=""button"" class=""ex"" id=""ex" & MyCommon.NZ(row2.Item("OfferID"), -1) & """ name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.remove", LanguageID) & """ " & " value=""X"" /></a>")
                  End If
                  Send("  </td>")
                End If
                  
                ' View column
                Send("  <td align=""center"">")
                If CardPK > 0 Then
                  If Logix.UserRoles.ViewRedemptionHistory Then
                    Send("    <a href=""XMLFeeds.aspx?OfferRedemptions=1&amp;CustomerPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "&amp;OfferID=" & MyCommon.NZ(row2.Item("OfferID"), "") & "&amp;Lang=" & LanguageID & "&amp;height=400&amp;width=600"" title=""" & Copient.PhraseLib.Lookup("term.redemptions", LanguageID) & """ class=""thickbox"">")
                    Send("      <input type=""button"" class=""view"" id=""view" & MyCommon.NZ(row2.Item("OfferID"), "") & """ name=""view" & MyCommon.NZ(row2.Item("OfferID"), "") & """ title=""" & Copient.PhraseLib.Lookup("term.redemptions", LanguageID) & """" & IIf(Logix.UserRoles.ViewRedemptionHistory AndAlso CardPK > 0, "", " disabled=""disabled""") & " value=""..."" />")
                    Send("    </a>")
                  Else
                    Send("      <input type=""button"" class=""view"" id=""view" & MyCommon.NZ(row2.Item("OfferID"), "") & """ name=""view" & MyCommon.NZ(row2.Item("OfferID"), "") & """ title=""" & Copient.PhraseLib.Lookup("term.redemptions", LanguageID) & """" & " disabled=""disabled" & " value=""..."" />")
                  End If
                Else
                  Send("    <input type=""button"" class=""view"" id=""view" & MyCommon.NZ(row2.Item("OfferID"), "") & """ name=""view" & MyCommon.NZ(row2.Item("OfferID"), "") & """ title=""" & Copient.PhraseLib.Lookup("term.redemptions", LanguageID) & """" & IIf(Logix.UserRoles.ViewRedemptionHistory AndAlso CardPK > 0, "", " disabled=""disabled""") & " value=""..."" />")
                End If
                Send("  </td>")
                  
                ' Adjust points
                Send("  <td style=""width:30px;"">")
                MyCommon.QueryStr = "dbo.pa_CustomerOfferHasPointsProgram"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.Int).Value = MyCommon.NZ(row2.Item("OfferID"), -1)
                MyCommon.LRTsp.Parameters.Add("@HasPointsProgram", SqlDbType.Bit).Direction = ParameterDirection.Output
                pointdt = MyCommon.LRTsp_select()
                IsPtsOffer = MyCommon.LRTsp.Parameters("@HasPointsProgram").Value
                MyCommon.Close_LRTsp()
                'Find if the customer has points if there is a points program
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
                  DisabledPtsAdj = ""
                  If OfferIsAirMiles Then
                    'US AirMiles offer points access needed
                    If (Logix.UserRoles.AccessAirmilesPointsBalances = False) Then DisabledPtsAdj = "disabled=""disabled"""
                  Else
                    'General offer points access needed
                    If (Logix.UserRoles.AccessPointsBalances = False) Then DisabledPtsAdj = "disabled=""disabled"""
                  End If
                  Sendb("  <input type=""button"" class=""adjust"" id=""ptsAdj" & MyCommon.NZ(row2.Item("OfferID"), "") & """ name=""ptsAdj"" value=""P""" & DisabledPtsAdj & IIf(HasPoints OrElse DisabledPtsAdj <> "", " style=""""", " style=""background-color:#CCCCCC;font-style:italic;""") & " title=""" & Copient.PhraseLib.Lookup("term.adjust", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.points", LanguageID), VbStrConv.Lowercase) & """ ")
                  Send("onClick=""javascript:openPopup('point-adjust.aspx?OfferID=" & MyCommon.NZ(row2.Item("OfferID"), "") & "&amp;CustomerPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "');"" />")
                End If
                HasPoints = False
                Send("  </td>")
                  
                ' Adjust stored value
                Send("  <td style=""width:30px;"">")
                MyCommon.QueryStr = "dbo.pa_CustomerOfferHasSVProgram"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.Int).Value = MyCommon.NZ(row2.Item("OfferID"), -1)
                MyCommon.LRTsp.Parameters.Add("@HasSVProgram", SqlDbType.Bit).Direction = ParameterDirection.Output
                svdt = MyCommon.LRTsp_select()
                IsSVOffer = MyCommon.LRTsp.Parameters("@HasSVProgram").Value
                MyCommon.Close_LRTsp()
                'Find if the customer has a stored value balance
                If svdt.Rows.Count > 0 Then
                  If svdt.Rows.Count = 1 Then
                    If MyCommon.NZ(svdt.Rows(0).Item("SVProgramID"), 0) > 0 Then
                      SVProgram = MyCommon.Extract_Val(MyCommon.NZ(svdt.Rows(0).Item("SVProgramID"), 0))
                    End If
                    If SVProgram > 0 Then
                      HasSV = SVTable.ContainsKey(SVProgram)
                    End If
                  Else
                    For Each svRow In svdt.Rows
                      If MyCommon.NZ(svRow.Item("SVProgramID"), 0) > 0 Then
                        SVProgram = MyCommon.Extract_Val(MyCommon.NZ(svRow.Item("SVProgramID"), 0))
                      End If
                      If SVProgram > 0 Then
                        HasSV = SVTable.ContainsKey(SVProgram)
                      End If
                      If HasSV Then Exit For
                    Next
                  End If
                End If
                If (IsSVOffer) Then
                  If (Logix.UserRoles.AccessStoredValue = False) Then
                    DisabledSVAdj = " disabled=""disabled"""
                  Else
                    DisabledSVAdj = ""
                  End If
                  Sendb("  <input type=""button"" class=""adjust"" id=""svAdj" & MyCommon.NZ(row2.Item("OfferID"), "") & """ name=""svAdj"" value=""S""" & DisabledSVAdj & IIf(HasSV, " style=""""", " style=""background-color:#CCCCCC;font-style:italic;""") & " title=""" & Copient.PhraseLib.Lookup("term.adjust", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.storedvalue", LanguageID), VbStrConv.Lowercase) & """ ")
                  Send("onClick=""javascript:openPopup('sv-adjust-redirect.aspx?OfferID=" & MyCommon.NZ(row2.Item("OfferID"), "") & "&amp;CustomerPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "&amp;Opener=" & CopientFileName & "');"" />")
                End If
                Send("  </td>")
                  
                'Adjust accumulation
                Send("  <td style=""width:30px;"">")
                MyCommon.QueryStr = "dbo.pa_CustomerOfferHasAccumulation"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.Int).Value = MyCommon.NZ(row2.Item("OfferID"), -1)
                MyCommon.LRTsp.Parameters.Add("@HasAccum", SqlDbType.Bit).Direction = ParameterDirection.Output
                acdt = MyCommon.LRTsp_select()
                IsAccumOffer = MyCommon.LRTsp.Parameters("@HasAccum").Value
                'Find if the customer has accumulation in the offer
                If acdt.Rows.Count > 0 Then
                  If acdt.Rows.Count = 1 Then
                    If MyCommon.NZ(acdt.Rows(0).Item("RewardOptionID"), 0) > 0 Then
                      RewardOptionID = MyCommon.Extract_Val(MyCommon.NZ(acdt.Rows(0).Item("RewardOptionID"), 0))
                    End If
                    If RewardOptionID > 0 Then
                      HasAccum = AccumTable.ContainsKey(RewardOptionID)
                    End If
                  Else
                    For Each acRow In acdt.Rows
                      If MyCommon.NZ(acRow.Item("RewardOptionID"), 0) > 0 Then
                        RewardOptionID = MyCommon.Extract_Val(MyCommon.NZ(acRow.Item("RewardOptionID"), 0))
                      End If
                      If RewardOptionID > 0 Then
                        HasAccum = AccumTable.ContainsKey(RewardOptionID)
                      End If
                      If HasAccum Then Exit For
                    Next
                  End If
                End If
                'Find if the offer is an Airmiles offer and use the correct permissions
                DisabledAccumAdj = ""
                If OfferIsAirMiles Then
                  'US AirMiles offer accumulation access needed
                  If (Logix.UserRoles.AccessAirmilesAccumBalances = False) Then DisabledAccumAdj = "disabled=""disabled"""
                Else
                  'General offer accumulation access needed
                  If (Logix.UserRoles.AccessAccumBalances = False) Then DisabledAccumAdj = "disabled=""disabled"""
                End If
                If (MyCommon.LRTsp.Parameters("@HasAccum").Value) Then
                  Sendb("  <input type=""button"" class=""adjust"" id=""accumAdj" & MyCommon.NZ(row2.Item("OfferID"), "") & """ name=""accumAdj"" value=""A""" & DisabledAccumAdj & IIf(HasAccum OrElse DisabledAccumAdj <> "", " style=""""", " style=""background-color:#CCCCCC;font-style:italic;""") & " title=""" & Copient.PhraseLib.Lookup("term.adjust", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.accumulation", LanguageID), VbStrConv.Lowercase) & """ ")
                  Sendb("onClick=""javascript:openPopup('CPEaccum-adjust.aspx?OfferID=" & MyCommon.NZ(row2.Item("OfferID"), "") & "&amp;CustomerPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "');"" />")
                End If
                MyCommon.Close_LRTsp()
                Send("  </td>")
                  
                ' Offer ID
                Send("  <td>" & MyCommon.NZ(row2.Item("OfferID"), -1) & "</td>")
                  
                ' Offer name and link
                If (Not restrictLinks) Then
                              Dim Name As String=""
                                 If (MyCommon.Fetch_UE_SystemOption(168) = "1" AndAlso MyCommon.NZ(row2.Item("BuyerID"), "") <> "") Then
                                    Name = "Buyer " + row2.Item("BuyerID").ToString() + " - " + MyCommon.SplitNonSpacedString(row2.Item("Name"), 25).ToString()
                                Else
                                    Name = MyCommon.NZ(MyCommon.SplitNonSpacedString(row2.Item("Name"), 25), Copient.PhraseLib.Lookup("term.unknown", LanguageID))
                                End If
                  If (MyCommon.NZ(row2.Item("EngineTypeID"), 0) = 0) Then 'CM offers
                    Send("  <td><a href=""" & URLOfferSum & "?OfferID=" & MyCommon.NZ(row2.Item("OfferID"), -1) & """ title=""" & MyCommon.NZ(row2.Item("Description"), "").ToString().Replace("""", "'") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row2.Item("Name"), ""), 25) & "</a></td>")
                  ElseIf (MyCommon.NZ(row2.Item("EngineTypeID"), 0) = 9) Then 'UE offers
                    Send("  <td><a href=""" & URLUEOfferSum & "?OfferID=" & MyCommon.NZ(row2.Item("OfferID"), -1) & """ title=""" & MyCommon.NZ(row2.Item("Description"), "").ToString().Replace("""", "'") & """>" & Name & "</a></td>")
                  ElseIf (MyCommon.NZ(row2.Item("EngineTypeID"), 0) = 2) Then 'CPE offers
                    Send("  <td><a href=""" & URLCPEOfferSum & "?OfferID=" & MyCommon.NZ(row2.Item("OfferID"), -1) & """ title=""" & MyCommon.NZ(row2.Item("Description"), "").ToString().Replace("""", "'") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row2.Item("Name"), ""), 25) & "</a></td>")
                         ElseIf (MyCommon.NZ(row2.Item("EngineTypeID"), 0) = 3) Then 'Web Site offers
                           Send("  <td><a href=""" & URLWEBOfferSum & "?OfferID=" & MyCommon.NZ(row2.Item("OfferID"), -1) & """ title=""" & MyCommon.NZ(row2.Item("Description"), "").ToString().Replace("""", "'") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row2.Item("Name"), ""), 25) & "</a></td>")
                  End If
                Else
                  If (MyCommon.NZ(row2.Item("EngineTypeID"), 0) = 0) Then  'CM offers
                    Send("<td><span title=""" & MyCommon.NZ(row2.Item("Description"), "") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row2.Item("Name"), ""), 25) & "</span></td>")
                  ElseIf (MyCommon.NZ(row2.Item("EngineTypeID"), 0) = 9) Then 'UE offers
                    Send("  <td><a href=""javascript:openPopup('" & URLUEOfferSum & "?OfferID=" & MyCommon.NZ(row2.Item("OfferID"), -1) & "&Popup=1')"" title=""" & MyCommon.NZ(row2.Item("Description"), "").ToString().Replace("""", "'") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row2.Item("Name"), ""), 25) & "</a></td>")
                   ElseIf  (MyCommon.NZ(row2.Item("EngineTypeID"), 0) = 2) Then 'CPE offers
                                Send("  <td><a href=""javascript:openPopup('" & URLCPEOfferSum & "?OfferID=" & MyCommon.NZ(row2.Item("OfferID"), -1) & "&Popup=1')"" title=""" & MyCommon.NZ(row2.Item("Description"), "").ToString().Replace("""", "'") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row2.Item("Name"), ""), 25) & "</a></td>")
                            ElseIf (MyCommon.NZ(row2.Item("EngineTypeID"), 0) = 3) Then 'Web offers
                                Send("  <td><a href=""javascript:openPopup('" & URLWEBOfferSum & "?OfferID=" & MyCommon.NZ(row2.Item("OfferID"), -1) & "&Popup=1')"" title=""" & MyCommon.NZ(row2.Item("Description"), "").ToString().Replace("""", "'") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row2.Item("Name"), ""), 25) & "</a></td>")
                  End If
                End If
                  
                ' Offer status
                Send("  <td>")
                Send("    " & Logix.GetOfferStatusHtml(Integer.Parse(OfferStatus), LanguageID))
                Send("  </td>")
                  
                ' Customer group
                If (MyCommon.NZ(row2.Item("CustomerGroupID"), -1) <> 1 And MyCommon.NZ(row2.Item("CustomerGroupID"), -1) <> 2) And (MyCommon.NZ(row2.Item("NewCardholders"), False) = False) Then
                  Send("  <td>")
                  ' get all the customer groups assigned to this offer.  Peek ahead after getting the current row to get any additional customer groups assigned to this offer
                  For j = i To sortedRows.Length - 1
                    If MyCommon.NZ(sortedRows(j).Item("OfferID"), 0) = MyCommon.NZ(row2.Item("OfferID"), 0) Then
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
              End If
            End If
            PrevOfferID = MyCommon.NZ(row2.Item("OfferID"), 0)
          Next
            
          If offerCt = 0 AndAlso Favorite = 1 Then
            Send("<tr>")
            Send("  <td colspan=""10"" align=""center"" style=""padding-top:5px;"">" & Copient.PhraseLib.Lookup("customer-inquiry.nofavorites", LanguageID) & "</td>")
            Send("</tr>")
          End If
        %>
      </tbody>
    </table>
    <%
      Send(" <div id=""pageIter"">")
      Send("  <div id=""searcher"" title=""Search terms"">")
      Send("   <input type=""text"" id=""offerterms"" name=""offerterms"" maxlength=""255"" style=""font-family:arial;font-size:12px;width:45%;"" value=""" & Request.QueryString("offerterms") & """ onkeydown=""submitOfferSearch(event);"" />")
      Send("   <input type=""button"" style=""font-family:arial;font-size:12px;width:45%;"" id=""btnOffer"" name=""btnOffer"" onclick=""searchOffers();"" value=""" & Copient.PhraseLib.Lookup("term.search", LanguageID) & """ />")
      Send("  </div>")
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
      Send("    <option value=""0""" & IIf(Favorite = 0 AndAlso Not ExcludeExpired, "", " selected=""selected""") & ">" & Copient.PhraseLib.Lookup("customer-inquiry.showall", LanguageID) & "</option>")
      Send("    <option value=""1""" & IIf(Favorite = 1, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("customer-inquiry.showfavorites", LanguageID) & "</option>")
      Send("    <option value=""2""" & IIf(Favorite = 0 AndAlso ExcludeExpired, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showall-noexpired", LanguageID) & "</option>")
      Send("    <option value=""5""" & IIf(Favorite = 0 AndAlso ShowCreditOnly, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showonlycredit", LanguageID) & "</option>")
      Send("    <option value=""6""" & IIf(Favorite = 0 AndAlso ExcludeExpired AndAlso Targeted, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showactivetargeted", LanguageID) & "</option>")
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
<%
  Send("<script type=""text/javascript"">")
  If (Request.QueryString("refresh") = "") Then
    Send(" if (document.mainform.searchterms != null) document.mainform.searchterms.focus();")
  End If
  Send(" if (document.getElementById(""listbar"") != null && document.getElementById(""pageIter"") !=null ) { ")
  Send("      document.getElementById(""listbar"").innerHTML = document.getElementById(""pageIter"").innerHTML;")
  Send("      document.getElementById(""pageIter"").innerHTML = """ & """")
  Send(" }")
  Send("  showOfferPage(1);")
  Send("</script>")
%>
<script type="text/javascript">
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
    'Stored value programs
    MyCommon.QueryStr = "select distinct SVProgramID from StoredValue with (NoLock) " & _
                        "where (QtyEarned - QtyUsed)>0 and CustomerPK=" & CustomerPK & " order by SVProgramID;"
    dt = MyCommon.LXS_Select
    If dt.Rows.Count > 0 Then
      For Each row In dt.Rows
        BalanceSVPrograms &= "," & MyCommon.NZ(row.Item("SVProgramID"), 0)
      Next
    End If
    'Accumulation ROIDs
    MyCommon.QueryStr = "select distinct RewardOptionID from CPE_RewardAccumulation with (NoLock) " & _
                        "where (TotalPrice>0 or QtyPurchased>0) and CustomerPK=" & CustomerPK & " and Deleted=0 order by RewardOptionID;"
    dt = MyCommon.LXS_Select
    If dt.Rows.Count > 0 Then
      For Each row In dt.Rows
        BalanceAccumRoids &= "," & MyCommon.NZ(row.Item("RewardOptionID"), 0)
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
                        "UNION " & _
                        " select DISTINCT O.IncentiveID as OfferID from CPE_Incentives as O " & _
                        " INNER JOIN CPE_RewardOptions as RO on RO.IncentiveID=O.IncentiveID " & _
                        " INNER JOIN CPE_DeliverableStoredValue as DSV on DSV.RewardOptionID=RO.RewardOptionID " & _
                        " where DSV.SVProgramID in (" & BalanceSVPrograms & ") " & _
                        " and O.Deleted=0 and RO.Deleted=0 and DSV.Deleted=0 " & _
                        "   union " & _
                        " select DISTINCT O.IncentiveID as OfferID from CPE_Incentives as O " & _
                        " INNER JOIN CPE_RewardOptions as RO on RO.IncentiveID=O.IncentiveID " & _
                        " INNER JOIN CPE_IncentiveStoredValuePrograms as ISV on ISV.RewardOptionID=RO.RewardOptionID " & _
                        " where ISV.SVProgramID in (" & BalanceSVPrograms & ") " & _
                        " and O.Deleted=0 and RO.Deleted=0 and ISV.Deleted=0 " & _
                        "   union " & _
                        " select DISTINCT O.OfferID from Offers as O " & _
                        " INNER JOIN OfferConditions as C on C.OfferID=O.OfferID " & _
                        " where C.ConditionTypeID=6 and C.LinkID in (" & BalanceSVPrograms & ") " & _
                        " and O.Deleted=0 and C.Deleted=0 " & _
                        "   union " & _
                        " select DISTINCT O.OfferID from Offers as O " & _
                        " INNER JOIN OfferRewards as R on R.OfferID=O.OfferID " & _
                        " where R.RewardTypeID=10 and R.LinkID in (" & BalanceSVPrograms & ") " & _
                        " and O.Deleted=0 and R.Deleted=0 " & _
                        "UNION " & _
                        " select IncentiveID as OfferID from CPE_RewardOptions as RO with (NoLock) " & _
                        " where RewardOptionID in (" & BalanceAccumRoids & ") " & _
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
  
  Function GetUnavailableCardTypeOffers(ByVal CustomerPK As Long) As Hashtable
    Dim MyCommon As New Copient.CommonInc
    Dim dtCardTypeIDs As DataTable
    Dim dtOfferIDs As DataTable
    Dim row As DataRow
    Dim CustomerCardTypeIDs As String = ""
    Dim CardTypeOffers As New Hashtable
    
    MyCommon.Open_LogixRT()
    MyCommon.Open_LogixXS()
    
    'Get a list of the CardTypeIDs held by the customer
    MyCommon.QueryStr = "select CardTypeID from CardIDs with (NoLock) where CustomerPK=" & CustomerPK & " order by CardTypeID;"
    dtCardTypeIDs = MyCommon.LXS_Select
    If dtCardTypeIDs.Rows.Count > 0 Then
      For Each row In dtCardTypeIDs.Rows
        If CustomerCardTypeIDs <> "" Then
          CustomerCardTypeIDs &= ","
        End If
        CustomerCardTypeIDs &= row.Item("CardTypeID")
      Next
    End If
    
    'Get a list of OfferIDs for offers that have card type conditions targeted to cards NOT held by the customer
    MyCommon.QueryStr = "select distinct RO.IncentiveID as OfferID from CPE_ST_IncentiveCardTypes as ICT with (NoLock) " & _
                        "left join CPE_RewardOptions as RO with (NoLock) on RO.RewardOptionID=ICT.RewardOptionID " & _
                        "where RO.RewardOptionID not in (" & _
                        "  select distinct RewardOptionID from CPE_ST_IncentiveCardTypes with (NoLock) " & _
                        "  where CardTypeID in (" & IIf(CustomerCardTypeIDs <> "", CustomerCardTypeIDs, "-1") & ") " & _
                        ");"
    dtOfferIDs = MyCommon.LRT_Select
    If dtOfferIDs.Rows.Count > 0 Then
      For Each row In dtOfferIDs.Rows
        CardTypeOffers.Add(MyCommon.NZ(row.Item("OfferID"), 0).ToString, MyCommon.NZ(row.Item("OfferID"), 0).ToString)
      Next
    End If
    
    MyCommon.Close_LogixRT()
    MyCommon.Close_LogixXS()
    
    Return CardTypeOffers
    
  End Function
  
  Function RemoveRestrictedRedemptionOffers(ByRef MyCommon As Copient.CommonInc, ByVal CustomerPK As Long, ByVal AssignedOffers As Data.DataTable) As Data.DataTable
    Dim RestrictedDT As Data.DataTable
    MyCommon.QueryStr = "Select RestrictedRedemption from Customers where CustomerPK=" & CustomerPK & ";"
    RestrictedDT = MyCommon.LXS_Select()
    If (RestrictedDT.Rows.Count > 0 AndAlso MyCommon.NZ(RestrictedDT.Rows(0).Item("RestrictedRedemption"), False)) Then
      Dim RemoveDT As DataRow() = AssignedOffers.Select("RestrictedRedemption=1")
      Dim RemoveCount As Integer = 0
      For RemoveCount = 0 To RemoveDT.Length() - 1
        AssignedOffers.Rows.Remove(RemoveDT(RemoveCount))
        AssignedOffers.AcceptChanges()
      Next
    End If
    
    Return AssignedOffers
  End Function

</script>
<%
done:
  Send_BodyEnd()
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  MyCommon = Nothing
  Logix = Nothing
%>
