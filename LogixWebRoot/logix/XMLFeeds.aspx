<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<script runat="server">
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim LogFile As String = "ModifyGroup." & Format(Now(), "yyyyMMdd") & ".txt"
  </script>
<%
  
  ' *****************************************************************************
  ' * FILENAME: XMLFeeds.aspx 
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
  
  Dim CopientFileName As String = "XMLFeeds.aspx"
  Dim CopientFileVersion As String = "7.3.1.138972" 
  Dim CopientProject As String = "Copient Logix"
  Dim CopientNotes As String = ""
  
  Dim AdminUserID As Long
  Dim rst As DataTable
  Dim row As DataRow
  Dim Category As String = "0"
  Dim Description As String = ""
  Dim CategoryDesc As String = ""
  Dim OfferName As String = ""
  Dim MonthAbbrs(-1) As String
  
  Dim ProductGroupID As String =""
  Dim Products As String =""
  Dim OperationType As Integer =-1
  Dim ProductType As Integer =-1
  Dim GName As String = ""
  Dim RewardID As String = "" 
  Dim IsCondition As Boolean = False
    
  MyCommon.AppName = "XMLFeeds.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
   
  If (LanguageID = 0) Then
    LanguageID = MyCommon.Extract_Val(Request.QueryString("LanguageID"))
  End If
  
 
  If (Request.QueryString("Category") <> "") Then
    Category = MyCommon.Extract_Val(Request.QueryString("Category"))
    Response.Expires = 0
    Response.Clear()
    Response.Expires = 0
    Response.ContentType = "text/xml"
    ' Response.Write("<?xml version='1.0' encoding='ISO-8859-1'?>" & vbCrLf)
    Response.Write("<data>" & vbLf)
    MyCommon.QueryStr = "select TOP 17 DATEPART(YYYY,ProdStartDate) as SDateYear," & _
                        "   DATEPART(mm,ProdStartDate) as SDateMonth," & _
                        "   DATEPART(DD,ProdStartDate) as SDateDay," & _
                        "   DATEPART(YYYY,ProdEndDate) as EDateYear," & _
                        "   DATEPART(mm,ProdEndDate) as EDateMonth," & _
                        "   DATEPART(DD,ProdEndDate) as EDateDay," & _
                        "   Name,OfferID,Description,CategoryDescription " & _
                        " from ( " & _
                        "  select ProdStartDate, ProdEndDate, Name, OfferID, O.Description, C.Description as CategoryDescription " & _
                        "  from Offers as O with (NoLock) " & _
                        "  left join Offercategories as C with (NoLock) on C.OfferCategoryID=O.OfferCategoryID " & _
                            "  where visible=1 and O.deleted=0 and IsTemplate=0 and ProdEndDate > getdate() and O.OfferCategoryID=@Category" & _
                            "  union  " & _
                            "  select StartDate as ProdStartDate, EndDate as ProdEndDate, I.IncentiveName as Name, I.IncentiveID as OfferID, I.Description, C.Description as CategoryDescription " & _
                            "  from CPE_Incentives as I with (NoLock) " & _
                            "  left join Offercategories as C with (NoLock) on C.OfferCategoryID=I.PromoClassID " & _
                            "  where I.deleted=0 and IsTemplate=0 and EndDate > getdate() and I.PromoClassID=@Category" & _
                            ") as OffersInCategory " & _
                            "order by ProdEndDate ASC "
       
        MyCommon.DBParameters.Add("@Category", SqlDbType.Int).Value = Convert.ToInt32(Category)
        rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
    If rst.Rows.Count > 0 Then
      For Each row In rst.Rows
        Description = row.Item("Description").Replace("&", "&amp;")
        OfferName = row.Item("Name").Replace("&", "&amp;")
        CategoryDesc = row.Item("CategoryDescription").Replace("&", "&amp;")
        MonthAbbrs = LoadMonthNames()
        
        'Response.Write(Left(MonthName(row.Item("DateMonth")), 3) & " " & row.Item("DateDay") & " " & row.Item("DateDay"))
        Response.Write("<event start=""" & GetMonthAbbreviation(row.Item("SDateMonth"), MonthAbbrs) & " " & row.Item("SDateDay") & " " & row.Item("SDateYear") & " 00:00:00 GMT""")
        Response.Write(" end=""" & GetMonthAbbreviation(row.Item("EDateMonth"), MonthAbbrs) & " " & row.Item("EDateDay") & " " & row.Item("EDateYear") & " 00:00:00 GMT"" isDuration=""true"" title=""" & OfferName & """ >" & _
        "&lt;a href=""javascript:ChangeParentDocument('" & row.Item("OfferID") & "')""&gt;OfferID:" & row.Item("OfferID") & "&lt;/a&gt; " & Description & _
        " &lt;br&gt; Category:" & CategoryDesc & "</event>")
      Next
    End If
    Response.Write("</data>")
  ElseIf (Request.Form("CollisionGroup") <> "") Then
    Collide(Request.Form("CollisionGroup"), Request.Form("EngineID"))
  ElseIf (Request.QueryString("CollisionGroup") <> "") Then
    Collide(Request.QueryString("CollisionGroup"), Request.QueryString("EngineID"))
  ElseIf (Request.QueryString("OtherOffersCheck") <> "") Then
    OtherOfferscheck(Request.QueryString("OtherOffersCheck"), Request.QueryString("CardPK"), Request.QueryString("OfferID"), Request.QueryString("CustomerGroupID"), Request.QueryString("AdminUserID"))
  ElseIf (Request.QueryString("AddOffer") <> "") Then
    AddOffer(Request.QueryString("AddOffer"), Request.QueryString("CustPK"), Request.QueryString("CardPK"), Request.QueryString("AdminUserID"))
  ElseIf (Request.Form("Reports") <> "") Then
    GenerateReport(Request.Form("Reports"))
  ElseIf (Request.Form("CustomReports") <> "") Then
    GenerateCustomReport(Request.Form("Reports"))
  ElseIf (Request.Form("CPEProductConditionLimits") <> "") Then
    GenerateCPEPRoductConditionLimits(Request.Form("CPEProductConditionLimits"), Request.Form("RewardOptionID"), Request.Form("Disqualifier"))
  ElseIf (Request.QueryString("OfferRedemptions") <> "") Then
    OfferRedemptions(Request.QueryString("CustomerPK"), Request.QueryString("CardPK"), Request.QueryString("OfferID"))
  ElseIf (Request.Form("CPETenderConditionValues") <> "") Then
    GenerateCPETenderConditionValues(Request.Form("CPETenderConditionValues"), Request.Form("RewardOptionID"), Request.Form("excluded"), Request.Form("exVal"))
  ElseIf (Request.QueryString("CAMOfferTransactions") <> "") Then
    CAMOfferTransactions(Long.Parse(Request.QueryString("CustPK")), Long.Parse(Request.QueryString("OfferID")))
  ElseIf (Request.QueryString("HandleLMGSave") <> "") Then
    HandleLMGSave(Request.QueryString("CheckDate"), Request.QueryString("OfferID"))
  'ElseIf (Request.QueryString("AdvSearchQuery") <> "") Then    
    'HandleActionsForAdvSearch(Request.QueryString("AdvSearchQuery")) 
  ElseIf (Request.QueryString("AdvSearchQuery") <> "") Then    
    HandleActionsForAdvSearch() 
  ElseIf (Request.Form("Mode") = "ModifyProductsProductGroups") Then   
    ProductGroupID = Request.Form("ProductGroupID")
    Products = Request.Form("Products")
    OperationType = Convert.ToInt32(Request.Form("OpertaionType"))
    ProductType = Convert.ToInt32(Request.Form("ProductType"))
    Response.Expires = 0
    Response.Clear()
    Response.ContentType = "text/html"
    ModifyProductsProductGroups(ProductGroupID,Products,OperationType,ProductType) 
  ElseIf (Request.Form("Mode") = "ModifyProducts") Then   
    GName = Request.Form("GName")
    RewardID = Request.Form("RewardID")
    Products = Request.Form("Products")
    OperationType = Convert.ToInt32(Request.Form("OperationType"))
    ProductType = Convert.ToInt32(Request.Form("ProductType"))  
    IsCondition = Request.Form("IsCondition")
    ModifyProducts(Products,OperationType,ProductType, GName, RewardID, Request.QueryString("AdminUserID"), IsCondition)
  ElseIf (Request.QueryString("GetResponseAttributes") <> "") Then
        If (Not String.IsNullOrEmpty(Request.QueryString("ConnectorId"))) Then
            GetResponseAttributes(Convert.ToInt32(Request.QueryString("ReferenceId")), Convert.ToInt32(Request.QueryString("ConnectorId")))
        ElseIf (String.IsNullOrEmpty(Request.QueryString("ConnectorId"))) Then
            GetResponseAttributes(Convert.ToInt32(Request.QueryString("ReferenceId")))
        End If
  Else
    Send("<b>" & Copient.PhraseLib.Lookup("term.noarguments", LanguageID) & "</b>")
  End If
        
  MyCommon.Close_LogixRT()
  Logix = Nothing
  MyCommon = Nothing
  Response.Flush()
  Response.End()
%>

<script runat="server">
    Public DefaultLanguageID
    Dim MyCryptLib As New Copient.CryptLib
  
  '----------------------------------------------------------------------------------
  
    Sub GetResponseAttributes(ByVal webMethodId As Integer, Optional ByVal connectorID As Integer = 0)
         LoadFilterData(webMethodId, MyCommon, Logix.UserRoles.EditConnectors, connectorID)
    End Sub
	
	Sub GetReportResponseAttributes(ByVal webMethodId As Integer, Optional ByVal connectorID As Integer = 0)
        LoadFilterData(webMethodId, MyCommon, True, connectorID)
    End Sub
  
  
  Sub OtherOfferscheck(ByVal CustomerPK As String, ByVal CardPK As String, ByVal OfferID As String, ByVal CustomerGroupID As String, ByVal AdminUserID As Integer)
    Dim rst As DataTable
    Dim row As DataRow
    Dim rst2 As DataTable
    Dim rst3 As DataTable
    Dim i As Integer = 0
    Dim OfferDesc As String = ""
    Dim Shaded As String = " class=""shaded"""
    Dim TempDate As Date
    Dim IsCAMOffer As Boolean = False
    Dim UserRoleIDs As Integer()
    Dim RoleMatch As Boolean = False
    Dim UserCanDelete As Boolean = False
    Dim x As Integer = 0
    Dim TempString As String = ""
    
    Response.Cache.SetCacheability(HttpCacheability.NoCache)
    
    Try
      MyCommon.Open_LogixRT()
      
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
      MyCommon.QueryStr = "select * from RolePermissions where RoleID in (select RoleID from AdminUserRoles where AdminUserID=14) and PermissionID=49"
      rst = MyCommon.LRT_Select
      If rst.Rows.Count > 0 Then
        UserCanDelete = True
      End If
      
            MyCommon.QueryStr = "select ICG.CustomerGroupID, CG.EditControlTypeID, CG.RoleID from CPE_RewardOptions as RO with (NoLock) " & _
                                "inner join CPE_IncentiveCustomerGroups as ICG with (NoLock) on ICG.RewardOptionID=RO.RewardOptionID " & _
                                "left join CustomerGroups as CG with (NoLock) on CG.CustomerGroupID=ICG.CustomerGroupID " & _
                                "where RO.IncentiveID=@OfferID and RO.Deleted=0 and ICG.Deleted=0"
            MyCommon.QueryStr += " union " & _
                              "select CG.CustomerGroupID, CG.EditControlTypeID,CG.RoleID " & _
                              "from Offers as O with (NoLock) " & _
                              "left join OfferConditions as OC with (NoLock) on OC.OfferID=O.OfferID " & _
                              "inner join CustomerGroups CG with (NoLock) on CG.CustomerGroupID=LinkID " & _
                              "where(O.Deleted = 0 And O.IsTemplate = 0 And OC.Deleted = 0 And CG.Deleted = 0 And OC.ConditionTypeID = 1 And o.OfferID =@OfferID)"
 
            MyCommon.DBParameters.Add("@OfferID", SqlDbType.Int).Value = OfferID
            rst3 = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            
      If rst3.Rows.Count > 0 Then
        For Each row In rst3.Rows
          i += 1
          RoleMatch = False
          If (MyCommon.NZ(row.Item("EditControlTypeID"), 0) = 3) Then 'it's limited per role, so loop through the user's roles and see if there's a match to the customer group's RoleID 
            For x = 0 To UserRoleIDs.Length - 1
              If UserRoleIDs(x) = MyCommon.NZ(row.Item("RoleID"), 0) Then
                RoleMatch = True
              End If
            Next
          End If
          'This is original line below: If (MyCommon.NZ(row.Item("EditControlTypeID"), 0) = 0 AndAlso UserCanDelete) Or (MyCommon.NZ(row.Item("EditControlTypeID"), 0) = 1) OrElse (MyCommon.NZ(row.Item("EditControlTypeID"), 0) = 3 AndAlso RoleMatch) Then
          'I'd originally added the "UserCanDelete" check, but dropped it in the final version because a) I think it's useful to see all the IDs with which the groups is associated,
          'regardless of the EditControlTypeIDs of those groups, and b) the main offers page should now throw an infomessage if the customer cannot be dropped from every associated group.
          If (MyCommon.NZ(row.Item("EditControlTypeID"), 0) = 0) Or (MyCommon.NZ(row.Item("EditControlTypeID"), 0) = 1) OrElse (MyCommon.NZ(row.Item("EditControlTypeID"), 0) = 3 AndAlso RoleMatch) Then
            TempString &= MyCommon.NZ(row.Item("CustomerGroupID"), 0)
            If i < rst3.Rows.Count Then
              TempString &= ", "
            End If
          End If
        Next
      End If
      Send("<center>" & Copient.PhraseLib.Detokenize("customer-inquiry.CustomerRemovalWarning", LanguageID, TempString) & "</center>")
      
      Send("<table summary=""" & Copient.PhraseLib.Lookup("term.offers", LanguageID) & """>")
      Send("  <thead>")
      Send("    <tr>")
      Send("      <th class=""th-id"" scope=""col"">" & Copient.PhraseLib.Lookup("term.id", LanguageID) & "</th>")
      Send("      <th class=""th-name"" scope=""col"">" & Copient.PhraseLib.Lookup("term.name", LanguageID) & "</th>")
      Send("    </tr>")
      Send("  </thead>")
      Send("  <tbody>")
      
      ' ok here we have the user PK we need to figure out what promotions they are in again to remove them from
            MyCommon.QueryStr = "select distinct Name,O.Description,O.OfferID,O.ProdEndDate,null as BuyerID from Offers as O with (NoLock) " & _
                                "left join OfferConditions as OC with (NoLock) on OC.OfferID=O.OfferID " & _
                                "where O.IsTemplate=0 and OC.ConditionTypeID=1 and OC.LinkID=@CustomerGroupID or OC.ExcludedID=@CustomerGroupID " & _
                                " union " & _
                                "select distinct I.IncentiveName as Name,I.Description,I.IncentiveID as OfferID,I.EndDate as ProdEndDate,buy.ExternalBuyerId as BuyerID " & _
                                "from CPE_IncentiveCustomerGroups ICG with (NoLock) " & _
                                "inner join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID = ICG.RewardOptionID " & _
                                "inner join CPE_Incentives I with (NoLock) on I.IncentiveID = RO.IncentiveID " & _
                                 " left outer join Buyers as buy with (nolock) on buy.BuyerId= I.BuyerId " & _
                                "where ICG.Deleted=0 and RO.Deleted=0 and I.Deleted=0 and I.IsTemplate=0 and ICG.CustomerGroupID=@CustomerGroupID " & _
                                "order by ProdEndDate ASC;"
            
            MyCommon.DBParameters.Add("@CustomerGroupID", SqlDbType.Int).Value = CustomerGroupID
            rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            
      For Each row In rst.Rows
    Dim Name As String=""
                    If (MyCommon.Fetch_UE_SystemOption(168) = "1" AndAlso MyCommon.NZ(row.Item("BuyerID"), "") <> "") Then
                    Name = "Buyer " + row.Item("BuyerID").ToString() + " - " + MyCommon.SplitNonSpacedString(row.Item("Name"), 25).ToString()
                Else
                    Name = MyCommon.NZ(MyCommon.SplitNonSpacedString(row.Item("Name"), 25), Copient.PhraseLib.Lookup("term.unknown", LanguageID))
                End If
        Send("    <tr" & Shaded & ">")
        Send("      <td valign=""top"">" & row.Item("OfferID") & "</td>")
        Sendb("      <td valign=""top"" title=""" & row.Item("Description") & """>" & Name)
        
        ' adjust the production end date as it's time is store at midnight, but the offer is still valid during the entire day of the end date.
        If (Date.TryParse(MyCommon.NZ(row.Item("ProdEndDate"), Now).ToString, TempDate)) Then
          TempDate = TempDate.AddDays(1)
          If (TempDate < Now) Then
            Sendb("<span style=""color:#ff0000;""> (" & Copient.PhraseLib.Lookup("term.expired", LanguageID) & ") </span>")
          End If
        End If
        
        ' Find and insert the offer's description
        MyCommon.QueryStr = "select EngineID from OfferIDs with (NoLock) where OfferID = " & row.Item("OfferID")
        rst2 = MyCommon.LRT_Select
        Select Case MyCommon.NZ(rst2.Rows(0).Item("EngineID"), 0)
          Case 0, 1
            MyCommon.QueryStr = "select Description from Offers with (NoLock) where OfferID = " & row.Item("OfferID")
          Case 6
            IsCAMOffer = True
            MyCommon.QueryStr = "select Description from CPE_Incentives with (NoLock) where IncentiveID = " & row.Item("OfferID")
          Case Else
            MyCommon.QueryStr = "select Description from CPE_Incentives with (NoLock) where IncentiveID = " & row.Item("OfferID")
        End Select
        rst2 = MyCommon.LRT_Select
        If (rst2.Rows.Count > 0) Then
          OfferDesc = MyCommon.NZ(rst2.Rows(0).Item("Description"), "")
        End If
        If (OfferDesc <> "") Then
          Sendb("<br /><small style=""color:#444444;"">" & OfferDesc & "</small>")
        End If
        Send("      </td>")
        Send("    </tr>")
        If Shaded = "" Then
          Shaded = " class=""shaded"""
        Else
          Shaded = ""
        End If
      Next
      Send("    <tr>")
      Send("      <td></td>")
      Send("      <td></td>")
      Send("    </tr>")
      Send("  </tbody>")
      Send("</table>")
      If IsCAMOffer Then
        Send("<form action=""" & GetIPortalQualifier() & "/logix/CAM/CAM-customer-offers.aspx"" method=""GET"" id=""inquiryform"" name=""inquiryform"">")
      Else
        Send("<form action=""" & GetIPortalQualifier() & "/logix/customer-offers.aspx"" method=""GET"" id=""inquiryform"" name=""inquiryform"">")
      End If
      Send("  <input type=""hidden"" id=""CustomerGroupID"" name=""CustomerGroupID"" value=""" & CustomerGroupID & """ />")
      Send("  <input type=""hidden"" id=""CustomerPK"" name=""CustomerPK"" value=""" & CustomerPK & """ />")
      Send("  <input type=""hidden"" id=""CustPK"" name=""CustPK"" value=""" & CustomerPK & """ />")
      If CardPK > 0 Then
        Send("  <input type=""hidden"" id=""CardPK"" name=""CardPK"" value=""" & CardPK & """ />")
      End If
      Send("  <input type=""hidden"" id=""OfferID"" name=""OfferID"" value=""" & OfferID & """ />")
      Send("  <input type=""hidden"" id=""search"" name=""search"" value=""" & Copient.PhraseLib.Lookup("term.search", LanguageID) & """ />")
      Send("  <input type=""submit"" class=""regular"" id=""RemoveFromOffer"" name=""RemoveFromOffer"" value=""" & Copient.PhraseLib.Lookup("term.remove", LanguageID) & """ />")
      Send("</form>")
      Send("<small>" & Logix.ToShortDateTimeString(DateTime.Now, MyCommon) & "</small>")
    Catch ex As Exception
      Send(ex.ToString)
    Finally
      MyCommon.Close_LogixRT()
    End Try
  End Sub
  
  '----------------------------------------------------------------------------------
  
  Sub AddOffer(ByVal OfferID As String, ByVal CustomerPK As String, ByVal CardPK As String, ByVal AdminUserID As Integer)
    Dim rst As DataTable
    Dim row As DataRow
    Dim CustomerGroupID As String = ""
    Dim ExtCardID As String = ""
    Dim CGList As String = ""
    Dim TempDate As Date
    Dim IsCAMOffer As Boolean = False
    Dim UserRoleIDs As Integer()
    Dim RoleMatch As Boolean = False
    Dim x As Integer = 0
    Dim UserCanAdd As Boolean = False
    
    Response.Cache.SetCacheability(HttpCacheability.NoCache)
    Try
      'Get the ExtCardID
      MyCommon.Open_LogixXS()
      If (MyCommon.Extract_Val(Request.QueryString("CardPK")) > 0) Then
        MyCommon.QueryStr = "select ExtCardID from CardIDs with (NoLock) where CardPK=@CardPK"
                MyCommon.DBParameters.Add("@CardPK", SqlDbType.Int).Value = CardPK
                rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)
        If rst.Rows.Count > 0 Then
                    ExtCardID = MyCryptLib.SQL_StringDecrypt(rst.Rows(0).Item("ExtCardID").ToString())
        End If
      End If
      MyCommon.Close_LogixXS()
      MyCommon.Open_LogixRT()
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
      MyCommon.QueryStr = "select * from RolePermissions where RoleID in (select RoleID from AdminUserRoles where AdminUserID=14) and PermissionID=49"
      rst = MyCommon.LRT_Select
      If rst.Rows.Count > 0 Then
        UserCanAdd = True
      End If
      
      Send("<center>")
      Send(Copient.PhraseLib.Detokenize("xmlfeeds.AddingToOffer", LanguageID, Copient.PhraseLib.Lookup("term." & IIf(CardPK > 0, "card", "customer"), LanguageID), OfferID))
      
      ' Figure out what customer group is used by the offer so we can determine what else it's in
            MyCommon.QueryStr = "select OC.LinkID, CG.Name, CG.EditControlTypeID, CG.RoleID from OfferConditions as OC with (NoLock) " & _
                                "left join CustomerGroups as CG with (NoLock) on OC.LinkID=CG.CustomerGroupID " & _
                                "where ConditionTypeID=1 and OC.Deleted=0 and OfferID=@OfferID " & _
                                " union " & _
                                "select ICG.CustomerGroupID as LinkID, CG.Name, CG.EditControlTypeID, CG.RoleID from CPE_IncentiveCustomerGroups as ICG with (NoLock) " & _
                                "left join CustomerGroups as CG with (NoLock) on ICG.CustomerGroupID=CG.CustomerGroupID " & _
                                "inner join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID=ICG.RewardOptionID " & _
                                "inner join CPE_Incentives I with (NoLock) on I.IncentiveID=RO.IncentiveID " & _
                                "where ICG.ExcludedUsers=0 AND ICG.Deleted=0 and CG.Deleted=0 and RO.Deleted=0 and I.Deleted=0 and I.IncentiveID=@OfferID;"
            
            MyCommon.DBParameters.Add("@OfferID", SqlDbType.Int).Value = OfferID
            rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            
      If (rst.Rows.Count > 0) Then
        For Each row In rst.Rows
          RoleMatch = False
          If (MyCommon.NZ(row.Item("EditControlTypeID"), 0) = 3) Then 'it's limited per role, so loop through the user's roles and see if there's a match to the customer group's RoleID 
            For x = 0 To UserRoleIDs.Length - 1
              If UserRoleIDs(x) = MyCommon.NZ(row.Item("RoleID"), 0) Then
                RoleMatch = True
              End If
            Next
          End If
          If (MyCommon.NZ(row.Item("EditControlTypeID"), 0) = 0) Or (MyCommon.NZ(row.Item("EditControlTypeID"), 0) = 1) OrElse (MyCommon.NZ(row.Item("EditControlTypeID"), 0) = 3 AndAlso RoleMatch) Then
            CustomerGroupID = row.Item("LinkID")
            CGList += "," & CustomerGroupID
            Send("<br />" & CustomerGroupID & " - " & row.Item("Name"))
          End If
        Next
        ' Remove the leading "," if necessary
        If (CGList.Length > 0) Then
          CGList = CGList.Substring(1)
        End If
      End If
      If (CGList.Trim = "") Then CGList = "-1"
      
      MyCommon.QueryStr = "select distinct Name,O.Description,O.OfferID,O.ProdEndDate, EngineID from Offers as O with (NoLock) " & _
                          "left join OfferConditions as OC with (NoLock) on OC.OfferID=O.OfferID " & _
                          "where O.IsTemplate=0 and OC.ConditionTypeID=1 and OC.Linkid in (" & CGList & ") or OC.ExcludedID in (" & CGList & ") " & _
                          " union " & _
                          "select distinct I.IncentiveName,I.Description,I.IncentiveID as OfferID,I.EndDate as ProdEndDate, EngineID " & _
                          "from CPE_IncentiveCustomerGroups ICG with (NoLock) " & _
                          "inner join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID = ICG.RewardOptionID " & _
                          "inner join CPE_Incentives I with (NoLock) on I.IncentiveID = RO.IncentiveID " & _
                          "where  ICG.Deleted=0 and RO.Deleted=0 and I.Deleted=0 and I.IsTemplate=0 and ICG.CustomerGroupID IN (" & CGList & ") " & _
                          "order by ProdEndDate ASC;"
      rst = MyCommon.LRT_Select
      Send("<table summary=""" & Copient.PhraseLib.Lookup("term.group", LanguageID) & """>")
      Send("  <thead>")
      Send("    <tr>")
      Send("      <th class=""th-id"" scope=""col"">" & Copient.PhraseLib.Lookup("term.id", LanguageID) & "</th>")
      Send("      <th class=""th-name"" scope=""col"">" & Copient.PhraseLib.Lookup("term.name", LanguageID) & "</th>")
      Send("    </tr>")
      Send("  </thead>")
      Send("  <tbody>")
      For Each row In rst.Rows
        If (MyCommon.NZ(row.Item("EngineID"), 0) = 6) Then IsCAMOffer = True
        Send("    <tr>")
        Send("      <td>" & row.Item("OfferID") & "</td>")
        Sendb("      <td title=""" & row.Item("Description") & """>" & row.Item("Name"))
        ' adjust the production end date as it's time is store at midnight, but the offer is still valid during the entire day of the end date.
        If (Date.TryParse(MyCommon.NZ(row.Item("ProdEndDate"), Now).ToString, TempDate)) Then
          TempDate = TempDate.AddDays(1)
          If (TempDate < Now) Then
            Sendb("<span style=""color:#ff0000;""> (" & Copient.PhraseLib.Lookup("term.expired", LanguageID) & ") </span>")
          End If
        End If
        Send("</td>")
        Send("    </tr>")
      Next
      Send("    <tr>")
      Send("      <td></td>")
      Send("      <td></td>")
      Send("    </tr>")
      Send("  </tbody>")
      Send("</table>")
      If IsCAMOffer Then
        Send("<form action=""" & GetIPortalQualifier() & "/logix/CAM/CAM-customer-addoffer.aspx"" method=""GET"" id=""inquiryform"" name=""inquiryform"" >")
      Else
        Send("<form action=""" & GetIPortalQualifier() & "/logix/customer-addoffer.aspx"" method=""GET"" id=""inquiryform"" name=""inquiryform"" >")
      End If
      Send("  <input type=""hidden"" id=""CustomerGroupID"" name=""CustomerGroupID"" value=""" & CGList & """ />")
      Send("  <input type=""hidden"" id=""SelectedOfferID"" name=""SelectedOfferID"" value=""" & OfferID & """ />")
      Send("  <input type=""hidden"" id=""CustomerPK"" name=""CustomerPK"" value=""" & CustomerPK & """ />")
      If CardPK > 0 Then
        Send("  <input type=""hidden"" id=""CardPK"" name=""CardPK"" value=""" & CardPK & """ />")
      End If
      Send("  <input type=""hidden"" id=""searchterms"" name=""searchterms"" value=""" & CustomerPK & """ />")
      Send("  <input type=""hidden"" id=""search"" name=""search"" value=""" & Copient.PhraseLib.Lookup("term.search", LanguageID) & """ />")
      Send("  <input type=""submit"" class=""regular"" id=""AddOffer"" name=""AddOffer"" value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """ />")
      Send("</form>")
      Send("<small>" & Logix.ToShortDateTimeString(DateTime.Now, MyCommon) & "</small>")
    Catch ex As Exception
      Send(ex.ToString)
    Finally
      MyCommon.Close_LogixRT()
    End Try
  End Sub
  
  '----------------------------------------------------------------------------------
  Sub HandleActionsForAdvSearch()
    Dim dt As DataTable
    Dim rst As DataTable
    Dim row As DataRow
    Dim OfferID As Integer
    Dim Name As String = ""
    Dim Engine As String = ""
    Dim OfferStartDate As Date
    Dim ExternalOfferID As String
    Dim OfferidsAll As String = ""
        Dim Index As Integer = 1
    Dim CustomerPK As Long = 1
    Dim advsearchquery As String  
    
    Response.Cache.SetCacheability(HttpCacheability.NoCache)
    If Session("AdvSearchquery") IsNot Nothing Then
      advsearchquery = Session("AdvSearchquery")    
    Else
      Send("<div>Session timed out. Please try again.</div>")    
      Exit Sub
    End If
   
    Try
      MyCommon.Open_LogixRT()
      MyCommon.QueryStr = advsearchquery                
      
      Dim _unnamed = Copient.PhraseLib.Lookup("term.unnamed", LanguageID)
      Send("<div id=""ActionsTB"">")
      Send("<table summary="""" style=""width:100%;"">")
        Send("  <tr>")
        Send("<td align=""right"">")              
        Send("<input type=""image"" src=""../images/folders/deploy.png"" alt=""" & Copient.PhraseLib.Lookup("folders.performaction", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("folders.performaction", LanguageID) & """ onclick=""javascript:clearErrorContents(); javascript:showactions();"" />")                  
        Send("</td>")
        Send("  </tr>")
        Send("</table>")        
        dt = MyCommon.LRT_Select
        Send("<br class=""half"" />")        
        Send("<table id= ""tb1"" summary="""" style=""width:100%;"">")
        Send("  <tr>")        
        Send("    <th scope=""col"" ><input name=""allofferIDs"" id=""allofferIDs"" type=""checkbox"" title=""" & Copient.PhraseLib.Lookup("hierarchy.SelectAllItems", LanguageID) & """ onclick=""javascript:handleAllItems();"" /></th>")                
        Send("    <th scope=""col"" style=""width:60px;"" >" & Copient.PhraseLib.Lookup("term.offerid", LanguageID) & "</th>")
        Send("    <th scope=""col"" >" & Copient.Phraselib.Lookup("tag.offerstartdate", LanguageID) & "</th>")
        Send("    <th scope=""col"" >" & Copient.PhraseLib.Lookup("term.xid", LanguageID) & "</th>")
        Send("    <th scope=""col"" >" & Copient.PhraseLib.Lookup("term.engine", LanguageID) & "</th>")
        Send("    <th scope=""col"" >" & Copient.PhraseLib.Lookup("term.name", LanguageID) & "</th>")                                     
        Send("  </tr>")
      Dim strBuf As New StringBuilder
      If dt.Rows.Count > 0 Then
        For Each row In dt.Rows
          
          OfferID = MyCommon.NZ(row.Item("OfferID"), 0)
          Name = MyCommon.NZ(row.Item("Name"), "")
          OfferStartDate = MyCommon.NZ(row.Item("ProdStartDate"), New DateTime(1990, 1, 1))
          ExternalOfferID = MyCommon.NZ(row.Item("ExtOfferID"), "")
                    'If Index < dt.Rows.Count Then
                    '  Index += 1
                    OfferidsAll = OfferidsAll & OfferID & ","
                    'Else
                    'OfferidsAll = OfferidsAll & OfferID
                    'End If
         
         
                    Engine = MyCommon.NZ(row.Item("PromoEngine"), "")
                    strBuf.Append("  <tr>" & _
                  "    <td>" & _
                  "      <input name=""linkID"" id=""linkID" & OfferID & """ type=""checkbox"" value=""" & OfferID & """ onclick=""submitToperformaction(" & OfferID & ", this.checked);"" />" & _
                  "    </td>" & _
                  "    <td>" & OfferID & "</td>" & _
                  "    <td>" & OfferStartDate & "</td>" & _
                  "    <td>" & ExternalOfferID & "</td>" & _
                  "    <td>" & Engine & "</td>" & _
                  "    <td><a href=""offer-redirect.aspx?OfferID=" & OfferID & """ target=""_blank"">" & IIf(Name <> "", Name, "(" & _unnamed & ")") & "</a></td>" & _
                  "  </tr>" & _
                  "  </tr>" & _
                  "<tr name=""errdesc"" id=""errdesc" & OfferID & """>" & _
                  "</tr>")
                Next
                Send(strBuf.ToString())
                Send("</table>")
                Send("</center>")
                Send("<input type=""hidden"" id=""itemlist"" name=""itemlist"" value=""" & OfferidsAll.Remove(OfferidsAll.LastIndexOf(","), 1) & """ />")
                Send("</div>")
            End If
        Catch ex As Exception
            Send(ex.ToString)
        Finally
            'If Session("AdvSearchquery") IsNot Nothing Then
            'Session.Remove("AdvSearchquery")
            'End If
            MyCommon.Close_LogixRT()
        End Try
  End Sub
 
    '----------------------------------------------------------------------------------
  
  Sub Collide(ByVal Group As String, ByVal PromoEngine As Integer)
    Dim rst As DataTable
    Dim row As DataRow
    Dim shaded As Boolean = True
    Dim FIRST As Boolean = True
    Dim Now As DateTime = DateTime.Now
    Dim builder As StringBuilder = New StringBuilder()
    Dim PrevExtProdID As String = ""
     Dim RewardAmtStr As String = ""
    Dim Localizer As Copient.Localization
    Localizer = New Copient.Localization(MyCommon)
    ' Default to PromoEngine 0 (CM) if not specified
    If (PromoEngine <> 0 And PromoEngine <> 2 And PromoEngine <> 9) Then PromoEngine = 0
    
    MyCommon.Open_LogixRT()

    If PromoEngine = 2  or PromoEngine=9 Then
      MyCommon.QueryStr = "dbo.pa_CPE_ProductGroup_Collision"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = PromoEngine
    Else
      MyCommon.QueryStr = "dbo.pa_CM_ProductGroup_Collision"
      MyCommon.Open_LRTsp()
    End If
    MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt).Value = Long.Parse(Group)
    rst = MyCommon.LRTsp_select
    MyCommon.Close_LRTsp()

    If rst.Rows.Count > 0 Then
      builder.Append("<table class=""list"" summary=""" & Copient.PhraseLib.Lookup("term.results", LanguageID) & """>")
      builder.Append("    <thead>")
      builder.Append("        <tr class=""shaded"">")
      builder.Append("            <th align=""left"" scope=""col"" class=""th-upc"">" & Copient.PhraseLib.Lookup("term.upc", LanguageID) & "</th>")
      builder.Append("            <th align=""left"" scope=""col"" class=""th-pgroup"">" & Copient.PhraseLib.Lookup("term.productgroup", LanguageID) & "</th>")
      builder.Append("            <th align=""left"" scope=""col"" class=""th-offer"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & "</th>")
      builder.Append("            <th align=""right"" scope=""col"" class=""th-amount"" style=""text-align:right;"">" & Copient.PhraseLib.Lookup("term.amount", LanguageID) & "</th>")
      builder.Append("        </tr>")
      builder.Append("    </thead>")
      builder.Append("    <tbody>")
            For Each row In rst.Rows
                  Dim Amount As String = MyCommon.NZ(row.Item("RewardAmount"), "").ToString.Substring(0, MyCommon.NZ(row.Item("RewardAmount"), " ").ToString.Length - 1)
                If PromoEngine = 2  or PromoEngine=9 Then
                Dim AmountTypeID As Integer = (MyCommon.NZ(row.Item("AmountTypeID"), 0))
              
                Dim RewardID As Long = MyCommon.NZ(row.Item("ROID"), 0)
                
                Select Case AmountTypeID
                    Case 1, 5, 9, 10, 11, 12
                        RewardAmtStr &= (Localizer.FormatCurrency_ForOffer(CDec(Amount), RewardID).ToString(MyCommon.GetAdminUser.Culture) & "&nbsp;")
                    Case 3
                        RewardAmtStr &= (Math.Round(CDec(Amount), 2).ToString(MyCommon.GetAdminUser.Culture) & "% " & "&nbsp;")
                    Case 4
                        RewardAmtStr &= (Copient.PhraseLib.Lookup("term.free", LanguageID) & "&nbsp;")
                    Case 2, 6, 13, 14, 15, 16
                        RewardAmtStr &= (Localizer.GetCached_Currency_Symbol(RewardID) & Math.Round(CDec(Amount), Localizer.GetCached_Currency_Precision(RewardID)).ToString(MyCommon.GetAdminUser.Culture) & "&nbsp;")
                    Case Else
                        RewardAmtStr &= (Amount & "&nbsp;")
                End Select
 Else
                   RewardAmtStr &= (Amount & "&nbsp;")
                End If
                               If (shaded) Then
                    builder.Append("        <tr class=""shaded"">")
                    shaded = False
                Else
                    builder.Append("        <tr>")
                    shaded = True
                End If
                If PrevExtProdID <> MyCommon.NZ(row.Item("ExtProductID"), "") Then
                    builder.Append("            <td>" & MyCommon.NZ(row.Item("ExtProductID"), "") & "</td>")
                Else
                    builder.Append("            <td></td>")
                End If
                builder.Append("            <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("GroupName"), ""), 25) & "</td>")
                If (PromoEngine = 0) Then
                    builder.Append("            <td><a href=""" & GetIPortalQualifier() & "offer-sum.aspx?OfferID=" & MyCommon.NZ(row.Item("PromoID"), "") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("PromoName"), ""), 25) & "</a></td>")
                ElseIf (PromoEngine = 2) Then
                    builder.Append("            <td><a href=""" & GetIPortalQualifier() & "CPEoffer-sum.aspx?OfferID=" & MyCommon.NZ(row.Item("PromoID"), "") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("PromoName"), ""), 25) & "</a></td>")
                ElseIf (PromoEngine = 9) Then
                    Dim GName As String = ""
                    If (MyCommon.Fetch_UE_SystemOption(168) = "1" AndAlso MyCommon.NZ(row.Item("BuyerID"), "") <> "") Then
                        GName = "Buyer " + row.Item("BuyerID").ToString() + " - " + MyCommon.NZ(row.Item("PromoName"), "").ToString()
                    Else
                        GName = MyCommon.NZ(row.Item("PromoName"), Copient.PhraseLib.Lookup("term.unknown", LanguageID))
                    End If
                    builder.Append("            <td><a href=""" & GetIPortalQualifier() & "UE/UEoffer-sum.aspx?OfferID=" & MyCommon.NZ(row.Item("PromoID"), "") & """>" & MyCommon.SplitNonSpacedString(GName, 25) & "</a></td>")
                End If
                builder.Append("            <td align=""right"">" & RewardAmtStr & "</td>")
                builder.Append("        </tr>")
                PrevExtProdID = MyCommon.NZ(row.Item("ExtProductID"), "")
            Next
      builder.Append("    </tbody>")
      builder.Append("</table>")
      'builder.Insert(0, "<div id=""infobar"" class=""green-background"">" & Copient.PhraseLib.Lookup("product-inquiry.reportcompleted", LanguageID) & ": " & QueryCount & " " & StrConv(Copient.PhraseLib.Lookup("term.queries", LanguageID), VbStrConv.Lowercase) & ", " & (DateTime.Now - Now).TotalSeconds.ToString & " " & StrConv(Copient.PhraseLib.Lookup("term.seconds", LanguageID), VbStrConv.Lowercase) & "</div>")
      Response.Write(builder)
    Else
      Response.Write("<div id=""infobar"" class=""red-background"">" & Copient.PhraseLib.Lookup("product-inquiry.noitems", LanguageID) & "</div>")
    End If
    
    MyCommon.Close_LogixRT()
  End Sub
  
  '----------------------------------------------------------------------------------
  
  Sub GenerateCPEPRoductConditionLimits(ByVal limits As String, ByVal roid As String, ByVal Disqualifier As Integer)
    Dim rst As DataTable
    Dim row As DataRow
    Dim rst2 As DataTable
    Dim i As Integer
    Dim bPArsed As Integer
    Dim Ids() As String
    Dim rst3 As DataTable
    Dim row3 As DataRow
    Dim TierDT As DataTable
    Dim TierDR As DataRow
    Dim ProductComboID As Integer = 1
    Dim Qty As Decimal
    Dim Type As Integer
    Dim AccumMin As Decimal
    Dim AccumLimit As Decimal
    Dim AccumPeriod As Integer
    Dim ShowAccumMsg As Boolean = False
    Dim CleanGroupName As String = ""
    Dim IsItem As Boolean = False
    Dim IsDollar As Boolean = False
    Dim Shaded As String = " class=""shaded"""
    Dim UniqueChecked As Boolean = False
    Dim NetPriceChecked As Boolean = False
    Dim TierLevel As Integer = 1
    Dim t As Integer
    Dim TierQty As Decimal = 0
    
    If (Request.Form("LanguageID") <> "") Then
      bPArsed = Integer.TryParse(Request.Form("lang"), LanguageID)
      If (Not bPArsed) Then LanguageID = 1
    End If
    
    MyCommon.Open_LogixRT()
    
        MyCommon.QueryStr = "select TierLevels from CPE_RewardOptions where RewardOptionID=@roid and TouchResponse=0 and Deleted=0;"
        
        MyCommon.DBParameters.Add("@roid", SqlDbType.Int).Value = roid
        rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        
    TierLevel = MyCommon.NZ(rst.Rows(0).Item("TierLevels"), 1)
    
    ' Preload the unit types
    MyCommon.QueryStr = "select UnitTypeID, PhraseID from CPE_UnitTypes UT with (NoLock)"
    rst3 = MyCommon.LRT_Select
    
    Send("<table summary=""" & Copient.PhraseLib.Lookup("term.groups", LanguageID) & """>")
    Send("    <thead>")
    Send("        <tr>")
    Send("            <th class=""th-group"" scope=""col"">" & Copient.PhraseLib.Lookup("term.group", LanguageID) & "</th>")
    Send("            <th class=""th-quantity"" scope=""col"">" & Copient.PhraseLib.Lookup("term.quantity", LanguageID) & "</th>")
    Send("            <th class=""th-unit"" scope=""col"">" & Copient.PhraseLib.Lookup("term.unit", LanguageID) & "</th>")
    Send("        </tr>")
    Send("    </thead>")
    Send("    <tbody>")
        MyCommon.QueryStr = "Select ProductGroupID,Name from ProductGroups with (NoLock) where ProductGroupID in (SELECT items FROM Split (@ProductGroupIDs, ',')) order by Name"
        MyCommon.DBParameters.Add("@ProductGroupIDs", SqlDbType.NVarChar).Value = limits
        rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
    ' Do regular limit stuff
    For Each row In rst.Rows
      ' Determine any limits associated with the current working group
      MyCommon.QueryStr = "Select IncentiveProductGroupID,QtyForIncentive,QtyUnitType,AccumMin,AccumLimit,AccumPeriod,UniqueProduct from CPE_IncentiveProductGroups with (NoLock) " & _
                                "where deleted=0 and RewardOptionID=@roid and ProductGroupID=@ProductGroupID"
      '; Send(MyCommon.QueryStr)
            MyCommon.DBParameters.Add("@roid", SqlDbType.Int).Value = roid
            MyCommon.DBParameters.Add("@ProductGroupID", SqlDbType.Int).Value = row.Item("ProductGroupID")
            rst2 = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
     
      If rst2.Rows.Count > 0 Then
        Qty = MyCommon.NZ(rst2.Rows(0).Item("QtyForIncentive"), 1)
        Type = MyCommon.NZ(rst2.Rows(0).Item("QtyUnitType"), 1)
        AccumMin = MyCommon.NZ(rst2.Rows(0).Item("AccumMin"), 0)
        AccumLimit = MyCommon.NZ(rst2.Rows(0).Item("AccumLimit"), 0)
        AccumPeriod = MyCommon.NZ(rst2.Rows(0).Item("AccumPeriod"), 0)
        UniqueChecked = MyCommon.NZ(rst2.Rows(0).Item("UniqueProduct"), False)
        NetPriceChecked = MyCommon.NZ(rst2.Rows(0).Item("NetPriceProduct"), False)
        If Type = 1 Then
          IsItem = True
        ElseIf Type = 2 Then
          IsDollar = True
        End If
      Else
        Qty = 1
        Type = 1
        AccumMin = 0
        AccumLimit = 0
        AccumPeriod = 0
      End If
      CleanGroupName = MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 15)
      Send("        <tr " & Shaded & ">")
      Send("            <td><label for=""t" & t & "_limit-" & row.Item("ProductGroupID") & """>" & CleanGroupName & "</label></td>")

      Send("<td>")
      For t = 1 To TierLevel
        If rst2.Rows.Count > 0 Then       
                    MyCommon.QueryStr = "Select Quantity from CPE_IncentiveProductGroupTiers with (NoLock) where RewardOptionID=@roid and TierLevel=@TierLevel and IncentiveProductGroupID=@IncentiveProductGroupID"
                    MyCommon.DBParameters.Add("@roid", SqlDbType.Int).Value = roid
                    MyCommon.DBParameters.Add("@TierLevel", SqlDbType.Int).Value = t
                    MyCommon.DBParameters.Add("@IncentiveProductGroupID", SqlDbType.Int).Value = rst2.Rows(0).Item("IncentiveProductGroupID")
                    TierDT = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
          
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
        ElseIf IsDollar Then
          TierQty = Math.Round(TierQty, 2)
        End If
        If TierLevel > 1 Then
          Sendb("            <label for=""t" & t & "_limit-" & row.Item("ProductGroupID") & """>" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</label>")
        End If
        Send("            <input type=""text"" class=""shorter"" maxlength=""9"" name=""t" & t & "_limit-" & row.Item("ProductGroupID") & """ id=""t" & t & "_limit-" & row.Item("ProductGroupID") & """ value=""" & TierQty & """ />" & IIf(TierLevel > 1, "<br />", ""))
      Next
      Send("</td>")
      Send("            <td><select name=""select-" & row.Item("ProductGroupID") & """ id=""select-" & row.Item("ProductGroupID") & """>")
      For Each row3 In rst3.Rows
        Sendb("<option")
        If (Type = row3.Item("UnitTypeID")) Then
          Sendb(" selected=""selected""")
        End If
        Sendb(" value=""" & row3.Item("UnitTypeID") & """>" & Copient.PhraseLib.Lookup(row3.Item("PhraseID"), LanguageID))
        Send("</option>")
      Next
      Send("            </td>")
      Send("        </tr>")
      Send("        <tr " & Shaded & "><td></td><td colspan=""2"">")
      Send("        <input type=""checkbox"" id=""Unique-" & row.Item("ProductGroupID") & """ name=""Unique-" & row.Item("ProductGroupID") & """ " & IIf(UniqueChecked, "checked=""checked""", "") & """ value=""1""/><label for=""Unique-" & row.Item("ProductGroupID") & """ id=""SelectUnique-" & row.Item("ProductGroupID") & """>" & Copient.PhraseLib.Lookup("term.uniqueproduct", LanguageID) & "</label>")
      If MyCommon.Fetch_UE_SystemOption(210) = "1" Then
        Send("        <input type=""checkbox"" id=""NetPrice-" & row.Item("ProductGroupID") & """ name=""NetPrice-" & row.Item("ProductGroupID") & """ " & IIf(NetPriceChecked, "checked=""checked""", "") & """ value=""1""/><label for=""NetPrice-" & row.Item("ProductGroupID") & """ id=""SelectNetPrice-" & row.Item("ProductGroupID") & """>" & Copient.PhraseLib.Lookup("term.netpriceproduct", LanguageID) & "</label>")
      End If
      Send("        </td></tr>")
      rst2 = Nothing
      If Shaded = " class=""shaded""" Then
        Shaded = ""
      Else
        Shaded = " class=""shaded"""
      End If
    Next
    Send("    </tbody>")
    Send("</table>")
    Send("<br />")

    
    If Disqualifier = 0 Then
      If rst.Rows.Count = 1 Then
        ' Do the accumulation stuff here
        Send("<b>" & Copient.PhraseLib.Lookup("term.accumulation", LanguageID) & "</b><br />")
        Send("<table summary=""" & Copient.PhraseLib.Lookup("term.accumulation", LanguageID) & """>")
        Send("    <thead>")
        Send("        <tr>")
        Send("            <th class=""th-minimum"" scope=""col"">" & Copient.PhraseLib.Lookup("term.minimum", LanguageID) & "</th>")
        Send("            <th class=""th-limit"" scope=""col"">" & Copient.PhraseLib.Lookup("term.limit", LanguageID) & "</th>")
        Send("            <th class=""th-period"" scope=""col"">" & Copient.PhraseLib.Lookup("term.period", LanguageID) & "</th>")
        Send("        </tr>")
        Send("    </thead>")
        Send("    <tbody>")
        Send("        <tr>")
        If IsItem Then
          AccumMin = Math.Truncate(AccumMin)
        ElseIf IsDollar Then
          AccumMin = Math.Round(AccumMin, 2)
        End If
        Send("            <td><input type=""text"" class=""shorter"" maxlength=""9"" name=""accummin-" & row.Item("ProductGroupID") & """ id=""accummin-" & row.Item("ProductGroupID") & """ value=""" & AccumMin & """ /></td>")
        If IsItem Then
          AccumLimit = Math.Truncate(AccumLimit)
        ElseIf IsDollar Then
          AccumLimit = Math.Round(AccumLimit, 2)
        End If
        Send("            <td><input type=""text"" class=""shorter"" maxlength=""9"" name=""accumlimit-" & row.Item("ProductGroupID") & """ id=""accumlimit-" & row.Item("ProductGroupID") & """ value=""" & AccumLimit & """ /> " & StrConv(Copient.PhraseLib.Lookup("term.per", LanguageID), VbStrConv.Lowercase) & "</td>")
        Send("            <td><input type=""text"" class=""shorter"" maxlength=""9"" name=""accumperiod-" & row.Item("ProductGroupID") & """ id=""accumperiod-" & row.Item("ProductGroupID") & """ value=""" & AccumPeriod & """ /> " & StrConv(Copient.PhraseLib.Lookup("term.days", LanguageID), VbStrConv.Lowercase) & "</td>")
        Send("        </tr>")
        Send("    </tbody>")
        Send("</table>")

        ProductComboID = 0
        ShowAccumMsg = (AccumMin > 0)
      End If
    End If
    
    'If Disqualifier = 0 Then
    '  ' Return the radios for single and/or
    '  If ProductComboID = 0 Then
    '    Send("<input type=""radio"" id=""ProductComboID0"" name=""ProductComboID"" value=""0"" checked=""checked"" /><label for=""ProductComboID0"">" & Copient.PhraseLib.Lookup("term.single", LanguageID) & "</label>")
    '    Send("<input type=""radio"" id=""ProductComboID1"" name=""ProductComboID"" value=""1"" disabled=""disabled"" /><label for=""ProductComboID1"">" & Copient.PhraseLib.Lookup("term.and", LanguageID) & "</label>")
    '    Send("<input type=""radio"" id=""ProductComboID2"" name=""ProductComboID"" value=""2"" disabled=""disabled"" /><label for=""ProductComboID2"">" & Copient.PhraseLib.Lookup("term.or", LanguageID) & "</label><br />")
    '  Else
    '    MyCommon.QueryStr = "select ProductComboID from CPE_RewardOptions with (NoLock) where RewardOptionID=" & roid
    '    rst = MyCommon.LRT_Select
    '    If (rst.Rows.Count > 0) Then
    '      ProductComboID = rst.Rows(0).Item("PRoductComboID")
    '      If (ProductComboID = 0) Then ProductComboID = 1
    '      ' Set the current selection
    '      Send("<input type=""radio"" id=""ProductComboID0"" name=""ProductComboID"" value=""0""")
    '      Send(" disabled=""disabled"" /><label for=""ProductComboID0"">" & Copient.PhraseLib.Lookup("term.single", LanguageID) & "</label>")
    '      Send("<input type=""radio"" id=""ProductComboID1"" name=""ProductComboID"" value=""1""")
    '      If (ProductComboID = 1) Then Send(" checked=""checked""")
    '      Send(" /><label for=""ProductComboID1"">" & Copient.PhraseLib.Lookup("term.and", LanguageID) & "</label>")
    '      Send("<input type=""radio"" id=""ProductComboID2"" name=""ProductComboID"" value=""2""")
    '      If (ProductComboID = 2) Then Send(" checked=""checked""")
    '      Send(" /><label for=""ProductComboID2"">" & Copient.PhraseLib.Lookup("term.or", LanguageID) & "</label><br />")
    '    End If
    '  End If

    'End If
    
    ' Alert the user 
    If (ShowAccumMsg) Then
      Send("<br /><div class=""red"">* " & Copient.PhraseLib.Lookup("CPE-con-accumulationenabled", LanguageID) & "</div>")
    End If
    
    MyCommon.Close_LogixRT()
  End Sub
  
  Sub GenerateCPETenderConditionValues(ByVal TenderIDs As String, ByVal ROID As String, ByVal ExcTender As String, ByVal ExcVal As String)
    Dim bParsed As Boolean
    Dim dt As DataTable
    Dim row As DataRow
    Dim i As Integer = 0
    Dim ExcludedTender As Boolean = False
    Dim ExcludedValue As Decimal = 0D
    
    If (Request.Form("LanguageID") <> "") Then
      bParsed = Integer.TryParse(Request.Form("lang"), LanguageID)
      If (Not bParsed) Then LanguageID = 1
    End If
    
    MyCommon.Open_LogixRT()

        MyCommon.QueryStr = "select ExcludedTender, ExcludedTenderAmtRequired from CPE_RewardOptions with (NoLock) where RewardOptionID=@ROID and TouchResponse=0 and Deleted=0;"
        MyCommon.DBParameters.Add("@ROID", SqlDbType.BigInt).Value = ROID
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
    
    If (dt.Rows.Count > 0) Then
      ExcludedTender = MyCommon.NZ(dt.Rows(0).Item("ExcludedTender"), False)
      ExcludedValue = MyCommon.NZ(dt.Rows(0).Item("ExcludedTenderAmtRequired"), 0D)
    End If
    
    If (ExcTender <> "") Then ExcludedTender = (ExcTender = "1")
    If (ExcVal <> "") Then Decimal.TryParse(ExcVal, ExcludedValue)
    
        MyCommon.QueryStr = "select ITT.IncentiveTenderID, TEN.TenderTypeID, TEN.Name, ITT.Value from " & _
                            "  (select TT.TenderTypeID, TT.Name from CPE_TenderTypes TT with (NoLock) " & _
                            "   where deleted=0 and TenderTypeID in (SELECT items FROM Split (@TenderTypeIDs, ','))) TEN " & _
                            "left join CPE_IncentiveTenderTypes ITT with (NoLock) on ITT.TenderTypeID = TEN.TenderTypeID and ITT.Deleted=0 and ITT.RewardOptionID=@ROID;"
        MyCommon.DBParameters.Add("@TenderTypeIDs", SqlDbType.NVarChar).Value = TenderIDs
        MyCommon.DBParameters.Add("@ROID", SqlDbType.BigInt).Value = ROID
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)

    If (dt.Rows.Count > 0) Then
      Send("<table summary=""" & Copient.PhraseLib.Lookup("term.values", LanguageID) & """>")
      Send("    <thead>")
      Send("        <tr>")
      Send("            <th class=""th-name"" scope=""col"">" & Copient.PhraseLib.Lookup("term.tender", LanguageID) & "</th>")
      Send("            <th class=""th-value"" style=""" & IIf(ExcludedTender, "display:none;", "") & " scope=""col"">" & Copient.PhraseLib.Lookup("term.value", LanguageID) & "</th>")
      Send("        </tr>")
      Send("    </thead>")
      Send("    <tbody>")
      For Each row In dt.Rows
        Send("     <tr>")
        Send("        <td>" & MyCommon.NZ(row.Item("Name"), "") & "</td>")
        Send("        <td " & IIf(ExcludedTender, "style=""display:none;""", "") & ">")
        Send("          <input type=""text"" class=""shorter"" maxlength=""12"" name=""tVal"" id=""tVal" & i & """ value=""" & MyCommon.NZ(row.Item("Value"), "0") & """ />")
        Send("          <input type=""hidden"" name=""tID"" id=""tID-" & MyCommon.NZ(row.Item("IncentiveTenderID"), "0") & """ value=""" & MyCommon.NZ(row.Item("IncentiveTenderID"), "0") & """ />")
        Send("          <input type=""hidden"" name=""ttID"" id=""ttID-" & MyCommon.NZ(row.Item("TenderTypeID"), "0") & """ value=""" & MyCommon.NZ(row.Item("TenderTypeID"), "0") & """ />")
        Send("        </td>")
        Send("     </tr>")
        i += 1
      Next
      Send("    <tbody>")
      Send("</table>")

    End If

    Send("<br class=""half"" />")
    Send("<div id=""exclusion"">")
    Send("  <input type=""checkbox"" id=""useasexcluded"" name=""useasexcluded""" & IIf(ExcludedTender, " checked=""checked""", "") & " onclick=""handleExcludedClick();""/>")
    Send("  <label for=""useasexcluded"">Use as excluded</label>")
    Send("  <div id=""alltenders"" style=""" & IIf(Not ExcludedTender, "display:none;", "") & """>")
    Send("    <br class=""half"" />")
    Send("    <table summary="""">")
    Send("      <thead>")
    Send("        <tr>")
    Send("          <th class=""th-name"" scope=""col"">" & Copient.PhraseLib.Lookup("term.tender", LanguageID) & "</th>")
    Send("          <th scope=""col"" style=""width:70px;"">" & Copient.PhraseLib.Lookup("term.value", LanguageID) & "</th>")
    Send("        </tr>")
    Send("      </thead>")
    Send("      <tbody>")
    Send("        <tr>")
    Send("          <td><label for=""exVal"">Other tender amount required</label></td>")
    Send("          <td><input type=""text"" class=""shorter"" maxlength=""12"" name=""exVal"" id=""exVal"" value=""" & ExcludedValue & """ /></td>")
    Send("        </tr>")
    Send("      </tbody>")
    Send("    </table>")
    Send("  </div>")
    Send("</div>")
    
    MyCommon.Close_LogixRT()
  End Sub

  '----------------------------------------------------------------------------------
  
  Sub OfferRedemptions(ByVal CustomerPK As String, ByVal CardPK As String, ByVal OfferID As String)
    Dim rst As DataTable
    Dim row As DataRow
    Dim rst2 As DataTable
    Dim dtHH As DataTable
    Dim LanguageID As Integer = 1
    Dim bParsed As Boolean
    Dim TotalRedeemCt As Integer = 0
    Dim TotalRedeemAmt As Double = 0.0
    Dim OfferDesc As String = ""
    Dim ExtCardID As String = ""
    Dim HHID As String = ""
    Dim PreviousHHID As String = ""
    Dim CurrentHHID As String = ""
    Dim HHRedeemCt As Integer = 0
    Dim HHRedeemAmt As Double = 0.0
    Dim HHTransactions As Integer = 0
    Dim j As Integer = 0
    Dim sCmMemberList As String = ""
    Dim iCmAutoHouseholdCustGrpOptionId As Integer = 24
    Dim bCmAutoHouseholdCustGrpEnabled As Boolean = False
    Dim IsHousehold As Boolean = False
    Dim CustExtIdList As String = ""
    Dim CustPKs As String = ""
    Dim LogixTransNums As String = ""
        Dim ExtCardIDOriginal As String = ""
    Response.Cache.SetCacheability(HttpCacheability.NoCache)
    
    If (Request.QueryString("Lang") <> "") Then
      bParsed = Integer.TryParse(Request.QueryString("lang"), LanguageID)
      If (Not bParsed) Then LanguageID = 1
    End If
    
    MyCommon.Open_LogixWH()
    MyCommon.Open_LogixRT()
    MyCommon.Open_LogixXS()
    
    If CardPK <> "" Then
            MyCommon.QueryStr = "select ExtCardID,ExtCardIDOriginal from CardIDs with (NoLock) where CardPK=@CardPK;"
            MyCommon.DBParameters.Add("@CardPK", SqlDbType.Int).Value = CardPK
            rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)
        
            If rst.Rows.Count > 0 Then
                ' No decrypt required as they are used in SQL Inline query.
                ExtCardID = MyCommon.NZ(rst.Rows(0).Item("ExtCardID"), "")
                ExtCardIDOriginal = MyCommon.NZ(rst.Rows(0).Item("ExtCardIDOriginal"), "")
            End If
    End If
    If CustomerPK <> "" Then
      MyCommon.QueryStr = "select C2.InitialCardID as CurrentHHID " & _
                          "from Customers as C1 " & _
                          "inner join Customers C2 on C2.CustomerPK=C1.HHPK " & _
                                "where C1.CustomerPK=@CustomerPK;"
            MyCommon.DBParameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
            dtHH = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)
      
      If dtHH.Rows.Count > 0 Then
        CurrentHHID = MyCryptLib.SQL_StringDecrypt(dtHH.Rows(0).Item("CurrentHHID").ToString())
      End If
      MyCommon.QueryStr = "select C2.InitialCardID as PreviousHHID " & _
                          "from Customers as C1 " & _
                          "inner join Customers C2 on C2.CustomerPK=C1.PreviousHHPK " & _
                                "where C1.CustomerPK=@CustomerPK;"
            MyCommon.DBParameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
            dtHH = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)
            If dtHH.Rows.Count > 0 Then
                PreviousHHID = MyCryptLib.SQL_StringDecrypt(dtHH.Rows(0).Item("PreviousHHID").ToString())
            End If
            MyCommon.QueryStr = "select CustomerTypeID from Customers with (NoLock) where CustomerPK=@CustomerPK;"
            MyCommon.DBParameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
            rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)
      If rst.Rows.Count > 0 Then
        If MyCommon.NZ(rst.Rows(0).Item("CustomerTypeID"), 0) = 1 Then
          IsHousehold = True
        End If
      End If
    End If
    
    MyCommon.QueryStr = "select EngineID from OfferIDs with (NoLock) where OfferID=@OfferID;"
        MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
        rst2 = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
    If (MyCommon.NZ(rst2.Rows(0).Item("EngineID"), 0) <> 2) Then
      ' CM offer (0), Catalina offer (1) or Web offer (3)
      If (MyCommon.NZ(rst2.Rows(0).Item("EngineID"), 0) = 0) Then
        ' CM offer
        bCmAutoHouseholdCustGrpEnabled = MyCommon.Fetch_CM_SystemOption(iCmAutoHouseholdCustGrpOptionId)
        If bCmAutoHouseholdCustGrpEnabled Then
          ' see if this is a household
          Dim lCustomerPK As Long
          Dim dt As DataTable
          MyCommon.QueryStr = "select CustomerPK from CardIDs with (NoLock) where ExtCardID = '" & ExtCardID & "' and CardTypeID=1;"
          rst2 = MyCommon.LXS_Select
          If rst2.Rows.Count > 0 Then
            lCustomerPK = MyCommon.NZ(rst2.Rows(0).Item(0), 0)
            If lCustomerPK > 0 Then
              ' get list of all members of household
              MyCommon.QueryStr = "select CustomerPK from Customers with (NoLock) where HHPK=" & lCustomerPK & ";"
              rst2 = MyCommon.LXS_Select
              If rst2.Rows.Count > 0 Then
                sCmMemberList = "'" & ExtCardID & "'"
                For Each row In rst2.Rows
                  lCustomerPK = MyCommon.NZ(row.Item(0), 0)
                  MyCommon.QueryStr = "select ExtCardID from CardIDs with (NoLock) where CustomerPK=" & lCustomerPK & " and CardTypeID=0;"
                  dt = MyCommon.LXS_Select
                                    If dt.Rows.Count > 0 Then
                                        'no need to decrypt as ExtCardId is used in SQL Inline
                                        sCmMemberList += ",'" & MyCommon.NZ(dt.Rows(0).Item(0), "") & "'"
                                    End If
                Next
              Else
                bCmAutoHouseholdCustGrpEnabled = False
              End If
            Else
              bCmAutoHouseholdCustGrpEnabled = False
            End If
          Else
            bCmAutoHouseholdCustGrpEnabled = False
          End If
        End If
      End If
      MyCommon.QueryStr = "select Description from Offers with (NoLock) where OfferID=@OfferID;"
            MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
            rst2 = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
    Else
      ' CPE offer (2)
            MyCommon.QueryStr = "select Description from CPE_Incentives with (NoLock) where IncentiveID=@OfferID;"
            MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
            rst2 = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
    End If
    If rst2.Rows.Count > 0 Then OfferDesc = MyCommon.NZ(rst2.Rows(0).Item("Description"), "")
    
    Send("<br class=""half"" />")
    Send("<center>")
    Sendb("<span style=""font-size:14px;color:brown;""><b>" & Copient.PhraseLib.Lookup("CPE_accum-adj-redemption-history", LanguageID) & OfferID)
    If IsHousehold Then
      Sendb("&nbsp;" & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & "&nbsp;" & Copient.PhraseLib.Lookup("term.household", LanguageID))
    Else
      Sendb("&nbsp;" & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & "&nbsp;" & Copient.PhraseLib.Lookup("term.customer", LanguageID))
    End If
        Send("&nbsp;" & MyCryptLib.SQL_StringDecrypt(ExtCardIDOriginal) & "</b></span>")
    Send("<br />")
    Send(OfferDesc & "<br />")
    Send("<br class=""half"" />")
    Send("<table style=""width:95%;"" summary=""" & Copient.PhraseLib.Lookup("term.list", LanguageID) & """>")
    Send("    <thead>")
    Send("    <tr>")
    Send("        <th scope=""col"" style=""width:30px;"">&nbsp;</th>")
    Send("        <th scope=""col"">" & Copient.PhraseLib.Lookup("term.date", LanguageID) & "</th>")
    Send("        <th scope=""col"">" & Copient.PhraseLib.Lookup("term.store", LanguageID) & "</th>")
    Send("        <th scope=""col"" style=""text-align:center;"">" & Copient.PhraseLib.Lookup("term.terminal", LanguageID) & "</th>")
    Send("        <th scope=""col"" style=""text-align:center;"">" & Copient.PhraseLib.Lookup("term.transaction", LanguageID) & "</th>")
    Send("        <th scope=""col"" style=""text-align:center;width:10px;"" title=""" & Copient.PhraseLib.Lookup("term.replayed", LanguageID) & """>R</th>")
    Send("        <th scope=""col"" style=""text-align:right;"">" & Copient.PhraseLib.Lookup("term.amount", LanguageID) & "</th>")
    Send("        <th scope=""col"" style=""text-align:right;""></th>")
    Send("        <th scope=""col"" style=""text-align:right;"">" & Copient.PhraseLib.Lookup("term.redemptions", LanguageID) & "</th>")
    Send("    </tr>")
    Send("    </thead>")
    Send("    <tbody>")
    
    Dim ShowPOSTimeStamp As Boolean = MyCommon.Fetch_CPE_SystemOption(131) = "1"
    If bCmAutoHouseholdCustGrpEnabled Then
      MyCommon.QueryStr = "select ExtLocationCode, RedemptionCount, RedemptionAmount, TransDate, TerminalNum, TransNum, " & _
                          "LogixTransNum, PresentedCustomerID, CustomerPrimaryExtID, PresentedCardTypeID, HHID, CustomerTypeID, Replayed, 0 AS TransContext, " & _
                          "SVAmount, SVProgramID, PointsAmount, PointsProgramID, POSTimeStamp " & _
                          "from TransRedemptionView with (NoLock) " & _
                          "where OfferID=" & Convert.ToInt32(OfferID) & " and CustomerPrimaryExtID in (" & sCmMemberList & ") " & _
                          "order by TransDate desc;"
    Else
      Dim OrderBy As String = If(ShowPOSTimeStamp, "POSTimeStamp", "TransDate")
      If IsHousehold Then
        'Get constituents' CustomerPKs
        MyCommon.QueryStr = "dbo.pt_Get_CustPKList"
        MyCommon.Open_LXSsp()
        MyCommon.LXSsp.Parameters.Add("@HHPK", SqlDbType.Int).Value = CustomerPK
        MyCommon.LXSsp.Parameters.Add("@CustPKs", SqlDbType.NVarChar, 4000).Direction = ParameterDirection.Output
        MyCommon.LXSsp.ExecuteNonQuery()
        CustPKs = MyCommon.LXSsp.Parameters("@CustPKs").Value
        MyCommon.Close_LXSsp()
        'Get constituents' ExtCardIDs
        MyCommon.QueryStr = "dbo.pt_Get_ExtCardIDList"
        MyCommon.Open_LXSsp()
        MyCommon.LXSsp.Parameters.Add("@CustPKs", SqlDbType.NVarChar, 4000).Value = CustPKs
        MyCommon.LXSsp.Parameters.Add("@ExtCardIDs", SqlDbType.NVarChar, 4000).Direction = ParameterDirection.Output
                MyCommon.LXSsp.ExecuteNonQuery()
                'No encryption in LogixWH
                Dim eCustExtIdList As String = ""
                Dim cntExtIds As Integer = 0
                eCustExtIdList = MyCommon.LXSsp.Parameters("@ExtCardIDs").Value
                For Each scustExtId In eCustExtIdList.Split(",")
                    If cntExtIds = 0 Then
                        CustExtIdList += "'" & MyCryptLib.SQL_StringDecrypt(scustExtId.Replace("'", "")) & "'"
                    Else
                        CustExtIdList += ",'" & MyCryptLib.SQL_StringDecrypt(scustExtId.Replace("'", "")) & "'"
                    End If
                    cntExtIds = cntExtIds + 1
                Next
                MyCommon.Close_LXSsp()
                'Get constituents' LogixTransNums
                MyCommon.QueryStr = "dbo.pt_Get_LogixTransNumList"
                MyCommon.Open_LXSsp()
                MyCommon.LXSsp.Parameters.Add("@HHPK", SqlDbType.Int).Value = CustomerPK
                MyCommon.LXSsp.Parameters.Add("@ExtCardIDs", SqlDbType.NVarChar).Value = eCustExtIdList
                MyCommon.LXSsp.Parameters.Add("@LogixTransNums", SqlDbType.NVarChar, 4000).Direction = ParameterDirection.Output
                MyCommon.LXSsp.ExecuteNonQuery()
                LogixTransNums = MyCommon.LXSsp.Parameters("@LogixTransNums").Value
                MyCommon.Close_LXSsp()
                'Run the main query
                'First select (TransContext 0)  = Records where the individuals were previously and are currently members of this household.
                'Second select (TransContext 1) = Records where the individuals were previously members of this household, but no longer are. (CustomerID doesn't match a current member of the household, but HouseholdID does match the current household.)
                'Third select (TransContext 2)  = Records where the individuals were previously members of a different household, but now are in the this one. (CustomerID matches a current member of the household, but HouseholdID doesn't match the current household.)
                'Fourth select (TransContext 0) = Records associated to the household itself.
                MyCommon.QueryStr = "select CustomerPrimaryExtID, TransDate, ExtLocationCode, RedemptionAmount, RedemptionCount, " & _
                                    "TerminalNum, TransNum, LogixTransNum, PresentedCustomerID, PresentedCardTypeID, HHID, CustomerTypeID, Replayed, 0 AS TransContext, " & _
                                    "SVAmount, SVProgramID, PointsAmount, PointsProgramID, POSTimeStamp " & _
                                    "from TransRedemptionView with (NoLock) " & _
                                    "where HHID = '" & MyCryptLib.SQL_StringDecrypt(ExtCardID) & "' AND CustomerPrimaryExtID IN (" & CustExtIdList & ") and OfferID=" & Convert.ToInt32(OfferID) & " " & _
                                    " UNION " & _
                                    "select CustomerPrimaryExtID, TransDate, ExtLocationCode, RedemptionAmount, RedemptionCount, " & _
                                    "TerminalNum, TransNum, LogixTransNum, PresentedCustomerID, PresentedCardTypeID, HHID, CustomerTypeID, Replayed, 1 AS TransContext, " & _
                                    "SVAmount, SVProgramID, PointsAmount, PointsProgramID, POSTimeStamp " & _
                                    "from TransRedemptionView with (NoLock) " & _
                                    "where HHID = '" & MyCryptLib.SQL_StringDecrypt(ExtCardID) & "' AND CustomerPrimaryExtID NOT IN (" & CustExtIdList & ") and OfferID=" & Convert.ToInt32(OfferID) & " " & _
                                    " UNION " & _
                                    "select CustomerPrimaryExtID, TransDate, ExtLocationCode, RedemptionAmount, RedemptionCount, " & _
                                    "TerminalNum, TransNum, LogixTransNum, PresentedCustomerID, PresentedCardTypeID, HHID, CustomerTypeID, Replayed, 2 AS TransContext, " & _
                                    "SVAmount, SVProgramID, PointsAmount, PointsProgramID, POSTimeStamp " & _
                                    "from TransRedemptionView with (NoLock) " & _
                                    "where ISNULL(HHID, '') <> '" & MyCryptLib.SQL_StringDecrypt(ExtCardID) & "' AND (CustomerPrimaryExtID IN (" & CustExtIdList & ") and CustomerTypeID<>1) AND OfferID=" & Convert.ToInt32(OfferID) & " " & _
                                    " UNION " & _
                                    "select CustomerPrimaryExtID, TransDate, ExtLocationCode, RedemptionAmount, RedemptionCount, " & _
                                    "TerminalNum, TransNum, LogixTransNum, PresentedCustomerID, PresentedCardTypeID, HHID, CustomerTypeID, Replayed, 0 AS TransContext, " & _
                                    "SVAmount, SVProgramID, PointsAmount, PointsProgramID, POSTimeStamp " & _
                                    "from TransRedemptionView with (NoLock) " & _
                                    "where ISNULL(HHID, '') <> '" & MyCryptLib.SQL_StringDecrypt(ExtCardID) & "' AND (CustomerPrimaryExtID='" & MyCryptLib.SQL_StringDecrypt(ExtCardID) & "' AND CustomerTypeID=1) AND OfferID=" &  Convert.ToInt32(OfferID) & " " & _
                                    "order by " & OrderBy & " desc;"
            Else
                MyCommon.QueryStr = "select ExtLocationCode, RedemptionCount, RedemptionAmount, TransDate, TerminalNum, TransNum, LogixTransNum, " & _
                                    "PresentedCustomerID, CustomerPrimaryExtID, PresentedCardTypeID, HHID, CustomerTypeID, Replayed, 0 as TransContext, " & _
                                    "SVAmount, SVProgramID, PointsAmount, PointsProgramID, POSTimeStamp " & _
                                    "from TransRedemptionView with (NoLock) " & _
                                    "where OfferID=" & Convert.ToInt32(OfferID) & " and CustomerPrimaryExtID='" & MyCryptLib.SQL_StringDecrypt(ExtCardID) & "' " & _
                                    "order by " & OrderBy & " desc;"
            End If
    End If

    rst = MyCommon.LWH_Select
    If (rst.Rows.Count > 0) Then
      j = 0
      For Each row In rst.Rows
        j += 1
        HHID = row.Item("HHID").ToString()
        If (HHID <> "") AndAlso (HHID = PreviousHHID) AndAlso (HHID <> CurrentHHID) Then
          Send("    <tr id=""hist" & j & """ style=""color:RED;"">")
          HHRedeemAmt += MyCommon.NZ(row.Item("RedemptionAmount"), 0.0)
          HHRedeemCt += MyCommon.NZ(row.Item("RedemptionCount"), 0)
          HHTransactions += 1
        ElseIf (HHID <> "") AndAlso (MyCommon.NZ(row.Item("TransContext"), 0) = 1) AndAlso IsHousehold Then
          Send("    <tr id=""hist" & j & """ style=""color:RED;"">")
          HHRedeemAmt += MyCommon.NZ(row.Item("RedemptionAmount"), 0.0)
          HHRedeemCt += MyCommon.NZ(row.Item("RedemptionCount"), 0)
          HHTransactions += 1
        ElseIf (HHID <> "") AndAlso ((MyCommon.NZ(row.Item("TransContext"), 0) = 2) OrElse (MyCommon.NZ(row.Item("TransContext"), 0) = 3)) AndAlso IsHousehold Then
          Send("    <tr id=""hist" & j & """ style=""color:GREEN;"">")
          HHRedeemAmt += MyCommon.NZ(row.Item("RedemptionAmount"), 0.0)
          HHRedeemCt += MyCommon.NZ(row.Item("RedemptionCount"), 0)
          HHTransactions += 1
        Else
          Send("    <tr id=""hist" & j & """>")
        End If
        Send("        <td><input class=""ex more"" type=""button"" value=""+"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.details", LanguageID) & """ onclick=""javascript:showDetail(" & j & ", this);"" /></td>")
        If ShowPOSTimeStamp Then
          Send("        <td>" & MyCommon.NZ(row.Item("POSTimeStamp"), "") & "</td>")
        Else
          Send("        <td>" & MyCommon.NZ(row.Item("TransDate"), "") & "</td>")
        End If
        Send("        <td>" & MyCommon.NZ(row.Item("ExtLocationCode"), "") & "</td>")
        Send("        <td style=""text-align:center;"">" & MyCommon.NZ(row.Item("TerminalNum"), "") & "</td>")
        Send("        <td style=""text-align:center;word-break:break-all""><span title=""" & MyCommon.NZ(row.Item("LogixTransNum"), "") & """>" & MyCommon.NZ(row.Item("TransNum"), "") & "</span></td>")
        Send("        <td style=""text-align:center;"">" & IIf(MyCommon.NZ(row.Item("Replayed"), 0) > 0, "<span title=""" & Copient.PhraseLib.Lookup("term.replayed", LanguageID) & """ style=""color:#dd0000;cursor:default;font-size:10px;font-weight:bold;"">R</span>", "") & "</td>")
        
        Dim DisplayAmount As String = ""
        Dim DisplayType As String = ""
        If (MyCommon.NZ(row.Item("PointsProgramID"), 0) > 0) Then
          DisplayAmount = MyCommon.NZ(row.Item("PointsAmount"), "")
          DisplayType = "<span title=""" & Copient.PhraseLib.Lookup("term.points", LanguageID) & """ style=""cursor:default;font-size:12px;font-weight:bold;"">P</span>"
        ElseIf (MyCommon.NZ(row.Item("SVProgramID"), 0) > 0) Then
          DisplayAmount = MyCommon.NZ(row.Item("SVAmount"), "")
          DisplayType = "<span title=""" & Copient.PhraseLib.Lookup("term.storedvalue", LanguageID) & """ style=""cursor:default;font-size:12px;font-weight:bold;"">S</span>"
        Else
          DisplayAmount = "$" & MyCommon.NZ(row.Item("RedemptionAmount"), "")
        End If
                
        Send("        <td style=""text-align:right;"">" & DisplayAmount & "</td>")
        Send("        <td style=""text-align:right;"">" & DisplayType & "</td>")
        Send("        <td style=""text-align:right;"">" & MyCommon.NZ(row.Item("RedemptionCount"), "") & "</td>")
        Send("    </tr>")
        Send("    <tr id=""histdetail" & j & """ style=""display:none;color:#777777;"">")
        Send("        <td></td>")
        Send("        <td colspan=""6"">")
        Send("          " & Copient.PhraseLib.Lookup("term.presented", LanguageID) & "&nbsp;" & Copient.PhraseLib.Lookup("term.id", LanguageID) & ":&nbsp;" & IIf(IsDBNull(row.Item("PresentedCustomerID")), "Unknown", row.Item("PresentedCustomerID").ToString()) & "<br />")
                Send("          " & Copient.PhraseLib.Lookup("term.resolved", LanguageID) & "&nbsp;" & Copient.PhraseLib.Lookup("term.id", LanguageID) & ":&nbsp;" & IIf(IsDBNull(row.Item("CustomerPrimaryExtID")), "Unknown",row.Item("CustomerPrimaryExtID").ToString()) & "<br />")
        Sendb("          " & Copient.PhraseLib.Lookup("term.household", LanguageID) & "&nbsp;" & Copient.PhraseLib.Lookup("term.id", LanguageID) & ":&nbsp;")
                If (row.Item("HHID").ToString() = "") AndAlso (MyCommon.NZ(row.Item("CustomerTypeID"), 0) = 1) Then
                    Send(IIf(IsDBNull(row.Item("CustomerPrimaryExtID")), "Unknown", row.Item("CustomerPrimaryExtID").ToString()) & "<br />")
                Else
                    Send(IIf(IsDBNull(row.Item("HHID")), "Unknown",row.Item("HHID").ToString()) & "<br />")
                End If
        
        Dim ProgramDT As DataTable
        If (MyCommon.NZ(row.Item("PointsProgramID"), 0) > 0) Then
          MyCommon.QueryStr = "Select ProgramName from PointsPrograms with (NoLock) where ProgramID=" & MyCommon.NZ(row.Item("PointsProgramID"), 0) & ";"
          ProgramDT = MyCommon.LRT_Select()
          If ProgramDT.Rows.Count > 0 Then
            Send("      " & Copient.PhraseLib.Lookup("term.pointsprogram", LanguageID) & " #" & MyCommon.NZ(row.Item("PointsProgramID"), 0) & ": " & MyCommon.NZ(ProgramDT.Rows(0).Item("ProgramName"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)))
          End If
        ElseIf (MyCommon.NZ(row.Item("SVProgramID"), 0) > 0) Then
          MyCommon.QueryStr = "Select Name from StoredValuePrograms with (NoLock) where SVProgramID=" & MyCommon.NZ(row.Item("SVProgramID"), 0) & ";"
          ProgramDT = MyCommon.LRT_Select()
          If ProgramDT.Rows.Count > 0 Then
            Send("      " & Copient.PhraseLib.Lookup("term.storedvalueprogram", LanguageID) & " #" & MyCommon.NZ(row.Item("SVProgramID"), 0) & ":  " & MyCommon.NZ(ProgramDT.Rows(0).Item("Name"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)))
          End If
        End If
        
        Send("        </td>")
        Send("    </tr>")
        TotalRedeemAmt += MyCommon.NZ(row.Item("RedemptionAmount"), 0.0)
        TotalRedeemCt += MyCommon.NZ(row.Item("RedemptionCount"), 0)
      Next
      Send("    <tr style=""height:15px;"">")
      Send("        <td colspan=""6""></td>")
      Send("        <td><hr></td>")
      Send("        <td><hr></td>")
      Send("        <td><hr></td>")
      Send("    </tr>")
      Send("    <tr>")
      Send("        <td colspan=""2"">" & Copient.PhraseLib.Lookup("term.total", LanguageID) & " " & Copient.PhraseLib.Lookup("CPE_accum-adj-transamt", LanguageID) & ": " & rst.Rows.Count & "</td>")
      Send("        <td colspan=""4""></td>")
      Send("        <td style=""text-align:right;"">" & Format(TotalRedeemAmt, "$ #,###,##0.000") & "</td>")
      Send("        <td></td>")
      Send("        <td style=""text-align:right;"">" & TotalRedeemCt & "</td>")
      Send("    </tr>")
      If (HHTransactions > 0) AndAlso (Not IsHousehold) Then
        Send("    <tr style=""color:RED;"">")
        Send("        <td colspan=""2"">" & Copient.PhraseLib.Lookup("term.previous", LanguageID) & " <acronym title=""" & Copient.PhraseLib.Lookup("term.household", LanguageID) & """>HH</acronym> " & Copient.PhraseLib.Lookup("CPE_accum-adj-transamt", LanguageID) & ": " & HHTransactions & "</td>")
        Send("        <td colspan=""4""></td>")
        Send("        <td style=""text-align:right;"">" & Format(HHRedeemAmt, "$ #,###,##0.000") & "</td>")
        Send("        <td></td>")
        Send("        <td style=""text-align:right;"">" & HHRedeemCt & "</td>")
        Send("    </tr>")
      End If
    Else
      Send("    <tr>")
      Send("        <td colspan=""8"" style=""text-align:center""><i>" & Copient.PhraseLib.Lookup("CPE_accum-adj-notranshistory", LanguageID) & "</i></td>")
      Send("    </tr>")
    End If
    Send("</tbody>")
    Send("</table>")
    Send("<br /><br />")
    Send("<small>" & Logix.ToShortDateTimeString(DateTime.Now, MyCommon) & "</small>")
    Send("</center>")
    
    MyCommon.Close_LogixWH()
    MyCommon.Close_LogixRT()
    MyCommon.Close_LogixXS()
    
    'HandleAccumLink(ExtCardID, OfferID)
  End Sub
  
  '----------------------------------------------------------------------------------
  
  Sub CAMOfferTransactions(ByVal CustomerPK As Long, ByVal OfferID As Long)
    Dim rst As DataTable
    Dim row As DataRow
    Dim LanguageID As Integer = 1
    Dim bParsed As Boolean
    Dim TotalRedeemCt As Integer = 0
    Dim TotalRedeemAmt As Double = 0.0
    Dim OfferDesc As String = ""
    Dim OfferName As String = ""
    Dim CustomerExtID As String = ""
    Dim CustName As String = ""
    
    Response.Cache.SetCacheability(HttpCacheability.NoCache)
    
    Try
      MyCommon.Open_LogixWH()
      MyCommon.Open_LogixRT()
      MyCommon.Open_LogixXS()
      
      If (Request.QueryString("Lang") <> "") Then
        bParsed = Integer.TryParse(Request.QueryString("lang"), LanguageID)
        If (Not bParsed) Then LanguageID = 1
      End If
    
      ' get the customers card number from the unique identifier
            MyCommon.QueryStr = "select PrimaryExtID, FirstName, LastName from Customers with (NoLock) where CustomerPK=@CustomerPK"
            MyCommon.DBParameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
            rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)
      If (rst.Rows.Count > 0) Then
        CustomerExtID = MyCommon.NZ(rst.Rows(0).Item("PrimaryExtID"), "")
        CustName = MyCommon.NZ(rst.Rows(0).Item("FirstName"), "") & " " & MyCommon.NZ(rst.Rows(0).Item("LastName"), "")
      End If
    
      ' get the offers description for display purposes
            MyCommon.QueryStr = "select IncentiveName, Description from CPE_Incentives with (NoLock) where IncentiveID = @OfferID"
            MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
            rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
      
      If (rst.Rows.Count > 0) Then
        OfferName = MyCommon.NZ(rst.Rows(0).Item("IncentiveName"), "")
        OfferDesc = MyCommon.NZ(rst.Rows(0).Item("Description"), "")
      End If
    
      Send("<br class=""half"" />")
      Send("<table style=""width:95%;font-size:10pt;color:#333333;border:solid 1px #333333;"" cellpadding=""0"" cellspacing=""0"" summary=""" & Copient.PhraseLib.Lookup("term.customer", LanguageID) & """>")
      Send("    <tr style=""background-color:#dddddd;"">")
      Send("      <td><b>" & Copient.PhraseLib.Lookup("term.customer", LanguageID) & ":</b></td>")
      Send("      <td>" & Copient.PhraseLib.Lookup("term.id", LanguageID) & ": " & CustomerExtID & " </td>")
      Send("      <td>" & Copient.PhraseLib.Lookup("term.name", LanguageID) & ": " & CustName & " </td>")
      Send("    </tr>")
      Send("    <tr style=""background-color:#eeeeee;"">")
      Send("      <td><b>" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & "</b>:</td>")
      Send("      <td>" & Copient.PhraseLib.Lookup("term.id", LanguageID) & ": " & OfferID & " </td>")
      Send("      <td>" & Copient.PhraseLib.Lookup("term.name", LanguageID) & ": " & OfferName & " </td>")
      Send("    </tr>")
      Send("    <tr style=""background-color:#eeeeee;"">")
      Send("      <td></td>")
      Send("      <td colspan=""2"">" & Copient.PhraseLib.Lookup("term.description", LanguageID) & ": " & OfferDesc & "</td>")
      Send("    </tr>")
      Send("</table>")
      Send("<br class=""half"" />")
      Send("<table style=""width:95%;"" summary=""" & Copient.PhraseLib.Lookup("term.list", LanguageID) & """>")
      Send("    <thead>")
      Send("    <tr>")
      Send("        <th scope=""col"">" & Copient.PhraseLib.Lookup("term.adjust", LanguageID) & "</th>")
      Send("        <th scope=""col"">" & Copient.PhraseLib.Lookup("term.date", LanguageID) & "</th>")
      Send("        <th scope=""col"">" & Copient.PhraseLib.Lookup("term.store", LanguageID) & "</th>")
      Send("        <th scope=""col"" style=""text-align:center;"">" & Copient.PhraseLib.Lookup("term.terminal", LanguageID) & "</th>")
      Send("        <th scope=""col"" style=""text-align:center;"">" & Copient.PhraseLib.Lookup("term.transaction", LanguageID) & "</th>")
      Send("        <th scope=""col"" style=""text-align:right;"">" & Copient.PhraseLib.Lookup("term.amount", LanguageID) & "</th>")
      Send("        <th scope=""col"" style=""text-align:right;"">" & Copient.PhraseLib.Lookup("term.redemptions", LanguageID) & "</th>")
      Send("    </tr>")
      Send("    </thead>")
      Send("    <tbody>")
    
            MyCommon.QueryStr = "select ExtLocationCode, RedemptionCount, RedemptionAmount, TransDate, TerminalNum, TransNum " & _
                                "from TransRedemptionView with (NoLock) " & _
                                "where OfferID=@OfferID and CustomerPrimaryExtID=@CustomerPrimaryExtID order by TransDate;"
            MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
            MyCommon.DBParameters.Add("@CustomerPrimaryExtID", SqlDbType.NVarChar).Value = MyCryptLib.SQL_StringDecrypt(CustomerExtID)
            rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixWH)
      If (rst.Rows.Count > 0) Then
        For Each row In rst.Rows
          Send("    <tr>")
          Sendb("       <td><input type=""button"" class=""adjust"" id=""ptsAdj" & MyCommon.NZ(row.Item("OfferID"), "") & """ name=""ptsAdj"" value=""P"" title=""" & Copient.PhraseLib.Lookup("term.adjust", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.points", LanguageID), VbStrConv.Lowercase) & """ ")
          Send("onClick=""javascript:openPopup('CAM-point-adjust.aspx?OfferID=" & MyCommon.NZ(row.Item("OfferID"), "") & "&CustomerPK=" & CustomerPK & "');"" /></td>")
          Send("        <td>" & MyCommon.NZ(row.Item("TransDate"), "") & "</td>")
          Send("        <td>" & MyCommon.NZ(row.Item("ExtLocationCode"), "") & "</td>")
          Send("        <td style=""text-align:center;"">" & MyCommon.NZ(row.Item("TerminalNum"), "") & "</td>")
          Send("        <td style=""text-align:center;word-break:break-all"">" & MyCommon.NZ(row.Item("TransNum"), "") & "</td>")
          Send("        <td style=""text-align:right;"">" & MyCommon.NZ(row.Item("RedemptionAmount"), "") & "</td>")
          Send("        <td style=""text-align:right;"">" & MyCommon.NZ(row.Item("RedemptionCount"), "") & "</td>")
          Send("    </tr>")
          TotalRedeemAmt += MyCommon.NZ(row.Item("RedemptionAmount"), 0.0)
          TotalRedeemCt += MyCommon.NZ(row.Item("RedemptionCount"), 0)
        Next
        Send("    <tr style=""height:15px;"">")
        Send("        <td colspan=""5""></td>")
        Send("        <td><hr></td>")
        Send("        <td><hr></td>")
        Send("    </tr>")
        Send("    <tr>")
        Send("        <td colspan=""3"">" & Copient.PhraseLib.Lookup("CPE_accum-adj-transamt", LanguageID) & ": " & rst.Rows.Count & "</td>")
        Send("        <td colspan=""2""></td>")
        Send("        <td style=""text-align:right;"">" & Format(TotalRedeemAmt, "$ #,###,##0.000") & "</td>")
        Send("        <td style=""text-align:right;"">" & TotalRedeemCt & "</td>")
        Send("    </tr>")
      Else
        Send("    <tr>")
        Send("        <td colspan=""7"" style=""text-align:center""><i>" & Copient.PhraseLib.Lookup("CPE_accum-adj-notranshistory", LanguageID) & "</i></td>")
        Send("    </tr>")
      End If
    
      Send("</tbody>")
      Send("</table>")
      Send("<br /><br />")
      Send("<center>")
      Send("<input type=""button"" id=""btnNewTrans"" name=""btnNewTrans"" value=""" & Copient.PhraseLib.Lookup("term.newtransaction", LanguageID) & """ />")
      Send("<br />")
      Send("<small>" & Logix.ToShortDateTimeString(DateTime.Now, MyCommon) & "</small>")
      Send("</center>")
    Catch ex As Exception
      Send(ex.ToString)
    Finally
      MyCommon.Close_LogixWH()
      MyCommon.Close_LogixRT()
      MyCommon.Close_LogixXS()
    End Try
    
  End Sub
  
  Sub HandleAccumLink(ByVal CustomerExtID As String, ByVal OfferId As Long)
    Dim rst As DataTable
    Dim EngineID As Integer = -1
    Dim AccumProgram As Boolean = False
    Dim bParsed As Boolean = False
    Dim Logix As New Copient.LogixInc
    Dim AdminUserID As Long
    
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
    
    If (Request.QueryString("Lang") <> "") Then
      bParsed = Integer.TryParse(Request.QueryString("lang"), LanguageID)
      If (Not bParsed) Then LanguageID = 1
    End If
    
    
        MyCommon.QueryStr = "select EngineID from OfferIDs with (NoLock) where OfferID =@OfferId"
        MyCommon.DBParameters.Add("@OfferId", SqlDbType.BigInt).Value = OfferId
        rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        If (rst.Rows.Count > 0) Then
            EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), -1)
        End If
    
    ' At this time, only CPE engine offers allow accumulation adjustments
    If (EngineID = 2) Then
      MyCommon.QueryStr = "select IPG.AccumMin " & _
                          "from CPE_IncentiveProductGroups as IPG with (NoLock) Inner Join CPE_RewardOptions as RO with (NoLock) on IPG.RewardOptionID=RO.RewardOptionID and IPG.Deleted=0 and IPG.ExcludedProducts=0 and RO.Deleted=0 " & _
                          "where RO.IncentiveID=@OfferId;"
            MyCommon.DBParameters.Add("@OfferId", SqlDbType.BigInt).Value = OfferId
            rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
      If (rst.Rows.Count > 0) Then
        If MyCommon.NZ(rst.Rows(0).Item("AccumMin"), 0) > 0 Then
          AccumProgram = True
        End If
      End If
      If (AccumProgram AndAlso Logix.UserRoles.EditAccumBalances) Then
        Send("<br />")
        Send("<br />")
        Send("<center>")
        Send("<input type=""button"" class=""large"" value=""" & Copient.PhraseLib.Lookup("CPE_accum-adj", LanguageID) & """ onclick=""javascript:openPopup('CPEaccum-adjust.aspx?OfferID=" & OfferId & "&CustomerExtId=" & CustomerExtID & " ');"" />")
        Send("</center>")
      End If
    End If
  End Sub
  
  '----------------------------------------------------------------------------------
  
  Sub GenerateReport(ByVal ReportID As String)
    Dim builder As StringBuilder = New StringBuilder()
    Dim bParsed As Boolean = False
    Dim LanguageID As Integer = 1
    Dim Frequency As String = 1
    
    Response.Cache.SetCacheability(HttpCacheability.NoCache)
    
    If (Request.Form("lang") <> "") Then
      bParsed = Integer.TryParse(Request.Form("lang"), LanguageID)
      If (Not bParsed) Then LanguageID = 1
    End If
    If (Request.Form("frequency") <> "") Then
      Frequency = Request.Form("frequency")
    End If
    
    builder.Append(CreateReport(LanguageID))
    Response.Write(builder)
  End Sub
  
  '----------------------------------------------------------------------------------
  
  Function CreateReport(ByVal LanguageID As Integer) As String
    Dim ReportStartDate As Date
    Dim ReportEndDate As Date
    Dim ReportWeeks As Integer
    Dim RowCount As Integer
    Dim CumulativeImpress As Double
    Dim CumulativeRedeem As Double
        Dim CumulativeTransactions As Double
        Dim CumulativeAmtRedeem As Double
        Dim RedemptionRate As Double
        Dim AmtRedeem As Double
        Dim Redemptions As Double
        Dim Transactions As Double
        Dim Impressions As Double
        Dim i As Integer
        Dim OfferID As String = ""
        Dim dst As System.Data.DataTable
        Dim bParsed As Boolean
        Dim ExportRequested As String
        Dim builder As StringBuilder = New StringBuilder()
        Dim frequency As String = ""
        Dim DefaultToEnhancedCustomReport As Integer = 0
    
        If MyCommon.Fetch_SystemOption(274) = "1" Then
            DefaultToEnhancedCustomReport = 1
        End If
 
        bParsed = DateTime.TryParse(Request.Form("reportstart"), ReportStartDate)
        If (Not bParsed) Then ReportStartDate = Now()
        bParsed = DateTime.TryParse(Request.Form("reportend"), ReportEndDate)
        If (Not bParsed) Then ReportEndDate = Now
        ReportWeeks = DateTime.Compare(ReportStartDate, ReportEndDate) / 7
        OfferID = Request.Form("offerId")
        ExportRequested = Request.Form("exportRpt")
    
        MyCommon.Open_LogixWH()
        MyCommon.QueryStr = "select NumImpressions, NumRedemptions, NumTransactions, AmountRedeemed, ReportingDate from OfferReporting with (nolock) " & _
                            "where OfferID = @OfferID " & _
                            "and ReportingDate between '" & ReportStartDate.ToString & "' and '" & ReportEndDate.ToString & "' " & _
                            "order by ReportingDate"
        MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
        dst = MyCommon.ExecuteQuery(Copient.DataBases.LogixWH)
    
        If (Request.Form("frequency") = "1") Then
            dst = RollupReportWeek(dst, ReportStartDate, ReportEndDate)
            frequency = "weekly"
        ElseIf (Request.Form("frequency") = "2") Then
            dst = FillInDays(dst, ReportStartDate, ReportEndDate)
            frequency = "daily"
        End If
    
        If (ExportRequested <> "") Then
            Response.AddHeader("Content-Disposition", "attachment; filename=Offer" & OfferID & "_Rpt.csv")
            Response.ContentType = "application/octet-stream"
            Return ExportReport(dst, frequency)
        End If
    
        RowCount = dst.Rows.Count
    
        If (RowCount > 0) Then
            builder.Append("<table id=""reportStats"" class=""raligntable"" cellpadding=""2"" cellspacing=""0"" summary=""" & Copient.PhraseLib.Lookup("term.reports", LanguageID) & """")
            builder.Append(Copient.PhraseLib.Lookup("term.impressions", LanguageID))
            builder.Append(""">")
            builder.Append("<tr style=""height:24px;"">")
            For i = 0 To (RowCount - 1)
                builder.Append("<th scope=""col"">&nbsp;&nbsp;&nbsp;&nbsp;" & dst.Rows(i).Item("ReportingDate") & "</th>")
            Next
            builder.Append("</tr>")
            builder.Append("<tr id=""rowBody1"" ondblclick=""javascript:toggleHighlight(1);"">")
            For i = 0 To (RowCount - 1)
                Impressions = MyCommon.NZ(dst.Rows(i).Item("NumImpressions"), 0)
                If (Impressions = 0) Then
                    builder.Append("<td>-</td>")
                Else
                    builder.Append("<td>" & Impressions & "</td>")
                End If
            Next
            builder.Append("</tr>")
            builder.Append("<tr id=""rowBody2"" class=""shaded"" ondblclick=""javascript:toggleHighlight(2);"">")
            For i = 0 To (RowCount - 1)
                CumulativeImpress = CumulativeImpress + MyCommon.NZ(dst.Rows(i).Item("NumImpressions"), 0)
                If (CumulativeImpress = 0) Then
                    builder.Append("<td>-</td>")
                Else
                    builder.Append("<td>" & CumulativeImpress & "</td>")
                End If
            Next
            builder.Append("</tr>")
            builder.Append("<tr id=""rowBody3"" ondblclick=""javascript:toggleHighlight(3);"">")
            For i = 0 To (RowCount - 1)
                builder.Append("<td>")
                Redemptions = MyCommon.NZ(dst.Rows(i).Item("NumRedemptions"), 0)
                If (Redemptions = 0) Then
                    builder.Append("-")
                Else
                    builder.Append(Redemptions)
                End If
                builder.Append("</td>")
            Next
            builder.Append("</tr>")
            builder.Append("<tr id=""rowBody4"" class=""shaded"" ondblclick=""javascript:toggleHighlight(4);"">")
            For i = 0 To (RowCount - 1)
                CumulativeRedeem = CumulativeRedeem + MyCommon.NZ(dst.Rows(i).Item("NumRedemptions"), 0)
                If (CumulativeRedeem = 0) Then
                    builder.Append("<td>-</td>")
                Else
                    builder.Append("<td>" & CumulativeRedeem & "</td>")
                End If
            Next
            builder.Append("</tr>")
            
            If DefaultToEnhancedCustomReport = 1 Then
                builder.Append("<tr id=""rowBody5"" ondblclick=""javascript:toggleHighlight(5);"">")
                For i = 0 To (RowCount - 1)
                    builder.Append("<td>")
                    Transactions = MyCommon.NZ(dst.Rows(i).Item("NumTransactions"), 0)
                    If (Transactions = 0) Then
                        builder.Append("-")
                    Else
                        builder.Append(Transactions)
                    End If
                    builder.Append("</td>")
                Next
                builder.Append("</tr>")
                builder.Append("<tr id=""rowBody6"" class=""shaded"" ondblclick=""javascript:toggleHighlight(6);"">")
                For i = 0 To (RowCount - 1)
                    CumulativeTransactions = CumulativeTransactions + MyCommon.NZ(dst.Rows(i).Item("NumTransactions"), 0)
                    If (CumulativeTransactions = 0) Then
                        builder.Append("<td>-</td>")
                    Else
                        builder.Append("<td>" & CumulativeTransactions & "</td>")
                    End If
                Next
                builder.Append("</tr>")
            End If
            
            builder.Append("<tr id=""rowBody7"" ondblclick=""javascript:toggleHighlight(7);"">")
            For i = 0 To (RowCount - 1)
                builder.Append("<td>")
                AmtRedeem = MyCommon.NZ(dst.Rows(i).Item("AmountRedeemed"), 0.0)
                If (AmtRedeem = 0) Then
                    builder.Append("-")
                Else
                    builder.Append(AmtRedeem.ToString("$#,##0.00"))
                End If
                builder.Append("</td>")
            Next
            builder.Append("</tr>")
            builder.Append("<tr id=""rowBody8"" class=""shaded"" ondblclick=""javascript:toggleHighlight(8);"">")
            For i = 0 To (RowCount - 1)
                CumulativeAmtRedeem = CumulativeAmtRedeem + MyCommon.NZ(dst.Rows(i).Item("AmountRedeemed"), 0.0)
                If (CumulativeAmtRedeem = 0.0) Then
                    builder.Append("<td>-</td>")
                Else
                    builder.Append("<td>" & CumulativeAmtRedeem.ToString("$#,##0.00") & "</td>")
                End If
            Next
            builder.Append("</tr>")
            builder.Append("<tr id=""rowBody9"" ondblclick=""javascript:toggleHighlight(9);"">")
            For i = 0 To (RowCount - 1)
                Impressions = MyCommon.NZ(dst.Rows(i).Item("NumImpressions"), 0)
                Redemptions = MyCommon.NZ(dst.Rows(i).Item("NumRedemptions"), 0)
                If (Impressions > 0) Then
                    RedemptionRate = Redemptions / Impressions
                Else
                    RedemptionRate = 0.0
                End If
                If (RedemptionRate = 0) Then
                    builder.Append("<td>-</td>")
                Else
                    builder.Append("<td>" & RedemptionRate.ToString("##.##%") & "</td>")
                End If
            Next
            builder.Append("</tr>")
            builder.Append("</table>")
            builder.Append("<input type=""hidden"" id=""colCt"" name=""colCt"" value=""" & RowCount & """ />")
            builder.Append("<br /><br />")
            MyCommon.Close_LogixWH()
        Else
            builder.Append("<b>" & Copient.PhraseLib.Lookup("reports.nodata", LanguageID) & "</b>")
        End If
        Return builder.ToString
    End Function
  
    '----------------------------------------------------------------------------------
  
    Function RollupReportWeek(ByVal dst As DataTable, ByVal ReportStartDate As Date, ByVal ReportEndDate As Date) As DataTable
        Dim dstWeek As New DataTable
        Dim i, j As Integer
        Dim numRedeem As Double
        Dim numTransact As Double
        Dim numImpression As Double
        Dim amtRedeem As Double
    
        If (dst.Rows.Count > 0) Then
            Dim CurrentStart As Date
            Dim CurrentEnd As Date
            Dim ReportWeeks As Integer
            Dim row As DataRow
            Dim rowCt As Integer
      
            dstWeek = dst.Copy()
            dstWeek.Clear()
            CurrentStart = ReportStartDate
            CurrentEnd = ReportStartDate.AddDays(6)
            ReportWeeks = DateDiff(DateInterval.Day, ReportStartDate, ReportEndDate) / 7
      
            For i = 0 To ReportWeeks
                If (DateTime.Compare(ReportEndDate, CurrentStart) >= 0) Then
                    dst.DefaultView.RowFilter = "ReportingDate >= '" & CurrentStart.ToString() & "' and ReportingDate <= '" & CurrentEnd.ToString() & "'"
                    rowCt = dst.DefaultView.Count
                    If (rowCt > 0) Then
                        For j = 0 To rowCt - 1
                            If Not IsDBNull(dst.DefaultView(j).Item("NumRedemptions")) Then numRedeem += dst.DefaultView(j).Item("NumRedemptions")
                            If Not IsDBNull(dst.DefaultView(j).Item("NumTransactions")) Then numTransact += dst.DefaultView(j).Item("NumTransactions")
                            If Not IsDBNull(dst.DefaultView(j).Item("AmountRedeemed")) Then amtRedeem += dst.DefaultView(j).Item("AmountRedeemed")
                            If Not IsDBNull(dst.DefaultView(j).Item("NumImpressions")) Then numImpression += dst.DefaultView(j).Item("NumImpressions")
                            If (j = dst.DefaultView.Count - 1) Then
                                row = dst.DefaultView(j).Row
                                row.Item("ReportingDate") = CurrentStart
                                row.Item("NumRedemptions") = numRedeem
                                row.Item("NumTransactions") = numTransact
                                row.Item("AmountRedeemed") = amtRedeem
                                row.Item("NumImpressions") = numImpression
                                dstWeek.ImportRow(row)
                            End If
                        Next
                    Else
                        row = dstWeek.NewRow()
                        row.Item("ReportingDate") = CurrentStart
                        row.Item("NumRedemptions") = 0
                        row.Item("NumTransactions") = 0
                        row.Item("AmountRedeemed") = 0.0
                        row.Item("NumImpressions") = 0
                        dstWeek.Rows.Add(row)
                    End If
                    numRedeem = 0
                    numTransact = 0
                    amtRedeem = 0.0
                    numImpression = 0
                    CurrentStart = CurrentEnd.AddDays(1)
                    CurrentEnd = CurrentStart.AddDays(6)
                End If
            Next
        End If
    
        Return dstWeek
    End Function
  
    '----------------------------------------------------------------------------------
  
    Function FillInDays(ByVal dst As DataTable, ByVal ReportStartDate As Date, ByVal ReportEndDate As Date) As DataTable
        Dim dstDay As New DataTable
        Dim CurrentDate As Date
        Dim RptDate As Date
        Dim row As DataRow
    
        dstDay = dst.Copy()
        dstDay.Clear()
    
        CurrentDate = ReportStartDate
        RptDate = ReportStartDate
    
        For Each row In dst.Rows
            RptDate = row.Item("ReportingDate")
            If (CurrentDate < RptDate) Then
                AddEmptyDays(dstDay, CurrentDate, RptDate)
                dstDay.ImportRow(row)
            Else
                dstDay.ImportRow(row)
            End If
            CurrentDate = RptDate.AddDays(1)
        Next
    
        If (ReportEndDate > RptDate) Then
            If (RptDate = ReportStartDate) Then
                AddEmptyDays(dstDay, RptDate, ReportEndDate.AddDays(1))
            Else
                AddEmptyDays(dstDay, RptDate.AddDays(1), ReportEndDate.AddDays(1))
            End If
        End If
    
        Return dstDay
    End Function
  
    '----------------------------------------------------------------------------------
  
    Sub AddEmptyDays(ByRef dst As DataTable, ByVal StartDate As Date, ByVal EndDate As Date)
        Dim CurrentDate As Date
        Dim row As DataRow
        CurrentDate = StartDate
        While (CurrentDate < EndDate)
            row = dst.NewRow()
            row.Item("ReportingDate") = CurrentDate
            row.Item("NumRedemptions") = 0
            row.Item("NumImpressions") = 0
            row.Item("NumTransactions") = 0
            dst.Rows.Add(row)
            CurrentDate = CurrentDate.AddDays(1)
        End While
    End Sub
  
    '----------------------------------------------------------------------------------
  
    Function ExportReport(ByVal dst As DataTable, ByVal frequency As String) As String
        Dim builder As StringBuilder = New StringBuilder()
    
        If (Not dst Is Nothing) Then
            builder.Append(",")
            builder.Append(WriteExportRow(dst, "ReportingDate", False))
            If (frequency.Contains("weekly")) Then
                builder.Append("Impressions (weekly),")
            Else
                builder.Append("Impressions (daily),")
            End If
            builder.Append(WriteExportRow(dst, "NumImpressions", False))
            builder.Append("Impressions (cumulative),")
            builder.Append(WriteExportRow(dst, "NumImpressions", True))
            If (frequency.Contains("weekly")) Then
                builder.Append("Redemptions (weekly),")
            Else
                builder.Append("Redemptions (daily),")
            End If
            builder.Append(WriteExportRow(dst, "NumRedemptions", False))
            builder.Append("Redemptions (cumulative),")
            builder.Append(WriteExportRow(dst, "NumRedemptions", True))
            If (frequency.Contains("weekly")) Then
                builder.Append("Mark Downs ($) (weekly),")
            Else
                builder.Append("Mark Downs ($) (daily),")
            End If
            builder.Append(WriteExportRow(dst, "AmountRedeemed", False))
            builder.Append("Mark Downs ($) (cumulative),")
            builder.Append(WriteExportRow(dst, "AmountRedeemed", True))
            builder.Append("Redemption Rate,")
            builder.Append(WriteRedemptionRow(dst))
        End If
		
        Return builder.ToString
    End Function
  
    '----------------------------------------------------------------------------------
  
    Function WriteExportRow(ByVal dst As DataTable, ByVal field As String, ByVal bCumulative As Boolean) As String
        Dim builder As StringBuilder = New StringBuilder()
        Dim RowCount, i As Integer
        Dim cumulative As Double
        Dim dt As Date
    
        RowCount = dst.Rows.Count
    
        For i = 0 To (RowCount - 1)
            If (field = "" OrElse IsDBNull(dst.Rows(i).Item(field))) Then
                builder.Append("0")
            Else
                If (bCumulative) Then
                    If Not IsDBNull(dst.Rows(i).Item(field)) Then cumulative += dst.Rows(i).Item(field)
                    builder.Append(cumulative)
                Else
                    If (IsDate(dst.Rows(i).Item(field))) Then
                        dt = dst.Rows(i).Item(field)
                        builder.Append(Logix.ToShortDateString(dt, MyCommon))
                    Else
                        builder.Append(dst.Rows(i).Item(field))
                    End If
                End If
            End If
            If (i = (RowCount - 1)) Then
                builder.Append(vbNewLine)
            Else
                builder.Append(",")
            End If
        Next
        Return builder.ToString()
    End Function
  
    '----------------------------------------------------------------------------------
  
    Function WriteRedemptionRow(ByVal dst As DataTable) As String
        Dim builder As StringBuilder = New StringBuilder()
        Dim RowCount, i As Integer
        Dim Impressions, Redemptions As Integer
        Dim RedemptionRate As Double
    
        RowCount = dst.Rows.Count
    
        For i = 0 To (RowCount - 1)
            Impressions = MyCommon.NZ(dst.Rows(i).Item("NumImpressions"), 0)
            Redemptions = MyCommon.NZ(dst.Rows(i).Item("NumRedemptions"), 0)
            If (Impressions > 0) Then
                RedemptionRate = Redemptions / Impressions
            Else
                RedemptionRate = 0.0
            End If
            builder.Append(RedemptionRate.ToString("0.####"))
            If (i = (RowCount - 1)) Then
                builder.Append(vbNewLine)
            Else
                builder.Append(",")
            End If
        Next
        Return builder.ToString()
    End Function
  
    Function GetIPortalQualifier() As String
        Dim IPortalStr As String = ""
        Dim IPortalStart As Integer = -1
        Dim IPortalEnd As Integer = -1
        Dim UriStr As String = ""
    
        ' check if this is an IPortal request, if so then get the iportal token from the url and
        ' prepend it later to all form action values or link hrefs.
        UriStr = Page.Request.Url.AbsoluteUri
        If (UriStr IsNot Nothing) Then
            IPortalStart = UriStr.IndexOf("/,")
            If (IPortalStart > -1) Then
                IPortalEnd = UriStr.IndexOf("+", IPortalStart)
                If (IPortalEnd > -1) Then
                    IPortalStr = Left(Request.Url.ToString, Request.Url.ToString.LastIndexOf("/") + 1)
                    IPortalStr &= UriStr.Substring(IPortalStart + 1, (IPortalEnd - IPortalStart))
                End If
            End If
        End If

        Return IPortalStr
    End Function
  
    Sub HandleLMGSave(ByVal CheckDate As String, ByVal OfferID As Long)
        Dim rst As DataTable
        Dim StartDate As New Date
        Dim EndDate As New Date
        Dim CheckDateClean As New Date
        Dim OutputString As String = ""
    
        Response.Cache.SetCacheability(HttpCacheability.NoCache)
        MyCommon.Open_LogixRT()
    
        Try
            If Date.TryParse(CheckDate, CheckDateClean) AndAlso IsNumeric(OfferID) Then
                MyCommon.QueryStr = "select StartDate, EndDate from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID & ";"
                rst = MyCommon.LRT_Select
                If rst.Rows.Count > 0 Then
                    StartDate = MyCommon.NZ(rst.Rows(0).Item("StartDate"), New Date(1980, 1, 1))
                    EndDate = MyCommon.NZ(rst.Rows(0).Item("EndDate"), New Date(1980, 1, 1))
                    CheckDateClean = New Date(CheckDateClean.Year, CheckDateClean.Month, CheckDateClean.Day)
                    If (CheckDateClean < StartDate) OrElse (CheckDateClean > EndDate) Then
                        OutputString = Copient.PhraseLib.Lookup("lmg.DateOutOfRange", LanguageID)
                    Else
                        OutputString = ""
                    End If
                Else
                    OutputString = Copient.PhraseLib.Detokenize("offer.DoesNotExist", LanguageID, OfferID)
                End If
            End If
        Catch ex As Exception
            Send(ex.ToString)
        Finally
        End Try
        MyCommon.Close_LogixRT()
    
        Sendb(OutputString)
    End Sub
  
    Function CleanProductCodes(ByVal ProductCodes As String) As String
        Dim cleanedString As String = ""
        Dim tempProducts As String = ""
        tempProducts = ProductCodes.Replace(vbLf, ",")
        Dim charsToTrim() As Char = {","c, " "c}
        tempProducts = tempProducts.Trim(charsToTrim)
        cleanedString = tempProducts.Replace(",,", ",")
        cleanedString = cleanedString.Replace(", ,", ",")
			
        Return cleanedString
    End Function
	
    Sub ModifyProductsProductGroups(ByVal ProductGroupID As String, ByVal Products As String, ByVal OperationType As Integer, ByVal ProductType As Integer)
        Dim tempProducts As String = ""
        'Dim tempProductsList() As String = Nothing
        Dim validItemList As List(Of String) = New List(Of String)
        Dim invalidItemList As List(Of String) = New List(Of String)
        Dim DuplicateItemCount As Integer = 0
        Dim OutputString As String = ""
        Dim tempTableInsertStatement As StringBuilder = New StringBuilder()
        Dim duplicateItemsList As StringBuilder = New StringBuilder()
        Dim maxLimit As Integer = 0

        MyCommon.Write_Log(LogFile, "Started at =" & DateTime.Now & vbCr, True)

        If (Not String.IsNullOrEmpty(Products)) Then
            tempProducts = CleanProductCodes(Products)
            'tempProductsList = tempProducts.Split(",")
            Integer.TryParse(MyCommon.Fetch_SystemOption(166), maxLimit)

            Dim dtProductsList As DataTable = New DataTable("TableParameter")
            dtProductsList.Columns.Add("ProductID")
            dtProductsList.Columns.Add("ProductTypeID")
            dtProductsList.Columns.Add("Description")
            dtProductsList.AcceptChanges()
            Dim strProduct = String.Empty
            For Each item In tempProducts.Split(",")
                Dim Dr As DataRow = dtProductsList.NewRow()
                strProduct = GenerateProductCodeWithPadding(ProductType, Trim(item))
                Dr(0) = strProduct
                Dr(1) = ProductType
                Dr(2) = ""
                dtProductsList.Rows.Add(Dr)
                dtProductsList.AcceptChanges()
            Next

            If (dtProductsList.Rows.Count > 0 AndAlso maxLimit = 0) Or (dtProductsList.Rows.Count > 0 AndAlso dtProductsList.Rows.Count <= maxLimit) Then
                Dim productUPC As String = ""

                MyCommon.QueryStr = "dbo.pa_PUA_UpdateProductNew"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@ProductsUpdateType", System.Data.SqlDbType.Structured).Value = dtProductsList
                Dim dtnew As DataTable = MyCommon.LRTsp_select()
                MyCommon.Close_LRTsp()

                Dim Validexpression As String = "IsValid = 1"
                Dim foundValidRows As DataRow() = dtnew.[Select](Validexpression)
                If foundValidRows.Length > 0 Then
                    For Each drValidItem In foundValidRows
                        validItemList.Add(Trim(drValidItem("TblProductID")))
                        tempTableInsertStatement.Append("insert into #TempProdPK values(" & drValidItem("TblProductID") & "); ")
                    Next
                End If

                Dim inValidexpression As String = "IsValid = 0 AND TblProductID > 0"
                Dim foundInValidRows As DataRow() = dtnew.[Select](inValidexpression)
                If foundInValidRows.Length > 0 Then
                    For Each drInValidItem In foundInValidRows
                        invalidItemList.Add(Trim(drInValidItem("TblProductID")))
                    Next
                End If

                Dim duplicateexpression As String = "TblProductID is null"
                Dim foundDuplicateRows As DataRow() = dtnew.[Select](duplicateexpression)
                If foundDuplicateRows.Length > 0 Then
                    For Each drDuplicateItem In foundDuplicateRows
                        duplicateItemsList.Append("Duplicate Item =" & Trim(drDuplicateItem("ExtProductId")) & vbCr)
                    Next
                    DuplicateItemCount = foundDuplicateRows.Count
                    MyCommon.Write_Log(LogFile, Copient.PhraseLib.Lookup("term.prod_dup", LanguageID), True)
                    MyCommon.Write_Log(LogFile, duplicateItemsList.ToString(), True)
                End If

                If (validItemList.Count > 0) Then
                    SaveProductToProductGroup(tempTableInsertStatement.ToString(), ProductGroupID, OperationType, ProductType, validItemList)
                End If

                If (invalidItemList.Count > 0) Then
                    OutputString &= "Fail" & "~|"
                    For Each invaliditem In invalidItemList
                        OutputString &= invaliditem & vbCrLf
                    Next
                Else
                    OutputString = "OK"
                End If
            Else
                OutputString = "Invalid" & "~|" & Copient.PhraseLib.Lookup("pgroup-edit.maxlimit", LanguageID) & " : " & maxLimit
            End If
        End If
        MyCommon.Write_Log(LogFile, "OutputString =" & OutputString & vbCr, True)
        MyCommon.Write_Log(LogFile, "ProductGroupID =" & ProductGroupID & ",OpertaionType=" & OperationType & ",ProductType=" & ProductType & vbCr, True)
        MyCommon.Write_Log(LogFile, "Operation Type : 0 -Full Replace, 1 - Add to Group, 2- Remove from group" & vbCr, True)
        MyCommon.Write_Log(LogFile, "Total No of Codes=" & (validItemList.Count + invalidItemList.Count) & vbCr, True)
        MyCommon.Write_Log(LogFile, "Total No of invalid codes=" & (invalidItemList.Count) & vbCr, True)
        MyCommon.Write_Log(LogFile, "Total No of valid codes=" & (validItemList.Count) & vbCr, True)
        MyCommon.Write_Log(LogFile, "Total No of duplicate Codes=" & DuplicateItemCount & vbCr, True)
        MyCommon.Write_Log(LogFile, "Completed on =" & DateTime.Now & vbCr, True)
        MyCommon.Write_Log(LogFile, "_________________________________________________________________" & vbCr, True)

        Response.Write(OutputString)
    End Sub
 
    Sub ModifyProducts(ByVal Products As String, ByVal OperationType As Integer, ByVal ProductType As Integer, ByVal GName As String, ByVal RewardID As String, ByVal AdminUserID As Integer, ByVal IsCondition As Boolean)
   

        Dim ProductGroupID As Integer = 0
        Dim infoMessage As String = ""
        Dim DataToSend As String = ""
        Dim rst, rst1 As dataTable
        Dim tempProducts As String = ""
        Dim tempProductsList() As String = Nothing
        Dim maxLimit As Integer = 0
        Dim validItemList As List(Of String) = New List(Of String)
        Dim invalidItemList As List(Of String) = New List(Of String)
        Dim DuplicateItemCount As Integer = 0
        Dim OutputString As String = ""
        Dim tempTableInsertStatement As StringBuilder = New StringBuilder()


        Response.Cache.SetCacheability(HttpCacheability.NoCache)
        MyCommon.Open_LogixRT()

        If Not IsCondition Then
            MyCommon.QueryStr = "select ProductGroupID, ExcludedProdGroupID from OfferRewards with (NoLock) where RewardID=@RewardID and deleted=0;"
            MyCommon.DBParameters.Add("@RewardID", SqlDbType.BigInt).Value = RewardID
            rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            
        Else
            MyCommon.QueryStr = "select LinkID as ProductGroupID, ExcludedID as ExcludedProdGroupID from OfferConditions with (NoLock) where ConditionID=@RewardID and deleted=0;"
            MyCommon.DBParameters.Add("@RewardID", SqlDbType.BigInt).Value = RewardID
            rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        End If
	
        '    GName = GetCgiValue("modprodgroupname")
    
        If (rst.Rows(0).Item("ProductGroupID") = 0) Then 'Create new product group
      
            MyCommon.QueryStr = "SELECT ProductGroupID FROM ProductGroups with (NoLock) WHERE Name = '" & IIf(GName.Contains("'"), GName.Replace("'", "''"), GName) & "' AND Deleted=0"
            rst1 = MyCommon.LRT_Select
            If (rst1.Rows.Count > 0) Then
                ProductGroupID = rst1.Rows(0).Item("ProductGroupID")
            Else
                MyCommon.QueryStr = "dbo.pt_ProductGroups_Insert"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 200).Value = GName
                GName = MyCommon.Parse_Quotes(Logix.TrimAll(GName))
                MyCommon.LRTsp.Parameters.Add("@AnyProduct", SqlDbType.Bit).Value = 0
                MyCommon.LRTsp.Parameters.Add("@AdminID", SqlDbType.Int).Value = AdminUserID
                MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                MyCommon.LRTsp.ExecuteNonQuery()
                ProductGroupID = MyCommon.LRTsp.Parameters("@ProductGroupID").Value
                Send("<input type=""hidden"" id=""NewCreatedProdGroupID"" name=""NewCreatedProdGroupID"" value=""" & ProductGroupID & """ />")
                MyCommon.Activity_Log(5, ProductGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.pgroup-create", LanguageID))
                MyCommon.Close_LRTsp()
            End If
        Else
            ProductGroupID = rst.Rows(0).Item("ProductGroupID")
        End If
            
        
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
                        If (Not validItemList.Contains(GenerateProductCodeWithPadding(ProductType, item))) Then
                            If (IsValidItemCode(ProductGroupID, ProductType, item, OperationType, infoMessage)) Then
                                validItemList.Add(item)
                                tempTableInsertStatement.Append(SaveProduct(item, ProductType, OperationType, ProductGroupID))
                            Else
                                invalidItemList.Add(item)
                            End If
                        End If
                    End If
                Next
                
                If (validItemList.Count > 0) Then
                    SaveProductToProductGroup(tempTableInsertStatement.ToString(), ProductGroupID, OperationType, ProductType, validItemList)
                End If
            Else
                infoMessage = "Invalid" & "~|" & Copient.PhraseLib.Lookup("pgroup-edit.maxlimit", LanguageID) & " : " & maxLimit
            End If
        End If

        If invalidItemList.Count > 0 AndAlso (infoMessage = "" OrElse infoMessage Is Nothing) Then
            infoMessage = "There are " & invalidItemList.Count & " invalid items"
        End If
    
        DataToSend = infoMessage & "|" & ProductGroupID
        Response.Write(DataToSend)

   
    End Sub
  
    
    Private Function SaveProduct(ByVal ExtProductID As String, ByVal ProductTypeID As Integer, ByVal Operation As Integer, ByVal ProductGroupID As String) As String
        
        Dim ProductDesc As String = ""
        Dim querystr As String = ""
        Dim dst As DataTable
        Dim ProductID As Integer
        Dim bRTConnectionOpened As Boolean = False

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                MyCommon.Open_LogixRT()
                bRTConnectionOpened = True
            End If
            MyCommon.QueryStr = "select ProductID from Products with (NoLock) where (ExtProductID = @ExtProductID and " & _
                                " ProductTypeID =@ProductTypeID)"
            MyCommon.DBParameters.Add("@ExtProductID", SqlDbType.NVarChar).Value = ExtProductID
            MyCommon.DBParameters.Add("@ProductTypeID", SqlDbType.Int).Value = ProductTypeID
            dst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
				
            If dst.Rows.Count > 0 Then
                querystr &= "insert into #TempProdPK values(" & MyCommon.NZ(dst.Rows(0).Item("ProductID"), 0) & "); "
            Else
               
                'Create New Product
                If Operation <> 2 Then
                          
                    MyCommon.QueryStr = "dbo.pa_PUA_UpdateProduct"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@ExtProductID", SqlDbType.NVarChar, 120).Value = ExtProductID
                    MyCommon.LRTsp.Parameters.Add("@ProductTypeID", SqlDbType.Int).Value = ProductTypeID
                    MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 200).Value = ProductDesc
                    MyCommon.LRTsp.Parameters.Add("@ProductID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                    MyCommon.LRTsp.ExecuteNonQuery()
                    ProductID = MyCommon.LRTsp.Parameters("@ProductID").Value
          
                    MyCommon.QueryStr = "Select PhraseID,Name from ProductTypes where ProductTypeID=" & ProductTypeID & ";"
                    Dim productTypeTable As DataTable = MyCommon.LRT_Select()
                    Dim typePhrase As Integer = 0
                    If (productTypeTable.Rows.Count > 0) Then
                        typePhrase = MyCommon.NZ(productTypeTable.Rows(0).Item("PhraseID"), 0)
                    End If
                    MyCommon.Activity_Log(5, ProductGroupID, AdminUserID, Copient.PhraseLib.Lookup("term.created", LanguageID) & " " & Copient.PhraseLib.Lookup("term.product", LanguageID).ToLower() & " " & Copient.PhraseLib.Lookup("term.id", LanguageID) & " " & ExtProductID & _
                                          IIf(typePhrase > 0, "(" & Copient.PhraseLib.Lookup(typePhrase, LanguageID) & ")", ""))
                    MyCommon.Close_LRTsp()
                           
                    'Add to table
                    querystr = "insert into #TempProdPK values(" & ProductID & "); "
                     
                End If
                    
            End If
        Catch ex As Exception
            MyCommon.Write_Log(LogFile, "Method Name : Saveproduct " & vbCr & ex.Message & vbCr, True)
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed AndAlso bRTConnectionOpened Then MyCommon.Close_LogixRT()
        End Try
        Return querystr
    End Function

    Private Function SaveProductToProductGroup(ByVal insertStatement As String, ByVal ProductGroupID As Integer, ByVal Operation As String, ByVal ProductType As Integer, ByVal ProductIDList As List(Of String)) As Boolean
        
        Dim querystr As String
        Dim Status As Integer
        Dim Sucess As Boolean = True
        Dim bRTConnectionOpened As Boolean = False

        Try
                        
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                MyCommon.Open_LogixRT()
                bRTConnectionOpened = True
            End If

            MyCommon.QueryStr = "IF OBJECT_ID('tempdb..#TempProdPK') IS NOT NULL BEGIN drop table #TempProdPK END "
            MyCommon.LRT_Execute()
            querystr = "create table #TempProdPK([TempPK] int PRIMARY KEY IDENTITY," & _
                 "[ProductID] bigint NOT NULL)"
            
            '0 -Full Replace, 1 - Add to Group, 2- Remove from group
                    
            querystr &= insertStatement
           
            MyCommon.QueryStr = querystr
			
            MyCommon.LRT_Execute()
            
            If Operation = 0 Then
                MyCommon.QueryStr = "dbo.pt_LGM_ProductGroup_Replace"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt, 20).Value = ProductGroupID
                MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
                MyCommon.LRTsp.CommandTimeout = 2400
                MyCommon.LRTsp.ExecuteNonQuery()
                Status = MyCommon.LRTsp.Parameters("@Status").Value
                MyCommon.Close_LRTsp()
                MyCommon.QueryStr = " drop table #TempProdPK;"
                MyCommon.LRT_Execute()
                If Status = -2 Then
                    Sucess = False
                End If
               
            ElseIf Operation = 1 Then
                MyCommon.QueryStr = "dbo.pt_LGM_ProductGroup_Insert"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt, 20).Value = ProductGroupID
                MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
                MyCommon.LRTsp.CommandTimeout = 2400
                MyCommon.LRTsp.ExecuteNonQuery()
                Status = MyCommon.LRTsp.Parameters("@Status").Value
                MyCommon.Close_LRTsp()
                MyCommon.QueryStr = " drop table #TempProdPK;"
                MyCommon.LRT_Execute()
        
                If Status = 0 Then
                    For Each item In ProductIDList
                        MyCommon.Activity_Log(5, ProductGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.pgroup-add", LanguageID) & " " & item)
                    Next
          
                End If

                If Status = -2 Then
                    Sucess = False
                End If
               
            ElseIf Operation = 2 Then
                MyCommon.QueryStr = "dbo.pt_LGM_ProductGroup_Remove"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt, 20).Value = ProductGroupID
                MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
                MyCommon.LRTsp.CommandTimeout = 2400
                MyCommon.LRTsp.ExecuteNonQuery()
                Status = MyCommon.LRTsp.Parameters("@Status").Value
                MyCommon.Close_LRTsp()
                MyCommon.QueryStr = " drop table #TempProdPK;"
                For Each validitem In ProductIDList
                    MyCommon.Activity_Log(5, ProductGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.pgroup-remove", LanguageID) & " " & validitem)
                Next
                MyCommon.LRT_Execute()
                If Status = -2 Then
                    Sucess = False
                End If
                
            End If
        
            'When full replace then automatically redeploy
            If Operation = 0 Then
                MyCommon.QueryStr = "update productgroups with (RowLock) set LastUpdate=getdate(), CMOAStatusFlag=2 where ProductGroupID=" & ProductGroupID
                MyCommon.LRT_Execute()
            Else
                MyCommon.QueryStr = "update productgroups with (RowLock) set  LastUpdate=getdate() where ProductGroupID=" & ProductGroupID
                MyCommon.LRT_Execute()
            End If
			
        Catch ex As Exception
            Sucess = False
            MyCommon.Write_Log(LogFile, "Method Name : SaveProductToProductGroup " & vbCr & ex.Message & vbCr, True)
        Finally
            MyCommon.QueryStr = "IF OBJECT_ID('tempdb..#TempProdPK') IS NOT NULL BEGIN drop table #TempProdPK END "
            MyCommon.LRT_Execute()
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed AndAlso bRTConnectionOpened Then MyCommon.Close_LogixRT()
        End Try

        Return Sucess
    End Function

  
    Function LoadMonthNames() As String()
        Dim MonthNames(-1) As String
    
        ' Localize the month abbreviation to the users language
        If MyCommon.GetAdminUser.Culture IsNot Nothing Then
            MonthNames = MyCommon.GetAdminUser.Culture.DateTimeFormat.AbbreviatedMonthNames
        End If

        Return MonthNames
    End Function

    Function GetMonthAbbreviation(ByVal MonthNumber As Integer, ByVal MonthAbbreviations As String()) As String
        Dim Abbr As String = ""
    
        If MonthAbbreviations IsNot Nothing AndAlso (MonthNumber - 1) >= MonthAbbreviations.Length Then
            Abbr = MonthAbbreviations(MonthNumber - 1)
        Else
            Abbr = Left(MonthName(MonthNumber), 3)
        End If
    
        Return Abbr
    End Function
    '----------------------------------------------------------------------------------
  
    Sub GenerateCustomReport(ByVal ReportID As String)
        Dim builder As StringBuilder = New StringBuilder()
        Dim bParsed As Boolean = False
        Dim LanguageID As Integer = 1
        Dim Frequency As String = 1
    
        Response.Cache.SetCacheability(HttpCacheability.NoCache)
    
        If (Request.Form("lang") <> "") Then
            bParsed = Integer.TryParse(Request.Form("lang"), LanguageID)
            If (Not bParsed) Then LanguageID = 1
        End If
        If (Request.Form("frequency") <> "") Then
            Frequency = Request.Form("frequency")
        End If
    
        builder.Append(CreateCustomReport(LanguageID))
        Response.Write(builder)
    End Sub
  
    '----------------------------------------------------------------------------------
  
    Function CreateCustomReport(ByVal LanguageID As Integer) As String
        Dim ReportStartDate As Date
        Dim ReportEndDate As Date
        Dim ReportWeeks As Integer
        Dim RowCount As Integer
        Dim CumulativeImpress As Double
        Dim CumulativeRedeem As Double
        Dim CumulativeTrans As Double
        Dim CumulativeAmtRedeem As Double
        Dim RedemptionRate As Double
        Dim AmtRedeem As Double
        Dim Transactions As Double
        Dim Redemptions As Double
        Dim Impressions As Double
        Dim i As Integer
        Dim OfferID As String = ""
        Dim dst As System.Data.DataTable
        Dim bParsed As Boolean
        Dim ExportRequested As String
        Dim builder As StringBuilder = New StringBuilder()
        Dim frequency As String = ""
        Dim ofrIds As String = Request.Form("ofrIds")
        Dim OptList As String()

        Dim TableHeight As Integer = 174  '158
        Dim SpanHeight As Integer = 16
        Dim ShadedClass As String = ""

        Dim DivBoxHeight As Integer = 191    '175
    
        ' calculating DivBox height below
        If Request.Form("impressionChecked") = "false" Then
            DivBoxHeight = DivBoxHeight - SpanHeight
        End If
        If Request.Form("redemptionChecked") = "false" Then
            DivBoxHeight = DivBoxHeight - SpanHeight
        End If
        If Request.Form("transactionChecked") = "false" Then
            DivBoxHeight = DivBoxHeight - SpanHeight
        End If
        If Request.Form("markdownChecked") = "false" Then
            DivBoxHeight = DivBoxHeight - SpanHeight
        End If
        If Request.Form("CumuImpressionsChecked") = "false" Then
            DivBoxHeight = DivBoxHeight - SpanHeight
        End If
        If Request.Form("CumuRedemptionsChecked") = "false" Then
            DivBoxHeight = DivBoxHeight - SpanHeight
        End If
        If Request.Form("CumuTransactionsChecked") = "false" Then
            DivBoxHeight = DivBoxHeight - SpanHeight
        End If
        If Request.Form("CumuMarkdownsChecked") = "false" Then
            DivBoxHeight = DivBoxHeight - SpanHeight
        End If
	
        Dim WhereClause As New StringBuilder("")
   
        '     If Request.Form("impressionChecked") = "true" Then 
        If (Request.Form("Impressions") <> "" AndAlso Request.Form("ImpressionType") <> "") Then
            WhereClause.Append(IIf(WhereClause.Length > 0, " and ", " Where "))
            ' WhereClause.Append("NumImpressions" & ConvertToOperand(Request.Form("ImpressionType")) & Request.Form("Impressions"))
            WhereClause.Append("NumImpressions" & Request.Form("ImpressionType") & Request.Form("Impressions"))
        End If
        '     End If
        '     If Request.Form("redemptionChecked") = "true" Then   
        If (Request.Form("Redemptions") <> "" AndAlso Request.Form("RedemptionType") <> "") Then
            WhereClause.Append(IIf(WhereClause.Length > 0, " and ", " Where "))
            ' WhereClause.Append("NumRedemptions" & ConvertToOperand(Request.Form("RedemptionType")) & Request.Form("Redemptions"))
            WhereClause.Append("NumRedemptions" & Request.Form("RedemptionType") & Request.Form("Redemptions"))
        End If
        '     End If
        '     If Request.Form("transactionChecked") = "true" Then   
        If (Request.Form("Transactions") <> "" AndAlso Request.Form("TransactionType") <> "") Then
            WhereClause.Append(IIf(WhereClause.Length > 0, " and ", " Where "))
            ' WhereClause.Append("NumTransactions" & ConvertToOperand(Request.Form("TransactionType")) & Request.Form("Transactions"))
            WhereClause.Append("NumTransactions" & Request.Form("TransactionType") & Request.Form("Transactions"))
        End If
        '     End If
        '     If Request.Form("markdownChecked") = "true" Then    
        If (Request.Form("MarkDowns") <> "" AndAlso Request.Form("MarkDownType") <> "") Then
            WhereClause.Append(IIf(WhereClause.Length > 0, " and ", " Where "))
            'WhereClause.Append("AmountRedeemed" & ConvertToOperand(Request.Form("MarkDownType")) & Request.Form("MarkDowns"))
            WhereClause.Append("AmountRedeemed" & Request.Form("MarkDownType") & Request.Form("MarkDowns"))
        End If
        '     End If

        If WhereClause.Length > 0 Then
            WhereClause.Append(" and OfferID = ")
        Else
            WhereClause.Append("where OfferID = ")
        End If
    
        ' calculating Table height below
        If Request.Form("impressionChecked") = "false" Then
            TableHeight = TableHeight - SpanHeight
        End If
        If Request.Form("redemptionChecked") = "false" Then
            TableHeight = TableHeight - SpanHeight
        End If
        If Request.Form("transactionChecked") = "false" Then
            TableHeight = TableHeight - SpanHeight
        End If
        If Request.Form("markdownChecked") = "false" Then
            TableHeight = TableHeight - SpanHeight
        End If
        If Request.Form("CumuImpressionsChecked") = "false" Then
            TableHeight = TableHeight - SpanHeight
        End If
        If Request.Form("CumuRedemptionsChecked") = "false" Then
            TableHeight = TableHeight - SpanHeight
        End If
        If Request.Form("CumuTransactionsChecked") = "false" Then
            TableHeight = TableHeight - SpanHeight
        End If
        If Request.Form("CumuMarkdownsChecked") = "false" Then
            TableHeight = TableHeight - SpanHeight
        End If
        ' calculating Table height above
    
        If (Request.Form("ReportingType") = "Between") Then
            bParsed = DateTime.TryParse(Request.Form("reportstart"), ReportStartDate)
            If (Not bParsed) Then ReportStartDate = Now()
            bParsed = DateTime.TryParse(Request.Form("reportend"), ReportEndDate)
            If (Not bParsed) Then ReportEndDate = Now()
        ElseIf (Request.Form("ReportingType") = "<=") Then
            bParsed = DateTime.TryParse(Request.Form("reportstart"), ReportEndDate)
            If (Not bParsed) Then
                ReportEndDate = Now()
            End If
            ReportStartDate = ReportEndDate.AddDays(-30)
        ElseIf (Request.Form("ReportingType") = "<") Then
            bParsed = DateTime.TryParse(Request.Form("reportstart"), ReportEndDate)
            If (Not bParsed) Then
                ReportEndDate = Now()
            End If
            ReportEndDate = ReportEndDate.Date.AddDays(-1)
            ReportStartDate = ReportEndDate.Date.AddDays(-30)
            ReportEndDate = ReportEndDate.AddTicks(-1)
        ElseIf (Request.Form("ReportingType") = ">=") Then
            bParsed = DateTime.TryParse(Request.Form("reportstart"), ReportStartDate)
            If (Not bParsed) Then ReportStartDate = Now()
            ReportEndDate = Now()
        ElseIf (Request.Form("ReportingType") = ">") Then
            bParsed = DateTime.TryParse(Request.Form("reportstart"), ReportStartDate)
            If (Not bParsed) Then
                ReportStartDate = Now()
            End If
            ReportStartDate = ReportStartDate.Date.AddDays(1)
            ReportEndDate = Now()
        Else   'Request.Form("ReportingType") = "="
            bParsed = DateTime.TryParse(Request.Form("reportstart"), ReportStartDate)
            If (Not bParsed) Then ReportStartDate = Now()
            ReportStartDate = ReportStartDate.Date
            ReportEndDate = ReportStartDate.AddDays(1).AddTicks(-1)
        End If

        ReportWeeks = DateTime.Compare(ReportStartDate, ReportEndDate) / 7

        ofrIds = Replace(ofrIds, vbCrLf, "")
        ofrIds = Replace(ofrIds, vbCr, "")
        OptList = ofrIds.Split(",")
        OfferID = Request.Form("offerId")

        ExportRequested = Request.Form("exportRpt")
    
        MyCommon.Open_LogixWH()
       
        For count = 0 To OptList.Length - 1
            If (OptList(count) <> "") Then
                OfferID = OptList(count)

                MyCommon.QueryStr = "select NumImpressions, NumRedemptions, NumTransactions, AmountRedeemed, ReportingDate from OfferReporting with (nolock) " & _
                                 WhereClause.ToString & OfferID & " " & _
                                 "and ReportingDate between '" & ReportStartDate.ToString & "' and '" & ReportEndDate.ToString & "' " & _
                                 "order by ReportingDate"
                          

                dst = MyCommon.LWH_Select
    
                If (Request.Form("frequency") = "1") Then
                    dst = RollupReportWeek(dst, ReportStartDate, ReportEndDate)
                    frequency = "weekly"
                ElseIf (Request.Form("frequency") = "2") Then
                    dst = FillInDays(dst, ReportStartDate, ReportEndDate)
                    frequency = "daily"
                End If
    
                If (ExportRequested <> "") Then
                    Response.ClearHeaders()
                    Response.AddHeader("Content-Disposition", "attachment; filename=Offer" & OfferID & "_Rpt.csv")
                    Response.ContentType = "application/octet-stream"
                    Response.Clear()
                    Return ExportReport(dst, frequency)
                End If

                ' Reset cumulative variables
                CumulativeImpress = 0
                CumulativeRedeem = 0
                CumulativeTrans = 0
                CumulativeAmtRedeem = 0
   
                RowCount = dst.Rows.Count
    
                If (RowCount > 0) Then
  
                    builder.Append("<div class=""box"" style=""width:97%; height: ")
                    builder.Append(DivBoxHeight.toString)
                    builder.Append("px;overflow-x:scroll;"">")
                    builder.Append("<table id=""reportStats"" class=""raligntable"" cellpadding=""2"" style="" margin-left:-4px; height:")
                    builder.Append(TableHeight.toString)
                    builder.Append("px;"" cellspacing=""0"" summary=""" & Copient.PhraseLib.Lookup("term.reports", LanguageID) & """")
                    builder.Append(Copient.PhraseLib.Lookup("term.impressions", LanguageID))
                    builder.Append(""">")
                    builder.Append("<tr style=""height:24px;"">")
                    For i = 0 To (RowCount - 1)
                        builder.Append("<th scope=""col"" style=""line-height: 20px; color: White;"">&nbsp;&nbsp;&nbsp;&nbsp;" & FormatDateTime(dst.Rows(i).Item("ReportingDate").ToString().Trim(), vbShortDate) & "</th>")
                    Next
                    builder.Append("</tr>")

                    If (Request.Form("impressionChecked") = "true") Then
                        builder.Append("<tr id=""rowBody1"" ")
                        builder.Append(ShadedClass)
                        If ShadedClass = "" And TableHeight > 62 Then
                            ShadedClass = "class=""shaded"""
                        Else
                            ShadedClass = ""
                        End If
                        builder.Append(" ondblclick=""javascript:toggleHighlight(1);"">")
                        For i = 0 To (RowCount - 1)
                            Impressions = MyCommon.NZ(dst.Rows(i).Item("NumImpressions"), 0)
                            builder.Append("<td>" & Impressions & "</td>")
                        Next
                        builder.Append("</tr>")
                    End If

                    If (Request.Form("CumuImpressionsChecked") = "true") Then
                        builder.Append("<tr id=""rowBody2"" ")
                        builder.Append(ShadedClass)
                        If ShadedClass = "" And (Request.Form("redemptionChecked") = "true" Or Request.Form("CumuRedemptionsChecked") = "true" Or Request.Form("markdownChecked") = "true" Or Request.Form("CumuMarkdownsChecked") = "true") Then
                            ShadedClass = "class=""shaded"""
                        Else
                            ShadedClass = ""
                        End If
                        builder.Append(" ondblclick=""javascript:toggleHighlight(2);"">")
                        For i = 0 To (RowCount - 1)
                            CumulativeImpress = CumulativeImpress + MyCommon.NZ(dst.Rows(i).Item("NumImpressions"), 0)
                            builder.Append("<td>" & CumulativeImpress & "</td>")
                        Next
                        builder.Append("</tr>")
                    End If

                    If (Request.Form("redemptionChecked") = "true") Then
                        builder.Append("<tr id=""rowBody3""")
                        builder.Append(ShadedClass)
                        If ShadedClass = "" And (Request.Form("CumuRedemptionsChecked") = "true" Or Request.Form("transactionChecked") = "true" Or Request.Form("CumuTransactionsChecked") = "true") Then
                            ShadedClass = "class=""shaded"""
                        Else
                            ShadedClass = ""
                        End If
                        builder.Append("  ondblclick=""javascript:toggleHighlight(3);"">")
                        For i = 0 To (RowCount - 1)
                            builder.Append("<td>")
                            Redemptions = MyCommon.NZ(dst.Rows(i).Item("NumRedemptions"), 0)
                            If (Redemptions = 0) Then
                                builder.Append("-")
                            Else
                                builder.Append(Redemptions)
                            End If
                            builder.Append("</td>")
                        Next
                        builder.Append("</tr>")
                    End If

                    If (Request.Form("CumuRedemptionsChecked") = "true") Then
                        builder.Append("<tr id=""rowBody4"" ")
                        builder.Append(ShadedClass)
                        If ShadedClass = "" And (Request.Form("transactionChecked") = "true" Or Request.Form("CumuTransactionsChecked") = "true") Then
                            ShadedClass = "class=""shaded"""
                        Else
                            ShadedClass = ""
                        End If
                        builder.Append(" ondblclick=""javascript:toggleHighlight(4);"">")
                        For i = 0 To (RowCount - 1)
                            CumulativeRedeem = CumulativeRedeem + MyCommon.NZ(dst.Rows(i).Item("NumRedemptions"), 0)
                            If (CumulativeRedeem = 0) Then
                                builder.Append("<td>-</td>")
                            Else
                                builder.Append("<td>" & CumulativeRedeem & "</td>")
                            End If
                        Next
                        builder.Append("</tr>")
                    End If

                    If (Request.Form("transactionChecked") = "true") Then
                        builder.Append("<tr id=""rowBody5""")
                        builder.Append(ShadedClass)
                        If ShadedClass = "" And (Request.Form("CumuTransactionsChecked") = "true" Or Request.Form("markdownChecked") = "true" Or Request.Form("CumuMarkdownsChecked") = "true") Then
                            ShadedClass = "class=""shaded"""
                        Else
                            ShadedClass = ""
                        End If
                        builder.Append("  ondblclick=""javascript:toggleHighlight(5);"">")
                        For i = 0 To (RowCount - 1)
                            builder.Append("<td>")
                            Transactions = MyCommon.NZ(dst.Rows(i).Item("NumTransactions"), 0)
                            If (Transactions = 0) Then
                                builder.Append("-")
                            Else
                                builder.Append(Transactions)
                            End If
                            builder.Append("</td>")
                        Next
                        builder.Append("</tr>")
                    End If

                    If (Request.Form("CumuTransactionsChecked") = "true") Then
                        builder.Append("<tr id=""rowBody6"" ")
                        builder.Append(ShadedClass)
                        If ShadedClass = "" And (Request.Form("markdownChecked") = "true" Or Request.Form("CumuMarkdownsChecked") = "true") Then
                            ShadedClass = "class=""shaded"""
                        Else
                            ShadedClass = ""
                        End If
                        builder.Append(" ondblclick=""javascript:toggleHighlight(6);"">")
                        For i = 0 To (RowCount - 1)
                            CumulativeTrans = CumulativeTrans + MyCommon.NZ(dst.Rows(i).Item("NumTransactions"), 0)
                            If (CumulativeTrans = 0) Then
                                builder.Append("<td>-</td>")
                            Else
                                builder.Append("<td>" & CumulativeTrans & "</td>")
                            End If
                        Next
                        builder.Append("</tr>")
                    End If

                    If (Request.Form("markdownChecked") = "true") Then
                        builder.Append("<tr id=""rowBody7"" ")
                        builder.Append(ShadedClass)
                        If ShadedClass = "" And Request.Form("CumuMarkdownsChecked") = "true" Then
                            ShadedClass = "class=""shaded"""
                        Else
                            ShadedClass = ""
                        End If
                        builder.Append(" ondblclick=""javascript:toggleHighlight(7);"">")
                        For i = 0 To (RowCount - 1)
                            builder.Append("<td>")
                            AmtRedeem = MyCommon.NZ(dst.Rows(i).Item("AmountRedeemed"), 0.0)
                            If (AmtRedeem = 0) Then
                                builder.Append("-")
                            Else
                                builder.Append(AmtRedeem.ToString("$#,##0.00"))
                            End If
                            builder.Append("</td>")
                        Next
                        builder.Append("</tr>")
                    End If

                    If (Request.Form("CumuMarkdownsChecked") = "true") Then
                        builder.Append("<tr id=""rowBody8"" ")
                        builder.Append(ShadedClass)
                        ShadedClass = ""
                        builder.Append(" ondblclick=""javascript:toggleHighlight(8);"">")
                        For i = 0 To (RowCount - 1)
                            CumulativeAmtRedeem = CumulativeAmtRedeem + MyCommon.NZ(dst.Rows(i).Item("AmountRedeemed"), 0.0)
                            If (CumulativeAmtRedeem = 0.0) Then
                                builder.Append("<td>-</td>")
                            Else
                                builder.Append("<td>" & CumulativeAmtRedeem.ToString("$#,##0.00") & "</td>")
                            End If
                        Next
                        builder.Append("</tr>")
                    End If

                    builder.Append("</table>")
                    builder.Append("<input type=""hidden"" id=""colCt"" name=""colCt"" value=""" & RowCount & """ />")
                    builder.Append("</div>")
                    MyCommon.Close_LogixWH()
                Else
                    builder.Append("<div class=""box"" style=""width: 97%; height: ")
                    builder.Append(DivBoxHeight.toString)
                    builder.Append("px;line-height: 20px;top:2px; color: White;"">")
                    builder.Append("<b>" & Copient.PhraseLib.Lookup("reports.nodata", LanguageID) & "</b>")
                    builder.Append("</div>")
                End If

            End If
        Next
         
        Return builder.ToString
    End Function
 
  </script>
