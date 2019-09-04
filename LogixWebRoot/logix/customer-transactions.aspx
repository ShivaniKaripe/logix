<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: customer-transactions.aspx 
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
  
  ' Page notes:
  '
  ' Referenced Request.QueryString variables
  '   CustPK - Primary key of the customer
  '   CustomerPK - alternative to 'CustPK'
  '   CardPK - Primary key of the card (supercedes the customer?)
  '   mode - Looks like the valid values are "summary" and Nothing
  '     exiturl - Only used if mode = "summary"
  '     cardnumber - Only used if mode = "summary"
  '   redemptionFilter - Indicates whether to display transactions with/without redemptions
  '   searchterms - Used in a lot of the queries. Very important
  '   Search - Only used with searchterms. Important if = "Search"
  '   searchPressed - Only used once. Important if = "Search"
  '   editterms - Used once to assign CustExtID
  '   transterms - Used to generate having and search filters.
  '   sortcol - Column on which to sort
  '   sortdir - Direction by which to sort
  '   pagenum - Page of the transactions that we're viewing
  '   showall - Show all transactions for the customer (related to paging, not to redemption status)
  '
  
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
  Dim rstTrans As DataTable
  Dim rstTemp As DataTable
  Dim row As DataRow
  Dim rst2 As DataTable
  Dim row2 As DataRow
  Dim rst3 As DataTable
  Dim dt As DataTable
  Dim CustomerPK As Long = 0
  Dim CardPK As Long
  Dim ExtCardID As String = ""
  Dim FullName As String = ""
  Dim TotalRedeemCt As Integer = 0
  Dim TotalRedeemAmt As Double = 0.0
  Dim CustExtID As String = ""
  Dim i As Integer = 0
  Dim transCt As Integer = 0
  Dim transOffers As StringBuilder
  Dim transRdmptAmt As StringBuilder
  Dim transRdmptCt As StringBuilder
  Dim OfferName As String = ""
  Dim XID As String = ""
  Dim IsPtsOffer As Boolean = False
  Dim IsAccumOffer As Boolean = False
  Dim UnknownPhrase As String = ""
  Dim IsHouseholdID As Boolean = False
  Dim HHPK As Integer = 0
  Dim HouseholdID As String = ""
  Dim HHCustIdList As New ArrayList(5)
  Dim PaddedExtID As String = StrDup(25, "0")
    Dim CustExtIdList As String = ""
    Dim dCustExtIdList As String = ""
  Dim CustPKs As String = ""
  Dim LogixTransNums As String = ""
  Dim Shaded As String = " class=""shaded"""
  Dim CustomerTypeID As Integer = 0
  Dim OfferID As Integer = 0
  Dim ClientUserID1 As String = ""
  Dim IDLength As Integer = 0
  Dim CustomerGroupIDs As String() = Nothing
  Dim loopCtr As Integer = 0
  Dim searchterms As String = ""
  Dim restrictLinks As Boolean = False
  Dim PointsIDBuf As New StringBuilder()
  Dim PointsNameBuf As New StringBuilder()
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim TransDateStart As String = ""
  Dim TransDateEnd As String = ""
  Dim CardTypeID As String = ""
  
  ' default urls for links from this page
  Dim URLOfferSum As String = "offer-sum.aspx"
  Dim URLCPEOfferSum As String = "CPEoffer-sum.aspx"
  Dim URLcgroupedit As String = "cgroup-edit.aspx"
  Dim URLpointedit As String = "point-edit.aspx"
  Dim URLtrackBack As String = ""
  Dim inCardNumber As String = ""
  ' tack on the customercare remote links if needed
  Dim extraLink As String = ""
  
  Dim PageNum As Integer = 0
  Dim MorePages As Boolean
  Dim linesPerPage As Integer = 10
  Dim sizeOfData As Integer
  Dim startPosition As Integer
  Dim endPosition As Integer
  Dim TransTerms As String
  Dim SearchFilter As String = ""
  Dim SearchFilterTH As String = ""
  Dim HavingFilterTR As String = ""
  Dim HavingFilterTH As String = ""
  Dim tempDate As Date
  Dim SortCol As String = "TransactionDate"
  Dim SortDir As String = "desc"
  Dim SortUrl As String = ""
  Dim NoOfdays As Integer = MyCommon.Extract_Val(MyCommon.NZ(MyCommon.Fetch_SystemOption(308),0))
  Dim redemptionFilter As Integer = 2
  Dim IsCMOnly As Boolean = False
  Dim CPEHHFilter As String = ""
  Dim IncludeAllCards As Integer = 0	'...CLOUDSOL-1252	
  
  Dim iCmAutoHouseholdCustGrpOptionId As Integer = 24
  Dim bCmAutoHouseholdCustGrpEnabled As Boolean = False
  Dim InstalledEngines(-1) As Integer
  
  Dim ShowAll As Boolean = False
  Dim ContextStyle As String = ""
  Dim TrxTotalDisplayOnUI As Integer = 0
  
  'Additions for Transaction Inquiry
  Dim sStartTime As String = ""
  Dim sEndTime As String = ""
  Dim LocationID As Integer
  Dim ExtLocationCode As String = ""
  Dim LocationName As String
  Dim sLast4CardID As String = ""
  Dim PresentedCustomerID As String
  Dim DateFilter As String = ""
  Dim doc As XDocument
  
    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If
  

    Response.Expires = 0
    MyCommon.AppName = "customer-transactions.aspx"
    MyCommon.Open_LogixRT()
    MyCommon.Open_LogixXS()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
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
  
    CustomerPK = Convert.ToInt32(Request.QueryString("CustPK"))
    CardPK = Convert.ToInt32(Request.QueryString("CardPK"))
  ExtLocationCode = Request.QueryString("LocCode")
  sLast4CardID = Request.QueryString("Last4")
  sStartTime = Request.QueryString("StartTime")
  sEndTime = Request.QueryString("EndTime")
  
 
  
    If CardPK > 0 Then
        ExtCardID = MyLookup.FindExtCardIDFromCardPK(CardPK)
    End If
  
    ' show all?
    If (Request.QueryString("showall") <> "") Then
        ShowAll = True
    Else
        ShowAll = False
    End If
  
    ' special handling for customer inquery direct link in 
    If (restrictLinks) Then
        URLOfferSum = ""
        URLCPEOfferSum = ""
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
  
    If Not (Convert.ToInt32(Request.QueryString("CustPK")) <> 0) Then
    If (Request.QueryString("LocCode") = "") Then
        Response.Status = "301 Moved Permanently"
        Response.AddHeader("Location", "customer-inquiry.aspx")
    End If
    End If
  
    'Set the default redemption filter value, then update if present in querystring
    redemptionFilter = MyCommon.Fetch_SystemOption(104)
    If (NoOfdays <= 0 AndAlso redemptionFilter = 4) Then
        'As the value for No.of days is zero.  Customer inquiry transactions default view  should not the have (All Transaction - x Days) as part of the options.
        redemptionFilter = 2
    End If
    If (Request.QueryString("redemptionFilter") <> "") Then
        redemptionFilter = Convert.ToInt32(Request.QueryString("redemptionFilter"))
    End If
  
    ' <2.3.12>- Fetch the system option 120, Transaction Total Display on UI
    Try
        TrxTotalDisplayOnUI = CInt(MyCommon.Fetch_SystemOption(120))
    Catch ex As Exception
        TrxTotalDisplayOnUI = -1
    End Try
  
    If (Convert.ToInt32(Request.QueryString("CustPK")) > 0 Or (Request.QueryString("searchterms") <> "" And _
    (Request.QueryString("Search") <> "" Or Request.QueryString("searchPressed") <> "")) Or _
      inCardNumber <> "" _
      ) Then
        ' someone wants to search for a customer.  First lets get their primary key from our database
        If (Convert.ToInt32(Request.QueryString("CustPK")) > 0 Or (Convert.ToInt32(Request.QueryString("CustomerPK")) > 0)) Then
            If (Convert.ToInt32(Request.QueryString("CustPK")) > 0) Then
                CustomerPK = Convert.ToInt32(Request.QueryString("CustPK"))
            ElseIf (Convert.ToInt32(Request.QueryString("CustomerPK")) > 0) Then
                CustomerPK = Convert.ToInt32(Request.QueryString("CustomerPK"))
            End If
            MyCommon.QueryStr = "select C.CustomerPK, C.CustomerTypeID, C.HHPK, C.FirstName, C.MiddleName, C.LastName, " & _
                                "CE.Address, CE.City, CE.State, CE.Zip, CE.PhoneAsEntered, CE.email " & _
                                "from Customers C with (NoLock) left join CustomerExt CE with (NoLock) on CE.CustomerPK = C.CustomerPK " & _
                                "left join Customers C2 with (NoLock) on C2.CustomerPK = C.HHPK " & _
                                "where C.CustomerPK=" & CustomerPK
        Else
            ' IF the page was called from an outside application set ClientUserID1 to the outside passed in value
            If (inCardNumber <> "" And Request.QueryString("mode") = "summary") Then
                ClientUserID1 = MyCommon.Pad_ExtCardID(inCardNumber, Copient.commonShared.CardTypes.CUSTOMER)
                searchterms = Request.QueryString("searchterms")
                MyCommon.QueryStr = "select C.CustomerPK, C.CustomerTypeID, C.HHPK, C.FirstName, C.MiddleName, C.LastName, " & _
                                    "CE.Address, CE.City, CE.State, CE.Zip, CE.PhoneAsEntered, CE.email " & _
                                    "from Customers C with (NoLock) left join CustomerExt CE with (NoLock) on CE.CustomerPK=C.CustomerPK " & _
                                    "left join Customers C2 with (NoLock) on C2.CustomerPK=C.HHPK " & _
                                    "where C.CustomerPK=" & CustomerPK & ";"
            End If
            If (Request.QueryString("searchterms") <> "" And ClientUserID1 = "") Then
                ClientUserID1 = MyCommon.Pad_ExtCardID(Request.QueryString("searchterms"), Copient.commonShared.CardTypes.CUSTOMER)
                MyCommon.QueryStr = "select C.CustomerPK, C.CustomerTypeID, C.HHPK C.FirstName, C.MiddleName, C.LastName, " & _
                                    "CE.Address, CE.City, CE.State, CE.Zip, CE.PhoneAsEntered, CE.email " & _
                                    "from Customers C with (NoLock) left join CustomerExt CE with (NoLock) on CE.CustomerPK=C.CustomerPK " & _
                                    "left join Customers C2 with (NoLock) on C2.CustomerPK=C.HHPK " & _
                                    "where C.CustomerPK=" & CustomerPK & " or CE.PhoneDigitsOnly = '" & MyCryptLib.SQL_StringEncrypt(MyCommon.DigitsOnly(Request.QueryString("searchterms"))) & _
                                    "' or CE.email = '" & MyCryptLib.SQL_StringEncrypt(MyCommon.Parse_Quotes(Request.QueryString("searchterms"))) & "' or C.LastName like '%" & MyCommon.Parse_Quotes(Request.QueryString("searchterms")) & "%';"
            End If
        End If
        rstResults = MyCommon.LXS_Select
        If (rstResults.Rows.Count = 1) Then
            ' ok we found a primary key for the external id provided
            CustomerPK = rstResults.Rows(0).Item("CustomerPK")
            ClientUserID1 = MyLookup.FindExtCardIDFromCardPK(Convert.ToInt32(Request.QueryString("CardPK")))
            IsHouseholdID = (MyCommon.NZ(rstResults.Rows(0).Item("CustomerTypeID"), 0) = 1)
            CardTypeID = MyCommon.NZ(rstResults.Rows(0).Item("CustomerTypeID"), 0)
        Else
            infoMessage = "" & Copient.PhraseLib.Lookup("customer.notfound", LanguageID) & ""
            infoMessage = infoMessage & " <a href=""customer-inquiry.aspx?mode=add&Search=Search" & extraLink & "&searchterms=" & Request.QueryString("searchterms") & """>" & Copient.PhraseLib.Lookup("term.add", LanguageID) & "</a>"
        End If
    End If
  
    If IsHouseholdID And CustomerPK > 0 Then
        InstalledEngines = MyCommon.GetInstalledEngines
        If MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM) And InstalledEngines.Length = 1 Then
            IsCMOnly = True
            bCmAutoHouseholdCustGrpEnabled = MyCommon.Fetch_CM_SystemOption(iCmAutoHouseholdCustGrpOptionId)
            If bCmAutoHouseholdCustGrpEnabled Then
                Dim lCustomerPk As Long
                ' get list of all members of household
                MyCommon.QueryStr = "select CustomerPK from Customers with (NoLock) where HHPK=" & CustomerPK & ";"
                rst2 = MyCommon.LXS_Select
                If rst2.Rows.Count > 0 Then
                    For Each row In rst2.Rows
                        lCustomerPk = MyCommon.NZ(row.Item(0), 0)
                        MyCommon.QueryStr = "select ExtCardID from CardIDs with (NoLock) where CustomerPK=" & lCustomerPk & " and CardTypeID=0;"
                        dt = MyCommon.LXS_Select
                        If dt.Rows.Count > 0 Then
                            ' Not required to decrypt as its passed into SQL SP/Inline SQL
                            HHCustIdList.Add(MyCommon.NZ(dt.Rows(0).Item(0), ""))
                        End If
                    Next
                End If
            End If
        End If
    End If
  
    UnknownPhrase = Copient.PhraseLib.Lookup("term.unknown", LanguageID)
  
    If CardPK > 0 Then
        Send_HeadBegin("term.customer", "term.transactions", MyCommon.Extract_Val(ExtCardID))
  Else If ExtLocationCode <> ""
    Send_HeadBegin("term.transactions", "term.transactions", ExtLocationCode)
    Else
        Send_HeadBegin("term.customer", "term.transactions")
    End If
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    If restrictLinks Then
        Send("<link rel=""stylesheet"" href=""../css/logix-restricted.css"" type=""text/css"" media=""screen"" />")
    End If
    Send_Scripts()
    Send_HeadEnd()
  
    Send_BodyBegin(1)
    Send_Bar(Handheld)
    Send_Help(CopientFileName)
    Send_Logos()
    If (Not restrictLinks) Then
        Send_Tabs(Logix, 3)
        If CardPK > 0 Then
            Send_Subtabs(Logix, 32, 6, LanguageID, CustomerPK, , CardPK)
    ElseIf ExtLocationCode <> "" Then
      Send_Subtabs(Logix, 35, 6, LanguageID, ExtLocationCode)
        Else
            Send_Subtabs(Logix, 32, 6, LanguageID, CustomerPK)
        End If
    Else
        If CardPK > 0 Then
            Send_Subtabs(Logix, 91, 7, LanguageID, CustomerPK, extraLink, CardPK)
        Else
            Send_Subtabs(Logix, 91, 7, LanguageID, CustomerPK, extraLink)
        End If
    End If
  
    If (Logix.UserRoles.AccessCustomerInquiry = False) Then
        Send_Denied(1, "perm.customer-access")
        GoTo done
    End If
  
    If (HHCustIdList.Count > 0) Then
        CustExtIdList = "'" & MyCryptLib.SQL_StringEncrypt(ExtCardID) & "'"
        dCustExtIdList = "'" & ExtCardID & "'"
        For i = 0 To HHCustIdList.Count - 1
            ' HHCustIdList is already encrypted
            CustExtIdList += ", '" & MyCommon.Parse_Quotes(HHCustIdList.Item(i).ToString) & "'"
            dCustExtIdList += ", '" & MyCryptLib.SQL_StringDecrypt(HHCustIdList.Item(i).ToString) & "'"
        Next
    ElseIf (Request.QueryString("searchterms") <> "") Then
        dCustExtIdList = MyCommon.Pad_ExtCardID(ClientUserID1, Convert.ToInt32(CardTypeID))
        ClientUserID1 = MyCryptLib.SQL_StringEncrypt(MyCommon.Pad_ExtCardID(ClientUserID1, Convert.ToInt32(CardTypeID)))
        CustExtID = ClientUserID1
        CustExtIdList = "'" & ClientUserID1 & "'"
        dCustExtIdList = "'" & dCustExtIdList & "'"
    ElseIf (Request.QueryString("editterms") <> "") Then
        dCustExtIdList = "'" & MyCommon.Parse_Quotes(Request.QueryString("editterms"))& "'"
        CustExtID = MyCryptLib.SQL_StringEncrypt(MyCommon.Parse_Quotes(Request.QueryString("editterms")))
        CustExtIdList = "'" & CustExtID & "'"
    ElseIf ClientUserID1 <> "" Then
        dCustExtIdList = "'" & ClientUserID1 & "'"
        CustExtIdList = "'" & MyCryptLib.SQL_StringEncrypt(ClientUserID1) & "'"
    ElseIf CustomerPK > 0 Then
        ' We have a customer, but we haven't discovered any External IDs. If we have a customer,
        ' then we should use all external IDs for that customer.
        Dim get_ext_ids_query = "SELECT ExtCardID FROM CardIDs WHERE CustomerPK = " & CustomerPK
        MyCommon.QueryStr = get_ext_ids_query
        Dim query_result As DataTable = MyCommon.LXS_Select
        If query_result.Rows.Count = 0 Then
            ' Orphaned customer record?
            CustExtIdList = "''"
            dCustExtIdList = "''"
            MyCommon.Write_Log("logix_error_log.txt", "CustomerPK " & CustomerPK & " has no associated cards", True)
        Else
            CustExtIdList = "'" & query_result.Rows(0).Item(0) & "'"
            dCustExtIdList = "'" & MyCryptLib.SQL_StringDecrypt(query_result.Rows(0).Item(0).ToString()) & "'"
            For i = 1 To query_result.Rows.Count - 1
                'Dont have to encrypt as this is passed into SQL SP /SQL Inline
                CustExtIdList += ", '" & query_result.Rows(i).Item(0) & "'"
                dCustExtIdList += ", '" & MyCryptLib.SQL_StringDecrypt(query_result.Rows(i).Item(0).ToString()) & "'"
            Next
        End If
    Else
        CustExtIdList = "''"
        dCustExtIdList = "''"
    End If
  
    MyCommon.Open_LogixWH()
  
  If (sStartTime <> "") And (sEndTime <> "")  Then
      DateFilter = " having max(TH.TransDate) between '" & sStartTime & "'" & _
                     " and '" & sEndTime & "' "  
  End If
  
    ' Process search request, if applicable 
    TransTerms = Request.QueryString("transterms")
    If (TransTerms <> "") Then
        If (Date.TryParse(TransTerms, tempDate)) Then
            HavingFilterTR = " having Max(TR.TransDate) between '" & tempDate.ToString("yyyy-MM-ddT00:00:00") & "'" & _
                           " and '" & tempDate.ToString("yyyy-MM-ddT23:59:59") & "'"
            HavingFilterTH = " having Max(TH.TransDate) between '" & tempDate.ToString("yyyy-MM-ddT00:00:00") & "'" & _
                         " and '" & tempDate.ToString("yyyy-MM-ddT23:59:59") & "'"
            SearchFilter = ""
            SearchFilterTH = ""
        Else
            'SearchFilter = " and (TR.CustomerPrimaryExtID='" & MyCommon.Parse_Quotes(MyCryptLib.SQL_StringEncrypt(TransTerms)) & "'" & _
            SearchFilter = " and (TR.CustomerPrimaryExtID='" & MyCommon.Parse_Quotes(TransTerms) & "'" & _
                           " or TR.ExtLocationCode='" & MyCommon.Parse_Quotes(TransTerms) & "'" & _
                           " or TR.LogixTransNum='" & MyCommon.Parse_Quotes(TransTerms) & "'" & _
                           " or TR.TransNum='" & MyCommon.Parse_Quotes(TransTerms) & "'" & _
                           " or TR.TerminalNum='" & MyCommon.Parse_Quotes(TransTerms) & "')"
            'SearchFilterTH = " and (TH.CustomerPrimaryExtID='" & MyCommon.Parse_Quotes(MyCryptLib.SQL_StringEncrypt(TransTerms)) & "'" & _
            SearchFilterTH = " and (TH.CustomerPrimaryExtID='" & MyCommon.Parse_Quotes(TransTerms) & "'" & _
                       " or TH.ExtLocationCode='" & MyCommon.Parse_Quotes(TransTerms) & "'" & _
                       " or TH.LogixTransNum='" & MyCommon.Parse_Quotes(TransTerms) & "'" & _
                       " or POSTransNum='" & MyCommon.Parse_Quotes(TransTerms) & "'" & _
                       " or TH.TerminalNum='" & MyCommon.Parse_Quotes(TransTerms) & "')"
        End If
    Else
        SearchFilter = ""
        SearchFilterTH = ""
        HavingFilterTR = ""
        HavingFilterTH = ""
    End If
  
    'If
    If Not IsCMOnly AndAlso IsHouseholdID Then
        CPEHHFilter = "and HHID is NULL) OR HHID=" & dCustExtIdList & " "
    End If
  
    
    ' Process sort request, if applicable
    If (Request.QueryString("sortcol") <> "") Then SortCol = Request.QueryString("sortcol")
    If (Request.QueryString("sortdir") <> "") Then
        SortDir = Request.QueryString("sortdir")
    Else
        SortDir = "desc"
    End If
  
    Dim ShowPOSTimeStamp As Boolean = IIf(MyCommon.Fetch_CPE_SystemOption(131) = "1", True, False)
    Dim TransactionTimeQuery As String = "Max(TR.TransDate) as TransactionDate"
    Dim TransactionTimeQuery_TH As String = "Max(TH.TransDate) as TransactionDate"
    Dim TransactionTimeCondQuery As String = "Max(TR.TransDate) "
    Dim TransactionTimeCondQuery_TH As String = "Max(TransDate) "
    If ShowPOSTimeStamp Then
        TransactionTimeQuery = "Max(TR.POSTimeStamp) as TransactionDate"
        TransactionTimeQuery_TH = "Max(POSTimeStamp) as TransactionDate"
        TransactionTimeCondQuery = "Max(TR.POSTimeStamp) "
        TransactionTimeCondQuery_TH = "Max(POSTimeStamp) "
    End If
  
  
    If IsHouseholdID Then
        'HOUSEHOLD
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
        'Dont have to encrypt as this is passed into SQL SP /SQL Inline
        CustExtIdList = MyCommon.LXSsp.Parameters("@ExtCardIDs").Value
        
        MyCommon.Close_LXSsp()
        If (String.IsNullOrEmpty(CustExtIdList.Replace("'", ""))) Then
            CustExtIdList = "'" & MyCryptLib.SQL_StringEncrypt(ExtCardID.ToString()) & "'"
            dCustExtIdList = "'" & ExtCardID & "'"
        Else
            Dim CustListData() As String = CustExtIdList.Split(",")
            CustExtIdList = CustExtIdList & ",'" & MyCryptLib.SQL_StringEncrypt(ExtCardID.ToString()) & "'"
            For Each cData In CustListData
                Dim tmpExtId = cData.Replace("'", "").Replace(",", "").Trim
                dCustExtIdList = "'" & MyCryptLib.SQL_StringDecrypt(tmpExtId) & "',"
            Next
            dCustExtIdList = dCustExtIdList & "'" & ExtCardID & "'"
        End If
        'Get constituents' LogixTransNums
        MyCommon.QueryStr = "dbo.pt_Get_LogixTransNumList"
        MyCommon.Open_LXSsp()
        MyCommon.LXSsp.Parameters.Add("@HHPK", SqlDbType.Int).Value = CustomerPK
        MyCommon.LXSsp.Parameters.Add("@ExtCardIDs", SqlDbType.NVarChar).Value = CustExtIdList
        MyCommon.LXSsp.Parameters.Add("@LogixTransNums", SqlDbType.NVarChar, 4000).Direction = ParameterDirection.Output
        MyCommon.LXSsp.ExecuteNonQuery()
        LogixTransNums = MyCommon.LXSsp.Parameters("@LogixTransNums").Value
        MyCommon.Close_LXSsp()
        'Run the main query
        'TransContext 0 = Records where the individuals were previously and are currently members of this household, or records for the household itself.
        'TransContext 1 = Records where the individuals were previously members of this household, but no longer are. (CustomerID doesn't match a current member of the household, but HouseholdID does match the current household.)
        'TransContext 2 = Records where the individuals were previously members of a different household, but now are in the this one. (CustomerID matches a current member of the household, but HouseholdID doesn't match the current household.)
        Dim dtContext0 As DataTable
        Dim dtContext1 As DataTable
        Dim dtContext2 As DataTable
        If redemptionFilter = 1 Then
            MyCommon.QueryStr = "select " & IIf(ShowAll, "", "top 20 ") & " TR.CustomerPrimaryExtID, " & TransactionTimeQuery & ", TR.ExtLocationCode, " & _
                                "sum(TR.RedemptionAmount) as RedemptionAmount, sum(TR.RedemptionCount) as RedemptionCount, " & _
                                "TR.TerminalNum, TR.LogixTransNum, TR.TransNum, count(*) as DetailRecords, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.HHID, TR.Replayed, 0 AS TransContext, TH.TransTotal as TransTotal " & _
                                "from TransRedemptionView as TR with (NoLock) Left Outer Join  TransHist TH ON TR.LogixTransNum=TH.LogixTransNum " & _
                                "where (TR.HHID = '" & ExtCardID & "' AND TR.CustomerPrimaryExtID IN (" & dCustExtIdList & ") AND TR.CustomerTypeID<>1) or (TR.CustomerPrimaryExtID = '" & ExtCardID & "' AND TR.CustomerTypeID=1) " & _
                                "group by TR.CustomerPrimaryExtID, TR.HHID, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.LogixTransNum, TR.TransNum, TR.TerminalNum, TR.ExtLocationCode, TR.Replayed, TH.TransTotal " & _
              " UNION " & _
              "select " & IIf(ShowAll, "", "top 20 ") & " TR.CustomerPrimaryExtID, " & TransactionTimeQuery & ", TR.ExtLocationCode, " & _
                                "sum(TR.RedemptionAmount) as RedemptionAmount, sum(TR.RedemptionCount) as RedemptionCount, " & _
                                "TR.TerminalNum, TR.LogixTransNum, TR.TransNum, count(*) as DetailRecords, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.HHID, TR.Replayed, 0 AS TransContext, TH.TransTotal as TransTotal " & _
                                "from ThirdPartyTransRedemption as TR with (NoLock) Left Outer Join  ThirdPartyTransHist TH ON TR.LogixTransNum=TH.LogixTransNum " & _
                                "where (TR.HHID = '" & ExtCardID & "' AND TR.CustomerPrimaryExtID IN (" & dCustExtIdList & ") AND TR.CustomerTypeID<>1) or (TR.CustomerPrimaryExtID = '" & ExtCardID & "' AND TR.CustomerTypeID=1) " & _
                                "group by TR.CustomerPrimaryExtID, TR.HHID, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.LogixTransNum, TR.TransNum, TR.TerminalNum, TR.ExtLocationCode, TR.Replayed, TH.TransTotal " & _
                                "order by " & SortCol & " " & SortDir
            rstTemp = MyCommon.LWH_Select
            MyCommon.QueryStr = "select " & IIf(ShowAll, "", "top 20 ") & " TR.CustomerPrimaryExtID, " & TransactionTimeQuery & ", TR.ExtLocationCode, " & _
                                "sum(TR.RedemptionAmount) as RedemptionAmount, sum(TR.RedemptionCount) as RedemptionCount, " & _
                                "TR.TerminalNum, TR.LogixTransNum, TR.TransNum, count(*) as DetailRecords, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.HHID, TR.Replayed, 1 AS TransContext, TH.TransTotal as TransTotal " & _
                                "from TransRedemptionView as TR with (NoLock) Left Outer Join  TransHist TH ON TR.LogixTransNum=TH.LogixTransNum " & _
                                "where TR.HHID = '" & ExtCardID & "' AND TR.CustomerPrimaryExtID NOT IN (" & dCustExtIdList & ") " & _
                                "group by TR.CustomerPrimaryExtID, TR.HHID, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.LogixTransNum, TR.TransNum, TR.TerminalNum, TR.ExtLocationCode, TR.Replayed, TH.TransTotal " & _
              " UNION " & _
              "select " & IIf(ShowAll, "", "top 20 ") & " TR.CustomerPrimaryExtID, " & TransactionTimeQuery & ", TR.ExtLocationCode, " & _
                                "sum(TR.RedemptionAmount) as RedemptionAmount, sum(TR.RedemptionCount) as RedemptionCount, " & _
                                "TR.TerminalNum, TR.LogixTransNum, TR.TransNum, count(*) as DetailRecords, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.HHID, TR.Replayed, 1 AS TransContext, TH.TransTotal as TransTotal " & _
                                "from ThirdPartyTransRedemption as TR with (NoLock) Left Outer Join  ThirdPartyTransHist TH ON TR.LogixTransNum=TH.LogixTransNum " & _
                                "where TR.HHID = '" & ExtCardID & "' AND TR.CustomerPrimaryExtID NOT IN (" & dCustExtIdList & ") " & _
                                "group by TR.CustomerPrimaryExtID, TR.HHID, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.LogixTransNum, TR.TransNum, TR.TerminalNum, TR.ExtLocationCode, TR.Replayed, TH.TransTotal " & _
                                "order by " & SortCol & " " & SortDir
            dtContext1 = MyCommon.LWH_Select
            If dtContext1.Rows.Count > 0 Then
                MergeDataTables(dtContext1, rstTemp)
            End If
            MyCommon.QueryStr = "select " & IIf(ShowAll, "", "top 20 ") & " TR.CustomerPrimaryExtID, " & TransactionTimeQuery & ", TR.ExtLocationCode, " & _
                                "sum(TR.RedemptionAmount) as RedemptionAmount, sum(TR.RedemptionCount) as RedemptionCount, " & _
                                "TR.TerminalNum, TR.LogixTransNum, TR.TransNum, count(*) as DetailRecords, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.HHID, TR.Replayed, 2 AS TransContext, TH.TransTotal as TransTotal " & _
                                "from TransRedemptionView as TR with (NoLock) Left Outer Join  TransHist TH ON TR.LogixTransNum=TH.LogixTransNum " & _
                                "where ISNULL(TR.HHID, '') <> '" & ExtCardID & "' AND (TR.CustomerPrimaryExtID IN (" & dCustExtIdList & ") and TR.CustomerTypeID<>1) " & _
                                "group by TR.CustomerPrimaryExtID, TR.HHID, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.LogixTransNum, TR.TransNum, TR.TerminalNum, TR.ExtLocationCode, TR.Replayed, TH.TransTotal " & _
              " UNION " & _
              "select " & IIf(ShowAll, "", "top 20 ") & " TR.CustomerPrimaryExtID, " & TransactionTimeQuery & ", TR.ExtLocationCode, " & _
                                "sum(TR.RedemptionAmount) as RedemptionAmount, sum(TR.RedemptionCount) as RedemptionCount, " & _
                                "TR.TerminalNum, TR.LogixTransNum, TR.TransNum, count(*) as DetailRecords, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.HHID, TR.Replayed, 2 AS TransContext, TH.TransTotal as TransTotal " & _
                                "from ThirdPartyTransRedemption as TR with (NoLock) Left Outer Join  ThirdPartyTransHist TH ON TR.LogixTransNum=TH.LogixTransNum " & _
                                "where ISNULL(TR.HHID, '') <> '" & ExtCardID & "' AND (TR.CustomerPrimaryExtID IN (" & dCustExtIdList & ") and TR.CustomerTypeID<>1) " & _
                                "group by TR.CustomerPrimaryExtID, TR.HHID, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.LogixTransNum, TR.TransNum, TR.TerminalNum, TR.ExtLocationCode, TR.Replayed, TH.TransTotal " & _
                                "order by " & SortCol & " " & SortDir
            dtContext2 = MyCommon.LWH_Select
            If dtContext2.Rows.Count > 0 Then
                MergeDataTables(dtContext2, rstTemp)
            End If
            
           
            
            MyCommon.QueryStr = "select " & IIf(ShowAll, "", "top 20 ") & "CustomerPrimaryExtID, " & TransactionTimeQuery_TH & ", ExtLocationCode, 0 as RedemptionAmount, 0 as RedemptionCount, " & _
                                "TerminalNum, LogixTransNum, POSTransNum as TransNum, 0 as DetailRecords, CustomerTypeID, PresentedCustomerID, PresentedCardTypeID, HHID, Replayed, 0 AS TransContext,isnull(TransTotal,0) as TransTotal  " & _
                                "from TransHist as TH with (NoLock) " & _
                                "where " & IIf(CPEHHFilter <> "", "(", "") & "((CustomerPrimaryExtID in (" & dCustExtIdList & ") " & CPEHHFilter & " and CustomerTypeID<>1) or (CustomerPrimaryExtID = '" & ExtCardID & "' AND CustomerTypeID=1)) " & SearchFilterTH & " and not exists " & _
                                "  (select LogixTransNum from TransRedemptionView as TR2 where " & IIf(CPEHHFilter <> "", "(", "") & "((CustomerPrimaryExtID in (" & dCustExtIdList & ") " & CPEHHFilter & " and CustomerTypeID<>1) or (CustomerPrimaryExtID = '" & ExtCardID & "' AND CustomerTypeID=1)) and TH.LogixTransNum=TR2.LogixTransNum) " & _
                                "group by CustomerPrimaryExtID, HHID, CustomerTypeID, PresentedCustomerID, PresentedCardTypeID, LogixTransNum, POSTransNum, TerminalNum, ExtLocationCode, Replayed,TransTotal " & HavingFilterTH & " " & _
              " UNION " & _
              "select " & IIf(ShowAll, "", "top 20 ") & "CustomerPrimaryExtID, " & TransactionTimeQuery_TH & ", ExtLocationCode, 0 as RedemptionAmount, 0 as RedemptionCount, " & _
                                "TerminalNum, LogixTransNum, POSTransNum as TransNum, 0 as DetailRecords, CustomerTypeID, PresentedCustomerID, PresentedCardTypeID, HHID, Replayed, 0 AS TransContext,isnull(TransTotal,0) as TransTotal  " & _
                                "from ThirdPartyTransHist as TH with (NoLock) " & _
                                "where " & IIf(CPEHHFilter <> "", "(", "") & "((CustomerPrimaryExtID in (" & dCustExtIdList & ") " & CPEHHFilter & " and CustomerTypeID<>1) or (CustomerPrimaryExtID = '" & ExtCardID & "' AND CustomerTypeID=1)) " & SearchFilterTH & " and not exists " & _
                                "  (select LogixTransNum from ThirdPartyTransRedemption as TR2 where " & IIf(CPEHHFilter <> "", "(", "") & "((CustomerPrimaryExtID in (" & dCustExtIdList & ") " & CPEHHFilter & " and CustomerTypeID<>1) or (CustomerPrimaryExtID = '" & ExtCardID & "' AND CustomerTypeID=1)) and TH.LogixTransNum=TR2.LogixTransNum) " & _
                                "group by CustomerPrimaryExtID, HHID, CustomerTypeID, PresentedCustomerID, PresentedCardTypeID, LogixTransNum, POSTransNum, TerminalNum, ExtLocationCode, Replayed,TransTotal " & HavingFilterTH & " " & _
                                "order by " & SortCol & " " & SortDir
            dtContext0 = MyCommon.LWH_Select
            If dtContext0.Rows.Count > 0 Then
                MergeDataTables(dtContext0, rstTemp)
            End If
            rstTrans = SelectIntoDataTable("", SortCol & " " & SortDir, ShowAll, rstTemp)
        ElseIf redemptionFilter = 2 Then
            MyCommon.QueryStr = "select CustomerPrimaryExtID, TransactionDate, ExtLocationCode, RedemptionAmount, RedemptionCount, TerminalNum, LogixTransNum, TransNum, DetailRecords, CustomerTypeID, PresentedCustomerID, PresentedCardTypeID, HHID, Replayed, TransContext, TransTotal " & _
                                "from (select " & IIf(ShowAll, "", "top 20 ") & " TR.CustomerPrimaryExtID, " & TransactionTimeQuery & ", TR.ExtLocationCode, " & _
                                "sum(TR.RedemptionAmount) as RedemptionAmount, sum(TR.RedemptionCount) as RedemptionCount, " & _
                                "TR.TerminalNum, TR.LogixTransNum, TR.TransNum, count(*) as DetailRecords, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.HHID, TR.Replayed, 0 AS TransContext, TH.TransTotal as TransTotal " & _
                                "from TransRedemptionView as TR with (NoLock) Left Outer Join  TransHist TH ON TR.LogixTransNum=TH.LogixTransNum " & _
                                "where (TR.HHID = '" & ExtCardID & "' AND TR.CustomerPrimaryExtID IN (" & dCustExtIdList & ") AND TR.CustomerTypeID<>1) or (TR.CustomerPrimaryExtID = '" & ExtCardID & "' AND TR.CustomerTypeID=1) " & _
                                "group by TR.CustomerPrimaryExtID, TR.HHID, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.LogixTransNum, TR.TransNum, TR.TerminalNum, TR.ExtLocationCode, TR.Replayed, TH.TransTotal" & _
                                " " & IIf(ShowAll, "", "order by " & SortCol & " " & SortDir) & ") AS LogixRedemption " & _
                                " UNION " & _
                                            "select CustomerPrimaryExtID, TransactionDate, ExtLocationCode, RedemptionAmount, RedemptionCount, TerminalNum, LogixTransNum, TransNum, DetailRecords, CustomerTypeID, PresentedCustomerID, PresentedCardTypeID, HHID, Replayed, TransContext, TransTotal " & _
                                "from (select " & IIf(ShowAll, "", "top 20 ") & " TR.CustomerPrimaryExtID, " & TransactionTimeQuery & ", TR.ExtLocationCode, " & _
                                "sum(TR.RedemptionAmount) as RedemptionAmount, sum(TR.RedemptionCount) as RedemptionCount, " & _
                                "TR.TerminalNum, TR.LogixTransNum, TR.TransNum, count(*) as DetailRecords, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.HHID, TR.Replayed, 0 AS TransContext, TH.TransTotal as TransTotal " & _
                                "from ThirdPartyTransRedemption as TR with (NoLock) Left Outer Join  ThirdPartyTransHist TH ON TR.LogixTransNum=TH.LogixTransNum " & _
                                "where (TR.HHID = '" & ExtCardID & "' AND TR.CustomerPrimaryExtID IN (" & dCustExtIdList & ") AND TR.CustomerTypeID<>1) or (TR.CustomerPrimaryExtID = '" & ExtCardID & "' AND TR.CustomerTypeID=1) " & _
                                "group by TR.CustomerPrimaryExtID, TR.HHID, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.LogixTransNum, TR.TransNum, TR.TerminalNum, TR.ExtLocationCode, TR.Replayed, TH.TransTotal" & _
                                " " & IIf(ShowAll, "", "order by " & SortCol & " " & SortDir) & ") AS ThirdPartyRedemption " & _
                                "order by " & SortCol & " " & SortDir
      rstTemp = MyCommon.LWH_Select
            MyCommon.QueryStr = "select CustomerPrimaryExtID, TransactionDate, ExtLocationCode, RedemptionAmount, RedemptionCount, TerminalNum, LogixTransNum, TransNum, DetailRecords, CustomerTypeID, PresentedCustomerID, PresentedCardTypeID, HHID, TransContext, TransTotal " & _
                                "from (select " & IIf(ShowAll, "", "top 20 ") & " TR.CustomerPrimaryExtID, " & TransactionTimeQuery & ", TR.ExtLocationCode, " & _
                                "sum(TR.RedemptionAmount) as RedemptionAmount, sum(TR.RedemptionCount) as RedemptionCount, " & _
                                "TR.TerminalNum, TR.LogixTransNum, TR.TransNum, count(*) as DetailRecords, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.HHID, TR.Replayed, 1 AS TransContext, TH.TransTotal as TransTotal " & _
                                "from TransRedemptionView as TR with (NoLock) Left Outer Join  TransHist TH ON TR.LogixTransNum=TH.LogixTransNum " & _
                                "where TR.HHID = '" & ExtCardID & "' AND TR.CustomerPrimaryExtID NOT IN (" & dCustExtIdList & ") " & _
                                "group by TR.CustomerPrimaryExtID, TR.HHID, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.LogixTransNum, TR.TransNum, TR.TerminalNum, TR.ExtLocationCode, TR.Replayed, TH.TransTotal " & _
                                " " & IIf(ShowAll, "", "order by " & SortCol & " " & SortDir) & ") AS LogixRedemption " & _
                                " UNION " & _
                                            "select CustomerPrimaryExtID, TransactionDate, ExtLocationCode, RedemptionAmount, RedemptionCount, TerminalNum, LogixTransNum, TransNum, DetailRecords, CustomerTypeID, PresentedCustomerID, PresentedCardTypeID, HHID, TransContext, TransTotal " & _
                                "from (select " & IIf(ShowAll, "", "top 20 ") & " TR.CustomerPrimaryExtID, " & TransactionTimeQuery & ", TR.ExtLocationCode, " & _
                                "sum(TR.RedemptionAmount) as RedemptionAmount, sum(TR.RedemptionCount) as RedemptionCount, " & _
                                "TR.TerminalNum, TR.LogixTransNum, TR.TransNum, count(*) as DetailRecords, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.HHID, TR.Replayed, 1 AS TransContext, TH.TransTotal as TransTotal " & _
                                "from ThirdPartyTransRedemption as TR with (NoLock) Left Outer Join  ThirdPartyTransHist TH ON TR.LogixTransNum=TH.LogixTransNum " & _
                                "where TR.HHID = '" & ExtCardID & "' AND TR.CustomerPrimaryExtID NOT IN (" & dCustExtIdList & ") " & _
                                "group by TR.CustomerPrimaryExtID, TR.HHID, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.LogixTransNum, TR.TransNum, TR.TerminalNum, TR.ExtLocationCode, TR.Replayed, TH.TransTotal " & _
                                " " & IIf(ShowAll, "", "order by " & SortCol & " " & SortDir) & ") AS ThirdPartyRedemption " & _
                                "order by " & SortCol & " " & SortDir
      dtContext1 = MyCommon.LWH_Select
      If dtContext1.Rows.Count > 0 Then
        MergeDataTables(dtContext1, rstTemp)
      End If
            MyCommon.QueryStr = "select CustomerPrimaryExtID, TransactionDate, ExtLocationCode, RedemptionAmount, RedemptionCount, TerminalNum, LogixTransNum, TransNum, DetailRecords, CustomerTypeID, PresentedCustomerID, PresentedCardTypeID, HHID, TransContext, TransTotal " & _
                                "from (select " & IIf(ShowAll, "", "top 20 ") & " TR.CustomerPrimaryExtID, " & TransactionTimeQuery & ", TR.ExtLocationCode, " & _
                                "sum(TR.RedemptionAmount) as RedemptionAmount, sum(TR.RedemptionCount) as RedemptionCount, " & _
                                "TR.TerminalNum, TR.LogixTransNum, TR.TransNum, count(*) as DetailRecords, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.HHID, TR.Replayed, 2 AS TransContext, TH.TransTotal as TransTotal " & _
                                "from TransRedemptionView as TR with (NoLock) Left Outer Join  TransHist TH ON TR.LogixTransNum=TH.LogixTransNum " & _
                                "where ISNULL(TR.HHID, '') <> '" & ExtCardID & "' AND (TR.CustomerPrimaryExtID IN (" & dCustExtIdList & ") and TR.CustomerTypeID<>1) " & _
                                "group by TR.CustomerPrimaryExtID, TR.HHID, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.LogixTransNum, TR.TransNum, TR.TerminalNum, TR.ExtLocationCode, TR.Replayed, TH.TransTotal " & _
                                " " & IIf(ShowAll, "", "order by " & SortCol & " " & SortDir) & ") AS LogixRedemption " & _
                                " UNION " & _
                                            "select CustomerPrimaryExtID, TransactionDate, ExtLocationCode, RedemptionAmount, RedemptionCount, TerminalNum, LogixTransNum, TransNum, DetailRecords, CustomerTypeID, PresentedCustomerID, PresentedCardTypeID, HHID, TransContext, TransTotal " & _
                                "from (select " & IIf(ShowAll, "", "top 20 ") & " TR.CustomerPrimaryExtID, " & TransactionTimeQuery & ", TR.ExtLocationCode, " & _
                                "sum(TR.RedemptionAmount) as RedemptionAmount, sum(TR.RedemptionCount) as RedemptionCount, " & _
                                "TR.TerminalNum, TR.LogixTransNum, TR.TransNum, count(*) as DetailRecords, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.HHID, TR.Replayed, 2 AS TransContext, TH.TransTotal as TransTotal " & _
                                "from ThirdPartyTransRedemption as TR with (NoLock) Left Outer Join  ThirdPartyTransHist TH ON TR.LogixTransNum=TH.LogixTransNum " & _
                                "where ISNULL(TR.HHID, '') <> '" & ExtCardID & "' AND (TR.CustomerPrimaryExtID IN (" & dCustExtIdList & ") and TR.CustomerTypeID<>1) " & _
                                "group by TR.CustomerPrimaryExtID, TR.HHID, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.LogixTransNum, TR.TransNum, TR.TerminalNum, TR.ExtLocationCode, TR.Replayed, TH.TransTotal " & _
                                " " & IIf(ShowAll, "", "order by " & SortCol & " " & SortDir) & ") AS ThirdPartyRedemption " & _
                                "order by " & SortCol & " " & SortDir
            dtContext2 = MyCommon.LWH_Select
            If dtContext2.Rows.Count > 0 Then
                MergeDataTables(dtContext2, rstTemp)
            End If
            rstTrans = SelectIntoDataTable("", SortCol & " " & SortDir, ShowAll, rstTemp)
        ElseIf redemptionFilter = 3 Then
            MyCommon.QueryStr = "select CustomerPrimaryExtID, " & TransactionTimeQuery_TH & ", ExtLocationCode, 0 as RedemptionAmount, 0 as RedemptionCount, " & _
                                "TerminalNum, LogixTransNum, POSTransNum as TransNum, 0 as DetailRecords, CustomerTypeID, PresentedCustomerID, PresentedCardTypeID, HHID, Replayed, 0 AS TransContext,isnull(TransTotal,0) as TransTotal " & _
                                "from TransHist as TH with (NoLock) " & _
                                "where " & IIf(CPEHHFilter <> "", "(", "") & "(CustomerPrimaryExtID in (" & dCustExtIdList & ") " & CPEHHFilter & " OR (CustomerPrimaryExtID='" & ExtCardID & "' and CustomerTypeID=1)) " & SearchFilterTH & " and not exists " & _
                                "  (select LogixTransNum from TransRedemptionView as TR where " & IIf(CPEHHFilter <> "", "(", "") & "(CustomerPrimaryExtID in (" & dCustExtIdList & ") " & CPEHHFilter & " OR (CustomerPrimaryExtID='" & ExtCardID & "' and CustomerTypeID=1)) and TH.LogixTransNum=TR.LogixTransNum) " & _
                                "group by CustomerPrimaryExtID, HHID, CustomerTypeID, PresentedCustomerID, PresentedCardTypeID, LogixTransNum, POSTransNum, TerminalNum, ExtLocationCode, Replayed,TransTotal " & HavingFilterTH & " " & _
              " UNION " & _
              "select CustomerPrimaryExtID, " & TransactionTimeQuery_TH & ", ExtLocationCode, 0 as RedemptionAmount, 0 as RedemptionCount, " & _
                                "TerminalNum, LogixTransNum, POSTransNum as TransNum, 0 as DetailRecords, CustomerTypeID, PresentedCustomerID, PresentedCardTypeID, HHID, Replayed, 0 AS TransContext,isnull(TransTotal,0) as TransTotal " & _
                                "from ThirdPartyTransHist as TH with (NoLock) " & _
                                "where " & IIf(CPEHHFilter <> "", "(", "") & "(CustomerPrimaryExtID in (" & dCustExtIdList & ") " & CPEHHFilter & " OR (CustomerPrimaryExtID='" & ExtCardID & "' and CustomerTypeID=1)) " & SearchFilterTH & " and not exists " & _
                                "  (select LogixTransNum from ThirdPartyTransRedemption as TR where " & IIf(CPEHHFilter <> "", "(", "") & "(CustomerPrimaryExtID in (" & dCustExtIdList & ") " & CPEHHFilter & " OR (CustomerPrimaryExtID='" & ExtCardID & "' and CustomerTypeID=1)) and TH.LogixTransNum=TR.LogixTransNum) " & _
                                "group by CustomerPrimaryExtID, HHID, CustomerTypeID, PresentedCustomerID, PresentedCardTypeID, LogixTransNum, POSTransNum, TerminalNum, ExtLocationCode, Replayed,TransTotal " & HavingFilterTH & " " & _
                                "order by " & SortCol & " " & SortDir
            rstTrans = MyCommon.LWH_Select
        ElseIf redemptionFilter = 4 Then
		  Dim NoOfdays1 As Integer = NoOfdays - 1 
          Dim startDate As String = Now.AddDays(-NoOfdays1).ToShortDateString() & " 00:00:00"
          Dim endDate As String = Now.ToShortDateString() & " 23:59:59"

          Dim timeCondition As String = " Having " & TransactionTimeCondQuery & " between '" & startDate & "' and '" & endDate & "' "
          Dim timeCondition_TH As String = " Having " & TransactionTimeCondQuery_TH & " between '" & startDate & "' and '" & endDate & "' "

            MyCommon.QueryStr = "select " & IIf(ShowAll, "", "top 20 ") & " TR.CustomerPrimaryExtID, " & TransactionTimeQuery & ", TR.ExtLocationCode, " & _
                           "sum(TR.RedemptionAmount) as RedemptionAmount, sum(TR.RedemptionCount) as RedemptionCount, " & _
                           "TR.TerminalNum, TR.LogixTransNum, TR.TransNum, count(*) as DetailRecords, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.HHID, TR.Replayed, 0 AS TransContext, TH.TransTotal as TransTotal " & _
                           "from TransRedemptionView as TR with (NoLock) Left Outer Join  TransHist TH ON TR.LogixTransNum=TH.LogixTransNum " & _
                           "where (TR.HHID = '" & ExtCardID & "' AND TR.CustomerPrimaryExtID IN (" & dCustExtIdList & ") AND TR.CustomerTypeID<>1) or (TR.CustomerPrimaryExtID = '" & ExtCardID & "' AND TR.CustomerTypeID=1) " & _
                           "group by TR.CustomerPrimaryExtID, TR.HHID, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.LogixTransNum, TR.TransNum, TR.TerminalNum, TR.ExtLocationCode, TR.Replayed, TH.TransTotal " & _
                            timeCondition & _
         " UNION " & _
         "select " & IIf(ShowAll, "", "top 20 ") & " TR.CustomerPrimaryExtID, " & TransactionTimeQuery & ", TR.ExtLocationCode, " & _
                           "sum(TR.RedemptionAmount) as RedemptionAmount, sum(TR.RedemptionCount) as RedemptionCount, " & _
                           "TR.TerminalNum, TR.LogixTransNum, TR.TransNum, count(*) as DetailRecords, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.HHID, TR.Replayed, 0 AS TransContext, TH.TransTotal as TransTotal " & _
                           "from ThirdPartyTransRedemption as TR with (NoLock) Left Outer Join  ThirdPartyTransHist TH ON TR.LogixTransNum=TH.LogixTransNum " & _
                           "where (TR.HHID = '" & ExtCardID & "' AND TR.CustomerPrimaryExtID IN (" & dCustExtIdList & ") AND TR.CustomerTypeID<>1) or (TR.CustomerPrimaryExtID = '" & ExtCardID & "' AND TR.CustomerTypeID=1) " & _
                           "group by TR.CustomerPrimaryExtID, TR.HHID, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.LogixTransNum, TR.TransNum, TR.TerminalNum, TR.ExtLocationCode, TR.Replayed, TH.TransTotal " & _
                            timeCondition & _
                           "order by " & SortCol & " " & SortDir
          rstTemp = MyCommon.LWH_Select
            MyCommon.QueryStr = "select " & IIf(ShowAll, "", "top 20 ") & " TR.CustomerPrimaryExtID, " & TransactionTimeQuery & ", TR.ExtLocationCode, " & _
                                "sum(TR.RedemptionAmount) as RedemptionAmount, sum(TR.RedemptionCount) as RedemptionCount, " & _
                                "TR.TerminalNum, TR.LogixTransNum, TR.TransNum, count(*) as DetailRecords, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.HHID, TR.Replayed, 1 AS TransContext, TH.TransTotal as TransTotal " & _
                                "from TransRedemptionView as TR with (NoLock) Left Outer Join  TransHist TH ON TR.LogixTransNum=TH.LogixTransNum " & _
                                "where TR.HHID = '" & ExtCardID & "' AND TR.CustomerPrimaryExtID NOT IN (" & dCustExtIdList & ") " & _
                                "group by TR.CustomerPrimaryExtID, TR.HHID, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.LogixTransNum, TR.TransNum, TR.TerminalNum, TR.ExtLocationCode, TR.Replayed, TH.TransTotal " & _
                                 timeCondition & _
              " UNION " & _
              "select " & IIf(ShowAll, "", "top 20 ") & " TR.CustomerPrimaryExtID, " & TransactionTimeQuery & ", TR.ExtLocationCode, " & _
                                "sum(TR.RedemptionAmount) as RedemptionAmount, sum(TR.RedemptionCount) as RedemptionCount, " & _
                                "TR.TerminalNum, TR.LogixTransNum, TR.TransNum, count(*) as DetailRecords, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.HHID, TR.Replayed, 1 AS TransContext, TH.TransTotal as TransTotal " & _
                                "from ThirdPartyTransRedemption as TR with (NoLock) Left Outer Join  ThirdPartyTransHist TH ON TR.LogixTransNum=TH.LogixTransNum " & _
                                "where TR.HHID = '" & ExtCardID & "' AND TR.CustomerPrimaryExtID NOT IN (" & dCustExtIdList & ") " & _
                                "group by TR.CustomerPrimaryExtID, TR.HHID, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.LogixTransNum, TR.TransNum, TR.TerminalNum, TR.ExtLocationCode, TR.Replayed, TH.TransTotal " & _
                                timeCondition & _
                                "order by " & SortCol & " " & SortDir
          dtContext1 = MyCommon.LWH_Select
          If dtContext1.Rows.Count > 0 Then
              MergeDataTables(dtContext1, rstTemp)
          End If
            MyCommon.QueryStr = "select " & IIf(ShowAll, "", "top 20 ") & " TR.CustomerPrimaryExtID, " & TransactionTimeQuery & ", TR.ExtLocationCode, " & _
                                "sum(TR.RedemptionAmount) as RedemptionAmount, sum(TR.RedemptionCount) as RedemptionCount, " & _
                                "TR.TerminalNum, TR.LogixTransNum, TR.TransNum, count(*) as DetailRecords, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.HHID, TR.Replayed, 2 AS TransContext, TH.TransTotal as TransTotal " & _
                                "from TransRedemptionView as TR with (NoLock) Left Outer Join  TransHist TH ON TR.LogixTransNum=TH.LogixTransNum " & _
                                "where ISNULL(TR.HHID, '') <> '" & ExtCardID & "' AND (TR.CustomerPrimaryExtID IN (" & dCustExtIdList & ") and TR.CustomerTypeID<>1) " & _
                                "group by TR.CustomerPrimaryExtID, TR.HHID, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.LogixTransNum, TR.TransNum, TR.TerminalNum, TR.ExtLocationCode, TR.Replayed, TH.TransTotal " & _
                                timeCondition & _
              " UNION " & _
              "select " & IIf(ShowAll, "", "top 20 ") & " TR.CustomerPrimaryExtID, " & TransactionTimeQuery & ", TR.ExtLocationCode, " & _
                                "sum(TR.RedemptionAmount) as RedemptionAmount, sum(TR.RedemptionCount) as RedemptionCount, " & _
                                "TR.TerminalNum, TR.LogixTransNum, TR.TransNum, count(*) as DetailRecords, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.HHID, TR.Replayed, 2 AS TransContext, TH.TransTotal as TransTotal " & _
                                "from ThirdPartyTransRedemption as TR with (NoLock) Left Outer Join  ThirdPartyTransHist TH ON TR.LogixTransNum=TH.LogixTransNum " & _
                                "where ISNULL(TR.HHID, '') <> '" & ExtCardID & "' AND (TR.CustomerPrimaryExtID IN (" & dCustExtIdList & ") and TR.CustomerTypeID<>1) " & _
                                "group by TR.CustomerPrimaryExtID, TR.HHID, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.LogixTransNum, TR.TransNum, TR.TerminalNum, TR.ExtLocationCode, TR.Replayed, TH.TransTotal " & _
                                timeCondition & _
                                "order by " & SortCol & " " & SortDir
          dtContext2 = MyCommon.LWH_Select
          If dtContext2.Rows.Count > 0 Then
              MergeDataTables(dtContext2, rstTemp)
          End If
            MyCommon.QueryStr = "select " & IIf(ShowAll, "", "top 20 ") & "CustomerPrimaryExtID, " & TransactionTimeQuery_TH & ", ExtLocationCode, 0 as RedemptionAmount, 0 as RedemptionCount, " & _
                                "TerminalNum, LogixTransNum, POSTransNum as TransNum, 0 as DetailRecords, CustomerTypeID, PresentedCustomerID, PresentedCardTypeID, HHID, Replayed, 0 AS TransContext,isnull(TransTotal,0) as TransTotal  " & _
                                "from TransHist as TH with (NoLock) " & _
                                "where " & IIf(CPEHHFilter <> "", "(", "") & "((CustomerPrimaryExtID in (" & dCustExtIdList & ") " & CPEHHFilter & " and CustomerTypeID<>1) or (CustomerPrimaryExtID = '" & ExtCardID & "' AND CustomerTypeID=1)) " & SearchFilterTH & " and not exists " & _
                                "  (select LogixTransNum from TransRedemptionView as TR2 where " & IIf(CPEHHFilter <> "", "(", "") & "((CustomerPrimaryExtID in (" & dCustExtIdList & ") " & CPEHHFilter & " and CustomerTypeID<>1) or (CustomerPrimaryExtID = '" & ExtCardID & "' AND CustomerTypeID=1)) and TH.LogixTransNum=TR2.LogixTransNum) " & _
                                      "group by CustomerPrimaryExtID, HHID, CustomerTypeID, PresentedCustomerID, PresentedCardTypeID, LogixTransNum, POSTransNum, TerminalNum, ExtLocationCode, Replayed,TransTotal " & HavingFilterTH & " " & _
                                timeCondition_TH & _
              " UNION " & _
              "select " & IIf(ShowAll, "", "top 20 ") & "CustomerPrimaryExtID, " & TransactionTimeQuery_TH & ", ExtLocationCode, 0 as RedemptionAmount, 0 as RedemptionCount, " & _
                                "TerminalNum, LogixTransNum, POSTransNum as TransNum, 0 as DetailRecords, CustomerTypeID, PresentedCustomerID, PresentedCardTypeID, HHID, Replayed, 0 AS TransContext,isnull(TransTotal,0) as TransTotal  " & _
                                "from ThirdPartyTransHist as TH with (NoLock) " & _
                                "where " & IIf(CPEHHFilter <> "", "(", "") & "((CustomerPrimaryExtID in (" & dCustExtIdList & ") " & CPEHHFilter & " and CustomerTypeID<>1) or (CustomerPrimaryExtID = '" & ExtCardID & "' AND CustomerTypeID=1)) " & SearchFilterTH & " and not exists " & _
                                "  (select LogixTransNum from ThirdPartyTransRedemption as TR2 where " & IIf(CPEHHFilter <> "", "(", "") & "((CustomerPrimaryExtID in (" & dCustExtIdList & ") " & CPEHHFilter & " and CustomerTypeID<>1) or (CustomerPrimaryExtID = '" & ExtCardID & "' AND CustomerTypeID=1)) and TH.LogixTransNum=TR2.LogixTransNum) " & _
                                      "group by CustomerPrimaryExtID, HHID, CustomerTypeID, PresentedCustomerID, PresentedCardTypeID, LogixTransNum, POSTransNum, TerminalNum, ExtLocationCode, Replayed,TransTotal " & HavingFilterTH & " " & _
                              timeCondition_TH & _
                                "order by " & SortCol & " " & SortDir
          dtContext0 = MyCommon.LWH_Select
          If dtContext0.Rows.Count > 0 Then
              MergeDataTables(dtContext0, rstTemp)
          End If
          rstTrans = SelectIntoDataTable("", SortCol & " " & SortDir, ShowAll, rstTemp)

        End If
  Else If ExtLocationCode <> "" Then
    Dim tempstr As String = ""
    If sLast4CardID <> "" Then
            'tempstr = " And (TH.PresentedCustomerID = '" & MyCryptLib.SQL_StringEncrypt(sLast4CardID) & "') "
            tempstr = " And (TH.PresentedCustomerID = '" & sLast4CardID & "') "
    End If
    
    MyCommon.QueryStr = "dbo.pc_Transaction_Select"
    MyCommon.Open_LWHsp()
    MyCommon.LWHsp.Parameters.Add("@ExtLocationCode", SqlDbType.NVarChar, 20).Value = ExtLocationCode
    MyCommon.LWHsp.Parameters.Add("@Max", SqlDbType.Int).Value = IIF(ShowAll, 100, 20)
    MyCommon.LWHsp.Parameters.Add("@SearchFilter", SqlDbType.NVarChar, 1000).Value = IIF(SearchFilterTH <> "", SearchFilterTH, tempstr) 
    MyCommon.LWHsp.Parameters.Add("@DateFilter", SqlDbType.NVarChar, 1000).Value = IIF(HavingFilterTH <> "", HavingFilterTH, DateFilter)
    MyCommon.LWHsp.Parameters.Add("@RedemptionFilter", SqlDbType.Int).Value = redemptionFilter
    MyCommon.LWHsp.Parameters.Add("@SortCol", SqlDbType.NVarChar, 20).Value = SortCol
    MyCommon.LWHsp.Parameters.Add("@SortDir", SqlDbType.NVarChar, 4).Value = SortDir
    rstTrans = MyCommon.LWHsp_select()
    MyCommon.Close_LXSsp()
    Else
        'CUSTOMER
		
'=============================================================================	CLOUDSOL-1252    
	Dim IACSearchTerms As String = Request.QueryString("searchterms")
    If (Request.QueryString("chkIncludeAllCards") <> "") Then
      IncludeAllCards = Convert.ToInt32(Request.QueryString("chkIncludeAllCards"))
    End If 
    If IncludeAllCards = 1 And IACSearchTerms = "" Then	'...Get constituents ExtCardIDs
      MyCommon.QueryStr = "dbo.pt_Get_ExtCardIDList"
      MyCommon.Open_LXSsp()
      MyCommon.LXSsp.Parameters.Add("@CustPKs", SqlDbType.NVarChar, 4000).Value = CustomerPK
	  MyCommon.LXSsp.Parameters.Add("@IncludeAltIDType3Cards", SqlDbType.Int).Value = 1
      MyCommon.LXSsp.Parameters.Add("@ExtCardIDs", SqlDbType.NVarChar, 4000).Direction = ParameterDirection.Output
      MyCommon.LXSsp.ExecuteNonQuery()
      CustExtIdList = MyCommon.LXSsp.Parameters("@ExtCardIDs").Value
      MyCommon.Close_LXSsp()
	  If (String.IsNullOrEmpty(CustExtIdList)) Then
                CustExtIdList = "'" & MyCryptLib.SQL_StringEncrypt(ExtCardID.ToString()) & "'"
                dCustExtIdList = "'" & ExtCardID.ToString() & "'"
      Else
                CustExtIdList = CustExtIdList & ",'" & MyCryptLib.SQL_StringEncrypt(ExtCardID.ToString()) & "'"
                Dim CustListData() As String = CustExtIdList.Split(",")
                For Each cData In CustListData
                    Dim tmpExtId = cData.Replace("'", "").Replace(",", "").Trim
                    dCustExtIdList = "'" & MyCryptLib.SQL_StringDecrypt(tmpExtId) & "',"
                Next
                dCustExtIdList = dCustExtIdList & "'" & ExtCardID & "'"
      End If	
    Else
      If IACSearchTerms <> "" Then
                CustExtIdList = MyCryptLib.SQL_StringEncrypt(IACSearchTerms)
                dCustExtIdList = IACSearchTerms
      End If
    End If
'=============================================================================	CLOUDSOL-1252 		
		
        If redemptionFilter = 1 Then 'get all transactions with redemptions (from TransRedemption view) and all transactions
            'without redemptions (from TransHist view, minus the transaction #s from TransRedemptionView),
            'and union them together.
            MyCommon.QueryStr = "select TR.CustomerPrimaryExtID, " & TransactionTimeQuery & ", TR.ExtLocationCode, sum(TR.RedemptionAmount) as RedemptionAmount, sum(TR.RedemptionCount) as RedemptionCount, " & _
                                "TR.TerminalNum, TR.LogixTransNum, TR.TransNum, count(*) as DetailRecords, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.HHID, TR.Replayed, 0 AS TransContext,TH.TransTotal as TransTotal " & _
                                "from TransRedemptionView as TR with (NoLock) Left Outer Join  TransHist TH ON TR.LogixTransNum = TH.LogixTransNum " & _
                                "where TR.CustomerTypeID in (0,1) and " & IIf(CPEHHFilter <> "", "(", "") & "(TR.CustomerPrimaryExtID in (" & dCustExtIdList & ") " & CPEHHFilter & ") " & SearchFilter & "" & _
                                "group by TR.CustomerPrimaryExtID, TR.HHID, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.LogixTransNum, TR.TransNum, TR.TerminalNum, TR.ExtLocationCode, TR.Replayed, TH.TransTotal " & HavingFilterTR & _
                                " UNION " & _
                                "select CustomerPrimaryExtID, " & TransactionTimeQuery_TH & ", ExtLocationCode, 0 as RedemptionAmount, 0 as RedemptionCount, " & _
                                "TerminalNum, LogixTransNum, POSTransNum as TransNum, 0 as DetailRecords, CustomerTypeID, PresentedCustomerID, PresentedCardTypeID, HHID, Replayed, 0 AS TransContext,isnull(TransTotal,0) as TransTotal " & _
                                "from TransHist as TH with (NoLock) " & _
                                "where " & IIf(CPEHHFilter <> "", "(", "") & "(CustomerPrimaryExtID in (" & dCustExtIdList & ") " & CPEHHFilter & ") " & SearchFilterTH & " and not exists " & _
                                "  (select LogixTransNum from TransRedemptionView as TR2 where " & IIf(CPEHHFilter <> "", "(", "") & "(CustomerPrimaryExtID in (" & dCustExtIdList & ") " & CPEHHFilter & ") and TH.LogixTransNum=TR2.LogixTransNum) " & _
                                "group by CustomerPrimaryExtID, HHID, CustomerTypeID, PresentedCustomerID, PresentedCardTypeID, LogixTransNum, POSTransNum, TerminalNum, ExtLocationCode, Replayed,TransTotal " & HavingFilterTH & _
              " UNION " & _
              "select TR.CustomerPrimaryExtID, " & TransactionTimeQuery & ", TR.ExtLocationCode, sum(TR.RedemptionAmount) as RedemptionAmount, sum(TR.RedemptionCount) as RedemptionCount, " & _
                                "TR.TerminalNum, TR.LogixTransNum, TR.TransNum, count(*) as DetailRecords, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.HHID, TR.Replayed, 0 AS TransContext,TH.TransTotal as TransTotal " & _
                                "from ThirdPartyTransRedemption as TR with (NoLock) Left Outer Join  ThirdPartytranshist TH ON TR.LogixTransNum = TH.LogixTransNum " & _
                                "where TR.CustomerTypeID in (0,1) and " & IIf(CPEHHFilter <> "", "(", "") & "(TR.CustomerPrimaryExtID in (" & dCustExtIdList & ") " & CPEHHFilter & ") " & SearchFilter & "" & _
                                "group by TR.CustomerPrimaryExtID, TR.HHID, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.LogixTransNum, TR.TransNum, TR.TerminalNum, TR.ExtLocationCode, TR.Replayed, TH.TransTotal " & HavingFilterTR & _
                                " UNION " & _
                                "select CustomerPrimaryExtID, " & TransactionTimeQuery_TH & ", ExtLocationCode, 0 as RedemptionAmount, 0 as RedemptionCount, " & _
                                "TerminalNum, LogixTransNum, POSTransNum as TransNum, 0 as DetailRecords, CustomerTypeID, PresentedCustomerID, PresentedCardTypeID, HHID, Replayed, 0 AS TransContext,isnull(TransTotal,0) as TransTotal " & _
                                "from ThirdPartytranshist as TH with (NoLock) " & _
                                "where " & IIf(CPEHHFilter <> "", "(", "") & "(CustomerPrimaryExtID in (" & dCustExtIdList & ") " & CPEHHFilter & ") " & SearchFilterTH & " and not exists " & _
                                "  (select LogixTransNum from ThirdPartyTransRedemption as TR2 where " & IIf(CPEHHFilter <> "", "(", "") & "(CustomerPrimaryExtID in (" & dCustExtIdList & ") " & CPEHHFilter & ") and TH.LogixTransNum=TR2.LogixTransNum) " & _
                                "group by CustomerPrimaryExtID, HHID, CustomerTypeID, PresentedCustomerID, PresentedCardTypeID, LogixTransNum, POSTransNum, TerminalNum, ExtLocationCode, Replayed,TransTotal " & HavingFilterTH & " " & _
                                "order by " & SortCol & " " & SortDir
            rstTrans = MyCommon.LWH_Select
        ElseIf redemptionFilter = 2 Then 'get only transactions that have redemptions (what we always used to do)
            MyCommon.QueryStr = "select TR.CustomerPrimaryExtID, " & TransactionTimeQuery & ", TR.ExtLocationCode, sum(TR.RedemptionAmount) as RedemptionAmount, sum(TR.RedemptionCount) as RedemptionCount, " & _
                                "TR.TerminalNum, TR.LogixTransNum, TR.TransNum, count(*) as DetailRecords, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.HHID, TR.Replayed, 0 AS TransContext, TH.TransTotal as TransTotal " & _
                                "from TransRedemptionView as TR with (NoLock) Left Outer Join  TransHist TH ON TR.LogixTransNum = TH.LogixTransNum " & _
                                "where TR.CustomerTypeID in (0,1) and " & IIf(CPEHHFilter <> "", "(", "") & "(TR.CustomerPrimaryExtID in (" & dCustExtIdList & ") " & CPEHHFilter & ") " & SearchFilter & "" & _
                                "group by TR.CustomerPrimaryExtID, TR.HHID, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.LogixTransNum, TR.TransNum, TR.TerminalNum, TR.ExtLocationCode, TR.Replayed, TH.TransTotal " & HavingFilterTR & _
              " UNION " & _
              "select TR.CustomerPrimaryExtID, " & TransactionTimeQuery & ", TR.ExtLocationCode, sum(TR.RedemptionAmount) as RedemptionAmount, sum(TR.RedemptionCount) as RedemptionCount, " & _
                                "TR.TerminalNum, TR.LogixTransNum, TR.TransNum, count(*) as DetailRecords, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.HHID, TR.Replayed, 0 AS TransContext, TH.TransTotal as TransTotal " & _
                                "from ThirdPartyTransRedemption as TR with (NoLock) Left Outer Join  ThirdPartytranshist TH ON TR.LogixTransNum = TH.LogixTransNum " & _
                                "where TR.CustomerTypeID in (0,1) and " & IIf(CPEHHFilter <> "", "(", "") & "(TR.CustomerPrimaryExtID in (" & dCustExtIdList & ") " & CPEHHFilter & ") " & SearchFilter & "" & _
                                "group by TR.CustomerPrimaryExtID, TR.HHID, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.LogixTransNum, TR.TransNum, TR.TerminalNum, TR.ExtLocationCode, TR.Replayed, TH.TransTotal " & HavingFilterTR & _
                                "order by " & SortCol & " " & SortDir
            rstTrans = MyCommon.LWH_Select
        ElseIf redemptionFilter = 3 Then 'get only transactions that do NOT have redemptions
            MyCommon.QueryStr = "select CustomerPrimaryExtID, " & TransactionTimeQuery_TH & ", ExtLocationCode, 0 as RedemptionAmount, 0 as RedemptionCount, " & _
                                "TerminalNum, LogixTransNum, POSTransNum as TransNum, 0 as DetailRecords, CustomerTypeID, PresentedCustomerID, PresentedCardTypeID, HHID, Replayed, 0 AS TransContext,isnull(TransTotal,0) as TransTotal " & _
                                "from TransHist as TH with (NoLock) " & _
                                "where " & IIf(CPEHHFilter <> "", "(", "") & "(CustomerPrimaryExtID in (" & dCustExtIdList & ") " & CPEHHFilter & ") " & SearchFilterTH & " and not exists " & _
                                "  (select LogixTransNum from TransRedemptionView as TR where " & IIf(CPEHHFilter <> "", "(", "") & "(CustomerPrimaryExtID in (" & dCustExtIdList & ") " & CPEHHFilter & ") and TH.LogixTransNum=TR.LogixTransNum) " & _
                                "group by CustomerPrimaryExtID, HHID, CustomerTypeID, PresentedCustomerID, PresentedCardTypeID, LogixTransNum, POSTransNum, TerminalNum, ExtLocationCode, Replayed,TransTotal " & HavingFilterTH & _
              " UNION " & _
              "select CustomerPrimaryExtID, " & TransactionTimeQuery_TH & ", ExtLocationCode, 0 as RedemptionAmount, 0 as RedemptionCount, " & _
                                "TerminalNum, LogixTransNum, POSTransNum as TransNum, 0 as DetailRecords, CustomerTypeID, PresentedCustomerID, PresentedCardTypeID, HHID, Replayed, 0 AS TransContext,isnull(TransTotal,0) as TransTotal " & _
                                "from ThirdPartytranshist as TH with (NoLock) " & _
                                "where " & IIf(CPEHHFilter <> "", "(", "") & "(CustomerPrimaryExtID in (" & dCustExtIdList & ") " & CPEHHFilter & ") " & SearchFilterTH & " and not exists " & _
                                "  (select LogixTransNum from ThirdPartyTransRedemption as TR where " & IIf(CPEHHFilter <> "", "(", "") & "(CustomerPrimaryExtID in (" & dCustExtIdList & ") " & CPEHHFilter & ") and TH.LogixTransNum=TR.LogixTransNum) " & _
                                "group by CustomerPrimaryExtID, HHID, CustomerTypeID, PresentedCustomerID, PresentedCardTypeID, LogixTransNum, POSTransNum, TerminalNum, ExtLocationCode, Replayed,TransTotal " & HavingFilterTH & _
                                "order by " & SortCol & " " & SortDir
            rstTrans = MyCommon.LWH_Select
      ElseIf redemptionFilter = 4 Then 'get only transactions that do NOT have redemptions
          Dim NoOfdays1 As Integer = NoOfdays - 1 
          Dim startDate As String = Now.AddDays(-NoOfdays1).ToShortDateString() & " 00:00:00"
          Dim endDate As String = Now.ToShortDateString() & " 23:59:59"
  
          Dim timeCondition As String = "Having " & TransactionTimeCondQuery & " between '" & startDate & "' and '" & endDate & "' "
          Dim timeCondition_TH As String = "Having " & TransactionTimeCondQuery_TH & " between '" & startDate & "' and '" & endDate & "' "
            MyCommon.QueryStr = "select TR.CustomerPrimaryExtID, " & TransactionTimeQuery & ", TR.ExtLocationCode, sum(TR.RedemptionAmount) as RedemptionAmount, sum(TR.RedemptionCount) as RedemptionCount, " & _
                                  "TR.TerminalNum, TR.LogixTransNum, TR.TransNum, count(*) as DetailRecords, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.HHID, TR.Replayed, 0 AS TransContext,TH.TransTotal as TransTotal " & _
                                  "from TransRedemptionView as TR with (NoLock) Left Outer Join  TransHist TH ON TR.LogixTransNum = TH.LogixTransNum " & _
                                  "where TR.CustomerTypeID in (0,1) and " & IIf(CPEHHFilter <> "", "(", "") & "(TR.CustomerPrimaryExtID in (" & dCustExtIdList & ") " & CPEHHFilter & ") " & SearchFilter & "" & _
                                        "group by TR.CustomerPrimaryExtID, TR.HHID, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.LogixTransNum, TR.TransNum, TR.TerminalNum, TR.ExtLocationCode, TR.Replayed, TH.TransTotal " & HavingFilterTR & _
                                 timeCondition & _
                                  " UNION " & _
                                  "select CustomerPrimaryExtID, " & TransactionTimeQuery_TH & ", ExtLocationCode, 0 as RedemptionAmount, 0 as RedemptionCount, " & _
                                  "TerminalNum, LogixTransNum, POSTransNum as TransNum, 0 as DetailRecords, CustomerTypeID, PresentedCustomerID, PresentedCardTypeID, HHID, Replayed, 0 AS TransContext,isnull(TransTotal,0) as TransTotal " & _
                                  "from TransHist as TH with (NoLock) " & _
                                  "where " & IIf(CPEHHFilter <> "", "(", "") & "(CustomerPrimaryExtID in (" & dCustExtIdList & ") " & CPEHHFilter & ") " & SearchFilterTH & " and not exists " & _
                                  "  (select LogixTransNum from TransRedemptionView as TR2 where " & IIf(CPEHHFilter <> "", "(", "") & "(CustomerPrimaryExtID in (" & dCustExtIdList & ") " & CPEHHFilter & ") and TH.LogixTransNum=TR2.LogixTransNum) " & _
                                        "group by CustomerPrimaryExtID, HHID, CustomerTypeID, PresentedCustomerID, PresentedCardTypeID, LogixTransNum, POSTransNum, TerminalNum, ExtLocationCode, Replayed,TransTotal " & HavingFilterTH & _
                                   timeCondition_TH & _
                " UNION " & _
                "select TR.CustomerPrimaryExtID, " & TransactionTimeQuery & ", TR.ExtLocationCode, sum(TR.RedemptionAmount) as RedemptionAmount, sum(TR.RedemptionCount) as RedemptionCount, " & _
                                  "TR.TerminalNum, TR.LogixTransNum, TR.TransNum, count(*) as DetailRecords, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.HHID, TR.Replayed, 0 AS TransContext,TH.TransTotal as TransTotal " & _
                                  "from ThirdPartyTransRedemption as TR with (NoLock) Left Outer Join  ThirdPartytranshist TH ON TR.LogixTransNum = TH.LogixTransNum " & _
                                  "where TR.CustomerTypeID in (0,1) and " & IIf(CPEHHFilter <> "", "(", "") & "(TR.CustomerPrimaryExtID in (" & dCustExtIdList & ") " & CPEHHFilter & ") " & SearchFilter & "" & _
                                        "group by TR.CustomerPrimaryExtID, TR.HHID, TR.CustomerTypeID, TR.PresentedCustomerID, TR.PresentedCardTypeID, TR.LogixTransNum, TR.TransNum, TR.TerminalNum, TR.ExtLocationCode, TR.Replayed, TH.TransTotal " & HavingFilterTR & _
                                   timeCondition & _
                                  " UNION " & _
                                  "select CustomerPrimaryExtID, " & TransactionTimeQuery_TH & ", ExtLocationCode, 0 as RedemptionAmount, 0 as RedemptionCount, " & _
                                  "TerminalNum, LogixTransNum, POSTransNum as TransNum, 0 as DetailRecords, CustomerTypeID, PresentedCustomerID, PresentedCardTypeID, HHID, Replayed, 0 AS TransContext,isnull(TransTotal,0) as TransTotal " & _
                                  "from ThirdPartytranshist as TH with (NoLock) " & _
                                  "where " & IIf(CPEHHFilter <> "", "(", "") & "(CustomerPrimaryExtID in (" & dCustExtIdList & ") " & CPEHHFilter & ") " & SearchFilterTH & " and not exists " & _
                                  "  (select LogixTransNum from ThirdPartyTransRedemption as TR2 where " & IIf(CPEHHFilter <> "", "(", "") & "(CustomerPrimaryExtID in (" & dCustExtIdList & ") " & CPEHHFilter & ") and TH.LogixTransNum=TR2.LogixTransNum) " & _
                                        "group by CustomerPrimaryExtID, HHID, CustomerTypeID, PresentedCustomerID, PresentedCardTypeID, LogixTransNum, POSTransNum, TerminalNum, ExtLocationCode, Replayed,TransTotal " & HavingFilterTH & " " & _
                                 timeCondition_TH & _
                                  "order by " & SortCol & " " & SortDir
          rstTrans = MyCommon.LWH_Select
          End If
      End If
    sizeOfData = rstTrans.Rows.Count
    PageNum = Convert.ToInt32(Request.QueryString("pagenum"))
    startPosition = PageNum * linesPerPage
    endPosition = IIf(sizeOfData < startPosition + linesPerPage, sizeOfData, startPosition + linesPerPage) - 1
    MorePages = IIf(sizeOfData > startPosition + linesPerPage, True, False)
  If ExtLocationCode <> "" Then
    If sizeOfData = 100 AndAlso ShowAll Then
      infomessage = "Max 100 Records Reached - Refine Search Criteria"
    End If
    
    SortUrl = "customer-transactions.aspx?LocCode=" & ExtLocationCode & IIF(sLast4CardID <> "","&amp;Last4=" & sLast4CardID, "") & IIf(sStartTime <> "", "&amp;StartTime=" & sStartTime, "") & IIf(sEndTime <> "", "&amp;EndTime=" & sEndTime, "") & "&amp;redemptionFilter=" & redemptionFilter & "&amp;transterms=" & TransTerms & "&amp;pagenum=0" & IIf(ShowAll, "&amp;ShowAll=True", "")
  Else
    SortUrl = "customer-transactions.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "&amp;redemptionFilter=" & redemptionFilter & "&amp;transterms=" & TransTerms & "&amp;pagenum=0" & IIf(ShowAll, "&amp;ShowAll=True", "")
  End If
%>
<script type="text/javascript">
<!--
function showTransDetail(row, btn ) {
  var elemTr = document.getElementById("trTrans" + row);
  
  if (elemTr != null && btn != null) {
    elemTr.style.display = (btn.value == "+") ? "" : "none";
    btn.value = (btn.value == "+") ? "-" : "+";  
  }
}

function searchTrans() {
  var elemTransTerms = document.getElementById("transterms");
  var elemLocCode = document.getElementById("LocCode");
  var transTerms = '';
  var qryStr = '';
  
  if (elemTransTerms != null) {
    transTerms = elemTransTerms.value;
  }
  <%
    Dim strTerms = Request.QueryString("searchterms")
    If (strTerms <> "") Then
      strTerms = strTerms.Replace("'", "\'")
      strTerms = strTerms.Replace("""", "\""")
    End If
   %>
   if (elemLocCode != null) {
    qryStr = 'customer-transactions.aspx?search=Search&searchterms=<%Sendb(strTerms)%>&LocCode=' + elemLocCode.value + '<%Sendb(IIF(sLast4CardID <> "","&Last4=" & sLast4CardID, ""))%><%Sendb(IIf(sStartTime <> "", "&StartTime=" & sStartTime, ""))%><%Sendb(IIf(sEndTime <> "", "&EndTime=" & sEndTime, ""))%>&redemptionFilter=<%Sendb(redemptionFilter)%>&offerSearch=Search<%Sendb(IIf(ShowAll, "&ShowAll=True", ""))%>&transterms=' + transTerms;
   }
   else {
    qryStr = 'customer-transactions.aspx?search=Search&searchterms=<%Sendb(strTerms)%>&CustPK=<%Sendb(CustomerPK)%><%Sendb(IIf(CardPK > 0, "&CardPK=" & CardPK, ""))%>&redemptionFilter=<%Sendb(redemptionFilter)%>&chkIncludeAllCards=<%Sendb(IncludeAllCards)%>&offerSearch=Search<%Sendb(IIf(ShowAll, "&ShowAll=True", ""))%>&transterms=' + transTerms;
  }
  document.location = qryStr;
}

function submitTransSearch(e) {
  var key = e.which ? e.which : e.keyCode;
  
  if (key == 13) {
    if (e && e.preventDefault) {
      e.preventDefault(); // DOM style
      searchTrans();
    } else {
      e.keyCode = 9;
      searchTrans();
    }
    return false;
  }
  return true;
}
//-->
</script>
<script type="text/javascript" src="../javascript/jquery.min.js"></script>
<script type="text/javascript" src="../javascript/thickbox.js">
  function HR1_onclick() {
  }
</script>
<form id="mainform" name="mainform" action="customer-transactions.aspx">
<div id="intro">
  <h1 id="title">
    <%
      If CardPK = 0 Then
        If (IsHouseholdID) Then
          Sendb(Copient.PhraseLib.Lookup("term.household", LanguageID))
        Else If (ExtLocationCode <> "") Then
          Sendb(Copient.PhraseLib.Lookup("term.store", LanguageID) & " #" & ExtLocationCode)
        Else
          Sendb(Copient.PhraseLib.Lookup("term.customer", LanguageID))
        End If
      Else
        If (IsHouseholdID) Then
          Sendb(Copient.PhraseLib.Lookup("term.householdcard", LanguageID) & " #" & ExtCardID)
        Else If (ExtLocationCode <> "") Then
          Sendb(Copient.PhraseLib.Lookup("term.store", LanguageID) & " #" & ExtLocationCode)
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
        Sendb(": " & MyCommon.TruncateString(FullName, 30))
      End If
      If (restrictLinks AndAlso URLtrackBack <> "") Then
        Send(" - <a href=""" & URLtrackBack & """>" & Copient.PhraseLib.Lookup("customer-inquiry.return", LanguageID) & "</a>")
      End If
    %>
  </h1>
  
    <%
'=============================================================================	CLOUDSOL-1252  	
	  If((IsHouseholdID) = False And (ExtLocationCode = "")) Then
	    If (Request.QueryString("chkIncludeAllCards") <> "") Then
          IncludeAllCards = Convert.ToInt32(Request.QueryString("chkIncludeAllCards"))
        End If 	  
	    Send("  <input type=""checkbox"" id=""chkIncludeAllCards"" name=""chkIncludeAllCards"" value=1 " & IIf(IncludeAllCards=1, " checked ", " ") & " onchange=mainform.submit() /> ") 
	    Send(" <label for=""chkIncludeAllCards"">" & Copient.PhraseLib.Lookup("term.includeallcards", LanguageID) & "</label> ")
      End If
'=============================================================================	CLOUDSOL-1252  	  
    %>
  
  <div id="controls" <% Sendb(IIf(restrictLinks AndAlso URLtrackBack <> "", " style=""width:115px;""", "")) %>>
    <%
      If (CustomerPK > 0 And Logix.UserRoles.ViewCustomerNotes) Then
        Send_CustomerNotes(CustomerPK, CardPK)
      End If
    %>
  </div>
</div>
<div id="main">
  <%
    If (Logix.UserRoles.ViewTransHistory = False) Then
      Send(Copient.PhraseLib.Lookup("error.forbidden", LanguageID))
      Send("</div>")
      Send("</form>")
      GoTo done
    End If
    If (infoMessage <> "") Then
      Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div><br />")
    End If
    If CustomerPK > 0 Then
    Send("  <input type=""hidden"" id=""CustPK"" name=""CustPK"" value=""" & CustomerPK & """ />")
    End If
    If CardPK > 0 Then
      Send("  <input type=""hidden"" id=""CardPK"" name=""CardPK"" value=""" & CardPK & """ />")
    End If
    If ExtLocationCode <> "" Then
      Send("  <input type=""hidden"" id=""LocCode"" name=""LocCode"" value=""" & ExtLocationCode & """ />")
    End If
    If sLast4CardID <> "" Then
      Send("  <input type=""hidden"" id=""Last4"" name=""Last4"" value=""" & sLast4CardID & """ />")
    End If
    If (sStartTime <> "") AndAlso (sEndTime <> "") Then
      Send("  <input type=""hidden"" id=""StartTime"" name=""StartTime"" value=""" & sStartTime & """ />")
      Send("  <input type=""hidden"" id=""EndTime"" name=""EndTime"" value=""" & sEndTime & """ />")
    End If
    If (Request.QueryString("mode") = "summary") Then
      Send("<input type=""hidden"" id=""mode"" name=""mode"" value=""summary"" />")
      Send("<input type=""hidden"" id=""exiturl"" name=""exiturl"" value=""" & URLtrackBack & """ />")
    End If
    If ShowAll Then
      Send("<input type=""hidden"" id=""showall"" name=""showall"" value=""True"" />")
    End If
  %>
  <div id="column">
    <%
      If rstTrans.Rows.Count >= 20 Then
        If ShowAll Then
          If CustomerPK >0 Then
          Send("<span style=""float:right;font-weight:bold;margin:2px 0 0 10px;""><a href=""customer-transactions.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & """>" & Copient.PhraseLib.Lookup("term.showbrief", LanguageID) & "</a></span>")
          Else If ExtLocationCode <> "" Then
            Send("<span style=""float:right;font-weight:bold;margin:2px 0 0 10px;""><a href=""customer-transactions.aspx?LocCode=" & ExtLocationCode & IIF(sLast4CardID <> "","&amp;Last4=" & sLast4CardID, "") & IIf(sStartTime <> "", "&amp;StartTime=" & sStartTime, "") & IIf(sEndTime <> "", "&amp;EndTime=" & sEndTime, "") & "&amp;redemptionFilter=" & redemptionFilter & """>" & Copient.PhraseLib.Lookup("term.showbrief", LanguageID) & "</a></span>")
          End If
        Else
          If CustomerPK >0 Then
          Send("<span style=""float:right;font-weight:bold;margin:2px 0 0 10px;""><a href=""customer-transactions.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "&amp;ShowAll=True"">" & Copient.PhraseLib.Lookup("term.showall", LanguageID) & "</a></span>")
          Else If ExtLocationCode <> "" Then
            Send("<span style=""float:right;font-weight:bold;margin:2px 0 0 10px;""><a href=""customer-transactions.aspx?LocCode=" & ExtLocationCode & IIF(sLast4CardID <> "","&amp;Last4=" & sLast4CardID, "") & IIf(sStartTime <> "", "&amp;StartTime=" & sStartTime, "") & IIf(sEndTime <> "", "&amp;EndTime=" & sEndTime, "") & "&amp;redemptionFilter=" & redemptionFilter &  "&amp;ShowAll=True"">" & Copient.PhraseLib.Lookup("term.showall", LanguageID) & "</a></span>")
          End If
        End If
      End If
      If IsHouseholdID Then
        Send("<span style=""float:right;font-weight:bold;margin:2px 0 0 10px;""><a href=""javascript:openPopup('customer-transactions-legend.aspx')"">" & Copient.PhraseLib.Lookup("term.legend", LanguageID) & "</a></span>")
      End If
      Send("<br clear=""all"" />")
      If (Logix.UserRoles.ViewTransHistory AndAlso (CustomerPK > 0 OrElse ExtLocationCode <> "")) Then
        Send(" <div id=""listbar"">")
        Send("  <div id=""searcher"">")
        Send("   <input type=""text"" style=""font-family:arial;font-size:12px;width:45%;"" id=""transterms"" name=""transterms"" maxlength=""256"" class=""mediumshort"" value=""" & TransTerms & """ onkeydown=""submitTransSearch(event);"" />")
        Send("   <input type=""button"" style=""font-family:arial;font-size:12px;width:45%;"" id=""btnOffer"" name=""btnOffer"" value=""" & Copient.PhraseLib.Lookup("term.search", LanguageID) & """ onclick=""searchTrans();"" />")
        Send("  </div>")
        Send("  <div id=""paginator"">")
        If (sizeOfData > 0) Then
          If (PageNum > 0) Then
            If ExtLocationCode <> "" Then
              Send("   <span id=""first""><a href=""customer-transactions.aspx?LocCode=" & ExtLocationCode & IIF(sLast4CardID <> "","&amp;Last4=" & sLast4CardID, "") &  IIF(sStartTime = "", "", "&StartTime="&sStartTime) & IIF(sEndTime = "", "", "&EndTime="&sEndTime) & "&amp;redemptionFilter=" & redemptionFilter & "&amp;pagenum=0&amp;transterms=" & TransTerms & "&amp;sortcol=" & SortCol & "&amp;sortdir=" & SortDir & IIf(ShowAll, "&amp;ShowAll=True", "") & """><b>|</b>◄" & Copient.PhraseLib.Lookup("term.first", LanguageID) & "</a>&nbsp;</span>")
              Send("   <span id=""previous""><a href=""customer-transactions.aspx?LocCode=" & ExtLocationCode & IIF(sLast4CardID <> "","&amp;Last4=" & sLast4CardID, "") & IIF(sStartTime = "", "", "&StartTime="&sStartTime) & IIF(sEndTime = "", "", "&EndTime="&sEndTime) & "&amp;redemptionFilter=" & redemptionFilter & "&amp;pagenum=" & PageNum - 1 & "&amp;transterms=" & TransTerms & "&amp;sortcol=" & SortCol & "&amp;sortdir=" & SortDir & IIf(ShowAll, "&amp;ShowAll=True", "") & """>◄" & Copient.PhraseLib.Lookup("term.previous", LanguageID) & "</a>&nbsp;</span>")
            Else
              Send("   <span id=""first""><a href=""customer-transactions.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "&amp;chkIncludeAllCards=" & IncludeAllCards & "&amp;redemptionFilter=" & redemptionFilter & "&amp;pagenum=0&amp;transterms=" & TransTerms & "&amp;sortcol=" & SortCol & "&amp;sortdir=" & SortDir & IIf(ShowAll, "&amp;ShowAll=True", "") & """><b>|</b>◄" & Copient.PhraseLib.Lookup("term.first", LanguageID) & "</a>&nbsp;</span>")
              Send("   <span id=""previous""><a href=""customer-transactions.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "&amp;chkIncludeAllCards=" & IncludeAllCards & "&amp;redemptionFilter=" & redemptionFilter & "&amp;pagenum=" & PageNum - 1 & "&amp;transterms=" & TransTerms & "&amp;sortcol=" & SortCol & "&amp;sortdir=" & SortDir & IIf(ShowAll, "&amp;ShowAll=True", "") & """>◄" & Copient.PhraseLib.Lookup("term.previous", LanguageID) & "</a>&nbsp;</span>")
            End If
          Else
            Send("   <span id=""first""><b>|</b>◄" & Copient.PhraseLib.Lookup("term.first", LanguageID) & "&nbsp;</span>")
            Send("   <span id=""previous"">◄" & Copient.PhraseLib.Lookup("term.previous", LanguageID) & "&nbsp;</span>")
          End If
        Else
          Send("   <span id=""first""><b>|</b>◄" & Copient.PhraseLib.Lookup("term.first", LanguageID) & "&nbsp;</span>")
          Send("   <span id=""previous"">◄" & Copient.PhraseLib.Lookup("term.previous", LanguageID) & "&nbsp;</span>")
        End If
        If sizeOfData = 0 Then
          Send("   &nbsp;[ " & Copient.PhraseLib.Lookup("term.noresults", LanguageID) & " ]&nbsp;")
        Else
          Send("   &nbsp;[ " & Copient.PhraseLib.Lookup("term.results", LanguageID) & " <b><span id=""startPos"">" & startPosition + 1 & "</span>-<span id=""endPos"">" & endPosition + 1 & "</span></b> " & Copient.PhraseLib.Lookup("term.of", LanguageID) & " <b>" & sizeOfData & "</b> ]&nbsp;")
        End If
        If (sizeOfData > 0) Then
          If (MorePages) Then
            If ExtLocationCode <> "" Then
              Send("   <span id=""next""><a href=""customer-transactions.aspx?LocCode=" & ExtLocationCode & IIF(sLast4CardID <> "","&amp;Last4=" & sLast4CardID, "") &  IIF(sStartTime = "", "", "&StartTime="&sStartTime) & IIF(sEndTime = "", "", "&EndTime="&sEndTime) & "&amp;redemptionFilter=" & redemptionFilter & "&amp;pagenum=" & PageNum + 1 & "&amp;transterms=" & TransTerms & "&amp;sortcol=" & SortCol & "&amp;sortdir=" & SortDir & IIf(ShowAll, "&amp;ShowAll=True", "") & """><b>|</b>" & Copient.PhraseLib.Lookup("term.next", LanguageID) & "►</a>&nbsp;</span>")
              Send("   <span id=""last""><a href=""customer-transactions.aspx?LocCode=" & ExtLocationCode & IIF(sLast4CardID <> "","&amp;Last4=" & sLast4CardID, "") & IIF(sStartTime = "", "", "&StartTime="&sStartTime) & IIF(sEndTime = "", "", "&EndTime="&sEndTime) & "&amp;redemptionFilter=" & redemptionFilter & "&amp;pagenum=" & (Math.Ceiling(sizeOfData / linesPerPage) - 1) & "&amp;transterms=" & TransTerms & "&amp;sortcol=" & SortCol & "&amp;sortdir=" & SortDir & IIf(ShowAll, "&amp;ShowAll=True", "") & """>" & Copient.PhraseLib.Lookup("term.last", LanguageID) & "►</a>&nbsp;</span>")
            Else
              Send("   <span id=""next""><a href=""customer-transactions.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "")  & "&amp;chkIncludeAllCards=" & IncludeAllCards & "&amp;redemptionFilter=" & redemptionFilter & "&amp;pagenum=" & PageNum + 1 & "&amp;transterms=" & TransTerms & "&amp;sortcol=" & SortCol & "&amp;sortdir=" & SortDir & IIf(ShowAll, "&amp;ShowAll=True", "") & """>" & Copient.PhraseLib.Lookup("term.next", LanguageID) & "►</a>&nbsp;</span>")
              Send("   <span id=""last""><a href=""customer-transactions.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "")  & "&amp;chkIncludeAllCards=" & IncludeAllCards & "&amp;redemptionFilter=" & redemptionFilter & "&amp;pagenum=" & (Math.Ceiling(sizeOfData / linesPerPage) - 1) & "&amp;transterms=" & TransTerms & "&amp;sortcol=" & SortCol & "&amp;sortdir=" & SortDir & IIf(ShowAll, "&amp;ShowAll=True", "") & """>" & Copient.PhraseLib.Lookup("term.last", LanguageID) & "►<b>|</b></a>&nbsp;</span>")
            End If
          Else
            Send("   <span id=""next"">" & Copient.PhraseLib.Lookup("term.next", LanguageID) & "►&nbsp;</span>")
            Send("   <span id=""last"">" & Copient.PhraseLib.Lookup("term.last", LanguageID) & "►<b>|</b>&nbsp;</span>")
          End If
        Else
          Send("   <span id=""next"">" & Copient.PhraseLib.Lookup("term.next", LanguageID) & "►&nbsp;</span>")
          Send("   <span id=""last"">" & Copient.PhraseLib.Lookup("term.last", LanguageID) & "►<b>|</b></span><br />")
        End If
        Send("  </div>")
        Send("  <div id=""filter"">")
        Send("   <select id=""redemptionFilter"" name=""redemptionFilter"" onchange=mainform.submit()>")
        Send("    <option value=""1""" & IIf(redemptionFilter = 1, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("customer-transactions.AllTransactions", LanguageID) & "</option>")
        Send("    <option value=""2""" & IIf(redemptionFilter = 2, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("customer-transactions.OnlyWithRedemptions", LanguageID) & "</option>")
        Send("    <option value=""3""" & IIf(redemptionFilter = 3, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("customer-transactions.OnlyWithoutRedemptions", LanguageID) & "</option>")
        If (NoOfdays > 0) Then
            Send("    <option value=""4""" & IIf(redemptionFilter = 4, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("customer-transactions.AllTransactions", LanguageID) & " - " & NoOfdays & " " & Copient.PhraseLib.Lookup("term.days", LanguageID).ToLower() & "</option>")
        End If
        
        Send("   </select>")
        Send("  </div>")
        Send(" </div>")
        
        Send("<div style=""overflow-x:auto;"">")
        Send(" <table summary=""" & Copient.PhraseLib.Lookup("term.transactionhistory", LanguageID) & """>")
        Send("  <thead>")
    %>
    <tr>
      <th align="left" class="th-button" scope="col">
        <% Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>
      </th>
      <th align="left" class="th-datetime" scope="col">
        <a href="<% Sendb(SortUrl & "&amp;sortcol=TransactionDate&amp;sortdir=" & IIf(SortDir = "desc", "asc", "desc")) %>">
          <% Sendb(Copient.PhraseLib.Lookup("term.date", LanguageID))%></a>
        <%
          If SortCol = "TransactionDate" Then
            If SortDir = "asc" Then
              Sendb("<span class=""sortarrow"">&#9660;</span>")
            Else
              Sendb("<span class=""sortarrow"">&#9650;</span>")
            End If
          End If
        %>
      </th>
      <th align="left" class="th-cardholder" scope="col">
        <a href="<% Sendb(SortUrl & "&amp;sortcol=PresentedCustomerID&amp;sortdir=" & IIf(SortDir = "desc", "asc", "desc")) %>">
          <% Sendb(Copient.PhraseLib.Lookup("term.cardnumber", LanguageID))%></a>
        <%
          If SortCol = "PresentedCustomerID" Then
            If SortDir = "asc" Then
              Sendb("<span class=""sortarrow"">&#9660;</span>")
            Else
              Sendb("<span class=""sortarrow"">&#9650;</span>")
            End If
          End If
        %>
      </th>
      <th align="left" class="th-cardholder" scope="col">
        <a href="<% Sendb(SortUrl & "&amp;sortcol=HHID&amp;sortdir=" & IIf(SortDir = "desc", "asc", "desc")) %>">
          <% Sendb(Copient.PhraseLib.Lookup("term.household", LanguageID))%></a>
        <%
          If SortCol = "HHID" Then
            If SortDir = "asc" Then
              Sendb("<span class=""sortarrow"">&#9660;</span>")
            Else
              Sendb("<span class=""sortarrow"">&#9650;</span>")
            End If
          End If
        %>
      </th>
      <th align="left" class="th-id" scope="col">
        <a href="<% Sendb(SortUrl & "&amp;sortcol=ExtLocationCode&amp;sortdir=" & IIf(SortDir = "desc", "asc", "desc")) %>">
          <% Sendb(Copient.PhraseLib.Lookup("term.store", LanguageID))%></a>
        <%
          If SortCol = "ExtLocationCode" Then
            If SortDir = "asc" Then
              Sendb("<span class=""sortarrow"">&#9660;</span>")
            Else
              Sendb("<span class=""sortarrow"">&#9650;</span>")
            End If
          End If
        %>
      </th>
      <th align="center" class="th-id" scope="col" style="text-align: center;">
        <a href="<% Sendb(SortUrl & "&amp;sortcol=TerminalNum&amp;sortdir=" & IIf(SortDir = "desc", "asc", "desc")) %>">
          <% Sendb(Copient.PhraseLib.Lookup("term.terminal", LanguageID))%></a>
        <%
          If SortCol = "TerminalNum" Then
            If SortDir = "asc" Then
              Sendb("<span class=""sortarrow"">&#9660;</span>")
            Else
              Sendb("<span class=""sortarrow"">&#9650;</span>")
            End If
          End If
        %>
      </th>
      <th align="center" class="th-transaction" scope="col" style="text-align: center;">
        <a href="<% Sendb(SortUrl & "&amp;sortcol=TransNum&amp;sortdir=" & IIf(SortDir = "desc", "asc", "desc")) %>"
          title="<% Sendb(Copient.PhraseLib.Lookup("term.transactionnumber", LanguageID)) %>">
          <% Sendb(Copient.PhraseLib.Lookup("term.txn", LanguageID) & "#")%></a>
        <%
          If SortCol = "TransNum" Then
            If SortDir = "asc" Then
              Sendb("<span class=""sortarrow"">&#9660;</span>")
            Else
              Sendb("<span class=""sortarrow"">&#9650;</span>")
            End If
          End If
        %>
      </th>
      <th align="center" class="th-replayed" scope="col" style="text-align: center;">
        <a href="<% Sendb(SortUrl & "&amp;sortcol=Replayed&amp;sortdir=" & IIf(SortDir = "desc", "asc", "desc")) %>"
          title="<% Sendb(Copient.PhraseLib.Lookup("term.replayed", LanguageID)) %>">
          <% Sendb("R")%></a>
        <%
          If SortCol = "Replayed" Then
            If SortDir = "asc" Then
              Sendb("<span class=""sortarrow"">&#9660;</span>")
            Else
              Sendb("<span class=""sortarrow"">&#9650;</span>")
            End If
          End If
        %>
      </th>
      <% If (redemptionFilter <> 3) Then%>
      <th align="right" class="th-redemptions" scope="col" style="text-align: right;">
        <a href="<% Sendb(SortUrl & "&amp;sortcol=RedemptionCount&amp;sortdir=" & IIf(SortDir = "desc", "asc", "desc")) %>"
          title="<% Sendb(Copient.PhraseLib.Lookup("term.redemptions", LanguageID)) %>">
          <% Sendb("Rdms")%></a>
        <%
          If SortCol = "RedemptionCount" Then
            If SortDir = "asc" Then
              Sendb("<span class=""sortarrow"">&#9660;</span>")
            Else
              Sendb("<span class=""sortarrow"">&#9650;</span>")
            End If
          End If
        %>
      </th>
      <% End If%>
      <% If (TrxTotalDisplayOnUI = 1) Then%>
      <th align="right" class="th-transactionsTotal" scope="col" style="text-align: left;">
        <a href="<% Sendb(SortUrl & "&amp;sortcol=TransTotal&amp;sortdir=" & IIf(SortDir = "desc", "asc", "desc")) %>"
          title="<% Sendb(Copient.PhraseLib.Lookup("term.TransactionTotal", LanguageID)) %>">
          <% Sendb("TransTotal")%></a>
        <%
          If SortCol = "TransTotal" Then
            If SortDir = "asc" Then
              Sendb("<span class=""sortarrow"">&#9660;</span>")
            Else
              Sendb("<span class=""sortarrow"">&#9650;</span>")
            End If
          End If
        %>
      </th>
      <% End If%>
    </tr>
    <%
      Send("  </thead>")
      Send("  <tbody>")
      Dim transRows As ArrayList
      If (rst.Rows.Count > 0) Then
        transCt = 0
        transRows = GetSubList(rstTrans, startPosition, endPosition)
        If transRows.Count > 0 Then
          For Each row In transRows
            transCt += 1
            Select Case Convert.ToInt32(MyCommon.NZ(row.Item("TransContext"), 0))
              Case 1
                ContextStyle = "color:#cc0000;"
              Case 2
                ContextStyle = "color:#009900;"
              Case 3
                ContextStyle = "color:#009900;"
              Case Else
                ContextStyle = "color:#000000;"
            End Select
                    Send("<tr" & IIf(Not IsCMOnly AndAlso ExtCardID = IIf(IsDBNull(row.Item("CustomerPrimaryExtID")), 0, row.Item("CustomerPrimaryExtID").ToString()) AndAlso IsHouseholdID, Shaded, "") & ">")
            Send("  <td><a href=""#""><input class=""ex more"" type=""button"" value=""+"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.details", LanguageID) & """" & IIf(MyCommon.NZ(row.Item("RedemptionCount"), 0) = 0, " disabled=""disabled""", " onclick=""javascript:showTransDetail(" & transCt & ", this);""") & " /></a></td>")
            Sendb("  <td style=""" & ContextStyle & """>")
            If MyCommon.NZ(row.Item("TransactionDate"), "1/1/1980") = "1/1/1980" Then
              TransDateStart = Now.ToString("yyyy-MM-dd 00:00:00")
              TransDateEnd = Now.ToString("yyyy-MM-dd 23:59:59")
              Sendb(Copient.PhraseLib.Lookup("term.unknown", LanguageID))
            Else
              TransDateStart = Date.Parse(row.Item("TransactionDate")).ToString("yyyy-MM-dd 00:00:00")
              TransDateEnd = Date.Parse(row.Item("TransactionDate")).ToString("yyyy-MM-dd 23:59:59")
              Sendb(Logix.ToShortDateTimeString(row.Item("TransactionDate"), MyCommon))
            End If
            'Sendb(DisplayTransactionDate(MyCommon.NZ(row.Item("TransactionDate"), "1/1/1980"), TransDateStart, TransDateEnd))
            Send("</td>")
            Send("  <td style=""" & ContextStyle & """>")
            Send("    <span title=""" & Copient.PhraseLib.Lookup("term.presented", LanguageID) & """>P:" & MyCommon.NZ(row.Item("PresentedCustomerID"), UnknownPhrase) & "</span><br />")
            Dim TempCardPK, TempCustPK As Long
                    If Not IsCMOnly AndAlso Not ExtCardID = IIf(IsDBNull(row.Item("CustomerPrimaryExtID")), 0, row.Item("CustomerPrimaryExtID").ToString()) AndAlso IsHouseholdID Then
                        Dim extCardID1 As String = ""
                        extCardID1 = MyCryptLib.SQL_StringEncrypt(MyCommon.Pad_ExtCardID(MyCommon.NZ(row.Item("CustomerPrimaryExtID").ToString(), ""), 1))

                        MyCommon.QueryStr = "Select CardPK, CustomerPK from CardIDs with (NoLock) where ExtCardID='" & extCardID1 & "';"
                        rstTemp = MyCommon.LXS_Select()
                        If rstTemp.Rows.Count > 0 Then
                            TempCardPK = MyCommon.NZ(rstTemp.Rows(0).Item("CardPK"), 0)
                            TempCustPK = MyCommon.NZ(rstTemp.Rows(0).Item("CustomerPK"), 0)
                        End If
                        Send("    <span title=""" & Copient.PhraseLib.Lookup("term.resolved", LanguageID) & """>R:" & "<a href=""customer-transactions.aspx?CustPK=" & TempCustPK & "&CardPK=" & TempCardPK & """>" & MyCommon.NZ(row.Item("CustomerPrimaryExtID"), UnknownPhrase) & "</a></span>")
                    Else
                        Dim CardID = MyCommon.NZ(row.Item("CustomerPrimaryExtID"), UnknownPhrase)
                        If ExtLocationCode <> "" Then
                            MyCommon.QueryStr = "Select CardPK, CustomerPK from CardIDs with (NoLock) where ExtCardID='" & MyCryptLib.SQL_StringEncrypt(CardID) & "';"
                            rstTemp = MyCommon.LXS_Select()
                            If rstTemp.Rows.Count > 0 Then
                                TempCardPK = MyCommon.NZ(rstTemp.Rows(0).Item("CardPK"), 0)
                                TempCustPK = MyCommon.NZ(rstTemp.Rows(0).Item("CustomerPK"), 0)
                            End If
                            Send("    <span title=""" & Copient.PhraseLib.Lookup("term.resolved", LanguageID) & """>R: <a href=""customer-general.aspx?CustPK=" & TempCustPK & "&CardPK=" & TempCardPK & """>" & CardID & "</a></span>")
                        Else
                            Send("    <span title=""" & Copient.PhraseLib.Lookup("term.resolved", LanguageID) & """>R:" & CardID & "</span>")
                        End If
                    End If
            Send("  </td>")
            If (MyCommon.NZ(row.Item("HHID"), "") = "") AndAlso (MyCommon.NZ(row.Item("CustomerTypeID"), 0) = 1) Then
                        Send("  <td style=""" & ContextStyle & """>" & IIf(IsDBNull(row.Item("CustomerPrimaryExtID")), "&nbsp;", row.Item("CustomerPrimaryExtID").ToString()) & "</td>")
            Else
                        Send("  <td style=""" & ContextStyle & """>" & IIf(IsDBNull(row.Item("HHID")), "&nbsp;", row.Item("HHID").ToString()) & "</td>")
            End If
            
            Send("  <td style=""" & ContextStyle & """>" & MyCommon.NZ(row.Item("ExtLocationCode"), UnknownPhrase) & "</td>")
            Send("  <td style=""" & ContextStyle & "text-align:center;"">" & MyCommon.NZ(row.Item("TerminalNum"), UnknownPhrase) & "</td>")
            
            Dim LogixTransNum = MyCommon.NZ(row.Item("LogixTransNum"), UnknownPhrase)
            If (MyCommon.Fetch_CPE_SystemOption(169) = 1) Then
              Send("  <td style=""" & ContextStyle & "text-align:center;""><span title=""" & MyCommon.NZ(row.Item("TransNum"), UnknownPhrase) & """><a href=""javascript:openPopup('/logix/customer-transaction-items.aspx?TransNum=" &LogixTransNum & "')"">" & LogixTransNum & "</a></span></td>")
            Else
              Send("  <td style=""" & ContextStyle & "text-align:center;""><span title=""" & MyCommon.NZ(row.Item("LogixTransNum"), UnknownPhrase) & """>" & MyCommon.NZ(row.Item("TransNum"), UnknownPhrase) & "</span></td>")
            End If
            Send("  <td style=""" & ContextStyle & """>" & IIf(MyCommon.NZ(row.Item("Replayed"), 0) > 0, "<span title=""" & Copient.PhraseLib.Lookup("term.replayed", LanguageID) & """ style=""color:#dd0000;cursor:default;font-size:10px;font-weight:bold;"">R</span>", "") & "</td>")
            If redemptionFilter <> 3 Then
              'If MyCommon.NZ(row.Item("Replayed"), 0) > 0 AndAlso (MyCommon.NZ(row.Item("RedemptionAmount"), 0) > 0) Then
              '  Send("  <td style=""" & ContextStyle & "text-align:right;""><span style=""background-color:#dddddd;color:#888888;"">" & MyCommon.NZ(row.Item("RedemptionAmount"), UnknownPhrase) & "</span></td>")
              'Else
              '  Send("  <td style=""" & ContextStyle & "text-align:right;"">" & MyCommon.NZ(row.Item("RedemptionAmount"), UnknownPhrase) & "</td>")
              'End If
              Send("  <td style=""" & ContextStyle & "text-align:right;"">" & MyCommon.NZ(row.Item("RedemptionCount"), 0) & "</td>")
            End If
            If TrxTotalDisplayOnUI = 1 Then
              'Restricting Trx.Total value decimal places to two			  
              If row.Item("TransTotal").ToString.Contains(".") Then
                Send("  <td style=""" & ContextStyle & """>" & row.Item("TransTotal").ToString.Substring(0, row.Item("TransTotal").ToString.IndexOf(".") + 3) & "</td>")
              Else
                Send("  <td style=""" & ContextStyle & """>" & MyCommon.NZ(row.Item("TransTotal"), UnknownPhrase) & "</td>")
              End If
            End If
            Send("</tr>")
            ' write detail rows
            If (MyCommon.NZ(row.Item("RedemptionCount"), 0) = 0) Then
              'No detail line needed
            Else
              Send("<tr id=""trTrans" & transCt & """ style=""display:none;color:#888888;"" class=""transdetail"">")
              'Send("  <td></td>")
              Dim DetailCPEHHFilter As String = "" 'This is a variant to the main CPEHHFilter that must be used in the detail row when the TransContext is 2 (for redemptions made in previous households).
              If MyCommon.NZ(row.Item("TransContext"), 0) = 2 AndAlso IsCMOnly = False AndAlso IsHouseholdID AndAlso CPEHHFilter <> "" Then
                DetailCPEHHFilter = "and HHID is NULL) OR HHID='" & MyCommon.NZ(row.Item("HHID"), "&nbsp;") & "' "
              Else
                DetailCPEHHFilter = ""
              End If
                        MyCommon.QueryStr = "select OfferID, RedemptionAmount, RedemptionCount, Replayed, SVAmount, SVProgramID, PointsAmount, PointsProgramID from TransRedemptionView with (NoLock) " & _
                                            "where (" & IIf(CPEHHFilter <> "", "(", "") & "CustomerPrimaryExtID in (" & dCustExtIdList & ") " & IIf(DetailCPEHHFilter <> "", DetailCPEHHFilter, CPEHHFilter) & " OR (CustomerPrimaryExtID='" & ExtCardID & "' and CustomerTypeID=1)) " & _
                                            "  and LogixTransNum='" & MyCommon.NZ(row.Item("LogixTransNum"), UnknownPhrase) & "' " & _
                                            "  and ExtLocationCode='" & MyCommon.NZ(row.Item("ExtLocationCode"), "") & "' " & _
                                            "  and TerminalNum='" & MyCommon.NZ(row.Item("TerminalNum"), "") & "' and TransNum='" & MyCommon.NZ(row.Item("TransNum"), "") & "' and " & _
                                            IIf(ShowPOSTimeStamp, "POSTimeStamp", "TransDate") & " between '" & TransDateStart & "' and '" & TransDateEnd & "' " & _
                                            "order by Replayed desc, TransDate, RedemptionAmount asc;"
              rst2 = MyCommon.LWH_Select
              If (rst2.Rows.Count > 0) Then
                transOffers = New StringBuilder(500)
                transRdmptAmt = New StringBuilder(100)
                transRdmptCt = New StringBuilder(100)
                Dim ProgramString As StringBuilder = New StringBuilder(100)
                For Each row2 In rst2.Rows
                   MyCommon.QueryStr = "select SUBSTRING(Name,1,70) as OfferName, ExtOfferID as XID from Offers with (NoLock) where OfferID=" & MyCommon.NZ(row2.Item("OfferID"), -1) &
                                                        " union all " &
                                                        "select SUBSTRING(IncentiveName,1,70) as OfferName, ClientOfferID as XID from CPE_Incentives with (NoLock) where IncentiveID=" & MyCommon.NZ(row2.Item("OfferID"), -1) & ";"
                  rst3 = MyCommon.LRT_Select
                  If (rst3.Rows.Count > 0) Then
                    OfferName = MyCommon.NZ(rst3.Rows(0).Item("OfferName"), "")
                    XID = MyCommon.NZ(rst3.Rows(0).Item("XID"), "[" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "]")
                  Else
                    OfferName = "[" & Copient.PhraseLib.Lookup("history.offer-delete", LanguageID) & "]"
                    XID = ""
                  End If
                  'First cell-----
                  If MyCommon.NZ(row.Item("Replayed"), 0) > 0 Then
                    transOffers.Append(IIf(MyCommon.NZ(row2.Item("RedemptionAmount"), 0) > 0, "<div style=""background-color:#dddddd;color:#888888"">", "<div style=""color:#cc0000;"">"))
                  Else
                    transOffers.Append("<div>")
                  End If
                  transOffers.Append(Copient.PhraseLib.Lookup("term.xid", LanguageID) & ": " & XID & "&nbsp;&nbsp;&nbsp;&nbsp;")
                  transOffers.Append(Copient.PhraseLib.Lookup("term.id", LanguageID) & ": ")
                  transOffers.Append(MyCommon.NZ(row2.Item("OfferID"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)) & "&nbsp;&nbsp;&nbsp;&nbsp;")
                  transOffers.Append(Copient.PhraseLib.Lookup("term.name", LanguageID) & ": " & MyCommon.SplitNonSpacedString(OfferName, 30) & "<br />")
                  transOffers.Append("</div>")
                  'Second cell-----
                  If MyCommon.NZ(row.Item("Replayed"), 0) > 0 Then
                    transRdmptAmt.Append(IIf(MyCommon.NZ(row2.Item("RedemptionAmount"), 0) > 0, "<div style=""background-color:#dddddd;color:#888888"">", "<div style=""color:#cc0000;"">"))
                  Else
                    transRdmptAmt.Append("<div>")
                  End If
                    
                  Dim DisplayRdmptAmt As Double = 0
                  Dim AmtDisplayString As String = ""
                  If (MyCommon.NZ(row2.Item("PointsProgramID"), 0) > 0) Then
                    If IsDBNull(row2.Item("PointsProgramID")) Then
                      AmtDisplayString = Copient.PhraseLib.Lookup("term.unknown", LanguageID)
                    Else
                      AmtDisplayString = Format(row2.Item("PointsAmount"), "#0")
                    End If
                      
                    ProgramString.Append("<input id=""pointsAdj" & MyCommon.NZ(row2.Item("PointsProgramID"), 0) & """ class=""adjust"" type=""button"" onclick=""javascript:openPopup('point-adjust-program.aspx?ProgramID=" & MyCommon.NZ(row2.Item("PointsProgramID"), 0) & "&CustomerPK=" & CustomerPK & "&CardPK=" & CardPK & "&Opener=" & CopientFileName & "');"" title=""" & Copient.PhraseLib.Lookup("customer-transactions.AdjustPoints", LanguageID) & """ style="""" value=""P"" name=""svAdj""><br />")
                  ElseIf (MyCommon.NZ(row2.Item("SVProgramID"), 0) > 0) Then
                    DisplayRdmptAmt = MyCommon.NZ(row2.Item("SVAmount"), -1)
                    AmtDisplayString = IIf(DisplayRdmptAmt < 0, Copient.PhraseLib.Lookup("term.unknown", LanguageID), Format(DisplayRdmptAmt, "#0.000"))
                    ProgramString.Append("<input id=""svAdj" & MyCommon.NZ(row2.Item("SVProgramID"), 0) & """ class=""adjust"" type=""button"" onclick=""javascript:openPopup('sv-adjust-program.aspx?ProgramID=" & MyCommon.NZ(row2.Item("SVProgramID"), 0) & "&CustomerPK=" & CustomerPK & "&CardPK=" & CardPK & "&Opener=" & CopientFileName & "');"" title=""" & Copient.PhraseLib.Lookup("customer-transactions.AdjustSV", LanguageID) & """ style="""" value=""S"" name=""svAdj""><br />")
                  Else
                    AmtDisplayString = "$" & MyCommon.NZ(row2.Item("RedemptionAmount"), Copient.PhraseLib.Lookup("term.unknown", LanguageID))
                    ProgramString.Append("<input class=""adjust"" type=""button"" style=""visibility:hidden;""><br />")
                  End If
                  transRdmptAmt.Append(AmtDisplayString & "<br /></div>")
                  transRdmptCt.Append("<div>" & MyCommon.NZ(row2.Item("RedemptionCount"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)) & "</div>")
                Next
                Send("  <td style=""text-align:left;"" valign=""top"" colspan=""6"">" & transOffers.ToString & "</td>")
                Send("  <td style=""text-align:right;"" valign=""top"">" & transRdmptAmt.ToString & "</td>")
                Send("  <td style=""text-align:right;"" valign=""top"">" & ProgramString.ToString() & "</td>")
                Send("  <td style=""text-align:right;"" valign=""top"">" & transRdmptCt.ToString & "</td>")
              End If
              Send("</tr>")
              TotalRedeemAmt += MyCommon.NZ(row.Item("RedemptionAmount"), 0.0)
              TotalRedeemCt += MyCommon.NZ(row.Item("RedemptionCount"), 0)
            End If
          Next
        Else
          Send("<tr>")
          Send("  <td colspan=""7"" style=""text-align:center""><i>" & Copient.PhraseLib.Lookup("customer-inquiry.nohistory", LanguageID) & "</i></td>")
          Send("</tr>")
        End If
      Else
        Send("<tr>")
        Send("  <td colspan=""7"" style=""text-align:center""><i>" & Copient.PhraseLib.Lookup("customer-inquiry.nohistory", LanguageID) & "</i></td>")
        Send("</tr>")
      End If
      MyCommon.Close_LogixWH()
      Send("  </tbody>")
      Send(" </table>")
      Send(" </div>")
      Send(" <hr class=""hidden"" id=""HR1"" onclick=""return HR1_onclick()"" />")
    End If
    %>
  </div>
  <br clear="all" />
</div>
</form>
<script runat="server">
  Function GetCustomerPK(ByRef MyCommon As Copient.CommonInc, ByVal PrimaryExtID As String) As Integer
    Dim dt As DataTable
    Dim CustomerPK As Integer = 0
    
    If (Not MyCommon.LXSadoConn.State = ConnectionState.Open) Then MyCommon.Open_LogixXS()
    
    MyCommon.QueryStr = "select CustomerPK from Customers with (NoLock) where PrimaryExtID='" & PrimaryExtID & "';"
    dt = MyCommon.LXS_Select
    If (dt.Rows.Count > 0) Then
      CustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)
    End If
    
    Return CustomerPK
  End Function
  
  Function GetSubList(ByVal dt As DataTable, ByVal startIndex As Integer, ByVal endIndex As Integer) As ArrayList
    Dim subList As New ArrayList(20)
    Dim i As Integer
    
    For i = startIndex To endIndex
      If (dt.Rows.Count - 1 >= i) Then
        subList.Add(dt.Rows(i))
      End If
    Next
    
    Return subList
  End Function
  
  Public Sub MergeDataTables(ByVal source As DataTable, ByVal destination As DataTable)
    destination.BeginLoadData()
    For i As Integer = 0 To source.Rows.Count - 1
      destination.LoadDataRow(source.Rows(i).ItemArray, True)
    Next
    destination.EndLoadData()
  End Sub
  
  Function SelectIntoDataTable(ByVal selectFilter As String, ByVal sortFilter As String, ByVal ShowAll As Boolean, ByVal sourceDataTable As DataTable) As DataTable
    Dim newDataTable As DataTable = sourceDataTable.Clone
    Dim dataRows As DataRow() = sourceDataTable.Select(selectFilter, sortFilter)
    Dim typeDataRow As DataRow
    Dim rowCount As Integer = 0
    
    For Each typeDataRow In dataRows
      newDataTable.ImportRow(typeDataRow)
      rowCount += 1
      If (ShowAll = False) AndAlso (rowCount = 20) Then
        Exit For
      End If
    Next
    
    Return newDataTable
  End Function
  
  'Public Sub GetTransactionDate(ByVal TransDateStart As String, ByVal TransDateEnd As String, ByVal DisplayTransDate As String)
    
  'End Sub
  
  'Public Sub SetTransactionStartDate(ByVal TransactionDate As String)
    
  'End Sub
  
  'If MyCommon.NZ(row.Item("TransactionDate"), "1/1/1980") = "1/1/1980" Then
  '              TransDateStart = Now.ToString("yyyy-MM-dd 00:00:00")
  '              TransDateEnd = Now.ToString("yyyy-MM-dd 23:59:59")
  '              Sendb(Copient.PhraseLib.Lookup("term.unknown", LanguageID))
  '            Else
  '              TransDateStart = Date.Parse(row.Item("TransactionDate")).ToString("yyyy-MM-dd 00:00:00")
  '              TransDateEnd = Date.Parse(row.Item("TransactionDate")).ToString("yyyy-MM-dd 23:59:59")
  '              Sendb(Format(row.Item("TransactionDate"), "dd MMM yyyy, HH:mm:ss"))
  '            End If
  
  'Function will have to run the query to populate depending on the system option for using the POSTimeStamp
  Public Function DisplayTransactionDate(ByVal ProcessDate As DateTime) As String
    Dim DisplayTransDate As String = ""
    'If DateTime.Compare(ProcessDate, "1/1/1980") = 0 Then 'Date from datebase was NULL 
    If ProcessDate = "1/1/1980" Then
      DisplayTransDate = Copient.PhraseLib.Lookup("term.unknown", LanguageID)
    Else
      DisplayTransDate = Date.Parse(Format(ProcessDate, "dd MMM yyyy, HH:mm:ss"))
    End If
    
    Return DisplayTransDate
  End Function
  
  Public Sub SetTransactionDates(ByRef TransDateStart As String, ByRef TransDateEnd As String, ByVal TransactionDate As DateTime)
    If TransactionDate = "1/1/1980" Then 'Date from datebase was NULL 
      TransDateStart = Now.ToString("yyyy-MM-dd 00:00:00")
      TransDateEnd = Now.ToString("yyyy-MM-dd 23:59:59")
    Else
      TransDateStart = Date.Parse(Format(TransactionDate, "yyyy-MM-dd 00:00:00"))
      TransDateEnd = Date.Parse(Format(TransactionDate, "yyyy-MM-dd 23:59:59"))
    End If
    
  End Sub
  
</script>
<%
done:
  Send_BodyEnd("mainform", "transterms")
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  MyCommon = Nothing
  Logix = Nothing
%>
