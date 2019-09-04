<%@ Page Language="vb" Debug="true" CodeFile="../LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: CAM-customer-transactions.aspx 
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
  Dim rstTrans As DataTable
  Dim row As DataRow
  Dim rst2 As DataTable
  Dim row2 As DataRow
  Dim rst3 As DataTable
  Dim dt As DataTable
  Dim CustomerPK As Long
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
  Dim offerPointBuf As StringBuilder
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
  Dim DisabledPtsAdj As String = ""
  Dim LogixTransNum As String = ""
  
  ' default urls for links from this page
  Dim URLCAMOfferSum As String = "CAM-offer-sum.aspx"
  Dim URLcgroupedit As String = "/logix/cgroup-edit.aspx"
  Dim URLpointedit As String = "/logix/point-edit.aspx"
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
  Dim HavingFilter As String = ""
  Dim tempDate As Date
  Dim SortCol As String = "TransactionDate"
  Dim SortDir As String = "desc"
  Dim SortUrl As String = ""
  Dim redemptionFilter As Integer = 2
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "CAM-customer-transactions.aspx"
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
  
  If (MyCommon.Extract_Val(Request.QueryString("CustPK")) > 0) Then
    CustomerPK = MyCommon.Extract_Val(Request.QueryString("CustPK"))
  ElseIf (MyCommon.Extract_Val(Request.QueryString("CustomerPK")) > 0) Then
    CustomerPK = MyCommon.Extract_Val(Request.QueryString("CustomerPK"))
  End If
  
  CardPK = MyCommon.Extract_Val(Request.QueryString("CardPK"))
  If CardPK = 0 Then
    CardPK = MyLookup.FindCardPK(CustomerPK, 2)
  End If
  
  If CardPK > 0 Then
    ExtCardID = MyLookup.FindExtCardIDFromCardPK(CardPK)
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
  
  If (MyCommon.Extract_Val(Request.QueryString("CustPK")) = 0) Then
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "/logix/customer-inquiry.aspx")
  End If
  
  'Set the default redemption filter value, then update if present in querystring
  redemptionFilter = MyCommon.Fetch_SystemOption(104)
  If (Request.QueryString("redemptionFilter") <> "") Then
    redemptionFilter = MyCommon.Extract_Val(Request.QueryString("redemptionFilter"))
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
      MyCommon.QueryStr = "select C.InitialCardID, C.CustomerPK, C.CustomerTypeID, C.HHPK, C.FirstName, C.MiddleName, C.LastName, " & _
                          "CE.Address, CE.City, CE.State, CE.Zip, CE.PhoneAsEntered as Phone, CE.email " & _
                          "from Customers C with (NoLock) " & _
                          "left join CustomerExt CE with (NoLock) on CE.CustomerPK=C.CustomerPK " & _
                          "where C.CustomerPK=" & CustomerPK & ";"
    Else
      ' IF the page was called from an outside application set ClientUserID1 to the outside passed in value
      If (inCardNumber <> "" And Request.QueryString("mode") = "summary") Then
        ExtCardID = MyCommon.Pad_ExtCardID(inCardNumber,2)
        MyCommon.QueryStr = "select CustomerPK from CardIDs with (NoLock) where ExtCardID='" & MyCryptLib.SQL_StringEncrypt(ExtCardID) & "';"
        dt = MyCommon.LXS_Select
        If dt.Rows.Count > 0 Then
          CustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)
        End If
        searchterms = Request.QueryString("searchterms")
        MyCommon.QueryStr = "select C.InitialCardID, C.CustomerPK, C.CustomerTypeID, C.HHPK, C.FirstName, C.MiddleName, C.LastName, " & _
                            "CE.Address, CE.City, CE.State, CE.Zip, CE.PhoneAsEntered as Phone, CE.email " & _
                            "from Customers C with (NoLock) " & _
                            "left join CustomerExt CE with (NoLock) on CE.CustomerPK=C.CustomerPK " & _
                            "where C.CustomerPK=" & CustomerPK & ";"
      End If
      If (Request.QueryString("searchterms") <> "" And ExtCardID = "") Then
        ExtCardID = MyCommon.Pad_ExtCardID( MyCommon.Parse_Quotes(Left(Request.QueryString("searchterms"), 26)),2)
                MyCommon.QueryStr = "select C.InitialCardID, C.CustomerPK, C.CustomerTypeID, C.HHPK, C.FirstName, C.MiddleName, C.LastName, " & _
                                    "CE.Address, CE.City, CE.State, CE.Zip, CE.PhoneAsEntered as Phone, CE.email " & _
                                    "from Customers C with (NoLock) " & _
                                    "left join CustomerExt CE with (NoLock) on CE.CustomerPK=C.CustomerPK " & _
                                    "where C.InitialCardID='" & MyCryptLib.SQL_StringEncrypt(ExtCardID) & "' or CE.PhoneDigitsOnly = '" & MyCryptLib.SQL_StringEncrypt(MyCommon.DigitsOnly(Request.QueryString("searchterms"))) & _
                                    "' or CE.email = '" & MyCryptLib.SQL_StringEncrypt(MyCommon.Parse_Quotes(Request.QueryString("searchterms"))) & "' or C.LastName like '%" & MyCommon.Parse_Quotes(Request.QueryString("searchterms")) & "%';"
      End If
    End If
    rstResults = MyCommon.LXS_Select
    
    If (rstResults.Rows.Count = 1) Then
      ' ok we found a primary key for the external id provided
            CustomerPK = rstResults.Rows(0).Item("CustomerPK")
            'No need to decrypt as its passed its passed to inline SQL Query
      ClientUserID1 = MyCommon.NZ(rstResults.Rows(0).Item("InitialCardID"), "")
      IsHouseholdID = (MyCommon.NZ(rstResults.Rows(0).Item("CustomerTypeID"), 0) = 1)
    Else
      infoMessage = "" & Copient.PhraseLib.Lookup("customer.notfound", LanguageID) & ""
      infoMessage = infoMessage & " <a href=""/logix/customer-inquiry.aspx?mode=add&Search=Search" & extraLink & "&searchterms=" & Request.QueryString("searchterms") & """>" & Copient.PhraseLib.Lookup("term.add", LanguageID) & "</a>"
    End If
        
  End If
  
  UnknownPhrase = Copient.PhraseLib.Lookup("term.unknown", LanguageID)
  
  If CardPK > 0 Then
    Send_HeadBegin("term.customer", "term.adjustments", MyCommon.Extract_Val(ExtCardID))
  Else
    Send_HeadBegin("term.customer", "term.adjustments")
  End If
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  If restrictLinks Then
    Send("<link rel=""stylesheet"" href=""/css/logix-restricted.css"" type=""text/css"" media=""screen"" />")
  End If
%>
<style type="text/css">
  .transOffers {
    color: #888888;
    padding-left: 60px;
  }
  .transRdmptAmt {
    color: #888888;
    text-align: right;
    width: 60px;
  }
  .transRdmptCt {
    color: #888888;
    padding-right: 0px;
    text-align: right;
    width: 67px;
  }
</style>  
<%
  Send_Scripts()
  Send_HeadEnd()
  
  Send_BodyBegin(1)
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  If (Not restrictLinks) Then
    Send_Tabs(Logix, 3)
    If CardPK > 0 Then
      Send_Subtabs(Logix, 33, 6, LanguageID, CustomerPK, , CardPK)
    Else
      Send_Subtabs(Logix, 33, 6, LanguageID, CustomerPK)
    End If
  Else
    If CardPK > 0 Then
      Send_Subtabs(Logix, 94, 6, LanguageID, CustomerPK, extraLink, CardPK)
    Else
      Send_Subtabs(Logix, 94, 6, LanguageID, CustomerPK, extraLink)
    End If
  End If
  If (Logix.UserRoles.AccessCustomerInquiry = False) Then
    Send_Denied(1, "perm.customer-access")
    GoTo done
  End If
  
  If (Request.QueryString("searchterms") <> "") Then
    CustExtID = ExtCardID
        'CustExtIdList = "'" & MyCommon.Parse_Quotes(MyCryptLib.SQL_StringEncrypt(CustExtID)) & "'"
        CustExtIdList = "'" & MyCommon.Parse_Quotes(CustExtID) & "'"
  ElseIf (Request.QueryString("editterms") <> "") Then
    CustExtID = Request.QueryString("editterms")
        'CustExtIdList = "'" & MyCommon.Parse_Quotes(MyCryptLib.SQL_StringEncrypt(CustExtID)) & "'"
        CustExtIdList = "'" & MyCommon.Parse_Quotes(CustExtID) & "'"
    ElseIf ClientUserID1 <> "" Then
        'No need to Encrypt as its derived from SQL and is already encrypted.
        CustExtIdList = "'" & ClientUserID1 & "'"
  Else
        'CustExtIdList = "'" & MyCryptLib.SQL_StringEncrypt(ExtCardID) & "'"
        CustExtIdList = "'" & ExtCardID & "'"
  End If
  
  MyCommon.Open_LogixWH()
  
  ' Process search request, if applicable 
  TransTerms = Request.QueryString("transterms")
  If (TransTerms <> "") Then
    If (Date.TryParse(TransTerms, tempDate)) Then
      HavingFilter = " having Max(TransDate) between '" & tempDate.ToString("yyyy-MM-ddT00:00:00") & "'" & _
                     " and '" & tempDate.ToString("yyyy-MM-ddT23:59:59") & "'"
      SearchFilter = ""
      SearchFilterTH = ""
    Else
      SearchFilter = " and (CustomerPrimaryExtID='" & TransTerms & "'" & _
                     " or ExtLocationCode='" & TransTerms & "'" & _
                     " or LogixTransNum='" & TransTerms & "'" & _
                     " or TransNum='" & TransTerms & "'" & _
                     " or TerminalNum='" & TransTerms & "')"
      SearchFilterTH = " and (CustomerPrimaryExtID='" & TransTerms & "'" & _
                       " or ExtLocationCode='" & TransTerms & "'" & _
                       " or LogixTransNum='" & TransTerms & "'" & _
                       " or POSTransNum='" & TransTerms & "'" & _
                       " or TerminalNum='" & TransTerms & "')"
    End If
  Else
    SearchFilter = ""
    SearchFilterTH = ""
    HavingFilter = ""
  End If
  
  ' Process sort request, if applicable
  If (Request.QueryString("sortcol") <> "") Then SortCol = Request.QueryString("sortcol")
  If (Request.QueryString("sortdir") <> "") Then
    SortDir = Request.QueryString("sortdir")
  Else
    SortDir = "desc"
  End If
  
  'If SortCol = "TerminalNum" Or SortCol = "POSTransNum" Then
  '  Try
  '    MyCommon.QueryStr = "select CustomerPrimaryExtId, Max(TransDate) as TransactionDate, " & _
  '                        "       ExtLocationCode, TerminalNum, POSTransNum, LogixTransNum " & _
  '                        "from TransHistory with (NoLock) " & _
  '                        "where CustomerTypeID=2 and CustomerPrimaryExtID in (" & CustExtIdList & ") " & SearchFilter & _
  '                        "group by CustomerPrimaryExtId, POSTransNum, TerminalNum, ExtLocationCode, LogixTransNum " & HavingFilter & " " & _
  '                        "order by cast(" & SortCol & " as int) " & SortDir
  '    rstTrans = MyCommon.LWH_Select

  '  Catch ex As Exception
  '    MyCommon.QueryStr = "select CustomerPrimaryExtId, Max(TransDate) as TransactionDate, " & _
  '                        "       ExtLocationCode, TerminalNum, POSTransNum, LogixTransNum " & _
  '                        "from TransHistory with (NoLock) " & _
  '                        "where CustomerTypeID=2 and CustomerPrimaryExtID in (" & CustExtIdList & ") " & SearchFilter & _
  '                        "group by CustomerPrimaryExtId, POSTransNum, TerminalNum, ExtLocationCode, LogixTransNum " & HavingFilter & " " & _
  '                        "order by " & SortCol & " " & SortDir
  '    rstTrans = MyCommon.LWH_Select

  '  End Try
  'Else
  '  MyCommon.QueryStr = "select CustomerPrimaryExtId, Max(TransDate) as TransactionDate, " & _
  '                      "       ExtLocationCode, TerminalNum, POSTransNum, LogixTransNum " & _
  '                      "from TransHistory with (NoLock) " & _
  '                      "where CustomerTypeID=2 and CustomerPrimaryExtID in (" & CustExtIdList & ") " & SearchFilter & _
  '                      "group by CustomerPrimaryExtId, POSTransNum, TerminalNum, ExtLocationCode, LogixTransNum " & HavingFilter & " " & _
  '                      "order by " & SortCol & " " & SortDir
  '  rstTrans = MyCommon.LWH_Select
  'End If
  
  If redemptionFilter = 1 Then 'get all transactions with redemptions (from TransRedemption view) and all transactions
    'without redemptions (from TransHist view, minus the transaction #s from TransRedemptionView),
    'and union them together.
    MyCommon.QueryStr = "select CustomerPrimaryExtID, Max(TransDate) as TransactionDate, ExtLocationCode, sum(RedemptionAmount) as RedemptionAmount, sum(RedemptionCount) as RedemptionCount, " & _
                        "TerminalNum, LogixTransNum, TransNum, count(*) as DetailRecords, PresentedCustomerID, PresentedCardTypeID, HHID, Replayed " & _
                        "from TransRedemptionView as TR with (NoLock) " & _
                        "where CustomerTypeID=2 and (CustomerPrimaryExtID in (" & CustExtIdList & ")) " & SearchFilter & "" & _
                        "group by CustomerPrimaryExtID, HHID, PresentedCustomerID, PresentedCardTypeID, LogixTransNum, TransNum, TerminalNum, ExtLocationCode, Replayed " & HavingFilter & _
                        " UNION " & _
                        "select CustomerPrimaryExtID, Max(TransDate) as TransactionDate, ExtLocationCode, 0 as RedemptionAmount, 0 as RedemptionCount, " & _
                        "TerminalNum, LogixTransNum, POSTransNum as TransNum, 0 as DetailRecords, PresentedCustomerID, PresentedCardTypeID, HHID, Replayed " & _
                        "from TransHist as TH with (NoLock) " & _
                        "where CustomerTypeID=2 and (CustomerPrimaryExtID in (" & CustExtIdList & ")) " & SearchFilterTH & " and not exists " & _
                        "  (select LogixTransNum from TransRedemptionView as TR2 where (CustomerPrimaryExtID in (" & CustExtIdList & ")) and TH.LogixTransNum=TR2.LogixTransNum) " & _
                        "group by CustomerPrimaryExtID, HHID, PresentedCustomerID, PresentedCardTypeID, LogixTransNum, POSTransNum, TerminalNum, ExtLocationCode, Replayed " & HavingFilter & " " & _
                        "order by " & SortCol & " " & SortDir
  ElseIf redemptionFilter = 2 Then 'get only transactions that have redemptions (what we always used to do)
    MyCommon.QueryStr = "select CustomerPrimaryExtID, Max(TransDate) as TransactionDate, ExtLocationCode, sum(RedemptionAmount) as RedemptionAmount, sum(RedemptionCount) as RedemptionCount, " & _
                        "TerminalNum, LogixTransNum, TransNum, count(*) as DetailRecords, PresentedCustomerID, PresentedCardTypeID, HHID, Replayed " & _
                        "from TransRedemptionView as TR with (NoLock) " & _
                        "where CustomerTypeID=2 and (CustomerPrimaryExtID in (" & CustExtIdList & ")) " & SearchFilter & "" & _
                        "group by CustomerPrimaryExtID, HHID, PresentedCustomerID, PresentedCardTypeID, LogixTransNum, TransNum, TerminalNum, ExtLocationCode, Replayed " & HavingFilter & " " & _
                        "order by " & SortCol & " " & SortDir
  ElseIf redemptionFilter = 3 Then 'get only transactions that do NOT have redemptions
    MyCommon.QueryStr = "select CustomerPrimaryExtID, Max(TransDate) as TransactionDate, ExtLocationCode, 0 as RedemptionAmount, 0 as RedemptionCount, " & _
                        "TerminalNum, LogixTransNum, POSTransNum as TransNum, 0 as DetailRecords, PresentedCustomerID, PresentedCardTypeID, HHID, Replayed " & _
                        "from TransHist as TH with (NoLock) " & _
                        "where CustomerTypeID=2 and (CustomerPrimaryExtID in (" & CustExtIdList & ")) " & SearchFilterTH & " and not exists " & _
                        "  (select LogixTransNum from TransRedemptionView as TR where (CustomerPrimaryExtID in (" & CustExtIdList & ")) and TH.LogixTransNum=TR.LogixTransNum) " & _
                        "group by CustomerPrimaryExtID, HHID, PresentedCustomerID, PresentedCardTypeID, LogixTransNum, POSTransNum, TerminalNum, ExtLocationCode, Replayed " & HavingFilter & " " & _
                        "order by " & SortCol & " " & SortDir
  End If
  rstTrans = MyCommon.LWH_Select
  
  sizeOfData = rstTrans.Rows.Count
  PageNum = MyCommon.Extract_Val(Request.QueryString("pagenum"))
  startPosition = PageNum * linesPerPage
  endPosition = IIf(sizeOfData < startPosition + linesPerPage, sizeOfData, startPosition + linesPerPage) - 1
  MorePages = IIf(sizeOfData > startPosition + linesPerPage, True, False)
  SortUrl = "CAM-customer-transactions.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "&amp;redemptionFilter=" & redemptionFilter & "&amp;transterms=" & TransTerms & "&amp;pagenum=0"
%>

<script type="text/javascript">
<!--
function showTransDetail(row, btn ) {
  var elemTr = document.getElementById("trTrans" + row);
  var isOpen = false;
  
  if (elemTr != null && btn != null) {
    elemTr.style.display = (btn.src.indexOf("plus.png") > -1) ? "" : "none";
    isOpen = (btn.src.indexOf('minus.png') > -1) ? true : false;
    if (isOpen) {
      btn.src = '/images/plus.png';
    } else {
      btn.src = '/images/minus.png';
    }
  }
}

function searchTrans() {
  var elemTransTerms = document.getElementById("transterms");
  var transTerms = '';
  
  if (elemTransTerms != null) { transTerms = elemTransTerms.value; }
  <%
    Dim strTerms As String = Request.QueryString("searchterms")
    If (strTerms <> "") Then
      strTerms = strTerms.Replace("'", "\'")
      strTerms = strTerms.Replace("""", "\""")
    End If
   %>
  var qryStr = 'CAM-customer-transactions.aspx?search=Search&searchterms=<%Sendb(strTerms)%>&CustPK=<%Sendb(CustomerPK)%><%Sendb(IIf(CardPK > 0, "&CardPK=" & CardPK, ""))%>&redemptionFilter=<%Sendb(redemptionFilter)%>&offerSearch=Search&transterms=' + transTerms;
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
      if (mode == 'ShowSelect') {
        updateSelectDialog(respTxt, args[0], args[1]);
      }
    }
  }
  self.xmlHttpReq.send(qryStr);
}

function showSelectDialog(logixTransNum, transNum, transDate) {
  var elem = document.getElementById("selectDialog");
  var elemBkgrd = document.getElementById("disabledBkgrd");
  var args = new Array(logixTransNum, transDate)
  var qryStr = 'CAMOfferSelection=1&LogixTransNum=' + logixTransNum + '&TransNum=' + transNum + '&TransDate=' + transDate + '&CustPK=<%Sendb(CustomerPK)%><%Sendb(IIf(CardPK > 0, "&CardPK=" & CardPK, ""))%>';
  
  if (elemBkgrd != null) {
    elemBkgrd.style.display = "block";
  }
  xmlhttpPost('CAM-Feeds.aspx?'+ qryStr, qryStr, 'ShowSelect', args);
}

function updateSelectDialog(respTxt, logixTransNum, transDate) {
  var elem = document.getElementById("selectDialog");
  var elemSearch = document.getElementById("txtSearch");
  var boxes = null;
  
  if (elem != null) {
    elem.innerHTML = respTxt;
    elem.style.display = "block";
    
    // set focus to the search box. Work-around for AJAX created text box focus issue.
    boxes = elem.getElementsByTagName("INPUT"); 
    for(var i = 0; i < boxes.length; i++) {
      if (boxes[i] != null && boxes[i].id == "txtSearch") {
        boxes[i].focus();
        boxes[i].select();
      }
    }
  }
}

function closeDialog () {
  var elem = document.getElementById("selectDialog");
  var elemBkgrd = document.getElementById("disabledBkgrd");
  
  if (elem != null) {
    elem.style.display = "none";
  }
  if (elemBkgrd != null) {
    elemBkgrd.style.display = "none";
  }
}

function searchFromDialog(logixTransNum, transNum, transDate) {
  var elemSearch = document.getElementById("txtSearch");
  var searchValue =  '';
  var args = new Array(logixTransNum, transDate)
  
  if (elemSearch != null) { searchValue = elemSearch.value; }
  var qryStr = 'CAMOfferSelection=1&LogixTransNum=' + logixTransNum + '&TransNum=' + transNum + '&TransDate=' + transDate + '&CustPK=<%Sendb(CustomerPK)%><%Sendb(IIf(CardPK > 0, "&CardPK=" & CardPK, ""))%>&SearchText=' + searchValue;
  xmlhttpPost('CAM-Feeds.aspx?'+ qryStr, qryStr, 'ShowSelect', args);
}

function handleSearchKeyDown(e, logixTransNum, transNum, transDate) {
  var keycode = 0;
  
  if (window.event) keycode = window.event.keyCode;
  else if (e) keycode = e.which;
  else return true;
  
  if (keycode == 13) {    
    searchFromDialog(logixTransNum, transNum, transDate);
  } else if (keycode == 27) {
    closeDialog();
  }
}
//-->
</script>
<script type="text/javascript" src="/javascript/jquery.min.js"></script>
<script type="text/javascript" src="/javascript/thickbox.js">
function HR1_onclick() {
}
</script>

<form id="mainform" name="mainform" action="CAM-customer-transactions.aspx">
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
    <div id="controls"<% Sendb(IIf(restrictLinks AndAlso URLtrackBack <> "", " style=""width:115px;""", "")) %>>
      <input type="button" class="regular" id="btnNew" name="btnNew" value="<%Sendb(Copient.PhraseLib.Lookup("term.new", LanguageID)) %>..." onclick="showSelectDialog('', '', '');" />&nbsp;&nbsp;
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
    %>
    <input type="hidden" id="CustomerPK" name="CustomerPK" value="<%Sendb(CustomerPK)%>" />
    <input type="hidden" id="CustPK" name="CustPK" value="<%Sendb(CustomerPK)%>" />
    <input type="hidden" id="CardPK" name="CardPK" value="<%Sendb(CardPK)%>" />
    <%
      If (Request.QueryString("mode") = "summary") Then
        Send("<input type=""hidden"" id=""mode"" name=""mode"" value=""summary"" />")
        Send("<input type=""hidden"" id=""exiturl"" name=""exiturl"" value=""" & URLtrackBack & """ />")
      End If
    %>
    <div id="column">
      <% If (Logix.UserRoles.ViewTransHistory AndAlso CustomerPK > 0) Then%>
        <%
          Send(" <div id=""listbar"">")
          Send("  <div id=""searcher"">")
          Send("   <input type=""text"" style=""font-family:arial;font-size:12px;width:91px;"" id=""transterms"" name=""transterms"" class=""mediumshort"" value=""" & TransTerms & """ onkeydown=""submitTransSearch(event);"" />")
          Send("   <input type=""button"" style=""font-family:arial;font-size:12px;"" id=""btnOffer"" name=""btnOffer"" value=""" & Copient.PhraseLib.Lookup("term.search", LanguageID) & """ onclick=""searchTrans();"" />")
          Send("  </div>")
          Send("  <div id=""paginator"">")
          If (sizeOfData > 0) Then
            If (PageNum > 0) Then
              Send("   <span id=""first""><a href=""CAM-customer-transactions.aspx?CustPK=" & CustomerPK & "&CardPK=" & CardPK & "&redemptionFilter=" & redemptionFilter & "&pagenum=0&transterms=" & TransTerms & "&sortcol=" & SortCol & "&sortdir=" & SortDir & """><b>|</b>◄" & Copient.PhraseLib.Lookup("term.first", LanguageID) & "</a>&nbsp;</span>")
              Send("   <span id=""previous""><a href=""CAM-customer-transactions.aspx?CustPK=" & CustomerPK & "&CardPK=" & CardPK & "&redemptionFilter=" & redemptionFilter & "&pagenum=" & PageNum - 1 & "&transterms=" & TransTerms & "&sortcol=" & SortCol & "&sortdir=" & SortDir & """>◄" & Copient.PhraseLib.Lookup("term.previous", LanguageID) & "</a>&nbsp;</span>")
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
              Send("   <span id=""next""><a href=""CAM-customer-transactions.aspx?CustPK=" & CustomerPK & "&CardPK=" & CardPK & "&redemptionFilter=" & redemptionFilter & "&pagenum=" & PageNum + 1 & "&transterms=" & TransTerms & "&sortcol=" & SortCol & "&sortdir=" & SortDir & """>" & Copient.PhraseLib.Lookup("term.next", LanguageID) & "►</a>&nbsp;</span>")
              Send("   <span id=""last""><a href=""CAM-customer-transactions.aspx?CustPK=" & CustomerPK & "&CardPK=" & CardPK & "&redemptionFilter=" & redemptionFilter & "&pagenum=" & (Math.Ceiling(sizeOfData / linesPerPage) - 1) & "&transterms=" & TransTerms & "&sortcol=" & SortCol & "&sortdir=" & SortDir & """>" & Copient.PhraseLib.Lookup("term.last", LanguageID) & "►<b>|</b></a>&nbsp;</span>")
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
          Send("   </select>")
          Send("  </div>")
          Send(" </div>")
        %>
        <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.transactionhistory", LanguageID)) %>">
          <thead>
            <tr>
              <th align="left" style="width: 15px;">
              </th>
              <th align="left" class="th-button" scope="col" style="text-align: center;">
                <%Sendb(Copient.PhraseLib.Lookup("term.adjust", LanguageID))%>
              </th>
              <th align="left" class="th-datetime" scope="col">
                <a href="<% Sendb(SortUrl & "&sortcol=TransactionDate&sortdir=" & IIf(SortDir = "desc", "asc", "desc")) %>">
                  <% Sendb(Copient.PhraseLib.Lookup("term.date", LanguageID))%>
                </a>
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
                <a href="<% Sendb(SortUrl & "&sortcol=CustomerPrimaryExtId&sortdir=" & IIf(SortDir = "desc", "asc", "desc")) %>">
                  <% Sendb(Copient.PhraseLib.Lookup("term.cardnumber", LanguageID))%>
                </a>
                <%
                  If SortCol = "CustomerPrimaryExtId" Then
                    If SortDir = "asc" Then
                      Sendb("<span class=""sortarrow"">&#9660;</span>")
                    Else
                      Sendb("<span class=""sortarrow"">&#9650;</span>")
                    End If
                  End If
                %>
              </th>
              <th align="left" class="th-id" scope="col">
                <a href="<% Sendb(SortUrl & "&sortcol=ExtLocationCode&sortdir=" & IIf(SortDir = "desc", "asc", "desc")) %>">
                  <% Sendb(Copient.PhraseLib.Lookup("term.store", LanguageID))%>
                </a>
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
                <a href="<% Sendb(SortUrl & "&sortcol=TerminalNum&sortdir=" & IIf(SortDir = "desc", "asc", "desc")) %>">
                  <% Sendb(Copient.PhraseLib.Lookup("term.terminal", LanguageID))%>
                </a>
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
                <a href="<% Sendb(SortUrl & "&sortcol=TransNum&sortdir=" & IIf(SortDir = "desc", "asc", "desc")) %>">
                  <% Sendb(Copient.PhraseLib.Lookup("term.txn", LanguageID) & "#")%>
                </a>
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
                <a href="<% Sendb(SortUrl & "&amp;sortcol=Replayed&amp;sortdir=" & IIf(SortDir = "desc", "asc", "desc")) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.replayed", LanguageID)) %>"><% Sendb("R")%></a>
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
              <th align="right" class="th-amount" scope="col" style="text-align: right;">
                <% Sendb(Copient.PhraseLib.Lookup("term.amount", LanguageID))%>
              </th>
              <th align="right" class="th-redemptions" scope="col" style="text-align:right;width:40px;">
                <% Sendb("Rdms")%>
              </th>
              <% End If%>
            </tr>
          </thead>
          <tbody>
            <%
              Dim transRows As ArrayList
              If (rst.Rows.Count > 0) Then
                transCt = 0
                transRows = GetSubList(rstTrans, startPosition, endPosition)
                If transRows.Count > 0 Then
                  For Each row In transRows
                    transCt += 1
                    LogixTransNum = MyCommon.NZ(row.Item("LogixTransNum"), "").ToString.Trim()
                    Send("<tr>")
                    If MyCommon.NZ(row.Item("RedemptionCount"), 0) = 0 Then
                      Sendb("  <td>&nbsp;</td>")
                    Else
                      Sendb("  <td><a href=""#""><img id=""plus" & transCt & """ src=""/images/plus.png"" style=""cursor:hand;"" onclick=""javascript:showTransDetail(" & transCt & ", this);"" /></a></td>")
                    End If
                    If (Logix.UserRoles.AccessPointsBalances = False) Then
                      DisabledPtsAdj = " disabled=""disabled"""
                    Else
                      DisabledPtsAdj = ""
                    End If
                    Sendb("  <td style=""text-align:center;""><input type=""button"" class=""adjust"" id=""ptsAdj" & LogixTransNum & """ name=""ptsAdj"" value=""P""" & DisabledPtsAdj & " title=""" & Copient.PhraseLib.Lookup("term.adjust", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.points", LanguageID), VbStrConv.Lowercase) & """")
                    Send(" onClick=""javascript:showSelectDialog('" & LogixTransNum & "','" & MyCommon.NZ(row.Item("TransNum"), "") & "','" & MyCommon.NZ(row.Item("TransactionDate"), "1/1/1980") & "');"" /><span style=""margin-left:8px;"">&nbsp;</span></td>")
                    Sendb("  <td>")
                    If MyCommon.NZ(row.Item("TransactionDate"), "1/1/1980") = "1/1/1980" Then
                      TransDateStart = Now.ToString("yyyy-MM-dd 00:00:00")
                      TransDateEnd = Now.ToString("yyyy-MM-dd 23:59:59")
                      Sendb(Copient.PhraseLib.Lookup("term.unknown", LanguageID))
                    Else
                      TransDateStart = Date.Parse(row.Item("TransactionDate")).ToString("yyyy-MM-dd 00:00:00")
                      TransDateEnd = Date.Parse(row.Item("TransactionDate")).ToString("yyyy-MM-dd 23:59:59")
                      Sendb(Format(row.Item("TransactionDate"), "dd MMM yyyy, HH:mm:ss"))
                    End If
                    Sendb("</td>")
                    Send("  <td>" & MyCommon.NZ(row.Item("CustomerPrimaryExtId"), UnknownPhrase) & "</td>")
                    Send("  <td>" & MyCommon.NZ(row.Item("ExtLocationCode"), UnknownPhrase) & "</td>")
                    Send("  <td style=""text-align:center;"">" & MyCommon.NZ(row.Item("TerminalNum"), UnknownPhrase) & "</td>")
                    Send("  <td style=""text-align:center;word-break:break-all;""><span title=""" & MyCommon.NZ(row.Item("LogixTransNum"), UnknownPhrase) & """>" & MyCommon.NZ(row.Item("TransNum"), UnknownPhrase) & "</span></td>")
                    Send("  <td>" & IIf(MyCommon.NZ(row.Item("Replayed"), 0) > 0, "<span title=""" & Copient.PhraseLib.Lookup("term.replayed", LanguageID) & """ style=""color:#dd0000;cursor:default;font-size:10px;font-weight:bold;"">R</span>", "") & "</td>")
                    If redemptionFilter <> 3 Then
                      ' calculate the redemption count and amount for this transaction
                      MyCommon.QueryStr = "select CustomerPrimaryExtId, max(TransDate) as TransactionDate, ExtLocationCode, sum(RedemptionAmount) as RedemptionAmount, " & _
                                          "  sum(RedemptionCount) as RedemptionCount, TerminalNum, TransNum, LogixTransNum " & _
                                          "from TransRedemptionView with (NoLock) where LogixTransNum = '" & LogixTransNum & "' " & _
                                          "group by CustomerPrimaryExtId, TransNum, TerminalNum, ExtLocationCode, LogixTransNum "
                      rst2 = MyCommon.LWH_Select
                      If (rst2.Rows.Count > 0) Then
                        If MyCommon.NZ(row.Item("Replayed"), 0) > 0 AndAlso (MyCommon.NZ(row.Item("RedemptionAmount"), 0) > 0) Then
                          Send("  <td style=""text-align:right;""><span style=""background-color:#dddddd;color:#888888;"">" & MyCommon.NZ(rst2.Rows(0).Item("RedemptionAmount"), UnknownPhrase) & "</span></td>")
                        Else
                          Send("  <td style=""text-align:right;"">" & MyCommon.NZ(rst2.Rows(0).Item("RedemptionAmount"), UnknownPhrase) & "</td>")
                        End If
                        Send("  <td style=""text-align:right;"">" & MyCommon.NZ(rst2.Rows(0).Item("RedemptionCount"), UnknownPhrase) & "</td>")
                        TotalRedeemAmt += MyCommon.NZ(rst2.Rows(0).Item("RedemptionAmount"), 0.0)
                        TotalRedeemCt += MyCommon.NZ(rst2.Rows(0).Item("RedemptionCount"), 0)
                      Else
                        If MyCommon.NZ(row.Item("Replayed"), 0) > 0 AndAlso (MyCommon.NZ(row.Item("RedemptionAmount"), 0) > 0) Then
                          Send("  <td style=""text-align:right;"">0.00</td>")
                        Else
                          Send("  <td style=""text-align:right;"">0.00</td>")
                        End If
                        Send("  <td style=""text-align:right;"">0</td>")
                      End If
                    End If
                    Send("</tr>")

                    ' write detail rows
                    If (MyCommon.NZ(row.Item("RedemptionCount"), 0) = 0) Then
                      'No detail line needed
                    Else
                      Send("<tr id=""trTrans" & transCt & """ style=""display:none;"">")
                      Send("  <td colspan=""10"" style=""padding:0;"">")
                      MyCommon.QueryStr = "select OfferID, RedemptionAmount, RedemptionCount, LogixTransNum, TransDate from TransRedemptionView with (NoLock) " & _
                                          "where LogixTransNum = '" & LogixTransNum & "' order by OfferID asc, TransDate desc"
                      rst2 = MyCommon.LWH_Select
                      If (rst2.Rows.Count > 0) Then
                        transOffers = New StringBuilder(500)
                        transRdmptAmt = New StringBuilder(100)
                        transRdmptCt = New StringBuilder(100)
                        offerPointBuf = New StringBuilder(200)
                        Send("<table style=""width:100%;"">")
                        For Each row2 In rst2.Rows
                          LogixTransNum = MyCommon.NZ(row2.Item("LogixTransNum"), "").ToString.Trim
                          MyCommon.QueryStr = "select IncentiveName as OfferName, ClientOfferID as xid from CPE_Incentives with (NoLock) where IncentiveId = " & MyCommon.NZ(row2.Item("OfferID"), -1) & ";"
                          rst3 = MyCommon.LRT_Select
                          If (rst3.Rows.Count > 0) Then
                            OfferName = MyCommon.NZ(rst3.Rows(0).Item("OfferName"), "")
                            XID = MyCommon.NZ(rst3.Rows(0).Item("xid"), "[" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "]")
                          Else
                            OfferName = "[" & Copient.PhraseLib.Lookup("history.offer-delete", LanguageID) & "]"
                            XID = ""
                          End If
                          MyCommon.QueryStr = "dbo.pa_CustomerOfferHasPointsProgram"
                          MyCommon.Open_LRTsp()
                          MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.Int).Value = MyCommon.NZ(row2.Item("OfferID"), -1)
                          MyCommon.LRTsp.Parameters.Add("@HasPointsProgram", SqlDbType.Bit).Direction = ParameterDirection.Output
                          MyCommon.LRTsp.ExecuteNonQuery()
                          IsPtsOffer = MyCommon.LRTsp.Parameters("@HasPointsProgram").Value
                          MyCommon.Close_LRTsp()
                          If (IsPtsOffer) Then
                            IIf(Logix.UserRoles.AccessPointsBalances = False, DisabledPtsAdj = " disabled=""disabled""", DisabledPtsAdj = "")
                            transOffers.Append("<input type=""button"" class=""adjust"" id=""ptsAdj" & MyCommon.NZ(row2.Item("OfferID"), "") & """ name=""ptsAdj"" value=""P""" & DisabledPtsAdj & " title=""" & Copient.PhraseLib.Lookup("term.adjust", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.points", LanguageID), VbStrConv.Lowercase) & """ ")
                            transOffers.Append("onClick=""javascript:openPopup('CAM-point-adjust.aspx?LogixTransNum=" & LogixTransNum & "&OfferID=" & MyCommon.NZ(row2.Item("OfferID"), "") & "&CustomerPK=" & CustomerPK & IIf(CardPK > 0, "&CardPK=" & CardPK, "") & "');"" /><span style=""margin-left:8px;"">&nbsp;</span>")
                          Else
                            transOffers.Append("<span style=""margin-left:30px;"">&nbsp;</span>")
                          End If
                          transOffers.Append(Copient.PhraseLib.Lookup("term.xid", LanguageID) & ": " & XID & "&nbsp;&nbsp;&nbsp;&nbsp;")
                          transOffers.Append(Copient.PhraseLib.Lookup("term.id", LanguageID) & ": ")
                          transOffers.Append(MyCommon.NZ(row2.Item("OfferID"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)) & "&nbsp;&nbsp;&nbsp;&nbsp;")
                          transOffers.Append(Copient.PhraseLib.Lookup("term.name", LanguageID) & ": " & MyCommon.SplitNonSpacedString(OfferName, 30) & "<br />")
                          transRdmptAmt.Append(MyCommon.NZ(row2.Item("RedemptionAmount"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)) & "<br />")
                          transRdmptCt.Append(MyCommon.NZ(row2.Item("RedemptionCount"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)) & "<br />")
                          Send("  <tr>")
                          Send("    <td class=""transOffers"">" & transOffers.ToString & "</td>")
                          Send("    <td class=""transRdmptAmt"">" & transRdmptAmt.ToString & "</td>")
                          Send("    <td class=""transRdmptCt"" style=""width:40px;"">" & transRdmptCt.ToString & "</td>")
                          Send("  </tr>")
                          transOffers.Remove(0, transOffers.Length)
                          transRdmptAmt.Remove(0, transRdmptAmt.Length)
                          transRdmptCt.Remove(0, transRdmptCt.Length)
                        Next
                        Send("</table>")
                      Else
                        Send("<table style=""width:100%;"">")
                        Send("  <tr><td class=""transOffers"" colspan=""3"">No redemptions found for this transaction</td></tr>")
                        Send("</table>")
                      End If
                      Send("  </td>")
                      Send("</tr>")
                    End If
                  Next
                Else
                  Send("<tr>")
                  Send("  <td colspan=""9"" style=""text-align:center""><i>" & Copient.PhraseLib.Lookup("customer-inquiry.nohistory", LanguageID) & "</i></td>")
                  Send("</tr>")
                End If
              Else
                Send("<tr>")
                Send("  <td colspan=""9"" style=""text-align:center""><i>" & Copient.PhraseLib.Lookup("customer-inquiry.nohistory", LanguageID) & "</i></td>")
                Send("</tr>")
              End If
              MyCommon.Close_LogixWH()
            %>
          </tbody>
        </table>
        <hr class="hidden" />
      <% End If%>
    </div>
    <br clear="all" />
  </div>
</form>

<div id="selectDialog" style="display:none; overflow:auto; position:absolute; top:140px; left:250px; width:500px; height:350px; z-index:300; background-color:#e0e0e0; border:outset 2px #606060;">
</div>

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
</script>
<%
  'If MyCommon.Fetch_SystemOption(75) Then
  '  If (CustomerPK > 0 And Logix.UserRoles.ViewCustomerNotes) Then
  '    Send_Notes(4, CustomerPK, AdminUserID)
  '  End If
  'End If
done:
  Send_FocusScript("mainform", "transterms")
  Send_WrapEnd()
  Send("<div id=""disabledBkgrd"" style=""position:absolute; top:0px; left:0px; right:0px; width:100%; height:100%; background-color:Gray; display:none; z-index:99; filter:alpha(opacity=50); -moz-opacity:.50; opacity:.50;""></div>")
  Send_PageEnd()
  
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  MyCommon = Nothing
  Logix = Nothing
%>
