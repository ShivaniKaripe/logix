<%@ Page Language="vb" Debug="true" CodeFile="../LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: CAM-customer-adjustments.aspx 
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
  Dim MyPoints As New Copient.Points
  Dim Logix As New Copient.LogixInc
  Dim rstResults As DataTable = Nothing
  Dim rst As DataTable
  Dim rstPrograms As DataTable
  Dim row As DataRow
  Dim rst2 As DataTable
  Dim dt As DataTable
  Dim dtAssigned As DataTable
  Dim CustomerPK As Long
  Dim CardPK As Long
  Dim ExtCardID As String = ""
  Dim FullName As String = ""
  Dim TotalRedeemCt As Integer = 0
  Dim TotalRedeemAmt As Double = 0.0
  Dim CustExtID As String = ""
  Dim i As Integer = 0
  Dim transCt As Integer = 0
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
  Dim ProgramID As Integer = 0
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
  Dim OfferTable As Hashtable = Nothing
  Dim OfferList As String = ""
  Dim CgXML As String = ""
  Dim AllCAMCardholdersID As Long = 0
  Dim Employee As Boolean = False
  
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
  Dim SortCol As String = "PP.ProgramID"
  Dim SortDir As String = "desc"
  Dim SortUrl As String = ""
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "CAM-customer-adjustments.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  MyCommon.Open_LogixWH()
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
                          "C.Employee, CE.Address, CE.City, CE.State, CE.Zip, CE.PhoneAsEntered as Phone, CE.email " & _
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
        MyCommon.QueryStr = "select C.CustomerPK, C.CustomerTypeID, C.HHPK, C.FirstName, C.MiddleName, C.LastName, " & _
                            "C.Employee, CE.Address, CE.City, CE.State, CE.Zip, CE.PhoneAsEntered as Phone, CE.email " & _
                            "from Customers C with (NoLock) " & _
                            "left join CustomerExt CE with (NoLock) on CE.CustomerPK=C.CustomerPK " & _
                            "where C.CustomerPK=" & CustomerPK & ";"
      End If
      If (Request.QueryString("searchterms") <> "" And ClientUserID1 = "") Then
        ClientUserID1 = MyCommon.Pad_ExtCardID( MyCommon.Parse_Quotes(Left(Request.QueryString("searchterms"), 26)),2)
                MyCommon.QueryStr = "select C.CustomerPK, C.CustomerTypeID, C.HHPK, C.FirstName, C.MiddleName, C.LastName, " & _
                                    "C.Employee, CE.Address, CE.City, CE.State, CE.Zip, CE.PhoneAsEntered as Phone, CE.email " & _
                                    "from Customers C with (NoLock) " & _
                                    "left join CustomerExt CE with (NoLock) on CE.CustomerPK=C.CustomerPK " & _
                                    "where C.PrimaryExtID='" & MyCryptLib.SQL_StringEncrypt(ClientUserID1) & "' or CE.PhoneDigitsOnly = '" & MyCryptLib.SQL_StringEncrypt(MyCommon.DigitsOnly(Request.QueryString("searchterms"))) & _
                                    "' or CE.email = '" & MyCryptLib.SQL_StringEncrypt(MyCommon.Parse_Quotes(Request.QueryString("searchterms"))) & "' or C.LastName like '%" & MyCommon.Parse_Quotes(Request.QueryString("searchterms")) & "%';"
      End If
    End If
    rstResults = MyCommon.LXS_Select
    
    If (rstResults.Rows.Count = 1) Then
      ' ok we found a primary key for the external id provided
      CustomerPK = rstResults.Rows(0).Item("CustomerPK")
      IsHouseholdID = (MyCommon.NZ(rstResults.Rows(0).Item("CustomerTypeID"), 0) = 1)
      Employee = (MyCommon.NZ(rstResults.Rows(0).Item("Employee"), 0) = 1)
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
  Send_Scripts()
%>

<script type="text/javascript">
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
    var qryStr = 'CAM-customer-adjustments.aspx?search=Search&searchterms=<%Sendb(strTerms)%>&CustPK=<%Sendb(CustomerPK)%><%Sendb(IIf(CardPK > 0, "&CardPK=" & CardPK, ""))%>&offerSearch=Search&transterms=' + transTerms;
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
    
  function expandRow(programID) {
    var trTranElem = document.getElementById("trTrans" + programID);
    var imgElem = document.getElementById("plus" + programID);
    var isOpen = false;
    var args = new Array(programID, 0)
    var qryStr = 'CAMProgramTransactions=1&ProgramID=' + programID + '&CustPK=<%Sendb(CustomerPK)%>&ExtCustID=<%Sendb(MyCommon.Parse_Quotes(ClientUserID1))%>';
    
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
</script>
<%
  Send_HeadEnd()
  Send_BodyBegin(1)
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  If (Not restrictLinks) Then
    Send_Tabs(Logix, 3)
    If CardPK > 0 Then
      Send_Subtabs(Logix, 33, 5, LanguageID, CustomerPK, , CardPK)
    Else
      Send_Subtabs(Logix, 33, 5, LanguageID, CustomerPK)
    End If
  Else
    If CardPK > 0 Then
      Send_Subtabs(Logix, 94, 5, LanguageID, CustomerPK, extraLink, CardPK)
    Else
      Send_Subtabs(Logix, 94, 5, LanguageID, CustomerPK, extraLink)
    End If
  End If
  If (Logix.UserRoles.AccessCustomerInquiry = False) Then
    Send_Denied(1, "perm.customer-access")
    GoTo done
  End If
  
  If (HHCustIdList.Count > 0) Then
    CustExtIdList = "'" & ExtCardID & "'"
    For i = 0 To HHCustIdList.Count - 1
      CustExtIdList += ", '" & MyCommon.Parse_Quotes(HHCustIdList.Item(i).ToString) & "'"
    Next
  Else
    If (Request.QueryString("searchterms") <> "") Then
      ClientUserID1 = MyCommon.Pad_ExtCardID(ClientUserID1,2)
      CustExtID = ClientUserID1
      CustExtIdList = "'" & MyCommon.Parse_Quotes(ClientUserID1) & "'"
    ElseIf (Request.QueryString("editterms") <> "") Then
      CustExtID = Request.QueryString("editterms")
      CustExtIdList = "'" & MyCommon.Parse_Quotes(CustExtID) & "'"
    ElseIf ClientUserID1 <> "" Then
      CustExtIdList = "'" & ClientUserID1 & "'"
    Else
      CustExtIdList = "''"
    End If
  End If
  
  ' Process search request, if applicable 
  TransTerms = Request.QueryString("transterms")
  If (TransTerms <> "") Then
    SearchFilter = " and (PP.ProgramID=" & MyCommon.Extract_Val(TransTerms) & " or PP.ProgramName like '%" & TransTerms & "%' " & _
                   "      or PP.Description like '%" & TransTerms & "%') "
  Else
    SearchFilter = ""
  End If
  
  ' Process sort request, if applicable
  If (Request.QueryString("sortcol") <> "") Then SortCol = Request.QueryString("sortcol")
  If (Request.QueryString("sortdir") <> "") Then
    SortDir = Request.QueryString("sortdir")
  Else
    SortDir = "desc"
  End If
  
  ' find all the offers for which a customer is eligible, then determine which of those offers have points programs
  CgXML = "<customergroups>"
  MyCommon.QueryStr = "select CustomerGroupID from CustomerGroups with (NoLock) where AnyCAMCardholder=1;"
  rst = MyCommon.LRT_Select
  If (rst.Rows.Count > 0) Then
    AllCAMCardholdersID = MyCommon.NZ(rst.Rows(0).Item("CustomerGroupID"), 0)
    CgXML &= "<id>" & AllCAMCardholdersID & "</id>"
  End If
  
  MyCommon.QueryStr = "select CustomerGroupID from GroupMembership with (NoLock) where CustomerPK=" & CustomerPK & " and Deleted=0"
  rst = MyCommon.LXS_Select()
  
  If rst.Rows.Count > 0 Then
    For Each row In rst.Rows
      CgXML &= "<id>" & MyCommon.NZ(row.Item("CustomerGroupID"), "") & "</id>"
    Next
  End If
  CgXML &= "</customergroups>"
  
  MyCommon.QueryStr = "dbo.pa_CAM_CustomerOffersCurrent"
  MyCommon.Open_LRTsp()
  MyCommon.LRTsp.Parameters.Add("@cgXml", SqlDbType.Xml).Value = CgXML
  MyCommon.LRTsp.Parameters.Add("@IsEmployee", SqlDbType.Bit).Value = IIf(Employee, 1, 0)
  If (Request.QueryString("offerterms") <> "") Then
    MyCommon.LRTsp.Parameters.Add("@Filter", SqlDbType.NVarChar, 50).Value = Request.QueryString("offerterms")
  End If
  MyCommon.LRTsp.Parameters.Add("@Favorite", SqlDbType.Bit).Value = 0
  MyCommon.LRTsp.Parameters.Add("@AdminUserID", SqlDbType.Int).Value = AdminUserID
  dtAssigned = MyCommon.LRTsp_select
  MyCommon.Close_LRTsp()
  
  ' restrict the adjustable programs to those offer for which the customer is eligible
  If dtAssigned.Rows.Count > 0 Then
    For Each row In dtAssigned.Rows
      If OfferList <> "" Then OfferList &= ","
      OfferList &= MyCommon.NZ(row.Item("OfferID"), -1)
    Next
    SearchFilter &= " and INC.IncentiveID IN (" & OfferList & ")"
  End If
  
  MyCommon.QueryStr = "select distinct PP.ProgramID, PP.ProgramName, PP.Description from PointsPrograms AS PP with (NoLock) " & _
                      "inner join CPE_IncentivePointsGroups AS IPG with (NoLock) on IPG.ProgramID = PP.ProgramID and IPG.Deleted=0 " & _
                      "inner join CPE_RewardOptions AS RO with (NoLock) on RO.RewardOptionID = IPG.RewardOptionID and RO.Deleted=0 " & _
                      "inner join CPE_Incentives AS INC with (NoLock) on INC.IncentiveID = RO.IncentiveID and INC.Deleted=0 and INC.EngineID=6 " & _
                      "where PP.Deleted=0 " & SearchFilter & " " & _
                      "union " & _
                      "select distinct PP.ProgramID, PP.ProgramName, PP.Description from PointsPrograms AS PP with (NoLock) " & _
                      "inner join CPE_DeliverablePoints AS DPT with (NoLock) on DPT.ProgramID = PP.ProgramID and DPT.Deleted=0 " & _
                      "inner join CPE_Deliverables as DEL with (NoLock) on DEL.OutputID = DPT.PKID and DEL.DeliverableTypeID=8 and DEL.Deleted=0 " & _
                      "inner join CPE_RewardOptions AS RO with (NoLock) on RO.RewardOptionID = DEL.RewardOptionID and RO.Deleted=0 " & _
                      "inner join CPE_Incentives AS INC with (NoLock) on INC.IncentiveID = RO.IncentiveID and INC.Deleted=0 and INC.EngineID=6 " & _
                      "where PP.Deleted=0 " & SearchFilter & " " & _
                      "order by " & SortCol & " " & SortDir & "; "
  rstPrograms = MyCommon.LRT_Select
  sizeOfData = rstPrograms.Rows.Count
  
  PageNum = MyCommon.Extract_Val(Request.QueryString("pagenum"))
  startPosition = PageNum * linesPerPage
  endPosition = IIf(sizeOfData < startPosition + linesPerPage, sizeOfData, startPosition + linesPerPage) - 1
  MorePages = IIf(sizeOfData > startPosition + linesPerPage, True, False)
  
  SortUrl = "CAM-customer-adjustments.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&CardPK=" & CardPK, "") & "&transterms=" & TransTerms & "&pagenum=0"
%>
<form id="mainform" name="mainform" action="CAM-customer-adjustments.aspx">
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
    <%
      If (CustomerPK > 0 And Logix.UserRoles.ViewCustomerNotes) Then
        Send_CustomerNotes(CustomerPK, CardPK)
      End If
    %>
  </div>
</div>
<div id="main">
  <%If (Logix.UserRoles.ViewTransHistory = False) Then
      Send(Copient.PhraseLib.Lookup("error.forbidden", LanguageID))
      Send("</div>")
      Send("</form>")
      GoTo done
    End If
  %>
  <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div><br />")%>
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
    <% If (Logix.UserRoles.AccessAdjustmentsPage AndAlso CustomerPK > 0) Then%>
    <div class="box" id="activity" <%if (customerpk = 0) then sendb(" style=""display: none;""") %>>
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.pointsprograms", LanguageID))%></span></h2>
      <%
        Send(" <div id=""listbar"">")
        Send("  <div id=""paginator"" style=""float:none;text-align:left;width:auto;"">")
        If (sizeOfData > 0) Then
          Send("   <input type=""text"" style=""font-family:arial;font-size:12px;"" id=""transterms"" name=""transterms"" class=""mediumshort"" value=""" & TransTerms & """ onkeydown=""submitTransSearch(event);"" />")
          Send("   <input type=""button"" style=""font-family:arial;font-size:12px;"" id=""btnOffer"" name=""btnOffer"" value=""" & Copient.PhraseLib.Lookup("term.search", LanguageID) & """ onclick=""searchTrans();"" />")
          Send("   <span style=""padding-right:60px;"">&nbsp;</span>")
          If (PageNum > 0) Then
            Send("   <span id=""first""><a href=""CAM-customer-adjustments.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&CardPK=" & CardPK, "") & "&pagenum=0&transterms=" & TransTerms & "&sortcol=" & SortCol & "&sortdir=" & SortDir & """><b>|</b>◄" & Copient.PhraseLib.Lookup("term.first", LanguageID) & "</a>&nbsp;</span>")
            Send("   <span id=""previous""><a href=""CAM-customer-adjustments.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&CardPK=" & CardPK, "") & "&pagenum=" & PageNum - 1 & "&transterms=" & TransTerms & "&sortcol=" & SortCol & "&sortdir=" & SortDir & """>◄" & Copient.PhraseLib.Lookup("term.previous", LanguageID) & "</a>&nbsp;</span>")
          Else
            Send("   <span id=""first""><b>|</b>◄" & Copient.PhraseLib.Lookup("term.first", LanguageID) & "&nbsp;</span>")
            Send("   <span id=""previous"">◄" & Copient.PhraseLib.Lookup("term.previous", LanguageID) & "&nbsp;</span>")
          End If
        Else
          Send("   <input type=""text"" class=""mediumshort"" id=""transterms"" name=""transterms"" onkeydown=""submitTransSearch(event);"" style=""font-family:arial;font-size:12px;"" value=""" & TransTerms & """ />")
          Send("   <input type=""button"" id=""btnTrans"" name=""btnTrans"" onclick=""searchTrans();"" style=""font-family:arial;font-size:12px;"" value=""Search"" />")
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
            Send("   <span id=""next""><a href=""CAM-customer-adjustments.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&CardPK=" & CardPK, "") & "&pagenum=" & PageNum + 1 & "&transterms=" & TransTerms & "&sortcol=" & SortCol & "&sortdir=" & SortDir & """>" & Copient.PhraseLib.Lookup("term.next", LanguageID) & "►</a>&nbsp;</span>")
            Send("   <span id=""last""><a href=""CAM-customer-adjustments.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&CardPK=" & CardPK, "") & "&pagenum=" & (Math.Ceiling(sizeOfData / linesPerPage) - 1) & "&transterms=" & TransTerms & "&sortcol=" & SortCol & "&sortdir=" & SortDir & """>" & Copient.PhraseLib.Lookup("term.last", LanguageID) & "►<b>|</b></a>&nbsp;</span>")
          Else
            Send("   <span id=""next"">" & Copient.PhraseLib.Lookup("term.next", LanguageID) & "►&nbsp;</span>")
            Send("   <span id=""last"">" & Copient.PhraseLib.Lookup("term.last", LanguageID) & "►<b>|</b>&nbsp;</span>")
          End If
        Else
          Send("   <span id=""next"">" & Copient.PhraseLib.Lookup("term.next", LanguageID) & "►&nbsp;</span>")
          Send("   <span id=""last"">" & Copient.PhraseLib.Lookup("term.last", LanguageID) & "►<b>|</b></span><br />")
        End If
        Send("  </div>")
        Send(" </div>")
      %>
      <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.pointsprograms", LanguageID)) %>">
        <thead>
          <tr>
            <th align="left" style="width: 15px;">
            </th>
            <th align="left" class="th-button" scope="col" style="text-align: center;">
              <%Sendb(Copient.PhraseLib.Lookup("term.adjust", LanguageID))%>
            </th>
            <th align="left" class="th-programid" scope="col">
              <a href="<% Sendb(SortUrl & "&sortcol=PP.ProgramID&sortdir=" & IIf(SortDir = "desc", "asc", "desc")) %>">
                <% Sendb(Copient.PhraseLib.Lookup("term.ProgramID", LanguageID))%></a>
              <%If SortCol = "PP.ProgramID" Then
                  If SortDir = "asc" Then
                    Sendb("<span class=""sortarrow"">&#9660;</span>")
                  Else
                    Sendb("<span class=""sortarrow"">&#9650;</span>")
                  End If
                End If%>
            </th>
            <th align="left" class="th-quantity" scope="col">
              <% Sendb(Copient.PhraseLib.Lookup("term.quantity", LanguageID))%>
            </th>
            <th align="left" class="th-name" scope="col">
              <a href="<% Sendb(SortUrl & "&sortcol=PP.ProgramName&sortdir=" & IIf(SortDir = "desc", "asc", "desc")) %>">
                <% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%></a>
              <%If SortCol = "PP.ProgramName" Then
                  If SortDir = "asc" Then
                    Sendb("<span class=""sortarrow"">&#9660;</span>")
                  Else
                    Sendb("<span class=""sortarrow"">&#9650;</span>")
                  End If
                End If%>
            </th>
            <th align="center" class="th-description" scope="col">
              <a href="<% Sendb(SortUrl & "&sortcol=PP.Description&sortdir=" & IIf(SortDir = "desc", "asc", "desc")) %>">
                <% Sendb(Copient.PhraseLib.Lookup("term.description", LanguageID))%></a>
              <%If SortCol = "PP.Description" Then
                  If SortDir = "asc" Then
                    Sendb("<span class=""sortarrow"">&#9660;</span>")
                  Else
                    Sendb("<span class=""sortarrow"">&#9650;</span>")
                  End If
                End If%>
            </th>
          </tr>
        </thead>
        <tbody>
          <%
            Dim transRows As ArrayList
            If (rstPrograms.Rows.Count > 0) Then
              transCt = 0
              transRows = GetSubList(rstPrograms, startPosition, endPosition)
              If transRows.Count > 0 Then
                For Each row In transRows
                  transCt += 1
                  ProgramID = MyCommon.NZ(row.Item("ProgramID"), 0)
                  If ProgramID > 0 Then
                    Send("<tr>")
                    Send("<td><img id=""plus" & ProgramID & """ src=""/images/plus.png"" style=""cursor:hand;"" onclick=""expandRow(" & ProgramID & ");"" /></td>")
                    If (Logix.UserRoles.AccessPointsBalances = False) Then
                      DisabledPtsAdj = " disabled=""disabled"""
                    Else
                      DisabledPtsAdj = ""
                    End If
                    Sendb("   <td style=""text-align:center;""><input type=""button"" class=""adjust"" id=""ptsAdj" & ProgramID & """ name=""ptsAdj"" value=""P""" & DisabledPtsAdj & " title=""" & Copient.PhraseLib.Lookup("term.adjust", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.points", LanguageID), VbStrConv.Lowercase) & """ ")
                    Send("onClick=""javascript:openPopup('CAM-point-adjust-program.aspx?Trans=-1&ProgramID=" & ProgramID & "&CustomerPK=" & CustomerPK & IIf(CardPK > 0, "&CardPK=" & CardPK, "") & "');"" /><span style=""margin-left:8px;""></span></td>")
                    Send("  <td>" & ProgramID & "</td>")
                    Send("  <td style=""text-align:right;"">" & MyPoints.GetBalance(CustomerPK, ProgramID) & "</td>")
                    Send("  <td>" & MyCommon.NZ(row.Item("ProgramName"), "") & "</td>")
                    Send("  <td>" & MyCommon.NZ(row.Item("Description"), "") & "</td>")
                    Send("</tr>")
                    
                    'create the Transactions row
                    Send("<tr id=""trTrans" & ProgramID & """ style=""display:none;background-color:#dddddd;"" >")
                    Send("  <td id=""tdTrans" & ProgramID & """ colspan=""6"">")
                    Send("    <img src=""/images/loadingAnimation.gif"" />")
                    Send("  </td>")
                    Send("</tr>")
                  End If
                Next
              Else
                Send("<tr>")
                Send("  <td colspan=""6"" style=""text-align:center""><i>" & Copient.PhraseLib.Lookup("term.NoPointsPrograms", LanguageID) & "</i></td>")
                Send("</tr>")
              End If
            Else
              Send("<tr>")
              Send("  <td colspan=""6"" style=""text-align:center""><i>" & Copient.PhraseLib.Lookup("term.NoPointsPrograms", LanguageID) & "</i></td>")
              Send("</tr>")
            End If
          %>
        </tbody>
      </table>
      <hr class="hidden" id="HR1" onclick="return HR1_onclick()" />
    </div>
    <% End If%>
  </div>
  <br clear="all" />
</div>
</form>

<script runat="server">
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
done:
  Send_BodyEnd("mainform", "transterms")
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  MyCommon.Close_LogixWH()
  MyCommon = Nothing
  Logix = Nothing
%>
