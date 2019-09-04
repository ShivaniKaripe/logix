<%@ Page Language="vb" Debug="true" CodeFile="../LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: LMG-rejections.aspx 
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
  Dim MyCam As New Copient.CAM
  Dim MyPoints As New Copient.Points
  Dim Logix As New Copient.LogixInc
  Dim dt As System.Data.DataTable
  Dim dt2 As System.Data.DataTable
  Dim dtemp As System.Data.DataTable
  Dim row As System.Data.DataRow
  Dim rst, rst2 As System.Data.DataTable
  Dim rstStores As System.Data.DataTable
  Dim rstOffers As System.Data.DataTable
  Dim Shaded As String = "shaded"
  Dim PageNum As Integer = 0
  Dim MorePages As Boolean
  Dim linesPerPage As Integer = 10
  Dim sizeOfData As Integer
  Dim i As Integer = 0
  Dim j As Integer = 0
  
  ' Variables used in searching
  Dim SearchSource As Integer = -1
  Dim SearchPromoEngine As string = ""
  Dim SearchClientLocationCode As String = ""
  Dim SearchCustomer As String = ""
  Dim SearchAirmile As String = ""
  Dim SearchAirmileExtCardID As String = ""
  Dim SearchBoxID As Integer = -1
  Dim SearchOfferID As Integer = -1
  Dim SearchIssuanceDate As Date = "1/1/1900"
  Dim SearchIssuanceDateEnd As Date = "1/1/1900"
  Dim myUrl As String = "/logix/CAM/LMG-rejections.aspx"
  Dim startVal As Integer = 0
  Dim endVal As Integer = 0
  Dim SearchString As String = ""
  Dim SortString As String = ""
  Dim SortText As String = "ClientLocationCode"
  Dim SortDirection As String
  
  ' Variables used when retrying, editing and saving records
  Dim RetryItem As Integer = 0
  Dim EditItem As Integer = 0
  Dim DeleteItem As Integer = 0
  Dim SaveItem As Integer = 0
  Dim SaveClientLocationCode As String = ""
  Dim SaveBoxID As Integer = 0
  Dim SaveTransactionNumber As String = "0"
  Dim SaveOfferID As Integer = 0
  Dim SaveRewardQty As Integer = 0
  Dim SaveIssuanceDate As Date = "1/1/1900 00:00:00"
  Dim SaveROID As Integer = 0
  Dim SaveProgramID As Integer = 0
  Dim SaveLocationID As Integer = 0
  Dim UserName As String = ""
  Dim FullName As String = ""
  Dim CurrentIssuanceTable As String = ""
  Dim IssuanceTableName As String = ""
  
  ' Variables used when copying to Points and PointsHistory
  Dim PromoVarID As Integer = 0
  Dim CustomerPK As Integer = 0
  Dim CardPK As Integer = 0
  Dim PrimaryExtID As String = ""
  Dim LogixTransNum As String = "0"
  Dim SourceTypeID As Integer = 0
  Dim CustomerTypeID As Integer = 2
  Dim Amount As Decimal = 0.0
  Dim ProgramID As Integer = 0
  Dim ROID As Integer = 0
  Dim OfferID As Integer = 0
  Dim OfferStartDate As String = ""
  Dim OfferEndDate As String = ""
  Dim ClientLocationCode As String = ""
  Dim LocationID As Integer = 0
  Dim BoxID As Integer = 0
  Dim TransactionNumber As String = "0"
  Dim IssuanceDate As String = ""
  Dim LMGRejectionPKID As Long
  Dim AirmileMemberID As String = ""
  Dim IsUSAM As Boolean = False
  
  Dim WhereString As String = ""
  Dim AffectedCount As Integer = 0
  Dim PointsBal As Long = 0
  Dim SoftLimit As Long = 0
  Dim HardLimit As Long = 0
  Dim ValidAdjustment As Boolean = False
  Dim TempValInt As Integer = 0
  Dim TempValString As String = ""
  
  Dim infoMessage As String = ""
  Dim statusMessage As String = ""
  
  Dim Handheld As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "LMG-rejections.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixEX()
  MyCommon.Open_LogixXS()
  MyCommon.Open_LogixWH()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  
  Send_HeadBegin("term.rejections")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
%>
<style type="text/css">
#rejections {
  overflow-x: auto;
  padding-bottom: 3px;
  width: 739px;
  }
* html #rejections {
  padding-bottom: 20px;
  }
.editline {
  background-color: #dddddd;
  border-radius: 0 0 6px 6px;
  -moz-border-radius: 0 0 6px 6px;
  }
.editline td input {
  font-size: 12px;
  }
.editline td select {
  font-size: 12px;
  }
  
/* SEARCH BAR */
#searchbar {
  background-color: #d0caea;
  font-family: Arial, Helvetica, sans-serif;
  height: 70px;
  margin: 2px 0 3px 0;
  padding: 2px 3px 0 3px;
  text-align: center;
  width: 733px;
  }
* html #searchbar {
  width: 739px;
  }
#searchbar input {
  color: #aaaaaa;
  font-size: 12px;
  height: 14px;
  margin: 0;
  width: 64px;
  }
* html #searchbar input {
  height: 20px;
  }
#searchbar select {
  font-size: 12px;
  margin: 0;
  }
#search {
  color: #000000 !important;
  }
#SearchBoxID, #SearchOfferID {
  width: 54px !important;
  }
</style>
<%
  Send_Scripts()
%>
<script type="text/javascript">
  function clearInput(input) {
    if ((document.getElementById(input).value == '<% Sendb(Copient.PhraseLib.Lookup("term.location", LanguageID)) %>') || (document.getElementById(input).value == '<% Sendb(Copient.PhraseLib.Lookup("term.terminal", LanguageID)) %>') || (document.getElementById(input).value == '<% Sendb(Copient.PhraseLib.Lookup("term.offerid", LanguageID)) %>') || (document.getElementById(input).value == '<% Sendb(Copient.PhraseLib.Lookup("term.startdate", LanguageID)) %>') || (document.getElementById(input).value == '<% Sendb(Copient.PhraseLib.Lookup("term.enddate", LanguageID)) %>') || (document.getElementById(input).value == '<% Sendb(Copient.PhraseLib.Lookup("term.customer", LanguageID)) %>') || (document.getElementById(input).value == '<% Sendb(Copient.PhraseLib.Lookup("cpesettings.153", LanguageID).Substring(8,12)) %>')) {
      document.getElementById(input).value = '';
      document.getElementById(input).style.color = '#000000';
    }
  }
  
  function resetInputs() {
    if (document.getElementById('SearchClientLocationCodeInput').value == '') {
      document.getElementById('SearchClientLocationCodeInput').value = '<% Sendb(Copient.PhraseLib.Lookup("term.location", LanguageID)) %>';
      document.getElementById('SearchClientLocationCodeInput').style.color = '#aaaaaa';
    }
    if (document.getElementById('SearchBoxIDInput').value == '') {
      document.getElementById('SearchBoxIDInput').value = '<% Sendb(Copient.PhraseLib.Lookup("term.terminal", LanguageID)) %>';
      document.getElementById('SearchBoxIDInput').style.color = '#aaaaaa';
    }
    if (document.getElementById('SearchOfferIDInput').value == '') {
      document.getElementById('SearchOfferIDInput').value = '<% Sendb(Copient.PhraseLib.Lookup("term.offerid", LanguageID)) %>';
      document.getElementById('SearchOfferIDInput').style.color = '#aaaaaa';
    }
    if (document.getElementById('SearchIssuanceDateInput').value == '') {
      document.getElementById('SearchIssuanceDateInput').value = '<% Sendb(Copient.PhraseLib.Lookup("term.startdate", LanguageID)) %>';
      document.getElementById('SearchIssuanceDateInput').style.color = '#aaaaaa';
    }
    if (document.getElementById('SearchIssuanceDateEndInput').value == '') {
      document.getElementById('SearchIssuanceDateEndInput').value = '<% Sendb(Copient.PhraseLib.Lookup("term.enddate", LanguageID)) %>';
      document.getElementById('SearchIssuanceDateEndInput').style.color = '#aaaaaa';
    }
    if (document.getElementById('SearchCustomerInput').value == '') {
      document.getElementById('SearchCustomerInput').value = '<% Sendb(Copient.PhraseLib.Lookup("term.customer", LanguageID)) %>';
      document.getElementById('SearchCustomerInput').style.color = '#aaaaaa';
    }
    if (document.getElementById('SearchAirmileInput').value == '') {
      document.getElementById('SearchAirmileInput').value = '<% Sendb(Copient.PhraseLib.Lookup("cpesettings.153", LanguageID).Substring(8,12)) %>';
      document.getElementById('SearchAirmileInput').style.color = '#aaaaaa';
    }
  }
  
  function submitSearch() {
    if (document.getElementById('SearchSourceInput').value == '-1') {
      document.getElementById('SearchSourceInput').value = '';
      document.getElementById('SearchSource').value = '';
    }
    if (document.getElementById('SearchClientLocationCodeInput').value == '<% Sendb(Copient.PhraseLib.Lookup("term.location", LanguageID)) %>') {
      document.getElementById('SearchClientLocationCodeInput').value = '';
      document.getElementById('SearchClientLocationCode').value = '';
    }
    if (document.getElementById('SearchBoxIDInput').value == '<% Sendb(Copient.PhraseLib.Lookup("term.terminal", LanguageID)) %>') {
      document.getElementById('SearchBoxIDInput').value = '';
      document.getElementById('SearchBoxID').value = '';
    }
    if (document.getElementById('SearchOfferIDInput').value == '<% Sendb(Copient.PhraseLib.Lookup("term.offerid", LanguageID)) %>') {
      document.getElementById('SearchOfferIDInput').value = '';
      document.getElementById('SearchOfferID').value = '';
    }
    if (document.getElementById('SearchIssuanceDateInput').value == '<% Sendb(Copient.PhraseLib.Lookup("term.startdate", LanguageID)) %>') {
      document.getElementById('SearchIssuanceDateInput').value = '';
      document.getElementById('SearchIssuanceDate').value = '';
    }
    if (document.getElementById('SearchIssuanceDateEndInput').value == '<% Sendb(Copient.PhraseLib.Lookup("term.enddate", LanguageID)) %>') {
      document.getElementById('SearchIssuanceDateEndInput').value = '';
      document.getElementById('SearchIssuanceDateEnd').value = '';
    }
    if (document.getElementById('SearchCustomerInput').value == '<% Sendb(Copient.PhraseLib.Lookup("term.Customer", LanguageID)) %>') {
      document.getElementById('SearchCustomerInput').value = '';
      document.getElementById('SearchCustomer').value = '';
    }
    if (document.getElementById('SearchAirmileInput').value == '<% Sendb(Copient.PhraseLib.Lookup("cpesettings.153", LanguageID).Substring(8,12)) %>') {
      document.getElementById('SearchAirmileInput').value = '';
      document.getElementById('SearchAirmile').value = '';
    }
    if (document.getElementById('SearchPromoEngineInput').value == 'all') {
      document.getElementById('SearchPromoEngineInput').value = '';
      document.getElementById('SearchPromoEngine').value = '';
    }
    document.searchform.submit();
  }
  
  function editItem(PKID) {
    var editLine = document.getElementById("editline" + PKID);
    if (editLine.style.display == 'none') {
      editLine.style.display = '';
    } else {
      editLine.style.display = 'none';
    }
  }

  function retryItem(PKID, sizeOfData) {
    var tokenValues = [];
    var msg = '';

    if (PKID == '-1') {
      msg = '<%Sendb(Copient.PhraseLib.Lookup("lmg.ConfirmRetryAll", LanguageID))%>';
      tokenValues[0] = sizeOfData;
      msg = detokenizeString(msg, tokenValues);
      if (confirm(msg)) {
        document.getElementById("RetryItem").value = '-1';
        document.mainform.action = "LMG-rejections.aspx"
        document.mainform.submit();
      }
    } else {
      msg = '<%Sendb(Copient.PhraseLib.Lookup("lmg.ConfirmRetry", LanguageID))%>';
      if (confirm(msg)) {
        document.getElementById("RetryItem").value = PKID;
        document.mainform.action = "LMG-rejections.aspx"
        document.mainform.submit();
      }
    }
  }

  function deleteItem(PKID, sizeOfData) {
    var tokenValues = [];
    var msg = '';

    if (PKID == '-1') {
      msg = '<%Sendb(Copient.PhraseLib.Lookup("lmg.ConfirmDeleteAll", LanguageID))%>';
      tokenValues[0] = sizeOfData;
      msg = detokenizeString(msg, tokenValues);
      if (confirm(msg)) {
        document.getElementById("DeleteItem").value = '-1';
        document.mainform.action = "LMG-rejections.aspx"
        document.mainform.submit();
      }
    } else {
      msg = '<%Sendb(Copient.PhraseLib.Lookup("lmg.ConfirmDelete", LanguageID))%>';
      if (confirm(msg)) {
        document.getElementById("DeleteItem").value = PKID;
        document.mainform.action = "LMG-rejections.aspx"
        document.mainform.submit();
      }
    }
  }
  
  function saveItem(PKID, sizeOfData, dateAlert) {
    var OfferID = '';
    var IssuanceDate = ''
    var confirmationText = '';
    var tokenValues = [];
    
    IssuanceDate = document.getElementById('IssuanceDate' + PKID).value;
    OfferID = document.getElementById('OfferID' + PKID).value;
    
    // First set the text that will appear in the confirmation popup
    if (PKID == '-1') {
      confirmationText = '<%Sendb(Copient.PhraseLib.Lookup("lmg.ConfirmDeleteAll", LanguageID))%>';
      tokenValues[0] = sizeOfData;
      confirmationText = detokenizeString(msg, tokenValues);
    } else {
      confirmationText = '<%Sendb(Copient.PhraseLib.Lookup("lmg.ConfirmDelete", LanguageID))%>';
    }
    if (dateAlert != '') {
      confirmationText = confirmationText + '\n\n' + dateAlert;
    }
    // Then, if confirmed, set the hidden inputs and submit the form
    if (confirm(confirmationText)) {
      document.getElementById("SaveItem").value = PKID;
      document.getElementById("SaveClientLocationCode").value = document.getElementById("ClientLocationCode" + PKID).value;
      document.getElementById("SaveBoxID").value = document.getElementById("BoxID" + PKID).value;
      document.getElementById("SaveTransactionNumber").value = document.getElementById("TransactionNumber" + PKID).value;
      document.getElementById("SaveOfferID").value = document.getElementById("OfferID" + PKID).value;
      document.getElementById("SaveRewardQty").value = document.getElementById("RewardQty" + PKID).value;
      document.getElementById("SaveIssuanceDate").value = document.getElementById("IssuanceDate" + PKID).value;
      document.getElementById('save' + PKID).removeAttribute('disabled');
      document.mainform.action = "LMG-rejections.aspx"
      document.mainform.submit();
    } else {
      document.getElementById('save' + PKID).removeAttribute('disabled');
    }
  }
  
  function handleSave(PKID, sizeOfData) {
    var IssuanceDate = ''
    var OfferID = ''
    var confirmationText = '';
    
    IssuanceDate = document.getElementById('IssuanceDate' + PKID).value;
    OfferID = document.getElementById('OfferID' + PKID).value;
    if ((OfferID == '') || (isNaN(OfferID))) { 
      OfferID = '0';
    }
    document.getElementById('save' + PKID).disabled = 'true';
    xmlhttpPost('/logix/XMLFeeds.aspx', 'HandleLMGSave=1&CheckDate=' + IssuanceDate + '&OfferID=' + OfferID, PKID, sizeOfData);
  }
  
  function xmlhttpPost(strURL, qryStr, PKID, sizeOfData) {
    var xmlHttpReq = false;
    var self = this;
    
    if (window.XMLHttpRequest) { // Mozilla/Safari
      self.xmlHttpReq = new XMLHttpRequest();
    } else if (window.ActiveXObject) { // IE
      self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
    }
    strURL += "?" + qryStr;
    self.xmlHttpReq.open('POST', strURL, true);
    self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    self.xmlHttpReq.send(qryStr);
    self.xmlHttpReq.onreadystatechange = function() {
      if (self.xmlHttpReq !=null && self.xmlHttpReq.readyState == 4) {
        respTxt = self.xmlHttpReq.responseText;
        saveItem(PKID, sizeOfData, respTxt);
      }
    }
  }
</script>
<%
  Send_HeadEnd()
  Send_BodyBegin(1)
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  Send_Tabs(Logix, 8)
  Send_Subtabs(Logix, 8, 5)
  
  
  If Request.QueryString("StatusMessage") <> "" Then
    statusMessage = MyCommon.Parse_Quotes(Request.QueryString("StatusMessage"))
  End If
  
  MyCommon.QueryStr = "select Username, Firstname, Lastname from AdminUsers with (NoLock) where AdminUserID=" & AdminUserID & ";"
  dt = MyCommon.LRT_Select
  If dt.Rows.Count > 0 Then
    UserName = MyCommon.NZ(dt.Rows(0).Item("UserName"), Copient.PhraseLib.Lookup("term.unknown", LanguageID))
    FullName = MyCommon.NZ(dt.Rows(0).Item("FirstName"), "") & " " & MyCommon.NZ(dt.Rows(0).Item("LastName"), "")
  End If
  
  If (Logix.UserRoles.AccessLMGRejections = False) Then
    Send_Denied(1, "perm.lmg-access")
    GoTo done
  End If
  
  MyCommon.QueryStr = "select top(1) TableName from IssuanceTables order by PKID desc;"
  dt = MyCommon.LEX_Select
  If dt.Rows.Count > 0 Then
    CurrentIssuanceTable = MyCommon.NZ(dt.Rows(0).Item("TableName"), "")
  End If
  
  ' RETRY, SAVE and DELETE ACTIONS
  If (Request.QueryString("RetryItem") <> "") OrElse (Request.QueryString("SaveItem") <> "") OrElse (Request.QueryString("DeleteItem") <> "") Then
    ' 1) Find and assign to variables all the data that will need to be logged, copied, etc.
    CustomerTypeID = 2  '(We're processing only CAM cards so CustomerTypeID should *always* be 2)
    SourceTypeID = 12   '(SourceType should now always be stored as 12, which means "re-sent")
    
    MyCommon.QueryStr = "select ClientLocationCode, BoxID, TransactionNumber, PrimaryExtID, OfferID, ROID, RewardQty, IssuanceDate, ProgramID, LogixTransNum, SourceTypeID, AirmileMemberID " & _
                        "from LMGRejection with (NoLock) "
    If (Request.QueryString("RetryItem") <> "") Then
      MyCommon.QueryStr &= "where PKID=" & MyCommon.Extract_Val(Request.QueryString("RetryItem")) & ";"
    ElseIf (Request.QueryString("SaveItem") <> "") Then
      MyCommon.QueryStr &= "where PKID=" & MyCommon.Extract_Val(Request.QueryString("SaveItem")) & ";"
    ElseIf (Request.QueryString("DeleteItem") <> "") Then
      MyCommon.QueryStr &= "where PKID=" & MyCommon.Extract_Val(Request.QueryString("DeleteItem")) & ";"
    End If
    dt = MyCommon.LEX_Select()
    If dt.Rows.Count > 0 Then
      ClientLocationCode = MyCommon.NZ(dt.Rows(0).Item("ClientLocationCode"), "")
      BoxID = MyCommon.NZ(dt.Rows(0).Item("BoxID"), 0)
      TransactionNumber = MyCommon.NZ(dt.Rows(0).Item("TransactionNumber"), 0)
      PrimaryExtID = MyCommon.NZ(dt.Rows(0).Item("PrimaryExtID"), "")
      OfferID = MyCommon.NZ(dt.Rows(0).Item("OfferID"), 0)
      ROID = MyCommon.NZ(dt.Rows(0).Item("ROID"), 0)
      Amount = MyCommon.NZ(dt.Rows(0).Item("RewardQty"), 0)
      IssuanceDate = MyCommon.NZ(dt.Rows(0).Item("IssuanceDate"), 0)
      ProgramID = MyCommon.NZ(dt.Rows(0).Item("ProgramID"), 0)
      LogixTransNum = MyCommon.NZ(dt.Rows(0).Item("LogixTransNum"), "0")
      AirmileMemberID = MyCommon.NZ(dt.Rows(0).Item("AirmileMemberID"), "0")
    End If
    PrimaryExtID = MyCommon.Pad_ExtCardID(PrimaryExtID,2)

    
    MyCommon.QueryStr = "select C.CustomerPK, CT.CustTypeID from CardIDs as C with (NoLock) " & _
                        "left join CardTypes as CT on CT.CardTypeID=C.CardTypeID " & _
                        "where ExtCardID='" & MyCryptLib.SQL_StringEncrypt(PrimaryExtID) & "';"
    dt = MyCommon.LXS_Select()
    If dt.Rows.Count > 0 Then
      CustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)
      CustomerTypeID = MyCommon.NZ(dt.Rows(0).Item("CustTypeID"), 0)
    End If
    MyCommon.QueryStr = "select PromoVarID from PointsPrograms with (NoLock) where ProgramID=" & ProgramID & ";"
    dt = MyCommon.LRT_Select()
    If dt.Rows.Count > 0 Then
      PromoVarID = MyCommon.NZ(dt.Rows(0).Item("PromoVarID"), 0)
    End If
    MyCommon.QueryStr = "select LocationID from Locations with (NoLock) where ExtLocationCode='" & ClientLocationCode & "';"
    dt = MyCommon.LRT_Select()
    If dt.Rows.Count > 0 Then
      LocationID = MyCommon.NZ(dt.Rows(0).Item("LocationID"), 0)
    End If
    If Request.QueryString("SearchSource") <> "" Then
      If MyCommon.Extract_Val(Request.QueryString("SearchSource")) >= 0 Then
        SearchSource = MyCommon.Extract_Val(Request.QueryString("SearchSource"))
      End If
    End If
    If Request.QueryString("SearchClientLocationCode") <> "" Then
      SearchClientLocationCode = MyCommon.Parse_Quotes(Request.QueryString("SearchClientLocationCode"))
    End If
    If Request.QueryString("SearchBoxID") <> "" Then
      SearchBoxID = MyCommon.Extract_Val(Request.QueryString("SearchBoxID"))
    End If
    If Request.QueryString("SearchOfferID") <> "" Then
      SearchOfferID = MyCommon.Extract_Val(Request.QueryString("SearchOfferID"))
    End If
    If Request.QueryString("SearchIssuanceDate") <> "" Then
      If IsDate(Request.QueryString("SearchIssuanceDate")) Then
        If Date.Parse(Request.QueryString("SearchIssuanceDate")) > "1/1/1900" Then
          SearchIssuanceDate = Date.Parse(Request.QueryString("SearchIssuanceDate"))
        End If
      End If
    End If
    If (Request.QueryString("RetryItem") <> "") Then
      LMGRejectionPKID = MyCommon.Extract_Val(Request.QueryString("RetryItem"))
    ElseIf (Request.QueryString("SaveItem") <> "") Then
      LMGRejectionPKID = MyCommon.Extract_Val(Request.QueryString("SaveItem"))
    ElseIf (Request.QueryString("DeleteItem") <> "") Then
      LMGRejectionPKID = MyCommon.Extract_Val(Request.QueryString("DeleteItem"))
    End If
    
    '2) Build up a "where" string based on search criteria (if any), to be used in subsequent queries
    If LMGRejectionPKID = -1 Then
      WhereString = " where PKID>=0"
      WhereString &= IIf(SearchSource >= 0, " and SourceTypeID=" & SearchSource, "")
      WhereString &= IIf(SearchClientLocationCode <> "", " and ClientLocationCode='" & SearchClientLocationCode & "'", "")
      WhereString &= IIf(SearchBoxID >= 0, " and BoxID=" & SearchBoxID, "")
      WhereString &= IIf(SearchOfferID >= 0, " and OfferID=" & SearchOfferID, "")
      WhereString &= IIf(SearchIssuanceDate > "1/1/1900", " and IssuanceDate>='" & SearchIssuanceDate & " 00:00:00' and IssuanceDate<='" & SearchIssuanceDate & " 23:59:59'", "")
      MyCommon.QueryStr &= WhereString & ";"
    Else
      WhereString = " where PKID=" & LMGRejectionPKID
    End If
    
    '3) Get the "save" values, if any
    SaveClientLocationCode = Request.QueryString("SaveClientLocationCode")
    SaveBoxID = MyCommon.Extract_Val(Request.QueryString("SaveBoxID"))
    SaveTransactionNumber = MyCommon.Extract_Val(Request.QueryString("SaveTransactionNumber"))
    SaveOfferID = MyCommon.Extract_Val(Request.QueryString("SaveOfferID"))
    SaveRewardQty = MyCommon.Extract_Val(Request.QueryString("SaveRewardQty"))
    If IsDate(Request.QueryString("SaveIssuanceDate")) Then
      SaveIssuanceDate = Date.Parse(Request.QueryString("SaveIssuanceDate"))
    End If
    If SaveOfferID > 0 Then
      MyCommon.QueryStr = "select I.IncentiveID, RO.RewardOptionID, DP.ProgramID " & _
                          "from CPE_Incentives as I with (NoLock) " & _
                          "left join CPE_RewardOptions as RO on RO.IncentiveID=I.IncentiveID " & _
                          "left join CPE_DeliverablePoints as DP on DP.RewardOptionID=RO.RewardOptionID " & _
                          "where I.IncentiveID=" & SaveOfferID & ";"
      dt = MyCommon.LRT_Select
      If dt.Rows.Count > 0 Then
        SaveROID = MyCommon.NZ(dt.Rows(0).Item("RewardOptionID"), 0)
        SaveProgramID = MyCommon.NZ(dt.Rows(0).Item("ProgramID"), 0)
      End If
    End If
    If SaveClientLocationCode <> "" Then
      MyCommon.QueryStr = "select LocationID from Locations with (NoLock) where ExtLocationCode='" & SaveClientLocationCode & "';"
      dt = MyCommon.LRT_Select
      If dt.Rows.Count > 0 Then
        SaveLocationID = MyCommon.NZ(dt.Rows(0).Item("LocationID"), 0)
      End If
    End If
    
    'Now we have all our data, so proceed to the retry/save or delete operations:
    If (Request.QueryString("RetryItem") <> "") OrElse (Request.QueryString("SaveItem") <> "") Then
      '4) Retry routine ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
      MyCommon.QueryStr = "select * from LMGRejection with (NoLock)" & WhereString & ";"
      rst = MyCommon.LEX_Select
      AffectedCount = rst.Rows.Count
      If AffectedCount > 0 Then
        'Determine if the adjustment is valid
        If Request.QueryString("RetryItem") <> "" Then
          ValidAdjustment = True
        ElseIf Request.QueryString("SaveItem") <> "" Then
          If LMGRejectionPKID = -1 Then
            ValidAdjustment = True
          Else
            PointsBal = MyPoints.GetBalance(CustomerPK, ProgramID)
            SoftLimit = MyCam.GetMaxAdjustment(OfferID, ProgramID)
            HardLimit = MyCam.GetOfferPointLimit(OfferID, ProgramID)
            If SaveRewardQty = 0 Then
              infoMessage &= Copient.PhraseLib.Lookup("lmg.ErrorZeroAdjustment", LanguageID)
            ElseIf (PointsBal + SaveRewardQty > HardLimit) Then
              infoMessage &= Copient.PhraseLib.Detokenize("lmg.ErrorLimitExceeded", LanguageID, HardLimit, (HardLimit - PointsBal))  'Offers point limit of {0} exceeded. Maximum adjustment is {1} points.
            ElseIf (PointsBal + SaveRewardQty < 0) Then
              infoMessage &= Copient.PhraseLib.Detokenize("lmg.ErrorNegativeBalance", LanguageID, SaveRewardQty, (PointsBal + SaveRewardQty), (-PointsBal))  'This adjustment ({0}) would cause a negative points balance ({1}).  Maximum adjustment is {2} points.
            ElseIf (Math.Abs(SaveRewardQty) > SoftLimit) Then
              infoMessage &= Copient.PhraseLib.Lookup("lmg.WarningLimit", LanguageID)
              ValidAdjustment = True
            Else
              ValidAdjustment = True
            End If
          End If
        End If
        If ValidAdjustment Then
          For Each row In rst.Rows
            '4b) Look up per-line data not included in LMGRejection:
            MyCommon.QueryStr = "select C.CustomerPK, CT.CustTypeID, CE.AirmileMemberID from CardIDs as C with (NoLock) " & _
                                "left join CardTypes as CT on CT.CardTypeID=C.CardTypeID " & _
                                "left join CustomerExt as CE on C.CustomerPK=CE.CustomerPK " & _
                                "where ExtCardID='" & MyCryptLib.SQL_StringEncrypt(row.Item("PrimaryExtID").ToString()) & "' and C.CardTypeID=" & MyCommon.NZ(row.Item("CardTypeID"), 2) & ";"
            rst2 = MyCommon.LXS_Select()
            If rst2.Rows.Count > 0 Then
              CustomerPK = MyCommon.NZ(rst2.Rows(0).Item("CustomerPK"), 0)
              CustomerTypeID = MyCommon.NZ(rst2.Rows(0).Item("CustTypeID"), 0)
              AirmileMemberID = MyCommon.NZ(rst2.Rows(0).Item("AirmileMemberID"), "")
            End If
            '4c) Insert a copy of the LMGRejection record(s) into the current issuance table:
            MyCommon.QueryStr = "insert into " & CurrentIssuanceTable & " with (RowLock)" & _
                                " (ClientLocationCode, LocationID, BoxID, TransactionNumber, PrimaryExtID, OfferID, ROID," & _
                                "  IssuanceDate, DeliverableType, Void, RewardQty, ProgramID, ChargebackVendorID, CashierNum," & _
                                "  ChargebackDept, ExtCRMInterface, SourceTypeID, PromoEngine, CustomerTypeID, RewardValue," & _
                                "  UniqueID, Expiration, Gross, Net, ExceedManualThreshold, LogixTransNum, VendorCouponCode, ManufacturerCoupon, AirmileMemberID, CardTypeID)" & _
                                " values " & _
                                " ('" & IIf(Request.QueryString("SaveClientLocationCode") <> "", SaveClientLocationCode, MyCommon.NZ(row.Item("ClientLocationCode"), "")) & "', " & _
                                "  " & IIf(Request.QueryString("SaveClientLocationCode") <> "", SaveLocationID, MyCommon.NZ(row.Item("LocationID"), 0)) & ", " & _
                                "  " & IIf(Request.QueryString("SaveBoxID") <> "", SaveBoxID, MyCommon.NZ(row.Item("BoxID"), 0)) & ", " & _
                                "  " & IIf(Request.QueryString("SaveTransactionNumber") <> "", SaveTransactionNumber, MyCommon.NZ(row.Item("TransactionNumber"), 0)) & ", " & _
                                "  '" & MyCommon.NZ(row.Item("PrimaryExtID"), "") & "', " & _
                                "  " & IIf(Request.QueryString("SaveOfferID") <> "", SaveOfferID, MyCommon.NZ(row.Item("OfferID"), 0)) & ", " & _
                                "  " & IIf(Request.QueryString("SaveOfferID") <> "", SaveROID, MyCommon.NZ(row.Item("ROID"), 0)) & ", " & _
                                "  '" & IIf(Request.QueryString("SaveIssuanceDate") <> "", SaveIssuanceDate, MyCommon.NZ(row.Item("IssuanceDate"), "1/1/1900")) & "', " & _
                                "  " & MyCommon.NZ(row.Item("DeliverableType"), 0) & ", " & _
                                "  " & MyCommon.NZ(row.Item("Void"), 0) & ", " & _
                                "  " & IIf(Request.QueryString("SaveRewardQty") <> "", SaveRewardQty, MyCommon.NZ(row.Item("RewardQty"), 0)) & ", " & _
                                "  " & IIf(Request.QueryString("SaveOfferID") <> "", SaveProgramID, MyCommon.NZ(row.Item("ProgramID"), 0)) & ", " & _
                                "  " & MyCommon.NZ(row.Item("ChargebackVendorID"), 0) & ", " & _
                                "  " & MyCommon.NZ(row.Item("CashierNum"), 0) & ", " & _
                                "  " & MyCommon.NZ(row.Item("ChargebackDept"), 0) & ", " & _
                                "  '" & MyCommon.NZ(row.Item("ExtCRMInterface"), "") & "', " & _
                                "  12, " & _
                                "  '" & MyCommon.NZ(row.Item("PromoEngine"), "CPE") & "', " & _
                                "  " & CustomerTypeID & ", " & _
                                "  " & MyCommon.NZ(row.Item("RewardValue"), "NULL") & ", " & _
                                "  " & MyCommon.NZ(row.Item("UniqueID"), "NULL") & ", " & _
                                "  " & IIf(IsDBNull(row.Item("Expiration")), "NULL", "'" & row.Item("Expiration") & "'") & ", " & _
                                "  " & MyCommon.NZ(row.Item("Gross"), "NULL") & ", " & _
                                "  " & MyCommon.NZ(row.Item("Net"), "NULL") & ", " & _
                                "  " & IIf(row.Item("ExceedManualThreshold"), 1, 0) & ", " & _
                                "  " & MyCommon.NZ(row.Item("LogixTransNum"), 0) & ", " & _
                                "  '" & MyCommon.NZ(row.Item("VendorCouponCode"), "") & "', " & _
                                "  " & MyCommon.NZ(row.Item("ManufacturerCoupon"), 0) & ", " & _
                                " '" & AirmileMemberID & "', '" & MyCommon.NZ(row.Item("CardTypeID"), 2) & "');"
            MyCommon.LEX_Execute()
            '4d) Insert these data into TransRedemption on WH
            MyCommon.QueryStr = "insert into TransRedemption with (RowLock) " & _
                                " (OfferID, ExtLocationCode, CustomerPrimaryExtID, RedemptionCount, RedemptionAmount, TransDate, TerminalNum, TransNum, LogixTransNum, CustomerTypeID) " & _
                                " values " & _
                                " (" & IIf(Request.QueryString("SaveOfferID") <> "", SaveOfferID, MyCommon.NZ(row.Item("OfferID"), 0)) & ", '" & _
                                "  " & IIf(Request.QueryString("SaveClientLocationCode") <> "", SaveClientLocationCode, MyCommon.NZ(row.Item("ClientLocationCode"), "")) & "', '" & _
                                "  " & MyCommon.NZ(row.Item("PrimaryExtID"), "") & "', " & _
                                "  1, " & _
                                "  0.00, '" & _
                                "  " & IIf(Request.QueryString("SaveIssuanceDate") <> "", SaveIssuanceDate, MyCommon.NZ(row.Item("IssuanceDate"), "1/1/1900")) & "', " & _
                                "  " & IIf(Request.QueryString("SaveBoxID") <> "", SaveBoxID, MyCommon.NZ(row.Item("BoxID"), 0)) & ", " & _
                                "  " & IIf(Request.QueryString("SaveTransactionNumber") <> "", SaveTransactionNumber, MyCommon.NZ(row.Item("TransactionNumber"), 0)) & ", " & _
                                "  " & MyCommon.NZ(row.Item("LogixTransNum"), 0) & ", " & _
                                "  " & CustomerTypeID & ");"
            MyCommon.LWH_Execute()
            '4e) Insert into PointsHistory on XS with copies of the LMGRejection data
            MyCommon.QueryStr = "insert into PointsHistory with (RowLock) " & _
                                " (ProgramID, CustomerPK, AdjAmount, EarnedUnderROID, LastUpdate, LastServerID, LocationID, LogixTransNum, SourceTypeID, POSTimeStamp) " & _
                                " values " & _
                                " (" & IIf(Request.QueryString("SaveOfferID") <> "", SaveProgramID, MyCommon.NZ(row.Item("ProgramID"), 0)) & ", " & _
                                "  " & CustomerPK & ", " & _
                                "  " & IIf(Request.QueryString("SaveRewardQty") <> "", SaveRewardQty, MyCommon.NZ(row.Item("RewardQty"), 0)) & ", " & _
                                "  " & IIf(Request.QueryString("SaveOfferID") <> "", SaveROID, MyCommon.NZ(row.Item("ROID"), 0)) & ", " & _
                                "  getdate(), " & _
                                "  -10, " & _
                                "  -10, " & _
                                "  " & MyCommon.NZ(row.Item("LogixTransNum"), 0) & ", " & _
                                "  " & MyCommon.NZ(row.Item("SourceTypeID"), 0) & ", " & _ 
                                "  getdate());"
            MyCommon.LXS_Execute()
            '4f) Also put data into Points; if already in Points, do an update instead of an insert.
            MyCommon.QueryStr = "select * from Points with (NoLock) where CustomerPK=" & CustomerPK & " and ProgramID=" & ProgramID & ";"
            rst2 = MyCommon.LXS_Select()
            If rst2.Rows.Count > 0 Then
              MyCommon.QueryStr = "update Points set Amount=Amount+" & Amount & " " & _
                                  "where CustomerPK=" & CustomerPK & " and " & _
                                  "ProgramID=" & ProgramID & ";"
              MyCommon.LXS_Execute()
            Else
              MyCommon.QueryStr = "insert into Points with (RowLock) (PromoVarID, CustomerPK, Amount, ProgramID) " & _
                                  " values " & _
                                  " (" & PromoVarID & ", " & _
                                  "  " & CustomerPK & ", " & _
                                  "  " & Amount & ", " & _
                                  "  " & ProgramID & ");"
              MyCommon.LXS_Execute()
            End If
            If Request.QueryString("RetryItem") <> "" Then
              LogAction("Retry", Now(), UserName, FullName, MyCommon.NZ(row.Item("ClientLocationCode"), ""), MyCommon.NZ(row.Item("BoxID"), 0), MyCommon.NZ(row.Item("TransactionNumber"), 0), MyCommon.NZ(row.Item("PrimaryExtID"), ""), MyCommon.NZ(row.Item("OfferID"), 0), MyCommon.NZ(row.Item("IssuanceDate"), "1/1/1900"), MyCommon.NZ(row.Item("ProgramID"), 0), Amount)
            ElseIf Request.QueryString("SaveItem") <> "" Then
              LogAction("Save", Now(), UserName, FullName, MyCommon.NZ(row.Item("ClientLocationCode"), ""), MyCommon.NZ(row.Item("BoxID"), 0), MyCommon.NZ(row.Item("TransactionNumber"), 0), MyCommon.NZ(row.Item("PrimaryExtID"), ""), MyCommon.NZ(row.Item("OfferID"), 0), MyCommon.NZ(row.Item("IssuanceDate"), "1/1/1900"), MyCommon.NZ(row.Item("ProgramID"), 0), Amount)
            End If
          Next
        End If
        '4g) Finally, delete the original record(s)
        MyCommon.QueryStr = "delete from LMGRejection " & WhereString & ";"
        MyCommon.LEX_Execute()
        Response.Redirect("/logix/CAM/LMG-rejections.aspx?StatusMessage=" & AffectedCount & " record(s) retried.")
      Else
        Response.Redirect("/logix/CAM/LMG-rejections.aspx?StatusMessage=No records retried.")
      End If
      
    ElseIf (Request.QueryString("DeleteItem") <> "") Then
      '5) Delete routine ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
      DeleteItem = Request.QueryString("DeleteItem")
      If LMGRejectionPKID > -1 Then
        AffectedCount = 1
        LogAction("Delete", Now(), UserName, FullName, ClientLocationCode, BoxID, TransactionNumber, PrimaryExtID, OfferID, IssuanceDate, ProgramID, Amount)
      Else
        MyCommon.QueryStr = "select * from LMGRejection" & WhereString & ";"
        rst = MyCommon.LEX_Select()
        AffectedCount = rst.Rows.Count
        If AffectedCount > 0 Then
          For Each row In rst.Rows
            LogAction("Delete", Now(), UserName, FullName, MyCommon.NZ(row.Item("ClientLocationCode"), ""), MyCommon.NZ(row.Item("BoxID"), 0), MyCommon.NZ(row.Item("TransactionNumber"), 0), MyCommon.NZ(row.Item("PrimaryExtID"), ""), MyCommon.NZ(row.Item("OfferID"), 0), MyCommon.NZ(row.Item("IssuanceDate"), "1/1/1900"), MyCommon.NZ(row.Item("ProgramID"), 0), Amount)
          Next
        End If
      End If
      MyCommon.QueryStr = "delete from LMGRejection" & WhereString & ";"
      MyCommon.LEX_Execute()
      Response.Redirect("/logix/CAM/LMG-rejections.aspx?StatusMessage=" & AffectedCount & " record(s) deleted.")
    End If
  End If
  
  PageNum = Request.QueryString("pagenum")
  If PageNum < 0 Then
    PageNum = 0
  End If
  MorePages = False
    
  If (Request.QueryString("SortText") <> "") Then
    SortText = Request.QueryString("SortText")
  End If
  If (Request.QueryString("pagenum") = "") Then
    If (Request.QueryString("SortDirection") = "ASC") Then
      SortDirection = "DESC"
    ElseIf (Request.QueryString("SortDirection") = "DESC") Then
      SortDirection = "ASC"
    Else
      SortDirection = "DESC"
    End If
  Else
    SortDirection = Request.QueryString("SortDirection")
  End If
  SortString = "&amp;SortText=" & SortText & "&amp;SortDirection=" & SortDirection
  
  'Gather inbound search terms
  If (Request.QueryString("SearchSourceInput") <> "") OrElse (Request.QueryString("SearchClientLocationCodeInput") <> "") OrElse (Request.QueryString("SearchBoxIDInput") <> "") OrElse (Request.QueryString("SearchOfferIDInput") <> "") OrElse (Request.QueryString("SearchIssuanceDateInput") <> "") OrELse (Request.QueryString("SearchIssuanceDateEndInput") <> "") OrElse (Request.QueryString("SearchCustomerInput") <> "") OrElse (Request.QueryString("SearchAirmileInput") <> "") OrElse (Request.QueryString("SearchPromoEngineInput") <> "") Then
    SearchString = "&amp;"
    If Request.QueryString("SearchSourceInput") <> "" Then
      If MyCommon.Extract_Val(Request.QueryString("SearchSourceInput")) >= 0 Then
        SearchSource = MyCommon.Extract_Val(Request.QueryString("SearchSourceInput"))
        SearchString &= "SearchSourceInput=" & SearchSource & "&amp;"
      End If
    End If
    If Request.QueryString("SearchClientLocationCodeInput") <> "" Then
      SearchClientLocationCode = MyCommon.Parse_Quotes(Request.QueryString("SearchClientLocationCodeInput"))
      SearchString &= "SearchClientLocationCodeInput=" & SearchClientLocationCode & "&amp;"
    End If
    If Request.QueryString("SearchBoxIDInput") <> "" Then
      SearchBoxID = MyCommon.Extract_Val(Request.QueryString("SearchBoxIDInput"))
      SearchString &= "SearchBoxIDInput=" & SearchBoxID & "&amp;"
    End If
    If Request.QueryString("SearchOfferIDInput") <> "" Then
      SearchOfferID = MyCommon.Extract_Val(Request.QueryString("SearchOfferIDInput"))
      SearchString &= "SearchOfferIDInput=" & SearchOfferID & "&amp;"
    End If
    If Request.QueryString("SearchIssuanceDateInput") <> "" Then
      If IsDate(Request.QueryString("SearchIssuanceDateInput")) Then
        If Date.Parse(DateValue(Request.QueryString("SearchIssuanceDateInput"))) > "1/1/1900" Then
          SearchIssuanceDate = Date.Parse(DateValue(Request.QueryString("SearchIssuanceDateInput")))
          SearchString &= "SearchIssuanceDateInput=" & SearchIssuanceDate & "&amp;"
          'If Request.QueryString("SearchIssuanceDateEnd") = "" Then
          '  SearchIssuanceDateEnd = SearchIssuanceDate
          'End If
        End If
      End If
    End If
    If Request.QueryString("SearchIssuanceDateEndInput") <> "" Then
      If IsDate(Request.QueryString("SearchIssuanceDateEndInput")) Then
        If Date.Parse(DateValue(Request.QueryString("SearchIssuanceDateEndInput"))) > "1/1/1900" Then
          SearchIssuanceDateEnd = Date.Parse(DateValue(Request.QueryString("SearchIssuanceDateEndInput")))
          SearchString &= "SearchIssuanceDateEndInput=" & SearchIssuanceDateEnd & "&amp;"
        End If
      End If
    End If
    If Request.QueryString("SearchCustomerInput") <> "" Then
      SearchCustomer = MyCommon.Parse_Quotes(Request.QueryString("SearchCustomerInput"))
	  'SearchCustomer = MyCommon.Extract_Val(Request.QueryString("SearchCustomerInput"))
      SearchString &= "SearchCustomerInput=" & SearchCustomer & "&amp;"
    End If
    If Request.QueryString("SearchAirmileInput") <> "" Then
	  'SearchString &= "SearchAirmile=" & SearchAirmile & "&amp;"
      SearchAirmile = MyCommon.Parse_Quotes(Request.QueryString("SearchAirmileInput"))
	  SearchString &= "SearchAirmileInput=" & SearchAirmile & "&amp;"
     'get the ExtCardID associated with the searched AirmileMemberID, include those in the LMG Query
      MyCommon.QueryStr = "select c.ExtCardID from CardIDs c with (NoLock) " & _
                                            "join CustomerExt e with (NoLock) on c.CustomerPK=e.CustomerPK " & _
                                            "where e.AirmileMemberID='" & SearchAirmile & "' and c.CardTypeID=1;"
      dtemp = MyCommon.LXS_Select
      If(dtemp.Rows.Count > 0) then
        SearchAirmileExtCardID = IIf(IsDBNull(dtemp.Rows(0).Item("ExtCardID")), "unknown", MyCryptLib.SQL_StringDecrypt(dtemp.Rows(0).Item("ExtCardID").ToString()))
        SearchString &= "SearchAirmileExtCardID=" & SearchAirmileExtCardID & "&amp;"
      End If
    End If
    If Request.QueryString("SearchPromoEngineInput") <> "" Then
      If MyCommon.Parse_Quotes(Request.QueryString("SearchPromoEngineInput")) <> "all" Then
        SearchPromoEngine = MyCommon.Parse_Quotes(Request.QueryString("SearchPromoEngineInput"))
        SearchString &= "SearchPromoEngineInput=" & SearchPromoEngine & "&amp;"
      End If
    End If
    MyCommon.QueryStr = "select PKID, ClientLocationCode, BoxID, TransactionNumber, PrimaryExtID, OfferID, ROID, IssuanceDate, DeliverableType, Void, " & _
                        "RewardQty, ProgramID, ChargebackVendorID, CashierNum, ChargebackDept, ExtCRMInterface, LMG.SourceTypeID, ST.Description as Source, PromoEngine, LMG.AirmileMemberID " & _
                        "from LMGRejection as LMG with (NoLock) " & _
                        "left join SourceTypes as ST with (NoLock) on ST.SourceTypeID=LMG.SourceTypeID " & _
                        "where PKID>=0"
    MyCommon.QueryStr &= IIf(SearchSource >= 0, " and LMG.SourceTypeID=" & SearchSource, "")
    MyCommon.QueryStr &= IIf((SearchPromoEngine <> "all" AndAlso SearchPromoEngine <> ""), " and PromoEngine='" & SearchPromoEngine & "'", "")
    MyCommon.QueryStr &= IIf(SearchClientLocationCode <> "", " and ClientLocationCode='" & SearchClientLocationCode & "'", "")
    MyCommon.QueryStr &= IIf(SearchBoxID >= 0, " and BoxID=" & SearchBoxID, "")
    MyCommon.QueryStr &= IIf(SearchOfferID >= 0, " and OfferID=" & SearchOfferID, "")
    MyCommon.QueryStr &= IIf(SearchIssuanceDate <> "1/1/1900", " and IssuanceDate>='" & SearchIssuanceDate & " 00:00:00'  ", "")
    MyCommon.QueryStr &= iif(SearchIssuanceDateEnd <> "1/1/1900", " and IssuanceDate<='" & SearchIssuanceDateEnd & " 23:59:59'", "")
    MyCommon.QueryStr &= IIf(SearchCustomer <> "", " and PrimaryExtID='" & SearchCustomer & "'", "")
    MyCommon.QueryStr &= IIf(SearchAirmile <> "", " and (LMG.AirmileMemberID='" & SearchAirmile & "' or PrimaryExtID='" & SearchAirmileExtCardID & "')", "")
  Else
    'No search terms, just query
    MyCommon.QueryStr = "select PKID, ClientLocationCode, BoxID, TransactionNumber, PrimaryExtID, OfferID, ROID, IssuanceDate, DeliverableType, Void, " & _
                        "RewardQty, ProgramID, ChargebackVendorID, CashierNum, ChargebackDept, ExtCRMInterface, LMG.SourceTypeID, ST.Description as Source, PromoEngine, LMG.AirmileMemberID " & _
                        "from LMGRejection as LMG with (NoLock) " & _
                        "left join SourceTypes as ST with (NoLock) on ST.SourceTypeID=LMG.SourceTypeID "
  End If
  If(Logix.UserRoles.ViewEditUSAMRejections AndAlso Logix.UserRoles.ViewEditCAMRejections) 
  MyCommon.QueryStr &= " order by " & SortText & " " & SortDirection & ";"
  Else If (Logix.UserRoles.ViewEditUSAMRejections)
    MyCommon.QueryStr &= " where PromoEngine='CPE' order by " & SortText & " " & SortDirection & ";"
  Else If (Logix.UserRoles.ViewEditCAMRejections)
    MyCommon.QueryStr &= " where PromoEngine='CAM' order by " & SortText & " " & SortDirection & ";"
  Else 
    MyCommon.QueryStr &= " where PromoEngine='' order by " & SortText & " " & SortDirection & ";"
  End If
  'infomessage = MyCommon.QueryStr
  dt = MyCommon.LEX_Select
  sizeOfData = dt.Rows.Count
  i = linesPerPage * PageNum
  startVal = linesPerPage * PageNum
  endVal = linesPerPage * PageNum + linesPerPage
  If startVal = 0 Then
    startVal = 1
  Else
    startVal += 1
  End If
  If endVal > sizeOfData Then endVal = sizeOfData
%>
<div id="intro">
  <h1 id="title">
    <% Sendb(Copient.PhraseLib.Lookup("term.rejections", LanguageID))%>
  </h1>
  <div id="controls">
  </div>
</div>
<div id="main">
  <form action="LMG-rejections.aspx" id="searchform" name="searchform">
    <%
      If (infoMessage <> "") Then
        Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
      End If
      If (statusMessage <> "") Then
        Send("<div id=""statusbar"" class=""green-background"">" & statusMessage & "</div>")
      End If
    %>
    <div id="searchbar">
      <!-- source, location, terminal, offer, date -->
      <select id="SearchSourceInput" name="SearchSourceInput" style="width:150px;">
        <%
          Send("<option value=""-1"">" & Copient.PhraseLib.Lookup("term.all", LanguageID) & " " & Copient.PhraseLib.Lookup("term.sources", LanguageID) & "</option>")
          If(Logix.UserRoles.ViewEditCAMRejections AndAlso Logix.UserRoles.ViewEditUSAMRejections) Then
          MyCommon.QueryStr = "select SourceTypeID, Description from SourceTypes with (NoLock)"
		  Else If (Logix.UserRoles.ViewEditCAMRejections) Then
			MyCommon.QueryStr = "select SourceTypeID, Description from SourceTypes with (NoLock) where (ProgramType Like '%CAM%' or ProgramType is NULL)"
		  Else If (Logix.UserRoles.ViewEditUSAMRejections) Then
			MyCommon.QueryStr = "select SourceTypeID, Description from SourceTypes with (NoLock) where (ProgramType Like '%USAM%' or ProgramType is Null)"
      Else 
			MyCommon.QueryStr = "select SourceTypeID, Description from SourceTypes with (NoLock) where ProgramType is Null"
		  End If
          rst = MyCommon.LEX_Select()
          If rst.Rows.Count > 0 Then
            For Each row In rst.Rows
              Send("<option value=""" & MyCommon.NZ(row.Item("SourceTypeID"), 0) & """" & IIf(SearchSource = MyCommon.NZ(row.Item("SourceTypeID"), 0), " selected=""selected""", "") & ">" & _
                  IIf(MyCommon.NZ(row.Item("Description"), "") <> "", row.Item("Description"), "&nbsp;") & "</option>")
            Next
          End If
        %>
      </select>
      <input type="text" id="SearchClientLocationCodeInput" name="SearchClientLocationCodeInput" title="<% Sendb(Copient.PhraseLib.Lookup("term.location", LanguageID))%>"<% Sendb(IIf(SearchClientLocationCode <> "", " value=""" & SearchClientLocationCode & """ style=""color:#000000;""", " value=""" & Copient.PhraseLib.Lookup("term.location", LanguageID) & """"))%> onclick="javascript:clearInput(this.id);" onblur="javascript:resetInputs();" />
      <input type="text" id="SearchBoxIDInput" name="SearchBoxIDInput" title="<% Sendb(Copient.PhraseLib.Lookup("term.terminal", LanguageID))%>"<% Sendb(IIf(SearchBoxID >= 0, " value=""" & SearchBoxID & """ style=""color:#000000""", " value=""" & Copient.PhraseLib.Lookup("term.terminal", LanguageID) & """"))%> onclick="javascript:clearInput(this.id);" onblur="javascript:resetInputs();" />
      <input type="text" id="SearchOfferIDInput" name="SearchOfferIDInput" title="<% Sendb(Copient.PhraseLib.Lookup("term.offerid", LanguageID))%>"<% Sendb(IIf(SearchOfferID >= 0, " value=""" & SearchOfferID & """ style=""color:#000000;""", " value=""" & Copient.PhraseLib.Lookup("term.offerid", LanguageID) & """"))%> onclick="javascript:clearInput(this.id);" onblur="javascript:resetInputs();" />
      <input type="text" id="SearchIssuanceDateInput" name="SearchIssuanceDateInput" title="<% Sendb(Copient.PhraseLib.Lookup("term.startdate", LanguageID))%>"<% Sendb(IIf(SearchIssuanceDate > "1/1/1900", " value=""" & SearchIssuanceDate & """ style=""color:#000000;""", " value=""" & IIf(Logix.UserRoles.ViewEditUSAMRejections, Copient.PhraseLib.Lookup("term.startdate", LanguageID), Copient.PhraseLib.Lookup("term.date", LanguageID)) & """"))%> onclick="javascript:clearInput(this.id);" onblur="javascript:resetInputs();" />
      <input type="text" id="SearchIssuanceDateEndInput" name="SearchIssuanceDateEndInput" title="<% Sendb(Copient.PhraseLib.Lookup("term.enddate", LanguageID))%>"<% Sendb(IIf(Logix.UserRoles.ViewEditUSAMRejections, IIf(SearchIssuanceDateEnd > "1/1/1900", " value=""" & SearchIssuanceDateEnd & """ style=""color:#000000;""", " value=""" & Copient.PhraseLib.Lookup("term.enddate", LanguageID) & """"), " style=""display:none"" "))%> onclick="javascript:clearInput(this.id);" onblur="javascript:resetInputs();" />
      <input type="text" id="SearchCustomerInput" name="SearchCustomerInput" title="<% Sendb(Copient.PhraseLib.Lookup("term.customer", LanguageID))%>"<% Sendb(IIf(Logix.UserRoles.ViewEditUSAMRejections, IIf(SearchCustomer <> "", " value=""" & SearchCustomer & """ style=""color:#000000;""", " value=""" & Copient.PhraseLib.Lookup("term.customer", LanguageID) & """"), " style=""display:none"" "))%> onclick="javascript:clearInput(this.id);" onblur="javascript:resetInputs();" />
      <input type="text" id="SearchAirmileInput" name="SearchAirmileInput" title="<% Sendb(Copient.PhraseLib.Lookup("cpesettings.153", LanguageID).Substring(8,12))%>"<% Sendb(IIf(Logix.UserRoles.ViewEditUSAMRejections, IIf(SearchAirmile <> "", " value=""" & SearchAirmile & """ style=""color:#000000;""", " value=""" & Copient.PhraseLib.Lookup("cpesettings.153", LanguageID).Substring(8,12) & """"), " style=""display:none"" "))%> onclick="javascript:clearInput(this.id);" onblur="javascript:resetInputs();" />
      <input type="button" id="search" name="search" style="height:20px;font-size:11px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.search", LanguageID))%>" onclick="javascript:submitSearch();" />
      <br />
      <div style="margin-top:4px;" align="left">
          <% 'Sendb("&#160;&#160;")%>
          <select id="SearchPromoEngineInput" name="SearchPromoEngineInput" <% Sendb(IIf((Logix.UserRoles.ViewEditUSAMRejections AndAlso Logix.UserRoles.ViewEditCAMRejections), _
            " style=""width:150px;"" ", " style=""display:none"" ")) %> >
            <%
              Send("<option value=""all"">" & Copient.PhraseLib.Lookup("term.all", LanguageID) & " " & Copient.PhraseLib.Lookup("term.airmiles", LanguageID) & "</option>")
              If (Logix.UserRoles.ViewEditUSAMRejections AndAlso Logix.UserRoles.ViewEditCAMRejections) Then
                Send("<option value=""CAM"""  & IIf(SearchPromoEngine = "CAM" , " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.canadianairmiles", LanguageID) & "</option>")
                Send("<option value=""CPE"""  & IIf(SearchPromoEngine = "CPE", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.us-airmiles", LanguageID) & "</option>")
              ElseIf (Not (Logix.UserRoles.ViewEditUSAMRejections)  AndAlso Logix.UserRoles.ViewEditCAMRejections) Then
                Send("<option value=""CAM"" selected=""selected"">" & Copient.PhraseLib.Lookup("term.canadianairmiles", LanguageID) & "</option>")
              ElseIf (Logix.UserRoles.ViewEditUSAMRejections AndAlso Not (Logix.UserRoles.ViewEditCAMRejections)) Then
                Send("<option value=""CPE"" selected=""selected"">" & Copient.PhraseLib.Lookup("term.us-airmiles", LanguageID) & "</option>")
              'Else
              '  Send("<option value=""500"" selected=""selected"">" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</option>")
              End If
            %>
          </select>
          <br/>
      </div>
      <div style="margin-top:4px;">
        <%
          If (PageNum > 0) Then
            Send("   <span id=""first""><a id=""firstPageLink"" href=""" & myUrl & "?pagenum=0" & SearchString & SortString & """><b>|</b>◄" & Copient.PhraseLib.Lookup("term.first", LanguageID) & "</a></span>&nbsp;")
            Send("   <span id=""previous""><a id=""previousPageLink"" href=""" & myUrl & "?pagenum=" & PageNum - 1 & SearchString & SortString & """>◄" & Copient.PhraseLib.Lookup("term.previous", LanguageID) & "</a></span>")
          Else
            Send("   <span id=""first""><b>|</b>◄" & Copient.PhraseLib.Lookup("term.first", LanguageID) & "</span>&nbsp;")
            Send("   <span id=""previous"">◄" & Copient.PhraseLib.Lookup("term.previous", LanguageID) & "</span>")
          End If
          If sizeOfData = 0 Then
            Send("   &nbsp;[ " & Copient.PhraseLib.Lookup("term.noresults", LanguageID) & " ]&nbsp;")
          Else
            Send("   &nbsp;[ <b>" & startVal & "</b> - <b>" & endVal & "</b> " & Copient.PhraseLib.Lookup("term.of", LanguageID) & " <b>" & sizeOfData & "</b> ]&nbsp;")
          End If
          If (sizeOfData > endVal) Then
            Send("   <span id=""next""><a id=""nextPageLink"" href=""" & myUrl & "?pagenum=" & PageNum + 1 & SearchString & SortString & """>" & Copient.PhraseLib.Lookup("term.next", LanguageID) & "►</a></span>&nbsp;")
            Send("   <span id=""last""><a id=""lastPageLink"" href=""" & myUrl & "?pagenum=" & (Math.Ceiling(sizeOfData / linesPerPage) - 1) & SearchString & SortString & """>" & Copient.PhraseLib.Lookup("term.last", LanguageID) & "►<b>|</b></a></span><br />")
          Else
            Send("   <span id=""next"">" & Copient.PhraseLib.Lookup("term.next", LanguageID) & "►</span>&nbsp;")
            Send("   <span id=""last"">" & Copient.PhraseLib.Lookup("term.last", LanguageID) & "►<b>|</b></span><br />")
          End If
        %>
      </div>
    </div>
  </form>
  <form action="LMG-rejections.aspx" id="mainform" name="mainform">
    <div id="column">
      <div id="rejections">
        <input type="hidden" id="RetryItem" name="RetryItem" value="" />
        <input type="hidden" id="DeleteItem" name="DeleteItem" value="" />
        <input type="hidden" id="SaveItem" name="SaveItem" value="" />
        <input type="hidden" id="SaveClientLocationCode" name="SaveClientLocationCode" value="" />
        <input type="hidden" id="SaveBoxID" name="SaveBoxID" value="" />
        <input type="hidden" id="SaveTransactionNumber" name="SaveTransactionNumber" value="" />
        <input type="hidden" id="SaveOfferID" name="SaveOfferID" value="" />
        <input type="hidden" id="SaveRewardQty" name="SaveRewardQty" value="" />
        <input type="hidden" id="SaveIssuanceDate" name="SaveIssuanceDate" value="" />
        <input type="hidden" id="ActivePage" name="ActivePage" value="" />
        <input type="hidden" id="SearchSource" name="SearchSource" value="<% Sendb(SearchSource) %>" />
        <input type="hidden" id="SearchPromoEngine" name="SearchPromoEngine" value="<% Sendb(SearchPromoEngine) %>" />
        <input type="hidden" id="SearchClientLocationCode" name="SearchClientLocationCode" value="<% Sendb(SearchClientLocationCode) %>" />
        <input type="hidden" id="SearchBoxID" name="SearchBoxID" value="<% Sendb(SearchBoxID) %>" />
        <input type="hidden" id="SearchOfferID" name="SearchOfferID" value="<% Sendb(SearchOfferID) %>" />
        <input type="hidden" id="SearchIssuanceDate" name="SearchIssuanceDate" value="<% Sendb(SearchIssuanceDate) %>" />
        <input type="hidden" id="SearchIssuanceDateEnd" name="SearchIssuanceDateEnd" value="<% Sendb(SearchIssuanceDateEnd) %>" />
        <input type="hidden" id="SearchCustomer" name="SearchCustomer" value="<% Sendb(SearchCustomer) %>" />
        <input type="hidden" id="SearchAirmile" name="SearchAirmile" value="<% Sendb(SearchAirmile) %>" />
        <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.rejections", LanguageID)) %>" style="width:100%;">
          <thead>
            <tr>
              <th style="text-align:center; width:88px; min-width:88px;">
                <% Sendb(Copient.PhraseLib.Lookup("term.action", LanguageID))%>
              </th>
              <th style="text-align:left; width:85px; min-width:85px;" scope="col">
                <a href="/logix/CAM/LMG-rejections.aspx?<% Sendb(SearchString) %>&amp;SortText=SourceTypeID&amp;SortDirection=<% Sendb(SortDirection) %>">
                  <% Sendb(Copient.PhraseLib.Lookup("term.source", LanguageID))%>
                </a>
                <%
                  If SortText = "SourceTypeID" Then
                    If SortDirection = "ASC" Then
                      Sendb("<span class=""sortarrow"">&#9660;</span>")
                    Else
                      Sendb("<span class=""sortarrow"">&#9650;</span>")
                    End If
                  Else
                  End If
                %>
              </th>
              <th style="text-align:left; width:70px; min-width:70px;" scope="col">
                <a href="/logix/CAM/LMG-rejections.aspx?<% Sendb(SearchString) %>&amp;SortText=ClientLocationCode&amp;SortDirection=<% Sendb(SortDirection) %>">
                  <% Sendb(Copient.PhraseLib.Lookup("term.location", LanguageID))%>
                </a>
                <%
                  If SortText = "ClientLocationCode" Then
                    If SortDirection = "ASC" Then
                      Sendb("<span class=""sortarrow"">&#9660;</span>")
                    Else
                      Sendb("<span class=""sortarrow"">&#9650;</span>")
                    End If
                  Else
                  End If
                %>
              </th>
              <th style="text-align:left; width:48px; min-width:48px;" scope="col">
                <a href="/logix/CAM/LMG-rejections.aspx?<% Sendb(SearchString) %>&amp;SortText=BoxID&amp;SortDirection=<% Sendb(SortDirection) %>">
                  <% Sendb(Left(Copient.PhraseLib.Lookup("term.terminal", LanguageID), 4))%>
                </a>
                <%
                  If SortText = "BoxID" Then
                    If SortDirection = "ASC" Then
                      Sendb("<span class=""sortarrow"">&#9660;</span>")
                    Else
                      Sendb("<span class=""sortarrow"">&#9650;</span>")
                    End If
                  Else
                  End If
                %>
              </th>
              <th style="text-align:left; width:55px; min-width:55px;" scope="col">
                <a href="/logix/CAM/LMG-rejections.aspx?<% Sendb(SearchString) %>&amp;SortText=TransactionNumber&amp;SortDirection=<% Sendb(SortDirection) %>">
                  <% Sendb(Left(Copient.PhraseLib.Lookup("term.transaction", LanguageID), 5))%>
                </a>
                <%
                  If SortText = "TransactionNumber" Then
                    If SortDirection = "ASC" Then
                      Sendb("<span class=""sortarrow"">&#9660;</span>")
                    Else
                      Sendb("<span class=""sortarrow"">&#9650;</span>")
                    End If
                  Else
                  End If
                %>
              </th>
              <th style="text-align: left; width:69px; min-width:69px;" scope="col">
                <a href="/logix/CAM/LMG-rejections.aspx?<% Sendb(SearchString) %>&amp;SortText=PrimaryExtID&amp;SortDirection=<% Sendb(SortDirection) %>">
                  <% Sendb(IIF(Logix.UserRoles.ViewEditUSAMRejections, Copient.PhraseLib.Lookup("term.customer", LanguageID) & " " & Copient.PhraseLib.Lookup("term.information", LanguageID), Copient.PhraseLib.Lookup("term.customer", LanguageID)))%>
                </a>
                <%
                  If SortText = "PrimaryExtID" Then
                    If SortDirection = "ASC" Then
                      Sendb("<span class=""sortarrow"">&#9660;</span>")
                    Else
                      Sendb("<span class=""sortarrow"">&#9650;</span>")
                    End If
                  Else
                  End If
                %>
              </th>
              <th style="text-align:left; width:60px; min-width:60px;" align="left" scope="col">
                <a href="/logix/CAM/LMG-rejections.aspx?<% Sendb(SearchString) %>&amp;SortText=OfferID&amp;SortDirection=<% Sendb(SortDirection) %>">
                  <% Sendb(Copient.PhraseLib.Lookup("term.offer", LanguageID))%>
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
              <th style="text-align:left; width:38px; min-width:38px;" scope="col">
                <a href="/logix/CAM/LMG-rejections.aspx?<% Sendb(SearchString) %>&amp;SortText=RewardQty&amp;SortDirection=<% Sendb(SortDirection) %>">
                  <% Sendb("Qty")%>
                </a>
                <%
                  If SortText = "RewardQty" Then
                    If SortDirection = "ASC" Then
                      Sendb("<span class=""sortarrow"">&#9660;</span>")
                    Else
                      Sendb("<span class=""sortarrow"">&#9650;</span>")
                    End If
                  Else
                  End If
                %>
              </th>
              <th style="text-align:left; width:128px; min-width:128px;" scope="col">
                <a href="/logix/CAM/LMG-rejections.aspx?<% Sendb(SearchString) %>&amp;SortText=IssuanceDate&amp;SortDirection=<% Sendb(SortDirection) %>">
                  <% Sendb(Copient.PhraseLib.Lookup("term.date", LanguageID))%>
                </a>
                <%
                  If SortText = "IssuanceDate" Then
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
              If sizeOfData > 0 Then
                ' First: the all records line
                If sizeOfData > 1 Then
                  If Logix.UserRoles.EditLMGRejections Then
                    Send("<tr class=""" & Shaded & """ style=""padding-bottom:4px;"">")
                    Send("  <!-- Special row for applying changes to all shown records -->")
                    Send("  <td align=""center"">")
                    Send("    <input type=""button"" id=""editall"" name=""editall"" class=""ed"" title=""" & Copient.PhraseLib.Lookup("term.edit", LanguageID) & " " & Copient.PhraseLib.Lookup("term.all", LanguageID) & """ value=""E"" onclick=""javascript:editItem('-1');""" & IIf(Logix.UserRoles.EditLMGRejections, "", " disabled=""disabled""") & " />")
                    Send("    <input type=""button"" id=""retryall"" name=""retryall"" class=""adjust"" title=""" & Copient.PhraseLib.Lookup("term.retry", LanguageID) & " " & Copient.PhraseLib.Lookup("term.all", LanguageID) & """ value=""✓"" onclick=""javascript:retryItem('-1', " & sizeOfData & ");""" & IIf(Logix.UserRoles.EditLMGRejections, "", " disabled=""disabled""") & " />")
                    Send("    <input type=""button"" id=""deleteall"" name=""deleteall"" class=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & " " & Copient.PhraseLib.Lookup("term.all", LanguageID) & """ value=""X"" onclick=""javascript:deleteItem('-1', " & sizeOfData & ");""" & IIf(Logix.UserRoles.EditLMGRejections, "", " disabled=""disabled""") & " />")
                    Send("  </td>")
                    Send("  <td colspan=""8"" style=""font-weight:bold;text-align:center;"">" & Copient.PhraseLib.Detokenize("lmg.AllRecordsInView", LanguageID, sizeOfData) & "</td>")
                    Send("</tr>")
                    Send("<tr class=""editline"" id=""editline-1"" style=""display:none;"">")
                    Send("  <td style=""text-align:center;""><input type=""button"" id=""save-1"" name=""save-1"" title=""" & Copient.PhraseLib.Lookup("term.save", LanguageID) & " &amp; " & Copient.PhraseLib.Lookup("term.retry", LanguageID) & " " & Copient.PhraseLib.Lookup("term.all", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.save", LanguageID) & " &amp; " & Copient.PhraseLib.Lookup("term.retry", LanguageID) & """ onclick=""javascript:handleSave('-1', " & sizeOfData & ");""" & IIf(Logix.UserRoles.EditLMGRejections, "", " disabled=""disabled""") & " style=""width:80px;"" /></td>")
                    Send("  <td></td>")
                    Send("  <td><input type=""text"" id=""ClientLocationCode-1"" name=""ClientLocationCode-1"" maxlength=""30"" value="""" style=""width:60px;"" /></td>")
                    Send("  <td><input type=""text"" id=""BoxID-1"" name=""BoxID-1"" maxlength=""5"" value="""" style=""width:35px;"" /></td>")
                    Send("  <td><input type=""text"" id=""TransactionNumber-1"" name=""TransactionNumber-1"" maxlength=""128"" value="""" style=""width:45px;"" /></td>")
                    Send("  <td></td>")
                    Send("  <td><input type=""text"" id=""OfferID-1"" name=""OfferID-1"" maxlength=""20"" value="""" style=""width:45px;"" /></td>")
                    Send("  <td><input type=""text"" id=""RewardQty-1"" name=""RewardQty-1"" maxlength=""9"" value="""" style=""width:22px;"" /></td>")
                    Send("  <td><input type=""text"" id=""IssuanceDate-1"" name=""IssuanceDate-1"" maxlength=""30"" value="""" style=""width:118px;"" /></td>")
                    Send("</tr>")
                  End If
                End If
                ' Next: all the per-record lines
                MyCommon.QueryStr = "select LocationID, ExtLocationCode, LocationName from Locations with (NoLock) " & _
                                    "where TestingLocation=0 order by ExtLocationCode;"
                rstStores = MyCommon.LRT_Select
                MyCommon.QueryStr = "select IncentiveID as OfferID, IncentiveName as OfferName, StartDate, EndDate from CPE_Incentives with (NoLock) " & _
                                    "where Deleted=0 order by IncentiveID;"
                rstOffers = MyCommon.LRT_Select
                While (i < sizeOfData And i < linesPerPage + linesPerPage * PageNum)
                  If Shaded = "shaded" Then
                    Shaded = ""
                  Else
                    Shaded = "shaded"
                  End If
                  'Details line
                  Send("<tr class=""" & Shaded & """>")
                  Send("  <!-- PKID " & dt.Rows(i).Item("PKID") & " -->")
                  Send("  <td align=""center"" title=""" & dt.Rows(i).Item("PKID") & """>")
                  Send("    <input type=""button"" id=""edit" & dt.Rows(i).Item("PKID") & """ name=""edit" & dt.Rows(i).Item("PKID") & """ class=""ed"" title=""" & Copient.PhraseLib.Lookup("term.edit", LanguageID) & """ value=""E"" onclick=""javascript:editItem('" & dt.Rows(i).Item("PKID") & "');""" & IIf(Logix.UserRoles.EditLMGRejections, "", " disabled=""disabled""") & " />")
                  Send("    <input type=""button"" id=""retry" & dt.Rows(i).Item("PKID") & """ name=""retry" & dt.Rows(i).Item("PKID") & """ class=""adjust"" title=""" & Copient.PhraseLib.Lookup("term.retry", LanguageID) & """ value=""✓"" onclick=""javascript:retryItem('" & dt.Rows(i).Item("PKID") & "');""" & IIf(Logix.UserRoles.EditLMGRejections, "", " disabled=""disabled""") & " />")
                  Send("    <input type=""button"" id=""delete" & dt.Rows(i).Item("PKID") & """ name=""delete" & dt.Rows(i).Item("PKID") & """ class=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ value=""X"" onclick=""javascript:deleteItem('" & dt.Rows(i).Item("PKID") & "');""" & IIf(Logix.UserRoles.EditLMGRejections, "", " disabled=""disabled""") & " />")
                  Send("  </td>")
                  Sendb("  <td>")
                  If MyCommon.NZ(dt.Rows(i).Item("DeliverableType"), -1) > -1 Then
                    MyCommon.QueryStr = "select Description from SourceTypes with (NoLock) " & _
                                        "where SourceTypeID=" & MyCommon.NZ(dt.Rows(i).Item("SourceTypeID"), 0) & ";"
                    dt2 = MyCommon.LEX_Select
                    If dt2.Rows.Count > 0 Then
                      Sendb(MyCommon.NZ(dt2.Rows(0).Item("Description"), ""))
                    Else
                      Sendb(dt.Rows(i).Item("DeliverableType"))
                    End If
                  Else
                    Sendb("&nbsp;")
                  End If
                  Send("</td>")
                  If MyCommon.NZ(dt.Rows(i).Item("ClientLocationCode"), "") <> "" Then
                    MyCommon.QueryStr = "select LocationID, LocationName from Locations with (NoLock) " & _
                                        "where ExtLocationCode='" & MyCommon.NZ(dt.Rows(i).Item("ClientLocationCode"), "") & "';"
                    dt2 = MyCommon.LRT_Select
                    If dt2.Rows.Count > 0 Then
                      Send("  <td><a href=""/logix/store-edit.aspx?LocationID=" & MyCommon.NZ(dt2.Rows(0).Item("LocationID"), "") & """>" & MyCommon.NZ(dt.Rows(i).Item("ClientLocationCode"), "") & "</a></td>")
                    Else
                      Send("  <td>" & MyCommon.NZ(dt.Rows(i).Item("ClientLocationCode"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)) & "</td>")
                    End If
                  Else
                    Send("  <td>&nbsp;</td>")
                  End If
                  Send("  <td>" & MyCommon.NZ(dt.Rows(i).Item("BoxID"), 0) & "</td>")
                  Send("  <td align=""left""><div style=""word-break: break-all;"">"& MyCommon.NZ(dt.Rows(i).Item("TransactionNumber"), 0) & "</div></td>")
                   If MyCommon.NZ(dt.Rows(i).Item("PrimaryExtID"), "") <> "" Then
                    If MyCommon.NZ(dt.Rows(i).Item("AirmileMemberID"), "") <> "" Then
                    MyCommon.QueryStr = "select C.CustomerPK, C.CardPK, C.ExtCardID, CE.AirmileMemberID from CardIDs as C with (NoLock) " & _
										"Left Join CustomerExt as CE with (NoLock) on C.CustomerPK=CE.CustomerPK " & _
                                        "where C.ExtCardID='" & MyCommon.NZ(dt.Rows(i).Item("PrimaryExtID"), "") & "' " & _
                                        "and C.CardTypeID=1;"
										IsUSAM = True
                    Else
                     MyCommon.QueryStr = "select CustomerPK, CardPK, ExtCardID, '' as AirmileMemberID from CardIDs with (NoLock) " & _
                                        "where ExtCardID='" & MyCommon.NZ(dt.Rows(i).Item("PrimaryExtID"), "") & "' " & _
                                        "and CardTypeID in (select CardTypeID from CardTypes where CustTypeID=2);"
                    End If
                    dt2 = MyCommon.LXS_Select
                    If dt2.Rows.Count > 0 Then
                      Sendb("  <td>")
                      For Each row In dt2.Rows
                        j += 1
						If Not (IsUSAM) Then
                        Sendb("<a href=""/logix/CAM/CAM-customer-general.aspx?CustPK=" & MyCommon.NZ(row.Item("CustomerPK"), 0) & _
                          IIf(MyCommon.NZ(row.Item("CardPK"), 0) > 0, "&amp;CardPK=" & row.Item("CardPK"), "") & """>" & MyCryptLib.SQL_StringDecrypt(row.Item("ExtCardID").ToString()) & "</a>")
						Else
                                        Sendb("<a href=""/logix/customer-general.aspx?CustPK=" & MyCommon.NZ(row.Item("CustomerPK"), 0) & _
                                          IIf(MyCommon.NZ(row.Item("CardPK"), 0) > 0, "&amp;CardPK=" & row.Item("CardPK"), "") & """>" & MyCryptLib.SQL_StringDecrypt(row.Item("ExtCardID").ToString()) & "</a>")
						End If 
                        If j < dt2.Rows.Count Then
                          Sendb("<br />")
                        End If
                      Next
		              If (Logix.UserRoles.ViewEditUSAMRejections) Then
                       	Sendb(IIf(MyCommon.NZ(row.Item("AirmileMemberID"), "") <> "", "<br />" & Copient.PhraseLib.Lookup("term.airmilememberid", LanguageID) & "<br />" & row.Item("AirmileMemberID"), ""))
                       	Sendb(IIf(MyCommon.NZ(dt.Rows(i).Item("AirmileMemberID"),"")<>"", "<br />" & Copient.PhraseLib.Lookup("term.lmgrejected", LanguageID).Substring(4,8) & "<br />" & dt.Rows(i).Item("AirmileMemberID"),"") & "</td>")
		              End If
                      Send("</td>")
                    Else
                      Send("  <td>" & MyCommon.NZ(dt.Rows(i).Item("PrimaryExtID"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)) & "</td>")
                    End If
                  Else
                    Send("  <td>&nbsp;</td>")
                  End If
                  'If MyCommon.NZ(dt.Rows(i).Item("AirmileMemberID"), "") = "" Then
                  '  Send("  <td>&nbsp;</td>")
                  'Else
                  '  Send("  <td>" & MyCommon.NZ(dt.Rows(i).Item("AirmileMemberID"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)) & "</td>")
                  'End If
                  If MyCommon.NZ(dt.Rows(i).Item("OfferID"), 0) = 0 Then
                    Send("  <td>&nbsp;</td>")
                  Else
                    Send("  <td><a href=""/logix/offer-redirect.aspx?OfferID=" & MyCommon.NZ(dt.Rows(i).Item("OfferID"), 0) & """>" & MyCommon.NZ(dt.Rows(i).Item("OfferID"), 0) & "</a></td>")
                  End If
                  Send("  <td>" & MyCommon.NZ(dt.Rows(i).Item("RewardQty"), 0) & "</td>")
                  If (Not IsDBNull(dt.Rows(i).Item("IssuanceDate"))) Then
                    Send("  <td>" & Format(dt.Rows(i).Item("IssuanceDate"), "dd MMM yyyy, HH:mm:ss") & "</td>")
                  Else
                    Send("  <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
                  End If
                  Send("</tr>")
                  ' Edit line
                  Send("<tr class=""editline"" id=""editline" & dt.Rows(i).Item("PKID") & """ style=""display:none;"">")
                  Send("  <td style=""text-align:center;""><input type=""button"" id=""save" & dt.Rows(i).Item("PKID") & """ name=""save" & dt.Rows(i).Item("PKID") & """ title=""" & Copient.PhraseLib.Lookup("term.save", LanguageID) & " &amp; " & Copient.PhraseLib.Lookup("term.retry", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.save", LanguageID) & " &amp; " & Copient.PhraseLib.Lookup("term.retry", LanguageID) & """ onclick=""javascript:handleSave('" & dt.Rows(i).Item("PKID") & "', 1);""" & IIf(Logix.UserRoles.EditLMGRejections, "", " disabled=""disabled""") & " style=""width:80px;"" /></td>")
                  Send("  <td></td>")
                  Send("  <td><input type=""text"" id=""ClientLocationCode" & dt.Rows(i).Item("PKID") & """ name=""ClientLocationCode" & dt.Rows(i).Item("PKID") & """ maxlength=""30"" value=""" & MyCommon.NZ(dt.Rows(i).Item("ClientLocationCode"), "") & """ style=""width:60px;"" /></td>")
                  Send("  <td><input type=""text"" id=""BoxID" & dt.Rows(i).Item("PKID") & """ name=""BoxID" & dt.Rows(i).Item("PKID") & """ maxlength=""5"" value=""" & MyCommon.NZ(dt.Rows(i).Item("BoxID"), 0) & """ style=""width:35px;"" /></td>")
                  Send("  <td><input type=""text"" id=""TransactionNumber" & dt.Rows(i).Item("PKID") & """ name=""TransactionNumber" & dt.Rows(i).Item("PKID") & """ maxlength=""128"" value=""" & MyCommon.NZ(dt.Rows(i).Item("TransactionNumber"), 0) & """ style=""width:45px;"" /></td>")
                  Send("  <td></td>")
                  Sendb("  <td><input type=""text"" id=""OfferID" & dt.Rows(i).Item("PKID") & """ name=""OfferID" & dt.Rows(i).Item("PKID") & """ maxlength=""20"" value=""" & MyCommon.NZ(dt.Rows(i).Item("OfferID"), 0) & """ style=""width:45px;"" /></td>")
                  Send("  <td><input type=""text"" id=""RewardQty" & dt.Rows(i).Item("PKID") & """ name=""RewardQty" & dt.Rows(i).Item("PKID") & """ maxlength=""9"" value=""" & MyCommon.NZ(dt.Rows(i).Item("RewardQty"), 0) & """ style=""width:22px;"" /></td>")
                  Send("  <td><input type=""text"" id=""IssuanceDate" & dt.Rows(i).Item("PKID") & """ name=""IssuanceDate" & dt.Rows(i).Item("PKID") & """ maxlength=""30"" value=""" & Format(dt.Rows(i).Item("IssuanceDate"), "M/d/yyyy H:mm:ss") & """ style=""width:118px;"" /></td>")
                  Send("</tr>")
                  i = i + 1
                End While
              Else
                Send("<tr class=""" & Shaded & """>")
                Send("  <td colspan=""9"" style=""text-align:center;"">" & Copient.PhraseLib.Lookup("lmgrejections.none", LanguageID) & "</td>")
                Send("</tr>")
              End If
            %>
          </tbody>
        </table>
      </div>
    </div>
  </form>
</div>

<script runat="server">
  Sub LogAction(ByVal Action As String, ByVal AccessDate As DateTime, ByVal UserName As String, ByVal FullName As String, ByVal ClientLocationCode As String, ByVal BoxID As Long, ByVal TransactionNumber As Long, ByVal PrimaryExtID As String, ByVal OfferID As Long, ByVal IssuanceDate As String, ByVal ProgramID As Long, ByVal RewardQty As Integer)
    Dim MyCommon As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    Dim LogFile As String = ""
    Dim IPAddress As String = ""
    
    MyCommon.Open_LogixRT()
    
    LogFile = "LMGRejectionAction." & MyCommon.Leading_Zero_Fill(Year(Today), 4) & MyCommon.Leading_Zero_Fill(Month(Today), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
    MyCommon.Write_Log(LogFile, Action & vbTab & Today & " " & TimeOfDay & vbTab & UserName & vbTab & FullName & vbTab & ClientLocationCode & vbTab & BoxID & vbTab & TransactionNumber & vbTab & PrimaryExtID & vbTab & OfferID & vbTab & IssuanceDate & vbTab & ProgramID & vbTab & RewardQty)
    
    MyCommon.Close_LogixRT()
  End Sub
</script>

<%
done:
  Send_BodyEnd("mainform", "SearchSource")
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixEX()
  MyCommon.Close_LogixXS()
  MyCommon.Close_LogixWH()
  Logix = Nothing
  MyCommon = Nothing
%>
