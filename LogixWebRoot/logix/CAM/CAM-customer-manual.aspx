<%@ Page Language="vb" Debug="true" CodeFile="../LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: CAM-customer-manual.aspx 
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
  Dim MyCam As New Copient.CAM
  Dim MyLookup As New Copient.CustomerLookup
  Dim TxDetail As New Copient.CAM.TransactionDetail
  Dim ProgDetail As New Copient.CAM.ProgramDetail
  Dim MyCryptLib As New Copient.CryptLib
  Dim MyPoints As New Copient.Points
  Dim Logix As New Copient.LogixInc
  Dim dt As DataTable = Nothing
  Dim dt2 As DataTable = Nothing
  Dim dtTrans As DataTable
  Dim row As DataRow
  Dim CustomerPK As Long
  Dim CardPK As Long
  Dim ExtCardID As String = ""
  Dim FirstName As String = ""
  Dim MiddleName As String = ""
  Dim LastName As String = ""
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
  Dim OfferID As Long = 0
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
  Dim StatusMessage As String = ""
  Dim Handheld As Boolean = False
  Dim TransDateStart As String = ""
  Dim TransDateEnd As String = ""
  Dim DisabledPtsAdj As String = ""
  Dim OfferTable As Hashtable = Nothing
  Dim OfferList As String = ""
  Dim CgXML As String = ""
  Dim AllCAMCardholdersID As Long = 0
  Dim Employee As Boolean = False
  Dim AdjAmount As Integer = 0
  Dim ProgramIDs(-1) As Integer
  Dim HighlightedPKID As Integer = 0
  Dim IsDate As Boolean = False
  
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
  Dim TransTerms As String = ""
  Dim SearchFilter As String = ""
  Dim tempDate As Date
  Dim SortCol As String = "PrimaryExtID"
  Dim SortDir As String = "asc"
  Dim SortUrl As String = ""
  Dim LogixTransNums(-1) As String
  Dim LogixTransNum As String = ""
  Dim CreateNewTrans As Boolean = False
  Dim ResultsMessage As String = ""
  Dim ClearEntry As Boolean = False
  Dim TransNum As String = ""
  Dim PKID As Integer = 0
  Dim WarningProgramID As Long = 0
  Dim ExecutePermitted As Boolean = False
  Dim LockCriteria As Boolean = False
  Dim LockTime As Boolean = False
  Dim ShowSave As Boolean = False
  Dim ShowResults As Boolean = False
  Dim ShowCreate As Boolean = False
  Dim TransNote As String = ""
  Dim LogixTransList(-1) As String
  Dim CustDetail As New Copient.Customer
  Dim TempAdjAmount As Integer = 0
  Dim SessionID As String = ""
  Dim StoreShort As String = ""
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "CAM-customer-manual.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  MyCommon.Open_LogixWH()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  UnknownPhrase = Copient.PhraseLib.Lookup("term.unknown", LanguageID)
  ExecutePermitted = Logix.UserRoles.ExecuteCAMAdjustment
  
  If Session("SessionID") IsNot Nothing AndAlso Session("SessionID").ToString.Trim <> "" Then
    SessionID = Session("SessionID").ToString.Trim
    CopientNotes &= "SessionID: " & SessionID
  End If
  
  ' lets check the logged in user and see if they are to be restricted to this page
  MyCommon.QueryStr = "select AU.StartPageID,AUSP.PageName ,AUSP.prestrict from AdminUsers as AU with (NoLock) " & _
                      "inner join AdminUserStartPages as AUSP with (NoLock) on AU.StartPageID=AUSP.StartPageID " & _
                      "where AU.AdminUserID=" & AdminUserID
  dt = MyCommon.LRT_Select
  If dt.Rows.Count > 0 Then
    If (MyCommon.NZ(dt.Rows(0).Item("prestrict"), False) = True) Then
      ' ok we got in here then we need to restrict the user from seeing any other pages
      restrictLinks = True
    End If
  End If
  
  Send_HeadBegin("term.customer", "term.manual", CustomerPK)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  If restrictLinks Then
    Send("<link rel=""stylesheet"" href=""../../css/logix-restricted.css"" type=""text/css"" media=""screen"" />")
  End If
  Send_Scripts(New String() {"datePicker.js"})
%>
<script type="text/javascript">
  var datePickerDivID = "datepicker";
  
  if (window.captureEvents){
    window.captureEvents(Event.CLICK);
    window.onclick=handlePageClick;
  } else {
    document.onclick=handlePageClick;
  }
  
  <% Send_Calendar_Overrides(MyCommon) %>

  function handlePageClick(e) {
    var calFrame = document.getElementById('calendariframe');
    var el=(typeof event!=='undefined')? event.srcElement : e.target        
    
    if (el != null) {
      var pickerDiv = document.getElementById(datePickerDivID);
      if (pickerDiv != null && pickerDiv.style.visibility == "visible") {
        if (el.id!="transdate-picker") {
          if (!isDatePickerControl(el.className)) {
            pickerDiv.style.visibility = "hidden";
            pickerDiv.style.display = "none"; 
            if (calFrame != null) {
              calFrame.style.visibility = 'hidden';
              calFrame.style.display = "none";
            }
          }
        } else  {
          pickerDiv.style.visibility = "visible";            
          pickerDiv.style.display = "block";     
          if (calFrame != null) {
            calFrame.style.visibility = 'visible';
            calFrame.style.display = "block";
          }
        }
      }
    }
  }
  
  function isDatePickerControl(ctrlClass) {
    var retVal = false;
    
    if (ctrlClass != null && ctrlClass.length >= 2) {
      if (ctrlClass.substring(0,2) == "dp") {
        retVal = true;
      }
    }
    return retVal;
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
    var qryStr = 'CAM-customer-manual.aspx?search=Search&searchterms=<%Sendb(strTerms)%>&CustPK=<%Sendb(CustomerPK)%>&offerSearch=Search&transterms=' + transTerms;
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

function deleteTransaction(pkid, transNum) {
  var elemAT = document.getElementById('actiontype');
  var elemP1 = document.getElementById('parm1');
  var elemP2 = document.getElementById('parm2');
  var tokenValues = [];
  var msg = '';
  
  tokenValues[0] = transNum;
  msg = '<%Sendb(Copient.PhraseLib.Lookup("customer-manual.ConfirmDeleteTransaction", LanguageID))%>';
  msg = detokenizeString(msg, tokenValues);

  if (elemAT != null && elemP1 != null && elemP2 != null) {
    if (confirm(msg)) {
      elemAT.value = 'delete';
      elemP1.value = pkid;
      elemP2.value = '<%Sendb(Request.QueryString("pagenum")) %>';
      document.mainform.submit();
    }
  }
}

function applyTransaction(pkid, transNum, rowID) {
  var elemAT = document.getElementById('actiontype');
  var elemP1 = document.getElementById('parm1');
  var elemP2 = document.getElementById('parm2');
  var elemP3 = document.getElementById('parm3');
  var elemP4 = document.getElementById('parm4');
  var elemOID = document.getElementById('OfferID' + rowID);
  var elemAdj = document.getElementById('Adjustment' + rowID);
  
  if (elemAT != null && elemP1 != null && elemP2 != null && elemP3 != null && elemP4 != null && elemOID != null && elemAdj != null) {
    if (isNaN(elemOID.value)) {
      alert('<%Sendb(Copient.PhraseLib.Lookup("customer-manual.InvalidOfferID", LanguageID)) %>')
    } else if (isNaN(elemAdj.value)) {
      alert('<%Sendb(Copient.PhraseLib.Lookup("customer-manual.InvalidOfferID", LanguageID)) %>')
    } else {
      elemAT.value = 'apply';
      elemP1.value = pkid;
      elemP2.value = '<%Sendb(Request.QueryString("pagenum")) %>';
      elemP3.value = elemOID.value
      elemP4.value = elemAdj.value
      document.mainform.submit();
    }
  } 
}

function saveTransaction(pkid, transNum, rowID) {
  var elemAT = document.getElementById('actiontype');
  var elemP1 = document.getElementById('parm1');
  var elemP2 = document.getElementById('parm2');
  var elemP3 = document.getElementById('parm3');
  var elemP4 = document.getElementById('parm4');
  var elemP5 = document.getElementById('parm5');
  var elemOID = document.getElementById('OfferID' + rowID);
  var elemAdj = document.getElementById('Adjustment' + rowID);
  var elemTdate = document.getElementById('tdate' + rowID);
  
  if (elemAT != null && elemP1 != null && elemP2 != null && elemP3 != null && elemP4 != null && elemOID != null && elemAdj != null) {
    if (isNaN(elemOID.value)) {
      alert('<%Sendb(Copient.PhraseLib.Lookup("customer-manual.InvalidOfferID", LanguageID)) %>')
    } else if (isNaN(elemAdj.value)) {
      alert('<%Sendb(Copient.PhraseLib.Lookup("customer-manual.InvalidOfferID", LanguageID)) %>')
    } else {
      elemAT.value = 'save';
      elemP1.value = pkid;
      elemP2.value = '<%Sendb(Request.QueryString("pagenum")) %>';
      elemP3.value = elemOID.value
      elemP4.value = elemAdj.value
      if (elemP5 != null && elemTdate!= null) elemP5.value = elemTdate.value
      document.mainform.submit();
    }
  } 
} 

function selectTrans(logixTransNum) {
  var elemAT = document.getElementById('actiontype');
  var elemP1 = document.getElementById('parm1');
  
  if (elemAT != null && elemP1 != null) {
    elemAT.value="selectTrans";
    elemP1.value = logixTransNum;
    document.mainform.submit();
  }  
}    

function launchOffers(logixTransNum, offerDate, offerStore, row) {
  var elemCustPK = document.getElementById('custPK' + row);
  var queryStr = '/logix/CAM/CAM-manual-offers.aspx?row=' + row
  
  if (elemCustPK != null) {
    queryStr += '&CustPK=' + elemCustPK.value + '&logixTransNum=' + logixTransNum
    // uncomment the line 3 down to send the transaction's date and store for list of offers
    // pertaining on to the store where the date is within the start and end date.
    // not seeming to find the offers though on CAM-manual-offers when this line is uncommented.
    queryStr += '&offerdate=' + offerDate + '&offerstore=' + offerStore
    openWidePopup(queryStr);
  } else {
    alert('<%Sendb(Copient.PhraseLib.Lookup("customer-manual.LoadingError", LanguageID)) %>');
  }
}

function showNote(noteRow) {
  var elem = document.getElementById('note' + noteRow);
  
  if (elem != null) {
    elem.style.display = (elem.style.display=='none') ? 'block' : 'none';
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
    Send_Subtabs(Logix, 32, 2, LanguageID, 0)
  Else
    Send_Subtabs(Logix, 91, 1, LanguageID, 0)
  End If
  If (Logix.UserRoles.AccessCustomerInquiry = False) Then
    Send_Denied(1, "perm.customer-access")
    GoTo done
  End If
  
  ResultsMessage = Request.QueryString("RetMsg")
  If ResultsMessage Is Nothing Then
    ResultsMessage = ""
  End If
  
  
  ExtCardID = Request.QueryString("card")
  If ExtCardID Is Nothing Then
    ExtCardID = ""
  End If
  If ExtCardID <> "" Then
    ExtCardID = MyCommon.Pad_ExtCardID(ExtCardID,2)
  End If
  
  ProgDetail.AdjustmentAmount = MyCommon.Extract_Val(Request.QueryString("adjustment"))
  TxDetail.TransNumber = Request.QueryString("transaction")
  If TxDetail.TransNumber Is Nothing Then
    TxDetail.TransNumber = ""
  End If
  TxDetail.TransOffer = MyCommon.Extract_Val(Request.QueryString("offerid"))
  TxDetail.TransDateStr = Request.QueryString("transdate")
  If TxDetail.TransDateStr Is Nothing Then
    TxDetail.TransDateStr = ""
  End If
  TxDetail.TransStore = Request.QueryString("store")
  If TxDetail.TransStore Is Nothing Then
    TxDetail.TransStore = ""
  End If
  TxDetail.TransTerminal = Request.QueryString("lane")
  If TxDetail.TransTerminal Is Nothing Then
    TxDetail.TransTerminal = ""
  End If
  TxDetail.Note = Request.QueryString("transNote")
  If TxDetail.Note Is Nothing Then
    TxDetail.Note = ""
  End If
  TxDetail.TransTimeStr = Request.QueryString("transTime")
  If TxDetail.TransTimeStr Is Nothing Then
    TxDetail.TransTimeStr = ""
  End If
  
  If TxDetail.TransStore = "" Then
    'Full code not provided, so try to work it out from the short code.
    If MyCommon.NZ(Request.QueryString("storeshort"), "").Trim <> "" Then
      StoreShort = Request.QueryString("storeshort").Trim
      StoreShort = StoreShort.PadLeft(4, "0")
      'See if there's a store the rightmost digits of whose ExtLocationCode matches the padded short code.
      MyCommon.QueryStr = "select top 1 ExtLocationCode from Locations with (NoLock) where RIGHT(ExtLocationCode, 4)='" & StoreShort & "';"
      dt = MyCommon.LRT_Select
      If dt.Rows.Count > 0 Then
        'Found it
        TxDetail.TransStore = MyCommon.NZ(dt.Rows(0).Item("ExtLocationCode"), "")
      Else
        'Didn't find it, so try matching the raw, unpadded shortcode.
        MyCommon.QueryStr = "select top 1 ExtLocationCode from Locations with (NoLock) where ExtLocationCode='" & Request.QueryString("storeshort").Trim & "';"
        dt2 = MyCommon.LRT_Select
        If dt2.Rows.Count > 0 Then
          TxDetail.TransStore = MyCommon.NZ(dt2.Rows(0).Item("ExtLocationCode"), "")
        End If
      End If
      If TxDetail.TransStore.Length < 4 Then
        StoreShort = TxDetail.TransStore
      End If
    End If
  Else
    'Full code provided.  Derive short code from it.
    StoreShort = Right(TxDetail.TransStore, 4)
    If TxDetail.TransStore.Length > 4 Then
      StoreShort = StoreShort.PadLeft(4, "0")
    End If
  End If
  
  'Load customer information
  If ExtCardID.Trim <> "" Then
    MyCommon.QueryStr = "select top 1 CustomerPK from CardIDs with (NoLock) where ExtCardID='" & MyCommon.Parse_Quotes(MyCryptLib.SQL_StringEncrypt(ExtCardID.Trim)) & "' and CardTypeID in " & _
                        "  (select CardTypeID from CardTypes where CustTypeID=2);"
    dt = MyCommon.LXS_Select
    If dt.Rows.Count > 0 Then
      CustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)
    End If
    ' create a customer if one doesn't already exist
    If CustomerPK = 0 AndAlso MyCam.VerifyCardNumber(ExtCardID, infoMessage) Then
      CustomerPK = MyCam.AddCustomer(ExtCardID, AdminUserID, infoMessage)
    End If
  End If
  
  If (infoMessage = "" AndAlso Request.QueryString("transSearch") <> "") Then
    Try
      IsDate = Date.TryParse(TxDetail.TransDateStr.Trim, MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, tempDate)
      If CustomerPK > 0 Then
        If TxDetail.TransNumber.Trim = "" OrElse TxDetail.TransDateStr.Trim = "" OrElse TxDetail.TransStore.Trim = "" OrElse TxDetail.TransTerminal.Trim = "" Then
          infoMessage = Copient.PhraseLib.Detokenize("customer-manual.EnterAllInformation", LanguageID, StoreShort, TxDetail.TransStore)  'Please enter all required transaction information (labeled in red).  Storeshort={0}, TransStore={1}.
        ElseIf Not IsDate Or tempDate.Year < 2000 Then
          infoMessage = Copient.PhraseLib.Lookup("customer-manual.InvalidDate", LanguageID)
        ElseIf IsDate AndAlso tempDate > Date.Now Then
          infoMessage = Copient.PhraseLib.Lookup("customer-manual.InvalidFutureDate", LanguageID)
        Else
          If MyCam.IsValidLocation(TxDetail.TransStore) Then
            LogixTransNums = MyCam.FindCustomerTransaction(ExtCardID, TxDetail)
            ShowResults = True
            LockCriteria = True
          Else
            infoMessage = Copient.PhraseLib.Detokenize("customer-manual.StoreNotFound", LanguageID, TxDetail.TransStore)  'Store not found in Logix: {0}
          End If
        End If
      Else
        infoMessage = Copient.PhraseLib.Detokenize("customer-manual.NoCustomerFound", LanguageID, ExtCardID)  'No customer found for card number: {0}
      End If
    Catch ex As Exception
      infoMessage = Copient.PhraseLib.Detokenize("customer-manual.ErrorEncountered", LanguageID, ex.ToString)  'Error encountered while searching for transaction:<br />{0}
    End Try
  ElseIf (Request.QueryString("transClear") <> "") Then
    ExtCardID = ""
    ProgDetail.AdjustmentAmount = 0
    TxDetail.TransNumber = ""
    TxDetail.TransOffer = 0
    TxDetail.TransDateStr = ""
    TxDetail.TransStore = ""
    TxDetail.TransTerminal = ""
    TxDetail.Note = ""
  ElseIf (Request.QueryString("createTrans") <> "") Then
    ShowSave = True
    ShowResults = False
    LockCriteria = True
  ElseIf (Request.QueryString("save") <> "") Then
    If CustomerPK = 0 Then
      infoMessage = Copient.PhraseLib.Detokenize("customer-manual.NoCustomerFound", LanguageID, ExtCardID)  'No customer found for card number: {0}
    ElseIf (ExtCardID <> "" AndAlso TxDetail.TransNumber <> "" AndAlso TxDetail.TransStore <> "" AndAlso TxDetail.TransTerminal <> "" AndAlso TxDetail.TransDateStr <> "") Then
      Try
        ProgDetail.ProgramID = 0
        ' add the time component if present and valid
        If Request.QueryString("transTime") <> "" Then
          TxDetail.TransDateStr = AppendTimeToDate(TxDetail.TransDateStr, Request.QueryString("transTime"))
        End If
        ' validate that the offer, if sent, is valid
        If TxDetail.TransOffer = 0 OrElse (Date.TryParse(TxDetail.TransDateStr.Trim, tempDate) AndAlso ValidateOffer(TxDetail.TransOffer, tempDate, infoMessage)) Then
          TxDetail.ViewInManualEntry = True
          LogixTransNum = MyCam.CreateCustomerTransaction(ExtCardID, CustomerPK, AdminUserID, TxDetail, ProgDetail)
          If LogixTransNum.Trim <> "" Then
            ResultsMessage = Copient.PhraseLib.Detokenize("customer-manual.TransactionCreated", LanguageID, TxDetail.TransNumber, ExtCardID)  'Transaction {0} was created and added to the list below for card {1}.
          End If
          ClearEntry = True
          Response.Status = "301 Moved Permanently"
          Response.AddHeader("Location", "CAM-customer-manual.aspx?RetMsg=" & ResultsMessage)
          GoTo done
        End If
      Catch sdneEx As Copient.StoreDoesNotExistException
        infoMessage = Copient.PhraseLib.Detokenize("customer-manual.StoreNotFound", LanguageID, sdneEx.GetStore())  'Store not found in Logix: {0}
      Catch teEx As Copient.TransactionExistsException
        LogixTransNums = teEx.GetLogixTransNums
        If LogixTransNums.Length >= 1 Then
          ' if only one match exists then simply use that transaction
          LogixTransNum = LogixTransNums(0)
          TxDetail.LogixTransNum = LogixTransNum
          infoMessage = SubmitPoint(AdminUserID, CustomerPK, TxDetail, ProgDetail)
          If infoMessage = "" Then
            If LogixTransNum.Trim <> "" Then
              ResultsMessage = Copient.PhraseLib.Detokenize("customer-manual.TransactionAdded", LanguageID, TxDetail.TransNumber, ExtCardID)  'Transaction {0} was added to the list below for card {1}.
            End If
            Response.Status = "301 Moved Permanently"
            Response.AddHeader("Location", "CAM-customer-manual.aspx?RetMsg=" & ResultsMessage)
          End If
        End If
      Catch gtEx As Copient.GeneralTransactionException
        infoMessage = gtEx.GetErrorMessage
      Catch ex As Exception
        infoMessage = Copient.PhraseLib.Detokenize("customer-manual.CreateError", LanguageID, ex.ToString)  'Error encountered during create:<br />{0}
      End Try
    Else
      infoMessage = Copient.PhraseLib.Lookup("customer-manual.EnterSearchCriteria", LanguageID)
    End If
    ShowSave = True
    LockCriteria = True
  ElseIf (Request.QueryString("execute") <> "") Then
    If CustomerPK = 0 Then
      infoMessage = Copient.PhraseLib.Detokenize("customer-manual.NoCustomerFound", LanguageID, ExtCardID)  'No customer found for card number: {0}
      ShowSave = True
      LockCriteria = False
    ElseIf (ExtCardID <> "" AndAlso TxDetail.TransNumber <> "" AndAlso TxDetail.TransStore <> "" _
            AndAlso TxDetail.TransTerminal <> "" AndAlso TxDetail.TransDateStr <> "" _
            AndAlso TxDetail.TransOffer <> 0 AndAlso ProgDetail.AdjustmentAmount <> 0) Then
      ' create the transaction
      Try
        WarningProgramID = MyCommon.Extract_Val(Request.QueryString("WarningProgramID"))
        ProgramIDs = GetProgramsForOffer(TxDetail.TransOffer, infoMessage)
        If infoMessage = "" Then
          If ProgramIDs.Length > 0 Then
            ProgDetail.ProgramID = ProgramIDs(0)
            ProgDetail.SourceTypeID = 1 ' manual entry
            ' add the time component if present and valid
            If Request.QueryString("transTime") <> "" Then
              TxDetail.TransDateStr = AppendTimeToDate(TxDetail.TransDateStr, Request.QueryString("transTime"))
            End If
            If IsValidAdjustment(TxDetail, CustomerPK, ProgDetail, infoMessage, WarningProgramID) Then
              TxDetail.ViewInManualEntry = False
              LogixTransNum = MyCam.CreateCustomerTransaction(ExtCardID, CustomerPK, AdminUserID, TxDetail, ProgDetail)
              TxDetail.LogixTransNum = LogixTransNum
              If LogixTransNum.Trim <> "" Then
                ResultsMessage = Copient.PhraseLib.Detokenize("customer-manual.TransactionCreated", LanguageID, TxDetail.TransNumber, ExtCardID)  'Transaction {0} was created and added to the list below for card {1}.
              End If
              ClearEntry = True
              MyCommon.QueryStr = "select RewardOptionID from CPE_RewardOptions with (NoLock) where IncentiveID =" & TxDetail.TransOffer
              dt = MyCommon.LRT_Select
              If dt.Rows.Count > 0 Then
                TxDetail.TransROID = MyCommon.NZ(dt.Rows(0).Item("RewardOptionID"), 0)
              End If
              infoMessage = AdjustPoint(AdminUserID, ExtCardID, CustomerPK, TxDetail, ProgDetail, SessionID, TxDetail.TransOffer)
              ' if no error was reported then send a message conveying the success of the adjustment
              If infoMessage = "" Then
                Try
                  ' to ensure the that the RedemptionCount for a manual adjustment is only 1, set 
                  ' set the adjustment amount to 1, then reset it after the trans history is recorded
                  TempAdjAmount = ProgDetail.AdjustmentAmount
                  ProgDetail.AdjustmentAmount = 1
                  CustDetail = New Copient.Customer
                  CustDetail.SetCustomerTypeID(2)
                  MyCam.AddToTransHistory(TxDetail, ProgDetail, CustDetail, CardPK, AdminUserID, MyCommon)
                  'reset the program adjustment amount
                  ProgDetail.AdjustmentAmount = TempAdjAmount
                  ResultsMessage = Copient.PhraseLib.Detokenize("customer-manual.TransactionAdjusted", LanguageID, TxDetail.TransNumber, ProgDetail.AdjustmentAmount, ExtCardID)  'Transaction {0} adjusted points balance by {1} for card {2}.
                Catch ex As Exception
                  MyCommon.Write_Log("CAM.txt", "Failed to create transaction " & TxDetail.TransNumber & " for card " & ExtCardID, True)
                End Try
                Response.Status = "301 Moved Permanently"
                Response.AddHeader("Location", "CAM-customer-manual.aspx?RetMsg=" & ResultsMessage)
                GoTo done
              End If
            End If
          Else
            infoMessage = Copient.PhraseLib.Detokenize("customer-manual.NoPointsPrograms", LanguageID, TxDetail.TransOffer) 'There are no points programs associated with offer {0}.
            ShowSave = True
            LockCriteria = True
          End If
        End If
      Catch sdneEx As Copient.StoreDoesNotExistException
        infoMessage = Copient.PhraseLib.Detokenize("customer-manual.StoreNotFound", LanguageID, sdneEx.GetStore())  'Store not found in Logix: {0}
        ShowSave = True
        LockCriteria = False
      Catch teEx As Copient.TransactionExistsException
        LogixTransNums = teEx.GetLogixTransNums
        If LogixTransNums.Length = 1 Then
          ' if only one match exists then simply use that transaction
          LogixTransNum = LogixTransNums(0)
          If LogixTransNum.Trim <> "" Then
            ResultsMessage = Copient.PhraseLib.Detokenize("customer-manual.TransactionExists", LanguageID, TxDetail.TransNumber, ExtCardID)  'Transaction {0} already exists in the list below for card {1}.
          End If
        Else
          infoMessage = Copient.PhraseLib.Lookup("customer-manual.MultipleMatches", LanguageID)
        End If
      Catch gtEx As Copient.GeneralTransactionException
        infoMessage = gtEx.GetErrorMessage
      Catch ex As Exception
        infoMessage = Copient.PhraseLib.Detokenize("customer-manual.CreateError", LanguageID, ex.ToString)  'Error encountered during create:<br />{0}
      End Try
    Else
      infoMessage = Copient.PhraseLib.Lookup("customer-manual.EnterValidValues", LanguageID)
    End If
    LockCriteria = True
    ShowSave = True
  ElseIf (Request.QueryString("actiontype") = "delete") Then
    MyCommon.QueryStr = "delete from PointsAdj_Pending with (RowLock) where PKID=" & MyCommon.Extract_Val(Request.QueryString("Parm1"))
    MyCommon.LXS_Execute()
    PageNum = MyCommon.Extract_Val(Request.QueryString("Parm2"))
  ElseIf (Request.QueryString("actiontype") = "selectTrans") Then
    ' get the LogixTransNum
    LogixTransNum = Request.QueryString("parm1")
    If LogixTransNum IsNot Nothing AndAlso LogixTransNum <> "" Then
      ShowSave = True
      LockCriteria = True
      LockTime = True
    End If
  ElseIf (Request.QueryString("actiontype") = "save") Then
    TxDetail.TransOffer = MyCommon.Extract_Val(Request.QueryString("Parm3"))
    ProgDetail.AdjustmentAmount = MyCommon.Extract_Val(Request.QueryString("Parm4"))
    TxDetail.TransDateStr = Request.QueryString("Parm5")
    ' validate that the offer, if sent, is valid
    If TxDetail.TransOffer = 0 OrElse (Date.TryParse(TxDetail.TransDateStr.Trim, tempDate) AndAlso ValidateOffer(TxDetail.TransOffer, tempDate, infoMessage)) Then
      HighlightedPKID = MyCommon.Extract_Val(Request.QueryString("Parm1"))
      PageNum = MyCommon.Extract_Val(Request.QueryString("Parm2"))
      TransTerms = TxDetail.TransNumber
      ' update with the new offer and adjustment amount
      MyCommon.QueryStr = "Update PointsAdj_Pending with (RowLock) set OfferID=" & TxDetail.TransOffer & ", " & _
                          " AdjAmount=" & ProgDetail.AdjustmentAmount & " " & _
                          "where PKID=" & MyCommon.Extract_Val(Request.QueryString("Parm1")) & ";"
      MyCommon.LXS_Execute()
    End If
    TxDetail.TransDateStr = ""
    TxDetail.TransOffer = 0
    ProgDetail.AdjustmentAmount = 0
  ElseIf (Request.QueryString("actiontype") = "apply") Then
    MyCommon.QueryStr = "select LogixTransNum, TransNum, TransDate, ExtLocationCode, TerminalNum, PEND.CustomerPK, ProgramID " & _
                        "from PointsAdj_Pending AS PEND with (NoLock) " & _
                        "inner join Customers AS CUST with (NoLock) on CUST.CustomerPK=PEND.CustomerPK and CustomerTypeID=2 " & _
                        "where PKID=" & MyCommon.Extract_Val(Request.QueryString("Parm1")) & ";"
    dt = MyCommon.LXS_Select
    If dt.Rows.Count > 0 Then
      TxDetail.LogixTransNum = MyCommon.NZ(dt.Rows(0).Item("LogixTransNum"), "")
      TxDetail.TransNumber = MyCommon.NZ(dt.Rows(0).Item("TransNum"), "")
      tempDate = MyCommon.NZ(dt.Rows(0).Item("TransDate"), New Date(1980, 1, 1))
      TxDetail.TransDateStr = tempDate.ToString
      TxDetail.TransStore = MyCommon.NZ(dt.Rows(0).Item("ExtLocationCode"), "")
      TxDetail.TransTerminal = MyCommon.NZ(dt.Rows(0).Item("TerminalNum"), "")
      CustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)
      MyCommon.QueryStr = "select top 1 ExtCardID from CardIDs with (NoLock) where CustomerPK=" & CustomerPK & " and CardTypeID in " & _
                          "  (select CardTypeID from CardTypes where CustTypeID=2);"
      dt2 = MyCommon.LXS_Select
      If dt2.Rows.Count > 0 Then
        ExtCardID = MyCryptLib.SQL_StringDecrypt(dt2.Rows(0).Item("ExtCardID").ToString())
      End If
      TxDetail.TransOffer = MyCommon.Extract_Val(Request.QueryString("Parm3"))
      ProgDetail.AdjustmentAmount = MyCommon.Extract_Val(Request.QueryString("Parm4"))
      If TxDetail.TransOffer > 0 AndAlso ProgDetail.AdjustmentAmount <> 0 Then
        Try
          HighlightedPKID = MyCommon.Extract_Val(Request.QueryString("Parm1"))
          TransTerms = TxDetail.TransNumber
          ' update with the new offer and adjustment amount
          MyCommon.QueryStr = "Update PointsAdj_Pending with (RowLock) set OfferID=" & TxDetail.TransOffer & ", " & _
                              " AdjAmount=" & ProgDetail.AdjustmentAmount & " " & _
                              "where PKID=" & MyCommon.Extract_Val(Request.QueryString("Parm1")) & ";"
          MyCommon.LXS_Execute()
          WarningProgramID = MyCommon.Extract_Val(Request.QueryString("WarningProgramID"))
          ProgramIDs = GetProgramsForOffer(TxDetail.TransOffer, infoMessage)
          If infoMessage = "" Then
            If ProgramIDs.Length > 0 Then
              ProgDetail.ProgramID = ProgramIDs(0)
              ProgDetail.SourceTypeID = 1 ' manual entry
              If IsValidAdjustment(TxDetail, CustomerPK, ProgDetail, infoMessage, WarningProgramID) Then
                ClearEntry = True
                MyCommon.QueryStr = "select RewardOptionID from CPE_RewardOptions with (NoLock) where IncentiveID =" & TxDetail.TransOffer
                dt = MyCommon.LRT_Select
                If dt.Rows.Count > 0 Then
                  TxDetail.TransROID = MyCommon.NZ(dt.Rows(0).Item("RewardOptionID"), 0)
                End If
                infoMessage = AdjustPoint(AdminUserID, MyCryptLib.SQL_StringDecrypt(ExtCardID), CustomerPK, TxDetail, ProgDetail, SessionID, TxDetail.TransOffer)
                ' if no error was reported then send a message conveying the success of the adjustment
                If infoMessage = "" Then
                  ' insert the transaction into the TransHistory and TransRedemption tables from the PointsAdj_Pending table
                  MyCommon.QueryStr = "select * from TransHistory with (NoLock) where LogixTransNum = '" & TxDetail.LogixTransNum & "'"
                  dt = MyCommon.LWH_Select
                  If dt.Rows.Count = 0 Then
                    MyCommon.QueryStr = "insert into TransHist with (RowLock) (LogixTransNum, CustomerPrimaryExtID, ExtLocationCode, " & _
                                        "  TransDate, TerminalNum, POSTransNum, CustomerTypeID) " & _
                                        "values ('" & TxDetail.LogixTransNum & "', '" & ExtCardID & "', '" & TxDetail.TransStore & "', '" & TxDetail.TransDateStr & "', " & _
                                        "        '" & TxDetail.TransTerminal & "', '" & TxDetail.TransNumber & "', 2);"
                    MyCommon.LWH_Execute()
                  End If
                  MyCommon.QueryStr = "insert into TransRedemption with (RowLock) (OfferID, ExtLocationCode, CustomerPrimaryExtID, RedemptionCount, " & _
                                      "    TransDate, TerminalNum, TransNum, LogixTransNum, CustomerTypeID, RedemptionAmount) " & _
                                      " values (" & TxDetail.TransOffer & ", '" & TxDetail.TransStore & "', '" & ExtCardID & "', " & _
                                      "         " & 1 & ", '" & TxDetail.TransDateStr & "', '" & TxDetail.TransTerminal & "', " & _
                                      "         '" & TxDetail.TransNumber & "', '" & TxDetail.LogixTransNum & "',2,0);"
                  MyCommon.LWH_Execute()
                  ' remove the entry from the PointsAdj_Pending table
                  MyCommon.QueryStr = "delete from PointsAdj_Pending with (RowLock) where LogixTransNum='" & TxDetail.LogixTransNum & "';"
                  MyCommon.LXS_Execute()
                  ResultsMessage = Copient.PhraseLib.Detokenize("customer-manual.TransactionAdjusted", LanguageID, TxDetail.TransNumber, ProgDetail.AdjustmentAmount, ExtCardID) 'Transaction {0} adjusted points balance by {1} for card {2}.
                  Response.Status = "301 Moved Permanently"
                  Response.AddHeader("Location", "CAM-customer-manual.aspx?RetMsg=" & ResultsMessage)
                  GoTo done
                End If
              End If
            Else
              infoMessage = Copient.PhraseLib.Detokenize("customer-manual.NoPointsPrograms", LanguageID, TxDetail.TransOffer)  'There are no points programs associated with offer {0}.
            End If
          End If
        Catch sdneEx As Copient.StoreDoesNotExistException
          infoMessage = Copient.PhraseLib.Detokenize("customer-manual.StoreNotFound", LanguageID, sdneEx.GetStore())  'Store not found in Logix: {0}
        Catch teEx As Copient.TransactionExistsException
          LogixTransNums = teEx.GetLogixTransNums
          If LogixTransNums.Length = 1 Then
            ' if only one match exists then simply use that transaction
            LogixTransNum = LogixTransNums(0)
            If LogixTransNum.Trim <> "" Then
              ResultsMessage = Copient.PhraseLib.Detokenize("customer-manual.TransactionExists", LanguageID, TxDetail.TransNumber, ExtCardID)  'Transaction {0} already exists in the list below for card {1}.
            End If
          Else
            infoMessage = Copient.PhraseLib.Lookup("customer-manual.MultipleMatches", LanguageID)
          End If
        Catch gtEx As Copient.GeneralTransactionException
          infoMessage = gtEx.GetErrorMessage
        Catch ex As Exception
          infoMessage = Copient.PhraseLib.Detokenize("customer-manual.CreateError", LanguageID, ex.ToString)  'Error encountered during create:<br />{0}
        End Try
      Else
        infoMessage = Copient.PhraseLib.Detokenize("customer-manual.InvalidEntry", LanguageID, TxDetail.TransNumber)  'Invalid entry for transaction {0}.
      End If
      PageNum = MyCommon.Extract_Val(Request.QueryString("Parm2"))
    Else
      infoMessage = Copient.PhraseLib.Lookup("customer-manual.TransactionError", LanguageID)
    End If
  End If
  
  ' when necessary, load transaction information
  If (LogixTransNum.Trim <> "" AndAlso Not ClearEntry) Then
    MyCommon.QueryStr = "select ExtLocationCode, TransDate, TerminalNum, POSTransNum, 0 as OfferID, 0 as AdjAmount from TransHistory with (NoLock) where LogixTransNum ='" & LogixTransNum & "';"
    dt = MyCommon.LWH_Select()
    If dt.Rows.Count = 0 Then
      MyCommon.QueryStr = "select ExtLocationCode, TransDate, TerminalNum, TransNum as POSTransNum, OfferID, AdjAmount from PointsAdj_Pending with (NoLock) where LogixTransNum ='" & LogixTransNum & "';"
      dt = MyCommon.LXS_Select
    End If
    If (dt.Rows.Count > 0) Then
      TxDetail.LogixTransNum = LogixTransNum
      TxDetail.TransNumber = MyCommon.NZ(dt.Rows(0).Item("POSTransNum"), "")
      tempDate = MyCommon.NZ(dt.Rows(0).Item("TransDate"), New Date(1980, 1, 1))
      TxDetail.TransDateStr = tempDate.ToString("MM/dd/yyyy")
      TxDetail.TransTimeStr = tempDate.ToString("hh:mm:ss tt")
      TxDetail.TransStore = MyCommon.NZ(dt.Rows(0).Item("ExtLocationCode"), "")
      TxDetail.TransTerminal = MyCommon.NZ(dt.Rows(0).Item("TerminalNum"), "")
      TxDetail.TransOffer = MyCommon.NZ(dt.Rows(0).Item("OfferID"), TxDetail.TransOffer)
      ProgDetail.AdjustmentAmount = MyCommon.NZ(dt.Rows(0).Item("AdjAmount"), ProgDetail.AdjustmentAmount)
    End If
  ElseIf (ClearEntry) Then
    ExtCardID = ""
    TxDetail.LogixTransNum = ""
    TxDetail.TransNumber = ""
    TxDetail.TransDateStr = ""
    TxDetail.TransTimeStr = ""
    TxDetail.TransStore = ""
    TxDetail.TransTerminal = ""
    TxDetail.TransOffer = 0
    ProgDetail.AdjustmentAmount = 0
  End If
  
  ' Process search request, if applicable 
  If TransTerms = "" Then
    TransTerms = Request.QueryString("transterms")
  End If
  If (TransTerms <> "") Then
    SearchFilter = " where (OfferID=" & MyCommon.Extract_Val(TransTerms) & " or InitialCardID='" & MyCryptLib.SQL_StringEncrypt(IIf(IDLength > 0, TransTerms.PadLeft(IDLength, "0"), TransTerms)) & "' " & _
                   "        or TransNum='" & TransTerms & "' or ExtLocationCode='" & TransTerms & "' " & _
                   "        or TerminalNum='" & TransTerms & "' "
    If Date.TryParse(TransTerms, tempDate) Then
      SearchFilter &= " or TransDate='" & tempDate.ToString & "' "
    End If
    SearchFilter &= ") and ViewInManualEntry = 1 "
    If (Request.QueryString("Parm1") <> "") Then
      HighlightedPKID = Request.QueryString("Parm1")
    End If
  Else
    SearchFilter = " where ViewInManualEntry = 1 "
  End If
  
  ' Process sort request, if applicable
  If (Request.QueryString("sortcol") <> "") Then SortCol = Request.QueryString("sortcol")
  If (Request.QueryString("sortdir") <> "") Then
    SortDir = Request.QueryString("sortdir")
  Else
    SortDir = "desc"
  End If
  
  MyCommon.QueryStr = "select PKID, OfferID, AdjAmount, TransDate, ExtLocationCode, TerminalNum, TransNum, PEND.CustomerPK, LogixTransNum, Note, InitialCardID as PrimaryExtID " & _
                      "from PointsAdj_Pending as PEND with (NoLock) " & _
                      "inner join Customers as CUST with (NoLock) on CUST.CustomerPK=PEND.CustomerPK " & _
                      " " & SearchFilter & " "
  dtTrans = MyCommon.LXS_Select
  sizeOfData = dtTrans.Rows.Count
  
  If PageNum = 0 Then PageNum = MyCommon.Extract_Val(Request.QueryString("pagenum"))
  If (linesPerPage * PageNum) > sizeOfData Then PageNum = PageNum - 1
  startPosition = PageNum * linesPerPage
  endPosition = IIf(sizeOfData < startPosition + linesPerPage, sizeOfData, startPosition + linesPerPage) - 1
  MorePages = IIf(sizeOfData > startPosition + linesPerPage, True, False)
  
  SortUrl = "CAM-customer-manual.aspx?CustPK=" & CustomerPK & "&amp;transterms=" & TransTerms & "&amp;pagenum=0"
%>
<form id="mainform" name="mainform" action="CAM-customer-manual.aspx">
<input type="hidden" id="WarningProgramID" name="WarningProgramID" value="<%Sendb(WarningProgramID)%>" />
<input type="hidden" name="actiontype" id="actiontype" value="" />
<input type="hidden" name="parm1" id="parm1" value="" />
<input type="hidden" name="parm2" id="parm2" value="" />
<input type="hidden" name="parm3" id="parm3" value="" />
<input type="hidden" name="parm4" id="parm4" value="" />
<input type="hidden" name="parm5" id="parm5" value="" />
<div id="intro">
  <h1 id="title">
  </h1>
  <div id="controls">
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
    If (StatusMessage <> "") Then
      Send("<div id=""infobar"" class=""green-background"">" & StatusMessage & "</div><br />")
    End If
    If (Request.QueryString("mode") = "summary") Then
      Send("<input type=""hidden"" id=""mode"" name=""mode"" value=""summary"" />")
      Send("<input type=""hidden"" id=""exiturl"" name=""exiturl"" value=""" & URLtrackBack & """ />")
    End If
  %>
  <div id="column">
    <div class="box" id="entry">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("customer-inquiry.transaction-detail", LanguageID))%>
        </span>
      </h2>
      <table summary="<%Sendb(Copient.PhraseLib.Lookup("customer-inquiry.transaction-detail", LanguageID))%>">
        <tr>
          <th style="color:red; width: 150px;">
            <label for="card"><%Sendb(Copient.PhraseLib.Lookup("term.card", LanguageID))%></label>
          </th>
          <th style="color:red; width: 100px;">
            <label for="transdate"><%Sendb(Copient.PhraseLib.Lookup("term.date", LanguageID))%></label>
          </th>
          <th style="color:red; width: 75px;">
            <label for="storeshort"><%Sendb(Copient.PhraseLib.Lookup("term.store", LanguageID))%></label>
            <span style="color:#000000;float:right;font-weight:normal;"><i><%Sendb(StrConv(Copient.PhraseLib.Lookup("term.or", LanguageID), VbStrConv.Lowercase) & "&nbsp;")%></i></span>
          </th>
          <th style="color:red; width: 75px;">
            <label for="store"><%Sendb(Copient.PhraseLib.Lookup("term.storecode", LanguageID))%></label>
          </th>
          <th style="color:red; width: 60px;">
            <label for="lane"><%Sendb(Copient.PhraseLib.Lookup("term.terminal", LanguageID))%></label>
          </th>
          <th style="color:red; width: 60px;">
            <label for="transaction"><%Sendb(Copient.PhraseLib.Lookup("customer-inquiry.txn", LanguageID))%></label>
          </th>
          <th>
            <%Sendb(Copient.PhraseLib.Lookup("term.action", LanguageID))%>
          </th>
        </tr>
        <% If LockCriteria Then%>
        <tr>
          <td>
            <input type="text" style="width:150px; color:Gray;" id="card" name="card" value="<%Sendb(ExtCardID) %>" readonly="readonly" />
          </td>
          <td>
            <input type="text" style="width:70px; color:Gray;" id="transdate" name="transdate" value="<%Sendb(TxDetail.TransDateStr) %>" readonly="readonly" />
          </td>
          <td>
            <input type="text" style="width:40px; color:Gray;" id="storeshort" name="storeshort" value="<%Sendb(StoreShort) %>" readonly="readonly" maxlength="4" />
          </td>
          <td>
            <input type="text" style="width:60px; color:Gray;" id="store" name="store" value="<%Sendb(TxDetail.TransStore) %>" readonly="readonly" maxlength="20" />
          </td>
          <td>
            <input type="text" style="width:60px; color:Gray;" id="lane" name="lane" value="<%Sendb(TxDetail.TransTerminal) %>" readonly="readonly" maxlength="4" />
          </td>
          <td>
            <input type="text" style="width:60px; color:Gray;" id="transaction" name="transaction" value="<%Sendb(TxDetail.TransNumber) %>" readonly="readonly" maxlength="128" />
          </td>
          <td>
            <input type="submit" id="transClear" name="transClear" value="<%Sendb(Copient.PhraseLib.Lookup("term.clear", LanguageID))%>" />
            <% If ShowSave Then%>
            <input type="submit" id="save" name="save" value="<%Sendb(Copient.PhraseLib.Lookup("term.save", LanguageID))%>" style="color: red;" />
            <% If ExecutePermitted Then%>
            <input type="submit" id="execute" name="execute" class="adjust" value="E" style="width: 30px;" />
            <% End If%>
            <% End If%>
          </td>
        </tr>
        <% Else%>
        <tr>
          <td>
            <input type="text" style="width:150px;" id="card" name="card" value="<%Sendb(ExtCardID) %>" />
          </td>
          <td>
            <input type="text" style="width:70px;" id="transdate" name="transdate" value="<%Sendb(TxDetail.TransDateStr) %>" />
            <img src="/images/calendar.png" class="calendar" id="transdate-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('transdate', event);" />
          </td>
          <td>
            <input type="text" style="width:40px;" id="storeshort" name="storeshort" value="<%Sendb(TxDetail.TransStore) %>" maxlength="4" />
          </td>
          <td>
            <input type="text" style="width:60px;" id="store" name="store" value="<%Sendb(TxDetail.TransStore) %>" maxlength="20" />
          </td>
          <td>
            <input type="text" style="width:60px;" id="lane" name="lane" value="<%Sendb(TxDetail.TransTerminal) %>" maxlength="4" />
          </td>
          <td>
            <input type="text" style="width:60px;" id="transaction" name="transaction" value="<%Sendb(TxDetail.TransNumber) %>" maxlength="128" />
          </td>
          <td>
            <input type="submit" id="transSearch" name="transSearch" value="<%Sendb(Copient.PhraseLib.Lookup("term.search", LanguageID))%>" />
          </td>
        </tr>
        <% End If%>
        <% If ShowSave Then%>
        <tr>
          <th colspan="2">
            <label for="transNote"><%Sendb(Copient.PhraseLib.Lookup("term.note", LanguageID))%></label>
          </th>
          <th>
            <label for="transTime"><%Sendb(Copient.PhraseLib.Lookup("term.time", LanguageID))%></label>
          </th>
          <th>
            <label for="adjustment"><%Sendb(Copient.PhraseLib.Lookup("term.adjustment", LanguageID))%></label>
          </th>
          <th>
            <label for="offerid"><%Sendb(Copient.PhraseLib.Lookup("term.offerid", LanguageID))%></label>
          </th>
          <td>
          </td>
        </tr>
        <tr>
          <td colspan="2">
            <textarea name="transNote" id="transNote" rows="3" cols="27"><%Sendb(TxDetail.Note)%></textarea>
          </td>
          <% If LockTime Then%>
          <td valign="top">
            <input type="text" style="width:60px; color:Gray;" id="transTime" name="transTime" value="<%Sendb(TxDetail.TransTimeStr)%>" maxlength="5" readonly="readonly" />
          </td>
          <% Else%>
          <td valign="top">
            <input type="text" style="width:60px;" id="transTime" name="transTime" value="<%Sendb(TxDetail.TransTimeStr)%>" maxlength="5" />
            <br />
            <small>(HH:mm, 24-hr)</small>
          </td>
          <% End If%>
          <% If ExecutePermitted Then%>
          <td valign="top">
            <input type="text" style="width:60px;" id="adjustment" name="adjustment" value="<%Sendb(IIf(ProgDetail.AdjustmentAmount <> 0, ProgDetail.AdjustmentAmount, "")) %>" maxlength="7" />
          </td>
          <td valign="top">
            <input type="text" style="width:60px;" id="offerid" name="offerid" value="<%Sendb(IIf(TxDetail.TransOffer > 0, TxDetail.TransOffer, "")) %>" maxlength="18" />
          </td>
          <% Else%>
          <td colspan="2">
          </td>
          <% End If%>
          <td>
          </td>
        </tr>
        <% End If%>
        <% If ResultsMessage.Trim <> "" Then%>
        <tr>
          <td colspan="7" class="green-background white" style="font-weight: bold;word-break:break-all;">
            <%Sendb(ResultsMessage)%>
          </td>
        </tr>
        <% End If%>
      </table>
      <hr class="hidden" />
    </div>
    <% If ShowResults Then%>
    <div class="box" id="transResults">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("customer-inquiry.transaction-results", LanguageID))%>
        </span>
      </h2>
      <%
        If LogixTransNums IsNot Nothing AndAlso LogixTransNums.Length > 0 Then
          ' add quotes around values for use in the queries
          ReDim LogixTransList(LogixTransNums.GetUpperBound(0))
          For i = 0 To LogixTransNums.GetUpperBound(0)
            LogixTransList(i) = "'" & LogixTransNums(i).Trim & "'"
          Next
          ' check if it's already in the holding area (i.e. PointsAdj_Pending)
          MyCommon.QueryStr = "select PKID, TransNum from PointsAdj_Pending with (NoLock) " & _
                              "where LogixTransNum in (" & String.Join(",", LogixTransList) & ");"
          dt = MyCommon.LXS_Select
          If dt.Rows.Count > 0 Then
            Send("<center><b>" & Copient.PhraseLib.Lookup("customer-manual.TransactionAlreadyAdded", LanguageID) & "</b></center>")
            Send("<br />")
            Send("<center><a href=""/logix/CAM/CAM-customer-manual.aspx?search=Search&searchterms=&CustPK=0&offerSearch=Search&transterms=" & MyCommon.NZ(dt.Rows(0).Item("TransNum"), "") & """>" & Copient.PhraseLib.Lookup("customer-manual.ClickToFind", LanguageID) & "</a></center>")
            ShowCreate = False
            Response.Status = "301 Moved Permanently"
            Response.AddHeader("Location", "/logix/CAM/CAM-customer-manual.aspx?search=Search&searchterms=&CustPK=0&offerSearch=Search&transterms=" & MyCommon.NZ(dt.Rows(0).Item("TransNum"), "") & "&Parm1=" & MyCommon.NZ(dt.Rows(0).Item("PKID"), 0))
          Else
            ' get the transaction information and then loop through it and write the rows
            MyCommon.QueryStr = "select top 10 LogixTransNum, CustomerPrimaryExtID, ExtLocationCode, TransDate, TerminalNum, POSTransNum " & _
                                "from TransHistory with (NoLock) " & _
                                "where LogixTransNum in (" & String.Join(",", LogixTransList) & ");"
            dt = MyCommon.LWH_Select
            If dt.Rows.Count > 0 Then
              Send("<table>")
              Send("  <tr>")
              Send("    <th style=""width:30px;""></th>")
              Send("    <th>" & Copient.PhraseLib.Lookup("term.card", LanguageID) & "</th>")
              Send("    <th>" & Copient.PhraseLib.Lookup("customer-inquiry.txn", LanguageID) & "</th>")
              Send("    <th>" & Copient.PhraseLib.Lookup("term.date", LanguageID) & "</th>")
              Send("    <th>" & Copient.PhraseLib.Lookup("term.store", LanguageID) & "</th>")
              Send("    <th>" & Copient.PhraseLib.Lookup("term.terminal", LanguageID) & "</th>")
              Send("  </tr>")
              For Each row In dt.Rows
                Send("<tr>")
                Send("  <td><input type=""button"" name=""select"" value=""..."" title=""Click to select this transaction."" onclick=""selectTrans('" & MyCommon.NZ(row.Item("LogixTransNum"), "") & "');"" /></td>")
                Send("  <td>" & MyCommon.NZ(row.Item("CustomerPrimaryExtID"), "") & "</td>")
                Send("  <td style=""word-wrap:break-word"">" & MyCommon.NZ(row.Item("POSTransNum"), "") & "</td>")
                Send("  <td>" & MyCommon.NZ(row.Item("TransDate"), "") & "</td>")
                Send("  <td>" & MyCommon.NZ(row.Item("ExtLocationCode"), "") & "</td>")
                Send("  <td>" & MyCommon.NZ(row.Item("TerminalNum"), "") & "</td>")
                Send("</tr>")
              Next
              Send("</table>")
              ShowCreate = True
            Else
              Send("<center>" & Copient.PhraseLib.Lookup("customer-manual.NoTransactionsFound", LanguageID) & "</center>")
              ShowCreate = True
            End If
          End If
        Else
          Send("<center>" & Copient.PhraseLib.Lookup("customer-manual.NoTransactionsFound", LanguageID) & "</center>")
          ShowCreate = True
        End If
        If ShowCreate Then
          Send("<br /><br class=""half""/><center><input type=""submit"" name=""createTrans"" value=""Create new transaction"" /></center>")
          Response.Status = "301 Moved Permanently"
          Response.AddHeader("Location", "/logix/CAM/CAM-customer-manual.aspx?card=" & Request.QueryString("card") & _
                             "&transdate=" & Request.QueryString("transdate") & _
                             "&storeshort=" & Request.QueryString("storeshort") & _
                             "&store=" & Request.QueryString("store") & _
                             "&lane=" & Request.QueryString("lane") & _
                             "&transaction=" & Request.QueryString("transaction") & _
                             "&createTrans=Create new transaction" & _
                             "&transterms=")
        End If
      %>
    </div>
    <% End If%>
    <div class="box" id="history">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("customer-inquiry.available-manual", LanguageID))%>
        </span>
      </h2>
      <%
        Send(" <div id=""listbar"">")
        Send("  <div id=""paginator"" style=""float:none;text-align:left;width:auto;"">")
        If (sizeOfData > 0) Then
          Send("   <input type=""text"" style=""font-family:arial;font-size:12px;"" id=""transterms"" name=""transterms"" class=""mediumshort"" value=""" & TransTerms & """ onkeydown=""submitTransSearch(event);"" />")
          Send("   <input type=""button"" style=""font-family:arial;font-size:12px;"" id=""btnOffer"" name=""btnOffer"" value=""" & Copient.PhraseLib.Lookup("term.search", LanguageID) & """ onclick=""searchTrans();"" />")
          Send("   <span style=""padding-left:60px;"">&nbsp;</span>")
          If (PageNum > 0) Then
            Send("   <span id=""first""><a href=""CAM-customer-manual.aspx?CustPK=" & CustomerPK & "&amp;pagenum=0&amp;transterms=" & TransTerms & "&amp;sortcol=" & SortCol & "&amp;sortdir=" & SortDir & """><b>|</b>◄" & Copient.PhraseLib.Lookup("term.first", LanguageID) & "</a>&nbsp;</span>")
            Send("   <span id=""previous""><a href=""CAM-customer-manual.aspx?CustPK=" & CustomerPK & "&amp;pagenum=" & PageNum - 1 & "&amp;transterms=" & TransTerms & "&amp;sortcol=" & SortCol & "&amp;sortdir=" & SortDir & """>◄" & Copient.PhraseLib.Lookup("term.previous", LanguageID) & "</a>&nbsp;</span>")
          Else
            Send("   <span id=""first""><b>|</b>◄" & Copient.PhraseLib.Lookup("term.first", LanguageID) & "&nbsp;</span>")
            Send("   <span id=""previous"">◄" & Copient.PhraseLib.Lookup("term.previous", LanguageID) & "&nbsp;</span>")
          End If
        Else
          Send("   <input type=""text"" class=""mediumshort"" id=""transterms"" name=""transterms"" onkeydown=""submitTransSearch(event);"" style=""font-family:arial;font-size:12px;"" value=""" & TransTerms & """ />")
          Send("   <input type=""button"" id=""btnTrans"" name=""btnTrans"" onclick=""searchTrans();"" style=""font-family:arial;font-size:12px;"" value=""" & Copient.PhraseLib.Lookup("term.search", LanguageID) & """ />")
          Send("   <span style=""padding-left:60px;"">&nbsp;</span>")
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
            Send("   <span id=""next""><a href=""CAM-customer-manual.aspx?CustPK=" & CustomerPK & "&amp;pagenum=" & PageNum + 1 & "&amp;transterms=" & TransTerms & "&amp;sortcol=" & SortCol & "&amp;sortdir=" & SortDir & """>" & Copient.PhraseLib.Lookup("term.next", LanguageID) & "►</a>&nbsp;</span>")
            Send("   <span id=""last""><a href=""CAM-customer-manual.aspx?CustPK=" & CustomerPK & "&amp;pagenum=" & (Math.Ceiling(sizeOfData / linesPerPage) - 1) & "&amp;transterms=" & TransTerms & "&amp;sortcol=" & SortCol & "&amp;sortdir=" & SortDir & """>" & Copient.PhraseLib.Lookup("term.last", LanguageID) & "►<b>|</b></a>&nbsp;</span>")
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
      <table summary="<% Sendb(Copient.PhraseLib.Lookup("customer-inquiry.available-manual", LanguageID)) %>" style="table-layout:auto;">
        <thead>
          <tr>
            <th align="left" style="width: 145px; text-align: center;" scope="col">
              <%Sendb(Copient.PhraseLib.Lookup("term.actions", LanguageID))%>
            </th>
            <th align="left" style="width: 60px;" scope="col">
              <% Sendb(Copient.PhraseLib.Lookup("term.OfferID", LanguageID))%>
            </th>
            <th align="left" style="width: 60px;" scope="col">
              <% Sendb(Copient.PhraseLib.Lookup("term.adjustment", LanguageID))%>
            </th>
            <th align="left" class="th-cardholder" scope="col">
              <a href="<% Sendb(SortUrl & "&amp;sortcol=PrimaryExtID&amp;sortdir=" & IIf(SortDir = "desc", "asc", "desc")) %>">
                <% Sendb(Copient.PhraseLib.Lookup("term.card", LanguageID))%>
              </a>
              <%
                If SortCol = "PrimaryExtID" Then
                  If SortDir = "asc" Then
                    Sendb("<span class=""sortarrow"">&#9660;</span>")
                  Else
                    Sendb("<span class=""sortarrow"">&#9650;</span>")
                  End If
                End If
              %>
            </th>
            <th align="center" class="th-shortdate" scope="col">
              <a href="<% Sendb(SortUrl & "&amp;sortcol=TransDate&amp;sortdir=" & IIf(SortDir = "desc", "asc", "desc")) %>">
                <% Sendb(Copient.PhraseLib.Lookup("term.datetime", LanguageID))%>
              </a>
              <%
                If SortCol = "TransDate" Then
                  If SortDir = "asc" Then
                    Sendb("<span class=""sortarrow"">&#9660;</span>")
                  Else
                    Sendb("<span class=""sortarrow"">&#9650;</span>")
                  End If
                End If
              %>
            </th>
            <th align="center" class="th-store" scope="col">
              <a href="<% Sendb(SortUrl & "&amp;sortcol=ExtLocationCode&amp;sortdir=" & IIf(SortDir = "desc", "asc", "desc")) %>">
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
            <th align="center" class="th-lane" scope="col">
              <a href="<% Sendb(SortUrl & "&amp;sortcol=TerminalNum&amp;sortdir=" & IIf(SortDir = "desc", "asc", "desc")) %>">
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
            <th align="center" class="th-txn" scope="col">
              <a href="<% Sendb(SortUrl & "&amp;sortcol=TransNum&amp;sortdir=" & IIf(SortDir = "desc", "asc", "desc")) %>">
                <% Sendb(Copient.PhraseLib.Lookup("customer-inquiry.txn", LanguageID))%>
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
          </tr>
        </thead>
        <tbody>
          <%
            Dim transRows As ArrayList
            If (dtTrans.Rows.Count > 0) Then
              transCt = 0
              transRows = GetSubList(dtTrans, startPosition, endPosition)
              If transRows.Count > 0 Then
                For Each row In transRows
                  transCt += 1
                  PKID = MyCommon.NZ(row.Item("PKID"), 0)
                  TransNum = MyCommon.NZ(row.Item("TransNum"), "")
                  OfferID = MyCommon.NZ(row.Item("OfferID"), 0)
                  AdjAmount = MyCommon.NZ(row.Item("AdjAmount"), 0)
                  CustomerPK = MyCommon.NZ(row.Item("CustomerPK"), 0)
                  LogixTransNum = MyCommon.NZ(row.Item("LogixTransNum"), "").ToString.Trim
                  TransNote = MyCommon.NZ(row.Item("Note"), "")
                  
                  If PKID = HighlightedPKID Then
                    Send("<tr id=""tr" & PKID & """ class=""rowHighlighted"">")
                  Else
                    Send("<tr id=""tr" & PKID & """>")
                  End If
                  If ExecutePermitted Then
                    Send("  <td><input type=""button"" class=""view"" name=""btnView"" id=""btnView"" title=""" & Copient.PhraseLib.Lookup("customer-inquiry.select-offer", LanguageID) & """ " & " value=""..."" onclick=""launchOffers('" & LogixTransNum & "','" & MyCommon.NZ(row.Item("TransDate"), "") & "', '" & MyCommon.NZ(row.Item("ExtLocationCode"), "") & "'," & transCt & ");"" />")
                    Send("      <input type=""button"" class=""adjust"" name=""btnExecute"" id=""btnExecute"" title=""" & Copient.PhraseLib.Lookup("customer-inquiry.apply-adjustment", LanguageID) & """ " & " value=""E"" onclick=""applyTransaction(" & PKID & ",'" & TransNum & "'," & transCt & ");"" />")
                    Send("      <input type=""button"" class=""adjust"" name=""btnSave"" id=""btnSave"" title=""" & Copient.PhraseLib.Lookup("term.save", LanguageID) & """ " & " value=""S"" onclick=""saveTransaction(" & PKID & ",'" & TransNum & "'," & transCt & ");"" />")
                    Send("      <input type=""button"" class=""ex"" name=""btnDelete"" id=""btnDelete"" title=""" & Copient.PhraseLib.Lookup("customer-inquiry.remove", LanguageID) & """ " & " value=""X"" onclick=""deleteTransaction(" & PKID & ",'" & TransNum & "');"" />")
                    If TransNote.Trim <> "" Then
                                  Send("      <img src=""/images/notes-some.png"" style="""" onclick=""showNote(" & transCt & ");"" title=""" & Copient.PhraseLib.Lookup("term.note", LanguageID) & """ />")
                    End If
                    Send("  </td>")
                    Send("  <td><input type=""text"" name=""OfferID" & transCt & """ id=""OfferID" & transCt & """ style=""width:60px;"" value=""" & IIf(OfferID > 0, OfferID, "") & """ maxlength=""18"" /></td>")
                    Send("  <td><input type=""text"" name=""Adjustment" & transCt & """ id=""Adjustment" & transCt & """ style=""width:60px;text-align:right;"" value=""" & IIf(AdjAmount <> 0, AdjAmount, "") & """ maxlength=""7"" /></td>")
                  Else
                    Send("  <td>")
                    If TransNote.Trim <> "" Then
                      Send("      <img src=""/images/notes-some.png"" style=""position:relative;top:4px;"" onclick=""showNote(" & transCt & ");"" title=""" & Copient.PhraseLib.Lookup("term.note", LanguageID) & """ />")
                    End If
                    Send("  </td>")
                    Send("  <td></td>")
                    Send("  <td></td>")
                  End If
                  Send("  <td>" & MyCryptLib.SQL_StringDecrypt(row.Item("PrimaryExtID").ToString()) & "<input type=""hidden"" name=""custPK" & transCt & """ id=""custPK" & transCt & """ value=""" & CustomerPK & """ /></td>")
                  Send("  <td>" & Format(MyCommon.NZ(row.Item("TransDate"), ""), "dd MMM yyyy, HH:mm") & "<input type=""hidden"" name=""tdate" & transCt & """ id=""tdate" & transCt & """ value=""" & MyCommon.NZ(row.Item("TransDate"), "") & """ /></td>")
                  Send("  <td>" & MyCommon.NZ(row.Item("ExtLocationCode"), "") & "</td>")
                  Send("  <td>" & MyCommon.NZ(row.Item("TerminalNum"), "") & "</td>")
                  Send("  <td  style=""word-break:break-all"">" & MyCommon.NZ(row.Item("TransNum"), "") & "</td>")
                  Send("</tr>")
                  
                  If TransNote.Trim <> "" Then
                    Send("<tr id=""note" & transCt & """ style=""display:none;"">")
                    Send("  <td colspan=""8"" style=""background:#dddddd;"">" & TransNote & "</td>")
                    Send("</tr>")
                  End If
                Next
              Else
                Send("<tr>")
                Send("  <td colspan=""8"" style=""text-align:center""><i>" & "No Manual Entries" & "</i></td>")
                Send("</tr>")
              End If
            Else
              Send("<tr>")
              Send("  <td colspan=""8"" style=""text-align:center""><i>" & Copient.PhraseLib.Lookup("customer-manual.NoManualEntries", LanguageID) & "</i></td>")
              Send("</tr>")
            End If
          %>
        </tbody>
      </table>
      <hr class="hidden" id="HR1" onclick="return HR1_onclick()" />
    </div>
  </div>
  <div id="datepicker" class="dpDiv">
    <br clear="all" />
  </div>
</div>
</form>

<script runat="server">
  Dim MyCommon As New Copient.CommonInc
  
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
  
  Function GetProgramsForOffer(ByVal OfferID As Long, ByRef infoMessage As String) As Integer()
    Dim Programs(-1) As Integer
    Dim dt As DataTable
    Dim ROID As Long
    Dim i As Integer
    
    Try
      If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
      
      MyCommon.QueryStr = "select RO.RewardOptionID from CPE_Incentives as INC with (NoLock) " & _
                          "inner join CPE_RewardOptions as RO with (NoLock) on RO.IncentiveID = INC.IncentiveID " & _
                          "where INC.Deleted=0 and RO.Deleted=0 and INC.EngineID=6 and INC.IncentiveID=" & OfferID
      dt = MyCommon.LRT_Select
      If dt.Rows.Count > 0 Then
        ROID = MyCommon.NZ(dt.Rows(0).Item("RewardOptionID"), 0)
        
        MyCommon.QueryStr = "select ProgramID from CPE_IncentivePointsGroups with (NoLock) where Deleted=0 and RewardOptionID = " & ROID & _
                            " union " & _
                            "select DP.ProgramID from CPE_Deliverables as DEL with (NoLock) " & _
                            "inner join CPE_DeliverablePoints as DP with (NoLock) on DP.DeliverableID = DEL.DeliverableID " & _
                            "where DEL.Deleted=0 and DP.Deleted=0 and DEL.RewardOptionID=" & ROID
        dt = MyCommon.LRT_Select
        If dt.Rows.Count > 0 Then
          ReDim Programs(dt.Rows.Count - 1)
          For i = 0 To Programs.GetUpperBound(0)
            Programs(i) = MyCommon.NZ(dt.Rows(i).Item("ProgramID"), 0)
          Next
        End If
      Else
        infoMessage = Copient.PhraseLib.Detokenize("customer-manual.NotAValidCAMOffer", LanguageID, OfferID)  'Offer {0} is not a valid CAM offer.
      End If
    Catch ex As Exception
      infoMessage = Copient.PhraseLib.Detokenize("error.encountered", LanguageID, ex.ToString)
    Finally
      If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
    End Try
    
    Return Programs
  End Function
  
  Function IsValidAdjustment(ByVal TxDetail As Copient.CAM.TransactionDetail, ByVal CustomerPK As Long, ByVal ProgDetail As Copient.CAM.ProgramDetail, _
                             ByRef infoMessage As String, ByRef WarningProgramID As Long) As Boolean
    Dim ValidAdj As Boolean = False
    Dim MyCAM As New Copient.CAM
    Dim MyPoints As New Copient.Points
    Dim PointsBal, AdjustAmt, WarningLimit As Long
    Dim ErrMessage As String = ""
    Dim TempDate As New Date
    
    Try
      If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
    
      AdjustAmt = ProgDetail.AdjustmentAmount
      PointsBal = MyPoints.GetBalance(CustomerPK, ProgDetail.ProgramID)
      WarningLimit = MyCAM.GetMaxAdjustment(TxDetail.TransOffer, ProgDetail.ProgramID)
      
      If ProgDetail.ProgramID <= 0 Then
        infoMessage &= Copient.PhraseLib.Detokenize("customer-manual.InvalidProgramID", LanguageID, ProgDetail.ProgramID)  'Invalid Program ID: {0}
      ElseIf AdjustAmt = 0 Then
        infoMessage &= Copient.PhraseLib.Lookup("customer-manual.ZeroPointAdjustments", LanguageID)  'Adjustments of zero points are not permitted.
      ElseIf (PointsBal + AdjustAmt < 0) Then
        infoMessage &= Copient.PhraseLib.Detokenize("customer-manual.NegativePointsBalance", LanguageID, (-PointsBal))  'This adjustment would cause a negative points balance.  Maximum adjustment is {0} points.
      ElseIf (Math.Abs(AdjustAmt) > WarningLimit) Then
        ' if a warning has already been issued then allow the adjustment
        If WarningProgramID = 0 Then
          infoMessage &= Copient.PhraseLib.Lookup("customer-manual.AdjustmentWarningLimit", LanguageID)  'This adjustment amount will cross the warning limit for maximum points adjusted for this offer. Click the "E" button again to confirm you wish to make this adjustment.
          WarningProgramID = ProgDetail.ProgramID
        Else
          ValidAdj = True
        End If
      ElseIf Not Date.TryParse(TxDetail.TransDateStr, TempDate) OrElse Not ValidateOffer(TxDetail.TransOffer, TempDate, ErrMessage) Then
        infoMessage &= ErrMessage
      Else
        ValidAdj = True
      End If
    Catch ex As Exception
      ValidAdj = False
      infoMessage = Copient.PhraseLib.Detokenize("error.encountered", LanguageID, ex.ToString)
    Finally
      If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
    End Try
    
    Return ValidAdj
  End Function
  
  Function AdjustPoint(ByVal AdminUserID As String, ByVal CustomerExtID As String, ByVal CustomerPK As Long, _
                       ByVal TxDetail As Copient.CAM.TransactionDetail, ByVal ProgDetail As Copient.CAM.ProgramDetail, _
                       ByVal SessionID As String, ByVal SelectedOfferID As Long) As String
    
    Dim ProgramID As Long
    Dim AdjustAmt As Long
    Dim LogText As String = ""
    Dim bNeedsRollback As Boolean = False
    Dim MyCam As New Copient.CAM
    Dim MyPoints As New Copient.Points
    Dim RetMsg As String = ""
    Dim Fields As New Copient.CommonInc.ActivityLogFields
    Dim AssocLinks(-1) As Copient.CommonInc.ActivityLink
    Dim Offers(-1) As Copient.Offer
    Dim i As Integer
    Dim PreAdjustBal As New Decimal(0)
    Dim ProgramName As String = ""
    Dim dt As System.Data.DataTable
    
    Try
      If (MyCommon.LXSadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixXS()
      If (MyCommon.LRTadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixRT()
      If (MyCommon.LEXadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixEX()
      
      ProgramID = ProgDetail.ProgramID
      AdjustAmt = ProgDetail.AdjustmentAmount
      
      MyCommon.QueryStr = "begin transaction"
      MyCommon.LXS_Execute()
      MyCommon.LEX_Execute()
      
      If (CustomerPK > 0 AndAlso ProgramID > 0 AndAlso AdjustAmt <> 0) Then
        PreAdjustBal = New Decimal(MyPoints.GetBalance(CustomerPK, ProgramID))
        
        MyCommon.QueryStr = "dbo.pa_CPE_TU_InsertData_PA"
        MyCommon.Open_LXSsp()
        MyCommon.LXSsp.Parameters.Add("@TableNum", SqlDbType.VarChar, 4).Value = "4"
        MyCommon.LXSsp.Parameters.Add("@Operation", SqlDbType.VarChar, 2).Value = "1"
        MyCommon.LXSsp.Parameters.Add("@Col1", SqlDbType.VarChar, 255).Value = ProgramID.ToString
        MyCommon.LXSsp.Parameters.Add("@Col2", SqlDbType.VarChar, 255).Value = CustomerPK.ToString
        MyCommon.LXSsp.Parameters.Add("@Col3", SqlDbType.VarChar, 255).Value = AdjustAmt.ToString
        MyCommon.LXSsp.Parameters.Add("@Col4", SqlDbType.VarChar, 255).Value = MyCommon.Extract_Val(TxDetail.TransROID)
        MyCommon.LXSsp.Parameters.Add("@Col5", SqlDbType.VarChar, 255).Value = 2 ' CustomerTypeID
        MyCommon.LXSsp.Parameters.Add("@Col6", SqlDbType.VarChar, 255).Value = TxDetail.LogixTransNum
        MyCommon.LXSsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = -9
        MyCommon.LXSsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = -9
        MyCommon.LXSsp.Parameters.Add("@WaitingACK", SqlDbType.Int).Value = 0
        MyCommon.LXSsp.ExecuteNonQuery()
        MyCommon.Close_LXSsp()
        
        MyCommon.Open_LogixRT()
        MyCommon.QueryStr = "select ProgramName from PointsPrograms with (NoLock) where ProgramID=" & ProgramID & ";"
        dt = MyCommon.LRT_Select
        If dt.Rows.Count > 0 Then
          ProgramName = MyCommon.NZ(dt.Rows(0).Item("ProgramName"), "")
        End If
        MyCommon.Close_LogixRT()
        
        If Not MyCam.SendPointsToIssuance(TxDetail, ProgDetail, CustomerExtID, AdminUserID, MyCommon) Then
          bNeedsRollback = True
          RetMsg = Copient.PhraseLib.Lookup("customer-manual.IssuanceError", LanguageID)
        Else
          Try
            LogText = Copient.PhraseLib.Lookup("history.customer-adjust-points", LanguageID) & " " & ProgramID
            If ProgramName <> "" Then
              LogText &= " (""" & ProgramName & """)"
            End If
            LogText &= " " & StrConv(Copient.PhraseLib.Lookup("term.by", LanguageID), VbStrConv.Lowercase) & " " & AdjustAmt
            
            ' find all the offers associated to the adjusted points program
			If (MyCommon.Fetch_SystemOption(321) = "0") Then
            Offers = MyPoints.GetAssociatedOffers(ProgramID)
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If Offers.Length > 0 Then
              ReDim AssocLinks(Offers.GetUpperBound(0))
              For i = 0 To Offers.GetUpperBound(0)
                AssocLinks(i) = New Copient.CommonInc.ActivityLink()
                AssocLinks(i).LinkID = Offers(i).GetOfferID
                AssocLinks(i).LinkTypeID = 1 ' Offers
                AssocLinks(i).Selected = (Offers(i).GetOfferID = SelectedOfferID)
              Next
            End If
			End if

            ' log the adjustment and any associated offers
            Fields.ActivityTypeID = 25
            Fields.ActivitySubTypeID = 12
            Fields.LinkID = CustomerPK
            Fields.AdminUserID = AdminUserID
            Fields.Description = LogText
            Fields.LinkID2 = MyCommon.Extract_Val(ProgramID)
            Fields.ActivityValue = AdjustAmt.ToString
            Fields.AssociatedLinks = AssocLinks
            Fields.SessionID = SessionID
            Fields.PreAdjustBalance = PreAdjustBal
            Fields.Adjustment = New Decimal(AdjustAmt)
            Fields.PostAdjustBalance = Decimal.Add(Fields.PreAdjustBalance, Fields.Adjustment)
            MyCommon.Activity_Log3(Fields)
          Catch ex As Exception
            MyCommon.Write_Log("CAM.txt", "Failed to log CAM manual points adjustment for the following reason: " & ex.ToString, True)
          End Try
        End If
      Else
        RetMsg = Copient.PhraseLib.Detokenize("customer-manual.UnableToAdjust", LanguageID, CustomerPK, ProgramID, AdjustAmt)  'Unable to adjust points balance. CustomerPK: {0}, ProgramID: {1}, Adjustment: {2}.
      End If
    Catch ex As Exception
      bNeedsRollback = True
      RetMsg = ex.ToString
    Finally
      If bNeedsRollback Then
        MyCommon.QueryStr = "rollback transaction"
      Else
        MyCommon.QueryStr = "commit transaction"
      End If
      MyCommon.LXS_Execute()
      MyCommon.LEX_Execute()
      
      MyCommon.Close_LogixRT()
      MyCommon.Close_LogixXS()
      MyCommon.Close_LogixEX()
    End Try
    
    Return RetMsg
  End Function
  
  Function SubmitPoint(ByVal AdminUserID As String, ByVal CustomerPK As Long, _
                       ByVal TxDetail As Copient.CAM.TransactionDetail, ByVal ProgDetail As Copient.CAM.ProgramDetail) As String
    Dim RetMsg As String = ""
    Dim ProgramID As Long
    Dim AdjustAmt As Long
    Dim Note As String = ""
    
    ProgramID = ProgDetail.ProgramID
    AdjustAmt = ProgDetail.AdjustmentAmount
    
    Try
      If (MyCommon.LXSadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixXS()
      If (MyCommon.LRTadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixRT()
      
      If TxDetail.TransDateStr = "" Then TxDetail.TransDateStr = "1980-01-01"
      If TxDetail.Note IsNot Nothing Then TxDetail.Note = Left(TxDetail.Note, 200)
      
      ' write this points adjustment to the pending adjustment table
      MyCommon.QueryStr = "insert into PointsAdj_Pending with (RowLock) (LogixTransNum, TransNum, TransDate, ExtLocationCode, TerminalNum, " & _
                          "                    CustomerPK, ProgramID, OfferID, AdjAmount, CreateDate, CreatedBy, Note) " & _
                          " values ('" & TxDetail.LogixTransNum & "', '" & TxDetail.TransNumber & "', " & _
                          "   '" & TxDetail.TransDateStr & "', '" & TxDetail.TransStore & "', '" & TxDetail.TransTerminal & "', " & _
                          "    " & CustomerPK & ", " & ProgramID & ", " & TxDetail.TransOffer & ", " & _
                          "    " & AdjustAmt & ", getdate(), " & AdminUserID & ",'" & MyCommon.Parse_Quotes(TxDetail.Note) & "');"
      
      MyCommon.LXS_Execute()
      If MyCommon.RowsAffected <= 0 Then
        RetMsg = Copient.PhraseLib.Lookup("customer-manual.ErrorEncounteredAdjusting", LanguageID)
      Else
        MyCommon.Activity_Log2(25, 12, CustomerPK, AdminUserID, Copient.PhraseLib.Lookup("customer-manual.CreatedInHoldingArea", LanguageID), ProgDetail.ProgramID, ProgDetail.AdjustmentAmount)
      End If
    Catch ex As Exception
      RetMsg = ex.ToString
    Finally
      MyCommon.Close_LogixRT()
      MyCommon.Close_LogixXS()
    End Try
    
    Return RetMsg
  End Function
  
  Public Function AppendTimeToDate(ByVal DateString As String, ByVal TimeString As String) As String
    Dim NewDateString As String = DateString
    Dim NewDate As New Date
    
    If DateString IsNot Nothing AndAlso TimeString IsNot Nothing Then
      If Date.TryParse(DateString & " " & TimeString.Trim, NewDate) Then
        NewDateString = NewDate.ToString("yyyy-MM-dd HH:mm:ss")
      End If
    End If
    
    Return NewDateString
  End Function
  
  Function ValidateOffer(ByVal OfferID As Long, ByVal TransDate As Date, ByRef ErrMessage As String) As Boolean
    Dim dt As DataTable
    
    Try
      If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
      
      ErrMessage = ""
      
      MyCommon.QueryStr = "select IncentiveID, EngineID, StartDate, EndDate, dateadd(d, 1, EndDate) as EvalEndDate, Deleted " & _
                          "from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID
      dt = MyCommon.LRT_Select
      If dt.Rows.Count > 0 Then
        If TransDate < MyCommon.NZ(dt.Rows(0).Item("StartDate"), New Date(2070, 10, 8)) Then
          ErrMessage = "Transaction date of " & TransDate.ToShortDateString & " occurs before the offer's " & _
                       "start date of " & MyCommon.NZ(dt.Rows(0).Item("StartDate"), New Date(2070, 10, 8))
        End If
        If Not (TransDate < MyCommon.NZ(dt.Rows(0).Item("EvalEndDate"), New Date(1980, 1, 1))) Then
          ErrMessage = "Transaction date of " & TransDate.ToShortDateString & " occurs after the offer's " & _
                        "end date of " & MyCommon.NZ(dt.Rows(0).Item("EndDate"), New Date(1980, 1, 1))
        End If
        If MyCommon.NZ(dt.Rows(0).Item("EngineID"), 0) <> 6 Then ErrMessage = "Offer " & OfferID & " is not a CAM offer"
        If MyCommon.NZ(dt.Rows(0).Item("Deleted"), True) Then ErrMessage = "Offer " & OfferID & " is a deleted offer."
      Else
        ErrMessage = Copient.PhraseLib.Detokenize("error.OfferNotFound", LanguageID, OfferID)  'Offer {0} not found.
      End If
      
    Catch ex As Exception
      ErrMessage = ex.ToString
    Finally
      If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
    End Try
    
    Return (ErrMessage = "")
  End Function
</script>

<%
done:
  Send_BodyEnd("mainform", "card")
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  MyCommon.Close_LogixWH()
  MyCommon = Nothing
  Logix = Nothing
%>
