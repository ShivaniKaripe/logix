<%@ Page Language="vb" Debug="true" CodeFile="../LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: CAM-point-adjust-program.aspx 
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
  
    Dim MyCommon As New Copient.CommonInc
    Dim MyCryptLib As New Copient.CryptLib
  Dim MyCAM As New Copient.CAM
  Dim TxDetail As New Copient.CAM.TransactionDetail
  Dim ProgDetail As New Copient.CAM.ProgramDetail
  Dim MyLookup As New Copient.CustomerLookup()
  Dim AdminUserID As Long
  Dim Logix As New Copient.LogixInc
  Dim dt As DataTable = Nothing
  Dim row As DataRow
  Dim CustomerExtID As String = ""
  Dim CustomerPK As Long
  Dim OfferID As Long
  Dim PageTitle As String = ""
  Dim OfferName As String = ""
  Dim ProgramID As Long
  Dim ProgramName As String = ""
  Dim ProgramDesc As String = ""
  Dim ProgramIsExternal As Boolean = False
  Dim AdjustPermitted As Boolean = False
  Dim ExecutePermitted As Boolean = False
  Dim EarnedROID As Integer = 0
  Dim EarnedCMOffer As Integer = 0
  Dim OfferDesc As String = ""
  Dim InfoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim HHPK As Integer = 0
  Dim i As Integer = 0
  Dim KeyCt As Integer = 0
  Dim RefreshParent As Boolean = False
  Dim CustFirstName As String = ""
  Dim CustMiddleName As String = ""
  Dim CustLastName As String = ""
  Dim LogixTransNum As String = ""
  Dim TransDate As New Date(1980, 1, 1)
  Dim TransStore As String = ""
  Dim TransTerminal As String = ""
  Dim TransNumber As String = ""
  Dim LogixTransNums(-1) As String
  Dim ErrorMessage As String = ""
  Dim WarningProgramID As Long = 0
  Dim FocusElem As String = ""
  Dim ShowResults As Boolean = False
  Dim LogixTransList(-1) As String
  Dim CustDetail As New Copient.Customer
  Dim TempAdjAmount As Integer = 0
  Dim OfferStart As String = ""
  Dim OfferEnd As String = ""
  Dim SessionID As String = ""
  Dim CardPK As Long
  

  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "CAM-point-adjust-program.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  MyCommon.Open_LogixWH()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  MyLookup.SetAdminUserID(AdminUserID)
  MyLookup.SetLanguageID(LanguageID)
  
  AdjustPermitted = Logix.UserRoles.EditPointsBalances
  ExecutePermitted = Logix.UserRoles.ExecuteCAMAdjustment

  If Session("SessionID") IsNot Nothing AndAlso Session("SessionID").ToString.Trim <> "" Then
    SessionID = Session("SessionID").ToString.Trim
    CopientNotes &= "SessionID: " & SessionID
  End If

  CustomerExtID = Request.QueryString("CustomerExtId")
  CustomerPK = MyCommon.Extract_Val(Request.QueryString("CustomerPK"))
  CardPK = MyCommon.Extract_Val(Request.QueryString("CardPK"))
  If CardPK = 0 Then
    CardPK = MyLookup.FindCardPK(CustomerPK, 2)
  End If
  
  ProgramID = MyCommon.Extract_Val(Request.QueryString("ProgramID"))
  LogixTransNum = Request.QueryString("LogixTransNum")
  If LogixTransNum Is Nothing Then LogixTransNum = ""
  OfferID = MyCommon.Extract_Val(Request.QueryString("TransOffer"))
  FocusElem = Request.QueryString("FocusElem")
  If FocusElem Is Nothing Then FocusElem = ""
  
  
  MyCommon.QueryStr = "select ProgramName, Description, ExternalProgram from PointsPrograms where ProgramID=" & ProgramID & ";"
  dt = MyCommon.LRT_Select()
  If (dt.Rows.Count > 0) Then
    ProgramName = MyCommon.NZ(dt.Rows(0).Item("ProgramName"), "")
    ProgramDesc = MyCommon.NZ(dt.Rows(0).Item("Description"), "")
    ProgramIsExternal = MyCommon.NZ(dt.Rows(0).Item("ExternalProgram"), False)
  End If
  
  ' handle transaction searching and creating
  If (Request.QueryString("btnTransSearch") <> "" OrElse Request.QueryString("btnTransCreate") <> "") Then
    TransNumber = Request.QueryString("TransNum")
    TransStore = Request.QueryString("TransStore")
    TransTerminal = Request.QueryString("TransTerminal")
    If Not (Date.TryParse(Request.QueryString("TransDate"), TransDate)) OrElse TransDate < Now.AddYears(-100) Then
      InfoMessage = Copient.PhraseLib.Lookup("customer-inquiry.InvalidDateOrFormat", LanguageID)
      TransDate = New Date(1980, 1, 1)
    End If
    
    TxDetail.TransNumber = TransNumber
    TxDetail.TransDateStr = TransDate.ToString
    TxDetail.TransStore = TransStore
    TxDetail.TransTerminal = TransTerminal
    TxDetail.TransOffer = OfferID

    If (InfoMessage = "" AndAlso Request.QueryString("btnTransSearch") <> "") Then
      If (CustomerExtID <> "") Then
        Try
          LogixTransNums = MyCAM.FindCustomerTransaction(CustomerExtID, TxDetail)
          If LogixTransNums.Length = 0 Then
            InfoMessage = Copient.PhraseLib.Lookup("CAM-point-adjust.TransactionNotFound", LanguageID)
          ElseIf LogixTransNums.Length = 1 Then
            LogixTransNum = LogixTransNums(0).Trim
          ElseIf LogixTransNums.Length > 1 Then
            ShowResults = True
          End If
        Catch ex As Exception
          InfoMessage = Copient.PhraseLib.Detokenize("customer-manual.CreateError", LanguageID, ex.ToString)
        End Try
      Else
        InfoMessage = Copient.PhraseLib.Lookup("CAM-point-adjust.CardNumberNotSent", LanguageID)
      End If
    
    ElseIf (InfoMessage = "" AndAlso Request.QueryString("btnTransCreate") <> "") Then
      If TransDate > Date.Now Then
        InfoMessage = Copient.PhraseLib.Lookup("customer-manual.InvalidFutureDate", LanguageID)
      Else
        If (TransNumber <> "" AndAlso TransStore <> "" AndAlso TransTerminal <> "" AndAlso Request.QueryString("TransDate") <> "") Then
          Try
            ProgDetail.ProgramID = MyCommon.Extract_Val(Request.QueryString("ProgramID"))
            ProgDetail.AdjustmentAmount = MyCommon.Extract_Val(Request.QueryString("Adjust"))
          
            TxDetail.ViewInManualEntry = False
            LogixTransNum = MyCAM.CreateCustomerTransaction(CustomerExtID, CustomerPK, AdminUserID, TxDetail, ProgDetail)
          Catch sdneEx As Copient.StoreDoesNotExistException
            InfoMessage = Copient.PhraseLib.Detokenize("customer-manual.StoreNotFound", LanguageID, sdneEx.GetStore())
          Catch teEx As Copient.TransactionExistsException
            LogixTransNums = teEx.GetLogixTransNums
            If LogixTransNums.Length = 1 Then
              ' if only one match exists then simply use that transaction
              LogixTransNum = LogixTransNums(0)
            Else
              InfoMessage = Copient.PhraseLib.Lookup("customer-manual.MultipleMatches", LanguageID)
            End If
          Catch gtEx As Copient.GeneralTransactionException
            InfoMessage = gtEx.GetErrorMessage
          Catch ex As Exception
            InfoMessage = Copient.PhraseLib.Detokenize("customer-manual.CreateError", LanguageID, ex.ToString)
          End Try

        Else
          InfoMessage = Copient.PhraseLib.Lookup("customer-manual.EnterSearchCriteria", LanguageID)
        End If
      End If
      
    End If
  End If

  ' when necessary, load transaction information
  If (LogixTransNum.Trim <> "") Then
    MyCommon.QueryStr = "select ExtLocationCode, TransDate, TerminalNum, POSTransNum from TransHistory with (NoLock) where LogixTransNum ='" & LogixTransNum.Trim & "';"
    dt = MyCommon.LWH_Select()

    If dt.Rows.Count = 0 Then
      MyCommon.QueryStr = "select ExtLocationCode, TransDate, TerminalNum, TransNum as POSTransNum, OfferID from PointsAdj_Pending with (NoLock) where LogixTransNum ='" & LogixTransNum & "';"
      dt = MyCommon.LXS_Select
    End If

    If (dt.Rows.Count > 0) Then
      TransStore = MyCommon.NZ(dt.Rows(0).Item("ExtLocationCode"), "")
      TransDate = MyCommon.NZ(dt.Rows(0).Item("TransDate"), Date.Parse("1/1/1980"))
      TransTerminal = MyCommon.NZ(dt.Rows(0).Item("TerminalNum"), "")
      TransNumber = MyCommon.NZ(dt.Rows(0).Item("POSTransNum"), "")
      'OfferID = MyCommon.NZ(dt.Rows(0).Item("OfferID"), 0)
      
      TxDetail.LogixTransNum = LogixTransNum.Trim
      TxDetail.TransNumber = TransNumber
      TxDetail.TransDateStr = TransDate.ToString
      TxDetail.TransStore = TransStore
      TxDetail.TransTerminal = TransTerminal
      TxDetail.TransOffer = OfferID
    End If
  End If
  
  ' load customer information
  MyCommon.QueryStr = "select CustomerPK, FirstName, MiddleName, LastName from Customers with (NoLock) " & _
                        "where CustomerPK=" & CustomerPK & ";"
  dt = MyCommon.LXS_Select()
  If (dt.Rows.Count > 0) Then
    CustomerPK = dt.Rows(0).Item("CustomerPK")
    CustFirstName = MyCommon.NZ(dt.Rows(0).Item("FirstName"), "")
    CustMiddleName = MyCommon.NZ(dt.Rows(0).Item("MiddleName"), "")
    CustLastName = MyCommon.NZ(dt.Rows(0).Item("LastName"), "")
    
    MyCommon.QueryStr = "select top 1 ExtCardID from CardIDs with (NoLock) where CardTypeID=2 and CustomerPK=" & CustomerPK
    dt = MyCommon.LXS_Select
    If dt.Rows.Count > 0 Then
      CustomerExtID = MyCryptLib.SQL_StringDecrypt(dt.Rows(0).Item("ExtCardID").ToString())
    End If
  Else
    InfoMessage = Copient.PhraseLib.Detokenize("CAM-point-adjust.UnableToFindCardnumber", LanguageID, CustomerExtID)
  End If
    
  If (Request.QueryString("save") <> "") Then
    WarningProgramID = MyCommon.Extract_Val(Request.QueryString("WarningProgramID"))
    ProgDetail.ProgramID = MyCommon.Extract_Val(Request.QueryString("ProgramID"))
    ProgDetail.AdjustmentAmount = MyCommon.Extract_Val(Request.QueryString("Adjust"))
    TxDetail.TransOffer = MyCommon.Extract_Val(Request.QueryString("transOffer" & ProgDetail.ProgramID))
    
    MyCommon.QueryStr = "select RewardOptionID, CONVERT(VARCHAR(10), INC.StartDate, 1) as StartDate, " & _
                        "   CONVERT(VARCHAR(10), INC.EndDate, 1) as EndDate " & _
                        "from CPE_RewardOptions as RO with (NoLock)" & _
                        "inner join CPE_Incentives as INC on INC.IncentiveID = RO.IncentiveID " & _
                        "   and RO.TouchResponse=0 and RO.Deleted=0 and INC.Deleted=0  " & _
                        "where RO.IncentiveID=" & TxDetail.TransOffer
    dt = MyCommon.LRT_Select
    If dt.Rows.Count > 0 Then
      TxDetail.TransROID = MyCommon.NZ(dt.Rows(0).Item("RewardOptionID"), 0)
      OfferStart = MyCommon.NZ(dt.Rows(0).Item("StartDate"), "[" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "]")
      OfferEnd = MyCommon.NZ(dt.Rows(0).Item("EndDate"), "[" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "]")
    End If

    ' ensure that the transaction date is within the offers start and end dates
    MyCommon.QueryStr = "select count(*) as WithinOfferDates from CPE_Incentives " & _
                        "where IncentiveID=" & TxDetail.TransOffer & _
                        " and '" & TxDetail.TransDateStr & "' between StartDate and DateAdd(d,1,EndDate)"
    dt = MyCommon.LRT_Select
    If dt.Rows.Count > 0 Then
      If MyCommon.NZ(dt.Rows(0).Item("WithinOfferDates"), 0) = 0 Then
        InfoMessage = Copient.PhraseLib.Detokenize("CAM-point-adjust.TxnDateOutsideOffer", LanguageID, TxDetail.TransDateStr, TxDetail.TransOffer, OfferStart, OfferEnd)
      End If
    End If
    
    If InfoMessage = "" AndAlso TxDetail.LogixTransNum.Trim <> "" AndAlso IsValidAdjustment(TxDetail.TransOffer, CustomerPK, InfoMessage, WarningProgramID) Then
      ProgDetail.ExceededThreshold = (WarningProgramID > 0)
      If MyCommon.Extract_Val(Request.QueryString("SourceTypeID")) = 11 Then
        ProgDetail.SourceTypeID = 11 ' this is a reversal refund
      Else
        ProgDetail.SourceTypeID = 1 ' default to manual entry
      End If

      If ExecutePermitted Then
        InfoMessage = AdjustPoint(AdminUserID, CustomerExtID, TxDetail, ProgDetail, SessionID, TxDetail.TransOffer)
      Else
        InfoMessage = SubmitPoint(AdminUserID, TxDetail, ProgDetail)
      End If

      ' NB: The above line was originally AdjustPoint(AdminUserID, EarnedROID).
      ' I hard-coded the ROID to zero since that is what needs to be recorded in order to
      ' indicate that the points were awarded via a manual adjustment (and *not* via a
      ' particular offer or reward). --Huw
      If InfoMessage = "" Then
        Try
          ' to ensure the that the RedemptionCount for a manual adjustment is only 1, set 
          ' set the adjustment amount to 1, then reset it after the trans history is recorded
          TempAdjAmount = ProgDetail.AdjustmentAmount
          ProgDetail.AdjustmentAmount = 1

          CustDetail = New Copient.Customer
          CustDetail.AddCard(New Copient.Card(CustomerExtID, 2, CustomerPK))
          CustDetail.SetCustomerTypeID(2)
          MyCAM.AddToTransHistory(TxDetail, ProgDetail, CustDetail, cardpk, AdminUserID, MyCommon)
          'reset the program adjustment amount
          ProgDetail.AdjustmentAmount = TempAdjAmount
        Catch ex As Exception
          MyCommon.Write_Log("CAM.txt", "Ex: " & ex.ToString & " Offer adjustment failed to create transaction " & TxDetail.TransNumber & " for card " & CustomerExtID, True)
        End Try

        Response.Status = "301 Moved Permanently"
        Response.AddHeader("Location", "CAM-point-adjust-program.aspx?CustomerExtId=" & CustomerExtID & "&CustomerPK=" & CustomerPK & _
                           "&ProgramID=" & ProgramID & "&ProgramName=" & ProgramName & "&RefreshParent=true" & _
                           "&historyTo=" & Server.UrlEncode(Request.QueryString("historyTo")) & _
                           "&historyFrom=" & Server.UrlEncode(Request.QueryString("historyFrom")) & _
                           "&LogixTransNum=" & Request.QueryString("LogixTransNum"))
        GoTo done
      End If
      
    End If
    
  ElseIf (Request.QueryString("HistoryEnabled.x") <> "" OrElse Request.QueryString("HistoryDisabled.x") <> "") Then
    ' Write a cookie and then reload the page
    Response.Cookies("CAMHistoryEnabled").Expires = "10/08/2100"
    Response.Cookies("CAMHistoryEnabled").Value = IIf(Request.QueryString("HistoryEnabled.x") <> "", "1", "0")

    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "CAM-point-adjust-program.aspx?CustomerExtId=" & CustomerExtID & "&CustomerPK=" & CustomerPK & _
                       "&ProgramID=" & ProgramID & "&ProgramName=" & ProgramName & "&RefreshParent=true" & _
                       "&historyTo=" & Server.UrlEncode(Request.QueryString("historyTo")) & _
                       "&historyFrom=" & Server.UrlEncode(Request.QueryString("historyFrom")) & _
                       "&LogixTransNum=" & Request.QueryString("LogixTransNum"))
    GoTo done
  End If
  
  If (Request.QueryString("RefreshParent") = "true") Then RefreshParent = True
  
  Send_HeadBegin("term.point", "term.pointsadjustment", ProgramID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
%>

<script type="text/javascript" language="javascript">
    var datePickerDivID = "datepicker";
    var bSkipUnload = false;
    var linkToHH = false;
    
    <% Send_Calendar_Overrides(MyCommon) %>

    function isValidEntry(programId) {
        var retVal = true;
        var elem = document.getElementById("adjust" + programId);

        if (elem != null) {
            if (isNaN(elem.value)) {
                retVal = false;
                alert('<%Sendb(Copient.PhraseLib.Lookup("points-adjust.entryerror", LanguageID)) %>');
                if (!elem.disabled)  {
                  elem.focus();
                  elem.select();
                }
            }
        }                    

        return retVal;
    }
    
    function ChangeParentDocument() {
      var refreshElem = document.getElementById("RefreshParent");
      
      if (opener != null && !opener.closed && bSkipUnload != true) {
        if (refreshElem != null && refreshElem.value == "true") {
            if (opener.location.href.indexOf('CAM-customer-adjustments') > -1) {
              opener.location = 'CAM-customer-adjustments.aspx?CustPK=<%Sendb(CustomerPK)%>&adjWin=1';
            } else if (opener.location.href.indexOf('CAM-customer-transactions') > -1) {
              opener.location = 'CAM-customer-transactions.aspx?CustPK=<%Sendb(CustomerPK)%>&adjWin=1';
            }
        }
      }
    }
    
    function Refresher(ProgramID, CustomerPK) {
      var refreshElem = document.getElementById("RefreshParent");
      var RefreshPrt = "false";
      
      bSkipUnload = true;
      
      if (refreshElem != null && refreshElem.value == 'true') {
        RefreshPrt = "true";
      }
      location = ('CAM-point-adjust-program.aspx?ProgramID=' + ProgramID + '&CustomerPK=' + CustomerPK + '&RefreshParent=' + RefreshPrt+ '&historyTo=<%Sendb(Server.UrlEncode(Request.QueryString("historyTo")))%>&historyFrom=<%Sendb(Server.UrlEncode(Request.QueryString("historyFrom")))%>&LogixTransNum=<%Sendb(LogixTransNum.Trim)%>')
    }
    
    function HandleSwitchToHH() {
      var refreshElem = document.getElementById("RefreshParent");

      linkToHH = true;
      if (refreshElem != null) {
        refreshElem.value = "true";
      }      
    }
    
    function removeLogixTransNum() {
      var elemLogixTransNum = document.getElementById("LogixTransNum");
      
      if (elemLogixTransNum != null) {
        elemLogixTransNum.value = "";
        document.mainform.submit();
      }
    }
    function validateEntry(programID) {
      var elemAdj = document.getElementById('adjust' + programID);
      var elemProg = document.getElementById('ProgramID');
      var elemAdjust = document.getElementById('Adjust');
      var elemSave = document.getElementById('Save');
      var elemFocus = document.getElementById('FocusElem');
            
      if (elemAdj != null && isValidEntry(programID)) {
        if (elemSave != null) { elemSave.value = 'save'; }
        if (elemProg != null) { elemProg.value = programID; }
        if (elemAdjust != null) { elemAdjust.value = elemAdj.value; }
        if (elemFocus != null) { elemFocus.value = 'adjust' + programID; }
        
        bSkipUnload = true;  
        document.mainform.submit();       
      }
    }
    
    function validateRefund(programID) {
      var elemAdj = document.getElementById('adjust' + programID);
      var elemProg = document.getElementById('ProgramID');
      var elemAdjust = document.getElementById('Adjust');
      var elemSave = document.getElementById('Save');
      var elemFocus = document.getElementById('FocusElem');
      var elemSrcType = document.getElementById('SourceTypeID');
            
      if (elemAdj != null && isValidEntry(programID)) {
        if (parseInt(elemAdj.value) < 0) {
          if (elemSave != null) { elemSave.value = 'save'; }
          if (elemProg != null) { elemProg.value = programID; }
          if (elemAdjust != null) { elemAdjust.value = elemAdj.value; }
          if (elemFocus != null) { elemFocus.value = 'adjust' + programID; }
          if (elemSrcType != null) { elemSrcType.value = '11'; }
          bSkipUnload = true;  
          document.mainform.submit();       
        } else {
          alert('<%Sendb(Copient.PhraseLib.Lookup("CAM-point-adjust.ReversalRefundsNegative", LanguageID))%>');
        }
      }
    }
    function selectTrans(logixTransNum) {
      var elemLTN = document.getElementById('LogixTransNum');

      if (elemLTN != null) {
        elemLTN.value = logixTransNum;
        document.mainform.submit();
      }  
    }

</script>

<% 
  Send_Scripts(New String() {"datePicker.js"})
  Send_HeadEnd()
  Send_BodyBegin(2)
  
  If (Logix.UserRoles.AccessPointsBalances = False) Then
    Send_Denied(2, "perm.customers-ptbalaccess")
    GoTo done
  End If
  
%>
<form id="mainform" name="mainform" action="">
  <input type="hidden" id="ProgramID" name="ProgramID" value="<% Sendb(ProgramID)%>" />
  <input type="hidden" id="CustomerPK" name="CustomerPK" value="<% Sendb(CustomerPK)%>" />
  <input type="hidden" id="CustomerExtId" name="CustomerExtId" value="<%Sendb(CustomerExtID)%>" />
  <input type="hidden" id="RefreshParent" name="RefreshParent" value="<% Sendb(RefreshParent.ToString.ToLower) %>" />
  <input type="hidden" id="LogixTransNum" name="LogixTransNum" value="<% Sendb(LogixTransNum.Trim())%>" />
  <input type="hidden" id="Save" name="Save" value="" />
  <input type="hidden" id="Adjust" name="Adjust" value="" />
  <input type="hidden" id="WarningProgramID" name="WarningProgramID" value="<%Sendb(WarningProgramID)%>" />
  <input type="hidden" id="FocusElem" name="FocusElem" value="<%Sendb(FocusElem)%>" />
  <input type="hidden" id="SourceTypeID" name="SourceTypeID" value="1" />

  <div id="intro">
    <h1 id="title">
      <% Sendb(Copient.PhraseLib.Lookup("term.pointsprogram", LanguageID) & " #" & ProgramID & ": " & MyCommon.TruncateString(ProgramName, 40))%>
    </h1>
    <div id="controls">
      <%
        If (Logix.UserRoles.EditPointsBalances) AndAlso (ProgramIsExternal = False) AndAlso LogixTransNum.Trim <> "" Then
          'Send_Save()
        End If
      %>
    </div>
    <hr class="hidden" />
  </div>
  <div id="main">
    <% If (InfoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & InfoMessage & "</div>")%>
    <div id="column">
      <%
        If ProgramDesc <> "" Then
          Sendb("<p id=""description"">" & MyCommon.SplitNonSpacedString(ProgramDesc, 50) & "</p>")
        End If

      %>
      <div class="box" id="transaction">
        <h2>
          <span>
            <%Sendb(Copient.PhraseLib.Lookup("term.transaction", LanguageID))%>
          </span>
        </h2>
        <table summary="<%Sendb(Copient.PhraseLib.Lookup("term.transaction", LanguageID))%>" style="width:95%;table-layout:auto;">
          <tr>
            <td><%Sendb(Copient.PhraseLib.Lookup("term.customer", LanguageID))%>:</td>
            <td><%Sendb(IIf(CustFirstName <> "", CustFirstName & " ", "") & IIf(CustMiddleName <> "", Left(CustMiddleName, 1) & ". ", "") & IIf(CustLastName <> "", CustLastName, ""))%></td>
            <td style="width:20px;"></td>
            <td><%Sendb(Copient.PhraseLib.Lookup("term.cardnumber", LanguageID))%>:</td>
            <td><%Sendb(CustomerExtID)%></td>
          </tr>
          <% If LogixTransNum.Trim() <> "" Then%>
          <tr>
            <td><%Sendb(Copient.PhraseLib.Lookup("term.transactionnumber", LanguageID))%>:</td>
            <td align="left"><div style="word-break: break-all;"><%Sendb(TransNumber)%></div></td>
            <td style="width:20px;"></td>
            <td><%Sendb(Copient.PhraseLib.Lookup("term.store", LanguageID))%>:</td>
            <td><%Send(TransStore)%></td>
            <td><input type="button" name="btnTransChange" id="btnTransChange" class="short" value="<%Sendb(Copient.PhraseLib.Lookup("term.change", LanguageID)) %>"  onclick="removeLogixTransNum();"/></td>            
          </tr>
          <tr>
            <td><%Sendb(Copient.PhraseLib.Lookup("term.date", LanguageID))%>:</td>
            <td><%Send(TransDate.ToString("dd MMM yyyy, HH:mm:ss"))%></td>
            <td style="width:20px;"></td>
            <td><%Sendb(Copient.PhraseLib.Lookup("term.terminal", LanguageID))%>:</td>
            <td><%Send(TransTerminal)%></td>
            <td></td>
          </tr>
          <% Else %>
          <tr>
            <td><%Sendb(Copient.PhraseLib.Lookup("term.transactionnumber", LanguageID))%>:</td>
            <td><input type="text" name="TransNum" id="TransNum" class="medium" value="<%Sendb(TransNumber)%>" maxlength="128" tabindex="1"/></td>
            <td style="width:20px;"></td>
            <td><%Sendb(Copient.PhraseLib.Lookup("term.store", LanguageID))%>:</td>
            <td><input type="text" name="TransStore" id="TransStore" class="short" value="<%Sendb(TransStore)%>"  maxlength="20" tabindex="3" /></td>
            <td><input type="submit" name="btnTransSearch" id="btnTransSearch" class="short" value="<%Sendb(Copient.PhraseLib.Lookup("term.search", LanguageID)) %>" /></td>            
          </tr>
          <tr>
            <td><%Sendb(Copient.PhraseLib.Lookup("term.date", LanguageID))%>:</td>
            <td><input type="text" name="TransDate" id="TransDate" class="medium" value="<%Sendb(iif(TransDate = "1/1/1980", "", TransDate.ToString()))%>"  tabindex="2" /></td>
            <td style="width:20px;"></td>
            <td><%Sendb(Copient.PhraseLib.Lookup("term.terminal", LanguageID))%>:</td>
            <td><input type="text" name="TransTerminal" id="TransTerminal" class="short" maxlength="4" value="<%Sendb(TransTerminal)%>" tabindex="4" /></td>
            <td>
             <% If Logix.UserRoles.CreateTransaction Then%>
              <input type="submit" name="btnTransCreate" id="btnTransCreate" class="short" value="<%Sendb(Copient.PhraseLib.Lookup("term.create", LanguageID)) %>" />
             <% End If %>
            </td>            
          </tr>
          <% End If %>
        </table>
        <hr class="hidden" />
      </div>
      
        <% If ShowResults Then%>
      <div class="box" id="transResults">
        <h2><span><% Sendb(Copient.PhraseLib.Lookup("customer-inquiry.transaction-results", LanguageID))%></span></h2>
        <%
          If LogixTransNums IsNot Nothing AndAlso LogixTransNums.Length > 0 Then
            ' add quotes around values for use in the queries
            ReDim LogixTransList(LogixTransNums.GetUpperBound(0))
            For i = 0 To LogixTransNums.GetUpperBound(0)
              LogixTransList(i) = "'" & LogixTransNums(i).Trim & "'"
            Next
            
            ' get the transaction information and then loop through it and write the rows
            MyCommon.QueryStr = "select top 20 LogixTransNum,CustomerPrimaryExtID, ExtLocationCode, TransDate, TerminalNum, POSTransNum from TransHistory with (NoLock) " & _
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
                Send("  <td><input type=""button"" name=""select"" value=""..."" title=""" & Copient.PhraseLib.Lookup("CAM-point-adjust.ClickToSelect", LanguageID) & """ onclick=""selectTrans('" & MyCommon.NZ(row.Item("LogixTransNum"), "") & "');"" /></td>")
                Send("  <td>" & MyCommon.NZ(row.Item("CustomerPrimaryExtID"), "") & "</td>")
                Send("  <td style=""word-break:break-all"">" & MyCommon.NZ(row.Item("POSTransNum"), "") & "</td>")
                Send("  <td>" & MyCommon.NZ(row.Item("TransDate"), "") & "</td>")
                Send("  <td>" & MyCommon.NZ(row.Item("ExtLocationCode"), "") & "</td>")
                Send("  <td>" & MyCommon.NZ(row.Item("TerminalNum"), "") & "</td>")
                Send("</tr>")
              Next
              Send("</table>")
            Else
              Send("<center>" & Copient.PhraseLib.Lookup("customer-manual.NoTransactionsFound", LanguageID) & "</center>")
            End If
          End If
          
        %>
      </div>
      <% End If %>
      
      <div class="box" id="history">
        <h2>
          <span><%Sendb(Copient.PhraseLib.Lookup("term.history", LanguageID))%></span>
        </h2>
        <% 
          If (ProgramID > 0) Then
            ShowProcessingOrHistory(CustomerPK, ProgramID, True, LogixTransNum)
          Else
            Send(Copient.PhraseLib.Lookup("point-adjust.nopointprograms", LanguageID) & "<br />")
          End If
        %>
        <hr class="hidden" />
      </div>
      
      <div class="box" id="ptAdj"<% If (Not Logix.UserRoles.EditPointsBalances) OrElse (HHPK > 0) Then Sendb(" style=""display:none;""")%>>
        <h2>
          <span>
            <%Sendb(Copient.PhraseLib.Lookup("term.pointsadjustment", LanguageID))%>
          </span>
        </h2>
        <% 
          If (ProgramID > 0) Then
            ShowPoints(CustomerPK, ProgramID, TxDetail, LogixTransNum.Trim, Logix, FocusElem)
          Else
            Send(Copient.PhraseLib.Lookup("point-adjust.nopointprograms", LanguageID) & "<br />")
          End If
        %>
        <hr class="hidden" />
      </div>
      
      <div class="box" id="processing">
        <h2><span><%Sendb(Copient.PhraseLib.Lookup("term.processing", LanguageID))%></span></h2>
        <span id="refreshlink" style="float:right;font-size:9px;position:relative;top:-22px;"><a href="javascript:Refresher(<%Sendb(ProgramID)%>,<%Sendb(CustomerPK)%>)"><%Sendb(Copient.PhraseLib.Lookup("term.refresh", LanguageID))%></a></span>
        <% 
          If (ProgramID > 0) Then
            ShowProcessingOrHistory(CustomerPK, ProgramID, False, LogixTransNum)
          Else
            Send(Copient.PhraseLib.Lookup("point-adjust.nopointprograms", LanguageID) & "<br />")
          End If
        %>
        <hr class="hidden" />
      </div>
      
      <div class="box" id="pending">
        <h2><span><%Sendb(Copient.PhraseLib.Lookup("term.pending", LanguageID))%></span></h2>
        <% 
          ShowPending(CustomerPK, TxDetail)
        %>
        <hr class="hidden" />
      </div>      
  
      </div>
  </div>
</form>

<script runat="server">
  Dim MyCommon As New Copient.CommonInc
  
  Sub ShowPoints(ByVal CustomerPK As Long, ByVal ProgramID As Long, ByVal TxDetail As Copient.CAM.TransactionDetail, _
                 ByVal LogixTransNum As String, ByVal Logix As Copient.LogixInc, ByRef FocusElem As String)
    Dim UpdateAccum As Boolean = False
    Dim dt As DataTable
    Dim row As DataRow
    Dim dtPrograms As DataTable
    Dim rowProgram As DataRow = Nothing
    Dim Points As Integer = 0
    Dim PendingAdj As Integer = 0
    Dim InProcessAdj As Integer = 0
    Dim i As Integer = 1
    Dim PromoVarID As Integer = 0
    Dim External As Boolean = False
    Dim MyPoints As New Copient.Points
    Dim TempDate As Date
    Dim OfferID As Long
    
    MyCommon.Open_LogixRT()
    MyCommon.Open_LogixXS()
    
    MyCommon.QueryStr = "select ProgramID, ProgramName, PromoVarID, ExternalProgram from PointsPrograms with (NoLock) " & _
                        "where ProgramID=" & ProgramID & ";"
    dtPrograms = MyCommon.LRT_Select
    If dtPrograms.Rows.Count > 0 Then
      External = MyCommon.NZ(dtPrograms.Rows(0).Item("ExternalProgram"), False)
    End If
    
    Send("<br class=""half"" />")
    Send(Copient.PhraseLib.Lookup("point-adjust.programnote", LanguageID) & " " & ProgramID)
    Send("<br /><br class=""half"" />")
    Send("        <table summary=""" & Copient.PhraseLib.Lookup("term.programs", LanguageID) & """>")
    Send("         <thead>")
    Send("          <tr>")
    Send("            <th class=""th-name"" scope=""col"">" & Copient.PhraseLib.Lookup("term.program", LanguageID) & "</th>")
    Send("            <th class=""th-id"" scope=""col"">" & Copient.PhraseLib.Lookup("term.id", LanguageID) & "</th>")
    Send("            <th class=""th-pending"" scope=""col"" style=""text-align:center"">" & Copient.PhraseLib.Lookup("term.processing", LanguageID) & "</th>")
    Send("            <th class=""th-pending"" scope=""col"" style=""text-align:center"">" & Copient.PhraseLib.Lookup("term.pending", LanguageID) & "</th>")
    Send("            <th class=""th-quantity"" scope=""col"" style=""text-align:center"">" & Copient.PhraseLib.Lookup("term.quantity", LanguageID) & "</th>")
    Send("            <th style=""width:90px;"" scope=""col"">" & Copient.PhraseLib.Lookup("term.offerid", LanguageID) & "</th>")
    Send("            <th style=""width:80px;"" scope=""col"">" & Copient.PhraseLib.Lookup("term.adjustment", LanguageID) & "</th>")
    Send("            <th scope=""col""></th>")
    Send("          </tr>")
    Send("         </thead>")
    Send("         <tbody>")
    
    For Each rowProgram In dtPrograms.Rows
      PromoVarID = MyCommon.NZ(rowProgram.Item("PromoVarID"), 0)
      ProgramID = MyCommon.NZ(rowProgram.Item("ProgramID"), 0)
      Points = MyPoints.GetBalance(CustomerPK, ProgramID, PromoVarID)
      InProcessAdj = MyPoints.GetInProcessAdjustment(CustomerPK, ProgramID)
      If Not Date.TryParse(TxDetail.TransDateStr, TempDate) OrElse TempDate < Now.AddYears(-100) Then TempDate = New Date(1980, 1, 1)
      PendingAdj = MyPoints.GetPendingForTransaction(CustomerPK, ProgramID, TxDetail.TransNumber, TempDate, _
                                                    TxDetail.TransStore, TxDetail.TransTerminal)
      
      Send("          <tr>")
      Send("            <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(rowProgram.Item("ProgramName"), "nbsp;"), 25) & "</td>")
      Send("            <td>" & ProgramID & "</td>")
      Send("            <td style=""text-align:right"">" & InProcessAdj & "</td>")
      Send("            <td style=""text-align:right"">" & PendingAdj & "</td>")
      Send("            <td style=""text-align:right"">" & Points & "</td>")
      If (External OrElse LogixTransNum.Trim = "") Then
        Send("            <td colspan=""2"" style=""text-align:center;"">" & Copient.PhraseLib.Lookup("term.not-available", LanguageID) & "</td>")
        If External Then
          Send("            <td>(" & Copient.PhraseLib.Lookup("term.external", LanguageID) & " " & Copient.PhraseLib.Lookup("term.program", LanguageID) & ")</td>")
        Else
          Send("            <td></td>")
        End If
      Else
        MyCommon.QueryStr = "select distinct RO.IncentiveID from CPE_RewardOptions as RO with (NoLock) " & _
                            "inner join CPE_Deliverables as DEL with (NoLock) on DEL.RewardOptionID = RO.RewardOptionID " & _
                            "inner join CPE_DeliverablePoints as DP with (NoLock) on DP.PKID = DEL.OutputID " & _
                            "where DEL.DeliverableTypeID=8 and DEL.Deleted=0 and RO.Deleted=0 and DP.Deleted=0 and DP.ProgramID=" & ProgramID
        dt = MyCommon.LRT_Select
        If dt.Rows.Count > 0 Then
          Send("            <td><select class=""short"" id=""transOffer" & ProgramID & """ name=""transOffer" & ProgramID & """>")
          Send("                  <option value=""""></option>")
          For Each row In dt.Rows
            OfferID = MyCommon.NZ(row.Item("IncentiveID"), 0)
            If OfferID > 0 AndAlso OfferID = MyCommon.Extract_Val(Request.QueryString("transOffer" & ProgramID)) Then
              Send("                  <option value=""" & OfferID & """ selected=""selected"">" & OfferID & "</option>")
            ElseIf OfferID > 0 Then
              Send("                  <option value=""" & OfferID & """>" & OfferID & "</option>")
            End If
          Next
          Send("                 </select></td>")
        Else
          Send("            <td><input type=""text"" class=""shorter"" id=""transOffer" & ProgramID & """ name=""transOffer" & ProgramID & """ style=""text-align:right;"" maxlength=""7"" value=""" & Request.QueryString("transOffer" & ProgramID) & """ /></td>")
        End If
        
        Send("            <td><input type=""text"" class=""shorter"" id=""adjust" & ProgramID & """ name=""adjust" & ProgramID & """ style=""text-align:right;"" maxlength=""7"" value=""" & Request.QueryString("adjust" & ProgramID) & """ /></td>")
        If FocusElem = "" Then FocusElem = "transOffer" & ProgramID
        If Logix.UserRoles.EditPointsBalances Then
          Send("            <td>")
          Sendb("  <input type=""button"" class=""adjust"" id=""ptsAdj" & ProgramID & """ name=""ptsAdj"" value=""E"" title=""" & Copient.PhraseLib.Lookup("term.adjust", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.points", LanguageID), VbStrConv.Lowercase) & """ ")
          Send("onClick=""javascript:validateEntry('" & ProgramID & "');"" />")
          Sendb("  <input type=""button"" class=""adjust"" id=""revRefund" & ProgramID & """ name=""revRefund"" value=""R"" title=""" & Copient.PhraseLib.Lookup("term.reversal-refund", LanguageID) & """ ")
          Send("onClick=""javascript:validateRefund('" & ProgramID & "');"" />")
          Send("            </td>")
        Else
          Send("          <td></td>")
        End If
      End If
      Send("          </tr>")
      i += 1
    Next
    Send("         </tbody>")
    Send("        </table>")
    MyCommon.Close_LogixRT()
    MyCommon.Close_LogixXS()
  End Sub
  
  Sub ShowProcessingOrHistory(ByVal CustomerPK As Long, ByVal ProgramID As Integer, _
                              ByVal ShowHistory As Boolean, ByVal LogixTransNum As String)
    Dim dt As DataTable
    Dim dt2 As DataTable
    Dim dt3 As DataTable
    Dim row As DataRow
    Dim sQueryBuilder As New StringBuilder()
    Dim ProgramName As String = ""
    Dim OfferNumber As Integer
    Dim LocationID As Integer
    Dim ExtLocationCode As String = ""
    Dim i As Integer = 0
    Dim StartDateStr As String = ""
    Dim EndDateStr As String = ""
    
    Dim PageNum As Integer = 0
    Dim MorePages As Boolean
    Dim linesPerPage As Integer = 15
    Dim Shaded As String = "shaded"
    Dim Cookie As HttpCookie = Nothing
    Dim HistoryEnabled As Boolean = True
    Dim AltText As String = ""
    
    PageNum = Request.QueryString("pagenum")
    If PageNum < 0 Then PageNum = 0
    MorePages = False
    
    Dim SortText As String = "ActivityDate"
    Dim SortDirection As String = ""
    
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
    
    MyCommon.Open_LogixRT()
    MyCommon.Open_LogixXS()
    MyCommon.Open_LogixEX()
    
    If (ShowHistory) Then
      sQueryBuilder.Append("select ProgramID, AdjAmount, EarnedUnderROID, EarnedUnderCMOfferID, LastUpdate, LastServerID, LocationID, SourceTypeID ")
      sQueryBuilder.Append("from PointsHistoryView with (NoLock) ")
      sQueryBuilder.Append("where ProgramID = " & ProgramID & " and CustomerPK = " & CustomerPK & " and LogixTransNum = '" & LogixTransNum & "' ")
      sQueryBuilder.Append(" order by LastUpdate desc;")
    Else
      sQueryBuilder.Append("select Convert(int, Col1) as ProgramID, Convert(int,Col3) as AdjAmount, Convert(int,Col4) as EarnedUnderROID, 0 as EarnedUnderCMOfferID, getDate() as LastUpdate ")
      sQueryBuilder.Append("from CPE_UploadTemp_PA with (NoLock) ")
      sQueryBuilder.Append("where Col1 = '" & ProgramID & "' ")
      sQueryBuilder.Append(" and Col2 = '" & CustomerPK & "';")
    End If
    
    MyCommon.QueryStr = sQueryBuilder.ToString
    dt = MyCommon.LXS_Select
    If (dt.Rows.Count > 0) Then
      Send("<table summary=""" & Copient.PhraseLib.Lookup("term.history", LanguageID) & """>")
      Send("  <tr>")
      Send("    <th class=""th-name"" scope=""col"">" & Copient.PhraseLib.Lookup("term.pointsprogram", LanguageID) & "</th>")
      Send("    <th class=""th-id"" scope=""col"">" & Copient.PhraseLib.Lookup("term.id", LanguageID) & "</th>")
      Send("    <th class=""th-adjustment"" scope=""col"">" & Copient.PhraseLib.Lookup("term.adjustment", LanguageID) & "</th>")
      Send("    <th class=""th-author"" scope=""col"">" & Copient.PhraseLib.Lookup("term.from", LanguageID) & "</th>")
      Send("    <th class=""th-datetime"" scope=""col"">" & Copient.PhraseLib.Lookup("term.lastupdated", LanguageID) & "</th>")
      Send("  </tr>")
      For Each row In dt.Rows
        ProgramID = MyCommon.NZ(row.Item("ProgramID"), 0)
        If (ShowHistory) Then
          LocationID = MyCommon.NZ(row.Item("LocationID"), 0)
        End If
        MyCommon.QueryStr = "select ExtLocationCode from Locations with (NoLock) where LocationID=" & LocationID
        dt2 = MyCommon.LRT_Select
        If (dt2.Rows.Count > 0) Then
          ExtLocationCode = MyCommon.NZ(dt2.Rows(0).Item("ExtLocationCode"), "")
        End If
        MyCommon.QueryStr = "select ProgramName from PointsPrograms with (NoLock) where ProgramID=" & ProgramID
        dt2 = MyCommon.LRT_Select
        If (dt2.Rows.Count > 0) Then
          ProgramName = MyCommon.NZ(dt2.Rows(0).Item("ProgramName"), "")
        End If
        If MyCommon.NZ(row.Item("EarnedUnderROID"), 0) = 0 AndAlso MyCommon.NZ(row.Item("EarnedUnderCMOfferID"), 0) = 0 Then
          OfferNumber = 0
        Else
          If MyCommon.NZ(row.Item("EarnedUnderROID"), 0) <> 0 Then
            MyCommon.QueryStr = "select IncentiveID as OfferID from CPE_RewardOptions as RO where RewardOptionID=" & MyCommon.NZ(row.Item("EarnedUnderROID"), 0)
            dt3 = MyCommon.LRT_Select
            If (dt3.Rows.Count > 0) Then
              OfferNumber = MyCommon.NZ(dt3.Rows(0).Item("OfferID"), 0)
            End If
          End If
          If MyCommon.NZ(row.Item("EarnedUnderCMOfferID"), 0) <> 0 Then
            OfferNumber = MyCommon.NZ(row.Item("EarnedUnderCMOfferID"), 0)
          End If
        End If
        Send("  <tr>")
        Send("    <td>" & MyCommon.SplitNonSpacedString(ProgramName, 25) & "</td>")
        Send("    <td>" & ProgramID & "</td>")
        Send("    <td>" & MyCommon.NZ(row.Item("AdjAmount"), "&nbsp;") & "</td>")
        If MyCommon.NZ(row.Item("EarnedUnderROID"), "&nbsp;") = 0 Then
          Send("    <td>" & Copient.PhraseLib.Lookup("term.logix-manual-entry", LanguageID) & "</td>")
        Else
          Sendb("    <td>")
          If OfferNumber > 0 Then
            Sendb(Copient.PhraseLib.Lookup("term.offer", LanguageID) & " " & OfferNumber & " ")
          End If
          If LocationID > 0 Then
            If OfferNumber > 0 Then
              Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.at", LanguageID), VbStrConv.Lowercase) & " ")
            End If
            Sendb(ExtLocationCode)
          ElseIf LocationID = -9 Then
            Sendb(Copient.PhraseLib.Lookup("term.logix-manual-entry", LanguageID))
          ElseIf LocationID = -10 Then
            MyCommon.QueryStr = "select Description from SourceTypes with (NoLock) where SourceTypeID=" & MyCommon.NZ(row.Item("SourceTypeID"), 0)
            dt2 = MyCommon.LEX_Select
            If (dt2.Rows.Count > 0) Then
              Sendb(MyCommon.NZ(dt2.Rows(0).Item("Description"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)))
            Else
              Sendb(Copient.PhraseLib.Lookup("term.unknown", LanguageID))
            End If
          End If
          Send("</td>")
        End If
        If (Not IsDBNull(row.Item("LastUpdate"))) Then
          Send("    <td>" & Format(row.Item("LastUpdate"), "dd MMM yyyy, HH:mm:ss") & "</td>")
        Else
          Send("    <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
        End If
        Send("  </tr>")
      Next
      Send("</table>")
    Else
      If (ShowHistory) Then
        Send("<i>" & Copient.PhraseLib.Lookup("cam-point-adjust.nohistory", LanguageID) & "</i>")
      Else
        Send("<i>" & Copient.PhraseLib.Lookup("point-adjust.noprocessing", LanguageID) & "</i>")
      End If
    End If
    MyCommon.Close_LogixRT()
    MyCommon.Close_LogixXS()
    MyCommon.Close_LogixEX()
  End Sub
  
  Sub ShowPending(ByVal CustomerPK As Long, ByVal TxDetail As Copient.CAM.TransactionDetail)
    Dim dt As DataTable
    Dim dt2 As DataTable
    Dim row As DataRow
    Dim ProgramID As Long
    Dim ProgramName As String = ""
    Dim Adjustment As Integer = 0
    Dim CreatorID As Integer = 0
    Dim CreatorName As String = ""
    Dim CreateDate As Date
    Dim TransDate As Date
    
    Try
      If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
      If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
      
      If Not Date.TryParse(TxDetail.TransDateStr, TransDate) Then TransDate = New Date(1980, 1, 1)
      
      MyCommon.QueryStr = "select PKID, OfferID, ProgramID, AdjAmount, CreateDate, CreatedBy from PointsAdj_Pending with (NoLock) " & _
                          "where TransNum = '" & TxDetail.TransNumber & "' and ExtLocationCode='" & TxDetail.TransStore & "' " & _
                          "  and TerminalNum = '" & TxDetail.TransTerminal & "' and TransDate = '" & TransDate.ToString & "' " & _
                          "  and CustomerPK='" & CustomerPK & "' and AdjAmount <> 0 " & _
                          "order by CreateDate, CustomerPK;"
      dt = MyCommon.LXS_Select
      If (dt.Rows.Count > 0) Then
        Send("<table summary=""" & Copient.PhraseLib.Lookup("term.pending", LanguageID) & """>")
        Send("  <tr>")
        Send("    <th class=""th-name"" scope=""col"">" & Copient.PhraseLib.Lookup("term.pointsprogram", LanguageID) & "</th>")
        Send("    <th class=""th-id"" scope=""col"">" & Copient.PhraseLib.Lookup("term.id", LanguageID) & "</th>")
        Send("    <th class=""th-adjustment"" scope=""col"">" & Copient.PhraseLib.Lookup("term.adjustment", LanguageID) & "</th>")
        Send("    <th class=""th-author"" scope=""col"">" & Copient.PhraseLib.Lookup("term.from", LanguageID) & "</th>")
        Send("    <th class=""th-datetime"" scope=""col"">" & Copient.PhraseLib.Lookup("term.created", LanguageID) & "</th>")
        Send("  </tr>")
        For Each row In dt.Rows
          ProgramID = MyCommon.NZ(row.Item("ProgramID"), 0)
          Adjustment = MyCommon.NZ(row.Item("AdjAmount"), 0)
          CreateDate = MyCommon.NZ(row.Item("CreateDate"), New Date(1980, 1, 1))
          CreatorID = MyCommon.NZ(row.Item("CreatedBy"), 1)
          
          MyCommon.QueryStr = "select ProgramName from PointsPrograms with (NoLock) where ProgramID=" & ProgramID
          dt2 = MyCommon.LRT_Select
          If (dt2.Rows.Count > 0) Then
            ProgramName = MyCommon.NZ(dt2.Rows(0).Item("ProgramName"), "")
          End If
          
          MyCommon.QueryStr = "select FirstName + ' ' + LastName as FullName from AdminUsers with (NoLock) where AdminUserID=" & CreatorID
          dt2 = MyCommon.LRT_Select
          If (dt2.Rows.Count > 0) Then
            CreatorName = MyCommon.NZ(dt2.Rows(0).Item("FullName"), "")
          End If
          
          Send("<tr>")
          Send("  <td>" & ProgramName & "</td>")
          Send("  <td>" & ProgramID & "</td>")
          Send("  <td>" & Adjustment & "</td>")
          Send("  <td>" & CreatorName & "</td>")
          Send("  <td>" & CreateDate.ToString("g") & "</td>")
          Send("</tr>")
        Next
        Send("</table>")
      Else
        Send("<i>" & Copient.PhraseLib.Lookup("cam-point-adjust.nopending", LanguageID) & "</i>")
      End If
    Catch ex As Exception
      Send("<i>" & Copient.PhraseLib.Lookup("cam-point-adjust.pendingerror", LanguageID) & "</i>")
    Finally
      If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
      If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
    End Try
    
  End Sub
  
  Function IsValidAdjustment(ByVal OfferID As Long, ByVal CustomerPK As Long, ByRef infoMessage As String, ByRef WarningProgramID As Long) As Boolean
    Dim ValidAdj As Boolean = False
    Dim MyCAM As New Copient.CAM
    Dim MyPoints As New Copient.Points
    Dim PointsBal, WarningLimit, PendingAdj, tempAdj As Long
    Dim ProgramID, AdjustAmt As Long
    Dim dt As DataTable
    Dim row As DataRow
    
    
    ProgramID = MyCommon.Extract_Val(Request.QueryString("programId"))
    AdjustAmt = MyCommon.Extract_Val(Request.QueryString("adjust" & ProgramID))
    PointsBal = MyPoints.GetBalance(CustomerPK, ProgramID)
    WarningLimit = MyCAM.GetMaxAdjustment(OfferID, ProgramID)
    
    MyCommon.Open_LogixXS()
    MyCommon.QueryStr = "select sum(Convert(int, IsNull(Col3,0))) as Pending from CPE_UploadTemp_PA with (NoLock) " & _
                        "where Col1='" & ProgramID & "' and Col2='" & CustomerPK & "';"
    dt = MyCommon.LXS_Select
    If (dt.Rows.Count > 0) Then
      PendingAdj = 0
      For Each row In dt.Rows
        tempAdj = MyCommon.NZ(row.Item("Pending"), 0)
        PendingAdj = PendingAdj + tempAdj
      Next
    End If
    MyCommon.Close_LogixXS()
        
    If OfferID <= 0 Then
      infoMessage &= Copient.PhraseLib.Lookup("CAM-point-adjust.SelectOffer", LanguageID)
    ElseIf ProgramID <= 0 Then
      infoMessage &= Copient.PhraseLib.Detokenize("CAM-point-adjust.InvalidProgram", LanguageID, ProgramID)
    ElseIf AdjustAmt = 0 Then
      infoMessage &= Copient.PhraseLib.Lookup("lmg.ErrorZeroAdjustment", LanguageID)
    ElseIf ((PointsBal + PendingAdj) + AdjustAmt < 0 AndAlso AdjustAmt < 0) Then
      infoMessage &= Copient.PhraseLib.Detokenize("points-adjust.NegativeBalWarning", LanguageID, (-(PointsBal + PendingAdj)))
    ElseIf (Math.Abs(AdjustAmt) > WarningLimit) Then
      ' if a warning has already been issued then allow the adjustment
      If WarningProgramID = 0 Then
        infoMessage &= Copient.PhraseLib.Lookup("points-adjust.OverMaxLimit", LanguageID)
        WarningProgramID = ProgramID
      Else
        ValidAdj = True
      End If
    Else
      ValidAdj = True
    End If
    
    Return ValidAdj
  End Function

  Function AdjustPoint(ByVal AdminUserID As String, ByVal CustomerExtID As String, _
                  ByVal TxDetail As Copient.CAM.TransactionDetail, ByVal ProgDetail As Copient.CAM.ProgramDetail, _
                  ByVal SessionID As String, ByVal SelectedOfferID As Long) As String

    Dim MyCam As New Copient.CAM
    Dim MyPoints As New Copient.Points
    Dim ProgramID As Long
    Dim PointsAdj As Long
    Dim CustomerPK As Long
    Dim LogText As String = ""
    Dim RetMsg As String = ""
    Dim bNeedsRollback As Boolean = False
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
      PointsAdj = ProgDetail.AdjustmentAmount
      CustomerPK = MyCommon.Extract_Val(Request.QueryString("CustomerPK"))
      
      MyCommon.QueryStr = "begin transaction"
      MyCommon.LXS_Execute()
      MyCommon.LEX_Execute()
      
      If (CustomerPK > 0 AndAlso ProgramID > 0) Then
        PreAdjustBal = New Decimal(MyPoints.GetBalance(CustomerPK, ProgramID))
        
        MyCommon.QueryStr = "dbo.pa_CPE_TU_InsertData_PA"
        MyCommon.Open_LXSsp()
        MyCommon.LXSsp.Parameters.Add("@TableNum", SqlDbType.VarChar, 4).Value = "4"
        MyCommon.LXSsp.Parameters.Add("@Operation", SqlDbType.VarChar, 2).Value = "1"
        MyCommon.LXSsp.Parameters.Add("@Col1", SqlDbType.VarChar, 255).Value = ProgramID.ToString
        MyCommon.LXSsp.Parameters.Add("@Col2", SqlDbType.VarChar, 255).Value = CustomerPK.ToString
        MyCommon.LXSsp.Parameters.Add("@Col3", SqlDbType.VarChar, 255).Value = PointsAdj.ToString
        MyCommon.LXSsp.Parameters.Add("@Col4", SqlDbType.VarChar, 255).Value = MyCommon.Extract_Val(TxDetail.TransROID)
        MyCommon.LXSsp.Parameters.Add("@Col5", SqlDbType.VarChar, 255).Value = 2 ' CustomerTypeID
        MyCommon.LXSsp.Parameters.Add("@Col6", SqlDbType.VarChar, 255).Value = TxDetail.LogixTransNum
        MyCommon.LXSsp.Parameters.Add("@Col7", SqlDbType.VarChar, 255).Value = ProgDetail.SourceTypeID
        MyCommon.LXSsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = -9
        MyCommon.LXSsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = IIf(ProgDetail.SourceTypeID = 11, -10, -9)
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
            LogText &= " " & StrConv(Copient.PhraseLib.Lookup("term.by", LanguageID), VbStrConv.Lowercase) & " " & MyCommon.Extract_Val(PointsAdj)
            
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
			End If

            ' log the adjustment and any associated offers
            Fields.ActivityTypeID = 25
            Fields.ActivitySubTypeID = 12
            Fields.LinkID = CustomerPK
            Fields.AdminUserID = AdminUserID
            Fields.Description = LogText
            Fields.LinkID2 = MyCommon.Extract_Val(ProgramID)
            Fields.ActivityValue = PointsAdj.ToString
            Fields.AssociatedLinks = AssocLinks
            Fields.SessionID = SessionID
            Fields.PreAdjustBalance = PreAdjustBal
            Fields.Adjustment = New Decimal(PointsAdj)
            Fields.PostAdjustBalance = Decimal.Add(Fields.PreAdjustBalance, Fields.Adjustment)
            MyCommon.Activity_Log3(Fields)
          Catch ex As Exception
            MyCommon.Write_Log("CAM.txt", "Failed to log CAM points program adjustment for the following reason: " & ex.ToString, True)
          End Try
        End If
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
  
  Function SubmitPoint(ByVal AdminUserID As String, ByVal TxDetail As Copient.CAM.TransactionDetail, ByVal ProgDetail As Copient.CAM.ProgramDetail) As String
    
    Dim RetMsg As String = ""
    Dim ProgramID As Long
    Dim CustomerPK As Long
    Dim AdjustAmt As Long

    ProgramID = ProgDetail.ProgramID
    AdjustAmt = ProgDetail.AdjustmentAmount
    CustomerPK = MyCommon.Extract_Val(Request.QueryString("CustomerPK"))

    Try
      If (MyCommon.LXSadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixXS()
      If (MyCommon.LRTadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixRT()
      
      If TxDetail.TransDateStr = "" Then TxDetail.TransDateStr = "1980-01-01"

      ' write this points adjustment to the pending adjustment table
      MyCommon.QueryStr = "insert into PointsAdj_Pending with (RowLock) (LogixTransNum, TransNum, TransDate, ExtLocationCode, TerminalNum, " & _
                          "                    CustomerPK, ProgramID, OfferID, AdjAmount, CreateDate, CreatedBy) " & _
                          " values ('" & TxDetail.LogixTransNum & "', '" & TxDetail.TransNumber & "', " & _
                          "   '" & TxDetail.TransDateStr & "', '" & TxDetail.TransStore & "', '" & TxDetail.TransTerminal & "', " & _
                          "    " & CustomerPK & ", " & ProgramID & ", " & TxDetail.TransOffer & ", " & _
                          "    " & AdjustAmt & ", getdate(), " & AdminUserID & ");"
      MyCommon.LXS_Execute()

      If MyCommon.RowsAffected <= 0 Then
        RetMsg = Copient.PhraseLib.Lookup("customer-manual.ErrorEncounteredAdjusting", LanguageID)
      End If
    Catch ex As Exception
      RetMsg = ex.ToString
    Finally
      MyCommon.Close_LogixRT()
      MyCommon.Close_LogixXS()
 
    End Try
    
    Return RetMsg
  End Function
  
</script>


<%
done:
  If (LogixTransNum.Trim <> "") Then
    If FocusElem <> "" Then
      Send_BodyEnd("mainform", FocusElem)
    Else
      Send_BodyEnd()
    End If
  Else
    Send_BodyEnd("mainform", "TransNum")
  End If
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  MyCommon.Close_LogixWH()
  MyCommon = Nothing
  Logix = Nothing
%>
