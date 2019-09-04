<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: CM-cashier-search.aspx 
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
  ' * Version : 5.10b1.0 
  ' *
  ' *****************************************************************************
  
  Dim CopientFileName As String = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
  Dim CopientFileVersion As String = "7.3.1.138972"
  Dim CopientProject As String = "Copient Logix"
  Dim CopientNotes As String = ""
  
  Dim AdminUserID As Long
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim dt As DataTable
  Dim row As DataRow
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim TokenString As String = ""
  Dim Cashier As String = ""
  Dim Store As String = ""
  Dim CardId As String = ""
  Dim CardType As String = ""
  Dim Desc As String = ""
  Dim DateTime1 As String = ""
  Dim DateTime2 As String = ""
  Dim CashierSelected As Integer
  Dim StoreSelected As Integer
  Dim CardIdSelected As Integer
  Dim CardTypeSelected As Integer
  Dim DescSelected As Integer
  Dim DateTimeSelected As Integer
  Dim EnginesInstalled(-1) As Integer
  Dim TempDate As Date
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  
  Response.Expires = 0
  MyCommon.AppName = "CM-cashier-search.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  Send_HeadBegin("term.customer", "term.history")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts(New String() {"datePicker.js"})

 %>
<script type="text/javascript" language="javascript">
  window.name = "AdvSearch"
  var datePickerDivID = "datepicker";

  if (window.captureEvents){
      window.captureEvents(Event.CLICK);
      window.onclick=handlePageClick;
  }
  else {
      document.onclick=handlePageClick;
  }

<% Send_Calendar_Overrides(MyCommon) %>

  function handlePageClick(e) {
      var calFrame = document.getElementById('calendariframe');
      var el=(typeof event!=='undefined')? event.srcElement : e.target        
    
      if (el != null) {
          var pickerDiv = document.getElementById(datePickerDivID);
          if (pickerDiv != null && pickerDiv.style.visibility == "visible") {
              if (el.id!="datetimeDate1picker" && el.id!="datetimeDate2picker") {
                  if (!isDatePickerControl(el.className)) {
                    pickerDiv.style.visibility = "hidden";
                    pickerDiv.style.display = "none";            
                    if (calFrame != null) {
                      calFrame.style.visibility = "hidden";
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
  
  function handleDateToFrom(selIndex, tdName, elemName) {
    var elemTD = document.getElementById(tdName);
    var elem = document.getElementById(elemName);
    
    if (elemTD != null) {
      if (selIndex == 3) {
        elemTD.style.display = "";
      } else {
        if (elem !=null) elem.value = "";
        elemTD.style.display = "none";
      }
    }
  }
  
  function submitForm() {
    document.mainform.submit();
    if (window.opener != null && !window.opener.closed) {
      window.opener.focus();
      window.close();
    }
  }
  
  function validateDates() {
    var elemCreated1 = document.getElementById("datetimeDate1");
    var elemCreated2 = document.getElementById("datetimrDate2");
    var bValid = true;

    if (elemCreated1 != null && elemCreated1.value != "") {
      bValid = bValid && isDate(elemCreated1.value); 
    }
    if (elemCreated2 != null && elemCreated2.value != "") {
      bValid = bValid && isDate(elemCreated2.value);
      // check if the end date is after or on the start date
      if (bValid) {
        if (isStartAfterEnd(elemCreated1.value, elemCreated2.value)) {
          bValid = false;
          alert('<%Sendb(Copient.PhraseLib.Lookup("reports.startdate", LanguageID)) %>');
        }
      } 
    }
    
    return bValid;
  }
  
  function isStartAfterEnd(startDate, endDate) {
    var isAfter = false;
  
    var dtStart = Date.parse(startDate);
    var dtEnd = Date.parse(endDate);
    
    if (dtStart > dtEnd) {
      isAfter = true;
    }
    
    return isAfter;
  }
  
  
  function updateDateControls() {
    var elemCreated = document.getElementById("datetimeOption");
    var elemTdCreated = document.getElementById("tdDateTime");
    
    if (elemCreated != null && elemCreated.value == "3") {
      if (elemTdCreated != null) elemTdCreated.style.display = "block";
    }
    if (elemStart != null && elemStart.value == "3") {
      if (elemTdStart != null) elemTdStart.style.display = "block";
    }
    if (elemEnd != null && elemEnd.value == "3") {
      if (elemTdEnd != null) elemTdEnd.style.display = "block";
    }
  }
</script>

<%
  Send_HeadEnd()
  Send_BodyBegin(3)
  
  If (Logix.UserRoles.AccessOffers = False) Then
    Send_Denied(2, "perm.offers-access")
    GoTo done
  End If
  
  If (Request.QueryString("tokens") <> "") Then
    Dim TokenRows As String()
    Dim TokenCols As String()
    Dim i As Integer
    
    TokenString = Request.QueryString("tokens")
    TokenRows = TokenString.Split("|")
    For i = 0 To TokenRows.GetUpperBound(0)
      TokenCols = TokenRows(i).Split(",")
      If (TokenCols.Length >= 4) Then
        Select Case TokenCols(0).ToUpper
          Case "CASHIER"
            CashierSelected = Integer.Parse(TokenCols(1))
            Cashier = TokenCols(2)
          Case "STORE"
            StoreSelected = Integer.Parse(TokenCols(1))
            Store = TokenCols(2)
          Case "CARDID"
            CardIdSelected = Integer.Parse(TokenCols(1))
            CardId = TokenCols(2)
          Case "CARDTYPE"
            CardTypeSelected = Integer.Parse(TokenCols(1))
            CardType = TokenCols(2)
          Case "DESC"
            DescSelected = Integer.Parse(TokenCols(1))
            Desc = TokenCols(2)
          Case "DATETIME"
            DateTimeSelected = Integer.Parse(TokenCols(1))
            DateTime1 = TokenCols(2)
            If Date.TryParse(DateTime1, MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, TempDate) Then
              DateTime1 = Logix.ToShortDateString(TempDate, MyCommon)
            End If
            DateTime2 = TokenCols(3)
            If Date.TryParse(DateTime2, MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, TempDate) Then
              DateTime2 = Logix.ToShortDateString(TempDate, MyCommon)
            End If
        End Select
      End If
    Next
  End If
%>
<form action="CM-cashier-report.aspx" id="mainform" name="mainform" method="post" target="CahierReportWin">
  <input type="hidden" name="mode" id="mode" value="advancedsearch" />
  <div id="intro">
    <h1 id="title">
      <% Sendb(Copient.PhraseLib.Lookup("term.CashierReport", LanguageID) & " " & Copient.PhraseLib.Lookup("term.searchterms", LanguageID))%>
    </h1>
    <div id="controls">
    </div>
  </div>
  <div id="main">
    <%
      If (infoMessage <> "") Then
        Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
      End If
    %>
    <div id="column">
      <div class="box" id="criteria">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.criteria", LanguageID))%>
          </span>
        </h2>
        <center>
          <table style="width: 65%;" summary="<% Sendb(Copient.PhraseLib.Lookup("term.advancedsearchcriteria", LanguageID))%>">
            <tr>
              <td>
                <label for="datetimeOption"><% Sendb(Copient.PhraseLib.Lookup("term.date", LanguageID))%>:</label>
              </td>
              <td>
                <select id="datetimeOption" name="datetimeOption" class="mediumshort" onchange="handleDateToFrom(this.selectedIndex, 'tdDateTime', 'datetimeDate2');">
                  <option value="0"<% If(DateTimeSelected=0) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.on", LanguageID))%>
                  </option>
                  <option value="1"<% If(DateTimeSelected=1) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.before", LanguageID))%>
                  </option>
                  <option value="2"<% If(DateTimeSelected=2) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.after", LanguageID))%>
                  </option>
                  <option value="3"<% If(DateTimeSelected=3) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.between", LanguageID))%>
                  </option>
                </select>
              </td>
              <td>
                <input id="datetimeDate1" name="datetimeDate1" type="text" value="<% Sendb(DateTime1) %>" class="mediumshort" />
                <img src="../images/calendar.png" class="calendar" id="datetimeDate1picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('datetimeDate1',event);" />
              </td>
            </tr>
            <tr id="tdDateTime" style="display:none;">
              <td>
              </td>
              <td>
              </td>
              <td>
                <input id="datetimeDate2" name="datetimeDate2" type="text" value="<% Sendb(DateTime2) %>" class="mediumshort" />
                <img src="../images/calendar.png" class="calendar" id="datetimeDate2picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('datetimeDate2',event);" />
              </td>
            </tr>
            <tr>
              <td>
                <label for="cashierOption"><% Sendb(Copient.PhraseLib.Lookup("term.cashier", LanguageID))%>:</label>
              </td>
              <td>
                <select id="cashierOption" name="cashierOption" class="mediumshort">
                  <option value="1"<% If(CashierSelected=1) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.contains", LanguageID))%>
                  </option>
                  <option value="2"<% If(CashierSelected=2) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.exact", LanguageID))%>
                  </option>
                  <option value="3"<% If(CashierSelected=3) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.startswith", LanguageID))%>
                  </option>
                  <option value="4"<% If(CashierSelected=4) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.endswith", LanguageID))%>
                  </option>
                  <option value="5"<% If(CashierSelected=5) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.excludes", LanguageID))%>
                  </option>
                </select>
              </td>
              <td>
                <input id="cashier" name="cashier" type="text" value="<% Sendb(Cashier) %>" class="mediumshort" />
              </td>
            </tr>
            <tr>
              <td>
                <label for="storeOption"><% Sendb(Copient.PhraseLib.Lookup("term.store", LanguageID))%>:</label>
              </td>
              <td>
                <select id="storeOption" name="storeOption" class="mediumshort">
                  <option value="1"<% If(StoreSelected=1) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.contains", LanguageID))%>
                  </option>
                  <option value="2"<% If(StoreSelected=2) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.exact", LanguageID))%>
                  </option>
                  <option value="3"<% If(StoreSelected=3) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.startswith", LanguageID))%>
                  </option>
                  <option value="4"<% If(StoreSelected=4) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.endswith", LanguageID))%>
                  </option>
                  <option value="5"<% If(StoreSelected=5) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.excludes", LanguageID))%>
                  </option>
                </select>
              </td>
              <td>
                <input id="store" name="store" type="text" value="<% Sendb(Store) %>" class="mediumshort" /></td>
            </tr>
            <tr>
              <td>
                <label for="idOption"><% Sendb(Copient.PhraseLib.Lookup("term.card", LanguageID) & " " & Copient.PhraseLib.Lookup("term.id", LanguageID))%>:</label>
              </td>
              <td>
                <select id="cardidOption" name="cardidOption" class="mediumshort">
                  <option value="1"<% If(CardIdSelected=1) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.contains", LanguageID))%>
                  </option>
                  <option value="2"<% If(CardIdSelected=2) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.exact", LanguageID))%>
                  </option>
                  <option value="3"<% If(CardIdSelected=3) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.startswith", LanguageID))%>
                  </option>
                  <option value="4"<% If(CardIdSelected=4) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.endswith", LanguageID))%>
                  </option>
                  <option value="5"<% If(CardIdSelected=5) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.excludes", LanguageID))%>
                  </option>
                </select>
              </td>
              <td>
                <input id="cardid" name="cardid" type="text" value="<% Sendb(CardId) %>" class="mediumshort" />
              </td>
            </tr>
            <tr>
              <td>
                <label for="cardtypeOption"><% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID) & " " & Copient.PhraseLib.Lookup("term.type", LanguageID))%>:</label>
              </td>
              <td>
                <select id="cardtypeOption" name="cardtypeOption" class="mediumshort">
                  <option value="1"<% If(CardTypeSelected=1) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.contains", LanguageID))%>
                  </option>
                  <option value="2"<% If(CardTypeSelected=2) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.exact", LanguageID))%>
                  </option>
                  <option value="3"<% If(CardTypeSelected=3) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.startswith", LanguageID))%>
                  </option>
                  <option value="4"<% If(CardTypeSelected=4) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.endswith", LanguageID))%>
                  </option>
                  <option value="5"<% If(CardTypeSelected=5) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.excludes", LanguageID))%>
                  </option>
                </select>
              </td>
              <td>
                <input id="cardtype" name="cardtype" type="text" value="<% Sendb(CardType) %>" class="mediumshort" />
              </td>
            </tr>
            <tr>
              <td>
                <label for="descOption"><% Sendb(Copient.PhraseLib.Lookup("term.action", LanguageID))%>:</label>
              </td>
              <td>
                <select id="descOption" name="descOption" class="mediumshort">
                  <option value="1"<% If(DescSelected=1) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.contains", LanguageID))%>
                  </option>
                  <option value="2"<% If(DescSelected=2) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.exact", LanguageID))%>
                  </option>
                  <option value="3"<% If(DescSelected=3) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.startswith", LanguageID))%>
                  </option>
                  <option value="4"<% If(DescSelected=4) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.endswith", LanguageID))%>
                  </option>
                  <option value="5"<% If(DescSelected=5) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.excludes", LanguageID))%>
                  </option>
                </select>
              </td>
              <td>
                <input id="desc" name="desc" type="text" value="<% Sendb(Desc) %>" class="mediumshort" />
              </td>
            </tr>
          </table>
          <br />
          <input type="button" class="regular" id="search" name="search" value="<% Sendb(Copient.PhraseLib.Lookup("term.search", LanguageID))%>" onclick="return submitForm();" />
        </center>
        <hr class="hidden" />
      </div>

      <div id="datepicker" class="dpDiv">
      </div>
      <%
        If Request.Browser.Type = "IE6" Then
          Send("<iframe src=""javascript:'';"" id=""calendariframe"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no"" style=""visibility:hidden;display:none;""></iframe>")
        End If
      %>
    </div>
  </div>
</form>

<script type="text/javascript">
<% Send_Date_Picker_Terms() %>
  updateDateControls();
</script>

<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd("mainform")
  Logix = Nothing
  MyCommon = Nothing
%>
