<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Text" %>
<%@ Import Namespace="System.Web" %>
<%@ Import Namespace="Copient.commonShared" %>
<%
  ' *****************************************************************************
  ' * FILENAME: customer-transaction-inquiry.aspx 
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
  Dim Logix As New Copient.LogixInc
  Dim dt As DataTable
  Dim dr As DataTable
  Dim rst As DataTable
  Dim row As DataRow
  Dim i As Integer = 0
  Dim Shaded As String = " class=""shaded"""
  Dim HasSearchResults As Boolean = False

  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim PageName As String = ""
  
  ' default urls for links from this page
  Dim URLCustGen As String = "customer-general.aspx"
  Dim URLtrackBack As String = ""
  
  Response.Expires = 0
  MyCommon.AppName = "customer-transaction-inquiry.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixWH()
  Dim MyLookup As New Copient.CustomerLookup(MyCommon)
  Dim EmptyCriteria As Boolean
  Dim SessionID As String = ""
  Dim LogixTransNum As Integer

  Dim LocID_Name As String
  Dim LocationID As Integer
  Dim ExtLocationCode As String
  Dim LocationName As String
  Dim sLast4CardID As String
  Dim PresentedCustomerID As String
  Dim BannersEnabled As Boolean = False
  Dim FocusType As Integer = 0
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  

  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  MyLookup.SetAdminUserID(AdminUserID)
  MyLookup.SetLanguageID(LanguageID)
  
  LogixTransNum = 0
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  FocusType = IIf(Request.QueryString("Focus") = "1", 1, 0)
  
  Send_HeadBegin("term.transaction")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts(New String() {"datePicker.js"})
%>
<script type="text/javascript">
  var datePickerDivID = "datepicker";
  var isIE =(navigator.appName.indexOf("Microsoft")!=-1) ? true : false;
  var isOpera = (navigator.appName.indexOf("Opera")!=-1) ? true : false;
  var fullAvailList = null;

  function elmName(){
        window.onunload = null;
        for(i=0; i<document.mainform.elements.length; i++)
        {
            document.mainform.elements[i].disabled=false;
            //alert(document.mainform.elements[i].name)
        }
        return true;
    }
    
  function ValidateTimes() {

    var elemDispStartHr = document.getElementById("start-hr");
    var elemDispStartMin = document.getElementById("start-min");
    var elemDispEndHr = document.getElementById("end-hr");
    var elemDispEndMin = document.getElementById("end-min");
    var retVal = false;
    var validMinute = new RegExp('^[0-5][0-9]$');

    if (elemDispStartHr != null && elemDispStartHr.value != "" && (!isInteger(elemDispStartHr.value) || (parseInt(elemDispStartHr.value) <= 0) || (parseInt(elemDispStartHr.value) > 12))) {
      alert('<%Sendb(Copient.PhraseLib.Lookup("transactions.InvalidStartHour", LanguageID)) %>');
    } else if (elemDispStartMin != null && elemDispStartMin.value != "" && !validMinute.test(elemDispStartMin.value)) {
      alert('<%Sendb(Copient.PhraseLib.Lookup("transactions.InvalidStartMin", LanguageID)) %>');
    } else if (elemDispEndHr != null && elemDispEndHr.value != "" && (!isInteger(elemDispEndHr.value) || (parseInt(elemDispEndHr.value) <= 0) || (parseInt(elemDispEndHr.value) > 12))) {
      alert('<%Sendb(Copient.PhraseLib.Lookup("transactions.InvalidEndHour", LanguageID)) %>');
    } else if (elemDispEndMin != null && elemDispEndMin.value != "" && !validMinute.test(elemDispEndMin.value)) {
      alert('<%Sendb(Copient.PhraseLib.Lookup("transactions.InvalidEndMin", LanguageID)) %>');
    } else {
      retVal = true;
    }

    return retVal;
  }
  
  // This is the javascript array holding the function list
  // The PrintJavascriptArray ASP function can be used to print this array.
  // Locations array
  <% 
    LocID_Name = IIF(Request.QueryString("sgroups-available") <> "", Request.QueryString("sgroups-available"), Request.QueryString("LocCode"))
    MyCommon.QueryStr = "select LocationName from Locations as L with (NoLock) where Deleted=0 and LocationName not in ('" & LocID_Name & "') and Deleted=0 order by LocationName;"

    rst = MyCommon.LRT_Select
    
    If (rst.rows.count>0 )
      Sendb("var functionlist = Array(")
      For Each row In rst.Rows
        Sendb("""" & MyCommon.NZ(row.item("LocationName"), "").ToString().Replace("""", "\""") & """,")
      Next
      Send(""""");")
      Sendb("  var vallist = Array(")
      For Each row In rst.Rows
        Sendb("""" & row.item("LocationName") & """,")
      Next
      Send(""""");")
    Else
      Sendb("var functionlist = Array(")
      Send("""" & "" & """);")
      Sendb("  var vallist = Array(")
      Send("""" & "" & """);")
    End If
  %>
  
// This is the function that refreshes the list after a keypress.
  // The maximum number to show can be limited to improve performance with
  // huge lists (1000s of entries).
  // The function clears the list, and then does a linear search through the
  // globally defined array and adds the matches back to the list.
  function handleKeyUp(maxNumToShow) {
    var selectObj, textObj, functionListLength;
    var i,  numShown;
    var searchPattern;
    var newOpt, optGp;
    
    document.getElementById("sgroups-available").size = "10";
    
    // Set references to the form elements
    selectObj = document.getElementById("sgroups-available");
    textObj = document.forms[0].functioninput;
    
    // Remember the function list length for loop speedup
    functionListLength = functionlist.length;
    
    // Set the search pattern depending
    if(document.forms[0].functionradio[0].checked == true) {
      searchPattern = "^"+textObj.value;
    } else {
      searchPattern = textObj.value;
    }
    searchPattern = cleanRegExpString(searchPattern);
    
    // Create a regular expression
    re = new RegExp(searchPattern,"gi");
    
    // Clear the options list (for IE, cloning the select box without its option and then replacing the 
    // existing one, is significantly faster than removing each option.
    if (textObj.value == '' && fullAvailList != null)  {
      document.getElementById('sgrouplist').replaceChild(fullAvailList, selectObj);
    } else {
      //var newSelectBox = selectObj.cloneNode(false);
      if (isIE) {
        <%
            Send("var newSelectBox = document.createElement('select');")
            Send("newSelectBox.setAttribute('id', 'sgroups-available');")
            Send("newSelectBox.setAttribute('name', 'sgroups-available');")
            Send("newSelectBox.setAttribute('class', 'longest');")
            Send("newSelectBox.setAttribute('size', '10');")
        %>
      } else {
        var newSelectBox = document.createElement('select');
        newSelectBox.id = 'sgroups-available';
        newSelectBox.name = 'sgroups-available';
        newSelectBox.className = 'longest';
        newSelectBox.size = '10';
        newSelectBox.multiple = false;
      }
      
      document.getElementById('sgrouplist').replaceChild(newSelectBox, selectObj);
      selectObj = document.getElementById("sgroups-available");
     
      // Loop through the array and re-add matching options
      numShown = 0;
      for(i = 0; i < functionListLength; i++) {
        if(functionlist[i].search(re) != -1) {
          if (vallist[i] != "") {
            var newOpt = document.createElement('OPTION');
            newOpt.value = vallist[i];
            if (isIE) { newOpt.innerText = functionlist[i]}; 
            newOpt.text =  functionlist[i]; 
            
            <% If (BannersEnabled) Then %>
              if (!isOpera) {
                optGp = GetOptionGroup(bannerlist[i], selectObj);
                if (optGp != null) {
                  optGp.appendChild(newOpt);
                  selectObj.appendChild(optGp);
                } else {
                  selectObj[numShown] = newOpt
                }                
              } else {
                selectObj[numShown] = newOpt
              }
            <% Else %>
              selectObj[numShown] = new Option(newOpt.text, newOpt.value);
            <% End If %>
            if (vallist[i] == 1) {
              selectObj[numShown].style.fontWeight = 'bold';
              selectObj[numShown].style.color = 'brown';
            }
            numShown++;
          }
        }
        // Stop when the number to show is reached
        if(numShown == maxNumToShow) {
          break;
        }
      }
      // When options list whittled to one, select that entry
      if(selectObj.length == 1) {
        try {
          selectObj.options[0].selected = true;
        } catch (ex) {
          // ignore if unable to select (workaround for problem in IE 6)
        }
      }
    }
    if (textObj.value == '' && fullAvailList == null){
      fullAvailList = selectObj.cloneNode(true);
    }
  }
  
  function handleKeyDown(e, slctName) {
    var key = e.which ? e.which : e.keyCode;
    if (key == 40) {
      var elemSlct = document.getElementById(slctName);
      if (elemSlct != null) { elemSlct.focus(); }
    }
  }


</script>
<%
  Send_HeadEnd()
  
  ' SEARCHING FOR A Transaction ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  If (Request.QueryString("Search") <> "" Or Request.QueryString("searchPressed") <> "") Then
  
    If Session("SessionID") IsNot Nothing Then
      Session.Remove("SessionID")
    End If

  End If
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  
  If (Request.QueryString("Search") <> "" Or Request.QueryString("searchPressed") <> "") Then
    Dim sDate As String = Request.QueryString("transdate")
    Dim sStartHr As String = Request.QueryString("form_startHr")
    Dim sStartMin As String = Request.QueryString("form_startMin")
    Dim sEndHr As String = Request.QueryString("form_endHr")
    Dim sEndMin As String = Request.QueryString("form_endMin")
    Dim dtStartTime As Date
    Dim dtEndTime As Date

    LocID_Name = Request.QueryString("LocCode")
    sLast4CardID = Request.QueryString("Last4")
    
    'validate time entry
    If (sStartHr = "" AndAlso sStartMin <> "") Then
        infoMessage = IIf(String.IsNullOrEmpty(infoMessage), Copient.PhraseLib.Lookup("transactions.InvalidStartHour", LanguageID), infoMessage)
    Else If (sStartHr <> "" AndAlso sStartMin = "") Then
        infoMessage = IIf(String.IsNullOrEmpty(infoMessage), Copient.PhraseLib.Lookup("transactions.InvalidStartMin", LanguageID), infoMessage)
    Else If (sEndHr = "" AndAlso sEndMin <> "") Then
        infoMessage = IIf(String.IsNullOrEmpty(infoMessage), Copient.PhraseLib.Lookup("transactions.InvalidEndHour", LanguageID), infoMessage)
    Else If (sEndHr <> "" AndAlso sEndMin = "") Then
        infoMessage = IIf(String.IsNullOrEmpty(infoMessage), Copient.PhraseLib.Lookup("transactions.InvalidEndMin", LanguageID), infoMessage)
    Else If (sDate = "" AndAlso (sStartHr <> "" OrElse sStartMin <> "" OrElse sEndHr <> "" OrElse sEndMin <> "")) Then
      infoMessage = IIf(String.IsNullOrEmpty(infoMessage), Copient.PhraseLib.Lookup("transactions.NoDate", LanguageID), infoMessage)
    Else If sDate <> "" Then
      If sStartHr = "" AndAlso sStartMin = "" Then
        dtStartTime = DateTime.ParseExact(sDate, "M/d/yyyy", Nothing)
      Else
        dtStartTime = DateTime.ParseExact(sDate & " " & sStartHr & ":" & sStartMin & " " & Request.QueryString("AMPM"), "M/d/yyyy h:mm tt", Nothing)
      End If
      
      If sEndHr = "" AndAlso sEndMin = "" Then
        dtEndTime = DateTime.ParseExact(sDate & " 11:59 PM", "M/d/yyyy h:mm tt", Nothing)
      Else
        dtEndTime = DateTime.ParseExact(sDate & " " & sEndHr & ":" & sEndMin & " " & Request.QueryString("AMPM2"), "M/d/yyyy h:mm tt", Nothing)
      End If
      
      If DateTime.Compare(dtStartTime,dtEndTime) >= 0  Then
        infoMessage = IIf(String.IsNullOrEmpty(infoMessage), Copient.PhraseLib.Lookup("transactions.InvalidDateRange", LanguageID), infoMessage)
      End If
    End If
    'end validate time entry
        
    If sLast4CardID <> "" AndAlso sLast4CardID.length <> 4 Then 
      infoMessage = IIf(String.IsNullOrEmpty(infoMessage), Copient.PhraseLib.Lookup("transactions.InvalidLast4", LanguageID), infoMessage)
    End If
    
    If (LocID_Name <> "") Then
      
      MyCommon.QueryStr = "select top 1 LocationID, ExtLocationCode, LocationName from Locations with (NoLock) where LocationID like '%" & LocID_Name & "%' or LocationName like '%" & LocID_Name & "%' order by LocationID;"
      dt = MyCommon.LRT_Select
      If dt.Rows.Count > 0 Then     
        LocationID = MyCommon.NZ(dt.Rows(0).Item("LocationID"), 0)
        ExtLocationCode = MyCommon.NZ(dt.Rows(0).Item("ExtLocationCode"), 0)
        LocationName = MyCommon.NZ(dt.Rows(0).Item("LocationName"), 0)
        MyCommon.QueryStr = "select top 1 LogixTransNum from TransHist with (NoLock) where ExtLocationCode like '%" & ExtLocationCode & "%'  and PresentedCustomerID like '%"& sLast4CardID &"';"
        dr = MyCommon.LWH_Select
        If dr.Rows.Count > 0 Then 
          HasSearchResults = true
        Else
         infoMessage = IIf(String.IsNullOrEmpty(infoMessage), Copient.PhraseLib.Lookup("transactions.NoTrans", LanguageID), infoMessage)
        End If
      Else
        infoMessage = IIf(String.IsNullOrEmpty(infoMessage), Copient.PhraseLib.Lookup("transactions.NoStores", LanguageID), infoMessage)
      End If        
                  
    Else
      ' IIf is added to prevent any Existing Error Message from getting modified
      infoMessage = IIf(String.IsNullOrEmpty(infoMessage), "Please select a store", infoMessage)
    End If
    
    If ((HasSearchResults=True) And (infoMessage = "")) Then
      Response.Redirect("customer-transactions.aspx?LocCode=" & ExtLocationCode & IIF(sLast4CardID <> "","&Last4=" & sLast4CardID, "") & IIF(sDate <> "", "&StartTime=" & dtStartTime.toString("M/d/yyyy H:mm") & "&EndTime=" & dtEndTime.toString("M/d/yyyy H:mm"), "") )
    End If

  End If
  
  
  Send_BodyBegin(1)
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  Send_Tabs(Logix, 3)
  Send_Subtabs(Logix, 35, 0, LanguageID)
  
  If (Logix.UserRoles.AccessTransactionInquiry = False) Then
    Send_Denied(1, "perm.transactions-access")
    GoTo done
  End If
%>
<form id="mainform" name="mainform" action="customer-transaction-inquiry.aspx" onsubmit="elmName(); return ValidateTimes();">
<div id="intro">
  <h1 id="title">
    <%
      Sendb(Copient.PhraseLib.Lookup("term.transaction", LanguageID))
    %>
  </h1>
  <div id="controls">
    <%
     
    %>
  </div>
</div>
<div id="main">
  <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
  <%Sendb(Copient.PhraseLib.Lookup("transactions.StoreSearch", LanguageID))%>
  <br />
  <br class="half" />
  <%
    If (Request.QueryString("mode") = "summary") Then
      Send("<input type=""hidden"" id=""mode"" name=""mode"" value=""summary"" />")
      Send("<input type=""hidden"" id=""exiturl"" name=""exiturl"" value=""" & URLtrackBack & """ />")
    End If
  %>
  
  <div id="column1">
      <div class="box" id="storegroups">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.stores", LanguageID))%>
          </span>
        </h2>
        <label for="sgroups-available">
          <b>
            <% Sendb(Copient.PhraseLib.Lookup("term.available", LanguageID) & ":")%>
          </b>
        </label>
        <br clear="all" />
        <%Send("<input type=""radio"" id=""functionradio1"" name=""functionradio"" onchange=""handleKeyUp();"" value=""1""" & IIf(Request.QueryString("functionradio") <> "2", " checked=""checked""", "") & "/>")%>
        <label for="functionradio1"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
        <%Send("<input type=""radio"" id=""functionradio2"" name=""functionradio"" onchange=""handleKeyUp();"" value=""2""" & IIf(Request.QueryString("functionradio") = "2", " checked=""checked""", "") & "/>")%>
        <label for="functionradio2"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label>
        <br />
        
        <input class="longer" onkeydown="handleKeyDown();" onkeyup="handleKeyUp();" id="functioninput" name="functioninput" maxlength="100" type="text" value="<%Sendb(Request.QueryString("functioninput")) %>" />
        <br />
        <br class="half" />
        <%
          Send("<span id=""sgrouplist"">")
          Send("<select class=""longest"" id=""sgroups-available"" name=""sgroups-available"" size=""10"">")
          Send("</select>")
          Send("</span>")
          Send("<br />")
          Send("<br />")
                    
          ' SELECTED STORE GROUPS
          Send("<label for=""sgroups-select""><b>" & Copient.PhraseLib.Lookup("term.selected", LanguageID) & ":</b></label>")
          Send("<br />")
          
          ' Buttons
          Sendb("<input type=""submit"" class=""regular select"" id=""stores-add1"" name=""stores-add1"" title=""" & Copient.PhraseLib.Lookup("term.select", LanguageID) & """ value=""&#9660; " & Copient.PhraseLib.Lookup("term.select", LanguageID) & """" )
          
          Send(" />")
          Send("<br />")

          ' List
          Send("<input type=""text"" class=""longest"" id=""LocCode"" name=""LocCode"" size=""10"" value=""" & LocID_Name & """>")
          
          Send("<br />")
          Send("<br class=""half"" />")
          
        %>
        <hr class="hidden" />
      </div>
      <a name="h01"></a>
    </div>
    
    <div id="gutter">
    </div>
  
  <div id="column2">
    <table width="100%" summary="<% Sendb(Copient.PhraseLib.Lookup("term.criteria", LanguageID))%>">
      <thead>
        <tr style="visibility: hidden;">
          <th style="width: 85px;">
          </th>
        </tr>
      </thead>
      <tbody>
        <table id="trCustom" >
          <tr>
          <td style="width: 85px;">
            <label for="transdate">
              <% Sendb(Copient.PhraseLib.Lookup("term.date", LanguageID))%>
              :</label>            
          </td>
          <td>
              <input class="short" id="transdate" name="transdate" maxlength="10"
                type="text" value="<% Sendb(Request.QueryString("transdate")) %>" />
              <img src="../images/calendar.png" class="calendar" id="picker-date" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
               title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('transdate', event);" />
                (m/d/yyyy)
          </td>
          </tr>
          <hr class="hidden" />
          <div id="datepicker" class="dpDiv">
          </div>
          <%
            If Request.Browser.Type = "IE6" Then
              Send("<iframe src=""javascript:'';"" id=""calendariframe"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no"" style=""visibility:hidden;display:none;""></iframe>")
            End If
          %>
          <tr>
          <td style="width: 85px;">
            <label for="start-time">
              <% Sendb(Copient.PhraseLib.Lookup("term.time", LanguageID))%>
              :</label>
          </td>
          <td>
              <input class="shortest" id="start-hr" maxlength="2" name="form_StartHr"
              type="text" value="<%Sendb(Request.QueryString("form_StartHr")) %>" />:<input class="shortest" id="start-min" maxlength="2" name="form_StartMin" type="text" value="<%Sendb(Request.QueryString("form_StartMin")) %>" />
              <select name="AMPM" id="AMPM">
              <%
                Send("<option value=""AM""" & IIf(Request.QueryString("AMPM") = "AM", " selected=""selected""", "") & "> AM </option>")
                Send("<option value=""PM""" & IIf(Request.QueryString("AMPM") = "PM", " selected=""selected""", "") & "> PM </option>")
              %>
              </select>
            
              <%Sendb(Copient.PhraseLib.Lookup("term.to", LanguageID))%>
                            
              <input class="shortest" id="end-hr" maxlength="2" name="form_EndHr"
              type="text" value="<%Sendb(Request.QueryString("form_EndHr")) %>" />:<input class="shortest" id="end-min" maxlength="2" name="form_EndMin" type="text" value="<%Sendb(Request.QueryString("form_EndMin")) %>" />
                
              <select name="AMPM2" id="AMPM2">
              <%
                Send("<option value=""AM""" & IIf(Request.QueryString("AMPM2") = "AM", " selected=""selected""", "") & "> AM </option>")
                Send("<option value=""PM""" & IIf(Request.QueryString("AMPM2") = "PM", " selected=""selected""", "") & "> PM </option>")
              %>
              </select>
               
          </td>
          </tr>
          <tr>
          <td style="width: 85px;">
            <label for="Last4">
              <%Sendb(Copient.PhraseLib.Lookup("transactions.Last4", LanguageID))%>
            </label>
          </td>
          <td>
            <input type="text" class="short" id="Last4" name="Last4" maxlength="4" value="<%Sendb(Request.QueryString("Last4")) %>" />
          </td>
          </tr>
        </table>
        <br />
        <tr id="trSubmit">
          <td>
            &nbsp;
          </td>
          <td>
            <input type="submit" class="regular" id="search" name="search" value="<% Sendb(Copient.PhraseLib.Lookup("term.search", LanguageID)) %>" />
          </td>
        </tr>
      </tbody>
    </table>
    <br />    
  </div>
</div>
</form>
<script type="text/javascript">

  handleKeyUp();
  
  <% Send_Date_Picker_Terms() %>
</script>
<%
done:
  Send_BodyEnd("mainform", "functioninput")
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixWH()
  MyCommon = Nothing
  Logix = Nothing
%>