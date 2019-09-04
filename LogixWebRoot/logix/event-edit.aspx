<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: events.aspx 
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
  Dim Logix As New Copient.LogixInc
  Dim dst As System.Data.DataTable
  Dim rst As System.Data.DataTable
  Dim row As DataRow
  Dim AdminUserID As Long
  Dim CloseAfterSave As Boolean = False
  Dim EventTitle As String = ""
  Dim pgName As String = ""
  
  Dim EventID As Integer = 0
  Dim Description As String = ""
  Dim Recurrence As Integer = 0
  Dim FixedDate As Boolean = True
  Dim EventDate As String = ""
  Dim Day As Integer = 0
  Dim Ordinal As Integer = 0
  Dim Month As Integer = 0
  Dim Year As Integer = 0
  
  Dim eSave As Boolean = False
  Dim eDelete As Boolean = False
  Dim eCreate As Boolean = False
  
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "events.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  EventID = MyCommon.NZ(Request.QueryString("EventID"), "0")
  
  
  ' --------------------------------------------------
  If (Request.QueryString("Delete") <> "") Then
    If (MyCommon.Extract_Val(Request.Form("EventID")) <> "") Then
      EventID = MyCommon.Extract_Val(Request.Form("EventID"))
      MyCommon.QueryStr = "UPDATE Events WITH (RowLock) SET Deleted=1 where EventID=" & EventID
      MyCommon.LRT_Execute()
      Response.Status = "301 Moved Permanently"
      Response.AddHeader("Location", "events.aspx")
    End If
  ElseIf (Request.QueryString("save") <> "") Then
    EventID = MyCommon.Extract_Val(Request.Form("EventID"))
    If EventID = 0 Then
      ' Insert a new event into the table
      MyCommon.QueryStr = "dbo.pt_Event_Insert"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 100).Value = Request.Form("description")
      MyCommon.LRTsp.Parameters.Add("@Recurrence", SqlDbType.Int).Value = Request.Form("recurrence")
      MyCommon.LRTsp.Parameters.Add("@FixedDate", SqlDbType.Bit).Value = Request.Form("fixeddate")
      MyCommon.LRTsp.Parameters.Add("@EventDate", SqlDbType.DateTime).Value = Request.Form("eventdate")
      MyCommon.LRTsp.Parameters.Add("@Ordinal", SqlDbType.Int).Value = Request.Form("ordinal")
      MyCommon.LRTsp.Parameters.Add("@Day", SqlDbType.Int).Value = Request.Form("day")
      MyCommon.LRTsp.Parameters.Add("@Month", SqlDbType.Int).Value = Request.Form("month")
      MyCommon.LRTsp.Parameters.Add("@Year", SqlDbType.Int).Value = Request.Form("year")
      MyCommon.LRTsp.Parameters.Add("@EventID", SqlDbType.Int).Direction = ParameterDirection.Output
      If (Logix.TrimAll(Request.Form("description")) = "") Then
        infoMessage = Copient.PhraseLib.Lookup("event.noname", LanguageID)
      ElseIf (Request.Form("fixeddate") = 1) And (Request.Form("eventdate") = "") Then
        infoMessage = Copient.PhraseLib.Lookup("event.baddate", LanguageID)
      Else
        MyCommon.LRTsp.ExecuteNonQuery()
        EventID = MyCommon.LRTsp.Parameters("@EventID").Value
      End If
      MyCommon.Close_LRTsp()
    Else
      ' Update an existing event
      If Logix.TrimAll(Request.Form("description")) = "" Then
        infoMessage = Copient.PhraseLib.Lookup("event.noname", LanguageID)
      Else
        MyCommon.QueryStr = "UPDATE Events WITH (RowLock) SET " & _
                            "Description='" & Request.Form("description") & "', " & _
                            "Recurrence=" & Request.Form("recurrence") & ", " & _
                            "FixedDate=" & Request.Form("fixeddate") & ", " & _
                            "EventDate=" & Request.Form("eventdate") & ", " & _
                            "Ordinal=" & Request.Form("ordinal") & ", " & _
                            "Day=" & Request.Form("day") & ", " & _
                            "Month=" & Request.Form("month") & ", " & _
                            "Year=" & Request.Form("year") & " " & _
                            "WHERE EventID=" & EventID & ";"
        MyCommon.LRT_Execute()
      End If
    End If
  End If
  ' --------------------------------------------------
  
  
  ' Grab the event
  MyCommon.QueryStr = "SELECT * FROM Events AS E WITH (NoLock) WHERE Deleted=0 AND EventID='" & EventID & "';"
  rst = MyCommon.LRT_Select
  If (rst.Rows.Count > 0) Then
    EventID = MyCommon.NZ(rst.Rows(0).Item("EventID"), 0)
    Description = MyCommon.NZ(rst.Rows(0).Item("Description"), "")
    Recurrence = MyCommon.NZ(rst.Rows(0).Item("Recurrence"), 0)
    FixedDate = MyCommon.NZ(rst.Rows(0).Item("FixedDate"), 1)
    EventDate = MyCommon.NZ(rst.Rows(0).Item("EventDate"), "")
    Ordinal = MyCommon.NZ(rst.Rows(0).Item("Ordinal"), 0)
    Day = MyCommon.NZ(rst.Rows(0).Item("Day"), 0)
    Month = MyCommon.NZ(rst.Rows(0).Item("Month"), 0)
    Year = MyCommon.NZ(rst.Rows(0).Item("Year"), 0)
  ElseIf (Request.QueryString("new") <> "New") And (EventID > 0) Then
    Send_HeadBegin("term.event", , EventID)
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts(New String() {"datePicker.js"})
    Send_HeadEnd()
    Send_BodyBegin(3)
    Send("")
    Send("<div id=""intro"">")
    Send("    <h1 id=""title"">" & Copient.PhraseLib.Lookup("term.event", LanguageID) & " #" & EventID & "</h1>")
    Send("</div>")
    Send("<div id=""main"">")
    Send("    <div id=""infobar"" class=""red-background"">")
    Send("        " & Copient.PhraseLib.Lookup("error.deleted", LanguageID))
    Send("    </div>")
    Send("</div>")
    Send_BodyEnd()
    GoTo done
  Else
    EventID = "0"
    Description = ""
    pgName = Copient.PhraseLib.Lookup("term.newevent", LanguageID)
  End If
  
  Send_HeadBegin("term.event", , EventID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts(New String() {"datePicker.js"})
  Send_HeadEnd()
  Send_BodyBegin(3)
  
  If (Logix.UserRoles.EditSystemConfiguration = False) Then
    Send_Denied(1, "perm.admin-configuration")
    GoTo done
  End If
%>

<script type="text/javascript" language="javascript">
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
            if (el.id!="fixed-date-picker") {

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
                  calFrame.style.visibility = "visible";
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
</script>

<form action="#" id="mainform" name="mainform" method="post">
  <div id="intro">
    <h1 id="title">
      <%
        If EventID = 0 Then
          Sendb(Copient.PhraseLib.Lookup("term.newevent", LanguageID))
        Else
          Sendb(Copient.PhraseLib.Lookup("term.event", LanguageID) & " #" & EventID & ": ")
          EventTitle = Description
          Sendb(MyCommon.TruncateString(EventTitle, 40))
        End If
      %>
    </h1>
    <div id="controls">
      <% Send_Save() %>
      <%
        If EventID > 0 Then
          Send_Delete()
        End If
      %>
    </div>
  </div>
  <div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <div id="column1">
      <div class="box" id="general">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.general", LanguageID))%>
          </span>
        </h2>
        <label for="description"><% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>:</label><br />
        <input type="text" class="long" id="description" name="description" value="<% Sendb(Description) %>" maxlength="100" />
        <input type="hidden" id="EventID" name="EventID" value="<% Sendb(EventID) %>" />
        <hr class="hidden" />
      </div>
      <div class="box" id="occurrence">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.occurrence", LanguageID))%>
          </span>
        </h2>
        <label for="recurrence"><% Sendb(Copient.PhraseLib.Lookup("term.frequency", LanguageID))%>:</label><br />
        <select id="recurrence" name="recurrence">
          <option value="1"<%If (Recurrence = 1) Then Sendb(" selected=""selected""") %>><% Sendb(Copient.PhraseLib.Lookup("event.occursonce", LanguageID))%></option>
          <option value="0"<%If (Recurrence = 0) Then Sendb(" selected=""selected""") %>><% Sendb(Copient.PhraseLib.Lookup("event.recursannually", LanguageID))%></option>
        </select>
        <br />
        
        <br class="half" />
        
        <input type="radio" id="fixed-on" name="fixed" value="1"<% Sendb(IIf(FixedDate, " checked=""checked""", "")) %> /><label for="fixed-on"><% Sendb(Copient.PhraseLib.Lookup("event.fixeddate", LanguageID))%>:</label><br />
        <div id="fixeddate" style="padding-left: 22px;margin-bottom:5px;">
          <input type="text" class="short" id="eventdate" name="eventdate" maxlength="10" value="<% Sendb(EventDate) %>" />
          <img src="../images/calendar.png" class="calendar" id="fixed-date-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('eventdate', event);" />
        </div>
        
        <input type="radio" id="fixed-off" name="fixed" value="0"<% Sendb(IIf(FixedDate, "", " checked=""checked""")) %> /><label for="fixed-off"><% Sendb(Copient.PhraseLib.Lookup("event.variabledate", LanguageID))%>:</label><br />
        <div id="variabledate" style="padding-left: 22px;margin-bottom:5px;">
          <select id="dayordinal" name="dayordinal">
            <option value="0">&nbsp;</option>
            <option value="1"<% Sendb(IIf(Ordinal = 1, " selected=""selected""", "")) %>><% Sendb(Copient.PhraseLib.Lookup("term.first", LanguageID))%></option>
            <option value="2"<% Sendb(IIf(Ordinal = 2, " selected=""selected""", "")) %>><% Sendb(Copient.PhraseLib.Lookup("term.second", LanguageID))%></option>
            <option value="3"<% Sendb(IIf(Ordinal = 3, " selected=""selected""", "")) %>><% Sendb(Copient.PhraseLib.Lookup("term.third", LanguageID))%></option>
            <option value="4"<% Sendb(IIf(Ordinal = 4, " selected=""selected""", "")) %>><% Sendb(Copient.PhraseLib.Lookup("term.fourth", LanguageID))%></option>
            <option value="5"<% Sendb(IIf(Ordinal = 5, " selected=""selected""", "")) %>><% Sendb(Copient.PhraseLib.Lookup("term.fifth", LanguageID))%></option>
          </select>
          <select id="day" name="day">
            <option value="0">&nbsp;</option>
            <option value="1"<% Sendb(IIf(Day = 1, " selected=""selected""", "")) %>><% Sendb(Copient.PhraseLib.Lookup("term.sunday", LanguageID))%></option>
            <option value="2"<% Sendb(IIf(Day = 2, " selected=""selected""", "")) %>><% Sendb(Copient.PhraseLib.Lookup("term.monday", LanguageID))%></option>
            <option value="3"<% Sendb(IIf(Day = 3, " selected=""selected""", "")) %>><% Sendb(Copient.PhraseLib.Lookup("term.tuesday", LanguageID))%></option>
            <option value="4"<% Sendb(IIf(Day = 4, " selected=""selected""", "")) %>><% Sendb(Copient.PhraseLib.Lookup("term.wednesday", LanguageID))%></option>
            <option value="5"<% Sendb(IIf(Day = 5, " selected=""selected""", "")) %>><% Sendb(Copient.PhraseLib.Lookup("term.thursday", LanguageID))%></option>
            <option value="6"<% Sendb(IIf(Day = 6, " selected=""selected""", "")) %>><% Sendb(Copient.PhraseLib.Lookup("term.friday", LanguageID))%></option>
            <option value="7"<% Sendb(IIf(Day = 7, " selected=""selected""", "")) %>><% Sendb(Copient.PhraseLib.Lookup("term.saturday", LanguageID))%></option>
          </select>
          <% Sendb(Copient.PhraseLib.Lookup("term.of", LanguageID))%><br />
          <select id="month" name="month" style="margin-top:3px;">
            <option value="0">&nbsp;</option>
            <option value="1"<% Sendb(IIf(Month = 1, " selected=""selected""", "")) %>><% Sendb(Copient.PhraseLib.Lookup("term.january", LanguageID))%></option>
            <option value="2"<% Sendb(IIf(Month = 2, " selected=""selected""", "")) %>><% Sendb(Copient.PhraseLib.Lookup("term.february", LanguageID))%></option>
            <option value="3"<% Sendb(IIf(Month = 3, " selected=""selected""", "")) %>><% Sendb(Copient.PhraseLib.Lookup("term.march", LanguageID))%></option>
            <option value="4"<% Sendb(IIf(Month = 4, " selected=""selected""", "")) %>><% Sendb(Copient.PhraseLib.Lookup("term.april", LanguageID))%></option>
            <option value="5"<% Sendb(IIf(Month = 5, " selected=""selected""", "")) %>><% Sendb(Copient.PhraseLib.Lookup("term.may", LanguageID))%></option>
            <option value="6"<% Sendb(IIf(Month = 6, " selected=""selected""", "")) %>><% Sendb(Copient.PhraseLib.Lookup("term.june", LanguageID))%></option>
            <option value="7"<% Sendb(IIf(Month = 7, " selected=""selected""", "")) %>><% Sendb(Copient.PhraseLib.Lookup("term.july", LanguageID))%></option>
            <option value="8"<% Sendb(IIf(Month = 8, " selected=""selected""", "")) %>><% Sendb(Copient.PhraseLib.Lookup("term.august", LanguageID))%></option>
            <option value="9"<% Sendb(IIf(Month = 9, " selected=""selected""", "")) %>><% Sendb(Copient.PhraseLib.Lookup("term.september", LanguageID))%></option>
            <option value="10"<% Sendb(IIf(Month = 10, " selected=""selected""", "")) %>><% Sendb(Copient.PhraseLib.Lookup("term.october", LanguageID))%></option>
            <option value="11"<% Sendb(IIf(Month = 11, " selected=""selected""", "")) %>><% Sendb(Copient.PhraseLib.Lookup("term.november", LanguageID))%></option>
            <option value="12"<% Sendb(IIf(Month = 12, " selected=""selected""", "")) %>><% Sendb(Copient.PhraseLib.Lookup("term.december", LanguageID))%></option>
          </select>
          <input type="text" id="year" name="year" class="shorter" maxlength="4" value="<% Sendb(IIf(Year = 0, DatePart(DateInterval.Year, Now), Year)) %>" style="margin-top:3px;" />
        </div>
        
      </div>
      <div id="datepicker" class="dpDiv">
      </div>
      <%
        If Request.Browser.Type = "IE6" Then
          Send("<iframe src=""javascript:'';"" id=""calendariframe"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no"" style=""visibility:hidden;display:none;""></iframe>")
        End If
      %>
    </div>
    
    <br clear="all" />
  </div>
</form>

<script type="text/javascript">
<% Send_Date_Picker_Terms() %>
<% If (CloseAfterSave) Then %>
    window.close();
<% End If %>
</script>

<%
done:
  Send_BodyEnd("mainform", "description")
  MyCommon.Close_LogixRT()
  MyCommon = Nothing
  Logix = Nothing
%>
