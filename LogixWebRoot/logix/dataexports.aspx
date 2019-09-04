<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: dataexports.aspx 
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
  Dim dst As DataTable
  Dim rst1 As DataTable
  Dim rst2 As DataTable
  Dim rst3 As DataTable
  Dim rstIssuance As DataTable
  Dim row As DataRow
  Dim row2 As DataRow
  Dim row3 As DataRow
  
  Dim selectedStr As String = ""
  Dim i As Integer = 0
  Dim tempdate As New Date
  Dim tod As String = "00:00"
  Dim todhour As String = "0"
  Dim todminute As String = "0"
  Dim StartDay As String = ""
  Dim IsValidDate As Boolean = False
  Dim ScheduleID As Integer = 0
  Dim ScheduleEnabled As Boolean = True
  Dim SchedulePath As String = ""
  Dim ScheduleDate As Date
  Dim OutputZipped As Boolean = False
  Dim OutputPath As String = ""
  Dim DaysChecked As New BitArray(8, False)
  Dim DayIndex As Integer = 0
  Dim LastRan As String = ""
  
  Dim RemotePKID As Integer = 0
  Dim RemoteEnabled As Boolean = False
  Dim RemotePath As String = ""
  Dim LastUpdate As String = ""
  Dim FilePathRequired As Boolean = False
  Dim RemoteDataType As Integer = 0
  Dim RemoteStyleDesc As String = ""
  Dim IssuanceEnabled As Boolean = False
  
  Dim AdminUserID As Long
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "dataexports.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  Send_HeadBegin("term.dataexports")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts(New String() {"datePicker.js"})
%>
<script type="text/javascript" language="javascript">
  var datePickerDivID = "datepicker";
  
  if (window.captureEvents){
    window.captureEvents(Event.CLICK);
    window.onclick=handlePageClick;
  } else {
    document.onclick=handlePageClick;
  }
function RestrictSpace() {
    if (event.keyCode == 32) {
        return false;
    }
}
 
<% Send_Calendar_Overrides(MyCommon) %>
  
  function handlePageClick(e) {
    var calFrame = document.getElementById('calendariframe');
    var el=(typeof event!=='undefined')? event.srcElement : e.target        
    
    if (el != null) {
      var pickerDiv = document.getElementById(datePickerDivID);
      if (pickerDiv != null && pickerDiv.style.visibility == "visible") {
        if (el.id!="deo-start-picker") {
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
  
  function updateRdoEnabled(pkid, checked) {
    var elemName = "rdo-enable" + pkid;
    var elem = document.getElementById(elemName);
    
    if (elem != null) {
      elem.value = (checked) ? "1" : "0"
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
  Send_Subtabs(Logix, 8, 4)
  
  
    
  If (Logix.UserRoles.AccessSystemSettings = False) Then
    Send_Denied(1, "perm.admin-settings")
    GoTo done
  End If
  
  IssuanceEnabled = (MyCommon.Extract_Val(MyCommon.Fetch_CPE_SystemOption(70)) = 1)
  
  ' query for data export options
  MyCommon.QueryStr = "select DEO.DataExportOptionID, DEO.DataExportTypeID, DET.Description as TypeDescription," & _
                      "DEO.StyleID, DES.Description as StyleDescription,DEO.Enabled, DEO.OutputPath, DEO.LastRunTime," & _
                      "DEC.ScheduleID, DEC.ScheduleTypeID, DEC.StartDateTime, DEO.Zipped, DET.PhraseID " & _
                      "from DataExportOptions as DEO " & _
                      "inner join DataExportSchedules as DEC on DEC.DataExportOptionID=DEC.DataExportOptionID " & _
                      "inner join DataExportTypes as DET on DET.DataExportTypeID=DEO.DataExportTypeID " & _
                      "inner join DataExportStyles as DES on DES.StyleID=DEO.StyleID;"
  rst1 = MyCommon.LRT_Select
  ' query for remote data options
  MyCommon.QueryStr = "select RDO.PKID, RDO.RemoteDataTypeID, RDT.Description as TypeDescription, RDO.StyleID, RDS.Description as StyleDescription, " & _
                      "RDO.Enabled, RDO.OutputPath, RDO.LastUpdate, RDS.FilePathRequired, RDT.PhraseID " & _
                      "from RemoteDataOptions as RDO " & _
                      "inner join RemoteDataTypes as RDT on RDT.RemoteDataTypeID=RDO.RemoteDataTypeID " & _
                      "inner join RemoteDataStyles as RDS on RDS.StyleID=RDO.StyleID and RDS.RemoteDataTypeID = RDO.RemoteDataTypeID " & _
                      "where RDO.RemoteDataTypeID <> 1 order by RemoteDataTypeID desc;"
  rst2 = MyCommon.LRT_Select
  ' remote data option for issuance only
  MyCommon.QueryStr = "select RDO.PKID, RDO.RemoteDataTypeID, RDT.Description as TypeDescription, RDO.StyleID, RDS.Description as StyleDescription, " & _
                      "RDO.Enabled, RDO.OutputPath, RDO.LastUpdate, RDS.FilePathRequired " & _
                      "from RemoteDataOptions as RDO " & _
                      "inner join RemoteDataTypes as RDT on RDT.RemoteDataTypeID=RDO.RemoteDataTypeID " & _
                      "inner join RemoteDataStyles as RDS on RDS.StyleID=RDO.StyleID and RDS.RemoteDataTypeID = RDO.RemoteDataTypeID " & _
                      "where RDO.RemoteDataTypeID = 1 order by RemoteDataTypeID desc;"
  rstIssuance = MyCommon.LRT_Select
    If (Server.HtmlEncode(Request.QueryString("save")) <> "") Then
    If rst1.Rows.Count > 0 Then
      IsValidDate = Date.TryParse(Request.QueryString("deo-start") & " " & Request.QueryString("hours") & ":" & _
                        Request.QueryString("minutes") & ":00", tempdate)
      If Not IsValidDate Then
        infoMessage = Copient.PhraseLib.Lookup("term.InvalidDate", LanguageID)
      Else
        If rst1.Rows.Count > 0 Then
                    ScheduleID = MyCommon.Extract_Val(Server.HtmlEncode(Request.QueryString("scheduleid")))
          ' update the deo output path and enabled state
                    MyCommon.QueryStr = "update DataExportOptions with (RowLock) set Enabled= " & IIf(Server.HtmlEncode(Request.QueryString("deo-enable")) = "", 0, 1) & ", " & _
                                        "OutputPath='" & MyCommon.Parse_Quotes(Server.HtmlEncode(Request.QueryString("deo-path"))) & "', LastUpdate=getdate(), " & _
                              "Zipped=" & IIf(Request.QueryString("deo-zipped") = "", 0, 1) & " " & _
                                        "where DataExportOptionID=" & MyCommon.Extract_Val(Server.HtmlEncode(Request.QueryString("deo-id")))
          MyCommon.LRT_Execute()
          
          ' update the start date and time for this data export
          MyCommon.QueryStr = "update DataExportSchedules with (RowLock) set StartDateTime='" & tempdate.ToString("yyyy-MM-dd HH:mm:ss") & "', LastUpdate=getdate() " & _
                              "where ScheduleID=" & ScheduleID
          MyCommon.LRT_Execute()
          
          ' clear the scheduled days before re-adding the selected days
          MyCommon.QueryStr = "delete DataExportScheduleValues with (RowLock) where ScheduleID = " & ScheduleID
          MyCommon.LRT_Execute()
          
          ' add the selected days to the schedule for this data export
          If (Request.QueryString("day") <> "") Then
            For i = 0 To Request.QueryString.GetValues("day").GetUpperBound(0)
              MyCommon.QueryStr = "insert into DataExportScheduleValues (ScheduleID, Value) values (" & ScheduleID & _
                                  ", " & MyCommon.Extract_Val(Request.QueryString.GetValues("day")(i)) & ");"
              MyCommon.LRT_Execute()
            Next
          End If
        End If
      End If
    End If
    
    If rst2.Rows.Count > 0 OrElse rstIssuance.Rows.Count > 0 Then
      If (Request.QueryString("rdo-enable") <> "") Then
        For i = 0 To Request.QueryString.GetValues("rdo-enable").GetUpperBound(0)
          ' update the remote data settings
          MyCommon.QueryStr = "update RemoteDataOptions with (RowLock) set Enabled=" & Request.QueryString.GetValues("rdo-enable")(i) & _
                                      ", OutputPath='" & MyCommon.Parse_Quotes(Request.QueryString.GetValues("rdo-path")(i)) & "', CPEUpdateFlag=1 " & _
                            "where PKID=" & MyCommon.Extract_Val(Request.QueryString.GetValues("remotepkid")(i))
          MyCommon.LRT_Execute()
        Next
      End If
    End If
    
    MyCommon.Activity_Log(24, 0, AdminUserID, Copient.PhraseLib.Lookup("history.settings", LanguageID))
    Response.Redirect("dataexports.aspx?dataexport=Data+exports")
  End If
%>
<form action="#" id="mainform" name="mainform">
  <div id="intro">
    <h1 id="title">
      <% Sendb(Copient.PhraseLib.Lookup("term.dataexports", LanguageID))%>
    </h1>
    <div id="controls">
      <%
        If (Logix.UserRoles.EditDataExports = True) Then
          Send_Save()
        End If
        'If MyCommon.Fetch_SystemOption(75) Then
        '  If (Logix.UserRoles.AccessNotes) Then
        '    Send_NotesButton(21, 0, AdminUserID)
        '  End If
        'End If
      %>
    </div>
  </div>
  <div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <div id="column1">
      <div class="box" id="dataexport">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.dataexport", LanguageID))%>
          </span>
        </h2>
        <%
          If rst1.Rows.Count > 0 Then
            For Each row In rst1.Rows
              ' split the date for use in separate UI form field controls.
              If (Date.TryParse(MyCommon.NZ(row.Item("StartDateTime"), ""), ScheduleDate)) Then
                StartDay = Logix.ToShortDateString(ScheduleDate, MyCommon)
                tod = ScheduleDate.ToString("HH:mm")
                If (tod <> "") Then
                  Dim tokens As String() = tod.Split(":")
                  If (tokens.Length = 2) Then
                    todhour = tokens(0)
                    todminute = tokens(1)
                  End If
                End If
              Else
                StartDay = ""
                todhour = ""
                todminute = ""
              End If
              
              If (IsDBNull(row.Item("LastRunTime"))) Then
                LastRan = Copient.PhraseLib.Lookup("term.never", LanguageID)
              Else
                LastRan = Logix.ToShortDateTimeString(row.Item("LastRunTime"), MyCommon)
              End If
              
              ScheduleEnabled = MyCommon.NZ(row.Item("Enabled"), False)
              ScheduleID = MyCommon.NZ(row.Item("ScheduleID"), 0)
              OutputZipped = MyCommon.NZ(row.Item("Zipped"), False)
              
              Send("<div style=""position:relative;"">")
              Send("<input type=""hidden"" id=""scheduleid"" name=""scheduleid"" value=""" & MyCommon.NZ(row.Item("ScheduleID"), -1) & """ />")
              Send("<input type=""hidden"" id=""deo-id"" name=""deo-id"" value=""" & MyCommon.NZ(row.Item("DataExportOptionID"), -1) & """ />")
              Send("<label for=""deo-enable""><b>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "</b></label><br />")
              Send("<br class=""half"" />")
              Send("<div id=""deo" & MyCommon.NZ(row.Item("DataExportOptionID"), 0) & """ style=""margin:0 0 5px 15px;"">")
              Send("  <input type=""checkbox"" id=""deo-enable"" name=""deo-enable"" value=""1""" & IIf(ScheduleEnabled, " checked=""checked"" ", " ") & "/><label for=""deo-enable"">" & Copient.PhraseLib.Lookup("term.enabled", LanguageID) & "</label><br />")
              Send("  <br />")
			  If MyCommon.Fetch_SystemOption("288") = 0 Then
                 Send("  <label for=""deo-path"">" & Copient.PhraseLib.Lookup("term.path", LanguageID) & ":</label><br />")
                 Send("  <input type=""text"" class=""longer"" id=""deo-path"" onkeypress=""return RestrictSpace()"" name=""deo-path"" value=""" & MyCommon.NZ(row.Item("OutputPath"), "") & """ maxlength=""100"" /><br />")
                 Send("  <br class=""half"" />")
			  End If
              Send("  <label for=""deo-zipped"">" & Copient.PhraseLib.Lookup("term.usegzip", LanguageID) & ":</label><input type=""checkbox"" id=""deo-zipped"" name=""deo-zipped"" value=""1""" & IIf(OutputZipped, " checked=""checked"" ", " ") & "/><br />")
              Send("  <br /><br class=""half"" />")
              Send("  <label for=""deo-start"">" & Copient.PhraseLib.Lookup("term.startdate", LanguageID) & ":</label><br />")
              Send("  <input type=""text"" class=""short"" id=""deo-start"" name=""deo-start"" value=""" & StartDay & """ />")
              Send("  <img src=""../images/calendar.png"" class=""calendar"" id=""deo-start-picker"" alt=""" & Copient.PhraseLib.Lookup("term.datepicker", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.datepicker", LanguageID) & """ onclick=""displayDatePicker('deo-start',event);"" /><br />")
              Send("  <br class=""half"" />")
              Send("  <label for=""hours"">" & Copient.PhraseLib.Lookup("term.time", LanguageID) & ":</label><br />")
              Send("  <select id=""hours"" name=""hours"">")
              For i = 0 To 23
                selectedStr = IIf(i = Integer.Parse(todhour), " selected=""selected""", "")
                Send("      <option value=""" & i.ToString("00") & """" & selectedStr & ">" & i.ToString("00") & "</option>")
              Next
              Send("  </select>:<select id=""minutes"" name=""minutes"">")
              For i = 0 To 59
                selectedStr = IIf(i = Integer.Parse(todminute), " selected=""selected""", "")
                Send("    <option value=""" & i.ToString("00") & """" & selectedStr & ">" & i.ToString("00") & "</option>")
              Next
              Send("  </select><br />")
              Send("  <br class=""half"" />")
              
              ' get the selected days for the schedule.  Note: the zero element is unused purposefully to have day value
              ' to correspond directly with the bit array index 
              MyCommon.QueryStr = "select value from DataExportScheduleValues where ScheduleID =" & ScheduleID
              rst3 = MyCommon.LRT_Select()
              For Each row3 In rst3.Rows
                DayIndex = MyCommon.NZ(row3.Item("value"), 0)
                If (DayIndex > 0 AndAlso DayIndex < DaysChecked.Count) Then
                  DaysChecked(DayIndex) = True
                End If
              Next
              
              Send("  " & Copient.PhraseLib.Lookup("term.day", LanguageID) & ":<br />")
              Send("  <input type=""checkbox"" id=""day1"" name=""day"" value=""1"" " & IIf(DaysChecked(1), " checked=""checked""", " ") & " /><label for=""day1"">" & Left(Copient.PhraseLib.Lookup("term.sunday", LanguageID), 3) & "</label> ")
              Send("  <input type=""checkbox"" id=""day2"" name=""day"" value=""2"" " & IIf(DaysChecked(2), " checked=""checked""", " ") & " /><label for=""day2"">" & Left(Copient.PhraseLib.Lookup("term.monday", LanguageID), 3) & "</label> ")
              Send("  <input type=""checkbox"" id=""day3"" name=""day"" value=""3"" " & IIf(DaysChecked(3), " checked=""checked""", " ") & " /><label for=""day3"">" & Left(Copient.PhraseLib.Lookup("term.tuesday", LanguageID), 3) & "</label> ")
              Send("  <input type=""checkbox"" id=""day4"" name=""day"" value=""4"" " & IIf(DaysChecked(4), " checked=""checked""", " ") & " /><label for=""day4"">" & Left(Copient.PhraseLib.Lookup("term.wednesday", LanguageID), 3) & "</label> ")
              Send("  <input type=""checkbox"" id=""day5"" name=""day"" value=""5"" " & IIf(DaysChecked(5), " checked=""checked""", " ") & " /><label for=""day5"">" & Left(Copient.PhraseLib.Lookup("term.thursday", LanguageID), 3) & "</label> ")
              Send("  <input type=""checkbox"" id=""day6"" name=""day"" value=""6"" " & IIf(DaysChecked(6), " checked=""checked""", " ") & " /><label for=""day6"">" & Left(Copient.PhraseLib.Lookup("term.friday", LanguageID), 3) & "</label> ")
              Send("  <input type=""checkbox"" id=""day7"" name=""day"" value=""7"" " & IIf(DaysChecked(7), " checked=""checked""", " ") & " /><label for=""day7"">" & Left(Copient.PhraseLib.Lookup("term.saturday", LanguageID), 3) & "</label>")
              Send("  <br />")
              Send("  <br clear=""left"" />")
              Send("  <br class=""half"" />")
              Send("  " & Copient.PhraseLib.Lookup("term.lastran", LanguageID) & " " & LastRan & "<br />")
              Send("  <br class=""half"" />")
              Send("</div>")
              Send("</div")
            Next
          Else
            Send(Copient.PhraseLib.Lookup("dataexports.no-deo", LanguageID) & "<br />")
          End If
        %>
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
    
    <div id="gutter">
    </div>
    
    <% If rst2.Rows.Count > 0 OrElse (IssuanceEnabled AndAlso rstIssuance.Rows.Count > 0) Then%>
    <div id="column2">
      <div class="box" id="remotedata">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.remotedata", LanguageID))%>
          </span>
        </h2>
        <%
            If rst2.Rows.Count > 0 Then
                If rst2.Rows(0).Item("TypeDescription").ToString().Contains("UE") Then
                    Send("<b>" & Copient.PhraseLib.Lookup(MyCommon.NZ(rst2.Rows(0).Item("PhraseID"), 0), LanguageID) & "</b><br />")
                End If
                For Each row In rst2.Rows
                    RemotePKID = MyCommon.NZ(row.Item("PKID"), 0)
                    RemoteDataType = MyCommon.NZ(row.Item("RemoteDataTypeID"), 0)
                    RemoteEnabled = MyCommon.NZ(row.Item("Enabled"), False)
                    RemotePath = MyCommon.NZ(row.Item("OutputPath"), "")
                    Dim RmtStyleDesc As String = MyCommon.NZ(row.Item("StyleDescription"), Copient.PhraseLib.Lookup("term.enabled", LanguageID))
                    If (IsDBNull(row.Item("LastUpdate"))) Then
                        LastUpdate = Copient.PhraseLib.Lookup("term.never", LanguageID)
                    Else
                        LastUpdate = row.Item("LastUpdate").ToString
                    End If
                    FilePathRequired = MyCommon.NZ(row.Item("FilePathRequired"), False)
                    If Not row.Item("TypeDescription").ToString().Contains("UE") Then
                        Send("<b>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "</b><br />")
                    End If
                    Send("<br class=""half"" />")
                    Send("<div id=""rdo" & RemotePKID & """ style=""margin:0 0 5px 15px;"">")
                    Send("  <input type=""hidden"" name=""remotepkid"" id=""remotepkid" & RemotePKID & """ value=""" & RemotePKID & """ />")
                    Send("  <input type=""hidden"" name=""rdo-enable"" id=""rdo-enable" & RemotePKID & """ value=""" & IIf(RemoteEnabled, "1", "0") & """ />")
                    Send("  <input type=""checkbox"" id=""rdo-chk-enable" & RemotePKID & """ name=""rdo-chk-enable"" value=""1""" & IIf(RemoteEnabled, " checked=""checked""", "") & " onclick=""javascript:updateRdoEnabled(" & RemotePKID & ", this.checked);"" />")
                    If row.Item("TypeDescription").ToString().Contains("UE") Then
                        Send("  <label for=""rdo-chk-enable" & RemotePKID & """>" & RmtStyleDesc & "</label><br />")
                    Else
                        Send("  <label for=""rdo-chk-enable" & RemotePKID & """>" & Copient.PhraseLib.Lookup("term.enabled", LanguageID) & "</label><br />")
                    End If


                    Send("  <br class=""half"" />")
                    If MyCommon.Fetch_SystemOption("288") = 0 Then
                        Send("  <label for=""rdo-path" & RemotePKID & """>" & Copient.PhraseLib.Lookup("term.path", LanguageID) & ":</label><br />")
                        Send("  <input type=""text"" class=""longer"" id=""rdo-path" & RemotePKID & """ name=""rdo-path"" onkeypress=""return RestrictSpace()"" value=""" & RemotePath & """ maxlength=""100"" /><br />")
                        Send("  <br class=""half"" />")
                    End If
                    Send("  " & Copient.PhraseLib.Lookup("term.lastupdated", LanguageID) & " " & LastUpdate & "<br />")
                    Send("  <br class=""half"" />")
                    Send("</div>")
                Next
            End If

            If IssuanceEnabled Then
                ' remote data option for issuance only
                MyCommon.QueryStr = "select RDO.PKID, RDO.RemoteDataTypeID, RDT.Description as TypeDescription, RDO.StyleID, RDS.Description as StyleDescription, " & _
                                    "RDO.Enabled, RDO.OutputPath, RDO.LastUpdate, RDS.FilePathRequired " & _
                                    "from RemoteDataOptions as RDO " & _
                                    "inner join RemoteDataTypes as RDT on RDT.RemoteDataTypeID=RDO.RemoteDataTypeID " & _
                                    "inner join RemoteDataStyles as RDS on RDS.StyleID=RDO.StyleID and RDS.RemoteDataTypeID = RDO.RemoteDataTypeID " & _
                                    "where RDO.RemoteDataTypeID = 1 order by RemoteDataTypeID desc;"
                rstIssuance = MyCommon.LRT_Select
                If rstIssuance.Rows.Count > 0 Then
                    i = 0
                    For Each row In rstIssuance.Rows
                        RemotePKID = MyCommon.NZ(row.Item("PKID"), 0)
                        RemoteDataType = MyCommon.NZ(row.Item("RemoteDataTypeID"), 0)
                        RemoteEnabled = MyCommon.NZ(row.Item("Enabled"), False)
                        RemotePath = MyCommon.NZ(row.Item("OutputPath"), "")
                        If (IsDBNull(row.Item("LastUpdate"))) Then
                            LastUpdate = Copient.PhraseLib.Lookup("term.never", LanguageID)
                        Else
                            LastUpdate = row.Item("LastUpdate").ToString
                        End If
                        FilePathRequired = MyCommon.NZ(row.Item("FilePathRequired"), False)
                        RemoteStyleDesc = MyCommon.NZ(row.Item("StyleDescription"), Copient.PhraseLib.Lookup("term.enabled", LanguageID))

                        If (i = 0) Then
                            Send("<b>" & MyCommon.NZ(row.Item("TypeDescription"), "") & "</b><br />")
                            Send("<br class=""half"" />")
                            Send("<div id=""rdo" & RemotePKID & """ style=""margin:0 0 5px 15px;"">")
                        End If

                        Send("  <input type=""hidden"" name=""remotepkid"" id=""remotepkid" & RemotePKID & """ value=""" & RemotePKID & """ />")
                        Send("  <input type=""hidden"" name=""rdo-enable"" id=""rdo-enable" & RemotePKID & """ value=""" & IIf(RemoteEnabled, "1", "0") & """ />")
                        Send("  <input type=""checkbox"" id=""rdo-chk-enable" & RemotePKID & """ name=""rdo-chk-enable"" value=""1""" & IIf(RemoteEnabled, " checked=""checked""", "") & " onclick=""javascript:updateRdoEnabled(" & RemotePKID & ", this.checked);"" />")
                        Send("  <label for=""rdo-chk-enable" & RemotePKID & """>" & RemoteStyleDesc & "</label><br />")
                        Send("  <input type=""hidden"" class=""longer"" id=""rdo-path" & RemotePKID & """ name=""rdo-path"" value=""" & RemotePath & """ />")
                        Send("  <br class=""half"" />")

                        If (i = rstIssuance.Rows.Count - 1) Then
                            Send("</div>")
                        End If

                        i += 1
                    Next
                End If
            End If
            If rst2.Rows.Count = 0 AndAlso rstIssuance.Rows.Count = 0 Then
                Send(Copient.PhraseLib.Lookup("dataexports.no-rdo", LanguageID) & "<br />")
            End If
        %>
        
        <hr class="hidden" />
      </div>
    </div>
    <% End If %>
    <br clear="all" />
  </div>
</form>

<script type="text/javascript">
  <% Send_Date_Picker_Terms() %>
</script>

<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (Logix.UserRoles.AccessNotes) Then
      'Send_Notes(21, 0, AdminUserID)
    End If
  End If
done:
  Send_BodyEnd("mainform", "name")
  MyCommon.Close_LogixRT()
  MyCommon = Nothing
  Logix = Nothing
%>