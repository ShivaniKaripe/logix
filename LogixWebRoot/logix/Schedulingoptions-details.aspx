<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Text.RegularExpressions" %>
<%@ Import Namespace="System.IO" %>
<%
  ' *****************************************************************************
  ' * FILENAME: Schedulingoptions-details.aspx 
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
  Dim dr As System.Data.DataRow
  Dim AppID As Integer
    Dim Name As String
    Dim iPhraseId As Integer
  Dim iFrequency As Integer
  Dim iDayOfWeek As Integer
  Dim iStartHr As Integer
  Dim iStartMin As Integer
  Dim bEnabled As Integer = True
  Dim bAllowEdit As Boolean = True
  Dim bWeekly As Boolean = False
  Dim sFilePath As String = ""
  Dim sLastRunStart As String = ""
  Dim sLastRunFinish As String = ""
  
  Dim Path As String = ""
  Dim Installed As Boolean = True
  Dim Visible As Boolean = True
  Dim infoMessage As String = ""
  Dim infoMessageGen As String = ""
  Dim Handheld As Boolean = False
  Dim OptSelected As String = ""
  Dim TempStr As String = ""
  Dim OptionObj As Copient.SystemOption = Nothing
  Dim CreatedDate As String = ""  
  Dim TotalTransactions As Integer

  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  
  Response.Expires = 0
  MyCommon.AppName = "Schedulingoptions-details.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  If (Request.QueryString("AppID") <> "") Then
    AppID = MyCommon.Extract_Val(Request.QueryString("AppID"))
  ElseIf (Request.QueryString("form_AppID") <> "") Then
    AppID = MyCommon.Extract_Val(Request.QueryString("form_AppID"))
  End If
  
  If (Request.QueryString("save") <> "") Then
    Dim sUpdate As String = ""
    If (Request.QueryString("dowradio") <> "") Then
      iDayOfWeek = MyCommon.NZ(Request.QueryString("dowradio"), 0)
      If sUpdate <> "" Then
        sUpdate &= ","
      End If
      sUpdate &= "ScheduledRunDay=" & iDayOfWeek
    End If
    If (Request.QueryString("DisableExtract") <> "") Then
      If sUpdate <> "" Then
        sUpdate &= ","
      End If
      If MyCommon.NZ(Request.QueryString("DisableExtract"), "") = "on" Then
        sUpdate &= "Enabled=0"
      Else
        sUpdate &= "Enabled=1"
      End If
    Else
      If sUpdate <> "" Then
        sUpdate &= ","
      End If
      sUpdate &= "Enabled=1"
    End If
    If (Request.QueryString("FilePath") <> "") Then
      sFilePath = MyCommon.NZ(Request.QueryString("FilePath"), "")
      If sUpdate <> "" Then
        sUpdate &= ","
      End If
      sUpdate &= "FilePath='" & sFilePath & "'"
    End If
    If (Request.QueryString("StartHour") <> "") Then
      iStartHr = MyCommon.NZ(Request.QueryString("StartHour"), 0)
      If sUpdate <> "" Then
        sUpdate &= ","
      End If
      sUpdate &= "ScheduledRunHour=" & iStartHr
    End If
    If (Request.QueryString("StartMin") <> "") Then
      iStartMin = MyCommon.NZ(Request.QueryString("StartMin"), 0)
      If sUpdate <> "" Then
        sUpdate &= ","
      End If
      sUpdate &= "ScheduledRunMinute=" & iStartMin
    End If
    If sUpdate <> "" Then
      If iStartHr < 0 Or iStartHr > 23 Then
        infoMessage = Copient.PhraseLib.Lookup("PDEoffer-gen.invalidtime", LanguageID)
      End If
      If iStartMin < 0 Or iStartMin > 59 Then
        infoMessage = Copient.PhraseLib.Lookup("PDEoffer-gen.invalidtime", LanguageID)
      End If
      If infoMessage = "" Then
        sUpdate &= ",LastUpdate=getdate()"
        MyCommon.QueryStr = "update Agent_Scheduling_Options with (RowLock) set " & sUpdate & _
                            " where AppID=" & AppID & ";"
        MyCommon.LRT_Execute()
        If (MyCommon.Fetch_SystemOption(48) = "1") Then
          Response.Redirect("Agent-Schedulingoptions.aspx")
        End If
      End If
    End If
  End If
  
  If (Request.QueryString("infoMessage") <> "") Then
    infoMessage = Request.QueryString("infoMessage")
  End If
   
  If AppID > 0 Then
    MyCommon.QueryStr = "select AppID, Name, PhraseID, Display, AllowEdit, LastUpdate, Frequency, FilePath, Enabled, LastRunStart, LastRunFinish," & _
                        " ScheduledRunDay, ScheduledRunHour, ScheduledRunMinute" & _
                        " from Agent_Scheduling_Options with (NoLock) where AppID=" & AppID & ";"
  
    dt = MyCommon.LRT_Select
    If dt.Rows.Count > 0 Then
      dr = dt.Rows(0)
      Name = dr.Item("Name")
      If (Not IsDBNull(dr.Item("PhraseID"))) Then
        If Not Integer.TryParse(dr.Item("PhraseID"), iPhraseId) Then iPhraseId = 0
        If iPhraseId > 0 Then
          Name = Copient.PhraseLib.Lookup(iPhraseId, LanguageID, dr.Item("Name"))
        End If
      End If
      If (IsDBNull(dr.Item("Enabled"))) Then
        bEnabled = True
      Else
        bEnabled = MyCommon.NZ(dr.Item("Enabled"), True)
      End If
      If (IsDBNull(dr.Item("AllowEdit"))) Then
        bAllowEdit = True
      Else
        bAllowEdit = MyCommon.NZ(dr.Item("AllowEdit"), True)
      End If
      If (IsDBNull(dr.Item("FilePath"))) Then
        sFilePath = ""
      Else
        sFilePath = MyCommon.NZ(dr.Item("FilePath"), "")
      End If
      If (IsDBNull(dr.Item("Frequency"))) Then
        iFrequency = 0
      Else
        If Not Integer.TryParse(dr.Item("Frequency"), iFrequency) Then iFrequency = 0
        If iFrequency = 2 Then
          bWeekly = True
          If IsDBNull(dr.Item("ScheduledRunDay")) Then
            iDayOfWeek = 6
          Else
            If Not Integer.TryParse(dr.Item("ScheduledRunDay"), iDayOfWeek) Then iDayOfWeek = 6
          End If
        End If
      End If
      If (IsDBNull(dr.Item("LastRunStart"))) Then
        sLastRunStart = Copient.PhraseLib.Lookup("term.unknown", LanguageID)
      Else
        sLastRunStart = Logix.ToShortDateTimeString(dr.Item("LastRunStart"), MyCommon)
      End If
      If (IsDBNull(dr.Item("LastRunFinish"))) Then
        sLastRunFinish = Copient.PhraseLib.Lookup("term.unknown", LanguageID)
      Else
        sLastRunFinish = Logix.ToShortDateTimeString(dr.Item("LastRunFinish"), MyCommon)
      End If
      If (IsDBNull(dr.Item("ScheduledRunHour"))) Then
        iStartHr = 0
      Else
        If Not Integer.TryParse(dr.Item("ScheduledRunHour"), iStartHr) Then iStartHr = 0
      End If
      If (IsDBNull(dr.Item("ScheduledRunMinute"))) Then
        iStartMin = 0
      Else
        If Not Integer.TryParse(dr.Item("ScheduledRunMinute"), iStartMin) Then iStartMin = 0
      End If
    End If
  End If
  
  'Generate Reconcilation Report
  'Save
  If (Request.QueryString("Generate") <> "") Then
    'form_Generate = Request.QueryString("startdate")
	If Request.QueryString("startdate") <> "" Then
	  TotalTransactions = ExportFile(Request.QueryString("startdate"), infoMessage)
	  If infoMessage <> "" Then
	    infoMessage = infoMessage
	  ElseIf TotalTransactions > 0 Then
	    infoMessageGen = Copient.PhraseLib.Lookup("Schedulingoptions-details.recongen", LanguageID)		
	  Else 
	    infoMessage = Copient.PhraseLib.Lookup("Schedulingoptions-details.norecords", LanguageID)	
	  End If
	Else
	  	infoMessage = Copient.PhraseLib.Lookup("Schedulingoptions-details.nodate", LanguageID)
	End If      
  End If
  
  
  
  Send_HeadBegin("term.schedulingoptions")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
  Send_HeadEnd()
  Send_BodyBegin(1)
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  Send_Tabs(Logix, 8)
  Send_Subtabs(Logix, 8, 4, 0, "", "Agent-Schedulingoptions.aspx")
    If (Logix.UserRoles.AccessSchedulingOptions = False) Then
        Send_Denied(1, "perm.accessschedulingoptions")
        GoTo done
    End If
	Send_Scripts(New String() {"datePicker.js", "popup.js"})
%>

<script runat="server">
    Public MyCommon As New Copient.CommonInc
    Dim MyCryptLib As New Copient.CryptLib
  Public Structure HeaderRecord
    
    Dim CreateTimeStamp As Date
    Dim FileId As String
    Dim NoOfTransactions As Integer
    Dim TotalDiscount As Decimal
  End Structure
  Public Structure TransactionRecords
    
    Dim ArrangementID As String
    Dim LoyaltyCardNumber As String
    Dim RedeemedQuantity As Integer
    Dim SVProgramID As Integer
    Dim CouponID As String
    Dim SiteID As String
    Dim DiscountAmount As Decimal
    Dim POSTimeStamp As String
  End Structure
  Public oHeaderRecord As HeaderRecord
  
  Public otransRecoerd As TransactionRecords
  
  Public dtThirdPartyTransactions As DataTable
  
  Function ExportFile(ByVal ReconiliationFileDay As Date, ByRef infoMessage As String) As Integer
            
    Dim Pathstr As String = getExportFilePath()
    
    Dim ReconciliationFile As String = Pathstr & "ShellReconciliation." & ReconiliationFileDay.ToString("MMddyyyy") & ".txt"
		
    Dim TotalTransactions As Integer
	
	  Dim objFS As FileStream
      Dim objSW As StreamWriter
	  
      Dim dt As DataTable
		
		
    Try
      If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
      If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

      Dim TransTimeLBound As DateTime = ReconiliationFileDay.AddHours(-21)
      Dim TransTimeUBound As DateTime = ReconiliationFileDay.AddHours(3).AddMinutes(-1)
	  Dim sDateTimeFormat As String = "yyyy-MM-ddTHH:mm:ss"
			
      MyCommon.QueryStr = "SELECT Sum(DiscountAmount) AS TotalDiscountAmount FROM ThirdPartyTransactions WHERE ProcessTransaction = 1 AND POSTimeStamp BETWEEN '" & TransTimeLBound & "' AND '" & TransTimeUBound & "' Group by ProcessTransaction"
      dt = MyCommon.LRT_Select
			
      If dt.Rows.Count > 0 Then
			    
        oHeaderRecord.TotalDiscount = dt.Rows(0).Item("TotalDiscountAmount")
				
        MyCommon.QueryStr = "SELECT ArrangementID, POSTimeStamp, SiteID, LoyaltyCardNumber, DiscountAmount, CouponID, SVProgramQuantity FROM ThirdPartyTransactions WHERE ProcessTransaction = 1 AND POSTimeStamp BETWEEN '" & TransTimeLBound & "' AND '" & TransTimeUBound & "' Order By POSTimeStamp Asc"
        dtThirdPartyTransactions = MyCommon.LRT_Select
				

        If dtThirdPartyTransactions.Rows.Count > 0 Then
            If Not System.IO.File.Exists(ReconciliationFile) Then
	          objFS = New FileStream(ReconciliationFile, FileMode.Create, FileAccess.Write)
		      objSW = New StreamWriter(objFS)
	        Else
	          infoMessage = Copient.PhraseLib.Lookup("Schedulingoptions-details.fileexists", LanguageID)
			  Exit Function
	        End If
          TotalTransactions = dtThirdPartyTransactions.Rows.Count
          'objSW.WriteLine("CreateTimeStamp|FileID|NumberOfdataRecords|DiscountTotalAmount")
          objSW.WriteLine("" & Format(Date.Now, sDateTimeFormat) & "|ShellReconciliation" & Date.Now.ToString("MMddyyyyHHmmss") & ".TXT|" & TotalTransactions & "|" & oHeaderRecord.TotalDiscount & "")

          

          For Each row As System.Data.DataRow In dtThirdPartyTransactions.Rows
            otransRecoerd.ArrangementID = row.Item("ArrangementID")
            otransRecoerd.POSTimeStamp = Format(row.Item("POSTimeStamp"), sDateTimeFormat)
            otransRecoerd.SiteID = row.Item("SiteID")
            otransRecoerd.LoyaltyCardNumber = ReFormatCard(MyCryptLib.SQL_StringDecrypt(row.Item("LoyaltyCardNumber").ToString()))
            otransRecoerd.DiscountAmount = row.Item("DiscountAmount")
            otransRecoerd.CouponID = row.Item("CouponID")
            otransRecoerd.RedeemedQuantity = row.Item("SVProgramQuantity") * -1

            objSW.WriteLine("" & otransRecoerd.POSTimeStamp & "|" & otransRecoerd.SiteID & "|00000|" & otransRecoerd.LoyaltyCardNumber & "|" & otransRecoerd.ArrangementID & "|" & otransRecoerd.DiscountAmount & "|" & otransRecoerd.CouponID & "|" & otransRecoerd.RedeemedQuantity & "")
          Next

          objSW.Close()

        End If
      End If
    Catch ex As Exception
            
    Finally
      If MyCommon.LRTadoConn.State = ConnectionState.Open Then MyCommon.Close_LogixRT()
      If MyCommon.LXSadoConn.State = ConnectionState.Open Then MyCommon.Close_LogixXS()
    End Try
    Return TotalTransactions
		
  End Function
	
	Public Function getExportFilePath() As String
        Const WORKSPACE_FILE_PATH_SYSTEM_OPTION As Integer = 29

        Dim FilePath As String = MyCommon.Fetch_SystemOption(WORKSPACE_FILE_PATH_SYSTEM_OPTION).Trim

        If FilePath <> "" AndAlso Not (Right(FilePath, 1) = "\") Then
            FilePath = FilePath & "\"
        ElseIf FilePath = "" Then
            FilePath = "\"
        End If
        Return FilePath

    End Function
    Public Function ReFormatCard(ByVal CardID As String) As String
        Dim FormattedCard As String = ""
        If CardID.Length > 10 Then
            FormattedCard = CardID.Substring(CardID.Length - 10)
        End If
        Return FormattedCard
    End Function
</script>
 
<script type="text/javascript">

  var datePickerDivID = "datepicker";
  
  if (window.captureEvents) {
    window.captureEvents(Event.CLICK);
    window.onclick = handlePageClick;
  } else {
    document.onclick = handlePageClick;
  }
  
  function disableUnload() {
    window.onunload = null;
  }
    
  // callback function for save changes on unload during navigate away
  function handleAutoFormSubmit() {
      elmName();
      if (ValidateTimes()) {
        document.mainform.submit();
      }
  }
    
  function elmName(){
    window.onunload = null;
    for(i=0; i<document.mainform.elements.length; i++)
    {
      document.mainform.elements[i].disabled=false;
    }
    return true;
  }

  function handleOnSubmit() {
    var retVal = false;
  
    retVal = ValidateTimes();
    return retVal;
  }
  
  function ValidateTimes() {
  var elemStartHr = document.getElementById("StartHour");
  var elemStartMin = document.getElementById("StartMin");
  var retVal = true;

  if (retVal == true && elemStartHr != null && (!isInteger(elemStartHr.value) || (parseInt(elemStartHr.value) < 0) || (parseInt(elemStartHr.value) > 23))) {
		alert('<%Sendb(Copient.PhraseLib.Lookup("logix-js.EnterStartHour", LanguageID)) %>');
    retVal = false;
  }
  if (retVal == true && elemStartMin != null && (!isInteger(elemStartMin.value) || (parseInt(elemStartMin.value) < 0) || (parseInt(elemStartMin.value) > 59))) {
		alert('<%Sendb(Copient.PhraseLib.Lookup("logix-js.EnterStartMinute", LanguageID)) %>');
    retVal = false;
  }
  return retVal;
}

function isIE() {
    return /msie/i.test(navigator.userAgent) && !/opera/i.test(navigator.userAgent);
  } 
       
       function handlePageClick(e) {
      var calFrame = document.getElementById('calendariframe');
      var el=(typeof event!=='undefined')? event.srcElement : e.target        
      
      if (el != null) {
        var pickerDiv = document.getElementById(datePickerDivID);
        if (pickerDiv != null && pickerDiv.style.visibility == "visible") {
          if (el.id!="start-picker") {
            if (!isDatePickerControl(el.className)) {
              pickerDiv.style.visibility = "hidden";
              pickerDiv.style.display = "none";  
              if (calFrame != null) {
                calFrame.style.visibility = 'hidden';
                calFrame.style.display = 'none';
              }
            }
          } else  {
              pickerDiv.style.visibility = "visible";            
              pickerDiv.style.display = "block";            
              if (calFrame != null) {
                calFrame.style.visibility = 'visible';
                calFrame.style.display = 'block';
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

<form action="Schedulingoptions-details.aspx" method="get" id="mainform" name="mainform"
  onsubmit="elmName(); return handleOnSubmit();">
  <div id="intro">
    <h1 id="title">
      <%
        Sendb(Copient.PhraseLib.Lookup("term.schedulingoptions", LanguageID))
      %>
    </h1>
    <div id="controls">
      <%
        If (Logix.UserRoles.EditSystemConfiguration = True And bAllowEdit) Then
          Send_Save()
        End If
      %>
    </div>
  </div>
  <div id="main">
    <input type="hidden" id="form_AppID" name="form_AppID" value="<%sendb(AppID) %>" />
    <%
      If (infoMessage <> "") Then
        Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
      End If
	  If (infoMessageGen <> "") Then
        Send("<div id=""modbar"">" & infoMessageGen & "</div>")
      End If
    %>
    <div id="column1">
      <div class="box" id="identification">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.general", LanguageID))%>
          </span>
        </h2>
        <%
          Send("<br />")
          Send(Copient.PhraseLib.Lookup("term.Name", LanguageID) & ": &nbsp;" & Name)
          Send("<br /><br />")

          Send(Copient.PhraseLib.Lookup("term.frequency", LanguageID) & ": &nbsp;")
          Select Case iFrequency
            Case 1
              Send(Copient.PhraseLib.Lookup("term.daily", LanguageID))
            Case 2
              Send(Copient.PhraseLib.Lookup("term.weekly", LanguageID))
            Case Else
              Send(Copient.PhraseLib.Lookup("term.unknown", LanguageID))
          End Select
          Send("<br /><br />")

          Send(Copient.PhraseLib.Lookup("term.lastrunstart", LanguageID) & ": &nbsp;")
          Send(sLastRunStart)
          Send("<br /><br />")

          Send(Copient.PhraseLib.Lookup("term.lastrunfinish", LanguageID) & ": &nbsp;")
          Send(sLastRunFinish)
          Send("<br /><br />")
        %>
        <label for="FilePath">
          <% Sendb(Copient.PhraseLib.Lookup("term.path", LanguageID) & ": &nbsp;")%>
        </label>
		<% 'For AppID 98 and 100, disabling the user input to the path. These two agents use the path defined in System OptionID # 194 %>
        <input class="longer" id="FilePath" maxlength="250" name="FilePath" type="text" value="<% sendb(sFilePath)%>" <% sendb(IIf(AppID=98 Or AppID=100, "disabled", "")) %> />
        <br />
        <br />
        <input class="checkbox" id="DisableExtract" name="DisableExtract" type="checkbox"
          <% if(not bEnabled)then sendb(" checked=""checked""") %> />
        <label for="DisableExtract">
          <% Sendb(Copient.PhraseLib.Lookup("term.disable", LanguageID))%>
        </label>
        <br />
        <br />
      </div>
    </div>
    <div id="gutter">
    </div>
    <div id="column2">
      <div class="box" id="Div1">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.time", LanguageID))%>
          </span>
        </h2>
        <% Sendb(Copient.PhraseLib.Lookup("term.timetorun", LanguageID))%>
        <br />
        <br />
        <label for="StartHour">
          <% Sendb(Copient.PhraseLib.Lookup("term.StartTime", LanguageID))%>
        </label>
        &nbsp;
        <input class="shortest" id="StartHour" maxlength="2" name="StartHour" type="text"
          value="<% sendb(iStartHr.ToString("00"))%>" />
        <% Sendb(":")%>
        <input class="shortest" id="StartMin" maxlength="2" name="StartMin" type="text" value="<% sendb(iStartMin.ToString("00"))%>" />
        <hr class="hidden" />
      </div>
      <%If bWeekly Then%>
      <div class="box" id="DayOfWeek">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.dayofweek", LanguageID))%>
          </span>
        </h2>
        &nbsp;
        <input class="radio" id="sunday" name="dowradio" value="1" type="radio" <% if(iDayOfWeek=1)then sendb(" checked=""checked""") %> />
        <label for="sunday">
          <% Sendb(Copient.PhraseLib.Lookup("term.sunday", LanguageID))%>
        </label>
        <br />
        &nbsp;
        <input class="radio" id="monday" name="dowradio" value="2" type="radio" <% if(iDayOfWeek=2)then sendb(" checked=""checked""") %> />
        <label for="monday">
          <% Sendb(Copient.PhraseLib.Lookup("term.monday", LanguageID))%>
        </label>
        <br />
        &nbsp;
        <input class="radio" id="tuesday" name="dowradio" value="3" type="radio" <% if(iDayOfWeek=3)then sendb(" checked=""checked""") %> />
        <label for="tuesday">
          <% Sendb(Copient.PhraseLib.Lookup("term.tuesday", LanguageID))%>
        </label>
        <br />
        &nbsp;
        <input class="radio" id="wednesday" name="dowradio" value="4" type="radio" <% if(iDayOfWeek=4)then sendb(" checked=""checked""") %> />
        <label for="wednesday">
          <% Sendb(Copient.PhraseLib.Lookup("term.wednesday", LanguageID))%>
        </label>
        <br />
        &nbsp;
        <input class="radio" id="thursday" name="dowradio" value="5" type="radio" <% if(iDayOfWeek=5)then sendb(" checked=""checked""") %> />
        <label for="thursday">
          <% Sendb(Copient.PhraseLib.Lookup("term.thursday", LanguageID))%>
        </label>
        <br />
        &nbsp;
        <input class="radio" id="friday" name="dowradio" value="6" type="radio" <% if(iDayOfWeek=6)then sendb(" checked=""checked""") %> />
        <label for="friday">
          <% Sendb(Copient.PhraseLib.Lookup("term.friday", LanguageID))%>
        </label>
        <br />
        &nbsp;
        <input class="radio" id="saturday" name="dowradio" value="7" type="radio" <% if(iDayOfWeek=7)then sendb(" checked=""checked""") %> />
        <label for="saturday">
          <% Sendb(Copient.PhraseLib.Lookup("term.saturday", LanguageID))%>
        </label>
        <br />
        <hr class="hidden" />
      </div>
      <%End If%>
	  <%If AppID = 96%>
	   <div class="box" id="Div1">
	  
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.genreconciliationfile", LanguageID))%>
          </span>
        </h2>
		<label for="Period">
		  <% Sendb(Copient.PhraseLib.Lookup("term.period", LanguageID))%>
		</label>
		<input type="text" class="short" id="startdate" name="startdate" maxlength="10" value="" />
        <img src="../images/calendar.png" class="calendar" id="start-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('startdate', this);" />		
		<input type="submit" id="Generate" name="Generate" style="width:80px;height:22px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.generate", LanguageID))%>"/>
        
		<%
        If Request.Browser.Type = "IE6" Then
          Send("<iframe src=""javascript:'';"" id=""calendariframe"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no"" style=""visibility:hidden;display:none;""></iframe>")
        End If
      %>
	  </div>
	 <%End If%> 
    </div>
  </div>
</form>

<script runat="server">
 
    
</script>

<%

  Send_BodyEnd("mainform", "name")
done:
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  Logix = Nothing
  MyCommon = Nothing
%>
