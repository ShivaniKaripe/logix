<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Text.RegularExpressions" %>
<%
  ' *****************************************************************************
  ' * FILENAME: connector-detail.aspx 
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
  Dim ExtractID As Integer
  Dim ExtractName As String
  Dim iPhraseId As Integer
  Dim iFrequency As Integer
  Dim iDayOfWeek As Integer
  Dim iStartHr As Integer
  Dim iStartMin As Integer
  Dim bEnabled As Integer = True
  Dim bRunImmediate as Boolean = False
  Dim iNumDays as Integer
  Dim bAllowEdit As Boolean = True
  Dim bWeekly As Boolean = False
  Dim sFilePath As String = ""
  Dim sLastRunStart As String = ""
  Dim sLastRunFinish As String = ""
  
  Dim Path As String = ""
  Dim Installed As Boolean = True
  Dim Visible As Boolean = True
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim OptSelected As String = ""
  Dim TempStr As String = ""
  Dim OptionObj As Copient.SystemOption = Nothing
  Dim CreatedDate As String = ""
  Dim bShelfLabelEnabled As Boolean = False

  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  
  Response.Expires = 0
  MyCommon.AppName = "CM-extract-detail.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  If MyCommon.Fetch_CM_SystemOption(105) = "1" Then
    bShelfLabelEnabled = True
  Else
    bShelfLabelEnabled = False
  End If

  If (Request.QueryString("ExtractID") <> "") Then
    ExtractID = MyCommon.Extract_Val(Request.QueryString("ExtractID"))
  ElseIf (Request.QueryString("form_ExtractID") <> "") Then
    ExtractID = MyCommon.Extract_Val(Request.QueryString("form_ExtractID"))
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
    If (Request.QueryString("RunImmediateExtract") <> "") Then
      If sUpdate <> "" Then
        sUpdate &= ","
      End If
      If MyCommon.NZ(Request.QueryString("RunImmediateExtract"), "") = "on" Then
        sUpdate &= "RunImmediate=1"
      Else
        sUpdate &= "RunImmediate=0"
      End If
    Else
      If sUpdate <> "" Then
        sUpdate &= ","
      End If
      sUpdate &= "RunImmediate=0"
    End If
    If (Request.QueryString("NumDays") <> "") Then
      iNumDays = MyCommon.NZ(Request.QueryString("NumDays"), 0)
      If sUpdate <> "" Then
        sUpdate &= ","
      End If
      sUpdate &= "NumDays=" & iNumDays
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
        MyCommon.QueryStr = "update CM_Extract_Options with (RowLock) set " & sUpdate & _
                            " where ExtractID=" & ExtractID & ";"
        MyCommon.LRT_Execute()
        If (MyCommon.Fetch_SystemOption(48) = "1") Then
          Response.Redirect("CM-extracts.aspx")
        End If
      End If
    End If
  End If
  
  If (Request.QueryString("infoMessage") <> "") Then
    infoMessage = Request.QueryString("infoMessage")
  End If
   
  If ExtractID > 0 Then
    MyCommon.QueryStr = "select ExtractId, ExtractName, PhraseID, Display, AllowEdit, LastUpdate, Frequency, FilePath, Enabled, LastRunStart, LastRunFinish," & _
                        " ScheduledRunDay, ScheduledRunHour, ScheduledRunMinute, RunImmediate, NumDays" & _
                        " from CM_Extract_Options with (NoLock) where ExtractID=" & ExtractID & ";"
  
    dt = MyCommon.LRT_Select
    If dt.Rows.Count > 0 Then
      dr = dt.Rows(0)
      ExtractName = dr.Item("ExtractName")
      If (Not IsDBNull(dr.Item("PhraseID"))) Then
        If Not Integer.TryParse(dr.Item("PhraseID"), iPhraseId) Then iPhraseId = 0
        If iPhraseId > 0 Then
          ExtractName = Copient.PhraseLib.Lookup(iPhraseId, LanguageID, dr.Item("ExtractName"))
        End If
      End If
      If (IsDBNull(dr.Item("Enabled"))) Then
        bEnabled = True
      Else
        bEnabled = MyCommon.NZ(dr.Item("Enabled"), True)
      End If
      If (IsDBNull(dr.Item("RunImmediate"))) Then
        bRunImmediate = False
      Else
        bRunImmediate = MyCommon.NZ(dr.Item("RunImmediate"), False)
      End If
      If (IsDBNull(dr.Item("NumDays"))) Then
        iNumDays = 0
      Else
        If Not Integer.TryParse(dr.Item("NumDays"), iNumDays) Then iNumDays = 0
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
  
  Send_HeadBegin("term.extractoptions", , ExtractID)
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
  Send_Subtabs(Logix, 8, 4, 0, "", "CM-extracts.aspx")
  If (Logix.UserRoles.AccessCMExtracts = False) Then
    Send_Denied(1, "perm.accesscmextracts")
    GoTo done
  End If
%>

<script type="text/javascript">

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

</script>

<form action="CM-extract-detail.aspx" method="get" id="mainform" name="mainform"
  onsubmit="elmName(); return handleOnSubmit();">
  <div id="intro">
    <h1 id="title">
      <%
        Sendb(Copient.PhraseLib.Lookup("term.extractoptions", LanguageID))
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
    <input type="hidden" id="form_ExtractID" name="form_ExtractID" value="<%sendb(ExtractID) %>" />
    <%
      If (infoMessage <> "") Then
        Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
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
          Send(Copient.PhraseLib.Lookup("term.Name", LanguageID) & ": &nbsp;" & ExtractName)
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
        <input class="longer" id="FilePath" maxlength="250" name="FilePath" type="text" value="<% sendb(sFilePath)%>" />
        <br />
        <br />
        <input class="checkbox" id="DisableExtract" name="DisableExtract" type="checkbox"
          <% if(not bEnabled)then sendb(" checked=""checked""") %> />
        <label for="DisableExtract">
          <% Sendb(Copient.PhraseLib.Lookup("term.disable", LanguageID))%>
        </label>
        <br />
        <br />
        <%If bShelfLabelEnabled Then%>
          <input class="checkbox" id="RunImmediateExtract" name="RunImmediateExtract" type="checkbox"
            <% if(bRunImmediate)then sendb(" checked=""checked""") %> />
          <label for="RunImmediateExtract">
            <% Sendb(Copient.PhraseLib.Lookup("term.runimmediate", LanguageID))%>
          </label>
          <br />
          <br />
          <label for="NumDays">
            <% Sendb(Copient.PhraseLib.Lookup("term.numdays", LanguageID))%>
          </label>
          &nbsp;
          <input class="shortest" id="NumDays" maxlength="2" name="NumDays" type="text"
            value="<% sendb(iNumDays.ToString("0"))%>" />
        <%End If%>
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
        <% Sendb(Copient.PhraseLib.Lookup("cm-extracts.timetorun", LanguageID))%>
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
