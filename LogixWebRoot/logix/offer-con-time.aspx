<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: offer-con-time.aspx 
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
  Dim rst As DataTable
  Dim row As DataRow
  Dim OfferID As Long
  Dim Name As String = ""
  Dim ConditionID As String
  Dim NumTiers As Integer
  Dim sunday, monday, tuesday, wednesday, thursday, friday, saturday As Integer
  Dim StartHour, EndHour, StartMinute, EndMinute As String
  Dim iStartHour, iEndHour, iStartMinute, iEndMinute As Integer
  Dim Disallow_Edit As Boolean = True
  Dim bUseTemplateLocks As Boolean
  Dim IsTemplate As Boolean = False
  Dim CloseAfterSave As Boolean = False
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim BannersEnabled As Boolean = False
  Dim bDisallowEditDay As Boolean = False
  Dim bDisallowEditTime As Boolean = False
  Dim sDisabled As String = ""

  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "offer-con-time.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  StartHour = "0"
  EndHour = "0"
  StartMinute = "0"
  EndMinute = "0"
  OfferID = Request.QueryString("OfferID")
  ConditionID = Request.QueryString("ConditionID")
  NumTiers = Request.QueryString("NumTiers")
  
  
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  
  ' fetch the name
  MyCommon.QueryStr = "Select Name,IsTemplate,FromTemplate from Offers with (NoLock) where OfferID=" & OfferID
  rst = MyCommon.LRT_Select
  If (rst.Rows.Count > 0) Then
    Name = rst.Rows(0).Item("Name")
    IsTemplate = MyCommon.NZ(rst.Rows(0).Item("IsTemplate"), False)
    If IsTemplate Then
      bUseTemplateLocks = False
    Else
      bUseTemplateLocks = MyCommon.NZ(rst.Rows(0).Item("FromTemplate"), False)
    End If
  End If
  
  If (IsTemplate Or bUseTemplateLocks) Then
    ' lets dig the permissions if its a template
    MyCommon.QueryStr = "select Disallow_Edit, DisallowEdit1, DisallowEdit2 from OfferConditions with (NoLock) where ConditionID=" & ConditionID
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      Disallow_Edit = MyCommon.NZ(rst.Rows(0).Item("Disallow_Edit"), True)
      bDisallowEditDay = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit1"), False)
      bDisallowEditTime = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit2"), False)
      If bUseTemplateLocks Then
        If Disallow_Edit Then
          bDisallowEditDay = True
          bDisallowEditTime = True
        Else
          Disallow_Edit = bDisallowEditDay And bDisallowEditTime
        End If
      End If
    End If
  End If
  
  Send_HeadBegin("term.offer", "term.timecondition", OfferID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
  Send_HeadEnd()
  If (IsTemplate) Then
    Send_BodyBegin(12)
  Else
    Send_BodyBegin(2)
  End If
  
  If (Logix.UserRoles.AccessOffers = False AndAlso Not IsTemplate) Then
    Send_Denied(2, "perm.offers-access")
    GoTo done
  End If
  If (Logix.UserRoles.AccessTemplates = False AndAlso IsTemplate) Then
    Send_Denied(2, "perm.offers-access-templates")
    GoTo done
  End If
  If (BannersEnabled AndAlso Not Logix.IsAccessibleOffer(AdminUserID, OfferID)) Then
    Send("<script type=""text/javascript"" language=""javascript"">")
    Send("  function ChangeParentDocument() { return true; } ")
    Send("</script>")
    Send_Denied(1, "banners.access-denied-offer")
    Send_BodyEnd()
    GoTo done
  End If
  
  Send("<script type=""text/javascript"">window.name=""offerConTime""</script>")
  
  If (Request.QueryString("save") <> "") Then

    If (Not (bUseTemplateLocks And bDisallowEditDay)) Or (Not (bUseTemplateLocks And bDisallowEditTime)) Then
      MyCommon.QueryStr = "select ConditionID,StartHour,StartMinute,EndHour,EndMinute,Sunday,Monday,Tuesday,Wednesday,Thursday,Friday,Saturday " & _
                          "from ConditionTimes with (NoLock) where ConditionID=" & ConditionID & ";"
      rst = MyCommon.LRT_Select
      If (rst.Rows.Count > 0) Then
        iStartHour = rst.Rows(0).Item("StartHour")
        iStartMinute = rst.Rows(0).Item("StartMinute")
        iEndHour = rst.Rows(0).Item("EndHour")
        iEndMinute = rst.Rows(0).Item("EndMinute")
        sunday = rst.Rows(0).Item("Sunday")
        monday = rst.Rows(0).Item("Monday")
        tuesday = rst.Rows(0).Item("Tuesday")
        wednesday = rst.Rows(0).Item("Wednesday")
        thursday = rst.Rows(0).Item("Thursday")
        friday = rst.Rows(0).Item("Friday")
        saturday = rst.Rows(0).Item("Saturday")
      Else
        iStartHour = 0
        iStartMinute = 0
        iEndHour = 0
        iEndMinute = 0
        sunday = 0
        monday = 0
        tuesday = 0
        wednesday = 0
        thursday = 0
        friday = 0
        saturday = 0
        MyCommon.QueryStr = "insert into ConditionTimes with (RowLock) (ConditionID,StartHour,StartMinute,EndHour,EndMinute,Sunday,Monday,Tuesday,Wednesday,Thursday,Friday,Saturday) " & _
                            "values(" & ConditionID & "," & iStartHour & "," & iStartMinute & "," & iEndHour & "," & iEndMinute & "," & _
                            sunday & "," & monday & "," & tuesday & "," & wednesday & "," & thursday & "," & friday & "," & saturday & ")"
        MyCommon.LRT_Execute()
      End If
      
      If Not (bUseTemplateLocks And bDisallowEditDay) Then
        sunday = 0
        monday = 0
        tuesday = 0
        wednesday = 0
        thursday = 0
        friday = 0
        saturday = 0
        If (Request.QueryString("avail-sun") = "on") Then sunday = 1
        If (Request.QueryString("avail-mon") = "on") Then monday = 1
        If (Request.QueryString("avail-tue") = "on") Then tuesday = 1
        If (Request.QueryString("avail-wed") = "on") Then wednesday = 1
        If (Request.QueryString("avail-thu") = "on") Then thursday = 1
        If (Request.QueryString("avail-fri") = "on") Then friday = 1
        If (Request.QueryString("avail-sat") = "on") Then saturday = 1
      End If
      
      If Not (IsNumeric(Request.QueryString("start_hour")) AndAlso IsNumeric(Request.QueryString("start_minute")) AndAlso IsNumeric(Request.QueryString("end_hour")) AndAlso IsNumeric(Request.QueryString("end_minute"))) Then
        infoMessage = "The specified time range is invalid."
      Else
        If Not (bUseTemplateLocks And bDisallowEditTime) Then
          iStartHour = Int(Request.QueryString("start_hour"))
          iStartMinute = Int(Request.QueryString("start_minute"))
          iEndHour = Int(Request.QueryString("end_hour"))
          iEndMinute = Int(Request.QueryString("end_minute"))
        End If
      End If
      
      If (iStartHour > iEndHour) Or (iStartHour = iEndHour And iStartMinute > iEndMinute) Then
        infoMessage = Copient.PhraseLib.Lookup("condition.badtime", LanguageID)
      Else
        If infoMessage = "" Then
          MyCommon.QueryStr = "update ConditionTimes with (RowLock) " & _
                              "set StartHour=" & iStartHour & _
                              ",StartMinute=" & iStartMinute & _
                              ",EndHour=" & iEndHour & _
                              ",EndMinute=" & iEndMinute & _
                              ",Sunday=" & sunday & _
                              ",Monday=" & monday & _
                              ",Tuesday=" & tuesday & _
                              ",Wednesday=" & wednesday & _
                              ",Thursday=" & thursday & _
                              ",Friday=" & friday & _
                              ",Saturday=" & saturday & _
                              " where ConditionID=" & ConditionID & ";"
          MyCommon.LRT_Execute()
          MyCommon.QueryStr = "update OfferConditions with (RowLock) set TCRMAStatusFlag=2,CRMAStatusFlag=2 where ConditionID=" & ConditionID
          MyCommon.LRT_Execute()
          MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where offerid=" & OfferID
          MyCommon.LRT_Execute()
          MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-time", LanguageID))
        End If
      End If
    End If
    
    If (Request.QueryString("IsTemplate") = "IsTemplate") Then
      ' time to update the status bits for the templates
      Dim form_Disallow_Edit As Integer = 0
      Dim iDisallowEditDay As Integer = 0
      Dim iDisallowEditTime As Integer = 0
      
      Disallow_Edit = False
      bDisallowEditDay = False
      bDisallowEditTime = False

      If (Request.QueryString("Disallow_Edit") = "on") Then
        form_Disallow_Edit = 1
        Disallow_Edit = True
      End If
      If (Request.QueryString("DisallowEditDay") = "on") Then
        iDisallowEditDay = 1
        bDisallowEditDay = True
      End If
      If (Request.QueryString("DisallowEditTime") = "on") Then
        iDisallowEditTime = 1
        bDisallowEditTime = True
      End If
      
      MyCommon.QueryStr = "update OfferConditions with (RowLock) set Disallow_Edit=" & form_Disallow_Edit & _
                          ",DisallowEdit1=" & iDisallowEditDay & _
                          ",DisallowEdit2=" & iDisallowEditTime & _
                          " where ConditionID=" & ConditionID
      MyCommon.LRT_Execute()
    End If
    
    If infoMessage = "" Then
      CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
    End If
  End If
  
  Send("<script type=""text/javascript"">")
    Send("function ChangeParentDocument() { ")
    Send("  if (opener != null) {")
    Send("    var newlocation = 'offer-con.aspx?OfferID=" & OfferID & "'; ")
    Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
  Send("    opener.location = 'offer-con.aspx?OfferID=" & OfferID & "'; ")
    Send("    } ")
    Send("  }")
    Send("}")
  Send("</script>")
%>
<form action="offer-con-time.aspx" id="mainform" name="mainform" method="get" onsubmit="return ValidateOfferConTimeForm()">
  <div id="intro">
    <input type="hidden" id="OfferID" name="OfferID" value="<% sendb(OfferID) %>" />
    <input type="hidden" id="Name" name="Name" value="<% sendb(Name) %>" />
    <input type="hidden" id="ConditionID" name="ConditionID" value="<% sendb(ConditionID) %>" />
    <input type="hidden" id="IsTemplate" name="IsTemplate" value="<% 
            if(istemplate)then 
            sendb("IsTemplate")
            else
            sendb("Not") 
            end if
        %>" />
    <%If (isTemplate) Then
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.timecondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
      Else
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.timecondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
      End If
    %>
    <div id="controls">
      <% If (IsTemplate) Then%>
      <span class="temp">
        <input type="checkbox" class="tempcheck" id="temp-employees" name="Disallow_Edit"<% if(disallow_edit)then sendb(" checked=""checked""") %> />
        <label for="temp-employees"><% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
      </span>
      <% End If%>
      <% If Not (IsTemplate) Then
           If (Logix.UserRoles.EditOffer And Not (bUseTemplateLocks And Disallow_Edit)) Then Send_Save()
         Else
           If (Logix.UserRoles.EditTemplates) Then Send_Save()
         End If    
      %>
    </div>
  </div>
  <div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <div id="column1">
      <div class="box" id="day">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.day", LanguageID))%>
          </span>
          <% If (IsTemplate) Then%>
          <span class="tempRequire">
            <input type="checkbox" class="tempcheck" id="DisallowEditDay1" name="DisallowEditDay" <% if(bDisallowEditDay)then send(" checked=""checked""") %> />
            <label for="temp-Tiers"><% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
          </span>
          <% ElseIf (bUseTemplateLocks And bDisallowEditDay) Then%>
          <span class="tempRequire">
            <input type="checkbox" class="tempcheck" id="DisallowEditDay2" name="DisallowEditDay" disabled="disabled" checked="checked" />
            <label for="temp-Tiers"><% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
          </span>
          <% End If%>
        </h2>
        <% Sendb(Copient.PhraseLib.Lookup("condition.day", LanguageID))%>
        <br />
        <br class="half" />
        <%
          MyCommon.QueryStr = "select * from ConditionTimes with (NoLock) where ConditionID=" & ConditionID & ";"
          rst = MyCommon.LRT_Select
          For Each row In rst.Rows
            sunday = 0
            monday = 0
            tuesday = 0
            wednesday = 0
            thursday = 0
            friday = 0
            saturday = 0
            If (row.Item("Sunday")) Then sunday = 1
            If (row.Item("Monday")) Then monday = 1
            If (row.Item("Tuesday")) Then tuesday = 1
            If (row.Item("Wednesday")) Then wednesday = 1
            If (row.Item("Thursday")) Then thursday = 1
            If (row.Item("Friday")) Then friday = 1
            If (row.Item("Saturday")) Then saturday = 1
            StartHour = row.Item("StartHour").ToString.PadLeft(2, "0")
            EndHour = row.Item("EndHour").ToString.PadLeft(2, "0")
            StartMinute = row.Item("StartMinute").ToString.PadLeft(2, "0")
            EndMinute = row.Item("EndMinute").ToString.PadLeft(2, "0")
          Next

          If (bUseTemplateLocks And bDisallowEditDay) Then
            sDisabled = " disabled=""disabled"""
          Else
            sDisabled = ""
          End If
        %>
        <input class="checkbox" id="avail-sun" name="avail-sun" type="checkbox"<% if(sunday=1)then sendb(" checked=""checked"" ") %><% sendb(sDisabled) %> /><label for="avail-sun"><% Sendb(Copient.PhraseLib.Lookup("term.sunday", LanguageID))%></label><br />
        <input class="checkbox" id="avail-mon" name="avail-mon" type="checkbox"<% if(monday=1)then sendb(" checked=""checked"" ") %><% sendb(sDisabled) %> /><label for="avail-mon"><% Sendb(Copient.PhraseLib.Lookup("term.monday", LanguageID))%></label><br />
        <input class="checkbox" id="avail-tue" name="avail-tue" type="checkbox"<% if(tuesday=1)then sendb(" checked=""checked"" ") %><% sendb(sDisabled) %> /><label for="avail-tue"><% Sendb(Copient.PhraseLib.Lookup("term.tuesday", LanguageID))%></label><br />
        <input class="checkbox" id="avail-wed" name="avail-wed" type="checkbox"<% if(wednesday=1)then sendb(" checked=""checked"" ") %><% sendb(sDisabled) %> /><label for="avail-wed"><% Sendb(Copient.PhraseLib.Lookup("term.wednesday", LanguageID))%></label><br />
        <input class="checkbox" id="avail-thu" name="avail-thu" type="checkbox"<% if(thursday=1)then sendb(" checked=""checked"" ") %><% sendb(sDisabled) %> /><label for="avail-thu"><% Sendb(Copient.PhraseLib.Lookup("term.thursday", LanguageID))%></label><br />
        <input class="checkbox" id="avail-fri" name="avail-fri" type="checkbox"<% if(friday=1)then sendb(" checked=""checked"" ") %><% sendb(sDisabled) %> /><label for="avail-fri"><% Sendb(Copient.PhraseLib.Lookup("term.friday", LanguageID))%></label><br />
        <input class="checkbox" id="avail-sat" name="avail-sat" type="checkbox"<% if(saturday=1)then sendb(" checked=""checked"" ") %><% sendb(sDisabled) %> /><label for="avail-sat"><% Sendb(Copient.PhraseLib.Lookup("term.saturday", LanguageID))%></label><br />
        <hr class="hidden" />
      </div>
    </div>
    
    <div id="gutter">
    </div>
    
    <div id="column2">
      <div class="box" id="time">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.time", LanguageID))%>
          </span>
          <% If (IsTemplate) Then%>
          <span class="tempRequire">
            <input type="checkbox" class="tempcheck" id="DisallowEditTime1" name="DisallowEditTime" <% if(bDisallowEditTime)then send(" checked=""checked""") %> />
            <label for="temp-Tiers"><% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
          </span>
          <% ElseIf (bUseTemplateLocks And bDisallowEditTime) Then%>
          <span class="tempRequire">
            <input type="checkbox" class="tempcheck" id="DisallowEditTime2" name="DisallowEditTime" disabled="disabled" checked="checked" />
            <label for="temp-Tiers"><% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
          </span>
          <% End If%>
        </h2>
        <% Sendb(Copient.PhraseLib.Lookup("condition.time", LanguageID))%>
        <br />
        <br class="half" />
        <label for="start_hour"><% Sendb(Copient.PhraseLib.Lookup("term.starts", LanguageID))%></label>
        <br />
        <% 
          If (bUseTemplateLocks And bDisallowEditTime) Then
            sDisabled = " disabled=""disabled"""
          Else
            sDisabled = ""
          End If
        %>
        <input class="shortest" id="start_hour" name="start_hour" maxlength="2" type="text" value="<% Sendb(starthour) %>"<% Sendb(sDisabled) %> />:
        <input class="shortest" id="start_minute" name="start_minute" maxlength="2" type="text" value="<% Sendb(startminute) %>"<% Sendb(sDisabled) %> />
        (<% Sendb(Copient.PhraseLib.Lookup("term.hhmm", LanguageID))%>)<br />
        <br class="half" />
        <label for="end_hour"><% Sendb(Copient.PhraseLib.Lookup("term.ends", LanguageID))%></label>
        <br />
        <input class="shortest" id="end_hour" name="end_hour" maxlength="2" type="text" value="<% Sendb(endhour) %>"<% Sendb(sDisabled) %> />:
        <input class="shortest" id="end_minute" name="end_minute" maxlength="2" type="text" value="<% Sendb(endminute) %>"<% Sendb(sDisabled) %> />
        (<% Sendb(Copient.PhraseLib.Lookup("term.hhmm", LanguageID))%>)<br />
        <br />
      </div>
    </div>
  </div>
</form>

<script type="text/javascript">
<% If (CloseAfterSave) Then %>
    window.close();
<% End If %>
</script>

<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd()
  MyCommon = Nothing
  Logix = Nothing
%>
