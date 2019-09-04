<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: issuance.aspx 
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
  Dim row As DataRow
  Dim AdminUserID As Long
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim dt2 As System.Data.DataTable
  Dim Enabled As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "issuance.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  Send_HeadBegin("term.issuance")
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
  Send_Subtabs(Logix, 8, 4)
  
  
  If (Logix.UserRoles.EditSystemConfiguration = False) Then
    Send_Denied(1, "perm.admin-configuration")
    GoTo done
  End If
  
  MyCommon.QueryStr = "select * from CPE_DeliverableTypes with (NoLock) WHERE DeliverableTypeID NOT IN (6,7) order by DisplayOrder;"
  dt2 = MyCommon.LRT_Select
    
  If (Request.QueryString("save") <> "") Then
    If (Request.QueryString("IssuCheck") = "on") Then
      MyCommon.QueryStr = "Update CPE_SystemOptions with (RowLock) set OptionValue='1' WHERE OptionID=70;"
      MyCommon.LRT_Execute()
      For Each row In dt2.Rows
        If (Request.QueryString("type" & MyCommon.NZ(row.Item("DeliverableTypeID"), 0) & "") = "on") Then
          MyCommon.QueryStr = "Update CPE_DeliverableTypes with (RowLock) set IssuanceEnabled=1 WHERE DeliverableTypeID=" & MyCommon.NZ(row.Item("DeliverableTypeID"), 0) & ";"
          MyCommon.LRT_Execute()
        Else
          MyCommon.QueryStr = "Update CPE_DeliverableTypes with (RowLock) set IssuanceEnabled=0 WHERE DeliverableTypeID=" & MyCommon.NZ(row.Item("DeliverableTypeID"), 0) & ";"
          MyCommon.LRT_Execute()
        End If
      Next
    Else
      MyCommon.QueryStr = "Update CPE_SystemOptions with (RowLock) set OptionValue='0' WHERE OptionID = 70;"
      MyCommon.LRT_Execute()
    End If
    
    MyCommon.Activity_Log(36, 0, AdminUserID, Copient.PhraseLib.Lookup("history.issuancesettings", LanguageID))
  End If
%>

<script type="text/javascript">

function enabledIssu() {
  //Issuance Checkbox variable
  var issu = document.getElementById("IssuCheck");
  
  //Checkbox variables
  <% 
    If dt2.Rows.Count > 0 Then
      For Each row In dt2.Rows
        Send("var type" & MyCommon.NZ(row.Item("DeliverableTypeID"), 0) & " = document.getElementById(""type" & MyCommon.NZ(row.Item("DeliverableTypeID"), 0) & """);")
      Next
    End If
  %>
  
  if(issu.checked){
  //Enable checkboxes
  <%
    If dt2.Rows.Count > 0 Then
      For Each row In dt2.Rows
        Send("type" & MyCommon.NZ(row.Item("DeliverableTypeID"), 0) & ".disabled = false;")
      Next
    End If
  %>
  } else {
  //Diabeling checkboxes 
  <%
    If dt2.Rows.Count > 0 Then
      For Each row In dt2.Rows
        Send("type" & MyCommon.NZ(row.Item("DeliverableTypeID"), 0) & ".disabled = true;")
      Next
    End If
  %>
  }
}
</script>

<form action="#" id="mainform" name="mainform">
  <div id="intro">
    <h1 id="title">
      <% Sendb(Copient.PhraseLib.Lookup("term.issuance", LanguageID))%>
    </h1>
    <div id="controls">
      <%
        If (Logix.UserRoles.EditSystemConfiguration = True) Then
          Send_Save()
        End If
        If MyCommon.Fetch_SystemOption(75) Then
          If (Logix.UserRoles.AccessNotes) Then
            Send_NotesButton(36, 0, AdminUserID)
          End If
        End If
      %>
    </div>
  </div>
  <div id="main">
    <%
      If (infoMessage <> "") Then
        Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
      End If
      If (MyCommon.Fetch_CPE_SystemOption(70)) Then
        Enabled = True
      End If
    %>
    <div id="Issu" style="float: left">
      <input type="checkbox" id="IssuCheck" name="IssuCheck" onclick="javascript:enabledIssu();"<%sendb(IIf(enabled, " checked=""checked""", "")) %> />
      <label for="IssuCheck"><% Sendb(Copient.PhraseLib.Lookup("term.addIssu", LanguageID).Trim)%></label><br />
      <%
        MyCommon.QueryStr = "select * from CPE_DeliverableTypes WHERE DeliverableTypeID NOT IN (6,7) order by DisplayOrder;"
        dt2 = MyCommon.LRT_Select
        If dt2.Rows.Count > 0 Then
          For Each row In dt2.Rows
            Send("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type=""checkbox"" id=""type" & MyCommon.NZ(row.Item("DeliverableTypeID"), 0) & """ name=""type" & MyCommon.NZ(row.Item("DeliverableTypeID"), 0) & """" & IIf(row.Item("IssuanceEnabled"), " checked=""checked""", "") & IIf(Enabled, "", " disabled=""disabled""") & " />")
            Send("  <label for=""type" & MyCommon.NZ(row.Item("DeliverableTypeID"), 0) & """ >" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "</label><br />")
          Next
        End If
      %>
    </div>
  </div>
</form>
<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (Logix.UserRoles.AccessNotes) Then
      Send_Notes(21, 0, AdminUserID)
    End If
  End If
done:
  Send_BodyEnd()
  MyCommon.Close_LogixRT()
  MyCommon = Nothing
  Logix = Nothing
%>
