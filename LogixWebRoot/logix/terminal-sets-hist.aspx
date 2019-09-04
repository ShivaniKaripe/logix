<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: terminal-sets-hist.aspx 
  ' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' * Copyright © 2002 - 2011.  All rights reserved by:
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
  Dim dst As System.Data.DataTable
  Dim row As System.Data.DataRow
  Dim TerminalSetID As Long
  Dim SetName As String = ""
  Dim Deleted As Boolean = False
  Dim sizeOfData As Integer
  Dim i As Integer = 0
  Dim maxEntries As Integer = 9999
  Dim Shaded As String = "shaded"
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "terminal-sets-hist.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
    
  TerminalSetID = MyCommon.Extract_Val(GetCgiValue("TerminalSetID"))
  
  If (TerminalSetID = 0) Then
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "store-edit.aspx?new=New")
  End If
  
  MyCommon.QueryStr = "select TerminalSetID, Name from TerminalSets with (NoLock) where TerminalSetID=" & TerminalSetID & ";"
  rst = MyCommon.LRT_Select()
  If (rst.Rows.Count = 0) Then
    infoMessage = Copient.PhraseLib.Lookup("error.deleted", LanguageID)
    Deleted = True
  Else
    SetName = MyCommon.NZ(rst.Rows(0).Item("Name"), "")
  End If
  
  Send_HeadBegin("term.terminal-set", "term.history", TerminalSetID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
  Send_HeadEnd()
  Send_BodyBegin(1)
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  Send_Tabs(Logix, 7)
  Send_Subtabs(Logix, 74, 5, , TerminalSetID)
  
  If Not Logix.UserRoles.EditTerminalSets Then
    Send_Denied(1, "perm.store-access")
    GoTo done
  End If
  If Not Logix.UserRoles.ViewHistory Then
    Send_Denied(1, "perm.admin-history")
    GoTo done
  End If
%>
<div id="intro">
  <h1 id="title">
    <%
      Sendb(Copient.PhraseLib.Lookup("term.terminal-set", LanguageID) & " #" & TerminalSetID)
      If SetName <> "" Then
        Sendb(": " & MyCommon.TruncateString(SetName, 40))
      End If
    %>
  </h1>
  <div id="controls">
    <%
      If MyCommon.Fetch_SystemOption(75) AndAlso (Deleted = False) Then
        If (TerminalSetID > 0 And Logix.UserRoles.AccessNotes) Then
          Send_NotesButton(13, TerminalSetID, AdminUserID)
        End If
      End If
    %>
  </div>
</div>
<div id="main">
  <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
  <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.history", LanguageID)) %>">
    <thead>
      <tr>
        <th align="left" scope="col" class="th-timedate">
          <% Sendb(Copient.PhraseLib.Lookup("term.timedate", LanguageID))%>
        </th>
        <th align="left" scope="col" class="th-user">
          <% Sendb(Copient.PhraseLib.Lookup("term.user", LanguageID))%>
        </th>
        <th align="left" scope="col" class="th-action">
          <% Sendb(Copient.PhraseLib.Lookup("term.action", LanguageID))%>
        </th>
      </tr>
    </thead>
    <tbody>
      <%
        MyCommon.QueryStr = "select AU.FirstName, AU.LastName, AL.ActivityDate, AL.Description from ActivityLog as AL with (NoLock) left join AdminUsers as AU with (NoLock) on AL.AdminID=AU.AdminUserID Where ActivityTypeID='48' and LinkID='" & TerminalSetID & "' order by ActivityDate desc;"
        dst = MyCommon.LRT_Select
        sizeOfData = dst.Rows.Count
        While (i < sizeOfData And i < maxEntries)
          Send("<tr class=""" & Shaded & """>")
          If (Not IsDBNull(dst.Rows(i).Item("ActivityDate"))) Then
            Send("  <td>" & Logix.ToShortDateTimeString(dst.Rows(i).Item("ActivityDate"), MyCommon) & "</td>")
          Else
            Send("  <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
          End If
          Send("  <td>" & MyCommon.NZ(dst.Rows(i).Item("FirstName"), "") & " " & MyCommon.NZ(dst.Rows(i).Item("LastName"), "") & "</td>")
          Send("  <td>" & MyCommon.NZ(dst.Rows(i).Item("Description"), "") & "</td>")
          Send("</tr>")
          If Shaded = "shaded" Then
            Shaded = ""
          Else
            Shaded = "shaded"
          End If
          i = i + 1
        End While
        If (sizeOfData = 0) Then
          Send("<tr>")
          Send("  <td colspan=""3""></td>")
          Send("</tr>")
        End If
      %>
    </tbody>
  </table>
</div>
<%
  If MyCommon.Fetch_SystemOption(75) AndAlso (Deleted = False) Then
    If (TerminalSetID > 0 And Logix.UserRoles.AccessNotes) Then
      Send_Notes(13, TerminalSetID, AdminUserID)
    End If
  End If
done:
  Send_BodyEnd()
  MyCommon.Close_LogixRT()
  MyCommon = Nothing
  Logix = Nothing
%>
