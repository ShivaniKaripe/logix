<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: cgroup-hist.aspx 
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
  Dim dst As System.Data.DataTable
  Dim row As System.Data.DataRow
  Dim CustomerGroupID As Long
  Dim GName As String
  Dim GNameTitle As String
  Dim sizeOfData As Integer
  Dim DefaultIDType As Integer
  Dim i As Integer = 0
  Dim Deleted As Boolean = False
  Dim maxEntries As Integer = 9999
  Dim Shaded As String = "shaded"
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "cgroup-hist.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  DefaultIDType = MyCommon.Fetch_SystemOption(30)
  CustomerGroupID = Request.QueryString("CustomerGroupID")
  GName = Request.QueryString("GroupName")
  ' Check in case it was a POST instead of get
  If (CustomerGroupID = 0 And Not Request.QueryString("save") <> "") Then
    CustomerGroupID = Request.Form("CustomerGroupID")
    GName = Request.Form("GroupName")
  End If
  
  If (CustomerGroupID = 0) Then
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "cgroup-edit.aspx?new=New")
  End If
  
  MyCommon.QueryStr = "select Name, Deleted from CustomerGroups with (NoLock) where CustomerGroupID=" & CustomerGroupID & ";"
  rst = MyCommon.LRT_Select()
  If (rst.Rows.Count = 0) Then
    infoMessage = Copient.PhraseLib.Lookup("error.deleted", LanguageID)
    Deleted = True
  Else
    GName = MyCommon.NZ(rst.Rows(0).Item("Name"), "")
    If (MyCommon.NZ(rst.Rows(0).Item("Deleted"), "1") = "1") Then
      infoMessage = Copient.PhraseLib.Lookup("error.deleted", LanguageID)
      Deleted = True
    End If
  End If
  
  Send_HeadBegin("term.customergroup", "term.history", CustomerGroupID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
  Send_HeadEnd()
  Send_BodyBegin(1)
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  Send_Tabs(Logix, 3)
  Send_Subtabs(Logix, 31, 4, , CustomerGroupID)
  
  If (Logix.UserRoles.AccessCustomerGroups = False) Then
    Send_Denied(1, "perm.cgroup-access")
    GoTo done
  End If
  If (Logix.UserRoles.ViewHistory = False) Then
    Send_Denied(1, "perm.admin-history")
    GoTo done
  End If
%>
<div id="intro">
  <h1 id="title">
    <%
      Sendb(Copient.PhraseLib.Lookup("term.customergroup", LanguageID) & " #" & CustomerGroupID)
      If GName <> "" Then
        GNameTitle = GName
        If (Len(GNameTitle) > 30) Then
            GNameTitle = Left(GNameTitle, 27) & "..."
        End If
        Sendb(": " & GNameTitle)
      End If
    %>
  </h1>
  <div id="controls">
    <%
      If MyCommon.Fetch_SystemOption(75) AndAlso (Deleted = False) Then
        If (CustomerGroupID > 0 And Logix.UserRoles.AccessNotes) Then
          Send_NotesButton(5, CustomerGroupID, AdminUserID)
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
        MyCommon.QueryStr = "select case when AL.AdminID is null then '" & Copient.PhraseLib.Lookup("term.systemuser", LanguageID) & "' else AU.FirstName end AS FirstName, case when AL.AdminID is null then AL.LinkID3 else AU.LastName end AS LastName, AL.ActivityDate, AL.Description from ActivityLog as AL with (NoLock) left join AdminUsers as AU with (NoLock) on AL.AdminID=AU.AdminUserID Where ActivityTypeID='4' and LinkID='" & CustomerGroupID & "' order by ActivityDate desc;"
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
    If (CustomerGroupID > 0 And Logix.UserRoles.AccessNotes) Then
      Send_Notes(5, CustomerGroupID, AdminUserID)
    End If
  End If
done:
  Send_BodyEnd()
  MyCommon.Close_LogixRT()
  MyCommon = Nothing
  Logix = Nothing
%>
