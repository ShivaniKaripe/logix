<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: user-list.aspx 
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
  Dim dst As System.Data.DataTable
  Dim dtUsers As System.Data.DataTable
  Dim rst2 As DataTable
  Dim row As System.Data.DataRow
  Dim row2 As DataRow
  Dim rowUser As DataRow
  Dim rows As DataRow() = Nothing
  Dim idNumber As Integer
  Dim idSearch As String
  Dim idSearchText As String
  Dim PageNum As Integer = 0
  Dim MorePages As Boolean
  Dim linesPerPage As Integer = 20
  Dim sizeOfData As Integer
  Dim i As Integer = 0
  Dim Shaded As String = "shaded"
  Dim SortText As String = "AdminUserID"
  Dim SortDirection As String = ""
  Dim RoleName As String = ""
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "user-list.aspx"
  MyCommon.Open_LogixRT()
  If MyCommon.IsIntegrationInstalled(Copient.CommonIncConfigurable.Integrations.PREFERENCE_MANAGER) Then
    MyCommon.Open_PrefManRT()
  End If

  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  
  PageNum = Request.QueryString("pagenum")
  If PageNum < 0 Then
    PageNum = 0
  End If
  MorePages = False
  
  If (Request.QueryString("SortText") <> "") Then
    SortText = Request.QueryString("SortText")
  End If
  
  If (Request.QueryString("pagenum") = "") Then
    If (Request.QueryString("SortDirection") = "ASC") Then
      SortDirection = "DESC"
    ElseIf (Request.QueryString("SortDirection") = "DESC") Then
      SortDirection = "ASC"
    Else
      SortDirection = "DESC"
    End If
  Else
    SortDirection = Request.QueryString("SortDirection")
  End If
  
  Send_HeadBegin("term.users")
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
  Send_Subtabs(Logix, 8, 9)
  
  
  If (Request.QueryString("searchterms") <> "") Then
    If (Integer.TryParse(Request.QueryString("searchterms"), idNumber)) Then
      idSearch = idNumber.ToString
    Else
      idSearch = "-1"
    End If
    idSearchText = MyCommon.Parse_Quotes(HttpUtility.UrlDecode(Request.QueryString("searchterms")))
    MyCommon.QueryStr = "SELECT DISTINCT AU.AdminUserID,AU.UserName,AU.FirstName,AU.LastName,AU.LastLogin,AU.LastAuth " & _
                        "FROM AdminUsers AS AU WITH (NoLock) WHERE (AdminUserID=" & idSearch & " or UserName like '%" & idSearchText & "%' or FirstName like '%" & idSearchText & "%' or LastName like '%" & idSearchText & "%')"
  Else
    MyCommon.QueryStr = "SELECT DISTINCT AU.AdminUserID,AU.UserName,AU.FirstName,AU.LastName,AU.LastLogin,AU.LastAuth " & _
                        "FROM AdminUsers AS AU WITH (NoLock) "
  End If
  
  If (Request.QueryString("SortText") <> "") Then
    MyCommon.QueryStr += "ORDER BY " & SortText & " " & SortDirection
  End If
  
  dst = MyCommon.LRT_Select
  
  ' store all the data into one datatable and filter it as necessary            
  dtUsers = New DataTable
  dtUsers.Columns.Add("AdminUserID", System.Type.GetType("System.Int64"))
  dtUsers.Columns.Add("UserName", System.Type.GetType("System.String"))
  dtUsers.Columns.Add("FirstName", System.Type.GetType("System.String"))
  dtUsers.Columns.Add("LastName", System.Type.GetType("System.String"))
  dtUsers.Columns.Add("Roles", System.Type.GetType("System.String"))
  dtUsers.Columns.Add("LastAuth", System.Type.GetType("System.DateTime"))
  If MyCommon.IsIntegrationInstalled(Copient.CommonIncConfigurable.Integrations.PREFERENCE_MANAGER) Then
    dtUsers.Columns.Add("EPMRoles", System.Type.GetType("System.String"))
  End If
  
  For Each row In dst.Rows
    rowUser = dtUsers.NewRow
    rowUser.Item("AdminUserID") = row.Item("AdminUserID")
    rowUser.Item("UserName") = row.Item("UserName")
    rowUser.Item("FirstName") = row.Item("FirstName")
    rowUser.Item("LastName") = row.Item("LastName")
    ' find the roles based on userid
    RoleName = ""
    MyCommon.QueryStr = "select Distinct AR.RoleName,AR.PhraseID from AdminUserRoles AUR with (NoLock) " & _
                        "inner join AdminRoles AR with (NoLock) on AR.RoleID = AUR.RoleID " & _
                        "where AdminUserID =" & row.Item("AdminUserID")
    rst2 = MyCommon.LRT_Select
    For Each row2 In rst2.Rows
      If (Not IsDBNull(row2.Item("PhraseID"))) Then
        RoleName &= Copient.PhraseLib.Lookup(row2.Item("PhraseID"), LanguageID) & " <br />"
      Else
        RoleName &= row2.Item("RoleName") & " <br />"
      End If
    Next
    rowUser.Item("Roles") = RoleName
    
    If MyCommon.IsIntegrationInstalled(Copient.CommonIncConfigurable.Integrations.PREFERENCE_MANAGER) Then
      RoleName = ""
      MyCommon.QueryStr = "select Distinct " & _
                          "  case when ISNULL(PMPT.Phrase, '') <> '' then PMPT.Phrase " & _
                          "       else AR.RoleName " & _
                          "  end as RoleName " & _
                          "from AdminUserRoles AUR with (NoLock)  " & _
                          "inner join AdminRoles AR with (NoLock) on AR.RoleID = AUR.RoleID  " & _
                          "left join PM_Phrases as PMP with (NoLock) on PMP.PhraseID = AR.PhraseID " & _
                          "left join PM_PhraseText as PMPT with (NoLock) on PMPT.PhraseID = PMP.PhraseID and LanguageID=" & LanguageID & " " & _
                          "where AdminUserID = " & row.Item("AdminUserID")
      rst2 = MyCommon.PMRT_Select
      For Each row2 In rst2.Rows
        RoleName &= row2.Item("RoleName") & " <br />"
      Next
      rowUser.Item("EPMRoles") = RoleName
    End If
    
    rowUser.Item("LastAuth") = row.Item("LastAuth")
    dtUsers.Rows.Add(rowUser)
  Next
  
  rows = dtUsers.Select("1=1", SortText & " " & SortDirection)
  sizeOfData = rows.Length
  
  ' set page variable(s)
  i = linesPerPage * PageNum
  
  ' check if user has permission to view this page
  If (Logix.UserRoles.ViewOthersInfo = False) Then
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "user-edit.aspx?UserID=" & AdminUserID)
  End If
  
  If (sizeOfData = 1 AndAlso Request.QueryString("searchterms") <> "") Then
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "user-edit.aspx?UserID=" & dst.Rows(i).Item("AdminUserID"))
  End If
%>
<div id="intro">
  <h1 id="title">
    <% Sendb(Copient.PhraseLib.Lookup("term.users", LanguageID))%>
  </h1>
  <div id="controls">
    <form action="user-edit.aspx" id="controlsform" name="controlsform">
      <%
        If (Logix.UserRoles.CreateAdminUsers = True) Then
          Send_New()
        End If
        'If MyCommon.Fetch_SystemOption(75) Then
        '  If (Logix.UserRoles.AccessNotes) Then
        '    Send_NotesButton(19, 0, AdminUserID)
        '  End If
        'End If
      %>
    </form>
  </div>
</div>
<div id="main">
  <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
  <% Send_Listbar(linesPerPage, sizeOfData, PageNum, Server.HtmlEncode(Request.QueryString("searchterms")), "&SortText=" & SortText & "&SortDirection=" & SortDirection)%>
  <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.users", LanguageID)) %>">
    <thead>
      <tr>
        <th align="left" class="th-id" scope="col">
          <a href="user-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&SortText=AdminUserID&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>
          </a>
          <%If SortText = "AdminUserID" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" class="th-username" scope="col">
          <a href="user-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&SortText=UserName&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.username", LanguageID))%>
          </a>
          <%If SortText = "UserName" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="right" class="th-firstname" scope="col" style="text-align: right;">
          <a href="user-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&SortText=FirstName&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.firstname", LanguageID))%>
          </a>
          <%If SortText = "FirstName" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" class="th-surname" scope="col">
          <a href="user-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&SortText=LastName&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.surname", LanguageID))%>
          </a>
          <%If SortText = "LastName" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" class="th-roles" scope="col">
          <!--
          <a href="user-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&SortText=Roles&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.roles", LanguageID))%>
          </a>
          -->
          <% Sendb(Copient.PhraseLib.Lookup("term.ams-roles", LanguageID))%>
          <%If SortText = "Roles" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
       <% If MyCommon.IsIntegrationInstalled(Copient.CommonIncConfigurable.Integrations.PREFERENCE_MANAGER) Then%>
        <th align="left" class="th-roles" scope="col">
          <% Sendb(Copient.PhraseLib.Lookup("term.epm-roles", LanguageID))%>
        </th>
       <% End If%>
        <th align="left" class="th-datetime" scope="col">
          <a href="user-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&SortText=LastAuth&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.lastactivity", LanguageID))%>
          </a>
          <%If SortText = "LastAuth" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
      </tr>
    </thead>
    <tbody>
      <%
        While (i < sizeOfData AndAlso i < linesPerPage + linesPerPage * PageNum)
          Send("<tr class=""" & Shaded & """>")
          Send("  <td>" & rows(i).Item("AdminUserID") & "</td>")
          Send("  <td><a href=""user-edit.aspx?UserID=" & rows(i).Item("AdminUserID") & """>" & MyCommon.SplitNonSpacedString(rows(i).Item("UserName"), 15) & "</a></td>")
          Send("  <td align=""right"">" & MyCommon.SplitNonSpacedString(MyCommon.NZ(rows(i).Item("FirstName"), "&nbsp;"), 20) & "</td>")
          Send("  <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(rows(i).Item("LastName"), "&nbsp;"), 20) & "</td>")
          Send("  <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(rows(i).Item("Roles"), "&nbsp;"), 25) & "</td>")
          If MyCommon.IsIntegrationInstalled(Copient.CommonIncConfigurable.Integrations.PREFERENCE_MANAGER) Then
            Send("  <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(rows(i).Item("EPMRoles"), "&nbsp;"), 25) & "</td>")
          End If
          If (Not IsDBNull(rows(i).Item("LastAuth"))) Then
            Send("  <td>" & Logix.ToShortDateTimeString(rows(i).Item("LastAuth"), MyCommon) & "</td>")
          Else
            Send("  <td>" & Copient.PhraseLib.Lookup("term.never", LanguageID) & "</td>")
          End If
          Send("</tr>")
          If Shaded = "shaded" Then
            Shaded = ""
          Else
            Shaded = "shaded"
          End If
          i = i + 1
        End While
        MyCommon.Close_LogixRT()
      %>
    </tbody>
  </table>
</div>
<%
  'If MyCommon.Fetch_SystemOption(75) Then
  '  If (Logix.UserRoles.AccessNotes) Then
  '    Send_Notes(19, 0, AdminUserID)
  '  End If
  'End If
done:
  Send_BodyEnd()
  Send_FocusScript("searchform", "searchterms")
  MyCommon.Close_LogixRT()
  If MyCommon.IsIntegrationInstalled(Copient.CommonIncConfigurable.Integrations.PREFERENCE_MANAGER) Then
    MyCommon.Close_PrefManRT()
  End If
  Logix = Nothing
  MyCommon = Nothing
%>
