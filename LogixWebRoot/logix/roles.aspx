<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: roles.aspx 
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
  Dim rst As DataTable
  Dim l_name As String
  Dim l_RID As Long
  dim sExtName as String = ""
  Dim AdminUserID As Long
  Dim CategoryName As String = ""
  Dim ShowByCategory As Boolean = False
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim InstalledEngines As String = "2"
  Dim InstalledEngineCt As Integer = -1
  Dim ExcludedPermissions As New Hashtable()
  Dim BannersDisabled As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "roles.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)

  Send_HeadBegin("term.roles")
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
  
  ' Exclude the Web Site Engine (engine id =3) as that has no corresponding permission exclusions
  MyCommon.QueryStr = "select EngineID from PromoEngines with (NoLock) where Installed=1 and EngineID not in (3);"
  dst = MyCommon.LRT_Select
  If (dst.Rows.Count > 0) Then
    InstalledEngineCt = dst.Rows.Count
    InstalledEngines = ""
    For Each row In dst.Rows
      If (InstalledEngines <> "") Then InstalledEngines &= ","
      InstalledEngines &= row.Item("EngineID")
    Next
  End If
  
  ' load up all the excluded permissions and a count of the engines excluded for this permissions
  MyCommon.QueryStr = "select Distinct PermissionID, Count(PermissionID) as EnginesExcluded " & _
                      "from PermissionExcludedEngines with (NoLock) where EngineID in (" & InstalledEngines & ") group by PermissionID;"
  dst = MyCommon.LRT_Select
  If (dst.Rows.Count > 0) Then
    For Each row In dst.Rows
      ExcludedPermissions.Add(MyCommon.NZ(row.Item("PermissionID"), "-1").ToString, MyCommon.NZ(row.Item("EnginesExcluded"), -1))
    Next
  End If
  
  BannersDisabled = (MyCommon.Fetch_SystemOption(66) <> "1")
  
  l_RID = MyCommon.Extract_Val(Request.QueryString("role"))
  
  If (Request.QueryString("add") <> "") Then
    l_name = MyCommon.NZ(Request.QueryString("name"), "")
    l_name = MyCommon.Parse_Quotes(Logix.TrimAll(l_name))
    If l_name = "" Then
      infoMessage = Copient.PhraseLib.Lookup("roles.noname", LanguageID)
    Else
      MyCommon.QueryStr = "SELECT RoleID FROM AdminRoles with (NoLock) WHERE RoleName = '" & l_name & "';"
      dst = MyCommon.LRT_Select
      If (dst.Rows.Count > 0) Then
        infoMessage = Copient.PhraseLib.Lookup("roles.nameused", LanguageID)
      Else
        MyCommon.QueryStr = "INSERT INTO AdminRoles with (RowLock) (RoleName, PhraseID) VALUES (N'" & l_name & "', NULL);"
        MyCommon.LRT_Execute()
        MyCommon.QueryStr = "SELECT RoleID FROM AdminRoles with (NoLock) WHERE RoleName = '" & l_name & "';"
        dst = MyCommon.LRT_Select
        If dst.Rows.Count > 0 Then
          MyCommon.Activity_Log(20, MyCommon.NZ(dst.Rows(0).Item("RoleID"), 0), AdminUserID, Copient.PhraseLib.Lookup("history.role-create", LanguageID))
        End If
      End If
    End If
  ElseIf (Request.QueryString("delete") <> "") Then
    If l_RID = 0 Or l_RID = 1 Then
      infoMessage = Copient.PhraseLib.Lookup("roles.nodelete", LanguageID)
    Else
      ' Make sure this role is not in use
      MyCommon.QueryStr = "select AdminUserID from adminuserroles with (NoLock) where roleid=" & l_RID
      rst = MyCommon.LRT_Select
      If (rst.Rows.Count > 0) Then
        infoMessage = Copient.PhraseLib.Lookup("roles.inuse", LanguageID)
      Else
        MyCommon.QueryStr = "DELETE FROM AdminRoles with (RowLock) WHERE RoleID = " & l_RID
        MyCommon.LRT_Execute()
        MyCommon.Activity_Log(20, l_RID, AdminUserID, Copient.PhraseLib.Lookup("history.role-delete", LanguageID))
        Response.Status = "301 Moved Permanently"
        Response.AddHeader("Location", "roles.aspx")
      End If
    End If
  ElseIf (Request.QueryString("selectperm") <> "") Then
    Dim tmpRoles() As String
    Dim z As Integer
    If Not (Request.QueryString("perm-avail") = "") Then
      tmpRoles = Request.QueryString("perm-avail").Split(",")
      For z = 0 To tmpRoles.GetUpperBound(0)
        MyCommon.QueryStr = "Select RoleID from RolePermissions with (NoLock) where RoleID=" & l_RID & " and PermissionID=" & tmpRoles(z)
        rst = MyCommon.LRT_Select()
        If rst.Rows.Count = 0 Then
          MyCommon.QueryStr = "INSERT into RolePermissions with (RowLock) (PermissionID,RoleID) values(" & tmpRoles(z) & "," & l_RID & ");"
          MyCommon.LRT_Execute()
        End If
      Next
      MyCommon.Activity_Log(20, l_RID, AdminUserID, Copient.PhraseLib.Lookup("history.role-add", LanguageID))
    End If
  ElseIf (Request.QueryString("removeperm") <> "") And Not (Request.QueryString("perm-select") = "") Then
    Dim tmpRoles2() As String
    Dim w As Integer
    tmpRoles2 = Request.QueryString("perm-select").Split(",")
    For w = 0 To tmpRoles2.GetUpperBound(0)
      MyCommon.QueryStr = "DELETE from RolePermissions with (RowLock) where PermissionID=" & tmpRoles2(w) & " and RoleID=" & l_RID & ";"
      MyCommon.LRT_Execute()
    Next
    MyCommon.Activity_Log(20, l_RID, AdminUserID, Copient.PhraseLib.Lookup("history.role-remove", LanguageID))
  ElseIf (Request.QueryString("save") <> "") Then
    If l_RID > 0 Then
      sExtName = MyCommon.NZ(Request.QueryString("extname"), "")
      sExtName = MyCommon.Parse_Quotes(Logix.TrimAll(sExtName))
      MyCommon.QueryStr = "update AdminRoles with (RowLock) set ExtRoleName = '" & sExtName & "' where RoleID=" & l_RID & ";"
      MyCommon.LRT_Execute()
    End If
  End If
%>
<form action="#" id="frmRoles" name="frmRoles">
  <div id="intro">
    <h1 id="title">
      <% Sendb(Copient.PhraseLib.Lookup("term.roles", LanguageID))%>
    </h1>
    <div id="controls">
      <%
        If MyCommon.Fetch_SystemOption(75) Then
          If (Logix.UserRoles.AccessNotes) Then
            Send_NotesButton(25, 0, AdminUserID)
          End If
        End If
      %>
    </div>
  </div>
  <div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <div id="column">
      <%
        If (Logix.UserRoles.EditRoles = False) Then
          Send("<select id=""roles"" name=""roles"" class=""long"" size=""15"">")
          MyCommon.QueryStr = "SELECT RoleID, RoleName FROM AdminRoles with (NoLock) order by RoleName"
          dst = MyCommon.LRT_Select
          For Each row In dst.Rows
            Send("  <option value=""" & row.Item("RoleID") & """>" & row.Item("RoleName") & "</option>")
          Next
          Send("</select>")
          Send("</div>")
          Send("</form>")
          GoTo done
        End If
      %>
      <div class="box" id="roleadd">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("roles.add", LanguageID))%>
          </span>
        </h2>
        <label for="name"><% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>:</label><br />
        <input type="text" id="name" name="name" class="long" maxlength="100" value="" />
        <input type="submit" class="regular" id="add" name="add" value="<% Sendb(Copient.PhraseLib.Lookup("term.add", LanguageID)) %>" /><br />
        <hr class="hidden" />
      </div>
      
      <% If (Request.QueryString("optView") = "0") Then%>
      <div class="box" id="roleedit">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("roles.edit", LanguageID))%>
          </span>
        </h2>
        <label for="role"><% Sendb(Copient.PhraseLib.Lookup("roles.editnote", LanguageID))%></label>
        <br />
        <select id="role" name="role" class="long">
          <option value="">
            <% Sendb(Copient.PhraseLib.Lookup("roles.select", LanguageID))%>
          </option>
          <%
            MyCommon.QueryStr = "SELECT RoleID, RoleName, PhraseID FROM AdminRoles with (NoLock) order by RoleName"
            dst = MyCommon.LRT_Select
            For Each row In dst.Rows
              Sendb("  <option value=""" & row.Item("RoleID") & """")
              If l_RID = row.Item("RoleID") Then
                Sendb(" selected=""selected""")
              End If
              Sendb(">")
              If IsDBNull(row.Item("PhraseID")) Then
                Sendb(row.Item("RoleName"))
              Else
                If (row.Item("PhraseID") = 0) Then
                  Sendb(row.Item("RoleName"))
                Else
                  Sendb(Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID))
                End If
              End If
              Send("</option>")
            Next
          %>
        </select>
        <span style="position:relative; left:160px; height:20px; border:solid 1px #808080; padding-top:3px;">
          <input type="radio" id="optView1" name="optView" value="0" checked="checked" />
          <label for="optView1"><% Sendb(Copient.PhraseLib.Lookup("term.viewalphabetically", LanguageID))%></label>
          <input type="radio" id="optView2" name="optView" value="1" onclick="document.frmRoles.submit();" />
          <label for="optView2"><% Sendb(Copient.PhraseLib.Lookup("term.viewbycategory", LanguageID))%>&nbsp;</label>
        </span>
        <br />
        <input type="submit" class="regular" id="edit" name="edit" value="<% Sendb(Copient.PhraseLib.Lookup("term.edit", LanguageID)) %>" />&nbsp;
        <input type="submit" class="regular" id="delete" name="delete" onclick="if(confirm('<% Sendb(Copient.PhraseLib.Lookup("confirm.delete", LanguageID)) %>')){}else{return false}" value="<% Sendb(Copient.PhraseLib.Lookup("term.delete", LanguageID)) %>" /><br />
        <br />
        <div style="float:left; position:relative;">
          <label for="perm-select"><b><% Sendb(Copient.PhraseLib.Lookup("roles.selected", LanguageID))%></b></label>
          <br />
          <select class="wideselector" multiple="multiple" id="perm-select" name="perm-select" style="height: 200px;">
            <%
              l_RID = MyCommon.Extract_Val(Request.QueryString("role"))
              MyCommon.QueryStr = "select distinct RP.PermissionID, P.Description, P.PhraseID, P.CategoryID " & _
                                  "FROM RolePermissions AS RP with (NoLock) " & _
                                  "LEFT JOIN Permissions AS P with (NoLock) ON RP.PermissionID=P.PermissionID " & _
                                  "WHERE RoleID=" & l_RID & "ORDER BY Description;"
              rst = MyCommon.LRT_Select
              For Each row In rst.Rows
                If (BannersDisabled AndAlso MyCommon.NZ(row.Item("CategoryID"), -1) = 8) Then
                  ' skip writing this banner option as banners are disabled
                Else
                  Send("    <option value=""" & row.Item("PermissionID") & """>" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID) & "</option>")
                End If
              Next
            %>
          </select>
        </div>
        <div style="float:left; padding:75px 2px 1px 2px; position:relative;">
          <%
            If l_RID = 0 Or l_RID = 1 Then
              Send("   <input type=""submit"" class=""arrowadd"" id=""selectperm"" name=""selectperm"" value=""&#171;"" title=""" & Copient.PhraseLib.Lookup("term.select", LanguageID) & """ disabled=""disabled"" /><br clear=""all"" />")
              Send("   <br class=""half"" />")
              Send("   <input type=""submit"" class=""arrowrem"" id=""removeperm"" name=""removeperm"" value=""&#187;"" title=""" & Copient.PhraseLib.Lookup("term.remove", LanguageID) & """ disabled=""disabled"" />")
            Else
              Send("   <input type=""submit"" class=""arrowadd"" id=""selectperm"" name=""selectperm"" value=""&#171;"" title=""" & Copient.PhraseLib.Lookup("term.select", LanguageID) & """ /><br clear=""all"" />")
              Send("   <br class=""half"" />")
              Send("   <input type=""submit"" class=""arrowrem"" id=""removeperm"" name=""removeperm"" value=""&#187;"" title=""" & Copient.PhraseLib.Lookup("term.remove", LanguageID) & """ />")
            End If
          %>
        </div>
        <div style="float:left; position:relative;">
          <label for="perm-avail"><b><% Sendb(Copient.PhraseLib.Lookup("roles.available", LanguageID))%></b></label>
          <br />
          <select class="wideselector" multiple="multiple" id="perm-avail" name="perm-avail" style="height: 200px;">
            <%
              MyCommon.QueryStr = "SELECT distinct RP.PermissionID, P.Description, P.PhraseID, P.CategoryID " & _
                                  "FROM RolePermissions AS RP with (NoLock) " & _
                                  "LEFT JOIN Permissions AS P with (NoLock) ON RP.PermissionID=P.PermissionID " & _
                                  "WHERE P.PermissionID not in (select PermissionID from RolePermissions where RoleID=" & l_RID & ") " & _
                                  "ORDER BY Description"
              dst = MyCommon.LRT_Select
              For Each row In dst.Rows
                If Not (IsExcludedPermission(MyCommon.NZ(row.Item("PermissionID"), "-1").ToString, ExcludedPermissions, InstalledEngineCt)) Then
                  If (BannersDisabled AndAlso MyCommon.NZ(row.Item("CategoryID"), -1) = 8) Then
                    ' skip writing this banner option as banners are disabled
                  Else
                    Send("    <option value=""" & row.Item("PermissionID") & """>" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID) & "</option>")
                  End If
                End If
              Next
            %>
          </select>
        </div>
        <br clear="left" />
        <br class="zero" />
      </div>
      <% If (Request.QueryString("role") <> "" and MyCommon.Fetch_SystemOption(90)) Then%>
        <%
          l_RID = MyCommon.Extract_Val(Request.QueryString("role"))
          MyCommon.QueryStr = "select ExtRoleName from Adminroles with (NoLock) where RoleID=" & l_RID & ";"
          rst = MyCommon.LRT_Select
          if rst.Rows.Count > 0 then
            sExtName = MyCommon.NZ(rst.Rows(0).Item("ExtRoleName"),"")
          end if
        %>
        <br class="half" />
        <div class="box" id="extrole2">
          <h2>
            <span>
              <% Sendb(Copient.PhraseLib.Lookup("roles.editextname", LanguageID))%>
            </span>
          </h2>
          <label for="name"><% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>:</label><br />
          <input type="text" id="extname" name="extname" class="long" maxlength="100" value="<% Sendb(sExtName)%>" />
          <input type="submit" class="regular" id="save" name="save" value="<% Sendb(Copient.PhraseLib.Lookup("term.save", LanguageID)) %>" /><br />
          <hr class="hidden" />
        </div>
      <% End If%>
      
      <% Else%>
      <div class="box" id="roleedit">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("roles.edit", LanguageID))%>
          </span>
        </h2>
        <label for="role"><% Sendb(Copient.PhraseLib.Lookup("roles.editnote", LanguageID))%></label>
        <br />
        <select id="role" name="role" class="long">
          <option value="">
            <% Sendb(Copient.PhraseLib.Lookup("roles.select", LanguageID))%>
          </option>
          <%
            MyCommon.QueryStr = "SELECT RoleID, RoleName, PhraseID FROM AdminRoles with (NoLock) order by RoleName"
            dst = MyCommon.LRT_Select
            For Each row In dst.Rows
              Sendb("  <option value=""" & row.Item("RoleID") & """")
              If l_RID = row.Item("RoleID") Then
                Sendb(" selected=""selected""")
              End If
              Sendb(">")
              If IsDBNull(row.Item("PhraseID")) Then
                Sendb(row.Item("RoleName"))
              Else
                If (row.Item("PhraseID") = 0) Then
                  Sendb(row.Item("RoleName"))
                Else
                  Sendb(Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID))
                End If
              End If
              Send("</option>")
            Next
          %>
        </select>
        <span style="position: relative; left: 160px; height: 20px; border: solid 1px #808080;
          padding-top: 3px;">
          <input type="radio" id="optView1" name="optView" value="0" onclick="document.frmRoles.submit();" />
          <label for="optView1"><% Sendb(Copient.PhraseLib.Lookup("term.viewalphabetically", LanguageID))%></label>
          <input type="radio" id="optView2" name="optView" value="1" checked="checked" />
          <label for="optView2"><% Sendb(Copient.PhraseLib.Lookup("term.viewbycategory", LanguageID))%>&nbsp;</label>
        </span>
        <br />
        <input type="submit" class="regular" id="edit" name="edit" value="<% Sendb(Copient.PhraseLib.Lookup("term.edit", LanguageID)) %>" />&nbsp;
        <input type="submit" class="regular" id="delete" name="delete" onclick="if(confirm('<% Sendb(Copient.PhraseLib.Lookup("confirm.delete", LanguageID)) %>')){}else{return false}" value="<% Sendb(Copient.PhraseLib.Lookup("term.delete", LanguageID)) %>" /><br />
        <br />
        <div style="float:left; position:relative;">
          <label for="perm-select"><b><% Sendb(Copient.PhraseLib.Lookup("roles.selected", LanguageID))%></b></label>
          <br />
          <select class="wideselector" multiple="multiple" id="perm-select" name="perm-select" style="height: 200px;">
            <%
              CategoryName = ""
              l_RID = MyCommon.Extract_Val(Request.QueryString("role"))
              MyCommon.QueryStr = "SELECT RP.RoleID, RP.PermissionID, PC.Description as CategoryName, PC.PhraseID as CategoryPhraseID, P.Description, P.PhraseID, PC.CategoryID " & _
                                  "FROM RolePermissions AS RP with (NoLock) " & _
                                  "LEFT JOIN Permissions AS P with (NoLock) ON RP.PermissionID=P.PermissionID " & _
                                  "LEFT JOIN PermissionCategories as PC on P.CategoryID=PC.CategoryID " & _
                                  "WHERE RoleID=" & l_RID & " ORDER BY PC.Description, P.Description;"
              rst = MyCommon.LRT_Select
              If (rst.Rows.Count > 0) Then
                For Each row In rst.Rows
                  If (BannersDisabled AndAlso MyCommon.NZ(row.Item("CategoryID"), -1) = 8) Then
                    ' skip adding the banner permissions if banners are not enabled
                  Else
                    If (CategoryName <> Copient.PhraseLib.Lookup(row.Item("CategoryPhraseID"), LanguageID)) Then
                      If (CategoryName <> "") Then
                        Send("</optgroup>")
                      End If
                      Send("   <optgroup label=""" & Copient.PhraseLib.Lookup(row.Item("CategoryPhraseID"), LanguageID) & """>  ")
                    End If
                    Send("    <option value=""" & row.Item("PermissionID") & """>" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID) & "</option>")
                    CategoryName = Copient.PhraseLib.Lookup(row.Item("CategoryPhraseID"), LanguageID)
                  End If
                Next
                Send("</optgroup>")
              End If
            %>
          </select>
        </div>
        <div style="float:left; padding:75px 2px 1px 2px; position:relative;">
          <%
            If l_RID = 0 Or l_RID = 1 Then
              Send("   <input type=""submit"" class=""arrowadd"" id=""selectperm"" name=""selectperm"" value=""&#171;"" title=""" & Copient.PhraseLib.Lookup("term.select", LanguageID) & """ disabled=""disabled"" /><br clear=""all"" />")
              Send("   <br class=""half"" />")
              Send("   <input type=""submit"" class=""arrowrem"" id=""removeperm"" name=""removeperm"" value=""&#187;"" title=""" & Copient.PhraseLib.Lookup("term.remove", LanguageID) & """ disabled=""disabled"" />")
            Else
              Send("   <input type=""submit"" class=""arrowadd"" id=""selectperm"" name=""selectperm"" value=""&#171;"" title=""" & Copient.PhraseLib.Lookup("term.select", LanguageID) & """ /><br clear=""all"" />")
              Send("   <br class=""half"" />")
              Send("   <input type=""submit"" class=""arrowrem"" id=""removeperm"" name=""removeperm"" value=""&#187;"" title=""" & Copient.PhraseLib.Lookup("term.remove", LanguageID) & """ />")
            End If
          %>
        </div>
        <div style="float:left; position:relative;">
          <label for="perm-avail"><b><% Sendb(Copient.PhraseLib.Lookup("roles.available", LanguageID))%></b></label>
          <br />
          <select class="wideselector" multiple="multiple" id="perm-avail" name="perm-avail" style="height:200px;">
            <%
              CategoryName = ""
              MyCommon.QueryStr = "SELECT distinct RP.PermissionID, PC.Description as CategoryName, PC.PhraseID as CategoryPhraseID, P.Description, P.PhraseID, PC.CategoryID " & _
                                  "FROM RolePermissions AS RP with (NoLock) " & _
                                  "LEFT JOIN Permissions AS P with (NoLock) ON RP.PermissionID=P.PermissionID " & _
                                  "LEFT JOIN PermissionCategories as PC on P.CategoryID=PC.CategoryID " & _
                                  "WHERE P.PermissionID not in(select PermissionID from RolePermissions where RoleID=" & l_RID & ") " & _
                                  "ORDER BY PC.Description, P.Description;"
              dst = MyCommon.LRT_Select
              If (dst.Rows.Count > 0) Then
                For Each row In dst.Rows
                  If (BannersDisabled AndAlso MyCommon.NZ(row.Item("CategoryID"), -1) = 8) Then
                    ' skip adding the banner permissions if banners are not enabled
                  Else
                    If (CategoryName <> Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("CategoryPhraseID"), 0), LanguageID)) Then
                      If (CategoryName <> "") Then
                        Send("</optgroup>")
                      End If
                      Send("   <optgroup label=""" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("CategoryPhraseID"), 0), LanguageID) & """>  ")
                    End If
                    If Not (IsExcludedPermission(MyCommon.NZ(row.Item("PermissionID"), "-1").ToString, ExcludedPermissions, InstalledEngineCt)) Then
                      Send("    <option value=""" & row.Item("PermissionID") & """>" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID) & "</option>")
                    End If
                    CategoryName = Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("CategoryPhraseID"), 0), LanguageID)
                  End If
                Next
                Send("</optgroup>")
              End If
            %>
          </select>
        </div>
        <br clear="left" />
        <br class="zero" />
      </div>
      <% If (Request.QueryString("role") <> "" and MyCommon.Fetch_SystemOption(90)) Then%>
        <%
          l_RID = MyCommon.Extract_Val(Request.QueryString("role"))
          MyCommon.QueryStr = "select ExtRoleName from Adminroles with (NoLock) where RoleID=" & l_RID & ";"
          rst = MyCommon.LRT_Select
          if rst.Rows.Count > 0 then
            sExtName = MyCommon.NZ(rst.Rows(0).Item("ExtRoleName"),"")
          end if
        %>
        <br class="half" />
        <div class="box" id="extrole1">
          <h2>
            <span>
              <% Sendb(Copient.PhraseLib.Lookup("roles.editextname", LanguageID))%>
            </span>
          </h2>
          <label for="name"><% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>:</label><br />
          <input type="text" id="Text1" name="extname" class="long" maxlength="100" value="<% Sendb(sExtName)%>" />
          <input type="submit" class="regular" id="save" name="save" value="<% Sendb(Copient.PhraseLib.Lookup("term.save", LanguageID)) %>" /><br />
          <hr class="hidden" />
        </div>
      <% End If%>
    <% End If%>
    </div>
    <br clear="all" />
  </div>
</form>
<script runat="server">
  ' Only Exclude a permission if it is excluded from ALL engines.
  Function IsExcludedPermission(ByVal PermissionID As String, ByRef ExcludedPermissions As Hashtable, ByVal InstalledEngineCt As Integer) As Boolean
    
    Dim ExcludedEngineCt As Integer = -1
    Dim Excluded As Boolean
    
    If (ExcludedPermissions.ContainsKey(PermissionID)) Then
      If Not Integer.TryParse(ExcludedPermissions.Item(PermissionID), ExcludedEngineCt) Then
        Excluded = False
      Else
                
        Excluded = (ExcludedEngineCt = InstalledEngineCt)
      End If
    Else
      Excluded = False
    End If
    
    Return Excluded
  End Function
  
</script>
<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (Logix.UserRoles.AccessNotes) Then
      Send_Notes(25, 0, AdminUserID)
    End If
  End If
done:
  Send_BodyEnd("frmRoles", "name")
  MyCommon.Close_LogixRT()
  Logix = Nothing
  MyCommon = Nothing
%>
