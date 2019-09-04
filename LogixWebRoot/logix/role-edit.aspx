﻿<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.CommonInc" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%
  ' *****************************************************************************
  ' * FILENAME: role-edit.aspx 
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
  Dim RoleID As Long = 0
  Dim RoleName As String = ""
  Dim ExtRoleName As String = ""
  Dim OptView As Integer = 1
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim dst As DataTable
  Dim dst2 As DataTable
  Dim rst As DataTable
  Dim rstAssociated As DataTable = Nothing
  Dim row As DataRow
  Dim bSave As Boolean
  Dim bDelete As Boolean
  Dim bCreate As Boolean
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim BannersDisabled As Boolean = False
  Dim BannersEnabled As Boolean = False
  Dim i As Integer = 0
  Dim HasAssociatedUsers As Boolean = False
  Dim FullUserName As String = ""
  Dim CategoryName As String = ""
  Dim InstalledEnginePKIDs As String = ""
  Dim InstalledEngineCt As Integer = -1
  Dim IncludedPermissions As New Hashtable()
  
  Dim HasAddCustPerm As Boolean = False
  Dim HasRemCustPerm As Boolean = False
  Dim DroppingAddCustPerm As Boolean = False
  Dim DroppingRemCustPerm As Boolean = False
  Dim InvalidRemoval As Boolean = False
  Dim tmpRoles() As String
  Dim z As Integer = 0
  Dim tmpRoles2() As String
  Dim w As Integer = 0
  Dim PointsAdjustmentLimit As Integer = 0
  Dim StoredValueAdjustmentLimit As Integer = 0
  Dim LimitPeriod As Integer = 0
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "role-edit.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  Try
    ' fill in if it was a get method
    If Request.RequestType = "GET" Then
      RoleID = IIf(Request.QueryString("RoleID") = "", 0, MyCommon.Extract_Val(Request.QueryString("RoleID")))
      If Request.QueryString("save") = "" Then
        bSave = False
      Else
        bSave = True
      End If
      If Request.QueryString("delete") = "" Then
        bDelete = False
      Else
        bDelete = True
      End If
      If Request.QueryString("mode") = "" Then
        bCreate = False
      Else
        bCreate = True
      End If
    Else
      RoleID = IIf(Request.Form("RoleID") = "", 0, MyCommon.Extract_Val(Request.Form("RoleID")))
      'If the role ID is not found in the form then try the query string
      If RoleID = 0 Then RoleID = IIf(Request.QueryString("RoleID") = "", 0, MyCommon.Extract_Val(Request.QueryString("RoleID")))
      If Request.Form("save") = "" Then
        bSave = False
      Else
        bSave = True
      End If
      If Request.Form("delete") = "" Then
        bDelete = False
      Else
        bDelete = True
      End If
      If Request.Form("mode") = "" Then
        bCreate = False
      Else
        bCreate = True
      End If
    End If
    
    
    BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
    BannersDisabled = (MyCommon.Fetch_SystemOption(66) <> "1")
    
    'Find all installed engines/subengines
    MyCommon.QueryStr = "select PKID as EngineSubTypePKID from PromoEngineSubTypes with (NoLock) " & _
                        "where Installed=1;"
    dst = MyCommon.LRT_Select
    If (dst.Rows.Count > 0) Then
      InstalledEngineCt = dst.Rows.Count
      InstalledEnginePKIDs = ""
      For Each row In dst.Rows
        If (InstalledEnginePKIDs <> "") Then InstalledEnginePKIDs &= ","
        InstalledEnginePKIDs &= row.Item("EngineSubTypePKID")
      Next
    End If
    
    'Load all the included permissions into a hashtable
	
	If MyCommon.Fetch_SystemOption(263) = "1" Then
      MyCommon.QueryStr = "select distinct PermissionID, Count(PermissionID) as EnginesIncluded " & _
                        "from PromoEnginePermissions with (NoLock) " & _
                        "where EngineSubTypePKID in (" & InstalledEnginePKIDs & ") " & _
                        "group by PermissionID;"
    Else
	  MyCommon.QueryStr = "select distinct PermissionID, Count(PermissionID) as EnginesIncluded " & _
                        "from PromoEnginePermissions with (NoLock) " & _
                        "where EngineSubTypePKID in (" & InstalledEnginePKIDs & ") " & _
						"and PermissionID <> 250 " & _
                        "group by PermissionID;"
	End If
	dst = MyCommon.LRT_Select
    If (dst.Rows.Count > 0) Then
      For Each row In dst.Rows
        IncludedPermissions.Add(MyCommon.NZ(row.Item("PermissionID"), "-1").ToString, MyCommon.NZ(row.Item("EnginesIncluded"), -1))
      Next
    End If
    
    Send_HeadBegin("term.roles", , RoleID)
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts()
%>
<script type="text/javascript">
  function toggleDropdown() {
    if (document.getElementById("actionsmenu") != null) {
      bOpen = (document.getElementById("actionsmenu").style.visibility != 'visible')
      if (bOpen) {
        document.getElementById("actionsmenu").style.visibility = 'visible';
        document.mainform.actions.value = '<% Sendb(Copient.PhraseLib.Lookup("term.actions", LanguageID)) %>▲';
      } else {
        document.getElementById("actionsmenu").style.visibility = 'hidden';
        document.mainform.actions.value = '<% Sendb(Copient.PhraseLib.Lookup("term.actions", LanguageID)) %>▼';
      }
    }
  }
 function limits_onchange() {
    var lperiod = document.getElementById("LimitPeriod");  
    var points = document.getElementById("PointsLimit");
	var svlimit = document.getElementById("StoredValueLimit");
	var alertmsg = '';
	
    if(!lperiod.value.match(/^\d+$/)) {
	   alert('<%Sendb(Copient.PhraseLib.Lookup("adjustmentLimit-LimitPeriod.positivevalue", LanguageID)) %>');
	   //lperiod.value = 1;
     } else {
       if (lperiod.value <= 0) {
		  alert('<%Sendb(Copient.PhraseLib.Lookup("adjustmentLimit-LimitPeriod.invalid-period", LanguageID)) %>');
		  //lperiod.value = 1;
		}	 
	 }
    if(points.value.match(/^\d+$/)) {
		if (points.value > 2147483647) {
		// means input value is greater than the Integer max. value		
		  alertmsg += '<%Sendb(Copient.PhraseLib.Lookup("term.points", LanguageID)) %>'
		  alertmsg += ' <%Sendb(Copient.PhraseLib.Lookup("adjustmentLimit.error", LanguageID)) %>'		
		  alertmsg += ' 2147483647.'
		  alert(alertmsg);
		  //points.value = 0;
		}
	 } else {
	   alert('<%Sendb(Copient.PhraseLib.Lookup("adjustmentLimit-Points.positivevalue", LanguageID)) %>');
	   //points.value = 0;
     }
    if(svlimit.value.match(/^\d+$/)) {
		if (svlimit.value > 2147483647) {
		// means input value is greater than the Integer max. value
		  alertmsg += '<%Sendb(Copient.PhraseLib.Lookup("term.storedvalue", LanguageID)) %>'
		  alertmsg += ' <%Sendb(Copient.PhraseLib.Lookup("adjustmentLimit.error", LanguageID)) %>'		
		  alertmsg += ' 2147483647.'
		  alert(alertmsg);
		  //svlimit.value = 0;
		}
	 } else {
	   alert('<%Sendb(Copient.PhraseLib.Lookup("adjustmentLimit-Storedvalue.positivevalue", LanguageID)) %>');
	   //svlimit.value = 0;
     }	 
  }
  
  var EnableAccessSelection='<%=MyCommon.Fetch_SystemOption(261)%>'
  function AutoSelectAccess(prmID) {

        //Logix.UserRoles.CreateProductGroups=78
        //Logix.UserRoles.DeleteProductGroups=79
        //Logix.UserRoles.EditProductGroups=80
        //Logix.UserRoles.AccessProductGroups=14
        var permissions = new Array(78,79,80,14);
        if (EnableAccessSelection == '1' && (prmID == permissions[0] || prmID == permissions[1] || prmID == permissions[2]) )		
        {
            if (document.getElementById(permissions[3]) != null) {
                accessElt = document.getElementById(permissions[3]);
                accessElt.setAttribute('selected', 'selected');
            }
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
  
  If (Logix.UserRoles.EditSystemConfiguration = False) Then
    Send_Denied(1, "perm.admin-configuration")
    GoTo done
  End If
  
  If (Request.QueryString("new") <> "") Then
    Response.Redirect("role-edit.aspx")
  End If
  
  If bSave Then
    'Saving a role (either new or existing)
    RoleName = MyCommon.NZ(Request.QueryString("RoleName"), "")
    RoleName = MyCommon.Parse_Quotes(Logix.TrimAll(RoleName))
    ExtRoleName = MyCommon.NZ(Request.QueryString("ExtRoleName"), "")
    ExtRoleName = MyCommon.Parse_Quotes(Logix.TrimAll(ExtRoleName))
	 If MyCommon.Fetch_SystemOption(224) Then
	If Not IsNumeric(Request.QueryString("LimitPeriod")) Then
      infoMessage = Copient.PhraseLib.Lookup("adjustmentLimit-LimitPeriod.positivevalue", LanguageID)
	Else
	  If MyCommon.NZ(Request.QueryString("LimitPeriod"), 0) <= 0 Then
	    infoMessage = Copient.PhraseLib.Lookup("adjustmentLimit-LimitPeriod.invalid-period", LanguageID)
	  Else
	    LimitPeriod = MyCommon.NZ(Request.QueryString("LimitPeriod"), 1)
	  End If
	End If	
	If Not IsNumeric(Request.QueryString("PointsLimit")) Then
      infoMessage & = Copient.PhraseLib.Lookup("adjustmentLimit-Points.positivevalue", LanguageID)
	Else
	  If MyCommon.NZ(Request.QueryString("PointsLimit"), 0) > 2147483647 Then
        infoMessage = Copient.PhraseLib.Lookup("term.points", LanguageID)  
		infoMessage & = " " & Copient.PhraseLib.Lookup("adjustmentLimit.error", LanguageID)
		infoMessage & = " 2147483647."
	  ElseIf MyCommon.NZ(Request.QueryString("PointsLimit"), 0) <= 0 Then
	    infoMessage & = Copient.PhraseLib.Lookup("adjustmentLimit-Points.positivevalue", LanguageID)	  
	  Else
        PointsAdjustmentLimit = MyCommon.NZ(Request.QueryString("PointsLimit"), 0)
	  End If
	End If
	If Not IsNumeric(Request.QueryString("StoredValueLimit")) Then
      infoMessage & = Copient.PhraseLib.Lookup("adjustmentLimit-Storedvalue.positivevalue", LanguageID)	
	Else
	  If MyCommon.NZ(Request.QueryString("StoredValueLimit"), 0) > 2147483647 Then
        infoMessage = Copient.PhraseLib.Lookup("term.storedvalue", LanguageID)  
		infoMessage & = " " & Copient.PhraseLib.Lookup("adjustmentLimit.error", LanguageID)
		infoMessage & = " 2147483647."
	  ElseIf MyCommon.NZ(Request.QueryString("StoredValueLimit"), 0) <= 0 Then
	    infoMessage & = Copient.PhraseLib.Lookup("adjustmentLimit-Storedvalue.positivevalue", LanguageID)			
	  Else
        StoredValueAdjustmentLimit = MyCommon.NZ(Request.QueryString("StoredValueLimit"), 0)
	  End If
	End If
    End If
    If RoleName = "" Then
      infoMessage = Copient.PhraseLib.Lookup("roles.noname", LanguageID)
    Else
      If (RoleID = 0) Then
        'New role
        MyCommon.QueryStr = "SELECT RoleID FROM AdminRoles with (NoLock) WHERE RoleName='" & RoleName & "';"
        dst = MyCommon.LRT_Select
        If (dst.Rows.Count > 0) Then
          infoMessage = Copient.PhraseLib.Lookup("roles.nameused", LanguageID)
        Else
          MyCommon.QueryStr = "INSERT INTO AdminRoles with (RowLock) (RoleName, ExtRoleName, PhraseID) VALUES (N'" & RoleName & "', N'" & ExtRoleName & "', NULL);"
          MyCommon.LRT_Execute()
          MyCommon.QueryStr = "SELECT RoleID FROM AdminRoles with (NoLock) WHERE RoleName='" & RoleName & "';"
          dst = MyCommon.LRT_Select
          RoleID = MyCommon.NZ(dst.Rows(0).Item("RoleID"), 0)
          MyCommon.Activity_Log(20, RoleID, AdminUserID, Copient.PhraseLib.Lookup("history.role-create", LanguageID))
          Response.Status = "301 Moved Permanently"
          Response.AddHeader("Location", "role-edit.aspx?RoleID=" & RoleID)
        End If
      Else
        'Existing role
        MyCommon.QueryStr = "SELECT RoleID FROM AdminRoles with (NoLock) WHERE RoleName='" & RoleName & "' and RoleID<>" & RoleID & ";"
        dst = MyCommon.LRT_Select
        If infoMessage = "" Then		
        If (dst.Rows.Count > 0) Then
          infoMessage = Copient.PhraseLib.Lookup("roles.nameused", LanguageID)
        Else
          MyCommon.QueryStr = "UPDATE AdminRoles with (RowLock) set RoleName='" & MyCommon.Parse_Quotes(RoleName) & "', ExtRoleName='" & MyCommon.Parse_Quotes(ExtRoleName) & "' " & _
                              "where RoleID=" & RoleID & ";"
          MyCommon.LRT_Execute()
		    If MyCommon.Fetch_SystemOption(224) Then
              MyCommon.QueryStr = "SELECT RoleID FROM AdminRoleAdjustmentLimits with (NoLock) WHERE RoleID=" & RoleID & ";"
              dst = MyCommon.LRT_Select
              If (dst.Rows.Count > 0) Then
                MyCommon.QueryStr = "UPDATE AdminRoleAdjustmentLimits with (RowLock) SET PointsAdjustmentLimit = " & PointsAdjustmentLimit & ", StoredValueAdjustmentLimit = " & StoredValueAdjustmentLimit & ", LimitPeriod = " & LimitPeriod & ", UpdatedByAdminID = " & AdminUserID & ", LastUpdate = GETDATE() WHERE RoleID=" & RoleID & ";"
                MyCommon.LRT_Execute()			
		      Else
                MyCommon.QueryStr = "INSERT INTO AdminRoleAdjustmentLimits with (RowLock) (RoleID, PointsAdjustmentLimit, StoredValueAdjustmentLimit, LimitPeriod, UpdatedByAdminID) VALUES (" & RoleID & ", " & PointsAdjustmentLimit & "," & StoredValueAdjustmentLimit & " ," & LimitPeriod & " ," & AdminUserID & ");"
                MyCommon.LRT_Execute()
			  End If
		    End If
          MyCommon.Activity_Log(20, RoleID, AdminUserID, Copient.PhraseLib.Lookup("history.role-edit", LanguageID))
          Response.Status = "301 Moved Permanently"
          Response.AddHeader("Location", "role-edit.aspx?RoleID=" & RoleID)
        End If
		End If
      End If
    End If
  ElseIf bDelete Then
    'Deleting a role
    If RoleID = 1 Then
      infoMessage = Copient.PhraseLib.Lookup("roles.nodelete", LanguageID)
    Else
      MyCommon.QueryStr = "SELECT RoleID FROM AdminUserRoles WITH (NoLock) WHERE RoleID=" & RoleID & ";"
      dst = MyCommon.LRT_Select
      If dst.Rows.Count > 0 Then
        infoMessage = Copient.PhraseLib.Lookup("roles.inuse", LanguageID)
      Else
        MyCommon.QueryStr = "SELECT CustomerGroupID FROM CustomerGroups WITH (NoLock) WHERE Deleted=0 AND EditControlTypeID=3 AND RoleID=" & RoleID & ";"
        dst = MyCommon.LRT_Select
        If dst.Rows.Count > 0 Then
          infoMessage = Copient.PhraseLib.Lookup("roles.cgroup-nodelete", LanguageID) & " ("
          For Each row In dst.Rows
            i += 1
            infoMessage &= "<a href=""cgroup-edit.aspx?CustomerGroupID=" & MyCommon.NZ(row.Item("CustomerGroupID"), 0) & """>" & MyCommon.NZ(row.Item("CustomerGroupID"), 0) & "</a>"
            If (dst.Rows.Count > 1) AndAlso (i < dst.Rows.Count) Then
              infoMessage &= ", "
            End If
          Next
          infoMessage &= ")."
        Else
          MyCommon.QueryStr = "DELETE FROM AdminRoles with (RowLock) WHERE RoleID=" & RoleID & ";"
          MyCommon.LRT_Execute()
		  If MyCommon.Fetch_SystemOption(224) Then
            MyCommon.QueryStr = "DELETE FROM AdminRoleAdjustmentLimits with (RowLock) WHERE RoleID=" & RoleID & ";"
            MyCommon.LRT_Execute()
		  End If	
          MyCommon.Activity_Log(20, RoleID, AdminUserID, Copient.PhraseLib.Lookup("history.role-delete", LanguageID))
          Response.Status = "301 Moved Permanently"
          Response.AddHeader("Location", "role-list.aspx")
        End If
      End If
    End If
  ElseIf (Request.QueryString("selectperm") <> "") Then
    If Not (Request.QueryString("pa") = "") Then
      tmpRoles = Request.QueryString("pa").Split(",")
      For z = 0 To tmpRoles.GetUpperBound(0)
        MyCommon.QueryStr = "Select RoleID from RolePermissions with (NoLock) where RoleID=" & RoleID & " and PermissionID=" & tmpRoles(z)
        rst = MyCommon.LRT_Select()
        If rst.Rows.Count = 0 Then
          MyCommon.QueryStr = "INSERT into RolePermissions with (RowLock) (PermissionID,RoleID) values(" & tmpRoles(z) & "," & RoleID & ");"
          MyCommon.LRT_Execute()
        End If
      Next
      MyCommon.Activity_Log(20, RoleID, AdminUserID, Copient.PhraseLib.Lookup("history.role-add", LanguageID))
    End If
  ElseIf (Request.QueryString("removeperm") <> "") And Not (Request.QueryString("ps") = "") Then
    tmpRoles2 = Request.QueryString("ps").Split(",")
    'Check to see if the role is used by a customer group; if so, it must continue to contain either
    'the "Add customer to groups" (#50) or "Delete customer from groups" (#49) permissions.
    MyCommon.QueryStr = "SELECT CustomerGroupID FROM CustomerGroups WITH (NoLock) WHERE Deleted=0 AND EditControlTypeID=3 AND RoleID=" & RoleID & ";"
    dst = MyCommon.LRT_Select
    If dst.Rows.Count > 0 Then
      MyCommon.QueryStr = "SELECT PermissionID FROM RolePermissions WITH (NoLock) WHERE RoleID=" & RoleID & " AND PermissionID=49;"
      dst2 = MyCommon.LRT_Select
      If dst2.Rows.Count > 0 Then
        HasRemCustPerm = True
      End If
      MyCommon.QueryStr = "SELECT PermissionID FROM RolePermissions WITH (NoLock) WHERE RoleID=" & RoleID & " AND PermissionID=50;"
      dst2 = MyCommon.LRT_Select
      If dst2.Rows.Count > 0 Then
        HasAddCustPerm = True
      End If
      For w = 0 To tmpRoles2.GetUpperBound(0)
        If (tmpRoles2(w) = 49) Then
          DroppingRemCustPerm = True
        ElseIf (tmpRoles2(w) = 50) Then
          DroppingAddCustPerm = True
        End If
      Next
      If (HasRemCustPerm AndAlso HasAddCustPerm) Then
        If (DroppingRemCustPerm AndAlso DroppingAddCustPerm) Then
          InvalidRemoval = True
        End If
      ElseIf HasRemCustPerm AndAlso DroppingRemCustPerm Then
        InvalidRemoval = True
      ElseIf HasAddCustPerm AndAlso DroppingAddCustPerm Then
        InvalidRemoval = True
      End If
    End If
    If InvalidRemoval Then
      infoMessage = Copient.PhraseLib.Lookup("roles.cgroup-noremove", LanguageID)
    Else
      For w = 0 To tmpRoles2.GetUpperBound(0)
        MyCommon.QueryStr = "DELETE from RolePermissions with (RowLock) where PermissionID=" & tmpRoles2(w) & " and RoleID=" & RoleID & ";"
        MyCommon.LRT_Execute()
      Next
      MyCommon.Activity_Log(20, RoleID, AdminUserID, Copient.PhraseLib.Lookup("history.role-remove", LanguageID))
    End If
  End If
  
  If Not bCreate Then
    ' No one clicked anything
    MyCommon.QueryStr = "select RoleID, RoleName, DisplayOrder, PhraseID, ExtRoleName from AdminRoles with (NoLock) " & _
                        "where RoleID=" & RoleID & ";"
    rst = MyCommon.LRT_Select()
    If (rst.Rows.Count > 0) Then
      RoleName = MyCommon.NZ(rst.Rows(0).Item("RoleName"), "")
      ExtRoleName = MyCommon.NZ(rst.Rows(0).Item("ExtRoleName"), "")
    ElseIf (RoleID > 0) Then
      Send("")
      Send("<div id=""intro"">")
      Send("  <h1 id=""title"">" & Copient.PhraseLib.Lookup("term.role", LanguageID) & " #" & RoleID & "</h1>")
      Send("</div>")
      Send("<div id=""main"">")
      Send("  <div id=""infobar"" class=""red-background"">")
      Send("    " & Copient.PhraseLib.Lookup("error.deleted", LanguageID))
      Send("  </div>")
      Send("</div>")
      GoTo done
    End If
  End If
  
  If Request.QueryString("OptView") <> "" Then
    OptView = MyCommon.Extract_Val(Request.QueryString("OptView"))
  End If
  
  MyCommon.QueryStr = "select AU.AdminUserID, AU.FirstName, AU.LastName, AU.UserName from AdminUsers as AU with (NoLock) " & _
                      "where AdminUserID in (select AdminUserID from AdminUserRoles where RoleID=" & RoleID & ") " & _
                      "order by LastName, FirstName;"
  rstAssociated = MyCommon.LRT_Select
  HasAssociatedUsers = (rstAssociated.Rows.Count > 0)
%>
<form action="#" id="mainform" name="mainform">
<input type="hidden" id="RoleID" name="RoleID" value="<% Sendb(RoleID) %>" />
<div id="intro">
  <%
    Sendb("<h1 id=""title"">")
    If RoleID = 0 Then
      Sendb(Copient.PhraseLib.Lookup("term.newrole", LanguageID))
    Else
      Sendb(Copient.PhraseLib.Lookup("term.role", LanguageID) & " #" & RoleID & ": " & MyCommon.TruncateString(RoleName, 40))
    End If
    Send("</h1>")
    Send("<div id=""controls"">")
    If (RoleID = 0) Then
      If (Logix.UserRoles.EditRoles) Then
        Send_Save()
      End If
    Else
      If (Logix.UserRoles.EditRoles) Then
        Send("<input type=""button"" class=""regular"" id=""actions"" name=""actions"" value=""" & Copient.PhraseLib.Lookup("term.actions", LanguageID) & " ▼"" onclick=""toggleDropdown();"" />")
        Send("<div class=""actionsmenu"" id=""actionsmenu"">")
        Send_Save()
        Send_Delete()
        Send_New()
        Send("</div>")
      End If
      If MyCommon.Fetch_SystemOption(75) Then
        If (Logix.UserRoles.AccessNotes) Then
          Send_NotesButton(25, RoleID, AdminUserID)
        End If
      End If
    End If
    Send("</div>")
  %>
</div>
<div id="main">
  <%
    If (infoMessage <> "") Then
      Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
    End If
  %>
  <div id="column1">
    <div class="box" id="identification">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.identification", LanguageID))%>
        </span>
      </h2>
      <%
        If (RoleName Is Nothing) Then
          RoleName = ""
        End If
        Send("<label for=""RoleName"">" & Copient.PhraseLib.Lookup("term.name", LanguageID) & ":</label><br />")
        Send("<input type=""text"" class=""longest"" id=""RoleName"" name=""RoleName"" maxlength=""100"" value=""" & RoleName.Replace("""", "&quot;") & """" & IIf(RoleID = 1, " readonly=""readonly""", "") & " />")
        Send("<br />")
        Send("<br class=""half"" />")
        If MyCommon.Fetch_SystemOption(90) Then
          Send("<label for=""ExtRoleName"">" & Copient.PhraseLib.Lookup("term.externalname", LanguageID) & ":</label><br />")
          Send("<input type=""text"" class=""longest"" id=""ExtRoleName"" name=""ExtRoleName"" maxlength=""100"" value=""" & ExtRoleName.Replace("""", "&quot;") & """ />")
          Send("<br />")
          Send("<br class=""half"" />")
        Else
          Send("<input type=""hidden"" class=""longest"" id=""ExtRoleName"" name=""ExtRoleName"" maxlength=""100"" value=""" & ExtRoleName.Replace("""", "&quot;") & """ />")
        End If
      %>
      <hr class="hidden" />
    </div>
    
    <div class="box" id="roleedit"<%Sendb(IIf(RoleID > 0, "", " style=""display:none;"""))%>>
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.permissions", LanguageID))%>
        </span>
      </h2>
      <div style="position:relative;float:right;top:-26px;">
        <input type="radio" id="optView1" name="optView" value="0"<%Sendb(IIf(OptView=0, " checked=""checked""", " onclick=""document.mainform.submit();"""))%> /><label for="optView1" style="font-size:10px;position:relative;top:-2px;"><%Sendb(Copient.PhraseLib.Lookup("term.ByName", LanguageID)) %></label>
        <input type="radio" id="optView2" name="optView" value="1"<%Sendb(IIf(OptView=1, " checked=""checked""", " onclick=""document.mainform.submit();"""))%> /><label for="optView2" style="font-size:10px;position:relative;top:-2px;"><%Sendb(Copient.PhraseLib.Lookup("term.ByCategory", LanguageID))%></label>
      </div>
      <%
        'Available list...
        If OptView = 0 Then
          '...by name
          MyCommon.QueryStr = "Select distinct P.PermissionID, P.Description, P.PhraseID, P.CategoryID from Permissions P with (NoLock)" & _
                              "where not exists (select PermissionID from RolePermissions RP with (NoLock) where P.PermissionID=RP.PermissionID and RoleID=" & RoleID & ") " & _
                              "order by P.Description;"
          dst = MyCommon.LRT_Select
          Send("<label for=""pa""><b>" & Copient.PhraseLib.Lookup("term.available", LanguageID) & ":</b></label><br />")
          If MyCommon.Fetch_SystemOption(261) Then
              Send("<select class=""wideselector"" multiple=""multiple"" id=""pa"" name=""pa"" onClick=AutoSelectAccess(value) style=""height:160px;"">")
          Else
              Send("<select class=""wideselector"" multiple=""multiple"" id=""pa"" name=""pa""  style=""height:160px;"">")
          End If
          For Each row In dst.Rows
            If (IsIncludedPermission(MyCommon.NZ(row.Item("PermissionID"), "-1").ToString, IncludedPermissions, InstalledEngineCt)) Then
              If (BannersDisabled AndAlso MyCommon.NZ(row.Item("CategoryID"), -1) = 8) Then
                ' skip writing this banner option as banners are disabled
                      Else
                          'CLOUDSOL-2857:ACS 7.01-AMS6.2 Unable to save CM Settings, error message appears when save button is clicked.
                          If Not (row.Item("PermissionID") = 235 And RoleID = 1) Then
                              Send("    <option id=""" & row.Item("PermissionID") & """ value=""" & row.Item("PermissionID") & """>" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID) & "</option>")
                          End If
                      End If
                  End If
          Next
          Send("</select>")
        Else
          '...by category
          CategoryName = ""
          MyCommon.QueryStr = "select distinct P.PermissionID, PC.Description as CategoryName, PC.PhraseID as CategoryPhraseID, P.Description, P.PhraseID, P.CategoryID " & _
                              "from Permissions P with (NoLock) " & _
                              "Left Join PermissionCategories PC with (NoLock) on P.CategoryID=PC.CategoryID " & _
                              "where not exists (select PermissionID from RolePermissions RP where P.PermissionID=RP.PermissionID and RoleID=" & RoleID & ") " & _
                              "order by PC.Description, P.Description;"
          dst = MyCommon.LRT_Select
          Send("<label for=""pa""><b>" & Copient.PhraseLib.Lookup("term.available", LanguageID) & ":</b></label><br />")
          If MyCommon.Fetch_SystemOption(261) Then
              Send("<select class=""wideselector"" multiple=""multiple"" id=""pa"" name=""pa"" onClick=AutoSelectAccess(value) style=""height:160px;"">")
          Else
              Send("<select class=""wideselector"" multiple=""multiple"" id=""pa"" name=""pa""  style=""height:160px;"">")
          End If
          If (dst.Rows.Count > 0) Then
            For Each row In dst.Rows
              If (BannersDisabled AndAlso MyCommon.NZ(row.Item("CategoryID"), -1) = 8) OR
			    (NOT(IsIncludedPermission(MyCommon.NZ(row.Item("PermissionID"), "-1").ToString, IncludedPermissions, InstalledEngineCt))) Then
                'Skip adding the banner permissions if banners are not enabled and/or not included by permission
              Else
                If (CategoryName <> Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("CategoryPhraseID"), 0), LanguageID)) Then
                  If (CategoryName <> "") Then
                    Send("</optgroup>")
                  End If
                  Send("   <optgroup label=""" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("CategoryPhraseID"), 0), LanguageID) & """>  ")
                          End If
                          'CLOUDSOL-2857:ACS 7.01-AMS6.2 Unable to save CM Settings, error message appears when save button is clicked.
                          If Not (row.Item("PermissionID") = 235 And RoleID = 1) Then
                              Send("    <option id=""" & row.Item("PermissionID") & """ value=""" & row.Item("PermissionID") & """>" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID) & "</option>")
                          End If
                          CategoryName = Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("CategoryPhraseID"), 0), LanguageID)
                          End If
            Next
            Send("</optgroup>")
          End If
          Send("</select>")
        End If
      %>
      <div style="margin-left:64px;">
        <input type="submit" class="regular select" id="selectperm" name="selectperm" value="▼ <% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID))%>" title="<% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID))%>"<%Sendb(IIf(RoleID <= 1 OrElse Logix.UserRoles.EditRoles = False, " disabled=""disabled""", ""))%> />
        <input type="submit" class="regular deselect" id="removeperm" name="removeperm" value="<% Sendb(Copient.PhraseLib.Lookup("term.remove", LanguageID))%> ▲" title="<% Sendb(Copient.PhraseLib.Lookup("term.remove", LanguageID))%>"<%Sendb(IIf(RoleID <= 1 OrElse Logix.UserRoles.EditRoles = False, " disabled=""disabled""", ""))%> />
      </div>
      <%
        'Selected roles...
        If OptView = 0 Then
          '...by name
          MyCommon.QueryStr = "SELECT DISTINCT RP.PermissionID, P.Description, P.PhraseID, P.CategoryID " & _
                              "FROM RolePermissions AS RP with (NoLock) " & _
                              "LEFT JOIN Permissions AS P with (NoLock) ON RP.PermissionID=P.PermissionID " & _
                              "WHERE RoleID=" & RoleID & "ORDER BY Description;"
          rst = MyCommon.LRT_Select
          Send("<label for=""ps""><b>" & Copient.PhraseLib.Lookup("term.selected", LanguageID) & ":</b></label><br />")
          Send("<select class=""wideselector"" multiple=""multiple"" id=""ps"" name=""ps"" style=""height: 160px;"">")
          For Each row In rst.Rows
            If (BannersDisabled AndAlso MyCommon.NZ(row.Item("CategoryID"), -1) = 8) Then
              ' skip writing this banner option as banners are disabled
                  Else
                      Send("    <option value=""" & row.Item("PermissionID") & """>" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID) & "</option>")
                      End If
          Next
          Send("</select>")
        Else
          '...by category
          CategoryName = ""
          MyCommon.QueryStr = "SELECT RP.RoleID, RP.PermissionID, PC.Description as CategoryName, PC.PhraseID as CategoryPhraseID, P.Description, P.PhraseID, PC.CategoryID " & _
                              "FROM RolePermissions AS RP with (NoLock) " & _
                              "LEFT JOIN Permissions AS P with (NoLock) ON RP.PermissionID=P.PermissionID " & _
                              "LEFT JOIN PermissionCategories as PC on P.CategoryID=PC.CategoryID " & _
                              "WHERE RoleID=" & RoleID & " ORDER BY PC.Description, P.Description;"
          rst = MyCommon.LRT_Select
          Send("<label for=""ps""><b>" & Copient.PhraseLib.Lookup("term.selected", LanguageID) & ":</b></label><br />")
          Send("<select class=""wideselector"" multiple=""multiple"" id=""ps"" name=""ps"" style=""height: 160px;"">")
          If (rst.Rows.Count > 0) Then
            For Each row In rst.Rows
              If (BannersDisabled AndAlso MyCommon.NZ(row.Item("CategoryID"), -1) = 8) Then
                'Skip adding the banner permissions if banners are not enabled
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
          Send("</select>")
        End If
      %>
    </div>
  </div>
  
  <div id="gutter">
  </div>
  
  <div id="column2">
    <div class="box" id="users"<%if(RoleID = 0)then sendb(" style=""display: none;""") %>>
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.associatedusers", LanguageID))%>
        </span>
      </h2>
      <div class="boxscroll">
        <%
          If (RoleID > 0) Then
            If rstAssociated.Rows.Count > 0 Then
              Send(Copient.PhraseLib.Detokenize("roles.UsersCount", LanguageID, rstAssociated.Rows.Count) & "<br />")
              Send("<br class=""half"" />")
              For Each row In rstAssociated.Rows
                FullUserName = ""
                If (MyCommon.NZ(row.Item("FirstName"), "") <> "") Then
                  FullUserName &= MyCommon.NZ(row.Item("FirstName"), "")
                End If
                If (MyCommon.NZ(row.Item("FirstName"), "") <> "") AndAlso (MyCommon.NZ(row.Item("FirstName"), "") <> "") Then
                  FullUserName &= " "
                End If
                If (MyCommon.NZ(row.Item("LastName"), "") <> "") Then
                  FullUserName &= MyCommon.NZ(row.Item("LastName"), "")
                End If
                If FullUserName = "" Then
                  FullUserName = "&nbsp;"
                End If
                If (Logix.UserRoles.ViewOthersInfo = False) Then
                  Send(FullUserName)
                Else
                  Sendb(" <a href=""user-edit.aspx?UserID=" & row.Item("AdminUserID") & """>" & FullUserName & "</a>")
                End If
                Send("<br />")
              Next
            Else
              Send("                " & Copient.PhraseLib.Lookup("term.none", LanguageID) & "<br />")
            End If
          Else
            Send("                " & Copient.PhraseLib.Lookup("term.none", LanguageID) & "<br />")
          End If
        %>
      </div>
      <hr class="hidden" />
    </div>
    
    <div class="box" id="groups"<%if(RoleID = 0)then sendb(" style=""display: none;""") %>>
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.associatedgroups", LanguageID))%>
        </span>
      </h2>
      <div class="boxscroll">
        <%
          If (RoleID > 0) Then
            MyCommon.QueryStr = "select CustomerGroupID, Name from CustomerGroups with (NoLock) where Deleted=0 and EditControlTypeID in (0,1,3) and RoleID=" & RoleID & " order by Name;"
            rst = MyCommon.LRT_Select
            If rst.Rows.Count > 0 Then
              Send(Copient.PhraseLib.Detokenize("roles.CustomerGroupCount", LanguageID, rst.Rows.Count) & "<br />")
              Send("<br class=""half"" />")
              For Each row In rst.Rows
                If (Logix.UserRoles.AccessCustomerGroups) Then
                  Sendb(" <a href=""cgroup-edit.aspx?CustomerGroupID=" & MyCommon.NZ(row.Item("CustomerGroupID"), 0) & """>" & MyCommon.NZ(row.Item("Name"), "&nbsp;") & "</a>")
                Else
                  Sendb(" " & MyCommon.NZ(row.Item("Name"), "&nbsp;"))
                End If
                Send("<br />")
              Next
            Else
              Send("                " & Copient.PhraseLib.Lookup("term.none", LanguageID) & "<br />")
            End If
          Else
            Send("                " & Copient.PhraseLib.Lookup("term.none", LanguageID) & "<br />")
          End If
        %>
      </div>
      <hr class="hidden" />
    </div>
	<% If MyCommon.Fetch_SystemOption(224) Then %>
	<%
      MyCommon.QueryStr = "SELECT PointsAdjustmentLimit, StoredValueAdjustmentLimit, LimitPeriod FROM AdminRoleAdjustmentLimits WITH (NoLock) WHERE RoleID=" & RoleID & ";"
      dst = MyCommon.LRT_Select
      If (dst.Rows.Count > 0) Then 
        PointsAdjustmentLimit = MyCommon.NZ(dst.Rows(0).Item("PointsAdjustmentLimit"), 0)
        StoredValueAdjustmentLimit = MyCommon.NZ(dst.Rows(0).Item("StoredValueAdjustmentLimit"), 0)
		LimitPeriod = MyCommon.NZ(dst.Rows(0).Item("LimitPeriod"), 1) 
		If LimitPeriod <= 0 Then
		  LimitPeriod = 1
		End If	
      End If	
	%>
   <div id="column2">
   <div class="box" id="adjustmentlimits"<%Sendb(IIf(RoleID > 0, "", " style=""display:none;"""))%>>
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.adjustment", LanguageID))%>&nbsp;<% Sendb(System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToLower(Copient.PhraseLib.Lookup("term.limits", LanguageID)))%>
        </span>
      </h2>
      <%
		Send("<label for=""LimitPeriod"">" & Copient.PhraseLib.Lookup("term.limit", LanguageID) & " " & System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToLower(Copient.PhraseLib.Lookup("term.period", LanguageID)) & ":</label>")
        Send("<input type=""text"" class=""short"" id=""LimitPeriod"" name=""LimitPeriod"" maxlength=""3"" title='" & Copient.PhraseLib.Lookup("points-adjust.perprogrampercustomerperdaylimit", LanguageID) & "' onchange=""limits_onchange()"" value=""" & IIf(infoMessage <> "", Request.QueryString("LimitPeriod"), LimitPeriod) & """"" />")
	    Sendb(Copient.PhraseLib.Lookup("term.days.lc", LanguageID))
        Send("<br />")
        Send("<br class=""half"" />")
	    Send("<label for=""PointsLimit"">" & Copient.PhraseLib.Lookup("term.points", LanguageID) & ":</label>")
        Send("<input type=""text"" class=""short"" id=""PointsLimit"" name=""PointsLimit"" maxlength=""10"" onchange=""limits_onchange()"" value=""" & IIf(infoMessage <> "", Request.QueryString("PointsLimit"), PointsAdjustmentLimit) & """"" />")
        Send("<br />")
        Send("<br class=""half"" />")
        Send("<label for=""StoredValueLimit"">" & Copient.PhraseLib.Lookup("term.storedvalue", LanguageID) & ":</label>")
        Send("<input type=""text"" class=""short"" id=""StoredValueLimit"" name=""StoredValueLimit"" maxlength=""10"" onchange=""limits_onchange()"" value=""" & IIf(infoMessage <> "", Request.QueryString("StoredValueLimit"), StoredValueAdjustmentLimit) & """"" />")
        Send("<br />")		
        Send("<br class=""half"" />")
      %>
      <hr class="hidden" />
    </div>   
   </div>   	
  <% End If %>
  </div>
  <br clear="all" />
  
</div>
</form>

<script runat="server">
  Function IsIncludedPermission(ByVal PermissionID As String, ByRef IncludedPermissions As Hashtable, ByVal InstalledEngineCt As Integer) As Boolean
    Dim IncludedEngineCt As Integer = -1
    Dim Included As Boolean
    
    If (IncludedPermissions.ContainsKey(PermissionID)) Then
      Included = True
    Else
      Included = False
    End If
    
    Return Included
  End Function
  
  ' Only exclude a permission if it is excluded from ALL engines.
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

<script type="text/javascript">
  if (window.captureEvents) {
    window.captureEvents(Event.CLICK);
    window.onclick = handlePageClick;
  } else {
    document.onclick = handlePageClick;
  }
</script>
<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (Logix.UserRoles.AccessNotes) Then
      Send_Notes(25, RoleID, AdminUserID)
    End If
  End If
done:
Finally
  MyCommon.Close_LogixRT()
End Try
Send_BodyEnd("mainform", "RoleName")
MyCommon = Nothing
Logix = Nothing
%>
