<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.CommonInc" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%
  ' *****************************************************************************
  ' * FILENAME: terminal-lockgroup-edit.aspx 
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
  Dim TerminalLockingGroupId As Long
  Dim GroupName As String
  Dim GroupDescription As String
  Dim LastUpdate As String
  Dim EngineType As Integer
  Dim EngineName As String
  Dim Deleted As Boolean = False
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim dst As DataTable
  Dim rst As DataTable
  Dim rst2 As DataTable
  Dim rst3 As DataTable
  Dim row As DataRow
  Dim row2 As DataRow
  Dim bSave As Boolean
  Dim bDelete As Boolean
  Dim bCreate As Boolean
  Dim dtGroups As DataTable = Nothing
  Dim dtOffers As DataTable = Nothing
  Dim sQuery As String
  Dim FocusField As String = "ExtTerminalCode"
  Dim SizeOfData As Integer
  Dim ShowActionButton As Boolean = False
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim BannerID As Integer = 0
  Dim BannerName As String = ""
  Dim BannersEnabled As Boolean = False
  Dim BannerIDs As String = ""
  Dim BannerEngines As String = ""
  Dim DefaultEngineID As Integer = 0
  Dim rstSelected As DataTable
  Dim SelectSize As Integer = 6
  Dim NameValue As String = ""
  Dim PrevNameValue As String = ""
  Dim AllBannersPermission As Boolean = False

  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  MyCommon.AppName = "terminal-lockgroup-edit.aspx"
  Response.Expires = 0
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  Try
    ' fill in if it was a get method
    If Request.RequestType = "GET" Then
      TerminalLockingGroupId = Request.QueryString("TerminalLockingGroupId")
      GroupName = Logix.TrimAll(Request.QueryString("GroupName"))
      GroupDescription = Request.QueryString("GroupDescription")
      BannerID = MyCommon.Extract_Val(Request.QueryString("BannerID"))
      EngineType = Request.QueryString("EngineID")
      If (Request.QueryString("terminals-add1") <> "" And Request.QueryString("terminals-available") <> "") Then
        Dim s As String = Request.QueryString("terminals-available")
        Dim a() As String
        Dim j As Integer
    
        a = s.Split(",")
        For j = 0 To a.GetUpperBound(0)
          MyCommon.QueryStr = "Update TerminalTypes with (RowLock) set LockingGroupId=" & TerminalLockingGroupId & ", LastUpdate=getdate() where TerminalTypeId=" & a(j) & ";"
          MyCommon.LRT_Execute()
          MyCommon.Activity_Log(21, a(j), AdminUserID, Copient.PhraseLib.Lookup("history.terminal-edit", LanguageID))
        Next
      ElseIf (Request.QueryString("terminals-rem1") <> "" And Request.QueryString("terminals-select") <> "") Then
        Dim s As String = Request.QueryString("terminals-select")
        Dim a() As String
        Dim j As Integer
    
        a = s.Split(",")
        For j = 0 To a.GetUpperBound(0)
          MyCommon.QueryStr = "Update TerminalTypes with (RowLock) set LockingGroupId=0, LastUpdate=getdate() where TerminalTypeId=" & a(j) & ";"
          MyCommon.LRT_Execute()
          MyCommon.Activity_Log(21, a(j), AdminUserID, Copient.PhraseLib.Lookup("history.terminal-edit", LanguageID))
        Next
      Else
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
      End If
    Else
      TerminalLockingGroupId = Request.Form("TerminalLockingGroupId")
      If TerminalLockingGroupId = 0 Then
        TerminalLockingGroupId = MyCommon.Extract_Val(Request.QueryString("TerminalLockingGroupId"))
      End If
      GroupName = Request.Form("GroupName")
      GroupDescription = Request.Form("GroupDescription")
      BannerID = MyCommon.Extract_Val(Request.Form("BannerID"))
      EngineType = Request.Form("EngineID")
      If (Request.Form("terminals-add1") <> "" And Request.Form("terminals-available") <> "") Then
        Dim s As String = Request.Form("terminals-available")
        Dim a() As String
        Dim j As Integer
    
        a = s.Split(",")
        For j = 0 To a.GetUpperBound(0)
          MyCommon.QueryStr = "Update TerminalTypes with (RowLock) set LockingGroupId=" & TerminalLockingGroupId & ", LastUpdate=getdate() where TerminalTypeId=" & a(j) & ";"
          MyCommon.LRT_Execute()
          MyCommon.Activity_Log(21, a(j), AdminUserID, Copient.PhraseLib.Lookup("history.terminal-edit", LanguageID))
        Next
      ElseIf (Request.Form("terminals-rem1") <> "" And Request.Form("terminals-select") <> "") Then
        Dim s As String = Request.Form("terminals-select")
        Dim a() As String
        Dim j As Integer
    
        a = s.Split(",")
        For j = 0 To a.GetUpperBound(0)
          MyCommon.QueryStr = "Update TerminalTypes with (RowLock) set LockingGroupId=0, LastUpdate=getdate() where TerminalTypeId=" & a(j) & ";"
          MyCommon.LRT_Execute()
          MyCommon.Activity_Log(21, a(j), AdminUserID, Copient.PhraseLib.Lookup("history.terminal-edit", LanguageID))
        Next
      Else
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
    End If
   
    
    BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
    
    Send_HeadBegin("term.terminallockgroups", , TerminalLockingGroupId)
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
    
    If (Request.QueryString("new") <> "") Then
      Response.Redirect("terminal-lockgroup-edit.aspx")
    End If
    
    MyCommon.QueryStr = "select EngineID from PromoEngines PE with (NoLock) where Installed=1 and DefaultEngine=1;"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      DefaultEngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), 0)
    End If
    
    If bSave Then
      If (TerminalLockingGroupId = 0) Then
        
        '**************************************
        ' For now only CPE engine, EngineId = 2
        '**************************************
        EngineType = 2
        
        If (GroupName = "") Then
          infoMessage = Copient.PhraseLib.Lookup("terminal-edit.noname", LanguageID)
        Else
          MyCommon.QueryStr = "SELECT TerminalLockingGroupID FROM TerminalLockingGroups with (NoLock) WHERE Name = '" & MyCommon.Parse_Quotes(GroupName) & "' AND Deleted=0 "
          If (BannersEnabled) Then
            MyCommon.QueryStr &= " and BannerID=" & BannerID
          End If
          rst = MyCommon.LRT_Select
          If (rst.Rows.Count > 0) Then
            infoMessage = Copient.PhraseLib.Lookup("terminal-edit.nameused", LanguageID)
          Else
            MyCommon.QueryStr = "dbo.pt_TerminalLockingGroups_Insert"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 100).Value = GroupName
            MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 1000).Value = GroupDescription
            MyCommon.LRTsp.Parameters.Add("@BannerID", SqlDbType.Int).Value = BannerID
            MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = EngineType
            MyCommon.LRTsp.Parameters.Add("@TerminalLockingGroupId", SqlDbType.Int).Direction = ParameterDirection.Output
            MyCommon.LRTsp.ExecuteNonQuery()
            TerminalLockingGroupId = MyCommon.LRTsp.Parameters("@TerminalLockingGroupId").Value
            MyCommon.Close_LRTsp()
            MyCommon.Activity_Log(41, TerminalLockingGroupId, AdminUserID, Copient.PhraseLib.Lookup("history.terminal-lockgroup-create", LanguageID))
          End If
        End If
      Else
        If (GroupName = "") Then
          infoMessage = Copient.PhraseLib.Lookup("terminal-edit.noname", LanguageID)
        Else
          MyCommon.QueryStr = "SELECT Name FROM TerminalLockingGroups with (NoLock) WHERE Name = '" & MyCommon.Parse_Quotes(GroupName) & "' AND Deleted=0 AND TerminalLockingGroupId <> " & TerminalLockingGroupId & " "
          If (BannersEnabled) Then
            MyCommon.QueryStr &= " and BannerID=" & BannerID
          Else
            '**************************************
            ' For now only CPE engine, EngineId = 2
            '**************************************
            EngineType = 2
          End If
          rst2 = MyCommon.LRT_Select
          If (rst2.Rows.Count > 0) Then
            infoMessage = Copient.PhraseLib.Lookup("terminal-edit.nameused", LanguageID)
          Else
            MyCommon.QueryStr = "dbo.pt_TerminalLockingGroups_Update"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@TerminalLockingGroupId", SqlDbType.Int).Value = TerminalLockingGroupId
            MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 100).Value = GroupName
            MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 1000).Value = GroupDescription
            MyCommon.LRTsp.ExecuteNonQuery()
            MyCommon.Activity_Log(41, TerminalLockingGroupId, AdminUserID, Copient.PhraseLib.Lookup("history.terminal-lockgroup-edit", LanguageID))
            MyCommon.Close_LRTsp()
          End If
        End If
      End If
    ElseIf bDelete Then
      MyCommon.QueryStr = "select distinct TerminalTypeId  from TerminalTypes with (NoLock)" & _
                          "where Deleted=0 and LockingGroupID=" & Request.QueryString("TerminalLockingGroupID")
      rst = MyCommon.LRT_Select
      If (rst.Rows.Count > 0) Then
        infoMessage = Copient.PhraseLib.Lookup("terminal-edit.inuse", LanguageID)
      Else
        MyCommon.QueryStr = "dbo.pt_TerminalLockingGroups_Delete"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@TerminalLockingGroupId", SqlDbType.Int).Value = TerminalLockingGroupId
        MyCommon.LRTsp.ExecuteNonQuery()
        MyCommon.Close_LRTsp()
        MyCommon.Activity_Log(41, TerminalLockingGroupId, AdminUserID, Copient.PhraseLib.Lookup("history.terminal-lockgroup-delete", LanguageID))
        Response.Status = "301 Moved Permanently"
        Response.AddHeader("Location", "terminal-lockgroup-list.aspx")
        TerminalLockingGroupId = 0
        GroupName = ""
        GroupDescription = ""
      End If
    End If
    
    LastUpdate = ""
    
    If Not bCreate Then
      ' no one clicked anything
      MyCommon.QueryStr = "select TLG.Name,TLG.Description, TLG.LastUpdate, TLG.Deleted, " & _
                          "PE.PhraseID, PE.EngineID as EngineID, TLG.BannerID, BAN.Name as BannerName from TerminalLockingGroups as TLG with (nolock) " & _
                          "left join PromoEngines as PE with (NoLock) on PE.EngineID=TLG.EngineID " & _
                          "left join Banners BAN with (NoLock) on TLG.BannerID = BAN.BannerID and BAN.Deleted=0 " & _
                          "where TerminalLockingGroupId=" & TerminalLockingGroupId
      rst = MyCommon.LRT_Select()
      If (rst.Rows.Count > 0) Then
        For Each row In rst.Rows
          If (GroupName = "") Then
            If Not row.Item("Name").Equals(System.DBNull.Value) Then
              GroupName = row.Item("Name")
            End If
          End If
          If (GroupDescription = "") Then
            If Not row.Item("Description").Equals(System.DBNull.Value) Then
              GroupDescription = row.Item("Description")
            End If
          End If
          If (LastUpdate = "") Then
            If Not row.Item("LastUpdate").Equals(System.DBNull.Value) Then
              LastUpdate = row.Item("LastUpdate")
            End If
          End If
          If row.Item("Deleted") Then
            Deleted = True
            infoMessage = Copient.PhraseLib.Lookup("terminal-edit.deleted", LanguageID)
          End If
          If (EngineName = "") Then
            If Not row.Item("PhraseID").Equals(System.DBNull.Value) Then
              EngineName = Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID)
            End If
          End If
          If (EngineType = 0) Then
            If Not row.Item("EngineID").Equals(System.DBNull.Value) Then
              EngineType = row.Item("EngineID")
            End If
          End If
          BannerID = MyCommon.NZ(rst.Rows(0).Item("BannerID"), 0)
          BannerName = MyCommon.NZ(rst.Rows(0).Item("BannerName"), Copient.PhraseLib.Lookup("term.multiple-banners", LanguageID))
        Next
      ElseIf (TerminalLockingGroupId > 0) Then
        Send("")
        Send("<div id=""intro"">")
        Send("    <h1 id=""title"">" & Copient.PhraseLib.Lookup("term.terminallockgroup", LanguageID) & " #" & TerminalLockingGroupId & "</h1>")
        Send("</div>")
        Send("<div id=""main"">")
        Send("    <div id=""infobar"" class=""red-background"">")
        Send("        " & Copient.PhraseLib.Lookup("error.deleted", LanguageID))
        Send("    </div>")
        Send("</div>")
        GoTo done
      End If
    End If
    
%>

<script type="text/javascript">
  var isIE =(navigator.appName.indexOf("Microsoft")!=-1) ? true : false;
  var isOpera = (navigator.appName.indexOf("Opera")!=-1) ? true : false;

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

    function handleBanners(bannerID) {
      var elem = document.getElementById("TermCodeSpan");
      var elemCode = document.getElementById("ExtTerminalCode");
      var engineID = -1;
            
      if (elem != null) {
       engineID = getBannerEngine(bannerID);
       var showTermCode = (engineID==2) ? false : true;
       elem.style.display = (showTermCode) ? "" : "none"; 
       if (elemCode != null) {
        elemCode.value = (showTermCode) ? elemCode.value : "";
       }
      }
    }
    
    function getBannerEngine(bannerID) {
      var index = -1;
      var engineID = -1;
      
      for (var i=0; i < bannerIDs.length && index==-1; i++) {
        if (bannerID == bannerIDs[i]) {
          index = i;
        }
      }
      
      if (bannerEngines.length > index) {
        engineID = bannerEngines[index]
      }
      
      return engineID;
    }
    
  // This is the function that refreshes the list after a keypress.
  // The maximum number to show can be limited to improve performance with
  // huge lists (1000s of entries).
  // The function clears the list, and then does a linear search through the
  // globally defined array and adds the matches back to the list.
  function handleKeyUp(maxNumToShow) {
    var selectObj, textObj, functionListLength;
    var i,  numShown;
    var searchPattern;
    var newOpt, optGp;
    var isAll = false;
    
    document.getElementById("terminals-available").size = "10";
    
    // Set references to the form elements
    selectObj = document.getElementById("terminals-available");
    textObj = document.forms[0].functioninput;
    
    // Remember the function list length for loop speedup
    functionListLength = functionlist.length;
    
    // Set the search pattern depending
    if(document.forms[0].functionradio1[0].checked == true) {
      searchPattern = "^"+textObj.value;
    } else {
      searchPattern = textObj.value;
    }
    searchPattern = cleanRegExpString(searchPattern);
    
    // Create a regulare expression
    
    re = new RegExp(searchPattern,"gi");
    // Clear the options list
    selectObj.length = 0;
    
    // Loop through the array and re-add matching options
    numShown = 0;
    for(i = 0; i < functionListLength; i++) {
      if(functionlist[i].search(re) != -1) {
        if (vallist2[i] != "") {
          isAll = (functionlist[i] == "<% Sendb(Copient.PhraseLib.Lookup("term.unassigned", LanguageID)) %>")
          var newOpt = document.createElement('OPTION');
          newOpt.value = vallist2[i];
          if (isIE) { newOpt.innerText = functionlist[i]}; 
          newOpt.text =  functionlist[i]; 
          <% If (BannersEnabled) Then %>
            if (!isOpera) {
              optGp = GetTerminalOptionGroup(grouplist[i], selectObj, isAll);
              if (optGp != null) {
                optGp.appendChild(newOpt);
                selectObj.appendChild(optGp);
              } else {
                selectObj[numShown] = newOpt
              }                
            } else {
              selectObj[numShown] = newOpt
            }
          <% Else %>
            selectObj[numShown] = new Option(newOpt.text, newOpt.value);
          <% End If %>
          
          if (isAll) {
            selectObj[numShown].style.fontWeight = 'bold';
            selectObj[numShown].style.color = 'brown';
          }
          numShown++;
        }                        
      }
      // Stop when the number to show is reached
      if(numShown == maxNumToShow) {
        break;
      }
    }
    // When options list whittled to one, select that entry
    if(selectObj.length == 1) {
      selectObj.options[0].selected = true;
    }
  }
  
  function GetTerminalOptionGroup(bannerName, elemSlct, isAll) {
    var elemGroup = null;
    
    if (bannerName == "") {
      elemGroup = document.getElementById("gpTermAllBanners");
      if (elemGroup == null) {
        elemGroup = document.createElement("OPTGROUP");
        if (isAll) {
          elemGroup.label = "<% Sendb(Copient.PhraseLib.Lookup("term.allbanners", LanguageID)) %>";
        } else {
          elemGroup.label = "<% Sendb(Copient.PhraseLib.Lookup("term.unassigned", LanguageID)) %>";
        }
        elemGroup.id = "gpTermAllBanners";
        elemSlct.appendChild(elemGroup);
      }
    } else {
      elemGroup = document.getElementById("gpTerm" + bannerName);
      if (elemGroup == null) {
        elemGroup = document.createElement("OPTGROUP");
        elemGroup.label = bannerName;
        elemGroup.id = "gpTerm" + bannerName;
        elemSlct.appendChild(elemGroup);
      }
    }
    return elemGroup;
  }
  
  function handleKeyDown(e, slctName) {
    var key = e.which ? e.which : e.keyCode;
    
    if (key == 40) {
      var elemSlct = document.getElementById(slctName);
      if (elemSlct != null) { elemSlct.focus(); }
    }
  }

  // Terminals Array 
  <%
    
    If (BannersEnabled) Then
      MyCommon.QueryStr = "select TT.TerminalTypeID, TT.Name, isnull(TLG.Name,'" & Copient.PhraseLib.Lookup("term.unassigned", LanguageID) & "') as GroupName from TerminalTypes TT with (NoLock) " & _ 
                          "left join TerminalLockingGroups TLG on TLG.TerminalLockingGroupID=TT.LockingGroupID and TLG.BannerId=TT.BannerID and TLG.Deleted=0 " & _
                          "where TT.Deleted=0 and TT.AnyTerminal=0 and TT.BannerID=" & BannerID & " and TT.EngineID=" & EngineType & " " & _
                          "and TT.TerminalTypeID not in " & _
                          " (select TerminalTypeID from TerminalTypes with (NoLock) where LockingGroupID=" & TerminalLockingGroupId & ") " & _
                          "order by GroupName;"
    Else
      MyCommon.QueryStr = "select TT.TerminalTypeID, TT.Name, isnull(TLG.Name,'" & Copient.PhraseLib.Lookup("term.unassigned", LanguageID) & "') as GroupName from TerminalTypes TT with (NoLock) " & _
                          "left join TerminalLockingGroups TLG on TLG.TerminalLockingGroupID=TT.LockingGroupID and TLG.EngineID=TT.EngineID and TLG.Deleted=0 " & _
                          "where TT.Deleted=0 and TT.AnyTerminal=0 and TT.EngineID=" & EngineType & " " & _
                          "and TT.TerminalTypeID not in " & _
                          " (select TerminalTypeID from TerminalTypes with (NoLock) where LockingGroupID=" & TerminalLockingGroupId & ") " & _
                          "order by GroupName;"
    End If
    
    rst = MyCommon.LRT_Select
    If (rst.rows.count>0)
      Sendb("var functionlist = Array(")
      For Each row In rst.Rows
        Sendb("""" & MyCommon.NZ(row.item("Name"), "").ToString().Replace("""", "\""") & """,")
      Next
      Send(""""");")

      Sendb("  var grouplist = Array(")
      For Each row In rst.Rows
        Sendb("""" & MyCommon.NZ(row.item("GroupName"), "").ToString().Replace("""", "\""") & """,")
      Next
      Send(""""");")
      
      Sendb("var vallist2 = Array(")
      For Each row In rst.Rows
        Sendb("""" & row.item("TerminalTypeID") & """,")
      Next
      Send(""""");")
    Else
      Sendb("var functionlist = Array(")
      Send("""" & "" & """);")

      Sendb("var vallist2 = Array(")
      Send("""" & "" & """);")

      Sendb("  var grouplist = Array(")
      Send("""" & "" & """);")
    End If
  %>
  
  function submitTerminals() {
    document.mainform.action = "terminal-lockgroup-edit.aspx#h01";
  }

</script>


<form action="#" id="mainform" name="mainform">
  <input type="hidden" id="TerminalLockingGroupId" name="TerminalLockingGroupId" value="<% Sendb(TerminalLockingGroupId) %>" />
  <div id="intro">
    <%
      
      Sendb("<h1 id=""title"">")
      If TerminalLockingGroupId = 0 Then
        Sendb(Copient.PhraseLib.Lookup("term.new-terminal-lockgroup", LanguageID))
      Else
        Dim NameTitle As String
        Sendb(Copient.PhraseLib.Lookup("term.terminallockgroup", LanguageID) & " #" & TerminalLockingGroupId & ": ")
        MyCommon.QueryStr = "SELECT Name FROM TerminalLockingGroups with (NoLock) WHERE TerminalLockingGroupId = " & TerminalLockingGroupId & ";"
        rst2 = MyCommon.LRT_Select
        If (rst2.Rows.Count > 0) Then
          NameTitle = MyCommon.NZ(rst2.Rows(0).Item("Name"), "")
        End If
        Sendb(MyCommon.TruncateString(NameTitle, 40))
      End If
      Sendb("</h1>")
    %>
    <div id="controls">
      <%
        If Not Deleted Then
          If (TerminalLockingGroupId = 0) Then
            If (Logix.UserRoles.EditTerminals) Then
              Send_Save()
            End If
          Else
            ShowActionButton = (Logix.UserRoles.EditTerminals)
            If (ShowActionButton) Then
              Send("<input type=""button"" class=""regular"" id=""actions"" name=""actions"" value=""" & Copient.PhraseLib.Lookup("term.actions", LanguageID) & " ▼"" onclick=""toggleDropdown();"" />")
              Send("<div class=""actionsmenu"" id=""actionsmenu"">")
              If (Logix.UserRoles.EditTerminals) Then
                Send_Save()
              End If
              If (Logix.UserRoles.EditTerminals) Then
                Send_Delete()
              End If
              If (Logix.UserRoles.EditTerminals) Then
                Send_New()
              End If
              Send("</div>")
            End If
            If MyCommon.Fetch_SystemOption(75) Then
              If (Logix.UserRoles.AccessNotes) Then
                Send_NotesButton(26, TerminalLockingGroupId, AdminUserID)
              End If
            End If
          End If
        End If
      %>
    </div>
  </div>
  <a name="h00" id="h00"></a>
  <div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <% If Deleted Then GoTo DeleteSkip%>
    <div id="column1">
      <div class="box" id="identification">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.identification", LanguageID))%>
          </span>
        </h2>
        <label for="TerminalName"  style="position:relative;"><% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>:</label><br />
        <%If (GroupName Is Nothing) Then GroupName = ""
          Sendb("<input type=""text"" class=""longest"" id=""GroupName"" name=""GroupName"" maxlength=""100"" value=""" & GroupName.Replace("""", "&quot;") & """ />")%>
        <br />
        <br class="half" />
        <label for="desc" style="position:relative;"><% Sendb(Copient.PhraseLib.Lookup("term.description", LanguageID))%>:</label><br />
        <textarea class="longest" cols="48" rows="3" id="GroupDescription" name="GroupDescription"><% Sendb(GroupDescription)%></textarea><br />
        <br class="half" />
        <%
          If (BannersEnabled) Then
            If (TerminalLockingGroupId = 0) Then
              MyCommon.QueryStr = "select distinct BAN.BannerID, BAN.Name, BE.EngineID from Banners BAN with (NoLock) " & _
                                   "inner join AdminUserBanners AUB with (NoLock) on AUB.BannerID = BAN.BannerID " & _
                                   "inner join BannerEngines BE with (NoLock) on BE.BannerID = BAN.BannerID " & _
                                   "where BAN.Deleted=0 and BAN.AllBanners=0 and BE.EngineId=2 and AdminUserID = " & AdminUserID & " order by BAN.Name;"
              rst = MyCommon.LRT_Select
              Send("<label for=""BannerID"">" & Copient.PhraseLib.Lookup("term.banner", LanguageID) & ":</label><br />")
              Send("<select class=""longest"" name=""BannerID"" id=""BannerID"" onchange=""handleBanners(this.value);"">")
              For Each row In rst.Rows
                Send("  <option value=""" & MyCommon.NZ(row.Item("BannerID"), -1) & """>" & MyCommon.NZ(row.Item("Name"), Copient.PhraseLib.Lookup("term.multiple-banners", LanguageID)) & "</option>")
                BannerIDs &= MyCommon.NZ(row.Item("BannerID"), -1) & ","
                BannerEngines &= MyCommon.NZ(row.Item("EngineID"), -1) & ","
              Next
              
              'If default engine id <> CPE, then don't allow multiple banners
              If DefaultEngineID = 2 Then
                Send("  <option value=""0"">[" & Copient.PhraseLib.Lookup("term.multiple-banners", LanguageID) & "]</option>")
                BannerIDs &= "0"
                BannerEngines &= DefaultEngineID.ToString
              End If
              
              Send("</select>")
            Else
              Send(Copient.PhraseLib.Lookup("term.banner", LanguageID) & ": " & MyCommon.SplitNonSpacedString(BannerName, 25))
            End If
            Send("<br /><br class=""half"" />")
          Else
            If (TerminalLockingGroupId = 0) Then
              ' limit to CPE engine only 
              MyCommon.QueryStr = "Select EngineID,DefaultEngine,PhraseID from PromoEngines with (NoLock) where EngineId=2 and Installed=1;"
              rst = MyCommon.LRT_Select
              If (rst.Rows.Count > 0) Then
                Send("<label for=""EngineID"">" & Copient.PhraseLib.Lookup("term.promotionengine", LanguageID) & ":</label><br />")
                Send("<select id=""EngineID"" name=""EngineID"" class=""medium"">")
                For Each row In rst.Rows
                  If MyCommon.NZ(row.Item("EngineID"), 0) = 2 Then
                    Send("  <option selected=""selected"" value=""" & row.Item("EngineID") & """>" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID) & " </option>")
                    EngineType = MyCommon.NZ(row.Item("EngineID"), 0)
                  Else
                    Send("  <option value=""" & row.Item("EngineID") & """>" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID) & " </option>")
                  End If
                Next
                Send("</select><br />")
              End If
            Else
              Send("<input type=""hidden"" id=""EngineID"" name=""EngineID"" value=""" & EngineType & """ />")
              Sendb("                    " & Copient.PhraseLib.Lookup("term.promotionengine", LanguageID) & ": " & EngineName & "<br />")
            End If
          End If
          
        %>
        <br class="half" />
        <%
          MyCommon.QueryStr = "select ActivityDate from ActivityLog with (NoLock) where ActivityTypeID='41' and LinkID='" & TerminalLockingGroupId & "' order by ActivityDate asc;"
          dst = MyCommon.LRT_Select
          SizeOfData = dst.Rows.Count
          If SizeOfData > 0 Then
            Sendb(Copient.PhraseLib.Lookup("term.created", LanguageID) & " ")
            Send(Logix.ToLongDateTimeString(dst.Rows(0).Item("ActivityDate"), MyCommon))
            Send("<br />")
            Sendb(Copient.PhraseLib.Lookup("term.edited", LanguageID) & " ")
            Send(Logix.ToLongDateTimeString(dst.Rows(SizeOfData - 1).Item("ActivityDate"), MyCommon))
          End If
        %>
        <hr class="hidden" />
      </div>
      <a name="h01"></a>
    </div>
    
    <div id="gutter">
    </div>
    
    <div id="column2">
      <div class="box" id="Terminals" <%if(TerminalLockingGroupId = 0)then sendb(" style=""visibility: hidden;""") %>>
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.terminals", LanguageID))%>
          </span>
        </h2>
        <label for="terminals-available">
          <b>
            <% Sendb(Copient.PhraseLib.Lookup("term.available", LanguageID) & ":")%>
          </b>
        </label>
        <br clear="all" />
        <input type="radio" id="functionradio1" name="functionradio1" checked="checked" /><label for="functionradio1"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
        <input type="radio" id="functionradio2" name="functionradio1" /><label for="functionradio2"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
        <input class="longer" onkeydown="handleKeyDown(event, 'terminals-available');" onkeyup="handleKeyUp(9999);" id="functioninput" name="functioninput" maxlength="100" type="text" value="" /><br />
        <br class="half" />
        <select class="longest" multiple="multiple" id="terminals-available" name="terminals-available" size="10">
        </select>
        <br />
        <br />
        <%
          ' First off: queries to get both the selected and excluded terminals
          If (BannersEnabled) Then
            MyCommon.QueryStr = "select TT.TerminalTypeID, TT.Name, isnull(TLG.Name,'" & Copient.PhraseLib.Lookup("term.unassigned", LanguageID) & "') as GroupName from TerminalTypes TT with (NoLock) " & _
                                "left join TerminalLockingGroups TLG on TLG.TerminalLockingGroupID=TT.LockingGroupID and TLG.BannerID=TT.BannerID and TLG.Deleted=0 " & _
                                "where TT.Deleted=0 and TT.AnyTerminal=0 and TT.BannerID=" & BannerID & " and TT.EngineID=" & EngineType & " " & _
                                "and LockingGroupID=" & TerminalLockingGroupId & _
                                "order by GroupName;"
          Else
            MyCommon.QueryStr = "select TT.TerminalTypeID, TT.Name, isnull(TLG.Name,'" & Copient.PhraseLib.Lookup("term.unassigned", LanguageID) & "') as GroupName from TerminalTypes TT with (NoLock) " & _
                                "left join TerminalLockingGroups TLG on TLG.TerminalLockingGroupID=TT.LockingGroupID and TLG.EngineID=TT.EngineID and TLG.Deleted=0 " & _
                                "where TT.Deleted=0 and TT.AnyTerminal=0 and TT.EngineID=" & EngineType & " " & _
                                "and LockingGroupID=" & TerminalLockingGroupId & _
                                "order by GroupName;"
          End If
          rstSelected = MyCommon.LRT_Select
          
          ' SELECTED TERMINALS
          Send("<label for=""terminals-select""><b>" & Copient.PhraseLib.Lookup("term.selected", LanguageID) & ":</b></label>")
          Send("<br />")
          
          ' Buttons
          Sendb("<input type=""submit"" class=""regular select"" id=""terminals-add1"" name=""terminals-add1"" title=""" & Copient.PhraseLib.Lookup("term.select", LanguageID) & """ onclick=""return submitTerminals();"" value=""&#9660; " & Copient.PhraseLib.Lookup("term.select", LanguageID) & """")
          If (Not Logix.UserRoles.EditTerminals) Then
            Sendb(" disabled=""disabled""")
          End If
          Send(" />")
          Sendb("<input type=""submit"" class=""regular deselect"" id=""terminals-rem1"" name=""terminals-rem1"" title=""" & Copient.PhraseLib.Lookup("term.unselect", LanguageID) & """ onclick=""return submitTerminals();"" value=""" & Copient.PhraseLib.Lookup("term.deselect", LanguageID) & " &#9650;""")
          If (Not Logix.UserRoles.EditTerminals) OrElse (rstSelected.Rows.Count = 0) Then
            Sendb(" disabled=""disabled""")
          End If
          Send(" />")
          Send("<br />")
          
          ' List
          Send("<select class=""longest"" multiple=""multiple"" id=""terminals-select"" name=""terminals-select"" size=""" & SelectSize & """>")
          If rstSelected.Rows.Count > 0 Then
            'PrevNameValue = ""
            For Each row In rstSelected.Rows
              'NameValue = MyCommon.NZ(row.Item("GroupName"), Copient.PhraseLib.Lookup("term.unassigned", LanguageID))
              'If (NameValue <> PrevNameValue) Then
              'If (PrevNameValue <> "") Then Send("</optgroup>")
              'Send("   <optgroup label=""" & NameValue & """>  ")
              'End If
              'PrevNameValue = NameValue
              Send("<option value=""" & row.Item("TerminalTypeID") & """>" & row.Item("Name") & "</option>")
            Next
            'Send("</optgroup>")
          End If
          Send("</select>")
          Send("<br />")
          Send("<br class=""half"" />")
          Send("</div>")
        %>
      </div>
    </div>
    <br clear="all" />
    <% DeleteSkip:%>
  </div>
</form>

<script type="text/javascript" language="javascript">
    <% If (BannersEnabled) Then %>
      var bannerIDs = Array(<% Sendb(BannerIDs) %>);
      var bannerEngines = Array(<% Sendb(BannerEngines) %>);
      
      if (document.getElementById("BannerID") != null) {
        handleBanners(document.getElementById("BannerID").value);
      }
    <% End If %>
    handleKeyUp(9999);
</script>

<script type="text/javascript">
    if (window.captureEvents){
        window.captureEvents(Event.CLICK);
        window.onclick=handlePageClick;
    }
    else {
        document.onclick=handlePageClick;
    }

  <% If (infoMessage <> "") Then %>
    // DOM2
    if (typeof window.addEventListener != "undefined")
      window.addEventListener( "load", JumpToTop, false );
    // IE
    else if (typeof window.attachEvent != "undefined") {
      window.attachEvent( "onload", JumpToTop );
    } else {
      if (window.onload != null) {
        var oldOnload = window.onload;
        window.onload = function ( e ) {
        oldOnload(e);
          JumpToTop();
        };
      } else
        window.onload = InitialiseScrollableArea;
    }
    function JumpToTop() {
      try {
        window.location.hash = 'h00';
        document.getElementById('main').scrollTop = 0;
      } catch (err) {
        // ignore
      }
    }
  <% End If %>

</script>

<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (TerminalLockingGroupId > 0 And Logix.UserRoles.AccessNotes) Then
      Send_Notes(37, TerminalLockingGroupId, AdminUserID)
    End If
  End If
done:
Finally
  MyCommon.Close_LogixRT()
End Try
Send_BodyEnd("mainform", "TerminalName")
MyCommon = Nothing
Logix = Nothing
%>
