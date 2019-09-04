<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.CommonInc" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%
  ' *****************************************************************************
  ' * FILENAME: scorecard-edit.aspx 
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
  Dim ScorecardID As Long
  Dim ScorecardTypeID As Integer
  Dim Description As String = ""
  Dim Priority As Integer = 0
  Dim Bold As Boolean = False
  Dim CreatedDate As Date
  Dim LastUpdate As Date
  Dim EngineID As Integer = -1
  Dim EngineName As String = ""
  Dim DefaultForEngine As Boolean = False
  Dim Defaultable As Boolean = False
  Dim PrintTotalLine As Boolean = False
  Dim TotalLinePosition As Integer = 0
  Dim DefaultTotalLinePosition As Integer = 0
  Dim Deleted As Boolean = False
  Dim LoneScorecard As Boolean = False
  
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim rst As DataTable
  Dim rst2 As DataTable
  Dim rstsc As DataTable
  Dim row As DataRow
  Dim bSave As Boolean
  Dim bDelete As Boolean
  Dim bCreate As Boolean
  Dim ShowActionButton As Boolean = False
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False

  'Scorecard Preference if RLM Integration and ScoreCard Preference set
   Dim PreferenceName As String = MyCommon.NZ(MyCommon.Fetch_CPE_SystemOption(188), "0")
   Dim bScorecardPref As Boolean = MyCommon.NZ(MyCommon.Fetch_CPE_SystemOption(163), 0) = 1 AndAlso PreferenceName <> "0"
   Dim PreferenceListItemName As String
   Dim ItemID As Integer 
   Dim PrefName As String
   Dim PrintZeroBalance As Boolean = False
   
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "scorecard-edit.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_PrefManRT()  
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  Try
    ' fill in if it was a get method
    If Request.RequestType = "GET" Then
      ScorecardID = Request.QueryString("ScorecardID")
      ScorecardTypeID = Request.QueryString("ScorecardTypeID")
      Description = Left(Request.QueryString("Description"), 28)
      If Request.QueryString("Priority") <> "" Then
        Priority = MyCommon.Extract_Val(Request.QueryString("Priority"))
      End If
      If MyCommon.Extract_Val(Request.QueryString("Bold")) = 1 Then
        Bold = True
      End If
      EngineID = Request.QueryString("EngineID")
      If (ScorecardTypeID = 2) OrElse (ScorecardTypeID = 4) Then
        EngineID = 2
      End If
      If MyCommon.Extract_Val(Request.QueryString("DefaultForEngine")) = 1 Then
        DefaultForEngine = True
      End If
      If MyCommon.Extract_Val(Request.QueryString("PrintTotalLine")) = 1 Then
        PrintTotalLine = True
      End If
      TotalLinePosition = MyCommon.Extract_Val(Request.QueryString("TotalLinePosition"))
      PreferenceListItemName = Request.QueryString("PreferenceListItemName")
      If MyCommon.Extract_Val(Request.QueryString("PrintZeroBalance")) = 1 Then
        PrintZeroBalance = True
      End If  	  
      If Request.QueryString("save") = "" AndAlso Request.QueryString("savePressed") = "" Then
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
      ScorecardID = Request.Form("ScorecardID")
      ScorecardTypeID = Request.Form("ScorecardTypeID")
      If ScorecardID = 0 Then
        ScorecardID = MyCommon.Extract_Val(Request.QueryString("ScorecardID"))
      End If
      Description = Left(Request.Form("Description"), 28)
      If Request.Form("Priority") <> "" Then
        Priority = MyCommon.Extract_Val(Request.Form("Priority"))
      End If
      If Request.Form("Bold") = 1 Then
        Bold = True
      End If
      EngineID = Request.Form("EngineID")
      If (ScorecardTypeID = 2) OrElse (ScorecardTypeID = 4) Then
        EngineID = 2
      End If
      If Request.Form("DefaultForEngine") = 1 Then
        DefaultForEngine = True
      End If
      If Request.Form("PrintTotalLine") = 1 Then
        PrintTotalLine = True
      End If
      TotalLinePosition = MyCommon.Extract_Val(Request.Form("TotalLinePosition"))
      PreferenceListItemName = Request.Form("PreferenceListItemName")
      If Request.Form("PrintZeroBalance") = 1 Then
        PrintZeroBalance = True
      End If	  
      If Request.Form("save") = "" AndAlso Request.QueryString("savePressed") = "" Then
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
    
    
    Send_HeadBegin("term.scorecard", , ScorecardID)
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

  function toggleTotalLinePosition() {
    var PrintTotalLineElem = document.getElementById("PrintTotalLine");
    var TotalLinePositionRow = document.getElementById("TotalLinePositionRow");

    if ((PrintTotalLineElem != null) && (TotalLinePositionRow != null)) {
      if (PrintTotalLineElem.checked == true) {
        TotalLinePositionRow.style.display = "";
      } else {
        TotalLinePositionRow.style.display = "none";
      }
    }
  }

  function handleKeyDown(e) {
    var keycode;
    var submitThing;
    
    if (window.event) keycode = window.event.keyCode;
    else if (e) keycode = e.which;
    else return true;
    
    if (keycode == 13) {
      submitThing = document.getElementById("savePressed");
      submitThing.value = "save";
    <% If Request.Browser.Browser <> "IE" Then %>
      document.mainform.submit();
    <% End If %>
      return false;
    } else {
      return true;
    }
    return true;
  }

   // This javascript function will set engine id in hidden field when user changes engine using select drop down.
  function OnEngineIDChange (selectedEngine)
  {
      var EngineIdHidden = document.getElementById("EngineID");
      EngineIdHidden.value = selectedEngine.value; 
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
    Response.Redirect("scorecard-edit.aspx?ScorecardTypeID=" & Request.QueryString("ScorecardTypeID"))
  End If
  
  If bSave Then
    ' Before anything else, check to see if there are any other scorecards already existing for the selected engine.  If the one
    ' we're saving IS the only one for its engine, we'll automatically set the "Default for Engine" bit as a convenience.
    MyCommon.QueryStr = "select ScorecardID, EngineID, DefaultForEngine from Scorecards with (NoLock) " & _
                        "where EngineID=" & EngineID & " and Deleted=0 and ScorecardTypeID=" & ScorecardTypeID & ";"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count = 0) Then
      LoneScorecard = True
    End If
    If (ScorecardID = 0) Then
      MyCommon.QueryStr = "dbo.pt_Scorecard_Insert"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@ScorecardTypeID", SqlDbType.Int).Value = ScorecardTypeID
      MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 100).Value = Description.Trim()
      MyCommon.LRTsp.Parameters.Add("@Priority", SqlDbType.Int).Value = 0 ' ...Hard-coding this to '0'; change it back to 'Priority' if you want to allow users to control priority
      MyCommon.LRTsp.Parameters.Add("@Bold", SqlDbType.Bit).Value = Bold
      MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = EngineID
      If LoneScorecard Then
        MyCommon.LRTsp.Parameters.Add("@DefaultForEngine", SqlDbType.Bit).Value = 1
      Else
        MyCommon.LRTsp.Parameters.Add("@DefaultForEngine", SqlDbType.Bit).Value = DefaultForEngine
      End If
      MyCommon.LRTsp.Parameters.Add("@PrintTotalLine", SqlDbType.Bit).Value = PrintTotalLine
      MyCommon.LRTsp.Parameters.Add("@TotalLinePosition", SqlDbType.Bit).Value = TotalLinePosition
      MyCommon.LRTsp.Parameters.Add("@PreferenceListItemName", SqlDbType.NVarChar, 50).Value = PreferenceListItemName
      MyCommon.LRTsp.Parameters.Add("@PrintZeroBalance", SqlDbType.Bit).Value = PrintZeroBalance	  
      MyCommon.LRTsp.Parameters.Add("@ScorecardID", SqlDbType.Int).Direction = ParameterDirection.Output
      If (Description.Trim() = "") Then
        infoMessage = Copient.PhraseLib.Lookup("scorecard-edit.noname", LanguageID)
      Else
        MyCommon.QueryStr = "SELECT ScorecardID FROM Scorecards with (NoLock) WHERE Description='" & MyCommon.Parse_Quotes(Description.Trim()) & "' AND Deleted=0 AND EngineID= '" & EngineID &"';"
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
          infoMessage = Copient.PhraseLib.Lookup("scorecard-edit.nameused", LanguageID)
        Else
          MyCommon.LRTsp.ExecuteNonQuery()
          ScorecardID = MyCommon.LRTsp.Parameters("@ScorecardID").Value
          MyCommon.Activity_Log(40, ScorecardID, AdminUserID, Copient.PhraseLib.Lookup("history.scorecard-create", LanguageID))
          Response.Status = "301 Moved Permanently"
          Response.AddHeader("Location", "scorecard-edit.aspx?ScorecardID=" & ScorecardID & "&ScorecardTypeID" & ScorecardTypeID)
        End If
      End If
      MyCommon.Close_LRTsp()
    Else
      MyCommon.QueryStr = "dbo.pt_Scorecard_Update"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@ScorecardTypeID", SqlDbType.Int).Value = ScorecardTypeID
      MyCommon.LRTsp.Parameters.Add("@ScorecardID", SqlDbType.Int).Value = ScorecardID
      MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 1000).Value = Description.Trim()
      MyCommon.LRTsp.Parameters.Add("@Priority", SqlDbType.Int).Value = 0 ' ...Hard-coding this to '0'; change it back to 'Priority' if you want to allow users to control priority
      MyCommon.LRTsp.Parameters.Add("@Bold", SqlDbType.Bit).Value = Bold
      MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = EngineID
      If LoneScorecard Then
        MyCommon.LRTsp.Parameters.Add("@DefaultForEngine", SqlDbType.Bit).Value = 1
      Else
        MyCommon.LRTsp.Parameters.Add("@DefaultForEngine", SqlDbType.Bit).Value = DefaultForEngine
      End If
      MyCommon.LRTsp.Parameters.Add("@PrintTotalLine", SqlDbType.Bit).Value = PrintTotalLine
      MyCommon.LRTsp.Parameters.Add("@TotalLinePosition", SqlDbType.Bit).Value = TotalLinePosition
      MyCommon.LRTsp.Parameters.Add("@PreferenceListItemName", SqlDbType.NVarChar, 50).Value = PreferenceListItemName
      MyCommon.LRTsp.Parameters.Add("@PrintZeroBalance", SqlDbType.Bit).Value = PrintZeroBalance  
      If (Description.Trim() = "") Then
        infoMessage = Copient.PhraseLib.Lookup("scorecard-edit.noname", LanguageID)
      Else
        MyCommon.QueryStr = "select ScorecardID from Scorecards with (NoLock) " & _
                            "where Description='" & MyCommon.Parse_Quotes(Description.Trim()) & "' and Deleted=0 and ScorecardID<>" & ScorecardID & " AND EngineID= '" & EngineID & "';"
        rst2 = MyCommon.LRT_Select
        If (rst2.Rows.Count > 0) Then
          infoMessage = Copient.PhraseLib.Lookup("scorecard-edit.nameused", LanguageID)
        Else
          MyCommon.LRTsp.ExecuteNonQuery()
          MyCommon.Activity_Log(40, ScorecardID, AdminUserID, Copient.PhraseLib.Lookup("history.scorecard-edit", LanguageID))
        End If
      End If
      MyCommon.Close_LRTsp()
    End If
  ElseIf bDelete Then
   'AMS-4916 Validate scorecard before deletion, also refactored above code in one stored procedure for a single call
    MyCommon.QueryStr = "pa_IsScoreCardInUse"
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add("@ScorecardId", SqlDbType.Int).Value = Request.QueryString("ScorecardID")
    rst = MyCommon.LRTsp_select()
    MyCommon.Close_LRTsp()
    If rst.Rows.Count > 0 Then
      infoMessage = Copient.PhraseLib.Lookup("scorecard.inuse", LanguageID)
    Else
      MyCommon.QueryStr = "dbo.pt_Scorecard_Delete"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@ScorecardID", SqlDbType.Int).Value = ScorecardID
      MyCommon.LRTsp.ExecuteNonQuery()
      MyCommon.Close_LRTsp()
      MyCommon.Activity_Log(40, ScorecardID, AdminUserID, Copient.PhraseLib.Lookup("history.scorecard-delete", LanguageID))
      Response.Status = "301 Moved Permanently"
      Response.AddHeader("Location", "scorecard-list.aspx?ScorecardTypeID=" & ScorecardTypeID)
      ScorecardID = 0
      Description = ""
    End If
  End If
  
  If Not bCreate Then
    ' no one clicked anything
    MyCommon.QueryStr = "select ScorecardID, ScorecardTypeID, Description, Priority, Bold, CreatedDate, LastUpdate, " & _
                        "EngineID, DefaultForEngine, PrintTotalLine, TotalLinePosition, Deleted , PreferenceListItemName, PrintZeroBalance " & _
                        "from Scorecards with (NoLock) where ScorecardID=" & ScorecardID & ";"
    rst = MyCommon.LRT_Select()
    If (rst.Rows.Count > 0) Then
      For Each row In rst.Rows
        ScorecardTypeID = MyCommon.NZ(row.Item("ScorecardTypeID"), 1)
        Description = MyCommon.NZ(row.Item("Description"), "")
        Priority = MyCommon.NZ(row.Item("Priority"), "")
        Bold = MyCommon.NZ(row.Item("Bold"), False)
        EngineID = MyCommon.NZ(row.Item("EngineID"), -1)
        DefaultForEngine = MyCommon.NZ(row.Item("DefaultForEngine"), False)
        PrintTotalLine = MyCommon.NZ(row.Item("PrintTotalLine"), False)
        TotalLinePosition = MyCommon.NZ(row.Item("TotalLinePosition"), 0)
        CreatedDate = MyCommon.NZ(row.Item("CreatedDate"), "1/1/2000")
        LastUpdate = MyCommon.NZ(row.Item("LastUpdate"), "1/1/2000")
        PreferenceListItemName = MyCommon.NZ(row.Item("PreferenceListItemName"), "")
        PrintZeroBalance = MyCommon.NZ(row.Item("PrintZeroBalance"), 0)		
        If MyCommon.NZ(row.Item("Deleted"), False) Then
          Deleted = True
          infoMessage = Copient.PhraseLib.Lookup("error.deleted", LanguageID)
        End If
      Next
    ElseIf (ScorecardID > 0) Then
      Send("")
      Send("<div id=""intro"">")
      Send("  <h1 id=""title"">" & Copient.PhraseLib.Lookup("term.scorecard", LanguageID) & " #" & ScorecardID & "</h1>")
      Send("</div>")
      Send("<div id=""main"">")
      Send("  <div id=""infobar"" class=""red-background"">")
      Send("    " & Copient.PhraseLib.Lookup("error.deleted", LanguageID))
      Send("  </div>")
      Send("</div>")
      GoTo done
    End If
    MyCommon.QueryStr = "select DefaultForEngine from Scorecards with (NoLock) " & _
                        "where EngineID=" & EngineID & " and ScorecardTypeID=" & ScorecardTypeID & " and DefaultForEngine=1 and Deleted=0;"
    rst2 = MyCommon.LRT_Select
    If rst2.Rows.Count > 0 Then
      Defaultable = False
    Else
      Defaultable = True
    End If
    DefaultTotalLinePosition = MyCommon.Fetch_SystemOption(101)
  End If
%>
<form action="#" id="mainform" name="mainform">
  <input type="hidden" id="ScorecardID" name="ScorecardID" value="<% Sendb(ScorecardID) %>" />
  <input type="hidden" id="ScorecardTypeID" name="ScorecardTypeID" value="<% Sendb(ScorecardTypeID) %>" />
  <input type="hidden" name="savePressed" id="savePressed" value="" />
  <div id="intro">
    <%
      Sendb("<h1 id=""title"">")
      If ScorecardID = 0 Then
        Sendb(Copient.PhraseLib.Lookup("term.new", LanguageID) & " ")
        If ScorecardTypeID = 1 Then
          Sendb(StrConv(Copient.PhraseLib.Lookup("term.points", LanguageID), VbStrConv.Lowercase) & " ")
        ElseIf ScorecardTypeID = 2 Then
          Sendb(StrConv(Copient.PhraseLib.Lookup("term.storedvalue", LanguageID), VbStrConv.Lowercase) & " ")
        ElseIf ScorecardTypeID = 3 Then
          Sendb(StrConv(Copient.PhraseLib.Lookup("term.discount", LanguageID), VbStrConv.Lowercase) & " ")
        ElseIf ScorecardTypeID = 4 Then
          Sendb(StrConv(Copient.PhraseLib.Lookup("term.limits", LanguageID), VbStrConv.Lowercase) & " ")
        End If
        Sendb(StrConv(Copient.PhraseLib.Lookup("term.scorecard", LanguageID), VbStrConv.Lowercase))
        Description = ""
      Else
        If ScorecardTypeID = 1 Then
          Sendb(Copient.PhraseLib.Lookup("term.points", LanguageID) & " ")
        ElseIf ScorecardTypeID = 2 Then
          Sendb(Copient.PhraseLib.Lookup("term.storedvalue", LanguageID) & " ")
        ElseIf ScorecardTypeID = 3 Then
          Sendb(Copient.PhraseLib.Lookup("term.discount", LanguageID) & " ")
        ElseIf ScorecardTypeID = 4 Then
          Sendb(Copient.PhraseLib.Lookup("term.limits", LanguageID) & " ")
        End If
        Sendb(StrConv(Copient.PhraseLib.Lookup("term.scorecard", LanguageID), VbStrConv.Lowercase))
        Sendb(" #" & ScorecardID)
        MyCommon.QueryStr = "select Description from Scorecards with (NoLock) where ScorecardID=" & ScorecardID & ";"
        rst2 = MyCommon.LRT_Select
        If (rst2.Rows.Count > 0) Then
          Sendb(": " & MyCommon.TruncateString(MyCommon.NZ(rst2.Rows(0).Item("Description"), ""), 40))
        End If
      End If
      Send("</h1>")
    %>
    <div id="controls">
      <%
        If Not Deleted Then
          If (ScorecardID = 0) Then
            If (Logix.UserRoles.EditScorecard) Then
              Send_Save()
            End If
          Else
            ShowActionButton = (Logix.UserRoles.EditScorecard)
            If (ShowActionButton) Then
              Send("<input type=""button"" class=""regular"" id=""actions"" name=""actions"" value=""" & Copient.PhraseLib.Lookup("term.actions", LanguageID) & " ▼"" onclick=""toggleDropdown();"" />")
              Send("<div class=""actionsmenu"" id=""actionsmenu"">")
              If (Logix.UserRoles.EditScorecard) Then
                Send_Save()
              End If
              If (Logix.UserRoles.EditScorecard) Then
                Send_Delete()
              End If
              If (Logix.UserRoles.EditScorecard) Then
                Send_New()
              End If
              Send("</div>")
            End If
            If MyCommon.Fetch_SystemOption(75) Then
              If (Logix.UserRoles.AccessNotes) Then
                Send_NotesButton(36, ScorecardID, AdminUserID)
              End If
            End If
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
      If Deleted Then
        GoTo DeleteSkip
      End If
    %>
    <div id="column1">
      <div class="box" id="identification">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.identification", LanguageID))%>
          </span>
        </h2>
        <table>
          <tr>
            <td>
              <label for="description"><% Sendb(Copient.PhraseLib.Lookup("term.description", LanguageID))%>:</label>
            </td>
            <td>
              <input type="text" class="mediumlong" id="description" name="description" maxlength="28" onkeydown="handleKeyDown(event);" value="<% Sendb(Description.Replace("""", "&quot;")) %>" />
            </td>
          </tr>
          <%
            Send("<tr>")
            If ScorecardID = 0 AndAlso ScorecardTypeID = 1 Then
              MyCommon.QueryStr = "select PE.EngineID, PE.Description, PE.PhraseID from PromoEngines as PE with (NoLock) " & _
                                  "where PE.Installed=1 and PE.EngineID in (2,6);"
              rst2 = MyCommon.LRT_Select
              Send("  <td>")
              Send("    <label for=""EngineID"">" & Copient.PhraseLib.Lookup("term.engine", LanguageID) & ":</label>")
              Send("  </td>")
              Send("  <td>")
              Send("    <select id=""EngineID"" name=""EngineID"">")
              If rst2.Rows.Count > 0 Then
                For Each row In rst2.Rows
                  Send("      <option value=""" & MyCommon.NZ(row.Item("EngineID"), 0) & """" & IIf(MyCommon.NZ(row.Item("EngineID"), -1) = EngineID, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "</option>")
                Next
              End If
              Send("    </select>")
              Send("  </td>")
            Else
              Send("  <td>")
              Send("    " & Copient.PhraseLib.Lookup("term.engine", LanguageID) & ":")
              Send("  </td>")
              Send("  <td>")
                  If (ScorecardID = 0) Then
                      MyCommon.QueryStr = "select PE.EngineID, PE.Description, PE.DefaultEngine, PE.PhraseID from PromoEngines as PE with (NoLock) " & _
                                "where PE.Installed=1 and PE.EngineID in (2,9);"
                      rst2 = MyCommon.LRT_Select
                      If (Not rst2 Is Nothing) Then
                          If (rst2.Rows.Count > 1) Then
                              Send("    <select id=""EngineID"" name=""EngineID"" onchange=""OnEngineIDChange(this)"">")
                              If rst2.Rows.Count > 0 Then
                                  For Each row In rst2.Rows
                                      Send("      <option value=""" & MyCommon.NZ(row.Item("EngineID"), 0) & """" & IIf(MyCommon.NZ(row.Item("DefaultEngine"), -1) = 1, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "</option>")
                                  Next
                              End If
                              Send("    </select>")
                              Send("    <input type=""hidden"" id=""EngineID"" name=""EngineID"" />")
                          Else
                              Send("    " & Copient.PhraseLib.Lookup(MyCommon.NZ(rst2.Rows(0).Item("PhraseID"), 0), LanguageID))
                              Send("    <input type=""hidden"" id=""EngineID"" name=""EngineID"" value=""" & MyCommon.NZ(rst2.Rows(0).Item("EngineID"), 0) & """ />")
                          End If
                      End If
                  Else
                      MyCommon.QueryStr = "select SC.ScorecardID,sc.EngineID, PE.PhraseID from Scorecards SC " & _
                                          "inner join PromoEngines PE on PE.EngineID = SC.EngineID " & _
                                          "where SC.Deleted=0 and SC.ScorecardTypeID=" & ScorecardTypeID & " and SC.ScorecardID=" & ScorecardID & ";"
                      rstsc = MyCommon.LRT_Select
                      If (Not rstsc Is Nothing) Then
                          If (rstsc.Rows.Count <> 0) Then
                              Send("    " & Copient.PhraseLib.Lookup(MyCommon.NZ(rstsc.Rows(0).Item("PhraseID"), 0), LanguageID))
                              Send("    <input type=""hidden"" id=""EngineID"" name=""EngineID"" value=""" & MyCommon.NZ(rstsc.Rows(0).Item("EngineID"), 0) & """ />")
                          End If
                      End If
                  End If
                      Send("  </td>")
                  End If
            Send("</tr>")
            
            If ScorecardID <> 0 Then
              Send("<tr>")
              Send("  <td>")
              Send("    <label for=""DefaultForEngine"">" & Copient.PhraseLib.Lookup("term.DefaultForEngine", LanguageID) & ":</label>")
              Send("  </td>")
              Send("  <td>")
              Send("    <input type=""checkbox"" id=""DefaultForEngine"" name=""DefaultForEngine"" value=""1""" & IIf(DefaultForEngine, " checked=""checked""", "") & IIf(Defaultable = False And DefaultForEngine = False, " disabled=""disabled""", "") & " />")
              Send("  </td>")
              Send("</tr>")
            End If
            
            Send("<tr" & IIf(ScorecardTypeID = 2 Or ScorecardTypeID = 4, " style=""display:none;""", "") & ">")
            Send("  <td>")
            Send("    <label for=""PrintTotalLine"">" & Copient.PhraseLib.Lookup("scorecard.printtotalline", LanguageID) & ":</label>")
            Send("  </td>")
            Send("  <td>")
            If ScorecardTypeID = 2 Then
              'Stored value
              Send("    <input type=""hidden"" id=""PrintTotalLine"" name=""PrintTotalLine"" value=""0"" />")
            Else
              'Points and discounts
              Send("    <input type=""checkbox"" id=""PrintTotalLine"" name=""PrintTotalLine"" value=""1""" & IIf(PrintTotalLine, " checked=""checked""", "") & " onclick=""javascript:toggleTotalLinePosition();"" />")
            End If
            Send("  </td>")
            Send("</tr>")
            
            Send("<tr id=""TotalLinePositionRow""" & IIf(PrintTotalLine, "", " style=""display:none;""") & ">")
            Send("  <td>")
            Send("    " & Copient.PhraseLib.Lookup("scorecard.totallineposition", LanguageID) & ":")
            Send("  </td>")
            Send("  <td>")
            Sendb("    <input type=""radio"" id=""TotalLinePosition0"" name=""TotalLinePosition"" value=""0""")
            If (ScorecardID > 0 AndAlso TotalLinePosition = 0) OrElse (ScorecardID > 0 AndAlso PrintTotalLine = False AndAlso DefaultTotalLinePosition = 0) OrElse (ScorecardID = 0 AndAlso DefaultTotalLinePosition = 0) Then
              Sendb(" checked=""checked""")
            End If
            Send(" /><label for=""TotalLinePosition0"">" & Copient.PhraseLib.Lookup("term.bottom", LanguageID) & "</label>&nbsp;")
            Sendb("    <input type=""radio"" id=""TotalLinePosition1"" name=""TotalLinePosition"" value=""1""")
            If (ScorecardID > 0 AndAlso TotalLinePosition = 1) OrElse (ScorecardID > 0 AndAlso PrintTotalLine = False AndAlso DefaultTotalLinePosition = 1) OrElse (ScorecardID = 0 AndAlso DefaultTotalLinePosition = 1) Then
              Sendb(" checked=""checked""")
            End If
            Sendb(" /><label for=""TotalLinePosition1"">" & Copient.PhraseLib.Lookup("term.top", LanguageID) & "</label>")
            Send("  </td>")
            Send("</tr>")
          %>
          <tr>
            <td>
              <label for="bold"><% Sendb(Copient.PhraseLib.Lookup("term.bold", LanguageID)) %>:</label>
            </td>
            <td>
              <input type="checkbox" id="bold" name="bold" value="1"<% Sendb(IIf(Bold, " checked=""checked""", "")) %>/>
            </td>
          </tr>
          <tr>
          <% Send("<tr" & IIf(bScorecardPref AndAlso (ScorecardTypeID = 1 Or ScorecardTypeID = 2), "", " style=""display:none;""") & ">") %>
            <td>
              <label for="PrintZeroBalance"><% Sendb(Copient.PhraseLib.Lookup("term.printzerobalance", LanguageID)) %>:</label>
            </td>
            <td>
              <input type="checkbox" id="PrintZeroBalance" name="PrintZeroBalance" value="1"<% Sendb(IIf(PrintZeroBalance, " checked=""checked""", "")) %>/>
            </td>
          </tr>
            <% If MyCommon.IsIntegrationInstalled(Integrations.PREFERENCE_MANAGER) Then%>
           <tr>
          <% Send("<tr" & IIf(bScorecardPref AndAlso (ScorecardTypeID = 1 Or ScorecardTypeID = 2), "", " style=""display:none;""") & ">") %>
            <td>
              <% Send("<label for=""PreferenceListItemName"">" & PreferenceName & ": </label>") %>
            </td>
            <td>
             <% 
              Send("    <select class=""medium"" id=""PreferenceListItemName"" name=""PreferenceListItemName"" >")
               MyCommon.QueryStr = "select PLI.ItemID as ItemID, PLI.Name as Name from PreferenceListItems as PLI with (NoLock)  Full Outer Join Preferences as P with (NoLock) on P.PreferenceID=PLI.PreferenceID where P.Name = '" & PreferenceName & "' ;"
               rst = MyCommon.PMRT_Select
               Send("      <option value=""0"">" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "</option>")         
               If rst.Rows.Count > 0 Then
                 For Each row In rst.Rows
                   ItemID = MyCommon.NZ(row.Item("ItemID"), 0) 
                   PrefName = MyCommon.NZ(row.Item("Name"), "")
                   If(PrefName = PreferenceListItemName) Then
                     Send("      <option value=""" & PrefName & """ selected=""selected"" >" & PrefName & "</option>") 
                   Else
                     Send("      <option value=""" & PrefName & """ >" & PrefName & "</option>") 
                   End If
                 Next
               End If
               Send("    </select>") 
              %>
            </td>
          </tr>		
             <%  End If %>
          <!-- Commenting out priority per Mark's suggestion.  The order of multiple scorecards is determined when the footer printed message is assembled.
          <tr>
            <td>
              <label for="priority"><% Sendb(Copient.PhraseLib.Lookup("term.priority", LanguageID)) %>:</label>
            </td>
            <td>
              <input type="text" class="shortest" id="priority" name="priority" maxlength="2" value="<% Sendb(Priority) %>" /> (0 = lowest, 99=highest)
            </td>
          </tr>
          -->
        </table>
        <%
          If ScorecardID <> 0 Then
            Send("<br class=""half"" />")
            Send(Copient.PhraseLib.Lookup("term.created", LanguageID) & " " & Logix.ToLongDateTimeString(CreatedDate, MyCommon))
            Send("<br />")
            Send(Copient.PhraseLib.Lookup("term.edited", LanguageID) & " " & Logix.ToLongDateTimeString(LastUpdate, MyCommon))
          End If
        %>
        <hr class="hidden" />
      </div>
    </div>
    
    <div id="gutter">
    </div>
    
    <div id="column2">
      <div class="box" id="offers"<%if(ScorecardID = 0)then sendb(" style=""visibility: hidden;""") %>>
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.associatedoffers", LanguageID))%>
          </span>
        </h2>
        <div class="boxscroll">
          <%
            If (Request.QueryString("ScorecardID") <> "") And (ScorecardTypeID >= 1 And ScorecardTypeID <= 4) Then
              If ScorecardTypeID = 1 Then 'points
                MyCommon.QueryStr = "select distinct CDP.ScorecardID, D.RewardOptionID, RO.IncentiveID as OfferID, I.IncentiveName as Name, I.EligibilityEndDate as ProdEndDate,buy.ExternalBuyerId as BuyerID " & _
                                    "from CPE_DeliverablePoints as CDP with (NoLock) " & _
                                    "inner join CPE_Deliverables as D with (NoLock) on D.OutputID=CDP.PKID " & _
                                    "inner join CPE_RewardOptions as RO with (NoLock) on RO.RewardOptionID=D.RewardOptionID " & _
                                    "inner join CPE_Incentives as I with (NoLock) on I.IncentiveID=RO.IncentiveID " & _
                                     "left outer join Buyers as buy with (nolock) on buy.BuyerId= I.BuyerId " & _
                                    "where CDP.ScorecardID = " & ScorecardID & " and D.DeliverableTypeID = 8 and CDP.Deleted = 0 and D.Deleted = 0 and RO.Deleted = 0 and I.Deleted = 0 " & _
                                    "order by IncentiveName;"
              ElseIf ScorecardTypeID = 2 Then 'stored value
                MyCommon.QueryStr = "select distinct CDSV.ScorecardID, D.RewardOptionID, RO.IncentiveID as OfferID, I.IncentiveName as Name, I.EligibilityEndDate as ProdEndDate,buy.ExternalBuyerId as BuyerID " & _
                                    "from CPE_DeliverableStoredValue as CDSV with (NoLock) " & _
                                    "inner join CPE_Deliverables as D with (NoLock) on D.OutputID=CDSV.PKID " & _
                                    "inner join CPE_RewardOptions as RO with (NoLock) on RO.RewardOptionID=D.RewardOptionID " & _
                                    "inner join CPE_Incentives as I with (NoLock) on I.IncentiveID=RO.IncentiveID " & _
                                     "left outer join Buyers as buy with (nolock) on buy.BuyerId= I.BuyerId " & _
                                    "where CDSV.ScorecardID = " & ScorecardID & " and D.DeliverableTypeID = 11 and CDSV.Deleted = 0 and D.Deleted = 0 and RO.Deleted = 0 and I.Deleted = 0 " & _
                                    "order by IncentiveName;"
              ElseIf ScorecardTypeID = 3 Then 'discounts
                MyCommon.QueryStr = "select distinct CD.ScorecardID, D.RewardOptionID, RO.IncentiveID as OfferID, I.IncentiveName as Name, I.EligibilityEndDate as ProdEndDate,buy.ExternalBuyerId as BuyerID " & _
                                    "from CPE_Discounts as CD with (NoLock) " & _
                                    "inner join CPE_Deliverables as D with (NoLock) on D.OutputID=CD.DiscountID " & _
                                    "inner join CPE_RewardOptions as RO with (NoLock) on RO.RewardOptionID=D.RewardOptionID " & _
                                    "inner join CPE_Incentives as I with (NoLock) on I.IncentiveID=RO.IncentiveID " & _
                                     "left outer join Buyers as buy with (nolock) on buy.BuyerId= I.BuyerId " & _
                                    "where CD.ScorecardID=" & ScorecardID & " and D.DeliverableTypeID = 2 and CD.Deleted=0 and D.Deleted=0 and RO.Deleted=0 And I.Deleted=0 " & _
                                    "order by IncentiveName;"
              ElseIf ScorecardTypeID = 4 Then 'limits
                MyCommon.QueryStr = "select I.IncentiveID as OfferID, I.IncentiveName as Name, I.EligibilityEndDate as ProdEndDate,buy.ExternalBuyerId as BuyerID " & _
                                    "from CPE_Incentives as I with (NoLock) " & _
                                     "left outer join Buyers as buy with (nolock) on buy.BuyerId= I.BuyerId " & _
                                    "where I.ScorecardID=" & ScorecardID & " and I.Deleted=0 " & _
                                    "order by IncentiveName;"
              End If
                  Dim assocName As String=""
              rst = MyCommon.LRT_Select
              If rst.Rows.Count > 0 Then
                For Each row In rst.Rows
                    If (MyCommon.Fetch_UE_SystemOption(168) = "1" AndAlso MyCommon.NZ(row.Item("BuyerID"), "") <> "") Then
                    assocName = "Buyer " + row.Item("BuyerID").ToString() + " - " + MyCommon.NZ(row.Item("Name"), "").ToString()
                    Else
                    assocName = MyCommon.NZ(row.Item("Name"),Copient.PhraseLib.Lookup("term.unknown", LanguageID))
                    End If
                  If (Logix.IsAccessibleOffer(AdminUserID, row.Item("OfferID"))) Then
                    Sendb(" <a href=""offer-redirect.aspx?OfferID=" & row.Item("OfferID") & """>" & assocName & "</a>")
                  Else
                    Sendb(assocName)
                  End If
                  Sendb(" (" & Logix.GetOfferStatus(row.Item("OfferID"), LanguageID) & ")")
                  Send("<br />")
                Next
              Else
                Send(Copient.PhraseLib.Lookup("term.none", LanguageID) & "<br />")
              End If
            Else
              Send(Copient.PhraseLib.Lookup("term.none", LanguageID) & "<br />")
            End If
          %>
        </div>
        <hr class="hidden" />
      </div>
    </div>
    <br clear="all" />
    <% DeleteSkip:%>
  </div>
</form>

<script type="text/javascript">
  if (window.captureEvents){
    window.captureEvents(Event.CLICK);
    window.onclick=handlePageClick;
  }
  else {
    document.onclick=handlePageClick;
  }
</script>
<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (ScorecardID > 0 And Logix.UserRoles.AccessNotes) Then
      Send_Notes(36, ScorecardID, AdminUserID)
    End If
  End If
done:
Finally
  MyCommon.Close_LogixRT()
  MyCommon.Close_PrefManRT()  
End Try
Send_BodyEnd("mainform", "description")
MyCommon = Nothing
Logix = Nothing
%>
