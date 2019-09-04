<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: agent-list.aspx 
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
  Dim row As System.Data.DataRow
  Dim idSearchText As String
  Dim PageNum As Integer = 0
  Dim MorePages As Boolean
  Dim linesPerPage As Integer = 100
  Dim sizeOfData As Integer
  Dim i As Integer = 0
  Dim SortText As String = "AppName"
  Dim SortDirection As String = "DESC"
  Dim Offline As Boolean = False
  Dim Watchdog As System.Data.DataTable
  Dim shaded As Boolean = True
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "agent-list.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  Send_HeadBegin("term.agents")
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
  Send_Subtabs(Logix, 8, 1)
  
  If (Logix.UserRoles.AccessSystemHealth = False) Then
    Send_Denied(1, "perm.admin-health")
    GoTo done
  End If
  
  PageNum = Request.QueryString("pagenum")
  If PageNum < 0 Then PageNum = 0
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
      SortDirection = "ASC"
    End If
  Else
    SortDirection = Request.QueryString("SortDirection")
  End If
  
  If (Request.QueryString("action") <> "" And Request.QueryString("appid") <> "") Then
    If (Request.QueryString("action") = "mute") Then
      MyCommon.QueryStr = "update LastSync with (RowLock) Set SendAlert = 0 where AppID = " & Request.QueryString("appid")
      MyCommon.LRT_Execute()
    ElseIf (Request.QueryString("action") = "unmute") Then
      MyCommon.QueryStr = "update LastSync with (RowLock) Set SendAlert = 1 where AppID = " & Request.QueryString("appid")
      MyCommon.LRT_Execute()
    End If
  End If

    If (Server.HtmlEncode(Request.QueryString("searchterms")) <> "") Then
        idSearchText = MyCommon.Parse_Quotes(Server.HtmlEncode(Request.QueryString("searchterms")))
    If (idSearchText.IndexOf("_") > -1) Then idSearchText = idSearchText.Replace("_", "[_]")
        MyCommon.QueryStr = "select Distinct LS.AppID, LS.PhraseID , LS.AppName, LastStartTime, LastEndTime, " & _
                        "LastTouchTime, LastAppErrorTime, RunFreq, IsNull(SendAlert, 0) SendAlert " & _
                        "from LastSync as LS with (NoLock) " & _
                        "left Join LastSyncEngines as LSE with (NoLock) on LSE.AppID=LS.AppID " & _
                        "left join PromoEngines as PE with (NoLock) on PE.EngineID=LSE.EngineID " & _
                        "where PE.Installed=1 or LSE.EngineID=-1 " & _
                        "and AppName like N'%" & idSearchText & "%' order by " & SortText & " " & SortDirection & ";"
  Else
        MyCommon.QueryStr = "select Distinct LS.AppID, LS.PhraseId , LS.AppName, LastStartTime, LastEndTime, " & _
                        "LastTouchTime, LastAppErrorTime, RunFreq, IsNull(SendAlert, 0) SendAlert " & _
                        "from LastSync as LS with (NoLock) " & _
                        "left Join LastSyncEngines as LSE with (NoLock) on LSE.AppID=LS.AppID " & _
                        "left join PromoEngines as PE with (NoLock) on PE.EngineID=LSE.EngineID " & _
                        "where PE.Installed=1 or LSE.EngineID=-1 " & _
                        "order by " & SortText & " " & SortDirection & ";"
  End If
  
  dst = MyCommon.LRT_Select
  sizeOfData = dst.Rows.Count
  i = linesPerPage * PageNum
  
%>
<div id="intro">
  <h1 id="title">
    <% Sendb(Copient.PhraseLib.Lookup("term.agents", LanguageID))%>
  </h1>
  <div id="controls">
    <form action="<%Sendb("log-view.aspx?filetype=-1&amp;fileyear=" & Year(Today) & "&amp;filemonth=" & Month(Today) & "&amp;fileday=" & Day(Today)) %>" target="_blank" id="controlsform" name="controlsform" style="float: right;">
      <%
        If (Logix.UserRoles.AccessLogs = True) Then
          Send("<input type=""submit"" class=""regular"" id=""logs"" name=""logs"" value=""" & Copient.PhraseLib.Lookup("term.logs", LanguageID) & "..."" />")
        End If
        'If MyCommon.Fetch_SystemOption(75) Then
        '  If (Logix.UserRoles.AccessNotes) Then
        '    Send_NotesButton(16, 0, AdminUserID)
        '  End If
        'End If
      %>
    </form>
  </div>
</div>
<div id="main">
  <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
  <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.health", LanguageID)) %>">
    <thead>
      <tr>
        <th align="left" class="th-alert" scope="col">
          <a href="agent-list.aspx?searchterms=<%Sendb(Server.HtmlEncode(Request.QueryString("searchterms")))%>&amp;SortText=SendAlert&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.alert", LanguageID))%>
          </a>
          <%
            If SortText = "SendAlert" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            End If
          %>
        </th>
        <th align="left" class="th-name" scope="col">
          <a href="agent-list.aspx?searchterms=<%Sendb(Server.HtmlEncode(Request.QueryString("searchterms")))%>&amp;SortText=AppName&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.application", LanguageID))%>
          </a>
          <%
            If SortText = "AppName" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            End If
          %>
        </th>
        <th align="left" class="th-frequency" scope="col">
          <a href="agent-list.aspx?searchterms=<%Sendb(Server.HtmlEncode(Request.QueryString("searchterms")))%>&amp;SortText=RunFreq&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.frequency", LanguageID))%>
          </a>
          <%
            If SortText = "RunFreq" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            End If
          %>
        </th>
        <th align="left" class="th-datetime" scope="col">
          <a href="agent-list.aspx?searchterms=<%Sendb(Server.HtmlEncode(Request.QueryString("searchterms")))%>&amp;SortText=LastStartTime&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.laststarted", LanguageID))%>
          </a>
          <%
            If SortText = "LastStartTime" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            End If
          %>
        </th>
        <th align="left" class="th-datetime" scope="col">
          <a href="agent-list.aspx?searchterms=<%Sendb(Server.HtmlEncode(Request.QueryString("searchterms")))%>&amp;SortText=LastTouchTime&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.lasttouch", LanguageID))%>
          </a>
          <%
            If SortText = "LastTouchTime" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            End If
          %>
        </th>
        <th align="left" class="th-datetime" scope="col">
          <a href="agent-list.aspx?searchterms=<%Sendb(Server.HtmlEncode(Request.QueryString("searchterms")))%>&amp;SortText=LastEndTime&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.lastfinished", LanguageID))%>
          </a>
          <%
            If SortText = "LastEndTime" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            End If
          %>
        </th>
      </tr>
    </thead>
    <tbody>
      <%
          Dim WatchdogRows As DataRow()
          MyCommon.QueryStr = "dbo.pa_WDog_CheckOfflineApps"
          MyCommon.Open_LRTsp()
          MyCommon.LRTsp.Parameters.Add("@LastErrorTime", SqlDbType.DateTime).Value = DateTime.Now.AddSeconds(-3600)
          Watchdog = MyCommon.LRTsp_select()
          MyCommon.Close_LRTsp()

          If (dst.Rows.Count > 0) Then
              For Each row In dst.Rows
                  WatchdogRows = Watchdog.Select("AppID=" & MyCommon.NZ(row.Item("AppID"), 0))
                  Offline = (WatchdogRows.Length > 0)
                  If (Offline) Then
                      Send("<tr class=""shadedred"">")
                      If (shaded) Then
                          shaded = False
                      Else
                          shaded = True
                      End If
                      Offline = False
                  Else
                      If (shaded) Then
                          Send("<tr class=""shaded"">")
                          shaded = False
                      Else
                          Send("<tr>")
                          shaded = True
                      End If
                  End If
                  If (MyCommon.NZ(row.Item("SendAlert"), False) = True) Then
                      Send("  <td><a href=""?action=mute&amp;appid=" & MyCommon.NZ(row.Item("AppID"), 0) & """><img src=""../images/unmute.png"" border=""0"" alt=""" & Copient.PhraseLib.Lookup("health.mute", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("health.mute", LanguageID) & """ /></a></td>")
                  Else
                      Send("  <td><a href=""?action=unmute&amp;appid=" & MyCommon.NZ(row.Item("AppID"), 0) & """><img src=""../images/mute.png"" border=""0"" alt=""" & Copient.PhraseLib.Lookup("health.unmute", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("health.unmute", LanguageID) & """ /></a></td>")
                  End If
                  If (MyCommon.NZ(row.Item("AppID"), 0) = 111) Then
                      Send("  <td>" & Copient.PhraseLib.Lookup(IIf(IsDBNull(row.Item("Phraseid")), 367, row.Item("Phraseid")), MyCommon.GetAdminUser().LanguageID, "Phrase not Found") & "</td>")
                  Else
                      Send("  <td><a href=""agent-detail.aspx?appid=" & MyCommon.NZ(row.Item("AppID"), 0) & """>" & Copient.PhraseLib.Lookup(IIf(IsDBNull(row.Item("Phraseid")), 367, row.Item("Phraseid")), MyCommon.GetAdminUser().LanguageID, "Phrase not Found") & "</a></td>")
                  End If
                  Send("  <td>" & MyCommon.NZ(row.Item("RunFreq"), "&nbsp;") & "</td>")
                  If (Not IsDBNull(row.Item("LastStartTime"))) Then
                      Send("  <td>" & Logix.ToShortDateTimeString(MyCommon.NZ(row.Item("LastStartTime"), "1/1/1900"), MyCommon) & "</td>")
                  Else
                      Send("  <td>" & Copient.PhraseLib.Lookup("term.never", LanguageID) & "</td>")
                  End If
                  If (Not IsDBNull(row.Item("LastTouchTime"))) Then
                      Send("  <td>" & Logix.ToShortDateTimeString(MyCommon.NZ(row.Item("LastTouchTime"), "1/1/1900"), MyCommon) & "</td>")
                  Else
                      Send("  <td>" & Copient.PhraseLib.Lookup("term.never", LanguageID) & "</td>")
                  End If
                  If (Not IsDBNull(row.Item("LastEndTime"))) Then
                      Send("  <td>" & Logix.ToShortDateTimeString(MyCommon.NZ(row.Item("LastEndTime"), "1/1/1900"), MyCommon) & "</td>")
                  Else
                      Send("  <td>" & Copient.PhraseLib.Lookup("term.never", LanguageID) & "</td>")
                  End If
                  Send("</tr>")
              Next
          Else
          End If
      %>
    </tbody>
  </table>
</div>
<%
  'If MyCommon.Fetch_SystemOption(75) Then
  '  If (Logix.UserRoles.AccessNotes) Then
  '    Send_Notes(16, 0, AdminUserID)
  '  End If
  'End If
done:
  Send_BodyEnd()
  MyCommon.Close_LogixRT()
  MyCommon = Nothing
  Logix = Nothing
%>
