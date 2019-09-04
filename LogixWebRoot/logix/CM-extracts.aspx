<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:5.99.1.49448.Unstable Build - WUSJJ250009-8HO %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: CM-extracts.aspx 
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
  ' * Version : 5.10b1.0 
  ' *
  ' *****************************************************************************
  
  Dim CopientFileName As String = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
  Dim CopientFileVersion As String = "5.99.1.49448"
  Dim CopientProject As String = "Copient Logix"
  Dim CopientNotes As String = ""
  
  Dim AdminUserID As Long
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  
  Dim Shaded As String = "shaded"
  Dim idNumber As Integer
  Dim idSearch As String
  Dim idSearchText As String
  Dim PageNum As Integer = 0
  Dim linesPerPage As Integer = 20
  Dim sizeOfData As Integer
  Dim PrctSignPos As Integer
  Dim i As Integer = 0
  Dim dst As DataTable
  Dim dr As DataRow
  Dim drs() As DataRow
  Dim iPhraseId As Integer
  Dim iDayOfWeek As Integer
  Dim iHour As Integer
  Dim iMinute As Integer
  Dim bEnabled As Boolean
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "CM-extracts.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  If (Logix.UserRoles.AccessCMExtracts = False) Then
    Send_Denied(1, "perm.accesscmextracts")
    GoTo done
  End If
    
  Dim SortText As String = "ExtractID"
  Dim SortDirection As String
  
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

  MyCommon.QueryStr = "select  ExtractId, ExtractName, PhraseID, Display, AllowEdit, LastUpdate, Frequency, FilePath, Enabled, LastRunStart, LastRunFinish, " & _
                      "ScheduledRunDay, ScheduledRunHour, ScheduledRunMinute,'' as Day, '' as Time" & _
                      " from CM_Extract_Options with (NoLock) where Display=1"
  
  If (Request.QueryString("searchterms") <> "") Then
    If (Integer.TryParse(Request.QueryString("searchterms"), idNumber)) Then
      idSearch = idNumber.ToString
    Else
      idSearch = "-1"
    End If
    idSearchText = MyCommon.Parse_Quotes(Request.QueryString("searchterms"))
    PrctSignPos = idSearchText.IndexOf("%")
    If (PrctSignPos > -1) Then
      idSearch = -1
      idSearchText = idSearchText.Replace("%", "[%]")
    End If
    If (idSearchText.IndexOf("_") > -1) Then
      idSearchText = idSearchText.Replace("_", "[_]")
    End If
    MyCommon.QueryStr &= " and (ExtractID=" & idSearch & " or ExtractName like N'%" & idSearchText & "%')"
  End If
  MyCommon.QueryStr &= ";"
  
  dst = MyCommon.LRT_Select
  sizeOfData = dst.Rows.Count
  
  If sizeOfData > 0 Then
    For Each dr In dst.Rows
      If (Not IsDBNull(dr.Item("PhraseID"))) Then
        If Not Integer.TryParse(dr.Item("PhraseID"), iPhraseId) Then iPhraseId = 0
        If iPhraseId > 0 Then
          dr.Item("ExtractName") = Copient.PhraseLib.Lookup(iPhraseId, LanguageID, dr.Item("ExtractName"))
        End If
      End If
      If (IsDBNull(dr.Item("Enabled"))) Then
        bEnabled = True
      Else
        bEnabled = MyCommon.NZ(dr.Item("Enabled"), True)
      End If
      If bEnabled Then
        If (Not IsDBNull(dr.Item("Frequency"))) Then
          If MyCommon.NZ(dr.Item("Frequency"), 0) = 2 Then
            If IsDBNull(dr.Item("ScheduledRunDay")) Then
              iDayOfWeek = 7
            Else
              If Not Integer.TryParse(dr.Item("ScheduledRunDay"), iDayOfWeek) Then iDayOfWeek = 7
            End If
            If iDayOfWeek > 0 And iDayOfWeek < 8 Then
              Select Case iDayOfWeek
                Case 1
                  dr.Item("Day") = Copient.PhraseLib.Lookup("term.sunday", LanguageID)
                Case 2
                  dr.Item("Day") = Copient.PhraseLib.Lookup("term.monday", LanguageID)
                Case 3
                  dr.Item("Day") = Copient.PhraseLib.Lookup("term.tuesday", LanguageID)
                Case 4
                  dr.Item("Day") = Copient.PhraseLib.Lookup("term.wednesday", LanguageID)
                Case 5
                  dr.Item("Day") = Copient.PhraseLib.Lookup("term.thursday", LanguageID)
                Case 6
                  dr.Item("Day") = Copient.PhraseLib.Lookup("term.friday", LanguageID)
                Case 7
                  dr.Item("Day") = Copient.PhraseLib.Lookup("term.saturday", LanguageID)
              End Select
            End If
          End If
        End If
        If (IsDBNull(dr.Item("ScheduledRunHour"))) Then
          iHour = 0
        Else
          If Not Integer.TryParse(dr.Item("ScheduledRunHour"), iHour) Then iHour = 0
          If iHour > 23 Then
            iHour = 0
          End If
        End If
        If (IsDBNull(dr.Item("ScheduledRunMinute"))) Then
          iMinute = 0
        Else
          If Not Integer.TryParse(dr.Item("ScheduledRunMinute"), iMinute) Then iMinute = 0
          If iMinute > 59 Then
            iMinute = 0
          End If
        End If
        dr.Item("Time") = iHour.ToString("00") & ":" & iMinute.ToString("00")
      Else
        dr.Item("Day") = Copient.PhraseLib.Lookup("term.disabled", LanguageID)
      End If
    Next
  End If
  drs = dst.Select("", SortText & " " & SortDirection)

  ' set i
  i = linesPerPage * PageNum

  Send_HeadBegin("term.cmextracts")
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
  
  If (Logix.UserRoles.AccessCMExtracts = False) Then
    Send_Denied(1, "perm.accesscmextracts")
    GoTo done
  End If
%>
<form action="#" id="mainform" name="mainform">
  <div id="intro">
    <h1 id="title">
      <% Sendb(Copient.PhraseLib.Lookup("term.cmextracts", LanguageID))%>
    </h1>
  </div>
<div id="main">
  <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
  <% Send_Listbar(linesPerPage, sizeOfData, PageNum, Request.QueryString("searchterms"), "&SortText=" & SortText & "&SortDirection=" & SortDirection)%>
  <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.cmextracts", LanguageID)) %>">
    <thead>
      <tr>
        <th align="left" class="th-id" scope="col">
          <a href="CM-extracts.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&SortText=ExtractID&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>
          </a>
          <%If SortText = "ExtractID" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" class="th-name" scope="col">
          <a href="CM-extracts.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&SortText=ExtractName&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>
          </a>
          <%If SortText = "ExtractName" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" class="th-status" scope="col">
          <a href="CM-extracts.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&SortText=Frequency&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.frequency", LanguageID))%>
          </a>
          <%If SortText = "Frequencty" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" class="th-status" scope="col">
          <a href="CM-extracts.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&SortText=ScheduledRunDay&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.day", LanguageID))%>
          </a>
          <%If SortText = "ScheduledRunDay" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" class="th-status" scope="col">
            <% Sendb(Copient.PhraseLib.Lookup("term.time", LanguageID))%>
        </th>
        <th align="left" class="th-datetime" scope="col">
            <% Sendb(Copient.PhraseLib.Lookup("term.lastrunstart", LanguageID))%>
        </th>
        <th align="left" class="th-datetime" scope="col">
            <% Sendb(Copient.PhraseLib.Lookup("term.lastrunfinish", LanguageID))%>
        </th>
      </tr>
    </thead>
    <tbody>
      <%
        While (i < sizeOfData And i < linesPerPage + linesPerPage * PageNum)
          Send("<tr class=""" & Shaded & """>")
            
          Send("  <td>" & drs(i).Item("ExtractID") & "</td>")
          If (drs(i).Item("ExtractID") > 0) Then
            Sendb("  <td>")
            Send("<a href=""CM-extract-detail.aspx?ExtractID=" & drs(i).Item("ExtractID") & """>" & MyCommon.SplitNonSpacedString(drs(i).Item("ExtractName"), 30) & "</a></td>")
          Else
            Send("  <td>" & MyCommon.SplitNonSpacedString(drs(i).Item("ExtractName"), 30) & "</td>")
          End If
          If (Not IsDBNull(drs(i).Item("Frequency"))) Then
            If MyCommon.NZ(drs(i).Item("Frequency"), 0) = 1 Then
              Send("  <td>" & Copient.PhraseLib.Lookup("term.daily", LanguageID) & "</td>")
            Else
              Send("  <td>" & Copient.PhraseLib.Lookup("term.weekly", LanguageID) & "</td>")
            End If
          Else
            Send("  <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
          End If
          Send("  <td>" & drs(i).Item("Day") & "</td>")
          Send("  <td>" & drs(i).Item("Time") & "</td>")
          If (Not IsDBNull(drs(i).Item("LastRunStart"))) Then
            Send("  <td>" & Logix.ToShortDateTimeString(drs(i).Item("LastRunStart"), MyCommon) & "</td>")
          Else
            Send("  <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
          End If
          If (Not IsDBNull(drs(i).Item("LastRunFinish"))) Then
            Send("  <td>" & Logix.ToShortDateTimeString(drs(i).Item("LastRunFinish"), MyCommon) & "</td>")
          Else
            Send("  <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
          End If
          Send("</tr>")
          If Shaded = "shaded" Then
            Shaded = ""
          Else
            Shaded = "shaded"
          End If
          i = i + 1
        End While
      %>
    </tbody>
  </table>
</div>
</form>

<%
done:
  Send_BodyEnd()
  MyCommon.Close_LogixRT()
  Logix = Nothing
  MyCommon = Nothing
%>
