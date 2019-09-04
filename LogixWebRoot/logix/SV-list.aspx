<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: SV-list.aspx 
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
  Dim Shaded As String = "shaded"
  Dim idNumber As Integer
  Dim idSearch As String
  Dim idSearchText As String
  Dim PageNum As Integer = 0
  Dim MorePages As Boolean
  Dim linesPerPage As Integer = 20
  Dim sizeOfData As Integer
  Dim i As Integer = 0
  Dim PrctSignPos As Integer
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "SV-list.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  PageNum = Request.QueryString("pagenum")
  If PageNum < 0 Then PageNum = 0
  MorePages = False
  
  Send_HeadBegin("term.storedvalueprogram")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
  Send_HeadEnd()
  Send_BodyBegin(1)
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  Send_Tabs(Logix, 5)
  Send_Subtabs(Logix, 50, 2)
  
  If (Logix.UserRoles.AccessStoredValuePrograms = False) Then
    Send_Denied(1, "perm.storedvalue-access")
    GoTo done
  End If
  
  Dim SortText As String = "SVProgramID"
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
      SortDirection = "DESC"
    End If
  Else
    SortDirection = Request.QueryString("SortDirection")
  End If
  
  ' any search terms inbound?
    If (Server.HtmlEncode(Request.QueryString("searchterms")) <> "") Then
        If (Integer.TryParse(Server.HtmlEncode(Request.QueryString("searchterms")), idNumber)) Then
      idSearch = idNumber.ToString
    Else
      idSearch = "-1"
    End If
        idSearchText = MyCommon.Parse_Quotes(Server.HtmlEncode(Request.QueryString("searchterms")))
    PrctSignPos = idSearchText.IndexOf("%")
    If (PrctSignPos > -1) Then
      idSearch = -1
      idSearchText = idSearchText.Replace("%", "[%]")
    End If
        If (idSearchText.IndexOf("_") > -1) Then idSearchText = idSearchText.Replace("_", "[_]")
        If (idSearchText.IndexOf("&amp;") > -1) Then idSearchText = idSearchText.Replace("&amp;", "&")
        MyCommon.QueryStr = "SELECT SVProgramID AS ProgramID, Name AS ProgramName, SVP.SVTypeID, PT.Phrase as Type, Value, SVP.CreatedDate, SVP.LastUpdate " &
                            "FROM StoredValuePrograms AS SVP WITH (NoLock) " &
                            "LEFT JOIN SVTypes AS SVT ON SVT.SVTypeID=SVP.SVTypeID " &
                            "LEFT JOIN PhraseText AS PT ON PT.PhraseID=SVT.PhraseID " &
                            "WHERE Deleted=0 AND Visible=1 AND PT.LanguageID=" & LanguageID & " AND (SVProgramID =" & idSearch & " OR Name LIKE N'%" & Server.HtmlDecode(idSearchText).Replace("'", "''") & "%') ORDER BY " & SortText & " " & SortDirection
  Else
    ' no search strings, just query
    MyCommon.QueryStr = "SELECT SVProgramID AS ProgramID, Name AS ProgramName, SVP.SVTypeID, PT.Phrase as Type, Value, SVP.CreatedDate, SVP.LastUpdate " & _
                        "FROM StoredValuePrograms AS SVP WITH (NoLock) " & _
                        "LEFT JOIN SVTypes AS SVT ON SVT.SVTypeID=SVP.SVTypeID " & _
                        "LEFT JOIN PhraseText AS PT ON PT.PhraseID=SVT.PhraseID " & _
                        "WHERE Deleted=0 AND Visible=1 AND PT.LanguageID=" & LanguageID & " ORDER BY " & SortText & " " & SortDirection
  End If
  
  'Response.Write("Testing query -> " & MyCommon.QueryStr & " <- against database <br />")
  dst = MyCommon.LRT_Select
  sizeOfData = dst.Rows.Count
  ' set page variable(s)
  i = linesPerPage * PageNum
    If (sizeOfData = 1 AndAlso Server.HtmlEncode(Request.QueryString("searchterms")) <> "") Then
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "sv-edit.aspx?ProgramGroupID=" & dst.Rows(i).Item("ProgramID"))
  End If
%>
<div id="intro">
  <h1 id="title">
    <% Sendb(Copient.PhraseLib.Lookup("term.storedvalueprograms", LanguageID))%>
  </h1>
  <div id="controls">
    <form action="SV-edit.aspx" id="controlsform" name="controlsform">
      <%
        If (Logix.UserRoles.CreateStoredValuePrograms) Then
          Send_New()
        End If
        'If MyCommon.Fetch_SystemOption(75) Then
        '  If (Logix.UserRoles.AccessNotes) Then
        '    Send_NotesButton(9, 0, AdminUserID)
        '  End If
        'End If
      %>
    </form>
  </div>
</div>
<div id="main">
  <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
  <%  Send_Listbar(linesPerPage, sizeOfData, PageNum, Server.HtmlEncode(Request.QueryString("searchterms")), "&SortText=" & SortText & "&SortDirection=" & SortDirection)%>
  <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.points", LanguageID)) %>">
    <thead>
      <tr>
        <th align="left" class="th-id" scope="col">
          <a href="SV-list.aspx?searchterms=<%Sendb(Server.HtmlEncode(Request.QueryString("searchterms")))%>&SortText=SVProgramID&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>
          </a>
          <%
            If SortText = "SVProgramID" Then
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
          <a href="SV-list.aspx?searchterms=<%Sendb(Server.HtmlEncode(Request.QueryString("searchterms")))%>&SortText=Name&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>
          </a>
          <%
            If SortText = "Name" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" scope="col" style="width:90px;">
          <a href="SV-list.aspx?searchterms=<%Sendb(Server.HtmlEncode(Request.QueryString("searchterms")))%>&SortText=SVTypeID&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.type", LanguageID))%>
          </a>
          <%
            If SortText = "SVTypeID" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" class="th-datetime" scope="col">
          <a href="SV-list.aspx?searchterms=<%Sendb(Server.HtmlEncode(Request.QueryString("searchterms")))%>&SortText=CreatedDate&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.created", LanguageID))%>
          </a>
          <%
            If SortText = "CreatedDate" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" class="th-datetime" scope="col">
          <a href="SV-list.aspx?searchterms=<%Sendb(Server.HtmlEncode(Request.QueryString("searchterms")))%>&SortText=LastUpdate&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.edited", LanguageID))%>
          </a>
          <%
            If SortText = "LastUpdate" Then
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
        While (i < sizeOfData And i < linesPerPage + linesPerPage * PageNum)
          Send("<tr class=""" & Shaded & """>")
          Send("  <td>" & dst.Rows(i).Item("ProgramID") & "</td>")
          Send("  <td><a href=""sv-edit.aspx?ProgramGroupID=" & dst.Rows(i).Item("ProgramID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("ProgramName"), 30) & "</a></td>")
          Send("  <td>" & MyCommon.NZ(dst.Rows(i).Item("Type"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)) & "</td>")
          If (Not IsDBNull(dst.Rows(i).Item("CreatedDate"))) Then
            Send("  <td>" & Logix.ToShortDateTimeString(dst.Rows(i).Item("CreatedDate"), MyCommon) & "</td>")
          Else
            Send("  <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
          End If
          If (Not IsDBNull(dst.Rows(i).Item("LastUpdate"))) Then
            Send("  <td>" & Logix.ToShortDateTimeString(dst.Rows(i).Item("LastUpdate"), MyCommon) & "</td>")
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
<%
  'If MyCommon.Fetch_SystemOption(75) Then
  '  If (Logix.UserRoles.AccessNotes) Then
  '    Send_Notes(9, 0, AdminUserID)
  '  End If
  'End If
done:
  Send_BodyEnd("searchform", "searchterms")
  MyCommon.Close_LogixRT()
  Logix = Nothing
  MyCommon = Nothing
%>
