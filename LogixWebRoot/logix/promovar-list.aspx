<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: promovar-list.aspx 
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
    Dim MyCryptLib As New Copient.CryptLib
  Dim Logix As New Copient.LogixInc
  Dim dst As System.Data.DataTable
  Dim row As System.Data.DataRow
  Dim sortedRows As System.Data.DataRow()
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
  MyCommon.AppName = "promovar-list.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  PageNum = Request.QueryString("pagenum")
  If PageNum < 0 Then PageNum = 0
  MorePages = False
  
  Send_HeadBegin("term.promovars")
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
  Send_Subtabs(Logix, 50, 3)

  
  If (Logix.UserRoles.AccessPointsPrograms = False) Then
    Send_Denied(1, "perm.points-access")
    GoTo done
  End If
  
  Dim SortText As String = "PromoVarID"
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
    If (idSearchText.IndexOf("_") > -1) Then idSearchText = idSearchText.Replace("_", "[_]")
    MyCommon.QueryStr = "select PV.PromoVarID, PV.ExternalID, PV.Name, PVT.PhraseID as TypePhraseID, '' as TypePhrase, PVT.Description as TypeDesc, PV.CreatedDate, PV.LastUpdate " & _
                        "from PromoVariables PV with (NoLock) inner join PromoVarTypes PVT with (NoLock) on PVT.TypeID = PV.VarTypeID " & _
                        "where PV.Deleted=0 AND (PromoVarID =" & idSearch & " or Name like N'%" & idSearchText & "%')" ' order by " & SortText & " " & SortDirection
  Else
    ' no search strings, just query
    MyCommon.QueryStr = "select PV.PromoVarID, PV.ExternalID, PV.Name, PVT.PhraseID as TypePhraseID, '' as TypePhrase, PVT.Description as TypeDesc, PV.CreatedDate, PV.LastUpdate " & _
                        "from PromoVariables PV with (NoLock) inner join PromoVarTypes PVT with (NoLock) on PVT.TypeID = PV.VarTypeID " & _
                        "where PV.Deleted=0" ' order by " & SortText & " " & SortDirection
  End If
  
  'Response.Write("Testing query -> " & MyCommon.QueryStr & " <- against database <br />")
  dst = MyCommon.LXS_Select
  
  ' Get the phrase text for each of the Phrase IDs 
  For Each row In dst.Rows
    If (IsDBNull(row.Item("TypePhraseID")) OrElse MyCommon.NZ(row.Item("TypePhraseID"), 0) = 0) Then
      row.Item("TypePhrase") = MyCommon.NZ(row.Item("TypeDesc"), "")
    Else
      row.Item("TypePhrase") = Copient.PhraseLib.Lookup(row.Item("TypePhraseID"), LanguageID)
    End If
  Next

  sortedRows = dst.Select("", SortText & " " & SortDirection)
  sizeOfData = sortedRows.Length
  ' set page variable(s)
  i = linesPerPage * PageNum
  
  If (sizeOfData = 1 AndAlso Request.QueryString("searchterms") <> "") Then
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "promovar-edit.aspx?PromoVarID=" & sortedRows(i).Item("PromoVarID"))
  End If
%>
<div id="intro">
  <h1 id="title">
    <% Sendb(Copient.PhraseLib.Lookup("term.promovars", LanguageID))%>
  </h1>
  <div id="controls">
    <form action="promovar-edit.aspx" id="controlsform" name="controlsform">
    <%
      Send_New()
      'If MyCommon.Fetch_SystemOption(75) Then
      '  If (Logix.UserRoles.AccessNotes) Then
      '    Send_NotesButton(10, 0, AdminUserID)
      '  End If
      'End If
    %>
    </form>
  </div>
</div>
<div id="main">
  <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
  <% Send_Listbar(linesPerPage, sizeOfData, PageNum, Request.QueryString("searchterms"), "&SortText=" & SortText & "&SortDirection=" & SortDirection)%>
  <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.promovars", LanguageID)) %>">
    <thead>
      <tr>
        <th align="left" class="th-id" scope="col">
          <a href="promovar-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&SortText=PromoVarID&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>
          </a>
          <%If SortText = "PromoVarID" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" class="th-id" scope="col">
          <a href="promovar-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&SortText=ExternalID&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.xid", LanguageID))%>
          </a>
          <%If SortText = "ExternalID" Then
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
          <a href="promovar-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&SortText=Name&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>
          </a>
          <%If SortText = "Name" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" class="th-type" scope="col">
          <a href="promovar-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&SortText=TypePhrase&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.type", LanguageID))%>
          </a>
          <%If SortText = "TypePhrase" Then
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
          <a href="promovar-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&SortText=CreatedDate&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.created", LanguageID))%>
          </a>
          <%If SortText = "CreatedDate" Then
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
          <a href="promovar-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&SortText=LastUpdate&SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.edited", LanguageID))%>
          </a>
          <%If SortText = "LastUpdate" Then
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
          Send("  <td>" & MyCommon.NZ(sortedRows(i).Item("PromoVarID"), "") & "</td>")
              Send("  <td>" & MyCryptLib.SQL_StringDecrypt(sortedRows(i).Item("ExternalID").ToString()) & "</td>")
          Send("  <td><a href=""promovar-edit.aspx?PromoVarID=" & dst.Rows(i).Item("PromoVarID") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(dst.Rows(i).Item("Name"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)), 25) & "</a></td>")
          Send("  <td>" & MyCommon.NZ(sortedRows(i).Item("TypePhrase"), "") & "</td>")

          If (Not IsDBNull(sortedRows(i).Item("CreatedDate"))) Then
            Send("  <td>" & Format(sortedRows(i).Item("CreatedDate"), "dd MMM yyyy, HH:mm:ss") & "</td>")
          Else
            Send("  <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
          End If
          If (Not IsDBNull(sortedRows(i).Item("LastUpdate"))) Then
            Send("  <td>" & Format(sortedRows(i).Item("LastUpdate"), "dd MMM yyyy, HH:mm:ss") & "</td>")
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
  '    Send_Notes(10, 0, AdminUserID)
  '  End If
  'End If
done:
  Send_BodyEnd("searchform", "searchterms")
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  Logix = Nothing
  MyCommon = Nothing
%>
