<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: reports-enhanced-list.aspx 
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
  Dim AdminUserID As Long
  Dim dt As System.Data.DataTable
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False

  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  Response.Expires = 0
  MyCommon.AppName = "reports-enhanced-list.aspx"
  
  Send_HeadBegin("term.reports")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
%>
<script type="text/javascript">
  function deleteReport(reportID) {
    var del = document.getElementById("delete");
    var tokenValues = [];
    var msg = '';

    msg = '<%Sendb(Copient.PhraseLib.Lookup("reports.DeleteReport", LanguageID))%>';
    tokenValues[0] = reportID;
    msg = detokenizeString(msg, tokenValues); 

    if (confirm(msg)) {
      del.value = reportID;
      document.mainform.submit();
    }
  }

  function getCSV(fileName) {
    var csv = document.getElementById("csv");
    if (fileName != "") {
      csv.value = fileName;
      document.mainform.submit();
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
  Send_Subtabs(Logix, 8, 8)
  
  If (Logix.UserRoles.AccessReports = False) Then
    Send_Denied(1, "perm.admin-reports")
    GoTo done
  End If
  
  Dim FilePath As String = Trim(MyCommon.Fetch_SystemOption(114))
  If Not (Right(FilePath, 1) = "\") Then
    FilePath = FilePath & "\"
  End If
  
  If (Request.Form("delete") <> "") Then
    MyCommon.QueryStr = "dbo.pt_Reports_Delete"
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add("@ReportID", SqlDbType.Int).Value = Request.Form.Item("delete")
    MyCommon.LRTsp.ExecuteNonQuery()
    MyCommon.Close_LRTsp()
    MyCommon.Activity_Log(29, Request.Form.Item("delete"), AdminUserID, Copient.PhraseLib.Lookup("history.report-delete", LanguageID))
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "reports-enhanced-list.aspx")
  ElseIf (Request.Form("csv") <> "") Then
    FilePath = FilePath & Request.Form("csv")
    Dim File As System.IO.FileInfo = New System.IO.FileInfo(FilePath)
    If File.Exists Then
      Response.Clear()
      Response.AddHeader("Content-Disposition", "attachment; filename=" & File.Name)
      Response.AddHeader("Content-Length", File.Length.ToString())
      Response.ContentType = "application/octet-stream"
      Response.WriteFile(File.FullName)
      Response.End()
    End If
  End If
  
  Dim PageNum As Integer = 0
  PageNum = Request.QueryString("pagenum")
  If PageNum < 0 Then
    PageNum = 0
  End If
  Dim MorePages As Boolean = False
  
  Dim SortText As String = "ReportID"
  If (Request.QueryString("SortText") <> "") Then
    SortText = Request.QueryString("SortText")
  End If
  
  Dim SortDirection As String = "DESC"
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
  
  Dim idNumber As Integer = 0
  Dim idSearch As String = ""
  Dim idSearchText As String = ""
  Dim linesPerPage As Integer = 20
  Dim sizeOfData As Integer = 0
  Dim i As Integer = 0
  MyCommon.QueryStr = "select R.*, RT.Name as ReportType, RT.PhraseID as ReportTypePhraseID, " & _
                      "  RS.Name as ReportStatus, RS.PhraseID as ReportStatusPhraseID, AU.UserName, AU.FirstName, AU.LastName " & _
                      "from Reports as R with (NoLock) " & _
                      "left join ReportTypes as RT on R.ReportTypeID=RT.ReportTypeID " & _
                      "left join ReportStatuses as RS on R.ReportStatusID=RS.ReportStatusID " & _
                      "left join AdminUsers as AU on R.AdminUserID=AU.AdminUserID " & _
                      "where R.Deleted=0"
  If (Request.QueryString("searchterms") <> "") Then
    If (Integer.TryParse(Request.QueryString("searchterms"), idNumber)) Then
      idSearch = idNumber.ToString
    Else
      idSearch = "-1"
    End If
    idSearchText = MyCommon.Parse_Quotes(Request.QueryString("searchterms"))
    Dim PrctSignPos As Integer = idSearchText.IndexOf("%")
    If (PrctSignPos > -1) Then
      idSearch = -1
      idSearchText = idSearchText.Replace("%", "[%]")
    End If
    If (idSearchText.IndexOf("_") > -1) Then
      idSearchText = idSearchText.Replace("_", "[_]")
    End If
    MyCommon.QueryStr &= " and (R.Name like N'%" & idSearchText & "%' or UserName like N'%" & idSearchText & "%')"
  End If
  MyCommon.QueryStr &= " order by " & SortText & " " & SortDirection
  dt = MyCommon.LRT_Select
  sizeOfData = dt.Rows.Count
  i = linesPerPage * PageNum
  
  If (sizeOfData = 1 AndAlso Request.QueryString("searchterms") <> "") Then
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "reports-enhanced-request.aspx?ReportID=" & dt.Rows(0).Item("ReportID"))
  End If
%>
<div id="intro">
  <h1 id="title">
    <% Sendb(Copient.PhraseLib.Lookup("term.enhanced", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.reports", LanguageID), VbStrConv.Lowercase))%>
  </h1>
  <div id="controls">
    <form action="#" id="controlsform" name="controlsform">
      <%
        Send("<input type=""button"" name=""new"" id=""new"" class=""regular"" value=""" & Copient.PhraseLib.Lookup("term.new", LanguageID) & """ onclick=""location.href='reports-enhanced-request.aspx'"" />")
      %>
    </form>
  </div>
</div>
<div id="main">
  <%
    If (infoMessage <> "") Then
      Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
    End If
    Send_Listbar(linesPerPage, sizeOfData, PageNum, Request.QueryString("searchterms"), "&SortText=" & SortText & "&SortDirection=" & SortDirection, False)
  %>
  <form action="reports-enhanced-list.aspx" method="post" id="mainform" name="mainform">
    <input type="hidden" id="csv" name="csv" value="" />
    <input type="hidden" id="delete" name="delete" value="" />
    <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.reports", LanguageID)) %>">
      <thead>
        <tr>
          <th align="left" class="th-button" scope="col" style="text-align: center;">
            <% Sendb(Left(Copient.PhraseLib.Lookup("term.remove", LanguageID), 3))%>
          </th>
          <th align="left" class="th-id" scope="col">
            <a href="reports-enhanced-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=ReportID&amp;SortDirection=<% Sendb(SortDirection) %>">
              <% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>
            </a>
            <%
              If SortText = "ReportID" Then
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
            <a href="reports-enhanced-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=Name&amp;SortDirection=<% Sendb(SortDirection) %>">
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
          <th align="left" class="th-category" scope="col">
            <a href="reports-enhanced-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=ReportTypeID&amp;SortDirection=<% Sendb(SortDirection) %>">
              <% Sendb(Copient.PhraseLib.Lookup("term.type", LanguageID))%>
            </a>
            <%
              If SortText = "ReportTypeID" Then
                If SortDirection = "ASC" Then
                  Sendb("<span class=""sortarrow"">&#9660;</span>")
                Else
                  Sendb("<span class=""sortarrow"">&#9650;</span>")
                End If
              Else
              End If
            %>
          </th>
          <th align="left" class="th-date" scope="col">
            <a href="reports-enhanced-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=Requested&amp;SortDirection=<% Sendb(SortDirection) %>">
              <% Sendb(Copient.PhraseLib.Lookup("term.requested", LanguageID))%>
            </a>
            <%
              If SortText = "Requested" Then
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
            <a href="reports-enhanced-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=Username&amp;SortDirection=<% Sendb(SortDirection) %>">
              <% Sendb(Copient.PhraseLib.Lookup("term.requester", LanguageID))%>
            </a>
            <%
              If SortText = "Username" Then
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
            <a href="reports-enhanced-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=ReportStatusID&amp;SortDirection=<% Sendb(SortDirection) %>">
              <% Sendb(Copient.PhraseLib.Lookup("term.status", LanguageID))%>
            </a>
            <%
              If SortText = "ReportStatusID" Then
                If SortDirection = "ASC" Then
                  Sendb("<span class=""sortarrow"">&#9660;</span>")
                Else
                  Sendb("<span class=""sortarrow"">&#9650;</span>")
                End If
              Else
              End If
            %>
          </th>
          <th class="th-button" align="left" scope="col" style="text-align:center;">
            <% Sendb(Copient.PhraseLib.Lookup("term.csv", LanguageID))%>
          </th>
          <th class="th-button" align="left" scope="col" style="text-align:center;">
            <% Sendb(Copient.PhraseLib.Lookup("term.view", LanguageID))%>
          </th>
        </tr>
      </thead>
      <tbody>
        <%
          Dim downloadEnabled As Boolean = False
          Dim Shaded As String = "shaded"
          While (i < sizeOfData And i < linesPerPage + linesPerPage * PageNum)
            Send("<tr class=""" & Shaded & """>")
            Send("  <td style=""text-align: center;""><input type=""button"" value=""X"" title=""Remove"" name=""ex"" id=""ex" & dt.Rows(i).Item("ReportID") & """ class=""ex"" onclick=""javascript:deleteReport(" & dt.Rows(i).Item("ReportID") & ");"" /></td>")
            Send("  <td>" & dt.Rows(i).Item("ReportID") & "</td>")
                Sendb("  <td><a href=""reports-enhanced-request.aspx?ReportID=" & dt.Rows(i).Item("ReportID") & "&ReportTypeID=" & dt.Rows(i).Item("ReportTypeID") & """>")
            If MyCommon.NZ(dt.Rows(i).Item("Name"), "") <> "" Then
              Sendb(MyCommon.SplitNonSpacedString(dt.Rows(i).Item("Name"), 25))
            Else
              Sendb("<i>" & Copient.PhraseLib.Lookup("term.unnamed", LanguageID) & "</i>")
            End If
            Send("</a></td>")
            If MyCommon.NZ(dt.Rows(i).Item("ReportTypePhraseID"), 0) > 0 Then
              Sendb("  <td>" & Copient.PhraseLib.Lookup(dt.Rows(i).Item("ReportTypePhraseID"), LanguageID) & "</td>")
            Else
              Sendb("  <td>" & MyCommon.NZ(dt.Rows(i).Item("ReportType"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)) & "</td>")
            End If
            If IsDBNull(dt.Rows(i).Item("Requested")) Then
              Send("  <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
            Else
              Send("  <td>" & Logix.ToShortDateTimeString(dt.Rows(i).Item("Requested"), MyCommon) & "</td>")
            End If
            Sendb("  <td>")
            If (Logix.UserRoles.ViewOthersInfo AndAlso MyCommon.NZ(dt.Rows(i).Item("AdminUserID"), 0) > 0) Then
              Sendb("<a href=""user-edit.aspx?UserID=" & dt.Rows(i).Item("AdminUserID") & """>" & MyCommon.NZ(dt.Rows(i).Item("UserName"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)) & "</a>")
            Else
              Sendb(MyCommon.NZ(dt.Rows(i).Item("UserName"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)))
            End If
            Send("</td>")
            Sendb("  <td")
            Select Case MyCommon.NZ(dt.Rows(i).Item("ReportStatusID"), 0)
              Case 1
                Sendb(" style=""color:#cc0000;"">")
                downloadEnabled = False
              Case 2
                Sendb(" style=""color:#cc9900;"">")
                downloadEnabled = False
              Case 3
                Sendb(" style=""color:#00cc00;"">")
                downloadEnabled = True
              Case Else
                Sendb(">")
                downloadEnabled = True
            End Select
            If (MyCommon.NZ(dt.Rows(i).Item("FileName"), "") = "") Then
              downloadEnabled = False
            End If
            If MyCommon.NZ(dt.Rows(i).Item("ReportStatusPhraseID"), 0) > 0 Then
              Sendb(Copient.PhraseLib.Lookup(dt.Rows(i).Item("ReportStatusPhraseID"), LanguageID))
            Else
              Sendb(MyCommon.NZ(dt.Rows(i).Item("ReportStatus"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)))
            End If
            Send("</td>")
            Send("<td style=""text-align:center;""><input type=""button""" & IIf(downloadEnabled, "", " disabled=""disabled""") & " class=""view"" id=""csv" & dt.Rows(i).Item("ReportID") & """ name=""csv"" onclick=""javascript:getCSV('" & MyCommon.NZ(dt.Rows(i).Item("FileName"), "") & "');"" title=""" & Copient.PhraseLib.Lookup("term.download", LanguageID) & " " & MyCommon.NZ(dt.Rows(i).Item("FileName"), "") & """ value=""▼"" /></td>")
            Send("<td style=""text-align:center;""><input type=""button""" & IIf(downloadEnabled, "", " disabled=""disabled""") & " class=""view"" id=""view" & dt.Rows(i).Item("ReportID") & """ name=""view"" onclick=""javscript:window.open('reports-enhanced-viewer.aspx?ReportID=" & dt.Rows(i).Item("ReportID") & "', '_blank');"" title=""" & Copient.PhraseLib.Lookup("term.view", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.report", LanguageID), VbStrConv.Lowercase) & """ value=""►"" /></td>")
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
  </form>
</div>

<%
done:
  Send_BodyEnd("searchform", "searchterms")
  MyCommon.Close_LogixRT()
  Logix = Nothing
  MyCommon = Nothing
%>