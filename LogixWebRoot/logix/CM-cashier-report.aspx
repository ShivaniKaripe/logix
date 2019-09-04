<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: CM-cashier-report.aspx 
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
  Dim CopientFileVersion As String = "7.3.1.138972"
  Dim CopientProject As String = "Copient Logix"
  Dim CopientNotes As String = ""
  
  Dim AdminUserID As Long
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim dt As DataTable
  Dim rst As DataTable
  Dim ExtCardID As String = ""
  Dim FirstName As String = ""
  Dim MiddleName As String = ""
  Dim LastName As String = ""
  Dim FullName As String = ""
  Dim IsHouseholdID As Boolean = False
  Dim sizeOfData As Integer
  Dim i As Integer = 0
  Dim x As Integer = 0
  Dim idSearch As String
  Dim idSearchText As String
  Dim PrctSignPos As Integer
  Dim PageNum As Integer = 0
  Dim MorePages As Boolean
  Dim linesPerPage As Integer = 20
  Dim Shaded As String = "shaded"
  Dim restrictLinks As Boolean = False
  Dim extraLink As String = ""
  Dim InfoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim sSearchQuery As String
  
  Dim WhereClause As String = ""
  Dim WhereBuf As New StringBuilder()
  Dim AdvSearchSQL As String = ""
  Dim CriteriaMsg As String = ""
  Dim CriteriaTokens As String = ""
  Dim CriteriaError As Boolean = False

  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  
  Response.Expires = 0
  MyCommon.AppName = "CM-cashier-report.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  PageNum = Request.QueryString("pagenum")
  If PageNum < 0 Then PageNum = 0
  MorePages = False
  
  
  Dim SortText As String = "ActivityDate"
  Dim SortDirection As String = ""
  
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
  
  ' lets check the logged in user and see if they are to be restricted to this page
  MyCommon.QueryStr = "select AU.StartPageID,AUSP.PageName ,AUSP.prestrict from AdminUsers as AU with (NoLock) " & _
                      "inner join AdminUserStartPages as AUSP with (NoLock) on AU.StartPageID=AUSP.StartPageID " & _
                      "where AU.AdminUserID=" & AdminUserID
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    If (MyCommon.NZ(rst.Rows(0).Item("prestrict"), False) = True) Then
      ' ok we got in here then we need to restrict the user from seeing any other pages
      restrictLinks = True
    End If
  End If
  
 
  Send_HeadBegin("term.customer", "term.history")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  If restrictLinks Then
    Send("<link rel=""stylesheet"" href=""../css/logix-restricted.css"" type=""text/css"" media=""screen"" />")
  End If
  Send_Scripts()
%>

<script type="text/javascript">
  
  function launchAdvSearch() {
    self.name = "CahierReportWin";
    openPopup("CM-cashier-search.aspx");
  }
  
  function editSearchCriteria() {
    var tokenStr = document.frmIter.advTokens.value;
    
    self.name = "CahierReportWin";
    openPopup("CM-cashier-search.aspx?tokens=" + tokenStr);
  }
</script>

<%
  Send_HeadEnd()
  Send_BodyBegin(1)
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  If Not restrictLinks Then
    Send_Tabs(Logix, 3)
  End If

  Send_Subtabs(Logix, 30, 3)
  
  If (Logix.UserRoles.ViewHistory = False) Then
    Send_Denied(1, "perm.admin-history")
    GoTo done
  End If
  
  ' handle an Advance Search Criteria
  If (Request.Form("mode") = "advancedsearch") Then
    Dim TempStr As String = ""
    Dim CritBuf As New StringBuilder()
    Dim CritTokenBuf As New StringBuilder()
    
    If (Request.Form("cashier").Trim <> "" AndAlso Request.Form("cashierOption") <> "") Then
      WhereBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
      WhereBuf.Append(GetOptionString(MyCommon, Integer.Parse(Request.Form("cashierOption")), Request.Form("cashier"), "AL.LinkID3"))
      If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
      CritBuf.Append(Copient.PhraseLib.Lookup("term.cashier", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("cashierOption"))) & " '" & Request.Form("cashier").Trim & "'")
      CritTokenBuf.Append("CASHIER," & Integer.Parse(Request.Form("cashierOption")) & "," & Request.Form("cashier").Trim & ",|")
    End If
    
    If (Request.Form("store").Trim <> "" AndAlso Request.Form("storeOption") <> "") Then
      WhereBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
      WhereBuf.Append(GetOptionString(MyCommon, Integer.Parse(Request.Form("storeOption")), Request.Form("store"), "AL.LinkID4"))
      If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
      CritBuf.Append(Copient.PhraseLib.Lookup("term.store", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("storeOption"))) & " '" & Request.Form("store").Trim & "'")
      CritTokenBuf.Append("STORE," & Integer.Parse(Request.Form("storeOption")) & "," & Request.Form("store").Trim & ",|")
    End If
    
    If (Request.Form("cardid").Trim <> "" AndAlso Request.Form("cardidOption") <> "") Then
      WhereBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
      WhereBuf.Append(GetOptionString(MyCommon, Integer.Parse(Request.Form("cardidOption")), Request.Form("cardid"), "AL.LinkID5"))
      If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
      CritBuf.Append(Copient.PhraseLib.Lookup("term.card", LanguageID) & " " & Copient.PhraseLib.Lookup("term.id", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("cardidOption"))) & " '" & Request.Form("cardid").Trim & "'")
      CritTokenBuf.Append("CARDID," & Integer.Parse(Request.Form("cardidOption")) & "," & Request.Form("cardid").Trim & ",|")
    End If
    
    If (Request.Form("cardtype").Trim <> "" AndAlso Request.Form("cardtypeOption") <> "") Then
      WhereBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
      WhereBuf.Append(GetOptionString(MyCommon, Integer.Parse(Request.Form("cardtypeOption")), Request.Form("cardtype"), "AL.LinkID6"))
      If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
      CritBuf.Append(Copient.PhraseLib.Lookup("term.id", LanguageID) & " " & Copient.PhraseLib.Lookup("term.type", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("cardtypeOption"))) & " '" & Request.Form("cardtype").Trim & "'")
      CritTokenBuf.Append("CARDTYPE," & Integer.Parse(Request.Form("cardtypeOption")) & "," & Request.Form("cardtype").Trim & ",|")
    End If
      
    If (Request.Form("desc").Trim <> "" AndAlso Request.Form("descOption") <> "") Then
      WhereBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
      WhereBuf.Append(GetOptionString(MyCommon, Integer.Parse(Request.Form("descOption")), Request.Form("desc"), "Convert(nvarchar(1000),Description)"))
      If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
      CritBuf.Append(Copient.PhraseLib.Lookup("term.action", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("descOption"))) & " '" & Request.Form("desc").Trim & "'")
      CritTokenBuf.Append("DESC," & Integer.Parse(Request.Form("descOption")) & "," & Request.Form("desc").Trim & ",|")
    End If
    
    If (Request.Form("datetimeDate1").Trim <> "" AndAlso Request.Form("datetimeOption") <> "") Then
      Try
        TempStr = GetDateOption(MyCommon, Integer.Parse(Request.Form("datetimeOption")), Request.Form("datetimeDate1"), Request.Form("datetimeDate2"), "AL.ActivityDate")

        If (TempStr <> "") Then
          WhereBuf.Append(" and " & TempStr)
          If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
          If (TempStr.IndexOf("between") > -1) Then
            CritBuf.Append(Copient.PhraseLib.Lookup("term.date", LanguageID) & " " & GetDateOptionType(Integer.Parse(Request.Form("datetimeOption"))) & " '" & Request.Form("datetimeDate1").Trim & "'")
            If Request.Form("datetimeDate2").Trim <> "" Then
              CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " '" & Request.Form("datetimeDate2").Trim & "'")
            End If
          Else
            CritBuf.Append(Copient.PhraseLib.Lookup("term.date", LanguageID) & " " & GetDateOptionType(Integer.Parse(Request.Form("datetimeOption"))) & " '" & Request.Form("datetimeDate1").Trim & "'")
            CritTokenBuf.Append("DATETIME," & Integer.Parse(Request.Form("datetimeOption")) & "," & Request.Form("datetimeDate1").Trim & ",|")
          End If
        End If
      Catch aex As ApplicationException
        CriteriaError = True
        TempStr = ""
        CriteriaMsg = aex.Message
      End Try
    End If
    
    CriteriaMsg &= CritBuf.ToString
    CriteriaTokens = CritTokenBuf.ToString
  End If

  sSearchQuery = "select AU.FirstName, AU.LastName, AL.ActivityDate, AL.Description, AL.LinkID2, AL.LinkID3, AL.LinkID4, AL.LinkID5, AL.LinkID6, AL.ActivitySubTypeID from ActivityLog as AL with (NoLock) " & _
                 "left join AdminUsers as AU with (NoLock) on AL.AdminID=AU.AdminUserID " & _
                 "where ActivityTypeID='25' and ActivitySubTypeID='2' and AL.LinkID3 is not null "
  
  If (Request.Form("mode") = "advancedsearch" OrElse Request.Form("advSql") <> "") Then
    If (Request.Form("advSql") <> "") Then WhereBuf = New StringBuilder(Server.UrlDecode(Request.Form("advSql")))
    If (Request.Form("advCrit") <> "" AndAlso CriteriaMsg = "") Then CriteriaMsg = Server.UrlDecode(Request.Form("advCrit"))
    If (Request.Form("advTokens") <> "" AndAlso CriteriaTokens = "") Then CriteriaTokens = Server.UrlDecode(Request.Form("advTokens"))
    
    AdvSearchSQL = WhereBuf.ToString
    sSearchQuery += AdvSearchSQL
  Else
    If (Request.QueryString("searchterms") <> "") Then
      idSearchText = MyCommon.Parse_Quotes(Request.QueryString("searchterms"))
      If (idSearchText <> "") Then
        PrctSignPos = idSearchText.IndexOf("%")
        If (PrctSignPos > -1) Then
          idSearch = -1
          idSearchText = idSearchText.Replace("%", "[%]")
        End If
        If (idSearchText.IndexOf("_") > -1) Then idSearchText = idSearchText.Replace("_", "[_]")
        sSearchQuery &= " and (LinkID3 like N'%" & idSearchText & "%'"
        sSearchQuery &= " or LinkID4 like N'%" & idSearchText & "%'"
        sSearchQuery &= " or LinkID5 like N'%" & idSearchText & "%'"
        sSearchQuery &= " or Description like '%" & idSearchText & "%')"
      End If
    End If
  End If
    
    
    
  sSearchQuery = sSearchQuery & " order by " & SortText & " " & SortDirection & ";"
  MyCommon.QueryStr = sSearchQuery
  dt = MyCommon.LRT_Select
  sizeOfData = dt.Rows.Count
  
  If (Request.QueryString("excel") <> "") Then
    InfoMessage = ExportListToExcel(dt, MyCommon, Logix)
    If InfoMessage = "" Then
      GoTo done
    End If
  End If

  i = linesPerPage * PageNum
%>
<div id="intro">
  <h1 id="title">
    <%
      Sendb(Copient.PhraseLib.Lookup("term.cashierhistoryreport", LanguageID))
    %>
  </h1>
  <div id="controls">
    <%
      If dt.Rows.Count > 0 Then
        Send_ExportToExcel()
      End If
    %>
  </div>
</div>
<div id="main">
  <%
    If (InfoMessage <> "") Then
      Send("<div id=""infobar"" class=""red-background"">" & InfoMessage & "</div>")
    End If
    Send_Listbar(linesPerPage, sizeOfData, PageNum, Request.QueryString("searchterms"), "&amp;SortText=" & SortText & "&amp;SortDirection=" & SortDirection, False)
    If (CriteriaMsg <> "") Then
      Send("<div id=""criteriabar"">" & CriteriaMsg & "<a href=""javascript:editSearchCriteria();"" class=""white"" style=""padding-left:15px;"">[" & Copient.PhraseLib.Lookup("term.edit", LanguageID) & "]</a><a href=""CM-cashier-report.aspx"" class=""white"" style=""padding-left:15px;"">[" & Copient.PhraseLib.Lookup("term.clear", LanguageID) & "]</a></div>")
    End If
  %>
  <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.history", LanguageID)) %>">
    <thead>
      <tr>
        <th align="left" scope="col" class="th-cashiertimedate">
          <a href="CM-cashier-report.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=ActivityDate&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.datetime", LanguageID))%>
          </a>
          <%
            If SortText = "ActivityDate" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" scope="col" class="th-cashier">
          <a href="CM-cashier-report.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=LinkID3&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.cashier", LanguageID))%>
          </a>
          <%
            If SortText = "LinkID3" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" scope="col" class="th-store">
          <a href="CM-cashier-report.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=LinkID4&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.store", LanguageID))%>
          </a>
          <%
            If SortText = "LinkID4" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" scope="col" class="th-cardid">
          <a href="CM-cashier-report.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=LinkID5&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.card", LanguageID) & " " & Copient.PhraseLib.Lookup("term.id", LanguageID))%>
          </a>
          <%
            If SortText = "LinkID5" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" scope="col" class="th-cardtype">
          <a href="CM-cashier-report.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=LinkID6&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID) & " " & Copient.PhraseLib.Lookup("term.type", LanguageID))%>
          </a>
          <%
            If SortText = "LinkID6" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" scope="col" class="th-cashieraction">
          <a href="CM-cashier-report.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=Description&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.action", LanguageID))%>
          </a>
          <%
            If SortText = "Description" Then
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
          If (Not IsDBNull(dt.Rows(i).Item("ActivityDate"))) Then
            Send("  <td>" & Format(dt.Rows(i).Item("ActivityDate"), "dd MMM yyyy HH:mm:ss") & "</td>")
          Else
            Send("  <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
          End If
          Send("  <td>" & MyCommon.NZ(dt.Rows(i).Item("LinkID3"), "") & "</td>")
          Send("  <td>" & MyCommon.NZ(dt.Rows(i).Item("LinkID4"), "") & "</td>")
          Send("  <td>" & MyCommon.NZ(dt.Rows(i).Item("LinkID5"), "") & "</td>")
          Send("  <td>" & MyCommon.NZ(dt.Rows(i).Item("LinkID6"), "") & "</td>")
          Send("  <td>" & MyCommon.NZ(dt.Rows(i).Item("Description"), "") & "</td>")
          Send("</tr>")
          If Shaded = "shaded" Then
            Shaded = ""
          Else
            Shaded = "shaded"
          End If
          i = i + 1
        End While
        If (sizeOfData = 0) Then
          Send("<tr>")
          Send("  <td colspan=""3""></td>")
          Send("</tr>")
        End If
      %>
    </tbody>
  </table>
</div>

<script type="text/javascript">  
 
  function handleExcel() {
    var sUrl = document.getElementById("ExcelUrl");
    var form = document.forms['excelform'];
    
    form.action = sUrl.value;
    form.method = "Post";
    form.submit();
  }
</script>

<script runat="server">
  Function GetOptionString(ByRef MyCommon As Copient.CommonInc, ByVal OptionIndex As Integer, _
                         ByVal OptionValue As String, ByVal FieldName As String) As String
    Dim FieldBuf As New StringBuilder()
    FieldBuf.Append(FieldName & " ")
    Select Case OptionIndex
      Case 1 ' contains
        FieldBuf.Append(" like '%" & MyCommon.Parse_Quotes(OptionValue.Trim) & "%' ")
      Case 2 ' exact
        FieldBuf.Append(" = '" & MyCommon.Parse_Quotes(OptionValue.Trim) & "' ")
      Case 3 ' starts with
        FieldBuf.Append(" like '" & MyCommon.Parse_Quotes(OptionValue.Trim) & "%' ")
      Case 4 ' ends with
        FieldBuf.Append(" like '%" & MyCommon.Parse_Quotes(OptionValue.Trim) & "' ")
      Case 5 ' excludes
        FieldBuf = New StringBuilder()
        FieldBuf.Append(" (" & FieldName & " not like '%" & MyCommon.Parse_Quotes(OptionValue.Trim) & "%' or " & FieldName & " is null) ")
      Case 6 ' is
        FieldBuf.Append(" = " & MyCommon.Parse_Quotes(OptionValue.Trim) & " ")
      Case 7 ' is not
        FieldBuf.Append(" <> " & MyCommon.Parse_Quotes(OptionValue.Trim) & " ")
      Case Else ' default to contains
        FieldBuf.Append(" like '%" & MyCommon.Parse_Quotes(OptionValue.Trim) & "%' ")
    End Select
    Return FieldBuf.ToString
  End Function
  
  Function GetOptionType(ByVal OptionIndex As Integer) As String
    Dim OptionType As String = "contains"
    Select Case OptionIndex
      Case 1 ' contains
        OptionType = "contains"
      Case 2 ' exact
        OptionType = "="
      Case 3 ' starts with
        OptionType = "starts with"
      Case 4 ' ends with
        OptionType = "ends with"
      Case 5 ' excludes
        OptionType = "excludes"
      Case 6 ' is
        OptionType = "is"
      Case 7 ' is not
        OptionType = "is not"
      Case Else ' default to contains
        OptionType = "contains"
    End Select
    Return OptionType
  End Function
  
  Function GetDateOption(ByRef MyCommon As Copient.CommonInc, ByVal OptionIndex As Integer, _
                         ByVal StartValue As String, ByVal EndValue As String, ByVal FieldName As String) As String
    Dim StartDate, EndDate As Date
    Dim FieldBuf As New StringBuilder()
    If (TryParseLocalizedDate(StartValue, StartDate, MyCommon) AndAlso OptionIndex <> 3) _
    OrElse (TryParseLocalizedDate(StartValue, StartDate, MyCommon) AndAlso TryParseLocalizedDate(EndValue, EndDate, MyCommon)) Then
      FieldBuf.Append(FieldName & " ")
      Select Case OptionIndex
        Case 0 ' on
          FieldBuf.Append(" between '" & StartDate.ToString("yyyy-MM-ddT00:00:00") & "' and '" & StartDate.ToString("yyyy-MM-ddT23:59:59") & "' ")
        Case 1 ' before
          FieldBuf.Append(" < '" & StartDate.ToString("yyyy-MM-dd") & "' ")
        Case 2 ' after
          FieldBuf.Append(" > '" & StartDate.ToString("yyyy-MM-dd") & "' ")
        Case 3 ' between
          FieldBuf.Append(" between '" & StartDate.ToString("yyyy-MM-dd") & "' and '" & EndDate.ToString("yyyy-MM-dd") & "' ")
        Case Else ' default to after
          FieldBuf.Append(" > '" & StartDate.ToString("yyyy-MM-dd") & "' ")
      End Select
    Else
      Throw New ApplicationException(Copient.PhraseLib.Lookup("term.invaliddateformat", LanguageID) & " (" & MyCommon.GetAdminUser.Culture.DateTimeFormat.ShortDatePattern & ").<br />")
    End If
    Return FieldBuf.ToString
  End Function
  
  Function GetDateOptionType(ByVal OptionIndex As Integer) As String
    Dim OptionType As String = "after"
    Select Case OptionIndex
      Case 0 ' on
        OptionType = Copient.PhraseLib.Lookup("term.on", LanguageID)
      Case 1 ' before
        OptionType = Copient.PhraseLib.Lookup("term.before", LanguageID)
      Case 2 ' after
        OptionType = Copient.PhraseLib.Lookup("term.after", LanguageID)
      Case 3 ' between
        OptionType = Copient.PhraseLib.Lookup("term.between", LanguageID)
      Case Else ' default to after
        OptionType = Copient.PhraseLib.Lookup("term.after", LanguageID)
    End Select
    Return OptionType
  End Function

  Private Function ExportListToExcel(ByRef dst As DataTable, ByRef MyCommon As Copient.CommonInc, ByRef Logix As Copient.LogixInc) As String
    Dim bStatus As Boolean
    Dim sMsg As String = ""
    Dim CmExport As New Copient.ExportXml
    Dim sFileFullPath As String
    Dim sFullPathFileName As String
    Dim sFileName As String = "CashierCustomerHistory.xls"
    Dim dtExport As DataTable
    Dim dr As DataRow
    Dim drExport As DataRow


    If dst.Rows.Count > 0 Then
      
      dtExport = New DataTable()
      dtExport.Columns.Add("Datetime", Type.GetType("System.String"))
      dtExport.Columns.Add("Cashier", Type.GetType("System.String"))
      dtExport.Columns.Add("Store", Type.GetType("System.String"))
      dtExport.Columns.Add("CardID", Type.GetType("System.String"))
      dtExport.Columns.Add("CardType", Type.GetType("System.String"))
      dtExport.Columns.Add("Action", Type.GetType("System.String"))
      
      For Each dr In dst.Rows
        drExport = dtExport.NewRow()
        drExport.Item("Datetime") = Format(dr.Item("ActivityDate"), "dd MMM yyyy HH:mm:ss")
        drExport.Item("Cashier") = dr.Item("LinkID3")
        drExport.Item("Store") = dr.Item("LinkID4")
        drExport.Item("CardID") = dr.Item("LinkID5")
        drExport.Item("CardType") = dr.Item("LinkID6")
        drExport.Item("Action") = dr.Item("Description")
        dtExport.Rows.Add(drExport)
      Next

      sFileFullPath = MyCommon.Fetch_SystemOption(29)
      sFullPathFileName = sFileFullPath & "\" & sFileName

      bStatus = CmExport.ExportToExcel(sFullPathFileName, dtExport)
      If bStatus Then
        Dim oRead As System.IO.StreamReader
        Dim LineIn As String
        Dim Bom As String = ChrW(65279)
        oRead = System.IO.File.OpenText(sFullPathFileName)
        Response.Clear()
        Response.ContentEncoding = Encoding.Unicode
        Response.ContentType = "application/octet-stream"
        Response.AddHeader("Content-Disposition", "attachment; filename=" & sFileName)
        
        'force little endian fffe bytes at front, why?  i dont know but is required.
        Sendb(Bom)
        While oRead.Peek <> -1
          LineIn = oRead.ReadLine()
          Send(LineIn)
        End While
        oRead.Close()
        Response.End()
        System.IO.File.Delete(sFullPathFileName)
      Else
        sMsg = CmExport.GetStatusMsg
      End If
    Else
      sMsg = Copient.PhraseLib.Lookup("offer-list.empty", LanguageID)
    End If
  
    Return sMsg
  End Function
  
  Function TryParseLocalizedDate(ByVal DateStr As String, ByRef LocalizedDate As Date, ByRef MyCommon As Copient.CommonInc) As Boolean
    Return Date.TryParseExact(DateStr, MyCommon.GetAdminUser.Culture.DateTimeFormat.ShortDatePattern, _
                           MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, LocalizedDate)
  End Function

</script>

<form id="frmIter" name="frmIter" method="post" action="">
  <input type="hidden" id="advSql" name="advSql" value="<% Sendb(Server.UrlEncode(AdvSearchSQL)) %>" />
  <input type="hidden" id="advCrit" name="advCrit" value="<% Sendb(Server.UrlEncode(CriteriaMsg)) %>" />
  <input type="hidden" id="advTokens" name="advTokens" value="<%Sendb(Server.UrlEncode(CriteriaTokens)) %>" />
</form>


<%
done:
  Send_BodyEnd()
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  MyCommon = Nothing
  Logix = Nothing
%>
