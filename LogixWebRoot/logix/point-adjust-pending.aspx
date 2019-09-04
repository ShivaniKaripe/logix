<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: point-adjust-pending.aspx 
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
  Dim dtPending As DataTable
  Dim dt As DataTable
  Dim row As DataRow
  Dim sortedRows() As DataRow
  
  Dim Handheld As Boolean
  Dim InfoMessage As String = ""
  Dim SortText As String = "PrimaryExtID"
  Dim SortDirection As String = ""
  Dim SearchText As String = ""
  Dim PageNum As Integer
  Dim MorePages As Boolean
  Dim linesPerPage As Integer = 20
  Dim sizeOfData As Integer
  Dim i As Integer
  Dim Shaded As String = "shaded"
  Dim OfferID, ProgramID, CreatorID, PKID As Long
  Dim DetailText As String = ""
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "point-adjust-pending.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  Send_HeadBegin("term.pending")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
%>
<script type="text/javascript">
  function hideShowDetails(elemName, imgName) {
    var elem = document.getElementById(elemName);
    var imgElem = document.getElementById(imgName);
    
    if (elem != null) {
      elem.style.display =  (elem.style.display != 'block') ? 'block' : 'none';
      if (imgElem != null) {
        if (elem.style.display == 'none') {
          imgElem.src = '/images/plus.png';
        } else {
          imgElem.src = '/images/minus.png';
        }
      }
    }
  }
  
  function handleCheckboxClick(action, pkid) {
    var elemApply = document.getElementById('apply' + pkid);
    var elemDelete = document.getElementById('delete' + pkid);
    
    if (action == 'apply') {
      if (elemApply.checked == true) {
        elemDelete.checked = false;
      }
    } else if (action == 'delete') {
      if (elemDelete.checked == true) {
        elemApply.checked = false;
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
  Send_Tabs(Logix, 3)
  Send_Subtabs(Logix, 32, 2, LanguageID, 0)
  
  If (Logix.UserRoles.AccessCustomerInquiry = False) Then
    Send_Denied(1, "perm.customer-access")
    GoTo done
  End If
  
  If Request.QueryString("Save") <> "" Then
    ' delete any marked as delete
    ' apply any marked to apply
  End If
  
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
  
  ' load up all the pending adjustments
  MyCommon.QueryStr = "select PKID, LogixTransNum, TransNum, TransDate, ExtLocationCode, CUST.PrimaryExtID, " & _
                      "  TerminalNum, PEND.CustomerPK, ProgramID, OfferID, AdjAmount, CreateDate, CreatedBy, " & _
                      "  CUST.CustomerPK, CUST.FirstName, CUST.MiddleName, CUST.LastName, '' as OfferName, '' as ProgramName, '' as Creator " & _
                      "from PointsAdj_Pending as PEND with (NoLock) " & _
                      "inner join Customers as CUST on CUST.CustomerPK = PEND.CustomerPK "
  SearchText = Request.QueryString("searchterms")
  If (SearchText <> "") Then
    MyCommon.QueryStr &= "where Cust.LastName like '%" & SearchText & "%' or PrimaryExtID = '" & MyCryptLib.SQL_StringEncrypt(SearchText) & "' " & _
                         "  or TransNum = '" & SearchText & "' "
  End If
  
  dtPending = MyCommon.LXS_Select
  
  ' populate the names for each ID
  For Each row In dtPending.Rows
    OfferID = MyCommon.NZ(row.Item("OfferID"), 0)
    If OfferID > 0 Then
      MyCommon.QueryStr = "select IncentiveName from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID
      dt = MyCommon.LRT_Select
      If dt.Rows.Count > 0 Then
        row.Item("OfferName") = MyCommon.NZ(dt.Rows(0).Item("IncentiveName"), "")
      End If
    End If
    ProgramID = MyCommon.NZ(row.Item("ProgramID"), 0)
    If ProgramID > 0 Then
      MyCommon.QueryStr = "select ProgramName from PointsPrograms with (NoLock) where ProgramID=" & ProgramID
      dt = MyCommon.LRT_Select
      If dt.Rows.Count > 0 Then
        row.Item("ProgramName") = MyCommon.NZ(dt.Rows(0).Item("ProgramName"), "")
      End If
    End If
    CreatorID = MyCommon.NZ(row.Item("CreatedBy"), 0)
    If CreatorID > 0 Then
      MyCommon.QueryStr = "select FirstName + ' ' + LastName as FullName from AdminUsers with (NoLock) where AdminUserID =" & CreatorID
      dt = MyCommon.LRT_Select
      If dt.Rows.Count > 0 Then
        row.Item("Creator") = MyCommon.NZ(dt.Rows(0).Item("FullName"), "")
      End If
    End If
  Next
  
  sortedRows = dtPending.Select("", SortText & " " & SortDirection)
  sizeOfData = sortedRows.Length
  i = linesPerPage * PageNum
%>
<form id="mainform" name="mainform" action="point-adjust-pending.aspx">
  <div id="intro">
    <h1 id="title">
      <%
        Sendb(Copient.PhraseLib.Lookup("term.pending", LanguageID) & " " & Copient.PhraseLib.Lookup("term.adjustments", LanguageID).ToLower)
      %>
    </h1>
    <div id="controls">
      <% Send_Save()%>
    </div>
  </div>
  <div id="main">
    <% If (InfoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & InfoMessage & "</div>")%>
    <br />
    <br class="half" />
    <% Send_Listbar(linesPerPage, sizeOfData, PageNum, Request.QueryString("searchterms"), "&amp;SortText=" & SortText & "&amp;SortDirection=" & SortDirection)%>
    <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.pending", LanguageID)) %>">
      <thead>
        <tr>
          <th align="center" style="width: 13px;" scope="col">
            &nbsp;</th>
          <th align="center" class="th-button" scope="col">
            <%Sendb(Copient.PhraseLib.Lookup("term.delete", LanguageID))%>
          </th>
          <th align="center" class="th-button" scope="col">
            <%Sendb(Copient.PhraseLib.Lookup("term.apply", LanguageID))%>
          </th>
          <th align="left" class=".th-cardholder" scope="col">
            <a href="point-adjust-pending.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=PrimaryExtID&amp;SortDirection=<% Sendb(SortDirection) %>">
              <% Sendb(Copient.PhraseLib.Lookup("term.cardnumber", LanguageID))%>
            </a>
            <%
              If SortText = "PrimaryExtID" Then
                If SortDirection = "ASC" Then
                  Sendb("<span class=""sortarrow"">&#9660;</span>")
                Else
                  Sendb("<span class=""sortarrow"">&#9650;</span>")
                End If
              End If
            %>
          </th>
          <th align="left" class="th-firstname" scope="col">
            <a href="point-adjust-pending.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=FirstName&amp;SortDirection=<% Sendb(SortDirection) %>">
              <% Sendb(Copient.PhraseLib.Lookup("term.firstname", LanguageID))%>
            </a>
            <%
              If SortText = "FirstName" Then
                If SortDirection = "ASC" Then
                  Sendb("<span class=""sortarrow"">&#9660;</span>")
                Else
                  Sendb("<span class=""sortarrow"">&#9650;</span>")
                End If
              End If
            %>
          </th>
          <th align="left" class="th-lastname" scope="col">
            <a href="point-adjust-pending.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=LastName&amp;SortDirection=<% Sendb(SortDirection) %>">
              <% Sendb(Copient.PhraseLib.Lookup("term.lastname", LanguageID))%>
            </a>
            <%
              If SortText = "LastName" Then
                If SortDirection = "ASC" Then
                  Sendb("<span class=""sortarrow"">&#9660;</span>")
                Else
                  Sendb("<span class=""sortarrow"">&#9650;</span>")
                End If
              End If
            %>
          </th>
          <th align="left" class="th-name" scope="col">
            <a href="point-adjust-pending.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=ProgramName&amp;SortDirection=<% Sendb(SortDirection) %>">
              <% Sendb(Copient.PhraseLib.Lookup("term.program", LanguageID))%>
            </a>
            <%
              If SortText = "ProgramName" Then
                If SortDirection = "ASC" Then
                  Sendb("<span class=""sortarrow"">&#9660;</span>")
                Else
                  Sendb("<span class=""sortarrow"">&#9650;</span>")
                End If
              End If
            %>
          </th>
          <th align="right" class="th-amount" scope="col">
            <a href="point-adjust-pending.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=AdjAmount&amp;SortDirection=<% Sendb(SortDirection) %>">
              <% Sendb(Copient.PhraseLib.Lookup("term.adjust", LanguageID))%>
            </a>
            <%
              If SortText = "AdjAmount" Then
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
          Shaded = "shaded"
          DetailText = Copient.PhraseLib.Lookup("point-adjust-pending.viewdetails", LanguageID)
          While (i < sizeOfData And i < linesPerPage + linesPerPage * PageNum)
            PKID = MyCommon.NZ(sortedRows(i).Item("PKID"), 0)
            Send("      <tr class=""" & Shaded & """>")
            Send("        <td align=""center""><a href=""javascript:hideShowDetails('tr" & PKID & "', 'img" & PKID & "');"" alt=""" & DetailText & """ title=""" & DetailText & """><img id=""img" & PKID & """ src=""/images/plus.png"" /></a></td>")
            Send("        <td align=""center""><input type=""checkbox"" name=""delete"" id=""delete" & PKID & """ value=""" & PKID & """ onclick=""handleCheckboxClick('delete', '" & PKID & "');"" /></td>")
            Send("        <td align=""center""><input type=""checkbox"" name=""apply"" id=""apply" & PKID & """ value=""" & PKID & """ onclick=""handleCheckboxClick('apply', '" & PKID & "');"" /></td>")
            Send("        <td><a href=""/logix/CAM/CAM-customer-transactions.aspx?CustPK=" & sortedRows(i).Item("CustomerPK") & """  target=""_blank"">" & MyCommon.SplitNonSpacedString(MyCryptLib.SQL_StringDecrypt(sortedRows(i).Item("PrimaryExtID").ToString()), 30) & "</a></td>")
            Send("        <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(sortedRows(i).Item("FirstName"), "&nbsp;"), 25) & "</td>")
            Send("        <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(sortedRows(i).Item("LastName"), "&nbsp;"), 25) & "</td>")
            Send("        <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(sortedRows(i).Item("ProgramName"), "&nbsp;"), 25) & "</td>")
            Send("        <td align=""right"">" & MyCommon.NZ(sortedRows(i).Item("AdjAmount"), 0) & "</td>")
            Send("      </tr>")
            ' write the transaction row
            Send("      <tr id=""tr" & PKID & """ style=""display:none;background-color:#e0e0e0;"">")
            Send("        <td></td>")
            Sendb("        <td colspan=""7""><b>" & Copient.PhraseLib.Lookup("term.transaction", LanguageID) & "</b>    <span style=""padding-left:25px;""><b>Date:</b> " & MyCommon.NZ(sortedRows(i).Item("TransDate"), "") & "</span>")
            Sendb("<span style=""padding-left:25px;""><b>" & Copient.PhraseLib.Lookup("term.number", LanguageID) & ":</b> " & MyCommon.NZ(sortedRows(i).Item("TransNum"), "") & "</span>")
            Sendb("<span style=""padding-left:25px;""><b>" & Copient.PhraseLib.Lookup("term.lane", LanguageID) & ":</b>   " & MyCommon.NZ(sortedRows(i).Item("TerminalNum"), "") & "</span>")
            Send("<span style=""padding-left:25px;""><b>" & Copient.PhraseLib.Lookup("term.store", LanguageID) & ":</b> " & MyCommon.NZ(sortedRows(i).Item("ExtLocationCode"), "") & "</span>")
            Sendb("<span style=""padding-left:25px;""><b>" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & ":</b> " & MyCommon.NZ(sortedRows(i).Item("OfferName"), "") & "</span>")
            Sendb("<br /><br class=""half"" /><b>" & Copient.PhraseLib.Lookup("term.adjustment", LanguageID) & "</b> ")
            Sendb("<span style=""padding-left:25px;""><b>" & Copient.PhraseLib.Lookup("term.createdby", LanguageID) & ":</b> " & MyCommon.NZ(sortedRows(i).Item("Creator"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)) & " ")
            Sendb("<span style=""padding-left:25px;""><b>" & Copient.PhraseLib.Lookup("term.date", LanguageID) & ":</b> " & MyCommon.NZ(sortedRows(i).Item("CreateDate"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)) & " ")
            Send("        </td>")
            Send("      </tr>")
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
  Send_BodyEnd("mainform")
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  MyCommon = Nothing
  Logix = Nothing
%>
