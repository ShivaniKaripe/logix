<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.Commoninc" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%
  ' *****************************************************************************
  ' * FILENAME: adjustmentUPC-list.aspx 
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
  Dim rst As System.Data.DataTable
  Dim row As DataRow
  Dim idNumber As Integer
  Dim idSearch As String
  Dim idSearchText As String
  Dim sSearchQuery As String
  Dim PageNum As Integer
  Dim MorePages As Boolean
  Dim linesPerPage As Integer = 20
  Dim sizeOfData As Integer
  Dim i As Integer
  Dim PrctSignPos As Integer
  Dim RangeBegin As Decimal = 0
  Dim RangeBeginString As String = ""
  Dim RangeEnd As Decimal = 0
  Dim RangeEndString As String = ""
  Dim RangeLocked As Boolean = True
  Dim Range As Decimal = 0
  Dim filterUPC As Integer = 0
  Dim OffsetBegin As Decimal = 0
  Dim OffsetEnd As Decimal = 0
  Dim IDLength As Integer = 0
  Dim Shaded As String = "shaded"
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim BannersEnabled As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "adjustmentUPC-list.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  PageNum = Request.QueryString("pagenum")
  If PageNum < 0 Then
    PageNum = 0
  End If
  MorePages = False
  
  
  Send_HeadBegin("term.adjustmentupcs")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
%>
<style type="text/css">
#summarybar {
  background-color: #ddddff;
  font-size: 11px;
  margin-top: 2px;
  padding: 3px;
  text-align: center;
  width: 733px;
}
* html #summarybar {
  width: 740px;
}
</style>
<%
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
  
  
    MyCommon.QueryStr = "select paddingLength from producttypes where producttypeid=1"
    rst = MyCommon.LRT_Select
    If rst IsNot Nothing Then
        IDLength = Convert.ToInt32(rst.Rows(0).Item("paddingLength"))
    End If
    rst = Nothing 
    'Integer.TryParse(MyCommon.Fetch_SystemOption(52), IDLength)
  BannersEnabled = (MyCommon.Fetch_SystemOption("66") = "1")
  If MyCommon.Fetch_CPE_SystemOption(100) = "" Then
    RangeBegin = 0
  Else
    RangeBegin = CDec(MyCommon.Fetch_CPE_SystemOption(100))
  End If
  RangeBeginString = MyCommon.Fetch_CPE_SystemOption(100).ToString.PadLeft(IDLength, "0")
  If MyCommon.Fetch_CPE_SystemOption(101) = "" Then
    RangeEnd = 0
  Else
    RangeEnd = CDec(MyCommon.Fetch_CPE_SystemOption(101))
  End If
  RangeEndString = MyCommon.Fetch_CPE_SystemOption(101).ToString.PadLeft(IDLength, "0")
  RangeLocked = IIf(MyCommon.Fetch_CPE_SystemOption(102), False, True)
  Range = (RangeEnd - RangeBegin) + 1
  If Request.QueryString("filterUPC") <> "" Then
    filterUPC = MyCommon.Extract_Val(Request.QueryString("filterUPC"))
  End If
  If Request.QueryString("OffsetBegin") <> "" Then
    OffsetBegin = MyCommon.Extract_Val(Request.QueryString("OffsetBegin"))
  End If
  If Request.QueryString("OffsetEnd") <> "" Then
    OffsetEnd = MyCommon.Extract_Val(Request.QueryString("OffsetEnd"))
  End If
  
  If RangeBegin > RangeEnd Then
    infoMessage = Copient.PhraseLib.Lookup("adjustmentUPC-list.IllogicalUPCRange", LanguageID) 'Illogical UPC range: beginning value is higher than the ending value.
  End If
  
  Dim SortText As String = "AdjustmentUPC"
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
  
  'Prior to querying, do a clean-up: eliminate any non-numeric AdjustmentUPCs from the PointsPrograms and StoredValuePrograms tables
  MyCommon.QueryStr = "update PointsPrograms with (RowLock) set AdjustmentUPC=Null where PatIndex('%[^0-9]%', AdjustmentUPC) > 0;"
  MyCommon.LRT_Execute()
  MyCommon.QueryStr = "update StoredValuePrograms with (RowLock) set AdjustmentUPC=Null where PatIndex('%[^0-9]%', AdjustmentUPC) > 0;"
  MyCommon.LRT_Execute()
  If MyCommon.Fetch_SystemOption(95) = 1 Then
  If filterUPC = 0 Then
    ' Show all active UPCs
    sSearchQuery = "select distinct CAST(AdjustmentUPC as decimal(26,0)) as AdjustmentUPC, AdjustmentUPC as AdjustmentUPCString, ProgramType, ProgramID, ProgramName from (" & _
                   " select distinct AdjustmentUPC, 0 as ProgramType, ProgramID, ProgramName from PointsPrograms where AdjustmentUPC Is Not NULL and RTrim(AdjustmentUPC)<>'' and deleted = 0" & _
                   " union " & _
                   " select distinct AdjustmentUPC, 1 as ProgramType, SVProgramID as ProgramID, Name as ProgramName from StoredValuePrograms where AdjustmentUPC Is Not NULL and RTrim(AdjustmentUPC)<>''" & _
                   ") as AdjustmentUPC where AdjustmentUPC<>'' and CAST(AdjustmentUPC as decimal(26,0))>0 "
  ElseIf filterUPC = 1 Then
    ' Show all active in-range UPCs
    sSearchQuery = "select distinct CAST(AdjustmentUPC as decimal(26,0)) as AdjustmentUPC, AdjustmentUPC as AdjustmentUPCString, ProgramType, ProgramID, ProgramName from (" & _
                   " select distinct AdjustmentUPC, 0 as ProgramType, ProgramID, ProgramName from PointsPrograms where AdjustmentUPC Is Not NULL and RTrim(AdjustmentUPC)<>'' and deleted = 0" & _
                   " union " & _
                   " select distinct AdjustmentUPC, 1 as ProgramType, SVProgramID as ProgramID, Name as ProgramName from StoredValuePrograms where AdjustmentUPC Is Not NULL and RTrim(AdjustmentUPC)<>''" & _
                   ") as AdjustmentUPC where AdjustmentUPC<>'' and CAST(AdjustmentUPC as decimal(26,0))>0 " & _
                   "and CAST(AdjustmentUPC as decimal(26,0))>=" & RangeBegin & " and CAST(AdjustmentUPC as decimal(26,0))<=" & RangeEnd & " "
  ElseIf filterUPC = 2 Then
    ' Show all active out-of-range UPCs
    sSearchQuery = "select distinct CAST(AdjustmentUPC as decimal(26,0)) as AdjustmentUPC, AdjustmentUPC as AdjustmentUPCString, ProgramType, ProgramID, ProgramName from (" & _
                   " select distinct AdjustmentUPC, 0 as ProgramType, ProgramID, ProgramName from PointsPrograms where AdjustmentUPC Is Not NULL and RTrim(AdjustmentUPC)<>'' and deleted = 0" & _
                   " union " & _
                   " select distinct AdjustmentUPC, 1 as ProgramType, SVProgramID as ProgramID, Name as ProgramName from StoredValuePrograms where AdjustmentUPC Is Not NULL and RTrim(AdjustmentUPC)<>''" & _
                   ") as AdjustmentUPC where AdjustmentUPC<>'' and CAST(AdjustmentUPC as decimal(26,0))>0 " & _
                   "and (CAST(AdjustmentUPC as decimal(26,0))<" & RangeBegin & " or CAST(AdjustmentUPC as decimal(26,0))>" & RangeEnd & ") "
  End If
  idSearchText = Request.QueryString("searchterms")
  If (idSearchText <> "") Then
    If (Integer.TryParse(Request.QueryString("searchterms"), idNumber)) Then
      idSearch = idNumber.ToString
    Else
      idSearch = "-1"
    End If
    PrctSignPos = idSearchText.IndexOf("%")
    If (PrctSignPos > -1) Then
      idSearch = "-1"
      idSearchText = idSearchText.Replace("%", "[%]")
    End If
    If (idSearchText.IndexOf("_") > -1) Then idSearchText = idSearchText.Replace("_", "[_]")
    idSearchText = MyCommon.Parse_Quotes(idSearchText)
    If IsNumeric(idSearchText) Then
      sSearchQuery = sSearchQuery & " and ((CAST(AdjustmentUPC as decimal(26,0))='" & idSearchText & "' or CAST(AdjustmentUPC as decimal(26,0))='" & idSearchText.TrimStart("0") & "' or CAST(AdjustmentUPC as decimal(26,0))='" & idSearchText.PadLeft(IDLength, "0") & "') or (ProgramName like N'%" & idSearchText & "%')) "
    Else
      sSearchQuery = sSearchQuery & " and ProgramName like N'%" & idSearchText & "%' "
    End If
  End If
  MyCommon.QueryStr = sSearchQuery & " order by " & SortText & " " & SortDirection & ";"
  dst = MyCommon.LRT_Select
  sizeOfData = dst.Rows.Count
  i = linesPerPage * PageNum
  End If 
%>
<div id="intro">
  <h1 id="title">
    <%
      Sendb(Copient.PhraseLib.Lookup("term.adjustmentupcs", LanguageID))
    %>
  </h1>
  <div id="controls">
  </div>
</div>
<div id="main">
  <%
    If (infoMessage <> "") Then
      Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
    End If
    Send_Listbar(linesPerPage, sizeOfData, PageNum, Request.QueryString("searchterms"), "&SortText=" & SortText & "&SortDirection=" & SortDirection)
    Send("<div id=""summarybar"">")
    If MyCommon.Fetch_CPE_SystemOption(100) = "" OrElse MyCommon.Fetch_CPE_SystemOption(101) = "" Then
      Sendb(Copient.PhraseLib.Lookup("adjustmentUPC-list.RangeUndefined", LanguageID))
    Else
      Sendb("  " & Copient.PhraseLib.Detokenize("adjustmentUPC-list.TotalCodeRange", LanguageID, (RangeEnd - RangeBegin) + 1, RangeBeginString, RangeEndString)) 'Total code range: {0} ({1} to {2}).
    End If
    If Not RangeLocked Then
      Sendb("  " & Copient.PhraseLib.Lookup("adjustmentUPC-list.OutOfRangeCodesPermitted", LanguageID))
    End If
    Send("  <br />")
    Send("</div>")
  %>
  <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.triggercodes", LanguageID)) %>">
    <thead>
      <tr>
        <th align="left" style="width:200px;" scope="col">
          <a href="adjustmentUPC-list.aspx?searchterms=<%Sendb(Request.QueryString("searchterms")) %>&amp;SortText=AdjustmentUPC&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.upc", LanguageID))%>
          </a>
          <%
            If SortText = "AdjustmentUPC" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" style="width:90px;" scope="col">
          <a href="adjustmentUPC-list.aspx?searchterms=<%Sendb(Request.QueryString("searchterms")) %>&amp;SortText=ProgramType&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.type", LanguageID))%>
          </a>
          <%
            If SortText = "ProgramType" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" scope="col">
          <a href="adjustmentUPC-list.aspx?searchterms=<%Sendb(Request.QueryString("searchterms")) %>&amp;SortText=ProgramName&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.program", LanguageID))%>
          </a>
          <%
            If SortText = "ProgramName" Then
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
        Shaded = "shaded"
        While (i < sizeOfData And i < linesPerPage + linesPerPage * PageNum)
          Send("      <tr class=""" & Shaded & """>")
          Send("        <td>" & MyCommon.NZ(dst.Rows(i).Item("AdjustmentUPCString"), "") & "</td>")
          Send("        <td>" & IIf(MyCommon.NZ(dst.Rows(i).Item("ProgramType"), 0) = 0, "<a href=""point-list.aspx"">" & Copient.PhraseLib.Lookup("term.points", LanguageID) & "</a>", "<a href=""sv-list.aspx"">" & Copient.PhraseLib.Lookup("term.storedvalue", LanguageID) & "</a>") & "</td>")
          If MyCommon.NZ(dst.Rows(i).Item("ProgramType"), 0) = 0 Then
            Send("        <td><a href=""point-edit.aspx?ProgramGroupID=" & MyCommon.NZ(dst.Rows(i).Item("ProgramID"), "") & """>" & MyCommon.NZ(dst.Rows(i).Item("ProgramName"), "") & "</a></td>")
          ElseIf MyCommon.NZ(dst.Rows(i).Item("ProgramType"), 0) = 1 Then
            Send("        <td><a href=""sv-edit.aspx?ProgramGroupID=" & MyCommon.NZ(dst.Rows(i).Item("ProgramID"), "") & """>" & MyCommon.NZ(dst.Rows(i).Item("ProgramName"), "") & "</a></td>")
          End If
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
<%
done:
  MyCommon.Close_LogixRT()
  Logix = Nothing
  MyCommon = Nothing
  Send_BodyEnd("searchform", "searchterms")
%>
