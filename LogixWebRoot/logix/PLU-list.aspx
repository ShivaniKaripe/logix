<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.Commoninc" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%
  ' *****************************************************************************
  ' * FILENAME: PLU-list.aspx 
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
  Dim MultipleOffers As Boolean = False
  Dim Range As Decimal = 0
  Dim filterPLU As Integer = 0
  Dim MessageID As Integer = 0
  Dim OffsetBegin As Decimal = 0
  Dim OffsetEnd As Decimal = 0
  Dim IDLength As Integer = 0
  Dim Shaded As String = "shaded"
  Dim Expired As Boolean = False
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim BannersEnabled As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "PLU-list.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  
  PageNum = Request.QueryString("pagenum")
  If PageNum < 0 Then
    PageNum = 0
  End If
  MorePages = False
  
  Send_HeadBegin("term.triggercodes")
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
  
  
  Integer.TryParse(MyCommon.Fetch_SystemOption(52), IDLength)
  BannersEnabled = (MyCommon.Fetch_SystemOption("66") = "1")
  RangeBegin = CDec(MyCommon.Fetch_SystemOption(198))
  RangeBeginString = MyCommon.Fetch_SystemOption(198).ToString.PadLeft(IDLength, "0")
    RangeEnd = CDec(MyCommon.Fetch_SystemOption(199))
    RangeEndString = MyCommon.Fetch_SystemOption(199).ToString.PadLeft(IDLength, "0")
  RangeLocked = IIf(MyCommon.Fetch_CPE_SystemOption(95), False, True)
  MultipleOffers = MyCommon.Fetch_CPE_SystemOption(94)
  Range = (RangeEnd - RangeBegin) + 1
  If Request.QueryString("filterPLU") <> "" Then
    filterPLU = MyCommon.Extract_Val(Request.QueryString("filterPLU"))
  End If
  If Request.QueryString("OffsetBegin") <> "" Then
    OffsetBegin = MyCommon.Extract_Val(Request.QueryString("OffsetBegin"))
  End If
  If Request.QueryString("OffsetEnd") <> "" Then
    OffsetEnd = MyCommon.Extract_Val(Request.QueryString("OffsetEnd"))
  End If
  
  MyCommon.QueryStr = "select MessageID from CPE_CashierMessages with (NoLock) where PLU=1;"
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    MessageID = MyCommon.NZ(rst.Rows(0).Item("MessageID"), 0)
  End If
  
  If RangeBegin > RangeEnd Then
    infoMessage = Copient.PhraseLib.Lookup("plu.InvalidPLURange", LanguageID)
  End If
  
  Dim SortText As String = "PLU"
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
  
  If filterPLU = 0 Then
    ' Show all active PLUs
    sSearchQuery = "select distinct PLU from CPE_IncentivePLUs as CIP with (NoLock) " & _
                   "left join CPE_RewardOptions as RO on RO.RewardOptionID=CIP.RewardOptionID " & _
                   "left join CPE_Incentives as I on I.IncentiveID=RO.IncentiveID " & _
                   "where PLU Is Not NULL and RTrim(PLU) <> '' and I.Deleted=0 "
  ElseIf filterPLU = 1 Then
    ' Show all active in-range PLUs
    MyCommon.QueryStr = "select distinct PLU into #tmpPLUnum from CPE_IncentivePLUs as CIP with (NoLock) " & _
                   "left join CPE_RewardOptions as RO on RO.RewardOptionID=CIP.RewardOptionID " & _
                   "left join CPE_Incentives as I on I.IncentiveID=RO.IncentiveID " & _
                   "where PLU Is Not NULL and RTrim(PLU) <> '' and I.Deleted=0 and IsNumeric(PLU)=1 "
    MyCommon.LRT_Execute()
    sSearchQuery = "select PLU from #tmpPLUnum " & _
                   "where CAST(PLU as decimal(26,0))>=" & RangeBegin & " and CAST(PLU as decimal(26,0))<=" & RangeEnd & " "
  ElseIf filterPLU = 2 Then
    ' Show all active out-of-range PLUs
    MyCommon.QueryStr = "select distinct PLU into #tmpPLUnum from CPE_IncentivePLUs as CIP with (NoLock) " & _
                   "left join CPE_RewardOptions as RO on RO.RewardOptionID=CIP.RewardOptionID " & _
                   "left join CPE_Incentives as I on I.IncentiveID=RO.IncentiveID " & _
                   "where PLU Is Not NULL and RTrim(PLU) <> '' and I.Deleted=0 and IsNumeric(PLU)=1; " & _
                   "select distinct PLU into #tmpPLUstr from CPE_IncentivePLUs as CIP with (NoLock) " & _
                   "left join CPE_RewardOptions as RO on RO.RewardOptionID=CIP.RewardOptionID " & _
                   "left join CPE_Incentives as I on I.IncentiveID=RO.IncentiveID " & _
                   "where PLU Is Not NULL and RTrim(PLU) <> '' and I.Deleted=0 and IsNumeric(PLU)=0; "
    MyCommon.LRT_Execute()
    sSearchQuery = "select PLU from #tmpPLUnum " & _
                   "where CAST(PLU as decimal(26,0))>=0 and (CAST(PLU as decimal(26,0))<" & RangeBegin & " or CAST(PLU as decimal(26,0))>" & RangeEnd & ") " & _
                   "union select PLU from #tmpPLUstr "
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
    sSearchQuery = sSearchQuery & " and (PLU='" & idSearchText & "' or PLU='" & idSearchText.TrimStart("0") & "' or PLU='" & idSearchText.PadLeft(IDLength, "0") & "') "
  End If
  MyCommon.QueryStr = sSearchQuery & " order by " & SortText & " " & SortDirection & ";"
  dst = MyCommon.LRT_Select
  sizeOfData = dst.Rows.Count
  i = linesPerPage * PageNum
%>
<div id="intro">
  <h1 id="title">
    <%
      Sendb(Copient.PhraseLib.Lookup("term.triggercodes", LanguageID))
    %>
  </h1>
  <div id="controls">
    <form action="PLU-list.aspx" id="controlsform" name="controlsform">
      <%
        If (Logix.UserRoles.EditSystemConfiguration = True) Then
          Send("<button type=""button"" class=""regular"" onclick=""javascript:openPopup('PLU-cmsg.aspx');"">" & Copient.PhraseLib.Lookup("term.message", LanguageID) & "</button>")
        End If
      %>
    </form>
  </div>
</div>
<div id="main">
  <%
    If (infoMessage <> "") Then
      Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
    End If
    Send_Listbar(linesPerPage, sizeOfData, PageNum, Request.QueryString("searchterms"), "&SortText=" & SortText & "&SortDirection=" & SortDirection)
    Send("<div id=""summarybar"">")
    Sendb("  " & Copient.PhraseLib.Detokenize("adjustmentUPC-list.TotalCodeRange", LanguageID, (RangeEnd - RangeBegin) + 1, RangeBeginString, RangeEndString))
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
          <a href="PLU-list.aspx?searchterms=<%Sendb(Request.QueryString("searchterms")) %>&amp;SortText=PLU&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.code", LanguageID))%>
          </a>
          <%
            If SortText = "PLU" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" class="" scope="col">
          <% Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>
        </th>
      </tr>
    </thead>
    <tbody>
      <%
        Shaded = "shaded"
        While (i < sizeOfData And i < linesPerPage + linesPerPage * PageNum)
          Send("      <tr class=""" & Shaded & """>")
          Send("        <td>" & MyCommon.NZ(dst.Rows(i).Item("PLU"), "") & "</td>")
          Sendb("        <td>")
          'If MyCommon.NZ(dst.Rows(i).Item("PLU"), -1) >= 0 Then
          MyCommon.QueryStr = "select CIP.RewardOptionID, I.IncentiveID, I.IncentiveName, I.StartDate, I.EndDate, I.IsTemplate, I.EngineID " & _
                              "from CPE_IncentivePLUs as CIP with (NoLock) " & _
                              "inner join CPE_RewardOptions as RO on RO.RewardOptionID=CIP.RewardOptionID " & _
                              "inner join CPE_Incentives as I on I.IncentiveID=RO.IncentiveID " & _
                              "where PLU='" & dst.Rows(i).Item("PLU") & "' " & _
                              "order by I.IncentiveName;"
          rst = MyCommon.LRT_Select
          If rst.Rows.Count > 0 Then
            For Each row In rst.Rows
                            If MyCommon.NZ(row.Item("EndDate"), "1/1/2000") < Now.Date Then
                Expired = True
              Else
                Expired = False
              End If
              Sendb("<a href=""offer-redirect.aspx?OfferID=" & MyCommon.NZ(row.Item("IncentiveID"), 0) & """" & IIf(Expired, " class=""greylink""", "") & ">" & MyCommon.NZ(row.Item("IncentiveName"), 0) & "</a>")
              If Expired Then
                Sendb(" <small>(" & Copient.PhraseLib.Lookup("term.expired", LanguageID) & ")</small>")
              End If
              Send("<br />")
            Next
            Sendb("")
          Else
            Sendb(Copient.PhraseLib.Lookup("term.none", LanguageID))
          End If
          'End If
          Send("</td>")
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
