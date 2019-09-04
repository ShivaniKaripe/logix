<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%
  ' *****************************************************************************
  ' * FILENAME: sv-adjust-redirect.aspx 
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
  Dim AdminUserID As Long
  Dim Logix As New Copient.LogixInc
  Dim dt As System.Data.DataTable = Nothing
  Dim row As System.Data.DataRow
  Dim CustomerPK As Long
  Dim CardPK As Long
  Dim OfferID As Long
  Dim Opener As String = ""
  Dim Programs As Integer = 0
  Dim InfoMessage As String = ""
  Dim Handheld As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "sv-adjust-redirect.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  OfferID = Request.QueryString("OfferID")
  CustomerPK = MyCommon.Extract_Val(Request.QueryString("CustomerPK"))
  CardPK = MyCommon.Extract_Val(Request.QueryString("CardPK"))
  Opener = Request.QueryString("Opener")
  If Opener = "" Then
    Opener = "customer-offers.aspx"
  End If
  
  MyCommon.QueryStr = "select OfR.RewardID as X, RSV.ProgramID, OfR.Deleted, SV.Name, SV.Value, SV.Description from OfferRewards as OfR " & _
                      "  inner join CM_RewardStoredValues as RSV on RSV.RewardStoredValuesID=OfR.LinkID " & _
                      "  inner join StoredValuePrograms as SV on SV.SVProgramID=RSV.ProgramID " & _
                      "  where OfferID=" & OfferID & " and RewardTypeID=10 and SV.Deleted=0 and OfR.Deleted=0 " & _
                      "  union " & _
                      "select OfC.ConditionID as X, OfC.LinkID as ProgramID, OfC.Deleted, SV.Name, SV.Value, SV.Description from OfferConditions as OfC " & _
                      "  inner join StoredValuePrograms as SV on SV.SVProgramID=OfC.LinkID " & _
                      "  where OfferID=" & OfferID & " and ConditionTypeID=6 and SV.Deleted=0 and OfC.Deleted=0 " & _
                      "  union " & _
                      "select RO.RewardOptionID as X, DSV.SVProgramID as ProgramID, DSV.Deleted, SV.Name, SV.Value, SV.Description from CPE_RewardOptions as RO " & _
                      "  inner join CPE_DeliverableStoredValue as DSV on DSV.RewardOptionID=RO.RewardOptionID " & _
                      "  inner join StoredValuePrograms as SV on SV.SVProgramID=DSV.SVProgramID " & _
                      "  where IncentiveID=" & OfferID & " and SV.Deleted=0 and DSV.Deleted=0 " & _
                      "  union " & _
                      "select RO.RewardOptionID as X, ISV.SVProgramID as ProgramID, ISV.Deleted, SV.Name, SV.Value, SV.Description from CPE_RewardOptions as RO " & _
                      "  inner join CPE_IncentiveStoredValuePrograms as ISV on ISV.RewardOptionID=RO.RewardOptionID " & _
                      "  inner join StoredValuePrograms as SV on SV.SVProgramID=ISV.SVProgramID " & _
                      "  where IncentiveID=" & OfferID & " and SV.Deleted=0 and ISV.Deleted=0 " & _
                      "  union " & _
                      "select RO.RewardOptionID, DI.SVProgramID, 0 as Deleted, SV.Name, SV.Value, SV.Description from CPE_Deliverables as DE with (NoLock) " & _
                      "  inner join CPE_RewardOptions as RO on RO.RewardOptionID=DE.RewardOptionID " & _
                      "  inner join CPE_Discounts as DI on DI.DiscountID=DE.OutputID " & _
                      "  inner join StoredValuePrograms as SV on SV.SVProgramID=DI.SVProgramID " & _
                      "  where RO.IncentiveID=" & OfferID & " and SV.Deleted=0 and DE.Deleted=0 and DE.DeliverableTypeID=2 and DI.AmountTypeID=7;"
  dt = MyCommon.LRT_Select
  Programs = dt.Rows.Count
  If Programs = 1 Then
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "sv-adjust-program.aspx?ProgramID=" & dt.Rows(0).Item("ProgramID") & "&CustomerPK=" & CustomerPK & _
                       IIf(CardPK > 0, "&CardPK=" & CardPK, "") & "&Opener=" & Opener & "&OfferID=" & OfferID)
    GoTo done
  End If
  Send_HeadBegin("term.storedvalue", "term.storedvalueadjustment", OfferID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
  Send_HeadEnd()
  Send_BodyBegin(2)
  
  If (Logix.UserRoles.AccessStoredValue = False) Then
    Send_Denied(2, "perm.customer-svaccess")
    GoTo done
  End If
%>
<div id="intro">
  <h1 id="title">
    <% Sendb(Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & ": " & Copient.PhraseLib.Lookup("term.storedvalueadjustment", LanguageID))%>
  </h1>
  <div id="controls">
  </div>
  <hr class="hidden" />
</div>
<div id="main">
  <% If (InfoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & InfoMessage & "</div>")%>
  <div id="column">
    <div class="box" id="programs">
      <h2>
        <span>
          <%Sendb(Copient.PhraseLib.Lookup("term.storedvalueprograms", LanguageID))%>
        </span>
      </h2>
      <%
        If dt.Rows.Count = 0 Then
          Send("<p>" & Copient.PhraseLib.Lookup("sv-adjust-redirect.NoSVPrograms", LanguageID) & "</p>")
        Else
          Send("<p>" & Copient.PhraseLib.Lookup("sv-adjust-redirect.MulitplePrograms", LanguageID) & "</p>")
          Send("<table summary=""" & Copient.PhraseLib.Lookup("term.programs", LanguageID) & """ width=""100%"">")
          Send("  <tr>")
          Send("    <th>" & Copient.PhraseLib.Lookup("term.id", LanguageID) & "</th>")
          Send("    <th>" & Copient.PhraseLib.Lookup("term.name", LanguageID) & "</th>")
          Send("    <th>" & Copient.PhraseLib.Lookup("term.unitvalue", LanguageID) & "</th>")
          Send("    <th>" & Copient.PhraseLib.Lookup("term.description", LanguageID) & "</th>")
          Send("  </tr>")
          For Each row In dt.Rows
            Send("<tr>")
            Send("  <td>")
            Send("    " & row.Item("ProgramID"))
            Send("  </td>")
            Send("  <td>")
            Send("    <a href=""sv-adjust-program.aspx?ProgramID=" & row.Item("ProgramID") & "&CustomerPK=" & CustomerPK & IIf(CardPK > 0, "&CardPK=" & CardPK, "") & "&Opener=customer-offers.aspx"">" & row.Item("Name") & "</a>")
            Send("  </td>")
            Send("  <td>")
            Send("    " & row.Item("Value"))
            Send("  </td>")
            Send("  <td>")
            Send("    " & row.Item("Description"))
            Send("  </td>")
            Send("</tr>")
          Next
          Send("</table>")
        End If
      %>
      <hr class="hidden" />
    </div>
  </div>
</div>
<%
done:
  Send_BodyEnd()
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  MyCommon = Nothing
  Logix = Nothing
%>
