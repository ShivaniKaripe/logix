﻿<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: customer-storedfranking.aspx 
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
  Dim MyLookup As New Copient.CustomerLookup
  Dim dt As DataTable
  Dim dt2 As DataTable
  Dim dtSF As DataTable
  Dim row As DataRow
  Dim CustomerPK As Long = 0
  Dim CardPK As Long = 0
  Dim ExtCardID As String = ""
  Dim IsHouseholdID As Boolean = False
  Dim bIsErrorMsg As Boolean
  Dim sSearchQuery As String
  Dim PageNum As Integer
  Dim MorePages As Boolean
  Dim linesPerPage As Integer = 20
  Dim sizeOfData As Integer
  Dim i As Integer
  Dim SortText As String = "create_date"
  Dim SortDirection As String
  Dim Shaded As String = "shaded"
  Dim FirstName As String = ""
  Dim MiddleName As String = ""
  Dim LastName As String = ""
  Dim Note As String = ""
  Dim FullName As String = ""
  Dim InfoMessage As String = ""
  Dim Handheld As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "customer-storedfranking.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  PageNum = Request.QueryString("pagenum")
  If PageNum < 0 Then
    PageNum = 0
  End If
  MorePages = False
  
  CustomerPK = MyCommon.Extract_Val(Request.QueryString("CustPK"))
  If CustomerPK = 0 Then
    CustomerPK = MyCommon.Extract_Val(Request.Form("CustPK"))
  End If
  
  CardPK = MyCommon.Extract_Val(Request.QueryString("CardPK"))
  If CardPK = 0 Then
    CardPK = MyCommon.Extract_Val(Request.Form("CardPK"))
  End If
  If CardPK > 0 Then
    ExtCardID = MyLookup.FindExtCardIDFromCardPK(CardPK)
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
  
  MyCommon.QueryStr = "select CustomerPK, CustomerTypeID from Customers with (NoLock) where CustomerPK=" & CustomerPK & ";"
  dt = MyCommon.LXS_Select
  If (dt.Rows.Count > 0) Then
    IsHouseholdID = MyCommon.NZ(dt.Rows(0).Item("CustomerTypeID"), 0) = 1
  End If
  
  sSearchQuery = "select * from StoredFranking as SF with (NoLock) where CustomerPK=" & CustomerPK
  MyCommon.QueryStr = sSearchQuery & " order by " & SortText & " " & SortDirection
  dtSF = MyCommon.LXS_Select
  sizeOfData = dtSF.Rows.Count
  i = linesPerPage * PageNum
    
  Send_HeadBegin("term.customer", "term.storedfranking", CustomerPK)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
  
  Send("<script type=""text/javascript"">")
  Send("function ChangeParentDocument() { ")
  Send("  if (opener != null && !opener.closed) { ")
  Send("    if (opener.location.href.indexOf('CAM') > -1) { ")
  Send("      opener.location = '/logix/CAM/CAM-customer-offers.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&CardPK=" & CardPK, "") & "'; ")
  Send("    } else { ")
  Send("      opener.location = '/logix/customer-offers.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&CardPK=" & CardPK, "") & "'; ")
  Send("    }")
  Send("  }")
  Send("}")
  Send("function handleSubmit() { ")
  Send("  var retVal = true;")
  Send("  var elem = document.getElementById('custnote');")
  Send("   ")
  Send("    if (elem != null) { ")
  Send("      retVal = (elem.value.length <= 1000); ")
  Send("      if (!retVal) { alert('" & Copient.PhraseLib.Lookup("error.notelength", LanguageID) & "'); }")
  Send("    }")
  Send("  return retVal;")
  Send("}")
  Send("</script>")
  
  Send_HeadEnd()
  Send_BodyBegin(2)
%>
<form action="customer-storedfranking.aspx" method="post" id="mainform" name="mainform" onsubmit="return handleSubmit();">
  <div id="intro">
    <%
      Send("<input type=""hidden"" id=""CustPK"" name=""CustPK"" value=""" & CustomerPK & """ />")
      If CardPK > 0 Then
        Send("<input type=""hidden"" id=""CardPK"" name=""CardPK"" value=""" & CardPK & """ />")
      End If
      Send("<h1 id=""title"">")
      If CardPK = 0 Then
        If (IsHouseholdID) Then
          Sendb(Copient.PhraseLib.Lookup("term.household", LanguageID))
        Else
          Sendb(Copient.PhraseLib.Lookup("term.customer", LanguageID))
        End If
      Else
        If (IsHouseholdID) Then
          Sendb(Copient.PhraseLib.Lookup("term.householdcard", LanguageID) & " #" & ExtCardID)
        Else
          Sendb(Copient.PhraseLib.Lookup("term.customercard", LanguageID) & " #" & ExtCardID)
        End If
      End If
      MyCommon.QueryStr = "select Prefix, FirstName, MiddleName, LastName, Suffix from Customers with (NoLock) where CustomerPK=" & CustomerPK & ";"
      dt2 = MyCommon.LXS_Select
      If dt2.Rows.Count > 0 Then
        FullName = IIf(MyCommon.NZ(dt2.Rows(0).Item("Prefix"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_Prefix, MyCommon.NZ(dt2.Rows(0).Item("Prefix"), "") & " ", "")
        FullName &= IIf(MyCommon.NZ(dt2.Rows(0).Item("FirstName"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_FirstName, MyCommon.NZ(dt2.Rows(0).Item("FirstName"), "") & " ", "")
        FullName &= IIf(MyCommon.NZ(dt2.Rows(0).Item("MiddleName"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_MiddleName, Left(MyCommon.NZ(dt2.Rows(0).Item("MiddleName"), ""), 1) & ". ", "")
        FullName &= IIf(MyCommon.NZ(dt2.Rows(0).Item("LastName"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_LastName, MyCommon.NZ(dt2.Rows(0).Item("LastName"), ""), "")
        FullName &= IIf(MyCommon.NZ(dt2.Rows(0).Item("Suffix"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_Suffix, " " & MyCommon.NZ(dt2.Rows(0).Item("Suffix"), ""), "")
      End If
      If FullName <> "" Then
        Sendb(": " & MyCommon.TruncateString(FullName, 30))
      End If
      Send("</h1>")
    %>
    <div id="controls">
    </div>
  </div>
  <div id="main">
    <%
      If (InfoMessage <> "" And bIsErrorMsg) Then
        Send("<div id=""infobar"" class=""red-background"">" & InfoMessage & "</div>")
      ElseIf (InfoMessage <> "") Then
        Send("<div id=""infobar"" class=""green-background"">" & InfoMessage & "</div>")
      End If
    %>
    <div id="column">
      <div class="box" id="storedfranking">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.storedfranking", LanguageID))%>
          </span>
        </h2>
        <div style="border:solid 0px; height:278px; overflow:scroll; overflow-x:hidden;">
          <%
            If (dtSF.Rows.Count > 0) Then
              Send("<table summary=""" & Copient.PhraseLib.Lookup("term.storedfranking", LanguageID) & """>")
              Send("  <thead>")
              Send("    <tr>")
              Send("      <th class=""th-status"" scope=""col"">")
              Send("        <a href=""customer-storedfranking.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "&amp;SortText=status&amp;SortDirection=" & SortDirection & """>")
              Send("          " & Copient.PhraseLib.Lookup("term.status", LanguageID))
              Send("        </a>")
              If SortText = "status" Then
                If SortDirection = "ASC" Then
                  Sendb("<span class=""sortarrow"">&#9660;</span>")
                Else
                  Sendb("<span class=""sortarrow"">&#9650;</span>")
                End If
              Else
              End If
              Send("      </th>")
              Send("      <th class=""th-datetime"" scope=""col"">")
              Send("        <a href=""customer-storedfranking.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "&amp;SortText=create_date&amp;SortDirection=" & SortDirection & """>")
              Send("          " & Copient.PhraseLib.Lookup("term.created", LanguageID))
              Send("        </a>")
              If SortText = "create_date" Then
                If SortDirection = "ASC" Then
                  Sendb("<span class=""sortarrow"">&#9660;</span>")
                Else
                  Sendb("<span class=""sortarrow"">&#9650;</span>")
                End If
              Else
              End If
              Send("      </th>")
              Send("      <th style=""min-width:90px;"" scope=""col""></th>")
              Send("      <th class=""th-datetime"" scope=""col"">")
              Send("        <a href=""customer-storedfranking.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "&amp;SortText=issue_date&amp;SortDirection=" & SortDirection & """>")
              Send("          " & Copient.PhraseLib.Lookup("term.issued", LanguageID))
              Send("        </a>")
              If SortText = "issue_date" Then
                If SortDirection = "ASC" Then
                  Sendb("<span class=""sortarrow"">&#9660;</span>")
                Else
                  Sendb("<span class=""sortarrow"">&#9650;</span>")
                End If
              Else
              End If
              Send("      </th>")
              Send("      <th style=""min-width:90px;"" scope=""col""></th>")
              Send("      <th class=""th-id"" scope=""col"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & "</th>")
              'Send("      <th scope=""col"">" & Copient.PhraseLib.Lookup("term.text", LanguageID) & "</th>")
              Send("    </tr>")
              Send("  </thead>")
              Send("  <tbody>")
              For Each row In dtSF.Rows
                Send("    <tr class=""" & Shaded & """ >")
                'Status
                If MyCommon.NZ(row.Item("status"), 0) = 0 Then
                  Send("      <td>" & Copient.PhraseLib.Lookup("term.new", LanguageID) & "</td>")
                ElseIf MyCommon.NZ(row.Item("status"), 0) = 1 Then
                  Send("      <td>" & Copient.PhraseLib.Lookup("term.uploaded", LanguageID) & "</td>")
                ElseIf MyCommon.NZ(row.Item("status"), 0) = 2 Then
                  Send("      <td>" & Copient.PhraseLib.Lookup("term.delivered", LanguageID) & "</td>")
                Else
                  Send("      <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
                End If
                'Creation date
                If (Not IsDBNull(row.Item("create_date"))) Then
                  Send("      <td>" & Logix.ToShortDateTimeString(row.Item("create_date"), MyCommon) & "</td>")
                Else
                  Send("      <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
                End If
                'Creation store
                If (Not IsDBNull(row.Item("origin_store"))) Then
                  MyCommon.QueryStr = "select ExtLocationCode from Locations with (NoLock) where LocationID=" & MyCommon.NZ(row.Item("origin_store"), 0) & ";"
                  dt = MyCommon.LRT_Select
                  If dt.Rows.Count > 0 Then
                    Sendb("      <td>" & Copient.PhraseLib.Lookup("term.at", LanguageID) & " ")
                    If Logix.UserRoles.AccessStores Then
                      Send("<a href=""store-edit.aspx?LocationID=" & MyCommon.NZ(row.Item("origin_store"), 0) & """ target=""main"">" & MyCommon.NZ(dt.Rows(0).Item("ExtLocationCode"), "") & "</a></td")
                    Else
                      Send(MyCommon.NZ(dt.Rows(0).Item("ExtLocationCode"), "") & "</td")
                    End If
                  Else
                    Send("      <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td")
                  End If
                Else
                  Send("      <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td")
                End If
                'Issuing date
                If (Not IsDBNull(row.Item("issue_date"))) Then
                  Send("      <td>" & Logix.ToShortDateTimeString(row.Item("issue_date"), MyCommon) & "</td>")
                Else
                  Send("      <td>&nbsp;</td>")
                End If
                'Issuing store
                If (Not IsDBNull(row.Item("issuing_store"))) Then
                  MyCommon.QueryStr = "select ExtLocationCode from Locations with (NoLock) where LocationID=" & MyCommon.NZ(row.Item("issuing_store"), 0) & ";"
                  dt = MyCommon.LRT_Select
                  If dt.Rows.Count > 0 Then
                    Sendb("      <td>" & Copient.PhraseLib.Lookup("term.at", LanguageID) & " ")
                    If Logix.UserRoles.AccessStores Then
                      Send("<a href=""store-edit.aspx?LocationID=" & MyCommon.NZ(row.Item("issuing_store"), 0) & """ target=""main"">" & MyCommon.NZ(dt.Rows(0).Item("ExtLocationCode"), "") & "</a></td")
                    Else
                      Send(MyCommon.NZ(dt.Rows(0).Item("ExtLocationCode"), "") & "</td")
                    End If
                  Else
                    Send("      <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td")
                  End If
                Else
                  Send("      <td>&nbsp;</td")
                End If
                'Offer
                MyCommon.QueryStr = "select IncentiveID from CPE_RewardOptions with (NoLock) where RewardOptionID=" & MyCommon.NZ(row.Item("rewardid"), 0) & ";"
                dt = MyCommon.LRT_Select
                If dt.Rows.Count > 0 Then
                  If Logix.UserRoles.AccessOffers Then
                    Send("      <td><a href=""offer-redirect.aspx?OfferID=" & MyCommon.NZ(dt.Rows(0).Item("IncentiveID"), 0) & """ target=""main"">" & MyCommon.NZ(dt.Rows(0).Item("IncentiveID"), 0) & "</a></td")
                  Else
                    Send("      <td>" & MyCommon.NZ(dt.Rows(0).Item("IncentiveID"), "") & "</td")
                  End If
                Else
                  Send("      <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td")
                End If
                ''Text
                'Send("      <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("franking_text"), ""), 40) & "</td>")
                Send("    </tr>")
                Shaded = IIf(Shaded = "shaded", "", "shaded")
              Next
              Send("  </tbody>")
              Send("</table>")
            Else
              Send("<i>" & Copient.PhraseLib.Lookup("customer.nostoredfranking", LanguageID) & "</i>")
            End If
          %>
        </div>
        <hr class="hidden" />
      </div>
    </div>
  </div>
</form>
<%
done:
  MyCommon.Close_LogixXS()
  MyCommon.Close_LogixRT()
  Send_BodyEnd("mainform", "custnote")
  Logix = Nothing
  MyCommon = Nothing
%>
