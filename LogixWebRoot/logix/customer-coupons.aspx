<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.Commoninc" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%
  ' *****************************************************************************
  ' * FILENAME: customer-coupons.aspx 
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
  Dim BarcodeDT As System.Data.DataTable
  Dim CustomerDT As System.Data.DataTable
  Dim dt As System.Data.DataTable
  Dim CustomerPK As Long = 0
  Dim CardPK As Long = 0
  Dim ExtCardID As String = ""
  Dim idNumber As Integer
  Dim sSearchQuery As String = ""
  Dim PageNum As Integer
  Dim MorePages As Boolean
  Dim linesPerPage As Integer = 20
  Dim sizeOfData As Integer
  Dim i As Integer
  Dim PrctSignPos As Integer
  Dim Shaded As String = "shaded"
  Dim SortText As String = "IssueDate"
  Dim SortDirection As String = "ASC"
  Dim FirstName As String = ""
  Dim MiddleName As String = ""
  Dim LastName As String = ""
  Dim FullName As String = ""
  Dim VoidBarcode As String = ""
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim ActivityText As String = ""
  Dim AdminID As System.Data.DataTable
  Dim ROID as Integer
 
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "customer-coupons.aspx"
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
  
  MyCommon.QueryStr = "select CustomerPK, CustomerTypeID from Customers with (NoLock) where CustomerPK=" & CustomerPK & ";"
  dt = MyCommon.LXS_Select
  Dim IsHouseholdID As Boolean = False
  If (dt.Rows.Count > 0) Then
    IsHouseholdID = MyCommon.NZ(dt.Rows(0).Item("CustomerTypeID"), 0) = 1
  End If
  
  If (Request.QueryString("VoidBarcode") <> "") Then
    VoidBarcode = Request.QueryString("VoidBarcode")
    MyCommon.QueryStr = "select Barcode from BarcodeDetails " & _
                        "where CustomerPK=" & CustomerPK & " and Barcode='" & VoidBarcode & "' and ISNULL(Voided, 0)=0;"
    dt = MyCommon.LXS_Select
    If (dt.Rows.Count > 0) Then
      MyCommon.QueryStr = "update BarcodeDetails set Voided=1, RedeemedDate=getdate(), RedeemedLocationID='-9', RedeemedCSR=" & AdminUserID & " " & _
                          "where CustomerPK=" & CustomerPK & " and Barcode='" & VoidBarcode & "';"
      MyCommon.LXS_Execute()
      MyCommon.Activity_Log2(25, 24, CustomerPK, AdminUserID, Copient.PhraseLib.Lookup("history.customer-void", LanguageID) & " " & VoidBarcode)
    End If
    Response.Redirect("/logix/customer-coupons.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&CardPK=" & CardPK, ""))
  End If
  
  Try
    Send_HeadBegin("term.customer", "term.uniquecoupons", CustomerPK)
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
  %>
<style type="text/css">
#searcher {
  margin: 0 3px 0 0;
  width: 728px;
  }
* html #searcher {
  width: 736px;
  }
#paginator {
  margin: 0 3px 0 0;
  width: 916px;
  }
* html #paginator {
  width: 915px;
  }
  <!--471,477,659,664-->
</style>
  <%
    Send_Scripts(New String() {"datePicker.js"})
    
    Send("<script type=""text/javascript"">")
    Send("var datePickerDivID = ""datepicker"";")
    Send("")
    Send_Calendar_Overrides(MyCommon)
    Send("")
    Send("function ChangeParentDocument() { ")
    Send("  if (opener != null && !opener.closed) { ")
    Send("    if (opener.location.href.indexOf('CAM') > -1) { ")
    Send("      opener.location = '/logix/CAM/CAM-customer-general.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&CardPK=" & CardPK, "") & "'; ")
    Send("    } else { ")
    Send("      opener.location = '/logix/customer-general.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&CardPK=" & CardPK, "") & "'; ")
    Send("    }")
    Send("  }")
    Send("}")
    Send("function voidBarcode(Barcode) {")
    Send("  if (confirm('" & Copient.PhraseLib.Lookup("customer-coupons.VoidBarcode", LanguageID) & "')) {")
    Send("    document.getElementById('VoidBarcode').value = Barcode;")
    Send("    document.mainform.submit();")
    Send("  } else {")
    Send("    return false;")
    Send("  }")
    Send("}")
    Send("</script>")
    
    Send_HeadEnd()
    Send_BodyBegin(2)
    
    If (Logix.UserRoles.AccessCustomerCoupons = False) Then
      Send_Denied(1, "perm.customers-accesscoupons")
      GoTo done
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
        SortDirection = "ASC"
      End If
    Else
      SortDirection = Request.QueryString("SortDirection")
    End If
    
    sSearchQuery = "select Barcode, ValidLocation, CustomerPK, SVProgramID, Channel, GeneratedOn, IsNull(ExpirationDate, '1/1/2100') as ExpirationDate, " & _
                   "RedeemedLocationID, RedeemedDate, RedeemedCSR, Voided, RedeemingTransactionID, RedeemingMemberID, " & _
                   "IssuingTransactionID, IssuingCSR, IssuingCostCenter, IssueDate, RejectionCode, RejectedDate, RewardOptionID " & _
                   "from BarcodeDetails as BD with (NoLock) " & _
                   "where CustomerPK=" & CustomerPK
    'SEARCHTERMS
    Dim idSearch As String
    Dim idSearchText As String = ""
    If (Request.QueryString("searchterms") <> "") Then
      idSearchText = Request.QueryString("searchterms")
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
      sSearchQuery &= " and (Barcode='" & idSearchText & "')"
    End If
    'FILTER
    Dim filterCoupon As Integer = 0
    If (Request.QueryString("filtercoupon") <> "") Then
      filterCoupon = Request.QueryString("filtercoupon")
      If filterCoupon = 0 Then
        'All coupons
      ElseIf filterCoupon = 1 Then
        'Unredeemed nonexpired
        sSearchQuery &= " and RedeemedDate is NULL and IsNull(ExpirationDate, '1/1/2100') >= getdate() and IsNull(Voided, 0)=0"
      ElseIf filterCoupon = 2 Then
        'Unredeemed expired
        sSearchQuery &= " and RedeemedDate is NULL and ExpirationDate is not NULL and ExpirationDate<getdate() and IsNull(Voided, 0)=0"
      ElseIf filterCoupon = 3 Then
        'Voided
        sSearchQuery &= " and Voided=1"
      End If
    End If
    'DATES
    Dim startDate As String = ""
    Dim endDate As String = ""
    Dim TempDate As Date
    
    If (Request.QueryString("startDate") <> "") Then
      startDate = Request.QueryString("startDate")
      If Date.TryParse(startDate, MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, TempDate) Then
        sSearchQuery &= " and IssueDate>='" & Copient.commonShared.ConvertToSqlDate(startDate, MyCommon.GetAdminUser.Culture) & "'"
      Else
        infoMessage = Copient.PhraseLib.Lookup("term.InvalidStartDate", LanguageID)
      End If
    End If
    If (Request.QueryString("endDate") <> "") Then
      endDate = Request.QueryString("endDate")
      If Date.TryParse(endDate, MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, TempDate) Then
        sSearchQuery &= " and ExpirationDate<='" & TempDate.ToString("yyyy-MM-dd") & " 23:59:59'"
      Else
        infoMessage = Copient.PhraseLib.Lookup("term.InvalidEndDate", LanguageID)
      End If
    End If
    
    MyCommon.QueryStr = sSearchQuery '& " order by " & SortText & " " & SortDirection
    BarcodeDT = MyCommon.LXS_Select
	BarcodeDT.Columns.Add("RewardName",GetType(String))
	If BarcodeDT.Rows.Count >0 Then
		For Each row As DataRow In BarcodeDT.Rows
			'ROID = BarcodeDT.Rows(i).Item("RewardOptionID") 
			If IsDBNull(row("RewardOptionID")) Then
				row("RewardName") = Copient.PhraseLib.Lookup("barcode.rewardnotfound",LanguageID)
			Else
				' ROID = row("RewardOptionID")
				' MyCommon.QueryStr = "select Value as 'RewardName' from PassThruTierValues as D "& _
													' "INNER JOIN CPE_Deliverables AS P " & _
													' "on D.PTPKID = P.OutputID " & _
													' "where P.RewardOptionID = " & ROID & " AND D.PassThruPresTagID = 11 AND P.DeliverableTypeID = 12"
				' dt = MyCommon.LRT_Select
				' row("RewardName") = dt.Rows(0).Item("RewardName")
				MyCommon.QueryStr = "dbo.pt_GetCouponRewardName"
				MyCommon.Open_LRTsp()
				MyCommon.LRTsp.Parameters.Add("@ROID", System.Data.SqlDbType.Bigint).Value = row("RewardOptionID")
				MyCommon.LRTsp.Parameters.Add("@RewardName", System.Data.SqlDbType.Nvarchar, 255).Direction = System.Data.ParameterDirection.Output
				MyCommon.LRTsp.ExecuteNonQuery()
				row("RewardName") =  MyCommon.LRTsp.Parameters("@RewardName").Value
				If IsDBNull(row("RewardName"))  then row("RewardName") = Copient.PhraseLib.Lookup("barcode.rewardnotfound",LanguageID)
				
			End If

		Next
		BarcodeDT.DefaultView.Sort = SortText & " " & SortDirection
		BarcodeDT = BarcodeDT.DefaultView.ToTable()
	End If
	
		
    sizeOfData = BarcodeDT.Rows.Count
    i = linesPerPage * PageNum
%>
<form action="#"  id="mainform" name="mainform">
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
      CustomerDT = MyCommon.LXS_Select
      If CustomerDT.Rows.Count > 0 Then
        FullName = IIf(MyCommon.NZ(CustomerDT.Rows(0).Item("Prefix"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_Prefix, MyCommon.NZ(CustomerDT.Rows(0).Item("Prefix"), "") & " ", "")
        FullName &= IIf(MyCommon.NZ(CustomerDT.Rows(0).Item("FirstName"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_FirstName, MyCommon.NZ(CustomerDT.Rows(0).Item("FirstName"), "") & " ", "")
        FullName &= IIf(MyCommon.NZ(CustomerDT.Rows(0).Item("MiddleName"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_MiddleName, Left(MyCommon.NZ(CustomerDT.Rows(0).Item("MiddleName"), ""), 1) & ". ", "")
        FullName &= IIf(MyCommon.NZ(CustomerDT.Rows(0).Item("LastName"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_LastName, MyCommon.NZ(CustomerDT.Rows(0).Item("LastName"), ""), "")
        FullName &= IIf(MyCommon.NZ(CustomerDT.Rows(0).Item("Suffix"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_Suffix, " " & MyCommon.NZ(CustomerDT.Rows(0).Item("Suffix"), ""), "")
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
      Send("<br class=""half"" />")
      If (infoMessage <> "") Then
        Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
      End If
      Dim QueryString As String = ""
      QueryString = "CustPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "")
      Send_ListbarBarcodes(linesPerPage, sizeOfData, PageNum, filterCoupon, QueryString, Request.QueryString("searchterms"), "&amp;SortText=" & SortText & "&amp;SortDirection=" & SortDirection, startDate, endDate)
      Send("<input type=""hidden"" id=""VoidBarcode"" name=""VoidBarcode"" value="""" />")
    %>
    <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.uniquecoupons", LanguageID)) %>" style="width:915px;">
      <thead>
        <tr>
          <th align="center" class="th-button" scope="col" style="text-align: center;">
            <a href="customer-coupons.aspx?CustPK=<%sendb(CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, ""))%>&amp;searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=Voided&amp;SortDirection=<% Sendb(SortDirection) %>">
              <% Sendb(Copient.PhraseLib.Lookup("term.void", LanguageID))%>
            </a>
            <%
              If SortText = "Voided" Then
                If SortDirection = "ASC" Then
                  Sendb("<span class=""sortarrow"">&#9660;</span>")
                Else
                  Sendb("<span class=""sortarrow"">&#9650;</span>")
                End If
              End If
            %>
          </th>
          <th align="left" class="th-code" scope="col">
            <a href="customer-coupons.aspx?CustPK=<%sendb(CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, ""))%>&amp;searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=Barcode&amp;SortDirection=<% Sendb(SortDirection) %>">
              <% Sendb(Copient.PhraseLib.Lookup("term.barcode", LanguageID))%>
            </a>
            <%
              If SortText = "Barcode" Then
                If SortDirection = "ASC" Then
                  Sendb("<span class=""sortarrow"">&#9660;</span>")
                Else
                  Sendb("<span class=""sortarrow"">&#9650;</span>")
                End If
              End If
            %>
          </th>
		  <th align="left" class="th-datetime" scope="col">
            <a href="customer-coupons.aspx?CustPK=<%sendb(CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, ""))%>&amp;searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=RewardName&amp;SortDirection=<% Sendb(SortDirection) %>">
              <% Sendb(Copient.PhraseLib.Lookup("term.rewardname", LanguageID))%>
            </a>
            <%
              If SortText = "RewardName" Then
                If SortDirection = "ASC" Then
                  Sendb("<span class=""sortarrow"">&#9660;</span>")
                Else
                  Sendb("<span class=""sortarrow"">&#9650;</span>")
                End If
              End If
            %>
          </th>
		  <th align="left" class="th-datetime" scope="col">
            <a href="customer-coupons.aspx?CustPK=<%sendb(CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, ""))%>&amp;searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=IssuingCostCenter&amp;SortDirection=<% Sendb(SortDirection) %>">
              <% Sendb(Copient.PhraseLib.Lookup("term.issuingcostcenter", LanguageID))%>
            </a>
            <%
              If SortText = "IssuingCostCenter" Then
                If SortDirection = "ASC" Then
                  Sendb("<span class=""sortarrow"">&#9660;</span>")
                Else
                  Sendb("<span class=""sortarrow"">&#9650;</span>")
                End If
              End If
            %>
          </th>
          <th align="left" class="th-datetime" scope="col">
            <a href="customer-coupons.aspx?CustPK=<%sendb(CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, ""))%>&amp;searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=IssueDate&amp;SortDirection=<% Sendb(SortDirection) %>">
              <% Sendb(Copient.PhraseLib.Lookup("term.issuedon", LanguageID))%>
            </a>
            <%
              If SortText = "IssueDate" Then
                If SortDirection = "ASC" Then
                  Sendb("<span class=""sortarrow"">&#9660;</span>")
                Else
                  Sendb("<span class=""sortarrow"">&#9650;</span>")
                End If
              End If
            %>
          </th>
          <th align="left" class="th-datetime" scope="col">
            <a href="customer-coupons.aspx?CustPK=<%sendb(CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, ""))%>&amp;searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=ExpirationDate&amp;SortDirection=<% Sendb(SortDirection) %>">
              <% Sendb(Copient.PhraseLib.Lookup("term.expires", LanguageID))%>
            </a>
            <%
              If SortText = "ExpirationDate" Then
                If SortDirection = "ASC" Then
                  Sendb("<span class=""sortarrow"">&#9660;</span>")
                Else
                  Sendb("<span class=""sortarrow"">&#9650;</span>")
                End If
              End If
            %>
          </th>
		  <th align="left" class="th-status" scope="col">
            <a href="customer-coupons.aspx?CustPK=<%sendb(CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, ""))%>&amp;searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=RedeemedLocationID&amp;SortDirection=<% Sendb(SortDirection) %>">
             <% Sendb(Copient.PhraseLib.Lookup("term.redmloc", LanguageID))%>
            </a>
            <%
              If SortText = "RedeemedLocationID" Then
                If SortDirection = "ASC" Then
                  Sendb("<span class=""sortarrow"">&#9660;</span>")
                Else
                  Sendb("<span class=""sortarrow"">&#9650;</span>")
                End If
              End If
            %>
          </th>
          <th align="left" class="th-datetime" scope="col">
            <a href="customer-coupons.aspx?CustPK=<%sendb(CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, ""))%>&amp;searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=RedeemedDate&amp;SortDirection=<% Sendb(SortDirection) %>">
              <% Sendb(Copient.PhraseLib.Lookup("term.redeemedon", LanguageID))%>
            </a>
            <%
              If SortText = "RedeemedDate" Then
                If SortDirection = "ASC" Then
                  Sendb("<span class=""sortarrow"">&#9660;</span>")
                Else
                  Sendb("<span class=""sortarrow"">&#9650;</span>")
                End If
              End If
            %>
          </th>
          <th align="left" class="th-csr" scope="col">
            <a href="customer-coupons.aspx?CustPK=<%sendb(CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, ""))%>&amp;searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=RedeemedCSR&amp;SortDirection=<% Sendb(SortDirection) %>">
              <% Sendb(Copient.PhraseLib.Lookup("term.admin", LanguageID))%>
            </a>
            <%
              If SortText = "RedeemedCSR" Then
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
          If sizeOfData > 0 Then
            Dim ExtLocationDT As DataTable
            MyCommon.Open_LogixRT()
            While (i < sizeOfData And i < linesPerPage + linesPerPage * PageNum)
              Send("      <tr class=""" & Shaded & """>")
              Send("        <td style=""text-align:center;"">")
              If MyCommon.NZ(BarcodeDT.Rows(i).Item("Voided"), 0) = 0 Then
                Send("          <input type=""button"" value=""X"" title=""" & Copient.PhraseLib.Lookup("term.void", LanguageID) & """ name=""ex"" id=""ex-" & BarcodeDT.Rows(i).Item("Barcode") & """ class=""ex"" onclick=""javascript:voidBarcode('" & BarcodeDT.Rows(i).Item("Barcode") & "')""" & IIf(Logix.UserRoles.EditCustomerCoupons, "", " disabled=""disabled""") & " />")
              Else
                Send("          <span style=""color:#aa0000;"">" & Copient.PhraseLib.Lookup("term.void", LanguageID) & "</span>")
              End If
              Send("        </td>")
              'Send("        <td>" & MyCommon.NZ(BarcodeDT.Rows(i).Item("Barcode"), "&nbsp;") & "</td>")
			  Send("        <td> <a href=""#"" onclick=""openWidePopup('/logix/coupon-details.aspx?barcode=" & MyCommon.NZ(BarcodeDT.Rows(i).Item("Barcode"), "&nbsp;") & "');"" />"& MyCommon.NZ(BarcodeDT.Rows(i).Item("Barcode"), "&nbsp;") &"</a></td>")
			  Send("        <td>" & MyCommon.NZ(BarcodeDT.Rows(i).Item("RewardName"),"&nbsp;") & "</td>")
			  Send("        <td>" & MyCommon.NZ(BarcodeDT.Rows(i).Item("IssuingCostCenter"), "&nbsp;") & "</td>")
              If IsDBNull(BarcodeDT.Rows(i).Item("IssueDate")) Then
                Send("        <td>" & Copient.PhraseLib.Lookup("term.unissued", LanguageID) & "</td>")
              Else
                Send("        <td>" & Logix.ToShortDateTimeString(BarcodeDT.Rows(i).Item("IssueDate"), MyCommon) & "</td>")
              End If
			  If (MyCommon.NZ(BarcodeDT.Rows(i).Item("ExpirationDate"), "1/1/2100") = "1/1/2100") Then
                Send("        <td>" & Copient.PhraseLib.Lookup("term.never", LanguageID) & "</td>")
              Else
                Send("        <td>" & Logix.ToShortDateTimeString(BarcodeDT.Rows(i).Item("ExpirationDate"), MyCommon) & "</td>")
              End If
              If MyCommon.NZ(BarcodeDT.Rows(i).Item("RedeemedLocationID"), 0) = "-9" Then
                Send("        <td>" & Copient.PhraseLib.Lookup("term.logix", LanguageID) & "</td>")
              Else
                'Send("        <td>" & MyCommon.NZ(BarcodeDT.Rows(i).Item("RedeemedLocationID"), "&nbsp;") & "</td>")
                If Not IsDBNull(BarcodeDT.Rows(i).Item("RedeemedLocationID")) Then
                  MyCommon.QueryStr = "select ExtLocationCode from Locations where ExtLocationCode='" & BarcodeDT.Rows(i).Item("RedeemedLocationID") &"'"
                  ExtLocationDT = MyCommon.LRT_Select()
                  If (ExtLocationDT.Rows.Count > 0) Then
                    Send("        <td>" & MyCommon.NZ(ExtLocationDT.Rows(0).Item("ExtLocationCode"), "&nbsp;") & "</td>")
                  Else
                    Send("        <td>&nbsp;</td>")
                  End If
                Else
                  Send("        <td>&nbsp;</td>")
                End If
              End If
              If IsDBNull(BarcodeDT.Rows(i).Item("RedeemedDate")) Then
                Send("        <td>" & Copient.PhraseLib.Lookup("term.unredeemed", LanguageID) & "</td>")
              Else
                Send("        <td>" & Logix.ToShortDateTimeString(BarcodeDT.Rows(i).Item("RedeemedDate"), MyCommon) & "</td>")
              End If
			  If IsDBNull(BarcodeDT.Rows(i).Item("RedeemedCSR")) Then
          Send("		  <td>&nbsp;</td>")
			  Else
          MyCommon.QueryStr = "select UserName from AdminUsers where AdminUserID=" & BarcodeDT.Rows(i).Item("RedeemedCSR").ToString()
          AdminID = MyCommon.LRT_Select()
          If (AdminID.Rows.Count > 0) Then
            Send("        <td>" & MyCommon.NZ(AdminID.Rows(0).Item("UserName"),"&nbsp;") & "</td>")
          Else
            Send("		  <td>&nbsp;</td>")
          End If
			  End If
              
              Send("      </tr>")
              If Shaded = "shaded" Then
                Shaded = ""
              Else
                Shaded = "shaded"
              End If
              i = i + 1
            End While
            MyCommon.Close_LogixRT()
          Else
            Send("      <tr class=""" & Shaded & """>")
            Sendb("        <td colspan=""9"" style=""text-align:center;"">")
            If filterCoupon = 0 Then
              If idSearchText <> "" OrElse startDate <> "" OrElse startDate <> "" Then
                Sendb(Copient.PhraseLib.Lookup("customer-coupons.NoMatchCoupons", LanguageID))
              Else
                Sendb(Copient.PhraseLib.Lookup("customer-coupons.NoCoupons", LanguageID))
              End If
            ElseIf filterCoupon = 1 Then
              If idSearchText <> "" OrElse startDate <> "" OrElse startDate <> "" Then
                Sendb(Copient.PhraseLib.Lookup("customer-coupons.NoMatchNonexpired", LanguageID))
              Else
                Sendb(Copient.PhraseLib.Lookup("customer-coupons.NoNonexpired", LanguageID))
              End If
            ElseIf filterCoupon = 2 Then
              If idSearchText <> "" OrElse startDate <> "" OrElse startDate <> "" Then
                Sendb(Copient.PhraseLib.Lookup("customer-coupons.NoMatchExpired", LanguageID))
              Else
                Sendb(Copient.PhraseLib.Lookup("customer-coupons.NoExpired", LanguageID))
              End If
            ElseIf filterCoupon = 3 Then
            If idSearchText <> "" OrElse startDate <> "" OrElse startDate <> "" Then
                Sendb(Copient.PhraseLib.Lookup("customer-coupons.NoMatchVoided", LanguageID))
              Else
                Sendb(Copient.PhraseLib.Lookup("customer-coupons.NoVoided", LanguageID))
            End If
            End If
            Send("</td>")
            Send("      </tr>")
          End If
        %>
      </tbody>
    </table>
  </div>
</form>

<%
done:
Finally
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  Logix = Nothing
  MyCommon = Nothing
End Try
Send_BodyEnd("searchform", "searchterms")
%>
