<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>

<script runat="server" >

  Function shouldDoCustomizedCustomerInquiry(ByRef MyCommon As Copient.CommonInc) As Boolean
    Const USE_CUSTOMIZED_CUSTOMER_INQUIRY As Integer = 107
    Return (MyCommon.Fetch_SystemOption(USE_CUSTOMIZED_CUSTOMER_INQUIRY) = 1)
  End Function
  
  Function validateCard(ByVal card As String, ByRef MyCommon As Copient.CommonInc) As String
    Const PHYSICAL_CARD_LENGTH As Integer = 12
    Const MEMBER_ID_LENGTH As Integer = 15
    card = Trim(card)
        
    Dim cardConverter As New Copient.CustomizedCustomerInquiry(MyCommon.Get_Install_Path() & "/AgentFiles/CustomizedCustomerInquiryCard.config")
    If (card.Length = PHYSICAL_CARD_LENGTH AndAlso IsNumeric(card)) Then
            card = cardConverter.getMemberIdFromCardNumber(card)
    ElseIf (card.Length = MEMBER_ID_LENGTH AndAlso IsNumeric(card)) Then
      Dim physical_card As String = cardConverter.getCardNumberFromMemberId(Long.Parse(card))
    Else
      Throw New ArgumentException(String.Format("{0} ({1})", Copient.PhraseLib.Lookup("term.invalid-cust-specific-card-number", LanguageID), card))
    End If
    Return card
  End Function
    
    
  Function transformCard(ByVal card As String, ByVal cardType As String, ByRef MyCommon As Copient.CommonInc) As String
    Const STANDARD_CUSTOMER_CARD_TYPEID As String = "0"
    
    If (cardType = STANDARD_CUSTOMER_CARD_TYPEID AndAlso Not isEmpty(card) AndAlso shouldDoCustomizedCustomerInquiry( MyCommon ) ) Then
      Try
        
        Return validateCard(card, MyCommon)
        
      Catch ex As Exception ' ensure any exceptions get logged
        MyCommon.Write_Log(MyCommon.LogPath & "/customerinquiry.txt", ex.Message, True)
        Throw
      End Try
    End If ' card is not empty and should do the transform/validation 
    
    Return card
  End Function
  
  Function ReplaceSpecialChar(ByVal inputString As String)
    'Replacing % with [%] will allow sql to search for % when using a like.
    'Replacing _ with [_] will allow sql to search for _ when using a like.
    Return inputString.Trim().Replace("%", "[%]").Replace("_", "[_]")
  End Function
</script>

<%
  ' *****************************************************************************
  ' * FILENAME: coupon-inquiry.aspx 
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
    Dim MyCryptLib As New Copient.CryptLib
  Dim rst As DataTable
  Dim result As Boolean = False
  Dim shaded As Boolean = True
  Dim Restricted As Boolean = False
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  'Dim SearchTerms As String = ""
  Dim i as Integer
  Dim sizeOfData as Integer
  'Dim AdminID as DataTable
  Dim BarcodeDT as  System.Data.DataTable
  Dim custQueryString as String  =""
  Dim VoidBarcode As String = ""
  Dim CustomerPK As Long = 0
  'Dim ROID as Integer
  Dim SortText As String = "IssueDate"
  Dim SortDirection As String = "ASC"
  Dim barcodeQueryString as String
  Dim PageNum As Integer
  Dim MorePages As Boolean
  Dim linesPerPage As Integer = 50
  
  Dim URLtrackBack As String = ""
  Dim inCardNumber As String = ""
  ' tack on the customercare remote links if needed
  Dim extraLink As String = ""
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "coupon-inquiry.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
    PageNum = Request.QueryString("pagenum")
  If PageNum < 0 Then
    PageNum = 0
  End If
  MorePages = False
  
  ' lets check the logged in user and see if they are to be restricted to this page
  MyCommon.QueryStr = "select AU.StartPageID,AUSP.PageName ,AUSP.prestrict from AdminUsers as AU with (NoLock) " & _
                      "inner join AdminUserStartPages as AUSP with (NoLock) on AU.StartPageID=AUSP.StartPageID " & _
                      "where AU.AdminUserID=" & AdminUserID
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    If (rst.Rows(0).Item("prestrict") = True) Then
      ' ok we got in here then we need to restrict the user from seeing any other pages
      Restricted = True
    End If
  End If
  
   If (Request.QueryString("SortText") <> "") Then
      SortText = Request.QueryString("SortText")
    End If
    If (Request.QueryString("SortDirection") = "ASC") Then
      SortDirection = "DESC"
      ElseIf (Request.QueryString("SortDirection") = "DESC") Then
        SortDirection = "ASC"
      Else
        SortDirection = "ASC"
      End If
	  	  
  'find the customer PK
  If(Request.QueryString("memid") <> "") Then
	'SearchTerms = Request.QueryString("memid").PadLeft(MyCommon.Fetch_SystemOption(53), "0")
    Const STANDARD_CUSTOMER_CARD_TYPEID As String = "0"
    Dim ExtCardID As String = ""
    ExtCardID = transformCard(MyCommon.Pad_ExtCardID(Request.QueryString("memid"),0), STANDARD_CUSTOMER_CARD_TYPEID, MyCommon)

    MyCommon.QueryStr = "select CustomerPK from CardIDs C where ExtCardID = '" & MyCryptLib.SQL_StringEncrypt(ExtCardID) & "' AND CardTypeID =0"
    rst = MyCommon.LXS_Select
    If rst.Rows.Count > 0 Then
      CustomerPK = rst.Rows(0).Item("CustomerPK")
      custQueryString = "C.CustomerPK = '" & CustomerPK & "'"
    Else
      infoMessage = Copient.PhraseLib.Lookup("coupon-inquiry.nocardnumber", LanguageID)
    End If
  End If
  

  If Request.QueryString("barcode") <> "" Then
    MyCommon.QueryStr = "Select MappedID from UniqueUPCManufacturerCodeMapping WHERE ManufacturerCode = SUBSTRING('" & Request.QueryString("barcode") & "',1,5)"
    rst = MyCommon.LXS_Select
    If rst.Rows.Count > 0 Then
      barcodeQueryString = "LEFT(barcode,6) = '" & rst.Rows(0).Item("MappedID") & "' + SUBSTRING('" & Request.QueryString("barcode") & "',6,10)"
    Else
      infoMessage = Copient.PhraseLib.Lookup("coupon-inquiry.nobarcode", LanguageID)
    End If
  End If
  
  If (Request.QueryString("coupon") <> "") OrElse (custQueryString <> "") OrElse (barcodeQueryString <> "") Then
    'SearchTerms = Request.QueryString("barcode").PadLeft(MyCommon.Fetch_SystemOption(52), "0")
    MyCommon.QueryStr = "select Voided, Barcode, IssuingCSR, IssuingCostCenter, IssueDate, ExpirationDate, RedeemedLocationID, RedeemedDate, RedeemedCSR, RewardOptionID, C.initialcardid, generatedon " & _
           "from BarcodeDetails B  with (NoLock) inner join Customers C on B.CustomerPK = C.CustomerPK where "
    If (Request.QueryString("coupon") <> "") Then
      MyCommon.QueryStr = MyCommon.QueryStr & "barcode = '" & MyCommon.Parse_Quotes(Request.QueryString("coupon")) & "'"
    End If
    If (Request.QueryString("coupon") <> "") AndAlso (custQueryString <> "") Then
      MyCommon.QueryStr = MyCommon.QueryStr & " AND "
    End If
    If (custQueryString <> "") Then
      MyCommon.QueryStr = MyCommon.QueryStr & custQueryString
    End If
    If (barcodeQueryString <> "") AndAlso ((Request.QueryString("coupon") <> "") Or (custQueryString <> "")) Then
      MyCommon.QueryStr = MyCommon.QueryStr & " AND "
    End If
	
    If (barcodeQueryString <> "") Then
      MyCommon.QueryStr = MyCommon.QueryStr & barcodeQueryString
    End If
	
    BarcodeDT = MyCommon.LXS_Select
	
    If BarcodeDT.Rows.Count > 0 Then
      result = True
      BarcodeDT.Columns.Add("RewardName", GetType(String))
      If BarcodeDT.Rows.Count > 0 Then
        For Each row As DataRow In BarcodeDT.Rows
          'ROID = BarcodeDT.Rows(i).Item("RewardOptionID") 
          If IsDBNull(row("RewardOptionID")) Then
            row("RewardName") = Copient.PhraseLib.Lookup("barcode.rewardnotfound", LanguageID)
          Else
            MyCommon.QueryStr = "dbo.pt_GetCouponRewardName"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@ROID", System.Data.SqlDbType.BigInt).Value = row("RewardOptionID")
            MyCommon.LRTsp.Parameters.Add("@RewardName", System.Data.SqlDbType.NVarChar, 255).Direction = System.Data.ParameterDirection.Output
            MyCommon.LRTsp.ExecuteNonQuery()
            row("RewardName") = MyCommon.LRTsp.Parameters("@RewardName").Value
            If IsDBNull(row("RewardName")) Then row("RewardName") = Copient.PhraseLib.Lookup("barcode.rewardnotfound", LanguageID)
					
          End If
        Next
        BarcodeDT.DefaultView.Sort = SortText & " " & SortDirection
        BarcodeDT = BarcodeDT.DefaultView.ToTable()
			
      End If
    Else
      infoMessage = Copient.PhraseLib.Lookup("coupon-inquiry.nobarcode", LanguageID)
    End If
    sizeOfData = BarcodeDT.Rows.Count
    i = linesPerPage * PageNum
  End If
   
   	If (Request.QueryString("export") <> "") Then
    ' they want to download the group get it from the database and stream it to the client
  
    If result = True Then
      Dim time As DateTime = Now()
      Response.AddHeader("Content-Disposition", "attachment; filename=CouponList." & MyCommon.Leading_Zero_Fill(Year(time), 4) & _
      MyCommon.Leading_Zero_Fill(Month(time), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(time), 2) & ".txt")
		
      Response.ContentType = "application/octet-stream"
      Send(Copient.PhraseLib.Lookup("term.upc", LanguageID) & ", " & Copient.PhraseLib.Lookup("term.assignedto", LanguageID) & "," & Copient.PhraseLib.Lookup("coupon-details.generatedon", LanguageID) & "," & Copient.PhraseLib.Lookup("term.redeemedon", LanguageID))
      For Each row In BarcodeDT.Rows
        Sendb(MyCommon.NZ(row.Item("Barcode"), ""))
        Sendb(",")
        Sendb(MyCryptLib.SQL_StringDecrypt(row.Item("initialcardid").ToString()))
        Sendb(",")
        Sendb(MyCommon.NZ(row.Item("generatedon"), ""))
        Sendb(",")
        Send(MyCommon.NZ(row.Item("RedeemedDate"), ""))
      Next
      GoTo done
    End If
  End If
   

    If (Request.QueryString("VoidBarcode") <> "") Then
		VoidBarcode = Request.QueryString("VoidBarcode")
		MyCommon.QueryStr = "select Barcode from BarcodeDetails " & _
							" where Barcode='" & VoidBarcode & "' and ISNULL(Voided, 0)=0;"
		rst = MyCommon.LXS_Select
		If (rst.Rows.Count > 0) Then
			MyCommon.QueryStr = "update BarcodeDetails set Voided=1, RedeemedDate=getdate(), RedeemedLocationID='-9', RedeemedCSR=" & AdminUserID & " " & _
								 " where Barcode='" & VoidBarcode & "';"
			MyCommon.LXS_Execute()
			MyCommon.Activity_Log2(25, 24, CustomerPK, AdminUserID, Copient.PhraseLib.Lookup("history.customer-void", LanguageID) & " " & VoidBarcode)
		End If
		Response.Redirect("/logix/coupon-inquiry.aspx?barcode=" & Request.QueryString("barcode") & "&memid="& Request.QueryString("memid") )
	End If
  
  ' lets check the logged in user and see if they are to be restricted to this page
  MyCommon.QueryStr = "select AU.StartPageID,AUSP.PageName ,AUSP.prestrict from AdminUsers as AU with (NoLock) " & _
                      "inner join AdminUserStartPages as AUSP with (NoLock) on AU.StartPageID=AUSP.StartPageID " & _
                      "where AU.AdminUserID=" & AdminUserID
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    If (rst.Rows(0).Item("prestrict") = True) Then
      ' ok we got in here then we need to restrict the user from seeing any other pages
      Restricted = True
    End If
  End If
  
  If (Request.QueryString("mode") = "summary") Then
    URLtrackBack = Request.QueryString("exiturl")
    inCardNumber = Request.QueryString("cardnumber")
    extraLink = "&mode=summary&exiturl=" & URLtrackBack & "&cardnumber=" & inCardNumber
  End If
   Send("<script type=""text/javascript"">")
    Send("var datePickerDivID = ""datepicker"";")
    Send("")
    Send_Calendar_Overrides(MyCommon)
    Send("")
    Send("function voidBarcode(Barcode) {")
    Send("  if (confirm('" & Copient.PhraseLib.Lookup("customer-coupons.VoidBarcode", LanguageID) & "')) {")
    Send("    document.getElementById('VoidBarcode').value = Barcode;")
    Send("    document.mainform.submit();")
    Send("  } else {")
    Send("    return false;")
    Send("  }")
    Send("}")
    Send("</script>")
	
  Send_HeadBegin("coupon-inquiry.couponinquiry")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld, Restricted)
  Send_Scripts()
  Send_HeadEnd()
  Send_BodyBegin(1)
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  If (Not Restricted) Then
    Send_Tabs(Logix, 3)
    Send_Subtabs(Logix, 34, 4)
  Else
    Send_Subtabs(Logix, 92, 2, LanguageID, ID, extraLink)
  End If
  
  If (Logix.UserRoles.AccessBarcodeInquiry = False) Then
    Send_Denied(1, "perm.accessbarcodeinquiry")
    GoTo done
  End If
%>
<style type="text/css">

#paginator {
  margin: 0 3px 0 0;
  width: 735px;
  }
* html #paginator {
  width: 735px;
  }

</style>
<script type="text/javascript">
	function exportSubmit()
	{
		$("#export").val("export");
		document.getElementById("mainform").submit();
	}
</script>
<form action="#"  id="mainform" name="mainform">
   <input type="hidden" id="export" name="export" value =""/>
  <div id="intro">
    <h1 id="title">
		<% Sendb(Copient.PhraseLib.Lookup("coupon-inquiry.couponinquiry", LanguageID)&":")%>
    </h1>
    <div id="controls">
		<%
		  Send("<input type=""button"" class=""regular"" id=""export"" name=""export"" value=""" & Copient.PhraseLib.Lookup("offer-list.export", LanguageID) & """" & IIf(result = False, "disabled=true", "") & " onclick=""exportSubmit();"" />")
			%>
    </div>
  </div>
  <div id="main">
    <%
		If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
		 Send("<input type=""hidden"" id=""VoidBarcode"" name=""VoidBarcode"" value="""" />")
	%>
	<table class = "list">
		<tr>
			<td style="width:25%;">
				<h2>
					<% Sendb(Copient.PhraseLib.Lookup("coupon-inquiry.coupondetails", LanguageID))%>
				</h2>
				<% Sendb(Copient.PhraseLib.Lookup("coupon-inquiry.enter12digitupc", LanguageID) & ":")%> 
				<br />
				<br class="half" />

			</td>
			<td style="width:25%;">
				<h2>
					<% Sendb(Copient.PhraseLib.Lookup("coupon-inquiry.memberdetails", LanguageID))%>
				</h2>
				<% Sendb(Copient.PhraseLib.Lookup("coupon-inquiry.memberid", LanguageID) &":")%>
				<br />
				<br class="half" />

			<td>
				<h2>
					<%Sendb(Copient.PhraseLib.Lookup("coupon-inquiry.barcodedetails", LanguageID) )%> 
				</h2>
					<%Sendb(Copient.PhraseLib.Lookup("coupon-inquiry.enter10digitbarcode", LanguageID)& ":")%>
				<br />
				<br class="half" />
			<td>
		</tr>
		<tr>
			<td>
				<input type="text" id="coupon" name="coupon" maxlength="100" value="<%Sendb(Request.QueryString("coupon")) %>" />
			</td>
			<td>
				<input type="text" id="memid" name="memid" maxlength="100" value="<%Sendb(Request.QueryString("memid")) %>" />
			</td>
			<td>
			<input type="text" id="barcode" name="barcode" maxlength="100" value="<%Sendb(Request.QueryString("barcode")) %>" />
				
				<input type="submit" class="regular" id="search" name="search" value="<% Sendb(Copient.PhraseLib.Lookup("term.search", LanguageID)) %>" />
			</td>

		</tr>
	</table
    <%
      If (Restricted) Then
        Send("<input type=""hidden"" id=""mode"" name=""mode"" value=""summary"" />")
        Send("<input type=""hidden"" id=""exiturl"" name=""exiturl"" value=" & URLtrackBack & " />")
        Send("<input type=""hidden"" id=""cardnumber"" name=""cardnumber"" value=" & inCardNumber & " />")
      End If
    %>
    
	
	
    <div id="results">
    <% 
		If (result) Then 'VoidBarcode=&coupon=&memid=&barcode=5289400010&search=Search#
			Send_ListbarBarcodes(linesPerPage, sizeOfData, PageNum, , , , "&amp;coupon=" & Request.QueryString("coupon") & "&amp;memid=" &  Request.QueryString("memid")& "&amp;barcode=" & Request.QueryString("barcode") & "&amp;SortText=" & SortText & "&amp;SortDirection=" & SortDirection, , ,false)
			Send("<table class =""list"" summary=""uniquecoupons"">")
			Send("	<thead>")
			'Send("  		<th align=""center"" class=""th-button"" scope=""col"" style=""text-align: center;"" >" & Copient.PhraseLib.Lookup("term.void", LanguageID) & "</th>")
			Send("			<th align=""left"" class=""th-button"" scope=""col"">")
			Send("				<a href=""coupon-inquiry.aspx?barcode="& Request.QueryString("barcode") & "&amp;memid=" & Request.QueryString("memid") & "&amp;SortText=Voided&amp;SortDirection="& SortDirection &""">")
			Send("			"&Copient.PhraseLib.Lookup("term.void", LanguageID) & "</a>")
				  If SortText = "Voided" Then
					If SortDirection = "ASC" Then
					  Send("<span class=""sortarrow"">&#9660;</span>")
					Else
					  Send("<span class=""sortarrow"">&#9650;</span>")
					End If
				  End If
			Send("				</th>")
			'Send("		<th align=""left"" class=""th-code"" scope=""col"">" &   Copient.PhraseLib.Lookup("term.barcode", LanguageID)& "</th>")
			Send("			<th align=""left"" class=""th-status"" scope=""col"">")
			Send("				<a href=""coupon-inquiry.aspx?barcode="& Request.QueryString("barcode") & "&amp;memid=" & Request.QueryString("memid") & "&amp;SortText=Barcode&amp;SortDirection="& SortDirection &""">")
			Send("			"&Copient.PhraseLib.Lookup("term.barcode", LanguageID) & "</a>")
				  If SortText = "Barcode" Then
					If SortDirection = "ASC" Then
					  Send("<span class=""sortarrow"">&#9660;</span>")
					Else
					  Send("<span class=""sortarrow"">&#9650;</span>")
					End If
				  End If
			Send("				</th>")
			'Send("		<th align=""left"" class=""th-rewardname"" scope=""col"">" & Copient.PhraseLib.Lookup("term.rewardname", LanguageID)& "</th>")
			Send("			<th align=""left"" class=""th-rewardname"" scope=""col"">")
			Send("				<a href=""coupon-inquiry.aspx?barcode="& Request.QueryString("barcode") & "&amp;memid=" & Request.QueryString("memid") & "&amp;SortText=RewardName&amp;SortDirection="& SortDirection &""">")
			Send("			"&Copient.PhraseLib.Lookup("term.rewardname", LanguageID) & "</a>")
				  If SortText = "RewardName" Then
					If SortDirection = "ASC" Then
					  Send("<span class=""sortarrow"">&#9660;</span>")
					Else
					  Send("<span class=""sortarrow"">&#9650;</span>")
					End If
				  End If
			Send("				</th>")
			'Send("	    <th align=""left"" class=""th-code"" scope=""col"">" & Copient.PhraseLib.Lookup("term.issucc", LanguageID) & "</th>")
			Send("			<th align=""left"" class=""th-issucc"" scope=""col"">")
			Send("				<a href=""coupon-inquiry.aspx?barcode="& Request.QueryString("barcode") & "&amp;memid=" & Request.QueryString("memid") & "&amp;SortText=IssuingCostCenter&amp;SortDirection="& SortDirection &""">")
			Send("			"&Copient.PhraseLib.Lookup("term.issucc", LanguageID) & "</a>")
				  If SortText = "IssuingCostCenter" Then
					If SortDirection = "ASC" Then
					  Send("<span class=""sortarrow"">&#9660;</span>")
					Else
					  Send("<span class=""sortarrow"">&#9650;</span>")
					End If
				  End If
			Send("				</th>")
			'Send("       <th align=""left"" class=""th-datetime"" scope=""col"">" &  Copient.PhraseLib.Lookup("term.issuedon", LanguageID) & "</th>")
			Send("			<th align=""left"" class=""th-barcodedate"" scope=""col"">")
			Send("				<a href=""coupon-inquiry.aspx?barcode="& Request.QueryString("barcode") & "&amp;memid=" & Request.QueryString("memid") & "&amp;SortText=IssueDate&amp;SortDirection="& SortDirection &""">")
			Send("			"&Copient.PhraseLib.Lookup("term.issuedon", LanguageID) & "</a>")
				  If SortText = "IssueDate" Then
					If SortDirection = "ASC" Then
					  Send("<span class=""sortarrow"">&#9660;</span>")
					Else
					  Send("<span class=""sortarrow"">&#9650;</span>")
					  End If
				  End If
			Send("				</th>")
			'Send("       <th align=""left"" class=""th-datetime"" scope=""col"">" &  Copient.PhraseLib.Lookup("term.expires", LanguageID) & "</th>")
			Send("			<th align=""left"" class=""th-barcodedate"" scope=""col"">")
			Send("				<a href=""coupon-inquiry.aspx?barcode="& Request.QueryString("barcode") & "&amp;memid=" & Request.QueryString("memid") & "&amp;SortText=ExpirationDate&amp;SortDirection="& SortDirection &""">")
			Send("			"&Copient.PhraseLib.Lookup("term.expires", LanguageID) & "</a>")
				  If SortText = "ExpirationDate" Then
					If SortDirection = "ASC" Then
					  Send("<span class=""sortarrow"">&#9660;</span>")
					Else
					  Send("<span class=""sortarrow"">&#9650;</span>")
					End If
				  End If
			Send("				</th>")
			'Send("       <th align=""left"" class=""th-status"" scope=""col"">" & Copient.PhraseLib.Lookup("term.redmloc", LanguageID) & "</th>")
			Send("			<th align=""left"" class=""th-status"" scope=""col"">")
			Send("				<a href=""coupon-inquiry.aspx?barcode="& Request.QueryString("barcode") & "&amp;memid=" & Request.QueryString("memid") & "&amp;SortText=RedeemedLocationID&amp;SortDirection="& SortDirection &""">")
			Send("			"&Copient.PhraseLib.Lookup("term.redmloc", LanguageID) & "</a>")
				  If SortText = "RedeemedLocationID" Then
					If SortDirection = "ASC" Then
					  Send("<span class=""sortarrow"">&#9660;</span>")
					Else
					  Send("<span class=""sortarrow"">&#9650;</span>")
					End If
				  End If
			Send("				</th>")
			'Send("       <th align=""left"" class=""th-datetime"" scope=""col"">" &  Copient.PhraseLib.Lookup("term.redeemedon", LanguageID) & "</th>")
			Send("			<th align=""left"" class=""th-barcodedate"" scope=""col"">")
			Send("				<a href=""coupon-inquiry.aspx?barcode="& Request.QueryString("barcode") & "&amp;memid=" & Request.QueryString("memid") & "&amp;SortText=RedeemedDate&amp;SortDirection="& SortDirection &""">")
			Send("			"&Copient.PhraseLib.Lookup("term.redeemedon", LanguageID) & "</a>")
				  If SortText = "RedeemedDate" Then
					If SortDirection = "ASC" Then
					  Send("<span class=""sortarrow"">&#9660;</span>")
					Else
					  Send("<span class=""sortarrow"">&#9650;</span>")
					End If
				  End If
			Send("				</th>")
			'Send("      <th align=""left"" class=""th-csr"" scope=""col"">" & "Admin" & "</th>")
			Send("  </thead>")
			Send("  <tbody>")
			Shaded = true
			
			sizeOfData = BarcodeDT.Rows.Count
            Dim ExtLocationDT As DataTable
            While (i < sizeOfData And i < linesPerPage + linesPerPage * PageNum)
			  If Shaded = true Then
				Send("      <tr class=""shaded"">")
			  else
				Send("      <tr class=>")
			  End If
               Send("        <td style=""text-align:center;"">")
              If MyCommon.NZ(BarcodeDT.Rows(i).Item("Voided"), 0) = 0 Then
                Send("          <input type=""button"" value=""X"" title=""" & Copient.PhraseLib.Lookup("term.void", LanguageID) & """ name=""ex"" id=""ex-" & BarcodeDT.Rows(i).Item("Barcode") & """ class=""ex"" onclick=""javascript:voidBarcode('" & BarcodeDT.Rows(i).Item("Barcode") & "')""" & IIf(Logix.UserRoles.EditCustomerCoupons, "", " disabled=""disabled""") & " />")
              Else
                Send("          <span style=""color:#aa0000;"">" & Copient.PhraseLib.Lookup("term.void", LanguageID) & "</span>")
              End If
              Send("        </td>")
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
                  MyCommon.QueryStr = "select ExtLocationCode from Locations where ExtLocationCode= '" & BarcodeDT.Rows(i).Item("RedeemedLocationID") & "'"
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
			 ' If IsDBNull(BarcodeDT.Rows(i).Item("RedeemedCSR")) Then
			'	Send("		  <td>&nbsp;</td>")
			 ' Else
			'	MyCommon.QueryStr = "select UserName from AdminUsers where AdminUserID=" & BarcodeDT.Rows(i).Item("RedeemedCSR").ToString()
			'	AdminID = MyCommon.LRT_Select()
			'	Send("        <td>" & MyCommon.NZ(AdminID.Rows(0).Item("UserName"),"&nbsp;") & "</td>")
			 ' End If
              
              Send("      </tr>")
              If Shaded = true Then
                Shaded = false
              Else
                Shaded = true
              End If
              i = i + 1
            End While

			Send("  </tbody>")
			Send("</table>")
		End If
	
	%>

		
    </div>
  </div>
</form>
<%
  'If MyCommon.Fetch_SystemOption(75) Then
  '  If (Logix.UserRoles.AccessNotes) Then
  '    Send_Notes(6, 0, AdminUserID)
  '  End If
  'End If
  Send_BodyEnd("mainform", "searchterms")
done:

  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  Logix = Nothing
  MyCommon = Nothing
%>
