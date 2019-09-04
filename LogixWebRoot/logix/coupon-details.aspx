<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: product-inquiry.aspx 
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
  'Dim row As DataRow
  Dim rst2 As DataTable
  'Dim row2 As DataRow
  Dim rst3 As DataTable
  'Dim row3 As DataRow
  'Dim amtMax As String
  'Dim description As String = ""
  'Dim productGroups() As String
  'Dim productGroupsIDs() As String
  'Dim productGroupList As String
  'Dim prodGroups As String
  'Dim promotionsList As String
  'Dim promotions() As String
  'Dim promotionsIDs() As String
  Dim result As Boolean = False
  'Dim shaded As Boolean = True
  ' Dim Restricted As Boolean = False
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  'Dim SearchTerms As String = ""
  'Dim TempDate As Date
  'Dim i as Integer
  'Dim sizeOfData as Integer
  Dim AdminID as DataTable
  Dim BarcodeDT as  System.Data.DataTable
  Dim custQueryString as String  =""
  Dim barcodeQueryString as String = ""
  Dim ExtLocationDT As DataTable
  Dim VoidBarcode As String = ""
  Dim CustomerPK As Long = 0
  Dim RewardName as String =""
  Dim URLtrackBack As String = ""
  Dim inCardNumber As String = ""
  ' tack on the customercare remote links if needed
  Dim extraLink As String = ""
  Dim ROID as Integer
  Dim dt as DataTable
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "coupon-deatails.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  ' lets check the logged in user and see if they are to be restricted to this page
  ' MyCommon.QueryStr = "select AU.StartPageID,AUSP.PageName ,AUSP.prestrict from AdminUsers as AU with (NoLock) " & _
                      ' "inner join AdminUserStartPages as AUSP with (NoLock) on AU.StartPageID=AUSP.StartPageID " & _
                      ' "where AU.AdminUserID=" & AdminUserID
  ' rst = MyCommon.LRT_Select
  ' If rst.Rows.Count > 0 Then
    ' If (rst.Rows(0).Item("prestrict") = True) Then
      ' ok we got in here then we need to restrict the user from seeing any other pages
      ' Restricted = True
    ' End If
  ' End If

  
	If (Request.QueryString("barcode") <> "")Then
		'SearchTerms = Request.QueryString("barcode").PadLeft(MyCommon.Fetch_SystemOption(52), "0")
		MyCommon.QueryStr = "select CustomerPK from BarcodeDetails where Barcode = '" & MyCommon.Parse_Quotes(Request.QueryString("barcode")) & "' "
		rst2 = MyCommon.LXS_Select
		If rst2.Rows.Count >0 Then
			CustomerPK = 999
		End If
		
		MyCommon.QueryStr = "select Voided, Barcode, IssuingCSR, IssuingCostCenter, IssueDate, ExpirationDate, RedeemedLocationID, RedeemingMemberID, RedeemedDate, RedeemedCSR, GeneratedOn, EffectiveDate, IssuingTransactionID, "& _
										"RedeemingTransactionID, CustomerPK, RewardOptionID from BarcodeDetails  with (NoLock) where Barcode = '" &  MyCommon.Parse_Quotes(Request.QueryString("barcode")) & "' "
		BarcodeDT = MyCommon.LXS_Select
		If BarcodeDT.Rows.Count >0 Then
			result=true
                  MyCommon.QueryStr = "select ExtCardIDOriginal as ExtCardID from CardIDs with (NoLock) where CardTypeID=0 and CustomerPK = " & _ 
                                       MyCommon.NZ(BarcodeDT.Rows(0).Item("RedeemingMemberID"),"0") & ";"
                  rst3 = MyCommon.LXS_Select
		Else
			infoMessage = "Couldn't find Barcode"
		End If
	Else
			infoMessage = "Invalid barcode"
	End If
   
   If (Request.QueryString("VoidBarcode") <> "") Then
		VoidBarcode = Request.QueryString("VoidBarcode")
		MyCommon.QueryStr = "select Barcode, CustomerPK from BarcodeDetails " & _
							"where Barcode='" & VoidBarcode & "' and ISNULL(Voided, 0)=0;"
		rst = MyCommon.LXS_Select
		If (rst.Rows.Count > 0) Then
			MyCommon.QueryStr = "update BarcodeDetails set Voided=1, RedeemedDate=getdate(), RedeemedLocationID='-9', RedeemedCSR=" & AdminUserID & " " & _
								"where Barcode='" & VoidBarcode & "';"
			MyCommon.LXS_Execute()
			MyCommon.Activity_Log2(25, 24, rst.Rows(0).Item("CustomerPK"), AdminUserID, Copient.PhraseLib.Lookup("history.customer-void", LanguageID) & " " & VoidBarcode)
		End If
		Response.Redirect("/logix/coupon-details.aspx?barcode=" &  VoidBarcode)
	End If
  
  ' lets check the logged in user and see if they are to be restricted to this page
  ' MyCommon.QueryStr = "select AU.StartPageID,AUSP.PageName ,AUSP.prestrict from AdminUsers as AU with (NoLock) " & _
                      ' "inner join AdminUserStartPages as AUSP with (NoLock) on AU.StartPageID=AUSP.StartPageID " & _
                      ' "where AU.AdminUserID=" & AdminUserID
  ' rst = MyCommon.LRT_Select
  ' If rst.Rows.Count > 0 Then
    ' If (rst.Rows(0).Item("prestrict") = True) Then
      ' ok we got in here then we need to restrict the user from seeing any other pages
      ' Restricted = True
    ' End If
  ' End If
  
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
  
  Send_HeadBegin("term.coupondetails")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
  Send_HeadEnd()
  Send_BodyBegin(2)
'  Send_Bar(Handheld)
  'Send_Help(CopientFileName)
'  Send_Logos()
  'If (Not Restricted) Then
    'Send_Tabs(Logix, 4)
'    Send_Subtabs(Logix, 40, 2)
  'Else
    'Send_Subtabs(Logix, 92, 2, LanguageID, ID, extraLink)
  'End If
  
  If (Logix.UserRoles.AccessBarcodeInquiry = False) Then
    Send_Denied(1, "perm.accessbarcodeinquiry")
    GoTo done
  End If
%>
<form action="#"  id="mainform" name="mainform">
  <div id="intro">
    <h1 id="title">
	<%
	  If (result) Then
		If IsDBNull(BarcodeDT.Rows(0).Item("RewardOptionID")) Then
				RewardName = ""
		Else
			' ROID = BarcodeDT.Rows(0).Item("RewardOptionID") 
			' MyCommon.QueryStr = "select Value as 'RewardName'  from PassThruTierValues as D "& _
												' "INNER JOIN CPE_Deliverables AS P " & _
												' "on D.PTPKID = P.OutputID " & _
												' "where P.RewardOptionID = " & ROID & " AND D.PassThruPresTagID = 11 AND P.DeliverableTypeID = 12"
			' dt = MyCommon.LRT_Select
			
			' If dt.Rows.Count > 0 Then
				' RewardName =  MyCommon.NZ(dt.Rows(0).Item("RewardName"),"&nbsp;") 
			' Else
				' RewardName = ""
			' End If
			
			MyCommon.QueryStr = "dbo.pt_GetCouponRewardName"
			MyCommon.Open_LRTsp()
			MyCommon.LRTsp.Parameters.Add("@ROID", System.Data.SqlDbType.Bigint).Value = BarcodeDT.Rows(0).Item("RewardOptionID") 
			MyCommon.LRTsp.Parameters.Add("@RewardName", System.Data.SqlDbType.Nvarchar, 255).Direction = System.Data.ParameterDirection.Output
			MyCommon.LRTsp.ExecuteNonQuery()
			RewardName =  MyCommon.NZ(MyCommon.LRTsp.Parameters("@RewardName").Value,"")
			
		End If
      Send("Coupon Detail - " & MyCommon.NZ(BarcodeDT.Rows(0).Item("Barcode"), "&nbsp;") & " - " & RewardName)  
	  End If 
	   
	  %>
    </h1>
    <div id = "controls">
      <!--div class="actionsmenu" id="actionsmenu"-->  
        
		<%
			'  If (result) Then
			'	  If MyCommon.NZ(BarcodeDT.Rows(0).Item("Voided"), 0) = 0 Then
			'		Send("          <input type=""button"" value=""Void Coupon"" title=""" & Copient.PhraseLib.Lookup("term.void", LanguageID) & """ name=""ex"" id=""ex-" & Request.QueryString("barcode") & """style=""color: #cc0000;font-size: 11px; font-weight: bold; width = 110px; height = 25px"" onclick=""javascript:voidBarcode('" & BarcodeDT.Rows(0).Item("Barcode") & "')""" & IIf(Logix.UserRoles.EditCustomerCoupons, "", " disabled=""disabled""") & " />")
			'	  Else
			'		Send("          <span style=""color:#aa0000;font-size:15px;"">" & Copient.PhraseLib.Lookup("term.void", LanguageID) & "</span>")
			'	  End If
			 ' End If
			  %>
      <!--/div-->
    </div>
  </div>
  
  <div id="main">
	 
	 <%If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>") 
	  Send("<input type=""hidden"" id=""VoidBarcode"" name=""VoidBarcode"" value="""" />")
	 %>
    <div id="coulmn">
	<%
	
	If (result) Then
		Send("<div class=""box"">")
        Send("<h2>")
        Send("	<span>" &  Copient.PhraseLib.Lookup("term.activity",LanguageID) & "</span>")
		
        Send("</h2>")
		If MyCommon.NZ(BarcodeDT.Rows(0).Item("Voided"), 0) = 0 Then
					Send("          <input type=""button"" value=""Void Coupon"" title=""" & Copient.PhraseLib.Lookup("term.void", LanguageID) & """ name=""ex"" id=""ex-" & Request.QueryString("barcode") & """style=""color: #cc0000;font-size: 11px; font-weight: bold; width = 110px; height = 25px"" onclick=""javascript:voidBarcode('" & BarcodeDT.Rows(0).Item("Barcode") & "')""" & IIf(Logix.UserRoles.EditCustomerCoupons, "", " disabled=""disabled""") & " />")
				  Else
					Send("          <span style=""color:#aa0000;font-size:15px;"">" & Copient.PhraseLib.Lookup("term.void", LanguageID) & "</span>")
				  End If
        Send("<table>")
        Send("	<tbody>")
        Send("	<tr>")
        Send("		<td>" & Copient.PhraseLib.Lookup("term.issuedon", LanguageID) &":</td>")

        If IsDBNull(BarcodeDT.Rows(0).Item("IssueDate")) Then
			Send("        <td>" & Copient.PhraseLib.Lookup("term.unissued", LanguageID) & "</td>")
        Else
            Send("        <td>" & Logix.ToShortDateTimeString(BarcodeDT.Rows(0).Item("IssueDate"), MyCommon) & "</td>")
        End If
        Send("		<td>"&Copient.PhraseLib.Lookup("term.redeemedon", LanguageID) & ":</td>")
        
        If IsDBNull(BarcodeDT.Rows(0).Item("RedeemedDate")) Then
			Send("        <td>" & Copient.PhraseLib.Lookup("term.unredeemed", LanguageID) & "</td>")
        Else
            Send("        <td>" & Logix.ToShortDateTimeString(BarcodeDT.Rows(0).Item("RedeemedDate"), MyCommon) & "</td>")
        End If
        Send("	</tr>")
        Send("	<tr>")
		Send("		<td>" & Copient.PhraseLib.Lookup("term.issuingcostcenter", LanguageID) & "</td>")
		Send("       <td>" & MyCommon.NZ(BarcodeDT.Rows(0).Item("IssuingCostCenter"), "&nbsp;") & "</td>")

        Send("		<td>" & Copient.PhraseLib.Lookup("term.redeemingloc", LanguageID) & "</td>")
        
        If MyCommon.NZ(BarcodeDT.Rows(0).Item("RedeemedLocationID"), 0) = "-9" Then
			Send("        <td>" & Copient.PhraseLib.Lookup("term.logix", LanguageID) & "</td>")
        Else
			'Send("        <td>" & MyCommon.NZ(BarcodeDT.Rows(i).Item("RedeemedLocationID"), "&nbsp;") & "</td>")
            If Not IsDBNull(BarcodeDT.Rows(0).Item("RedeemedLocationID")) Then
				MyCommon.QueryStr = "select ExtLocationCode from Locations where ExtLocationCode='" & BarcodeDT.Rows(0).Item("RedeemedLocationID") & "'"
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
        Send("	</tr>")
        Send("	<tr>")
        Send("		<td>" & Copient.PhraseLib.Lookup("term.issuingtrans", LanguageID) &":</td>")
        
        Send("		<td>"&  MyCommon.NZ(BarcodeDT.Rows(0).Item("IssuingTransactionID"),"&nbsp;")&"</td>")
        Send("		<td>" &Copient.PhraseLib.Lookup("term.redeemingtrans", LanguageID) & ":</td>")
        
        Send("		<td>"&  MyCommon.NZ(BarcodeDT.Rows(0).Item("RedeemingTransactionID"),"&nbsp;")&"</td>")
		Send("	</tr>")
		Send("	<tr>")
		Send("		<td>"&Copient.PhraseLib.Lookup("coupon-details.generatedon", LanguageID)&": </td>")
		Send("		<td>" & MyCommon.NZ(BarcodeDT.Rows(0).Item("GeneratedOn"),"&nbsp;") & "</td>")

		Send("		<td>"&Copient.PhraseLib.Lookup("term.expires", LanguageID) &"</td>")
		
		If (MyCommon.NZ(BarcodeDT.Rows(0).Item("ExpirationDate"), "1/1/2100") = "1/1/2100") Then
			Send("        <td>" & Copient.PhraseLib.Lookup("term.never", LanguageID) & "</td>")
        Else
            Send("        <td>" & Logix.ToShortDateTimeString(BarcodeDT.Rows(0).Item("ExpirationDate"), MyCommon) & "</td>")
        End If
		Send("	</tr>")
		Send("	<tr>")
		Send("		<td>"& Copient.PhraseLib.Lookup("term.assignedto", LanguageID) & ": </td>")
		MyCommon.QueryStr = "Select ExtCardIDOriginal as ExtCardID, CardPK from CardIDs where CustomerPK = " & BarcodeDT.Rows(0).Item("CustomerPK")
		rst = MyCommon.LXS_Select
		if(rst.Rows.Count> 0) then 
		Send("		<td>"& MyCryptLib.SQL_StringDecrypt(rst.Rows(0).Item("ExtCardID").ToString()) & "</td>")
		else
			Send("		<td> &nbsp; </td>")
		End If
		Send("  <td>"& Copient.PhraseLib.Lookup("term.redeemingmemberid", LanguageID) & ":</td>")
		if(rst3.Rows.Count> 0) then
      Send("  <td>"& MyCryptLib.SQL_StringDecrypt(rst3.Rows(0).Item("ExtCardID").ToString()) & "</td>")
    else
      Send("  <td></td>")
    end if
               'Send("  <td>"& MyCommon.NZ(BarcodeDT.Rows(0).Item("RedeemingMemberID"),"") & "</td>")
		Send("	</tr>")
		Send(" <tr>")
		Send("  <td>" & Copient.PhraseLib.Lookup("term.admin", LanguageID) & ":</td>")
		If IsDBNull(BarcodeDT.Rows(0).Item("RedeemedCSR")) Then
          Send("		  <td>&nbsp;</td>")
        Else
          MyCommon.QueryStr = "select UserName from AdminUsers where AdminUserID=" & BarcodeDT.Rows(0).Item("RedeemedCSR").ToString()
          AdminID = MyCommon.LRT_Select()
          If (AdminID.Rows.Count > 0) Then
            Send("        <td>" & MyCommon.NZ(AdminID.Rows(0).Item("UserName"),"&nbsp;") & "</td>")
          Else
            Send("		  <td>&nbsp;</td>")
          End If
		End If
		Send("   </tbody>")
		Send(" </table>")
		Send("</div>")
	End If
	%>
    <br/>
   
    </div>
  </div>
</form>
<%
  'If MyCommon.Fetch_SystemOption(75) Then
  '  If (Logix.UserRoles.AccessNotes) Then
  '    Send_Notes(6, 0, AdminUserID)
  '  End If
  'End If
done:
  Send_BodyEnd("mainform", "searchterms")
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  Logix = Nothing
  MyCommon = Nothing
%>
