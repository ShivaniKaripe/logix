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
    Dim MyCryptlib As New Copient.CryptLib
  Dim Logix As New Copient.LogixInc
  Dim rst As DataTable
  Dim cgnameDT as DataTable
  Dim row as DataRow
  Dim shaded As Boolean = True
  Dim Restricted As Boolean = False
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim i as Integer
  Dim sizeOfData as Integer
  Dim AdminID as DataTable
  Dim BarcodeDT as New System.Data.DataTable
  Dim custQueryString as String  =""
  Dim CustomerPK As Long = 0
  Dim ROID as Integer
  Dim SortText As String = "RequestPK"
  Dim SortDirection As String = "ASC"
  Dim DataAdapter As New System.Data.SqlClient.SqlDataAdapter

  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "coupon-batch-report.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
	
	If (Request.QueryString("SortText") <> "") Then
      SortText = Request.QueryString("SortText")
	  MyCommon.QueryStr &= " order by " & SortText
    End If
    If (Request.QueryString("SortDirection") = "ASC") Then
      SortDirection = "DESC"
    ElseIf (Request.QueryString("SortDirection") = "DESC") Then
        SortDirection = "ASC"
    Else
        SortDirection = "DESC"
    End If
	
	
	MyCommon.QueryStr = "Select RequestPK, RequestCompletedOn, CustomerGroupID from BarcodeBatchRequestQueue where NOT (BarcodeGenerationStart is  null and RequestCompletedOn is not null) " 
	rst = MyCommon.LXS_Select
	If rst.Rows.Count >0 Then
		
		rst.Columns.Add("CustomerGroup",GetType(String))
		rst.Columns.Add("NumberOfBarcodes",GetType(Integer))
		for each row in rst.rows
			
			MyCommon.QueryStr = "select Name from CustomerGroups where CustomerGroupID = " & row.Item("CustomerGroupID")
			 cgnameDT = MyCommon.LRT_Select
			 If cgnameDT.Rows.Count>0 Then
				row("CustomerGroup") = cgnameDT.Rows(0).Item("Name")
			End If
			MyCommon.QueryStr = "DECLARE @bstart as DateTime; DECLARE @bstop as DateTime ; " & _
									"Select @bstart = BarcodeGenerationStart, @bstop = RequestCompletedOn from BarcodeBatchRequestQueue where RequestPK = " & row.Item("RequestPK") & _
									"	select count(b.barcode) as 'NumberOfBarcodes' from BarcodeDetails as b  where  b.GeneratedOn >= @bstart and  b.GeneratedOn<= @bstop "
			BarcodeDT = MyCommon.LXS_Select
			row("NumberOfBarcodes") = BarcodeDT.Rows(0).Item("NumberOfBarcodes")
		Next
	End If
	  
	rst.DefaultView.Sort = SortText & " " & SortDirection
	rst = rst.DefaultView.ToTable()

	If(Request.QueryString("export") <> "") Then
		Dim time as DateTime 

		  BarcodeDT= New DataTable
		  MyCommon.QueryStr = "dbo.pt_GetCouponBatchList"
			MyCommon.Open_LXSsp()
			MyCommon.LXSsp.Parameters.Add("@RequestPK", SqlDbType.BigInt).Value = MyCommon.Extract_Val(MyCommon.NZ(Request.QueryString("RequestPK"),-1))
			MyCommon.LXSsp.ExecuteNonQuery()
			DataAdapter.SelectCommand = MyCommon.LXSsp
			DataAdapter.Fill(BarcodeDT)
			MyCommon.Close_LXSsp()
			
		  time= BarcodeDT.Rows(0).Item("BatchGeneratedOn")
		  Response.AddHeader("Content-Disposition", "attachment; filename=BarcodeBatch." &  MyCommon.Leading_Zero_Fill(Year(time), 4) & _ 
		  MyCommon.Leading_Zero_Fill(Month(time), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(time), 2) & ".txt")
			
		  Response.ContentType = "application/octet-stream"
		  Send("MemberID, Barcode")
		  For Each row In BarcodeDT.Rows
            Sendb(MyCryptlib.SQL_StringDecrypt(row.Item("InitialCardID").ToString()))
			Sendb(",")
			Send(MyCommon.NZ(row.Item("barcode"), ""))	
		  Next
		  GoTo done
	End If
	  	  	
  Send_HeadBegin("term.barcodebatchreport")
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
    Send_Subtabs(Logix, 34, 5)
  Else
    Send_Subtabs(Logix, 92, 2, LanguageID, ID)
  End If
  
  If (Logix.UserRoles.AccessBarcodeInquiry = False) Then
    Send_Denied(1, "perm.accessbarcodeinquiry")
    GoTo done
  End If
%>
<script type="text/javascript">
	function setRequestPK(rpk)
	{
		$('#RequestPK').val(rpk);
	}
</script>
<form action="#"  id="mainform" name="mainform">
  <div id="intro">
    <h1 id="title">
		Batch Report
    </h1>
    <div id="controls">
    </div>
  </div>
  <div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
		
	%>
	
    <%
      If (Restricted) Then
        Send("<input type=""hidden"" id=""mode"" name=""mode"" value=""summary"" />")
      End If
    %>
    <input type = "hidden"  id = "RequestPK" name = "RequestPK" />
	
	
    <div id="results">
    <% 
		'If (result) Then
			Send("<table class =""list"" summary=""uniquecoupons"">")
			Send("	<thead>")
			Send("			<th align=""left"" class=""th-code"" scope=""col"">")
			Send("				<a href=""coupon-batch-report.aspx?&SortText=RequestPK&amp;SortDirection="& SortDirection &""">")
			Send("			"& Copient.PhraseLib.Lookup("barcodebatchreport.requestid", LanguageID) & "</a>") 
				  If SortText = "RequestPK" Then
					If SortDirection = "ASC" Then
					  Send("<span class=""sortarrow"">&#9660;</span>")
					Else
					  Send("<span class=""sortarrow"">&#9650;</span>")
					End If
				  End If
			Send("				</th>")
			'Send("		<th align=""left"" class=""th-code"" scope=""col"">" &   Copient.PhraseLib.Lookup("term.barcode", LanguageID)& "</th>")
			Send("			<th align=""left"" scope=""col"">")
			Send("				<a href=""coupon-batch-report.aspx?&SortText=RequestCompletedOn&amp;SortDirection="& SortDirection &""">")
			Send("			"& Copient.PhraseLib.Lookup("term.timecompleted", LanguageID) & "</a>") 
				  If SortText = "RequestCompletedOn" Then
					If SortDirection = "ASC" Then
					  Send("<span class=""sortarrow"">&#9660;</span>")
					Else
					  Send("<span class=""sortarrow"">&#9650;</span>")
					End If
				  End If
			Send("				</th>")
						'Send("		<th align=""left"" class=""th-code"" scope=""col"">" &   Copient.PhraseLib.Lookup("term.barcode", LanguageID)& "</th>")
			Send("			<th align=""left"" class="""" scope=""col"">")
			Send("				<a href=""coupon-batch-report.aspx?&SortText=CustomerGroup&amp;SortDirection="& SortDirection &""">")
			Send("			"& Copient.PhraseLib.Lookup("term.customergroup", LanguageID) & "</a>") 
				  If SortText = "CustomerGroup" Then
					If SortDirection = "ASC" Then
					  Send("<span class=""sortarrow"">&#9660;</span>")
					Else
					  Send("<span class=""sortarrow"">&#9650;</span>")
					End If
				  End If
			Send("				</th>")
			'Send("		<th align=""left"" class=""th-rewardname"" scope=""col"">" & Copient.PhraseLib.Lookup("term.rewardname", LanguageID)& "</th>")
			Send("			<th align=""left"" class="""" scope=""col"">")
			Send("				<a href=""coupon-batch-report.aspx?&SortText=NumberOfBarcodes&amp;SortDirection="& SortDirection &""">")
			Send("			"& Copient.PhraseLib.Lookup("coupon-batch-report.barcodescreated", LanguageID) & "</a>")
				  If SortText = "NumberOfBarcodes" Then
					If SortDirection = "ASC" Then
					  Send("<span class=""sortarrow"">&#9660;</span>")
					Else
					  Send("<span class=""sortarrow"">&#9650;</span>")
					End If
				  End If
			Send("				</th>")
			Send("			<th align=""left"" class="""" scope=""col"">")

			Send("			"&  Copient.PhraseLib.Lookup("term.export", LanguageID) & "")
				  
			Send("				</th>")

			Send("  </thead>")
			Send("  <tbody>")
			For each row in rst.Rows
			  If Shaded = true Then
				Send("    <tr class=""shaded"">")
				Shaded = False
			  else
				Send("    <tr class=>")
				Shaded = True
			  End If
			  Send("      <td>" & row.Item("RequestPK") & "</td>")
			  If isDBNull(row.Item("RequestCompletedOn"))  Then
			    Send("      <td>" &  Copient.PhraseLib.Lookup("term.requestnotcompleted", LanguageID)  & "</td>")
			  Else
			    Send("      <td><a href=""#"" onclick=""openPopup('/logix/coupon-batch-details.aspx?RequestPK=" & row.Item("RequestPK") & "')"">"& row.Item("RequestCompletedOn")  & "</a></td>")
			  End If

				Send("      <td>" &MyCommon.TruncateString(row.Item("CustomerGroup"), 45) & "</td>")

			  Send("      <td>" & row.Item("NumberOfBarcodes") & "</td>")
			    If isDBNull(row.Item("RequestCompletedOn"))  Then
			  Send("      <td><input type=""submit"" class="""" id=""export"" name=""export"" disabled=""disabled"" value=""" & Copient.PhraseLib.Lookup("term.export", LanguageID) & """ onclick=""setRequestPK(" & row.Item("RequestPK") & ")"" /></td>")
			  else
			  Send("      <td><input type=""submit"" class="""" id=""export"" name=""export"" value=""" & Copient.PhraseLib.Lookup("term.export", LanguageID) & """ onclick=""setRequestPK(" & row.Item("RequestPK") & ")"" /></td>")
			  end if
			  Send("    </tr>")
			 
			  
			Next

			Send("  </tbody>")
			Send("</table>")
	'	End If
	
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
