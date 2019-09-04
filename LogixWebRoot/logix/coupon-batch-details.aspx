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
  Dim result As Boolean = False
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim BarcodeDT as New DataTable
  Dim custQueryString as String  =""
  Dim barcodeQueryString as String = ""
  Dim VoidBarcode As String = ""
  Dim CustomerPK As Long = 0
  Dim RewardName as String =""
  Dim URLtrackBack As String = ""
  Dim inCardNumber As String = ""
  Dim extraLink As String = ""
  Dim RequestPK as Integer
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "coupon-deatails.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  RequestPK = MyCommon.Extract_Val(MyCommon.NZ(Request.QueryString("RequestPK"),-1))
  If RequestPK =-1 Then	InfoMessage = "Invalid RequestPK"

  


  
	If (Request.QueryString("RequestPK") <> "")Then
		Dim DataAdapter As New System.Data.SqlClient.SqlDataAdapter

		MyCommon.QueryStr = "dbo.pt_GetCouponBatchList"
        MyCommon.Open_LXSsp()
        MyCommon.LXSsp.Parameters.Add("@RequestPK", SqlDbType.BigInt).Value = RequestPK
        MyCommon.LXSsp.ExecuteNonQuery()
		DataAdapter.SelectCommand = MyCommon.LXSsp
    	DataAdapter.Fill(BarcodeDT)
        MyCommon.Close_LXSsp()
		
	 Else
			' infoMessage = "Invalid barcode"
	 End If
	 
	If (Request.QueryString("export") <> "") Then
    ' they want to download the group get it from the database and stream it to the client
  
    If (BarcodeDT.Rows.Count > 0) Then
	   Dim time as DateTime = BarcodeDT.Rows(0).Item("BatchGeneratedOn")
      Response.AddHeader("Content-Disposition", "attachment; filename=BarcodeBatch." &  MyCommon.Leading_Zero_Fill(Year(time), 4) & _ 
		MyCommon.Leading_Zero_Fill(Month(time), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(time), 2) & ".txt")
		
      Response.ContentType = "application/octet-stream"
	  Send( Copient.PhraseLib.Lookup("term.memberid", LanguageID)  & "," & Copient.PhraseLib.Lookup("term.barcode", LanguageID) )
      For Each row In BarcodeDT.Rows
        Sendb(MyCryptlib.SQL_StringDecrypt(row.Item("InitialCardID").ToString()))
        Sendb(",")
        Send(MyCommon.NZ(row.Item("barcode"), ""))	
      Next
      GoTo done
    Else
      infoMessage = Copient.PhraseLib.Lookup("pgroup-edit.noelements", LanguageID)
    End If
  End If

 

  
  Send_HeadBegin()
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
  Send_HeadEnd()
  Send_BodyBegin(2)

  If (Logix.UserRoles.AccessBarcodeInquiry = False) Then
    Send_Denied(1, "perm.accessbarcodeinquiry")
    GoTo done
  End If
%>
<form action="#"  id="mainform" name="mainform">

  <div id="intro">
    <h1 id="title">

    </h1>
    <div id = "controls">

		<%
			Send_Export()
			 Send("<input type=""hidden"" id=""RequestPK"" name=""RequestPK"" value=""" & RequestPK & """ />")
			  %>

    </div>
  </div>
  
  <div id="main">
	 
	 <%If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>") 
	 %>
    <div id="coulmn">
	<%
		  Dim Shaded As String = "shaded"
		Send("<div class=""box"">")
        Send("<h2>")
        Send("	<span>" & Copient.PhraseLib.Lookup("term.barcode", LanguageID) & "</span>")
		
        Send("</h2>")

        Send("<table>")
		Send("<thead>")
		Send("<th>" & Copient.PhraseLib.Lookup("term.memberid", LanguageID) & "</th>")
		Send("<th>" & Copient.PhraseLib.Lookup("term.barcode", LanguageID) & "</th>")
		Send("</thead>")
        Send("	<tbody>")
		For each row in BarcodeDT.Rows
			Send("<tr class=""" & Shaded & """>")
			Send("<td>" & MyCryptlib.SQL_StringDecrypt(row.Item("InitialCardID").ToString()) & " </td>")
			Send("<td>" & row.Item("Barcode") & "</td>")
			Send("<tr>")
		If Shaded = "shaded" Then
            Shaded = ""
          Else
            Shaded = "shaded"
          End If
		Next
		Send("   </tbody>")
		Send(" </table>")
		Send("</div>")

	%>
    <br/>
   
    </div>
  </div>
</form>
<%

    Send_BodyEnd("mainform")
done:

  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  Logix = Nothing
  MyCommon = Nothing
%>
