<%@ Page Language="vb" Debug="true" CodeFile="cwCB.vb" Inherits="cwCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  Dim MyCommon As New Copient.CommonInc
  Dim MyCryptLib As New Copient.CryptLib
  Dim Logix As New Copient.LogixInc
  Dim rst As System.Data.DataTable
  Dim row As System.Data.DataRow
  Dim CustomerPK As Long = Request.QueryString("CustPK")
  Dim PrimaryExtID As Long = Request.QueryString("PrimaryExtID")
  
  Dim TempValue As String = ""
  
  Response.Expires = 0
  MyCommon.AppName = "list.aspx"
  
  MyCommon.Open_LogixXS()
  
  If (Request.QueryString("submit") = "Submit") Then
    MyCommon.QueryStr = "insert into ShopList (CustomerPK,Item) values (" & CustomerPK & ",'" & Request.QueryString("item") & "')"
    MyCommon.LXS_Execute()
    ' store the change away for later reporting by the customer inquiry web service to sync up outside customer care applications.
  ElseIf ( Request.QueryString("listpk") <> "" ) Then
    ' they want to delete a list item.
    MyCommon.QueryStr = "delete from ShopList where listpk=" & Request.QueryString("listpk")
    MyCommon.LXS_Execute()
  End If
%>
<!-- IE6 quirks mode -->
<!DOCTYPE html 
     PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
     "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="en" xml:lang="en">
<head>
<title>NCR Customer Website</title>
<meta name="copyright" content="&copy; Copyright 2011, NCR" />
<meta name="description" content="NCR customer-facing website" />
<meta name="content-type" content="text/html; charset=utf-8" />
<meta name="robots" content="noindex, nofollow" />
<meta http-equiv="cache-control" content="no-cache" />
<meta http-equiv="pragma" content="no-cache" />
<link rel="icon" href="images/favicon.ico" type="image/x-icon" />
<link rel="shortcut icon" href="images/favicon.ico" type="image/x-icon" />
<link rel="stylesheet" href="css/cw-screen.css" type="text/css" media="screen" />
<script src="javascript/cw.js" type="text/javascript"></script>
</head>

<body class="popup" >
<script type="text/javascript">
function RefreshParent() {
    window.close();
   // opener.location.reload();
} 
</script>
<div id="wrap">
<a id="top" name="top"></a>


<h2>Edit your list</h2>
<form action="#" id="editform" name="editform">
  <input type="hidden" name="CustPK" id="CustPK" value="<% Sendb(CustomerPK) %>" />
  <table summary="Details" >
    <tbody>
	    <tr>
        <td>New Item:
          <input type="text" id="item" name="item" value="" />
        </td>
      </tr>
      <tr>
        <td align="center">
          <input type="submit" class="medium" id="submit" name="submit" value="Submit" />
        </td>
      </tr>
<%
  MyCommon.QueryStr = "select listpk,Item from ShopList where CustomerPK=" & CustomerPK
  rst = MyCommon.LXS_Select
  For Each row In rst.Rows
    Sendb("<tr><td align=""left"" bgcolor=#EEE >")
    Sendb("<a href='list.aspx?CustPK=" & CustomerPK & "&listpk=" & row.Item("listpk") & "'>X</a> ")
    Sendb(row.Item("Item"))
    Sendb("</tr></td")
  Next
%>
    </tbody>
  </table>
</form>

<a id="bottom" name="bottom"></a>
</div>
</body>
</html>

<%
  MyCommon.Close_LogixXS()
  MyCommon = Nothing
  Logix = Nothing
%>