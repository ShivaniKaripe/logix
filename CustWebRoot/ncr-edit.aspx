<%@ Page Language="vb" Debug="true" CodeFile="ncr-cwCB.vb" Inherits="cwCB" %>
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
  Dim FirstName As String = ""
  Dim LastName As String = ""
  Dim Address As String = ""
  Dim City As String = ""
  Dim State As String = ""
  Dim Zip As String = ""
  Dim Phone As String = ""
  Dim Email As String = ""
  Dim DOB_month As String = ""
  Dim DOB_day As String = ""
  Dim DOB_year As String = ""
  Dim DOB As String = ""
  
  Dim TempValue As String = ""
  Dim DOBParts() As String = {"", "", ""}
  
  Response.Expires = 0
  MyCommon.AppName = "ncr-edit.aspx"
  
  MyCommon.Open_LogixXS()
  
  If (Request.QueryString("submit") = "Submit") Then
    FirstName = MyCommon.Parse_Quotes(Request.QueryString("firstname"))
    LastName = MyCommon.Parse_Quotes(Request.QueryString("lastname"))
    Address = MyCommon.Parse_Quotes(Request.QueryString("address"))
    City = MyCommon.Parse_Quotes(Request.QueryString("city"))
    State = MyCommon.Parse_Quotes(Request.QueryString("state"))
    Zip = MyCommon.Parse_Quotes(Request.QueryString("zip"))
    Email = MyCommon.Parse_Quotes(Request.QueryString("email"))
    Phone = MyCommon.Parse_Quotes(Request.QueryString("phone1"))
    DOB_month = Request.QueryString("dob1")
    DOB_day = Request.QueryString("dob2")
    DOB_year = Request.QueryString("dob3")
    If (DOB_month.Trim = "" AndAlso DOB_day.Trim = "" AndAlso DOB_year.Trim = "") Then
      DOB = "NULL"
    Else
      DOB = DOB_month.Trim.PadLeft(2, "0") & DOB_day.Trim.PadLeft(2, "0") & DOB_year.Trim.PadLeft(4, "0")
    End If
    
    MyCommon.QueryStr = "update Customers with (RowLock) set " & _
                        "FirstName='" & FirstName & "', " & _
                        "LastName='" & LastName & "', " & _
                        "CPEStoreSendFlag=1 " & _
                        "where CustomerPK=" & CustomerPK
    MyCommon.LXS_Execute()
    MyCommon.QueryStr = "update CustomerExt with (RowLock) set " & _
                        "Address='" & Address & "', " & _
                        "City='" & City & "', " & _
                        "State='" & State & "', " & _
                        "Zip='" & Zip & "', " & _
                        "PhoneAsEntered='" & Phone & "', " & _
                        "PhoneDigitsOnly='" & MyCommon.DigitsOnly(Phone) & "', " & _
                        "Email='" & Email & "', " & _
                        "DOB='" & DOB & "' " & _
                        "where CustomerPK=" & CustomerPK
    MyCommon.LXS_Execute()
  End If
%>
<!-- IE6 quirks mode -->
<!DOCTYPE html 
     PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
     "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="en" xml:lang="en">
<head>
<title>Marsh Customer Website</title>
<meta name="copyright" content="&copy; Copyright 2007, Marsh Supermarkets, Inc." />
<meta name="description" content="Marsh Supermarkets customer-facing website" />
<meta name="content-type" content="text/html; charset=utf-8" />
<meta name="robots" content="noindex, nofollow" />
<meta http-equiv="cache-control" content="no-cache" />
<meta http-equiv="pragma" content="no-cache" />
<link rel="icon" href="images/favicon.ico" type="image/x-icon" />
<link rel="shortcut icon" href="images/favicon.ico" type="image/x-icon" />
<link rel="stylesheet" href="css/ncr-cw-screen.css" type="text/css" media="screen" />
<script src="javascript/cw.js" type="text/javascript"></script>
</head>

<body class="popup" onunload="RefreshParent()">
<script type="text/javascript">
function RefreshParent() {
    window.close();
    opener.location.reload();
} 
</script>
<div id="wrap">
<a id="top" name="top"></a>

<%
  MyCommon.QueryStr = "select * from Customers where CustomerPK=" & CustomerPK
  rst = MyCommon.LXS_Select
  For Each row In rst.Rows
    FirstName = MyCommon.NZ(row.Item("FirstName"), "")
    LastName = MyCommon.NZ(row.Item("LastName"), "")
  Next
  MyCommon.QueryStr = "select * from CustomerExt where CustomerPK=" & CustomerPK
  rst = MyCommon.LXS_Select
  For Each row In rst.Rows
    Address = MyCommon.NZ(row.Item("Address"), "")
    City = MyCommon.NZ(row.Item("City"), "")
    State = MyCommon.NZ(row.Item("State"), "")
    Zip = MyCommon.NZ(row.Item("Zip"), "")
    Email = MyCommon.NZ(row.Item("Email"), "")
    Phone = MyCommon.NZ(row.Item("PhoneAsEntered"), "")
    DOB = MyCommon.NZ(row.Item("DOB"), "")
  Next
%>

<h2>Edit your details</h2>
<form action="#" id="editform" name="editform" action="get">
  <input type="hidden" name="CustPK" id="CustPK" value="<% Sendb(CustomerPK) %>" />
  <table summary="Details">
    <tbody>
      <tr>
        <td>
          <label for="firstname">First name:</label>
        </td>
        <td colspan="3">
          <input type="text" id="firstname" name="firstname" value="<% Sendb(FirstName) %>" />
        </td>
      </tr>
      <tr>
        <td>
          <label for="lastname">Last name:</label>
        </td>
        <td colspan="3">
          <input type="text" id="lastname" name="lastname" value="<% Sendb(LastName) %>" />
        </td>
      </tr>
      <tr>
        <td>
          <label for="address">Address:</label>
        </td>
        <td colspan="3">
          <input type="text" id="address" name="address" value="<% Sendb(Address) %>" />
        </td>
      </tr>
      <tr>
        <td>
          <label for="city">City:</label>
        </td>
        <td colspan="3">
          <input type="text" id="city" name="city" value="<% Sendb(City) %>" />
        </td>
      </tr>
      <tr>
        <td>
          <label for="state">State:</label>
        </td>
        <td>
          <input type="text" class="shortest" id="state" name="state" value="<% Sendb(State) %>" />
        </td>
        <td>
          <label for="zip">ZIP:</label>
        </td>
        <td>
          <input type="text" class="short" id="zip" name="zip" value="<% Sendb(Zip) %>" />
        </td>
      </tr>
      <tr>
        <td>
          <label for="email">Email:</label>
        </td>
        <td colspan="3">
          <input type="text" id="email" name="email" value="<% Sendb(Email) %>" />
        </td>
      </tr>
      <tr>
        <td>
          <label for="phone1">Phone:</label>
        </td>
        <td colspan="3">
          <%
            Send("<input type=""text"" id=""phone1"" name=""phone1"" maxlength=""50"" value=""" & Phone & """ />")
          %>
        </td>
      </tr>
      <tr>
        <td>
          <label for="dob1">Birth date:</label>
        </td>
        <td colspan="3">
        <%
          TempValue = DOB
          DOBParts = ParseDateOfBirth(TempValue)
          Send("    <input type=""text"" style=""width:20px;"" id=""dob1"" name=""dob1"" maxlength=""2"" value=""" & DOBParts(0) & """ />/")
          Send("    <input type=""text"" style=""width:20px;"" id=""dob2"" name=""dob2"" maxlength=""2"" value=""" & DOBParts(1) & """ />/")
          Send("    <input type=""text"" style=""width:38px;"" id=""dob3"" name=""dob3"" maxlength=""4"" value=""" & DOBParts(2) & """ />")
        %>
          (m/d/y)          
        </td>
      </tr>
      <tr>
        <td colspan="4" align="center">
          <input type="submit" class="medium" id="submit" name="submit" value="Submit" />
        </td>
      </tr>
    </tbody>
  </table>
</form>

<a id="bottom" name="bottom"></a>
</div>
</body>
</html>

<script runat="server">
  
  Function ParseDateOfBirth(ByVal DateOfBirth As String) As String()
    Dim DOBParts() As String = {"", "", ""}
    If (DateOfBirth IsNot Nothing) Then
      Select Case DateOfBirth.Length
        Case 4
          DOBParts(0) = ""
          DOBParts(1) = ""
          DOBParts(2) = DateOfBirth
        Case 8
          DOBParts(0) = DateOfBirth.Substring(0, 2)
          DOBParts(1) = DateOfBirth.Substring(2, 2)
          DOBParts(2) = DateOfBirth.Substring(4)
      End Select
    End If
    Return DOBParts
  End Function
</script>

<%
  MyCommon.Close_LogixXS()
  MyCommon = Nothing
  Logix = Nothing
%>
