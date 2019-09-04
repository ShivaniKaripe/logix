<%@ Page Language="C#" AutoEventWireup="true" CodeFile="PopUpDenied.aspx.cs" Inherits="logix_Denied" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body class="popup" >
    <form id="form1" runat="server">
  <div id="custom1"></div>
<div id="wrap">
<div id="custom2"></div>
<a id="top" name="top"></a>
<div id="intro">
<h1 id="title"><%=PhraseLib.Lookup("term.accessdenied", LanguageID)%> </h1>

</div>
<div id="main">
 <%=PhraseLib.Lookup("error.forbidden", LanguageID)%>
 <asp:Label ID="lblError" runat="server">
 </asp:Label>
       
</div>
  <a id="bottom" name="bottom"></a>
<div id="footer">
   <%=PhraseLib.Lookup("about.copyright", LanguageID)%>
</div>
<div id="custom3"></div>
</div> <!-- End wrap -->
<div id="custom4"></div>
    </form>
</body>
</html>
