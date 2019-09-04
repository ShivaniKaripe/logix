<%@ Page Title="" Language="C#" MasterPageFile="~/logix/User.master" AutoEventWireup="true" CodeFile="PageDenied.aspx.cs" Inherits="logix_PageDenied" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
<div id="intro">
<h1 id="title"><%=PhraseLib.Lookup("term.accessdenied", LanguageID)%> </h1>

</div>
<div id="main">
 <%=PhraseLib.Lookup("error.forbidden", LanguageID)%>
 <asp:Label ID="lblError" runat="server">
 </asp:Label>
       
</div>
</asp:Content>

