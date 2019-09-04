<%@ Page Title="" Language="C#" MasterPageFile="~/logix/User.master" AutoEventWireup="true" CodeFile="error-message.aspx.cs" Inherits="logix_error_message" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
<div id="intro">
  <h1 id="title">
  <%=Server.HtmlEncode(MainHeading)%>
  </h1>
</div>
<div id="main">
  <div id="infobar" class="red-background">
   <%=Server.HtmlEncode(ErrorMessage)%>
  </div>
</div>
</asp:Content>

