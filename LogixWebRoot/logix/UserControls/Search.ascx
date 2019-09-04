<%@ Control Language="C#" AutoEventWireup="true" CodeFile="Search.ascx.cs" Inherits="logix_UserControls_Search" %>
 <div id="searcher" title="Search terms"> 
  <asp:TextBox id="txtSearch" CssClass="searchterms" runat="server"  MaxLength="100"></asp:TextBox>
  <asp:Button ID="btnSearch" CssClass="searcher" runat="server" Text="Search" onclick="btnSearch_Click" />
  <br />
  </div>