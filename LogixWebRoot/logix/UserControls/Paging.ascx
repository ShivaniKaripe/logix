<%@ Control Language="C#" AutoEventWireup="true" CodeFile="Paging.ascx.cs" Inherits="logix_UserControls_Paging" %>
<div id="paginator">
   <span id="first">
   <asp:LinkButton ID="lnkFirst"  Text="<b>|</b>◄First" runat="server" 
     onclick="lnkFirst_Click"></asp:LinkButton>
   <asp:Label ID="lblFirst" Text="<b>|</b>◄First" runat="server"></asp:Label>
   </span>&nbsp;
   <span id="previous">
   <asp:LinkButton ID="lnkPrevious" Visible="false" Text="◄Previous" runat="server" 
     onclick="lnkPrevious_Click"></asp:LinkButton><asp:Label ID="lblPrevious" Text="◄Previous" runat="server">
     </asp:Label>
     </span>&nbsp;[
   <asp:Label ID="lblPage" Text="" runat="server"></asp:Label>
   ]&nbsp;
   <span id="next">
    <asp:LinkButton ID="lnkNext" Visible="false" Text="Next►" runat="server" 
     onclick="lnkNext_Click"></asp:LinkButton>
   <asp:Label ID="lblNext" Text="Next►" runat="server"></asp:Label>
   </span>&nbsp;
   <span id="last">
      <asp:LinkButton ID="lnkLast" Text="Last►<b>|</b>" runat="server" 
     onclick="lnkLast_Click"></asp:LinkButton>
   <asp:Label ID="lblLast" Text="Last►<b>|</b>" runat="server"></asp:Label>
   </span><br />
  </div>