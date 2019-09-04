<%@ Control Language="C#" AutoEventWireup="true"  CodeFile="CollapsableDiv.ascx.cs" Inherits="logix_UserControls_CollapsableDiv" %>
 <div class="resizer">
  <a href="javascript:resizeDiv('<%=TargetDivID%>','<%=imgID.ClientID%>','<%=ToolTip%>');">
  <asp:Image runat="server" ID="imgID" ImageUrl ="/images/arrowup-off.png"  />
  </a>
  </div>