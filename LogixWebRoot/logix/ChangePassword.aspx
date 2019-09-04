<%@ Page Title="" Language="C#" Debug="true" AutoEventWireup="true" CodeFile="ChangePassword.aspx.cs" Inherits="ChangePassword"  ValidateRequest="false"  %>
<head runat="server">
    <link rel="stylesheet" href="../css/logix-screen.css" />
</head>
<body>
<div id="custom1"></div>
<div id="wrap">
<div id="custom2"></div>
<form id="mainform" runat="server">
   <div id="logos">
  <div id="logix" title='<%=PhraseLib.Lookup("term.logix", LanguageID) %>'></div>
  <%
    if (System.IO.File.Exists(Server.MapPath("/images/logos/licenseecustom.png")))
    { %>
  <div id="licenseecustom" title='<%= PhraseLib.Lookup("term.licensee",LanguageID) %>'></div>
  <%}
    else
    { %>
  <div id="licensee" title='<%=PhraseLib.Lookup("term.licensee", LanguageID)%>'></div>
    <%} %>
  <br clear="all" />
</div> 
    <div id="tabs">
<asp:PlaceHolder runat="server" ID="phMenu" ></asp:PlaceHolder>

  <br clear="all" />
</div>
    <div id="intro">
</div>
    <div id="main">
         <h1 id="title">
    <span>
        <% =PhraseLib.Lookup("term.change", LanguageID) %> <% =PhraseLib.Lookup("term.password", LanguageID) %>
        </span>
  </h1>
   <div id="infobar"  runat="server" clientidmode="Static" visible="false" />
<asp:Panel ID="ChangePassword1" runat="server" CssClass="MinutesList" >
        <table cellpadding="1" cellspacing="0" style="border-collapse:collapse;">
            <tr>
                <td>
                    <table cellpadding="0">
                      <tr>
                            <td align="right">
                                <asp:Label ID="CurrentPasswordLabel" runat="server" 
                                    AssociatedControlID="CurrentPassword"></asp:Label>
                            </td>
                            <td>
                                <asp:TextBox ID="CurrentPassword" runat="server" TextMode="Password" Visible="true" >></asp:TextBox>
                                <asp:RequiredFieldValidator ID="CurrentPasswordRequired" runat="server" 
                                    ControlToValidate="CurrentPassword"  
                                    ToolTip="Password is required." ValidationGroup="ChangePassword1"></asp:RequiredFieldValidator>
                            </td>
                        </tr>
                        <tr>
                            <td align="right">
                                <asp:Label ID="NewPasswordLabel" runat="server" 
                                    AssociatedControlID="NewPassword"></asp:Label>
                            </td>
                            <td>
                                <asp:TextBox ID="NewPassword" runat="server" TextMode="Password"></asp:TextBox>
                                <asp:RequiredFieldValidator ID="NewPasswordRequired" runat="server" 
                                    ControlToValidate="NewPassword"  
                                    ToolTip="New Password is required." ValidationGroup="ChangePassword1"></asp:RequiredFieldValidator>
                            </td>
                        </tr>
                        <tr>
                            <td align="right">
                                <asp:Label ID="ConfirmNewPasswordLabel" runat="server" 
                                    AssociatedControlID="ConfirmNewPassword"></asp:Label>
                            </td>
                            <td>
                                <asp:TextBox ID="ConfirmNewPassword" runat="server" TextMode="Password"></asp:TextBox>
                                <asp:RequiredFieldValidator ID="ConfirmNewPasswordRequired" runat="server" 
                                    ControlToValidate="ConfirmNewPassword"                                     
                                    ToolTip="Confirm New Password is required." ValidationGroup="ChangePassword1"></asp:RequiredFieldValidator>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="2">
                                <asp:CompareValidator ID="NewPasswordCompare" runat="server" 
                                    ControlToCompare="NewPassword" ControlToValidate="ConfirmNewPassword" 
                                    Display="Dynamic"                                     
                                    ValidationGroup="ChangePassword1" EnableClientScript="false"></asp:CompareValidator>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="2" style="color:Red;">
                                <asp:Literal ID="FailureText" runat="server" EnableViewState="False"></asp:Literal>
                            </td>
                        </tr>
                        <tr>
                            <td align="right">
                                <asp:Button ID="ChangePasswordPushButton" runat="server" 
                                    CommandName="ChangePassword" 
                                    ValidationGroup="ChangePassword1" OnClick="ChangePasswordPushButton_Click" />
                            </td>
                            <td>
                                <asp:Button ID="CancelPushButton" runat="server" CausesValidation="False" 
                                    CommandName="Cancel" />
                             <asp:Button ID="Continue" runat="server"  Visible="false" CausesValidation="true" 
                                    onclick="HomePageButton_Click" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>

</asp:Panel>
        </div>
    </form>
    </body>