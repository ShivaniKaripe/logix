<%@ Control Language="C#" AutoEventWireup="true" CodeFile="TemplateFieldLockControl.ascx.cs" Inherits="logix_UserControls_TemplateFieldLockControl" %>

<span id="TempDisallow" runat="server" class="newTemp">
    <asp:CheckBox ID="chkDisallow_Edit" runat="server" CssClass="tempcheck" OnCheckedChanged="chkDisallow_Edit_OnCheckedChanged" AutoPostBack="true" />
    <asp:label id="lblLocked" runat="server"><%= PhraseLib.Lookup("term.locked", LanguageId)%></asp:label>
    <asp:LinkButton ID="lnkBtnArrow" title='<%= PhraseLib.Lookup("cpeoffer-rew-disc-clicktoview", LanguageId)%>' 
    runat="server" OnClick="lnkBtnArrow_OnClick">
    &#9660;</asp:LinkButton>
</span>
<div id="divTemplatefields" class="newTemplatefields" runat="server" >
    <asp:Label ID="lblFieldLevelPerms" runat="server" Font-Bold="true"> <%= PhraseLib.Lookup("temp.fieldlevelperms", LanguageId) %> </asp:Label>
    </br>
    <asp:Repeater ID="repTemplateFields" runat="server" OnItemDataBound="repTemplateFields_OnItemDataBound">
       <HeaderTemplate>
         <table>
       </HeaderTemplate>
       <ItemTemplate>
        <tr style="background-color:#e0e0e0;">
            <td style="width:10%;">
            <asp:CheckBox ID="chkFieldLock" runat="server" Checked="false"
            OnCheckedChanged="chkFieldLock_OnCheckedChanged" AutoPostBack="true" />
            </td>
            <td style="width:60%;">
            <asp:Label ID="lblFieldName" runat="server" Text='<%# DataBinder.Eval(Container.DataItem, "FieldName") %>'/>
            </td>
            <td style="width:30%;">
            <asp:Label Id="lblStatus" runat="server" Text='<%# DataBinder.Eval(Container.DataItem, "StatusString") %>' />
            </td>
        </tr>
       </ItemTemplate>     
       <FooterTemplate>
          </table>
       </FooterTemplate>         
    </asp:Repeater>
</div>


