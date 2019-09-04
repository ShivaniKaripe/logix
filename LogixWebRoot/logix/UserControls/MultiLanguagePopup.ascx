<%@ Control Language="C#" AutoEventWireup="true" Debug = "true"  CodeFile="MultiLanguagePopup.ascx.cs" Inherits="logix_UserControls_MultiLanguagePopup" ClassName="logix_UserControls_MultiLanguagePopup" %>
<div class="mlwrap" id="divMLWrap" style="z-index: 500;" runat="server">
    <asp:TextBox runat="server" id="tbMLI" name="tbMLIStandard" maxlength="1000" class="middle" 
        
        onkeydown="return (event.keyCode!=13);" 
        value="" />
    <asp:Image runat="server" ImageUrl="/images/mg.png" class="mg" id="imgMLI" 
        title="" />
    <div runat="server" class="ml" id="divML" style="display:none; opacity: 1; margin-top:-2em;" >
        <asp:Button runat="server" id="btnMLClose" class="mlclose" value="X" 
        title="Close" />
        <br class="half" />
                    <asp:Repeater ID="repMLIInputs" runat="server" OnItemDataBound="repMLIInputs_ItemDataBound">
                        <ItemTemplate>
                            <asp:HiddenField ID="hfLangId" runat="server" Value='<%# DataBinder.Eval(Container.DataItem, "LanguageID")%>' />
                            <asp:Label ID="lblLanguageName" runat="server" Text='<%# DataBinder.Eval(Container.DataItem, "Name")%>' />
                            <br />
                            <asp:TextBox runat="server" id="tbTranslation"  
                                onkeydown="return (event.keyCode!=13);" value='<%# DataBinder.Eval(Container.DataItem, "Translation")%>'
                              />
                        </ItemTemplate>
                        <SeparatorTemplate>
                            <br />
                        </SeparatorTemplate>
                    </asp:Repeater>
        <br />
    </div>
    <br />
    <br class="half" />
</div>

