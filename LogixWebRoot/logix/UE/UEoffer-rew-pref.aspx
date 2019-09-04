<%@ Page Title="" Language="C#" AutoEventWireup="true" CodeFile="UEoffer-rew-pref.aspx.cs"
  Inherits="logix_UE_UEoffer_rew_pref" %>

<%@ Reference Control="~/logix/UserControls/MultiLanguagePopup.ascx" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
  <base target="_self" />
  <title></title>
  <script type="text/javascript" src="/javascript/logix.js"></script>
  <script type="text/javascript" src="/javascript/jquery.min.js"></script>
</head>
<body class="popup">
  <script type="text/javascript" language="javascript">
    function CloseModel() {
      window.close();

    }
    function SendChangeOnServer(obj) {
      //use keyup instead keypress because:
      //- keypress will not work on backspace and delete
      //- keypress is called before the character is added to the textfield (at least in google chrome) 
      var searchText = $.trim(obj.value);
      var c = String.fromCharCode(event.keyCode);
      var isWordCharacter = c.match(/\w/);
      var isBackspaceOrDelete = (event.keyCode == 8 || event.keyCode == 46);
      var isAllowedBlankSpace = (event.keyCode == 32 && searchText.length > 0);

      if (isWordCharacter || isBackspaceOrDelete || isAllowedBlankSpace)
        document.getElementById('ReloadThePanel').click();
    }
    $(document).ready(function () {
      var object = $get('functioninput');
      if (object != null) {
        object.focus();
      }
    });

    window.onunload = function () { window.opener.location = '/logix/UE/UEoffer-rew.aspx?OfferID=<%= OfferID %>'; }
  </script>
  <form id="mainform" runat="server">
  <asp:ScriptManager ID="smScriptManager1" runat="server" ScriptMode="Auto" EnablePartialRendering="true"
    EnablePageMethods="true">
  </asp:ScriptManager>
  <div id="intro">
    <h1 id='title' runat="server">
      Title</h1>
    <div id='controls'>
      <span class="temp" id="TempDisallow" runat="server">
        <asp:CheckBox ID="chkDisallow_Edit" runat="server" CssClass="tempcheck" />
        <label for="Disallow_Edit">
          <%=PhraseLib.Lookup("term.locked", LanguageID)%>
        </label>
      </span>
      <asp:Button AccessKey="s" CssClass="regular" ID="btnSave" runat="server" Text=""
        Visible="true" OnClick="btnSave_Click" />
    </div>
  </div>
  <div id="main">
    <asp:UpdatePanel ID="UpdatePanelMain" runat="server" UpdateMode="Conditional">
      <ContentTemplate>
        <div id="infobar" clientidmode="Static" class="red-background" runat="server" visible="false">
        </div>
        <div id="column1">
          <div class="box" id="selector">
            <h2>
              <span>
                <%=PhraseLib.Lookup("term.preferencereward", LanguageID)%>
              </span>
            </h2>
            <asp:RadioButton runat="server" ID="functionradio1" GroupName="functionradio" Checked="true" /><label
              for="functionradio1"><%=PhraseLib.Lookup("term.startingwith", LanguageID)%></label>
            <asp:RadioButton runat="server" ID="functionradio2" GroupName="functionradio" /><label
              for="functionradio2"><%=PhraseLib.Lookup("term.containing", LanguageID)%></label><br />
            <asp:UpdatePanel ID="UpdatePanel1" runat="server" UpdateMode="Conditional">
              <ContentTemplate>
                <asp:TextBox runat="server" CssClass="medium" ID="functioninput" ClientIDMode="Static"
                  onkeyup="SendChangeOnServer(this);" MaxLength="100" AutoPostBack="false" style="width:197px"/><br />
              </ContentTemplate>
            </asp:UpdatePanel>
            <asp:UpdatePanel ID="UpdatePanel2" runat="server" UpdateMode="Conditional">
              <ContentTemplate>
                <asp:Button ID="ReloadThePanel" runat="server" Style="display: none;" ClientIDMode="Static"
                  OnClick="ReloadThePanel_Click" />
                <div id="cgList">
                  <asp:ListBox ID="lstAvailable" runat="server" SelectionMode="Single" DataTextField="PhraseText"
                    DataValueField="PreferenceID" Rows="12" CssClass="longer"></asp:ListBox>
                </div>
                <br />
                <b>
                  <%= PhraseLib.Lookup("term.selectedpreference", LanguageID)%>:</b><br />
                <asp:Button runat="server" class="regular" ID="select1" Text="" OnClick="select1_Click">
                </asp:Button>
                <br class="half" />
                <asp:ListBox ID="lstSelected" runat="server" SelectionMode="Single" Rows="2" CssClass="longer"
                  DataTextField="PhraseText" DataValueField="PreferenceID"></asp:ListBox>
                <br />
                <br />
                <%=PhraseLib.Lookup("term.datatype", LanguageID)%>:
                <asp:Label ID="lblDataType" runat="server" Text=""></asp:Label>
                <br />
                <%=PhraseLib.Lookup("term.multiple-values", LanguageID)%>:
                <asp:Label ID="lblMultiValued" runat="server" Text=""></asp:Label>
              </ContentTemplate>
              <Triggers>
                <asp:AsyncPostBackTrigger ControlID="ReloadThePanel" EventName="Click" />
              </Triggers>
            </asp:UpdatePanel>
          </div>
        </div>
        <div id="gutter">
        </div>
        <div id="column2" runat="server" clientidmode="Static">
          <div class="box" id="preferencevalues">
            <h2>
              <span>
                <%=PhraseLib.Lookup("term.values", LanguageID)%>
              </span>
            </h2><br />
            <asp:UpdatePanel ID="UpdatePanel3" runat="server">
              <ContentTemplate>
                <asp:Repeater ID="rptTierValues" runat="server" OnItemDataBound="rptTierValues_ItemDataBound"
                  OnItemCommand="rptTierValues_ItemCommand">
                  <ItemTemplate>
                    <asp:Label ID="lblTierText" runat="server" Text=""></asp:Label>
                    <%= PhraseLib.Lookup("term.value", LanguageID) + ":" %>&nbsp;&nbsp;
                    <asp:DropDownList ID="ddlPreferenceData" runat="server" DataTextField="PhraseText"
                      DataValueField="Value" CssClass="mediumlong">
                    </asp:DropDownList>
                    <asp:Button ID="btnAdd" runat="server" width="60px" Text="" CommandArgument='<%# Eval("TierLevel") %>'
                      CommandName="Add" /><br />
                    <br class="half" />
                    <asp:ListBox ID="lstSelectedPreference" runat="server" SelectionMode="Multiple" DataTextField="PhraseText"
                      DataValueField="Value" CssClass="longer"></asp:ListBox>
                    <br />
                    <br class="half" />
                    <asp:Button ID="btnRemove" runat="server" Text="" CommandArgument='<%# Eval("TierLevel") %>'
                      CommandName="Remove" /><br />
                  </ItemTemplate>
                  <SeparatorTemplate>
                    <br />
                    <br />
                  </SeparatorTemplate>
                </asp:Repeater>
              </ContentTemplate>
            </asp:UpdatePanel>
          </div>
        </div>
      </ContentTemplate>
    </asp:UpdatePanel>
  </div>
  </form>
</body>
</html>
