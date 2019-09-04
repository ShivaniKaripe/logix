<%@ Page Title="" Language="C#" MasterPageFile="~/logix/PopUpMaster.master" AutoEventWireup="true"
  CodeFile="OfferTCProgramCondition.aspx.cs" Inherits="logix_OfferTCProgramCondition" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
  <base target="_self" />
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

    window.onunload = function () { window.opener.location = '/logix/UE/UEoffer-con.aspx?OfferID=<%= OfferID %>'; }
  </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
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
                <%=PhraseLib.Lookup("term.trackablecouponprogram", LanguageID)%>
              </span>
            </h2>
            <asp:RadioButton runat="server" ID="functionradio1" GroupName="functionradio" Checked="true" /><label
              for="functionradio1"><%=PhraseLib.Lookup("term.startingwith", LanguageID)%></label>
            <asp:RadioButton runat="server" ID="functionradio2" GroupName="functionradio" /><label
              for="functionradio2"><%=PhraseLib.Lookup("term.containing", LanguageID)%></label><br />
            <asp:UpdatePanel ID="UpdatePanel1" runat="server" UpdateMode="Conditional">
              <ContentTemplate>
                <asp:TextBox runat="server" CssClass="medium" ID="functioninput" ClientIDMode="Static"
                  onkeyup="SendChangeOnServer(this);" MaxLength="100" AutoPostBack="false" /><br />
              </ContentTemplate>
            </asp:UpdatePanel>
            <asp:UpdatePanel ID="UpdatePanel2" runat="server" UpdateMode="Conditional">
              <ContentTemplate>
                <asp:Button ID="ReloadThePanel" runat="server" Style="display: none;" ClientIDMode="Static"
                  OnClick="ReloadThePanel_Click" />
                <div id="cgList">
                  <asp:ListBox ID="lstAvailable" runat="server" SelectionMode="Single" DataTextField="Name"
                    DataValueField="ProgramID" Rows="12" CssClass="longer"></asp:ListBox>
                </div>
                <br />
                <asp:Button runat="server" class="regular" ID="select1" Text="" OnClick="select1_Click">
                </asp:Button>
                <asp:Button runat="server" class="regular" ID="deselect1" Text="" Enabled="false"
                  OnClick="deselect1_Click" /><br />
                <br class="half" />
                <asp:ListBox ID="lstSelected" runat="server" SelectionMode="Single" Rows="2" CssClass="longer"
                  DataTextField="Name" DataValueField="ProgramID"></asp:ListBox>
                <br />
              </ContentTemplate>
              <Triggers>
                <asp:AsyncPostBackTrigger ControlID="ReloadThePanel" EventName="Click" />
              </Triggers>
            </asp:UpdatePanel>
          </div>
        </div>
        <div id="gutter">
        </div>
        <%--        <div id="column2" runat="server">
          <div class="box" id="hhoptions">
            <h2>
              <span>
                <%=PhraseLib.Lookup("term.value", LanguageID)%>
              </span>
            </h2>
            <tr>
              <td>
                <asp:Label ID="lblValueNeeded" runat="server"></asp:Label>
              </td>
            </tr>
            <tr>
              <td>
                <asp:TextBox runat="server" ID="txtValueNeeded"></asp:TextBox>
              </td>
            </tr>
            <table>
            </table>
          </div>
        </div>--%>
      </ContentTemplate>
    </asp:UpdatePanel>
  </div>
</asp:Content>
