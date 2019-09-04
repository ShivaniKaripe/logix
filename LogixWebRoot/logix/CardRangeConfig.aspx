<%@ Page Title="" Language="C#" MasterPageFile="~/logix/User.master" AutoEventWireup="true" CodeFile="CardRangeConfig.aspx.cs" Inherits="logix_CardRangeConfig" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
    <script type="text/javascript" src="../javascript/scrollsaver.min.js"></script>
    <script type="text/javascript">
        function clearfields() {
            document.getElementById('txtStartrange').value = '';
            document.getElementById('txtEndRange').value = '';
            return false;
        }
        function isNumber(evt) {
            evt = (evt) ? evt : window.event;
            var charCode = (evt.which) ? evt.which : evt.keyCode;
            if (charCode > 31 && (charCode < 48 || charCode > 57)) {
                return false;
            }
            return true;
        }
        function stripInvalidCharacters(control) {
            var $this = $(control);
            $this.val($this.val().replace(/\D/g, ''));
        }

        function Validation() {
            var blnvalid = true;
            var startrange = $('#txtStartrange').val();
            var endRange = $('#txtEndRange').val();
            if (startrange =="") {
                $('#infobar').html('<%=PhraseLib.Lookup(8769, LanguageID)%>');
                $('#infobar').show();
                $('#statusbar').hide();
                blnvalid = false;
                return blnvalid;
            }
            if (startrange == "0") {
                $('#infobar').html('<%=PhraseLib.Lookup(8794, LanguageID)%>');
                $('#infobar').show();
                $('#statusbar').hide();
                blnvalid = false;
                return blnvalid;
            }
            if (endRange == "") {
                $('#infobar').html('<%=PhraseLib.Lookup(8770, LanguageID)%>');
                $('#infobar').show();
                $('#statusbar').hide();
                blnvalid = false;
                return blnvalid;
            }
            if (endRange == "0") {
                $('#infobar').html('<%=PhraseLib.Lookup(8795, LanguageID)%>');
                $('#infobar').show();
                $('#statusbar').hide();
                blnvalid = false;
                return blnvalid;
            }
        }
    </script>
  
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <div id="intro">
        <h1 id="htitle" runat="server"></h1>
    </div>

    <div id="main">
        <div id="infobar" class="red-background" runat="server" clientidmode="Static" style="display:none;">
        </div>
        <div id="statusbar" class="green-background" runat="server" clientidmode="Static" style="display:none;">
        </div>
        <div class="columnfull">
            <div id="listbar" style="overflow: hidden;">
                <div class="customsearch">
                    <asp:DropDownList ID="ddlCardTypes" AutoPostBack="False" runat="server" ClientIDMode="Static" Style="height: 100%;max-width:150px;">
                    </asp:DropDownList>&nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:Label ID="lbStartRange" runat="server" Text="Start Range"></asp:Label>
                    <asp:TextBox ID="txtStartrange" CssClass="searchterms" style="height:19px; width:100px;" runat="server" ClientIDMode="Static" onkeyup="stripInvalidCharacters(this)" onkeypress="return isNumber(event)" MaxLength="28"></asp:TextBox>&nbsp;&nbsp;&nbsp;
                    <asp:Label ID="lbEndRange" runat="server" Text="End Range"></asp:Label>
                    <asp:TextBox ID="txtEndRange" CssClass="searchterms" style="height:19px; width:100px;" runat="server" ClientIDMode="Static" onkeyup="stripInvalidCharacters(this)" onkeypress="return isNumber(event)" MaxLength="28"></asp:TextBox>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:Button ID="btnAdd" runat="server" Text="Add" OnClick="btnAdd_Click" OnClientClick="return Validation()" style="height:100%"/>&nbsp;&nbsp;
                    <asp:Button ID="btnClear" OnClientClick="return clearfields()" runat="server" Text="Clear" style="height:100%"/>
                </div>
            </div>
            <div style="overflow: hidden;">
                <br />
                <asp:Repeater ID="repCardType" runat="server" OnItemDataBound="repCardType_ItemDataBound">
                    <ItemTemplate>
                        <div class="box">
                            <span>
                                <h2>
                                    <asp:Label ID="lblCardType" runat="server" Text='<%# PhraseLib.Lookup(Convert.ToInt32(Eval("PhraseID")), LanguageID) %>'  /></h2>
                            </span>
                            <asp:HiddenField ID="hdnCardTypeID" runat="server" Value='<%#Eval("CardTypeID") %>' />
                            <section class="cardrangesection">
                            <div class="container">
                            <asp:Repeater ID="repRangeList" runat="server" OnItemDataBound="repRangeList_ItemDataBound" OnItemCommand="repRangeList_ItemCommand">
                                <HeaderTemplate>
                                    <table class="tblcardrange">
                                        <tr>
                                            <th></th>
                                            <th>
                                                <b><asp:Label ID="lblStartRange" runat="server"/></b>
                                            </th>
                                            <th>
                                                <b><asp:Label ID="lblEndRange" runat="server" /></b>
                                            </th>
                                        </tr>
                                        <tr><th colspan="3"><hr /></th></tr>
                                       
                                </HeaderTemplate>
                                <ItemTemplate>
                                    <tr>
                                        <td>
                                            <asp:Button ID="btnRemove" runat="server" Text="X" class="ex" Font-Bold="true" CommandName="Delete" CommandArgument='<%# Eval("CardRangeID ") %>' />
                                            <asp:HiddenField ID="hdnCardRangeID" runat="server" Value='<%#Eval("CardRangeID") %>' />
                                        </td>
                                        <td>
                                            <asp:Label ID="lblMinValue" runat="server" Text='<%#Eval("StartRange") %>' />
                                        </td>
                                        <td>
                                            <asp:Label ID="lblMaxValue" runat="server" Text='<%#Eval("EndRange") %>' />
                                        </td>
                                    </tr>
                                </ItemTemplate>
                                <FooterTemplate>
                                    </table>
                                </FooterTemplate>
                            </asp:Repeater>
                                </div>
                     </section>
                        </div>
                    </ItemTemplate>
                </asp:Repeater>
            </div>
        </div>
    </div>
</asp:Content>

