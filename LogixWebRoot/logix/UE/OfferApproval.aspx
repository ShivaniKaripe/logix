<%@ Page Language="C#" AutoEventWireup="true" MasterPageFile="~/logix/User.master" Debug="true" CodeFile="OfferApproval.aspx.cs" Inherits="logix_OfferApproval" %>

<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div id="intro">
        <h1 id="htitle" style="font-size: 16px;" runat="server"></h1>
        <div id="controls">
            <asp:Button AccessKey="s" CssClass="regular" ID="btnSave" runat="server" Text=""
                Visible="true" OnClick="BtnSave_Click" />
        </div>

    </div>
    <div id="main">
        <script type="text/javascript">

            $(document).ready(function () {
                ChangeControlsDisplay();
            });
               
           
            document.onclick = function (evt) {
                    
                    var target = document.all ? event.srcElement : evt.target;
                   if (target.href) {
                    if(document.aspnetForm != undefined){
                       // var isChecked = document.getElementById('<%=enableapproval.ClientID %>').checked;
                        if (IsFormChanged(document.aspnetForm)) {
                            var bConfirm = confirm('Warning:  Unsaved data will be lost.  Are you sure you wish to continue?');
                            return bConfirm;
                       }
                    }
                }
               };
            function ChangeControlsDisplay() {
                if ('<%=isBannerEnabled%>' == 'False') {
                    document.getElementById('bannerdiv').style.display = "none";
                    document.getElementById('hiddendiv').style.display = "none";
                }
                if (document.getElementById('<%=enableapproval.ClientID %>').checked) {
                    document.getElementById('usersdiv').style.display = 'block';
                    document.getElementById('selector1').style.display = 'block';
                    document.getElementById('hiddendiv').style.display = ('<%=isBannerEnabled%>' == 'False') ? 'none' : 'block';
                }
                else {
                    document.getElementById('usersdiv').style.display = 'none';
                    document.getElementById('selector1').style.display = 'none';
                    document.getElementById('hiddendiv').style.display = 'none';
                }
            }
            function FireEvent(){
                document.getElementById('<%=dummybtn.ClientID %>').click();
                
            }
        </script>
        <div id="infobar" runat="server" clientidmode="Static" style="display: none;" />
        <div id="column1">
            <asp:Button id="dummybtn" runat="server" style="display: none;" OnClick="radiodeploy_CheckedChanged" />
            <br class="half" />
            <div id="bannerdiv" style="height: 35px;">
                <asp:Label ID="lblbanner" runat="server" Text="" Visible="false"></asp:Label>
                <asp:DropDownList ID="bannerddl" runat="server" Visible="false" DataTextField="Name"
                    AutoPostBack="true" DataValueField="BannerID" OnSelectedIndexChanged="Bannerddl_SelectedIndexChanged">
                </asp:DropDownList>
            </div>
            <div id="usersdiv">
                <asp:Label ID="lbldefaultapprover" runat="server" Text=""></asp:Label>
                <asp:DropDownList ID="defaultapproverddl" runat="server" DataTextField="Name"
                    DataValueField="AdminUserID" style="    width: 150px;">
                </asp:DropDownList>

                <br />
                <br class="half" />
                <div id="selector" class="box" style="width: 330px;">
                    <h2 style="font-size: small;"><span>
                        <% if(isBannerEnabled == true)  { %>
                        <%= PhraseLib.Lookup("term.banner", LanguageID) + " " + PhraseLib.Lookup("users.deploymentpermission", LanguageID)  %>
                        <%} %>
                        <%else { %>
                        <%= PhraseLib.Lookup("users.deploymentpermission", LanguageID) %>
                        <%} %>
                        
                    </span></h2>
                    <div id="userlist">
                        <br class="half" />
                        <asp:ListBox ID="lstAvailable" runat="server" SelectionMode="Single" DataTextField="Name" AutoPostBack="true"
                            DataValueField="AdminUserID" Rows="22" CssClass="longer" OnSelectedIndexChanged="LstAvailable_SelectedIndexChanged"
                             style="margin: 15px;margin-top: 0px;overflow: scroll;"></asp:ListBox>
                    </div>
                </div>
            </div>
        </div>
        <div id="column2">
            <br class="half" />
            <asp:CheckBox ID="enableapproval" runat="server" OnCheckedChanged="Enableapproval_CheckedChanged" AutoPostBack="true" />
            <asp:Label ID="lblenableapproval" runat="server" Text=""></asp:Label><br />
            <br />
            <div id="hiddendiv" style="height: 35px;">
            </div>
            <div id="selector1" class="box">
                <h2 style="font-size: small;"><span><%= PhraseLib.Lookup("term.approverselect", LanguageID)%>
                </span></h2>
                 <asp:Table ID="approvedeploytbl" runat="server" style="margin: 10px;">
                    <asp:TableRow>
                        <asp:TableCell>
                            <asp:RadioButton ID="radiodeploy" GroupName="functionradio" AutoPostBack="true" runat="server" onchange="FireEvent()" OnCheckedChanged="radiodeploy_CheckedChanged"/>
                        </asp:TableCell>
                        <asp:TableCell>
                            <asp:Label ID="lbldeploy" runat="server" Text=""></asp:Label>
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell>
                            <asp:RadioButton ID="radioreqapproval" GroupName="functionradio" AutoPostBack="true" runat="server" OnCheckedChanged="radioreqapproval_CheckedChanged" />
                        </asp:TableCell>
                        <asp:TableCell>
                            <asp:Label ID="lblreqapproval" runat="server" Text=""></asp:Label>
                        </asp:TableCell>
                    </asp:TableRow>
                </asp:Table>
                <div id="approverlist">
                    <br class="half" />
                    <asp:ListBox ID="approverlistbox" runat="server" SelectionMode="Multiple" DataTextField="Name"
                        DataValueField="AdminUserID" Rows="10" CssClass="longer" style="margin-left: 20px;overflow: scroll;"></asp:ListBox>
                </div>
                <br class="half" />
                <asp:Button runat="server" class="regular select" ID="btnselect" Text="" CommandName="Select" OnClick="Btnselect_Click"
                     style="margin-left: 30px;"></asp:Button>
                <asp:Button runat="server" class="regular select" ID="btndeselect" Text="" CommandName="Deselect" OnClick="Btndeselect_Click"
                     style="margin-left: 40px;"></asp:Button>
                <div id="selectedapproverlist">
                    <br class="half" />
                    <asp:ListBox ID="selectedapproverlistbox" runat="server" SelectionMode="Multiple" DataTextField="Name"
                        DataValueField="AdminUserID" Rows="5" CssClass="longer" style="margin-left: 20px;margin-bottom: 18px;overflow: scroll;"></asp:ListBox>
                </div>
            </div>
        </div>
    </div>
</asp:Content>
