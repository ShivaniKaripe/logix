<%@ Page Title="" Language="C#" MasterPageFile="~/logix/User.master" AutoEventWireup="true"
    CodeFile="Buyer-edit.aspx.cs" Inherits="logix_Buyer_edit" %>

<%@ Register TagPrefix="uc" TagName="UI" Src="~/logix/UserControls/Notes.ascx" %>
<%@ Register TagPrefix="uc" TagName="Popup" Src="~/logix/UserControls/Notes-Popup.ascx" %>
<%@ Register Src="UserControls/CollapsableDiv.ascx" TagName="CollapsableDiv" TagPrefix="uc1" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <input type="hidden" id="TempFolder" name="TempFolder" runat="server" />
    <input type="hidden" id="templist" name="templist" runat="server" enableviewstate="true" />
    <div id="intro">
        <h1 id="htitle" runat="server">
        </h1>
        <div id="controls">
            <asp:Button ID="btnSave" class="regular select" Text="" runat="server" Visible="true"
                OnClick="btnSave_Click" OnClientClick="UpdateBuyerInfo();" />
            <asp:Button ID="btnDelete" class="regular delete" Text="" runat="server" Visible="true"
                OnClick="btnDelete_Click" OnClientClick="AlertUser();" />
            <uc:UI ID="ucNotesUI" runat="server" />
        </div>
    </div>
    <div id="main">
        <script type="text/javascript">

            $(document).ready(function () {
                var foldername = "";
                foldername += document.getElementById('ctl00_ContentPlaceHolder1_TempFolder').value;
                document.getElementById('folderNames').innerHTML = foldername;
            });

            function ShowPopUp() {
                var buyerid;
                var parent = window.parent.location.href;
                buyerid = parent.split("=")[1];
                if (buyerid == undefined) {
                    buyerid = 0;
                }
                javascript: openPopup('folder-browse.aspx?buyerid=' + buyerid);

            }
            function UpdateFolderName() {
                var foldername = "";
                foldername += document.getElementById('ctl00_ContentPlaceHolder1_TempFolder').value;
                document.getElementById('folderNames').innerHTML = foldername;
            }
            function UpdateBuyerInfo() {
                var folderList = document.getElementById('folderList').value;
                document.getElementById('ctl00_ContentPlaceHolder1_templist').value = folderList;

            }
            function AlertUser() {
                if (confirm('Are you sure you want to delete this?')) {
                    return true;
                }
                else {
                    event.preventDefault ? event.preventDefault() : event.returnValue = false;
                }
            }

        </script>
        <div id="infobar" class="red-background" runat="server" clientidmode="Static" style="display: none" />
        <div id="column1">
            <div class="box"  style="height:475px">
                <h2 id="hidentification" runat="server" style="float: left;">
                </h2>
                <br />
                <br />
                <div id="identityBody" style="margin-left: 10px; overflow:scroll; height:450px;">
                    <asp:Label ID="lblName" runat="server" /><br />
                    <asp:TextBox ID="txtName" runat="server" class="Medium" MaxLength="30" Style="margin-top: 3px;" />
                    <asp:RequiredFieldValidator EnableClientScript="false" ControlToValidate="txtName"
                        runat="server" ID="requirefieldName" Display="None" />
                    <br />
                    <br />
                    <asp:Label ID="lblUsers" runat="server" Text="Users:" /><br />
                    <div id="cgList">
                        <%--<asp:ListView ID="lstAvailable" runat="server" >
                        </asp:ListView>--%>
                        <asp:ListBox ID="lstAvailable" runat="server" SelectionMode="Multiple" DataTextField="FirstName"
                            DataValueField="ID" Rows="10" CssClass="longer" Style="margin-top: 3px;"></asp:ListBox>
                    </div>
                    <br />
                    <div id="divbtn" style="margin-left:1px;">
                     <asp:Button runat="server" class="regular select" ID="btnselect" Text="" OnClick="btnselect_Click">
                    </asp:Button>
                    <asp:Button runat="server" class="regular select" ID="btndeselect" Text="" OnClick="btndeselect_Click">
                    </asp:Button><br />
                    </div>

                    <br class="half" />
                    <asp:ListBox ID="lstSelected" runat="server" SelectionMode="Multiple" Rows="6" CssClass="longer"
                        DataTextField="firstname" DataValueField="ID"></asp:ListBox>
                    <br />
                    <br />
                    &nbsp &nbsp &nbsp
                    <table>
                        <tr style="width: 1px;">
                           
                            <td id="folderNames" class="c1">
                            </td>
                            <td>
                                <input id="folderList" type="hidden" value="" name="folderList" />
                                 <asp:Label ID="lblDefaultFolder" runat="server" Text="" /></td>
                               <td> <input id="btnBrowse" class="regular" style="margin-left: 30px;" type="button" onclick="ShowPopUp()"
                                    value="<%=PhraseLib.Lookup("term.browse", LanguageID) %>" name="btnBrowse" />
                            </td>
                        </tr>
                    </table>
                </div>
            </div>
        </div>
        <div id="column2" style="padding-left: 10px;">
            <div class="box" style="height:475px;">
                <h2 id="hdepartments" runat="server" style="float: left;">
                </h2>
                <br />
                <br />
                &nbsp &nbsp &nbsp
                <div id="Div3" style="margin-left:5px; margin-top:3px;">
                <div id="Div1" style="z-index:32; OVERFLOW:auto; WIDTH:97%; height:210px; ">
                   
                    <asp:ListBox ID="lstDeptAvailable" runat="server" SelectionMode="Multiple" DataTextField="ExternalID"
                        DataValueField="NodeID" Rows="13" CssClass="longer" Style="margin-left: 3px;"></asp:ListBox>     
                   
                     
                </div>
               
                 <div id="div4" style="margin-left:2.1px; margin-top:7px;">
                <asp:Button runat="server" class="regular select" ID="btnselectdept" Text="" OnClick="btnselectdept_Click">
                </asp:Button>
                <asp:Button runat="server" class="regular select" ID="btndeselectdept" Text="" OnClick="btndeselectdept_Click" /><br />
                </div>
                <br class="half" />
                <div id="Div2" style="z-index:32; OVERFLOW:auto; WIDTH:97%; height:100px;">
                <asp:ListBox ID="lstDeptSelected" runat="server" SelectionMode="Multiple" Rows="6"
                    DataTextField="ExternalID"  DataValueField="NodeID"  Style="margin-left: 3px;" CssClass="longer"></asp:ListBox>
                </div>
                <br />
                </div>
            </div>
            <br />
            &nbsp &nbsp &nbsp
            <br />
        </div>
    </div>
    <uc:Popup ID="ucNotes_Popup" runat="server" />
</asp:Content>
