<%@ Page Language="C#" AutoEventWireup="true" MasterPageFile="~/logix/User.master" Debug="true" CodeFile="PendingApproval.aspx.cs" Inherits="logix_UE_PendingApproval" %>

<%@ Register Src="~/logix/UserControls/ListBar.ascx" TagName="ListBar" TagPrefix="uc1" %>
<%@ Register Src="~/logix/UserControls/Paging.ascx" TagName="Paging" TagPrefix="uc2" %>
<%@ Register Src="~/logix/UserControls/Search.ascx" TagName="Search" TagPrefix="uc3" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <link rel="stylesheet" href="../../css/logix-screen.css" />
    <style type="text/css">
        .button
        {
            background-size: 15px;
            border: none;    
            width: 15px;
        }
        .approve
        {
            background-image:url("../../images/approve.png");
        }
        .reject
        {
            background-image:url("../../images/reject.png");
        }
        .viewreport
        {
            background-image:url("../../images/viewreport.png");
        }
    </style>
    <script type="text/javascript">
        function checkLength(){
            elem = document.getElementById("rejectText");
            elem.onkeydown = function() {
                var key = event.keyCode || event.charCode;
                if( key == 8 || key == 46 ){
                }
                else{
                    return check(elem);
                }
            };
            elem.onkeyup = function() {
                var key = event.keyCode || event.charCode;
                if( key == 8 || key == 46 ){
                }
                else{
                    return check(elem);
                }
            };
        }
        function check(elem){
            if(elem.value.length >= 500)
                    {
                        document.getElementById("rejectText").value = elem.value.substring(0, 500);
                        return false;
                    }
                    else
                        return true;

        }
        function xmlhttpPost(strURL, action) {
            var xmlHttpReq = false;
            var self = this;
            var tokens = new Array();
            var runbackground = "";
            if (window.XMLHttpRequest) { // Mozilla/Safari
                self.xmlHttpReq = new XMLHttpRequest();
            } else if (window.ActiveXObject) { // IE
                self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
            }
            self.xmlHttpReq.open('POST', strURL, true);
            self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
            self.xmlHttpReq.send();

        }
        function ApproveOffer(offerId, approvalType){
            var isOCDEnabled = ('<%=isOCDEnabled%>' == 'True') ? true : false;
            xmlhttpPost('../OfferFeeds.aspx?Mode=ApproveOffer&OfferID=' + offerId + '&ApprovalType=' + approvalType + '&OCDEnabled=' + isOCDEnabled, 'ApproveOffer'  );
            window.location.reload();
        }
        function toggleDialog(elemName, shown) {
            var elem = document.getElementById(elemName);
            var fadeElem = document.getElementById('fadeDiv');
            if (elem != null) {
                elem.style.display = (shown) ? 'block' : 'none';
            }
            if (fadeElem != null) {
                fadeElem.style.display = (shown) ? 'block' : 'none';
            }
        }
        function showRejectConfirmation() {
                toggleDialog('oawreject', true);
          }
        function hideRejectConfirmation() {
                document.getElementById('rejectText').value = '';
                toggleDialog('oawreject', false);
          }
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div id="intro">
        <h1 id="htitle" style="font-size: 16px;" runat="server"></h1>
        <div id="controls">
        </div>

    </div>
    <div id="main">
        <div id="accessdenied" runat="server" style="display:none;">
            <h1><span><%=PhraseLib.Lookup("term.accessdenied", LanguageID) %></span></h1><br />
            <div id="forbidden">
                <%=PhraseLib.Lookup("error.forbidden", LanguageID) %>
            </div><br />
            <div id="reqpermission">
                <%=PhraseLib.Lookup("term.requiredpermission", LanguageID) + " : " + PhraseLib.Lookup("term.offerapproval", LanguageID) %>
            </div>
        </div>
        <div id="infobar" runat="server" clientidmode="Static" style="display: none;" />
        <uc3:Search ID="Search" runat="server" />
        <uc2:Paging ID="Paging" runat="server" />
        <div id="parentfilter" runat="server">
        <div id="filter" title="Filter">
            <asp:DropDownList ID="filterddl" runat="server" AutoPostBack="true" OnSelectedIndexChanged="filterddl_SelectedIndexChanged">
                <asp:ListItem Value="0"></asp:ListItem>
                <asp:ListItem Value="1"></asp:ListItem>
            </asp:DropDownList>
        </div></div><br /><br /><br />
        <div id="oawreject" style="display: none;">
            <div id="oawrejectwrap" style="width: 420px;">
                <div class="box" id="oawrejectbox" style="height: auto;">
                    <asp:HiddenField ID="OfferID" runat="server" Value="" />
                    <h2><span><%=PhraseLib.Lookup("term.offerrejection", LanguageID) %></span></h2>
                    <p>
                        <br />
                        <%=PhraseLib.Lookup("term.rejectionreason", LanguageID) %>:
                    </p>
                    <p style="text-align: center; padding: 1px">
                        <textarea rows="4" cols="20" id="rejectText" name="rejectText" class="boxsizingBorder" style="resize: none" maxlength="500"
                             onchange="return checkLength();" onkeyup="return checkLength();" onkeydown="return checkLength();"></textarea>
                        <br class="half" /><br />
                          <small>(<%= PhraseLib.Lookup("offerrejection.rejectiontext", LanguageID) %>)</small>
                          <br />
                        <span style="padding: 15px"></span>
                        <br />
                        <br />
                        <asp:Button ID="reject" runat="server" OnClick="reject_Click" CssClass="large" />
                        <asp:Button ID="cancel" runat="server" OnClick="cancel_Click" CssClass="large" />
                    </p>
                </div>
            </div>
        </div>
        <div class="Shaded" style="overflow-x: auto; white-space: nowrap;display: inline-block;">
            <AMSControls:AMSGridView ID="gvPendingOfferList" runat="server" CssClass="list"
                GridLines="None" CellSpacing="3" AutoGenerateColumns="False" AllowSorting="True"
                ShowHeader="true" ShowHeaderWhenEmpty="true" OnSorting="gvPendingOfferList_Sorting" 
                OnRowDataBound="gvPendingOfferList_RowDataBound" style="white-space: normal;" HeaderStyle-ForeColor="#0000cc">
                <AlternatingRowStyle CssClass="" />
                <RowStyle CssClass="shaded" />
                <Columns>
                    <asp:TemplateField SortExpression="ID" ItemStyle-Width="30px">
                        <ItemTemplate>
                            <asp:Label ID="ID" runat="server" Text='<%# Bind("IncentiveID") %>' Width="50px"
                                Style="word-wrap: normal; word-break: break-all;"></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField SortExpression="OfferName" ItemStyle-Width="150px">
                        <ItemTemplate>
                            <asp:HyperLink runat="server" ID="OfferName" Text='<%# Eval("IncentiveName") %>'
                                NavigateUrl='<%# String.Format("~\\logix\\UE\\UEoffer-sum.aspx?OfferID={0}",Eval("IncentiveID")) %>'
                                Width="140px" Style="word-wrap: normal; word-break: break-all;"></asp:HyperLink>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField SortExpression="StartDate" ItemStyle-Width="70px">
                        <ItemTemplate>
                            <asp:Label ID="StartDate" runat="server" Text='<%# Bind("StartDate") %>' Width="70px" Style="word-wrap: normal; word-break: break-all;"></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField SortExpression="SubmittedBy" ItemStyle-Width="150px">
                        <ItemTemplate>
                            <asp:Label ID="SubmittedBy" runat="server" Text='<%# Bind("SubmittedBy") %>' Width="140px" Style="display: inline; word-wrap: normal; word-break: break-all;"></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField SortExpression="WaitingSince" ItemStyle-Width="120px">
                        <ItemTemplate>
                            <asp:Label ID="WaitingSince" runat="server" Text='<%# Bind("WaitingSince") %>' Width="120px"
                                Style="word-wrap: normal; word-break: break-all;"></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField ItemStyle-Width="80px">
                        <ItemTemplate>
                            <asp:Button ID="Approve" OnClick="Approve_Click" CssClass="button approve" runat="server" Height="11pt" Style="margin-top: 3px;"/>
                            <asp:Button ID="Reject" OnClick="Reject_Click" CssClass="button reject" runat="server" Height="11pt" Style="margin-top: 3px;" 
                                />
                            <asp:Button ID="CollisionReport" OnClick="CollisionReport_Click" CssClass="button viewreport" 
                                runat="server" Height="11pt" Visible="false" Style="margin-top: 3px;"/>
                        </ItemTemplate>
                    </asp:TemplateField>
                </Columns>
            </AMSControls:AMSGridView>
        </div>
    </div>
</asp:Content>
