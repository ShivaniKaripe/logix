<%@ Page Title="" Language="C#" MasterPageFile="~/logix/User.master" AutoEventWireup="true"
    CodeFile="buyer-list.aspx.cs" Inherits="logix_buyer_list" %>

<%@ Register Src="UserControls/ListBar.ascx" TagName="ListBar" TagPrefix="uc1" %>
<%@ Register Src="UserControls/Paging.ascx" TagName="Paging" TagPrefix="uc2" %>
<%@ Register Src="UserControls/Search.ascx" TagName="Search" TagPrefix="uc3" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <div id="intro">
        <h1 id="title">
        <asp:Label ID="lblTitle" runat="server" Text="" />
        </h1>
        <div id="controls">
            <asp:Label ID="testLabel" runat="server" Text="" />
            <asp:Button ID="newBtn" runat="server" Text="" class="regular" OnClick="newBtn_Click"
                TabIndex="-1" />
        </div>
    </div>
    <div id="main">
        <div id="infobar" class="red-background" runat="server" clientidmode="Static" style="display: none" />
        <uc1:ListBar ID="ListBar1" runat="server" />
        <div class="Shaded"style="overflow-x:auto;white-space:nowrap;" >
            <AMSControls:AMSGridView ID="gvCouponProgramList" runat="server" CssClass="list"
        GridLines="None" CellSpacing="3" AutoGenerateColumns="False" AllowSorting="True"
        OnSorting="gvCouponProgramList_Sorting" ShowHeaderWhenEmpty="true" RowStyle-Wrap="true"  >
        <RowStyle CssClass="shaded"  />
        <AlternatingRowStyle CssClass=""  />
        <Columns >
          <asp:BoundField DataField="ID" HeaderText="Id" HeaderStyle-CssClass="th-id"
            SortExpression="b.buyerid"  />
          <asp:HyperLinkField DataTextField="externalbuyerid" HeaderStyle-CssClass="th-name"
            DataNavigateUrlFields="ID,encodedexternalbuyerid"   HeaderText="Buyer" Target="_self" SortExpression="externalbuyerid" DataNavigateUrlFormatString="~\logix\buyer-edit.aspx?externalbuyerid={1}&id={0}"/>
         <%-- <asp:BoundField DataField="FirstName" HeaderText="FirstName" HeaderStyle-CssClass="th-datetime" 
            SortExpression="firstname"  />--%>
            <asp:TemplateField HeaderText="FirstName" HeaderStyle-CssClass="th-datetime" SortExpression="firstname" >
            <ItemTemplate>
                <%# Eval("FirstName").ToString().Replace(Environment.NewLine, "<br/>")%>
            </ItemTemplate>
            </asp:TemplateField>
             <asp:TemplateField HeaderText="Surname" HeaderStyle-CssClass="th-datetime" SortExpression="LastName" >
            <ItemTemplate>
                <%# Eval("LastName").ToString().Replace(Environment.NewLine, "<br/>")%>
            </ItemTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="UserName" HeaderStyle-CssClass="th-datetime" SortExpression="UserName" >
            <ItemTemplate>
                <%# Eval("UserName").ToString().Replace(Environment.NewLine, "<br/>")%>
            </ItemTemplate>
            </asp:TemplateField>
              <asp:TemplateField HeaderText="Departments" HeaderStyle-CssClass="th-datetime" >
            <ItemTemplate>
                <%# Eval("Departments").ToString().Replace(Environment.NewLine, "<br/>")%>
            </ItemTemplate>
            </asp:TemplateField>
          <asp:BoundField DataField="Lastupdated" HeaderText="Last Updated" HeaderStyle-CssClass="th-datetime"
            SortExpression="LastUpdated" />
        </Columns>
      </AMSControls:AMSGridView> </div>
    </div>
</asp:Content>
