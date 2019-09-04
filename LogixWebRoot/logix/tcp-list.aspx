<%@ Page Title="" Language="C#" MasterPageFile="~/logix/User.master" AutoEventWireup="true"
  CodeFile="tcp-list.aspx.cs" Inherits="TCProgramList" %>

<%@ Register Src="UserControls/ListBar.ascx" TagName="ListBar" TagPrefix="uc1" %>
<%@ Register Src="UserControls/Paging.ascx" TagName="Paging" TagPrefix="uc2" %>
<%@ Register Src="UserControls/Search.ascx" TagName="Search" TagPrefix="uc3" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
  <div id="intro">
    <h1 id="title">
     <span>
        <%=PhraseLib.Lookup("term.trackablecouponprogram", LanguageID)%>
        </span>
    </h1>
    <div id="controls">
      <asp:Label ID="testLabel" runat="server" Text="" />
      <asp:Button ID="newBtn" runat="server" Text="New" class="regular" OnClick="newBtn_Click"
        TabIndex="-1" />
    </div>
  </div>
  <div id="main">
    <div id="infobar" class="red-background" runat="server" clientidmode="Static" style="display: none" />
    <uc1:ListBar ID="ListBar1" runat="server" />
    <div class="Shaded">
      <AMSControls:AMSGridView ID="gvCouponProgramList" runat="server" CssClass="list"
        GridLines="None" CellSpacing="3" AutoGenerateColumns="False" AllowSorting="True"
        OnSorting="gvCouponProgramList_Sorting" ShowHeaderWhenEmpty="true">
        <RowStyle CssClass="shaded" />
        <AlternatingRowStyle CssClass="" />
        <Columns>
          <asp:BoundField DataField="ProgramID" HeaderText="Id" HeaderStyle-CssClass="th-id"
            SortExpression="ProgramID" />
          <asp:HyperLinkField DataTextField="Name" HeaderStyle-CssClass="th-name" DataTextFormatString="{0:c}"
            DataNavigateUrlFields="ProgramID" DataNavigateUrlFormatString="~\logix\tcp-edit.aspx?tcprogramid={0}"
            HeaderText="Name" Target="_self" SortExpression="Name" ControlStyle-Width="250px" />
          <asp:BoundField DataField="ExpireDate" HeaderText="Expiry" HeaderStyle-CssClass="th-datetime"
            SortExpression="ExpireDate" />
          <asp:BoundField DataField="CreatedDate" HeaderText="Created" HeaderStyle-CssClass="th-datetime"
            SortExpression="CreatedDate" />
          <asp:BoundField DataField="LastUpdate" HeaderText="Edited" HeaderStyle-CssClass="th-datetime"
            SortExpression="LastUpdate" />
        </Columns>
      </AMSControls:AMSGridView>
    </div>
  </div>
</asp:Content>
