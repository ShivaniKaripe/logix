<%@ Page Title="" Language="C#" MasterPageFile="~/logix/User.master" AutoEventWireup="true"
  CodeFile="tcp-hist.aspx.cs" Inherits="logix_tcprogram_hist" %>

<%@ Register TagPrefix="uc" TagName="UI" Src="~/logix/UserControls/Notes.ascx" %>
<%@ Register TagPrefix="uc" TagName="Popup" Src="~/logix/UserControls/Notes-Popup.ascx" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
  <script type="text/javascript" src="/javascript/logix.js"></script>
  <script type="text/javascript" src="/javascript/jquery.min.js"></script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
  <div id="intro">
    <h1 id="title" runat="server">
      Title</h1>
    <div id="controls">
      <uc:UI ID="ucNotesUI" runat="server" />
    </div>
  </div>
  <div id="main">
    <div id="infobar" clientidmode="Static" class="red-background" runat="server" visible="false">
    </div>
    <table class="list" summary="<%= PhraseLib.Lookup("term.history", LanguageID) %>">
      <thead>
        <tr>
          <th align="left" scope="col" class="th-timedate">
            <%= PhraseLib.Lookup("term.timedate", LanguageID) %>
          </th>
          <th align="left" scope="col" class="th-user">
            <%= PhraseLib.Lookup("term.user", LanguageID) %>
          </th>
          <th align="left" scope="col" class="th-action">
            <%= PhraseLib.Lookup("term.action", LanguageID) %>
          </th>
        </tr>
      </thead>
      <tbody>
        <asp:Repeater ID="rptProgramHistory" runat="server">
          <ItemTemplate>
            <tr class="<%# Container.ItemIndex % 2 == 0 ? "shaded" : "" %>">
              <td>
                <%# Eval("ActivityDate") %>
              </td>
              <td>
                <%# Eval("AdminUser.Name")%>
              </td>
              <td>
                <%# Eval("Description") %>
              </td>
            </tr>
          </ItemTemplate>
        </asp:Repeater>
      </tbody>
    </table>
  </div>
  <uc:Popup ID="ucNotes_Popup" runat="server" />
</asp:Content>
