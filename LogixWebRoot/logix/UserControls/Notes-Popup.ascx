<%@ Control Language="C#" AutoEventWireup="true" CodeFile="Notes-Popup.ascx.cs" Inherits="logix_UserControls_Notes_Popup" %>
<script type="text/javascript" src="/javascript/jquery.min.js"></script>
<script type="text/javascript" src="/javascript/jquery-ui-1.10.3/jquery-ui-min.js"></script>
<link rel="stylesheet" href="/javascript/jquery-ui-1.10.3/style/jquery-ui.min.css" />
<style type="text/css">
  .ui-dialog .ui-dialog-content
  {
    background-color: #c0c0c0;
    border: solid 1px #dddddd;
  }
  
  .ui-dialog-titlebar
  {
    background: #0066ff;
    color: #ffffff;
    font-weight: bold;
    line-height: 20px;
  }
</style>
<script language="javascript" type="text/javascript">
  function toggleNotesDisplay() {
    $("div#notes").toggle();
  }
  function toggleNotesInputBox() {
    if ($("div#notes").css("display") == "block") {
      if ($("div#notesscroll").height() == "165") {
        $("div#notesscroll").height(295);
      }
      else {
        $("div#notesscroll").height(165);
      }
      $("div#noteadddiv").toggle();
      $("div#notesinput").toggle();
    }
  }
</script>
<div id="notes" style="display: none">
  <div id="notesbody">
    <span id="notestitle">
      <%= PhraseLib.Lookup("term.notes", LanguageID)%></span> <span id="notesclose" title='<%= PhraseLib.Lookup("term.close", LanguageID)%>'>
        <a href="javascript:toggleNotesDisplay()">x</a> </span>
    <div id="notesdisplay">
      <div id="notesscroll">
        <asp:Repeater ID="rptNotes" runat="server" OnItemCommand="rptNotes_ItemCommand" 
          onitemcreated="rptNotes_ItemCreated">
          <ItemTemplate>
            <div id="<%# "note" + Eval("NoteID") %>" style="<%# Convert.ToBoolean(Eval("Private")) && AdminUserID != Convert.ToInt32(Eval("AdminUser.ID")) ? "display:none": "" %>"
              class="<%# "note" + (Convert.ToBoolean(Eval("Private")) ? " private" : "") + (Convert.ToBoolean(Eval("Important")) ? " important" : "") %>">
              <asp:HiddenField ID="IsPrivate" runat="server" />
              <a name="<%# "n" + Eval("NoteID") %>"></a><span class="notedate">
                <%# Eval("CreatedDate") %></span> <span class="noteuser">
                  <%# Eval("AdminUser.Name").ToString() %>
                </span><span class="notedelete" style="<%# UserPermissions.DeleteNotes ? "": "display: none" %>">
                  [<asp:LinkButton ID="lbdelete" CommandName="Delete" CommandArgument='<%# Eval("NoteID") %>'
                    runat="server">X</asp:LinkButton>]</span>
              <br />
              <%# Eval("Note") %>
            </div>
          </ItemTemplate>
        </asp:Repeater>
        <div id="emptyNotes" class="note" runat="server">
          <span style="color: #808050;">
            <%= PhraseLib.Lookup("notes.none", LanguageID) %></span>
          <br />
        </div>
      </div>
      <div id="noteadddiv" clientidmode="Static" style="text-align: center;" runat="server">
        <input id="noteadd" class="regular" type="button" onclick="toggleNotesInputBox();"
          value='<%= PhraseLib.Lookup("term.add", LanguageID) %>' name="noteadd" />
        <br />
      </div>
      <div id="notesinput" clientidmode="Static" style="display: none;" runat="server">
        <textarea ID="notetext" name="notetext" runat="server" ClientIDMode="Static"></textarea>
        <br />
        <asp:CheckBox ID="cbPrivate" runat="server" />
        <label for="cbPrivate">
          <%= PhraseLib.Lookup("term.private", LanguageID) %></label>
        <asp:CheckBox ID="cbImportant" Style="display: none;" runat="server" />
        <label style="display: none;" for="cbImportant">
          <%= PhraseLib.Lookup("term.important", LanguageID) %></label>
        <br />
        <asp:Button ID="notesave" class="regular" Style="margin-top: 6px;" value='<%= PhraseLib.Lookup("term.save", LanguageID) %>'
          name="notesave" runat="server" OnClick="notesave_Click" />
        <input id="notecancel" class="regular" type="button" onclick="toggleNotesInputBox();"
          style="margin-top: 6px;" value='<%= PhraseLib.Lookup("term.cancel", LanguageID) %>'
          name="notecancel" />
        <br />
      </div>
    </div>
  </div>
  <div id="notesshadow">
    <img alt="" src="/images/notesshadow.png" />
  </div>
</div>
