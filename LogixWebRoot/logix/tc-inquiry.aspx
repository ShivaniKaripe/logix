<%@ Page Title="" Language="C#" MasterPageFile="~/logix/User.master" AutoEventWireup="true"
  CodeFile="tc-inquiry.aspx.cs" Inherits="logix_tc_inquiry" %>

<%@ Register Src="UserControls/ListBar.ascx" TagName="ListBar" TagPrefix="uc1" %>
<%@ Register Src="UserControls/Paging.ascx" TagName="Paging" TagPrefix="uc2" %>
<%@ Register Src="UserControls/Search.ascx" TagName="Search" TagPrefix="uc3" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
  <style type="text/css">
  
    .ui-dialog {
      border: 2px solid black;
}

  </style>
  <script language="javascript" type="text/javascript">

    if (window.captureEvents) {
      window.captureEvents(Event.CLICK);
      window.onclick = function (e) {
        handlePageClickNew(e, 'divActionMenu', 'btnAction');
      }
    } else {

      document.onclick = function (e) {
        handlePageClickNew(e, 'divActionMenu', 'btnAction');
      }
    }

    function toggleDropdown() {
      $("#divActionMenu").toggle();
      if ($("#divActionMenu").css("display") == "none") {
        $("#btnAction").prop('value', '<%=PhraseLib.Lookup("term.actions",LanguageID)%>' + ' ▼');
      }
      else {
        $("#btnAction").prop('value', '<%=PhraseLib.Lookup("term.actions",LanguageID)%>' + ' ▲');
      }
      return false;
    }

    function SelectAllCheckboxes(spanChk) {

      $('#<%=gvCouponList.ClientID %>').find("input:checkbox").each(function () {
        if (this != spanChk) {
          this.checked = spanChk.checked;
        }
      });      
    }
    function createdialog(id) {

      var dlgsv = $("#" + id).dialog({

        modal: true,
        draggable: true,
        resizable: false,
        show: 'Transfer',
        hide: 'Transfer',
        width: 860,
        title: 'Coupon History',
        autoOpen: false,
        minHeight: 10,
        minwidth: 10,
        closeText: '<%=PhraseLib.Lookup("term.close", LanguageID)%>',
        closeOnEscape: true,
        dialogclass:'uidialogtitle',
        overlay: {
          opacity: 0.65
        }
      });

    }
    function opendialog(link, id, code) {
      $('#' + id).dialog('open');
      //            var myDialogX = $(link).position().left;
      //            var myDialogY = $(link).position().top + $(link).outerHeight();
      $('#' + id).dialog('option', 'title', "Trackable Coupon History - " + code);

    }
    function deleteConfirmation() {
      var _result = validateCheckBoxes();
      if (_result == true) {
        var confirmationResult = confirm('<%=PhraseLib.Lookup("terminal-set.confirmdelete", LanguageID)%>');
        if (confirmationResult == true) {
          var hidVal = document.getElementById('<%= hidLockStatus.ClientID %>').value;
          if (hidVal == "True") {
            document.getElementById('<%= infobar.ClientID %>').style.display = "block";
            $('div#infobar').html('<%=PhraseLib.Lookup("TrackableCouponsLock.Delete", LanguageID)%>');
            return false;
          }
          return true;
        }
        else {
          return false;
        }
      }
      else {
        return false;
      }
    }
    function validateCheckBoxes() {
      if ($('#<%=gvCouponList.ClientID%> input:checkbox:checked').length > 0) {
        return true;
      }
      else {
        alert('<%=PhraseLib.Lookup("folders.NothingSelectedtoPerformAction", LanguageID)%>');
              return false;
      }
    }
         
  </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
  <div id="intro">
    <h1 id="title">
      <%=PhraseLib.Lookup("tcinquiry.title", LanguageID)%>
    </h1>
    <div id="controls">
      <div id="divAction" runat="server" visible="true">
        <asp:Button ID="btnAction" class="regular" runat="server" ClientIDMode="Static" OnClientClick="return toggleDropdown();" />
        <input type="hidden" id="hidLockStatus" runat="server" value="0" />
        <div id="divActionMenu" class="actionsmenunew">
          <asp:Button ID="btnUnLock" Text="Unlock" class="regular" runat="server" OnClick="btnUnLock_Click"
            OnClientClick="if (!validateCheckBoxes()) return false;" />
          <asp:Button ID="btnDelete" Text="Delete" class="regular" runat="server" OnClick="btnDelete_Click"
            OnClientClick="if (!deleteConfirmation()) return false;"/>
        </div>
      </div>
    </div>
  </div>
  <div id="main">
    <div id="infobar" class="red-background" runat="server" clientidmode="Static" style="display:none" />
   <div id="searcher1" class ="customsearch"> <asp:TextBox id="txtSearch" runat="server"  MaxLength="100" Width="42%"></asp:TextBox>
  <asp:Button ID="btnSearch" runat="server" Text="Search" onclick="btnSearch_Click"  /></div>
    <div class="Shaded" style="overflow:auto">
      <AMSControls:AMSGridView ID="gvCouponList" runat="server" CssClass="list" GridLines="None" 
        CellSpacing="3" AutoGenerateColumns="False" AllowSorting="True" OnSorting="gvCouponList_Sorting"
        OnRowCreated="gvCouponList_RowCreated" DataKeyNames="CouponId,Locked" HeaderStyle-VerticalAlign="Top" >
        <RowStyle CssClass="shaded" />
        <AlternatingRowStyle CssClass="" />
        <Columns>
          <asp:TemplateField>
             <ItemTemplate>
              <asp:CheckBox ID="chkSelect" runat="server" />
            </ItemTemplate>
          </asp:TemplateField>
          <asp:TemplateField>
            <ItemTemplate>
              <p style="width: 400px; word-wrap: break-word; margin-top: 0.9em;" runat="server" ID="tccouponcode"><%# DataBinder.Eval(Container.DataItem, "CouponCode") %></p>
            </ItemTemplate>
          </asp:TemplateField>
          <asp:BoundField DataField="CouponProgramId" HeaderStyle-CssClass="th-id" Visible="false" />
          <asp:TemplateField HeaderStyle-CssClass="th-id">
            <ItemTemplate>
          <a class="linkcss" href = 'tcp-edit.aspx?tcprogramid= <%# DataBinder.Eval(Container.DataItem, "CouponProgramId")%>' > <%# DataBinder.Eval(Container.DataItem, "Program.Name")%></a>
            </ItemTemplate>
          </asp:TemplateField>
          <asp:BoundField DataField="RemainingUses" HeaderStyle-CssClass="th-id" />
          <asp:BoundField DataField="InitialUses" HeaderStyle-CssClass="th-id" />
          <asp:TemplateField HeaderStyle-CssClass="th-date" Visible="false">
            <ItemTemplate>
              <p style="word-wrap: break-word; margin-top: 0.9em;" runat="server" ID="tcexpiredate"><%# DataBinder.Eval(Container.DataItem, "ExpireDate", "{0:g}")%></p>
            </ItemTemplate>
          </asp:TemplateField>
          <asp:TemplateField HeaderStyle-CssClass="th-name" >
            <ItemTemplate>
              <asp:Label ID="lblStatus" runat="server" Text='<%#Eval("Locked").ToString().ToUpper() == "TRUE" ? "Locked" : "Unlock" %>'></asp:Label>
            </ItemTemplate>
          </asp:TemplateField>
          <asp:BoundField DataField="CouponId" HeaderStyle-CssClass="th-id" Visible="false" />
          <asp:TemplateField HeaderStyle-CssClass="th-id">
            <ItemTemplate>
            
              <a style="<%# DataBinder.Eval(Container.DataItem, "History.Count").ToString() != "0" ? "": "display:none" %>"
              class="linkcss" onclick='javascript:opendialog(this,"divHistory<%#DataBinder.Eval(Container.DataItem, "CouponID")%>","<%#DataBinder.Eval(Container.DataItem, "CouponCode") %>");'>
                <%= PhraseLib.Lookup("term.view", LanguageID) %></a>
              <%#RegisterPopUpScript((long)DataBinder.Eval(Container.DataItem, "CouponID") )%>
              <div id="divHistory<%#DataBinder.Eval(Container.DataItem, "CouponID")%>" >
                <asp:Repeater ID="repHistory" runat="server">
              
                  <HeaderTemplate>
                  <div style="height:300px;overflow-x:scroll;overflow-y:scroll">
                    <table class="list" style="padding:3px;width:840px;">
                      <tr>
                        <th class="th-datetime" >
                          <%= PhraseLib.Lookup("term.timedate", LanguageID) %>
                        </th>
                        <th class="th-status">
                          <%= PhraseLib.Lookup("term.action", LanguageID)%>
                        </th>
                        <th class="th-name" >
                          <%= PhraseLib.Lookup("term.customerid", LanguageID)%>
                        </th>
                        <th >
                          <%= PhraseLib.Lookup("term.transaction", LanguageID)%> #
                        </th>
                        <th class="th-name">
                          <%= PhraseLib.Lookup("term.transaction", LanguageID)%>
                          <%= PhraseLib.Lookup("term.status", LanguageID)%>
                        </th>
                        <th>
                          <%= PhraseLib.Lookup("term.redeem", LanguageID) %>
                          <%= PhraseLib.Lookup("term.count", LanguageID) %>
                        </th>
                        <th class="th-status">
                          <%= PhraseLib.Lookup("term.location", LanguageID) %>
                        </th>
                      </tr>
                      
                  </HeaderTemplate>
                 
                  <ItemTemplate >
                 
                    <tr align="left" class="<%# Container.ItemIndex % 2 == 0 ? "shaded" : "" %>" >                    
                      <td >
                        <%# DataBinder.Eval(Container.DataItem, "CreationDate")%>
                      </td>
                      <td>
                                <%# GetActionText(DataBinder.Eval(Container.DataItem,"Type").ToString()) %> 
                      </td>
                      <td>
                        <a class="linkcss"  href="/logix/customer-general.aspx?CustPK=<%# DataBinder.Eval(Container.DataItem, "Customer.CustomerPK") %>">                          
                            <%# DataBinder.Eval(Container.DataItem, "Customer.InitialCardID") %>
                        </a>
                      </td>
                      <td >
                        <%# DataBinder.Eval(Container.DataItem, "LogixTransNum")%>
                      </td>
                      <td>
                        <%# GetTransactionStatus(Convert.ToInt32(DataBinder.Eval(Container.DataItem, "TransStatus")))%>                     
                      </td>
                      <td>
                        <%# DataBinder.Eval(Container.DataItem, "RedeemCount")%>
                      </td>
                      <td>
                      <a class="linkcss" href="/logix/store-edit.aspx?LocationID=<%# DataBinder.Eval(Container.DataItem, "Location.LocationID") %>">
                        <%# DataBinder.Eval(Container.DataItem, "Location.LocationCode")%>
                        </a>
                      </td>
                      
                    </tr>
                  </ItemTemplate>
                  <FooterTemplate>
                    </table>
                    </div>
                  </FooterTemplate>
          
                </asp:Repeater>
              </div>
            </ItemTemplate>
          </asp:TemplateField>
        </Columns>
      </AMSControls:AMSGridView>
    </div>
  </div>
</asp:Content>
