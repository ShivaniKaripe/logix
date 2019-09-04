<%@ Page Language="C#" MasterPageFile="~/logix/ue/UEUser.master" AutoEventWireup="true"
    CodeFile="UENodeHealth.aspx.cs" Inherits="UENodeHealth" %>

<%@ Register Src="~/logix/UserControls/ServerHealthTabs.ascx" TagName="ServerHealthTabs"
    TagPrefix="uc1" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <uc1:ServerHealthTabs ID="ucServerHealthTabs" runat="server" AppName="ServerHealth.aspx" />
      <input type="hidden" id="hdnURL" runat="server" clientidmode="Static" />
    <div id="serverhealth-main" class="serverhealth-main">
        <div class="nodeHealth_Detailed">
            <div class="column1">
                <div class="box" id="identification">
                    <h2>
                        <span>
                            <%= PhraseLib.Lookup("term.identification", LanguageID) %></span>
                    </h2>
                    <table>
                        <tr>
                            <td>
                                <b>
                                    <%= PhraseLib.Lookup("term.name", LanguageID) %>
                                    :</b>
                            </td>
                            <td>
                                <asp:Label ID="lblNodeName" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <%= PhraseLib.Lookup("term.ipaddress",LanguageID) %>
                                :
                            </td>
                            <td>
                                <asp:Label ID="lblIpAddress" runat="server"></asp:Label>
                            </td>
                        </tr>
                         <tr>
                            <td colspan="2">
                                <input id="Alert" runat="server" class="Alert" type="checkbox" value="" /><%= PhraseLib.Lookup("term.sendmailonerror", LanguageID) + " " + PhraseLib.Lookup("term.at", LanguageID) + " " + PhraseLib.Lookup("term.this", LanguageID).ToLower() + " " + PhraseLib.Lookup("term.node", LanguageID).ToLower()%>
                            </td>
                        </tr>
                         <tr>
                            <td colspan="2">
                                <input id="Report" runat="server" class="Report" type="checkbox" value="" /><%= PhraseLib.Lookup("term.enable", LanguageID) + " " + PhraseLib.Lookup("term.error", LanguageID).ToLower() + " " + PhraseLib.Lookup("term.reporting", LanguageID).ToLower() + " " + PhraseLib.Lookup("term.for", LanguageID).ToLower() + " " + PhraseLib.Lookup("term.this", LanguageID).ToLower() + " " + PhraseLib.Lookup("term.node", LanguageID).ToLower()%>
                            </td>
                        </tr>
                    </table>
                </div>
                <div class="box" id="communication" runat="server">
                        <h2>
                            <span>
                                <%= PhraseLib.Lookup("term.communication", LanguageID)%>
                            </span>
                        </h2>
                        <table>
                            <tr>
                                <td >
                                    <%= PhraseLib.Lookup("store-edit.lastofferupdate", LanguageID)%>:
                                </td>
                                <td style="text-align: center; width: 200px;">
                                    <asp:Label ID="lblLastUpdateOffer" runat="server" Text=""></asp:Label>
                                </td>
                                <td></td>
                                <td></td>
                            </tr>
                            <tr>
                                <td>
                                    <%= PhraseLib.Lookup("term.lastipl", LanguageID)%>:
                                </td>
                                <td style="text-align: center;">
                                    <asp:Label ID="lblLastIPL" runat="server" Text=""></asp:Label>
                                </td>
                                <td></td>
                                <td></td>
                            </tr>
                        </table>
                    </div>
                <div class="box" id="divPB" runat="server">
                    <h2>
                        <span>
                            <%= PhraseLib.Lookup("term.promotionbroker", LanguageID)%>
                        </span>
                    </h2>
                    <table>
                        <tr>
                            <td>
                                <%= PhraseLib.Lookup("term.status", LanguageID)%>
                                :
                            </td>
                            <td>
                                <asp:Label ID="lblPBStatus" runat="server"></asp:Label>
                            </td>
                        </tr>
                           <tr>
                            <td>
                                <%= PhraseLib.Lookup("store-edit.lastcommunication", LanguageID)%>
                                :
                            </td>
                            <td>
                                <asp:Label ID="lblPBLastHeard" runat="server"></asp:Label>
                            </td>
                        </tr>
                    </table>
                </div>
                 <div class="box" id="divCB" runat="server">
                    <h2>
                        <span>
                            <%= PhraseLib.Lookup("term.customerbroker", LanguageID)%>
                        </span>
                    </h2>
                    <table>
                        <tr>
                            <td>
                                <%= PhraseLib.Lookup("term.status", LanguageID)%>
                                :
                            </td>
                            <td>
                                <asp:Label ID="lblCBStatus" runat="server"></asp:Label>
                            </td>
                        </tr>
                           <tr>
                            <td>
                                <%= PhraseLib.Lookup("store-edit.lastcommunication", LanguageID)%>
                                :
                            </td>
                            <td>
                                <asp:Label ID="lblCBLastHeard" runat="server"></asp:Label>
                            </td>
                        </tr>
                          <tr>
                            <td>
                                <%= PhraseLib.Lookup("store-edit.lastlookup", LanguageID)%>
                                :
                            </td>
                            <td>
                                <asp:Label ID="lblLastLookUp" runat="server"></asp:Label>
                            </td>
                        </tr>
                            <tr>
                            <td>
                                <%= PhraseLib.Lookup("term.Lasttransupload", LanguageID)%>
                                :
                            </td>
                            <td>
                                <asp:Label ID="lblLastTransUpload" runat="server"></asp:Label>
                            </td>
                        </tr>
                            <tr>
                            <td>
                                <%= PhraseLib.Lookup("term.lasttransdownload", LanguageID)%>
                                :
                            </td>
                            <td>
                                <asp:Label ID="lblLastTransDownload" runat="server"></asp:Label>
                            </td>
                        </tr>
                    </table>
                </div>
                
            </div>
            <div id="gutter">
            </div>
            <div class="column1">
                <div class="box" id="warnings" runat="server" clientidmode="Static">
                    <h2>
                        <span>
                            <%= PhraseLib.Lookup("term.warnings", LanguageID)%></span>
                    </h2>
                    <div   style="max-height:275px;overflow:auto;">
                    <AMSControls:AMSGridView ID="gvNodeWarnings" runat="server" PageSize="50" ShowHeader="true"
                        AutoGenerateColumns="false" Height="30px" GridLines="None" CellSpacing="2" AllowSorting="false"
                        ShowHeaderWhenEmpty="true" CssClass="fixHeader">
                        <Columns>
                            <asp:BoundField DataField="Severity" />
                            <asp:TemplateField >
                                <ItemTemplate>
                                    <a href="javascript:openPopup('../health-resolutions.aspx?FromServerHealth=1&ParamID=<%#Eval("ParamID")%>');" title='<%= PhraseLib.Lookup("store-health.resolution-note", LanguageID)%>'>
                                        <%#Eval("ParamID")%></a>
                                </ItemTemplate>
                            </asp:TemplateField>
                             <asp:BoundField DataField="Description" />
                            <asp:BoundField DataField="Duration" />
                        </Columns>
                        <RowStyle CssClass="shaded" />
                        <AlternatingRowStyle CssClass="" />
                    </AMSControls:AMSGridView>
                     <asp:Label ID="lblError" runat="server" />
                </div>
              </div>
            </div>
        </div>
    </div>
   <script type="text/javascript">
       $(".Alert").change(function () {
           if (PostReportAlertDatatoHealthService($("#hdnURL").val() + "/alert/nodes/<%= lblNodeName.Text %>", this.checked) == false)
               $(".Alert")[0].checked = !$(".Alert")[0].checked;
       });

       $(".Report").change(function () {
           if (PostReportAlertDatatoHealthService($("#hdnURL").val() + "/report/nodes/<%= lblNodeName.Text %>", this.checked) == false)
               $(".Report")[0].checked = !$(".Report")[0].checked;
       });
   </script>
</asp:Content>
