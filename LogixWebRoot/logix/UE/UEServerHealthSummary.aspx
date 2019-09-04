<%@ Page Language="C#" MasterPageFile="~/logix/ue/UEUser.master" AutoEventWireup="true"
    CodeFile="UEServerHealthSummary.aspx.cs" Inherits="UEServerHealthSummary" %>

<%@ Register Src="~/logix/UserControls/ServerHealthTabs.ascx" TagName="ServerHealthTabs"
    TagPrefix="uc1" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
    <script type="text/javascript">

        function CheckforTerminal() {
            if ($('#ddlTerminalSet').val() == '') return false;
            else return true;
        }

        $(document).ready(function () {
            resizecolumns("gvFilesHeader", "gvFiles", "divFiles");
            resizecolumns("gvWarningsHeader", "gvWarnings", "divWarnings");
        });
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <input type="hidden" runat="server" id="hdnLocationID" clientidmode="Static" />
    <input type="hidden" runat="server" id="hdnURL" clientidmode="Static" />
    <input type="hidden" id="hdnPageCount" runat="server" clientidmode="Static" />
    <input type="hidden" id="hdnPageSize" value="30" runat="server" clientidmode="Static" />
    <input type="hidden" id="hdnPageCount1" runat="server" clientidmode="Static" />
    <input type="hidden" id="hdnPageSize1" value="20" runat="server" clientidmode="Static" />
    <uc1:ServerHealthTabs ID="ucServerHealthTabs" runat="server" AppName="ServerHealth.aspx" />
    <div id="serverhealth-main" class="serverhealth-main">
        <div id="divServerHealthSummary" runat="server">
            <div class="serverhealth_tab">
                <div id="column1">
                    <div class="box" id="identification" runat="server">
                        <h2>
                            <span>
                                <%= PhraseLib.Lookup("term.identification", LanguageID) %></span>
                        </h2>
                        <table>
                           <%-- <tr>
                                <td>
                                    <b>
                                        <%= PhraseLib.Lookup("term.server", LanguageID) %>: </b>
                                </td>
                                <td>
                                    <asp:Label ID="lblServerName" runat="server"></asp:Label>
                                </td>
                            </tr>--%>
                           <%-- <tr>
                                <td>
                                    <%= PhraseLib.Lookup("term.ipaddress", LanguageID) %>:
                                </td>
                                <td>
                                    <asp:Label ID="lblIpAddress" runat="server"></asp:Label>
                                </td>
                            </tr>--%>
                            <tr>
                                <td>
                                    <%= PhraseLib.Lookup("term.promotionengine", LanguageID)%>:
                                </td>
                                <td>
                                    <%= PhraseLib.Lookup("term.ue", LanguageID)%>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <asp:HyperLink ID="linkConfig" runat="server" />
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
                            <thead>
                                <tr>
                                    <td></td>
                                    <td style="text-align: center;"><%= PhraseLib.Lookup("term.time", LanguageID)%></td>
                                    <td style="text-align: center;"><%= PhraseLib.Lookup("term.server", LanguageID)%></td>
                                    <td style="text-align: center;"><%= PhraseLib.Lookup("term.ip", LanguageID)%></td>
                                </tr>
                            </thead>
                            <tbody>
                            <tr>
                                <td >
                                    <%= PhraseLib.Lookup("store-edit.lastofferupdate", LanguageID)%>:
                                </td>
                                <td style="text-align:center">
                                    <asp:Label ID="lblLastUpdateOffer" runat="server" Text="" ></asp:Label>(<a target="_blank"
                                        href="../log-view.aspx?filetype=200&localserverid=2"><%= PhraseLib.Lookup("term.log", LanguageID)%></a>)
                                </td>
                                <td style="text-align:center"> <asp:Label ID="lblFetchServerName" runat="server" Text=""></asp:Label></td>
                                <td style="text-align:center"> <asp:Label ID="lblFetchServerIP" runat="server" Text=""></asp:Label></td>
                            </tr>
                            <tr>
                                <td>
                                    <%= PhraseLib.Lookup("term.lastipl", LanguageID)%>:
                                </td>
                                <td style="text-align:center">
                                    <asp:Label ID="lblLastIPL" runat="server" Text=""></asp:Label>
                                </td>
                                <td style="text-align:center"> <asp:Label ID="lblIPLServerName" runat="server" Text=""></asp:Label></td>
                                <td style="text-align:center"> <asp:Label ID="lblIPLServerIP" runat="server" Text=""></asp:Label></td>
                            </tr>
                                </tbody>
                        </table>
                    </div>
                </div>
                <div id="gutter">
                </div>
                <div id="column2">
                    <div class="box" id="warnings" runat="server" clientidmode="Static">
                        <h2>
                            <span>
                                <%= PhraseLib.Lookup("term.warnings", LanguageID)%></span>
                        </h2>
                        <br />
                        <table id="gvWarningsHeader">
                            <tr>
                                <th>
                                    <%= PhraseLib.Lookup("term.description", LanguageID)%>
                                </th>
                                <th>
                                    <%= PhraseLib.Lookup("term.duration", LanguageID)%>
                                </th>
                            </tr>
                        </table>
                        <div id="divWarnings" style="height: 135px; overflow: auto;">
                            <AMSControls:AMSGridView ID="gvWarnings" runat="server" ShowHeader="false" AutoGenerateColumns="false"
                                GridLines="None" CellSpacing="2" AllowSorting="false" ShowHeaderWhenEmpty="true"
                                CssClass="fixHeader">
                                <Columns>
                                    <asp:TemplateField ItemStyle-CssClass="Description">
                                        <ItemTemplate>
                                            <a title="<%#Eval("tooltip")%>" href="<%#Eval("URL")%>">
                                                <%#Eval("NodeIP")%>&nbsp;</a><span title='<%#Eval("tooltip")%>'><%#Eval("Description")%></span>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="Duration" ItemStyle-CssClass="Duration" />
                                </Columns>
                                <RowStyle CssClass="shaded" />
                                <AlternatingRowStyle CssClass="" />
                            </AMSControls:AMSGridView>
                            <div id="loadmoreajaxloader1" runat="server" clientidmode="Static" style="background-color: White;">
                                <center>
                                    <img alt='' src="../../images/loader.gif" /></center>
                            </div>
                        </div>
                    </div>
                </div>
                <br clear="all" />
                <div class="box" style="width: 730px; overflow:auto;">
                    <h2>
                        <span>
                            <%= PhraseLib.Lookup("serverhealth.brokerfiles", LanguageID)%></span>
                    </h2>
                    <br />
                    <table id="gvFilesHeader">
                        <tr>
                            <th>
                                <%= PhraseLib.Lookup("term.record", LanguageID)%>
                            </th>
                            <th>
                                <%= PhraseLib.Lookup("term.file", LanguageID)%>
                            </th>
                            <th>
                                <%= PhraseLib.Lookup("term.age", LanguageID)%>
                            </th>
                            <th>
                                <%= PhraseLib.Lookup("term.created", LanguageID)%>
                            </th>
                            <th>
                                <%= PhraseLib.Lookup("term.download", LanguageID)%>
                            </th>
                        </tr>
                    </table>
                    <div id="divFiles" style="max-height: 300px; overflow: auto;">
                        <AMSControls:AMSGridView ID="gvFiles" runat="server" ShowHeader="false" GridLines="None"
                            CellSpacing="2" AllowSorting="false" AutoGenerateColumns="false" ShowHeaderWhenEmpty="true"
                            CssClass="fixHeader">
                            <Columns>
                                <asp:TemplateField ItemStyle-CssClass="Offer">
                                    <ItemTemplate>
                                        <a href='<%# Eval("RecordLink")%>'>
                                            <%# Eval("RecordText")%></a>
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:BoundField ItemStyle-CssClass="FileName" DataField="FileName" ItemStyle-Wrap="true" />
                                <asp:BoundField ItemStyle-CssClass="Age" DataField="Age" ItemStyle-Wrap="true" />
                                <asp:BoundField ItemStyle-CssClass="Created" DataField="Created" ItemStyle-Wrap="true" />
                                <asp:TemplateField>
                                    <ItemTemplate>
                                        <a class="Download" href='<%# Eval("Path")%><%# Eval("FileName")%>' title='<%=PhraseLib.Lookup("term.download",LanguageID)%>'
                                            target="_blank">
                                            <img alt='<%=PhraseLib.Lookup("term.download",LanguageID)%>' src="../../images/download.png"
                                                border="0" style="cursor: pointer;" />
                                        </a>
                                    </ItemTemplate>
                                </asp:TemplateField>
                            </Columns>
                        </AMSControls:AMSGridView>
                        <div id="loadmoreajaxloader" runat="server" clientidmode="Static" style="background-color: White;">
                            <center>
                                <img alt='' src="../../images/loader.gif" /></center>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
    <script type="text/javascript">

        var pageIndex = 1;
        var pageCount = parseInt($("#hdnPageCount").val());
        var pageSize = parseInt($("#hdnPageSize").val());

        var pageIndex1 = 1;
        var pageCount1 = parseInt($("#hdnPageCount1").val());
        var pageSize1 = parseInt($("#hdnPageSize1").val());

        //Load GridView Rows when DIV is scrolled
       
        LoadMoreRecords("divFiles", "gvFiles", $("#hdnURL").val() + "/enterprise/logixfiles?", "UEServerHealthSummary.aspx/GetFiles", pageIndex, pageSize, pageCount, 'div#loadmoreajaxloader',<%=LanguageID%>,'<%= PhraseLib.Lookup("term.nomorerecords", LanguageID) %>');

        LoadMoreRecords("divWarnings", "gvWarnings", $("#hdnURL").val() + "/allerrors?", "UEServerHealthSummary.aspx/GetWarnings", pageIndex1, pageSize1, pageCount1, 'div#loadmoreajaxloader1',<%=LanguageID%>,'<%= PhraseLib.Lookup("term.nomorerecords", LanguageID) %>');

        function OnSuccess(response, divid, loaderdivid) {
            try {
                var obj = JSON.parse(response.d);
                if (typeof obj != "string") {
                    $.each(obj, function (idx, item) {
                        var row = $("[id$=" + divid + "] tr").eq(1).clone(true);
              
                     if(divid == 'gvFiles'){
                            $(".Offer", row).html("<a href='"+ item.RecordLink+"' class='Offer' >"+ item.RecordText +"</a>");            
                            $(".FileName", row).html(item.FileName);
                            $(".Age", row).html(item.Age);
                            $(".Created", row).html(item.Created);
                            $(".Download", row).attr("href",item.Path+item.FileName);                        
                        }else {

                            $(".Description", row).html(item.Description);
                            $(".Duration", row).html(item.Duration);
                        }

                         $("[id$=" + divid + "]").append(row);
                    });
                } else {
                    $(loaderdivid).html("<center>" + response.d + "<\center>");
                }
            }
            catch (e) {
                $(loaderdivid).html("<center>" + e + "<\center>");
            }
        }
           
    </script>
</asp:Content>
