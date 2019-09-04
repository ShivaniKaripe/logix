<%@ Page Language="C#" MasterPageFile="~/logix/ue/UEUser.master" AutoEventWireup="true"
    CodeFile="UEEngineHealth.aspx.cs" Inherits="UEEngineHealth" %>

<%@ Register Src="~/logix/UserControls/ServerHealthTabs.ascx" TagName="ServerHealthTabs"
    TagPrefix="uc1" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
    <script type="text/javascript">
        $(document).ready(function () {
            resizecolumns("gvEngineFilesHeader", "gvEngineFiles", "divEngineFiles");
            resizecolumns("gvBrokerFilesHeader", "gvBrokerFiles", "divBrokerFiles");
        });
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <input type="hidden" id="hdnURL" runat="server" clientidmode="Static" />
    <input type="hidden" id="hdnNodeName" runat="server" clientidmode="Static" />
    <input type="hidden" id="hdnPageCount" runat="server" clientidmode="Static" />
    <input type="hidden" id="hdnPageSize" value="10" runat="server" clientidmode="Static" />
    <input type="hidden" id="hdnPageCount1" runat="server" clientidmode="Static" />
    <input type="hidden" id="hdnPageSize1" value="10" runat="server" clientidmode="Static" />
    <input type="hidden" id="hdnLocationID" runat="server" clientidmode="Static" />
    <input type="hidden" id="hdnEnterpriseEngine" runat="server" clientidmode="Static" />
    <input type="hidden" id="hdnPendingFilesURL" runat="server" clientidmode="Static" />
    <uc1:ServerHealthTabs ID="ucServerHealthTabs" runat="server" AppName="ServerHealth.aspx" />
    <div id="serverhealth-main" class="serverhealth-main">
        <div class="engineHealth_Detailed">
            <div class="column1">
                <div class="box" id="engineidentification">
                    <h2>
                        <span>
                            <%= PhraseLib.Lookup("term.identification", LanguageID) %></span>
                    </h2>
                    <table>
                        <tr>
                            <td>
                                <b>
                                    <%= PhraseLib.Lookup("term.engine", LanguageID) %>
                                    :</b>
                            </td>
                            <td>
                                <asp:Label ID="lblEngineName" runat="server"></asp:Label>
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
                            <td>
                                <%= PhraseLib.Lookup("term.store",LanguageID) %>
                                :
                            </td>
                            <td>
                                <asp:Label ID="lblStore" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <%= PhraseLib.Lookup("term.type",LanguageID) %>
                                :
                            </td>
                            <td>
                                <asp:Label ID="lblType" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <%= PhraseLib.Lookup("term.status",LanguageID) %>
                                :
                            </td>
                            <td>
                                <asp:Label ID="lblStatus" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2">
                                <input id="Alert" runat="server" class="Alert" type="checkbox" value="" /><%= PhraseLib.Lookup("term.sendmailonerror", LanguageID) + " " + PhraseLib.Lookup("term.at", LanguageID) + " " + PhraseLib.Lookup("term.this", LanguageID).ToLower() + " " + PhraseLib.Lookup("term.engine", LanguageID).ToLower()%>
                            </td>
                        </tr>
                         <tr>
                            <td colspan="2">
                                <input id="Report" runat="server" class="Report" type="checkbox" value="" /><%= PhraseLib.Lookup("term.enable", LanguageID) + " " + PhraseLib.Lookup("term.error", LanguageID).ToLower() + " " + PhraseLib.Lookup("term.reporting", LanguageID).ToLower() + " " + PhraseLib.Lookup("term.for", LanguageID).ToLower() + " " + PhraseLib.Lookup("term.this", LanguageID).ToLower() + " " + PhraseLib.Lookup("term.engine", LanguageID).ToLower()%>
                            </td>
                        </tr>
                          <tr>
                            <td colspan="2">
                                <asp:HyperLink ID="linkConfig" runat="server" />
                            </td>
                        </tr>
                    </table>
                </div>
                <div class="box">
                    <h2>
                        <span>
                            <%= PhraseLib.Lookup("term.communication", LanguageID)%>
                        </span>
                    </h2>
                    <table>
                        <tr>
                            <td>
                                <%= PhraseLib.Lookup("store-edit.lastlookup", LanguageID)%>:
                            </td>
                            <td>
                                <asp:Label ID="lblCardholderlookup" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <%= PhraseLib.Lookup("term.Lasttransupload", LanguageID)%>
                                :
                            </td>
                            <td>
                                <asp:Label ID="lblTransactionupload" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <%= PhraseLib.Lookup("store-edit.lastofferupdate", LanguageID)%>
                                :
                            </td>
                            <td>
                                <asp:Label ID="lblsync" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <%= PhraseLib.Lookup("store-edit.lastcommunication", LanguageID)%>
                                :
                            </td>
                            <td>
                                <asp:Label ID="lblLastHeard" runat="server"></asp:Label>
                            </td>
                        </tr>
                    </table>
                </div>
            </div>
            <div id="gutter">
            </div>
            <div class="column1">
                <div class="box" id="warnings" runat="server" clientidmode="Static" >
                    <h2>
                        <span>
                            <%= PhraseLib.Lookup("term.warnings", LanguageID)%></span>
                    </h2>
                    <div style="max-height: 275px;
                    overflow: auto;">
                    <AMSControls:AMSGridView ID="gvEngineWarnings" runat="server" PageSize="50" ShowHeader="true"
                        AutoGenerateColumns="false" Height="30px" GridLines="None" CellSpacing="2" AllowSorting="false"
                        ShowHeaderWhenEmpty="true" CssClass="fixHeader">
                        <Columns>
                            <asp:BoundField DataField="Severity" />
                            <asp:TemplateField>
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
            <br clear="all" />
            <div class="box" style="width: 700px;">
                <h2>
                    <span>
                        <%=PhraseLib.Lookup("term.offer", LanguageID)+" "+PhraseLib.Lookup("serverhealth.brokerfiles", LanguageID).ToLower()%></span>
                </h2>
                <br />
                <table id="gvBrokerFilesHeader">
                    <tr>
                        <th>
                            <%= PhraseLib.Lookup("term.offer", LanguageID)%>
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
                <div id="divBrokerFiles" style="max-height: 150px; overflow: auto;">
                    <AMSControls:AMSGridView ID="gvBrokerFiles" runat="server" ShowHeader="false" GridLines="None"
                        CellSpacing="2" AllowSorting="false" AutoGenerateColumns="false" ShowHeaderWhenEmpty="true"
                        CssClass="fixHeader">
                        <Columns>
                            <asp:TemplateField ItemStyle-CssClass="Offer">
                                <ItemTemplate>
                                    <a href='<%# "UEoffer-sum.aspx?OfferID=" + Eval("OfferID")%>'>
                                        <%=PhraseLib.Lookup("term.offer",LanguageID)%>
                                        #<%# Eval("OfferID") %></a>
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
                    <div id="loadmoreajaxloader1" runat="server" clientidmode="Static" style="background-color: White;">
                        <center>
                            <img alt='' src="../../images/loader.gif" /></center>
                    </div>
                </div>
            </div>
            <div class="box" style="width: 700px;">
                <h2>
                    <span>
                        <%=PhraseLib.Lookup("term.offer", LanguageID)+" "+ PhraseLib.Lookup("serverhealth.enginefiles", LanguageID).ToLower()%></span>
                </h2>
                <br />
                <table id="gvEngineFilesHeader">
                    <tr>
                        <th>
                            <%= PhraseLib.Lookup("term.offer", LanguageID)%>
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
                    </tr>
                </table>
                <div id="divEngineFiles" style="max-height: 150px; overflow: auto;">
                    <AMSControls:AMSGridView ID="gvEngineFiles" runat="server" ShowHeader="false" GridLines="None"
                        CellSpacing="2" AllowSorting="false" AutoGenerateColumns="false" ShowHeaderWhenEmpty="true"
                        CssClass="fixHeader">
                        <Columns>
                            <asp:TemplateField>
                                <ItemTemplate>
                                    <a class="Offer" href='<%# "UEoffer-sum.aspx?OfferID=" + Eval("OfferID")%>'>
                                        <%=PhraseLib.Lookup("term.offer",LanguageID)%>
                                        #<%# Eval("OfferID") %></a>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:BoundField ItemStyle-CssClass="FileName" DataField="FileName" ItemStyle-Wrap="true" />
                            <asp:BoundField ItemStyle-CssClass="Age" DataField="Age" ItemStyle-Wrap="true" />
                            <asp:BoundField ItemStyle-CssClass="Created" DataField="Created" ItemStyle-Wrap="true" />
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
    <script type="text/javascript">
        var pageIndex = 1;
        var pageCount = parseInt($("#hdnPageCount").val());
        var pageSize = parseInt($("#hdnPageSize").val());

        var pageIndex1 = 1;
        var pageCount1 = parseInt($("#hdnPageCount1").val());
        var pageSize1 = parseInt($("#hdnPageSize1").val());

        //Load GridView Rows when DIV is scrolled
        var url =  $("#hdnPendingFilesURL").val() + "?";
        LoadMoreRecords("divEngineFiles", "gvEngineFiles", url, "UEEngineHealth.aspx/GetEngineFiles", pageIndex, pageSize, pageCount, 'div#loadmoreajaxloader',<%=LanguageID%>,'<%= PhraseLib.Lookup("term.nomorerecords", LanguageID) %>');

        if ($("#hdnEnterpriseEngine").val() == "true") {
            url = $("#hdnURL").val() + "/enterprise/logixfiles?";
        } else {
            url = $("#hdnURL").val() + "/stores/" + $("#hdnLocationID").val() + "/logixfiles?";
        }

        LoadMoreRecords("divBrokerFiles", "gvBrokerFiles", url, "UEEngineHealth.aspx/GetEngineFiles", pageIndex1, pageSize1, pageCount1, 'div#loadmoreajaxloader1',<%=LanguageID%>,'<%= PhraseLib.Lookup("term.nomorerecords", LanguageID) %>');


        function OnSuccess(response, divid, loaderdivid) {
            try {
                var obj = JSON.parse(response.d);

                if (typeof obj != "string") {
                    $.each(obj, function (idx, item) {
                        var row = $("[id$=" + divid + "] tr").eq(1).clone(true);
                     
                        $(".Offer", row).html("<a href='UEoffer-sum.aspx?OfferID=" + item.OfferID+"' class='Offer' >"+'<%=PhraseLib.Lookup("term.offer",LanguageID)%>'+" #"+ item.OfferID+"</a>");
                 
                        $(".FileName", row).html(item.FileName);
                        $(".Age", row).html(item.Age);
                        $(".Created", row).html(item.Created);
                        $(".Download", row).attr("href",item.Path+item.FileName);                        

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
           
    $(".Alert").change(function() {   
        if(PostReportAlertDatatoHealthService($("#hdnURL").val()  + "/alert/engines/<%= lblEngineName.Text %>", this.checked)==false)
             $(".Alert")[0].checked = !$(".Alert")[0].checked;
    });

     $(".Report").change(function() {
        if(PostReportAlertDatatoHealthService($("#hdnURL").val()  + "/report/engines/<%= lblEngineName.Text %>", this.checked)==false)
             $(".Report")[0].checked = !$(".Report")[0].checked;
    });
    </script>
</asp:Content>
