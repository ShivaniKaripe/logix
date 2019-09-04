<%@ Page Language="C#" MasterPageFile="~/logix/ue/UEUser.master" AutoEventWireup="true"
    CodeFile="UEEngineHealthSummary.aspx.cs" Inherits="UEEngineHealthSummary" %>

<%@ Register Src="~/logix/UserControls/ServerHealthTabs.ascx" TagName="ServerHealthTabs"
    TagPrefix="uc1" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
    <script type="text/javascript">
        $(document).ready(function () {
            resizecolumns("header", "gvEngines", "dvGrid");
        });
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <input type="hidden" runat="server" id="hdnLocationID" />
    <input type="hidden" runat="server" id="hdnURL" clientidmode="Static" />
    <input type="hidden" runat="server" id="hdnPageIndex" clientidmode="Static" />
    <input type="hidden" runat="server" id="hdnPageSize" clientidmode="Static" />
    <input type="hidden" runat="server" id="hdnSortText" clientidmode="Static" />
    <input type="hidden" runat="server" id="hdnSortDir" clientidmode="Static" />
    <input type="hidden" runat="server" id="hdnSearch" clientidmode="Static" />
    <input type="hidden" runat="server" id="hdnFilter" clientidmode="Static" />
    <input type="hidden" runat="server" id="hdnPageCount" clientidmode="Static" />
    <uc1:ServerHealthTabs ID="ucServerHealthTabs" runat="server" AppName="ServerHealth.aspx" />
    <div id="serverhealth-main" class="serverhealth-main" style="overflow: hidden;">
        <div class="engineHealth_Summary">
            <div id="listbar" class="customsearch" style="background-color: #cceecc;">
                <div id="searcher" title="<%=PhraseLib.Lookup("term.searchby",LanguageID) %> <%=PhraseLib.Lookup("term.locations",LanguageID) %>/<%=PhraseLib.Lookup("term.engines",LanguageID) %>" style="width: auto;">
                    <asp:TextBox ID="txtSearch" Name="EngineSearchTerm" MaxLength="100" runat="server"
                        Text="" Width="50%" />
                    <asp:Button runat="server" ID="btnSearch" name="search" Text="Search" OnClick="EngineSearchChanged_Event" /><br />
                </div>
                <div id="filter" title="<%=PhraseLib.Lookup("term.filter",LanguageID) %>" style="float: right;
                    padding-right: 20px;">
                    <asp:DropDownList ID="filterEngineHealth" runat="server" ClientIDMode="Static" name="filterEngineHealth"
                        OnSelectedIndexChanged="EngineFilterChanged_Event" AutoPostBack="true">
                    </asp:DropDownList>
                </div>
                <hr class="hidden" />
            </div>
            <br />
            <table id="header">
                <tr>
                    <th>
                    </th>
                    <th>
                        <asp:LinkButton ID="lnkEngine" runat="server" OnClick="SortChanged_Event" CommandArgument="engine"
                            ClientIDMode="Static"><%=PhraseLib.Lookup("term.engine", LanguageID)%></asp:LinkButton>
                        <span id="div_nodeIp" name="div_nodeIp" clientidmode="Static" runat="server"></span>
                    </th>
                    <th>
                        <asp:LinkButton ID="lnkstorename" runat="server" OnClick="SortChanged_Event" CommandArgument="location"
                            ClientIDMode="Static"><%=PhraseLib.Lookup("term.store", LanguageID)%></asp:LinkButton>
                        <span id="div_storeName" runat="server" clientidmode="Static"></span>
                    </th>
                    <th>
                        <%=PhraseLib.Lookup("term.errors", LanguageID)%>
                    </th>
                    <th> <asp:LinkButton ID="lnkReport" runat="server" OnClick="SortChanged_Event" CommandArgument="report"
                            ClientIDMode="Static"><%=PhraseLib.Lookup("term.report", LanguageID)%></asp:LinkButton>
                        <span id="div_report" runat="server" clientidmode="Static"></span>                       
                    </th>
                    <th>
                     <asp:LinkButton ID="lnkAlert" runat="server" OnClick="SortChanged_Event" CommandArgument="alert"
                            ClientIDMode="Static"> <%=PhraseLib.Lookup("term.alert", LanguageID)%></asp:LinkButton>
                        <span id="div_alert" runat="server" clientidmode="Static"></span>
                       
                    </th>
                </tr>
            </table>
            <div id="dvGrid" style="height: 560px; width: 100%; overflow: auto; background-color: White;">
                <AMSControls:AMSGridView ID="gvEngines" AutoGenerateColumns="false" runat="server"
                    ShowHeader="false" DataKeyNames="RowNum" GridLines="None" AllowSorting="false"
                    ShowHeaderWhenEmpty="true" CssClass="fixHeader gridView" AllowPaging="false"
                    OnRowDataBound="gvEngines_OnRowDataBound" CellSpacing="2" Width="100%">
                    <Columns>
                        <asp:TemplateField>
                            <ItemTemplate>
                                <a id="link" clientidmode="Static" runat="server" class="nodeSummaryRow" >
                                    <img alt="Details" runat="server" id="plusimage" clientidmode="Static" src="../../images/plus2.png"
                                        border="0" style="cursor: pointer;" />
                                </a>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField ItemStyle-CssClass="Engine">
                            <ItemTemplate>
                                <asp:HyperLink runat="server" DataNavigateUrlFields="NodeName" NavigateUrl='<%# "UEEngineHealth.aspx?NodeName=" + Eval("NodeName")%>'
                                    ToolTip='<%# Eval("NodeName")%>'> <%#Eval("NodeIP")%>
                                </asp:HyperLink>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:BoundField DataField="StoreName" ItemStyle-Wrap="true" ItemStyle-CssClass="Location col" />
                        <asp:BoundField DataField="Errors" ItemStyle-Wrap="true" ItemStyle-CssClass="Errors" />
                        <asp:TemplateField ItemStyle-HorizontalAlign="Center">
                            <ItemTemplate>
                                <img runat="server" alt="Report" clientidmode="Static" class="Report" id="Report"
                                    src="../../images/report-on.png" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField ItemStyle-HorizontalAlign="Center">
                            <ItemTemplate>
                                <img runat="server" alt="Alert" clientidmode="Static" id="Alert" class="Alert" src="../../images/email.png" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField>
                            <ItemTemplate>
                                </td></tr>
                                <tr id="div<%# Eval("RowNum").ToString()%>" style="display: none;" class="errordetails">
                                    <td style="background-color: White;" />
                                    <td colspan="5">
                                        <asp:GridView ID="gvErrors" runat="server" AutoGenerateColumns="false" ShowHeader="true"
                                            GridLines="None" DataKeyNames="RowNum" Width="100%">
                                            <Columns>
                                                <asp:BoundField DataField="Severity" />
                                                <asp:TemplateField>
                                                    <ItemTemplate>
                                                        <a href="javascript:openPopup('../health-resolutions.aspx?FromServerHealth=1&ParamID=<%#Eval("ParamID")%>');" title="<%=PhraseLib.Lookup("store-health.resolution-note", LanguageID)%>">
                                                            <%#Eval("ParamID")%></a>
                                                    </ItemTemplate>
                                                </asp:TemplateField>
                                                <asp:BoundField DataField="Description" />
                                                <asp:BoundField DataField="Duration" />
                                            </Columns>
                                            <RowStyle CssClass="" />
                                            <AlternatingRowStyle CssClass="" />
                                        </asp:GridView>
                                    </td>
                                </tr>
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
    <script type="text/javascript">

        var pageIndex = 1;
        var pageCount = parseInt($("#hdnPageCount").val());
        var pageSize = parseInt($("#hdnPageSize").val());

        $('#dvGrid').css({ 'height': (($(window).height()) - 300) + 'px' });

        $(window).resize(function () { // On resize
            $('#dvGrid').css({ 'height': (($(window).height()) - 300) + 'px' });
        });

        var filter = "";
        switch ($("#hdnFilter").val()) {
            case "0":
                filter = "AllErrors";
                break;

            case "1":
                filter = "ShowAll";
                break;

            case "2":
                filter = "AllEnterpriseEngines";
                break;

            case "3":
                filter = "AllStoreEngines";
                break;

            case "4":
                filter = "DisconnectedEngines";
                break;

            case "5":
                filter = "CommunicationsOK";
                break;
        }
        //Load GridView Rows when DIV is scrolled
        var url = $("#hdnURL").val() + "/engines?&filter=" + filter + "&search=" + encodeURIComponent($("#hdnSearch").val()) + "&sort=" + $("#hdnSortText").val() + "&sortdir=" + $("#hdnSortDir").val() + "&";
        if ($("#hdnFilter").val() == 4 || $("#hdnFilter").val() == 0)
            url = url + "report=true" + "&";
        else
            url = url + "report=all" + "&";
        LoadMoreRecords("dvGrid", "gvEngines", url, "UEEngineHealthSummary.aspx/GetEngines", pageIndex, pageSize, pageCount, 'div#loadmoreajaxloader','<%=LanguageID%>','<%= PhraseLib.Lookup("term.nomorerecords", LanguageID) %>');

        //Function to recieve XML response append rows to GridView
        function OnSuccess(response, gridviewid, loaderdivid) {
            try {
                pageIndex++;

                var pendingRecords = JSON.parse(response.d);

                if (typeof pendingRecords != "string") {
                    var RowNum = pageIndex * pageSize;

                    $.each(pendingRecords, function (idx, item) {
                        var row = $("[id$=" + gridviewid + "] tr").eq(0).clone(true);
                        var errorRow = $("[id$=div1] tr").eq(0).clone(true);

                        RowNum++;

                        $(".Engine", row).html("<a title='" + item.nodeName + "' href=\"UEEngineHealth.aspx?NodeName=" + item.nodeName + "\">" + item.nodeIp + "</a>");
                        $(".Location", row).html(item.storeName);
                        $(".Errors", row).html(item.errorString);

                        row.attr("id", "row" + RowNum.toString());
                        errorRow.html("<td style='background-color: White;' /><td colspan='5'><table style=\"width:100%;border-collapse:collapse;\">" + item.errorHtml + "</table></td>");
                        errorRow.attr("id", "div" + RowNum.toString());
                        errorRow.css("display", "none");
                        errorRow.attr("class", "errordetails");

                        $(".nodeSummaryRow img", row).attr("id", "imgdiv" + RowNum.toString());

                        if (item.errorHtml == '') {
                            $(".nodeSummaryRow img", row).attr("src", "../../images/plus2-disabled.png");
                            $(".nodeSummaryRow", row).attr("href", "");
                        } else {
                            $(".nodeSummaryRow img", row).attr("src", "../../images/plus2.png");
                            $(".nodeSummaryRow", row).attr("href", "JavaScript:divexpandcollapse('div" + RowNum.toString() + "')");
                        }

                        $(".Report", row).attr("onclick", "javascript:ToggleReport(this,'" + item.nodeName + "','" + $("#hdnURL").val() + "','row" + RowNum.toString() + "','div" + RowNum.toString() + "');");
                        $(".Alert", row).attr("onclick", "javascript:ToggleAlert(this,'" + item.nodeName + "','" + $("#hdnURL").val() + "');");

                        if (item.Report == true)
                            $(".Report", row).attr("src", "../../images/report-on.png");
                        else
                            $(".Report", row).attr("src", "../../images/report-off.png");

                        if (item.Alert == true)
                            $(".Alert", row).attr("src", "../../images/email.png");
                        else
                            $(".Alert", row).attr("src", "../../images/email-off.png");

                        $("[id$=" + gridviewid + "]").append(row);
                        $("[id$=" + gridviewid + "]").append(errorRow);

                        HighlightRow(RowNum, item.Report, '<%= PhraseLibExtension.PhraseLib.Lookup("term.high", LanguageID)%>', '<%= PhraseLibExtension.PhraseLib.Lookup("term.medium", LanguageID)%>', '<%= PhraseLibExtension.PhraseLib.Lookup("term.low", LanguageID)%>');

                    });
                } else {
                    $(loaderdivid).html("<center>" + response.d + "<\center>");
                }
            } catch (e) {
                $(loaderdivid).html(e);
            }

        }

        function divexpandcollapse(divname) {
            var img = "img" + divname;
            if ($("#" + img).attr("src") == "../../images/plus2.png") {
              $("#" + divname).css('display', '');
                $("#" + img).attr("src", "../../images/minus2.png");
            } else {
                $("#" + divname).css('display', 'none') 
                $("#" + img).attr("src", "../../images/plus2.png");
            }
        }

        function ToggleReport(imgReport, hostName, url, row, div) {
            if (imgReport.src.indexOf("report-off.png") > 0 && PostReportAlertDatatoHealthService(url + "/report/engines/" + hostName, true)) {
                    $(imgReport).attr("src","../../images/report-on.png");
                    if ($("#hdnFilter").val() == 1)
                        HighlightRow(parseInt(row.replace("row", "")), true, '<%= PhraseLibExtension.PhraseLib.Lookup("term.high", LanguageID)%>', '<%= PhraseLibExtension.PhraseLib.Lookup("term.medium", LanguageID)%>', '<%= PhraseLibExtension.PhraseLib.Lookup("term.low", LanguageID)%>');
            }
                else if (PostReportAlertDatatoHealthService(url + "/report/engines/" + hostName, false)) {
                if ($("#hdnFilter").val() == 4 || $("#hdnFilter").val() == 0) {
                    $("#dvGrid").scrollTop($("#dvGrid")[0].scrollTop + 10);
                    $("#" + row).css('display', 'none');
                    $("#" + div).css('display', 'none');
                } else if ($("#hdnFilter").val() == 1)
                    HighlightRow(parseInt(row.replace("row", "")), false, '<%= PhraseLibExtension.PhraseLib.Lookup("term.high", LanguageID)%>', '<%= PhraseLibExtension.PhraseLib.Lookup("term.medium", LanguageID)%>', '<%= PhraseLibExtension.PhraseLib.Lookup("term.low", LanguageID)%>');
                           
                 $(imgReport).attr("src","../../images/report-off.png");                
            }
        }

        function ToggleAlert(imgAlert, hostName, url) {
            if (imgAlert.src.indexOf("email-off.png") > 0 && PostReportAlertDatatoHealthService(url + "/alert/engines/" + hostName, true))
                $(imgAlert).attr("src", "../../images/email.png");
            else if (PostReportAlertDatatoHealthService(url + "/alert/engines/" + hostName, false))
                $(imgAlert).attr("src", "../../images/email-off.png");
        }

     
    </script>
</asp:Content>
