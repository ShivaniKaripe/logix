<%@ Control Language="C#" AutoEventWireup="true" CodeFile="ServerHealthTabs.ascx.cs"
    Inherits="ServerHealthTabs" %>
<script type="text/javascript" src="../../javascript/infinitescroll.js"></script>
<script type="text/javascript" src="../../javascript/logix.js"></script>
<script type="text/javascript" src="../../javascript/jquery.xdomainrequest.min.js"></script>
<div id="serverhealth-intro" class="serverhealth-intro">
    <h1 id="title">
        <asp:Label ID="lblTitle" runat="server"  />&nbsp;<asp:DropDownList runat="server" style="height: 17px;" ID="ddlEngines" AutoPostBack="true" OnSelectedIndexChanged="ddlEngines_SelectedIndexChanged" />
    </h1>
    <div id="controls">
        <input value='<%=Copient.PhraseLib.Lookup("term.logs", LanguageID)%>' type="button"
            id="btnLogs" onclick="window.open('/logix/log-view.aspx?filetype=-1&fileyear=' + <%=DateTime.Now.Year.ToString() %> + '&filemonth=' + <%=DateTime.Now.Month.ToString()  %>+'&fileday=' + <%=DateTime.Now.Day.ToString() %>, '_blank');"
            class="regular" />
        <br />
    </div>
    <asp:Label ID="infobar" runat="server" style="height: auto;" CssClass="warnings white-background infobar"
        EnableViewState="true" />
    <div id="serverhealth-subtabs" class="serverhealth-subtabs">
        <asp:LinkButton ID="ServerHealthSummary" CssClass="serverhealth_subtabs" OnClick="ServerHealthSummary_Click"
            runat="server">
        <%=Copient.PhraseLib.Lookup("term.summary", LanguageID)%>
        </asp:LinkButton>
        <asp:LinkButton ID="NodeHealth" CssClass="serverhealth_subtabs" OnClick="NodeHealth_Click"
            runat="server">
        <%=Copient.PhraseLib.Lookup("term.nodehealth", LanguageID) %>
        </asp:LinkButton>
        <asp:LinkButton ID="EngineHealth" CssClass="serverhealth_subtabs" OnClick="EngineHealth_Click"
            runat="server">
        <%=Copient.PhraseLib.Lookup("term.enginehealth", LanguageID) %>
        </asp:LinkButton>
    </div>
    <script type="text/javascript">
        function PostReportAlertDatatoHealthService(healthServiceURL, Enabled) {
            var returnValue = false;
            $.ajax({
                type: "POST",
                url: "UEServerHealthSummary.aspx/ToggleReportAlert",
                async: false,
                data: "{ URL: \"" + healthServiceURL + "\",LanguageID:<%=LanguageID%>,Enabled:" + Enabled + " }",
                contentType: "application/json; charset=utf-8",
                dataType: "json"
            })
            .done(function (response) {
                if (response.d != '')
                    returnValue = false;
                else
                    returnValue = true;
            })
            .fail(function (response) {
                returnValue = false;
            });

            return returnValue;
        }

        function HighlightRow(rowNum, IsReportEnabled, high, medium, low) {
            var row = $("#row" + rowNum);
            var errorHTML = (".Errors", row).html();

            if (rowNum % 2 != 0)
                row.attr("class", "shaded");
            else
                row.removeAttr("class");

            if (IsReportEnabled) {
                if (errorHTML != '') {
                    if (errorHTML.indexOf(high) > 0)
                        row.attr("class", "shadeddarkred");
                    else if (errorHTML.indexOf(medium) > 0)
                        row.attr("class", "shadedred");
                    else if (errorHTML.indexOf(low) > 0)
                        row.attr("class", "shadedlightred");
                }
            }
        }
    </script>
</div>
