<%@ Page Title="" Language="C#" MasterPageFile="~/logix/User.master" AutoEventWireup="true"
    CodeFile="CollidingOffers-list.aspx.cs" Inherits="logix_CollidingOffers_list" %>

<%@ Register Src="~/logix/UserControls/Search.ascx" TagName="Search" TagPrefix="uc1" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
    <script type="text/javascript" src="/javascript/date.js"></script>
    <script type="text/javascript" src="/javascript/globalize.min.js"></script>
    <script type="text/javascript" src="/javascript/spin.min.js"></script>
    <script type="text/javascript">
        function gridHeaderMove() {
            if (document.getElementById("dvGrid") != null) {
                if ($('#<%= gvCollidingOfferList.ClientID %> tbody tr').length > 1) {

                    $('#dvGridHeader table').html('<tr>' + $('#<%= gvCollidingOfferList.ClientID %> tbody tr:first-child').html() + '</tr>');
                    $('#<%= gvCollidingOfferList.ClientID %> tbody tr:nth-child(1)').css('display', 'none');
                }
            }
        }

        function gridScroll() {
            var pageIndex = 0;
            var recordcount = null;
            var searchtext;
            var sortkey;
            var sortorder;
            var pagesize;
            var param;
            var adminuserId;
            var tempCount = 0;
            var totalRecords = 0;
            //Load GridView Rows when DIV is scrolled
            recordcount = $('#Totalrecordcount').val();
            sortkey = $('#sortkey').val();
            sortorder = $('#sortorder').val();
            pagesize = $('#pagesize').val();
            adminuserId = $('#adminuserId').val();
            tempCount = recordcount;
            if (parseInt(recordcount) >= parseInt(pagesize)) {
                totalRecords = pagesize;
            }
            else {
                totalRecords = recordcount;
            }
            if (totalRecords == "" || recordcount == "") {
                $('#<%= Count.ClientID %>').html('');
    }
    else {
        $('#<%= Count.ClientID %>').html('<%= PhraseLib.Lookup("term.showing", LanguageID) %> ' + totalRecords + ' <%= PhraseLib.Lookup("term.of", LanguageID) %> ' + recordcount);
    }
    $("#dvGrid").on("scroll", function (e) {
        var $o = $(e.currentTarget);
        if ($o[0].scrollHeight - $o.scrollTop() <= $o.outerHeight()) {
            GetRecords();
        }
    });
    //Function to make AJAX call to the Web Method
    function GetRecords() {
        pageIndex++;
        searchtext = $('#searchtext').val();
        tempCount = $('#Totalrecordcount').val();
        tempCount = tempCount - pagesize;
        if (pageIndex > 0 && (recordcount == null || tempCount > 0)) {
            //Show Loader

            if ($('div.spinner').length == 0) {
                $('#gridheader tbody tr:first-child a').each(function () { $(this).replacewith($(this).html()); });
                $('#<%= gvCollidingOfferList.ClientID %>').closest('div').css('opacity', 0.5);
                var opts = {

                    lines: 13, // the number of lines to draw 
                    length: 20, // the length of each line 
                    width: 10, // the line thickness 
                    radius: 25, // the radius of the inner circle 
                    corners: 1, // corner roundness (0..1) 
                    rotate: 0, // the rotation offset 
                    direction: 1, // 1: clockwise, -1: counterclockwise 
                    color: '#000', // #rgb or #rrggbb or array of colors 
                    speed: 1, // rounds per second 
                    trail: 60, // afterglow percentage 
                    shadow: false, // whether to render a shadow 
                    hwaccel: false, // whether to use hardware acceleration 
                    classname: 'spinner', // the css class to assign to the spinner 
                    zindex: 2e9 // the z-index (defaults to 2000000000) 
                };
                var div = document.getElementById('dvGrid');
                var spinner = new Spinner(opts).spin(div);
                $(".spinner").css('top', "30%");
            }

            $.support.cors = true;
            $.ajax({
                type: "POST",
                url: "/Connectors/AjaxProcessingFunctions.asmx/GetCollisionReportOfferList",
                data: JSON.stringify({ pageindex: pageIndex, sortKey: sortkey, sortOrder: sortorder, searchingText: searchtext, userid: adminuserId }),
                contentType: "application/json; charset=utf-8",
                dataType: "json"
            })
            .done(function (response) {
                OnSuccess(response);
            })
            .fail(function (xhr, status, error) {
                $('#infobar').show();
                $('#infobar').html(xhr.status + " " + xhr.statusText);
            })
        }
    }

    //Function to recieve response of type AMSResult<List<T>>  and append rows to GridView
    function OnSuccess(response) {
        if (response.d.ResultType != 1) {
            $('#infobar').show();
            $('#infobar').html(response.d.MessageString);
            return;
        }

        response = response.d.Result;


        var collidingOfferlist = response;

        totalRecords = parseInt(totalRecords) + parseInt(collidingOfferlist.length);
        $('#<%= Count.ClientID %>').html('Showing ' + totalRecords + ' of ' + recordcount);
        $.each(collidingOfferlist, function (i, re) {
            var trcss = '';
            if (i % 2 == 0) {
                trcss = 'shaded';
            }
            $('#<%= gvCollidingOfferList.ClientID %>').append('<tr class=' + trcss + '><td>' + $.trim(re.ClientOfferID) + '</td><td>' + re.IncentiveID + '</td><td>' + re.BuyerID + '</td><td><a href="\\logix\\UE\\UEoffer-sum.aspx?OfferID=' + re.IncentiveID + '">' + re.IncentiveName + '</a></td><td>' +
                         ConvertToShortDateTime(re.CollisionRanOn) + '</td><td>' + re.CollisionCount + '</td><td><a href="\\logix\\CollidingOffers-Report.aspx?ID=' + re.IncentiveID + '">' + '<%= PhraseLib.Lookup("term.viewreport", LanguageID) %>' + '</a></td></tr>');
      });
      $('div.spinner').remove();
      $('#<%= gvCollidingOfferList.ClientID %>').closest('div').css('opacity', 1);

}

    function ConvertToShortDateTime(objDate) {
        if (objDate == null) {
            return '';
        }
        var localeDateTime = new Date(parseInt(objDate.replace('/Date(', '')));
        var date = localeDateTime.toString(Globalize.culture().calendar.patterns.d + " " + Globalize.culture().calendar.patterns.T);
        return date;
    }
}
$(document).ready(gridScroll);
    </script>
    <style type="text/css">
        table.setWidth tr {
            width: 540px;
        }

        table.SetWidth th:first-child {
            width: 117px;
            word-wrap: normal;
            work-break: break-all;
        }

        table.SetWidth th:nth-child(2) {
            width: 52px;
            word-wrap: normal;
            work-break: break-all;
        }

        table.SetWidth th:nth-child(3) {
            width: 62px;
            word-wrap: normal;
            work-break: break-all;
        }

        table.SetWidth th:nth-child(4) {
            width: 121px;
            word-wrap: normal;
            work-break: break-all;
        }

        table.SetWidth th:nth-child(5) {
            width: 162px;
            word-wrap: normal;
            work-break: break-all;
        }

        table.SetWidth th:nth-child(6) {
            width: 104px;
            word-wrap: normal;
            work-break: break-all;
        }

        table.SetWidth th:nth-child(7) {
            width: 132px;
            word-wrap: normal;
            work-break: break-all;
        }
    </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <script type="text/javascript">
        Sys.WebForms.PageRequestManager.getInstance().add_endRequest(gridScroll);
    </script>
    <div id="intro">
        <h1 id="title">
            <asp:Label ID="lblTitle" runat="server" Text="" />
        </h1>
        <div id="controls">
            <asp:Label ID="testLabel" runat="server" Text="" />
        </div>
    </div>
    <div id="main">
        <div id="infobar" class="red-background" runat="server" clientidmode="Static" style="display: none" />
        <div class="searcher1" style="vertical-align: bottom">
            <uc1:Search ID="ListSearch" runat="server" />
            <div style="padding-top: 6px; padding-left: 550px;">
                <asp:Label ID="Count" runat="server" Text="Showing"></asp:Label>
            </div>
        </div>
        <br />
        <br />
        <div id="dvGridHeader" style="overflow: hidden">
            <table id="gridHeader" class="SetWidth">
            </table>
        </div>
        <div id="dvGrid" style="height: 80%; overflow: auto; position: static">
            <AMSControls:AMSGridView ID="gvCollidingOfferList" runat="server" CssClass="list"
                GridLines="None" CellSpacing="3" AutoGenerateColumns="False" AllowSorting="True"
                ShowHeader="true" ShowHeaderWhenEmpty="true" OnSorting="gvCollidingOfferList_Sorting"
                OnLoad="gvCollidingOfferList_Load"
                OnRowDataBound="gvCollidingOfferList_RowDataBound">
                <AlternatingRowStyle CssClass="" />
                <RowStyle CssClass="shaded" />
                <Columns>
                    <asp:TemplateField SortExpression="ClientOfferID">
                        <ItemTemplate>
                            <asp:Label ID="XID" runat="server" Text='<%# Bind("ClientOfferID") %>' Width="94px"
                                Style="word-wrap: normal; word-break: break-all;"></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField SortExpression="IncentiveID">
                        <ItemTemplate>
                            <asp:Label ID="ID" runat="server" Text='<%# Bind("IncentiveID") %>' Width="40px"
                                Style="word-wrap: normal; word-break: break-all;"></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField SortExpression="BuyerID">
                        <ItemTemplate>
                            <asp:Label ID="Buyer" runat="server" Text='<%# Bind("BuyerID") %>' Width="50px" Style="word-wrap: normal; word-break: break-all;"></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField SortExpression="IncentiveName">
                        <ItemTemplate>
                            <asp:HyperLink runat="server" ID="OfferName" Text='<%# Eval("IncentiveName") %>'
                                NavigateUrl='<%# String.Format("~\\logix\\UE\\UEoffer-sum.aspx?OfferID={0}",Eval("IncentiveID")) %>'
                                Width="100px" Style="word-wrap: normal; word-break: break-all;"></asp:HyperLink>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField SortExpression="CollisionRanOn">
                        <ItemTemplate>
                            <asp:Label ID="ReportRun" runat="server" Text='<%# Bind("CollisionRanOn") %>' Width="135px"
                                Style="word-wrap: normal; word-break: break-all;"></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField SortExpression="CollisionCount">
                        <ItemTemplate>
                            <asp:Label ID="CollisionCount" runat="server" Text='<%# Bind("CollisionCount") %>'
                                Width="87px" Style="word-wrap: normal; word-break: break-all;"></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField>
                        <ItemTemplate>
                            <asp:HyperLink runat="server" ID="ViewReport" Text="ViewReport" NavigateUrl='<%# String.Format("~\\logix\\CollidingOffers-Report.aspx?ID={0}",Eval("IncentiveID")) %>'
                                Width="93px" Style="word-wrap: normal; word-break: break-all;"></asp:HyperLink>
                        </ItemTemplate>
                    </asp:TemplateField>
                </Columns>
            </AMSControls:AMSGridView>
        </div>
        <asp:HiddenField runat="server" ID="searchtext" ClientIDMode="Static" />
        <asp:HiddenField runat="server" ID="Totalrecordcount" ClientIDMode="Static" />
        <asp:HiddenField runat="server" ID="sortkey" ClientIDMode="Static" />
        <asp:HiddenField runat="server" ID="sortorder" ClientIDMode="Static" />
        <asp:HiddenField runat="server" ID="pagesize" ClientIDMode="Static" />
        <asp:HiddenField runat="server" ID="adminuserId" ClientIDMode="Static" />
    </div>
</asp:Content>
