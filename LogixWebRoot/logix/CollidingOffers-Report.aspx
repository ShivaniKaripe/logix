<%@ Page Title="" Language="C#" MasterPageFile="~/logix/User.master" AutoEventWireup="true"
    CodeFile="CollidingOffers-Report.aspx.cs" Inherits="logix_CollidingOffers_list" %>

<%@ Register Src="UserControls/ListBar.ascx" TagName="ListBar" TagPrefix="uc1" %>
<%@ Register Src="UserControls/Paging.ascx" TagName="Paging" TagPrefix="uc2" %>
<%@ Register Src="UserControls/Search.ascx" TagName="Search" TagPrefix="uc3" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
    <script src="/javascript/jquery.min.js"></script>
    <script type="text/javascript" src="/javascript/spin.min.js"></script>
    <script type="text/javascript" src="/javascript/jquery-ui-1.10.3/jquery-ui-min.js"></script>



    <script type="text/javascript">
        var recordcount = 0;
        var sortkey = "";
        var sortorder = ""; //  0;
        var pageIndex = 0;
        var pageSize = 50;  //pick this value from System Options once we make it configurable
        var totalPages = 1;
        var CDS_OfferId = 0;
        var selfURL = "";
        $(document).ready(function () {
            gridHeaderMove();
            recordcount = $('#lblCount').text();
            sortkey = $('#hdnSortkey').val();
            sortorder = $('#hdnSortorder').val();
            CDS_OfferId = $('#lblOfferID').text();
            totalPages = Math.ceil(recordcount / pageSize);
            selfURL = "CollidingOffers-Report.aspx?ID=" + CDS_OfferId;
            $.support.cors = true;
            window.addEventListener("resize", resizeProductCollisionsGrid);
        });

        function resizeProductCollisionsGrid() {
            var ht = $('div.columnfull').height() - ($('#dvGridHeader').height() * 2);
            $('#dvGrid').css("height", ht);
        }

        function gridHeaderMove() {
            if (document.getElementById("dvGrid") != null) {
                var offID = '<%=PhraseLib.Lookup("term.offerid", LanguageID)%>';
                var offName = '<%=PhraseLib.Lookup("term.offername", LanguageID)%>';
                offID = offID.replace('&#39;', "'");
                offName = offName.replace('&#39;', "'");
                if ($('#<%= gvData.ClientID %> tbody tr').length > 1) {
                    $('#<%= gvData.ClientID %> > tbody:nth-child(1) > tr > td:nth-child(3)').hide(); //hiding the fake column
                    $('#<%= gvData.ClientID %> > tbody:nth-child(1) > tr > th:nth-child(3)').hide();
                    var str = $('#<%= gvData.ClientID %> tbody tr:first-child').html();
                    str = str.replace("<th scope=\"col\">" + offID + "</th>", "<th scope=\"col\" >" + offID + "</th><th scope=\"col\" >" + offName + "</th></tr>");
                    $('#dvGridHeader table').html('<tr>' + str + '</tr>');
                    $('#<%= gvData.ClientID %> > tbody:nth-child(1) > tr:nth-child(1)').css('display', 'none');
                    resizeProductCollisionsGrid();
                }
            }
        }

        function gridScroll(obj) {

            var div = document.getElementById('dvGrid');
            var div2 = document.getElementById('dvGridHeader');
            //****** Scrolling HeaderDiv along with DataDiv ******
            div2.scrollLeft = div.scrollLeft;
            var hasVerticalScrollbar = obj.scrollHeight > obj.clientHeight;
            if (hasVerticalScrollbar == false)  //Do not Process on Horizontal Scroll
                return;
            if ($(obj).scrollTop() + $(obj).outerHeight() >= $(obj)[0].scrollHeight) {
                GetRecords();
            }
        }

        function fnresolution() {

            var option = $('#<%=rdResolution.ClientID %> input:checked').val();
            $('#infobar').hide();
            if (option == 1) {
                // 1. check whether current offer's product group is used in any other offers
                if ($('#hdnProductList').val() == '') //product group used in the offer is not used in any other offer's condition / discount
                {
                    rmCollidingProducts(1);
                }
                else {
                    //3. Get the list of offers using current product group (use offer id to fetch current product group)(Separate by , in case of multiple)
                    rmCollidingProducts(2);
                }
            }
            else if (option == 2) {
                window.location = "/logix/UE/UEOffer-gen.aspx?OfferID=" + CDS_OfferId;
            }
            else if (option == 3) {
                // re-run collision report
                RunCollisionBackground();
            }
        }

        function toggleDialog(elemName, shown) {
            var elem = document.getElementById(elemName);
            var fadeElem = document.getElementById('fadeDiv');
            if (elem != null) {
                elem.style.display = (shown) ? 'block' : 'none';
            }
            if (fadeElem != null) {
                fadeElem.style.display = (shown) ? 'block' : 'none';
            }
        }

        function rmCollidingProducts(str) {
            if (str == 1) {
                var confirmationBox = document.getElementById("dialog-confirm");
                if (confirmationBox != null) {
                    toggleDialog('dialog-confirm', true);
                }
            }
            else {
                var confirmationBox = document.getElementById("dialog-confirm-ExistPG");
                if (confirmationBox != null) {
                    toggleDialog('dialog-confirm-ExistPG', true);
                }
            }
        }


        //Function to make AJAX call to the Web Method
        function GetRecords() {
            pageIndex = $('#hdnPageIndex').val();
            pageIndex++;
            if (pageIndex > 0 && pageIndex <= totalPages) {

                //Show Loader


                if ($('div.spinner').length == 0) {
                    $('#gridheader tbody tr:first-child a').each(function () { $(this).replacewith($(this).html()); });
                    $('#<%= gvData.ClientID %>').closest('div').css('opacity', 0.5);
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
                }

                //Call

                sortkey = $('#hdnSortkey').val();
                sortorder = $('#hdnSortorder').val();
                var tenantId = 0;

                $.ajax({
                    type: "POST",
                    url: "CollidingOffers-Report.aspx/GetCollidingProducts",
                    data: JSON.stringify({ OfferID: CDS_OfferId, pageIndex: pageIndex, sortKey: sortkey, sortOrder: sortorder }),
                    contentType: "application/json; charset=utf-8",
                    dataType: "json"
                })
                .done(function (response) {
                    OnSuccess(response);
                })
                .fail(function (xhr, status, error) {
                    alert(xhr.status + " " + xhr.statusText);
                    $('#infobar').show();
                    $('#infobar').html(xhr.status + " " + xhr.statusText);
                })
            }
        }

        //Function to recieve response of type Products<List<T>>  and append rows to GridView
        function OnSuccess(response) {
            if (response.d.ResultType != 1) {
                $('#infobar').show();
                $('#infobar').html(response.d.MessageString);
                return;
            }
            response = response.d;
            if (response.Result.Products != null) {
                response = response.Result.Products;
                if (response.length > 0) {

                    var str = "";
                    var more = 0;
                    $.each(response, function (i, re) {
                        var trcss = '';
                        if (i % 2 == 0) {
                            trcss = 'shaded';
                        }
                        str = str + "<tr class=" + trcss + "><td>" + re.ExtProductID + "</td><td>" + re.Description + "</td>";
                        str = str + "<td><div><table class='reportchildGrid' cellspacing='0'><tbody>";
                        $.each(re.Offers, function (j, inn) {
                            str = str + "<tr><td>" + inn.IncentiveID + "</td><td><a href='\\logix\\UE\\UEOffer-sum.aspx?OfferID=" + inn.IncentiveID + "'>" + inn.IncentiveName + "</a></td></tr>";
                        });

                        str = str + "</tbody></table></div>";

                        if (re.OfferCount > 100) {
                            more = re.OfferCount - 100
                            str = str + '<span>+ ' + more + '<%=PhraseLib.Lookup("term.more", LanguageID)%></span>'
                        }
                        str = str + "</td>";
                        str = str + "</tr>";
                    });



                    $('#<%= gvData.ClientID %> > tbody:nth-child(1)').append(str);
                    $('#hdnPageIndex').val(pageIndex);

                }
            }

            // remove spinner
            $('div.spinner').remove();
            $('#<%= gvData.ClientID %>').closest('div').css('opacity', 1);

        }

        function xmlhttpPost(strURL, action) {
            var xmlHttpReq = false;
            var self = this;
            var tokens = new Array();
            var runbackground = "";
            var tt = selfURL;
            if (window.XMLHttpRequest) { // Mozilla/Safari
                self.xmlHttpReq = new XMLHttpRequest();
            } else if (window.ActiveXObject) { // IE
                self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
            }
            self.xmlHttpReq.open('POST', strURL, true);
            self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');

            // window.location(strURL);
            self.xmlHttpReq.onreadystatechange = function () {
                //    alert("ok");
                if (self.xmlHttpReq != null && self.xmlHttpReq.readyState == 4) {
                    if (action = "ProductCollisionsBackgroundUE") {
                        runbackground = self.xmlHttpReq.responseText.toString();
                        if (runbackground.replace("\r\n", "") == "True") {
                            reRunCDStatus();
                        }
                    }
                }
            }
            self.xmlHttpReq.send();

        }

        function RunCollisionBackground() {
            xmlhttpPost('OfferFeeds.aspx?Mode=ProductCollisionsBackgroundUE&OfferID=' + CDS_OfferId + '&DeferDeploy=false', 'ProductCollisionsBackgroundUE');
        }

        function cancelCollision() {
            var confirmstr = '<%=PhraseLib.Lookup("confirm.cancelcollisiondetection", LanguageID)%>';
            confirmstr = confirmstr.replace('&#39;', "'");
            return confirm(confirmstr);
        }

        function reRunCDStatus() {
            $('#infobar').hide();
            $('#dvResolution').children().prop('disabled', true);

            $('#statusbar').show();
            $('#statusbar').html(' <%=PhraseLib.Lookup("term.collisiondetectioninprogress", LanguageID)%> <a id="self"> <%=PhraseLib.Lookup("term.refresh", LanguageID)%></a>');
            $('#self').prop('href', selfURL);
            $('#self').show();
            $('#canceldeploy').show();
        }

        function OfferChangeStatus() {
            $('#infobar').show();
            $('#rdResolution_2').attr('checked', true);
            $('#rdResolution_1').prop('disabled', true);
            $('#rdResolution_0').prop('disabled', true);
            $('#self').hide();
            $('#statusbar').hide();
            $('#infobar').html('<%=PhraseLib.Lookup("term.offerparamschangedreruncollision", LanguageID)%>');
        }

        function PGEmptyAfterResolution() {
            $('#rdResolution_0').prop('disabled', true);
            $('#infobar').show();
            $('#infobar').html('<%=PhraseLib.Lookup("term.collisionNotAvailable", LanguageID)%>');
        }


        function NoEditPermission() {
            $('#infobar').show();
            $('#infobar').html('<%=PhraseLib.Lookup("collision.NoEditPermission", LanguageID) %>');
        }



        String.prototype.format = function () {
            var str = this;
            for (var i = 0; i < arguments.length; i++) {
                var reg = new RegExp("\\{" + i + "\\}", "gm");
                str = str.replace(reg, arguments[i]);
            }
            return str;
        }

    </script>

</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <div id="intro">
        <h1 id="title">
            <%=PhraseLib.Lookup("term.collisionreport", LanguageID)%>:
            <asp:Label ID="lblOfferID" ClientIDMode="Static" runat="server" />
            -
            <asp:Label ID="lblOfferName" runat="server" />
        </h1>
        <div id="controls">
            <asp:Button ID="canceldeploy" name="canceldeploy" runat="server" Text="Add" OnClick="btnCancelDeploy_Click"
                OnClientClick="return cancelCollision()" Style="display: none" ClientIDMode="Static" />&nbsp;&nbsp;
            <%--<input runat="server" name="canceldeploy" id="canceldeploy" OnClientClick="return cancelCollision();"
                OnClick="btnMoveDown_Click"
                type="button" value='<%=PhraseLib.Lookup("term.cancelcollisiondetection", LanguageID)%>'
                style="display: none" />--%>
        </div>
    </div>
    <div id="main">
        <div id="infobar" class="red-background" style="display: none" runat="server" clientidmode="Static">
        </div>
        <div id="statusbar" class="green-background" style="display: none" runat="server" clientidmode="Static">
            <a id="self">
                <%=PhraseLib.Lookup("term.refresh", LanguageID)%></a>
        </div>
        <div class="column1">
            <div class="box">
                <h2>
                    <%=PhraseLib.Lookup("term.currentofferinfo", LanguageID)%>
                </h2>
                <table>
                    <tr>
                        <td style="font-weight: bold">
                            <%=PhraseLib.Lookup("term.id", LanguageID)%>:
                        </td>
                        <td>
                            <asp:Label ID="lblID" runat="server"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td style="font-weight: bold">
                            <%=PhraseLib.Lookup("term.externalid", LanguageID)%>:
                        </td>
                        <td>
                            <asp:Label ID="lblExtID" runat="server"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td style="font-weight: bold">
                            <%=PhraseLib.Lookup("term.name", LanguageID)%>:
                        </td>
                        <td>
                            <asp:Label ID="lblName" runat="server"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td style="font-weight: bold">
                            <%=PhraseLib.Lookup("term.description", LanguageID)%>:
                        </td>
                        <td>
                            <asp:Label ID="lblDescription" runat="server"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td style="font-weight: bold">
                            <%=PhraseLib.Lookup("term.startdate", LanguageID)%>:
                        </td>
                        <td>
                            <asp:Label ID="lblStartDate" runat="server"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td style="font-weight: bold">
                            <%=PhraseLib.Lookup("term.enddate", LanguageID)%>:
                        </td>
                        <td>
                            <asp:Label ID="lblEndDate" runat="server"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td style="font-weight: bold">
                            <%=PhraseLib.Lookup("term.buyerid", LanguageID)%>:
                        </td>
                        <td>
                            <asp:Label ID="lblBID" runat="server"></asp:Label>
                        </td>
                    </tr>
                </table>
            </div>
        </div>
        <div id="gutter">
        </div>
        <div class="column2" id="dvResolution">
            <div class="box" style="height: 165px">
                <h2>
                    <%=PhraseLib.Lookup("term.resolution", LanguageID)%>
                </h2>
                <asp:RadioButtonList ID="rdResolution" runat="server" ClientIDMode="Static">
                    <asp:ListItem Value="1"></asp:ListItem>
                    <asp:ListItem Value="2"></asp:ListItem>
                    <asp:ListItem Value="3"></asp:ListItem>
                </asp:RadioButtonList>
                <input type="button" id="btnSubmit" value='<%=PhraseLib.Lookup("term.execute", LanguageID)%>'
                    onclick="javascript: fnresolution();" />
            </div>
        </div>
        <div class="columnfull" style="height: 50%">
            <div class="box" style="height: 100%">
                <h2 style="width: 100%">
                    <%=PhraseLib.Lookup("term.productcollisionson", LanguageID)%>
                    <asp:Label ID="lblReportDate" runat="server"></asp:Label>
                    <span style="float: right; margin-right: 15px">
                        <%=PhraseLib.Lookup("term.total", LanguageID)%>:
                        <asp:Label ID="lblCount" runat="server" ClientIDMode="Static"></asp:Label></span>
                </h2>
                <div id="dvGridHeader" style="overflow: hidden">
                    <table id="gridHeader" class="reportfixHeader1" cellspacing="5">
                    </table>
                </div>
                <asp:UpdatePanel ID="up" runat="server">
                    <ContentTemplate>
                        <div id="dvGrid" style="overflow: auto;" onscroll="javascript:gridScroll(this);">
                            <AMSControls:AMSGridView ID="gvData" runat="server" PageSize="50" ShowHeader="true"
                                GridLines="None" CellSpacing="5" AllowSorting="true" DataKeyNames="ExtProductID"
                                ShowHeaderWhenEmpty="true" CssClass="reportfixHeader" AutoGenerateColumns="false"
                                OnRowDataBound="gvData_RowDataBound" OnSorting="gvData_Sorting">
                                <AlternatingRowStyle CssClass="" />
                                <RowStyle CssClass="shaded" />
                                <Columns>
                                    <asp:BoundField DataField="ExtProductID" SortExpression="ExtProductID" ItemStyle-Width="100px" />
                                    <%--  <asp:BoundField DataField="ProductType" SortExpression="ProductType" />--%>
                                    <asp:BoundField DataField="Description" SortExpression="Description" />
                                    <asp:TemplateField HeaderText="fakeColumn">
                                        <ItemTemplate>
                                            <asp:Label ID="Label1" runat="server"></asp:Label>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:GridView ID="gvOffers" runat="server" ShowHeader="false" AutoGenerateColumns="false"
                                                CssClass="reportchildGrid" GridLines="None" ShowFooter="false">
                                                <Columns>
                                                    <asp:BoundField DataField="IncentiveID" />
                                                    <asp:TemplateField>
                                                        <ItemTemplate>
                                                            <asp:HyperLink ID="OfferName" runat="server" Text='<%# Bind("IncentiveName") %>'
                                                                NavigateUrl='<%# String.Format("~\\logix\\UE\\UEOffer-sum.aspx?OfferID={0}",Eval("IncentiveID")) %>'></asp:HyperLink>
                                                        </ItemTemplate>
                                                    </asp:TemplateField>
                                                </Columns>
                                            </asp:GridView>
                                            <asp:Label ID="lblOfferCount" runat="server" Visible="false"></asp:Label>
                                            <asp:HiddenField ID="hdnCount" runat="server" Value='<%# Bind("OfferCount") %>' />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                </Columns>
                            </AMSControls:AMSGridView>
                            <asp:Button ID="btnTemp" runat="server" ClientIDMode="Static" Style="visibility: hidden; display: none;"
                                OnClientClick="javascript:SetDivPosition()" />
                            <asp:HiddenField runat="server" ID="hdnSortkey" ClientIDMode="Static" />
                            <asp:HiddenField runat="server" ID="hdnSortorder" ClientIDMode="Static" />
                            <asp:HiddenField runat="server" ID="hdnPageIndex" ClientIDMode="Static" Value="1" />
                        </div>
                    </ContentTemplate>
                </asp:UpdatePanel>
            </div>
        </div>
    </div>
    <div id="dialog-confirm" style="display: none; position: absolute; top: 170px; left: 475px;">
        <div id="confirmingwrap" style="width: 420px;">
            <div class="box" id="confirmingbox" style="height: auto;">
                <h2>
                    <%=PhraseLib.Lookup("term.ProductCollisions", LanguageID)%>
                </h2>
                <p>
                    <%=PhraseLib.Lookup("term.confirmRemoveColliding", LanguageID)%>
                </p>
                <br />
                <p style="text-align: center; padding: 1px">
                    <asp:Button ID="deleteItems" runat="server" Text="" CssClass="large" OnClick="deleteItems_click" />
                    <asp:Button ID="cancel" CssClass="large" runat="server" Text="" OnClientClick="javascript:toggleDialog('dialog-confirm', false);" />
                </p>
            </div>
        </div>
    </div>
    <div id="dialog-confirm-ExistPG" style="display: none; position: absolute; top: 170px; left: 475px;">
        <div id="confirmingwrap1" style="width: 420px;">
            <div class="box" id="confirmingbox1" style="height: auto;">
                <h2>
                    <%=PhraseLib.Lookup("term.ProductCollisions", LanguageID)%>
                </h2>
                <p>
                    <%=PhraseLib.Lookup("term.copyPGwithColProducts", LanguageID)%>
                </p>
                <br />
                <p style="text-align: center; padding: 1px">
                    <asp:Button ID="copyPG" runat="server" Text="" CssClass="large" OnClick="deleteItems_click" />
                    <asp:Button ID="cancel1" CssClass="large" runat="server" Text="" OnClientClick="javascript:toggleDialog('dialog-confirm-ExistPG', false);" />
                </p>
            </div>
        </div>
    </div>
    <div id="fadeDiv">
    </div>
    <asp:HiddenField runat="server" ID="hdnProductList" ClientIDMode="Static" />
    <asp:HiddenField runat="server" ID="hdnOfferList" ClientIDMode="Static" />
    <asp:HiddenField runat="server" ID="hdnIsPGEmptyAfterResolution" ClientIDMode="Static" Value="false" />
    <script type="text/javascript">
        Sys.Application.add_load(gridHeaderMove);
    </script>
</asp:Content>
