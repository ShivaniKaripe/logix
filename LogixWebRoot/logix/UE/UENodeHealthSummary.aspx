<%@ Page Language="C#" MasterPageFile="~/logix/ue/UEUser.master" AutoEventWireup="true"
  CodeFile="UENodeHealthSummary.aspx.cs" Inherits="UENodeHealthSummary" %>

<%@ Register Src="~/logix/UserControls/ServerHealthTabs.ascx" TagName="ServerHealthTabs"
  TagPrefix="uc1" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
  <script type="text/javascript">
    $(document).ready(function () {
      resizecolumns("header", "gvNodes", "dvGrid");
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
    <div class="nodeHealth_Summary">
      <div id="listbar" class="customsearch" style="background-color: #cceecc;">
        <div id="searcher" title='<%=PhraseLib.Lookup("term.searchby",LanguageID) %> <%=PhraseLib.Lookup("term.nodes",LanguageID) %>'
          style="width: auto;">
          <asp:TextBox ID="txtSearch" Name="NodeSearchTerm" MaxLength="100" runat="server"
            Text="" Width="50%" />
          <asp:Button runat="server" ID="btnSearch" name="search" Text="Search" OnClick="NodeSearchChanged_Event" /><br />
        </div>
        <div id="filter" title="<%=PhraseLib.Lookup("term.filter",LanguageID) %>" style="float: right;">
          <asp:DropDownList ID="filterNodeHealth" runat="server" ClientIDMode="Static" name="filterNodeHealth"
            OnSelectedIndexChanged="NodeFilterChanged_Event" AutoPostBack="true">
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
            <asp:LinkButton ID="lnkNode" runat="server" OnClick="SortChanged_Event" CommandArgument="node"
              ClientIDMode="Static"><%=PhraseLib.Lookup("term.node",LanguageID) %></asp:LinkButton>
            <span id="div_node" name="div_node" clientidmode="Static" runat="server"></span>
          </th>
          <th>
            <%=PhraseLib.Lookup("term.errors",LanguageID) %>
          </th>
          <th>
            <asp:LinkButton ID="lnkReport" runat="server" OnClick="SortChanged_Event" CommandArgument="report"
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
        <AMSControls:AMSGridView ID="gvNodes" AutoGenerateColumns="false" runat="server"
          ShowHeader="false" DataKeyNames="RowNum" GridLines="None" AllowSorting="false"
          ShowHeaderWhenEmpty="true" CssClass="fixHeader gridView" AllowPaging="false" OnRowDataBound="gvNodes_OnRowDataBound"
          CellSpacing="2" Width="100%">
          <Columns>
            <asp:TemplateField>
              <ItemTemplate>
                <a id="link" clientidmode="Static" class="nodeSummaryRow" href="JavaScript:divexpandcollapse('pblinkdiv<%# Eval("RowNum").ToString()%>','imgdiv<%# Eval("RowNum").ToString()%>','pberrordiv<%# Eval("RowNum").ToString()%>','pbimgdiv<%# Eval("RowNum").ToString()%>');divexpandcollapse('cblinkdiv<%# Eval("RowNum").ToString()%>','imgdiv<%# Eval("RowNum").ToString()%>','cberrordiv<%# Eval("RowNum").ToString()%>','cbimgdiv<%# Eval("RowNum").ToString()%>');">
                  <img alt="Details" id="imgdiv<%# Eval("RowNum").ToString()%>" clientidmode="Static"
                    src="../../images/plus2.png" border="0" title="<%=PhraseLib.Lookup("term.viewhidedetails",LanguageID) %>"
                    style="cursor: pointer;" />
                </a>
                <br class="half" />
              </ItemTemplate>
            </asp:TemplateField>
            <asp:TemplateField ItemStyle-CssClass="Node">
              <ItemTemplate>
                <asp:HyperLink ID="HyperLink1" runat="server" DataNavigateUrlFields="NodeIP" NavigateUrl='<%# "UENodeHealth.aspx?NodeName=" + Eval("NodeName")%>'
                  ToolTip='<%# Eval("NodeName")%>'> <%#Eval("NodeIP")%>
                </asp:HyperLink>
              </ItemTemplate>
            </asp:TemplateField>
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
                <tr id="pblinkdiv" class="pblinkdivClass" runat="server" clientidmode="Static" style="background-color: White;
                  display: none;">
                  <td />
                  <td colspan="4">
                    <a id="pblink" runat="server" clientidmode="Static">
                      <img alt="Details" id="pbimgdiv" class="pbimgClass" runat="server" clientidmode="Static"
                        src="../../images/plus2.png" border="0" style="cursor: pointer;" />
                    </a>&nbsp;<%=PhraseLib.Lookup("term.promotionbroker", LanguageID)%>
                  </td>
                </tr>
                <tr id="pberrordiv<%# Eval("RowNum").ToString()%>" class="PBerrorsClass  errordetails"
                  style="display: none">
                  <td style="background-color: White" />
                  <td colspan="4">
                    <div>
                      <asp:GridView ID="gvPBErrors" runat="server" AutoGenerateColumns="false" ShowHeader="true"
                        GridLines="None" DataKeyNames="RowNum" Width="100%">
                        <Columns>
                          <asp:BoundField DataField="Severity" />
                          <asp:TemplateField>
                            <ItemTemplate>
                              <a href="javascript:openPopup('../health-resolutions.aspx?FromServerHealth=1&ParamID=<%#Eval("ParamID")%>');"
                                title="<%=PhraseLib.Lookup("store-health.resolution-note",LanguageID) %>">
                                <%#Eval("ParamID")%></a>
                            </ItemTemplate>
                          </asp:TemplateField>
                          <asp:BoundField DataField="Description" />
                          <asp:BoundField DataField="Duration" />
                        </Columns>
                        <RowStyle CssClass="" />
                        <AlternatingRowStyle CssClass="" />
                      </asp:GridView>
                    </div>
                  </td>
                </tr>
                <tr id="cblinkdiv" runat="server" clientidmode="Static" class="cblinkdivClass" style="background-color: White;
                  display: none;">
                  <td />
                  <td colspan="4">
                    <div>
                      <a id="cblink" runat="server" clientidmode="Static">
                        <img alt="Details" runat="server" class="cbimgClass" id="cbimgdiv" clientidmode="Static"
                          src="../../images/plus2.png" border="0" style="cursor: pointer;" />
                      </a>&nbsp;<%=PhraseLib.Lookup("term.customerbroker", LanguageID)%></div>
                  </td>
                </tr>
                <tr id="cberrordiv<%# Eval("RowNum").ToString()%>" class="CBerrorsClass errordetails"
                  style="display: none">
                  <td style="background-color: White" />
                  <td colspan="4">
                    <div>
                      <asp:GridView ID="gvCBErrors" runat="server" AutoGenerateColumns="false" ShowHeader="true"
                        GridLines="None" DataKeyNames="RowNum" Width="100%">
                        <Columns>
                          <asp:BoundField DataField="Severity" />
                          <asp:TemplateField>
                            <ItemTemplate>
                              <a href="javascript:openPopup('../health-resolutions.aspx?FromServerHealth=1&ParamID=<%#Eval("ParamID")%>');"
                                title="<%=PhraseLib.Lookup("store-health.resolution-note", LanguageID)%>">
                                <%#Eval("ParamID")%></a>
                            </ItemTemplate>
                          </asp:TemplateField>
                          <asp:BoundField DataField="Description" />
                          <asp:BoundField DataField="Duration" />
                        </Columns>
                        <RowStyle CssClass="" />
                        <AlternatingRowStyle CssClass="" />
                      </asp:GridView>
                    </div>
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
                filter = "DisconnectedNodes";
                break;

            case "3":
                filter = "CommunicationsOK";
                break;
        } 
        //Load GridView Rows when DIV is scrolled
        var url = $("#hdnURL").val() + "/nodes?&filter=" + filter + "&search=" +encodeURIComponent($("#hdnSearch").val()) + "&sort=" + $("#hdnSortText").val() + "&sortdir=" + $("#hdnSortDir").val() + "&";
        if ($("#hdnFilter").val() == 2 || $("#hdnFilter").val() == 0)
            url = url + "report=true" + "&";
        else
             url = url + "report=all" + "&";
        LoadMoreRecords("dvGrid", "gvNodes", url, "UENodeHealthSummary.aspx/GetNodes", pageIndex, pageSize, pageCount, 'div#loadmoreajaxloader', <%=LanguageID%>,'<%= PhraseLib.Lookup("term.nomorerecords", LanguageID) %>');

        //Function to recieve XML response append rows to GridView
        function OnSuccess(response, gridviewid, loaderdivid) {
            try {
                pageIndex++;

                var nodes = JSON.parse(response.d);

                if (typeof nodes != "string") {
                    var RowNum = (pageIndex-1) * pageSize;

                    $.each(nodes, function (idx, item) {
                        var row0 = $("[id$=" + gridviewid + "] tr").eq(0).clone(true);
                        var row1 = $("[id$=" + gridviewid + "] tr[id$='pblinkdiv1']").clone(true);
                        var row2 = $("[id$=" + gridviewid + "] tr[id$='pberrordiv1']").clone(true);
                        var row4 = $("[id$=" + gridviewid + "] tr[id$='cblinkdiv1']").clone(true);
                        var row5 = $("[id$=" + gridviewid + "] tr[id$='cberrordiv1']").clone(true);

                        RowNum++;
                       
                         row0.attr("id", "row"+RowNum.toString());

                        $(".Node", row0).html("<a title='"+item.nodeName+"' href=\"UENodeHealth.aspx?NodeName=" + item.nodeName + "\">" + item.nodeIp + "</a>");
                        $(".Errors", row0).html(item.errorString);

                        $(".nodeSummaryRow img", row0).attr("id", "imgdiv" + RowNum.toString());                      
                        $(".nodeSummaryRow img", row0).attr("src", "../../images/plus2.png");                      
                        $(".nodeSummaryRow", row0).attr("href", "JavaScript:divexpandcollapse('pblinkdiv"+RowNum.toString()+"','imgdiv"+RowNum.toString()+"','pberrordiv"+RowNum.toString()+"','pbimgdiv"+RowNum.toString()+"');divexpandcollapse('cblinkdiv"+RowNum.toString()+"','imgdiv"+RowNum.toString()+"','cberrordiv"+RowNum.toString()+"','cbimgdiv"+RowNum.toString()+"')");
                      
                        $(".Report", row0).attr("onclick", "javascript:ToggleReport(this,'" + item.nodeName + "','" + $("#hdnURL").val() + "',"+RowNum.toString()+");");
                        $(".Alert", row0).attr("onclick", "javascript:ToggleAlert(this,'" + item.nodeName + "','" + $("#hdnURL").val() + "');");


                        if (item.Report == true)
                            $(".Report", row0).attr("src", "../../images/report-on.png");
                        else
                            $(".Report", row0).attr("src", "../../images/report-off.png");                            

                        if (item.Alert == true)
                            $(".Alert", row0).attr("src", "../../images/email.png");
                        else
                            $(".Alert", row0).attr("src", "../../images/email-off.png");

                      //Promotion Broker
                      row1= populateLink(row1,item.PBerrorHtml,"pbimgdiv" + RowNum.toString(), "pbimgClass", "pblinkdiv" + RowNum.toString(), "pblinkdivClass", "pberrordiv"+RowNum.toString(),'pblink',item.hasPB);
                      row2 = populateErrors(row2, "pberrordiv"+RowNum.toString(),"PBerrorsClass",item.PBerrorHtml);
                      
                        //Customer Broker
                        row4= populateLink(row4,item.CBerrorHtml,"cbimgdiv" + RowNum.toString(), "cbimgClass", "cblinkdiv" + RowNum.toString(), "cblinkdivClass", "cberrordiv"+RowNum.toString(),'cblink',item.hasCB);
                        row5 = populateErrors(row5, "cberrordiv"+RowNum.toString(),"CBerrorsClass",item.CBerrorHtml);
                      
                      
                        $("[id$=" + gridviewid + "]").append(row0);
                         $("[id$=" + gridviewid + "]").append(row1);
                          $("[id$=" + gridviewid + "]").append(row2);
                            $("[id$=" + gridviewid + "]").append(row4);
                             $("[id$=" + gridviewid + "]").append(row5);

                         HighlightRow(RowNum, item.Report,'<%= PhraseLibExtension.PhraseLib.Lookup("term.high", LanguageID)%>','<%= PhraseLibExtension.PhraseLib.Lookup("term.medium", LanguageID)%>','<%= PhraseLibExtension.PhraseLib.Lookup("term.low", LanguageID)%>');

                    });
                } else {
                    $(loaderdivid).html("<center>" + response.d + "<\center>");
                }
            } catch (e) {
                $(loaderdivid).html(e);
            }

       }

       function populateErrors(errorrow,errorDiv,errorsClass,errorHtml){
          errorrow.attr("id", errorDiv);
            errorrow.css('display', 'none');
      
          if(errorHtml != '') {
            errorrow.html("<td style='background-color:White'/><td colspan='4'><div><table style=\"width: 100%; border-collapse: collapse;\">" + errorHtml + "</div></td>");
          }

          return errorrow;
       }

       function populateLink(linkrow,errorHtml,imgDiv, imgClass, linkDiv, linkDivClass,errorDiv,link,show){
         
            $("."+imgClass, linkrow).attr("id", imgDiv);

            linkrow.attr("id", linkDiv);
            linkrow.css('display', 'none');

            if(show)
                 linkrow.attr("show","show");
            else
                linkrow.attr("show","hide");
            
            if (errorHtml == '') {
                $("."+imgClass, linkrow).attr("src", "../../images/plus2-disabled.png");
                $("#"+link, linkrow).attr("href", "");
            } else {
                $("."+imgClass, linkrow).attr("src", "../../images/plus2.png");              
                $("#"+link, linkrow).attr("href", "JavaScript:divexpandcollapse('"+errorDiv + "','"+imgDiv+"')");
            }

            return linkrow;
           
       }

        function divexpandcollapse(divid, imgid, dividOpt, imgidOpt) {
            var attr = $("#" + divid).attr('show');

            if (typeof attr !== typeof undefined && attr == "hide") 
                return;

            if ($("#" + divid).css('display') == "none" && $("#" + divid)) {
                $("#" + divid).css('display', '');
                $("#" + imgid).attr("src", "../../images/minus2.png");
            } else {
                $("#" + divid).css('display', 'none');
                $("#" + imgid).attr("src", "../../images/plus2.png");

                if (dividOpt != '' && $("#" + dividOpt).css('display') != "none")
                    divexpandcollapse(dividOpt, imgidOpt, '', '');
            }
        }

         function ToggleReport(imgReport, hostName, url,rownum) {
            if (imgReport.src.indexOf("report-off.png") > 0 && PostReportAlertDatatoHealthService(url + "/report/nodes/" + hostName, true)) {
                $(imgReport).attr("src","../../images/report-on.png");             
                if ($("#hdnFilter").val() == 1)
                    HighlightRow(rownum, true,'<%= PhraseLibExtension.PhraseLib.Lookup("term.high", LanguageID)%>','<%= PhraseLibExtension.PhraseLib.Lookup("term.medium", LanguageID)%>','<%= PhraseLibExtension.PhraseLib.Lookup("term.low", LanguageID)%>');
            }
            else if(PostReportAlertDatatoHealthService(url + "/report/nodes/" + hostName, false)) {
                if ($("#hdnFilter").val() == 2 || $("#hdnFilter").val() == 0) {
                    $("#dvGrid").scrollTop($("#dvGrid")[0].scrollTop + 10);
                    $("#row" + rownum).css('display', 'none');
                    $("#pblinkdiv" + rownum).css('display', 'none');
                    $("#pberrordiv" + rownum).css('display', 'none');
                    $("#cblinkdiv" + rownum).css('display', 'none');
                    $("#cberrordiv" + rownum).css('display', 'none');
                } else if ($("#hdnFilter").val() == 1)
                   HighlightRow(rownum, false,'<%= PhraseLibExtension.PhraseLib.Lookup("term.high", LanguageID)%>','<%= PhraseLibExtension.PhraseLib.Lookup("term.medium", LanguageID)%>','<%= PhraseLibExtension.PhraseLib.Lookup("term.low", LanguageID)%>');
                
                $(imgReport).attr("src","../../images/report-off.png");
            }
        }

        function ToggleAlert(imgAlert, hostName, url) {
            if (imgAlert.src.indexOf("email-off.png") > 0 && PostReportAlertDatatoHealthService(url + "/alert/nodes/" + hostName, true)) 
                 $(imgAlert).attr("src", "../../images/email.png");            
            else if(PostReportAlertDatatoHealthService(url + "/alert/nodes/" + hostName, false))
                 $(imgAlert).attr("src","../../images/email-off.png");
        }

  </script>
</asp:Content>
