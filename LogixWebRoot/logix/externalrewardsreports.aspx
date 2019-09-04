<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%  ' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
    ' *****************************************************************************
    ' * FILENAME: externalrewardsreports.aspx 
    ' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ' * Copyright © 2002 - 2009.  All rights reserved by:
    ' *
    ' * NCR Corporation
    ' * 1435 Win Hentschel Boulevard
    ' * West Lafayette, IN  47906
    ' * voice: (888) 346-7199 fax: (765) 496-6489
    ' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ' *
    ' * PROJECT : NCR Advanced Marketing Solution
    ' *
    ' * MODULE  : Logix
    ' *
    ' * PURPOSE : 
    ' *
    ' * NOTES   : 
    ' *
    ' * Version : 7.3.1.138972 
    ' *
    ' *****************************************************************************
  
    Dim CopientFileName As String = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
    Dim CopientFileVersion As String = "7.3.1.138972"
    Dim CopientProject As String = "Copient Logix"
    Dim CopientNotes As String = ""
  
    Dim rst As New DataTable
    Dim MyCommon As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    Dim AdminUserID As Long
    Dim i As Integer = 0
    Dim partnerStr As New StringBuilder()
    Dim operationTypeStr As New StringBuilder()
    Dim StyleDownloadBtn As String = "visibility:hidden;"
    
    Dim infoMessage As String = ""
  
    Dim Handheld As Boolean = False
   
    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If
        
    Response.Expires = 0
    MyCommon.AppName = "externalrewardsreports.aspx"
    MyCommon.Open_LogixRT()
    MyCommon.Open_Logix3P()
    
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
    
  
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts(New String() {"datePicker.js"})

    
    MyCommon.QueryStr = "SELECT InternalPartnerId,Name FROM ExternalRewards_InternalPartner with (NoLock) "
    rst = MyCommon.L3P_Select
    If rst.Rows.Count > 0 Then
        For Each row In rst.Rows
            partnerStr.Append("<option value=" & MyCommon.NZ(row.Item("InternalPartnerId"), 0) & "> " & MyCommon.NZ(row.Item("Name"), "0") & " </option>")
        Next
    
    End If
    
    MyCommon.QueryStr = "SELECT OperationTypeChar,OperationName from ExternalRewards_OperationType with (NoLock)"
    rst = MyCommon.L3P_Select
    If rst.Rows.Count > 0 Then
        For Each row In rst.Rows
            operationTypeStr.Append("<option value='" & MyCommon.NZ(row.Item("OperationTypeChar"), "0") & "'> " & MyCommon.NZ(row.Item("OperationName"), "0") & " </option>")
        Next
    
    End If

%>
<script type="text/javascript">
    var divElems = new Array("criteriabody");
    var divVals = new Array(1);
    var divImages = new Array("imgGeneral");

    function showDiv(elemName) {
        var elem = document.getElementById(elemName);

        if (elem != null) {
            elem.style.display = (elem.style.display == "none") ? "block" : "none";
        }
    }

    var datePickerDivID = "datepicker";

    var highlightedRow = -1;
    var highlightedCol = -1;

    if (window.captureEvents) {
        window.captureEvents(Event.CLICK);
        window.onclick = handlePageClick;
    } else {
        document.onclick = handlePageClick;
    }


    function handlePageClick(e) {
        var calFrame = document.getElementById('calendariframe');
        var el = (typeof event !== 'undefined') ? event.srcElement : e.target

        if (el != null) {
            var pickerDiv = document.getElementById(datePickerDivID);
            if (pickerDiv != null && pickerDiv.style.visibility == "visible") {
                if (el.id != "prod-start-picker" && el.id != "prod-end-picker") {
                    if (!isDatePickerControl(el.className)) {
                        pickerDiv.style.visibility = "hidden";
                        pickerDiv.style.display = "none";
                        if (calFrame != null) {
                            calFrame.style.visibility = 'hidden';
                            calFrame.style.display = "none";
                        }
                    }
                } else {
                    pickerDiv.style.visibility = "visible";
                    pickerDiv.style.display = "block";
                    if (calFrame != null) {
                        calFrame.style.visibility = 'visible';
                        calFrame.style.display = "block";
                    }
                }
            }
        }
    }

    function isDatePickerControl(ctrlClass) {
        var retVal = false;

        if (ctrlClass != null && ctrlClass.length >= 2) {
            if (ctrlClass.substring(0, 2) == "dp") {
                retVal = true;
            }
        }

        return retVal;
    }

    function handleFrequency() {
        var frequency = document.getElementById("frequency").value;

        if (frequency == 3) {
            document.getElementById("days").style.display = 'block';
        }
        else {
            var checkboxes = document.getElementsByName('day');
            for (var i = 0, n = checkboxes.length; i < n; i++) {
                if (checkboxes[i].checked) {
                    checkboxes[i].checked = false;
                }
            }
            document.getElementById("days").style.display = 'none';
        }
    }

    function clearForm() {
        var opts = document.getElementById('partners').options;
        for (var i = 0, opt; opt = opts[i]; i++) {
            opt.selected = false;
        }
        var opts = document.getElementById('operationtype').options;
        for (var i = 0, opt; opt = opts[i]; i++) {
            opt.selected = false;
        }

        document.getElementById("frequency").options[0].selected = true;
        document.mainform.reportstart.value = "";
        document.mainform.reportend.value = "";
        var checkboxes = document.getElementsByName('day');
        for (var i = 0, n = checkboxes.length; i < n; i++) {
            if (checkboxes[i].checked) {
                checkboxes[i].checked = false;
            }
        }
        document.getElementById("days").style.display = "none";
        if (document.getElementById("reportlist") != undefined || document.getElementById("reportlist") != null) document.getElementById("reportlist").innerHTML = "";

        document.getElementById("download").disabled = true;

    }

    function handleDownload() {
        var form = document.mainform;

        if (validateForm()) {
            form.exportRpt.value = "1";
            form.action = "#";
            form.method = "post";
            form.submit();
        }
    }

    function validateForm() {
        var startDate = null;
        var endDate = null;
        var isweekdayChked = false;

        if (document.getElementById('reportstart').value != "" && document.getElementById('reportend').value != "") {
            startDate = new Date($('#reportstart').val());
            endDate = new Date($('#reportend').val());
        }

        if (document.getElementsByName('day') != null) {
            var checkboxes = document.getElementsByName('day');
            var vals = "";
            for (var i = 0, n = checkboxes.length; i < n; i++) {
                if (checkboxes[i].checked) {
                    isweekdayChked = true;
                }
            }
        }

        if (document.getElementById('partners').value == null || document.getElementById('partners').value == "") {
            alert('<%Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID))%>' + ' ' + '<%Sendb(Copient.PhraseLib.Lookup("term.partner", LanguageID))%>');
            return false;
        }
        else if (document.getElementById('operationtype').value == null || document.getElementById('operationtype').value == "") {
            alert('<%Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID))%>' + ' ' + '<%Sendb(Copient.PhraseLib.Lookup("term.operation-type", LanguageID))%>');
            return false;
        }
        else if (document.getElementById('reportstart').value == "" || document.getElementById('reportend').value == "") {
            alert('<%Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID))%>' + ' ' + '<%Sendb(Copient.PhraseLib.Lookup("term.daterange", LanguageID))%>');
            return false;
        }
        else if (endDate < startDate) {
            alert('<%Sendb(Copient.PhraseLib.Lookup("externalrewards.error-invalidaterange", LanguageID))%>');
            return false;
        }
        else if (document.mainform.frequency.value == 3 && isweekdayChked == false) {
            alert('<%Sendb(Copient.PhraseLib.Lookup("externalrewards.error-invaliddayofweek", LanguageID))%>');
            return false;
        }
        return true;
    }
    function GenerateExternalRewardsReport() {
        if (validateForm()) {
            xmlhttpPost("ExternalRewardsFeeds.aspx");
        }
    }

    function xmlhttpPost(strURL) {
        var xmlHttpReq = false;
        var self = this;

        document.getElementById("partners").disabled = true;
        document.getElementById("frequency").disabled = true;
        document.getElementById("operationtype").disabled = true;
        document.getElementById("reportstart").disabled = true;
        document.getElementById("reportend").disabled = true;

        // resetRow(highlightedRow);
        highlightedRow = -1;

        document.getElementById("reportlist").innerHTML = "<div class=\"loading\"><br \/><img id=\"clock\" src=\"../images/clock22.png\" \/><br \/>" + '<% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID)) %><\/div>';

        // Mozilla/Safari
        if (window.XMLHttpRequest) {
            self.xmlHttpReq = new XMLHttpRequest();
        }
        // IE
        else if (window.ActiveXObject) {
            self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
        }
        self.xmlHttpReq.open('POST', strURL, true);
        self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
        self.xmlHttpReq.onreadystatechange = function () {
            if (self.xmlHttpReq.readyState == 4) {
                updatepage(self.xmlHttpReq.responseText);
            }
        }
        self.xmlHttpReq.send(getPostData());
    }

    function getPostData() {
        var partners = "";

        var selObj = document.getElementById('partners');
        for (var i = 0; i < selObj.options.length; i++) {
            if (selObj.options[i].selected) {
                partners += "," + selObj.options[i].value;
            }
        }
        if (partners != "") partners = partners.substring(1)

        var operationtype = "";
        selObj = document.getElementById('operationtype');
        for (var i = 0; i < selObj.options.length; i++) {
            if (selObj.options[i].selected) {
                operationtype += ",'" + selObj.options[i].value + "'";
            }
        }
        if (operationtype != "") operationtype = operationtype.substring(1)

        var startDate = document.mainform.reportstart.value;
        var endDate = document.mainform.reportend.value;
        var freq = document.mainform.frequency.value;
        var weekdays = "";

        var checkboxes = document.getElementsByName('day');
        var vals = "";
        for (var i = 0, n = checkboxes.length; i < n; i++) {
            if (checkboxes[i].checked) {
                weekdays += ",'" + checkboxes[i].value + "'";
            }
        }
        if (weekdays != null && weekdays != "")
            weekdays = weekdays.substring(1);

        qstr = 'GenerateExternalRewardsReport=1&partner=' + escape(partners); // NOTE: no '?' before querystring
        qstr += '&operationtype=' + escape(operationtype);
        qstr += "&startdate=" + escape(startDate);
        qstr += "&enddate=" + escape(endDate);
        qstr += '&frequency=' + escape(freq);
        qstr += '&day=' + escape(weekdays);

        return qstr;
    }

    function updatepage(str) {
        
        var nodata = '<%Sendb(Copient.PhraseLib.Lookup("reports.nodata", LanguageID))%>';

        if (str != null) {
            document.getElementById("reportlist").innerHTML = str;
            var n = str.search(nodata);
            if (n == -1)
                document.getElementById("download").disabled = false;
            else
                document.getElementById("download").disabled = true;

        }
        document.getElementById("partners").disabled = false;
        document.getElementById("frequency").disabled = false;
        document.getElementById("operationtype").disabled = false;
        document.getElementById("reportstart").disabled = false;
        document.getElementById("reportend").disabled = false;

    }

    
</script>
<%
    Dim dtResult As DataTable = Nothing
    

    Send_HeadEnd()
    Send_BodyBegin(1)
    Send_Bar(Handheld)
    Send_Help(CopientFileName)
    Send_Logos()
    Send_Tabs(Logix, 8)
    Send_Subtabs(Logix, 8, 10)
  
  
    If (Logix.UserRoles.AccessReports = False) Then
        Send_Denied(1, "perm.admin-reports")
        GoTo done
    End If
  
    ' Do not initially run the query and hide the results table
    'Write the data as a CSV file for download
    If (Request.Form("exportRpt") = "1") Then
        If Session("RptDataTable") IsNot Nothing Then
            dtResult = DirectCast(Session("RptDataTable"), DataTable)
        End If
        Response.ClearHeaders()
        Response.AddHeader("Content-Disposition", "attachment; filename=ExternalRewardsRpt" & Now() & ".csv")
        Response.ContentType = "application/octet-stream"
        Response.Clear()
        'Write the column headers
        If (dtResult.Rows.Count > 0) Then
            Dim columnCnt As Integer = dtResult.Columns.Count
            For Each column As DataColumn In dtResult.Columns
                Sendb(column.ColumnName & ",")
            Next
            Send("")
            'Write the data rows
            i = 0
            For Each row In dtResult.Rows
                For index As Integer = 0 To columnCnt - 1
                    Sendb(MyCommon.NZ(row.Item(index), "") & ",")
                Next
                Send("")
                i = i + 1
            Next
            Response.Flush()
            Response.End()
        End If
    End If
%>
<div id="intro">
    <h1 id="title">
        <% Sendb(Copient.PhraseLib.Lookup("term.externalrewardsreports", LanguageID))%>
    </h1>
    <div id="controls">
        <input type="button" class="regular" id="download" name="download" value="<% Sendb(Copient.PhraseLib.Lookup("term.download", LanguageID)) %>"
            disabled="disabled" title="<% Sendb(Copient.PhraseLib.Lookup("report-custom-downloadreport", LanguageID)) %>"
            onclick="handleDownload();" />
        <input class="regular" type="button" id="btnClear" name="btnClear" value="<% Sendb(Copient.PhraseLib.Lookup("term.reset", LanguageID))%>"
            onclick="clearForm();" style="width: 70px;" />
        <input class="regular" type="button" id="generateReport" name="generateReport" value="<% Sendb(Copient.PhraseLib.Lookup("term.GetResults", LanguageID))%>"
            onclick="javascript:GenerateExternalRewardsReport();" />
    </div>
</div>
<div id="main">
    <div id="infobar" class="red-background" style="display: none">
    </div>
    <div class="box" id="criteria">
        <h2 style="float: left;">
            <span>
                <% Sendb(Copient.PhraseLib.Lookup("term.criteria", LanguageID))%>
            </span>
        </h2>
        <% Send_BoxResizer("criteriabody", "imgGeneral", Copient.PhraseLib.Lookup("term.criteria", LanguageID), True)%>
        <div id="criteriabody">
            <form action="#" name="mainform" id="mainform" method="post">
            <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.criteria", LanguageID)) %>">
                <tr>
                    <td>
                        <b>
                            <% Sendb(Copient.PhraseLib.Lookup("term.partner", LanguageID))%></b>
                    </td>
                    <td>
                        <select id="partners" name="partners" multiple="multiple" style="height: 125px; width: 125px;"
                            size="5">
                            <% Send(partnerStr.ToString())%>
                        </select>
                    </td>
                    <td>
                        <b>
                            <% Sendb(Copient.PhraseLib.Lookup("term.operation", LanguageID))%>
                        </b>
                    </td>
                    <td>
                        <select id="operationtype" name="operationtype" multiple="multiple" style="height: 125px;
                            width: 125px;" size="5">
                            <% Send(operationTypeStr.ToString())%>
                        </select>
                    </td>
                    <td style="vertical-align: top;">
                        <div id="Div1" style="width: 200px">
                        </div>
                    </td>
                </tr>
                <tr>
                    <td>
                        <b>
                            <label for="date">
                                <% Sendb(Copient.PhraseLib.Lookup("term.date", LanguageID))%>
                            </label>
                        </b>
                    </td>
                    <td colspan="3">
                        <div id="datepicker" class="dpDiv">
                        </div>
                        <%
                            If Request.Browser.Type = "IE6" Then
                                Send("<iframe src=""javascript:'';"" id=""calendariframe"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no"" style=""visibility:hidden;display:none;""></iframe>")
                            End If
                        %>
                        <input type="text" class="short" id="reportstart" name="reportstart" maxlength="10"
                            value="" />
                        <img src="../images/calendar.png" class="calendar" id="prod-start-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
                            title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
                            onclick="displayDatePicker('reportstart', event);" />
                        <b>
                            <% Sendb(Copient.PhraseLib.Lookup("term.to", LanguageID))%>
                        </b>
                        <input type="text" class="short" id="reportend" name="reportend" maxlength="10" value="" />
                        <img src="../images/calendar.png" class="calendar" id="prod-end-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
                            title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
                            onclick="displayDatePicker('reportend', event);" />
                   &nbsp;
                        <b>
                            <% Sendb(Convert.ToString(Copient.PhraseLib.Lookup("term.by", LanguageID)))%>
                        </b>
                   &nbsp;&nbsp;
                        <select id="frequency" name="frequency" onchange="javascript:handleFrequency();">
                            <option value="0">
                                 <% Sendb(Copient.PhraseLib.Lookup("term.none", LanguageID))%>
                            </option>
                            <option value="1">
                                <% Sendb(Copient.PhraseLib.Lookup("term.week", LanguageID))%>
                            </option>
                            <option value="2">
                                <% Sendb(Copient.PhraseLib.Lookup("term.month", LanguageID))%>
                            </option>
                            <option value="3">
                                <% Sendb(Copient.PhraseLib.Lookup("term.day", LanguageID) & " " & Copient.PhraseLib.Lookup("term.of", LanguageID) & " " & Copient.PhraseLib.Lookup("term.week", LanguageID))%>
                            </option>
                        </select>
                    </td>
                    <td>
                        <div id="days" style="display: none">
                            <input type="checkbox" id="sunday" name="day" value="sunday" /><label for="sunday"><% Sendb(Convert.ToString(Copient.PhraseLib.Lookup("term.sunday", LanguageID)).Substring(0, 3))%></label>
                            <input type="checkbox" id="monday" name="day" value="monday" /><label for="monday"><% Sendb(Convert.ToString(Copient.PhraseLib.Lookup("term.monday", LanguageID)).Substring(0, 3))%></label>
                            <input type="checkbox" id="tuesday" name="day" value="tuesday" /><label for="tuesday"><% Sendb(Convert.ToString(Copient.PhraseLib.Lookup("term.tuesday", LanguageID)).Substring(0, 3))%></label>
                            <input type="checkbox" id="wednesday" name="day" value="wednesday" /><label for="wednesday"><% Sendb(Convert.ToString(Copient.PhraseLib.Lookup("term.wednesday", LanguageID)).Substring(0, 3))%></label>
                            <br />
                            <input type="checkbox" id="Thursday" name="day" value="Thursday" /><label for="Thursday"><% Sendb(Convert.ToString(Copient.PhraseLib.Lookup("term.thursday", LanguageID)).Substring(0, 3))%></label>
                            <input type="checkbox" id="friday" name="day" value="friday" /><label for="friday"><% Sendb(Convert.ToString(Copient.PhraseLib.Lookup("term.friday", LanguageID)).Substring(0, 3))%></label>
                            <input type="checkbox" id="saturday" name="day" value="saturday" /><label for="saturday"><% Sendb(Convert.ToString(Copient.PhraseLib.Lookup("term.saturday", LanguageID)).Substring(0, 3))%></label>
                        </div>
                    </td>
                </tr>
              
            </table>
            <input type="hidden" name="exportRpt" id="exportRpt" value="0" />
            </form>
        </div>
    </div>
    <div id="gutter">
    </div>
    <div class="box" id="results" style="height: 400px;">
        <h2>
            <% Sendb(Copient.PhraseLib.Lookup("term.results", LanguageID))%>
        </h2>
        <div style="height: 300px; overflow: auto;" id="reportlist" class="reportclass">
        </div>
    </div>
</div>
<%
done:
    
    MyCommon.Close_LogixRT()
    MyCommon.Close_Logix3P()
    Logix = Nothing
    MyCommon = Nothing
           
    If Session("RptDataTable") IsNot Nothing Then
        Session.Remove("RptDataTable")
    End If
%>
