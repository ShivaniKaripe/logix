<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
    ' *****************************************************************************
    ' * FILENAME: pgroup-inquiry.aspx 
    ' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ' * Copyright © 2002 - 2009.  All rights reserved by:
    ' *
    ' * NCR Corporation
    ' * 1435 Win Hentschel Boulevard
    ' * West Lafayette, IN  47906
    ' * voice: 888-346-7199  fax: 765-464-1369
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
  
    Dim AdminUserID As Long
    Dim MyCommon As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    Dim rst As DataTable
    Dim row As DataRow
    Dim x As Integer
    Dim y As Integer
    Dim result As Boolean
    Dim shaded As Boolean
    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False
    Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249)="1",True,False)
  
    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If
  
    Response.Expires = 0
    MyCommon.AppName = "pgroup-inquiry.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
    shaded = True
    result = False
  
  
    Send_HeadBegin("term.group")
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts()
    Send_HeadEnd()
    Send_BodyBegin(1)
%>
<script type="text/javascript" src="../../javascript/jquery-1.10.2.min.js"></script>
<script type="text/javascript" language="javascript">
    var PageStart = 0;
    var PageEnd = 0;
    var DefaultPageSize = 5000
    var IsEngineInstalled = '<%=MyCommon.IsEngineInstalled(9)%>' == 'True' ? true : false
    var Fetch_UE_SystemOption_168 = '<%MyCommon.Fetch_UE_SystemOption(168)%>'

    window.onload = GetProductGroupsList();
     
function xmlhttpPost(strURL) {
    var xmlHttpReq = false;
    var self = this;
    document.getElementById("GroupSelector").disabled = true
    document.getElementById("butt").disabled = true
    document.getElementById("results").innerHTML = "<div class=\"loading\"><br \/><img id=\"clock\" src=\"../images/clock22.png\" \/><br \/>" + '<% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID)) %><\/div>';
    
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
    self.xmlHttpReq.onreadystatechange = function() {
        if (self.xmlHttpReq.readyState == 4) {
            updatepage(self.xmlHttpReq.responseText);
        }
    }
    self.xmlHttpReq.send(getquerystring());
}

function xmlhttpPost2(strURL,value) {
    var xmlHttpReq = false;
    var self = this;
    // document.getElementById("GroupSelector").disabled = true
    // document.getElementById("butt").disabled = true
    document.getElementById("results").innerHTML = "<div class=\"loading\"><br /><img id=\"clock\" src=\"../images/clock22.png\" \/><br \/>" + '<% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID)) %><\/div>';
    
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
    self.xmlHttpReq.onreadystatechange = function() {
        if (self.xmlHttpReq.readyState == 4) {
            updatepage2(self.xmlHttpReq.responseText);
        }
    }
    //alert('CollisionGroup=' + escape(value) + '&EngineID=' + document.getElementById("EngineID").value);
    self.xmlHttpReq.send('CollisionGroup=' + escape(value) + '&EngineID=' + document.getElementById("EngineID").value);
}

function getquerystring() {
    var form = document.forms['f1'];
    var word = form.GroupSelector.value;
    qstr = 'CollisionGroup=' + escape(word);  // NOTE: no '?' before querystring
    
    return qstr;
}

function updatepage(str){
    document.getElementById("results").innerHTML = str;
    document.getElementById("GroupSelector").disabled = false
    document.getElementById("butt").disabled = false
    // document.getElementById("results").innerHTML = "Testing";
}

function updatepage2(str){
    document.getElementById("results").innerHTML = str;
           document.getElementById("functioninput").disabled = false;
           document.getElementById("functionselect").disabled = false;
    document.getElementById("butt2").disabled = false;
    // document.getElementById("GroupSelector").disabled = false
    // document.getElementById("butt").disabled = false
    // document.getElementById("results").innerHTML = "Testing";
}


// This is the function that refreshes the list after a keypress.
// The maximum number to show can be limited to improve performance with
// huge lists (1000s of entries).
// The function clears the list, and then does a linear search through the
// globally defined array and adds the matches back to the list.
function handleKeyUp() {
    if ($("#functioninput").val() != "") {
        var params = "mi=PRODUCT_GROUPS_INQUIRY&PageStart=1&PageEnd=999999&SearchText=" + $("#functioninput").val() + "&SearchType=" + ($("#functionradio1").is(':checked') ? "1" : "2")
        Ajax_GetProductGroupsList(params);
    }
    else {
        GetProductGroupsList()
    }
}

// this function gets the selected value and loads the appropriate
// php reference page in the display frame
// it can be modified to perform whatever action is needed, or nothing
function handleSelectClick() {
    selectObj = document.forms[0].functionselect;
    textObj = document.forms[0].functioninput;
    
    //selectedValue = selectObj.options[selectObj.selectedIndex].text;
    //selectedValue = selectedValue.replace(/_/g, '-') ;
    selectedValue = document.getElementById("functionselect").value;
    if(selectedValue != "") {
        document.getElementById("functionselect").size = "5";
        document.getElementById("functionselect").disabled = true;
        document.getElementById("functioninput").disabled = true;
        document.getElementById("butt2").disabled = true;
        //document.location.href = "XMLFeeds.aspx?CollisionGroup=" + selectedValue + "";
        xmlhttpPost2("XMLFeeds.aspx",selectedValue);
    }
}

function productsGroupsList_Prev(){
    if (PageStart != 0 ) {

        PageStart = PageStart - parseInt(DefaultPageSize);

        PageEnd = PageStart + parseInt(DefaultPageSize);

            GetProductGroupsList();
        }

        return false;
    }

    function productsGroupsList_Next(){
        if (PageStart == 0) {
            PageStart = parseInt(DefaultPageSize);
        }
        else{
            PageStart = PageEnd;
        }

        PageEnd = PageStart + parseInt(DefaultPageSize);

        GetProductGroupsList();

        return false;
    }

    function GetProductGroupsList() {
        if(PageEnd == 0) { PageStart =0; PageEnd = DefaultPageSize }

        var params = "mi=PRODUCT_GROUPS_INQUIRY&PageStart=" + (PageStart + 1) + "&PageEnd=" + PageEnd + "&SearchText=&SearchType="

        Ajax_GetProductGroupsList(params);
    }
    
    function Ajax_GetProductGroupsList(params) {
        var url = "\UE\\UEoffer-rew-discount.aspx";
        $("#functionselect").html("<option>Loading...</option>")

        var objXMLHttpRequest = new XMLHttpRequest();
        objXMLHttpRequest.onreadystatechange = function () {
            if (objXMLHttpRequest.readyState == 4 && objXMLHttpRequest.status == 200) {
                var arrSplitResult = objXMLHttpRequest.responseText.split("_AMS_SPLITTER_AMS_");
                if (arrSplitResult[0] == 'T') {
                    try {
                        window.objJsonTable = JSON.parse(arrSplitResult[1]);
                        window.objStrHTML = "";

                        if ((window.objJsonTable.length <= 0 && PageEnd == DefaultPageSize) || (PageStart == 0)) {
                            $("#lnkPrev").hide();
                        }
                        else {
                            $("#lnkPrev").show();
                        }
                        if (window.objJsonTable.length < (DefaultPageSize-1)){
                            $("#lnkNext").hide();
                        }
                        else {
                            $("#lnkNext").show();
                        }
                        $("#functionselect").html("")
                        $.each(window.objJsonTable, function(index, tRow) {
                            
                            if(tRow.ProductGroupID == null){
                                tRow.ProductGroupID = 0;
                            }
                            //window.objStrHTML += "<option name='" + tRow.Name + "' value='" + tRow.ProductGroupID + "'>" + tRow.Name + "</option>"
                            
                            if (IsEngineInstalled && Fetch_UE_SystemOption_168 == "1" && tRow.externalbuyerid != null) {
                                window.objStrHTML += "<option value='" + tRow.ProductGroupID + "'>" + tRow.Name + " - " & "Buyer " + tRow.externalbuyerid + " - " + tRow.Name + "</option>"
                            }
                        else{
                                window.objStrHTML += "<option value='" + tRow.ProductGroupID + "'>" + tRow.ProductGroupID + " - " + tRow.Name + "</option>"
                        }
                        });
                        
                    $("#functionselect").html(window.objStrHTML);
                    
                    delete window.objJsonTable;
                    delete window.objStrHTML;
                }
                catch (e) {
                }
            }
            else {
                    
            }
        };
        }
    objXMLHttpRequest.open("POST", url, true);
    objXMLHttpRequest.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
    objXMLHttpRequest.send(params);
}



</script>
<%
    Send_Bar(Handheld)
    Send_Help(CopientFileName)
    Send_Logos()
    Send_Tabs(Logix, 4)
    Send_Subtabs(Logix, 40, 3)
  
    If (Logix.UserRoles.AccessGroupInquiry = False) Then
        Send_Denied(1, "perm.pgroup-groupinquiry")
        GoTo done
    End If
  
    ' check engine to make sure we shoud allow this
    ' MyCommon.QueryStr = "select EngineID,DefaultEngine,Installed from PromoEngines with (NoLock) where EngineID=2 and DefaultEngine=1"
    ' rst = MyCommon.LRT_Select
  
    ' If (rst.Rows.Count > 0) Then
%>
<!--
  <div id="main">
    Not Supported for CPE
  </div>
</div>
-->
<%      
    'GoTo done
    'End If
%>
<div id="intro">
    <h1 id="title">
        <% Sendb(Copient.PhraseLib.Lookup("term.productgroupinquiry", LanguageID))%>
    </h1>
    <div id="controls">
        <%
            'If MyCommon.Fetch_SystemOption(75) Then
            '  If (Logix.UserRoles.AccessNotes) Then
            '    Send_NotesButton(7, 0, AdminUserID)
            '  End If
            'End If
        %>
    </div>
</div>
<div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <% Sendb(Copient.PhraseLib.Lookup("pgroup-inquiry.main", LanguageID))%>
    <br />
    <br class="half" />
    <form onsubmit="handleSelectClick();return false;" action="" id="mainform" name="mainform">
    <% 
        MyCommon.QueryStr = "select EngineID,Description,PhraseID,DefaultEngine from PromoEngines with (NoLock) where Installed=1 and (EngineID<3 or EngineID=9);"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
            Send("<label for=""EngineID"">" & Copient.PhraseLib.Lookup("term.engine", LanguageID) & ":</label>")
            Send("<select id=""EngineID"" name=""EngineID"">")
            For Each row In rst.Rows
                Sendb("  <option value=""" & row.Item("EngineID") & """" & IIf(row.Item("DefaultEngine") = 1, " selected=""selected""", "") & ">")
                If MyCommon.NZ(row.Item("PhraseID"), 0) > 0 Then
                    Sendb(Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID))
                Else
                    Sendb(row.Item("Description"))
                End If
                Send("</option>")
            Next
            Send("</select>")
            Send("<br />")
            Send("<br class=""half"" />")
        Else
            Send("<input type=""hidden"" id=""EngineID"" name=""EngineID"" value=""0"" />")
        End If
    %>
    <div style="float: left;position: relative; width: 108%; height: 80%;"">
   
        <div class="column3x" id="selector" style="overflow: auto;">
    <input type="radio" id="functionradio1" name="functionradio" <% if(MyCommon.Fetch_SystemOption(175)= "1" ) then sendb(" checked=""checked""") %> /><label
        for="functionradio1"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
    &nbsp;
    <input type="radio" id="functionradio2" name="functionradio" <% if(MyCommon.Fetch_SystemOption(175)= "2" ) then sendb(" checked=""checked""") %> /><label
        for="functionradio2"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
    <input type="text" id="functioninput" name="functioninput" onkeyup="handleKeyUp();"
        maxlength="100" style="font-size: 10pt; width: 34ex;" value="" /><br />
    <div id="pgList" style="width:410px">
        <a id="lnkPrev" style="float:left;display:none;" href="#" onclick="productsGroupsList_Prev()"><< Prev</a>
        <a id="lnkNext" style=" float:right; margin-right:10px;display:none;" href="#" onclick="productsGroupsList_Next()">Next >></a>
        <br />
        <select onclick="handleSelectClick();" id="functionselect" name="functionselect"
        size="5" style="font-size: 10pt; width: 400px;">
       
    </select>
    </div>
    </div>
        </div>
    <br />
    <input type="button" id="butt2" name="butt2" onclick="handleKeyUp();" value="<% Sendb(Copient.PhraseLib.Lookup("pgroup-inquiry.load", LanguageID))%>" />
    </form>
    <div id="results">
    </div>
</div>
<%-- 
<% If Request.QueryString("CollisionGroup") <> "" Then %>
<script type="text/javascript" language="javascript">
    xmlhttpPost("XMLFeeds.aspx", "<% sendb(Request.QueryString("CollisionGroup")) %>" )
</script>
<% End If%>
--%>
<%
    'If MyCommon.Fetch_SystemOption(75) Then
    '  If (Logix.UserRoles.AccessNotes) Then
    '    Send_Notes(7, 0, AdminUserID)
    '  End If
    'End If
done:
    Send_BodyEnd("mainform", "functioninput")
    MyCommon.Close_LogixRT()
    Logix = Nothing
    MyCommon = Nothing
%>
