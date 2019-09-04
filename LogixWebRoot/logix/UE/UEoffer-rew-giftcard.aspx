﻿<%@ Page Language="C#" AutoEventWireup="true" CodeFile="UEoffer-rew-giftcard.aspx.cs"
    Inherits="logix_UE_UEoffer_rew_giftcard" %>

<%--<%@ Register Src="~/logix/UserControls/MultiLanguagePopup.ascx" TagName="MLI" TagPrefix="uc" %>--%>
<%@ Reference Control="~/logix/UserControls/MultiLanguagePopup.ascx" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <base target="_self" />
    <title></title>
    <script type="text/javascript" src="/javascript/logix.js"></script>
    <script type="text/javascript" src="/javascript/jquery.min.js"></script>
    <script type="text/javascript" src="/javascript/jquery-ui-1.10.3/jquery-ui-min.js"></script>
    <link rel="stylesheet" href="/javascript/jquery-ui-1.10.3/style/jquery-ui.min.css" />
    <style type="text/css">
        .ui-dialog .ui-dialog-content {
            background-color: #c0c0c0;
            border: solid 1px #dddddd;
        }

        .ui-dialog-titlebar {
            background: #0066ff;
            color: #ffffff;
            font-weight: bold;
            line-height: 20px;
        }

        #vsError {
            display: none;
            visibility: hidden;
            width: 655px;
        }

        .table-div {
            margin-left: 60px;
            margin-top: 30px;
            border-style: solid;
            border-color: Green;
        }
    </style>
</head>
<body class="<%=objOffer.IsTemplate? "popup template" : "popup"%>" onunload="ChangeParentDocument();">
    <script language="javascript" type="text/javascript">

        function ChangeParentDocument() {
    <%= GetRefreshScript() %>

        }
        function EnableControls() {
            if (Page_ClientValidate() && Page_IsValid) {
                //Removing the disabled attribute of the controls as there values are not accessible while saving on server side if disabled attribute is set.
                $("#message *").each(function () {
                    //alert($(this).attr("id"));
                    if ($(this).attr("disabled")) {
                        $(this).removeAttr("disabled");
                    }
                });
            }
        }
        function ValidatePage() {
            Page_ClientValidate();
            if (Page_IsValid) {
                document.getElementById("vsError").style.display = "none";
                document.getElementById("vsError").style.visibility = "hidden";
            }
            else {
                document.getElementById("vsError").style.display = "block";
                document.getElementById("vsError").style.visibility = "visible";
            }
        }


        function CheckNumeric(e) {
            var NumberDecimalSeparator = '<%= CurrentUser.AdminUser.Culture.NumberFormat.NumberDecimalSeparator %>';
            var decimalkeyCode = NumberDecimalSeparator.charCodeAt(0);
            if (window.event) // IE 
            {
                if ((e.keyCode < 48 || e.keyCode > 57) & e.keyCode != 8 & e.keyCode != decimalkeyCode) {
                    event.returnValue = false;
                    return false;
                }
            }
            else { // Fire Fox
                if ((e.which < 48 || e.which > 57) & e.which != 8 & e.which != decimalkeyCode) {
                    e.preventDefault();
                    return false;
                }
            }
        }

        //For checking alpha numeric
        function CheckAlphaNumeric(key) {
            var keycode = (key.which) ? key.which : key.keyCode
            if (!(keycode == 8 || keycode == 46) && (keycode < 48 || keycode > 57) && (keycode < 65 || keycode > 90) && (keycode < 97 || keycode > 122)) {
                (key.which) ? key.preventDefault() : event.returnValue = false;
                return false;
            }
        }

        //Template related scripts
        function xmlhttpPost(strURL) {
            var xmlHttpReq = false;
            var self = this;

            document.getElementById("results").innerHTML = "<div class=\"loading\"><img id=\"clock\" src=\"../../images/clock22.png\" \/><br \/>" + '<\/div>';

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
            self.xmlHttpReq.send("LanguageID=" + $('#language').val());

            SetPageLockStatus();
            return false;
        }


        function updatepage(str) {
            document.getElementById("results").innerHTML = str;
        }
        function ShowFieldList() {
            var elemList = document.getElementById("templatefields");

            if (elemList != null) {
                if (elemList.style.display == 'block') {
                    elemList.style.display = 'none';
                } else {
                    elemList.style.display = 'block';
                }
            }
        }

        function updateTableLockStatus() {
            var elemTbl = document.getElementById("tblTempFields");
            var elemTr = null, elemTd = null, elemChk = null;

            SetPageLockStatus();

            if (elemTbl != null) {
                elemTrs = elemTbl.getElementsByTagName("TR");
                for (var i = 1; i < elemTrs.length; i++) {
                    elemTd = elemTrs[i].firstChild.nextSibling;
                    if (elemTd != null) {
                        updateLockStatus(elemTd.firstChild);
                    }
                }
            }
        }


        function updateLockStatus(elem) {
            var pageElem = document.mainform.chkDisallow_Edit;
            if (pageElem != null) {
                var pageLockChecked = pageElem.checked;
                if (elem != null) {
                    var td1 = elem.parentNode;
                    if (td1 != null) {
                        var tr = td1.parentNode;
                        if (tr != null) {
                            //var td3 = tr.lastChild;
                            var td3 = $(tr).children('td:last');
                            if (td3 != null) {
                                if (pageLockChecked) {
                                    td3.html((td3.html() == 'Locked') ? 'Unlocked' : 'Locked');
                                } else {
                                    td3.html((elem.checked) ? 'Locked' : 'Unlocked');
                                }

                                if (elem.checked)
                                    UpdateLockedFieldsStatus(elem.value, "add");
                                else
                                    UpdateLockedFieldsStatus(elem.value, "remove");

                            }
                        }
                    }
                }
            }
        }

        var lockedFields = new Array();

        $(document).ready(function () {
            updateHiddenFieldValue();
        });

        function updateHiddenFieldValue() {
            if ($("#hdnLockedTemplateFields").val() != "") {
                $.merge(lockedFields, ($("#hdnLockedTemplateFields").val().split(",")));
            }
        }

        function UpdateLockedFieldsStatus(value, action) {
            if (action == "add" && lockedFields.indexOf(value) == -1)
                lockedFields.push(value);
            else if (action == "remove" && lockedFields.indexOf(value) != -1) {
                lockedFields.splice(lockedFields.indexOf(value), 1);
            }

            $("#hdnLockedTemplateFields").val(lockedFields);
        }

        function SetPageLockStatus() {
            if ($("#chkDisallow_Edit").is(":checked") == true)
                $("#hdnPageLock").val("Locked");
            else if ($("#chkDisallow_Edit").is(":checked") == false)
                $("#hdnPageLock").val("Unlocked");
        }
//Template related scripts
    </script>
    <form id="mainform" runat="server">
        <input type="hidden" id="PercentError" name="PercentErrorString" value='<%=PhraseLib.Lookup("error.GC-percent-over", LanguageID)%>' />
        <input type="hidden" id="ValueError" name="ValueErrorString" value='<%=PhraseLib.Lookup("error.GC-value-over", LanguageID)%>' />
        <input type="hidden" id="hdnLockedTemplateFields" value="" runat="server" />
        <input type="hidden" id="hdnPageLock" value="Unlocked" runat="server" />
        <input type="hidden" id="language" value="<%=LanguageID%>" />
        <div id="results" style="position: absolute; z-index: 99; top: 31px; right: 21px;">
        </div>
        <div id="custom1">
        </div>
        <div id="wrap">
            <div id="custom2">
            </div>
            <div id="intro">
                <h1 id='title' runat="server">Title</h1>
                <div id='controls'>
                    <span class="temp" id="TempDisallow" runat="server">
                        <asp:CheckBox ID="chkDisallow_Edit" runat="server" CssClass="tempcheck" onclick="updateTableLockStatus();" />
                        <label for="chkDisallow_Edit">
                            <%=PhraseLib.Lookup("term.locked", LanguageID)%>
                        </label>
                        <a href="javascript:ShowFieldList();" title='<%=PhraseLib.Lookup("cpeoffer-rew-disc-clicktoview", LanguageID)%>'>&#9660;</a> </span>
                    <asp:Button AccessKey="s" CssClass="regular" ID="btnSave" runat="server" Text=""
                        OnClientClick="ValidatePage();EnableControls();" Visible="true" OnClick="btnSave_Click" ValidationGroup="btnSaveGroup" />
                </div>
            </div>
            <div id="main">
                <asp:Label ID="vsError" ForeColor="" CssClass="errsummary red-background" runat="server" />
                <div id="infobar" class="red-background" runat="server" visible="false">
                </div>
                <div id="column4">
                    <div class="box" id="message">
                        <h2>
                            <span>
                                <%= PhraseLib.Lookup("term.data", LanguageID)%>
                            </span>
                        </h2>
                        <div style="height: 10%; margin-bottom: 5px;">
                            <div class="table-div">
                                <table>
                                    <tr style="padding-bottom: 10px; height: 30px;">
                                        <td style="width: 128px">
                                            <asp:Label ID="lblValueType" runat="server" Text="ValueType:" />
                                        </td>
                                        <td>
                                            <asp:DropDownList ID="ddlValueType" runat="server" Width="160px" AutoPostBack="true" OnSelectedIndexChanged="ddlValueType_OnSelectedIndexChanged"
                                                onchange="ManageControl(this.value);" />
                                        </td>
                                    </tr>
                                    <tr style="padding-bottom: 10px; height: 30px;">
                                        <td>
                                            <asp:Label ID="lblProrationRate" runat="server"><%= PhraseLib.Lookup("term.prorationtype", LanguageID)%>:</asp:Label>
                                        </td>
                                        <td>
                                            <asp:DropDownList ID="ddlProrationRate" runat="server" Width="160px" />
                                        </td>
                                    </tr>
                                    <tr style="padding-bottom: 10px; height: 30px;">
                                        <td>
                                            <asp:Label ID="lblNameOfCard" runat="server"><%= PhraseLib.Lookup("term.gc-name", LanguageID)%>:</asp:Label>
                                        </td>
                                        <td>
                                            <asp:TextBox ID="txtNameOfCard" runat="server" class="middle" MaxLength="30" Text="" />
                                        </td>
                                        <td>
                                            <asp:RequiredFieldValidator runat="server" ID="requirefieldNameOfCard" ControlToValidate="txtNameOfCard"
                                                SetFocusOnError="true" />
                                        </td>
                                    </tr>
                                    <tr style="padding-bottom: 10px; height: 30px;">
                                        <td>
                                            <asp:Label ID="lblChargeBack" runat="server"><%= PhraseLib.Lookup("term.gc-chargeback", LanguageID)%>:</asp:Label>
                                        </td>
                                        <td>
                                            <asp:DropDownList ID="ddlChargeBack" runat="server" Width="160px" />
                                            <asp:RequiredFieldValidator ID="RFV_ddlChargeBack" runat="server" InitialValue="" ControlToValidate="ddlChargeBack"
                                                SetFocusOnError="true" Display="Dynamic" ValidationGroup="btnSaveGroup">
                                                <%=PhraseLib.Lookup("term.chargebackDeptValidation", LanguageID)%>
                                            </asp:RequiredFieldValidator>
                                        </td>
                                    </tr>
                                    <tr style="padding-bottom: 10px; height: 30px;">
                                        <td>
                                            <asp:Label ID="lblCardIdentifier" runat="server"><%= PhraseLib.Lookup("term.gc-cardidentifier", LanguageID)%>:</asp:Label>
                                        </td>
                                        <td>
                                            <asp:TextBox ID="txtCardIdentifier" runat="server" class="middle" MaxLength="30"
                                                Text="" onkeypress="CheckAlphaNumeric(event);" />
                                        </td>
                                        <td>
                                            <asp:RegularExpressionValidator ID="IdentiferValidator" runat="server" ControlToValidate="txtCardIdentifier"
                                                ValidationExpression="^[a-zA-Z0-9]+$" SetFocusOnError="true" />
                                        </td>
                                    </tr>
                                </table>
                            </div>
                        </div>
                        <asp:Repeater ID="repGiftcard" runat="server" OnItemDataBound="repGiftcard_ItemDataBound"
                            OnItemCreated="repGiftcard_ItemCreated">
                            <ItemTemplate>
                                <label style="font-weight: bold; margin-left: 20px;">
                                    <%#(objOffer.NumbersOfTier>1?PhraseLib.Lookup("term.tier", LanguageID):"" ) %>
                                    <%#(objOffer.NumbersOfTier > 1 ? DataBinder.Eval(Container.DataItem, "TierLevel") + ":" : "")%></label>
                                <div style="height: 10%;">
                                    <div style="margin-left: 60px">
                                        <table>

                                            <tr>
                                                <td style="width: 128px">
                                                    <asp:Label ID="lblValue" runat="server" Style="float: left"><%= PhraseLib.Lookup("term.value", LanguageID)%>:</asp:Label>
                                                    <asp:Label ID="VTypeStart" runat="server" Style="float: right"><%= ValueTypeInitialPrefix %></asp:Label>
                                                </td>
                                                <td style="width: 275px">
                                                    <asp:TextBox ID="txtValue" runat="server" class="middle" MaxLength="9" Text='<%#((Decimal)DataBinder.Eval(Container.DataItem, "Amount")).ToString(CurrentUser.AdminUser.Culture) %>'
                                                        onkeypress="CheckNumeric(event);" />
                                                    <asp:Label ID="VTypeEnd" runat="server"><%= ValueTypeInitialSuffix %></asp:Label>
                                                </td>
                                                <td>
                                                    <asp:RequiredFieldValidator runat="server" ID="requirefieldValue" ControlToValidate="txtValue"
                                                        SetFocusOnError="true" Display="Dynamic" />
                                                    <asp:CustomValidator ID="customValidator" ControlToValidate="txtValue" Display="Dynamic"
                                                        SetFocusOnError="true" ErrorMessage="Invalid Value!" ForeColor="red" ClientValidationFunction="ValidateVals"
                                                        EnableClientScript="true" runat="server" />
                                                </td>
                                            </tr>

                                            <tr>
                                                <td>
                                                    <asp:Label ID="lblBuyDescription" runat="server"><%= PhraseLib.Lookup("term.gc-buydesc", LanguageID)%>:</asp:Label>
                                                </td>
                                                <td>
                                                    <div id="divBuyDescriptionMLI" runat="server" style="width: 190px" />
                                                </td>
                                            </tr>


                                        </table>
                                    </div>
                                </div>
                            </ItemTemplate>
                            <SeparatorTemplate>
                                <br />
                            </SeparatorTemplate>
                        </asp:Repeater>
                    </div>
                    <div id="column5">
                        <div class="box" id="Div4">
                            <h2>
                                <span><%=PhraseLib.Lookup("term.distribution", LanguageID)%></span>
                            </h2>
                            <div style="height: 50px; overflow-y: auto;">
                                <div style="margin: 15px 0 10px 0">
                                    <asp:CheckBox ID="chkRollUp" runat="server" CssClass="tempcheck" Text="" />
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
                <a id="bottom" name="bottom"></a>
                <div id="footer">
                    <%=PhraseLib.Lookup("about.copyright", LanguageID)%>
                </div>
                <div id="custom3">
                </div>
            </div>
            <!-- End wrap -->
            <div id="custom4">
            </div>
            <script type="text/javascript">

                var symbol = "<%= CurrencySymbol %>";
                var abbr = "<%= CurrencyAbbr %>";
                var ddlValType = "<%= ddlValueType.ClientID %>";
                var NumberDecimalSeparator = "<%= CurrentUser.AdminUser.Culture.NumberFormat.NumberDecimalSeparator %>";
                function ValidateVals(sender, args) {
                    //alert("function called.");
                    //if(args.Value < 0 && args.Value > 999)
                    var selectedValType = $("#" + ddlValType).val();
                    var enteredValue = args.Value;
                    if (NumberDecimalSeparator != ".") {
                        enteredValue = enteredValue.replace(NumberDecimalSeparator, ".");
                    }
                    if (selectedValType == 1) {
                        if (enteredValue > 999999999)
                            args.IsValid = false;
                    }
                    else if (selectedValType == 3) {
                        if (enteredValue > 100)
                            args.IsValid = false;
                    }
                }
                function ManageControl(selectedValue) {
                    $("span").each(function () {
                        //Percent OFF
                        if (selectedValue == 3) {
                            if ($(this).text() == $('#ValueError').val()) {
                                {
                                    var v = document.getElementById($(this).attr("id"));
                                    v.maximumvalue = "100";
                                    v.innerHTML = $('#PercentError').val();
                                    v.style.display = "none";
                                }
                            }
                        }
                        //Dollar
                        else if (selectedValue == 1) {
                            if ($(this).text() == $('#PercentError').val()) {
                                {
                                    var v = document.getElementById($(this).attr("id"));
                                    v.maximumvalue = "999999999";
                                    v.innerHTML = $('#ValueError').val();
                                    v.style.display = "none";
                                }
                            }

                        }
                    });
                }
            </script>
    </form>
</body>
</html>
