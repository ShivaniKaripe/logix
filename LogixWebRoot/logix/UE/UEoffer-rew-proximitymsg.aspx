<%@ Page Language="C#" AutoEventWireup="true" CodeFile="UEoffer-rew-proximitymsg.aspx.cs"
    Inherits="logix_UE_UEoffer_rew_proximitymsg" %>

<%@ Register TagName="ucTemplateLockableFields" TagPrefix="uc" Src="~/logix/UserControls/TemplateFieldLockControl.ascx" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <base target="_self" />
    <title></title>
    <script type="text/javascript" src="/javascript/logix.js"></script>
    <script type="text/javascript" src="/javascript/jquery.min.js"></script>
    <script type="text/javascript" src="/javascript/jquery-ui-1.10.3/jquery-ui-min.js"></script>
    <link rel="stylesheet" href="/javascript/jquery-ui-1.10.3/style/jquery-ui.min.css" />
    <style>
        .popup {
            overflow-y: auto;
        }
    </style>
    <script type="text/javascript">

        function ChangeParentDocument() {
            opener.location = '/logix/UE/UEoffer-rew.aspx?OfferID=<%#OfferID%>';
        }

        $(document).ready(function () {
            var attrId = "";
            var position = 0;
            //var precision =  UnitPrecision %>;
            var errorMessage = "";

            $('.cpepmsg').each(function (i, obj) {
                if (attrId == "") {
                    attrId = $(this).attr('id');
                }
                SetHeightTextArea(this);
                return;
            });
            $('.cpepmsg').on('keyup keypress blur change', function (e) {
                SetHeightTextArea(this);
            });
            $('.awayitems').each(function (i, obj) {
                SetThresholdLabel(this);
            });
            $('.awayitems').change(function (e) {
                if (CheckNumeric(e)) {

                }
                else {
                    errorMessage += '<%# phraseawayfromnumeric %>' + "<br/>";
                }
                SetThresholdLabel(this);
            });

            function SetHeightTextArea(obj) {
                var length = Number($(obj).val().toString().length);
                var rowlength = Number($(obj).attr('rows'));
                var collength = 54;
                if (collength != length) {
                    if (length != 0 && length / collength > 3)
                        $(obj).attr('rows', 1 + (length / collength));
                    else
                        $(obj).attr('rows', 3);
                }
            }
            function CheckNumeric(e) {
                var NumberDecimalSeparator = '<%# NumberDecimalSeparator %>';
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
                return true;
            }
            function ValidateAwayData() {
                $('.awayitems').each(function (i, obj) {
                    var temp;
                    var precision = document.getElementById("hdnPrecision").value;
                    //                    if ($.isNumeric(parseFloat(($(this).val()).replace(',','.').replace(' ','')))) {
                    //GetPrecision();
                    if (precision == 0)
                        temp = parseFloat(($(this).val()).replace(',', '.')).toFixed(precision);
                    else
                        temp = parseFloat(parseFloat(($(this).val()).replace(',', '.'))).toFixed(precision);
                    if (isNaN(temp)) {
                        temp = ($(this).val());
                    }
                    $(this).val(temp);
                });
            }


            function SetThresholdLabel(obj) {
                var numberseperator = '<%# NumberDecimalSeparator%>';
                var awayFromValue = parseFloat(($(obj).val()).replace(',', '.'));
                var temp = awayFromValue;
                var precision = document.getElementById("hdnPrecision").value;
                if ($.isNumeric(($(obj).val()).replace(',', '.'))) {
                    //GetPrecision();
                    if (precision == 0)
                        temp = parseFloat(($(obj).val()).replace(',', '.')).toFixed(precision);
                    else
                        temp = parseFloat(parseFloat(($(obj).val()).replace(',', '.'))).toFixed(precision);
                    var requireddata = ($("#repProximityMsg_" + $(obj).attr('id').substr(16, 5) + "_requiredData").val()).replace(',', '.');
                    if (numberseperator == ",") {
                        var strng = '<%# phraseforthresholdmessage %> ' + (requireddata - temp).toFixed(precision).toString().replace('.', ',') + " ";
                   }
                   else {
                       var strng = '<%# phraseforthresholdmessage %> ' + (requireddata - temp).toFixed(precision).toString() + " ";
                    }
                    var units = $("#repProximityMsg_" + $(obj).attr('id').substr(16, 5) + "_awayAbbr").text();

                    if (awayFromValue != 0 && (requireddata != "Undefined")) {
                        $("#repProximityMsg_" + $(obj).attr('id').substr(16, 5) + "_thresholdLabel").text(" " + strng);
                        $("#repProximityMsg_" + $(obj).attr('id').substr(16, 5) + "_thresholdAbbr").text(" " + units);
                    }
                    else {
                        $("#repProximityMsg_" + $(obj).attr('id').substr(16, 5) + "_thresholdLabel").text("");
                        $("#repProximityMsg_" + $(obj).attr('id').substr(16, 5) + "_thresholdAbbr").text("");
                    }
                }
                else {
                    $("#repProximityMsg_" + $(obj).attr('id').substr(16, 5) + "_thresholdLabel").text("");
                    $("#repProximityMsg_" + $(obj).attr('id').substr(16, 5) + "_thresholdAbbr").text("");
                }
            }

            $('#btnSave').click(function () {
                var tierCount = $('.requireditems').length;
                var translationCount = $('.cpepmsg').length;
                var requiredValue = [];
                var differenceValue = [];
                var awayValue = [];
                var description = [];
                var messagetype = $('#ddlMessageType').val();
                var tierLevel = 1;
                var counterValue = translationCount / tierCount;
                var precision = document.getElementById("hdnPrecision").value;
                errorMessage = "";
                ValidateAwayData();

                $('.requireditems').each(function (i, obj) {
                    if (($(this).val()).replace(',', '.') != "undefined") {
                        requiredValue[i] = parseFloat(($(this).val()).replace(',', '.')) - 0;
                        if (i == 0) {
                            differenceValue[i] = parseFloat(($(this).val()).replace(',', '.')) - 0;
                        }
                        else {
                            differenceValue[i] = (parseFloat(($(this).val()).replace(',', '.')) - requiredValue[i - 1]).toFixed(precision);
                        }
                    }
                });

                $('.awayitems').each(function (i, obj) {
                    var difference;

                    awayValue[i] = ($(this).val());
                    if (awayValue[i] == '') {
                        errorMessage += '<%# phraseawayfromnumeric %>' + "<br/>";
                    }
                    if (!isNaN(differenceValue[i])) {
                        if (parseFloat(($(this).val()).replace(',', '.')).toFixed(precision) <= 0) {
                            errorMessage += '<%#phraseawayvaluenontiers %> ' + (i + 1) + "<br />";
                            return;
                        }
                        else if (i == 0 && parseFloat(($(this).val()).replace(',', '.')).toFixed(precision) >= requiredValue[i] && requiredValue[i] > 0) {
                            if (messagetype == 1 || messagetype == 9)
                                difference = requiredValue[i] - 1;
                            else
                                difference = requiredValue[i] - (1 / (Math.pow(10, precision)));
                            errorMessage += '<%#phraseawayvaluewithtiers %>' + difference.toFixed(precision) + ' <%# phrasefortier %> ' + (i + 1) + "<br />";
                            return;
                        }
                        else if (parseFloat(($(this).val()).replace(',', '.')).toFixed(precision) >= differenceValue[i] && differenceValue[i] > 0) {
                            if (i == 0) {
                                if (messagetype == 1 || messagetype == 9)
                                    difference = differenceValue[i] - 1;
                                else
                                    difference = differenceValue[i] - (1 / (Math.pow(10, precision)));
                            }
                            else {
                                difference = differenceValue[i];
                            }
                            errorMessage += '<%#phraseawayvaluewithtiers %>' + difference + ' <%# phrasefortier %> ' + (i + 1) + "<br />";
                            return;
                        }
                    }
                    else {

                        errorMessage += '<%#phraseconditionwarning %> ' + (i + 1) + "<br />";
                        return;
                    }
                });
                var counter = 0;
                var defaultFlag = false;
                var languages = $("#hdnLanguages").val().split(",");

                $('.cpepmsg').each(function (i, obj) {

                    defaultFlag = false;
                    if (languages[counter].split(":").length == 2)
                        defaultFlag = true;
                    var tempString = ($(this).val()).replace(',', '.');
                    //alert(languages[counter]);
                    if (tempString != "") {
                        var count = tempString.match(/"|ValueRequired|"/g || []);
                        if (count == null || count.length != 1) {
                            errorMessage += '<%#phrasetagrequired %> ' + (tierLevel) + ' <%# phraseForLanguage %> ' + languages[counter].split(":")[0] + "<br />";
                            //return;
                        }
                    }
                    else {
                        if (defaultFlag) {
                            errorMessage +='<%#phrasemessagerequiredwarning %> ' + (tierLevel) + "<br />";
                        }
                    }

                    if ((i + 1) % counterValue == 0) {
                        tierLevel++;
                    }
                    if (counter + 1 == counterValue) {
                        //defaultFlag = true;
                        counter = 0;
                    } else {
                        counter++;
                    }
                });

                if (errorMessage == "") {
                    return true;
                } else {
                    $('#infobar').html(errorMessage);
                    $('#infobar').show();
                    return false;
                }
            });


            $('.awayitems').keypress(function (e) {
                var str = String.fromCharCode(!e.charCode ? e.which : e.charCode);
                var messagetype = $('#ddlMessageType').val();
                var pattern = /[a-zA-Z]/g;
                var isMatch = pattern.test(str);
                if (messagetype != 1 && messagetype != 9 && !isMatch) {
                    return true;
                }
                else
                    if (str != "." && !isMatch)
                        return true;
                e.preventDefault();
                return false;
            });



            $("#btnPreview").click(function () {
                var popW = 700;
                var popH = 522;
                var value = $("#repProximityMsg_" + attrId.substr(16, 5) + "_awaydata").val();
                var tempMessage = $("#" + attrId).val();
                var Message = tempMessage.toString().split("|ValueRequired|").join(value);

                var myUrl = 'UEoffer-rew-proximitymsgpreview.aspx?Message=' + Message;
                window.open(myUrl, "", "width=" + popW + ", height=" + popH + ", top=" + calcTop(popH) + ", left=" + calcLeft(popW) + ", toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=yes");
            });

            $(".cpepmsg").blur(function () {
                attrId = $(this).attr("id");
            });

            $(".cpepmsg").mouseup(function () {
                position = $(this).getCursorPosition();
            }).keyup(function () {
                position = $(this).getCursorPosition();
            });

            (function ($, undefined) {
                $.fn.getCursorPosition = function () {
                    var el = $(this).get(0);
                    var pos = 0;
                    // IE Support
                    if (document.selection) {
                        el.focus();

                        var r = document.selection.createRange();
                        if (r == null) {
                            return 0;
                        }

                        var re = el.createTextRange(),
                            rc = re.duplicate();
                        re.moveToBookmark(r.getBookmark());
                        rc.setEndPoint('EndToStart', re);

                        return rc.text.length;
                    }
                    // Firefox and Chrome support
                    else if (el.selectionStart || el.selectionStart == '0')
                        pos = el.selectionStart;
                    return pos;
                }

                $.fn.extend({
                    insertAtCaret: function (myValue) {
                        if (document.selection) {
                            this.focus();
                            sel = document.selection.createRange();
                            sel.text = myValue;
                            this.focus();
                        }
                        else if (this.selectionStart || this.selectionStart == '0') {
                            var startPos = this.selectionStart;
                            var endPos = this.selectionEnd;
                            var scrollTop = this.scrollTop;
                            this.value = this.value.substring(0, startPos) + myValue + this.value.substring(endPos, this.value.length);
                            this.focus();
                            this.selectionStart = startPos + myValue.length;
                            this.selectionEnd = startPos + myValue.length;
                            this.scrollTop = scrollTop;
                        } else {
                            this.value += myValue;
                            this.focus();
                        }
                    }
                })
            })(jQuery);

            $(".insert_tag").click(function () {
                var Str = $("#" + attrId).val();
                var length = Str.length;
                var startStr = Str.substr(0, position);
                var endStr = Str.substr(position, length);
                $("#" + attrId).html(startStr + "|ValueRequired|" + endStr);
                $("#" + attrId).val(startStr + "|ValueRequired|" + endStr);
            });
        });

        function CheckValid(control) {
            if (control != null) {
                var val = control.value
                var RE = /^[0-9]{1,3}(,[0-9]{3})*(([\\.,]{1}[0-9]*)|())$/;
                if (val != '' & !RE.test(val)) {
                    var errorMessage ='<%# phraseawayfromnumeric %>' + "<br/>";
                    $('#infobar').html(errorMessage);
                    $('#infobar').show();
                    control.focus();
                    control.value = "";
                    return false;
                }
                else {
                    $('#infobar').html("");
                    $('#infobar').hide();
                }
            }
        }
    </script>
</head>
<body class="popup" onunload="ChangeParentDocument();">
    <form id="mainform" runat="server">
        <asp:HiddenField ID="hdnLanguages" runat="server" />
        <asp:HiddenField ID="hdnPrecision" runat="server" ClientIDMode="Static" Value="0" />

        <div id="wrap">
            <div id="custom2">
            </div>
            <div id="intro">
                <table>
                    <tr>
                        <td>
                            <h1 id='title' runat="server"></h1>
                        </td>
                        <td>
                            <div id='controls'>
                                <asp:Button CssClass="regular" ID="btnPreview" runat="server" />
                                <asp:Button AccessKey="s" CssClass="regular" ID="btnSave" runat="server" Text="Save"
                                    Visible="true" OnClick="Save_Click" />
                                <uc:ucTemplateLockableFields ID="ucTemplateLockableFields" runat="server" Visible="false"
                                    OnLoad="ucTemplateLockableFields_Onload" />
                            </div>
                        </td>
                    </tr>
                </table>
            </div>
            <div id="main">
                <div id="infobar" class="red-background" runat="server" style="display: none;">
                </div>
                <div id="column2x">
                    <div class="box" id="message">
                        <h2>
                            <span>
                                <%= PhraseLib.Lookup("term.data", LanguageID)%>
                            </span>
                        </h2>
                        <div style="height: 30px;">
                            <div class="table-div">
                                <table>
                                    <tr>
                                        <td style="width: 128px">
                                            <asp:Label ID="lblMessageType" runat="server"></asp:Label>
                                        </td>
                                        <td>
                                            <asp:DropDownList ID="ddlMessageType" runat="server" Width="155px" AutoPostBack="true"
                                                OnSelectedIndexChanged="Message_Changed" />
                                        </td>
                                    </tr>
                                </table>
                            </div>
                        </div>
                        <hr style="margin-top: 10px;" />
                        <asp:Repeater ID="repProximityMsg" runat="server" OnItemDataBound="BindProximityMsgDesc">
                            <ItemTemplate>
                                <br />
                                <label style="font-weight: bold;">
                                    <%#(objOffer.NumbersOfTier>1?PhraseLib.Lookup("term.tier", LanguageID) + " " + DataBinder.Eval(Container.DataItem, "TierLevel") + ":":"" ) %>
                                </label>
                                <div class="table-div">
                                    <table>
                                        <tr>
                                            <td style="width: 80%">
                                                <asp:Label ID="requiredLabel" runat="server" Text="" />
                                            </td>
                                            <td style="width: 2%">
                                                <asp:Label ID="requiredSymbol" runat="server" Text=""></asp:Label>
                                            </td>
                                            <td style="width: 10%">
                                                <asp:TextBox ID="requiredData" CssClass="requireditems" runat="server" class="short"
                                                    type="text" Text="" ReadOnly="true" disabled="true" />
                                            </td>
                                            <td style="width: 5%">
                                                <asp:Label ID="requiredAbbr" runat="server" Text=""></asp:Label>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td style="width: 80%">
                                                <asp:Label ID="awayLabel" runat="server" Text="" />
                                            </td>
                                            <td style="width: 2%">
                                                <asp:Label ID="awaySymbol" runat="server" Text=""></asp:Label>
                                            </td>
                                            <td style="width: 10%">
                                                <asp:TextBox ID="awaydata" CssClass="awayitems" runat="server" class="short" type="text"
                                                    value='<%#((Decimal)DataBinder.Eval(Container.DataItem, "TriggerValue")).ToString(UnitPrecisionFormat, System.Globalization.CultureInfo.InvariantCulture) %>' onblur="CheckValid(this)" />
                                            </td>
                                            <td>
                                                <asp:Label ID="awayAbbr" runat="server" Text=""></asp:Label>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td colspan="4">
                                                <asp:Label runat="server" ID="thresholdLabel" />
                                                <asp:Label runat="server" ID="thresholdAbbr" />
                                            </td>
                                        </tr>
                                    </table>
                                    <div runat="server">
                                    </div>
                                    <asp:Repeater ID="repProximityMsgDesc" runat="server">
                                        <ItemTemplate>
                                            <br />
                                            <label>
                                                <%# (SystemSettings.IsMultiLanguageEnabled() ? 
                                                PhraseLib.Lookup(Languages.Where(l => l.LanguageID == (int)DataBinder.Eval(Container.DataItem, "LanguageId")).FirstOrDefault().PhraseTerm, LanguageID) + " " +
                                                                                                                                              (CustomerFacingLangID == (int)DataBinder.Eval(Container.DataItem, "LanguageId") ? "(" + PhraseLib.Lookup("term.default", LanguageID) + ")" : "") : "")
                                                %>
                                            </label>
                                            <div class="pmsgwrap">
                                                <textarea id="prmessage" runat="server" name="pmdesc" class="cpepmsg" rows="3" cols="38"
                                                    style="width: 98%; overflow-y: scroll;"><%#DataBinder.Eval(Container.DataItem, "Message") %></textarea>
                                            </div>
                                        </ItemTemplate>
                                    </asp:Repeater>
                                </div>
                            </ItemTemplate>
                        </asp:Repeater>
                    </div>
                </div>
                <div id="gutter">
                </div>
                <div id="column1x">
                    <div class="box" id="tags" style="position: fixed">
                        <h2>
                            <span>
                                <%= PhraseLib.Lookup("term.tags", LanguageID)%>
                            </span>
                        </h2>
                        <br class="half" />
                        <div id="ed_toolbar" style="background-color: #d0d0d0; text-align: center;">
                            <div id="tools">
                                <asp:Button ID="ed_normal" CssClass="ed_button insert_tag" runat="server" />
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </form>
</body>
</html>
