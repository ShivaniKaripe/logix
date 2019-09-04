<%@ Page Language="C#" Debug="true" AutoEventWireup="true" CodeFile="UEoffer-con-brokerinstantwin.aspx.cs"
  Inherits="logix_UE_UEOffer_con_BrokerInstantWin" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
  <title></title>
  <script type="text/javascript" src="/javascript/logix.js"></script>
  <script type="text/javascript" src="/javascript/jquery.min.js"></script>
  <script type="text/javascript" src="/javascript/jquery-ui-1.10.3/jquery-ui-min.js"></script>
  <link rel="stylesheet" href="/javascript/jquery-ui-1.10.3/style/jquery-ui.min.css" />
  <style type="text/css">
    #Image1
    {
      opacity: 0.4;
      filter: alpha(opacity=40); /* For IE8 and earlier */
    }
    
    #Image1:hover
    {
      opacity: 1.0;
      filter: alpha(opacity=100); /* For IE8 and earlier */
    }
    
    .ui-dialog .ui-dialog-content
    {
      background-color: #c0c0c0;
      border: solid 1px #dddddd;
    }
    input[type="radio"]
    {
      margin: 7px;
    }
    .ui-dialog-titlebar
    {
      background: #0066ff;
      color: #ffffff;
      font-weight: bold;
      line-height: 20px;
    }
    #vsError
    {
      display: none;
      visibility: hidden;
      width: 655px;
    }
    .table-div
    {
      margin-left: 60px;
      margin-top: 30px;
      border-style: solid;
      border-color: Green;
    }
  </style>
  <script type="text/javascript">
    function isNumber(evt) {
      debugger;
      var e = evt || window.event;
      var key = e.keyCode || e.which;
      if (!e.shiftKey && !e.altKey && !e.ctrlKey &&
      // numbers   
    key >= 48 && key <= 57 ||
      // Numeric keypad
    key >= 96 && key <= 105 ||
      // Backspace and Tab and Enter
    key == 8 || key == 9 || key == 13 ||
      // Home and End
    key == 35 || key == 36 ||
      // left and right arrows
    key == 37 || key == 39 ||
      // Del and Ins
    key == 46 || key == 45) {
        // input is VALID
      }
      else {
        // input is INVALID
        e.returnValue = false;
        if (e.preventDefault) e.preventDefault();
      }
    }
    function CheckZero(control) {
      if (control.value == "0")
        control.value = "";
    }
    function ParseBool(val) {
      if (val.toUpperCase() == "TRUE")
        return true;
      else
        return false;
    }
    $(document).ready(function () {
      //ready will be loaded once all the page methods are executed
      var isTemplate = $("#hdnIsTemplate").val();
      var fromTemplate = $("#hdnFromTemplate").val();
      var DisallowEdit = $("#hdnDisallowEdit").val();
      $("#infobar")[0].style.display = "none";
      if ($("#hdnInstantWinID").val() == "" || $("#hdnInstantWinID").val() == "0") {
        loadDefaultUI();
      }
      else {
        loadUIforInstantWinCondition(fromTemplate);
      }
      if ($("#hdnOfferID").val() == "0" || $("#hdnOfferID").val() == "")
        $("#btnSave")[0].disabled = true;

      if (ParseBool(isTemplate)) {
        if (ParseBool(DisallowEdit)) {
          $("#chkDisallow_Edit").prop("checked", true);
          $("#iwBody")[0].className = "popup template";
        }
      }
      else {
        $("#iwBody")[0].style.display = "popup";
      }
      if (ParseBool(fromTemplate) && ParseBool(DisallowEdit)) {
        $("#content").children().prop('disabled', true);
        var nodes = document.getElementById("content").getElementsByTagName('*');
        for (var i = 0; i < nodes.length; i++) {
          nodes[i].disabled = true;
        }
      }
      updateTotalRewards();
    });

    function loadUIforInstantWinCondition() {
      var IWType = $("#hdnProgramType").val().toString().toLowerCase();
      var AwardLimitEnterprise = $("#hdnAwardLimitEnterprise").val().toString().toLowerCase();
      var ChanceOfWinning = $("#hdnChanceOfWinning").val().toString().toLowerCase();
      var ChanceOfWinningEnterprise = $("#hdnChanceOfWinningEnterprise").val().toString().toLowerCase();
      var NumPrizesAllowed = $("#hdnNumPrizesAllowed").val().toString().toLowerCase();
      var Unlimited = $("#hdnUnlimited").val().toString().toLowerCase();
      if (IWType == "random") {
        enableRandomDiv(true);
        $("#rbtnRandom").prop("checked", true);
        $("#tboxRandomUsr").val(ChanceOfWinning);
      }
      else {
        enableRandomDiv(false);
        $("#rbtnSequence").prop("checked", true);
        $("#tboxSequenceUsr").val(ChanceOfWinning);
      }
      if (ParseBool(ChanceOfWinningEnterprise)) {
        $("#rbtnListChanceOfWin_1").prop("checked", true);
        $("#rbtnListChanceOfWin_0").prop("checked", false);
      }
      else {
        $("#rbtnListChanceOfWin_1").prop("checked", false);
        $("#rbtnListChanceOfWin_0").prop("checked", true);
      }

      if (ParseBool(Unlimited)) {
        $("#rbtnUnlimited").prop("checked", true);
        $("#rbtnLimited").prop("checked", false);
        DisableAwardDivUnlimited(true);
      }
      else {
        $("#rbtnLimited").prop("checked", true);
        $("#rbtnUnlimited").prop("checked", false);
        if (ParseBool(AwardLimitEnterprise)) {
          $("#rbtnListAwardLimit_1").prop("checked", true);
          $("#rbtnListAwardLimit_0").prop("checked", false);
        }
        else {
          $("#rbtnListAwardLimit_1").prop("checked", false);
          $("#rbtnListAwardLimit_0").prop("checked", true);
        }
        DisableAwardDivUnlimited(false);
        $("#tboxAwardLimitNumber").val(NumPrizesAllowed);
      }
      if ($("#hdnNoOfStores").val() == "-1")
        $("#lblTotalAwards")[0].innerText = "-";
      else
        $("#lblTotalAwards")[0].innerText = "0";
    }
    function loadDefaultUI() {


      $("#rbtnRandom").prop("checked", true);
      $("#divChanceOfWinRandomUsr")[0].style.display = "block";
      $("#divChanceOfWinSequenceUsr")[0].style.display = "none";

      $("#rbtnUnlimited").prop("checked", true);
      $("#rbtnListAwardLimit_1").prop("checked", true);
      $("#rbtnListAwardLimit_1")[0].disabled = true;
      $("#rbtnListAwardLimit_0")[0].disabled = true;
      $("#tboxAwardLimitNumber")[0].disabled = true;
      $("#tboxAwardLimitNumber").val("");

      if ($("#hdnNoOfStores").val() == "-1")
        $("#lblTotalAwards")[0].innerText = "-";
      else
        $("#lblTotalAwards")[0].innerText = "0";
    }
    function updateTotalRewards() {
      var noOfStores = parseInt($("#hdnNoOfStores").val())
      if ($("#rbtnUnlimited")[0].checked) {
        $("#lblTotalAwards")[0].innerText = $("#hdnIWUnlimited").val();
        $("#imghelp")[0].title = $("#hdnIWUnlimited").val();
        updateWinners();
        return;
      }
      if (noOfStores == "-1") {
        $("#lblTotalAwards")[0].innerText = "-";
        $("#imghelp")[0].title = $("#hdnIWStoregroupError").val();
        updateWinners();
        return
      }
      if ($("#rbtnLimited")[0].checked) {
        var val = $("#tboxAwardLimitNumber").val()
        if (val == "") {
          $("#lblTotalAwards")[0].innerText = "-";
          $("#imghelp")[0].title = $("#hdnIWAwardLimitError").val();
          updateWinners();
          return;
        }
        if ($("#rbtnListAwardLimit_0")[0].checked) {
          $("#lblTotalAwards")[0].innerText = val * noOfStores;
          $("#imghelp")[0].title = val + " " + $("#hdnAwardAtStore").val() + $("#hdnTotal").val() + " " + noOfStores + " " + $("#hdnStrStore").val() + ".";
        }
        if ($("#rbtnListAwardLimit_1")[0].checked) {
          $("#lblTotalAwards")[0].innerText = val;          
          $("#imghelp")[0].title = val + " " + $("#hdnAwardAcross").val() + " " + noOfStores + " " + $("#hdnStrStore").val() + ".";
        }
        updateWinners();
        return;
      }
    }
    function enableRandomDiv(flag) {
      if (flag) {
        $("#divChanceOfWinRandomUsr")[0].style.display = "block";
        $("#divChanceOfWinSequenceUsr")[0].style.display = "none";

      }
      else {
        $("#divChanceOfWinSequenceUsr")[0].style.display = "block";
        $("#divChanceOfWinRandomUsr")[0].style.display = "none";
      }
      return true;
    }

    function DisableAwardDivUnlimited(flag) {
      $("#rbtnListAwardLimit_1")[0].disabled = flag;
      $("#rbtnListAwardLimit_0")[0].disabled = flag;
      $("#tboxAwardLimitNumber")[0].disabled = flag;
      if (flag) {
        $("#tboxAwardLimitNumber").val("");
        $("#rbtnListAwardLimit_1").prop("checked", flag);
      }
      return true;
    }


    function ChangeParentDocument() {
      var newlocation = "/logix/UE/UEoffer-con.aspx?OfferID=" + $("#hdnOfferID").val();
      if (opener != null) {
        if (opener.location.href.indexOf(newlocation) > -1) {
          opener.location = "/logix/UE/UEoffer-con.aspx?OfferID=" + $("#hdnOfferID").val();
        }
      }
    }

    function validateSave() {
      var flag = true;
      var ErrorMsg;
      var InvalidRandom = $("#rbtnRandom")[0].checked && ($("#tboxRandomUsr").val() == "" || parseInt($("#tboxRandomUsr").val()) < 0)
      var InvalidSequence = $("#rbtnSequence")[0].checked && ($("#tboxSequenceUsr").val() == "" || parseInt($("#tboxSequenceUsr").val()) < 0)

      if (InvalidRandom || InvalidSequence) {
        flag = false;
        ErrorMsg = "" + $("#hdnIWIntegerError").val() + " " + $("#hdnIWchanceofWin").val();
      }

      if (flag && !($("#rbtnListChanceOfWin_1")[0].checked) && !($("#rbtnListChanceOfWin_0")[0].checked)) {
        flag = false;
        ErrorMsg = $("#hdnIWChanceOfWinError").val();
      }

      if (flag && $("#rbtnLimited")[0].checked) {
        if ($("#tboxAwardLimitNumber").val() == "" || parseInt($("#tboxAwardLimitNumber").val()) < 0) {
          flag = false;
          ErrorMsg = $("#hdnIWIntegerError").val() + " " + $("#hdnIWawardLimit").val();
        }
        if (flag && !($("#rbtnListAwardLimit_1")[0].checked) && !($("#rbtnListAwardLimit_0")[0].checked)) {
          flag = false;
          ErrorMsg = $("#hdnIWAWLimitAppliedError").val();
        }
      }

      if (!flag) {
        if ($("#infobar") != null) {
          $("#infobar")[0].innerHTML = ErrorMsg;
          $("#infobar")[0].style.display = "block";
        }
      }
      return flag;
    }
    function RecalculateCount() {
      Winners = -1;
      abort = true;
      if (Progress)
        Progress.abort();
      if (xdr)
        xdr.abort();
      updateWinners();
    }
    var Winners = -1;
    var Progress;
    var xdr;
    var abort = false;
    function updateWinners() {
        debugger
        abort = false;
        var totalrwrd = parseInt($("#lblTotalAwards")[0].innerText);
        if (isNaN(totalrwrd)) {
            $('#lblAwardsRemaining').html($("#lblTotalAwards")[0].innerText);
        }
        else {
            if (!isNaN(Winners) && Winners >= 0 && (totalrwrd - Winners) > 0) {
                $('#lblAwardsRemaining').html(totalrwrd - Winners);
            }
            else {
                $('#lblAwardsRemaining').html(0);
            }
        }

        var strurl = '/UE/UEoffer-con-brokerinstantwin.aspx/LoadBrokerData';
        var storeArray = $("#hdnStores").val();
        if (Winners < 0) {
            $('#lblWinnersCount').hide();
            $('#ImgWinnersCount').show();
            Progress = $.ajax({
                type: 'POST',
                url: "UEoffer-con-brokerinstantwin.aspx/FetchWinners",
                data: '{offerId:"' + $("#hdnOfferID").val() + '",storeNames:"' + storeArray + '"}',
                contentType: "application/json; charset=utf-8",
                dataType: 'json'
            })
            .done(function (data) {
                onSuccess(data.d, totalrwrd);
            })
            .fail(function (data) {
                onErrror(data.d);
            });
        }
    }
    function onSuccess(val, totalrwrd) {
      debugger
      $('#lblWinnersCount').show();
      $('#ImgWinnersCount').hide();
      Winners = val;
      if (!isNaN(Winners)) {
        $('#lblWinnersCount').html(Winners);
        if (!isNaN(totalrwrd) && !isNaN(Winners)) {
          if (Winners >= 0 && (totalrwrd - Winners) > 0) {
            $('#lblAwardsRemaining').html(totalrwrd - Winners);
          }
          else {
            $('#lblAwardsRemaining').html(0);
          }
        }
      }
      else
        onErrror(Winners)
    }
    function onErrror(data) {
      debugger
      if (!abort) {
        Winners = 0;
        $('#ImgWinnersCount').hide();
        $('#lblWinnersCount').show();
        var temperrText = $("#hdnIWBrokerError").val().toString();
        $('#lblWinnersCount').html(temperrText);
        $('#lblAwardsRemaining').html(temperrText);
      }
    }
  </script>
</head>
<body id="iwBody" class="popup" onunload="ChangeParentDocument();">
  <form id="frmIWCondition" runat="server" clientidmode="Static">
  <input id="hdnOfferID" type="hidden" runat="server" clientidmode="Static" />
  <input id="hdnDissalowEdit" type="hidden" runat="server" clientidmode="Static" />
  <input id="hdnIsTemplate" type="hidden" runat="server" clientidmode="Static" />
  <input id="hdnFromTemplate" type="hidden" runat="server" clientidmode="Static" />
  <input id="hdnInstantWinID" type="hidden" runat="server" clientidmode="Static" />
  <input id="hdnAwardLimitEnterprise" type="hidden" runat="server" clientidmode="Static" />
  <input id="hdnChanceOfWinning" type="hidden" runat="server" clientidmode="Static" />
  <input id="hdnChanceOfWinningEnterprise" type="hidden" runat="server" clientidmode="Static" />
  <input id="hdnDisallowEdit" type="hidden" runat="server" clientidmode="Static" />
  <input id="hdnNumPrizesAllowed" type="hidden" runat="server" clientidmode="Static" />
  <input id="hdnUnlimited" type="hidden" runat="server" clientidmode="Static" />
  <input id="hdnProgramType" type="hidden" runat="server" clientidmode="Static" />
  <input id="hdnDeleted" type="hidden" runat="server" clientidmode="Static" />
  <input id="hdnGetWinnersURL" type="hidden" runat="server" clientidmode="Static" />
  <input id="hdnNoOfStores" type="hidden" runat="server" clientidmode="Static" />
  <input id="hdnIWBrokerError" type="hidden" value='<%=PhraseLib.Lookup("term.IWBrokerError", LanguageID)  %>' />
  <input id="hdnIWAwardLimitError" type="hidden" value='<%=PhraseLib.Lookup("term.IWAwardLimitError", LanguageID)  %>' />
  <input id="hdnIWStoregroupError" type="hidden" value='<%=PhraseLib.Lookup("term.IWStoregroupError", LanguageID)  %>' />
  <input id="hdnIWIntegerError" type="hidden" value='<%=PhraseLib.Lookup("term.IWIntegerError", LanguageID)  %>' />
  <input id="hdnIWChanceOfWinError" type="hidden" value='<%=PhraseLib.Lookup("term.IWChanceOfWinError", LanguageID)  %>' />
  <input id="hdnIWAWLimitAppliedError" type="hidden" value='<%=PhraseLib.Lookup("term.IWAWLimitAppliedError", LanguageID)  %>' />
  <input id="hdnIWoddsofWinning" type="hidden" value='<% =PhraseLib.Lookup("offer-gen.oddsofwinning", LanguageID) %>' />
  <input id="hdnIWawardLimit" type="hidden" value='<% =PhraseLib.Lookup("term.IWawardLimit", LanguageID) %>' />
  <input id="hdnIWchanceofWin" type="hidden" value='<% =PhraseLib.Lookup("term.IWchanceofWin", LanguageID) %>' />
  <input id="hdnIWUnlimited" type="hidden" value='<% =PhraseLib.Lookup("term.unlimited", LanguageID) %>' />
  <input id="hdnValidationError" type="hidden" value='<% =PhraseLib.Lookup("term.IWValidationError", LanguageID) %>' />
  <input id="hdnStrStore" type="hidden" value='<% =PhraseLib.Lookup("term.stores", LanguageID) %>' />
  <input id="hdnAwardAtStore" type="hidden" value='<% =PhraseLib.Lookup("term.awardatstore", LanguageID) %>' />
  <input id="hdnTotal" type="hidden" value='<% =PhraseLib.Lookup("term.total", LanguageID) %>' />
  <input id="hdnAwardAcross" type="hidden" value='<% =PhraseLib.Lookup("term.awardacross", LanguageID) %>' />
  <input id="hdnStores" type="hidden" runat="server" clientidmode="Static" />
  <div id="content">
    <div id="intro">
      <h1 id='title' runat="server">
        Title</h1>
      <div id='controls'>
        <span class="temp" id="TempDisallow" runat="server">
          <asp:CheckBox ID="chkDisallow_Edit" runat="server" CssClass="tempcheck" />
          <label id="lblLocked" for="chkDisallow_Edit">
            <%=PhraseLib.Lookup("term.locked", LanguageID)%>
          </label>
        </span>
        <asp:Button AccessKey="s" CssClass="regular" ID="btnSave" runat="server" Text=""
          OnClientClick="return validateSave()" OnClick="btnSave_Click" />
      </div>
    </div>
    <div id="main">
      <br />
      <div id="infobar" class="red-background" runat="server" style="width: 96%; height: auto;"
        clientidmode="Static">
      </div>
      <table style="width: 96%; height: auto;" align="center">
        <tr>
          <td valign="top" style="width: 50%">
            <asp:Panel runat="server" ID="divIWPrgType" ClientIDMode="Static" class="box">
              <h2 style="width: inherit;">
                <% =PhraseLib.Lookup("term.IWprgType", LanguageID) %>
              </h2>
              <asp:RadioButton ID="rbtnRandom" runat="server" ClientIDMode="Static" GroupName="IWprgType"
                onClick="enableRandomDiv(true)" />
              <br />
              <asp:RadioButton ID="rbtnSequence" runat="server" ClientIDMode="Static" GroupName="IWprgType"
                onClick="enableRandomDiv(false)" Text="" />
              <br />
              <br />
            </asp:Panel>
            <div id="divChanceOfWin" runat="server" clientidmode="Static" class="box">
              <h2>
                <% =PhraseLib.Lookup("term.IWchanceofWin", LanguageID) %>
              </h2>
              <br />
              <div id="divChanceOfWinRandomUsr" runat="server" clientidmode="Static">
                <% =PhraseLib.Lookup("term.Random1", LanguageID)%>
                <asp:TextBox ID="tboxRandomUsr" runat="server" MaxLength="8" Width="75px" onpaste="return isNumber(event)" onkeydown="return isNumber(event)" 
                  oninput="CheckZero(this)" ClientIDMode="Static"></asp:TextBox>
                <% =PhraseLib.Lookup("term.Random2", LanguageID)%>
              </div>
              <div id="divChanceOfWinSequenceUsr" runat="server" clientidmode="Static">
                <% =PhraseLib.Lookup("term.every", LanguageID) %>
                <asp:TextBox ID="tboxSequenceUsr" runat="server" MaxLength="8" Width="75px" onpaste="return isNumber(event)" onkeydown="return isNumber(event)"
                  oninput="CheckZero(this)" ClientIDMode="Static"></asp:TextBox>
                <% =PhraseLib.Lookup("term.custWinner", LanguageID) %>
              </div>
              <asp:Panel id="divchanceofWinList" runat="server" clientidmode="Static">
                <br />
                <h3>
                  <% =PhraseLib.Lookup("term.IWchanceofWin", LanguageID)%><% =PhraseLib.Lookup("term.areApplied", LanguageID)%></h3>
                <asp:RadioButtonList ID="rbtnListChanceOfWin" runat="server" ClientIDMode="Static"
                  RepeatLayout="Flow">
                  <asp:ListItem Text="" Value="0"></asp:ListItem>
                  <asp:ListItem Text="" Value="1"></asp:ListItem>
                </asp:RadioButtonList>
              </asp:Panel>
              <br />
            </div>
            <div id="divAwardLimit" runat="server" clientidmode="Static" class="box">
              <h2>
                <% =PhraseLib.Lookup("term.IWawardLimit", LanguageID) %>
              </h2>
              <!--  -->
              <asp:RadioButton ID="rbtnUnlimited" runat="server" ClientIDMode="Static" GroupName="AwardLimit"
                Text="" onClick="DisableAwardDivUnlimited(true);" onchange="updateTotalRewards()" />
              <br />
              <asp:RadioButton ID="rbtnLimited" runat="server" ClientIDMode="Static" GroupName="AwardLimit"
                Text="" onClick="DisableAwardDivUnlimited(false);" onchange="updateTotalRewards()" />
              <asp:TextBox ID="tboxAwardLimitNumber" runat="server" MaxLength="8" ClientIDMode="Static"
                Width="75px" onpaste="return isNumber(event)" onkeydown="return isNumber(event)" oninput="CheckZero(this)" onblur="updateTotalRewards()"></asp:TextBox>
              <label for="rbtnLimited">
                <% =PhraseLib.Lookup("term.winsduring", LanguageID) %><!--winners during the --></label>
              <br />
              <br />
              <h3>
                <% =PhraseLib.Lookup("term.appliedawardapplied", LanguageID) %></h3>
              <asp:Panel id="divAwardLimitrbtnList" runat="server" clientidmode="Static">
                <asp:RadioButtonList ID="rbtnListAwardLimit" runat="server" ClientIDMode="Static"
                  onchange="updateTotalRewards()" RepeatLayout="Flow">
                  <asp:ListItem Text="" Value="0"></asp:ListItem>
                  <asp:ListItem Text="" Value="1"></asp:ListItem>
                </asp:RadioButtonList>
              </asp:Panel>
              <br />
            </div>
          </td>
          <td valign="top" style="width: 50%">
            <div id="divwardStatus" runat="server" clientidmode="Static" class="box" style="min-height: 180px;
              height: auto">
              <h2>
                <% =PhraseLib.Lookup("term.IWawardStatus", LanguageID) %>
              </h2>
              <table id="tblawardStatus" align="left" style="border-collapse: separate; border-spacing: 0 1.3em;">
                <tr>
                  <td>
                    <% =PhraseLib.Lookup("term.IWmaxAwards", LanguageID) %>
                  </td>
                  <td>
                    <asp:Label ID="lblTotalAwards" runat="server" Text="" />
                  </td>
                  <td>
                    <asp:Image ID="imghelp" runat="server" Height="20" Width="20" ToolTip="" AlternateText="(i)"
                      ImageUrl="~/images/information.png" BorderWidth="0" />
                  </td>
                </tr>
                <tr>
                  <td>
                    <% =PhraseLib.Lookup("term.IWtotalWinners", LanguageID) %>
                  </td>
                  <td>
                    <asp:Label ID="lblWinnersCount" runat="server" Text="" />
                    <asp:Image ID="ImgWinnersCount" src="../../images/loader.gif" Style="display: none"
                      alt="Loading" runat="server" Height="10px" Width="40px" />
                  </td>
                  <td>
                    <asp:Image ID="Image1" runat="server" Height="20" Width="20" 
                      ImageUrl="~/images/view_refresh.png" BorderWidth="0" onclick="RecalculateCount();" />
                  </td>
                </tr>
                <tr>
                  <td>
                    <% =PhraseLib.Lookup("term.IWawardsRem", LanguageID) %>
                  </td>
                  <td>
                    <asp:Label ID="lblAwardsRemaining" runat="server" Text="" />
                  </td>
                </tr>
              </table>
              <br />
            </div>
          </td>
        </tr>
      </table>
    </div>
  </div>
  </form>
</body>
</html>
