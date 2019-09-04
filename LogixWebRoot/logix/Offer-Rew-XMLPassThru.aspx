<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Offer-Rew-XMLPassThru.aspx.cs"
  ValidateRequest="false" Inherits="logix_Offer_Rew_XMLPassThru" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
  <base target="_self" />
  <title></title>
  <script type="text/javascript" src="/javascript/logix.js"></script>
  <script type="text/javascript" src="/javascript/jquery.min.js"></script>
  <script type="text/javascript" src="/javascript/jquery-ui-1.10.3/jquery-ui-min.js"></script>
  <link rel="stylesheet" href="/javascript/jquery-ui-1.10.3/style/jquery-ui.min.css" />
  <style type="text/css">
    .ui-dialog .ui-dialog-content
    {
      background-color: #c0c0c0;
      border: solid 1px #dddddd;
    }
    
    .ui-dialog-titlebar
    {
      background: #0066ff;
      color: #ffffff;
      font-weight: bold;
      line-height: 20px;
    }
  </style>
  <script language="javascript" type="text/javascript">

    var caretPos = 0;
    var caretTID = 0;
    var bCapture = true;
    var srcElement = document.getElementById('txtData');
    caretID = setInterval("captureCursorPosition();", 200);

    function radioclicksv() {
      var objtxt = document.getElementById("functioninputsv");
      handleKeyUpsv(200, objtxt)
    }
    String.prototype.PadLeft = function (pad_length, pad_string) {
      var output = this
      while (output.length < pad_length) {
        output = pad_string + output;
      }
      return output;
    }
    String.prototype.PadRight = function (pad_length, pad_string) {
      var output = this;
      while (output.length < pad_length) {
        output = output + pad_string;
      }
      return output;
    }
    function radioclickpoint() {
      var objtxt = document.getElementById("functioninputpoint");
      handleKeyUppoint(200, objtxt)
    }

    $(document).ready(function () {
      var DisableSV = ($("#hdDisableSV").val() == 1);
      if (DisableSV) {
        $("#ed_SVID").attr('disabled', true);
      }

      var DisablePoint = ($("#hdDisablePoint").val() == 1);
      if (DisablePoint) {
        $("#ed_PointsID").attr('disabled', true);
      }

    });
    
    function captureCursorPosition() {

      if (bCapture && srcElement != null) {
        caretPos = getCaret(srcElement);
      }
    }

    function setCursorPosition(elem) {

      if (elem != null) {
        if (elem.createTextRange) {
          var range = elem.createTextRange();
          range.move('character', caretPos);
          range.select();
        } else {
          if (elem.selectionStart) {
            elem.focus();
            elem.setSelectionRange(caretPos, caretPos);
          }
          else
            elem.focus();
        }
      }
      bCapture = true;

      bCapture = true;

    }
    function getCaret(el) {
      var CaretPos = 0;
      // IE Support
      if (document.selection) {
        el.focus();
        var Sel = document.selection.createRange();
        var Sel2 = Sel.duplicate();
        Sel2.moveToElementText(el);
        var CaretPos = 0;
        var CharactersAdded = 1;
        while (Sel2.inRange(Sel)) {
          //old GetCaretPosition always counts 1 for linetermination
          if (Sel2.htmlText.substr(0, 2) == "\r\n") {
            CaretPos += 2;
            CharactersAdded = 2;
          } else {
            CaretPos++;
            CharactersAdded = 1;
          }
          Sel2.moveStart('character');
        }
        CaretPos -= CharactersAdded;
      }
      // Firefox support
      else if (el.selectionStart || el.selectionStart == '0')
        CaretPos = el.selectionStart;
      return (CaretPos);
    }
    function edInsertContent(myValue) {


      if (srcElement == null) {

        return;
      }
      myField = srcElement;
      setCursorPosition(srcElement);
      caretPos = caretPos + myValue.length;

      //IE support
      if (document.selection) {
        myField.focus();
        sel = document.selection.createRange();
        sel.text = myValue;
        myField.focus();
      }
      //MOZILLA/NETSCAPE support
      else if (myField.selectionStart || myField.selectionStart == '0') {
        var startPos = myField.selectionStart;
        var endPos = myField.selectionEnd;
        var scrollTop = myField.scrollTop;
        myField.value = myField.value.substring(0, startPos)
		              + myValue
                      + myField.value.substring(endPos, myField.value.length);
        myField.focus();
        myField.selectionStart = startPos + myValue.length;
        myField.selectionEnd = startPos + myValue.length;
        myField.scrollTop = scrollTop;
      } else {
        myField.value += myValue;
        myField.focus();
      }


    }


    function handlesvclick(obj) {
        if(obj.value != ""){
            $("#svselector").dialog("close");
            edInsertContent("<!-- StoredValueID -->" + obj.value);
        }

    }
    function handlepointclick(obj) {
        if(obj.value != ""){
            $("#pointselector").dialog("close");
            edInsertContent("<!-- PointsID -->" + obj.value);
        }
      
    
    }

    function copytext(obj) {
        var value1 = obj.value + "";
        if(value1 != ""){
            var obj = document.getElementById("txtCode");
            obj.value = value1;
            validateCode('txtCode');
        }
    }

    function handlecodeclick(obj) {
      $("#tcselector").dialog("close");
      var value = obj.value + "";
      if(IsNumeric(value, false, false)){
      value ="<!-- TriggerCode -->" + value.PadLeft(codesettingsJSON.PadLength, codesettingsJSON.PadLetter);
      }
      else{
          value = "<!-- TriggerCode -->" + value;
      }
      edInsertContent(value);

    }


    function handleKeyUpsv(maxNumToShow, textObj) {

      var selectObj;
      var i, numShown;
      var searchPattern;

      //document.getElementById("functionselectsv").size = "10";

      // Set references to the form elements
      selectObj = document.getElementById("lbSV");

      // Set the search pattern depending
      if (document.getElementById('functionradiosva').checked == true) {
        searchPattern = "^" + textObj.value;
      } else {
        searchPattern = textObj.value;
      }
      searchPattern = cleanRegExpString(searchPattern);

      // Create a regulare expression
      re = new RegExp(searchPattern, "gi");

      // Clear the options list
      selectObj.length = 0;

      // Loop through the array and re-add matching options
      numShown = 0;
      $.each(svJSON, function (i, el) {
        if (el.ProgramName.search(re) != -1) {
          selectObj[numShown] = new Option(el.ProgramName, el.SVProgramID);
          numShown++;
          if (numShown >= maxNumToShow)
            return false;
        }
      });

      // When options list whittled to one, select that entry
      if (selectObj.length == 1) {
        selectObj.options[0].selected = true;
      }
    }

    function handleKeyUppoint(maxNumToShow, textObj) {

      var selectObj;
      var i, numShown;
      var searchPattern;

      //document.getElementById("functionselectsv").size = "10";

      // Set references to the form elements
      selectObj = document.getElementById("lbPoints");


      // Set the search pattern depending
      if (document.getElementById('functionradiopointa').checked == true) {
        searchPattern = "^" + textObj.value;
      } else {
        searchPattern = textObj.value;
      }

      searchPattern = cleanRegExpString(searchPattern);

      // Create a regulare expression
      re = new RegExp(searchPattern, "gi");

      // Clear the options list
      selectObj.length = 0;

      // Loop through the array and re-add matching options
      numShown = 0;
      $.each(pointJSON, function (i, el) {
        if (el.ProgramName.search(re) != -1) {
          selectObj[numShown] = new Option(el.ProgramName, el.ProgramID);
          numShown++;
          if (numShown >= maxNumToShow)
            return false;
        }
      });

      // When options list whittled to one, select that entry
      if (selectObj.length == 1) {
        selectObj.options[0].selected = true;
      }
    }
  </script>
</head>
<body class="<%=objOffer.IsTemplate? "popup template" : "popup"%>" onunload="ChangeParentDocument();">
  <script language="javascript" type="text/javascript">
var svJSON=jQuery.parseJSON('<%=StoredValueJSON%>');
var pointJSON=jQuery.parseJSON('<%=PointsJSON%>');
var codesettingsJSON=jQuery.parseJSON('<%=CodeSettingsJSON%>');

  function ChangeParentDocument() {
    <%= GetRefreshScript() %>

 }
 
    function pageLoad() {
      var item=  $('textarea[id^=txtData]')[0];
    $('#hdnPointsHtml').val($("#pointselector")[0].innerHTML);
     $('#hdnSVHtml').val($("#svselector")[0].innerHTML);
       $('#hdnCodeHtml').val($("#tcselector")[0].innerHTML);
    item.focus();
    var dlgsv = $("#svselector").dialog({

        modal: true,
        draggable: true,
        resizable: false,
        show: 'Transfer',
        hide: 'Transfer',
        width: 300,
        title: '<%=PhraseLib.Lookup("offer-rew.selectsv", LanguageID)%>',
        autoOpen: false,
        minHeight: 10,
        minwidth: 10,
        closeText: 'X',
        closeOnEscape: true,
        overlay: {
        opacity: 0.65 },
         open:function(){
        $("#svselector")[0].innerHTML=$('#hdnSVHtml').val();
        }

       
       

    });
    $("#ed_SVID").click(function () {
        $("#svselector").dialog("open");
    });

    var dlgpoint = $("#pointselector").dialog({

        modal: true,
        draggable: true,
        resizable: false,
        show: 'Transfer',
        hide: 'Transfer',
        width: 300,
        title:  '<%=PhraseLib.Lookup("offer-rew.selectpoints", LanguageID)%>',
        autoOpen: false,
        minHeight: 10,
        minwidth: 10,
        closeText: 'X',
        closeOnEscape: true,
       
        overlay: {
        opacity: 0.65 },
         open:function(){
        $("#pointselector")[0].innerHTML=$('#hdnPointsHtml').val();
        }

    });
      $("#ed_PointsID").click(function () {
        $("#pointselector").dialog("open");
    });

     var dlgtc = $("#tcselector").dialog({

        modal: true,
        draggable: true,
        resizable: false,
        show: 'Transfer',
        hide: 'Transfer',
        width: 300,
        title: '<%=PhraseLib.Lookup("term.triggercode", LanguageID)%>',
        autoOpen: false,
        //Clean message and text box value
        open: function() {
                  var srcElement =  $get('lblCodeInfo');
                  srcElement.innerHTML="";
                  srcElement =  $get('txtCode');
                  srcElement.value="";
         },
        minHeight: 10,
        minwidth: 10,
        closeText: 'X',
        closeOnEscape: true,
        overlay: {
        opacity: 0.65 },
         open:function(){
        $("#tcselector")[0].innerHTML=$('#hdnCodeHtml').val();
        }

    });
     $("#ed_trgCode").click(function () {
        $("#tcselector").dialog("open");
    });
 }
      function isNormalInteger(str) {
          if(str != "")
              return (IsNumeric(str, false, false));
          else
              return false;
}
 function validateCode(elid) {
    
     var txtObj = $get(elid);
     var infobar = document.getElementById('lblCodeInfo');
     var txtval = txtObj.value;
     
         infobar.innerHTML = ""
         infobar.style.display = "none";
         if(txtval == ""){
             infobar.innerHTML =  '<%=PhraseLib.Detokenize("ueoffer-con.positiveinteger", LanguageID)%>';
                 infobar.style.display = "";
         }
         else if(codesettingsJSON.RangeLocked == true){
             var retVal = isNormalInteger(txtObj.value);
             if( retVal == false)
             {
                 infobar.innerHTML =  '<%=PhraseLib.Detokenize("ueoffer-con.positiveinteger", LanguageID)%>';
                 infobar.style.display = "";
                 //return;
             }
             else if (txtval >= codesettingsJSON.RangeBegin && txtval <= codesettingsJSON.RangeEnd) {
                 handlecodeclick(txtObj);
                 txtObj.value = "";
             }
             else {


                 infobar.innerHTML = '<%=strOutOfRangeMessage%>'; ;
                 infobar.style.display = "";
             }
         }
         else{
             handlecodeclick(txtObj);
             txtObj.value = "";
         }
 }

  </script>
  <form id="mainform" runat="server">
  <asp:ScriptManager ID="smScriptManager1" runat="server" ScriptMode="Auto" EnablePartialRendering="true"
    EnablePageMethods="true">
  </asp:ScriptManager>
  <div id="custom1">
  </div>
  <div id="wrap">
    <div id="custom2">
    </div>
    <a id="top" name="top"></a>
   <input type="hidden" id="hdnPointsHtml" />
   <input type="hidden" id="hdnSVHtml" />
   <input type="hidden" id="hdnCodeHtml" />
   <input type="hidden" id="hdDisableSV" runat="server" />
    <input type="hidden" id="hdDisablePoint" runat="server" />
    <asp:UpdatePanel runat="server" UpdateMode="Conditional" ID="updatePanel1">
      <ContentTemplate>
        <div id="intro">
          <h1 id='title' runat="server">
            Title</h1>
          <div id='controls'>
            <span class="temp" id="TempDisallow" runat="server">
              <asp:CheckBox ID="chkDisallow_Edit" runat="server" CssClass="tempcheck" />
              <label for="Disallow_Edit">
                <%=PhraseLib.Lookup("term.locked", LanguageID)%>
              </label>
            </span>
            <asp:Button AccessKey="s" CssClass="regular" ID="btnSave" runat="server" Text=""
              Visible="true" OnClick="btnSave_Click" />
          </div>
        </div>
        <div id="main">
          <div id="infobar" class="red-background" runat="server" visible="false">
          </div>
          <div id="column2x">
            <div class="box" id="message">
              <h2>
                <span>
                  <%= PhraseLib.Lookup("term.data", LanguageID)%>
                </span>
              </h2>
              <div style="height: 460px; overflow-y: auto;">
                <asp:Repeater ID="repXMLPassThroughData" runat="server" OnItemDataBound="repXMLPassThroughData_ItemDataBound">
                  <ItemTemplate>
                    <asp:Label ID="lblMessage" runat="server"></asp:Label>
                    <asp:Label ID="lblLanguageID" Visible="false" runat="server" Text='<%#DataBinder.Eval(Container.DataItem, "LanguageID") %>'></asp:Label>
                    <asp:Label ID="lblTierLevel" Visible="false" runat="server" Text='<%#DataBinder.Eval(Container.DataItem, "TierLevel") %>'></asp:Label>
                    <br />
                    <div class="pmsgwrap">
                      <asp:TextBox class="CPEpmsg" TextMode="MultiLine" MaxLength="2000" Columns="38" Rows="8"
                        runat="server" ID="txtData" ClientIDMode="Static" Text='<%#DataBinder.Eval(Container.DataItem, "Data") %>'
                        onfocus="javascript:srcElement=this;setCursorPosition(this);" onblur="javascript:bCapture=false;">
                      </asp:TextBox>
                      <br />
                    </div>
                  </ItemTemplate>
                  <SeparatorTemplate>
                    <hr class="hidden" />
                  </SeparatorTemplate>
                </asp:Repeater>
              </div>
            </div>
            <div class="box" id="distribution" runat="server" style="display: none;">
              <h2>
                <span>
                  <%=PhraseLib.Lookup("term.distribution", LanguageID)%>
                </span>
              </h2>
              <table style="width: 50%; display: none;" summary="<%=PhraseLib.Lookup("term.distribution", LanguageID)%>">
                <tr>
                  <td>
                    <label for="t1_value">
                      <%=PhraseLib.Lookup("term.value", LanguageID)%>:</label>
                  </td>
                  <td>
                    <asp:Repeater ID="repValues" runat="server">
                      <ItemTemplate>
                        <label>
                          <%=PhraseLib.Lookup("term.tier", LanguageID)%>
                          <%#(objOffer.NumbersOfTier>1?DataBinder.Eval(Container.DataItem, "TierLevel"):"" )+":"%></label>
                        <asp:TextBox runat="server" ID="txtValue" class="shorter" MaxLength="9" Text='<%#DataBinder.Eval(Container.DataItem, "Value") %>'></asp:TextBox>
                      </ItemTemplate>
                      <SeparatorTemplate>
                        <br />
                      </SeparatorTemplate>
                    </asp:Repeater>
                  </td>
                </tr>
                <tr>
                  <td colspan="2">
                    <hr />
                  </td>
                </tr>
                <tr>
                  <td colspan="2">
                    <asp:CheckBox ID="chkRequiredToDeliver" runat="server" Checked="true" />
                    <label for="chkRequiredToDeliver">
                      <%=PhraseLib.Lookup("ue-reward.reward-required", LanguageID)%></label>
                  </td>
                </tr>
              </table>
            </div>
          </div>
          <div id="gutter">
          </div>
          <div id="column1x">
            <div class="box" id="tags">
              <h2>
                <span>
                  <%=PhraseLib.Lookup("term.tags", LanguageID)%>
                </span>
              </h2>
              <br class="half" />
              <div id="ed_toolbar" style="background-color: #d0d0d0; text-align: center;">
                <div id="tools">
                  <input type="button" id="ed_trgCode" class="ed_button" value='<%=PhraseLib.Lookup("term.triggercode", LanguageID)%>' />
                  <input type="button" id="ed_PointsID" class="ed_button" value='<%=PhraseLib.Lookup("term.pointsid", LanguageID)%>' />
                  <input type="button" id="ed_SVID" class="ed_button" value='<%=PhraseLib.Lookup("term.storedvalueid", LanguageID)%>' />
                </div>
              </div>
            </div>
          </div>
        </div>
      </ContentTemplate>
    </asp:UpdatePanel>
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
  <div id="svselector">
    <input type="radio" id="functionradiosva" name="functionradiosv" checked="checked"
      onclick="radioclicksv();" /><label for="functionradiosv"><%=PhraseLib.Lookup("term.startingwith", LanguageID)%></label>
    <input type="radio" id="functionradiosvb" name="functionradiosv" onclick="radioclicksv();" /><label
      for="functionradiosv"><%=PhraseLib.Lookup("term.containing", LanguageID)%></label><br />
    <input type="text" class="medium" id="functioninputsv" name="functioninputsv" onkeyup="handleKeyUpsv( 200,this);"
      value="" style="width: 210px;" /><br />
    <asp:ListBox ID="lbSV" ClientIDMode="Static" runat="server" Rows="10" Style="width: 220px;"
      onclick="handlesvclick(this);" DataTextField="ProgramName" DataValueField="SVProgramID">
    </asp:ListBox>
    <br />
    <br />
    <br />
  </div>
  <div id="pointselector">
    <input type="radio" id="functionradiopointa" name="functionradiopoint" onclick="radioclickpoint();"
      checked="checked" /><label for="functionradiopoint"><%=PhraseLib.Lookup("term.startingwith", LanguageID)%></label>
    <input type="radio" id="functionradiopointb" name="functionradiopoint" onclick="radioclickpoint();" /><label
      for="functionradiopoint"><%=PhraseLib.Lookup("term.containing", LanguageID)%></label><br />
    <input type="text" class="medium" id="functioninputpoint" name="functioninputpoint"
      onkeyup="handleKeyUppoint(200,this);" value="" style="width: 210px;" /><br />
    <asp:ListBox ID="lbPoints" runat="server" ClientIDMode="Static" Rows="10" Style="width: 220px;"
      onclick="handlepointclick(this);" DataTextField="ProgramName" DataValueField="ProgramID">
    </asp:ListBox>
    <br />
    <br />
    <br />
  </div>
  <div id="tcselector">
    <p style="padding-left: 2em">
      <asp:Label ID="lblCodeInfo" runat="server" ClientIDMode="Static" CssClass="red-background"
        Style="text-align: left; display: none; float: left; color: #ffffff"></asp:Label></p>
    <asp:TextBox ID="txtCode" runat="server" Style="width: 160px"></asp:TextBox>
    <input type="button" id="btnCodeSave" onclick="validateCode('<%=txtCode.ClientID%>');"
      value='<%=PhraseLib.Lookup("term.add",LanguageID)%>' /><br />
    <asp:Label ID="lblDisplay" runat="server" Style="text-align: left; float: left; padding-left: 2em"></asp:Label><br />
    <br />
    <br />
    <asp:ListBox ID="lbCodes" runat="server" Rows="10" Style="width: 220px; overflow:auto;" onclick="copytext(this);">
    </asp:ListBox>
    <br />
    <br />
    <br />
  </div>
  </form>
</body>
</html>
