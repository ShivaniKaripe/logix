<%@ Page Language="C#" AutoEventWireup="true" CodeFile="OfferEligibilityPointCondition.aspx.cs" Inherits="OfferEligibilityPointCondition" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
  <base target="_self"/>
  <title></title>
  <script type="text/javascript" src="/javascript/logix.js"></script>
   <script type="text/javascript" src="/javascript/jquery.min.js"></script>
   <script type="text/javascript" language="javascript">
     function CloseModel() {
         window.close();
         window.opener.location.reload();//reload the parent window
     }
     function SendChangeOnServer(obj) {
       //use keyup instead keypress because:
       //- keypress will not work on backspace and delete
       //- keypress is called before the character is added to the textfield (at least in google chrome) 
       var searchText = $.trim(obj.value);

       var c = String.fromCharCode(event.keyCode);
       var isWordCharacter = c.match(/\w/);
       var isBackspaceOrDelete = (event.keyCode == 8 || event.keyCode == 46);

       if (isWordCharacter || isBackspaceOrDelete)
         document.getElementById('ReloadThePanel').click();
     }
     


     function SetFoucs() {
       var object = $get('functioninput');
       if (object != null) {
         SetCursorPosition(object, object.value.length);
       }
     }

     $(document).ready(function () {
       var object = $get('functioninput');
       if (object != null) {
         object.focus();
       }
     });


 function AlertMessage(msg) {
     alert(msg);
     return false;
 }
  </script>
</head>
<body class="popup" >
<script type="text/javascript" language= "javascript"  >
    function ValidateQty() {
        var retVal = true;
        var qtyElem = document.getElementById('txtValueNeeded');
        var qtyVal = qtyElem.value.replace(/^\s+|\s+$/g, '');
        if (qtyVal == "" || isNaN(qtyVal) || !isInteger(qtyVal) || parseInt(qtyVal) == 0) {
            retVal = false;
            alert('<%=PhraseLib.Lookup("CPEoffer-con-point.positiveinteger", LanguageID) %>');
            qtyElem.focus();
            qtyElem.select();
        }
        return retVal;
    }
</script>
  <form id="mainform" runat="server" >
 <asp:ScriptManager ID="smScriptManager1" runat="server" ScriptMode="Auto" EnablePartialRendering="true"  EnablePageMethods="true">
    </asp:ScriptManager>
 <div id="custom1"></div>
<div id="wrap">
<div id="custom2"></div>
<a id="top" name="top"></a>
  <div id="intro">
    <h1 id='title' runat="server"> Title</h1>
    <div id='controls'>
      <span class="temp" id="TempDisallow" runat="server">
    <asp:CheckBox ID="chkDisallow_Edit" runat="server" CssClass="tempcheck" />
        <label for="Disallow_Edit" >
        <%=PhraseLib.Lookup("term.locked", LanguageID)%>
        </label>
      </span>
     <asp:Button AccessKey="s" CssClass="regular" ID="btnSave" runat="server" Text="" Visible="true" OnClientClick= "javascript:return ValidateQty();" OnClick="btnSave_Click"/>
    </div>
  </div>
  <div id="main">
    <asp:UpdatePanel ID="UpdatePanelMain" runat="server" UpdateMode="Conditional">
    <ContentTemplate>
    <div id="infobar" class="red-background" runat="server" visible="false"></div>
    <div id="column1">
       
     
    <div class="box" id="selector">
 
        <h2>
          <span>
            <%=PhraseLib.Lookup("term.pointsprogram", LanguageID)%>
          </span>          
         
        </h2>
        <asp:RadioButton runat="server" ID="functionradio1" GroupName="functionradio" 
          Checked="true" /><label for="functionradio1"><%=PhraseLib.Lookup("term.startingwith", LanguageID)%></label>
        <asp:RadioButton runat="server" ID="functionradio2" GroupName="functionradio" 
           /><label for="functionradio2"><%=PhraseLib.Lookup("term.containing", LanguageID)%></label><br />
        <asp:TextBox runat="server" CssClass="medium" ID="functioninput"  ClientIDMode="Static" onkeyup="SendChangeOnServer(this);"
          MaxLength="100" AutoPostBack="false"  />
           <asp:Button runat="server" class="regular " ID="btnCreate" 
            style="padding-left: 5px; padding-right: 5px;" Text="" onclick="btnCreate_Click" />
          <br />
        <div id="cgList">
       
          <asp:ListBox ID="lstAvailable" runat="server" SelectionMode="Single" DataTextField="ProgramName" DataValueField="ProgramID" Rows="12" CssClass="longer"></asp:ListBox>
      
        </div>
        <br />      
        <asp:Button runat="server" class="regular select" ID="select1" style="padding-left: 5px; padding-right: 5px;" Text= "" 
          onclick="select1_Click"></asp:Button>
        <asp:Button runat="server" class="regular select" ID="deselect1" style="padding-left: 5px; padding-right: 5px;" Text= "" Enabled="false" onclick="deselect1_Click"/><br />
        <br class="half" />
        <asp:ListBox ID="lstSelected" runat="server" SelectionMode="Single" Rows="2" CssClass="longer"  DataTextField="ProgramName" DataValueField="ProgramID" ></asp:ListBox>
        <br />       
</div>
    </div>
    <div id="gutter" ></div>
    <div id="column2" runat="server">
   <div class="box" id="hhoptions">
        <h2>
          <span>
            <%=PhraseLib.Lookup("term.value", LanguageID)%>
          </span>
        </h2>
        <tr>
        <td><asp:Label ID ="lblValueNeeded" runat="server"></asp:Label></td>
        </tr>
        <tr>
        <td><asp:TextBox runat ="server" ID ="txtValueNeeded"></asp:TextBox></td>
        </tr>
        <table>
        </table>
      </div>
  </div>
  </ContentTemplate>
  <Triggers>
        <asp:AsyncPostBackTrigger ControlID="ReloadThePanel" EventName="Click" />
    </Triggers>

  </asp:UpdatePanel>
  </div>
  <asp:Button ID="ReloadThePanel" runat="server" style="display:none;" ClientIDMode="Static"
    onclick="ReloadThePanel_Click"  />
  <a id="bottom" name="bottom"></a>
<div id="footer">
   <%=PhraseLib.Lookup("about.copyright", LanguageID)%>
</div>
<div id="custom3"></div>
</div> <!-- End wrap -->
<div id="custom4"></div>
</form>
</body>
</html>

