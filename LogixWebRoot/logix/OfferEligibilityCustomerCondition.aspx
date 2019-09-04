<%@ Page Language="C#" AutoEventWireup="true" CodeFile="OfferEligibilityCustomerCondition.aspx.cs"
  Inherits="logix_OfferEligibilityCustomerCondition" %>

<%// version:7.3.1.138972.Official Build (SUSDAY10202) %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
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

        var c= String.fromCharCode(event.keyCode);
        var isWordCharacter = c.match(/\w/);
        var isBackspaceOrDelete = (event.keyCode == 8 || event.keyCode == 46);

        if(isWordCharacter || isBackspaceOrDelete)
          document.getElementById('ReloadThePanel').click();
      }
      
   function SetFoucs() {
     var object = $get('functioninput');
     if (object != null) {
       SetCursorPosition(object, object.value.length);
     }

   }
   function ConfirmRegulerConditionDelete(groups) {

      
     if(confirm(groups)==true)
     {
       document.getElementById('btnDummySave').click();
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
  <form id="mainform" runat="server" >
 <asp:ScriptManager ID="smScriptManager1" runat="server" ScriptMode="Auto" EnablePartialRendering="true"  EnablePageMethods="true">
    </asp:ScriptManager>
     
<input type="hidden" id="hdnOfferID" name="OfferID" runat="server" />
<input type="hidden" id="hdnConditionID" name="ConditionID" runat="server" />
<input type="hidden" id="hdnEngineID" name="EngineID" runat="server" />
<input type="hidden" id="hdnIsTemplate" name="IsTemplate" runat="server" />
<input type="hidden" id="hdnFromTemplate" name="FromTemplate" runat="server" />
<input type="hidden" id="hdnOfferName" name="OfferName" runat="server" />
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
     <asp:Button AccessKey="s" CssClass="regular" ID="btnSave" runat="server" Text=""  OnClick="btnSave_Click"/>
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
            <%=PhraseLib.Lookup("term.customergroup", LanguageID)%>
          </span>
          
          <span class="tempRequire" runat="server" id="TempRequired" visible="false">
          <asp:CheckBox CssClass="tempcheck" ID="require_cg" runat="server" />
          <label for="require_cg"><%=PhraseLib.Lookup("term.required", LanguageID)%></label>
          </span>
         
          <span class="tempRequire" runat="server" id="FromTempRequire" visible="false">*
            <%=PhraseLib.Lookup("term.required", LanguageID)%>
          </span>
          
        </h2>
        <asp:RadioButton runat="server" ID="functionradio1" GroupName="functionradio" 
          Checked="true" oncheckedchanged="functionradio_CheckedChanged" /><label for="functionradio1"><%=PhraseLib.Lookup("term.startingwith", LanguageID)%></label>
        <asp:RadioButton runat="server" ID="functionradio2" GroupName="functionradio" 
          oncheckedchanged="functionradio_CheckedChanged" /><label for="functionradio2"><%=PhraseLib.Lookup("term.containing", LanguageID)%></label><br />
        <asp:TextBox runat="server" CssClass="medium" ID="functioninput"  ClientIDMode="Static" onkeyup="SendChangeOnServer(this);"
          MaxLength="100" AutoPostBack="false"  />
         <asp:Button runat="server" class="regular " ID="btnCreate" 
            style="padding-left: 5px; padding-right: 5px;" Text="" onclick="btnCreate_Click" />
        <br />
        <div id="cgList">
       
          <asp:ListBox ID="lstAvailable" runat="server" SelectionMode="Multiple" DataTextField="Name" DataValueField="CustomerGroupID" Rows="12" CssClass="longer"></asp:ListBox>
      
        </div>
        <br />
        <br class="half" />
        <b><%=PhraseLib.Lookup("term.selectedcustomers", LanguageID)%>:</b>
        <br />
        <asp:Button runat="server" class="regular select" ID="select1" style="padding-left: 5px; padding-right: 5px;" Text="" 
          OnClick="select1_Click" />
        <asp:Button runat="server" class="regular select" ID="deselect1" style="padding-left: 5px; padding-right: 5px;" 
          Text="" Enabled="false" OnClick="deselect1_Click" /><br />
        <br class="half" />
        <asp:ListBox ID="lstSelected" runat="server" SelectionMode="Multiple" Rows="2" CssClass="longer"  DataTextField="Name" DataValueField="CustomerGroupID" ></asp:ListBox>
       
         <br />
        <br class="half" />
        <b><%=PhraseLib.Lookup("term.excludedcustomers", LanguageID)%>:</b>
        <br />
       <asp:Button runat="server" class="regular select" ID="select2" style="padding-left: 5px; padding-right: 5px;" Text="" Enabled="false" OnClick="select2_Click" />
        <asp:Button runat="server" class="regular select" ID="deselect2" style="padding-left: 5px; padding-right: 5px;" Text=""  Enabled="false" OnClick="deselect2_Click" /><br />
        <br class="half" />
        <asp:ListBox ID="lstExcluded" runat="server" SelectionMode="Multiple" Rows="2" CssClass="longer"  DataTextField="Name" DataValueField="CustomerGroupID" ></asp:ListBox>
        <hr class="hidden" />
</div>
    </div>
    <div id="gutter" ></div>
    <div id="column2" runat="server" visible="false">
   <div class="box" id="hhoptions">
        <h2>
          <span>
            <%=PhraseLib.Lookup("term.options", LanguageID)%>
          </span>
        </h2>
        <span id="spnOffline" runat="server"  visible="false">
          <asp:CheckBox  CssClass="tempcheck" runat="server" ID="chkOffline" />&nbsp;
        <label for="chkOffline"><%=PhraseLib.Lookup("ueoffer-con-customer.metwhenoffline", LanguageID)%></label> <br />
        </span>
          <span id="spnHouseHold" runat="server"  visible="false">
         <asp:CheckBox  CssClass="tempcheck"  runat="server" ID="chkHouseHold" />&nbsp;
        <label for="chkHouseHold"><%=PhraseLib.Lookup("term.enable", LanguageID)%>&nbsp;<%=PhraseLib.Lookup("term.householding", LanguageID).ToLower()%></label></span></div>
  </div>
  </ContentTemplate>
  <Triggers>
        <asp:AsyncPostBackTrigger ControlID="ReloadThePanel" EventName="Click" />
    </Triggers>

  </asp:UpdatePanel>
  </div>
  <asp:Button ID="ReloadThePanel" runat="server" style="display:none;" ClientIDMode="Static"
    onclick="ReloadThePanel_Click"  />

    <asp:Button ID="btnDummySave" runat="server" style="display:none;" 
    ClientIDMode="Static" onclick="btnDummySave_Click"
    />
  
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
