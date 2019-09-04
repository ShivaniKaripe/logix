<%@ Page Language="C#" AutoEventWireup="true" CodeFile="OptInGroupMigration.aspx.cs" Inherits="logix_OptInGroupMigration" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
   <base target="_self"/>
  <title></title>
  <script type="text/javascript" src="/javascript/logix.js"></script>
  <script type="text/javascript" src="/javascript/jquery.min.js"></script>
  <script language="javascript" type="text/javascript">
    function CloseModel() {
//        window.returnValue = true; // the value which is return to the parent page
//        window.close();
        window.close();
        window.opener.location.reload();//to load parent window

    }

    function OpenConditionWindow() {
      
        feature = "dialogWidth:800px;dialogHeight:600px;status:no;help:no";
        var ValEngineID = document.getElementById('hdnEngineID').value;
        var ValOfferID = document.getElementById('hdnOfferID').value;
        if (ValEngineID == '9') {
            //          window.showModalDialog('UE/UEoffer-con-customer.aspx?mode=optout&OfferID=' + ValOfferID, '', feature);
            openPopup('UE/UEoffer-con-customer.aspx?mode=optout&OfferID=' + ValOfferID);
        }
        if (ValEngineID == '2') {
            //window.showModalDialog('CPEoffer-con-customer.aspx?mode=optout&OfferID=' + ValOfferID, '', feature);
            openPopup('CPEoffer-con-customer.aspx?mode=optout&OfferID=' + ValOfferID);
        }
        if (ValEngineID == '0') {
            //window.showModalDialog('Offer-con-customer.aspx?mode=optout&OfferID=' + ValOfferID, '', feature);
            openPopup('Offer-con-customer.aspx?mode=optout&OfferID=' + ValOfferID);

        }

    }
    function MigrateCustomers(operation) {


      var imgobj = document.getElementById("btnnewgroup");
      var hdnoper = document.getElementById("hdnOperation");
      hdnoper.value = operation;
      var btnsave = document.getElementById("btnSave");
      if (operation == "new") {
        if (imgobj != null && imgobj.src.indexOf('save.png') > 0) {
          btnsave.click();
        }
      }
      else {
        if (operation == "select")
          OpenConditionWindow();

        btnsave.click();
      }



    }
    function CheckChanges(obj) {
      var imgobj = document.getElementById("btnnewgroup");
      var hyper = document.getElementById("hypernewgroup");
      if (obj.value.trim() != "") {


        imgobj.src = "../images/save.png";
        hyper.disabled = false;
      }
      else {

        imgobj.src = "../images/save-off.png";
        hyper.disabled = true;
      }
    }
  </script>
</head>
<body class="popup">
    <form id="form1" runat="server">
    <input type="hidden" id="hdnOfferID" name="OfferID" runat="server" />
    <input type="hidden" id="hdnOperation" name="hdnOperation" runat="server" />
    <input type="hidden" id="hdnEngineID" name="EngineID" runat="server" />
    <asp:Button runat="server" ID="btnSave" style="display:none"  ClientIDMode="Static"
      onclick="btnSave_Click" />
    <div id="wrap">
<div id="custom2"></div>
<a id="top" name="top"></a>
<div id="intro">
    <h1 id='title' runat="server"><%=PhraseLib.Lookup("term.offer", LanguageID)%> #<%=hdnOfferID.Value%>:<%=PhraseLib.Lookup("offer.optingroupmigration",LanguageID) %></h1>
   
</div>
<div id="main" style="width:450px">


    <div id="column1" style="width:400px">
    <div id="infobar" class="red-background" runat="server" visible="false" style="width:400px"></div>
    <div class="box">
     <h2><%=PhraseLib.Lookup("offer.optinmigrationoptions",LanguageID)%>           
     </h2>
    <h3><%=PhraseLib.Lookup("term.copycustomer",LanguageID) %></h3> 
   <div style="padding-left:20px">  
  
   
   

     <label for="txtCloneGroup">
     <%=PhraseLib.Lookup("term.tonewcustomergroup", LanguageID) %></label>:
    <asp:TextBox runat="server" CssClass="medium" ID="txtNewGroup" onkeyup="CheckChanges(this);"
      MaxLength="255" AutoPostBack="false" />&nbsp;<a href="javascript:MigrateCustomers('new')"  id="hypernewgroup" disabled="true" style="vertical-align:bottom"><img src="../images/save-off.png" name="btnnewgroup" id="btnnewgroup"  />
</a>
<br /><br />
   <a href="javascript:MigrateCustomers('select')"><%=PhraseLib.Lookup("term.toselectedcustomergroup",LanguageID) %></a>
   </div>
  <br />
  <br />

   <div><a href="javascript:MigrateCustomers('discard')"><%=PhraseLib.Lookup("term.discardcustomer", LanguageID) %></a></div>
   <br />
  </div>
</div>
</div>

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
