<%@ Page Language="C#" AutoEventWireup="true" CodeFile="CacheSettings.aspx.cs" Inherits="logix_cachesettings" %>

<%// version:7.3.1.138972.Official Build (SUSDAY10202) %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
  <base target="_self" />
  <title></title>
  <script type="text/javascript" src="/javascript/logix.js"></script>
  <script type="text/javascript" src="/javascript/jquery.min.js"></script>
  <script type="text/javascript">
    function ToggleSaveButton(enabled) {
      $('#btnnewgroup').attr("src", (enabled) ? '../images/save.png' : '../images/save-off.png');
    }
    window.onload = function () {
      $("#txtCacheTimeout")[0].focus();
      ToggleSaveButton(false);
    }
    function IsCacheTimeoutValueChanged(obj) {
      if (obj.value != obj.defaultValue)
        ToggleSaveButton(true);
      else
        ToggleSaveButton(false);
    }
  </script>
</head>
<body class="popup">
  <script type="text/javascript" language="javascript">
    function CheckTimeOutValue() {
      if ($("#btnnewgroup").attr('src') == '../images/save-off.png') {
        return false;
      }
      var numberRegex = /^[1-9]\d*$/;
      var inputBox = $("#txtCacheTimeout");
      var retVal = numberRegex.test(inputBox.attr('value'));
      if (retVal == false) {
        inputBox[0].focus();
        alert('<%=PhraseLib.Lookup("CPEoffer-con-tender.positivevalue", LanguageID)%>');
      }
      else {
        ToggleSaveButton(false);
      }
      return retVal;
    }
  </script>
  <form id="mainform" runat="server">
  <div id="custom1">
  </div>
  <div id="wrap">
    <div id="custom2">
    </div>
    <a id="top" name="top"></a>
    <div id="intro">
      <h1 id='title' runat="server">
        Cache Settings</h1>
      
    </div>
    <div id="main">
      <br />
      <div id="infobar" class="red-background" runat="server" visible="false">
      </div>
      <div id="column">
      <div class="box" id="autorefreshcache">
        <h2>
          <span>
            <%=PhraseLib.Lookup("term.autorefreshcache", LanguageID)%>
          </span>
        </h2>
       <table width="100%">
       <tr>
       <td width="80%">
        <asp:Label ID="lbCacheInterval" runat="server" Text="Cache Interval"/>:
        <asp:TextBox ID="txtCacheTimeout" CssClass="medium" ClientIDMode="Static" runat="server"
          onkeyup="javascript:IsCacheTimeoutValueChanged(this);" MaxLength="15" AutoPostBack="false" />
       </td>
       <td width="20%" style="text-align:right">
        <asp:LinkButton runat="server" ID="lnkSave" style="vertical-align:bottom;" OnClientClick="javascript:return CheckTimeOutValue();"
          onclick="lnkSave_Click"><img src="../images/save.png" name="btnnewgroup" id="btnnewgroup"  /></asp:LinkButton>
       </td>
       </tr>
       </table>
       
          
      </div>
      <br />
      <div class="box" id="clearcachemanual">
        <h2>
          <span>
            <%=PhraseLib.Lookup("term.clearcachedata", LanguageID)%>
          </span>
        </h2>
        <asp:Repeater ID="rptCache" runat="server" 
          onitemcommand="rptCache_ItemCommand" >
          <HeaderTemplate>
            <ul id="repeaterul" style="padding-left: 1em; list-style-type: none;">
              <asp:LinkButton ID="lnkCache" CommandArgument="-1" runat="server"><%# PhraseLib.Lookup("cache.allcache", LanguageID) %></asp:LinkButton>
             
          </HeaderTemplate>
          <ItemTemplate>
            <li>
             <asp:LinkButton ID="lnkCache" CommandArgument='<%# Eval("CachedObjectID") %>' runat="server"><%# (Eval("PhraseName") == null || Eval("PhraseName").ToString() == String.Empty) ? Eval("Name") : PhraseLib.Lookup(Eval("PhraseName").ToString(), LanguageID) %></asp:LinkButton>
             
            </li>
          </ItemTemplate>
          <FooterTemplate>
            </ul>
          </FooterTemplate>
        </asp:Repeater>
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
  </form>
</body>
</html>
