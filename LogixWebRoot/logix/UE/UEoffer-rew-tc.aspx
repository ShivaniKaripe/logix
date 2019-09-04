<%@ Page Language="C#" AutoEventWireup="true" Debug="true" CodeFile="UEoffer-rew-tc.aspx.cs" Inherits="logix_UE_UEoffer_rew_tc" %>
<%@ Reference Control="~/logix/UserControls/MultiLanguagePopup.ascx" %>
<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
     <script type="text/javascript" src="/javascript/logix.js"></script>
  <script type="text/javascript" src="/javascript/jquery.min.js"></script>
</head>
<body class="popup" onunload="ChangeParentDocument();">
    <script type="text/javascript">

        function CloseModel() {
            window.close();

        }
        function ChangeParentDocument() {
            opener.location = '/logix/UE/UEoffer-rew.aspx?OfferID=<%= OfferID %>';

        }
            function IsProgramSelected() {
           
                var programSel = $("#lstAvailable").val();
                if (programSel == null) {
                    alert('<%= PhraseLib.Lookup("CPE-rew-membership.selectpoints", LanguageID) %>');
                }
           
        }


        var lockedFields = new Array();
        function SendChangeOnServer(obj) {
            var searchText = $.trim(obj.value);
            var c = String.fromCharCode(event.keyCode);
            var isWordCharacter = c.match(/\w/);
            var isSpecialChar = (event.keyCode == 192 || event.keyCode == 188 || event.keyCode == 190 || event.keyCode == 222 ||
                                    event.keyCode == 16 || event.keyCode == 189 || event.keyCode == 187 || event.keyCode == 186
                                    || event.keyCode == 191 || event.keyCode == 219 || event.keyCode == 220 || event.keyCode == 221);
            var isBackspaceOrDelete = (event.keyCode == 8 || event.keyCode == 46);
            var isAllowedBlankSpace = (event.keyCode == 32 && searchText.length > 0);

            if (isWordCharacter || isSpecialChar || isBackspaceOrDelete || isAllowedBlankSpace)
                document.getElementById('ReloadThePanel').click();
        }
        function updatepage(str) {
            document.getElementById("results").innerHTML = str;
        }

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

        function SetPageLockStatus() {
            if ($("#chkDisallow_Edit").is(":checked") == true)
                $("#hdnPageLock").val("Locked");
            else if ($("#chkDisallow_Edit").is(":checked") == false)
                $("#hdnPageLock").val("Unlocked");
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
        $(document).ready(function () {
            var object = $get('functioninput');
            if (object != null) {
                object.focus();
            }
            updateHiddenFieldValue();
        });
        
    </script>
    <form id="mainform" runat="server">
    <asp:ScriptManager ID="smScriptManager1" runat="server" EnablePartialRendering="true"
    EnablePageMethods="true">
  </asp:ScriptManager>
         <input type="hidden" id="hdnLockedTemplateFields" value="" runat="server" />
         <input type="hidden" id="hdnPageLock" value="Unlocked" runat="server" />
        <div id="results" style="position: absolute; z-index: 99; top: 31px; right: 21px;">
        </div>
        <asp:UpdatePanel ID="UpdatePanelMain" runat="server" UpdateMode="Conditional">
      <ContentTemplate>
    <div id="intro">
        <h1 id="title" runat="server"></h1>
     <div id='controls'>
                <span class="temp" id="TempDisallow" runat="server">
                    <asp:CheckBox ID="chkDisallow_Edit" runat="server" CssClass="tempcheck" onclick="updateTableLockStatus();" />
                    <label for="chkDisallow_Edit">
                        <%=PhraseLib.Lookup("term.locked", LanguageID)%>
                    </label>
                    <a href="javascript:ShowFieldList();" title='<%=PhraseLib.Lookup("cpeoffer-rew-disc-clicktoview", LanguageID)%>'>&#9660;</a> </span>
      <asp:Button AccessKey="s" CssClass="regular" ID="btnSave" runat="server" Text=""
        Visible="true" OnClick="btnSave_Click"  />
    </div>
    </div>
        <div id="main">
          
        <div id="infobar" class="red-background" runat="server" clientidmode="Static" visible="false" />
            <div id="column1">
                <div id="selector" class="box">
                    <h2> <span>
                <%=PhraseLib.Lookup("term.trackablecouponprogram", LanguageID)%>
              </span></h2>
                    <asp:RadioButton runat="server" ID="functionradio1" GroupName="functionradio" Checked="true" /><label
                        for="functionradio1"><%=PhraseLib.Lookup("term.startingwith", LanguageID)%></label>
                    <asp:RadioButton runat="server" ID="functionradio2" GroupName="functionradio" /><label
                        for="functionradio2"><%=PhraseLib.Lookup("term.containing", LanguageID)%></label><br />
                   <asp:UpdatePanel ID="UpdatePanel1" runat="server" UpdateMode="Conditional">
                       <ContentTemplate>
                            <asp:TextBox runat="server" CssClass="medium" ID="functioninput" ClientIDMode="Static"
                  onkeyup="SendChangeOnServer(this);" MaxLength="100" AutoPostBack="false" />
                       </ContentTemplate>
                   </asp:UpdatePanel>
               <br />
               <asp:UpdatePanel ID="UpdatePanel2" runat="server" UpdateMode="Conditional">
              <ContentTemplate>
                       <asp:Button ID="ReloadThePanel" runat="server" Style="display: none;"  ClientIDMode="Static"
                  OnClick="ReloadThePanel_Click" />
                    <div id="tcpList">
                         <br class="half" />
                        <asp:ListBox ID="lstAvailable" runat="server" SelectionMode="Single" DataTextField="Name"
                            DataValueField="ProgramID" Rows="10" CssClass="longer"></asp:ListBox>
                    </div>
                    <br class="half" />
                    <asp:Repeater ID="repTier_selectedTCP" runat="server" OnItemDataBound="repTier_selectedTCP_ItemDataBound"
                        OnItemCommand="repTier_selectedTCP_ItemCommand">
                        <ItemTemplate>
                           <b>
                            <%#(Offer.NumbersOfTier > 1 ? PhraseLib.Lookup("term.tier", LanguageID):"" ) %>
                                <%#(Offer.NumbersOfTier > 1 ? DataBinder.Eval(Container.DataItem, "TierLevel") + ":" : "")%></b>
                             
                        <asp:Button runat="server" class="regular select" ID="btnselect" Text="" OnClientClick="IsProgramSelected();" CommandArgument='<%# Eval("TierLevel") %>' CommandName="Select"></asp:Button>
                        <asp:Button runat="server" class="regular select" ID="btndeselect" Text="" CommandArgument='<%# Eval("TierLevel") %>' CommandName="Deselect"></asp:Button>
                  <br />
               
                        <asp:ListBox ID="lstSelected" runat="server" SelectionMode="Single" DataTextField="Name"
                            DataValueField="ProgramID" Rows="5" CssClass="longer" style="width:295px;"></asp:ListBox>
                  <br class="half" />
                        </ItemTemplate>
                        <SeparatorTemplate><br class="half" /></SeparatorTemplate>
                    </asp:Repeater>

              </ContentTemplate>
              <Triggers>
                <asp:AsyncPostBackTrigger ControlID="ReloadThePanel" EventName="Click" />
              </Triggers>

               </asp:UpdatePanel>
              
                   
                     
                </div>
               
                <div id="options" runat="server" class="box" style="width:375px;">
                    <h2><span><%=PhraseLib.Lookup("term.advancedoptions", LanguageID)%></span></h2>
                   <br />
                     <asp:Label ID="lblprinttype" runat="server" ClientIDMode="Static"></asp:Label>&nbsp;&nbsp;
                    <asp:DropDownList ID="ddlprinttype" runat="server" AutoPostBack="true" ClientIDMode="Static" DataTextField="Phrase" DataValueField="PrintTypeID" 
                        OnSelectedIndexChanged="ddlprinttype_SelectedIndexChanged"></asp:DropDownList>&nbsp;&nbsp;
                    <asp:Label ID="lblsubtype" runat="server" ClientIDMode="Static" Visible="false"></asp:Label>&nbsp;&nbsp;
                     <asp:DropDownList ID="ddlsubtype" runat="server" ClientIDMode="Static" DataTextField="TypeDescription" DataValueField="PrintSubTypeID" 
                         Visible="false"></asp:DropDownList>
                    <br />
                    <br />
                    <div id="coupon" runat="server">
                        <fieldset style="width:300px; border-color:#fff;">
                         <legend> <span><%=PhraseLib.Lookup("coupon.delivery", LanguageID) %></span></legend>   
                       
                        <br />
                            <asp:RadioButtonList ID="deliverytypes" runat="server" DataValueField="TCDeliveryTypeID" DataTextField="Phrase">
                            </asp:RadioButtonList>
                        </fieldset>
                   
                        
                    </div>
                    <br />
                    
                    
                   
                </div>
            </div>
            <div id="gutter"> </div>
                <div id="column2" runat="server">
                    <div id="Distribution" class="box">
               <h2><span><%=PhraseLib.Lookup("term.distribution", LanguageID)%></span></h2>
                       
                         <asp:Repeater ID="repTier_Desc" runat="server" OnItemDataBound="repTier_Desc_ItemDataBound" OnItemCreated="repTier_Desc_ItemCreated">
                             
              <ItemTemplate>
                  <b>
                            <%#(Offer.NumbersOfTier > 1 ? "<br />" + PhraseLib.Lookup("term.tier", LanguageID):"" ) %>
                                <%#(Offer.NumbersOfTier > 1 ? DataBinder.Eval(Container.DataItem, "TierLevel") + ":" : "")%></b><br />
                             <table>
                                <td  style="text-align: right;padding-bottom:15px"><asp:Label ID="lbldesc" runat="server" ClientIDMode="Static" Text="Description : "></asp:Label></td>
                                <td ><div align="centre"  id="divBuyDescriptionMLI" runat="server" /></td>                             
                            </table>                  
                                
                        </ItemTemplate> 
                             <SeparatorTemplate></SeparatorTemplate>
                             </asp:Repeater>
                        <hr />
                        <asp:CheckBox ID="successful" runat="server" ClientIDMode="Static" Checked="true" />&nbsp;&nbsp;
                        <asp:Label ID="lblsucdelivery" runat="server" ClientIDMode="Static"></asp:Label>
           </div>
                </div>
           <br />
             </ContentTemplate>
    </asp:UpdatePanel>
        </div>
    </form>
</body>
</html>
