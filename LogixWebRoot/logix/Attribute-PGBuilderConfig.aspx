<%@ Page Title="" Language="C#" MasterPageFile="~/logix/User.master" AutoEventWireup="true"
  CodeFile="Attribute-PGBuilderConfig.aspx.cs" Inherits="attribute_pgbuilderconfig" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">    
  <style type="text/css">
    #divGrouping
    {
        margin-left: 10px;
        margin-top: 5px;
        margin-bottom: 5px;
    }
    .LabelHierarchyIDStyle
    {
        display: inline-block;
        width: 50px;
        text-align: right;
        margin-left: 4px;
    }
    .DDLLevelStyle
    {
        margin-left: 5px;
        max-width: 120px;
       
    }
    .DDLLevelStyle Select
    {
        max-width: 100px;
    }
  </style>
  <script type="text/javascript">
    function EnableDisableMoveButtons(obj) {
      if ($(obj).find(":selected").length == 1) {
        if ($(obj).find(":selected").index() > 0) {
          $('#<%= btnMoveUp.ClientID %>').removeAttr("disabled");
        }
        else {
          $('#<%= btnMoveUp.ClientID %>').attr('disabled', 'disabled');
        }
        if ($(obj).find(":selected").index() != (obj.length - 1)) {
          $('#<%= btnMoveDown.ClientID %>').removeAttr("disabled");
        }
        else {
          $('#<%= btnMoveDown.ClientID %>').attr('disabled', 'disabled');
        }
      }
      else {
        $('#<%= btnMoveUp.ClientID %>').attr('disabled', 'disabled');
        $('#<%= btnMoveDown.ClientID %>').attr('disabled', 'disabled');
      }
    }
    var scrollpos = null;
    function pageLoad() {
      if (scrollpos != null) {
        var selectedid = $("#<%= lbSelectedAttributeTypes.ClientID %>").prop("selectedIndex");
        $("#<%= lbSelectedAttributeTypes.ClientID %>").prop("selectedIndex", selectedid);
        scrollpos = null;
    }
    var radNoGrouping = document.getElementById("radNoGrouping");
    if(radNoGrouping && radNoGrouping.checked)
        ToggleGrouping(true);
    }
    function setListSelectedScrollPos() {
      scrollpos = $('#<%= lbSelectedAttributeTypes.ClientID %>').scrollTop();
  }
  function ToggleGrouping(flag) {
      var div = document.getElementById("divGrouping");
      if (div)
        div.disabled = flag;
  }
  function WarnForNoLevelSelection() {
      var radGrouping = document.getElementById("radGroupByLevel");
      var warnFlag = false;
      var warnMsg = '<%= PhraseLib.Lookup("pab.config.selectalllevels",LanguageID) %>';
      if (radGrouping && radGrouping.checked) {
          //find ddls with selected index = 0 and give one warning for them
          $("#divGrouping select").each(function (e, t) {
              //alert("Looping" + t.selectedIndex);
              if(t.selectedIndex == 0)
                warnFlag = true;
          });
          if(warnFlag)
          {
              alert(warnMsg);
            return false;
          }
      }
  }
  </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
  <div id="intro">
    <h1 id="htitle" runat="server">
    </h1>
    <div id="controls">
      <asp:Button ID="btnSave" class="regular select" Text="" runat="server" Visible="true"
        OnClick="btnSave_Click" OnClientClick="return WarnForNoLevelSelection()" />
    </div>
  </div>
  <div id="main">
    <div id="infobar" class="red-background" runat="server" visible="false" clientidmode="Static">
    </div>
    <div id="statusbar" class="green-background" runat="server" visible="false" clientidmode="Static">
    </div>
    <div class="columnfull">
        <div class="box" style="overflow: hidden">
            <h2 id="hGrouping" style="overflow: hidden" runat="server">
            </h2>
            <div style="float: left;">
                <asp:RadioButton ID="radGroupByLevel" runat="server" GroupName="GroupingOptions" onclick="javascript:ToggleGrouping(false)" ClientIDMode="Static" />
                <br />
                <div id="divGrouping" runat="server" clientidmode="Static">
                    <asp:Label ID="lblGroupingNotAvailable" runat="server" Visible="false"></asp:Label>
                    <asp:Repeater ID="repGrouping" runat="server">
                        <ItemTemplate>
                            <asp:Label ID="lblHierarchyID" runat="server" CssClass="LabelHierarchyIDStyle" Text='<%# DataBinder.Eval(Container.DataItem, "ExtHierarchyID") %>' /> : 
                            <asp:DropDownList ID="ddlLevels" runat="server" CssClass="DDLLevelStyle" OnDataBound="ddlLevels_DataBound"
                            DataSource='<%# DataBinder.Eval(Container.DataItem, "ListLevelNames") %>'></asp:DropDownList>
                            <br />
                            <br />
                        </ItemTemplate>
                    </asp:Repeater>
                </div>
                <br />
                <asp:RadioButton ID="radNoGrouping" runat="server" GroupName="GroupingOptions" onclick="javascript:ToggleGrouping(true)" ClientIDMode="Static" />
            </div>
        </div>
      <div class="box" style="overflow: hidden">
        <h2 id="hselection" runat="server" style="float: left;">
        </h2>
        <br clear="all" />
        <div style="float: left; width: 300px;">
          <br />
          <asp:Label ID="lblAvailableAttributeType" runat="server" />
          <br />
          <asp:ListBox ID="lbAvailableAttributeTypes" runat="server" SelectionMode="Multiple"
            DataTextField="AttributeName" DataValueField="AttributeTypeID" Rows="12" CssClass="longerwideselector">
          </asp:ListBox>
          <asp:Label ID="lblContainsProducts" runat="server" />
        </div>
        <div style="float: left; width: 120px; margin-right: 4px">
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <center>
            <asp:Button ID="select1" class="regular select" runat="server" Text="Button" OnClick="select1_Click" />
          </center>
          <br />
          <br />
          <br />
          <br />
          <center>
            <asp:Button ID="deselect1" class="regular deselect" runat="server" Text="Button"
              OnClick="deselect1_Click" /></center>
        </div>
        <div style="float: left; width: 300px;">
          <br />
          <asp:Label ID="lblSelectedAttributeType" runat="server" />
          <br />
          <asp:ListBox ID="lbSelectedAttributeTypes" runat="server" SelectionMode="Multiple"
            DataTextField="AttributeName" DataValueField="AttributeTypeID" Rows="12" CssClass="longerwideselector"
            onchange="javascript:EnableDisableMoveButtons(this)"></asp:ListBox>
          <br />
          <br />
          <div style="float: left;">
            <asp:Button ID="btnMoveUp" class="regular select" runat="server" Text="Button" OnClientClick="javascript:setListSelectedScrollPos()"
              OnClick="btnMoveUp_Click" />
          </div>
          <div style="float: right; margin-right: 20px">
            <asp:Button ID="btnMoveDown" class="regular select" runat="server" Text="Button"
              OnClientClick="javascript:setListSelectedScrollPos()" OnClick="btnMoveDown_Click" />
          </div>
          <br />
          <br />
          <br />
          <br />
        </div>
      </div>
    </div>
  </div>
</asp:Content>
