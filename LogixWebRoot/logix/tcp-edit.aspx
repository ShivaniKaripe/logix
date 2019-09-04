<%@ Page Title="" Language="C#" MasterPageFile="~/logix/User.master" AutoEventWireup="true" 
  CodeFile="tcp-edit.aspx.cs" Inherits="logix_tcp_edit" %>

<%@ Register TagPrefix="uc" TagName="UI" Src="~/logix/UserControls/Notes.ascx" %>
<%@ Register TagPrefix="uc" TagName="Popup" Src="~/logix/UserControls/Notes-Popup.ascx" %>
<%@ Register Src="UserControls/CollapsableDiv.ascx" TagName="CollapsableDiv" TagPrefix="uc1" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
  <script type="text/javascript">

    $(document).ready(function () {
      $('#txtDescription').keypress(function (e) {
        if ($(this).val().length > 999) {
          e.preventDefault();
        }
      });
    });

    if (window.captureEvents) {
      window.captureEvents(Event.CLICK);
      window.onclick = function (e) {
        handlePageClickNew(e, 'divActionMenu', 'btnAction');
      }
    } else {

      document.onclick = function (e) {
        handlePageClickNew(e, 'divActionMenu', 'btnAction');
      }
    }

    function pageLoad() {
      $("#txtDatepicker").datepicker({
        showOn: "button",
        buttonImage: "../images/calendar.png",
        buttonImageOnly: true,
        showButtonPanel: true,
        changeMonth: true,
        changeYear: true
        });
    }

    function toggleDropdown() {
      $("#divActionMenu").toggle();
      if ($("#divActionMenu").css("display") == "none") {
        $("#btnAction").prop('value', '<%=PhraseLib.Lookup("term.actions",LanguageID)%>' + ' ▼');
      }
      else {
        $("#btnAction").prop('value', '<%=PhraseLib.Lookup("term.actions",LanguageID)%>' + ' ▲');
      }
      return false;
    }
  </script>
</asp:Content>
<asp:Content ID="Content2" ClientIDMode="Static" ContentPlaceHolderID="ContentPlaceHolder1"
  runat="Server">
  <div id="intro">
    <h1 id="htitle" runat="server">
    </h1>
    <div id="controls">
      <asp:Button ID="btnSave" class="regular select" Text="" runat="server" Visible="true" OnClick="btnSave_Click" />
      <asp:Button ID="btnAction" ClientIDMode="Static" runat="server" class="regular actionbutton" OnClientClick="return toggleDropdown();" />
      <div id="divActionMenu" class="actionsmenunew">
        <asp:Button ID="btnUpdate" AccessKey ="s" class="regular select" runat="server" OnClick="btnSave_Click" />
        <asp:Button ID="btnDelete" AccessKey ="d" class="regular deletebutton" runat="server" OnClick="btnDelete_Click" />
        <asp:Button ID="btnNew" AccessKey ="n" class="regular select" runat="server" OnClick="btnNew_Click" />
      </div>
      <uc:UI ID="ucNotesUI" runat="server" />
    </div>
  </div>
  <div id="main">
    <AMSControls:AMSValidationSummary ID="vsError" ForeColor="" EnableClientScript="false"  CssClass="errsummary red-background" runat="server"  />
    <div id="column1">
      <div class="box">
        <h2 id="hidentification" runat="server" style="float: left;">
        </h2>
        <uc1:CollapsableDiv ID="Collapsadividentification" runat="server" TargetDivID="identityBody" />
        <br clear="all" />
        <div id="identityBody">
          <asp:Label ID="lblName" runat="server" />
          <asp:TextBox ID="txtName" runat="server" class="longest" MaxLength="200" />
           <asp:RequiredFieldValidator EnableClientScript="false" ControlToValidate ="txtName" runat="server"  ID="requirefieldName" Display ="None" />
        
          <asp:Label ID="lblExternalID" runat="server"  />
          <asp:TextBox ID="txtExternalID" runat="server" class="longest" MaxLength="20" />
          <asp:RequiredFieldValidator EnableClientScript="false" ControlToValidate ="txtExternalID"  runat="server"  ID="requirefieldExternalID" Display ="None" />
         
          <asp:Label ID="lblDescription" runat="server" />
          <textarea ID="txtDescription" runat="server" rows="3" cols="48"></textarea>
          <asp:RegularExpressionValidator EnableClientScript="false" Display = "None" ControlToValidate = "txtDescription" ID="DescriptionLengthValidator" ValidationExpression = "^[\s\S]{0,1000}$" runat="server" ></asp:RegularExpressionValidator>
          <br />
          <small>
            <asp:Label ID="lblDescriptionLimitMsg" runat="server" />
          </small>
          <br />
          <br />
          <asp:Label ID="lblExpire" runat="server" />
          <asp:Label ID="ExpireDate" runat="server" />
          <br />
        </div>
      </div>
      <div class="box">
        <h2 id="hRedemptioninformation" runat="server" style="float: left;" />
        <uc1:CollapsableDiv ID="CollapsadivRedemptioninformation" runat="server" TargetDivID="redemptionBody" />
        <br clear="all" />
        <div id="redemptionBody">
          <asp:Label ID="lblMaxRedempCount" runat="server" />
          <asp:TextBox ID="txtMaxRedempCount" runat="server" class="short" MaxLength="3" Text="1" />
          <asp:RequiredFieldValidator EnableClientScript="false" ControlToValidate ="txtMaxRedempCount" runat="server"  ID="requirefieldMaxRedempCount" Display ="None" />
          <asp:RangeValidator  EnableClientScript="false" ID="RangeValidatorMaxRedempCount" runat ="server" ControlToValidate ="txtMaxRedempCount" Display ="None" Type="Integer" MinimumValue="1" MaximumValue="255"  />
          <small>
            <asp:Label ID="lblMaxMinInfoMsg" runat="server" />
          </small>
          <br />
        </div>
      </div>
      <div class="box">
        <h2 id="hCouponUploadSumm" runat="server" style="float: left;">
        </h2>
        <uc1:CollapsableDiv ID="CollapsableDivCouponUploadSumm" runat="server" TargetDivID="couponUploadbody" />
        <br clear="all" />
        <div id="couponUploadbody">
          <asp:Label ID="lblLastUpload" runat="server" />
          <asp:Label ID="lblCouponUploadDate" runat="server" />
          <br />
          <asp:Label ID="lblStatusMsg" runat="server" />
          <asp:Label ID="lblCouponuploadSumm" runat="server" />
          <br />
          <br />
          <asp:Label ID="lblCouponusCount" runat="server" />
          <asp:Label ID="CouponCount" runat="server" />
          <br />
        </div>
      </div>
    </div>
    <div id="gutter">
    </div>
    <div id="column2">
      <div class="box" id="divExpiration" runat="server" visible="false">
        <h2 id="hExpiration" runat="server" style="float: left;" />
        <uc1:CollapsableDiv ID="CollapsableDivExpiration" runat="server" TargetDivID="expirationBody" />
        <br clear="all" />
        <div id="expirationBody">
          <asp:Label ID="lblExpirationType" runat="server" />
          <asp:DropDownList ID="ddlExpireTypes" AutoPostBack="True" runat="server" Style="height: 100%;max-width:200px;" OnSelectedIndexChanged="ddlExpireTypes_SelectedIndexChanged" />
          <br />
          <br />
          <asp:Panel ID="pnlExpirationPeriod" runat="server" >
             <asp:Label ID="lblExpirationPeriodType" runat="server" />
             <asp:DropDownList ID="ddlExpirePeriodTypes" AutoPostBack="False" runat="server" Style="height: 100%;max-width:150px;" />
             <br />
             <br />

             <asp:Label ID="lblExpirationPeriod" runat="server" />
             <asp:TextBox ID="txtExpirationPeriod" runat="server" class="short" MaxLength="6" Text="0" />
             <asp:RangeValidator EnableClientScript="false" ID="rvExpirePeriod" runat="server" ControlToValidate="txtExpirationPeriod" Display="None" Type="Integer" MinimumValue="0" MaximumValue="999999" />
             <br />
             <br />
          </asp:Panel>
          <asp:Panel ID="pnlExpirationDate" runat="server" >
             <asp:Label ID="lblExpirationTime" runat="server" />
             <asp:DropDownList ID="ddlExpireTimeHours" AutoPostBack="True" OnSelectedIndexChanged="handleExpireDateTimeChange" runat="server" />:
             <asp:DropDownList ID="ddlExpireTimeMinutes"  AutoPostBack="True" OnSelectedIndexChanged="handleExpireDateTimeChange" runat="server" />
             <br />
             <br />
             <asp:Label ID="lblExpirationDatePicker" runat="server" />
             <asp:TextBox id="txtDatepicker" OnTextChanged="handleExpireDateTimeChange" Style="height: 100%;max-width:200px;" runat="server"/>
             <br />
          </asp:Panel>
        </div>
      </div>
      <div class="box" id="offers">
        <h2 id="hAssociatedoffer" runat="server" style="float: left;">
        </h2>
        <uc1:CollapsableDiv ID="CollapsableDivAssociatedoffer" runat="server" TargetDivID="divOffers" />
        <br clear="all" />
        <div class="boxscrollhalf" id="divOffers" runat="server" EnableViewState ="true">
          <asp:Label ID="lblAssociatedOffer" runat ="server" />
        </div>
      </div>
    </div>
  </div>
  <uc:Popup ID="ucNotes_Popup" runat="server" />
</asp:Content>
