<%@ Control Language="C#" AutoEventWireup="true" CodeFile="ListBar.ascx.cs" Inherits="logix_UserControls_ListBar" %>
 <%@ Register src="Search.ascx" tagname="Search" tagprefix="uc1" %>
<%@ Register src="Paging.ascx" tagname="Paging" tagprefix="uc2" %>
 <div id="listbar">
   
   
   <uc1:Search ID="ListSearch" runat="server" />
  <uc2:Paging ID="ListPaging" runat="server" PageSize="20" /> 
  <div id="filter" title="Filter">
  
  </div>

  <hr class="hidden">
</div>