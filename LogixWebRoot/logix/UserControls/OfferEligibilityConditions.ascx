<%@ Control Language="C#" AutoEventWireup="true" CodeFile="OfferEligibilityConditions.ascx.cs"
  Inherits="logix_UserControls_OfferEligibilityConditions" ViewStateMode ="Enabled"  %>
<asp:ScriptManager ID="smScriptManager1" runat="server" ScriptMode="Auto" EnablePartialRendering="true"
  EnablePageMethods="true">
</asp:ScriptManager>
 <script type="text/javascript" src="/javascript/logix.js"></script>
<script language="javascript" type="text/javascript">
  function OptOutWindow() {
    var chkOptIn = document.getElementById('<%=chkOptIn.ClientID%>').checked;
     if (chkOptIn == false) {
//      feature = "dialogWidth:450px;dialogHeight:250px;status:no;help:no";
//      var ReturnValue = false;
//     ReturnValue = window.showModalDialog('/logix/OptInGroupMigration.aspx?OfferID=<%= OfferID%>&EngineID=<%=objOffer.EngineID%>', '', feature);
    
     window.openMiniPopup('/logix/OptInGroupMigration.aspx?OfferID=<%= OfferID%>&EngineID=<%=objOffer.EngineID%>');  

      //__doPostBack(null, null);

    }
  }
  function ConfirmDeletion() {
    var strConfirmation = '<%= PhraseLib.Lookup("condition.confirmdelete", LanguageID)%>';
    if (confirm(strConfirmation)) {
      return true;
    }
    return false;
  }

  function OpenConditionWindow(ConditionID, ConditionTypeID) {

  if(ConditionTypeID==-1)
  {
    var objddl = document.getElementById('<%=ddlOptInConditions.ClientID%>');
    if(objddl!=null)
    {
      ConditionTypeID= objddl.value;
    }
    else
    {
      return;
    }
    }
    if (ConditionTypeID == <%=CustomerGroupConditionTypeID %>) 
       openPopup('/logix/OfferEligibilityCustomerCondition.aspx?OfferID=<%= OfferID%>&OfferName=<%=HttpUtility.UrlEncode(objOffer.OfferName)%>&EngineID=<%=objOffer.EngineID%>&IsTemplate=<%=objOffer.IsTemplate%>&FromTemplate=<%= objOffer.FromTemplate%>&ConditionID=' + ConditionID + '&ConditionTypeID=' + ConditionTypeID);

    else if (ConditionTypeID == <%=PointsConditionTypeID %>) 
        openPopup('/logix/OfferEligibilityPointCondition.aspx?OfferID=<%= OfferID%>&OfferName=<%=HttpUtility.UrlEncode(objOffer.OfferName)%>&EngineID=<%=objOffer.EngineID%>&IsTemplate=<%=objOffer.IsTemplate%>&FromTemplate=<%= objOffer.FromTemplate%>&ConditionID=' + ConditionID+ '&ConditionTypeID=' + ConditionTypeID)

    else if (ConditionTypeID == <%=SVConditionTypeID %>) 
      openPopup('/logix/OfferEligibilitySVCondition.aspx?OfferID=<%= OfferID%>&OfferName=<%=HttpUtility.UrlEncode(objOffer.OfferName)%>&EngineID=<%=objOffer.EngineID%>&IsTemplate=<%=objOffer.IsTemplate%>&FromTemplate=<%= objOffer.FromTemplate%>&ConditionID=' + ConditionID+ '&ConditionTypeID=' + ConditionTypeID);

    if (ConditionID == 0)
      return false;
  }
  function ChkLockedClick(obj) {   
    if (obj.checked)
      document.forms['mainform'].IsOptInPanelLocked.value = 1;
    else
      document.forms['mainform'].IsOptInPanelLocked.value = 0;
  }

  function chkConditionLockClick(ConditionID, objChk) {
   var name1='EligibilityCondID';
   var name2='EligibilityCondVal';
   var elm1= document.getElementById(name1+ ConditionID);
   var elm2= document.getElementById(name2+ ConditionID);
   var val=0;
   if(objChk.checked)
    val=1;
   if(elm1 ==null)
   {
   $("#mainform").append("<input type='hidden' name='"+  name1 + "' id='" + name1 + ConditionID +"'  value='" + ConditionID + "' />");
   $("#mainform").append("<input type='hidden' name='"+  name2 + "' id='" + name2 + ConditionID +"'  value='" + val  + "' />");
   }
   else
   {
   elm2.value=val;
   }
   } 
  function chkOptIn(obj) {
    var objbtn = document.getElementById('<%=btnAdd.ClientID%>');
    var objddl = document.getElementById('<%=ddlOptInConditions.ClientID%>');
   if (objddl != null) {
      if (obj.checked) {
        if (objddl.options.length > 0 && objddl.options[0].text!="No Condition") {
          //objddl.style.display = '';
          //objbtn.style.display = '';
          objddl.disabled = false;
          objbtn.disabled = false;

        }
        else {
          objddl.disabled = true;
          objbtn.disabled = true;

        }
      }
      else {

        objddl.disabled = true;
        objbtn.disabled = true;
      }
    }

  }
</script>
<style type="text/css">
        .text-wrap
        {
            word-break: break-all;
            display:inline-block;
        }
    </style>
<asp:UpdatePanel ID="UpdatePanelMain" runat="server" UpdateMode="Conditional">
  <ContentTemplate>
  <asp:HiddenField ID="hdnPath" runat="server" />
    <div id="infobar" class="red-background" runat="server" visible="false">
      <asp:Label ID="lblError" runat="server"></asp:Label></div>
      
    <asp:Panel ID="panelOptIn" runat="server" Enabled="true">
      <div class="box" id="OptInConditions">
        <h2>
          
          <span>
            <asp:CheckBox ID="chkOptIn" runat="server"  AutoPostBack="false"   onclick="chkOptIn(this);"/><asp:Label ID="lblTitle" runat="server" class="tempcheck"></asp:Label>
          </span><span class="tempRequire" id="spanChkLocked" visible="false" runat="server">
            <asp:CheckBox ID="chkOptInLocked" runat="server" Text="Locked" onclick="ChkLockedClick(this)">
            </asp:CheckBox>
          </span>
        </h2>
        <asp:Label ID="lblGlobalCondition" Text="" runat="server"></asp:Label><br />
        <asp:DropDownList ID="ddlOptInConditions" runat="server" >
        </asp:DropDownList>
        <asp:Button ID="btnAdd" Text="" runat="server"  class="regular"    />
      </div>
    </asp:Panel>
    
    <asp:Panel ID="panelEligibilityCondition" runat="server" Visible="false" CssClass="box">
      
        <h2>
          <span>
            <asp:Label ID="lblOptedCondition" runat="server"><%= PhraseLib.Lookup("offereligibilityconditions.eligibile", LanguageID)%></asp:Label></span>
        </h2>
         <table id="tblHeader" class="list" summary="conditions">
          <thead>
              <tr>
                <th align="left" scope="col" class="th-del">
                  <%= PhraseLib.Lookup("term.delete", LanguageID)%>
                </th>
                <th align="left" scope="col" class="th-andor">
                  <%= PhraseLib.Lookup("term.andor", LanguageID)%>
                </th>
                <th align="left" scope="col" class="th-type">
                  <%= PhraseLib.Lookup("term.type", LanguageID)%>
                </th>
                <th align="left" scope="col" class="th-details" colspan="1">
                  <%= PhraseLib.Lookup("term.details", LanguageID)%>
                </th>
                <th align="left" scope="col" class="th-information">
                  <%= PhraseLib.Lookup("term.information", LanguageID)%>
                </th>

                <th align="left" scope="col" class="th-locked" <%=HideLockedColumn()%>>
                  <%= Copient.PhraseLib.Lookup("term.locked", LanguageID)%>
                </th>
              </tr>
               </thead>
               <tbody>
              <asp:Repeater ID="repCustomerConditions" runat="server" OnItemDataBound="repCustomerConditions_ItemDataBound" >
              <HeaderTemplate>
               <tr class="shadeddark">
              <td id="Td1" runat = "server" colspan = "4">
                <h3>
                 <%=PhraseLib.Lookup("term.customerconditions", LanguageID)%>
                </h3>
              </td>
              <td ></td>
               <td  <%=HideLockedColumn()%>>
                    
                  </td>
            </tr>
              </HeaderTemplate>
              <ItemTemplate>
              <tr class="shaded">
                  <td>
                  
                  <asp:Button ID="btnCustomerDelete" Visible="false" Enabled="false" title='<%#PhraseLib.Lookup("term.delete", LanguageID)%>' CommandArgument='<%#DataBinder.Eval(Container.DataItem, "ConditionID")%>' CommandName="CustomerDelete" runat="server"  CssClass="ex" Text="X" />
                  </td>
                  <td>
                    <asp:Label runat="server" ID="lblAndOr"> <%# DataBinder.Eval(Container.DataItem, "AndOr")%></asp:Label>
                  </td>
                  <td>
                    <a href="javascript:OpenConditionWindow('<%#DataBinder.Eval(Container.DataItem, "ConditionID")%>','<%#DataBinder.Eval(Container.DataItem, "ConditionTypeID")%>' )">
                      <%=PhraseLib.Lookup("term.customer", LanguageID)%></a>
                  </td>
                  <td>
                    <asp:Label Visible="false" runat="server" ID="lblDetails"><%#DataBinder.Eval(Container.DataItem, "Details")%></asp:Label>
                   <asp:HyperLink runat="server" ID="lnkDetails" Visible="false" CssClass="text-wrap" NavigateUrl='<%#"/logix/cgroup-edit.aspx?CustomerGroupID=" + DataBinder.Eval(Container.DataItem, "CustomerGroupID") %>'  > <%#DataBinder.Eval(Container.DataItem, "Details")%></asp:HyperLink>
                  </td>
                  <td colspan="1">
                    <%# DataBinder.Eval(Container.DataItem, "Infomation")%>
                  </td>
                  <td  <%=HideLockedColumn()%> class="templine">
                    <asp:Label ID="lblLocked" runat="server" Visible='<%#objOffer.FromTemplate%>'> <%# (DataBinder.Eval(Container.DataItem, "Locked").ToString() =="True"? "Yes": "No")%></asp:Label>
                  <asp:CheckBox ID="chkLocked" runat="server" Visible='<%#objOffer.IsTemplate %>' onclick='<%# "chkConditionLockClick(" + DataBinder.Eval(Container.DataItem, "ConditionID") + ",this)"%>' Checked='<%#DataBinder.Eval(Container.DataItem, "Locked")%>' ></asp:CheckBox>
                  </td>
                </tr>
              </ItemTemplate>
              </asp:Repeater>
              <asp:Repeater ID="repPointConditions" runat="server" 
                   OnItemCommand="repPointConditions_ItemCommand" 
                   onitemdatabound="repPointConditions_ItemDataBound"   >
              <HeaderTemplate>
               <tr class="shadeddark">
              <td id="Td1" runat = "server" colspan = "4">
                <h3>
                 <%=PhraseLib.Lookup("term.pointsconditions", LanguageID)%>
                </h3>
              </td>
              <td ></td>
               <td  <%=HideLockedColumn()%>>
                    
                  </td>
            </tr>
              </HeaderTemplate>
               <ItemTemplate>
               <asp:Label ID="lblConditionTypeID" runat="server" Visible="false" Text='<%#DataBinder.Eval(Container.DataItem, "ConditionTypeID")%>'></asp:Label>
               <asp:Label ID="lblJoinTypeID" runat="server" Visible="false" Text='<%#DataBinder.Eval(Container.DataItem, "JoinTypeID")%>' ></asp:Label>
              <tr class="shaded">
                  <td>
                  <asp:Button ID="btnPointsDelete" title='<%#PhraseLib.Lookup("term.delete", LanguageID)%>' onclick='<%# "return confirm(\"" + PhraseLib.Lookup("condition.confirmdelete", LanguageID) + "\");"%>'
                   CommandArgument='<%#DataBinder.Eval(Container.DataItem, "ConditionID")%>' CommandName="PointDelete" runat="server"  CssClass="ex" Text="X" />
                  </td>
                  <td>
                  <asp:LinkButton ID="lnkJoinType" runat="server"  CommandArgument='<%#DataBinder.Eval(Container.DataItem, "ConditionID")%>' CommandName="ChangePointJoinType"  ></asp:LinkButton>
                 
                  </td>
                  <td>
                    <a href="javascript:OpenConditionWindow('<%#DataBinder.Eval(Container.DataItem, "ConditionID")%>','<%#DataBinder.Eval(Container.DataItem, "ConditionTypeID")%>' )">
                      <%=PhraseLib.Lookup("term.points", LanguageID)%></a>
                  </td>
                  <td>
                 
                   <asp:HyperLink runat="server" ID="lnkDetails"  NavigateUrl='<%#"/logix/point-edit.aspx?ProgramGroupID=" + DataBinder.Eval(Container.DataItem, "ProgramID") %>'  > <%#DataBinder.Eval(Container.DataItem, "PointsProgram.ProgramName")%></asp:HyperLink>
                  </td>
                  <td>
                  <%# DataBinder.Eval(Container.DataItem, "Quantity")%>&nbsp;<%=PhraseLib.Lookup("term.points", LanguageID).ToLower()%>&nbsp;<%=PhraseLib.Lookup("term.required", LanguageID).ToLower()%></td>
                  <td  <%=HideLockedColumn()%> class="templine">
                    
                     <asp:Label ID="lblLocked" runat="server" Visible='<%#objOffer.FromTemplate%>'> <%# (DataBinder.Eval(Container.DataItem, "DisallowEdit").ToString() == "True" ? "Yes" : "No")%></asp:Label>
                  <asp:CheckBox ID="chkLocked" runat="server" Visible='<%#objOffer.IsTemplate %>' onclick='<%# "chkConditionLockClick(" + DataBinder.Eval(Container.DataItem, "ConditionID") + ",this)"%>' Checked='<%#DataBinder.Eval(Container.DataItem, "DisallowEdit")%>' ></asp:CheckBox>
                  </td>
                </tr>
              </ItemTemplate>
              </asp:Repeater>

               <asp:Repeater ID="repSvConditions" runat="server" 
                   onitemcommand="repSvConditions_ItemCommand" 
                   onitemdatabound="repSvConditions_ItemDataBound" >
              <HeaderTemplate>
               <tr class="shadeddark">
              <td id="Td1" runat = "server" colspan = "4">
                <h3>
                 <%=PhraseLib.Lookup("term.storedvalueconditions", LanguageID)%>
                </h3>
              </td>
              <td ></td>
               <td  <%=HideLockedColumn()%>>
                    
                  </td>
            </tr>
              </HeaderTemplate>
                <ItemTemplate>
               <asp:Label ID="lblConditionTypeID" runat="server" Visible="false" Text='<%#DataBinder.Eval(Container.DataItem, "ConditionTypeID")%>'></asp:Label>
               <asp:Label ID="lblJoinTypeID" runat="server" Visible="false" Text='<%#DataBinder.Eval(Container.DataItem, "JoinTypeID")%>' ></asp:Label>
              <tr class="shaded">
                  <td>
                  <asp:Button ID="btnSVDelete" CommandArgument='<%#DataBinder.Eval(Container.DataItem, "ConditionID")%>' onclick='<%# "return confirm(\"" + PhraseLib.Lookup("condition.confirmdelete", LanguageID) + "\");"%>' CommandName="SVDelete" runat="server"  title='<%#PhraseLib.Lookup("term.delete", LanguageID)%>' CssClass="ex" Text="X" />
                  </td>
                  <td>
                  <asp:LinkButton ID="lnkJoinType" runat="server"  CommandArgument='<%#DataBinder.Eval(Container.DataItem, "ConditionID")%>' CommandName="ChangeSVJoinType"  ></asp:LinkButton>
                 
                  </td>
                  <td>
                    <a href="javascript:OpenConditionWindow('<%#DataBinder.Eval(Container.DataItem, "ConditionID")%>','<%#DataBinder.Eval(Container.DataItem, "ConditionTypeID")%>' )">
                      <%=PhraseLib.Lookup("term.storedvalue", LanguageID)%></a>
                  </td>
                  <td>
                 
                   <asp:HyperLink runat="server" ID="lnkDetails"  NavigateUrl='<%#"/logix/SV-edit.aspx?ProgramGroupID=" + DataBinder.Eval(Container.DataItem, "ProgramID") %>'  > <%#DataBinder.Eval(Container.DataItem, "SVProgram.ProgramName")%></asp:HyperLink>
                  </td>
                  <td>
                  <%#GetStoredValueDesc(Container.DataItem as CMS.AMS.Models.SVCondition )%></td>
                  <td  <%=HideLockedColumn()%> class="templine">
                     <asp:Label ID="lblLocked" runat="server" Visible='<%#objOffer.FromTemplate%>'> <%# (DataBinder.Eval(Container.DataItem, "DisallowEdit").ToString() == "True" ? "Yes" : "No")%></asp:Label>
                  <asp:CheckBox ID="chkLocked" runat="server" Visible='<%#objOffer.IsTemplate %>' onclick='<%# "chkConditionLockClick(" + DataBinder.Eval(Container.DataItem, "ConditionID") + ",this)"%>' Checked='<%#DataBinder.Eval(Container.DataItem, "DisallowEdit")%>' ></asp:CheckBox>
                  </td>
                </tr>
              </ItemTemplate>
              </asp:Repeater>
      </tbody>
           </table>
        
    </asp:Panel>
  </ContentTemplate>
 
</asp:UpdatePanel>
