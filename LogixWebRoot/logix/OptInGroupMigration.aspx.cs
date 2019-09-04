using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using CMS.AMS.Models;
using CMS.AMS;
using CMS;
using System.Data;
public partial class logix_OptInGroupMigration : AuthenticatedUI
{
  CMS.AMS.Contract.ICustomerGroupCondition m_CustGroupCondition;
  CMS.AMS.Contract.ICustomerGroups m_CustGroup;
  CMS.AMS.Contract.IOffer m_Offer;
  Copient.CommonInc MyCommon = new Copient.CommonInc();
  string historyString;
  protected override void OnInit(EventArgs e)
  {
    AppName = "OptInGroupMigration.aspx";
    base.OnInit(e);
  }
  protected void Page_Load(object sender, EventArgs e)
  {
    try
    {
      m_CustGroupCondition = CurrentRequest.Resolver.Resolve<CMS.AMS.Contract.ICustomerGroupCondition>();
      m_CustGroup = CurrentRequest.Resolver.Resolve<CMS.AMS.Contract.ICustomerGroups>();
      m_Offer = CurrentRequest.Resolver.Resolve<CMS.AMS.Contract.IOffer>();
      hdnOfferID.Value = Request.QueryString["OfferID"];
      hdnEngineID.Value = Request.QueryString["EngineID"];
      AssignPageTitle("term.offer", "offer.optingroupmigration", hdnOfferID.Value.ToString());

      if (!IsPostBack)
      {

        DefaultCustomerGroup = m_Offer.GetOfferDefaultCustomerGroup(hdnOfferID.Value.ConvertToLong(), hdnEngineID.Value.ConvertToInt32());
        DefaultCustomerGroup.GroupMembers = m_CustGroup.GetCustomersByGroupID(DefaultCustomerGroup.CustomerGroupID);
      }
      infobar.Visible = false;
    }
    catch (Exception ex)
    {
      infobar.InnerText = ErrorHandler.ProcessError(ex);
      infobar.Visible = true;
    }
  }
  protected void btnSave_Click(object sender, EventArgs e)
  {
    try
    {
      if (hdnOperation.Value == "new")
      {
        CreateNewGroup();
      }
      if (hdnOperation.Value == "discard")
      {
        if (!DeleteGroup())
          return;
      }
      if (hdnOperation.Value == "select")
      {
        CopyToSelectedGroup();
      }
      m_Offer.UpdateOfferStatusToModified(hdnOfferID.Value.ConvertToLong(), hdnEngineID.Value.ConvertToInt32(), CurrentUser.AdminUser.ID);
    }
    catch (Exception ex)
    {
      infobar.InnerText = ErrorHandler.ProcessError(ex);
      infobar.Visible = true;
    }
  }
  private void WriteToActivityLog()
  {
    if (MyCommon.LRTadoConn.State == ConnectionState.Closed)
      MyCommon.Open_LogixRT();
    MyCommon.Activity_Log(3, hdnOfferID.Value.ConvertToLong(), CurrentUser.AdminUser.ID, historyString);
    MyCommon.Close_LogixRT();
  }
  private bool DeleteGroup()
  {
    CMS.AMS.Models.CustomerGroupConditions cusconditions = m_CustGroupCondition.GetOfferCustomerCondition(hdnOfferID.Value.ConvertToLong(), hdnEngineID.Value.ConvertToInt32());
    if (cusconditions.IncludeCondition.Count <= 1 && hdnEngineID.Value.ConvertToInt32() != 0) //Exception for CM Engine
    {
      ////Only one Default condition exists, need to chcek if other conditions exists or not
      //if (m_Offer.IsOtherOfferConditionsExistsExceptCustomer(hdnOfferID.Value.ConvertToLong(), hdnEngineID.Value.ConvertToInt32()))
      //{
      infobar.InnerText = PhraseLib.Lookup("OptInGroup-DeleteGroup", LanguageID);
      infobar.Visible = true;
      return false;
      //}

    }


    m_Offer.DeleteOfferEligibleConditions(hdnOfferID.Value.ConvertToLong(), hdnEngineID.Value.ConvertToInt32());
    if (DefaultCustomerGroup != null)
    {
      m_Offer.DeleteCustomerConditionByGroupID(hdnOfferID.Value.ConvertToLong(), hdnEngineID.Value.ConvertToInt32(), DefaultCustomerGroup.CustomerGroupID);

      m_CustGroup.DeleteCustomerGroup(DefaultCustomerGroup.CustomerGroupID);
      historyString = PhraseLib.Lookup("cgroup-edit.delete", LanguageID) + ": " + DefaultCustomerGroup.CustomerGroupID;
      WriteToActivityLog();
    }
    ClientScript.RegisterStartupScript(this.GetType(), "pageclose", "<script>CloseModel()</script>");
    return true;
  }
  private void CreateNewGroup()
  {
    
      if (DefaultCustomerGroup != null)
      {

        if (txtNewGroup.Text.Trim().ToLower() == DefaultCustomerGroup.Name.Trim().ToLower())
        {
          infobar.Visible = true;
          infobar.InnerText = PhraseLib.Lookup("Offer-Condition-OptOut.duplicatename", LanguageID) + "-" + DefaultCustomerGroup.Name;
          return;
        }
        if (m_CustGroup.GetCustomerGroupByName(txtNewGroup.Text.Trim()) != null)
        {
          infobar.Visible = true;
          infobar.InnerText = PhraseLib.Lookup("Offer-Condition-OptOut.alreadyexist", LanguageID);
          return;
        }
        DefaultCustomerGroup.Name = txtNewGroup.Text.Trim();
        DefaultCustomerGroup.IsOptinGroup = false;
        m_CustGroup.CreateUpdateCustomerGroup(DefaultCustomerGroup);
        m_Offer.DeleteOfferEligibleConditions(hdnOfferID.Value.ConvertToLong (), hdnEngineID.Value.ConvertToInt32 ());
        historyString = PhraseLib.Lookup("history.cgroup-copy", LanguageID) + ":" + txtNewGroup.Text.Trim();
        WriteToActivityLog();
        ClientScript.RegisterStartupScript(this.GetType(), "pageclose", "<script>CloseModel()</script>");
      }
      else
      {
        infobar.InnerText = "No default group exists";
        infobar.Visible = true;
      }

   

  }
  private void CopyToSelectedGroup()
  {

    CustomerGroup custGroup = m_Offer.GetOfferDefaultCustomerGroup(hdnOfferID.Value.ConvertToLong(), hdnEngineID.Value.ConvertToInt32());
    if (custGroup == null)
    {
      //delete Customer Eligibility Conditions
      m_Offer.DeleteOfferEligibleConditions(hdnOfferID.Value.ConvertToLong(), hdnEngineID.Value.ConvertToInt32());

      //Copy all the customers from deleted default group to selected groups in customer conditions
      CustomerGroupConditions CustConditions = m_CustGroupCondition.GetOfferCustomerCondition(hdnOfferID.Value.ConvertToLong(), hdnEngineID.Value.ConvertToInt32());
      if (DefaultCustomerGroup != null && DefaultCustomerGroup.GroupMembers != null && CustConditions != null)
      {
        foreach (CMS.AMS.Models.Customer customer in DefaultCustomerGroup.GroupMembers)
        {
          m_Offer.SatisfyCustomerCondition(CustConditions, customer.CustomerPK);
        }
      }

      //Group Deleted Successfully
      m_CustGroup.DeleteCustomerGroup(DefaultCustomerGroup.CustomerGroupID);
      ClientScript.RegisterStartupScript(this.GetType(), "pageclose", "<script>CloseModel()</script>");
    }


  }
  private CustomerGroup DefaultCustomerGroup
  {
    get
    {
      return (CustomerGroup)(ViewState["DefaultGroup"]);
    }

    set
    {
      ViewState["DefaultGroup"] = value;
    }
  }

 
}