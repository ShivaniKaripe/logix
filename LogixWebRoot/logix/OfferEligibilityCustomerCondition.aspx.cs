using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using CMS;
using CMS.AMS;
using CMS.AMS.Contract;
using System.Data;
using CMS.AMS.Models;
public partial class logix_OfferEligibilityCustomerCondition : AuthenticatedUI
{
    #region Variables

    bool IsTemplate = false;
    long OfferID = 0;
    bool FromTemplate = false;
    bool DisabledAttribute = false;
    int EngineID = -1;
    long ConditionID = 0;
    string historyString;


    IOffer m_offer;
    ICustomerGroups m_CustGroup;
    ICustomerGroupCondition m_CustCondition;
    IOfferApprovalWorkflowService m_OAWService;

    Copient.CommonInc MyCommon = new Copient.CommonInc();
    Copient.LogixInc Logix = new Copient.LogixInc();
    bool bCreateGroupOrProgramFromOffer = false;
    bool isTranslatedOffer = false;
    bool bEnableRestrictedAccessToUEOfferBuilder = false;
    bool bOfferEditable = false;
    bool bEnableAdditionalLockoutRestrictionsOnOffers = false;

    protected override void OnInit(EventArgs e)
    {
        AppName = "OfferEligibilityCustomerCondition-Select.aspx";
        base.OnInit(e);

    }
    /// <summary>
    /// Customer Groups exist in regualar customer condition
    /// </summary>
    private List<CMS.AMS.Models.CustomerGroup> IncludedConditionGroup
    {

        get
        {
            return ViewState["IncludedRegularConditionCustomerGroups"] as List<CMS.AMS.Models.CustomerGroup>;
        }
        set
        {
            ViewState["IncludedRegularConditionCustomerGroups"] = value;
        }

    }
    /// <summary>
    /// Customer Groups exist in regualar excluded condition
    /// </summary>
    private List<CMS.AMS.Models.CustomerGroup> ExcludedConditionGroup
    {

        get
        {
            return ViewState["ExcludedRegularConditionCustomerGroups"] as List<CMS.AMS.Models.CustomerGroup>;
        }
        set
        {
            ViewState["ExcludedRegularConditionCustomerGroups"] = value;
        }

    }

    private List<CMS.AMS.Models.CustomerGroup> IncludedGroup
    {

        get
        {
            return ViewState["IncludedGroup"] as List<CMS.AMS.Models.CustomerGroup>;
        }
        set
        {
            ViewState["IncludedGroup"] = value;
        }

    }
    private List<CMS.AMS.Models.CustomerGroup> ExcludedGroup
    {

        get
        {
            return ViewState["ExcludedGroup"] as List<CMS.AMS.Models.CustomerGroup>;
        }
        set
        {
            ViewState["ExcludedGroup"] = value;
        }

    }
    private List<CMS.AMS.Models.CustomerGroup> AvailableFilteredCustomerGroup
    {

        get
        {
            return Session["AvailableFilteredCustomerGroup"] as List<CMS.AMS.Models.CustomerGroup>;
        }
        set
        {
            Session["AvailableFilteredCustomerGroup"] = value;
        }

    }
    private CMS.AMS.Models.CustomerGroupConditions OfferEligibileCustomerGroupCondition
    {
        get
        {
            return ViewState["OfferEligibileCustomerGroupCondition"] as CMS.AMS.Models.CustomerGroupConditions;
        }
        set
        {
            ViewState["OfferEligibileCustomerGroupCondition"] = value;
        }
    }

    private List<CMS.AMS.Models.CustomerGroup> AllGroups
    {
        get
        {
            return Session["AllGroups"] as List<CMS.AMS.Models.CustomerGroup>;
        }
        set
        {
            Session["AllGroups"] = value;
        }
    }
    #endregion Variables

    #region Protected Methods

    protected void Page_Load(object sender, EventArgs e)
    {
        try
        {


            Response.Expires = -1;
            infobar.Visible = false;
            ResolveDepedencies();
            GetQueryStrings();
            bCreateGroupOrProgramFromOffer = MyCommon.Fetch_CM_SystemOption(134) == "1" ? true : false;
            bEnableRestrictedAccessToUEOfferBuilder = MyCommon.Fetch_SystemOption(249) == "1" ? true : false;
            AssignPageTitle("term.offer", "term.eligibilitycustomercondition", OfferID.ToString());
            if (!IsPostBack)
            {



                SetUpAndLocalizePage();
                CustomerGroupConditions objCustomerGroupConditions = m_CustCondition.GetOfferCustomerCondition(OfferID, EngineID);
                if (objCustomerGroupConditions != null)
                {
                    IncludedConditionGroup = (from p in objCustomerGroupConditions.IncludeCondition
                                              where p.Deleted == false & p.CustomerGroupID != 0
                                              select p.CustomerGroup).ToList();
                    ExcludedConditionGroup = (from p in objCustomerGroupConditions.ExcludeCondition
                                              where p.Deleted == false & p.CustomerGroupID != 0
                                              select p.CustomerGroup).ToList();
                }
                GetOfferEligibleCustomerCondition();
                GetAllCustomerGroup();

                IncludedGroup = (from p in OfferEligibileCustomerGroupCondition.IncludeCondition
                                 where p.Deleted == false
                                 select p.CustomerGroup).ToList();
                List<CustomerGroup> IncludedGroupsWithPhrase = IncludedGroup.Where(p => p.PhraseID != null).ToList();
                foreach (CustomerGroup cgroup in IncludedGroupsWithPhrase)
                {
                    cgroup.Name = PhraseLib.Lookup((Int32)cgroup.PhraseID, LanguageID).Replace("&#39;", "'");
                }

                ExcludedGroup = (from p in OfferEligibileCustomerGroupCondition.ExcludeCondition
                                 where p.Deleted == false
                                 select p.CustomerGroup).ToList();

                SetAvailableData();
                chkDisallow_Edit.Checked = OfferEligibileCustomerGroupCondition.DisallowEdit;
                chkHouseHold.Checked = OfferEligibileCustomerGroupCondition.HouseHoldEnabled;
                chkOffline.Checked = OfferEligibileCustomerGroupCondition.EvaluateOfflineCustomer;
                SetButtons();
                DisableControls();
            }
            else
            {
                GetValuesFromHidden();
                ScriptManager.RegisterStartupScript(this, this.GetType(), "selectAndFocus", " SetFoucs();", true);

            }




        }
        catch (Exception ex)
        {
            infobar.InnerText = ErrorHandler.ProcessError(ex);
            infobar.Visible = true;
        }

    }
    protected override void AuthorisePage()
    {
        if (CurrentUser.UserPermissions.AccessOffers == false && !IsTemplate)
        {
            Server.Transfer("PopUpDenied.aspx?PhraseName=perm.offers-access", false);
            return;
        }
        if (CurrentUser.UserPermissions.AccessTemplates == false && IsTemplate)
        {
            Server.Transfer("PopUpDenied.aspx?PhraseName=perm.offers-access-templates", false);
            return;
        }

    }

    protected void select1_Click(object sender, EventArgs e)
    {
        if (lstAvailable.SelectedItem != null)
        {

            foreach (int i in lstAvailable.GetSelectedIndices())
            {

                IncludedGroup.Add(AvailableFilteredCustomerGroup[i]);
            }
            HandleSelectedForSpecialGroup();
            SetAvailableData();

        }
        SetButtons();
    }

    protected void ReloadThePanel_Click(object sender, EventArgs e)
    {
        SetAvailableData();
    }

    protected void functionradio_CheckedChanged(object sender, EventArgs e)
    {
        SetAvailableData();
    }

    protected void deselect1_Click(object sender, EventArgs e)
    {

        if (lstSelected.SelectedItem != null)
        {
            if (ExcludedGroup.Count > 0 && lstSelected.GetSelectedIndices().Count() == IncludedGroup.Count())
            {
                infobar.InnerText = PhraseLib.Lookup("term-ValidationOnDeleteForAllSelected", LanguageID).Replace("&#39;", "'");
                infobar.Visible = true;
                return;
            }

            //need to reverse the order - find the issue in case of removing the items from index 0 
            var desc = from j in lstSelected.GetSelectedIndices().ToList()
                       orderby j descending
                       select j;

            foreach (int i in desc)
            {
                IncludedGroup.RemoveAt(i);
            }
            SetAvailableData();
        }
        SetButtons();
    }

    protected void select2_Click(object sender, EventArgs e)
    {
        if (lstAvailable.SelectedItem != null)
        {

            foreach (int i in lstAvailable.GetSelectedIndices())
            {
                if (AvailableFilteredCustomerGroup[i].CustomerGroupID == SystemCacheData.GetAnyCardHolderGroup().CustomerGroupID || AvailableFilteredCustomerGroup[i].CustomerGroupID == SystemCacheData.GetAnyCustomerGroup().CustomerGroupID || AvailableFilteredCustomerGroup[i].CustomerGroupID == SystemCacheData.GetNewCardHolderGroup().CustomerGroupID)
                {
                    infobar.InnerText = AvailableFilteredCustomerGroup[i].Name + " " + PhraseLib.Lookup("offer-eligibility-validateexlgroup", LanguageID).Replace("&#39;", "'");
                    infobar.Visible = true;
                    break;
                }
                ExcludedGroup.Add(AvailableFilteredCustomerGroup[i]);
            }

            SetAvailableData();

        }
        SetButtons();
    }

    protected void deselect2_Click(object sender, EventArgs e)
    {
        if (lstExcluded.SelectedItem != null)
        {
            //need to reverse the order - find the issue in case of removing the items from index 0 
            var desc = from j in lstExcluded.GetSelectedIndices().ToList()
                       orderby j descending
                       select j;
            foreach (int i in desc)
            {

                ExcludedGroup.RemoveAt(i);
            }
            SetAvailableData();

        }
        SetButtons();
    }

    protected void btnSave_Click(object sender, EventArgs e)
    {
        if (!(lstSelected.Items.Count > 0))
        {
            infobar.InnerText = PhraseLib.Lookup("term-validateincludedgroupset", LanguageID).Replace("&#39;", "'");
            infobar.Visible = true;
            return;
        }
        if (OfferEligibileCustomerGroupCondition.ConditionID == 0)
        {
            if (IsDefaultGroupNameExsits())
            {
                infobar.Visible = true;
                infobar.InnerText = String.Format(String.Format(PhraseLib.Lookup("OfferEligibilityCustomerCondition.validatedefaultgroupname", LanguageID), Constants.DEFAULT_OFFER_GROUP_NAME), hdnOfferName.Value).Replace("&#39;", "'");
                return;
            }
        }

        var deletedexclist = OfferEligibileCustomerGroupCondition.ExcludeCondition.Where(p => !ExcludedGroup.Any(exc => exc.CustomerGroupID == p.CustomerGroupID));
        if (OfferEligibileCustomerGroupCondition.ConditionID != 0)
        {
            //if it is an existing eligibility condition and user attempt to remove group from excluded condition which is exist in regular excluded condition then ask user to delete remove the same from regualr condition as well
            var mustExcludedList = deletedexclist.Where(p => ExcludedConditionGroup.Any(exc => exc.CustomerGroupID == p.CustomerGroupID)).Select(z => z.CustomerGroup);
            string strGroups = String.Empty;
            foreach (CustomerGroup item in mustExcludedList)
            {
                if (strGroups != "") { strGroups = strGroups + ","; }
                strGroups = strGroups + item.Name;
            }
            if (strGroups != String.Empty)
            {
                ScriptManager.RegisterStartupScript(this, this.GetType(), "RegulerConditionDelete", "ConfirmRegulerConditionDelete('" + String.Format(PhraseLib.Lookup("offer-eligibility-deconfirmation", LanguageID), strGroups) + "');", true);
            }
            else
            {
                Save();
            }
        }
        else
        {
            Save();
        }
    }

    protected void lst_DataBinding(object sender, EventArgs e)
    {
        ListBox lst = (ListBox)sender;
        int Counter = 0;
        for (Counter = 0; Counter < lst.Items.Count; Counter++)
        {
            ListItem lstItem = lst.Items[Counter];
            if (Counter > 4)
                break;
            if (lstItem.Value.ConvertToLong() == SystemCacheData.GetAnyCustomerGroup().CustomerGroupID ||
              lstItem.Value.ConvertToLong() == SystemCacheData.GetAllCAMCardHolderGroup().CustomerGroupID ||
              lstItem.Value.ConvertToLong() == SystemCacheData.GetNewCardHolderGroup().CustomerGroupID ||
              lstItem.Value.ConvertToLong() == SystemCacheData.GetAnyCardHolderGroup().CustomerGroupID)
                lstItem.Attributes.Add("style", "color:brown;font-weight:bold;");


        }

    }

    protected void btnDummySave_Click(object sender, EventArgs e)
    {
        Save();
    }

    protected void btnCreate_Click(object sender, EventArgs e)
    {
        string Name = string.Empty;
        if (MyCommon.Parse_Quotes(Logix.TrimAll(functioninput.Text)) != null)
            Name = Convert.ToString(MyCommon.Parse_Quotes(Logix.TrimAll(functioninput.Text)));
        if (!String.IsNullOrEmpty(Name))
        {
            int AvailableListCount = AvailableFilteredCustomerGroup.Where(p => p.Name.Equals(functioninput.Text, StringComparison.OrdinalIgnoreCase)).ToList().Count;
            int IncludedGroupCount = IncludedGroup.Where(p => p.Name.Equals(functioninput.Text, StringComparison.OrdinalIgnoreCase)).ToList().Count;
            int ExcludedGroupCount = ExcludedGroup.Where(p => p.Name.Equals(functioninput.Text, StringComparison.OrdinalIgnoreCase)).ToList().Count;

            bool isValidText = true;
            if (functioninput.Text.ToLower().Equals(PhraseLib.Lookup("term.anycardholder", LanguageID).ToLower()) || functioninput.Text.ToLower().Equals(PhraseLib.Lookup("term.anycustomer", LanguageID).ToLower()) || functioninput.Text.ToLower().Equals(PhraseLib.Lookup("term.newcardholders", LanguageID).ToLower()))
                isValidText = false;
            if (!isValidText)
            {
                string alertMessage = Copient.PhraseLib.Lookup("term.enter", LanguageID) + " " + Copient.PhraseLib.Lookup("term.valid", LanguageID).ToLower() + " " + Copient.PhraseLib.Lookup("term.customergroup", LanguageID).ToLower() + " " + Copient.PhraseLib.Lookup("term.name", LanguageID).ToLower();
                ScriptManager.RegisterStartupScript(UpdatePanelMain, UpdatePanelMain.GetType(), "AlertMessage", " AlertMessage('" + alertMessage + "'); ", true);
            }
            else if (IncludedGroupCount > 0 || ExcludedGroupCount > 0)
            {
                string alertMessage = Copient.PhraseLib.Lookup("term.customergroup", LanguageID) + ": " + Name + " " + Copient.PhraseLib.Lookup("term.is", LanguageID).ToLower() + " " + Copient.PhraseLib.Lookup("term.already", LanguageID).ToLower() + " " + Copient.PhraseLib.Lookup("term.selected", LanguageID).ToLower();
                ScriptManager.RegisterStartupScript(UpdatePanelMain, UpdatePanelMain.GetType(), "AlertMessage", " AlertMessage('" + alertMessage + "');", true);
            }
            else if (AvailableListCount > 0)
            {
                string alertMessage = Copient.PhraseLib.Lookup("term.existing", LanguageID) + " " + Copient.PhraseLib.Lookup("term.customergroup", LanguageID).ToLower() + ": " + Name + " " + Copient.PhraseLib.Lookup("offer.message", LanguageID).ToLower();
                ScriptManager.RegisterStartupScript(UpdatePanelMain, UpdatePanelMain.GetType(), "AlertMessage", " AlertMessage('" + alertMessage + "');", true);

                IncludedGroup.Add(AvailableFilteredCustomerGroup.Where(p => p.Name.Equals(functioninput.Text, StringComparison.OrdinalIgnoreCase)).FirstOrDefault());
                HandleSelectedForSpecialGroup();
                SetAvailableData();
                SetButtons();
            }
            else
            {
                IncludedGroup.Add(CreateCustomerGroup());
                SetAvailableData();
                SetButtons();
            }
        }
        else
        {
            string alertMessage = Copient.PhraseLib.Lookup("term.enter", LanguageID) + " " + Copient.PhraseLib.Lookup("term.valid", LanguageID).ToLower() + " " + Copient.PhraseLib.Lookup("term.customergroup", LanguageID).ToLower() + " " + Copient.PhraseLib.Lookup("term.name", LanguageID).ToLower();
            ScriptManager.RegisterStartupScript(UpdatePanelMain, UpdatePanelMain.GetType(), "AlertMessage", " AlertMessage('" + alertMessage + "');", true);
        }
    }

    #endregion Protected Methods

    #region Private Methods

    private void Save()
    {
        bool isNewCondition = false;
        string strIncludedGroup = string.Empty;
        string strExcludedGroup = string.Empty;
        try
        {


            if (OfferEligibileCustomerGroupCondition.ConditionID == 0)
            {
                isNewCondition = true;

            }
            if (chkDisallow_Edit.Visible)
                OfferEligibileCustomerGroupCondition.DisallowEdit = chkDisallow_Edit.Checked;

            if (spnHouseHold.Visible)
                OfferEligibileCustomerGroupCondition.HouseHoldEnabled = chkHouseHold.Checked;

            if (spnOffline.Visible)
                OfferEligibileCustomerGroupCondition.EvaluateOfflineCustomer = chkOffline.Checked;

            //Updated Include List
            var deletedlist = OfferEligibileCustomerGroupCondition.IncludeCondition.Where(p => !IncludedGroup.Any(inc => inc.CustomerGroupID == p.CustomerGroupID));
            foreach (CMS.AMS.Models.CustomerConditionDetails custdetail in deletedlist)
            {
                custdetail.Deleted = true;
            }
            historyString = PhraseLib.Lookup("history.con-customer-edit", LanguageID) + ": ";
            foreach (CMS.AMS.Models.CustomerGroup CustGroup in IncludedGroup)
            {

                if (!OfferEligibileCustomerGroupCondition.IncludeCondition.Exists(p => p.CustomerGroupID == CustGroup.CustomerGroupID))
                {
                    CMS.AMS.Models.CustomerConditionDetails condetail = new CMS.AMS.Models.CustomerConditionDetails();
                    condetail.CustomerGroupID = CustGroup.CustomerGroupID;
                    OfferEligibileCustomerGroupCondition.IncludeCondition.Add(condetail);
                    strIncludedGroup = strIncludedGroup + CustGroup.CustomerGroupID.ToString() + ",";
                }

            }
            historyString = historyString + strIncludedGroup.TrimEnd(',');
            //Update Exclude List exc

            var deletedexclist = OfferEligibileCustomerGroupCondition.ExcludeCondition.Where(p => !ExcludedGroup.Any(exc => exc.CustomerGroupID == p.CustomerGroupID));
            foreach (CMS.AMS.Models.CustomerConditionDetails custdetail in deletedexclist)
            {
                custdetail.Deleted = true;
            }
            bool IsExcludedExist = false;
            foreach (CMS.AMS.Models.CustomerGroup CustGroup in ExcludedGroup)
            {

                if (!OfferEligibileCustomerGroupCondition.ExcludeCondition.Exists(p => p.CustomerGroupID == CustGroup.CustomerGroupID))
                {
                    CMS.AMS.Models.CustomerConditionDetails condetail = new CMS.AMS.Models.CustomerConditionDetails();
                    condetail.CustomerGroupID = CustGroup.CustomerGroupID;
                    strExcludedGroup = strExcludedGroup + CustGroup.CustomerGroupID.ToString() + ",";
                    OfferEligibileCustomerGroupCondition.ExcludeCondition.Add(condetail);
                    IsExcludedExist = true;
                }
            }
            if (ExcludedConditionGroup != null)
            {
                //if it is a new condition then add excluded customer groups which are currently exist in regualr customer condition
                if (isNewCondition)
                {
                    foreach (CMS.AMS.Models.CustomerGroup CustGroup in ExcludedConditionGroup)
                    {

                        if (!OfferEligibileCustomerGroupCondition.ExcludeCondition.Exists(p => p.CustomerGroupID == CustGroup.CustomerGroupID))
                        {
                            CMS.AMS.Models.CustomerConditionDetails condetail = new CMS.AMS.Models.CustomerConditionDetails();
                            condetail.CustomerGroupID = CustGroup.CustomerGroupID;
                            strExcludedGroup = strExcludedGroup + CustGroup.CustomerGroupID.ToString() + ",";
                            OfferEligibileCustomerGroupCondition.ExcludeCondition.Add(condetail);
                            IsExcludedExist = true;
                        }
                    }
                }
                else
                {
                    //if it is an existing eligibility condition and user attempt to remove group from excluded condition which is exist in regular excluded condition then ask user to delete remove the same from regualr condition as well
                    var mustExcludedList = deletedexclist.Where(p => ExcludedConditionGroup.Any(exc => exc.CustomerGroupID == p.CustomerGroupID)).Select(z => z.CustomerGroup);
                    List<long> ExcludedGroupIds = new List<long>();
                    foreach (CustomerGroup item in mustExcludedList)
                    {
                        ExcludedGroupIds.Add(item.CustomerGroupID);

                    }
                    if (ExcludedGroupIds.Count > 0)
                    {
                        //Delete the excluded condition
                        m_CustCondition.DeleteRegulerExcludedConditionsByCustomerGroupIDs(OfferID, EngineID, ExcludedGroupIds);
                    }
                }
            }
            if (IsExcludedExist)
            {
                historyString = historyString + " " + PhraseLib.Lookup("term.excluding", LanguageID) + " " + strExcludedGroup.TrimEnd(',');
            }
            m_offer.CreateUpdateOfferEligibleCustomerCondition(OfferID, EngineID, OfferEligibileCustomerGroupCondition);
            if (isNewCondition)
            {
                CMS.AMS.Models.CustomerGroup CustomerGroup = new CMS.AMS.Models.CustomerGroup();
                CustomerGroup.Name = string.Format(Constants.DEFAULT_OFFER_GROUP_NAME, hdnOfferName.Value);
                CustomerGroup.IsOptinGroup = true;
                m_CustGroup.CreateOptInCustomerGroup(CustomerGroup);
                CMS.AMS.Models.CustomerGroupConditions CustomerCondition = new CMS.AMS.Models.CustomerGroupConditions();
                CustomerCondition.DisallowEdit = OfferEligibileCustomerGroupCondition.DisallowEdit;
                CustomerCondition.RequiredFromTemplate = OfferEligibileCustomerGroupCondition.RequiredFromTemplate;
                CustomerCondition.IncludeCondition = new List<CMS.AMS.Models.CustomerConditionDetails>();
                CustomerCondition.IncludeCondition.Add(new CMS.AMS.Models.CustomerConditionDetails());
                CustomerCondition.IncludeCondition[0].CustomerGroupID = CustomerGroup.CustomerGroupID;
                m_offer.CreateDefaultCustomerCondition(OfferID, EngineID, CustomerCondition);
            }
            m_offer.UpdateOfferStatusToModified(OfferID, EngineID, CurrentUser.AdminUser.ID);
            m_OAWService.ResetOfferApprovalStatus(OfferID);
            WriteToActivityLog();
            ScriptManager.RegisterStartupScript(this, this.GetType(), "Close", "CloseModel()", true);
        }
        catch (Exception ex)
        {
            infobar.InnerText = ErrorHandler.ProcessError(ex);
            infobar.Visible = true;
        }
    }

    private void DisableControls()
    {
        if (!IsTemplate)
            TempDisallow.Visible = false;
        else
            TempRequired.Visible = true;
        if (!IsTemplate)
        {
            DisabledAttribute = ((CurrentUser.UserPermissions.EditOffer && !(FromTemplate && OfferEligibileCustomerGroupCondition.DisallowEdit)) ? false : true);
        }
        else
            DisabledAttribute = CurrentUser.UserPermissions.EditTemplates ? false : true;

        //If disable is set to false, check Buyer conditions
        if (EngineID == 9 && !DisabledAttribute)
        {
            if (MyCommon.LRTadoConn.State == ConnectionState.Closed)
                MyCommon.Open_LogixRT();
            DisabledAttribute = ((CurrentUser.UserPermissions.EditOffersRegardlessBuyer || MyCommon.IsOfferCreatedWithUserAssociatedBuyer(CurrentUser.AdminUser.ID, OfferID)) ? false : true);
            MyCommon.Close_LogixRT();
        }
        if (DisabledAttribute)
        {
            functionradio1.Enabled = false;
            functionradio2.Enabled = false;
            functioninput.Enabled = false;
            chkOffline.Enabled = false;
            chkHouseHold.Enabled = false;
            lstAvailable.Enabled = false;
            lstSelected.Enabled = false;
            lstExcluded.Enabled = false;
            select1.Enabled = false;
            deselect1.Enabled = false;
            select2.Enabled = false;
            deselect2.Enabled = false;
            btnSave.Visible = false;
            btnCreate.Enabled = false;
        }

        if (!bCreateGroupOrProgramFromOffer || !CurrentUser.UserPermissions.CreateCustomerGroups || EngineID != 0)
        {
            btnCreate.Visible = false;
        }
        if ((bEnableRestrictedAccessToUEOfferBuilder && isTranslatedOffer) || (bEnableAdditionalLockoutRestrictionsOnOffers && !bOfferEditable) || m_OAWService.CheckIfOfferIsAwaitingApproval(OfferID).Result)
        {
            btnSave.Visible = false;
        }
    }

    private void WriteToActivityLog()
    {
        if (MyCommon.LRTadoConn.State == ConnectionState.Closed)
            MyCommon.Open_LogixRT();
        MyCommon.Activity_Log(3, OfferID, CurrentUser.AdminUser.ID, historyString);
        MyCommon.Close_LogixRT();
    }

    private bool IsDefaultGroupNameExsits()
    {


        if (m_CustGroup.GetCustomerGroupByName(string.Format(Constants.DEFAULT_OFFER_GROUP_NAME, hdnOfferName.Value)) != null)
            return true;
        else
            return false;

    }
    private void HandleSelectedForSpecialGroup()
    {
        var group = IncludedGroup.Where(p => p.CustomerGroupID == SystemCacheData.GetAnyCardHolderGroup().CustomerGroupID || p.CustomerGroupID == SystemCacheData.GetAnyCustomerGroup().CustomerGroupID).SingleOrDefault();
        if (group != null)
        {
            IncludedGroup.Clear();
            IncludedGroup.Add(group);
            select1.Enabled = false;
        }
        else
            select1.Enabled = true;
    }
    private void SetButtons()
    {
        if (lstSelected.Items.Count == 0)
        {
            select1.Enabled = true;
        }
        if (lstSelected.Items.Count > 0)
        {
            deselect1.Enabled = true;
            select2.Enabled = true;
        }
        else
        {
            deselect1.Enabled = false;
            select2.Enabled = false;
        }

        if (lstExcluded.Items.Count > 0)
        {
            deselect2.Enabled = true;
        }
        else
        {
            deselect2.Enabled = false;
        }

        if (lstSelected.Items.Count == 1 && (lstSelected.Items[0].Value.ConvertToLong() == SystemCacheData.GetAnyCardHolderGroup().CustomerGroupID || lstSelected.Items[0].Value.ConvertToLong() == SystemCacheData.GetAnyCustomerGroup().CustomerGroupID))
        {
            select1.Enabled = false;
        }

    }
    private void GetValuesFromHidden()
    {
        OfferID = hdnOfferID.Value.ConvertToLong();
        EngineID = hdnEngineID.Value.ConvertToInt32();
        ConditionID = hdnConditionID.Value.ConvertToLong();
        IsTemplate = hdnIsTemplate.Value.ConvertToBool();
        FromTemplate = hdnFromTemplate.Value.ConvertToBool();
        if (!IsTemplate)
            DisabledAttribute = ((CurrentUser.UserPermissions.EditOffer && !(FromTemplate && OfferEligibileCustomerGroupCondition.DisallowEdit)) ? false : true);
        else
            DisabledAttribute = CurrentUser.UserPermissions.EditTemplates ? false : true;

    }
    private void GetQueryStrings()
    {
        OfferID = Request["OfferID"].ConvertToLong();
        EngineID = Request["EngineID"].ConvertToInt32();
        ConditionID = Request["ConditionID"].ConvertToLong();
        IsTemplate = Request["IsTemplate"].ConvertToBool();
        FromTemplate = Request["FromTemplate"].ConvertToBool();
        isTranslatedOffer = MyCommon.IsTranslatedUEOffer(Convert.ToInt32(OfferID), MyCommon);

        //to do need to comment after integration
        bool _disallowEdit = false;
        if (OfferEligibileCustomerGroupCondition != null)
        {
            _disallowEdit = OfferEligibileCustomerGroupCondition.DisallowEdit;
        }
        if (!IsTemplate)
            DisabledAttribute = ((CurrentUser.UserPermissions.EditOffer && !(FromTemplate && _disallowEdit)) ? false : true);
        else
            DisabledAttribute = CurrentUser.UserPermissions.EditTemplates ? false : true;
        hdnConditionID.Value = ConditionID.ToString();
        hdnOfferID.Value = OfferID.ToString();
        hdnEngineID.Value = EngineID.ToString();
        hdnIsTemplate.Value = IsTemplate.ConvertToInt32().ToString();
        hdnFromTemplate.Value = FromTemplate.ConvertToInt32().ToString();
        hdnOfferName.Value = Request["OfferName"].ConvertToString();
        bOfferEditable = MyCommon.IsOfferEditablePastLockOutPeriod(CurrentUser.UserPermissions.EditOfferPastLockoutPeriod, MyCommon, Convert.ToInt32(OfferID));
        bEnableAdditionalLockoutRestrictionsOnOffers = MyCommon.Fetch_SystemOption(260) == "1" ? true : false;
    }
    private void ResolveDepedencies()
    {

        m_offer = CurrentRequest.Resolver.Resolve<CMS.AMS.Contract.IOffer>();
        m_CustGroup = CurrentRequest.Resolver.Resolve<CMS.AMS.Contract.ICustomerGroups>();
        m_CustCondition = CurrentRequest.Resolver.Resolve<CMS.AMS.Contract.ICustomerGroupCondition>();
        m_OAWService = CurrentRequest.Resolver.Resolve<IOfferApprovalWorkflowService>();
    }
    private void SetUpAndLocalizePage()
    {
        if (IsTemplate)
            title.InnerText = PhraseLib.Lookup("term.template", LanguageID) + " #" + OfferID + " " + PhraseLib.Lookup("term.eligibilitycustomercondition", LanguageID).Replace("&#39;", "'");
        else
            title.InnerText = PhraseLib.Lookup("term.offer", LanguageID) + " #" + OfferID + " " + PhraseLib.Lookup("term.eligibilitycustomercondition", LanguageID).Replace("&#39;", "'");
        btnSave.Text = PhraseLib.Lookup("term.save", LanguageID);
        select1.Text = "▼" + PhraseLib.Lookup("term.select", LanguageID);
        select2.Text = "▼" + PhraseLib.Lookup("term.select", LanguageID);
        deselect1.Text = PhraseLib.Lookup("term.deselect", LanguageID) + "▲";
        deselect2.Text = PhraseLib.Lookup("term.deselect", LanguageID) + "▲";
        btnCreate.Text = PhraseLib.Lookup("term.create", LanguageID);

        if (EngineID == 2)
            spnHouseHold.Visible = true;


        if (EngineID == 9)
            spnOffline.Visible = true;
    }

    private void SetAvailableData()
    {

        GetAllCustomerGroup();

        AvailableFilteredCustomerGroup = AllGroups.Where(p => !IncludedGroup.Any(inc => inc.CustomerGroupID == p.CustomerGroupID))
                                                            .Where(p => !ExcludedGroup.Any(exc => exc.CustomerGroupID == p.CustomerGroupID)).ToList();
        //Hide regular excluded customer groups for new condition
        if (ConditionID == 0 && ExcludedConditionGroup != null && ExcludedConditionGroup.Count > 0)
            AvailableFilteredCustomerGroup = AvailableFilteredCustomerGroup.Where(p => !ExcludedConditionGroup.Any(inc => inc.CustomerGroupID == p.CustomerGroupID)).ToList();


        string strFilter = functioninput.Text;

        if (functionradio1.Checked)
            AvailableFilteredCustomerGroup = AvailableFilteredCustomerGroup.Where(p => p.Name.StartsWith(strFilter, StringComparison.OrdinalIgnoreCase)).ToList();
        else
            AvailableFilteredCustomerGroup = AvailableFilteredCustomerGroup.Where(p => p.Name.IndexOf(strFilter, StringComparison.OrdinalIgnoreCase) >= 0).ToList();


        lstSelected.DataSource = IncludedGroup;
        lstSelected.DataBind();
        lstExcluded.DataSource = ExcludedGroup;
        lstExcluded.DataBind();

        lstAvailable.DataSource = AvailableFilteredCustomerGroup;
        lstAvailable.DataBind();

    }

    private void GetOfferEligibleCustomerCondition()
    {
        if (ConditionID == 0)
        {
            OfferEligibileCustomerGroupCondition = new CMS.AMS.Models.CustomerGroupConditions();
            OfferEligibileCustomerGroupCondition.IncludeCondition = new List<CMS.AMS.Models.CustomerConditionDetails>();
            OfferEligibileCustomerGroupCondition.ExcludeCondition = new List<CMS.AMS.Models.CustomerConditionDetails>();
            OfferEligibileCustomerGroupCondition.ConditionTypeID = m_CustCondition.GetCustomerGroupConditionTypeID(EngineID);
            OfferEligibileCustomerGroupCondition.EngineID = EngineID;

        }
        else
        {
            OfferEligibileCustomerGroupCondition = m_CustCondition.GetConditionByID(ConditionID);
            if (OfferEligibileCustomerGroupCondition == null)
                throw new Exception("Invalid ConditionID");
        }


    }
    private void GetAllCustomerGroup()
    {
        bool IsAnyCustomerEnabled = false;
        //if (EngineID == 2)
        //{
        // // IsAnyCustomerEnabled = (SystemCacheData.GetSystemOption_CPE_ByOptionId(125) == "1" ? true : false);
        //  if (IsAnyCustomerEnabled)
        //    IsAnyCustomerEnabled = m_offer.IsAnyCustomerAllowedForOffer(OfferID);
        //}

        AllGroups = m_CustGroup.GetCustomerGroups();

        if (!IsAnyCustomerEnabled)
        {
            var CustGroup = SystemCacheData.GetAnyCustomerGroup();
            var selectedgroup = AllGroups.Where(p => p.CustomerGroupID == CustGroup.CustomerGroupID).SingleOrDefault();
            AllGroups.Remove(selectedgroup);
        }
        //if CAM is Not installed, remove CAM specific groups - refer to AMS-14578
        if (EngineID != 6)
        {
            var camgroup = AllGroups.Where(p => p.CustomerGroupID == SystemCacheData.GetAllCAMCardHolderGroup().CustomerGroupID).SingleOrDefault();
            AllGroups.Remove(camgroup);

        }
        else
        {
            //do CMS engine specific things
        }
        List<CustomerGroup> lstCustomerGroupwithPhrases = AllGroups.Where(p => p.PhraseID != null).ToList();
        foreach (CustomerGroup cgroup in lstCustomerGroupwithPhrases)
        {
            cgroup.Name = PhraseLib.Lookup((Int32)cgroup.PhraseID, LanguageID).Replace("&#39;", "'");
        }
    }

    private CMS.AMS.Models.CustomerGroup CreateCustomerGroup()
    {
        CMS.AMS.Models.CustomerGroup NewCustGroup = null;

        try
        {
            bool saved = m_CustGroup.CreateUpdateCustomerGroup(new CustomerGroup { Name = Logix.TrimAll(functioninput.Text) });

            if (saved)
            {
                NewCustGroup = new CustomerGroup();
                NewCustGroup = m_CustGroup.GetCustomerGroupByName(Logix.TrimAll(functioninput.Text));
                if (MyCommon.LRTadoConn.State == ConnectionState.Closed)
                    MyCommon.Open_LogixRT();
                MyCommon.Activity_Log(7, NewCustGroup.CustomerGroupID, CurrentUser.AdminUser.ID, Copient.PhraseLib.Lookup("history.cgroup-create", LanguageID));
            }
        }
        catch (Exception err)
        {
            ScriptManager.RegisterStartupScript(UpdatePanelMain, UpdatePanelMain.GetType(), "AlertMessage", " AlertMessage('" + err.Message + "');", true);

        }
        finally
        {
            MyCommon.Close_LogixRT();
        }
        return NewCustGroup;
    }
    #endregion Private Methods
}