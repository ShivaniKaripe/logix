using System;
using System.Collections.Generic;
using System.Web.UI;
using CMS.AMS.Contract;
using CMS.AMS.Models;
using CMS.AMS;
using CMS;
using CMS.Contract;
using System.Web.UI.WebControls;
using System.Linq;
using System.Data;

public partial class logix_OfferApproval : AuthenticatedUI
{

    #region Private Variables
    private Copient.CommonInc mCommon;
    private ILogger m_logger;
    private IOfferApprovalWorkflowService m_OAWServices;
    private bool isOAWChecked;
    private bool isOAWEnabled;
    #endregion 

    #region properties
    protected bool isBannerEnabled
    {
        get { return (bool)ViewState["isBannerEnabled"]; }
        set { ViewState["isBannerEnabled"] = value; }
    }
    private List<OfferApprover> AllOfferApprovers
    {
        get { return ViewState["AllOfferApprovers"] as List<OfferApprover>; }
        set { ViewState["AllOfferApprovers"] = value;}
    }
    private List<OfferApprover> AllUsersWithDeploymentPermission
    {
        get { return (List<OfferApprover>)ViewState["AllUsersWithDeploymentPermission"]; }
        set { ViewState["AllUsersWithDeploymentPermission"] = value; }
    }
    private List<OfferApprover> AvailableFilteredOfferApprovers
    {
        get { return (List<OfferApprover>)ViewState["AvailableFilteredOfferApprovers"]; }
        set { ViewState["AvailableFilteredOfferApprovers"] = value; }
    }
    private List<OfferApprover> IncludedOfferApprovers
    {
        get { return (List<OfferApprover>)ViewState["IncludedOfferApprovers"]; }
        set { ViewState["IncludedOfferApprovers"] = value; }
    }
    #endregion 

    #region Protected Methods
    protected void Page_Load(object sender, EventArgs e)
    {
        AssignPageTitle("term.offerapprovalworkflow");
        ResolveDependencies();
        if (!Page.IsPostBack)
        {
            (this.Master as logix_LogixMasterPage).Tab_Name = "8_4";
            isBannerEnabled = mCommon.Fetch_SystemOption(66).Equals("1") ? true : false;
            ChangeBarDisplay();
            FillPageControlTextAndData();
        }
    }
    protected void Bannerddl_SelectedIndexChanged(object sender, EventArgs e)
    {
        try
        {
            ChangeBarDisplay();
            SetUserData();
            EnableDisableUserSettings();
        }
        catch (Exception ex)
        {
            m_logger.WriteCritical(ex.ToString());
            DisplayError(PhraseLib.Lookup("error.offerapproval-update", LanguageID));
        }

    }
    protected void BtnSave_Click(object sender, EventArgs e)
    {
        try
        {
            int defaultApproverId = GetDefaultApprover();
            bool isOAWEnabled = defaultApproverId > 0 ? true : false;
            bool isdefaultApproverChanged = (defaultapproverddl.SelectedValue.ConvertToInt32() != defaultApproverId);
            if ((enableapproval.Checked && !isOAWEnabled) || (isdefaultApproverChanged && isOAWEnabled) || (isOAWEnabled && !enableapproval.Checked))
                UpdateDefaultApprover();
            EnableDisableUserSettings();
        }
        catch(Exception ex)
        {
            m_logger.WriteCritical(ex.ToString());
            DisplayError(PhraseLib.Lookup("error.offerapproval-update", LanguageID));
        }
        
    }
    protected void Enableapproval_CheckedChanged(object sender, EventArgs e)
    {
        ChangeBarDisplay();
        EnableDisableUserSettings();
        if (CurrentUser.UserPermissions.ApprovalManager)
        {
            bool isOAWEnabled = GetDefaultApprover() > 0 ? true : false;
            if (enableapproval.Checked && !isOAWEnabled)
                DisplayInfoMsg(PhraseLib.Lookup("info.enableOfferApproval", LanguageID));
        }
    }
    protected void LstAvailable_SelectedIndexChanged(object sender, EventArgs e)
    {
        try
        {
            ChangeBarDisplay();
            ChangeControlsDisplay();
            SetApproverData();
        }
        catch(Exception ex)
        {
            m_logger.WriteCritical(ex.ToString());
            DisplayError(PhraseLib.Lookup("error.offerapproval-update", LanguageID));
        }
    }
    protected void Btnselect_Click(object sender, EventArgs e)
    {
        try
        {
            ChangeControlsDisplay();
            if (approverlistbox.SelectedValue.ConvertToInt32() > 0)
            {
                string selectedIds = "";
                ChangeBarDisplay();
                int[] approvers = approverlistbox.GetSelectedIndices();
                foreach(var approver in approvers)
                {
                    if (selectedIds == "") selectedIds = approverlistbox.Items[approver].Value;
                    else selectedIds += ", " + approverlistbox.Items[approver].Value;
                }
                UpdateUserApprovers(selectedIds);
                SetApproverData();
            }
        }
        catch (Exception ex)
        {
            m_logger.WriteCritical(ex.ToString());
            DisplayError(PhraseLib.Lookup("error.offerapproval-update", LanguageID));
        }
    }
    protected void Btndeselect_Click(object sender, EventArgs e)
    {
        try
        {
            ChangeControlsDisplay();
            if (selectedapproverlistbox.SelectedValue.ConvertToInt32() > 0)
            {
                string selectedIds = "";
                ChangeBarDisplay();
                int[] approvers = selectedapproverlistbox.GetSelectedIndices();
                foreach (var approver in approvers)
                {
                    if (selectedIds == "") selectedIds = selectedapproverlistbox.Items[approver].Value;
                    else selectedIds += ", " + selectedapproverlistbox.Items[approver].Value;
                }
                RemoveUserApprovers(selectedIds);
                SetApproverData();
            }
        }
        catch (Exception ex)
        {
            m_logger.WriteCritical(ex.ToString());
            DisplayError(PhraseLib.Lookup("error.offerapproval-update", LanguageID));
        }
    }
    protected void radioreqapproval_CheckedChanged(object sender, EventArgs e)
    {
        btnselect.Enabled = true;
        btndeselect.Enabled = true;
        ChangeBarDisplay();
        ChangeControlsDisplay();
        GetApproversForAUser();
        if(IncludedOfferApprovers != null && IncludedOfferApprovers.Count > 0 && IncludedOfferApprovers[0].AdminUserID == 0)
        {
            RemoveUserApprovers();
        }
    }
    protected void radiodeploy_CheckedChanged(object sender, EventArgs e)
    {
        ChangeBarDisplay();
        ChangeControlsDisplay();
        btnselect.Enabled = false;
        btndeselect.Enabled = false;
        GetApproversForAUser();
        if(!(IncludedOfferApprovers == null && IncludedOfferApprovers.Count > 0 && IncludedOfferApprovers[0].AdminUserID == 0))
        {
            RemoveUserApprovers();
            UpdateUserApprovers("0");
            SetApproverData();
        }
    }
    protected override void AuthorisePage()
    {
        if (CurrentUser.UserPermissions.ApprovalManager == false)
        {
            btnSave.Visible = false;
            btnselect.Enabled = false;
            btndeselect.Enabled = false;
            radiodeploy.Enabled = false;
            radioreqapproval.Enabled = false;
            defaultapproverddl.Enabled = false;
            enableapproval.Enabled = false;
            selectedapproverlistbox.Enabled = false;
            approverlistbox.Enabled = false;
        }
    }
    #endregion

    #region Private Methods
    private void EnableDisableUserSettings()
    {
        bool enable = true;
        int defaultApproverId = GetDefaultApprover();
        isOAWEnabled = (defaultApproverId > 0) ? true : false;
        if (!isOAWEnabled)
        {
            enable = false;
        }
        btnselect.Enabled = (enable) ? true : false;
        btndeselect.Enabled = (enable) ? true : false;
        radiodeploy.Enabled = (enable) ? true : false;
        radioreqapproval.Enabled = (enable) ? true : false;
        selectedapproverlistbox.Enabled = (enable) ? true : false;
        approverlistbox.Enabled = (enable) ? true : false;
        ChangeControlsDisplay();
    }
    private void ChangeControlsDisplay()
    {
        ScriptManager.RegisterStartupScript(this, this.GetType(), "ChangeDisplay", "ChangeControlsDisplay()", true);
    }
    private void ResolveDependencies()
    {
        mCommon = new Copient.CommonInc();
        m_OAWServices = CurrentRequest.Resolver.Resolve<IOfferApprovalWorkflowService>();
        m_logger = CurrentRequest.Resolver.Resolve<ILogger>();
    }
    private void FillPageControlTextAndData()
    {
        try
        {
            htitle.InnerText = PhraseLib.Lookup("term.offerapprovalworkflow", LanguageID);
            lbldeploy.Text = PhraseLib.Lookup("deployoffers.withoutapproval", LanguageID);
            lblreqapproval.Text = PhraseLib.Lookup("deployoffers.withapproval", LanguageID) + ": ";
            btnselect.Text = "▼" + " " + PhraseLib.Lookup("term.select", LanguageID);
            btndeselect.Text = "▲" + " " + PhraseLib.Lookup("term.deselect", LanguageID);
            btnSave.Text = PhraseLib.Lookup("term.save", LanguageID);
            if (isBannerEnabled)
            {
                lblbanner.Visible = true;
                bannerddl.Visible = true;
                lblbanner.Text = PhraseLib.Lookup("term.banner", LanguageID);
                lbldefaultapprover.Text = PhraseLib.Lookup("term.defaultapprover-banner", LanguageID);
                lblenableapproval.Text = PhraseLib.Lookup("term.enableworkflow-banner", LanguageID);
            }
            else
            {
                lbldefaultapprover.Text = PhraseLib.Lookup("term.defaultapprover", LanguageID);
                lblenableapproval.Text = PhraseLib.Lookup("term.enableworkflow", LanguageID);
            }
            SetAvailableData();
        }
        catch (Exception ex)
        {
            m_logger.WriteCritical(ex.ToString());
            DisplayError(PhraseLib.Lookup("error.offerapproval-update", LanguageID));
        }
    }
    private void SetAvailableData()
    {
            if (isBannerEnabled)
            {
                AMSResult<DataTable> dtBanners = m_OAWServices.GetAllBanners();
                if(dtBanners.ResultType == AMSResultType.Success)
                {
                    if(dtBanners.Result != null)
                    {
                        bannerddl.DataSource = dtBanners.Result;
                        bannerddl.DataBind();
                        bannerddl.SelectedValue = "1";
                    }
                }
                else
                {
                    DisplayError(PhraseLib.Lookup(" error.fetching-banners", LanguageID));
                }
                
            }
            SetUserData();
    }
    private void SetUserData()
    {
        GetAllOfferApprovers();
        GetUsersWithDeploymentPermission();
        if (AllUsersWithDeploymentPermission != null)
        {
            lstAvailable.DataSource = AllUsersWithDeploymentPermission;
            lstAvailable.DataBind();
            if (lstAvailable.Items.Count > 0)
            {
                lstAvailable.SelectedIndex = 0;
            }
        }
        if (AllOfferApprovers != null)
        {
            defaultapproverddl.DataSource = AllOfferApprovers;
            defaultapproverddl.DataBind();
            SetApproverData();
        }
        SetSavedData();
    }
    private void UpdateDefaultApprover()
    {
        int selecteddefaultApproverId;
        bool approverExists = false;
        AMSResult<bool> recordUpdated = new AMSResult<bool>();
        isOAWChecked = enableapproval.Checked;
        selecteddefaultApproverId = defaultapproverddl.SelectedValue.ConvertToInt32();
        if (selecteddefaultApproverId > 0)
        {
            int defaultApproverId = GetDefaultApprover();
            if (defaultApproverId > 0)
            {
                approverExists = true;
            }
            if (defaultApproverId != selecteddefaultApproverId && isOAWChecked)
            {
                recordUpdated = m_OAWServices.InsertUpdateDefaultApprover(selecteddefaultApproverId, ((isBannerEnabled) ? bannerddl.SelectedValue.ConvertToInt32() : 0), approverExists);
            }
            else if (!isOAWChecked && defaultApproverId != 0)
            {
                recordUpdated = m_OAWServices.RemoveDefaultApprover(defaultApproverId, ((isBannerEnabled) ? bannerddl.SelectedValue.ConvertToInt32() : 0));
            }
            if (recordUpdated.ResultType != AMSResultType.Success)
            {
                DisplayError(PhraseLib.Lookup("error.offerapproval-update", LanguageID));
            }
            else
            {
                SetSavedData();
                DisplaySuccessMsg(PhraseLib.Lookup("info.OAW-updated", LanguageID));
            }
        }
        else
        {
            DisplayError(PhraseLib.Lookup("error.nodefaultapprover", LanguageID));
        }
        
    }
    private void UpdateUserApprovers(string approvers)
    {
        int adminUserId = lstAvailable.SelectedValue.ConvertToInt32();

        AMSResult<bool> recordUpdated = m_OAWServices.InsertUpdateApprovers(adminUserId, approvers, (isBannerEnabled) ? bannerddl.SelectedValue.ConvertToInt32() : 0);
        if(recordUpdated.ResultType == AMSResultType.Success)
        {
            DisplaySuccessMsg(PhraseLib.Lookup("info.approverselection-updated", LanguageID));
        }
        else
        {
            DisplayError(PhraseLib.Lookup("error.offerapproval-update", LanguageID));
        }
    }
    private void RemoveUserApprovers(string removeApprovers = "")
    {
        int adminUserId = lstAvailable.SelectedValue.ConvertToInt32();
        int bannerId = (isBannerEnabled) ? bannerddl.SelectedValue.ConvertToInt32() : 0;
        AMSResult<bool> recordUpdated = m_OAWServices.RemoveApproverForAUser(removeApprovers, bannerId, adminUserId);
        if (recordUpdated.ResultType == AMSResultType.Success)
        {
            DisplaySuccessMsg(PhraseLib.Lookup("info.approverselection-updated", LanguageID));
        }
        else
        {
            DisplayError(PhraseLib.Lookup("error.offerapproval-update", LanguageID));
        }
    }
    private void SetSavedData()
    {
        int defaultApproverId = GetDefaultApprover();
        if (defaultApproverId > 0)
        {
            if(defaultapproverddl.Items.Count > 0)
                defaultapproverddl.SelectedValue = defaultApproverId.ToString();
            enableapproval.Checked = true;
        }
        else
        {
            enableapproval.Checked = false;
            if (defaultapproverddl.Items.Count > 0)
                defaultapproverddl.SelectedIndex = 0;
        }
    }
    private void SetApproverData()
    {
        OfferApprover availableOfferApprover = new OfferApprover();
        GetApproversForAUser();
        if (IncludedOfferApprovers != null && IncludedOfferApprovers.Count > 0)
        {
            if(IncludedOfferApprovers.Count == 1 && IncludedOfferApprovers[0].AdminUserID == 0)
            {
                radiodeploy.Checked = true;
                radioreqapproval.Checked = false;
                approverlistbox.DataSource = AllOfferApprovers;
                approverlistbox.DataBind();
                selectedapproverlistbox.Items.Clear();
                btndeselect.Enabled = false;
                btnselect.Enabled = false;
            }
            else
            {
                radioreqapproval.Checked = true;
                radiodeploy.Checked = false;
                btndeselect.Enabled = true;
                btnselect.Enabled = true;
                selectedapproverlistbox.DataSource = IncludedOfferApprovers;
                selectedapproverlistbox.DataBind();
                AvailableFilteredOfferApprovers = AllOfferApprovers.Where(p => !IncludedOfferApprovers.Any(p2 => p2.AdminUserID == p.AdminUserID)).ToList();
                if (AvailableFilteredOfferApprovers != null && AvailableFilteredOfferApprovers.Count > 0)
                {
                    approverlistbox.DataSource = AvailableFilteredOfferApprovers;
                    approverlistbox.DataBind();
                }
                else
                {
                    approverlistbox.Items.Clear();
                }
            }
        }
        else
        {
            radioreqapproval.Checked = true;
            btndeselect.Enabled = true;
            btnselect.Enabled = true;
            approverlistbox.DataSource = AllOfferApprovers;
            approverlistbox.DataBind();
            selectedapproverlistbox.Items.Clear();
        }
    }
    private int GetDefaultApprover()
    {
        int defaultApproverId = 0;
        AMSResult<int> defaultApprover = m_OAWServices.GetDefaultApprover((isBannerEnabled) ? bannerddl.SelectedValue.ConvertToInt32() : 0);
        if (defaultApprover.ResultType == AMSResultType.Success)
        {
            if(defaultApprover.Result > 0)
                defaultApproverId = defaultApprover.Result;
        }
        else
        {
            DisplayError(PhraseLib.Lookup("error.fetching-defaultapprover", LanguageID));
        }
        return defaultApproverId;
    }
    private void GetApproversForAUser()
    {
        AMSResult<List<OfferApprover>> offerApprovers = m_OAWServices.GetOfferApprovers(lstAvailable.SelectedValue.ConvertToInt32(), (isBannerEnabled) ? bannerddl.SelectedValue.ConvertToInt32() : 0);
        if (offerApprovers.ResultType == AMSResultType.Success)
        {
            if(offerApprovers.Result != null)
                IncludedOfferApprovers = offerApprovers.Result;
        }
        else
        {
            DisplayError(PhraseLib.Lookup("error.fetching-approvers", LanguageID));
        }
    }
    private void GetAllOfferApprovers()
    {
        AMSResult<List<OfferApprover>> result = m_OAWServices.GetUsersWithOfferApprovalPermission((isBannerEnabled) ? bannerddl.SelectedValue.ConvertToInt32() : 0);
        if(result.ResultType == AMSResultType.Success)
        {
            if (result.Result != null)
                AllOfferApprovers = result.Result;
        }
        else
        {
            DisplayError(PhraseLib.Lookup("error.fetching-users", LanguageID));
        }
    }
    private void GetUsersWithDeploymentPermission()
    {
        AMSResult<List<OfferApprover>> result = m_OAWServices.GetUsersWithDeploymentPermission((isBannerEnabled) ? bannerddl.SelectedValue.ConvertToInt32() : 0);
        if (result.ResultType == AMSResultType.Success)
        {
           
                AllUsersWithDeploymentPermission = result.Result;
        }
        else
        {
            DisplayError(PhraseLib.Lookup("error.fetching-users", LanguageID));
        }
    }
    private void DisplayError(String errorText)
    {
        infobar.Attributes["class"] = "red-background";
        infobar.InnerText = errorText;
        infobar.Style["display"] = "block";
    }
    private void DisplayInfoMsg(String message)
    {
        infobar.Attributes["class"] = "orange-background";
        infobar.InnerText = message;
        infobar.Style["display"] = "block";
    }
    private void DisplaySuccessMsg(String message)
    {
        infobar.Attributes["class"] = "green-background";
        infobar.InnerText = message;
        infobar.Style["display"] = "block";
    }
    private void ChangeBarDisplay()
    {
        if (infobar.Style["display"] == "block") infobar.Style["display"] = "none";
    }
    #endregion










  
}