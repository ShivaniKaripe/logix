using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using CMS.AMS;
using CMS.AMS.Models;
using CMS.AMS.Contract;
using CMS;
using CMS.DB;
using System.Text;
using System.Collections;
using System.Data;
using Copient;
using System.Data.SqlClient;
public partial class logix_UE_UEoffer_rew_tc : AuthenticatedUI
{
    #region Variables

    protected int OfferID = 0;
    int TCRewardID = 0;
    int NumTiers;
    int DeliverableID;
    int RewardID = 0;
    int DefaultLanguageID = 0;
    bool mliOverriddenInTemplate;
    Hashtable OverrideDiv = null;
    Control resultControl = null;
    Control ct = null;
    CommonInc m_CommonInc;
    bool DisabledAttribute = false;
    string OvrdFldClass;
    ITrackableCouponProgramService m_TCProgram;
    IOfferApprovalWorkflowService m_OAWService;
    ICouponRewardService m_CRService;
    IActivityLogService m_ActivityLog;
    CouponReward couponReward;
    IOffer m_Offer;
    CMS.AMS.Common m_Commondata;
    List<CouponTierTranslation> tierTransData = new List<CouponTierTranslation>();
    AMSResult<TrackableCouponProgram> savedProgram = new AMSResult<TrackableCouponProgram>();
    #endregion

    #region properties

    protected CMS.AMS.Models.Offer Offer
    {
        get { return ViewState["Offer"] as CMS.AMS.Models.Offer; }
        set { ViewState["Offer"] = value; }
    }
    private List<TrackableCouponProgram> AvailableFilteredTCProgram
    {

        get
        {

            return ViewState["AvailableFilteredTCProgram"] as List<TrackableCouponProgram>;
        }
        set
        {
            ViewState["AvailableFilteredTCProgram"] = value;
        }
    }

    private List<TrackableCouponProgram> IncludedTCProgram
    {

        get
        {
            return ViewState["IncludedTCProgram"] as List<TrackableCouponProgram>;
        }
        set
        {
            ViewState["IncludedTCProgram"] = value;
        }

    }
    private TrackableCouponProgram SavedTCProgram
    {

        get
        {
            return ViewState["SavedTCProgram"] as TrackableCouponProgram;
        }
        set
        {
            ViewState["SavedTCProgram"] = value;
        }

    }
    private List<CMS.AMS.Models.Language> lstLanguage
    {
        get { return ViewState["Language"] as List<CMS.AMS.Models.Language>; }
        set { ViewState["Language"] = value; }
    }
    private bool isMultiLanguageEnabled
    {
        get { return (bool)ViewState["isMLEnabled"]; }
        set { ViewState["isMLEnabled"] = value; }
    }
    #endregion

    #region Protected Methods

    
    protected void Page_Load(object sender, EventArgs e)
    {
        ResolveDependencies();
        GetQueryString();
        LoadOfferSettings();
        
        isMultiLanguageEnabled = SystemSettings.IsMultiLanguageEnabled();
        DefaultLanguageID = SystemSettings.GetSystemDefaultLanguage().LanguageID;
        AssignPageTitle("term.offer", "term.trackablecoupon", OfferID.ToString());
        GetCRewardObj();
        SetDisabledAttribute();
        if (!IsPostBack)
        {
            SetControlText();
            SetTierValues();
            if (TCRewardID > 0)
                SetAvailableData(true);
            else
                SetAvailableData(false);
        }
        
      
        SetUpAndLocalizePage();
        DisableControls();
      
    }
    protected void ReloadThePanel_Click(object sender, EventArgs e)
    {
        SetAvailableData(false);
    }
    protected void btnSave_Click(object sender, EventArgs e)
    {
        try
        {
            bool bSave = true;
            int progCount = 0;
            foreach (RepeaterItem rep in repTier_selectedTCP.Items)
            {
                if (((ListBox)rep.FindControl("lstSelected")).Items.Count == 1)
                    progCount++;
            }
            if (progCount == 0)
                bSave = false;
           
           if(!bSave)
                    DisplayError(PhraseLib.Lookup("CPE-rew-membership.selectpoints", LanguageID));
           else
            SaveCouponReward();
        }
        catch (Exception ex)
        {
            DisplayError(ex);
        }
    }
    protected void ddlprinttype_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (ddlprinttype.SelectedValue.ConvertToInt32() != 1)
        {
            ddlsubtype.Visible = true;
            lblsubtype.Visible = true;

        }
        else
        {
            lblsubtype.Visible = false;
            ddlsubtype.Visible = false;
        }
    }

    protected void repTier_selectedTCP_ItemDataBound(object sender, RepeaterItemEventArgs e)
    {
        if (e.Item.ItemType == ListItemType.Item || e.Item.ItemType == ListItemType.AlternatingItem)
        {
            Int32 TierLevel = e.Item.ItemIndex;
            ((Button)e.Item.FindControl("btnselect")).Text = "▼" + " " + PhraseLib.Lookup("term.select", LanguageID);
            ((Button)e.Item.FindControl("btndeselect")).Text = "▲" + " " + PhraseLib.Lookup("term.deselect", LanguageID);
            List<TrackableCouponProgram> selectedList = new List<TrackableCouponProgram>();
            if (TCRewardID > 0)
            {
                GetAllTCProgram();
                selectedList = AvailableFilteredTCProgram.ToList();
                    selectedList = selectedList.Where(p => p.ProgramID == ((CouponTier)e.Item.DataItem).ProgramID).ToList();
            }
            
           
            ((ListBox)e.Item.FindControl("lstSelected")).DataSource = selectedList;
            ((ListBox)e.Item.FindControl("lstSelected")).DataBind();

        }
    }
    protected void repTier_selectedTCP_ItemCommand(object source, RepeaterCommandEventArgs e)
    {
        
        TrackableCouponProgram progSelected = new TrackableCouponProgram();
        if (IncludedTCProgram == null)
            IncludedTCProgram = new List<TrackableCouponProgram>();
        
        if (e.CommandName.ToString() == "Select")
        {
            if (lstAvailable.SelectedValue == "") return;
            progSelected = AvailableFilteredTCProgram.Where(p => p.ProgramID == lstAvailable.SelectedValue.ConvertToInt32()).SingleOrDefault();
            IncludedTCProgram.Add(progSelected);
        }
        else if (e.CommandName.ToString() == "Deselect")
        {
            if (((ListBox)e.Item.FindControl("lstSelected")).SelectedValue == "") return;
            if (NumTiers == 1)
            {
                progSelected = IncludedTCProgram.Where(p => p.ProgramID == ((ListBox)e.Item.FindControl("lstSelected")).SelectedValue.ConvertToInt32()).SingleOrDefault();
                if (progSelected == null) return;
                IncludedTCProgram = IncludedTCProgram.Where(p => p.ProgramID != progSelected.ProgramID).ToList();
            }
               
            
            progSelected = null;
        }
        
        BindDataToListBox(e, progSelected);
        DisableControls();
    }
    
    
    #endregion
   
    #region Private Methods

    public void SetDisabledAttribute()
    {
        if (Offer.FromTemplate)
            DisabledAttribute = ((CurrentUser.UserPermissions.EditOffer && !couponReward.DisallowEdit) ? false : true);
        else if (Offer.IsTemplate)
            DisabledAttribute = CurrentUser.UserPermissions.EditTemplates ? false : true;
        else
            DisabledAttribute = CurrentUser.UserPermissions.EditOffer ? false : true;

        //If disable is set to false, check Buyer conditions
        if (!DisabledAttribute)
        {
            if (m_CommonInc.LRTadoConn.State == ConnectionState.Closed)
                m_CommonInc.Open_LogixRT();
            DisabledAttribute = ((CurrentUser.UserPermissions.EditOffersRegardlessBuyer || m_CommonInc.IsOfferCreatedWithUserAssociatedBuyer(CurrentUser.AdminUser.ID, OfferID)) ? false : true);
            DisabledAttribute = m_OAWService.CheckIfOfferIsAwaitingApproval(OfferID).Result;
            m_CommonInc.Close_LogixRT();
            //Hide save button if Disable is true
            btnSave.Visible = !DisabledAttribute;
        }
    }
    private void DisableControls()
    {
        foreach (RepeaterItem rep in repTier_selectedTCP.Items)
        {
            if (((ListBox)rep.FindControl("lstSelected")).Items.Count > 0)
            {
                ((Button)rep.FindControl("btnselect")).Enabled = false;
            }
            else
                ((Button)rep.FindControl("btnselect")).Enabled = true;
        }
       
    }
    private void BindDataToListBox(RepeaterCommandEventArgs e, TrackableCouponProgram progSelected)
    {
        List<TrackableCouponProgram> includedList = new List<TrackableCouponProgram>();
        if (progSelected != null)
            includedList.Add(progSelected);
        includedList = includedList.OrderBy(o => o.Name).ToList();
        ((ListBox)e.Item.FindControl("lstSelected")).DataSource = includedList;
        ((ListBox)e.Item.FindControl("lstSelected")).DataBind();

        SetAvailableData(false);
    }
    private void GetCRewardObj()
    {
        int i;
        couponReward = m_CRService.GetCouponRewardbyID(TCRewardID);

        if (couponReward.CouponTiers == null)
        {
            couponReward.CouponTiers = new List<CouponTier>();
            for (i = 0; i < NumTiers; i++)
            {
                CouponTier couponTier = new CouponTier();
                couponTier.TierLevel = i + 1;
                couponTier.Description = "";
                couponReward.CouponTiers.Add(couponTier);
                couponTier.CouponTierTranslations = new List<CouponTierTranslation>();
                couponTier.CouponTierTranslations.Add(new CouponTierTranslation() { CouponTierID = 0, DeliverableMessage = "", Id = 0, LanguageId = 0 });
            }
        }
        else if (couponReward.CouponTiers.Count < NumTiers)
        {
            for (i = couponReward.CouponTiers.Count; i < NumTiers; i++)
            {
                CouponTier couponTier = new CouponTier();
                couponTier.TierLevel = i + 1;
                couponTier.Description = "";
                couponReward.CouponTiers.Add(couponTier);
                couponTier.CouponTierTranslations = new List<CouponTierTranslation>();
                couponTier.CouponTierTranslations.Add(new CouponTierTranslation() { CouponTierID = 0, DeliverableMessage = "", Id = 0, LanguageId = 0 });
            }
        }
    }
    private void SetTierValues()
    {
        try
        {
            repTier_selectedTCP.DataSource = couponReward.CouponTiers;
            repTier_selectedTCP.DataBind();
            repTier_Desc.DataSource = couponReward.CouponTiers;
            repTier_Desc.DataBind();
            if(Offer.FromTemplate) DisableAttributes();
        }
        catch (Exception ex)
        {
            DisplayError(ex.ToString());
        }
    }

    private void DisableAttributes()
    {
        foreach (RepeaterItem rep in repTier_selectedTCP.Items)
        {
            Button btnselect = (Button)rep.FindControl("btnselect");
            if (btnselect != null) btnselect.Enabled = !DisabledAttribute;

            Button btndeselect = (Button)rep.FindControl("btndeselect");
            if (btndeselect != null) btndeselect.Enabled = !DisabledAttribute;

            ListBox lb = (ListBox)rep.FindControl("lstSelected");
            if (lb != null) lb.Enabled = !DisabledAttribute;
        }
        foreach (RepeaterItem rep in repTier_Desc.Items)
        {
            TextBox txt = (TextBox)rep.FindControl("ucMLI$tbMLI");
            if (txt != null) txt.Enabled = !DisabledAttribute;
        }
        ddlprinttype.Enabled = !DisabledAttribute;
        ddlsubtype.Enabled = !DisabledAttribute;
        deliverytypes.Enabled = !DisabledAttribute;
        successful.Enabled = !DisabledAttribute;
        lstAvailable.Enabled = !DisabledAttribute;
        functioninput.Enabled = !DisabledAttribute;
        functionradio1.Enabled = !DisabledAttribute;
        functionradio2.Enabled = !DisabledAttribute;
    }
    private void DisplayError(String ErrorText)
    {
        infobar.InnerText = ErrorText;
        infobar.Visible = true;
    }

    private void DisplayError(Exception ex)
    {
        DisplayError(ErrorHandler.ProcessError(ex));
    }

    private Hashtable GetOverriddenFields()
    {
        Hashtable overrideFields = new Hashtable();
        //Get the Lockable fields data.
        DataTable dt = m_Commondata.GetFieldLevelPermissions(OfferID, AppName);
        List<string> LockedTemplateFields = new List<string>();
        foreach (DataRow row in dt.Rows)
        {
            if (row["DeliverableID"].ConvertToInt32() == DeliverableID)
            {
                if (row["Tiered"].ConvertToBool() == true)
                {
                    for (int i = 0, j = 0; i < Offer.NumbersOfTier; i++)
                    {
                        overrideFields.Add(repTier_selectedTCP.ID + "$ctl0" + j + "$" + row["ControlName"].ConvertToString(), row["Editable"].ConvertToBool());
                        j += 2;
                    }
                    for (int i = 0, j = 0; i < Offer.NumbersOfTier; i++)
                    {
                        overrideFields.Add(repTier_Desc.ID + "$ctl0" + j + "$" + row["ControlName"].ConvertToString(), row["Editable"].ConvertToBool());
                        j += 2;
                    }
                }
                else
                {
                    overrideFields.Add(row["ControlName"].ConvertToString(), row["Editable"].ConvertToBool());
                }
                LockedTemplateFields.Add(row["FieldID"].ConvertToString());
                if (row["ControlName"].ConvertToString().Contains("ucMLI$tbMLI"))
                    mliOverriddenInTemplate = true;
            }
        }
        hdnLockedTemplateFields.Value = string.Join(",", LockedTemplateFields);

        return overrideFields;
    }

    private void SetCouponRewardData()
    {
        couponReward = new CouponReward();
        couponReward.ROID = RewardID;
        couponReward.RewardOptionPhase = 3;
        couponReward.Required = successful.Checked;
        couponReward.DisallowEdit = chkDisallow_Edit.Checked;
    }
    private List<CouponTier> SetCouponTierData()
    {
        List<CouponTier> CouponTierList = new List<CouponTier>();
       
        int repeaterItemCount = repTier_selectedTCP.Items.Count;
        for(int i = 0; i < repeaterItemCount; i++)
        {
            ListBox lb = ((ListBox)repTier_selectedTCP.Items[i].FindControl("lstSelected"));
            CouponTier CouponTierData = new CouponTier();
            if(lb.Items.Count != 0)
                CouponTierData.ProgramID = lb.Items[0].Value.ConvertToInt32();
            CouponTierData.TierLevel = i + 1;
            if (ddlsubtype.Visible == true)
            {
                CouponTierData.PrintTypeID = ddlprinttype.SelectedValue.ConvertToInt32();
                CouponTierData.PrintSubTypeID = ddlsubtype.SelectedValue.ConvertToInt32();
            }
            else
                CouponTierData.PrintTypeID = ddlprinttype.SelectedValue.ConvertToInt32();
            CouponTierData.TCDeliveryTypeID = deliverytypes.SelectedValue.ConvertToInt32();
            CouponTierData.CouponTierTranslations = SetCouponTranslationData(i, CouponTierData);
            CouponTierList.Add(CouponTierData);
        }
        return CouponTierList;
    }
    private void SaveCouponReward()
    {
        AMSResult<bool> success;
        string historyString;

        SetCouponRewardData();
        List<CouponTier> CouponTierList = SetCouponTierData();
        couponReward.CouponTiers = CouponTierList;

        if (TCRewardID > 0)
        {
            historyString = PhraseLib.Lookup("term.edited", LanguageID) + " " + PhraseLib.Lookup("term.trackablecoupon", LanguageID).ToLower() + " " + PhraseLib.Lookup("term.reward", LanguageID).ToLower();
            couponReward.CouponRewardID = TCRewardID;
            couponReward.DeliverableID = DeliverableID;
           success = m_CRService.UpdateCouponReward(OfferID, LanguageID, couponReward);
        }
        else
        {
            historyString = PhraseLib.Lookup("term.created", LanguageID) + " " + PhraseLib.Lookup("term.trackablecoupon", LanguageID).ToLower() + " " + PhraseLib.Lookup("term.reward", LanguageID).ToLower(); 
                success = m_CRService.CreateCouponReward(OfferID, LanguageID, couponReward);
        }
         
        if (success.ResultType == AMSResultType.Success)
        {
            if (Offer.IsTemplate)
            {
                //time to update the status bits for the templates
                int form_Disallow_Edit = 0;
                string[] LockFieldsList = null;

                form_Disallow_Edit = chkDisallow_Edit.Checked == true ? 1 : 0;
                if (!string.IsNullOrEmpty(hdnLockedTemplateFields.Value))
                {
                    LockFieldsList = hdnLockedTemplateFields.Value.Split(',');
                    m_Commondata.PurgeFieldLevelPermissions(AppName, Offer.OfferID, 0);
                    m_Commondata.UpdateDeliverableAndFieldLevelPermissions(couponReward.DeliverableID.ConvertToInt32(), Offer.OfferID, form_Disallow_Edit, LockFieldsList);
                }
                else
                {
                    m_Commondata.PurgeFieldLevelPermissions(AppName, Offer.OfferID, 0);
                }
            }
            m_OAWService.ResetOfferApprovalStatus(OfferID);
            m_ActivityLog.Activity_Log(ActivityTypes.Offer, OfferID, CurrentUser.AdminUser.ID, historyString);
            ScriptManager.RegisterStartupScript(this, this.GetType(), "Close", "CloseModel()", true);
        }
        else if (success.ResultType == AMSResultType.SQLException || success.ResultType == AMSResultType.Exception)
        {
            DisplayError(success.MessageString);
        }

    }
    private List<CouponTierTranslation> SetCouponTranslationData(int tierLevel,CouponTier ctier)
    {
        List<CouponTierTranslation> CouponTierTrans=null;
        if (repTier_Desc != null)
        {
            Repeater repeater = (Repeater)repTier_Desc;
             FindNestedControl(repeater.Items[tierLevel], "tbMLI");
            if(resultControl !=null)
            {
                CouponTierTrans = new List<CouponTierTranslation>();
                TextBox txtDesc = (TextBox)resultControl;
                ctier.Description = txtDesc.Text.Trim();
                CouponTierTranslation ctr = new CouponTierTranslation()
                {
                    CouponTierID = ctier.Id,
                    DeliverableMessage = txtDesc.Text.Trim(),
                    LanguageId = DefaultLanguageID
                };
                CouponTierTrans.Add(ctr);
                    resultControl=null;
            }
            if(isMultiLanguageEnabled)
            {
                FindNestedControl(repeater.Items[tierLevel], "repMLIInputs");
                if (resultControl != null)
                {
                    Repeater repMLI = (Repeater)resultControl;
                    CouponTierTrans = SetTierTranslations(repMLI, 1, ctier.Description);
                    resultControl = null;
                }
            }
            
          
        }

        return CouponTierTrans;
    }
    private void ResolveDependencies()
    {
        CurrentRequest.Resolver.AppName = "UEoffer-rew-tc.aspx";
        m_TCProgram = CurrentRequest.Resolver.Resolve<ITrackableCouponProgramService>();
        m_OAWService = CurrentRequest.Resolver.Resolve<IOfferApprovalWorkflowService>();
        m_CRService = CurrentRequest.Resolver.Resolve<ICouponRewardService>();
        m_ActivityLog = CurrentRequest.Resolver.Resolve<IActivityLogService>();
        m_Offer = CurrentRequest.Resolver.Resolve<IOffer>();
        m_Commondata = CurrentRequest.Resolver.Resolve<CMS.AMS.Common>();
        m_CommonInc = CurrentRequest.Resolver.Resolve<CommonInc>();
    }

    private void LoadOfferSettings()
    {
        Offer = m_Offer.GetOffer(OfferID, LoadOfferOptions.None);
        // IsTemplate = Offer.IsTemplate;
        //  FromTemplate = Offer.FromTemplate;
        NumTiers = Offer.NumbersOfTier;
    }
    private void GetQueryString()
    {
        OfferID = Request.QueryString["OfferID"].ConvertToInt32();
        RewardID = Request.QueryString["RewardID"].ConvertToInt32();
        TCRewardID = !(Request.QueryString["TCDeliverableID"].ConvertToInt32() == null) ? Request.QueryString["TCDeliverableID"].ConvertToInt32() : 0;
        DeliverableID = Request.QueryString["DeliverableID"].ConvertToInt32();
    }
    private StringBuilder GetTemplateScript(Hashtable overrideFields, Hashtable overrideDiv, string editable)
    {
        StringBuilder script = new StringBuilder("xmlhttpPost(\"UEtemplateFeeds.aspx?OfferID=" + Offer.OfferID + "&PageName=" + AppName + "&DeliverableID=" + DeliverableID + "&PageEditable=" + editable + "\");");

        //Update the locked fields.
        if (overrideFields.Count > 0)
        {
            foreach (DictionaryEntry de in overrideFields)
            {
                OvrdFldClass = de.Value.ConvertToString().ToUpper() == "TRUE" ? "enabledTemplateField" : "disabledTemplateField";
                if (de.Key.ToString().Contains("btnselect"))
                {
                    foreach (RepeaterItem rep in repTier_selectedTCP.Items)
                    {
                        ListBox lb = (ListBox)rep.FindControl("lstSelected");
                        if (lb != null) lb.CssClass = OvrdFldClass;
                    }
                }
                else if (de.Key.ToString().Contains("ucMLI"))
                {
                    foreach (RepeaterItem rep in repTier_Desc.Items)
                    {
                        TextBox txt = (TextBox)rep.FindControl("ucMLI$tbMLI");
                        if (txt != null) txt.CssClass = OvrdFldClass;
                    }
                }
                else if (de.Key.ToString().Contains("ddlprinttype"))
                {
                    ddlprinttype.CssClass = OvrdFldClass;
                    ddlsubtype.CssClass = OvrdFldClass;
                }
                else if (de.Key.ToString().Contains("successful"))
                    successful.CssClass = OvrdFldClass;
                else if (de.Key.ToString().Contains("deliverytypes"))
                    deliverytypes.CssClass = OvrdFldClass;
            }
        }
        return script;
    }
    private void RegisterScript(string key, string script)
    {
        // Get a ClientScriptManager reference from the Page class.
        ClientScriptManager cs = Page.ClientScript;

        // Check to see if the startup script is already registered.
        if (!cs.IsStartupScriptRegistered(this.GetType(), key))
        {
            cs.RegisterStartupScript(this.GetType(), key, script, true);
        }
    }

    private void SetControlText()
    {
        title.InnerText = PhraseLib.Lookup("term.offer", LanguageID) + " #" + OfferID + " " + PhraseLib.Lookup("term.trackablecoupon", LanguageID) + " " + PhraseLib.Lookup("term.reward", LanguageID).ToLower();
        btnSave.Text = PhraseLib.Lookup("term.save", LanguageID);
        lblsucdelivery.Text = PhraseLib.Lookup("ue-reward.reward-required", LanguageID);
        lblprinttype.Text = PhraseLib.Lookup("cpesettingsvalue.55", LanguageID) + " " + PhraseLib.Lookup("term.type", LanguageID) + " : ";
        functioninput.Text = "";
        lblsubtype.Text = PhraseLib.Lookup("term.barcode", LanguageID) + " " + PhraseLib.Lookup("term.type", LanguageID) + " : ";
        try
        {
                ddlprinttype.DataSource = m_CRService.GetPrintTypes(LanguageID);
                ddlprinttype.DataBind();
                ddlsubtype.DataSource = m_CRService.GetPrintSubTypes();
                ddlsubtype.DataBind();

                deliverytypes.DataSource = m_CRService.GetTCDeliveryTypes(LanguageID);
                deliverytypes.DataBind();
                deliverytypes.SelectedIndex = 0;
        }
        catch (Exception ex)
        {
            DisplayError(ex);
        }
    }
    private void SetUpAndLocalizePage()
    {
        int overriddenFieldsCount = 0;
        Hashtable overriddenFields = null;
        overriddenFields = new Hashtable();
        OverrideDiv = new Hashtable();
        if (Offer.IsTemplate && !IsPostBack)
        {
            string editable = m_Commondata.IsOfferEditable(DeliverableID);
             overriddenFields = GetOverriddenFields();
            overriddenFieldsCount = overriddenFields.Count;
            RegisterScript("xmlhttpPost", GetTemplateScript(overriddenFields, OverrideDiv, editable).ToString());
        }
        else if(Offer.FromTemplate)
        {
            overriddenFields = GetOverriddenFields();
            overriddenFieldsCount = overriddenFields.Count;
            RegisterScript("xmlhttpPost", GetOfferFromTemplateScript(overriddenFields).ToString());
            SetMLIPopupProperty();
            TempDisallow.Visible = false;
        }
        else if(!Offer.IsTemplate)
        {
            SetMLIPopupProperty();
            TempDisallow.Visible = false;
        }
        
    }

    private void SetMLIPopupProperty()
    {
        //Set the popup enable\disable
        //mliOverriddenInTemplate is true when the buydescription control for multi\single language is overridden
        if (isMultiLanguageEnabled
            &&
            (Offer.FromTemplate && (couponReward.DisallowEdit && !mliOverriddenInTemplate) || (!couponReward.DisallowEdit && mliOverriddenInTemplate))
            ||
            (!Offer.FromTemplate && !Offer.IsTemplate && DisabledAttribute))  //DisabledAttribute false means edit permissions are there for offer
        {
            foreach (RepeaterItem item in repTier_Desc.Items)
            {
                if (item.FindControl("ucMLI") != null)
                {
                    logix_UserControls_MultiLanguagePopup popup = (logix_UserControls_MultiLanguagePopup)item.FindControl("ucMLI");
                    popup.DisablePopup = true;
                }
            }
        }
    }

    private StringBuilder GetOfferFromTemplateScript(Hashtable overrideFields)
    {
        StringBuilder script = new StringBuilder();
        if (overrideFields.Count > 0)
        {
            string OvrdFldDisabled = "False";
            foreach (DictionaryEntry de in overrideFields)
            {
                OvrdFldDisabled = de.Value.ToString().ToUpper() == "TRUE" ? "true" : "false";
                if (de.Key.ToString().Contains("btnselect"))
                        DisablePrograms(OvrdFldDisabled);
                   
                    else if (de.Key.ToString().Contains("ucMLI"))
                    {
                        foreach (RepeaterItem rep in repTier_Desc.Items)
                        {
                            TextBox txt = (TextBox)rep.FindControl("ucMLI$tbMLI");
                            if (txt != null) txt.Enabled = OvrdFldDisabled.ConvertToBool();
                        }
                    }
                    else if (de.Key.ToString().Equals("ddlprinttype"))
                    {
                        ddlprinttype.Enabled = OvrdFldDisabled.ConvertToBool();
                        ddlsubtype.Enabled = OvrdFldDisabled.ConvertToBool();
                    }
                    else if(de.Key.ToString().Equals("deliverytypes"))
                        deliverytypes.Enabled = OvrdFldDisabled.ConvertToBool();
                    
                    else if (de.Key.ToString().Equals("successful"))
                        successful.Enabled = OvrdFldDisabled.ConvertToBool();
            }
        }
        else if (DisabledAttribute)
        {
            btnSave.Visible = false;
        }
        return script;
    }

    private void DisablePrograms(string OvrdFldDisabled)
    {
        foreach (RepeaterItem rep in repTier_selectedTCP.Items)
        {
            Button btnselect = (Button)rep.FindControl("btnselect");
            if (btnselect != null) btnselect.Enabled = OvrdFldDisabled.ConvertToBool();

            Button btndeselect = (Button)rep.FindControl("btndeselect");
            if (btndeselect != null) btndeselect.Enabled = OvrdFldDisabled.ConvertToBool();

            ListBox lb = (ListBox)rep.FindControl("lstSelected");
            if (lb != null) lb.Enabled = OvrdFldDisabled.ConvertToBool();
        }
        lstAvailable.Enabled = OvrdFldDisabled.ConvertToBool();
        functioninput.Enabled = OvrdFldDisabled.ConvertToBool();
        functionradio1.Enabled = OvrdFldDisabled.ConvertToBool();
        functionradio2.Enabled = OvrdFldDisabled.ConvertToBool();
    }
    private void SetAvailableData(bool LoadSavedData)
    {
        try
        {
            List<TrackableCouponProgram> filterlist = new List<TrackableCouponProgram>();
            TrackableCouponProgram includedprog = new TrackableCouponProgram();
            GetAllTCProgram();
            filterlist = AvailableFilteredTCProgram.ToList();
           
            if (LoadSavedData)
            {
                SetSavedOfferTCReward();
                IncludedTCProgram = AvailableFilteredTCProgram.Where(p => p.ProgramID == 0).ToList();
                foreach (CouponTier cTier in couponReward.CouponTiers)
                {
                    includedprog = AvailableFilteredTCProgram.Where(p => p.ProgramID == cTier.ProgramID).SingleOrDefault();
                    if(includedprog != null) IncludedTCProgram.Add(includedprog);
                }
                if (NumTiers == 1)
                {
                    foreach (CouponTier cTier in couponReward.CouponTiers)
                    {
                        filterlist = filterlist.Where(p => p.ProgramID != cTier.ProgramID).ToList();
                    }
                }
               
            }
            else
            {
                if (NumTiers == 1)
                {
                    if (IncludedTCProgram != null)
                        foreach (TrackableCouponProgram prog in IncludedTCProgram)
                        {
                            filterlist = filterlist.Where(p => p.ProgramID != prog.ProgramID).ToList();
                        }
                }
            }

            filterlist = filterlist.OrderBy(o => o.Name).ToList();
            lstAvailable.DataSource = filterlist;
            lstAvailable.DataBind();
        }
        catch (Exception ex)
        {
            DisplayError(ex);
        }

    }

    private void SetSavedOfferTCReward()
    {
        successful.Checked = couponReward.Required;
        ddlprinttype.SelectedValue = couponReward.CouponTiers[0].PrintTypeID.ToString();
        if (couponReward.CouponTiers[0].PrintSubTypeID == 0)
        {
            lblsubtype.Visible = false;
            ddlsubtype.Visible = false;
        }
        else
        {
            lblsubtype.Visible = true;
            ddlsubtype.Visible = true;
            ddlsubtype.SelectedValue = couponReward.CouponTiers[0].PrintSubTypeID.ToString();
        }
        deliverytypes.SelectedValue = couponReward.CouponTiers[0].TCDeliveryTypeID.ToString();
        chkDisallow_Edit.Checked = couponReward.DisallowEdit;
    }

    private string GetSearchText()
    {
        string SearchText = "";
        if (functioninput.Text.Equals("%"))
            SearchText = "[%]";
        else if (functioninput.Text.Equals("_"))
            SearchText = "[_]";
        else
            SearchText = functioninput.Text;
        return SearchText;
    }
    private void GetAllTCProgram()
    {
        int RecordCount = 100;
        string SearchText = GetSearchText();
        AMSResult<List<TrackableCouponProgram>> AllTCProgram = m_TCProgram.GetAllTrackableCouponPrograms(RecordCount, SearchText, functionradio1.Checked);
        if (AllTCProgram.ResultType != AMSResultType.Success)
        {
            AvailableFilteredTCProgram = new List<TrackableCouponProgram>();
            DisplayError(AllTCProgram.GetLocalizedMessage(LanguageID));
        }
        AvailableFilteredTCProgram = AllTCProgram.Result;
       
    }
    private void AddLanguageControl(RepeaterItemEventArgs e)
    {
        Control divControl = e.Item.FindControl("divBuyDescriptionMLI");
        if (divControl != null)
        {
            if (!isMultiLanguageEnabled)
            {
                logix_UserControls_MultiLanguagePopup ucMLI = (logix_UserControls_MultiLanguagePopup)LoadControl("~/logix/UserControls/MultiLanguagePopup.ascx");
                ucMLI.ID = "ucMLI";
                ucMLI.IsMultiLanguageEnabled = isMultiLanguageEnabled;
                divControl.Controls.Add(ucMLI);
            }
            else
            {
                logix_UserControls_MultiLanguagePopup ucMLI = (logix_UserControls_MultiLanguagePopup)LoadControl("~/logix/UserControls/MultiLanguagePopup.ascx");
                ucMLI.ID = "ucMLI";
                ucMLI.MLIdentifierValue = Convert.ToInt64(DataBinder.Eval(e.Item.DataItem, "id"));
                ucMLI.MLTableName = "DeliverableCouponTierTranslation";
                ucMLI.MLColumnName = "DeliverableMessage";
                ucMLI.MLITranslationColumn = "DeliverableCouponTierID";
                ucMLI.IsMultiLanguageEnabled = isMultiLanguageEnabled;
                //ucMLI.DisablePopup = disablePopup;
                if (DataBinder.Eval(e.Item.DataItem, "Description") != null)
                    ucMLI.MLDefaultLanguageStandardValue = DataBinder.Eval(e.Item.DataItem, "Description").ToString();
                divControl.Controls.Add(ucMLI);
            }
        }
    }
    private void FindNestedControl(Control control, string childControlId)
    {
        Control tempControl = control.FindControl(childControlId);

        if (tempControl == null)
        {
            foreach (Control c in control.Controls)
            {
                FindNestedControl(c, childControlId);
            }
        }
        else
            resultControl = tempControl;
    }

    private List<CouponTierTranslation> SetTierTranslations(Repeater repeater, int cTierId, string defaultDescription)
    {
        CouponTierTranslation ct = null;
        List<CouponTierTranslation> listct = new List<CouponTierTranslation>();

        foreach (RepeaterItem ri in repeater.Items)
        {

            HiddenField hfLangId = (HiddenField)ri.FindControl("hfLangId");
            int langId = Convert.ToInt32(hfLangId.Value);
            TextBox tbTranslation = (TextBox)ri.FindControl("tbTranslation");

            ct = new CouponTierTranslation();
            ct.LanguageId = langId;
            ct.CouponTierID = cTierId;
            if (langId == DefaultLanguageID)
                ct.DeliverableMessage = defaultDescription.Trim();
            else
                ct.DeliverableMessage = tbTranslation.Text.Trim();

            listct.Add(ct);
        }
        return listct;
    }

   
    #endregion




   
    protected void repTier_Desc_ItemDataBound(object sender, RepeaterItemEventArgs e)
    {
        if (!isMultiLanguageEnabled)
        {
            FindNestedControl(e.Item, "tbMLI");
            if (resultControl != null)
            {
                TextBox Desc = (TextBox)resultControl;
                if (e.Item.DataItem != null && DataBinder.Eval(e.Item.DataItem, "Description") !=null)
                    Desc.MaxLength = 1000;
                    Desc.Text = DataBinder.Eval(e.Item.DataItem, "Description").ToString();
            }
        }
        resultControl = null;
        Repeater repeater = (Repeater)sender;
        FindNestedControl(e.Item, "repMLIInputs");
        if (resultControl != null)
        {
            Repeater rep = (Repeater)resultControl;
            rep.ItemDataBound += new RepeaterItemEventHandler(repMLIInputs_ItemDataBound);

            resultControl = null;
        }
    }
    
    private void repMLIInputs_ItemDataBound(object sender, RepeaterItemEventArgs e)
    {
        //Get the translations from the multilanguage usercontrol
        if (e.Item.DataItem != null)
        {
            MultiLanguageInput mli = (MultiLanguageInput)e.Item.DataItem;
            CouponTierTranslation tierTrans = new CouponTierTranslation();
           
            tierTrans.LanguageId = mli.LanguageID;
            tierTrans.CouponTierID = mli.IdentifierValue.ConvertToInt32();
            tierTrans.DeliverableMessage = mli.Translation.Trim();
            tierTransData.Add(tierTrans);
            

        }
    }


    protected void repTier_Desc_ItemCreated(object sender, RepeaterItemEventArgs e)
    {
        AddLanguageControl(e);
    }
}