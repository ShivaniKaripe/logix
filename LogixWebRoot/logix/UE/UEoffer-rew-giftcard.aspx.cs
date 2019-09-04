using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.UI;
using System.Web.UI.WebControls;
using CMS;
using CMS.AMS;
using CMS.AMS.Contract;
using CMS.AMS.Models;
using System.Data;
using System.Collections;
using System.Text;
using Copient;

public partial class logix_UE_UEoffer_rew_giftcard : AuthenticatedUI
{
    # region Global vairables
    int gcrTierTemplateFieldCount;
    int gcrTemplateFieldCount;
    bool IsNewGiftCard = false;
    bool DisabledAttribute = false;
    bool ProductCondition = false;
    private bool percentOffAllowed = true;
    int DefaultLanguageID = 0;
    int OfferID;
    int DeliverableID;
    int RewardID;
    int Phase;
    string Description;
    string OvrdFldClass = "";
    IOffer m_Offer;
    IGiftCardRewardService m_GiftCardRewardService;
    CMS.AMS.Common m_Commondata;
    CommonInc m_CommonInc;
    Control resultControl = null;
    Hashtable OverrideDiv = new Hashtable();
    bool OvrdFldEditable = false;
    UnitTypeInformation currencyInformation;
    List<GiftCardTierTranslation> gcTierTranslations = new List<GiftCardTierTranslation>();
    ILocalizationService m_LocalizationService;
    IActivityLogService m_ActivityLogService;
    IOfferApprovalWorkflowService m_OAWService;
    private bool restrictProrationTypeToAllConditional;
    bool disablePopup;
    bool mliOverriddenInTemplate;
    # endregion Global vairables
    # region Properties
    public string ValueTypeInitialPrefix
    {
        get { return ViewState["ValueTypeInitialPrefix"].ToString(); }
        set { ViewState["ValueTypeInitialPrefix"] = value; }
    }
    public string ValueTypeInitialSuffix
    {
        get { return ViewState["ValueTypeInitialSuffix"].ToString(); }
        set { ViewState["ValueTypeInitialSuffix"] = value; }
    }
    public string CurrencySymbol
    {
        get { return ViewState["CurrencySymbol"].ToString(); }
        set { ViewState["CurrencySymbol"] = value; }
    }
    public string CurrencyAbbr
    {
        get { return ViewState["CurrencyAbbr"].ToString(); }
        set { ViewState["CurrencyAbbr"] = value; }
    }

    private bool isMultiLanguageEnabled
    {
        get { return (bool)ViewState["isMLEnabled"]; }
        set { ViewState["isMLEnabled"] = value; }
    }
    protected CMS.AMS.Models.Offer objOffer
    {
        get { return ViewState["Offer"] as CMS.AMS.Models.Offer; }
        set { ViewState["Offer"] = value; }
    }

    private GiftCard objGiftCard
    {
        get { return ViewState["objGiftCard"] as GiftCard; }
        set { ViewState["objGiftCard"] = value; }
    }

    private List<CMS.AMS.Models.Language> lstLanguage
    {
        get { return ViewState["Language"] as List<CMS.AMS.Models.Language>; }
        set { ViewState["Language"] = value; }
    }
    # endregion Properties

    # region Page events
    protected void ddlValueType_OnSelectedIndexChanged(object sender, EventArgs e)
    {
        DropDownList ddlValueType = (DropDownList)sender;
        //AL-5916 : restrictions to percent off and prorationtypes in GCR popup
        DataTable dt = SystemCacheData.LoadDefaultProrationsForGiftCard();
        if (restrictProrationTypeToAllConditional && ddlValueType.SelectedValue.ConvertToInt32() == (int)CPEAmountTypes.PercentageOff)
            dt.DefaultView.RowFilter = "ID=" + (int)UEProrationTypes.AllConidtionalItems;
        else
            dt.DefaultView.RowFilter = "";

        ddlProrationRate.DataSource = dt;
        ddlProrationRate.DataBind();
        foreach (RepeaterItem ri in repGiftcard.Items)
        {
            UpdateCurrencyControls(Convert.ToInt32(ddlValueType.SelectedValue), ri);
        }
    }
    protected void repGiftcard_ItemCreated(object sender, RepeaterItemEventArgs e)
    {
        AddLanguageControl(e);
    }
    protected void repGiftcard_ItemDataBound(object sender, RepeaterItemEventArgs e)
    {
        Repeater repeater = (Repeater)sender;
        FindNestedControl(e.Item, "repMLIInputs");
        if (resultControl != null)
        {
            Repeater rep = (Repeater)resultControl;
            rep.ItemDataBound += new RepeaterItemEventHandler(repMLIInputs_ItemDataBound);

            resultControl = null;
        }
    }
    void repMLIInputs_ItemDataBound(object sender, RepeaterItemEventArgs e)
    {
        //Get the translations from the multilanguage usercontrol
        if (e.Item.DataItem != null)
        {
            MultiLanguageInput mli = (MultiLanguageInput)e.Item.DataItem;

            GiftCardTierTranslation gct = new GiftCardTierTranslation();
            gct.LanguageId = mli.LanguageID;
            gct.GiftCardTierId = mli.IdentifierValue;
            gct.BuyDescription = mli.Translation;

            gcTierTranslations.Add(gct);
        }
    }
    protected override void OnInit(EventArgs e)
    {
        AppName = "UEoffer-rew-giftcard.aspx";
        base.OnInit(e);

    }

    protected void Page_Load(object sender, EventArgs e)
    {
        ResolveDependencies();
        GetQueryString();
        chkRollUp.Text = PhraseLib.Lookup("term.rollup", LanguageID);
        DefaultLanguageID = SystemSettings.GetSystemDefaultLanguage().LanguageID;
        isMultiLanguageEnabled = SystemSettings.IsMultiLanguageEnabled();
        if (objOffer == null)
        {
            if (OfferID == 0)
            {
                DisplayError(PhraseLib.Lookup("error.invalidoffer", LanguageID));
                return;
            }
            objOffer = m_Offer.GetOffer(OfferID, CMS.AMS.Models.LoadOfferOptions.CustomerCondition);
            lstLanguage = SystemSettings.GetAllActiveLanguages((Engines)objOffer.EngineID);
        }

        if (!IsPostBack)
        {

            if (DeliverableID != 0)
            {
                AMSResult<GiftCard> result = m_GiftCardRewardService.GetGiftCardReward(DeliverableID, objOffer.EngineID);
                if (result.ResultType == AMSResultType.Success)
                {
                    objGiftCard = result.Result;
                }
                else
                {
                    DisplayError(result.GetLocalizedMessage<GiftCard>(LanguageID));
                    return;
                }
            }
        }
        DisableControls();
        if (!IsPostBack)
            LoadPageData();
        SetUpAndLocalizePage();
    }

    protected void btnSave_Click(object sender, EventArgs e)
    {
        try
        {
            CreateObjectAndSave();
        }
        catch (Exception ex)
        {
            DisplayError(ex.Message);
        }
    }
    //Refresh the Parent UE page to display newly added/Updated Reward
    protected string GetRefreshScript()
    {
        return "opener.location='/logix/UE/UEoffer-rew.aspx?OfferID=" + objOffer.OfferID + "'; ";
    }
    # endregion Page events

    # region private methods
    private void SetControlsTexts()
    {
        AssignPageTitle("term.offer", "term.GiftCard", OfferID.ToString());
        btnSave.Text = PhraseLib.Lookup("term.save", LanguageID);
        lblValueType.Text = PhraseLib.Lookup("term.value", LanguageID) + PhraseLib.Lookup("term.type", LanguageID) + ":";

        if (objOffer.IsTemplate)
            title.InnerText = PhraseLib.Lookup("term.template", LanguageID) + " #" + OfferID + " " + PhraseLib.Lookup("term.GiftCard", LanguageID);
        else if (objOffer.FromTemplate)
            title.InnerText = PhraseLib.Lookup("term.offer", LanguageID) + " #" + OfferID + " " + PhraseLib.Lookup("term.GiftCard", LanguageID);
        else
            title.InnerText = PhraseLib.Lookup("term.offer", LanguageID) + " #" + OfferID + " " + PhraseLib.Lookup("term.GiftCard", LanguageID);
    }
    private void SetUpAndLocalizePage()
    {
        int overriddenFieldsCount = 0;
        Hashtable overriddenFields = null;
        SetControlsTexts();
        //If offer is template, register script to update the control styles on client side.
        if (objOffer.IsTemplate && !IsPostBack)
        {
            string editable = m_Commondata.IsOfferEditable(DeliverableID);
            overriddenFields = GetOverriddenFields();
            overriddenFieldsCount = overriddenFields.Count;
            RegisterScript("xmlhttpPost", GetTemplateScript(overriddenFields, OverrideDiv, editable).ToString());
        }
        //If the offer instantiated from Template, register script to disable the controls appropriately.
        else if (objOffer.FromTemplate)
        {
            overriddenFields = GetOverriddenFields();
            overriddenFieldsCount = overriddenFields.Count;
            RegisterScript("xmlhttpPost", GetOfferFromTemplateScript(overriddenFields).ToString());
            SetMLIPopupProperty();
            TempDisallow.Visible = false;
        }
        else
        {
            SetMLIPopupProperty();
            TempDisallow.Visible = false;
        }
        UpdateSaveButton(overriddenFieldsCount);
    }
    //Load data specific to this page
    private void LoadPageData()
    {
        SetGiftCardForLoad();
    }
    private void SetGiftCardForLoad()
    {
        //If tier level is increased for offer, update gift card with respective tier data.
        //If it is decreased we are updating gift card using SP in UEOffer-gen page
        if (objGiftCard != null)
        {
            int existingTiers = objGiftCard.GiftCardTiers.Count();
            int newTiers = objOffer.NumbersOfTier;
            if (newTiers > existingTiers)
            {
                for (int j = existingTiers + 1; j <= newTiers; j++)
                {
                    GiftCardTier gt = new GiftCardTier();
                    gt.TierLevel = j;
                    gt.GCTierTranslations = new List<GiftCardTierTranslation>();
                    objGiftCard.GiftCardTiers.Add(gt);
                }
            }
        }
        if (objGiftCard == null)
        {
            objGiftCard = new CMS.AMS.Models.GiftCard();
            objGiftCard.Id = 0;
            objGiftCard.RewardOptionId = RewardID;
            objGiftCard.RewardOptionPhase = Phase;
            objGiftCard.RewardTypeID = (int)DELIVERABLE_TYPES.GIFTCARD;
            objGiftCard.Required = true;
            objGiftCard.Rollup = false;
            objGiftCard.GiftCardTiers = new List<GiftCardTier>();
            GiftCardTier m_GiftCardTier;

            for (int i = 0; i < objOffer.NumbersOfTier; i++)
            {
                m_GiftCardTier = new GiftCardTier();
                m_GiftCardTier.TierLevel = i + 1;
                m_GiftCardTier.GCTierTranslations = new List<GiftCardTierTranslation>();
                objGiftCard.GiftCardTiers.Add(m_GiftCardTier);
            }
        }

        repGiftcard.DataSource = objGiftCard.GiftCardTiers;
        repGiftcard.DataBind();
        //Load default values
        LoadDefaultData();
        objGiftCard.GiftCardTiers.Clear();
    }
    private void LoadDefaultData()
    {
        int prorationTypeId = 0, chargeBackDeptId = 0, amountTypeId = 0;

        currencyInformation = m_LocalizationService.LoadCurrencyInfoForOffer(RewardID, CurrentUser.AdminUser.LanguageID);
        CurrencyAbbr = currencyInformation.Abbrevation;
        CurrencySymbol = currencyInformation.Symbol;

        chkRollUp.Checked = objGiftCard.Rollup;
        chkDisallow_Edit.Checked = objGiftCard.DisallowEdit;

        LoadValueTypeDropDown(ProductCondition);
        bool prodCondExists = DoesProductConditionExists();

        //AL-5916 : restrictions to percent off and proration types in GCR popup
        DataTable dt = SystemCacheData.LoadDefaultProrationsForGiftCard();
        if (restrictProrationTypeToAllConditional && amountTypeId == (int)CPEAmountTypes.PercentageOff)
            dt.DefaultView.RowFilter = "ID=" + (int)UEProrationTypes.AllConidtionalItems;
        else
            dt.DefaultView.RowFilter = "";
        LoadDropDown(ddlProrationRate, "PHRASE", "ID", dt);

        LoadDropDown(ddlChargeBack, "NAME", "CHARGEBACKDEPTID", m_Commondata.GetDefaultChargeBack(LanguageID, prodCondExists));

        foreach (RepeaterItem item in repGiftcard.Items)
        {
            SetBuyDescription(item);

            if (objGiftCard.GiftCardTiers.Count >= repGiftcard.Items.Count)
            {
                prorationTypeId = objGiftCard.GiftCardTiers[0].ProrationTypeId;
                if (prorationTypeId == 0)
                    prorationTypeId = ddlProrationRate.Items[0].Value.ConvertToInt32();
                ddlProrationRate.SelectedValue = prorationTypeId.ToString();

                chargeBackDeptId = objGiftCard.GiftCardTiers[0].ChargeBackDeptId;
                if (chargeBackDeptId == 0)
                    chargeBackDeptId = ddlChargeBack.Items[0].Value.ConvertToInt32();
                ddlChargeBack.SelectedValue = chargeBackDeptId.ToString();

                amountTypeId = objGiftCard.GiftCardTiers[0].AmountTypeId;
                if (amountTypeId == 0)
                    amountTypeId = ddlValueType.Items[0].Value.ConvertToInt32();
                ddlValueType.SelectedValue = amountTypeId.ToString();
                UpdateCurrencyControls(amountTypeId, item);

                if (txtNameOfCard.Text == "" && objGiftCard.GiftCardTiers[0].Name != null)
                {
                    txtNameOfCard.Text = objGiftCard.GiftCardTiers[0].Name.ToString();
                }

                if (txtCardIdentifier.Text == "" && objGiftCard.GiftCardTiers[0].CardIdentifier != null)
                {
                    txtCardIdentifier.Text = objGiftCard.GiftCardTiers[0].CardIdentifier.ToString();
                }

            }


            RequiredFieldValidator RFV = item.FindControl("requirefieldValue") as RequiredFieldValidator;
            RFV.ErrorMessage = PhraseLib.Lookup("giftcardEdit.invalidvalue", LanguageID);
            RFV = requirefieldNameOfCard as RequiredFieldValidator;
            RFV.ErrorMessage = PhraseLib.Lookup("giftcardEdit.invalidname", LanguageID);
            RegularExpressionValidator REV = IdentiferValidator as RegularExpressionValidator;
            REV.ErrorMessage = PhraseLib.Lookup("giftcardEdit.invalidid", LanguageID);
            vsError.Text = PhraseLib.Lookup("error.summary", LanguageID);

            if (!objOffer.IsTemplate && DisabledAttribute)
                SetDisableAttribute(item);
            if (!prodCondExists || (SystemCacheData.GetSystemOption_UE_ByOptionId(237) == "0"))
            {
                if (ddlProrationRate != null && ddlProrationRate.Enabled)
                {
                    ddlProrationRate.Enabled = false;
                }
            }

        }

        if (!objOffer.IsTemplate)
        {
            if (DisabledAttribute)
                chkRollUp.Attributes["disabled"] = DisabledAttribute.ToString();
        }
    }

    private bool DoesProductConditionExists()
    {
        bool conditionExists = false;
        AMSResult<List<RegularProductCondition>> listProdCondition = m_Offer.GetRegularProductConditionsByOfferId(OfferID);
        if (listProdCondition.Result.Count > 0 && listProdCondition.Result[0].IncentiveProductGroupId > 0)
        {
            conditionExists = true;
        }
        return conditionExists;
    }
    private void CreateObjectAndSave()
    {
        bool isValid = true;
        string errMsg = String.Empty;
        decimal previousTierValue = -1;
        decimal currentTierValue = 0;
        objGiftCard.GiftCardTiers.Clear();
        objGiftCard.DisallowEdit = chkDisallow_Edit.Checked;
        objGiftCard.RewardOptionPhase = Phase;
        objGiftCard.RewardOptionId = RewardID;
        objGiftCard.Rollup = chkRollUp.Checked;
        //We dont have Required to deliver option in giftcard. Hence setting it to false by default.
        objGiftCard.Required = false;
        GiftCardTier objTierData;

        if (!string.IsNullOrWhiteSpace(txtNameOfCard.Text))
        {

            foreach (RepeaterItem item in repGiftcard.Items)
        {
            //To do: Fill all the tiers data  and update the giftcard with tiers
            objTierData = new GiftCardTier();
            objTierData.GCTierTranslations = new List<GiftCardTierTranslation>();

            if (ddlProrationRate != null)
            {
                objTierData.ProrationTypeId = ddlProrationRate.SelectedValue.ConvertToInt32();
            }

            objTierData.AmountTypeId = ddlValueType.SelectedValue.ConvertToInt32();

            TextBox txtValue = item.FindControl("txtValue") as TextBox;
            if (txtValue != null)
            {
                currentTierValue = m_CommonInc.Extract_Decimal(txtValue.Text, CurrentUser.AdminUser.Culture);
                if (currentTierValue > 0)
                {
                    if (previousTierValue < currentTierValue)
                    {
                        previousTierValue = currentTierValue;
                        objTierData.Amount = currentTierValue;
                    }
                    else if (previousTierValue >= currentTierValue)
                    {
                        isValid = false;
                        errMsg = PhraseLib.Lookup("condition.tiervalues", LanguageID);
                        UpdateCurrencyControls(objTierData.AmountTypeId, item);
                    }
                }
                else
                {
                    isValid = false;
                    errMsg = PhraseLib.Lookup("error.invalid-amount-billion", LanguageID);
                }
            }
            if (txtNameOfCard != null)
            {
                objTierData.Name = txtNameOfCard.Text;
            }
            //Get buydescription
            FindNestedControl(item, "tbMLI");
            if (resultControl != null)
            {
                TextBox txtBuyDesc = (TextBox)resultControl;
                objTierData.BuyDescription = txtBuyDesc.Text;

                resultControl = null;
            }
            if (ddlChargeBack != null)
            {
                objTierData.ChargeBackDeptId = ddlChargeBack.SelectedValue.ConvertToInt32();
            }

            if (txtCardIdentifier != null)
            {
                objTierData.CardIdentifier = txtCardIdentifier.Text;
            }

            objTierData.TierLevel = item.ItemIndex + 1;

            //Add translations
            if (isMultiLanguageEnabled)
            {
                FindNestedControl(item, "repMLIInputs");
                if (resultControl != null)
                {
                    Repeater repMLI = (Repeater)resultControl;

                    List<GiftCardTierTranslation> list = SetTierTranslations(repMLI, objTierData.Id, objTierData.BuyDescription);
                    if (list != null)
                        objTierData.GCTierTranslations = list;

                    resultControl = null;
                }
            }

            //Add the updated tier to giftcard
            objGiftCard.GiftCardTiers.Add(objTierData);
            }

        }
        else
        {
            isValid = false;
            errMsg = PhraseLib.Lookup("giftcardEdit.invalidname", LanguageID);
            requirefieldNameOfCard.IsValid = false;
        }


        if (!isValid)
        {
            DisplayError(errMsg);
            return;
        }

        IsNewGiftCard = (objGiftCard.Id == 0);
        //save giftcard object
        AMSResult<bool> result = m_GiftCardRewardService.CreateUpdateGiftCardReward(objGiftCard, objOffer.OfferID, objOffer.EngineID);
        //save the templates permission 
        if (objOffer.IsTemplate)
        {
            //time to update the status bits for the templates
            int form_Disallow_Edit = 0;
            string[] LockFieldsList = null;

            form_Disallow_Edit = chkDisallow_Edit.Checked == true ? 1 : 0;
            if (!string.IsNullOrEmpty(hdnLockedTemplateFields.Value))
            {
                LockFieldsList = hdnLockedTemplateFields.Value.Split(',');
                m_Commondata.PurgeFieldLevelPermissions(AppName, objOffer.OfferID, 0);
                m_Commondata.UpdateDeliverableAndFieldLevelPermissions(objGiftCard.RewardID.ConvertToInt32(), objOffer.OfferID, form_Disallow_Edit, LockFieldsList);
            }
            else
            {
                m_Commondata.PurgeFieldLevelPermissions(AppName, objOffer.OfferID, 0);
            }
        }

        if (result.ResultType != AMSResultType.Success)
            DisplayError(result.GetLocalizedMessage<bool>(LanguageID));
        else
        {
            Description = IsNewGiftCard ? "reward.creategiftcard" : "reward.editgiftcard";
            m_Offer.UpdateOfferStatusToModified(OfferID, (int)Engines.UE, CurrentUser.AdminUser.ID);
            m_OAWService.ResetOfferApprovalStatus(OfferID);
            m_ActivityLogService.Activity_Log(ActivityTypes.GiftCard, OfferID, CurrentUser.AdminUser.ID, PhraseLib.Lookup(Description, LanguageID));

            //ScriptManager.RegisterStartupScript(this, this.GetType(), "Close", "window.close();", true);
            RegisterScript("Close", "window.close();");
        }
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
                ucMLI.MLIdentifierValue = Convert.ToInt64(DataBinder.Eval(e.Item.DataItem, "Id"));
                ucMLI.MLTableName = "GiftCardTierTranslation";
                ucMLI.MLColumnName = "BuyDescription";
                ucMLI.MLITranslationColumn = "GiftCardTierId";
                ucMLI.IsMultiLanguageEnabled = isMultiLanguageEnabled;
                //ucMLI.DisablePopup = disablePopup;
                if (DataBinder.Eval(e.Item.DataItem, "BuyDescription") != null)
                    ucMLI.MLDefaultLanguageStandardValue = DataBinder.Eval(e.Item.DataItem, "BuyDescription").ToString();
                divControl.Controls.Add(ucMLI);
            }
        }
    }

    private void GetQueryString()
    {
        OfferID = Request.QueryString["OfferID"].ConvertToInt32();
        DeliverableID = Request.QueryString["DeliverableID"].ConvertToInt32();
        Phase = Request.QueryString["Phase"].ConvertToInt32();
        RewardID = Request.QueryString["RewardID"].ConvertToInt32();
        ProductCondition = Request.QueryString["productCondition"].ConvertToBool();
        percentOffAllowed = Request.QueryString["PercentOffAllowed"].ConvertToBool();
        restrictProrationTypeToAllConditional = Request.QueryString["RestrictProrationTypeToAllConditional"].ConvertToBool();
    }

    private void DisplayError(string err)
    {
        infobar.Attributes["class"] = "red-background";
        infobar.InnerHtml = err;
        infobar.Visible = true;
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
    private void DisableControls()
    {
        //CurrentUser.UserPermissions = RefreshUserPermissions(CurrentUser.AdminUser.ID);
        if (objOffer.FromTemplate)
            DisabledAttribute = ((CurrentUser.UserPermissions.EditOffer && !objGiftCard.DisallowEdit) ? false : true);
        else if (objOffer.IsTemplate)
            DisabledAttribute = CurrentUser.UserPermissions.EditTemplates ? false : true;
        else
            DisabledAttribute = CurrentUser.UserPermissions.EditOffer ? false : true;

        //If disable is set to false, check Buyer conditions
        if (!DisabledAttribute)
        {
            if (m_CommonInc.LRTadoConn.State == ConnectionState.Closed)
                m_CommonInc.Open_LogixRT();
            DisabledAttribute = ((CurrentUser.UserPermissions.EditOffersRegardlessBuyer || m_CommonInc.IsOfferCreatedWithUserAssociatedBuyer(CurrentUser.AdminUser.ID, OfferID)) ? false : true);
            m_CommonInc.Close_LogixRT();
            //Hide save button if Disable is true
            btnSave.Visible = !DisabledAttribute;
        }

    }
    private List<GiftCardTierTranslation> SetTierTranslations(Repeater repeater, long gcTierId, string defaultBuyDescription)
    {
        GiftCardTierTranslation gct = null;
        List<GiftCardTierTranslation> listGct = new List<GiftCardTierTranslation>();

        //((TextBox)repMLI.Items[0].FindControl("tbTranslation")).Text
        foreach (RepeaterItem ri in repeater.Items)
        {

            HiddenField hfLangId = (HiddenField)ri.FindControl("hfLangId");
            int langId = Convert.ToInt32(hfLangId.Value);
            TextBox tbTranslation = (TextBox)ri.FindControl("tbTranslation");

            gct = new GiftCardTierTranslation();
            gct.LanguageId = langId;
            gct.GiftCardTierId = gcTierId;
            if (langId == SystemCacheData.GetSystemOption_General_ByOptionId(125).ConvertToInt32())
                gct.BuyDescription = defaultBuyDescription;
            else
                gct.BuyDescription = tbTranslation.Text;

            listGct.Add(gct);
        }
        return listGct;
    }
    private void ResolveDependencies()
    {
        m_GiftCardRewardService = CurrentRequest.Resolver.Resolve<CMS.AMS.Contract.IGiftCardRewardService>();
        m_Offer = CurrentRequest.Resolver.Resolve<CMS.AMS.Contract.IOffer>();
        m_Commondata = CurrentRequest.Resolver.Resolve<CMS.AMS.Common>();
        m_CommonInc = CurrentRequest.Resolver.Resolve<CommonInc>();

        m_Commondata.Open_LogixRT();
        m_LocalizationService = CurrentRequest.Resolver.Resolve<ILocalizationService>();
        m_ActivityLogService = CurrentRequest.Resolver.Resolve<IActivityLogService>();
        m_OAWService = CurrentRequest.Resolver.Resolve<IOfferApprovalWorkflowService>();
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

    private void UpdateSaveButton(int overriddenFieldsCount)
    {
        if ((!objOffer.IsTemplate && !CurrentUser.UserPermissions.EditOffer) ||
            (objOffer.IsTemplate && !CurrentUser.UserPermissions.EditTemplates) ||
            (objOffer.FromTemplate && objGiftCard.DisallowEdit && overriddenFieldsCount == 0) ||
            (objOffer.FromTemplate && !objGiftCard.DisallowEdit && overriddenFieldsCount == (6 * objOffer.NumbersOfTier + 2)) ||
            m_OAWService.CheckIfOfferIsAwaitingApproval(OfferID).Result)              //((6*number of tiers) + 2) fields will be locked if all fields are selected in GCR template
            btnSave.Visible = false;

    }

    private void UpdateCurrencyControls(int selectedValue, RepeaterItem item)
    {
        CustomValidator RV = item.FindControl("customValidator") as CustomValidator;
        if (selectedValue == (int)CPEAmountTypes.PercentageOff)
        {
            ValueTypeInitialPrefix = string.Empty;
            ValueTypeInitialSuffix = "%";

            RV.ErrorMessage = PhraseLib.Lookup("error.GC-percent-over", LanguageID);
        }
        else if (selectedValue == (int)CPEAmountTypes.FixedAmountOff)
        {
            ValueTypeInitialPrefix = CurrencySymbol;
            ValueTypeInitialSuffix = CurrencyAbbr;

            RV.ErrorMessage = PhraseLib.Lookup("error.GC-value-over", LanguageID);
        }
    }
    private Hashtable GetOverriddenFields()
    {
        Hashtable overrideFields = new Hashtable();
        //Get the Lockable fields data.
        DataTable dt = m_Commondata.GetFieldLevelPermissions(OfferID, AppName);
        List<string> LockedTemplateFields = new List<string>();
        foreach (DataRow row in dt.Rows)
        {
            if (row["Tiered"].ConvertToBool() == true)
            {
                for (int i = 0, j = 0; i < objOffer.NumbersOfTier; i++)
                {
                    overrideFields.Add(repGiftcard.ID + "$ctl0" + j + "$" + row["ControlName"].ConvertToString(), row["Editable"].ConvertToBool());
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
        hdnLockedTemplateFields.Value = string.Join(",", LockedTemplateFields);

        return overrideFields;
    }
    private StringBuilder GetOfferFromTemplateScript(Hashtable overrideFields)
    {
        StringBuilder script = new StringBuilder();
        if (overrideFields.Count > 0)
        {
            string OvrdFldDisabled = "False";
            foreach (DictionaryEntry de in overrideFields)
            {
                script.Append("elem = document.mainform." + de.Key.ToString() + ";");
                OvrdFldDisabled = de.Value.ToString().ToUpper() == "TRUE" ? "false" : "true";
                script.Append("if (elem != null) { elem.disabled = " + OvrdFldDisabled + "; }");
            }
        }
        return script;
    }
    private StringBuilder GetTemplateScript(Hashtable overrideFields, Hashtable overrideDiv, string editable)
    {
        StringBuilder script = new StringBuilder("xmlhttpPost(\"UEtemplateFeeds.aspx?OfferID=" + objOffer.OfferID + "&PageName=" + AppName + "&PageEditable=" + editable + "\");");

        //Update the locked fields.
        if (overrideFields.Count > 0)
        {
            foreach (DictionaryEntry de in overrideFields)
            {
                script.Append("  var elem = null;");
                if (overrideDiv[de.Key.ConvertToString()].ConvertToString().ToUpper() == "TRUE")
                    script.Append("  elem = document.getElementById(" + de.Key.ConvertToString() + "Div);");
                else
                    script.Append("  elem = document.mainform." + de.Key.ConvertToString() + ";");

                OvrdFldClass = de.Value.ConvertToString().ToUpper() == "TRUE" ? "enabledTemplateField" : "disabledTemplateField";
                script.Append("  if (elem != null) { elem.setAttribute('className', '" + OvrdFldClass + "'); }");
                script.Append("  if (elem != null) { elem.className = '" + OvrdFldClass + "'; }");
            }
        }
        return script;
    }
    private void SetMLIPopupProperty()
    {
        //Set the popup enable\disable
        //mliOverriddenInTemplate is true when the buydescription control for multi\single language is overridden
        if (isMultiLanguageEnabled
            &&
            (objOffer.FromTemplate && (objGiftCard.DisallowEdit && !mliOverriddenInTemplate) || (!objGiftCard.DisallowEdit && mliOverriddenInTemplate))
            ||
            (!objOffer.FromTemplate && !objOffer.IsTemplate && DisabledAttribute))  //DisabledAttribute false means edit permissions are there for offer
        {
            foreach (RepeaterItem item in repGiftcard.Items)
            {
                if (item.FindControl("ucMLI") != null)
                {
                    logix_UserControls_MultiLanguagePopup popup = (logix_UserControls_MultiLanguagePopup)item.FindControl("ucMLI");
                    popup.DisablePopup = true;
                }
            }
        }
    }
    private void LoadValueTypeDropDown(bool productCondition)
    {
        //filter value types based on Product condition. If product condition not set, skip Percent.
        SystemCacheData.ClearAllCacheData();
        DataTable dt = SystemCacheData.LoadDefaultValueTypesForGiftCard(CurrentUser.AdminUser.LanguageID);
        if (productCondition && percentOffAllowed)
        {
            dt.DefaultView.RowFilter = "";
        }
        else
        {
            dt.DefaultView.RowFilter = "AmountTypeID = 1";
        }
        ddlValueType.DataSource = dt.DefaultView;
        ddlValueType.DataTextField = "PHRASE";
        ddlValueType.DataValueField = "AmountTypeID";
        ddlValueType.DataBind();
    }
    private void SetBuyDescription(RepeaterItem item)
    {
        if (!isMultiLanguageEnabled)
        {
            FindNestedControl(item, "tbMLI");
            if (resultControl != null)
            {
                TextBox tbBuyDesc = (TextBox)resultControl;
                tbBuyDesc.Text = objGiftCard.GiftCardTiers[item.ItemIndex].BuyDescription;
            }
        }
        resultControl = null;
    }
    private void SetDisableAttribute(RepeaterItem item)
    {
        //set disable attributes
        TextBox txt = item.FindControl("txtValue") as TextBox;
        if (txt != null)
            txt.Attributes["disabled"] = "\"" + DisabledAttribute.ToString() + "\"";
        if (txtNameOfCard != null)
            txtNameOfCard.Attributes["disabled"] = "\"" + DisabledAttribute.ToString() + "\"";
        txt = item.FindControl("ucMLI$tbMLI") as TextBox;
        if (txt != null)
        {
            txt.Attributes["disabled"] = "\"" + DisabledAttribute.ToString() + "\"";
        }
        if (txtCardIdentifier != null)
            txtCardIdentifier.Attributes["disabled"] = "\"" + DisabledAttribute.ToString() + "\"";

        if (ddlValueType != null)
            ddlValueType.Attributes["disabled"] = "\"" + DisabledAttribute.ToString() + "\"";

        if (ddlProrationRate != null)
            ddlProrationRate.Attributes["disabled"] = "\"" + DisabledAttribute.ToString() + "\"";

        if (ddlChargeBack != null)
            ddlChargeBack.Attributes["disabled"] = "\"" + DisabledAttribute.ToString() + "\"";
    }
    private void LoadDropDown(DropDownList ddl, string textField, string valueField, DataTable sourceTable)
    {
        if (ddl != null)
        {
            ddl.DataSource = sourceTable;
            ddl.DataTextField = textField;
            ddl.DataValueField = valueField;
            ddl.DataBind();
        }

        // adding select item for ddlChargeBack
        if (ddl.ClientID == "ddlChargeBack")
        {
            ddl.Items.Insert(0, new ListItem("--" + PhraseLib.Lookup("term.select", LanguageID) + "--", "", true));
        }

    }
    # endregion priavate methods
}