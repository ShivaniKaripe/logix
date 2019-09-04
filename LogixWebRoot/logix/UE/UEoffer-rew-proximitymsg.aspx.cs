using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using CMS;
using CMS.AMS.Contract;
using CMS.AMS.Models;
using CMS.AMS;
using System.Data;
using System.Web.UI.HtmlControls;
using System.Globalization;

public partial class logix_UE_UEoffer_rew_proximitymsg : AuthenticatedUI
{
    #region Global variables
    private IOffer m_Offer;
    private CMS.AMS.Common m_Commondata;
    Copient.CommonInc m_CommonInc;
    private ILocalizationService m_LocalizationService;
    private IProximityMessageRewardService m_ProximityMessageRewardService;
    private IOfferApprovalWorkflowService m_OAWService;

    protected int OfferID;
    private int Phase;
    private int RewardID;
    private int PMID;

    private int DefaultLanguageID;
    private bool IsMultiLanguageEnabled;
    private int ThresholdTypeId;
    private UnitTypeInformation unitTypeInformation;

    //UI related variables
    protected int CurrencyPrecision;
    protected string phrasedecimalwarning;
    protected string phraseawayfromnumeric;
    protected string phraseawayvaluenontiers;
    protected string phraseawayvaluewithtiers;
    protected string phraseconditionwarning;
    protected string phrasetagrequired;
    protected string phrasetextareanum;
    protected string phrasemessagerequiredwarning;
    protected string phrasefortier;
    protected string phraseForLanguage;
    protected string phraseforthresholdmessage;
    protected string UnitPrecisionFormat;
    protected int UnitPrecision;
    protected string NumberDecimalSeparator;
    //Units related variables.
    private string UnitName;
    private string UnitSymbol;
    private string UnitAbbr;
    private string RequiredLabel;
    private string AwayLabel;
    private string TagLabel;
    bool DisabledAttribute = false;
    //Permissions related variables
    private bool UserEditFlag;
    public Int32 CustomerFacingLangID = 1;


    Copient.CommonInc MyCommon = new Copient.CommonInc();
    bool isTranslatedOffer = false;
    bool bEnableRestrictedAccessToUEOfferBuilder = false;
    #endregion

    #region Properties
    protected Offer objOffer
    {
        get { return ViewState["Offer"] as Offer; }
        set { ViewState["Offer"] = value; }
    }

    protected ProximityMessage objProximityMessage
    {
        get { return ViewState["ProximityMessage"] as ProximityMessage; }
        set { ViewState["ProximityMessage"] = value; }
    }

    protected int RewardOptionID
    {
        get { return Int32.Parse(ViewState["RewardOptionID"].ToString()); }
        set { ViewState["RewardOptionID"] = value; }
    }

    protected List<string> QuantityName
    {
        get { return ViewState["QuantityName"] as List<string>; }
        set { ViewState["QuantityName"] = value; }
    }

    protected List<int> QuantityUnit
    {
        get { return ViewState["QuantityUnit"] as List<int>; }
        set { ViewState["QuantityUnit"] = value; }
    }

    protected List<string> QuantityValue
    {
        get { return ViewState["QuantityValue"] as List<string>; }
        set { ViewState["QuantityValue"] = value; }
    }

    protected List<Language> Languages
    {
        get { return ViewState["Languages"] as List<Language>; }
        set { ViewState["Languages"] = value; }
    }

    protected int PMCount
    {
        get { return Int32.Parse(ViewState["PMCount"].ToString()); }
        set { ViewState["PMCount"] = value; }
    }

    protected int PreviousThresholdTypeId
    {
        get { return Int32.Parse(ViewState["PreviousThresholdTypeId"].ToString()); }
        set { ViewState["PreviousThresholdTypeId"] = value; }
    }
    #endregion

    #region Protected Methods
    /// <summary>
    /// 
    /// </summary>
    /// <param name="e"></param>
    protected override void OnInit(EventArgs e)
    {
        AppName = "UEoffer-rew-proximitymsg.aspx";
        base.OnInit(e);

    }

    /// <summary>
    /// This function is called when proximity message reward is selected.
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    protected void Page_Load(object sender, EventArgs e)
    {
        AMSResult<SystemOption> objResult = SystemSettings.GetGeneralSystemOption(125);
        bEnableRestrictedAccessToUEOfferBuilder = MyCommon.Fetch_SystemOption(249) == "1" ? true : false;
        if (objResult.ResultType == AMSResultType.Success)
        {
            Int32.TryParse(objResult.Result.OptionValue, out CustomerFacingLangID);
        } ResolveDependencies();
        GetQueryString();

        DefaultLanguageID = SystemSettings.GetSystemDefaultLanguage().LanguageID;
        IsMultiLanguageEnabled = SystemSettings.IsMultiLanguageEnabled();

        if (objOffer == null)
        {
            if (OfferID == 0)
            {
                //DisplayError(PhraseLib.Lookup("error.invalidoffer", LanguageID));
                return;
            }
            m_Commondata.QueryStr = "select RewardOptionId from CPE_RewardOptions where IncentiveID = " + OfferID;
            var result = m_Commondata.LRT_Select();
            if (result.Rows.Count > 0)
                RewardOptionID = result.Rows[0][0].ConvertToInt32();

            objOffer = m_Offer.GetOffer(OfferID, CMS.AMS.Models.LoadOfferOptions.AllRegularConditions);
            GetOfferRelatedData();
            Languages = SystemSettings.GetAllActiveLanguages((Engines)objOffer.EngineID);

            AssignHiddenFieldLanguages();
        }

        ucTemplateLockableFields.OfferId = OfferID;
        if (PMID != 0)
            ucTemplateLockableFields.DeliverableId = RewardID;
        if (!IsPostBack)
        {
            CheckPermissions();
            SetupAndLocalizePage();
            LoadPageData();
            ApplyPermissions();
            Page.Header.DataBind();
            EnableDisableTemplateFields();
        }
    }
    private void AssignHiddenFieldLanguages()
    {
        List<string> languages = new List<string>();
        foreach (Language lan in Languages)
        {
            if (lan.LanguageID == DefaultLanguageID)
                languages.Add(lan.Name + ":Default");
            else
                languages.Add(lan.Name);
        }

        hdnLanguages.Value = string.Join(",", languages);
    }
    private void SetupAndLocalizePage()
    {
        SetControlsTexts();

    }
    private void UpdateSaveButton()
    {
        if ((!objOffer.IsTemplate && !CurrentUser.UserPermissions.EditOffer) ||
            (objOffer.IsTemplate && !CurrentUser.UserPermissions.EditTemplates) ||
            (objOffer.FromTemplate && ucTemplateLockableFields.RewardTemplateFieldSource.DisallowEdit && ucTemplateLockableFields.ExceptionFields.Count == 0) ||
            (objOffer.FromTemplate && !ucTemplateLockableFields.RewardTemplateFieldSource.DisallowEdit && ucTemplateLockableFields.ExceptionFields.Count == (2 * objOffer.NumbersOfTier + 2)))              //((6*number of tiers) + 2) fields will be locked if all fields are selected in GCR template
            btnSave.Visible = false;
        if (bEnableRestrictedAccessToUEOfferBuilder && isTranslatedOffer)
        {
            btnSave.Visible = false;
        }

    }
    private void DisableControls()
    {
        CurrentUser.UserPermissions = RefreshUserPermissions(CurrentUser.AdminUser.ID);
        if (objOffer.FromTemplate)
            DisabledAttribute = ((CurrentUser.UserPermissions.EditOffer && !ucTemplateLockableFields.RewardTemplateFieldSource.DisallowEdit) ? false : true);
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
    private void SetControlsTexts()
    {
        int languageid = base.CurrentUser.AdminUser.LanguageID;
        AssignPageTitle("term.offer", "term.proximitymessagereward", OfferID.ToString());
        btnSave.Text = PhraseLib.Lookup("term.save", languageid);
        btnPreview.Text = PhraseLib.Lookup("term.preview", languageid);
        lblMessageType.Text = PhraseLib.Lookup("term.messagetype", languageid) + ":";

        if (objOffer.IsTemplate)
        {
            title.InnerText = PhraseLib.Lookup("term.template", languageid) + " #" + OfferID + " " + PhraseLib.Lookup("term.proximitymessagereward", languageid);
        }
        else
            title.InnerText = PhraseLib.Lookup("term.offer", languageid) + " #" + OfferID + " " + PhraseLib.Lookup("term.proximitymessagereward", languageid);

        phrasedecimalwarning = PhraseLib.Lookup("term.decimalawayfromwarning", languageid);
        phraseawayfromnumeric = PhraseLib.Lookup("term.AwayFromNumeric", languageid);
        phraseawayvaluenontiers = PhraseLib.Lookup("term.AwayFromValueNonTiers", languageid);
        phraseawayvaluewithtiers = PhraseLib.Lookup("term.AwayFromValueWithTiers", languageid);
        phraseconditionwarning = PhraseLib.Lookup("term.conditionwarning", languageid).ToString().Replace("[]", (ddlMessageType.SelectedValue.ConvertToInt32() == 9 ? "Points" : "Product"));
        phrasetagrequired = PhraseLib.Lookup("term.tagrequired", languageid);
        phrasetextareanum = PhraseLib.Lookup("term.textareanumber", languageid);
        phrasemessagerequiredwarning = PhraseLib.Lookup("term.messagerequiredwarning", languageid);
        phrasefortier = PhraseLib.Lookup("term.fortier", languageid);
        phraseForLanguage = PhraseLib.Lookup("term.language", languageid).ToLower();
        phraseforthresholdmessage = PhraseLib.Lookup("term.thresholdmessage", languageid);
        NumberDecimalSeparator = CurrentUser.AdminUser.Culture.NumberFormat.NumberDecimalSeparator;
    }

    protected void ucTemplateLockableFields_Onload(object sender, EventArgs e)
    {
        FromTemplate();
        IsTemplate();
        UpdateSaveButton();
    }
    private void IsTemplate()
    {
        if (objOffer.IsTemplate && !DisabledAttribute)
        {
            RewardTemplateFieldContainer rc = ucTemplateLockableFields.RewardTemplateFieldSource;
            if (rc != null)
            {
                foreach (CMS.AMS.Models.TemplateField temp in rc.TemplateFieldList)
                {
                    if (temp.Tiered)
                    {
                        for (int i = 0, j = 0; i < objOffer.NumbersOfTier; i++)
                        {
                            if (temp.FieldName == "Message")
                            {
                                List<Language> lang = SystemSettings.GetAllActiveLanguages((Engines)objOffer.EngineID);
                                var k = 0;
                                foreach (Language lan in lang)
                                {
                                    var tt = (HtmlControl)this.FindControl(repProximityMsg.ID + "$ctl0" + j + "$" + "repProximityMsgDesc" + "$ctl0" + k + "$" + temp.ControlName);
                                    if (rc.DisallowEdit == true && temp.Editable)
                                    {
                                        tt.Style.Add("background", "#bfffff");
                                    }
                                    else if (!rc.DisallowEdit && !temp.Editable)
                                        tt.Style.Add("background", "#ffdddd");
                                    k++;
                                }
                            }
                            else
                            {
                                var t = (WebControl)this.FindControl(repProximityMsg.ID + "$ctl0" + j + "$" + temp.ControlName);
                                if (rc.DisallowEdit == true && temp.Editable)
                                {
                                    t.Style.Add("background-color", "#bfffff");
                                }
                                else if (!rc.DisallowEdit && !temp.Editable)
                                {
                                    t.Style.Add("background-color", "#ffdddd");
                                }
                            }
                            j += 1;
                        }
                    }
                    else
                    {
                        var t = (WebControl)this.FindControl(temp.ControlName);
                        if (rc.DisallowEdit == true && temp.Editable)
                        {
                            t.Style.Add("background-color", "#bfffff");
                        }
                        else if (!rc.DisallowEdit && !temp.Editable)
                        {
                            t.Style.Add("background-color", "#ffdddd");
                        }
                    }
                }
            }
        }
    }
    private void FromTemplate()
    {
        if (objOffer.FromTemplate)
        {
            RewardTemplateFieldContainer rc = ucTemplateLockableFields.RewardTemplateFieldSource;
            if (rc != null)
            {
                foreach (CMS.AMS.Models.TemplateField temp in rc.TemplateFieldList)
                {
                    if (temp.Tiered)
                    {
                        for (int i = 0, j = 0; i < objOffer.NumbersOfTier; i++)
                        {
                            if (temp.FieldName == "Message" && !temp.Editable)
                            {
                                List<Language> lang = SystemSettings.GetAllActiveLanguages((Engines)objOffer.EngineID);
                                var k = 0;
                                foreach (Language lan in lang)
                                {
                                    var tt = (HtmlControl)this.FindControl(repProximityMsg.ID + "$ctl0" + j + "$" + "repProximityMsgDesc" + "$ctl0" + k + "$" + temp.ControlName);
                                    if (tt != null)
                                        tt.Attributes.Add("disabled", "true");
                                    k++;
                                }
                            }
                            else
                            {
                                if (!temp.Editable)
                                {
                                    var t = (WebControl)this.FindControl(repProximityMsg.ID + "$ctl0" + j + "$" + temp.ControlName);
                                    t.Attributes.Add("disabled", "true");
                                }
                            }
                            j += 1;
                        }
                    }
                    else
                    {
                        if (!temp.Editable)
                        {
                            if (temp.FieldName == "Tag")
                            {
                                var t = (WebControl)this.FindControl(temp.ControlName);
                                t.Attributes.Add("disabled", "true");
                            }
                            else
                            {
                                var t = (DropDownList)this.FindControl(temp.ControlName);
                                t.Attributes.Add("disabled", "true");
                            }
                        }
                    }
                }
            }
        }
    }
    private void EnableDisableTemplateFields()
    {
        if (objOffer.IsTemplate)
        {
            ucTemplateLockableFields.Visible = true;
            ucTemplateLockableFields.OfferId = OfferID;
            ucTemplateLockableFields.PageName = AppName;
            ucTemplateLockableFields.LanguageId = DefaultLanguageID;
            ucTemplateLockableFields.RewardLockStatus = objProximityMessage.DisallowEdit;
        }
    }

    /// <summary>
    /// This method is called whenever new tier is created to display
    /// textarea for description.
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    protected void BindProximityMsgDesc(object sender, RepeaterItemEventArgs e)
    {
        if (IsMultiLanguageEnabled)
        {
            Repeater childRepeater = (Repeater)e.Item.FindControl("repProximityMsgDesc");
            ProximityMessageTier tempProximityTier = (ProximityMessageTier)e.Item.DataItem;
            if (tempProximityTier.PMTierTranslations.Count != Languages.Count)
            {
                ProximityMessageTierTranslation tempPMTierTranslation;
                foreach (var lan in Languages)
                {
                    if (!tempProximityTier.PMTierTranslations.Exists(l => l.LanguageId == lan.LanguageID))
                    {
                        tempPMTierTranslation = new ProximityMessageTierTranslation();
                        tempPMTierTranslation.LanguageId = lan.LanguageID;
                        tempPMTierTranslation.Message = "";
                        tempProximityTier.PMTierTranslations.Add(tempPMTierTranslation);
                    }
                }
            }
            childRepeater.DataSource = tempProximityTier.PMTierTranslations;
            childRepeater.DataBind();
        }
        else
        {
            List<ProximityMessageTier> tempProximityTier = new List<ProximityMessageTier>();
            Repeater childRepeater = (Repeater)e.Item.FindControl("repProximityMsgDesc");
            tempProximityTier.Add((ProximityMessageTier)e.Item.DataItem);
            childRepeater.DataSource = tempProximityTier;
            childRepeater.DataBind();
        }
    }

    /// <summary>
    /// This method is called when save button is clicked on UI.
    /// Updates offer object and saves it.
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    protected void Save_Click(object sender, EventArgs e)
    {
        try
        {
            UpdateObjectAndSave();
        }
        catch (Exception ex)
        {
            //DisplayError(ex.Message);
        }
    }

    /// <summary>
    /// This method is called when message type is changed in UI.
    /// Updates UI and also updates object. 
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    protected void Message_Changed(object sender, EventArgs e)
    {
        ThresholdTypeId = Int32.Parse(ddlMessageType.SelectedValue);
        UpdateUnitRelatedInformation();
        LoadDefaultData();
    }

    #endregion

    #region Private Methods
    /// <summary>
    /// Resolve all required dependencies.
    /// </summary>
    private void ResolveDependencies()
    {
        m_Offer = CurrentRequest.Resolver.Resolve<CMS.AMS.Contract.IOffer>();
        m_Commondata = CurrentRequest.Resolver.Resolve<CMS.AMS.Common>();
        m_Commondata.Open_LogixRT();
        m_LocalizationService = CurrentRequest.Resolver.Resolve<ILocalizationService>();
        m_ProximityMessageRewardService = CurrentRequest.Resolver.Resolve<CMS.AMS.Contract.IProximityMessageRewardService>();
        m_OAWService = CurrentRequest.Resolver.Resolve<IOfferApprovalWorkflowService>();
        m_CommonInc = CurrentRequest.Resolver.Resolve<Copient.CommonInc>();
    }

    /// <summary>
    /// Get Offer related data. Ex: OfferId, RewardOptionId, Phase, RewardID, ConditionType.
    /// </summary>
    private void GetQueryString()
    {
        OfferID = Request.QueryString["OfferID"].ConvertToInt32();
        Phase = Request.QueryString["Phase"].ConvertToInt32();
        RewardID = Request.QueryString["RewardID"].ConvertToInt32();
        PMID = Request.QueryString["PMID"].ConvertToInt32();
        isTranslatedOffer = MyCommon.IsTranslatedUEOffer(Convert.ToInt32(OfferID), MyCommon);
    }

    /// <summary>
    /// We get offer related conditions and data related to it.
    /// </summary>
    private void GetOfferRelatedData()
    {
        string tempStr = "";
        int unitType;
        QuantityUnit = new List<int>();
        QuantityName = new List<string>();
        QuantityValue = new List<string>();
        DataTable result1;

        #region Product Condition related data
        var productConditionDetails = m_Offer.GetRegularProductConditionsByOfferId(OfferID).Result;
        if (productConditionDetails.Count == 1 && (double)productConditionDetails.FirstOrDefault().QtyValue != 0.1 && (double)productConditionDetails.FirstOrDefault().QtyValue != 0.01 && (productConditionDetails.FirstOrDefault().QtyUnitType != 4 || (productConditionDetails.FirstOrDefault().QtyUnitType != 1 ? productConditionDetails.FirstOrDefault().QtyUnitType != 1 : true)))
        {
            if (((double)productConditionDetails.FirstOrDefault().QtyValue >= 2 && (double)productConditionDetails.FirstOrDefault().QtyUnitType == 1) || (double)productConditionDetails.FirstOrDefault().QtyUnitType >1)
            {
                unitType = productConditionDetails.FirstOrDefault().QtyUnitType;
                QuantityUnit.Add(unitType);

                //Get QuantityName
                m_Commondata.QueryStr = "select PhraseID from CPE_UnitTypes where UnitTypeID = " + unitType;
                result1 = m_Commondata.LRT_Select();
                if (result1.Rows.Count > 0)
                {
                    QuantityName.Add(PhraseLib.Lookup((int)result1.Rows[0][0], CurrentUser.AdminUser.LanguageID));
                }

                //Get QuantityValue(s)
                if (productConditionDetails.FirstOrDefault().RegularProductConditionTiers.Count > 0)
                {
                    foreach (var row in productConditionDetails.FirstOrDefault().RegularProductConditionTiers)
                    {
                        tempStr += row.TierLevel + "," + row.Quantity + ";";
                    }
                    for (var i = productConditionDetails.FirstOrDefault().RegularProductConditionTiers.Count; i < objOffer.NumbersOfTier; i++)
                    {
                        tempStr += (i + 1) + ",0;";
                    }
                    QuantityValue.Add(tempStr);
                }
            }
        }
        #endregion

        #region Points Condition related data
        tempStr = "";
        int PhraseID = 0;
        //Get Points Condition related data.
        if (objOffer.PointsProgramConditions != null)
        {
            if (objOffer.PointsProgramConditions.Count == 1)
            {
                var pointsConditionDetails = m_Offer.GetRegularPointConditionsByOfferId(OfferID).Result;
                unitType = (int)CPEUnitTypes.Points;

                //Get QuantityName
                m_Commondata.QueryStr = "select PhraseID from CPE_UnitTypes where UnitTypeID = " + unitType;
                result1 = m_Commondata.LRT_Select();
                if (result1.Rows.Count > 0)
                {
                    PhraseID = (int)result1.Rows[0][0];
                }

                //Get QuantityValue(s)
                if (pointsConditionDetails.FirstOrDefault().RegularPointConditionTiers.Count > 0)
                {
                    foreach (var row in pointsConditionDetails.FirstOrDefault().RegularPointConditionTiers)
                    {
                        if (row.Quantity.ToString() != "1")
                            tempStr += row.TierLevel.ToString() + "," + row.Quantity.ToString() + ";";
                    }
                    for (var i = pointsConditionDetails.FirstOrDefault().RegularPointConditionTiers.Count; i < objOffer.NumbersOfTier; i++)
                    {
                        tempStr += (i + 1) + ",0;";
                    }
                    if (tempStr != "")
                    {
                        QuantityValue.Add(tempStr);
                        QuantityUnit.Add(unitType);
                        QuantityName.Add(PhraseLib.Lookup(PhraseID, CurrentUser.AdminUser.LanguageID));
                    }
                }
            }
        }
        #endregion

        #region Proximity Message related data
        AMSResult<List<ProximityMessage>> tempPM = m_ProximityMessageRewardService.GetProximityMessageReward(OfferID, 9);
        if (tempPM.ResultType == AMSResultType.Success)
        {
            PMCount = tempPM.Result.Count;
            if (PMCount == 1)
            {
                PreviousThresholdTypeId = tempPM.Result.FirstOrDefault().ThresholdTypeId;
                if (PreviousThresholdTypeId == (int)CPEUnitTypes.Points && QuantityUnit.Count() == 2)
                    ThresholdTypeId = QuantityUnit[0];
                else
                    ThresholdTypeId = (int)CPEUnitTypes.Points;
            }
            else if (PMCount == 0)
                ThresholdTypeId = QuantityUnit[0];
        }
        else
        {
            ThresholdTypeId = QuantityUnit[0];
        }

        #endregion
    }

    private void CheckPermissions()
    {
        CurrentUser.UserPermissions = RefreshUserPermissions(CurrentUser.AdminUser.ID);
        UserEditFlag = CurrentUser.UserPermissions.EditOffer;
    }

    /// <summary>
    /// This method creates and populates PM related data.
    /// </summary>
    private void LoadPageData()
    {
        SetProximityMessageForLoad();
        List<ProximityMessage> tempObjPM;
        if (!IsPostBack)
        {
            AMSResult<List<ProximityMessage>> tempPM = m_ProximityMessageRewardService.GetProximityMessageReward(OfferID, 9);
            if (tempPM.ResultType == AMSResultType.Success && tempPM.Result.Count > 0)
            {
                tempObjPM = tempPM.Result;
                if (PMID == 0)
                {
                    if (PMCount == 0)
                    {
                        objProximityMessage = tempObjPM.FirstOrDefault();
                        if (objProximityMessage != null)
                            ThresholdTypeId = objProximityMessage.ThresholdTypeId;
                    }
                }
                else
                {
                    objProximityMessage = tempObjPM.Where(p => p.Id == PMID).FirstOrDefault();
                    if (objProximityMessage != null)
                        ThresholdTypeId = objProximityMessage.ThresholdTypeId;
                }
            }
        }
        UpdateUnitRelatedInformation();
        LoadDefaultData();
    }

    /// <summary>
    /// Initializes or updates PM related data.
    /// </summary>
    private void SetProximityMessageForLoad()
    {
        if (objProximityMessage != null)
        {
        }
        else
        {
            objProximityMessage = new CMS.AMS.Models.ProximityMessage();
            objProximityMessage.Id = 0;
            objProximityMessage.RewardID = RewardID;
            objProximityMessage.RewardOptionPhase = Phase;
            objProximityMessage.RewardTypeID = (int)DELIVERABLE_TYPES.PROXIMITY_MESSAGE;
            objProximityMessage.ProximityMessageTiers = new List<ProximityMessageTier>();

            ProximityMessageTier tempProxMessgTier;

            for (int i = 1; i <= objOffer.NumbersOfTier; i++)
            {
                tempProxMessgTier = new ProximityMessageTier();
                tempProxMessgTier.TierLevel = i;
                tempProxMessgTier.PMTierTranslations = new List<ProximityMessageTierTranslation>();
                if (IsMultiLanguageEnabled)
                {
                    ProximityMessageTierTranslation tempPMTierTranslation;
                    foreach (var j in SystemSettings.GetAllActiveLanguages(Engines.UE))
                    {
                        tempPMTierTranslation = new ProximityMessageTierTranslation();
                        tempPMTierTranslation.ProximityMessageTierId = tempProxMessgTier.Id;
                        tempPMTierTranslation.LanguageId = j.LanguageID;
                        tempPMTierTranslation.Message = "";
                        tempProxMessgTier.PMTierTranslations.Add(tempPMTierTranslation);
                    }
                }
                objProximityMessage.ProximityMessageTiers.Add(tempProxMessgTier);
            }
        }
        repProximityMsg.DataSource = objProximityMessage.ProximityMessageTiers;
        repProximityMsg.DataBind();
    }

    /// <summary>
    /// Updates Unit name, abbreviation, symbol, precision and labels.
    /// </summary>
    private void UpdateUnitRelatedInformation()
    {
        switch (ThresholdTypeId)
        {
            case (int)CPEUnitTypes.Items:
                UnitName = QuantityName[0];
                UnitAbbr = PhraseLib.Lookup("term.items", CurrentUser.AdminUser.LanguageID);
                UnitPrecision = 0;
                UnitSymbol = "";
                RequiredLabel = PhraseLib.Lookup("term.quantityrequired", base.CurrentUser.AdminUser.LanguageID);
                AwayLabel = PhraseLib.Lookup("term.quantityaway", base.CurrentUser.AdminUser.LanguageID);
                TagLabel = PhraseLib.Lookup("term.qtyrequired", base.CurrentUser.AdminUser.LanguageID);
                break;

            case (int)CPEUnitTypes.Dollars:
                unitTypeInformation = m_LocalizationService.LoadCurrencyInfoForOffer(RewardOptionID, CurrentUser.AdminUser.LanguageID);
                UnitName = PhraseLib.Lookup(unitTypeInformation.Name, CurrentUser.AdminUser.LanguageID);
                QuantityName[0] = UnitName;
                UnitAbbr = unitTypeInformation.Abbrevation;
                UnitPrecision = unitTypeInformation.Precision;
                CurrencyPrecision = UnitPrecision;
                UnitSymbol = unitTypeInformation.Symbol;
                RequiredLabel = PhraseLib.Lookup("term.amountrequired", base.CurrentUser.AdminUser.LanguageID);
                AwayLabel = PhraseLib.Lookup("term.amountaway", base.CurrentUser.AdminUser.LanguageID);
                TagLabel = PhraseLib.Lookup("term.amtrequired", base.CurrentUser.AdminUser.LanguageID);
                break;

            case (int)CPEUnitTypes.Volume:
                unitTypeInformation = m_LocalizationService.LoadQuantityInfoForOffer(RewardOptionID, CurrentUser.AdminUser.LanguageID, (int)CPEUnitTypes.Volume);
                UnitName = QuantityName[0];
                UnitAbbr = unitTypeInformation.Abbrevation;
                UnitPrecision = CurrencyPrecision = unitTypeInformation.Precision;
                UnitSymbol = "";
                RequiredLabel = PhraseLib.Lookup("term.quantityrequired", base.CurrentUser.AdminUser.LanguageID);
                AwayLabel = PhraseLib.Lookup("term.quantityaway", base.CurrentUser.AdminUser.LanguageID);
                TagLabel = PhraseLib.Lookup("term.qtyrequired", base.CurrentUser.AdminUser.LanguageID);
                break;
            case (int)CPEUnitTypes.Weight:
                unitTypeInformation = m_LocalizationService.LoadQuantityInfoForOffer(RewardOptionID, CurrentUser.AdminUser.LanguageID, (int)CPEUnitTypes.Weight);
                UnitName = QuantityName[0];
                UnitAbbr = unitTypeInformation.Abbrevation;
                UnitPrecision = CurrencyPrecision = unitTypeInformation.Precision;
                UnitSymbol = "";
                RequiredLabel = PhraseLib.Lookup("term.quantityrequired", base.CurrentUser.AdminUser.LanguageID);
                AwayLabel = PhraseLib.Lookup("term.quantityaway", base.CurrentUser.AdminUser.LanguageID);
                TagLabel = PhraseLib.Lookup("term.qtyrequired", base.CurrentUser.AdminUser.LanguageID);
                break;

            case (int)CPEUnitTypes.Length:
                unitTypeInformation = m_LocalizationService.LoadQuantityInfoForOffer(RewardOptionID, CurrentUser.AdminUser.LanguageID, (int)CPEUnitTypes.Length);
                UnitName = QuantityName[0];
                UnitAbbr = unitTypeInformation.Abbrevation;
                UnitPrecision = CurrencyPrecision = unitTypeInformation.Precision;
                UnitSymbol = "";
                RequiredLabel = PhraseLib.Lookup("term.quantityrequired", base.CurrentUser.AdminUser.LanguageID);
                AwayLabel = PhraseLib.Lookup("term.quantityaway", base.CurrentUser.AdminUser.LanguageID);
                TagLabel = PhraseLib.Lookup("term.qtyrequired", base.CurrentUser.AdminUser.LanguageID);
                break;
            case (int)CPEUnitTypes.SurfaceArea:
                unitTypeInformation = m_LocalizationService.LoadQuantityInfoForOffer(RewardOptionID, CurrentUser.AdminUser.LanguageID, (int)CPEUnitTypes.SurfaceArea);
                UnitName = QuantityName[0];
                UnitAbbr = unitTypeInformation.Abbrevation;
                UnitPrecision = CurrencyPrecision = unitTypeInformation.Precision;
                UnitSymbol = "";
                RequiredLabel = PhraseLib.Lookup("term.quantityrequired", base.CurrentUser.AdminUser.LanguageID);
                AwayLabel = PhraseLib.Lookup("term.quantityaway", base.CurrentUser.AdminUser.LanguageID);
                TagLabel = PhraseLib.Lookup("term.qtyrequired", base.CurrentUser.AdminUser.LanguageID);
                break;

            case (int)CPEUnitTypes.WeightVolume:
                UnitName = QuantityName[0];
                UnitAbbr = PhraseLib.Lookup("term.lbsgals", CurrentUser.AdminUser.LanguageID);
                UnitPrecision = 3;
                UnitSymbol = "";
                RequiredLabel = PhraseLib.Lookup("term.quantityrequired", base.CurrentUser.AdminUser.LanguageID);
                AwayLabel = PhraseLib.Lookup("term.quantityaway", base.CurrentUser.AdminUser.LanguageID);
                TagLabel = PhraseLib.Lookup("term.qtyrequired", base.CurrentUser.AdminUser.LanguageID);
                break;

            case (int)CPEUnitTypes.Points:
                UnitName = QuantityName[0];
                UnitAbbr = PhraseLib.Lookup("term.points", CurrentUser.AdminUser.LanguageID);
                UnitPrecision = 0;
                UnitSymbol = "";
                RequiredLabel = PhraseLib.Lookup("term.quantityrequired", base.CurrentUser.AdminUser.LanguageID);
                AwayLabel = PhraseLib.Lookup("term.quantityaway", base.CurrentUser.AdminUser.LanguageID);
                TagLabel = PhraseLib.Lookup("term.qtyrequired", base.CurrentUser.AdminUser.LanguageID);
                break;
        }
        //Set hidden field value for javascript usage
        hdnPrecision.Value = UnitPrecision.ToString();
        switch (UnitPrecision)
        {
            case 0:
                UnitPrecisionFormat = "0";
                break;

            case 1:
                UnitPrecisionFormat = "0.0";
                break;

            case 2:
                UnitPrecisionFormat = "0.00";
                break;

            case 3:
                UnitPrecisionFormat = "0.000";
                break;

            case 4:
                UnitPrecisionFormat = "0.0000";
                break;

            case 5:
                UnitPrecisionFormat = "0.00000";
                break;
        }
    }

    /// <summary>
    /// Populates all the fields on the UI
    /// </summary>
    private void LoadDefaultData()
    {
        LoadBasicData();

        if (!IsPostBack)
            UpdateProximityReward();

        LoadTierData();
        LoadTierConditionRelatedData();
    }

    /// <summary>
    /// Populates dropdown based on the condition set in the offer.
    /// </summary>
    private void LoadBasicData()
    {
        DataTable table = new DataTable();
        table.Columns.Add("DataValue", typeof(int));
        table.Columns.Add("DataText", typeof(string));

        for (var i = 0; i < QuantityUnit.Count; i++)
        {
            if (PMID == 0)
            {
                if (PMCount == 0 && (Convert.ToInt32(QuantityValue[i].Split(',')[0]) != 1 ? false : true))
                    table.Rows.Add(QuantityUnit[i], QuantityName[i]);
                else if (PMCount == 1 && PreviousThresholdTypeId != QuantityUnit[i] && (Convert.ToInt32(QuantityValue[i].Split(',')[0]) != 1 ? false : true))
                    table.Rows.Add(QuantityUnit[i], QuantityName[i]);
            }
            else if (PMCount == 1 && (Convert.ToInt32(QuantityValue[i].Split(',')[0]) != 1 ? false : true))
                table.Rows.Add(QuantityUnit[i], QuantityName[i]);
            else if (ThresholdTypeId == QuantityUnit[i] && (Convert.ToInt32(QuantityValue[i].Split(',')[0]) != 1 ? false : true))
                table.Rows.Add(QuantityUnit[i], QuantityName[i]);
        }

        ddlMessageType.DataSource = table;
        ddlMessageType.DataValueField = "DataValue";
        ddlMessageType.DataTextField = "DataText";
        ddlMessageType.DataBind();
        ddlMessageType.SelectedValue = ThresholdTypeId.ToString();
        ed_normal.Text = TagLabel;
    }

    /// <summary>
    /// Updates Proximity Reward based on the number of tiers set.
    /// </summary>
    private void UpdateProximityReward()
    {
        if (objOffer.NumbersOfTier < objProximityMessage.ProximityMessageTiers.Count)
        {
            List<ProximityMessageTier> tempPMTier = new List<ProximityMessageTier>();
            for (int i = 0; i < objOffer.NumbersOfTier; i++)
            {
                tempPMTier.Add(objProximityMessage.ProximityMessageTiers[i]);
            }
            objProximityMessage.ProximityMessageTiers = tempPMTier;
            m_ProximityMessageRewardService.CreateUpdateProximityMessageReward(objProximityMessage, OfferID, 9, RewardOptionID);
        }
        else if (objOffer.NumbersOfTier > objProximityMessage.ProximityMessageTiers.Count)
        {
            int i = objProximityMessage.ProximityMessageTiers.Count;
            ProximityMessageTier tempProximityTierData;
            for (; i < objOffer.NumbersOfTier; i++)
            {
                tempProximityTierData = new ProximityMessageTier();
                tempProximityTierData.TierLevel = i + 1;
                tempProximityTierData.PMTierTranslations = new List<ProximityMessageTierTranslation>();
                tempProximityTierData.Message = "";
                tempProximityTierData.TriggerValue = Decimal.Parse("0");
                objProximityMessage.ProximityMessageTiers.Add(tempProximityTierData);
            }
        }
    }
    /// <summary>
    /// Populates Proximity Reward tier related data.
    /// </summary>
    private void LoadTierData()
    {
        foreach (var tier in objProximityMessage.ProximityMessageTiers)
        {

            tier.TriggerValue = m_CommonInc.Extract_Decimal(tier.TriggerValue.ToString(UnitPrecisionFormat), CurrentUser.AdminUser.Culture);
        }

        repProximityMsg.DataSource = objProximityMessage.ProximityMessageTiers;
        repProximityMsg.DataBind();
    }

    /// <summary>
    /// Updates product / points condition related data in each tiers including unit types.
    /// </summary>
    private void LoadTierConditionRelatedData()
    {
        int i = 0;
        int index = 0;
        if (QuantityUnit.Count == 2)
        {
            index = ThresholdTypeId == QuantityUnit[1] ? 1 : 0;
        }
        foreach (RepeaterItem pmItem in repProximityMsg.Items)
        {
            Label tempLabel = pmItem.FindControl("requiredLabel") as Label;
            tempLabel.Text = RequiredLabel;

            tempLabel = pmItem.FindControl("requiredSymbol") as Label;
            tempLabel.Text = UnitSymbol;

            TextBox tempTextBox = pmItem.FindControl("requiredData") as TextBox;
            if (Decimal.Parse(QuantityValue[index].Split(';')[i].Split(',')[1]).ToString(UnitPrecisionFormat, CultureInfo.InvariantCulture) != "0")
                tempTextBox.Text = Decimal.Parse(QuantityValue[index].Split(';')[i++].Split(',')[1]).ToString(UnitPrecisionFormat, CurrentUser.AdminUser.Culture);
            else
                tempTextBox.Text = "Undefined";

            tempLabel = pmItem.FindControl("requiredAbbr") as Label;
            tempLabel.Text = UnitAbbr;

            tempLabel = pmItem.FindControl("awayLabel") as Label;
            tempLabel.Text = AwayLabel;

            tempLabel = pmItem.FindControl("awaySymbol") as Label;
            tempLabel.Text = UnitSymbol;

            tempLabel = pmItem.FindControl("awayAbbr") as Label;
            tempLabel.Text = UnitAbbr;

        }
    }

    private void ApplyPermissions()
    {
        if (!UserEditFlag)
        {
            btnSave.Visible = false;
        }
    }

    /// <summary>
    /// Updates the Proximity Reward object and saves the data.
    /// </summary>
    private void UpdateObjectAndSave()
    {
        bool IsNewPMR = (objProximityMessage.Id == 0);

        objProximityMessage.ProximityMessageTiers.Clear();
        objProximityMessage.RewardOptionPhase = Phase;
        objProximityMessage.RewardID = IsNewPMR ? 0 : RewardID;
        objProximityMessage.Required = false;
        objProximityMessage.RewardTypeID = (int)DELIVERABLE_TYPES.PROXIMITY_MESSAGE;
        objProximityMessage.ThresholdTypeId = Int32.Parse(ddlMessageType.SelectedValue);

        ProximityMessageTier tempProximityTierData;

        foreach (RepeaterItem pmItem in repProximityMsg.Items)
        {
            tempProximityTierData = new ProximityMessageTier();
            TextBox tempAwayDataTextbox = pmItem.FindControl("awaydata") as TextBox;
            TextBox tempMessage = pmItem.FindControl("prmessage") as TextBox;

            tempProximityTierData.TriggerValue = tempAwayDataTextbox != null ? Decimal.Parse(tempAwayDataTextbox.Text, CultureInfo.InvariantCulture) : Decimal.Parse("0");
            tempProximityTierData.TierLevel = pmItem.ItemIndex + 1;
            if (IsMultiLanguageEnabled)
            {
                tempProximityTierData.PMTierTranslations = new List<ProximityMessageTierTranslation>();
                ProximityMessageTierTranslation tempPMTierTranslation;
                string tempData = "";
                for (int i = 0; i < Languages.Count; i++)
                {
                    tempPMTierTranslation = new ProximityMessageTierTranslation();
                    tempData = Request.Form["repProximityMsg$ctl0" + pmItem.ItemIndex + "$repProximityMsgDesc$ctl0" + i + "$prmessage"];
                    if (i == 0)
                    {
                        tempProximityTierData.Message = tempData;
                    }

                    if (tempData != "")
                    {
                        tempPMTierTranslation.ProximityMessageTierId = tempProximityTierData.Id;
                        tempPMTierTranslation.LanguageId = Languages[i].LanguageID;
                        tempPMTierTranslation.Message = tempData;
                        tempProximityTierData.PMTierTranslations.Add(tempPMTierTranslation);
                    }
                }
            }
            else
            {
                tempProximityTierData.Message = Request.Form["repProximityMsg$ctl0" + pmItem.ItemIndex + "$repProximityMsgDesc$ctl00$prmessage"];
                tempProximityTierData.PMTierTranslations = new List<ProximityMessageTierTranslation>();
            }

            objProximityMessage.ProximityMessageTiers.Add(tempProximityTierData);
        }

        AMSResult<bool> result = m_ProximityMessageRewardService.CreateUpdateProximityMessageReward(objProximityMessage, OfferID.ConvertToLong(), 9, RewardOptionID);
        RewardID = objProximityMessage.RewardID.ConvertToInt32();
        //save the templates permission 
        if (objOffer.IsTemplate)
        {
            //time to update the status bits for the templates
            int form_Disallow_Edit = 0;
            string[] LockFieldsList = null;

            form_Disallow_Edit = ucTemplateLockableFields.RewardLockStatus == true ? 1 : 0;
            if (ucTemplateLockableFields.ExceptionFields != null)
            {
                LockFieldsList = ucTemplateLockableFields.ExceptionFields.Select(x => x.ToString()).ToArray();
                m_Commondata.PurgeFieldLevelPermissions(AppName, objOffer.OfferID, RewardID);
                m_Commondata.UpdateDeliverableAndFieldLevelPermissions(RewardID, objOffer.OfferID, form_Disallow_Edit, LockFieldsList);
                LockFieldsList = ucTemplateLockableFields.RewardTemplateFieldSource.TemplateFieldList.Where(x => LockFieldsList.Contains(x.FieldId.ToString()) == false).Select(x => x.FieldId.ToString()).ToArray();
                m_Commondata.UpdateDeliverableAndUnlockFieldLevelPermissions(RewardID, objOffer.OfferID, form_Disallow_Edit == 1 ? 0 : 1, LockFieldsList);
            }
            else
            {
                m_Commondata.PurgeFieldLevelPermissions(AppName, objOffer.OfferID, RewardID);
            }
        }
        if (result.ResultType == AMSResultType.Success)
        {
            m_Commondata.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" + CurrentUser.AdminUser.ID + ", StatusFlag=1 where IncentiveID=" + OfferID + ";";
            m_Commondata.LRT_Execute();
            m_OAWService.ResetOfferApprovalStatus(OfferID);
            RegisterScript("Close", "window.close();");
        }
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


}

    #endregion
