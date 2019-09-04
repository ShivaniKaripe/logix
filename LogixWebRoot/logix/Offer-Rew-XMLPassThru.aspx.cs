using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using CMS;
using CMS.AMS;
using CMS.AMS.Contract;
using CMS.AMS.Models;
using System.Xml;
using System.Data;
public partial class logix_Offer_Rew_XMLPassThru : AuthenticatedUI
{
    #region Variables
  IPassThroughRewards m_PassThroughRewards;
  IStoredValueProgramService m_SVProgram;
  IPointsProgramService m_PointsProgram;
    IOfferApprovalWorkflowService m_OAWService;
  IOffer m_Offer;
  bool DisabledAttribute = false;
  int OfferID;
  int DeliverableID;
  int Phase;
  bool isMultiLanguageEnabled = false;
  bool isBannerEnabled = false;
  int DefaultLanguageID = 0;
  bool IsNewXMLPassThru = false;
  string Description;
  Copient.CommonInc MyCommon = new Copient.CommonInc();
  protected string strOutOfRangeMessage = string.Empty;
   Int32 CustomerFacingLangID = 1;
   bool isTranslatedOffer = false;
   bool bEnableRestrictedAccessToUEOfferBuilder = false;
   bool bEnableAdditionalLockoutRestrictionsOnOffers = false;
   bool bOfferEditable = false;
  private PassThrough objPassThrough
  {
    get { return ViewState["XMLPassThrough"] as PassThrough; }
    set { ViewState["XMLPassThrough"] = value; }
  }
  protected CMS.AMS.Models.Offer objOffer
  {
    get { return ViewState["Offer"] as CMS.AMS.Models.Offer; }
    set { ViewState["Offer"] = value; }
  }
  private List<CMS.AMS.Models.Language> lstLanguage
  {
    get { return ViewState["Language"] as List<CMS.AMS.Models.Language>; }
    set { ViewState["Language"] = value; }
  }
  protected string StoredValueJSON { get; set; }
  protected string PointsJSON { get; set; }
  protected string CodeSettingsJSON { get; set; }

  protected override void OnInit(EventArgs e)
  {
    AppName = "Offer-Rew-XMLPassThru.aspx";
    base.OnInit(e);

  }

    #endregion Variables
    #region Protected Methods
  protected void Page_Load(object sender, EventArgs e)
  {

    ResolveDependencies();
    GetQueryString();
    Int32.TryParse(MyCommon.Fetch_SystemOption(125),out CustomerFacingLangID);
   
      

    if (objOffer == null)
    {
      if (OfferID == 0)
      {
        DisplayError("Invalid Offer ID");
        return;
      }
      objOffer = m_Offer.GetOffer(OfferID, CMS.AMS.Models.LoadOfferOptions.CustomerCondition);
      bEnableRestrictedAccessToUEOfferBuilder = MyCommon.Fetch_SystemOption(249) == "1" ? true : false;
      isTranslatedOffer = MyCommon.IsTranslatedUEOffer(OfferID,MyCommon);
      bEnableAdditionalLockoutRestrictionsOnOffers = MyCommon.Fetch_SystemOption(260) == "1" ? true : false;
      bOfferEditable = MyCommon.IsOfferEditablePastLockOutPeriod(CurrentUser.UserPermissions.EditOfferPastLockoutPeriod, MyCommon, OfferID);
    }
    SetUpAndLocalizePage();
    CheckPermission();
    if (!IsPostBack)
    {
      if (DeliverableID != 0)
      {
        AMSResult<PassThrough> result = m_PassThroughRewards.GetPassThroughReward(DeliverableID, objOffer.EngineID);
        if (result.ResultType == AMSResultType.Success)
        {
          objPassThrough = result.Result;
        }
        else
        {
          DisplayError(result.GetLocalizedMessage<PassThrough>(LanguageID));
          return;
        }

      }
      LoadPageData();

    }
    DisableControls();
  }

  protected void repXMLPassThroughData_ItemDataBound(object sender, RepeaterItemEventArgs e)
  {
    int TireLanguageID = 0;
    int TierLevel = 0;
    if (e.Item.ItemType == ListItemType.Item || e.Item.ItemType == ListItemType.AlternatingItem)
    {

      PassThroughTier tierdata = e.Item.DataItem as PassThroughTier;
      Label lblMessage = e.Item.FindControl("lblMessage") as Label;
      Label lblTierLevel = e.Item.FindControl("lblTierLevel") as Label;
      Label lblLanguageID = e.Item.FindControl("lblLanguageID") as Label;
      if (lblTierLevel != null)
        TierLevel = lblTierLevel.Text.ConvertToInt32();
      if (lblLanguageID != null)
        TireLanguageID = lblLanguageID.Text.ConvertToInt32();
      if (lblMessage != null)
      {

        if (TierLevel > 1)
          lblMessage.Style.Add("color", (TierLevel % 2 == 0 ? "#009900" : "#000099"));

        if (objOffer.NumbersOfTier > 1)
        {
          lblMessage.Text = "<b>" + PhraseLib.Lookup("term.tier", LanguageID) + " " + TierLevel + "</b>";
        }
        if (isMultiLanguageEnabled)
        {
          var lang = (from lan in lstLanguage
                      where lan.LanguageID == TireLanguageID
                      select lan).SingleOrDefault();

          lblMessage.Text = lblMessage.Text + " " + PhraseLib.Lookup(lang.PhraseTerm, LanguageID) + (lang.LanguageID == CustomerFacingLangID ? PhraseLib.Lookup("term.default", LanguageID) : "");
        }
      }

      TextBox txt = e.Item.FindControl("txtData") as TextBox;
      if (txt != null)
      {

        if (TierLevel == 1)
        {

          //txt.Rows = 10;

        }

      }

    }
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
    protected string GetRefreshScript()
    {
        string strval = string.Empty;
        switch (objOffer.EngineID)
        {

            case 3:
                strval = "opener.location='/logix/web-offer-rew.aspx?OfferID=" + objOffer.OfferID + "'; ";
                break;
            case 5:
                strval = "opener.location='/logix/email-offer-rew.asp?OfferID=" + objOffer.OfferID + "'; ";
                break;
            case 6:
                strval = "opener.location='/logix/CAM/CAM-offer-rew.aspx??OfferID=" + objOffer.OfferID + "'; ";
                break;
            case 9:
                strval = "opener.location='/logix/UE/UEoffer-rew.aspx?OfferID=" + objOffer.OfferID + "'; ";
                break;
            default:
                strval = "opener.location='/logix/CPEoffer-rew.aspx?OfferID=" + objOffer.OfferID + "'; ";
                break;

        }
        return strval;
    }

    #endregion Protected Methods

    #region Private Methods

  private void CreateObjectAndSave()
  {
    bool isValid = true;
    bool isTextEntered = false;
    string errMsg = String.Empty;
    objPassThrough.TiersData.Clear();
    objPassThrough.DisallowEdit = chkDisallow_Edit.Checked;
    objPassThrough.Required = chkRequiredToDeliver.Checked;
    foreach (RepeaterItem item in repXMLPassThroughData.Items)
    {
      Label lblTierLevel = item.FindControl("lblTierLevel") as Label;
      Label lblLanguageID = item.FindControl("lblLanguageID") as Label;
      TextBox txtData = item.FindControl("txtData") as TextBox;

      if (lblTierLevel != null && lblLanguageID != null && txtData != null)
      {

        string strXML = txtData.Text;
        if (!string.IsNullOrWhiteSpace(txtData.Text))
          isTextEntered = true;

        if (txtData.Text.Trim() != string.Empty && !ValidateXML(txtData.Text))
        {
          isValid = false;
          var lang = (from lan in lstLanguage
                      where lan.LanguageID == lblLanguageID.Text.ConvertToInt32()
                      select lan).SingleOrDefault();
          if (objOffer.NumbersOfTier > 1)
            errMsg = String.Format(PhraseLib.Lookup("term.invalidxmltier", LanguageID), lblTierLevel.Text, lang.Name);
          else
            errMsg = String.Format(PhraseLib.Lookup("term.invalidxml", LanguageID), lang.Name);

          break;
        }
        //validate XML and then add it, if invalid display the message and abort the save
        PassThroughTier objTierData = new PassThroughTier();
        objTierData.LanguageID = lblLanguageID.Text.ConvertToInt32();
        objTierData.TierLevel = lblTierLevel.Text.ConvertToInt32();
        objTierData.Data = txtData.Text.TrimEnd(Environment.NewLine.ToCharArray()).TrimStart(Environment.NewLine.ToCharArray()).Trim();

        //Update Values
        RepeaterItem valItem = repValues.Items[objTierData.TierLevel - 1];
        TextBox txtVal = valItem.FindControl("txtValue") as TextBox;
        if (txtVal != null)
          objTierData.Value = txtVal.Text.ConvertToDecimal();

        objPassThrough.TiersData.Add(objTierData);
      }
    }
    if (!isTextEntered)
    {
      DisplayError(PhraseLib.Lookup("term.enterxmlpassthru", LanguageID));
      return;

    }
    if (!isValid)
    {
      DisplayError(errMsg);
      return;
    }
    IsNewXMLPassThru = (objPassThrough.PassThroughID == 0);
    AMSResult<bool> result = m_PassThroughRewards.CreateUpdatePassThroughReward(objPassThrough, objOffer.OfferID, objOffer.EngineID);
    if (result.ResultType != AMSResultType.Success)
      DisplayError(result.GetLocalizedMessage<bool>(LanguageID));
    else
    {
      Description = IsNewXMLPassThru ? "reward.createxmlpassthru" : "reward.editxmlpassthru";
            m_OAWService.ResetOfferApprovalStatus(OfferID);
      WriteToActivityLog(PhraseLib.Lookup(Description, LanguageID));
      ScriptManager.RegisterStartupScript(this, this.GetType(), "Close", "window.close();", true);
    }
  }
  private void WriteToActivityLog(string Description)
  {
    if (MyCommon.LRTadoConn.State == ConnectionState.Closed)
      MyCommon.Open_LogixRT();
    MyCommon.Activity_Log(3, OfferID, CurrentUser.AdminUser.ID, Description);
    MyCommon.Close_LogixRT();
  }
  private bool ValidateXML(string data)
  {
    try
    {
      XmlDocument doc = new XmlDocument();
      doc.LoadXml(data);

      return true;
    }
    catch (Exception e)
    {
      return false;
    }

  }
  private void CheckPermission()
  {
    //if (CurrentUser.UserPermissions.AccessOffers = false && !objOffer.IsTemplate)
    //{
    //  Response.Redirect("~/logix/PopupDenied.aspx?BodyType=2&PermPhraseName=perm.offers-access&Title=" + Page.Title, false);

    //}
    //if (CurrentUser.UserPermissions.AccessTemplates = false && objOffer.IsTemplate)
    //{
    //  Server.Transfer("~/logix/PopupDenied.aspx?BodyType=2&PermPhraseName=perm.offers-access-templates&Title=" + Page.Title);


    //}


  }
  private void DisplayError(string err)
  {
    infobar.Attributes["class"] = "red-background";
    infobar.InnerHtml = err;
    infobar.Visible = true;
  }

  private void GetQueryString()
  {
    OfferID = Request.QueryString["OfferID"].ConvertToInt32();
    DeliverableID = Request.QueryString["DeliverableID"].ConvertToInt32();
    Phase = Request.QueryString["Phase"].ConvertToInt32();
  }
  
  private void LoadPageData()
  {
    if (objOffer.CustomerGroupConditions != null && objOffer.CustomerGroupConditions.IncludeCondition[0].CustomerGroupID == SystemCacheData.GetAnyCustomerGroup().CustomerGroupID)
    {
      if (objOffer.EngineID != (int)Engines.UE || !m_SVProgram.IsAnyCustomerSVProgramExist())
      {
        hdDisableSV.Value = "1";
        var data = from sv in new List<SVProgram>() select new { SVProgramID = sv.SVProgramID, ProgramName = sv.ProgramName };
        StoredValueJSON = data.ToJSON();
        lbSV.DataSource = data;
        lbSV.DataBind();
      }
      else
      {
        LoadSV();
      }
      if (objOffer.EngineID != (int)Engines.UE || !m_PointsProgram.IsAnyCustomerPointProgramExist())
      {
        hdDisablePoint.Value = "1"; 
        var data2 = from point in new List<PointsProgram>() select new { ProgramID = point.ProgramID, ProgramName = point.ProgramName };
        PointsJSON = data2.ToJSON();
        lbPoints.DataSource = data2;
        lbPoints.DataBind();
      }
      else
      {
        LoadPoints();
      }
    }
    else {
      LoadSV();
      LoadPoints();
    }    
    LoadCode();
    SetPassThroughForLoad();
  }
  private void LoadSV()
  {

    var svPrograms = objOffer.CustomerGroupConditions.IncludeCondition[0].CustomerGroupID == SystemCacheData.GetAnyCustomerGroup().CustomerGroupID ? m_SVProgram.GetStoredValueAllowAnyCustomerPrograms(false):m_SVProgram.GetStoredValuePrograms();
    var data = from sv in svPrograms select new { SVProgramID = sv.SVProgramID, ProgramName = sv.ProgramName };
 
    StoredValueJSON = data.ToJSON();
    lbSV.DataSource = data;
    lbSV.DataBind();

  }
  private void LoadCode()
  {
    var codesettings = SystemSettings.GetTriggerCodeSettings((Engines)objOffer.EngineID);
    CodeSettingsJSON = codesettings.ToJSON();
    decimal index = codesettings.RangeBegin;
    int counter = 1;
    if (codesettings.PadLength > 0)
      txtCode.MaxLength = codesettings.PadLength;
    string str = string.Empty;
    if (!(codesettings.RangeBegin == 0 && codesettings.RangeEnd == 0))
    {
      if (codesettings.RangeBegin != codesettings.RangeEnd)
        if (codesettings.RangeBegin > codesettings.RangeEnd)
          str = PhraseLib.Lookup("ueoffer-con-plu.InvalidRangeDefinition", LanguageID);
        else
          str = PhraseLib.Detokenize("ueoffer-con-plu.RangeBounds", codesettings.RangeBeginString, codesettings.RangeEndString);
      else
        str = PhraseLib.Detokenize("ueoffer-con-plu.RangeBegin", codesettings.RangeBeginString);

    }
    else
      str = PhraseLib.Lookup("ueoffer-con-plu.NoRange", LanguageID);


    if (codesettings.RangeLocked)
      str = str + " " + PhraseLib.Lookup("ueoffer-con-plu.OutOfRangeNotAccepted", LanguageID);
    else
      str = str + " " + PhraseLib.Lookup("ueoffer-con-plu.OutOfRangeAccepted", LanguageID);

    strOutOfRangeMessage = PhraseLib.Detokenize("ueoffer-con-plu.outofrangenotallowed", codesettings.RangeBeginString, codesettings.RangeEndString);

    lblDisplay.Text = str;
    while (counter <= 100 && index <= codesettings.RangeEnd)
    {
      lbCodes.Items.Add(new ListItem() { Value = index.ToString(), Text = index.ToString().PadLeft(codesettings.PadLength, codesettings.PadLetter) });
      counter++;
      index++;
    }

  }
  private void LoadPoints()
  {
    var pointPrograms = objOffer.CustomerGroupConditions.IncludeCondition[0].CustomerGroupID == SystemCacheData.GetAnyCustomerGroup().CustomerGroupID ? m_PointsProgram.GetAllowAnyCustomerPointsPrograms(): m_PointsProgram.GetPointsPrograms();
    var data = from point in pointPrograms select new { ProgramID = point.ProgramID, ProgramName = point.ProgramName };
    PointsJSON = data.ToJSON();
    lbPoints.DataSource = data;
    lbPoints.DataBind();

  }
  private void SetPassThroughForLoad()
  {
    if (objPassThrough == null)
    {
      objPassThrough = new CMS.AMS.Models.PassThrough();
      objPassThrough.TiersData = new List<CMS.AMS.Models.PassThroughTier>();
      objPassThrough.PassThroughRewardID = 0;
      objPassThrough.RewardID = DeliverableID;
      objPassThrough.RewardOptionPhase = Phase;
      objPassThrough.RewardTypeID = 12;
      objPassThrough.LSInterfaceID = 2;
      objPassThrough.Required = true;
      objPassThrough.ActionTypeID = 0;
      objPassThrough.RewardOptionPhase = Phase;

    }
    lstLanguage = SystemSettings.GetAllActiveLanguages((Engines) objOffer.EngineID);
    isMultiLanguageEnabled = SystemSettings.IsMultiLanguageEnabled();
    DefaultLanguageID = SystemSettings.GetSystemDefaultLanguage().LanguageID;
    //isBannerEnabled = SystemSettings.IsBannerEnabled();
    List<PassThroughTier> UpdatedTireData = new List<PassThroughTier>();
    for (int i = 1; i <= objOffer.NumbersOfTier; i++)
    {
      var tiredata = (from td in objPassThrough.TiersData
                      where td.TierLevel == i && lstLanguage.Exists(l => l.LanguageID == td.LanguageID)
                      select td).ToList();
      UpdatedTireData.AddRange(tiredata);
      var addeddata = (from lan in lstLanguage
                       where !tiredata.Exists(d => d.LanguageID == lan.LanguageID)
                       select new PassThroughTier() { LanguageID = lan.LanguageID, TierLevel = i }).ToList();
      UpdatedTireData.AddRange(addeddata);
    }
    var TierValues = (from td in UpdatedTireData
                      select new { TierLevel = td.TierLevel, Value = td.Value }).Distinct();
    objPassThrough.TiersData.Clear();

    repXMLPassThroughData.DataSource = UpdatedTireData;
    repXMLPassThroughData.DataBind();

    repValues.DataSource = TierValues;
    repValues.DataBind();
    chkRequiredToDeliver.Checked = objPassThrough.Required;
    chkDisallow_Edit.Checked = objPassThrough.DisallowEdit;

  }
  private void ResolveDependencies()
  {
    m_PassThroughRewards = CurrentRequest.Resolver.Resolve<CMS.AMS.Contract.IPassThroughRewards>();
    m_Offer = CurrentRequest.Resolver.Resolve<CMS.AMS.Contract.IOffer>();
    m_SVProgram = CurrentRequest.Resolver.Resolve<CMS.AMS.Contract.IStoredValueProgramService>();
    m_PointsProgram = CurrentRequest.Resolver.Resolve<CMS.AMS.Contract.IPointsProgramService>();
        m_OAWService = CurrentRequest.Resolver.Resolve<IOfferApprovalWorkflowService>();

  }
  private void DisableControls()
  {

    if (!objOffer.IsTemplate)
      DisabledAttribute = ((CurrentUser.UserPermissions.EditOffer && !(objOffer.FromTemplate && objPassThrough.DisallowEdit)) ? false : true);
    else
      DisabledAttribute = CurrentUser.UserPermissions.EditTemplates ? false : true;
    //If disable is set to false, check Buyer conditions
    if (objOffer.EngineID == 9 && !DisabledAttribute)
    {
        if (MyCommon.LRTadoConn.State == ConnectionState.Closed)
            MyCommon.Open_LogixRT();
        DisabledAttribute = ((CurrentUser.UserPermissions.EditOffersRegardlessBuyer || MyCommon.IsOfferCreatedWithUserAssociatedBuyer(CurrentUser.AdminUser.ID, OfferID)) ? false : true);
        MyCommon.Close_LogixRT();
    }

    if (DisabledAttribute)
    {
      btnSave.Visible = false;
    }
        if ((bEnableRestrictedAccessToUEOfferBuilder && isTranslatedOffer) || (bEnableAdditionalLockoutRestrictionsOnOffers && !bOfferEditable) || m_OAWService.CheckIfOfferIsAwaitingApproval(OfferID).Result)
    {
        btnSave.Visible = false;
    }
  }
  private void SetUpAndLocalizePage()
  {
    if (objOffer.IsTemplate)
      title.InnerText = PhraseLib.Lookup("term.template", LanguageID) + " #" + OfferID + " " + PhraseLib.Lookup("term.xmlpass-through", LanguageID);
    else
      title.InnerText = PhraseLib.Lookup("term.offer", LanguageID) + " #" + OfferID + " " + PhraseLib.Lookup("term.xmlpass-through", LanguageID);
    btnSave.Text = PhraseLib.Lookup("term.save", LanguageID);
    if (!objOffer.IsTemplate)
      TempDisallow.Visible = false;


    AssignPageTitle("term.offer", "term.xmlpass-throughreward", OfferID.ToString());
    if (objOffer.EngineID == (int)Engines.UE)
    {
      distribution.Visible = true;
    }
  }
    #endregion Private Methods
}