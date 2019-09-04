using System;
using System.Collections.Generic;
using System.Web.UI;
using System.Web.UI.WebControls;
using CMS;
using CMS.AMS;
using CMS.AMS.Contract;
using CMS.AMS.Models;
using Copient;
using CMS.Models;
using System.Web.Services;
using System.ServiceModel.Web;

public partial class logix_UE_UEOffer_con_BrokerInstantWin : AuthenticatedUI
{
  #region global variables
  int Phase;
  IOffer m_Offer;
  ILocalizationService m_LocalizationService;
  IActivityLogService m_ActivityLogService;
  IInstantWinConditionService m_InstantWinService;
    IOfferApprovalWorkflowService m_OAWService;
  CommonInc m_CommonInc;
  LogixInc Logix;
  CMS.AMS.Common m_Commondata;
  ICacheData m_CacheData;

  public int LanguageID = 1;//Default English
  public long OfferID = 0;
  public bool DissallowEdit = false;
  public bool IsTemplate = false;
  public int InstantWinID = 0;

  public CMS.AMS.Models.Offer objOffer
  {
    get { return ViewState["Offer"] as CMS.AMS.Models.Offer; }
    set { ViewState["Offer"] = value; }
  }
  private List<CMS.AMS.Models.Language> lstLanguage
  {
    get { return ViewState["Language"] as List<CMS.AMS.Models.Language>; }
    set { ViewState["Language"] = value; }
  }
  private List<WebControl> pageUIControlsList
  {
    get { return ViewState["PageUIControls"] as List<WebControl>; }
    set { ViewState["PageUIControls"] = value; }
  }
  private InstantWinCondition objInstantWin
  {
    get { return ViewState["objInstantWin"] as InstantWinCondition; }
    set { ViewState["objInstantWin"] = value; }
  }
  private ICacheData CacheData
  {
    get { return ViewState["CacheData"] as CacheData; }
    set { ViewState["CacheData"] = value; }
  }
  private Int32 RoId
  {
    get { return ViewState["roid"].ConvertToInt32(); }
    set { ViewState["roid"] = value; }
  }
  #endregion

  #region PageMethods
  protected override void OnInit(EventArgs e)
  {
    AppName = "UEoffer-con-brokerinstantwin.aspx";
    base.OnInit(e);
  }

  protected void Page_Load(object sender, EventArgs e)
  {
    ResolveDependencies();
    GetQueryString();
    //LanguageID = SystemSettings.GetSystemDefaultLanguage().LanguageID;
    Image1.AlternateText = PhraseLib.Lookup("term.reload", LanguageID);
    Image1.ToolTip = PhraseLib.Lookup("term.reload", LanguageID);
    if (OfferID == 0)
    {
      DisplayError(PhraseLib.Lookup("error.invalidoffer", LanguageID));
      return;
    }

    if (objOffer == null)
    {
      objOffer = m_Offer.GetOffer(OfferID, CMS.AMS.Models.LoadOfferOptions.CustomerCondition);
      if (objOffer == null)
      {
        DisplayError(PhraseLib.Lookup("error.invalidoffer", LanguageID));
        return;
      }
      lstLanguage = SystemSettings.GetAllActiveLanguages((Engines)objOffer.EngineID);
      RoId = GetOfferRewardOptionID(OfferID);

    }

    if (!IsPostBack)
    {
      AMSResult<InstantWinCondition> result = m_InstantWinService.GetInstantWinCondition(RoId);
      if (result.ResultType == AMSResultType.Success)
      {
        objInstantWin = result.Result;
        if (objInstantWin == null)
        {
          objInstantWin = new InstantWinCondition();
          objInstantWin.ProgramType = InstantWinProgramType.Random;
        }
        hdnInstantWinID.Value = objInstantWin.IncentiveInstantWinID.ToString();
      }
      else
      {
        hdnInstantWinID.Value = "";
        DisplayError(result.GetLocalizedMessage<InstantWinCondition>(LanguageID));
        return;
      }
      GetInstantWinData();
      AssignPhrases();
      LoadIWUI();
    }
  }

  [WebMethod]
  public static String FetchWinners(String offerId, String storeNames)
  {
    var ajaxProcessingFunctions = new AjaxProcessingFunctions();
    ICacheData cache = CurrentRequest.Resolver.Resolve<ICacheData>();
    string PromoBrkrAddress = String.Format("http://{0}/ams-broker-promotion/instantwin/allwinners", cache.GetSystemOption_UE_ByOptionId(186));
    return ajaxProcessingFunctions.HttpPost(PromoBrkrAddress, offerId.ConvertToInt32(), storeNames);
  }

  private void LoadIWUI()
  {
    rbtnRandom.Checked = objInstantWin.ProgramType == InstantWinProgramType.Random ? true : false;
    rbtnSequence.Checked = objInstantWin.ProgramType == InstantWinProgramType.Sequential ? true : false;
    rbtnListChanceOfWin.SelectedValue = objInstantWin.ChanceOfWinningEnterprise ? "1" : "0";
    rbtnListAwardLimit.SelectedValue = objInstantWin.AwardLimitEnterprise ? "1" : "0";
    //Disable the program type section, if offer is deployed
    if (objInstantWin.IncentiveInstantWinID > 0)
    {
      bool Deployed = m_InstantWinService.IsInstantWinConditionDeployed(objInstantWin.IncentiveInstantWinID).Result;
      divIWPrgType.Enabled = !Deployed;
      divAwardLimitrbtnList.Enabled = !Deployed ;
      divchanceofWinList.Enabled = !Deployed;

    }
    if (!objOffer.IsTemplate)
      TempDisallow.Visible = false;
        if (m_OAWService.CheckIfOfferIsAwaitingApproval(OfferID).Result)
            btnSave.Visible = false;
  }
  #endregion

  #region UI Server events
  protected void btnSave_Click(object sender, EventArgs e)
  {
    SaveData();
  }
  #endregion

  #region PrivateMethods
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

  private void AssignPhrases()
  {
    AssignPageTitle("term.offer", "term.instantwincondition", OfferID.ToString());
    if (objOffer.IsTemplate)
      title.InnerText = PhraseLib.Lookup("term.template", LanguageID) + " #" + OfferID + " " + PhraseLib.Lookup("term.instantwincondition", LanguageID);
    else if (objOffer.FromTemplate)
      title.InnerText = PhraseLib.Lookup("term.offer", LanguageID) + " #" + OfferID + " " + PhraseLib.Lookup("term.instantwincondition", LanguageID);
    else
      title.InnerText = PhraseLib.Lookup("term.offer", LanguageID) + " #" + OfferID + " " + PhraseLib.Lookup("term.instantwincondition", LanguageID);

    rbtnRandom.Text = PhraseLib.Lookup("term.IWrandom", LanguageID);
    rbtnSequence.Text = PhraseLib.Lookup("term.IWSequence", LanguageID);

    rbtnListChanceOfWin.Items[0].Text = PhraseLib.Lookup("term.perstore", LanguageID);//Per store
    rbtnListChanceOfWin.Items[1].Text = PhraseLib.Lookup("term.IWallStores", LanguageID);//Across all the stores in this offer
    rbtnListChanceOfWin.Items[0].Value = "0";
    rbtnListChanceOfWin.Items[1].Value = "1";

    rbtnListAwardLimit.Items[0].Text = PhraseLib.Lookup("term.perstore", LanguageID);//Per store
    rbtnListAwardLimit.Items[1].Text = PhraseLib.Lookup("term.IWallStores", LanguageID);//Across all the stores in this offer
    rbtnListAwardLimit.Items[0].Value = "0";
    rbtnListAwardLimit.Items[1].Value = "1";

    rbtnUnlimited.Text = PhraseLib.Lookup("term.unlimitedWinners", LanguageID);//Unlimited number of winners
    rbtnLimited.Text = PhraseLib.Lookup("term.totalno", LanguageID);//Total of
    btnSave.Text = PhraseLib.Lookup("term.save", LanguageID);
  }

  private void ResolveDependencies()
  {
    m_InstantWinService = CurrentRequest.Resolver.Resolve<CMS.AMS.Contract.IInstantWinConditionService>();
    m_Offer = CurrentRequest.Resolver.Resolve<CMS.AMS.Contract.IOffer>();
    m_Commondata = CurrentRequest.Resolver.Resolve<CMS.AMS.Common>();
    m_CommonInc = CurrentRequest.Resolver.Resolve<CommonInc>();
    m_CacheData = CurrentRequest.Resolver.Resolve<CacheData>();
    m_LocalizationService = CurrentRequest.Resolver.Resolve<ILocalizationService>();
    m_ActivityLogService = CurrentRequest.Resolver.Resolve<IActivityLogService>();
        m_OAWService = CurrentRequest.Resolver.Resolve<IOfferApprovalWorkflowService>();

    m_Commondata = new CMS.AMS.Common(Environment.MachineName, "UEOffer_con_UEInstantWin.aspx");
    LanguageID = CurrentUser.AdminUser.LanguageID;
    m_Commondata.Open_LogixRT();
  }

  private void GetQueryString()
  {
    OfferID = Request.QueryString["OfferID"].ConvertToInt32();
    Phase = Request.QueryString["Phase"].ConvertToInt32();
    InstantWinID = Request.QueryString["InstantWinID"].ConvertToInt32();
    IsTemplate = ((Request.QueryString["IsTemplate"] != null) && (Request.QueryString["IsTemplate"] != "Not")) ? true : false;
  }

  private void DisplayError(string err)
  {
    infobar.Attributes["class"] = "red-background";
    infobar.InnerHtml = err;
    infobar.Visible = true;
    imghelp.ToolTip = err;
  }

  private void RegisterJavaScript(string key, string script)
  {
    // Get a ClientScriptManager reference from the Page class.
    ClientScriptManager cs = Page.ClientScript;

    // Check to see if the startup script is already registered.
    if (!cs.IsStartupScriptRegistered(this.GetType(), key))
    {
      cs.RegisterStartupScript(this.GetType(), key, script, true);
    }
  }

  private void GetInstantWinData()
  {
    if (objInstantWin != null)
    {
      hdnAwardLimitEnterprise.Value = objInstantWin.AwardLimitEnterprise.ToString();
      hdnChanceOfWinning.Value = objInstantWin.ChanceOfWinning.ToString();
      hdnChanceOfWinningEnterprise.Value = objInstantWin.ChanceOfWinningEnterprise.ToString();
      hdnDisallowEdit.Value = objInstantWin.DisallowEdit.ToString();
      hdnNumPrizesAllowed.Value = objInstantWin.NumPrizesAllowed.ToString();
      hdnUnlimited.Value = objInstantWin.Unlimited.ToString();
      hdnProgramType.Value = (objInstantWin.ProgramType == InstantWinProgramType.Random) ? "random" : "sequence";
      hdnDeleted.Value = objInstantWin.Deleted.ToString();

      hdnOfferID.Value = OfferID.ToString();
      hdnFromTemplate.Value = objOffer.FromTemplate.ToString();
      hdnIsTemplate.Value = objOffer.IsTemplate.ToString();
      string PromoBrkrIP = m_CacheData.GetSystemOption_UE_ByOptionId(186);
      hdnGetWinnersURL.Value = "http://" + PromoBrkrIP + "/ams-broker-promotion/instantwin/allwinners";

      StoreGroup StoresSelected = StoreGroup.UnKnown;
      var StoreInformation = m_Offer.GetCountOfUEOfferLocations(OfferID, ref StoresSelected).Result;
      hdnNoOfStores.Value = StoreInformation.Item1.ToString();
      hdnStores.Value = StoreInformation.Item2;

    }
  }

  private void SaveData()
  {
    //Library methods calls and processing
    if (objInstantWin == null)
      objInstantWin = new InstantWinCondition();
    objInstantWin.RewardOptionId = RoId;
    objInstantWin.ProgramType = (rbtnRandom.Checked) ? InstantWinProgramType.Random : InstantWinProgramType.Sequential;
    objInstantWin.ChanceOfWinningEnterprise = rbtnListChanceOfWin.SelectedValue == "1" ? true : false;
    objInstantWin.ChanceOfWinning = ((objInstantWin.ProgramType == InstantWinProgramType.Random) ? tboxRandomUsr.Text.Trim() : tboxSequenceUsr.Text.Trim()).ConvertToInt32();
    objInstantWin.NumPrizesAllowed = ((rbtnUnlimited.Checked) ? "" : tboxAwardLimitNumber.Text.Trim()).ConvertToInt32();
    if (rbtnUnlimited.Checked)
      objInstantWin.AwardLimitEnterprise = true;
    else
      objInstantWin.AwardLimitEnterprise = (rbtnListAwardLimit.SelectedValue == "1") ? true : false;

    objInstantWin.Unlimited = rbtnUnlimited.Checked;
    objInstantWin.DisallowEdit = chkDisallow_Edit.Checked;
    objInstantWin.Deleted = false;
    AMSResult<bool> result = m_InstantWinService.CreateUpdateInstantWinCondition(objInstantWin, OfferID, CurrentUser.AdminUser.ID);

    if (result.ResultType == AMSResultType.Success)
    {
            m_OAWService.ResetOfferApprovalStatus(OfferID);
      //success notification and close popup
      RegisterScript("Close", "window.close();");
    }
  }
  #endregion
}
