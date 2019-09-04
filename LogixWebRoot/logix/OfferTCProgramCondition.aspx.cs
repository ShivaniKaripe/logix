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
using System.Data;

public partial class logix_OfferTCProgramCondition : AuthenticatedUI
{
  #region Variables
  ITrackableCouponConditionService m_TCProgramCondition;
  ITrackableCouponProgramService m_TCProgram;
  IOffer m_Offer;
  IActivityLogService m_ActivityLog;
    IOfferApprovalWorkflowService m_OAWService;

  bool IsTemplate = false;
  protected long OfferID = 0;
  bool FromTemplate = false;
  bool DisabledAttribute = false;
  long ConditionID = 0;
  int ConditionTypeID = 0;
  string historyString;
  Engines Engine;
  CMS.AMS.Models.Offer Offer;
  Copient.CommonInc MyCommon = new Copient.CommonInc();
  bool isTranslatedOffer = false;
  bool bEnableRestrictedAccessToUEOfferBuilder = false;
  bool bEnableAdditionalLockoutRestrictionsOnOffers =false;
  bool bOfferEditable = false;

  private const int TRACKABLE_COUPON_EXPIRE_DATE_SYSOPTION_ID = 325;
  private bool bExpireDateEnabled = false;

  #endregion

  #region Properties

  private TrackableCouponProgram IncludedTCProgram {

    get {
      return ViewState["IncludedTCProgram"] as TrackableCouponProgram;
    }
    set {
      ViewState["IncludedTCProgram"] = value;
    }

  }
  private TrackableCouponProgram SavedTCProgram {

    get {
      return ViewState["SavedTCProgram"] as TrackableCouponProgram;
    }
    set {
      ViewState["SavedTCProgram"] = value;
    }

  }
  private List<TrackableCouponProgram> AvailableFilteredTCProgram {

    get {
      return ViewState["AvailableFilteredTCProgram"] as List<TrackableCouponProgram>;
    }
    set {
      ViewState["AvailableFilteredTCProgram"] = value;
    }

  }
  private TCProgramCondition OfferTCProgramCondition {
    get {
      return ViewState["OfferTCProgramCondition"] as TCProgramCondition;
    }
    set {
      ViewState["OfferTCProgramCondition"] = value;
    }
  }
  #endregion

  #region Events
  protected void Page_Load(object sender, EventArgs e) {
    ResolveDepedencies();
    GetQueryStrings();
    LoadOfferSettings();
    if (!IsPostBack) {
      AssignPageTitle("term.offer","term.trackablecouponcondition",OfferID.ToString());
      SetUpAndLocalizePage();
      GetOfferTCProgramCondition();
      SetAvailableData(true);
      SetButtons();
      DisableControls();
      
    }
  }

  protected void btnSave_Click(object sender, EventArgs e) {
    try {
      if (!(lstSelected.Items.Count > 0)) {
        infobar.InnerText = PhraseLib.Lookup("CPE-rew-membership.selectpoints", LanguageID);
        infobar.Visible = true;
        return;
      }
      if (OfferTCProgramCondition == null) {
        OfferTCProgramCondition = new TCProgramCondition();
      }
      ConditionTypeID = m_TCProgramCondition.GetTCConditionTypeID(Engine);
      if (chkDisallow_Edit.Visible)
        OfferTCProgramCondition.DisallowEdit = chkDisallow_Edit.Checked;
      OfferTCProgramCondition.Deleted = false;
      OfferTCProgramCondition.ConditionID = ConditionID;
      OfferTCProgramCondition.ConditionTypeID = ConditionTypeID;
      OfferTCProgramCondition.EngineID = (int)Engine;
      OfferTCProgramCondition.RequiredFromTemplate = false;
      OfferTCProgramCondition.ProgramID = lstSelected.Items[0].Value.ConvertToLong();
      m_Offer.CreateUpdateOfferTrackableCouponCondition(OfferID, Engine, OfferTCProgramCondition);

      // Only update the program expiration when the expiration feature is off
      // or the program expiration type is offer end date
      if (!bExpireDateEnabled || (IncludedTCProgram.TCExpireType == 1))
      {
         m_TCProgram.UpdateTCProgramExpiryDate(OfferTCProgramCondition.ProgramID.ConvertToInt32(), Offer.EndDate);
      }

      m_Offer.UpdateOfferStatusToModified(OfferID, (int)Engine, CurrentUser.AdminUser.ID);
            m_OAWService.ResetOfferApprovalStatus(OfferID);
      historyString = PhraseLib.Lookup("history.CustomerTrackableCouponConditionEdit", LanguageID) + ":" + lstSelected.Items[0].Text;
      m_ActivityLog.Activity_Log(ActivityTypes.Offer, OfferID.ConvertToInt32(), CurrentUser.AdminUser.ID, historyString);
      ScriptManager.RegisterStartupScript(this, this.GetType(), "Close", "CloseModel()", true);
    }
    catch (Exception ex) {
      infobar.InnerText = ErrorHandler.ProcessError(ex);
      infobar.Visible = true;
    }
  }
  protected void select1_Click(object sender, EventArgs e) {
    if (lstAvailable.SelectedItem != null) {
      IncludedTCProgram = AvailableFilteredTCProgram.Where(p => p.ProgramID == lstAvailable.SelectedValue.ConvertToInt32()).SingleOrDefault();
      SetAvailableData();
    }
    SetButtons();
  }
  protected void deselect1_Click(object sender, EventArgs e) {
    if (lstSelected.SelectedItem != null) {
      string strFilter = functioninput.Text;
      IncludedTCProgram = null;
      SetAvailableData();
    }
    SetButtons();
  }
  protected void ReloadThePanel_Click(object sender, EventArgs e) {
    SetAvailableData(true);
  }
  #endregion

  #region Private Methods
  private void ResolveDepedencies() {
    m_TCProgramCondition = CurrentRequest.Resolver.Resolve<CMS.AMS.Contract.ITrackableCouponConditionService>();
    m_Offer = CurrentRequest.Resolver.Resolve<CMS.AMS.Contract.IOffer>();
    m_TCProgram = CurrentRequest.Resolver.Resolve<ITrackableCouponProgramService>();
    m_ActivityLog = CurrentRequest.Resolver.Resolve<IActivityLogService>();
        m_OAWService = CurrentRequest.Resolver.Resolve<IOfferApprovalWorkflowService>();
  }

  private void GetQueryStrings() {
    OfferID = Request.QueryString["OfferID"].ConvertToLong();
    ConditionID = Request.QueryString["ConditionID"].ConvertToLong();
    bEnableRestrictedAccessToUEOfferBuilder = MyCommon.Fetch_SystemOption(249) == "1" ? true : false;
    isTranslatedOffer = MyCommon.IsTranslatedUEOffer(Convert.ToInt32(OfferID), MyCommon);
    bEnableAdditionalLockoutRestrictionsOnOffers = MyCommon.Fetch_SystemOption(260) == "1" ? true : false;
    bOfferEditable = MyCommon.IsOfferEditablePastLockOutPeriod(CurrentUser.UserPermissions.EditOfferPastLockoutPeriod, MyCommon, Convert.ToInt32(OfferID));
    bExpireDateEnabled = SystemCacheData.GetSystemOption_General_ByOptionId(TRACKABLE_COUPON_EXPIRE_DATE_SYSOPTION_ID).Equals("1");
    
  }

  private void SetUpAndLocalizePage() {
    if (IsTemplate)
      title.InnerText = PhraseLib.Lookup("term.template", LanguageID) + " #" + OfferID + " " + PhraseLib.Lookup("term.trackablecouponcondition", LanguageID);
    else
      title.InnerText = PhraseLib.Lookup("term.offer", LanguageID) + " #" + OfferID + " " + PhraseLib.Lookup("term.trackablecouponcondition", LanguageID);
    btnSave.Text = PhraseLib.Lookup("term.save", LanguageID);
    select1.Text = "▼" + PhraseLib.Lookup("term.select", LanguageID);
    deselect1.Text = PhraseLib.Lookup("term.deselect", LanguageID) + "▲";

  }

  private void GetOfferTCProgramCondition() {
    if (ConditionID > 0) {
      OfferTCProgramCondition = m_TCProgramCondition.GetConditionByID(ConditionID);
    }

    if (OfferTCProgramCondition == null)
      OfferTCProgramCondition = new TCProgramCondition();
    else {
      SavedTCProgram = OfferTCProgramCondition.TCProgram;
      IncludedTCProgram = OfferTCProgramCondition.TCProgram;
      chkDisallow_Edit.Checked = OfferTCProgramCondition.DisallowEdit;
    }
  }

  private void LoadOfferSettings() {
    Offer = m_Offer.GetOffer(OfferID, LoadOfferOptions.None);
    Engine = (Engines)Offer.EngineID;
    IsTemplate = Offer.IsTemplate;
    FromTemplate = Offer.FromTemplate;
  }

  private void SetAvailableData(bool ReloadData = false) {
    try {
      if (ReloadData) {
        GetAllTCProgram();
        AddSavedTCProgram();
      }
      string strFilter = functioninput.Text;

      List<TrackableCouponProgram> filterlist;

      filterlist = AvailableFilteredTCProgram.ToList();
      if (IncludedTCProgram != null)
        filterlist = filterlist.Where(p => p.ProgramID != IncludedTCProgram.ProgramID).ToList();

      filterlist = filterlist.OrderBy(o => o.Name).ToList();

      List<TrackableCouponProgram> inc = new List<TrackableCouponProgram>();
      if (IncludedTCProgram != null)
        inc.Add(IncludedTCProgram);
      lstSelected.DataSource = inc;
      lstSelected.DataBind();

      lstAvailable.DataSource = filterlist;
      lstAvailable.DataBind();
    }
    catch (Exception ex) {
      infobar.InnerText = ErrorHandler.ProcessError(ex);
      infobar.Visible = true;
    }
  }

  private void AddSavedTCProgram() {

    string strFilter = functioninput.Text;
    if (SavedTCProgram != null) {
      if (functionradio1.Checked && SavedTCProgram.Name.StartsWith(strFilter, StringComparison.OrdinalIgnoreCase))
        AvailableFilteredTCProgram.Add(SavedTCProgram);
      else if (functionradio2.Checked && SavedTCProgram.Name.IndexOf(strFilter, StringComparison.OrdinalIgnoreCase) >= 0)
        AvailableFilteredTCProgram.Add(SavedTCProgram);
    }
  }

  private void SetButtons() {
    if (lstSelected.Items.Count == 0) {
      select1.Enabled = true;
      select1.Attributes["class"] = "regular select";
    }
    if (lstSelected.Items.Count > 0) {
      deselect1.Enabled = true;
      select1.Enabled = false;
      select1.Attributes["class"] = "regular";
      deselect1.Attributes["class"] = "regular deselect";
    }
    else {
      deselect1.Enabled = false;
      deselect1.Attributes["class"] = "regular";

    }
  }

  private void DisableControls() {
    if (!IsTemplate) {
      TempDisallow.Visible = false;
      DisabledAttribute = ((CurrentUser.UserPermissions.EditOffer && !(FromTemplate && OfferTCProgramCondition.DisallowEdit)) ? false : true);
    }
    else
      DisabledAttribute = CurrentUser.UserPermissions.EditTemplates ? false : true;
    //If disable is set to false, check Buyer conditions
    if (Offer.EngineID == 9 && !DisabledAttribute)
    {
        if (MyCommon.LRTadoConn.State == ConnectionState.Closed)
            MyCommon.Open_LogixRT();
        DisabledAttribute = ((CurrentUser.UserPermissions.EditOffersRegardlessBuyer || MyCommon.IsOfferCreatedWithUserAssociatedBuyer(CurrentUser.AdminUser.ID, OfferID)) ? false : true);
        MyCommon.Close_LogixRT();
    }

    if (DisabledAttribute) {
      functionradio1.Enabled = false;
      functionradio2.Enabled = false;
      functioninput.Enabled = false;
      lstAvailable.Enabled = false;
      lstSelected.Enabled = false;
      select1.Enabled = false;
      deselect1.Enabled = false;
      btnSave.Visible = false;
    }
    if ((bEnableRestrictedAccessToUEOfferBuilder && isTranslatedOffer) || (bEnableAdditionalLockoutRestrictionsOnOffers && !bOfferEditable))
    {
        btnSave.Visible = false;
    }
        if (m_OAWService.CheckIfOfferIsAwaitingApproval(OfferID).Result)
        {
            btnSave.Visible = false;
        }
    }

  private void GetAllTCProgram() {
    int RecordCount = 100;
    AMSResult<List<TrackableCouponProgram>> AllTCProgram = m_TCProgram.GetAvailableTrackableCouponPrograms(RecordCount, functioninput.Text, functionradio1.Checked);
    if (AllTCProgram.ResultType != AMSResultType.Success) {
      AvailableFilteredTCProgram = new List<TrackableCouponProgram>();
      infobar.Visible = true;
      infobar.InnerHtml = AllTCProgram.GetLocalizedMessage(LanguageID);
    }
    AvailableFilteredTCProgram = AllTCProgram.Result;
  }

  #endregion

}