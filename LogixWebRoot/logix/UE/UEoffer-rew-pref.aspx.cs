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
using System.Data;
using System.Data.SqlClient;

public partial class logix_UE_UEoffer_rew_pref : AuthenticatedUI
{
  #region Variables
  IOffer m_Offer;
  IActivityLogService m_ActivityLog;
  IPreferenceService m_Preference;
  IPreferenceRewardService m_PreferenceReward;
    IOfferApprovalWorkflowService m_OAWService;
  bool IsTemplate = false;
  protected long OfferID = 0;
  bool FromTemplate = false;
  bool DisabledAttribute = false;
  long RewardID = 0;
  Int64 DeliverableID = 0;
  Int32 PreferenceRewardID = 0;
  bool bEnableAdditionalLockoutRestrictionsOnOffers = false;
  bool bOfferEditable = false;
  Copient.CommonInc MyCommon = new Copient.CommonInc();
  CMS.AMS.Models.Offer Offer;
  private CMS.DB.IDBAccess m_dbAccess;
  private CMS.DB.SQLParametersList sqlParams = new CMS.DB.SQLParametersList();
  private Int32? NumTiers {
    get {
      return ViewState["NumTiers"] as Int32?;
    }
    set {
      ViewState["NumTiers"] = value;
    }
  }
  private List<PreferenceDataTypes> lstPreferenceDataTypes {
    get {
      return ViewState["PreferenceDataTypes"] as List<PreferenceDataTypes>;
    }
    set {
      ViewState["PreferenceDataTypes"] = value;
    }
  }
  private PreferenceReward OfferPreferenceReward {
    get {
      return ViewState["PreferenceReward"] as PreferenceReward;
    }
    set {
      ViewState["PreferenceReward"] = value;
    }
  }
  private List<Preference> AllPreferences {
    get {
      return ViewState["AllPreferences"] as List<Preference>;
    }
    set {
      ViewState["AllPreferences"] = value;
    }
  }
  private Preference SelectedPreference {
    get {
      return ViewState["SelectedPreference"] as Preference;
    }
    set {
      ViewState["SelectedPreference"] = value;
    }
  }
  private List<PreferenceTiers> PreferenceTierData {
    get {
      return ViewState["PreferenceTierData"] as List<PreferenceTiers>;
    }
    set {
      ViewState["PreferenceTierData"] = value;
    }
  }
  #endregion

  protected void Page_Load(object sender, EventArgs e) {
    ResolveDepedencies();
    GetQueryStrings();
    LoadOfferSettings();
    if (!IsPostBack) {
      AssignPageTitle("term.offer", "term.preferencereward", OfferID.ToString());
      SetUpAndLocalizePage();
      GetAllPreferences();
      if (LoadPreferenceDataTypes() == false) return;
      GetOfferPreferenceReward();
      SetAvailableData();
      SetButtons();
      DisableControls();
    }
    else
    {
       DisplayErrorHide();
    }
  }
  protected void btnSave_Click(object sender, EventArgs e) {
    try {
      if (SelectedPreference == null) {
        DisplayError(PhraseLib.Lookup("ueoffer-con-pref.SelectPreference", LanguageID));
        return;
      }
      if (PreferenceTierData == null || PreferenceTierData.Count == 0) {
        return;
      }
      List<PreferenceTiers> lstInvalidTierValues = (from p in PreferenceTierData
                                                    where p.PreferenceItems.Count(pi => pi.Selected == true) == 0
                                                    select p).ToList();
      if (lstInvalidTierValues.Count > 0) {
        DisplayError(String.Format(PhraseLib.Lookup("ueoffer-con-pref.SupplyTierValue", LanguageID), lstInvalidTierValues[0].TierLevel));
        return;
      }

            /* if Preference allows mulitple values - should check min seleted values*/
            if (!CheckValidationListMinValue())
                return;

      if (OfferPreferenceReward != null) {
      }
      else {
        OfferPreferenceReward = new PreferenceReward();
      }
      string historyString = PhraseLib.Lookup((OfferPreferenceReward.PreferenceRewardID == 0 ? "history.rew-pref-create" : "history.rew-pref-edit"), LanguageID) + ":" + SelectedPreference.PhraseText;
      if (TempDisallow.Visible) OfferPreferenceReward.DisallowEdit = chkDisallow_Edit.Checked;
      OfferPreferenceReward.Deleted = false;
      OfferPreferenceReward.PreferenceID = (Int64)SelectedPreference.PreferenceID;
      OfferPreferenceReward.RewardTypeID = 15;
      OfferPreferenceReward.RewardOptionId = RewardID;
      List<String> TierValues;
      if (OfferPreferenceReward.PreferenceRewardTiers == null) {
        OfferPreferenceReward.PreferenceRewardTiers = new List<PreferenceRewardTier>();
      }
      for (int counter = OfferPreferenceReward.PreferenceRewardTiers.Count; counter < NumTiers; counter++) {
        PreferenceRewardTier rewardtier = new PreferenceRewardTier();
        rewardtier.TierLevel = counter + 1;
        rewardtier.PreferenceRewardID = OfferPreferenceReward.PreferenceRewardID;
        rewardtier.PreferenceRewardTierValues = new List<PreferenceRewardTierValue>();
        OfferPreferenceReward.PreferenceRewardTiers.Add(rewardtier);
      }
      PreferenceRewardTierValue rewardtierval;
      foreach (RepeaterItem item in rptTierValues.Items) {
        ListBox selecteditems = (ListBox)item.FindControl("lstSelectedPreference");
        TierValues = selecteditems.Items.Cast<ListItem>().Select(i => i.Value).ToList();
        PreferenceRewardTier rewardtier = OfferPreferenceReward.PreferenceRewardTiers.Where(p => p.TierLevel == (item.ItemIndex + 1)).FirstOrDefault();
        if (OfferPreferenceReward.PreferenceRewardID > 0) {
          rewardtier.PreferenceRewardTierValues.RemoveAll(x => !TierValues.Contains(x.PreferenceValue));
        }
        foreach (String preferencevalue in TierValues) {
          if (rewardtier.PreferenceRewardTierValues.Count(p => p.PreferenceValue == preferencevalue) == 0) {
            rewardtierval = new PreferenceRewardTierValue();
            rewardtierval.PreferenceRewardTierID = (Int64)rewardtier.PreferenceRewardTierID;
            rewardtierval.PreferenceValue = preferencevalue;
            rewardtier.PreferenceRewardTierValues.Add(rewardtierval);
          }
        }
      }
      AMSResult<bool> SavePreferenceReward = m_PreferenceReward.CreateUpdatePreferenceReward(OfferID, OfferPreferenceReward); 
      if (SavePreferenceReward.ResultType != AMSResultType.Success) {
        DisplayError(SavePreferenceReward.MessageString);
        return;
      }
      m_Offer.UpdateOfferStatusToModified(OfferID, (Int32)Engines.UE, CurrentUser.AdminUser.ID);
            m_OAWService.ResetOfferApprovalStatus(OfferID);
      m_ActivityLog.Activity_Log(ActivityTypes.Offer, OfferID.ConvertToInt32(), CurrentUser.AdminUser.ID, historyString);
      ScriptManager.RegisterStartupScript(this, this.GetType(), "Close", "CloseModel()", true);
    }
    catch (Exception ex) {
      DisplayError(ex);
    }
  }
  protected void select1_Click(object sender, EventArgs e) {
    if (lstAvailable.SelectedItem != null) {
      if (column2.Visible == false) {
        column2.Visible = true;
        UpdatePanelMain.Update();
      }
      SelectedPreference = AllPreferences.Where(p => p.PreferenceID == lstAvailable.SelectedValue.ConvertToInt32()).SingleOrDefault();
      lblDataType.Text = lstPreferenceDataTypes.Where(p => p.DataTypeID == (int)SelectedPreference.DataTypeID).FirstOrDefault().PhraseText;
      lblMultiValued.Text = (SelectedPreference.MultiValue ? PhraseLib.Lookup("term.yes", LanguageID) : PhraseLib.Lookup("term.no", LanguageID));
      SetAvailableData();
      DisplayTierValues();
    }
    SetButtons();
  }
  protected void ReloadThePanel_Click(object sender, EventArgs e) {
    SetAvailableData(true);
  }

    #region Private Methods

    private bool CheckValidationListMinValue()
    {
        bool result = true;
        if (SelectedPreference.MultiValue == true)
        {
            int minValue = m_Preference.GetPreferenceItemslistminvalue(Convert.ToInt32(SelectedPreference.PreferenceID));
            foreach (RepeaterItem item in rptTierValues.Items)
            {
                List<String> TierValues;
                ListBox selecteditems = (ListBox)item.FindControl("lstSelectedPreference");
                TierValues = selecteditems.Items.Cast<ListItem>().Select(itm => itm.Value).ToList();
                if (TierValues.Count < minValue)
                {
                    DisplayError(String.Format(PhraseLib.Lookup("ueoffer-con-pref.MinValue", LanguageID), minValue));
                    result = false;
                    break;
                }
            }
        }
        return result;
    }

    private void DisplayError(Exception ex) {
    DisplayError(ErrorHandler.ProcessError(ex));
  }
  private void DisplayError(String ErrorText) {
    infobar.InnerText = ErrorText;
    infobar.Visible = true;
  }

    private void DisplayErrorHide()
    {
        infobar.InnerText = "";
        infobar.Visible = false;
    }

    private void ResolveDepedencies() {
    CurrentRequest.Resolver.AppName = "UEoffer-rew-pref.aspx";
    m_Offer = CurrentRequest.Resolver.Resolve<IOffer>();
    m_Preference = CurrentRequest.Resolver.Resolve<IPreferenceService>();
    m_PreferenceReward = CurrentRequest.Resolver.Resolve<IPreferenceRewardService>();
    m_ActivityLog = CurrentRequest.Resolver.Resolve<IActivityLogService>();
    m_dbAccess = CurrentRequest.Resolver.Resolve<CMS.DB.IDBAccess>();
        m_OAWService = CurrentRequest.Resolver.Resolve<IOfferApprovalWorkflowService>();
  }
  private void GetQueryStrings() {
    OfferID = Request.QueryString["OfferID"].ConvertToLong();
    RewardID = Request.QueryString["RewardID"].ConvertToLong();
    DeliverableID = Request.QueryString["DeliverableID"].ConvertToLong();
    PreferenceRewardID = Request.QueryString["PreferenceRewardID"].ConvertToInt32();
    bOfferEditable = MyCommon.IsOfferEditablePastLockOutPeriod(CurrentUser.UserPermissions.EditOfferPastLockoutPeriod, MyCommon, Convert.ToInt32(OfferID));
    bEnableAdditionalLockoutRestrictionsOnOffers = MyCommon.Fetch_SystemOption(260) == "1" ? true : false;
  }
  private void DisplayTierValues(bool ProcessSavedData = false) {
    if (SelectedPreference == null)
      return;
    PreferenceTierData = new List<PreferenceTiers>();
    AMSResult<List<PreferenceItems>> preferenceitems = m_Preference.GetPreferenceItemsbyPreferenceID(SelectedPreference.DataTypeID, (Int32)SelectedPreference.PreferenceID, LanguageID);
    if (preferenceitems.ResultType != AMSResultType.Success) {
      DisplayError(preferenceitems.MessageString);
      return;
    }
    List<PreferenceItems> preferenceValues = preferenceitems.Result;
    PreferenceTiers prefTier;
    for (byte counter = 1; counter <= NumTiers; counter++) {
      prefTier = new PreferenceTiers();
      prefTier.TierLevel = counter;
      prefTier.PreferenceItems = preferenceValues.Clone();
      PreferenceTierData.Add(prefTier);
      if (ProcessSavedData && OfferPreferenceReward != null && OfferPreferenceReward.PreferenceRewardTiers.Count >= counter) {
        List<String> lstPreference = (from p in OfferPreferenceReward.PreferenceRewardTiers[counter - 1].PreferenceRewardTierValues
                                      select p.PreferenceValue).ToList();
        List<PreferenceItems> lstSavedPreference = PreferenceTierData[counter - 1].PreferenceItems.Where(w => lstPreference.Contains(w.Value)).ToList();
        for (int counter1 = 0; counter1 < lstSavedPreference.Count; counter1++) {
          lstSavedPreference[counter1].Selected = true;
        }
      }
    }
    rptTierValues.DataSource = PreferenceTierData;
    rptTierValues.DataBind();
  }
  private void SetUpAndLocalizePage() {
    if (IsTemplate)
      title.InnerText = PhraseLib.Lookup("term.template", LanguageID) + " #" + OfferID + " " + PhraseLib.Lookup("term.preferencereward", LanguageID).ToLower();
    else
      title.InnerText = PhraseLib.Lookup("term.offer", LanguageID) + " #" + OfferID + " " + PhraseLib.Lookup("term.preferencereward", LanguageID).ToLower();
    btnSave.Text = PhraseLib.Lookup("term.save", LanguageID);
    select1.Text = "▼" + PhraseLib.Lookup("term.select", LanguageID);
  }
  private void GetOfferPreferenceReward() {
    if (Offer == null) return;
    if (PreferenceRewardID > 0) {
      AMSResult<PreferenceReward> preferencereward = m_PreferenceReward.GetPreferenceRewardByID(PreferenceRewardID);
      if (preferencereward.ResultType != AMSResultType.Success) {
        DisplayError(preferencereward.MessageString);
        return;
      }
      OfferPreferenceReward = preferencereward.Result;
      SelectedPreference = AllPreferences.Where(p => p.PreferenceID == OfferPreferenceReward.PreferenceID).FirstOrDefault();
      if (SelectedPreference == null) {
        AMSResult<Preference> preference = m_Preference.GetPreferenceByID(OfferPreferenceReward.PreferenceID, LanguageID);
        if (preference.ResultType != AMSResultType.Success) {
          DisplayError(preference.MessageString);
          return;
        }
        SelectedPreference = preference.Result;
      }
      lblDataType.Text = lstPreferenceDataTypes.Where(p => p.DataTypeID == (int)SelectedPreference.DataTypeID).FirstOrDefault().PhraseText;
      lblMultiValued.Text = (SelectedPreference.MultiValue ? PhraseLib.Lookup("term.yes", LanguageID) : PhraseLib.Lookup("term.no", LanguageID));
      if (Offer != null && Offer.IsTemplate && OfferPreferenceReward != null) {
        chkDisallow_Edit.Checked = OfferPreferenceReward.DisallowEdit;
      }
      DisplayTierValues(true);
    }
    else {
      lblDataType.Text = PhraseLib.Lookup("term.none", LanguageID);
      lblMultiValued.Text = PhraseLib.Lookup("term.no", LanguageID);
      column2.Visible = false;
    }
  }

  private void LoadOfferSettings() {
    Offer = m_Offer.GetOffer(OfferID, LoadOfferOptions.None);
    IsTemplate = Offer.IsTemplate;
    FromTemplate = Offer.FromTemplate;
    NumTiers = Offer.NumbersOfTier;
  }

  private void SetAvailableData(bool ReloadData = false) {
    try {
      if (ReloadData) {
        GetAllPreferences();
      }
      string strFilter = functioninput.Text;
      List<Preference> filterlist = new List<Preference>();
      if (SelectedPreference != null)
        filterlist = AllPreferences.Where(p => p.PreferenceID != SelectedPreference.PreferenceID).ToList();
      else
        filterlist = AllPreferences;
      List<Preference> inc = new List<Preference>();
      if (SelectedPreference != null) inc.Add(SelectedPreference);
      lstSelected.DataSource = inc;
      lstSelected.DataBind();

      lstAvailable.DataSource = filterlist;
      lstAvailable.DataBind();

      if (lstAvailable.Items.Count == 1) {
        lstAvailable.Items[0].Selected = true;
      }
    }
    catch (Exception ex) {
      DisplayError(ex);
    }
  }
  private void SetButtons() {
    if (lstSelected.Items.Count == 0) {
      select1.Attributes["class"] = "regular select";
    }
    else {
      select1.Attributes["class"] = "regular";
    }
  }

  private void DisableControls() {
    if (!IsTemplate) {
      TempDisallow.Visible = false;
      DisabledAttribute = ((CurrentUser.UserPermissions.EditOffer && !(FromTemplate && (OfferPreferenceReward == null ? false : OfferPreferenceReward.DisallowEdit))) ? false : true);
    }
    else
      DisabledAttribute = CurrentUser.UserPermissions.EditTemplates ? false : true;
    if (DisabledAttribute) {
      functionradio1.Enabled = false;
      functionradio2.Enabled = false;
      functioninput.Enabled = false;
      lstAvailable.Enabled = false;
      lstSelected.Enabled = false;
      select1.Enabled = false;
      btnSave.Visible = false;
      foreach (RepeaterItem item in rptTierValues.Items) {
        if (item.ItemType == ListItemType.Item || item.ItemType == ListItemType.AlternatingItem) {
          ((DropDownList)item.FindControl("ddlPreferenceData")).Enabled = false;
          ((ListBox)item.FindControl("lstSelectedPreference")).Enabled = false;
          ((Button)item.FindControl("btnAdd")).Enabled = false;
          ((Button)item.FindControl("btnRemove")).Enabled = false;
        }
      }
    }
    if (bEnableAdditionalLockoutRestrictionsOnOffers && !bOfferEditable)
    {
        btnSave.Visible = false;
    }
  }
  private void GetAllPreferences() {
    Int32 PageSize = 100;
    AMSResult<List<Preference>> lstPreference = m_Preference.GetPreferences(PageSize, functioninput.Text, functionradio1.Checked, LanguageID);
    if (lstPreference.ResultType != AMSResultType.Success) {
      DisplayError(lstPreference.MessageString);
      return;
    }
    AllPreferences = lstPreference.Result;
  }
  private bool LoadPreferenceDataTypes() {
    AMSResult<List<PreferenceDataTypes>> preferencedatatypes = m_Preference.GetPreferenceDataTypes(LanguageID);
    if (preferencedatatypes.ResultType != AMSResultType.Success) {
      DisplayError(preferencedatatypes.MessageString);
      return false;
    }
    lstPreferenceDataTypes = preferencedatatypes.Result;
    return true;
  }
  #endregion
  protected void rptTierValues_ItemDataBound(object sender, RepeaterItemEventArgs e) {
      if (e.Item.ItemType == ListItemType.Item || e.Item.ItemType == ListItemType.AlternatingItem)
      {
    Int32 TierLevel = e.Item.ItemIndex;
    ((Button)e.Item.FindControl("btnAdd")).Text = PhraseLib.Lookup("term.add", LanguageID);
    ((Button)e.Item.FindControl("btnRemove")).Text = PhraseLib.Lookup("term.remove", LanguageID);

      if (NumTiers != null && NumTiers > 1)
        ((Label)e.Item.FindControl("lblTierText")).Text = "<b>" + PhraseLib.Lookup("term.tier", LanguageID) + " " + ((PreferenceTiers)(e.Item.DataItem)).TierLevel + "</b><br/>";
      ((DropDownList)e.Item.FindControl("ddlPreferenceData")).DataSource = ((PreferenceTiers)(e.Item.DataItem)).PreferenceItems.Where(p => p.Selected == false);
      ((DropDownList)e.Item.FindControl("ddlPreferenceData")).DataBind();
      ((ListBox)e.Item.FindControl("lstSelectedPreference")).DataSource = ((PreferenceTiers)(e.Item.DataItem)).PreferenceItems.Where(p => p.Selected == true);
      ((ListBox)e.Item.FindControl("lstSelectedPreference")).DataBind();
      if ((SelectedPreference.DataTypeID == PreferenceDataType.ListBox && SelectedPreference.MultiValue == true && ((PreferenceTiers)(e.Item.DataItem)).PreferenceItems.Where(p => p.Selected == false).ToList().Count == 0)
        || (SelectedPreference.DataTypeID == PreferenceDataType.ListBox && SelectedPreference.MultiValue == false && ((PreferenceTiers)(e.Item.DataItem)).PreferenceItems.Where(p => p.Selected == true).ToList().Count > 0)
        || (SelectedPreference.DataTypeID != PreferenceDataType.ListBox && ((PreferenceTiers)(e.Item.DataItem)).PreferenceItems.Where(p => p.Selected == true).ToList().Count > 0)) {
        ((Button)e.Item.FindControl("btnAdd")).Enabled = false;
      }

            UpdatePanelMain.Update();
        }
  }
  protected void rptTierValues_ItemCommand(object source, RepeaterCommandEventArgs e) {
    if (SelectedPreference == null)
      return;
    if (e.CommandName.ToString() == "Add" && ((DropDownList)e.Item.FindControl("ddlPreferenceData")).SelectedItem != null) {
      Byte TierLevel = e.CommandArgument.ConvertToByte();
      List<PreferenceItems> lstPreferences = (from item in PreferenceTierData
                                              where item.TierLevel == TierLevel
                                              select item.PreferenceItems).FirstOrDefault().ToList();
      if (lstPreferences != null && lstPreferences.Count > 0) {
        PreferenceItems prefitem = (from p in lstPreferences
                                    where p.Value == ((DropDownList)e.Item.FindControl("ddlPreferenceData")).SelectedItem.Value
                                    select p).FirstOrDefault();
        if (prefitem != null)
          prefitem.Selected = true;
      }
      ((DropDownList)e.Item.FindControl("ddlPreferenceData")).DataSource = lstPreferences.Where(p => p.Selected == false);
      ((DropDownList)e.Item.FindControl("ddlPreferenceData")).DataBind();
      ((ListBox)e.Item.FindControl("lstSelectedPreference")).DataSource = lstPreferences.Where(p => p.Selected == true);
      ((ListBox)e.Item.FindControl("lstSelectedPreference")).DataBind();
      if (SelectedPreference.DataTypeID == PreferenceDataType.ListBox) {
        ((Button)e.Item.FindControl("btnAdd")).Enabled = ((SelectedPreference.MultiValue == false || lstPreferences.Where(p => p.Selected == false).ToList().Count == 0) ? false : true);
        if (SelectedPreference.MultiValue == true)
        {
            int i = m_Preference.GetPreferenceItemslistmaxvalue(Convert.ToInt32(SelectedPreference.PreferenceID));
            if (lstPreferences.Where(p => p.Selected == true).ToList().Count < i)
            {
                ((Button)e.Item.FindControl("btnAdd")).Enabled = true;
            }
            else
            {
                ((Button)e.Item.FindControl("btnAdd")).Enabled = false;
            }
        }
      }
      else {
        ((Button)e.Item.FindControl("btnAdd")).Enabled = false;
      }
    }
    else if (e.CommandName.ToString() == "Remove" && ((ListBox)e.Item.FindControl("lstSelectedPreference")).SelectedItem != null) {
      Byte TierLevel = e.CommandArgument.ConvertToByte();
      List<PreferenceItems> lstPreferences = (from item in PreferenceTierData
                                              where item.TierLevel == TierLevel
                                              select item.PreferenceItems).FirstOrDefault().ToList();
      if (lstPreferences != null && lstPreferences.Count > 0) {
        foreach (ListItem item in ((ListBox)e.Item.FindControl("lstSelectedPreference")).Items) {
          PreferenceItems prefitem = (from p in lstPreferences
                                      where p.Value == item.Value
                                      select p).FirstOrDefault();
          if (prefitem != null) prefitem.Selected = !item.Selected;
        }
        ((DropDownList)e.Item.FindControl("ddlPreferenceData")).DataSource = lstPreferences.Where(p => p.Selected == false);
        ((DropDownList)e.Item.FindControl("ddlPreferenceData")).DataBind();
        ((ListBox)e.Item.FindControl("lstSelectedPreference")).DataSource = lstPreferences.Where(p => p.Selected == true);
        ((ListBox)e.Item.FindControl("lstSelectedPreference")).DataBind();
      }
      if ((SelectedPreference.DataTypeID == PreferenceDataType.ListBox && lstPreferences.Where(p => p.Selected == false).ToList().Count > 0) || SelectedPreference.DataTypeID != PreferenceDataType.ListBox) {
        ((Button)e.Item.FindControl("btnAdd")).Enabled = true;
      }
    }

        UpdatePanelMain.Update();
    }
}
[Serializable]
public class PreferenceTiers
{
  public Byte TierLevel { get; set; }
  public List<PreferenceItems> PreferenceItems { get; set; }
}