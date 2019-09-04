using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using CMS;
using CMS.AMS;
using CMS.AMS.Contract;
using CMS.AMS.Models;
using System.Threading;
using Microsoft.Practices.Unity;

public partial class attribute_pgbuilderconfig : AuthenticatedUI
{
  #region Variable Declaration
  private IActivityLogService m_ActivityLogService;
  private IAttributeService m_AttributeService;
  private IProductService m_ProductService;
  private CMS.AMS.Common m_Common;
  private AMSResult<List<HierarchyLevels>> amsResultHierarchyLevels;
  #endregion

  #region Override Methods
  protected override void AuthorisePage() {
    if (CurrentUser.UserPermissions.EditSystemConfiguration == false) {
      Server.Transfer("PageDenied.aspx?PhraseName=perm.admin-configuration&TabName=8_4", false);
      return;
    }
  }
  #endregion

  #region Events
  protected void Page_Load(object sender, EventArgs e) {
    infobar.InnerHtml = statusbar.InnerHtml = "";
    infobar.Visible = statusbar.Visible = false;
    (this.Master as logix_LogixMasterPage).Tab_Name = "8_4";
    AssignPageTitle("term.attributepg-builderconfiguration");
    SetPageData();
    ResolveDependencies();
    btnSave.Enabled = m_Common.Fetch_UE_SystemOption(157) == "1";
    if (!IsPostBack)
    {
      LoadData();
      SetButtons();
      amsResultHierarchyLevels = m_ProductService.GetProductHierarchyLevels();
      if (amsResultHierarchyLevels.ResultType != AMSResultType.Success)
        DisplayError(amsResultHierarchyLevels.MessageString);
      else
      {
        LoadProdHierarchyLevels();
        SetGroupingControls();
        ShowPreviousLevelUpdateStatus();
      }
    }
  }

  protected void btnSave_Click(object sender, EventArgs e) {
      bool saveHLSelections = false;
    if (lbSelectedAttributeTypes.Items.Count == 0 && lbAvailableAttributeTypes.Items.Count == 0) {
      return;
    }
    try {
      List<AttributeType> lstAttributes = new List<AttributeType>();
      AttributeType objAttribute;
      foreach (ListItem item in lbAvailableAttributeTypes.Items) {
        objAttribute = new AttributeType();
        objAttribute.AttributeTypeID = item.Value.ConvertToInt16();
        objAttribute.DisplayOrder = null;
        objAttribute.AttributeName = item.Text;
        lstAttributes.Add(objAttribute);
      }
      Byte counter = 1;
      foreach (ListItem item in lbSelectedAttributeTypes.Items) {
        objAttribute = new AttributeType();
        objAttribute.AttributeTypeID = item.Value.ConvertToInt16();
        objAttribute.DisplayOrder = counter;
        objAttribute.AttributeName = item.Text;
        lstAttributes.Add(objAttribute);
        counter++;
      }
      m_AttributeService.UpdateAttributeTypeDisplayOrder(lstAttributes);
      statusbar.Visible = true;
      statusbar.InnerHtml += PhraseLib.Lookup("term.changessaved", LanguageID);
      m_ActivityLogService.Activity_Log(ActivityTypes.AttributePGBuilderConfig, 0, CurrentUser.AdminUser.ID, "Modified Attribute Product Group Builder Configuration");
      SetButtons(true);

      //Hierarchy Level update
      DataTable dt = PrepareHierarchyLevelSelectionDT(radNoGrouping.Checked);
      if (dt != null && dt.Rows.Count > 0)
      {
        saveHLSelections = SaveGroupingSelection(dt);
        if (saveHLSelections)
        {
          //foreach(DataRow row in dt.Rows)
          TriggerLevelUpdateAsync(dt);
          infobar.Visible = true;
          infobar.Attributes["class"] = "modbar";
              infobar.Style["background-color"] = "#cc6000";
              infobar.InnerHtml += PhraseLib.Lookup("pab.config.updatinglevels", LanguageID);
          }

      }
    }
    catch (Exception ex) {
      DisplayError(ex.Message);
    }
  }

  protected void select1_Click(object sender, EventArgs e) {
    if (lbAvailableAttributeTypes.Items.Count == 0 || lbAvailableAttributeTypes.GetSelectedIndices().Length == 0)
      return;
    //btnSave.Enabled = true;
    int CurrentDisplayOrder = lbSelectedAttributeTypes.Items.Count;
    List<ListItem> selectedItems = (from item in lbAvailableAttributeTypes.Items.OfType<ListItem>()
                                    where item.Selected
                                    select item).ToList<ListItem>();
    foreach (ListItem item in selectedItems) {
      item.Selected = false;
      lbSelectedAttributeTypes.Items.Add(item);
      lbAvailableAttributeTypes.Items.Remove(item);
    }
    SetButtons();
  }
  protected void deselect1_Click(object sender, EventArgs e) {
    if (lbSelectedAttributeTypes.Items.Count == 0 || lbSelectedAttributeTypes.GetSelectedIndices().Length == 0)
      return;
    //btnSave.Enabled = true;
    List<ListItem> selectedItems = (from item in lbSelectedAttributeTypes.Items.OfType<ListItem>()
                                    where item.Selected
                                    select item).ToList<ListItem>();
    foreach (ListItem item in selectedItems) {
      item.Selected = false;
      lbAvailableAttributeTypes.Items.Add(item);
      lbSelectedAttributeTypes.Items.Remove(item);
    }
    SetButtons();
  }
  protected void btnMoveDown_Click(object sender, EventArgs e) {
    if (lbSelectedAttributeTypes.Items.Count == 0 || lbSelectedAttributeTypes.GetSelectedIndices().Length != 1)
      return;
    //btnSave.Enabled = true;
    int SelectedIndex = lbSelectedAttributeTypes.SelectedIndex;

    lbSelectedAttributeTypes.Items.Insert(SelectedIndex + 2, lbSelectedAttributeTypes.Items[SelectedIndex]);
    lbSelectedAttributeTypes.Items.RemoveAt(SelectedIndex);
    lbSelectedAttributeTypes.SelectedIndex = SelectedIndex + 1;
    SelectedIndex++;
    btnMoveUp.Enabled = (SelectedIndex > 0);
    btnMoveDown.Enabled = (SelectedIndex != (lbSelectedAttributeTypes.Items.Count - 1));
    lbSelectedAttributeTypes.SelectedIndex = SelectedIndex;
  }
  protected void btnMoveUp_Click(object sender, EventArgs e) {
    if (lbSelectedAttributeTypes.Items.Count == 0 || lbSelectedAttributeTypes.GetSelectedIndices().Length != 1)
      return;
    //btnSave.Enabled = true;

    int SelectedIndex = lbSelectedAttributeTypes.SelectedIndex;
    if (SelectedIndex == 0)
      return;
    lbSelectedAttributeTypes.Items.Insert(SelectedIndex - 1, lbSelectedAttributeTypes.Items[SelectedIndex]);
    lbSelectedAttributeTypes.Items.RemoveAt(SelectedIndex + 1);
    lbSelectedAttributeTypes.SelectedIndex = SelectedIndex - 1;
    SelectedIndex--;
    btnMoveUp.Enabled = (SelectedIndex != 0);
    btnMoveDown.Enabled = (SelectedIndex < (lbSelectedAttributeTypes.Items.Count - 1));
    lbSelectedAttributeTypes.Items[SelectedIndex].Selected = true;
  }
  protected void ddlLevels_DataBound(object sender, EventArgs e)
  {
    DropDownList ddl = sender as DropDownList;
    if (ddl.Items.Count > 0)
    {
      ddl.Items.Insert(0, PhraseLib.Lookup("pab.config.selectlevel", LanguageID));
      string extHierarchyID = (ddl.Parent.FindControl("lblHierarchyID") as Label).Text;
      SetDDLLevelSelection(ddl, extHierarchyID);
    }
  }
  #endregion

  #region Private Methods
  private void TriggerLevelUpdateAsync(DataTable dt)
  {
    ThreadPool.QueueUserWorkItem(new WaitCallback(LevelUpdateCallBack), dt);
  }
  private void LevelUpdateCallBack(object state)
  {
    //object[] objParams = state as object[];
    DataTable dtHL = state as DataTable;
    string extHierarchyID = string.Empty;
    Int16 groupingLevel = 0;

    ResolverBuilder requestResolver = new ResolverBuilder();

    try
    {
      requestResolver.Build();
      requestResolver.Container.RegisterType<IProductService, ProductService>(new HierarchicalLifetimeManager());
      IProductService ps = requestResolver.Container.Resolve<IProductService>();
      foreach (DataRow row in dtHL.Rows)
      {
        extHierarchyID = row["ExtHierarchyID"].ToString();
        groupingLevel = row["LevelInHierarchy"].ConvertToInt16();
        AMSResult<bool> amsResult = m_ProductService.UpdateHierarchyLevel(extHierarchyID, groupingLevel);
        if (amsResult.ResultType != AMSResultType.Success)
        {
          infobar.Visible = true;
          infobar.InnerHtml = PhraseLib.Lookup("pab.config.levelupdateerror", LanguageID) + ": " + amsResult.MessageString;
        }
      }
    }
    catch (Exception ex)
    {
      infobar.Visible = true;
      infobar.InnerHtml = PhraseLib.Lookup("pab.config.levelupdateerror", LanguageID) + ": " + ex.Message;
    }
  }
  private bool SaveGroupingSelection(DataTable dt)
  {
    bool hlSaveFlag = false;
    if (dt != null)
      hlSaveFlag = m_ProductService.SaveProductHierarchyGroupingLevel(dt).Result;
    return hlSaveFlag;
  }
  private DataTable PrepareHierarchyLevelSelectionDT(bool noGroupingFlag)
  {
    DataTable dt = new DataTable();
    string extHierarchyID = string.Empty;
    int levelInHierarchy = 0;
    string currentSelectedLevelName = string.Empty;
    int selectedIndexFromLevels = 0;
    DropDownList ddl = null;

    dt.Columns.Add("ExtHierarchyID");
    dt.Columns.Add("LevelInHierarchy", typeof(Int16));

    foreach (RepeaterItem ri in repGrouping.Items)
    {
      ddl = ri.FindControl("ddlLevels") as DropDownList;
      extHierarchyID = (ri.FindControl("lblHierarchyID") as Label).Text;
      if (noGroupingFlag)
      {
        dt.Rows.Add(extHierarchyID, null);
        ddl.SelectedIndex = 0;
      }
      else
      {
        currentSelectedLevelName = ddl.SelectedValue;
        //HierarchyLevels hl1 = ri.DataItem as HierarchyLevels;
        HierarchyLevels hl = m_ProductService.GetProductHierarchyLevel(extHierarchyID).Result;
        if (hl != null)
        {
          selectedIndexFromLevels = hl.ListLevelNames.IndexOf(currentSelectedLevelName);
          if (selectedIndexFromLevels >= 0)
          {
            levelInHierarchy = hl.ListLevelInHiearchy[selectedIndexFromLevels];
            //The level selected in UI should not be what is already saved to prevent redundant calls
            if (levelInHierarchy != hl.SelectedLevel)
              dt.Rows.Add(extHierarchyID, levelInHierarchy);
          }
        }
      }
    }
    return dt;
  }
  private void SetDDLLevelSelection(DropDownList ddl, string extHierarchyID)
  {
    if (amsResultHierarchyLevels != null && amsResultHierarchyLevels.Result.Count > 0)
    {
      HierarchyLevels hL = amsResultHierarchyLevels.Result.SingleOrDefault(p => p.ExtHierarchyID == extHierarchyID);
      int index = hL.ListLevelInHiearchy.IndexOf(hL.SelectedLevel);
      if (hL != null && hL.SelectedLevel > 0 && index != -1 && hL.ListLevelNames.Capacity >= index)
      {
        string savedLevelName = hL.ListLevelNames[index];
        int indexOfSavedLevel = ddl.Items.IndexOf(new ListItem(savedLevelName));
        if (indexOfSavedLevel != -1 && ddl.Items.Count >= indexOfSavedLevel)
          ddl.SelectedIndex = indexOfSavedLevel;
      }
    }
  }
  private void ResolveDependencies() {
    m_ActivityLogService = CurrentRequest.Resolver.Resolve<IActivityLogService>();
    m_AttributeService = CurrentRequest.Resolver.Resolve<AttributeService>();
    m_ProductService = CurrentRequest.Resolver.Resolve<ProductService>();
    m_Common = CurrentRequest.Resolver.Resolve<CMS.AMS.Common>();
  }
  private void SetPageData() {
    htitle.InnerText = PhraseLib.Lookup("term.attributepg-builderconfiguration", LanguageID);
    hselection.InnerText = PhraseLib.Lookup("term.datacolumnselection", LanguageID);
    btnSave.Text = PhraseLib.Lookup("term.save", LanguageID);
    btnMoveUp.Text = PhraseLib.Lookup("term.moveup", LanguageID) + " ▲";
    btnMoveDown.Text = PhraseLib.Lookup("term.movedown", LanguageID) + " ▼";
    select1.Text = PhraseLib.Lookup("term.select", LanguageID) + " ►";
    deselect1.Text = "◄ " + PhraseLib.Lookup("term.deselect", LanguageID);
    lblAvailableAttributeType.Text = PhraseLib.Lookup("term.attributetype-available", LanguageID);
    lblSelectedAttributeType.Text = PhraseLib.Lookup("term.attributetype-selected", LanguageID);
    //Grouping phrases
    hGrouping.InnerText = PhraseLib.Lookup("pabconfig.datagrouping", LanguageID);
    radGroupByLevel.Text = PhraseLib.Lookup("pab.config.grouping", LanguageID);
    radNoGrouping.Text = PhraseLib.Lookup("pab.config.nogrouping", LanguageID);
  }
  private void DisplayError(string errorString) {
    infobar.InnerHtml = errorString;
    infobar.Visible = true;
  }
  private void SetButtons(bool ToggleMoveButtons = false) {
    select1.Enabled = (lbAvailableAttributeTypes.Items.Count > 0);
    deselect1.Enabled = (lbSelectedAttributeTypes.Items.Count > 0);
    if (ToggleMoveButtons) {
      btnMoveUp.Enabled = (lbSelectedAttributeTypes.Items.Count > 1 &&
                           lbSelectedAttributeTypes.GetSelectedIndices().Length == 1 &&
                           lbSelectedAttributeTypes.SelectedIndex != 0);
      btnMoveDown.Enabled = (lbSelectedAttributeTypes.Items.Count > 1 &&
                             lbSelectedAttributeTypes.GetSelectedIndices().Length == 1 &&
                             lbSelectedAttributeTypes.SelectedIndex != (lbSelectedAttributeTypes.Items.Count - 1));
    }
    else {
      btnMoveUp.Enabled = btnMoveDown.Enabled = false;
    }
  }
  private void LoadData() {
    AMSResult<List<AttributeType>> AttributeTypes = m_AttributeService.GetAllAttributeTypes();
    if (AttributeTypes.ResultType != AMSResultType.Success) {
      DisplayError(AttributeTypes.MessageString);
      return;
    }
    SetAvailableData(AttributeTypes.Result);
  }
  private void SetAvailableData(List<AttributeType> lstAttributeTypes) {
    if (lstAttributeTypes != null && lstAttributeTypes.Count > 0) {
      lstAttributeTypes = lstAttributeTypes.OrderBy(x => x.DisplayOrder).ToList();
      List<AttributeType> AvailableAttributeTypes = lstAttributeTypes.Where(p => p.DisplayOrder == null).ToList();
      List<AttributeType> SelectedAttributeTypes = lstAttributeTypes.Where(p => p.DisplayOrder != null).ToList();

        // Adding below code  for Cloudsol 1284 , as HBC want to change ExtProductId As Item UPC"".we can't change this changes into database level as many other Query and Pages used it.
      foreach (AttributeType attributeType in AvailableAttributeTypes)
      {
        if (attributeType.AttributeName.Equals("ExtProductID", StringComparison.InvariantCultureIgnoreCase) && attributeType.AttributeTypeID.Equals(32767))
        {
          attributeType.AttributeName = PhraseLib.Lookup("term.itemsupc", LanguageID);  //"Item UPC";
        }
      }

      lbAvailableAttributeTypes.DataSource = AvailableAttributeTypes;
      lbAvailableAttributeTypes.DataBind();

      foreach (AttributeType attributeType in SelectedAttributeTypes)
      {
        if (attributeType.AttributeName.Equals("ExtProductID", StringComparison.InvariantCultureIgnoreCase) && attributeType.AttributeTypeID.Equals(32767))
        {
          attributeType.AttributeName = PhraseLib.Lookup("term.itemsupc", LanguageID);  //"Item UPC";
        }
      }

      lbSelectedAttributeTypes.DataSource = SelectedAttributeTypes;
      lbSelectedAttributeTypes.DataBind();
    }
  }
  private void LoadProdHierarchyLevels()
  {
    if (amsResultHierarchyLevels.Result != null && amsResultHierarchyLevels.Result.Count > 0)
    {
      repGrouping.DataSource = amsResultHierarchyLevels.Result;
      repGrouping.DataBind();
    }
  }
  private void SetGroupingControls()
  {
    if (amsResultHierarchyLevels.Result != null)
    {
      if (amsResultHierarchyLevels.Result.Count == 0)
      {
        radGroupByLevel.Enabled = false;
        radNoGrouping.Checked = true;
        SetNoGroupingMessage();
      }
      else
      {
        bool levelSelectionFlag = amsResultHierarchyLevels.Result.Exists(m => m.SelectedLevel > 0);

        if (levelSelectionFlag)
          radGroupByLevel.Checked = true;
        else
          radNoGrouping.Checked = true;
      }
    }
    else
    {
      radGroupByLevel.Enabled = false;
      radNoGrouping.Checked = true;
    }
  }
  private void SetNoGroupingMessage()
  {
    lblGroupingNotAvailable.Visible = true;
    lblGroupingNotAvailable.Text = PhraseLib.Lookup("pab.config.groupingnotavailable", LanguageID);
  }
  private void ShowPreviousLevelUpdateStatus()
  {
    if (amsResultHierarchyLevels.Result != null)
    {
      bool levelUpdateInProgress = amsResultHierarchyLevels.Result.Exists(m => m.LevelUpdateStatus == (byte)HierarchyLevelUpdateStatus.ReadyForUpdate);
      if (levelUpdateInProgress)
      {
        infobar.Visible = true;
        infobar.Attributes["class"] = "modbar";
        infobar.Style["background-color"] = "#cc6000";
        infobar.InnerHtml = PhraseLib.Lookup("pab.config.updatinglevels", LanguageID);
      }
    }
  }
  #endregion
}