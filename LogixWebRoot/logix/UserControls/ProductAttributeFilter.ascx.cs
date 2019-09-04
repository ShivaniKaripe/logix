﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Script.Serialization;
using System.Web.UI.WebControls;
using CMS;
using CMS.AMS;
using CMS.AMS.Contract;
using CMS.AMS.Models;
using CMS.Contract;
using System.Web.UI;
using System.Reflection;
using Newtonsoft.Json;
using System.Web.UI.HtmlControls;

public partial class logix_UserControls_ProductAttributeFilter : System.Web.UI.UserControl
{
  #region Public Property
  DataTable dtFilterData = new DataTable();
  public string AppName { get; set; }
  AMSResult<List<AttributeType>> objAttributeType;
  public static int TotalCount;
  public int LanguageID { get; set; }
  public List<string> ProductIDInculded = new List<string>();
  public List<string> ProductIDExcluded = new List<string>();
  List<Product> exclList = new List<Product>();
  bool btnStatus = true;
  bool? _IsFilterUpdated = false;

  bool? _IsLevelFilterUpdated = false;

  public bool ReloadHierarchyTreeClicked
  {
    get
    {
      return (hdnIsEditHierarchyInProgress.Value == "1");
    }
    set
    {
      hdnIsEditHierarchyInProgress.Value = value ? "1" : "0";
    }
  }

  public string HierarchyTreeURL
  {
    get
    {
      return Session["hierarchytreeURL"] == null ? "" : Session["hierarchytreeURL"].ToString();
    }
    set
    {
      Session["hierarchytreeURL"] = value;
      LoadHierarchyTree();
    }
  }

  public string LocateHierarchyURL
  {
    get { return hdnLocateHierarchyURL.Value; }
  }
  public AMSResult<DataTable> GridViewCacheData
  {
    get
    {
      return Session["GridViewCacheProducts"] == null ? new AMSResult<DataTable>() : Session["GridViewCacheProducts"] as AMSResult<DataTable>;
    }
    set
    {
      Session["GridViewCacheProducts"] = value;
    }
  }

  public DataTable LevelGridViewDetails
  {
    get
    {
      if (ViewState["LevelGridViewDetails"] == null)
        ViewState["LevelGridViewDetails"] = prepareGridviewtable();
      return ViewState["LevelGridViewDetails"] as DataTable;
    }
    set
    {
      if (value != null && (value is DataTable) && (value as DataTable).Rows.Count == 0)
        value = prepareGridviewtable();
      ViewState["LevelGridViewDetails"] = value;
    }
  }

  public AMSResult<List<AttributeType>> AttributeTypes
  {
    get
    {
      return Session["AttributeType"] as AMSResult<List<AttributeType>>;
    }
    set
    {
      Session["AttributeType"] = value;
    }
  }

  public bool ClearFilters
  {
    get
    {
      return ViewState["ClearFilters"] == null ? false : ViewState["ClearFilters"].ConvertToBool();
    }
    set
    {
      ViewState["ClearFilters"] = value;
    }
  }
  public DataTable GridViewData
  {
    get
    {
      return Session["FilteredProductswithAttr"] as DataTable;
    }
    set
    {
      Session["FilteredProductswithAttr"] = value;
    }
  }
  public DataTable AllChildNodes
  {
    get
    {
      return Session["AllChildNodes"] as DataTable;
    }
    set
    {
      Session["AllChildNodes"] = value;
    }
  }

  public string CurrentNodes
  {
    get
    {
      return Session["CurrentNodes"].ToString();
    }
    set
    {
      Session["CurrentNodes"] = value;
    }
  }
  public bool ButtonStatus
  {
    get
    {
      return btnStatus;
    }
    set
    {
      btnStatus = value;
    }
  }
  public DataTable DTSelectedAttributeValuesViewState
  {
    get
    {
      return ViewState["dtFilters"] as DataTable;
    }
    set
    {
      ViewState["dtFilters"] = value;
    }
  }
  public DataTable DTAppliedAttributeValuesViewState
  {
    get
    {
      return ViewState["dtAppliedFilters"] as DataTable;
    }
    set
    {
      ViewState["dtAppliedFilters"] = value;
    }
  }
  public List<AttributeType> AttributeTypeViewState
  {
    get
    {
      return ViewState["AttributeType"] as List<AttributeType>;
    }
    set
    {
      ViewState["AttributeType"] = value;
    }
  }
  public List<AttributeType> AttributeTypeToDisplayViewState
  {
    get
    {
      return ViewState["AttributeTypeToDisplay"] as List<AttributeType>;
    }
    set
    {
      ViewState["AttributeTypeToDisplay"] = value;
    }
  }
  public bool IsFilterUpdated
  {
    get
    {
      _IsFilterUpdated = ViewState["FilterUpdated"] as bool?;
      _IsFilterUpdated = _IsFilterUpdated == null ? false : _IsFilterUpdated;
      return (_IsFilterUpdated.ConvertToBool());
    }
    set
    {
      ViewState["FilterUpdated"] = value;
    }
  }
  public bool IsLevelFilterUpdated
  {
    get
    {
      _IsLevelFilterUpdated = ViewState["LevelFilterUpdated"] as bool? ?? false;
      return (_IsLevelFilterUpdated.ConvertToBool());
    }
    set
    {
      ViewState["LevelFilterUpdated"] = value;
      if (value)
      {
        hfIsGridUpdated.Value = hfLevelInfoTable.Value = hfIncludedItems.Value = hfExcludedItems.Value = "";
        LevelGridViewDetails = null;
        hlShowDetails.Text = PhraseLib.Lookup("term.view_detail_pglist", LanguageID, "Phrase Not Found");
        hlShowDetails.Enabled = true;
      }
    }
  }

  public String _sortBy
  {
    get
    {
      return ViewState["SortBy"] as String;
    }
    set
    {
      ViewState["SortBy"] = value;
    }
  }
  public String _sortOrder
  {
    get
    {
      return ViewState["SortOrder"] as String;
    }
    set
    {
      ViewState["SortOrder"] = value;
    }
  }
  public int Count
  {
    get
    {
      return Convert.ToInt32(ViewState["Count"]);
    }
    set
    {
      ViewState["Count"] = value;
    }
  }

  public Int32 BuyerID { get; set; }
  public Boolean IsEditPermitted { get; set; }
  public ILogger m_Logger { get; set; }
  public IPhraseLib PhraseLib
  {
    get;
    set;
  }

  private string _SelectedNodeIDs = "";
  public string SelectedNodeIDs { get { return GetSelectedNodeIDs().Trim(','); } set { _SelectedNodeIDs = value; } }

  public string SelectedHierarchy
  {
    get
    {
      ViewState["SelectedHierarchy "] = m_ProductGroup.GetHierarchyName(SelectedNodeIDs).Result;
      return ViewState["SelectedHierarchy "].ToString();
    }
    set
    {
      ViewState["SelectedHierarchy "] = value;
    }
  }

  public bool SaveData
  {
    set
    {
      if (value)
        updateProductGroupwithNodesAndAttribute();
    }
  }

  private long _ProductGroupID = -1;
  public long ProductGroupID
  {
    set
    {
      if (_ProductGroupID != value)
      {
        _ProductGroupID = value;
        if (!string.IsNullOrEmpty(HierarchyTreeURL) && IsAttributeSwitch && IsPostBack)
        {
          hdndivphselectedtree.Value = "";
          //Update the ProductGroupID in QueryString
          var nameValues = HttpUtility.ParseQueryString(HierarchyTreeURL.Substring(HierarchyTreeURL.IndexOf("?") + 1));
          nameValues.Set("ProductGroupID", _ProductGroupID.ToString());
          HierarchyTreeURL = HierarchyTreeURL.Substring(0, HierarchyTreeURL.IndexOf("?") + 1) + nameValues.ToString();
        }
      }
    }
    get { return _ProductGroupID; }
  }
  public string PrevSelectedNodeIDs
  {
    get
    {
      return ViewState["PrevSelectedNodeIDs"] as String;
    }
    set
    {
      ViewState["PrevSelectedNodeIDs"] = value;
    }

  }
  public bool IsGroupGrid
  {
    get
    {
      if (ViewState["IsGroupGrid"] != null)
        return Convert.ToBoolean(ViewState["IsGroupGrid"]);
      else
        return false;
    }
    set
    {
      ViewState["IsGroupGrid"] = value;
    }
  }

  public string GroupedOn
  {
    get
    {
      if (ViewState["GroupedOn"] != null)
        return Convert.ToString(ViewState["GroupedOn"]);
      else
        return "";
    }
    set
    {
      ViewState["GroupedOn"] = value;
    }
  }

  public bool IsPGAttributeType { get; set; }
  public bool IsAttributeSwitch { get; set; }

  public string PABStage { get { return hdnPABStage.Value; } }
  #endregion

  #region Fields
  private const string NODE_PRODUCT_COUNT_FLAG = "FetchNodesProductCount";
  CMS.AMS.Common m_common;
  IOffer m_Offer;
  IErrorHandler m_ErrorHandler;
  IAttributeService m_Attribute;
  IProductService m_Product;
  IProductGroupService m_ProductGroup;
  public int PageSize = 50;
  public int PageIndex = 0;
  Image sortImage = new Image();
  DataTable dtUpdate;
  System.IO.TextWriter writer;
  #endregion
  #region Protected Methods
  protected void Page_Load(object sender, EventArgs e)
  {
    DataTable dt = new DataTable();
    producthierarchy.InnerHtml = "";
    ResolveDependencies();
    loadphrasesForTexts();
    if (IsPGAttributeType == false)
      return;
    SetContolsVisibility();

    ScriptManager.RegisterOnSubmitStatement(this, this.GetType(), "Key_AbortAsyncCountRequest", "AbortAsyncCountRequest()");
    if (PABStage == "1" && producthierarchy.InnerHtml == string.Empty && writer != null)   //This condition comes up when the Hierarchy is returned from asynchronous call
      producthierarchy.InnerHtml = writer.ToString();

    if (!IsPostBack || IsAttributeSwitch)
    {
      //clear the previous content of the grid and its cache
      ClearSelectionOnGrid();
      ddlAttributeType.Attributes.Add("onChange", "getClear()");
      PopulateAttributeTypeToDisplay();
      hdnProductGroupID.Value = ProductGroupID.ConvertToString();
      PopulateControlText();

      exclList = m_Product.GetExcludedProducts(ProductGroupID).Result;
      hdnExludedProductsCount.Value = exclList.Count.ToString();

      AMSResult<DataTable> objAttributeValues = m_Product.GetFilterDataByPGID(ProductGroupID);

      dt = ClearFilters == false ? (DataTable)objAttributeValues.Result : null;
      DataTable dtTemp = GenerateBlankDTForAttributes();
      if (dt != null)
      {
        foreach (DataRow dRow in dt.Rows)
        {
          dtTemp.Rows.Add(dRow["AttributeSetID"], dRow["ProductAttributeTypeID"].ConvertToInt16(), dRow["AttributeName"].ToString(),
              dRow["ProductAttributeValueID"].ConvertToInt32(), dRow["AttributeValue"].ToString(), dRow["ExcludeFlag"]);
        }
        DTAppliedAttributeValuesViewState = dtTemp;
        //Store PAB applied attribute type\values to use in fetching count during page load on client side
        if (DTAppliedAttributeValuesViewState.Rows.Count == 0)
          hdnPABAVPairsJson.Value = NODE_PRODUCT_COUNT_FLAG;
        else
          hdnPABAVPairsJson.Value = PrepareJSON(DTAppliedAttributeValuesViewState.DefaultView.ToTable(false, "AttributeSetID", "AttributeTypeID", "AttributeValueID", "ExcludeFlag"));

        BindAttributeSetDDL();
        DTSelectedAttributeValuesViewState = GenerateBlankDTForAttributes();
        if (DTSelectedAttributeValuesViewState.Rows.Count > 0)
        {
          repFilter.DataSource = GetDataSourceForSelectedSetUI(DTSelectedAttributeValuesViewState);
          repFilter.DataBind();
        }
        if (DTAppliedAttributeValuesViewState.Rows.Count > 0)
        {
          repWrapMultiSet.DataSource = GetDataSourceForMultiAttributeSetBinding(DTAppliedAttributeValuesViewState);
          repWrapMultiSet.DataBind();
        }
      }
      else
      {
        DTAppliedAttributeValuesViewState = dtTemp;
        DTSelectedAttributeValuesViewState = GenerateBlankDTForAttributes();
      }
      if (!string.IsNullOrEmpty(SelectedNodeIDs))
      {
        PopluateAttributes();
        //SetProductCountMessage(Convert.ToInt32(HeirarchyProductCount));
        EnableControls();
        //Enable disable attribute set dropdown and excludeset radio
        EnableDisableSelectors(false);
      }
      IsGroupGrid = GetGridType();
    }
    if (DTAppliedAttributeValuesViewState != null && DTAppliedAttributeValuesViewState.Rows.Count > 0)
    {
      setEditBoxHiddenFieldValue();
      //EnableControls(false);
    }
    gvGroupLevel.Attributes.Add("style", "word-break:break-all; word-wrap:break-word");
    gvData.Attributes.Add("style", "word-break:break-all; word-wrap:break-word");
    UpdateTemplateProperties();
    UpdateShowDetailsLink();
  }

  private void UpdateShowDetailsLink()
  {
    if (IsGroupGrid)
    {
      if (!IsLevelFilterUpdated && !m_ProductGroup.IsExclusionProcessesd(ProductGroupID).Result)
      {
        hlShowDetails.Enabled = false;
        hlShowDetails.Text = PhraseLib.Lookup("term.exclusion_proess", LanguageID, "Changes to the product group are in process by the agent");
      }

    }

  }

  private bool GetGridType()
  {
    DataTable resDt = m_Product.GetHierarchyGroupedLevelInfo(SelectedHierarchy).Result;
    if (resDt != null && resDt.Rows.Count>0)
    {
      GroupedOn= resDt.Rows[0]["GroupingLevelName"].ToString();
      hfGroupedOn.Value = GroupedOn;
      return true;
    }
    return false;
  }
  private void UpdateTemplateProperties()
  {
    if (!IsEditPermitted)
    {
      List<string> exceptIds = new List<string>() { "hlShowDetails", "hlBack", "btnTemp" };
      disableHierarchyTree.Value = "true";
      DisableControls(panelPAB, exceptIds);
    }
    else
      disableHierarchyTree.Value = "false";
  }
  protected void DisableControls(Control parent, List<string> exceptIDs)
  {
    foreach (Control c in parent.Controls)
    {
      if (!exceptIDs.Contains(c.ID))
      {
        Type type = c.GetType();
        PropertyInfo prop = type.GetProperty("Enabled");
        if (prop != null)
        {
          prop.SetValue(c, false, null);
        }
        if (type.Name == "RepeaterItem")
        {
          // For Included Attribute Set
          Repeater rptInclude = (Repeater)c.FindControl("repIncludeSet");
          Repeater rptExclude = (Repeater)c.FindControl("repExcludeSet");
          if (rptInclude != null && rptInclude.Items.Count > 0)
          {
            for (int item = 0; item < rptInclude.Items.Count; item++)
            {
              LinkButton tempIncludelb = (LinkButton)rptInclude.Items[item].FindControl("lnkBtnIncludeSet");
              if (tempIncludelb != null)
              {
                HtmlGenericControl divIncludeSetContainer1 = (HtmlGenericControl)rptInclude.Parent;
                divIncludeSetContainer1.Attributes.Add("onmouseover", null);
                divIncludeSetContainer1.Attributes.Add("onmouseout", null);
                HtmlGenericControl divExcludeSetContainer1 = (HtmlGenericControl)rptExclude.Parent;
                divExcludeSetContainer1.Attributes.Add("onmouseover", null);
                divExcludeSetContainer1.Attributes.Add("onmouseout", null);
                if (tempIncludelb.CssClass == "filterButton")
                  tempIncludelb.CssClass = tempIncludelb.CssClass.Replace("filterButton", "filterButton1");

                HtmlGenericControl divIncludeEditSet1 = (HtmlGenericControl)rptInclude.Items[item].FindControl("divIncludeEditSet");
                divIncludeEditSet1.Attributes.Add("title", "");
              }
            }
          }

          // For Excluded Attribute Set

          if (rptExclude != null && rptExclude.Items.Count > 0)
          {
            for (int item = 0; item < rptExclude.Items.Count; item++)
            {
              LinkButton tempExcludelb = (LinkButton)rptExclude.Items[item].FindControl("lnkBtnExcludeSet");
              if (tempExcludelb != null)
              {
                HtmlGenericControl divIncludeSetContainer1 = (HtmlGenericControl)rptInclude.Parent;
                divIncludeSetContainer1.Attributes.Add("onmouseover", null);
                divIncludeSetContainer1.Attributes.Add("onmouseout", null);
                HtmlGenericControl divExcludeSetContainer1 = (HtmlGenericControl)rptExclude.Parent;
                divExcludeSetContainer1.Attributes.Add("onmouseover", null);
                divExcludeSetContainer1.Attributes.Add("onmouseout", null);
                if (tempExcludelb.CssClass == "filterButton")
                  tempExcludelb.CssClass = tempExcludelb.CssClass.Replace("filterButton", "filterButton1");
                HtmlGenericControl divExcludeEditSet1 = (HtmlGenericControl)rptExclude.Items[item].FindControl("divExcludeEditSet");
                divExcludeEditSet1.Attributes.Add("title", "");
              }
            }
          }

        }
        if (c.Controls.Count > 0)
        {
          this.DisableControls(c, exceptIDs);
        }
      }
    }
  }

  protected void gvHeaderTest_RowDataBound(object sender, GridViewRowEventArgs e)
  {
    e.Row.Cells[0].Visible = false;
    e.Row.Cells[2].Visible = false;
  }

  protected void gvData_Sorting(object sender, GridViewSortEventArgs e)
  {
    IsFilterUpdated = true;
    _sortOrder = gvData.SortOrder;
    _sortBy = gvData.SortKey;
    btnTemp.Text = "0";
    RefreshGrid();
    //SetProductCountMessage(Convert.ToInt32(HeirarchyProductCount), false);
    backImg.Visible = hlBack.Visible = true;
  }
  protected void ddlAttributeType_SelectedIndexChanged(object sender, EventArgs e)
  {
    if (!string.IsNullOrEmpty(hdnEjectedValues.Value))
      hdnEjectedValues.Value = "";
    ddlAttributeValue.Enabled = false;
    btnAddFilter.Enabled = false;
    PopulateAttributeValue();
    ddlAttributeValue.Enabled = true;
    btnAddFilter.Enabled = true;
  }
  protected void repWrapMultiSet_ItemCommand(object source, RepeaterCommandEventArgs e)
  {
    IsLevelFilterUpdated = IsFilterUpdated = true;
    Repeater repSource = source as Repeater;
    int setID = 0;
    bool excludeFlag = false;
    hdnExcludeSetExists.Value = Boolean.FalseString;

    if (e.CommandArgument != null)
    {
      setID = e.CommandArgument.ConvertToInt16();
    }

    if (repSource.ID == "repIncludeSet")
    {
      DTSelectedAttributeValuesViewState = GetAttributeSetDTBeingEdited(setID, false);
      ddlAttributeSet.SelectedIndex = 0;

      //Find if the current included set clicked for edit has excluded set for displaying warning message using hdnExcludeSetExists
      bool exclusionExistsForSet = (DTAppliedAttributeValuesViewState.AsEnumerable()
                      .Count(r => r["AttributeSetID"].ConvertToInt16() == setID && r["ExcludeFlag"].ConvertToBool() == true)) == 0 ? false : true;
      hdnExcludeSetExists.Value = exclusionExistsForSet.ToString();
    }
    else if (repSource.ID == "repExcludeSet")
    {
      DTSelectedAttributeValuesViewState = GetAttributeSetDTBeingEdited(setID, true);
      //BindAttributeSetDDL();
      ddlAttributeSet.Items.Insert(1, setID.ToString());
      ddlAttributeSet.SelectedIndex = 1;

      excludeFlag = true;
    }
    //Set edited element id for keeping the clicked set highlighted
    hdnEditedElementID.Value = repSource.Parent.ClientID;
    DisableControls(repSource.Parent.Parent.Parent.Parent, new List<string>());
    repFilter.DataSource = GetDataSourceForSelectedSetUI(DTSelectedAttributeValuesViewState);
    repFilter.DataBind();

    //EnableControls(true);
    ClearSelectionOnGrid();
    //Reset the count
    //HeirarchyProductCount = null;
    hdnConsiderExclusions.Value = "0";

    DetailedProductList.Visible = false;
    //Enable disable attribute set dropdown and radio
    EnableDisableSelectors(false, true, excludeFlag);
    EnableControls();
    PopluateAttributes();
  }
  protected void repFilter_ItemCommand(object source, RepeaterCommandEventArgs e)
  {

    if (e.CommandName == "FilterUpdate")
    {
      //copy the view state table into another data table.
      dtUpdate = DTSelectedAttributeValuesViewState.Copy().AsEnumerable().Where(r => r.Field<Int16>("AttributeTypeID") == Convert.ToInt32(e.CommandArgument)).CopyToDataTable();
      //remove values for selected typeID from viewstate so that later it can bind
      DTSelectedAttributeValuesViewState.AsEnumerable().Where(r => r.Field<Int16>("AttributeTypeID") == Convert.ToInt32(e.CommandArgument)).ToList().ForEach(row => row.Delete());
      DTSelectedAttributeValuesViewState.AcceptChanges();
      IsLevelFilterUpdated = IsFilterUpdated = true;
      repFilter.DataSource = GetDataSourceForSelectedSetUI(DTSelectedAttributeValuesViewState);
      repFilter.DataBind();
      //RefreshGrid();
      PopulateAttributeType();
      if (ddlAttributeType.Items.Count > 0)// && ddlAttributeType.SelectedItem.Value.ConvertToInt32() == Convert.ToInt32(e.CommandArgument))
      {
        for (int i = 0; i < ddlAttributeType.Items.Count; i++)
        {
          if (Convert.ToInt32(ddlAttributeType.Items[i].Value) == Convert.ToInt32(e.CommandArgument))
            ddlAttributeType.SelectedIndex = i;
        }
        hdnEjectedValues.Value = "";
        foreach (DataRow item in dtUpdate.Rows)
        {
          hdnEjectedValues.Value += item["Value"].ToString() + ",";
        }
        hdnEjectedValues.Value = hdnEjectedValues.Value.Remove(hdnEjectedValues.Value.Length - 1);
        PopulateAttributeValue();

      }
      //ddlAttributeType.Enabled = true;
      //ddlAttributeValue.Enabled = true;
      EnableControls();
    }
  }
  protected void btnDummyForcatchingEvent_Click(object sender, EventArgs e)
  {
    if (string.IsNullOrEmpty(SelectedNodeIDs))
      return;
    PopluateAttributes();
    EnableControls();
    ShowControls(true);
    PrevSelectedNodeIDs = SelectedNodeIDs;
    ClearSelectionOnGrid();
    hdnConsiderExclusions.Value = "0";
    //HeirarchyProductCount = null;
    //SetProductCountMessage(Convert.ToInt32(HeirarchyProductCount));
    //Setting this value will trigger count fetch based on the nodes
    hdnPABAVPairsJson.Value = NODE_PRODUCT_COUNT_FLAG;
    BindAttributeSetDDL();
    EnableDisableSelectors();
  }
  protected void btnAddFilter_Click(object sender, EventArgs e)
  {
    IsFilterUpdated = true;
    ClearFilters = false;
    if (!String.IsNullOrEmpty(hdnEjectedValues.Value))
      hdnEjectedValues.Value = String.Empty;
    if (!String.IsNullOrEmpty(hdnSelectedValues.Value))
    {
      string[] selectedIds = (hdnSelectedIDs.Value.ToString()).Split(',');
      string[] selectedValues = (hdnSelectedValues.Value.ToString()).Split(',');

      int nextAttributeSetID = 0;
      bool currentSetEditFlag = IsCurrentSetInEditMode(out nextAttributeSetID);

      if (!currentSetEditFlag)
        nextAttributeSetID = GetNextAttributeSetID();

      if (ddlAttributeType.Items.Count > 0)
      {
        for (int i = 0; i < selectedIds.Length; i++)
        {
          DTSelectedAttributeValuesViewState.Rows.Add(nextAttributeSetID, Convert.ToInt64(ddlAttributeType.SelectedItem.Value),
              ddlAttributeType.SelectedItem.Text.ToString(), Convert.ToInt64(selectedIds[i]), selectedValues[i].ToString(), radExcludedSet.Checked,
              currentSetEditFlag);
        }
      }
    }
    hdnSelectedIDs.Value = String.Empty;
    hdnSelectedValues.Value = String.Empty;
    if (DTSelectedAttributeValuesViewState.Rows.Count > 0)
    {
      repFilter.DataSource = GetDataSourceForSelectedSetUI(DTSelectedAttributeValuesViewState);
      repFilter.DataBind();
    }
    PopluateAttributes();
    EnableControls();
  }
  protected void gvData_RowDataBound(object sender, GridViewRowEventArgs e)
  {
    if (e.Row.Cells[0].GetType() == typeof(System.Web.UI.WebControls.DataControlFieldCell))
    {
      TableCell tc = e.Row.Cells[1];
      if (tc.Controls.Count > 0)
      {
        CheckBox cb = (CheckBox)tc.Controls[0];
        if (!(cb == null))
        {
          cb.Enabled = true;
          cb.Checked = false;
          cb.Attributes.Add("onclick", "javascript:ToggleMasterCheckbox(this)");
          string val = e.Row.Cells[2].Text.ToString();

          if (ProductIDExcluded.Contains(val) ||
              (exclList.Exists(p => p.ProductID.ToString().Equals(val)) && hdnConsiderExclusions.Value == "1" && hdnIsMasterCBChecked.Value != "0") ||
              hdnIsMasterCBChecked.Value == "1")
          {
            cb.Checked = true;
          }
          if (ProductIDInculded.Contains(val))
          {
            cb.Checked = false;
          }
        }
      }
    }


    e.Row.Cells[0].Visible = false;
    e.Row.Cells[2].Visible = false;
  }
  protected void EditTempBtn_Click(object sender, EventArgs e)
  {
    var EditValuesText = from v in DTAppliedAttributeValuesViewState.AsEnumerable()
                         group v by v.Field<string>("Name") into g
                         select new { Name = g.Key, values = g };
    string edittext = "";
    foreach (var item in EditValuesText)
    {
      edittext += "<b>" + item.Name + "</b>:";
      foreach (var i in item.values)
      {
        edittext += i["value"] + ",";
      }
      edittext = edittext.Remove(edittext.LastIndexOf(','));
      edittext += "</br>";
    }
    Page.ClientScript.RegisterStartupScript(GetType(), "hwa", "editAttributeSetDialog('" + edittext + "');", true);
  }
  /// <summary>
  /// This handler is called when scrolling happens in the old grid without grouping
  /// </summary>
  /// <param name="sender"></param>
  /// <param name="e"></param>
  protected void btnTemp_Click(object sender, EventArgs e)
  {
    try
    {
      string strExcludedProductIDs = "";
      GridViewCheckedItems();
      PageIndex = btnTemp.Text.ConvertToInt32();
      PageIndex++;
      btnTemp.Text = PageIndex.ToString();
      AMSResult<List<AttributeType>> objType = new AMSResult<List<AttributeType>>();
      AMSResult<DataTable> dtResult = new AMSResult<DataTable>();
      List<AttributeValue> lstAttributeValue = DTSelectedAttributeValuesViewState.Rows.Count > 0 ? DTSelectedAttributeValuesViewState.ToGenericList<AttributeValue>() :
        DTAppliedAttributeValuesViewState.ToGenericList<AttributeValue>();
      if (IsFilterUpdated == true)
      {
        strExcludedProductIDs = String.Join(",", ProductIDExcluded.ToArray());
      }
      Int64 RowNumIndexStart = (PageIndex * PageSize);
      Int32 CachedPageNum = pageNum.Value.ConvertToInt32();
      if (GridViewCacheData.Result == null)
      {
        pageNum.Value = (pageNum.Value.ConvertToInt32() + 1).ToString();
        CachedPageNum = pageNum.Value.ConvertToInt32();
        if (DTAppliedAttributeValuesViewState.Rows.Count > 0)
        {
          DataTable dt = DTAppliedAttributeValuesViewState.DefaultView.ToTable(false, "AttributeSetID", "AttributeTypeID", "AttributeValueID", "ExcludeFlag");
          GridViewCacheData = m_Product.GetProductsByNodeAndAttributeValues(AllChildNodes, CachedPageNum, _sortBy, _sortOrder, dt, AttributeTypeToDisplayViewState, hdnProductGroupID.Value.ConvertToLong(), strExcludedProductIDs, IsFilterUpdated, out TotalCount);
        }
        else if (lstAttributeValue != null && lstAttributeValue.Count > 0)
          GridViewCacheData = m_Product.GetProductsByAttributeValues(CachedPageNum, PageSize * 20, _sortBy, _sortOrder, lstAttributeValue, AttributeTypeToDisplayViewState, hdnProductGroupID.Value.ConvertToLong(), strExcludedProductIDs, IsFilterUpdated, out TotalCount);
        else
          GridViewCacheData = m_Product.GetProductsByNode(SelectedNodeIDs, CachedPageNum, (PageSize * 20), _sortBy, _sortOrder, AttributeTypeToDisplayViewState, hdnProductGroupID.Value.ConvertToLong(), strExcludedProductIDs, IsFilterUpdated, out TotalCount);
        pageNum.Value = CachedPageNum.ToString();
        TotalCount = hdnInclProducts.Value.ConvertToInt32();
      }
      //Once reach to last record in cache and there are few more record exist load those record.
      if (GridViewCacheData.Result.Rows.Count <= TotalCount && (GridViewCacheData.Result.Rows.Count / (PageIndex * 50)) >= Count)
      {
        int TempIndex = PageIndex;
        PageIndex = (GridViewCacheData.Result.Rows.Count / (PageIndex * 50) + 1);
        pageNum.Value = (pageNum.Value.ConvertToInt32() + 1).ToString();
        CachedPageNum = pageNum.Value.ConvertToInt32();
        if (DTAppliedAttributeValuesViewState.Rows.Count > 0)
        {
          DataTable dt = DTAppliedAttributeValuesViewState.DefaultView.ToTable(false, "AttributeSetID", "AttributeTypeID", "AttributeValueID", "ExcludeFlag");
          GridViewCacheData = m_Product.GetProductsByNodeAndAttributeValues(AllChildNodes, PageIndex, _sortBy, _sortOrder, dt, AttributeTypeToDisplayViewState, hdnProductGroupID.Value.ConvertToLong(), strExcludedProductIDs, IsFilterUpdated, out TotalCount);
        }
        Count = PageIndex + 1;
        //Keep existing page index because all data gets loaded.
        PageIndex = TempIndex;

      }
      var rows = (from DataRow row in GridViewCacheData.Result.Rows
                  where row.Field<Int64>("RowNum") > RowNumIndexStart && row.Field<Int64>("RowNum") <= (RowNumIndexStart + PageSize)
                  select row);
      if (rows.Any())
        dtResult.Result = rows.CopyToDataTable();
      dtResult.ResultType = GridViewCacheData.ResultType;
      if ((PageIndex + 1) % 20 == 0)
      {
        GridViewCacheData.Result = null;
      }

      if (dtResult.ResultType == AMSResultType.Success)
      {
          if (GridViewData != null && dtResult.Result != null && dtResult.Result.Rows.Count > 0)
        {
          var drs = dtResult.Result.AsEnumerable().Take(dtResult.Result.Rows.Count);
          if (drs.Any())
            drs.CopyToDataTable(GridViewData, LoadOption.OverwriteChanges);
        }
          if (GridViewData != null && dtResult.Result != null && (GridViewData.Rows.Count >= hdnTotalRecords.Value.ConvertToInt32() || dtResult.Result.Rows.Count == 0))
          lbNeedReload.Text = "False";
        gvData.DataSource = GridViewData;
        IsFilterUpdated = false;
        exclList = m_Product.GetExcludedProducts(ProductGroupID).Result;
        //SetProductCountMessage(Convert.ToInt32(HeirarchyProductCount), false);
      }
      else
      {
        BindGridHeader();
      }
      gvData.DataBind();
      backImg.Visible = hlBack.Visible = true;
    }
    finally
    {
      GridDataLoading.Value = "False";
    }
  }
  protected void btnApplyFilter_Click(object sender, EventArgs e)
  {
    IsLevelFilterUpdated = true;
    ClearFilters = false;
    if (DTSelectedAttributeValuesViewState != null && DTSelectedAttributeValuesViewState.Rows.Count > 0)
    {
      int setID = 0;
      bool excludeFlag = radExcludedSet.Checked;
      bool editFlag = IsCurrentSetInEditMode(out setID);
      if (editFlag)
        hdnExludedProductsCount.Value = "0";
      //Merge currently selected set into the applied set data table
      MergeSelectedSetAppliedSetDT(editFlag, setID, excludeFlag);

      repWrapMultiSet.DataSource = GetDataSourceForMultiAttributeSetBinding(DTAppliedAttributeValuesViewState);
      repWrapMultiSet.DataBind();

      //Bind the attribute set dropdown
      BindAttributeSetDDL();
      //Enable disable attribute set dropdown and radio
      EnableDisableSelectors(true);

      ResetCreateAttributeSet(true);
      ClearSelectionOnGrid();

      hdnConsiderExclusions.Value = "0";
      hlShowDetails.Visible = true;

      //Input for javascript async count call while pageload
      DataTable dtAVPairs = DTAppliedAttributeValuesViewState.DefaultView.ToTable(false, "AttributeSetID", "AttributeTypeID", "AttributeValueID", "ExcludeFlag");
      //dtAVPairs.ToJSON(); -- Gives circular reference error
      hdnPABAVPairsJson.Value = PrepareJSON(dtAVPairs);
      //HeirarchyProductCount = null;
      //SetProductCountMessage(Convert.ToInt32(HeirarchyProductCount));
      setEditBoxHiddenFieldValue();

      //Clear the clicked set id from hidden field
      hdnEditedElementID.Value = string.Empty;

      EnableControls();
      ClearEditFlag(setID, excludeFlag);
      //Clear the field hdnExcludeSetExists to prevent warning message
      hdnExcludeSetExists.Value = Boolean.FalseString;
    }
  }

  protected void btnClear_Click(object sender, EventArgs e)
  {
    ResetAppliedAttributeSet(radExcludedSet.Checked);
    ResetCreateAttributeSet();
    EnableControls();

    //Clear the clicked set id from hidden field
    hdnEditedElementID.Value = string.Empty;
    //Clear the field hdnExcludeSetExists to prevent warning message
    hdnExcludeSetExists.Value = Boolean.FalseString;

    //If there are no more sets remaining fetch the product count on page load based on selected nodes
    if (DTAppliedAttributeValuesViewState.Rows.Count == 0)
      hdnPABAVPairsJson.Value = NODE_PRODUCT_COUNT_FLAG;
    else
      hdnPABAVPairsJson.Value = PrepareJSON(DTAppliedAttributeValuesViewState.DefaultView.ToTable(false, "AttributeSetID", "AttributeTypeID", "AttributeValueID", "ExcludeFlag"));
  }
  protected void hlBack_Click(object sender, EventArgs e)
  {
    ShowControls(true);
    //Fix for : First Value in Attribute Value Select List is automatically getting selected.
    ListItem item1 = ddlAttributeValue.Items.Cast<ListItem>().Where(p => p.Selected == true).FirstOrDefault();
    if (item1 != null)
      item1.Selected = false;

    EnableControls();
    //SetProductCountMessage(Convert.ToInt32(HeirarchyProductCount), false);
    BindAttributeSetDDL();
    EnableDisableSelectors();
    //Update the level information
    if (IsGroupGrid)
      LevelGridViewDetails = (DataTable)JsonConvert.DeserializeObject(hfLevelInfoTable.Value, (typeof(DataTable)));

  }
  protected void hlShowDetails_Click(object sender, EventArgs e)
  {
    ShowControls(false);
    //If Grouping is not enabled show normal grid

    if (!IsGroupGrid)
    {
      RefreshGrid();
    }
    else
    {
      hfGroupgrid.Value = "true";
      PopulateLevelGV(0, 100);
    }
  }
  protected void btnReloadHierarchytree_Click(object sender, EventArgs e)
  {
    //EnableControls(false);
    hdnPABStage.Value = "1";
    //repFilter.DataSource = null;
    //repFilter.DataBind();
    DTSelectedAttributeValuesViewState.Clear();
    DTAppliedAttributeValuesViewState.Clear();
    repFilter.DataSource = GetDataSourceForSelectedSetUI(DTSelectedAttributeValuesViewState);
    repFilter.DataBind();
    repWrapMultiSet.DataSource = GetDataSourceForMultiAttributeSetBinding(DTAppliedAttributeValuesViewState);
    repWrapMultiSet.DataBind();
    ClearFilters = true;
    LoadHierarchyTree();
    ddlAttributeType.Items.Clear();
    ddlAttributeValue.Items.Clear();
    ClearSelectionOnGrid();
    //For Level Group grid
    IsLevelFilterUpdated = true;
  }
  protected void ddlAttributeSet_SelectedIndexChanged(object sender, EventArgs e)
  {
    DropDownList ddl = sender as DropDownList;
    ResetCreateAttributeSet();
    EnableDisableSelectors(false, false, false, false, true);
  }
  protected void radIncludedSet_CheckedChanged(object sender, EventArgs e)
  {
    ResetCreateAttributeSet();
    EnableDisableSelectors();
    if (ddlAttributeSet.Items.Count > 0)
      ddlAttributeSet.SelectedIndex = 0;
  }
  protected void btnLocateHierarchyTree_Click(object sender, EventArgs e)
  {
    LoadHierarchyTree();
  }
  #endregion

  #region Private Methods
  /// <summary>
  /// Get the set which has been clicked for edit
  /// </summary>
  /// <param name="attributeSetID"></param>
  /// <param name="excludeFlag"></param>
  /// <returns></returns>
  private string PrepareJSON(DataTable dt)
  {
    List<Dictionary<string, object>> list = new List<Dictionary<string, object>>();

    foreach (DataRow row in dt.Rows)
    {
      var dict = new Dictionary<string, object>();
      foreach (DataColumn col in dt.Columns)
        dict.Add(col.ToString(), row[col]);

      list.Add(dict);
    }

    return JSONHelper.ToJSON(list);
  }
  private DataTable GetAttributeSetDTBeingEdited(int attributeSetID, bool excludeFlag)
  {
    DataTable editSetDT = GenerateBlankDTForAttributes();
    var rows = from row in DTAppliedAttributeValuesViewState.AsEnumerable()
               where row.Field<Int16>("AttributeSetID") == attributeSetID && row.Field<bool>("ExcludeFlag") == excludeFlag
               select row;
    foreach (DataRow row in rows)
    {
      //Set applied row edit flag to true
      row["EditFlag"] = true;

      //copy the row fields to new table
      DataRow newRow = editSetDT.NewRow();
      newRow["AttributeSetID"] = row["AttributeSetID"];
      newRow["AttributeTypeID"] = row["AttributeTypeID"];
      newRow["Name"] = row["Name"];
      newRow["AttributeValueID"] = row["AttributeValueID"];
      newRow["value"] = row["value"];
      newRow["ExcludeFlag"] = row["ExcludeFlag"];
      newRow["EditFlag"] = row["EditFlag"];
      editSetDT.Rows.Add(newRow);
    }
    DTAppliedAttributeValuesViewState.AcceptChanges();
    return editSetDT;
  }
  private void ClearSelectionOnGrid()
  {
    ProductIDInculded.Clear();
    ProductIDExcluded.Clear();
    gvData.DataSource = null;
    gvData.DataBind();
    btnTemp.Text = "0";
    GridViewCacheData = new AMSResult<DataTable>();
    GridViewData = new DataTable();
    hdnConsiderExclusions.Value = "1";
    hdnIsMasterCBChecked.Value = "";
  }
  private void ResetCreateAttributeSet(bool appliedSetFlag = false)
  {
    hdnEjectedValues.Value = string.Empty;
    DTSelectedAttributeValuesViewState.Rows.Clear();
    repFilter.DataSource = DTSelectedAttributeValuesViewState;
    repFilter.DataBind();
    PopluateAttributes(appliedSetFlag);
    //This is taking 90% of time. Commented to improve the performance. Need to check all the scenarios
    //if (!IsGroupGrid)
    //RefreshGrid();
  }
  /// <summary>
  /// Function to remove the excluded set or included+excluded set in edit mode 
  /// </summary>
  /// <param name="excludeFlag"></param>
  private void ResetAppliedAttributeSet(bool excludeFlag)
  {
    int setID = 0;

    if (IsCurrentSetInEditMode(out setID))
    {
      if (excludeFlag)
      {
        DTAppliedAttributeValuesViewState.AsEnumerable()
        .Where(r => r.Field<Int16>("AttributeSetID") == setID && r.Field<bool>("ExcludeFlag") == true)
        .ToList()
        .ForEach(row => row.Delete());

        DTAppliedAttributeValuesViewState.AcceptChanges();
        BindAttributeSetDDL();
        EnableDisableSelectors(false, false, true, true);
      }
      else
      {
        //Remove the set
        DTAppliedAttributeValuesViewState.AsEnumerable()
        .Where(r => r.Field<Int16>("AttributeSetID") == setID)
        .ToList()
        .ForEach(row => row.Delete());

        DTAppliedAttributeValuesViewState.AcceptChanges();

        //Renumber the setids after removing above setid 
        DTAppliedAttributeValuesViewState.AsEnumerable()
            .ToList()
            .ForEach(delegate(DataRow row)
        {
          int tempSetID = row.Field<Int16>("AttributeSetID");
          if (tempSetID > setID)
            row["AttributeSetID"] = tempSetID - 1;
        }
        );

        //Bind the attribute set drop down after renumbering
        BindAttributeSetDDL();
        EnableDisableSelectors(false, false, false, true);
      }

      repWrapMultiSet.DataSource = GetDataSourceForMultiAttributeSetBinding(DTAppliedAttributeValuesViewState);
      repWrapMultiSet.DataBind();
    }
  }

  private bool IsCurrentSetInEditMode(out int setID)
  {
    setID = 0;
    if (DTAppliedAttributeValuesViewState != null && DTAppliedAttributeValuesViewState.Rows.Count > 0)
    {
      var row = DTAppliedAttributeValuesViewState.AsEnumerable().FirstOrDefault(r => r["EditFlag"] != null && r["EditFlag"].ConvertToBool() == true);

      if (row != null)
      {
        setID = row.Field<Int16>("AttributeSetID");
        return true;
      }
    }
    return false;
  }
  /// <summary>
  /// Find setid based on exclusion\inclusion\edit mode
  /// </summary>
  /// <returns></returns>
  private int GetNextAttributeSetID()
  {
    int nextAttributeSetID = 0;

    if (radExcludedSet.Checked && ddlAttributeSet.SelectedIndex > 0 && ddlAttributeSet.SelectedItem != null)
      nextAttributeSetID = ddlAttributeSet.SelectedItem.Text.ConvertToInt16();
    else
    {
      nextAttributeSetID = DTAppliedAttributeValuesViewState.Compute("max(AttributeSetID)", string.Empty).ConvertToInt16();
      nextAttributeSetID = nextAttributeSetID == 0 ? 1 : nextAttributeSetID + 1;
    }
    return nextAttributeSetID;
  }
  private void MergeSelectedSetAppliedSetDT(bool editFlag, int setID, bool excludeFlag)
  {
    //If edit flag is set then first remove the old included or excluded rows and merge freshly prepared set
    if (editFlag)
    {
      DTAppliedAttributeValuesViewState.AsEnumerable()
          .Where(r => r.Field<Int16>("AttributeSetID") == setID && r.Field<bool>("ExcludeFlag") == excludeFlag)
          .ToList()
          .ForEach(row => row.Delete());
      DTAppliedAttributeValuesViewState.AcceptChanges();
    }

    DTAppliedAttributeValuesViewState.Merge(DTSelectedAttributeValuesViewState);

    DTSelectedAttributeValuesViewState.Clear();
  }
  private void ClearEditFlag(int setID, bool excludeFlag)
  {
    var rows = from row in DTAppliedAttributeValuesViewState.AsEnumerable()
               where row.Field<Int16>("AttributeSetID") == setID && row.Field<bool>("ExcludeFlag") == excludeFlag
               select row;
    foreach (DataRow row in rows)
    {
      row["EditFlag"] = false;
    }
    DTAppliedAttributeValuesViewState.AcceptChanges();
  }
  private void LoadBuyerNodesintoSession(int BuyerID)
  {
    if (BuyerID > 0 && Session["AllBuyerNodes" + BuyerID.ToString()] == null)
    {
      Session["AllBuyerNodes" + BuyerID.ToString()] = m_ProductGroup.GetAllNodesAssociatedWithBuyer(BuyerID).Result;
    }
  }
  private void PopulateAttributeValue(bool appliedSetFlag = false)
  {
    ddlAttributeValue.Attributes.Add("multiple", "multiple");
    ddlAttributeValue.Items.Clear();
    List<AttributeKeyValue> objs = new List<AttributeKeyValue>();
    if (ddlAttributeType.SelectedItem != null)
    {
      AttributeType objAttributeType = new AttributeType();
      objAttributeType.AttributeTypeID = ddlAttributeType.SelectedItem.Value.ConvertToInt16();
      List<string> SelectedValues = new List<string>();
      foreach (DataRow item in DTSelectedAttributeValuesViewState.Rows)
      {
        SelectedValues.Add(item["AttributeValueID"].ToString());
        AttributeKeyValue obj = new AttributeKeyValue();
        obj.AttributeTypeID = item["AttributeTypeID"].ConvertToInt16();
        obj.AttributeValueID = item["AttributeValueID"].ConvertToInt32();
        objs.Add(obj);
      }
      string str = new JavaScriptSerializer().Serialize(objs);
      str = str.Replace("\"", "'");
      hdnKeyValue.Value = str;

      hdnSelectedAttributeValues.Value = String.Join(",", SelectedValues);// Use by web service to load chunk data 
      DataTable dt = !appliedSetFlag ? GetMergedInputDTForAttributeValueRequery() : null;
      AMSResult<List<AttributeValue>> objResult;
      if (dt != null && dt.Rows.Count > 0)
        objResult = m_Attribute.ReGetAllLinkedAttributeValueByType_WithAttributes(objAttributeType, String.Empty, 0, AllChildNodes, dt, SelectedValues.Count > 0 ? SelectedValues : null);
      else
        objResult = m_Attribute.GetAllLinkedAttributeValueByType(objAttributeType, String.Empty, 0, AllChildNodes, SelectedValues.Count > 0 ? SelectedValues : null);
      if (objResult.ResultType == AMSResultType.Success)
      {
        if (objResult.Result.Count > 0)
        {
          ddlAttributeValue.DataTextField = "Value";
          ddlAttributeValue.DataValueField = "AttributeValueID";
          ddlAttributeValue.DataSource = objResult.Result;
          ddlAttributeValue.DataBind();
        }
      }
    }
  }

  private void SetContolsVisibility()
  {
    hidReadOnly.Value = IsEditPermitted.ToString();
    bool _enabled = true;
    if (hidReadOnly.Value.Equals("false", StringComparison.InvariantCultureIgnoreCase))
      _enabled = false;
    gvData.Enabled = _enabled;
    //btnAddFilter.Enabled = _enabled;
    ButtonStatus = _enabled;
    lblFilter.Enabled = _enabled;
    hlBack.Visible = backImg.Visible = false;
    lbTotalProducts.Visible = true;
  }
  private void BindGridHeader()
  {
    DataTable dt = new DataTable();
    dt.Columns.Add("RowNum", typeof(int));
    dt.Columns.Add("Excluded", typeof(bool));
    dt.Columns.Add("ProductID", typeof(long));
    if (AttributeTypeToDisplayViewState != null)
    {
      foreach (AttributeType item in AttributeTypeToDisplayViewState)
        dt.Columns.Add(item.AttributeName, typeof(string));
    }
    //dt.Columns.Add(PhraseLib.Lookup("term.description", LanguageID), typeof(string));
    if (dt.Columns.Count > 0)
      GenerateColumn(dt, gvData);
    gvData.DataSource = dt;
    gvData.DataBind();
    lblexcludedcheckboxdetail.Text = String.Empty;

  }
  private DataTable GenerateBlankDTForAttributes()
  {
    DataTable dt = new DataTable();
    dt.Columns.Add("AttributeSetID", typeof(Int16));
    dt.Columns.Add("AttributeTypeID", typeof(Int16));
    dt.Columns.Add("Name", typeof(string));
    dt.Columns.Add("AttributeValueID", typeof(Int32));
    dt.Columns.Add("value", typeof(string));
    dt.Columns.Add("ExcludeFlag", typeof(bool));
    dt.Columns.Add("EditFlag", typeof(bool));
    return dt;
  }
  private DataTable GenerateBlankDTForSaving()
  {
    DataTable dt = new DataTable();
    dt.Columns.Add("AutoID", typeof(int));
    dt.Columns.Add("AttributeSetID", typeof(string));
    dt.Columns.Add("AttributeTypeID", typeof(Int16));
    dt.Columns.Add("AttributeValueID", typeof(Int32));
    return dt;
  }

  private void PopluateAttributes(bool appliedSetFlag = false)
  {
    ddlAttributeType.Enabled = false;
    ddlAttributeValue.Enabled = false;
    PopulateAttributeType(appliedSetFlag);
    if (ddlAttributeType.Items.Count > 0)
      ddlAttributeType.Enabled = true;
    PopulateAttributeValue(appliedSetFlag);
    if (ddlAttributeValue.Items.Count > 0)
      ddlAttributeValue.Enabled = true;

  }
  private DataTable GetMergedInputDTForAttributeValueRequery()
  {
    DataTable dtInputSet = DTSelectedAttributeValuesViewState.DefaultView.ToTable(false, "AttributeSetID", "AttributeTypeID", "AttributeValueID", "ExcludeFlag");

    if (radExcludedSet.Checked && ddlAttributeSet.SelectedIndex > 0)
    {
      //Find and merge the included set if the current set is being excluded with a set id selected
      int includeSetIDForExclusion = ddlAttributeSet.SelectedItem.Text.ConvertToInt16();
      if (includeSetIDForExclusion > 0 && DTAppliedAttributeValuesViewState != null && DTAppliedAttributeValuesViewState.Rows.Count > 0)
      {
        var rowColl = from row in DTAppliedAttributeValuesViewState.AsEnumerable()
                      where row.Field<Int16>("AttributeSetID") == includeSetIDForExclusion && row.Field<bool>("ExcludeFlag") == false
                      select row;

        if (rowColl != null && rowColl.Count() > 0)
        {
          DataTable dtIncludeSet = rowColl.CopyToDataTable().DefaultView.ToTable(false, "AttributeSetID", "AttributeTypeID", "AttributeValueID", "ExcludeFlag");
          dtInputSet.Merge(dtIncludeSet);
        }
      }
    }
    return dtInputSet;
  }
  private void PopulateAttributeType(bool appliedSetFlag = false)
  {
    ddlAttributeType.Items.Clear();
    //In case set is applied, we have to get fresh values for attribute type\value dropdowns
    DataTable dtInput = !appliedSetFlag ? GetMergedInputDTForAttributeValueRequery() : null;

    if (dtInput != null && dtInput.Rows.Count > 0)
      objAttributeType = m_Attribute.ReGetAllLinkedAttributeTypes_WithAttributes(AllChildNodes, dtInput);
    else
    {
      if (AttributeTypes == null)
      {
        AttributeTypes = m_Attribute.GetAllLinkedAttributeTypes(AllChildNodes);
      }
      objAttributeType = AttributeTypes;
    }

    if (objAttributeType.ResultType == AMSResultType.Success)
    {
      if (objAttributeType.Result.Count > 0)
      {
        ViewState["AttributeType"] = objAttributeType.Result;
        ddlAttributeType.DataTextField = "AttributeName";
        ddlAttributeType.DataValueField = "AttributeTypeID";
        ddlAttributeType.DataSource = objAttributeType.Result;
        ddlAttributeType.DataBind();
      }
    }
  }
  private void PopulateControlText()
  {
    lblAttributeType.Text = PhraseLib.Lookup("term.attributetype", LanguageID) + ": ";
    lblAttributeValue.Text = PhraseLib.Lookup("term.value", LanguageID) + ": ";
    btnAddFilter.Text = PhraseLib.Lookup("term.addfilter", LanguageID);//Add attribute to set
    lblFilter.Text = PhraseLib.Lookup("term.currentfilters", LanguageID);//Set Contains:
  }
  private void PopulateAttributeTypeToDisplay()
  {
    objAttributeType = m_Attribute.GetAllAttributeTypesTodisplay();
    if (objAttributeType.ResultType == AMSResultType.Success)
    {
      if (objAttributeType.Result.Count > 0)
        AttributeTypeToDisplayViewState = objAttributeType.Result;
      else
        AttributeTypeToDisplayViewState = new List<AttributeType>();
    }
  }

  private void GenerateColumn(DataTable dt, GridView gridView)
  {
    gridView.Columns.Clear();
    gridView.AutoGenerateColumns = false;
    foreach (DataColumn col in dt.Columns)
    {
      if (col.ColumnName != "Excluded")
      {
        //Declare the bound field and allocate memory for the bound field.
        BoundField bfield = new BoundField();

        //Initalize the DataField value.
        bfield.DataField = col.ColumnName;
        //Initialize the HeaderText field value.
        bfield.HeaderText = col.ColumnName;

        //... As per CLOUDSOL-1284 , we need to change ExtProductID to ItemUPC.  We can't changed it into Database 
        if (col.ColumnName.Equals("ExtProductID", StringComparison.InvariantCultureIgnoreCase))
        {
          bfield.HeaderText = PhraseLib.Lookup("term.itemsupc", LanguageID);// "Item UPC";
        }
        //bfield.ItemStyle.CssClass = "overflow:hidden;text-overflow:ellipsis;white-space:nowrap;";
        if (!col.ColumnName.Equals("Description", StringComparison.InvariantCultureIgnoreCase))
        {
          bfield.SortExpression = col.ColumnName;
        }
        else
          bfield.HeaderText = PhraseLib.Lookup("term.description", LanguageID);
        //bfield.ItemStyle.Width = Unit.Pixel(35);
        //Add the newly created bound field to the GridView.
        gridView.Columns.Add(bfield);
      }
      else
      {
        //Declare the bound field and allocate memory for the bound field.
        CheckBoxField cfield = new CheckBoxField();

        //Initalize the DataField value.
        cfield.DataField = col.ColumnName;
        //Initialize the HeaderText field value.
        cfield.HeaderText = PhraseLib.Lookup("term.excluded", LanguageID);
        //cfield.ItemStyle.Width = Unit.Pixel(5);
        cfield.HeaderText = @"<div style=""width:40px""><input type=""checkbox"" id=""checkall"" onclick=""javascript:ToggleCheckbox(this.checked);""/>" + "<img src=\"/images/Information.png\" alt=\"" + PhraseLib.Lookup("term.excludedprodcheckbox", LanguageID) + "\" title=\"" + PhraseLib.Lookup("term.excludedprodcheckbox", LanguageID) + "\" " + " />" + "</div>";
        //Add the newly created bound field to the GridView.
        lblexcludedcheckboxdetail.Text = "<img src=\"/images/arrow.png\" style=\"float:left;padding-left:12px;\"/>" +
          "<span style=\"float:left;padding-top:14px;\">" + PhraseLib.Lookup("term.excludedprodcheckbox", LanguageID) + "</span/>" + "<br clear=\"all\"/>";
        gridView.Columns.Add(cfield);
      }
    }
  }
  private void RefreshGrid()
  {
    TotalCount = hdnInclProducts.Value.ConvertToInt32();
    AMSResult<DataTable> dtResult = new AMSResult<DataTable>();
    string strExcludedProductIDs = "";
    exclList = m_Product.GetExcludedProducts(ProductGroupID).Result;
    AMSResult<DataTable> GridViewCacheData = GetGridiViewCacheData(out strExcludedProductIDs, exclList.Count);
    if (GridViewCacheData.ResultType == AMSResultType.Success)
    {
      if (ProductIDExcluded.Count > 0 && (exclList != null && exclList.Count > 0))
        exclList.RemoveAll(p => !ProductIDExcluded.Contains(p.ProductID.ToString()));
      if (gvData.Rows.Count <= 0 || IsFilterUpdated)
      {
        if (PageIndex == 0)
        {
          hdnTotalProducts.Value = hdnTotalRecords.Value = TotalCount.ConvertToString();
        }

        gvData.Columns.Clear();

        if (GridViewCacheData.Result.Rows.Count > 0)
          GridViewData = (from row in GridViewCacheData.Result.AsEnumerable()
                          where row.Field<Int64>("RowNum") > 0 && row.Field<Int64>("RowNum") <= PageSize
                          select row).CopyToDataTable();
        else
          GridViewData = GridViewCacheData.Result;

        if (GridViewData.Rows.Count >= TotalCount)
          lbNeedReload.Text = "False";
        else
          lbNeedReload.Text = "True";
        btnTemp.Text = "0";
        GenerateColumn(GridViewData, gvData);
        gvData.DataSource = GridViewData;
        gvData.DataBind();

      }
      else
      {
        GenerateColumn(GridViewData, gvData);
        gvData.DataSource = GridViewData;
        gvData.DataBind();
      }

    }
    else
    {
      BindGridHeader();
      SetProductCountInputs("0", string.Empty, string.Empty);
    }
  }

  private AMSResult<DataTable> GetGridiViewCacheData(out String strExcludedProductIDs, int excludeCount)
  {
    strExcludedProductIDs = "";
    AMSResult<DataTable> GridViewCacheData = new AMSResult<DataTable>();
    List<AttributeValue> lstAttributeValue = DTAppliedAttributeValuesViewState.ToGenericList<AttributeValue>();
    if (lstAttributeValue.Count > 0)
    {
      if (IsFilterUpdated)
      {
        GridViewCheckedItems();
        strExcludedProductIDs = String.Join(",", ProductIDExcluded.ToArray());
      }
      if (DTAppliedAttributeValuesViewState.Rows.Count > 0)
      {
        DataTable dt = DTAppliedAttributeValuesViewState.DefaultView.ToTable(false, "AttributeSetID", "AttributeTypeID", "AttributeValueID", "ExcludeFlag");
        GridViewCacheData = m_Product.GetProductsByNodeAndAttributeValues(AllChildNodes, PageIndex, _sortBy, _sortOrder, dt, AttributeTypeToDisplayViewState, ProductGroupID, strExcludedProductIDs, IsFilterUpdated, out TotalCount);
      }
    }
    else
    {
      GridViewCheckedItems();
      strExcludedProductIDs = String.Join(",", ProductIDExcluded.ToArray());
      GridViewCacheData = m_Product.GetProductsByNode(SelectedNodeIDs, PageIndex, (PageSize * 20), _sortBy, _sortOrder, AttributeTypeToDisplayViewState, ProductGroupID, strExcludedProductIDs, IsFilterUpdated, out TotalCount);
    }
    if (Count == 0)
    {
      Count = 2;
    }
    //Set the inputs for product count to be updated during page load on client side if count is not available already
    int tempCount = -1;
    if (Int32.TryParse(hdnInclProducts.Value, out tempCount) && tempCount >= 0)
      SetProductCountInputs(hdnInclProducts.Value, string.Empty, string.Empty);
    else
      SetProductCountInputs(TotalCount.ToString(), excludeCount.ToString(), string.Empty);

    this.GridViewCacheData = GridViewCacheData;
    return GridViewCacheData;
  }
  private void ResolveDependencies()
  {
    CurrentRequest.Resolver.AppName = this.AppName;
    PhraseLib = CurrentRequest.Resolver.Resolve<IPhraseLib>();
    m_Logger = CurrentRequest.Resolver.Resolve<ILogger>();
    m_common = CurrentRequest.Resolver.Resolve<CMS.AMS.Common>();
    m_Offer = CurrentRequest.Resolver.Resolve<IOffer>();
    m_ErrorHandler = CurrentRequest.Resolver.Resolve<IErrorHandler>();
    m_Product = CurrentRequest.Resolver.Resolve<IProductService>();
    m_Attribute = CurrentRequest.Resolver.Resolve<IAttributeService>();
    m_ProductGroup = CurrentRequest.Resolver.Resolve<IProductGroupService>();
    m_common.Open_LogixRT();
  }
  private void BindAttributeSetDDL()
  {
    ddlAttributeSet.Items.Clear();
    AddDefaultTextToSetIDSelector();
    if (DTAppliedAttributeValuesViewState != null && DTAppliedAttributeValuesViewState.Rows.Count > 0)
    {
      DataTable sortedAppliedSetDT = (DTAppliedAttributeValuesViewState.AsEnumerable().OrderBy(x => x["AttributeSetID"])).CopyToDataTable();
      foreach (DataRow row in sortedAppliedSetDT.Rows)
      {
        string setID = row["AttributeSetID"].ToString();

        if (!ddlAttributeSet.Items.Contains(new ListItem(setID)))
        {
          bool exclusionExistsForSet = (sortedAppliedSetDT.AsEnumerable()
                                      .Count(r => r["AttributeSetID"].ToString() == setID && r["ExcludeFlag"].ConvertToBool() == true)) == 0 ? false : true;
          if (!exclusionExistsForSet)
            ddlAttributeSet.Items.Add(setID);
        }
      }
    }
  }
  private DataTable CreateSelectedSetDT()
  {
    DataTable dtCustomTable = new DataTable();
    dtCustomTable.Columns.Add("AttributeTypeID", typeof(int));
    dtCustomTable.Columns.Add("AttributeValue", typeof(string));
    dtCustomTable.Columns.Add("ToolTip", typeof(string));
    dtCustomTable.Columns.Add("AttributeTitle", typeof(string));
    //Purpose of AttributeSetID here is only for passing command argument in the control which helps during edit in repWrapMultiSet_ItemCommand 
    //because the datasource and dataitem and the setid label text is not available
    dtCustomTable.Columns.Add("AttributeSetID", typeof(string));
    return dtCustomTable;
  }
  private DataTable GetDataSourceForSelectedSetUI(DataTable dtSelectedSet)
  {
    DataTable dtCustom = new DataTable();
    if (dtSelectedSet.Rows.Count > 0)
    {
      //Get groups for each attribute typeid
      var drAll2 = from attribute in dtSelectedSet.AsEnumerable()
                   group attribute by attribute.Field<Int16>("AttributeTypeID")
                     into g
                     select new { AttributeTypeID = g.Key, Group = g };

      dtCustom = CreateSelectedSetDT();
      foreach (var item in drAll2)
      {
        IGrouping<short, DataRow> group = item.Group;
        DataRow customRow = dtCustom.NewRow();
        GetCustomizedAVDataRow(group, customRow);
        //if(radExcludedSet.Checked)
        //    GetCustomizedAVDataRow(group, customRow, true);
        //else
        //    GetCustomizedAVDataRow(group, customRow, false);
        dtCustom.Rows.Add(customRow);
      }
    }
    lblFilter.Visible = true;
    return dtCustom;
  }
  /// <summary>
  /// Function to customize few columns for repeater going through each passed group 
  /// </summary>
  /// <param name="group"></param>
  /// <param name="customRow"></param>
  private void GetCustomizedAVDataRow<T>(IGrouping<T, DataRow> group, DataRow customRow)
  {
    //DataRow customRow = new DataRow();
    string strValues = "";
    int tmpAttributeTypeID = 0;
    string strToolTip = "";
    string strTitle = "";
    int counter = 0;
    int setID = 0;

    foreach (DataRow i in group)
    {
      setID = i["AttributeSetID"].ConvertToInt16();
      strTitle = i["Name"].ToString();
      strToolTip = strToolTip + ", " + i["value"].ToString();
      tmpAttributeTypeID = i["AttributeTypeID"].ConvertToInt32();
      if (counter < 1)
      {
        strValues = strValues + i["value"].ToString();
      }
      else if (counter == 1)
      {
        strValues = strValues + "...";
      }
      counter++;
    }
    if (strValues.Length > 10)
    {
      strValues = strValues.Substring(0, 10) + "...";
    }
    if (strTitle.Length > 12)
    {
      strTitle = strTitle.Substring(0, 12) + "...";
    }
    customRow["AttributeTypeID"] = tmpAttributeTypeID;
    customRow["AttributeValue"] = strValues;
    customRow["ToolTip"] = strToolTip.TrimStart(',');
    customRow["AttributeTitle"] = strTitle;
    customRow["AttributeSetID"] = setID;
  }
  private DataTable CreateMultiSetDT()
  {
    DataTable dtCustomTable = new DataTable();
    dtCustomTable.Columns.Add("AttributeSetID", typeof(Int16));
    dtCustomTable.Columns.Add("AttributeIncludeDT", typeof(DataTable));
    dtCustomTable.Columns.Add("AttributeExcludeDT", typeof(DataTable));
    dtCustomTable.Columns.Add("ExcludeString", typeof(string));
    dtCustomTable.Columns.Add("ExcludeSetStyle", typeof(string));
    dtCustomTable.Columns.Add("JoinStyle", typeof(string));
    dtCustomTable.Columns.Add("JoiningString", typeof(string));
    dtCustomTable.PrimaryKey = new DataColumn[] { dtCustomTable.Columns[0] };
    return dtCustomTable;
  }

  private DataTable GetDataSourceForMultiAttributeSetBinding(DataTable dtAppliedSets)
  {
    string excludePhrase = PhraseLib.Lookup("pab.excluding", LanguageID, "Phrase not found");
    DataTable dtIncludeSet = CreateSelectedSetDT();
    DataTable dtExcludeSet = CreateSelectedSetDT();
    DataTable dtMultiSetCustom = CreateMultiSetDT();
    string styleMargin = "margin-left:28px; margin-right: 5px; margin-bottom: 5px;";
    string styleMarginForJoin = "margin-top:15px; margin-left:5px; margin-right: 5px;";
    string styleDisplay = "display:block;";
    string styleHide = "display:none;";

    if (dtAppliedSets.Rows.Count > 0)
    {
      //Get groups for each attributesetid and typeid
      var drAll2 = (from avRow in dtAppliedSets.AsEnumerable()
                    group avRow by new
                    {
                      AttributeSetID = avRow["AttributeSetID"],
                      AttributeTypeID = avRow["AttributeTypeID"],
                      ExcludeFlag = avRow["ExcludeFlag"],
                    }
                      into g
                      select new { AttributeSetID = g.Key.AttributeSetID, AttributeTypeID = g.Key.AttributeTypeID, ExcludeFlag = g.Key.ExcludeFlag, Group = g })
                   .OrderBy(r => r.AttributeSetID)
                   .ThenBy(s => s.ExcludeFlag);

      int index = 0;
      foreach (var item in drAll2)    //Foreach attribute setid and typeid
      {
        int nextSetID = 0;
        bool excludeFlag = item.ExcludeFlag.ConvertToBool();
        IGrouping<object, DataRow> group = item.Group;
        DataRow customRow = excludeFlag ? dtExcludeSet.NewRow() : dtIncludeSet.NewRow();
        GetCustomizedAVDataRow(group, customRow);
        if (excludeFlag)
          dtExcludeSet.Rows.Add(customRow);
        else
          dtIncludeSet.Rows.Add(customRow);

        //Prepare datatable containing columns - setid and other datatable containing all attributeid\values for a set
        var nextElement = drAll2.ElementAtOrDefault(index + 1);
        if (nextElement != null)
        {
          nextSetID = nextElement.AttributeSetID.ConvertToInt16();
          if (nextSetID > item.AttributeSetID.ConvertToInt16())
          {
            string excludeString = excludeFlag ? excludePhrase : string.Empty; //PhraseLib.Lookup("", LanguageID, "Phrase not found");//Less no of executions inside this statement
            string excludeStyle = string.Empty;
            if (dtExcludeSet.Rows.Count > 0)
              excludeStyle = styleMargin + styleDisplay;
            else
              excludeStyle = styleHide;

            DataRow existingRow = dtMultiSetCustom.Rows.Find(item.AttributeSetID);
            if (existingRow != null)
            {
              existingRow["AttributeExcludeDT"] = dtExcludeSet;
              existingRow["ExcludeString"] = excludeString;
              existingRow["ExcludeSetStyle"] = excludeStyle;
              existingRow["JoinStyle"] = styleMarginForJoin;
            }
            else
              dtMultiSetCustom.Rows.Add(item.AttributeSetID, dtIncludeSet, dtExcludeSet, excludeString, excludeStyle, styleMarginForJoin, PhraseLib.Lookup("term.or", LanguageID, "Phrase not found"));

            dtIncludeSet = CreateSelectedSetDT();
            dtExcludeSet = CreateSelectedSetDT();
          }
        }
        else
        {
          if (nextSetID == 0) //true for last or only one set
          {
            string excludeString = excludeFlag ? excludePhrase : string.Empty; //PhraseLib.Lookup("", LanguageID, "Phrase not found");
            string excludeStyle = string.Empty;
            if (dtExcludeSet.Rows.Count > 0)
              excludeStyle = styleMargin + styleDisplay;
            else
              excludeStyle = styleHide;

            DataRow existingRow = dtMultiSetCustom.Rows.Find(item.AttributeSetID);
            if (existingRow != null)
            {
              existingRow["AttributeExcludeDT"] = dtExcludeSet;
              existingRow["ExcludeString"] = excludeString;
              existingRow["ExcludeSetStyle"] = excludeStyle;
            }
            else
              dtMultiSetCustom.Rows.Add(item.AttributeSetID, dtIncludeSet, dtExcludeSet, excludeString, excludeStyle);
          }
        }
        index++;
      }
    }
    lblFilter.Visible = true;
    return dtMultiSetCustom;
  }
  private void GridViewCheckedItems()
  {
    int counter = 0;
    ProductIDExcluded.Clear();
    ProductIDInculded.Clear();
    foreach (GridViewRow row in gvData.Rows)
    {
      CheckBox chk = row.Cells[1].Controls[0] as CheckBox;
      if (chk != null && chk.Checked)
        ProductIDExcluded.Add(gvData.DataKeys[counter].Values["ProductID"].ToString());
      else
        ProductIDInculded.Add(gvData.DataKeys[counter].Values["ProductID"].ToString());
      counter++;
    }
  }

  public string GetSelectedNodeIDs()
  {

    string NodeID = "";
    bool LoadChildNode = false;
    //When this function called from aspx directoly, forcibly resolve depencency since load does not call.
    if (m_Product == null)
    {
      ResolveDependencies();
      //m_Product = CurrentRequest.Resolver.Resolve<IProductService>();
    }

    //Try to get it from Hierarchy tree in screen1 if any changes are made
    if (_SelectedNodeIDs != string.Empty)
    {
      NodeID = _SelectedNodeIDs.Replace("N", "");
    }
    //if no changes made to hierarchy selction and you might be working on saved productgroup; so use hidden variable
    if (NodeID == "" && PABStage == "2")
    {
      NodeID = hdnSelctedNodeIDs.Value;
      if (Session["CurrentNodes"] == null)
      {
        CurrentNodes = hdnSelctedNodeIDs.Value;
        PrevSelectedNodeIDs = CurrentNodes;
        LoadChildNode = true;
        //Update the grid type and rouped on column name when ever node is changed 
        IsGroupGrid = GetGridType();
      }
      else
      {
        if (PrevSelectedNodeIDs != NodeID)
        {
          CurrentNodes = hdnSelctedNodeIDs.Value;
          PrevSelectedNodeIDs = CurrentNodes;
          LoadChildNode = true;
          //Update the grid type and rouped on column name when ever node is changed 
          IsGroupGrid = GetGridType();
        }
      }
    }
    //If nodeid is not found then get from db - this case occurs when we hit save button
    if (NodeID == string.Empty)
    {

      NodeID = m_ProductGroup.GetNodeIDsOfSelectedNodes(ProductGroupID).Result;
      LoadChildNode = true;
    }
    if (LoadChildNode)
    {

      AllChildNodes = m_Product.GetAllChildNodes(NodeID).Result;
      AttributeTypes = null;
    }
    return NodeID;
  }

  private void MergeProductsCurrentExcludedAndDBExcluded()
  {
    List<Product> pro = new List<Product>();
    GridViewCheckedItems();
    //Save the Excluded items as iteself even if they are not loaded on grid view
    pro = m_Product.GetExcludedProducts(ProductGroupID).Result;
    pro.RemoveAll(p => ProductIDInculded.Contains(p.ProductID.ToString()));
    //Remove all the excluded 
    pro.RemoveAll(p => ProductIDExcluded.Contains(p.ProductID.ToString()));
    //update the excluded list with tha latest list we have 
    ProductIDExcluded.AddRange(pro.Select(i => i.ProductID.ToString()).ToList<string>());
  }
  private void updateProductGroupwithNodesAndAttribute()
  {
    DataTable dt = DTAppliedAttributeValuesViewState.DefaultView.ToTable(false, "AttributeSetID", "AttributeTypeID", "AttributeValueID", "ExcludeFlag");
    string includeprods, excludeprods;
    ResolveDependencies();
    bool IsExcludeAllChecked = false;
    if (!IsGroupGrid)
    {
      if (hdnConsiderExclusions.Value == "1")
      {
        MergeProductsCurrentExcludedAndDBExcluded();
      }
      else
      {
        GridViewCheckedItems();
      }

      if (hdnIsMasterCBChecked.Value == "1")
      {
        GridViewCheckedItems();
        IsExcludeAllChecked = true;
      }
      else if (hdnIsMasterCBChecked.Value == "0")
      {
        GridViewCheckedItems();
      }
      includeprods = String.Join(",", ProductIDInculded.ToArray());
      excludeprods = String.Join(",", ProductIDExcluded.ToArray());
      if (IsExcludeAllChecked)
        m_ProductGroup.UpdateAttributeProductGroups(ProductGroupID, SelectedNodeIDs, PABStage == "2" ? dt : GenerateBlankDTForSaving(), includeprods, IsExcludeAllChecked);
      else
        m_ProductGroup.UpdateAttributeProductGroups(ProductGroupID, SelectedNodeIDs, PABStage == "2" ? dt : GenerateBlankDTForSaving(), excludeprods, IsExcludeAllChecked);

      ClearSelectionOnGrid();
    }
    //If group grid, then update the prods based on include items and exclude items..
    else
    {
      if (hdnIsMasterCBChecked.Value == "1")
      {
        IsExcludeAllChecked = true;
      }
      excludeprods = hfExcludedItems.Value;
      includeprods = hfIncludedItems.Value;
      DataTable pgProdInfo = PrepareProductExclusionInfo();
      DataTable pgLevelInfo = PrepareLevelExclusion();

      AMSResult<bool> res = m_ProductGroup.UpdateAttributeProductGroups_Level(ProductGroupID, SelectedNodeIDs,
        PABStage == "2" ? dt : GenerateBlankDTForSaving(), IsExcludeAllChecked, pgLevelInfo, pgProdInfo, IsLevelFilterUpdated);
      //Error case
      if (!res.Result)
      {
          m_Logger.WriteError("Exception while saving: " + res.MessageString);          
      }
    }
  }

  private DataTable PrepareLevelExclusion()
  {
    DataTable pgLevelInfo = new DataTable();
    //Always get the info from hidden field
    if (hfLevelInfoTable.Value != "" && hfLevelInfoTable.Value != "[]")
      LevelGridViewDetails = (DataTable)JsonConvert.DeserializeObject(hfLevelInfoTable.Value, (typeof(DataTable)));
    if (LevelGridViewDetails != null)
    {
      pgLevelInfo = LevelGridViewDetails.DefaultView.ToTable(false, "ID", "DisplayLevel", "Excluded", "ConsiderLastState");
      DataColumn newColumn1 = new DataColumn("ProductGroupID", typeof(System.Int32));
      newColumn1.DefaultValue = ProductGroupID;
      pgLevelInfo.Columns.Add(newColumn1);
      newColumn1.SetOrdinal(1);
      pgLevelInfo.Columns["ID"].ColumnName = "Pkid";
      pgLevelInfo.Columns["DisplayLevel"].ColumnName = "LevelID";
      pgLevelInfo.Columns["Excluded"].ColumnName = "ExcludFlag";
      pgLevelInfo.Columns["ConsiderLastState"].ColumnName = "PreviouesState";
    }
    return pgLevelInfo;
  }
  private DataTable PrepareProductExclusionInfo()
  {
    DataTable includeDT = new DataTable();
    DataTable excludeDT = new DataTable();
    if (hfIncludedItems.Value != "" && hfIncludedItems.Value != "[]")
      includeDT = (DataTable)JsonConvert.DeserializeObject(hfIncludedItems.Value, (typeof(DataTable)));
    if (hfExcludedItems.Value != "" && hfExcludedItems.Value != "[]")
      excludeDT = (DataTable)JsonConvert.DeserializeObject(hfExcludedItems.Value, (typeof(DataTable)));
    if (includeDT.Rows.Count == 0)
      includeDT = CreateProdTable(excludeDT);
    if (excludeDT.Rows.Count == 0)
      excludeDT = CreateProdTable(includeDT);

    DataColumn newColumn1 = new DataColumn("ProductGroupID", typeof(System.Int32));
    newColumn1.DefaultValue = ProductGroupID;
    includeDT.Columns.Add(newColumn1);
    newColumn1.SetOrdinal(0);
    newColumn1 = new DataColumn("ExcludeFlag", typeof(System.Boolean));
    newColumn1.DefaultValue = false;
    includeDT.Columns.Add(newColumn1);


    newColumn1 = new DataColumn("ProductGroupID", typeof(System.Int32));
    newColumn1.DefaultValue = ProductGroupID;
    excludeDT.Columns.Add(newColumn1);
    newColumn1.SetOrdinal(0);
    newColumn1 = new DataColumn("ExcludeFlag", typeof(System.Boolean));
    newColumn1.DefaultValue = true;
    excludeDT.Columns.Add(newColumn1);

    excludeDT.Merge(includeDT);

    DataColumn newColumn = new DataColumn("Pkid", typeof(System.Int32));
    excludeDT.Columns.Add(newColumn);
    newColumn.SetOrdinal(0);
    return excludeDT;
  }

  private DataTable prepareGridviewtable()
  {
    DataTable dt = new DataTable();
    DataColumn[] cols = { new DataColumn("ID", typeof(Int32)), new DataColumn("DisplayLevel", typeof(Int32)), 
                              new DataColumn("Excluded", typeof(Int32)), new DataColumn("ConsiderLastState", typeof(Int32)) };
    dt.Columns.AddRange(cols);
    return dt;
  }

  private DataTable CreateProdTable(DataTable tempdt)
  {
    DataTable dt = new DataTable();
    DataColumn[] cols;
    if (tempdt.Rows.Count == 0)
    {
      DataColumn[] tempcols = { new DataColumn("LevelID", typeof(String)), 
                              new DataColumn("ProductID", typeof(Int32)) };
      cols = tempcols;
    }
    else
    {

      DataColumn[] tempcols = { new DataColumn("LevelID", typeof(String)), 
                              new DataColumn("ProductID", tempdt.Columns["ProductID"].DataType) };
      cols = tempcols;
    }
    dt.Columns.AddRange(cols);
    return dt;
  }
  private string GetAttributeValueIDs()
  {
    string output = string.Empty;
    DataTable dtTemp = new DataTable();
    if (DTAppliedAttributeValuesViewState as DataTable != null)
    {
      dtTemp = DTAppliedAttributeValuesViewState;
      dtTemp.DefaultView.Sort = "AttributeValueID";
      dtTemp = dtTemp.DefaultView.ToTable();
    }

    for (int i = 0; i < dtTemp.Rows.Count; i++)
    {
      output = output + dtTemp.Rows[i]["AttributeValueID"].ToString();
      output += (i < (dtTemp.Rows.Count - 1)) ? "," : string.Empty;
    }
    return output;
  }
  private void loadphrasesForTexts()
  {
    btnClear.Text = PhraseLib.Lookup("term.clear", LanguageID, "Phrase Not Found") + " " + PhraseLib.Lookup("term.Set", LanguageID, "Phrase Not Found");
    btnApplyFilter.Text = PhraseLib.Lookup("term.apply_attr_set", LanguageID, "Phrase Not Found");
    hlBack.Text = PhraseLib.Lookup("term.bck2attr_selection", LanguageID, "Phrase Not Found");
    lblFilter0.Text = PhraseLib.Lookup("term.include_prod_with_attr", LanguageID, "Phrase Not Found");
    hlShowDetails.Text = PhraseLib.Lookup("term.view_detail_pglist", LanguageID, "Phrase Not Found");
    radExcludedSet.Text = PhraseLib.Lookup("pab.excludeset", LanguageID, "Phrase Not Found");
    radIncludedSet.Text = PhraseLib.Lookup("pab.createincludedset", LanguageID, "Phrase Not Found");
  }

  //dt.Columns.Add("ID", typeof(int)).SetOrdinal(0);
  //calculate the total number of products inside tha heirarchy nd update the count
  private void LoadHierarchyTree()
  {
    if (string.IsNullOrEmpty(HierarchyTreeURL))
    {
      return;
    }

    string url = HierarchyTreeURL + "&BuyerID=" + ((BuyerID == null || BuyerID <= 0) ? -1 : BuyerID);
    if (m_Product == null)
    {
      ResolveDependencies();
    }

    long ParentID = 0;
    long HierarchyID = 0;
    if (LocateHierarchyURL != string.Empty)
    {
      //Trying to locate item in Hierarchy Tree
      url += LocateHierarchyURL + "&SelectedNodeIDs=-1";
      hdnPABStage.Value = "1";
      hdnSelctedNodes.Value = "";
      hdnSelctedNodeIDs.Value = "";
      hdnLocateHierarchyURL.Value = "";

    }
    else if (!IsPostBack || IsAttributeSwitch)
    {
      //first time page is loaded or it's a switch from standard group to Attribute group
      //if there are selected nodes already then load the hierarchytree and selected nodes should be checked by default
      if (ProductGroupID > 0)
      {
        PHNode phnode = m_ProductGroup.GetTopPHNode(ProductGroupID).Result;
        if (phnode != null)
        {
          if (phnode.ParentID > 0)
          {
            url += "&Selected=L1ID" + phnode.ParentID.ToString();
          }
          else
          {
            url += "&Selected=L0ID" + phnode.HierarchyID;
          }
          ParentID = phnode.ParentID;
          HierarchyID = phnode.HierarchyID;
        }

        hdnSelctedNodes.Value = m_ProductGroup.GetExternalIDsOfSelectedNodes(ProductGroupID,HierarchyID).Result;
        hdnSelctedNodeIDs.Value = m_ProductGroup.GetNodeIDsOfSelectedNodes(ProductGroupID).Result;

        if (hdnSelctedNodes.Value != string.Empty)
        {
          hdnPABStage.Value = "2";
          url += "&SelectedNodeIDs=" + hdnSelctedNodeIDs.Value;
        }
        else
        {
          hdnPABStage.Value = "1";
          url += "&SelectedNodeIDs=-1";
        }
      }
      else
      {
        hdnPABStage.Value = "1";
        hdnSelctedNodes.Value = "";
        hdnSelctedNodeIDs.Value = "";
        url += "&SelectedNodeIDs=-1";
      }
    }
    else
    {
      //Trying to edit the selected nodes-->Set the stage to 1 and reload hierarchytree with selected nodes checked
      hdnPABStage.Value = "1";
      if (hdnSelctedNodes.Value != string.Empty)
      {
        PHNode phnode = m_ProductGroup.GetPHNode(Convert.ToInt32(hdnSelctedNodeIDs.Value.Split(',')[0])).Result;
        if (phnode != null)
        {
          if (phnode.ParentID > 0)
          {
            url += "&Selected=L1ID" + phnode.ParentID.ToString();
          }
          else
          {
            url += "&Selected=L0ID" + phnode.HierarchyID;
          }

          ParentID = phnode.ParentID;
          HierarchyID = phnode.HierarchyID;
        }
        url += "&SelectedNodeIDs=" + hdnSelctedNodeIDs.Value;
      }
      else
      {
        url += "&SelectedNodeIDs=-1";
      }
    }

    if (url != "")
    {
      if (hdnPABStage.Value == "2")
      {
        divphselectedtree.InnerHtml = LoadSelectedNodesHierarchy(hdnSelctedNodes.Value, ParentID, HierarchyID);
        hdndivphselectedtree.Value = Microsoft.JScript.GlobalObject.escape(divphselectedtree.InnerHtml);
      }
      else
      {
        LoadBuyerNodesintoSession(BuyerID);
        writer = new System.IO.StringWriter();
        HttpContext.Current.Server.Execute(url, writer, true);
        producthierarchy.InnerHtml = writer.ToString();

      }
    }
  }

  private string GetDisplayText()
  {

    string systemOption62 = m_common.Fetch_SystemOption(62);
    string displayText = "";

    switch (systemOption62)
    {
      case "0":
        displayText = "PH1.Name ";
        break;
      case "1":
        displayText = "   case  " +
              "       when PH1.ExternalID is NULL then PH1.Name " +
              "       when PH1.ExternalID = '' then PH1.Name " +
              "       when PH1.ExternalID not like '%' + PH1.Name + '%' then PH1.ExternalID + '-' + PH1.Name " +
              "       else PH1.ExternalID " +
              "   end ";
        break;
      case "2":
        displayText = "   case  " +
              "       when PH1.DisplayID is NULL or PH1.DisplayID='' then PH1.Name " +
              "       else PH1.DisplayID + '-' + PH1.Name " +
              "   end ";
        break;
      default:
        displayText = "PH1.Name ";
        break;
    }

    return displayText;

  }
  private string LoadSelectedNodesHierarchy(string selctedNodeNames, long ParentNodeID, long HierarchyID)
  {
    StringBuilder html = new StringBuilder("");
    string tooltip = PhraseLib.Lookup("term.edithierarchyselection", LanguageID);
    int NewLeft = 0;

    string displayText = GetDisplayText();

    DataTable Nodes = new DataTable();

    if (ParentNodeID > 0)
    {
      Nodes = m_ProductGroup.GetProductHierarchyByNodeID(ParentNodeID, displayText).Result;
    }

    //Top Node
    html.Append("<br />");
    html.Append("<img src=\"/images/clear.png \" style=\"height:1px;width:" + NewLeft.ToString() + "px; \" />");
    html.Append("<span onclick=\"return WarnUser();\" class=\"hrow\" title=\"" + tooltip + "\" onmouseover=\"highlightdiv(true)\" onmouseout=\"highlightdiv(false)\" ><img  border=\"0\" src=\"/images/hierarchy.png\" /><span style=\"left: 5px;\">&nbsp;" + PhraseLib.Lookup("term.producthierarchies", LanguageID) + "</span></span><br />");
    //Hierarchy Node
    string HierarchyName = m_ProductGroup.GetHierarchyName(HierarchyID, displayText).Result;
    NewLeft += 17;
    html.Append("<img src=\"/images/clear.png \" style=\"height:1px;width:" + NewLeft.ToString() + "px; \" />");
    html.Append("<span onclick=\"return WarnUser();\" class=\"hrow\" title=\"" + tooltip + "\" onmouseover=\"highlightdiv(true)\" onmouseout=\"highlightdiv(false)\" ><img  border=\"0\" src=\"/images/hierarchy.png\" /><span style=\"left: 5px;\">&nbsp;" + HierarchyName + "</span></span><br />");

    //Inner Nodes
    foreach (DataRow row in Nodes.AsEnumerable().Reverse<DataRow>())
    {
      NewLeft += 17;
      html.Append("<img src=\"/images/clear.png \" style=\"height:1px;width:" + NewLeft.ToString() + "px; \" />");
      html.Append("<span onclick=\"return WarnUser();\" class=\"hrow\" title=\"" + tooltip + "\" onmouseover=\"highlightdiv(true)\" onmouseout=\"highlightdiv(false)\" ><img  border=\"0\" src=\"/images/folder.png\" /><span style=\"left: 5px;\">&nbsp;" + row["NodeName"].ToString() + "</span></span><br />");

    }
    //Selcted Nodes
    NewLeft += 17;
    foreach (string item in selctedNodeNames.Split(','))
    {
      if (item != string.Empty)
      {
        html.Append("<img src=\"/images/clear.png \" style=\"height:1px;width:" + NewLeft.ToString() + "px; \" />");
        html.Append("<span onclick=\"return WarnUser();\" class=\"hrow\" title=\"" + tooltip + "\" onmouseover=\"highlightdiv(true)\" onmouseout=\"highlightdiv(false)\" ><img  border=\"0\" src=\"/images/folder.png\" /><span style=\"left: 5px;\">&nbsp;" + item + "</span></span><br />");
      }

    }

    return html.ToString();

  }
  private void AddDefaultTextToSetIDSelector()
  {
    if (ddlAttributeSet.Items.Count == 0)
      ddlAttributeSet.Items.Add(new ListItem(PhraseLib.Lookup(376, LanguageID)));
  }
  private void EnableDisableSelectors(bool appliedSetFlag = false, bool editSetFlag = false, bool excludeSetFlag = false, bool clearSetFlag = false, bool ddlAttSetIndexChangeFlag = false)
  {
    if (editSetFlag)  //Disable all selectors during edit
    {
      radIncludedSet.Enabled = false;
      radExcludedSet.Enabled = false;
      ddlAttributeSet.Enabled = false;

      if (excludeSetFlag)
      {
        radExcludedSet.Checked = true;
        radIncludedSet.Checked = false;
      }
      else
      {
        radExcludedSet.Checked = false;
        radIncludedSet.Checked = true;
      }

      if (ddlAttributeSet.SelectedIndex == 0)
      {
        radExcludedSet.Checked = false;
        radIncludedSet.Checked = true;
      }
    }
    else if (appliedSetFlag) //When set is being applied
    {
      radExcludedSet.Checked = false;
      radIncludedSet.Checked = true;
      radIncludedSet.Enabled = true;
      ddlAttributeSet.Enabled = false;
      if (ddlAttributeSet.Items.Count > 1)
        radExcludedSet.Enabled = true;
      else
        radExcludedSet.Enabled = false;
    }
    else if (clearSetFlag)
    {
      if (ddlAttributeSet.Items.Count > 1)
        radExcludedSet.Enabled = true;
      else
        radExcludedSet.Enabled = false;
      radIncludedSet.Enabled = true;
      ddlAttributeSet.Enabled = false;
      ddlAttributeSet.SelectedIndex = 0;
      radExcludedSet.Checked = false;
      radIncludedSet.Checked = true;
    }
    else if (ddlAttSetIndexChangeFlag)
    {
      if (ddlAttributeSet.SelectedIndex == 0)
      {
        radExcludedSet.Checked = false;
        radIncludedSet.Checked = true;
        ddlAttributeSet.Enabled = false;
      }
      else
        ddlAttributeSet.Enabled = true;

    }
    else
    {
      if (ddlAttributeSet.Items.Count > 1)
        radExcludedSet.Enabled = true;
      else
        radExcludedSet.Enabled = false;
      ddlAttributeSet.Enabled = false;
      radExcludedSet.Checked = false;
      radIncludedSet.Checked = true;
    }
  }

  private void EnableControls()
  {
    ddlAttributeType.Enabled = ddlAttributeType.Items.Count > 0;
    ddlAttributeValue.Enabled = ddlAttributeValue.Items.Count > 0;
    btnAddFilter.Enabled = ddlAttributeType.Enabled && ddlAttributeValue.Enabled;

    if (DTSelectedAttributeValuesViewState != null && DTSelectedAttributeValuesViewState.Rows.Count > 0)
      btnApplyFilter.Enabled = btnClear.Enabled = true;
    else
      btnApplyFilter.Enabled = btnClear.Enabled = false;

    UpdateTemplateProperties();
  }
  private void ShowControls(bool flag)
  {
    attributes.Visible = flag;
    hlShowDetails.Visible = flag;
    attibuteFilters.Visible = flag;
    divphselectedtree.Visible = flag;
    backImg.Visible = hlBack.Visible = !flag;
    DetailedProductList.Visible = !flag;
  }
  /// <summary>
  /// Set the count
  /// </summary>
  /// <param name="matchingProduts"></param>
  /// <param name="firstCall"></param>
  private void SetProductCountMessage(int matchingProducts)
  {
    hdnInclProducts.Value = lbTotalProducts.Text = matchingProducts.ToString();
  }
  private void SetProductCountInputs(string totalProductCount, string exlcudedProductCount, string avPairs)
  {
    //Clear off the field hdnPABAVPairsJson to prevent async call for count and use the count available here
    hdnPABAVPairsJson.Value = avPairs;
    hdnInclProducts.Value = totalProductCount;
    hdnExludedProductsCount.Value = exlcudedProductCount;
  }
  private void setEditBoxHiddenFieldValue()
  {
    Dictionary<string, string> EditKeyValue = new Dictionary<string, string>();
    var EditValuesText = from v in DTAppliedAttributeValuesViewState.AsEnumerable()
                         group v by v.Field<string>("Name") into g
                         select new { Name = g.Key, values = g };
    foreach (var item in EditValuesText)
    {
      string values = "";
      foreach (var i in item.values)
      {
        values += i["value"] + "," + " ";
      }
      values = values.Remove(values.LastIndexOf(','));
      EditKeyValue.Add(item.Name, values);
    }
    EditBoxMessageString.Value = EditKeyValue.ToJSON();
  }
  #endregion
  private void PopulateLevelGV(int pageIndex, int noOfRecord)
  {
    AMSResult<DataTable> dtResult = new AMSResult<DataTable>();
    AMSResult<DataSet> res = new AMSResult<DataSet>();
    string strExcludedProductIDs = "";
    int pageCount = 0, LevelCount = 0;
    //Keep the AttributeTypeToDisplay and AttributeValues in session, for using them in Ajax calls to load grid
    Session["AttributeTypeToDisplay"] = AttributeTypeToDisplayViewState;
    DataTable dt = DTAppliedAttributeValuesViewState.DefaultView.ToTable(false, "AttributeSetID", "AttributeTypeID", "AttributeValueID", "ExcludeFlag");
    Session["AttributeValues"] = dt;
    Session["AllChildNodes"] = AllChildNodes;
    Session["IsLevelFilterUpdated"] = IsLevelFilterUpdated;

    DataTable ExcludeLeveldt = new DataTable();
    DataColumn dc = new DataColumn("DisplayLevel", typeof(String));
    ExcludeLeveldt.Columns.Add(dc);
    if (IsLevelFilterUpdated && LevelGridViewDetails.Rows.Count > 0)
    {
      if (LevelGridViewDetails.Columns["Excluded"].DataType == typeof(Int32))
      {
        LevelGridViewDetails.AsEnumerable()
          .Where(r => (r.Field<Int32?>("Excluded") ?? 0) == 1).ToList()
            .ForEach(row => ExcludeLeveldt.Rows.Add(row["DisplayLevel"]));
      }
      else if (LevelGridViewDetails.Columns["Excluded"].DataType == typeof(Int64))
      {
        LevelGridViewDetails.AsEnumerable()
        .Where(r => (r.Field<Int64?>("Excluded") ?? 0) == 1).ToList()
        .ForEach(row => ExcludeLeveldt.Rows.Add(row["DisplayLevel"]));
      }
      else if (LevelGridViewDetails.Columns["Excluded"].DataType == typeof(bool))
      {
        LevelGridViewDetails.AsEnumerable()
          .Where(r => (r.Field<bool?>("Excluded") ?? false)).ToList()
            .ForEach(row => ExcludeLeveldt.Rows.Add(row["DisplayLevel"]));
      }
    }
    string sortOrder = "ASC", sortBy = "DisplayLevel";
    if (_sortBy == sortBy)
      sortOrder = _sortOrder;

    res = m_Product.GetLevelGroups(PageIndex, 100, sortBy, sortOrder, dt, AttributeTypeToDisplayViewState, ProductGroupID, strExcludedProductIDs,
      IsLevelFilterUpdated, AllChildNodes, ExcludeLeveldt, out LevelCount, out TotalCount);
    if (res.ResultType == AMSResultType.Success)
    {
      gvGroupLevel.DataSource = res.Result.Tables[1];
      gvGroupLevel.DataBind();
      if (hfLevelInfoTable.Value == "")
      {
        LevelGridViewDetails = res.Result.Tables[0];
        //Update the Exclude flag correctly based on exclude count and total count
        LevelGridViewDetails.AsEnumerable()
          .Where(r => r.Field<Int32>("ExcludedCount") != r.Field<Int32>("TotalCount")).ToList()
          .ForEach(row => row["Excluded"] = 0);
        hfLevelInfoTable.Value = JsonConvert.SerializeObject(LevelGridViewDetails);
      }
      //Clear off the field hdnPABAVPairsJson to prevent async call for count and use the count available here
      //hdnPABAVPairsJson.Value = string.Empty;
      int excludeCount = LevelGridViewDetails.AsEnumerable().Sum(r => r["ExcludedCount"].ConvertToInt32());
      //hdnExludedProductsCount.Value = excludeCount.ToString();
      //hdnInclProducts.Value = TotalCount.ToString();
      SetProductCountInputs(TotalCount.ToString(), excludeCount.ToString(), string.Empty);

      if (TotalCount > 0)
      {
        pageCount = (LevelCount / noOfRecord) + ((LevelCount % noOfRecord) > 0 ? 1 : 0);
        hfTotalPage.Value = pageCount.ToString();
        hfPageIndex.Value = "1";//loaded the 0th page, now load from 1st page
      }
    }
    else
    {
      gvGroupLevel.DataSource = null;
      gvGroupLevel.DataBind();
    }
  }
  /// <summary>
  /// Update the Group level header and set the column css.
  /// </summary>
  /// <param name="sender"></param>
  /// <param name="e"></param>
  protected void GroupGrid_RowDataBound(object sender, GridViewRowEventArgs e)
  {
    hfGroupedOn.Value = GroupedOn;
    //For Level header, get the group value and set-varma 
    if (e.Row.RowType == DataControlRowType.Header)
    {
      LinkButton LnkHeaderText = e.Row.Cells[2].Controls[0] as LinkButton;
      LnkHeaderText.Text = GroupedOn;
      foreach (var item in e.Row.Cells)
      {
        if (item is TableCell && (item as TableCell).Controls[0] is LinkButton)
        {
          LnkHeaderText = (item as TableCell).Controls[0] as LinkButton;
          LnkHeaderText.Attributes.Add("onclick", "UpdateProductChanges();");
          if (LnkHeaderText.Text == "ExtProductID")
            LnkHeaderText.Text = PhraseLib.Lookup("term.itemsupc", LanguageID);
          if (LnkHeaderText.Text == "Description")//For Description,dont enable Sorting.
            (item as TableCell).Text = "Description";
          hfItemUPC.Value = PhraseLib.Lookup("term.itemsupc", LanguageID);
          var flag = (((System.Web.UI.WebControls.DataControlFieldCell)(item)).ContainingField.SortExpression == _sortBy) ? true : false;
          if (flag && (!string.IsNullOrEmpty(_sortBy)))
            addSortImage((System.Web.UI.WebControls.DataControlFieldCell)(item));
        }
      }

    }
    //Dont display the RowID, Excluded, ProductID columns
    e.Row.Cells[1].Visible = false;
    e.Row.Cells[3].Visible = false;
    e.Row.Cells[4].Visible = false;
  }

  private void addSortImage(DataControlFieldCell headerCell)
  {
    Literal img = new Literal();
    // Create the sorting image based on the sort direction.
    if (_sortOrder == "ASC")
      img.Text = "&nbsp;<span class='sortarrow'>&#9660;</span>";
    else
      img.Text = "&nbsp;<span class='sortarrow'>&#9650;</span>";

    // Add the image to the appropriate header cell.
    headerCell.Controls.Add(img);
  }
  protected void GroupGrid_Sorting(object sender, GridViewSortEventArgs e)
  {
    _sortOrder = gvGroupLevel.SortOrder;
    _sortBy = gvGroupLevel.SortKey;
    //Update the level details as user might changed some thing in ui.
    if (hfIsGridUpdated.Value == "1")
      LevelGridViewDetails = (DataTable)JsonConvert.DeserializeObject(hfLevelInfoTable.Value, (typeof(DataTable)));
    PopulateLevelGV(0, 100);
    //SetProductCountMessage(Convert.ToInt32(HeirarchyProductCount), false);
    backImg.Visible = hlBack.Visible = true;
  }

  //Update the level item exclude state 
  public string ProcessExcludeDataItem(object DisplayLevel, object IsExcluded)
  {
    string res = "";
    bool excluded;
    //if any modifications happend in the grid, consider grid state
    if (hfIsGridUpdated.Value == "1")
    {
      DataRow[] dr = LevelGridViewDetails.Select("DisplayLevel='" + DisplayLevel.ToString().Replace("'", "''") + "'");
      //If the level state is changed..
      if (dr != null && Convert.ToInt32(dr[0]["ConsiderLastState"]) == 0)
      {
        IsExcluded = dr[0]["Excluded"];
      }

      //If the product inside the level is included after excluding level then the level should not be excluded.
      DataTable includeDT = new DataTable();
      DataRow[] drs = null;
      if (hfIncludedItems.Value != "" && hfIncludedItems.Value != "[]")
      {
        includeDT = (DataTable)JsonConvert.DeserializeObject(hfIncludedItems.Value, (typeof(DataTable)));
        drs = includeDT.Select("LevelID='" + DisplayLevel.ToString().Replace("'", "''") + "'");
      }
      if (IsExcluded.ConvertToBool() && drs != null && drs.Count() > 0)
      {
        IsExcluded = false;
      }
    }
    excluded = (IsExcluded == DBNull.Value) ? false : Convert.ToBoolean(IsExcluded);
    if (excluded)
      res = "checked";
    return res;
  }
}
