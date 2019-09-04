using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using CMS.DB;
using CMS.AMS;
using CMS.AMS.Models;
using CMS.AMS.Contract;
using System.Data;
using System.Web.UI.HtmlControls;

public partial class logix_UserControls_TemplateFieldLockControl : System.Web.UI.UserControl
{
    IDBAccess dbAccess;
    CMS.AMS.Common m_Commondata;
    private bool rewardLockStatus;
    # region Public properties
    public IPhraseLib PhraseLib { get; set; }
    public string PageName { get; set; }
    public int OfferId { get; set; }
    public int LanguageId { get; set; }
    public int DeliverableId { get; set; }
    public bool RewardLockStatus { 
        get { return RewardTemplateFieldSource.DisallowEdit; }
        set
        {
            rewardLockStatus = value;
        }
    } 
    public List<int> ExceptionFields { 
        get
        {
            return (List<int>)ViewState["ExceptionFields"];
        }
        set
        {
            ViewState["ExceptionFields"] = value;
        }
    }

    public RewardTemplateFieldContainer RewardTemplateFieldSource { 
        get 
        {
            return (RewardTemplateFieldContainer)ViewState["RewardTemplateFieldSource"]; 
        }
        set 
        {
            ViewState["RewardTemplateFieldSource"] = value;
        } 
    }
    # endregion Public properties
    # region Control events
    protected void Page_Load(object sender, EventArgs e)
    {
        PhraseLib = CurrentRequest.Resolver.Resolve<IPhraseLib>();
        dbAccess = CurrentRequest.Resolver.Resolve<IDBAccess>();
        m_Commondata = CurrentRequest.Resolver.Resolve<CMS.AMS.Common>();

        if (!IsPostBack)
        {
            ExceptionFields = new List<int>();
            PopulateFields();
        }
    }
    protected void lnkBtnArrow_OnClick(Object sender, EventArgs e)
    {
        HtmlGenericControl divControl = (HtmlGenericControl)this.FindControl("divTemplatefields");
        if (divControl.Style["display"] == "inline")
            divControl.Style["display"] = "none";
        else if (divControl.Style["display"] == "none" || divControl.Style["display"] == null)
            divControl.Style["display"] = "inline";
    }
    protected void repTemplateFields_OnItemDataBound(object sender, RepeaterItemEventArgs riEA)
    {
        CheckBox chk = (CheckBox)riEA.Item.FindControl("chkFieldLock");
        Label lblStatus = (Label)riEA.Item.FindControl("lblStatus");
        CMS.AMS.Models.TemplateField tf = (CMS.AMS.Models.TemplateField)riEA.Item.DataItem;
        if (chk != null)
        {
            if (RewardTemplateFieldSource.DisallowEdit)
                chk.Checked = tf.Editable ? true : false;
            else if (!RewardTemplateFieldSource.DisallowEdit)
                chk.Checked = !tf.Editable ? true : false;
          
        }
    }
    protected void chkFieldLock_OnCheckedChanged(object sender, EventArgs e)
    {
        CheckBox chk = (CheckBox)sender;

        RepeaterItem ri = (RepeaterItem)chk.Parent;
        Label lbl = (Label)ri.FindControl("lblFieldName");
        Label lblStatus = (Label)ri.FindControl("lblStatus");
        CMS.AMS.Models.TemplateField tf = RewardTemplateFieldSource.TemplateFieldList.Single(p => p.FieldName == lbl.Text);
        
        if (chk.Checked)
        {
            if(!ExceptionFields.Contains(tf.FieldId))
             ExceptionFields.Add(tf.FieldId);
            if (!RewardTemplateFieldSource.DisallowEdit)
                tf.Editable = false;
            else
                tf.Editable = true;
        }
        else
        {
            if (ExceptionFields.Contains(tf.FieldId))
                ExceptionFields.Remove(tf.FieldId);
            if (!RewardTemplateFieldSource.DisallowEdit)
                tf.Editable = true;
            else
                tf.Editable = false;
        }
        UpdateFieldStatusString(tf);
        repTemplateFields.DataSource = RewardTemplateFieldSource.TemplateFieldList;
        repTemplateFields.DataBind();
    }
    protected void chkDisallow_Edit_OnCheckedChanged(object sender, EventArgs e)
    {
        if (chkDisallow_Edit.Checked)
        {
            RewardTemplateFieldSource.DisallowEdit = true;
        }
        else
        {
            RewardTemplateFieldSource.DisallowEdit = false;
        }
        ToggleTemplateFields();
    }
    # endregion Control events
# region Private methods
    private RewardTemplateFieldContainer GetTemplateFields()
    {
        DataTable dt;
        CMS.AMS.Models.TemplateField templateField;
        RewardTemplateFieldSource = new RewardTemplateFieldContainer();
        RewardTemplateFieldSource.TemplateFieldList = new List<CMS.AMS.Models.TemplateField>();
        RewardTemplateFieldSource.DisallowEdit = rewardLockStatus;
        m_Commondata = CurrentRequest.Resolver.Resolve<CMS.AMS.Common>();
        PageName = "UEoffer-rew-proximitymsg.aspx";
        dt = m_Commondata.GetAllFieldLevelPermissions(OfferId,DeliverableId, PageName);
        string isoffereditable = m_Commondata.IsOfferEditable(DeliverableId);
        if (isoffereditable == "False")
        {
            RewardTemplateFieldSource.DisallowEdit = true;
        }
        else
        {
            RewardTemplateFieldSource.DisallowEdit = false;
        }
        foreach (DataRow dr in dt.Rows)
        {
            templateField = new CMS.AMS.Models.TemplateField();
            templateField.FieldId = Convert.ToInt32(dr["FieldID"]);
            templateField.FieldName = dr["FieldName"].ToString();
            templateField.Editable = Convert.ToBoolean(dr["Editable"]);
            templateField.ControlName = dr["ControlName"].ToString();
            templateField.Tiered = Convert.ToBoolean(dr["Tiered"]);
            UpdateFieldStatusString(templateField);
            RewardTemplateFieldSource.TemplateFieldList.Add(templateField);
            if (RewardTemplateFieldSource.DisallowEdit && templateField.Editable && !ExceptionFields.Contains(templateField.FieldId))
                ExceptionFields.Add(templateField.FieldId);
            else if (!RewardTemplateFieldSource.DisallowEdit && !templateField.Editable && !ExceptionFields.Contains(templateField.FieldId))
                ExceptionFields.Add(templateField.FieldId);


        }
        return RewardTemplateFieldSource;
    }
    private void PopulateFields()
    {
        RewardTemplateFieldContainer templateRewardContainer = GetTemplateFields();

        chkDisallow_Edit.Checked = templateRewardContainer.DisallowEdit;

        repTemplateFields.DataSource = templateRewardContainer.TemplateFieldList;
        repTemplateFields.DataBind();
    }
    private void UpdateFieldStatusString(CMS.AMS.Models.TemplateField tf)
    {
        tf.StatusString = tf.Editable ? "UnLocked" : "Locked";
       
    }
    private void ToggleTemplateFields()
    {
        ExceptionFields.Clear();
        foreach (CMS.AMS.Models.TemplateField tf in RewardTemplateFieldSource.TemplateFieldList)
        {
            tf.Editable = !tf.Editable;
            if (tf.Editable && RewardTemplateFieldSource.DisallowEdit && !ExceptionFields.Contains(tf.FieldId))
                ExceptionFields.Add(tf.FieldId);
            else if (!RewardTemplateFieldSource.DisallowEdit && !tf.Editable && !ExceptionFields.Contains(tf.FieldId))
                ExceptionFields.Add(tf.FieldId);
            UpdateFieldStatusString(tf);
           
        }
        repTemplateFields.DataSource = RewardTemplateFieldSource.TemplateFieldList;
        repTemplateFields.DataBind();
    }
    # endregion Private methods
}