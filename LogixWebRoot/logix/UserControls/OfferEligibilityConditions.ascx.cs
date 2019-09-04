using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using CMS.AMS;
using CMS.Contract;
using CMS.AMS.Contract;
using CMS.AMS.Models;
using System.Reflection;
using System.Data;
using CMS;
using System.ComponentModel;

public partial class logix_UserControls_OfferEligibilityConditions : System.Web.UI.UserControl
{
    public IPhraseLib PhraseLib
    {
        get;
        set;
    }

    CMS.AMS.Common m_common;

    ICustomerGroupCondition m_CustomerGroupCondition;
    IOfferApprovalWorkflowService m_OAWService;
    IOffer m_Offer;
    IErrorHandler m_ErrorHandler;
    Boolean DeleteBtnEnabled = false;
    bool isTranslatedOffer = false;
    bool bEnableRestrictedAccessToUEOfferBuilder = false;
    Copient.CommonInc MyCommon = null;
    Copient.LogixInc Logix = null;
    bool bOfferEditable = false;
    bool bEnableAdditionalLockoutRestrictionsOnOffers = false;
    

    object offerconditions;
    public string AppName { get; set; }
    public long OfferID { get; set; }
    public int LanguageID { get; set; }
    public int AdminUserID { get; set; }
    public bool IsOptInDisabled { get; set; }
    public bool IsOptInBlockLocked { get; set; }
    protected int CustomerGroupConditionTypeID;
    protected int PointsConditionTypeID;
    protected int SVConditionTypeID;
    protected bool m_Disable;
    protected CMS.AMS.Models.Offer objOffer
    {
        get
        {
            return ViewState["Offer"] as CMS.AMS.Models.Offer;
        }
        set
        {
            ViewState["Offer"] = value;
        }
    }

    [Description("Enables or disables conrols inside UserControl"), Category("Behavior")]
    public bool Disable
    {
        get
        {
            return m_Disable;
        }
        set
        {
            m_Disable = value;
        }
    }

    protected string GetStoredValueDesc(SVCondition program)
    {
        string desc = string.Empty;
        if (program != null)
        {
            desc = program.Quantity + " " + PhraseLib.Lookup("term.units", LanguageID).ToLower() + " ";
            if (program.SVProgram.SVType.SVTypeID != 1)
            {

                desc = desc + "($" + Math.Round(program.Quantity * program.SVProgram.Value, program.SVProgram.SVType.ValuePrecision).ToString() + ") ";
            }

            desc = desc + PhraseLib.Lookup("term.required", LanguageID).ToLower();

        }
        return desc;
    }

    protected void Page_Load(object sender, EventArgs e)
    {
        bool bConnectionOpened = false;
        try
        {

            ResolveDependencies();

            if (!IsPostBack)
            {
                if (MyCommon.LRTadoConn.State == ConnectionState.Closed)
                {
                    MyCommon.Open_LogixRT();
                    bConnectionOpened = true;
                }
                hdnPath.Value = Request.Url.LocalPath + "?offerid=" + OfferID;
                bEnableRestrictedAccessToUEOfferBuilder = m_common.Fetch_SystemOption(249) == "1" ? true : false;
                isTranslatedOffer = MyCommon.IsTranslatedUEOffer(Convert.ToInt32(OfferID), MyCommon);
                bEnableAdditionalLockoutRestrictionsOnOffers  =  m_common.Fetch_SystemOption(260) == "1" ? true : false;
                bOfferEditable = MyCommon.IsOfferEditablePastLockOutPeriod(Logix.UserRoles.EditOfferPastLockoutPeriod, MyCommon, Convert.ToInt32(OfferID));
                LocalizeControl();
                LoadEligibilityConditions();
                LoadConditionTypes();
                btnAdd.Attributes.Add("onclick", "return OpenConditionWindow(0,-1);");
                if (objOffer != null && objOffer.IsOptable)
                {
                    panelEligibilityCondition.Visible = true;
                    chkOptIn.Attributes.Add("onclick", "javascript:OptOutWindow();");
                    chkOptIn.Checked = true;

                }
                else
                {
                    chkOptIn.Checked = false;
                }
                if (!chkOptIn.Checked)
                {
                    btnAdd.Enabled = false;
                    ddlOptInConditions.Enabled = false;
                }
                if (objOffer.IsTemplate == true)
                {
                    spanChkLocked.Visible = true;
                    chkOptInLocked.Checked = IsOptInBlockLocked;
                }
                BindConditionRepeters();
                if (objOffer.FromTemplate == true && IsOptInBlockLocked == true)
                {

                    spanChkLocked.Visible = false;
                    //panelEligibilityCondition.Enabled = false;
                    panelOptIn.Enabled = false;

                }
                if (IsOptInDisabled == true)
                {
                    panelOptIn.Enabled = false;
                }


            }
            //Disable controls
            foreach (RepeaterItem item in repPointConditions.Items)
            {
                var delBtn = (Button)item.FindControl("btnPointsDelete");
                if (delBtn != null)
                    delBtn.Enabled = Disable;
                if (objOffer.EligiblePointsProgramConditions[item.ItemIndex].DisallowEdit == true)
                    delBtn.Enabled = false;
                if ((bEnableRestrictedAccessToUEOfferBuilder && isTranslatedOffer) || (bEnableAdditionalLockoutRestrictionsOnOffers && !bOfferEditable) || m_OAWService.CheckIfOfferIsAwaitingApproval(OfferID).Result)
                {
                    delBtn.Visible = false;
                }
            }
            foreach (RepeaterItem item in repSvConditions.Items)
            {
                var delBtn = (Button)item.FindControl("btnSVDelete");
                if (delBtn != null)
                    delBtn.Enabled = Disable;
                if (objOffer.EligibleSVProgramConditions[item.ItemIndex].DisallowEdit == true)
                    delBtn.Enabled = false;
                if ((bEnableRestrictedAccessToUEOfferBuilder && isTranslatedOffer) || (bEnableAdditionalLockoutRestrictionsOnOffers && !bOfferEditable) || m_OAWService.CheckIfOfferIsAwaitingApproval(OfferID).Result)
                {
                    delBtn.Visible = false;
                }
            }

            btnAdd.Enabled = (Disable && chkOptIn.Checked);
            if ((bEnableRestrictedAccessToUEOfferBuilder && isTranslatedOffer) || (bEnableAdditionalLockoutRestrictionsOnOffers && !bOfferEditable) || m_OAWService.CheckIfOfferIsAwaitingApproval(OfferID).Result)
            {
                btnAdd.Visible = false;
            }
            //panelOptIn.Enabled = true;
            //chkOptIn.Checked = true;

            //}
            //AllowAccess();
            //DeleteBtnEnabled = true;
        }
        catch (Exception excp)
        {
            infobar.Visible = true;
            lblError.Text = m_ErrorHandler.ProcessError(excp);
        }
        finally
        {
            m_common.Close_LogixRT(); 
            if (bConnectionOpened) 
            {
                MyCommon.Close_LogixRT();
            }
        }
    }
    private void LocalizeControl()
    {
        btnAdd.Text = PhraseLib.Lookup("term.add", LanguageID);
        chkOptInLocked.Text = PhraseLib.Lookup("term.locked", LanguageID);
        lblGlobalCondition.Text = PhraseLib.Lookup("offer-con.addglobal", LanguageID) + ":";
        lblTitle.Text = PhraseLib.Lookup("term.optinconditions", LanguageID);
    }

    private void BindConditionRepeters()
    {

        repCustomerConditions.DataSource = GetCustomerConditions();
        repCustomerConditions.DataBind();
        repPointConditions.DataSource = objOffer.EligiblePointsProgramConditions;
        repPointConditions.DataBind();
        repSvConditions.DataSource = objOffer.EligibleSVProgramConditions;
        repSvConditions.DataBind();


        if (repCustomerConditions.Items.Count == 0)
            repCustomerConditions.Visible = false;
        if (repPointConditions.Items.Count == 0)
            repPointConditions.Visible = false;
        if (repSvConditions.Items.Count == 0)
            repSvConditions.Visible = false;
    }
    private object GetCustomerConditions()
    {
        if (objOffer.EligibleCustomerGroupConditions != null)
        {
            var conditions = (from p in objOffer.EligibleCustomerGroupConditions.IncludeCondition
                              select new { ConditionID = objOffer.EligibleCustomerGroupConditions.ConditionID, AndOr = "", ConditionTypeID = objOffer.EligibleCustomerGroupConditions.ConditionTypeID, CustomerConditionDetailsID = p.CustomerConditionDetailsID, Details = p.CustomerGroup.Name, CustomerGroupID = p.CustomerGroupID, Infomation = "", include = true, Locked = objOffer.EligibleCustomerGroupConditions.DisallowEdit })
                              .Union
                              (from p in objOffer.EligibleCustomerGroupConditions.ExcludeCondition
                               select new { ConditionID = objOffer.EligibleCustomerGroupConditions.ConditionID, AndOr = "", ConditionTypeID = objOffer.EligibleCustomerGroupConditions.ConditionTypeID, CustomerConditionDetailsID = p.CustomerConditionDetailsID, Details = p.CustomerGroup.Name, CustomerGroupID = p.CustomerGroupID, Infomation = "", include = false, Locked = objOffer.EligibleCustomerGroupConditions.DisallowEdit });



            return conditions;
        }
        else
            return null;
    }

    private void ResolveDependencies()
    {
        CurrentRequest.Resolver.AppName = this.AppName;
        PhraseLib = CurrentRequest.Resolver.Resolve<IPhraseLib>();
        m_common = CurrentRequest.Resolver.Resolve<CMS.AMS.Common>();

        m_CustomerGroupCondition = CurrentRequest.Resolver.Resolve<ICustomerGroupCondition>();
        m_Offer = CurrentRequest.Resolver.Resolve<IOffer>();
        m_ErrorHandler = CurrentRequest.Resolver.Resolve<IErrorHandler>();
        m_OAWService = CurrentRequest.Resolver.Resolve<IOfferApprovalWorkflowService>();
        MyCommon = new Copient.CommonInc();
        Logix = new Copient.LogixInc();
        Object o = MyCommon;
        Logix.Load_Roles(ref o, AdminUserID);
        m_common.Open_LogixRT();
    }
    private void LoadConditionTypes()
    {
        bool displayCustGroup = true;
        bool displayPoint = true;
        bool displaysv = true;

        CustomerGroupConditionTypeID = m_CustomerGroupCondition.GetCustomerGroupConditionTypeID(objOffer.EngineID);
        PointsConditionTypeID = m_CustomerGroupCondition.GetPointsConditionTypeID(objOffer.EngineID);
        SVConditionTypeID = m_CustomerGroupCondition.GetStoredValueConditionTypeID(objOffer.EngineID);


        if (objOffer.EligibleCustomerGroupConditions != null && objOffer.EligibleCustomerGroupConditions.ConditionID > 0)
        {
            displayCustGroup = false;
        }
        else
        {
            displayPoint = false;
            displaysv = false;
        }

        if (objOffer.EligiblePointsProgramConditions != null && objOffer.EligiblePointsProgramConditions.Count > 0 && objOffer.EngineID == 2)
            displayPoint = false;
        if (objOffer.EligibleSVProgramConditions != null && objOffer.EligibleSVProgramConditions.Count > 0 && objOffer.EngineID == 2)
            displaysv = false;

        List<ConditionType> lstConditionTypes = m_CustomerGroupCondition.GetAllConditionTypes(objOffer.EngineID, objOffer.EngineSubTypeID);
        ddlOptInConditions.DataTextField = "Description";
        ddlOptInConditions.DataValueField = "ConditionTypeID";
        ddlOptInConditions.DataSource = from c in lstConditionTypes
                                        where (c.ConditionTypeID == CustomerGroupConditionTypeID && displayCustGroup) || (c.ConditionTypeID == PointsConditionTypeID & displayPoint) || (c.ConditionTypeID == SVConditionTypeID && displaysv)
                                        select new { ConditionTypeID = c.ConditionTypeID, Description = PhraseLib.Lookup((int)c.PhraseID, LanguageID) };
        //var ConditionTypes = lstConditionTypes.FindAll(c => (c.ConditionTypeID == CustomerGroupConditionTypeID && displayCustGroup) || (c.ConditionTypeID == PointsConditionTypeID & displayPoint) || (c.ConditionTypeID == SVConditionTypeID && displaysv));

        ddlOptInConditions.DataBind();

        if (ddlOptInConditions.Items.Count == 0)
        {
            btnAdd.Enabled = false;
            ddlOptInConditions.Items.Add("No Condition");
            ddlOptInConditions.Enabled = false;

        }

    }
    private void LoadEligibilityConditions()
    {
        objOffer = m_Offer.GetOffer(OfferID, LoadOfferOptions.OfferDetail | LoadOfferOptions.AllEligibilityConditions);

    }

    protected string HideLockedColumn()
    {
        if (objOffer.IsTemplate == false && objOffer.FromTemplate == false)
        {
            return "style='display:none;width=0px'";
        }
        return "";
    }
    private int recordcount = 0;
    protected void repCustomerConditions_ItemDataBound(object sender, System.Web.UI.WebControls.RepeaterItemEventArgs e)
    {
        try
        {
            if (e.Item.DataItem != null)
            {
                //Button 
                dynamic x = e.Item.DataItem;


                //Customer Group Link
                bool include = x.include;
                long CustomerConditionDetailID = x.CustomerConditionDetailsID;
                if (include)
                {
                    if (recordcount > 0)
                    {
                        if (e.Item.FindControl("lblAndOr") != null)
                            ((Label)e.Item.FindControl("lblAndOr")).Text = GetJoinTypeText(JoinTypes.Or);


                    }
                    else
                    {
                        if (e.Item.FindControl("btnCustomerDelete") != null && !bEnableRestrictedAccessToUEOfferBuilder && !isTranslatedOffer && !bEnableAdditionalLockoutRestrictionsOnOffers && bOfferEditable)
                            ((Button)e.Item.FindControl("btnCustomerDelete")).Visible = true; ;
                    }
                    var cusDetail = objOffer.EligibleCustomerGroupConditions.IncludeCondition.Where(p => p.CustomerConditionDetailsID == CustomerConditionDetailID).SingleOrDefault();
                    if (cusDetail.CustomerGroup.CustomerGroupID > 2 && cusDetail.CustomerGroup.NewCardHolders == false)
                    {
                        if (e.Item.FindControl("lnkDetails") != null)
                            e.Item.FindControl("lnkDetails").Visible = true;

                    }
                    else
                    {
                        if (e.Item.FindControl("lblDetails") != null)
                            e.Item.FindControl("lblDetails").Visible = true;
                    }


                }
                else
                {
                    var cusDetail = objOffer.EligibleCustomerGroupConditions.ExcludeCondition.Where(p => p.CustomerConditionDetailsID == CustomerConditionDetailID).SingleOrDefault();
                    if (cusDetail.CustomerGroup.CustomerGroupID > 2 && cusDetail.CustomerGroup.NewCardHolders == false)
                    {
                        if (e.Item.FindControl("lnkDetails") != null)
                        {
                            e.Item.FindControl("lnkDetails").Visible = true;
                            e.Item.FindControl("lblDetails").Visible = true;
                            (e.Item.FindControl("lblDetails") as Label).Text = "excluding ";
                        }
                    }
                    else
                    {
                        if (e.Item.FindControl("lblDetails") != null)
                        {
                            e.Item.FindControl("lblDetails").Visible = true;
                            (e.Item.FindControl("lblDetails") as Label).Text = "excluding " + (e.Item.FindControl("lblDetails") as Label).Text;
                        }
                    }

                }
                if (e.Item.FindControl("lblLocked") != null && recordcount > 0)
                    e.Item.FindControl("lblLocked").Visible = false;
                if (e.Item.FindControl("chkLocked") != null && recordcount > 0)
                    e.Item.FindControl("chkLocked").Visible = false;
                recordcount = recordcount + 1;
            }
        }
        catch (Exception excp)
        {
            infobar.Visible = true;
            lblError.Text = m_ErrorHandler.ProcessError(excp);
        }
        finally
        { }
    }



    private void DeletePointCondition(long ConditionID)
    {
        m_Offer.DeleteOfferEligiblePointsCondition(objOffer.OfferID, objOffer.EngineID, ConditionID);
        m_Offer.UpdateOfferStatusToModified(objOffer.OfferID, objOffer.EngineID, AdminUserID);
        Response.Redirect(hdnPath.Value, false);
    }
    private void ToggelConditionJoinType(long ConditionID, int ConditionTypeID, JoinTypes jointype)
    {
        if (jointype == JoinTypes.And)

            jointype = JoinTypes.Or;

        else
            jointype = JoinTypes.And;

        if (objOffer.EngineID == 2 || objOffer.EngineID == 9 || objOffer.EngineID == 0)
        {
            m_Offer.ToggleOfferEligibleConditionJoinType(objOffer.OfferID, jointype, ConditionTypeID);
        }

        m_Offer.UpdateOfferStatusToModified(objOffer.OfferID, objOffer.EngineID, AdminUserID);
        Response.Redirect(hdnPath.Value, false);
    }
    private void DeleteSVCondition(long ConditionID)
    {
        m_Offer.DeleteOfferEligibleSVCondition(objOffer.OfferID, objOffer.EngineID, ConditionID);
        m_Offer.UpdateOfferStatusToModified(objOffer.OfferID, objOffer.EngineID, AdminUserID);
        Response.Redirect(hdnPath.Value, false);
    }
    protected void repPointConditions_ItemCommand(object source, RepeaterCommandEventArgs e)
    {
        int ConditionTypeID = -1;
        int JoinTypeID = -1;
        try
        {
            Label lbl = e.Item.FindControl("lblConditionTypeID") as Label;
            if (lbl != null)
            {
                ConditionTypeID = lbl.Text.ConvertToInt32();
            }
            lbl = e.Item.FindControl("lblJoinTypeID") as Label;
            if (lbl != null)
            {
                JoinTypeID = lbl.Text.ConvertToInt32();
            }
            switch (e.CommandName)
            {

                case "PointDelete":

                    DeletePointCondition(e.CommandArgument.ConvertToLong());

                    break;
                case "ChangePointJoinType":
                    ToggelConditionJoinType(e.CommandArgument.ConvertToLong(), ConditionTypeID, (JoinTypes)JoinTypeID);
                    break;
                default:
                    break;
            }
        }
        catch (Exception excp)
        {
            infobar.Visible = true;
            lblError.Text = m_ErrorHandler.ProcessError(excp);
        }
    }
    //protected void RepeaterConditions_ItemCommand(object source, RepeaterCommandEventArgs e)
    //{
    //  try
    //  {
    //    switch (e.CommandName)
    //    {

    //      case "DeleteClick":

    //        DeleteCondition(e.CommandArgument.ConvertToLong());

    //        break;
    //      default:
    //        break;
    //    }
    //  }
    //  catch (Exception excp)
    //  {
    //    infobar.Visible = true;
    //    lblError.Text = m_ErrorHandler.ProcessError(excp);
    //  }
    //  finally
    //  { }
    //}
    private int pointsrecordcount = 0;
    protected void repPointConditions_ItemDataBound(object sender, RepeaterItemEventArgs e)
    {

        if (e.Item.DataItem != null)
        {

            CMS.AMS.Models.PointsCondition PointCondition = e.Item.DataItem as CMS.AMS.Models.PointsCondition;
            if (pointsrecordcount > 0)
            {
                if (e.Item.FindControl("lnkJoinType") != null)
                {
                    ((LinkButton)e.Item.FindControl("lnkJoinType")).Text = GetJoinTypeText(PointCondition.JoinTypeID);
                    if ((bEnableRestrictedAccessToUEOfferBuilder && isTranslatedOffer) || (bEnableAdditionalLockoutRestrictionsOnOffers && !bOfferEditable))
                        ((LinkButton)e.Item.FindControl("lnkJoinType")).Enabled = false;
                }

                if (e.Item.FindControl("lblJoinTypeID") != null)
                {
                    ((Label)e.Item.FindControl("lblJoinTypeID")).Text = ((int)PointCondition.JoinTypeID).ToString();

                }

            }
            pointsrecordcount = pointsrecordcount + 1;
        }



        ////Customer Group Link
        //bool include = x.include;
        //long CustomerConditionDetailID = x.CustomerConditionDetailsID;
        //if (include)
        //{


        //  var cusDetail = objOffer.EligibleCustomerGroupConditions.IncludeCondition.Where(p => p.CustomerConditionDetailsID == CustomerConditionDetailID).SingleOrDefault();
        //  if (cusDetail.CustomerGroup.CustomerGroupID > 2 && cusDetail.CustomerGroup.NewCardHolders == false)
        //  {
        //    if (e.Item.FindControl("lnkDetails") != null)
        //      e.Item.FindControl("lnkDetails").Visible = true;

        //  }
        //  else
        //  {
        //    if (e.Item.FindControl("lblDetails") != null)
        //      e.Item.FindControl("lblDetails").Visible = true;
        //  }
        //  pointsrecordcount = pointsrecordcount + 1;

        //}
        //else
        //{
        //  var cusDetail = objOffer.EligibleCustomerGroupConditions.ExcludeCondition.Where(p => p.CustomerConditionDetailsID == CustomerConditionDetailID).SingleOrDefault();
        //  if (cusDetail.CustomerGroup.CustomerGroupID > 2 && cusDetail.CustomerGroup.NewCardHolders == false)
        //  {
        //    if (e.Item.FindControl("lnkDetails") != null)
        //    {
        //      e.Item.FindControl("lnkDetails").Visible = true;
        //      e.Item.FindControl("lblDetails").Visible = true;
        //      (e.Item.FindControl("lblDetails") as Label).Text = "excluding ";
        //    }
        //  }
        //  else
        //  {
        //    if (e.Item.FindControl("lblDetails") != null)
        //    {
        //      e.Item.FindControl("lblDetails").Visible = true;
        //      (e.Item.FindControl("lblDetails") as Label).Text = "excluding " + (e.Item.FindControl("lblDetails") as Label).Text;
        //    }
        //  }

        //}




    }
    protected void repSvConditions_ItemCommand(object source, RepeaterCommandEventArgs e)
    {

        int ConditionTypeID = -1;
        int JoinTypeID = -1;
        try
        {
            Label lbl = e.Item.FindControl("lblConditionTypeID") as Label;
            if (lbl != null)
            {
                ConditionTypeID = lbl.Text.ConvertToInt32();
            }
            lbl = e.Item.FindControl("lblJoinTypeID") as Label;
            if (lbl != null)
            {
                JoinTypeID = lbl.Text.ConvertToInt32();
            }
            switch (e.CommandName)
            {

                case "SVDelete":

                    DeleteSVCondition(e.CommandArgument.ConvertToLong());
                    break;
                case "ChangeSVJoinType":
                    ToggelConditionJoinType(e.CommandArgument.ConvertToLong(), ConditionTypeID, (JoinTypes)JoinTypeID);
                    break;
                default:
                    break;
            }
        }
        catch (Exception excp)
        {
            infobar.Visible = true;
            lblError.Text = m_ErrorHandler.ProcessError(excp);
        }
    }
    int svrecordcount = 0;
    protected void repSvConditions_ItemDataBound(object sender, RepeaterItemEventArgs e)
    {
        if (e.Item.DataItem != null)
        {

            CMS.AMS.Models.SVCondition SVCondition = e.Item.DataItem as CMS.AMS.Models.SVCondition;
            if (svrecordcount > 0)
            {
                if (e.Item.FindControl("lnkJoinType") != null)
                {
                    ((LinkButton)e.Item.FindControl("lnkJoinType")).Text = GetJoinTypeText(SVCondition.JoinTypeID);
                    if ((bEnableRestrictedAccessToUEOfferBuilder && isTranslatedOffer) || (bEnableAdditionalLockoutRestrictionsOnOffers && !bOfferEditable))
                        ((LinkButton)e.Item.FindControl("lnkJoinType")).Enabled = false;

                }

                if (e.Item.FindControl("lblJoinTypeID") != null)
                {
                    ((Label)e.Item.FindControl("lblJoinTypeID")).Text = ((int)SVCondition.JoinTypeID).ToString();

                }

            }
            svrecordcount = svrecordcount + 1;
        }


    }
    private string GetJoinTypeText(JoinTypes type)
    {
        if (type == JoinTypes.And)
        {
            return PhraseLib.Lookup("term.and", LanguageID);
        }
        else
        {
            return PhraseLib.Lookup("term.or", LanguageID);
        }

    }
}
