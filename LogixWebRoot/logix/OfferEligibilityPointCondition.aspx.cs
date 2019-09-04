using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using CMS;
using CMS.AMS;
using CMS.AMS.Contract;
using System.Data;
using CMS.AMS.Models;
using CMS.DB;
public partial class OfferEligibilityPointCondition : AuthenticatedUI
{
    #region Variables

    bool IsTemplate = false;
    long OfferID = 0;
    bool FromTemplate = false;
    bool DisabledAttribute = false;
    int EngineID = -1;
    long ConditionID = 0;
    //long PointsConditionID = 0;
    int ConditionTypeID = 0;
    string historyString;
    Copient.CommonInc MyCommon = new Copient.CommonInc();

    IPointsCondition m_PointsCondition;
    IOffer m_Offer;
    IPointsProgramService m_PointProgram;
    IDBAccess m_dbAccess;
    IOfferApprovalWorkflowService m_OAWService;

    Copient.LogixInc Logix = new Copient.LogixInc();
    bool bCreateGroupOrProgramFromOffer = false;
    bool isTranslatedOffer = false;
    bool bEnableRestrictedAccessToUEOfferBuilder = false;
    bool bOfferEditable = false;
    bool bEnableAdditionalLockoutRestrictionsOnOffers = false;

    #endregion

    #region Properties
    private List<CMS.AMS.Models.PointsProgram> AllPointsProgram
    {
        get
        {
            return Session["AllPointsProgram"] as List<CMS.AMS.Models.PointsProgram>;
        }
        set
        {
            Session["AllPointsProgram"] = value;
        }
    }
    private List<CMS.AMS.Models.PointsProgram> IncludedPointProgram
    {

        get
        {
            return ViewState["IncludedPointProgram"] as List<CMS.AMS.Models.PointsProgram>;
        }
        set
        {
            ViewState["IncludedPointProgram"] = value;
        }

    }
    private List<CMS.AMS.Models.PointsProgram> AvailableFilteredPointProgram
    {

        get
        {
            return Session["AvailableFilteredPointProgram"] as List<CMS.AMS.Models.PointsProgram>;
        }
        set
        {
            Session["AvailableFilteredPointProgram"] = value;
        }

    }
    private CMS.AMS.Models.PointsCondition OfferEligibilePointsCondition
    {
        get
        {
            return ViewState["OfferEligibilePointsCondition"] as CMS.AMS.Models.PointsCondition;
        }
        set
        {
            ViewState["OfferEligibilePointsCondition"] = value;
        }
    }
    #endregion

    #region Protected Methods
    protected override void OnInit(EventArgs e)
    {
        AppName = "OfferEligibilityPointCondition.aspx";
        base.OnInit(e);

    }
    protected void Page_Load(object sender, EventArgs e)
    {
        ResolveDepedencies();
        GetQueryStrings();
        bCreateGroupOrProgramFromOffer = MyCommon.Fetch_CM_SystemOption(134) == "1" ? true : false;
        bEnableRestrictedAccessToUEOfferBuilder = MyCommon.Fetch_SystemOption(249) == "1" ? true : false;
        AssignPageTitle("term.offer", "term.eligibilitypointscondition", OfferID.ToString());
        if (!IsPostBack)
        {
            SetUpAndLocalizePage();
            GetOfferEligiblePointsCondition();

            SetAvailableData();
            SetButtons();
            DisableControls();

        }
        else
        {
            ScriptManager.RegisterStartupScript(this, this.GetType(), "selectAndFocus", " SetFoucs();", true);
        }
    }


    protected void select1_Click(object sender, EventArgs e)
    {
        if (lstAvailable.SelectedItem != null)
        {

            foreach (int i in lstAvailable.GetSelectedIndices())
            {

                IncludedPointProgram.Add(AvailableFilteredPointProgram[i]);
            }
            SetAvailableData();

        }
        SetButtons();
    }
    protected void deselect1_Click(object sender, EventArgs e)
    {
        if (lstSelected.SelectedItem != null)
        {
            foreach (int i in lstSelected.GetSelectedIndices())
            {
                IncludedPointProgram.RemoveAt(i);
            }
            SetAvailableData();
        }
        SetButtons();
    }
    protected void ReloadThePanel_Click(object sender, EventArgs e)
    {
        SetAvailableData();
    }
    protected void btnSave_Click(object sender, EventArgs e)
    {
        try
        {
            if (!(lstSelected.Items.Count > 0))
            {
                infobar.InnerText = "Please select at least one inclusion group";
                infobar.Visible = true;
                return;
            }
            if (OfferEligibilePointsCondition == null)
            {
                OfferEligibilePointsCondition = new CMS.AMS.Models.PointsCondition();
            }
            if (OfferEligibilePointsCondition.ConditionID == 0)
                OfferEligibilePointsCondition.JoinTypeID = CMS.AMS.Models.JoinTypes.And;
            if (chkDisallow_Edit.Visible)
                OfferEligibilePointsCondition.DisallowEdit = chkDisallow_Edit.Checked;
            OfferEligibilePointsCondition.Deleted = false;
            OfferEligibilePointsCondition.ConditionID = ConditionID;

            OfferEligibilePointsCondition.ConditionTypeID = ConditionTypeID;
            OfferEligibilePointsCondition.EngineID = EngineID;
            //OfferEligibilePointsCondition.PointsConditionID = PointsConditionID;
            OfferEligibilePointsCondition.RequiredFromTemplate = false;
            OfferEligibilePointsCondition.Quantity = txtValueNeeded.Text.ConvertToInt32();
            OfferEligibilePointsCondition.ProgramID = lstSelected.Items[0].Value.ConvertToLong();
            if (OfferEligibilePointsCondition.Quantity == 0)
            {
                infobar.InnerText = Copient.PhraseLib.Lookup("pointscondition.invalidValueNeeded", LanguageID);
                infobar.Visible = true;
            }
            else
            {
                m_Offer.CreateUpdateOfferEligiblePointsCondition(OfferID, EngineID, OfferEligibilePointsCondition);
                m_Offer.UpdateOfferStatusToModified(OfferID, EngineID, CurrentUser.AdminUser.ID);
                m_OAWService.ResetOfferApprovalStatus(OfferID);
                historyString = PhraseLib.Lookup("history.CustomerPointConditionEdit", LanguageID) + ":" + lstSelected.Items[0].Text + " requires " + txtValueNeeded.Text.ConvertToInt32();
                WriteToActivityLog();
                ScriptManager.RegisterStartupScript(this, this.GetType(), "Close", "CloseModel()", true);
            }
        }
        catch (Exception ex)
        {
            infobar.InnerText = ErrorHandler.ProcessError(ex);
            infobar.Visible = true;
        }
    }
    protected void btnCreate_Click(object sender, EventArgs e)
    {
        string Name = string.Empty;
        if (MyCommon.Parse_Quotes(Logix.TrimAll(functioninput.Text)) != null)
            Name = Convert.ToString(MyCommon.Parse_Quotes(Logix.TrimAll(functioninput.Text)));
        if (!String.IsNullOrEmpty(Name))
        {
            int AvailableListCount = AvailableFilteredPointProgram.Where(p => p.ProgramName.Equals(functioninput.Text, StringComparison.OrdinalIgnoreCase)).ToList().Count;
            int IncludedGroupCount = IncludedPointProgram.Where(p => p.ProgramName.Equals(functioninput.Text, StringComparison.OrdinalIgnoreCase)).ToList().Count;

            if (IncludedGroupCount > 0)
            {
                string alertMessage = Copient.PhraseLib.Lookup("term.pointsprogram", LanguageID) + ": " + Name + " " + Copient.PhraseLib.Lookup("term.is", LanguageID).ToLower() + " " + Copient.PhraseLib.Lookup("term.already", LanguageID).ToLower() + " " + Copient.PhraseLib.Lookup("term.selected", LanguageID).ToLower();
                ScriptManager.RegisterStartupScript(UpdatePanelMain, UpdatePanelMain.GetType(), "AlertMessage", " AlertMessage('" + alertMessage + "');", true);
            }
            else if (AvailableListCount > 0)
            {

                string alertMessage = Copient.PhraseLib.Lookup("term.existing", LanguageID) + " " + Copient.PhraseLib.Lookup("term.pointsprogram", LanguageID).ToLower() + ": " + Name + " " + Copient.PhraseLib.Lookup("offer.message", LanguageID).ToLower();
                ScriptManager.RegisterStartupScript(UpdatePanelMain, UpdatePanelMain.GetType(), "AlertMessage", " AlertMessage('" + alertMessage + "');", true);

                //First remove the selected point if exist any 

                IncludedPointProgram.Clear();

                IncludedPointProgram.Add(AvailableFilteredPointProgram.Where(p => p.ProgramName.Equals(functioninput.Text, StringComparison.OrdinalIgnoreCase)).FirstOrDefault());
                SetAvailableData();
                SetButtons();


            }
            else
            {
                IncludedPointProgram.Clear();

                //Then add newly created points program to the selected list
                IncludedPointProgram.Add(CreatePointsProgram());
                SetAvailableData();

                SetButtons();
            }
        }
        else
        {
            string alertMessage = Copient.PhraseLib.Lookup("term.enter", LanguageID) + " " + Copient.PhraseLib.Lookup("term.valid", LanguageID).ToLower() + " " + Copient.PhraseLib.Lookup("term.pointsprogram", LanguageID).ToLower() + " " + Copient.PhraseLib.Lookup("term.name", LanguageID).ToLower();
            ScriptManager.RegisterStartupScript(UpdatePanelMain, UpdatePanelMain.GetType(), "AlertMessage", " AlertMessage('" + alertMessage + "');", true);


        }
    }

    #endregion

    #region Private Methods
    private void ResolveDepedencies()
    {
        m_PointsCondition = CurrentRequest.Resolver.Resolve<CMS.AMS.Contract.IPointsCondition>();
        m_Offer = CurrentRequest.Resolver.Resolve<CMS.AMS.Contract.IOffer>();

        m_PointProgram = CurrentRequest.Resolver.Resolve<IPointsProgramService>();
        m_dbAccess = CMS.AMS.CurrentRequest.Resolver.Resolve<IDBAccess>();
        m_OAWService = CurrentRequest.Resolver.Resolve<IOfferApprovalWorkflowService>();
    }
    private void SetUpAndLocalizePage()
    {

        if (IsTemplate)
            title.InnerText = PhraseLib.Lookup("term.template", LanguageID) + " #" + OfferID + " " + PhraseLib.Lookup("term.eligibilitypointscondition", LanguageID);
        else
            title.InnerText = PhraseLib.Lookup("term.offer", LanguageID) + " #" + OfferID + " " + PhraseLib.Lookup("term.eligibilitypointscondition", LanguageID);
        btnSave.Text = PhraseLib.Lookup("term.save", LanguageID);
        select1.Text = "▼" + PhraseLib.Lookup("term.select", LanguageID);
        deselect1.Text = PhraseLib.Lookup("term.deselect", LanguageID) + "▲";
        lblValueNeeded.Text = PhraseLib.Lookup("condition.valueneeded", LanguageID);
        btnCreate.Text = PhraseLib.Lookup("term.create", LanguageID);
    }
    private void GetQueryStrings()
    {
        OfferID = Request.QueryString["OfferID"].ConvertToLong();
        EngineID = Request.QueryString["EngineID"].ConvertToInt32();
        ConditionID = Request.QueryString["ConditionID"].ConvertToLong();
        //PointsConditionID = Request.QueryString["PointsConditionID"].ConvertToLong();
        IsTemplate = Request["IsTemplate"].ConvertToBool();
        FromTemplate = Request["FromTemplate"].ConvertToBool();
        ConditionTypeID = Request["ConditionTypeID"].ConvertToInt32();
        isTranslatedOffer = MyCommon.IsTranslatedUEOffer(Convert.ToInt32(OfferID), MyCommon);
        bOfferEditable = MyCommon.IsOfferEditablePastLockOutPeriod(CurrentUser.UserPermissions.EditOfferPastLockoutPeriod, MyCommon, Convert.ToInt32(OfferID));
        bEnableAdditionalLockoutRestrictionsOnOffers = MyCommon.Fetch_SystemOption(260) == "1" ? true : false;
    }
    private void GetAllPointsProgram()
    {
        AllPointsProgram = m_PointProgram.GetPointsPrograms();
    }
    private void SetAvailableData()
    {
        GetAllPointsProgram();
        AvailableFilteredPointProgram = AllPointsProgram.Where(p => !IncludedPointProgram.Any(inc => inc.ProgramID == p.ProgramID)).ToList();

        string strFilter = functioninput.Text;

        if (functionradio1.Checked)
            AvailableFilteredPointProgram = AvailableFilteredPointProgram.Where(p => p.ProgramName.StartsWith(strFilter, StringComparison.OrdinalIgnoreCase)).ToList();
        else
            AvailableFilteredPointProgram = AvailableFilteredPointProgram.Where(p => p.ProgramName.IndexOf(strFilter, StringComparison.OrdinalIgnoreCase) >= 0).ToList();


        lstSelected.DataSource = IncludedPointProgram;
        lstSelected.DataBind();

        lstAvailable.DataSource = AvailableFilteredPointProgram;
        lstAvailable.DataBind();

    }
    private void SetButtons()
    {
        if (lstSelected.Items.Count == 0)
        {
            select1.Enabled = true;
        }
        if (lstSelected.Items.Count > 0)
        {
            deselect1.Enabled = true;
            select1.Enabled = false;
        }
        else
        {
            deselect1.Enabled = false;
        }
    }
    protected override void AuthorisePage()
    {
        if (CurrentUser.UserPermissions.AccessOffers == false && !IsTemplate)
        {
            Server.Transfer("PopUpDenied.aspx?PhraseName=perm.offers-access", false);
            return;
        }
        if (CurrentUser.UserPermissions.AccessTemplates == false && IsTemplate)
        {
            Server.Transfer("PopUpDenied.aspx?PhraseName=perm.offers-access-templates", false);
            return;
        }

    }
    private void DisableControls()
    {
        if (!IsTemplate)
            TempDisallow.Visible = false;
        if (!IsTemplate)
        {
            DisabledAttribute = ((CurrentUser.UserPermissions.EditOffer && !(FromTemplate && OfferEligibilePointsCondition.DisallowEdit)) ? false : true);
        }
        else
            DisabledAttribute = CurrentUser.UserPermissions.EditTemplates ? false : true;
        //If disable is set to false, check Buyer conditions
        if (EngineID == 9 && !DisabledAttribute)
        {
            if (MyCommon.LRTadoConn.State == ConnectionState.Closed)
                MyCommon.Open_LogixRT();
            DisabledAttribute = ((CurrentUser.UserPermissions.EditOffersRegardlessBuyer || MyCommon.IsOfferCreatedWithUserAssociatedBuyer(CurrentUser.AdminUser.ID, OfferID)) ? false : true);
            MyCommon.Close_LogixRT();
        }

        if (DisabledAttribute)
        {
            functionradio1.Enabled = false;
            functionradio2.Enabled = false;
            functioninput.Enabled = false;
            lstAvailable.Enabled = false;
            lstSelected.Enabled = false;
            txtValueNeeded.Enabled = false;
            select1.Enabled = false;
            deselect1.Enabled = false;
            btnSave.Visible = false;
            btnCreate.Enabled = false;
        }

        if (!bCreateGroupOrProgramFromOffer || !CurrentUser.UserPermissions.CreatePointsPrograms || EngineID != 0)
        {
            btnCreate.Visible = false;
        }
        if ((bEnableRestrictedAccessToUEOfferBuilder && isTranslatedOffer) || (bEnableAdditionalLockoutRestrictionsOnOffers && !bOfferEditable) || m_OAWService.CheckIfOfferIsAwaitingApproval(OfferID).Result)
        {
            btnSave.Visible = false;
        }
    }
    private void GetOfferEligiblePointsCondition()
    {
        if (ConditionID > 0)
        {
            OfferEligibilePointsCondition = m_PointsCondition.GetConditionByID(ConditionID);
        }
        if (IncludedPointProgram == null)
        {
            IncludedPointProgram = new List<CMS.AMS.Models.PointsProgram>();
        }
        if (OfferEligibilePointsCondition == null)
            OfferEligibilePointsCondition = new CMS.AMS.Models.PointsCondition();
        else
        {
            CMS.AMS.Models.PointsProgram objPointsProgram = new CMS.AMS.Models.PointsProgram();
            objPointsProgram.ProgramID = OfferEligibilePointsCondition.ProgramID;

            IncludedPointProgram.Add(OfferEligibilePointsCondition.PointsProgram);
            SetValues(OfferEligibilePointsCondition);
        }
    }
    private void SetValues(CMS.AMS.Models.PointsCondition OfferEligibilePointsCondition)
    {
        txtValueNeeded.Text = OfferEligibilePointsCondition.Quantity.ToString();
        chkDisallow_Edit.Checked = OfferEligibilePointsCondition.DisallowEdit;
    }
    private void WriteToActivityLog()
    {
        if (MyCommon.LRTadoConn.State == ConnectionState.Closed)
            MyCommon.Open_LogixRT();
        MyCommon.Activity_Log(3, OfferID, CurrentUser.AdminUser.ID, historyString);
        MyCommon.Close_LogixRT();
    }


    private CMS.AMS.Models.PointsProgram CreatePointsProgram()
    {
        CMS.AMS.Models.PointsProgram NewPointProgram = null;

        try
        {
            NewPointProgram = new PointsProgram { ProgramName = Logix.TrimAll(functioninput.Text) };
            bool saved = CreatePointsProgram(NewPointProgram);

            if (saved)
                return NewPointProgram;
            else
                return null;
        }
        catch (Exception err)
        {
            ScriptManager.RegisterStartupScript(UpdatePanelMain, UpdatePanelMain.GetType(), "AlertMessage", " AlertMessage('" + err.Message + "');", true);
            return null;
        }

    }

    private bool CreatePointsProgram(PointsProgram pointsProgram)
    {
        bool result = false;
        SQLParametersList lst = new SQLParametersList();

        try
        {
            lst.Add("@CAMProgram", SqlDbType.Bit).Value = 0;
            lst.Add("@ExternalProgram", SqlDbType.Bit).Value = 0;
            lst.Add("@AutoDelete", SqlDbType.Bit).Value = 1;

            if (pointsProgram.ProgramID == 0)
            {
                lst.Add("@ProgramName", SqlDbType.NVarChar, 200).Value = pointsProgram.ProgramName;
                lst.Add("@ProgramID", SqlDbType.BigInt).Direction = ParameterDirection.Output;
                m_dbAccess.ExecuteNonQuery(DataBases.LogixRT, CommandType.StoredProcedure, "pt_PointsPrograms_Insert", lst);
                pointsProgram.ProgramID = lst["@ProgramID"].Value.ConvertToLong();
                result = true;
                if (MyCommon.LRTadoConn.State == ConnectionState.Closed)
                    MyCommon.Open_LogixRT();
                MyCommon.Activity_Log(7, pointsProgram.ProgramID, CurrentUser.AdminUser.ID, Copient.PhraseLib.Lookup("history.point-create", LanguageID));

                lst = new SQLParametersList();
                lst.Add("@ProgramID", SqlDbType.NVarChar, 200).Value = pointsProgram.ProgramID;
                lst.Add("@VarID", SqlDbType.BigInt).Direction = ParameterDirection.Output;
                m_dbAccess.ExecuteNonQuery(DataBases.LogixXS, CommandType.StoredProcedure, "dbo.pc_PointsVar_Create", lst);
                long PromoVarID = lst["@VarID"].Value.ConvertToLong();

                if (PromoVarID != 0)
                {
                    MyCommon.QueryStr = " update PointsPrograms with (RowLock) SET " +
                                        " PromoVarID=" + PromoVarID + " " + " where ProgramID=" + pointsProgram.ProgramID + ";";
                    MyCommon.LRT_Execute();
                }
            }
        }
        catch (Exception ex)
        {
            throw ex;
        }
        finally
        {
            MyCommon.Close_LogixRT();
        }
        lst = null;
        return result;
    }

}
#endregion
