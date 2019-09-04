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
using CMS.DB;
public partial class OfferEligibilitySVCondition : AuthenticatedUI
{
    #region Variables

    bool IsTemplate = false;
    long OfferID = 0;
    bool FromTemplate = false;
    bool DisabledAttribute = false;
    int EngineID = -1;
    long ConditionID = 0;
    int ConditionTypeID = 0;
    string historyString;
    Copient.CommonInc MyCommon = new Copient.CommonInc();
    Copient.LogixInc Logix = new Copient.LogixInc();
    bool bCreateGroupOrProgramFromOffer = false;
    bool isTranslatedOffer = false;
    bool bEnableRestrictedAccessToUEOfferBuilder = false;
    bool bOfferEditable = false;
    bool bEnableAdditionalLockoutRestrictionsOnOffers = false;
    IStoredValueProgramService m_StoredValueProgram;
    IStoredValueCondition m_StoredValueCondition;
    IOfferApprovalWorkflowService m_OAWService;
    IDBAccess m_dbAccess;
    IOffer m_Offer;

    #endregion

    #region Properties
    private List<CMS.AMS.Models.SVProgram> AllSVPrograms
    {
        get
        {
            return Session["AllSVPrograms"] as List<CMS.AMS.Models.SVProgram>;
        }
        set
        {
            Session["AllSVPrograms"] = value;
        }
    }
    private List<CMS.AMS.Models.SVProgram> IncludedSVProgram
    {

        get
        {
            return ViewState["IncludedSVProgram"] as List<CMS.AMS.Models.SVProgram>;
        }
        set
        {
            ViewState["IncludedSVProgram"] = value;
        }

    }
    private List<CMS.AMS.Models.SVProgram> AvailableFilteredSVPrograms
    {

        get
        {
            return Session["AvailableFilteredSVPrograms"] as List<CMS.AMS.Models.SVProgram>;
        }
        set
        {
            Session["AvailableFilteredSVPrograms"] = value;
        }

    }
    private CMS.AMS.Models.SVCondition OfferEligibileSVCondition
    {
        get
        {
            return ViewState["OfferEligibileSVCondition"] as CMS.AMS.Models.SVCondition;
        }
        set
        {
            ViewState["OfferEligibileSVCondition"] = value;
        }
    }
    #endregion

    #region Protected Methods
    protected override void OnInit(EventArgs e)
    {
        AppName = "OfferEligibilitySVCondition.aspx";
        base.OnInit(e);

    }

    protected void Page_Load(object sender, EventArgs e)
    {
        ResolveDepedencies();
        GetQueryStrings();
        bCreateGroupOrProgramFromOffer = MyCommon.Fetch_CM_SystemOption(134) == "1" ? true : false;
        bEnableRestrictedAccessToUEOfferBuilder = MyCommon.Fetch_SystemOption(249) == "1" ? true : false;
        AssignPageTitle("term.offer", "term.eligibilitystoredvaluecondition", OfferID.ToString());
        if (!IsPostBack)
        {
            SetUpAndLocalizePage();
            GetOfferEligibleSVCondition();
            SetAvailableData();
            SetButtons();
            DisableControls();

        }
        else
            ScriptManager.RegisterStartupScript(this, this.GetType(), "selectAndFocus", " SetFoucs();", true);
    }

    protected void select1_Click(object sender, EventArgs e)
    {
        if (lstAvailable.SelectedItem != null)
        {
            foreach (int i in lstAvailable.GetSelectedIndices())
            {
                IncludedSVProgram.Add(AvailableFilteredSVPrograms[i]);
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
                IncludedSVProgram.RemoveAt(i);
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
            if (OfferEligibileSVCondition == null)
            {
                OfferEligibileSVCondition = new CMS.AMS.Models.SVCondition();
            }
            if (chkDisallow_Edit.Visible)
                OfferEligibileSVCondition.DisallowEdit = chkDisallow_Edit.Checked;
            if (OfferEligibileSVCondition.ConditionID == 0)
                OfferEligibileSVCondition.JoinTypeID = CMS.AMS.Models.JoinTypes.And;
            OfferEligibileSVCondition.Deleted = false;
            OfferEligibileSVCondition.ConditionID = ConditionID;
            OfferEligibileSVCondition.EngineID = EngineID;
            OfferEligibileSVCondition.ConditionTypeID = ConditionTypeID;
            OfferEligibileSVCondition.RequiredFromTemplate = false;
            OfferEligibileSVCondition.Quantity = txtValueNeeded.Text.ConvertToInt32();
            OfferEligibileSVCondition.SVProgramID = lstSelected.Items[0].Value.ConvertToLong();
            if (OfferEligibileSVCondition.Quantity == 0)
            {
                infobar.InnerText = Copient.PhraseLib.Lookup("pointscondition.invalidValueNeeded", LanguageID);
                infobar.Visible = true;
            }
            else
            {
                m_Offer.CreateUpdateOfferEligibleStoredValueCondition(OfferID, EngineID, OfferEligibileSVCondition);
                m_Offer.UpdateOfferStatusToModified(OfferID, EngineID, CurrentUser.AdminUser.ID);
                m_OAWService.ResetOfferApprovalStatus(OfferID);
                historyString = PhraseLib.Lookup("history.CustomerStoredValueConditionEdit", LanguageID) + ":" + lstSelected.Items[0].Text + " requires " + txtValueNeeded.Text.ConvertToInt32();
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
            int AvailableListCount = AvailableFilteredSVPrograms.Where(p => p.ProgramName.Equals(Name, StringComparison.OrdinalIgnoreCase)).ToList().Count;
            int IncludedGroupCount = IncludedSVProgram.Where(p => p.ProgramName.Equals(functioninput.Text, StringComparison.OrdinalIgnoreCase)).ToList().Count;

            if (IncludedGroupCount > 0)
            {
                string alertMessage = Copient.PhraseLib.Lookup("term.storedvalueprogram", LanguageID) + ": " + Name + " " + Copient.PhraseLib.Lookup("term.is", LanguageID).ToLower() + " " + Copient.PhraseLib.Lookup("term.already", LanguageID).ToLower() + " " + Copient.PhraseLib.Lookup("term.selected", LanguageID).ToLower();
                ScriptManager.RegisterStartupScript(UpdatePanelMain, UpdatePanelMain.GetType(), "AlertMessage", " AlertMessage('" + alertMessage + "');", true);
            }
            else if (AvailableListCount > 0)
            {

                string alertMessage = Copient.PhraseLib.Lookup("term.existing", LanguageID) + " " + Copient.PhraseLib.Lookup("term.storedvalueprogram", LanguageID).ToLower() + ": " + Name + " " + Copient.PhraseLib.Lookup("offer.message", LanguageID).ToLower();
                ScriptManager.RegisterStartupScript(UpdatePanelMain, UpdatePanelMain.GetType(), "AlertMessage", " AlertMessage('" + alertMessage + "');", true);

                //First remove the selected point if exist any 

                IncludedSVProgram.Clear();

                IncludedSVProgram.Add(AvailableFilteredSVPrograms.Where(p => p.ProgramName.Equals(functioninput.Text, StringComparison.OrdinalIgnoreCase)).FirstOrDefault());
                SetAvailableData();
                SetButtons();


            }
            else
            {
                IncludedSVProgram.Clear();

                //Then add newly created points program to the selected list
                IncludedSVProgram.Add(CreatePointsProgram());
                SetAvailableData();
                SetButtons();
            }
        }
        else
        {
            string alertMessage = Copient.PhraseLib.Lookup("term.enter", LanguageID) + " " + Copient.PhraseLib.Lookup("term.valid", LanguageID).ToLower() + " " + Copient.PhraseLib.Lookup("term.storedvalueprogram", LanguageID).ToLower() + " " + Copient.PhraseLib.Lookup("term.name", LanguageID).ToLower();
            ScriptManager.RegisterStartupScript(UpdatePanelMain, UpdatePanelMain.GetType(), "AlertMessage", " AlertMessage('" + alertMessage + "');", true);


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

    #endregion

    #region Private Methods
    private void ResolveDepedencies()
    {
        m_StoredValueCondition = CurrentRequest.Resolver.Resolve<CMS.AMS.Contract.IStoredValueCondition>();
        m_Offer = CurrentRequest.Resolver.Resolve<CMS.AMS.Contract.IOffer>();
        m_StoredValueProgram = CurrentRequest.Resolver.Resolve<IStoredValueProgramService>();
        m_dbAccess = CMS.AMS.CurrentRequest.Resolver.Resolve<IDBAccess>();
        m_OAWService = CurrentRequest.Resolver.Resolve<IOfferApprovalWorkflowService>();
    }

    private void SetUpAndLocalizePage()
    {

        if (IsTemplate)
            title.InnerText = PhraseLib.Lookup("term.template", LanguageID) + " #" + OfferID + " " + PhraseLib.Lookup("term.eligibilitystoredvaluecondition", LanguageID);
        else
            title.InnerText = PhraseLib.Lookup("term.offer", LanguageID) + " #" + OfferID + " " + PhraseLib.Lookup("term.eligibilitystoredvaluecondition", LanguageID);
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
        ConditionTypeID = Request.QueryString["ConditionTypeID"].ConvertToInt32();
        IsTemplate = Request["IsTemplate"].ConvertToBool();
        FromTemplate = Request["FromTemplate"].ConvertToBool();
        isTranslatedOffer = MyCommon.IsTranslatedUEOffer(Convert.ToInt32(OfferID), MyCommon);
        bOfferEditable = MyCommon.IsOfferEditablePastLockOutPeriod(CurrentUser.UserPermissions.EditOfferPastLockoutPeriod, MyCommon, Convert.ToInt32(OfferID));
        bEnableAdditionalLockoutRestrictionsOnOffers = MyCommon.Fetch_SystemOption(260) == "1" ? true : false;
    }

    private void GetAllSVPrograms()
    {
        AllSVPrograms = m_StoredValueProgram.GetStoredValuePrograms();
    }

    private void SetAvailableData()
    {
        GetAllSVPrograms();
        AvailableFilteredSVPrograms = AllSVPrograms.Where(p => !IncludedSVProgram.Any(inc => inc.SVProgramID == p.SVProgramID)).ToList();

        string strFilter = functioninput.Text;

        if (functionradio1.Checked)
            AvailableFilteredSVPrograms = AvailableFilteredSVPrograms.Where(p => p.ProgramName.StartsWith(strFilter, StringComparison.OrdinalIgnoreCase)).ToList();
        else
            AvailableFilteredSVPrograms = AvailableFilteredSVPrograms.Where(p => p.ProgramName.IndexOf(strFilter, StringComparison.OrdinalIgnoreCase) >= 0).ToList();


        lstSelected.DataSource = IncludedSVProgram;
        lstSelected.DataBind();

        lstAvailable.DataSource = AvailableFilteredSVPrograms;
        lstAvailable.DataBind();

        chkDisallow_Edit.Checked = OfferEligibileSVCondition.DisallowEdit;

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

    private void GetOfferEligibleSVCondition()
    {
        if (ConditionID > 0)
        {
            OfferEligibileSVCondition = m_StoredValueCondition.GetConditionByID(ConditionID);
        }
        if (OfferEligibileSVCondition == null)
            OfferEligibileSVCondition = new CMS.AMS.Models.SVCondition();
        else
            SetValues(OfferEligibileSVCondition);
        if (IncludedSVProgram == null)
        {
            IncludedSVProgram = new List<CMS.AMS.Models.SVProgram>();
        }
        if (OfferEligibileSVCondition.ProgramID > 0) IncludedSVProgram.Add(OfferEligibileSVCondition.SVProgram);
    }

    private void SetValues(CMS.AMS.Models.SVCondition OfferEligibileSVCondition)
    {
        txtValueNeeded.Text = OfferEligibileSVCondition.Quantity.ToString();
        chkDisallow_Edit.Checked = OfferEligibileSVCondition.DisallowEdit;
    }

    private void DisableControls()
    {
        if (!IsTemplate)
            TempDisallow.Visible = false;
        if (!IsTemplate)
            DisabledAttribute = ((CurrentUser.UserPermissions.EditOffer && !(FromTemplate && OfferEligibileSVCondition.DisallowEdit)) ? false : true);
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
        if (!bCreateGroupOrProgramFromOffer || !CurrentUser.UserPermissions.CreateStoredValuePrograms || EngineID != 0)
        {
            btnCreate.Visible = false;
        }
        if ((bEnableRestrictedAccessToUEOfferBuilder && isTranslatedOffer) || (bEnableAdditionalLockoutRestrictionsOnOffers && !bOfferEditable) || m_OAWService.CheckIfOfferIsAwaitingApproval(OfferID).Result)
        {
            btnSave.Visible = false;
        }
    }

    private void WriteToActivityLog()
    {
        if (MyCommon.LRTadoConn.State == ConnectionState.Closed)
            MyCommon.Open_LogixRT();
        MyCommon.Activity_Log(3, OfferID, CurrentUser.AdminUser.ID, historyString);
        MyCommon.Close_LogixRT();
    }

    private CMS.AMS.Models.SVProgram CreatePointsProgram()
    {
        CMS.AMS.Models.SVProgram NewSVPointProgram = null;

        try
        {
            NewSVPointProgram = new CMS.AMS.Models.SVProgram { ProgramName = Logix.TrimAll(functioninput.Text) };
            bool saved = CreateSVPointsProgram(NewSVPointProgram);

            if (saved)
                return NewSVPointProgram;
            else
                return null;
        }
        catch (Exception err)
        {
            ScriptManager.RegisterStartupScript(UpdatePanelMain, UpdatePanelMain.GetType(), "AlertMessage", " AlertMessage('" + err.Message + "');", true);
            return null;
        }

    }

    private bool CreateSVPointsProgram(CMS.AMS.Models.SVProgram SVPointsProgram)
    {
        bool result = false;
        SQLParametersList lst = new SQLParametersList();

        try
        {
            if (SVPointsProgram.SVProgramID == 0)
            {
                lst.Add("@Name", SqlDbType.NVarChar, 200).Value = SVPointsProgram.ProgramName;
                lst.Add("@Description", SqlDbType.NVarChar, 200).Value = "";
                lst.Add("@ExpirePeriod", SqlDbType.Int).Value = 1;
                lst.Add("@Value", SqlDbType.NVarChar, 200).Value = "1";
                lst.Add("@OneUnitPerRec", SqlDbType.Bit).Value = 0;
                lst.Add("@SVExpireType", SqlDbType.Int).Value = 1;
                lst.Add("@SVExpirePeriodType", SqlDbType.Int).Value = 1;
                lst.Add("@ExpireTOD", SqlDbType.VarChar, 5).Value = "";
                lst.Add("@ExpireDate", SqlDbType.DateTime).Value = Convert.ToDateTime("12/31/2025 23:59");
                lst.Add("@ExpireCentralServerTZ", SqlDbType.Bit).Value = 0;
                lst.Add("@SVTypeID", SqlDbType.Int).Value = 1;
                lst.Add("@UOMLimit", SqlDbType.Int).Value = 1;
                lst.Add("@AllowReissue", SqlDbType.Int).Value = 0;
                lst.Add("@ScorecardID", SqlDbType.Int).Value = 0;
                lst.Add("@ScorecardDesc", SqlDbType.NVarChar, 100).Value = "";
                lst.Add("@ScorecardBold", SqlDbType.Bit).Value = 1;
                lst.Add("@DisallowRedeemInEarnTrans", SqlDbType.Int).Value = 0;
                lst.Add("@AllowNegativeBal", SqlDbType.Int).Value = 0;
                lst.Add("@AdjustmentUPC", SqlDbType.NVarChar, 100).Value = "";
                lst.Add("@RedemptionRestrictionID", SqlDbType.Int).Value = 0;
                lst.Add("@MemberRedemptionID", SqlDbType.Int).Value = 0;
                lst.Add("@SVProgramID", SqlDbType.BigInt).Direction = ParameterDirection.Output;

                m_dbAccess.ExecuteNonQuery(DataBases.LogixRT, CommandType.StoredProcedure, "pt_StoredValuePrograms_Insert", lst);
                SVPointsProgram.SVProgramID = lst["@SVProgramID"].Value.ConvertToLong();
                result = true;
                if (MyCommon.LRTadoConn.State == ConnectionState.Closed)
                    MyCommon.Open_LogixRT();
                MyCommon.Activity_Log(26, SVPointsProgram.SVProgramID, CurrentUser.AdminUser.ID, Copient.PhraseLib.Lookup("history-sv-create", LanguageID));
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
    #endregion

}