using System;
using System.Web.UI;
using CMS.AMS.Contract;
using CMS.AMS.Models;
using CMS.AMS;
using CMS;
using CMS.Contract;
using System.Web.UI.WebControls;
using System.Data;

public partial class logix_UE_PendingApproval : AuthenticatedUI
{

    #region Private Variables
    private ILogger m_logger;
    private int adminUserID;
    private bool searchtext;
    private string ProgramName = string.Empty;
    private int pageIndex = 0;
    private int pageSize = 20;
    private int startRowNum = 0;
    private int RecordCount = 0;
    private string sortingText = "";
    private int matchOfferId = -1;
    protected bool isOCDEnabled = false;
    string offer_Name = string.Empty;
    private Copient.CommonInc mCommon;
    private IOfferApprovalWorkflowService oawService;
    private IOffer offerService;
    private bool isBannerEnabled;
    #endregion 

    #region Protected Methods
    protected void Page_Load(object sender, EventArgs e)
    {
        adminUserID = ((AuthenticatedUI)this.Page).CurrentUser.AdminUser.ID;
        AssignPageTitle("term.approvals");
        ResolveDependencies();
        SetUserControls();
        GetSearchText();
        GetSortingText();
        isBannerEnabled = mCommon.Fetch_SystemOption(66).Equals("1") ? true : false;
        if (!Page.IsPostBack)
        {
            gvPendingOfferList.SortKey = "IncentiveID";
            gvPendingOfferList.SortOrder = "Desc";
            (this.Master as logix_LogixMasterPage).Tab_Name = "2_8"; // tab name for Approval is 2_8 in AppMenus (database)
            FillPageControlTextAndData(0);
        }
    }
    protected void Reject_Click(object sender, EventArgs e)
    {
        CallScript("showRejectConfirmation", "showRejectConfirmation()");
        GridViewRow gvRow = (GridViewRow)((Button)sender).NamingContainer;
        long offerId = ((Label)gvRow.FindControl("ID")).Text.ConvertToLong();
        OfferID.Value = offerId.ToString();
    }
    protected void Approve_Click(object sender, EventArgs e)
    {
        try
        {
            int approvalType = 13;
            GridViewRow gvRow = (GridViewRow)((Button)sender).NamingContainer;
            long offer_Id = ((Label)gvRow.FindControl("ID")).Text.ConvertToLong();
            AMSResult<bool> isOCDEnabledResult = offerService.IsCollisionDetectionEnabled(Engines.UE, offer_Id);
            if (isOCDEnabledResult.ResultType == AMSResultType.Success && isOCDEnabledResult.Result == true)
            {
                isOCDEnabled = true;
            }
            AMSResult<int> approvalTypeResult = oawService.GetOfferStatusFlag(offer_Id);
            if(approvalTypeResult.ResultType == AMSResultType.Success)
            {
                approvalType = approvalTypeResult.Result;
            }

            CallScript("ApproveOffer", "ApproveOffer(" + "'" + offer_Id + "'" + ", " + "'" + approvalType + "'" + ")");
        }
        catch (Exception ex)
        {
            m_logger.WriteCritical(ex.ToString());
            DisplayError(PhraseLib.Lookup("error.checklogs", LanguageID));
        }
    }
    protected void gvPendingOfferList_Sorting(object sender, GridViewSortEventArgs e)
    {
        GetSortingText();
        FilterOffers();
    }
    protected void gvPendingOfferList_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        int collisionCount = 0;
        long offer_Id = 0;
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            if ((e.Row.RowState == DataControlRowState.Normal) || (e.Row.RowState == (DataControlRowState.Alternate | DataControlRowState.Normal)))
            {
                offer_Id = ((Label)e.Row.Cells[0].FindControl("ID")).Text.ConvertToLong();
                collisionCount = GetCollisionCount(offer_Id);
                if (collisionCount > 0)
                {
                    e.Row.Cells[0].FindControl("CollisionReport").Visible = true;
                }
                ((Button)e.Row.Cells[0].FindControl("Approve")).ToolTip = PhraseLib.Lookup("term.approve", LanguageID);
                ((Button)e.Row.Cells[0].FindControl("Reject")).ToolTip = PhraseLib.Lookup("term.reject", LanguageID);
                ((Button)e.Row.Cells[0].FindControl("CollisionReport")).ToolTip = PhraseLib.Lookup("term.viewcollisionreport", LanguageID);
            }
        }
    }

    protected void CollisionReport_Click(object sender, EventArgs e)
    {
        GridViewRow gvRow = (GridViewRow)((Button)sender).NamingContainer;
        long offer_Id = ((Label)gvRow.FindControl("ID")).Text.ConvertToLong();
        Response.Redirect("../CollidingOffers-Report.aspx?ID=" + offer_Id);
    }
    protected void reject_Click(object sender, EventArgs e)
    {
        string rejectionMsg = string.Empty;
        if (Request.Form["rejectText"].ToString().Trim() != "")
        {
            rejectionMsg = Request.Form["rejectText"].ToString().Trim();
        }
        AMSResult<bool> offerRejected = oawService.RejectOffer(OfferID.Value.ConvertToLong(), adminUserID, rejectionMsg);
        if (offerRejected.Result)
        {
            string logMessage = Copient.PhraseLib.Lookup("term.offer-rejected", LanguageID);
            if(rejectionMsg != "")
            {
                logMessage = logMessage + " " + Copient.PhraseLib.Lookup("term.rejectionreason", LanguageID) + ": " + rejectionMsg;
            }
            mCommon.Activity_Log(3, OfferID.Value.ConvertToLong(), adminUserID, logMessage);
            FilterOffers();
        }
        else
        {
            DisplayError(PhraseLib.Lookup("error.checklogs", LanguageID));
        }
        CallScript("hideRejectConfirmation", "hideRejectConfirmation()");
    }

    protected void cancel_Click(object sender, EventArgs e)
    {
        CallScript("hideRejectConfirmation", "hideRejectConfirmation()");
    }
    protected void filterddl_SelectedIndexChanged(object sender, EventArgs e)
    {
        FilterOffers();
    }
    protected override void AuthorisePage()
    {
        if (CurrentUser.UserPermissions.OfferApproval == false)
        {
            htitle.Visible = false;
            gvPendingOfferList.Visible = false;
            Search.Visible = false;
            Paging.Visible = false;
            parentfilter.Visible = false;
            accessdenied.Style["display"] = "block";
        }
    }
    #endregion 

    #region Private Methods
    private void SetUserControls()
    {
        this.Form.DefaultButton = (Search.SearchButton.UniqueID);
        Paging.PageSize = 20;
        Search.OnSearch += SearchControl_OnSearch;
        Paging.OnFirstPageClick += PageingControl_OnFirstPageClick;
        Paging.OnPreviousPageClick += PageingControl_OnPreviousPageClick;
        Paging.OnNextPageClick += PageingControl_OnNextPageClick;
        Paging.OnLastPageClick += PageingControl_OnLastPageClick;
    }
    private void PageingControl_OnLastPageClick(object sender, EventArgs e)
    {
        FilterOffers();
    }

    private void PageingControl_OnNextPageClick(object sender, EventArgs e)
    {
        FilterOffers();
    }

    private void PageingControl_OnPreviousPageClick(object sender, EventArgs e)
    {
        FilterOffers();
    }

    private void SearchControl_OnSearch(object sender, EventArgs e)
    {
        try
        {
            FilterOffers();
        }
        catch (Exception ex)
        {
            m_logger.WriteCritical(ex.ToString());
            DisplayError(PhraseLib.Lookup("error.checklogs", LanguageID));
        }
    }

    private void PageingControl_OnFirstPageClick(object sender, EventArgs e)
    {
        FilterOffers();
    }
    private void FilterOffers()
    {
        if (filterddl.SelectedValue == "0")
        {
            FetchPendingOffersData(Paging.PageIndex);
        }
        else if (filterddl.SelectedValue == "1")
        {
            FetchPendingOffersData(Paging.PageIndex, true);
        }
    }
    private void ResolveDependencies()
    {
        mCommon = new Copient.CommonInc();
        oawService = CurrentRequest.Resolver.Resolve<IOfferApprovalWorkflowService>();
        offerService = CurrentRequest.Resolver.Resolve<IOffer>();
    }
    private void FillPageControlTextAndData(int pageIndex)
    {
        filterddl.Items[0].Text = PhraseLib.Lookup("term.alloffers", LanguageID);
        filterddl.Items[1].Text = PhraseLib.Lookup("term.expiredoffers", LanguageID);
        reject.Text = PhraseLib.Lookup("term.reject", LanguageID);
        cancel.Text = PhraseLib.Lookup("term.cancel", LanguageID);
        htitle.InnerText = PhraseLib.Lookup("term.pendingoffers", LanguageID);
        for (int i = 0; i < gvPendingOfferList.Columns.Count; i++)
        {
            switch (i)
            {
                case 0:
                    gvPendingOfferList.Columns[i].HeaderText = PhraseLib.Lookup("term.id", LanguageID);
                    break;
                case 1:
                    gvPendingOfferList.Columns[i].HeaderText = PhraseLib.Lookup("term.offername", LanguageID);
                    break;
                case 2:
                    gvPendingOfferList.Columns[i].HeaderText = PhraseLib.Lookup("term.startdate", LanguageID);
                    break;
                case 3:
                    gvPendingOfferList.Columns[i].HeaderText = PhraseLib.Lookup("term.submittedby", LanguageID);
                    break;
                case 4:
                    gvPendingOfferList.Columns[i].HeaderText = PhraseLib.Lookup("term.waitingsince", LanguageID);
                    break;
                case 5:
                    gvPendingOfferList.Columns[i].HeaderText = PhraseLib.Lookup("term.action", LanguageID);
                    break;
            }
        }
        FetchPendingOffersData(0);

    }
    private void FetchPendingOffersData(int pageIndex, bool showExpired = false)
    {
        AMSResult<DataTable> pendingOffersList = new AMSResult<DataTable>();
        pendingOffersList = oawService.GetPendingOffersList(adminUserID, isBannerEnabled, pageIndex, Paging.PageSize, sortingText, offer_Name, matchOfferId, out RecordCount, showExpired);
        if (pendingOffersList.ResultType != AMSResultType.Success)
        {
            DisplayError(pendingOffersList.MessageString);
        }
        else
        {
            Paging.RecordCount = RecordCount;
            Paging.PageIndex = pageIndex;
            Paging.DataBind();
        }
        if(pendingOffersList.Result == null)
        {
            pendingOffersList.Result = new DataTable();
        }
        gvPendingOfferList.DataSource = pendingOffersList.Result;
        gvPendingOfferList.DataBind();
    }
    private void DisplayError(String errorText)
    {
        infobar.Attributes["class"] = "red-background";
        infobar.InnerText = errorText;
        infobar.Style["display"] = "block";
    }
    private void DisplaySuccessMsg(String message)
    {
        infobar.Attributes["class"] = "green-background";
        infobar.InnerText = message;
        infobar.Style["display"] = "block";
    }
    private void GetSortingText()
    {
        if (gvPendingOfferList.SortKey.Length > 0 && gvPendingOfferList.SortOrder.Length > 0)
        {
            if(gvPendingOfferList. SortKey == "ID")
            {
                sortingText = " ORDER BY OA.IncentiveID " + gvPendingOfferList.SortOrder;
            }
            else if (gvPendingOfferList.SortKey == "OfferName")
            {
                sortingText = " ORDER BY CI.IncentiveName " + gvPendingOfferList.SortOrder;
            }
            else if (gvPendingOfferList.SortKey == "StartDate")
            {
                sortingText = " ORDER BY CI.StartDate " + gvPendingOfferList.SortOrder;
            }
            else if (gvPendingOfferList.SortKey == "SubmittedBy")
            {
                sortingText = " ORDER BY (AU.FirstName + AU.LastName) " + gvPendingOfferList.SortOrder;
            }
            else if (gvPendingOfferList.SortKey == "WaitingSince")
            {
                sortingText = " ORDER BY CI.LastUpdate " + gvPendingOfferList.SortOrder;
            }
        }
        else
        {
            sortingText = " ORDER BY OA.IncentiveID DESC";
        }
    }
    private void GetSearchText()
    {
        string searchText = Search.SearchText.Trim();
        if (searchText.Length > 0)
        {
            bool result = int.TryParse(searchText, out matchOfferId);
            if (!result)
            {
                matchOfferId = -1;
            }
            offer_Name = searchText;
        }
    }
    private int GetCollisionCount(long offerID)
    {
        int count = 0;
        AMSResult<int> reportCount = offerService.GetCollidingOfferCount(offerID, adminUserID, 9);
        if(reportCount.ResultType == AMSResultType.Success && reportCount.Result > 0)
        {
            count = reportCount.Result;
        }
        return count;
    }
    private void CallScript(string key, string script)
    {
        ScriptManager.RegisterStartupScript(this, this.GetType(), key, script, true);
    }
    #endregion


    
}