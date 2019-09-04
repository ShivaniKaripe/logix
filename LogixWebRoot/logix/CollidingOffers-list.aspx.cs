using System;
using System.Collections.Generic;
using System.Data;
using System.Web.UI.WebControls;
using CMS.AMS;
using CMS.AMS.Contract;
using CMS.AMS.Models;
using CMS;

public partial class logix_CollidingOffers_list : AuthenticatedUI
{
    #region Fields
    DataTable dt = new DataTable();
    string OfferName = string.Empty;
    int ID = -1;
    public int PageSize = 50;
    public int PageIndex = 0;
    protected int AdminUserID;
    public String CollisionDetectionServiceURL;
    public String TenantID = String.Empty;
    public String OfferIDNavigate = String.Empty;
    #endregion

    protected void Page_Load(object sender, EventArgs e)
    {
        AdminUserID = ((AuthenticatedUI)this.Page).CurrentUser.AdminUser.ID;
        lblTitle.Text = Copient.PhraseLib.Lookup("term.collisionreports", LanguageID);
        ((logix_LogixMasterPage)this.Master).Tab_Name = "2_7";
        AssignPageTitle("term.collisionreports");
        infobar.Style["display"] = "none";
        ListSearch.OnSearch += new EventHandler(btnSearch_Click);
        GetSearchText();
        CollisionDetectionServiceURL = SystemCacheData.GetSystemOption_UE_ByOptionId(185).TrimEnd('/');
        if (String.IsNullOrWhiteSpace(CollisionDetectionServiceURL))
        {
          DisplayError(PhraseLib.Lookup("term.undefinedcollisionserviceurl", LanguageID));
          return;
        }
        TenantID = "0";
        if (!IsPostBack)
        {
            gvCollidingOfferList.SortKey = "CollisionRanOn";
            gvCollidingOfferList.SortOrder = "Desc";
            FetchData(PageIndex);
        }
    }

    protected void gvCollidingOfferList_Load(object sender, EventArgs e)
    {
        System.Web.UI.ScriptManager.RegisterStartupScript(this, this.GetType(), "callfunction", "gridHeaderMove()", true);
    }

    protected void btnSearch_Click(object sender, EventArgs e)
    {
        try
        {

            FetchData(PageIndex);
            if (gvCollidingOfferList.Rows.Count == 1)
            {
                Response.Redirect("\\logix\\CollidingOffers-Report.aspx?ID=" + OfferIDNavigate, false);
            }
        }
        catch (Exception ex)
        {
            DisplayError(ErrorHandler.ProcessError(ex));
        }
    }

    private void DisplayError(string err)
    {
        infobar.InnerHtml = err;
        infobar.Style["display"] = "block";
    }

    protected override void AuthorisePage()
    {
      if (CurrentUser.UserPermissions.AccessOffers == false)
      {
        Server.Transfer("PageDenied.aspx?PhraseName=perm.offers-access&TabName=2_7", false);
      }
    }

    private void GetSearchText()
    {
        string searchText = ListSearch.SearchText;
        if (searchText.Length > 0)
        {
            bool result = int.TryParse(searchText, out ID);
            if (!result)
            {
                ID = -1;
            }
            OfferName = searchText;
        }
    }

    protected void gvCollidingOfferList_Sorting(object sender, GridViewSortEventArgs e)
    {
        FetchData(PageIndex);
    }

    private void FetchData(int pageIndex)
    {
        ICollisionDetectionService collisionService;
        AMSResult<List<CMS.AMS.Models.OCD.Offer>> listcollidingOffer;
        collisionService = CurrentRequest.Resolver.Resolve<ICollisionDetectionService>();
        //String searchingText = "o";
        String searchingText = ListSearch.SearchText;
        int userid = -1;
        for (int i = 0; i < gvCollidingOfferList.Columns.Count; i++)
        {
            switch (i)
            {
                case 0:
                    gvCollidingOfferList.Columns[i].HeaderText = PhraseLib.Lookup("term.xid", LanguageID);
                    break;
                case 1:
                    gvCollidingOfferList.Columns[i].HeaderText = PhraseLib.Lookup("term.id", LanguageID);
                    break;
                case 2:
                    gvCollidingOfferList.Columns[i].HeaderText = PhraseLib.Lookup("term.buyer", LanguageID);
                    break;
                case 3:
                    gvCollidingOfferList.Columns[i].HeaderText = PhraseLib.Lookup("term.offername", LanguageID);
                    break;
                case 4:
                    gvCollidingOfferList.Columns[i].HeaderText = PhraseLib.Lookup("term.reportRun", LanguageID);
                    break;
                case 5:
                    gvCollidingOfferList.Columns[i].HeaderText = PhraseLib.Lookup("term.CollisionCount", LanguageID);
                    break;
                case 6:
                    gvCollidingOfferList.Columns[i].HeaderText = PhraseLib.Lookup("term.viewreport", LanguageID);
                    break;
            }
        }
        userid = AdminUserID;
        listcollidingOffer = collisionService.GetCollisionReports(pageIndex, gvCollidingOfferList.SortKey, gvCollidingOfferList.SortOrder, Server.UrlEncode(searchingText), userid);
        if (listcollidingOffer.ResultType != AMSResultType.Success)
        {
          if (listcollidingOffer.MessageString.Equals("ArgumentException") || listcollidingOffer.MessageString.Equals("WebException") || listcollidingOffer.ResultType == AMSResultType.UriFormatException)
            {
                DisplayError(PhraseLib.Lookup("term.CDSPath", LanguageID));
            }
            else
                DisplayError(listcollidingOffer.GetLocalizedMessage<List<CMS.AMS.Models.OCD.Offer>>(LanguageID));
            gvCollidingOfferList.DataSource = null;
            gvCollidingOfferList.DataBind();
        }
        else
        {
            gvCollidingOfferList.DataSource = listcollidingOffer.Result;
            gvCollidingOfferList.DataBind();
            if (listcollidingOffer.Result != null && listcollidingOffer.Result.Count == 1)
            {
                OfferIDNavigate = listcollidingOffer.Result[0].IncentiveID.ToString();
            }
            if (listcollidingOffer.Result != null)
            {
                Totalrecordcount.Value = listcollidingOffer.Result.Count.ToString();
            }
            sortkey.Value = gvCollidingOfferList.SortKey;
            sortorder.Value = gvCollidingOfferList.SortOrder;
            searchtext.Value = searchingText;
            pagesize.Value = PageSize.ToString();
            adminuserId.Value = userid.ToString();
        }
    }
    protected void gvCollidingOfferList_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            (e.Row.FindControl("ViewReport") as HyperLink).Text = PhraseLib.Lookup("term.viewreport", LanguageID);
        }
    }
}