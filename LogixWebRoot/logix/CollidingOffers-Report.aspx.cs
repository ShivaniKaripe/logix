using System;
using System.Collections.Generic;
using System.Data;
using System.Web.UI.WebControls;
using CMS.AMS;
using CMS.AMS.Contract;
using CMS.AMS.Models;
using System.Web;
using System.Web.UI;
using CMS.AMS.Models;

public partial class logix_CollidingOffers_list : AuthenticatedUI
{
    string sortingText = "";
    long iOfferID_CDS;
    public String CollisionDetectionServiceURL;
    Copient.CommonInc MyCommon = new Copient.CommonInc();
    ICollisionDetectionService m_CollisionDetectionService;
    IOffer offer;
    CMS.AMS.Models.OCD.Offer obj;
    protected int AdminUserID;
    int pageIndex = 1;
    const String LogFile = "Sammy_Test.txt";


    protected void Page_Load(object sender, EventArgs e)
    {
        infobar.Attributes["style"] = "display: none;";
        infobar.InnerText = String.Empty;
        statusbar.Attributes["style"] = "display: none;";
        infobar.InnerText = String.Empty;
        deleteItems.Text = PhraseLib.Lookup("term.deleteallitems", LanguageID);
        cancel.Text = PhraseLib.Lookup("term.cancel", LanguageID);
        
        copyPG.Text = PhraseLib.Lookup("term.copyPG", LanguageID);
        copyPG.ToolTip = PhraseLib.Lookup("term.copyPG", LanguageID);
        cancel1.Text = PhraseLib.Lookup("term.cancel", LanguageID);
        cancel1.ToolTip = PhraseLib.Lookup("term.cancel", LanguageID);
        canceldeploy.Text = PhraseLib.Lookup("term.cancelcollisiondetection", LanguageID);
        if (CurrentUser.UserPermissions.EditOffer == false)
        {
          rdResolution.Items[0].Enabled = false;
          rdResolution.Items[1].Enabled = false;
        }
        if (!Page.IsPostBack)
        {
            AdminUserID = ((AuthenticatedUI)this.Page).CurrentUser.AdminUser.ID;
            ((logix_LogixMasterPage)this.Master).Tab_Name = "2_7";
            AssignPageTitle("term.collisionreport");
            if (!string.IsNullOrEmpty(Request.QueryString["ID"]))
                iOfferID_CDS = Convert.ToInt64(Request.QueryString["ID"]);
            else
                Response.Redirect("CollidingOffers-list.aspx");

            lblOfferID.Text = iOfferID_CDS.ToString();
            CollisionDetectionServiceURL = SystemCacheData.GetSystemOption_UE_ByOptionId(185).TrimEnd('/');
            InitializePhrases();
            BindGrid(Convert.ToInt32(lblOfferID.Text), pageIndex, "ExtProductID", "Desc");
            BindOffer(Convert.ToInt32(lblOfferID.Text));
            CheckCDStatus();
            BindPGOfferList();
        }
     
    }

    protected override void AuthorisePage()
    {
      if (CurrentUser.UserPermissions.AccessOffers == false)
      {
        Server.Transfer("PageDenied.aspx?PhraseName=perm.offers-access&TabName=2_7", false);
      }
    }

    protected void deleteItems_click(object sender, EventArgs e)
    {
      try
      {
        Int32 removeType = 2;
        if (hdnProductList.Value == String.Empty)
        {
          removeType = 1;
        }
        Int64.TryParse(Request.QueryString["ID"], out iOfferID_CDS);
        var ajaxProcessingFunctions = new AjaxProcessingFunctions();
        AMSResult<Boolean> response = ajaxProcessingFunctions.RemoveCollidingProducts(Convert.ToInt32(iOfferID_CDS), removeType);
        if (response.ResultType != AMSResultType.Success)
        {
          infobar.Attributes["style"] = "";
          infobar.InnerText = response.MessageString;
          return;
        }
        m_CollisionDetectionService = CurrentRequest.Resolver.Resolve<ICollisionDetectionService>();
        AMSResult<int> newcount = m_CollisionDetectionService.DetectOfferCollision(Convert.ToInt64(lblOfferID.Text), AdminUserID);
        if (newcount.ResultType == AMSResultType.Success)
        {
          if (newcount.Result == 0)
          {
            Response.Redirect("CollidingOffers-list.aspx");
          }
          else
          {
            Response.Redirect("CollidingOffers-Report.aspx?ID=" + Convert.ToInt64(lblOfferID.Text));
          }
        }

      }
      catch (Exception ex)
      {
        infobar.Attributes["style"] = "";
        infobar.InnerText = ex.Message;
      }
    }

    protected void btnCancelDeploy_Click(object sender, EventArgs e)
    {
      try
      {
        var ajaxProcessingFunctions = new AjaxProcessingFunctions();
        AMSResult<Boolean> response = ajaxProcessingFunctions.UpdateCollideOfferStatus(iOfferID_CDS, 2);
        if (response.ResultType == AMSResultType.Success)
        {
          canceldeploy.Attributes["style"] = "display: none;";
          infobar.Attributes["style"] = "";
          infobar.InnerText = PhraseLib.Lookup("term.collisiondetectioncancelled", LanguageID);
          statusbar.Attributes["style"] = "display: none;";
          if (hdnIsPGEmptyAfterResolution.Value == "true")
          {
            rdResolution.Items[0].Attributes.Add("Disabled", "");
          }
        }
        else
        {
          infobar.Attributes["style"] = "";
          infobar.InnerText = response.MessageString;
        }

      }
      catch (Exception ex)
      {
        infobar.Attributes["style"] = "";
        infobar.InnerText = ex.Message;
      }
    }

    private void InitializePhrases()
    {
        //Initializing Phrases
        rdResolution.Items[0].Text = PhraseLib.Lookup("term.removecollidingproducts", LanguageID);
        rdResolution.Items[1].Text = PhraseLib.Lookup("term.resolvecollisionsmanually", LanguageID); //"Edit offer to manually resolve collisions";
        rdResolution.Items[2].Text = PhraseLib.Lookup("term.reruncollisionreport", LanguageID); //"Re-run collision report";

        //Assigning Headings
        lblOfferID.Text = Convert.ToString(iOfferID_CDS);

        //Initializing grid sorting parameters
        hdnSortkey.Value = "ExtProductID";
        hdnSortorder.Value = "DESC";
        gvData.SortKey = "ExtProductID";
        gvData.SortOrder = "Desc";

    }

    private void BindGrid(int offerId, int pageIndex, string sortKey, string sortOrder)
    {
        ICollisionDetectionService collisionService;
        
        AMSResult<CMS.AMS.Models.OCD.ProductList> lst = new AMSResult<CMS.AMS.Models.OCD.ProductList>();
        collisionService = CurrentRequest.Resolver.Resolve<ICollisionDetectionService>();
        lst = collisionService.GetCollidingProducts(offerId, pageIndex, sortKey, sortOrder);
        //Assign Header Name
        for (int i = 0; i < gvData.Columns.Count; i++)
        {
            switch (i)
            {
                case 0:
                    gvData.Columns[i].HeaderText = PhraseLib.Lookup("term.productid", LanguageID);//"Product ID";// 
                    break;
                //case 1:
                //  gvData.Columns[i].HeaderText = PhraseLib.Lookup("term.type", LanguageID);//"Type"; //
                //  break;
                case 1:
                    gvData.Columns[i].HeaderText = PhraseLib.Lookup("term.description", LanguageID);//"Description";//
                    break;
                case 3:
                    gvData.Columns[i].HeaderText = PhraseLib.Lookup("term.offerid", LanguageID);//"Offer ID";// 
                    break;
            }
        }
        if (lst.ResultType == AMSResultType.Success && lst.Result!=null && lst.Result.ProductsCount > 0)
        {
            lblCount.Text = lst.Result.ProductsCount.ToString();
            gvData.DataSource = lst.Result.Products;
            gvData.DataBind();
            //hdnSortkey.Value = gvData.SortKey;
            //hdnSortorder.Value = gvData.SortOrder;
            ((HiddenField)up.FindControl("hdnSortkey")).Value = gvData.SortKey;
            ((HiddenField)up.FindControl("hdnSortorder")).Value = gvData.SortOrder;
        }
    }

    protected void gvData_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            List<CMS.AMS.Models.OCD.Report.Offer> obj = new List<CMS.AMS.Models.OCD.Report.Offer>();
            obj = ((CMS.AMS.Models.OCD.Product)(e.Row.DataItem)).Offers;
            if (obj.Count > 0)
            {
                GridView gvOffers = (GridView)(e.Row.FindControl("gvOffers"));
                gvOffers.DataSource = obj;
                gvOffers.DataBind();
                Int64 count = ((CMS.AMS.Models.OCD.Product)(e.Row.DataItem)).OfferCount;
                if (count > 100)
                {
                    Label lblOfferCount = (Label)(e.Row.FindControl("lblOfferCount"));
                    count = count - 100;
                    lblOfferCount.Text = "+" + count.ToString() + " " + PhraseLib.Lookup("term.more", LanguageID);
                    lblOfferCount.Visible = true;
                }
            }
        }
    }

    protected void gvData_Sorting(object sender, GridViewSortEventArgs e)
    {
        ((HiddenField)up.FindControl("hdnPageIndex")).Value = "1";
        BindGrid(Convert.ToInt32(lblOfferID.Text), pageIndex, gvData.SortKey, gvData.SortOrder);
    }

    private void BindOffer(int offerId)
    {
        ICollisionDetectionService collisionService;
        AMSResult<CMS.AMS.Models.OCD.Offer> lst = new AMSResult<CMS.AMS.Models.OCD.Offer>();
        collisionService = CurrentRequest.Resolver.Resolve<ICollisionDetectionService>();
        lst = collisionService.GetOfferDetail(offerId);
        if (lst.ResultType == AMSResultType.Success)
        {
            obj = new CMS.AMS.Models.OCD.Offer();
            obj = (CMS.AMS.Models.OCD.Offer)lst.Result;
            if (obj != null)
            {
                lblOfferName.Text = MyCommon.TruncateString(obj.IncentiveName, 37);
                if (!string.IsNullOrEmpty(obj.BuyerID))
                    lblBID.Text = obj.BuyerID.ToString();
                lblDescription.Text = obj.OfferDescription;
                lblEndDate.Text = obj.EndDate.ToString("MM/dd/yy");
                lblExtID.Text = obj.ClientOfferID;
                lblName.Text = obj.IncentiveName;
                lblStartDate.Text = obj.StartDate.ToString("MM/dd/yy");
                lblID.Text = obj.IncentiveID.ToString();
                lblReportDate.Text = obj.CollisionRanOn.ToString();
            }
            else
                Response.Redirect("CollidingOffers-list.aspx");
        }
    }

    private void CheckCDStatus()
    {
        MyCommon.AppName = "CollidingOffers-Report.aspx";
        CurrentRequest.Resolver.AppName = MyCommon.AppName;
        m_CollisionDetectionService = CurrentRequest.Resolver.Resolve<ICollisionDetectionService>();
        AMSResult<CMS.AMS.Models.OCD.QueueStatus> AwaitingDetectionResp = m_CollisionDetectionService.GetOfferQueueStatus(iOfferID_CDS);
        if ((AwaitingDetectionResp.ResultType == AMSResultType.Success && (AwaitingDetectionResp.Result == CMS.AMS.Models.OCD.QueueStatus.NotStarted || AwaitingDetectionResp.Result == CMS.AMS.Models.OCD.QueueStatus.InProgress)))
        {
            RegisterScript("reRunCD", "javascript:reRunCDStatus();");
        }
        else
        {


            CheckOfferStatus();

        }

    }

    private void CheckOfferStatus()
    {
        if (obj != null)
        {

            if (obj.StatusFlag == 1 && obj.ReportUpdated == false)
            {
                RegisterScript("OfferChangeStatus", "javascript:OfferChangeStatus();");
            }
            else
            {
                //IsPGEmptyRemoveCollision
                IsPGEmptyAfterResolution();
            }
        }
    }

    private void IsPGEmptyAfterResolution()
    {
        m_CollisionDetectionService = CurrentRequest.Resolver.Resolve<ICollisionDetectionService>();
        AMSResult<bool> status = m_CollisionDetectionService.IsPGEmptyAfterResolution(1, Convert.ToInt32(lblOfferID.Text));
        if (CurrentUser.UserPermissions.EditOffer == true && status.Result == true)
        {
            RegisterScript("OfferChangeStatus", "javascript:PGEmptyAfterResolution();");
            hdnIsPGEmptyAfterResolution.Value = "true";
        }
        else if (CurrentUser.UserPermissions.EditOffer == false && status.Result == true)
        {
          RegisterScript("OfferChangeStatus", "javascript:NoEditPermission();");
        }
    }

    private void RegisterScript(string key, string script)
    {
        // Get a ClientScriptManager reference from the Page class.
        ClientScriptManager cs = Page.ClientScript;

        // Check to see if the startup script is already registered.
        if (!cs.IsStartupScriptRegistered(this.GetType(), key))
        {
            cs.RegisterStartupScript(this.GetType(), key, script, true);
        }
    }

    private void BindPGOfferList()
    {
        IOffer OfferProductGroups = CurrentRequest.Resolver.Resolve<IOffer>();
        string Errormsg = string.Empty;
        AMSResult<List<string>> amsresult = new AMSResult<List<string>>();
        amsresult = OfferProductGroups.GetOfferPGs(iOfferID_CDS);
        if ((amsresult.ResultType == AMSResultType.Success))
        {
            hdnProductList.Value = amsresult.Result[0].ToString();
            hdnOfferList.Value = amsresult.Result[1].ToString();

        }
        else
        {
            // Send(Copient.PhraseLib.Lookup(amsresult.MessageString, LanguageID));
        }
    }

    [System.Web.Services.WebMethod]
    public static AMSResult<CMS.AMS.Models.OCD.ProductList> GetCollidingProducts(Int32 OfferID, Int32 pageIndex, String sortKey, String sortOrder)
    {
        var ajaxProcessingFunctions = new AjaxProcessingFunctions();
        return ajaxProcessingFunctions.GetCollidingProducts(OfferID, pageIndex, sortKey, sortOrder);
    }

}