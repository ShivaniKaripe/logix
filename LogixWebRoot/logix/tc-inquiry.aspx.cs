using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

using CMS.AMS;
using CMS.AMS.Contract;
using CMS.AMS.Models;

public partial class logix_tc_inquiry : AuthenticatedUI
{
  ITrackableCouponService trackableCouponProgram;
  int pageSize = 20;
  int startRowNum = 0;
  int RecordCount = 0;
  string sortingText = "";
  string searchParam = "";

  private const int TRACKABLE_COUPON_EXPIRE_DATE_SYSOPTION_ID = 325;
  private bool bExpireDateEnabled = false;

  protected void Page_Load(object sender, EventArgs e)
  {
    (Master as logix_LogixMasterPage).Tab_Name = "5_3_1";
    AssignPageTitle("coupon-inquiry.couponinquiry");
    infobar.Style["display"] = "none";
    this.Form.DefaultButton = btnSearch.UniqueID;
    this.txtSearch.MaxLength = 150;
    if (!IsPostBack)
    {
      EnableButton(false);
      btnAction.Text = PhraseLib.Lookup("term.actions", LanguageID) + " ▼";
      btnUnLock.Text = PhraseLib.Lookup("term.unlock", LanguageID);
      btnDelete.Text = PhraseLib.Lookup("term.delete", LanguageID);
      btnSearch.Text = PhraseLib.Lookup("term.search", LanguageID);            
        }
  }

  #region Override Methods

  protected override void AuthorisePage()
  {
    if (CurrentUser.UserPermissions.AccessTrackableCouponPrograms == false)
    {
      Server.Transfer("PageDenied.aspx?PhraseName=perm.trackablecoupon-access&TabName=5_3_1", false);
      return;
    }
  }

  #endregion

  protected string RegisterPopUpScript(long CouponID)
  {

    ScriptManager.RegisterStartupScript(this, this.GetType(), "popup" + CouponID.ToString(), "createdialog('divHistory'" + CouponID + "');", true);
    return "";

  }

  void PageingControl_OnFirstPageClick(object sender, EventArgs e)
  {
    try
    {
      GetSortingText();
      GetSearchText();
      FetchData(0);
    }
    catch (Exception ex)
    {
      DisplayError(ErrorHandler.ProcessError(ex));
    }

  }

  protected void btnSearch_Click(object sender, EventArgs e)
  {
    try
    {
      if (txtSearch.Text.Trim().Length <= 0)
      {
        DisplayError(PhraseLib.Lookup("term.searchblank", LanguageID));
        gvCouponList.DataSource = null;
        gvCouponList.DataBind();
        EnableButton(false);
        return;
      }
      GetSearchText();
      FetchData(0);
      if (RecordCount <= 0)
      {
        DisplayError(PhraseLib.Lookup("lmgrejections.none", LanguageID));
      }
    }
    catch (Exception ex)
    {
      DisplayError(ErrorHandler.ProcessError(ex));
    }
  }

  protected void gvCouponList_Sorting(object sender, GridViewSortEventArgs e)
  {
    try
    {
      GetSortingText();
      GetSearchText();
      FetchData(0);
    }
    catch (Exception ex)
    {
      DisplayError(ErrorHandler.ProcessError(ex));
    }

  }

  private void FetchData(int pageIndex)
  {
    AMSResult<List<TrackableCouponView>> listTCPmodel;
    trackableCouponProgram = CurrentRequest.Resolver.Resolve<ITrackableCouponService>();


    listTCPmodel = trackableCouponProgram.SearchTrackableCouponByCode(pageIndex, 10, sortingText, searchParam, out RecordCount);

    if (listTCPmodel.ResultType != AMSResultType.Success)
    {
      DisplayError(listTCPmodel.GetLocalizedMessage(LanguageID));
      gvCouponList.DataSource = null;
      gvCouponList.DataBind();

      EnableButton(false);
      return;
    }
    gvCouponList.Columns[1].HeaderText = PhraseLib.Lookup("term.couponcodeupc", LanguageID);
    //gvCouponList.Columns[2].HeaderText = PhraseLib.Lookup("term.programid", LanguageID);
    gvCouponList.Columns[3].HeaderText = PhraseLib.Lookup("term.programname", LanguageID);
    gvCouponList.Columns[5].HeaderText = PhraseLib.Lookup("term.originaluses", LanguageID);
    gvCouponList.Columns[4].HeaderText = PhraseLib.Lookup("tcinquiry.RemainingUses", LanguageID);
    gvCouponList.Columns[6].HeaderText = PhraseLib.Lookup("storedvalue.expiredate", LanguageID);
    gvCouponList.Columns[7].HeaderText = PhraseLib.Lookup("term.status", LanguageID);
    gvCouponList.Columns[9].HeaderText = PhraseLib.Lookup("term.history", LanguageID);

    bExpireDateEnabled = SystemCacheData.GetSystemOption_General_ByOptionId(TRACKABLE_COUPON_EXPIRE_DATE_SYSOPTION_ID).Equals("0") ? false : true;
    if (bExpireDateEnabled)
    {
       gvCouponList.Columns[6].Visible = true;
    }

    gvCouponList.DataSource = listTCPmodel.Result;
    gvCouponList.DataBind();
    if (listTCPmodel.Result.Count <= 0)
    {
      EnableButton(false);
    }
    else
    {
      hidLockStatus.Value = ((TrackableCouponView)listTCPmodel.Result[0]).Locked.ToString();
      EnableButton(true);
    }

  }

  private void GetSortingText()
  {
    if (gvCouponList.SortKey.Length > 0 && gvCouponList.SortOrder.Length > 0)
    {
      sortingText = " ORDER BY " + gvCouponList.SortKey + " " + gvCouponList.SortOrder;
    }
  }

  private void GetSearchText()
  {
    searchParam = txtSearch.Text.Trim();
  }
  protected string GetActionText(string strType)
  {
    string returnText = "";
    switch (strType)
    {
      case "Scan":
        returnText = PhraseLib.Lookup("term.scan", LanguageID);
        break;
      case "Redemption":
        returnText = PhraseLib.Lookup("term.redemption", LanguageID);
        break;
      case "Unlock":
        returnText = PhraseLib.Lookup("term.unlock", LanguageID);
        break;
      case "Query":
        returnText = PhraseLib.Lookup("term.query", LanguageID);
        break;
    }
    return returnText;
  }
  protected string GetTransactionStatus(int TransStatus)
  {
    string returnText = "";
    switch (TransStatus)
    {
      case 0:
        returnText = PhraseLib.Lookup(7211, LanguageID);
        break;
      case 1:
        returnText = PhraseLib.Lookup(7212, LanguageID);
        break;
      case 2:
        returnText = PhraseLib.Lookup(7213, LanguageID);
        break;
      case 3:
        returnText = PhraseLib.Lookup(7214, LanguageID);
        break;
      case 4:
        returnText = PhraseLib.Lookup(7215, LanguageID);
        break;
      case 5:
        returnText = PhraseLib.Lookup(7216, LanguageID);
        break;
    }
    return returnText;
  }
  private void DisplayError(string err)
  {
    infobar.Attributes["class"] = "red-background";
    infobar.InnerHtml = err;
    infobar.Style["display"] = "block";
  }

  private void EnableButton(bool enableButton)
  {
    btnDelete.Enabled = enableButton;
    btnUnLock.Enabled = enableButton;
  }

  protected void btnDelete_Click(object sender, EventArgs e)
  {

    try
    {
      foreach (GridViewRow row in gvCouponList.Rows)
      {
        CheckBox cb = (CheckBox)row.FindControl("chkSelect");
        if (cb.Checked == true)
        {
          bool LockStatus = Convert.ToBoolean(gvCouponList.DataKeys[row.DataItemIndex].Values[1].ToString());
          long CouponID = Convert.ToInt64(gvCouponList.DataKeys[row.DataItemIndex].Values[0].ToString());
          trackableCouponProgram = CurrentRequest.Resolver.Resolve<ITrackableCouponService>();
          AMSResult<bool> objResult = trackableCouponProgram.DeleteTrackableCouponById(CouponID);
          if (objResult.ResultType != AMSResultType.Success)
          {
            DisplayError(ErrorHandler.ProcessError(objResult.GetLocalizedMessage(LanguageID)));
            gvCouponList.DataSource = null;
            gvCouponList.DataBind();

          }
          GetSortingText();
          GetSearchText();
          FetchData(0);

        }
      }
    }
    catch (Exception ex)
    {
      DisplayError(ErrorHandler.ProcessError(ex));
    }
  }

  protected void btnUnLock_Click(object sender, EventArgs e)
  {
    AMSResult<byte> objResult = null;
    foreach (GridViewRow row in gvCouponList.Rows)
    {
      CheckBox cb = (CheckBox)gvCouponList.Rows[row.DataItemIndex].FindControl("chkSelect");
      if (cb.Checked == true)
      {
        string CouponCode = ((System.Web.UI.HtmlControls.HtmlGenericControl)gvCouponList.Rows[row.DataItemIndex].FindControl("tccouponcode")).InnerText.ToString();// .Cells[1].Text.Trim();
        bool LockStatus = Convert.ToBoolean(gvCouponList.DataKeys[row.DataItemIndex].Values[1].ToString());
        trackableCouponProgram = CurrentRequest.Resolver.Resolve<ITrackableCouponService>();

        if (LockStatus == true)
        {
          objResult = trackableCouponProgram.UnlockTrackableCoupon(CouponCode);

          if (objResult.ResultType == AMSResultType.Success)
          {
            GetSortingText();
            GetSearchText();
            FetchData(0);
          }
          else
          {
            DisplayError(objResult.GetLocalizedMessage(LanguageID));
            gvCouponList.DataSource = null;
            gvCouponList.DataBind();
          }
        }
        else
        {
          DisplayError(PhraseLib.Lookup("term.couponunlock", LanguageID));
          GetSortingText();
          GetSearchText();
          FetchData(0);
        }

      }
    }
  }

  protected void gvCouponList_RowCreated(object sender, GridViewRowEventArgs e)
  {
    if (e.Row.RowType == DataControlRowType.DataRow)
    {
      TrackableCouponView tv = (e.Row.DataItem as TrackableCouponView);
      if (tv == null)
        return;
      ScriptManager.RegisterStartupScript(this, this.GetType(), "popup" + tv.CouponId, "createdialog('divHistory" + tv.CouponId + "');", true);
      Repeater rep = (Repeater)e.Row.FindControl("repHistory");
      if (rep != null)
      {        
        rep.DataSource = tv.History;
        rep.DataBind();
      }

    }
  }
}