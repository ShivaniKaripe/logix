using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

using CMS.AMS;
using CMS.AMS.Contract;
using CMS.AMS.Models;

public partial class TCProgramList : AuthenticatedUI
{
  ITrackableCouponProgramService trackableCouponProgram;
  List<TrackableCouponProgram> listTCPmodel, listTCPModelPage;
  int pageSize = 20;
  int startRowNum = 0;
  int RecordCount = 0;
  string sortingText = "";
  string ProgramName = string.Empty;
  int ProgramID = -1;
  protected void Page_Load(object sender, EventArgs e)
  {
    this.Form.DefaultButton = (ListBar1.SearchControl.SearchButton.UniqueID);
    ((logix_LogixMasterPage)this.Master).Tab_Name= "5_3";
    AssignPageTitle("term.trackablecouponprogram");
    infobar.Style["display"] = "none";
    ListBar1.PageingControl.PageSize = 20;
    ListBar1.SearchControl.OnSearch +=new EventHandler(btnSearch_Click);
    ListBar1.PageingControl.OnFirstPageClick += new EventHandler(PageingControl_OnFirstPageClick);
    ListBar1.PageingControl.OnPreviousPageClick += new EventHandler(PageingControl_OnFirstPageClick);
    ListBar1.PageingControl.OnNextPageClick += new EventHandler(PageingControl_OnFirstPageClick);
    ListBar1.PageingControl.OnLastPageClick += new EventHandler(PageingControl_OnFirstPageClick);
    GetSearchText();
    GetSortingText();
    if (!IsPostBack)
    {
      gvCouponProgramList.SortKey = "ProgramID";
      gvCouponProgramList.SortOrder = "Desc";
      FetchData(0);
      newBtn.Text = PhraseLib.Lookup("term.new", LanguageID);
    }
    newBtn.Visible = CurrentUser.UserPermissions.CreateTrackableCouponPrograms;
  }

  #region Override Methods

  protected override void AuthorisePage()
  {
    if (CurrentUser.UserPermissions.AccessTrackableCouponPrograms == false)
    {
      Server.Transfer("PageDenied.aspx?PhraseName=perm.trackablecoupon-access&TabName=5_3", false);
      return;
    }
  }

  #endregion

  void PageingControl_OnFirstPageClick(object sender, EventArgs e)
  {

    FetchData(ListBar1.PageingControl.PageIndex);
  }
 
  protected void newBtn_Click(object sender, EventArgs e)
  {
    Response.Redirect("/logix/tcp-edit.aspx", false);
  }
 
  protected void gvCouponProgramList_Sorting(object sender, GridViewSortEventArgs e)
  {
    GetSortingText();
    FetchData(ListBar1.PageingControl.PageIndex);

  }  

  protected void btnSearch_Click(object sender, EventArgs e)
  {
    try {
      FetchData(0);
      if (gvCouponProgramList.Rows.Count == 1) {
        Response.Redirect("~\\logix\\tcp-edit.aspx?tcprogramid=" + gvCouponProgramList.Rows[0].Cells[0].Text, false);
      }
    }
    catch (Exception ex) {
      DisplayError(ErrorHandler.ProcessError(ex));
    }

  }

  private void FetchData(int pageIndex)
  {
    AMSResult<List<TrackableCouponProgram>> listTCPmodel;
    trackableCouponProgram = CurrentRequest.Resolver.Resolve<ITrackableCouponProgramService>();

    for (int i = 0; i < gvCouponProgramList.Columns.Count; i++) {
      switch (i) {
        case 0:
          gvCouponProgramList.Columns[i].HeaderText = PhraseLib.Lookup("term.id", LanguageID);
          break;
        case 1:
          gvCouponProgramList.Columns[i].HeaderText = PhraseLib.Lookup("term.name", LanguageID);
          break;
        case 2:
          gvCouponProgramList.Columns[i].HeaderText = PhraseLib.Lookup("storedvalue.expiredate", LanguageID).Replace("&#39;", "'");
          break;
        case 3:
          gvCouponProgramList.Columns[i].HeaderText = PhraseLib.Lookup("term.created", LanguageID);
          break;
        case 4:
          gvCouponProgramList.Columns[i].HeaderText = PhraseLib.Lookup("term.edited", LanguageID);
          break;
      }
    }
    listTCPmodel = trackableCouponProgram.GetTrackableCouponPrograms(pageIndex, ListBar1.PageingControl.PageSize, sortingText, ProgramName, ProgramID, out RecordCount);
    if (listTCPmodel.ResultType != AMSResultType.Success)
    {
      DisplayError(listTCPmodel.GetLocalizedMessage(LanguageID));
      gvCouponProgramList.DataSource = null;
      gvCouponProgramList.DataBind();
    }
    else
    {
      gvCouponProgramList.DataSource = listTCPmodel.Result;
      gvCouponProgramList.DataBind();
      ListBar1.PageingControl.RecordCount = RecordCount;
      ListBar1.PageingControl.PageIndex = pageIndex;
      ListBar1.PageingControl.DataBind();
    }
  }

  private void GetSortingText()
  {
    if (gvCouponProgramList.SortKey.Length > 0 && gvCouponProgramList.SortOrder.Length > 0)
    {
      sortingText = " ORDER BY " + gvCouponProgramList.SortKey + " " + gvCouponProgramList.SortOrder;
    }
    else
    {
      sortingText = " ORDER BY ProgramID DESC"; 
    }
  }

  private void DisplayError(string err) {
    infobar.InnerHtml = err;
    infobar.Style["display"] = "block";
  }

  private void GetSearchText()
  {
    string searchText =ListBar1.SearchControl.SearchText.Trim();
    if (searchText.Length > 0)
    {
      bool result = int.TryParse(searchText, out ProgramID);
      if (!result)
      {
        ProgramID = -1;
      }
      ProgramName = searchText;
    }
  }

} 