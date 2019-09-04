using System;
using System.Collections.Generic;
using System.Data;
using System.Web.UI.WebControls;
using CMS.AMS;
using CMS.AMS.Contract;
using CMS.AMS.Models;

public partial class logix_buyer_list : AuthenticatedUI
{
  private IBuyerRoleData buyerdataservice;
  private bool searchtext;
  private int BuyerID = -1;
  private string ProgramName = string.Empty;
  private int pageIndex = 0;
  private int pageSize = 20;
  private int startRowNum = 0;
  private int RecordCount = 0;

  private string sortingText = "";
  internal AMSResult<List<Buyer>> buyers = null;

  protected void Page_Load(object sender, EventArgs e)
  {
    lblTitle.Text = Copient.PhraseLib.Lookup("term.buyers", LanguageID);
    newBtn.Text = Copient.PhraseLib.Lookup("term.new", LanguageID);
    this.Form.DefaultButton = (ListBar1.SearchControl.SearchButton.UniqueID);
    ((logix_LogixMasterPage)this.Master).Tab_Name = "8_4";
    AssignPageTitle("term.Buyers");
    infobar.Style["display"] = "none";
    ListBar1.PageingControl.PageSize = 20;
    ListBar1.SearchControl.OnSearch += new EventHandler(btnSearch_Click);
    ListBar1.PageingControl.OnFirstPageClick += new EventHandler(PageingControl_OnFirstPageClick);
    ListBar1.PageingControl.OnPreviousPageClick += new EventHandler(PageingControl_OnFirstPageClick);
    ListBar1.PageingControl.OnNextPageClick += new EventHandler(PageingControl_OnFirstPageClick);
    ListBar1.PageingControl.OnLastPageClick += new EventHandler(PageingControl_OnFirstPageClick);
    GetSearchText();
    GetSortingText();
    if (!IsPostBack)
    {
      gvCouponProgramList.SortKey = "b.externalbuyerid";
      gvCouponProgramList.SortOrder = "Desc";
      FetchData(0);
    }
  }

  protected override void AuthorisePage()
  {
    if (CurrentUser.UserPermissions.ViewBuyerRoles == false)
    {
      Server.Transfer("PageDenied.aspx?" + "View buyer Roles" + "&TabName=8_4", false);
      return;
    }
  }

  protected void newBtn_Click(object sender, EventArgs e)
  {
    Response.Redirect("Buyer-edit.aspx", false);
  }

  private void GetSearchText()
  {
    string searchText = ListBar1.SearchControl.SearchText.Trim();
    if (searchText.Length > 0)
    {
      bool result = int.TryParse(searchText, out BuyerID);
      if (!result)
      {
        BuyerID = -1;
      }
      ProgramName = searchText;
    }
  }

  protected void btnSearch_Click(object sender, EventArgs e)
  {
    try
    {
      FetchData(0);
      //if (gvCouponProgramList.Rows.Count == 1)
      //{
      //    Response.Redirect("~\\logix\\tcp-edit.aspx?tcprogramid=" + gvCouponProgramList.Rows[0].Cells[0].Text, false);
      //}
    }
    catch (Exception ex)
    {
      DisplayError(ErrorHandler.ProcessError(ex));
    }
  }

  private void PageingControl_OnFirstPageClick(object sender, EventArgs e)
  {
    FetchData(ListBar1.PageingControl.PageIndex);
  }

  private void FetchData(int pageIndex)
  {
    buyerdataservice = CurrentRequest.Resolver.Resolve<IBuyerRoleData>();

    buyers = buyerdataservice.GetAllAvailableBuyerRoles(pageIndex, sortingText, ListBar1.PageingControl.PageSize, out RecordCount);
    if (buyers.Result == null)
    {
      gvCouponProgramList.DataSource = null;
      gvCouponProgramList.DataBind();
    }
    else
    {
      if (ListBar1.SearchControl.SearchText == "")
      {
        gvCouponProgramList.DataSource = ReturnBuyersData();
        gvCouponProgramList.DataBind();       
      }
      else
      {
        DataTable dt = ReturnBuyersData();
        DataTable dnew = new DataTable();
        dnew.Columns.Add("ID");
        dnew.Columns.Add("externalbuyerid");
        dnew.Columns.Add("encodedexternalbuyerid");
        dnew.Columns.Add("FirstName");
        dnew.Columns.Add("LastName");
        dnew.Columns.Add("UserName");
        dnew.Columns.Add("Departments");
        dnew.Columns.Add("Lastupdated");
        foreach (DataRow row in dt.Rows)
        {
          if (row["ID"].ToString().Contains(ListBar1.SearchControl.SearchText) || row["ID"].ToString().Equals(ListBar1.SearchControl.SearchText))
          {
            dnew.ImportRow(row);
          }
          else if (row["externalbuyerid"].ToString().Contains(ListBar1.SearchControl.SearchText) || row["externalbuyerid"].ToString().Equals(ListBar1.SearchControl.SearchText))
          {
            dnew.ImportRow(row);
          }
          else if (row["FirstName"].ToString().Contains(ListBar1.SearchControl.SearchText) || row["FirstName"].ToString().Equals(ListBar1.SearchControl.SearchText))
          {
            dnew.ImportRow(row);
          }
          else if (row["LastName"].ToString().Contains(ListBar1.SearchControl.SearchText) || row["LastName"].ToString().Equals(ListBar1.SearchControl.SearchText))
          {
            dnew.ImportRow(row);
          }
          else if (row["UserName"].ToString().Contains(ListBar1.SearchControl.SearchText) || row["UserName"].ToString().Equals(ListBar1.SearchControl.SearchText))
          {
            dnew.ImportRow(row);
          }
        }
        gvCouponProgramList.DataSource = dnew;
        gvCouponProgramList.DataBind();
       
      }
    }

       ListBar1.PageingControl.RecordCount = RecordCount;
        ListBar1.PageingControl.PageIndex = pageIndex;
        ListBar1.PageingControl.DataBind();
  }

  private DataTable ReturnBuyersData()
  {
    DataTable dt = new DataTable();
    dt.Columns.Add("ID");
    dt.Columns.Add("externalbuyerid");
    dt.Columns.Add("encodedexternalbuyerid");
    dt.Columns.Add("FirstName");
    dt.Columns.Add("LastName");
    dt.Columns.Add("UserName");
    dt.Columns.Add("Departments");
    dt.Columns.Add("Lastupdated");

    foreach (var buyer in buyers.Result)
    {
      int id = buyer.ID;
      string extId = buyer.ExternalID;
      string Firstname = String.Empty;
      string Lasttname = String.Empty;
      string departments = string.Empty;
      string Username = string.Empty;
      string lastEdited = buyer.LastUpdated.ToString();
      if (buyer.Department.Departments.Count > 0)
      {
        foreach (var item in buyer.Department.Departments)
        {
          departments += item.ExternalID + "<br/>";
        }
        departments = departments.Remove(departments.Length - 1);
      }
      if (buyer.AdminUser.Count > 0)
      {
        foreach (var user in buyer.AdminUser)
        {
          if (user != null)
          {
            Username += user.UserName + "<br/>";
            Firstname += user.FirstName + "<br/>";
            Lasttname += user.LastName + "<br/>";
          }
        }
        if (!String.IsNullOrEmpty(Firstname) && !string.IsNullOrEmpty(Lasttname) && !string.IsNullOrEmpty(Username))
        {
          Firstname = Firstname.Remove(Firstname.Length - 1);
          Lasttname = Lasttname.Remove(Lasttname.Length - 1);
          Username = Username.Remove(Username.Length - 1);
        }
      }
      dt.Rows.Add(id, extId, Server.UrlEncode(extId), Firstname, Lasttname, Username, departments, lastEdited);
    }
    return dt;
  }

  private void DisplayError(string err)
  {
    infobar.InnerHtml = err;
    infobar.Style["display"] = "block";
  }

  private void GetSortingText()
  {
    if (gvCouponProgramList.SortKey.Length > 0 && gvCouponProgramList.SortOrder.Length > 0)
    {
      sortingText = " ORDER BY " + gvCouponProgramList.SortKey + " " + gvCouponProgramList.SortOrder;
    }
    else
    {
      sortingText = " ORDER BY b.BuyerId DESC";
    }
  }

  protected void gvCouponProgramList_Sorting(object sender, GridViewSortEventArgs e)
  {
    GetSortingText();
    FetchData(ListBar1.PageingControl.PageIndex);
  }

}