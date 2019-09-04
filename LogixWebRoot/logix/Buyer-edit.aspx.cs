using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using CMS.AMS.Models;
using CMS.AMS.Contract;
using CMS.AMS;
using CMS.Models;


public partial class logix_Buyer_edit : AuthenticatedUI
{
    #region Private Variables
    private IBuyerRoleData m_buyerRoleService;
    private IDepartment m_departmentService;
    private IAdminUserData m_adminUserDataService;
    #endregion Private Variables

    AMSResult<string> folderName = new AMSResult<string>();
    public int folderid = 0;

    #region Properties
    private Buyer buyer
    {
        get
        {
            return ViewState["buyer"] as Buyer;
        }
        set
        {
            ViewState["buyer"] = value;
        }
    }
    private List<AdminUser> ALLAdminuser
    {
        get
        {
            return Session["AllAdminuser"] as List<AdminUser>;
        }
        set
        {
            Session["AllAdminuser"] = value;
        }
    }
    private List<AdminUser> AvailableFilteredAdminUser
    {
        get
        {
            return Session["AvailableFilteredAdminUser"] as List<AdminUser>;
        }
        set
        {
            Session["AvailableFilteredAdminUser"] = value;
        }
    }
    private List<AdminUser> IncludeAdminUser
    {
        get
        {
            return ViewState["IncludeAdminUser"] as List<AdminUser>;

        }
        set
        {
            ViewState["IncludeAdminUser"] = value;
        }
    }
    private List<PHNode> AllDepartments
    {
        get
        {
            return Session["AllDepartments"] as List<PHNode>;
        }
        set
        {
            Session["AllDepartments"] = value;
        }
    }
    private List<PHNode> AvailableFilteredDepartments
    {
        get
        {
            return Session["AvailableFilteredDepartments"] as List<PHNode>;
        }
        set
        {
            Session["AvailableFilteredDepartments"] = value;
        }
    }
    private List<PHNode> IncludedDepartments
    {
        get
        {
            return ViewState["IncludedDepartments"] as List<PHNode>;
        }
        set
        {
            ViewState["IncludedDepartments"] = value;
        }
    }
    #endregion

    #region Protected Functions
    protected override void OnInit(EventArgs e)
    {
        AppName = "Buyer-edit.aspx";
        base.OnInit(e);
    }

    protected void Page_Load(object sender, EventArgs e)
    {
        try
        {
            lblDefaultFolder.Text = String.Concat(PhraseLib.Lookup("term.defaultfolder", LanguageID), ":").Trim();
            ResolveDependencies();
            if (!Page.IsPostBack)
            {
                (this.Master as logix_LogixMasterPage).Tab_Name = "8_4";
                buyer = VerifyPageUrl();
                FillPageControlText(buyer);
                SetAvailableDataForAdminUsers();
                SetAvailableDataForDepartments();

            }
            ucNotes_Popup.NotesUpdate += new EventHandler(ucNotes_Popup_NotesUpdate);
            SetUpUserControls();
        }
        catch (Exception ex)
        {
            DisplayMessage(ex.Message);
        }
    }

    private void SetHeightForListbox()
    {
        if (lstDeptAvailable.Items.Count > 0)
        {
            lstDeptAvailable.Style.Remove("width");
            lstDeptAvailable.Rows = lstDeptAvailable.Items.Count;
        }
        else
        {
            lstDeptAvailable.Style.Add("width", "330px");
            lstDeptAvailable.Rows = 13;
        }

        if (lstDeptSelected.Items.Count > 0)
        {
            lstDeptSelected.Style.Remove("width");
            if (lstDeptSelected.Items.Count > 6)
                lstDeptSelected.Rows = lstDeptSelected.Items.Count;
            else
            {
                lstDeptSelected.Rows = 6;
            }
        }
        else
        {
            lstDeptSelected.Style.Add("width", "330px");
            lstDeptSelected.Rows = 6;
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

    protected void ucNotes_Popup_NotesUpdate(object sender, EventArgs e)
    {
        try
        {
            ucNotesUI.reloadNotesSrc();
        }
        catch (Exception ex)
        {
            DisplayMessage(ex.Message); ;
        }
    }

    protected void btnDelete_Click(Object sender, EventArgs e)
    {
        if (!Page.IsValid)
            return;
        try
        {
            if (CurrentUser.UserPermissions.EditBuyerRoles == false)
            {
                //infobar.InnerHtml = "User does not have Edit Permission";
                //infobar.Style["display"] = "block";
                DisplayMessage("User does not have Edit Permission");
            }
            else
            {

                AMSResult<bool> result;
                if (buyer != null)
                    result = m_buyerRoleService.DeleteBuyerRoleByExternalId(buyer.ExternalID);
                else
                    result = m_buyerRoleService.DeleteBuyerRoleByExternalId(txtName.Text);
                if (result.ResultType == AMSResultType.Success)
                {
                    DisplayMessage(result.MessageString, true);
                    Response.Redirect("Buyer-list.aspx", false);
                }
                else
                    DisplayMessage(result.MessageString);
            }
        }
        catch (Exception ex)
        {
            //infobar.InnerText = ErrorHandler.ProcessError(ex);
            //infobar.Visible = true;
            DisplayMessage(ex.Message);
        }
    }

    protected void btnSave_Click(Object sender, EventArgs e)
    {
        string logMsg = String.Empty;
        try
        {
            if (CurrentUser.UserPermissions.EditBuyerRoles == false)
            {
                DisplayMessage("User does not have Edit Permission");
            }
            else
            {
                int id = 0;

                Buyer newbuyer = new Buyer();
                newbuyer.AdminUser = IncludeAdminUser;
                BuyerDepts buyerdepts = new BuyerDepts();
                buyerdepts.Departments = IncludedDepartments;
                newbuyer.Department = buyerdepts;
                newbuyer.ExternalID = txtName.Text;

                if (buyer != null)
                {
                    AMSResult<bool> result = m_buyerRoleService.UpdateBuyerRoleByInternalId(buyer.ID, newbuyer, LanguageID);
                    updateFolder(buyer.ID);
                    if (result.ResultType == AMSResultType.Success)
                    {
                        DisplayMessage(result.MessageString, true);
                        if (newbuyer.ExternalID != buyer.ExternalID)
                          htitle.InnerText = buyer == null ? PhraseLib.Lookup("term.new", LanguageID) + " " + PhraseLib.Lookup("term.buyer", LanguageID) : PhraseLib.Lookup("term.buyers", LanguageID) + " #" + buyer.ID + ":" + " " + newbuyer.ExternalID;
                        Response.Redirect("buyer-edit.aspx?externalbuyerid=" + Server.UrlEncode(newbuyer.ExternalID) + "&id=" + buyer.ID);
                    }
                    else
                        DisplayMessage(result.MessageString);
                }
                else
                {
                    AMSResult<int> result = m_buyerRoleService.CreateBuyerRole(newbuyer, LanguageID);
                    if (result.ResultType == AMSResultType.Success)
                    {
                        AMSResult<Buyer> buyercreated = m_buyerRoleService.LookupBuyerRoleByInternalId(result.Result);
                        DisplayMessage(result.MessageString, true);
                        if (templist.Value != "")
                        {
                            id = result.Result;
                            int folderid = Convert.ToInt32(templist.Value);
                            m_buyerRoleService.updateFolder(id, folderid);
                            TempFolder.Value = folderName.Result;
                            ScriptManager.RegisterClientScriptBlock(this.Page, this.GetType(), "myfun1", "UpdateFolderName()", true);
                        }
                        Response.Redirect("buyer-edit.aspx?externalbuyerid=" + Server.UrlEncode(buyercreated.Result.ExternalID) + "&id=" + buyercreated.Result.ID);
                    }
                    else
                    {
                        DisplayMessage(result.MessageString);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            //infobar.InnerText = ErrorHandler.ProcessError(ex);
            //infobar.Visible = true;
            DisplayMessage(ex.Message);
        }
    }

    protected void btnselect_Click(object sender, EventArgs e)
    {
        if (lstAvailable.SelectedItem != null)
        {
            foreach (var i in lstAvailable.GetSelectedIndices())
            {
                IncludeAdminUser.Add(AvailableFilteredAdminUser[i]);
            }
            SetAvailableDataForAdminUsers();
        }
        if (buyer != null)
        {
            updateFolder(buyer.ID);
        }
        else
        {
            TempFolder.Value = folderName.Result;
            ScriptManager.RegisterClientScriptBlock(this.Page, this.GetType(), "myfun1", "UpdateFolderName()", true);
        }
    }

    protected void btndeselect_Click(object sender, EventArgs e)
    {
        if (lstSelected.SelectedItem != null)
        {
            int[] i = lstSelected.GetSelectedIndices();
            int temp = 0;
            foreach (var item in i)
            {
                IncludeAdminUser.RemoveAt(item);
                if (i.Count() > temp + 1 && i[temp + 1] != null)
                    i[temp + 1] = i[temp + 1] - 1;
                temp++;
            }
            SetAvailableDataForAdminUsers();
        }
        if (buyer != null)
        {
            updateFolder(buyer.ID);
        }
        else
        {
            TempFolder.Value = folderName.Result;
            ScriptManager.RegisterClientScriptBlock(this.Page, this.GetType(), "myfun1", "UpdateFolderName()", true);
        }
    }

    protected void btnselectdept_Click(object sender, EventArgs e)
    {
        if (lstDeptAvailable.SelectedItem != null)
        {
            foreach (var i in lstDeptAvailable.GetSelectedIndices())
            {
                IncludedDepartments.Add(AvailableFilteredDepartments[i]);
            }
            SetAvailableDataForDepartments();
        }
        if (buyer != null)
        {
            updateFolder(buyer.ID);
        }
        else
        {
            TempFolder.Value = folderName.Result;
            ScriptManager.RegisterClientScriptBlock(this.Page, this.GetType(), "myfun1", "UpdateFolderName()", true);
        }
    }

    protected void btndeselectdept_Click(object sender, EventArgs e)
    {
        if (lstDeptSelected.SelectedItem != null)
        {
            int[] i = lstDeptSelected.GetSelectedIndices();
            int temp = 0;
            foreach (var item in i)
            {
                IncludedDepartments.RemoveAt(item);
                if (i.Count() > temp + 1 && i[temp + 1] != null)
                    i[temp + 1] = i[temp + 1] - 1;
                temp++;
            }
            SetAvailableDataForDepartments();
        }
        if (buyer != null)
        {
            updateFolder(buyer.ID);
        }
        else
        {
            TempFolder.Value = folderName.Result;
            ScriptManager.RegisterClientScriptBlock(this.Page, this.GetType(), "myfun1", "UpdateFolderName()", true);
        }
    }

    #endregion Protected Functions

    #region Private Functions
    private void FillPageControlText(Buyer buyer)
    {
        Copient.CommonInc MyCommon = new Copient.CommonInc();
        btnSave.Text = PhraseLib.Lookup("term.save", LanguageID);
        btnDelete.Text = PhraseLib.Lookup("term.delete", LanguageID);
        btnDelete.Visible = (buyer == null) || (CurrentUser.UserPermissions.EditBuyerRoles == false) ? false : true;
        btnSave.Visible = CurrentUser.UserPermissions.EditBuyerRoles == false ? false : true;
        AssignPageTitle("term.buyer", string.Empty, buyer == null ? string.Empty : buyer.ID.ToString());
        htitle.InnerText = buyer == null ? PhraseLib.Lookup("term.new", LanguageID) + " " + PhraseLib.Lookup("term.buyer", LanguageID) : PhraseLib.Lookup("term.buyer", LanguageID) + " #" + buyer.ID + ":" + " " + buyer.ExternalID;
        hidentification.InnerText = PhraseLib.Lookup("term.identification", LanguageID);
        lblName.Text = String.Concat(PhraseLib.Lookup("term.buyerroleidentifier", LanguageID), ":");
        lblUsers.Text = String.Concat(PhraseLib.Lookup("term.users", LanguageID), ":");
        lblDefaultFolder.Text = String.Concat(PhraseLib.Lookup("term.defaultfolder", LanguageID), ":");
        btnselect.Text = PhraseLib.Lookup("term.select", LanguageID) + " " + "▼";
        btndeselect.Text = PhraseLib.Lookup("term.deselect", LanguageID) + " " + "▲";
        hdepartments.InnerText = PhraseLib.Lookup("term.departments", LanguageID);
        btnselectdept.Text = PhraseLib.Lookup("term.select", LanguageID) + " " + "▼";
        btndeselectdept.Text = PhraseLib.Lookup("term.deselect", LanguageID) + " " + "▲";
        ucNotesUI.Visible = buyer == null ? false : MyCommon.Fetch_SystemOption(75).Equals("1") ? true : false;
        txtName.Text = buyer == null ? "" : buyer.ExternalID;
        if (buyer != null)
        {
            Dictionary<int, string> buyerFolder = m_buyerRoleService.GetFolderNameByBuyerId(buyer.ID).Result;
            string folderNames = string.Empty;
            if (buyerFolder.Values.Count > 0)
                folderNames = buyerFolder.FirstOrDefault().Value;
            TempFolder.Value = folderNames;
        }
        else
        {
            if (templist.Value != "")
            {
                int folderid = Convert.ToInt32(templist.Value);
                AMSResult<string> foldername = m_buyerRoleService.GetFolderNameByFolderId(folderid);
                TempFolder.Value = foldername.Result;
                ScriptManager.RegisterClientScriptBlock(this.Page, this.GetType(), "myfun1", "UpdateFolderName()", true);
            }
        }
    }

    private void SetUpUserControls()
    {
        if (buyer != null)
        {
            ucNotesUI.NoteType = NoteTypes.Buyer;
            ucNotesUI.LinkID = buyer.ID;
            ucNotes_Popup.NoteType = NoteTypes.Buyer;
            ucNotes_Popup.LinkID = buyer.ID;
            ucNotes_Popup.ActivityType = ActivityTypes.Buyer;
        }
    }

    private void updateFolder(int id)
    {
        Dictionary<int, string> buyerFolder = m_buyerRoleService.GetFolderNameByBuyerId(id).Result;
        string folderNames = string.Empty;
        if (buyerFolder.Values.Count > 0)
            folderNames = buyerFolder.FirstOrDefault().Value;
        TempFolder.Value = folderNames;
        ScriptManager.RegisterClientScriptBlock(this.Page, this.GetType(), "myfun1", "UpdateFolderName()", true);
    }

    private void DisplayMessage(string Msg, bool message = false)
    {
        if (message)
            infobar.Attributes["class"] = "green-background";
        else
            infobar.Attributes["class"] = "red-background";
        infobar.InnerHtml = Msg;
        infobar.Style["display"] = "block";
    }

    private void ResolveDependencies()
    {
        m_adminUserDataService = CurrentRequest.Resolver.Resolve<IAdminUserData>();
        m_buyerRoleService = CurrentRequest.Resolver.Resolve<IBuyerRoleData>();
        m_departmentService = CurrentRequest.Resolver.Resolve<IDepartment>();
    }

    private Buyer VerifyPageUrl()
    {
        string paramval = Request.QueryString["externalbuyerid"];
        if (!String.IsNullOrEmpty(paramval))
        {
            buyer = GetBuyerByExternalID(paramval);
        }
        return buyer;
    }

    private Buyer GetBuyerByExternalID(string externalID)
    {
        int id = Int32.Parse(Request.QueryString["id"]);
        AMSResult<Buyer> buyer = m_buyerRoleService.LookupBuyerRoleByInternalId(id);
        if (buyer.ResultType != AMSResultType.Success)
        {
            return null;
        }
        else
        {
            IncludeAdminUser = buyer.Result.AdminUser.Where(a => a != null).ToList();
            foreach (var admin in IncludeAdminUser)
                admin.FirstName = admin.FirstName + " " + admin.LastName + "(" + admin.UserName + ")";
            IncludedDepartments = buyer.Result.Department.Departments.Where(d => d != null).ToList();
            foreach (var dept in IncludedDepartments)
                dept.ExternalID = dept.RootHierarhcyName + ":" + dept.ExternalID;
            return (Buyer)buyer.Result;
        }
    }

    private void GetAdminUsers()
    {
        ALLAdminuser = m_adminUserDataService.GetAvailableAdminUsers().Result;
        foreach (var admin in ALLAdminuser)
            admin.FirstName = admin.FirstName + " " + admin.LastName + "(" + admin.UserName + ")";

    }

    private void SetAvailableDataForAdminUsers()
    {
        GetAdminUsers();
        if (IncludeAdminUser == null)
            IncludeAdminUser = new List<AdminUser>();
        AvailableFilteredAdminUser = ALLAdminuser.Where(a => !IncludeAdminUser.Any(inc => inc.ID == a.ID)).ToList();
        lstSelected.DataSource = IncludeAdminUser;
        lstSelected.DataBind();
        lstAvailable.DataSource = AvailableFilteredAdminUser;
        lstAvailable.DataBind();

    }

    private void SetAvailableDataForDepartments()
    {
        GetDepartments();
        if (IncludedDepartments == null)
            IncludedDepartments = new List<PHNode>();
        AvailableFilteredDepartments = AllDepartments.Where(d => !IncludedDepartments.Any(inc => inc.NodeID == d.NodeID)).ToList();
        lstDeptSelected.DataSource = IncludedDepartments;
        lstDeptSelected.DataBind();

        foreach (ListItem item in lstDeptSelected.Items)
            item.Attributes["title"] = item.Text;

        lstDeptAvailable.DataSource = AvailableFilteredDepartments;
        lstDeptAvailable.DataBind();
        foreach (ListItem item in lstDeptAvailable.Items)
            item.Attributes["title"] = item.Text;
        //SetHeightForListbox();

    }

    private void GetDepartments()
    {
        AllDepartments = m_departmentService.GetProductDepartments().Result;
        foreach (var dept in AllDepartments)
            dept.ExternalID = dept.RootHierarhcyName + ":" + dept.ExternalID;

    }

    private void GetFolderNames()
    {
        if (buyer != null)
        {
            string query = "select distinct FolderID from FolderItems as FI (NoLock)  where LinkID=" + buyer.ID + " and LinkTypeID=2;";
        }
    }
    #endregion Private Functions
}