using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using CMS;
using CMS.AMS.Models;
using System.IO;
using System.Text;
using CMS.AMS;
using CMS.AMS.Contract;
public class AppMenuEventArg : EventArgs
{
  public AppMenu AppMenu { get; set; }
}
public partial class UE_MasterPage : System.Web.UI.MasterPage
{
  
  public delegate void OverridePageMenu(object myObject, AppMenuEventArg args);
  public event OverridePageMenu OnOverridePageMenu;
  protected AuthenticatedUI AunthUI
  {
    get
    {
      return this.Page as AuthenticatedUI;
    }


  }
    public string Tab_Name { get; set; }
    protected void Page_Load(object sender, EventArgs e)
    {
      if(!IsPostBack)
        SetHeader();

    }
    private void SetHeader()
    {
      time.Attributes.Add("title", DateTime.Now.ToString(@"HH:mm:ss, G\MT zzz"));
      time.InnerText = DateTime.Now.ToString("HH:mm") + " | ";
      if (AunthUI.Handheld)
        date.InnerText = DateTime.Now.ToShortDateString() + " | ";
      else
        date.InnerText = DateTime.Now.ToLongDateString() + " | ";
      string Name = AunthUI.CurrentUser.AdminUser.Name;
      useredit.HRef = "/logix/user-edit.aspx?UserID=" + AunthUI.CurrentUser.AdminUser.ID;
      useredit.InnerText = (Name.Length > 20 ? Name.Substring(0, 19) + "..." : Name);
      if (!string.IsNullOrEmpty(Tab_Name))
      {
        IAppMenuService AppMenuSvc = CurrentRequest.Resolver.Resolve<IAppMenuService>();
        AMSResult<AppMenu> result = AppMenuSvc.GetApplicationMenus(AunthUI.CurrentUser, Tab_Name, Request.QueryString);
        if (result.ResultType != AMSResultType.Success)
        {
          lblErrorMsg.Text= result.GetLocalizedMessage(AunthUI.LanguageID);
          return;

        }
        AppMenu appMenu = result.Result;
        AppMenuEventArg e = new AppMenuEventArg();
        e.AppMenu = appMenu;
        if (OnOverridePageMenu != null)
          OnOverridePageMenu(this, e);
        SetUpMenu(e.AppMenu);

      }



    }
    private void SetUpMenu(AppMenu appMenu)
    {
      int counter = 1;
      foreach (CMS.AMS.Models.Menu menu in appMenu.Menus)
      {
        if (menu != null)
        {
          HyperLink hyp = new HyperLink();
          hyp.NavigateUrl = menu.NavigateURL;
          hyp.CssClass = menu.Highlighet? "on": "";
          if (counter < 10) {
            hyp.Attributes["accesskey"] = counter.ToString();
            counter++;
          }
          hyp.ID = "tab" + (menu.AppMenuID).ToString();
          hyp.ClientIDMode = ClientIDMode.Static;
          hyp.Attributes["title"] = menu.TitlePhraseID==0?menu.Caption:AunthUI.PhraseLib.Lookup(menu.TitlePhraseID, AunthUI.LanguageID);
          hyp.Text =menu.PhraseID==0?menu.Caption:AunthUI.PhraseLib.Lookup(menu.PhraseID, AunthUI.LanguageID);
          phMenu.Controls.Add(hyp);
          phMenu.Controls.Add(new LiteralControl("\n"));
          if (menu.Highlighet)
          {
            SetUpSubMenu(menu);
          }

        }
      }
    }
    private void SetUpSubMenu(CMS.AMS.Models.Menu mainmenu)
    {
      int submenukeycounter = 0;
      string[] submenuaccesskeys = {"!", "@", "#", "$", "%", "^", "&", "*", "(", ")"};
      CMS.AMS.Models.Menu HighlightMenu = null;
      foreach (CMS.AMS.Models.Menu menu in mainmenu.Menus)
      {
        if (menu != null)
        {
          HyperLink hyp = new HyperLink();
          hyp.NavigateUrl = menu.NavigateURL;
          hyp.CssClass = menu.Highlighet ? "on" : "";
          if (submenukeycounter < 10) {
            hyp.Attributes["accesskey"] = submenuaccesskeys[submenukeycounter];
            submenukeycounter++;
          }
          hyp.ID = "subtab" + (menu.AppMenuID).ToString();
          hyp.ClientIDMode = ClientIDMode.Static;
          hyp.Attributes["title"] = menu.TitlePhraseID == 0 ? menu.Caption : AunthUI.PhraseLib.Lookup(menu.TitlePhraseID, AunthUI.LanguageID);
          hyp.Text = menu.PhraseID == 0 ? menu.Caption : AunthUI.PhraseLib.Lookup(menu.PhraseID, AunthUI.LanguageID);
          phSubMenu.Controls.Add(hyp);
          phSubMenu.Controls.Add(new LiteralControl("\n"));
          if (menu.Highlighet)
          {
            HighlightMenu = menu;
          }

        }
      }
      if (HighlightMenu == null || HighlightMenu.Menus==null)
      {
        //Allow AuthorisePage to be called from ASPX page in case User does not have access to the Highlighted Sub-Tab
        if (mainmenu.AppMenuID == 8)
          return;
        MenuError();

        return;
      }
      foreach (CMS.AMS.Models.Menu menu in HighlightMenu.Menus)
      {
        HyperLink hyp = new HyperLink();
        hyp.NavigateUrl = menu.NavigateURL;
        hyp.CssClass = menu.Highlighet ? "on" : "";
        if (submenukeycounter < 10) {
          hyp.Attributes["accesskey"] = submenuaccesskeys[submenukeycounter];
          submenukeycounter++;
        }
        hyp.ID = "subtab" + (menu.AppMenuID).ToString();
        hyp.ClientIDMode = ClientIDMode.Static;
        hyp.Attributes["title"] = menu.TitlePhraseID == 0 ? menu.Caption : AunthUI.PhraseLib.Lookup(menu.TitlePhraseID, AunthUI.LanguageID);
        hyp.Attributes["style"] = "float: right; left: auto; right: 11px;";
        hyp.Text = menu.PhraseID == 0 ? menu.Caption : AunthUI.PhraseLib.Lookup(menu.PhraseID, AunthUI.LanguageID);
        phSubMenu.Controls.Add(hyp);
        phSubMenu.Controls.Add(new LiteralControl("\n"));
        if (menu.Highlighet)
        {
          HighlightMenu = menu;
        }
      }
      if (HighlightMenu == null || HighlightMenu.Menus == null)
      {
        MenuError();

        return;
      }


    }
    private void MenuError()
    {

      if (AunthUI.AppName.ToLower() != "pagedenied.aspx")
        Server.Transfer("PageDenied.aspx?PhraseName=term.accesspage&TabName="+ Tab_Name, false);
    }
  
// 
}
