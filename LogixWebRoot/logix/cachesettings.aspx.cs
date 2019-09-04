using System;
using System.Collections.Generic;
using System.Web.UI.WebControls;
using CMS;
using CMS.AMS.Models;
using CMS.AMS;
public partial class logix_cachesettings : AuthenticatedUI
{

  protected override void OnInit(EventArgs e) 
  {
    AppName = "cachesettings.aspx";
    base.OnInit(e);
  }

  protected void Page_Load(object sender, EventArgs e) 
  {
    infobar.InnerHtml = "";
    infobar.Visible = false;
    if (!IsPostBack) 
    {
      AssignPageTitle("term.configuration", "term.cache", "");
      title.InnerText = PhraseLib.Lookup("term.cachesettings", LanguageID);
      lbCacheInterval.Text = PhraseLib.Lookup("term.cacheinterval", LanguageID);
      txtCacheTimeout.Text = SystemSettings.GetGeneralSystemOption(165).Result.OptionValue;
      rptCache.DataSource = SystemCacheData.GetCachedObjectNames();
      rptCache.DataBind();
    }
  }

  protected override void AuthorisePage()
  {
    if (!CurrentUser.UserPermissions.AccessSystemSettings)
    {
      Server.Transfer("PopUpdenied.aspx?PhraseName=perm.admin-settings", false);
    }   
  }

  protected void rptCache_ItemCommand(object source, RepeaterCommandEventArgs e)
  {
    List<CachedObject> lstCachedObject = new List<CachedObject>();
    CachedObject objCachedObject = new CachedObject();
    objCachedObject.CachedObjectID = e.CommandArgument.ConvertToInt32();
    lstCachedObject.Add(objCachedObject);
    SystemCacheData.ClearCacheData(lstCachedObject);
    switch (objCachedObject.CachedObjectID)
    {
        case -1:    
        case 1:
            Copient.SystemOptionsCache.RemoveCache(System.Web.HttpContext.Current.Request.Url.Host);
            break;
    }
    DisplaySuccess(PhraseLib.Lookup("cache.clearcache", LanguageID));
  }

  private void DisplayError(string msg)
  {
    infobar.InnerHtml = msg;
    infobar.Visible = true;
    infobar.Attributes["class"] = "red-background";
  }

  private void DisplaySuccess(string msg)
  {
    infobar.InnerHtml = msg;
    infobar.Visible = true;
    infobar.Attributes["class"] = "green-background";
  }

  protected void lnkSave_Click(object sender, EventArgs e)
  {
    List<SystemOption> lstSystemOption = new List<SystemOption>();
    SystemOption objSystemOption = new SystemOption();
    objSystemOption.OptionId = 165;
    objSystemOption.OptionValue = txtCacheTimeout.Text;
    lstSystemOption.Add(objSystemOption);

    AMSResult<bool> Result = SystemSettings.UpdateGeneralSystemOptions(lstSystemOption);
    if (Result.ResultType != AMSResultType.Success)
      DisplayError(Result.GetLocalizedMessage(LanguageID));
    else
      DisplaySuccess(PhraseLib.Lookup("cache.caheintervalsaved", LanguageID));     
  }
}