using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;
using System.Text;
using CMS;
using CMS.Contract;
using CMS.AMS.Contract;
using CMS.AMS;
using CMS.AMS.Models;
using Copient;
using System.Data;

/// <summary>
/// Summary description for BaseUi
/// </summary>
public class BaseUI : System.Web.UI.Page
{
  private CommonInc myCommon;

  public bool Handheld
  {
    get;
    set;
  }
  public int LanguageID
  {
    get;
    set;
  }
  public string AppName
  {
    get;
    set;
  }

  public User CurrentUser
  {
    get
    {
      if (Session["AdminUser"] != null)
      {
        return Session["AdminUser"] as User;
      }
      else
        return null;
    }
    set
    {
      Session["AdminUser"] = value;
    }
  }
  public ILogger Logger
  {
    get;
    private set;
  }
  public IErrorHandler ErrorHandler
  {
    get;
    private set;
  }
  public IPhraseLib PhraseLib
  {
    get;
    private set;
  }
  public ISystemSettings SystemSettings
  {
    get;
    private set;
  }
  public ICacheData SystemCacheData
  {
    get;
    private set;
  }

  protected override void OnInit(EventArgs e)
  {
    Handheld = false;
    AppName = Request.Url.Segments[Page.Request.Url.Segments.GetUpperBound(0)];
    base.OnInit(e);
    ResolveDepedencies();
    myCommon = new CommonInc();
    if (myCommon.LRTadoConn.State != System.Data.ConnectionState.Open)
      myCommon.Open_LogixRT();
    if (Request != null && Request.Browser != null && Request.Browser.Platform != null && Request.ServerVariables != null && Request.ServerVariables["HTTP_USER_AGENT"] != null)
    {
      Handheld = DetectHandheld(Request.Browser["IsMobileDevice"].ConvertToBoolean(), Request.Browser.Platform, Request.ServerVariables["HTTP_USER_AGENT"].ConvertToString());
    }


  }

  public string GetFormValue(string VarName)
  {
    string TempVal = null;
    TempVal = "";
    if (Request.QueryString[VarName] == null)
    {
      TempVal = "";
    }
    else
    {
      TempVal = Request.QueryString[VarName];
    }

    if (string.IsNullOrEmpty(TempVal))
    {
      if (Request.Form[VarName] == null)
      {
        TempVal = "";
      }
      else
      {
        TempVal = Request.Form[VarName];
      }
    }

    return TempVal;

  }

  private void ResolveDepedencies()
  {
    if (string.IsNullOrWhiteSpace(AppName))
    {
      throw new Exception("App Name is not assigned");
    }
    CurrentRequest.Resolver.AppName = AppName;
    Logger = CurrentRequest.Resolver.Resolve<ILogger>();
    ErrorHandler = CurrentRequest.Resolver.Resolve<IErrorHandler>();
    SystemSettings = CurrentRequest.Resolver.Resolve<ISystemSettings>();
    PhraseLib = CurrentRequest.Resolver.Resolve<IPhraseLib>();
    SystemCacheData = CurrentRequest.Resolver.Resolve<ICacheData>();
  }
  private bool DetectHandheld(bool MobileDevice, string Platform, string UserAgent)
  {
    bool retval = false;

    if (MobileDevice)
      retval = true;
    else if (Platform.IndexOf("WinCE") > -1 || Platform.IndexOf("Palm") > -1 || Platform.IndexOf("Pocket") > -1)
      retval = true;
    else if (UserAgent.IndexOf("iPhone") > -1)
      retval = true;

    return retval;
  }
  protected void AddStyleToPage()
  {


    bool EPMInstalled = false;
    CommonInc.IntegrationValues IntegrationVals = new CommonInc.IntegrationValues();
    int StyleValue = -1;

    Page.Header.Controls.Add(new LiteralControl(@"<link rel=""icon"" href=""/images/favicon.ico"" type=""image/x-icon"" />" + Environment.NewLine));
    Page.Header.Controls.Add(new LiteralControl(@"<link rel=""shortcut icon"" href=""/images/favicon.ico"" type=""image/x-icon"" />" + Environment.NewLine));
    Page.Header.Controls.Add(new LiteralControl(@"<link rel=""apple-touch-icon"" href=""/images/touchicon.png"" />" + Environment.NewLine));
    if (Handheld)
    {
      Page.Header.Controls.Add(new LiteralControl(@"<link rel='stylesheet' href='/css/logix-handheld.css' type='text/css' media='screen, handheld' />" + Environment.NewLine));
    }
    Page.Header.Controls.Add(new LiteralControl(@"<link rel='stylesheet' href='/css/logix-aural.css' type='text/css' media='aural' />" + Environment.NewLine));
    Page.Header.Controls.Add(new LiteralControl(@"<link rel='stylesheet' href='/css/logix-print.css' type='text/css' media='braille, embossed, print, tty' />" + Environment.NewLine));
    EPMInstalled = myCommon.IsIntegrationInstalled(CommonIncConfigurable.Integrations.PREFERENCE_MANAGER, ref IntegrationVals);
    StringBuilder sb = new StringBuilder();
    if (EPMInstalled)
    {

      sb.AppendLine("<style type='text/css'>");
      if (Request.Cookies["Style"] != null) Int32.TryParse(Request.Cookies["Style"].Value, out StyleValue);
      switch (StyleValue)
      {
        case 1:
          sb.AppendLine("#tabs a, #tabs a.on {background: url('/images/tab_narrow1.png') no-repeat scroll 0 0 transparent; left: 7px; width: 82px;}");
          sb.AppendLine("#tabs a:hover {background: url('/images/tab-hover_narrow1.png') no-repeat;}");
          sb.AppendLine("#tabs a.on {background: url('/images/tab-on_narrow1.png') no-repeat;}");
          sb.AppendLine("#tabs a.on:hover {background: url('/images/tab-on_narrow1.png') no-repeat;}");
          break;
        case 2:
          sb.AppendLine("#tabs a, #tabs a.on {left: 7px; width: 82px;}");
          sb.AppendLine("#tabs a:hover {background: url('/images/ncr/tab-hover_narrow1.png') no-repeat;}");
          sb.AppendLine("#tabs a.on {background: url('/images/ncr/tab-on_narrow1.png') no-repeat; font-weight: bold; height: 25px;}");
          sb.AppendLine("#tabs a.on:hover {background: url('/images/ncr/tab-on_narrow1.png') no-repeat;}");
          break;
        case 3:
          break;
        default:
          sb.AppendLine("#tabs a, #tabs a.on {left: 7px; width: 82px;}");
          sb.AppendLine("#tabs a:hover {background: url('/images/ncrgreen/tab-hover_narrow1.png') no-repeat;}");
          sb.AppendLine("#tabs a.on {background: url('/images/ncrgreen/tab-on_narrow1.png') no-repeat; font-weight: bold; height: 25px;}");
          sb.AppendLine("#tabs a.on:hover {background: url('/images/ncrgreen/tab-on_narrow1.png') no-repeat;}");
          break;
      }
      sb.AppendLine("</style>");
      Page.Header.Controls.Add(new LiteralControl(sb.ToString()));
      //Browser-specific multilanguage-input tweak
    }
    if (Request.Browser.Browser == "IE" || Request.Browser.Browser == "Opera")
    {
      sb.Clear();
      sb.AppendLine(@"<style type=""text/css"">");
      sb.AppendLine("  .ml { width: 88% !important; }");
      sb.AppendLine("</style>");
      Page.Header.Controls.Add(new LiteralControl(sb.ToString()));
    }
  }

  protected Int32 GetOfferRewardOptionID(long OfferID)
  {
    DataTable rst = new DataTable();
    myCommon.QueryStr = "select RewardOptionID from CPE_RewardOptions with (NoLock) where IncentiveID=" + OfferID + " and TouchResponse=0 and Deleted=0;";
    rst = myCommon.LRT_Select();
    if (rst.Rows.Count > 0)
      return rst.Rows[0][0].ConvertToInt32();
    else
      return -1;
  }

  protected bool IsOfferDeployed(long OfferID)
  {
    DataTable rst = new DataTable();
    myCommon.QueryStr = "SELECT StatusFlag FROM CPE_ST_Incentives WITH (NOLOCK) where IncentiveID=" + OfferID + "  and Deleted=0;";
    rst = myCommon.LRT_Select();
    if (rst != null && rst.Rows.Count > 0)
      return true;
    else
      return false;

  }
  protected void AssignPageTitle(string PageTitlePhrase = "", string PageSubTitlePhrase = "", string PageID = "")
  {

    Page.Title = GetPageTitle(PageTitlePhrase, PageSubTitlePhrase, PageID);
  }
  protected string GetPageTitle(string PageTitlePhrase = "", string PageSubTitlePhrase = "", string PageID = "")
  {
    string title = PhraseLib.Lookup("term.logix", CurrentUser.AdminUser.LanguageID);
    if (PageTitlePhrase != string.Empty)
    {
        title = title + " > " + PhraseLib.Lookup(PageTitlePhrase, CurrentUser.AdminUser.LanguageID);
    }
    if (PageID != string.Empty)
    {
      title = title + " " + CMS.Utilities.Left(PageID, 100);
    }
    if (PageSubTitlePhrase != string.Empty)
    {
        title = title + " > " + PhraseLib.Lookup(PageSubTitlePhrase, CurrentUser.AdminUser.LanguageID);
    }
    return title;
  }

  protected void AddMetaToPage()
  {
    int TempLanguageID = 0;
    if (CurrentUser == null)

      TempLanguageID = 1;

    else
      TempLanguageID = CurrentUser.AdminUser.LanguageID;
    StringBuilder sb = new StringBuilder();
    AssignPageTitle();
    string CopientFileName = Page.Request.Url.Segments[Page.Request.Url.Segments.Length - 1];
    string CopientFileVersion = "7.3.1.138972";
    string CopientProject = "Copient Logix";
    sb.AppendLine("<!-- ");
    sb.AppendLine("Project:   " + CopientProject == string.Empty ? "..." : CopientProject);
    sb.AppendLine("FileName:  " + CopientFileName == string.Empty ? "..." : CopientFileName);
    sb.AppendLine("Version:   " + CopientFileVersion == string.Empty ? "..." : CopientFileVersion);
    sb.AppendLine("Notes:     " + "...");
    sb.AppendLine("--> ");
    List<AppVersion> lstAppVersion = SystemSettings.GetInstalledVersions();

    foreach (AppVersion app in lstAppVersion)
    {
      sb.AppendLine(@"<meta name=""version"" content=""" + app.MajorVersion + "." + app.MinorVersion + " " + PhraseLib.Lookup("term.build", TempLanguageID) + " " + app.Build + "." + app.Revision);
      sb.AppendLine(@" (" + app.InstallDate.ToString("MMMM d, yyyy") + @")"" />");
    }
    sb.AppendLine(@"<meta name=""author"" content=""" + PhraseLib.Lookup("about.copientaddress", TempLanguageID) + @""" />");
    sb.AppendLine(@"<meta name=""copyright"" content=""" + PhraseLib.Lookup("about.copyright", TempLanguageID) + @""" />");
    sb.AppendLine(@"<meta name=""description"" content=""" + PhraseLib.Lookup("about.description", TempLanguageID) + @""" />");
    sb.AppendLine(@"<meta name=""content-type"" content=""text/html; charset=utf-8"" />");
    sb.AppendLine(@"<meta name=""robots"" content=""noindex, nofollow"" />");
    sb.AppendLine(@"<meta name=""viewport"" content=""width=782"" />");
    sb.AppendLine(@"<meta http-equiv=""cache-control"" content=""no-cache"" />");
    sb.AppendLine(@"<meta http-equiv=""pragma"" content=""no-cache"" />");
    sb.AppendLine(@"<meta http-equiv=""X-UA-Compatible"" content=""IE=9"" />");

    Page.Header.Controls.Add(new LiteralControl(sb.ToString()));
  }
}
