using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using CMS;
using CMS.Contract;
using CMS.AMS.Contract;
using CMS.AMS;
using CMS.AMS.Models;
using Copient;
using System.Web.UI;
using System.Web.Security;
/// <summary>
/// Summary description for AuthenticatedUI
/// </summary>
public class AuthenticatedUI : BaseUI
{
  private CMS.AMS.Common common;
  private AuthLib authLib;
  private CommonInc myCommon;
  private LogixInc myLogix;
    
    private const string AntiXsrfTokenKey = "__AntiXsrfToken";
    private const string AuthToken = "AuthToken";
    private string _antiXsrfTokenValue = string.Empty;
    private string _authToken = string.Empty;
    protected void Page_Init(object sender, EventArgs e)
    {
        var requestCookieAuthToken = Request.Cookies[AuthToken];
        if (requestCookieAuthToken != null)
        {
            _authToken = requestCookieAuthToken.Value;
        }
        Page.PreLoad += master_Page_PreLoad;
    }
    protected void SetCookieAndViewState()
    {
        _antiXsrfTokenValue = Guid.NewGuid().ToString("N");
        var responseCookie = new HttpCookie(AntiXsrfTokenKey)
        {
            HttpOnly = true,
            Value = _antiXsrfTokenValue
        };
        if (FormsAuthentication.RequireSSL &&
            Request.IsSecureConnection)
        {
            responseCookie.Secure = true;
        }
        Response.Cookies.Set(responseCookie);
        ViewState[AntiXsrfTokenKey] = _antiXsrfTokenValue;
        ViewState[AuthToken] = _authToken;
    }
    protected void master_Page_PreLoad(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            SetCookieAndViewState();
        }
        else
        {
            var requestCookie = Request.Cookies[AntiXsrfTokenKey];
            Guid requestCookieGuidValue;
            if (requestCookie != null
                && Guid.TryParse(requestCookie.Value, out requestCookieGuidValue))
            {
                _antiXsrfTokenValue = requestCookie.Value;
            }
            if ((string)ViewState[AntiXsrfTokenKey] != _antiXsrfTokenValue
                || (string)ViewState[AuthToken] != _authToken)
            {
                throw new InvalidOperationException("Validation of Anti-XSRF token failed.");
            }
            else
            {
                SetCookieAndViewState();
            }
        }
    }
  protected override void OnInit(EventArgs e)
  {
    base.OnInit(e);
    authLib = CurrentRequest.Resolver.Resolve<AuthLib>();
    common = CurrentRequest.Resolver.Resolve<CMS.AMS.Common>();
    myCommon = new CommonInc();
    myLogix = new LogixInc();

    AuthenticateUser();
    AddMetaToPage();
    this.AddStyleToPage();
    base.AddStyleToPage();
    
  }
  protected override void Render(HtmlTextWriter writer)
  {

    
    AuthorisePage();
    base.Render(writer);
  }

  protected virtual void AuthorisePage()
  {
    
  }
  private  void AddStyleToPage()
  {
    List<string> lstStyles = SystemSettings.GetStyleFileNamesForAuthenicatedPages(CurrentUser.AdminUser.ID);
    foreach (string stylename in lstStyles)
    {
      Page.Header.Controls.Add(new LiteralControl(@"<link rel='stylesheet' href='/css/" + stylename + "' type='text/css' media='screen'  />" + Environment.NewLine));
       
    }
  }

  private void AuthenticateUser()
  {
    common = CurrentRequest.Resolver.Resolve<CMS.AMS.Common>();
    int AdminUserID;
    if (common.LRT_Connection_State() != System.Data.ConnectionState.Open)
      common.Open_LogixRT();

    if (myCommon.LRTadoConn.State != System.Data.ConnectionState.Open)
      myCommon.Open_LogixRT();
    string TransferKey = string.Empty;
    string Authtoken = "";
    string MyURI = string.Empty;


    //1st, check the transferkey and see if the user is being transfered into AMS from another product (PrefMan)
    if (!string.IsNullOrEmpty(GetFormValue("transferkey")))
    {
      Logger.WriteDebug("AppName=" + AppName + " - Checking the TransferKey (" + GetFormValue("transferkey") + ")", "auth.txt");

      TransferKey = GetFormValue("transferkey");
      AdminUserID = authLib.Auth_TransferKey_Verify(TransferKey);

      Logger.WriteDebug("AppName=" + AppName + " - After TransferKey_Verify AdminUserID=" + AdminUserID, "auth.txt");
      if (AdminUserID != 0)
      {
        Response.Cookies["AuthToken"].Value = Authtoken;
        CurrentUser = GetUser(AdminUserID);
        return;
      }
    }

    Authtoken = "";
    if (Request.Cookies["AuthToken"] != null) //if allready validated
    {
      Authtoken = Request.Cookies["AuthToken"].Value;
    }
    Logger.WriteDebug("AppName=" + AppName + " - AuthToken='" + Authtoken + "'   Transferkey='" + GetFormValue("transferkey") + "'", "auth.txt");
    AdminUserID = 0;
    AdminUserID = authLib.Auth_Token_Verify(Authtoken);
    Logger.WriteDebug("AppName=" + AppName + " - After checking AuthToken, AdminUserID=" + AdminUserID, "auth.txt");

    if (AdminUserID == 0)
    {
      MyURI = System.Web.HttpUtility.UrlEncode(Request.Url.AbsoluteUri);
      Response.Redirect("/logix/login.aspx?mode=invalid&bounceback=" + MyURI);
    }
    else
    {
      if (CurrentUser == null || (CurrentUser != null && CurrentUser.AdminUser.ID != AdminUserID))
        CurrentUser = GetUser(AdminUserID);
      authLib.Fetch_User(CurrentUser);
      System.Threading.Thread.CurrentThread.CurrentCulture = CurrentUser.AdminUser.Culture;
      System.Threading.Thread.CurrentThread.CurrentUICulture = CurrentUser.AdminUser.Culture;
      LanguageID = CurrentUser.AdminUser.LanguageID;
    }
    if (common.LRT_Connection_State() == System.Data.ConnectionState.Open)
      common.Close_LogixRT();
    if (myCommon.LRTadoConn.State == System.Data.ConnectionState.Open)
      myCommon.Close_LogixRT();
  }
  private User GetUser(int AdminUserId)
  {
    object common = myCommon;
    User user = new User();
    user.Type = CMS.Models.UserType.AdminUser;
    user.AdminUser.ID = AdminUserId;
    myLogix.Load_Roles(ref common, (long)AdminUserId);
    user.UserPermissions = CMS.Utilities.ConvertToObject<LogixInc.RolesStruct, Permisssions>(myLogix.UserRoles);
    LanguageID = user.AdminUser.LanguageID;
    return user;
  }

  public Permisssions RefreshUserPermissions(int AdminUserId)
  {
      return GetUser(AdminUserId).UserPermissions;
  }
}