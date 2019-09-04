using System;
using System.Collections.Generic;
using System.Web;
using System.Web.Security;
using System.Web.UI;

/// <summary>
/// Summary description for UnAuthenticatedUI
/// </summary>
public class UnAuthenticatedUI: BaseUI
{
    private const string AntiXsrfTokenKey = "__AntiXsrfToken";
    private const string AntiXsrfUserNameKey = "__AntiXsrfUserName";
    private string _antiXsrfTokenValue;
    protected void Page_Init(object sender, EventArgs e)
    {
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
        ViewState[AntiXsrfUserNameKey] = Context.User.Identity.Name ?? String.Empty;
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
            if ((string)ViewState[AntiXsrfTokenKey] != _antiXsrfTokenValue || (string)ViewState[AntiXsrfUserNameKey] != (Context.User.Identity.Name ?? String.Empty))
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
    this.AddStyleToPage();
  }
  private void AddStyleToPage()
  {
    List<string> lstStyles = SystemSettings.GetStyleFileNamesForUnauthenicatedPages();
    foreach (string stylename in lstStyles)
    {
      Page.Header.Controls.Add(new LiteralControl(@"<link rel='stylesheet' href='/css/" + stylename + "' type='text/css' media='screen'  />" + Environment.NewLine));
    }
  }
}