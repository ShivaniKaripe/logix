using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;
using CMS.AMS.Contract;
using CMS.AMS.Models;
using CMS.Contract;
using CMS.DB;
using CMS.Models;
using System.Text.RegularExpressions;
using System.Data.SqlClient;
using CMS.AMS;
using CMS;

public partial class ChangePassword : UnAuthenticatedUI
{
    IAdminUserData m_adminUserDataService;
    public String Username
    {
        get { return ViewState["userName"] as String; }
        set { ViewState["userName"] = value; }
    }
    public static string GetReferrerPageName()
    {
        string functionReturnValue = null;

        if ((((System.Web.HttpContext.Current.Request.UrlReferrer) != null)))
        {
            functionReturnValue = HttpContext.Current.Request.UrlReferrer.ToString();
        }
        else
        {
            functionReturnValue = "N/A";
        }
        return functionReturnValue;
    }
    protected void Page_Load(object sender, EventArgs e)
    {
        ResolveDependencies();
        DeleteAuthTokenCookie();
        ChangePasswordPushButton.Text = PhraseLib.Lookup("term.ChangePassword", LanguageID);
        CancelPushButton.Text = PhraseLib.Lookup("term.clear", LanguageID);
        Continue.Text = PhraseLib.Lookup("term.continue", LanguageID);
        CurrentPasswordLabel.Text = PhraseLib.Lookup("term.CurrentPassword", LanguageID);
        NewPasswordLabel.Text = PhraseLib.Lookup("term.NewPassword", LanguageID);
        ConfirmNewPasswordLabel.Text = PhraseLib.Lookup("term.ConfirmNewPassword", LanguageID);
        CurrentPasswordRequired.Text = PhraseLib.Lookup("term.fieldrequired", LanguageID);
        NewPasswordRequired.Text = PhraseLib.Lookup("term.fieldrequired", LanguageID);
        ConfirmNewPasswordRequired.Text = PhraseLib.Lookup("term.fieldrequired", LanguageID);
        CurrentPasswordRequired.ErrorMessage = PhraseLib.Lookup("term.Passwordrequired", LanguageID);
        NewPasswordRequired.ErrorMessage = PhraseLib.Lookup("term.NewPasswordrequired", LanguageID);
        ConfirmNewPasswordRequired.ErrorMessage = PhraseLib.Lookup("term.ConfNewPasswordrequired", LanguageID);

        if (GetReferrerPageName() == "N/A")
        {
           // DeleteAuthTokenCookie();
            Response.Redirect("/logix/login.aspx");
            return;
        }
  
        if (!Page.IsPostBack)
        {
            
            infobar.Attributes["class"] = "red-background";
            if (Request.Form["message"] == Convert.ToInt32(PasswordValidation.PasswordExpire).ToString())
                DisplayError(PhraseLib.Lookup("term.passwordexpired", LanguageID));
            else
                DisplayError(PhraseLib.Lookup("term.Passwordchange", LanguageID));
            ViewState["userName"] = Request.Form["userName"];
            if (ViewState["userName"] != null)
            {
                Username = Request.Form["userName"];
            }
            CurrentPassword.Enabled = true;
            NewPassword.Enabled = true;
            ConfirmNewPassword.Enabled = true;
            ChangePasswordPushButton.Enabled = true;
            CancelPushButton.Enabled = true; 
        }
        else
        {
            DisplayError(null);
            infobar.Visible = false;
        }
       
    }

    private void ResolveDependencies()
    {
        CurrentRequest.Resolver.AppName = "ChangePassword.aspx";
        m_adminUserDataService = CurrentRequest.Resolver.Resolve<IAdminUserData>();
          
    }
    protected void HomePageButton_Click(object sender, EventArgs e)
    {
        Response.Redirect("/logix/login.aspx",false);
    }

    protected void ChangePasswordPushButton_Click(object sender, EventArgs e)
    {

        ValidatePassword();
    }
    private void DisplayError(String ErrorText)
    {
        infobar.InnerText = ErrorText;
        infobar.Style["display"] = "block";
        infobar.Visible = true;
    }
    // validating password from change password page
    private void ValidatePassword()
    {
       string confirmnewpassword = ConfirmNewPassword.Text;
        string currentpassword = CurrentPassword.Text;
        DataTable dt = m_adminUserDataService.GetAdminUserIDbyUserName(Username,currentpassword);

            if (dt.Rows.Count > 0)
            {
                string username = dt.Rows[0]["UserName"].ToString();
                //Validating Password
                if (ConfirmNewPassword.Text != NewPassword.Text)
                {
                    DisplayError(PhraseLib.Lookup("term.Passwordmatch", LanguageID));
                    return;
                }
               
                AMSResult<bool> u_Passwordvalid = m_adminUserDataService.ValidatePassword(confirmnewpassword, username,LanguageID,true);
                
                if (!(u_Passwordvalid.ResultType == AMSResultType.Success))
                {
                    infobar.Attributes["class"] = "red-background";
                    DisplayError(u_Passwordvalid.MessageString);
                }
                else
                {
                    //updating Password
                    AMSResult<bool> u_Passwordvalidupdate = m_adminUserDataService.UpdatePasswordLogin(confirmnewpassword, username, LanguageID);
                    if (u_Passwordvalidupdate.ResultType == AMSResultType.Success)
                    {
                        infobar.Attributes["class"] = "green-background";
                        DisplayError(u_Passwordvalidupdate.MessageString);
                        Continue.Visible = true;
                        CurrentPassword.Enabled = false;
                        NewPassword.Enabled = false;
                        ConfirmNewPassword.Enabled = false;
                        ChangePasswordPushButton.Enabled = false;
                        CancelPushButton.Enabled = false;
                    }
                    else
                    {
                        infobar.Attributes["class"] = "red-background";
                        DisplayError(u_Passwordvalidupdate.MessageString);

                    }
                }
            }
            else
            {
                infobar.Attributes["class"] = "red-background";
                DisplayError(PhraseLib.Lookup("term.invalidpassword", LanguageID));
            }
        }
    public void DeleteAuthTokenCookie()
    {
        //Delete Session and cookies for session-id and authtoken on logout
        Session.Abandon();
        Session.RemoveAll();

        if (Request.Cookies["ASP.NET_SessionId"] != null)
        {
            Response.Cookies["ASP.NET_SessionId"].Value = string.Empty;
            Response.Cookies["ASP.NET_SessionId"].Expires = DateTime.Now.AddMonths(-20);

        }
    }

 
}