using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using CMS.AMS;
using CMS.AMS.Models;
using CMS.AMS.Contract;
using CMS;
using CMS.DB;
using System.Text;
using System.Collections;
using System.Data;
using Copient;
public partial class logix_UE_Default : AuthenticatedUI
{
    string previewStr;
    protected void Page_Load(object sender, EventArgs e)
    {
        ResolveDependencies();
        GetQueryStrings();
        AssignPageTitle("term.patternprev");
        SetControlText();
    }

    private void ResolveDependencies()
    {
        CurrentRequest.Resolver.AppName = "CouponPattern-Preview.aspx";
    }

    private void GetQueryStrings()
    {
        previewStr = Request.QueryString["pattern"].ToString();
    }

    private void SetControlText()
    {
        title.InnerText = PhraseLib.Lookup("term.coupon", LanguageID) + " " + PhraseLib.Lookup("term.pattern", LanguageID).ToLower() + " " + PhraseLib.Lookup("term.preview", LanguageID).ToLower();
        ptext.Text = PhraseLib.Lookup("couponpattern-preview", LanguageID);
        lblpatternprev.Text = Server.HtmlEncode(previewStr);
    }
}