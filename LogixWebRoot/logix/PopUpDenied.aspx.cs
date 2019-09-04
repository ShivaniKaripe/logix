using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class logix_Denied : AuthenticatedUI
{
  private  string PhraseName
  {
    get;
    set;
    
  }
    protected void Page_Load(object sender, EventArgs e)
    {
      
      if (PreviousPage != null)
        this.Title = PreviousPage.Title;

      if (Request.QueryString["PhraseName"] != null)
      {

        PhraseName=Request.QueryString["PhraseName"].ToString();
      }
      if(!string.IsNullOrWhiteSpace(PhraseName))
      {
        lblError.Text = "<br/> " + Server.HtmlEncode(PhraseLib.Lookup("term.requiredpermission", LanguageID) + " : " + PhraseLib.Lookup(PhraseName, LanguageID));
      }

       
    }
    protected override void OnInit(EventArgs e)
    {
      AppName = "PopUpDenied.aspx";
      base.OnInit(e);
    }

}