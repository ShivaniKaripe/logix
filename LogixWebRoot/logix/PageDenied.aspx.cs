using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class logix_PageDenied : AuthenticatedUI
{
  private string PhraseName
  {
    get;
    set;

  }
  protected override void OnInit(EventArgs e)
  {
    base.OnInit(e);
    this.AppName = "PageDenied.aspx";
  }
  protected void Page_Load(object sender, EventArgs e)
  {

    if (PreviousPage != null)
      this.Title = PreviousPage.Title;

    if (Request.QueryString["PhraseName"] != null)
    {

      PhraseName = Request.QueryString["PhraseName"].ToString();
    }
    if (Request.QueryString["TabName"] != null)
    {

      (this.Master as logix_LogixMasterPage).Tab_Name = Request.QueryString["TabName"].ToString();
    }
    if (!string.IsNullOrWhiteSpace(PhraseName))
    {
      lblError.Text = "<br/> " + Server.HtmlEncode(PhraseLib.Lookup("term.requiredpermission", LanguageID) + " : " + PhraseLib.Lookup(PhraseName, LanguageID));
    }


  }
}