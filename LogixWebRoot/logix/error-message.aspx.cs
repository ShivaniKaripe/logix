using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class logix_error_message : AuthenticatedUI
{
    protected void Page_Load(object sender, EventArgs e)
    {
      if (PreviousPage != null)
        this.Title = PreviousPage.Title;
      if (Request.QueryString["TabName"] != null)
      {

        (this.Master as logix_LogixMasterPage).Tab_Name = Request.QueryString["TabName"].ToString();
      }
      if (Request.QueryString["MainHeading"] != null)
      {

        MainHeading = Request.QueryString["MainHeading"].ToString();
      }
      if (Request.QueryString["ErrorMessage"] != null)
      {

        ErrorMessage = Request.QueryString["ErrorMessage"].ToString();
      }
      

    }
    protected  string MainHeading    {
      get;
      set;

    }
    protected string ErrorMessage
    {
      get;
      set;

    }
}