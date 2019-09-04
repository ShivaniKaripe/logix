using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using CMS.AMS;
using CMS.AMS.Contract;
using CMS.AMS.Models;
using System.Data;

public partial class logix_configurator_UEoffer_rew_proximitymsgpreview : AuthenticatedUI
{
    #region Global Variables
    private string Message;
    #endregion

    #region Protected Methods
    protected void Page_Load(object sender, EventArgs e)
    {
        GetQueryString();
        proximityMessage.InnerText = Message;
    }
    #endregion

    #region Private Methods
    private void GetQueryString()
    {
        Message = Request.QueryString["Message"];
    }
    #endregion
}