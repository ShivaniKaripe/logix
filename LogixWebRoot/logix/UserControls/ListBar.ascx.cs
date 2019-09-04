using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class logix_UserControls_ListBar : System.Web.UI.UserControl
{
  public logix_UserControls_Search SearchControl
  {
    get { return ListSearch; }
  }
  public logix_UserControls_Paging PageingControl
  {
    get { return ListPaging; }
  }
  protected void Page_Load(object sender, EventArgs e)
  {
    
  }
}