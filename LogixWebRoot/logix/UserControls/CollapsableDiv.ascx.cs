using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class logix_UserControls_CollapsableDiv : System.Web.UI.UserControl
{
  public string TargetDivID { get; set; }
  public string ToolTip { get; set; }
  public bool IsCollapsed {get;set;}
  protected override void OnLoad(EventArgs e)
  {

    imgID.ID = this.ID + imgID.ID;
    imgID.Attributes.Add("onmouseover", "handleResizeHover(true,'" + TargetDivID + "','" + imgID.ClientID + "')");
    imgID.Attributes.Add("onmouseout", "handleResizeHover(false,'" + TargetDivID + "','" + imgID.ClientID + "')");
    if (IsCollapsed)
      imgID.ToolTip = "Show " + ToolTip;
    else
      imgID.ToolTip = "Hide " + ToolTip;
    
    base.OnLoad(e);
  }
}