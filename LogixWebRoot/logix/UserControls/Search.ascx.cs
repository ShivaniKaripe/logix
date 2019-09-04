using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class logix_UserControls_Search : System.Web.UI.UserControl
{
  public event EventHandler OnSearch;
  public string SearchText
  {
    get
    {
      return txtSearch.Text;
    }
    set
    {
      if (value != null)
        txtSearch.Text = value;
    }
  }
  public string SerachButtonCaption
  {
    get;
    set;
  }
  public string SerachButtonToolTip
  {
    get;
    set;
  }
  public Button  SearchButton
  {
    get
    {

      return btnSearch;
    }
  }
  protected override void OnLoad(EventArgs e)
  {
    Localized();
    base.OnLoad(e);
    
  }
  protected void btnSearch_Click(object sender, EventArgs e)
  {
    if (OnSearch != null)
      OnSearch(sender, e);
  }
  private void Localized()
  {
    if (string.IsNullOrEmpty(SerachButtonCaption))
    {

      btnSearch.Text = (this.Page as AuthenticatedUI).PhraseLib.Lookup("term.Search", (this.Page as AuthenticatedUI).LanguageID);
    }
    else
      btnSearch.Text = SerachButtonCaption;


    if (!string.IsNullOrEmpty(SerachButtonToolTip))
    {
      btnSearch.ToolTip = SerachButtonToolTip;
    }
  }
}