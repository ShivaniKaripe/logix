using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.ComponentModel;
using  CMS;
public partial class logix_UserControls_Paging : System.Web.UI.UserControl
{
  public event EventHandler OnNextPageClick;
  public event EventHandler OnLastPageClick;
  public event EventHandler OnFirstPageClick;
  public event EventHandler OnPreviousPageClick;
  [DefaultValue(20)]
  public int PageSize
  {
    get
    {

      return ViewState[this.ID + "PageZise"].ConvertToInt32();
    }
    set
    {
      ViewState[this.ID + "PageZise"] = value;
    }
  }
  public int PageIndex {
    get
    {

      return ViewState[this.ID + "PageIndex"].ConvertToInt32();
    }
    set
    {
      ViewState[this.ID + "PageIndex"] = value;
    }
  }
  public int RecordCount
  {
    get
    {

      return ViewState[this.ID + "RecordCount"].ConvertToInt32();
    }
    set
    {
      ViewState[this.ID + "RecordCount"] = value;
    }
  }
  private int CurrentPageStartIndex
  {
    get
    {
      return (PageIndex * PageSize) + 1;
    }
  }
  private int CurrentPageEndIndex
  {
    get
    {
      int endindex = (PageIndex + 1) * PageSize;
      return endindex < RecordCount ? endindex : RecordCount;
    }
  }
  private int PageCount
  {
    get
    {
      int remainder = RecordCount%PageSize;
      if(remainder>0)
         return (RecordCount/PageSize)+1;
      else
          return (RecordCount/PageSize);
  
     
    }
  }
  public void DataBind()
  {

    int pageIndex = this.PageIndex;
    int pageCount = this.PageCount;

    lnkFirst.Visible = lnkPrevious.Visible = (pageIndex > 0);
    lnkNext.Visible = lnkLast.Visible = (pageIndex < (pageCount - 1));

    lblFirst.Visible = lblPrevious.Visible = !(pageIndex > 0);
    lblNext.Visible = lblLast.Visible = !(pageIndex < (pageCount - 1));
    if (RecordCount == 0)
    {
      lblPage.Text = (this.Page as AuthenticatedUI).PhraseLib.Lookup("term.noresults", (this.Page as AuthenticatedUI).LanguageID);
    }
    else
      lblPage.Text = "<b>" + CurrentPageStartIndex.ToString() + "</b> - <b>" + CurrentPageEndIndex.ToString() + "</b> " + (this.Page as AuthenticatedUI).PhraseLib.Lookup("term.of", (this.Page as AuthenticatedUI).LanguageID) + " <b>" + RecordCount.ToString() + "</b>";
  }
  protected void Page_Load(object sender, EventArgs e)
  {
    lnkFirst.Text = "<b>|</b>◄" + (this.Page as AuthenticatedUI).PhraseLib.Lookup("term.first", (this.Page as AuthenticatedUI).LanguageID);
    lblFirst.Text = "<b>|</b>◄" + (this.Page as AuthenticatedUI).PhraseLib.Lookup("term.first", (this.Page as AuthenticatedUI).LanguageID);
    lnkPrevious.Text = "◄" + (this.Page as AuthenticatedUI).PhraseLib.Lookup("term.previous", (this.Page as AuthenticatedUI).LanguageID);
    lblPrevious.Text = "◄" + (this.Page as AuthenticatedUI).PhraseLib.Lookup("term.previous", (this.Page as AuthenticatedUI).LanguageID);
    lnkNext.Text = (this.Page as AuthenticatedUI).PhraseLib.Lookup("term.next", (this.Page as AuthenticatedUI).LanguageID) + "►";
    lblNext.Text = (this.Page as AuthenticatedUI).PhraseLib.Lookup("term.next", (this.Page as AuthenticatedUI).LanguageID) + "►";
    lnkLast.Text = (this.Page as AuthenticatedUI).PhraseLib.Lookup("term.last", (this.Page as AuthenticatedUI).LanguageID) + "►<b>|</b>";
    lblLast.Text = (this.Page as AuthenticatedUI).PhraseLib.Lookup("term.last", (this.Page as AuthenticatedUI).LanguageID) + "►<b>|</b>";
  }
  protected void lnkFirst_Click(object sender, EventArgs e)
  {
    PageIndex = 0;
    if (OnFirstPageClick!=null)
      OnFirstPageClick(this, e);
  }
  protected void lnkPrevious_Click(object sender, EventArgs e)
  {
    PageIndex =PageIndex- 1;
    if (OnPreviousPageClick != null)
      OnPreviousPageClick(this, e);
  }
  protected void lnkNext_Click(object sender, EventArgs e)
  {
    PageIndex = PageIndex + 1;
    if (OnNextPageClick != null)
      OnNextPageClick(this, e);
  }
  protected void lnkLast_Click(object sender, EventArgs e)
  {
    PageIndex = PageCount-1;
    if (OnLastPageClick != null)
      OnLastPageClick(this, e);
  }
}