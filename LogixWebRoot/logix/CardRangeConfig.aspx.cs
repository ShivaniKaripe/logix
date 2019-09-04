using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.UI.WebControls;
using CMS;
using CMS.AMS.Contract;
using CMS.AMS;
using CMS.AMS.Models;

public partial class logix_CardRangeConfig : AuthenticatedUI
{
    #region Private Variables
    private ICardTypeService m_CardTypeService;
    private ICacheData m_CacheData;
    #endregion

    #region Override Methods
    protected override void AuthorisePage()
    {
        if (CurrentUser.UserPermissions.EditSystemConfiguration == false)
        {
            Server.Transfer("PageDenied.aspx?PhraseName=perm.admin-configuration&TabName=8_4", false);
            return;
        }
    }
    #endregion

    protected void Page_Load(object sender, EventArgs e)
    {
        infobar.InnerHtml = statusbar.InnerHtml = "";
        (this.Master as logix_LogixMasterPage).Tab_Name = "8_4";
        AssignPageTitle("term.cardranges");
        ResolveDependencies();
        SetPageData();
        Int32 intCardRangeConfig = m_CacheData.GetSystemOption_UE_ByOptionId(221).ConvertToInt32();
        if (intCardRangeConfig==0)
        {
            Server.Transfer("configuration.aspx", false);
            return;
        }
        if (!IsPostBack)
        {
            //Get List of Numeric Cardtypes from Cache itself for performance. Card Description is fetched from PhraseLib for Localization
            List<CardType> lstNumericCards = SystemCacheData.CardTypes.Where(m => m.NumericOnly == true && m.CardTypeID != 2 && m.CardTypeID != 8).ToList();
            ddlCardTypes.DataSource = (from item in lstNumericCards
                                      select new { Text = PhraseLib.Lookup(item.PhraseID, LanguageID).Replace("&#39;", "'") , Value = item.CardTypeID.ToString() }).ToList();
            ddlCardTypes.DataTextField = "Text";
            ddlCardTypes.DataValueField = "Value";
            ddlCardTypes.DataBind();
            populateCardTypeWithRange();            
        }
    }

    private void populateCardTypeWithRange() {
        AMSResult<List<CMS.AMS.Models.CardType>> lstCardTypesHavingRange = m_CardTypeService.GetAllNumericCardTypesWithRange();
        if (lstCardTypesHavingRange.ResultType == AMSResultType.SQLException || lstCardTypesHavingRange.ResultType == AMSResultType.Exception)
        {
            statusbar.Attributes.Add("style", "display:none"); 
            infobar.Attributes.Add("display", "block");
            infobar.InnerText = lstCardTypesHavingRange.MessageString;
            return;
        }
        repCardType.DataSource = lstCardTypesHavingRange.Result;
        repCardType.DataBind();
    }

    private void ResolveDependencies()
    {
        m_CardTypeService = CurrentRequest.Resolver.Resolve<ICardTypeService>();
        m_CacheData = CurrentRequest.Resolver.Resolve<ICacheData>();
    }

    protected void repRangeList_ItemCommand(object sender, RepeaterCommandEventArgs e)
    {
        if (e.CommandName == "Delete")
        {
            Int64 cardRangeID = Convert.ToInt64(e.CommandArgument.ConvertToLong());
            m_CardTypeService.DeleteCardRange(cardRangeID);
            populateCardTypeWithRange();
            infobar.Attributes.Add("style", "display:none");
            statusbar.Attributes.Add("style", "display:block");   
            statusbar.InnerText = PhraseLib.Lookup(8767, LanguageID);
        }
    }

    private void SetPageData()
    {
        htitle.InnerText = PhraseLib.Lookup(8764, LanguageID);
        lbStartRange.Text = PhraseLib.Lookup(8765, LanguageID) + ": ";
        lbEndRange.Text = PhraseLib.Lookup(8766, LanguageID) + ": ";
        btnAdd.Text = PhraseLib.Lookup(128, LanguageID);
        btnClear.Text = PhraseLib.Lookup(2776, LanguageID);
    }


    protected void repCardType_ItemDataBound(object sender, RepeaterItemEventArgs e)
    {
        if (e.Item.ItemType == ListItemType.Item || e.Item.ItemType == ListItemType.AlternatingItem)
        {
            Repeater repeater = (Repeater)e.Item.FindControl("RepRangeList");
            repeater.DataSource = ((CardType)(e.Item.DataItem)).lstCardRange;
            repeater.DataBind();
        }
    }

    protected void repRangeList_ItemDataBound(object sender, RepeaterItemEventArgs e)
    {
        if (e.Item.ItemType == ListItemType.Header)
        {
            Label lblStartRange = ((Label)e.Item.FindControl("lblStartRange"));
            lblStartRange.Text = PhraseLib.Lookup(8765, LanguageID);
            Label lblEndRange = ((Label)e.Item.FindControl("lblEndRange"));
            lblEndRange.Text = PhraseLib.Lookup(8766, LanguageID);
        }
        if (e.Item.ItemType == ListItemType.Item || e.Item.ItemType == ListItemType.AlternatingItem)
        {
            
            Button deleteButton = ((Button)e.Item.FindControl("btnRemove"));
            deleteButton.ToolTip = PhraseLib.Lookup(125, LanguageID);
            deleteButton.OnClientClick = String.Concat("return confirm('", PhraseLib.Lookup(8777, LanguageID), "');");
        }
    }

    protected void btnAdd_Click(object sender, EventArgs e)
    {
        try
        {
            String startRange = txtStartrange.Text;
            String endRange = txtEndRange.Text;
            Int32 maxCardLength = SystemCacheData.GetCardTypeByCardTypeID(ddlCardTypes.SelectedItem.Value.ConvertToInt32()).MaxIDLength;
            if (String.IsNullOrWhiteSpace(startRange))
            {
                statusbar.Attributes.Add("style", "display:none");
                infobar.Attributes.Add("style", "display:block");
                infobar.InnerText = PhraseLib.Lookup(8769, LanguageID);
                return;
            }
            else if (startRange == "0")
            {
                statusbar.Attributes.Add("style", "display:none");
                infobar.Attributes.Add("style", "display:block");
                infobar.InnerText = PhraseLib.Lookup(8794, LanguageID);
                return;
            }
            else if (String.IsNullOrWhiteSpace(endRange))
            {
                statusbar.Attributes.Add("style", "display:none");
                infobar.Attributes.Add("style", "display:block");
                infobar.InnerText = PhraseLib.Lookup(8770, LanguageID);
                return;
            }
            else if (endRange == "0")
            {
                statusbar.Attributes.Add("style", "display:none");
                infobar.Attributes.Add("style", "display:block");
                infobar.InnerText = PhraseLib.Lookup(8795, LanguageID);
                return;
            }
            else if (startRange.ConvertToDecimal() >= endRange.ConvertToDecimal())
            {
                statusbar.Attributes.Add("style", "display:none");
                infobar.Attributes.Add("style", "display:block");
                infobar.InnerText = PhraseLib.Lookup(8775, LanguageID);
                return;
            }
            else if (startRange.Length > maxCardLength)
            {
                statusbar.Attributes.Add("style", "display:none");
                infobar.Attributes.Add("style", "display:block");
                infobar.InnerText = String.Format(PhraseLib.Lookup(8771, LanguageID), maxCardLength);
                return;
            }
            else if (endRange.Length > maxCardLength)
            {
                statusbar.Attributes.Add("style", "display:none");
                infobar.Attributes.Add("style", "display:block");
                infobar.InnerText = String.Format(PhraseLib.Lookup(8772, LanguageID), maxCardLength);
                return;
            }
            else if (startRange.IsDigitsOnly() == false)
            {
                statusbar.Attributes.Add("style", "display:none");
                infobar.Attributes.Add("style", "display:block");
                infobar.InnerText = PhraseLib.Lookup(8773, LanguageID);
                return;
            }
            else if (endRange.IsDigitsOnly() == false)
            {
                statusbar.Attributes.Add("style", "display:none");
                infobar.Attributes.Add("style", "display:block");
                infobar.InnerText = PhraseLib.Lookup(8774, LanguageID);
                return;
            }
            else
            {
                CardRange cardRange = new CardRange() { CardTypeID = ddlCardTypes.SelectedItem.Value.ConvertToInt32(), StartRange = startRange.ConvertToDecimal(), EndRange = endRange.ConvertToDecimal() };
                m_CardTypeService.AddCardRange(cardRange);
                populateCardTypeWithRange();
                infobar.Attributes.Add("style", "display:none");
                statusbar.Attributes.Add("style", "display:block");
                statusbar.InnerText = PhraseLib.Lookup(8768, LanguageID);
                txtStartrange.Text = String.Empty;
                txtEndRange.Text = String.Empty;
            }
        }
        catch (Exception ex)
        {
            statusbar.Attributes.Add("style", "display:none");
            infobar.Attributes.Add("style", "display:block");
            infobar.InnerText = ex.Message;
            return;
        }
        
    }
}