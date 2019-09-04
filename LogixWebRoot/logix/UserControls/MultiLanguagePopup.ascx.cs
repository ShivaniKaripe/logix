using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;
using CMS.DB;
using CMS.AMS;
using CMS.AMS.Contract;
using CMS.AMS.Models;
using System.Data;

public partial class logix_UserControls_MultiLanguagePopup : System.Web.UI.UserControl
{
    public string MLTableName { get; set; }
    public string MLColumnName { get; set; }
    public long MLIdentifierValue { get; set; }
    public string MLITranslationColumn { get; set; }
    public string MLDefaultLanguageStandardValue { get; set; }
    //public bool FromTemplate { get; set; }
    public bool DisablePopup { get; set; }
    public bool IsMultiLanguageEnabled { get; set; }
    private int defaultLanguageId;
    ISystemSettings systemSettings;
    IDBAccess dbAccess;
    public int CustomerFacingLangID = 1;
    protected void Page_Load(object sender, EventArgs e)
    {
        if (IsMultiLanguageEnabled)
        {
            systemSettings = CurrentRequest.Resolver.Resolve<ISystemSettings>();
            dbAccess = CurrentRequest.Resolver.Resolve<IDBAccess>();
            //tbRdesc.Attributes.Add("onfocus","javascript:ShowMLI('"+ tbRdesc.ClientID + "', event, false);");
            Int32.TryParse(systemSettings.GetGeneralSystemOption(125).Result.OptionValue, out CustomerFacingLangID);
            if (!this.Page.IsPostBack)
            {
                defaultLanguageId = ((AuthenticatedUI)this.Page).LanguageID;
                LoadTranslations(MLIdentifierValue);
            }
            if(repMLIInputs != null)
            {
                foreach (RepeaterItem rep in repMLIInputs.Items)
                {
                    if (rep != null)
                    {
                        TextBox tbTranslation = ((TextBox)rep.FindControl("tbTranslation"));
                        int langId = Convert.ToInt32(((HiddenField)rep.FindControl("hfLangId")).Value);
                        if (tbTranslation != null && CustomerFacingLangID == langId)
                        {
                            if ((tbTranslation.ID.IndexOf("_default")) == -1)
                            {
                                tbTranslation.ID += "_default";
                            }
                        }
                    }
                }
            }
            
            //LoadTranslations(358);
            tbMLI.Attributes.Add("onfocus", "javascript:ShowMLI('" + tbMLI.ClientID + "','" + divML.ClientID + "','" + divMLWrap.ClientID + "', event);");
            btnMLClose.Attributes.Add("onclick", "return HideMLI('" + tbMLI.ClientID + "','" + divML.ClientID + "', event);");
            imgMLI.Attributes.Add("onmouseout", "return HideMLI('" + tbMLI.ClientID + "','" + divML.ClientID + "', event);");
            imgMLI.Attributes.Add("onmouseover", "javascript:ShowMLI('" + tbMLI.ClientID + "','" + divML.ClientID + "','" + divMLWrap.ClientID + "', event);");
            imgMLI.Attributes.Add("onclick", "javascript:ShowMLI('" + tbMLI.ClientID + "','" + divML.ClientID + "','" + divMLWrap.ClientID + "', event);");
            tbMLI.Attributes.Add("onBlur", "javascript:return findHTMLTags('" + tbMLI.ClientID + "','" + Copient.PhraseLib.Lookup("categories.invalidname", defaultLanguageId) + "',event);");
            TextBox tb = (TextBox)repMLIInputs.Controls[0].FindControl("tbTranslation_default");
            if (tb == null)
            {
                tb = (TextBox)repMLIInputs.Controls[0].FindControl("tbTranslation");
            }
            tb.Attributes.Add("onBlur", "javascript:return findHTMLTags('" + tb.ClientID + "','" + Copient.PhraseLib.Lookup("categories.invalidname", defaultLanguageId) + "',event);");
        }
        else
        {
            imgMLI.Visible = false;
            //tbMLI.Text = MLDefaultLanguageStandardValue;
        }
    }

    private void LoadTranslations(long mlIdentifierValue)
    {
        repMLIInputs.DataSource = ConstructMultiLanguageInputList(mlIdentifierValue);
        repMLIInputs.DataBind();
    }
    private List<MultiLanguageInput> ConstructMultiLanguageInputList(long mlIdentifierValue)
    {
        List<MultiLanguageInput> tempMLIList = new List<MultiLanguageInput>();
        MultiLanguageInput mli = null;
        SQLParametersList paramsList = new SQLParametersList();
        string query = "SELECT L.LanguageID, L.Name, L.MSNetCode, L.JavaLocaleCode, L.PhraseTerm, L.RightToLeftText, T." + MLColumnName + " AS Translation " +
                     " FROM Languages AS L LEFT JOIN " + MLTableName + " AS T ON T.LanguageID=L.LanguageID AND T."+MLITranslationColumn+"=@MLIdentifierValue " +
                     " WHERE L.LanguageID in (SELECT TLV.LanguageID FROM TransLanguagesCF_UE AS TLV) " +
                     " ORDER BY CASE WHEN L.LanguageID=1 THEN 1 ELSE 2 END, L.LanguageID;";
        paramsList.Add("@MLIdentifierValue", System.Data.SqlDbType.BigInt).Value = mlIdentifierValue;
        DataTable dt = dbAccess.ExecuteQuery(DataBases.LogixRT, System.Data.CommandType.Text, query, paramsList);

        foreach (DataRow row in dt.Rows)
        {
            mli = new MultiLanguageInput();
            mli.LanguageID = Convert.ToInt32(row["LanguageID"]);
            if (mli.LanguageID == CustomerFacingLangID)
                mli.Translation = MLDefaultLanguageStandardValue;
            else
                mli.Translation = row["Translation"].ToString();
            mli.Name = (this.Page as AuthenticatedUI).PhraseLib.Lookup(row["PhraseTerm"].ToString(), defaultLanguageId);
            mli.MSNetCode = row["MSNetCode"].ToString();
            mli.JavaLocaleCode = row["JavaLocaleCode"].ToString();
            mli.PhraseTerm = row["PhraseTerm"].ToString();
            mli.RightToLeftText = row["RightToLeftText"].ToString();
            mli.IdentifierValue = mlIdentifierValue;
            tempMLIList.Add(mli);
        }
        return tempMLIList;
    }
    protected void repMLIInputs_ItemDataBound(object sender, RepeaterItemEventArgs e)
    {
        if (e.Item.FindControl("tbTranslation") != null)
        {

            TextBox tbTranslationBox = (TextBox)e.Item.FindControl("tbTranslation");
            int langId = Convert.ToInt32(((HiddenField)e.Item.FindControl("hfLangId")).Value);
            if (CustomerFacingLangID == langId)
            {
                ((Label)e.Item.FindControl("lblLanguageName")).Text += " (" + (this.Page as AuthenticatedUI).PhraseLib.Lookup(0327, defaultLanguageId) + ")";
                tbTranslationBox.ID += "_default";
                tbMLI.Text = ((MultiLanguageInput)e.Item.DataItem).Translation;
            }


            if (DisablePopup)
                tbTranslationBox.Enabled = false;
        }
    }
}