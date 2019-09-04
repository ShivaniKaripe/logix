using CMS.AMS;
using CMS.AMS.Contract;
using CMS.AMS.Models;
using Copient;
using System;
using System.Collections.Generic;
using System.Data;

public partial class logix_UE_pos_channel_imageurl : AuthenticatedUI
{

    #region Variables

    int OfferID;
    int DeliverableID;
    int PTPKID;

    bool IsNewPassThru = false;
    string Description = string.Empty;

    private string logFile = string.Format("ErrorLog.{0}.txt", DateTime.Now.ToString("yyyyMMdd"));

    private PassThrough objPassThrough
    {
        get { return ViewState["XMLPassThrough"] as PassThrough; }
        set { ViewState["XMLPassThrough"] = value; }
    }

    protected CMS.AMS.Models.Offer objOffer
    {
        get { return ViewState["Offer"] as CMS.AMS.Models.Offer; }
        set { ViewState["Offer"] = value; }
    }

    IPassThroughRewards m_PassThroughRewards;
    IOffer m_Offer;
    Copient.CommonInc MyCommon = new Copient.CommonInc();

    #endregion

    protected void Page_Load(object sender, EventArgs e)
    {
        ResolveDependencies();
        GetQueryStrings();
        if (objOffer == null)
        {
            if (OfferID == 0)
            {
                DisplayError(PhraseLib.Lookup("error.invalidoffer", LanguageID));
                return;
            }
            objOffer = m_Offer.GetOffer(OfferID, CMS.AMS.Models.LoadOfferOptions.CustomerCondition);
        }
        AssignPageTitle("mediatype.imgurl");
        if (!IsPostBack)
        {
            SetControlText();

            if (DeliverableID != 0)
            {
                AMSResult<PassThrough> result = m_PassThroughRewards.GetPassThroughReward(DeliverableID, objOffer.EngineID);
                if (result.ResultType == AMSResultType.Success)
                {
                    objPassThrough = result.Result;
                }
                else
                {
                    DisplayError(result.GetLocalizedMessage<PassThrough>(LanguageID));
                    return;
                }
                BindImageUrl();
            }
            SetPassThroughForLoad();
        }
        HideDisplayInfoMsg();
        PreviewImg.Visible = false;
    }

    protected void btnPreview_Click(object sender, EventArgs e)
    {
        PreviewImg.Visible = true;
        if (txtImageUrl.Text != "")
        {
            PreviewImg.ImageUrl = MyCommon.Fetch_UE_SystemOption(239) + txtImageUrl.Text;
        }
    }

    protected void btnSave_Click(object sender, EventArgs e)
    {
        if (string.IsNullOrWhiteSpace(txtImageUrl.Text))
        {
            DisplayError(PhraseLib.Lookup("term.enterimageurl", LanguageID));
            return;
        }
        if (CreateObjectAndSave())
        {
            btnPreview.Enabled = true;
            DisplayInfoMsg(PhraseLib.Lookup("term.changessaved", LanguageID));
        }
        else
        {
            DisplayError(PhraseLib.Lookup("term.ProcessingError", LanguageID));
        }
    }

    private void BindImageUrl()
    {
        if (objPassThrough != null)
        {
            List<PassThroughTier> TierObj = objPassThrough.TiersData;
            txtImageUrl.Text = TierObj[0].Data.ToString();
            btnPreview.Enabled = true;
        }
    }

    private void SetPassThroughForLoad()
    {
        if (objPassThrough == null)
        {
            objPassThrough = new CMS.AMS.Models.PassThrough();
            objPassThrough.TiersData = new List<CMS.AMS.Models.PassThroughTier>();
            objPassThrough.PassThroughRewardID = 0;
            objPassThrough.RewardID = DeliverableID;
            objPassThrough.RewardOptionPhase = 1;
            objPassThrough.RewardTypeID = 12;
            objPassThrough.LSInterfaceID = 2;
            objPassThrough.ActionTypeID = 0;
            objPassThrough.Required = true;
            objPassThrough.Deleted = false;
        }
    }
    private void ResolveDependencies()
    {
        CurrentRequest.Resolver.AppName = "pos-channel-imageurl.aspx";
        m_PassThroughRewards = CurrentRequest.Resolver.Resolve<CMS.AMS.Contract.IPassThroughRewards>();
        m_Offer = CurrentRequest.Resolver.Resolve<CMS.AMS.Contract.IOffer>();
    }

    private void GetQueryStrings()
    {
        Int32.TryParse(Request.QueryString["DeliverableID"], out DeliverableID);
        Int32.TryParse(Request.QueryString["PTPKID"], out PTPKID);
        Int32.TryParse(Request.QueryString["OfferID"], out OfferID);
    }

    private void SetControlText()
    {
        title.InnerText = PhraseLib.Lookup("term.offer", LanguageID) + " #" + OfferID + " " + PhraseLib.Lookup("term.posNotifications", LanguageID);
        ptext.Text = PhraseLib.Lookup("mediatype.imgurl", LanguageID);

        btnPreview.Text = PhraseLib.Lookup("term.preview", LanguageID);
        btnSave.Text = PhraseLib.Lookup("term.save", LanguageID);
    }

    private bool CreateObjectAndSave()
    {
        bool output = false;
        string errMsg = String.Empty;
        objPassThrough.TiersData.Clear();
        objPassThrough.RewardOptionPhase = 1;
        objPassThrough.RewardTypeID = 12;
        objPassThrough.Deleted = false;
        objPassThrough.Required = true;
        objPassThrough.PassThroughRewardID = 0;
        objPassThrough.LSInterfaceID = 2;
        objPassThrough.ActionTypeID = 0;

        PassThroughTier objTierData = new PassThroughTier();
        objTierData.LanguageID = LanguageID;
        objTierData.TierLevel = 1;
        objTierData.Data = txtImageUrl.Text;
        objTierData.Value = 0;

        objPassThrough.TiersData.Add(objTierData);

        IsNewPassThru = (objPassThrough.PassThroughID == 0);
        AMSResult<bool> result = m_PassThroughRewards.CreateUpdatePassThroughReward(objPassThrough, objOffer.OfferID, objOffer.EngineID);
        if (result.ResultType != AMSResultType.Success)
            DisplayError(result.GetLocalizedMessage<bool>(LanguageID));
        else
        {
            output = true;
            Description = IsNewPassThru ? "CPE_Notification.createimgurlmsg" : "CPE_Notification.editimgurlmsg";
            WriteToActivityLog(PhraseLib.Lookup(Description, LanguageID));
        }
        return output;
    }

    private void WriteToActivityLog(string Description)
    {
        if (MyCommon.LRTadoConn.State == ConnectionState.Closed)
            MyCommon.Open_LogixRT();
        MyCommon.Activity_Log(3, OfferID, CurrentUser.AdminUser.ID, Description);
        MyCommon.Close_LogixRT();
    }

    private void DisplayError(String errorText)
    {
        infobar.Attributes["class"] = "red-background";
        infobar.InnerText = errorText;
        infobar.Style["display"] = "block";
    }

    private void DisplayInfoMsg(String message)
    {
        infobar.Attributes["class"] = "green-background";
        infobar.InnerText = message;
        infobar.Style["display"] = "block";
    }

    private void HideDisplayInfoMsg()
    {
        infobar.Attributes["class"] = "orange-background";
        infobar.InnerText = "";
        infobar.Style["display"] = "none";
    }

}