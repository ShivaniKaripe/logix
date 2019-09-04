using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using CMS.AMS;
using CMS.AMS.Contract;
using CMS.AMS.Models;
using CMS;
using System.Web.Security;

public partial class logix_tcp_edit : AuthenticatedUI
{

    #region PageLevelVariables
    ITrackableCouponProgramService tcpService;
    IActivityLogService activityLogService;
    IOffer offerService;
    #endregion

    #region CSRF CODE

    private const string AntiXsrfTokenKey = "__AntiXsrfToken";
    private const string AntiXsrfUserNameKey = "__AntiXsrfUserName";
    private string _antiXsrfTokenValue;

    private const int TRACKABLE_COUPON_EXPIRE_DATE_SYSOPTION_ID = 325;
    private bool bExpireDateEnabled = false;

    protected void Page_Init(object sender, EventArgs e)
    {
        //First, check for the existence of the Anti-XSS cookie
        var requestCookie = Request.Cookies[AntiXsrfTokenKey];
        Guid requestCookieGuidValue;

        //If the CSRF cookie is found, parse the token from the cookie.
        //Then, set the global page variable and view state user
        //key. The global variable will be used to validate that it matches in the view state form field in the Page.PreLoad
        //method.
        if (requestCookie != null
        && Guid.TryParse(requestCookie.Value, out requestCookieGuidValue))
        {
            //Set the global token variable so the cookie value can be
            //validated against the value in the view state form field in
            //the Page.PreLoad method.
            _antiXsrfTokenValue = requestCookie.Value;

            //Set the view state user key, which will be validated by the
            //framework during each request
            Page.ViewStateUserKey = _antiXsrfTokenValue;
        }
        //If the CSRF cookie is not found, then this is a new session.
        else
        {
            //Generate a new Anti-XSRF token
            _antiXsrfTokenValue = Guid.NewGuid().ToString("N");

            //Set the view state user key, which will be validated by the
            //framework during each request
            Page.ViewStateUserKey = _antiXsrfTokenValue;

            //Create the non-persistent CSRF cookie
            var responseCookie = new HttpCookie(AntiXsrfTokenKey)
            {
                //Set the HttpOnly property to prevent the cookie from
                //being accessed by client side script
                HttpOnly = true,

                //Add the Anti-XSRF token to the cookie value
                Value = _antiXsrfTokenValue
            };

            //If we are using SSL, the cookie should be set to secure to
            //prevent it from being sent over HTTP connections
            if (FormsAuthentication.RequireSSL &&
            Request.IsSecureConnection)
                responseCookie.Secure = true;

            //Add the CSRF cookie to the response
            Response.Cookies.Set(responseCookie);
        }

        Page.PreLoad += Page_PreLoad;
    }

    protected void Page_PreLoad(object sender, EventArgs e)
    {
        ////During the initial page load, add the Anti-XSRF token and user
        ////name to the ViewState
        //if (!IsPostBack)
        //{
        //    //Set Anti-XSRF token
        //    ViewState[AntiXsrfTokenKey] = Page.ViewStateUserKey;

        //    //If a user name is assigned, set the user name
        //    ViewState[AntiXsrfUserNameKey] =
        //    Context.User.Identity.Name ?? String.Empty;
        //}
        ////During all subsequent post backs to the page, the token value from
        ////the cookie should be validated against the token in the view state
        ////form field. Additionally user name should be compared to the
        ////authenticated users name
        //else
        //{
        //    //Validate the Anti-XSRF token
        //    if ((string)ViewState[AntiXsrfTokenKey] != _antiXsrfTokenValue
        //    || (string)ViewState[AntiXsrfUserNameKey] !=
        //    (Context.User.Identity.Name ?? String.Empty))
        //    {
        //        throw new InvalidOperationException("Validation of Anti - XSRF token failed.");
        //    }
        //}
    }

    #endregion


    #region events
    protected void Page_Load(object sender, EventArgs e)
    {
        try
        {
            (this.Master as logix_LogixMasterPage).Tab_Name = "5_3_2";
            (this.Master as logix_LogixMasterPage).OnOverridePageMenu += new logix_LogixMasterPage.OverridePageMenu(logix_tcp_edit_OnOverridePageMenu);

            btnDelete.Attributes.Add("onclick", "return confirm('" + PhraseLib.Lookup("term.confirmdeleteprogram", LanguageID) + "')");
            if (!Page.IsPostBack)
            {
                //CSRF CODE ENDS

                //Set Anti-XSRF token
                ViewState[AntiXsrfTokenKey] = Page.ViewStateUserKey;

                //If a user name is assigned, set the user name
                ViewState[AntiXsrfUserNameKey] =
                Context.User.Identity.Name ?? String.Empty;

                //CSRF CODE ENDS

                TrackableCouponProgram tcProgram = VerifyPageUrl();
                FillPageControlText(tcProgram);
                AssignPageTitle("term.trackablecouponprogram", String.Empty, ProgramID.ToString());
                ApplyPermission();
                if (SystemCacheData.GetSystemOption_UE_ByOptionId(152) == "0")
                {
                    lblCouponusCount.Visible = false;
                    CouponCount.Visible = false;
                }
            }
            else
            {
                //Validate the Anti-XSRF token
                if ((string)ViewState[AntiXsrfTokenKey] != _antiXsrfTokenValue || (string)ViewState[AntiXsrfUserNameKey] != (Context.User.Identity.Name ?? String.Empty))
                {
                    throw new InvalidOperationException("Validation of Anti - XSRF token failed.");
                }
            }
            ucNotes_Popup.NotesUpdate += new EventHandler(ucNotes_Popup_NotesUpdate);
            SetUpUserControls();
            txtName.Focus();
        }
        catch (Exception ex)
        {
            DisplayError(ErrorHandler.ProcessError(ex));
        }
    }
    protected void btnSave_Click(Object sender, EventArgs e)
    {
        if (!Page.IsValid)
            return;
        bool isNew = false;
        string logMsg = String.Empty;
        try
        {

            tcpService = CurrentRequest.Resolver.Resolve<ITrackableCouponProgramService>();
            activityLogService = CurrentRequest.Resolver.Resolve<IActivityLogService>();
            if (Convert.ToInt32(ProgramID) == 0) isNew = true;
            TrackableCouponProgram tcProgramModel = new TrackableCouponProgram();
            tcProgramModel.ProgramID = Convert.ToInt32(ProgramID);
            tcProgramModel.Name = txtName.Text.Trim();
            tcProgramModel.Description = txtDescription.InnerText.Trim();
            tcProgramModel.ExtProgramID = txtExternalID.Text.Trim();
            tcProgramModel.MaxRedeemCount = Convert.ToByte(txtMaxRedempCount.Text);
            tcProgramModel.TCExpireType = Convert.ToInt32(ddlExpireTypes.SelectedValue);
            int ExpirationPeriod;
            switch (tcProgramModel.TCExpireType)
            {
                case 1:
                    tcProgramModel.ExpirePeriod = 0;
                    tcProgramModel.TCExpirePeriodType = 0;
                    tcProgramModel.ExpireDate = null;
                    break;
                case 2:
                    DateTime dtSpecified;
                    tcProgramModel.ExpirePeriod = 0;
                    tcProgramModel.TCExpirePeriodType = 0;
                    if (DateTime.TryParse(ExpireDate.Text, out dtSpecified))
                    {
                        tcProgramModel.ExpireDate = dtSpecified;
                    }
                    else
                    {
                        tcProgramModel.ExpireDate = null;
                    }
                    if (tcProgramModel.ExpireDate < DateTime.Now)
                    {
                        DisplayError(Copient.PhraseLib.Lookup("logix-js.EnterValidExpDate", LanguageID).Replace("&#39;", "\'"));
                        return;
                    }
                    break;
                case 3:
                    if (!Int32.TryParse(txtExpirationPeriod.Text, out ExpirationPeriod))
                    {
                        DisplayError(Copient.PhraseLib.Lookup("sv-edit.InvalidExpirePeriod", LanguageID));
                        return;
                    }
                    tcProgramModel.ExpirePeriod = ExpirationPeriod;
                    tcProgramModel.TCExpirePeriodType = Convert.ToInt32(ddlExpirePeriodTypes.SelectedValue);
                    tcProgramModel.ExpireDate = null;
                    break;
                case 4:
                    if (!Int32.TryParse(txtExpirationPeriod.Text, out ExpirationPeriod))
                    {
                        DisplayError(Copient.PhraseLib.Lookup("sv-edit.InvalidExpirePeriod", LanguageID));
                        return;
                    }
                    tcProgramModel.ExpirePeriod = ExpirationPeriod;
                    tcProgramModel.TCExpirePeriodType = 2;
                    tcProgramModel.ExpireDate = null;
                    break;
            }
            AMSResult<bool> retVal = tcpService.CreateUpdateTrackableCouponProgram(tcProgramModel);
            if (retVal.ResultType != AMSResultType.Success)
                DisplayError(retVal.GetLocalizedMessage<bool>(LanguageID));
            else
            {
                logMsg = isNew == true ? String.Concat(PhraseLib.Lookup("term.trackablecouponprogram", LanguageID), " ", PhraseLib.Lookup("term.created", LanguageID)) : String.Concat(PhraseLib.Lookup("term.trackablecouponprogram", LanguageID), " ", PhraseLib.Lookup("term.edited", LanguageID));
                activityLogService.Activity_Log(ActivityTypes.TCProgram, tcProgramModel.ProgramID.ConvertToInt32(), CurrentUser.AdminUser.ID, logMsg);
                Response.Redirect("~/logix/tcp-edit.aspx?tcprogramid=" + tcProgramModel.ProgramID, false);
            }
        }
        catch (Exception ex)
        {
            DisplayError(ErrorHandler.ProcessError(ex));
        }

    }
    protected void btnNew_Click(Object sender, EventArgs e)
    {
        try
        {
            Response.Redirect("~/logix/tcp-edit.aspx", false);
        }
        catch (Exception ex)
        {

            DisplayError(ErrorHandler.ProcessError(ex));
        }

    }
    protected void btnDelete_Click(Object sender, EventArgs e)
    {
        try
        {
            if (!String.IsNullOrEmpty(AssociateOfferID))
            {
                DisplayError(PhraseLib.Lookup("tcpedit.programused", LanguageID));
                return;
            }

            tcpService = CurrentRequest.Resolver.Resolve<ITrackableCouponProgramService>();
            activityLogService = CurrentRequest.Resolver.Resolve<IActivityLogService>();
            AMSResult<bool> retVal = tcpService.DeleteTrackableCouponProgram(Convert.ToInt32(ProgramID));
            if (retVal.ResultType != AMSResultType.Success)
                DisplayError(retVal.GetLocalizedMessage<bool>(LanguageID));
            else
            {
                activityLogService.Activity_Log(ActivityTypes.TCProgram, ProgramID.ConvertToInt32(), CurrentUser.AdminUser.ID, String.Concat(PhraseLib.Lookup("term.trackablecouponprogram", LanguageID), " ", PhraseLib.Lookup("term.deleted", LanguageID)));
                Response.Redirect("~/logix/tcp-list.aspx", false);
            }

        }
        catch (Exception ex)
        {

            DisplayError(ErrorHandler.ProcessError(ex));

        }

    }
    protected void ucNotes_Popup_NotesUpdate(object sender, EventArgs e)
    {
        try
        {
            ucNotesUI.reloadNotesSrc();
        }
        catch (Exception ex)
        {
            DisplayError(ErrorHandler.ProcessError(ex));
        }
    }

    protected void handleExpireDateTimeChange(Object sender, EventArgs e)
    {
        ExpireDate.Text = Server.HtmlEncode(txtDatepicker.Text) + " " + ddlExpireTimeHours.SelectedValue + ":" + ddlExpireTimeMinutes.SelectedValue;
    }
    #endregion

    #region Override Methods
    void logix_tcp_edit_OnOverridePageMenu(object myObject, AppMenuEventArg args)
    {
        AppMenu appmenu = args.AppMenu;
        //Top Level TAB
        var menu = (from m in appmenu.Menus
                    where m.Highlighet == true
                    select m).SingleOrDefault();
        if (menu == null)
            return;
        var submenu = (from m in menu.Menus
                       where m.Highlighet == true
                       select m

                      ).SingleOrDefault();
        if (submenu == null)
            return;

        if (Request.QueryString["tcprogramid"].ConvertToInt32() == 0)
        {
            var pagemenus = from m in submenu.Menus
                            where m.Highlighet == true && m.PageSpecific == true
                            select m;
            submenu.Menus = pagemenus.ToList();


        }
    }
    protected override void AuthorisePage()
    {
        if (CurrentUser.UserPermissions.AccessTrackableCouponPrograms == false)
        {
            Server.Transfer("PageDenied.aspx?PhraseName=perm.trackablecoupon-access&TabName=5_3_2", false);
            return;
        }
    }
    #endregion

    #region Properties
    public String ProgramID
    {
        get { return ViewState["ProgramID"] as String; }
        set { ViewState["ProgramID"] = value; }
    }
    public String AssociateOfferID
    {
        get { return ViewState["AssociateOfferID"] as String; }
        set { ViewState["AssociateOfferID"] = value; }
    }
    #endregion

    #region Private Methods
    private void ApplyPermission()
    {
        if (!CurrentUser.UserPermissions.CreateTrackableCouponPrograms)
        {
            btnSave.Visible = false;
            btnNew.Visible = false;
        }
        if (!CurrentUser.UserPermissions.EditTrackableCouponPrograms)
            btnUpdate.Visible = false;
        if (!CurrentUser.UserPermissions.DeleteTrackableCouponPrograms)
            btnDelete.Visible = false;

        if (!CurrentUser.UserPermissions.CreateTrackableCouponPrograms && !CurrentUser.UserPermissions.EditTrackableCouponPrograms && !CurrentUser.UserPermissions.DeleteTrackableCouponPrograms)
            btnAction.Visible = false;
    }
    private void SetUpUserControls()
    {
        ucNotesUI.NoteType = NoteTypes.TCProgram;
        ucNotesUI.LinkID = ProgramID.ConvertToInt32();
        ucNotes_Popup.NoteType = NoteTypes.TCProgram;
        ucNotes_Popup.LinkID = ProgramID.ConvertToInt32(); ;
        ucNotes_Popup.ActivityType = ActivityTypes.TCProgram;
    }
    private TrackableCouponProgram VerifyPageUrl()
    {
        int pgID = 0;
        string paramVal = Request.QueryString["tcprogramid"];
        TrackableCouponProgram tcProgram = null;
        if (!String.IsNullOrEmpty(paramVal))
        {
            if (Int32.TryParse(paramVal, out pgID))
            {
                tcProgram = GetProgramByID(pgID);
                if (tcProgram == null)
                    Server.Transfer("error-message.aspx?MainHeading=" + PhraseLib.Lookup("term.trackablecouponprogram", LanguageID) + " #" + pgID + "&ErrorMessage=" + PhraseLib.Lookup("term.itemnotfound", LanguageID) + "&TabName=5_3_2", false);
            }
            else
                Response.Redirect("~/logix/tcp-edit.aspx", false);
        }
        return tcProgram;
    }
    private void SetLastLoadMessage(TrackableCouponProgram tcProgram)
    {
        lblStatusMsg.Visible = true;
        lblCouponuploadSumm.Visible = true;
        lblStatusMsg.Text = PhraseLib.Lookup("term.statusmessage", LanguageID) + ":";
        if (tcProgram == null || tcProgram.LastLoadMsg == null)
        {
            lblStatusMsg.Visible = false;
        }
        else
        {
            lblCouponuploadSumm.Text = Copient.PhraseLib.DecodeEmbededTokens(tcProgram.LastLoadMsg, LanguageID);
        }
        if (tcProgram == null || tcProgram.LastLoaded == null)
        {
            lblCouponUploadDate.Text = PhraseLib.Lookup("tcpedit.nocouponuploadmsg", LanguageID);
            //lblCouponuploadSumm.Text = String.Empty;
            //lblStatusMsg.Text = String.Empty;
            //lblStatusMsg.Visible = false;
            //lblCouponuploadSumm.Visible = false;
        }
        else
        {
            lblCouponUploadDate.Text = tcProgram.LastLoaded.Value.ToLongDateString() + " " + tcProgram.LastLoaded.Value.ToLongTimeString();
            //lblStatusMsg.Visible = true;
            //lblCouponuploadSumm.Visible = true;
            //lblStatusMsg.Text = PhraseLib.Lookup("term.statusmessage", LanguageID) + ":";
            //lblCouponuploadSumm.Text = tcProgram.LastLoadMsg;
        }
    }

    private void FillPageControlText(TrackableCouponProgram tcProgram)
    {
        Copient.CommonInc MyCommon = new Copient.CommonInc();
        #region LabelsNeverbeChange
        btnSave.Text = PhraseLib.Lookup("term.save", LanguageID);
        hidentification.InnerText = Collapsadividentification.ToolTip = PhraseLib.Lookup("term.identification", LanguageID);
        btnUpdate.Text = PhraseLib.Lookup("term.save", LanguageID);
        btnNew.Text = PhraseLib.Lookup("term.new", LanguageID);
        btnDelete.Text = PhraseLib.Lookup("term.delete", LanguageID);
        lblExternalID.Text = String.Concat(PhraseLib.Lookup("term.externalid", LanguageID), ":");
        lblName.Text = String.Concat(PhraseLib.Lookup("term.name", LanguageID), ":");
        lblDescription.Text = String.Concat(PhraseLib.Lookup("term.description", LanguageID), ":");
        lblDescriptionLimitMsg.Text = String.Concat("(", PhraseLib.Lookup("CPEoffergen.description", LanguageID), ")");
        hRedemptioninformation.InnerText = CollapsadivRedemptioninformation.ToolTip = PhraseLib.Lookup("term.redemptioninfo", LanguageID);
        lblMaxRedempCount.Text = String.Concat(PhraseLib.Lookup("term.maxredemptioncount", LanguageID), ":");
        lblMaxMinInfoMsg.Text = String.Concat(PhraseLib.Lookup("term.minimum", LanguageID), ":1", " ", PhraseLib.Lookup("term.maximum", LanguageID), ":255");
        lblExpire.Text = String.Concat(PhraseLib.Lookup("storedvalue.expiredate", LanguageID), ":");
        hCouponUploadSumm.InnerText = CollapsableDivCouponUploadSumm.ToolTip = PhraseLib.Lookup("tcpedit.associatedcoupons", LanguageID);
        hAssociatedoffer.InnerText = CollapsableDivAssociatedoffer.ToolTip = PhraseLib.Lookup("term.associatedoffers", LanguageID);
        lblLastUpload.Text = PhraseLib.Lookup("term.lastupload", LanguageID) + ":";
        btnAction.Text = PhraseLib.Lookup("term.actions", LanguageID) + " ▼";
        lblCouponusCount.Text = PhraseLib.Lookup("tcpedit.couponscount", LanguageID) + ": ";
        lblExpire.Text = String.Concat(PhraseLib.Lookup("storedvalue.expiredate", LanguageID), ":");
        requirefieldName.ErrorMessage = PhraseLib.Lookup("tcpedit.invalidname", LanguageID);
        requirefieldExternalID.ErrorMessage = PhraseLib.Lookup("tcpedit.invalidexternalprogramid", LanguageID);
        DescriptionLengthValidator.ErrorMessage = PhraseLib.Lookup("CPEoffergen.description", LanguageID);
        RangeValidatorMaxRedempCount.ErrorMessage = PhraseLib.Lookup("tcpedit.invalidredeemcount", LanguageID);
        requirefieldMaxRedempCount.ErrorMessage = PhraseLib.Lookup("tcpedit.blankredemptioninfoerror", LanguageID);
        hExpiration.InnerText = PhraseLib.Lookup("term.expiration", LanguageID);
        lblExpirationType.Text = PhraseLib.Lookup("storedvalue.expiretype", LanguageID) + ": ";
        lblExpirationPeriodType.Text = PhraseLib.Lookup("storedvalue.expireperiodtype", LanguageID) + ": ";
        lblExpirationPeriod.Text = PhraseLib.Lookup("storedvalue.expireperiod", LanguageID) + ": ";
        lblExpirationDatePicker.Text = PhraseLib.Lookup("storedvalue.expiredate", LanguageID) + ": ";
        lblExpirationTime.Text = PhraseLib.Lookup("storedvalue.expiretime", LanguageID) + ": ";
        rvExpirePeriod.ErrorMessage = PhraseLib.Lookup("sv-edit.InvalidExpirePeriod", LanguageID);
        FillExpireType();
        FillExpirePeriodType();
        FillExpireTime();

        #endregion

        ProgramID = tcProgram == null ? "0" : tcProgram.ProgramID.ToString();
        btnSave.Visible = tcProgram == null ? true : false;
        btnAction.Visible = !btnSave.Visible;
        htitle.InnerText = tcProgram == null ? PhraseLib.Lookup("term.new", LanguageID) + " " + PhraseLib.Lookup("term.trackablecouponprogram", LanguageID).ToLower() : PhraseLib.Lookup("term.trackablecouponprogram", LanguageID) + " #" + tcProgram.ProgramID + ": " + tcProgram.Name.TruncateString(15);
        txtName.Text = tcProgram == null ? String.Empty : tcProgram.Name;
        txtDescription.InnerText = tcProgram == null ? String.Empty : tcProgram.Description.Trim();
        txtExternalID.Text = tcProgram == null ? String.Empty : tcProgram.ExtProgramID;
        ExpireDate.Text = (tcProgram == null || tcProgram.ExpireDate == null) ? PhraseLib.Lookup("tcpedit.expiredatenotset", LanguageID) : tcProgram.ExpireDate.ConvertToDate().ToShortDateString();
        txtMaxRedempCount.Text = tcProgram == null ? "1" : Convert.ToString(tcProgram.MaxRedeemCount);
        SetLastLoadMessage(tcProgram);
        CouponCount.Text = tcProgram == null ? "0" : Convert.ToString(tcProgram.AssosiatedCouponCount);
        ucNotesUI.Visible = tcProgram == null ? false : MyCommon.Fetch_SystemOption(75).Equals("1") ? true : false;
        lblAssociatedOffer.Text = PhraseLib.Lookup("term.none", LanguageID);
        AssociateOfferID = String.Empty;

        bExpireDateEnabled = SystemCacheData.GetSystemOption_General_ByOptionId(TRACKABLE_COUPON_EXPIRE_DATE_SYSOPTION_ID).Equals("1");
        ddlExpireTypes.SelectedValue = ((tcProgram == null) || (!bExpireDateEnabled)) ? "1" : tcProgram.TCExpireType.ToString();
        ddlExpirePeriodTypes.SelectedValue = tcProgram == null ? "0" : tcProgram.TCExpirePeriodType.ToString();
        txtExpirationPeriod.Text = tcProgram == null ? "0" : tcProgram.ExpirePeriod.ToString();
        ddlExpireTypes_SelectedIndexChanged(this, EventArgs.Empty);
        txtDatepicker.Text = (tcProgram == null || tcProgram.ExpireDate == null) ? "" : ExpireDate.Text;
        ddlExpireTimeHours.SelectedValue = (tcProgram == null || tcProgram.ExpireDate == null) ? "00" : tcProgram.ExpireDate.Value.Hour.ToString("00");
        ddlExpireTimeMinutes.SelectedValue = (tcProgram == null || tcProgram.ExpireDate == null) ? "00" : tcProgram.ExpireDate.Value.Minute.ToString("00");

        if (bExpireDateEnabled)
        {
            divExpiration.Visible = true;
        }

        if (tcProgram != null)
        {
            List<CMS.AMS.Models.Offer> offersObj = GetAttachedOffersByID(tcProgram.ProgramID);
            if (offersObj != null && offersObj.Count > 0)
            {
                lblAssociatedOffer.Text = String.Empty;
                offerService = CurrentRequest.Resolver.Resolve<IOffer>();
                foreach (var item in offersObj)
                {
                    AssociateOfferID = item.OfferID.ToString();
                    if (SystemCacheData.GetSystemOption_General_ByOptionId(66) == "1")
                    {
                        AMSResult<bool> ResultObj = offerService.IsAccessibleBannerEnabledOffer(CurrentUser.AdminUser.ID, item.OfferID);
                        if (ResultObj.ResultType != AMSResultType.Success)
                        {
                            DisplayError(ResultObj.GetLocalizedMessage<bool>(LanguageID));
                            return;
                        }
                        if (ResultObj.Result)
                            lblAssociatedOffer.Text += "<a href='offer-redirect.aspx?OfferID=" + item.OfferID + "'>" + item.OfferName + "</a>" + "</br>";
                        else
                            lblAssociatedOffer.Text += item.OfferName + "</br>";
                    }
                    else
                    {
                        lblAssociatedOffer.Text += "<a href='offer-redirect.aspx?OfferID=" + item.OfferID + "'>" + item.OfferName + "</a>" + "</br>";
                    }
                }
            }
        }
    }

    private void FillExpireType()
    {
        tcpService = CurrentRequest.Resolver.Resolve<ITrackableCouponProgramService>();
        AMSResult<List<TrackableCouponExpireType>> lstExpireTypes = tcpService.GetTrackableCouponExpireTypes();
        if (lstExpireTypes.ResultType != AMSResultType.Success)
        {
            DisplayError(lstExpireTypes.GetLocalizedMessage<List<TrackableCouponExpireType>>(LanguageID));
            return;
        }
        ddlExpireTypes.DataSource = (from item in lstExpireTypes.Result
                                     select new { Text = PhraseLib.Lookup(item.PhraseID, LanguageID).Replace("&#39;", "'"), Value = item.TCExpireTypeID.ToString() }).ToList();
        ddlExpireTypes.DataTextField = "Text";
        ddlExpireTypes.DataValueField = "Value";
        ddlExpireTypes.DataBind();
    }

    private void FillExpirePeriodType()
    {
        tcpService = CurrentRequest.Resolver.Resolve<ITrackableCouponProgramService>();
        AMSResult<List<TrackableCouponExpirePeriodType>> lstExpirePeriodTypes = tcpService.GetTrackableCouponExpirePeriodTypes();
        if (lstExpirePeriodTypes.ResultType != AMSResultType.Success)
        {
            DisplayError(lstExpirePeriodTypes.GetLocalizedMessage<List<TrackableCouponExpirePeriodType>>(LanguageID));
            return;
        }
        ddlExpirePeriodTypes.DataSource = (from item in lstExpirePeriodTypes.Result
                                           select new { Text = PhraseLib.Lookup(item.PhraseID, LanguageID).Replace("&#39;", "'"), Value = item.TCExpirePeriodTypeID.ToString() }).ToList();
        ddlExpirePeriodTypes.DataTextField = "Text";
        ddlExpirePeriodTypes.DataValueField = "Value";
        ddlExpirePeriodTypes.DataBind();
    }

    private void FillExpireTime()
    {
        List<string> hours = new List<string>();
        List<string> minutes = new List<string>();

        for (int i = 0; i <= 59; i++)
        {
            if (i < 24)
            {
                hours.Add(i.ToString("00"));
            }
            minutes.Add(i.ToString("00"));
        }

        ddlExpireTimeHours.DataSource = hours;
        ddlExpireTimeHours.DataBind();
        ddlExpireTimeMinutes.DataSource = minutes;
        ddlExpireTimeMinutes.DataBind();
    }

    protected void ddlExpireTypes_SelectedIndexChanged(object sender, EventArgs e)
    {
        int selection = Convert.ToInt32(ddlExpireTypes.SelectedValue);
        switch (selection)
        {
            case 1:
            default:
                pnlExpirationPeriod.Style["Display"] = "none";
                pnlExpirationDate.Style["Display"] = "none";
                lblExpire.Visible = true;
                ExpireDate.Visible = true;
                break;
            case 2:
                pnlExpirationPeriod.Style["Display"] = "none";
                pnlExpirationDate.Style["Display"] = "block";
                lblExpire.Visible = false;
                ExpireDate.Visible = false;
                break;
            case 3:
                pnlExpirationPeriod.Style["Display"] = "block";
                pnlExpirationDate.Style["Display"] = "none";
                ddlExpirePeriodTypes.Enabled = true;
                lblExpire.Visible = false;
                ExpireDate.Visible = false;
                break;
            case 4:
                pnlExpirationPeriod.Style["Display"] = "block";
                pnlExpirationDate.Style["Display"] = "none";
                ddlExpirePeriodTypes.SelectedValue = "2";
                ddlExpirePeriodTypes.Enabled = false;
                lblExpire.Visible = false;
                ExpireDate.Visible = false;
                break;
        }
    }

    private TrackableCouponProgram GetProgramByID(Int32 pID)
    {
        tcpService = CurrentRequest.Resolver.Resolve<ITrackableCouponProgramService>();
        AMSResult<TrackableCouponProgram> tcProgram = SystemCacheData.GetSystemOption_UE_ByOptionId(152) == "0" ? tcpService.GetTrackableCouponProgramById(pID) : tcpService.GetTrackableCouponProgramById(pID, true);
        if (tcProgram.ResultType != AMSResultType.Success)
        {
            DisplayError(tcProgram.GetLocalizedMessage<TrackableCouponProgram>(LanguageID));
            return null;
        }
        else
            return (TrackableCouponProgram)tcProgram.Result;

    }

    private List<CMS.AMS.Models.Offer> GetAttachedOffersByID(Int32 pID)
    {
        offerService = CurrentRequest.Resolver.Resolve<IOffer>();
        AMSResult<List<CMS.AMS.Models.Offer>> offerResObj = offerService.GetofferByTrackableProgramID(pID);
        if (offerResObj.ResultType != AMSResultType.Success)
        {
            DisplayError(offerResObj.GetLocalizedMessage<List<CMS.AMS.Models.Offer>>(LanguageID));
            return null;
        }
        else
            return (List<CMS.AMS.Models.Offer>)offerResObj.Result;
    }

    private void DisplayError(string err)
    {
        //CustomValidator CustomValidatorCtrl = new CustomValidator();
        //CustomValidatorCtrl.IsValid = false;
        //CustomValidatorCtrl.ErrorMessage = err;
        //this.Page.Controls.Add(err);
        vsError.AddErrorMessage(err);
    }
    #endregion

}