using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using CMS;
using CMS.AMS;
using CMS.AMS.Contract;
using System.Data;
using CMS.AMS.Models;

public partial class logix_tcprogram_hist : AuthenticatedUI
{
  int TCProgramID = 0;

  IActivityLogService m_ActivityLog;
  ITrackableCouponProgramService m_TCProgramService;

  protected override void OnInit(EventArgs e)
  {
    base.OnInit(e);
    ((logix_LogixMasterPage)this.Master).Tab_Name = "5_3_3";
  }


  protected void Page_Load(object sender, EventArgs e)
  {
    try
    {
      Response.Expires = -1;
      infobar.Visible = false;
      ResolveDepedencies();
      GetQueryStrings();
      SetUpUserControls();
      if (!IsPostBack)
      {
        SetUpAndLocalizePage();
        AssignPageTitle("term.trackablecouponprogram", "term.history", TCProgramID.ToString());
        LoadHistoryData(TCProgramID);

      }
      ucNotes_Popup.NotesUpdate += new EventHandler(ucNotes_Popup_NotesUpdate);
    }
    catch (Exception ex)
    {
      infobar.InnerText = ErrorHandler.ProcessError(ex);
      infobar.Visible = true;
    }
  }

  private bool IsTrackableCouponExists()
  {
    bool IsExist = false;
    if (TCProgramID <= 0)
      return false;
    AMSResult<TrackableCouponProgram> tcProgram = m_TCProgramService.GetTrackableCouponProgramById(TCProgramID);
    if (tcProgram.ResultType == AMSResultType.Success)
    {
      IsExist = (tcProgram.Result != null);
    }
    return IsExist;
  }

  protected void ucNotes_Popup_NotesUpdate(object sender, EventArgs e)
  {
    try
    {
      ucNotesUI.reloadNotesSrc();
      LoadHistoryData(TCProgramID);
    }
    catch (Exception ex)
    {
      infobar.InnerText = ErrorHandler.ProcessError(ex);
      infobar.Visible = true;
    }
  }

  #region Override Methods
  protected override void AuthorisePage()
  {
    if (CurrentUser.UserPermissions.AccessTrackableCouponPrograms == false)
    {
      Server.Transfer("PageDenied.aspx?PhraseName=perm.trackablecoupon-access&TabName=5_3_3", false);
      return;
    }
  }
  #endregion

  private void SetUpAndLocalizePage()
  {
    title.InnerText = "Trackable Coupon Program #" + TCProgramID;
  }
  private void SetUpUserControls()
  {
    Copient.CommonInc MyCommon = new Copient.CommonInc();
    ucNotesUI.Visible = MyCommon.Fetch_SystemOption(75).Equals("1") ? true : false;
    ucNotesUI.NoteType = NoteTypes.TCProgram;
    ucNotesUI.LinkID = TCProgramID;
    ucNotes_Popup.NoteType = NoteTypes.TCProgram;
    ucNotes_Popup.LinkID = TCProgramID;
    ucNotes_Popup.ActivityType = ActivityTypes.TCProgram;
  }

  private void LoadHistoryData(int TCProgramID)
  {
    if (TCProgramID > 0)
    {
      rptProgramHistory.DataSource = m_ActivityLog.GetActivityTypeHistory(TCProgramID, ActivityTypes.TCProgram);
      rptProgramHistory.DataBind();
    }
  }
  private void GetQueryStrings()
  {
    string paramVal = Request.QueryString["tcprogramid"];
    if (!String.IsNullOrEmpty(paramVal))
    {
      if (Int32.TryParse(paramVal, out TCProgramID))
      {
        if (!IsTrackableCouponExists())
          Server.Transfer("error-message.aspx?MainHeading=" + PhraseLib.Lookup("term.trackablecouponprogram", LanguageID) + " #" + TCProgramID + "&ErrorMessage=" + PhraseLib.Lookup("term.itemnotfound", LanguageID) + "&TabName=5_3_3", false);
      }
      else
        Response.Redirect("~/logix/tcp-edit.aspx", false);
    }
    else
      Response.Redirect("~/logix/tcp-edit.aspx", false);
  }

  private void ResolveDepedencies()
  {
    m_ActivityLog = CurrentRequest.Resolver.Resolve<IActivityLogService>();
    m_TCProgramService = CurrentRequest.Resolver.Resolve<ITrackableCouponProgramService>();
  }
}