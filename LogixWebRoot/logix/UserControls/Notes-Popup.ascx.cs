using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using CMS;
using CMS.Contract;
using CMS.AMS;
using CMS.AMS.Contract;
using CMS.AMS.Models;


public partial class logix_UserControls_Notes_Popup : System.Web.UI.UserControl
{
  public IPhraseLib PhraseLib;
  public Permisssions UserPermissions;
  IErrorHandler m_ErrorHandler;
  INotesService m_Notes;
  IActivityLogService m_ActivityLog;
  Copient.CommonInc MyCommon = new Copient.CommonInc();

  public int LinkID { get; set; }
  public NoteTypes NoteType { get; set; }
  public ActivityTypes ActivityType { get; set; }

  private string HistoryString = string.Empty;
  protected int AdminUserID;
  public int LanguageID;
  private int ActivityTypePhraseID = 0;

  public event EventHandler NotesUpdate;
  
  protected void Page_Load(object sender, EventArgs e) {
    ResolveDependencies();
    AdminUserID = ((AuthenticatedUI)this.Page).CurrentUser.AdminUser.ID;
    LanguageID = ((AuthenticatedUI)this.Page).LanguageID;
    PhraseLib = ((AuthenticatedUI)this.Page).PhraseLib;
    m_ErrorHandler = ((AuthenticatedUI)this.Page).ErrorHandler;
    notesave.Text = PhraseLib.Lookup("term.save", LanguageID);
    UserPermissions = ((AuthenticatedUI)this.Page).CurrentUser.UserPermissions;
    if (UserPermissions.CreateNotes == false)
      noteadddiv.Visible = false;
    if (UserPermissions.AccessNotes && !IsPostBack)
      LoadNotes();
  }

  protected void notesave_Click(object sender, EventArgs e) {
    Notes objNote = new Notes();    
    if (notetext.InnerText.Trim() != String.Empty) {
      objNote.AdminUser = new CMS.Models.AdminUser();
      objNote.AdminUser.FirstName = ((AuthenticatedUI)this.Page).CurrentUser.AdminUser.FirstName;
      objNote.AdminUser.LastName = ((AuthenticatedUI)this.Page).CurrentUser.AdminUser.LastName;
      objNote.Note = notetext.InnerText;
      objNote.NoteTypeID = NoteType;
      objNote.LinkID = LinkID;
      objNote.Private = cbPrivate.Checked;
      objNote.Important = cbImportant.Checked;
      objNote.AdminUser.ID = AdminUserID;
      objNote.AdminUser.LanguageID = LanguageID;
      if (m_Notes.AddNote(objNote) == true) {
        if (objNote.Private == false) {
          HistoryString = PhraseLib.Lookup("history.note-add", LanguageID);
          ActivityTypePhraseID = m_ActivityLog.GetActivityTypePhraseID(ActivityType);
          if (ActivityTypePhraseID > 0)
            HistoryString += " " + PhraseLib.Lookup("term.to", LanguageID).ToLower() + " " + PhraseLib.Lookup(ActivityTypePhraseID, LanguageID).ToLower();
          if (objNote.LinkID == 0) {
            switch (objNote.NoteTypeID) {
              case NoteTypes.Offers:
              case NoteTypes.CustomerGroup:
              case NoteTypes.ProductGroup:
              case NoteTypes.PointsProgram:
              case NoteTypes.StoredValueProgram:
              case NoteTypes.Promovar:
              case NoteTypes.Graphic:
              case NoteTypes.Layout:
              case NoteTypes.Store:
              case NoteTypes.StoreGroup:
              case NoteTypes.Agent:
              case NoteTypes.Report:
              case NoteTypes.User:
              case NoteTypes.Banner:
              case NoteTypes.Department:
              case NoteTypes.Terminal:
                HistoryString += " " + PhraseLib.Lookup("term.list", LanguageID).ToLower();
                break;
            }
          }
          m_ActivityLog.Activity_Log(ActivityType, LinkID, AdminUserID, HistoryString);
        }
        notetext.InnerText = String.Empty;
        cbPrivate.Checked = false;
        cbImportant.Checked = false;
        if(NotesUpdate!=null)
          NotesUpdate(sender, e);
        LoadNotes();
        ScriptManager.RegisterStartupScript(this,this.GetType(), "Script", "toggleNotesDisplay()", true);
      }
    }
  }

  protected void rptNotes_ItemCreated(object sender, RepeaterItemEventArgs e) {
    Notes note = (Notes)e.Item.DataItem;
    if (note != null) {
      ((HiddenField)e.Item.FindControl("IsPrivate")).Value = note.Private.ToString();
    }
  }

  protected void rptNotes_ItemCommand(object source, RepeaterCommandEventArgs e) {
    if (e.CommandName.ToString() == "Delete") {
      long NoteID = e.CommandArgument.ConvertToLong();
      bool IsPrivate = ((HiddenField)e.Item.FindControl("IsPrivate")).Value.ConvertToBoolean();
      m_Notes.DeleteNote(NoteID, NoteType);
      if (IsPrivate == false) {
        HistoryString = PhraseLib.Lookup("history.note-delete", LanguageID);
        ActivityTypePhraseID = m_ActivityLog.GetActivityTypePhraseID(ActivityType);
        if (ActivityTypePhraseID > 0)
          HistoryString += " " + PhraseLib.Lookup("term.from", LanguageID).ToLower() + " " + PhraseLib.Lookup(ActivityTypePhraseID, LanguageID).ToLower();
        if (LinkID == 0) {
          switch (NoteType) {
            case NoteTypes.Offers:
            case NoteTypes.CustomerGroup:
            case NoteTypes.ProductGroup:
            case NoteTypes.PointsProgram:
            case NoteTypes.StoredValueProgram:
            case NoteTypes.Promovar:
            case NoteTypes.Graphic:
            case NoteTypes.Layout:
            case NoteTypes.Store:
            case NoteTypes.StoreGroup:
            case NoteTypes.Agent:
            case NoteTypes.Report:
            case NoteTypes.User:
            case NoteTypes.Banner:
            case NoteTypes.Department:
            case NoteTypes.Terminal:
              HistoryString += " " + PhraseLib.Lookup("term.list", LanguageID).ToLower();
              break;
          }
        }
        m_ActivityLog.Activity_Log(ActivityType, LinkID, AdminUserID, HistoryString);
      }
      if (NotesUpdate != null)
        NotesUpdate(source, e);
      LoadNotes();
      ScriptManager.RegisterStartupScript(this, this.GetType(), "Script", "toggleNotesDisplay()", true);
    }
  }

  private void LoadNotes() {
    List<Notes> lstNotes = m_Notes.GetNotes(LinkID, NoteType);
    rptNotes.DataSource = lstNotes;
    rptNotes.DataBind();
    if (lstNotes.Count > 0)
      emptyNotes.Visible = false;
  }
  private void ResolveDependencies() {
    m_Notes = CurrentRequest.Resolver.Resolve<INotesService>();
    m_ActivityLog = CurrentRequest.Resolver.Resolve<IActivityLogService>();
  }

}