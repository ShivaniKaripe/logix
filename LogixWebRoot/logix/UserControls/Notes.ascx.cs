using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using CMS.AMS;
using CMS.Contract;
using CMS.AMS.Contract;
using CMS.AMS.Models;

public partial class logix_UserControls_Notes : System.Web.UI.UserControl
{
  public IPhraseLib PhraseLib;
  IErrorHandler m_ErrorHandler;
  INotesService m_Notes;

  public int LinkID { get; set; }
  public NoteTypes NoteType { get; set; }

  private int AdminUserID;
  public int LanguageID;
  private bool HasVisibleNotes = false;
  private bool HasImportantNotes = false;
  private bool HasNewNotes = false;

  protected void Page_Load(object sender, EventArgs e) {
    ResolveDependencies();
    AdminUserID = ((AuthenticatedUI)this.Page).CurrentUser.AdminUser.ID;
    LanguageID = ((AuthenticatedUI)this.Page).LanguageID;
    PhraseLib = ((AuthenticatedUI)this.Page).PhraseLib;
    m_ErrorHandler = ((AuthenticatedUI)this.Page).ErrorHandler;
    if (((AuthenticatedUI)this.Page).CurrentUser.UserPermissions.AccessNotes == false) {
      notesbutton.Visible = false;
      return;
    }
    if(!IsPostBack)
      reloadNotesSrc();
  }

  public void reloadNotesSrc() {
    List<Notes> lstNotes = m_Notes.GetNotesBtnInfo(LinkID, NoteType);
    foreach (Notes note in lstNotes) {
      if (note.Private == false || (note.Private == true && note.AdminUser.ID == AdminUserID))
        HasVisibleNotes = true;
      if ((note.Private == false || (note.Private == true && note.AdminUser.ID == AdminUserID)) && note.Important)
        HasImportantNotes = true;
      if (note.CreatedDate.Subtract(DateTime.Today).Days == 0)
        HasNewNotes = true;
    }
    if (HasVisibleNotes) {
      if (HasImportantNotes) {
        notesbutton.Src = (HasNewNotes) ? "/images/notes-newimportant.png" : "/images/notes-someimportant.png";
      }
      else {
        notesbutton.Src = (HasNewNotes) ? "/images/notes-new.png" : "/images/notes-some.png";
      }
      notesbutton.Alt = (lstNotes.Count == 1) ? lstNotes.Count + " " + PhraseLib.Lookup("term.note", LanguageID).ToLower() : lstNotes.Count + " " + PhraseLib.Lookup("term.notes", LanguageID).ToLower();
      notesbutton.Attributes.Add("title", notesbutton.Alt);
    }
    else {
      notesbutton.Src = "/images/notes-none.png";
      notesbutton.Alt = PhraseLib.Lookup("term.notes", LanguageID);
    }
  }
  private void ResolveDependencies() {
    m_Notes = CurrentRequest.Resolver.Resolve<INotesService>();
  }
}