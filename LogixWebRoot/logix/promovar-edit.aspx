<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: promovar-edit.aspx 
  ' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' * Copyright © 2002 - 2009.  All rights reserved by:
  ' *
  ' * NCR Corporation
  ' * 1435 Win Hentschel Boulevard
  ' * West Lafayette, IN  47906
  ' * voice: 888-346-7199  fax: 765-464-1369
  ' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' *
  ' * PROJECT : NCR Advanced Marketing Solution
  ' *
  ' * MODULE  : Logix
  ' *
  ' * PURPOSE : 
  ' *
  ' * NOTES   : 
  ' *
  ' * Version : 7.3.1.138972 
  ' *
  ' *****************************************************************************
  
  Dim CopientFileName As String = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
  Dim CopientFileVersion As String = "7.3.1.138972" 
  Dim CopientProject As String = "Copient Logix"
  Dim CopientNotes As String = ""
  
  Dim AdminUserID As Long
  Dim row As System.Data.DataRow
    Dim MyCommon As New Copient.CommonInc
    Dim MyCryptLib As New Copient.CryptLib
  Dim Logix As New Copient.LogixInc
  Dim dstPromoVar As System.Data.DataTable
  Dim rst As DataTable
  Dim pgDescription As String
  Dim pgPromoVarID As String
  Dim pgName As String
  Dim pvID As String
  Dim PromoVarDesc As String
  Dim PromoVarName As String
  Dim ExternalID As String
  Dim PromoVarID As Integer = -1
  Dim UpdateToHost As Boolean = False
  Dim PromoVarCreated As Date = Nothing
  Dim PromoVarUpdated As Date = Nothing
  Dim longDate As New DateTime
  Dim longDateString As String
  Dim PromoVarNameTitle As String = ""
  Dim VarTypeName As String = ""
  Dim ShowActionButton As Boolean = False
  Dim statusMessage As String = ""
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "promovar-edit.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  If (Request.QueryString("infoMessage") <> "") Then
    infoMessage = Request.QueryString("infoMessage")
  End If
  
  If (Request.QueryString("new") <> "") Then
    Response.Redirect("promovar-edit.aspx")
  End If
    
  ' any GET parms inbound?
  If (Request.QueryString("Delete") <> "") Then
    ' expunge record if there is one
    If (MyCommon.Extract_Val(Request.QueryString("PromoVarID")) > 0) Then
      MyCommon.QueryStr = "pt_Promotion_Variables_Delete"
      MyCommon.Open_LXSsp()
      MyCommon.LXSsp.Parameters.Add("@PromoVarID", SqlDbType.BigInt).Value = MyCommon.Extract_Val(Request.QueryString("PromoVarID"))
      MyCommon.LXSsp.ExecuteNonQuery()
      MyCommon.Close_LXSsp()
    End If
    'Record history
    MyCommon.Activity_Log(27, PromoVarID, AdminUserID, Copient.PhraseLib.Lookup("history.promovar-delete", LanguageID))
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "promovar-list.aspx")
    GoTo done
  ElseIf (Request.QueryString("PromoVarID") = Copient.PhraseLib.Lookup("term.new", LanguageID)) Then
    ' add a record
    PromoVarName = MyCommon.Parse_Quotes(Request.QueryString.Item("name"))
    PromoVarName = Logix.TrimAll(PromoVarName)
    PromoVarDesc = MyCommon.Parse_Quotes(Request.QueryString.Item("desc"))
    ExternalID = MyCommon.Parse_Quotes(Request.QueryString.Item("xid"))
    
    MyCommon.QueryStr = "dbo.pt_Promotion_Variables_Insert"
    MyCommon.Open_LXSsp()
    MyCommon.LXSsp.Parameters.Add("@ExternalID", SqlDbType.NVarChar, 400).Value = MyCryptLib.SQL_StringEncrypt(Request.QueryString.Item("xid"))
    MyCommon.LXSsp.Parameters.Add("@Name", SqlDbType.NVarChar, 50).Value = Logix.TrimAll(Request.QueryString.Item("name"))
    MyCommon.LXSsp.Parameters.Add("@Description", SqlDbType.NVarChar, 50).Value = Request.QueryString.Item("desc")
    MyCommon.LXSsp.Parameters.Add("@VarTypeID", SqlDbType.Int).Value = MyCommon.Extract_Val(Request.QueryString.Item("vartype"))
    MyCommon.LXSsp.Parameters.Add("@UpdateToHost", SqlDbType.Bit).Value = MyCommon.Extract_Val(Request.QueryString.Item("updatehost"))
    MyCommon.LXSsp.Parameters.Add("@PromoVarID", SqlDbType.BigInt).Direction = ParameterDirection.Output

    If PromoVarName = "" Then
      infoMessage = Copient.PhraseLib.Lookup("promovar.noname", LanguageID)
    Else
      MyCommon.QueryStr = "SELECT PromoVarID FROM PromoVariables with (NoLock) WHERE ExternalID = '" & MyCryptLib.SQL_StringEncrypt(PromoVarName) & "' AND Deleted=0"
      rst = MyCommon.LXS_Select
      If (rst.Rows.Count > 0) Then
        infoMessage = Copient.PhraseLib.Lookup("promovar.edit-nameused", LanguageID)
      Else
        MyCommon.LXSsp.ExecuteNonQuery()
        PromoVarID = MyCommon.LXSsp.Parameters("@PromoVarID").Value
        MyCommon.Activity_Log(27, PromoVarID, AdminUserID, Copient.PhraseLib.Lookup("history.promovar-create", LanguageID))
      End If
    End If
    MyCommon.Close_LXSsp()

    If (PromoVarName <> "") Then
      Response.Status = "301 Moved Permanently"
      If (PromoVarID = 0) Then
        Response.AddHeader("Location", "promovar-edit.aspx?PromoVarID=" & PromoVarID & "&infoMessage=" & Copient.PhraseLib.Lookup("point-edit.nameused", LanguageID))
      Else
        Response.AddHeader("Location", "promovar-edit.aspx?PromoVarID=" & PromoVarID)
      End If
      GoTo done
    End If
  ElseIf (Request.QueryString("save") <> "") Then
    ' somebody clicked save
    PromoVarID = MyCommon.Extract_Val(Request.QueryString("PromoVarID"))
    With Request.QueryString
      PromoVarName = MyCommon.Parse_Quotes(.Item("name"))
      If PromoVarName = "" Then
        infoMessage = Copient.PhraseLib.Lookup("promovar.noname", LanguageID)
      Else
                MyCommon.QueryStr = "SELECT PromoVarID FROM PromoVariables with (NoLock) WHERE ExternalID = '" & MyCryptLib.SQL_StringEncrypt(PromoVarName) & "' AND Deleted=0"
        rst = MyCommon.LXS_Select
        If (rst.Rows.Count > 0) Then
          infoMessage = Copient.PhraseLib.Lookup("promovar.edit-nameused", LanguageID)
        Else
          MyCommon.QueryStr = "dbo.pt_Promotion_Variables_Update"
          MyCommon.Open_LXSsp()
          MyCommon.LXSsp.Parameters.Add("@PromoVarID", SqlDbType.BigInt).Value = PromoVarID
          MyCommon.LXSsp.Parameters.Add("@ExternalID", SqlDbType.NVarChar, 400).Value = MyCryptLib.SQL_StringEncrypt(Request.QueryString.Item("xid"))
          MyCommon.LXSsp.Parameters.Add("@Name", SqlDbType.NVarChar, 50).Value = Request.QueryString.Item("name")
          MyCommon.LXSsp.Parameters.Add("@Description", SqlDbType.NVarChar, 50).Value = Request.QueryString.Item("desc")
          MyCommon.LXSsp.Parameters.Add("@VarTypeID", SqlDbType.Int).Value = MyCommon.Extract_Val(Request.QueryString.Item("vartype"))
          MyCommon.LXSsp.Parameters.Add("@UpdateToHost", SqlDbType.Bit).Value = MyCommon.Extract_Val(Request.QueryString.Item("updatehost"))
          MyCommon.LXSsp.ExecuteNonQuery()
          
          MyCommon.Activity_Log(27, PromoVarID, AdminUserID, Copient.PhraseLib.Lookup("history.promovar-edit", LanguageID))
        End If
      End If
    End With
  ElseIf (Request.QueryString("PromoVarID") <> "") Then
    ' simple edit/search mode
    PromoVarID = MyCommon.NZ(Request.QueryString("PromoVarID"), "0")
  ElseIf (Request.Form("PromoVarID") <> "") Then
    PromoVarID = MyCommon.Extract_Val(Request.Form("PromoVarID"))
  Else
    ' no group id passed ... what now ?
    PromoVarID = "0"
  End If
  
  ' grab this promo var
  MyCommon.QueryStr = "select PromoVarID, VarTypeID, Description, ExternalID, Name, UpdateToHost, CreatedDate, LastUpdate " & _
                      "from PromoVariables with (NoLock) where Deleted=0 AND PromoVarID=" & PromoVarID & ";"
  dstPromoVar = MyCommon.LXS_Select
  If (dstPromoVar.Rows.Count > 0) Then
    PromoVarID = MyCommon.NZ(dstPromoVar.Rows(0).Item("PromoVarID"), 0)
    PromoVarDesc = MyCommon.NZ(dstPromoVar.Rows(0).Item("Description"), "")
    ExternalID = MyCryptLib.SQL_StringDecrypt(dstPromoVar.Rows(0).Item("ExternalID").ToString())
    PromoVarName = MyCommon.NZ(dstPromoVar.Rows(0).Item("Name"), "")
    UpdateToHost = MyCommon.NZ(dstPromoVar.Rows(0).Item("UpdateToHost"), False)
    PromoVarCreated = MyCommon.NZ(dstPromoVar.Rows(0).Item("CreatedDate"), Nothing)
    PromoVarUpdated = MyCommon.NZ(dstPromoVar.Rows(0).Item("LastUpdate"), Nothing)
  ElseIf (Request.QueryString("new") = "") And (PromoVarID > 0) Then
    Send_HeadBegin("term.pointsprogram", , PromoVarID)
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts()
    Send_HeadEnd()
    Send_BodyBegin(1)
    Send_Bar(Handheld)
    Send_Help(CopientFileName)
    Send_Logos()
    Send_Tabs(Logix, 5)
    Send_Subtabs(Logix, 53, 4, , PromoVarID)
    Send("")
    Send("<div id=""intro"">")
    Send("    <h1 id=""title"">" & Copient.PhraseLib.Lookup("term.promovar", LanguageID) & " #" & PromoVarID & "</h1>")
    Send("</div>")
    Send("<div id=""main"">")
    Send("    <div id=""infobar"" class=""red-background"">")
    Send("        " & Copient.PhraseLib.Lookup("error.deleted", LanguageID))
    Send("    </div>")
    Send("</div>")
    Send_BodyEnd()
    GoTo done
  Else
    PromoVarID = 0
    PromoVarDesc = ""
    PromoVarName = Copient.PhraseLib.Lookup("term.newpromovar", LanguageID)
  End If
  
  Send_HeadBegin("term.promovar", , PromoVarID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
  Send_HeadEnd()
  Send_BodyBegin(1)
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  Send_Tabs(Logix, 5)
  Send_Subtabs(Logix, 53, 4, , PromoVarID)
  
  If (Logix.UserRoles.AccessPointsPrograms = False) Then
    Send_Denied(1, "perm.points-access") ' TO DO: change after adding permissions
    Send_BodyEnd()
    GoTo done
  End If
%>
<script type="text/javascript">
    function toggleDropdown() {
        if (document.getElementById("actionsmenu") != null) {
            bOpen = (document.getElementById("actionsmenu").style.visibility != 'visible')
            if (bOpen) {
                document.getElementById("actionsmenu").style.visibility = 'visible';
                document.mainform.actions.value = '<% Sendb(Copient.PhraseLib.Lookup("term.actions", LanguageID)) %>▲';
            } else {
                document.getElementById("actionsmenu").style.visibility = 'hidden';
                document.mainform.actions.value = '<% Sendb(Copient.PhraseLib.Lookup("term.actions", LanguageID)) %>▼';
            }
        }
   }
</script>
<form action="#" method="get" id="mainform" name="mainform" onsubmit="return disableSaveCheck();">
  <div id="intro">
    <h1 id="title">
      <%
        If PromoVarID = 0 Then
          Sendb(Copient.PhraseLib.Lookup("term.newpromovar", LanguageID))
        Else
          Sendb(Copient.PhraseLib.Lookup("term.promovar", LanguageID) & " #" & PromoVarID & ": ")
          MyCommon.QueryStr = "SELECT PromoVarID,Name FROM PromoVariables with (NoLock) WHERE deleted=0 and PromoVarID=" & PromoVarID & ";"
          rst = MyCommon.LXS_Select
          If (rst.Rows.Count > 0) Then
            PromoVarNameTitle = MyCommon.NZ(rst.Rows(0).Item("Name"), "")
          End If
          Sendb(MyCommon.TruncateString(PromoVarNameTitle, 40))
        End If
      %>
    </h1>
    <div id="controls">
      <%
        If (pvID = 0) Then
          If (Logix.UserRoles.CreatePointsPrograms) Then
            Send_Save()
          End If
        Else
          ShowActionButton = (Logix.UserRoles.EditPointsPrograms) OrElse (Logix.UserRoles.DeletePointsPrograms) OrElse (Logix.UserRoles.EditPointsPrograms)
          If (ShowActionButton) Then
            Send("<input type=""button"" class=""regular"" id=""actions"" name=""actions"" value=""" & Copient.PhraseLib.Lookup("term.actions", LanguageID) & " ▼"" onclick=""toggleDropdown();"" />")
            Send("<div class=""actionsmenu"" id=""actionsmenu"">")
            If (Logix.UserRoles.EditPointsPrograms) Then
              Send_Save()
            End If
            If (Logix.UserRoles.DeletePointsPrograms) Then
              Send_Delete()
            End If
            If (Logix.UserRoles.CreatePointsPrograms) Then
              Send_New()
            End If
            Send("</div>")
          End If
          If MyCommon.Fetch_SystemOption(75) Then
            If (Logix.UserRoles.AccessNotes) Then
              Send_NotesButton(10, PromoVarID, AdminUserID)
            End If
          End If
        End If
      %>
    </div>
  </div>
  <div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <% If (statusMessage <> "") Then Send("<div id=""statusbar"" class=""green-background"">" & statusMessage & "</div>")%>
    <div id="column1">
      <div class="box" id="identity">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.identification", LanguageID))%>
          </span>
        </h2>
        <input type="hidden" id="ProgramGroupID" name="ProgramGroupID" value="<% sendb(PromoVarID) %>" />
        <input type="hidden" id="PromoVarID" name="PromoVarID" value="<% sendb(pgPromoVarID) %>" />
        <label for="id"><% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>:</label><br />
        <input type="text" class="shorter" id="id" name="id" maxlength="4" value="<% Sendb(pvID) %>" /><br />
        <label for="name"><% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>:</label><br />
        <input type="text" class="longest" id="name" name="name" maxlength="50" value="<% Sendb(pgName) %>" /><br />
        <label for="xid"><% Sendb(Copient.PhraseLib.Lookup("term.externalid", LanguageID))%>:</label><br />
        <input type="text" class="longest" id="xid" name="xid" maxlength="50" value="<% Sendb(pgName) %>" /><br />
        <label for="desc"><% Sendb(Copient.PhraseLib.Lookup("term.description", LanguageID))%>:</label><br />
        <textarea class="longest" id="desc" name="desc" cols="48" rows="3"><% Sendb(pgDescription)%></textarea><br />
        <label for="vartype"><% Sendb(Copient.PhraseLib.Lookup("term.type", LanguageID))%>:</label><br />
        <select class="longest" name="vartype" id="vartype">
        <%
          MyCommon.QueryStr = "select distinct TypeID, Description, PhraseID from PromoVarTypes with (NoLock) order by Description;"
          rst = MyCommon.LXS_Select
          For Each row In rst.Rows
            If (MyCommon.NZ(row.Item("PhraseID"), 0) = 0) Then
              VarTypeName = MyCommon.NZ(row.Item("Description"), "")
            Else
              VarTypeName = Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID)
            End If
            Send("<option value=""" & MyCommon.NZ(row.Item("TypeID"), 0) & """ >" & VarTypeName & "</option>")
          Next
        %>
        </select><br /><br class="half" />
        <input type="checkbox" id="updatehost" name="updatehost" value="1" />
        <label for="updatehost"><% Sendb(Copient.PhraseLib.Lookup("term.updatetohost", LanguageID))%></label><br />
        <br class="half" />
        <%
          If PromoVarCreated = Nothing Then
          Else
            Sendb(Copient.PhraseLib.Lookup("term.created", LanguageID) & " ")
            longDate = PromoVarCreated
            longDateString = longDate.ToString("dddd, d MMMM yyyy, HH:mm:ss")
            Send(longDateString)
            Send("<br />")
          End If
          If PromoVarUpdated = Nothing Then
          Else
            Sendb(Copient.PhraseLib.Lookup("term.edited", LanguageID) & " ")
            longDate = PromoVarUpdated
            longDateString = longDate.ToString("dddd, d MMMM yyyy, HH:mm:ss")
            Send(longDateString)
          End If
        %>
        <br class="half" />
        <hr class="hidden" />
      </div>
    </div>

    <div id="gutter">
    </div>

    <div id="column2">
    </div>
  </div>
</form>
<script type="text/javascript">
    if (window.captureEvents){
        window.captureEvents(Event.CLICK);
        window.onclick=handlePageClick;
    }
    else {
        document.onclick=handlePageClick;
    }
</script>
<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (PromoVarID > 0 And Logix.UserRoles.AccessNotes) Then
      Send_Notes(10, PromoVarID, AdminUserID)
    End If
  End If
  Send_BodyEnd("mainform", "name")
done:
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  Logix = Nothing
  MyCommon = Nothing
%>
