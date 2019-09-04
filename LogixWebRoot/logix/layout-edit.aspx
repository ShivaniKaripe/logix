<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: layout-edit.aspx 
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
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim dt As DataTable
  Dim row As DataRow
  Dim LayoutID As Long = 0
  Dim bExistingLayout As Boolean = True
  Dim LayoutName As String = ""
  Dim Width As Integer = 640
  Dim Height As Integer = 480
  Dim NumRecs As Integer = 0
  Dim CellID As Integer = 0
  Dim DisableCellEdit As String = ""
  Dim ShowActionButton As Boolean = False
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim OfferCtr As Integer = 0
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  
  Response.Expires = 0
  MyCommon.AppName = "layout-edit.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  LayoutID = MyCommon.Extract_Decimal(Request.QueryString("LayoutID"), MyCommon.GetAdminUser.Culture)
  
  If (Request.QueryString("new") <> "") Then
    Response.Redirect("layout-edit.aspx")
  End If
  
  If (Request.QueryString("save") <> "" OrElse Request.QueryString("unsavedChanges") = "true") Then
    LayoutName = Left(Logix.TrimAll(MyCommon.Strip_Quotes(Request.QueryString("name"))), 255)
    Width = MyCommon.Extract_Decimal(Request.QueryString("layoutwidth"), MyCommon.GetAdminUser.Culture)
    Height = MyCommon.Extract_Decimal(Request.QueryString("layoutheight"), MyCommon.GetAdminUser.Culture)
    If (LayoutName = "") Then
      infoMessage = Copient.PhraseLib.Lookup("layout.noname", LanguageID)
    Else
      'create new layout
      MyCommon.QueryStr = "select count(*) as NumRecs from ScreenLayouts with (NoLock) where Name='" & MyCommon.Parse_Quotes(LayoutName) & "' and Deleted=0 and LayoutID<>" & LayoutID & ";"
      dt = MyCommon.LRT_Select()
      If (dt.Rows.Count > 0) Then
        NumRecs = MyCommon.NZ(dt.Rows(0).Item("NumRecs"), 0)
      End If
      If (NumRecs > 0) Then
        infoMessage = Copient.PhraseLib.Lookup("layout.duplicatename", LanguageID)
      Else
        If (LayoutID = 0) Then
          MyCommon.QueryStr = "insert into ScreenLayouts with (RowLock) (Name, Width, Height, Deleted, LastUpdate) values (N'" & LayoutName & "', " & Width & ", " & Height & ", 0, getdate());"
          MyCommon.LRT_Execute()
          MyCommon.QueryStr = "select @@Identity as PK from ScreenLayouts with (NoLock);"
          dt = MyCommon.LRT_Select
          If Not (dt.Rows.Count = 0) Then
            LayoutID = MyCommon.NZ(dt.Rows(0).Item("PK"), 0)
          End If
          MyCommon.Activity_Log(13, LayoutID, AdminUserID, Copient.PhraseLib.Lookup("history.layout-create", LanguageID))
        Else
          MyCommon.QueryStr = "update ScreenLayouts with (RowLock) set Name=N'" & MyCommon.Parse_Quotes(LayoutName) & "', Width=" & Width & ", Height=" & Height & ", LastUpdate=getdate() where LayoutID=" & LayoutID & ";"
          MyCommon.LRT_Execute()
          MyCommon.Activity_Log(13, LayoutID, AdminUserID, Copient.PhraseLib.Lookup("history.layout-edit", LanguageID))
        End If
      End If
    End If
  End If
  
  If (Request.QueryString("new") <> "") Then
    bExistingLayout = False
  ElseIf (Request.QueryString("delete") <> "") Then
    
    ' check that there are no deployed offers that use this layout
    MyCommon.QueryStr = "dbo.pa_AssociatedOffers_ST"
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add("@LinkType", SqlDbType.Int).Value = 7
    MyCommon.LRTsp.Parameters.Add("@LinkID", SqlDbType.Int).Value = LayoutID
    dt = MyCommon.LRTsp_select
    MyCommon.Close_LRTsp()
      
    If (dt.Rows.Count > 0) Then
      infoMessage = Copient.PhraseLib.Lookup("term.inusedeployment", LanguageID) & " : ("
      For OfferCtr = 0 To dt.Rows.Count - 1
        infoMessage &= MyCommon.NZ(dt.Rows(OfferCtr).Item("IncentiveID"), "")
      Next
      infoMessage &= ")"
    Else
      MyCommon.QueryStr = "update ScreenCells with (RowLock) set Deleted=1, LastUpdate=getdate() where LayoutID=" & LayoutID & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "update ScreenLayouts with (RowLock) set Deleted=1, LastUpdate=getdate() where LayoutID=" & LayoutID & ";"
      MyCommon.LRT_Execute()
      MyCommon.Activity_Log(13, LayoutID, AdminUserID, Copient.PhraseLib.Lookup("history.layout-delete", LanguageID))
      Response.Status = "301 Moved Permanently"
      Response.AddHeader("Location", "layout-list.aspx?infoMessage=" & Copient.PhraseLib.Lookup("history.layout-delete", LanguageID) & " " & LayoutID)
      GoTo done
    End If
  ElseIf (Request.QueryString("edit") <> "") Then
    CellID = MyCommon.Extract_Decimal(Request.QueryString("CellID"), MyCommon.GetAdminUser.Culture)
    If (CellID > 0) Then
      Response.Status = "301 Moved Permanently"
      Response.AddHeader("Location", "layout-cell.aspx?CellID=" & CellID & "&LayoutID=" & LayoutID)
      GoTo done
    Else
      infoMessage = Copient.PhraseLib.Lookup("layout.nocellselected", LanguageID)
    End If
  ElseIf (Request.QueryString("remove") <> "" And Request.QueryString("CellID") <> "") Then
    CellID = MyCommon.Extract_Decimal(Request.QueryString("CellID"), MyCommon.GetAdminUser.Culture)
    MyCommon.QueryStr = "update ScreenCells with (RowLock) set Deleted=1, LastUpdate=getdate() where CellID=" & CellID & " and Deleted=0;"
    MyCommon.LRT_Execute()
    MyCommon.Activity_Log(13, LayoutID, AdminUserID, Copient.PhraseLib.Lookup("history.layout-cell-delete", LanguageID))
  ElseIf (Request.QueryString("newcell") <> "") Then
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "layout-cell.aspx?CellID=0" & "&LayoutID=" & LayoutID)
    GoTo done
  End If
  
  If (LayoutID > 0) Then
    ' display the existing layout
    MyCommon.QueryStr = "select Name, Width, Height, Deleted from ScreenLayouts with (RowLock) where LayoutID=" & LayoutID & ";"
    dt = MyCommon.LRT_Select()
    If (dt.Rows.Count > 0) Then
      LayoutName = MyCommon.NZ(dt.Rows(0).Item("Name"), "")
      Width = MyCommon.NZ(dt.Rows(0).Item("Width"), 0)
      Height = MyCommon.NZ(dt.Rows(0).Item("Height"), 0)
      If dt.Rows(0).Item("Deleted") Then
        Send_HeadBegin("term.layout", , LayoutID)
        Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
        Send_Metas()
        Send_Links(Handheld)
        Send_Scripts()
        Send_HeadEnd()
        Send_BodyBegin(1)
        Send_Bar(Handheld)
        Send_Help(CopientFileName)
        Send_Logos()
        Send_Tabs(Logix, 6)
        Send_Subtabs(Logix, 62, 3, , LayoutID)
        Send("")
        Send("<div id=""intro"">")
        Send("  <h1 id=""title"">" & Copient.PhraseLib.Lookup("term.layout", LanguageID) & " #" & LayoutID & IIf(LayoutName <> "", ": " & LayoutName, "") & "</h1>")
        Send("</div>")
        Send("<div id=""main"">")
        Send("  <div id=""infobar"" class=""red-background"">")
        Send("    " & Copient.PhraseLib.Lookup("error.deleted", LanguageID))
        Send("  </div>")
        Send("</div>")
        Send_BodyEnd()
        GoTo done
      End If
    End If
  Else
    bExistingLayout = False
  End If
  
  Send_HeadBegin("term.layout", , LayoutID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
  Send_HeadEnd()
  Send_BodyBegin(1)
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  Send_Tabs(Logix, 6)
  Send_Subtabs(Logix, 62, 3, , LayoutID)
  
  If (Logix.UserRoles.AccessLayouts = False) Then
    Send_Denied(1, "perm.layouts-access")
    Send_BodyEnd()
    GoTo done
  End If
%>

<script type="text/javascript" language="javascript">
function handleSubmit() {
    var changed = checkForChanges();
    return validateEntry();
}

function validateEntry() {
    var elemName = document.getElementById("name");
    if (elemName != null && elemName.value == "") {
        alert('<% Sendb(Copient.PhraseLib.Lookup("layout.noname", LanguageID))%>')
        elemName.focus();
        return false;
    }
    return true;
}

function checkForChanges() {
    var elemChanged = document.getElementById("unsavedChanges")
    var elemName = document.getElementById("name");
    var elemWidth = document.getElementById("layoutWidth");
    var elemHeight = document.getElementById("layoutHeight");
    
    if (elemChanged != null && elemName != null && elemName.value != elemName.defaultValue) {
        elemChanged.value = "true";
        return true;
    }
    if (elemChanged != null && elemWidth != null && elemWidth.value != elemWidth.defaultValue) {
        elemChanged.value = "true";
        return true;
    }
    if (elemChanged != null && elemHeight != null && elemHeight.value != elemHeight.defaultValue) {
        elemChanged.value = "true";
        return true;
    }
}

function openPreviewWin() {
  var url = "layout-preview.aspx?LayoutID=<% Sendb(LayoutID) %>"
  var popW = <% Sendb(Width) %>;
	var popH = <% Sendb(Height) %>;
  previewWindow = window.open(url,"Popup", "width="+popW+", height="+popH+", top="+calcTop(popH)+", left="+calcLeft(popW)+", toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=no, resizable=yes");
  previewWindow.focus();
}

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

<form action="#" id="mainform" name="mainform" method="get" onsubmit="return handleSubmit();">
  <input type="hidden" id="unsavedChanges" name="unsavedChanges" value="false" />
  <div id="intro">
    <h1 id="title">
      <% 
        If LayoutID = 0 Then
          Sendb(Copient.PhraseLib.Lookup("term.newlayout", LanguageID))
        Else
          Sendb(Copient.PhraseLib.Lookup("term.layout", LanguageID) & " #" & LayoutID & ": ")
          If (Len(LayoutName) <= 40) Then
            Sendb(LayoutName)
          Else
            Dim LayoutNameShort As String
            LayoutNameShort = Left(LayoutName, 40)
            Sendb(LayoutNameShort & "...")
          End If
        End If
      %>
    </h1>
    <div id="controls">
      <%
        If (LayoutID = 0) Then
          If (Logix.UserRoles.CreateLayouts) Then
            Send_Save()
          End If
        Else
          ShowActionButton = (Logix.UserRoles.CreateLayouts) OrElse (Logix.UserRoles.EditLayouts) OrElse (Logix.UserRoles.DeleteLayouts)
          If (ShowActionButton) Then
            Send("<input type=""button"" class=""regular"" id=""actions"" name=""actions"" value=""" & Copient.PhraseLib.Lookup("term.actions", LanguageID) & " ▼"" onclick=""toggleDropdown();"" />")
            Send("<div class=""actionsmenu"" id=""actionsmenu"">")
            If (Logix.UserRoles.EditLayouts) Then
              Send_Save()
            End If
            If (Logix.UserRoles.DeleteLayouts) Then
              Send_Delete()
            End If
            If (Logix.UserRoles.CreateLayouts) Then
              Send_New()
            End If
            Send("</div>")
          End If
          If MyCommon.Fetch_SystemOption(75) Then
            If (Logix.UserRoles.AccessNotes) Then
              Send_NotesButton(12, LayoutID, AdminUserID)
            End If
          End If
        End If
      %>
    </div>
  </div>
  <div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <div id="column1">
      <div class="box" id="identity">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.identification", LanguageID))%>
          </span>
        </h2>
        <table id="identitytbl" style="width:320px; border-collapse:collapse; margin:0; padding:0;">
          <tr>
            <td>
              <input type="hidden" id="LayoutID" name="LayoutID" value="<% Sendb(LayoutID) %>" />
              <label for="name"><% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>:</label><br />
              <% If (Layoutname is nothing) then LayoutName = "" %>
              <input type="text" class="longer" id="name" name="name" maxlength="100" value="<% Sendb(LayoutName.Replace("""", "&quot;")) %>" />
            </td>
          </tr>
        </table>
       <hr class="hidden" />
      </div>
    </div>
    
    <div id="gutter">
    </div>
    
    <div id="column2">
      <div class="box" id="dimensions">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.dimensions", LanguageID))%>
          </span>
        </h2>
         <table id="dimensionstbl" style="width:125px; border-collapse:collapse; margin:0; padding:0;" >
          <tr>
            <td>
              <label for="layoutwidth"><% Sendb(Copient.PhraseLib.Lookup("term.width", LanguageID))%>:</label><br />
              <input type="text" class="shorter" id="layoutwidth" name="layoutwidth" maxlength="5" value="<% Sendb(Width) %>" /><br />
            </td>
            <td><br />X</td>
            <td>
              <label for="layoutheight"><% Sendb(Copient.PhraseLib.Lookup("term.height", LanguageID))%>:</label><br />
              <input type="text" class="shorter" id="layoutheight" name="layoutheight" maxlength="5" value="<% Sendb(Height) %>" /><br />
            </td>
          </tr>
        </table>
      </div>
    </div>
    <% If bExistingLayout Then%>
    <div id="column">
      <div class="box" id="cells">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.cells", LanguageID))%>
          </span>
        </h2>
        <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.cells", LanguageID))%>">
          <thead>
            <tr>
              <th class="th-select" scope="col">
                <% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID))%>
              </th>
              <th class="th-name" scope="col">
                <% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>
              </th>
              <th class="th-xpos" scope="col">
                <% Sendb(Copient.PhraseLib.Lookup("term.xpos", LanguageID))%>
              </th>
              <th class="th-ypos" scope="col">
                <% Sendb(Copient.PhraseLib.Lookup("term.ypos", LanguageID))%>
              </th>
              <th class="th-width" scope="col">
                <% Sendb(Copient.PhraseLib.Lookup("term.width", LanguageID))%>
              </th>
              <th class="th-height" scope="col">
                <% Sendb(Copient.PhraseLib.Lookup("term.height", LanguageID))%>
              </th>
              <th class="th-background" scope="col">
                <% Sendb(Copient.PhraseLib.Lookup("term.background", LanguageID))%>
              </th>
            </tr>
          </thead>
          <tbody>
            <%
              MyCommon.QueryStr = "select SC.CellID, SC.Name, SC.X, SC.Y, SC.Width, SC.Height, SC.BackgroundImg, OSA.Name as AdName from ScreenCells as SC with (RowLock) Left Join OnScreenAds as OSA with (RowLock) on SC.BackgroundImg=OnScreenAdID and OSA.Deleted=0 where SC.LayoutID=" & LayoutID & " and SC.Deleted=0;"
              dt = MyCommon.LRT_Select
              NumRecs = dt.Rows.Count
              DisableCellEdit = IIf((NumRecs = 0), " disabled=""disabled""", "")
              If (NumRecs = 0) Then
                Send("<tr>")
                Send("  <td colspan=""8""><i>" & Copient.PhraseLib.Lookup("layout.nocelldefined", LanguageID) & "</i></td>")
                Send("</tr>")
              Else
                For Each row In dt.Rows
                  Send("<tr>")
                  Send("  <td><input type=""radio"" name=""CellID"" id=""CellID" & MyCommon.NZ(row.Item("CellID"), "0") & """ value=""" & MyCommon.NZ(row.Item("CellID"), "0") & """ /></td>")
                  Send("  <td><label for=""CellID" & MyCommon.NZ(row.Item("CellID"), "0") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), "&nbsp;"), 25) & "</label></td>")
                  Send("  <td>" & MyCommon.NZ(row.Item("X"), "0") & "</td>")
                  Send("  <td>" & MyCommon.NZ(row.Item("Y"), "0") & "</td>")
                  Send("  <td>" & MyCommon.NZ(row.Item("Width"), "0") & "</td>")
                  Send("  <td>" & MyCommon.NZ(row.Item("Height"), "0") & "</td>")
                  Send("  <td>" & MyCommon.NZ(row.Item("AdName"), Copient.PhraseLib.Lookup("term.none", LanguageID)) & "</td>")
                  Send("</tr>")
                Next
              End If
            %>
            <tr style="height: 10px;">
              <td colspan="7">
              </td>
            </tr>
            <tr>
              <%  If (Logix.UserRoles.EditLayouts) Then%>
              <td colspan="6">
                <input type="submit" class="regular" id="edit" name="edit" value="<% Sendb(Copient.PhraseLib.Lookup("term.edit", LanguageID))%>"<% sendb(disablecelledit) %> />&nbsp;
                <input type="submit" class="regular" id="remove" name="remove" value="<% Sendb(Copient.PhraseLib.Lookup("term.remove", LanguageID))%>"<% sendb(disablecelledit) %> />&nbsp;
                <input type="submit" class="regular" id="newcell" name="newcell" value="<% Sendb(Copient.PhraseLib.Lookup("term.new", LanguageID))%>" />
              </td>
              <%  Else%>
              <td colspan="6">
              </td>
              <%  End If%>
              <td colspan="1">
                <input type="button" class="regular" id="preview" name="preview" onclick="javascript:openPreviewWin();" value="<% Sendb(Copient.PhraseLib.Lookup("term.preview", LanguageID))%>" />
              </td>
            </tr>
          </tbody>
        </table>
        <br class="half" />
        <hr class="hidden" />
      </div>
    </div>
    <% End If%>
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
    If (LayoutID > 0 And Logix.UserRoles.AccessNotes) Then
      Send_Notes(12, LayoutID, AdminUserID)
    End If
  End If
  Send_BodyEnd("mainform", "name")
done:
  MyCommon.Close_LogixRT()
  Logix = Nothing
  MyCommon = Nothing
%>
