<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.CommonInc" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%
  ' *****************************************************************************
  ' * FILENAME: meg-edit.aspx 
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
  Dim dst As DataTable
  Dim rst As DataTable
  Dim rstAssociated As DataTable = Nothing
  Dim row As DataRow
  Dim MutualExclusionGroupID As Long = 0
  Dim Name As String = ""
  Dim ExtGroupID As String = ""
  Dim Description As String = ""
  Dim ItemLevel As Boolean = False
  Dim LastUpdate As String = ""
  Dim bSave As Boolean
  Dim bDelete As Boolean
  Dim bCreate As Boolean
  Dim ShowActionButton As Boolean = False
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim HasAssociatedOffers As Boolean = False
  Dim ExpiredOnly As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "meg-edit.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  Try
    ' fill in if it was a get method
    If Request.RequestType = "GET" Then
      MutualExclusionGroupID = IIf(Request.QueryString("MutualExclusionGroupID") = "", 0, MyCommon.Extract_Val(Request.QueryString("MutualExclusionGroupID")))
      Name = Logix.TrimAll(Request.QueryString("Name"))
	  ExtGroupID = Logix.TrimAll(Request.QueryString("ExtGroupID"))
      Description = Left(Logix.TrimAll(Request.QueryString("Description")), 1000)
      ItemLevel = IIf(Request.QueryString("ItemLevel") = "1", True, False)
      If Request.QueryString("save") = "" Then
        bSave = False
      Else
        bSave = True
      End If
      If Request.QueryString("delete") = "" Then
        bDelete = False
      Else
        bDelete = True
      End If
      If Request.QueryString("mode") = "" Then
        bCreate = False
      Else
        bCreate = True
      End If
    Else
      MutualExclusionGroupID = IIf(Request.Form("MutualExclusionGroupID") = "", 0, MyCommon.Extract_Val(Request.Form("MutualExclusionGroupID")))
      If MutualExclusionGroupID <= 0 Then
        MutualExclusionGroupID = IIf(Request.QueryString("MutualExclusionGroupID") = "", 0, MyCommon.Extract_Val(Request.QueryString("MutualExclusionGroupID")))
      End If
      Name = Logix.TrimAll(Request.Form("Name"))
	  ExtGroupID = Logix.TrimAll(Request.Form("ExtGroupID"))
      Description = Left(Logix.TrimAll(Request.Form("Description")), 1000)
      ItemLevel = IIf(Request.Form("ItemLevel") = "1", True, False)
      If Request.Form("save") = "" Then
        bSave = False
      Else
        bSave = True
      End If
      If Request.Form("delete") = "" Then
        bDelete = False
      Else
        bDelete = True
      End If
      If Request.Form("mode") = "" Then
        bCreate = False
      Else
        bCreate = True
      End If
    End If
    
    Send_HeadBegin("term.MutualExclusionGroup", , MutualExclusionGroupID)
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts()
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
<%
  Send_HeadEnd()
  Send_BodyBegin(1)
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  Send_Tabs(Logix, 8)
  Send_Subtabs(Logix, 8, 4)
  
  If (Logix.UserRoles.EditSystemConfiguration = False) Then
    Send_Denied(1, "perm.admin-configuration")
    GoTo done
  End If
  
  If (Request.QueryString("new") <> "") Then
    Response.Redirect("meg-edit.aspx")
  End If
  
  MyCommon.QueryStr = "select MEGO.OfferID, IncentiveName as Name, StartDate, EndDate,buy.ExternalBuyerId as BuyerID " & _
                    "from MutualExclusionGroupOffers as MEGO with (NoLock) " & _
                    "left join CPE_Incentives as I with (NoLock) on I.IncentiveID=MEGO.OfferID " & _
                    "left outer join Buyers as buy with (nolock) on buy.BuyerId= I.BuyerId " & _
                    "where MutualExclusionGroupID=" & MutualExclusionGroupID & " and Deleted=0;"
  rstAssociated = MyCommon.LRT_Select
  HasAssociatedOffers = (rstAssociated.Rows.Count > 0)
  
  'If there are associated offers, determine if they're all expired.
  If HasAssociatedOffers Then
    For Each row In rstAssociated.Rows
      If (MyCommon.NZ(row.Item("EndDate"), Today) < Today) Then
        ExpiredOnly = True
      Else
        ExpiredOnly = False
        Exit For
      End If
    Next
  End If
  
  If bSave Then
    If (Name = "") Then
      infoMessage = Copient.PhraseLib.Lookup("meg.noname", LanguageID)
    Else
      If (MutualExclusionGroupID = 0) Then
        MyCommon.QueryStr = "SELECT MutualExclusionGroupID FROM MutualExclusionGroups with (NoLock) WHERE Deleted=0 and Name='" & MyCommon.Parse_Quotes(Name) & "';"
        dst = MyCommon.LRT_Select
        If (dst.Rows.Count > 0) Then
          infoMessage = Copient.PhraseLib.Lookup("meg.nameused", LanguageID)
        Else
          MyCommon.QueryStr = "dbo.pt_MutualExclusionGroups_Insert"
          MyCommon.Open_LRTsp()
          MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 255).Value = Name
		  MyCommon.LRTsp.Parameters.Add("@ExtGroupID", SqlDbType.NVarChar, 20).Value = ExtGroupID
          MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 1000).Value = Description
          MyCommon.LRTsp.Parameters.Add("@ItemLevel", SqlDbType.Bit).Value = ItemLevel
          MyCommon.LRTsp.Parameters.Add("@MutualExclusionGroupID", SqlDbType.Int).Direction = ParameterDirection.Output
          MyCommon.LRTsp.ExecuteNonQuery()
          MutualExclusionGroupID = MyCommon.LRTsp.Parameters("@MutualExclusionGroupID").Value
          MyCommon.Close_LRTsp()
          MyCommon.Activity_Log(52, MutualExclusionGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.meg-create", LanguageID))
          Response.Redirect("meg-edit.aspx?MutualExclusionGroupID=" & MutualExclusionGroupID)
        End If
      Else
        ' update the existing group
        MyCommon.QueryStr = "SELECT MutualExclusionGroupID FROM MutualExclusionGroups with (NoLock) WHERE Deleted=0 and Name='" & MyCommon.Parse_Quotes(Name) & "' and MutualExclusionGroupID<>" & MutualExclusionGroupID & ";"
        dst = MyCommon.LRT_Select
        If (dst.Rows.Count > 0) Then
          infoMessage = Copient.PhraseLib.Lookup("meg.nameused", LanguageID)
        Else
          MyCommon.QueryStr = "update MutualExclusionGroups with (RowLock) " &
                                        "set Name=N'" & MyCommon.Parse_Quotes(Name) & "', ExtGroupID=N'" & MyCommon.Parse_Quotes(ExtGroupID) & "', Description=N'" & MyCommon.Parse_Quotes(Description) & "', ItemLevel=" & IIf(ItemLevel, 1, 0) & ", LastUpdate=getdate() " &
                                        "where MutualExclusionGroupID=" & MutualExclusionGroupID & ";"
          MyCommon.LRT_Execute()
          MyCommon.Activity_Log(52, MutualExclusionGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.meg-edit", LanguageID))
        End If
      End If
    End If
    
  ElseIf bDelete Then
    If (rstAssociated.Rows.Count > 0) And (ExpiredOnly = False) Then
      infoMessage = Copient.PhraseLib.Lookup("meg.inuse", LanguageID)
    Else
      MyCommon.QueryStr = "dbo.pt_MutualExclusionGroups_Delete"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@MutualExclusionGroupID", SqlDbType.BigInt).Value = MutualExclusionGroupID
      MyCommon.LRTsp.ExecuteNonQuery()
      MyCommon.Close_LRTsp()
      If ExpiredOnly Then
        MyCommon.QueryStr = "delete from MutualExclusionGroupOffers where MutualExclusionGroupID=" & MutualExclusionGroupID & ";"
        MyCommon.LRT_Execute()
      End If
      MyCommon.Activity_Log(52, MutualExclusionGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.meg-delete", LanguageID))
      Response.Status = "301 Moved Permanently"
      Response.AddHeader("Location", "MEG-list.aspx")
    End If
  End If
  
  LastUpdate = ""
  
  If Not bCreate Then
    ' No one clicked anything
    MyCommon.QueryStr = "select MutualExclusionGroupID, Name, Description, ItemLevel, LastUpdate, ExtGroupID " &
                            "from MutualExclusionGroups as MEG with (NoLock) " &
                            "where Deleted=0 and MutualExclusionGroupID=" & MutualExclusionGroupID & ";"
    rst = MyCommon.LRT_Select()
    If (rst.Rows.Count > 0) Then
      MutualExclusionGroupID = MyCommon.NZ(rst.Rows(0).Item("MutualExclusionGroupID"), 0)
      Name = MyCommon.NZ(rst.Rows(0).Item("Name"), "")
	  ExtGroupID = MyCommon.NZ(rst.Rows(0).Item("ExtGroupID"), "")
      Description = MyCommon.NZ(rst.Rows(0).Item("Description"), "")
      ItemLevel = IIf(rst.Rows(0).Item("ItemLevel") = 0, False, True)
      If (IsDBNull(rst.Rows(0).Item("LastUpdate"))) Then
        LastUpdate = Copient.PhraseLib.Lookup("term.unknown", LanguageID)
      Else
        LastUpdate = Logix.ToLongDateTimeString(rst.Rows(0).Item("LastUpdate"), MyCommon)
      End If
    ElseIf (MutualExclusionGroupID > 0) Then
      Send("")
      Send("<div id=""intro"">")
      Send("  <h1 id=""title"">" & Copient.PhraseLib.Lookup("term.MutualExclusionGroup", LanguageID) & " #" & MutualExclusionGroupID & "</h1>")
      Send("</div>")
      Send("<div id=""main"">")
      Send("  <div id=""infobar"" class=""red-background"">")
      Send("    " & Copient.PhraseLib.Lookup("error.deleted", LanguageID))
      Send("  </div>")
      Send("</div>")
      GoTo done
    End If
  End If
%>
<form action="MEG-edit.aspx" id="mainform" name="mainform">
<input type="hidden" id="MutualExclusionGroupID" name="MutualExclusionGroupID" value="<% Sendb(MutualExclusionGroupID) %>" />
<input type="hidden" id="ExpiredOnly" name="ExpiredOnly" value="<% Sendb(IIf(ExpiredOnly AndAlso MutualExclusionGroupID>0, 1, 0)) %>" />
<div id="intro">
  <%
    Sendb("<h1 id=""title"">")
    If MutualExclusionGroupID = 0 Then
      Sendb(Copient.PhraseLib.Lookup("term.NewMutualExclusionGroup", LanguageID))
    Else
      Sendb(Copient.PhraseLib.Lookup("term.MutualExclusionGroup", LanguageID) & " #" & MutualExclusionGroupID & ": " & MyCommon.TruncateString(Description, 40))
    End If
    Send("</h1>")
    Send("<div id=""controls"">")
    If (MutualExclusionGroupID = 0) Then
      If (Logix.UserRoles.EditMutualExclusionGroups) Then
        Send_Save()
      End If
    Else
      ShowActionButton = (Logix.UserRoles.EditMutualExclusionGroups)
      If (ShowActionButton) Then
        Send("<input type=""button"" class=""regular"" id=""actions"" name=""actions"" value=""" & Copient.PhraseLib.Lookup("term.actions", LanguageID) & " ▼"" onclick=""toggleDropdown();"" />")
        Send("<div class=""actionsmenu"" id=""actionsmenu"">")
        If (Logix.UserRoles.EditMutualExclusionGroups) Then
          Send_Save()
          If ExpiredOnly Then
            Send("<input type=""submit"" accesskey=""d"" class=""regular"" id=""delete"" name=""delete"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("meg.confirmdelete", LanguageID) & "')){}else{return false}"" value=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ />")
          Else
            Send("<input type=""submit"" accesskey=""d"" class=""regular"" id=""delete"" name=""delete"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("confirm.delete", LanguageID) & "')){}else{return false}"" value=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ />")
          End If
          Send_New()
        End If
        Send("</div>")
      End If
      If MyCommon.Fetch_SystemOption(75) Then
        If (Logix.UserRoles.AccessNotes) Then
          Send_NotesButton(21, MutualExclusionGroupID, AdminUserID)
        End If
      End If
    End If
    Send("</div>")
  %>
</div>
<div id="main">
  <%
    If (infoMessage <> "") Then
      Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
    End If
  %>
  <div id="column1">
    <div class="box" id="identification">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.identification", LanguageID))%>
        </span>
      </h2>
      <%
        If (Name Is Nothing) Then Name = ""
        If (Description Is Nothing) Then Description = ""
        
        Send("<label for=""Name"">" & Copient.PhraseLib.Lookup("term.name", LanguageID) & ":</label><br />")
        Send("<input type=""text"" class=""longest"" id=""Name"" name=""Name"" maxlength=""255"" value=""" & Name.Replace("""", "&quot;") & """ />")
        Send("<br />")
        Send("<br class=""half"" />")
        
		Send("<label for=""ExtGroupID"">" & Copient.PhraseLib.Lookup("term.ExternalID", LanguageID) & ":</label><br />")
          Send("<input type=""text"" class=""longest"" id=""ExtGroupID"" name=""ExtGroupID"" maxlength=""20"" value=""" & ExtGroupID.Replace("""", "&quot;") & """ />")
          Send("<br />")
          Send("<br class=""half"" />")
		
        Send("<label for=""Description"">" & Copient.PhraseLib.Lookup("term.description", LanguageID) & ":</label><br />")
          Send("<textarea class=""longest"" id=""Description"" name=""Description"" oninput=""limitText(this,1000);"">" & Description.Replace("""", "&quot;") & "</textarea>")
        Send("<br />")
        Send("<br class=""half"" />")
        
        Sendb(Copient.PhraseLib.Lookup("term.level", LanguageID) & ":")
        If MutualExclusionGroupID > 0 Then
          Send(" " & Copient.PhraseLib.Lookup(IIf(ItemLevel, "term.itemlevel", "term.offerlevel"), LanguageID) & "<br />")
          Send("<input type=""hidden"" id=""ItemLevel"" name=""ItemLevel"" value=""" & IIf(ItemLevel, 1, 0) & """ />")
        Else
          Send("<br />")
          Send("<input type=""radio"" id=""ItemLevelTrue"" name=""ItemLevel""" & IIf(ItemLevel, " checked=""checked""", "") & " value=""1"" /><label for=""ItemLevelTrue"">" & Copient.PhraseLib.Lookup("term.itemlevel", LanguageID) & "</label><br />")
          Send("<input type=""radio"" id=""ItemLevelFalse"" name=""ItemLevel""" & IIf(ItemLevel, "", " checked=""checked""") & " value=""0"" /><label for=""ItemLevelFalse"">" & Copient.PhraseLib.Lookup("term.offerlevel", LanguageID) & "</label><br />")
        End If
        
        If (MutualExclusionGroupID > 0) Then
          Send("<br class=""half"" />")
          Sendb(Copient.PhraseLib.Lookup("term.edited", LanguageID) & " " & LastUpdate)
          Send("<br />")
        End If
      %>
      <hr class="hidden" />
    </div>
  </div>
  
  <div id="gutter">
  </div>
  
  <div id="column2">
    <div class="box" id="offers" <%if(MutualExclusionGroupID = 0)then sendb(" style=""visibility: hidden;""") %>>
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.associatedoffers", LanguageID))%>
        </span>
      </h2>
      <div class="boxscroll">
        <% 
             Dim assocName As String=""
          If (MutualExclusionGroupID > 0) Then
            If rstAssociated.Rows.Count > 0 Then
              For Each row In rstAssociated.Rows
                If (MyCommon.Fetch_UE_SystemOption(168) = "1" AndAlso MyCommon.NZ(row.Item("BuyerID"), "") <> "") Then
                assocName = "Buyer " + row.Item("BuyerID").ToString() + " - " + MyCommon.NZ(row.Item("Name"), "").ToString()
                Else
                assocName = MyCommon.NZ(row.Item("Name"),Copient.PhraseLib.Lookup("term.unknown", LanguageID))
                End If
                If (Logix.IsAccessibleOffer(AdminUserID, row.Item("OfferID"))) Then
                  Sendb(" <a href=""offer-redirect.aspx?OfferID=" & row.Item("OfferID") & """>" & assocName & "</a>")
                Else
                  Sendb(assocName)
                End If
                If (MyCommon.NZ(row.Item("EndDate"), Today) < Today) Then
                  Sendb(" (" & Copient.PhraseLib.Lookup("term.expired", LanguageID) & ")")
                End If
                Send("<br />")
              Next
            Else
              Send("                " & Copient.PhraseLib.Lookup("term.none", LanguageID) & "<br />")
            End If
          Else
            Send("                " & Copient.PhraseLib.Lookup("term.none", LanguageID) & "<br />")
          End If
        %>
      </div>
      <hr class="hidden" />
    </div>
  </div>
  <br clear="all" />
  
</div>
</form>

<script type="text/javascript">
  if (window.captureEvents) {
    window.captureEvents(Event.CLICK);
    window.onclick = handlePageClick;
  }
  else {
    document.onclick = handlePageClick;
  }
</script>

<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (Logix.UserRoles.AccessNotes) Then
      Send_Notes(21, MutualExclusionGroupID, AdminUserID)
    End If
  End If
done:
Finally
  MyCommon.Close_LogixRT()
End Try
Send_BodyEnd("mainform", "Name")
MyCommon = Nothing
Logix = Nothing
%>
