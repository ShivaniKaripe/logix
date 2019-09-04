<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: categories.aspx 
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
  
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim dst As System.Data.DataTable
  Dim row As DataRow
  Dim l_name As String
  Dim l_OCID As Long
  Dim AdminUserID As Long
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim DeleteStatusID As Integer = 0
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "categories.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  Send_HeadBegin("term.categories")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
%>

<script type="text/javascript" language="javascript">
    function handleDelete_Click() {
      var elem = document.getElementById("categories");
      var retVal = true;
      
      if (elem != null) {
        if (elem.selectedIndex > -1) {
          retVal = confirm('<% Sendb(Copient.PhraseLib.Lookup("confirm.delete", LanguageID)) %>')
        } else {
          retVal = false;
          alert('<% Sendb(Copient.PhraseLib.Lookup("categories.notselected", LanguageID)) %>');
        }
      }
      return retVal;
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
  
  If (Request.QueryString("add") <> "") Then
    l_name = MyCommon.NZ(Request.QueryString("name"), "")
    l_name = MyCommon.Parse_Quotes(Logix.TrimAll(l_name))
    If (l_name = "") Then
      infoMessage = Copient.PhraseLib.Lookup("categories.noname", LanguageID)
    Else
      MyCommon.QueryStr = "SELECT OfferCategoryID FROM OfferCategories with (NoLock) WHERE Deleted=0 and Description = '" & l_name & "'"
      dst = MyCommon.LRT_Select
      If (dst.Rows.Count > 0) Then
        infoMessage = Copient.PhraseLib.Lookup("categories.nameused", LanguageID)
      Else
        MyCommon.QueryStr = "INSERT INTO OfferCategories with (RowLock) (Description) VALUES (N'" & l_name & "')"
        MyCommon.LRT_Execute()
        MyCommon.Activity_Log(16, l_OCID, AdminUserID, Copient.PhraseLib.Lookup("history.category-create", LanguageID))
      End If
    End If
  ElseIf (Request.QueryString("delete") <> "") Then
    l_OCID = MyCommon.Extract_Val(Request.QueryString("categories"))
    
    MyCommon.QueryStr = "dbo.pt_OfferCategories_Delete"
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add("@CategoryID", SqlDbType.BigInt).Value = l_OCID
    MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.BigInt).Direction = ParameterDirection.Output
    MyCommon.LRTsp.ExecuteNonQuery()
    DeleteStatusID = MyCommon.LRTsp.Parameters("@Status").Value
    MyCommon.Close_LRTsp()
    
    If (DeleteStatusID = -1) Then
      infoMessage = Copient.PhraseLib.Lookup("categories.lastcategory", LanguageID)
    ElseIf (DeleteStatusID = -2) Then
      infoMessage = Copient.PhraseLib.Lookup("categories.inuse", LanguageID)
    Else
      MyCommon.Activity_Log(16, l_OCID, AdminUserID, Copient.PhraseLib.Lookup("history.category-delete", LanguageID))
    End If
  End If
%>
<form action="#" id="mainform" name="mainform">
  <div id="intro">
    <h1 id="title">
      <% Sendb(Copient.PhraseLib.Lookup("term.categories", LanguageID))%>
    </h1>
    <div id="controls">
      <%
        If MyCommon.Fetch_SystemOption(75) Then
          If (Logix.UserRoles.AccessNotes) Then
            Send_NotesButton(21, 0, AdminUserID)
          End If
        End If
      %>
    </div>
  </div>
  <div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <div id="column1">
      <%If (Logix.UserRoles.EditCategories = False) Then
          Send("<select class=""long"" id=""categories"" name=""categories"" size=""15"">")
          MyCommon.QueryStr = "SELECT OfferCategoryID, Description FROM OfferCategories with (NoLock) where Deleted=0 order by Description"
          dst = MyCommon.LRT_Select
          For Each row In dst.Rows
            Send("    <option value=""" & MyCommon.NZ(row.Item("OfferCategoryID"), 0) & """>" & MyCommon.NZ(row.Item("Description"), "") & "</option>")
          Next
          Send("</select>")
          Send("</div>")
          Send("</form>")
          GoTo done
        End If
      %>
      <div class="box" id="catadd">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("categories.add", LanguageID))%>
          </span>
        </h2>
        <label for="name"><% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>:</label><br />
        <input type="text" class="mediumlong" id="name" name="name" value="" maxlength="50" />
        <input type="submit" class="regular" id="add" name="add" value="<% Sendb(Copient.PhraseLib.Lookup("term.add", LanguageID)) %>"<%if(not logix.userroles.editcategories)then sendb(" disabled=""disabled """) %> /><br />
        <hr class="hidden" />
      </div>
      <div class="box" id="catdel">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("categories.delete", LanguageID))%>
          </span>
        </h2>
        <select class="longest" id="categories" name="categories" size="15">
          <%
            MyCommon.QueryStr = "SELECT OfferCategoryID, Description FROM OfferCategories with (NoLock) where Deleted=0 order by Description"
            dst = MyCommon.LRT_Select
            For Each row In dst.Rows
              Send("<option value=""" & MyCommon.NZ(row.Item("OfferCategoryID"), 0) & """>" & MyCommon.NZ(row.Item("Description"), "") & "</option>")
            Next
          %>
        </select>
        <br />
        <input type="submit" class="regular" id="delete" name="delete" onclick="javascript:return handleDelete_Click();" value="<% Sendb(Copient.PhraseLib.Lookup("term.delete", LanguageID)) %>"<%if(not logix.userroles.editcategories)then sendb(" disabled=""disabled""") %> /><br />
      </div>
    </div>
    <br clear="all" />
  </div>
</form>
<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (Logix.UserRoles.AccessNotes) Then
      Send_Notes(21, 0, AdminUserID)
    End If
  End If
done:
  Send_BodyEnd("mainform", "name")
  MyCommon.Close_LogixRT()
  MyCommon = Nothing
  Logix = Nothing
%>
