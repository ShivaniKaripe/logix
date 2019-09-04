<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.CommonInc" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%
  ' *****************************************************************************
  ' * FILENAME: category-edit.aspx 
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
  Dim OfferCategoryID As Long = 0
  Dim Description As String = ""
  Dim ExtCategoryID As String = ""
  Dim BaseOfferID As Long = 0
  Dim LastUpdate As String = ""
  Dim SortOrder As String = ""
  Dim TempInt As Integer
  Dim IconFileName As String = ""
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim dst As DataTable
  Dim rst As DataTable
  Dim rstAssociated As DataTable = Nothing
  Dim row As DataRow
  Dim bSave As Boolean
  Dim bDelete As Boolean
  Dim bCreate As Boolean
  Dim DeptNameTitle As String = ""
  Dim ShowActionButton As Boolean = False
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim BannersEnabled As Boolean = False
  Dim HasAssociatedOffers As Boolean = False
  Dim ExpiredOnly As Boolean = False
  Dim DeleteStatusID As Integer = 0
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "category-edit.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  Try
    
    ' fill in if it was a get method
    If Request.RequestType = "GET" Then
      OfferCategoryID = IIf(Request.QueryString("OfferCategoryID") = "", 0, MyCommon.Extract_Val(Request.QueryString("OfferCategoryID")))
      Description = Logix.TrimAll(Request.QueryString("Description"))
      ExtCategoryID = Logix.TrimAll(Request.QueryString("ExtCategoryID"))
      If (Request.QueryString("BaseOfferID") <> "") Then
        BaseOfferID = Int(Request.QueryString("BaseOfferID"))
      End If
      If (Request.QueryString("SortOrder") <> "") Then
        SortOrder = Logix.TrimAll(Request.QueryString("SortOrder"))
      End If
      If (Request.QueryString("IconFileName") <> "") Then
        IconFileName = Logix.TrimAll(Request.QueryString("IconFileName"))
      End If
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
      OfferCategoryID = IIf(Request.Form("OfferCategoryID") = "", 0, MyCommon.Extract_Val(Request.Form("OfferCategoryID")))
      If OfferCategoryID <= 0 Then
        OfferCategoryID = IIf(Request.QueryString("OfferCategoryID") = "", 0, MyCommon.Extract_Val(Request.QueryString("OfferCategoryID")))
      End If
      Description = Logix.TrimAll(Request.Form("Description"))
      ExtCategoryID = Logix.TrimAll(Request.Form("ExtCategoryID"))
      If (Request.QueryString("BaseOfferID") <> "") Then
        BaseOfferID = Int(Request.QueryString("BaseOfferID"))
      End If
      If (Request.QueryString("SortOrder") <> "") Then
        SortOrder = Logix.TrimAll(Request.QueryString("SortOrder"))
      End If
      If (Request.QueryString("IconFileName") <> "") Then
        IconFileName = Logix.TrimAll(Request.QueryString("IconFileName"))
      End If
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
    
    
    BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
    
    Send_HeadBegin("term.category", , OfferCategoryID)
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
    Response.Redirect("category-edit.aspx")
  End If
  
  If (Request.QueryString("ExpiredOnly") = "1") Then
    ExpiredOnly = True
  End If
  
  If bSave Then
    If (Description = "") Then
      infoMessage = Copient.PhraseLib.Lookup("categories.noname", LanguageID)
    ElseIf (ExtCategoryID = "") Then
      infoMessage = Copient.PhraseLib.Lookup("categories.noextid", LanguageID)
    ElseIf (SortOrder <> "") And (Not Integer.TryParse(SortOrder, TempInt) OrElse TempInt <= 0) Then
      infoMessage = Copient.PhraseLib.Lookup("categories.positivesortorder", LanguageID)
    Else
      If (OfferCategoryID = 0) Then
        MyCommon.QueryStr = "SELECT OfferCategoryID FROM OfferCategories with (NoLock) WHERE Deleted=0 and Description='" & MyCommon.Parse_Quotes(Description) & "';"
        dst = MyCommon.LRT_Select
        If (dst.Rows.Count > 0) Then
          infoMessage = Copient.PhraseLib.Lookup("categories.nameused", LanguageID)
        Else
          MyCommon.QueryStr = "SELECT OfferCategoryID FROM OfferCategories with (NoLock) WHERE Deleted=0 and ExtCategoryID='" & ExtCategoryID & "';"
          dst = MyCommon.LRT_Select
          If (dst.Rows.Count > 0) Then
            infoMessage = Copient.PhraseLib.Lookup("categories.extidused", LanguageID)
          Else
            MyCommon.QueryStr = "dbo.pt_OfferCategories_Insert"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 200).Value = Description
            MyCommon.LRTsp.Parameters.Add("@ExtCategoryID", SqlDbType.NVarChar, 20).Value = ExtCategoryID
            MyCommon.LRTsp.Parameters.Add("@BaseOfferID", SqlDbType.BigInt).Value = BaseOfferID
            MyCommon.LRTsp.Parameters.Add("@OfferCategoryID", SqlDbType.Int).Direction = ParameterDirection.Output
            MyCommon.LRTsp.Parameters.Add("@SortOrder", SqlDbType.Int).Value = IIf(SortOrder = "", 0, TempInt)
            MyCommon.LRTsp.Parameters.Add("@IconFileName", SqlDbType.NVarChar, 255).Value = IconFileName
            MyCommon.LRTsp.ExecuteNonQuery()
            OfferCategoryID = MyCommon.LRTsp.Parameters("@OfferCategoryID").Value
            MyCommon.Close_LRTsp()
            MyCommon.Activity_Log(16, OfferCategoryID, AdminUserID, Copient.PhraseLib.Lookup("history.category-create", LanguageID))
            Response.Redirect("category-edit.aspx?OfferCategoryID=" & OfferCategoryID)
          End If
        End If
      Else
        ' update the existing category
        MyCommon.QueryStr = "SELECT OfferCategoryID FROM OfferCategories with (NoLock) WHERE Deleted=0 and Description='" & MyCommon.Parse_Quotes(Description) & "' and OfferCategoryID<>" & OfferCategoryID & ";"
        dst = MyCommon.LRT_Select
        If (dst.Rows.Count > 0) Then
          infoMessage = Copient.PhraseLib.Lookup("categories.nameused", LanguageID)
        Else
          MyCommon.QueryStr = "SELECT OfferCategoryID FROM OfferCategories with (NoLock) WHERE Deleted=0 and ExtCategoryID='" & ExtCategoryID & "' and OfferCategoryID<>" & OfferCategoryID & ";"
          dst = MyCommon.LRT_Select
          If (dst.Rows.Count > 0) Then
            infoMessage = Copient.PhraseLib.Lookup("categories.extidused", LanguageID)
          Else
            MyCommon.QueryStr = "update OfferCategories with (RowLock) set Description='" & MyCommon.Parse_Quotes(Description) & "', ExtCategoryID='" & ExtCategoryID & "', " & _
                                "BaseOfferID=" & IIf(BaseOfferID = 0, "NULL", BaseOfferID) & ", LastUpdate=getdate()" & ", SortOrder=" & IIf(SortOrder = "", "NULL", TempInt) & _
                                ", IconFileName=" & IIf(IconFileName = "", "NULL", "'" & MyCommon.Parse_Quotes(IconFileName) & "'") & " where OfferCategoryID=" & OfferCategoryID & ";"
            MyCommon.LRT_Execute()
            MyCommon.Activity_Log(16, OfferCategoryID, AdminUserID, Copient.PhraseLib.Lookup("history.category-edit", LanguageID))
          End If
        End If
      End If
    End If
    
  ElseIf bDelete Then
    MyCommon.QueryStr = "dbo.pt_OfferCategories_Delete"
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add("@CategoryID", SqlDbType.BigInt).Value = OfferCategoryID
    MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.BigInt).Direction = ParameterDirection.Output
    MyCommon.LRTsp.ExecuteNonQuery()
    DeleteStatusID = MyCommon.LRTsp.Parameters("@Status").Value
    MyCommon.Close_LRTsp()
    If (DeleteStatusID = -1) Then
      infoMessage = Copient.PhraseLib.Lookup("categories.lastcategory", LanguageID)
    ElseIf (DeleteStatusID = -2) Then
      infoMessage = Copient.PhraseLib.Lookup("categories.inuse", LanguageID)
    Else
      If ExpiredOnly Then
        'The category that the user's deleting is associated to expired offers only.
        'The user has previously confirmed that he wants to delete anyway, so we'll
        'delete the category assignments on those offers.
        MyCommon.QueryStr = "update CPE_Incentives set PromoClassID=0 where PromoClassID=" & OfferCategoryID & ";"
        MyCommon.LRT_Execute()
        MyCommon.QueryStr = "update Offers set OfferCategoryID=0 where OfferCategoryID=" & OfferCategoryID & ";"
        MyCommon.LRT_Execute()
      End If
      MyCommon.Activity_Log(16, OfferCategoryID, AdminUserID, Copient.PhraseLib.Lookup("history.category-delete", LanguageID))
      Response.Status = "301 Moved Permanently"
      Response.AddHeader("Location", "category-list.aspx")
    End If
  End If
  
  LastUpdate = ""
  
  If Not bCreate Then
    ' No one clicked anything
    MyCommon.QueryStr = "select OfferCategoryID, Description, ExtCategoryID, BaseOfferID, LastUpdate, SortOrder, IconFileName " & _
                        "from OfferCategories as OC with (NoLock) " & _
                        "where Deleted=0 and OfferCategoryID=" & OfferCategoryID & ";"
    rst = MyCommon.LRT_Select()
    If (rst.Rows.Count > 0) Then
      OfferCategoryID = MyCommon.NZ(rst.Rows(0).Item("OfferCategoryID"), 0)
      Description = MyCommon.NZ(rst.Rows(0).Item("Description"), "")
      ExtCategoryID = MyCommon.NZ(rst.Rows(0).Item("ExtCategoryID"), "")
      BaseOfferID = MyCommon.NZ(rst.Rows(0).Item("BaseOfferID"), 0)
      SortOrder = MyCommon.NZ(rst.Rows(0).Item("SortOrder"), "")
      IconFileName = MyCommon.NZ(rst.Rows(0).Item("IconFileName"), "")
      If (IsDBNull(rst.Rows(0).Item("LastUpdate"))) Then
        LastUpdate = Copient.PhraseLib.Lookup("term.unknown", LanguageID)
      Else
        LastUpdate = Logix.ToLongDateTimeString(rst.Rows(0).Item("LastUpdate"), MyCommon)
      End If
    ElseIf (OfferCategoryID > 0) Then
      Send("")
      Send("<div id=""intro"">")
      Send("  <h1 id=""title"">" & Copient.PhraseLib.Lookup("term.category", LanguageID) & " #" & OfferCategoryID & "</h1>")
      Send("</div>")
      Send("<div id=""main"">")
      Send("  <div id=""infobar"" class=""red-background"">")
      Send("    " & Copient.PhraseLib.Lookup("error.deleted", LanguageID))
      Send("  </div>")
      Send("</div>")
      GoTo done
    End If
  End If
  
  MyCommon.QueryStr = "select IncentiveID as OfferID, IncentiveName as Name, StartDate, EndDate,buy.ExternalBuyerId as BuyerID from CPE_Incentives CPE   with (NoLock) left outer join Buyers as buy with (nolock) on buy.BuyerId= CPE.BuyerId " & _
                      "where PromoClassID=" & OfferCategoryID & " and Deleted=0 " & _
                      " union " & _
                      "select OfferID, Name, ProdStartDate as StartDate, ProdEndDate as EndDate,NULL as BuyerID from Offers with (NoLock) " & _
                      "where OfferCategoryID=" & OfferCategoryID & " and Deleted=0;"
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
%>
<form action="#" id="mainform" name="mainform">
<input type="hidden" id="OfferCategoryID" name="OfferCategoryID" value="<% Sendb(OfferCategoryID) %>" />
<input type="hidden" id="ExpiredOnly" name="ExpiredOnly" value="<% Sendb(IIf(ExpiredOnly AndAlso OfferCategoryID>0, 1, 0)) %>" />
<div id="intro">
  <%
    Sendb("<h1 id=""title"">")
    If OfferCategoryID = 0 Then
      Sendb(Copient.PhraseLib.Lookup("term.newcategory", LanguageID))
    Else
      Sendb(Copient.PhraseLib.Lookup("term.category", LanguageID) & " #" & OfferCategoryID & ": " & MyCommon.TruncateString(Description, 40))
    End If
    Send("</h1>")
    Send("<div id=""controls"">")
    If (OfferCategoryID = 0) Then
      If (Logix.UserRoles.EditCategories) Then
        Send_Save()
      End If
    Else
      ShowActionButton = (Logix.UserRoles.EditCategories)
      If (ShowActionButton) Then
        Send("<input type=""button"" class=""regular"" id=""actions"" name=""actions"" value=""" & Copient.PhraseLib.Lookup("term.actions", LanguageID) & " ▼"" onclick=""toggleDropdown();"" />")
        Send("<div class=""actionsmenu"" id=""actionsmenu"">")
        If (Logix.UserRoles.EditCategories) Then
          Send_Save()
          If ExpiredOnly Then
            Send("<input type=""submit"" accesskey=""d"" class=""regular"" id=""delete"" name=""delete"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("confirm.categorydelete", LanguageID) & "')){}else{return false}"" value=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ />")
          Else
            Send("<input type=""submit"" accesskey=""d"" class=""regular"" id=""delete"" name=""delete"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("confirm.delete", LanguageID) & "')){}else{return false}"" value=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ />")
          End If
          Send_New()
        End If
        Send("</div>")
      End If
      If MyCommon.Fetch_SystemOption(75) Then
        If (Logix.UserRoles.AccessNotes) Then
          Send_NotesButton(21, OfferCategoryID, AdminUserID)
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
        Send("<label for=""Description"">" & Copient.PhraseLib.Lookup("term.name", LanguageID) & ":</label><br />")
        If (Description Is Nothing) Then
          Description = ""
        End If
        Sendb("<input type=""text"" class=""longest"" id=""Description"" name=""Description"" maxlength=""100"" value=""" & Description.Replace("""", "&quot;") & """ />")
        Send("<br />")
        Send("<br class=""half"" />")
        Send("<label for=""ExtCategoryID"">" & Copient.PhraseLib.Lookup("term.externalid", LanguageID) & ":</label><br />")
        Send("<input type=""text"" class=""longest"" id=""ExtCategoryID"" name=""ExtCategoryID"" maxlength=""20"" value=""" & ExtCategoryID & """ />")
        Send("<br />")
        Send("<br class=""half"" />")
        If (OfferCategoryID > 0) Then
          Send("<label for=""BaseOfferID"">" & Copient.PhraseLib.Lookup("term.baseoffer", LanguageID) & ":</label><br />")
          Send("<select class=""longest"" id=""BaseOfferID"" name=""BaseOfferID"">")
          Send("  <option value=""0"">" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "</option>")
          MyCommon.QueryStr = "select IncentiveID as OfferID, IncentiveName as Name, StartDate, EndDate from CPE_Incentives with (NoLock) " & _
                              "where PromoClassID=" & OfferCategoryID & " and Deleted=0" & _
                              " union " & _
                              "select OfferID, Name, ProdStartDate as StartDate, ProdEndDate as EndDate from Offers with (NoLock) " & _
                              "where OfferCategoryID=" & OfferCategoryID & " and Deleted=0;"
          rst = MyCommon.LRT_Select
          If rst.Rows.Count > 0 Then
            For Each row In rst.Rows
              Sendb("  <option value=""" & MyCommon.NZ(row.Item("OfferID"), 0) & """" & IIf(MyCommon.NZ(row.Item("OfferID"), 0) = BaseOfferID, " selected=""selected""", "") & ">")
              Sendb(MyCommon.NZ(row.Item("OfferID"), 0))
              If MyCommon.NZ(row.Item("Name"), "") <> "" Then
                Sendb(" - " & MyCommon.NZ(row.Item("Name"), ""))
              End If
              Send("</option>")
            Next
          End If
          Send("</select><br />")
          Send("<br />")
          If MyCommon.Fetch_CPE_SystemOption(132) Then
            Send("<label for=""SortOrder"">" & Copient.PhraseLib.Lookup("term.sortorder", LanguageID) & ":</label><br />")
            Send("<input type=""text"" class=""longest"" id=""SortOrder"" name=""SortOrder"" maxlength=""10"" value=""" & SortOrder & """/><br /><br />")
            Send("<label for=""IconFileName"">" & Copient.PhraseLib.Lookup("term.iconfilename", LanguageID) & ":</label><br />")
            If (IconFileName Is Nothing) Then
              IconFileName = ""
            End If
            Send("<input type=""text"" class=""longest"" id=""IconFileName"" name=""IconFileName"" maxlength=""255""  value=""" & IconFileName.Replace("""", "&quot;") & """/><br /><br />")
          End If
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
    <div class="box" id="offers" <%if(OfferCategoryID = 0)then sendb(" style=""visibility: hidden;""") %>>
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.associatedoffers", LanguageID))%>
        </span>
      </h2>
      <div class="boxscroll">
        <% 
            Dim assocName As String=""
          If (OfferCategoryID > 0) Then
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
      Send_Notes(21, OfferCategoryID, AdminUserID)
    End If
  End If
done:
Finally
  MyCommon.Close_LogixRT()
End Try
Send_BodyEnd("mainform", "Description")
MyCommon = Nothing
Logix = Nothing
%>
