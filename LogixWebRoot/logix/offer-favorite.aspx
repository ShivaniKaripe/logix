<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: offer-favorite.aspx 
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
  Dim rst As DataTable
  Dim row As DataRow
  Dim rst2 As DataTable
  Dim row2 As DataRow
  Dim i As Integer = 0
  Dim FavoriteCount As Integer = 0
  Dim OfferID As Long
  Dim Name As String = ""
  Dim isTemplate As Boolean
  Dim historyString As String = ""
  Dim tmpString As String = ""
  Dim shaded As String = ""
  Dim MaxRoleID As Integer = 0
  Dim MaxUserID As Integer = 0
  Dim byRole As Boolean = True
  Dim CloseAfterSave As Boolean = False
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim EngineID As Integer = 2
  Dim CustomerInquiry As Boolean = False
  Dim BannersEnabled As Boolean = True
  Dim UserCount As Integer = 0
  Dim TotalUserCount As Integer = 0
  Dim bUseTemplateLocks As Boolean = False
  Dim Disallow_AdvancedOption As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "offer-favorite.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  If Request.QueryString("OfferID") <> "" Then
    OfferID = Request.QueryString("OfferID")
  End If
  If Request.QueryString("CustomerInquiry") <> "" Then
    CustomerInquiry = True
  End If
  
  
  If OfferID <> 0 Then
    MyCommon.QueryStr = "select EngineID from OfferIDs where OfferID=" & OfferID & ";"
    rst = MyCommon.LRT_Select
    If (Request.QueryString("EngineID") <> "") Then
      EngineID = MyCommon.Extract_Val(Request.QueryString("EngineID"))
    Else
      If rst.Rows.Count > 0 Then
        EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), 0)
      Else
        EngineID = 0
      End If
    End If
  End If
  If (Request.QueryString("Disallow_AdvancedOption") <> "") Then
    Disallow_AdvancedOption = Request.QueryString("Disallow_AdvancedOption")
  End If
  If (Request.QueryString("bUseTemplateLocks") <> "") Then
    bUseTemplateLocks = Request.QueryString("bUseTemplateLocks")
  End If
  
  MyCommon.QueryStr = "select top 1 AdminUserID from AdminUsers order by AdminUserID DESC;"
  rst = MyCommon.LRT_Select
  MaxUserID = MyCommon.NZ(rst.Rows(0).Item("AdminUserID"), 0)
  MyCommon.QueryStr = "select top 1 RoleID from AdminRoles order by RoleID DESC;"
  rst = MyCommon.LRT_Select
  MaxRoleID = MyCommon.NZ(rst.Rows(0).Item("RoleID"), 0)
  
  MyCommon.QueryStr = "select count(AdminUserID) as TotalUsers from AdminUsers with (NoLock);"
  rst = MyCommon.LRT_Select
  If (rst.Rows.Count > 0) Then
    TotalUserCount = MyCommon.NZ(rst.Rows(0).Item("TotalUsers"), 0)
  End If
  
  If (Request.QueryString("by") = "user") Then
    byRole = False
  Else
    byRole = True
  End If
  
  If (Request.QueryString("save") <> "") Then
    ' Delete all the favorites for this offer
    MyCommon.QueryStr = "delete from AdminUserOffers with (RowLock) where OfferID=" & OfferID & ";"
    MyCommon.LRT_Execute()
    
    If byRole Then
      ' Save by role
      MyCommon.QueryStr = "select RoleID from AdminRoles with (NoLock) order by DisplayOrder DESC;"
      rst = MyCommon.LRT_Select
      For Each row In rst.Rows
        tmpString = "role" & row.Item("RoleID")
        If (Request.QueryString(tmpString) = "on") Then
          MyCommon.QueryStr = "select AdminUserID from AdminUserRoles where RoleID=" & MyCommon.NZ(row.Item("roleID"), 0) & " and AdminUserID > 0;"
          rst2 = MyCommon.LRT_Select
          For Each row2 In rst2.Rows
            MyCommon.QueryStr = "insert into AdminUserOffers with (RowLock) (AdminUserID, OfferID, Priority, FavoredBy, FavoredDate) " & _
                                "values(" & MyCommon.NZ(row2.Item("AdminUserID"), 0) & ", " & OfferID & ", 1, " & AdminUserID & ", '" & Now() & "')"
            MyCommon.LRT_Execute()
            UserCount += 1
          Next
        End If
      Next
    Else
      ' Save by user
      MyCommon.QueryStr = "select AdminUserID from AdminUsers with (NoLock) order by LastName;"
      rst = MyCommon.LRT_Select
      For Each row In rst.Rows
        tmpString = "user" & row.Item("AdminUserID")
        If (Request.QueryString(tmpString) = "on") Then
          MyCommon.QueryStr = "insert into AdminUserOffers with (RowLock) (AdminUserID, OfferID, Priority, FavoredBy, FavoredDate) " & _
                              "values(" & MyCommon.NZ(row.Item("AdminUserID"), 0) & ", " & OfferID & ", 1, " & AdminUserID & ", '" & Now() & "')"
          MyCommon.LRT_Execute()
          i = i + 1
          UserCount += 1
        End If
      Next
    End If
    historyString = Copient.PhraseLib.Lookup("history.offer-favorites", LanguageID)
    CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
    MyCommon.Activity_Log(3, OfferID, AdminUserID, historyString)
  End If
  
  Send_HeadBegin("term.offer", "term.favorites", OfferID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
%>
<style type="text/css">
.favoritedetails {
  background-color: #dddddd;
}
</style>
<%
  Send_Scripts()
%>
<script type="text/javascript">
  function handleAllUsers() {
    var elem = null;
    var i = 1;
    
    var elemAll = document.getElementById("chkAllUsers");
    
    if (elemAll != null) {
      elem = document.getElementById("user" + i);
      while (i <= <% Sendb(MaxUserID) %>) {
        if (elem != null) {
          elem.checked = elemAll.checked;
          i++;
          elem = document.getElementById("user" + i);
        } else {
          i++;
          elem = document.getElementById("user" + i);
        }
      }
    }
  }
  function handleAllRoles() {
    var elem = null;
    var i = 1;
    
    var elemAll = document.getElementById("chkAllRoles");
    
    if (elemAll != null) {
      elem = document.getElementById("role" + i);
      while (i <= <% Sendb(MaxRoleID) %>) {
        if (elem != null) {
          elem.checked = elemAll.checked;
          i++;
          elem = document.getElementById("role" + i);
        } else {
          i++;
          elem = document.getElementById("role" + i);
        }
      }
    }
  }
  
  function toggleBy() {
    var elemByRole = document.getElementById("byrole");
    var elemByUser = document.getElementById("byuser");
    var elemRoles = document.getElementById("roles");
    var elemUsers = document.getElementById("users");
    
    if (elemByUser.checked) {
      elemRoles.style.display = "none";
      elemUsers.style.display = "block";
    } else {
      elemRoles.style.display = "block";
      elemUsers.style.display = "none";
    }
  }
</script>
<%
  Send_HeadEnd()
  If (isTemplate) Then
    Send_BodyBegin(12)
  Else
    Send_BodyBegin(2)
  End If
  
  If (Logix.UserRoles.FavoriteOffersForOthers = False) Then
    Send_Denied(2, "perm.offers-favorite-others")
    GoTo done
  End If
  If (BannersEnabled AndAlso Not Logix.IsAccessibleOffer(AdminUserID, OfferID)) Then
    Send("<script type=""text/javascript"" language=""javascript"">")
    Send("  function ChangeParentDocument() { return true; } ")
    Send("</script>")
    Send_Denied(1, "banners.access-denied-offer")
    Send_BodyEnd()
    GoTo done
  End If
  
  Send("<script type=""text/javascript"">")
  Send("function ChangeParentDocument() { ")
  If (Request.QueryString("save") <> "") Then
    Send("  var elem = null;")
    Send("  var msg = '" & UserCount & " " & Copient.PhraseLib.Lookup("term.of", LanguageID) & " " & TotalUserCount & _
         " " & StrConv(Copient.PhraseLib.Lookup("term.users", LanguageID), VbStrConv.Lowercase) & _
         " " & Copient.PhraseLib.Lookup("offer-gen.havefavorite", LanguageID) & "';")
    Send("  if (opener != null) { ")
    Send("    elem = opener.document.getElementById('favImg');")
    Send("    if (elem != null) { ")
    Send("      elem.setAttribute(""alt"", msg);")
    Send("      elem.setAttribute(""title"", msg);")
    Send("    }")
    Send("  }")
  End If
  Send("    }")
  Send("</script>")
%>
<form action="" id="mainform" name="mainform">
  <input type="hidden" id="OfferID" name="OfferID" value="<% sendb(OfferID) %>" />
  <input type="hidden" id="EngineID" name="EngineID" value="<% Sendb(EngineID) %>" />
  <%
    If CustomerInquiry Then
      Send("<input type=""hidden"" id=""CustomerInquiry"" name=""CustomerInquiry"" value=""1"" />")
    End If
  %>
  <div id="intro">
    <%
      If OfferID > 0 Then
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.favorites", LanguageID), VbStrConv.Lowercase) & "</h1>")
      Else
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.favorites", LanguageID), VbStrConv.Lowercase) & "</h1>")
      End If
    %>
    <div id="controls">
      <%
        If (Logix.UserRoles.FavoriteOffersForOthers) Then
          If (Not bUseTemplateLocks) Then
            Send_Save()
          ElseIf (bUseTemplateLocks And (Disallow_AdvancedOption = False)) Then
            Send_Save()
          End If
        End If
      %>
    </div>
  </div>
  <div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <div id="column">
      <%
        Sendb(Copient.PhraseLib.Lookup("offer-favorites.ManageFavoritesBy", LanguageID) & " ")
        Sendb("<input type=""radio"" id=""byrole"" name=""by"" onclick=""toggleBy();""" & IIf(byRole, " checked=""checked""", "") & " value=""role"" /><label for=""byrole"">" & StrConv(Copient.PhraseLib.Lookup("term.role", LanguageID), VbStrConv.Lowercase) & "</label>")
        Sendb("<input type=""radio"" id=""byuser"" name=""by"" onclick=""toggleBy();""" & IIf(byRole, "", " checked=""checked""") & " value=""user"" /><label for=""byuser"">" & StrConv(Copient.PhraseLib.Lookup("term.user", LanguageID), VbStrConv.Lowercase) & "</label>")
        Send("<br />")
      %>
      <br class="half" />
      <div class="box" id="roles"<%Sendb(IIf(byRole, "", " style=""display:none;"""))%>>
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.roles", LanguageID))%>
          </span>
        </h2>
        <div style="height:450px; overflow:auto;">
          <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.roles", LanguageID))%>">
            <thead>
              <tr>
                <th style="text-align:center;"><input type="checkbox" id="chkAllRoles" name="chkAllRoles" title="Select all items" onclick="handleAllRoles();" /></th>
                <th style="width:35%;"><% Sendb(Copient.PhraseLib.Lookup("term.role", LanguageID))%></th>
                <th style="width:60%;"><% Sendb(Copient.PhraseLib.Lookup("term.users", LanguageID))%></th>
              </tr>
            </thead>
            <tbody>
            <%
              ' Query to get role info
              MyCommon.QueryStr = "select AR.RoleName, AR.DisplayOrder, AR.PhraseID, AUR.RoleID, count(AUR.AdminUserID) as UserCount " & _
                                  "from AdminUserRoles as AUR with (NoLock) " & _
                                  "inner join AdminRoles as AR with (NoLock) on AR.RoleID=AUR.RoleID " & _
                                  "where AdminUserID > 0 " & _
                                  "group by AUR.RoleID, RoleName, DisplayOrder, PhraseID;"
              rst = MyCommon.LRT_Select
              i = 0
              If rst.Rows.Count > 0 Then
                For Each row In rst.Rows
                  If shaded = "" Then
                    shaded = " class=""shaded"""
                  Else
                    shaded = ""
                  End If
                  ' Query to find out about individual users
                  MyCommon.QueryStr = "select AU.AdminUserID, AUR.RoleID, AU.FirstName, AU.LastName, AU.UserName, " & _
                                      "  case when AUO.OfferID is NULL then cast(0 as bit) else cast(1 as bit) end as Favorite, " & _
                                      "AUO.FavoredBy, AUO.FavoredDate, AUX.FirstName as FavorerFirstName, " & _
                                      "AUX.LastName as FavorerLastName, AUX.UserName as FavorerUserName " & _
                                      "from AdminUsers as AU " & _
                                      "left join AdminUserOffers as AUO on AUO.AdminUserID=AU.AdminUserID and AUO.OfferID=" & OfferID & " " & _
                                      "left join AdminUsers as AUX on AUX.AdminUserID=AUO.FavoredBy " & _
                                      "left join AdminUserRoles as AUR on AUR.AdminUserID=AU.AdminUserID " & _
                                      "where RoleID=" & MyCommon.NZ(row.Item("RoleID"), 0) & " " & _
                                      "order by LastName, FirstName;"
                  rst2 = MyCommon.LRT_Select
                  FavoriteCount = 0
                  For Each row2 In rst2.Rows
                    If row2.Item("Favorite") = True Then
                      FavoriteCount += 1
                    End If
                  Next
                  Send("<tr" & shaded & ">")
                  Send("  <td style=""text-align:center;"">")
                  Send("    <input type=""checkbox"" id=""role" & MyCommon.NZ(row.Item("RoleID"), 0) & """ name=""role" & MyCommon.NZ(row.Item("RoleID"), 0) & """" & IIf(FavoriteCount = MyCommon.NZ(row.Item("UserCount"), 0), " checked=""checked""", "") & " />")
                  Send("  </td>")
                  Send("  <td>")
                  Send("    <label for=""role" & MyCommon.NZ(row.Item("RoleID"), "") & """>" & MyCommon.NZ(row.Item("RoleName"), "") & "</label>")
                  Send("  </td>")
                  Send("  <td>")
                  Send("    " & Copient.PhraseLib.Detokenize("offer-favorite.UserCount", LanguageID, FavoriteCount, MyCommon.NZ(row.Item("UserCount"), 0)))
                  Send("  </td>")
                  Send("</tr>")
                  i += 1
                Next
                shaded = ""
              End If
            %>
            </tbody>
          </table>
        </div>
        <hr class="hidden" />
      </div>
      
      <div class="box" id="users"<%Sendb(IIf(byRole, " style=""display:none;""", ""))%>>
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.users", LanguageID))%>
          </span>
        </h2>
        <div style="height:450px; overflow:auto;">
          <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.roles", LanguageID))%>">
            <thead>
              <tr>
                <th style="text-align:center;"><input type="checkbox" id="chkAllUsers" name="chkAllUsers" title="Select all users" onclick="handleAllUsers();" /></th>
                <th style="width:35%;"><% Sendb(Copient.PhraseLib.Lookup("term.user", LanguageID))%></th>
                <th style="width:35%;"><% Sendb(Copient.PhraseLib.Lookup("term.favoredby", LanguageID))%></th>
                <th style="width:25%;"><% Sendb(Copient.PhraseLib.Lookup("term.date", LanguageID))%></th>
              </tr>
            </thead>
            <tbody>
            <%
              ' Query to get user info
              MyCommon.QueryStr = "select distinct AU.AdminUserID, AU.FirstName, AU.LastName, AU.UserName, " & _
                                  "  case when AUO.OfferID is NULL then cast(0 as bit) else cast(1 as bit) end as Favorite, " & _
                                  "AUO.FavoredBy, AUO.FavoredDate, AUX.FirstName as FavorerFirstName, " & _
                                  "AUX.LastName as FavorerLastName, AUX.UserName as FavorerUserName " & _
                                  "from AdminUsers as AU " & _
                                  "left join AdminUserOffers as AUO on AUO.AdminUserID=AU.AdminUserID and AUO.OfferID=" & OfferID & " " & _
                                  "left join AdminUsers as AUX on AUX.AdminUserID=AUO.FavoredBy " & _
                                  "left join AdminUserRoles as AUR on AUR.AdminUserID=AU.AdminUserID " & _
                                  "order by LastName, FirstName;"
              rst = MyCommon.LRT_Select
              If rst.Rows.Count > 0 Then
                For Each row In rst.Rows
                  If shaded = "" Then
                    shaded = " class=""shaded"""
                  Else
                    shaded = ""
                  End If
                  Send("    <tr" & shaded & ">")
                  Send("      <td style=""text-align:center;""><input type=""checkbox"" id=""user" & MyCommon.NZ(row.Item("AdminUserID"), "") & """ name=""user" & MyCommon.NZ(row.Item("AdminUserID"), "") & """" & IIf(row.Item("Favorite"), " checked=""checked""", "") & " /></td>")
                  If (MyCommon.NZ(row.Item("FirstName"), "") = "" AndAlso MyCommon.NZ(row.Item("LastName"), "") = "") Then
                    Send("      <td><label for=""user" & MyCommon.NZ(row.Item("AdminUserID"), "") & """><i>" & MyCommon.NZ(row.Item("UserName"), "") & "</i></label></td>")
                  Else
                    Send("      <td><label for=""user" & MyCommon.NZ(row.Item("AdminUserID"), "") & """>" & MyCommon.NZ(row.Item("FirstName"), "") & " " & MyCommon.NZ(row.Item("LastName"), "") & "</label></td>")
                  End If
                  If (MyCommon.NZ(row.Item("FavorerFirstName"), "") = "" AndAlso MyCommon.NZ(row.Item("FavorerLastName"), "") = "") Then
                    Send("      <td>" & MyCommon.NZ(row.Item("FavorerUserName"), "") & "</td>")
                  Else
                    Send("      <td>" & MyCommon.NZ(row.Item("FavorerFirstName"), "") & " " & MyCommon.NZ(row.Item("FavorerLastName"), "") & "</td>")
                  End If
                  Send("      <td>" & MyCommon.NZ(row.Item("FavoredDate"), "") & "</td>")
                  Send("    </tr>")
                Next
              End If
            %>
            </tbody>
          </table>
        </div>
      </div>
      
    </div>
    
  </div>
</form>

<script type="text/javascript">
<% If (CloseAfterSave) Then %>
    window.close();
<% End If %>
</script>

<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd("", "")
  MyCommon = Nothing
  Logix = Nothing
%>
