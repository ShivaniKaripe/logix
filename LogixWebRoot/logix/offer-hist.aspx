<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="CMS.AMS.Models" %>
<%@ Import Namespace="CMS.AMS" %>
<%@ Import Namespace="CMS.AMS.Contract" %>
<%
  ' *****************************************************************************
  ' * FILENAME: offer-hist.aspx 
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
  Dim dst As System.Data.DataTable
  Dim row As System.Data.DataRow
  Dim rst2 As DataTable
  Dim rst3 As DataTable
  Dim OfferID As Long
  Dim GName As String
  Dim Deleted As Boolean = False
  Dim sizeOfData As Integer
  Dim DefaultIDType As Integer
  Dim i As Integer = 0
  Dim maxEntries As Integer = 9999
  Dim Shaded As String = "shaded"
  Dim IsTemplate As Boolean = False
  Dim FromTemplate As Boolean = False
  Dim EngineID As Integer
  Dim EngineSubTypeID As Integer = 0
  Dim infoMessage As String = ""
  Dim modMessage As String = ""
  Dim Handheld As Boolean = False
  Dim BannersEnabled As Boolean = True
  Dim StatusCode As Copient.LogixInc.STATUS_FLAGS = Copient.LogixInc.STATUS_FLAGS.STATUS_UNKNOWN
  Dim StatusText As String = ""
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "offer-hist.aspx"
  CurrentRequest.Resolver.AppName = MyCommon.AppName
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  
  DefaultIDType = MyCommon.Fetch_SystemOption(30)
  OfferID = Request.QueryString("OfferID")
  GName = Request.QueryString("GroupName")
  
  ' Check in case it was a POST instead of get
  If (OfferID = 0 And Not Request.QueryString("save") <> "") Then
    OfferID = Request.Form("OfferID")
    GName = Request.Form("Name")
  End If
  
  If (OfferID = 0) Then
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "offer-gen.aspx?new=New")
  End If
  
    Dim isTranslatedOffer As Boolean =MyCommon.IsTranslatedUEOffer(OfferID,  MyCommon)
    Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249)="1",True,False)
    
    Dim bEnableAdditionalLockoutRestrictionsOnOffers As Boolean = IIf(MyCommon.Fetch_SystemOption(260) = "1", True, False)
    Dim bOfferEditable As Boolean = MyCommon.IsOfferEditablePastLockOutPeriod(Logix.UserRoles.EditOfferPastLockoutPeriod, MyCommon, OfferID)
    
  MyCommon.QueryStr = "select EngineID from OfferIDs with (NoLock) where OfferID=" & OfferID & ";"
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), 0)
  End If
  
  'Determine the engine subtype
  If EngineID = 2 OrElse EngineID = 9 Then
    MyCommon.QueryStr = "select EngineSubTypeID from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID & ";"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
      EngineSubTypeID = MyCommon.NZ(rst.Rows(0).Item("EngineSubTypeID"), 0)
    End If
  End If
  
  If (EngineID = Copient.CommonInc.InstalledEngines.CPE OrElse EngineID = Copient.CommonInc.InstalledEngines.UE) Then
    MyCommon.QueryStr = "Select IncentiveName as Name, IsTemplate, FromTemplate, Deleted, DeployDeferred,buy.ExternalBuyerId as BuyerID from CPE_Incentives CPE with (NoLock) " & _
                        "left outer join Buyers as buy with (nolock) on buy.BuyerId= CPE.BuyerId " & _
                        " where IncentiveID=" & OfferID & ";"
  ElseIf (EngineID = Copient.CommonInc.InstalledEngines.Website OrElse EngineID = 5 OrElse EngineID = 6) Then
    MyCommon.QueryStr = "Select IncentiveName as Name, IsTemplate, FromTemplate, Deleted, DeployDeferred,buy.ExternalBuyerId as BuyerID from CPE_Incentives CPE with (NoLock) " & _
                        "left outer join Buyers as buy with (nolock) on buy.BuyerId= CPE.BuyerId " & _
                        " where IncentiveID=" & OfferID & ";"
  Else
    MyCommon.QueryStr = "select Name, IsTemplate, FromTemplate, ProdStartDate, ProdEndDate, StatusFlag, DeployDeferred, Deleted,NULL as BuyerID from Offers with (NoLock) where OfferID=" & OfferID & ";"
  End If
  
  rst = MyCommon.LRT_Select()
  If (rst.Rows.Count > 0) Then
    For Each row In rst.Rows
      If (MyCommon.Fetch_UE_SystemOption(168) = "1" AndAlso MyCommon.NZ(row.Item("BuyerID"), "") <> "") Then
        GName = "Buyer " + row.Item("BuyerID").ToString() + " -" + MyCommon.NZ(row.Item("Name"), "").ToString()
      Else
        GName = MyCommon.NZ(row.Item("Name"), Copient.PhraseLib.Lookup("term.unknown", LanguageID))
      End If
      'GName = MyCommon.NZ(row.Item("Name"), "")
      IsTemplate = row.Item("IsTemplate")
      FromTemplate = row.Item("FromTemplate")
      If (row.Item("Deleted") = "1") Then
        Deleted = True
      End If
    Next
  Else
    Deleted = True
  End If
  
  If (EngineID = Engines.UE) Then
    Dim IsCollisionDetectionEnabled As Boolean = False
    Dim CollisionDetectionEnabledResp As AMSResult(Of Boolean)
    Dim m_Offer As IOffer = CurrentRequest.Resolver.Resolve(Of IOffer)()
    Dim m_CollisionDetectionService As ICollisionDetectionService = CurrentRequest.Resolver.Resolve(Of ICollisionDetectionService)()
    CollisionDetectionEnabledResp = m_Offer.IsCollisionDetectionEnabled(Engines.UE, OfferID)
    If (CollisionDetectionEnabledResp.ResultType = AMSResultType.Success AndAlso CollisionDetectionEnabledResp.Result = True) Then IsCollisionDetectionEnabled = True
    
    If IsCollisionDetectionEnabled Then
      Dim AwaitingDetectionResp As AMSResult(Of Models.OCD.QueueStatus) = m_CollisionDetectionService.GetOfferQueueStatus(OfferID)
      If (AwaitingDetectionResp.ResultType = AMSResultType.Success AndAlso (AwaitingDetectionResp.Result = OCD.QueueStatus.NotStarted OrElse AwaitingDetectionResp.Result = OCD.QueueStatus.InProgress)) Then
        OfferLockedforCollisionDetection = True
      End If
    End If
  End If

  
  Send_HeadBegin("term.offer", "term.history", OfferID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
  Send_HeadEnd()
  If (IsTemplate) Then
    Send_BodyBegin(11)
  Else
    Send_BodyBegin(1)
  End If
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  Send_Tabs(Logix, 2)
  
  If EngineID = 2 Then
    If Deleted Then
      Send_Subtabs(Logix, 26, 9, , OfferID)
      infoMessage = Copient.PhraseLib.Lookup("error.deleted", LanguageID)
    Else
      If (IsTemplate) Then
        Send_Subtabs(Logix, 25, 9, , OfferID)
      Else
        Send_Subtabs(Logix, 24, 9, , OfferID)
      End If
    End If
  ElseIf EngineID = 3 Then
    If Deleted Then
      Send_Subtabs(Logix, 29, 8, , OfferID)
      infoMessage = Copient.PhraseLib.Lookup("error.deleted", LanguageID)
    Else
      Send_Subtabs(Logix, 27, 8, , OfferID)
    End If
    'ElseIf EngineID = 4 Then
    '  If Deleted Then
    '    Send_Subtabs(Logix, 202, 8, , OfferID)
    '    infoMessage = Copient.PhraseLib.Lookup("error.deleted", LanguageID)
    '  Else
    '    If (IsTemplate) Then
    '      Send_Subtabs(Logix, 201, 8, , OfferID)
    '    Else
    '      Send_Subtabs(Logix, 200, 8, , OfferID)
    '    End If
    '  End If
  ElseIf EngineID = 5 Then
    If Deleted Then
      Send_Subtabs(Logix, 204, 8, , OfferID)
      infoMessage = Copient.PhraseLib.Lookup("error.deleted", LanguageID)
    Else
      Send_Subtabs(Logix, 203, 8, , OfferID)
    End If
  ElseIf EngineID = 6 Then
    If Deleted Then
      Send_Subtabs(Logix, 206, 9, , OfferID)
      infoMessage = Copient.PhraseLib.Lookup("error.deleted", LanguageID)
    Else
      If (IsTemplate) Then
        Send_Subtabs(Logix, 205, 9, , OfferID)
      Else
        Send_Subtabs(Logix, 205, 9, , OfferID)
      End If
    End If
  ElseIf EngineID = 9 Then
    If Deleted Then
      Send_Subtabs(Logix, 210, 9, , OfferID)
      infoMessage = Copient.PhraseLib.Lookup("error.deleted", LanguageID)
    Else
      If (IsTemplate) Then
        Send_Subtabs(Logix, 209, 9, , OfferID)
      Else
        Send_Subtabs(Logix, 208, 9, , OfferID)
      End If
    End If
  Else
    If Deleted Then
      Send_Subtabs(Logix, 23, 8, , OfferID)
      infoMessage = Copient.PhraseLib.Lookup("error.deleted", LanguageID)
    Else
      If (IsTemplate) Then
        Send_Subtabs(Logix, 22, 8, , OfferID)
      Else
        Send_Subtabs(Logix, 21, 8, , OfferID)
      End If
    End If
  End If
  
  If (Logix.UserRoles.AccessOffers = False AndAlso Not IsTemplate) Then
    Send_Denied(1, "perm.offers-access")
    GoTo done
  End If
  If (Logix.UserRoles.AccessTemplates = False AndAlso IsTemplate) Then
    Send_Denied(1, "perm.offers-access-templates")
    GoTo done
  End If
  If (Logix.UserRoles.ViewHistory = False) Then
    Send_Denied(1, "perm.admin-history")
    GoTo done
  End If
  If (Logix.UserRoles.AccessInstantWinOffers = False AndAlso EngineID = 2 AndAlso EngineSubTypeID = 1) Then
    Send_Denied(1, "perm.offers-accessinstantwin")
    GoTo done
  End If
  If (BannersEnabled AndAlso Not Logix.IsAccessibleOffer(AdminUserID, OfferID)) Then
    Send_Denied(1, "banners.access-denied-offer")
    Send_BodyEnd()
    GoTo done
  End If
%>
<%
  Send("<div id=""intro"">")
  If (IsTemplate) Then
    Sendb("  <h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID)
    If (GName <> "") Then
      Sendb(": " & MyCommon.TruncateString(GName, 50))
    End If
    Send("</h1>")
    Send("<div id=""controls"">")
    If MyCommon.Fetch_SystemOption(75) AndAlso (Deleted = False) Then
            If (OfferID > 0 And Logix.UserRoles.AccessNotes AndAlso (Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse  bOfferEditable)) Then
        Send_NotesButton(3, OfferID, AdminUserID)
      End If
    End If
    Send("</div>")
  Else
    Sendb("  <h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID)
    If (GName <> "") Then
      Sendb(": " & MyCommon.TruncateString(GName, 50))
    End If
    Send("</h1>")
    Send("<div id=""controls"">")
    If (MyCommon.Fetch_SystemOption(75)) Then
            If (OfferID > 0 AndAlso Logix.UserRoles.AccessNotes AndAlso (Deleted = False) AndAlso (Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse  bOfferEditable)) Then
        Send_NotesButton(3, OfferID, AdminUserID)
      End If
    End If
    Send("</div>")
  End If
  Send("</div>")
%>
<div id="main">
  <%
    Select Case EngineID
      Case 2, 3, 5, 6, 9
        MyCommon.QueryStr = "select StatusFlag from CPE_Incentives where IncentiveID=" & OfferID & ";"
      Case Else
        MyCommon.QueryStr = "select StatusFlag from Offers where OfferID=" & OfferID & ";"
    End Select

    rst2 = MyCommon.LRT_Select
    StatusText = Logix.GetOfferStatus(OfferID, LanguageID, StatusCode)
    If Not IsTemplate Then
      If (rst2.Rows.Count > 0 AndAlso MyCommon.NZ(rst2.Rows(0).Item("StatusFlag"), 0) <> 2) Then
                If (StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_ACTIVE) AndAlso (MyCommon.NZ(rst2.Rows(0).Item("StatusFlag"), 0) > 0) Then
          If (MyCommon.NZ(rst.Rows(0).Item("DeployDeferred"), False) = False) Then
            modMessage = Copient.PhraseLib.Lookup("alert.modpostdeploy", LanguageID)
            Send("<div id=""modbar"">" & modMessage & "</div>")
          End If
        End If
      End If
    End If
    
    If Not Deleted Then
      ' Send the status bar, but not if it's a new offer or a template, or if there's already a modMessage being shown.
      If (Not IsTemplate) AndAlso (rst.Rows.Count > 0) AndAlso (modMessage = "") Then
        If (EngineID = 2) OrElse (EngineID = 3) OrElse (EngineID = 5) OrElse (EngineID = 6) OrElse (EngineID = 9) Then
          MyCommon.QueryStr = "select incentiveId from CPE_Incentives with (NoLock) where CreatedDate = LastUpdate and IncentiveID=" & OfferID
          rst3 = MyCommon.LRT_Select
          If (rst3.Rows.Count = 0) Then
            Send_Status(OfferID, 2)
          End If
        Else
          MyCommon.QueryStr = "select OfferID from Offers with (NoLock) where CreatedDate = LastUpdate and OfferID=" & OfferID
          rst3 = MyCommon.LRT_Select
          If (rst3.Rows.Count = 0) Then
            Send_Status(OfferID)
          End If
        End If
      End If
    End If
  %>
  <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
  <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.history", LanguageID)) %>">
    <thead>
      <tr>
        <th align="left" scope="col" class="th-timedate">
          <% Sendb(Copient.PhraseLib.Lookup("term.timedate", LanguageID))%>
        </th>
        <th align="left" scope="col" class="th-user">
          <% Sendb(Copient.PhraseLib.Lookup("term.user", LanguageID))%>
        </th>
        <th align="left" scope="col" class="th-user">
          <% Sendb(Copient.PhraseLib.Lookup("term.buyer", LanguageID))%>
        </th>
        <th align="left" scope="col" class="th-action">
          <% Sendb(Copient.PhraseLib.Lookup("term.action", LanguageID))%>
        </th>
      </tr>
    </thead>
    <tbody>
      <%
        MyCommon.QueryStr = "select AU.FirstName, AU.LastName, AL.ActivityDate, AL.Description,B.ExternalBuyerId from ActivityLog as AL with (NoLock) left join AdminUsers as AU with (NoLock) on AL.AdminID=AU.AdminUserID left join Buyers as B with (NoLock) on AL.BuyerID=B.BuyerID Where ActivityTypeID='3' and LinkID='" & OfferID & "' order by ActivityDate desc, ActivityID desc;"
        dst = MyCommon.LRT_Select
        sizeOfData = dst.Rows.Count
        While (i < sizeOfData And i < maxEntries)
          Send("<tr class=""" & Shaded & """>")
          If (Not IsDBNull(dst.Rows(i).Item("ActivityDate"))) Then
            Send("  <td>" & Logix.ToShortDateTimeString(dst.Rows(i).Item("ActivityDate"), MyCommon) & "</td>")
          Else
            Send("  <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
          End If
          Send("  <td>" & MyCommon.NZ(dst.Rows(i).Item("FirstName"), "") & " " & MyCommon.NZ(dst.Rows(i).Item("LastName"), "") & "</td>")
          Send("  <td>" & MyCommon.NZ(dst.Rows(i).Item("ExternalBuyerId"), "") & "</td>")
          Send("  <td>" & MyCommon.NZ(dst.Rows(i).Item("Description"), "") & "</td>")
          Send("</tr>")
          If Shaded = "shaded" Then
            Shaded = ""
          Else
            Shaded = "shaded"
          End If
          i = i + 1
        End While
        If (sizeOfData = 0) Then
          Send("<tr>")
          Send("  <td colspan=""3""></td>")
          Send("</tr>")
        End If
      %>
    </tbody>
  </table>
</div>
<%
  If MyCommon.Fetch_SystemOption(75) AndAlso (Deleted = False) Then
        If (OfferID > 0 And Logix.UserRoles.AccessNotes AndAlso (Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer)AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse  bOfferEditable)) Then
      Send_Notes(3, OfferID, AdminUserID)
    End If
  End If
done:
  Send_BodyEnd()
  MyCommon.Close_LogixRT()
  MyCommon = Nothing
  Logix = Nothing
%>
