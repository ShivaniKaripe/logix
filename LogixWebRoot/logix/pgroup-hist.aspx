<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
    ' *****************************************************************************
    ' * FILENAME: pgroup-hist.aspx 
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
    Dim ProductGroupID As Long
    Dim GName As String
    Dim Deleted As Boolean = False
    Dim sizeOfData As Integer
    Dim DefaultIDType As Integer
    Dim i As Integer = 0
    Dim maxEntries As Integer = 9999
    Dim Shaded As String = "shaded"
    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False
  
    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If
  
    Response.Expires = 0
    MyCommon.AppName = "pgroup-hist.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
    DefaultIDType = MyCommon.Fetch_SystemOption(30)
    ProductGroupID = Request.QueryString("ProductGroupID")
    GName = Request.QueryString("GroupName")
    ' Check in case it was a POST instead of get
    If (ProductGroupID = 0 And Not Request.QueryString("save") <> "") Then
        ProductGroupID = Request.Form("ProductGroupID")
        GName = Request.Form("GroupName")
    End If
  
    If (ProductGroupID = 0) Then
        Response.Status = "301 Moved Permanently"
        Response.AddHeader("Location", "pgroup-edit.aspx?new=New")
    End If
  
    MyCommon.QueryStr = "select Name, Deleted from ProductGroups with (NoLock) where ProductGroupID=" & ProductGroupID & ";"
    rst = MyCommon.LRT_Select()
    If (rst.Rows.Count = 0) Then
        infoMessage = Copient.PhraseLib.Lookup("error.deleted", LanguageID)
        Deleted = True
    Else
        GName = MyCommon.NZ(rst.Rows(0).Item("Name"), "")
        If (MyCommon.NZ(rst.Rows(0).Item("Deleted"), "1") = "1") Then
            infoMessage = Copient.PhraseLib.Lookup("error.deleted", LanguageID)
            Deleted = True
        End If
    End If
  
    Send_HeadBegin("term.productgroup", "term.history", ProductGroupID)
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts()
    Send_HeadEnd()
    Send_BodyBegin(1)
    Send_Bar(Handheld)
    Send_Help(CopientFileName)
    Send_Logos()
    Send_Tabs(Logix, 4)
    Send_Subtabs(Logix, 41, 5, , ProductGroupID)
  
    If (Logix.UserRoles.AccessProductGroups = False) Then
        Send_Denied(1, "perm.pgroup-access")
        GoTo done
    ElseIf (MyCommon.IsEngineInstalled(9) AndAlso ProductGroupID > 0) Then
        MyCommon.QueryStr = "dbo.pa_IsProductGroupAccessibleToBuyer"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt).Value = ProductGroupID
        MyCommon.LRTsp.Parameters.Add("@checkPGlistBasedOnBuyerID", SqlDbType.Bit).Value = (Logix.UserRoles.ViewProductgroupRegardlessBuyer = False)
        MyCommon.LRTsp.Parameters.Add("@adminuserid", SqlDbType.Int).Value = AdminUserID
        Dim IsPGAccessible As Boolean = (MyCommon.LRTsp_select().Rows.Count > 0)
        MyCommon.Close_LRTsp()
        If (IsPGAccessible = False) Then
            Send_Denied("1", "perm.viewProductgroupsRegardlessBuyer")
            GoTo done
        End If
    End If
    If (Logix.UserRoles.ViewHistory = False) Then
        Send_Denied(1, "perm.admin-history")
        GoTo done
    End If
%>
<div id="intro">
    <h1 id="title">
        <%
            Sendb(Copient.PhraseLib.Lookup("term.productgroup", LanguageID) & " #" & ProductGroupID)
            If GName <> "" Then
                Sendb(": " & GName)
            End If
        %>
    </h1>
    <div id="controls">
        <%
            If MyCommon.Fetch_SystemOption(75) AndAlso (Deleted = False) Then
                If (ProductGroupID > 0 And Logix.UserRoles.AccessNotes) Then
                    Send_NotesButton(7, ProductGroupID, AdminUserID)
                End If
            End If
        %>
    </div>
</div>
<div id="main">
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
                MyCommon.QueryStr = "select AU.FirstName, AU.LastName, B.ExternalBuyerId, AL.ActivityDate, AL.Description from ActivityLog as AL with (NoLock) " & _
                    " left join AdminUsers as AU with (NoLock) on AL.AdminID=AU.AdminUserID left join Buyers as B with (NoLock) on AL.BuyerID=B.BuyerId Where ActivityTypeID='5' and LinkID='" & ProductGroupID & "' order by ActivityDate desc;"
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
        If (ProductGroupID > 0 And Logix.UserRoles.AccessNotes) Then
            Send_Notes(7, ProductGroupID, AdminUserID)
        End If
    End If
done:
    Send_BodyEnd()
    MyCommon.Close_LogixRT()
    Logix = Nothing
    MyCommon = Nothing
%>
