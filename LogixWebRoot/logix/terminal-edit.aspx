<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.CommonInc" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%
  ' *****************************************************************************
  ' * FILENAME: terminal-edit.aspx 
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
  Dim TerminalId As Long
  Dim TerminalName As String
  Dim TerminalDescription As String
  Dim ExtTerminalCode As String
  Dim LastUpdate As String
  Dim EngineType As Integer
  Dim EngineName As String
  Dim Deleted As Boolean = False
  Dim LayoutID As Integer
  Dim SpecificPromos As String
  Dim SpecificPromosChecked As String
  Dim LayoutsDisplay As String = "none"
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim dst As DataTable
  Dim rst As DataTable
  Dim rst2 As DataTable
  Dim row As DataRow
  Dim row2 As DataRow
  Dim bSave As Boolean
  Dim bDelete As Boolean
  Dim bCreate As Boolean
  Dim FocusField As String = "ExtTerminalCode"
  Dim FuelProcessing As Boolean
  Dim FuelProcessingChecked As String = ""
  Dim TerminalNameTitle As String = ""
  Dim SizeOfData As Integer
  Dim ShowActionButton As Boolean = False
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim BannerID As Integer = 0
  Dim BannerName As String = ""
  Dim BannersEnabled As Boolean = False
  Dim BannerIDs As String = ""
  Dim BannerEngines As String = ""
  Dim DefaultEngineID As Integer = 0
  Dim disabledattribute As String = ""
  Dim TerminalLockingGroupID As Long
  Dim OldTerminalLockingGroupID As Long
  Dim TerminalLockingGroupName As String
  Dim CPEInstalled As Boolean = False
  Dim CMInstalled As Boolean = False
  Dim MultiEngine As Boolean = False
  Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249)="1",True,False)
  Dim conditionalQuery = String.Empty
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "terminal-edit.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
    
    '******************************* BEGIN AJAX SECTION ******************************* 
    'Check if its a ajax call
    If GetCgiValue("MI") <> Nothing Then
        Dim ReturnString As String = ""
        Dim TempconditionalQuery As String = ""
        Dim UE_SystemOption_168 As String = MyCommon.Fetch_UE_SystemOption(168)
        
        If GetCgiValue("mi") = "ASSOCIATED_OFFERS_LIST" Then
                        
            If (bEnableRestrictedAccessToUEOfferBuilder) Then
                TempconditionalQuery = GetRestrictedAccessToUEBuilderQuery(MyCommon, Logix, "I")
            End If
              
            If (GetCgiValue("TerminalID") <> "") Then
                
                MyCommon.QueryStr = "pt_TerminalAssociatedOfferList"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@AdminUserId", SqlDbType.Int).Value = AdminUserID
                MyCommon.LRTsp.Parameters.Add("@TerminalTypeid", SqlDbType.Int).Value = GetCgiValue("TerminalID")
                MyCommon.LRTsp.Parameters.Add("@PageStart", SqlDbType.Int).Value = GetCgiValue("PageStart")
                MyCommon.LRTsp.Parameters.Add("@PageEnd", SqlDbType.Int).Value = GetCgiValue("PageEnd")
                MyCommon.LRTsp.Parameters.Add("@CreateUEOffer", SqlDbType.Bit).Value = IIf(Logix.UserRoles.CreateUEOffers = True, 1, 0)
                MyCommon.LRTsp.Parameters.Add("@AccessTranslatedUEOffers", SqlDbType.Bit).Value = IIf(Logix.UserRoles.AccessTranslatedUEOffers = True, 1, 0)
                MyCommon.LRTsp.Parameters.Add("@EnableRestrictedAccessToUEOfferBuilder", SqlDbType.Bit).Value = IIf(bEnableRestrictedAccessToUEOfferBuilder = True, 1, 0)
                
                rst = MyCommon.LRTsp_select()
                MyCommon.Close_LRTsp()
                
                MyCommon.ConvertDataTabletoJson(ReturnString, rst)
                ReturnString = "T_AMS_SPLITTER_AMS_" + ReturnString
                
            End If
            Response.Write(ReturnString)
            Response.End()
            Return
        End If
        
    End If
       
    
    '******************************* END AJAX SECTION   ******************************* 
  
  Try
    ' fill in if it was a get method
    If Request.RequestType = "GET" Then
      TerminalId = Request.QueryString("TerminalId")
      TerminalName = Logix.TrimAll(Request.QueryString("TerminalName"))
      TerminalDescription = Logix.TrimAll(Request.QueryString("TerminalDescription"))
                                
      ExtTerminalCode = Logix.TrimAll(Request.QueryString("ExtTerminalCode"))
      EngineType = Request.QueryString("EngineID")
      LayoutID = MyCommon.Extract_Val(Request.QueryString("layouts"))
      SpecificPromos = MyCommon.Extract_Val(Request.QueryString("specificpromos"))
      FuelProcessing = IIf(Request.QueryString("fuelprocessing") = "1", 1, 0)
      BannerID = MyCommon.Extract_Val(Request.QueryString("BannerID"))
      TerminalLockingGroupID = MyCommon.Extract_Val(Request.QueryString("TerminalLockingGroupId"))
      OldTerminalLockingGroupID = MyCommon.Extract_Val(Request.QueryString("OldTerminalLockingGroupId"))
      
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
      TerminalId = Request.Form("TerminalId")
      If TerminalId = 0 Then
        TerminalId = MyCommon.Extract_Val(Request.QueryString("TerminalId"))
      End If
      TerminalName = Request.Form("TerminalName")
      TerminalDescription = Request.Form("TerminalDescription")
      ExtTerminalCode = Request.Form("ExtTerminalCode")
      EngineType = Request.Form("EngineID")
      LayoutID = MyCommon.Extract_Val(Request.Form("layouts"))
      SpecificPromos = MyCommon.Extract_Val(Request.Form("specificpromos"))
      FuelProcessing = IIf(Request.Form("fuelprocessing") = "1", 1, 0)
      BannerID = MyCommon.Extract_Val(Request.Form("BannerID"))
      TerminalLockingGroupID = MyCommon.Extract_Val(Request.Form("TerminalLockingGroupId"))
      OldTerminalLockingGroupID = MyCommon.Extract_Val(Request.Form("OldTerminalLockingGroupId"))
      
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
    
    If (SpecificPromos > 0) Then SpecificPromosChecked = " checked=""checked"""
    
    
    BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
    
    Send_HeadBegin("term.terminal", , TerminalId)
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
 
    Send_Scripts()
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
      Response.Redirect("terminal-edit.aspx")
    End If
    
    If MyCommon.IsEngineInstalled(2) Then CPEInstalled = True
    If MyCommon.IsEngineInstalled(0) Then CMInstalled = True
    If CPEInstalled AndAlso CMInstalled Then MultiEngine = True
    
    MyCommon.QueryStr = "select EngineID from PromoEngines PE with (NoLock) where Installed=1 and DefaultEngine=1;"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      DefaultEngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), 0)
    End If
    
    If (TerminalId = 0) Then
      ' get the engine id for the banner (if necessary)
      If (BannersEnabled AndAlso BannerID > 0) Then
        MyCommon.QueryStr = "select EngineID from BannerEngines BE with (NoLock) where BannerID = " & BannerID
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
          EngineType = MyCommon.NZ(rst.Rows(0).Item("EngineID"), 0)
        End If
      ElseIf (BannersEnabled AndAlso BannerID = 0) Then
        MyCommon.QueryStr = "select EngineID from PromoEngines PE with (NoLock) where Installed=1 and DefaultEngine=1;"
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
          EngineType = MyCommon.NZ(rst.Rows(0).Item("EngineID"), 0)
        End If
      End If
    End If
    
    If bSave Then
      If (TerminalId = 0) Then
              
        MyCommon.QueryStr = "dbo.pt_TerminalTypes_Insert"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 100).Value = TerminalName
        MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 1000).Value = TerminalDescription
        MyCommon.LRTsp.Parameters.Add("@ExtTerminalCode", SqlDbType.NVarChar, 50).Value = ExtTerminalCode
        MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = EngineType
        MyCommon.LRTsp.Parameters.Add("@LayoutID", SqlDbType.Int).Value = LayoutID
        MyCommon.LRTsp.Parameters.Add("@SpecificPromosOnly", SqlDbType.Int).Value = SpecificPromos
        MyCommon.LRTsp.Parameters.Add("@FuelProcessing", SqlDbType.Bit).Value = IIf(FuelProcessing, 1, 0)
        MyCommon.LRTsp.Parameters.Add("@BannerID", SqlDbType.Int).Value = BannerID
        MyCommon.LRTsp.Parameters.Add("@TerminalTypeId", SqlDbType.Int).Direction = ParameterDirection.Output
        'TerminalName = MyCommon.Parse_Quotes(TerminalName)
        If ExtTerminalCode = "" And (EngineType = InstalledEngines.CM OrElse EngineType = InstalledEngines.Catalina) Then
          infoMessage = Copient.PhraseLib.Lookup("terminal-edit.nonameandcode", LanguageID)
        ElseIf (TerminalName = "") Then
          infoMessage = Copient.PhraseLib.Lookup("terminal-edit.noname", LanguageID)
        ElseIf Not (EngineType = InstalledEngines.CM or EngineType = InstalledEngines.UE OrElse EngineType = InstalledEngines.Catalina) AndAlso IsNumeric(ExtTerminalCode) Then
          infoMessage = Copient.PhraseLib.Lookup("terminal-edit.badcode", LanguageID)
        ElseIf (EngineType = InstalledEngines.CM OrElse EngineType = InstalledEngines.Catalina) AndAlso (Not IsNumeric(ExtTerminalCode) OrElse ((ExtTerminalCode < 1) OrElse (Int(ExtTerminalCode) <> ExtTerminalCode))) Then
          infoMessage = Copient.PhraseLib.Lookup("terminal-edit.badcode", LanguageID)
        Else
          MyCommon.QueryStr = "SELECT TerminalTypeID FROM TerminalTypes with (NoLock) WHERE Name = '" & MyCommon.Parse_Quotes(TerminalName) & "' AND EngineID=" & EngineType & " AND AnyTerminal=0 AND Deleted=0 "
          If (BannersEnabled) Then
            MyCommon.QueryStr &= " and BannerID=" & BannerID
          End If
          rst = MyCommon.LRT_Select
          If (rst.Rows.Count > 0) Then
            infoMessage = Copient.PhraseLib.Lookup("terminal-edit.nameused", LanguageID)
          Else
            MyCommon.QueryStr = "SELECT TerminalTypeID FROM TerminalTypes with (NoLock) WHERE EngineID In (0,1,4) AND AnyTerminal=0 AND Deleted=0 AND ExtTerminalCode = '" & ExtTerminalCode & "'"
            If (BannersEnabled) Then
              MyCommon.QueryStr &= " and BannerID=" & BannerID
            End If
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count > 0) And Not (EngineType = InstalledEngines.CPE) Then
              infoMessage = Copient.PhraseLib.Lookup("terminal-edit.codeused", LanguageID)
            Else
              MyCommon.LRTsp.ExecuteNonQuery()
              TerminalId = MyCommon.LRTsp.Parameters("@TerminalTypeId").Value
              MyCommon.Close_LRTsp()
              MyCommon.Activity_Log(21, TerminalId, AdminUserID, Copient.PhraseLib.Lookup("history.terminal-create", LanguageID))
            End If
          End If
        End If
      Else
        MyCommon.QueryStr = "dbo.pt_TerminalTypes_Update"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@TerminalTypeId", SqlDbType.Int).Value = TerminalId
        MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 100).Value = TerminalName
        MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 1000).Value = TerminalDescription
        MyCommon.LRTsp.Parameters.Add("@ExtTerminalCode", SqlDbType.NVarChar, 50).Value = ExtTerminalCode
        MyCommon.LRTsp.Parameters.Add("@LayoutID", SqlDbType.Int).Value = LayoutID
        MyCommon.LRTsp.Parameters.Add("@FuelProcessing", SqlDbType.Bit).Value = IIf(FuelProcessing, 1, 0)
        MyCommon.LRTsp.Parameters.Add("@SpecificPromosOnly", SqlDbType.Int).Value = SpecificPromos
        'TerminalName = MyCommon.Parse_Quotes(TerminalName)
        If ExtTerminalCode = "" AndAlso (EngineType = InstalledEngines.CM OrElse EngineType = InstalledEngines.Catalina) Then
          infoMessage = Copient.PhraseLib.Lookup("terminal-edit.nonameandcode", LanguageID)
        ElseIf (TerminalName = "") Then
          infoMessage = Copient.PhraseLib.Lookup("terminal-edit.noname", LanguageID)
        ElseIf Not (EngineType = InstalledEngines.CM OrElse EngineType = InstalledEngines.Catalina) AndAlso IsNumeric(ExtTerminalCode) Then
          infoMessage = Copient.PhraseLib.Lookup("terminal-edit.badcode", LanguageID)
        ElseIf (EngineType = InstalledEngines.CM OrElse EngineType = InstalledEngines.Catalina) AndAlso (Not IsNumeric(ExtTerminalCode) OrElse ((ExtTerminalCode < 1) OrElse (Int(ExtTerminalCode) <> ExtTerminalCode))) Then
          infoMessage = Copient.PhraseLib.Lookup("terminal-edit.badcode", LanguageID)
        Else
          MyCommon.QueryStr = "SELECT Name FROM TerminalTypes with (NoLock) WHERE Name = '" & MyCommon.Parse_Quotes(TerminalName) & "' AND Deleted=0 AND AnyTerminal=0 AND TerminalTypeId <> " & TerminalId & " "
          If (BannersEnabled) Then
            MyCommon.QueryStr &= " and BannerID=" & BannerID
          End If
          rst2 = MyCommon.LRT_Select
          If (rst2.Rows.Count > 0) Then
            infoMessage = Copient.PhraseLib.Lookup("terminal-edit.nameused", LanguageID)
          Else
            MyCommon.QueryStr = "SELECT TerminalTypeID FROM TerminalTypes with (NoLock) WHERE EngineID In (0,1,4) AND Deleted=0 AND TerminalTypeID <> " & TerminalId & " AND AnyTerminal=0 AND ExtTerminalCode='" & ExtTerminalCode & "'"
            If (BannersEnabled) Then
              MyCommon.QueryStr &= " and BannerID=" & BannerID
            End If
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count > 0) And Not (EngineType = InstalledEngines.CPE) Then
              infoMessage = Copient.PhraseLib.Lookup("terminal-edit.codeused", LanguageID)
            Else
              MyCommon.LRTsp.ExecuteNonQuery()
              MyCommon.Activity_Log(21, TerminalId, AdminUserID, Copient.PhraseLib.Lookup("history.terminal-edit", LanguageID))
            End If
          End If
        End If
        MyCommon.Close_LRTsp()
        If TerminalLockingGroupID <> OldTerminalLockingGroupID Then
          MyCommon.QueryStr = "update TerminalTypes with (RowLock) set LockingGroupID=" & TerminalLockingGroupID & " where TerminalTypeID=" & TerminalId & ";"
          MyCommon.LRT_Execute()
        End If
      End If
      If infoMessage = "" Then
        Response.Status = "301 Moved Permanently"
        Response.AddHeader("Location", "terminal-edit.aspx?TerminalID=" & TerminalId)
      End If
    ElseIf bDelete Then
            MyCommon.QueryStr = "select distinct O.OfferID, Description from OfferTerminals as OT with (NoLock) " & _
                                "left join Offers as O with (NoLock) on O.OfferID=OT.OfferID " & _
                                "where O.Deleted=0 and OT.TerminalTypeID=" & Request.Form("TerminalID") & _
                                " union " & _
                                "select distinct CPE.IncentiveID as OfferID, Description from OfferTerminals as OT with (NoLock) " & _
                                "left join CPE_Incentives as CPE with (NoLock) on CPE.IncentiveID=OT.OfferID " & _
                                "where CPE.Deleted=0 and OT.TerminalTypeID=" & Request.Form("TerminalID")
      rst = MyCommon.LRT_Select
      If (rst.Rows.Count > 0) Then
        infoMessage = Copient.PhraseLib.Lookup("terminal-edit.inuse", LanguageID)
      Else
        MyCommon.QueryStr = "dbo.pt_TerminalTypes_Delete"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@TerminalTypeId", SqlDbType.Int).Value = TerminalId
        MyCommon.LRTsp.ExecuteNonQuery()
        MyCommon.Close_LRTsp()
        MyCommon.Activity_Log(21, TerminalId, AdminUserID, Copient.PhraseLib.Lookup("history.terminal-delete", LanguageID))
        Response.Status = "301 Moved Permanently"
        Response.AddHeader("Location", "terminal-list.aspx")
        TerminalId = 0
        TerminalName = ""
        TerminalDescription = ""
        ExtTerminalCode = ""
      End If
    End If
    
    LastUpdate = ""
    
    If Not bCreate Then
      ' no one clicked anything
      MyCommon.QueryStr = "select TT.ExtTerminalCode,TT.Name,TT.Description, TT.LayoutID, TT.SpecificPromosOnly, TT.FuelProcessing, TT.LastUpdate, TT.Deleted, " & _
                          "PE.PhraseID, PE.EngineID as EngineID, TT.BannerID, BAN.Name as BannerName, TT.LockingGroupId from TerminalTypes as TT with (nolock) " & _
                          "left join PromoEngines as PE with (NoLock) on PE.EngineID=TT.EngineID " & _
                          "left join Banners BAN with (NoLock) on TT.BannerID = BAN.BannerID and BAN.Deleted=0 " & _
                          "where TerminalTypeId=" & TerminalId
      rst = MyCommon.LRT_Select()
      If (rst.Rows.Count > 0) Then
        For Each row In rst.Rows
          If (ExtTerminalCode = "") Then
            If Not row.Item("ExtTerminalCode").Equals(System.DBNull.Value) Then
              ExtTerminalCode = row.Item("ExtTerminalCode")
            End If
          End If
          If (TerminalName = "") Then
            If Not row.Item("Name").Equals(System.DBNull.Value) Then
              TerminalName = row.Item("Name")
            End If
          End If
          If (TerminalDescription = "") Then
            If Not row.Item("Description").Equals(System.DBNull.Value) Then
              TerminalDescription = row.Item("Description")
            End If
          End If
          If (LastUpdate = "") Then
            If Not row.Item("LastUpdate").Equals(System.DBNull.Value) Then
              LastUpdate = row.Item("LastUpdate")
            End If
          End If
          If row.Item("Deleted") Then
            Deleted = True
            infoMessage = Copient.PhraseLib.Lookup("terminal-edit.deleted", LanguageID)
          End If
          If (EngineName = "") Then
            If Not row.Item("PhraseID").Equals(System.DBNull.Value) Then
              EngineName = Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID)
            End If
          End If
          If (EngineType = 0) Then
            If Not row.Item("EngineID").Equals(System.DBNull.Value) Then
              EngineType = row.Item("EngineID")
            End If
          End If
          If Not row.Item("LayoutID").Equals(System.DBNull.Value) Then
            LayoutID = row.Item("LayoutID")
          End If
          FuelProcessing = MyCommon.NZ(row.Item("FuelProcessing"), 0)

          SpecificPromos = MyCommon.NZ(row.Item("SpecificPromosOnly"), 0)
          If (SpecificPromos > 0) Then SpecificPromosChecked = " checked=""checked"""
          BannerID = MyCommon.NZ(rst.Rows(0).Item("BannerID"), 0)
          BannerName = MyCommon.NZ(rst.Rows(0).Item("BannerName"), Copient.PhraseLib.Lookup("term.multiple-banners", LanguageID))
          TerminalLockingGroupID = MyCommon.NZ(rst.Rows(0).Item("LockingGroupID"), 0)
          OldTerminalLockingGroupID = TerminalLockingGroupID
        Next
        If MyCommon.Fetch_CPE_SystemOption(86) = "2" AndAlso TerminalLockingGroupID > 0 Then
          MyCommon.QueryStr = "select TLG.Name as GroupName from TerminalLockingGroups as TLG with (NoLock) " & _
                "where TLG.TerminalLockingGroupID=" & TerminalLockingGroupID & " and TLG.Deleted=0;"
          rst2 = MyCommon.LRT_Select
          If rst2.Rows.Count > 0 Then
            TerminalLockingGroupName = rst2.Rows(0).Item("GroupName")
          Else
            TerminalLockingGroupName = Copient.PhraseLib.Lookup("error.deleted", LanguageID)
          End If
        Else
          TerminalLockingGroupName = ""
        End If
      ElseIf (TerminalId > 0) Then
        Send("")
        Send("<div id=""intro"">")
        Send("    <h1 id=""title"">" & Copient.PhraseLib.Lookup("term.terminal", LanguageID) & " #" & TerminalId & "</h1>")
        Send("</div>")
        Send("<div id=""main"">")
        Send("    <div id=""infobar"" class=""red-background"">")
        Send("        " & Copient.PhraseLib.Lookup("error.deleted", LanguageID))
        Send("    </div>")
        Send("</div>")
        GoTo done
      End If
    End If
    
    LayoutsDisplay = IIf(EngineType = InstalledEngines.CPE OrElse EngineType = InstalledEngines.Website, "block", "none")
    FocusField = IIf(EngineType = InstalledEngines.CM OrElse EngineType = InstalledEngines.Catalina, "ExtTerminalCode", "TerminalName")
    FuelProcessingChecked = IIf(FuelProcessing, " checked=""checked""", "")
%>

<script type="text/javascript" src="../javascript/jquery-1.10.2.min.js"></script>
<script type="text/javascript">
    var PageStart = 0;
    var PageEnd = 0;
    var DefaultPageSize = 5000
    var Fetch_UE_SystemOption_168 = '<%MyCommon.Fetch_UE_SystemOption(168)%>'

    function associatedOfferList_Prev() {
        if (PageStart != 0) {

            PageStart = PageStart - parseInt(DefaultPageSize);

            PageEnd = PageStart + parseInt(DefaultPageSize);

            GetAssociatedOfferList();
        }

        return false;
    }

    function associatedOfferList_Next() {
        if (PageStart == 0) {
            PageStart = parseInt(DefaultPageSize);
        }
        else {
            PageStart = PageEnd;
        }

        PageEnd = PageStart + parseInt(DefaultPageSize);

        GetAssociatedOfferList();

        return false;
    }

    function GetAssociatedOfferList() {
        var terminalId = '<%=Request.QueryString("TerminalId")%>'
        if (PageEnd == 0) { PageStart = 0; PageEnd = DefaultPageSize }

        var params = "mi=ASSOCIATED_OFFERS_LIST&TerminalID=" + terminalId + "&PageStart=" + (PageStart + 1) + "&PageEnd=" + PageEnd

        Ajax_GetAssociatedOfferList(params);
    }

    function Ajax_GetAssociatedOfferList(params) {
        var url = "terminal-edit.aspx";
        var assocName = ""
        var unknownText = '<%=Copient.PhraseLib.Lookup("term.unknown", LanguageID)%>'
        var expiredText = '<%=Copient.PhraseLib.Lookup("term.expired", LanguageID)%>'
        var IsOfferExpired = false
        var todayDate = new Date();
        var prodEndDate = null

        var objXMLHttpRequest = new XMLHttpRequest();
        objXMLHttpRequest.onreadystatechange = function () {
            if (objXMLHttpRequest.readyState == 4 && objXMLHttpRequest.status == 200) {
                var arrSplitResult = objXMLHttpRequest.responseText.split("_AMS_SPLITTER_AMS_");
                if (arrSplitResult[0] == 'T') {
                    try {
                        window.objJsonTable = JSON.parse(arrSplitResult[1]);
                        window.objStrHTML = "";

                        if (window.objJsonTable.length <= (DefaultPageSize - 1)) {
                            $("#lnkNext").hide();
                        }
                        else {
                            $("#lnkNext").show();
                        }

                        if (PageStart < 1 || window.objJsonTable.length == 0) {
                            $("#lnkPrev").hide();
                        }
                        else {
                            $("#lnkPrev").show();
                        }

                        if (PageStart > 1) {
                            $("#lnkPrev").show();
                        }

                        if (window.objJsonTable.length != 0) {
                            $("#functionselect").html("Loading..")
                            $("#functionselect").html("")
                            $("#functionselect").append("<table>")

                            $.each(window.objJsonTable, function (index, tRow) {
                                IsOfferExpired = false
                                if (Fetch_UE_SystemOption_168 == "1" && (tRow.BuyerID != null && tRow.BuyerID != "")) {
                                    assocName = "Buyer " + tRow.BuyerID + " - " + (tRow.Name == null || "" ? "" : tRow.Name)
                                }
                                else {
                                    assocName = (tRow.Name == null || "" ? unknownText : tRow.Name)
                                }

                                if (tRow.ProdEndDate != null && tRow.ProdEndDate != "") {
                                    prodEndDate = new Date(tRow.ProdEndDate);
                                    if (prodEndDate < todayDate) {
                                        IsOfferExpired = true
                                    }
                                }

                                if (tRow.IsAccessibleOffer == "1") {
                                    $("#functionselect").append("<tr><td><a href='offer-redirect.aspx?OfferID=" + tRow.OfferID + "'>" + assocName + "</a>" + (IsOfferExpired ? " (" + expiredText + ")" : "") + "</td></tr>")
                                }
                                else {
                                    $("#functionselect").append("<tr><td>" + assocName + (IsOfferExpired ? " (" + expiredText + ")" : "") + "</td></tr>")
                                }

                            });

                            $("#functionselect").append("</table>")
                        }
                        
                        //$("#functionselect").html(window.objStrHTML);

                        delete window.objJsonTable;
                        //delete window.objStrHTML;
                    }
                    catch (e) {
                    }
                }
                else {

                }
            };
        }
        objXMLHttpRequest.open("POST", url, true);
        objXMLHttpRequest.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
        objXMLHttpRequest.send(params);
    }

    window.onload =GetAssociatedOfferList()
    
    function handleLayouts(engineId) {
        var elem = document.getElementById("layoutSpan");
        var elemLayout = document.getElementById("layouts");
        var elemPromos = document.getElementById("specificpromos");
        var elemTerm = document.getElementById("ExtTerminalCode");
        var elemLblTerm = document.getElementById("lblExtTermCode");
        var elemFuel = document.getElementById("fuelprocessing");
        var elemFuelSpan = document.getElementById("fuelSpan");

        if (elem != null) {
          elem.style.display = (engineId == 2 || engineId == 3) ? "block" : "none";
          if (engineId != 2 && engineId != 3 && engineId != 9) {
            if (elemLayout != null) { elemLayout.options[0].selected = true; }
            if (elemPromos != null) { elemPromos.checked = false; }
            if (elemFuel != null) { elemFuel.checked = false; }
            if (elemTerm != null) { elemTerm.style.display = ''; }
            if (elemLblTerm != null) { elemLblTerm.style.display = ''; }
          } else if (engineId == 2 || engineId == 3 || engineId == 9) {
            if (elemTerm != null) { elemTerm.style.display = 'none'; }
            if (elemLblTerm != null) { elemLblTerm.style.display = 'none'; }
            if (elemFuelSpan != null) {
              elemFuelSpan.style.display = (engineId == 2 || engineId == 9) ? "block" : "none";
              if (elemFuel != null && engineId == 3) { elemFuel.checked = false; }
            }
          }
        }
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

    function handleBanners(bannerID) {
      var elem = document.getElementById("TermCodeSpan");
      var elemCode = document.getElementById("ExtTerminalCode");
      var engineID = -1;
            
      if (elem != null) {
       engineID = getBannerEngine(bannerID);
       var showTermCode = (engineID==2) ? false : true;
       elem.style.display = (showTermCode) ? "" : "none"; 
       if (elemCode != null) {
        elemCode.value = (showTermCode) ? elemCode.value : "";
       }
      }
    }
    
    function getBannerEngine(bannerID) {
      var index = -1;
      var engineID = -1;
      
      for (var i=0; i < bannerIDs.length && index==-1; i++) {
        if (bannerID == bannerIDs[i]) {
          index = i;
        }
      }
      
      if (bannerEngines.length > index) {
        engineID = bannerEngines[index]
      }
      
      return engineID;
    }
</script>

<form action="terminal-edit.aspx" id="mainform" name="mainform" onsubmit="return saveForm();" method="post">
  <input type="hidden" id="TerminalId" name="TerminalId" value="<% Sendb(TerminalId) %>" />
  <input type="hidden" id="TerminalLockingGroupID" name="TerminalLockingGroupID" value="<% Sendb(TerminalLockingGroupID) %>" />
  <input type="hidden" id="OldTerminalLockingGroupID" name="OldTerminalLockingGroupID" value="<% Sendb(OldTerminalLockingGroupID) %>" />
  <div id="intro">
    <%
      Sendb("<h1 id=""title"">")
      If TerminalId = 0 Then
        Sendb(Copient.PhraseLib.Lookup("term.newterminal", LanguageID))
      Else
        Sendb(Copient.PhraseLib.Lookup("term.terminal", LanguageID) & " #" & TerminalId & ": ")
        MyCommon.QueryStr = "SELECT Name FROM TerminalTypes with (NoLock) WHERE TerminalTypeId = " & TerminalId & ";"
        rst2 = MyCommon.LRT_Select
        If (rst2.Rows.Count > 0) Then
          TerminalNameTitle = MyCommon.NZ(rst2.Rows(0).Item("Name"), "")
        End If
        Sendb(MyCommon.TruncateString(TerminalNameTitle, 40))
      End If
      Sendb("</h1>")
    %>
    <div id="controls">
      <%
        If Not Deleted Then
          If (TerminalId = 0) Then
            If (Logix.UserRoles.EditTerminals) Then
              Send_Save()
            End If
          Else
            ShowActionButton = (Logix.UserRoles.EditTerminals)
            If (ShowActionButton) Then
              Send("<input type=""button"" class=""regular"" id=""actions"" name=""actions"" value=""" & Copient.PhraseLib.Lookup("term.actions", LanguageID) & " ▼"" onclick=""toggleDropdown();"" />")
              Send("<div class=""actionsmenu"" id=""actionsmenu"">")
              If (Logix.UserRoles.EditTerminals) Then
                Send_Save()
              End If
              If (Logix.UserRoles.EditTerminals) Then
                Send_Delete()
              End If
              If (Logix.UserRoles.EditTerminals) Then
                Send_New()
              End If
              Send("</div>")
            End If
            If MyCommon.Fetch_SystemOption(75) Then
              If (Logix.UserRoles.AccessNotes) Then
                Send_NotesButton(26, TerminalId, AdminUserID)
              End If
            End If
          End If
        End If
      %>
    </div>
  </div>
  <div id="main">
    <%
      If (infoMessage <> "") Then
        Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
      End If
      If Deleted Then
        GoTo DeleteSkip
      End If
    %>
    <div id="column1">
      <div class="box" id="identification">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.identification", LanguageID))%>
          </span>
        </h2>
        <span id="TermCodeSpan">
          <label id="lblExtTermCode" for="ExtTerminalCode" style="position: relative;"><% Sendb(Copient.PhraseLib.Lookup("term.code", LanguageID))%>:</label><br style="line-height: 0.1;" />
          <% Sendb("<input type=""text"" class=""longest"" id=""ExtTerminalCode"" name=""ExtTerminalCode"" maxlength=""50"" value=""" & ExtTerminalCode & """ />")%>
          <br class="half" />
        </span>
        <label for="TerminalName" style="position: relative;"><% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>:</label><br />
        <%
          If (TerminalName Is Nothing) Then TerminalName = ""
          Sendb("<input type=""text"" class=""longest"" id=""TerminalName"" name=""TerminalName"" maxlength=""100"" value=""" & TerminalName.Replace("""", "&quot;") & """ />")
        %>
        <br />
        <br class="half" />
        <label for="desc" style="position: relative;"><% Sendb(Copient.PhraseLib.Lookup("term.description", LanguageID))%>:</label><br />
        <textarea class="longest" cols="48" rows="3" id="desc" name="TerminalDescription"><% Sendb(TerminalDescription)%></textarea><br />
        <br class="half" />
        <span id="layoutSpan" style="display: <% Sendb(LayoutsDisplay) %>;">
          <label for="layouts"><% Sendb(Copient.PhraseLib.Lookup("term.layout", LanguageID))%>:</label><br />
          <select name="layouts" id="layouts" class="longest">
            <option value="0"><% Sendb(Copient.PhraseLib.Lookup("term.none", LanguageID))%></option>
            <%
              MyCommon.QueryStr = "select LayoutID, Name from ScreenLayouts with (NoLock) where Deleted=0 order by Name;"
              rst = MyCommon.LRT_Select
              For Each row In rst.Rows
                If (LayoutID = MyCommon.NZ(row.Item("LayoutID"), 0)) Then
                  Send("<option value=""" & row.Item("LayoutID") & """ selected=""selected"">" & row.Item("Name") & " </option>")
                Else
                  Send("<option value=""" & row.Item("LayoutID") & """>" & row.Item("Name") & " </option>")
                End If
              Next
            %>
          </select>
          <br />
          <br class="half" />
          <span id="fuelSpan">
            <input type="checkbox" id="fuelprocessing" name="fuelprocessing" value="1"<% sendb(fuelprocessingchecked) %> />
            <label for="fuelprocessing"><% Sendb(Copient.PhraseLib.Lookup("term.fuelprocessing", LanguageID))%></label>
          </span>
        </span>
        <br class="half" />
        <%
          If (BannersEnabled) Then
            If (TerminalId = 0) Then
              MyCommon.QueryStr = "select distinct BAN.BannerID, BAN.Name, BE.EngineID from Banners BAN with (NoLock) " & _
                                   "inner join AdminUserBanners AUB with (NoLock) on AUB.BannerID = BAN.BannerID " & _
                                   "inner join BannerEngines BE with (NoLock) on BE.BannerID = BAN.BannerID " & _
                                   "where BAN.Deleted=0 and BAN.AllBanners=0 and AdminUserID = " & AdminUserID & " order by BAN.Name;"
              rst = MyCommon.LRT_Select
              Send("<label for=""BannerID"">" & Copient.PhraseLib.Lookup("term.banner", LanguageID) & ":</label><br />")
              Send("<select class=""longest"" name=""BannerID"" id=""BannerID"" onchange=""handleBanners(this.value);"">")
              For Each row In rst.Rows
                Send("  <option value=""" & MyCommon.NZ(row.Item("BannerID"), -1) & """>" & MyCommon.NZ(row.Item("Name"), Copient.PhraseLib.Lookup("term.multiple-banners", LanguageID)) & "</option>")
                BannerIDs &= MyCommon.NZ(row.Item("BannerID"), -1) & ","
                BannerEngines &= MyCommon.NZ(row.Item("EngineID"), -1) & ","
              Next
              Send("  <option value=""0"">[" & Copient.PhraseLib.Lookup("term.multiple-banners", LanguageID) & "]</option>")
              BannerIDs &= "0"
              BannerEngines &= DefaultEngineID.ToString
              Send("</select>")
            Else
              Send("<input type=""hidden"" id=""EngineID"" name=""EngineID"" value=""" & EngineType & """ />")
              Send(Copient.PhraseLib.Lookup("term.banner", LanguageID) & ": " & MyCommon.SplitNonSpacedString(BannerName, 25))
            End If
            Send("<br /><br class=""half"" />")
          Else
            If (TerminalId = 0) Then
              ' spit out the engines available 
              'MyCommon.QueryStr = "Select EngineID,DefaultEngine,PhraseID from PromoEngines with (NoLock) where Installed=1;"  
              'Only allow for CPE and/or CM terminals
              MyCommon.QueryStr = "Select EngineID,DefaultEngine,PhraseID,Description from PromoEngines with (NoLock) where Installed=1 and EngineID in (0,2,9);"
              rst = MyCommon.LRT_Select
              If (rst.Rows.Count > 0) Then
                If rst.Rows.Count = 1 Then
                  EngineType = rst.Rows(0).Item("EngineID")
                  Send("<label for=""EngineIDlbl"">" & Copient.PhraseLib.Lookup("term.promotionengine", LanguageID) & ":</label><br />")
                  Send("<label name=""EngineIDlbl"" id=""EngineIDlbl"" >" & rst.Rows(0).Item("Description") & "</label><br />")
                  Send("<input type=""hidden"" id=""EngineID"" name=""EngineID"" value=""" & EngineType & """ />")
                Else
                  Send("<label for=""EngineID"">" & Copient.PhraseLib.Lookup("term.promotionengine", LanguageID) & ":</label><br />")
                  Send("<select id=""EngineID"" name=""EngineID"" class=""medium"" onchange=""javascript:handleLayouts(this.value);"">")
                  For Each row In rst.Rows
                    If MyCommon.NZ(row.Item("DefaultEngine"), 0) = 1 Then
                      Send("  <option selected=""selected"" value=""" & row.Item("EngineID") & """>" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID) & " </option>")
                      EngineType = MyCommon.NZ(row.Item("EngineID"), 0)
                    Else
                      Send("  <option value=""" & row.Item("EngineID") & """>" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID) & " </option>")
                    End If
                  Next
                  Send("</select><br />")
                End If
              End If
            Else
              Send("<input type=""hidden"" id=""EngineID"" name=""EngineID"" value=""" & EngineType & """ />")
              Sendb("                    " & Copient.PhraseLib.Lookup("term.promotionengine", LanguageID) & ": " & EngineName & "<br />")
            End If
          End If
        %>
        <br class="half" />
        <%
          MyCommon.QueryStr = "select ActivityDate from ActivityLog with (NoLock) where ActivityTypeID='21' and LinkID='" & TerminalId & "' order by ActivityDate asc;"
          dst = MyCommon.LRT_Select
          SizeOfData = dst.Rows.Count
          If SizeOfData > 0 Then
            Sendb(Copient.PhraseLib.Lookup("term.created", LanguageID) & " ")
            Send(Logix.ToLongDateTimeString(dst.Rows(0).Item("ActivityDate"), MyCommon))
            Send("<br />")
            Sendb(Copient.PhraseLib.Lookup("term.edited", LanguageID) & " ")
            Send(Logix.ToLongDateTimeString(dst.Rows(SizeOfData - 1).Item("ActivityDate"), MyCommon))
          End If
        %>
        <hr class="hidden" />
      </div>
    </div>
    
    <div id="gutter">
    </div>
    
    <div id="column2">
      <%
        ' CPE only
        If TerminalId > 0 AndAlso EngineType = InstalledEngines.CPE AndAlso MyCommon.Fetch_CPE_SystemOption(86) = "2" Then
          Send("<div class=""box"" id=""lockgroups"">")
          Send("<h2> <span>")
          Sendb(Copient.PhraseLib.Lookup("term.terminallockgroup", LanguageID))
          Send("</span> </h2>")
          Send("<input type=""radio"" id=""Pfunctionradio1"" name=""Pfunctionradio"" checked=""checked""" & disabledattribute & "/><label for=""Pfunctionradio1"">" & Copient.PhraseLib.Lookup("term.startingwith", LanguageID) & "</label>")
          Send("<input type=""radio"" id=""Pfunctionradio2"" name=""Pfunctionradio""" & disabledattribute & "/><label for=""Pfunctionradio2"">" & Copient.PhraseLib.Lookup("term.containing", LanguageID) & "</label><br />")
          Send("<input class=""medium"" onkeyup=""PhandleKeyUp(200);"" id=""Pfunctioninput"" name=""Pfunctioninput"" type=""text"" maxlength=""100"" value=""""" & disabledattribute & "/><br />")
          Send("<select class=""longer"" id=""Pfunctionselect"" name=""Pfunctionselect"" size=""10""" & disabledattribute & "/>")
          MyCommon.QueryStr = "select TerminalLockingGroupId as GroupID,Name as GroupName from TerminalLockingGroups with (NoLock) where deleted=0 and TerminalLockingGroupId is not null"
          If (BannersEnabled) Then
            MyCommon.QueryStr &= " and BannerID=" & BannerID
          Else
            MyCommon.QueryStr &= " and EngineID=" & EngineType
          End If
          MyCommon.QueryStr &= " order by GroupName;"
          rst2 = MyCommon.LRT_Select
          For Each row2 In rst2.Rows
            Send("<option value=" & row2.Item("GroupID") & ">" & row2.Item("GroupName") & "</option>")
          Next
          Send("</select> <br /> <br class=""half"" />")
          Send("<label for=""Pselected""><b>" & Copient.PhraseLib.Lookup("term.selectedprogram", LanguageID) & "</b></label><br />")
          Send("<input class=""regular"" id=""Pselect1"" name=""Pselect1"" type=""button"" value=""&#9660; " & Copient.PhraseLib.Lookup("term.select", LanguageID) & """ onclick=""PhandleSelectClick('Pselect1');""" & IIf(((rst2.Rows.Count = 0) Or (TerminalLockingGroupID > 0)), " disabled=""disabled""", "") & " />&nbsp;")
          Send("<input class=""regular"" id=""Pdeselect1"" name=""Pdeselect1"" type=""button"" value=""" & Copient.PhraseLib.Lookup("term.deselect", LanguageID) & " &#9650;"" onclick=""PhandleSelectClick('Pdeselect1');""" & IIf(TerminalLockingGroupID = 0, " disabled=""disabled""", "") & " /><br />")
          Send("<br class=""half"" />")
          Send("<select class=""longer"" id=""Pselected"" name=""Pselected"" size=""2""" & disabledattribute & ">")
          If TerminalLockingGroupID > 0 Then
            Send("<option value=""" & TerminalLockingGroupID & """>" & TerminalLockingGroupName & "</option>")
          End If
          Send("</select> <hr class=""hidden"" /> </div>")
        End If
      %>
      
    </div>
    <br clear="all" />
      <a id="lnkPrev" style="float:left" href="#" onclick="associatedOfferList_Prev()"><< Prev</a>
        <a id="lnkNext" style=" float:right; margin-right:18px" href="#" onclick="associatedOfferList_Next()">Next >></a>
        <br />
      <div class="box" id="offers" style="margin-right:15px" <%if(terminalid = 0)then sendb(" style=""visibility: hidden;""") %>">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.associatedoffers", LanguageID))%>
          </span>
        </h2>
        <div  id="functionselect" name="functionselect" style="height: 200px; overflow-y: scroll">
          
        </div>
        <hr class="hidden" />
      </div>
    <% DeleteSkip:%>
  </div>
</form>

<script type="text/javascript" language="javascript">
    <% If (BannersEnabled) Then %>
      var bannerIDs = Array(<% Sendb(BannerIDs) %>);
      var bannerEngines = Array(<% Sendb(BannerEngines) %>);
      
      if (document.getElementById("BannerID") != null) {
        handleBanners(document.getElementById("BannerID").value);
      }
    <% End If %>
    if (document.getElementById("EngineID") != null) {
        handleLayouts(document.getElementById("EngineID").value);
    }
</script>

<script type="text/javascript">

function saveForm(){
    var Pselected = document.getElementById('Pselected');
   
    if (Pselected != null && Pselected.options.length > 0) {
        document.getElementById("TerminalLockingGroupID").value = Pselected.options[0].value;
    }
    return true;
}


// This is the javascript array holding the function list
// The PrintJavascriptArray ASP function can be used to print this array.
<%
    MyCommon.QueryStr = "select TerminalLockingGroupId as GroupID,Name as GroupName from TerminalLockingGroups with (NoLock) where deleted=0 and TerminalLockingGroupId is not null"
    If (BannersEnabled) Then
        MyCommon.QueryStr &= " and BannerID=" & BannerID
    Else
        MyCommon.QueryStr &= " and EngineID=" & EngineType
    End If
    MyCommon.QueryStr &= " order by GroupName;"
    rst2 = MyCommon.LRT_Select
    
    If (rst2.rows.count>0)
        Sendb("var Pfunctionlist = Array(")
        For Each row2 In rst2.Rows
            Sendb("""" & MyCommon.NZ(row2.item("GroupName"), "").ToString().Replace("""", "\""") & """,")
        Next
        Send(""""");")
        Sendb("var Pvallist = Array(")
        For Each row2 In rst2.Rows
            Sendb("""" & row2.item("GroupID") & """,")
        Next
        Send(""""");")
    End If
%>

// This is the function that refreshes the list after a keypress.
// The maximum number to show can be limited to improve performance with
// huge lists (1000s of entries).
// The function clears the list, and then does a linear search through the
// globally defined array and adds the matches back to the list.
function PhandleKeyUp(maxNumToShow) {
  var selectObj, textObj, PfunctionListLength;
  var i, numShown;
  var searchPattern;
  var selectedList;
  var elem = document.getElementById("Pfunctionselect");
    
  if (elem != null)
  {

    elem.size = "10";
    
    // Set references to the form elements
    selectObj = document.forms[0].Pfunctionselect;
    textObj = document.forms[0].Pfunctioninput;
    selectedList = document.getElementById("Pselected");

    // Remember the function list length for loop speedup
    PfunctionListLength = Pfunctionlist.length;
    
    // Set the search pattern depending
    if(document.forms[0].Pfunctionradio[0].checked == true)
    {
        searchPattern = "^"+textObj.value;
    }
    else
    {
        searchPattern = textObj.value;
    }
    searchPattern = cleanRegExpString(searchPattern);

    // Create a regular expression

    re = new RegExp(searchPattern,"gi");
    // Clear the options list
    selectObj.length = 0;

    // Loop through the array and re-add matching options
    numShown = 0;
    for(i = 0; i < PfunctionListLength; i++)
    {
        if(Pfunctionlist[i].search(re) != -1)
        {
            if (Pvallist[i] != "" && (selectedList.options.length < 1 || Pvallist[i] != selectedList.options[0].value) ) {
                selectObj[numShown] = new Option(Pfunctionlist[i],Pvallist[i]);
                numShown++;
            }
        }
        // Stop when the number to show is reached
        if(numShown == maxNumToShow)
        {
            break;
        }
    }
    // When options list whittled to one, select that entry
    if(selectObj.length == 1)
    {
        selectObj.options[0].selected = true;
    }
  }
}

function PremoveUsed()
{
    PhandleKeyUp(99999);
    // this function will remove items from the functionselect box that are used in 
    // selected and excluded boxes

    var funcSel = document.getElementById('Pfunctionselect');
    var elSel = document.getElementById('Pselected');
    var i,j;
  
    for (i = elSel.length - 1; i>=0; i--) {
        for(j=funcSel.length-1;j>=0; j--) {
            if(funcSel.options[j].value == elSel.options[i].value){
                funcSel.options[j] = null;
            }
        }
    }
}


// this function gets the selected value and loads the appropriate
// php reference page in the display frame
// it can be modified to perform whatever action is needed, or nothing
function PhandleSelectClick(itemSelected)
{
    textObj = document.forms[0].Pfunctioninput;
     
    selectObj = document.forms[0].Pfunctionselect;
    selectedValue = document.getElementById("Pfunctionselect").value;
    if(selectedValue != ""){ selectedText = selectObj[document.getElementById("Pfunctionselect").selectedIndex].text; }
    
    selectboxObj = document.forms[0].Pselected;
    selectedboxValue = document.getElementById("Pselected").value;
    if(selectedboxValue != ""){ selectedboxText = selectboxObj[document.getElementById("Pselected").selectedIndex].text; }
    
    if(itemSelected == "Pselect1") {
        if(selectedValue != ""){
            // add items to selected box
            if(selectedValue == 1) {
                document.getElementById('Pselect1').disabled=true;
                // someone's adding all customers we need to empty the select box
                for (i = selectboxObj.length - 1; i>=0; i--) {
                    selectboxObj.options[i] = null;
                }
            }
            document.getElementById('Pdeselect1').disabled=false;
            document.getElementById('Pselect1').disabled=true;
            selectboxObj[selectboxObj.length] = new Option(selectedText,selectedValue);
        }
    }
    
    if(itemSelected == "Pdeselect1") {
        if(selectedboxValue != ""){
            // remove items from selected box
            document.getElementById("Pselected").remove(document.getElementById("Pselected").selectedIndex)
            if(selectedboxValue == 1) {
                document.getElementById('Pselect1').disabled=false;
            }
            if(selectboxObj.length == 0) {
                // nothing in the select box so disable deselect
                document.getElementById('Pdeselect1').disabled=true;
            }
        }
        if (document.getElementById("Pselected").options.length == 0) {
          document.getElementById('Pselect1').disabled=false;
          document.getElementById("TerminalLockingGroupID").value = 0;
        }
    }
    
    // remove items from large list that are in the other lists
    PremoveUsed();
    return true;
}

</script>

<script type="text/javascript">
    if (window.captureEvents){
        window.captureEvents(Event.CLICK);
        window.onclick=handlePageClick;
    }
    else {
        document.onclick=handlePageClick;
    }
    PhandleKeyUp(99999);
</script>

<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (TerminalId > 0 And Logix.UserRoles.AccessNotes) Then
      Send_Notes(26, TerminalId, AdminUserID)
    End If
  End If
  
done:
Finally
  MyCommon.Close_LogixRT()
End Try
Send_BodyEnd("mainform", "TerminalName")
MyCommon = Nothing
Logix = Nothing
%>