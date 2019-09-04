<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%@ Import Namespace="CMS.AMS" %>
<%@ Import Namespace="CMS.AMS.Contract" %>
<%  ' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<script type="text/javascript">

    var nVer = navigator.appVersion;
    var nAgt = navigator.userAgent;
    var browserName = navigator.appName;
    var nameOffset, verOffset, ix;

    var browser = navigator.appName;

    // In Opera, the true version is after "Opera" or after "Version"
    if ((verOffset = nAgt.indexOf("Opera")) != -1) {
        browserName = "Opera";
    }
    // In MSIE, the true version is after "MSIE" in userAgent
    else if ((verOffset = nAgt.indexOf("MSIE")) != -1) {
        browserName = "IE";
    }
    // In Chrome, the true version is after "Chrome"
    else if ((verOffset = nAgt.indexOf("Chrome")) != -1) {
        browserName = "Chrome";
    }
    // In Safari, the true version is after "Safari" or after "Version"
    else if ((verOffset = nAgt.indexOf("Safari")) != -1) {
        browserName = "Safari";
    }
    // In Firefox, the true version is after "Firefox"
    else if ((verOffset = nAgt.indexOf("Firefox")) != -1) {
        browserName = "Firefox";
    }
    // In most other browsers, "name/version" is at the end of userAgent
    else if ((nameOffset = nAgt.lastIndexOf(' ') + 1) <
          (verOffset = nAgt.lastIndexOf('/'))) {
        browserName = nAgt.substring(nameOffset, verOffset);
        fullVersion = nAgt.substring(verOffset + 1);
        if (browserName.toLowerCase() == browserName.toUpperCase()) {
            browserName = navigator.appName;
        }
    }

    if (browserName == "IE") {
        document.attachEvent("onclick", PageClick);
    }
    else {
        document.onclick = function (evt) {
            var target = document.all ? event.srcElement : evt.target;
            if (target.href) {
                if (IsFormChanged(document.mainform)) {
                    var bConfirm = confirm('Warning:  Unsaved data will be lost.  Are you sure you wish to continue?');
                    return bConfirm;
                }
            }
        };
    }
    function PageClick(evt) {
        var target = document.all ? event.srcElement : evt.target;

        if (target.href) {
            if (IsFormChanged(document.mainform)) {
                var bConfirm = confirm('Warning:  Unsaved data will be lost.  Are you sure you wish to continue?');
                return bConfirm;
            }
        }
    }
</script>
<%
    ' *****************************************************************************
    ' * FILENAME: lgroup-edit.aspx
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
    Dim GroupId As Long
    Dim GroupName As String
    Dim GName As String = ""
    Dim GroupDescription As String
    Dim ExtGroupId As String = ""
    Dim LastUpdate As String
    Dim CreatedDate As String
    Dim MyCommon As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    Dim rst As DataTable
    Dim row As DataRow
    Dim rst2 As DataTable
    Dim row2 As DataRow
    Dim bSave As Boolean
    Dim bAddStore As Boolean
    Dim bDelete As Boolean
    Dim bCreate As Boolean
    Dim bDownload As Boolean
    Dim bClose As Boolean
    Dim dtGroups As DataTable = Nothing
    Dim dtOffers As DataTable = Nothing
    Dim sQuery As String
    Dim EngineType As Integer
    Dim EngineName As String = ""
    Dim LocAvailableCount As Integer
    Dim LocAssignedCount As Integer
    Dim dtLocAvailable As DataTable = Nothing
    Dim dtLocAssigned As DataTable = Nothing
    Dim LocationList As String
    Dim Locations() As String
    Dim bAdd As Boolean
    Dim bAddAll As Boolean
    Dim bRemove As Boolean
    Dim GroupNameTitle As String = ""
    Dim Deleted As Boolean = False
    Dim ShowActionButton As Boolean = False
    Dim BannersEnabled As Boolean = False
    Dim BannerID As Integer = 0
    Dim BannerCt As Integer = 0
    Dim BannerName As String = ""
    Dim AllBannersPermission As Boolean = False
    Dim IE6ScrollFix As String = ""
    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False
    Dim OfferID As Integer = 0
    Dim EngineID As Integer = -1
    Dim LocEngineID As Integer = -1
    Dim CreatedFromOffer As Boolean = False
    Dim AllLocGroupID As Integer = 0
    Dim operateAtEnterprise As Boolean

    Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249)="1",True,False)
    Dim conditionalQuery = String.Empty
    
    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If

    ' Hierarchy stuff
    Dim ParentNodeIdList As String
    Dim SelectedNodeId As String
    Dim NodeName As String
    Dim bUp As Boolean
    Dim bDown As Boolean
    Dim NodeIds(-1) As String
    Dim ParentId As String
    Dim dr As DataRow
    Dim dtParents As DataTable = Nothing
    Dim dtParents1 As DataTable = Nothing
    Dim dtChildren As DataTable = Nothing
    Dim i As Integer
    Dim longDate As New DateTime
    Dim longDateString As String
    Dim ItemPKID As Integer = -1
    Dim ShowViewSelected As Boolean = False
    Dim File As HttpPostedFile
    Dim UploadOperation As Integer = 0
    Dim m_OfferService As IOffer
    Dim bCmToUeEnabled As Boolean = False
    Dim BrickAndMortarLocationId As Long = 0

    MyCommon.AppName = "lgroup-edit.aspx"
    CurrentRequest.Resolver.AppName = MyCommon.AppName
    m_OfferService = CurrentRequest.Resolver.Resolve(Of IOffer)()
    ' open database connection
    MyCommon.Open_LogixRT()

    If (Request.QueryString("new") <> "" OrElse Request.Form("new") <> "") Then
        Response.Redirect("lgroup-edit.aspx")
    End If

    BannersEnabled = MyCommon.Fetch_SystemOption(66)
    bCmToUeEnabled = (MyCommon.Fetch_CM_SystemOption(135) = "1")
  
    ' AL-6636 Removed UE@Enterprise system option so time zones will be sent to any UE system
    operateAtEnterprise = (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) AndAlso (MyCommon.Fetch_CPE_SystemOption(91) = "1")) OrElse (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE))

    If Request.QueryString("LargeFile") = "true" Then
        infoMessage = Copient.PhraseLib.Lookup("error.UploadTooLarge", LanguageID)
    End If

    Try
        AdminUserID = Verify_AdminUser(MyCommon, Logix)
        Response.Expires = 0

        ' fill in if it was a get method
        If Request.RequestType = "GET" Then
            GroupId = MyCommon.Extract_Val(Request.QueryString("LocationGroupID"))
            GroupName = Request.QueryString("GroupName")
            EngineType = Request.QueryString("EngineID")
            GroupDescription = Logix.TrimAll(Request.QueryString("GroupDescription"))
            ParentNodeIdList = Request.QueryString("ParentNodeIdList")
            SelectedNodeId = Request.QueryString("SelectedNodeId")
            NodeName = Request.QueryString("NodeName")
            BannerID = MyCommon.Extract_Val(Request.QueryString("banner"))
            If (Request.QueryString("download") = "") Then
                bDownload = False
            Else
                bDownload = True
            End If
            If Request.QueryString("save") = "" Then
                bSave = False
            Else
                bSave = True
            End If
            If Request.QueryString("store-add2") = "" Then
                bAddStore = False
            Else
                bAddStore = True
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
            If Request.QueryString("up") = "" Then
                bUp = False
            Else
                bUp = True
            End If
            If Request.QueryString("down") = "" Then
                bDown = False
            Else
                bDown = True
            End If
            If Request.QueryString("store-add") = "" Then
                bAdd = False
            Else
                bAdd = True
                MyCommon.Activity_Log(11, GroupId, AdminUserID, Copient.PhraseLib.Lookup("history.lgroup-add", LanguageID))
            End If
            If Request.QueryString("store-add-all") = "" Then
                bAddAll = False
            Else
                bAddAll = True
                MyCommon.Activity_Log(11, GroupId, AdminUserID, Copient.PhraseLib.Lookup("history.lgroup-add", LanguageID))
            End If
            If Request.QueryString("store-rem") = "" Then
                bRemove = False
            Else
                bRemove = True
                MyCommon.Activity_Log(11, GroupId, AdminUserID, Copient.PhraseLib.Lookup("history.lgroup-remove", LanguageID))
            End If
            If Request.QueryString("close") = "" Then
                bClose = False
            Else
                bClose = True
            End If
            UploadOperation = MyCommon.Extract_Val(Request.QueryString("Operation"))
        Else
            GroupId = MyCommon.Extract_Val(Request.Form("LocationGroupID"))
            If GroupId = 0 Then
                GroupId = MyCommon.Extract_Val(Request.QueryString("LocationGroupID"))
            End If
            GroupName = Request.Form("GroupName")
            EngineType = Request.Form("EngineID")
            GroupDescription = Logix.TrimAll(Request.Form("GroupDescription"))
            ParentNodeIdList = Request.Form("ParentNodeIdList")
            SelectedNodeId = Request.Form("SelectedNodeId")
            NodeName = Request.Form("NodeName")
            BannerID = MyCommon.Extract_Val(Request.Form("banner"))

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
            If Request.Form("download") = "" Then
                bDownload = False
            Else
                bDownload = True
            End If
            If Request.Form("store-add2") = "" Then
                bAddStore = False
            Else
                bAddStore = True
            End If
            If Request.Form("mode") = "" Then
                bCreate = False
            Else
                bCreate = True
            End If
            If Request.Form("up") = "" Then
                bUp = False
            Else
                bUp = True
            End If
            If Request.Form("down") = "" Then
                bDown = False
            Else
                bDown = True
            End If
            If Request.Form("store-add") = "" Then
                bAdd = False
            Else
                bAdd = True
                MyCommon.Activity_Log(11, GroupId, AdminUserID, Copient.PhraseLib.Lookup("history.lgroup-add", LanguageID))
            End If
            If Request.Form("store-add-all") = "" Then
                bAddAll = False
            Else
                bAddAll = True
                MyCommon.Activity_Log(11, GroupId, AdminUserID, Copient.PhraseLib.Lookup("history.lgroup-add", LanguageID))
            End If
            If Request.Form("store-rem") = "" Then
                bRemove = False
            Else
                bRemove = True
                MyCommon.Activity_Log(11, GroupId, AdminUserID, Copient.PhraseLib.Lookup("history.lgroup-remove", LanguageID))
            End If
            If Request.Form("close") = "" Then
                bClose = False
            Else
                bClose = True
            End If
            UploadOperation = MyCommon.Extract_Val(Request.Form("Operation"))
        End If

        OfferID = MyCommon.Extract_Val(Request.Form("OfferID"))
        If OfferID = 0 Then OfferID = MyCommon.Extract_Val(Request.QueryString("OfferID"))

        EngineID = IIf(Request.Form("EngineID") = "", -1, MyCommon.Extract_Val(Request.Form("EngineID")))
        If EngineID = -1 Then EngineID = IIf(Request.QueryString("EngineID") = "", -1, MyCommon.Extract_Val(Request.QueryString("EngineID")))

        CreatedFromOffer = OfferID > 0 And EngineID > 0
        If CreatedFromOffer And GroupId = 0 Then
            GName = Copient.PhraseLib.Lookup("term.offer", LanguageID) & " " & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.group", LanguageID), VbStrConv.Lowercase)
            MyCommon.QueryStr = "select count(*) as GroupCount from CustomerGroups where Name like '" & GName & "%';"
            rst = MyCommon.LRT_Select
            If rst.Rows(0).Item("GroupCount") > 0 Then
                GName = GName & " (" & rst.Rows(0).Item("GroupCount") & ")"
            End If
        End If

        ' START - Hierarchy ******************************************
        If bDown Then
            If ParentNodeIdList = "" Then
                ParentNodeIdList = SelectedNodeId
            Else
                ParentNodeIdList += "," & SelectedNodeId
            End If
            SelectedNodeId = ""
            NodeName = ""
        End If

        If bUp Then
            If ParentNodeIdList <> "" Then
                Dim n As Integer
                n = ParentNodeIdList.LastIndexOf(",")
                If n > 0 Then
                    SelectedNodeId = ParentNodeIdList.Substring(n + 1)
                    ParentNodeIdList = ParentNodeIdList.Substring(0, n)
                Else
                    SelectedNodeId = ParentNodeIdList
                    ParentNodeIdList = ""
                End If
            End If
            NodeName = ""
        End If

        If ParentNodeIdList = "" Then
            ' No hierarchy is selected, so no parents
            ' existing Hierarchies makeup children
            ParentId = ""
            'sQuery = "select HierarchyId as Id, [Name] as Name from LocationHierarchies with (NoLock)"
            sQuery = "select HierarchyId as Id,Name = " & _
                    "   case  " & _
                    "       when ExternalID is NULL then Name " & _
                    "       when ExternalID = '' then Name " & _
                    "       when ExternalID not like '%' + Name + '%' then ExternalID + '-' + Name " & _
                    "       else ExternalID " & _
                    "   end " & _
                    "from LocationHierarchies with (NoLock);"
        Else
            NodeIds = ParentNodeIdList.Split(",")
            'MyCommon.QueryStr = "select HierarchyId as Id, [Name] as Name from LocationHierarchies with (NoLock) where HierarchyId = " & NodeIds(0)
            MyCommon.QueryStr = "select HierarchyId as Id,Name = " & _
                                "   case  " & _
                                "       when ExternalID is NULL then Name " & _
                                "       when ExternalID = '' then Name " & _
                                "       when ExternalID not like '%' + Name + '%' then ExternalID + '-' + Name " & _
                                "       else ExternalID " & _
                                "   end " & _
                                "from LocationHierarchies with (NoLock) where HierarchyID=" & NodeIds(0)
            dtParents = MyCommon.LRT_Select()
            If NodeIds.Length = 1 Then
                ' No node is selected, selected hierarchy is parent
                ' existing root nodes for this hierarchy makeup children
                ParentId = "0"
                'sQuery = "select NodeId as Id, NodeName as Name from LHNodes with (NoLock) where ParentId = 0 and HierarchyId = " & NodeIds(0)
                sQuery = "select NodeId as Id, Name = " & _
                        "   case  " & _
                        "       when ExternalID is NULL then NodeName " & _
                        "       when ExternalID = '' then NodeName " & _
                        "       when ExternalID not like '%' + NodeName + '%' then ExternalID + '-' + NodeName " & _
                        "       else ExternalID " & _
                        "   end " & _
                        "from LHNodes with (NoLock) where ParentId = 0 and HierarchyId = " & NodeIds(0)
            Else
                For i = 1 To NodeIds.Length - 1
                    'MyCommon.QueryStr = "select NodeId as Id, NodeName as Name from LHNodes with (NoLock) where NodeId = " & NodeIds(i)
                    MyCommon.QueryStr = "select NodeId as Id, Name= " & _
                                        "   case  " & _
                                        "       when ExternalID is NULL then NodeName " & _
                                        "       when ExternalID = '' then NodeName " & _
                                        "       when ExternalID not like '%' + NodeName + '%' then ExternalID + '-' + NodeName " & _
                                        "       else ExternalID " & _
                                        "   end " & _
                                        "from LHNodes with (NoLock) where NodeId = " & NodeIds(i)
                    dtParents1 = MyCommon.LRT_Select()
                    dtParents.Merge(dtParents1)
                Next
                ' parents consist of hierarchy and listed nodes
                ' children made up of nodes with parentId = last parent
                ParentId = dtParents.Rows(dtParents.Rows.Count - 1).Item("id").ToString
                'sQuery = "select NodeId as Id, NodeName as Name from LHNodes with (NoLock) where ParentId = " & ParentId
                sQuery = "select NodeId as Id,Name= " & _
                        "   case  " & _
                        "       when ExternalID is NULL then NodeName " & _
                        "       when ExternalID = '' then NodeName " & _
                        "       when ExternalID not like '%' + NodeName + '%' then ExternalID + '-' + NodeName " & _
                        "       else ExternalID " & _
                        "   end " & _
                        "from LHNodes with (NoLock) where ParentId = " & ParentId
            End If
        End If
        MyCommon.QueryStr = sQuery
        dtChildren = MyCommon.LRT_Select()

        If SelectedNodeId = "" AndAlso dtChildren.Rows.Count > 0 Then
            SelectedNodeId = dtChildren.Rows(0).Item("id")
        End If
        ' END - Hierarchy ******************************************

        If Request.Files.Count >= 1 Then
            File = Request.Files.Get(0)
            If File.ContentType <> "text/plain" Then
                infoMessage = Copient.PhraseLib.Lookup("lgroup-edit.badfile", LanguageID)
                Response.AddHeader("Location", "lgroup-edit.aspx?LocationGroupID=" & GroupId)
            Else
                Dim LG As New Copient.LocationGroup(MyCommon, GroupId)
                LG.AdminUserID = AdminUserID
                'LG.LogFile = "" 'Add in a file location for logging
                LG.LanguageID = LanguageID
                Try
                    Dim UploadedText As String = Copient.PhraseLib.Lookup("history.lgroup-upload", LanguageID)
                    Select Case UploadOperation
                        Case 1
                            LG.AddLocations(Request.Files.Get(0))
                            UploadedText &= " (" & Copient.PhraseLib.Lookup("term.addedtogroup", LanguageID) & ")"
                        Case 2
                            LG.RemoveLocations(Request.Files.Get(0))
                            UploadedText &= " (" & Copient.PhraseLib.Lookup("term.removefromgroup", LanguageID) & ")"
                        Case Else
                            LG.ReplaceLocations(Request.Files.Get(0))
                            UploadedText &= " (" & Copient.PhraseLib.Lookup("term.fullreplacement", LanguageID) & ")"
                    End Select
                    MyCommon.Activity_Log(11, GroupId, AdminUserID, UploadedText)
                    Response.AddHeader("Location", "lgroup-edit.aspx?LocationGroupID=" & GroupId)
                Catch LocationEx As Copient.LocationException
                    infoMessage = LocationEx.Message
                Catch ex As Exception
                    infoMessage = "(" & Copient.PhraseLib.Lookup("term.ProcessingError", LanguageID) & "): " + ex.Message
                Finally

                End Try
            End If
        End If
        ' Download button clicked.
        If bDownload Then
            MyCommon.QueryStr = "select L.ExtLocationCode,L.locationname, l.TimeZone from locations as L with (NoLock) inner join LocGroupItems as LGI with (NoLock) on L.locationId = LGI.locationid where LGI.locationGroupID='" & GroupId & "' AND L.deleted =0 AND LGI.deleted =0"
            rst = MyCommon.LRT_Select()

            Response.Clear() 'AL-1401 Clear stream so that javascript from above is not displayed as part of the downloaded file
            If (rst.Rows.Count > 0) Then
                Response.AddHeader("Content-Disposition", "attachment; filename=LG" & GroupId & ".txt")
                Response.ContentType = "application/octet-stream"
                For Each row In rst.Rows
                    Sendb(MyCommon.NZ(row.Item("ExtLocationCode"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)))
                    Sendb(",")
                    If Not operateAtEnterprise Then
                        Send(MyCommon.NZ(row.Item("locationname"), 0))
                    Else
                        Sendb(MyCommon.NZ(row.Item("locationname"), 0))
                        Sendb(",")
                        Send(MyCommon.NZ(row.Item("TimeZone"), ""))
                    End If
                Next
                GoTo done
            Else
                infoMessage = Copient.PhraseLib.Lookup("lgroup-edit.noelements", LanguageID)
            End If
        End If
        ' lets see if they clicked save or delete
        If bSave OrElse (CreatedFromOffer AndAlso GroupId = 0) Then
            ' overwrite the engine id if this store is a member of a banner (if necessary)
            If (BannersEnabled AndAlso BannerID > 0) Then
                MyCommon.QueryStr = "select EngineID from BannerEngines BE with (NoLock) where BannerID=" & BannerID
                rst = MyCommon.LRT_Select
                If (rst.Rows.Count > 0) Then
                    EngineType = MyCommon.NZ(rst.Rows(0).Item("EngineID"), -1)
                End If
            End If

            If (GroupId = 0) Then
                MyCommon.QueryStr = "dbo.pt_LocationGroups_Insert"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 200).Value = Logix.TrimAll(GroupName)
                GroupName = Logix.TrimAll(GroupName)
                MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 1000).Value = GroupDescription
                MyCommon.LRTsp.Parameters.Add("@ExtGroupId", SqlDbType.NVarChar, 20).Value = ""
                MyCommon.LRTsp.Parameters.Add("@ExtSeqNum", SqlDbType.NVarChar, 20).Value = ""
                MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = EngineType
                If (BannersEnabled AndAlso BannerID > 0) Then
                    MyCommon.LRTsp.Parameters.Add("@BannerID", SqlDbType.Int).Value = BannerID
                End If
                MyCommon.LRTsp.Parameters.Add("@LocationGroupId", SqlDbType.BigInt).Direction = ParameterDirection.Output
                'GroupName = MyCommon.Parse_Quotes(GroupName)
                If (GroupName = "") Then
                    infoMessage = Copient.PhraseLib.Lookup("lgroup-edit.noname", LanguageID)
                Else
                    MyCommon.QueryStr = "SELECT LocationGroupID FROM LocationGroups WHERE Name = '" & MyCommon.Parse_Quotes(GroupName) & "' AND Deleted=0;"
                    rst = MyCommon.LRT_Select
                    If (rst.Rows.Count > 0) Then
                        infoMessage = Copient.PhraseLib.Lookup("lgroup-edit.nameused", LanguageID)
                    Else
                        MyCommon.LRTsp.ExecuteNonQuery()
                        GroupId = MyCommon.LRTsp.Parameters("@LocationGroupId").Value
                        MyCommon.Close_LRTsp()
                        MyCommon.Activity_Log(11, GroupId, AdminUserID, Copient.PhraseLib.Lookup("history.lgroup-create", LanguageID))
                    End If
                End If
            Else
                MyCommon.QueryStr = "dbo.pt_LocationGroups_Update"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@LocationGroupId", SqlDbType.BigInt).Value = GroupId
                MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 200).Value = Logix.TrimAll(GroupName)
                GroupName = Logix.TrimAll(GroupName)
                MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 1000).Value = GroupDescription
                'GroupName = MyCommon.Parse_Quotes(GroupName)
                If (GroupName = "") Then
                    infoMessage = Copient.PhraseLib.Lookup("lgroup-edit.noname", LanguageID)
                Else
                    MyCommon.QueryStr = "SELECT Name, LocationGroupID FROM LocationGroups WHERE Name = '" & MyCommon.Parse_Quotes(GroupName) & "' AND Deleted=0 AND LocationGroupID <> " & GroupId & ";"
                    rst = MyCommon.LRT_Select
                    If (rst.Rows.Count > 0) Then
                        infoMessage = Copient.PhraseLib.Lookup("lgroup-edit.nameused", LanguageID)
                    Else
                        MyCommon.LRTsp.ExecuteNonQuery()
                        MyCommon.Close_LRTsp()
                        SendNotificationsOfItemChange(GroupId, 3)
                        MyCommon.Activity_Log(11, GroupId, AdminUserID, Copient.PhraseLib.Lookup("history.lgroup-edit", LanguageID))
                    End If
                End If
            End If
            If infoMessage = "" Then
                Response.Status = "301 Moved Permanently"
                Response.AddHeader("Location", "lgroup-edit.aspx?LocationGroupID=" & GroupId)
            End If

        ElseIf bDelete Then
            'MyCommon.QueryStr = "select O.Name from OfferLocations as OL left join Offers as O on O.OfferID=OL.OfferID where OL.LocationGroupID=" & GroupId
            MyCommon.QueryStr = "select 1 as EngineID, O.Name as Name,O.OfferID as OfferID,O.ProdEndDate from OfferLocations as OL left join Offers as O on O.OfferID=OL.OfferID " & _
                                " where O.Deleted=0 and O.IsTemplate=0 and OL.Deleted=0 and OL.LocationGroupID=" & GroupId & _
                                " union " & _
                                "select OI.EngineID, I.IncentiveName as Name, I.IncentiveID as OfferID, I.EndDate as ProdEndDate from OfferLocations OL with (NoLock) " & _
                                "inner join CPE_Incentives I with (NoLock) on OL.OfferID = I.IncentiveID " & _
                                "inner join OfferIDs OI with (NoLock) on OI.OfferID = I.IncentiveID " & _
                                "where OL.LocationGroupID=" & GroupId & " and OL.Deleted=0 and I.Deleted=0 and I.IsTemplate=0 " & _
                                "order by Name;"
            rst = MyCommon.LRT_Select
            If rst.Rows.Count = 0 Then
                If (GroupId < 1) Then
                    infoMessage = Copient.PhraseLib.Lookup("lgroup-edit.nodelete", LanguageID)
                Else
                    MyCommon.QueryStr = "dbo.pt_LocationGroups_Delete"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@LocationGroupId", SqlDbType.BigInt).Value = GroupId
                    MyCommon.LRTsp.ExecuteNonQuery()
                    MyCommon.Close_LRTsp()
                    MyCommon.Activity_Log(11, GroupId, AdminUserID, Copient.PhraseLib.Lookup("history.lgroup-delete", LanguageID))
                    Response.Status = "301 Moved Permanently"
                    Response.AddHeader("Location", "lgroup-list.aspx")
                    GroupId = 0
                    GroupName = ""
                    GroupDescription = ""
                End If
            Else
                infoMessage = Copient.PhraseLib.Lookup("lgroup-edit.inuse", LanguageID)
            End If

        ElseIf bClose Then
            Response.Status = "301 Moved Permanently"
        End If

        CreatedDate = ""
        LastUpdate = ""

        If bAdd Then
            ' In this case the LocationList contains LocationId's from the LHContainer table
            LocationList = Request.Form("level-avail")
            If LocationList = "" Then LocationList = Request.QueryString("level-avail")

            If LocationList <> "" Then
                Locations = LocationList.Split(",")
                For i = 0 To Locations.Length - 1
                    MyCommon.QueryStr = "dbo.pt_LocGroupItems_Insert"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@LocationGroupId", SqlDbType.BigInt, 8).Value = GroupId.ToString
                    MyCommon.LRTsp.Parameters.Add("@LocationId", SqlDbType.BigInt, 8).Value = Locations(i)
                    MyCommon.LRTsp.Parameters.Add("@PkId", SqlDbType.BigInt).Direction = ParameterDirection.Output
                    MyCommon.LRTsp.ExecuteNonQuery()
                    MyCommon.Close_LRTsp()
                    SendNotificationsOfItemChange(GroupId, 3)
                    MyCommon.Activity_Log(11, GroupId, AdminUserID, Copient.PhraseLib.Lookup("history.lgroup-add", LanguageID) & ": " & Locations(i))
                Next
            End If
        ElseIf bAddAll Then
            If SelectedNodeId <> "" Then
                Dim HierarchyId As String
                Dim CurrentNode As String
                Dim dtLocations As DataTable
                If ParentNodeIdList = "" Then
                    ' Hierarchy level
                    HierarchyId = SelectedNodeId
                    CurrentNode = "0"
                Else
                    ' Node level
                    HierarchyId = NodeIds(0)
                    CurrentNode = SelectedNodeId
                End If
                sQuery = "select distinct LocationId as Id from GetBranchLocations(" & CurrentNode & "," & HierarchyId & ") where LocationId not in"
                sQuery += " (select LocationId from LocGroupItems with (NoLock) where Deleted = 0 and LocationGroupId = " & GroupId & ")"
                MyCommon.QueryStr = sQuery
                dtLocations = MyCommon.LRT_Select()
                If dtLocations.Rows.Count > 0 Then
                    For Each dr In dtLocations.Rows
                        MyCommon.QueryStr = "dbo.pt_LocGroupItems_Insert"
                        MyCommon.Open_LRTsp()
                        MyCommon.LRTsp.Parameters.Add("@LocationGroupId", SqlDbType.BigInt, 8).Value = GroupId.ToString
                        MyCommon.LRTsp.Parameters.Add("@LocationId", SqlDbType.BigInt, 8).Value = dr.Item("Id")
                        MyCommon.LRTsp.Parameters.Add("@PkId", SqlDbType.BigInt).Direction = ParameterDirection.Output
                        MyCommon.LRTsp.ExecuteNonQuery()
                        MyCommon.Close_LRTsp()
                        SendNotificationsOfItemChange(GroupId, 3)
                        MyCommon.Activity_Log(11, GroupId, AdminUserID, Copient.PhraseLib.Lookup("history.lgroup-add", LanguageID) & ": " & dr.Item("Id"))
                    Next
                End If
            End If
        End If

        If bRemove Then
            ' In this case the LocationList contains Primary Key values (PkId) from the LocationGroupItems table
            LocationList = Request.Form("contents")
            If LocationList = "" Then LocationList = Request.QueryString("contents")

            If LocationList <> "" Then
                Locations = LocationList.Split(",")
                For i = 0 To Locations.Length - 1
                    MyCommon.QueryStr = "dbo.pt_LocGroupItems_Delete"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@PkId", SqlDbType.BigInt, 8).Value = Locations(i)
                    MyCommon.LRTsp.ExecuteNonQuery()
                    MyCommon.Close_LRTsp()
                    SendNotificationsOfItemChange(GroupId, 3)
                Next
                If (MyCommon.Fetch_UE_SystemOption(191) = "1") Then m_OfferService.ProcessOfferCollisionDetectionStoreGroupChanges(GroupId.ToString())
            End If
        End If

        If bAddStore Then
            ' in this case we need to add all selected stores from the "fulllist" box
            LocationList = Request.Form("fulllist")
            If LocationList = "" Then LocationList = Request.QueryString("fulllist")

            If LocationList <> "" Then
                Locations = LocationList.Split(",")
                For i = 0 To Locations.Length - 1
                    MyCommon.QueryStr = "dbo.pt_LocGroupItems_Insert"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@LocationGroupId", SqlDbType.BigInt, 8).Value = GroupId.ToString
                    MyCommon.LRTsp.Parameters.Add("@LocationId", SqlDbType.BigInt, 8).Value = Locations(i)
                    MyCommon.LRTsp.Parameters.Add("@PkId", SqlDbType.BigInt).Direction = ParameterDirection.Output
                    MyCommon.LRTsp.ExecuteNonQuery()
                    MyCommon.Close_LRTsp()
                    SendNotificationsOfItemChange(GroupId, 3)
                    MyCommon.Activity_Log(11, GroupId, AdminUserID, Copient.PhraseLib.Lookup("history.lgroup-add", LanguageID) & ": " & Locations(i))
                Next
                If (MyCommon.Fetch_UE_SystemOption(191) = "1") Then m_OfferService.ProcessOfferCollisionDetectionStoreGroupChanges(GroupId.ToString())
            End If
        End If

        If Not bCreate Then
            ' no one clicked anything
            MyCommon.QueryStr = "select ExtGroupID,Name,LG.Description,CreatedDate,LG.LastUpdate,LG.EngineID,PE.Description as EngineName, PE.PhraseID as EnginePhraseID, BannerID, " & _
                                "isnull(BrickAndMortarLocationID,0) as BrickAndMortarLocationID from LocationGroups as LG " & _
                                "left join PromoEngines as PE on PE.EngineID=LG.EngineID where deleted=0 and LocationGroupId = " & GroupId
            rst = MyCommon.LRT_Select()
            If (rst.Rows.Count = 0 And GroupId <> 0) Then
                infoMessage = Copient.PhraseLib.Lookup("error.deleted", LanguageID)
                Deleted = True
            End If
            For Each row In rst.Rows
                If (ExtGroupId = "") Then
                    If Not row.Item("ExtGroupId").Equals(System.DBNull.Value) Then
                        ExtGroupId = row.Item("ExtGroupId")
                    End If
                End If
                If (GroupName = "") Then
                    If Not row.Item("Name").Equals(System.DBNull.Value) Then
                        GroupName = row.Item("Name")
                    End If
                End If
                If (EngineName = "") Then
                    EngineName = Copient.PhraseLib.Lookup(CInt(MyCommon.NZ(row.Item("EnginePhraseID"), 0)), LanguageID, MyCommon.NZ(row.Item("EngineName"), ""))
                End If
                If Not row.Item("EngineID").Equals(System.DBNull.Value) Then
                    EngineType = row.Item("EngineID")
                End If
                If (GroupDescription = "") Then
                    If Not row.Item("Description").Equals(System.DBNull.Value) Then
                        GroupDescription = row.Item("Description")
                    End If
                End If
                If (CreatedDate = "") Then
                    If Not row.Item("CreatedDate").Equals(System.DBNull.Value) Then
                        CreatedDate = row.Item("CreatedDate")
                    End If
                End If
                If (LastUpdate = "") Then
                    If Not row.Item("LastUpdate").Equals(System.DBNull.Value) Then
                        LastUpdate = row.Item("LastUpdate")
                    End If
                End If
                If (BannersEnabled) Then
                    BannerID = MyCommon.NZ(row.Item("BannerID"), 0)
                    MyCommon.QueryStr = "select Name from Banners with (NoLock) where BannerID=" & BannerID
                    rst2 = MyCommon.LRT_Select
                    If (rst2.Rows.Count > 0) Then
                        BannerName = MyCommon.NZ(rst2.Rows(0).Item("Name"), "")
                    End If
                End If
        
                If (bCmToUeEnabled) Then
                  BrickAndMortarLocationId = MyCommon.NZ(row.Item("BrickAndMortarLocationID"), 0)
                End If

            Next
            sQuery = "select distinct b.LocationGroupId,b.Name from LocGroupItems a,LocationGroups b with (NoLock)"
            sQuery += " where b.LocationGroupId = a.LocationGroupId and a.Deleted = 0 and a.LocationId = " & GroupId
            MyCommon.QueryStr = sQuery
            dtGroups = MyCommon.LRT_Select
            sQuery = "select distinct d.OfferId,d.Name from LocGroupItems a,LocationGroups b,OfferLocations c,Offers d with (NoLock)"
            sQuery += " where d.deleted = 0 and d.OfferId = c.OfferId and c.Deleted = 0 and c.Excluded = 0 and c.LocationGroupId = b.LocationGroupId"
            sQuery += " and b.LocationGroupId = a.LocationGroupId and a.Deleted = 0 and a.LocationId = " & GroupId
            MyCommon.QueryStr = sQuery
            dtOffers = MyCommon.LRT_Select
        End If

        If Not dtParents Is Nothing AndAlso dtParents.Rows.Count > 0 AndAlso dtChildren.Rows.Count > 0 Then
            sQuery = "select a.LocationID as Id, a.ExtLocationcode as Code, a.LocationName as Name, b.PKID from Locations a, LHContainer b"
            sQuery += " with (NoLock) where a.LocationID=b.LocationID and b.NodeID=" & SelectedNodeId
            If GroupId > 0 Then
                sQuery += " and a.LocationID not in ("
                sQuery += "select LocationID from LocGroupItems with (NoLock) where Deleted=0 and LocationGroupID=" & GroupId & ")"
            End If
            MyCommon.QueryStr = sQuery
            dtLocAvailable = MyCommon.LRT_Select()
            LocAvailableCount = dtLocAvailable.Rows.Count
        Else
            LocAvailableCount = 0
        End If

        If GroupId > 0 Then
            sQuery = "select b.PKID as Id, a.ExtLocationCode as Code, a.LocationName as Name from Locations a, LocGroupItems b"
            sQuery += " with (NoLock) where a.LocationID=b.LocationID and b.Deleted=0 and b.LocationGroupID=" & GroupId
            MyCommon.QueryStr = sQuery
            dtLocAssigned = MyCommon.LRT_Select()
            LocAssignedCount = dtLocAssigned.Rows.Count
        End If

        If (Request.Form("ItemPK") <> "") Then
            Integer.TryParse(Request.Form("ItemPK"), ItemPKID)
        ElseIf (Request.QueryString("ItemPK") <> "") Then
            Integer.TryParse(Request.QueryString("ItemPK"), ItemPKID)
        End If

        Send_HeadBegin("term.storegroup", , GroupId)
        Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
        Send_Metas()
        Send_Links(Handheld)
        Send_Scripts()
        Send_HeadEnd()
        If CreatedFromOffer Then
            Send_BodyBegin(3)
        Else
            Send_BodyBegin(1)
            Send_Bar(Handheld)
            Send_Help(CopientFileName)
            Send_Logos()
            Send_Tabs(Logix, 7)
            Send_Subtabs(Logix, 71, 4, , GroupId)
        End If
        If (Logix.UserRoles.AccessStoreGroups = False) Then
            Send_Denied(1, "perm.lgroup-access")
            GoTo done
        End If

        If (BannersEnabled And GroupId > 0) Then
            ' check if the user is allowed to view this bannered location group
            MyCommon.QueryStr = "select BannerID from LocationGroups LG with (NoLock) " & _
                                "where LocationGroupID = " & GroupId & " and (BannerID is Null or BannerID =0 " & _
                                "or BannerID in (select BannerID from AdminUserBanners where AdminUserID=" & AdminUserID & "))"
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count = 0) Then
                Send_Denied(1, "banners.access-denied-offer")
                Send_BodyEnd()
                GoTo done
            End If
        End If

        ' determine whether to show the View selected hierarchy nodes button
        MyCommon.QueryStr = "select count(NodeID) as NodeCount from LocationGroupNodes with (NoLock) where LocationGroupID=" & GroupId
        rst = MyCommon.LRT_Select
        ShowViewSelected = (MyCommon.NZ(rst.Rows(0).Item("NodeCount"), 0) > 0)
%>
<script type="text/javascript">
  // This is the javascript array holding the function list
  // The PrintJavascriptArray ASP function can be used to print this array.
  <%
    If (BannersEnabled) Then
      MyCommon.QueryStr = "select BE.BannerID from AdminUserBanners AUB with (NoLock) " & _
                          " inner join Banners BAN with (NoLock) on BAN.BannerID=AUB.BannerID " & _
                          " inner join BannerEngines BE with (NoLock) on BE.BannerID=BAN.BannerID " & _
                          " where AdminUserID=" & AdminUserID & " and BE.EngineID=" & EngineType & " and BAN.AllBanners=1;"
      rst = MyCommon.LRT_Select
      AllBannersPermission = (rst.Rows.Count > 0)
      MyCommon.QueryStr = "select LocationID, LocationName, ExtLocationCode from Locations " & _
                          " where Deleted=0 and EngineID=" & EngineType & " and LocationTypeID=1 and LocationID not in " & _
                          " (select LocationID from LocGroupItems where LocationGroupID=" & GroupId & " and Deleted=0) " & _
                          " and (" & IIf(AllBannersPermission, "BannerID is null or ", "") & "BannerID = " & BannerID & ") " & _
                          " order by ExtLocationCode;"
    Else
      MyCommon.QueryStr = "select LocationID, LocationName, ExtLocationCode from Locations " & _
                          " where Deleted=0 and EngineID=" & EngineType & " and LocationTypeID=1 and LocationID not in " & _
                          " (select LocationID from LocGroupItems where LocationGroupID=" & GroupId & " and Deleted=0) " & _
                          " order by ExtLocationCode;"
    End If
    rst = MyCommon.LRT_Select
    If (rst.rows.count>0)
      Sendb("var functionlist = Array(")
      For Each row In rst.Rows
        Sendb("""" & MyCommon.NZ(row.item("ExtLocationCode"), "").ToString().Replace("""", "\""") & " - " & MyCommon.NZ(row.item("LocationName"), "").ToString().Replace("""", "\""") &  """,")
      Next
      Send(""""");")
      Sendb("var vallist = Array(")
      For Each row In rst.Rows
        Sendb("""" & row.item("LocationID") & """,")
      Next
      Send(""""");")
    Else
      Sendb("var functionlist = Array(")
      Send(""""");")
      Sendb("var vallist = Array(")
      Send(""""");")
    End If
  %>
  <%
    If (not dtLocAssigned is nothing andalso dtLocAssigned.Rows.Count>0)
      Sendb("var functionlist2 = Array(")
      For Each row In dtLocAssigned.Rows
        Sendb("""" & MyCommon.NZ(row.Item("Code"), "").ToString().Replace("""", "\""") & " - " & MyCommon.NZ(row.Item("Name"), "").ToString().Replace("""", "\""") &  """,")
      Next
      Send(""""");")
      Sendb("var vallist2 = Array(")
      For Each row In dtLocAssigned.Rows
        Sendb("""" & row.Item("Id") & """,")
      Next
      Send(""""");")
    Else
      Sendb("var functionlist2 = Array(")
      Send(""""");")
      Sendb("var vallist2 = Array(")
      Send(""""");")
    End If
  %>
//        function chooseFile() {
//      document.getElementById("browse").click();
//   }
//   function fileonclick()
//   {
//   var filename=document.getElementById("browse").value;
//    document.getElementById("lblfileupload").innerText = filename.replace("C:\\fakepath\\", "");
//  }
  function handleKeyUp(maxNumToShow) {
    var selectObj, textObj, functionListLength;
    var i, numShown;
    var searchPattern;
    var selectedList;

    document.getElementById("fulllist").size = "8";

    // Set references to the form elements
    selectObj = document.forms[0].fulllist;
    textObj = document.forms[0].functioninput;

    // Remember the function list length for loop speedup
    functionListLength = functionlist.length;

    // Set the search pattern depending
    if(document.forms[0].functionradio[0].checked == true) {
      searchPattern = "^"+textObj.value;
    } else {
      searchPattern = textObj.value;
    }
    searchPattern = cleanRegExpString(searchPattern);

    // Create a regular expression

    re = new RegExp(searchPattern,"gi");
    // Clear the options list
    selectObj.length = 0;

    // Loop through the array and re-add matching options
    numShown = 0;
    for(i = 0; i < functionListLength; i++) {
      if(functionlist[i].search(re) != -1) {
        if (vallist[i] != "") {
          selectObj[numShown] = new Option(functionlist[i],vallist[i]);
          numShown++;
        }
      }
      // Stop when the number to show is reached
      if(numShown == maxNumToShow) {
        break;
      }
    }
    // When options list whittled to one, select that entry
    if(selectObj.length == 1) {
      selectObj.options[0].selected = true;
    }
  }

  function handleKeyDown(e) {
    var key = e.which ? e.which : e.keyCode;

    if (key == 40) {
      var elemSlct = document.getElementById("fulllist");
      if (elemSlct != null && elemSlct.options.length > 0) {
        elemSlct.options[0].selected = true;
        elemSlct.focus();
      }
    }
  }

  function handleMemberKeyUp(maxNumToShow) {
    var selectObj, textObj, functionListLength;
    var i,  numShown;
    var searchPattern;
    var selectedList;

    document.getElementById("contents").size = "8";

    // Set references to the form elements
    selectObj = document.forms[0].contents;
    textObj = document.forms[0].memberinput;

    // Remember the function list length for loop speedup
    functionListLength = functionlist2.length;

    // Set the search pattern depending
    if(document.forms[0].memberradio[0].checked == true) {
      searchPattern = "^"+textObj.value;
    } else {
      searchPattern = textObj.value;
    }
    searchPattern = cleanRegExpString(searchPattern);

    // Create a regular expression

    re = new RegExp(searchPattern,"gi");
    // Clear the options list
    selectObj.length = 0;

    // Loop through the array and re-add matching options
    numShown = 0;
    for(i = 0; i < functionListLength; i++) {
      if(functionlist2[i].search(re) != -1) {
        if (vallist2[i] != "") {
          selectObj[numShown] = new Option(functionlist2[i],vallist2[i]);
          numShown++;
        }
      }
      // Stop when the number to show is reached
      if(numShown == maxNumToShow) {
        break;
      }
    }
    // When options list whittled to one, select that entry
    if(selectObj.length == 1) {
      selectObj.options[0].selected = true;
    }
  }

  function handleMemberKeyDown(e) {
    var key = e.which ? e.which : e.keyCode;

    if (key == 40) {
      var elemSlct = document.getElementById("contents");
      if (elemSlct != null && elemSlct.options.length > 0) {
        elemSlct.options[0].selected = true;
        elemSlct.focus();
      }
    }
  }

  function SubmitForm() {
    document.mainform.submit();
  }

  function launchSearch() {
    openPopup('lhierarchy-search.aspx?LocationGroupID=<%Sendb(GroupId) %>');
  }

    function storeSelectedToAdd(btnClicked)
    {
        var selected = $("#fulllist").val();
        if (selected) {
            return checkForStoreChange(btnClicked);
        }
        else {
            alert('<% Sendb(Copient.PhraseLib.Lookup("lgroup-edit.selectstore", LanguageID)) %>');
        }
    }
  function checkForStoreChange(btnClicked) {
    var retVal = true;
    var promptAboutOffers = false
    var promptMsg = ""

    if (btnClicked=='store-add' || btnClicked=='store-add2' || btnClicked=='store-add-all') {
      promptAboutOffers = true;
      promptMsg = '<% Sendb(Copient.PhraseLib.Lookup("lgroup-edit.addstoreconfirm", LanguageID)) %>';
    } else if (btnClicked=='store-rem') {
      promptAboutOffers = true;
      promptMsg = '<% Sendb(Copient.PhraseLib.Lookup("lgroup-edit.removestoreconfirm", LanguageID)) %>';
    }

    if (promptAboutOffers) {
      retVal = confirm(promptMsg);
    }

    return retVal;
  }

  function showOffersWindow() {
    var offerWin = openPopup('lgroup-offers.aspx?LocationGroupID=<% Sendb(GroupId) %>');
  }

  function launchHierarchy() {
    var popW = 700;
    var popH = 570;
    var url = 'lhierarchytree.aspx?LocationGroupID=<% Sendb(GroupId) %>&OfferID=<%Sendb(OfferID)%>&EngineID=<%Sendb(EngineID)%>';

    lhierWindow = window.open(url,"hierTree", "width="+popW+", height="+popH+", top="+calcTop(popH)+", left="+calcLeft(popW)+", toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=yes");
    lhierWindow.focus();
  }

  function launchNodes() {
    var popW = 700;
    var popH = 570;
    var elemGroupName = document.getElementById("GroupName");
    var groupName = '';

    if (elemGroupName != null) {
      groupName = elemGroupName.value;
    }

    var url = 'lgroup-edit-nodes.aspx?LocationGroupID=<% Sendb(GroupId) %>&Name=' + escape(groupName);

    nodeWindow = window.open(url,"LNodes", "width="+popW+", height="+popH+", top="+calcTop(popH)+", left="+calcLeft(popW)+", toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=yes");
    nodeWindow.focus();
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

  function showWait()
  {
    var elemLoading = document.getElementById("loading");
    var elemFile = document.getElementById("browse");
    var elemUploader = document.getElementById("uploader");
    elemUploader.style.display = 'none';
    if (elemLoading != null && elemFile != null) {
      elemLoading.style.display = "block";
    }

  }
</script>
<form id="mainform" name="mainform" action="#" method="post">
<%
    If CreatedFromOffer Then
        Send("<input type=""hidden"" id=""OfferID"" name=""OfferID"" value=""" & OfferID & """ />")
        Send("<input type=""hidden"" id=""EngineID"" name=""EngineID"" value=""" & EngineID & """ />")
        If Request.RequestType = "GET" Then
            Send("<input type=""hidden"" id=""slct"" name=""slct"" value=""" & Request.QueryString("slct") & """ />")
            Send("<input type=""hidden"" id=""ex"" name=""ex"" value=""" & Request.QueryString("ex") & """ />")
            Send("<input type=""hidden"" id=""condChanged"" name=""condChanged"" value=""" & Request.QueryString("condChanged") & """ />")
        Else
            Send("<input type=""hidden"" id=""slct"" name=""slct"" value=""" & Request.Form("slct") & """ />")
            Send("<input type=""hidden"" id=""ex"" name=""ex"" value=""" & Request.Form("ex") & """ />")
            Send("<input type=""hidden"" id=""condChanged"" name=""condChanged"" value=""" & Request.Form("condChanged") & """ />")
        End If
    End If
%>
<input type="hidden" id="submitButton" name="submitButton" value="" />
<div id="intro">
    <h1 id="title">
        <%  If GroupId = 0 Then
                Sendb(Copient.PhraseLib.Lookup("term.newstoregroup", LanguageID))
            Else
                Sendb(Copient.PhraseLib.Lookup("term.storegroup", LanguageID) & " #" & GroupId & ": ")
                MyCommon.QueryStr = "SELECT Name, LocationGroupID FROM LocationGroups WHERE LocationGroupId = " & GroupId & ";"
                rst2 = MyCommon.LRT_Select
                If (rst2.Rows.Count > 0) Then
                    GroupNameTitle = MyCommon.NZ(rst2.Rows(0).Item("Name"), "")
                    Sendb(MyCommon.TruncateString(GroupNameTitle, 40))
                End If
            End If
        %>
    </h1>
    <div id="controls">
        <%
            If (GroupId = 0) Then
                If (Logix.UserRoles.CreateStoreGroups) Then
                    Send_Save()
                End If
            Else
                ShowActionButton = (Logix.UserRoles.CreateStoreGroups) OrElse (Logix.UserRoles.EditStoreGroups) OrElse (Logix.UserRoles.DeleteStoreGroups)
                If (ShowActionButton) Then
                    Send("<input type=""button"" class=""regular"" id=""actions"" name=""actions"" onclick=""toggleDropdown();"" value=""" & Copient.PhraseLib.Lookup("term.actions", LanguageID) & " ▼"" />")
                    Send("<div class=""actionsmenu"" id=""actionsmenu"">")
                    If (Logix.UserRoles.EditStoreGroups And BrickAndMortarLocationId = 0) Then
                        Send_Save()
                    End If
                    If (Logix.UserRoles.DeleteStoreGroups And BrickAndMortarLocationId = 0) Then
                        Send_Delete()
                    End If
                    If (Logix.UserRoles.EditStoreGroups And BrickAndMortarLocationId = 0) AndAlso (GroupId > 0) Then
                        Send_Upload()
                    End If
                    If (Logix.UserRoles.AccessStoreGroups And BrickAndMortarLocationId = 0) Then
                        Send_Download()
                    End If
                    If (Logix.UserRoles.CreateStoreGroups And Not CreatedFromOffer) Then
                        Send_New()
                    End If
                    If CreatedFromOffer Then
                        Send_Close()
                    End If
                    If Request.Browser.Type = "IE6" Then
                        Send("<iframe src=""javascript:'';"" id=""actionsiframe"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no"" style=""height:75px;""></iframe>")
                    End If
                    Send("</div>")
                End If
                If (MyCommon.Fetch_SystemOption(75) And Deleted = False And Not CreatedFromOffer) Then
                    If (Logix.UserRoles.AccessNotes) Then
                        Send_NotesButton(14, GroupId, AdminUserID)
                    End If
                End If
            End If
            Sendb(" <input type=""hidden"" id=""LocationGroupID1"" name=""LocationGroupId"" value=""" & GroupId & """ />")
        %>
    </div>
</div>
<%
    If Request.Browser.Type = "IE6" Then
        IE6ScrollFix = " onscroll=""javascript:document.getElementById('actionsmenu').style.visibility='hidden';"""
    End If
%>
<div id="main" <% Sendb(IE6ScrollFix) %>>
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"" style=""word-wrap: break-word;"">" & infoMessage & "</div>")%>
    <%
        If Deleted Then
            Send("</div>")
            Send("</form>")
            GoTo done
        End If
    %>
    <div id="column1">
        <div class="box" id="identification">
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.identification", LanguageID))%></span></h2>
            <label for="GroupName">
                <% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>:</label><br />
            <%  If (GroupName Is Nothing) Then GroupName = ""
                Sendb("<input class=""" & IIf(CreatedFromOffer, "long", "longest") & """ id=""GroupName"" name=""GroupName"" maxlength=""100"" type=""text"" value=""" & GroupName.Replace("""", "&quot;") & """ " & IIf(BrickAndMortarLocationId > 0, " disabled", "") & "/>")%>
            <br />
            <label for="desc">
                <% Sendb(Copient.PhraseLib.Lookup("term.description", LanguageID))%>:</label><br />
            <textarea class="<% Sendb(IIf(CreatedFromOffer, "long", "longest")) %>" id="desc"
                cols="48" rows="3" name="GroupDescription"><% Sendb(GroupDescription)%></textarea><br />
            <br class="half" />
            <%
                If ExtGroupId <> "" AndAlso ExtGroupId <> "0" Then
                    Sendb(Copient.PhraseLib.Lookup("term.xid", LanguageID) & ": " & ExtGroupId)
                End If
                If (GroupId = 0) Then
                    If (BannersEnabled) Then
                        MyCommon.QueryStr = "select distinct BAN.BannerID, BAN.Name from Banners BAN with (NoLock) " & _
                                            "inner join AdminUserBanners AUB with (NoLock) on AUB.BannerID = BAN.BannerID " & _
                                            "where BAN.Deleted=0 and BAN.AllBanners=0 and AdminUserID = " & AdminUserID & " order by BAN.Name;"
                        rst = MyCommon.LRT_Select
                        BannerCt = rst.Rows.Count
                        If (BannerCt > 0) Then
                            BannerID = MyCommon.Extract_Val(Request.Form("banner"))
                            If BannerID = 0 Then BannerID = MyCommon.Extract_Val(Request.QueryString("banner"))
                            Send("<br class=""half"" />")
                            Send("<label for=""BannerID"">" & Copient.PhraseLib.Lookup("term.banners", LanguageID) & ":</label><br />")
                            Send("  <select class=""" & IIf(CreatedFromOffer, "long", "longest") & """ name=""banner"" id=""banner"">")
                            For Each row In rst.Rows
                                Send("    <option value=""" & MyCommon.NZ(row.Item("BannerID"), -1) & """" & IIf(BannerID = MyCommon.NZ(row.Item("BannerID"), -1), " selected=""selected""", "") & ">" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
                            Next
                            Send("  </select>")
                            Send("&nbsp;<br class=""half"" />")
                        End If
                    Else
                        ' Spit out the engines available
                        MyCommon.QueryStr = "select EngineID, Description, PhraseID, DefaultEngine from PromoEngines with (NoLock) where Installed=1 and EngineID in (0, 2, 9);"
                        rst = MyCommon.LRT_Select
                        If (rst.Rows.Count > 0) Then
                            Send("<label for=""EngineID"">" & Copient.PhraseLib.Lookup("term.promotionengine", LanguageID) & ":</label><br />")
                            Send("<select class=""medium"" name=""EngineID"">")
                            For Each row In rst.Rows
                                Sendb("<option value=""" & row.Item("EngineID") & """" & IIf(MyCommon.NZ(row.Item("DefaultEngine"), 0) = 1, " selected=""selected""", "") & ">")
                                Sendb(IIf(MyCommon.NZ(row.Item("PhraseID"), 0) > 0, Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID), MyCommon.NZ(row.Item("Description"), "")))
                                Send("</option>")
                            Next
                            Send("</select><br />")
                        Else
                        End If
                    End If
                Else
                    Send("<br class=""half"" />")
                    If (BannersEnabled) Then
                        Send(Copient.PhraseLib.Lookup("term.banner", LanguageID) & ": " & MyCommon.SplitNonSpacedString(BannerName, 25) & "<br />")
                        Send("<input type=""hidden"" name=""banner"" id=""banner"" value=""" & BannerID & """ />")
                    End If
                    Send(Copient.PhraseLib.Lookup("term.promotionengine", LanguageID) & ": " & EngineName & "<br />")
                End If

                If CreatedDate = Nothing Then
                Else
                    Sendb("   " & Copient.PhraseLib.Lookup("term.created", LanguageID) & " ")
                    longDate = CreatedDate
                    longDateString = Logix.ToLongDateTimeString(longDate, MyCommon)
                    Sendb(longDateString)
                    Send("<br />")
                End If

                If LastUpdate = Nothing Then
                Else
                    Sendb("   " & Copient.PhraseLib.Lookup("term.edited", LanguageID) & " ")
                    longDate = LastUpdate
                    longDateString = Logix.ToLongDateTimeString(longDate, MyCommon)
                    Sendb(longDateString)
                    Send("<br />")
                End If
                Send("<br class=""half"" />")

                MyCommon.QueryStr = "select count(*) as GCount from LocGroupItems where LocationGroupID = " & GroupId & " And Deleted = 0"
                rst2 = MyCommon.LRT_Select()
                For Each row2 In rst2.Rows
                    If (GroupId > 1) Then
                        Sendb(Copient.PhraseLib.Lookup("term.contains", LanguageID) & " ")
                        Sendb(row2.Item("GCount") & " ")
                        If row2.Item("GCount") = 1 Then
                            Send(StrConv(Copient.PhraseLib.Lookup("term.store", LanguageID), VbStrConv.Lowercase))
                        Else
                            Send(StrConv(Copient.PhraseLib.Lookup("term.stores", LanguageID), VbStrConv.Lowercase))
                        End If
                    End If
                Next
            %>
            <hr class="hidden" />
        </div>
        <div class="box" id="members" <%if(groupid = 0)then send(" style=""visibility: hidden;""") %>>
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.contents", LanguageID))%></span></h2>
            <input type="radio" id="memberradio1" name="memberradio" <% if(MyCommon.Fetch_SystemOption(175)= "1" ) then sendb(" checked=""checked""") %> /><label
                for="memberradio1"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
            <input type="radio" id="memberradio2" name="memberradio" <% if(MyCommon.Fetch_SystemOption(175)= "2" ) then sendb(" checked=""checked""") %> /><label
                for="memberradio2"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
            <input type="text" class="medium" id="memberinput" name="memberinput" onkeydown="handleMemberKeyDown(event);"
                onkeyup="handleMemberKeyUp(200);" maxlength="100" value="" /><br />
            <select class="<% Sendb(IIf(CreatedFromOffer, "long", "longest")) %>" id="contents"
                name="contents" multiple="multiple" size="8" style="height: 176px;">
                <%
                    If LocAssignedCount > 0 Then
                        i = 0
                        For Each dr In dtLocAssigned.Rows
                            Send("    <option value=""" & dr.Item("Id") & """>" & dr.Item("Code") & " - " & dr.Item("Name") & "</option>")
                        Next
                    End If
                %>
            </select>
            <br />
            <%
                If (Logix.UserRoles.EditStoreGroups = True) Then
                    Sendb("    <input type=""submit"" id=""store-rem"" name=""store-rem"" style=""margin: 2px 0px 2px 2px;"" value=""" & Copient.PhraseLib.Lookup("lgroup-edit.removeselected", LanguageID) & """ ")
                    If LocAssignedCount > 0 and BrickAndMortarLocationId = 0 Then
                        Sendb(" onclick=""return checkForStoreChange(this.name);"" /><br />")
                    Else
                        Sendb(" disabled=""disabled"" /><br />")
                    End If
                End If
            %>
            <hr class="hidden" />
        </div>
    </div>
    <div id="gutter">
    </div>
    <div id="column2">
        <% If (BrickAndMortarLocationId = 0) Then%>
        <div class="box" id="listadd" <%if(groupid = 0)then send(" style=""visibility: hidden;""") %>>
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("lgroup-edit.addfromlist", LanguageID))%></span></h2>
            <br class="half" />
            <div>
                <input type="button" class="large" id="hierarchy" name="hierarchy" style="width: auto;"
                    value="<%Sendb(Copient.PhraseLib.Lookup("pgroup-edit.modifyusinghierarchy", LanguageID) & "...")%>"
                    onclick="launchHierarchy();" />
                <% If (ShowViewSelected) Then%>
                <input type="button" class="large" id="btnNodes" name="btnNodes" value="<%Sendb(Copient.PhraseLib.Lookup("pgroup-edit.viewselected", LanguageID) & "...")%>"
                    onclick="launchNodes();" />
                <% End If%>
            </div>
            <br class="clear" />
            <hr />
            <input type="radio" id="functionradio1" name="functionradio" <% if(MyCommon.Fetch_SystemOption(175)= "1" ) then sendb(" checked=""checked""") %> /><label
                for="functionradio1"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
            <input type="radio" id="functionradio2" name="functionradio" <% if(MyCommon.Fetch_SystemOption(175)= "2" ) then sendb(" checked=""checked""") %> /><label
                for="functionradio2"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
            <input type="text" class="medium" id="functioninput" name="functioninput" onkeydown="handleKeyDown(event);"
                onkeyup="handleKeyUp(200);" value="" /><br />
            <select class="<% Sendb(IIf(CreatedFromOffer, "long", "longest")) %>" id="fulllist"
                name="fulllist" multiple="multiple" size="8" style="height: 200px;">
            </select>
            <br />
            <% If (Logix.UserRoles.EditStoreGroups = True) Then%>
            <input type="submit" class="large" id="store-add2" name="store-add2" onclick="return storeSelectedToAdd(this.name);"
                style="margin: 2px 0px 2px 2px; width: auto" value="<%Sendb(Copient.PhraseLib.Lookup("lgroup-edit.addselected", LanguageID))%>" />
            <% End If%>
            <hr class="hidden" />
        </div>
        <% End If%>
        <% If Not CreatedFromOffer Then%>
        <div class="box" id="offers" <%if(groupid = 0)then send(" style=""visibility: hidden;""") %>>
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.recently", LanguageID) & " " & Copient.PhraseLib.Lookup("term.associatedoffers", LanguageID).ToLower())%></span></h2>
            <div class="boxscroll">
                <%
                    If(bEnableRestrictedAccessToUEOfferBuilder) Then
                        conditionalQuery=GetRestrictedAccessToUEBuilderQuery(MyCommon,Logix,"I")
                    End If
            
                    Dim OfferCount As Integer = 0

                    MyCommon.QueryStr = "select LocationGroupID from LocationGroups with (NoLock) where AllLocations=1 and Deleted=0;"
                    rst = MyCommon.LRT_Select
                    If rst.Rows.Count > 0 Then
                        AllLocGroupID = MyCommon.NZ(rst.Rows(0).Item("LocationGroupID"), 0)
                    End If

                    MyCommon.QueryStr = "select EngineID from LocationGroups with (NoLock) where LocationGroupID=" & GroupId
                    rst = MyCommon.LRT_Select
                    If rst.Rows.Count > 0 Then
                        LocEngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), -1)
                    End If
                    If LocEngineID >= 0 Then
                        If LocEngineID = 0 Then
                            If (BannersEnabled) Then
                                MyCommon.QueryStr = "select top 500 OL.PKID, 1 as EngineID, O.Name as Name,O.OfferID as OfferID,O.ProdEndDate,NULL as BuyerID from OfferLocations as OL " & _
                                                    "left join Offers as O on O.OfferID=OL.OfferID left join BannerOffers as BO on BO.OfferID=OL.OfferID " & _
                                                    "where O.Deleted=0 and O.IsTemplate=0 and OL.Deleted=0 and OL.LocationGroupID in (" & GroupId & ") and O.EngineID=" & LocEngineID & _
                                                    "order by OL.LocationGroupID desc;"
                            Else
                                MyCommon.QueryStr = "select top 500 OL.PKID, 1 as EngineID, O.Name as Name,O.OfferID as OfferID,O.ProdEndDate,NULL as BuyerID from OfferLocations as OL left join Offers as O on O.OfferID=OL.OfferID " & _
                                                  "where O.Deleted=0 and O.IsTemplate=0 and OL.Deleted=0 and OL.LocationGroupID in (" & GroupId & "," & AllLocGroupID & ") and O.EngineID=" & LocEngineID & _
                                                  "order by OL.LocationGroupID desc;"
                            End If
                        Else
                            MyCommon.QueryStr = "select top 500 OL.PKID, OI.EngineID, I.IncentiveName as Name, I.IncentiveID as OfferID, I.EndDate as ProdEndDate,buy.ExternalBuyerId as BuyerID from OfferLocations OL with (NoLock) " & _
                                                "inner join CPE_Incentives I with (NoLock) on OL.OfferID = I.IncentiveID " & _
                                                "inner join OfferIDs OI with (NoLock) on OI.OfferID = I.IncentiveID " & _
                                                 "left outer join Buyers as buy with (nolock) on buy.BuyerId= I.BuyerId " & _
                                                "where OL.LocationGroupID in (" & GroupId & "," & AllLocGroupID & ") and OL.Deleted=0 and I.Deleted=0 and I.IsTemplate=0  and I.EngineId =  " & LocEngineID & ""
              If(bEnableRestrictedAccessToUEOfferBuilder AndAlso Not String.IsNullOrEmpty(conditionalQuery)) Then MyCommon.QueryStr &=conditionalQuery & " "                
                            MyCommon.QueryStr &= "order by OL.LocationGroupID desc;"

                        End If
                    Else
                        MyCommon.QueryStr = "Select top 500 * from (" & _
                                          "select OL.PKID, 1 as EngineID, O.Name as Name,O.OfferID as OfferID,O.ProdEndDate,NULL as BuyerID from OfferLocations as OL left join Offers as O on O.OfferID=OL.OfferID " & _
                                          " where O.Deleted=0 and O.IsTemplate=0 and OL.Deleted=0 and OL.LocationGroupID in (" & GroupId & "," & AllLocGroupID & ") " & _
                                          " union " & _
                                          "select OL.PKID, OI.EngineID, I.IncentiveName as Name, I.IncentiveID as OfferID, I.EndDate as ProdEndDate,buy.ExternalBuyerId as BuyerID from OfferLocations OL with (NoLock) " & _
                                          "inner join CPE_Incentives I with (NoLock) on OL.OfferID = I.IncentiveID " & _
                                          "inner join OfferIDs OI with (NoLock) on OI.OfferID = I.IncentiveID " & _
                                           "left outer join Buyers as buy with (nolock) on buy.BuyerId= I.BuyerId " & _
                                          "where OL.LocationGroupID in (" & GroupId & "," & AllLocGroupID & ") and OL.Deleted=0 and I.Deleted=0 and I.IsTemplate=0 and I.EngineId =  " & LocEngineID & ""
            If(bEnableRestrictedAccessToUEOfferBuilder AndAlso Not String.IsNullOrEmpty(conditionalQuery)) Then MyCommon.QueryStr &=conditionalQuery & " "                
                        MyCommon.QueryStr &= ") Table1 " & _
                                          "order by PKID desc;"
                    End If

                    rst = MyCommon.LRT_Select
                    OfferCount = rst.Rows.Count
                    Dim assocName As String = ""
                    If (OfferCount > 0) Then
                        For Each row In rst.Rows
                            If (MyCommon.Fetch_UE_SystemOption(168) = "1" AndAlso MyCommon.NZ(row.Item("BuyerID"), "") <> "") Then
                                assocName = "Buyer " + row.Item("BuyerID").ToString() + " - " + MyCommon.NZ(row.Item("Name"), "").ToString()
                            Else
                                assocName = MyCommon.NZ(row.Item("Name"), Copient.PhraseLib.Lookup("term.unknown", LanguageID))
                            End If
                            Sendb("  <div style=""float:left;padding-top:3px;"">")

                            If (Logix.IsAccessibleOffer(AdminUserID, row.Item("OfferID"))) Then
                                Sendb(" <a href=""offer-redirect.aspx?OfferID=" & row.Item("OfferID") & """>" & assocName & "</a>")
                            Else
                                Sendb(assocName)
                            End If

                            If (MyCommon.NZ(row.Item("ProdEndDate"), Today) < Today) Then Sendb(" (" & Copient.PhraseLib.Lookup("term.expired", LanguageID) & ")")
                            Send("</div><br clear=""all"" />")
                        Next
                    End If
                %>
            </div>
            <br class="half" />
            <% If (Logix.UserRoles.EditStoreGroups = True) Then%>
            <input type="button" class="large" id="updateOffers" style="width: auto" name="updateOffers"
                onclick="showOffersWindow();" value="<% Sendb(Copient.PhraseLib.Lookup("term.updateoffers", LanguageID)) %>..." />
            <% End If%>
            <hr class="hidden" />
        </div>
        <% End If%>
    </div>
    <br clear="all" />
</div>
</form>
<div id="uploader" style="display: none;">
    <div id="uploadwrap">
        <div class="box" id="uploadbox">
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.upload", LanguageID))%>
                </span>
            </h2>
            <form action="lgroup-edit.aspx<%Sendb(IIf(CreatedFromOffer, "?OfferID=" & OfferID & "&EngineID=" & EngineID & "&slct=" & Request.QueryString("slct") & "&ex=" & Request.QueryString("ex"), "")) %>"
            id="uploadform" name="uploadform" onsubmit="showWait(); " method="post" enctype="multipart/form-data">
            <%
                Sendb("<input type=""button"" class=""ex"" id=""uploaderclose"" name=""uploaderclose"" value=""X"" style=""float:right;position:relative;top:-27px;"" ")
                Send("onclick=""javascript:document.getElementById('uploader').style.display='none';"" />")
                If (operateAtEnterprise) Then
                    Sendb(Copient.PhraseLib.Lookup("lgroup-edit.uploadtime", LanguageID))
                Else
                    Sendb(Copient.PhraseLib.Lookup("lgroup-edit.upload", LanguageID))
                End If

                Send("<br /><br />")
                Sendb("<input type=""radio"" name=""operation"" id=""operation1"" value=""0"" checked=""checked"" />")
                Send("<label for=""operation1"">" & Copient.PhraseLib.Lookup("term.FullReplace", LanguageID) & "</label>&nbsp;&nbsp;")
                Sendb("<input type=""radio"" name=""operation"" id=""operation2"" value=""1""  />")
                Send("<label for=""operation2"">" & Copient.PhraseLib.Lookup("term.AddToGroup", LanguageID) & "</label>&nbsp;&nbsp;")
                Sendb("<input type=""radio"" name=""operation"" id=""operation3"" value=""2""  />")
                Send("<label for=""operation3"">" & Copient.PhraseLib.Lookup("term.removefromgroup", LanguageID) & "</label>")
                Send("<br />")
            %>
            <br />
            <br class="half" />
            <%
                If (Logix.UserRoles.EditStoreGroups) Then
                    Send("     <input type=""hidden"" id=""LocationGroupID2"" name=""LocationGroupID"" value=""" & GroupId & """ />")
                    Send("     <input type=""file"" id=""browse"" name=""browse"" value=""" & Copient.PhraseLib.Lookup("term.browse", LanguageID) & """ />")
                    '         Send("<div id=""divfile"" style=""height:0px;overflow:hidden"">")
                    'Send("<input type=""file"" id=""browse"" name=""fileInput"" onchange=""fileonclick()"" />")
                    'Send("</div>")
                    'Send("<button type=""button"" onclick=""chooseFile();"">"&Copient.PhraseLib.Lookup("term.browse", LanguageID) &"</button>")
                    'Send("<label id=""lblfileupload"" name=""lblfileupload"">"&Copient.PhraseLib.Lookup("term.nofilesselected", LanguageID)&"</label>")
                    Send("     <input type=""submit"" class=""regular"" id=""uploadfile"" name=""uploadfile"" value=""" & Copient.PhraseLib.Lookup("term.upload", LanguageID) & """ />")
                    Send("     <br />")
                End If
            %>
            </form>
            <hr class="hidden" />
        </div>
    </div>
    <%
        If Request.Browser.Type = "IE6" Then
            Send("<iframe src=""javascript:'';"" id=""uploadiframe-pg"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no""></iframe>")
        End If
    %>
</div>
<div id="loading" style="display: none">
    <div id="loadingwrap" style="display: block">
        <div class="box" id="loadingbox">
            <form action="lgroup-edit.aspx" id="WaitForm" method="get">
            <div class="loading" style="display: block">
                <br />
                <img alt="<% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID)) %>" id="clock"
                    src="../images/clock22.png" />
                <br />
                <% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID))%>
            </div>
            </form>
            <hr class="hidden" />
        </div>
    </div>
</div>
<script type="text/javascript">
    if (window.captureEvents) {
        window.captureEvents(Event.CLICK);
        window.onclick = handlePageClick;
    }
    else {
        document.onclick = handlePageClick;
    }
    handleKeyUp(9999);
</script>
<%
    If MyCommon.Fetch_SystemOption(75) AndAlso (Deleted = False) Then
        If (GroupId > 0 And Logix.UserRoles.AccessNotes) Then
            Send_Notes(14, GroupId, AdminUserID)
        End If
    End If
    Send_BodyEnd("mainform", "GroupName")
done:
    ' Catch ex As Exception
    ' MyCommon.Error_Processor("Catch", ex.Message, "lgroup-edit.aspx", "Locations")
    ' Throw ex
Finally
    MyCommon.Close_LogixRT()
End Try
MyCommon = Nothing
Logix = Nothing
%>