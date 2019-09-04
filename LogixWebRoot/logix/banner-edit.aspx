<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
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
    ' * FILENAME: banner-edit.aspx 
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
    Dim WebClient As System.Net.WebClient

    Dim AdminUserID As Long
    Dim MyCommon As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    Dim rst, dt As DataTable
    Dim dtOffers As DataTable = Nothing
    Dim dtGroups As DataTable = Nothing
    Dim dtStores As DataTable = Nothing
    Dim row As DataRow
    Dim BannerID As Integer = -1
    Dim BannerDescription As String = ""
    Dim BannerName As String = ""
    Dim BannerNameTitle As String = ""
    Dim EngineID As Integer = -1
    Dim CreatedDate As Date
    Dim LastUpdate As Date
    Dim statusMessage As String = ""
    Dim infoMessage As String = ""
    Dim ShowActionButton As Boolean = False
    Dim Handheld As Boolean = False
    Dim OptionSelected As Boolean = False
    Dim IsDeleted As Boolean = False
    Dim UserTextFull As String = ""
    Dim UserTextDisplay As String = ""
    Dim BannerUserIDs As String() = Nothing
    Dim i As Integer
    Dim ExistingUserIDs As String()
    Dim NewUserIDs As String()
    Dim AddedUsers As New ArrayList(10)
    Dim RemovedUsers As New ArrayList(10)
    Dim UserStrArray As String()
    Dim IsAllBanners As Boolean = False
    Dim BannerInUse As Boolean = False
    Dim CustomerGroupID As Integer = 0
    Dim AllBannersUsed As Boolean = False
    Dim DefaultChargeback As Integer = 0
    Dim LocHierarchyID As Integer = 0
    Dim ProdHierarchyID As Integer = 0
    Dim TenderDeptID As Integer = 0
    Dim DisplayStr As String = ""
    Dim DeptID As Integer = 0
    Dim DeptXID As String = ""
    Dim DeptName As String = ""
    Dim IsDefaultDept As Boolean = False
    Dim EngineType As Integer = -1
    Dim TerminalID As Integer = 0
    Dim TerminalName As String = ""
    Dim ExtTerminalCode As String = ""
    Dim IE6ScrollFix As String = ""
    Dim ExtBannerID As String = ""
    Dim IsDefaultBanner As Boolean = False

    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If


    Response.Expires = 0
    MyCommon.AppName = "banner-edit.aspx"
    MyCommon.Open_LogixRT()
    MyCommon.Open_LogixXS()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)

    If (Request.QueryString("infoMessage") <> "") Then
        infoMessage = Request.QueryString("infoMessage")
    End If

    BannerID = IIf(Request.QueryString("BannerID") <> "", MyCommon.Extract_Val(Request.QueryString("BannerID")), -1)

    If (Request.QueryString("new") = "") Then
        BannerName = MyCommon.Parse_Quotes(Request.QueryString.Item("name"))
        BannerName = Logix.TrimAll(BannerName)
        BannerDescription = MyCommon.Parse_Quotes(Request.QueryString.Item("desc"))
        EngineID = IIf(Request.QueryString("engine") <> "", MyCommon.Extract_Val(Request.QueryString("engine")), -1)
        IsAllBanners = (Request.QueryString("allbanners") = "true")
        IsDefaultBanner = (Request.QueryString("defaultbanner") = "true")
    End If

    MyCommon.QueryStr = "select Top 500 OfferID, OfferName from " & _
                        "(select BO.OfferID, O.Name as OfferName from BannerOffers BO with (NoLock) " & _
                        "   inner join Offers O with (NoLock) on O.OfferID = BO.OfferID where BannerID=" & BannerID & " " & _
                        "union " & _
                        "select BO.OfferID, I.IncentiveName as OfferName from BannerOffers BO with (NoLock) " & _
                        "   inner join CPE_Incentives I with (NoLock) on I.IncentiveID = BO.OfferID where BannerID=" & BannerID & " " & _
                        ") as OfferList Order By OfferID desc;"
    dtOffers = MyCommon.LRT_Select

    MyCommon.QueryStr = "select LocationGroupID, Name from LocationGroups where BannerID=" & BannerID & " and Deleted=0;"
    dtGroups = MyCommon.LRT_Select

    MyCommon.QueryStr = "select LocationID, LocationName from Locations where BannerID=" & BannerID & " and Deleted=0;"
    dtStores = MyCommon.LRT_Select

    BannerInUse = (dtOffers.Rows.Count > 0 OrElse dtGroups.Rows.Count > 0 OrElse dtStores.Rows.Count > 0)

    ' any GET parms inbound?
    If (Request.QueryString("Delete") <> "") Then
        ' check if any offers are in process for this banner, if so then no deleting of the banner is permitted

        If (BannerInUse) Then
            infoMessage = Copient.PhraseLib.Lookup("banner-edit.inuse", LanguageID)
        Else
            If (BannerID > 0) Then
                MyCommon.QueryStr = "update Banners with (RowLock) set Deleted=1 where BannerID=" & BannerID & ";"
                MyCommon.LRT_Execute()

                MyCommon.QueryStr = "delete from BannerEngines with (RowLock) where BannerID = " & BannerID & ";"
                MyCommon.LRT_Execute()

                MyCommon.QueryStr = "delete from AdminUserBanners with (RowLock) where BannerID=" & BannerID & ";"
                MyCommon.LRT_Execute()

                ' find the associated all banner cardholder customer group and delete it
                If (BannerID > 0) Then
                    MyCommon.QueryStr = "select CustomerGroupID from CustomerGroups with (NoLock) where BannerID=" & BannerID
                    rst = MyCommon.LRT_Select
                    If (rst.Rows.Count > 0) Then
                        For Each row In rst.Rows
                            CustomerGroupID = MyCommon.NZ(row.Item("CustomerGroupID"), -1)
                            MyCommon.QueryStr = "dbo.pt_CustomerGroups_Delete"
                            MyCommon.Open_LRTsp()
                            MyCommon.LRTsp.Parameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = CustomerGroupID
                            MyCommon.LRTsp.ExecuteNonQuery()
                            MyCommon.Close_LRTsp()
                            MyCommon.Activity_Log(4, CustomerGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.cgroup-delete", LanguageID))
                        Next
                    End If
                End If

                MyCommon.QueryStr = "delete from ChargebackDepts with (RowLock) where BannerID = " & BannerID & ";"
                MyCommon.LRT_Execute()

                MyCommon.QueryStr = "delete from TerminalTypes with (RowLock) where BannerID = " & BannerID & ";"
                MyCommon.LRT_Execute()

                'Record history
                MyCommon.Activity_Log(28, BannerID, AdminUserID, Copient.PhraseLib.Lookup("history.banner-delete", LanguageID))
                Response.Status = "301 Moved Permanently"
                Response.AddHeader("Location", "banner-list.aspx")
                GoTo done
            End If
        End If

    ElseIf (Request.QueryString("BannerID") = "new") Then
        ' add a record
        BannerName = Request.QueryString.Item("name")
        BannerName = Logix.TrimAll(BannerName)
        BannerDescription = Request.QueryString.Item("desc")
        EngineID = IIf(Request.QueryString("engine") <> "", MyCommon.Extract_Val(Request.QueryString("engine")), -1)
        IsAllBanners = (Request.QueryString("allbanners") = "true")
        ExtBannerID = Logix.TrimAll(Request.QueryString("xid"))
        IsDefaultBanner = (Request.QueryString("defaultbanner") = "true")

        If (BannerName = "") Then
            infoMessage = Copient.PhraseLib.Lookup("banner-edit.noname", LanguageID)
            BannerID = -1
        Else
            MyCommon.QueryStr = "select BannerID from Banners with (NoLock) where Name = '" & MyCommon.Parse_Quotes(BannerName) & "' and Deleted=0;"
            rst = MyCommon.LRT_Select

            If (rst.Rows.Count > 0) Then
                infoMessage = Copient.PhraseLib.Lookup("banner-edit.nameused", LanguageID)
                BannerID = -1
            Else
                If (IsAllBanners AndAlso AllBannersExists(EngineID, BannerID, MyCommon)) Then
                    infoMessage = Copient.PhraseLib.Lookup("banner-edit.all-banners-exist", LanguageID)
                    BannerID = -1
                Else
                    MyCommon.QueryStr = "dbo.pt_Banner_Insert"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 100).Value = BannerName
                    MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 400).Value = BannerDescription
                    MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = EngineID
                    MyCommon.LRTsp.Parameters.Add("@AllBanners", SqlDbType.Bit).Value = IIf(IsAllBanners, 1, 0)
                    MyCommon.LRTsp.Parameters.Add("@ExtBannerID", SqlDbType.NVarChar, 50).Value = ExtBannerID
                    MyCommon.LRTsp.Parameters.Add("@DefaultBanner", SqlDbType.Int).Value = IIf(IsDefaultBanner, 1, 0)
                    MyCommon.LRTsp.Parameters.Add("@BannerID", SqlDbType.Int).Direction = ParameterDirection.Output
                    MyCommon.LRTsp.ExecuteNonQuery()
                    BannerID = MyCommon.LRTsp.Parameters("@BannerID").Value
                    MyCommon.Close_LRTsp()
                    MyCommon.Activity_Log(28, BannerID, AdminUserID, Copient.PhraseLib.Lookup("history.banner-create", LanguageID))

                    ' create a tender chargeback department for this banner
                    MyCommon.QueryStr = "dbo.pt_ChargebackDept_Insert"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 200).Value = Copient.PhraseLib.Lookup("term.tender", LanguageID)
                    MyCommon.LRTsp.Parameters.Add("@ExternalID", SqlDbType.NVarChar, 20).Value = "0000"
                    MyCommon.LRTsp.Parameters.Add("@BannerID", SqlDbType.Int).Value = BannerID
                    MyCommon.LRTsp.Parameters.Add("@ChargebackDeptID", SqlDbType.Int).Direction = ParameterDirection.Output
                    MyCommon.LRTsp.ExecuteNonQuery()
                    TenderDeptID = MyCommon.LRTsp.Parameters("@ChargebackDeptID").Value
                    MyCommon.Activity_Log(17, TenderDeptID, AdminUserID, Copient.PhraseLib.Lookup("history.department-create", LanguageID))

                    '' create a customer group for all cardholders in this banner
                    'MyCommon.QueryStr = "dbo.pt_CustomerGroups_Insert"
                    'MyCommon.Open_LRTsp()
                    'MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 200).Value = Copient.PhraseLib.Lookup("banner-edit.anycardholder", LanguageID) & " " & BannerName
                    'MyCommon.LRTsp.Parameters.Add("@CustomerGroupID", SqlDbType.Int).Direction = ParameterDirection.Output
                    'MyCommon.LRTsp.ExecuteNonQuery()
                    'CustomerGroupID = MyCommon.LRTsp.Parameters("@CustomerGroupID").Value
                    'MyCommon.Close_LRTsp()
                    'MyCommon.Activity_Log(4, CustomerGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.cgroup-create", LanguageID))

                    'MyCommon.QueryStr = "update CustomerGroups with (RowLock) set BannerID = " & BannerID & " where CustomerGroupID=" & CustomerGroupID
                    'MyCommon.LRT_Execute()

                    '' create a new cardholders customer group for this banner
                    'MyCommon.QueryStr = "dbo.pt_CustomerGroups_Insert"
                    'MyCommon.Open_LRTsp()
                    'MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 200).Value = Copient.PhraseLib.Lookup("banner-edit.newcardholders", LanguageID) & " " & BannerName
                    'MyCommon.LRTsp.Parameters.Add("@CustomerGroupID", SqlDbType.Int).Direction = ParameterDirection.Output
                    'MyCommon.LRTsp.ExecuteNonQuery()
                    'CustomerGroupID = MyCommon.LRTsp.Parameters("@CustomerGroupID").Value
                    'MyCommon.Close_LRTsp()
                    'MyCommon.Activity_Log(4, CustomerGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.cgroup-create", LanguageID))

                    'MyCommon.QueryStr = "update CustomerGroups with (RowLock) set BannerID = " & BannerID & ", NewCardholders=1 where CustomerGroupID = " & CustomerGroupID
                    'MyCommon.LRT_Execute()
                End If
            End If
        End If
    ElseIf ((Request.QueryString("save") <> "" OrElse Request.QueryString("select") <> "" OrElse Request.QueryString("deselect") <> "") AndAlso Request.QueryString("BannerID") <> "") Then
        ' somebody clicked save
        BannerID = MyCommon.Extract_Val(Request.QueryString("BannerID"))
        BannerName = Request.QueryString.Item("name")
        BannerName = Logix.TrimAll(BannerName)
        BannerDescription = Request.QueryString.Item("desc")
        IsAllBanners = (Request.QueryString("allbanners") = "true")
        ExtBannerID = Logix.TrimAll(Request.QueryString("xid"))
        IsDefaultBanner = (Request.QueryString("defaultbanner") = "true")

        If (Request.QueryString("name") = "") Then
            infoMessage = Copient.PhraseLib.Lookup("banner-edit.noname", LanguageID)
        Else
            MyCommon.QueryStr = "select BannerID from Banners with (NoLock) where Name = '" & MyCommon.Parse_Quotes(BannerName) & "' and Deleted=0 and BannerID <> " & BannerID & ";"
            rst = MyCommon.LRT_Select

            If (rst.Rows.Count > 0) Then
                infoMessage = Copient.PhraseLib.Lookup("banner-edit.nameused", LanguageID)
            Else
                If (IsAllBanners AndAlso AllBannersExists(EngineID, BannerID, MyCommon)) Then
                    infoMessage = Copient.PhraseLib.Lookup("banner-edit.all-banners-exist", LanguageID)
                Else
                    MyCommon.QueryStr = "dbo.pt_Banner_Update"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@BannerID", SqlDbType.Int).Value = BannerID
                    MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 100).Value = BannerName
                    MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 400).Value = BannerDescription
                    MyCommon.LRTsp.Parameters.Add("@AllBanners", SqlDbType.Bit).Value = IIf(IsAllBanners, 1, 0)
                    MyCommon.LRTsp.Parameters.Add("@DefaultBanner", SqlDbType.Int).Value = IIf(IsDefaultBanner, 1, 0)
                    MyCommon.LRTsp.Parameters.Add("@ExtBannerID", SqlDbType.NVarChar, 50).Value = ExtBannerID
                    MyCommon.LRTsp.ExecuteNonQuery()
                    MyCommon.Close_LRTsp()

                    If (Request.QueryString("save") <> "") Then
                        MyCommon.Activity_Log(28, BannerID, AdminUserID, Copient.PhraseLib.Lookup("history.banner-edit", LanguageID))
                    End If

                End If
            End If

            ' handle add and remove users to banner
            If (Request.QueryString("userschanged") = "true") Then
                ExistingUserIDs = Request.QueryString("existinguserids").Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
                Array.Sort(ExistingUserIDs)
                NewUserIDs = Request.QueryString("newuserids").Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
                Array.Sort(NewUserIDs)
                AddedUsers = GetAddedUsers(ExistingUserIDs, NewUserIDs)
                RemovedUsers = GetRemovedUsers(ExistingUserIDs, NewUserIDs)

                If (AddedUsers.Count > 0) Then
                    If (Logix.UserRoles.AddUsersToBanners) Then
                        For i = 0 To AddedUsers.Count - 1
                            MyCommon.QueryStr = "insert into AdminUserBanners (AdminUserID, BannerID) values (" & AddedUsers(i) & ", " & BannerID & ");"
                            MyCommon.LRT_Execute()
                        Next
                        UserStrArray = AddedUsers.ToArray(System.Type.GetType("System.String"))
                        MyCommon.Activity_Log(28, BannerID, AdminUserID, Copient.PhraseLib.Lookup("history.banner-addusers", LanguageID) & String.Join(", ", UserStrArray))
                    Else
                        infoMessage = Copient.PhraseLib.Lookup("error.permission", LanguageID)
                    End If
                End If

                If (RemovedUsers.Count > 0) Then
                    If (Logix.UserRoles.RemoveUsersFromBanners) Then
                        For i = 0 To RemovedUsers.Count - 1
                            MyCommon.QueryStr = "delete from AdminUserBanners where AdminUserID = " & RemovedUsers(i) & " and BannerID = " & BannerID & ";"
                            MyCommon.LRT_Execute()
                        Next
                        UserStrArray = RemovedUsers.ToArray(System.Type.GetType("System.String"))
                        MyCommon.Activity_Log(28, BannerID, AdminUserID, Copient.PhraseLib.Lookup("history.banner-removeusers", LanguageID) & String.Join(", ", UserStrArray))
                    Else
                        infoMessage = Copient.PhraseLib.Lookup("error.permission", LanguageID)
                    End If
                End If

            End If

        End If

    ElseIf (Request.QueryString("lochierarchy-select") <> "") Then
        ' update the location hierarchy associations with selected hierarchies
        If (Request.QueryString("lochierarchy-available") <> "") Then
            For i = 0 To Request.QueryString.GetValues("lochierarchy-available").GetUpperBound(0)
                LocHierarchyID = IIf(Request.QueryString.GetValues("lochierarchy-available")(i) <> "", Request.QueryString.GetValues("lochierarchy-available")(i), -1)
                If (LocHierarchyID > -1) Then
                    MyCommon.QueryStr = "select HierarchyID from BannerLocHierarchies with (NoLock) where BannerID=" & BannerID & " and HierarchyID=" & LocHierarchyID & ";"
                    rst = MyCommon.LRT_Select
                    If (rst.Rows.Count = 0) Then
                        MyCommon.QueryStr = "insert into BannerLocHierarchies with (RowLock) (BannerID, HierarchyID) values (" & BannerID & "," & LocHierarchyID & ");"
                    End If
                    MyCommon.LRT_Execute()
                End If
            Next
        End If
    ElseIf (Request.QueryString("lochierarchy-deselect") <> "") Then
        ' update the location hierarchy associations to remove deselected hierarchies
        If (Request.QueryString("lochierarchy-selected") <> "") Then
            For i = 0 To Request.QueryString.GetValues("lochierarchy-selected").GetUpperBound(0)
                LocHierarchyID = IIf(Request.QueryString.GetValues("lochierarchy-selected")(i) <> "", Request.QueryString.GetValues("lochierarchy-selected")(i), -1)
                If (LocHierarchyID > -1) Then
                    MyCommon.QueryStr = "select HierarchyID from BannerLocHierarchies with (NoLock) where BannerID=" & BannerID & " and HierarchyID=" & LocHierarchyID & ";"
                    rst = MyCommon.LRT_Select
                    If (rst.Rows.Count > 0) Then
                        MyCommon.QueryStr = "delete from BannerLocHierarchies with (RowLock) where BannerID=" & BannerID & " and HierarchyID=" & LocHierarchyID & ";"
                    End If
                    MyCommon.LRT_Execute()
                End If
            Next
        End If

    ElseIf (Request.QueryString("prodhierarchy-select") <> "") Then
        ' update the product hierarchy associations with selected hierarchies
        If (Request.QueryString("prodhierarchy-available") <> "") Then
            For i = 0 To Request.QueryString.GetValues("prodhierarchy-available").GetUpperBound(0)
                ProdHierarchyID = IIf(Request.QueryString.GetValues("prodhierarchy-available")(i) <> "", Request.QueryString.GetValues("prodhierarchy-available")(i), -1)
                If (ProdHierarchyID > -1) Then
                    MyCommon.QueryStr = "select HierarchyID from BannerProdHierarchies with (NoLock) where BannerID=" & BannerID & " and HierarchyID=" & ProdHierarchyID & ";"
                    rst = MyCommon.LRT_Select
                    If (rst.Rows.Count = 0) Then
                        MyCommon.QueryStr = "insert into BannerProdHierarchies with (RowLock) (BannerID, HierarchyID) values (" & BannerID & "," & ProdHierarchyID & ");"
                    End If
                    MyCommon.LRT_Execute()
                End If
            Next
        End If
    ElseIf (Request.QueryString("prodhierarchy-deselect") <> "") Then
        ' update the product hierarchy associations to remove deselected hierarchies
        If (Request.QueryString("prodhierarchy-selected") <> "") Then
            For i = 0 To Request.QueryString.GetValues("prodhierarchy-selected").GetUpperBound(0)
                ProdHierarchyID = IIf(Request.QueryString.GetValues("prodhierarchy-selected")(i) <> "", Request.QueryString.GetValues("prodhierarchy-selected")(i), -1)
                If (ProdHierarchyID > -1) Then
                    MyCommon.QueryStr = "select HierarchyID from BannerProdHierarchies with (NoLock) where BannerID=" & BannerID & " and HierarchyID=" & ProdHierarchyID & ";"
                    rst = MyCommon.LRT_Select
                    If (rst.Rows.Count > 0) Then
                        MyCommon.QueryStr = "delete from BannerProdHierarchies with (RowLock) where BannerID=" & BannerID & " and HierarchyID=" & ProdHierarchyID & ";"
                    End If
                    MyCommon.LRT_Execute()
                End If
            Next
        End If

        'ElseIf () Then
        '  ' update the product hierarchy associations
        '  ProdHierarchyID = IIf(Request.QueryString("prodhierarchy") <> "", MyCommon.Extract_Val(Request.QueryString("prodhierarchy")), -1)
        '  If (LocHierarchyID > -1) Then
        '    MyCommon.QueryStr = "select HierarchyID from BannerProdHierarchies with (NoLock) where BannerID=" & BannerID
        '    rst = MyCommon.LRT_Select
        '    If (rst.Rows.Count > 0) Then
        '      MyCommon.QueryStr = "update BannerProdHierarchies set HierarchyID = " & ProdHierarchyID & " where BannerID=" & BannerID
        '    Else
        '      MyCommon.QueryStr = "insert into BannerProdHierarchies (BannerID, HierarchyID) values (" & BannerID & "," & ProdHierarchyID & ");"
        '    End If
        '    MyCommon.LRT_Execute()
        '  End If

    ElseIf (Request.QueryString("adddept") <> "") Then
        BannerID = MyCommon.Extract_Val(Request.QueryString("BannerID"))
        DeptName = Request.QueryString("deptname")
        DeptName = Logix.TrimAll(DeptName)
        DeptXID = Request.QueryString("deptxid")

        MyCommon.QueryStr = "SELECT ChargeBackDeptID FROM ChargeBackDepts with (NoLock) " & _
                            "WHERE Name='" & MyCommon.Parse_Quotes(DeptName) & "' " & _
                            "AND BannerID=" & BannerID & ";"
        rst = MyCommon.LRT_Select

        If (DeptName = "" OrElse DeptXID = "") Then
            infoMessage = Copient.PhraseLib.Lookup("departments.noname", LanguageID)
        ElseIf (DeptXID = "0000") Then
            infoMessage = Copient.PhraseLib.Lookup("departments.numberused", LanguageID)
        ElseIf (CleanUPC(DeptXID) = "False") Then
            infoMessage = Copient.PhraseLib.Lookup("departments.badcode", LanguageID)
        Else
            If (rst.Rows.Count > 0) Then
                infoMessage = Copient.PhraseLib.Lookup("departments.nameused", LanguageID)
            Else
                MyCommon.QueryStr = "SELECT ChargeBackDeptID FROM ChargeBackDepts with (NoLock) WHERE ExternalID = '" & DeptXID & "' " & _
                                    "and BannerID=" & BannerID
                rst = MyCommon.LRT_Select
                If (rst.Rows.Count > 0) Then
                    infoMessage = Copient.PhraseLib.Lookup("departments.numberused", LanguageID)
                Else
                    MyCommon.QueryStr = "dbo.pt_ChargebackDept_Insert"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 200).Value = DeptName
                    MyCommon.LRTsp.Parameters.Add("@ExternalID", SqlDbType.NVarChar, 120).Value = DeptXID
                    MyCommon.LRTsp.Parameters.Add("@BannerID", SqlDbType.Int).Value = BannerID
                    MyCommon.LRTsp.Parameters.Add("@ChargebackDeptID", SqlDbType.Int).Direction = ParameterDirection.Output
                    MyCommon.LRTsp.ExecuteNonQuery()
                    DeptID = MyCommon.LRTsp.Parameters("@ChargebackDeptID").Value
                    MyCommon.Activity_Log(17, DeptID, AdminUserID, Copient.PhraseLib.Lookup("history.department-create", LanguageID))
                End If
            End If
        End If

    ElseIf (Request.QueryString("removedept") <> "") Then
        BannerID = MyCommon.Extract_Val(Request.QueryString("BannerID"))
        DeptID = MyCommon.Extract_Val(Request.QueryString("dept"))
        If (DeptID = 0) Then
            infoMessage = Copient.PhraseLib.Lookup("departments.nodelete", LanguageID)
        Else
            MyCommon.QueryStr = "select top 1 O.offerid from offers as O with (NoLock) left join offerrewards as OFR with (NoLock) on OFR.offerid=O.offerid " & _
                                "left join discounts as DISC with (NoLock) on DISC.discountid=OFR.linkid " & _
                                "where OFR.rewardtypeid=1 and O.prodenddate>getdate() and " & _
                                "O.deleted=0 and DISC.chargebackdeptid=" & DeptID
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count > 0) Then
                infoMessage = Copient.PhraseLib.Lookup("departments.inuse", LanguageID)
            Else
                MyCommon.QueryStr = "DELETE FROM ChargeBackDepts with (RowLock) WHERE ChargeBackDeptID = " & DeptID
                MyCommon.LRT_Execute()
                MyCommon.Activity_Log(17, DeptID, AdminUserID, Copient.PhraseLib.Lookup("history.department-delete", LanguageID))

                ' If this is the banner's default chargeback department being deleted than unset it for the banner
                MyCommon.QueryStr = "Update Banners set DefaultChargebackDeptID=NULL where BannerID=" & BannerID & " and DefaultChargebackDeptID=" & DeptID
                MyCommon.LRT_Execute()
            End If
        End If

    ElseIf (Request.QueryString("defaultdept") <> "") Then
        BannerID = MyCommon.Extract_Val(Request.QueryString("BannerID"))
        DeptID = MyCommon.Extract_Val(Request.QueryString("dept"))
        MyCommon.QueryStr = "update Banners with (RowLock) set DefaultChargebackDeptID=" & DeptID & " where BannerID=" & BannerID & " and Deleted=0;"
        MyCommon.LRT_Execute()
        MyCommon.Activity_Log(28, BannerID, AdminUserID, Copient.PhraseLib.Lookup("history.banner-edit", LanguageID))

    ElseIf (Request.QueryString("addterm") <> "") Then
        BannerID = MyCommon.Extract_Val(Request.QueryString("BannerID"))
        TerminalName = Request.QueryString("termname")
        TerminalName = Logix.TrimAll(TerminalName)
        ExtTerminalCode = Request.QueryString("termxid")

        MyCommon.QueryStr = "select EngineID from BannerEngines BE with (NoLock) where BannerID = " & BannerID
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
            EngineType = MyCommon.NZ(rst.Rows(0).Item("EngineID"), 0)
        End If

        MyCommon.QueryStr = "dbo.pt_TerminalTypes_Insert"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 100).Value = TerminalName
        MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 1000).Value = ""
        MyCommon.LRTsp.Parameters.Add("@ExtTerminalCode", SqlDbType.NVarChar, 50).Value = ExtTerminalCode
        MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = EngineType
        MyCommon.LRTsp.Parameters.Add("@LayoutID", SqlDbType.Int).Value = 0
        MyCommon.LRTsp.Parameters.Add("@SpecificPromosOnly", SqlDbType.Int).Value = 0
        MyCommon.LRTsp.Parameters.Add("@FuelProcessing", SqlDbType.Bit).Value = 0
        MyCommon.LRTsp.Parameters.Add("@BannerID", SqlDbType.Int).Value = BannerID
        MyCommon.LRTsp.Parameters.Add("@TerminalTypeId", SqlDbType.Int).Direction = ParameterDirection.Output

        ExtTerminalCode = Logix.TrimAll(ExtTerminalCode)

        If (TerminalName = "") Or (ExtTerminalCode = "" And (EngineType = 0 Or EngineType = 1)) Then
            infoMessage = Copient.PhraseLib.Lookup("terminal-edit.noname", LanguageID)
        ElseIf Not (EngineType = 0 Or EngineType = 1 Or EngineType = 4) AndAlso IsNumeric(ExtTerminalCode) Then
            infoMessage = Copient.PhraseLib.Lookup("terminal-edit.badcode", LanguageID)
        ElseIf (EngineType = 0 Or EngineType = 1 Or EngineType = 4) AndAlso (Not IsNumeric(ExtTerminalCode) OrElse ((ExtTerminalCode < 1) Or (Int(ExtTerminalCode) <> ExtTerminalCode))) Then
            infoMessage = Copient.PhraseLib.Lookup("terminal-edit.badcode", LanguageID)
        Else
            MyCommon.QueryStr = "SELECT TerminalTypeID FROM TerminalTypes with (NoLock) WHERE Name = '" & MyCommon.Parse_Quotes(TerminalName) & "' AND EngineID=" & EngineType & " AND AnyTerminal=0 AND Deleted=0 "
            MyCommon.QueryStr &= " and BannerID=" & BannerID

            rst = MyCommon.LRT_Select
            If (rst.Rows.Count > 0) Then
                infoMessage = Copient.PhraseLib.Lookup("terminal-edit.nameused", LanguageID)
            Else
                MyCommon.QueryStr = "SELECT TerminalTypeID FROM TerminalTypes with (NoLock) " & _
                                    "WHERE EngineID In (0,1,4) AND AnyTerminal=0 AND Deleted=0 " & _
                                    "AND ExtTerminalCode='" & ExtTerminalCode & "' AND BannerID=" & BannerID & ";"
                rst = MyCommon.LRT_Select
                If (rst.Rows.Count > 0) And Not (EngineType = 2) Then
                    infoMessage = Copient.PhraseLib.Lookup("terminal-edit.codeused", LanguageID)
                Else
                    MyCommon.LRTsp.ExecuteNonQuery()
                    TerminalID = MyCommon.LRTsp.Parameters("@TerminalTypeId").Value
                    MyCommon.Close_LRTsp()
                    MyCommon.Activity_Log(21, TerminalID, AdminUserID, Copient.PhraseLib.Lookup("history.terminal-create", LanguageID))
                End If
            End If
        End If

    ElseIf (Request.QueryString("removeterm") <> "") Then
        TerminalID = MyCommon.Extract_Val(Request.QueryString("terminal"))
        MyCommon.QueryStr = "select distinct O.offerid,description from offerterminals as OT with (NoLock) left join offers as O with (NoLock) on O.offerid=OT.offerid " & _
                            "where O.deleted=0 and ot.terminalTypeid=" & TerminalID
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
            infoMessage = Copient.PhraseLib.Lookup("terminal-edit.inuse", LanguageID)
        ElseIf TerminalID = 0 Then
            infoMessage = Copient.PhraseLib.Lookup("terminal-edit.noneSelected", LanguageID)
        Else
            If (TerminalID > 0) Then
                MyCommon.QueryStr = "dbo.pt_TerminalTypes_Delete"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@TerminalTypeId", SqlDbType.Int).Value = TerminalID
                MyCommon.LRTsp.ExecuteNonQuery()
                MyCommon.Close_LRTsp()
                MyCommon.Activity_Log(21, TerminalID, AdminUserID, Copient.PhraseLib.Lookup("history.terminal-delete", LanguageID))
            End If
        End If

    ElseIf (Request.QueryString("BannerID") <> "" AndAlso Request.QueryString("new") = "") Then
        ' simple edit/search mode
        If (Request.QueryString("BannerID") <> "") Then
            BannerID = MyCommon.Extract_Val(Request.QueryString("BannerID"))
        Else
            BannerID = -1
        End If

    ElseIf (Request.Form("BannerID") <> "" AndAlso Request.Form("new") = "") Then
        If (Request.QueryString("BannerID") <> "") Then
            BannerID = MyCommon.Extract_Val(Request.Form("BannerID"))
        Else
            BannerID = -1
        End If

    Else
        ' no group id passed ... what now ?
        BannerID = -1
    End If

    ' load the banner
    If (BannerID > -1) Then
        MyCommon.QueryStr = "select BannerID, Name, Description, AllBanners, CreatedDate, LastUpdate, " & _
                            "Deleted, DefaultChargebackDeptID, ExtBannerID " & _
                            "from Banners with (NoLock) where BannerID=" & BannerID
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
            BannerID = MyCommon.NZ(rst.Rows(0).Item("BannerID"), -1)
            BannerName = MyCommon.NZ(rst.Rows(0).Item("Name"), "")
            BannerDescription = MyCommon.NZ(rst.Rows(0).Item("Description"), "")
            IsAllBanners = MyCommon.NZ(rst.Rows(0).Item("AllBanners"), False)
            CreatedDate = MyCommon.NZ(rst.Rows(0).Item("CreatedDate"), "")
            LastUpdate = MyCommon.NZ(rst.Rows(0).Item("LastUpdate"), "")
            IsDeleted = MyCommon.NZ(rst.Rows(0).Item("Deleted"), False)
            If (IsDeleted) Then infoMessage = Copient.PhraseLib.Lookup("error.deleted", LanguageID)
            DefaultChargeback = MyCommon.NZ(rst.Rows(0).Item("DefaultChargebackDeptID"), 0)
            ExtBannerID = MyCommon.NZ(rst.Rows(0).Item("ExtBannerID"), "")
        Else
            BannerName = Copient.PhraseLib.Lookup("term.newprogram", LanguageID)
        End If

        MyCommon.QueryStr = "select HierarchyID from BannerLocHierarchies with (NoLock) where BannerID = " & BannerID
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
            LocHierarchyID = MyCommon.NZ(rst.Rows(0).Item("HierarchyID"), -1)
        End If

        MyCommon.QueryStr = "select HierarchyID from BannerProdHierarchies with (NoLock) where BannerID = " & BannerID
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
            ProdHierarchyID = MyCommon.NZ(rst.Rows(0).Item("HierarchyID"), -1)
        End If

    End If

    Send_HeadBegin("term.banner", , BannerID)
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
    Send_Subtabs(Logix, 8, 2, , BannerID)

    If (Logix.UserRoles.AccessBanners = False) Then
        Send_Denied(1, "perm.access-banners")
        Send_BodyEnd()
        GoTo done
    ElseIf (MyCommon.Fetch_SystemOption(66) <> "1") Then
        Send_Denied(1, "banners.disabled-note")
        GoTo done
    End If
%>

<script type="text/javascript" language="javascript">
  // convert all characters to lowercase to simplify testing
  var agt=navigator.userAgent.toLowerCase();
  var is_ie = ((agt.indexOf("msie") != -1) && (agt.indexOf("opera") == -1));
  
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
  function handleKeyDown(e) {
    var key = e.which ? e.which : e.keyCode;
    if (key == 40) {
      var elemSlct = document.getElementById("functionselect");
      if (elemSlct != null && elemSlct.options.length > 0) {
          elemSlct.options[0].selected = true;
          elemSlct.focus();
          e.returnValue = false;
      }
    } else if (key == 13) {
      handleSelectClick('select');
      if (document.getElementById("functioninput") != null) {
        document.getElementById("functioninput").value = "";
        handleKeyUp(99999);
      }
    }
  }
    function decodeString(encodedStr) {
        var parser = new DOMParser;
        var dom = parser.parseFromString(
            '<!doctype html><body>' + encodedStr,
            'text/html');
        var decodedString = dom.body.textContent;
        return decodedString;
    }
  // This is the function that refreshes the list after a keypress.
  // The maximum number to show can be limited to improve performance with
  // huge lists (1000s of entries).
  // The function clears the list, and then does a linear search through the
  // globally defined array and adds the matches back to the list.
  function handleKeyUp(maxNumToShow) {
      var selectObj, textObj, functionListLength;
      var i,  numShown;
      var searchPattern;
      var selectedList = new Array;
      
      document.getElementById("functionselect").size = "6";
      
      // Set references to the form elements
      selectObj = document.getElementById("functionselect");
      textObj = document.getElementById("functioninput");
      selectedList = getSelectedValues();
      
      // Remember the function list length for loop speedup
      functionListLength = functionlist.length;
      
      // Set the search pattern depending
      if(document.getElementById("functionradio1").checked == true)
      {
          searchPattern = "^"+textObj.value;
      }
      else
      {
          searchPattern = textObj.value;
      }
      searchPattern = cleanRegExpString(searchPattern);
      
      // Create a regulare expression
      
      re = new RegExp(searchPattern,"gi");
      // Clear the options list
      selectObj.length = 0;
      
      // Loop through the array and re-add matching options
      numShown = 0;
      for(i = 0; i < functionListLength; i++)
      {
          if(functionlist[i].search(re) != -1)
          {
              if (vallist[i] != "" && !isSelectedValue(vallist[i],selectedList)) {
                  selectObj[numShown] = new Option(decodeString(functionlist[i]),vallist[i]);
                  selectObj[numShown].title = titlelist[i];
                  selectObj[numShown].style.whiteSpace = 'pre';
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
  
  // this function gets the selected value and loads the appropriate
  // php reference page in the display frame
  // it can be modified to perform whatever action is needed, or nothing
  function handleSelectClick(itemSelected) {
      textObj = document.forms[0].functioninput;
      selectObj = document.forms[0].functionselect;
      selectboxObj = document.forms[0].selected;
      
      changeFlag = document.getElementById("userschanged");
      if (changeFlag != null) {
        changeFlag.value = "true";
      }
      
      if(itemSelected == "select") {
        for (var i=0; i < selectObj.options.length; i++) {
          if (selectObj[i].selected) {
            selectedValue = selectObj[i].value;
            if(selectedValue != "") {
              selectedText = selectObj[i].text;
              if (!is_ie) {
                selectedText = padUserName(selectedText);
              }
              selectboxObj[selectboxObj.length] = new Option(selectedText,selectedValue);
              selectboxObj[selectboxObj.length-1].title = selectedText;
              selectboxObj[selectboxObj.length-1].style.whiteSpace = 'pre';
            }
          }
        }
      } else if (itemSelected == "deselect") {
        for (var i=0; i < selectboxObj.options.length; i++) {
          if (selectboxObj[i].selected) {
            selectboxObj.remove(i)
            i--;
          }
        }
      }      
      
      handleKeyUp(99999);
      
      return true;
  }
  
  function padUserName(optionText) {
    var firstPart = '';
    var secondPart = '';
    var pos = -1;
    
    pos = optionText.indexOf(' ');
    
    if (pos > -1) {
      firstPart = optionText.substring(0, pos);
      while (firstPart.length < 19) {
        firstPart += " "
      }
      secondPart = optionText.substring(pos);
    } else {
      firstPart = optionText;
      secondPart = '';
    }
    
    return firstPart + secondPart;
  }
  
  function getSelectedValues() {
    var elem = document.getElementById('selected');
    var arr = new Array;
    
    if (elem != null) {
      for(var i=0; i < elem.options.length; i++) {
        arr[i] = elem.options[i].value;
      }
    }
    return arr;  
  }
  
  function isSelectedValue(val, arr) {
    var i = 0;
    var retVal = false;
    for (i = 0; (i < arr.length ) && (!retVal); i++) {
      retVal = (val == arr[i]);
    }
    return retVal;
  }
  
  function saveUsers() {
    var changeFlag = document.getElementById("userschanged");
    var elem = document.getElementById("newuserids");
    var elemSel = document.getElementById("selected");
    var userids = '';

    if (changeFlag != null && changeFlag.value == "true") {
      if (elem != null && elemSel != null) {
        for(var i=0; i < elemSel.options.length; i++) {
          if (i>0) { userids += ","; }
          userids += elemSel.options[i].value;
        }
        elem.value = userids;
      }
    }
  } 
  
  function submitForm() {
    saveUsers();
    return true;
  } 
</script>

<form action="#" method="get" id="mainform" name="mainform" onsubmit="return submitForm();">
  <div id="intro">
    <h1 id="title">
      <%
        If BannerID = -1 Then
          Sendb(Copient.PhraseLib.Lookup("term.newbanner", LanguageID))
        Else
          Sendb(Copient.PhraseLib.Lookup("term.banner", LanguageID) & " #" & BannerID & ": ")
          MyCommon.QueryStr = "SELECT Name FROM Banners with (NoLock) WHERE BannerId = " & BannerID & ";"
          rst = MyCommon.LRT_Select
          If (rst.Rows.Count > 0) Then
            BannerNameTitle = MyCommon.NZ(rst.Rows(0).Item("Name"), "")
          End If
          If (Len(BannerNameTitle) <= 31) Then
            Sendb(BannerNameTitle)
          Else
            Sendb(Left(BannerNameTitle, 31) & "...")
          End If
        End If
      %>
    </h1>
    <div id="controls">
      <%
        If (BannerID = -1 And Not IsDeleted) Then
          If (Logix.UserRoles.CreateBanners) Then
            Send_Save()
          End If
        ElseIf (Not IsDeleted) Then
          ShowActionButton = (Logix.UserRoles.CreateBanners) OrElse (Logix.UserRoles.EditBanners) OrElse (Logix.UserRoles.DeleteBanners)
          If (ShowActionButton) Then
            Send("<input type=""button"" class=""regular"" id=""actions"" name=""actions"" value=""" & Copient.PhraseLib.Lookup("term.actions", LanguageID) & " ▼"" onclick=""toggleDropdown();"" />")
            Send("<div class=""actionsmenu"" id=""actionsmenu"">")
            If (Logix.UserRoles.EditBanners) Then
              Send_Save()
            End If
            If (Logix.UserRoles.DeleteBanners AndAlso BannerID > 0) Then
              Send_Delete()
            End If
            If (Logix.UserRoles.CreateBanners) Then
              Send_New()
            End If
            If Request.Browser.Type = "IE6" Then
              Send("<iframe src=""javascript:'';"" id=""actionsiframe"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no"" style=""height:73px;""></iframe>")
            End If
            Send("</div>")
          End If
          If MyCommon.Fetch_SystemOption(75) Then
            If (Logix.UserRoles.AccessNotes) Then
              Send_NotesButton(20, BannerID, AdminUserID)
            End If
          End If
        End If
      %>
    </div>
  </div>
  <%
    If Request.Browser.Type = "IE6" Then
      IE6ScrollFix = " onscroll=""javascript:document.getElementById('actionsmenu').style.visibility='hidden';"""
    End If
  %>
  <div id="main"<% Sendb(IE6ScrollFix) %>>
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <% If (statusMessage <> "") Then Send("<div id=""statusbar"" class=""green-background"">" & statusMessage & "</div>")%>
    <% If (IsDeleted) Then GoTo done%>
    <div id="column1">
      <div class="box" id="identity">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.identification", LanguageID))%>
          </span>
        </h2>
        <input type="hidden" id="BannerID" name="BannerID" value="<% SendB(IIf(BannerID > -1, BannerID, "new")) %>" />
        <label for="name"><% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>:</label><br />
        <input type="text" class="longest" id="name" name="name" maxlength="100" value="<% Sendb(BannerName.Replace("""", "&quot;")) %>" /><br />
        <label for="code"><% Sendb(Copient.PhraseLib.Lookup("term.code", LanguageID))%>:</label><br />
        <input type="text" class="longest" id="xid" name="xid" maxlength="50" value="<% Sendb(ExtBannerID.Replace("""", "&quot;")) %>" /><br />
        <label for="desc"><% Sendb(Copient.PhraseLib.Lookup("term.description", LanguageID))%>:</label><br />
        <textarea class="longest" id="desc" name="desc" cols="48" rows="3"><% Sendb(BannerDescription)%></textarea><br />
        <br class="half" />
        <%
          If (BannerID = -1) Then
            MyCommon.QueryStr = "select EngineID, Description, PhraseID, DefaultEngine from PromoEngines with (NoLock) " & _
                                "where Installed=1 and BannerSupported=1 order by EngineID;"
            dt = MyCommon.LRT_Select
            If (dt.Rows.Count = 1) Then
              Sendb(Copient.PhraseLib.Lookup("term.engine", LanguageID) & ": ")
              If (IsDBNull(dt.Rows(0).Item("PhraseID"))) Then
                Send(MyCommon.NZ(dt.Rows(0).Item("Description"), ""))
              Else
                Sendb(Copient.PhraseLib.Lookup(dt.Rows(0).Item("PhraseID"), LanguageID).ToString.Trim)
              End If
              EngineID = MyCommon.NZ(dt.Rows(0).Item("EngineID"), -1)
              Send("<input type=""hidden"" name=""engine"" id=""engine"" value=""" & EngineID & """ />")
            ElseIf (dt.Rows.Count > 1) Then
              Send("<label for=""engine"">" & Copient.PhraseLib.Lookup("term.engine", LanguageID) & ":</label><br />")
              Send(Space(8) & "<select name=""engine"" id=""engine"" class=""longest"" >")
              For Each row In dt.Rows
                If (EngineID = MyCommon.NZ(row.Item("EngineID"), -1)) Then
                  OptionSelected = True
                ElseIf (EngineID < 0 AndAlso MyCommon.NZ(row.Item("DefaultEngine"), 0) = 1) Then
                  OptionSelected = True
                Else
                  OptionSelected = False
                End If
                Sendb(Space(10) & "<option value=""" & MyCommon.NZ(row.Item("EngineID"), -1) & """" & IIf(OptionSelected, " selected=""selected"" ", "") & ">")
                If (IsDBNull(row.Item("PhraseID"))) Then
                  Send(MyCommon.NZ(row.Item("Description"), "").ToString.Trim)
                Else
                  Sendb(Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID).ToString.Trim)
                End If
                Send(Space(10) & "</option>")
              Next
              Send(Space(8) & "</select>")
            End If
          Else
            MyCommon.QueryStr = "Select PE.EngineID, PE.Description, PE.PhraseID from PromoEngines PE with (NoLock) " & _
                                "inner join BannerEngines BE with (NoLock) on BE.EngineID = PE.EngineID where BE.BannerID = " & BannerID & ";"
            dt = MyCommon.LRT_Select
            If (dt.Rows.Count > 0) Then
              Sendb(Copient.PhraseLib.Lookup("term.engine", LanguageID) & ": ")
              If (IsDBNull(dt.Rows(0).Item("PhraseID"))) Then
                Send(MyCommon.NZ(dt.Rows(0).Item("Description"), ""))
              Else
                Sendb(Copient.PhraseLib.Lookup(dt.Rows(0).Item("PhraseID"), LanguageID).ToString.Trim)
              End If
              EngineID = MyCommon.NZ(dt.Rows(0).Item("EngineID"), -1)
              Send("<input type=""hidden"" name=""engine"" id=""engine"" value=""" & EngineID & """ />")
            End If
          End If

          MyCommon.QueryStr = "select BAN.BannerID from BannerEngines BE with (NoLock) " & _
                      "inner join Banners BAN with (NoLock) on BAN.BannerID = BE.BannerID " & _
                      "where BE.EngineID=" & EngineID & " and BAN.AllBanners=1;"
          dt = MyCommon.LRT_Select
          AllBannersUsed = (dt.Rows.Count > 0)
          If (MyCommon.Fetch_SystemOption(67) = "1" AndAlso BannerID > -1 AndAlso (Not AllBannersUsed Or IsAllBanners)) Then
        %>
        <br />
        <br class="half" />
        <input type="checkbox" name="allbanners" id="allbanners" value="true"<% Sendb(IIf(IsAllBanners, " checked=""checked"" ", "")) %> />
        <label for="allbanners"><%Sendb(Copient.PhraseLib.Lookup("banners.all-banners", LanguageID))%></label>
        <% End If%>
        <%
          MyCommon.QueryStr = "select BAN.BannerID from BannerEngines BE with (NoLock) " & _
                      "inner join Banners BAN with (NoLock) on BAN.BannerID = BE.BannerID " & _
                      "where BE.EngineID=" & EngineID & " and BAN.DefaultBanner=1;"
          dt = MyCommon.LRT_Select
          
          If (dt.Rows.Count = 0) OrElse (dt.Rows.Count > 0 AndAlso BannerID = MyCommon.NZ(dt.Rows(0).Item("BannerID"), -1)) Then
            Send("<br /><br class=""half"" />")
            Send("<input type=""checkbox"" name=""defaultbanner"" id=""defaultbanner"" value=""true""" & IIf(dt.Rows.Count = 0, "", " checked=""checked""") & " /> ")
            Send("<label for=""defaultbanner"">" & Copient.PhraseLib.Lookup("banners.default-banner", LanguageID) & "</label>")
          End If
        %>
        &nbsp;<br />
        <hr class="hidden" />
      </div>
      <%
        'If (BannerID > -1 And (Logix.UserRoles.EditBannerChargebacks OrElse Logix.UserRoles.EditBannerLocHierarchy OrElse Logix.UserRoles.EditBannerProdHierarchy)) Then
        If (BannerID > -1 And (Logix.UserRoles.EditBannerLocHierarchy OrElse Logix.UserRoles.EditBannerProdHierarchy)) Then
      %>
      <!-- <div class="box" id="general"> -->
      <div class="box" id="storehierarchies">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.storehierarchies", LanguageID))%>
          </span>
        </h2>
        <%
          If (Logix.UserRoles.EditBannerLocHierarchy) Then
            ' Available hierarchies
            Sendb("<label for=""lochierarchy-available""><b>" & Copient.PhraseLib.Lookup("term.available", LanguageID) & ":</b></label><br />")
            MyCommon.QueryStr = "select HierarchyID, ExternalID, Name from LocationHierarchies with (NoLock) " & _
                                "where HierarchyID not in (select HierarchyID from BannerLocHierarchies with (NoLock) where BannerID=" & BannerID & ") " & _
                                "order By ExternalID;"
            dt = MyCommon.LRT_Select
            Send("<select class=""longest"" id=""lochierarchy-available"" name=""lochierarchy-available"" multiple=""multiple"" size=""3"">")
            If (dt.Rows.Count > 0) Then
              For Each row In dt.Rows
                Sendb("<option value=""" & MyCommon.NZ(row.Item("HierarchyID"), -1) & """")
                If (LocHierarchyID = MyCommon.NZ(row.Item("HierarchyID"), -1)) Then
                  Sendb(" selected=""selected"" ")
                End If
                Send(">" & MyCommon.NZ(row.Item("ExternalID"), "") & _
                     " - " & MyCommon.NZ(row.Item("Name"), "") & "</option>")
              Next
            End If
            Send("</select><br />")
        %>
        <br class="half" />
        <input type="submit" class="regular select" id="lochierarchy-select" name="lochierarchy-select" value="&#9660; <% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID))%>" />&nbsp;
        <input type="submit" class="regular deselect" id="lochierarchy-deselect" name="lochierarchy-deselect" value="<% Sendb(Copient.PhraseLib.Lookup("term.deselect", LanguageID))%> &#9650;" /><br />
        <br class="half" />
        <%
          ' Selected hierarchies
          Sendb("<label for=""lochierarchy-selected""><b>" & Copient.PhraseLib.Lookup("term.selected", LanguageID) & ":</b></label><br />")
          MyCommon.QueryStr = "select LH.HierarchyID, LH.ExternalID, LH.Name from BannerLocHierarchies BLH with (NoLock) " & _
                              "inner join LocationHierarchies LH with (NoLock) on LH.HierarchyID=BLH.HierarchyID " & _
                              "where BannerID=" & BannerID & ";"
          dt = MyCommon.LRT_Select
          Send("<select class=""longest"" id=""lochierarchy-selected"" name=""lochierarchy-selected"" multiple=""multiple"" size=""3"">")
          If (dt.Rows.Count > 0) Then
            For Each row In dt.Rows
              Sendb("<option value=""" & MyCommon.NZ(row.Item("HierarchyID"), -1) & """")
              If (LocHierarchyID = MyCommon.NZ(row.Item("HierarchyID"), -1)) Then
                Sendb(" selected=""selected"" ")
              End If
              Send(">" & MyCommon.NZ(row.Item("ExternalID"), "") & _
                   " - " & MyCommon.NZ(row.Item("Name"), "") & "</option>")
            Next
          End If
          Send("</select>")
          Send("<br />&nbsp;<br />")
        End If
        %>
      </div>
      <div class="box" id="producthierarchies">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.producthierarchies", LanguageID))%>
          </span>
        </h2>
        <%
          If (Logix.UserRoles.EditProductHierarchy) Then
            ' Available hierarchies
            Sendb("<label for=""prodhierarchy-available""><b>" & Copient.PhraseLib.Lookup("term.available", LanguageID) & ":</b></label><br />")
            MyCommon.QueryStr = "select HierarchyID, ExternalID, Name from ProdHierarchies with (NoLock) " & _
                                "where HierarchyID not in (select HierarchyID from BannerProdHierarchies with (NoLock) where BannerID=" & BannerID & ") " & _
                                "order By ExternalID;"
            dt = MyCommon.LRT_Select
            Send("<select class=""longest"" id=""prodhierarchy-available"" name=""prodhierarchy-available"" multiple=""multiple"" size=""3"">")
            If (dt.Rows.Count > 0) Then
              For Each row In dt.Rows
                Sendb("<option value=""" & MyCommon.NZ(row.Item("HierarchyID"), -1) & """")
                If (LocHierarchyID = MyCommon.NZ(row.Item("HierarchyID"), -1)) Then
                  Sendb(" selected=""selected"" ")
                End If
                Send(">" & MyCommon.NZ(row.Item("ExternalID"), "") & _
                     " - " & MyCommon.NZ(row.Item("Name"), "") & "</option>")
              Next
            End If
            Send("</select><br />")
        %>
        <br class="half" />
        <input type="submit" class="regular select" id="prodhierarchy-select" name="prodhierarchy-select" value="&#9660; <% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID))%>" />&nbsp;
        <input type="submit" class="regular deselect" id="prodhierarchy-deselect" name="prodhierarchy-deselect" value="<% Sendb(Copient.PhraseLib.Lookup("term.deselect", LanguageID))%> &#9650;" /><br />
        <br class="half" />
        <%
          ' Selected hierarchies
          Sendb("<label for=""prodhierarchy-selected""><b>" & Copient.PhraseLib.Lookup("term.selected", LanguageID) & ":</b></label><br />")
          MyCommon.QueryStr = "select LH.HierarchyID, LH.ExternalID, LH.Name from BannerProdHierarchies BLH with (NoLock) " & _
                              "inner join ProdHierarchies LH with (NoLock) on LH.HierarchyID=BLH.HierarchyID " & _
                              "where BannerID=" & BannerID & ";"
          dt = MyCommon.LRT_Select
          Send("<select class=""longest"" id=""prodhierarchy-selected"" name=""prodhierarchy-selected"" multiple=""multiple"" size=""3"">")
          If (dt.Rows.Count > 0) Then
            For Each row In dt.Rows
              Sendb("<option value=""" & MyCommon.NZ(row.Item("HierarchyID"), -1) & """")
              If (LocHierarchyID = MyCommon.NZ(row.Item("HierarchyID"), -1)) Then
                Sendb(" selected=""selected"" ")
              End If
              Send(">" & MyCommon.NZ(row.Item("ExternalID"), "") & _
                   " - " & MyCommon.NZ(row.Item("Name"), "") & "</option>")
            Next
          End If
          Send("</select>")
          Send("<br />&nbsp;<br />")
        End If
        %>
      </div>
      <% End If%>
      <% If (BannerID > -1 AndAlso Logix.UserRoles.EditDepartments AndAlso Not IsAllBanners) Then%>
      <div class="box" id="departments">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.departments", LanguageID))%>
          </span>
        </h2>
        <%
          'MyCommon.QueryStr = "select ChargebackDeptID, ExternalID, Name from ChargebackDepts with (NoLock) where BannerID=" & BannerID & " and Deleted=0 order by ExternalID;"
          MyCommon.QueryStr = "select ChargebackDeptID, ExternalID, CD.Name, CASE WHEN BAN.DefaultChargebackDeptID is NULL THEN 0 ELSE 1 END as DefaultDept " & _
                              "from ChargebackDepts CD with (NoLock) " & _
                              "left join Banners BAN with (NoLock) on BAN.DefaultChargebackDeptID = CD.ChargebackDeptID and BAN.Deleted=0 and CD.Deleted=0 " & _
                              "where CD.BannerID=" & BannerID & " order by ExternalID;"
          dt = MyCommon.LRT_Select
          Send("<select class=""longest"" name=""dept"" id=""dept"" size=""5"" style=""font-family:Monospace;"">")
          For Each row In dt.Rows
            If (MyCommon.NZ(row.Item("ExternalID"), "") <> "") Then
              DisplayStr = MyCommon.NZ(row.Item("ExternalID"), "")
              For i = 0 To 6 - IIf(DisplayStr.Length <= 6, DisplayStr.Length, 6)
                DisplayStr &= "&nbsp;"
              Next
              DisplayStr &= " "
            Else
              DisplayStr = ""
            End If
            DisplayStr &= MyCommon.NZ(row.Item("Name"), "")
            IsDefaultDept = (MyCommon.NZ(row.Item("DefaultDept"), 0) = 1)
            If IsDefaultDept Then
              DisplayStr &= "(" & Copient.PhraseLib.Lookup("term.default", LanguageID) & ")"
            End If
            Send("  <option title=""" & DisplayStr & """ value=""" & MyCommon.NZ(row.Item("ChargebackDeptID"), -1) & """ " & IIf(IsDefaultDept, " style=""color:brown;font-weight:bold;"" ", "") & ">" & DisplayStr & "</option>")
          Next
          Send("</select>")
          Send("<br /><br class=""half"" />")
          Send("<input type=""submit"" class=""mediumshort"" name=""defaultdept"" id=""defaultdept"" value=""" & Copient.PhraseLib.Lookup("term.default", LanguageID) & """ />")
          Send("&nbsp;")
          Send("<input type=""submit"" class=""mediumshort"" name=""removedept"" id=""removedept"" value=""" & Copient.PhraseLib.Lookup("term.remove", LanguageID) & """ />")
          Send("<br class=""half"" />")
          Send("<hr />")
          Send("<br class=""half"" />")
          Sendb("<div style=""width:50px;float:left;""><label for=""deptxid"">" & Copient.PhraseLib.Lookup("term.code", LanguageID) & ":</label></div>")
          Send("<input type=""text"" class=""long"" maxlength=""120"" name=""deptxid"" id=""deptxid"" /><br /><br class=""half"" />")
          Sendb("<div style=""width:50px;float:left;""><label for=""deptname"">" & Copient.PhraseLib.Lookup("term.name", LanguageID) & ":</label></div>")
          Send("<input type=""text"" class=""long"" name=""deptname"" id=""deptname"" /><br /><br class=""half"" />")
          Send("<span style=""margin-left:50px;""><input type=""submit"" class=""mediumshort"" name=""adddept"" id=""adddept"" value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """ /></span>")
        %>
      </div>
      <% End If%>
      <% If (BannerID > -1 AndAlso Logix.UserRoles.EditTerminals AndAlso Not IsAllBanners) Then%>
      <div class="box" id="terminals">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.terminals", LanguageID))%>
          </span>
        </h2>
        <%
          MyCommon.QueryStr = "select TerminalTypeID, Name from TerminalTypes with (NoLock) where BannerID=" & BannerID & " and Deleted=0 order by Name;"
          dt = MyCommon.LRT_Select
          Send("<select class=""longest"" name=""terminal"" id=""terminal"" size=""5"" style=""font-family:Monospace;"">")
          For Each row In dt.Rows
            DisplayStr = MyCommon.NZ(row.Item("Name"), "")
            Send("  <option value=""" & MyCommon.NZ(row.Item("TerminalTypeID"), -1) & """>" & DisplayStr & "</option>")
          Next
          Send("</select>")
          Send("<br /><br class=""half"" />")
          Send("<input type=""submit"" class=""mediumshort"" name=""removeterm"" id=""removeterm"" value=""" & Copient.PhraseLib.Lookup("term.remove", LanguageID) & """/>")
          Send("<br class=""half"" />")
          Send("<hr />")
          Send("<br class=""half"" />")
          Sendb("<div style=""width:50px;float:left;""><label for=""termxid"">" & Copient.PhraseLib.Lookup("term.code", LanguageID) & ":</label></div>")
          Send("<input type=""text"" class=""long"" name=""termxid"" id=""termxid"" /><br /><br class=""half"" />")
          Sendb("<div style=""width:50px;float:left;""><label for=""termname"">" & Copient.PhraseLib.Lookup("term.name", LanguageID) & ":</label></div>")
          Send("<input type=""text"" class=""long"" name=""termname"" id=""termname"" /><br /><br class=""half"" />")
          Send("<span style=""margin-left:50px;""><input type=""submit"" class=""mediumshort"" name=""addterm"" id=""addterm"" value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """/></span>")
        %>
      </div>
      <% End If%>
    </div>
    
    <div id="gutter">
    </div>
    
    <div id="column2">
      <div class="box" id="users"<%Sendb(IIf(BannerID > -1, "", " style=""visibility:hidden;"" "))%>>
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.users", LanguageID))%>
          </span>
        </h2>
        <input type="radio" id="functionradio1" name="functionradio" checked="checked" /><label for="functionradio1"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
        <input type="radio" id="functionradio2" name="functionradio" /><label for="functionradio2"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
        <input type="text" class="medium" id="functioninput" name="functioninput" onkeydown="handleKeyDown(event);" onkeyup="handleKeyUp(200);" value="" /><br />
        <br class="half" />
        <label for="functionselect"><b><%Sendb(Copient.PhraseLib.Lookup("term.available", LanguageID))%>:</b></label><br />
        <select class="longest" id="functionselect" name="functionselect" multiple="multiple" size="6" style="font-family: Monospace;">
        </select>
        <br />
        <br class="half" />
        <input type="submit" class="regular select" id="select" name="select" value="&#9660; <% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID))%>" onclick="handleSelectClick('select');"<% Sendb(IIf (Logix.UserRoles.AddUsersToBanners, "", " disabled=""disabled"" ")) %> />&nbsp;
        <input type="submit" class="regular deselect" id="deselect" name="deselect" value="<% Sendb(Copient.PhraseLib.Lookup("term.deselect", LanguageID))%> &#9650;" onclick="handleSelectClick('deselect');"<% Sendb(IIf (Logix.UserRoles.RemoveUsersFromBanners, "", " disabled=""disabled"" ")) %> /><br />
        <br class="half" />
        <label for="selected"><b><%Sendb(Copient.PhraseLib.Lookup("term.selected", LanguageID))%>:</b></label><br />
        <select class="longest" id="selected" name="selected" multiple="multiple" size="6" style="font-family: Monospace;">
          <%
            ' alright lets find the currently selected users on page load
            MyCommon.QueryStr = "select AU.AdminUserID, AU.UserName, AU.FirstName, AU.LastName from AdminUsers AU with (NoLock) " & _
                                "inner join AdminUserBanners AUB with (NoLock) on AUB.AdminUserID = AU.AdminUserID " & _
                                "where AUB.BannerID = " & BannerID & " order by AU.UserName;"
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count > 0) Then
              ReDim BannerUserIDs(rst.Rows.Count - 1)
              i = 0
              For Each row In rst.Rows
                UserTextFull = MyCommon.NZ(row.Item("UserName"), "") & "  (" & MyCommon.NZ(row.Item("FirstName"), "") & " " & MyCommon.NZ(row.Item("LastName"), "") & ")"
                UserTextDisplay = MyCommon.NZ(row.Item("UserName"), "").ToString().PadRight(19, " ").Replace(" ", "&nbsp;") & "  " & MyCommon.NZ(row.Item("FirstName"), "") & " " & MyCommon.NZ(row.Item("LastName"), "")
                Send(Space(10) & "<option title=""" & UserTextFull & """ value=""" & MyCommon.NZ(row.Item("AdminUserId"), -1) & """>" & UserTextDisplay & "</option>")
                BannerUserIDs(i) = MyCommon.NZ(row.Item("AdminUserId"), -1)
                i += 1
              Next
            End If
          %>
        </select>
        <input type="hidden" name="existinguserids" id="existinguserids" value="<%If (BannerUserIDs IsNot Nothing) Then Sendb(String.Join(",", BannerUserIDs))%>" />
        <input type="hidden" name="newuserids" id="newuserids" value="" />
        <input type="hidden" name="userschanged" id="userschanged" value="false" />
        <br />
        &nbsp;<br class="half" />
      </div>
      <div class="box" id="offers"<%Sendb(IIf(BannerID > -1, "", " style=""display:none;"" "))%>>
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.recently", LanguageID) & " " & Copient.PhraseLib.Lookup("term.associatedoffers", LanguageID).ToLower())%>
          </span>
        </h2>
        <%
          If Not dtOffers Is Nothing AndAlso dtOffers.Rows.Count > 0 Then
            Send("    " & Copient.PhraseLib.Lookup("banner-edit.offers", LanguageID))
            Send("    <div class=""boxscroll"">")
            For Each row In dtOffers.Rows
              If (Logix.IsAccessibleOffer(AdminUserID, MyCommon.NZ(row.Item("OfferID"), 0))) Then
                Sendb(" <a href=""offer-redirect.aspx?OfferID=" & MyCommon.NZ(row.Item("OfferID"), 0) & """>" & MyCommon.NZ(row.Item("OfferName"), "") & "</a>")
              Else
                Sendb(MyCommon.NZ(row.Item("OfferName"), ""))
              End If
              Send("<br />")
            Next
            Send("    </div>")
          Else
            Send("    <div class=""boxscroll"">")
            Sendb(Copient.PhraseLib.Lookup("banner-edit.nooffers", LanguageID))
            Send("    </div>")
          End If
        %>
        <hr class="hidden" />
      </div>
      <div class="box" id="groups"<%Sendb(IIf(BannerID > -1, "", " style=""display:none;"" "))%>>
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.associatedgroups", LanguageID))%>
          </span>
        </h2>
        <%
          If Not dtGroups Is Nothing AndAlso dtGroups.Rows.Count > 0 Then
            Send("    " & Copient.PhraseLib.Lookup("banner-edit.groups", LanguageID))
            Send("    <div class=""boxscroll"">")
            For Each row In dtGroups.Rows
              Send("     <a href=""lgroup-edit.aspx?LocationGroupId=" & MyCommon.NZ(row.Item("LocationGroupId"), 0) & """>" & MyCommon.NZ(row.Item("Name"), "") & "</a><br />")
            Next
            Send("    </div>")
          Else
            Send("    <div class=""boxscroll"">")
            Sendb(Copient.PhraseLib.Lookup("banner-edit.nogroups", LanguageID) & "<br />")
            Send("    </div>")
          End If
        %>
        <hr class="hidden" />
      </div>
      <div class="box" id="stores"<%Sendb(IIf(BannerID > -1, "", " style=""display:none;"" "))%>>
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.associatedstores", LanguageID))%>
          </span>
        </h2>
        <%
          If Not dtStores Is Nothing AndAlso dtStores.Rows.Count > 0 Then
            Send("    " & Copient.PhraseLib.Lookup("banner-edit.stores", LanguageID))
            Send("    <div class=""boxscroll"">")
            For Each row In dtStores.Rows
              Send("     <a href=""store-edit.aspx?LocationID=" & MyCommon.NZ(row.Item("LocationID"), 0) & """>" & MyCommon.NZ(row.Item("LocationName"), "") & "</a><br />")
            Next
            Send("    </div>")
          Else
            Send("    <div class=""boxscroll"">")
            Sendb(Copient.PhraseLib.Lookup("banner-edit.nostores", LanguageID) & "<br />")
            Send("    </div>")
          End If
        %>
        <hr class="hidden" />
      </div>
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
    
    <%
    MyCommon.QueryStr = "select AdminUserID, UserName, FirstName, LastName from AdminUsers with (NoLock) " & _
                        "order by UserName;"
    rst = MyCommon.LRT_Select
    If (rst.rows.count > 0)
        Sendb("var functionlist = Array(")
        For Each row In rst.Rows
            If (Request.Browser.Browser.ToUpper = "FIREFOX") Then
                UserTextDisplay = MyCommon.NZ(row.Item("UserName"), "").ToString().PadRight(19, " ").Replace(" ", "&nbsp;") & "  " & MyCommon.NZ(row.Item("FirstName"), "") & " " & MyCommon.NZ(row.Item("LastName"), "")
                'UserTextDisplay = MyCommon.NZ(row.Item("UserName"), "").ToString().PadRight(18, " ") & "  " & MyCommon.NZ(row.Item("FirstName"), "") & " " & MyCommon.NZ(row.Item("LastName"), "")
            Else
                UserTextDisplay = MyCommon.NZ(row.Item("UserName"), "").ToString().PadRight(19, " ").Replace(" ", "&nbsp;") & "  " & MyCommon.NZ(row.Item("FirstName"), "") & " " & MyCommon.NZ(row.Item("LastName"), "")
            End If
            UserTextDisplay = UserTextDisplay.Replace("""", "\""")
            Sendb("""" & UserTextDisplay & """,")
        Next
        Send(""""");")
        Sendb("var titlelist = Array(")
        For Each row In rst.Rows
            UserTextDisplay = MyCommon.NZ(row.Item("UserName"), "").ToString().PadRight(19, " ").Replace(" ", " ") & "  " & MyCommon.NZ(row.Item("FirstName"), "") & " " & MyCommon.NZ(row.Item("LastName"), "")
            UserTextDisplay = UserTextDisplay.Replace("""", "\""")
            Sendb("""" & UserTextDisplay & """,")
        Next
        Send(""""");")
        Sendb("var vallist = Array(")
        For Each row In rst.Rows
          Sendb("""" & MyCommon.NZ(row.item("AdminUserID"), 0) & """,")
        Next
        Sendb(""""");")
      Else
        Sendb("var functionlist = Array(")
        Send("""" & " "  & """);")
        Sendb("var vallist = Array(")
        Send("""" & " " & """);")
      End If
    %>
    handleKeyUp(99999);    
</script>

<script runat="server">
  Function GetAddedUsers(ByVal ExistingUserIDs As String(), ByVal NewUserIDs As String()) As ArrayList
    Dim AddedUsers As New ArrayList(10)
    Dim i As Integer
    
    For i = 0 To NewUserIDs.GetUpperBound(0)
      If (Array.IndexOf(ExistingUserIDs, NewUserIDs(i)) = -1) Then
        AddedUsers.Add(NewUserIDs(i))
      End If
    Next
    
    Return AddedUsers
  End Function
  
  Function GetRemovedUsers(ByVal ExistingUserIDs As String(), ByVal NewUserIDs As String()) As ArrayList
    Dim RemovedUsers As New ArrayList(10)
    Dim i As Integer
    
    For i = 0 To ExistingUserIDs.GetUpperBound(0)
      If (Array.IndexOf(NewUserIDs, ExistingUserIDs(i)) = -1) Then
        RemovedUsers.Add(ExistingUserIDs(i))
      End If
    Next
    
    Return RemovedUsers
  End Function
  
  Function GetBannerEngineID(ByVal BannerID As Integer, ByRef MyCommon As Copient.CommonInc)
    Dim EngineID As Integer = -1
    Dim dt As DataTable
    
    MyCommon.QueryStr = "select EngineID from BannerEngines with (NoLock) where BannerID=" & BannerID
    dt = MyCommon.LRT_Select
    If (dt.Rows.Count > 0) Then
      EngineID = MyCommon.NZ(dt.Rows(0).Item("EngineID"), -1)
    End If
    
    Return EngineID
  End Function
  
  Function AllBannersExists(ByVal EngineID As Integer, ByVal BannerID As Integer, ByRef MyCommon As Copient.CommonInc)
    Dim AllBannersUsed As Boolean = False
    Dim dt As DataTable
    
    MyCommon.QueryStr = "select BAN.BannerID from BannerEngines BE with (NoLock) " & _
                        "inner join Banners BAN with (NoLock) on BAN.BannerID = BE.BannerID " & _
                        "where BE.EngineID=" & EngineID & " and BAN.AllBanners=1 and BAN.BannerID <> " & BannerID & ";"
    dt = MyCommon.LRT_Select
    AllBannersUsed = (dt.Rows.Count > 0)
    
    Return AllBannersUsed
  End Function
</script>

<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (BannerID > 0 And Logix.UserRoles.AccessNotes) Then
      Send_Notes(20, BannerID, AdminUserID)
    End If
  End If
  Send_BodyEnd("mainform", "name")
done:
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  Logix = Nothing
  MyCommon = Nothing
%>
