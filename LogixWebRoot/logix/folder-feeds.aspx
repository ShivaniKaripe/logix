<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" EnableSessionState="True" %>

<%@ Import Namespace="CMS.AMS.Security" %>

<%@ Import Namespace="CMS.AMS.Models" %>
<%@ Import Namespace="System.Threading" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%@ Import Namespace="CMS.AMS.Models" %>
<%@ Import Namespace="CMS.AMS.Contract" %>
<%@ Import Namespace="Microsoft.Practices.Unity" %>
<%@ Import Namespace="CMS.AMS" %>
<%'VB185060: For making Mass operation aynchronous, implemented ThreadPool.QueueUserWorkItem. We will set folder Massoperation status to
'FOLDER_IN_USE before the start of operation and in the end we set it to FOLDER_NOT_IN_USE+ SUCCESS(optional)
'(SUCCESS-If operation is successfully performed, which makes operation status to be shown in Green otherwise operation status is shown in Red).
'If we need Resolver created by dependency injection in any case, stop disposing the resolver by setting Candispose property to false.
%>
<script runat="server">
    Dim MyCommon As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    Dim MyExport As New Copient.ExportXml(MyCommon)
    Dim MyCPEOffer As New Copient.CPEOffer
    Dim MyCMOffer As New Copient.CMOffer
    Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249) = "1", True, False)
    Public Const FOLDER_IN_USE As String = "~FIU~"
    Public Const FOLDER_NOT_IN_USE As String = "~FNIU~"
    Public Const SUCCESS As String = "Success!"
    Public Const UPDATE_START_DATE As Integer = 0
    Public Const UPDATE_END_DATE As Integer = 1
    Public Const UPDATE_START_AND_END_DATE As Integer = 2
    Dim bOverrideMassUpdateRestriction As Boolean = IIf(MyCommon.Fetch_SystemOption(226) = "1", True, False)
    Dim bStoreUser As Boolean = False
    Dim sValidLocIDs As String = ""
    Dim sValidSU As String = ""
    Dim wherestr As String = ""
    Dim buyerwherestr As String = ""
    Dim sJoin As String = ""
    Dim buyerJoin As String = ""
</script>
<%
    Dim CopientFileName As String = "FolderFeeds.aspx"
    Dim CopientFileVersion As String = "7.3.1.138972"
    Dim CopientProject As String = "Copient Logix"
    Dim CopientNotes As String = ""

    Dim AdminUserID As Integer
    Dim ItemIDs As String = ""
    Dim LinkIDs As String = ""
    Dim LinkTypeIDs As String = ""
    Dim FolderID As Integer
    Dim FolderIDs As String
    Dim FromOfferList As Boolean
    Dim OffersWithoutConditions As Boolean
    Dim WFStatus As Integer
    Dim ActionItem As Integer
    Dim NoOfDuplicateOffers As Integer = 0
    Dim strStatusDupOffers As String = ""
    Dim flush As Boolean = True
    Dim rst As DataTable
    Dim iLen As Integer = 0
    Dim i As Integer = 0

    MyCommon.AppName = "FolderFeeds.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)

    'Store User
    If (MyCommon.Fetch_CM_SystemOption(131) = "1") Then
        MyCommon.QueryStr = "select LocationID from StoreUsers where UserID=" & AdminUserID & ";"
        rst = MyCommon.LRT_Select
        iLen = rst.Rows.Count
        If iLen > 0 Then
            bStoreUser = True
            sValidSU = AdminUserID
            For i = 0 To (iLen - 1)
                If i = 0 Then
                    sValidLocIDs = rst.Rows(0).Item("LocationID")
                Else
                    sValidLocIDs &= "," & rst.Rows(i).Item("LocationID")
                End If
            Next

            MyCommon.QueryStr = "select UserID from StoreUsers where LocationID in (" & sValidLocIDs & ") and NOT UserID=" & AdminUserID & ";"
            rst = MyCommon.LRT_Select
            iLen = rst.Rows.Count
            If iLen > 0 Then
                For i = 0 To (iLen - 1)
                    sValidSU &= "," & rst.Rows(i).Item("UserID")
                Next
            End If
        End If
    End If
    Response.Expires = 0
    Response.Clear()
    Response.ContentType = "text/html"

    FolderID = MyCommon.Extract_Val(Request.Form("FolderID"))
    Select Case Request.QueryString("Action")
        Case "CopyExpiredOffer"
            CopyExpiredOffer(MyCommon.Extract_Val(Request.Form("OfferID")), Request.Form("FolderList"), AdminUserID)
        Case "LoadFolderItems"
            Dim DeployStatus As String = GetFolderMassOperationStatus(FolderID)
            If (Not DeployStatus.StartsWith(FOLDER_NOT_IN_USE)) Then
                Send(DeployStatus)
                Return
            End If
            WFStatus = Request.QueryString("WFStatus")
            If WFStatus <> 0 Then
                SendFolderItems(FolderID, String.Empty, WFStatus)
            Else
                SendFolderItems(FolderID)
            End If
        Case "LoadFilteredFolderItems"
            WFStatus = Request.Form("WFStatus")
            If WFStatus <> 0 Then
                SendFolderItems(FolderID, String.Empty, WFStatus)
            Else
                SendFolderItems(FolderID)
            End If
        Case "AddItemsToFolder"
            LinkIDs = Request.Form("LinkIDs")
            LinkTypeIDs = Request.Form("LinkTypeIDs")
            WFStatus = Request.Form("WFStatus")
            If WFStatus <> 0 Then
                AddItemsToFolder(FolderID, LinkIDs, LinkTypeIDs, AdminUserID, WFStatus)
            Else
                AddItemsToFolder(FolderID, LinkIDs, LinkTypeIDs, AdminUserID)
            End If
        Case "RemoveItemsFromFolder"
            ItemIDs = Request.Form("ItemIDs")
            RemoveItemsFromFolder(FolderID, ItemIDs, AdminUserID)
        Case "SendFoundOffers"
            SendFoundOffers(FolderID)
        Case "CreateFolder"
            Dim IsdefaultUEFolder = Request.Form("IsdefaultUEFolder") IsNot Nothing AndAlso Request.Form("IsdefaultUEFolder") = True
            CreateFolder(FolderID, Request.Form("FolderName"), Request.Form("AccessLevel"), Request.Form("FolderStartDate"), Request.Form("FolderEndDate"), Request.Form("FolderTheme"), AdminUserID, IsdefaultUEFolder)
        Case "RenameFolder"
            RenameFolder(FolderID, Request.Form("FolderName"), AdminUserID)
        Case "DeleteFolder"
            DeleteFolder(FolderID, AdminUserID)
        Case "SaveOfferFolders"
            SaveOfferFolders(MyCommon.Extract_Val(Request.Form("OfferID")), Request.Form("FolderList"), AdminUserID)
        Case "SendFolderSearch"
            SendFolderSearch(Request.Form("searchterms"))
        Case "SendOfferSearch"
            SendOfferSearch(Request.Form("searchterms"))
        Case "GenDivModFolder"
            SendDivForModFolder(FolderID, AdminUserID)
        Case "ModifyFolder"
            Dim IsdefaultUEFolder = Request.Form("IsdefaultUEFolder") IsNot Nothing AndAlso Request.Form("IsdefaultUEFolder") = True

            Dim FolderName As String = Request.Form("ModFolderName"),
                FolderStartDate As String = Request.Form("ModFolderStartDate"),
                FolderEndDate As String = Request.Form("ModFolderEndDate"),
                FolderTheme As String = Request.Form("ModFolderTheme"),
                IsMassUpdateEnabled As Boolean = Request.Form("isMassupdateEnabled")
            Dim isValid As Boolean = ValidateFolderInfo(FolderID, FolderName, FolderStartDate, FolderEndDate, FolderTheme, IsdefaultUEFolder)
            flush = False
            If (isValid) Then
                Threading.ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf ModifyFolderAsync), New Object() {FolderID, FolderName, FolderStartDate, FolderEndDate, FolderTheme, IsMassUpdateEnabled, IsdefaultUEFolder})
            End If
        Case "ShowFolderInfo"
            SendFolderInfo(FolderID)
        Case "SendOfferFoldersWithFutureDate"
            SendOfferFoldersWithFutureDate(Request.Form("FolderList"), AdminUserID)
        Case "DuplicateOffers"
            Dim SrcFolderId As Integer = Request.Form("Fid")
            FolderIDs = Request.Form("FolderIDs")
            ItemIDs = Request.Form("ItemIDs")
            If Not String.IsNullOrEmpty(Request.Form("FromOfferList")) Then
                FromOfferList = Request.Form("FromOfferList")
            End If
            ActionItem = Request.Form("ActionItem")
            NoOfDuplicateOffers = Request.Form("DuplicateCnt")
            flush = False
            Threading.ThreadPool.QueueUserWorkItem(AddressOf DublicateOffersToFolderAsync, New Object() {FolderIDs, ItemIDs, FromOfferList, ActionItem, strStatusDupOffers, NoOfDuplicateOffers, SrcFolderId})
        Case "DeployOffers"
            ItemIDs = Request.Form("ItemIDs")
            FromOfferList = Request.Form("FromOfferList")
            OffersWithoutConditions = Request.Form("OffersWithoutConditions")
            Dim deploytransreqskip = IIf(GetCgiValue("deploytransreqskip") = "", 0, GetCgiValue("deploytransreqskip"))
            Dim deferdeploy = GetCgiValue("deferdeploy")
            flush = False
            Threading.ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf MassDeployWorker), New Object() {FolderID, ItemIDs, FromOfferList, OffersWithoutConditions, deploytransreqskip, deferdeploy})
        Case "DeferDeployOffers"
            ItemIDs = Request.Form("ItemIDs")
            FromOfferList = Request.Form("FromOfferList")
            OffersWithoutConditions = Request.Form("OffersWithoutConditions")
            Dim deploytransreqskip = IIf(GetCgiValue("deploytransreqskip") = "", 0, GetCgiValue("deploytransreqskip"))
            Dim deferdeploy = GetCgiValue("deferdeploy")
            flush = False
            Threading.ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf MassDeferDeployWorker), New Object() {FolderID, ItemIDs, FromOfferList, OffersWithoutConditions, deploytransreqskip, deferdeploy})
        Case "ApplyFolderStartDatesToOffer"
            ItemIDs = Request.Form("ItemIDs")
            flush = False
            Threading.ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf ApplyFolderStartEndDatesToOfferAsync), New Object() {ItemIDs, AdminUserID, Request.Form("FID"), True, False})
        Case "ApplyFolderEndDatesToOffer"
            ItemIDs = Request.Form("ItemIDs")
            flush = False
            Threading.ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf ApplyFolderStartEndDatesToOfferAsync), New Object() {ItemIDs, AdminUserID, Request.Form("FID"), False, True})
        Case "ApplyFolderStartEndDatesToOffer"
            ItemIDs = Request.Form("ItemIDs")
            flush = False
            Threading.ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf ApplyFolderStartEndDatesToOfferAsync), New Object() {ItemIDs, AdminUserID, Request.Form("FID"), True, True})
        Case "NavigatetoReports"
            ItemIDs = Request.Form("ItemIDs")
            FromOfferList = Request.Form("FromOfferList")

            Dim folder As Integer = IIf(Request.Form("FolderID") <> Nothing, Request.Form("FolderID"), 0)
            NavigatetoReports(ItemIDs, FromOfferList)
        Case "DeleteSelectedOffers"
            ItemIDs = Request.Form("ItemIDs")

            Threading.ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf DeleteSelectedOffersAsync), New Object() {ItemIDs, AdminUserID, Request.Form("FID")})
        Case "SendOutbound"
            ItemIDs = Request.Form("ItemIDs")
            FromOfferList = Request.Form("FromOfferList")
            Dim folder As Integer = IIf(Request.Form("FolderID") <> Nothing, Request.Form("FolderID"), 0)
            flush = False
            Threading.ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf MassSendOutBoundAsync), New Object() {ItemIDs, FromOfferList, folder})

        Case "WFStatustoPreValidate"
            ItemIDs = Request.Form("ItemIDs")
            FromOfferList = Request.Form("FromOfferList")
            WFStatustoPreValidate(ItemIDs, FromOfferList)
        Case "WFStatustoPostValidate"
            ItemIDs = Request.Form("ItemIDs")
            FromOfferList = Request.Form("FromOfferList")
            WFStatustoPostValidate(ItemIDs, FromOfferList)
        Case "WFStatustoReadytoDeploy"
            ItemIDs = Request.Form("ItemIDs")
            FromOfferList = Request.Form("FromOfferList")
            WFStatustoReadytoDeploy(ItemIDs, FromOfferList)
        Case "TransferOffers"
            Dim SourceFolderID As String = Request.Form("sFolder")
            Dim DestFolderID As String = Request.Form("dFolder")
            Dim v = Request.Form("ItemIDs")
            FromOfferList = Request.Form("FromOfferList")
            flush = False
            Threading.ThreadPool.QueueUserWorkItem(AddressOf TransferOffersAsync, New Object() {SourceFolderID, DestFolderID, v, FromOfferList})
        Case ("SaveBuyerFolders")
            SaveBuyerFolders(MyCommon.Extract_Val(Request.Form("id")), Request.Form("FolderList"), AdminUserID)
        Case "CheckDefaultFolder"
            CheckDefaultFolder(MyCommon)
        Case "GetOffersIteratively"
            Dim PageIndex As Integer = Request.Form("pageindex")
            Dim Checked As Boolean = Request.Form("checked")
            WFStatus = Request.Form("WFStatus")
            Send(GetofferTable(FolderID, PageIndex, Checked, WFStatus))
        Case "GetFolderItemIDs"
            WFStatus = Request.Form("WFStatus")
            GetFolderItemsList(FolderID, WFStatus)
        Case "UpdateFolderStatus"
            FromOfferList = Request.Form("FromOfferList")
            If (FromOfferList) Then
                Application("MassOperaionStaus") = ""
            End If
            SetFolderMassOperationStatus(FolderID, FOLDER_NOT_IN_USE)
        Case "GetSubFolders"
            GetSubFolders(Convert.ToInt32(Request.Form("ParentFolderID")))
        Case "GetParentFolders"
            GetParentFolders(Convert.ToInt32(Request.Form("FolderID")))
        Case "ShowRequestApprovalActionItem"

            If Logix.UserRoles.DeployNonTemplateOffers OrElse Logix.UserRoles.DeployTemplateOffers Then
                ShowRequestApprovalActionItem(Request.Form("ItemIds"), Boolean.Parse(Request.Form("FromOfferList")))
            End If
        Case "RequestApproval"
            flush = False
            Threading.ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf MassRequestApprovalWorker), New Object() {Request.Form("ApprovalType"), Request.Form("ItemIDs"), Request.Form("FromOfferList"), Request.Form("FolderID")})
        Case "HandleDeployAndApprovalRequest"
            OffersWithoutConditions = Request.Form("OffersWithoutConditions")
            Dim deploytransreqskip = IIf(GetCgiValue("deploytransreqskip") = "", 0, GetCgiValue("deploytransreqskip"))
            Dim deferdeploy = GetCgiValue("deferdeploy")
            HandleDeployAndApprovalRequest(Request.Form("ItemIds"), Boolean.Parse(Request.Form("FromOfferList")), Integer.Parse(Request.Form("DeploymentType")), FolderID, OffersWithoutConditions, deploytransreqskip, deferdeploy)
        Case Else
            Send("<b>" & Copient.PhraseLib.Lookup("feeds.noarguments", LanguageID) & "!</b>")
            Send(Request.RawUrl)
    End Select
    If (flush) Then
        Response.Flush()
        Response.End()
    End If
%>
<script runat="server">
    Public DefaultLanguageID As Integer
    Public LogFile As String
    Public CollisionsDetected As Boolean = False
    Enum SearchState As Integer
        EXPANDED = 1
        COLLAPSED = 2
    End Enum

    Public Function GetofferTable(ByVal FolderId As Integer, ByVal pageIndex As Integer, ByVal checked As Boolean, Optional ByVal WFStatus As Integer = 0) As String
        Dim ds As DataSet = GetOffersDataset(FolderId, pageIndex, WFStatus)
        'Form HTML table from the Data Table 
        Return ds.Tables(1)(0)(0) + "AMS_SPLITTER_AMS" + GetHTMLOfferTable(ds.Tables(0), checked)
    End Function

    Private Sub SetFolderMassOperationStatus(ByVal FolderID As Integer, ByVal Status As String)
        If FolderID > 0 Then
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            MyCommon.QueryStr = " Update Folders set MassOperationStatus = @Status where FolderID=@FolderID"
            MyCommon.DBParameters.Add("@Status", SqlDbType.NVarChar).Value = Status
            MyCommon.DBParameters.Add("@FolderID", SqlDbType.Int).Value = FolderID
            MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
            MyCommon.Close_LRTsp()
        Else
            Application("MassOperaionStaus") = Status
        End If
    End Sub

    Private Function GetFolderMassOperationStatus(ByVal FolderID As Integer) As String
        Dim tempdt As DataTable
        Dim DeployStatus As String = ""
        If FolderID > 0 Then
            Try
                If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
                MyCommon.QueryStr = " Select MassOperationStatus from Folders where FolderID=" & FolderID
                tempdt = MyCommon.LRT_Select
                If tempdt.Rows.Count > 0 Then
                    DeployStatus = tempdt.Rows(0)(0)
                End If
            Catch e As Exception
            Finally
                MyCommon.Close_LRTsp()
            End Try
        End If
        Return DeployStatus
    End Function

    Public Sub MassDeployWorker(ByVal State As Object)
        Dim obj As Object() = State
        Dim FolderID As Integer = obj(0)
        Dim ItemIDs = obj(1)
        Dim FromOfferList = obj(2)
        Dim OffersWithoutConditions = obj(3)
        Dim deploytransreqskip = obj(4)
        Dim deferdeploy = obj(5)
        Dim DeployStatus As String = String.Empty
        Dim err As String = String.Empty
        Try
            If (FromOfferList) Then
                SetFolderMassOperationStatus(-1, FOLDER_IN_USE + Copient.PhraseLib.Lookup("folders.massdeploy", LanguageID) & AdminName)
            Else
                SetFolderMassOperationStatus(FolderID, FOLDER_IN_USE + Copient.PhraseLib.Lookup("folders.massdeploy", LanguageID) & AdminName)
            End If
            MassDeployOffers(ItemIDs, FromOfferList, OffersWithoutConditions, AdminUserID, deploytransreqskip, deferdeploy, FolderID, DeployStatus)
        Catch ex As Exception
            err = "Error: " + ex.ToString
            MyCommon.Write_Log(LogFile, err)
        Finally
            If (FromOfferList) Then
                SetFolderMassOperationStatus(-1, FOLDER_NOT_IN_USE + DeployStatus)
            ElseIf (CollisionsDetected = True) Then
                SetFolderMassOperationStatus(FolderID, FOLDER_NOT_IN_USE + Copient.PhraseLib.Lookup("term.folderoffers-failed-collision", LanguageID))
            Else
                SetFolderMassOperationStatus(FolderID, FOLDER_NOT_IN_USE + DeployStatus)
            End If
        End Try
        'Dispose the resolver
        'CurrentRequest.Resolver.Dispose()
    End Sub

    Public Sub MassDeferDeployWorker(ByVal State As Object)
        Dim obj As Object() = State
        Dim FolderID As Integer = obj(0)
        Dim ItemIDs = obj(1)
        Dim FromOfferList = obj(2)
        Dim OffersWithoutConditions = obj(3)
        Dim deploytransreqskip = obj(4)
        Dim deferdeploy = obj(5)
        Dim DeployStatus As String = String.Empty
        Dim err As String = String.Empty
        Try
            If (FromOfferList) Then
                SetFolderMassOperationStatus(-1, FOLDER_IN_USE + " Mass defer-deploy is going on for the selected Offers" & AdminName)
            Else
                SetFolderMassOperationStatus(FolderID, FOLDER_IN_USE + "Mass defer-deploy is going on for the Offers in this Folder by " & AdminName)
            End If
            MassDeferDeployOffers(ItemIDs, FromOfferList, OffersWithoutConditions, AdminUserID, deploytransreqskip, deferdeploy, FolderID, DeployStatus)
        Catch ex As Exception
            err = "Error: " + ex.ToString
            MyCommon.Write_Log(LogFile, err)
        Finally
            If (FromOfferList) Then
                SetFolderMassOperationStatus(-1, FOLDER_NOT_IN_USE + DeployStatus)
            ElseIf (CollisionsDetected = True) Then
                SetFolderMassOperationStatus(FolderID, FOLDER_NOT_IN_USE + Copient.PhraseLib.Lookup("term.folderoffers-failed-collision", LanguageID))
            Else
                SetFolderMassOperationStatus(FolderID, FOLDER_NOT_IN_USE + DeployStatus)
            End If
        End Try
        'Dispose the resolver
        'CurrentRequest.Resolver.Dispose()
    End Sub

    ' ********************************************************************************
    ' *** Returns an HTML table of the current folder item for the specified FolderID
    ' *** Should only be called for use in either View or Remove mode.
    ' ********************************************************************************
    Sub SendFolderItems(ByVal FolderID As Long, Optional ByVal StatusText As String = "", Optional ByVal WFStatus As Integer = 0, Optional ByVal pageIndex As Integer = 1)

        Dim dt As DataTable
        Dim TempBuf As New StringBuilder()
        Dim Name As String = ""
        Dim SortDirection = "ASC"
        Dim SortText = "AOLV.OfferID"
        Dim Theme As String = ""
        Dim FolderStartDate As String = ""
        Dim FolderEndDate As String = ""
        Dim itemIdsAll As String = ""
        Dim IsEngineInstalled As Boolean = MyCommon.GetInstalledEngines().Length > 1
        'Set direction and orderby text, if any
        If (Request.QueryString("SortText") = "AOLV.OfferID") Then
            SortText = Request.QueryString("SortText")
        ElseIf (Request.QueryString("SortText") = "AOLV.ProdStartDate") Then
            SortText = Request.QueryString("SortText")
        ElseIf (Request.QueryString("SortText") = "AOLV.ProdEndDate") Then
            SortText = Request.QueryString("SortText")
        ElseIf (Request.QueryString("SortText") = "AOLV.ExtOfferID") Then
            SortText = Request.QueryString("SortText")
        ElseIf (Request.QueryString("SortText") = "AOLV.PromoEngine") Then
            SortText = Request.QueryString("SortText")
        ElseIf (Request.QueryString("SortText") = "AOLV.Name") Then
            SortText = Request.QueryString("SortText")
        ElseIf (Request.QueryString("SortText") = "AOLV.Status") Then
            SortText = Request.QueryString("SortText")
        Else
            'text already set, don't do anything
        End If

        If (Request.QueryString("SortDirection") = "ASC") Then
            SortDirection = "DESC"
        ElseIf (Request.QueryString("SortDirection") = "DESC") Then
            SortDirection = "ASC"
        Else
            'Direction already set, don't do anything
        End If

        Response.Cache.SetCacheability(HttpCacheability.NoCache)
        Dim st As String = Copient.PhraseLib.Lookup("folders.sendoutboundwarning", LanguageID)
        Dim BannersEnabled As Boolean = (MyCommon.Fetch_SystemOption(66) = "1")
        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.Fetch_SystemOption(123) = "1" Then
                MyCommon.QueryStr = " Select Th.ThemeDescription,Fl.StartDate,Fl.EndDate from Folders Fl " &
                                    " left outer join FolderThemes FT on Fl.FolderID=FT.FolderID " &
                                    " left outer join Themes Th on FT.ThemeID=Th.ThemeID where Fl.FolderID=" & FolderID
                dt = MyCommon.LRT_Select
                If dt.Rows.Count > 0 Then
                    Theme = MyCommon.NZ(dt.Rows(0).Item("ThemeDescription"), "Not Selected")
                    If Not IsDBNull(dt.Rows(0).Item("StartDate")) Then
                        FolderStartDate = Format(dt.Rows(0).Item("StartDate"), "MM/dd/yyyy")
                        If FolderStartDate = "01/01/1900" Then FolderStartDate = ""
                    End If
                    If Not IsDBNull(dt.Rows(0).Item("EndDate")) Then
                        FolderEndDate = Format(dt.Rows(0).Item("EndDate"), "MM/dd/yyyy")
                        If FolderEndDate = "01/01/1900" Then FolderEndDate = ""
                    End If
                End If
            End If

            If bStoreUser Then
                sJoin = " Full Outer Join OfferLocUpdate olu with (NoLock) on AOLV.OfferID=olu.OfferID "
                wherestr = " and (LocationID in (" & sValidLocIDs & ") or (CreatedByAdminID in (" & sValidSU & ") and Isnull(LocationID,0)=0)) "
            End If

            Dim PermissionType As Integer = -1
            If (bEnableRestrictedAccessToUEOfferBuilder) Then
                If (Logix.UserRoles.CreateUEOffers AndAlso Not Logix.UserRoles.AccessTranslatedUEOffers) Then
                    PermissionType = 1
                ElseIf (Not Logix.UserRoles.CreateUEOffers AndAlso Logix.UserRoles.AccessTranslatedUEOffers) Then
                    PermissionType = 2
                ElseIf (Not Logix.UserRoles.CreateUEOffers AndAlso Not Logix.UserRoles.AccessTranslatedUEOffers) Then
                    PermissionType = 3
                End If
            End If
            If WFStatus = 0 Then
                MyCommon.QueryStr = "dbo.pa_FolderItem_Select"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@FolderID", SqlDbType.Int).Value = FolderID
                MyCommon.LRTsp.Parameters.Add("@SortText", SqlDbType.NVarChar, 50).Value = If(SortText = "AOLV.Status", "AOLV.OfferID", SortText)
                MyCommon.LRTsp.Parameters.Add("@SortDirection", SqlDbType.NVarChar, 50).Value = SortDirection
                MyCommon.LRTsp.Parameters.Add("@PageIndex", SqlDbType.Int).Value = pageIndex
                MyCommon.LRTsp.Parameters.Add("@PageSize", SqlDbType.Int).Value = 1000
                MyCommon.LRTsp.Parameters.Add("@sJoin", SqlDbType.VarChar, 100).Value = sJoin
                MyCommon.LRTsp.Parameters.Add("@wherestr", SqlDbType.VarChar, 100).Value = wherestr
                MyCommon.LRTsp.Parameters.Add("@PageCount", SqlDbType.Int, 4).Direction = ParameterDirection.Output
                MyCommon.LRTsp.Parameters.Add("@TotalOffers", SqlDbType.Int, 4).Direction = ParameterDirection.Output

                'If ViewOffersRegardlessBuyer permission is not set, get buyer Specific Offer
                If (MyCommon.IsEngineInstalled(9) AndAlso MyCommon.IsUserAssociatedWithAnyBuyer(AdminUserID) AndAlso Not (Logix.UserRoles.ViewOffersRegardlessBuyer) Or bStoreUser) Then
                    MyCommon.LRTsp.Parameters.Add("@AdminUserId", SqlDbType.Int).Value = AdminUserID
                    MyCommon.LRTsp.Parameters.Add("@BannersEnabled", SqlDbType.Bit).Value = BannersEnabled
                    MyCommon.LRTsp.Parameters.Add("@BuyerFilteringEnabled", SqlDbType.Bit).Value = True
                ElseIf BannersEnabled = True Then
                    MyCommon.LRTsp.Parameters.Add("@AdminUserId", SqlDbType.Int).Value = AdminUserID
                    MyCommon.LRTsp.Parameters.Add("@BannersEnabled", SqlDbType.Bit).Value = BannersEnabled
                End If
            Else
                MyCommon.QueryStr = "dbo.pa_FolderItem_Select_WFStatus"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@FolderID", SqlDbType.Int).Value = FolderID
                MyCommon.LRTsp.Parameters.Add("@SortText", SqlDbType.NVarChar, 50).Value = If(SortText = "AOLV.Status", "AOLV.OfferID", SortText)
                MyCommon.LRTsp.Parameters.Add("@SortDirection", SqlDbType.NVarChar, 50).Value = SortDirection
                MyCommon.LRTsp.Parameters.Add("@WFStatus", SqlDbType.Int).Value = WFStatus
                MyCommon.LRTsp.Parameters.Add("@PageIndex", SqlDbType.Int).Value = pageIndex
                MyCommon.LRTsp.Parameters.Add("@PageSize", SqlDbType.Int).Value = 1000
                MyCommon.LRTsp.Parameters.Add("@PageCount", SqlDbType.Int, 4).Direction = ParameterDirection.Output
                MyCommon.LRTsp.Parameters.Add("@TotalOffers", SqlDbType.Int, 4).Direction = ParameterDirection.Output
                If (MyCommon.IsEngineInstalled(9) AndAlso MyCommon.IsUserAssociatedWithAnyBuyer(AdminUserID) AndAlso Not (Logix.UserRoles.ViewOffersRegardlessBuyer)) Then
                    MyCommon.LRTsp.Parameters.Add("@AdminUserId", SqlDbType.Int).Value = AdminUserID
                End If
            End If
            If (bEnableRestrictedAccessToUEOfferBuilder) Then MyCommon.LRTsp.Parameters.Add("@PermissionType", SqlDbType.Int).Value = PermissionType
            dt = MyCommon.LRTsp_select
            Dim TotalOffers = MyCommon.LRTsp.Parameters("@TotalOffers").Value
            MyCommon.Close_LRTsp()

            If (StatusText = "") Then
                If MyCommon.Fetch_SystemOption(123) = "1" Then
                    StatusText = TotalOffers & " " & Copient.PhraseLib.Lookup("term.item(s)", LanguageID) & " Theme: " & Theme & " StartDate: " & FolderStartDate & " EndDate: " & FolderEndDate
                Else
                    StatusText = TotalOffers & " " & Copient.PhraseLib.Lookup("term.item(s)", LanguageID)
                End If
            End If

            TempBuf.AppendLine("<input type=""hidden"" id=""statustext"" name=""statustext"" value=""" & StatusText & """ />")
            If dt.Rows.Count > 0 Then
                TempBuf.AppendLine("<table id = ""tb1"" summary=""" & Copient.PhraseLib.Lookup("term.contents", LanguageID) & """ style=""width:100%;"">")
                'AMS-2289 : To make Infobar's width in sync with Offer name, FolderStatus div is moved inside the table
                If WFStatus <> 0 Then
                    StatusText = StatusText & " | Workflow status filter is ON"
                End If
                'Update the Folder Deploy status 
                Dim status = GetFolderMassOperationStatus(FolderID)
                If status.StartsWith(FOLDER_NOT_IN_USE) Then
                    status = status.Substring(status.LastIndexOf("~") + 1)
                    If status <> String.Empty Then
                        TempBuf.AppendLine("<tr>")
                        TempBuf.AppendLine("<td colspan=""8"">")
                        TempBuf.AppendLine("<div id=""FolderStatus"" style=""color: whitesmoke; background-color: " & IIf(status.StartsWith(Copient.PhraseLib.Lookup("term.success", LanguageID)), "green", "red") & ";"">")
                        TempBuf.AppendLine("<img src=""..\images\desktop\window\close-on.png""  id=""statusClose""/>")
                        TempBuf.AppendLine("<p>" & status & "</p>")
                        TempBuf.AppendLine("</div>")
                        TempBuf.AppendLine("</td>")
                        TempBuf.AppendLine("</tr>")

                        'SetFolderDeployStatus(FolderID, FOLDER_NOT_IN_USE)
                    End If

                End If

                TempBuf.AppendLine("  <tr>")
                TempBuf.Append("    <th style=""width:20px;""><input name=""allitemIDs"" id=""allitemIDs"" type=""checkbox"" title=""" & Copient.PhraseLib.Lookup("hierarchy.SelectAllItems", LanguageID) & """ onclick=""javascript:handleAllItems();"" /></th>")
                TempBuf.Append("    <th style=""width:70px;white-space: nowrap;""><a id=""offeridLink"" href=""javascript:loadFolderItems(" & FolderID & ", '" & SortDirection & "', 'AOLV.OfferID');"" >" & Copient.PhraseLib.Lookup("term.id", LanguageID).Replace(" ", "&nbsp;") & "</a>")
                If SortText = "AOLV.OfferID" Then
                    If SortDirection = "ASC" Then
                        TempBuf.Append("<span class=""sortarrow"">&#9660;</span>")
                    ElseIf SortDirection = "DESC" Then
                        TempBuf.Append("<span class=""sortarrow"">&#9650;</span>")
                    End If
                Else
                End If
                TempBuf.AppendLine("</th>")
                TempBuf.Append("    <th><a id=""xidLink"" href=""javascript:loadFolderItems(" & FolderID & ", '" & SortDirection & "', 'AOLV.ExtOfferID');"" >" & Copient.PhraseLib.Lookup("term.xid", LanguageID) & "</a>")
                If SortText = "AOLV.ExtOfferID" Then
                    If SortDirection = "ASC" Then
                        TempBuf.Append("<span class=""sortarrow"">&#9660;</span>")
                    ElseIf SortDirection = "DESC" Then
                        TempBuf.Append("<span class=""sortarrow"">&#9650;</span>")
                    End If
                Else
                End If
                TempBuf.AppendLine("</th>")
                If (IsEngineInstalled) Then
                    TempBuf.Append("    <th><a id=""promoengineLink"" href=""javascript:loadFolderItems(" & FolderID & ", '" & SortDirection & "', 'AOLV.PromoEngine');"" >" & Copient.PhraseLib.Lookup("term.engine", LanguageID) & "</a>")
                    If SortText = "AOLV.PromoEngine" Then
                        If SortDirection = "ASC" Then
                            TempBuf.Append("<span class=""sortarrow"">&#9660;</span>")
                        ElseIf SortDirection = "DESC" Then
                            TempBuf.Append("<span class=""sortarrow"">&#9650;</span>")
                        End If
                    Else
                    End If
                    TempBuf.AppendLine("</th>")
                End If
                TempBuf.Append("    <th><a id=""nameLink"" href=""javascript:loadFolderItems(" & FolderID & ", '" & SortDirection & "', 'AOLV.Name');"" >" & Copient.PhraseLib.Lookup("term.name", LanguageID) & "</a>")
                If SortText = "AOLV.Name" Then
                    If SortDirection = "ASC" Then
                        TempBuf.Append("<span class=""sortarrow"">&#9660;</span>")
                    ElseIf SortDirection = "DESC" Then
                        TempBuf.Append("<span class=""sortarrow"">&#9650;</span>")
                    End If
                Else
                End If
                TempBuf.AppendLine("</th>")
                TempBuf.Append("    <th><a id=""offerstartdateLink"" href=""javascript:loadFolderItems(" & FolderID & ", '" & SortDirection & "', 'AOLV.ProdStartDate');"" >" & Copient.PhraseLib.Lookup("term.starts", LanguageID) & "</a>")
                If SortText = "AOLV.ProdStartDate" Then
                    If SortDirection = "ASC" Then
                        TempBuf.Append("<span class=""sortarrow"">&#9660;</span>")
                    ElseIf SortDirection = "DESC" Then
                        TempBuf.Append("<span class=""sortarrow"">&#9650;</span>")
                    End If
                Else
                End If
                TempBuf.AppendLine("</th>")
                TempBuf.Append("    <th><a id=""offerenddateLink"" href=""javascript:loadFolderItems(" & FolderID & ", '" & SortDirection & "', 'AOLV.ProdEndDate');"" >" & Copient.PhraseLib.Lookup("term.ends", LanguageID) & "</a>")
                If SortText = "AOLV.ProdEndDate" Then
                    If SortDirection = "ASC" Then
                        TempBuf.Append("<span class=""sortarrow"">&#9660;</span>")
                    ElseIf SortDirection = "DESC" Then
                        TempBuf.Append("<span class=""sortarrow"">&#9650;</span>")
                    End If
                Else
                End If
                TempBuf.AppendLine("</th>")
                TempBuf.Append("    <th><a id=""offerStatusLink"" href=""javascript:loadFolderItems(" & FolderID & ", '" & SortDirection & "', 'AOLV.Status');"">" & Copient.PhraseLib.Lookup("term.status", LanguageID) & "</a>")
                If Request.QueryString("SortText") = "AOLV.Status" Then
                    If SortDirection = "ASC" Then
                        TempBuf.Append("<span class=""sortarrow"">&#9660;</span>")
                    ElseIf SortDirection = "DESC" Then
                        TempBuf.Append("<span class=""sortarrow"">&#9650;</span>")
                    End If
                End If
                TempBuf.AppendLine("</th>")
                TempBuf.AppendLine("  </tr>")
                Dim count As Integer = 1
                Dim Statuses As New Hashtable(20)

                Dim colBValues As String() = {""}
                Dim temp As Int32 = 0
                For Each rown As DataRow In dt.Rows
                    ReDim Preserve colBValues(0 To UBound(colBValues) + 1)
                    colBValues(temp) = (rown.Item("LinkID").ToString())
                    temp = temp + 1
                Next
                Statuses = Logix.GetStatusForOffers(colBValues, LanguageID)

                'Form the Offers table
                TempBuf.AppendLine(GetHTMLOfferTable(dt, False))
                TempBuf.AppendLine("</table>")

                TempBuf.AppendLine("  <Div class=""loading"">")
                TempBuf.AppendLine("  <img id=""loader"" alt=""Loading"" src=""..\images\loader.gif"" style=""display: none;align: centre"" />")
                TempBuf.AppendLine("  </Div>")
            Else
                If (FolderID = 0) Then
                    TempBuf.AppendLine("<p>" & Copient.PhraseLib.Lookup("folders.SelectFolder", LanguageID) & "</p>")
                Else
                    TempBuf.AppendLine("<p>" & Copient.PhraseLib.Lookup("folders.EmptyFolder", LanguageID) & "</p>")
                End If
            End If
            TempBuf.AppendLine("<input type=""hidden"" id=""itemlist"" name=""itemlist"" value=""" & itemIdsAll & """ />")
            Send(TempBuf.ToString)
        Catch ex As Exception
            Send(ex.ToString)
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try
    End Sub

    Public Sub GetSubFolders(ByVal ParentFolderID As Integer)
        Dim Sb As StringBuilder = New StringBuilder()
        MyCommon.QueryStr = "select FolderID from Folders where ParentFolderID = @ParentFolderID"
        MyCommon.DBParameters.Add("@ParentFolderID", SqlDbType.Int).Value = ParentFolderID
        Dim dtsubfolder As DataTable = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        For i = 0 To dtsubfolder.Rows.Count - 1
            Sb.Append(dtsubfolder.Rows(i)(0).ToString())
            If (i < dtsubfolder.Rows.Count - 1) Then
                Sb.Append(",")
            End If
        Next
        Send(Sb.ToString())
    End Sub

    Public Sub GetParentFolders(ByVal FolderID As Integer)
        Dim Sb As StringBuilder = New StringBuilder()
        MyCommon.QueryStr = "select ParentFolderID from Folders where FolderID = @FolderID"
        MyCommon.DBParameters.Add("@FolderID", SqlDbType.Int).Value = FolderID
        Dim dtparentfolder As DataTable = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        For i = 0 To dtparentfolder.Rows.Count - 1
            Sb.Append(dtparentfolder.Rows(i)(0).ToString())
        Next
        Send(Sb.ToString())
    End Sub
    Public Function IsOfferApprovalWorkflowEnabled(ByVal offersDt As DataTable) As Boolean
        Dim isOAWEnabled As Boolean = False
        Dim BannersEnabled As Boolean = (MyCommon.Fetch_SystemOption(66) = "1")

        MyCommon.QueryStr = "dbo.pt_OfferApprovalWorkflow_Enabled"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@BannersEnabled", SqlDbType.Bit).Value = BannersEnabled
        MyCommon.LRTsp.Parameters.Add("@OfferDT", SqlDbType.Structured).Value = offersDt
        MyCommon.LRTsp.Parameters.Add("@Enabled", SqlDbType.Bit).Direction = ParameterDirection.Output
        MyCommon.LRTsp.ExecuteNonQuery()
        isOAWEnabled = MyCommon.LRTsp.Parameters("@Enabled").Value
        MyCommon.Close_LRTsp()

        Return isOAWEnabled
    End Function

    Public Sub ShowRequestApprovalActionItem(ByVal offers As String, ByVal FromOfferList As Boolean)
        Dim offerIds As String() = offers.Split(", ")
        Dim showApprovalActionItem As Integer = -1
        Dim isOAWEnabed As Boolean = False
        Dim TempBuf As New StringBuilder()
        Dim BannersEnabled As Boolean = (MyCommon.Fetch_SystemOption(66) = "1")
        Try
            LogFile = "MassActionLog." & MyCommon.Leading_Zero_Fill(Year(Today), 4) & MyCommon.Leading_Zero_Fill(Month(Today), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
            Dim offersDT As DataTable = GetSelectedOfferIds(offerIds, FromOfferList)
            Dim deployableOffers As DataTable = GetDeployableOffers(offersDT, BannersEnabled)
            Dim nonDeployableOffers As DataTable = GetNonDeployableOffers(offersDT, BannersEnabled)
            Dim pendingApprovalOffers As DataTable = GetPendingApprovalOffers(offersDT)
            If offersDT IsNot Nothing AndAlso offersDT.Rows.Count > 0 Then

                isOAWEnabed = IsOfferApprovalWorkflowEnabled(offersDT)

                If isOAWEnabed Then
                    If pendingApprovalOffers.Rows.Count = offersDT.Rows.Count Then
                        showApprovalActionItem = -1
                    ElseIf deployableOffers.Rows.Count = offersDT.Rows.Count Then
                        showApprovalActionItem = 0
                    ElseIf nonDeployableOffers.Rows.Count = offersDT.Rows.Count Then
                        showApprovalActionItem = 1
                    ElseIf pendingApprovalOffers.Rows.Count = 0 Then
                        showApprovalActionItem = 2
                    ElseIf deployableOffers.Rows.Count = 0 Then
                        showApprovalActionItem = 4
                    ElseIf nonDeployableOffers.Rows.Count = 0 Then
                        showApprovalActionItem = 3
                    ElseIf offersDT.Rows.Count = (pendingApprovalOffers.Rows.Count + nonDeployableOffers.Rows.Count + deployableOffers.Rows.Count) Then
                        showApprovalActionItem = 2
                    End If
                Else
                    showApprovalActionItem = 0
                End If

            End If
        Catch ex As Exception
            MyCommon.Write_Log(LogFile, ex.ToString())
        Finally
            Send(showApprovalActionItem)
        End Try


    End Sub

    Function GetSelectedOfferIds(ByVal itemIds As String(), ByVal FromOfferList As Boolean) As DataTable
        Dim offersDT As DataTable = New DataTable()
        offersDT.Columns.Add("OfferID")

        If FromOfferList Then
            For Each itemId In itemIds
                offersDT.Rows.Add(itemId)
            Next
        Else
            Dim itemsDT As DataTable = New DataTable()
            itemsDT.Columns.Add("OfferID")
            For Each itemId In itemIds
                itemsDT.Rows.Add(itemId)
            Next
            MyCommon.QueryStr = "dbo.pt_GetOfferID_FolderItems"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@OffersDT", SqlDbType.Structured).Value = itemsDT
            offersDT = MyCommon.LRTsp_select
            MyCommon.Close_LRTsp()
        End If


        Return offersDT
    End Function

    Public Sub HandleDeployAndApprovalRequest(ByVal itemIds As String, ByVal FromOfferList As Boolean, ByVal deploymentType As Integer, ByVal FolderID As Integer, ByVal OffersWithoutConditions As Boolean, ByVal deploytransreqskip As String, ByVal deferdeploy As String)
        Dim items As String() = itemIds.Split(", ")
        Dim offersDT As DataTable = GetSelectedOfferIds(items, FromOfferList)
        Dim deployableOffers As DataTable = GetDeployableOffers(offersDT, MyCommon.Fetch_SystemOption(66))
        Dim nonDeployableOffers As DataTable = GetNonDeployableOffers(offersDT, MyCommon.Fetch_SystemOption(66))
        Dim status As String = ""
        Dim status1 As String = ""
        Dim deployItems As String = ""
        Dim nonDeployItems As String = ""
        Dim isPendingOffersExists As Boolean = IIf(offersDT.Rows.Count = (deployableOffers.Rows.Count + nonDeployableOffers.Rows.Count), False, True)
        Try
            For Each row In deployableOffers.Rows
                If deployItems = "" Then
                    deployItems = row("OfferId").ToString()
                Else
                    deployItems &= ", " & row("OfferId").ToString()
                End If
            Next
            For Each row In nonDeployableOffers.Rows
                If nonDeployItems = "" Then
                    nonDeployItems = row("OfferId").ToString()
                Else
                    nonDeployItems &= ", " & row("OfferId").ToString()
                End If
            Next

            Select Case (deploymentType)
                Case 1 'deploy only deployable offers
                    DeployMixOffers(deployItems, FromOfferList, FolderID, OffersWithoutConditions, deploytransreqskip, deferdeploy, isPendingOffersExists)

                Case 2 'only request approval
                    RequestApprovalForMixOffers(nonDeployItems, FromOfferList, FolderID, 13, isPendingOffersExists)

                Case 3 'Only Request Approval with Deploy
                    RequestApprovalForMixOffers(nonDeployItems, FromOfferList, FolderID, 14, isPendingOffersExists)

                Case 4 'only request approval with defer deploy
                    RequestApprovalForMixOffers(nonDeployItems, FromOfferList, FolderID, 15, isPendingOffersExists)

                Case 5 'deploy and request approval
                    RequestApprovalAndDeployOffers(13, deployItems, nonDeployItems, FromOfferList, FolderID, OffersWithoutConditions, deploytransreqskip, deferdeploy, isPendingOffersExists)

                Case 6 'deploy and request approval with deploy
                    RequestApprovalAndDeployOffers(14, deployItems, nonDeployItems, FromOfferList, FolderID, OffersWithoutConditions, deploytransreqskip, deferdeploy, isPendingOffersExists)

                Case 7 'deploy and request approval with defer deploy
                    RequestApprovalAndDeployOffers(15, deployItems, nonDeployItems, FromOfferList, FolderID, OffersWithoutConditions, deploytransreqskip, deferdeploy, isPendingOffersExists)

                Case 8 'defer deploy only deployable offers
                    DeferDeployMixOffers(deployItems, FromOfferList, FolderID, OffersWithoutConditions, deploytransreqskip, deferdeploy, isPendingOffersExists)

                Case 9 'defer deploy and request approval
                    RequestApprovalAndDeferDeployOffers(13, deployItems, nonDeployItems, FolderID, OffersWithoutConditions, deploytransreqskip, deferdeploy, isPendingOffersExists)

                Case 10 'defer deploy and request approval with deploy
                    RequestApprovalAndDeferDeployOffers(14, deployItems, nonDeployItems, FolderID, OffersWithoutConditions, deploytransreqskip, deferdeploy, isPendingOffersExists)

                Case 11 'defer deploy and request approval with defer deploy
                    RequestApprovalAndDeferDeployOffers(15, deployItems, nonDeployItems, FolderID, OffersWithoutConditions, deploytransreqskip, deferdeploy, isPendingOffersExists)

            End Select
        Catch ex As Exception
            MyCommon.Write_Log(LogFile, ex.ToString())
        End Try
    End Sub

    Sub DeployMixOffers(ByVal itemIds As String, ByVal fromOfferList As Boolean, ByVal FolderID As Integer, ByVal OffersWithoutConditions As Boolean, ByVal deploytransreqskip As String, ByVal deferdeploy As String, ByVal isPendingOffersExists As Boolean)
        Dim status As String = String.Empty
        If (fromOfferList) Then
            SetFolderMassOperationStatus(-1, FOLDER_IN_USE + Copient.PhraseLib.Lookup("offers.massdeploy", LanguageID) & AdminName)
        Else
            SetFolderMassOperationStatus(FolderID, FOLDER_IN_USE + Copient.PhraseLib.Lookup("folders.massdeploy", LanguageID) & AdminName)
        End If
        MassDeployOffers(itemIds, True, OffersWithoutConditions, AdminUserID, deploytransreqskip, deferdeploy, FolderID, status, isPendingOffersExists)
        If (fromOfferList) Then
            SetFolderMassOperationStatus(-1, FOLDER_NOT_IN_USE + status)
        ElseIf (CollisionsDetected = True) Then
            SetFolderMassOperationStatus(FolderID, FOLDER_NOT_IN_USE + Copient.PhraseLib.Lookup("term.folderoffers-failed-collision", LanguageID))
        Else
            SetFolderMassOperationStatus(FolderID, FOLDER_NOT_IN_USE + status)
        End If
    End Sub

    Sub DeferDeployMixOffers(ByVal itemIds As String, ByVal fromOfferList As Boolean, ByVal FolderID As Integer, ByVal OffersWithoutConditions As Boolean, ByVal deploytransreqskip As String, ByVal deferdeploy As String, ByVal isPendingOffersExists As Boolean)
        Dim status As String = String.Empty
        SetFolderMassOperationStatus(FolderID, FOLDER_IN_USE + Copient.PhraseLib.Lookup("folders.massdeferdeploy", LanguageID) & AdminName)
        MassDeferDeployOffers(itemIds, True, OffersWithoutConditions, AdminUserID, deploytransreqskip, deferdeploy, FolderID, status, isPendingOffersExists)
        If (CollisionsDetected = True) Then
            SetFolderMassOperationStatus(FolderID, FOLDER_NOT_IN_USE + Copient.PhraseLib.Lookup("term.folderoffers-failed-collision", LanguageID))
        Else
            SetFolderMassOperationStatus(FolderID, FOLDER_NOT_IN_USE + status)
        End If
    End Sub

    Sub RequestApprovalForMixOffers(ByVal itemIds As String, ByVal fromOfferList As Boolean, ByVal FolderID As Integer, ByVal approvalType As Integer, ByVal isPendingOffersExists As Boolean)
        Dim status As String = String.Empty
        If (fromOfferList) Then
            SetFolderMassOperationStatus(-1, FOLDER_IN_USE + Copient.PhraseLib.Lookup("offers.massapproval", LanguageID) & AdminName)
        Else
            SetFolderMassOperationStatus(FolderID, FOLDER_IN_USE + Copient.PhraseLib.Lookup("folders.massapproval", LanguageID) & AdminName)
        End If
        MassRequestApproval(approvalType, itemIds, True, FolderID, status, isPendingOffersExists)
        If (fromOfferList) Then
            SetFolderMassOperationStatus(-1, FOLDER_NOT_IN_USE + status)
        ElseIf (CollisionsDetected = True) Then
            SetFolderMassOperationStatus(FolderID, FOLDER_NOT_IN_USE + Copient.PhraseLib.Lookup("term.folderoffers-failed-collision", LanguageID))
        Else
            SetFolderMassOperationStatus(FolderID, FOLDER_NOT_IN_USE + status)
        End If
    End Sub

    Sub RequestApprovalAndDeployOffers(ByVal approvalType As Integer, ByVal deployItems As String, ByVal nonDeployItems As String, ByVal fromOfferList As Boolean, ByVal folderId As Integer, ByVal OffersWithoutConditions As Boolean, ByVal deploytransreqskip As String, ByVal deferdeploy As String, ByVal isPendingOffersExists As Boolean)
        Dim status As String = String.Empty
        Dim status1 As String = String.Empty
        If (fromOfferList) Then
            SetFolderMassOperationStatus(-1, FOLDER_IN_USE + Copient.PhraseLib.Lookup("offers.massdeployapproval", LanguageID) & AdminName)
        Else
            SetFolderMassOperationStatus(folderId, FOLDER_IN_USE + Copient.PhraseLib.Lookup("folders.massdeployapproval", LanguageID) & AdminName)
        End If
        MassDeployOffers(deployItems, True, OffersWithoutConditions, AdminUserID, deploytransreqskip, deferdeploy, folderId, status)
        MassRequestApproval(approvalType, nonDeployItems, True, folderId, status1)
        If status.Contains("Success") AndAlso status1.Contains("Success") Then
            status = Copient.PhraseLib.Lookup("term.success1", LanguageID) + " " + Copient.PhraseLib.Lookup("offers.deployapproval-processed", LanguageID)
            If isPendingOffersExists Then
                status &= " " + Copient.PhraseLib.Lookup("term.ignore-pendingoffers", LanguageID)
            End If
        ElseIf Not status1.Contains("Success") Then
            status = status1
        End If
        If (fromOfferList) Then
            SetFolderMassOperationStatus(-1, FOLDER_NOT_IN_USE + status)
        ElseIf (CollisionsDetected = True) Then
            SetFolderMassOperationStatus(folderId, FOLDER_NOT_IN_USE + Copient.PhraseLib.Lookup("term.folderoffers-failed-collision", LanguageID))
        Else
            SetFolderMassOperationStatus(folderId, FOLDER_NOT_IN_USE + status)
        End If
    End Sub
    Sub RequestApprovalAndDeferDeployOffers(ByVal approvalType As Integer, ByVal deployItems As String, ByVal nonDeployItems As String, ByVal folderId As Integer, ByVal OffersWithoutConditions As Boolean, ByVal deploytransreqskip As String, ByVal deferdeploy As String, ByVal isPendingOffersExists As Boolean)
        Dim status As String = String.Empty
        Dim status1 As String = String.Empty
        SetFolderMassOperationStatus(folderId, FOLDER_IN_USE + Copient.PhraseLib.Lookup("folders.massapprovaldeferdeploy", LanguageID) & AdminName)
        MassDeferDeployOffers(deployItems, True, OffersWithoutConditions, AdminUserID, deploytransreqskip, deferdeploy, folderId, status)
        MassRequestApproval(approvalType, nonDeployItems, True, folderId, status1)
        If status.Contains("Success") AndAlso status1.Contains("Success") Then
            status = Copient.PhraseLib.Lookup("term.success1", LanguageID) + " " + Copient.PhraseLib.Lookup("offers.deferdeployapproval-processed", LanguageID)
            If isPendingOffersExists Then
                status &= " " + Copient.PhraseLib.Lookup("term.ignore-pendingoffers", LanguageID)
            End If
        ElseIf Not status1.Contains("Success") Then
            status = status1
        End If
        If (CollisionsDetected = True) Then
            SetFolderMassOperationStatus(folderId, FOLDER_NOT_IN_USE + Copient.PhraseLib.Lookup("term.folderoffers-failed-collision", LanguageID))
        Else
            SetFolderMassOperationStatus(folderId, FOLDER_NOT_IN_USE + status)
        End If
    End Sub
    Function GetDeployableOffers(ByVal offersDT As DataTable, ByVal BannersEnabled As Boolean) As DataTable
        Dim deployableOffers As DataTable = New DataTable()
        MyCommon.QueryStr = "dbo.pt_GetDeployableOffers"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@BannersEnabled", SqlDbType.Bit).Value = BannersEnabled
        MyCommon.LRTsp.Parameters.Add("@AdminUserID", SqlDbType.Int).Value = AdminUserID
        MyCommon.LRTsp.Parameters.Add("@OfferDT", SqlDbType.Structured).Value = offersDT
        deployableOffers = MyCommon.LRTsp_select
        MyCommon.Close_LRTsp()

        Return deployableOffers
    End Function

    Function GetNonDeployableOffers(ByVal offersDT As DataTable, ByVal BannersEnabled As Boolean) As DataTable
        Dim nondeployableOffers As DataTable = New DataTable()
        MyCommon.QueryStr = "dbo.pt_GetNonDeployableOffers"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@BannersEnabled", SqlDbType.Bit).Value = BannersEnabled
        MyCommon.LRTsp.Parameters.Add("@AdminUserID", SqlDbType.Int).Value = AdminUserID
        MyCommon.LRTsp.Parameters.Add("@OfferDT", SqlDbType.Structured).Value = offersDT
        nondeployableOffers = MyCommon.LRTsp_select
        MyCommon.Close_LRTsp()

        Return nondeployableOffers
    End Function

    Function GetPendingApprovalOffers(ByVal offersDT As DataTable) As DataTable
        Dim pendingapprovalOffers As DataTable = New DataTable()
        MyCommon.QueryStr = "dbo.pt_GetOffersPendingApproval"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@AdminUserID", SqlDbType.Int).Value = AdminUserID
        MyCommon.LRTsp.Parameters.Add("@OfferDT", SqlDbType.Structured).Value = offersDT
        pendingapprovalOffers = MyCommon.LRTsp_select
        MyCommon.Close_LRTsp()

        Return pendingapprovalOffers
    End Function

    Public Sub MassRequestApprovalWorker(ByVal State As Object)
        Dim obj As Object() = State
        Dim approvalType As Integer = obj(0)
        Dim ItemIDs As String = obj(1)
        Dim FromOfferList As Boolean = obj(2)
        Dim folderId As Integer = obj(3)
        Dim Status As String = String.Empty
        Dim err As String = String.Empty
        Try
            If (FromOfferList) Then
                SetFolderMassOperationStatus(-1, FOLDER_IN_USE + Copient.PhraseLib.Lookup("folders.massapproval", LanguageID) & AdminName)
            Else
                SetFolderMassOperationStatus(folderId, FOLDER_IN_USE + Copient.PhraseLib.Lookup("folders.massapproval", LanguageID) & AdminName)
            End If
            MassRequestApproval(approvalType, ItemIDs, FromOfferList, folderId, Status)
        Catch ex As Exception
            err = "Error: " + ex.ToString
            MyCommon.Write_Log(LogFile, err)
        Finally
            If (FromOfferList) Then
                SetFolderMassOperationStatus(-1, FOLDER_NOT_IN_USE + Status)
            ElseIf (CollisionsDetected = True) Then
                SetFolderMassOperationStatus(FolderID, FOLDER_NOT_IN_USE + Copient.PhraseLib.Lookup("term.folderoffers-failed-collision", LanguageID))
            Else
                SetFolderMassOperationStatus(FolderID, FOLDER_NOT_IN_USE + Status)
            End If
        End Try
    End Sub

    Public Sub MassRequestApproval(ByVal approvalType As Integer, ByVal itemIDs As String, ByVal fromOfferList As Boolean, ByVal folderId As Integer, Optional ByRef status As String = "", Optional ByVal isPendingOffersExists As Boolean = False)
        LogFile = "MassActionLog." & MyCommon.Leading_Zero_Fill(Year(Today), 4) & MyCommon.Leading_Zero_Fill(Month(Today), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
        MyCommon.Write_Log(LogFile, "----------------------------------------------------------------")
        Dim requestResolver = New ResolverBuilder()
        requestResolver.Build()
        Try
            RegisterDependencies(requestResolver)

            Dim offerDeploymentValidator As IOfferDeploymentValidator = requestResolver.Container.Resolve(Of IOfferDeploymentValidator)()
            Dim offerService As IOffer = requestResolver.Container.Resolve(Of IOffer)()
            Dim oawService As IOfferApprovalWorkflowService = requestResolver.Container.Resolve(Of OfferApprovalWorkflowService)()

            Dim ValidationResult As AMSResult(Of DataTable) = New AMSResult(Of DataTable)()
            Dim OffersWithEnginesDT As DataTable
            Dim IsValid As Boolean = False


            'Get the dictionary of OfferIds with corresponding Engine ID 
            OffersWithEnginesDT = GetEngineID(itemIDs, fromOfferList)

            MyCommon.Write_Log(LogFile, "Performed Action(Request Approval) on " & OffersWithEnginesDT.Rows.Count & " offers from Offer List.", True)

            ValidationResult = offerDeploymentValidator.ValidateOffers(OffersWithEnginesDT, False, False, True, True, False, folderId, AdminUserId:=AdminUserID, LangID:=LanguageID)
            If ValidationResult.ResultType = AMSResultType.Success Then
                IsValid = True
            End If
            If (ValidationResult.Result.Rows.Count > 0) Then
                Dim InvalidOfferStr As String = ValidationResult.Result.Rows(0).Item("ReturnMessage").ToString().TrimEnd(",")
                Dim ValidOfferStr As String = ValidationResult.Result.Rows(1).Item("ReturnMessage").ToString().TrimEnd(",")
                Dim ValidOfferList As New List(Of Int64)
                If ValidOfferStr.Length > 0 Then
                    ValidOfferList = ValidOfferStr.Split(",").Select(Function(x) Int64.Parse(x)).ToList()
                End If
                Dim validOffers As New List(Of Int64)
                Dim lstCollidingOffers As AMSResult(Of List(Of Int64))

                If ValidOfferStr.Length > 0 Then
                    MyCommon.Write_Log(LogFile, String.Format("Collision Detection Initiated for Offer IDs: {0}", ValidOfferStr), True)
                    lstCollidingOffers = offerService.ProcessOfferCollisionDetectionFolderDeployment(ValidOfferStr)
                    MyCommon.Write_Log(LogFile, String.Format("Collision Detection Completed for Offer IDs: {0}", ValidOfferStr), True)
                    If lstCollidingOffers.ResultType = AMSResultType.Success AndAlso lstCollidingOffers.Result IsNot Nothing Then
                        CollisionsDetected = (lstCollidingOffers.Result.Count > 0)
                        validOffers = ValidOfferList.Except(lstCollidingOffers.Result).ToList()
                    ElseIf lstCollidingOffers.Result Is Nothing Then
                        validOffers = ValidOfferList
                    End If
                End If
                If validOffers.Count > 0 Then
                    Dim offersDT As DataTable = New DataTable()
                    offersDT.Columns.Add("OfferID")
                    For Each offer In validOffers
                        offersDT.Rows.Add(offer)
                    Next
                    MyCommon.QueryStr = "dbo.pt_RequestApproval_Offers"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@AdminUserID", SqlDbType.Int).Value = AdminUserID
                    MyCommon.LRTsp.Parameters.Add("@ApprovalType", SqlDbType.Int).Value = approvalType
                    MyCommon.LRTsp.Parameters.Add("@OffersDT", SqlDbType.Structured).Value = offersDT
                    MyCommon.LRTsp.ExecuteNonQuery()
                    MyCommon.Close_LRTsp()
                    oawService.SendNotificationEmail(1, 0, AdminUserID, , GetOfferApproversList(offersDT))
                End If
            End If
            If (IsValid) Then
                status = Copient.PhraseLib.Lookup("term.success1", LanguageID) + " " + Copient.PhraseLib.Lookup("offers.approval-processed", LanguageID)
                If isPendingOffersExists Then
                    status &= " " + Copient.PhraseLib.Lookup("term.ignore-pendingoffers", LanguageID)
                End If
            Else
                status = Copient.PhraseLib.Lookup("offers.approval-failed", LanguageID)

            End If
        Catch ex As Exception
            MyCommon.Write_Log(LogFile, ex.ToString(), True)
            status = ex.ToString()
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            requestResolver.Container.Dispose()
        End Try

    End Sub
    Function GetOfferApproversList(ByVal offersDT As DataTable) As List(Of Integer)
        Dim approvers As List(Of Integer) = New List(Of Integer)()
        MyCommon.QueryStr = "dbo.pt_GetOfferApproversList"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@AdminUserID", SqlDbType.Int).Value = AdminUserID
        MyCommon.LRTsp.Parameters.Add("@OffersDT", SqlDbType.Structured).Value = offersDT
        Dim approversDT As DataTable = MyCommon.LRTsp_select
        MyCommon.Close_LRTsp()
        If approversDT IsNot Nothing Then
            For Each row In approversDT.Rows
                approvers.Add(Integer.Parse(row("AdminUserId")))
            Next
        End If
        Return approvers
    End Function
    ''' <summary>
    ''' This Sub gets the all the ItemIds of the Folder.
    ''' </summary>
    ''' <param name="FolderID"></param>
    ''' <param name="WFStatus"></param>
    ''' <remarks></remarks>
    Public Sub GetFolderItemsList(ByVal FolderID As Long, Optional ByVal WFStatus As Integer = -1)
        Dim dt As DataTable
        Dim i As Int32
        Dim Sb As StringBuilder = New StringBuilder()
        Dim BannersEnabled As Boolean = (MyCommon.Fetch_SystemOption(66) = "1")
        MyCommon.QueryStr = "dbo.pa_GetFolderItems"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@FolderID", SqlDbType.Int).Value = FolderID

        'If Permissions are set
        If (MyCommon.IsEngineInstalled(9) AndAlso MyCommon.IsUserAssociatedWithAnyBuyer(AdminUserID) AndAlso Not (Logix.UserRoles.ViewOffersRegardlessBuyer)) Then
            MyCommon.LRTsp.Parameters.Add("@AdminUserId", SqlDbType.Int).Value = AdminUserID
        End If

        'If Permissions are set View is there and Edit is not there, then bring only the respective user Items
        If (MyCommon.IsEngineInstalled(9) AndAlso MyCommon.IsUserAssociatedWithAnyBuyer(AdminUserID) AndAlso (Logix.UserRoles.ViewOffersRegardlessBuyer) AndAlso Not (Logix.UserRoles.EditOffersRegardlessBuyer)) Then
            MyCommon.LRTsp.Parameters.Add("@AdminUserId", SqlDbType.Int).Value = AdminUserID
        End If

        If WFStatus <> 0 Then
            MyCommon.LRTsp.Parameters.Add("@WFStatus", SqlDbType.Int).Value = WFStatus
        End If
        If BannersEnabled Then
            If Not MyCommon.LRTsp.Parameters.Contains("@AdminUserId") Then
                MyCommon.LRTsp.Parameters.Add("@AdminUserId", SqlDbType.Int).Value = AdminUserID
            End If
            MyCommon.LRTsp.Parameters.Add("@BannerEnabled", SqlDbType.Bit).Value = BannersEnabled
        End If
        'Gets all the Item IDs
        dt = MyCommon.LRTsp_select
        MyCommon.Close_LRTsp()
        'Form all the ItemIds as a comma seperated string and send back
        For i = 0 To dt.Rows.Count - 1
            Sb.Append(dt.Rows(i)(0).ToString())
            If (i < dt.Rows.Count - 1) Then
                Sb.Append(",")
            End If
        Next
        Send(Sb.ToString())
    End Sub


    ''' <summary>
    ''' Forms the HTML table from given Datatable
    ''' </summary>
    ''' <param name="dt"></param>
    ''' <param name="checked">checkbox in Data row can be checked?</param>
    ''' <returns>HTML table corresponding to given DataTable</returns>
    ''' <remarks></remarks>
    Public Function GetHTMLOfferTable(ByVal dt As DataTable, Optional ByVal checked As Boolean = False) As String
        Dim TempBuf As New StringBuilder()
        Dim ItemID As Integer
        Dim BuyerId As Integer
        Dim IsCheckboxDisabled As Boolean = False
        Dim PromoEngine As String
        Dim Name As String = ""
        Dim SortDirection = "ASC"
        Dim SortText = "AOLV.OfferID"
        Dim Theme As String = ""
        Dim FolderStartDate As String = ""
        Dim FolderEndDate As String = ""
        Dim lastDeployValidationTag As String
        Dim lastDeployValidationMessage As String
        Dim lastMassUpdateValidationTag As String
        Dim overrideColorStyle As String
        Dim Statuses As New Hashtable(20)
        Dim OfferStatus1 As Copient.LogixInc.STATUS_FLAGS
        Dim count As Integer = 1
        Dim IsEngineInstalled As Boolean = MyCommon.GetInstalledEngines().Length > 1
        Dim colBValues As String() = {""}
        Dim temp As Int32 = 0
        Dim CollisionReportID As Int32 = 0
        Dim EngineID As Int32 = 0
        Dim DisplayOfferLinkPopup As Boolean = False
        Dim IsViewEnabledAndEditDisabled As Boolean = False
        Dim CollisionReportEnabled As Boolean = False

        For Each rown As DataRow In dt.Rows
            ReDim Preserve colBValues(0 To UBound(colBValues) + 1)
            colBValues(temp) = (rown.Item("LinkID").ToString())
            temp = temp + 1
        Next
        Statuses = Logix.GetStatusForOffers(colBValues, LanguageID)
        If (Request.QueryString("SortText") = "AOLV.Status") Then
            Dim sortOrder As String = Request.QueryString("SortDirection")
            If (sortOrder = "ASC") Then
                sortOrder = "DESC"
            Else
                sortOrder = "ASC"
            End If
            dt = Logix.SortOfferStatuses(dt, Statuses, sortOrder)
        End If
        If (MyCommon.IsEngineInstalled(9) AndAlso MyCommon.IsUserAssociatedWithAnyBuyer(AdminUserID) AndAlso (Logix.UserRoles.ViewOffersRegardlessBuyer) AndAlso Not (Logix.UserRoles.EditOffersRegardlessBuyer) Or bStoreUser) Then
            IsViewEnabledAndEditDisabled = True
        Else
            IsViewEnabledAndEditDisabled = False
        End If
        For Each row In dt.Rows
            lastDeployValidationMessage = String.Empty
            overrideColorStyle = String.Empty
            IsCheckboxDisabled = False
            If (IsViewEnabledAndEditDisabled) Then
                BuyerId = MyCommon.NZ(row.Item("BuyerId"), -1)
                'If offer is not created with buyer or current users is not associated with Offer's Buyer, disable the offer
                If BuyerId = -1 Or Not BuyersAssociated.Contains(BuyerId) Then
                    IsCheckboxDisabled = True
                End If
            End If
            ItemID = MyCommon.NZ(row.Item("ItemID"), 0)
            PromoEngine = MyCommon.NZ(row.Item("PromoEngine"), "")
            Name = MyCommon.NZ(row.Item("Name"), "")
            lastDeployValidationTag = MyCommon.NZ(row.Item("LastDeployValidationMessage"), String.Empty)
            'AMS-14135: Fix to remove <font> tags from validation message. This was added in ueoffer-sum validate offer method.
            If Not String.IsNullOrEmpty(lastDeployValidationTag) AndAlso lastDeployValidationTag.Contains("font color") Then
                Dim tags() As String = {""">", "</"}
                lastDeployValidationTag = lastDeployValidationTag.Split(tags, StringSplitOptions.None)(1)
            End If
            lastMassUpdateValidationTag = MyCommon.NZ(row.Item("LastMassUpdateValidationMessage"), String.Empty)
            CollisionReportID = MyCommon.NZ(row.Item("CollisionReportID"), 0)
            EngineID = MyCommon.NZ(row.Item("EngineID"), 0)

            If (lastDeployValidationTag <> String.Empty And lastDeployValidationTag <> "term.validationsuccessful") OrElse (EngineID = 9 AndAlso CollisionReportID <> 0) Then
                'Dim len As Integer = lastDeployValidationTag.IndexOf(">")
                'lastDeployValidationMessage = lastDeployValidationTag.Substring(len + 1, lastDeployValidationTag.IndexOf("</font>") - len - 1)
                overrideColorStyle = " style=""color:red;"
                DisplayOfferLinkPopup = True
            Else
                DisplayOfferLinkPopup = False
            End If
            If (lastDeployValidationTag <> String.Empty And lastDeployValidationTag <> "term.validationsuccessful") Then
                'Dim len As Integer = lastDeployValidationTag.IndexOf(">")
                lastDeployValidationMessage = lastDeployValidationTag
                overrideColorStyle = " style=""color:red;"
                DisplayOfferLinkPopup = True
            ElseIf (lastMassUpdateValidationTag <> String.Empty And lastMassUpdateValidationTag <> "term.validationsuccessful") Then
                lastDeployValidationMessage = lastMassUpdateValidationTag
                overrideColorStyle = " style=""color:red;"
                DisplayOfferLinkPopup = True
            End If

            TempBuf.AppendLine("  <tr id=""OfferRecord" & ItemID & """ title=""" & lastDeployValidationMessage & """" & overrideColorStyle & """>")
            TempBuf.AppendLine("    <td style=""text-align:center;"">")
            TempBuf.AppendLine("      <input name=""itemID"" id=""itemID" & ItemID & """ type=""checkbox"" " & IIf(IsCheckboxDisabled, "disabled=""disabled""", """") & IIf(checked, "checked=""checked""", """") & "value=""" & PromoEngine & """ onclick=""submitToRemoveItems(" & ItemID & ",this.value, this.checked);"" />")
            TempBuf.AppendLine("    </td>")
            TempBuf.AppendLine("    <td>" & MyCommon.NZ(row.Item("LinkID"), "") & "</td>")
            TempBuf.AppendLine("    <td >" & MyCommon.NZ(row.Item("ExtOfferID"), "") & "</td>")

            If (IsEngineInstalled) Then
                TempBuf.AppendLine("    <td style=""white-space: nowrap;"">" & Copient.PhraseLib.Lookup(CInt(MyCommon.NZ(row.Item("EnginePhraseID"), 0)), LanguageID, MyCommon.NZ(row.Item("PromoEngine"), "")) & "</td>")
            End If

            If DisplayOfferLinkPopup = True Then
                CollisionReportEnabled = (CollisionReportID <> 0)
                Dim OfferValidationStr As String
                If CollisionReportEnabled Then
                    OfferValidationStr = Copient.PhraseLib.Lookup("term.offer-collisiondetected", LanguageID).Replace("&#39;", "\'")
                Else
                    OfferValidationStr = Copient.PhraseLib.Lookup("term.offer-predeploymentvalidation-failed", LanguageID).Replace("&#39;", "\'")
                End If
                TempBuf.AppendLine("    <td  style=""white-space: nowrap;""><a href=""#"" onclick=""javascript:return loadOfferNavigationDialog('" + MyCommon.NZ(row.Item("LinkID"), "").ToString() + "', '" + CollisionReportEnabled.ToString() + "', '" + OfferValidationStr + "');"" target=""_blank""" & overrideColorStyle & """>" & IIf(Name <> "", Name, "(unnamed)") & "</a></td>")
            Else
                TempBuf.AppendLine("    <td  style=""white-space: nowrap;""><a href=""offer-redirect.aspx?OfferID=" & MyCommon.NZ(row.Item("LinkID"), "") & """ target=""_blank""" & overrideColorStyle & """>" & IIf(Name <> "", Name, "(unnamed)") & "</a></td>")
            End If

            If IsDBNull(row.Item("ProdStartDate")) Then
                TempBuf.AppendLine("    <td ></td>")
            Else
                TempBuf.AppendLine("    <td >" & Logix.ToShortDateStringQuick(row.Item("ProdStartDate"), MyCommon) & "</td>")
            End If
            If IsDBNull(row.Item("ProdEndDate")) Then
                TempBuf.AppendLine("    <td ></td>")
            Else
                TempBuf.AppendLine("    <td >" & Logix.ToShortDateStringQuick(row.Item("ProdEndDate"), MyCommon) & "</td>")
            End If

            OfferStatus1 = Statuses.Item(row.Item("LinkID").ToString())
            TempBuf.AppendLine("    <td >" & Logix.GetOfferStatusHtml(OfferStatus1, LanguageID) & "</td>")
            TempBuf.AppendLine("  </tr>")
            TempBuf.AppendLine("<tr name=""errdesc"" id=""errdesc" & ItemID & """>")
            TempBuf.AppendLine("</tr>")
        Next
        Return TempBuf.ToString
    End Function

    Public Function GetOffersDataset(ByVal FolderID As Long, ByVal pageIndex As Integer, Optional ByVal StatusText As String = "", Optional ByVal WFStatus As Integer = 0) As DataSet
        Dim TempBuf As New StringBuilder()
        Dim Name As String = ""
        Dim SortDirection = "ASC"
        Dim SortText = "AOLV.OfferID"
        Dim Theme As String = ""
        Dim FolderStartDate As String = ""
        Dim FolderEndDate As String = ""
        Dim itemIdsAll As String = ""
        Dim IsEngineInstalled As Boolean = MyCommon.GetInstalledEngines().Length > 1
        Dim dt As DataTable = New DataTable
        Dim BannersEnabled As Boolean = (MyCommon.Fetch_SystemOption(66) = "1")

        'Set direction and orderby text, if any
        If (GetCgiValue("SortText") <> "") Then
            SortText = GetCgiValue("SortText")
        Else
            'text already set, don't do anything
        End If

        If (GetCgiValue("SortDirection") = "ASC") Then
            SortDirection = "DESC"
        ElseIf (GetCgiValue("SortDirection") = "DESC") Then
            SortDirection = "ASC"
        Else
            'Direction already set, don't do anything
        End If

        Response.Cache.SetCacheability(HttpCacheability.NoCache)

        If WFStatus = 0 Then
            MyCommon.QueryStr = "dbo.pa_FolderItem_Select"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@FolderID", SqlDbType.Int).Value = FolderID
            MyCommon.LRTsp.Parameters.Add("@SortText", SqlDbType.NVarChar, 50).Value = If(SortText = "AOLV.Status", "AOLV.OfferID", SortText)
            MyCommon.LRTsp.Parameters.Add("@SortDirection", SqlDbType.NVarChar, 50).Value = SortDirection
            MyCommon.LRTsp.Parameters.Add("@PageIndex", SqlDbType.Int).Value = pageIndex
            MyCommon.LRTsp.Parameters.Add("@PageSize", SqlDbType.Int).Value = 1000
            MyCommon.LRTsp.Parameters.Add("@PageCount", SqlDbType.Int, 4).Direction = ParameterDirection.Output
            MyCommon.LRTsp.Parameters.Add("@TotalOffers", SqlDbType.Int, 4).Direction = ParameterDirection.Output
            'If ViewOffersRegardlessBuyer permission is not set, get buyer Specific Offer
            If (MyCommon.IsEngineInstalled(9) AndAlso MyCommon.IsUserAssociatedWithAnyBuyer(AdminUserID) AndAlso Not (Logix.UserRoles.ViewOffersRegardlessBuyer) Or bStoreUser) Then
                MyCommon.LRTsp.Parameters.Add("@AdminUserId", SqlDbType.Int).Value = AdminUserID
                MyCommon.LRTsp.Parameters.Add("@BannersEnabled", SqlDbType.Bit).Value = BannersEnabled
                MyCommon.LRTsp.Parameters.Add("@BuyerFilteringEnabled", SqlDbType.Bit).Value = True
            ElseIf BannersEnabled = True Then
                MyCommon.LRTsp.Parameters.Add("@AdminUserId", SqlDbType.Int).Value = AdminUserID
                MyCommon.LRTsp.Parameters.Add("@BannersEnabled", SqlDbType.Bit).Value = BannersEnabled
            End If
        Else
            MyCommon.QueryStr = "dbo.pa_FolderItem_Select_WFStatus"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@FolderID", SqlDbType.Int).Value = FolderID
            MyCommon.LRTsp.Parameters.Add("@SortText", SqlDbType.NVarChar, 50).Value = If(SortText = "AOLV.Status", "AOLV.OfferID", SortText)
            MyCommon.LRTsp.Parameters.Add("@SortDirection", SqlDbType.NVarChar, 50).Value = SortDirection
            MyCommon.LRTsp.Parameters.Add("@WFStatus", SqlDbType.Int).Value = WFStatus
            MyCommon.LRTsp.Parameters.Add("@PageIndex", SqlDbType.Int).Value = pageIndex
            MyCommon.LRTsp.Parameters.Add("@PageSize", SqlDbType.Int).Value = 1000
            MyCommon.LRTsp.Parameters.Add("@PageCount", SqlDbType.Int, 4).Direction = ParameterDirection.Output
            MyCommon.LRTsp.Parameters.Add("@TotalOffers", SqlDbType.Int, 4).Direction = ParameterDirection.Output
            If (MyCommon.IsEngineInstalled(9) AndAlso MyCommon.IsUserAssociatedWithAnyBuyer(AdminUserID) AndAlso Not (Logix.UserRoles.ViewOffersRegardlessBuyer)) Then
                MyCommon.LRTsp.Parameters.Add("@AdminUserId", SqlDbType.Int).Value = AdminUserID
            End If
        End If

        dt = MyCommon.LRTsp_select
        dt.TableName = "Offers"
        Dim ds As DataSet = dt.DataSet
        dt = New DataTable("PageCount")
        dt.Columns.Add("PageCount")
        dt.Rows.Add()
        dt.Rows(0)(0) = MyCommon.LRTsp.Parameters("@PageCount").Value
        ds.Tables.Add(dt)
        MyCommon.Close_LRTsp()
        Return ds
    End Function


    Sub SendFolderInfo(ByVal FolderID As Long)
        Dim Theme As String = ""
        Dim FolderStartDate As String = ""
        Dim FolderEndDate As String = ""
        Dim dt As DataTable
        Dim bRTConnectionOpened As Boolean = False
        Dim StatusText As String = ""
        Response.Cache.SetCacheability(HttpCacheability.NoCache)

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                MyCommon.Open_LogixRT()
                bRTConnectionOpened = True
            End If
            If MyCommon.Fetch_SystemOption(123) = "1" Then
                MyCommon.QueryStr = " Select Th.ThemeDescription,Fl.StartDate,Fl.EndDate from Folders Fl " &
                                    " left outer join FolderThemes FT on Fl.FolderID=FT.FolderID " &
                                    " left outer join Themes Th on FT.ThemeID=Th.ThemeID where Fl.FolderID=" & FolderID
                dt = MyCommon.LRT_Select
                If dt.Rows.Count > 0 Then
                    Theme = MyCommon.NZ(dt.Rows(0).Item("ThemeDescription"), "Not Selected")
                    If Not IsDBNull(dt.Rows(0).Item("StartDate")) Then
                        FolderStartDate = Format(dt.Rows(0).Item("StartDate"), "MM/dd/yyyy")
                    End If
                    If Not IsDBNull(dt.Rows(0).Item("EndDate")) Then
                        FolderEndDate = Format(dt.Rows(0).Item("EndDate"), "MM/dd/yyyy")
                    End If
                End If
                StatusText = " Theme: " & Theme & " StartDate: " & FolderStartDate & " EndDate: " & FolderEndDate
            End If
            Send(StatusText)
        Catch ex As Exception
            Send(ex.ToString)
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed AndAlso bRTConnectionOpened Then MyCommon.Close_LogixRT()
        End Try
    End Sub

    Sub SendSearchBox(ByVal State As Integer, ByVal SearchTerms As String, ByVal ResultCount As Integer)
        'Send("<div class=""box"">")
        'Send("  <h2>Search</h2>")
        'Send("  <div id=""searchbody""" & IIf(State = SearchState.COLLAPSED, " style=""display:none;""", "") & ">")
        'Send("    <input type=""text"" id=""searchterms"" name=""searchterms"" value=""" & SearchTerms & """ />")
        'Send("    <input type=""button"" id=""search"" name=""search"" onclick=""javascript:submitSearch();"" value=""" & Copient.PhraseLib.Lookup("term.search", LanguageID) & """ />")

        'Send("  </div>")
        'Send("</div>")
    End Sub

    Sub SendFoundOffers(ByVal FolderID As Long, Optional ByVal StatusText As String = "")
        Dim dt, dt1 As DataTable
        Dim row, row1 As DataRow
        Dim TempBuf As New StringBuilder()
        Dim LinkID As Integer
        Dim PromoEngine As String
        Dim Name As String = ""
        Dim FolderName As String = ""
        Dim OfferEndDate As Date
        Dim OfferStartDate As Date
        Dim Warningmessage As String = ""
        Dim idNumber As Integer = 0
        Dim idSearchText As String = ""
        Dim idSearch As String = ""
        Dim PrctSignPos As Integer = 0
        Const OFFER_LINK_TYPE As Integer = 1
        Dim IsViewEnabledAndEditDisabled = False
        Dim BuyerId As Int32
        Dim IsCheckboxDisabled = False
        Response.Cache.SetCacheability(HttpCacheability.NoCache)

        Try
            If (FolderID = 0) Then
                TempBuf.AppendLine("<p>" & Copient.PhraseLib.Lookup("folders.SelectToAdd", LanguageID) & "</p>")
            Else
                If (Request.Form("searchterms") <> "") Then
                    If (Integer.TryParse(Request.Form("searchterms"), idNumber)) Then
                        idSearch = idNumber.ToString
                    Else
                        idSearch = "-1"
                    End If
                    idSearchText = MyCommon.Parse_Quotes(Request.Form("searchterms"))
                    PrctSignPos = idSearchText.IndexOf("%")
                    If (PrctSignPos > -1) Then
                        idSearch = "-1"
                        idSearchText = idSearchText.Replace("%", "[%]")
                    End If
                    If (idSearchText.IndexOf("_") > -1) Then idSearchText = idSearchText.Replace("_", "[_]")

                    'If ViewOffersRegardlessBuyer permission is not set, get buyer Specific Offer
                    If (MyCommon.IsEngineInstalled(9) AndAlso MyCommon.IsUserAssociatedWithAnyBuyer(AdminUserID) AndAlso Not (Logix.UserRoles.ViewOffersRegardlessBuyer)) Then
                        buyerJoin = " inner join BuyerRoleusers BRU with (NoLock) on AOLV.BuyerID = BRU.BuyerID "
                        buyerwherestr = " BRU.adminUSerID = " & AdminUserID & " and "
                    End If
                    If (MyCommon.IsEngineInstalled(9) AndAlso MyCommon.IsUserAssociatedWithAnyBuyer(AdminUserID) AndAlso Logix.UserRoles.ViewOffersRegardlessBuyer AndAlso Not (Logix.UserRoles.EditOffersRegardlessBuyer)) Then
                        IsViewEnabledAndEditDisabled = True
                    End If

                    If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
                    MyCommon.QueryStr = "select AOLV.* from AllOffersListview AOLV with (NoLock) " & buyerJoin &
                                                  " where " & buyerwherestr & " AOLV.Deleted = 0 And ( AOLV.OfferID=" & idSearch & " or AOLV.Name like N'%" & idSearchText & "%') " &
                                           "   and OfferID not in (select LinkID from FolderItems with (NoLock) where FolderID=" & FolderID & " and LinkTypeID=1) "
                    If (bEnableRestrictedAccessToUEOfferBuilder) Then
                        MyCommon.QueryStr &= GetRestrictedAccessToUEBuilderQuery(MyCommon, Logix, "AOLV")
                    End If

                    MyCommon.QueryStr &= " order by Name;"
                    dt = MyCommon.LRT_Select
                    TempBuf.AppendLine("<i>" & Copient.PhraseLib.Detokenize("folders.SearchResults", LanguageID, Request.Form("searchterms"), dt.Rows.Count) & "</i><br />")  'Search for "{0}" returned {1} result(s).
                    If dt.Rows.Count > 0 Then
                        If MyCommon.Fetch_SystemOption(132) = "1" AndAlso MyCommon.Fetch_SystemOption(191) = "1" Then
                            TempBuf.AppendLine("<table summary="""" style=""width:100%;white-space: nowrap;"">")
                            TempBuf.AppendLine("  <tr>")
                            TempBuf.AppendLine("    <th style=""width:20px;""></th>")
                            TempBuf.AppendLine("    <th style=""width:100px;"">" & Copient.PhraseLib.Lookup("term.offerid", LanguageID) & "</th>")
                            TempBuf.AppendLine("    <th>" & Copient.PhraseLib.Lookup("term.name", LanguageID) & "</th>")
                            TempBuf.AppendLine("    <th>" & Copient.PhraseLib.Lookup("folders.AssociatedFolder", LanguageID) & "</th>")
                            TempBuf.AppendLine("  </tr>")
                            For Each row In dt.Rows
                                LinkID = MyCommon.NZ(row.Item("OfferID"), 0)
                                PromoEngine = MyCommon.NZ(row.Item("PromoEngine"), "")
                                Name = MyCommon.NZ(row.Item("Name"), "")
                                OfferStartDate = Format(MyCommon.NZ(row.Item("ProdStartDate"), "1/1/1900"), "MM/dd/yyyy")
                                OfferEndDate = Format(MyCommon.NZ(row.Item("ProdEndDate"), "1/1/1900"), "MM/dd/yyyy")
                                IsCheckboxDisabled = False
                                If (IsViewEnabledAndEditDisabled) Then
                                    BuyerId = MyCommon.NZ(row.Item("BuyerId"), -1)
                                    'If offer is not created with buyer or current users is not associated with Offer's Buyer, disable the offer
                                    If BuyerId = -1 Or Not BuyersAssociated.Contains(BuyerId) Then
                                        IsCheckboxDisabled = True
                                    End If
                                End If

                                MyCommon.QueryStr = "select top 1 FI.FolderID,FD.FolderName from folderitems FI with (nolock) inner join Folders FD with (nolock) on FI.FolderID=FD.FolderID where LinkID=" & LinkID & ""
                                dt1 = MyCommon.LRT_Select
                                If dt1.Rows.Count > 0 Then
                                    For Each row1 In dt1.Rows
                                        FolderName = MyCommon.NZ(row1.Item("FolderName"), "")
                                    Next
                                    TempBuf.AppendLine("  <tr>")
                                    TempBuf.AppendLine("    <td>")
                                    TempBuf.AppendLine("      <input name=""linkID"" id=""linkID" & LinkID & """ type=""checkbox"" disabled=""disabled"" value=""" & LinkID & """ onclick=""submitToAddItems(" & LinkID & "," & OFFER_LINK_TYPE & ", '" & PromoEngine & "', this.checked);"" />")
                                    TempBuf.AppendLine("    </td>")
                                    TempBuf.AppendLine("    <td>" & MyCommon.NZ(row.Item("OfferID"), "") & "</td>")
                                    TempBuf.AppendLine("    <td><a href=""offer-redirect.aspx?OfferID=" & MyCommon.NZ(row.Item("OfferID"), "") & """ target=""_blank"">" & IIf(Name <> "", Name, "(" & Copient.PhraseLib.Lookup("term.unnamed", LanguageID) & ")") & "</a></td>")
                                    TempBuf.AppendLine("    <td>" & FolderName & "</td>")
                                    TempBuf.AppendLine("  </tr>")
                                Else
                                    FolderName = ""
                                    Warningmessage = ShowWarning(FolderID, OfferStartDate, OfferEndDate)
                                    TempBuf.AppendLine("  <tr>")
                                    TempBuf.AppendLine("    <td>")
                                    TempBuf.AppendLine("      <input name=""linkID"" id=""linkID" & LinkID & """ type=""checkbox"" value=""" & LinkID & """ onclick=""submitToAddItems(" & LinkID & "," & OFFER_LINK_TYPE & ", '" & PromoEngine & "', this.checked, '" & Warningmessage & "');"" " & IIf(IsCheckboxDisabled, "disabled=""disabled""", """") & "/>")
                                    TempBuf.AppendLine("    </td>")
                                    TempBuf.AppendLine("    <td>" & MyCommon.NZ(row.Item("OfferID"), "") & "</td>")
                                    TempBuf.AppendLine("    <td><a href=""offer-redirect.aspx?OfferID=" & MyCommon.NZ(row.Item("OfferID"), "") & """ target=""_blank"">" & IIf(Name <> "", Name, "(" & Copient.PhraseLib.Lookup("term.unnamed", LanguageID) & ")") & "</a></td>")
                                    TempBuf.AppendLine("    <td>" & FolderName & "</td>")
                                    TempBuf.AppendLine("  </tr>")
                                End If
                            Next
                            TempBuf.AppendLine("</table>")
                        Else
                            TempBuf.AppendLine("<table summary="""" style=""width:100%;white-space: nowrap;"">")
                            TempBuf.AppendLine("  <tr>")
                            TempBuf.AppendLine("    <th style=""width:20px;""></th>")
                            TempBuf.AppendLine("    <th style=""width:100px;"">" & Copient.PhraseLib.Lookup("term.offerid", LanguageID) & "</th>")
                            TempBuf.AppendLine("    <th>" & Copient.PhraseLib.Lookup("term.name", LanguageID) & "</th>")
                            TempBuf.AppendLine("  </tr>")
                            For Each row In dt.Rows
                                LinkID = MyCommon.NZ(row.Item("OfferID"), 0)
                                PromoEngine = MyCommon.NZ(row.Item("PromoEngine"), "")
                                Name = MyCommon.NZ(row.Item("Name"), "")
                                IsCheckboxDisabled = False
                                If (IsViewEnabledAndEditDisabled) Then
                                    BuyerId = MyCommon.NZ(row.Item("BuyerId"), -1)
                                    'If offer is not created with buyer or current users is not associated with Offer's Buyer, disable the offer
                                    If BuyerId = -1 Or Not BuyersAssociated.Contains(BuyerId) Then
                                        IsCheckboxDisabled = True
                                    End If
                                End If
                                TempBuf.AppendLine("  <tr>")
                                TempBuf.AppendLine("    <td>")
                                TempBuf.AppendLine("      <input name=""linkID"" id=""linkID" & LinkID & """ type=""checkbox"" value=""" & LinkID & """ onclick=""submitToAddItems(" & LinkID & "," & OFFER_LINK_TYPE & ", '" & PromoEngine & "', this.checked);"" " & IIf(IsCheckboxDisabled, "disabled=""disabled""", """") & "/>")
                                TempBuf.AppendLine("    </td>")
                                TempBuf.AppendLine("    <td>" & MyCommon.NZ(row.Item("OfferID"), "") & "</td>")
                                TempBuf.AppendLine("    <td><a href=""offer-redirect.aspx?OfferID=" & MyCommon.NZ(row.Item("OfferID"), "") & """ target=""_blank"">" & IIf(Name <> "", Name, "(" & Copient.PhraseLib.Lookup("term.unnamed", LanguageID) & ")") & "</a></td>")
                                TempBuf.AppendLine("  </tr>")
                            Next
                            TempBuf.AppendLine("</table>")
                        End If
                    End If

                    If (StatusText = "") Then
                        StatusText = Copient.PhraseLib.Detokenize("folders.ItemsFoundMatching", LanguageID, dt.Rows.Count, idSearchText)
                    Else
                        StatusText &= " (" & Copient.PhraseLib.Detokenize("folders.ItemsFoundMatching", LanguageID, dt.Rows.Count, idSearchText) & ")"
                    End If
                    TempBuf.AppendLine("<input type=""hidden"" id=""offerCount"" name=""offerCount"" value=""" & dt.Rows.Count & """ />")
                    TempBuf.AppendLine("<input type=""hidden"" id=""statustext"" name=""statustext"" value=""" & StatusText & """ />")
                    'SendSearchBox(SearchState.COLLAPSED, idSearchText, dt.Rows.Count)
                Else
                    TempBuf.AppendLine("<input type=""hidden"" id=""statustext"" name=""statustext"" value=""" & Copient.PhraseLib.Lookup("term.ready", LanguageID) & "."" />")
                    TempBuf.AppendLine("<b>" & Copient.PhraseLib.Lookup("folders.EnterSearchTerms", LanguageID) & "</b>")
                    'SendSearchBox(SearchState.EXPANDED, "", 0)
                End If
            End If

            Send(TempBuf.ToString)
        Catch ex As Exception
            Send(ex.ToString)
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try
    End Sub

    Function GetOfferIDs(ByVal ItemIDs As String) As Dictionary(Of String, String)
        Dim list As Dictionary(Of String, String) = New Dictionary(Of String, String)
        Dim dt As New DataTable
        Try

            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

            MyCommon.QueryStr = " SELECT linkid AS OfferID,ItemID FROM folderitems a INNER JOIN OfferIds c ON c.OfferID = a.LinkID  where ItemID IN(SELECT items FROM Split (@ItemIDs, ','))"
            MyCommon.DBParameters.Add("@ItemIDs", SqlDbType.NVarChar).Value = ItemIDs
            dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            For Each row In dt.Rows
                list.Add(row.Item("ItemID"), row.Item("OfferID"))
            Next

        Catch ex As Exception

        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try

        Return list
    End Function

    Sub TransferOffers(ByVal SourceFolderID As String, ByVal DestFolderID As String, ByVal ItemIDs As String, ByVal FromOfferList As Boolean, Optional ByRef status As String = "")
        Dim LinkCount As Integer = 0
        Dim ItemCount As Integer = 0
        Dim OfferIDs As String = String.Empty
        Dim ConnectionOpenedLocally As Boolean = False
        Dim ExpiredOfferFilterList As New List(Of String)
        Dim ActiveOfferFilterList As New List(Of String)
        Dim OfferFilterList As New List(Of String)
        Dim NewItemIds As String()
        Dim Offeriddt As DataTable = New DataTable()
        Dim restrictOperationOnActiveExpiredOffers As Boolean = IIf(MyCommon.Fetch_SystemOption(272) = "1", True, False)
        Dim list As Dictionary(Of String, String) = New Dictionary(Of String, String)
        Dim offersDT As DataTable
        Offeriddt.Columns.Add("OfferID")
        Dim isPendingOffersExists As Boolean
        'Response.Cache.SetCacheability(HttpCacheability.NoCache)
        LogFile = "MassActionLog." & MyCommon.Leading_Zero_Fill(Year(Today), 4) & MyCommon.Leading_Zero_Fill(Month(Today), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
        MyCommon.Write_Log(LogFile, "----------------------------------------------------------------")
        Try
            If (ItemIDs IsNot Nothing AndAlso ItemIDs.Trim <> String.Empty) Then
                offersDT = GetEngineID(ItemIDs, FromOfferList)
                offersDT = RemovePendingOffers(offersDT, isPendingOffersExists)

                MyCommon.QueryStr = "dbo.pt_GetFolderItemsByOfferID"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@OffersT", SqlDbType.Structured).Value = offersDT
                Dim itemsDT As DataTable = MyCommon.LRTsp_select
                MyCommon.Close_LRTsp()

                ItemIDs = ""
                For Each row In itemsDT.Rows
                    If ItemIDs = "" Then
                        ItemIDs = row("ItemID").ToString()
                    Else
                        ItemIDs &= "," & row("ItemID").ToString()
                    End If
                Next

                ItemCount = IndexOfCount(ItemIDs, ","c) + 1
                MyCommon.Write_Log(LogFile, "Selected Action - Transfer Offers - to another folder (id: - " & DestFolderID & "). Selected OfferIDs Count: " & ItemCount, True)
                If (FromOfferList) Then
                    MyCommon.Write_Log(LogFile, "Selected Offer ID(s):" & ItemIDs, True)
                    For Each IDs In ItemIDs.Split(",")
                        If (restrictOperationOnActiveExpiredOffers) Then
                            'AMSPS-3146 :Filter out Active &Expired offers during offer Transfer 
                            Dim StatusCode As Copient.LogixInc.STATUS_FLAGS = Copient.LogixInc.STATUS_FLAGS.STATUS_UNKNOWN
                            Logix.GetOfferStatus(IDs, LanguageID, StatusCode)
                            If (StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_EXPIRED) Then
                                ExpiredOfferFilterList.Add(IDs)
                            ElseIf (StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_ACTIVE) Then
                                ActiveOfferFilterList.Add(IDs)
                            Else
                                Offeriddt.Rows.Add(IDs)
                                OfferFilterList.Add(IDs)
                            End If
                        Else
                            Offeriddt.Rows.Add(IDs)
                            OfferFilterList.Add(IDs)
                        End If
                    Next
                Else
                    'Need to retreive OfferID as the values passed to this method has folderLinkID
                    list = GetOfferIDs(ItemIDs)
                    MyCommon.Write_Log(LogFile, "Selected Offer ID(s):" & String.Join(",", list.Values), True)
                    If (list.Count > 0) Then
                        For Each IDs In ItemIDs.Split(",")
                            If (restrictOperationOnActiveExpiredOffers) Then
                                'AMSPS-3146 :Filter out Active &Expired offers during offer Transfer 
                                Dim StatusCode As Copient.LogixInc.STATUS_FLAGS = Copient.LogixInc.STATUS_FLAGS.STATUS_UNKNOWN
                                Logix.GetOfferStatus(list.Item(IDs), LanguageID, StatusCode)
                                If (StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_EXPIRED) Then
                                    ExpiredOfferFilterList.Add(list.Item(IDs))
                                ElseIf (StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_ACTIVE) Then
                                    ActiveOfferFilterList.Add(list.Item(IDs))
                                Else
                                    Offeriddt.Rows.Add(IDs)
                                    OfferFilterList.Add(list.Item(IDs))
                                End If
                            Else
                                Offeriddt.Rows.Add(IDs)
                                OfferFilterList.Add(list.Item(IDs))
                            End If
                        Next
                    End If
                End If

                If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                    MyCommon.Open_LogixRT()
                    ConnectionOpenedLocally = True
                End If
                If ItemIDs IsNot Nothing AndAlso ItemIDs.Trim <> String.Empty Then
                    If Not FromOfferList Then
                        'called from folders.aspx

                        If (restrictOperationOnActiveExpiredOffers) Then
                            If (ExpiredOfferFilterList.Count > 0) Then
                                MyCommon.Write_Log(LogFile, "Filtering out Expired offers (Offer ID(s):" & String.Join(",", ExpiredOfferFilterList) & ") as expired offers cannot be transferred when system option #272 is enabled.", True)
                            End If
                            If (ActiveOfferFilterList.Count > 0) Then
                                MyCommon.Write_Log(LogFile, "Filtering out Active offers (Offer ID(s):" & String.Join(",", ActiveOfferFilterList) & ") as active offers cannot be transferred when system option #272 is enabled.", True)
                            End If
                        End If
                        MyCommon.Write_Log(LogFile, "Transfer Initiated for - (Offer ID(s):" & String.Join(",", OfferFilterList) & ") ", True)
                        MyCommon.QueryStr = "dbo.pa_FolderOffers_Transfer"
                        MyCommon.Open_LRTsp()
                        MyCommon.LRTsp.Parameters.Add("@AdminUserID", SqlDbType.Int).Value = AdminUserID
                        MyCommon.LRTsp.Parameters.Add("@SFolderID", SqlDbType.Int).Value = SourceFolderID
                        MyCommon.LRTsp.Parameters.Add("@DFolderID", SqlDbType.Int).Value = DestFolderID
                        MyCommon.LRTsp.Parameters.Add("@Offers", SqlDbType.Structured).Value = Offeriddt
                        MyCommon.LRTsp.Parameters.Add("@LinkCount", SqlDbType.Int, 4).Direction = ParameterDirection.Output
                        MyCommon.LRTsp.ExecuteNonQuery()
                        LinkCount = MyCommon.LRTsp.Parameters("@LinkCount").Value
                        MyCommon.Close_LRTsp()
                    Else
                        'called from offer-list.aspx
                        If (restrictOperationOnActiveExpiredOffers) Then
                            If (ExpiredOfferFilterList.Count > 0) Then
                                MyCommon.Write_Log(LogFile, "Filtering out Expired offers (Offer ID(s):" & String.Join(",", ExpiredOfferFilterList) & ") as expired offers cannot be transferred when system option #272 is enabled.", True)
                            End If
                            If (ActiveOfferFilterList.Count > 0) Then
                                MyCommon.Write_Log(LogFile, "Filtering out Active offers (Offer ID(s):" & String.Join(",", ActiveOfferFilterList) & ") as active offers cannot be transferred when system option #272 is enabled .", True)
                            End If
                        End If
                        MyCommon.Write_Log(LogFile, "Transfer Initiated for - (Offer ID(s):" & String.Join(",", OfferFilterList) & ") ", True)
                        MyCommon.QueryStr = "dbo.pa_FolderOffersAdvSearch_Transfer"
                        MyCommon.Open_LRTsp()
                        MyCommon.LRTsp.Parameters.Add("@DFolderID", SqlDbType.Int).Value = DestFolderID
                        MyCommon.LRTsp.Parameters.Add("@Offers", SqlDbType.Structured).Value = Offeriddt
                        Dim dt As DataTable = MyCommon.LRTsp_select
                        OfferIDs = MyCommon.NZ(dt.Rows(0).Item("OfferID"), "0")
                        MyCommon.Close_LRTsp()
                    End If
                End If

                If Not FromOfferList Then
                    If ItemCount = (ExpiredOfferFilterList.Count + ActiveOfferFilterList.Count) Then
                        MyCommon.Write_Log(LogFile, "Transfer completed.No Offers found to transfer after filtering out the selected list.", True)
                        status = Copient.PhraseLib.Lookup("term.warning", LanguageID) + ": " + Copient.PhraseLib.Lookup("bulkoffertransfer-warningall", LanguageID)
                    Else
                        MyCommon.Write_Log(LogFile, "Transferred OfferIDs: " & String.Join(",", OfferFilterList), True)
                        If (ExpiredOfferFilterList.Count + ActiveOfferFilterList.Count) > 0 Then
                            status = Copient.PhraseLib.Lookup("term.warning", LanguageID) + ": " + Copient.PhraseLib.Lookup("bulkoffertransfer-warning", LanguageID)
                        ElseIf ItemCount = LinkCount Then
                            status = Copient.PhraseLib.Lookup("term.success", LanguageID) + ": " + Copient.PhraseLib.Lookup("transfer-success", LanguageID)
                            If isPendingOffersExists Then
                                status &= " " + Copient.PhraseLib.Lookup("term.ignore-pendingoffers", LanguageID)
                            End If
                        Else
                            status = Copient.PhraseLib.Lookup("transfer-failure", LanguageID)
                        End If
                    End If
                Else
                    If ItemCount = (ExpiredOfferFilterList.Count + ActiveOfferFilterList.Count) Then
                        MyCommon.Write_Log(LogFile, "Transfer completed.No Offers found to transfer after filtering out the selected list.", True)
                        status = Copient.PhraseLib.Lookup("term.warning", LanguageID) + ":" + Copient.PhraseLib.Lookup("bulkoffertransfer-warningall", LanguageID)
                    Else
                        If OfferIDs = "0" Then
                            MyCommon.Write_Log(LogFile, "Transferred OfferIDs: " & String.Join(",", OfferFilterList), True)
                            If (ExpiredOfferFilterList.Count + ActiveOfferFilterList.Count) > 0 Then
                                status = Copient.PhraseLib.Lookup("term.warning", LanguageID) + ": " + Copient.PhraseLib.Lookup("bulkoffertransfer-warning", LanguageID)
                            ElseIf ItemCount = OfferFilterList.Count Then
                                status = Copient.PhraseLib.Lookup("term.success", LanguageID) + ": " + Copient.PhraseLib.Lookup("transfer-success", LanguageID)
                                If isPendingOffersExists Then
                                    status &= " " + Copient.PhraseLib.Lookup("term.ignore-pendingoffers", LanguageID)
                                End If
                            End If
                        Else
                            MyCommon.Write_Log(LogFile, "Transfer failed. " & OfferIDs & " Offer(s) exists in multiple source folders.", True)
                            status = OfferIDs + Copient.PhraseLib.Lookup("transfer-failure", LanguageID)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MyCommon.Write_Log(LogFile, ex.ToString(), True)
            'Send(ex.ToString)
            status = "Error:" + ex.Message
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed AndAlso ConnectionOpenedLocally Then MyCommon.Close_LogixRT()
        End Try
    End Sub

    ' ********************************************************************************
    ' *** Removes all the specified items already assigned to the specified FolderID
    ' *** Should only be called for use in the Remove mode.
    ' ********************************************************************************
    Sub RemoveItemsFromFolder(ByVal FolderID As Long, ByVal ItemIDs As String, ByVal AdminUserID As Integer)
        Dim ResponseText As String = ""
        Dim LinkCount As Integer = 0
        Dim dt As DataTable
        Dim ItemCount As Integer

        Response.Cache.SetCacheability(HttpCacheability.NoCache)

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

            If ItemIDs IsNot Nothing AndAlso ItemIDs.Trim <> String.Empty Then
                LinkCount = IndexOfCount(ItemIDs, ","c) + 1
                MyCommon.QueryStr = "dbo.pa_FolderItem_Delete"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@FolderID", SqlDbType.Int).Value = FolderID
                MyCommon.LRTsp.Parameters.Add("@ItemIDs", SqlDbType.NVarChar).Value = ItemIDs
                MyCommon.LRTsp.ExecuteNonQuery()
                ResponseText = LinkCount & " item" & IIf(LinkCount = 1, "", "s") & " removed from selected folder."
                MyCommon.Close_LRTsp()
            End If

            ' determine if there are any items still associated with the folder after the remove
            MyCommon.QueryStr = "select count(ItemID) as ItemCount from FolderItems with (NoLock) where FolderID=@FolderID"
            MyCommon.DBParameters.Add("@FolderID", SqlDbType.Int).Value = FolderID
            dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            If dt.Rows.Count > 0 Then
                ItemCount = MyCommon.NZ(dt.Rows(0).Item("ItemCount"), 0)
            End If
            Send(ItemCount & "|")
            SetFolderMassOperationStatus(FolderID, FOLDER_NOT_IN_USE)
            MyCommon.Activity_Log2(45, 23, FolderID, AdminUserID, Copient.PhraseLib.Lookup("history.folder-removeitem", LanguageID))
            SendFolderItems(FolderID, ResponseText)
        Catch ex As Exception
            Send(ex.ToString)
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try
    End Sub

    Sub SetLastDeployValidationMessage(ByVal OfferIds As String, ByVal OfferDescriptions As String, ByVal OffersCMEngineFlag As String)
        Dim offerIdArray() As String = OfferIds.Split(",").ToArray()
        Dim offerDesc() As String = OfferDescriptions.Split("|").ToArray()

        Dim i As Int32
        Dim msg As String
        If offerIdArray.Length > 1 Then
            For i = 0 To offerIdArray.Length - 2
                If String.IsNullOrEmpty(OfferDescriptions) Then
                    msg = Copient.PhraseLib.Lookup("term.validationsuccessful", LanguageID)
                Else
                    msg = "<font color=""red"">" & offerDesc(i) & "</font>"
                End If
                If OffersCMEngineFlag(i).Equals(Boolean.TrueString) Then
                    MyCommon.QueryStr = "Update Offers " &
                                      "  Set LastDeployValidationMessage=@Message " &
                                      "  where OfferId=@OfferId"

                Else
                    MyCommon.QueryStr = "Update CPE_Incentives " &
                                      "  Set LastDeployValidationMessage=@Message " &
                                      "  where IncentiveId=@OfferId"
                End If
                MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = Convert.ToInt64(offerIdArray(i))
                MyCommon.DBParameters.Add("@Message", SqlDbType.NVarChar).Value = msg
                MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
            Next
        End If
    End Sub

    Sub DeleteSelectedOffersAsync(ByVal State As Object)
        Dim ItemIDs As String = State(0)
        Dim AdminUserID As Integer = State(1)
        Dim FolderID As Integer = State(2)
        Dim Status As String = ""
        Dim err As String = String.Empty
        Dim isPendingOffersExists As Boolean
        Try
            SetFolderMassOperationStatus(FolderID, FOLDER_IN_USE + " Mass Action is going on to delete Offers.")
            RemoveSelectedOffers(ItemIDs, AdminUserID, FolderID, Status, isPendingOffersExists)
            If Status.IndexOf(SUCCESS) >= 0 Then
                Status = Copient.PhraseLib.Lookup("term.success1", LanguageID) + " " + Copient.PhraseLib.Lookup("delete-success", LanguageID)
                If isPendingOffersExists Then
                    Status &= " " + Copient.PhraseLib.Lookup("term.ignore-pendingoffers", LanguageID)
                End If
            End If
        Catch ex As Exception
            err = "Error: " + ex.ToString
            MyCommon.Write_Log(LogFile, err)
        Finally
            SetFolderMassOperationStatus(FolderID, FOLDER_NOT_IN_USE + Status)
        End Try
    End Sub

    Sub RemoveSelectedOffers(ByVal ItemIDs As String, ByVal AdminUserID As Integer, ByVal FolderID As Integer, ByRef Status As String, ByRef isPendingOffersExists As Boolean)
        Dim aItemIDs As String() = Nothing
        Dim aPromoEngine As String() = Nothing
        Dim bBeginTransactionRT As Boolean = False
        Dim iOfferID As Long
        Dim dt, dt1, rst As DataTable
        Dim roid As Integer = 0
        Dim EngineID As Integer
        Dim ISCMOffer As Boolean = False
        Dim NotAppliedOfferLst As String = ""
        Dim isValid As Boolean
        Dim offersDT As DataTable
        Try

            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()


            LogFile = "MassActionLog." & MyCommon.Leading_Zero_Fill(Year(Today), 4) & MyCommon.Leading_Zero_Fill(Month(Today), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
            MyCommon.Write_Log(LogFile, "Mass Deletion Of Offers started....", True)

            'CurrentRequest.Resolver.AppName = "folder-feeds.aspx"
            'Dim m_Offer As IOffer = CurrentRequest.Resolver.Resolve(Of IOffer)()
            'Dim m_customerGroup As ICustomerGroups = CurrentRequest.Resolver.Resolve(Of ICustomerGroups)()
            offersDT = GetEngineID(ItemIDs, False)
            offersDT = RemovePendingOffers(offersDT, isPendingOffersExists)

            MyCommon.QueryStr = "dbo.pt_GetFolderItemsByOfferID"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@OffersT", SqlDbType.Structured).Value = offersDT
            Dim itemsDT As DataTable = MyCommon.LRTsp_select
            MyCommon.Close_LRTsp()
            Dim itemsList As List(Of String) = New List(Of String)()
            For Each row In itemsDT.Rows
                itemsList.Add(row("ItemID").ToString())
            Next
            aItemIDs = itemsList.ToArray


            MyCommon.QueryStr = "Begin Transaction;"
            MyCommon.LRT_Execute()
            bBeginTransactionRT = True

            For i = 0 To aItemIDs.GetUpperBound(0)
                isValid = True
                MyCommon.QueryStr = "select LinkID from FolderItems with (NoLock) where ItemID =@aItemID"
                'MyCommon.LRT_Execute()
                MyCommon.DBParameters.Add("@aItemID", SqlDbType.Int).Value = Convert.ToInt64(aItemIDs(i))
                dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                
                iOfferID = Convert.ToInt64(dt.Rows(0)("LinkID"))
                MyCommon.QueryStr = "select RewardOptionID, TierLevels from CPE_RewardOptions with (NoLock) where TouchResponse=0 and IncentiveID=" & iOfferID & ";"
                dt1 = MyCommon.LRT_Select
                If dt1.Rows.Count > 0 Then
                    roid = MyCommon.NZ(dt1.Rows(0).Item("RewardOptionID"), 0)
                End If
                MyCommon.QueryStr = "select EngineID from OfferIds with (NoLock) where OfferID=" & iOfferID & ";"
                rst = MyCommon.LRT_Select
                If rst.Rows.Count > 0 Then
                    If MyCommon.Extract_Val(MyCommon.NZ(rst.Rows(0).Item("EngineID"), 0)) = 0 Then
                        ISCMOffer = True
                    End If
                End If

                'check if the offer is active
                If MyCommon.Fetch_SystemOption(285) = "1" Then
                    Dim StatusCode As Copient.LogixInc.STATUS_FLAGS = Copient.LogixInc.STATUS_FLAGS.STATUS_UNKNOWN
                    Logix.GetOfferStatus(iOfferID, LanguageID, StatusCode)
                    If (StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_ACTIVE) Then
                        NotAppliedOfferLst &= iOfferID & ","

                        MyCommon.Write_Log(LogFile, "Following offers are not deleted as those are active:" & Environment.NewLine & "" & NotAppliedOfferLst & "", True)
                        isValid = False
                    End If
                End If
                If isValid Then
                    If ISCMOffer Then

                        'Dim optInGroup As CustomerGroup = m_Offer.GetOfferDefaultCustomerGroup(iOfferID, 0)
                        MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=2,Deleted=1,UpdateLevel=UpdateLevel+1 where OfferID=" & iOfferID
                        MyCommon.LRT_Execute()

                        'Mark Client ID deleted if this is external offer.
                        MyCommon.QueryStr = "dbo.pt_ExtOfferID_Delete"
                        MyCommon.Open_LRTsp()
                        MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.NVarChar, 20).Value = iOfferID
                        MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = 0
                        MyCommon.LRTsp.ExecuteNonQuery()
                        MyCommon.Close_LRTsp()
                        'Remove any Offer Eligibility Conditions associated with the Offer.    
                        'm_Offer.DeleteOfferEligibleConditions(iOfferID, 0)
                        'If (optInGroup IsNot Nothing) Then
                        '    m_customerGroup.DeleteCustomerGroup(optInGroup.CustomerGroupID)
                        'End If

                        'remove the banners assigned to this offer
                        If (MyCommon.Fetch_SystemOption(66) = "1") Then
                            MyCommon.QueryStr = "delete from BannerOffers with (RowLock) where OfferID = " & iOfferID
                            MyCommon.LRT_Execute()
                        End If

                        MyCommon.Activity_Log(3, iOfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-delete", LanguageID))
                    Else

                        MyCommon.QueryStr = "select EngineID from CPE_Incentives with (NoLock) where IncentiveID=" & iOfferID & ";"
                        dt1 = MyCommon.LRT_Select
                        If dt1.Rows.Count > 0 Then
                            EngineID = MyCommon.NZ(dt1.Rows(0).Item("EngineID"), 0)
                        End If
                        'Dim optInGroup As CustomerGroup = m_Offer.GetOfferDefaultCustomerGroup(iOfferID, EngineID)
                        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=2, Deleted=1, LastUpdate=getdate(), UpdateLevel=UpdateLevel+1 where IncentiveID=" & iOfferID
                        MyCommon.LRT_Execute()
                        'Mark the shadow table offer as deleted as well.
                        MyCommon.QueryStr = "update CPE_ST_Incentives with (RowLock) set Deleted=1, LastUpdate=getdate(), UpdateLevel = " &
                        " (select UpdateLevel from CPE_Incentives with (NoLock)where IncentiveID=" & iOfferID & ") where IncentiveID=" & iOfferID
                        MyCommon.LRT_Execute()
                        'Mark Client ID deleted if this is enternal offer.
                        MyCommon.QueryStr = "dbo.pt_ExtOfferID_Delete"
                        MyCommon.Open_LRTsp()
                        MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.NVarChar, 20).Value = iOfferID
                        MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = EngineID
                        MyCommon.LRTsp.ExecuteNonQuery()
                        MyCommon.Close_LRTsp()
                        'Remove any Offer Eligibility Conditions associated with the Offer.
                        'm_Offer.DeleteOfferEligibleConditions(iOfferID, EngineID)
                        'If (optInGroup IsNot Nothing) Then
                        '    m_customerGroup.DeleteCustomerGroup(optInGroup.CustomerGroupID)
                        'End If

                        'Remove the banners assigned to this offer
                        If (MyCommon.Fetch_SystemOption(66) = "1") Then
                            MyCommon.QueryStr = "delete from BannerOffers with (RowLock) where OfferID = " & iOfferID
                            MyCommon.LRT_Execute()
                        End If
                        'Also remove/update the triggers for any associated EIW conditions
                        MyCommon.QueryStr = "update CPE_EIWTriggers with (RowLock) set Removed=1, LastUpdate=getdate() where RewardOptionID=" & roid & " and Removed=0;"
                        MyCommon.LRT_Execute()
                        'Record activity 
                        MyCommon.Activity_Log(3, iOfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-delete", LanguageID))
                    End If
                End If
            Next

            If bBeginTransactionRT Then
                MyCommon.QueryStr = "Commit Transaction;"
                MyCommon.LRT_Execute()
                If NotAppliedOfferLst <> "" Then
                    NotAppliedOfferLst = NotAppliedOfferLst.TrimEnd(CChar(","))
                    Status = "Some of the Offer(s) [OfferID: " & NotAppliedOfferLst & "] are not deleted because those offer(s) are Active"
                Else
                    Status = Copient.PhraseLib.Lookup("term.success1", LanguageID) + " " + Copient.PhraseLib.Lookup("delete-success", LanguageID)
                End If
                MyCommon.Write_Log(LogFile, Status, True)
            End If
        Catch ex As Exception
            MyCommon.QueryStr = "Rollback Transaction;"
            MyCommon.LRT_Execute()
            Status = ex.ToString()
            MyCommon.Write_Log(LogFile, Status, True)
        End Try

    End Sub


    Sub ApplyFolderStartEndDatesToOfferAsync(ByVal State As Object)
        Dim ItemIDs As String = State(0)
        Dim AdminUserID As Integer = State(1)
        Dim FolderID As Integer = State(2)
        Dim ChangeFolderStartDates As Boolean = State(3)
        Dim ChangeFolderEndDates As Boolean = State(4)
        Dim Status As String = ""
        Dim isValid As Boolean = True
        Dim rows() As DataRow
        Dim dtErrorOffers As New DataTable
        Dim result As String = ""
        Dim isPendingOffersExists As Boolean
        Dim bAllowTimeWithStartEndDates As Boolean = False
        bAllowTimeWithStartEndDates = MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE) AndAlso (MyCommon.Fetch_UE_SystemOption(200) = "1")

        Try

            If ChangeFolderStartDates AndAlso ChangeFolderEndDates Then
                If bAllowTimeWithStartEndDates Then
                    Status = " Mass Action is going on to apply Folder Start and End Date and Time to the selected Offers By " + AdminName
                Else
                    Status = " Mass Action is going on to apply Folder Start and End Dates to the selected Offers By " + AdminName
                End If
            ElseIf ChangeFolderStartDates Then
                If bAllowTimeWithStartEndDates Then
                    Status = " Mass Action is going on to apply Folder Start Date and Time to the selected Offers By " + AdminName
                Else
                    Status = " Mass Action is going on to apply Folder Start Dates to the selected Offers By " + AdminName
                End If
            ElseIf ChangeFolderEndDates Then
                If bAllowTimeWithStartEndDates Then
                    Status = " Mass Action is going on to apply Folder End Date and Time to the selected Offers By " + AdminName
                Else
                    Status = " Mass Action is going on to apply Folder End Dates to the selected Offers By " + AdminName
                End If
            End If

            SetFolderMassOperationStatus(FolderID, FOLDER_IN_USE + Status)

            Dim dtOfferIDs As New DataTable
            Dim dtUEOfferIDs As New DataTable
            dtUEOfferIDs.Columns.Add("OfferID")

            dtOfferIDs = GetEngineID(ItemIDs, False)
            dtOfferIDs = RemovePendingOffers(dtOfferIDs, isPendingOffersExists)
            rows = dtOfferIDs.Select("EngineID = " & Engines.UE)

            dtOfferIDs.Columns.Remove("EngineID")

            For Each row In rows
                dtUEOfferIDs.Rows.Add(row(0))
            Next

            ValidateInfo(dtOfferIDs, dtUEOfferIDs, AdminUserID, bAllowTimeWithStartEndDates, isValid)

            If isValid Then
                If MyCommon.Fetch_SystemOption(284) = "1" Then
                    MyCommon.QueryStr = "dbo.pa_Validate_OfferFolderdates"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.AddWithValue("@OffersT", dtOfferIDs)
                    MyCommon.LRTsp.Parameters.AddWithValue("@FolderID", FolderID)
                    MyCommon.LRTsp.Parameters.AddWithValue("@IsMassUpdateEnabled", False)
                    dtErrorOffers = MyCommon.LRTsp_select
                    If dtErrorOffers.Rows.Count > 0 AndAlso MyCommon.NZ(dtErrorOffers.Rows(0)(0), -1) > 0 Then
                        result = [String].Join(Environment.NewLine, dtErrorOffers.AsEnumerable().[Select](Function(row) row.Field(Of Int64)(0)))
                        MyCommon.Write_Log(LogFile, "Offer dates cannot be changed for following user modified offers:" & Environment.NewLine & "" & result & "", True)
                    End If

                    Dim ValidOfferRows = dtOfferIDs.AsEnumerable.Except(dtErrorOffers.AsEnumerable, DataRowComparer.[Default])
                    If ValidOfferRows.Any Then
                        dtOfferIDs = ValidOfferRows.CopyToDataTable()
                        dtOfferIDs.AcceptChanges()
                    Else
                        dtOfferIDs.Clear()
                    End If
                End If
                ApplyFolderStartEndDatesToOffer(dtOfferIDs, AdminUserID, FolderID, Status, ChangeFolderStartDates, ChangeFolderEndDates, result, dtErrorOffers, isPendingOffersExists, ItemIDs)

            Else
                If bAllowTimeWithStartEndDates Then
                    Status = Copient.PhraseLib.Lookup("mass-datetimechange-failed", LanguageID)
                Else
                    Status = Copient.PhraseLib.Lookup("mass-datechange-failed", LanguageID)
                End If
            End If

        Catch ex As Exception
            Status = ex.Message
        Finally
            SetFolderMassOperationStatus(FolderID, FOLDER_NOT_IN_USE + Status)
        End Try
    End Sub


    Sub ApplyFolderStartEndDatesToOffer(ByVal OfferDt As DataTable, ByVal AdminUserID As Integer, ByVal FolderID As Integer, ByRef Status As String, ByVal ChangeFolderStartDate As Boolean, ByVal ChangeFolderEndDate As Boolean, ByVal UserModifiedOffers As String, ByVal dtErrorOffers As DataTable, ByVal isPendingOffersExists As Boolean, Optional ByRef ItemIDs As String = "")
        Dim aItemIDs As String() = Nothing
        Dim bBeginTransactionRT As Boolean = False
        Dim FolderEndDate As Date
        Dim FolderStartDate As Date
        Dim dt As DataTable
        Dim SetClause As String = String.Empty
        Dim isAllowed As Boolean = True
        Dim Operation As Integer
        Dim RecCount As Integer
        Dim RestrictExpiredOffers As String
        'CLOUDSOL:2163 
        Dim RestrictFolderOperationOnActiveOrExpiredOffers As Boolean = IIf(MyCommon.Fetch_SystemOption(272) = "1", True, False)
        Dim NotAppliedOfferLst As String = String.Empty
        Dim ActiveOfferLst As String = String.Empty
        Dim ActiveOfferDt As New DataTable
        Dim options As String = ""

        'CLOUDSOL-3497
        Dim IsOfferEndDateExpired As Boolean = False
        Dim LockOffersAfterExpiration As Boolean = IIf(MyCommon.Fetch_UE_SystemOption(80) = "1", True, False)
        Dim OffersLockedDeployment As New List(Of Int64)
        Dim ValidOfferStr As String = ""
        Dim bAllowTimeWithStartEndDates As Boolean = False
        bAllowTimeWithStartEndDates = MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE) AndAlso (MyCommon.Fetch_UE_SystemOption(200) = "1")

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

            LogFile = "MassActionLog." & MyCommon.Leading_Zero_Fill(Year(Today), 4) & MyCommon.Leading_Zero_Fill(Month(Today), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
            If bAllowTimeWithStartEndDates Then
                MyCommon.Write_Log(LogFile, "Mass Update of Offers Date and Time started....", True)
            Else
                MyCommon.Write_Log(LogFile, "Mass Update of Offers Date started....", True)
            End If

            MyCommon.QueryStr = "select StartDate,EndDate from Folders with (NoLock) where FolderID=" & FolderID
            dt = MyCommon.LRT_Select
            If dt.Rows.Count > 0 Then

                If ChangeFolderEndDate AndAlso ChangeFolderStartDate Then
                    Operation = UPDATE_START_AND_END_DATE
                Else
                    If ChangeFolderStartDate Then
                        Operation = UPDATE_START_DATE
                    Else
                        Operation = UPDATE_END_DATE
                    End If
                End If
                If Not IsDBNull(dt.Rows(0).Item("StartDate")) Then
                    If bAllowTimeWithStartEndDates Then
                        FolderStartDate = Format(dt.Rows(0).Item("StartDate"), "MM/dd/yyyy HH:mm:ss")
                    Else
                        FolderStartDate = Format(dt.Rows(0).Item("StartDate"), "MM/dd/yyyy")
                    End If
                Else
                    If bAllowTimeWithStartEndDates Then
                        Status = Copient.PhraseLib.Lookup("folder.nofolderstartdatetime", LanguageID)
                    Else
                        Status = Copient.PhraseLib.Lookup("folder.nofolderstartdate", LanguageID)
                    End If
                    MyCommon.Write_Log(LogFile, Status, True)
                    Exit Sub
                End If

                If Not IsDBNull(dt.Rows(0).Item("EndDate")) Then
                    If bAllowTimeWithStartEndDates Then
                        FolderEndDate = Format(dt.Rows(0).Item("EndDate"), "MM/dd/yyyy HH:mm:ss")
                    Else
                        FolderEndDate = Format(dt.Rows(0).Item("EndDate"), "MM/dd/yyyy")
                    End If
                    If FolderEndDate < Date.Today() Then
                        IsOfferEndDateExpired = True
                    End If
                Else
                    If bAllowTimeWithStartEndDates Then
                        Status = Copient.PhraseLib.Lookup("folder.nofolderenddatetime", LanguageID)
                    Else
                        Status = Copient.PhraseLib.Lookup("folder.nofolderenddate", LanguageID)
                    End If
                    MyCommon.Write_Log(LogFile, Status, True)
                    Exit Sub
                End If
                If (RestrictFolderOperationOnActiveOrExpiredOffers AndAlso OfferDt.Rows.Count > 0) Then

                    For Each dr As DataRow In OfferDt.Rows
                        Dim StatusCode As Copient.LogixInc.STATUS_FLAGS = Copient.LogixInc.STATUS_FLAGS.STATUS_UNKNOWN
                        Logix.GetOfferStatus(MyCommon.NZ(dr.Item("OfferID"), "0"), LanguageID, StatusCode)
                        If (Operation = UPDATE_START_DATE AndAlso (StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_ACTIVE OrElse StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_EXPIRED)) Then
                            NotAppliedOfferLst &= dr.Item("OfferID") & ","
                        ElseIf (Operation = UPDATE_START_AND_END_DATE) Then
                            If (StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_ACTIVE) Then
                                ActiveOfferLst &= dr.Item("OfferID") & ","
                            ElseIf (StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_EXPIRED) Then
                                NotAppliedOfferLst &= dr.Item("OfferID") & ","
                            End If
                        ElseIf (Operation = UPDATE_END_DATE AndAlso StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_EXPIRED) Then
                            NotAppliedOfferLst &= dr.Item("OfferID") & ","
                        End If
                    Next

                    Dim OfferFilter As String = ""
                    If (Operation = UPDATE_START_AND_END_DATE) Then
                        'Appending all Active and expired offers to not NotAppliedOfferLst because offer with status other than Active/Expires status are processed first
                        'When Operation is : Perform Action -> Apply folder start/end dates to Offer - Active offers should be applied with only End dates 
                        'i.e Do not apply Folder Start Dates to the offer.Apply Folder End Dates to the offer.
                        OfferFilter = IIf(NotAppliedOfferLst.Length > 0, NotAppliedOfferLst, "") & IIf(ActiveOfferLst.Length > 0, ActiveOfferLst.Trim(","), "")
                        If (ActiveOfferLst.Length > 0) Then
                            Dim dtrMatchResult1 As DataRow() = OfferDt.Select("OfferID in (" & ActiveOfferLst.Trim(",") & ")")
                            ActiveOfferDt = OfferDt.Clone()
                            For Each dr As DataRow In dtrMatchResult1
                                ActiveOfferDt.ImportRow(dr)
                            Next
                        End If
                    Else
                        OfferFilter = NotAppliedOfferLst.Trim(",")
                    End If

                    If dt.Rows.Count > 0 AndAlso OfferDt.Rows.Count > 0 AndAlso (Operation = UPDATE_START_AND_END_DATE Or Operation = UPDATE_END_DATE) AndAlso LockOffersAfterExpiration AndAlso IsOfferEndDateExpired Then
                        'CLOUDSOL-3497 - If Offer got expired and UE_SystemOption 80 value is 1 then Offer needs to be deployed
                        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                            MyCommon.Open_LogixRT()
                        End If
                        MyCommon.QueryStr = "dbo.pa_FetchOffersDeployStatus"
                        MyCommon.Open_LRTsp()
                        MyCommon.LRTsp.Parameters.Add("@OfferData", SqlDbType.Structured).Value = OfferDt
                        Dim UpdateLevelCountDT As DataTable = MyCommon.LRTsp_select

                        Dim requestResolver = New ResolverBuilder()
                        requestResolver.Build()
                        RegisterDependencies(requestResolver)
                        Dim offerDeploymentValidator As IOfferDeploymentValidator = requestResolver.Container.Resolve(Of IOfferDeploymentValidator)()
                        Dim ValidationResult As AMSResult(Of DataTable) = New AMSResult(Of DataTable)()
                        ValidationResult = offerDeploymentValidator.ValidateOffers(UpdateLevelCountDT, False, True, True, True, False, AdminUserId:=AdminUserID, LangID:=LanguageID)
                        ValidOfferStr = ValidationResult.Result.Rows(1).Item("ReturnMessage").ToString().TrimEnd(",")
                    End If

                    If (Not String.IsNullOrEmpty(OfferFilter)) Then

                        Dim dtrMatchResult As DataRow() = OfferDt.Select("OfferID in (" & OfferFilter & ")")
                        For Each dr As DataRow In dtrMatchResult
                            OfferDt.Rows.Remove(dr)
                        Next
                        OfferDt.AcceptChanges()
                    End If
                End If

                If (RestrictFolderOperationOnActiveOrExpiredOffers AndAlso dtErrorOffers.Rows.Count > 0) Then

                    For Each dr As DataRow In dtErrorOffers.Rows
                        Dim StatusCode As Copient.LogixInc.STATUS_FLAGS = Copient.LogixInc.STATUS_FLAGS.STATUS_UNKNOWN
                        Logix.GetOfferStatus(MyCommon.NZ(dr.Item(0), "0"), LanguageID, StatusCode)
                        If (Operation = UPDATE_START_DATE AndAlso StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_ACTIVE) Then
                            NotAppliedOfferLst &= dr.Item(0) & ","
                        ElseIf (Operation = UPDATE_START_AND_END_DATE) Then
                            If (StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_ACTIVE) Then
                                ActiveOfferLst &= dr.Item(0) & ","
                            End If
                        End If

                    Next

                End If


                MyCommon.QueryStr = "Begin Transaction;"
                MyCommon.LRT_Execute()
                bBeginTransactionRT = True

                options = MyCommon.Fetch_SystemOption(226)
                RestrictExpiredOffers = MyCommon.Fetch_SystemOption(272)

                If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                    MyCommon.Open_LogixRT()
                End If

                MyCommon.QueryStr = "dbo.pa_OffersDate_Update"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("AdminUserID", SqlDbType.Int).Value = AdminUserID
                MyCommon.LRTsp.Parameters.Add("@FStartdate", SqlDbType.DateTime).Value = FolderStartDate
                MyCommon.LRTsp.Parameters.Add("@FEnddate", SqlDbType.DateTime).Value = FolderEndDate

                MyCommon.LRTsp.Parameters.Add("@OffersT", SqlDbType.Structured).Value = OfferDt
                MyCommon.LRTsp.Parameters.Add("@SoptVal", SqlDbType.Int).Value = options
                MyCommon.LRTsp.Parameters.Add("@Operation", SqlDbType.Int).Value = Operation
                MyCommon.LRTsp.Parameters.Add("@allowOverride", SqlDbType.Int).Value = RestrictExpiredOffers

                dt = MyCommon.LRTsp_select
                If (IsOfferApprovalWorkflowEnabled(OfferDt)) Then
                    ResetOfferApprovalStatus_MultipleOffers(MyCommon, OfferDt)
                End If
                If (RestrictFolderOperationOnActiveOrExpiredOffers AndAlso Operation = UPDATE_START_AND_END_DATE AndAlso ActiveOfferDt.Rows.Count > 0) Then
                    'When Operation is : Perform Action -> Apply folder start/end dates to Offer - Active offers should be applied with only End dates 
                    'i.e Do not apply Folder Start Dates to the offer.Apply Folder End Dates to the offer.
                    'Passing all active offers under operation: UPDATE_END_DATE
                    Dim dt1 As New DataTable
                    MyCommon.QueryStr = "dbo.pa_OffersDate_Update"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("AdminUserID", SqlDbType.Int).Value = AdminUserID
                    MyCommon.LRTsp.Parameters.Add("@FStartdate", SqlDbType.DateTime).Value = FolderStartDate
                    MyCommon.LRTsp.Parameters.Add("@FEnddate", SqlDbType.DateTime).Value = FolderEndDate

                    MyCommon.LRTsp.Parameters.Add("@OffersT", SqlDbType.Structured).Value = ActiveOfferDt
                    MyCommon.LRTsp.Parameters.Add("@SoptVal", SqlDbType.Int).Value = options
                    MyCommon.LRTsp.Parameters.Add("@Operation", SqlDbType.Int).Value = UPDATE_END_DATE
                    MyCommon.LRTsp.Parameters.Add("@allowOverride", SqlDbType.Int).Value = RestrictExpiredOffers
                    Dim recount As Integer = 0
                    dt1 = MyCommon.LRTsp_select
                    If (IsOfferApprovalWorkflowEnabled(ActiveOfferDt)) Then
                        ResetOfferApprovalStatus_MultipleOffers(MyCommon, ActiveOfferDt)
                    End If
                    If (dt1.Rows.Count > 0) Then
                        'Processing Active offer list seperately becoz only end date should be updated for active offer in case of Applying folder start/end date to offer
                        ' but finally the processing details of Active and other offers should be consolidated
                        recount = Convert.ToInt32(dt1.Rows(0)("Rcount")) + IIf(dt.Rows.Count > 0, Convert.ToInt32(dt.Rows(0)("Rcount")), 0)
                        If (dt.Rows.Count > 0) Then
                            dt.Rows.Clear()
                        Else
                            dt = dt1.Clone()
                        End If

                        Dim row As DataRow
                        row = dt.NewRow()
                        row("Rcount") = recount
                        dt.Rows.Add(row)

                        OfferDt.Merge(ActiveOfferDt)
                    End If
                End If
            End If

            If (dt.Rows.Count > 0) Then
                RecCount = Convert.ToInt32(dt.Rows(0)("Rcount"))
                If RecCount <> OfferDt.Rows.Count Then
                    If bAllowTimeWithStartEndDates Then
                        Status = "Some of the Offer(s) date and time doesn't fall in folder date time range"
                    Else
                        Status = "Some of the Offer(s) dates doesn't fall in folder dates range"
                    End If
                    MyCommon.Write_Log(LogFile, Status, True)
                End If
            End If

            If bBeginTransactionRT Then
                MyCommon.QueryStr = "Commit Transaction;"
                MyCommon.LRT_Execute()
                If RecCount <> OfferDt.Rows.Count AndAlso (RestrictExpiredOffers = "1" OrElse options = "0") Then
                    If bAllowTimeWithStartEndDates Then
                        Status = "The Start/End date and time of valid offers only have been changed successfully due to restrictions."
                    Else
                        Status = "The Start/End dates of valid offers only have been changed successfully due to restrictions."
                    End If
                ElseIf NotAppliedOfferLst.Length = 0 AndAlso ActiveOfferLst.Length = 0 AndAlso UserModifiedOffers = "" Then
                    If bAllowTimeWithStartEndDates Then
                        Status = Copient.PhraseLib.Lookup("term.success1", LanguageID) + " " + Copient.PhraseLib.Lookup("info.datetimechanged", LanguageID)
                    Else
                        Status = Copient.PhraseLib.Lookup("term.success1", LanguageID) + " " + Copient.PhraseLib.Lookup("info.dateschanged", LanguageID)
                    End If
                    If isPendingOffersExists Then
                        Status &= " " & Copient.PhraseLib.Lookup("term.ignore-pendingoffers", LanguageID)
                    End If
                Else
                    If bAllowTimeWithStartEndDates Then
                        Status = Copient.PhraseLib.Lookup("error.datetimenotchanged", LanguageID)
                    Else
                        Status = Copient.PhraseLib.Lookup("error.datesnotchanged", LanguageID)
                    End If
                End If
                MyCommon.Write_Log(LogFile, Status, True)

                If (RestrictFolderOperationOnActiveOrExpiredOffers) Then
                    If (Operation = UPDATE_START_AND_END_DATE) Then
                        If (NotAppliedOfferLst.Length > 0) Then
                            Status = Status & Copient.PhraseLib.Lookup("term.someoffers", LanguageID) & " [OfferID: " & NotAppliedOfferLst.Replace(",", ", ").Trim(",") & "] " & Copient.PhraseLib.Lookup("term.expiredofferstate", LanguageID)
                            MyCommon.Write_Log(LogFile, "Some of the Offer(s) [OfferID: " & NotAppliedOfferLst.Replace(",", ", ").Trim(",") & "] is/are  in Expired status.", True)
                        End If
                        If (ActiveOfferLst.Length > 0) Then
                            Status = Status & Copient.PhraseLib.Lookup("term.someoffers", LanguageID) & " [OfferID: " & ActiveOfferLst.Replace(",", ", ").Trim(",") & "] " & Copient.PhraseLib.Lookup("term.activeofferstate", LanguageID)
                            MyCommon.Write_Log(LogFile, "Some of the Offer(s) [OfferID: " & ActiveOfferLst.Replace(",", ", ").Trim(",") & "] is/are in Active status.", True)
                        End If
                    ElseIf (Operation = UPDATE_END_DATE) Then
                        If (NotAppliedOfferLst.Length > 0) Then
                            Status = Status & Copient.PhraseLib.Lookup("term.someoffers", LanguageID) & " [OfferID: " & NotAppliedOfferLst.Replace(",", ", ").Trim(",") & "] " & Copient.PhraseLib.Lookup("term.expiredofferstate", LanguageID)
                            MyCommon.Write_Log(LogFile, "Some of the Offer(s) [OfferID: " & NotAppliedOfferLst.Replace(",", ", ").Trim(",") & "] is/are  in Expired status.", True)
                        End If
                    Else
                        If (NotAppliedOfferLst.Length > 0) Then
                            Status = Status & Copient.PhraseLib.Lookup("term.someoffers", LanguageID) & " [OfferID: " & NotAppliedOfferLst.Replace(",", ", ").Trim(",") & "] " & Copient.PhraseLib.Lookup("term.active-expiredstate", LanguageID)
                            MyCommon.Write_Log(LogFile, "Some of the Offer(s) [OfferID: " & NotAppliedOfferLst.Replace(",", ", ").Trim(",") & "] is/are  in Active/Expired status.", True)
                        End If
                    End If
                End If
            End If
            If UserModifiedOffers <> "" Then
                Status = Status & Copient.PhraseLib.Lookup("term.someoffers", LanguageID) & " [OfferID: " & UserModifiedOffers & "] are user modified"
            End If
            If ValidOfferStr.Length > 0 AndAlso LockOffersAfterExpiration AndAlso IsOfferEndDateExpired Then
                MyCommon.Write_Log(LogFile, "Auto deployment of Expired offers started.", True)

                Try
                    OffersLockedDeployment = ValidOfferStr.Split(",").Select(Function(x) Int64.Parse(x)).ToList()
                    Dim requestResolver = New ResolverBuilder()
                    requestResolver.Build()
                    RegisterDependencies(requestResolver)
                    Dim offerService As IOffer = requestResolver.Container.Resolve(Of IOffer)()
                    Dim deploymentResult As AMSResult(Of Boolean)
                    deploymentResult = offerService.DeployOffers(OffersLockedDeployment, AdminUserID)
                    MyCommon.Write_Log(LogFile, "Expired Offers got auto deployed.", True)
                    MyCommon.Write_Log(LogFile, String.Format("Deployed offers are {0}. {1}", OffersLockedDeployment.Count, deploymentResult.MessageString), True)
                    For i As Integer = 0 To OffersLockedDeployment.Count - 1
                        MyCommon.Activity_Log(3, OffersLockedDeployment(i), AdminUserID, Copient.PhraseLib.Lookup("history.offer-deploy", LanguageID))
                    Next
                Catch ex As Exception
                    MyCommon.Write_Log(LogFile, "Error: " + ex.ToString)
                End Try
            End If
        Catch ex As Exception
            If bBeginTransactionRT Then
                MyCommon.QueryStr = "Rollback Transaction;"
                MyCommon.LRT_Execute()
            End If
            Status = ex.ToString()
            MyCommon.Write_Log(LogFile, Status, True)
        End Try

    End Sub


    Sub ValidateInfo(ByVal dtOfferIDs As DataTable, ByVal dtUEOfferIDs As DataTable, ByVal AdminUserID As Integer, ByVal bAllowTimeWithStartEndDates As Boolean, ByRef isValid As Boolean)

        LogFile = "MassActionLog." & MyCommon.Leading_Zero_Fill(Year(Today), 4) & MyCommon.Leading_Zero_Fill(Month(Today), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
        If bAllowTimeWithStartEndDates Then
            MyCommon.Write_Log(LogFile, "Validation of Offers started for Mass change of Start/End Date and Time....", True)
        Else
            MyCommon.Write_Log(LogFile, "Validation of Offers started for Mass change of Start/End Dates....", True)
        End If

        isValid = True

        If (dtOfferIDs.Rows.Count = 0) Then
            MyCommon.Write_Log(LogFile, Copient.PhraseLib.Lookup("folder.nooffers", LanguageID), True)
            isValid = False
            Exit Sub
        End If

        If MyCommon.Fetch_SystemOption(191) = "1" Then
            Dim dtMultipleOffers As DataTable
            MyCommon.QueryStr = "dbo.pa_CheckMultipleFolderOfferExists"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.AddWithValue("@OffersT", dtOfferIDs)
            dtMultipleOffers = MyCommon.LRTsp_select

            If dtMultipleOffers.Rows.Count > 0 AndAlso MyCommon.NZ(dtMultipleOffers.Rows(0)(0), -1) > 0 Then
                MyCommon.Write_Log(LogFile, Copient.PhraseLib.Lookup("folders.multiplefoldererror", LanguageID), True)
                isValid = False
            End If
        End If

        If (Not Logix.UserRoles.EditOffer) Then
            MyCommon.Write_Log(LogFile, Copient.PhraseLib.Lookup("folders.nopermissionerror", LanguageID), True)
            isValid = False
        End If


        If MyCommon.IsUserAssociatedWithAnyBuyer(AdminUserID) Then

            Dim dtErrorOffers As DataTable
            MyCommon.QueryStr = "dbo.pa_CheckBuyerPermission"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.AddWithValue("@OffersT", dtUEOfferIDs)
            MyCommon.LRTsp.Parameters.AddWithValue("@AdminUserID", AdminUserID)
            dtErrorOffers = MyCommon.LRTsp_select

            If dtErrorOffers.Rows.Count > 0 AndAlso MyCommon.NZ(dtErrorOffers.Rows(0)(0), -1) > 0 Then
                MyCommon.Write_Log(LogFile, Copient.PhraseLib.Lookup("folders.nopermissionerror", LanguageID), True)
                isValid = False
            End If
        End If
    End Sub

    Sub RegisterDependencies(ByRef resolverBuilder As ResolverBuilder)
        resolverBuilder.Container.RegisterType(Of ICustomerGroups, CustomerGroups)()
        resolverBuilder.Container.RegisterType(Of ICustomerGroupCondition, CustomerGroupConditionService)()
        resolverBuilder.Container.RegisterType(Of IPointsCondition, PointsConditionService)()
        resolverBuilder.Container.RegisterType(Of IStoredValueCondition, StoredValueConditionService)()
        resolverBuilder.Container.RegisterType(Of IStoredValueProgramService, StoredValueProgramService)()
        resolverBuilder.Container.RegisterType(Of IPointsProgramService, PointsProgramService)()
        resolverBuilder.Container.RegisterType(Of ITrackableCouponConditionService, TrackableCouponConditionService)()
        resolverBuilder.Container.RegisterType(Of IPassThroughRewards, PassThroughRewardService)()
        resolverBuilder.Container.RegisterType(Of ICollisionDetectionService, CollisionDetectionService)()
        resolverBuilder.Container.RegisterType(Of ITrackableCouponProgramService, TrackableCouponProgramService)()
        resolverBuilder.Container.RegisterType(Of ICustomerService, CustomerService)()
        resolverBuilder.Container.RegisterType(Of ILocationsService, LocationsService)()
        resolverBuilder.Container.RegisterType(Of IOfferDeploymentValidator, OfferDeploymentValidator)()
        resolverBuilder.Container.RegisterType(Of IOffer, OfferService)()
        resolverBuilder.Container.RegisterType(Of ILoginWithNEP, LoginWithNEP)()
        resolverBuilder.Container.RegisterType(Of IRestServiceHelper, RESTServiceHelper)()
        resolverBuilder.Container.RegisterType(Of IOfferApprovalWorkflowService, OfferApprovalWorkflowService)()
        resolverBuilder.Container.RegisterType(Of INotificationService, NotificationService)()
    End Sub

    Sub MassDeployOffers(ByVal ItemIDs As String, ByVal FromOfferList As Boolean, ByVal OffersWithoutConditions As Boolean, ByVal AdminUserID As Integer,
                         ByVal deploytransreqskip As String, ByVal deferdeploy As String, Optional ByVal FolderID As Integer = -1, Optional ByRef status As String = "", Optional ByVal isPendingOffersExists As Boolean = False)
        'Response.Cache.SetCacheability(HttpCacheability.NoCache)
        LogFile = "MassActionLog." & MyCommon.Leading_Zero_Fill(Year(Today), 4) & MyCommon.Leading_Zero_Fill(Month(Today), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
        MyCommon.Write_Log(LogFile, "----------------------------------------------------------------")
        Dim requestResolver = New ResolverBuilder()
        requestResolver.Build()

        Try
            RegisterDependencies(requestResolver)

            Dim offerDeploymentValidator As IOfferDeploymentValidator = requestResolver.Container.Resolve(Of IOfferDeploymentValidator)()
            Dim offerService As IOffer = requestResolver.Container.Resolve(Of IOffer)()
            Dim ValidationResult As AMSResult(Of DataTable) = New AMSResult(Of DataTable)()
            Dim OffersWithEnginesDT As DataTable
            Dim IsValid As Boolean

            'Get the dictionary of OfferIds with corresponding Engine ID 
            OffersWithEnginesDT = GetEngineID(ItemIDs, FromOfferList)
            If FromOfferList Then
                MyCommon.Write_Log(LogFile, "Performed Action(Deploy) on " & OffersWithEnginesDT.Rows.Count & " offers from Offer List.", True)
            Else
                MyCommon.Write_Log(LogFile, "Performed Action(Deploy) on " & OffersWithEnginesDT.Rows.Count & " offers from Folders.", True)
            End If
            If deferdeploy = "" Then
                deferdeploy = "0"
            End If
            ValidationResult = offerDeploymentValidator.ValidateOffers(OffersWithEnginesDT, False, False, True, True, False, FolderID, AdminUserId:=AdminUserID, LangID:=LanguageID)
            If ValidationResult.ResultType = AMSResultType.Success Then
                IsValid = True
            Else
                IsValid = False
            End If
            If (ValidationResult.Result.Rows.Count > 0) Then
                Dim InvalidOfferStr As String = ValidationResult.Result.Rows(0).Item("ReturnMessage").ToString().TrimEnd(",")
                Dim ValidOfferStr As String = ValidationResult.Result.Rows(1).Item("ReturnMessage").ToString().TrimEnd(",")
                Dim ValidOfferList As New List(Of Int64)
                If ValidOfferStr.Length > 0 Then
                    ValidOfferList = ValidOfferStr.Split(",").Select(Function(x) Int64.Parse(x)).ToList()
                End If
                Dim OffersPendingDeployment As New List(Of Int64)
                Dim lstCollidingOffers As AMSResult(Of List(Of Int64))
                Dim deploymentResult As AMSResult(Of Boolean)

                If ValidOfferStr.Length > 0 Then
                    MyCommon.Write_Log(LogFile, String.Format("Collision Detection Initiated for Offer IDs: {0}", ValidOfferStr), True)
                    lstCollidingOffers = offerService.ProcessOfferCollisionDetectionFolderDeployment(ValidOfferStr)
                    MyCommon.Write_Log(LogFile, String.Format("Collision Detection Completed for Offer IDs: {0}", ValidOfferStr), True)
                    If lstCollidingOffers.ResultType = AMSResultType.Success AndAlso lstCollidingOffers.Result IsNot Nothing Then
                        CollisionsDetected = (lstCollidingOffers.Result.Count > 0)
                        OffersPendingDeployment = ValidOfferList.Except(lstCollidingOffers.Result).ToList()
                    ElseIf lstCollidingOffers.Result Is Nothing Then
                        OffersPendingDeployment = ValidOfferList
                    End If
                End If

                'Deploy All Offers which have Passed Collision Detection
                If OffersPendingDeployment.Count > 0 Then
                    deploymentResult = offerService.DeployOffers(OffersPendingDeployment, AdminUserID)
                    MyCommon.Write_Log(LogFile, String.Format("Deployment completed for {0} offers. {1}", OffersPendingDeployment.Count, deploymentResult.MessageString), True)
                End If
            End If
            If (IsValid) Then
                status = Copient.PhraseLib.Lookup("term.success1", LanguageID) + " " + Copient.PhraseLib.Lookup("deploy-success", LanguageID)
                If isPendingOffersExists Then
                    status &= " " + Copient.PhraseLib.Lookup("term.ignore-pendingoffers", LanguageID)
                End If
            Else
                status = Copient.PhraseLib.Lookup("deploy-failed", LanguageID)

            End If
        Catch ex As Exception
            MyCommon.Write_Log(LogFile, ex.ToString(), True)
            status = ex.ToString()
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            requestResolver.Container.Dispose()
        End Try
    End Sub
    Function RemovePendingOffers(ByVal offersDT As DataTable, ByRef isPendingOffersExists As Boolean) As DataTable
        Dim dtOffers As DataTable = New DataTable()
        MyCommon.QueryStr = "dbo.pt_RemovePendingOffers"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@OffersT", SqlDbType.Structured).Value = offersDT
        dtOffers = MyCommon.LRTsp_select
        MyCommon.Close_LRTsp()
        isPendingOffersExists = IIf(dtOffers.Rows.Count = offersDT.Rows.Count, False, True)
        Return dtOffers
    End Function
    Sub MassDeferDeployOffers(ByVal ItemIDs As String, ByVal FromOfferList As Boolean, ByVal OffersWithoutConditions As Boolean, ByVal AdminUserID As Integer,
    ByVal deploytransreqskip As String, ByVal deferdeploy As String, Optional ByVal FolderID As Integer = -1, Optional ByRef status As String = "", Optional ByVal isPendingOffersExists As Boolean = False)
        'Response.Cache.SetCacheability(HttpCacheability.NoCache)
        LogFile = "MassActionLog." & MyCommon.Leading_Zero_Fill(Year(Today), 4) & MyCommon.Leading_Zero_Fill(Month(Today), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
        MyCommon.Write_Log(LogFile, "----------------------------------------------------------------")
        Dim requestResolver = New ResolverBuilder()
        requestResolver.Build()
        Dim filteredOffers As New List(Of String)
        Dim eligibleOffers As New List(Of String)
        Dim OfferIDs As String = ""
        Dim filteredIDS As String = ""
        Try
            RegisterDependencies(requestResolver)

            Dim offerDeploymentValidator As IOfferDeploymentValidator = requestResolver.Container.Resolve(Of IOfferDeploymentValidator)()
            Dim offerService As IOffer = requestResolver.Container.Resolve(Of IOffer)()

            Dim ValidationResult As AMSResult(Of DataTable) = New AMSResult(Of DataTable)()
            Dim OffersWithEnginesDT As DataTable
            Dim IsValid As Boolean

            Dim bEnableSeperatePermissionForDeferDeploy = IIf(MyCommon.Fetch_SystemOption(262) = "1", True, False)
            Dim bRestrictOperationForActiveOrExpired = IIf(MyCommon.Fetch_SystemOption(272) = "1", True, False)
            Dim bRestrictUserModifiedOffers = IIf(MyCommon.Fetch_SystemOption(284) = "1", True, False)

            'Get the dictionary of OfferIds with corresponding Engine ID 
            OffersWithEnginesDT = GetEngineID(ItemIDs, FromOfferList)
            OfferIDs = [String].Join(",", OffersWithEnginesDT.AsEnumerable().[Select](Function(row1) row1.Field(Of Int64)(0)))


            If FromOfferList Then
                MyCommon.Write_Log(LogFile, "Started Action(Defer Deploy) on " & OffersWithEnginesDT.Rows.Count & " offers from Offer List.", True)
            Else
                MyCommon.Write_Log(LogFile, "Started Action(Defer Deploy) on " & OffersWithEnginesDT.Rows.Count & " offers from Folders.", True)
            End If

            MyCommon.Write_Log(LogFile, "Checking Offer(s) eligibility for defer deploy (OfferID(s) : " & OfferIDs & ")", True)

            eligibleOffers = OfferIDs.Split(",").ToList()

            If (Not Logix.UserRoles.DeferDeployOffersPastLockoutPeriod) Then
                'Filter Lockedout offers , If the Folder is locked out then all the folders selected under this folder were in lockout state
                'Breaking the loop after checking first offer as the lockout period is calculated against folder start date.
                For Each offerid In eligibleOffers
                    If (MyCommon.IsLockOutPeriod(MyCommon, offerid)) Then
                        MyCommon.Write_Log(LogFile, "Defer Deploy cannot be performed on selected offers (Offer ID(s):" & OfferIDs & ") as the User has no permission to defer deploy offers which are past lockout period.", True)
                        eligibleOffers.Clear()
                        Exit For
                    End If
                Next
            End If

            If (eligibleOffers.Count > 0) Then
                Dim dtItemIDs As DataTable = New DataTable()
                Dim dtOffers As DataTable = New DataTable()
                dtItemIDs.TableName = "OfferIDs"
                dtItemIDs.Columns.Add("id", System.Type.GetType("System.Int64"))

                For Each offerid In eligibleOffers
                    dtItemIDs.Rows.Add(offerid)
                Next
                If dtItemIDs.Rows.Count > 0 Then
                    If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
                    MyCommon.QueryStr = "dbo.pa_GetOfferDetailsForDeferDeployOperation"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.AddWithValue("@OffersT", dtItemIDs)
                    dtOffers = MyCommon.LRTsp_select
                    MyCommon.Close_LRTsp()
                End If

                Dim row() As DataRow = Nothing


                If ((bEnableSeperatePermissionForDeferDeploy AndAlso Not Logix.UserRoles.DeferDeployTemplateOffers) OrElse (Not bEnableSeperatePermissionForDeferDeploy AndAlso Not Logix.UserRoles.DeployTemplateOffers)) Then
                    row = dtOffers.Select("FromTemplate=1")
                    For Each dr In row
                        Dim offerId As String = dr.Item("OfferID")
                        If (eligibleOffers.Contains(offerId)) Then
                            filteredOffers.Add(offerId)
                        End If
                    Next
                    If (filteredOffers.Count > 0) Then
                        MyCommon.Write_Log(LogFile, "Filtering out template offers (Offer ID(s):" & String.Join(",", filteredOffers) & ") as User has no permission to defer deploy template offers.", True)
                        eligibleOffers.RemoveAll(Function(item) filteredOffers.Contains(item))
                        filteredIDS = filteredIDS & String.Join(",", filteredOffers)
                    End If
                End If

                filteredOffers.Clear()
                If ((bEnableSeperatePermissionForDeferDeploy AndAlso Not Logix.UserRoles.DeferDeployNonTemplateOffers) OrElse (Not bEnableSeperatePermissionForDeferDeploy AndAlso Not Logix.UserRoles.DeployNonTemplateOffers)) Then
                    row = dtOffers.Select("FromTemplate=0")
                    For Each dr In row
                        Dim offerId As String = dr.Item("OfferID")
                        If (eligibleOffers.Contains(offerId)) Then
                            filteredOffers.Add(offerId)
                        End If
                    Next
                    If (filteredOffers.Count > 0) Then
                        MyCommon.Write_Log(LogFile, "Filtering out non-template offers (Offer ID(s):" & String.Join(",", filteredOffers) & ") as User has no permission to defer deploy non-template offers.", True)
                        eligibleOffers.RemoveAll(Function(item) filteredOffers.Contains(item))
                        filteredIDS = filteredIDS & String.Join(",", filteredOffers)
                    End If
                End If

                'Filtering Templates from the selected list
                filteredOffers.Clear()
                row = dtOffers.Select("IsTemplate=1")
                For Each dr In row
                    Dim offerId As String = dr.Item("OfferID")
                    If (eligibleOffers.Contains(offerId)) Then
                        filteredOffers.Add(offerId)
                    End If
                Next
                If (filteredOffers.Count > 0) Then
                    MyCommon.Write_Log(LogFile, "Filtering out Templates  (Template ID(s):" & String.Join(",", filteredOffers) & ") as templates cannot be defer-deployed.", True)
                    eligibleOffers.RemoveAll(Function(item) filteredOffers.Contains(item))
                    filteredIDS = filteredIDS & String.Join(",", filteredOffers)
                End If

                'Filter Expired offers
                filteredOffers.Clear()
                For Each offerid In eligibleOffers
                    Dim StatusCode As Copient.LogixInc.STATUS_FLAGS = Copient.LogixInc.STATUS_FLAGS.STATUS_UNKNOWN
                    Logix.GetOfferStatus(offerid, LanguageID, StatusCode)
                    If (StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_EXPIRED) Then
                        filteredOffers.Add(offerid)
                    End If
                Next
                If (filteredOffers.Count > 0) Then
                    MyCommon.Write_Log(LogFile, "Filtering out Expired offers (Offer ID(s):" & String.Join(",", filteredOffers) & ") as expired offers cannot be defer-deployed.", True)
                    eligibleOffers.RemoveAll(Function(item) filteredOffers.Contains(item))
                    filteredIDS = filteredIDS & String.Join(",", filteredOffers)
                End If



                filteredOffers.Clear()
                If (eligibleOffers.Count > 0) Then
                    'Filter Active offers which are processed for deployment i.e donot defer deploy offers which are processed for deployment
                    filteredOffers.Clear()
                    row = dtOffers.Select("StatusFlag = 2 AND OfferID IN(" & String.Join(",", eligibleOffers) & ")")
                    For Each dr In row
                        Dim offerId As String = dr.Item("OfferID")
                        filteredOffers.Add(offerId)
                    Next
                    If (filteredOffers.Count > 0) Then
                        MyCommon.Write_Log(LogFile, "Filtering out offers (Offer ID(s):" & String.Join(",", filteredOffers) & ") which are processed for deployment.", True)
                        eligibleOffers.RemoveAll(Function(item) filteredOffers.Contains(item))
                        filteredIDS = filteredIDS & String.Join(",", filteredOffers)
                    End If
                Else
                    status = Copient.PhraseLib.Lookup("bulkdeferdeploy-completed", LanguageID)
                    MyCommon.Write_Log(LogFile, "Defer Deployment completed.No Offers found to defer-deploy after filtering out the selected list.", True)
                End If


                If (eligibleOffers.Count > 0) Then
                    'Filter Active offers which are not modified i.e Active offers which are modified are eligible for defer-deploy operation
                    filteredOffers.Clear()
                    row = dtOffers.Select("StatusFlag <> 2 AND OfferID IN(" & String.Join(",", eligibleOffers) & ")")
                    For Each dr In row
                        Dim offerId As String = dr.Item("OfferID")
                        Dim StatusCode As Copient.LogixInc.STATUS_FLAGS = Copient.LogixInc.STATUS_FLAGS.STATUS_UNKNOWN
                        Logix.GetOfferStatus(offerId, LanguageID, StatusCode)
                        If (StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_ACTIVE) Then
                            If (MyCommon.NZ(dr.Item("StatusFlag"), 0) > 0) Then
                                'Offer is active but Modified 
                            Else
                                filteredOffers.Add(offerId)
                            End If
                        End If
                    Next
                    If (filteredOffers.Count > 0) Then
                        MyCommon.Write_Log(LogFile, "Filtering out Active offers (Offer ID(s):" & String.Join(",", filteredOffers) & ") as active offers which are NOT modified cannot be defer-deployed.", True)
                        eligibleOffers.RemoveAll(Function(item) filteredOffers.Contains(item))
                        filteredIDS = filteredIDS & String.Join(",", filteredOffers)
                    End If
                    filteredOffers.Clear()
                Else
                    status = Copient.PhraseLib.Lookup("bulkdeferdeploy-completed", LanguageID)
                    MyCommon.Write_Log(LogFile, "Defer Deployment completed.No Offers found to defer-deploy after filtering out the selected list.", True)
                End If

                filteredOffers.Clear()
                If (eligibleOffers.Count > 0) Then
                    'Filtering User modified Offers
                    If (bRestrictUserModifiedOffers) Then
                        row = dtOffers.Select("UserModifiedOffer=1 AND OfferID IN(" & String.Join(",", eligibleOffers) & ")")
                        For Each dr In row
                            Dim offerId As String = dr.Item("OfferID")
                            If (eligibleOffers.Contains(offerId)) Then
                                filteredOffers.Add(offerId)
                            End If
                        Next

                        If (filteredOffers.Count > 0) Then
                            MyCommon.Write_Log(LogFile, "Filtering out UserModified offers (Offer ID(s):" & String.Join(",", filteredOffers) & ") as defer-deploy is restricted on user modified offers.", True)
                            eligibleOffers.RemoveAll(Function(item) filteredOffers.Contains(item))
                            filteredIDS = filteredIDS & String.Join(",", filteredOffers)
                        End If

                        filteredOffers.Clear()
                    End If
                Else
                    status = Copient.PhraseLib.Lookup("bulkdeferdeploy-completed", LanguageID)
                    MyCommon.Write_Log(LogFile, "Defer Deployment completed.No Offers found to defer-deploy after filtering out the selected list.", True)
                End If

                filteredOffers.Clear()
                If (eligibleOffers.Count > 0) Then

                    MyCommon.QueryStr = " SELECT OfferID,EngineID FROM OfferIds WHERE OfferID in(" & String.Join(",", eligibleOffers) & ");   "
                    OffersWithEnginesDT = MyCommon.LRT_Select


                    ValidationResult = offerDeploymentValidator.ValidateOffers(OffersWithEnginesDT, False, True, True, True, False, FolderID, AdminUserId:=AdminUserID)
                    If ValidationResult.ResultType = AMSResultType.Success Then
                        IsValid = True
                    Else
                        IsValid = False
                    End If

                    Dim InvalidOfferStr As String = ValidationResult.Result.Rows(0).Item("ReturnMessage").ToString().TrimEnd(",")
                    Dim ValidOfferStr As String = ValidationResult.Result.Rows(1).Item("ReturnMessage").ToString().TrimEnd(",")
                    Dim ValidOfferList As New List(Of Int64)
                    If ValidOfferStr.Length > 0 Then
                        ValidOfferList = ValidOfferStr.Split(",").Select(Function(x) Int64.Parse(x)).ToList()
                    End If
                    Dim OffersPendingDeferDeployment As New List(Of Int64)
                    Dim lstCollidingOffers As AMSResult(Of List(Of Int64))
                    Dim deferdeploymentResult As AMSResult(Of Boolean)

                    If ValidOfferStr.Length > 0 Then
                        MyCommon.Write_Log(LogFile, String.Format("Collision Detection Initiated for Offer IDs: {0}", ValidOfferStr), True)
                        lstCollidingOffers = offerService.ProcessOfferCollisionDetectionFolderDeployment(ValidOfferStr)
                        MyCommon.Write_Log(LogFile, String.Format("Collision Detection Completed for Offer IDs: {0}", ValidOfferStr), True)
                        If lstCollidingOffers.ResultType = AMSResultType.Success AndAlso lstCollidingOffers.Result IsNot Nothing Then
                            CollisionsDetected = (lstCollidingOffers.Result.Count > 0)
                            OffersPendingDeferDeployment = ValidOfferList.Except(lstCollidingOffers.Result).ToList()
                        ElseIf lstCollidingOffers.Result Is Nothing Then
                            OffersPendingDeferDeployment = ValidOfferList
                        End If
                    Else
                        status = Copient.PhraseLib.Lookup("deferdeploy-failed", LanguageID)
                        MyCommon.Write_Log(LogFile, "Defer Deployment completed.No Offers found to defer-deploy after filtering out the selected list.", True)
                    End If

                    'Deploy All Offers which have Passed Collision Detection
                    If OffersPendingDeferDeployment.Count > 0 Then
                        deferdeploymentResult = offerService.DeferDeployOffers(OffersPendingDeferDeployment, AdminUserID)
                        MyCommon.Write_Log(LogFile, String.Format("Defer Deployment completed for {0} offers. {1}", OffersPendingDeferDeployment.Count, deferdeploymentResult.MessageString), True)
                    Else
                        status = Copient.PhraseLib.Lookup("deferdeploy-failed", LanguageID)
                        MyCommon.Write_Log(LogFile, "Defer Deployment completed.No Offers found to defer-deploy after filtering out the selected list.", True)
                    End If

                    If (IsValid) Then
                        If (Not String.IsNullOrEmpty(filteredIDS)) Then
                            status = Copient.PhraseLib.Lookup("term.success1", LanguageID) + " " + Copient.PhraseLib.Lookup("bulkdeferdeploy-warning", LanguageID)
                        Else
                            status = Copient.PhraseLib.Lookup("term.success1", LanguageID) + " " + Copient.PhraseLib.Lookup("deferdeploy-success", LanguageID)
                            If isPendingOffersExists Then
                                status &= " " + Copient.PhraseLib.Lookup("term.ignore-pendingoffers", LanguageID)
                            End If
                        End If

                    Else
                        status = Copient.PhraseLib.Lookup("deferdeploy-failed", LanguageID)

                    End If
                Else
                    status = Copient.PhraseLib.Lookup("bulkdeferdeploy-completed", LanguageID)
                    MyCommon.Write_Log(LogFile, "Defer Deployment completed.No Offers found to defer-deploy after filtering out the selected list.", True)
                End If
            Else
                status = Copient.PhraseLib.Lookup("bulkdeferdeploy-completed", LanguageID)
                MyCommon.Write_Log(LogFile, "Defer Deployment completed.No Offers found to defer-deploy after filtering out the selected list.", True)
            End If
        Catch ex As Exception
            MyCommon.Write_Log(LogFile, ex.ToString(), True)
            status = ex.ToString()
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            requestResolver.Container.Dispose()
        End Try
    End Sub

    Sub NavigatetoReports(ByVal ItemIDs As String, ByVal FromOfferList As Boolean)
        Dim OfferIDs As New StringBuilder("")
        Dim Status As String
        'AMSPS-2009
        Dim OfferIDsOnly As New StringBuilder("")

        LogFile = "MassActionLog." & MyCommon.Leading_Zero_Fill(Year(Today), 4) & MyCommon.Leading_Zero_Fill(Month(Today), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
        'AMSPS-2231
        'MyCommon.Write_Log(LogFile, "ItemIDs(before) - " + ItemIDs, True)
        If FromOfferList = False Then ItemIDs = GetItemIdsSortByOfferIds(Replace(ItemIDs, ",NaN", ""))
        'AMSPS-2231 above
        MyCommon.Write_Log(LogFile, "Action - Navigate to reports is started for the selected Offers by " + AdminName, True)

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

            If FromOfferList Then
                MyCommon.Write_Log(LogFile, "Action - Navigate to reports - Performed on offers from offer list.", True)
                OfferIDs.Append(ItemIDs)
                For Each Item In ItemIDs.Split(",")
                    OfferIDs.Append("<option value=""" & Item & """>" & Item)
                Next
                'AMSPS-2009
                OfferIDsOnly.Append(ItemIDs)

            Else
                MyCommon.Write_Log(LogFile, "Action - Navigate to reports - Performed on offers from folder by " + AdminName, True)
                Dim dtOfferIDs As DataTable
                dtOfferIDs = GetEngineID(ItemIDs, False)

                For Each row In dtOfferIDs.Rows
                    OfferIDs.Append("<option value=""" & row(0).ToString() & """>" & row(0).ToString())
                    'AMSPS-2009
                    If OfferIDsOnly.ToString() <> String.Empty Then
                        OfferIDsOnly.Append(",")
                    End If
                    OfferIDsOnly.Append(row(0).ToString())

                Next
            End If

            If OfferIDs.ToString() <> String.Empty Then
                Session.Add("OFFERIDS", OfferIDs.ToString())
                'AMSPS-2009
                Session.Add("OFFERIDSONLY", OfferIDsOnly.ToString())

                Status = SUCCESS + "Action - Navigate to reports - was successful."
                MyCommon.Write_Log(LogFile, Status, True)
            Else
                Status = "Action - Navigate to reports - was failed. Please check MassActionLog for more information."
                MyCommon.Write_Log(LogFile, Status, True)
            End If
        Catch ex As Exception
            Status = ex.ToString()
            MyCommon.Write_Log(LogFile, Status, True)
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try
    End Sub

    'AMSPS-2231
    Function GetItemIdsSortByOfferIds(ByVal ItemIds As String) As String
        Dim rst As DataTable
        Dim ItemIDsSortedByOfferIds As String = ""
        Dim i As Long
        MyCommon.QueryStr = "Select ItemID FROM [LogixRT].[dbo].[FolderItems] WHERE ItemID in (SELECT items FROM Split (@ItemIDs, ',')) ORDER BY LinkID"
        MyCommon.DBParameters.Add("@ItemIDs", SqlDbType.NVarChar).Value = ItemIds
        rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        
        If rst.Rows.Count > 0 Then
            For i = 0 To rst.Rows.Count - 1
                ItemIDsSortedByOfferIds = ItemIDsSortedByOfferIds + MyCommon.NZ(rst.Rows(i).Item("ItemID"), 0).ToString
                If (i < rst.Rows.Count - 1) Then
                    ItemIDsSortedByOfferIds = ItemIDsSortedByOfferIds + ","
                End If
            Next i
        End If
        GetItemIdsSortByOfferIds = ItemIDsSortedByOfferIds
    End Function
    'AMSPS-2231 above

    Sub MassSendOutBoundAsync(ByVal State As Object)
        Dim ItemIDs As String = State(0)
        Dim FromOfferList As Boolean = State(1)
        Dim FolderID As Integer = State(2)
        Dim Status As String = ""
        Dim err As String = String.Empty
        Try
            If (FromOfferList) Then
                SetFolderMassOperationStatus(-1, FOLDER_IN_USE + " Mass Operation is started for the selected Offers by" + " " + AdminName)
            Else
                SetFolderMassOperationStatus(FolderID, FOLDER_IN_USE + " Mass Send OutBound is started for the Offers in this Folder by" + " " + AdminName)
            End If

            MassSendOutBound(ItemIDs, FromOfferList, Status)
        Catch ex As Exception
            err = "Error: " + ex.ToString
            MyCommon.Write_Log(LogFile, err)
        Finally
            If (FromOfferList) Then
                SetFolderMassOperationStatus(-1, FOLDER_NOT_IN_USE + Status)
            Else
                SetFolderMassOperationStatus(FolderID, FOLDER_NOT_IN_USE + Status)
            End If
        End Try
    End Sub

    Sub ValidateForSendOutBound(ByVal OfferIDs As DataTable, ByRef IsValid As Boolean)
        Dim iCRMType As Integer
        Dim rst As DataTable
        Dim StatusCode As Copient.LogixInc.STATUS_FLAGS = Copient.LogixInc.STATUS_FLAGS.STATUS_UNKNOWN

        Dim SystemOption80 As String = MyCommon.Fetch_CPE_SystemOption(80)
        Dim SystemOption48 As String = MyCommon.Fetch_CM_SystemOption(48)

        Dim Message As New StringBuilder("")
        Dim CRMENgineID As Integer = 0
        Dim InboundCRMEngineID As Integer = 0
        IsValid = True
        Try

            If Not Integer.TryParse(MyCommon.Fetch_SystemOption(25), iCRMType) Then iCRMType = 0

            If iCRMType <= 0 Then
                Message.Append("Validation Failed!!! CRM is not enabled" + vbLf)
                IsValid = False
            End If

            For Each row In OfferIDs.Rows
                If SystemOption80 = "1" Then
                    Logix.CalcOfferStatusText(row(0), LanguageID, Now, StatusCode)
                    If StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_EXPIRED Then
                        Message.Append("OfferID=" + row(0).ToString() + ": " + Copient.PhraseLib.Lookup("cpeoffer-sum.deployalertforexpire", LanguageID) + vbLf)
                        IsValid = False
                    End If
                End If

                If Convert.ToInt16(row(1)) = Engines.CM Then
                    MyCommon.QueryStr = "select CRMEngineID,InboundCRMEngineID from Offers with (NoLock) where OfferID = " & row(0).ToString() & " and Deleted=0 "
                    rst = MyCommon.LRT_Select

                    If rst.Rows.Count > 0 Then
                        InboundCRMEngineID = MyCommon.NZ(rst.Rows(0).Item("InboundCRMEngineID"), 0)
                        CRMENgineID = MyCommon.NZ(rst.Rows(0).Item("CRMEngineID"), 0)
                    End If
                    If CRMENgineID = 0 Then
                        Message.Append("Validation Failed!!! CRM is not enabled for OfferID=" + row(0).ToString() + vbLf)
                        IsValid = False
                    End If

                    If SystemOption48 = "0" AndAlso InboundCRMEngineID <> 0 Then
                        Message.Append("Manual Send Outbound for offers created externally is not enabled for OfferID=" + row(0).ToString() + vbLf)
                        IsValid = False
                    End If
                Else
                    MyCommon.QueryStr = "select CRMEngineID from CPE_INCENTIVES with (NoLock) where INCENTIVEID = " & row(0).ToString() & " and Deleted=0 AND CRMEngineID=0"
                    rst = MyCommon.LRT_Select
                    If rst.Rows.Count > 0 Then
                        Message.Append("Validation Failed!!! CRM is not enabled for OfferID=" + row(0).ToString() + vbLf)
                        IsValid = False
                    End If
                End If

            Next

        Catch ex As Exception
            Message.Append(ex.Message)
        Finally
            MyCommon.Write_Log(LogFile, Message.ToString(), True)
        End Try

    End Sub

    Sub MassSendOutBound(ByVal ItemIDs As String, ByVal FromOfferList As Boolean, ByRef Status As String)
        Dim ValidationMessage As String = ""
        Dim dtOfferIDs As New DataTable
        dtOfferIDs.Columns.Add("OfferID")
        Dim isPendingOffersExists As Boolean

        LogFile = "MassActionLog." & MyCommon.Leading_Zero_Fill(Year(Today), 4) & MyCommon.Leading_Zero_Fill(Month(Today), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
        MyCommon.Write_Log(LogFile, "Started Mass Send OutBound", True)

        Try
            dtOfferIDs = GetEngineID(ItemIDs, FromOfferList)
            dtOfferIDs = RemovePendingOffers(dtOfferIDs, isPendingOffersExists)
            dtOfferIDs.Columns.Remove("EngineID")

            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

            MyCommon.QueryStr = "dbo.pa_ValidateSendOutBound"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.AddWithValue("@OffersT", dtOfferIDs)
            MyCommon.LRTsp.Parameters.Add("@Message", SqlDbType.NVarChar, -1).Direction = ParameterDirection.Output
            MyCommon.LRTsp.ExecuteNonQuery()
            ValidationMessage = MyCommon.LRTsp.Parameters("@Message").Value

            If ValidationMessage.Trim() = String.Empty Then

                MyCommon.QueryStr = "dbo.pa_SendOutBound"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.AddWithValue("@OffersT", dtOfferIDs)
                MyCommon.LRTsp.Parameters.AddWithValue("@AdminUserID", AdminUserID)
                MyCommon.LRTsp.Parameters.AddWithValue("@LanguageID", LanguageID)
                MyCommon.LRTsp.ExecuteNonQuery()
                Status = Copient.PhraseLib.Lookup("term.success1", LanguageID) + " " + Copient.PhraseLib.Lookup("sendoutbound-success", LanguageID)
                If isPendingOffersExists Then
                    Status &= " " + Copient.PhraseLib.Lookup("term.ignore-pendingoffers", LanguageID)
                End If
            Else
                MyCommon.Write_Log(LogFile, "Validation Failed. " + ValidationMessage, True)
                Status = Copient.PhraseLib.Lookup("sendoutbound-failure", LanguageID)

            End If

        Catch ex As Exception
            Status = ex.ToString()
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            MyCommon.Write_Log(LogFile, Status, True)
        End Try
    End Sub


    Function GetEngineID(ByVal ItemIDs As String, ByVal fromOfferList As Boolean) As DataTable
        Dim dtOffers As DataTable = New DataTable()
        Try
            Dim dtItemIDs As DataTable = New DataTable()
            dtItemIDs.TableName = "OfferIDs"
            dtItemIDs.Columns.Add("id", System.Type.GetType("System.Int64"))

            For Each item In ItemIDs.Split(",")
                dtItemIDs.Rows.Add(Convert.ToInt64(item))
            Next
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

            If dtItemIDs.Rows.Count > 0 Then

                MyCommon.QueryStr = "dbo.pa_GetOfferIDsAndEngineIDs"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@IsFromOfferList", SqlDbType.Bit).Value = IIf(fromOfferList, 1, 0)
                MyCommon.LRTsp.Parameters.AddWithValue("@OffersT", dtItemIDs)
                dtOffers = MyCommon.LRTsp_select
                MyCommon.Close_LRTsp()
            End If
        Catch ex As Exception

        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()

        End Try

        Return dtOffers
    End Function


    Sub WFStatustoPreValidate(ByVal ItemIDs As String, ByVal FromOfferList As Boolean)

        'Dim itemid As Long = 0
        Dim dt As DataTable
        Dim iOfferId As Long
        Dim iItemID As Integer
        Dim aItemIDs As String() = Nothing
        Dim ErrorMessage As String = ""
        Dim FailedOffers As String = ""
        Dim FailedOffersDescription As String = ""
        Dim ResponseValidatePreVal As String
        Dim OffersRequireReval As String = ""
        Dim ErrorLogDescription As New StringBuilder()
        Dim LogDescription As New StringBuilder()
        Dim bBeginTransactionRT As Boolean = False

        Response.Cache.SetCacheability(HttpCacheability.NoCache)

        LogFile = "MassActionLog." & MyCommon.Leading_Zero_Fill(Year(Today), 4) & MyCommon.Leading_Zero_Fill(Month(Today), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
        MyCommon.Write_Log(LogFile, "----------------------------------------------------------------")

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            aItemIDs = ItemIDs.Split(",")
            If FromOfferList Then
                MyCommon.Write_Log(LogFile, "Action - Changing status to pre validate - Performed on " & aItemIDs.Length & " offers from offer list.", True)
                MyCommon.QueryStr = "Begin Transaction;"
                MyCommon.LRT_Execute()
                bBeginTransactionRT = True
                For i = 0 To aItemIDs.GetUpperBound(0)
                    iOfferId = aItemIDs(i)
                    ErrorMessage = ValidateForPreValidate(MyCommon, iOfferId)
                    If (ErrorMessage = "") Then
                        ResponseValidatePreVal = AssignPreValidate(MyCommon, iOfferId)
                        If ResponseValidatePreVal <> "" Then
                            OffersRequireReval = iOfferId & vbCrLf
                        End If
                    Else
                        If FailedOffers = "" Then
                            ErrorLogDescription.AppendLine("The Action- Changing status to pre validate - was not successful. Following offers were failed during validation:")
                        End If
                        ErrorLogDescription.AppendLine(iOfferId & "  |  " & ErrorMessage)
                        FailedOffers = FailedOffers & iOfferId & ","
                        FailedOffersDescription = FailedOffersDescription & ErrorMessage & "|"
                    End If
                Next
            Else
                MyCommon.Write_Log(LogFile, "Action - Changing status to pre validate - Performed on " & aItemIDs.Length & " offers from folders.", True)
                MyCommon.QueryStr = "Begin Transaction;"
                MyCommon.LRT_Execute()
                bBeginTransactionRT = True
                For i = 0 To aItemIDs.GetUpperBound(0)
                    iItemID = aItemIDs(i)
                    MyCommon.QueryStr = "select LinkID from FolderItems with (NoLock) where ItemID=" & iItemID
                    MyCommon.LRT_Execute()
                    dt = MyCommon.LRT_Select()
                    For Each row In dt.Rows
                        iOfferId = MyCommon.NZ(row.Item("LinkID"), 0)
                        ErrorMessage = ValidateForPreValidate(MyCommon, iOfferId)
                        If (ErrorMessage = "") Then
                            ResponseValidatePreVal = AssignPreValidate(MyCommon, iOfferId)
                            If ResponseValidatePreVal <> "" Then
                                OffersRequireReval = iOfferId & vbCrLf
                            End If
                        Else
                            If FailedOffers = "" Then
                                ErrorLogDescription.AppendLine("The Action- Changing status to pre validate - was not successful. Following offers were failed during validation:")
                            End If
                            ErrorLogDescription.AppendLine(iOfferId & "  |  " & ErrorMessage)
                            FailedOffers = FailedOffers & iItemID & ","
                            FailedOffersDescription = FailedOffersDescription & ErrorMessage & "|"
                        End If
                    Next
                Next
            End If

            If FailedOffers = "" AndAlso bBeginTransactionRT Then
                MyCommon.QueryStr = "Commit Transaction;"
                MyCommon.LRT_Execute()
                If OffersRequireReval <> "" Then
                    LogDescription.AppendLine("The Action- Changing status to pre validate - was successful for all offers. Following offers require revalidation:")
                    LogDescription.AppendLine(OffersRequireReval)
                    MyCommon.Write_Log(LogFile, LogDescription.ToString(), True)
                    Send("ReqReValidation ||" & OffersRequireReval)
                Else
                    MyCommon.Write_Log(LogFile, "The Action- Changing status to pre validate - was successful for all offers.", True)
                    Send("OK")
                End If
            ElseIf FailedOffers <> "" AndAlso bBeginTransactionRT Then
                MyCommon.QueryStr = "Rollback Transaction;"
                MyCommon.LRT_Execute()
                MyCommon.Write_Log(LogFile, ErrorLogDescription.ToString(), True)
                Send(FailedOffers & ":" & FailedOffersDescription)
            End If
        Catch ex As Exception
            If bBeginTransactionRT Then
                MyCommon.QueryStr = "Rollback Transaction;"
                MyCommon.LRT_Execute()
            End If
            MyCommon.Write_Log(LogFile, ex.ToString(), True)
            Send(ex.ToString)
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try
    End Sub

    Sub WFStatustoPostValidate(ByVal ItemIDs As String, ByVal FromOfferList As Boolean)

        'Dim itemid As Long = 0
        Dim dt As DataTable
        Dim iOfferId As Long
        Dim iItemID As Integer
        Dim aItemIDs As String() = Nothing
        Dim ErrorMessage As String = ""
        Dim FailedOffers As String = ""
        Dim FailedOffersDescription As String = ""
        Dim ErrorLogDescription As New StringBuilder()
        Dim bBeginTransactionRT As Boolean = False

        Response.Cache.SetCacheability(HttpCacheability.NoCache)

        LogFile = "MassActionLog." & MyCommon.Leading_Zero_Fill(Year(Today), 4) & MyCommon.Leading_Zero_Fill(Month(Today), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
        MyCommon.Write_Log(LogFile, "----------------------------------------------------------------")

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            aItemIDs = ItemIDs.Split(",")
            If aItemIDs.Length > 0 Then
                ' order is important for Santa Bucks since parent OfferID is always less than Child OfferID
                Array.Sort(aItemIDs)
            End If
            If FromOfferList Then
                MyCommon.Write_Log(LogFile, "Action - Changing status to post validate - Performed on " & aItemIDs.Length & " offers from offer list.", True)
                MyCommon.QueryStr = "Begin Transaction;"
                MyCommon.LRT_Execute()
                bBeginTransactionRT = True
                For i = 0 To aItemIDs.GetUpperBound(0)
                    iOfferId = aItemIDs(i)
                    ErrorMessage = ValidateForPostValidate(MyCommon, iOfferId)
                    If (ErrorMessage = "") Then
                        AssignPostValidate(MyCommon, iOfferId)
                    Else
                        If FailedOffers = "" Then
                            ErrorLogDescription.AppendLine("The Action- Changing status to post validate - was not successful. Following offers were failed during validation:")
                        End If
                        ErrorLogDescription.AppendLine(iOfferId & "  |  " & ErrorMessage)
                        FailedOffers = FailedOffers & iOfferId & ","
                        FailedOffersDescription = FailedOffersDescription & ErrorMessage & "|"
                    End If
                Next
            Else
                MyCommon.Write_Log(LogFile, "Action - Changing status to post validate - Performed on " & aItemIDs.Length & " offers from folders.", True)
                MyCommon.QueryStr = "Begin Transaction;"
                MyCommon.LRT_Execute()
                bBeginTransactionRT = True
                For i = 0 To aItemIDs.GetUpperBound(0)
                    iItemID = aItemIDs(i)
                    MyCommon.QueryStr = "select LinkID from FolderItems with (NoLock) where ItemID=" & iItemID
                    MyCommon.LRT_Execute()
                    dt = MyCommon.LRT_Select()
                    For Each row In dt.Rows
                        iOfferId = MyCommon.NZ(row.Item("LinkID"), 0)
                        ErrorMessage = ValidateForPostValidate(MyCommon, iOfferId)
                        If (ErrorMessage = "") Then
                            AssignPostValidate(MyCommon, iOfferId)
                        Else
                            If FailedOffers = "" Then
                                ErrorLogDescription.AppendLine("The Action- Changing status to post validate - was not successful. Following offers were failed during validation:")
                            End If
                            ErrorLogDescription.AppendLine(iOfferId & "  |  " & ErrorMessage)
                            FailedOffers = FailedOffers & iItemID & ","
                            FailedOffersDescription = FailedOffersDescription & ErrorMessage & "|"
                        End If
                    Next
                Next
            End If

            If FailedOffers = "" AndAlso bBeginTransactionRT Then
                MyCommon.QueryStr = "Commit Transaction;"
                MyCommon.LRT_Execute()
                MyCommon.Write_Log(LogFile, "The Action- Changing status to post validate - was successful for all offers.", True)
                Send("OK")
            ElseIf FailedOffers <> "" AndAlso bBeginTransactionRT Then
                MyCommon.QueryStr = "Rollback Transaction;"
                MyCommon.LRT_Execute()
                MyCommon.Write_Log(LogFile, ErrorLogDescription.ToString(), True)
                Send(FailedOffers & ":" & FailedOffersDescription)
            End If
        Catch ex As Exception
            If bBeginTransactionRT Then
                MyCommon.QueryStr = "Rollback Transaction;"
                MyCommon.LRT_Execute()
            End If
            MyCommon.Write_Log(LogFile, ex.ToString(), True)
            Send(ex.ToString)
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try
    End Sub

    Sub WFStatustoReadytoDeploy(ByVal ItemIDs As String, ByVal FromOfferList As Boolean)

        'Dim itemid As Long = 0
        Dim dt As DataTable
        Dim iOfferId As Long
        Dim iItemID As Integer
        Dim aItemIDs As String() = Nothing
        Dim ErrorMessage As String = ""
        Dim FailedOffers As String = ""
        Dim FailedOffersDescription As String = ""
        Dim ErrorLogDescription As New StringBuilder()
        Dim bBeginTransactionRT As Boolean = False

        Response.Cache.SetCacheability(HttpCacheability.NoCache)

        LogFile = "MassActionLog." & MyCommon.Leading_Zero_Fill(Year(Today), 4) & MyCommon.Leading_Zero_Fill(Month(Today), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
        MyCommon.Write_Log(LogFile, "----------------------------------------------------------------")

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            aItemIDs = ItemIDs.Split(",")
            If FromOfferList Then
                MyCommon.Write_Log(LogFile, "Action - Changing status to Ready to Deploy - Performed on " & aItemIDs.Length & " offers from offers list.", True)
                MyCommon.QueryStr = "Begin Transaction;"
                MyCommon.LRT_Execute()
                bBeginTransactionRT = True
                For i = 0 To aItemIDs.GetUpperBound(0)
                    iOfferId = aItemIDs(i)
                    ErrorMessage = ValidateForReadytoDeploy(MyCommon, iOfferId)
                    If (ErrorMessage = "") Then
                        AssignReadytoDeploy(MyCommon, iOfferId)
                    Else
                        If FailedOffers = "" Then
                            ErrorLogDescription.AppendLine("The Action- Changing status to Ready to Deploy - was not successful. Following offers were failed during validation:")
                        End If
                        ErrorLogDescription.AppendLine(iOfferId & "  |  " & ErrorMessage)
                        FailedOffers = FailedOffers & iOfferId & ","
                        FailedOffersDescription = FailedOffersDescription & ErrorMessage & "|"
                    End If
                Next
            Else
                MyCommon.Write_Log(LogFile, "Action - Changing status to Ready to Deploy - Performed on " & aItemIDs.Length & " offers from folders.", True)
                MyCommon.QueryStr = "Begin Transaction;"
                MyCommon.LRT_Execute()
                bBeginTransactionRT = True
                For i = 0 To aItemIDs.GetUpperBound(0)
                    iItemID = aItemIDs(i)
                    MyCommon.QueryStr = "select LinkID from FolderItems with (NoLock) where ItemID=" & iItemID
                    MyCommon.LRT_Execute()
                    dt = MyCommon.LRT_Select()
                    For Each row In dt.Rows
                        iOfferId = MyCommon.NZ(row.Item("LinkID"), 0)
                        ErrorMessage = ValidateForReadytoDeploy(MyCommon, iOfferId)
                        If (ErrorMessage = "") Then
                            AssignReadytoDeploy(MyCommon, iOfferId)
                        Else
                            If FailedOffers = "" Then
                                ErrorLogDescription.AppendLine("The Action- Changing status to Ready to Deploy - was not successful. Following offers were failed during validation:")
                            End If
                            ErrorLogDescription.AppendLine(iOfferId & "  |  " & ErrorMessage)
                            FailedOffers = FailedOffers & iItemID & ","
                            FailedOffersDescription = FailedOffersDescription & ErrorMessage & "|"
                        End If
                    Next
                Next
            End If

            If FailedOffers = "" AndAlso bBeginTransactionRT Then
                MyCommon.QueryStr = "Commit Transaction;"
                MyCommon.LRT_Execute()
                MyCommon.Write_Log(LogFile, "The Action- Changing status to Ready to Deploy - was successful for all offers.", True)
                Send("OK")
            ElseIf FailedOffers <> "" AndAlso bBeginTransactionRT Then
                MyCommon.QueryStr = "Rollback Transaction;"
                MyCommon.LRT_Execute()
                MyCommon.Write_Log(LogFile, ErrorLogDescription.ToString(), True)
                Send(FailedOffers & ":" & FailedOffersDescription)
            End If
        Catch ex As Exception
            If bBeginTransactionRT Then
                MyCommon.QueryStr = "Rollback Transaction;"
                MyCommon.LRT_Execute()
            End If
            MyCommon.Write_Log(LogFile, ex.ToString(), True)
            Send(ex.ToString)
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try
    End Sub

    Sub DuplicateOffersToFolder(ByVal FolderIDs As String, ByVal ItemIDs As String, ByVal FromOfferList As Boolean, ByVal ActionItem As Integer, ByRef strStatusDupOffers As String)
        'Dim itemid As Long = 0
        Dim LinkCount As Integer = 0
        Dim ItemCount As Integer = 0
        Dim NewOfferIDs As String = ""
        Dim iOfferId As Long
        Dim SourceOfferID As Long
        Dim aItemIDs As String() = Nothing
        Dim sNewOfferIDs As String() = Nothing
        Dim dtSourceOffer As DataTable
        Dim rst As DataTable
        Dim LogDescription As New StringBuilder()
        Dim bCopyInboundCrmEngineID As Boolean = True
        Dim offerIds As DataTable = Nothing
        Dim tempOffers As String = Nothing
        Dim nonTempOffers As String = Nothing
        Dim bEnableCopyOffer As Boolean = IIf(MyCommon.Fetch_SystemOption(286) = "1", True, False)
        Dim drarray() As DataRow

        ' Response.Cache.SetCacheability(HttpCacheability.NoCache)
        LogFile = "MassActionLog." & MyCommon.Leading_Zero_Fill(Year(Today), 4) & MyCommon.Leading_Zero_Fill(Month(Today), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
        MyCommon.Write_Log(LogFile, "----------------------------------------------------------------")

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

            If MyCommon.Fetch_CM_SystemOption(107) = "1" Then
                bCopyInboundCrmEngineID = True
            Else
                bCopyInboundCrmEngineID = False
            End If
            If ItemIDs IsNot Nothing AndAlso ItemIDs.Trim <> String.Empty AndAlso ActionItem = 1 Then

                If bEnableCopyOffer Then
                    If (Not Logix.UserRoles.CreateOfferFromBlank) AndAlso ((Logix.UserRoles.CopyOfferCreatedFromBlank AndAlso Not Logix.UserRoles.CopyOfferCreatedFromTemplate) OrElse (Not Logix.UserRoles.CopyOfferCreatedFromBlank AndAlso Logix.UserRoles.CopyOfferCreatedFromTemplate)) Then
                        MyCommon.QueryStr = "SELECT IncentiveID, CI.FromTemplate AS FromTemplate, FI.ItemID AS ItemID FROM CPE_Incentives CI with (NoLock) " &
                                            " INNER JOIN FolderItems FI with (NoLock) ON  FI.LinkID = CI.IncentiveID  WHERE FI.ItemID in (SELECT items FROM Split (@ItemIDs, ','))"
                        
                        MyCommon.DBParameters.Add("@ItemIDs", SqlDbType.NVarChar).Value = ItemIDs
                        offerIds = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                        
                        If (offerIds IsNot Nothing OrElse offerIds.Rows.Count > 0) Then
                            If Logix.UserRoles.CopyOfferCreatedFromTemplate Then
                                ItemIDs = Nothing
                                drarray = offerIds.Select("FromTemplate=1")
                                For i = 0 To (drarray.Length - 1)
                                    ItemIDs = ItemIDs & MyCommon.NZ(drarray(i)("ItemID"), -1) & ","
                                Next
                                If (Not String.IsNullOrEmpty(ItemIDs)) Then
                                    ItemIDs = ItemIDs.Trim().Remove(ItemIDs.Length - 1)
                                End If
                                drarray = offerIds.Select("FromTemplate=0")
                                For i = 0 To (drarray.Length - 1)
                                    nonTempOffers = nonTempOffers & MyCommon.NZ(drarray(i)("IncentiveID"), -1) & ","
                                Next
                                If (Not String.IsNullOrEmpty(nonTempOffers)) Then
                                    nonTempOffers = nonTempOffers.Trim().Remove(nonTempOffers.Length - 1)
                                End If
                            ElseIf Logix.UserRoles.CopyOfferCreatedFromBlank Then
                                ItemIDs = Nothing
                                drarray = offerIds.Select("FromTemplate=0")
                                For i = 0 To (drarray.Length - 1)
                                    ItemIDs = ItemIDs & MyCommon.NZ(drarray(i)("ItemID"), -1) & ","
                                Next
                                If (Not String.IsNullOrEmpty(ItemIDs)) Then
                                    ItemIDs = ItemIDs.Trim().Remove(ItemIDs.Length - 1)
                                End If
                                drarray = offerIds.Select("FromTemplate=1")
                                For i = 0 To (drarray.Length - 1)
                                    tempOffers = tempOffers & MyCommon.NZ(drarray(i)("IncentiveID"), -1) & ","
                                Next
                                If (Not String.IsNullOrEmpty(tempOffers)) Then
                                    tempOffers = tempOffers.Trim().Remove(tempOffers.Length - 1)
                                End If
                            End If
                        End If
                    End If
                End If
                If (Not String.IsNullOrEmpty(ItemIDs)) Then
                    ItemCount = IndexOfCount(ItemIDs, ","c) + 1
                    MyCommon.Write_Log(LogFile, "Perfomed Action - Duplicate Offers to folder - on " & ItemCount & " offers", True)
                    MyCommon.QueryStr = "dbo.pa_Duplicate_Offers"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@FromOfferList", SqlDbType.NVarChar).Value = FromOfferList
                    MyCommon.LRTsp.Parameters.Add("@ItemIDs", SqlDbType.NVarChar).Value = ItemIDs
                    MyCommon.LRTsp.Parameters.Add("@CopyInboundCRM", SqlDbType.Bit).Value = IIf(bCopyInboundCrmEngineID, 1, 0)
                    MyCommon.LRTsp.Parameters.Add("@NewOfferIDs", SqlDbType.NVarChar, -1).Direction = ParameterDirection.Output
                    MyCommon.LRTsp.ExecuteNonQuery()
                    'ResponseText = LinkCount & " item" & IIf(LinkCount = 1, "", "s") & " removed from selected folder."
                    NewOfferIDs = MyCommon.LRTsp.Parameters("@NewOfferIDs").Value.ToString()
                    MyCommon.Close_LRTsp()
                    Dim BannersEnabled As Boolean = MyCommon.Fetch_SystemOption(66)
                    Dim offersDT As DataTable = New DataTable()
                    Dim offers As String() = NewOfferIDs.Split(", ")
                    offersDT.Columns.Add("OfferID")
                    For Each offer In offers
                        offersDT.Rows.Add(Long.Parse(offer))
                    Next
                    MyCommon.QueryStr = "dbo.pt_InsertOfferApprovalRecord"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@BannersEnabled", SqlDbType.Bit).Value = BannersEnabled
                    MyCommon.LRTsp.Parameters.Add("@AdminUserId", SqlDbType.Int).Value = AdminUserID
                    MyCommon.LRTsp.Parameters.Add("@OfferDT", SqlDbType.Structured).Value = offersDT
                    MyCommon.LRTsp.ExecuteNonQuery()

                Else
                    If bEnableCopyOffer Then
                        If (Not Logix.UserRoles.CreateOfferFromBlank) AndAlso ((Logix.UserRoles.CopyOfferCreatedFromBlank AndAlso Not Logix.UserRoles.CopyOfferCreatedFromTemplate) OrElse (Not Logix.UserRoles.CopyOfferCreatedFromBlank AndAlso Logix.UserRoles.CopyOfferCreatedFromTemplate)) Then
                            If Logix.UserRoles.CopyOfferCreatedFromBlank Then
                                strStatusDupOffers = "Offers created from template: " & tempOffers & " are not copied as user does not have permission to copy offers created from template."
                            ElseIf (Logix.UserRoles.CopyOfferCreatedFromTemplate) Then
                                strStatusDupOffers = "Offers created from blank: " & nonTempOffers & " are not copied as user does not have permission to copy offers created from blank."
                            End If
                            MyCommon.Write_Log(LogFile, strStatusDupOffers.ToString(), True)
                        End If
                    End If
                End If
            End If
            If (NewOfferIDs IsNot Nothing AndAlso NewOfferIDs.Trim <> String.Empty) OrElse ActionItem = 7 Then
                If (NewOfferIDs IsNot Nothing AndAlso NewOfferIDs.Trim <> String.Empty AndAlso ActionItem = 1) Then
                    LinkCount = IndexOfCount(NewOfferIDs, ","c) + 1
                End If
                MyCommon.QueryStr = "dbo.pa_FolderOffers_Insert"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@OfferIDs", SqlDbType.NVarChar).Value = IIf(ActionItem = 7, ItemIDs, NewOfferIDs)
                MyCommon.LRTsp.Parameters.Add("@FolderIDs", SqlDbType.NVarChar).Value = FolderIDs
                MyCommon.LRTsp.ExecuteNonQuery()
                MyCommon.Close_LRTsp()
                If bEnableCopyOffer Then
                    If (Not Logix.UserRoles.CreateOfferFromBlank) AndAlso ((Logix.UserRoles.CopyOfferCreatedFromBlank AndAlso Not Logix.UserRoles.CopyOfferCreatedFromTemplate) OrElse (Not Logix.UserRoles.CopyOfferCreatedFromBlank AndAlso Logix.UserRoles.CopyOfferCreatedFromTemplate)) Then
                        If Logix.UserRoles.CopyOfferCreatedFromBlank Then
                            LogDescription.AppendLine("Offers created from blank are copied successfully to the folder but offers created from template with offerIds: " & tempOffers & " are not copied as user has no permission to copy offers created from template.")
                            LogDescription.AppendLine("Action Successful. Following are the new offers added to folder(s) :" & IIf(ActionItem = 7, ItemIDs, NewOfferIDs) & ":")
                        ElseIf Logix.UserRoles.CopyOfferCreatedFromTemplate Then
                            LogDescription.AppendLine("Offers created from template are copied successfully to the folder but offers created from blank with offerIDs: " & nonTempOffers & " are not copied as user has no permission to copy offers created from blank.")
                            LogDescription.AppendLine("Action Successful. Following are the new offers added to folder(s) :" & IIf(ActionItem = 7, ItemIDs, NewOfferIDs) & ":")
                        End If
                    Else
                        LogDescription.AppendLine("Action Successful. Following are the new offers added to folder(s) :" & IIf(ActionItem = 7, ItemIDs, NewOfferIDs) & ":")
                    End If
                Else
                    LogDescription.AppendLine("Action Successful. Following are the new offers added to folder(s) :" & IIf(ActionItem = 7, ItemIDs, NewOfferIDs) & ":")
                End If
            End If

            If (Not String.IsNullOrEmpty(ItemIDs)) Then
                If ItemCount = LinkCount Then
                    aItemIDs = ItemIDs.Split(",")
                    sNewOfferIDs = NewOfferIDs.Split(",")
                    If ActionItem <> 7 Then
                        For i = 0 To aItemIDs.GetUpperBound(0)
                            iOfferId = sNewOfferIDs(i)
                            SourceOfferID = aItemIDs(i)
                            If FromOfferList = False Then
                                MyCommon.QueryStr = "Select linkid from FolderItems with (NoLock) where ItemId = @SourceOfferID"
                                
                                MyCommon.DBParameters.Add("@SourceOfferID", SqlDbType.Int).Value = SourceOfferID
                                dtSourceOffer = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                                
                                If dtSourceOffer.Rows.Count > 0 Then
                                    SourceOfferID = dtSourceOffer.Rows(0)(0)
                                End If
                            End If
                            MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(),CreatedByAdminID=" & AdminUserID & ", LastUpdatedByAdminID =" & AdminUserID & ", StatusFlag=1 where IncentiveID =" & iOfferId
                            MyCommon.LRT_Execute()
                            MyCommon.Activity_Log(3, iOfferId, AdminUserID, Copient.PhraseLib.Lookup("term.duplicatedoffer", LanguageID) & ": " & SourceOfferID)
                            MyCommon.QueryStr = "select OfferId FROM Offers WHERE OfferID=" & iOfferId
                            rst = MyCommon.LRT_Select
                            If rst.Rows.Count > 0 Then
                                CreateNewLocalPromotionVariables(iOfferId, MyCommon)
                            End If
                        Next i
                    End If
                    MyCommon.Write_Log(LogFile, LogDescription.ToString(), True)
                    If ActionItem = 1 Then
                        If bEnableCopyOffer Then
                            If (Not Logix.UserRoles.CreateOfferFromBlank) AndAlso ((Logix.UserRoles.CopyOfferCreatedFromBlank AndAlso Not Logix.UserRoles.CopyOfferCreatedFromTemplate) OrElse (Not Logix.UserRoles.CopyOfferCreatedFromBlank AndAlso Logix.UserRoles.CopyOfferCreatedFromTemplate)) Then
                                If Logix.UserRoles.CopyOfferCreatedFromBlank Then
                                    strStatusDupOffers = "Successfully copied offers created from blank. Offers created from template: " & tempOffers & " are not copied as user does not have permission to copy offers created from template."
                                ElseIf (Logix.UserRoles.CopyOfferCreatedFromTemplate) Then
                                    strStatusDupOffers = "Successfully copied offers created from template. Offers created from blank: " & nonTempOffers & " are not copied as user does not have permission to copy offers created from blank."
                                End If
                            Else
                                strStatusDupOffers = Copient.PhraseLib.Lookup("term.success", LanguageID) + ": " + Copient.PhraseLib.Lookup("offer-dup-success", LanguageID)
                            End If
                        Else
                            strStatusDupOffers = Copient.PhraseLib.Lookup("term.success", LanguageID) + ": " + Copient.PhraseLib.Lookup("offer-dup-success", LanguageID)
                        End If
                    End If
                Else
                    If ActionItem = 1 Then
                        strStatusDupOffers = Copient.PhraseLib.Lookup("offer-dup-failure", LanguageID)
                    End If
                End If
            End If
        Catch ex As Exception
            MyCommon.Write_Log(LogFile, ex.ToString(), True)
            If ex.Message = "error.couldnot-processoffers" Then
                strStatusDupOffers = Copient.PhraseLib.Lookup(ex.Message, LanguageID)
            Else
                strStatusDupOffers = ex.ToString()
            End If
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try
    End Sub
    Sub CreateNewLocalPromotionVariables(ByVal lOfferId As Long, ByRef Mycommon As Copient.CommonInc)
        Dim lRewardID As Long
        Dim lPromoVarId As Long
        Dim rst As DataTable
        Dim row As DataRow

        ' create local promotion variables for this new offer
        Mycommon.QueryStr = "select OfferID from Offers with (NoLock) where OfferID=" & lOfferId &
                            " and DistPeriodLimit > 0.00 and DistPeriod <> 0 and DistPeriodVarID=0;"
        rst = Mycommon.LRT_Select
        For Each row In rst.Rows
            Mycommon.Open_LogixXS()
            Mycommon.QueryStr = "dbo.pc_DistributionVar_Create"
            Mycommon.Open_LXSsp()
            Mycommon.LXSsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = lOfferId
            Mycommon.LXSsp.Parameters.Add("@VarID", SqlDbType.BigInt).Direction = ParameterDirection.Output
            Mycommon.LXSsp.ExecuteNonQuery()
            lPromoVarId = Mycommon.LXSsp.Parameters("@VarID").Value
            Mycommon.Close_LXSsp()
            Mycommon.Close_LogixXS()
            Mycommon.QueryStr = "update Offers with (RowLock) set DistPeriodVarID=" & lPromoVarId & " where OfferID=" & lOfferId & ";"
            Mycommon.LRT_Execute()
        Next

        ' create local promotion variables for this new offer's rewards
        Mycommon.QueryStr = "select RewardID from OfferRewards with (NoLock) where OfferID=" & lOfferId &
                            " and RewardLimit > 0.00 and RewardDistPeriod <> 0 and RewardDistLimitVarID=0;"
        rst = Mycommon.LRT_Select
        For Each row In rst.Rows
            lRewardID = Mycommon.NZ(row.Item("RewardID"), 0)
            Mycommon.Open_LogixXS()
            Mycommon.QueryStr = "dbo.pc_RewardLimitVar_Create"
            Mycommon.Open_LXSsp()
            Mycommon.LXSsp.Parameters.Add("@RewardID", SqlDbType.BigInt).Value = lRewardID
            Mycommon.LXSsp.Parameters.Add("@VarID", SqlDbType.BigInt).Direction = ParameterDirection.Output
            Mycommon.LXSsp.ExecuteNonQuery()
            lPromoVarId = Mycommon.LXSsp.Parameters("@VarID").Value
            Mycommon.Close_LXSsp()
            Mycommon.Close_LogixXS()
            Mycommon.QueryStr = "update OfferRewards with (RowLock) set RewardDistLimitVarID=" & lPromoVarId & " where RewardID=" & lRewardID & ";"
            Mycommon.LRT_Execute()
        Next
    End Sub
    Function ValidateOfferToAssignFolder(ByVal LinkIDs As String) As Boolean
        Dim aItemIDs As String() = Nothing
        Dim dt As DataTable
        aItemIDs = LinkIDs.Split(",")
        For i = 0 To aItemIDs.GetUpperBound(0)
            MyCommon.QueryStr = "Select FolderID from FolderItems where LinkID=@aItemID"
            MyCommon.DBParameters.Add("@aItemID", SqlDbType.Int).Value = aItemIDs(i)
            dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            If MyCommon.Fetch_SystemOption(191) = "1" AndAlso dt.Rows.Count > 0 Then
                Send("Cannot add the offer to this folder as the offer is associated to another folder")
                Return False
            End If
        Next
        Return True
    End Function
    ' ********************************************************************************
    ' *** Adds all the specified items not already assigned to the specified FolderID
    ' *** Should only be called for use in the Add mode.
    ' ********************************************************************************
    Sub AddItemsToFolder(ByVal FolderID As Long, ByVal LinkIDs As String, ByVal LinkTypeIDs As String, ByVal AdminUserID As Integer, Optional ByVal WFStatus As Integer = 0)
        Dim ResponseText As String = ""
        Dim LinkCount As Integer = 0

        Response.Cache.SetCacheability(HttpCacheability.NoCache)

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If LinkIDs IsNot Nothing AndAlso LinkIDs.Length > 0 Then
                If Not ValidateOfferToAssignFolder(LinkIDs) Then
                    Exit Sub
                End If
                If LinkIDs.Trim <> String.Empty AndAlso LinkTypeIDs.Trim <> String.Empty Then
                    LinkCount = IndexOfCount(LinkIDs, ","c) + 1
                    MyCommon.QueryStr = "dbo.pa_FolderItem_Insert"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@FolderID", SqlDbType.Int).Value = FolderID
                    MyCommon.LRTsp.Parameters.Add("@LinkIDs", SqlDbType.NVarChar).Value = LinkIDs
                    MyCommon.LRTsp.Parameters.Add("@LinkTypeIDs", SqlDbType.NVarChar).Value = LinkTypeIDs
                    MyCommon.LRTsp.ExecuteNonQuery()
                    ResponseText = Copient.PhraseLib.Detokenize("folders.ItemsAdded", LanguageID, LinkCount)
                    MyCommon.Close_LRTsp()
                    If MyCommon.Fetch_SystemOption(192) = "1" Then
                        MyCommon.QueryStr = "dbo.pt_FolderOffersDates_Default"
                        MyCommon.Open_LRTsp()
                        MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = 2
                        MyCommon.LRTsp.Parameters.Add("@FolderID", SqlDbType.Int).Value = FolderID
                        MyCommon.LRTsp.Parameters.Add("@LinkIDs", SqlDbType.NVarChar).Value = LinkIDs
                        MyCommon.LRTsp.ExecuteNonQuery()
                        MyCommon.Close_LRTsp()
                    End If
                End If
            End If
            SetFolderMassOperationStatus(FolderID, FOLDER_NOT_IN_USE)
            MyCommon.Activity_Log2(45, 22, FolderID, AdminUserID, Copient.PhraseLib.Lookup("history.folder-additem", LanguageID))
            If WFStatus <> 0 Then
                SendFolderItems(FolderID, ResponseText, WFStatus)
            Else
                SendFolderItems(FolderID, ResponseText)
            End If

        Catch ex As Exception
            Send(ex.ToString)
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()

        End Try
    End Sub

    Sub CreateFolder(ByVal ParentFolderID As Integer, ByVal FolderName As String, ByVal AccessLevel As Integer, ByVal FolderStartDate As String, ByVal FolderEndDate As String, ByVal FolderTheme As String, ByVal AdminUserID As Integer, Optional ByVal IsUEDefaultFolder As Boolean = False)
        Dim FolderID As Integer = 0
        Dim foldersdate As Date
        Dim folderedate As Date
        Dim dt As DataTable
        Dim bAllowTimeWithStartEndDates As Boolean = False
        bAllowTimeWithStartEndDates = MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE) AndAlso (MyCommon.Fetch_UE_SystemOption(200) = "1")

        If String.IsNullOrWhiteSpace(FolderName) Then
            Send(Copient.PhraseLib.Lookup("folders.EnterFolderName", LanguageID))
            Exit Sub
        End If
        If isEmpty(FolderStartDate) AndAlso isEmpty(FolderEndDate) Then
            If bAllowTimeWithStartEndDates Then
                Send(Copient.PhraseLib.Lookup("folder.EnterStartDateTime", LanguageID) & " " & Copient.PhraseLib.Lookup("folder.EnterEndDateTime", LanguageID))
            Else
                Send(Copient.PhraseLib.Lookup("folder.EnterStartDate", LanguageID) & " " & Copient.PhraseLib.Lookup("folder.EnterEndDate", LanguageID))
            End If
            Exit Sub
        End If
        If Not isEmpty(FolderStartDate) AndAlso isEmpty(FolderEndDate) Then
            If bAllowTimeWithStartEndDates Then
                Send(Copient.PhraseLib.Lookup("folder.EnterEndDateTime", LanguageID))
            Else
                Send(Copient.PhraseLib.Lookup("folder.EnterEndDate", LanguageID))
            End If
            Exit Sub
        End If
        If Not isEmpty(FolderEndDate) AndAlso isEmpty(FolderStartDate) Then
            If bAllowTimeWithStartEndDates Then
                Send(Copient.PhraseLib.Lookup("folder.EnterStartDateTime", LanguageID))
            Else
                Send(Copient.PhraseLib.Lookup("folder.EnterStartDate", LanguageID))
            End If
            Exit Sub
        End If
        If Not isEmpty(FolderStartDate) Then
            If Not Date.TryParse(FolderStartDate, MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, foldersdate) Then
                If bAllowTimeWithStartEndDates Then
                    Send(Copient.PhraseLib.Lookup("folder.EnterStartDateTime", LanguageID))
                Else
                    Send(Copient.PhraseLib.Lookup("folder.EnterStartDate", LanguageID))
                End If
                Exit Sub
            End If
        End If
        If Not isEmpty(FolderEndDate) Then
            If Not Date.TryParse(FolderEndDate, MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, folderedate) Then
                If bAllowTimeWithStartEndDates Then
                    Send(Copient.PhraseLib.Lookup("folder.EnterStartDateTime", LanguageID))
                Else
                    Send(Copient.PhraseLib.Lookup("folder.EnterStartDate", LanguageID))
                End If
                Exit Sub
            End If
        End If
        If ((Not isEmpty(FolderStartDate)) OrElse (Not isEmpty(FolderEndDate))) Then
            If folderedate < foldersdate Then
                If bAllowTimeWithStartEndDates Then
                    Send(Copient.PhraseLib.Lookup("folders.BadEndDateTime", LanguageID))
                Else
                    Send(Copient.PhraseLib.Lookup("folders.BadEndDate", LanguageID))
                End If
                Exit Sub
            ElseIf folderedate >= CDate("1/1/9999") Then
                Send(Copient.PhraseLib.Lookup("term.datetoolarge", LanguageID))
                Exit Sub
            End If
        End If

        If MyCommon.Fetch_SystemOption(123) = "1" AndAlso (FolderTheme Is Nothing OrElse FolderTheme.Trim = String.Empty) Then
            Send(Copient.PhraseLib.Lookup("folder.EnterTheme", LanguageID))
            Exit Sub
        End If
        Response.Cache.SetCacheability(HttpCacheability.NoCache)

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If ((Not isEmpty(FolderStartDate)) OrElse (Not isEmpty(FolderEndDate))) Then
                If folderedate < foldersdate Then
                    Send(Copient.PhraseLib.Lookup("folders.BadEndDate", LanguageID))
                    Exit Sub
                ElseIf folderedate >= CDate("1/1/9999") Then
                    Send(Copient.PhraseLib.Lookup("term.datetoolarge", LanguageID))
                    Exit Sub
                End If
            End If
            MyCommon.QueryStr = "Select FolderID from Folders where FolderName= @FolderName and ParentFolderID= @ParentFolderID"
            MyCommon.DBParameters.Add("@FolderName", SqlDbType.NVarChar, 50).Value = FolderName
            MyCommon.DBParameters.Add("@ParentFolderID", SqlDbType.Int).Value = ParentFolderID
            dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            If dt.Rows.Count > 0 Then
                Send(Copient.PhraseLib.Lookup("folders.NameExists", LanguageID))
                Exit Sub
            End If

            'AL-6583: Default Folder
            If IsUEDefaultFolder = True Then
                MyCommon.QueryStr = "Select * from Folders where DefaultUEFolder=1"
                dt = MyCommon.LRT_Select()
                If dt.Rows.Count > 0 Then
                    Send("NO|" + Copient.PhraseLib.Lookup("term.DefaultfolderExists", LanguageID))
                    Exit Sub
                End If
            End If

            MyCommon.QueryStr = "dbo.pt_Folders_Insert"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@ParentFolderID", SqlDbType.Int).Value = IIf(ParentFolderID > 0, ParentFolderID, 0)
            MyCommon.LRTsp.Parameters.Add("@FolderName", SqlDbType.NVarChar, 50).Value = FolderName
            MyCommon.LRTsp.Parameters.Add("@AccessLevel", SqlDbType.Int).Value = AccessLevel
            MyCommon.LRTsp.Parameters.Add("@CreatedBy", SqlDbType.Int).Value = AdminUserID
            MyCommon.LRTsp.Parameters.Add("@FolderStartDate", SqlDbType.DateTime).Value = IIf(Not isEmpty(FolderStartDate), foldersdate, DBNull.Value)
            MyCommon.LRTsp.Parameters.Add("@FolderEndDate", SqlDbType.DateTime).Value = IIf(Not isEmpty(FolderEndDate), folderedate, DBNull.Value)
            MyCommon.LRTsp.Parameters.Add("@FolderID", SqlDbType.Int).Direction = ParameterDirection.Output
            MyCommon.LRTsp.Parameters.Add("@IsUEDefaultFolder", SqlDbType.Bit).Value = IsUEDefaultFolder
            MyCommon.LRTsp.ExecuteNonQuery()
            FolderID = MyCommon.LRTsp.Parameters("@FolderID").Value
            MyCommon.Close_LRTsp()
            If MyCommon.Fetch_SystemOption(123) = "1" Then
                MyCommon.QueryStr = "pt_FolderThemes_Insert"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@FolderID", SqlDbType.Int).Value = IIf(FolderID > 0, FolderID, 0)
                MyCommon.LRTsp.Parameters.Add("@Theme", SqlDbType.NVarChar, 20).Value = IIf(Not isEmpty(FolderTheme), FolderTheme, DBNull.Value)
                MyCommon.LRTsp.ExecuteNonQuery()
                MyCommon.Close_LRTsp()
            End If
            MyCommon.Activity_Log2(45, 1, FolderID, AdminUserID, Copient.PhraseLib.Lookup("history.folder-create", LanguageID))
            Sendb("OK|" & FolderID & "|" & IsUEDefaultFolder)
        Catch ex As Exception
            Send("NO|" & ex.ToString)
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()

        End Try
    End Sub

    Sub RenameFolder(ByVal FolderID As Integer, ByVal NewFolderName As String, ByVal AdminUserID As Integer)

        Response.Cache.SetCacheability(HttpCacheability.NoCache)

        Dim dt As DataTable

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

            MyCommon.QueryStr = "Select ParentFolderID,FolderName from Folders where FolderID=" & FolderID
            dt = MyCommon.LRT_Select()

            If dt.Rows.Count > 0 AndAlso dt.Rows(0).Item(1) <> NewFolderName Then
                MyCommon.QueryStr = "Select FolderID from Folders where FolderName= @FolderName"
                MyCommon.DBParameters.Add("@FolderName", SqlDbType.NVarChar, 100).Value = NewFolderName.ToString()
                MyCommon.DBParameters.Add("@FolderID", SqlDbType.Int).Value = FolderID
                dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                If dt.Rows.Count > 0 Then
                    Send("NO " & Copient.PhraseLib.Lookup("folders.NameExists", LanguageID))
                    Exit Sub
                End If
            End If
            MyCommon.DBParameters.Clear()
            MyCommon.QueryStr = "update Folders with (RowLock) set FolderName= @FolderName where FolderID= @FolderID "
            MyCommon.DBParameters.Add("@FolderName", SqlDbType.NVarChar, 100).Value = NewFolderName.ToString()
            MyCommon.DBParameters.Add("@FolderID", SqlDbType.Int).Value = FolderID
            MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)

            ' MyCommon.Activity_Log2(45, 3, FolderID, AdminUserID, Copient.PhraseLib.Lookup("history.folder-rename", LanguageID))
            ' Sendb("OK")
        Catch ex As Exception
            Send("NO" & ex.ToString)
            'Finally
            '    If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()

        End Try
    End Sub

    Sub DeleteFolder(ByVal FolderID As Integer, ByVal AdminUserID As Integer)
        Dim dt As DataTable
        Response.Cache.SetCacheability(HttpCacheability.NoCache)

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.Fetch_SystemOption(134) = "1" Then
                MyCommon.QueryStr = "select incentiveid from cpe_incentives with (NoLock) where incentiveid in " &
                                     " (select LinkID from folderitems with (NoLock) where folderid=" & FolderID & " Or folderid in " &
                                      " (select FolderID from Folders with (NoLock) where ParentFolderID=" & FolderID & ")) and deleted = 0"
                If MyCommon.LRT_Select.Rows.Count > 0 Then
                    Sendb(Copient.PhraseLib.Lookup("folders.CannotDelete", LanguageID))
                    Exit Sub
                Else
                    DeleteFolderAndContents(FolderID)
                End If
            Else
                DeleteFolderAndContents(FolderID)
            End If

            MyCommon.Activity_Log2(45, 2, FolderID, AdminUserID, Copient.PhraseLib.Lookup("history.folder-delete", LanguageID))
            'AL-6583: Default Folder
            MyCommon.QueryStr = "Select * from Folders where DefaultUEFolder=1 "
            dt = MyCommon.LRT_Select()
            If dt.Rows.Count > 0 Then
                Send("OK|" + "1")
            Else
                Send("OK|" + "0")
            End If
        Catch ex As Exception
            Send("NO|" & ex.ToString)
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try
    End Sub

    Private Sub DeleteFolderAndContents(ByVal FolderID As Integer)
        Dim dt As DataTable
        Dim row As DataRow

        ' delete the folders with themes also only if system option 123 is true
        If MyCommon.Fetch_SystemOption(123) = "1" Then
            MyCommon.QueryStr = "delete from FolderThemes with (RowLock) where FolderID=" & FolderID
            MyCommon.LRT_Execute()
        End If
        ' delete the specified folder
        MyCommon.QueryStr = "delete from Folders with (RowLock) where FolderID=" & FolderID
        MyCommon.LRT_Execute()

        ' delete all folder content
        MyCommon.QueryStr = "delete from FolderItems with (RowLock) where FolderID=" & FolderID
        MyCommon.LRT_Execute()

        ' find all the subfolders for the specified folder and delete them as well
        MyCommon.QueryStr = "select FolderID from Folders with (NoLock) where ParentFolderID=" & FolderID
        dt = MyCommon.LRT_Select
        For Each row In dt.Rows
            DeleteFolderAndContents(MyCommon.NZ(row.Item("FolderID"), 0))
        Next
    End Sub

    Sub SendDivForModFolder(ByVal FolderID As Integer, ByVal AdminUserID As Integer)
        Dim dt As DataTable
        Dim row As DataRow
        Dim FolderStartDate, FolderEndDate As Date
        Dim sFolderStartDate As String = ""
        Dim sFolderEndDate As String = ""
        Dim sFolderName As String = ""
        Dim TempBuf As New StringBuilder()
        Dim index As Integer = 0
        Dim ThemeDescription As String = ""
        Dim bRTConnectionOpened As Boolean = False
        Dim bAllowTimeWithStartEndDates As Boolean = False
        Dim FolderStartHr As String = ""
        Dim FolderStartMin As String = ""
        Dim FolderEndHr As String = ""
        Dim FolderEndMin As String = ""

        bAllowTimeWithStartEndDates = MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE) AndAlso (MyCommon.Fetch_UE_SystemOption(200) = "1")

        Response.Cache.SetCacheability(HttpCacheability.NoCache)

        Try

            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                MyCommon.Open_LogixRT()
                bRTConnectionOpened = True
            End If

            MyCommon.QueryStr = "select StartDate,EndDate,FolderName from Folders with (NoLock) where FolderID=" & FolderID
            dt = MyCommon.LRT_Select
            If dt.Rows.Count > 0 Then
                If (Not IsDBNull(dt.Rows(0).Item("StartDate"))) Then
                    FolderStartDate = dt.Rows(0).Item("StartDate")
                    sFolderStartDate = Logix.ToShortDateString(FolderStartDate, MyCommon)
                    If sFolderStartDate = "01/01/1900" Then sFolderStartDate = ""

                    If bAllowTimeWithStartEndDates Then
                        FolderStartHr = FolderStartDate.Hour.ToString()
                        If FolderStartHr.Length <= 1 Then FolderStartHr = "0" & FolderStartHr
                        FolderStartMin = FolderStartDate.Minute.ToString()
                        If FolderStartMin.Length <= 1 Then FolderStartMin = "0" & FolderStartMin
                    End If
                End If
                If (Not IsDBNull(dt.Rows(0).Item("EndDate"))) Then
                    FolderEndDate = dt.Rows(0).Item("EndDate")
                    sFolderEndDate = Logix.ToShortDateString(FolderEndDate, MyCommon)
                    If sFolderEndDate = "01/01/1900" Then sFolderEndDate = ""
                    If bAllowTimeWithStartEndDates Then
                        FolderEndHr = FolderEndDate.Hour.ToString()
                        If FolderEndHr.Length <= 1 Then FolderEndHr = "0" & FolderEndHr
                        FolderEndMin = FolderEndDate.Minute.ToString()
                        If FolderEndMin.Length <= 1 Then FolderEndMin = "0" & FolderEndMin
                    End If
                End If
                If (Not IsDBNull(dt.Rows(0).Item("FolderName"))) Then
                    sFolderName = dt.Rows(0).Item("FolderName").ToString()
                    sFolderName = sFolderName.Replace("""", "&quot;")
                End If
            End If

            TempBuf.AppendLine("<div class=""foldertitlebar"">")
            TempBuf.AppendLine(" <span class=""dialogtitle"">" & Copient.PhraseLib.Lookup("folders.RenameFolder", LanguageID) & "</span>")
            TempBuf.AppendLine(" <span id=""dialogclose"" class=""dialogclose"" onclick=""toggleDialog('modifyfolder', false);"">X</span>")
            TempBuf.AppendLine("</div>")
            TempBuf.AppendLine("<div id=""modifyfoldererror"" style=""color: red;"" >")
            TempBuf.AppendLine("</div> ")
            TempBuf.AppendLine("<div class=""dialogcontents"">")
            TempBuf.AppendLine(" <br class=""half"" />")
            TempBuf.AppendLine("<label for=""EditFolderName"">" & Copient.PhraseLib.Lookup("term.foldername", LanguageID) & ":" & "</label><br />")
            TempBuf.AppendLine("<input type=""text"" id=""editFolderName"" name=""editFolderName"" value=""" & sFolderName & """ class=""mediumlong"" maxlength=""50"" />")
            TempBuf.AppendLine(" <br class=""half"" />")
            TempBuf.AppendLine("<label for=""lblfoldermodify"">" & Copient.PhraseLib.Lookup("folders.FolderDate", LanguageID) & ":" & "</label><br />")
            TempBuf.AppendLine("<input type=""text"" class=""short"" id=""modifyfolderstart"" name=""modifyfolderstart"" value=""" & sFolderStartDate & """ />")
            TempBuf.AppendLine("<img src=""../images/calendar.png"" class=""calendar"" id=""foldermodified-start-picker"" alt=""" & Copient.PhraseLib.Lookup("term.datepicker", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.datepicker", LanguageID) & """ onclick=""displayDatePicker('modifyfolderstart',event);"" />")
            If bAllowTimeWithStartEndDates Then
                TempBuf.AppendLine("<input type=""text"" class=""shortest"" id=""modifyfolderstartHr"" maxlength=""2"" onkeypress=""return NumberChecking(event,false,false)"" name=""modifyfolderstartHr"" value=""" & FolderStartHr & """ /> :")
                TempBuf.AppendLine("<input type=""text"" class=""shortest"" id=""modifyfolderstartMin"" maxlength=""2"" onkeypress=""return NumberChecking(event,false,false)"" name=""modifyfolderstartMin"" value=""" & FolderStartMin & """ />")
            End If
            TempBuf.AppendLine(Copient.PhraseLib.Lookup("term.to", LanguageID))
            TempBuf.AppendLine("<input type=""text"" class=""short"" id=""modifyfolderend"" name=""modifyfolderend"" value=""" & sFolderEndDate & """ />")
            TempBuf.AppendLine("<img src=""../images/calendar.png"" class=""calendar"" id=""foldermodified-end-picker"" alt=""" & Copient.PhraseLib.Lookup("term.datepicker", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.datepicker", LanguageID) & """ onclick=""displayDatePicker('modifyfolderend',event);"" />")
            If bAllowTimeWithStartEndDates Then
                TempBuf.AppendLine("<input type=""text"" class=""shortest"" id=""modifyfolderEndHr"" maxlength=""2"" onkeypress=""return NumberChecking(event,false,false)"" name=""modifyfolderEndHr"" value=""" & FolderEndHr & """ /> :")
                TempBuf.AppendLine("<input type=""text"" class=""shortest"" id=""modifyfolderEndMin"" maxlength=""2"" onkeypress=""return NumberChecking(event,false,false)"" name=""modifyfolderEndMin"" value=""" & FolderEndMin & """ />")
            End If

            TempBuf.AppendLine(MyCommon.GetAdminUser.Culture.DateTimeFormat.ShortDatePattern)

            TempBuf.AppendLine("<br />")

            MyCommon.QueryStr = "select TH.ThemeDescription from FolderThemes FT with (NoLock) inner join Themes TH on FT.ThemeID=TH.ThemeID where FT.FolderID=" & FolderID
            dt = MyCommon.LRT_Select
            If dt.Rows.Count > 0 Then
                ThemeDescription = dt.Rows(0).Item("ThemeDescription")
            End If

            MyCommon.QueryStr = "select ThemeId,ThemeDescription from Themes with (NoLock) where themedescription not in ('" & ThemeDescription & "');"
            dt = MyCommon.LRT_Select
            If dt.Rows.Count > 0 Then
                TempBuf.AppendLine(" <table summary=" & Copient.PhraseLib.Lookup("term.dates", LanguageID) & " >")
                TempBuf.AppendLine(" <tr>")
                TempBuf.AppendLine("  <td>")
                TempBuf.AppendLine(" <label for=""lblModifyTheme"">Theme:</label><br />")
                TempBuf.AppendLine(" <select name=""ModifyTheme"" id=""ModifyTheme"" >")
                TempBuf.AppendLine(" <option value=""" & index & """ selected=""selected"">" & ThemeDescription & "</option>")

                For Each row In dt.Rows
                    index = index + 1
                    TempBuf.AppendLine("<option value=""" & index & """>" & MyCommon.NZ(row.Item("ThemeDescription"), "") & "</option>")
                Next
            End If
            TempBuf.AppendLine("</select>")
            TempBuf.AppendLine("  </td>")
            TempBuf.AppendLine("</tr>")
            TempBuf.AppendLine("<br />")
            If MyCommon.IsEngineInstalled(9) Then
                MyCommon.QueryStr = "Select FolderID from Folders where DefaultUEFolder = 1"
                dt = MyCommon.LRT_Select
                Dim IsthisSelected As Boolean = False
                If dt.Rows.Count > 0 Then
                    IsthisSelected = MyCommon.NZ(dt.Rows(0)("FolderID"), False)
                End If
                If dt.Rows.Count = 0 Or (dt.Rows.Count = 1 AndAlso MyCommon.NZ(dt.Rows(0)("FolderID"), -1) = FolderID) Then
                    TempBuf.AppendLine("<tr>")
                    TempBuf.AppendLine("<td>")
                    TempBuf.AppendLine("<input type=""checkbox"" name=""defaultUEFolder-Mod"" id=""defaultUEFolder-Mod"" " & If(IsthisSelected, "checked=""checked""", "") & """ />")
                    TempBuf.AppendLine("<label for=""defaultUEFolder-Mod"">" & Copient.PhraseLib.Lookup("term.defaultuefolder", LanguageID) & "</label> <br /> <br />")
                    TempBuf.AppendLine("</td>")
                    TempBuf.AppendLine("</tr>")
                End If
            End If
            TempBuf.AppendLine("<tr>")
            TempBuf.AppendLine("<td>")
            TempBuf.AppendLine("<input type=""button"" name=""btnModFolder"" id=""btnModFolder"" value=""" & Copient.PhraseLib.Lookup("term.modify", LanguageID) & """ onclick=""javascript:savemodifiedFolder();"" />")
            TempBuf.AppendLine("</td>")
            TempBuf.AppendLine("</tr>")
            TempBuf.AppendLine(" </table>")

            TempBuf.AppendLine("<div id=""divMessage"" align=""Center"" style=""display:none""><label for=""lblErrorMsg"" style=""color:red;font-weight: bold"">" & Copient.PhraseLib.Lookup("folders.ConfirmStartEndDateChange", LanguageID) & "</label>")
            TempBuf.AppendLine(" <table>")
            TempBuf.AppendLine(" <tr>")
            TempBuf.AppendLine("  <td>")
            TempBuf.AppendLine("<input type=""button"" name=""btnOk"" id=""btnOk"" value=" & Copient.PhraseLib.Lookup("term.ok", LanguageID) & " onclick=""javascript:OKClicked();"" />")
            TempBuf.AppendLine("  </td>")
            TempBuf.AppendLine("  <td>")
            TempBuf.AppendLine("<input type=""button"" name=""btnCancel"" id=""btnCancel"" value=" & Copient.PhraseLib.Lookup("term.cancel", LanguageID) & " onclick=""javascript:CancelClicked();"" />")
            TempBuf.AppendLine("  </td>")
            TempBuf.AppendLine("</tr>")
            TempBuf.AppendLine(" </table>")
            TempBuf.AppendLine("</div>")
            Send(TempBuf.ToString)
        Catch ex As Exception
            Send(ex.ToString)
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed AndAlso bRTConnectionOpened Then MyCommon.Close_LogixRT()
        End Try
    End Sub



    Sub UpdateFolderDate(ByRef foldersdate As Date, ByRef folderedate As Date, ByVal FolderId As Integer)
        Try
            MyCommon.QueryStr = "Update Folders set StartDate='" & foldersdate & "',EndDate='" & folderedate & "' where FolderID=" & FolderId
            MyCommon.LRT_Execute()
        Catch ex As Exception
            Send(ex.ToString)
        End Try
    End Sub

    Sub CheckUpdateActiveOfferDates(ByVal OldFolderStartDate As Date, ByVal OldFolderEndDate As Date, ByVal OfferStartDate As Date, ByVal OfferEndDate As Date, ByVal foldersdate As Date, ByVal folderedate As Date, ByRef UpdateOfferEndDates As Boolean, ByRef updatefolderdates As Boolean)
        If bOverrideMassUpdateRestriction Then
            UpdateOfferEndDates = True
        Else
            If OldFolderStartDate = OfferStartDate AndAlso OldFolderEndDate = OfferEndDate Then
                If OfferStartDate >= foldersdate AndAlso OfferStartDate <= folderedate Then
                    UpdateOfferEndDates = True
                End If
            ElseIf OldFolderStartDate = OfferStartDate AndAlso OldFolderEndDate <> OfferEndDate Then
                If OfferStartDate >= foldersdate AndAlso OfferEndDate <= folderedate Then
                    updatefolderdates = True
                End If
            ElseIf OldFolderStartDate <> OfferStartDate AndAlso OldFolderEndDate = OfferEndDate Then
                If OfferStartDate >= foldersdate AndAlso OfferStartDate <= folderedate Then
                    UpdateOfferEndDates = True
                End If
            ElseIf OldFolderStartDate <> OfferStartDate AndAlso OldFolderEndDate <> OfferEndDate Then
                If OfferStartDate >= foldersdate AndAlso OfferEndDate <= folderedate Then
                    updatefolderdates = True
                End If
            End If
        End If
    End Sub

    Sub checkUpdateOfferDates(ByVal OldFolderStartDate As Date, ByVal OfferStartDate As Date, ByVal OldFolderEndDate As Date, ByVal OfferEndDate As Date, ByVal foldersdate As Date, ByVal folderedate As Date, ByRef UpdateOfferStartDates As Boolean, ByRef UpdateOfferEndDates As Boolean, ByRef updatefolderdates As Boolean)
        If bOverrideMassUpdateRestriction Then
            UpdateOfferStartDates = True
            UpdateOfferEndDates = True
        Else
            If OldFolderStartDate = OfferStartDate AndAlso OldFolderEndDate = OfferEndDate Then
                UpdateOfferStartDates = True
                UpdateOfferEndDates = True
            ElseIf OldFolderStartDate = OfferStartDate AndAlso OldFolderEndDate <> OfferEndDate Then
                If OfferEndDate >= foldersdate AndAlso OfferEndDate <= folderedate Then
                    UpdateOfferStartDates = True
                End If
            ElseIf OldFolderStartDate <> OfferStartDate AndAlso OldFolderEndDate = OfferEndDate Then
                If OfferStartDate >= foldersdate AndAlso OfferStartDate <= folderedate Then
                    UpdateOfferEndDates = True
                End If
            ElseIf OldFolderStartDate <> OfferStartDate AndAlso OldFolderEndDate <> OfferEndDate Then
                If OfferStartDate >= foldersdate AndAlso OfferEndDate <= folderedate Then
                    updatefolderdates = True
                End If
            End If
        End If

    End Sub

    Function ValidateFolderInfo(ByVal FolderId As Integer, ByVal FolderName As String, ByVal FolderStartDate As String, ByVal FolderEndDate As String, ByVal FolderTheme As String, Optional ByVal IsUEDefaultFolder As Boolean = False) As Boolean

        Dim dt As DataTable
        Dim foldersdate As Date
        Dim folderedate As Date
        Dim OldFolderStartDate, OldFolderEndDate As Date
        Dim UpdateOfferStartDates As Boolean = False
        Dim UpdateOfferEndDates As Boolean = False
        Dim updatefolderdates As Boolean = False
        Dim OfferExpired As Boolean = False
        Dim bRTConnectionOpened As Boolean = False
        Dim dtOffers, dtFolders As DataTable
        Dim bAllowTimeWithStartEndDates As Boolean = False
        bAllowTimeWithStartEndDates = MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE) AndAlso (MyCommon.Fetch_UE_SystemOption(200) = "1")

        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
            MyCommon.Open_LogixRT()
            bRTConnectionOpened = True
        End If

        'AL-6583: Default Folder
        If IsUEDefaultFolder = True Then
            MyCommon.QueryStr = "Select * from Folders where DefaultUEFolder=1 and Folderid !=" & FolderId
            dt = MyCommon.LRT_Select()
            If dt.Rows.Count > 0 Then
                Send("NO " + Copient.PhraseLib.Lookup("term.DefaultfolderExists", LanguageID))
                Return False
            End If
        End If

        'update folder name if it is not empty
        If Not isEmpty(FolderName) Then
            RenameFolder(FolderId, FolderName, AdminUserID)
        End If

        'Update default folder state
        MyCommon.QueryStr = "Update Folders set DefaultUEFolder = " & IIf(IsUEDefaultFolder, 1, 0) & "where Folderid=" & FolderId
        MyCommon.LRT_Execute()

        If isEmpty(FolderStartDate) AndAlso isEmpty(FolderEndDate) Then
            If bAllowTimeWithStartEndDates Then
                Send("NO " + Copient.PhraseLib.Lookup("folder.EnterStartDateTime", LanguageID) & " " & Copient.PhraseLib.Lookup("folder.EnterEndDateTime", LanguageID))
            Else
                Send("NO " + Copient.PhraseLib.Lookup("folder.EnterStartDate", LanguageID) & " " & Copient.PhraseLib.Lookup("folder.EnterEndDate", LanguageID))
            End If
            Return False
        End If
        If Not isEmpty(FolderStartDate) AndAlso isEmpty(FolderEndDate) Then
            If bAllowTimeWithStartEndDates Then
                Send("NO " + Copient.PhraseLib.Lookup("folder.EnterEndDateTime", LanguageID))
            Else
                Send("NO " + Copient.PhraseLib.Lookup("folder.EnterEndDate", LanguageID))
            End If
            Return False
        End If
        If Not isEmpty(FolderEndDate) AndAlso isEmpty(FolderStartDate) Then
            If bAllowTimeWithStartEndDates Then
                Send("NO " + Copient.PhraseLib.Lookup("folder.EnterStartDateTime", LanguageID))
            Else
                Send("NO " + Copient.PhraseLib.Lookup("folder.EnterStartDate", LanguageID))
            End If
            Return False
        End If
        If Not isEmpty(FolderStartDate) Then
            If Not Date.TryParse(FolderStartDate, MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, foldersdate) Then
                If bAllowTimeWithStartEndDates Then
                    Send("NO " + Copient.PhraseLib.Lookup("folder.EnterStartDateTime", LanguageID))
                Else
                    Send("NO " + Copient.PhraseLib.Lookup("folder.EnterStartDate", LanguageID))
                End If
                Return False
            End If
        End If
        If Not isEmpty(FolderEndDate) Then
            If Not Date.TryParse(FolderEndDate, MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, folderedate) Then
                If bAllowTimeWithStartEndDates Then
                    Send("NO " + Copient.PhraseLib.Lookup("folder.EnterEndDateTime", LanguageID))
                Else
                    Send("NO " + Copient.PhraseLib.Lookup("folder.EnterEndDate", LanguageID))
                End If
                Return False
            End If
        End If
        If ((Not isEmpty(FolderStartDate)) OrElse (Not isEmpty(FolderEndDate))) Then
            If folderedate < foldersdate Then
                If bAllowTimeWithStartEndDates Then
                    Send("NO " + Copient.PhraseLib.Lookup("folders.BadEndDateTime", LanguageID))
                Else
                    Send("NO " + Copient.PhraseLib.Lookup("folders.BadEndDate", LanguageID))
                End If
                Return False
            ElseIf folderedate >= CDate("1/1/9999") Then
                Send("NO " + Copient.PhraseLib.Lookup("term.datetoolarge", LanguageID))
                Return False
            End If
        End If
        'If (MyCommon.Fetch_SystemOption(133) = "1" AndAlso MyCommon.Fetch_SystemOption(272) = "1") Then

        'Date.TryParse(FolderStartDate, MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, foldersdate)
        'Date.TryParse(FolderEndDate, MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, folderedate)

        'MyCommon.QueryStr = "select StartDate,EndDate from Folders with (NoLock) where FolderID=" & FolderId
        'dtFolders = MyCommon.LRT_Select
        'If dtFolders.Rows.Count > 0 Then
        '    If ((Not IsDBNull(dtFolders.Rows(0).Item("StartDate"))) OrElse (Not IsDBNull(dtFolders.Rows(0).Item("EndDate")))) Then
        '        OldFolderEndDate = Format(dtFolders.Rows(0).Item("EndDate"), "MM/dd/yyyy")
        '        OldFolderStartDate = Format(dtFolders.Rows(0).Item("StartDate"), "MM/dd/yyyy")
        '    End If
        'End If

        'If OldFolderStartDate <> foldersdate OrElse OldFolderEndDate <> folderedate Then
        'MyCommon.QueryStr = "select StartDate,EndDate,isnull(LastDeployClick,'1900-01-01') as LastDeployClick,IncentiveID,dbo.Calc_offer_Status(LinkID) as status from cpe_incentives CI with (NoLock) INNER JOIN FolderItems FI with (NoLock) ON CI.IncentiveID = FI.LinkID AND FI.FolderID=" & FolderId & " where CI.Deleted = 0"
        'dtOffers = MyCommon.LRT_Select
        'If dtOffers.Rows.Count > 0 Then
        'Dim result() As DataRow = dtOffers.Select("status =" & STATUS_FLAGS.STATUS_ACTIVE & " OR status =" & STATUS_FLAGS.STATUS_EXPIRED)
        'If result.Count > 0 Then
        'Send("NO " + Copient.PhraseLib.Lookup("folder.datesupdateerror", LanguageID))
        'Return False
        'End If
        'End If
        'End If
        'End If
        Return True
    End Function

    Public Sub ModifyFolderAsync(ByVal State As Object)
        Dim obj As Object() = State
        Dim FolderID As Integer = obj(0)
        Dim ModFolderName = obj(1)
        Dim ModFolderStartDate = obj(2)
        Dim ModFolderEndDate = obj(3)
        Dim ModFolderTheme = obj(4)
        Dim isMassupdateEnabled = obj(5)
        Dim IsdefaultUEFolder = obj(6)
        Dim status As String = ""
        Dim err As String = String.Empty
        Try
            SetFolderMassOperationStatus(FolderID, FOLDER_IN_USE + " Mass update of offers is in progress for the Offers in this Folder by" + " " + AdminName)
            ModifyFolder(FolderID, ModFolderName, ModFolderStartDate, ModFolderEndDate, ModFolderTheme, isMassupdateEnabled, IsdefaultUEFolder, status)
        Catch ex As Exception
            err = "Error: " + ex.ToString
            MyCommon.Write_Log(LogFile, err)
        Finally
            SetFolderMassOperationStatus(FolderID, FOLDER_NOT_IN_USE + status)
        End Try

    End Sub
    Public Sub DublicateOffersToFolderAsync(ByVal State As Object)
        Dim obj As Object() = State
        Dim FolderIDs As String = obj(0)
        Dim ItemIDs As String = obj(1)
        Dim FromOfferList As Boolean = obj(2)
        Dim ActionItem As Integer = obj(3)
        Dim strStatusDupOffers As String = obj(4)
        Dim NoOfDuplicateOffers As Integer = obj(5)
        Dim SrcFolderId As Integer = obj(6)
        Dim sNewFolderIDs As String() = Nothing
        Dim status As String = ""
        Dim err As String = String.Empty
        sNewFolderIDs = FolderIDs.Split(",")
        'Updating source and Destination folders status as folder in use.
        Try
            If (FromOfferList) Then
                SetFolderMassOperationStatus(-1, FOLDER_IN_USE + " Mass Operation is in progress for the Offers by" + " " + AdminName)
            Else
                SetFolderMassOperationStatus(SrcFolderId, FOLDER_IN_USE + " Mass Operation is in progress for the Offers in this Folder by" + " " + AdminName)
            End If

            For Each FolderID As String In sNewFolderIDs
                SetFolderMassOperationStatus(FolderID, FOLDER_IN_USE + " Mass Operation is in progress for the Offers in this Folder by" + " " + AdminName)
            Next

            For IndexDupOffers As Integer = 1 To NoOfDuplicateOffers
                DuplicateOffersToFolder(FolderIDs, ItemIDs, FromOfferList, ActionItem, strStatusDupOffers)
            Next IndexDupOffers
        Catch ex As Exception
            err = "Error: " + ex.ToString
            MyCommon.Write_Log(LogFile, err)
        Finally
            For Each FolderID As String In sNewFolderIDs
                SetFolderMassOperationStatus(FolderID, FOLDER_NOT_IN_USE + strStatusDupOffers)
            Next
            If (FromOfferList) Then
                SetFolderMassOperationStatus(-1, FOLDER_NOT_IN_USE + strStatusDupOffers)
            Else
                SetFolderMassOperationStatus(SrcFolderId, FOLDER_NOT_IN_USE + strStatusDupOffers)
            End If
        End Try
    End Sub

    Public Sub TransferOffersAsync(ByVal state As Object)
        Dim obj As Object() = state
        Dim SourceFolderID As String = obj(0)
        Dim DestFolderID As String = obj(1)
        Dim ItemIDs As String = obj(2)
        Dim FromOfferList As Boolean = obj(3)
        Dim status As String = ""
        Dim err As String = String.Empty
        Try
            'Updating source and Destination folders status as folder in use.
            If (FromOfferList) Then
                SetFolderMassOperationStatus(-1, FOLDER_IN_USE + " Mass Operation is in progress for the Offers by" + " " + AdminName)
            Else
                SetFolderMassOperationStatus(Convert.ToInt32(SourceFolderID), FOLDER_IN_USE + " Mass Operation is in progress for the Offers in this Folder by" + " " + AdminName)
            End If
            SetFolderMassOperationStatus(Convert.ToInt32(DestFolderID), FOLDER_IN_USE + " Mass Operation is in progress for the Offers in this Folder by" + " " + AdminName)
            TransferOffers(SourceFolderID, DestFolderID, ItemIDs, FromOfferList, status)
        Catch ex As Exception
            err = "Error: " + ex.ToString
            MyCommon.Write_Log(LogFile, err)
        Finally
            If (FromOfferList) Then
                SetFolderMassOperationStatus(-1, FOLDER_NOT_IN_USE + status)
            Else
                SetFolderMassOperationStatus(Convert.ToInt32(SourceFolderID), FOLDER_NOT_IN_USE + status)
            End If
            SetFolderMassOperationStatus(Convert.ToInt32(DestFolderID), FOLDER_NOT_IN_USE + status)
        End Try
    End Sub


    Sub ModifyFolder(ByVal FolderId As Integer, ByVal FolderName As String, ByVal FolderStartDate As String, ByVal FolderEndDate As String, ByVal FolderTheme As String, ByVal isMassupdateEnabled As Boolean, Optional ByVal IsUEDefaultFolder As Boolean = False, Optional ByRef Status As String = "")
        Dim dtFolders, dtOffers, dtErrorOffers As DataTable
        Dim row As DataRow
        Dim offerID As Integer
        Dim offerStatustext As String
        Dim foldersdate, folderedate, OldFolderStartDate, OldFolderEndDate, OfferStartDate, OfferEndDate, LastDeployDate As Date
        Dim UpdateOfferStartDates As Boolean = False
        Dim UpdateOfferEndDates As Boolean = False
        Dim updatefolderdates As Boolean = False
        Dim OfferExpired As Boolean = False
        Dim OfferActive As Boolean = False
        Dim bRTConnectionOpened As Boolean = False
        Dim Message As New StringBuilder("")
        Dim isValid As Boolean = True
        Dim UsermModifiedOfferStatus As String = ""
        Dim OfferIDs As String = ""
        Dim RestrictFolderOperationOnActiveOrExpiredOffers As Boolean = IIf(MyCommon.Fetch_SystemOption(272) = "1", True, False)

        Date.TryParse(FolderStartDate, MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, foldersdate)
        Date.TryParse(FolderEndDate, MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, folderedate)


        Dim SQLQuery As New StringBuilder("")
        Dim PhraseWrongStartDate As String = Copient.PhraseLib.Lookup("folders.WrongStartDate", LanguageID)
        Dim PhraseWrongEndDate As String = Copient.PhraseLib.Lookup("folders.WrongEndDate", LanguageID)
        Dim PhraseOfferNotInDateRange As String = Copient.PhraseLib.Lookup("folders.OfferNotInDateRange", LanguageID)
        Dim Phraseofferdateschanged As String = Copient.PhraseLib.Lookup("history.offerdateschanged", LanguageID)
        Dim Phraseofferstartdateschanged As String = Copient.PhraseLib.Lookup("history.offerstartdateschanged", LanguageID)
        Dim Phraseofferenddateschanged As String = Copient.PhraseLib.Lookup("history.offerenddateschanged", LanguageID)
        Dim Phrasecannotupdateexpired As String = Copient.PhraseLib.Lookup("folder.cannotupdateexpired", LanguageID)
        Dim Phrasecannotupdateactive As String = Copient.PhraseLib.Lookup("folder.cannotupdateactive", LanguageID)
        Dim Phrasecannotupdateconditions As String = Copient.PhraseLib.Lookup("folder.cannotupdateconditions", LanguageID)
        Dim PhraseDeploy As String = Copient.PhraseLib.Lookup("history.offer-deploy", LanguageID)
        Dim Phrasecannotupdateusermodoffers As String = Copient.PhraseLib.Lookup("folder.cannotupdateusermodifiedoffers", LanguageID)

        LogFile = "MassActionLog." & MyCommon.Leading_Zero_Fill(Year(Today), 4) & MyCommon.Leading_Zero_Fill(Month(Today), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
        MyCommon.Write_Log(LogFile, "Mass Update of Offers started....", True)

        Try

            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                MyCommon.Open_LogixRT()
                bRTConnectionOpened = True
            End If

            MyCommon.QueryStr = "Update folderthemes set themeid=(select themeid from themes where themedescription='" & FolderTheme & "') where Folderid=" & FolderId
            MyCommon.LRT_Execute()


            MyCommon.QueryStr = "select StartDate,EndDate from Folders with (NoLock) where FolderID=" & FolderId
            dtFolders = MyCommon.LRT_Select
            If dtFolders.Rows.Count > 0 Then
                If ((Not IsDBNull(dtFolders.Rows(0).Item("StartDate"))) OrElse (Not IsDBNull(dtFolders.Rows(0).Item("EndDate")))) Then
                    OldFolderEndDate = Format(dtFolders.Rows(0).Item("EndDate"), "MM/dd/yyyy")
                    OldFolderStartDate = Format(dtFolders.Rows(0).Item("StartDate"), "MM/dd/yyyy")
                End If
            End If
            If isMassupdateEnabled Then
                MyCommon.QueryStr = "select StartDate,EndDate,isnull(LastDeployClick,'1900-01-01') as LastDeployClick,IncentiveID,dbo.Calc_offer_Status(LinkID) as status from cpe_incentives CI with (NoLock) INNER JOIN FolderItems FI with (NoLock) ON CI.IncentiveID = FI.LinkID AND FI.FolderID=" & FolderId & " where CI.Deleted =0 and CI.StatusFlag not in (13,14,15)"
                dtOffers = MyCommon.LRT_Select

                If MyCommon.Fetch_SystemOption(284) = "1" Then
                    MyCommon.QueryStr = "select IncentiveID from cpe_incentives CI with (NoLock) INNER JOIN FolderItems FI with (NoLock) ON CI.IncentiveID = FI.LinkID AND FI.FolderID=" & FolderId & " where CI.Deleted = 0"
                    dtOffers = MyCommon.LRT_Select
                    MyCommon.QueryStr = "dbo.pa_Validate_OfferFolderdates"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.AddWithValue("@OffersT", dtOffers)
                    MyCommon.LRTsp.Parameters.AddWithValue("@FolderID", FolderId)
                    MyCommon.LRTsp.Parameters.AddWithValue("@IsMassUpdateEnabled", True)

                    dtErrorOffers = MyCommon.LRTsp_select
                    If dtErrorOffers.Rows.Count > 0 AndAlso MyCommon.NZ(dtErrorOffers.Rows(0)(0), -1) > 0 Then
                        Dim result As String = [String].Join(Environment.NewLine, dtErrorOffers.AsEnumerable().[Select](Function(row1) row1.Field(Of Int64)(0)))
                        MyCommon.Write_Log(LogFile, "Offer dates cannot be changed for following user modified offers:" & Environment.NewLine & "" & result & "", True)
                        UsermModifiedOfferStatus = Phrasecannotupdateusermodoffers
                    End If

                    Dim ValidOfferRows = dtOffers.AsEnumerable.Except(dtErrorOffers.AsEnumerable, DataRowComparer.[Default])
                    If ValidOfferRows.Any Then
                        dtOffers = ValidOfferRows.CopyToDataTable()
                        dtOffers.AcceptChanges()
                        OfferIDs = [String].Join(", ", dtOffers.AsEnumerable().[Select](Function(row1) row1.Field(Of Int64)(0)))
                        MyCommon.QueryStr = "select StartDate,EndDate,isnull(LastDeployClick,'1900-01-01') as LastDeployClick,IncentiveID,dbo.Calc_offer_Status(IncentiveID) as status from cpe_incentives with (NoLock) Where IncentiveID in (" & OfferIDs & ")"
                        dtOffers = MyCommon.LRT_Select
                    Else
                        dtOffers.Clear()
                    End If

                End If

                If dtOffers.Rows.Count > 0 Then
                    For Each row In dtOffers.Rows
                        UpdateOfferStartDates = False
                        UpdateOfferEndDates = False
                        updatefolderdates = False
                        OfferExpired = False

                        offerID = row.Item("IncentiveID")
                        offerStatustext = row.Item("status")
                        OfferStartDate = Format(row.Item("StartDate"), "MM/dd/yyyy")
                        OfferEndDate = Format(row.Item("EndDate"), "MM/dd/yyyy")
                        LastDeployDate = Format(row.Item("LastDeployClick"), "MM/dd/yyyy")

                        MyCommon.QueryStr = ""

                        'Validate Folder Dates                                
                        If offerStatustext = STATUS_FLAGS.STATUS_EXPIRED AndAlso RestrictFolderOperationOnActiveOrExpiredOffers = True AndAlso (OldFolderEndDate <> folderedate OrElse OldFolderStartDate <> foldersdate) Then
                            Message.AppendLine("Processing OfferID=" + offerID.ToString() + " " + PhraseWrongEndDate)
                            isValid = False
                        End If

                        If offerStatustext <> STATUS_FLAGS.STATUS_EXPIRED OrElse RestrictFolderOperationOnActiveOrExpiredOffers = False Then
                            If LastDeployDate = "01/01/1900" Then ' Offer never deployed
                                checkUpdateOfferDates(OldFolderStartDate, OfferStartDate, OldFolderEndDate, OfferEndDate, foldersdate, folderedate, UpdateOfferStartDates, UpdateOfferEndDates, updatefolderdates)
                            Else
                                If (offerStatustext = STATUS_FLAGS.STATUS_ACTIVE AndAlso folderedate <> OldFolderEndDate AndAlso folderedate > Date.now) Then
                                    CheckUpdateActiveOfferDates(OldFolderStartDate, OldFolderEndDate, OfferStartDate, OfferEndDate, foldersdate, folderedate, UpdateOfferEndDates, updatefolderdates)
                                    OfferActive = True
                                ElseIf (offerStatustext = STATUS_FLAGS.STATUS_ACTIVE AndAlso foldersdate <> OldFolderStartDate) Then
                                    OfferActive = True
                                ElseIf offerStatustext <> STATUS_FLAGS.STATUS_ACTIVE
                                    checkUpdateOfferDates(OldFolderStartDate, OfferStartDate, OldFolderEndDate, OfferEndDate, foldersdate, folderedate, UpdateOfferStartDates, UpdateOfferEndDates, updatefolderdates)
                                End If

                                If (UpdateOfferStartDates = True AndAlso OfferStartDate <> foldersdate) Or (UpdateOfferEndDates = True AndAlso OfferEndDate <> folderedate) Then
                                    SQLQuery.Append("Update CPE_Incentives with (RowLock) set StatusFlag=2, UpdateLevel=UpdateLevel+1, LastDeployClick=getdate(), DeployDeferred=0, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where IncentiveID=" & offerID & ";")
                                    SQLQuery.Append("insert into ActivityLog with (RowLock) (LinkID, ActivityTypeID, AdminID, Description, ActivityDate) Values(" & offerID.ToString() & ",3," & AdminUserID.ToString() & ",'" & PhraseDeploy & "',GETDATE());")
                                End If
                            End If
                        Else
                            OfferExpired = True
                        End If

                        If UpdateOfferStartDates = True AndAlso UpdateOfferEndDates = True Then
                            SQLQuery.Append("Update cpe_incentives set StartDate='" & foldersdate & "',EligibilityStartDate='" & foldersdate & "',TestingStartDate='" & foldersdate & "', EndDate='" & folderedate & "',EligibilityEndDate='" & folderedate & "',TestingEndDate='" & folderedate & "' where incentiveid=" & offerID & ";")
                            SQLQuery.Append("insert into ActivityLog with (RowLock) (LinkID, ActivityTypeID, AdminID, Description, ActivityDate) Values(" & offerID.ToString() & ",3," & AdminUserID.ToString() & ",'" & Phraseofferdateschanged & "',GETDATE());")
                        ElseIf UpdateOfferStartDates = True Then
                            SQLQuery.Append("Update cpe_incentives set StartDate='" & foldersdate & "',EligibilityStartDate='" & foldersdate & "', TestingStartDate='" & foldersdate & "' where incentiveid=" & offerID & ";")
                            SQLQuery.Append("insert into ActivityLog with (RowLock) (LinkID, ActivityTypeID, AdminID, Description, ActivityDate) Values(" & offerID.ToString() & ",3," & AdminUserID.ToString() & ",'" & Phraseofferstartdateschanged & "',GETDATE());")
                        ElseIf UpdateOfferEndDates = True Then
                            SQLQuery.Append("Update cpe_incentives set EndDate='" & folderedate & "',EligibilityEndDate='" & folderedate & "', TestingEndDate='" & folderedate & "' where incentiveid=" & offerID & ";")
                            SQLQuery.Append("insert into ActivityLog with (RowLock) (LinkID, ActivityTypeID, AdminID, Description, ActivityDate) Values(" & offerID.ToString() & ",3," & AdminUserID.ToString() & ",'" & Phraseofferenddateschanged & "',GETDATE());")
                        Else
                            If updatefolderdates = False AndAlso (OfferExpired = False OrElse (UpdateOfferStartDates = False AndAlso UpdateOfferEndDates = False)) AndAlso bOverrideMassUpdateRestriction = False Then
                                isValid = False
                                Message.AppendLine("Processing OfferID=" + offerID.ToString() + " " + PhraseOfferNotInDateRange)
                            End If
                        End If
                    Next

                    If isValid Then
                        If (SQLQuery.ToString() <> String.Empty) Then
                            MyCommon.QueryStr = SQLQuery.ToString()
                            MyCommon.LRT_Execute()
                            SQLQuery.Clear()
                        End If
                        If UsermModifiedOfferStatus <> "" Then
                            Status = UsermModifiedOfferStatus
                        Else
                            Status = Copient.PhraseLib.Lookup("term.success", LanguageID) + Copient.PhraseLib.Lookup("Mass-update-comp", LanguageID)
                        End If
                    Else
                        If OfferExpired Then
                            Message.Append("Error: " + Phrasecannotupdateexpired)
                        ElseIf OfferActive Then
                            Message.Append("Error: " + Phrasecannotupdateactive)
                        Else
                            Message.Append("Error: " + Phrasecannotupdateconditions)
                        End If
                        Status = Copient.PhraseLib.Lookup("mass-update-failure", LanguageID)
                    End If

                    MyCommon.Write_Log(LogFile, Message.ToString(), True)
                    MyCommon.Write_Log(LogFile, Status, True)
                End If
                If UsermModifiedOfferStatus <> "" Then
                    Status = UsermModifiedOfferStatus
                End If
            Else
                Status = String.Empty
                MyCommon.Write_Log(LogFile, "Mass Update Not Enabled.", True)
            End If

        Catch ex As Exception
            Status = "Error: " + ex.ToString
            MyCommon.Write_Log(LogFile, Status)
        Finally
            UpdateFolderDate(foldersdate, folderedate, FolderId)
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed AndAlso bRTConnectionOpened Then MyCommon.Close_LogixRT()
        End Try

    End Sub

    Sub SaveBuyerFolders(ByVal id As Long, ByVal FolderList As String, ByVal AdminuserID As Integer)

        Dim dt As DataTable
        Dim FolderEndDate As Date
        Try
            If id > 0 Then
                If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

                If FolderList <> "" Then
                    MyCommon.QueryStr = "select EndDate from folders with (nolock) where FolderID = @FolderID"
                    MyCommon.DBParameters.Add("@FolderID", SqlDbType.Int).Value = FolderList
                    dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                    If dt.Rows.Count > 0 Then
                        If (Not IsDBNull(dt.Rows(0).Item("EndDate"))) Then
                            FolderEndDate = dt.Rows(0).Item("EndDate")
                            If FolderEndDate < Date.Today Then
                                Sendb(Copient.PhraseLib.Lookup("folders.SelectFolderWithFutureDate", LanguageID))
                            End If
                        End If
                    End If
                End If

                MyCommon.QueryStr = "dbo.pa_FolderBuyers_Insert"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@ID", SqlDbType.BigInt).Value = id
                MyCommon.LRTsp.Parameters.Add("@FolderIDs", SqlDbType.NVarChar).Value = FolderList
                MyCommon.LRTsp.ExecuteNonQuery()
                MyCommon.Close_LRTsp()
            End If

            Sendb("OK|" & id)

        Catch ex As Exception
            Send("NO|" & ex.ToString)
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()

        End Try
    End Sub

    Sub SaveOfferFolders(ByVal OfferID As Long, ByVal FolderList As String, ByVal AdminUserID As Integer)
        Dim EngineID As Integer = 2
        Dim dt As DataTable
        Dim AllowMultipleBanners As Boolean = False
        Dim BannersEnabled As Boolean = False
        Dim BannerID As Integer
        Dim FolderEndDate As Date
        Dim EnableAdditionalLockoutResitriction As Boolean = IIf(MyCommon.Fetch_SystemOption(260) = "1", True, False)

        BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
        AllowMultipleBanners = (MyCommon.Fetch_SystemOption(67) = "1")
        Response.Cache.SetCacheability(HttpCacheability.NoCache)

        Try
            If OfferID > 0 Then
                If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
                If MyCommon.Fetch_SystemOption(143) = "1" Then
                    If FolderList <> "" Then
                        MyCommon.QueryStr = "select EndDate from folders with (nolock) where FolderID IN (SELECT items FROM Split (@FolderList, ','))"
                        MyCommon.DBParameters.Add("@FolderList", SqlDbType.NVarChar).Value = FolderList
                        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                        
                        If dt.Rows.Count > 0 Then
                            If (Not IsDBNull(dt.Rows(0).Item("EndDate"))) Then
                                FolderEndDate = dt.Rows(0).Item("EndDate")
                                If FolderEndDate < Date.Today Then
                                    Sendb(Copient.PhraseLib.Lookup("folders.SelectFolderWithFutureDate", LanguageID))
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End If
                If MyCommon.Fetch_SystemOption(131) = "1" Then
                    If FolderList <> "" AndAlso BannersEnabled AndAlso Not AllowMultipleBanners Then
                        MyCommon.QueryStr = "SELECT BannerID FROM BannerOffers WHERE OfferID=" & OfferID
                        dt = MyCommon.LRT_Select
                        If dt.Rows.Count > 0 Then
                            BannerID = MyCommon.NZ(dt.Rows(0).Item("BannerID"), 0)
                        End If
                        MyCommon.QueryStr = "SELECT bt.Lockoutdays FROM BannerThemes bt INNER JOIN  FolderThemes ft ON ft.ThemeID = bt.ThemeID " &
                                    " AND bt.BannerID =@BannerID INNER JOIN Folders fo ON ft.FolderID=fo.FolderID " &
                                    " WHERE fo.FolderID IN (SELECT items FROM Split (@FolderList, ','))"
                        dt = MyCommon.LRT_Select
                        MyCommon.DBParameters.Add("@BannerID", SqlDbType.Int).Value = BannerID
                        MyCommon.DBParameters.Add("@FolderList", SqlDbType.NVarChar).Value = FolderList
                        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                        If dt.Rows.Count = 0 Then
                            Sendb(Copient.PhraseLib.Lookup("folder.BadTheme", LanguageID))
                            Exit Sub
                        End If
                    End If
                End If
                If EnableAdditionalLockoutResitriction AndAlso Not Logix.UserRoles.EditOfferPastLockoutPeriod Then
                    If FolderList <> "" AndAlso BannersEnabled AndAlso Not AllowMultipleBanners Then
                        Dim isInLockoutPeriod As Boolean = IsSelectedFolderInLockOutPeriod(MyCommon, FolderList, OfferID)
                        If isInLockoutPeriod Then
                            Sendb(Copient.PhraseLib.Lookup("folder.inlockoutperiod", LanguageID))
                            Exit Sub
                        End If
                    End If
                End If
                MyCommon.QueryStr = "dbo.pa_FolderOffers_Insert"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@OfferIDs", SqlDbType.NVarChar).Value = OfferID
                MyCommon.LRTsp.Parameters.Add("@FolderIDs", SqlDbType.NVarChar).Value = FolderList
                MyCommon.LRTsp.ExecuteNonQuery()
                MyCommon.Close_LRTsp()

                MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=1, LastUpdate=getdate(),LastUpdatedByAdminID=" & AdminUserID & " where IncentiveID=" & OfferID
                MyCommon.LRT_Execute()

                MyCommon.Activity_Log2(3, 22, OfferID, AdminUserID, String.Format("{0},{1}{2}{3},{4},{5},{6}", Copient.PhraseLib.Lookup("history.folder-additem", LanguageID), Copient.PhraseLib.Lookup("offer.status1", LanguageID), " ", Copient.PhraseLib.Lookup("history.offerstartdate-edit", LanguageID), Copient.PhraseLib.Lookup("history.offerenddate-edit", LanguageID), Copient.PhraseLib.Lookup("history.testingstartdate-edit", LanguageID), Copient.PhraseLib.Lookup("history.testingenddate-edit", LanguageID)), , FolderList)

                If MyCommon.Fetch_SystemOption(192) = "1" Then
                    MyCommon.QueryStr = "dbo.pt_OfferDates_Default"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = EngineID
                    MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
                    MyCommon.LRTsp.ExecuteNonQuery()
                    MyCommon.Close_LRTsp()
                End If
            End If
            Sendb("OK|" & OfferID)
        Catch ex As Exception
            Send("NO|" & ex.ToString)
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()

        End Try
    End Sub

    Public Function IsSelectedFolderInLockOutPeriod(ByRef MyCommon As Copient.CommonInc, ByVal DestinationFolderId As Integer, ByVal SourceOfferId As Integer) As Boolean

        Dim bInLockoutPeriod As Boolean = False
        Dim LockOutDays As Integer

        Dim FolderStartDate As Date
        Dim dt As DataTable
        Dim BannerID As Integer = 0



        MyCommon.QueryStr = "SELECT BannerID FROM BannerOffers WHERE OfferID=" & SourceOfferId
        dt = MyCommon.LRT_Select
        If dt.Rows.Count > 0 Then
            BannerID = MyCommon.NZ(dt.Rows(0).Item("BannerID"), 0)
        End If


        MyCommon.QueryStr = "SELECT bt.Lockoutdays FROM BannerThemes bt INNER JOIN  FolderThemes ft ON ft.ThemeID = bt.ThemeID " &
                          " AND bt.BannerID = " & BannerID & " INNER JOIN Folders fo ON ft.FolderID=fo.FolderID " &
                          " WHERE fo.FolderID = " & DestinationFolderId & ""

        dt = MyCommon.LRT_Select
        If (dt.Rows.Count > 0) Then
            LockOutDays = MyCommon.NZ(dt.Rows(0).Item("lockoutdays"), 0)
            MyCommon.QueryStr = "select startdate from folders with (NoLock) where FolderId=" & DestinationFolderId

            dt = MyCommon.LRT_Select
            If (dt.Rows.Count > 0) Then
                If Not IsDBNull(dt.Rows(0).Item("startdate")) Then
                    FolderStartDate = dt.Rows(0).Item("startdate")

                    If FolderStartDate <= Date.Now.AddDays(LockOutDays) Then
                        bInLockoutPeriod = True
                    End If

                Else
                    bInLockoutPeriod = False
                End If
            End If
        End If

        Return bInLockoutPeriod
    End Function

    Sub CopyExpiredOffer(ByVal OfferID As Long, ByVal FolderList As String, ByVal AdminUserID As Integer)
        Dim EngineID As Integer = 2
        Dim dt As DataTable
        Dim AllowMultipleBanners As Boolean = False
        Dim BannersEnabled As Boolean = False
        Dim BannerID As Integer
        Dim FolderEndDate As Date
        Dim copiedOfferID As Integer = -1
        BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
        AllowMultipleBanners = (MyCommon.Fetch_SystemOption(67) = "1")
        Response.Cache.SetCacheability(HttpCacheability.NoCache)

        Try
            If OfferID > 0 Then
                If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

                If FolderList <> "" Then
                    MyCommon.QueryStr = "select EndDate from folders with (nolock) where FolderID IN (SELECT items FROM Split (@FolderList, ','))"
                    MyCommon.DBParameters.Add("@FolderList", SqlDbType.NVarChar).Value = FolderList
                    dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                    If dt.Rows.Count > 0 Then
                        If (Not IsDBNull(dt.Rows(0).Item("EndDate"))) Then
                            FolderEndDate = dt.Rows(0).Item("EndDate")
                            If FolderEndDate < Date.Today Then
                                Sendb("NO|" & Copient.PhraseLib.Lookup("folders.SelectFolderWithFutureDate", LanguageID))
                                Exit Sub
                            End If
                        End If
                    End If
                End If

                If FolderList <> "" Then
                    Dim isInLockoutPeriod As Boolean = IsSelectedFolderInLockOutPeriod(MyCommon, FolderList, OfferID)
                    If (isInLockoutPeriod AndAlso Not Logix.UserRoles.EditOfferPastLockoutPeriod) Then
                        Sendb("NO|" & Copient.PhraseLib.Lookup("folder.inlockoutperiod", LanguageID))
                        Exit Sub
                    End If
                End If

                If MyCommon.Fetch_SystemOption(131) = "1" Then
                    If FolderList <> "" AndAlso BannersEnabled AndAlso Not AllowMultipleBanners Then
                        MyCommon.QueryStr = "SELECT BannerID FROM BannerOffers WHERE OfferID=" & OfferID
                        dt = MyCommon.LRT_Select
                        If dt.Rows.Count > 0 Then
                            BannerID = MyCommon.NZ(dt.Rows(0).Item("BannerID"), 0)
                        End If
                        MyCommon.QueryStr = "SELECT bt.Lockoutdays FROM BannerThemes bt INNER JOIN  FolderThemes ft ON ft.ThemeID = bt.ThemeID " &
                                    " AND bt.BannerID =@BannerID INNER JOIN Folders fo ON ft.FolderID=fo.FolderID " &
                                    " WHERE fo.FolderID IN (SELECT items FROM Split (@FolderList, ','))"
                        MyCommon.DBParameters.Add("@BannerID", SqlDbType.Int).Value = BannerID
                        MyCommon.DBParameters.Add("@FolderList", SqlDbType.NVarChar).Value = FolderList
                        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                        If dt.Rows.Count = 0 Then
                            Sendb("NO|" & Copient.PhraseLib.Lookup("folder.BadTheme", LanguageID))
                            Exit Sub
                        End If
                    End If
                End If

                Dim bUseDisplayDates As Boolean = IIf(MyCommon.Fetch_UE_SystemOption(143) = "1", True, False)
                Dim EngineSubTypeID As Integer = 0
                MyCommon.QueryStr = "select EngineSubTypeID from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID & ";"
                Dim rst As DataTable = MyCommon.LRT_Select
                If rst.Rows.Count > 0 Then
                    EngineSubTypeID = MyCommon.NZ(rst.Rows(0).Item("EngineSubTypeID"), 0)
                End If

                MyCommon.QueryStr = "dbo.pc_Copy_CPE_Offer"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@SourceOfferID", SqlDbType.BigInt).Value = OfferID
                MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = 9
                MyCommon.LRTsp.Parameters.Add("@EngineSubTypeID", SqlDbType.Int).Value = EngineSubTypeID
                MyCommon.LRTsp.Parameters.Add("@AdminUserID", SqlDbType.BigInt).Value = AdminUserID
                MyCommon.LRTsp.Parameters.Add("@NewOfferID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                MyCommon.LRTsp.ExecuteNonQuery()
                copiedOfferID = MyCommon.LRTsp.Parameters("@NewOfferID").Value
                MyCommon.Close_LRTsp()
                If (copiedOfferID > 0) Then
                    If bUseDisplayDates Then
                        'Updating TemplatePermission table with the Disallow_DisplayDates based on the UE SystemOption #143
                        UpdateTemplatePermissions(MyCommon, OfferID, copiedOfferID, 143)
                        SaveOfferDisplayDates(MyCommon, OfferID, copiedOfferID, AdminUserID, EngineID)
                    End If
                    SetPromotionDisplay(MyCommon, OfferID)
                    MyCommon.Activity_Log(3, copiedOfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-copy", LanguageID))

                    MyCommon.QueryStr = "dbo.pa_FolderOffers_Insert"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@OfferIDs", SqlDbType.NVarChar).Value = copiedOfferID
                    MyCommon.LRTsp.Parameters.Add("@FolderIDs", SqlDbType.NVarChar).Value = FolderList
                    MyCommon.LRTsp.ExecuteNonQuery()
                    MyCommon.Close_LRTsp()

                    MyCommon.Activity_Log2(3, 22, copiedOfferID, AdminUserID, Copient.PhraseLib.Lookup("history.folder-additem", LanguageID), , FolderList)

                    MyCommon.QueryStr = "dbo.pt_OfferDates_Default"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = EngineID
                    MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = copiedOfferID
                    MyCommon.LRTsp.ExecuteNonQuery()
                    MyCommon.Close_LRTsp()

                    Sendb("OK|" & copiedOfferID)
                Else
                    Send("NO|" & "-1")
                End If
            End If

        Catch ex As Exception
            Send("NO|" & ex.ToString)
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()

        End Try
    End Sub

    'Updating Promotion Display flag and Prorate on display flag when system options are off.      
    Sub SetPromotionDisplay(ByRef Common As Copient.CommonInc, ByVal OfferID As Long)
        Dim isEnabled As Boolean
        isEnabled = IIf(Common.Fetch_UE_SystemOption(145) = "1", True, False)
        If isEnabled = False Then
            Common.QueryStr = "update CPE_Incentives set PromotionDisplay=0 where IncentiveID=" & OfferID & ";"
            Common.LRT_Execute()
        End If
        'Updating Prorate on Display flag when system option 154 is off.   
        isEnabled = IIf(Common.Fetch_UE_SystemOption(154) = "1", True, False)
        If isEnabled = False Then
            Common.QueryStr = "update CPE_Incentives set ProrateonDisplay=0 where IncentiveID=" & OfferID & ";"
            Common.LRT_Execute()
        End If
    End Sub


    Sub UpdateTemplatePermissions(ByRef Common As Copient.CommonInc, ByVal SourceOfferID As Long, ByVal OfferID As Long, ByVal systemOption As Integer)

        Dim dtTempPermission As New DataTable
        Dim Disallow_DisplayDates As Integer = Integer.MinValue

        If (systemOption = 143) Then
            Common.QueryStr = "SELECT Disallow_DisplayDates from TemplatePermissions with (NoLock) WHERE OfferID = " & SourceOfferID & ";"
            dtTempPermission = Common.LRT_Select()
            If dtTempPermission.Rows.Count > 0 Then
                Disallow_DisplayDates = Convert.ToInt16(dtTempPermission.Rows(0).Item("Disallow_DisplayDates"))
            End If
            Common.QueryStr = "UPDATE TemplatePermissions with (RowLock) Set Disallow_DisplayDates=" & Disallow_DisplayDates & " WHERE OfferID = " & OfferID
            Common.LRT_Execute()
        End If
    End Sub

    Sub SaveOfferDisplayDates(ByRef Common As Copient.CommonInc, ByVal SourceOfferID As Long, ByVal OfferID As Long, ByVal AdminUserID As Integer, ByVal EngineId As Integer)

        Common.QueryStr = "SELECT DisplayStartDate,DisplayEndDate FROM OfferAccessoryFields with (NoLock) WHERE OfferID = " & SourceOfferID & ";"
        Dim dtODisp As New DataTable
        Dim startDate As String = ""
        Dim endDate As String = ""
        dtODisp = Common.LRT_Select()
        If dtODisp.Rows.Count > 0 Then
            startDate = Common.NZ(dtODisp.Rows(0).Item("DisplayStartDate"), Nothing)
            endDate = Common.NZ(dtODisp.Rows(0).Item("DisplayEndDate"), Nothing)
        End If
        Common.QueryStr = "dbo.pa_UpdateOfferAccessoryFields"
        Common.Open_LRTsp()
        Common.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
        Common.LRTsp.Parameters.Add("@DisplayStartTime", SqlDbType.DateTime).Value = IIf(String.IsNullOrEmpty(startDate), DBNull.Value, startDate)
        Common.LRTsp.Parameters.Add("@DisplayEndTime", SqlDbType.DateTime).Value = IIf(String.IsNullOrEmpty(endDate), DBNull.Value, endDate)

        Common.LRTsp.Parameters.Add("@RedemThresholdPerHour", SqlDbType.BigInt).Value = DBNull.Value
        Common.LRTsp.Parameters.Add("@AdminUserId", SqlDbType.BigInt).Value = AdminUserID
        Common.LRTsp.Parameters.Add("@EngineID", SqlDbType.BigInt).Value = EngineId

        Common.LRTsp.Parameters.Add("@OptionID", SqlDbType.BigInt).Value = 85
        Common.LRTsp.ExecuteNonQuery()
        Common.Close_LRTsp()

    End Sub

    Sub SendOfferFoldersWithFutureDate(ByVal FolderList As String, ByVal AdminUserID As Integer)
        Dim dt As DataTable
        Dim FolderEndDate As Date
        Response.Cache.SetCacheability(HttpCacheability.NoCache)

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If FolderList <> "" Then
                MyCommon.QueryStr = "select EndDate from folders with (nolock) where FolderID in (SELECT items FROM Split (@FolderList, ','))"
                MyCommon.DBParameters.Add("@FolderList", SqlDbType.NVarChar).Value = FolderList
                dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                
                If dt.Rows.Count > 0 Then
                    For Each dr As DataRow In dt.Rows
                        If (Not IsDBNull(dr.Item("EndDate"))) Then
                            FolderEndDate = dr.Item("EndDate")
                            If FolderEndDate < Date.Today Then
                                Send(Copient.PhraseLib.Lookup("folders.SelectFolderWithFutureDate", LanguageID))
                                Exit Sub
                            End If
                        Else
                            Send(Copient.PhraseLib.Lookup("folders.SelectFolderWithFutureDate", LanguageID))
                            Exit Sub
                        End If
                    Next
                End If
            End If
            Send("OK|" & FolderList)
        Catch ex As Exception
            Send("NO|" & ex.Message)
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try
    End Sub

    Sub SendOfferSearch(ByVal SearchTerms As String)
        Dim dt As DataTable
        Dim row As DataRow
        Dim TempBuf As New StringBuilder()
        Dim FolderID, OfferID As Integer
        Dim FolderName, OfferName As String

        Response.Cache.SetCacheability(HttpCacheability.NoCache)

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

            If bStoreUser Then
                sJoin = " Full Outer Join OfferLocUpdate olu with (NoLock) on AOLV.OfferID=olu.OfferID "
                wherestr = " and (LocationID in (" & sValidLocIDs & ") or (CreatedByAdminID in (" & sValidSU & ") and Isnull(LocationID,0)=0)) "
            End If
            'If ViewOffersRegardlessBuyer permission is not set, get buyer Specific Offer
            If (MyCommon.IsEngineInstalled(9) AndAlso MyCommon.IsUserAssociatedWithAnyBuyer(AdminUserID) AndAlso Not (Logix.UserRoles.ViewOffersRegardlessBuyer) Or bStoreUser) Then
                buyerJoin = " inner join BuyerRoleusers BRU on AOLV.BuyerID = BRU.BuyerID "
                buyerwherestr = " BRU.adminUSerID = " & AdminUserID & " and"
            End If
            MyCommon.QueryStr = "select F.FolderID, F.FolderName, AOLV.OfferID, AOLV.Name as OfferName from FolderItems as FI with (NoLock) " &
                                "inner join AllOffersListView as AOLV with (NoLock) on AOLV.OfferID = FI.LinkID and FI.LinkTypeID=1  " &
                                "inner join Folders as F with (NoLock) on F.FolderID = FI.FolderID  " & sJoin & buyerJoin &
                                "where " & buyerwherestr & " AOLV.deleted=0 and AOLV.Name like '%" & MyCommon.Parse_Quotes(SearchTerms) & "%' " & wherestr

            If (bEnableRestrictedAccessToUEOfferBuilder) Then
                MyCommon.QueryStr &= GetRestrictedAccessToUEBuilderQuery(MyCommon, Logix, "AOLV")
            End If

            MyCommon.QueryStr &= " order by OfferName;"

            dt = MyCommon.LRT_Select
            If dt.Rows.Count > 0 Then
                TempBuf.AppendLine("<table summary="""" style=""width:97%;white-space: nowrap;"">")
                TempBuf.AppendLine("  <tr>")
                TempBuf.AppendLine("    <th>" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & "</th>")
                TempBuf.AppendLine("    <th>" & Copient.PhraseLib.Lookup("term.foldername", LanguageID) & "</th>")
                TempBuf.AppendLine("  </tr>")
                For Each row In dt.Rows
                    FolderID = MyCommon.NZ(row.Item("FolderID"), 0)
                    FolderName = MyCommon.NZ(row.Item("FolderName"), "[" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "]")
                    OfferID = MyCommon.NZ(row.Item("OfferID"), 0)
                    OfferName = MyCommon.NZ(row.Item("OfferName"), "[" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "]")
                    TempBuf.AppendLine("  <tr>")
                    TempBuf.AppendLine("    <td>" & OfferName & "</td>")
                    TempBuf.AppendLine("    <td><a href=""javascript:folderLinkClicked(" & FolderID & ");"">" & FolderName & "</a></td>")
                    TempBuf.AppendLine("  </tr>")
                Next
                TempBuf.AppendLine("</table>")
            Else
                TempBuf.AppendLine("<b>" & Copient.PhraseLib.Lookup("folders.NoOffersFound", LanguageID) & "</b>")
            End If

            Send(TempBuf.ToString)
        Catch ex As Exception
            Send(ex.ToString)
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()

        End Try
    End Sub

    Function ShowWarning(ByVal FolderID As Integer, ByVal OfferStartDate As Date, ByVal OfferEndDate As Date) As String
        Dim dt As DataTable
        Dim FolderEndDate As Date
        Dim FolderStartDate As Date
        Dim WarningMessage As String = ""
        Response.Cache.SetCacheability(HttpCacheability.NoCache)

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            ' find the enddate of the folder and show the warning as well
            MyCommon.QueryStr = "select StartDate,EndDate from Folders with (NoLock) where FolderID=" & FolderID
            dt = MyCommon.LRT_Select
            If dt.Rows.Count > 0 Then
                If ((Not IsDBNull(dt.Rows(0).Item("StartDate"))) OrElse (Not IsDBNull(dt.Rows(0).Item("EndDate")))) Then
                    FolderEndDate = Format(dt.Rows(0).Item("EndDate"), "MM/dd/yyyy")
                    FolderStartDate = Format(dt.Rows(0).Item("StartDate"), "MM/dd/yyyy")
                    If OfferEndDate > FolderEndDate OrElse OfferStartDate < FolderStartDate Then
                        WarningMessage = ("<b>" & Copient.PhraseLib.Lookup("folders.OfferNotInFolderDateRange", LanguageID) & "</b>")
                    Else
                        WarningMessage = ("<b>" & Copient.PhraseLib.Lookup("folders.DateMatched", LanguageID) & "</b>")
                    End If
                Else
                    WarningMessage = ("<b>" & Copient.PhraseLib.Lookup("folders.ExpiryDateUnavailable", LanguageID) & "</b>")
                End If
            Else
                WarningMessage = ("<b>" & Copient.PhraseLib.Lookup("folders.ExpiryDateUnavailable", LanguageID) & "</b>")
            End If

        Catch ex As Exception
            Send(ex.ToString)
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try
        Return WarningMessage
    End Function

    Sub SendFolderSearch(ByVal SearchTerms As String)
        Dim dt As DataTable
        Dim row As DataRow
        Dim TempBuf As New StringBuilder()
        Dim FolderID, ParentID As Integer
        Dim FolderName, ParentName As String

        Response.Cache.SetCacheability(HttpCacheability.NoCache)

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

            MyCommon.QueryStr = "select F1.FolderID, F1.FolderName, F2.FolderID as ParentID, F2.FolderName as ParentName " &
                                  "from Folders as F1 with (NoLock) " &
                                  "left join Folders as F2 with (NoLock) on F2.FolderID = F1.ParentFolderID " &
                                  "where F1.FolderName like '%" & MyCommon.Parse_Quotes(SearchTerms) & "%'" &
                                  " order by FolderName;"

            dt = MyCommon.LRT_Select
            If dt.Rows.Count > 0 Then
                TempBuf.AppendLine("<table summary="""" style=""width:97%;white-space: nowrap;"">")
                TempBuf.AppendLine("  <tr>")
                TempBuf.AppendLine("    <th>" & Copient.PhraseLib.Lookup("term.foldername", LanguageID) & "</th>")
                TempBuf.AppendLine("    <th>" & Copient.PhraseLib.Lookup("term.parentname", LanguageID) & "</th>")
                TempBuf.AppendLine("  </tr>")
                For Each row In dt.Rows
                    FolderID = MyCommon.NZ(row.Item("FolderID"), 0)
                    FolderName = MyCommon.NZ(row.Item("FolderName"), "[" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "]")
                    ParentID = MyCommon.NZ(row.Item("ParentID"), 0)
                    ParentName = MyCommon.NZ(row.Item("ParentName"), "[" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "]")
                    TempBuf.AppendLine("  <tr>")
                    TempBuf.AppendLine("    <td><a href=""javascript:folderLinkClicked(" & FolderID & ");"">" & FolderName & "</a></td>")
                    TempBuf.AppendLine("    <td><a href=""javascript:folderLinkClicked(" & ParentID & ");"">" & ParentName & "</a></td>")
                    TempBuf.AppendLine("  </tr>")
                Next
                TempBuf.AppendLine("</table>")
            Else
                TempBuf.AppendLine("<b>" & Copient.PhraseLib.Lookup("folders.NoFoldersFound", LanguageID) & "</b>")
            End If

            Send(TempBuf.ToString)
        Catch ex As Exception
            Send(ex.ToString)
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()

        End Try
    End Sub

    Function IndexOfCount(ByVal str As String, ByVal searchChar As Char) As Integer
        Dim Count As Integer = 0

        For i As Integer = 0 To str.Length - 1
            If str(i) = searchChar Then Count += 1
        Next

        Return Count
    End Function
    '-----------------------------------------------------------------------------------------------------------------------

    'BR02 Common functions should be placed in some common library. may be CMOffer/CPEOffer??

    Function ValidateForDeploy(ByRef MyCommon As Copient.CommonInc, ByVal OfferID As Long, ByRef IsCMOffer As Boolean, ByVal deploytransreqskip As String) As String
        'first check if it is a cm offer or a cpe/ue
        'Dim row As DataRow
        Dim rst, dt As DataTable
        Dim Response As String = ""
        Dim Isdeployable As Boolean = False
        Dim ErrorMsg As String = ""
        Dim roid As Integer = 0
        Dim IsUEOffer As Boolean = False
        Dim bStatus As Boolean

        MyCommon.QueryStr = "select RewardOptionID, TierLevels from CPE_RewardOptions with (NoLock) where TouchResponse=0 and IncentiveID=" & OfferID & ";"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
            roid = MyCommon.NZ(rst.Rows(0).Item("RewardOptionID"), 0)
        End If

        MyCommon.QueryStr = "select EngineID from OfferIds with (NoLock) where OfferID=" & OfferID & ";"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
            If MyCommon.Extract_Val(MyCommon.NZ(rst.Rows(0).Item("EngineID"), 0)) = 9 Then
                IsUEOffer = True
            End If
        End If

        MyCommon.QueryStr = "select OfferId FROM Offers WHERE OfferID=" & OfferID
        rst = MyCommon.LRT_Select
        MyCommon.QueryStr = "select IncentiveId FROM CPE_Incentives WHERE IncentiveId=" & OfferID
        dt = MyCommon.LRT_Select

        If rst.Rows.Count > 0 Then
            'this is cm offer
            IsCMOffer = True
            bStatus = MyExport.ValidateOfferForDeploy(OfferID, LanguageID, True, False)
            If Not bStatus Then
                Response = MyExport.GetErrorMsg
            End If
        ElseIf dt.Rows.Count > 0 Then
            'this is cpe/ue offer
            Isdeployable = MyCPEOffer.IsDeployableOffer(Logix, MyCommon, OfferID, roid, IsUEOffer, ErrorMsg)
            If deploytransreqskip <> "1" Then
                If Isdeployable Then Isdeployable = MeetsTranslationRequirements(MyCommon, OfferID, roid, ErrorMsg)
            End If
            If MyCommon.Fetch_SystemOption(131) = "1" AndAlso MyCommon.Fetch_SystemOption(67) = "0" Then
                If Isdeployable Then Isdeployable = MeetsLockOutRequirement(Logix, MyCommon, OfferID)
                If (Not Isdeployable) Then
                    ErrorMsg = Copient.PhraseLib.Lookup("cpeoffer-sum.deployalertforlockout", LanguageID)
                End If
            End If
            If Not Isdeployable AndAlso ErrorMsg <> "" Then
                Response = ErrorMsg
            End If
        End If

        Return Response

    End Function

    Function MeetsLockOutRequirement(ByRef Logix As Copient.LogixInc, ByRef MyCommon As Copient.CommonInc, ByVal OfferId As Integer) As Boolean

        Dim bMeetsLockoutReq As Boolean = True
        Dim FolderId As Long
        Dim LockOutDays As Integer
        Dim FolderStartDate As Date
        Dim dt As DataTable
        Dim BannerID As Integer = 0

        If Logix.UserRoles.DeployOffersPastLockoutDate = False Then

            MyCommon.QueryStr = "SELECT BannerID FROM BannerOffers WHERE OfferID=" & OfferId
            dt = MyCommon.LRT_Select
            If dt.Rows.Count > 0 Then
                BannerID = MyCommon.NZ(dt.Rows(0).Item("BannerID"), 0)
            End If

            MyCommon.QueryStr = "select folderid from folderitems with (nolock) where linkid =" & OfferId
            dt = MyCommon.LRT_Select
            If dt.Rows.Count > 0 Then
                FolderId = dt.Rows(0).Item("folderid")
            End If

            MyCommon.QueryStr = "SELECT bt.Lockoutdays FROM BannerThemes bt INNER JOIN  FolderThemes ft ON ft.ThemeID = bt.ThemeID " &
                              " AND bt.BannerID = " & BannerID & " INNER JOIN Folders fo ON ft.FolderID=fo.FolderID " &
                              " WHERE fo.FolderID = " & FolderId & ""

            dt = MyCommon.LRT_Select
            If (dt.Rows.Count > 0) Then
                LockOutDays = MyCommon.NZ(dt.Rows(0).Item("lockoutdays"), 0)
                MyCommon.QueryStr = "select startdate from folders with (NoLock) where FolderId=" & FolderId

                dt = MyCommon.LRT_Select
                If (dt.Rows.Count > 0) Then
                    If Not IsDBNull(dt.Rows(0).Item("startdate")) Then
                        FolderStartDate = dt.Rows(0).Item("startdate")

                        If FolderStartDate <= Date.Now.AddDays(LockOutDays) Then
                            bMeetsLockoutReq = False
                        End If

                    Else
                        bMeetsLockoutReq = True
                    End If
                End If
            End If
        End If
        Return bMeetsLockoutReq
    End Function

    Public Function CheckForOfferCondition(ByRef MyCommon As Copient.CommonInc, ByVal OfferID As Long) As Boolean

        Dim rst As DataTable
        Dim bConditionAssigned As Boolean = False
        Try
            MyCommon.QueryStr = "select ConditionID, ConditionTypeID from OfferConditions OC with (NoLock) " &
                           "inner join Offers O with (NoLock) on O.OfferID = OC.OfferID " &
                                    "where OC.Deleted = 0 and O.Deleted = 0 and OC.OfferID =" & OfferID & " "
            rst = MyCommon.LRT_Select
            bConditionAssigned = (rst.Rows.Count > 0)

        Catch e As Exception

        End Try
        Return bConditionAssigned

    End Function

    Function MeetsTranslationRequirements(ByRef Common As Copient.CommonInc, ByVal OfferID As Long, ByVal ROID As Long, ByRef ErrorMsg As String) As Boolean
        Dim ReturnVal As Boolean = True
        Dim ErrorTerm As String = ""
        Dim TranslatedErrorTerm As String = ""
        Dim DeployType As String = ""

        ErrorMsg = ""
        If Common.Fetch_SystemOption(124) = "1" Then  'see if multi-language is enabled
            ErrorTerm = CheckForTranslationDeployError(Common, ROID)
            If Not (ErrorTerm = "") Then
                'if this section weren't in place, then you could get around the warning message by clicking defer deploy and then clicking deploy, without specifically addressing the warning
                If GetCgiValue("deferdeploy") <> "" Then
                    DeployType = "deferdeploy"
                Else
                    DeployType = "deploy"
                End If
                TranslatedErrorTerm = Copient.PhraseLib.DecodeEmbededTokens(ErrorTerm, LanguageID)
                ErrorMsg = Copient.PhraseLib.Detokenize("term.ReqTransFailed", LanguageID, TranslatedErrorTerm) &
                  " <input type=""submit"" class=""regular"" id=""" & DeployType & """ name=""" & DeployType & """ value=""" & Copient.PhraseLib.Lookup("term.yes", LanguageID) & """ onclick=""document.getElementById('deploytransreqskip').value='1';"" />" &
                  "<input type=""hidden"" id=""deploytransreqskip"" name=""deploytransreqskip"" value="""" />"
                ReturnVal = False
            End If
        End If  'multi-language enabled

        Return ReturnVal

    End Function

    Function CheckForTranslationDeployError(ByRef Common As Copient.CommonInc, ByVal ROID As Long) As String
        Dim ReturnVal As String
        Common.QueryStr = "dbo.pa_ReqTranslationCheck"
        Common.Open_LRTsp()
        Common.LRTsp.Parameters.Add("@ROID", SqlDbType.BigInt).Value = ROID
        Common.LRTsp.Parameters.Add("@ErrorTerm", SqlDbType.VarChar, 2000).Direction = ParameterDirection.Output
        Common.LRTsp.ExecuteNonQuery()
        ReturnVal = Common.LRTsp.Parameters("@ErrorTerm").Value
        Common.Close_LRTsp()
        Return ReturnVal
    End Function

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Validate SendOutBound


    Function ValidateForPreValidate(ByRef Common As Copient.CommonInc, ByVal OfferID As Long) As String

        Dim dt As DataTable
        Dim iWorkflowStatus As Integer = 0
        Dim ResponseMessage As String = ""
        Dim StatusCode As Copient.LogixInc.STATUS_FLAGS = Copient.LogixInc.STATUS_FLAGS.STATUS_UNKNOWN
        Dim offerStatus, offerStatustext As String

        Try
            If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()
            offerStatus = Logix.GetOfferStatus(OfferID, LanguageID, StatusCode)
            offerStatustext = Logix.GetOfferStatusText(StatusCode, LanguageID)

            Common.QueryStr = "select WorkFlowStatus  from  Offers with (nolock)  Where OfferID=" & OfferID & " and Deleted=0"
            dt = Common.LRT_Select
            If dt.Rows.Count > 0 Then
                iWorkflowStatus = MyCommon.NZ(dt.Rows(0).Item("WorkflowStatus"), 0)
            End If

            If iWorkflowStatus = 1 Then
                ResponseMessage = "The status can not be changed as it is already Pre-Validate"
            ElseIf offerStatustext.Trim.ToUpper = "EXPIRED" Then
                ResponseMessage = "The status cannot be changed to Pre-Validate as the Offer is expired"
            ElseIf OfferDeployed(Common, OfferID) Then
                ResponseMessage = "The status cannot be changed to Pre-Validate as the Offer is deployed"
            End If

            Return ResponseMessage
        Catch Ex As Exception
            Common.Write_Log(LogFile, Ex.Message, True)
            Return ResponseMessage
        End Try
    End Function

    Function OfferDeployed(ByRef Common As Copient.CommonInc, ByVal OfferID As Long) As Boolean

        Dim Deployed As Boolean = False
        Dim dt As DataTable

        Common.QueryStr = "select CMOADeploySuccessDate, StatusFlag, DeployDeferred from Offers with (NoLock) where OfferId=" & OfferID
        dt = Common.LRT_Select
        If Not IsDBNull(dt.Rows(0).Item("CMOADeploySuccessDate")) Then
            Deployed = True
        End If

        If (Common.NZ(dt.Rows(0).Item("StatusFlag"), -1) <> 2) Then
            If (Common.NZ(dt.Rows(0).Item("StatusFlag"), 0) > 0) Then
                If (Common.NZ(dt.Rows(0).Item("DeployDeferred"), False) = False) Then
                    Deployed = False
                End If
            End If
        End If

        Return Deployed

    End Function

    Function ValidateForPostValidate(ByRef Common As Copient.CommonInc, ByVal OfferID As Long) As String
        Dim dt As DataTable
        Dim iWorkflowStatus As Integer = 0
        Dim ResponseMessage As String = ""
        Dim bStatus As Boolean

        Try
            If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()

            bStatus = MyExport.ValidateOfferForDeploy(OfferID, LanguageID, True, True)
            If bStatus Then
                bStatus = MyExport.TransferOfferToTest(OfferID)
                If Not bStatus Then
                    ResponseMessage = MyExport.GetErrorMsg
                End If
            Else
                ResponseMessage = MyExport.GetErrorMsg
            End If
            If ResponseMessage <> "" Then
                Common.QueryStr = "select WorkFlowStatus  from  Offers with (nolock)  Where OfferID=" & OfferID & " and Deleted=0"
                dt = Common.LRT_Select
                If dt.Rows.Count > 0 Then
                    iWorkflowStatus = MyCommon.NZ(dt.Rows(0).Item("WorkflowStatus"), 0)
                End If

                If iWorkflowStatus <> 1 Then
                    ResponseMessage = "The status can be changed to post validate only if it is pre validate"
                End If
            End If

            Return ResponseMessage
        Catch Ex As Exception
            Common.Write_Log(LogFile, Ex.Message, True)
            Return ResponseMessage
        End Try
    End Function

    Function ValidateForReadytoDeploy(ByRef Common As Copient.CommonInc, ByVal OfferID As Long) As String
        Dim dt As DataTable
        Dim iWorkflowStatus As Integer = 0
        Dim ResponseMessage As String = ""
        Try
            If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()
            Common.QueryStr = "select WorkFlowStatus  from  Offers with (nolock)  Where OfferID=" & OfferID & " and Deleted=0"
            dt = Common.LRT_Select
            If dt.Rows.Count > 0 Then
                iWorkflowStatus = MyCommon.NZ(dt.Rows(0).Item("WorkflowStatus"), 0)
            End If

            If iWorkflowStatus <> 2 Then
                ResponseMessage = "The status can be changed to ready to deploy only if it is post validate"
            End If

            Return ResponseMessage
        Catch Ex As Exception
            Common.Write_Log(LogFile, Ex.Message, True)
            Return ResponseMessage
        End Try
    End Function

    Function AssignPreValidate(ByRef Common As Copient.CommonInc, ByVal OfferID As Long) As String

        Dim ResponseMessage As String = ""
        Dim StatusCode As Copient.LogixInc.STATUS_FLAGS = Copient.LogixInc.STATUS_FLAGS.STATUS_UNKNOWN
        Dim offerStatus, offerStatustext As String
        Dim bUseTestDates As Boolean

        Try
            If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()
            offerStatus = Logix.GetOfferStatus(OfferID, LanguageID, StatusCode)
            offerStatustext = Logix.GetOfferStatusText(StatusCode, LanguageID)
            bUseTestDates = (MyCommon.Fetch_SystemOption(88) = "1")

            Common.QueryStr = "update Offers with (RowLock) set WorkflowStatus=1 where OfferID=" & OfferID
            Common.LRT_Execute()
            Common.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("term.prevalidate", LanguageID))
            ResponseMessage = MyExport.SendWorkflowOutbound(OfferID, 1, AdminUserID, LanguageID)
            If offerStatustext.Trim.ToUpper <> "DEVELOPMENT" AndAlso offerStatustext.Trim.ToUpper <> "SCHEDULED" Then
                If offerStatustext.Trim.ToUpper = "TESTING" Then
                    If bUseTestDates Then
                        ResponseMessage = Copient.PhraseLib.Lookup("term.revalidationrequired", LanguageID)
                    End If
                Else
                    ResponseMessage = Copient.PhraseLib.Lookup("term.revalidationrequired", LanguageID)
                End If
            End If

            Return ResponseMessage
        Catch Ex As Exception
            Common.Write_Log(LogFile, Ex.Message, True)
            Return ResponseMessage
        End Try

    End Function

    Sub AssignPostValidate(ByRef Common As Copient.CommonInc, ByVal OfferID As Long)

        Dim ResponseMessage As String = ""

        Try
            If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()

            Common.QueryStr = "update Offers with (RowLock) set WorkflowStatus=2 where OfferID=" & OfferID
            Common.LRT_Execute()
            Common.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("term.postvalidate", LanguageID))
            ResponseMessage = MyExport.SendWorkflowOutbound(OfferID, 2, AdminUserID, LanguageID)
        Catch Ex As Exception
            Common.Write_Log(LogFile, Ex.Message, True)
        End Try

    End Sub

    Sub AssignReadytoDeploy(ByRef Common As Copient.CommonInc, ByVal OfferID As Long)

        Dim ResponseMessage As String = ""

        Try
            If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()

            Common.QueryStr = "update Offers with (RowLock) set WorkflowStatus=3 where OfferID=" & OfferID
            Common.LRT_Execute()
            Common.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("term.readytodeploy", LanguageID))
            ResponseMessage = MyExport.SendWorkflowOutbound(OfferID, 3, AdminUserID, LanguageID)
        Catch Ex As Exception
            Common.Write_Log(LogFile, Ex.Message, True)
        End Try

    End Sub

    Sub SendOutBound(ByRef Common As Copient.CommonInc, ByVal OfferID As Long)
        Dim iCRMType As Integer
        Dim rst As DataTable
        'Dim row As DataRow
        Try

            If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()
            If Not Integer.TryParse(Common.Fetch_SystemOption(25), iCRMType) Then iCRMType = 0

            Select Case iCRMType
                Case 1
                    ' Old Teradata CRM
                    Common.QueryStr = "update Offers with (RowLock) set CRMEngineUpdateLevel=CRMEngineUpdateLevel+1 where OfferID=" & OfferID
                    Common.LRT_Execute()

                    ' create an entry, if necessary, for use in TCRM agent processing
                    Common.QueryStr = "select LinkID from CRMEngineUpdateLevels with (NoLock) where EngineID=1 and ItemType=1 and LinkID=" & OfferID
                    rst = Common.LRT_Select
                    If rst.Rows.Count = 0 Then
                        Common.QueryStr = "insert into CRMEngineUpdateLevels with (RowLock) (EngineID, LinkID, ItemType, LastUpdateLevel, LastUpdate) " &
                                            "  values (1, " & OfferID & ",1,0,getdate());"
                        Common.LRT_Execute()
                    End If
                Case 2
                    ' CRM
                    Common.QueryStr = "update Offers with (RowLock) set LastCRMSendDate=getdate(), CRMEngineUpdateLevel=CRMEngineUpdateLevel+1, CRMSendToExport=1,CRMSendStatus=1 where OfferID=" & OfferID
                    Common.LRT_Execute()
                Case 3
                    ' All
                    ' Old Teradata CRM
                    Common.QueryStr = "update Offers with (RowLock) set CRMEngineUpdateLevel=CRMEngineUpdateLevel+1 where OfferID=" & OfferID
                    Common.LRT_Execute()

                    ' create an entry, if necessary, for use in TCRM agent processing
                    Common.QueryStr = "select LinkID from CRMEngineUpdateLevels with (NoLock) where EngineID=1 and ItemType=1 and LinkID=" & OfferID
                    rst = Common.LRT_Select
                    If rst.Rows.Count = 0 Then
                        Common.QueryStr = "insert into CRMEngineUpdateLevels with (RowLock) (EngineID, LinkID, ItemType, LastUpdateLevel, LastUpdate) " &
                                            "  values (1, " & OfferID & ",1,0,getdate());"
                        Common.LRT_Execute()
                    End If
                    ' CRM
                    Common.QueryStr = "update Offers with (RowLock) set LastCRMSendDate=getdate(), CRMEngineUpdateLevel=CRMEngineUpdateLevel+1, CRMSendToExport=1,CRMSendStatus=1 where OfferID=" & OfferID
                    Common.LRT_Execute()
            End Select
            Common.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-sendtocrm", LanguageID))
        Catch Ex As Exception
            Common.Write_Log(LogFile, Ex.Message, True)
        End Try
    End Sub

    Sub SendOutBoundCPE(ByRef Common As Copient.CommonInc, ByVal OfferID As Long)

        Try
            Common.QueryStr = "update CPE_Incentives with (RowLock) set LastCRMSendDate=getdate(), CRMEngineUpdateLevel=CRMEngineUpdateLevel+1, CRMSendToExport=1 where IncentiveID=" & OfferID
            Common.LRT_Execute()
            Common.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-sendtocrm", LanguageID))
        Catch Ex As Exception
            Common.Write_Log(LogFile, Ex.Message, True)
        End Try
    End Sub

    Sub CheckDefaultFolder(ByRef Common As Copient.CommonInc)
        Dim dt As DataTable
        MyCommon.QueryStr = "Select * from Folders where DefaultUEFolder=1"
        dt = MyCommon.LRT_Select()
        If dt.Rows.Count > 0 Then
            Send("1")
        Else
            Send("0")
        End If

    End Sub
</script>
