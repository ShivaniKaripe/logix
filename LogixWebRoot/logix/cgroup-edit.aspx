<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.CommonShared" %>
<%@ Import Namespace="CMS.AMS" %>
<%@ Import Namespace="CMS.AMS.Contract" %>

<%
    ' *****************************************************************************
    ' * FILENAME: cgroup-edit.aspx 
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
    Dim CustomerGroupID As Long
    Dim GName As String
    Dim CreatedDate As String
    Dim MyCommon As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    Dim MyCam As New Copient.CAM
    Dim MyCryptLib As New Copient.CryptLib
    Dim rst As DataTable
    Dim row As DataRow
    Dim rst2 As DataTable
    Dim row2 As DataRow
    Dim dt As DataTable
    Dim dt2 As DataTable
    Dim GroupSize As Integer
    Dim outputStatus As Integer
    Dim DefaultCustTypeID As Integer
    Dim CustTypeID As Integer
    Dim CardTypeID As Integer
    Dim CardTypeDesc As String = ""
    Dim File As HttpPostedFile
    Dim InstallPath As String
    Dim rowCount As Integer
    Dim deployDate As String
    Dim longDate As New DateTime
    Dim longDateString As String
    Dim lastUpdate As String
    Dim lastUpload As String
    Dim dst As System.Data.DataTable
    Dim AddAsHHChecked As String = ""
    Dim HouseholdingEnabled As Boolean
    Dim ClientUserID1 As String = ""
    Dim IDLength As Integer = 256 ' Allow maximum cardid
    Dim GNameTitle As String = ""
    Dim XID As String = ""
    Dim infoMessage As String = ""
    Dim infoMessage2 As String = ""
    Dim infoMessage3 As String = ""
    Dim ShowActionButton As Boolean = False
    Dim statusMessage As String = ""
    Dim Handheld As Boolean = False
    Dim OfferCtr As Integer = 0
    Dim IE6ScrollFix As String = ""
    Dim i As Integer = 0
    Dim OfferID As Integer = 0
    Dim EngineID As Integer = -1
    Dim CreatedFromOffer As Boolean = False
    Dim CAMInstalled As Boolean = False
    Dim CAMGroup As String = ""
    Dim IsCAMGroup As Boolean = False
    Dim CardIDs(-1) As String
    Dim CardCount As Integer = 0
    Dim EditControlTypeID As Integer = 0
    Dim RoleID As Integer = 0
    Dim disableButton As Boolean = False
    Dim RootURI As String = ""
    Dim IntegrationVals As New Copient.CommonInc.IntegrationValues
    Dim PrefPageName As String = ""
    Dim InUseByPreference As Boolean = False
    Dim AddHHToGroup As Boolean = False
    Dim CreateCustomer As Boolean = True
    Dim OptionText As String = ""
    Dim GroupSizeCustomers As Integer = 0
    Dim GMICount As Integer
    Dim IsOptInGroup As Boolean = False
    Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249)="1",True,False)
    Dim conditionalQuery = String.Empty
    Dim bCheckDigit As Boolean = False
    Dim bIsValidCard As Boolean = False
    Dim CMInstalled As Boolean = False
    Dim DisableRedeploy As Boolean = False
    Dim DualEngine As Boolean = False

    CurrentRequest.Resolver.AppName = "cgroup-edit.aspx"
    Dim m_Offer As IOffer = CurrentRequest.Resolver.Resolve(Of IOffer)()
    Dim analyticsCGService As CMS.AMS.Contract.IAnalyticsCustomerGroups = CMS.AMS.CurrentRequest.Resolver.Resolve(Of CMS.AMS.Contract.IAnalyticsCustomerGroups)()
    Dim isAnalyticsCG As Boolean = False

    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If

    Response.Expires = 0
    MyCommon.AppName = "cgroup-edit.aspx"

    MyCommon.Open_LogixXS()
    MyCommon.Open_LogixRT()
    If MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER, IntegrationVals) Then
        MyCommon.Open_PrefManRT()
    End If

    AdminUserID = Verify_AdminUser(MyCommon, Logix)

    DefaultCustTypeID = MyCommon.Fetch_SystemOption(30)
    AddAsHHChecked = IIf(DefaultCustTypeID = 1, " checked=""checked""", "")
    HouseholdingEnabled = IIf(MyCommon.Fetch_SystemOption(50) = 1, True, False)
    CAMInstalled = MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CAM)
    CMInstalled = MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)

    MyCommon.QueryStr = "select * FROM PromoEngines where Installed = 1"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 1 Then
        DualEngine = True
    End If

    If(CMInstalled And Not DualEngine) Then
        DisableRedeploy = True
    End If

    CustomerGroupID = Request.QueryString("CustomerGroupID")
    isAnalyticsCG = analyticsCGService.IsAnalyticsCustomerGroup(CustomerGroupID)
    GName = Request.QueryString("GroupName")
    If (Request.QueryString("CAMGroup") <> "") Then
        CAMGroup = Request.QueryString("CAMGroup")
    Else
        CAMGroup = "0"
    End If

    'If (HouseholdingEnabled) Then
    '  CustTypeID = IIf(Request.QueryString("clientusertype") = "", 0, MyCommon.Extract_Val(Request.QueryString("clientusertype")))
    'Else
    '  CustTypeID = DefaultCustTypeID
    'End If
    CardTypeID = IIf(Request.QueryString("clientusertype") = "", 0, MyCommon.Extract_Val(Request.QueryString("clientusertype")))
    MyCommon.QueryStr = "select CustTypeID from CardTypes with (NoLock) where CardTypeID=" & CardTypeID & ";"
    rst = MyCommon.LXS_Select
    If rst.Rows.Count > 0 Then
        CustTypeID = MyCommon.NZ(rst.Rows(0).Item("CustTypeID"), 0)
    End If

    OfferID = MyCommon.Extract_Val(Request.QueryString("OfferID"))
    EngineID = IIf(Request.QueryString("EngineID") = "", -1, MyCommon.Extract_Val(Request.QueryString("EngineID")))
    CreatedFromOffer = OfferID > 0 AndAlso EngineID > 0
    If CreatedFromOffer AndAlso CustomerGroupID = 0 Then
        GName = Copient.PhraseLib.Lookup("term.offer", LanguageID) & " " & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.group", LanguageID), VbStrConv.Lowercase)
        MyCommon.QueryStr = "select count(*) as GroupCount from CustomerGroups where Name like @Name + '%'"
        MyCommon.DBParameters.Add("@Name", SqlDbType.NVarChar, 255).Value = GName
        rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        If rst.Rows(0).Item("GroupCount") > 0 Then
            GName = GName & " (" & rst.Rows(0).Item("GroupCount") & ")"
        End If
    End If

    If (CustomerGroupID = 0 AndAlso Not Request.QueryString("save") <> "") Then
        If Not CreatedFromOffer Then
            GName = Request.Form("GroupName")
        End If
        CustomerGroupID = Request.Form("customerGroupID")
    End If

    If CustomerGroupID > 0 Then
        MyCommon.QueryStr = "select EditControlTypeID, RoleID from CustomerGroups with (NoLock) where CustomerGroupID=" & CustomerGroupID & ";"
        dt = MyCommon.LRT_Select
        If dt.Rows.Count > 0 Then
            EditControlTypeID = MyCommon.NZ(dt.Rows(0).Item("EditControlTypeID"), 0)
            RoleID = MyCommon.NZ(dt.Rows(0).Item("RoleID"), 0)
        End If
    End If

    InstallPath = MyCommon.Get_Install_Path(Request.PhysicalPath)

    'run the stored procedure to return all unfilled batch requests, if there is at least one open request, display an info message
    MyCommon.QueryStr = "select * from BarcodeBatchRequestQueue with (NoLock) where CustomerGroupID=" & CustomerGroupID & " AND RequestCompletedOn Is NULL;"
    dt2 = MyCommon.LXS_Select
    If (dt2.Rows.Count > 0 AndAlso MyCommon.Fetch_SystemOption(112)) Then
        infoMessage2 = Copient.PhraseLib.Lookup("cgroup-edit.GeneratingBarcodes", LanguageID)
        disableButton = True
    End If



    If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) Then
        MyCommon.QueryStr = "select CMOAStatusFlag from CustomerGroups with (NoLock) where CustomerGroupID=" & CustomerGroupID
        rst = MyCommon.LRT_Select()
        If (rst.Rows.Count > 0) Then
            If rst.Rows(0).Item(0) = "2" Then
                ' only show the awaiting deployment if there are Locations this group will be sent to.
                MyCommon.QueryStr = "select distinct LocationID from CustomerGroupLocUpdate with (NoLock) where EngineID=1 and CustomerGroupID=" & CustomerGroupID & ";"
                rst = MyCommon.LRT_Select
                If (rst.Rows.Count > 0) Then
                    statusMessage = Copient.PhraseLib.Lookup("confirm.redeployed", LanguageID)
                End If
            ElseIf rst.Rows(0).Item(0) = "-1" Then
                infoMessage = Copient.PhraseLib.Lookup("status.warning", LanguageID)
            End If
        End If
    End If

    If (Request.QueryString("ValidateBarcode") = "True") AndAlso disableButton = False Then
        'handle and validate barcodegeneration request
        Dim valid As Boolean = True
        Dim SVProgramID As Long = 0
        If Request.QueryString("SVProgramID") <> "" Then
            SVProgramID = MyCommon.Extract_Val(Request.QueryString("SVProgramID"))
        End If
        Dim LocationGroupID As Long = 0
        If Request.QueryString("LocationGroupID") <> "" Then
            LocationGroupID = MyCommon.Extract_Val(Request.QueryString("LocationGroupID"))
        End If
        Dim ValidLocation As Long = 0
        Dim UPC As String = ""
        UPC = Request.QueryString("UPC")
        Dim LocationID As Long = 0
        Dim ExtLocationID As String
        Dim ROID As Integer

        If Request.QueryString("LocationID") <> "" Then
            ExtLocationID = Request.QueryString("LocationID")
        End If
        Dim RedemptionRestrictionID As Integer = 0
        If Request.QueryString("RedemptionRestrictionID") <> "" Then
            RedemptionRestrictionID = MyCommon.Extract_Val(Request.QueryString("RedemptionRestrictionID"))
        End If
        Dim NumOfBarcodes As Integer
        If Request.QueryString("NumOfBarcodes") <> "" Then
            NumOfBarcodes = MyCommon.Extract_Val(Request.QueryString("NumOfBarcodes"))
            If NumOfBarcodes = 0 Then
                infoMessage = Copient.PhraseLib.Lookup("customer-edit.invalidnumofbarcodes", LanguageID)
                valid = False
            End If
        Else
            NumOfBarcodes = 1
        End If

        If (RedemptionRestrictionID = 2) Then
            MyCommon.QueryStr = "select LocationGroupID from LocationGroups where LocationGroupID = '" & LocationGroupID & "' and Deleted ='False';"
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count > 0) Then
                row = rst.Rows(0)
                LocationGroupID = row.Item("LocationGroupID")
            Else
                infoMessage3 = Copient.PhraseLib.Lookup("customer-edit.UnableToValidateLocation", LanguageID)
                valid = False
            End If
        ElseIf (RedemptionRestrictionID = 1) Then
            MyCommon.QueryStr = "select LocationID from Locations where ExtLocationCode = '" & ExtLocationID & "' and Deleted = 'False';"
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count > 0) Then
                row = rst.Rows(0)
                LocationID = row.Item("LocationID")
            Else
                infoMessage3 = Copient.PhraseLib.Lookup("customer-edit.UnableToValidateLocation", LanguageID)
                valid = False
            End If
        ElseIf (RedemptionRestrictionID = 0) Then
            LocationID = 0
            LocationGroupID = 0
        Else
            infoMessage3 = Copient.PhraseLib.Lookup("customer-edit.UnableToValidateLocation", LanguageID)
            valid = False
        End If

        If (SVProgramID = 0 OrElse Request.QueryString("SVProgramID") = "") Then
            infoMessage3 = Copient.PhraseLib.Lookup("customer-edit.UnableToValidateSV", LanguageID)
            valid = False
        End If

        If (UPC = "" OrElse Len(UPC) < 10 Or Len(UPC) > 10 Or Not IsNumeric(UPC)) Then
            infoMessage3 = Copient.PhraseLib.Lookup("customer-edit.UnableToValidateUPC", LanguageID)
            valid = False
        End If

        If Request.QueryString("OfferID") <> "" Then
            MyCommon.QueryStr = "select RewardOptionID from CPE_RewardOptions where IncentiveID = " & MyCommon.Extract_Val(Request.QueryString("OfferID"))
            rst = MyCommon.LRT_Select
            If rst.Rows.Count > 0 Then
                ROID = MyCommon.NZ(rst.Rows(0).Item("RewardOptionID"), -1)
            Else
                infoMessage3 = "Invalid Offer ID"
                valid = False
            End If
        Else
            ROID = -1
        End If

        If (valid) Then

            If (RedemptionRestrictionID = 2) Then
                ValidLocation = LocationGroupID
            ElseIf (RedemptionRestrictionID = 1) Then
                ValidLocation = LocationID
            Else
                ValidLocation = 0
            End If

            MyCommon.QueryStr = "INSERT INTO BarcodeBatchRequestQueue with (RowLock) (CustomerGroupID,SVProgramID,UPC," & _
               "RequestedOn,ValidLocation,RedemptionRestrictionID,NumberOfBarcodes, ROID) values(" & CustomerGroupID & "," & SVProgramID & "," & UPC & _
               "," & "getDate()," & ValidLocation & "," & RedemptionRestrictionID & "," & NumOfBarcodes & "," & IIf(ROID > -1, ROID, "NULL") & ");"
            MyCommon.LXS_Execute()
            Response.Redirect("cgroup-edit.aspx?CustomerGroupID=" & CustomerGroupID)
        End If

    End If

    If (Request.QueryString("new") <> "") Then
        Response.Redirect("cgroup-edit.aspx")
    End If

    If (Request.QueryString("download") <> "") Then
        MyCommon.QueryStr = "select Name from CustomerGroups with (NoLock) where CustomerGroupID=" & CustomerGroupID & " and deleted =0"
        rst = MyCommon.LRT_Select()
        For Each row In rst.Rows
            If (GName = "") Then GName = MyCommon.NZ(row.Item("Name"), "")
        Next
        MyCommon.QueryStr = "select C.ExtCardIDOriginal as ExtCardID, CT.ExtCardTypeID from CardIDs  as C with (NoLock) " & _
                            "inner join CardTypes as CT with (NoLock) on CT.CardTypeID = C.CardTypeID " & _
                            "where C.CustomerPK in (select CustomerPK from GroupMembership with (NoLock) where CustomerGroupID=" & CustomerGroupID & " and Deleted =0)"
        rst = MyCommon.LXS_Select()
        If (rst.Rows.Count > 0) Then
            Response.Clear()
            Response.AddHeader("Content-Disposition", "attachment; filename=CG" & CustomerGroupID & ".txt")
            Response.ContentType = "application/octet-stream"
            For Each row In rst.Rows
                Sendb(IIf(IsDBNull(row.Item("ExtCardID")), Copient.PhraseLib.Lookup("term.unknown", LanguageID), MyCryptLib.SQL_StringDecrypt(row.Item("ExtCardID").ToString())))
                Sendb(",")
                Send(MyCommon.NZ(row.Item("ExtCardTypeID"), 0))
            Next
            GoTo done
        Else
            infoMessage = Copient.PhraseLib.Lookup("cgroup-edit.noelements", LanguageID)
        End If
    End If

    If Request.QueryString("LargeFile") = "true" Then
        infoMessage = Copient.PhraseLib.Lookup("error.UploadTooLarge", LanguageID)
    End If

    If Request.Files.Count >= 1 Then
        File = Request.Files.Get(0)
        If File.ContentType <> "text/plain" Then
            infoMessage = Copient.PhraseLib.Lookup("cgroup-edit.badfile", LanguageID)
        Else
            Dim UploadFileName As String
            Dim TimeStampStr As String
            TimeStampStr = MyCommon.Leading_Zero_Fill(Day(Date.Now), 2) & MyCommon.Leading_Zero_Fill(Hour(Date.Now), 2) & MyCommon.Leading_Zero_Fill(Minute(Date.Now), 2) & MyCommon.Leading_Zero_Fill(Second(Date.Now), 2)
            UploadFileName = "U" & CustomerGroupID & "-" & TimeStampStr & ".dat"
            File.SaveAs(MyCommon.Fetch_SystemOption(29) & "\" & UploadFileName)
            MyCommon.QueryStr = "dbo.pt_GMInsertQueue_Insert"
            MyCommon.Open_LXSsp()
            MyCommon.LXSsp.Parameters.Add("@FileName", SqlDbType.NVarChar, 255).Value = UploadFileName
            MyCommon.LXSsp.Parameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = CustomerGroupID
            If MyCommon.Extract_Val(Request.Form("format")) = 1 Then
                MyCommon.LXSsp.Parameters.Add("@CardTypeID", SqlDbType.Int).Value = MyCommon.Extract_Val(Request.Form("cardtypeid"))
            ElseIf MyCommon.Extract_Val(Request.Form("format")) = 2 Then
                MyCommon.LXSsp.Parameters.Add("@CardTypeID", SqlDbType.Int).Value = -1
            End If
            MyCommon.LXSsp.Parameters.Add("@OperationType", SqlDbType.Int).Value = MyCommon.Extract_Val(Request.Form("operation"))
            MyCommon.LXSsp.ExecuteNonQuery()
            MyCommon.Activity_Log(4, CustomerGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.cgroup-upload", LanguageID))
            MyCommon.Close_LXSsp()
            Response.Status = "301 Moved Permanently"
            If Not CreatedFromOffer Then
                Response.AddHeader("Location", "cgroup-edit.aspx?CustomerGroupID=" & CustomerGroupID)
            Else
                Select Case EngineID
                    Case 7
                        Response.AddHeader("Location", "cgroup-edit.aspx?CustomerGroupID=" & CustomerGroupID & _
                                           "&OfferID=" & OfferID & "&EngineID=" & EngineID & "&slct=" & Request.QueryString("slct") & _
                                           "&ex=" & Request.QueryString("ex") & "&condChanged=" & Request.QueryString("condChanged"))
                    Case Else
                        Response.AddHeader("Location", "cgroup-edit.aspx?CustomerGroupID=" & CustomerGroupID)
                End Select
            End If
            GoTo done
        End If
    End If

    If (Request.QueryString("save") <> "") OrElse (CreatedFromOffer AndAlso CustomerGroupID = 0) Then
        If (CustomerGroupID = 0) Then
            MyCommon.QueryStr = "dbo.pt_CustomerGroups_Insert"
            MyCommon.Open_LRTsp()
            GName = Logix.TrimAll(GName)
            MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 200).Value = GName
            MyCommon.LRTsp.Parameters.Add("@CustomerGroupID", SqlDbType.BigInt).Direction = ParameterDirection.Output
            MyCommon.LRTsp.Parameters.Add("@CAMCustomerGroup", SqlDbType.Bit).Value = Val(CAMGroup)
            MyCommon.LRTsp.Parameters.Add("@EditControlTypeID", SqlDbType.Int).Value = EditControlTypeID
            MyCommon.LRTsp.Parameters.Add("@RoleID", SqlDbType.Int).Value = RoleID
            If (GName = "") Then
                infoMessage = Copient.PhraseLib.Lookup("cgroup-edit.noname", LanguageID)
            Else
                MyCommon.QueryStr = "SELECT CustomerGroupID FROM CustomerGroups WITH (NoLock) WHERE Name='" & MyCommon.Parse_Quotes(GName) & "' AND Deleted=0;"
                dst = MyCommon.LRT_Select
                If (dst.Rows.Count > 0) Then
                    infoMessage = Copient.PhraseLib.Lookup("cgroup-edit.nameused", LanguageID)
                Else
                    MyCommon.LRTsp.ExecuteNonQuery()
                    CustomerGroupID = MyCommon.LRTsp.Parameters("@CustomerGroupID").Value
                    MyCommon.Activity_Log(4, CustomerGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.cgroup-create", LanguageID))
                End If
            End If
            MyCommon.Close_LRTsp()
            If infoMessage = "" AndAlso CustomerGroupID <> 0 Then
                Response.Status = "301 Moved Permanently"
                Response.AddHeader("Location", "cgroup-edit.aspx?CustomerGroupID=" & CustomerGroupID)
            End If
        Else
            If (Request.QueryString("editcontroltypeid") <> "") Then
                EditControlTypeID = MyCommon.Extract_Val(Request.QueryString("editcontroltypeid"))
            End If
            If (Request.QueryString("roleid") <> "") Then
                RoleID = MyCommon.Extract_Val(Request.QueryString("roleid"))
            End If
            If (Request.QueryString("IsOptinGroup") <> "") Then
                IsOptInGroup = Request.QueryString("IsOptinGroup")
            End If
            MyCommon.QueryStr = "dbo.pt_CustomerGroups_Update"
            MyCommon.Open_LRTsp()
            GName = Logix.TrimAll(GName)
            MyCommon.LRTsp.Parameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = CustomerGroupID
            MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 200).Value = Logix.TrimAll(GName)
            MyCommon.LRTsp.Parameters.Add("@EditControlTypeID", SqlDbType.Int).Value = EditControlTypeID
            MyCommon.LRTsp.Parameters.Add("@RoleID", SqlDbType.Int).Value = RoleID
            MyCommon.LRTsp.Parameters.Add("@IsOptInGroup", SqlDbType.Bit).Value = IsOptInGroup
            If (GName = "") Then
                infoMessage = Copient.PhraseLib.Lookup("cgroup-edit.noname", LanguageID)
            Else
                MyCommon.QueryStr = "SELECT Name, CustomerGroupID FROM CustomerGroups WITH (NoLock) WHERE Name='" & MyCommon.Parse_Quotes(GName) & "' AND Deleted=0 AND CustomerGroupID<>" & CustomerGroupID & ";"
                dst = MyCommon.LRT_Select
                If (dst.Rows.Count > 0) Then
                    infoMessage = Copient.PhraseLib.Lookup("cgroup-edit.nameused", LanguageID)
                Else
                    MyCommon.LRTsp.ExecuteNonQuery()
                    MyCommon.Activity_Log(4, CustomerGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.cgroup-edit", LanguageID))
                End If
            End If
            MyCommon.Close_LRTsp()
            If infoMessage = "" Then
                Response.Status = "301 Moved Permanently"
                Response.AddHeader("Location", "cgroup-edit.aspx?CustomerGroupID=" & CustomerGroupID)
            End If
        End If
    ElseIf (Request.QueryString("redeploy") <> "") Then
        Dim SetFlags As String = ""
        ' The user wants to redeploy, so do a quick check on offerRewards to make sure there are at least some rewards
        If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) Then SetFlags = " CMOAStatusFlag=2"
        If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE)) Then
            If (SetFlags <> "") Then SetFlags = SetFlags & ","
            SetFlags = SetFlags & " CPEStatusFlag=2"
        End If
        If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE)) Then
            If (SetFlags <> "") Then SetFlags = SetFlags & ","
            SetFlags = SetFlags & " UEStatusFlag=2"
        End If

        MyCommon.QueryStr = "update customergroups with (RowLock) set " & SetFlags & " , updatelevel=updatelevel+1 where CustomerGroupID=" & CustomerGroupID
        MyCommon.LRT_Execute()
        statusMessage = Copient.PhraseLib.Lookup("confirm.redeployed", LanguageID)
        MyCommon.Activity_Log(4, CustomerGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.cgroup-redeploy", LanguageID))

        Response.Status = "301 Moved Permanently"
        Response.AddHeader("Location", "cgroup-edit.aspx?CustomerGroupID=" & CustomerGroupID & "&statMsg=" & statusMessage)

    ElseIf (Request.QueryString("delete") <> "") Then
        If MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER, IntegrationVals) Then
            ' check if any custom preferences use this as a targeted customer group.
            MyCommon.QueryStr = "select PCG.PreferenceID, PREF.Name, PREF.NamePhraseID, PREF.UserCreated, " & _
                                "  case when UPT.Phrase is null then PREF.Name else Convert(nvarchar(200), UPT.Phrase) end as PhrasedName " & _
                                "from PreferenceCustomerGroups as PCG with (NoLock) " & _
                                "inner join Preferences as PREF with (NoLock) on PREF.PreferenceID = PCG.PreferenceID " & _
                                "left join UserPhraseText as UPT with (NoLock) on UPT.PhraseID = PREF.NamePhraseID " & _
                                "where PCG.CustomerGroupID=" & CustomerGroupID & " and PREF.Deleted = 0;"
            rst2 = MyCommon.PMRT_Select
            InUseByPreference = (rst2.Rows.Count > 0)
        End If

        MyCommon.QueryStr = "select distinct 1 as EngineID, O.OfferID,O.Name,O.ProdEndDate from offerconditions as OC with (NoLock) " & _
                            "  left join Offers as O with (NoLock) on O.Offerid=OC.offerID " & _
                            "  where (OC.linkid = @CustomerGroupID or OC.excludedid = @CustomerGroupID) and OC.ConditionTypeID=1 and O.deleted=0 and O.IsTemplate=0 and OC.Deleted=0 " & _
                            " UNION " & _
                            " select distinct 1 as EngineID, O.OfferID,O.Name,O.ProdEndDate from offerrewards as OFR with (NoLock) " & _
                            "  left join Offers as O with (NoLock) on O.Offerid=OFR.offerID " & _
                            "  left join RewardCustomerGroupTiers as RCGT with (NoLock) on RCGT.RewardID=OFR.RewardID " & _
                            "  where (RCGT.CustomerGroupID = @CustomerGroupID) and (OFR.RewardTypeID=5 or OFR.RewardTypeID=6) and O.deleted=0 and O.IsTemplate=0 and OFR.Deleted=0 " & _
                            " UNION " & _
                            " SELECT DISTINCT 2 as EngineID, I.IncentiveID as OfferID, I.IncentiveName as Name, I.EndDate as ProdEndDate from CPE_IncentiveCustomerGroups ICG with (NoLock) " & _
                            "  INNER JOIN CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID=ICG.RewardOptionID " & _
                            "  INNER JOIN CPE_Incentives I with (NoLock) on I.IncentiveID = RO.IncentiveID " & _
                            "  INNER JOIN CustomerGroups CG with (NoLock) on ICG.CustomerGroupID = CG.CustomerGroupID " & _
                            "  WHERE ICG.CustomerGroupID = @CustomerGroupID and ICG.Deleted = 0 and RO.Deleted=0 and I.Deleted=0 and I.IsTemplate=0 and CG.Deleted=0 " & _
                            " UNION " & _
                            " SELECT DISTINCT 2 as EngineID, I.IncentiveID as OfferID, I.IncentiveName as Name, I.EndDate as ProdEndDate from CPE_Deliverables D with (NoLock) " & _
                            "  INNER JOIN CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID=D.RewardOptionID " & _
                            "  INNER JOIN CPE_Incentives I with (NoLock) on I.IncentiveID = RO.IncentiveID " & _
                            "  WHERE D.Deleted=0 and RO.Deleted=0 and I.Deleted=0 and I.IsTemplate=0 and D.DeliverableTypeId IN (5,6) and OutputID = @CustomerGroupID" & _
                            " UNION " & _
                            " SELECT I.EngineID, I.IncentiveID as OfferID, I.IncentiveName as Name, I.EndDate as ProdEndDate FROM CPE_Incentives I with (NoLock) " & _
                            "  INNER JOIN OfferEligibilityConditions OEC ON I.IncentiveID = OEC.OfferID " & _
                            "  INNER JOIN Conditions C ON C.ConditionID = OEC.ConditionID " & _
                            "  INNER JOIN CustomerConditionDetails CCD ON C.ConditionID = CCD.ConditionID " & _
                            " WHERE OEC.Deleted = 0 AND I.Deleted = 0 AND CCD.CustomerGroupID = @CustomerGroupID" & _
                            " order by Name "
        MyCommon.DBParameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = CustomerGroupID
        rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        If rst.Rows.Count = 0 Then
            If (CustomerGroupID <= 2) Then
                infoMessage = Copient.PhraseLib.Lookup("cgroup-edit.nodelete", LanguageID)
            ElseIf InUseByPreference Then
                infoMessage = Copient.PhraseLib.Lookup("cgroup-edit.inuse", LanguageID) & " (" & Copient.PhraseLib.Lookup("term.preference", LanguageID) & " '" & MyCommon.NZ(rst2.Rows(0).Item("Name"), "") & "')"
            Else
                ' check that there are no deployed offers that use this customer group
                MyCommon.QueryStr = "dbo.pa_AssociatedOffers_ST"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@LinkType", SqlDbType.Int).Value = 1
                MyCommon.LRTsp.Parameters.Add("@LinkID", SqlDbType.Int).Value = CustomerGroupID
                rst2 = MyCommon.LRTsp_select
                MyCommon.Close_LRTsp()
                If (rst2.Rows.Count > 0) Then
                    infoMessage = Copient.PhraseLib.Lookup("term.inusedeployment", LanguageID) & " : ("
                    For OfferCtr = 0 To rst2.Rows.Count - 1
                        infoMessage &= MyCommon.NZ(rst2.Rows(OfferCtr).Item("IncentiveID"), "")
                    Next
                    infoMessage &= ")"
                Else
                    MyCommon.QueryStr = "dbo.pt_CustomerGroups_Delete"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = CustomerGroupID
                    MyCommon.LRTsp.ExecuteNonQuery()
                    MyCommon.Close_LRTsp()
                    MyCommon.Activity_Log(4, CustomerGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.cgroup-delete", LanguageID))
                    Response.Status = "301 Moved Permanently"
                    Response.AddHeader("Location", "cgroup-list.aspx")
                    CustomerGroupID = 0
                    GName = ""
                End If

            End If
        Else
            infoMessage = Copient.PhraseLib.Lookup("cgroup-edit.inuse", LanguageID)
        End If
    ElseIf (Request.QueryString("add") <> "") Then
        'If(CardTypeID = 6) Then
        ClientUserID1 = Request.QueryString("clientuserid1")
        Dim CustomerPK As Long
        Dim dtTemp As DataTable
        'End If
        If (ClientUserID1 = "") Then
            If (CardTypeID = 4) Then
                infoMessage = Copient.PhraseLib.Lookup("cgroup-edit.UsernameAlphanumeric", LanguageID)
            ElseIf (CardTypeID = 6) Then
                infoMessage = Copient.PhraseLib.Lookup("cgroup-edit.EmailIdAlphanumeric", LanguageID)
            Else
                infoMessage = Copient.PhraseLib.Lookup("cgroup-edit.badid", LanguageID)
            End If
        End If
        ' check if this is a CAM customer group, if so, then validate the card number
        MyCommon.QueryStr = "select CAMCustomerGroup from CustomerGroups with (NoLock) where CustomerGroupID=" & CustomerGroupID & " and deleted=0 and IsNull(CAMCustomerGroup,0) = 1"
        rst = MyCommon.LRT_Select()
        If rst.Rows.Count > 0 Then
            If MyCommon.NZ(rst.Rows(0).Item("CAMCustomerGroup"), False) Then
                ClientUserID1 = MyCommon.Pad_ExtCardID(MyCommon.Extract_Val(Request.QueryString("clientuserid1")), 2)
                If MyCam.VerifyCardNumber(ClientUserID1, infoMessage) Then
                    CustTypeID = 2
                    CardTypeID = 2
                Else
                    infoMessage = Copient.PhraseLib.Detokenize("cgroup-edit.InvalidCAMCard", LanguageID, Request.QueryString("clientuserid1"))  'Card number {0} is not in a valid CAM card number format.
                End If
            Else
                If (MyCommon.IsEngineInstalled(6)) Then
                    infoMessage = ""
                    If MyCam.VerifyCardNumber(ClientUserID1, infoMessage) Then
                        infoMessage = Copient.PhraseLib.Detokenize("cgroup-edit.InvalidCAMCard", LanguageID, Request.QueryString("clientuserid1"))  'Card number {0} is not valid because it is in the CAM card number format.
                    Else
                        infoMessage = ""
                    End If
                End If
            End If
        End If
        ClientUserID1 = MyCommon.Pad_ExtCardID(Request.QueryString("clientuserid1"), CardTypeID)
        If ClientUserID1 = "" Then
            infoMessage = Copient.PhraseLib.Lookup("inputValidator.blankcustid", LanguageID)
        Else
            MyCommon.QueryStr = "select CustomerPK from CardIDs with (NoLock) where ExtCardID=@ClientUserID1 and CardTypeID=@CardTypeID;"
            MyCommon.DBParameters.Add("@ClientUserID1", SqlDbType.NVarChar, 400).Value = MyCryptLib.SQL_StringEncrypt(ClientUserID1)
            MyCommon.DBParameters.Add("@CardtypeID", SqlDbType.Int).Value = CardTypeID
            rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)
            CreateCustomer = (MyCommon.Fetch_InterfaceOption(11) = "1")
            If (rst.Rows.Count = 0) AndAlso (Logix.UserRoles.CreateCustomer = False) Then
                infoMessage = Copient.PhraseLib.Lookup("cgroup-edit.nocreate", LanguageID)
            ElseIf (rst.Rows.Count = 0) AndAlso (CreateCustomer = False) Then
                infoMessage = Copient.PhraseLib.Lookup("cgroup-edit.cardnotexist", LanguageID)
            End If
        End If
        AddHHToGroup = MyCommon.Fetch_SystemOption(138)
        bCheckDigit = MyCommon.Fetch_SystemOption(227)
        If (bCheckDigit AndAlso CardTypeID = 0) Then
            bIsValidCard = MyCommon.RewardCardCheckDigit(ClientUserID1)
            If Not bIsValidCard Then
                infoMessage = Copient.PhraseLib.Lookup("customer.invalidcard", LanguageID) & " (" & ClientUserID1 & ")"
            End If
        End If
        Dim AutoAddCards As Boolean = MyCommon.Fetch_SystemOption(146)
        If infoMessage = "" Then
            MyCommon.QueryStr = "dbo.pt_GroupMembership_Insert"
            MyCommon.Open_LXSsp()
            MyCommon.LXSsp.Parameters.Add("@ExtCardID", SqlDbType.NVarChar, 400).Value = MyCryptLib.SQL_StringEncrypt(ClientUserID1, True )
            MyCommon.LXSsp.Parameters.Add("@CardTypeID", SqlDbType.Int).Value = CardTypeID
            MyCommon.LXSsp.Parameters.Add("@CustomerTypeID", SqlDbType.Int).Value = CustTypeID
            MyCommon.LXSsp.Parameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = CustomerGroupID
            MyCommon.LXSsp.Parameters.Add("@AddHHToGroup", SqlDbType.Bit).Value = AddHHToGroup
            MyCommon.LXSsp.Parameters.Add("@AutoAddCardOption", SqlDbType.Bit).Value = AutoAddCards
            MyCommon.LXSsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
            MyCommon.LXSsp.Parameters.Add("@ExtCardIDOriginal", SqlDbType.NVarChar, 400).Value= MyCryptLib.SQL_StringEncrypt(ClientUserID1, False)
            If (MyCommon.AllowToProcessCustomerCard(ClientUserID1, CardTypeID, Nothing) = False) Then
                infoMessage = Copient.PhraseLib.Lookup("cgroup-edit.badid", LanguageID)
            Else If CardTypeID = 6 Then
                If (Not MyCommon.EmailAddressCheck(ClientUserID1)) Then
                    infoMessage = Copient.PhraseLib.Lookup("emailValidation", LanguageID)
                End If
            End If
            If infoMessage = "" Then
                MyCommon.LXSsp.ExecuteNonQuery()
                outputStatus = MyCommon.LXSsp.Parameters("@Status").Value
                If (outputStatus = -1) Then
                    infoMessage = Copient.PhraseLib.Lookup("cgroup-edit.duplicateid", LanguageID) & " (" & ClientUserID1 & ")"
                Else
                    CardIDs = FindAllCustomerCards(ClientUserID1, CardTypeID, MyCommon)
                    statusMessage = GetCardMessage(CardIDs, 1)
                    Dim MyLookup As New Copient.CustomerLookup()
                    MyLookup.SetLanguageID(LanguageID)
                    CardTypeDesc = MyLookup.findCardTypeDescription(CardTypeID)
                    MyCommon.Activity_Log(4, CustomerGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.cgroup-add", LanguageID) & " " & ClientUserID1 & IIf(CardTypeDesc <> "", " " & CardTypeDesc, ""))
                    '' If its a default customer group then add data in Activity_Log for CUstomer History Tab  --- START
                    MyCommon.QueryStr = "SELECT * FROM CustomerGroups WHERE IsOptinGroup = 1 AND CustomerGroupID = @CustomerGroupID"
                    MyCommon.DBParameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = CustomerGroupID
                    dtTemp = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                    If dtTemp.Rows.Count > 0 Then
                        Dim dtTemp2 As DataTable
                        MyCommon.QueryStr = "SELECT CustomerPK FROM CardIDs with (NoLock) WHERE ExtCardID=@ExtCardID AND CardTypeID=@CardTypeID"
                        MyCommon.DBParameters.Add("@ExtCardID", SqlDbType.NVarChar).Value = MyCryptLib.SQL_StringEncrypt(ClientUserID1)
                        MyCommon.DBParameters.Add("@CardTypeID", SqlDbType.Int).Value = CardTypeID
                        dtTemp2 = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)
                        If dtTemp2.Rows.Count > 0 Then
                            CustomerPK = Convert.ToInt64(dtTemp2(0).Item("CustomerPK"))
                        End If

                        Dim dtTemp3 As DataTable
                        MyCommon.QueryStr = "SELECT RO.IncentiveID AS OfferID, RO.Name AS OfferName FROM CPE_RewardOptions RO INNER JOIN CPE_Incentives I ON RO.IncentiveID = I.IncentiveID " & _
                                             "INNER JOIN CPE_IncentiveCustomerGroups ICG ON RO.RewardOptionID = ICG.RewardOptionID WHERE(ICG.CustomerGroupID = @CustomerGroupID)" & _
                                             "UNION SELECT O.OfferID, [Name] AS OfferName FROM Offers O INNER JOIN OfferConditions OC ON O.OfferID = OC.OfferID WHERE(OC.LinkID = @CustomerGroupID)"
                        MyCommon.DBParameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = CustomerGroupID
                        dtTemp3 = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                        If dtTemp3.Rows.Count > 0 Then
                            Dim objOfferID As Long
                            Dim objOfferName As String
                            objOfferID = Convert.ToInt64(dtTemp3(0).Item("OfferID"))
                            objOfferName = dtTemp3(0).Item("OfferName").ToString()
                            MyCommon.Activity_Log(25, CustomerPK, AdminUserID, Copient.PhraseLib.Detokenize("term.customeroptedin", 1, objOfferID, objOfferName))
                        End If
                    End If
                    '' If its a default customer group then add data in Activity_Log for CUstomer History Tab  --- END
                End If
            End If

            'End If
            MyCommon.Close_LXSsp()
            MyCommon.QueryStr = "update CustomerGroups with (RowLock) set LastUpdate=getdate() where CustomerGroupID=" & CustomerGroupID & ";"
            MyCommon.LRT_Execute()
            SendNotificationsOfItemChange(CustomerGroupID, 1)
        End If
        If infoMessage = "" Then
            Response.Status = "301 Moved Permanently"
            Response.AddHeader("Location", "cgroup-edit.aspx?CustomerGroupID=" & CustomerGroupID)
        End If
    ElseIf (Request.QueryString("remove") <> "") Then
        If (Request.QueryString("membershipid") <> "") Then
            MyCommon.Open_LogixRT()
            Dim CustomerPK As Long
            ReDim CardIDs(Request.QueryString.GetValues("membershipid").GetUpperBound(0))
            For i = 0 To Request.QueryString.GetValues("membershipid").GetUpperBound(0)
                MyCommon.QueryStr = "dbo.pt_GroupMembership_Delete_ByID"
                MyCommon.Open_LXSsp()
                MyCommon.LXSsp.Parameters.Add("@MembershipID", SqlDbType.BigInt).Value = IIf(Request.QueryString.GetValues("membershipid")(i) <> "", Request.QueryString.GetValues("membershipid")(i), -1)
                MyCommon.LXSsp.ExecuteNonQuery()
                MyCommon.Close_LXSsp()
                MyCommon.QueryStr = "select ExtCardID, CustomerPK from CardIDs with (NoLock) where CustomerPK in " & _
                                    "  (select CustomerPK from GroupMembership with (NoLock) " & _
                                    "   where MembershipID=" & IIf(Request.QueryString.GetValues("membershipid")(i) <> "", Request.QueryString.GetValues("membershipid")(i), -1) & ");"

                rst = MyCommon.LXS_Select
                If (rst.Rows.Count > 0) Then
                    Dim dtTemp As DataTable
                    CardIDs(i) = IIf(IsDBNull(rst.Rows(0).Item("ExtCardID")), 0, MyCryptLib.SQL_StringDecrypt(rst.Rows(0).Item("ExtCardID").ToString()))
                    CustomerPK = MyCommon.NZ(rst.Rows(0).Item("CustomerPK"), 0)
                    MyCommon.QueryStr = "SELECT * FROM CustomerGroups WHERE IsOptinGroup = 1 AND CustomerGroupID = @CustomerGroupID"
                    MyCommon.DBParameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = CustomerGroupID
                    dtTemp = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                    If dtTemp.Rows.Count > 0 Then
                        Dim dtTemp3 As DataTable
                        MyCommon.QueryStr = "SELECT RO.IncentiveID AS OfferID, RO.Name AS OfferName FROM CPE_RewardOptions RO INNER JOIN CPE_Incentives I ON RO.IncentiveID = I.IncentiveID " & _
                                             "INNER JOIN CPE_IncentiveCustomerGroups ICG ON RO.RewardOptionID = ICG.RewardOptionID WHERE(ICG.CustomerGroupID = @CustomerGroupID)" & _
                                             "UNION SELECT O.OfferID, [Name] AS OfferName FROM Offers O INNER JOIN OfferConditions OC ON O.OfferID = OC.OfferID WHERE(OC.LinkID = @CustomerGroupID)"
                        MyCommon.DBParameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = CustomerGroupID
                        dtTemp3 = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                        If dtTemp3.Rows.Count > 0 Then
                            Dim objOfferID As Long
                            Dim objOfferName As String
                            objOfferID = Convert.ToInt64(dtTemp3(0).Item("OfferID"))
                            objOfferName = dtTemp3(0).Item("OfferName").ToString()
                            MyCommon.Activity_Log(25, CustomerPK, AdminUserID, Copient.PhraseLib.Detokenize("term.customeroptedout", 1, objOfferID, objOfferName))
                        End If
                    End If

                End If
                'If (rst.Rows.Count > 0) Then
                '  ReDim CardIDs(rst.Rows.Count - 1)
                '  For j = 0 To CardIDs(j)
                '    CardIDs(j) = rst.Rows(j).Item("ExtCardID")
                '  Next
                '  statusMessage = GetCardMessage(CardIDs, 2)
                '  MyCommon.Activity_Log(4, CustomerGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.cgroup-remove", LanguageID) & " " & String.Join(",", CardIDs))
                'Else
                '  MyCommon.Activity_Log(4, CustomerGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.cgroup-remove", LanguageID) & " " & Left(Request.QueryString("clientuserid1"), 26))
                'End If
            Next
            If CardIDs.Length > 0 Then
                statusMessage = GetCardMessage(CardIDs, 2)
                MyCommon.Activity_Log(4, CustomerGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.cgroup-remove", LanguageID) & " " & String.Join(", ", CardIDs))
            Else
                MyCommon.Activity_Log(4, CustomerGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.cgroup-remove", LanguageID) & " " & Left(Request.QueryString("clientuserid1"), 26))
            End If

            MyCommon.QueryStr = "Update CustomerGroups with (RowLock) set LastUpdate=getdate() where CustomerGroupID=" & CustomerGroupID
            MyCommon.LRT_Execute()
            SendNotificationsOfItemChange(CustomerGroupID, 1)
        Else
            infoMessage = Copient.PhraseLib.Lookup("cgroup-edit.idnotselected", LanguageID)
        End If
        Response.Status = "301 Moved Permanently"
        Response.AddHeader("Location", "cgroup-edit.aspx?CustomerGroupID=" & CustomerGroupID)
    ElseIf (Request.QueryString("mremove") <> "") Then

        ClientUserID1 = MyCommon.Pad_ExtCardID(Request.QueryString("clientuserid1"), CardTypeID)

        Dim dtCustPK As DataTable
        Dim CustomerPK As Long = 0
        MyCommon.QueryStr = "select CustomerPK from CardIDs with (NoLock) where ExtCardID=@ClientUserID1 and CardTypeID=@CardTypeID;"
        MyCommon.DBParameters.Add("@ClientUserID1", SqlDbType.NVarChar, 400).Value = MyCryptLib.SQL_StringEncrypt(ClientUserID1)
        MyCommon.DBParameters.Add("@CardtypeID", SqlDbType.Int).Value = CardTypeID
        dtCustPK = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)
        If dtCustPK.Rows.Count > 0 Then CustomerPK = MyCommon.NZ(dtCustPK.Rows(0).Item("CustomerPK"), 0)
        MyCommon.Open_LogixRT()
        MyCommon.QueryStr = "dbo.pt_GroupMembership_Delete"
        MyCommon.Open_LXSsp()
        MyCommon.LXSsp.Parameters.Add("@ExtCardID", SqlDbType.NVarChar, 400).Value = MyCryptLib.SQL_StringEncrypt(ClientUserID1)
        MyCommon.LXSsp.Parameters.Add("@CardTypeID", SqlDbType.Int).Value = CardTypeID
        MyCommon.LXSsp.Parameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = CustomerGroupID
        MyCommon.LXSsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
        MyCommon.LXSsp.ExecuteNonQuery()
        outputStatus = MyCommon.LXSsp.Parameters("@Status").Value
        MyCommon.Close_LXSsp()

        If outputStatus <> -1 Then
            MyCommon.Activity_Log(4, CustomerGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.cgroup-remove", LanguageID) & " " & ClientUserID1)
            CardIDs = FindAllCustomerCards(ClientUserID1, CardTypeID, MyCommon)
            statusMessage = GetCardMessage(CardIDs, 2)
            Dim dtTemp As DataTable
            MyCommon.QueryStr = "SELECT * FROM CustomerGroups WHERE IsOptinGroup = 1 AND CustomerGroupID = @CustomerGroupID"
            MyCommon.DBParameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = CustomerGroupID
            dtTemp = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            If dtTemp.Rows.Count > 0 Then
                Dim dtTemp3 As DataTable
                MyCommon.QueryStr = "SELECT RO.IncentiveID AS OfferID, RO.Name AS OfferName FROM CPE_RewardOptions RO INNER JOIN CPE_Incentives I ON RO.IncentiveID = I.IncentiveID " & _
                                     "INNER JOIN CPE_IncentiveCustomerGroups ICG ON RO.RewardOptionID = ICG.RewardOptionID WHERE(ICG.CustomerGroupID = @CustomerGroupID)" & _
                                     "UNION SELECT O.OfferID, [Name] AS OfferName FROM Offers O INNER JOIN OfferConditions OC ON O.OfferID = OC.OfferID WHERE(OC.LinkID = @CustomerGroupID)"
                MyCommon.DBParameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = CustomerGroupID
                dtTemp3 = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                If dtTemp3.Rows.Count > 0 Then
                    Dim objOfferID As Long
                    Dim objOfferName As String
                    objOfferID = Convert.ToInt64(dtTemp3(0).Item("OfferID"))
                    objOfferName = dtTemp3(0).Item("OfferName").ToString()
                    MyCommon.Activity_Log(25, CustomerPK, AdminUserID, Copient.PhraseLib.Detokenize("term.customeroptedout", 1, objOfferID, objOfferName))
                End If
            End If
        Else
            infoMessage = Copient.PhraseLib.Detokenize("cgroup-edit.CustomerNotInGroup", LanguageID, ClientUserID1)
        End If

        MyCommon.QueryStr = "Update CustomerGroups with (RowLock) set LastUpdate=getdate() where CustomerGroupID=" & CustomerGroupID
        MyCommon.LRT_Execute()
        SendNotificationsOfItemChange(CustomerGroupID, 1)
        If infoMessage = "" Then
            Response.Status = "301 Moved Permanently"
            Response.AddHeader("Location", "cgroup-edit.aspx?CustomerGroupID=" & CustomerGroupID)
        End If
    ElseIf (Request.QueryString("close") <> "") Then
        Response.Status = "301 Moved Permanently"
    End If

    If (Request.QueryString("mode") <> "Create") Then
        MyCommon.QueryStr = "select Name,ExtGroupID,CreatedDate,LastUpdate,LastLoaded,CAMCustomerGroup, IsOptinGroup from CustomerGroups with (NoLock) where CustomerGroupID=" & CustomerGroupID & " and deleted=0"
        rst = MyCommon.LRT_Select()
        If rst.Rows.Count > 0 Then
            For Each row In rst.Rows
                If (GName = "") Then GName = MyCommon.NZ(row.Item("Name"), "")
                CreatedDate = MyCommon.NZ(row.Item("CreatedDate"), "1/1/1900")
                lastUpdate = MyCommon.NZ(row.Item("LastUpdate"), "1/1/1900")
                lastUpload = MyCommon.NZ(row.Item("LastLoaded"), "1/1/1900")
                XID = MyCommon.NZ(row.Item("ExtGroupID"), "")
                IsCAMGroup = MyCommon.NZ(row.Item("CAMCustomerGroup"), False)
                IsOptInGroup = row.Item("IsOptinGroup")
            Next

            'group counts which included nulls in EXTCardID column in CardIDs caused inaccurate count for customers in a customer group.
            If MyCommon.Fetch_SystemOption(163) Then
                MyCommon.QueryStr = "select distinct count(1) as GCount from GroupMembership GM with (NoLock) JOIN CardIDs CARDS with (NoLock) ON GM.CustomerPk = CARDS.CustomerPK where GM.CustomerGroupID =" & CustomerGroupID & " and GM.deleted = 0 "
                rst = MyCommon.LXS_Select()
                For Each row In rst.Rows
                    GroupSize = row.Item("GCount")
                Next

                MyCommon.QueryStr = "select Count(1) as CardCount from CardIDs CARDS with (NoLock) " & _
                                    "inner join  (select CustomerPK, MembershipID from GroupMembership with (NoLock) " & _
                                    "where CustomerGroupID=" & CustomerGroupID & " and Deleted=0) as GM on GM.CustomerPK = CARDS.CustomerPK"
                rst = MyCommon.LXS_Select
                If rst.Rows.Count > 0 Then
                    CardCount = MyCommon.NZ(rst.Rows(0).Item("CardCount"), GroupSize)
                End If
            End If

            MyCommon.QueryStr = "select distinct CustomerPK as CCount from GroupMembership where deleted = 0 and CustomerGroupID = " & CustomerGroupID
            rst = MyCommon.LXS_Select()
            If (rst.Rows.Count > 0) Then
                GroupSizeCustomers = rst.Rows.Count
            End If

            MyCommon.QueryStr = "select top 100 GM.CustomerPK, CARDS.ExtCardIDOriginal AS ExtCardID, CARDS.CardTypeID, GM.MembershipID, CT.Description as CardDesc, CT.PhraseID " & _
                                    "from CardIDs CARDS with (NoLock) " & _
                                    "inner join CardTypes as CT with (NoLock) on CT.CardTypeID = CARDS.CardTypeID " & _
                                    "inner join  (select top 100 CustomerPK, MembershipID from GroupMembership with (NoLock) " & _
                                    "where CustomerGroupID=" & CustomerGroupID & " and Deleted=0) as GM on GM.CustomerPK = CARDS.CustomerPK where CT.CardTypeID <> 8"

            rst = MyCommon.LXS_Select()
            Dim rowCnt As Integer = 0
            For Each row In rst.Rows
                Dim e_ExtCardID As String = ""
                e_ExtCardID = MyCryptLib.SQL_StringDecrypt(rst.Rows(rowCnt)("ExtCardID"))
                If MyCommon.Fetch_SystemOption(144) Then
                    'Mask the AltID last four digits
                    'CASE WHEN CID.CardTypeID=3 THEN LEFT(CID.ExtCardID,LEN(CID.ExtCardID)-4) ELSE CID.ExtCardID END AS ExtCardID  
                    If (CStr(rst.Rows(rowCnt)("CardTypeID")) = "3") Then
                        If (e_ExtCardID.Length < 11) Then
                            e_ExtCardID = e_ExtCardID
                        Else
                            e_ExtCardID = e_ExtCardID.Remove(10) & "****"

                        End If
                    End If
                End If
                rst.Rows(rowCnt)("ExtCardID") = e_ExtCardID
                rowCnt = rowCnt+1
            Next
        ElseIf (Request.QueryString("new") <> "New") AndAlso (CustomerGroupID > 0) Then
            ' check if this is a deleted customer group
            MyCommon.QueryStr = "select Name from CustomerGroups with (NoLock) where CustomerGroupID=" & CustomerGroupID & " and deleted =1"
            rst = MyCommon.LRT_Select()
            If (rst.Rows.Count > 0) Then
                GName = MyCommon.NZ(rst.Rows(0).Item("Name"), "")
            Else
                GName = ""
            End If

            Send_HeadBegin("term.customergroup", , CustomerGroupID)
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
                Send_Tabs(Logix, 3)
                Send_Subtabs(Logix, 31, 3, , CustomerGroupID)
            End If
            Send("<div id=""intro"">")
            Send("    <h1 id=""title"">" & Copient.PhraseLib.Lookup("term.customergroup", LanguageID) & " #" & Request.QueryString("CustomerGroupID") & " - " & GName & "</h1>")
            Send("</div>")
            Send("<div id=""main"">")
            Send("    <div id=""infobar"" class=""red-background"">")
            Send("        " & Copient.PhraseLib.Lookup("error.deleted", LanguageID))
            Send("    </div>")
            Send("</div>")
            Send_BodyEnd()
            GoTo done
        End If
    End If

    Send_HeadBegin("term.customergroup", , CustomerGroupID)
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
        Send_Tabs(Logix, 3)
        Send_Subtabs(Logix, 31, 3, , CustomerGroupID)
    End If

    If (Logix.UserRoles.AccessCustomerGroups = False) Then
        Send_Denied(1, "perm.cgroup-access")
        Send_BodyEnd()
        GoTo done
    End If

%>

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
          var bConfirm = confirm('<% Sendb(Copient.PhraseLib.Lookup("term.warning1", LanguageID))%>');
          return bConfirm;
        }
      }
    };
  }
  function PageClick(evt) {
    var target = document.all ? event.srcElement : evt.target;

    if (target.href) {
      if (IsFormChanged(document.mainform)) {
        var bConfirm = confirm('<% Sendb(Copient.PhraseLib.Lookup("term.warning1", LanguageID))%>');
        return bConfirm;
      }
    }
  }

    $(document).ready(function () {
        var isAnalyticsCG = <%=isAnalyticsCG.ToString.ToLower()%>;
        if(isAnalyticsCG) {
            //This customer group cannot be edited. So disable all editable fields and clickable buttons.
            $('#GroupName').attr('disabled', 'disabled');
            $('#add').attr('disabled', 'disabled');
            $('#mremove').attr('disabled', 'disabled');
            $('#clientusertype').attr('disabled', 'disabled');
            $('#clientuserid1').attr('disabled', 'disabled');
            $('#remove').attr('disabled', 'disabled');
            $('#editcontroltypeid').attr('disabled', 'disabled');
            $('#save').attr('disabled', 'disabled');
            $('#upload').attr('disabled', 'disabled');
        }
    });
</script>
<script type="text/javascript" language="javascript">
    function openViewBarcodes(CustomerGroupID) {
    self.name = "cgroupEditWin";
    <%
    Send("openPopup(""cgroup-edit-viewbarcodes.aspx?CustomerGroupID="" + CustomerGroupID);")
    %>
  }
    function openGenerateBarcodes(CustomerGroupID) {
    self.name = "cgroupEditWin";
    document.getElementById("generatebarcodes").disabled = 'disabled';
    <%
       Send("openPopup(""cgroup-edit-generatebarcodes.aspx?CustomerGroupID="" + CustomerGroupID + """ & IIf(GroupSizeCustomers >1, "&DisableNumberOfBarcodes=true","") & "&Parent="+ MyCommon.AppName +""");")
    %>
  }  
  function disableSaveCheck() {
    window.onunload = null;
    return true;
  }
//       function chooseFile() {
//      document.getElementById("browse").click();
//   }
//   function fileonclick()
//   {
//   var filename=document.getElementById("browse").value;
//    document.getElementById("lblfileupload").innerText = filename.replace("C:\\fakepath\\", "");
//  }
  function isValidPath() {
    var retVal = true;
    var frmElem = document.uploadform.browse
    var agt = navigator.userAgent.toLowerCase();
    var browser = '<% Sendb(Request.Browser.Browser) %>'
    
    if (browser == 'IE') {
      if (frmElem != null) {
        var filePath = frmElem.value
       
        if (filePath.length >=2) {
          if (filePath.charAt(1)!=":") {
            alert('<% Sendb(Copient.PhraseLib.Lookup("term.invalidfilepath", LanguageID))%>');
            retVal = false;
          }
        } else {
          alert('<% Sendb(Copient.PhraseLib.Lookup("term.invalidfilepath", LanguageID))%>');
          retVal = false;
        }
      }
    }
    return retVal;
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
  
  function toggleRoleIDs() {
    var ri = document.getElementById("roleid");
    var ect = document.getElementById("editcontroltypeid");
    var ect_value = 0;
    if (ect != null) {
      ect_value = ect.value;
      if (ect_value == 3) {
        ri.style.display = 'inline';
      } else {
        ri.style.display = 'none';
      }
    }
  }
  
  function toggleCardType(type) {
    var selector = document.getElementById("cardtypeid");
    if (selector != null) {
      if (type == 2) {
        selector.disabled = "disabled";
      } else {
        selector.disabled = "";
      }
    }
  }
</script>
<form action="cgroup-edit.aspx" id="mainform" name="mainform" method="get" onsubmit="return disableSaveCheck();">
<%
  If CreatedFromOffer Then
    Send("<input type=""hidden"" id=""OfferID"" name=""OfferID"" value=""" & OfferID & """ />")
    Send("<input type=""hidden"" id=""EngineID"" name=""EngineID"" value=""" & EngineID & """ />")
    Send("<input type=""hidden"" id=""slct"" name=""slct"" value=""" & Request.QueryString("slct") & """ />")
    Send("<input type=""hidden"" id=""ex"" name=""ex"" value=""" & Request.QueryString("ex") & """ />")
    Send("<input type=""hidden"" id=""condChanged"" name=""condChanged"" value=""" & Request.QueryString("condChanged") & """ />")
  End If
%>
<div id="intro">
  <%
    If CustomerGroupID = 0 Then
      GNameTitle = Copient.PhraseLib.Lookup("term.newcustomergroup", LanguageID)
    Else
      MyCommon.QueryStr = "SELECT Name, CustomerGroupID FROM CustomerGroups with (NoLock) WHERE CustomerGroupId = " & CustomerGroupID & ";"
      rst2 = MyCommon.LRT_Select
      If (rst2.Rows.Count > 0) Then
        GNameTitle = MyCommon.NZ(rst2.Rows(0).Item("Name"), "")
        If (Len(GNameTitle) > 30) Then
          GNameTitle = Left(GNameTitle, 27) & "..."
        End If
        GNameTitle = Copient.PhraseLib.Lookup("term.customergroup", LanguageID) & " #" & CustomerGroupID & ": " & GNameTitle
      End If
    End If
  %>
  <h1 id="title">
    <%If (IsCAMGroup) Then
        Sendb(Copient.PhraseLib.Lookup("term.cam", LanguageID) & " " & GNameTitle)
      Else
        Sendb(GNameTitle)
      End If
    %>
  </h1>
  <div id="controls">
    <%
      If (CustomerGroupID = 0) Then
        If (Logix.UserRoles.CreateCustomerGroups) Then
          Send_Save()
        End If
      Else
        ShowActionButton = (Logix.UserRoles.AccessCustomerGroups) OrElse (Logix.UserRoles.CreateCustomerGroups) OrElse (Logix.UserRoles.EditCustomerGroups) OrElse (Logix.UserRoles.DeleteCustomerGroups)
        If (ShowActionButton) Then
          Send("<input type=""button"" class=""regular"" id=""actions"" name=""actions"" value=""" & Copient.PhraseLib.Lookup("term.actions", LanguageID) & " ▼"" onclick=""toggleDropdown();"" />")
          Send("<div class=""actionsmenu"" id=""actionsmenu"">")
          If (Logix.UserRoles.EditCustomerGroups) Then
            Send_Save()
          End If
          If (Logix.UserRoles.DeleteCustomerGroups) Then
            Dim bEnableBuckOffers As Boolean
            bEnableBuckOffers = CMInstalled AndAlso (MyCommon.Fetch_CM_SystemOption(137) = "1")
            If bEnableBuckOffers Then
              Dim rst3 As DataTable
              MyCommon.QueryStr = "select isnull(ChildOfferID,0) as ChildOfferID from CM_BuckOffers with (NoLock) where ChildOfferID<>0 and CustomerGroupID=@CustomerGroupID;"
              MyCommon.DBParameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = CustomerGroupID
              rst3 = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
              If rst3.Rows.Count = 0 Then
                Send_Delete()
              End If
            Else
              Send_Delete()
            End If
          End If
          If (Logix.UserRoles.EditCustomerGroups AndAlso CustomerGroupID > 0 AndAlso Logix.UserRoles.CreateCustomer) Then
            Send_Upload()
          End If
          If (Logix.UserRoles.AccessCustomerGroups And Not CreatedFromOffer) Then
            Send_Download()
          End If
          If (Logix.UserRoles.CreateCustomerGroups And Not CreatedFromOffer) Then
            Send_New()
          End If
          If (Logix.UserRoles.EditCustomerGroups And Not CreatedFromOffer) Then
			If (DisableRedeploy = False) Then
            	Send_ReDeploy()
			End If
          End If
          If CreatedFromOffer Then
            Send_Close()
          End If
          If Request.Browser.Type = "IE6" Then
            Send("<iframe src=""javascript:'';"" id=""actionsiframe"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no"" style=""height:145px;""></iframe>")
          End If
          Send("</div>")
        End If
        If MyCommon.Fetch_SystemOption(75) And Not CreatedFromOffer Then
          If (Logix.UserRoles.AccessNotes) Then
            Send_NotesButton(5, CustomerGroupID, AdminUserID)
          End If
        End If
      End If
      Send("<input type=""hidden"" id=""CustomerGroupID"" name=""CustomerGroupID"" value=""" & CustomerGroupID & """ />")
      Send("<input type=""hidden"" id=""IsOptinGroup"" name=""IsOptinGroup"" value=""" & IsOptInGroup & """ />")
    %>
  </div>
</div>
<%
  If Request.Browser.Type = "IE6" Then
    IE6ScrollFix = " onscroll=""javascript:document.getElementById('uploader').style.display='none';document.getElementById('actionsmenu').style.visibility='hidden';"""
  End If
%>
<div id="main" <% Sendb(IE6ScrollFix) %>>
  <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
  <% If (infoMessage2 <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage2 & "</div>")%>
  <% If (infoMessage3 <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage3 & "</div>")%>
  <% If (statusMessage <> "") Then Send("<div id=""statusbar"" class=""green-background"">" & statusMessage & "</div>")%>
  <div id="column1">
    <div class="box" id="identity">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.identification", LanguageID))%>
        </span>
      </h2>
      <label for="GroupName">
        <%Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>:</label><br />
      <%
        If (GName Is Nothing) Then GName = ""
        Sendb("<input type=""text"" class=""" & IIf(CreatedFromOffer, "long", "longest") & """ id=""GroupName"" name=""GroupName"" maxlength=""100"" value=""" & GName.Replace("""", "&quot;") & """ " & IIf(IsOptInGroup, "readonly", "") & "/><br />")
      %>
      <br class="half" />
      <%
        If XID <> "" AndAlso XID <> "0" Then
          Sendb(Copient.PhraseLib.Lookup("term.xid", LanguageID) & ": " & XID & "<br />")
        End If
        If CreatedDate = Nothing Then
        Else
          Sendb(Copient.PhraseLib.Lookup("term.created", LanguageID) & " ")
          longDate = CreatedDate
          longDateString = Logix.ToLongDateTimeString(longDate, MyCommon)
          Send(longDateString)
          Send("<br />")
        End If
        If lastUpdate = Nothing Then
        Else
          Sendb(Copient.PhraseLib.Lookup("term.edited", LanguageID) & " ")
          longDate = lastUpdate
          longDateString = Logix.ToLongDateTimeString(longDate, MyCommon)
          Send(longDateString)
          Send("<br />")
        End If
        If (lastUpload = Nothing) OrElse (lastUpload = "1/1/1900") Then
          If (CustomerGroupID <> 0) Then
            Sendb(Copient.PhraseLib.Lookup("term.neveruploaded", LanguageID))
            Send("<br />")
          End If
        Else
          Sendb(Copient.PhraseLib.Lookup("term.lastupload", LanguageID) & " ")
          longDate = lastUpload
          longDateString = Logix.ToLongDateTimeString(longDate, MyCommon)
          Send(longDateString)
          Send("<br />")
        End If
        If CustomerGroupID <> 0 Then
          Send("<br class=""half"" />")
          Sendb(Copient.PhraseLib.Lookup("term.contains", LanguageID) & " ")
          If MyCommon.Fetch_SystemOption(163) Then
            If (GroupSize = 1) Then
              Response.Write(GroupSize & " ")
              Sendb(StrConv(Copient.PhraseLib.Lookup("term.cards", LanguageID), VbStrConv.Lowercase))
            Else
              Response.Write(GroupSize & " ")
              Sendb(StrConv(Copient.PhraseLib.Lookup("term.cards", LanguageID), VbStrConv.Lowercase))
            End If
            Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.for", LanguageID), VbStrConv.Lowercase))
          End If

          If (GroupSizeCustomers = 1) Then
            Response.Write(" " & GroupSizeCustomers & " ")
            Sendb(StrConv(Copient.PhraseLib.Lookup("term.customer", LanguageID), VbStrConv.Lowercase))
            Sendb(" " & Copient.PhraseLib.Lookup("term.id", LanguageID))
          Else
            Response.Write(" " & GroupSizeCustomers & " ")
            Sendb(StrConv(Copient.PhraseLib.Lookup("term.customer", LanguageID), VbStrConv.Lowercase))
            Sendb(" " & Copient.PhraseLib.Lookup("term.ids", LanguageID))
          End If
		  
          ' if there are multiple cards for at least one customer in the the group then display the card count
          If MyCommon.Fetch_SystemOption(163) Then
            If GroupSize <> CardCount Then
              Sendb(" (" & CardCount & " ")
              If (CardCount = 1) Then
                Sendb(Copient.PhraseLib.Lookup("term.card", LanguageID).ToLower)
              Else
                Sendb(Copient.PhraseLib.Lookup("term.cards", LanguageID).ToLower)
              End If
              Sendb(")")
            End If
          End If
        End If
        MyCommon.QueryStr = "select count(1) as GMICount from GMInsertQueue with (NoLock) where CustomerGroupID=" & CustomerGroupID
        rst2 = MyCommon.LXS_Select
        For Each row In rst2.Rows
          GMICount = row.Item("GMICount")
        Next
        If (GMICount > 0) Then
          Send("<br />")
          Send("<span class=""red"">" & GMICount & " " & Copient.PhraseLib.Lookup("cgroup-edit.awaiting", LanguageID) & "</span>")
          Send("<small><a href=""cgroup-edit.aspx?CustomerGroupID=" & CustomerGroupID & "&OfferID=" & OfferID & "&EngineID=" & EngineID & """>" & Copient.PhraseLib.Lookup("term.refresh", LanguageID) & "</a></small>")
        End If
      %>
      <hr class="hidden" />
      <%-- 'Checkbox to create CAM customer group --%>
      <% If (CAMInstalled And CustomerGroupID = 0) Then%>
      <input type="checkbox" id="CAMGroup" name="CAMGroup" value="1" />
      <label id="lbCAMGroup" for="CAMGroup">
        <% Send(Copient.PhraseLib.Lookup("term.camcustomergroup", LanguageID))%></label>
      <br />
      <%End If%>
    </div>
    <% If Not CreatedFromOffer Then%>
    <div class="box" id="eligibleoffers" <% if(customergroupid=0)then sendb(" style=""visibility: hidden;""") %>>
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.associatedeligibleoffers", LanguageID))%>
        </span>
      </h2>
      <div class="boxscrollhalf">
        <% 
          If (CustomerGroupID <> 0) Then
            Dim lstOffers As List(Of CMS.AMS.Models.Offer)
            lstOffers = m_Offer.GetEligibleOffersByCustomerGroupID(CustomerGroupID)
            
            If(bEnableRestrictedAccessToUEOfferBuilder) Then
                lstOffers = GetRoleBasedUEOffers(lstOffers,MyCommon,Logix)
            End If
            If lstOffers.Count > 0 Then
              For Each offer As CMS.AMS.Models.Offer In lstOffers
                If (Logix.IsAccessibleOffer(AdminUserID, offer.OfferID)) Then
                  Sendb(" <a href=""offer-redirect.aspx?OfferID=" & offer.OfferID & """>" & offer.OfferName & "</a>")
                Else
                  Sendb(offer.OfferName)
                End If
                If (MyCommon.NZ(offer.EndDate, Now().AddDays(-1D)) < Today) Then
                  Sendb(" (" & Copient.PhraseLib.Lookup("term.expired", LanguageID) & ")")
                End If
                Send("<br />")
              Next
            Else
              Send("     " & Copient.PhraseLib.Lookup("term.none", LanguageID))
            End If
          End If
        %>
      </div>
      <hr class="hidden" />
    </div>
    <div class="box" id="offers" <% if(customergroupid=0)then sendb(" style=""visibility: hidden;""") %>>
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.associatedoffers", LanguageID))%>
        </span>
      </h2>
      <div class="boxscroll">
        <% 
            If(bEnableRestrictedAccessToUEOfferBuilder) Then
                conditionalQuery =GetRestrictedAccessToUEBuilderQuery(MyCommon,Logix,"")
            End If
            
          If (CustomerGroupID <> 0) Then
            'MyCommon.QueryStr = "select O.OfferID,OC.ConditionID,OC.linkid,OC.excludedid,O.Name from offerconditions as OC " & _
            ' "left join Offers as O on O.Offerid=OC.offerID " & _
            ' "where (linkid=" & CustomerGroupID & " or excludedid=" & CustomerGroupID & ") and ConditionTypeID=1 and OC.Deleted=0 and O.ProdEndDate >= getdate() order by O.Name"
            MyCommon.QueryStr = "select distinct 1 as EngineID, O.OfferID,O.Name,O.ProdEndDate,NULL as BuyerID from offerconditions as OC with (NoLock) " & _
                                "  left join Offers as O with (NoLock) on O.Offerid=OC.offerID " & _
                                "  where (OC.linkid=" & CustomerGroupID & " or OC.excludedid=" & CustomerGroupID & ") and OC.ConditionTypeID=1 and O.deleted=0 and O.IsTemplate=0 and OC.Deleted=0 " & _
                                " UNION " & _
                                " select distinct 1 as EngineID, O.OfferID,O.Name,O.ProdEndDate,NULL as BuyerID from offerrewards as OFR with (NoLock) " & _
                                "  left join Offers as O with (NoLock) on O.Offerid=OFR.offerID " & _
                                "  left join RewardCustomerGroupTiers as RCGT with (NoLock) on RCGT.RewardID=OFR.RewardID " & _
                                "  where (RCGT.CustomerGroupID=" & CustomerGroupID & ") and (OFR.RewardTypeID=5 or OFR.RewardTypeID=6) and O.deleted=0 and O.IsTemplate=0 and OFR.Deleted=0 " & _
                                " UNION " & _
                                " SELECT DISTINCT 2 as EngineID, I.IncentiveID as OfferID, I.IncentiveName as Name, I.EndDate as ProdEndDate,buy.ExternalBuyerId as BuyerID from CPE_IncentiveCustomerGroups ICG with (NoLock) " & _
                                "  INNER JOIN CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID=ICG.RewardOptionID " & _
                                "  INNER JOIN CPE_Incentives I with (NoLock) on I.IncentiveID = RO.IncentiveID " & _
                                "  INNER JOIN CustomerGroups CG with (NoLock) on ICG.CustomerGroupID = CG.CustomerGroupID " & _
                                "left outer join Buyers as buy with (nolock) on buy.BuyerId= I.BuyerId " & _
                                "  WHERE ICG.CustomerGroupID=" & CustomerGroupID & " and ICG.Deleted = 0 and RO.Deleted=0 and I.Deleted=0 and I.IsTemplate=0 and CG.Deleted=0 " 
            If(bEnableRestrictedAccessToUEOfferBuilder AndAlso Not String.IsNullOrEmpty(conditionalQuery)) Then MyCommon.QueryStr &=conditionalQuery & " "                
            MyCommon.QueryStr &=  " UNION " & _
                                " SELECT DISTINCT 2 as EngineID, I.IncentiveID as OfferID, I.IncentiveName as Name, I.EndDate as ProdEndDate,buy.ExternalBuyerId as BuyerID from CPE_Deliverables D with (NoLock) " & _
                                "  INNER JOIN CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID=D.RewardOptionID " & _
                                "  INNER JOIN CPE_Incentives I with (NoLock) on I.IncentiveID = RO.IncentiveID " & _
                                "left outer join Buyers as buy with (nolock) on buy.BuyerId= I.BuyerId " & _
                                "  WHERE D.Deleted=0 and RO.Deleted=0 and I.Deleted=0 and I.IsTemplate=0 and D.DeliverableTypeId IN (5,6) and OutputID=" & CustomerGroupID  
            If(bEnableRestrictedAccessToUEOfferBuilder AndAlso Not String.IsNullOrEmpty(conditionalQuery)) Then MyCommon.QueryStr &=conditionalQuery & " " 
            MyCommon.QueryStr &= " order by Name;"
            rst2 = MyCommon.LRT_Select
            rowCount = rst2.Rows.Count
                Dim Name As String =""
            If rowCount > 0 Then
              For Each row In rst2.Rows
                If (MyCommon.Fetch_UE_SystemOption(168) = "1" AndAlso MyCommon.NZ(row.Item("BuyerID"), "") <> "") Then
                    Name = "Buyer " + row.Item("BuyerID").ToString() + " - " + MyCommon.NZ(row.Item("Name"), "").ToString()
                Else
                    Name = MyCommon.NZ(row.Item("Name"), Copient.PhraseLib.Lookup("term.unknown", LanguageID))
                End If
                If (Logix.IsAccessibleOffer(AdminUserID, MyCommon.NZ(row.Item("OfferID"), 0))) Then
                  Sendb(" <a href=""offer-redirect.aspx?OfferID=" & MyCommon.NZ(row.Item("OfferID"), 0) & """>" & MyCommon.NZ(Name, "") & "</a>")
                Else
                  Sendb(MyCommon.NZ(Name, ""))
                End If
                If (MyCommon.NZ(row.Item("ProdEndDate"), Now().AddDays(-1D)) < Today) Then
                  Sendb(" (" & Copient.PhraseLib.Lookup("term.expired", LanguageID) & ")")
                End If
                Send("<br />")
              Next
            Else
              Send("     " & Copient.PhraseLib.Lookup("term.none", LanguageID))
            End If
          End If
        %>
      </div>
      <hr class="hidden" />
    </div>
    <% End If%>
    <% If MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER) Then%>
    <div class="box" id="preferences" style="height: 100px; <% if(customergroupid=0)then sendb("visibility: hidden;") %>">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.associatedpreferences", LanguageID))%>
        </span>
      </h2>
      <div class="boxscroll" style="height: 65px;">
        <% 
          If (CustomerGroupID <> 0) Then
            MyCommon.QueryStr = "select PCG.PreferenceID, PREF.Name, PREF.NamePhraseID, PREF.UserCreated, " & _
                                "  case when UPT.Phrase is null then PREF.Name else Convert(nvarchar(200), UPT.Phrase) end as PhrasedName " & _
                                "from PreferenceCustomerGroups as PCG with (NoLock) " & _
                                "inner join Preferences as PREF with (NoLock) on PREF.PreferenceID = PCG.PreferenceID " & _
                                "left join UserPhraseText as UPT with (NoLock) on UPT.PhraseID = PREF.NamePhraseID " & _
                                "where PCG.CustomerGroupID=" & CustomerGroupID & " and PREF.Deleted = 0;"
            rst2 = MyCommon.PMRT_Select
            rowCount = rst2.Rows.Count
            If rowCount > 0 Then
              For Each row In rst2.Rows
                PrefPageName = IIf(MyCommon.NZ(row.Item("UserCreated"), False), "prefscustom-edit.aspx", "prefsstd-edit.aspx")
          
                RootURI = IntegrationVals.HTTP_RootURI
                If RootURI IsNot Nothing AndAlso RootURI.Length > 0 AndAlso Right(RootURI, 1) <> "/" Then
                  RootURI &= "/"
                End If

                Sendb("  <a href=""authtransfer.aspx?SendToURI=" & RootURI & "UI/" & PrefPageName & "?prefid=" & MyCommon.NZ(row.Item("PreferenceID"), 0) & """>")
                Send(MyCommon.NZ(row.Item("PhrasedName"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)) & "</a>")

                Send("<br />")
              Next
            Else
              Send("     " & Copient.PhraseLib.Lookup("term.none", LanguageID))
            End If
          End If
        %>
      </div>
      <hr class="hidden" />
    </div>
    <% End If%>
    <% If (MyCommon.IsEngineInstalled(2) And Not CreatedFromOffer And Not IsCAMGroup) Then%>
    <div class="box" id="validationCPE" <% if(customergroupid=0)then send(" style=""visibility: hidden;""") %>>
      <h2>
        <span>
          <%
            Dim dtEngine As DataTable
            Dim sEngine As String = ""
            MyCommon.QueryStr = "select PhraseID from PromoEngines with (NoLock) where EngineId=2"
            dtEngine = MyCommon.LRT_Select()
            If dtEngine.Rows.Count > 0 Then
              sEngine = " (" & Copient.PhraseLib.Lookup(dtEngine.Rows(0).Item(0), LanguageID) & ")"
            End If
            Sendb(Copient.PhraseLib.Lookup("term.validationreport", LanguageID) & sEngine)
          %>
        </span>
      </h2>
      <%
        Dim dtValid As DataTable
        Dim rowOK(), rowWatches(), rowWarnings() As DataRow
        Dim objTemp As Object
        Dim GraceHours As Integer
        Dim GraceCount As Double
          
        objTemp = MyCommon.Fetch_CPE_SystemOption(41)
        If Not (Integer.TryParse(objTemp.ToString, GraceHours)) Then
          GraceHours = 4
        End If
          
        objTemp = MyCommon.Fetch_CPE_SystemOption(42)
        If Not (Double.TryParse(objTemp.ToString, GraceCount)) Then
          GraceCount = 0.1D
        End If
          
        MyCommon.QueryStr = "dbo.pa_ValidationReport_CustGroup"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@CustomerGroupID", SqlDbType.Int).Value = CustomerGroupID
        MyCommon.LRTsp.Parameters.Add("@GraceHours", SqlDbType.Int).Value = GraceHours
        MyCommon.LRTsp.Parameters.Add("@GraceCount", SqlDbType.Decimal, 2).Value = GraceCount
          
        dtValid = MyCommon.LRTsp_select()
          
        rowOK = dtValid.Select("Status=0", "LocationName")
        rowWatches = dtValid.Select("Status=1", "LocationName")
        rowWarnings = dtValid.Select("Status=2", "LocationName")
          
        Send("<a id=""validLinkCPE"" href=""javascript:openPopup('validation-report.aspx?type=cg&amp;id=" & CustomerGroupID & "&amp;level=0');"">")
        Send(Copient.PhraseLib.Lookup("cgroup-edit.validlocations", LanguageID) & " (" & rowOK.Length & ")</a><br />")
        Send("<a id=""watchLinkCPE"" href=""javascript:openPopup('validation-report.aspx?type=cg&amp;id=" & CustomerGroupID & "&amp;level=1');"">")
        Send(Copient.PhraseLib.Lookup("cgroup-edit.watchlocations", LanguageID) & " (" & rowWatches.Length & ")</a><br />")
        Send("<a id=""warningLinkCPE"" href=""javascript:openPopup('validation-report.aspx?type=cg&amp;id=" & CustomerGroupID & "&amp;level=2');"">")
        Send(Copient.PhraseLib.Lookup("cgroup-edit.warninglocations", LanguageID) & " (" & rowWarnings.Length & ")</a><br />")
      %>
      <hr class="hidden" />
    </div>
    <% End If%>
    <% If (MyCommon.IsEngineInstalled(1) And Not CreatedFromOffer) Then%>
    <div class="box" id="validationCatalina" <% if(customergroupid=0)then send(" style=""visibility: hidden;""") %>>
      <h2>
        <span>
          <%
            Dim dtEngine As DataTable
            Dim sEngine As String = ""
            MyCommon.QueryStr = "select PhraseID from PromoEngines with (NoLock) where EngineId=1"
            dtEngine = MyCommon.LRT_Select()
            If dtEngine.Rows.Count > 0 Then
              sEngine = " (" & Copient.PhraseLib.Lookup(dtEngine.Rows(0).Item(0), LanguageID) & ")"
            End If
            Sendb(Copient.PhraseLib.Lookup("term.validationreport", LanguageID) & sEngine)
          %>
        </span>
      </h2>
      <%
        Dim dtValid As DataTable
        Dim rowOK(), rowWaiting(), rowWatches(), rowWarnings() As DataRow
        Dim objTemp As Object
        Dim GraceHours As Integer
        Dim GraceHoursWarn As Integer
        Dim iGroupLocations As Integer
          
        objTemp = MyCommon.Fetch_CM_SystemOption(10)
        If Not (Integer.TryParse(objTemp.ToString, GraceHours)) Then
          GraceHours = 4
        End If
          
        objTemp = MyCommon.Fetch_CM_SystemOption(11)
        If Not (Integer.TryParse(objTemp.ToString, GraceHoursWarn)) Then
          GraceHoursWarn = 24
        End If
          
        MyCommon.QueryStr = "dbo.pa_CM_ValidationReport_CustGroup"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@CustomerGroupID", SqlDbType.Int).Value = CustomerGroupID
        MyCommon.LRTsp.Parameters.Add("@GraceHours", SqlDbType.Int).Value = GraceHours
        MyCommon.LRTsp.Parameters.Add("@GraceHoursWarn", SqlDbType.Int).Value = GraceHoursWarn
          
        dtValid = MyCommon.LRTsp_select()
        iGroupLocations = dtValid.Rows.Count
          
        rowOK = dtValid.Select("Status=0", "LocationName")
        rowWaiting = dtValid.Select("Status=1", "LocationName")
        rowWatches = dtValid.Select("Status=2", "LocationName")
        rowWarnings = dtValid.Select("Status=3", "LocationName")
          
        Send("<a id=""validLinkCatalina"" href=""javascript:openPopup('CM-validation-report.aspx?type=cg&id=" & CustomerGroupID & "&level=0&gh=" & GraceHours & "&ghw=" & GraceHoursWarn & "');"">")
        Send(Copient.PhraseLib.Lookup("cgroup-edit.validlocations", LanguageID) & " (" & rowOK.Length & " of " & iGroupLocations & ")</a><br />")
        Send("<a id=""waitingLinkCatalina"" href=""javascript:openPopup('CM-validation-report.aspx?type=cg&id=" & CustomerGroupID & "&level=1&gh=" & GraceHours & "&ghw=" & GraceHoursWarn & "');"">")
        Send(Copient.PhraseLib.Lookup("cgroup-edit.waitlocations", LanguageID) & " (" & rowWaiting.Length & " of " & iGroupLocations & ")</a><br />")
        Send("<a id=""watchLinkCatalina"" href=""javascript:openPopup('CM-validation-report.aspx?type=cg&id=" & CustomerGroupID & "&level=2&gh=" & GraceHours & "&ghw=" & GraceHoursWarn & "');"">")
        Send(Copient.PhraseLib.Lookup("cgroup-edit.watchlocations", LanguageID) & " (" & rowWatches.Length & " of " & iGroupLocations & ")</a><br />")
        Send("<a id=""warningLinkCatalina"" href=""javascript:openPopup('CM-validation-report.aspx?type=cg&id=" & CustomerGroupID & "&level=3&gh=" & GraceHours & "&ghw=" & GraceHoursWarn & "');"">")
        Send(Copient.PhraseLib.Lookup("cgroup-edit.warninglocations", LanguageID) & " (" & rowWarnings.Length & " of " & iGroupLocations & ")</a><br />")
      %>
      <hr class="hidden" />
    </div>
    <% End If%>
    <% If (MyCommon.Fetch_SystemOption(112) AndAlso CustomerGroupID > 0) Then%>
    <div class="box" id="CustomerBarcaodes">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("cgroup-edit.CustomerBarcodes", LanguageID))%>
        </span>
      </h2>
      <br />
      <%
        If (Logix.UserRoles.EditCustomerGroups) Then
          Send("<input type=""button"" class=""large"" id=""generatebarcodes"" name=""generatebarcodes"" value=""" & Copient.PhraseLib.Lookup("cgroup-edit.GenerateBarcodes", LanguageID) & """ onClick=""openGenerateBarcodes(" & CustomerGroupID & ");""" & IIf(disableButton, " disabled=""disabled""", "") & " /> ")
          Sendb("<input type=""hidden"" class=""large"" id=""viewbarcodes"" name=""viewbarcodes"" value=""" & Copient.PhraseLib.Lookup("cgroup-edit.ViewBarcodes", LanguageID) & """ onclick=""openViewBarcodes(" & CustomerGroupID & ");"" /><br />")
        End If
      %>
      <br />
      <hr class="hidden" />
    </div>
    <%End If%>
  </div>
  <div id="gutter">
  </div>
  <div id="column2">
    <div class="box" id="addcustomers" <% if(customergroupid=0)then send(" style=""visibility: hidden;""") %>>
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("cgroup-edit.addremove", LanguageID))%>
        </span>
      </h2>
      <input type="text" class="medium" id="clientuserid1" name="clientuserid1" value="" style="width:50%;"
        maxlength="<%Sendb(IIf(IDLength > 0, IDLength, 26)) %>" />
      <%
          If (IsCAMGroup = False) Then
              MyCommon.QueryStr = "select CardTypeID, Description, PhraseID from CardTypes with (NoLock) "
              If HouseholdingEnabled Then
                  MyCommon.QueryStr &= "where CardTypeID not in (2,"
              Else
                  MyCommon.QueryStr &= "where CardTypeID not in (1, 2,"
              End If
              MyCommon.QueryStr &= " 8);" 'Consumer Account Number shouldn't be displayed on UI
              rst2 = MyCommon.LXS_Select
              Send("<select id=""clientusertype"" name=""clientusertype"" style=""width:45%;"">")
              For Each row2 In rst2.Rows
                  Send("<option value=""" & MyCommon.NZ(row2.Item("CardTypeID"), 0) & """" & IIf(DefaultCustTypeID = MyCommon.NZ(row2.Item("CardTypeID"), 0), "selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup(CInt(MyCommon.NZ(row2.Item("PhraseID"), 0)), LanguageID, MyCommon.NZ(row2.Item("Description"), "")) & "</option>")
              Next
              Send("</select>")
              Send("<br />")
          End If
          If (Logix.UserRoles.EditCustomerGroups) Then
              Sendb("<input type=""submit"" class=""regular"" id=""add"" name=""add"" value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """ />")
              Sendb("<input type=""submit"" class=""regular"" id=""mremove"" name=""mremove"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("confirm.remove", LanguageID) & "')){}else{return false}"" style=""width:120px;"" value=""" & Copient.PhraseLib.Lookup("term.removemanually", LanguageID) & """ /><br />")
          End If
      %>
      <br />
      <%Sendb(Copient.PhraseLib.Lookup("cgroup-edit.listnote", LanguageID))%>
      <br />
      <select class="longer" id="membershipid" name="membershipid" size="10" multiple="multiple">
        <%
          If (GroupSizeCustomers > 0) Then
            Dim CardTypePhraseID As Int32 = 0
            For Each row In rst.Rows
              CardTypePhraseID = MyCommon.Extract_Val(MyCommon.NZ(row.Item("PhraseID"), ""))
              OptionText = MyCommon.NZ(row.Item("ExtCardID"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)) & " (" & IIf(CardTypePhraseID > 0, Copient.PhraseLib.Lookup(CardTypePhraseID, LanguageID), MyCommon.NZ(row.Item("CardDesc"), "")) & ") "
              OptionText = OptionText.Replace("""", "&quot;")
              Send("<option value=""" & MyCommon.NZ(row.Item("MembershipID"), 0) & """ alt=""" & OptionText & """ title=""" & OptionText & """>" & OptionText & "</option>")
            Next
          End If
        %>
      </select>
      <br />
      <%
        If (Logix.UserRoles.EditCustomerGroups) Then
          Send("<input type=""submit"" class=""large"" id=""remove"" name=""remove"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("confirm.remove", LanguageID) & "')){}else{return false}"" value=""" & Copient.PhraseLib.Lookup("term.removefromlist", LanguageID) & """ /><br />")
        End If
      %>
      <hr class="hidden" />
    </div>
    <%
      If (CustomerGroupID > 0 And Not CreatedFromOffer) Then
        MyCommon.QueryStr = "select CMOADeployStatus,CMOADeployRpt,CMOARptDate,CMOADeploySuccessDate from CustomerGroups with (nolock) where CustomerGroupID=" & CustomerGroupID & " and Deleted=0"
        rst = MyCommon.LRT_Select()
        If rst.Rows.Count > 0 Then
          For Each row In rst.Rows
            If (MyCommon.NZ(row.Item("CMOARptDate").ToString, "") = "" AndAlso MyCommon.NZ(row.Item("CMOADeployRpt").ToString, "") = "") Then
              GoTo nodeployment
            End If
          Next
          Sendb("<div class=""box"" id=""deployment"">")
          Send("  <h2>")
          Send("    <span>" & Copient.PhraseLib.Lookup("term.deployment", LanguageID) & "</span>")
          Send("  </h2>")
          Send("<h3>" & Copient.PhraseLib.Lookup("term.lastattempted", LanguageID) & ":</h3>")
          deployDate = MyCommon.NZ(row.Item("CMOARptDate"), "")
          If deployDate = "" Then
            Send(Copient.PhraseLib.Lookup("term.never", LanguageID) & "<br />")
          Else
            Send(Logix.ToLongDateTimeString(CDate(deployDate), MyCommon) & "<br />")
          End If
          Send("<br class=""half"" />")
          Send("<h3>" & Copient.PhraseLib.Lookup("term.lastsuccessful", LanguageID) & ":</h3>")
          deployDate = MyCommon.NZ(row.Item("CMOADeploySuccessDate"), "")
          If deployDate = "" Then
            Send(Copient.PhraseLib.Lookup("term.never", LanguageID) & "<br />")
          Else
            Send(Logix.ToLongDateTimeString(CDate(deployDate), MyCommon) & "<br />")
          End If
          Send("<br class=""half"" />")
          Send("<h3>" & Copient.PhraseLib.Lookup("term.status", LanguageID) & ":</h3>")
          Send(MyCommon.NZ(row.Item("CMOADeployRpt"), "") & "<br />")
          Send("<hr class=""hidden"" />")
          Send("</div>")
nodeployment:
        End If
      End If
    %>
    <div class="box" id="roles" <% if(customergroupid=0)then send(" style=""visibility: hidden;""") %>>
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.editcontrol", LanguageID))%>
        </span>
      </h2>
      <%
        Send("<p>" & Copient.PhraseLib.Lookup("cgroup-edit.roles", LanguageID) & "</p>")
        Send("<select id=""editcontroltypeid"" name=""editcontroltypeid"" onchange=""toggleRoleIDs();"">")
        MyCommon.QueryStr = "select EditControlTypeID, Name, PhraseID from EditControlTypes with (NoLock) order by EditControlTypeID;"
        dt = MyCommon.LRT_Select
        If dt.Rows.Count > 0 Then
          For Each row In dt.Rows
            Send("  <option value=""" & MyCommon.NZ(row.Item("EditControlTypeID"), 0) & """" & IIf(MyCommon.NZ(row.Item("EditControlTypeID"), 0) = EditControlTypeID, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "</option>")
          Next
        End If
        Send("</select>")
        Send("<select id=""roleid"" name=""roleid""" & IIf(EditControlTypeID = 3, "", " style=""display:none;""") & ">")
        MyCommon.QueryStr = "select AR.RoleID, AR.RoleName, AR.DisplayOrder, AR.PhraseID, AR.ExtRoleName from AdminRoles as AR " & _
                            "where AR.RoleID in (select RoleID from RolePermissions where PermissionID in (50,49)) " & _
                            "order by RoleName;"
        dt = MyCommon.LRT_Select
        If dt.Rows.Count > 0 Then
          For Each row In dt.Rows
            Send("  <option value=""" & MyCommon.NZ(row.Item("RoleID"), 0) & """" & IIf(MyCommon.NZ(row.Item("RoleID"), 0) = RoleID, " selected=""selected""", "") & ">" & MyCommon.NZ(row.Item("RoleName"), "&nbsp;") & "</option>")
          Next
        End If
        Send("</select>")
      %>
    </div>
  </div>
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
      <form action="cgroup-edit.aspx<%Sendb(IIf(CreatedFromOffer, "?OfferID=" & OfferID & "&EngineID=" & EngineID & "&slct=" & Request.QueryString("slct") & "&ex=" & Request.QueryString("ex"), "")) %>"
      id="uploadform" name="uploadform" onsubmit="return isValidPath();" method="post"
      enctype="multipart/form-data">
      <%
        Send("<input type=""button"" class=""ex"" id=""uploaderclose"" name=""uploaderclose"" value=""X"" style=""float:right;position:relative;top:-27px;"" onclick=""javascript:document.getElementById('uploader').style.display='none';"" />")
        Send(Copient.PhraseLib.Lookup("cgroup-edit.upload", LanguageID))
        Send("<br class=""half"" />")
        Send("<h3>" & Copient.PhraseLib.Lookup("term.method", LanguageID) & "</h3>")
        Send("<input type=""radio"" name=""operation"" id=""operation1"" value=""0"" checked=""checked"" /><label for=""operation1"">" & Copient.PhraseLib.Lookup("cgroup-edit.uploadmethod1", LanguageID) & "</label><br />")
        Send("<input type=""radio"" name=""operation"" id=""operation2"" value=""1""  /><label for=""operation2"">" & Copient.PhraseLib.Lookup("cgroup-edit.uploadmethod2", LanguageID) & "</label><br />")
        Send("<input type=""radio"" name=""operation"" id=""operation3"" value=""2""  /><label for=""operation3"">" & Copient.PhraseLib.Lookup("cgroup-edit.uploadmethod3", LanguageID) & "</label><br />")
        Send("<br />")
        Send("<h3>" & Copient.PhraseLib.Lookup("term.format", LanguageID) & "</h3>")
        Send("<input type=""radio"" name=""format"" id=""format1"" value=""1"" onclick=""toggleCardType(1);"" checked=""checked"" /><label for=""format1"">" & Copient.PhraseLib.Lookup("cgroup-edit.uploadformat1", LanguageID) & "</label><br />")
          
        Send("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;of this type: <select name=""cardtypeid"" id=""cardtypeid"">")
        MyCommon.QueryStr = "select CardTypeID, Description, CustTypeID, PhraseID, ExtCardTypeID from CardTypes with (NoLock) " & _
                            "where CustTypeID in (" & IIf(IsCAMGroup, "2", "0,1") & ") " & _
                            "order by CardTypeID;"
        dst = MyCommon.LXS_Select
        If dst.Rows.Count > 0 Then
          For Each row In dst.Rows
            Send("  <option value=""" & MyCommon.NZ(row.Item("CardTypeID"), 0) & """>" & IIf(MyCommon.NZ(row.Item("PhraseID"), 0) > 0, Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID), MyCommon.NZ(row.Item("Description"), "&nbsp;")) & "</option>")
          Next
        End If
        Send("</select><br />")
          
        Send("<input type=""radio"" name=""format"" id=""format2"" value=""2"" onclick=""toggleCardType(2);"" /><label for=""format2"">" & Copient.PhraseLib.Lookup("cgroup-edit.uploadformat2", LanguageID) & "</label><br />")
        Send("<br />")
        If (Logix.UserRoles.EditCustomerGroups) Then
          Send("<input type=""hidden"" name=""CustomerGroupID"" value=""" & CustomerGroupID & """ />")
          Send("<input type=""file"" id=""browse"" name=""browse"" value=""" & Copient.PhraseLib.Lookup("term.browse", LanguageID) & """ />")
        '         Send("<div id=""divfile"" style=""height:0px;overflow:hidden"">")
        'Send("<input type=""file"" id=""browse"" name=""fileInput"" onchange=""fileonclick()"" />")
        'Send("</div>")
        '            If Request.Browser.Type = "IE9" Then
        '                Send("<div id=""divfile"" style=""display:none;"">")
        'Send("<button type=""button"" onclick=""chooseFile();"">"&Copient.PhraseLib.Lookup("term.browse", LanguageID) &"</button>")
        '                Send("</div>")
        '            Else
        '                Send("<button type=""button"" onclick=""chooseFile();"">" & Copient.PhraseLib.Lookup("term.browse", LanguageID) & "</button>")
        '            End If
        'Send("<label id=""lblfileupload"" name=""lblfileupload"">"&Copient.PhraseLib.Lookup("term.nofilesselected", LanguageID)&"</label>")
          Send("<input type=""submit"" class=""regular"" id=""uploadfile"" name=""uploadfile"" value=""" & Copient.PhraseLib.Lookup("term.upload", LanguageID) & """ />")
          Send("<br />")
        End If
      %>
      </form>
      <hr class="hidden" />
    </div>
  </div>
  <%
    If Request.Browser.Type = "IE6" Then
      Send("<iframe src=""javascript:'';"" id=""uploadiframe-cg"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no""></iframe>")
    End If
  %>
</div>
<script type="text/javascript">
  if (window.captureEvents) {
    window.captureEvents(Event.CLICK);
    window.onclick = handlePageClick;
  }
  else {
    document.onclick = handlePageClick;
  }
</script>
<script runat="server">
  Function FindAllCustomerCards(ByVal ExtCardID As String, ByVal CardTypeID As Integer, ByRef MyCommon As Copient.CommonInc) As String()
    Dim CardIDs(-1) As String
        Dim rst As DataTable
        Dim MyCryptLib As New Copient.CryptLib
    Dim i As Integer
    
    ' find all the card numbers for the customer being added to the group
    MyCommon.QueryStr = "select ExtCardID from CardIDs with (NoLock) where CustomerPK in " & _
                        "  (select CustomerPK from CardIDs where ExtCardID=@ExtCardID and CardTypeID=@CardTypeID)"
    MyCommon.DBParameters.Add("@ExtCardID", SqlDbType.NVarChar, 400).Value = MyCryptLib.SQL_StringEncrypt(ExtCardID)
    MyCommon.DBParameters.Add("@CardTypeID", SqlDbType.Int).Value = CardTypeID
    rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)
    If rst.Rows.Count > 0 Then
      ReDim CardIDs(rst.Rows.Count - 1)
      For i = 0 To rst.Rows.Count - 1
        CardIDs(i) = MyCryptLib.SQL_StringDecrypt(rst.Rows(i).Item("ExtCardID").ToString())
      Next
    End If
    Return CardIDs
  End Function
  
  Function GetCardMessage(ByVal CardIDs As String(), ByVal MessageType As Integer) As String
    Dim CardMessage As String = ""
    
    If MessageType = 1 Then
      CardMessage = Copient.PhraseLib.Lookup("cgroup-edit.added-to-group", LanguageID)
    Else
      CardMessage = Copient.PhraseLib.Lookup("cgroup-edit.removed-from-group", LanguageID)
    End If
    
    If CardIDs.Length > 0 Then
      CardMessage = Copient.PhraseLib.Detokenize("cgroup-edit.WithAssignedCards", LanguageID, CardIDs.Length)
    End If
    
    Return CardMessage
  End Function
</script>
<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (CustomerGroupID > 0 AndAlso Logix.UserRoles.AccessNotes And Not CreatedFromOffer) Then
      Send_Notes(5, CustomerGroupID, AdminUserID)
    End If
  End If
  Send_BodyEnd("mainform", "GroupName")
done:
  If MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER) Then
    MyCommon.Close_PrefManRT()
  End If
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  MyCommon = Nothing
  Logix = Nothing
%>
