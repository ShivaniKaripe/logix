<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%@ Import Namespace="System.Globalization" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="CMS.AMS" %>
<%@ Import Namespace="CMS.AMS.Contract" %>
<%
    ' *****************************************************************************
    ' * FILENAME: SV-edit.aspx 
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
    Dim row, row2 As System.Data.DataRow
    Dim MyCommon As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    Dim Localization As Copient.Localization
    Dim dstPrograms As System.Data.DataTable
    Dim dstAssociated As System.Data.DataTable = Nothing
    Dim dstSV As System.Data.DataTable
    Dim rst As DataTable
    Dim rst2 As DataTable
    Dim drsCM() As DataRow = Nothing
    Dim pgTotalSV As Long
    Dim pgDescription As String
    Dim pgPromoVarID As String
    Dim pgCreated As String
    Dim pgUpdated As String
    Dim pgValue As String
    Dim pgValuePrecision As String
    Dim pgExpire As String
    Dim pgName As String
    Dim sExtProgramID As String = ""
    Dim ProgramDescription As String
    Dim ProgramName As String
    Dim ProgramValue As String
    Dim ProgramExpire As String
    Dim ProgramID As Long
    Dim ScorecardID As Integer = 0
    Dim ScorecardDesc As String = ""
    Dim ScorecardBold As Boolean = False
    Dim AdjustmentUPC As String = ""
    Dim assocName As String
    Dim assocID As String
    Dim l_pgID As String
    Dim longDate As New DateTime
    Dim CpeEngineOnly As Boolean = False
    Dim ProgramNameTitle As String = ""
    Dim StoreInUnits As Boolean = False
    Dim svExpireType As Integer = 2
    Dim svExpirePeriodType As Integer = 0
    Dim svExpireTOD As String = "00:00"
    Dim svExpireTODHr As String = "0"
    Dim svExpireTODMin As String = "0"
    Dim svExpireDate As Date
    Dim svExpireDateStr As String = ""
    Dim svExpireHr As Integer = 0
    Dim svExpireMin As Integer = 0
    Dim UomLimit As Integer = 1
    Dim AllowReissue As Boolean = False
    Dim ExpireCentralServerTZ As Boolean = False
    Dim SVTypeID As Integer = 2
    Dim selectedStr As String = ""
    Dim ShowActionButton As Boolean = False
    Dim OfferCtr As Integer = 0
    Dim OfferUpdateNum As Integer = 0
    Dim OfferDeployNum As Integer = 0
    Dim IE6ScrollFix As String = ""
    Dim i As Integer
    Dim statusMessage As String = ""
    Dim infoMessage As String = ""
    Dim RewardsSetScorecard As Boolean = False
    Dim Handheld As Boolean = False
    Dim AutoDelete As Boolean = False
    Dim ValidSetUPC As Boolean = False
    Dim ValidSingleUPC As Boolean = False
    Dim BeginUPC As Decimal = 0
    Dim EndUPC As Decimal = 0
    Dim UPCDec As Decimal = 0
    Dim IDLength As Integer = 0
    Dim CPEInstalled As Boolean = False
    Dim POSVAdjs As Boolean = False
    Dim RedemptionRestrictionID As Integer = 0
    Dim MemberRedemptionID As Integer = 0
    Dim AllowAnyCustomer_UE As Boolean = False
    Dim AllowAnyCustomer_UE_PKID As Long = 0

    Dim RangeBegin As Decimal = 0
    Dim RangeBeginString As String = ""
    Dim RangeEnd As Decimal = 0
    Dim RangeEndString As String = ""
    Dim Range As Decimal = 0
    Dim x As Decimal = 0
    Dim counter As Integer = 1
    Dim MaxLength As Integer = 0
    Dim UEInstalled As Boolean = False
    Dim ReturnHandlingTypeID As Integer = 1
    Dim DisallowRedeemInTrans As Boolean = False
    Dim VisibleToCustomers As Boolean = False
    Dim AllowNegativeBal As Boolean = False
    Dim MultiLanguageEnabled As Boolean = False
    Dim DefaultLanguageID As Integer = 0
    Dim DefaultLanguageCode As String = ""
    Dim MLI As New Copient.Localization.MultiLanguageRec
    Dim bCmInstalled As Boolean = False
    Dim bAllowFuelPartner As Boolean = False
    Dim bFuelPartner As Boolean = False
    Dim bAutoRedeem As Boolean = False
    Dim bAllowAdjustments As Boolean = False
    Dim bEditUomLimit As Boolean = False
    Dim lstEligibleOffers As New List(Of Models.Offer)
    Dim IsAnyCustomerOffersExist_UE As Boolean = False
    Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249)="1",True,False)
    Dim conditionalQuery = String.Empty
    Dim bAllowExpirationExtension = False
    Dim bProgramActive = False
    Dim NewExtID as String

    CurrentRequest.Resolver.AppName = "SV-edit.aspx"
    Dim m_Offer As IOffer = CurrentRequest.Resolver.Resolve(Of IOffer)()

    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If

    Response.Expires = 0
    MyCommon.AppName = "SV-edit.aspx"
    MyCommon.Open_LogixRT()
    MyCommon.Open_LogixXS()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
    Localization = New Copient.Localization(MyCommon)

    CPEInstalled = MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE)
    UEInstalled = MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE)
    POSVAdjs = IIf(MyCommon.Fetch_SystemOption(94) = 1, True, False)

    MultiLanguageEnabled = IIf(MyCommon.Fetch_SystemOption(124) = "1", True, False)
    Integer.TryParse(MyCommon.Fetch_SystemOption(1), DefaultLanguageID)
    If DefaultLanguageID > 0 Then
        MyCommon.QueryStr = "select MSNetCode from Languages with (NoLock) where LanguageID=" & DefaultLanguageID & ";"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
            DefaultLanguageCode = MyCommon.NZ(rst.Rows(0).Item("MSNetCode"), "")
        End If
    End If
    MyCommon.QueryStr = "select paddingLength from producttypes where producttypeid=1"
    rst = MyCommon.LRT_Select
    If rst IsNot Nothing Then
        IDLength = Convert.ToInt32(rst.Rows(0).Item("paddingLength"))
    End If
    'Integer.TryParse(MyCommon.Fetch_SystemOption(52), IDLength)
    If (Request.QueryString("AdjustmentUPC") <> "") Then
        If (IDLength > 0) Then
            AdjustmentUPC = Left(Trim(Request.QueryString("AdjustmentUPC")), 26).PadLeft(IDLength, "0")
        Else
            AdjustmentUPC = Left(Trim(Request.QueryString("AdjustmentUPC")), 26)
        End If
    End If

    If MyCommon.Fetch_CPE_SystemOption(100) = "" OrElse MyCommon.Fetch_CPE_SystemOption(101) = "" Then
        RangeBegin = 0
        RangeEnd = 0
        Range = 0
    Else
        RangeBegin = CDec(MyCommon.Fetch_CPE_SystemOption(100))
        RangeEnd = CDec(MyCommon.Fetch_CPE_SystemOption(101))
        Range = (RangeEnd - RangeBegin) + 1
    End If
    RangeBeginString = RangeBegin.ToString.PadLeft(IDLength, "0")
    RangeEndString = RangeEnd.ToString.PadLeft(IDLength, "0")
    MaxLength = MyCommon.Extract_Val(IDLength)

    bAllowExpirationExtension = IIf(MyCommon.Fetch_SystemOption(281) = "1", True, False)

    If (Request.QueryString("infoMessage") <> "") Then
        infoMessage = Request.QueryString("infoMessage")
    End If

    If (Request.QueryString("new") <> "") Then
        Response.Redirect("sv-edit.aspx")
    End If

    If(bEnableRestrictedAccessToUEOfferBuilder) Then
        conditionalQuery=GetRestrictedAccessToUEBuilderQuery(MyCommon,Logix,"I")
    End If


    l_pgID = MyCommon.Extract_Val(Request.QueryString("ProgramGroupID"))
    If (l_pgID > 0) Then
        MyCommon.QueryStr = "SELECT DISTINCT I.IncentiveName as Name, I.IncentiveID as OfferID, 2 as EngineID,buy.ExternalBuyerId as BuyerID from CPE_IncentiveStoredValuePrograms ISVP with (NoLock)  " & _
                            "INNER JOIN CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID=ISVP.RewardOptionID  " & _
                            "INNER JOIN CPE_Incentives I with (NoLock) on I.IncentiveID = RO.IncentiveID  " & _
                            "INNER JOIN StoredValuePrograms SVP with (NoLock) on ISVP.SVProgramID = SVP.SVProgramID  " & _
                             "left outer join Buyers as buy with (nolock) on buy.BuyerId= I.BuyerId " & _
                            "WHERE ISVP.SVProgramID=" & l_pgID & " and ISVP.Deleted = 0 and RO.Deleted=0 and I.Deleted=0 and I.IsTemplate=0 and SVP.Deleted=0  "
        If(bEnableRestrictedAccessToUEOfferBuilder AndAlso Not String.IsNullOrEmpty(conditionalQuery)) Then MyCommon.QueryStr &=conditionalQuery & " "
        MyCommon.QueryStr &="UNION  " & _
                           "SELECT DISTINCT I.IncentiveName as Name, I.IncentiveID as OfferID, 2 as EngineID,buy.ExternalBuyerId as BuyerID from CPE_DeliverableStoredValue DSV with (NoLock) " & _
                           "INNER JOIN CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID=DSV.RewardOptionID  " & _
                           "INNER JOIN CPE_Incentives I with (NoLock) on I.IncentiveID = RO.IncentiveID  " & _
                            "left outer join Buyers as buy with (nolock) on buy.BuyerId= I.BuyerId " & _
                           "WHERE DSV.Deleted=0 and RO.Deleted=0 and I.Deleted=0 and I.IsTemplate=0 and SVProgramID=" & l_pgID & " "
        If(bEnableRestrictedAccessToUEOfferBuilder AndAlso Not String.IsNullOrEmpty(conditionalQuery)) Then MyCommon.QueryStr &=conditionalQuery & " "
        MyCommon.QueryStr &= "UNION  " & _
             "SELECT DISTINCT I.IncentiveName as Name, I.IncentiveID as OfferID, 2 as EngineID,buy.ExternalBuyerId as BuyerID from CPE_DeliverableMonStoredValue DMSV with (NoLock) " & _
             "INNER JOIN CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID=DMSV.RewardOptionID  " & _
             "INNER JOIN CPE_Incentives I with (NoLock) on I.IncentiveID = RO.IncentiveID  " & _
              "left outer join Buyers as buy with (nolock) on buy.BuyerId= I.BuyerId " & _
             "WHERE DMSV.Deleted=0 and RO.Deleted=0 and I.Deleted=0 and I.IsTemplate=0 and SVProgramID=" & l_pgID & " "
        If(bEnableRestrictedAccessToUEOfferBuilder AndAlso Not String.IsNullOrEmpty(conditionalQuery)) Then MyCommon.QueryStr &=conditionalQuery & " "
        MyCommon.QueryStr &="UNION " & _
                       "SELECT DISTINCT I.IncentiveName as Name, I.IncentiveID as OfferID, I.EngineID,buy.ExternalBuyerId as BuyerID from CPE_Deliverables D with (NoLock) " & _
                       "INNER JOIN CPE_RewardOptions RO on RO.RewardOptionID = D.RewardOptionID and RO.Deleted=0 " & _
                       "INNER JOIN CPE_Incentives I on I.IncentiveID = RO.IncentiveID and I.Deleted = 0 " & _
                       "INNER JOIN CPE_Discounts DISC on DISC.DiscountID = D.OutputID and DISC.Deleted=0 " & _
                        "left outer join Buyers as buy with (nolock) on buy.BuyerId= I.BuyerId " & _
                       "WHERE D.DeliverableTypeID=2 and D.Deleted=0 and DISC.SVProgramID = " & l_pgID & " "
        If(bEnableRestrictedAccessToUEOfferBuilder AndAlso Not String.IsNullOrEmpty(conditionalQuery)) Then MyCommon.QueryStr &=conditionalQuery & " "
        MyCommon.QueryStr &="UNION " & _
                      "select distinct O.Name as Name, O.OfferID, 0 as EngineID,NULL as BuyerID " & _
                      "from StoredValuePrograms SVP with (NoLock) " & _
                      "inner join OfferConditions OC with (NoLock) on OC.deleted=0 and OC.ConditionTypeID=6 and OC.LinkId=SVP.SVProgramId " & _
                      "inner join Offers O with (NoLock) on O.deleted=0 and O.OfferId=OC.OfferID " & _
                      "where SVP.deleted=0 and SVP.SVProgramId=" & l_pgID & " " & _
                      "union " & _
                      "select distinct O.Name as Name, O.OfferID, 0 as EngineID,NULL as BuyerID " & _
                      "from StoredValuePrograms SVP with (NoLock) " & _
                      "inner join OfferRewards ORW with (NoLock) on ORW.deleted = 0 and ORW.RewardTypeId in (10,14) " & _
                      "inner join CM_RewardStoredValues as RSV with (NoLock) on RSV.RewardStoredValuesID = ORW.LinkId and RSV.ProgramId=SVP.SVProgramId " & _
                      "inner join Offers O with (NoLock) on O.deleted=0 and O.OfferId=ORW.OfferID " & _
                      "where SVP.deleted=0 and SVP.SVProgramId=" & l_pgID & " " & _
                      "union " & _
                      "select distinct O.Name as Name, O.OfferID, 0 as EngineID,NULL as BuyerID " & _
                      "from StoredValuePrograms SVP with (NoLock) " & _
                      "inner join OfferRewards ORW with (NoLock) on ORW.deleted = 0 and ORW.RewardTypeId = 1 and ORW.RewardAmountTypeId =10 " & _
                      "inner join Discounts D on D.DiscountId = ORW.LinkId " & _
                      "inner join CM_RewardStoredValues as RSV with (NoLock) on RSV.RewardStoredValuesID = D.SVLinkId and RSV.ProgramId=SVP.SVProgramId " & _
                      "inner join Offers O with (NoLock) on O.deleted=0 and O.OfferId=ORW.OfferID " & _
                      "where SVP.deleted = 0 and SVP.SVProgramId=" & l_pgID & ";"
        dstAssociated = MyCommon.LRT_Select
        drsCM = dstAssociated.Select("EngineID=0")
    End If

    lstEligibleOffers = m_Offer.GetEligibleOffersBySVProgramID(l_pgID)

    bCmInstalled = MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)
    If bCmInstalled Then
        If MyCommon.Fetch_CM_SystemOption(56) = "1" Then
            bAllowFuelPartner = True
        End If
        Dim sFuelPointsPrograms As String = MyCommon.NZ(MyCommon.Fetch_CM_SystemOption(59), "")
        If sFuelPointsPrograms <> "" Then
            sFuelPointsPrograms &= ","
            bEditUomLimit = sFuelPointsPrograms.Contains(l_pgID.ToString & ",")
        End If
    End If

    ' any GET parms inbound?
    If (Request.QueryString("Delete") <> "") Then
        l_pgID = MyCommon.Extract_Val(Request.QueryString("ProgramGroupID"))
        If (l_pgID > 0 AndAlso Not dstAssociated Is Nothing AndAlso dstAssociated.Rows.Count = 0 AndAlso lstEligibleOffers.Count = 0) Then

            ' check that there are no deployed offers that use this stored value program
            MyCommon.QueryStr = "dbo.pa_AssociatedOffers_ST"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@LinkType", SqlDbType.Int).Value = 4
            MyCommon.LRTsp.Parameters.Add("@LinkID", SqlDbType.Int).Value = l_pgID
            rst = MyCommon.LRTsp_select
            MyCommon.Close_LRTsp()

            If (rst.Rows.Count > 0) Then
                infoMessage = Copient.PhraseLib.Lookup("term.inusedeployment", LanguageID) & " : ("
                For OfferCtr = 0 To rst.Rows.Count - 1
                    infoMessage &= MyCommon.NZ(rst.Rows(OfferCtr).Item("IncentiveID"), "")
                Next
                infoMessage &= ")"
            Else
                ' expunge record if there is one
                If (MyCommon.Extract_Val(Request.QueryString("ProgramGroupID")) <> "") Then
                    MyCommon.QueryStr = "pt_StoredValuePrograms_Delete"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@SVProgramID", SqlDbType.BigInt).Value = l_pgID
                    MyCommon.LRTsp.ExecuteNonQuery()
                    MyCommon.Close_LRTsp()
                End If
                'Record history
                MyCommon.Activity_Log(26, l_pgID, AdminUserID, Copient.PhraseLib.Lookup("history.sv-delete", LanguageID))
                Response.Status = "301 Moved Permanently"
                Response.AddHeader("Location", "SV-list.aspx")
                GoTo done
            End If
        Else
            infoMessage = Copient.PhraseLib.Lookup("sv-edit.inuse", LanguageID)
        End If
    ElseIf (Request.QueryString("save") <> "" AndAlso MyCommon.Extract_Val(Request.QueryString("ProgramGroupID")) = 0) Then
        ' add a record
        ProgramName = Logix.TrimAll(Request.QueryString.Item("name"))
        ProgramDescription = iif(isEmpty(Request.QueryString.Item("desc")),"",Request.QueryString.Item("desc"))
        ProgramValue = MyCommon.Extract_Decimal(Request.QueryString("value"), MyCommon.GetAdminUser.Culture)
        ProgramExpire = MyCommon.Extract_Val(Request.QueryString("expireperiod"))
        StoreInUnits = IIf(Request.QueryString("storeInUnits") = "1", True, False)
        svExpireType = MyCommon.Extract_Val(Request.QueryString("expiretype"))
        svExpirePeriodType = MyCommon.Extract_Val(Request.QueryString("expireperiodtype"))
        svExpireTOD = Request.QueryString("expiretod")
        svExpireDateStr = Request.QueryString("expiredatetime")
        VisibleToCustomers = MyCommon.Extract_Val(Request.QueryString("VisibleToCustomers"))
        If String.IsNullOrWhiteSpace(svExpireDateStr) = False Then
            Dim svExpireDateTime As DateTime
            DateTime.TryParse(svExpireDateStr, MyCommon.GetAdminUser.Culture, DateTimeStyles.None, svExpireDateTime)
            If svExpireDateTime < DateTime.Now AndAlso svExpireType = 1 Then
                infoMessage = Copient.PhraseLib.Lookup("logix-js.EnterValidExpDate", LanguageID).Replace("&#39;", "\'")
            End If
            svExpireDateStr = svExpireDateTime.ToString(CultureInfo.CurrentCulture)
        End If
        If Not ValidateMonthExpiry(ProgramExpire, svExpirePeriodType) Then
            infoMessage = Copient.PhraseLib.Lookup("error.invalidsvmonthexpiry", LanguageID)
        End If
        SVTypeID = MyCommon.Extract_Val(Request.QueryString("svtypeid"))
        If (SVTypeID = 0) Then
            SVTypeID = 2
        End If
        If (SVTypeID = 3 OrElse SVTypeID = 5 OrElse (SVTypeID = 1 And bEditUomLimit)) Then
            UomLimit = MyCommon.Extract_Val(Request.QueryString("uomlimit"))
            If (UomLimit <= 0) Then
                UomLimit = 1
            End If
        Else
            UomLimit = 1
        End If
        ExpireCentralServerTZ = (Request.QueryString("expirecentralserverTZ") = "on")
        AllowReissue = (Request.QueryString("allowreissue") = "1")
        ScorecardID = MyCommon.Extract_Val(Request.QueryString("ScorecardID"))
        ScorecardDesc = Request.QueryString("ScorecardDesc")
        If (CPEInstalled) AndAlso (POSVAdjs) AndAlso (SVTypeID <> 2 AndAlso SVTypeID <> 4) Then
            BeginUPC = CDec(MyCommon.Fetch_CPE_SystemOption(100))
            EndUPC = CDec(MyCommon.Fetch_CPE_SystemOption(101))
            If AdjustmentUPC = "" Then
                ValidSetUPC = True
            ElseIf Not AllDigits(AdjustmentUPC) Then
                ValidSetUPC = False
            Else
                If (CDec(AdjustmentUPC) < 0) Then
                    ValidSetUPC = False
                Else
                    'Is UPC allowed to be outside the set of UPCs
                    If MyCommon.Fetch_CPE_SystemOption(102) = "0" Then
                        UPCDec = CDec(AdjustmentUPC)
                        If UPCDec > (BeginUPC - 1) AndAlso UPCDec < (EndUPC + 1) Then
                            ValidSetUPC = True
                        End If
                    Else
                        ValidSetUPC = True
                    End If
                End If
            End If
            'The UPC is not allowed to be on more than one program
            If AllDigits(AdjustmentUPC) Then
                If AdjustmentUPC = "" OrElse CDec(AdjustmentUPC) = 0 Then
                    ValidSingleUPC = True
                Else
                    MyCommon.QueryStr = "select SVProgramID As ID from StoredValuePrograms with (NoLock) " & _
                                        "where AdjustmentUPC='" & AdjustmentUPC & "' " & _
                                        "union " & _
                                        "select ProgramID as ID from PointsPrograms with (NoLock) " & _
                                        "where AdjustmentUPC='" & AdjustmentUPC & "';"
                    rst = MyCommon.LRT_Select()
                    If rst.Rows.Count = 0 Then
                        ValidSingleUPC = True
                    End If
                End If
            Else
                If AdjustmentUPC = "" Then
                    ValidSingleUPC = True
                End If
            End If
        Else
            ValidSetUPC = True
            ValidSingleUPC = True
        End If
        RedemptionRestrictionID = MyCommon.Extract_Val(Request.QueryString("redemptionRestrictionID"))
        MemberRedemptionID = MyCommon.Extract_Val(Request.QueryString("memberredemptionID"))
        ReturnHandlingTypeID = MyCommon.Extract_Val(Request.QueryString("returnsHandling"))
        'Handle the fuel partner flag
        If bAllowFuelPartner AndAlso SVTypeID = 1 Then
            bFuelPartner = (Request.QueryString("fuelpartner") = "1")
        End If
        Dim tempProgramExpire As Integer = 0
        Integer.TryParse(IIf(ProgramExpire <= 0, 1, ProgramExpire), tempProgramExpire)
        MyCommon.QueryStr = "dbo.pt_StoredValuePrograms_Insert"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 200).Value = ProgramName
        MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 1000).Value = ProgramDescription
        MyCommon.LRTsp.Parameters.Add("@ExpirePeriod", SqlDbType.Int).Value = tempProgramExpire
        MyCommon.LRTsp.Parameters.Add("@Value", SqlDbType.NVarChar, 200).Value = ProgramValue
        MyCommon.LRTsp.Parameters.Add("@OneUnitPerRec", SqlDbType.Bit).Value = IIf(StoreInUnits, 1, 0)
        MyCommon.LRTsp.Parameters.Add("@SVExpireType", SqlDbType.Int).Value = svExpireType
        MyCommon.LRTsp.Parameters.Add("@SVExpirePeriodType", SqlDbType.Int).Value = svExpirePeriodType
        MyCommon.LRTsp.Parameters.Add("@ExpireTOD", SqlDbType.VarChar, 5).Value = svExpireTOD
        If (IsDate(svExpireDateStr)) Then
            MyCommon.LRTsp.Parameters.Add("@ExpireDate", SqlDbType.DateTime).Value = Date.Parse(svExpireDateStr)
        End If
        MyCommon.LRTsp.Parameters.Add("@ExpireCentralServerTZ", SqlDbType.Bit).Value = IIf(ExpireCentralServerTZ, 1, 0)
        MyCommon.LRTsp.Parameters.Add("@SVTypeID", SqlDbType.Int).Value = SVTypeID
        MyCommon.LRTsp.Parameters.Add("@UOMLimit", SqlDbType.Int).Value = UomLimit
        MyCommon.LRTsp.Parameters.Add("@AllowReissue", SqlDbType.Int).Value = AllowReissue
        MyCommon.LRTsp.Parameters.Add("@ScorecardID", SqlDbType.Int).Value = ScorecardID
        MyCommon.LRTsp.Parameters.Add("@ScorecardDesc", SqlDbType.NVarChar, 100).Value = ScorecardDesc
        MyCommon.LRTsp.Parameters.Add("@ScorecardBold", SqlDbType.Bit).Value = 1 ' Hardcoding to 1; to restore flexibility, change back to IIf(ScorecardBold, 1, 0)
        MyCommon.LRTsp.Parameters.Add("@DisallowRedeemInEarnTrans", SqlDbType.Int).Value = IIf(Request.QueryString("disallowRedeemInTrans") = "1", 1, 0)
        MyCommon.LRTsp.Parameters.Add("@AllowNegativeBal", SqlDbType.Int).Value = IIf(Request.QueryString("allowNegBal") = "1", 1, 0)
        If (CPEInstalled) AndAlso (POSVAdjs) AndAlso (SVTypeID <> 2 AndAlso SVTypeID <> 4) Then
            If AllDigits(AdjustmentUPC) Then
                If AdjustmentUPC = "" OrElse CDec(AdjustmentUPC) = 0 Then
                    MyCommon.LRTsp.Parameters.Add("@AdjustmentUPC", SqlDbType.NVarChar, 100).Value = ""
                Else
                    MyCommon.LRTsp.Parameters.Add("@AdjustmentUPC", SqlDbType.NVarChar, 100).Value = AdjustmentUPC
                End If
            Else
                MyCommon.LRTsp.Parameters.Add("@AdjustmentUPC", SqlDbType.NVarChar, 100).Value = ""
            End If
        ElseIf UEInstalled Then
            MyCommon.LRTsp.Parameters.Add("@ReturnHandlingTypeID", SqlDbType.Int).Value = ReturnHandlingTypeID
            MyCommon.LRTsp.Parameters.Add("@AdjustmentUPC", SqlDbType.NVarChar, 100).Value = ""
        Else
            MyCommon.LRTsp.Parameters.Add("@AdjustmentUPC", SqlDbType.NVarChar, 100).Value = ""
        End If
        MyCommon.LRTsp.Parameters.Add("@RedemptionRestrictionID", SqlDbType.Int).Value = RedemptionRestrictionID
        MyCommon.LRTsp.Parameters.Add("@MemberRedemptionID", SqlDbType.Int).Value = MemberRedemptionID

        ' Save the first 30 characters of the name as the external ID when feature
        ' is enabled, SV type is Points and Expire Type is fixed date/time
        If bAllowExpirationExtension AndAlso SVTypeID = 1 AndAlso svExpireType = 1 Then
            NewExtID = Left(ProgramName, 30)
            MyCommon.LRTsp.Parameters.Add("@ExtProgramID", SqlDbType.NVarChar, 30).Value = NewExtID
        End If
        MyCommon.LRTsp.Parameters.Add("@VisibleToCustomers", SqlDbType.Bit).Value = IIf(VisibleToCustomers, 1, 0)
        MyCommon.LRTsp.Parameters.Add("@SVProgramID", SqlDbType.BigInt).Direction = ParameterDirection.Output

        If (ProgramName = "") Then
            infoMessage = Copient.PhraseLib.Lookup("sv-no-storeValueName", LanguageID)
        Else
            MyCommon.QueryStr = "SELECT SVProgramID FROM StoredValuePrograms with (NoLock) WHERE Name='" & MyCommon.Parse_Quotes(ProgramName) & "' AND Deleted=0;"
            rst = MyCommon.LRT_Select
            
            'Check SV Partition Max Expire Date
            Dim ExpireDateMax As Date = Microsoft.VisualBasic.Now.AddDays(366)
            Dim ExpireDate As Date
            'Fixed date
            If (svExpireType = 1) Then
                If (IsDate(svExpireDateStr)) Then
                    ExpireDate = Date.Parse(svExpireDateStr)
                End If
                'Period Days with exact time or Exact Period
            ElseIf (svExpireType = 2 OrElse svExpireType = 3) Then
                'Period type days
                If (svExpirePeriodType = 1) Then
                    ExpireDate = Microsoft.VisualBasic.Now.AddDays(ProgramExpire)
                    'Period type hours
                ElseIf (svExpirePeriodType = 2) Then
                    ExpireDate = Microsoft.VisualBasic.Now.AddHours(ProgramExpire)
                    'Period type Month
                ElseIf (svExpirePeriodType = 3) Then
                    ExpireDate = Microsoft.VisualBasic.Now.AddMonths(ProgramExpire)
                End If
                'New program so no offer is associated yet
            ElseIf (svExpireType = 4) Then
                'Default to 1 Day
                MyCommon.LRTsp.Parameters("@ExpirePeriod").Value = 1
                'X months after the end of the current month
            ElseIf (svExpireType = 5) Then
                ExpireDate = Microsoft.VisualBasic.Now.AddMonths(ProgramExpire)
            End If
      
            If (svExpireType <> 4) Then
                If (DateTime.Compare(ExpireDate, ExpireDateMax)) > 0 Then
                    infoMessage = "Stored value program expire date " & ExpireDate.ToString("MM\/dd\/yyyy") & " is greater than max expire date " & ExpireDateMax.ToString("MM\/dd\/yyyy") & " ."
                End If
            End If
            Dim culture As New CultureInfo("en-US")
            Dim numInfo As NumberFormatInfo = culture.NumberFormat
            Dim decimalProgramValue As Decimal = Decimal.Parse(ProgramValue, numInfo)
            If (rst.Rows.Count > 0) Then
                infoMessage = Copient.PhraseLib.Lookup("sv-name-used", LanguageID)
            ElseIf (decimalProgramValue < 0) Then
                infoMessage = Copient.PhraseLib.Lookup("sv.negative", LanguageID)
            ElseIf (SVTypeID = 1 AndAlso decimalProgramValue <> MyCommon.MakeInt(ProgramValue, -1)) Then
                infoMessage = Copient.PhraseLib.Lookup("sv.integer", LanguageID)
            ElseIf (decimalProgramValue <> 0.01 AndAlso (SVTypeID = 2 OrElse SVTypeID = 3)) Then
                infoMessage = Copient.PhraseLib.Lookup("sv.onecent", LanguageID)
            ElseIf (decimalProgramValue <> 0.001 AndAlso (SVTypeID = 4 OrElse SVTypeID = 5)) Then
                infoMessage = Copient.PhraseLib.Lookup("sv.onemill", LanguageID)
            ElseIf (Convert.ToInt32(Convert.ToDouble(ProgramExpire)) = 0 AndAlso (svExpireType = 2 OrElse svExpireType = 3)) Then
                infoMessage = Copient.PhraseLib.Lookup("sv.zeroexpire", LanguageID)
            ElseIf (ValidSetUPC = False AndAlso (SVTypeID <> 2 AndAlso SVTypeID <> 4)) Then
                infoMessage = Copient.PhraseLib.Detokenize("sv-edit.UPCOutside", LanguageID, BeginUPC, EndUPC)
            ElseIf (ValidSingleUPC = False AndAlso (SVTypeID <> 2 AndAlso SVTypeID <> 4)) Then
                infoMessage = Copient.PhraseLib.Lookup("sv-edit.UPCInUse", LanguageID)
            ElseIf (infoMessage <> "")
            Else
                MyCommon.LRTsp.ExecuteNonQuery()
                ProgramID = MyCommon.LRTsp.Parameters("@SVProgramID").Value

                If UEInstalled Then
                    AllowAnyCustomer_UE = MyCommon.NZ(Request.QueryString("hdnAllowAnyCustomerUE"), False)
                    MyCommon.QueryStr = "dbo.pt_SVProgramsPromoEngineSettings_InsertUpdate"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = 9
                    MyCommon.LRTsp.Parameters.Add("@SVProgramID", SqlDbType.BigInt).Value = ProgramID
                    MyCommon.LRTsp.Parameters.Add("@AllowAnyCustomer", SqlDbType.Bit).Value = AllowAnyCustomer_UE
                    MyCommon.LRTsp.Parameters.Add("@PKID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                    MyCommon.LRTsp.ExecuteNonQuery()
                    AllowAnyCustomer_UE_PKID = MyCommon.LRTsp.Parameters("@PKID").Value
                    MyCommon.Close_LRTsp()
                End If
                'save program name multilanguage input
                MLI.ItemID = l_pgID
                MLI.MLTableName = "SVProgramTranslations"
                MLI.MLColumnName = "ProgramName"
                MLI.MLIdentifierName = "SVProgramID"
                MLI.StandardTableName = "StoredValuePrograms"
                MLI.StandardColumnName = "Name"
                MLI.StandardIdentifierName = "SVProgramID"
                MLI.InputName = "name"
                Localization.SaveTranslationInputs(MyCommon, MLI, Request.QueryString, 2)
                'Save multilanguage inputs
                MLI.ItemID = ProgramID
                MLI.MLTableName = "SVProgramTranslations"
                MLI.MLColumnName = "ScorecardDesc"
                MLI.MLIdentifierName = "SVProgramID"
                MLI.StandardTableName = "StoredValuePrograms"
                MLI.StandardColumnName = "ScorecardDesc"
                MLI.StandardIdentifierName = "SVProgramID"
                MLI.InputName = "ScorecardDesc"
                Localization.SaveTranslationInputs(MyCommon, MLI, Request.QueryString, 2)

                MyCommon.Activity_Log(26, ProgramID, AdminUserID, Copient.PhraseLib.Lookup("history-sv-create", LanguageID))
            End If
        End If

        If bFuelPartner AndAlso SVTypeID = 1 Then
            MyCommon.QueryStr = "Update StoredValuePrograms Set FuelPartner=1 where SVProgramID=" & ProgramID & ";"
            MyCommon.LRT_Execute()
        End If

        'Perform necessary functions for the fuel partner interface
        If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) AndAlso (MyCommon.Fetch_CM_SystemOption(70) = "1") Then
            'Save or update information for third party transaction table
            If SVTypeID = 1 AndAlso _
               (Request.QueryString("RequiredPoints") <> "") AndAlso (Request.QueryString("MonetaryValue") <> "") AndAlso (IsNumeric(Request.QueryString("RequiredPoints"))) _
               AndAlso (IsNumeric(Request.QueryString("MonetaryValue"))) Then
                MyCommon.QueryStr = "select * from StoredValuePointsConversion where SVProgramID=" & ProgramID & ";"
                rst = MyCommon.LXS_Select()
                If rst.Rows.Count > 0 Then
                    MyCommon.QueryStr = "Update StoredValuePointsConversion set SVMonetaryValue=" & Request.QueryString("MonetaryValue") & ", SVPointsValue=" & _
                                                              Request.QueryString("RequiredPoints") & " where SVProgramID=" & ProgramID & ";"
                    MyCommon.LXS_Execute()
                Else
                    MyCommon.QueryStr = "Insert into dbo.StoredValuePointsConversion(SVProgramID, SVMonetaryValue, SVPointsValue) values (" & _
                                                              ProgramID & ", " & Request.QueryString("MonetaryValue") & ", " & Request.QueryString("RequiredPoints") & ");"
                    MyCommon.LXS_Execute()
                End If
            End If
            If ((Request.QueryString("RequiredPoints") = "") AndAlso (Request.QueryString("MonetaryValue") <> "")) OrElse _
               ((Request.QueryString("RequiredPoints") <> "") AndAlso (Request.QueryString("MonetaryValue") = "")) Then
                infoMessage = "RequiredPoints or MonetaryValue is empty"
            End If
            If ((Request.QueryString("RequiredPoints") <> "") AndAlso (Request.QueryString("MonetaryValue") <> "")) AndAlso _
               Not ((IsNumeric(MyCommon.Extract_Val(Request.QueryString("RequiredPoints")))) AndAlso (IsNumeric(MyCommon.Extract_Val(Request.QueryString("RequiredPoints"))))) Then
                infoMessage = "RequiredPoints or MonetaryValue value is invalid1"
            End If
            If ((Request.QueryString("RequiredPoints") <> "") OrElse (Request.QueryString("MonetaryValue") <> "")) AndAlso SVTypeID <> 1 Then
                infoMessage = "Third party conversion data can only be saved with a points stored value program"
            End If
        End If 'MyCommon.Fetch_CM_SystemOption(70) = "1"

        MyCommon.Close_LRTsp()

        ' we may need something similar for CM leaving for now
        'MyCommon.QueryStr = "dbo.pc_PointsVar_Create"
        'MyCommon.Open_LXSsp()
        'MyCommon.LXSsp.Parameters.Add("@ProgramID", SqlDbType.BigInt).Value = ProgramID
        'MyCommon.LXSsp.Parameters.Add("@VarID", SqlDbType.BigInt).Direction = ParameterDirection.Output
        'If (ProgramName = "") Then
        '    infoMessage = Copient.PhraseLib.Lookup("point-edit.noname", LanguageID)
        'Else
        '    MyCommon.LXSsp.ExecuteNonQuery()
        'End If
        'PromoVarID = MyCommon.LXSsp.Parameters("@VarID").Value
        'MyCommon.Close_LXSsp()

        'If (ProgramName = "") Then
        '    infoMessage = Copient.PhraseLib.Lookup("point-edit.noname", LanguageID)
        'Else
        '    MyCommon.QueryStr = "UPDATE PointsPrograms SET " & _
        '                        "Description = N'" & ProgramDescription & "', " & _
        '                        "PromoVarID = " & PromoVarID & _
        '                        "WHERE ProgramID = " & ProgramID
        '    MyCommon.LRT_Execute()
        'End If

        If infoMessage <> "" Then
            Response.Status = "301 Moved Permanently"
            Response.AddHeader("Location", "sv-edit.aspx?ProgramGroupID=" & ProgramID & "&infoMessage=" & infoMessage)
        ElseIf (infoMessage = "" AndAlso ProgramID > 0) Then
            Response.Status = "301 Moved Permanently"
            Response.AddHeader("Location", "sv-edit.aspx?ProgramGroupID=" & ProgramID)
        End If

        'If (ProgramName = "") Then
        '  infoMessage = Copient.PhraseLib.Lookup("sv-no-name", LanguageID)
        'Else
        '  Response.Status = "301 Moved Permanently"
        '  If (ProgramID = 0) Then
        '    If SVTypeID = 1 AndAlso ProgramValue < 0 Then
        '      Response.AddHeader("Location", "sv-edit.aspx?ProgramGroupID=" & ProgramID & "&infoMessage=" & Copient.PhraseLib.Lookup("sv.negative", LanguageID))
        '    Else
        '      Response.AddHeader("Location", "sv-edit.aspx?ProgramGroupID=" & ProgramID & "&infoMessage=" & Copient.PhraseLib.Lookup("sv-name-used", LanguageID))
        '    End If
        '  Else
        '    Response.AddHeader("Location", "sv-edit.aspx?ProgramGroupID=" & ProgramID)
        '  End If
        '  GoTo done
        'End If

    ElseIf (((Request.QueryString("save") <> "") Or (Request.QueryString("SaveProp") <> "") Or (Request.QueryString("SavePropDeploy") <> "") Or (Request.QueryString("SavePropDeployExtend") <> "")) AndAlso MyCommon.Extract_Val(Request.QueryString("ProgramGroupID")) > 0) Then
        ' somebody clicked save
        l_pgID = MyCommon.Extract_Val(Request.QueryString("ProgramGroupID"))
        ProgramName = MyCommon.Parse_Quotes(Request.QueryString.Item("name"))
        ProgramName = Logix.TrimAll(ProgramName)
        ProgramDescription = MyCommon.Parse_Quotes(Request.QueryString.Item("desc"))
        ProgramValue = MyCommon.Parse_Quotes(MyCommon.Extract_Decimal(Request.QueryString("value"), MyCommon.GetAdminUser.Culture))
        svExpireType = MyCommon.Extract_Val(Request.QueryString("expiretype"))
        ProgramExpire = IIf(svExpireType = 1, 1, MyCommon.Parse_Quotes(MyCommon.Extract_Val(Request.QueryString("expireperiod"))))
        svExpirePeriodType = MyCommon.Extract_Val(Request.QueryString("expireperiodtype"))
        svExpireTOD = IIf(svExpireType <> 2, "", Request.QueryString("expiretod"))
        svExpireDateStr = IIf(svExpireType <> 1, "", GetCgiValue("expiredatetime"))
        If String.IsNullOrWhiteSpace(svExpireDateStr) = False Then
            Dim svExpireDateTime As DateTime
            DateTime.TryParse(svExpireDateStr, MyCommon.GetAdminUser.Culture, DateTimeStyles.None, svExpireDateTime)
            If svExpireDateTime < DateTime.Now AndAlso svExpireType = 1 Then
                infoMessage = Copient.PhraseLib.Lookup("logix-js.EnterValidExpDate", LanguageID).Replace("&#39;", "\'")
            End If
            svExpireDateStr = svExpireDateTime.ToString(CultureInfo.CurrentCulture)
        End If
        If Not ValidateMonthExpiry(ProgramExpire, svExpirePeriodType) Then
            infoMessage = Copient.PhraseLib.Lookup("error.invalidsvmonthexpiry", LanguageID)
        End If
        SVTypeID = MyCommon.Extract_Val(Request.QueryString("svtypeid"))

        'Perform necessary functions for the fuel partner interface
        If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) AndAlso (MyCommon.Fetch_CM_SystemOption(70) = "1") Then
            'Save or update information for third party transaction table
            If SVTypeID = 1 AndAlso (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) AndAlso (MyCommon.Fetch_CM_SystemOption(70) = "1") Then
                MyCommon.QueryStr = "select * from StoredValuePointsConversion where SVProgramID=" & l_pgID & ";"
                rst = MyCommon.LXS_Select()
                If rst.Rows.Count > 0 Then
                    If ((Request.QueryString("RequiredPoints") = "") AndAlso (Request.QueryString("MonetaryValue") = "")) Then
                        'Row Exists, but we need to delete it
                        MyCommon.QueryStr = "Delete from StoredValuePointsConversion where SVProgramID=" & l_pgID & ";"
                        MyCommon.LXS_Execute()
                    Else
                        'Row Exists, update it
                        MyCommon.QueryStr = "Update StoredValuePointsConversion set SVMonetaryValue=" & Request.QueryString("MonetaryValue") & ", SVPointsValue=" & _
                                                                  Request.QueryString("RequiredPoints") & " where SVProgramID=" & l_pgID & ";"
                        MyCommon.LXS_Execute()
                    End If
                Else
                    MyCommon.QueryStr = "Insert into dbo.StoredValuePointsConversion(SVProgramID, SVMonetaryValue, SVPointsValue) values (" & _
                                                              l_pgID & ", " & Request.QueryString("MonetaryValue") & ", " & Request.QueryString("RequiredPoints") & ");"
                    MyCommon.LXS_Execute()
                End If
            End If
            If ((Request.QueryString("RequiredPoints") = "") AndAlso (Request.QueryString("MonetaryValue") <> "")) OrElse _
               ((Request.QueryString("RequiredPoints") <> "") AndAlso (Request.QueryString("MonetaryValue") = "")) Then
                infoMessage = "RequiredPoints or MonetaryValue is empty"
            End If
            If (Request.QueryString("RequiredPoints") <> "") AndAlso IsNumeric(Request.QueryString("RequiredPoints")) OrElse _
                ((Request.QueryString("RequiredPoints") = "") AndAlso (Request.QueryString("MonetaryValue") = "")) Then
            Else
                infoMessage = "RequiredPoints or MonetaryValue value is invalid2 Reqd:" & Request.QueryString("RequiredPoints")
            End If
            If (Request.QueryString("MonetaryValue") <> "") AndAlso IsNumeric(Request.QueryString("MonetaryValue")) OrElse _
                ((Request.QueryString("RequiredPoints") = "") AndAlso (Request.QueryString("MonetaryValue") = "")) Then
            Else
                infoMessage = "RequiredPoints or MonetaryValue value is invalid3 Monetary:" & Request.QueryString("MonetaryValue")
            End If
        End If 'MyCommon.Fetch_CM_SystemOption(70) = "1"

        If (SVTypeID = 0) Then
            SVTypeID = 2
        End If

        If (SVTypeID = 3 OrElse SVTypeID = 5 OrElse (SVTypeID = 1 And bEditUomLimit)) Then
            UomLimit = MyCommon.Extract_Val(Request.QueryString("uomlimit"))
            If (UomLimit <= 0) Then
                UomLimit = 1
            End If
        Else
            UomLimit = 1
        End If
        ExpireCentralServerTZ = (Request.QueryString("expirecentralserverTZ") = "on")
        AllowReissue = (Request.QueryString("allowreissue") = "1")
        AutoDelete = (Request.QueryString("autodelete") = "1")
        ScorecardID = MyCommon.Extract_Val(Request.QueryString("ScorecardID"))
        ScorecardDesc = MyCommon.Parse_Quotes(Request.QueryString("ScorecardDesc"))
        If (CPEInstalled) AndAlso (POSVAdjs) AndAlso (SVTypeID <> 2 AndAlso SVTypeID <> 4) Then
            BeginUPC = CDec(MyCommon.Fetch_CPE_SystemOption(100))
            EndUPC = CDec(MyCommon.Fetch_CPE_SystemOption(101))
            If AdjustmentUPC = "" Then
                ValidSetUPC = True
            ElseIf Not AllDigits(AdjustmentUPC) Then
                ValidSetUPC = False
            Else
                If (CDec(AdjustmentUPC) < 0) Then
                    ValidSetUPC = False
                Else
                    'Is UPC allowed to be outside the set of UPCs
                    If MyCommon.Fetch_CPE_SystemOption(102) = "0" Then
                        UPCDec = CDec(AdjustmentUPC)
                        If UPCDec > (BeginUPC - 1) AndAlso UPCDec < (EndUPC + 1) Then
                            ValidSetUPC = True
                        End If
                    Else
                        ValidSetUPC = True
                    End If
                End If
            End If
            'The UPC is not allowed to be on more than one program
            If AllDigits(AdjustmentUPC) Then
                If AdjustmentUPC = "" OrElse CDec(AdjustmentUPC) = 0 Then
                    ValidSingleUPC = True
                Else
                    MyCommon.QueryStr = "select SVProgramID As ID from StoredValuePrograms with (NoLock) " & _
                                        "where AdjustmentUPC='" & AdjustmentUPC & "' and SVProgramID<>" & l_pgID & " " & _
                                        "union " & _
                                        "select ProgramID as ID from PointsPrograms with (NoLock) " & _
                                        "where AdjustmentUPC='" & AdjustmentUPC & "';"
                    rst = MyCommon.LRT_Select()
                    If rst.Rows.Count = 0 Then
                        ValidSingleUPC = True
                    End If
                End If
            Else
                If AdjustmentUPC = "" Then
                    ValidSingleUPC = True
                End If
            End If
        Else
            ValidSetUPC = True
            ValidSingleUPC = True
        End If
        RedemptionRestrictionID = MyCommon.Extract_Val(Request.QueryString("redemptionRestrictionID"))
        MemberRedemptionID = MyCommon.Extract_Val(Request.QueryString("memberredemptionID"))

        If bAllowFuelPartner AndAlso SVTypeID = 3 Then
            bFuelPartner = (Request.QueryString("fuelpartner") = "1")
            bAutoRedeem = (Request.QueryString("autoredeem") = "1")
            bAllowAdjustments = (Request.QueryString("allowadjust") = "1")
        ElseIf bAllowFuelPartner AndAlso SVTypeID = 1 Then
            bFuelPartner = (Request.QueryString("fuelpartner") = "1")
        Else
            bFuelPartner = False
            bAutoRedeem = False
            bAllowAdjustments = False
        End If

        ' If we don't know the expiration type yet, look it up
        If bAllowExpirationExtension AndAlso svExpireType = 0 Then
            Dim dt As DataTable
            MyCommon.QueryStr = "SELECT SVExpireType FROM StoredValuePrograms WITH (NoLock) WHERE Deleted=0 AND SVProgramID='" & l_pgID & "';"
            dt = MyCommon.LRT_Select
            svExpireType = MyCommon.NZ(dt.Rows(0).Item("SVExpireType"), 2)
        End If

        With Request.QueryString
            'Check SV Partition Max Expire Date
            Dim ExpireDateMax As Date = Microsoft.VisualBasic.Now.AddDays(366)
            Dim ExpireDate As Date
            'Fixed date
            If (svExpireType = 1) Then
                If (IsDate(svExpireDateStr)) Then
                    ExpireDate = Date.Parse(svExpireDateStr)
                End If
                'Period Days with exact time or Exact Period
            ElseIf (svExpireType = 2 OrElse svExpireType = 3) Then
                'Period type days
                If (svExpirePeriodType = 1) Then
                    ExpireDate = Microsoft.VisualBasic.Now.AddDays(ProgramExpire)
                    'Period type hours
                ElseIf (svExpirePeriodType = 2) Then
                    ExpireDate = Microsoft.VisualBasic.Now.AddHours(ProgramExpire)
                    'Period type Month
                ElseIf (svExpirePeriodType = 3) Then
                    ExpireDate = Microsoft.VisualBasic.Now.AddMonths(ProgramExpire)
                End If
                'X Days after earn offer
            ElseIf (svExpireType = 4) Then
                'use the incentive id from dstAssociated offers list to check max expire date
                If (dstAssociated.Rows.Count > 0) Then
                    For Each row In dstAssociated.Rows
                        MyCommon.QueryStr = "Select I.EndDate From CPE_Incentives as I " & _
                                            "Inner Join CPE_RewardOptions as RO on I.IncentiveID = RO.IncentiveID " & _
                                            "Inner Join CPE_Deliverables as D on RO.RewardOptionID = D.RewardOptionID " & _
                                            "Where D.DeliverableTypeID = 11 and I.IncentiveID = " & row.Item("OfferID").ToString & ";"
                        rst = MyCommon.LRT_Select
                        If (rst.Rows.Count > 0) Then
                            For Each row1 In rst.Rows
                                ExpireDate = Date.Parse(row1.Item("EndDate").ToString())
                                ExpireDate = ExpireDate.AddDays(ProgramExpire)
                                If (DateTime.Compare(ExpireDate, ExpireDateMax)) > 0 Then
                                    infoMessage = "Stored value program expire date is greater than " & ExpireDateMax.ToString("MM\/dd\/yyyy") & ", due to Offer " & row.Item("OfferID").ToString & "."
                                End If
                            Next
                        End If
                        If (infoMessage <> "") Then Exit For
                    Next
                End If
                'X months after the end of the current month
            ElseIf (svExpireType = 5) Then
                ExpireDate = Microsoft.VisualBasic.Now.AddMonths(ProgramExpire)
            End If
      
            If (svExpireType <> 4) Then
                If (DateTime.Compare(ExpireDate, ExpireDateMax)) > 0 Then
                    infoMessage = "Stored value program expire date " & ExpireDate.ToString("MM\/dd\/yyyy") & " is greater than max expire date " & ExpireDateMax.ToString("MM\/dd\/yyyy") & " ."
                End If
            End If
            Dim culture As New CultureInfo("en-US")
            Dim numInfo As NumberFormatInfo = culture.NumberFormat
            Dim decimalProgramValue As Decimal = Decimal.Parse(ProgramValue, numInfo)
            Dim isInteger As Boolean = IsNumeric(decimalProgramValue)
            If MyCommon.Parse_Quotes(.Item("name")) = "" Then
                infoMessage = Copient.PhraseLib.Lookup("sv-no-storeValueName", LanguageID)
            ElseIf (decimalProgramValue < 0) Then
                infoMessage = Copient.PhraseLib.Lookup("sv.negative", LanguageID)
            ElseIf (SVTypeID = 1 AndAlso isInteger <> True) Then
                infoMessage = Copient.PhraseLib.Lookup("sv.integer", LanguageID)
            ElseIf (decimalProgramValue <> 0.01 AndAlso (SVTypeID = 2 OrElse SVTypeID = 3)) Then
                infoMessage = Copient.PhraseLib.Lookup("sv.onecent", LanguageID)
            ElseIf (decimalProgramValue <> 0.001 AndAlso (SVTypeID = 4 OrElse SVTypeID = 5)) Then
                infoMessage = "ProgramValue=" & decimalProgramValue & " | " & Copient.PhraseLib.Lookup("sv.onemill", LanguageID)
            ElseIf (Convert.ToInt32(Convert.ToDouble(ProgramExpire)) = 0 AndAlso (svExpireType = 2 OrElse svExpireType = 3)) Then
                infoMessage = Copient.PhraseLib.Lookup("sv.zeroexpire", LanguageID)
            ElseIf (ValidSetUPC = False AndAlso (SVTypeID <> 2 AndAlso SVTypeID <> 4)) Then
                infoMessage = Copient.PhraseLib.Detokenize("sv-edit.UPCOutside", LanguageID, BeginUPC, EndUPC)
            ElseIf (ValidSingleUPC = False AndAlso (SVTypeID <> 2 AndAlso SVTypeID <> 4)) Then
                infoMessage = Copient.PhraseLib.Lookup("sv-edit.UPCInUse", LanguageID)
            ElseIf (infoMessage <> "")
            Else
                MyCommon.QueryStr = "SELECT SVProgramID,Name,Description,ExpirePeriod,Value FROM StoredValuePrograms with (NoLock) WHERE Name = '" & MyCommon.Parse_Quotes(.Item("name")) & "' AND Deleted=0 AND SVProgramID <> " & l_pgID
                rst = MyCommon.LRT_Select
                If (ProgramName = "") Then
                    infoMessage = Copient.PhraseLib.Lookup("sv-no-storeValueName", LanguageID)
                ElseIf (rst.Rows.Count > 0) Then
                    infoMessage = Copient.PhraseLib.Lookup("sv-name-used", LanguageID)
                ElseIf (MyCommon.Extract_Val(Request.QueryString("ScorecardID")) > 0) AndAlso ((Request.QueryString("ScorecardDesc") = "" And Not MultiLanguageEnabled) OrElse (Request.QueryString("ScorecardDesc") = "" And Request.QueryString("ScorecardDesc_" & DefaultLanguageCode) = "" And MultiLanguageEnabled)) Then
                    infoMessage = Copient.PhraseLib.Lookup("sv-edit.EnterScorecard", LanguageID)
                Else
                    MyCommon.QueryStr = "UPDATE StoredValuePrograms with (RowLock) SET " &
                                        "Name=N'" & ProgramName & "'," &
                                        "Value=" & ProgramValue & ", " &
                                        "ExpirePeriod=" & ProgramExpire & ", " &
                                        "Description=N'" & ProgramDescription & "', " &
                                        "OneUnitPerRec=" & IIf(Request.QueryString("storeInUnits") = "1", 1, 0) & ", " &
                                        "SVExpireType=" & IIf(Request.QueryString("expiretype") = "", "SVExpireType", svExpireType) & ", " &
                                        "SVExpirePeriodType=" & IIf(Request.QueryString("expireperiodtype") = "", "SVExpirePeriodType", svExpirePeriodType) & ", " &
                                                  "ExpireTOD='" & svExpireTOD & "', " &
                                        "SVTypeID=" & SVTypeID & ", UnitOfMeasureLimit=" & UomLimit & ", AllowReissue=" & IIf(AllowReissue, 1, 0) & ", " &
                                        "AutoDelete=" & IIf(AutoDelete, 1, 0) & ", " &
                                        "ExpireCentralServerTZ=" & IIf(ExpireCentralServerTZ, 1, 0) & ", " &
                                        "ScorecardID=" & Request.QueryString("ScorecardID") & ", " &
                                        "ScorecardDesc=N'" & MyCommon.Parse_Quotes(.Item("ScorecardDesc")) & "', " &
                                        "ScorecardBold=" & IIf(Request.QueryString("ScorecardBold") = "1", 1, 0) & ", " &
                                        "VisibleToCustomers=" & IIf(MyCommon.Parse_Quotes(Logix.TrimAll(.Item("VisibleToCustomers"))) = "1", 1, 0) & ", "
                    If IsDate(svExpireDateStr) Then
                        MyCommon.QueryStr &= "ExpireDate='" & svExpireDateStr & "', "
                    Else
                        MyCommon.QueryStr &= "ExpireDate=NULL, "
                    End If
                    If (CPEInstalled) AndAlso (POSVAdjs) AndAlso (SVTypeID <> 2 AndAlso SVTypeID <> 4) Then
                        If AllDigits(AdjustmentUPC) Then
                            If AdjustmentUPC = "" OrElse CDec(AdjustmentUPC) = 0 Then
                                MyCommon.QueryStr &= "AdjustmentUPC=NULL, "
                            Else
                                MyCommon.QueryStr &= "AdjustmentUPC=N'" & AdjustmentUPC & "', "
                            End If
                        Else
                            MyCommon.QueryStr &= "AdjustmentUPC=NULL, "
                        End If
                    Else
                        MyCommon.QueryStr &= "AdjustmentUPC=NULL, "
                    End If
                    MyCommon.QueryStr &= "RedemptionRestrictionID=" & RedemptionRestrictionID & ", "
                    MyCommon.QueryStr &= "MemberRedemptionID=" & MemberRedemptionID & ", "
                    If UEInstalled Then
                        MyCommon.QueryStr &= "ReturnHandlingTypeID=" & IIf(Request.QueryString("returnsHandling") <> "", MyCommon.Extract_Val(Request.QueryString("returnsHandling")), 1) & ", " & _
                                              "DisallowRedeemInEarnTrans=" & IIf(Request.QueryString("disallowRedeemInTrans") = "1", 1, 0) & ", " & _
                                              "AllowNegativeBal=" & IIf(Request.QueryString("allowNegBal") = "1", 1, 0) & ", "
                    End If

                    If bAllowFuelPartner AndAlso SVTypeID = 3 Then
                        MyCommon.QueryStr &= "FuelPartner=" & IIf(bFuelPartner, 1, 0) & ", "
                        MyCommon.QueryStr &= "AutoRedeem=" & IIf(bAutoRedeem, 1, 0) & ", "
                        MyCommon.QueryStr &= "AllowAdjustments=" & IIf(bAllowAdjustments, 1, 0) & ", "
                    End If
                    If bAllowFuelPartner AndAlso SVTypeID = 1 Then
                        MyCommon.QueryStr &= "FuelPartner=" & IIf(bFuelPartner, 1, 0) & ", "
                    End If

                    ' Save the first 30 characters of the name as the external ID when feature
                    ' is enabled, SV type is Points and Expire Type is fixed date/time
                    If bAllowExpirationExtension AndAlso SVTypeID = 1 AndAlso svExpireType = 1 Then
                        NewExtID = Left(ProgramName, 30)
                        MyCommon.QueryStr &= "ExtProgramID=N'" & NewExtID & "', "
                    End If

                    MyCommon.QueryStr &= "LastUpdate=getDate(), CMOAStatusFlag=1, CPEStatusFlag=1 " & _
                                         "WHERE SVProgramID=" & MyCommon.Parse_Quotes(.Item("ProgramGroupID"))
                    MyCommon.LRT_Execute()

                    If UEInstalled Then
                        AllowAnyCustomer_UE = MyCommon.NZ(Request.QueryString("hdnAllowAnyCustomerUE"), False)
                        MyCommon.QueryStr = "dbo.pt_SVProgramsPromoEngineSettings_InsertUpdate"
                        MyCommon.Open_LRTsp()
                        MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = 9
                        MyCommon.LRTsp.Parameters.Add("@SVProgramID", SqlDbType.BigInt).Value = l_pgID
                        MyCommon.LRTsp.Parameters.Add("@AllowAnyCustomer", SqlDbType.Bit).Value = AllowAnyCustomer_UE
                        MyCommon.LRTsp.Parameters.Add("@PKID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                        MyCommon.LRTsp.ExecuteNonQuery()
                        AllowAnyCustomer_UE_PKID = MyCommon.LRTsp.Parameters("@PKID").Value
                        MyCommon.Close_LRTsp()
                    End If
                    'save program name multilanguage input
                    MLI.ItemID = l_pgID
                    MLI.MLTableName = "SVProgramTranslations"
                    MLI.MLColumnName = "ProgramName"
                    MLI.MLIdentifierName = "SVProgramID"
                    MLI.StandardTableName = "StoredValuePrograms"
                    MLI.StandardColumnName = "Name"
                    MLI.StandardIdentifierName = "SVProgramID"
                    MLI.InputName = "name"
                    Localization.SaveTranslationInputs(MyCommon, MLI, Request.QueryString, 2)
                    'Save multilanguage inputs
                    MLI.ItemID = l_pgID
                    MLI.MLTableName = "SVProgramTranslations"
                    MLI.MLColumnName = "ScorecardDesc"
                    MLI.MLIdentifierName = "SVProgramID"
                    MLI.StandardTableName = "StoredValuePrograms"
                    MLI.StandardColumnName = "ScorecardDesc"
                    MLI.StandardIdentifierName = "SVProgramID"
                    MLI.InputName = "ScorecardDesc"
                    Localization.SaveTranslationInputs(MyCommon, MLI, Request.QueryString, 2)
                    MyCommon.QueryStr = " update CPE_DeliverableMonSVTranslations with (RowLock) SET Description =N'" & ProgramDescription & "' WHERE SVProgramID=" & l_pgID & " AND LanguageID=" & LanguageID & " ;"
                    MyCommon.LRT_Execute()
                    MyCommon.Activity_Log(26, l_pgID, AdminUserID, Copient.PhraseLib.Lookup("history.sv-edit", LanguageID))
                End If
            End If
        End With

        If (Request.QueryString("SaveProp") <> "") Then
            If (Not drsCM Is Nothing) AndAlso (drsCM.Length > 0) Then
                MyCommon.QueryStr = "dbo.pa_CM_PropagateStoredValues"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@ProgramID", SqlDbType.BigInt).Value = l_pgID
                MyCommon.LRTsp.Parameters.Add("@Redeploy", SqlDbType.Bit).Value = 0
                MyCommon.LRTsp.Parameters.Add("@UserId", SqlDbType.Int).Value = AdminUserID
                MyCommon.LRTsp.Parameters.Add("@UpdateHist", SqlDbType.VarChar, 100).Value = Copient.PhraseLib.Lookup("history.sv-update", LanguageID)
                MyCommon.LRTsp.Parameters.Add("@DeployHist", SqlDbType.VarChar, 100).Value = Copient.PhraseLib.Lookup("history.sv-deploy", LanguageID)
                MyCommon.LRTsp.Parameters.Add("@UpdateNum", SqlDbType.Int).Direction = ParameterDirection.Output
                MyCommon.LRTsp.Parameters.Add("@DeployNum", SqlDbType.Int).Direction = ParameterDirection.Output
                MyCommon.LRTsp.ExecuteNonQuery()
                OfferUpdateNum = MyCommon.LRTsp.Parameters("@UpdateNum").Value
                OfferDeployNum = MyCommon.LRTsp.Parameters("@DeployNum").Value
                MyCommon.Close_LRTsp()
            End If
            statusMessage = OfferUpdateNum & " " & Copient.PhraseLib.Lookup("sv-edit.OffersUpdated", LanguageID)

        ElseIf (Request.QueryString("SavePropDeploy") <> "") Then
            If (Not drsCM Is Nothing) AndAlso (drsCM.Length > 0) Then
                MyCommon.QueryStr = "dbo.pa_CM_PropagateStoredValues"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@ProgramID", SqlDbType.BigInt).Value = l_pgID
                MyCommon.LRTsp.Parameters.Add("@Redeploy", SqlDbType.Bit).Value = 1
                MyCommon.LRTsp.Parameters.Add("@UserId", SqlDbType.Int).Value = AdminUserID
                MyCommon.LRTsp.Parameters.Add("@UpdateHist", SqlDbType.VarChar, 100).Value = Copient.PhraseLib.Lookup("history.sv-update", LanguageID)
                MyCommon.LRTsp.Parameters.Add("@DeployHist", SqlDbType.VarChar, 100).Value = Copient.PhraseLib.Lookup("history.sv-deploy", LanguageID)
                MyCommon.LRTsp.Parameters.Add("@UpdateNum", SqlDbType.Int).Direction = ParameterDirection.Output
                MyCommon.LRTsp.Parameters.Add("@DeployNum", SqlDbType.Int).Direction = ParameterDirection.Output
                MyCommon.LRTsp.ExecuteNonQuery()
                OfferUpdateNum = MyCommon.LRTsp.Parameters("@UpdateNum").Value
                OfferDeployNum = MyCommon.LRTsp.Parameters("@DeployNum").Value
                MyCommon.Close_LRTsp()
            End If
            statusMessage = OfferUpdateNum & " " & Copient.PhraseLib.Lookup("sv-edit.OffersUpdated", LanguageID)
            statusMessage = statusMessage & " " & OfferDeployNum & " " & Copient.PhraseLib.Lookup("sv-edit.OffersDeployed", LanguageID)

        ElseIf (Request.QueryString("SavePropDeployExtend") <> "") Then
            MyCommon.QueryStr = "dbo.pa_ExtendStoredValues"
            MyCommon.Open_LXSsp()
            MyCommon.LXSsp.Parameters.Add("@ProgramID", SqlDbType.BigInt).Value = l_pgID
            MyCommon.LXSsp.Parameters.Add("@ExpireDate", SqlDbType.DateTime).Value = Date.Parse(svExpireDateStr)
            MyCommon.LXSsp.ExecuteNonQuery()
            MyCommon.Close_LXSsp()

            MyCommon.QueryStr = "dbo.pa_CM_PropagateStoredValues"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@ProgramID", SqlDbType.BigInt).Value = l_pgID
            MyCommon.LRTsp.Parameters.Add("@Redeploy", SqlDbType.Bit).Value = 1
            MyCommon.LRTsp.Parameters.Add("@UserId", SqlDbType.Int).Value = AdminUserID
            MyCommon.LRTsp.Parameters.Add("@UpdateHist", SqlDbType.VarChar, 100).Value = Copient.PhraseLib.Lookup("history.sv-update", LanguageID)
            MyCommon.LRTsp.Parameters.Add("@DeployHist", SqlDbType.VarChar, 100).Value = Copient.PhraseLib.Lookup("history.sv-deploy", LanguageID)
            MyCommon.LRTsp.Parameters.Add("@UpdateNum", SqlDbType.Int).Direction = ParameterDirection.Output
            MyCommon.LRTsp.Parameters.Add("@DeployNum", SqlDbType.Int).Direction = ParameterDirection.Output
            MyCommon.LRTsp.ExecuteNonQuery()

            OfferUpdateNum = MyCommon.LRTsp.Parameters("@UpdateNum").Value
            OfferDeployNum = MyCommon.LRTsp.Parameters("@DeployNum").Value
            MyCommon.Close_LRTsp()

            statusMessage = OfferUpdateNum & " " & Copient.PhraseLib.Lookup("sv-edit.OffersUpdated", LanguageID)
            statusMessage = statusMessage & " " & OfferDeployNum & " " & Copient.PhraseLib.Lookup("sv-edit.OffersDeployed", LanguageID)
        End If

    ElseIf (Request.QueryString("ProgramGroupID") <> "") Then
        ' simple edit/search mode
        l_pgID = MyCommon.NZ(Request.QueryString("ProgramGroupID"), "0")
    ElseIf (Request.Form("ProgramID") <> "") Then
        l_pgID = MyCommon.Extract_Val(Request.Form("ProgramID"))
    Else
        ' no group id passed ... what now ?
        l_pgID = "0"
    End If

    ' grab this stored value program
    MyCommon.QueryStr = "SELECT SVProgramID, Name, SVP.Description, Value, ValuePrecision, CreatedDate, LastUpdate, SVP.Description, OneUnitPerRec, AutoDelete, " &
                        "SVExpireType, SVExpirePeriodType, ExpirePeriod, ExpireTOD, ExpireDate, SVP.SVTypeID, UnitofMeasureLimit, AllowReissue, ExpireCentralServerTZ, ExtProgramID, " &
                        "ScorecardID, ScorecardDesc, ScorecardBold, AdjustmentUPC, RedemptionRestrictionID, ReturnHandlingTypeID, DisallowRedeemInEarnTrans, " &
                        "AllowNegativeBal, MemberRedemptionID, FuelPartner, AutoRedeem, AllowAdjustments, SVP.VisibleToCustomers " &
                        "FROM StoredValuePrograms AS SVP WITH (NoLock) " &
                        "LEFT JOIN SVTypes AS SVT WITH (NoLock) ON SVT.SVTypeID=SVP.SVTypeID " &
                        "WHERE Deleted=0 AND SVProgramID='" & l_pgID & "';"
    dstPrograms = MyCommon.LRT_Select
    If (dstPrograms.Rows.Count > 0) Then
        pgName = MyCommon.NZ(dstPrograms.Rows(0).Item("Name"), "")
        pgCreated = MyCommon.NZ(dstPrograms.Rows(0).Item("CreatedDate"), "")
        pgUpdated = MyCommon.NZ(dstPrograms.Rows(0).Item("LastUpdate"), "")
        ' pgPromoVarID = MyCommon.NZ(dstPrograms.Rows(0).Item("PromoVarID"), 0)
        pgDescription = MyCommon.NZ(dstPrograms.Rows(0).Item("Description"), 0)
        pgValue = (Convert.ToDecimal(dstPrograms.Rows(0).Item("Value"))).ToString(CultureInfo.InvariantCulture)
        pgValuePrecision = MyCommon.NZ(dstPrograms.Rows(0).Item("ValuePrecision"), 0)
        pgExpire = MyCommon.NZ(dstPrograms.Rows(0).Item("ExpirePeriod"), 1)
        l_pgID = MyCommon.NZ(dstPrograms.Rows(0).Item("SVProgramID"), 0)
        StoreInUnits = MyCommon.NZ(dstPrograms.Rows(0).Item("OneUnitPerRec"), False)
        AutoDelete = MyCommon.NZ(dstPrograms.Rows(0).Item("AutoDelete"), True)
        svExpireType = MyCommon.NZ(dstPrograms.Rows(0).Item("SVExpireType"), 2)
        svExpirePeriodType = MyCommon.NZ(dstPrograms.Rows(0).Item("SVExpirePeriodType"), 0)
        svExpireDateStr = MyCommon.NZ(dstPrograms.Rows(0).Item("ExpireDate"), "")
        SVTypeID = MyCommon.NZ(dstPrograms.Rows(0).Item("SVTypeID"), 2)
        UomLimit = MyCommon.NZ(dstPrograms.Rows(0).Item("UnitofMeasureLimit"), 1)
        ExpireCentralServerTZ = MyCommon.NZ(dstPrograms.Rows(0).Item("ExpireCentralServerTZ"), False)
        AllowReissue = MyCommon.NZ(dstPrograms.Rows(0).Item("AllowReissue"), False)
        VisibleToCustomers = MyCommon.NZ(dstPrograms.Rows(0).Item("VisibleToCustomers"), 0)
        If (svExpireDateStr <> "") Then
            svExpireDate = dstPrograms.Rows(0).Item("ExpireDate")
            svExpireHr = svExpireDate.TimeOfDay.Hours
            svExpireMin = svExpireDate.TimeOfDay.Minutes
            svExpireDateStr = Logix.ToShortDateString(svExpireDate, MyCommon)
        End If
        svExpireTOD = MyCommon.NZ(dstPrograms.Rows(0).Item("ExpireTOD"), "")
        If (svExpireTOD <> "") Then
            Dim tokens As String() = svExpireTOD.Split(":")
            If (tokens.Length = 2) Then
                svExpireTODHr = tokens(0)
                svExpireTODMin = tokens(1)
            End If
        End If
        sExtProgramID = MyCommon.NZ(dstPrograms.Rows(0).Item("ExtProgramID"), "")
        ScorecardID = MyCommon.NZ(dstPrograms.Rows(0).Item("ScorecardID"), 0)
        ScorecardDesc = MyCommon.NZ(dstPrograms.Rows(0).Item("ScorecardDesc"), "")
        ScorecardBold = MyCommon.NZ(dstPrograms.Rows(0).Item("ScorecardBold"), False)
        AdjustmentUPC = MyCommon.NZ(dstPrograms.Rows(0).Item("AdjustmentUPC"), "")
        RedemptionRestrictionID = MyCommon.NZ(dstPrograms.Rows(0).Item("RedemptionRestrictionID"), 0)
        ReturnHandlingTypeID = MyCommon.NZ(dstPrograms.Rows(0).Item("ReturnHandlingTypeID"), 1)
        DisallowRedeemInTrans = MyCommon.NZ(dstPrograms.Rows(0).Item("DisallowRedeemInEarnTrans"), 0)
        AllowNegativeBal = MyCommon.NZ(dstPrograms.Rows(0).Item("AllowNegativeBal"), 0)
        MemberRedemptionID = MyCommon.NZ(dstPrograms.Rows(0).Item("MemberRedemptionID"), 0)
        ' Let's see if any stored value deliverables using this program have ScorecardID and ScorecardDesc set
        MyCommon.QueryStr = "select ScorecardID, ScorecardDesc from CPE_DeliverableStoredValue with (NoLock) " & _
                            "where SVProgramID=" & l_pgID & " and ScorecardID>0;"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
            RewardsSetScorecard = True
        End If
        If bAllowFuelPartner AndAlso SVTypeID = 3 Then
            bFuelPartner = MyCommon.NZ(dstPrograms.Rows(0).Item("FuelPartner"), False)
            bAutoRedeem = MyCommon.NZ(dstPrograms.Rows(0).Item("AutoRedeem"), False)
            bAllowAdjustments = MyCommon.NZ(dstPrograms.Rows(0).Item("AllowAdjustments"), False)
        ElseIf bAllowFuelPartner AndAlso SVTypeID = 1 Then
            bFuelPartner = MyCommon.NZ(dstPrograms.Rows(0).Item("FuelPartner"), False)
        Else
            bFuelPartner = False
            bAutoRedeem = False
            bAllowAdjustments = False
        End If

        ' User can extend expiration for Points SV with a fixed date/time expiration
        ' If customer SV points exist, SVExtensionAgent will update customer records
        If bAllowExpirationExtension AndAlso SVTypeID = 1 AndAlso SVExpireType = 1
            MyCommon.QueryStr = "select SVProgramID from StoredValue where SVProgramID = @SVProgramID union " & _
                                "select SVProgramID from SVHistory where SVProgramID = @SVProgramID"
            MyCommon.DBParameters.Add("@SVProgramID", SqlDbType.BigInt).Value = l_pgID
            Dim dt As DataTable = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)
            If (dt.Rows.Count > 0) Then
                bProgramActive = True
            End If
        End If

        If UEInstalled Then
            If AllowAnyCustomer_UE_PKID = 0 Then
                MyCommon.QueryStr = "SELECT PKID, AllowAnyCustomer from SVProgramsPromoEngineSettings " & _
                            "WHERE SVProgramID = @SVProgramID AND EngineID = @EngineID"
                MyCommon.DBParameters.Add("@SVProgramID", SqlDbType.BigInt).Value = l_pgID
                MyCommon.DBParameters.Add("@EngineID", SqlDbType.Int).Value = 9
                Dim dt As DataTable = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                If (dt.Rows.Count > 0) Then
                    AllowAnyCustomer_UE_PKID = CMS.Utilities.NZ(dt.Rows(0).Item("PKID"), 0)
                    AllowAnyCustomer_UE = CMS.Utilities.NZ(dt.Rows(0).Item("AllowAnyCustomer"), False)
                End If
            End If
            If (AllowAnyCustomer_UE) Then
                MyCommon.QueryStr = "dbo.pa_IsAnyCustomerOffersExistForSVProgam"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = 9
                MyCommon.LRTsp.Parameters.Add("@SVProgramID", SqlDbType.BigInt).Value = l_pgID
                MyCommon.LRTsp.Parameters.Add("@result", SqlDbType.Bit).Direction = ParameterDirection.Output
                MyCommon.LRTsp.ExecuteNonQuery()
                IsAnyCustomerOffersExist_UE = MyCommon.LRTsp.Parameters("@result").Value
                MyCommon.Close_LRTsp()
            End If
        End If
    ElseIf (Request.QueryString("new") <> "New") And (l_pgID > 0) Then
        ' check if this is a deleted stored value program
        MyCommon.QueryStr = "select Name from StoredValuePrograms with (NoLock) where SVProgramID=" & l_pgID & " and deleted =1"
        rst = MyCommon.LRT_Select()
        If (rst.Rows.Count > 0) Then
            pgName = MyCommon.NZ(rst.Rows(0).Item("Name"), "")
        Else
            pgName = ""
        End If

        Send_HeadBegin("term.storedvalue", , l_pgID)
        Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
        Send_Metas()
        Send_Links(Handheld)
        Send_Scripts(New String() {"datePicker.js"})
        Send_HeadEnd()
        Send_BodyBegin(1)
        Send_Bar(Handheld)
        Send_Logos()
        Send_Tabs(Logix, 5)
        Send_Subtabs(Logix, 52, 5, , l_pgID)
        Send("")
        Send("<div id=""intro"">")
        Send("  <h1 id=""title"">" & Copient.PhraseLib.Lookup("term.storedvalue", LanguageID) & " #" & l_pgID & ": " & pgName & "</h1>")
        Send("</div>")
        Send("<div id=""main"">")
        Send("  <div id=""infobar"" class=""red-background"">")
        Send("    " & Copient.PhraseLib.Lookup("error.deleted", LanguageID))
        Send("  </div>")
        Send("</div>")
        Send_BodyEnd()
        GoTo done
    Else
        pgPromoVarID = 0
        l_pgID = "0"
        pgDescription = ""
        pgCreated = ""
        pgUpdated = ""
        pgName = ""
    End If

    CpeEngineOnly = MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) And _
                Not (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) And _
                Not (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.Catalina))
    Send_HeadBegin("term.storedvalue", , l_pgID)
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send("<style type=""text/css"">")
    Send("#expire div {")
    Send("  margin: 0 0 5px 10px;")
    Send("}")
    Send("</style>")
    Send_Scripts(New String() {"datePicker.js"})
    Send_HeadEnd()
    Send_BodyBegin(1)
    Send_Bar(Handheld)
    Send_Help(CopientFileName)
    Send_Logos()
    Send_Tabs(Logix, 5)
    Send_Subtabs(Logix, 52, 5, , l_pgID)

    If (Logix.UserRoles.AccessStoredValuePrograms = False) Then
        Send_Denied(1, "perm.storedvalue-access")
        Send_BodyEnd()
        GoTo done
    End If
%>
<script type="text/javascript">
  var datePickerDivID = "datepicker";
  var buttonClicked = "";
  
  $(document).ready(function () {
      WireUpEvents();   
  });
  //Register all the events in this function
  function WireUpEvents() {
      $('#expireperiod').on("focusout", handleExpirePeriodFocusOut)
      //$(document).on("submit", handleExpirePeriodFocusOut)
  }
  function handleExpirePeriodFocusOut()
  {
      var alertText = "<%= Copient.PhraseLib.Lookup("error.invalidsvmonthexpiry", LanguageID)%>";
        //AMS-4379: For Months do not allow value > 120
      if($('#expireperiodtype option:selected').val() == 3 && this.value > 120)
      {
          alert(alertText);
          this.focus();
          //return false; //prevent ongoing event from continuing
      }
  }
  <% Send_Calendar_Overrides(MyCommon) %>

  <% If (Logix.UserRoles.EditStoredValuePrograms And l_pgID > 0) Then %>
  window.onunload = function() {
    if (document.mainform.name.value != document.mainform.name.defaultValue || document.mainform.desc.value != document.mainform.desc.defaultValue) {
      saveChanges = confirm('<%Sendb(Copient.PhraseLib.Lookup("sv-edit.ChangesMade", LanguageID))%>');
      if (saveChanges) {
        if (document.mainform.elements['Save'] == null) {
          saveElem = document.createElement("input");
          saveElem.type = 'hidden';
          saveElem.id = 'Save';
          saveElem.name = 'Save';
          saveElem.value = 'save';
          document.mainform.appendChild(saveElem);
        }
        handleAutoFormSubmit();
      } 
    }
  };
  <% End If %>
  if (window.captureEvents){
    window.captureEvents(Event.CLICK);
    window.onclick=handlePageClick;
  } else {
    document.onclick=handlePageClick;
  }
  
  function handlePageClick(e) {
    var calFrame = document.getElementById('calendariframe');
    var el=(typeof event!=='undefined')? event.srcElement : e.target
    
    if (el != null) {
      var pickerDiv = document.getElementById(datePickerDivID);
      if (pickerDiv != null && pickerDiv.style.visibility == "visible") {
        if (el.id!="expire-date-picker") {
          if (!isDatePickerControl(el.className)) {
            pickerDiv.style.visibility = "hidden";
            pickerDiv.style.display = "none";
            if (calFrame != null) {
              calFrame.style.visibility = 'hidden';
              calFrame.style.display = "none";
            }
          }
        } else  {
          pickerDiv.style.visibility = "visible";            
          pickerDiv.style.display = "block";            
          if (calFrame != null) {
            calFrame.style.visibility = 'visible';
            calFrame.style.display = "block";
          }
        }
      }
    }
    if (el.id != 'actions') {
      if (document.getElementById("actionsmenu") != null) {
        var  bOpen = (document.getElementById("actionsmenu").style.visibility == 'visible');
        if (bOpen) {
          toggleDropdown();
        }
      }
    }
  }
  
  function isDatePickerControl(ctrlClass) {
    var retVal = false;
    
    if (ctrlClass != null && ctrlClass.length >= 2) {
      if (ctrlClass.substring(0,2) == "dp") {
        retVal = true;
      }
    }
    
    return retVal;
  }
  
  function setButtonClicked(btnName) {
    buttonClicked = btnName;
  }
  
  // callback function for save changes on unload during navigate away
  function handleAutoFormSubmit() {
    window.onunload = null;
    document.mainform.action="SV-edit.aspx"
    document.mainform.submit();
  }
  
  function handleSubmit() {
    var retVal = true;
    var elem = document.getElementById("expiretype");
    var elemExPdType = document.getElementById("expireperiodtype");
    
    if (buttonClicked=='save') {
      retVal = checkEmptyValuesBeforeSubmit();
      if (! retVal) {
        return false;
        }
      retVal = checkBeforeSubmit();
      retVal = retVal && validateExpiration();
      
      if (retVal) {
        window.onunload = null;
        formatExpTOD();
        formatExpDate();
        if (elemExPdType != null) {
          elemExPdType.disabled = false;
        }
        if (elem != null) {
          elem.disabled = false;
        }
      }
    }
    
    return retVal;
  }
  
  function checkBeforeSubmit() {
    var elemBulk = document.getElementById("storeInUnits");
    var retVal = true;
    
    if (elemBulk != null) {
      if (elemBulk.checked) {
        retVal = confirm('<%Sendb(Copient.PhraseLib.Lookup("storedvalue.disable-note", LanguageID))%>');
      }
    }
    return retVal;
  }

  function checkEmptyValuesBeforeSubmit() {
    var elemExPrd = document.getElementById("expireperiod");
    var elemPrName = document.getElementById("name");
    var elem = document.getElementById("expiretype");
    var retVal = true;
    
    if (elem.value =="2" || elem.value=="3" || elem.value =="4" || elem.value=="5") {
    if (elemExPrd != null) {
      if (elemExPrd.value == "") {
        alert('<%Sendb(Copient.PhraseLib.Lookup("sv-edit.InvalidExpirePeriod", LanguageID).Replace("&#39;", "\'"))%>');
        elemExPrd.focus();
        return false;
      }
    }
    }
    

    if (elemPrName != null) {
      if (elemPrName.value == "") {
        alert('<%Sendb(Copient.PhraseLib.Lookup("sv-no-storeValueName", LanguageID))%>');
        elemPrName.focus();
        return false;
      }
    }

    return retVal;
  }
  
  function handleExpireType() {
    var elem = document.getElementById("expiretype");
    var elemExPdType = document.getElementById("expireperiodtype");
    var elemExPd = document.getElementById("expireperiod");
    var elemExDate = document.getElementById("expiredate");
    var elemExTod = document.getElementById("expiretod");
    
    if (elem != null) {
      if (elem.options[elem.options.selectedIndex].value == "1") {
        if (elemExPdType != null) {
          elemExPdType.disabled = false;
        }
        document.getElementById("trExpPdType").style.display = "none";
        document.getElementById("trExpPd").style.display = "none";
        document.getElementById("trExpTod").style.display = "none";
        document.getElementById("trExpDate").style.display = "block";
        document.getElementById("trExpTime").style.display = "block";
      } else if (elem.options[elem.options.selectedIndex].value=="2") {
        if (elemExPdType != null) {
          elemExPdType.options[0].selected = true;
          elemExPdType.disabled = true;
        }
        document.getElementById("trExpPdType").style.display = "block";
        document.getElementById("trExpPd").style.display = "block";
        document.getElementById("trExpTod").style.display = "block";
        document.getElementById("trExpDate").style.display = "none";
        document.getElementById("trExpTime").style.display = "none";
      } else if (elem.options[elem.options.selectedIndex].value=="3") {
        document.getElementById("trExpPdType").style.display = "block";
        if (elemExPdType != null) {
          elemExPdType.disabled = false;
        }
        document.getElementById("trExpPd").style.display = "block";
        document.getElementById("trExpTod").style.display = "none";
        document.getElementById("trExpDate").style.display = "none";
        document.getElementById("trExpTime").style.display = "none";
      } else if (elem.options[elem.options.selectedIndex].value=="4") {
        if (elemExPdType != null) {
          elemExPdType.options[0].selected = true;
          elemExPdType.disabled = true;
        }
        document.getElementById("trExpPdType").style.display = "block";
        document.getElementById("trExpPd").style.display = "block";
        document.getElementById("trExpTod").style.display = "none";
        document.getElementById("trExpDate").style.display = "none";
        document.getElementById("trExpTime").style.display = "none";
      } else if (elem.options[elem.options.selectedIndex].value=="5") {
        if (elemExPdType != null) {
          elemExPdType.options[2].selected = true;
          elemExPdType.disabled = true;
        }
        document.getElementById("trExpPdType").style.display = "block";
        document.getElementById("trExpPd").style.display = "block";
        document.getElementById("trExpTod").style.display = "none";
        document.getElementById("trExpDate").style.display = "none";
        document.getElementById("trExpTime").style.display = "none";
      }
    }
    <% If (bProgramActive) Then %>
       elem.disabled = true;
    <% End If %>
  }

   // You can only extend the expiration so make sure new values are greater than old
   function handleExpireDateTimeChange() {
      var elemDateTime = document.getElementById("expiredatetime");
      var elemExDate = document.getElementById("expiredate");
      var elemHour = document.getElementById("expTodHours");
      var elemMin = document.getElementById("expTodMinutes");
      
      var oldDate = new Date(elemDateTime.value);
      var newDate = new Date(elemExDate.value);
      var oldHour = '<%= svExpireTODHr %>';
      var oldMin =  '<%= svExpireTODMin %>';

      if ((oldDate < newDate) ||
          ((oldDate.getTime() === newDate.getTime()) &&
           ((oldHour < elemHour) ||
            (oldHour == elemHour) && (oldMin < elemMin)))) {
         // Good new date - make sure customer records haven't been purged
         var purgeLen =  '<%= MyCommon.Fetch_SystemOption(45) %>';
         var curDate = new Date();
         if (oldDate.getTime() + purgeLen < curDate.getTime()) {
            // No user data to extend, use regular buttons and display message
            document.getElementById("save").disabled = false;
            document.getElementById("SaveProp").disabled = false;
            document.getElementById("SavePropDeploy").disabled = false;
            document.getElementById("SavePropDeployExtend").disabled = true;
            alert('<%Sendb(Copient.PhraseLib.Lookup("sv-edit.ExpirationExtensionAlreadyPurged", LanguageID).Replace("&#39;", "\'"))%>');
            return false;
         }
         document.getElementById("save").disabled = true;
         document.getElementById("SaveProp").disabled = true;
         document.getElementById("SavePropDeploy").disabled = true;
         document.getElementById("SavePropDeployExtend").disabled = false;
         return true;
      } else if ((oldDate.getTime() === newDate.getTime()) &&
                 (svExpireTod == elemExTod)) {
         // Date unchanged
         document.getElementById("save").disabled = false;
         document.getElementById("SaveProp").disabled = false;
         document.getElementById("SavePropDeploy").disabled = false;
         document.getElementById("SavePropDeployExtend").disabled = true;
         return true;
      }
      
      // New expire date is before old one:  error
      document.getElementById("save").disabled = false;
      document.getElementById("SaveProp").disabled = false;
      document.getElementById("SavePropDeploy").disabled = false;
      document.getElementById("SavePropDeployExtend").disabled = true;
      alert('<%Sendb(Copient.PhraseLib.Lookup("sv-edit.InvalidExpirationExtension", LanguageID).Replace("&#39;", "\'"))%>');

      return false;      
  }
  
  function validateExpiration() {
    var elem = document.getElementById("expiretype");
    var elemExPd = document.getElementById("expireperiod");
    var elemExDate = document.getElementById("expiredate");
    var retVal = true;

    if (elem != null) {
      if (elem.value == "1") {
        if (elemExDate != null) {
		    retVal = IsValidLocalizedDate(elemExDate.value, '<% Sendb(MyCommon.GetAdminUser.Culture.DateTimeFormat.ShortDatePattern) %>');
       	    if (!retVal) {
            elemExDate.focus();
            elemExDate.select();
            }
        }
      } else if (elem.value =="2" || elem.value=="3" || elem.value =="4" || elem.value=="5") {
        if (elemExPd != null) {
          retVal = !isNaN(elemExPd.value.replace(",", "."));
          if (!retVal) {
            alert('<%Sendb(Copient.PhraseLib.Lookup("sv-edit.InvalidExpirePeriod", LanguageID).Replace("&#39;", "\'"))%>');
            elemExPd.focus();
            elemExPd.select();
          }
		  else {
		   if (elemExPd.value.replace(",", ".") < 0) {
		     retVal = !retVal;
             alert('<%Sendb(Copient.PhraseLib.Lookup("sv-edit.InvalidExpirePeriod", LanguageID).Replace("&#39;", "\'"))%>');
             elemExPd.focus();
             elemExPd.select();		     
		   }
		  }		  
        }
      }
    }
    return retVal;    
  }
  
  function formatExpTOD() {
    var elem = document.getElementById("expiretype");
    var elemTod = document.getElementById("expiretod");
    var elemHour = document.getElementById("expTodHours");
    var elemMin = document.getElementById("expTodMinutes");
    
    if (elem != null && elem.value == 2) {    
      if (elemTod != null && elemHour != null && elemMin != null) {
        elemTod.value = elemHour.value + ":" + elemMin.value;
      }
    } else if (elem != null && elemTod != null) {
      elemTod.value = "";
    }
  }
  
  function formatExpDate() {
    var elem = document.getElementById("expiretype");
    var elemDate = document.getElementById("expiredate");
    var elemDateTime = document.getElementById("expiredatetime");
    var elemHour = document.getElementById("expHours");
    var elemMin = document.getElementById("expMinutes");
    
    if (elem != null && elem.value == 1) {    
      if (elemDate != null && elemDateTime != null && elemHour != null && elemMin != null) {
        //var dtPartArr = elemDate.value.split("/");
        elemDateTime.value = elemDate.value +  " " + elemHour.value + ":" + elemMin.value; //dtPartArr[2] + "-" + dtPartArr[0] + "-" + dtPartArr[1] + " " + elemHour.value + ":" + elemMin.value; 
      }
    } else if (elem != null && elemDate != null && elemDateTime != null)  {
      elemDate.value = "";
      elemDateTime = ""
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
  
  function handleSVType(typeID, pgID) {
    var elemDiv = document.getElementById("uomDiv");
    var valueInput = document.getElementById("value");
    var uovInput = document.getElementById("uovInput");
    var uovSummary = document.getElementById("uovSummary");
    var upcDiv = document.getElementById("ManualAdjustmentUPC");
    
    if (elemDiv != null) {
      if (typeID == 3 || typeID == 5) {
        elemDiv.style.display = 'inline';
      } else {
        elemDiv.style.display = 'none';
      }
    }
    var NumberDecimalSeparator = '<% Sendb(MyCommon.GetAdminUser.Culture.NumberFormat.NumberDecimalSeparator) %>'
    if (pgID == 0) {
      if (typeID == 2 || typeID == 3) {
        valueInput.value = "0" + NumberDecimalSeparator + "01";
        uovSummary.innerHTML = "$0" + NumberDecimalSeparator + "01";
        uovInput.style.display = 'none';
        uovSummary.style.display = 'inline';
      } else if (typeID == 4 || typeID == 5) {
        valueInput.value = "0" + NumberDecimalSeparator + "001";
        uovSummary.innerHTML = "$0" + NumberDecimalSeparator + "001";
        uovInput.style.display = 'none';
        uovSummary.style.display = 'inline';
      } else {
        valueInput.value = "1";
        uovSummary.innerHTML = "1";
        uovInput.style.display = 'inline';
        uovSummary.style.display = 'none';
      }
    }
    <%If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE)) AndAlso (MyCommon.Fetch_SystemOption(94) = 1) Then%>
    if (typeID == 2 || typeID == 4) {
      upcDiv.style.display = 'none';
    } else {
      upcDiv.style.display = '';
    }
    <%End If%>
  }
  
  function toggleScorecardText() {
    if (document.getElementById("ScorecardID").value == 0) {
      document.getElementById("scdesc").style.display = 'none';
      document.getElementById("ScorecardDesc").value = '';
    } else {
      document.getElementById("scdesc").style.display = '';
    }
  }
  
  function selectPLU(idLen) {
    var PLUelem = document.getElementById("AdjustmentUPC");
    var selElem = document.getElementById("selector");
    
    if (PLUelem != null && selElem != null) {
      PLUelem.value = padLeft(selElem.options[selElem.selectedIndex].value, idLen);
    }
  }
  
  function padLeft(str, totalLength) {
    var pd = '';
    
    str = str.toString();
    if (totalLength > str.length) {
      for (var i=0; i < (totalLength-str.length); i++) {
        pd += '0';
      }      
    }
    
    return pd + str.toString();
  }
  
  function handleFuelPartner() {
    var elemFuelPartner = document.getElementById("fuelpartner");
    if (elemFuelPartner != null) {
      if (elemFuelPartner.checked == true) {
        document.getElementById("autoredeem").disabled=false;
        document.getElementById("allowadjust").disabled=false;
      } else {
        document.getElementById("autoredeem").disabled=true;
        document.getElementById("allowadjust").disabled=true;
      }
    }
  }

  function updateInputVal(cbElement, hdnfield_id) {
    $("#"+hdnfield_id).val(cbElement.checked);
  }

  $(document).ready(function() {
    var AllowAnyCustomerUE = $("#cbAllowAnyCustomerUE")
    if (AllowAnyCustomerUE.length > 0) {
      updateInputVal(AllowAnyCustomerUE[0], "hdnAllowAnyCustomerUE");
    }
  });
  
</script>
<form action="#" method="get" id="mainform" name="mainform" onsubmit="return handleSubmit();">
<div id="intro">
  <h1 id="title">
    <%
      If l_pgID = 0 Then
        Sendb(Copient.PhraseLib.Lookup("term.newstoredvalue", LanguageID))
      Else
        Sendb(Copient.PhraseLib.Lookup("term.storedvalue", LanguageID) & " #" & l_pgID & ": ")
        Sendb(MyCommon.TruncateString(pgName, 40))
      End If
    %>
  </h1>
  <div id="controls">
    <%
      If (l_pgID = 0) Then
        If (Logix.UserRoles.CreateStoredValuePrograms) Then
          Send_Save(" onclick=""setButtonClicked('save');"" ")
        End If
      Else
        ShowActionButton = (Logix.UserRoles.CreateStoredValuePrograms) OrElse (Logix.UserRoles.EditStoredValuePrograms) OrElse (Logix.UserRoles.DeleteStoredValuePrograms)
        If (ShowActionButton) Then
          Send("<input type=""button"" class=""regular"" id=""actions"" name=""actions"" value=""" & Copient.PhraseLib.Lookup("term.actions", LanguageID) & " ▼"" onclick=""toggleDropdown();"" />")
          Send("<div class=""actionsmenu"" id=""actionsmenu"">")
          
          If (Logix.UserRoles.EditStoredValuePrograms) Then
            Send_Save(" onclick=""setButtonClicked('save');"" ")
          Else
            Send_Save(" onclick=""setButtonClicked('save');"" hidden")
          End If
          
          If (Logix.UserRoles.EditStoredValuePrograms) AndAlso (Logix.UserRoles.UpdateOffersUsingStoredValueProgram) AndAlso (Not drsCM Is Nothing) AndAlso (drsCM.Length > 0) Then
            Send_SV_Propagate(" onclick=""setButtonClicked('save');"" ")
          Else
            Send_SV_Propagate(" onclick=""setButtonClicked('save');"" hidden")
          End If
              
          If (Logix.UserRoles.EditStoredValuePrograms) AndAlso (Logix.UserRoles.RedeployOffersUsingStoredValueProgram) AndAlso (Not drsCM Is Nothing) AndAlso (drsCM.Length > 0) Then
            Send_SV_Deploy(" onclick=""setButtonClicked('save');"" ")
          Else
            Send_SV_Deploy(" onclick=""setButtonClicked('save');"" hidden")
          End If
              
          If  (Logix.UserRoles.EditStoredValuePrograms) AndAlso (Logix.UserRoles.RedeployOffersUsingStoredValueProgram) AndAlso (bProgramActive) Then
              Send_SV_Extend(" onclick=""setButtonClicked('save');"" ")
          Else
              Send_SV_Extend(" onclick=""setButtonClicked('save');"" hidden")
          End If
          
          If (Logix.UserRoles.DeleteStoredValuePrograms) Then
            Send_Delete(" onclick=""setButtonClicked('delete');"" ")
          End If
          
          If (Logix.UserRoles.CreateStoredValuePrograms) Then
            Send_New()
          End If
          
          If Request.Browser.Type = "IE6" Then
            Send("<iframe src=""javascript:'';"" id=""actionsiframe"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no"" style=""height:75px;""></iframe>")
          End If
          Send("</div>")
          If MyCommon.Fetch_SystemOption(75) Then
            If (Logix.UserRoles.AccessNotes) Then
              Send_NotesButton(9, l_pgID, AdminUserID)
            End If
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
<div id="main" <% Sendb(IE6ScrollFix) %>>
  <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
  <% If (statusMessage <> "") Then Send("<div id=""statusbar"" class=""green-background"">" & statusMessage & "</div>")%>
  <div id="column1">
    <div class="box" id="identity">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.identification", LanguageID))%>
        </span>
      </h2>
      <input type="hidden" id="ProgramGroupID" name="ProgramGroupID" value="<% sendb(l_pgID) %>" />
      <label for="name">
        <% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>:</label><br />
      <% If (pgName Is Nothing) Then pgName = ""%>
        <%If l_pgID = 0 Then
                %>
        <input type="text" class="longest" id="name" name="name" maxlength="50" value="<% Sendb(pgName.Replace("""", "&quot;")) %>" /><br />
        <%
            Else
                MLI.ItemID = l_pgID
                MLI.MLTableName = "SVProgramTranslations"
                MLI.MLColumnName = "ProgramName"
                MLI.MLIdentifierName = "SVProgramID"
                MLI.StandardTableName = "StoredValuePrograms"
                MLI.StandardColumnName = "Name"
                MLI.StandardIdentifierName = "SVProgramID"
                MLI.StandardValue = pgName
                MLI.InputName = "name"
                MLI.InputID = "name"
                MLI.InputType = "text"
                MLI.MaxLength = 50
                MLI.CSSStyle = "width:350px;"
                Send(Localization.SendTranslationInputs(MyCommon, MLI, Request.QueryString, 2))
            End If

        %>
      
      <label for="desc">
        <% Sendb(Copient.PhraseLib.Lookup("term.description", LanguageID))%>:</label><br />
      <textarea class="longest" id="desc" name="desc" cols="48" rows="3" maxlength="1000"><% Sendb(pgDescription.Trim)%></textarea><br />
      <br class="half" />
      <small>
        <%Send("(" & Copient.PhraseLib.Lookup("CPEoffergen.description", LanguageID) & ")")%></small><br
          class="half" />
      <br class="half" />
      <%
        If sExtProgramID <> "" Then
          Sendb(Copient.PhraseLib.Lookup("term.externalid", LanguageID) & ": " & sExtProgramID)
          Send("<br /><br class=""half"" />")
        End If
        If (l_pgID <> 0) Then
          MyCommon.QueryStr = "SELECT (SUM(Cast(QtyEarned as bigint)) - Sum(Cast(QtyUsed as bigint))) AS TotalSV, COUNT(Distinct CustomerPK) as Customers FROM StoredValue WITH (NOLOCK) WHERE SVProgramID=" & l_pgID & ";"
          dstSV = MyCommon.LXS_Select
          pgTotalSV = MyCommon.NZ(dstSV.Rows(0).Item(0), "0")
          Sendb(Copient.PhraseLib.Lookup("term.contains", LanguageID) & " ")
          Dim SVPhrase As String
          Dim CardholderPhrase As String
          If (pgTotalSV = 1) Then
            SVPhrase = StrConv(Copient.PhraseLib.Lookup("term.unit", LanguageID), VbStrConv.Lowercase)
          Else
            SVPhrase = StrConv(Copient.PhraseLib.Lookup("term.units", LanguageID), VbStrConv.Lowercase)
          End If
          If (MyCommon.NZ(dstSV.Rows(0).Item("Customers"), 0) = 1) Then
            CardholderPhrase = StrConv(Copient.PhraseLib.Lookup("term.cardholder", LanguageID), VbStrConv.Lowercase)
          Else
            CardholderPhrase = StrConv(Copient.PhraseLib.Lookup("term.cardholders", LanguageID), VbStrConv.Lowercase)
          End If
          Sendb(pgTotalSV & " " & SVPhrase)
          If pgTotalSV > 0 Then
            If SVTypeID > 1 Then
              Dim ValueString As String = ""              
              Dim TotalValue As Decimal = pgTotalSV * Convert.ToDecimal(pgValue)
              ValueString = Math.Round(TotalValue, Localization.Get_Default_Currency_Precision()).ToString(MyCommon.GetAdminUser.Culture) & " " & Localization.Get_Default_Currency_Symbol()
              Sendb(" (" & ValueString & ")")
            End If
            Sendb(" " & Copient.PhraseLib.Lookup("term.heldby", LanguageID) & " ")
            Sendb(MyCommon.NZ(dstSV.Rows(0).Item("Customers"), 0) & " " & CardholderPhrase)
          End If
          Send("<br /><br class=""half"" />")
          Send("<input type=""checkbox"" id=""autodelete"" name=""autodelete"" value=""1""" & IIf(AutoDelete, " checked=""checked""", "") & " /><label for=""autodelete"">" & Copient.PhraseLib.Lookup("sv-edit.AutoDelete", LanguageID) & "</label><br />")
        End If
      %>
      &nbsp;
      <hr class="hidden" />
    </div>
    <div class="box" id="general">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.general", LanguageID))%>
        </span>
      </h2>
      <%
        If bCmInstalled Then
          MyCommon.QueryStr = "select PhraseID, SVTypeID from SVTypes with (NoLock) order by SVTypeID;"
          rst = MyCommon.LRT_Select
          Send("<div" & IIf(l_pgID <> 0, " style=""display:none;""", "") & ">")
          Send("<label for=""svtypeid"" style=""position:relative;"">" & Copient.PhraseLib.Lookup("term.type", LanguageID) & ":</label> ")
          Send("<select id=""svtypeid"" name=""svtypeid"" onchange=""handleSVType(this.value," & l_pgID & ");"" >")
          For Each row In rst.Rows
            If (MyCommon.NZ(row.Item("SVTypeID"), -1) = SVTypeID) Then
              Send("<option value=""" & MyCommon.NZ(row.Item("SVTypeID"), -1) & """ selected=""selected"">" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "</option>")
            Else
              Send("<option value=""" & MyCommon.NZ(row.Item("SVTypeID"), -1) & """>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "</option>")
            End If
          Next
          Send("</select><br />")
          Send("</div>")
          If l_pgID <> 0 Then
            Sendb(Copient.PhraseLib.Lookup("term.type", LanguageID) & ": ")
            For Each row In rst.Rows
              If (MyCommon.NZ(row.Item("SVTypeID"), -1) = SVTypeID) Then
                If SVTypeID = 1 And bEditUomLimit Then
                  Send(Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & " (" & Copient.PhraseLib.Lookup("term.fuel", LanguageID) & ")" & "<br />")
                Else
                  Send(Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "<br />")
                End If
              End If
            Next
          End If
        ElseIf (MyCommon.IsEngineInstalled(2) Or MyCommon.IsEngineInstalled(6) Or MyCommon.IsEngineInstalled(9)) Then 'Check for CPE or CAM OR UE
          MyCommon.QueryStr = "select PhraseID, SVTypeID from SVTypes with (NoLock) where SVTypeID in (1,2"
          If (MyCommon.IsEngineInstalled(9)) Then
            MyCommon.QueryStr &= ") order by SVTypeID;"
          Else
            MyCommon.QueryStr &= ",4) order by SVTypeID;"
          End If
          
          rst = MyCommon.LRT_Select
          Send("<div" & IIf(l_pgID <> 0, " style=""display:none;""", "") & ">")
          Send("<label for=""svtypeid"" style=""position:relative;"">" & Copient.PhraseLib.Lookup("term.type", LanguageID) & ":</label> ")
          Send("<select id=""svtypeid"" name=""svtypeid"" onchange=""handleSVType(this.value," & l_pgID & ");"" >")
          For Each row In rst.Rows
            If (MyCommon.NZ(row.Item("SVTypeID"), -1) = SVTypeID) Then
              Send("<option value=""" & MyCommon.NZ(row.Item("SVTypeID"), -1) & """ selected=""selected"">" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "</option>")
            Else
              Send("<option value=""" & MyCommon.NZ(row.Item("SVTypeID"), -1) & """>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "</option>")
            End If
          Next
          Send("</select><br />")
          Send("</div>")
          If l_pgID <> 0 Then
            Sendb(Copient.PhraseLib.Lookup("term.type", LanguageID) & ": ")
            For Each row In rst.Rows
              If (MyCommon.NZ(row.Item("SVTypeID"), -1) = SVTypeID) Then
                Send(Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "<br />")
              End If
            Next
          End If
        Else
          Send("<input type=""hidden"" name=""svtypeid"" id=""svtypeid"" value=""1"" />")
          Sendb(Copient.PhraseLib.Lookup("term.type", LanguageID) & ": ")
          Send(Copient.PhraseLib.Lookup("term.points", LanguageID) & "<br />")
        End If
        Send("<div id=""uomDiv"" style=""position:relative;" & IIf((SVTypeID = 3 OrElse SVTypeID = 5 OrElse (SVTypeID = 1 And bEditUomLimit)), "", "display:none;") & """>")
        Send("<br class=""half"" />")
        Send(Copient.PhraseLib.Lookup("term.unitofmeasurelimit", LanguageID) & ": ")
        If l_pgID = 0 Then
          Send("<input type=""text"" class=""shorter"" id=""uomlimit"" name=""uomlimit"" value=""" & UomLimit & """ />")
        Else
          If SVTypeID <> 1 Then
            Send(UomLimit & "<br />")
            Send("<input type=""hidden"" id=""uomlimit"" name=""uomlimit"" value=""" & UomLimit & """ />")
          Else
            Send("<input type=""text"" class=""shorter"" id=""uomlimit"" name=""uomlimit"" value=""" & UomLimit & """ />")
          End If
        End If
        Send("</div>")
        Send("<br class=""half"" />")
        Send("<label for=""value"" style=""position:relative;"">" & Copient.PhraseLib.Lookup("term.unitvalue", LanguageID) & ": </label>")
        Dim currencySymbol As String = Localization.Get_Default_Currency_Symbol()
        If l_pgID = 0 Then
          Dim tempval As Decimal = MyCommon.NZ(pgValue, "0.01")              
          If (bCmInstalled Or MyCommon.IsEngineInstalled(2) Or MyCommon.IsEngineInstalled(9)) Then
            Sendb("<span id=""uovInput"" style=""position:relative;display:none;""><input type=""text"" class=""shorter"" id=""value"" name=""value"" maxlength=""6"" value=""" & tempval.ToString(MyCommon.GetAdminUser.Culture) & """ /></span>")
            Sendb("<span id=""uovSummary"" style=""position:relative;"">" & currencySymbol & tempval.ToString(MyCommon.GetAdminUser.Culture) & "</span>")
          Else
            tempval = MyCommon.NZ(pgValue, "1")
            Sendb("<span id=""uovInput"" style=""position:relative;""><input type=""text"" class=""shorter"" id=""value"" name=""value"" maxlength=""6"" value=""" & tempval.ToString(MyCommon.GetAdminUser.Culture) & """ /></span>")
          End If
        Else
          If SVTypeID = 1 Then
            Send(MyCommon.MakeInt(pgValue, 1) & "<br />")
            Send("<input type=""hidden"" id=""value"" name=""value"" value=""" & pgValue.ToString(MyCommon.GetAdminUser.Culture) & """ />")
          Else
            Dim tmp As Decimal = MyCommon.NZ(pgValue, 0)
            Send(currencySymbol & tmp.ToString(MyCommon.GetAdminUser.Culture) & "<br />")
            Send("<input type=""hidden"" id=""value"" name=""value"" value=""" & tmp.ToString(MyCommon.GetAdminUser.Culture) & """ />")
          End If
        End If
        Send("<br class=""half"" />")
        If (bCmInstalled) Then
          Send("<div>")
          Send("  <input type=""checkbox"" name=""allowreissue"" id=""allowreissue"" value=""1""" & IIf(AllowReissue, " checked=""checked""", "") & " />")
          Send("  <label for=""allowreissue"">" & Copient.PhraseLib.Lookup("term.allowreissue", LanguageID) & "</label><br />")
          If (MyCommon.Fetch_CPE_SystemOption(49) = "1") Then
            Send("<input type=""checkbox"" id=""storeInUnits"" name=""storeInUnits"" value=""1""" & IIf(StoreInUnits, " checked=""checked""", "") & " />")
            Send("<label for=""storeInUnits"">" & Copient.PhraseLib.Lookup("sv-storenote", LanguageID) & "</label><br />")
          End If
          If bAllowFuelPartner And SVTypeID = 3 Then
            Send("<input type=""checkbox"" id=""fuelpartner"" name=""fuelpartner"" value=""1"" onclick=""handleFuelPartner();""" & IIf(bFuelPartner, " checked=""checked""", "") & " /><label for=""fuelpartner"">" & Copient.PhraseLib.Lookup("term.fuelpartnerprogram", LanguageID) & "</label><br />")
            If bFuelPartner Then
              Send("<input type=""checkbox"" id=""autoredeem"" name=""autoredeem"" value=""1""" & IIf(bAutoRedeem, " checked=""checked""", "") & " /><label for=""autoredeem"">" & Copient.PhraseLib.Lookup("term.autoredeem", LanguageID) & "</label><br />")
              Send("<input type=""checkbox"" id=""allowadjust"" name=""allowadjust"" value=""1""" & IIf(bAllowAdjustments, " checked=""checked""", "") & " /><label for=""allowadjust"">" & Copient.PhraseLib.Lookup("term.allowadjust", LanguageID) & "</label><br />")
            Else
              Send("<input type=""checkbox"" id=""autoredeem"" name=""autoredeem"" value=""1"" disabled=""disabled""" & IIf(bAutoRedeem, " checked=""checked""", "") & " /><label for=""autoredeem"">" & Copient.PhraseLib.Lookup("term.autoredeem", LanguageID) & "</label><br />")
              Send("<input type=""checkbox"" id=""allowadjust"" name=""allowadjust"" value=""1"" disabled=""disabled""" & IIf(bAllowAdjustments, " checked=""checked""", "") & " /><label for=""allowadjust"">" & Copient.PhraseLib.Lookup("term.allowadjust", LanguageID) & "</label><br />")
            End If
          End If
          Send("</div>")
        End If
      %>
    </div>
    <div class="box" id="scorecards" <% Sendb(IIf(MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) OrElse MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CAM), "", " style=""display:none;""")) %>>
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.scorecard", LanguageID))%>
        </span>
      </h2>
      <table>
        <%
          If RewardsSetScorecard Then
            Send("<tr>")
            Send("  <td colspan=""2"">")
            Send("    <small>" & Copient.PhraseLib.Lookup("sv-edit.RewardsSetScorecard", LanguageID) & "</small>")
            Send("  </td>")
            Send("</tr>")
          End If
          Send("<tr>")
          Send("  <td style=""width:82px;"">")
          Send("    <label for=""ScorecardID"">" & Copient.PhraseLib.Lookup("term.scorecard", LanguageID) & ":</label>")
          Send("  </td>")
          Send("  <td>")
          If RewardsSetScorecard Then
            Send("    <input type=""hidden"" id=""ScorecardID"" name=""ScorecardID"" value=""" & ScorecardID & """ />")
            Send("    <select class=""medium"" id=""ScorecardIDLocked"" name=""ScorecardIDLocked"" onchange=""toggleScorecardText();"" disabled=""disabled"">")
            MyCommon.QueryStr = "select ScorecardID, Description from Scorecards with (NoLock) where ScorecardTypeID=2 and Deleted=0 and EngineID=2;"
            rst2 = MyCommon.LRT_Select
            Send("      <option value=""0"">" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "</option>")
            If rst2.Rows.Count > 0 Then
              For Each row In rst2.Rows
                Send("      <option value=""" & MyCommon.NZ(row.Item("ScorecardID"), 0) & """" & IIf(MyCommon.NZ(row.Item("ScorecardID"), 0) = ScorecardID, " selected=""selected""", "") & ">" & MyCommon.NZ(row.Item("Description"), "") & "</option>")
              Next
            End If
            Send("    </select>")
          Else
            Send("    <select class=""medium"" id=""ScorecardID"" name=""ScorecardID"" onchange=""toggleScorecardText();""" & IIf(RewardsSetScorecard, " disabled=""disabled""", "") & ">")
            MyCommon.QueryStr = "select ScorecardID, Description from Scorecards with (NoLock) where ScorecardTypeID=2 and Deleted=0 and EngineID=2;"
            rst2 = MyCommon.LRT_Select
            Send("      <option value=""0"">" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "</option>")
            If rst2.Rows.Count > 0 Then
              For Each row In rst2.Rows
                Send("      <option value=""" & MyCommon.NZ(row.Item("ScorecardID"), 0) & """" & IIf(MyCommon.NZ(row.Item("ScorecardID"), 0) = ScorecardID, " selected=""selected""", "") & ">" & MyCommon.NZ(row.Item("Description"), "") & "</option>")
              Next
            End If
            Send("    </select>")
          End If
          Send("  </td>")
          Send("</tr>")
          Send("<tr id=""scdesc""" & IIf(ScorecardID = 0, " style=""display:none;""", "") & ">")
          Send("  <td>")
          Send("    <label for=""ScorecardDesc"">" & Copient.PhraseLib.Lookup("term.scorecardtext", LanguageID) & ":</label>")
          Send("  </td>")
          Send("  <td>")
          'Send("    <input type=""text"" class=""medium"" id=""ScorecardDesc"" name=""ScorecardDesc"" maxlength=""31"" value=""" & ScorecardDesc & """" & IIf(RewardsSetScorecard, " disabled=""disabled""", "") & " />")
          MLI.ItemID = l_pgID
          MLI.MLTableName = "SVProgramTranslations"
          MLI.MLColumnName = "ScorecardDesc"
          MLI.MLIdentifierName = "SVProgramID"
          MLI.StandardTableName = "StoredValuePrograms"
          MLI.StandardColumnName = "ScorecardDesc"
          MLI.StandardIdentifierName = "SVProgramID"
          MLI.StandardValue = ScorecardDesc
          MLI.InputName = "ScorecardDesc"
          MLI.InputID = "ScorecardDesc"
          MLI.InputType = "text"
          MLI.MaxLength = 31
          MLI.CSSStyle = "width:230px;"
          MLI.Disabled = IIf(RewardsSetScorecard, True, False)
          Send(Localization.SendTranslationInputs(MyCommon, MLI, Request.QueryString, 2))
          Send("  </td>")
          Send("</tr>")
        %>
        <!-- Commenting out bolding functionality; to restore it, uncomment this section
          <tr>
            <%
              Send("<td><label for=""ScorecardBold"">" & Copient.PhraseLib.Lookup("term.bold", LanguageID) & ":</label></td>")
              Send("<td><input type=""checkbox"" id=""ScorecardBold"" name=""ScorecardBold"" value=""1""" & IIf(ScorecardBold, " checked=""checked""", "") & " /></td>")
            %>
          </tr>
          -->
      </table>
    </div>
    <% If MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE) Then%>
    <div class="box" id="advancedOptions">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.advancedoptions", LanguageID))%>
        </span>
      </h2>
     <%-- As part of UE offline implementation, we need to show advance options for UE also.
     <%
        Dim UEAlone As Boolean = False
        If MyCommon.IsEngineInstalled(0) = False AndAlso MyCommon.IsEngineInstalled(2) = False Then
          UEAlone = True
        End If
      %>
      <div <%Sendb(IIf(UEAlone, "style=""display:none;""", ""))%>>--%>
      <div>
        <label for="returnsHandling">
          <%Send(Copient.PhraseLib.Lookup("programs.returnhandling", LanguageID))%>:</label><br />
        <select id="returnsHandling" name="returnsHandling" class="longest">
          <%
            MyCommon.QueryStr = "select ReturnHandlingTypeID, Name, PhraseID from UE_ReturnHandlingTypes with (NoLock);"
            rst2 = MyCommon.LRT_Select
            For Each row2 In rst2.Rows
              Send("<option value=""" & MyCommon.NZ(row2.Item("ReturnHandlingTypeID"), 0) & """" & IIf(ReturnHandlingTypeID = MyCommon.NZ(row2.Item("ReturnHandlingTypeID"), 0), " selected=""selected""", "") & ">" & _
                    "" & Copient.PhraseLib.Lookup(MyCommon.NZ(row2.Item("PhraseID"), 0), LanguageID, MyCommon.NZ(row2.Item("Name"), "Unknown")))
            Next
          %>
        </select><br />
        <br class="half" />
      </div>
      <input type="checkbox" name="disallowRedeemInTrans" id="disallowRedeemInTrans" value="1"
        <%Send(IIf(DisallowRedeemInTrans, "checked=""checked"" ", "")) %> />
      <label for="disallowRedeemInTrans">
        <%Send(Copient.PhraseLib.Lookup("programs.disallowredeemintrans", LanguageID))%></label><br />
      <div>
        <input type="checkbox" name="allowNegBal" id="allowNegBal" value="1" <%Send(IIf(AllowNegativeBal, "checked=""checked"" ", "")) %> />
        <label for="allowNegBal">
          <%Send(Copient.PhraseLib.Lookup("programs.allowtogonegative", LanguageID))%></label><br />
      </div>
       
        
      <br class="half" />
      <hr class="hidden" />
    </div>
    <% End If%>
    <div class="box" id="ManualAdjustmentUPC" <% If (Not MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE)) OrElse (MyCommon.Fetch_SystemOption(94) = 0) OrElse (SVTypeID = 2 OrElse SVTypeID = 4) Then Send(" style=""display:none;""") %>>
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.adjustmentupc", LanguageID))%>
        </span>
      </h2>
      <%
        If Not (RangeBegin = 0 AndAlso RangeEnd = 0 AndAlso MyCommon.Fetch_CPE_SystemOption(102) = 0) Then
          Sendb("<input type=""text"" id=""AdjustmentUPC"" name=""AdjustmentUPC"" style=""width:200px;""")
          If AllDigits(AdjustmentUPC) Then
            If AdjustmentUPC = "" OrElse CDec(AdjustmentUPC) = 0 Then
              Send(" value="""" maxlength=""" & MaxLength & """ /><br />")
            Else
              Send(" value=""" & AdjustmentUPC & """ maxlength=""" & MaxLength & """ /><br />")
            End If
          Else
            Send(" value="""" maxlength=""" & MaxLength & """ /><br />")
          End If
        End If
        If Not (RangeBegin = 0 AndAlso RangeEnd = 0) Then
          If RangeBegin <> RangeEnd Then
            If RangeBegin > RangeEnd Then
              Sendb(Copient.PhraseLib.Lookup("sv-edit.RangeViolation", LanguageID))
            Else
              Sendb(Copient.PhraseLib.Detokenize("sv-edit.RangeBeginEnd", LanguageID, RangeBeginString, RangeEndString))
            End If
          Else
            Sendb(Copient.PhraseLib.Detokenize("sv-edit.RangeBegin", LanguageID, RangeBeginString))
          End If
        Else
          Sendb(Copient.PhraseLib.Lookup("ueoffer-con-plu.NoRange", LanguageID))
        End If
        If MyCommon.Fetch_CPE_SystemOption(102) Then
          Sendb(" " & Copient.PhraseLib.Lookup("ueoffer-con-plu.OutOfRangeAccepted", LanguageID))
        Else
          Sendb(" " & Copient.PhraseLib.Lookup("ueoffer-con-plu.OutOfRangeNotAccepted", LanguageID))
        End If
        Send("<br />")
        Send("<br class=""half"" />")
        Send("<hr />")
        'Selector
        If Not (RangeBegin = 0 AndAlso RangeEnd = 0 AndAlso MyCommon.Fetch_CPE_SystemOption(102) = 0) Then
          Send("<label for=""selector"">" & Copient.PhraseLib.Lookup("ueoffer-con-plu.TopUnusedCodes", LanguageID) & "</label><br />")
          Send("<select id=""selector"" name=""selector"" size=""10"" style=""width:220px;"" ondblclick=""javascript:selectPLU(" & IDLength & ");"">")
          x = RangeBegin
          counter = 1
          While (counter <= 100) AndAlso (x <= RangeEnd)
            MyCommon.QueryStr = "select CAST(AdjustmentUPC as decimal(26,0)) as AdjustmentUPC from StoredValuePrograms with (NoLock) " & _
                                " where IsNull(AdjustmentUPC, '') <> '' and AdjustmentUPC='" & x.ToString.PadLeft(IDLength, "0") & "' " & _
                                " union " & _
                                "select CAST(AdjustmentUPC as decimal(26,0)) as AdjustmentUPC from PointsPrograms with (NoLock) " & _
                                " where IsNull(AdjustmentUPC, '') <> '' and AdjustmentUPC='" & x.ToString.PadLeft(IDLength, "0") & "' "
            rst = MyCommon.LRT_Select
            If rst.Rows.Count = 0 Then
              Send("  <option value=""" & x & """>" & x.ToString.PadLeft(IDLength, "0") & "</option>")
              counter += 1
            End If
            x += 1
          End While
          Send("</select>")
        End If
      %>
    </div>
  </div>
  <div id="gutter">
  </div>
  <div id="column2">
    <div class="box" id="expire">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.expiration", LanguageID))%>
        </span>
      </h2>
      <label for="expiretype">
        <% Sendb(Copient.PhraseLib.Lookup("storedvalue.expiretype", LanguageID))%>:</label>
      <select name="expiretype" id="expiretype" class="long" onchange="handleExpireType();">
        <%
          MyCommon.QueryStr = "select SVExpireTypeID, Description, PhraseID from SVExpireTypes with (NoLock);"
          rst = MyCommon.LRT_Select
          For Each row In rst.Rows
            selectedStr = IIf(MyCommon.NZ(row.Item("SVExpireTypeID"), 0) = svExpireType, " selected=""selected""", "")
            Send("<option value=""" & MyCommon.NZ(row.Item("SVExpireTypeID"), 0) & """" & selectedStr & ">" & Copient.PhraseLib.Lookup(CInt(MyCommon.NZ(row.Item("PhraseID"), 0)), LanguageID, MyCommon.NZ(row.Item("Description"), "")) & "</option>")
          Next
        %>
      </select>
      <br />
      <br class="half" />
      <div id="trExpPdType">
        <label for="expireperiodtype">
          <% Sendb(Copient.PhraseLib.Lookup("storedvalue.expireperiodtype", LanguageID))%>:</label>
        <select name="expireperiodtype" id="expireperiodtype" class="short">
          <%
            MyCommon.QueryStr = "select SVExpirePeriodTypeID, Description, PhraseID from SVExpirePeriodTypes with (NoLock);"
            rst = MyCommon.LRT_Select
            For Each row In rst.Rows
              If (MyCommon.NZ(row.Item("SVExpirePeriodTypeID"), 0) <> 0) Then
                selectedStr = IIf(MyCommon.NZ(row.Item("SVExpirePeriodTypeID"), 0) = svExpirePeriodType, " selected=""selected""", "")
                Send("<option value=""" & MyCommon.NZ(row.Item("SVExpirePeriodTypeID"), 0) & """" & selectedStr & ">" & Copient.PhraseLib.Lookup(CInt(MyCommon.NZ(row.Item("PhraseID"), 0)), LanguageID, MyCommon.NZ(row.Item("Description"), "")) & "</option>")
              End If
            Next
          %>
        </select>
      </div>
      <div id="trExpPd">
        <label for="expireperiod">
          <% Sendb(Copient.PhraseLib.Lookup("storedvalue.expireperiod", LanguageID))%>:</label>
        <input type="text" class="short" id="expireperiod" name="expireperiod" maxlength="5"
          value="<% Sendb(pgExpire) %>" />
      </div>
      <div id="trExpTod">
        <label for="expTodHours">
          <% Sendb(Copient.PhraseLib.Lookup("storedvalue.expiretod", LanguageID))%>:</label>
        <input type="hidden" name="expiretod" id="expiretod" value="<% Sendb(svExpireTOD) %>" />
        <select id="expTodHours" name="expTodHours">
          <%
            For i = 0 To 23
              selectedStr = IIf(i = Integer.Parse(svExpireTODHr), " selected=""selected""", "")
              Send("<option value=""" & i.ToString("00") & """" & selectedStr & ">" & i.ToString("00") & "</option>")
            Next
          %>
        </select>:<select id="expTodMinutes" name="expTodMinutes">
          <%
            For i = 0 To 59
              selectedStr = IIf(i = Integer.Parse(svExpireTODMin), " selected=""selected""", "")
              Send("<option value=""" & i.ToString("00") & """" & selectedStr & ">" & i.ToString("00") & "</option>")
            Next
          %>
        </select>
      </div>
      <div id="trExpTime">
        <label for="expHours">
          <% Sendb(Copient.PhraseLib.Lookup("storedvalue.expiretime", LanguageID))%>:</label>
        <select id="expHours" name="expHours" <%Send(IIf(bProgramActive, "onchange=""handleExpireDateTimeChange()"" ", "")) %> >
          <%
            For i = 0 To 23
              selectedStr = IIf(i = svExpireHr, " selected=""selected""", "")
              Send("<option value=""" & i.ToString("00") & """" & selectedStr & ">" & i.ToString("00") & "</option>")
            Next
          %>
        </select>:<select id="expMinutes" name="expMinutes" <%Send(IIf(bProgramActive, "onchange=""handleExpireDateTimeChange()"" ", "")) %> >
          <%
            For i = 0 To 59
              selectedStr = IIf(i = svExpireMin, " selected=""selected""", "")
              Send("<option value=""" & i.ToString("00") & """" & selectedStr & ">" & i.ToString("00") & "</option>")
            Next
          %>
        </select>
      </div>
      <div id="trExpDate">
        <label for="expiredate">
          <% Sendb(Copient.PhraseLib.Lookup("storedvalue.expiredate", LanguageID))%>:</label>
        <input type="hidden" id="expiredatetime" name="expiredatetime" value="<% Sendb(svExpireDateStr)%>" />
        <input type="text" class="short" name="expiredate" id="expiredate" <%Send(IIf(bProgramActive, "onchange=""handleExpireDateTimeChange()"" ", "")) %> value="<% Sendb(svExpireDateStr)%>" />
        <img src="../images/calendar.png" class="calendar" id="expire-date-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
          title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('expiredate', event);" />
      </div>
      <% If MyCommon.Fetch_CM_SystemOption(87) = "1" Then%>
      <div id="trExactExp">
        <input type="checkbox" name="expirecentralserverTZ" id="expirecentralserverTZ" <%Send(IIf(ExpireCentralServerTZ, "checked=""checked"" ", "")) %> />
        <label for="expirecentralserverTZ">
          <%Send(Copient.PhraseLib.Lookup("storedvalue.expirecentralserverTZ", LanguageID))%></label><br />
      </div>
      <% End If%>
    </div>
    <div id="datepicker" class="dpDiv">
    </div>
    <%
      If Request.Browser.Type = "IE6" Then
        Send("<iframe src=""javascript:'';"" id=""calendariframe"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no"" style=""visibility:hidden;display:none;""></iframe>")
      End If
    %>
    <%
      If CPEInstalled Then
        Send("<div class=""box"" id=""redemption"">")
        Send("  <h2>")
        Send("    <span>")
        Send("      " & Copient.PhraseLib.Lookup("term.redemption", LanguageID))
        Send("    </span>")
        Send("  </h2>")
        Send("  <label for=""redemptionRestrictionID"">" & Copient.PhraseLib.Lookup("term.redemptionrestriction", LanguageID) & ":</label>")
        Send("  <select name=""redemptionRestrictionID"" id=""redemptionRestrictionID"">")
        MyCommon.QueryStr = "select RedemptionRestrictionID, Description, PhraseID from RedemptionRestrictions with (NoLock);"
        rst = MyCommon.LRT_Select
        For Each row In rst.Rows
          Sendb("    <option value=""" & MyCommon.NZ(row.Item("RedemptionRestrictionID"), 0) & """" & IIf(MyCommon.NZ(row.Item("RedemptionRestrictionID"), 0) = RedemptionRestrictionID, " selected=""selected""", "") & ">")
          If MyCommon.NZ(row.Item("PhraseID"), 0) > 0 Then
            Sendb(Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID))
          Else
            Sendb(MyCommon.NZ(row.Item("Description"), ""))
          End If
          Send("</option>")
        Next
        Send("  </select>")
        Send("<br />")
        Send("  <label for=""MemberRedemptionID"">" & Copient.PhraseLib.Lookup("term.memberredemptionrestriction", LanguageID) & ":</label>")
        Send("  <select name=""memberredemptionID"" id=""memberredemptionID"">")
        MyCommon.QueryStr = "select MemberRedemptionID, Description, PhraseID from MemberRedemptionRestrictions with (NoLock);"
        rst = MyCommon.LRT_Select
        For Each row In rst.Rows
          Sendb("    <option value=""" & MyCommon.NZ(row.Item("MemberRedemptionID"), 0) & """" & IIf(MyCommon.NZ(row.Item("MemberRedemptionID"), 0) = MemberRedemptionID, " selected=""selected""", "") & ">")
          If MyCommon.NZ(row.Item("PhraseID"), 0) > 0 Then
            Sendb(Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID))
          Else
            Sendb(MyCommon.NZ(row.Item("Description"), ""))
          End If
          Send("</option>")
        Next
        Send("  </select>")
        Send("  <hr class=""hidden"" />")
        Send("</div>")
      Else
        Send("<input type=""hidden"" id=""redemptionRestrictionID"" name=""redemptionRestrictionID"" value=""" & RedemptionRestrictionID & """ />")
      End If
    %>
    <% If MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE) Then%>
    <div class="box" id="EngineSpecificOptions">
      <h2 style="float: left;">
        <span>
          <%Send(Copient.PhraseLib.Lookup("term.enginespecificsettings", LanguageID))%>
        </span>
      </h2>
      <% Send_BoxResizer("EngineSpecificOptionsBody", "imgEngineSpecificOptionsBody", "Engine-Specific Settings", True)%>
      <div id="EngineSpecificOptionsBody">
        <!--Below Code to be used when any Engine Specific Options for CM/CPE are available-->
        <%--<% If MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM) Then%>
      <label style="font-size:small; font-weight:bold;">CM Settings:</label><br />
      <% End If%>
      <% If MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) Then%>
        <label style="font-size:small; font-weight:bold;">CPE Settings:</label><br />
      <% End If%>--%>
        <% If MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE) Then%>
        <h3><%Send(Copient.PhraseLib.Lookup("term.uesettings", LanguageID))%></h3>
        <input type="hidden" id="hdnAllowAnyCustomerUE" name="hdnAllowAnyCustomerUE" />
        <input type='checkbox' name='cbAllowAnyCustomerUE' id='cbAllowAnyCustomerUE' value='1'
          onchange='javascript:updateInputVal(this, "hdnAllowAnyCustomerUE");' <% if(AllowAnyCustomer_UE) Then sendb(" checked=""checked""") %>
          <% if(IsAnyCustomerOffersExist_UE) Then sendb(" disabled=""disabled""") %> />
        <label for="cbAllowAnyCustomerUE">
          <%Send(Copient.PhraseLib.Lookup("term.allowanycustomerearnredeem", LanguageID))%></label>
        <% End If%><br />
         
        <input type="checkbox" name="VisibleToCustomers" id="VisibleToCustomers" value="1"
        <%Send(IIf(VisibleToCustomers, "checked=""checked"" ", ""))%> />
      <label for="VisibleToCustomers">
        <%Send(Copient.PhraseLib.Lookup("programs.includeforcustomerportal", LanguageID))%></label><br />
          
        <br class="half" />
        <hr class="hidden" />
      </div>
    </div>
    <% End If%>
    <% If (l_pgID > 0) Then%>
    <div class="box" id="eligibleoffers" <% if(l_pgID=0)then sendb(" style=""visibility: hidden;""") %>>
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.associatedeligibleoffers", LanguageID))%>
        </span>
      </h2>
      <div class="boxscrollhalf">
        <% 
          If (l_pgID <> 0) Then
            Dim lstOffers As List(Of CMS.AMS.Models.Offer)
            lstOffers = m_Offer.GetEligibleOffersBySVProgramID(l_pgID)
            If(bEnableRestrictedAccessToUEOfferBuilder) Then
                lstOffers = GetRoleBasedUEOffers(lstOffers,MyCommon,Logix)
            End If
            If lstOffers.Count > 0 Then
              For Each offer As CMS.AMS.Models.Offer In lstOffers
                If (MyCommon.Fetch_UE_SystemOption(168) = "1" AndAlso MyCommon.NZ(offer.BuyerID, "") <> "") Then
                offer.OfferName = "Buyer " + offer.BuyerID.ToString() + " -" + MyCommon.NZ(offer.OfferName, "").ToString()
                Else
                offer.OfferName = MyCommon.NZ(offer.OfferName,Copient.PhraseLib.Lookup("term.unknown", LanguageID))
                End If
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
    <% End If%>
    <% If (l_pgID > 0) Then%>
    <div class="box" id="offers">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.associatedoffers", LanguageID))%>
        </span>
      </h2>
      <hr class="hidden" />
      <div class="boxscroll">
        <% 
          
          If Not dstAssociated Is Nothing AndAlso dstAssociated.Rows.Count > 0 Then
            For Each row In dstAssociated.Rows
                 If (MyCommon.Fetch_UE_SystemOption(168) = "1" AndAlso MyCommon.NZ(row.Item("BuyerID"), "") <> "") Then
                assocName = "Buyer " + row.Item("BuyerID").ToString() + " - " + MyCommon.NZ(row.Item("Name"), "").ToString()
                Else
                assocName = MyCommon.NZ(row.Item("Name"),Copient.PhraseLib.Lookup("term.unknown", LanguageID))
                End If
              'assocName = row.Item("Name")
              assocID = row.Item("OfferID")
              If (Logix.IsAccessibleOffer(AdminUserID, row.Item("OfferID"))) Then
                Sendb(" <a href=""offer-redirect.aspx?OfferID=" & assocID & """>" & assocName & "</a><br />")
              Else
                Sendb(assocName & "<br />")
              End If
            Next
          Else
            Send("  " & Copient.PhraseLib.Lookup("term.none", LanguageID))
          End If
        %>
      </div>
      <hr class="hidden" />
    </div>
    <% End If%>
    <%If (MyCommon.Fetch_CM_SystemOption(70) = "1") AndAlso (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) AndAlso _
           (SVTypeID = 1 OrElse l_pgID = 0) Then%>
    <div class="box" id="thirdpartyconversion">
      <h2>
        <span>
          <% Sendb("Third Party Conversion")%>
        </span>
      </h2>
      <%
        Dim SVPoints As String = ""
        Dim SVMonetary As String = ""
        MyCommon.QueryStr = "select SVMonetaryValue, SVPointsValue from StoredValuePointsConversion with (NoLock) where SVProgramID = " & l_pgID & ";"
        rst2 = MyCommon.LXS_Select
        If rst2.Rows.Count > 0 Then
          SVPoints = rst2.Rows(0).Item("SVPointsValue")
          SVMonetary = rst2.Rows(0).Item("SVMonetaryValue")
        End If
        Send("<br class=""half"" />")
        Send("<label for=""RequiredPoints"" style=""position:relative;"">Required Points: </label>")
        Send("<input type=""text"" class=""shorter"" id=""RequiredPoints"" name=""RequiredPoints"" maxlength=""20"" value=""" & SVPoints & """ /><br />")
        Send("<br class=""half"" />")
        Send("<label for=""MonetaryValue"" style=""position:relative;"">Monetary Value: </label>")
        Send("<input type=""text"" class=""shorter"" id=""MonetaryValue"" name=""MonetaryValue"" maxlength=""20"" value=""" & SVMonetary & """ /><br />")
        Send("</div>")
      %>
    </div>
    <% End If%>
    <br clear="all" />
  </div>
</div>
</form>
<script runat="server">
    Private Function ValidateMonthExpiry(monthValue As String, expirePeriodType As String) As Boolean
        If (expirePeriodType = 3 And monthValue > 120) Then
            Return False
        End If
        Return True
    End Function
  Public Function AllDigits(ByVal txt As String) As Boolean
    Dim ch As String
    Dim i As Integer
    
    AllDigits = True
    If Len(txt) > 0 Then
      For i = 1 To Len(txt)
        ' See if the next character is a non-digit.
        ch = Mid$(txt, i, 1)
        If ch < "0" Or ch > "9" Then
          ' This is not a digit.
          AllDigits = False
          Exit For
        End If
      Next i
    Else
      AllDigits = False
    End If
  End Function
</script>
<script type="text/javascript">
<% Send_Date_Picker_Terms() %>
  handleExpireType();
  
  if (window.captureEvents){
    window.captureEvents(Event.CLICK);
    window.onclick=handlePageClick;
  } else {
    document.onclick=handlePageClick;
  }
</script>
<%
    If MyCommon.Fetch_SystemOption(75) Then
        If (l_pgID > 0 And Logix.UserRoles.AccessNotes) Then
            Send_Notes(9, l_pgID, AdminUserID)
        End If
    End If
    Send_WrapEnd()
    Send_PageEnd()
done:
    MyCommon.Close_LogixRT()
    MyCommon.Close_LogixXS()
    Logix = Nothing
    MyCommon = Nothing
%>
