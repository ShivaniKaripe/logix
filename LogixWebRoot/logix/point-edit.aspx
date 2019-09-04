<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="CMS.AMS" %>
<%@ Import Namespace="CMS.AMS.Contract" %>
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
      if (target.href && target.className != "calendar") {
        if (IsFormChanged(document.mainform)) {
          var bConfirm = confirm('Warning:  Unsaved data will be lost.  Are you sure you wish to continue?');
          return bConfirm;
        }
      }
    };
  }

  function PageClick(evt) {
    var target = document.all ? event.srcElement : evt.target;

    if (target.href && target.className != "calendar") {
      if (IsFormChanged(document.mainform)) {
        var bConfirm = confirm('Warning:  Unsaved data will be lost.  Are you sure you wish to continue?');
        return bConfirm;
      }
    }
  }
</script>
<%
    ' *****************************************************************************
    ' * FILENAME: point-edit.aspx 
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
    Dim MyCryptLib As New Copient.CryptLib
    Dim Logix As New Copient.LogixInc
    Dim Localization As Copient.Localization
    Dim dstPoints As System.Data.DataTable
    Dim dstPrograms As System.Data.DataTable
    Dim dstAssociated As System.Data.DataTable
    Dim rst As DataTable
    Dim rst2 As DataTable
    'Dim rstPoints As DataTable
    Dim row As System.Data.DataRow
    Dim row2 As System.Data.DataRow
    Dim pgDescription As String = String.Empty
    Dim pgPromoVarID As String = String.Empty
    Dim pgTotalPoints As Long
    Dim pgCreated As String = String.Empty
    Dim pgUpdated As String = String.Empty
    Dim pgName As String = String.Empty
    Dim ProgramDescription As String = String.Empty
    Dim ProgramName As String = String.Empty
    Dim PromoVarID As String = String.Empty
    Dim ProgramID As Long
    Dim rowCount As Integer
    Dim assocName As String = String.Empty
    Dim assocID As String = String.Empty
    Dim l_pgID As String = String.Empty
    Dim IsExternalProgram As Boolean
    Dim ExternalID As String = ""
    Dim PartnerCode As String = ""
    Dim PartnerID As String = ""
    Dim ExtHostFuelProgram As Boolean = False
    Dim ExtHostCardBINMin As Long = 0 
    Dim ExtHostCardBINMax As Long = 0 
    Dim ExtHostTypeID As Integer = 0
    Dim ExtHostDesc As String = ""
    Dim DecimalValues As Boolean = False
    Dim longDate As New DateTime
    Dim longDateString As String = String.Empty
    Dim shortDateString As String = String.Empty
    Dim CpeEngineOnly As Boolean = False
    Dim CAMInstalled As Boolean = False
    Dim LastUpload As String
    Dim LastUploadMsg As String = ""
    Dim File As HttpPostedFile
    'Dim OptionText As String
    Dim CustExtID As String = String.Empty
    Dim BalanceAmt As Double
    Dim Status As Integer
    Dim CustBalPK As Integer
    Dim ScorecardID As Integer = 0
    Dim ScorecardDesc As String = ""
    Dim ScorecardBold As Boolean = False
    'Dim ScorecardLevel As Integer = 0
    Dim ProgramNameTitle As String = ""
    Dim AdjustmentUPC As String = ""
    Dim statusMessage As String = ""
    Dim ShowActionButton As Boolean = False
    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False
    'Dim OfferCtr As Integer = 0
    Dim AutoDelete As Boolean = False
    Dim RewardsSetScorecard As Boolean = False
    Dim CAMProgram As Boolean = False
    Dim ValidSetUPC As Boolean = False
    Dim ValidSingleUPC As Boolean = False
    Dim BeginUPC As Decimal = 0
    Dim EndUPC As Decimal = 0
    Dim UPCDec As Decimal = 0
    Dim IDLength As Integer = 0

    Dim RangeBegin As Decimal = 0
    Dim RangeBeginString As String = ""
    Dim RangeEnd As Decimal = 0
    Dim RangeEndString As String = ""
    Dim Range As Decimal = 0
    Dim i As Decimal = 0
    Dim counter As Integer = 1
    Dim MaxLength As Integer = 0
    Dim bCMInstalled As Boolean = False
    Dim bEnableCategories As Boolean = False
    Dim iCategoryID As Integer
    Dim XID As String = ""

    Dim sr As System.IO.StreamReader
    Dim UploadFileName As String

    Dim UEInstalled As Boolean = False
    Dim ReturnHandlingTypeID As Integer = 1
    Dim DisallowRedeemInTrans As Boolean = False
    Dim VisibleToCustomers As Boolean = False
    Dim AllowNegativeBal As Boolean = False
    Dim bPendingEnabled as Boolean = False
    Dim bDefaultPendingIsUseRedeem as Boolean = True
    Dim bApplyEarnedPending as Boolean = True
    Dim MultiLanguageEnabled As Boolean = False
    Dim DefaultLanguageID As Integer = 0
    Dim DefaultLanguageCode As String = ""
    Dim MLI As New Copient.Localization.MultiLanguageRec
    Dim AllowAnyCustomer_UE As Boolean = False
    Dim AllowAnyCustomer_UE_PKID As Long = 0
    Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249) = "1", True, False)
    Dim conditionalQuery = String.Empty
    CurrentRequest.Resolver.AppName = "point-edit.aspx"
    Dim m_Offer As IOffer = CurrentRequest.Resolver.Resolve(Of IOffer)()

    Dim bPointsMigrationEnabled As Boolean = False
    Dim lMigrationProgramID As Long = 0
    Dim MigrationDatetTime As String = String.Empty
    Dim MigrationDate As String = String.Empty
    Dim MigrationHr As String = String.Empty
    Dim MigrationMin As String = String.Empty
    Dim sDateOnlyFormat As String = "MM/dd/yyyy"
    Dim sHourOnlyFormat As String = "HH"
    Dim sMinutesOnlyFormat As String = "mm"
    Dim tempDateTime As Date
    Dim LastMigrationDate As Date = Nothing
    Dim bValidMigrationData As Boolean = True
    Dim IsAnyCustomerOffersExist_UE As Boolean = False
    Dim lMinimumAutoGeneratedPromoVarID As Long = 0


    Dim bPointsDeletionEnabled As Boolean = False
    Dim DeletionDateTime As String
    Dim DeletionDate As String
    Dim DeletionHr As String
    Dim DeletionMin As String
    Dim bValidDeletionDate As Boolean = True
    Dim LastDeletionDate As Date = Nothing
    Dim bEnableDeleteDate As Boolean = False

    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If

    Response.Expires = 0
    MyCommon.AppName = "point-edit.aspx"
    MyCommon.Open_LogixRT()
    MyCommon.Open_LogixXS()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
    Localization = New Copient.Localization(MyCommon)

    UEInstalled = MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE)
    bCMInstalled = MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)
    If bCMInstalled Then
        bEnableCategories = MyCommon.Fetch_CM_SystemOption(22) = "1"
    Else
        bEnableCategories = False
    End If

    bPointsMigrationEnabled = MyCommon.Fetch_SystemOption(164) = "1"
    bPointsDeletionEnabled = MyCommon.Fetch_SystemOption(205) = "1"

    If Not Long.TryParse(MyCommon.Fetch_CM_SystemOption(106), lMinimumAutoGeneratedPromoVarID) Then lMinimumAutoGeneratedPromoVarID = 0

    bPendingEnabled = MyCommon.Fetch_SystemOption(251)
    bDefaultPendingIsUseRedeem = MyCommon.Fetch_SystemOption(252)
    bApplyEarnedPending = bDefaultPendingIsUseRedeem

    MultiLanguageEnabled = MyCommon.Fetch_SystemOption(124) = "1"
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
            AdjustmentUPC = MyCommon.Parse_Quotes(Left(Trim(Request.QueryString("AdjustmentUPC")), 26)).PadLeft(IDLength, "0")
        Else
            AdjustmentUPC = MyCommon.Parse_Quotes(Left(Trim(Request.QueryString("AdjustmentUPC")), 26))
        End If
    End If

    If (Request.QueryString("infoMessage") <> "") Then
        infoMessage = Request.QueryString("infoMessage")
    End If

    If (Request.QueryString("new") <> "") Then
        Response.Redirect("point-edit.aspx")
    End If

    If Request.QueryString("LargeFile") = "true" Then
        infoMessage = Copient.PhraseLib.Lookup("error.UploadTooLarge", LanguageID)
    End If

    If (Request.QueryString("isexternal") = "1") Then
        IsExternalProgram = True
    End If

    If (Request.QueryString("externalid") <> "") Then
        ExternalID = Request.QueryString("externalid")
    End If
    If (Request.QueryString("partnercode") <> "") Then
      PartnerCode = Request.QueryString("partnercode")
    End If
    If (Request.QueryString("partnerid") <> "") Then
      PartnerID = Request.QueryString("partnerid")
    End If
    ExtHostFuelProgram = IIf(Request.QueryString("exthostfuelprogram") = "1", True, False)
    If (Request.QueryString("exthostcardbinmin") <> "") Then
       ExtHostCardBINMin = Request.QueryString("exthostcardbinmin")
    End If
    If (Request.QueryString("exthostcardbinmax") <> "") Then
       ExtHostCardBINMax = Request.QueryString("exthostcardbinmax")
    End If

    If (Request.QueryString("exthosttypeid") <> "") Then
        ExtHostTypeID = Request.QueryString("exthosttypeid")
    End If

    If (Request.QueryString("decimalvalues") = "1") Then
        DecimalValues = True
    End If

    CAMInstalled = MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CAM)
    If (Request.QueryString("CAMProgram") = "1") Then
        CAMProgram = True
    Else
        CAMProgram = False
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

    If Request.Files.Count >= 1 Then
        File = Request.Files.Get(0)

        ' AL-4783 Prevent errors when no file is selected
        If String.IsNullOrWhiteSpace(File.FileName) Then
            infoMessage = Copient.PhraseLib.Lookup("term.upload-file-not-found", LanguageID)
        Else

            l_pgID = MyCommon.Extract_Val(Request.Form("ProgramID"))
            UploadFileName = "PointsAdj." & l_pgID.ToString & "." & Now().ToString("yyyyMMddhhmmss") & ".ppp"
            Dim uploadFilePath As String = MyCommon.Fetch_SystemOption(29) & "\" & UploadFileName

            Try
                File.SaveAs(uploadFilePath)
                sr = New System.IO.StreamReader(File.InputStream)
                sr.Close()
                MyCommon.QueryStr = "dbo.pt_PointsInsertQueue_Insert"
                MyCommon.Open_LXSsp()
                MyCommon.LXSsp.Parameters.Add("@FileName", SqlDbType.NVarChar, 255).Value = UploadFileName
                MyCommon.LXSsp.Parameters.Add("@ProgramID", SqlDbType.BigInt).Value = l_pgID
                MyCommon.LXSsp.Parameters.Add("@PKID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                MyCommon.LXSsp.ExecuteNonQuery()
                MyCommon.Activity_Log(7, l_pgID, AdminUserID, Copient.PhraseLib.Lookup("history.points-upload", LanguageID))
                MyCommon.Close_LXSsp()
            Catch ex As Exception
                infoMessage = Copient.PhraseLib.Detokenize("error.encountered", LanguageID, ex.ToString)
            End Try

        End If
    End If

    If (Request.QueryString("download") <> "" AndAlso Request.QueryString("PromoVarID") <> "") Then
        MyCommon.QueryStr = "select MAX(C.CardPK) as CardPK, ExtCardID, C.CustomerPK, P.Amount from Points as P with (NoLock) " & _
                            "inner join CardIDs as C with (NoLock) on C.CustomerPK = P.CustomerPK " & _
                            "where P.PromoVarID = " & MyCommon.Extract_Val(Request.QueryString("PromoVarID")) & " and Amount <> 0 " & _
                            "group by ExtCardID, C.CustomerPK, P.Amount;"
        rst = MyCommon.LXS_Select()
        If (rst.Rows.Count > 0) Then
            Response.Clear()
            Response.AddHeader("Content-Disposition", "attachment; filename=PNT" & Request.QueryString("ProgramGroupID") & ".csv")
            Response.ContentType = "application/octet-stream"
            Send("CustomerID,Balance")
            For Each row In rst.Rows                
                Dim tmpExtCID As String = MyCryptLib.SQL_StringDecrypt(row.Item("ExtCardID").ToString())
                Send(IIf(tmpExtCID="", Copient.PhraseLib.Lookup("term.unknown", LanguageID), tmpExtCID) & "," & Int(MyCommon.NZ(row.Item("Amount"), 0)))
            Next
            GoTo done
        Else
            infoMessage = Copient.PhraseLib.Lookup("point-edit.noelements", LanguageID)
        End If
    End If

    ' any GET parms inbound?
    If (Request.QueryString("Delete") <> "") Then
        l_pgID = MyCommon.Extract_Val(Request.QueryString("ProgramGroupID"))
        MyCommon.QueryStr = "SELECT DISTINCT O.Name, O.OfferID " & _
                            "FROM PointsPrograms AS PP WITH (NoLock) " & _
                            "INNER JOIN OfferConditions AS OC WITH (NoLock) " & _
                            "ON PP.ProgramID = OC.LinkID AND OC.Deleted=0 " & _
                            "AND PP.Deleted = 0 " & _
                            "AND OC.ConditionTypeID = 3 " & _
                            "INNER JOIN Offers AS O WITH (NoLock) " & _
                            "ON OC.OfferID = O.OfferID " & _
                            "AND O.Deleted = 0 " & _
                            "WHERE PP.ProgramID = " & l_pgID & " " & _
                            "UNION " & _
                            "SELECT DISTINCT O.Name, O.OfferID " & _
                            "FROM RewardPoints AS RP WITH (NoLock) " & _
                            "INNER JOIN OfferRewards AS OFFR WITH (NoLock) " & _
                            "ON RP.RewardPointsID = OFFR.LinkID " & _
                            "AND (OFFR.RewardTypeID = 2 OR OFFR.RewardTypeID = 13) AND OFFR.Deleted=0 " & _
                            "INNER JOIN Offers AS O WITH (NoLock) " & _
                            "ON OFFR.OfferID = O.OfferID " & _
                            "AND O.Deleted = 0 " & _
                            "WHERE RP.ProgramID = " & l_pgID & " " & _
                            " UNION " & _
                            "SELECT DISTINCT I.IncentiveName as Name, I.IncentiveID as OfferID from CPE_IncentivePointsGroups IPG with (NoLock) " & _
                            "INNER JOIN CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID=IPG.RewardOptionID " & _
                            "INNER JOIN CPE_Incentives I with (NoLock) on I.IncentiveID = RO.IncentiveID " & _
                            "INNER JOIN PointsPrograms PP with (NoLock) on IPG.ProgramID = PP.ProgramID " & _
                            "WHERE IPG.ProgramID = " & l_pgID & " and IPG.Deleted = 0 and RO.Deleted=0 and I.Deleted=0 and PP.Deleted=0 " & _
                            "UNION " & _
                            "SELECT DISTINCT I.IncentiveName as Name, I.IncentiveID as OfferID from CPE_DeliverablePoints DP " & _
                            "INNER JOIN CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID=DP.RewardOptionID " & _
                            "INNER JOIN CPE_Incentives I with (NoLock) on I.IncentiveID = RO.IncentiveID " & _
                            "WHERE DP.Deleted=0 and RO.Deleted=0 and I.Deleted=0 and ProgramID=" & l_pgID & " " & _
                            "UNION " & _
                            "SELECT DISTINCT I.IncentiveName as Name, I.IncentiveID as OfferID from PointsCondition PC " & _
                            "INNER JOIN OfferEligibilityConditions OEC with (NoLock) ON OEC.ConditionID = PC.ConditionID AND OEC.Deleted = 0 " & _
                            "INNER JOIN CPE_Incentives I with (NoLock) on I.IncentiveID = OEC.OfferID " & _
                            "WHERE I.Deleted=0 and PC.ProgramID=" & l_pgID
        'Send(MyCommon.QueryStr)
        dstAssociated = MyCommon.LRT_Select

        ' AL-1388 Check if the program is not used anywhere and delete
        If dstAssociated.Rows.Count > 0 Then
            infoMessage = Copient.PhraseLib.Lookup("point-edit.inuse", LanguageID)
        Else
            MyCommon.QueryStr = "dbo.pt_PointsPrograms_Delete"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@ProgramID", SqlDbType.BigInt).Value = l_pgID
            MyCommon.LRTsp.ExecuteNonQuery()
            MyCommon.Close_LRTsp()
            MyCommon.Activity_Log(5, ProgramID, AdminUserID, Copient.PhraseLib.Lookup("history.point-delete", LanguageID))

            pgPromoVarID = MyCommon.Extract_Val(Request.QueryString("PromoVarID"))
            MyCommon.QueryStr = "dbo.pt_Promo_Variables_Delete"
            MyCommon.Open_LXSsp()
            MyCommon.LXSsp.Parameters.Add("@PromoVarID", SqlDbType.BigInt).Value = pgPromoVarID
            MyCommon.LXSsp.ExecuteNonQuery()
            MyCommon.Close_LXSsp()

            Response.Status = "301 Moved Permanently"
            Response.AddHeader("Location", "point-list.aspx")
        End If

    ElseIf (Request.QueryString("ProgramGroupID") = "new") Then
        ' add a record
        Dim bSuccess As Boolean = True

        ExternalID = MyCommon.Parse_Quotes(Request.QueryString.Item("externalid"))
        If (IsExternalProgram AndAlso ExternalID = "" AndAlso MyCommon.Fetch_SystemOption(223) = "1") Then
            infoMessage = Copient.PhraseLib.Lookup("error.noextid", LanguageID)
            bSuccess = False
        End If

        ProgramName = MyCommon.Parse_Quotes(Request.QueryString.Item("name"))
        ProgramName = Logix.TrimAll(ProgramName)
        If (ProgramName = "") Then
            infoMessage = Copient.PhraseLib.Lookup("point-edit.noname", LanguageID)
            bSuccess = False
        ElseIf (bSuccess) Then
            MyCommon.QueryStr = "select ProgramID from PointsPrograms with (NoLock) where ProgramName = '" & ProgramName & "' and Deleted=0;"
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count > 0) Then
                infoMessage = Copient.PhraseLib.Lookup("point-edit.nameused", LanguageID)
                bSuccess = False
            Else
                PromoVarID = 0
                If Not Long.TryParse(Request.QueryString("PromoVarID"), PromoVarID) Then PromoVarID = 0

                If PromoVarID > 0 Then
                    If PromoVarID >= lMinimumAutoGeneratedPromoVarID Then
                        infoMessage = String.Format(Copient.PhraseLib.Lookup("promovar.too-big", LanguageID), lMinimumAutoGeneratedPromoVarID)
                        bSuccess = False
                    End If
                End If

                If bSuccess Then
                    MyCommon.QueryStr = "dbo.pt_PointsPrograms_Insert"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@ProgramName", SqlDbType.NVarChar, 200).Value = Request.QueryString.Item("name")
                    MyCommon.LRTsp.Parameters.Add("@CAMProgram", SqlDbType.Bit).Value = IIf(CAMProgram, 1, 0)
                    MyCommon.LRTsp.Parameters.Add("@ExternalProgram", SqlDbType.Bit).Value = IIf(IsExternalProgram, 1, 0)
                    MyCommon.LRTsp.Parameters.Add("@AutoDelete", SqlDbType.Bit).Value = 1
                    MyCommon.LRTsp.Parameters.Add("@ProgramID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                    MyCommon.LRTsp.ExecuteNonQuery()
                    ProgramID = MyCommon.LRTsp.Parameters("@ProgramID").Value
                    MyCommon.Close_LRTsp()
                    MyCommon.Activity_Log(7, ProgramID, AdminUserID, Copient.PhraseLib.Lookup("history.point-create", LanguageID))

                    If PromoVarID > 0 Then
                        MyCommon.QueryStr = "dbo.pc_PointsVar_Create_Specific"
                        MyCommon.Open_LXSsp()
                        MyCommon.LXSsp.Parameters.Add("@ProgramID", SqlDbType.BigInt).Value = ProgramID
                        MyCommon.LXSsp.Parameters.Add("@VarID", SqlDbType.BigInt).Value = PromoVarID
                        MyCommon.LXSsp.Parameters.Add("@Success", SqlDbType.Bit).Direction = ParameterDirection.Output
                        MyCommon.LXSsp.ExecuteNonQuery()
                        bSuccess = MyCommon.LXSsp.Parameters("@Success").Value
                        MyCommon.Close_LXSsp()
                        If Not bSuccess Then
                            MyCommon.QueryStr = "dbo.pt_PointsPrograms_Delete"
                            MyCommon.Open_LRTsp()
                            MyCommon.LRTsp.Parameters.Add("@ProgramID", SqlDbType.BigInt).Value = ProgramID
                            MyCommon.LRTsp.ExecuteNonQuery()
                            MyCommon.Close_LRTsp()
                            infoMessage = String.Format(Copient.PhraseLib.Lookup("promovar.duplicate-id", LanguageID), PromoVarID)
                        End If
                    Else
                        MyCommon.QueryStr = "dbo.pc_PointsVar_Create"
                        MyCommon.Open_LXSsp()
                        MyCommon.LXSsp.Parameters.Add("@ProgramID", SqlDbType.BigInt).Value = ProgramID
                        MyCommon.LXSsp.Parameters.Add("@VarID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                        MyCommon.LXSsp.ExecuteNonQuery()
                        PromoVarID = MyCommon.LXSsp.Parameters("@VarID").Value
                        MyCommon.Close_LXSsp()
                    End If
                End If

                If bSuccess Then
                    ProgramDescription = MyCommon.Parse_Quotes(Request.QueryString.Item("desc"))
                    ProgramDescription = Logix.TrimAll(ProgramDescription)
                    iCategoryID = MyCommon.Extract_Val(Request.QueryString("form_Category"))
                    MyCommon.QueryStr = "update PointsPrograms with (RowLock) SET " & _
                                        "Description=N'" & ProgramDescription & "', " & _
                                        "PromoVarID=" & PromoVarID & ", " & _
                                        "CategoryID=" & iCategoryID & " " & _
                                        "where ProgramID=" & ProgramID & ";"
                    MyCommon.LRT_Execute()
                    If (IsExternalProgram) Then
                        MyCommon.QueryStr = "update PointsPrograms with (RowLock) SET " & _
                                            "ExternalProgram=1, " & _
                                            "ExtHostTypeID=" & ExtHostTypeID & ", " & _
                                "ExtHostProgramID='" & ExternalID & "', " & _
                                "ExtHostPartnerCode='" & PartnerCode & "', " & _
                                "ExtHostPartnerID='" & PartnerID & "', " & _
                                "ExtHostFuelProgram=" & IIf(ExtHostFuelProgram, "1", "0") & ", " & _
                                "ExtHostCardBINMin=" & ExtHostCardBINMin & ", " & _
                                "ExtHostCardBINMax=" & ExtHostCardBINMax & ", "
                        If DecimalValues Then
                            MyCommon.QueryStr &= "DecimalValues=1 "
                        Else
                            MyCommon.QueryStr &= "DecimalValues=0 "
                        End If
                        MyCommon.QueryStr &= "where ProgramID=" & ProgramID & ";"
                        MyCommon.LRT_Execute()
                        If MyCommon.Fetch_SystemOption(80) = "1" Then
                            MyCommon.QueryStr = "update PromoVariables with (RowLock) SET " & _
                                                "ExternalID='" & MyCryptLib.SQL_StringEncrypt(ExternalID) & "', LastUpdate=GetDate() " & _
                                                "where PromoVarID=" & PromoVarID & ";"
                            MyCommon.LXS_Execute()
                        End If
                    End If

                    Response.Status = "301 Moved Permanently"
                    If (ProgramID = 0) Then
                        Response.AddHeader("Location", "point-edit.aspx?ProgramGroupID=" & ProgramID & "&infoMessage=" & infoMessage)
                    Else
                        Response.AddHeader("Location", "point-edit.aspx?ProgramGroupID=" & ProgramID)
                    End If
                    GoTo done
                End If
            End If
        End If
    ElseIf (Request.QueryString("save") <> "") Then
        ' somebody clicked save
        l_pgID = MyCommon.Extract_Val(Request.QueryString("ProgramGroupID"))
        AutoDelete = Request.QueryString("AutoDelete") = "1"
        ScorecardID = MyCommon.Extract_Val(Request.QueryString("ScorecardID"))
        ExternalID = MyCommon.Parse_Quotes(Request.QueryString.Item("externalid"))
        If MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) Then
            If MyCommon.Fetch_CPE_SystemOption(100) = "" Then
                BeginUPC = 0
            Else
                BeginUPC = CDec(MyCommon.Fetch_CPE_SystemOption(100))
            End If
            If MyCommon.Fetch_CPE_SystemOption(101) = "" Then
                EndUPC = 0
            Else
                EndUPC = CDec(MyCommon.Fetch_CPE_SystemOption(101))
            End If
            If AdjustmentUPC = "" Then
                ValidSetUPC = True
            ElseIf Not AllDigits(AdjustmentUPC) Then
                ValidSetUPC = False
                infoMessage = Copient.PhraseLib.Lookup("error.InvalidUPCEntry", LanguageID)
            Else
                If (CDec(AdjustmentUPC) < 0) Then
                    ValidSetUPC = False
                    infoMessage = Copient.PhraseLib.Lookup("error.InvalidUPCEntry", LanguageID)
                ElseIf AdjustmentUPC = 0 AndAlso BeginUPC = 0 AndAlso EndUPC = 0 Then
                    ValidSetUPC = True
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
                                        "where AdjustmentUPC='" & AdjustmentUPC & "' and ProgramID<>" & l_pgID & ";"
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

        If bPointsMigrationEnabled Then
            With Request.QueryString
                MigrationDate = .Item("migration-date")
                If MigrationDate = "" Then
                    MigrationDate = Now.ToShortDateString()
                End If
                If IsDate(MigrationDate) Then
                    MigrationHr = .Item("migration-hr")
                    If MigrationHr = "" Then
                        MigrationHr = "00"
                    End If

                    MigrationMin = .Item("migration-min")
                    If MigrationMin = "" Then
                        MigrationMin = "00"
                    End If
                    MigrationDatetTime = MigrationDate & " " & MigrationHr & ":" & MigrationMin & ":00"
                Else
                    bValidMigrationData = False
                End If
                lMigrationProgramID = .Item("migration-id")
            End With
        End If

        If bPointsDeletionEnabled Then
            With Request.QueryString
                If (.Item("enabledeletedate") = "1") Then
                    bEnableDeleteDate = True
                    DeletionDate = .Item("deletion-date")
                    If IsDate(DeletionDate) Then
                        DeletionHr = .Item("deletion-hr")
                        If DeletionHr = "" Then
                            DeletionHr = "00"
                        End If

                        DeletionMin = .Item("deletion-min")
                        If DeletionMin = "" Then
                            DeletionMin = "00"
                        End If
                        DeletionDateTime = DeletionDate & " " & DeletionHr & ":" & DeletionMin & ":00"
                    Else
                        bValidDeletionDate = False
                    End If
                Else
                    bEnableDeleteDate = False
                End If
            End With
        End If

        With Request.QueryString
            If MyCommon.Parse_Quotes(Logix.TrimAll(.Item("name"))) = "" Then
                infoMessage = Copient.PhraseLib.Lookup("point-edit.noname", LanguageID)
            ElseIf ValidSetUPC = False AndAlso BeginUPC = 0 AndAlso EndUPC = 0 Then
                infoMessage = Copient.PhraseLib.Lookup("error.UndefineUPCCodeRange", LanguageID)
            ElseIf ValidSetUPC = False Then
                infoMessage = Copient.Detokenize("sv-edit.UPCOutside", LanguageID, BeginUPC, EndUPC)
            ElseIf ValidSingleUPC = False Then
                infoMessage = Copient.PhraseLib.Lookup("sv-edit.UPCInUse", LanguageID)
            ElseIf bValidMigrationData = False Then
                infoMessage = Copient.PhraseLib.Lookup("point-edit.InvalidMigrationDate", LanguageID)
            ElseIf bValidDeletionDate = False Then
                infoMessage = Copient.PhraseLib.Lookup("point-edit.InvalidDeletionDate", LanguageID)
            ElseIf (IsExternalProgram AndAlso ExternalID = "" AndAlso MyCommon.Fetch_SystemOption(223) = "1") Then
                infoMessage = Copient.PhraseLib.Lookup("error.noextid", LanguageID)
            Else
                MyCommon.QueryStr = "SELECT ProgramID, ProgramName FROM PointsPrograms " & _
                                    "WHERE ProgramName = '" & MyCommon.Parse_Quotes(.Item("name")) & "' " & _
                                    "AND ProgramID <> " & l_pgID & " " & _
                                    "AND Deleted = 0;"
                rst = MyCommon.LRT_Select
                If (rst.Rows.Count > 0) Then
                    infoMessage = Copient.PhraseLib.Lookup("point-edit.nameused", LanguageID)
                ElseIf (MyCommon.Extract_Val(Request.QueryString("ScorecardID")) > 0) AndAlso ((Request.QueryString("ScorecardDesc") = "" And Not MultiLanguageEnabled) OrElse (Request.QueryString("ScorecardDesc") = "" And Request.QueryString("ScorecardDesc_" & DefaultLanguageCode) = "" And MultiLanguageEnabled)) Then
                    infoMessage = "When selecting a scorecard, please also enter scorecard text."
                Else
                    MyCommon.QueryStr = "UPDATE PointsPrograms with (RowLock) SET " &
                                        "ProgramName=N'" & MyCommon.Parse_Quotes(Logix.TrimAll(.Item("name"))) & "'," &
                                        "Description=N'" & MyCommon.Parse_Quotes(Logix.TrimAll(.Item("desc"))) & "', " &
                                        "AutoDelete=" & IIf(Request.QueryString("autodelete") = "1", 1, 0) & ", " &
                                        "LastUpdate=getDate(), " &
                                        "ScorecardID=" & MyCommon.Extract_Val(Request.QueryString("ScorecardID")) & "," &
                                        "VisibleToCustomers=" & IIf(MyCommon.Parse_Quotes(Logix.TrimAll(.Item("VisibleToCustomers"))) = "1", 1, 0)
                    If MyCommon.Extract_Val(Request.QueryString("ScorecardID")) = 0 Then
                        MyCommon.QueryStr &= ", ScorecardDesc=NULL"
                    Else
                        MyCommon.QueryStr &= ", ScorecardDesc=N'" & MyCommon.Parse_Quotes(.Item("ScorecardDesc")) & "'"
                    End If
                    If MyCommon.Extract_Val(Request.QueryString("CAMProgram")) = "1" Then
                        MyCommon.QueryStr &= ", CAMProgram=1"
                    Else
                        MyCommon.QueryStr &= ", CAMProgram=0"
                    End If
                    MyCommon.QueryStr &= ", ScorecardBold=" & IIf(Request.QueryString("ScorecardBold") = "1", 1, 0)
                    If AdjustmentUPC = "" OrElse CDec(AdjustmentUPC) = 0 Then
                        MyCommon.QueryStr &= ", AdjustmentUPC=NULL"
                    Else
                        MyCommon.QueryStr &= ", AdjustmentUPC=N'" & AdjustmentUPC & "'"
                    End If
                    If bEnableCategories Then
                        iCategoryID = MyCommon.Extract_Val(Request.QueryString("form_Category"))
                        MyCommon.QueryStr &= ", CategoryID=" & iCategoryID
                    Else
                        MyCommon.QueryStr &= ", CategoryID=NULL"
                    End If
                    If IsExternalProgram Then
                        MyCommon.QueryStr &= ", ExtHostProgramID='" & ExternalID & "'"
                        MyCommon.QueryStr &= ", ExtHostPartnerCode='" & PartnerCode & "'"
                        MyCommon.QueryStr &= ", ExtHostPartnerID='" & PartnerID & "'"
                        MyCommon.QueryStr &= ", ExtHostFuelProgram=" & IIf(ExtHostFuelProgram, 1, 0)
                        MyCommon.QueryStr &= ", ExtHostCardBINMin=" &  ExtHostCardBINMin
                        MyCommon.QueryStr &= ", ExtHostCardBINMax=" &  ExtHostCardBINMax 
                    End If



                    If UEInstalled Then
                        MyCommon.QueryStr &= ", ReturnHandlingTypeID=" & IIf(Request.QueryString("returnsHandling") <> "", MyCommon.Extract_Val(Request.QueryString("returnsHandling")), 1) & ", " & _
                                              "DisallowRedeemInEarnTrans=" & IIf(Request.QueryString("disallowRedeemInTrans") = "1", 1, 0) & ", " & _
                                              "AllowNegativeBal=" & IIf(Request.QueryString("allowNegBal") = "1", 1, 0) & " "
                    End If

                    If bPendingEnabled Then
                        MyCommon.QueryStr &= ", ApplyEarnedPendingPoints='" & IIf(Request.QueryString("applyRedeemedPending") = "1", 0, 1) & "'"
                    End If

                    MyCommon.QueryStr &= " WHERE ProgramID=" & MyCommon.Parse_Quotes(.Item("ProgramGroupID")) & ";"
                    MyCommon.LRT_Execute()

                    If UEInstalled Then
                        AllowAnyCustomer_UE = MyCommon.NZ(Request.QueryString("hdnAllowAnyCustomerUE"), False)
                        MyCommon.QueryStr = "dbo.pt_PointsProgramsPromoEngineSettings_InsertUpdate"
                        MyCommon.Open_LRTsp()
                        MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = 9
                        MyCommon.LRTsp.Parameters.Add("@ProgramID", SqlDbType.BigInt).Value = l_pgID
                        MyCommon.LRTsp.Parameters.Add("@AllowAnyCustomer", SqlDbType.Bit).Value = AllowAnyCustomer_UE
                        MyCommon.LRTsp.Parameters.Add("@PKID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                        MyCommon.LRTsp.ExecuteNonQuery()
                        AllowAnyCustomer_UE_PKID = MyCommon.LRTsp.Parameters("@PKID").Value
                        MyCommon.Close_LRTsp()
                    End If

                    ' Update the ExternalID field (if necessary)
                    If IsExternalProgram AndAlso MyCommon.Fetch_SystemOption(80) = "1" Then
                        MyCommon.QueryStr = "UPDATE PromoVariables with (RowLock) SET " & _
                                            "ExternalID='" & MyCryptLib.SQL_StringEncrypt(ExternalID) & "', LastUpdate=GetDate() " & _
                                            "WHERE LinkID=" & MyCommon.Parse_Quotes(.Item("ProgramGroupID")) & ";"
                        MyCommon.LXS_Execute()
                    End If

                    If bPointsMigrationEnabled Then

                        MyCommon.QueryStr = "select MigrationProgramID, MigrationDate from PointsProgramMigrations with (NoLock) " & _
                                "where ProgramID=" & l_pgID & " and Deleted=0;"
                        rst2 = MyCommon.LRT_Select
                        If rst2.Rows.Count > 0 Then
                            MyCommon.QueryStr = "update PointsProgramMigrations with (rowLock) " & _
                                                "set MigrationProgramID=" & lMigrationProgramID & ",MigrationDate='" & MigrationDatetTime & "',LastUpdate=getdate(),LastUpdatedByAdminID=" & AdminUserID & " " & _
                                                "where ProgramID=" & l_pgID & " and Deleted=0;"
                        Else
                            MyCommon.QueryStr = "insert into PointsProgramMigrations with (rowLock) (ProgramID, MigrationProgramID, MigrationDate, LastMigrationDate, LastEmailDate, Deleted, LastUpdate,LastUpdatedByAdminID) " & _
                                                "values (" & l_pgID & "," & lMigrationProgramID & ",'" & MigrationDatetTime & "',null,null,0,getdate()," & AdminUserID & ");"
                        End If
                        MyCommon.LRT_Execute()
                    End If

                    ''save points deletion date
                    If bPointsDeletionEnabled Then
                        If bEnableDeleteDate Then
                            MyCommon.QueryStr = "select * from PointsProgramDeletions with (NoLock) where ProgramID = " & l_pgID & " and Deleted=0;"
                            rst2 = MyCommon.LRT_Select
                            If rst2.Rows.Count > 0 Then
                                MyCommon.QueryStr = "UPDATE PointsProgramDeletions with (rowLock) " & _
                                                    "set DeletionDate = '" & DeletionDateTime & "', LastUpdate=getdate(),LastUpdatedByAdminID=" & AdminUserID & " " & _
                                                    "where ProgramID=" & l_pgID & " and Deleted=0;"
                            Else
                                MyCommon.QueryStr = "INSERT into PointsProgramDeletions with (rowLock) (ProgramID,DeletionDate,Deleted,LastUpdate,LastUpdatedByAdminID) " & _
                                    " values(" & l_pgID & ",'" & DeletionDateTime & "',0,getdate()," & AdminUserID & ");"
                            End If
                            MyCommon.LRT_Execute()
                        Else
                            MyCommon.QueryStr = "select * from PointsProgramDeletions with (NoLock) where ProgramID = " & l_pgID & ";"
                            rst2 = MyCommon.LRT_Select
                            If rst2.Rows.Count > 0 Then
                                MyCommon.QueryStr = "DELETE from PointsProgramDeletions with (rowLock) " & _
                                                    "where ProgramID=" & l_pgID & ";"
                                MyCommon.LRT_Execute()
                            End If
                        End If
                    End If
                    'save program name
                    MLI.ItemID = l_pgID
                    MLI.MLTableName = "PointsProgramTranslations"
                    MLI.MLColumnName = "ProgramName"
                    MLI.MLIdentifierName = "ProgramID"
                    MLI.StandardTableName = "PointsPrograms"
                    MLI.StandardColumnName = "ProgramName"
                    MLI.StandardIdentifierName = "ProgramID"
                    MLI.InputName = "name"
                    Localization.SaveTranslationInputs(MyCommon, MLI, Request.QueryString, 2)
                    'Save scorecard multilanguage inputs
                    MLI.ItemID = l_pgID
                    MLI.MLTableName = "PointsProgramTranslations"
                    MLI.MLColumnName = "ScorecardDesc"
                    MLI.MLIdentifierName = "ProgramID"
                    MLI.StandardTableName = "PointsPrograms"
                    MLI.StandardColumnName = "ScorecardDesc"
                    MLI.StandardIdentifierName = "ProgramID"
                    MLI.InputName = "ScorecardDesc"
                    Localization.SaveTranslationInputs(MyCommon, MLI, Request.QueryString, 2)

                    MyCommon.Activity_Log(7, l_pgID, AdminUserID, Copient.PhraseLib.Lookup("history.point-edit", LanguageID))
                End If
            End If
        End With

    ElseIf (Request.QueryString("add") <> "") Then
        l_pgID = MyCommon.NZ(Request.QueryString("ProgramGroupID"), "0")
        pgPromoVarID = MyCommon.Extract_Val(Request.QueryString("PromoVarID"))
        CustExtID = Request.QueryString("clientuserid1")
        BalanceAmt = MyCommon.Extract_Val(Request.QueryString("balance"))
        If (CustExtID <> "" And BalanceAmt > 0) Then
            CustExtID = MyCommon.Pad_ExtCardID(CustExtID, 0)
            MyCommon.QueryStr = "dbo.pa_CustomerPoints_Add"
            MyCommon.Open_LXSsp()
            MyCommon.LXSsp.Parameters.Add("@PromoVarID", SqlDbType.BigInt).Value = pgPromoVarID
            MyCommon.LXSsp.Parameters.Add("@ProgramID", SqlDbType.BigInt).Value = l_pgID
            MyCommon.LXSsp.Parameters.Add("@ExtCardID", SqlDbType.NVarChar, 400).Value = MyCryptLib.SQL_StringEncrypt(CustExtID)
            MyCommon.LXSsp.Parameters.Add("@CardTypeID", SqlDbType.Int).Value = 0 ' default to customer
            MyCommon.LXSsp.Parameters.Add("@Amount", SqlDbType.Decimal, 12).Value = BalanceAmt
            MyCommon.LXSsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
            MyCommon.LXSsp.ExecuteNonQuery()
            Status = MyCommon.LXSsp.Parameters("@Status").Value
            If (Status = -1) Then
                infoMessage += "  " & Copient.PhraseLib.Lookup("term.customer", LanguageID) & " " & CustExtID & " " & Copient.PhraseLib.Lookup("term.notfound", LanguageID).ToLower & "."
            ElseIf (Status = 1) Then
                statusMessage += "  " & Copient.PhraseLib.Lookup("term.customer", LanguageID) & " " & CustExtID & " " & Copient.PhraseLib.Lookup("point-edit.updated", LanguageID)
                MyCommon.Activity_Log(7, l_pgID, AdminUserID, Copient.PhraseLib.Lookup("history.point-edit.updated", LanguageID) & " (" & CustExtID & ").")
            ElseIf (Status = 2) Then
                statusMessage += "  " & Copient.PhraseLib.Lookup("term.customer", LanguageID) & " " & CustExtID & " " & Copient.PhraseLib.Lookup("point-edit.added", LanguageID)
                MyCommon.Activity_Log(7, l_pgID, AdminUserID, Copient.PhraseLib.Lookup("history.point-edit.added", LanguageID) & " (" & CustExtID & ").")
            End If
        Else
            If (CustExtID = "") Then infoMessage += " " & Copient.PhraseLib.Lookup("point-edit.customernotentered", LanguageID)
            If (Request.QueryString("balance") = "") Then infoMessage += " " & Copient.PhraseLib.Lookup("point-edit.balancenotentered", LanguageID)
        End If
    ElseIf (Request.QueryString("remove") <> "") Then
        l_pgID = MyCommon.NZ(Request.QueryString("ProgramGroupID"), "0")
        pgPromoVarID = MyCommon.Extract_Val(Request.QueryString("PromoVarID"))
        CustBalPK = MyCommon.Extract_Val(Request.QueryString("custbalID"))
        If (Request.QueryString("custbalID") <> "") Then
            MyCommon.QueryStr = "delete from points with (RowLock) where PKID = " & CustBalPK
            MyCommon.LXS_Execute()
            statusMessage += Copient.PhraseLib.Lookup("point-edit.removebalance", LanguageID)
        End If
        MyCommon.Activity_Log(7, l_pgID, AdminUserID, Copient.PhraseLib.Lookup("point-edit.removebalance", LanguageID))
    ElseIf (Request.QueryString("ProgramGroupID") <> "") Then
        ' simple edit/search mode
        l_pgID = MyCommon.NZ(Request.QueryString("ProgramGroupID"), "0")
    ElseIf (Request.Form("ProgramID") <> "") Then
        l_pgID = MyCommon.Extract_Val(Request.Form("ProgramID"))
    Else
        ' no group id passed ... what now ?
        l_pgID = "0"
    End If

    ' grab this point program
    MyCommon.QueryStr = "SELECT PP.ProgramID, PP.ProgramName, PP.CreatedDate, PP.LastUpdate, PP.Description, PP.PromoVarID, " & _
                      "PP.ExternalProgram, PP.AutoDelete, PP.ScorecardID, PP.ScorecardDesc, PP.ScorecardBold, PP.AdjustmentUPC, PP.CAMProgram, PP.CategoryID, " & _
                      "PP.ExtHostTypeID, PP.ExtHostProgramID, PP.ReturnHandlingTypeID, PP.DisallowRedeemInEarnTrans, PP.AllowNegativeBal, PP.ApplyEarnedPendingPoints, PP.VisibleToCustomers, PP.ExtHostPartnerCode, " & _
                      "PP.ExtHostPartnerID, PP.ExtHostFuelProgram, PP.ExtHostCardBINMin, PP.ExtHostCardBINMax FROM PointsPrograms AS PP WITH (NoLock) " & _
                      "WHERE Deleted=0 AND ProgramID='" & l_pgID & "';"
    dstPrograms = MyCommon.LRT_Select
    If (dstPrograms.Rows.Count > 0) Then
        pgDescription = MyCommon.NZ(dstPrograms.Rows(0).Item("Description"), "")
        pgCreated = MyCommon.NZ(dstPrograms.Rows(0).Item("CreatedDate"), "")
        pgUpdated = MyCommon.NZ(dstPrograms.Rows(0).Item("LastUpdate"), "")
        pgPromoVarID = MyCommon.NZ(dstPrograms.Rows(0).Item("PromoVarID"), 0)
        l_pgID = MyCommon.NZ(dstPrograms.Rows(0).Item("ProgramID"), 0)
        pgName = MyCommon.NZ(dstPrograms.Rows(0).Item("ProgramName"), "")
        IsExternalProgram = MyCommon.NZ(dstPrograms.Rows(0).Item("ExternalProgram"), False)
        AutoDelete = MyCommon.NZ(dstPrograms.Rows(0).Item("AutoDelete"), True)
        ScorecardID = MyCommon.NZ(dstPrograms.Rows(0).Item("ScorecardID"), 0)
        ScorecardDesc = MyCommon.NZ(dstPrograms.Rows(0).Item("ScorecardDesc"), "")
        CAMProgram = MyCommon.NZ(dstPrograms.Rows(0).Item("CAMProgram"), 0)
        ScorecardBold = MyCommon.NZ(dstPrograms.Rows(0).Item("ScorecardBold"), False)
        AdjustmentUPC = MyCommon.NZ(dstPrograms.Rows(0).Item("AdjustmentUPC"), "")
        iCategoryID = MyCommon.NZ(dstPrograms.Rows(0).Item("CategoryID"), 0)
        VisibleToCustomers = MyCommon.NZ(dstPrograms.Rows(0).Item("VisibleToCustomers"), 0)
        If IsExternalProgram Then
            If IsDBNull(dstPrograms.Rows(0).Item("ExtHostTypeID")) Then
                MyCommon.QueryStr = "select ExternalID from PromoVariables with (NoLock) where PromoVarID =" & pgPromoVarID
                rst = MyCommon.LXS_Select
                If (rst.Rows.Count > 0) Then
                    ExternalID = MyCryptLib.SQL_StringDecrypt(rst.Rows(0).Item("ExternalID").ToString())
                End If
            Else
                ExternalID = MyCommon.NZ(dstPrograms.Rows(0).Item("ExtHostProgramID"), 0)
        PartnerCode = MyCommon.NZ(dstPrograms.Rows(0).Item("ExtHostPartnerCode"), "")
        PartnerID = MyCommon.NZ(dstPrograms.Rows(0).Item("ExtHostPartnerID"), "")
        ExtHostTypeID = MyCommon.NZ(dstPrograms.Rows(0).Item("ExtHostTypeID"), 0)
        ExtHostFuelProgram = MyCommon.NZ(dstPrograms.Rows(0).Item("ExtHostFuelProgram"), True)
        ExtHostCardBINMin = MyCommon.NZ(dstPrograms.Rows(0).Item("ExtHostCardBINMin"), 0)
        ExtHostCardBINMax = MyCommon.NZ(dstPrograms.Rows(0).Item("ExtHostCardBINMax"), 0)
            End If
        Else
            XID = MyCommon.NZ(dstPrograms.Rows(0).Item("ExtHostProgramID"), "")
        End If
        ReturnHandlingTypeID = MyCommon.NZ(dstPrograms.Rows(0).Item("ReturnHandlingTypeID"), 1)
        DisallowRedeemInTrans = MyCommon.NZ(dstPrograms.Rows(0).Item("DisallowRedeemInEarnTrans"), 0)
        AllowNegativeBal = MyCommon.NZ(dstPrograms.Rows(0).Item("AllowNegativeBal"), 0)
        bApplyEarnedPending = MyCommon.NZ(dstPrograms.Rows(0).Item("ApplyEarnedPendingPoints"), bDefaultPendingIsUseRedeem)

        If UEInstalled Then
            If AllowAnyCustomer_UE_PKID = 0 Then
                MyCommon.QueryStr = "SELECT PKID, AllowAnyCustomer from PointsProgramsPromoEngineSettings " & _
                                    "WHERE ProgramID = @ProgramID AND EngineID = @EngineID"
                MyCommon.DBParameters.Add("@ProgramID", SqlDbType.BigInt).Value = l_pgID
                MyCommon.DBParameters.Add("@EngineID", SqlDbType.Int).Value = 9
                Dim dt As DataTable = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                If (dt.Rows.Count > 0) Then
                    AllowAnyCustomer_UE_PKID = CMS.Utilities.NZ(dt.Rows(0).Item("PKID"), 0)
                    AllowAnyCustomer_UE = CMS.Utilities.NZ(dt.Rows(0).Item("AllowAnyCustomer"), False)
                End If
            End If
            If (AllowAnyCustomer_UE) Then
                MyCommon.QueryStr = "dbo.pa_IsAnyCustomerOffersExistForPointsProgam"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = 9
                MyCommon.LRTsp.Parameters.Add("@PointProgramID", SqlDbType.BigInt).Value = l_pgID
                MyCommon.LRTsp.Parameters.Add("@result", SqlDbType.Bit).Direction = ParameterDirection.Output
                MyCommon.LRTsp.ExecuteNonQuery()
                IsAnyCustomerOffersExist_UE = MyCommon.LRTsp.Parameters("@result").Value
                MyCommon.Close_LRTsp()
            End If
        End If

        ' Let's see if any points deliverables using this program have ScorecardID and ScorecardDesc set
        MyCommon.QueryStr = "select ScorecardID, ScorecardDesc from CPE_DeliverablePoints with (NoLock) " & _
                            "where ProgramID=" & l_pgID & " and ScorecardID>0;"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
            RewardsSetScorecard = True
        End If
        If ExtHostTypeID > 0 Then
            MyCommon.QueryStr = "select Description from ExtHostTypes with (NoLock) where ExtHostTypeID=" & ExtHostTypeID & ";"
            rst2 = MyCommon.LRT_Select
            If rst2.Rows.Count > 0 Then
                ExtHostDesc = MyCommon.NZ(rst2.Rows(0).Item("Description"), "")
            End If
        End If

        If bPointsMigrationEnabled Then
            MyCommon.QueryStr = "select MigrationProgramID, MigrationDate, LastMigrationDate, LastEmailDate from PointsProgramMigrations with (NoLock) " & _
                                "where ProgramID=" & l_pgID & " and Deleted=0;"
            rst2 = MyCommon.LRT_Select
            If rst2.Rows.Count > 0 Then
                tempDateTime = MyCommon.NZ(rst2.Rows(0).Item("MigrationDate"), Nothing)
                lMigrationProgramID = MyCommon.NZ(rst2.Rows(0).Item("MigrationProgramID"), 0)
                LastMigrationDate = MyCommon.NZ(rst2.Rows(0).Item("LastMigrationDate"), Nothing)
            Else
                tempDateTime = Nothing
                lMigrationProgramID = 0
                LastMigrationDate = Nothing
            End If
            If tempDateTime = Nothing Then
                tempDateTime = Now.AddYears(1)
            End If
            MigrationDate = tempDateTime.ToString(sDateOnlyFormat)
            MigrationHr = tempDateTime.ToString(sHourOnlyFormat)
            MigrationMin = tempDateTime.ToString(sMinutesOnlyFormat)
        End If

        ''grab the points deletion information
        If bPointsDeletionEnabled Then
            MyCommon.QueryStr = "select DeletionDate, LastDeletionDate, LastEmailDate from PointsProgramDeletions with (NoLock) " & _
                                "where ProgramID=" & l_pgID & " and Deleted=0;"
            rst2 = MyCommon.LRT_Select
            If rst2.Rows.Count > 0 Then
                bEnableDeleteDate = True
                tempDateTime = MyCommon.NZ(rst2.Rows(0).Item("DeletionDate"), Nothing)
                DeletionDate = tempDateTime.ToString(sDateOnlyFormat)
                DeletionHr = tempDateTime.ToString(sHourOnlyFormat)
                DeletionMin = tempDateTime.ToString(sMinutesOnlyFormat)
                LastDeletionDate = MyCommon.NZ(rst2.Rows(0).Item("LastDeletionDate"), Nothing)
            Else
                bEnableDeleteDate = False
                DeletionDate = ""
                DeletionHr = ""
                DeletionMin = ""
            End If
        End If

    ElseIf (Request.QueryString("new") <> Copient.PhraseLib.Lookup("term.new", LanguageID)) AndAlso (IIf(l_pgID <> "", l_pgID, 0) > 0) Then
        ' check if this is a deleted points program
        MyCommon.QueryStr = "select ProgramName from PointsPrograms with (NoLock) where ProgramID=" & l_pgID & " and deleted =1"
        rst = MyCommon.LRT_Select()
        If (rst.Rows.Count > 0) Then
            pgName = MyCommon.NZ(rst.Rows(0).Item("ProgramName"), "")
        Else
            pgName = ""
        End If

        Send_HeadBegin("term.pointsprogram", , l_pgID)
        Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
        Send_Metas()
        Send_Links(Handheld)

        Send_Scripts(New String() {"datePicker.js"})

        Send_HeadEnd()
        Send_BodyBegin(1)
        Send_Bar(Handheld)
        Send_Help(CopientFileName)
        Send_Logos()
        Send_Tabs(Logix, 5)
        Send_Subtabs(Logix, 51, 5, , l_pgID)
        Send("")
        Send("<div id=""intro"">")
        Send("    <h1 id=""title"">" & Copient.PhraseLib.Lookup("term.pointsprogram", LanguageID) & " #" & l_pgID & ": " & pgName & "</h1>")
        Send("</div>")
        Send("<div id=""main"">")
        Send("    <div id=""infobar"" class=""red-background"">")
        Send("        " & Copient.PhraseLib.Lookup("error.deleted", LanguageID))
        Send("    </div>")
        Send("</div>")
        Send_BodyEnd()
        GoTo done
    Else
        pgPromoVarID = 0
        l_pgID = "0"
        pgDescription = ""
        pgCreated = ""
        pgUpdated = ""
        pgName = Copient.PhraseLib.Lookup("term.newprogram", LanguageID)
    End If

    CpeEngineOnly = MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) And _
                Not (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) And _
                Not (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.Catalina))
    Send_HeadBegin("term.pointsprogram", , l_pgID)
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)

    Send_Scripts(New String() {"datePicker.js"})

    Send_HeadEnd()
    Send_BodyBegin(1)
%>
<script type="text/javascript" language="javascript">
  window.name = "pointEdit";
    var datePickerDivID = "datepicker";
    
    <% Send_Calendar_Overrides(MyCommon) %>

    function disableUnload() {
      window.onunload = null;
    }
    
    
    function elmName(){
      window.onunload = null;
      for(i=0; i<document.mainform.elements.length; i++)
      {
          document.mainform.elements[i].disabled=false;
          //alert(document.mainform.elements[i].name)
      }
      return true;
    }
    
    if (window.captureEvents){
      window.captureEvents(Event.CLICK);
      window.onclick=handlePageClick;
    }
    else {
      document.onclick=handlePageClick;
    }
    
    function handlePageClick(e) {
      var calFrame = document.getElementById('calendariframe');
      var el = (typeof event !== 'undefined') ? event.srcElement : e.target;        
      
      if (el != null) {
        var pickerDiv = document.getElementById(datePickerDivID);
        if (pickerDiv != null && pickerDiv.style.visibility == "visible") {
          if (el.id!="migration-picker" && el.id!="deletion-picker") {
            if (!isDatePickerControl(el.className)) {
              pickerDiv.style.visibility = "hidden";
              pickerDiv.style.display = "none";  
              if (calFrame != null) {
                calFrame.style.visibility = 'hidden';
                calFrame.style.display = 'none';
              }
            }
          } 
          else  {
            pickerDiv.style.visibility = "visible";            
            pickerDiv.style.display = "block";            
            if (calFrame != null) {
              calFrame.style.visibility = 'visible';
              calFrame.style.display = 'block';
            }
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


  <% If (Logix.UserRoles.EditPointsPrograms AndAlso l_pgID > 0) Then %>
  window.onunload= function(){
    if (document.mainform.name.value != document.mainform.name.defaultValue || document.mainform.desc.value != document.mainform.desc.defaultValue) {
      saveChanges = confirm('<%Sendb(Copient.PhraseLib.Lookup("sv-edit.ChangesMade", LanguageID)) %>');
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
  
  // callback function for save changes on unload during navigate away
  function handleAutoFormSubmit() {
    window.onunload = null;
    document.mainform.action = "point-edit.aspx";
    if (ValidateDateTime() && ValidatePromoVarId()) {
      document.mainform.submit();
    }
  }
  
  function toggleDropdown() {
    if (document.getElementById("actionsmenu") != null) {
      bOpen = (document.getElementById("actionsmenu").style.visibility != 'visible');
      if (bOpen) {
        document.getElementById("actionsmenu").style.visibility = 'visible';
        document.mainform.actions.value = '<% Sendb(Copient.PhraseLib.Lookup("term.actions", LanguageID)) %>▲';
      } else {
        document.getElementById("actionsmenu").style.visibility = 'hidden';
        document.mainform.actions.value = '<% Sendb(Copient.PhraseLib.Lookup("term.actions", LanguageID)) %>▼';
      }
    }
  }
   
  function handleExtProg_Click(IntegrationType, bChecked) {
    var elemExtID = document.getElementById("externalid");
    var elemExtHost = document.getElementById("exthosttypeid");
    var elemDecimalValues = document.getElementById("decimalvalues");
    
    if (elemExtID != null) {
      if (!bChecked) {
        elemExtID.value = "";
        elemExtID.disabled = true;
        elemExtHost.disabled = true;
        elemDecimalValues.checked = false;
        elemDecimalValues.disabled = true;
      } else {
        elemExtID.disabled = false;
        elemExtHost.disabled = false;
        elemDecimalValues.disabled = false;
        elemExtID.focus();
      }
    }
  }
  
  function handleEnableDeleteDate_Click(IntegrationType, bChecked) {
    var elemDeletionDate = document.getElementById("deletion-date");
    var elemDeletionHr = document.getElementById("deletion-hr");
    var elemDeletionMin = document.getElementById("deletion-min");
    
    if (elemDeletionDate != null) {
      if (!bChecked) {
        elemDeletionDate.disabled = true;
        elemDeletionHr.disabled = true;
        elemDeletionMin.disabled = true;
      } else {
        elemDeletionDate.disabled = false;
        elemDeletionHr.disabled = false;
        elemDeletionMin.disabled = false;
        elemDeletionDate.focus();
      }
    }
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

  function ValidateDateTime() {
  var elemMigrationDate = document.getElementById("migration-date");
  var elemMigrationHr = document.getElementById("migration-hr");
  var elemMigrationMin = document.getElementById("migration-min");  
  var retVal = true;
  
  <% If bPointsMigrationEnabled Then%>
    if (elemMigrationDate != null) {
      retVal = isDate(elemMigrationDate.value);
    }
    if (retVal == true && elemMigrationHr != null && (!isInteger(elemMigrationHr.value) || (parseInt(elemMigrationHr.value) < 0) || (parseInt(elemMigrationHr.value) > 23))) {
      alert('<%Sendb(Copient.PhraseLib.Lookup("point-edit.InvalidMigrationHour", LanguageID)) %>');
      retVal = false;
    }
    if (retVal == true && elemMigrationMin != null && (!isInteger(elemMigrationMin.value) || (parseInt(elemMigrationMin.value) < 0) || (parseInt(elemMigrationMin.value) > 59))) {
		  alert('<%Sendb(Copient.PhraseLib.Lookup("point-edit.InvalidMigrationMinute", LanguageID)) %>');
      retVal = false;
    }
  <% End If %>
  
  var retValDeletion = true;
  <% If bPointsDeletionEnabled Then%>
    var elemDeletionDate = document.getElementById("deletion-date");
    var elemDeletionHr = document.getElementById("deletion-hr");
    var elemDeletionMin = document.getElementById("deletion-min");
    var elemenabledeletedate = document.getElementById("enabledeletedate");
    if (elemenabledeletedate != null && elemenabledeletedate.checked == "1")
    {
      if (elemDeletionDate != null) {
        retValDeletion = isDate(elemDeletionDate.value);
      }
      if (retValDeletion == true && elemDeletionHr != null && (!isInteger(elemDeletionHr.value) || (parseInt(elemDeletionHr.value) < 0) || (parseInt(elemDeletionHr.value) > 23))) {
        alert('<%Sendb(Copient.PhraseLib.Lookup("point-edit.InvalidDeletionHour", LanguageID)) %>');
        retValDeletion = false;
      }
      if (retValDeletion == true && elemDeletionMin != null && (!isInteger(elemDeletionMin.value) || (parseInt(elemDeletionMin.value) < 0) || (parseInt(elemDeletionMin.value) > 59))) {
		    alert('<%Sendb(Copient.PhraseLib.Lookup("point-edit.InvalidDeletionMinute", LanguageID)) %>');
        retValDeletion = false;
      }
    }
  <% End If %>
  return retVal && retValDeletion;
}

  function ValidatePromoVarId() {
  var elemPromoVarId = document.getElementById("promovarid");
  var retVal = true;

  <% If lMinimumAutoGeneratedPromoVarID > 0 Then%>
      if (retVal == true && elemPromoVarId != null && (!isInteger(elemPromoVarId.value) || (parseInt(elemPromoVarId.value) < 0))) {
      alert('<%Sendb(Copient.PhraseLib.Lookup("error.requires_valid_positive_integer", LanguageID)) %>');
      retVal = false;
    }
  <% End If %>

  return retVal;
}

function handleOnSubmit() {
    var retVal = true;
    
    retVal = retVal && validateDescriptionLength();
     if (retVal) {
        window.onunload = null;
        retVal = ValidateDateTime();
        if (retVal) {
          retVal = ValidatePromoVarId();
        };
      }
    return retVal;
}

  function validateDescriptionLength() {
    var elemDesc = document.getElementById("desc");
    var retVal = true;
    if (elemDesc.value != null) {
      if (elemDesc.value.length > 1000) {
        alert('<%Sendb(Copient.PhraseLib.Lookup("CPEoffergen.description", LanguageID))%>');
        elemDesc.focus();
        return false;
      }
    }
    return retVal;
  }

  function updateInputVal(cbElement, hdnfield_id) {
    $("#"+hdnfield_id).val(cbElement.checked);
  }

  $(document).ready(function() {
    var AllowAnyCustomerUE = $("#cbAllowAnyCustomerUE");
    if (AllowAnyCustomerUE.length > 0) {
      updateInputVal(AllowAnyCustomerUE[0], "hdnAllowAnyCustomerUE");
      document.getElementById("hdnAllowAnyCustomerUE").defaultValue = AllowAnyCustomerUE[0].checked;
    }
  });
//   function chooseFile() {
//      document.getElementById("browse").click();
//  }
//  function fileonclick() {
//      var filename = document.getElementById("browse").value;
//      document.getElementById("lblfileupload").innerText = filename.replace("C:\\fakepath\\", "");
//  }
</script>
<%
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  Send_Tabs(Logix, 5)
  Send_Subtabs(Logix, 51, 5, , l_pgID)
  
  If (Logix.UserRoles.AccessPointsPrograms = False) Then
    Send_Denied(1, "perm.points-access")
    Send_BodyEnd()
    GoTo done
  End If
%>
<form id="mainform" name="mainform" action="point-edit.aspx" method="get" onsubmit="elmName(); return handleOnSubmit();">
<%
  If (Request.QueryString("New") = Copient.PhraseLib.Lookup("term.new", LanguageID)) Or (Request.QueryString("ProgramGroupID") = Copient.PhraseLib.Lookup("term.new", LanguageID).ToLower) Or (Request.QueryString("ProgramGroupID") = "0") Or (Request.QueryString("ProgramGroupID") = "") Then
  Else
    Send("<input type=""hidden"" id=""CAMProgram"" name=""CAMProgram"" value=""" & IIf(CAMProgram, "1", "0") & """ />")
  End If
%>
<div id="intro">
  <h1 id="title">
    <%
      If l_pgID = 0 Then
        Sendb(Copient.PhraseLib.Lookup("term.newpointsprogram", LanguageID))
      Else
        If CAMProgram Then
          Sendb(Copient.PhraseLib.Lookup("term.cam", LanguageID) & " ")
        End If
        Sendb(Copient.PhraseLib.Lookup("term.pointsprogram", LanguageID) & " #" & l_pgID & ": ")
        MyCommon.QueryStr = "SELECT ProgramID,ProgramName FROM PointsPrograms with (NoLock) WHERE ProgramId = " & l_pgID & ";"
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
          ProgramNameTitle = MyCommon.NZ(rst.Rows(0).Item("ProgramName"), "")
        End If
        Sendb(MyCommon.TruncateString(ProgramNameTitle, 35))
      End If
    %>
  </h1>
  <div id="controls">
    <%
      If (l_pgID = 0) Then
        If (Logix.UserRoles.CreatePointsPrograms) Then
          Send_Save()
        End If
      Else
        ShowActionButton = (Logix.UserRoles.CreatePointsPrograms) OrElse (Logix.UserRoles.EditPointsPrograms) OrElse (Logix.UserRoles.DeletePointsPrograms)
        If (ShowActionButton) Then
          Send("<input type=""button"" class=""regular"" id=""actions"" name=""actions"" value=""" & Copient.PhraseLib.Lookup("term.actions", LanguageID) & " ▼"" onclick=""toggleDropdown();"" />")
          Send("<div class=""actionsmenu"" id=""actionsmenu"">")
          If (Logix.UserRoles.EditPointsPrograms) Then
            Send_Save()
          End If
          If (Logix.UserRoles.DeletePointsPrograms) Then
            Dim bEnableBuckOffers As Boolean
            bEnableBuckOffers = bCMInstalled AndAlso (MyCommon.Fetch_CM_SystemOption(137) = "1")
            If bEnableBuckOffers Then
              Dim rst3 As DataTable
              MyCommon.QueryStr = "select isnull(BuckOfferID,0) as BuckOfferID from PointsPrograms with (NoLock) where ProgramId = " & l_pgID & ";"
              MyCommon.DBParameters.Add("@ProgramID", SqlDbType.BigInt).Value = l_pgID
              rst3 = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
              If rst3.Rows.Count > 0 Then
                Dim lParentOfferId As Long
                
                lParentOfferId = rst3.Rows(0).Item(0)
                If lParentOfferId > 0 Then
                  MyCommon.QueryStr = "select ParentOfferID from CM_BuckOffers with (NoLock) where ChildOfferID=0 and ParentOfferID=@ParentOfferID;"
                  MyCommon.DBParameters.Add("@ParentOfferID", SqlDbType.BigInt).Value = lParentOfferId
                  rst3 = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                  If rst3.Rows.Count = 0 Then
                    infoMessage = "This points program was created for Buck offer '" & lParentOfferId & "' but this offer is no longer a Buck Parent"
                    Send_Delete()
                  Else
                    MyCommon.QueryStr = "select OfferID from Offers with (NoLock) where Deleted=0 and OfferID=@ParentOfferID;"
                    MyCommon.DBParameters.Add("@ParentOfferID", SqlDbType.BigInt).Value = lParentOfferId
                    rst3 = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                    If rst3.Rows.Count = 0 Then
                      infoMessage = "This points program was created for Buck offer '" & lParentOfferId & "' but the offer has been deleted"
                      Send_Delete()
                    End If
                  End If
                Else
                  Send_Delete()
                End If
              Else
                Send_Delete()
              End If
            Else
              Send_Delete()
            End If
          End If
          If (Logix.UserRoles.EditPointsPrograms) AndAlso (l_pgID > 0) Then
            Send_Upload()
          End If
          If (Logix.UserRoles.AccessPointsPrograms) Then
            Send_Download()
          End If
          If (Logix.UserRoles.CreatePointsPrograms) Then
            Send_New()
          End If
          Send("</div>")
        End If
        If MyCommon.Fetch_SystemOption(75) Then
          If (Logix.UserRoles.AccessNotes) Then
            Send_NotesButton(8, l_pgID, AdminUserID)
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
    If (statusMessage <> "") Then
      Send("<div id=""statusbar"" class=""green-background"">" & statusMessage & "</div>")
    End If
  %>
  <%
    If (dstPrograms.Rows.Count > 0 And l_pgID <> 0) Then
  %>
  <div id="column1">
    <div class="box" id="identity">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.identification", LanguageID))%>
        </span>
      </h2>
      <input type="hidden" id="ProgramGroupID" name="ProgramGroupID" value="<% sendb(l_pgID) %>" />
      <input type="hidden" id="PromoVarID" name="PromoVarID" value="<% sendb(pgPromoVarID) %>" />
      <input type="hidden" id="isexternal" name="isexternal" value="<% sendb(IIf(IsExternalProgram, "1", "0")) %>" />
      <%
        If IsExternalProgram Then
          Send("<input type=""hidden"" id=""exthosttypeid"" name=""exthosttypeid"" value=""" & ExtHostTypeID & """ />")
        End If
      %>
      <label for="name">
        <% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>:</label><br />
      <% If (pgName Is Nothing) Then pgName = ""%>
        <%
            MLI.ItemID = l_pgID
            MLI.MLTableName = "PointsProgramTranslations"
            MLI.MLColumnName = "ProgramName"
            MLI.MLIdentifierName = "ProgramID"
            MLI.StandardTableName = "PointsPrograms"
            MLI.StandardColumnName = "ProgramName"
            MLI.StandardIdentifierName = "ProgramID"
            MLI.StandardValue = pgName
            MLI.InputName = "name"
            MLI.InputID = "name"
            MLI.InputType = "text"
            MLI.MaxLength = 50
            MLI.CSSStyle = "width:350px;"
            Send(Localization.SendTranslationInputs(MyCommon, MLI, Request.QueryString, 2))
            %>
    
      <label for="desc">
        <% Sendb(Copient.PhraseLib.Lookup("term.description", LanguageID))%>:</label><br />
      <textarea class="longest" id="desc" name="desc" cols="48" rows="3" maxlength="1000"><% Sendb(pgDescription)%></textarea><br />
      <br class="half" />
      <small>
        <%Send("(" & Copient.PhraseLib.Lookup("CPEoffergen.description", LanguageID) & ")")%></small><br
          class="half" />
      <br class="half" />
      <%
        If IsExternalProgram Then
          Send("<br class=""half"" />")
          Send("<label for=""externalid"">" & Copient.PhraseLib.Lookup("term.externalid", LanguageID) & " (" & IIf(ExtHostTypeID > 0, ExtHostDesc, "") & "):</label><br />")
          Send("<input type=""text"" class=""longest"" id=""externalid"" name=""externalid"" maxlength=""100"" value=""" & ExternalID & """ /><br />")
          Send("<label for=""partnercode"">" & Copient.PhraseLib.Lookup("term.partnercode", LanguageID) & " (" & IIf(ExtHostTypeID > 0, ExtHostDesc, "") & "):</label><br />")
          Send("<input type=""text"" class=""longest"" id=""partnercode"" name=""partnercode"" maxlength=""100"" value=""" & PartnerCode & """ /><br />")
          Send("<label for=""partnerid"">" & Copient.PhraseLib.Lookup("term.partnerID", LanguageID) & " (" & IIf(ExtHostTypeID > 0, ExtHostDesc, "") & "):</label><br />")
          Send("<input type=""text"" class=""longest"" id=""partnerid"" name=""partnerid"" maxlength=""100"" value=""" & PartnerID & """ /><br />")
          Send("<label for=""exthostcardbinmin"">" & Copient.PhraseLib.Lookup("term.binrange", LanguageID) & " " & Copient.PhraseLib.Lookup("term.min", LanguageID) & " (" & IIf(ExtHostTypeID > 0, ExtHostDesc, "") & "):</label><br />")
          Send("<input type=""text"" class=""longest"" id=""exthostcardbinmin"" name=""exthostcardbinmin"" maxlength=""100"" value=""" & ExtHostCardBINMin & """ /><br />")
          Send("<label for=""exthostcardbinmax"">" & Copient.PhraseLib.Lookup("term.binrange", LanguageID) & " " & Copient.PhraseLib.Lookup("term.max", LanguageID) & " (" & IIf(ExtHostTypeID > 0, ExtHostDesc, "") & "):</label><br />")
          Send("<input type=""text"" class=""longest"" id=""exthostcardbinmax"" name=""exthostcardbinmax"" maxlength=""100"" value=""" & ExthostCardBINMax & """ /><br />")
          Send("<input type=""checkbox"" id=""exthostfuelprogram"" name=""exthostfuelprogram"" value=""1""" & IIf(ExtHostFuelProgram, " checked=""checked""", "") & " /><label for=""exthostfuelprogram"">" & Copient.PhraseLib.Lookup("term.fuelprogram", LanguageID) & " (" & IIf(ExtHostTypeID > 0, ExtHostDesc, "") & ")</label><br />")
        Else
          If XID <> "" AndAlso XID <> "0" Then
            Sendb(Copient.PhraseLib.Lookup("term.xid", LanguageID) & ": " & XID & "<br />")
          End If
        End If
          
        Send("<br class=""half"" />")
          
        Send("<input type=""checkbox"" id=""autodelete"" name=""autodelete"" value=""1""" & IIf(AutoDelete, " checked=""checked""", "") & " /><label for=""autodelete"">" & Copient.PhraseLib.Lookup("sv-edit.AutoDelete", LanguageID) & "</label><br />")
        Send("<br />")
        If (CpeEngineOnly) Then
          MyCommon.QueryStr = "SELECT SUM(Cast(Amount as bigint)) AS TotalPoints, COUNT(PKID) as NumRecs FROM Points with (NoLock) WHERE Amount > 0 and ProgramID=" & l_pgID & ";"
        Else
          MyCommon.QueryStr = "SELECT SUM(Cast(Amount as bigint)) AS TotalPoints, COUNT(PKID) as NumRecs FROM Points with (NoLock) WHERE Amount > 0 and PromoVarID=" & pgPromoVarID & ";"
        End If
        dstPoints = MyCommon.LXS_Select
        pgTotalPoints = MyCommon.NZ(dstPoints.Rows(0).Item(0), "0")
        Sendb(Copient.PhraseLib.Lookup("term.created", LanguageID) & " ")
        If pgCreated = "" Then
        Else
          longDate = pgCreated
          longDateString = Logix.ToLongDateTimeString(longDate, MyCommon)
          Send(longDateString)
        End If
        Send("<br />")
        Sendb(Copient.PhraseLib.Lookup("term.edited", LanguageID) & " ")
        If pgUpdated = "" Then
        Else
          longDate = pgUpdated
          longDateString = Logix.ToLongDateTimeString(longDate, MyCommon)
          Send(longDateString)
        End If
        Send("<br />")

        MyCommon.QueryStr = "select LastLoaded, LastLoadMsg from PointsPrograms where ProgramID=" & l_pgID
        rst = MyCommon.LRT_Select()
        If (rst.Rows.Count > 0) Then
          LastUpload = MyCommon.NZ(rst.Rows(0).Item("LastLoaded"), "1/1/1900")
          LastUploadMsg = MyCommon.NZ(rst.Rows(0).Item("LastLoadMsg"), "")
        End If

        MyCommon.QueryStr = "select ProgramID from PointsInsertQueue where ProgramID=" & l_pgID
        rst = MyCommon.LXS_Select
        If (rst.Rows.Count > 0) Then
          Send("<span class=""red"">" & rst.Rows.Count & " " & Copient.PhraseLib.Lookup("cgroup-edit.awaiting", LanguageID) & "</span>")
          Send("<br />")
        End If
        Send("<br class=""half"" />")
        Sendb(Copient.PhraseLib.Lookup("term.contains", LanguageID) & " ")
        Dim PointPhrase As String
        Dim CardholderPhrase As String
        If (pgTotalPoints = 1) Then
          PointPhrase = StrConv(Copient.PhraseLib.Lookup("term.point", LanguageID), VbStrConv.Lowercase)
        Else
          PointPhrase = StrConv(Copient.PhraseLib.Lookup("term.points", LanguageID), VbStrConv.Lowercase)
        End If
        If (MyCommon.NZ(dstPoints.Rows(0).Item("NumRecs"), 0) = 1) Then
          CardholderPhrase = StrConv(Copient.PhraseLib.Lookup("term.cardholder", LanguageID), VbStrConv.Lowercase)
        Else
          CardholderPhrase = StrConv(Copient.PhraseLib.Lookup("term.cardholders", LanguageID), VbStrConv.Lowercase)
        End If
        Response.Write(pgTotalPoints & " ")
        Sendb(PointPhrase & " " & Copient.PhraseLib.Lookup("term.heldby", LanguageID) & " ")
        Sendb(MyCommon.NZ(dstPoints.Rows(0).Item("NumRecs"), 0) & " " & CardholderPhrase)
        Send("<br />")
        If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM) _
         Or MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.Catalina)) Then
          Dim iExtPromoVarId As Integer = 0
          If (MyCommon.Fetch_CM_SystemOption(42) = "1") Then
            MyCommon.QueryStr = "select ExternalID from PromoVariables with (NoLock) where PromoVarID=" & pgPromoVarID & ";"
            rst = MyCommon.LXS_Select
            If (rst.Rows.Count > 0) Then
              iExtPromoVarId = IIf(IsDBNull(rst.Rows(0).Item(0)), 0,MyCryptLib.SQL_StringDecrypt(rst.Rows(0).Item(0).ToString()))
            End If
          End If
          If iExtPromoVarId > 0 Then
            Sendb(Copient.PhraseLib.Lookup("term.var", LanguageID) & ": " & pgPromoVarID & "  (" & Copient.PhraseLib.Lookup("term.xid", LanguageID) & ": " & iExtPromoVarId & ")" & "<br />")
          Else
            Sendb(Copient.PhraseLib.Lookup("term.var", LanguageID) & ": " & pgPromoVarID & "<br />")
          End If
        End If
      %>
      <hr class="hidden" />
    </div>
    <div class="box" id="scorecards" <% Sendb(IIf(MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) OrElse MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CAM), "", " style=""display:none;""")) %>>
      <!-- BZ2079: UE-feature-removal -->
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
            Send("    <small>" & Copient.PhraseLib.Lookup("point-edit.RewardSetScorecard", LanguageID) & "</small>")
            Send("  </td>")
            Send("</tr>")
          End If
          Send("<tr>")
          Send("  <td style=""width:82px;"">")
          Send("    <label for=""ScorecardID"">" & Copient.PhraseLib.Lookup("term.scorecard", LanguageID) & ":</label>")
          Send("  </td>")
          Send("  <td>")
          Send("    <select class=""medium"" id=""ScorecardID"" name=""ScorecardID"" onchange=""toggleScorecardText();""" & IIf(RewardsSetScorecard, " disabled=""disabled""", "") & ">")
          MyCommon.QueryStr = "select ScorecardID, Description from Scorecards with (NoLock) where ScorecardTypeID=1 and Deleted=0 "
          If CAMProgram Then
            MyCommon.QueryStr &= " and EngineID=6;"
          Else
            MyCommon.QueryStr &= " and EngineID=2;"
          End If
          rst2 = MyCommon.LRT_Select
          Send("      <option value=""0"">" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "</option>")
          If rst2.Rows.Count > 0 Then
            For Each row In rst2.Rows
              Send("      <option value=""" & MyCommon.NZ(row.Item("ScorecardID"), 0) & """" & IIf(MyCommon.NZ(row.Item("ScorecardID"), 0) = ScorecardID, " selected=""selected""", "") & ">" & MyCommon.NZ(row.Item("Description"), "") & "</option>")
            Next
          End If
          Send("    </select>")
          Send("  </td>")
          Send("</tr>")
          Send("<tr id=""scdesc""" & IIf(ScorecardID = 0, " style=""display:none;""", "") & ">")
          Send("  <td>")
          Send("    <label for=""ScorecardDesc"">" & Copient.PhraseLib.Lookup("term.scorecardtext", LanguageID) & ":</label>")
          Send("  </td>")
          Send("  <td>")
          Send("    <input type=""text"" class=""medium"" id=""ScorecardDesc"" name=""ScorecardDesc"" maxlength=""31"" value=""" & ScorecardDesc & """" & IIf(RewardsSetScorecard, " disabled=""disabled""", "") & " />")
          MLI.ItemID = l_pgID
          MLI.MLTableName = "PointsProgramTranslations"
          MLI.MLColumnName = "ScorecardDesc"
          MLI.MLIdentifierName = "ProgramID"
          MLI.StandardTableName = "PointsPrograms"
          MLI.StandardColumnName = "ScorecardDesc"
          MLI.StandardIdentifierName = "ProgramID"
          MLI.StandardValue = ScorecardDesc
          MLI.InputName = "ScorecardDesc"
          MLI.InputID = "ScorecardDesc"
          MLI.InputType = "text"
          MLI.MaxLength = 31
          MLI.CSSStyle = "width:230px;"
          MLI.Disabled = RewardsSetScorecard
          Send(Localization.SendTranslationInputs(MyCommon, MLI, Request.QueryString, 2))
          Send("  </td>")
          Send("</tr>")
          'Commenting out bolding functionality; to restore it, uncomment this section
          'Send("<tr>")
          'Send("  <td><label for=""ScorecardBold"">" & Copient.PhraseLib.Lookup("term.bold", LanguageID) & ":</label></td>")
          'Send("  <td><input type=""checkbox"" id=""ScorecardBold"" name=""ScorecardBold"" value=""1""" & IIf(ScorecardBold, " checked=""checked""", "") & " /></td>")
          'Send("</tr>")
        %>
      </table>
    </div>
    <div class="box" id="AdjustmentUPCbox" <% If (Not MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE)) OrElse (MyCommon.Fetch_SystemOption(95) = 0) Then Send(" style=""display:none;""") %>>
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.adjustmentupc", LanguageID))%>
        </span>
      </h2>
      <%
        If Not (RangeBegin = 0 AndAlso RangeEnd = 0 AndAlso MyCommon.Fetch_CPE_SystemOption(102) = 0) Then
          Send("<input type=""text"" id=""AdjustmentUPC"" name=""AdjustmentUPC"" style=""width:200px;"" value=""" & IIf(Not AdjustmentUPC = "", AdjustmentUPC, "") & """ maxlength=""" & MaxLength & """ /><br />")
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
          Send("<label for=""selector"">Top unused codes (double-click to select):</label><br />")
          Send("<select id=""selector"" name=""selector"" size=""10"" style=""width:220px;"" ondblclick=""javascript:selectPLU(" & IDLength & ");"">")
          i = RangeBegin
          counter = 1
          While (counter <= 100) AndAlso (i <= RangeEnd)
            MyCommon.QueryStr = "select CAST(AdjustmentUPC as decimal(26,0)) as AdjustmentUPC from StoredValuePrograms with (NoLock) " & _
                                " where IsNull(AdjustmentUPC, '') <> '' and AdjustmentUPC='" & i.ToString.PadLeft(IDLength, "0") & "' " & _
                                " union " & _
                                "select CAST(AdjustmentUPC as decimal(26,0)) as AdjustmentUPC from PointsPrograms with (NoLock) " & _
                                " where IsNull(AdjustmentUPC, '') <> '' and AdjustmentUPC='" & i.ToString.PadLeft(IDLength, "0") & "' "
            rst = MyCommon.LRT_Select
            If rst.Rows.Count = 0 Then
              Send("  <option value=""" & i & """>" & i.ToString.PadLeft(IDLength, "0") & "</option>")
              counter += 1
            End If
            i += 1
          End While
          Send("</select>")
        End If
      %>
    </div>
    <div class="box" id="Categories" <% If (NOT bEnableCategories) Then Send(" style=""display:none;""") %>>
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.category", LanguageID))%>
        </span>
      </h2>
      <%
        MyCommon.QueryStr = "select OfferCategoryID, Description from OfferCategories with (NoLock) where Deleted=0 order by Description"
        rst2 = MyCommon.LRT_Select()
        Send("<label for=""form_Category"">" & Copient.PhraseLib.Lookup("term.category", LanguageID) & ":</label><br />")
        Send("<select class=""medium"" id=""form_Category"" name=""form_Category"">")
        For Each row2 In rst2.Rows
          If (iCategoryID = row2.Item("OfferCategoryID")) Then
            Sendb("<option value=""" & row2.Item("OfferCategoryID") & """ selected=""selected"">" & row2.Item("Description") & "</option>")
          Else
            Sendb("<option value=""" & row2.Item("OfferCategoryID") & """>" & row2.Item("Description") & "</option>")
          End If
        Next
        Send("</select>")
        Send("<br />")
      %>
    </div>
    <% If bPointsMigrationEnabled Then%>
    <%
      If LastMigrationDate = Nothing Then
        longDateString = Copient.PhraseLib.Lookup("term.never", LanguageID)
        shortDateString = ""
      Else
        longDateString = Logix.ToLongDateTimeString(LastMigrationDate, MyCommon)
        shortDateString = Logix.ToShortDateTimeString(LastMigrationDate, MyCommon)
      End If
    %>
    <div class="box" id="PointsMigration">
      <input type="hidden" id="migration-LastDate" name="migration-LastDate" value="<%sendb(shortDateString) %>" />
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.PointsMigration", LanguageID))%>
        </span>
      </h2>
      <span>
        <%
          Sendb(Copient.PhraseLib.Lookup("term.enter-datetime", LanguageID))
        %>
      </span>
      <br class="half" />
      <br class="half" />
      <label for="migration-date">
        <% Sendb(Copient.PhraseLib.Lookup("term.date", LanguageID))%>:</label><br />
      <input class="short" id="migration-date" name="migration-date" maxlength="10" type="text"
        value="<% If Not String.IsNullOrEmpty(MigrationDate) Then sendb(Logix.ToShortDateString(MigrationDate, MyCommon)) %>" />
      <img src="../images/calendar.png" class="calendar" id="migration-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
        title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('migration-date', event);" />
      <input class="shortest" id="migration-hr" maxlength="2" name="migration-hr" type="text"
        value="<% sendb(MigrationHr)%>" />:
      <input class="shortest" id="migration-min" maxlength="2" name="migration-min" type="text"
        value="<% sendb(MigrationMin)%>" />
      <br />
      <br class="half" />
      <label for="migration-id">
        <%Send(Copient.PhraseLib.Lookup("term.pointsprogram", LanguageID))%>:</label><br />
      <select id="migration-id" name="migration-id" class="longest">
        <%
          MyCommon.QueryStr = "select 0 as ProgramID, 'None' as ProgramName union select ProgramID, ProgramName from PointsPrograms with (NoLock) where deleted=0 and ProgramId <> " & l_pgID & ";"
          rst2 = MyCommon.LRT_Select
          For Each row2 In rst2.Rows
            Send("  <option value=""" & MyCommon.NZ(row2.Item("ProgramID"), 0) & """" & IIf(lMigrationProgramID = MyCommon.NZ(row2.Item("ProgramID"), 0), " selected=""selected""", "") & ">" & _
                  "" & MyCommon.NZ(row2.Item("ProgramName"), "Unknown"))
          Next
        %>
      </select>
      <%
        Sendb(Copient.PhraseLib.Lookup("point-edit.LastMigration", LanguageID) & ": ")
        Send(longDateString)
      %>
    </div>
    <%
      If Request.Browser.Type = "IE6" Then
        Send("<iframe src=""javascript:'';"" id=""calendariframe"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no"" style=""visibility:hidden;display:none;""></iframe>")
      End If
    %>
    <% End If%>
    <div id="datepicker" class="dpDiv">
    </div>
    <% If bPointsDeletionEnabled Then%>
    <%
      If LastDeletionDate = Nothing Then
        longDateString = Copient.PhraseLib.Lookup("term.never", LanguageID)
        shortDateString = ""
      Else
        longDateString = Logix.ToLongDateTimeString(LastDeletionDate, MyCommon)
        shortDateString = Logix.ToShortDateTimeString(LastDeletionDate, MyCommon)
      End If
    %>
    <div class="box" id="PointsDeletion">
      <input type="hidden" id="deletion-LastDate" name="deletion-LastDate" value="<%sendb(shortDateString) %>" />
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.PointsDeletion", LanguageID))%>
        </span>
      </h2>
      <%
        Send("<br class=""half"" />")
        Send("<input type=""checkbox"" id=""enabledeletedate"" name=""enabledeletedate"" value=""1""" & IIf(bEnableDeleteDate, " checked=""checked""", "") & "onclick=""javascript:handleEnableDeleteDate_Click(1, this.checked);""  /><label for=""enabledeletedate"">" & Copient.PhraseLib.Lookup("point-edit.enabledeletion", LanguageID) & "</label><br />")
        Send("<br />")
      %>
      <span>
        <%
          Sendb(Copient.PhraseLib.Lookup("term.enter-datetime", LanguageID))
        %>
      </span>
      <br class="half" />
      <br class="half" />
      <label for="deletion-date">
        <% Sendb(Copient.PhraseLib.Lookup("term.date", LanguageID))%>:</label><br />
      <input class="short" id="deletion-date" name="deletion-date" maxlength="10" type="text"
        <%
         If Not bEnableDeleteDate Then
           Send(" disabled=""disabled""")
         End If 
        %> value="<% If Not String.IsNullOrEmpty(DeletionDate) Then sendb(Logix.ToShortDateString(DeletionDate, MyCommon)) %>" />
      <img src="../images/calendar.png" class="calendar" id="deletion-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
        title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('deletion-date', event);" />
      <input class="shortest" id="deletion-hr" maxlength="2" name="deletion-hr" type="text"
        <%
         If Not bEnableDeleteDate Then
           Send(" disabled=""disabled""")
         End If 
        %> value="<% sendb(DeletionHr)%>" />:
      <input class="shortest" id="deletion-min" maxlength="2" name="deletion-min" type="text"
        <%
         If Not bEnableDeleteDate Then
           Send(" disabled=""disabled""")
         End If 
        %> value="<% sendb(DeletionMin)%>" />
      <br />
      <br class="half" />
      <%
        Sendb(Copient.PhraseLib.Lookup("point-edit.LastDeletion", LanguageID) & ": ")
        Send(longDateString)
      %>
    </div>
    <%
      If Request.Browser.Type = "IE6" Then
        Send("<iframe src=""javascript:'';"" id=""calendariframe"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no"" style=""visibility:hidden;display:none;""></iframe>")
      End If
    %>
    <% End If%>
    <div class="box" id="lastuploadattempt" <% if(l_pgID=0)then send(" style=""visibility: hidden;""") %>>
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("pgroup-edit-lastuploaded", LanguageID))%>
        </span>
      </h2>
      <br class="half" />
      <%
        ' last update date
        Sendb(Copient.PhraseLib.Lookup("term.lastupload", LanguageID) & ": ")
        If (LastUpload Is Nothing) OrElse (LastUpload = "1/1/1900") Then
          If (l_pgID <> 0) Then
            Sendb(Copient.PhraseLib.Lookup("term.neveruploaded", LanguageID))
            Send("<br />")
          End If
        Else
          longDate = MyCommon.NZ(LastUpload, "1/1/1900")
          longDateString = Logix.ToLongDateTimeString(longDate, MyCommon)
          Send(longDateString)
          Send("<br />")
        End If

        ' last update message
        If LastUploadMsg IsNot Nothing AndAlso LastUploadMsg.Trim <> "" Then
          Sendb(Copient.PhraseLib.Lookup("term.statusmessage", LanguageID) & ": ")
          If LastUploadMsg.ToLower.IndexOf("upload processing completed") > -1 Then
            Sendb(Copient.PhraseLib.Lookup("term.successful", LanguageID))
          Else
            Sendb("<br /><p style=""margin:2px 10px;font-family:courier;font-size:11px;"">" & LastUploadMsg.Trim & "</p>")
          End If
        End If

      %>
      <hr class="hidden" />
    </div>
    <%--
      <div class="box" id="addcustomers"<% if(l_pgid=0)then send(" style=""display:none;""") %>>
        <h2><span><% Sendb(Copient.PhraseLib.Lookup("points-edit.addremove", LanguageID))%></span></h2>
        <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.customers", LanguageID))%>">
            <tr>
              <td style="width: 100px;"><% Sendb(Copient.PhraseLib.Lookup("term.customerid", LanguageID))%>:</td>
              <td colspan="2">
                <input class="small" id="clientuserid1" maxlength="19" name="clientuserid1" type="text" value="" />
              </td>
            </tr>
            <tr>
              <td><% Sendb(Copient.PhraseLib.Lookup("term.balance", LanguageID))%>:</td>
              <td>
                <input class="small" id="balance" name="balance" type="text" value="" />
              </td>
              <%
                If (Logix.UserRoles.CRUDPoints) Then
                  Send("<td><input class=""add"" id=""add"" name=""add"" type=""submit"" value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """ /></td>")
                Else
                  Send("<td></td>")
                End If
              %>
            </tr>
            <tr>
              <td colspan="3"><%Sendb(Copient.PhraseLib.Lookup("points-edit.listnote", LanguageID))%></td>
            </tr>
            <tr>
              <td colspan="3">
                <select class="longest" style="font-family: monospace;" id="custbalID" name="custbalID" size="10">
                  <%
                    MyCommon.QueryStr = "select top 100 P.PKID, C.PrimaryExtID, P.Amount from Points P " & _
                                        "inner join Customers C on P.CustomerPK = C.CustomerPK where P.PromoVarID =" & pgPromoVarID & _
                                        " order by C.PrimaryExtID;"
                    rstPoints = MyCommon.LXS_Select
                    If (rstPoints.Rows.Count > 0) Then
                      OptionText = "Customer Ext ID".PadRight(19, " ") & "  " & "Balance".PadLeft(7, " ")
                      OptionText = OptionText.Replace(" ", "&nbsp;")
                      Send("  <option value=""0"">" & OptionText & "</option>")
                      Send("  <option value=""0"">-------------------&nbsp;&nbsp;-------</option>")
                      For Each row In rstPoints.Rows
                        OptionText = MyCommon.NZ(row.Item("PrimaryExtID"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)).ToString.PadRight(19, " ") & "  " & _
                                      FormatNumber(MyCommon.NZ(row.Item("Amount"), 0), 0, TriState.False, TriState.False, TriState.False).PadLeft(7, " ")
                        OptionText = OptionText.Replace(" ", "&nbsp;")
                        Send("  <option value=""" & row.Item("PKID") & """>" & OptionText & "</option>")
                      Next
                    End If
                  %>
                </select>
              </td>
            </tr>
          <%
            If (Logix.UserRoles.CRUDPoints) Then
              Send("<tr>")
              Send("  <td colspan=""3"">")
              Send("    <input class=""remove"" id=""remove"" name=""remove"" type=""submit"" value=""" & Copient.PhraseLib.Lookup("term.remove", LanguageID) & """ />")
              Send("  </td>")
              Send("</tr>")
            End If
          %>
        </table>
        <hr class="hidden" />
      </div>
    --%>
  </div>
  <div id="gutter">
  </div>
  <div id="column2">
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
          onchange='javascript:updateInputVal(this, "hdnAllowAnyCustomerUE");' <% if(AllowAnyCustomer_UE)then sendb(" checked=""checked""") %>
          <% if(IsAnyCustomerOffersExist_UE)then sendb(" disabled=""disabled""") %> />
        <label for="cbAllowAnyCustomerUE">
          <%Send(Copient.PhraseLib.Lookup("term.allowanycustomerearnredeem", LanguageID))%></label>
        <% End If%>
          <br />
          <input type="checkbox" name="VisibleToCustomers" id="VisibleToCustomers" value="1"
        <%Send(IIf(VisibleToCustomers, "checked=""checked"" ", ""))%> />
      <label for="VisibleToCustomers">
        <%Send(Copient.PhraseLib.Lookup("programs.includeforcustomerportal", LanguageID))%></label>
        <br class="half" />

        <hr class="hidden" />
      </div>
      <br />
      <br class="half" />
      <hr class="hidden" />
    </div>
    <% End If%>
    <% If (l_pgID > 0) Then%>
    <div class="box" id="eligibleoffers" <% if (l_pgID=0) then sendb(" style=""visibility: hidden;""") %>>
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.associatedeligibleoffers", LanguageID))%>
        </span>
      </h2>
      <div class="boxscrollhalf">
        <% 
          If (l_pgID <> 0) Then
            Dim lstOffers As List(Of CMS.AMS.Models.Offer)
            lstOffers = m_Offer.GetEligibleOffersByPointsProgramID(l_pgID)
            If (bEnableRestrictedAccessToUEOfferBuilder) Then
              lstOffers = GetRoleBasedUEOffers(lstOffers, MyCommon, Logix)
            End If
            If lstOffers.Count > 0 Then
              For Each offer As CMS.AMS.Models.Offer In lstOffers
                If (MyCommon.Fetch_UE_SystemOption(168) = "1" AndAlso MyCommon.NZ(offer.BuyerID, "") <> "") Then
                  offer.OfferName = "Buyer " + offer.BuyerID.ToString() + " -" + MyCommon.NZ(offer.OfferName, "").ToString()
                Else
                  offer.OfferName = MyCommon.NZ(offer.OfferName, Copient.PhraseLib.Lookup("term.unknown", LanguageID))
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
    <div class="box" id="offers">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.associatedoffers", LanguageID))%>
        </span>
      </h2>
      <div class="boxscroll">
        <%
        
          If (bEnableRestrictedAccessToUEOfferBuilder) Then
            conditionalQuery = GetRestrictedAccessToUEBuilderQuery(MyCommon, Logix, "I")
          End If
            
          MyCommon.QueryStr = "SELECT DISTINCT 1 as EngineID, O.Name, O.OfferID,NULL as BuyerID " & _
                                  "FROM PointsPrograms AS PP WITH (NoLock) " & _
                                  "INNER JOIN OfferConditions AS OC WITH (NoLock) " & _
                                  "ON PP.ProgramID = OC.LinkID " & _
                                  "AND PP.Deleted = 0 " & _
                                  "AND OC.ConditionTypeID = 3 " & _
                                  "INNER JOIN Offers AS O WITH (NoLock) " & _
                                  "ON OC.OfferID = O.OfferID " & _
                                  "AND O.Deleted = 0 And O.IsTemplate=0 " & _
                                  "AND OC.Deleted = 0 " & _
                                  "WHERE PP.ProgramID = " & l_pgID & " " & _
                                  "UNION " & _
                                  "SELECT DISTINCT 1 as EngineID, O.Name, O.OfferID,NULL as BuyerID " & _
                                  "FROM RewardPoints AS RP WITH (NoLock) " & _
                                  "INNER JOIN OfferRewards AS OFFR WITH (NoLock) " & _
                                  "ON RP.RewardPointsID = OFFR.LinkID " & _
                                  "AND (OFFR.RewardTypeID = 2 OR OFFR.RewardTypeID = 13) " & _
                                  "INNER JOIN Offers AS O WITH (NoLock) " & _
                                  "ON OFFR.OfferID = O.OfferID " & _
                                  "AND O.Deleted = 0 And O.IsTemplate=0 " & _
                                  "AND OFFR.Deleted = 0 " & _
                                  "AND O.ProdEndDate >= getdate() " & _
                                  "WHERE RP.ProgramID = " & l_pgID & " " & _
                                  "UNION " & _
                                  "SELECT DISTINCT 2 as EngineID, I.IncentiveName as Name, I.IncentiveID as OfferID,buy.ExternalBuyerId as BuyerID from CPE_IncentivePointsGroups IPG with (NoLock) " & _
                                  "INNER JOIN CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID=IPG.RewardOptionID " & _
                                  "INNER JOIN CPE_Incentives I with (NoLock) on I.IncentiveID = RO.IncentiveID " & _
                                  "INNER JOIN PointsPrograms PP with (NoLock) on IPG.ProgramID = PP.ProgramID " & _
                                  "left outer join Buyers as buy with (nolock) on buy.BuyerId= I.BuyerId " & _
                                  "WHERE IPG.ProgramID = " & l_pgID & " and IPG.Deleted = 0 and RO.Deleted=0 and I.Deleted=0 and PP.Deleted=0 and I.IsTemplate=0 "
          If (bEnableRestrictedAccessToUEOfferBuilder AndAlso Not String.IsNullOrEmpty(conditionalQuery)) Then MyCommon.QueryStr &= conditionalQuery & " "
          MyCommon.QueryStr &= "UNION " & _
                                 "SELECT DISTINCT 2 as EngineID, I.IncentiveName as Name, I.IncentiveID as OfferID,buy.ExternalBuyerId as BuyerID from CPE_DeliverablePoints DP " & _
                                 "INNER JOIN CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID=DP.RewardOptionID " & _
                                 "INNER JOIN CPE_Incentives I with (NoLock) on I.IncentiveID = RO.IncentiveID " & _
                                 "left outer join Buyers as buy with (nolock) on buy.BuyerId= I.BuyerId " & _
                                 "WHERE DP.Deleted=0 and RO.Deleted=0 and I.Deleted=0 and I.IsTemplate=0 and ProgramID=" & l_pgID
          If (bEnableRestrictedAccessToUEOfferBuilder AndAlso Not String.IsNullOrEmpty(conditionalQuery)) Then MyCommon.QueryStr &= conditionalQuery & " "
          MyCommon.QueryStr &= "ORDER BY Name;"
          dstAssociated = MyCommon.LRT_Select
          rowCount = dstAssociated.Rows.Count
          If rowCount > 0 Then
            For Each row In dstAssociated.Rows
              If (MyCommon.Fetch_UE_SystemOption(168) = "1" AndAlso MyCommon.NZ(row.Item("BuyerID"), "") <> "") Then
                assocName = "Buyer " + MyCommon.NZ(row.Item("BuyerID"), "").ToString() + " - " + MyCommon.NZ(row.Item("Name"), "").ToString()
              Else
                assocName = MyCommon.NZ(row.Item("Name"), "").ToString()
              End If
              'assocName = MyCommon.NZ(row.Item("Name"), "")
              assocID = MyCommon.NZ(row.Item("OfferID"), 0)
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
    <% If MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE) Then%>
    <div class="box" id="advancedOptions">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.advancedoptions", LanguageID))%>
        </span>
      </h2>
      <label for="returnsHandling">
        <%Send(Copient.PhraseLib.Lookup("programs.returnhandling", LanguageID))%>:</label><br />
      <select id="returnsHandling" name="returnsHandling" class="longest">
        <%
          MyCommon.QueryStr = "select ReturnHandlingTypeID, Name, PhraseID from UE_ReturnHandlingTypes with (NoLock);"
          rst2 = MyCommon.LRT_Select
          For Each row2 In rst2.Rows
            Send("  <option value=""" & MyCommon.NZ(row2.Item("ReturnHandlingTypeID"), 0) & """" & IIf(ReturnHandlingTypeID = MyCommon.NZ(row2.Item("ReturnHandlingTypeID"), 0), " selected=""selected""", "") & ">" & _
                  "" & Copient.PhraseLib.Lookup(MyCommon.NZ(row2.Item("PhraseID"), 0), LanguageID, MyCommon.NZ(row2.Item("Name"), "Unknown")))
          Next
        %>
      </select>
      <br />
      <br class="half" />
      <input type="checkbox" name="disallowRedeemInTrans" id="disallowRedeemInTrans" value="1"
        <%Send(IIf(DisallowRedeemInTrans, "checked=""checked"" ", "")) %> />
      <label for="disallowRedeemInTrans">
        <%Send(Copient.PhraseLib.Lookup("programs.disallowredeemintrans", LanguageID))%></label><br />
      <input type="checkbox" name="allowNegBal" id="allowNegBal" value="1" <%Send(IIf(AllowNegativeBal, "checked=""checked"" ", "")) %> />
      <label for="allowNegBal">
        <%Send(Copient.PhraseLib.Lookup("programs.allowtogonegative", LanguageID))%></label><br />
      <%If bPendingEnabled Then%>
        <input type="checkbox" name="applyRedeemedPending" id="applyRedeemedPending" value="1" <%Send(IIf(bApplyEarnedPending, "", "checked=""checked"" ")) %> />
        <label for="applyRedeemedPending">
          <%Send(Copient.PhraseLib.Lookup("term.applypending", LanguageID))%></label><br />
      <%End If%>
        
      <br class="half" />
      <hr class="hidden" />
    </div>
    <% End If%>
  </div>
  <%
  ElseIf (Request.QueryString("New") = Copient.PhraseLib.Lookup("term.new", LanguageID)) OrElse (Request.QueryString("ProgramGroupID") = Copient.PhraseLib.Lookup("term.new", LanguageID).ToLower) OrElse (Request.QueryString("ProgramGroupID") = "0") OrElse (Request.QueryString("ProgramGroupID") = "") Then
  %>
  <div id="column1">
    <div class="box" id="newIdentity">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.identification", LanguageID))%>
        </span>
      </h2>
      <%
        If (l_pgID <> 0) Then
          Send(Copient.PhraseLib.Lookup("term.id", LanguageID) & ": " & l_pgID)
          Send("<br />")
        End If
        Send("<input type=""hidden"" id=""ProgramGroupID"" name=""ProgramGroupID"" value=""new"" />")
        Send("<br class=""half"" />")
        Send("<label for=""name"">" & Copient.PhraseLib.Lookup("term.name", LanguageID) & ": </label><br />")
        Send("<input type=""text"" class=""longest"" id=""name"" name=""name"" maxlength=""50"" /><br />")
        Send("<label for=""desc"">" & Copient.PhraseLib.Lookup("term.description", LanguageID) & ": </label><br />")
        Send("<textarea class=""longest"" id=""desc"" name=""desc"" cols=""48"" rows=""3"" maxlength=""1000""></textarea><br />")
        Send("<br class=""half"" />")
        Send("<small>" & Copient.PhraseLib.Lookup("CPEoffergen.description", LanguageID) & "</small><br class=""half"" />")
        If (CAMInstalled) Then
          Send("<input type=""checkbox"" id=""CAMProgram"" name=""CAMProgram"" value=""1"" />")
          Send("<label for=""CAMProgram"">" & Copient.PhraseLib.Lookup("term.camprogram", LanguageID) & "</label>")
          Send("<br class=""half"" />")
          Send("<br />")
        End If
        
        If Logix.UserRoles.AssignPromoVarIdToPointsPrograms And lMinimumAutoGeneratedPromoVarID > 0 Then
          Send("<br class=""half"" />")
          Send("<label for=""promovarid"">" & Copient.PhraseLib.Lookup("term.promovarid", LanguageID) & ": </label><br />")
          Send("<input type=""text"" class=""longest"" id=""promovarid"" name=""promovarid"" maxlength=""8"" /><br />")
          Send("<small>" & Copient.PhraseLib.Lookup("promovar.manual-info", LanguageID) & "</small><br class=""half"" />")
        End If
          
        If MyCommon.Fetch_SystemOption(80) = 0 Then
          'No external points integration
          Send("<input type=""hidden"" id=""exthosttypeid"" name=""exthosttypeid"" value=""0"" />")
        ElseIf MyCommon.Fetch_SystemOption(80) > 0 Then
          'There is an external points integration
          MyCommon.QueryStr = "select PhraseID, Description from SystemOptionValues with (NoLock) where OptionID=80 and OptionValue=" & MyCommon.Fetch_SystemOption(80) & ";"
          rst2 = MyCommon.LRT_Select
          If rst2.Rows.Count > 0 Then
            If MyCommon.NZ(rst2.Rows(0).Item("PhraseID"), 0) > 0 Then
              ExtHostDesc = Copient.PhraseLib.Lookup(rst2.Rows(0).Item("PhraseID"), LanguageID)
            Else
              ExtHostDesc = MyCommon.NZ(rst2.Rows(0).Item("Description"), "")
            End If
          End If
          Send("<br class=""half"" />")
          Send("<input type=""checkbox"" id=""isexternal"" name=""isexternal"" value=""1""" & IIf(IsExternalProgram, "checked=""checked""", "") & " onclick=""javascript:handleExtProg_Click(1, this.checked);"" />")
          Send("<label for=""isexternal"">" & Copient.PhraseLib.Lookup("term.externalprogram", LanguageID) & " (" & ExtHostDesc & ")</label><br />")
          Send("<br class=""half"" />")
          MyCommon.QueryStr = "select ExtHostTypeID, Description, AllowDecimalValues from ExtHostTypes with (NoLock) where ExtPointsOptionValue=" & MyCommon.Fetch_SystemOption(80) & ";"
          rst2 = MyCommon.LRT_Select
          If rst2.Rows.Count = 0 Then
            Send("<input type=""hidden"" id=""exthosttypeid"" name=""exthosttypeid"" value=""0"" />")
          ElseIf rst2.Rows.Count = 1 Then
            Send("<input type=""hidden"" id=""exthosttypeid"" name=""exthosttypeid"" value=""" & MyCommon.NZ(rst2.Rows(0).Item("ExtHostTypeID"), 0) & """ />")
          Else
            Send("<label for=""exthosttypeid"">" & Copient.PhraseLib.Lookup("term.host", LanguageID) & ":</label><br />")
            Send("<select id=""exthosttypeid"" name=""exthosttypeid"">")
            For Each row In rst2.Rows
              Send("  <option value=""" & MyCommon.NZ(row.Item("ExtHostTypeID"), 0) & """>" & MyCommon.NZ(row.Item("Description"), "&nbsp;") & "</option>")
            Next
            Send("</select><br />")
            Send("<br class=""half"" />")
          End If
          Send("<label for=""externalid"">" & Copient.PhraseLib.Lookup("term.externalid", LanguageID) & ":</label><br />")
          Send("<input type=""text"" class=""longest"" id=""externalid"" name=""externalid"" maxlength=""100"" value="""" /><br />")
          If rst2.Rows.Count > 0 Then
            'Decimal values limited to only the Excentus/CRM external type
            If MyCommon.NZ(rst2.Rows(0).Item("AllowDecimalValues"), False) AndAlso MyCommon.NZ(rst2.Rows(0).Item("ExtHostTypeID"), 0) = 2 Then
              Send("<br class=""half"" />")
              Send("<input type=""checkbox"" id=""decimalvalues"" name=""decimalvalues"" value=""1"" disabled=""disabled""" & IIf(DecimalValues, " checked=""checked""", "") & " />")
              Send("<label for=""decimalvalues"">" & Copient.PhraseLib.Lookup("term.decimalvalues", LanguageID) & "</label><br />")
            End If
          End If
        End If
      %>
      <hr class="hidden" />
    </div>
  </div>
  <%
  Else
  %>
  <div id="column1">
    <div class="box" id="identification">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.identification", LanguageID))%>
        </span>
      </h2>
      <% Sendb("ID " & l_pgID & " " & Copient.PhraseLib.Lookup("term.notfound", LanguageID))%>
    </div>
    <br clear="all" />
  </div>
  <% End If%>
</div>
<!-- End main -->
</form>
<div id="uploader" style="display: none;">
  <div id="uploadwrap">
    <div class="box" id="uploadbox">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.upload", LanguageID))%>
        </span>
      </h2>
      <form action="point-edit.aspx" id="uploadform" name="uploadform" method="post" enctype="multipart/form-data">
      <%
        Sendb("<input type=""button"" class=""ex"" id=""uploaderclose"" name=""uploaderclose"" value=""X"" style=""float:right;position:relative;top:-27px;"" ")
        Sendb("onclick=""javascript:document.getElementById('uploader').style.display='none';"" />")
      %>
      <% Sendb(Copient.PhraseLib.Lookup("point-edit.upload", LanguageID))%>
      <br />
      <br class="half" />
      <%
        If (Logix.UserRoles.EditPointsBalances) Then
          Send("     <input type=""hidden"" name=""ProgramID"" value=""" & l_pgID & """ />")
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
</div>
<script runat="server">
  Public Function AllDigits(ByVal txt As String) As Boolean
    Dim ch As String
    Dim i As Integer
    
    AllDigits = True
    If Len(txt) > 0 Then
      For i = 1 To Len(txt)
        ' See if the next character is a non-digit.
        ch = Mid$(txt, i, 1)
        If ch < "0" OrElse ch > "9" Then
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
</script>
<%
    If MyCommon.Fetch_SystemOption(75) Then
        If (l_pgID > 0 AndAlso Logix.UserRoles.AccessNotes) Then
            Send_Notes(8, l_pgID, AdminUserID)
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
