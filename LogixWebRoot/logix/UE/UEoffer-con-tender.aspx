<%@ Page Language="vb" Debug="true" CodeFile="ue-cb.vb" Inherits="UECB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
    ' *****************************************************************************
    ' * FILENAME: UEoffer-con-tender.aspx 
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

    Dim MyCommon As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    Dim rst, dt, TierDT As DataTable
    Dim row As DataRow
    Dim OfferID As Long
    Dim Name As String = ""
    Dim isTemplate As Boolean
    Dim FromTemplate As Boolean
    Dim Disallow_Edit As Boolean = True
    Dim DisabledAttribute As String = ""
    Dim roid As Integer
    Dim ConditionID As String
    Dim Ids() As String
    Dim i As Integer
    Dim historyString As String = ""
    Dim CloseAfterSave As Boolean = False
    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False
    Dim RequirePP As Boolean = False
    Dim HasRequiredPP As Boolean = False
    Dim tmpString As String = ""
    Dim EngineID As Integer = 2
    Dim BannersEnabled As Boolean = True
    Dim IncentiveTenderID As Integer
    Dim TenderValue As Decimal
    Dim TenderType As Integer
    Dim TierTenderVal As Decimal
    Dim ExcludedTender As Boolean = False
    Dim ExcludedValue As Decimal = 0D
    Dim SelGrp As String = ""
    Dim TierLevel, t As Integer
    Dim ValidTier As Boolean = False
    Dim ValidDecimalPlaces As Boolean = False
    Dim t1, t2 As Decimal
    Dim Localizer As Copient.Localization
    Dim CurSymbol As String = ""
    Dim CurAbbr As String = ""
    Dim CurPrecision As Integer = 2
    Dim tempTenderValue As Decimal
    Dim SupportGlobalAndTieredConditions As Integer = MyCommon.Fetch_UE_SystemOption(197)
    Dim UseSameTierValue As Integer = 0
    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If

    Response.Expires = 0
    MyCommon.AppName = "UEoffer-con-tender.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)

    BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")

    OfferID = Request.QueryString("OfferID")

    'This would redirect to summary page if offer is awaiting deployment or if offer is awaiting recommendations.
    CheckIfValidOffer(MyCommon, OfferID)

    ConditionID = Request.QueryString("ConditionID")

    Dim isTranslatedOffer As Boolean =MyCommon.IsTranslatedUEOffer(OfferID,  MyCommon)
    Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249)="1",True,False)

    Dim bEnableAdditionalLockoutRestrictionsOnOffers As Boolean = IIf(MyCommon.Fetch_SystemOption(260) = "1", True, False)
    Dim bOfferEditable As Boolean = MyCommon.IsOfferEditablePastLockOutPeriod(Logix.UserRoles.EditOfferPastLockoutPeriod, MyCommon, OfferID)

    If (Request.QueryString("EngineID") <> "") Then
        EngineID = MyCommon.Extract_Val(Request.QueryString("EngineID"))
    Else
        MyCommon.QueryStr = "select EngineID from OfferIDs where OfferID=" & OfferID & ";"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
            EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), 0)
        End If
    End If

    MyCommon.QueryStr = "select RewardOptionID, ExcludedTender, ExcludedTenderAmtRequired, TierLevels from CPE_RewardOptions with (NoLock) where IncentiveID=" & OfferID & " and TouchResponse=0 and Deleted=0;"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
        roid = MyCommon.NZ(rst.Rows(0).Item("RewardOptionID"), -1)
        ExcludedTender = MyCommon.NZ(rst.Rows(0).Item("ExcludedTender"), False)
        ExcludedValue = MyCommon.NZ(rst.Rows(0).Item("ExcludedTenderAmtRequired"), 0)
        TierLevel = MyCommon.NZ(rst.Rows(0).Item("TierLevels"), 0)
    End If
    If Request.QueryString("ExcludedTender") <> "" Then
        If MyCommon.Extract_Val(Request.QueryString("ExcludedTender")) = 1 Then
            ExcludedTender = True
        ElseIf MyCommon.Extract_Val(Request.QueryString("ExcludedTender")) = 0 Then
            ExcludedTender = False
        End If
    ElseIf MyCommon.Extract_Val(Request.QueryString("excluded")) <> "" Then
        If MyCommon.Extract_Val(Request.QueryString("excluded")) = 1 Then ExcludedTender = True
    End If

    ' get the localized symbols, abbreviations, and precision for the offer's currency
    Localizer = New Copient.Localization(MyCommon)
    CurSymbol = Localizer.Get_Currency_Symbol(roid)
    CurAbbr = Localizer.Get_Currency_Abbreviation(roid)
    CurPrecision = Localizer.Get_Currency_Precision(roid)

    'Get the TenderTypeIDs of the selected tenders
    If Request.QueryString("selGroups") <> "" Then
        SelGrp = Request.QueryString("selGroups")
        TenderType = MyCommon.Extract_Val(Request.QueryString("selGroups"))
    Else
        MyCommon.QueryStr = "select TenderTypeID from CPE_IncentiveTenderTypes where RewardOptionID=" & roid
        dt = MyCommon.LRT_Select()
        If dt.Rows.Count > 0 Then
            For t = 1 To dt.Rows.Count
                If dt.Rows.Count = 1 Then
                    SelGrp = dt.Rows(t - 1).Item("TenderTypeID")
                ElseIf t = dt.Rows.Count Then
                    SelGrp = SelGrp & dt.Rows(t - 1).Item("TenderTypeID")
                Else
                    SelGrp = SelGrp & dt.Rows(t - 1).Item("TenderTypeID") & ","
                End If
            Next
        Else
            SelGrp = 0
        End If
    End If

    If Request.QueryString("IncentiveTenderID") <> "" Then
        IncentiveTenderID = MyCommon.Extract_Val(Request.QueryString("IncentiveTenderID"))
        MyCommon.QueryStr = "select TenderTypeID from CPE_IncentiveTenderTypes with (NoLock) where RewardOptionID=" & roid & " and IncentiveTenderID=" & IncentiveTenderID
        rst = MyCommon.LRT_Select()
        If rst.Rows.Count > 0 Then
            TenderType = MyCommon.NZ(rst.Rows(0).Item("TenderTypeID"), 0)
        End If
        'ElseIf Request.QueryString("tID") <> "" Then
        '  IncentiveTenderID = MyCommon.Extract_Val(Request.QueryString("tID"))
    Else
        IncentiveTenderID = 0
        TenderType = 0
    End If


    'Get UseSameTierValue
    If TierLevel > 1 And SupportGlobalAndTieredConditions = 1 Then
        MyCommon.QueryStr = "select ITTT.TierLevel, ITTT.Value from CPE_IncentiveTenderTypeTiers As ITTT with (NoLock) left join CPE_IncentiveTenderTypes as ITT with (NoLock) " & _
                            "on ITTT.IncentiveTenderID=ITT.IncentiveTenderID where ITT.IncentiveTenderID=" & IncentiveTenderID & " and ITT.Deleted=0;"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
            Dim lastValue As Decimal = 0.0
            For Each row In rst.Rows
                If (MyCommon.NZ(row.Item("TierLevel"), 0) > 1) and MyCommon.NZ(row.Item("Value"), 0) <> lastValue Then
                    UseSameTierValue = 0
                    Exit For
                Else
                    lastValue = MyCommon.NZ(row.Item("Value"), 0)
                End If
            Next
            If MyCommon.NZ(row.Item("TierLevel"), 0) = TierLevel Then
                UseSameTierValue = 1
            End If
        End If
    Else
        UseSameTierValue = 0
    End If

    ' see if someone is saving
    If (Request.QueryString("save") <> "" And roid > 0) Then
        'Tier level validation code
        If TierLevel > 1 And Not ExcludedTender Then
            For t = 2 To TierLevel
                t2 = MyCommon.Extract_Val(Request.QueryString("t" & t & "_tVal"))
                t1 = MyCommon.Extract_Val(Request.QueryString("t" & t - 1 & "_tVal"))
                If t2 > t1 Then
                    ValidTier = True
                Else
                    ValidTier = False
                    Exit For
                End If
            Next
        Else
            ValidTier = True
        End If

        ' validate entry is not more precise than the currencies for the offer
        For t = 1 To TierLevel
            t1 = MyCommon.Extract_Decimal(Request.QueryString("t" & t & "_tVal"), MyCommon.GetAdminUser.Culture)

            ValidDecimalPlaces = (GetDecimalPlaces(t1) <= CurPrecision)
            If Not ValidDecimalPlaces Then Exit For
        Next

        If ValidTier AndAlso ValidDecimalPlaces Then
            ' check to see if a tender condition is required by the template, if applicable
            MyCommon.QueryStr = "select TenderTypeID from CPE_IncentiveTenderTypes with (NoLock) where RewardOptionID=" & roid & _
                                " and RequiredFromTemplate=1 and Deleted=0;"
            rst = MyCommon.LRT_Select
            HasRequiredPP = (rst.Rows.Count > 0)
            If (Request.QueryString("selGroups") <> "") Then
                historyString = "Altered tender condition types: TenderID " & Request.QueryString("selGroups")
                If (Request.QueryString("t1_tVal") <> "" AndAlso Request.QueryString("ttID") <> "") Then
                    'IncentiveTenderID = MyCommon.Extract_Val(Request.QueryString("tID"))
                    TenderValue = MyCommon.Extract_Decimal(Request.QueryString("t1_tVal"), MyCommon.GetAdminUser.Culture)
                    TenderType = MyCommon.Extract_Val(Request.QueryString("ttID"))
                    If TenderValue > 0 Then
                        If TenderValue >= 1000000 Then
                            infoMessage = Copient.PhraseLib.Lookup("ueoffer-con-tender.LessThan1Million", LanguageID)
                        Else
                            'delete tier records for this tender if there are any
                            If IncentiveTenderID > 0 Then
                                MyCommon.QueryStr = "delete from CPE_IncentiveTenderTypeTiers where RewardOptionID=" & roid & " and IncentiveTenderID=" & IncentiveTenderID
                                MyCommon.LRT_Execute()
                            End If
                            'Add Tender or update if it is already there.
                            If IncentiveTenderID > 0 Then
                                MyCommon.QueryStr = "update CPE_IncentiveTenderTypes set TenderTypeID=" & TenderType & ", Value=" & Math.Round(TenderValue, CurPrecision) & ", LastUpdate=getdate() where IncentiveTenderID=" & IncentiveTenderID
                                MyCommon.LRT_Execute()
                            Else
                                MyCommon.QueryStr = "dbo.pa_CPE_AddTenderType"
                                MyCommon.Open_LRTsp()
                                MyCommon.LRTsp.Parameters.Add("@ROID", SqlDbType.Int, 4).Value = roid
                                MyCommon.LRTsp.Parameters.Add("@TenderTypeID", SqlDbType.Int, 4).Value = TenderType
                                MyCommon.LRTsp.Parameters.Add("@Value", SqlDbType.Decimal, 15).Value = Math.Round(TenderValue, CurPrecision)
                                MyCommon.LRTsp.Parameters.Add("@IncentiveTenderID", SqlDbType.Int, 4).Direction = ParameterDirection.Output
                                MyCommon.LRTsp.ExecuteNonQuery()
                                IncentiveTenderID = MyCommon.LRTsp.Parameters("@IncentiveTenderID").Value
                                MyCommon.Close_LRTsp()
                            End If
                            'Write Tiers
                            For t = 1 To TierLevel
                                TierTenderVal = MyCommon.Extract_Decimal(Request.QueryString("t" & t & "_tVal"), MyCommon.GetAdminUser.Culture)
                                MyCommon.QueryStr = "dbo.pa_CPE_AddTenderTypeTiers"
                                MyCommon.Open_LRTsp()
                                MyCommon.LRTsp.Parameters.Add("@IncentiveTenderID", SqlDbType.Int, 4).Value = IncentiveTenderID
                                MyCommon.LRTsp.Parameters.Add("@ROID", SqlDbType.Int, 4).Value = roid
                                MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int, 4).Value = t
                                MyCommon.LRTsp.Parameters.Add("@Value", SqlDbType.Decimal, 15).Value = Math.Round(TierTenderVal, CurPrecision)
                                MyCommon.LRTsp.ExecuteNonQuery()
                                MyCommon.Close_LRTsp()
                                historyString &= " TierLevel " & t & " requires " & TierTenderVal & "; "
                            Next
                        End If
                    ElseIf Request.QueryString("exVal") = "" Then
                        infoMessage = Copient.PhraseLib.Lookup("CPEoffer-con-tender.positivevalue", LanguageID)
                    End If
                End If
                If Request.QueryString("exVal") <> "" Then
                    TenderType = MyCommon.Extract_Val(GetCgiValue("ttID"))
                    If TenderType <= 0 Then TenderType = MyCommon.Extract_Val(GetCgiValue("selGroups"))
                    TierTenderVal = MyCommon.Extract_Val(Request.QueryString("exVal"))

                    ' set the ExcludedTender flag in RewardOptions
                    If TierTenderVal >= 1000000 Then
                        infoMessage = Copient.PhraseLib.Lookup("ueoffer-con-tender.LessThan1Million", LanguageID)
                    Else
                        'Add Tender or update if it is already there.
                        If IncentiveTenderID > 0 Then
                            MyCommon.QueryStr = "update CPE_IncentiveTenderTypes set TenderTypeID=" & TenderType & ", Value=" & Math.Round(TierTenderVal, CurPrecision) & ", LastUpdate=getdate() where IncentiveTenderID=" & IncentiveTenderID
                            MyCommon.LRT_Execute()
                        Else
                            MyCommon.QueryStr = "dbo.pa_CPE_AddTenderType"
                            MyCommon.Open_LRTsp()
                            MyCommon.LRTsp.Parameters.Add("@ROID", SqlDbType.Int, 4).Value = roid
                            MyCommon.LRTsp.Parameters.Add("@TenderTypeID", SqlDbType.Int, 4).Value = TenderType
                            MyCommon.LRTsp.Parameters.Add("@Value", SqlDbType.Decimal, 15).Value = Math.Round(TierTenderVal, CurPrecision)
                            MyCommon.LRTsp.Parameters.Add("@IncentiveTenderID", SqlDbType.Int, 4).Direction = ParameterDirection.Output
                            MyCommon.LRTsp.ExecuteNonQuery()
                            IncentiveTenderID = MyCommon.LRTsp.Parameters("@IncentiveTenderID").Value
                            MyCommon.Close_LRTsp()
                        End If
                        MyCommon.QueryStr = "update CPE_RewardOptions with (RowLock) set ExcludedTender=" & IIf(Request.QueryString("useasexcluded") = "on", 1, 0) & ", ExcludedTenderAmtRequired=" & TierTenderVal & " where RewardOptionID=" & roid & ";"
                        MyCommon.LRT_Execute()
                        MyCommon.QueryStr = "delete from CPE_IncentiveTenderTypeTiers where RewardOptionID=" & roid & " and IncentiveTenderID=" & IncentiveTenderID
                        MyCommon.LRT_Execute()

                        MyCommon.QueryStr = "dbo.pa_CPE_AddTenderTypeTiers"
                        MyCommon.Open_LRTsp()
                        MyCommon.LRTsp.Parameters.Add("@IncentiveTenderID", SqlDbType.Int, 4).Value = IncentiveTenderID
                        MyCommon.LRTsp.Parameters.Add("@ROID", SqlDbType.Int, 4).Value = roid
                        MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int, 4).Value = 1
                        MyCommon.LRTsp.Parameters.Add("@Value", SqlDbType.Decimal, 15).Value = Math.Round(TierTenderVal, CurPrecision)
                        MyCommon.LRTsp.ExecuteNonQuery()
                        MyCommon.Close_LRTsp()
                    End If
                Else
                    MyCommon.QueryStr = "update CPE_RewardOptions with (RowLock) set ExcludedTender=0, ExcludedTenderAmtRequired=0 where RewardOptionID=" & roid & ";"
                    MyCommon.LRT_Execute()
                End If
            Else
                ' no tender types are currently selected, so remove all the existing ones for this offer.
                ' A javascript alert will be shown

                'MyCommon.QueryStr = "delete from CPE_IncentiveTenderTypes with (RowLock) where RewardOptionID=" & roid
                'MyCommon.LRT_Execute()
                'MyCommon.QueryStr = "delete from CPE_IncentiveTenderTypeTiers with (RowLock) where RewardOptionID=" & roid
                'MyCommon.LRT_Execute()
                'MyCommon.QueryStr = "update CPE_IncentiveTenderTypes with (RowLock) set Deleted=1, LastUpdate=getdate() where Deleted=0 and RewardOptionID=" & roid
                'MyCommon.LRT_Execute()
                'MyCommon.QueryStr = "update CPE_RewardOptions with (RowLock) set ExcludedTender=0, ExcludedTenderAmtRequired=0 where RewardOptionID=" & roid & ";"
                'MyCommon.LRT_Execute()
            End If

            MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID
            MyCommon.LRT_Execute()
            If (infoMessage = "") Then
                CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
            Else
                CloseAfterSave = False
            End If
            ResetOfferApprovalStatus(OfferID)
            MyCommon.Activity_Log(3, OfferID, AdminUserID, Left(historyString, 250))
        ElseIf Not ValidDecimalPlaces Then
            infoMessage = Copient.PhraseLib.Detokenize("error.invalidDecimalPrecision", LanguageID, t1, CurPrecision)
        Else
            If Not ValidTier Then
                infoMessage = Copient.PhraseLib.Lookup("error.tiervalues", LanguageID)
                IncentiveTenderID = 0
            End If
        End If
    End If

    ' dig the offer info out of the database
    ' no one clicked anything
    MyCommon.QueryStr = "Select IncentiveID, IsTemplate, ClientOfferID,IncentiveName as Name,CPE.Description,PromoClassID,CRMEngineID,Priority," & _
                        "StartDate,EndDate,EveryDOW,EligibilityStartDate,EligibilityEndDate,TestingStartDate,TestingEndDate,P1DistQtyLimit,P1DistTimeType,P1DistPeriod," & _
                        "P2DistQtyLimit,P2DistTimeType,P2DistPeriod,P3DistQtyLimit,P3DistTimeType,P3DistPeriod,EnableImpressRpt,EnableRedeemRpt,CreatedDate,CPE.LastUpdate,CPE.Deleted, CPEOARptDate, CPEOADeploySuccessDate, CPEOADeployRpt," & _
                        "CRMRestricted,StatusFlag,OC.Description as CategoryName,IsTemplate,FromTemplate from CPE_Incentives as CPE with (NoLock) " & _
                        "left join OfferCategories as OC with (NoLock) on CPE.PromoClassID=OfferCategoryID where IncentiveID=" & Request.QueryString("OfferID") & ";"
    rst = MyCommon.LRT_Select
    For Each row In rst.Rows
        Name = MyCommon.NZ(row.Item("Name"), "")
        isTemplate = MyCommon.NZ(row.Item("IsTemplate"), False)
        FromTemplate = MyCommon.NZ(row.Item("FromTemplate"), False)
    Next

    'update the templates permission if necessary
    If (Request.QueryString("save") <> "" AndAlso Request.QueryString("IsTemplate") = "IsTemplate") Then
        ' time to update the status bits for the templates
        Dim form_Disallow_Edit As Integer = 0
        Dim form_Require_PP As Integer = 0

        If (Request.QueryString("Disallow_Edit") = "on") Then
            form_Disallow_Edit = 1
        End If

        If (Request.QueryString("require_pp") <> "") Then
            form_Require_PP = 1
        End If

        If (form_Disallow_Edit = 1 AndAlso form_Require_PP = 1) Then
            infoMessage = Copient.PhraseLib.Lookup("offer-con.lockeddenied", LanguageID)
            MyCommon.QueryStr = "update CPE_IncentiveTenderTypes with (RowLock) set DisallowEdit=" & form_Disallow_Edit & _
                                ", RequiredFromTemplate=0 where RewardOptionID=" & roid & " and Deleted=0;"
        Else
            MyCommon.QueryStr = "update CPE_IncentiveTenderTypes with (RowLock) set DisallowEdit=" & form_Disallow_Edit & _
                                ", RequiredFromTemplate=" & form_Require_PP & " " & _
                                " where RewardOptionID=" & roid & " and Deleted=0;"
        End If
        MyCommon.LRT_Execute()

        ' if necessary, create an empty condition
        If (form_Require_PP = 1) Then
            MyCommon.QueryStr = "select TenderTypeID from CPE_IncentiveTenderTypes with (NoLock) where RewardOptionID=" & roid & " and Deleted=0;"
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count = 0) Then
                MyCommon.QueryStr = "insert into CPE_IncentiveTenderTypes (RewardOptionID,Value,Deleted,LastUpdate,RequiredFromTemplate) " & _
                                    " values(" & roid & "," & IIf(Request.QueryString("Value"), Request.QueryString("Value"), "0") & ",0,getdate(),1)"
                MyCommon.LRT_Execute()
            End If
        End If

        If (infoMessage = "") Then
            CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
        Else
            CloseAfterSave = False
        End If
    End If

    If (isTemplate Or FromTemplate) Then
        ' lets dig the permissions if its a template
        MyCommon.QueryStr = "select DisallowEdit, RequiredFromTemplate from CPE_IncentiveTenderTypes with (NoLock) where RewardOptionID=" & roid & " and Deleted=0;"
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
            Disallow_Edit = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit"), False)
            RequirePP = MyCommon.NZ(rst.Rows(0).Item("RequiredFromTemplate"), False)
        Else
            Disallow_Edit = False
        End If
    End If
    Dim m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
    If Not IsTemplate Then
        DisabledAttribute = IIf(Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And Disallow_Edit), "", "disabled=""disabled""")
    Else
        DisabledAttribute = IIf(Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer, "", "disabled=""disabled""")
    End If

    Send_HeadBegin("term.offer", "term.tendercondition", OfferID)
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts()
%>
<script type="text/javascript">
// This is the javascript array holding the function list
// The PrintJavascriptArray ASP function can be used to print this array.
<%
   MyCommon.QueryStr = "select distinct TT.TenderTypeID, TT.Name from CPE_TenderTypes as TT with (NoLock) " & _ 
                        "where TT.Deleted=0 and TT.TenderTypeID not in " & _ 
                        "(select TenderTypeID from CPE_IncentiveTenderTypes where RewardOptionID=" & roid & " and IncentiveTenderID not in (" & IncentiveTenderID & ")) order by Name;"
  rst = MyCommon.LRT_Select
  If (rst.rows.count>0)
    Sendb("var functionlist = Array(")
    For Each row In rst.Rows
        Sendb("""" & MyCommon.NZ(row.item("Name"), "").ToString().Replace("""", "\""") & """,")
    Next
    Sendb(""""");")
    Sendb("var vallist = Array(")
    For Each row In rst.Rows
        Sendb("""" & MyCommon.NZ(row.item("TenderTypeID"), 0) & """,")
    Next
    Sendb(""""");")
  Else
    Sendb("var functionlist = Array(")
    Send("""" & "" & """);")
    Sendb("var vallist = Array(")
    Send("""" & "" & """);")
  End If
%>

// This is the function that refreshes the list after a keypress.
// The maximum number to show can be limited to improve performance with
// huge lists (1000s of entries).
// The function clears the list, and then does a linear search through the
// globally defined array and adds the matches back to the list.
function handleKeyUp(maxNumToShow) {
  var selectObj, textObj, functionListLength;
  var i,  numShown;
  var searchPattern;
  
  document.getElementById("functionselect").size = "16";
  
  // Set references to the form elements
  selectObj = document.forms[0].functionselect;
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
  
  // Create a regulare expression
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

function removeUsed() {
  handleKeyUp(99999);
  // this function will remove items from the functionselect box that are used in 
  // selected and excluded boxes
  
  var funcSel = document.getElementById('functionselect');
  var elSel = document.getElementById('selected');
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
function handleSelectClick(itemSelected) {
  textObj = document.forms[0].functioninput;
  
  selectObj = document.forms[0].functionselect;
  selectedValue = document.getElementById("functionselect").value;
  if(selectedValue != ""){ selectedText = selectObj[document.getElementById("functionselect").selectedIndex].text; }
  
  selectboxObj = document.forms[0].selected;
  selectedboxValue = document.getElementById("selected").value;
  if(selectedboxValue != ""){ selectedboxText = selectboxObj[document.getElementById("selected").selectedIndex].text; }
  
  if(itemSelected == "select1") {
    if(selectedValue != "") {
      // add items to selected box
      
      document.getElementById('deselect1').disabled=false;
      //if (selectboxObj.length > 0) {
      //  selectboxObj[0] = null;
      //}
      selectboxObj[selectboxObj.length] = new Option(selectedText,selectedValue);
      
      //Set the Tender Type ID when a group is selected
      document.getElementById('ttID').value = selectedValue;
      
      if(selectboxObj.length == 1) {
        document.getElementById('select1').disabled=true;
      }
      if(selectboxObj.length == 0) {
        // nothing in the select box so disable deselect
        document.getElementById('deselect1').disabled=true;
      }
    }
  }
  
  if(itemSelected == "deselect1") {
  document.getElementById('select1').disabled=false;
    if(selectedboxValue != "") {
      // remove items from selected box
      document.getElementById("selected").remove(document.getElementById("selected").selectedIndex)
      if(selectboxObj.length == 1) {
        document.getElementById('select1').disabled=true;
      }
      if(selectboxObj.length == 0) {
        // nothing in the select box so disable deselect
        document.getElementById('deselect1').disabled=true;
        document.getElementById('select1').disabled=false;
      }
    }
    document.getElementById("selected").selectedIndex = - 1;
  }
  // remove items from large list that are in the other lists
  removeUsed();
  updateButtons();
  return true;
}

function saveForm() {
  var funcSel = document.getElementById('functionselect');
  var elSel = document.getElementById('selected');
  var i,j;
  var selectList = "";
  var excludededList = "";
  var htmlContents = "";
  
  if (!validateEntry()) {
    return false;
  }
  
  if(!ValidSave()) {
    return false;
  }
  
  // assemble the list of values from the selected box
  for (i = elSel.length - 1; i>=0; i--) {
    if(elSel.options[i].value != ""){
      if(selectList != "") { selectList = selectList + ","; }
      selectList = selectList + elSel.options[i].value;
    }
  }
  
  // ok time to build up the hidden variables to pass for saving
  htmlContents = "<input type=\"hidden\" name=\"selGroups\" value=" + selectList + "> ";
  document.getElementById("hiddenVals").innerHTML = htmlContents;
  // alert(htmlContents);
  enableTiers();
  return true;
}

function ValidSave(){
  var elem = document.getElementById("selected"); 
  var msg = '';
  var retVal = true;
  
  if (elem != null && elem.options.length == 0) {
    retVal = false;
    msg += '* <% Sendb(Copient.PhraseLib.Lookup("CPEoffer-con-product.selectgroup", LanguageID)) %>'
    elem.focus();
  }
  
  if (msg != '') {
    alert(msg);
  }
  
  return retVal;
}

function getquerystring() {    
  var funcSel = document.getElementById('functionselect');
  var exElem = document.getElementById('excluded');
  var excluded = '';
  var exValElem = document.getElementById('exVal');
  var exVal = '';
  var elSel = document.getElementById('selected');
  var i,j;
  var selectList = "";
  var excludededList = "";
  var htmlContents = "";
  
  // assemble the list of values from the selected box
  for (i = elSel.length - 1; i>=0; i--) {
    if(elSel.options[i].value != ""){
      if(selectList != "") { selectList = selectList + ","; }
      selectList = selectList + elSel.options[i].value;
    }
  }
  
  if (funcSel != null) { selOpt = funcSel.value; }
  if (exElem != null) { excluded = exElem.value; }
  if (exValElem != null) { exVal = exValElem.value; }
  
  qstr = '<%sendb("LanguageID=" & LanguageID) %>' + '&CPETenderConditionValues=' + escape(selectList) + '&RewardOptionID=' + document.getElementById('roid').value + '&excluded=' + excluded + '&exVal=' + exVal;  // NOTE: no '?' before querystring
  return qstr;
}

//function updatepage(str) {
//  document.getElementById("results").innerHTML = str;
//}

function validateEntry() {
  var retVal = true;
  var qtyElem = document.getElementById("tVal0");
  var exValElem = document.getElementById("exVal");
  var allElem = document.getElementById("alltenders");
  var i = 0;
  var msg = '';
  var elemName = '';
  
  while (qtyElem != null) {
    // trim the string
    var qtyVal = qtyElem.value.replace(/^\s+|\s+$/g, ''); 
    if (qtyVal == "" || isNaN(qtyVal) || parseFloat(qtyVal) < 0) {
      retVal = false;
      if (msg != '') { msg += '\n\r\n\r'; }
      msg += '* <% Sendb(Copient.PhraseLib.Lookup("CPEoffer-con-tender.positivevalue", LanguageID)) %>';
      qtyElem.focus();
      qtyElem.select();
    } else if (qtyVal == "" || isNaN(qtyVal) || parseFloat(qtyVal) > 999999.999) {
      retVal = false;
      if (msg != '') { msg += '\n\r\n\r'; }
      msg += '* <% Sendb(Copient.PhraseLib.Lookup("ueoffer-con-tender.LessThan1Million", LanguageID)) %>';
      qtyElem.focus();
      qtyElem.select();
    }
    i ++;
    elemName = "tVal" + i
    qtyElem = document.getElementById(elemName);
  }
  
  if (exValElem != null && allElem != null && allElem.style.display != "none") {
    var val = exValElem.value.replace(/^\s+|\s+$/g, ''); 
    if (val == "" || isNaN(val) || parseFloat(val) <= 0.000) {
      retVal = false;
      if (msg != '') { msg += '\n\r\n\r'; }
      msg += '* <% Sendb(Copient.PhraseLib.Lookup("CPEoffer-con-tender.positivevalue", LanguageID)) %>';
      exValElem.focus();
      exValElem.select();
    } else if (val == "" || isNaN(val) || parseFloat(val) > 999999.999) {
      retVal = false;
      if (msg != '') { msg += '\n\r\n\r'; }
      msg += '* <% Sendb(Copient.PhraseLib.Lookup("ueoffer-con-tender.LessThan1Million", LanguageID)) %>';
      exValElem.focus();
      exValElem.select();
    }
  }
  
  if (msg != '') {
    alert(msg);
  }
  return retVal;
}

function updateButtons() {
  if (document.getElementById('selected').length > 0) {
    document.getElementById('select1').disabled=true; 
    document.getElementById('deselect1').disabled=false;
    if (document.getElementById('save') != null) {
      document.getElementById('save').disabled=false;
    }
  } else {
    document.getElementById('select1').disabled=false; 
    document.getElementById('deselect1').disabled=true;
    if (document.getElementById('save') != null) {
      document.getElementById('save').disabled=true;
    }
  }
  <%
  Dim EditRegardlessOfBuyer1 = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
    If Not IsTemplate Then
      If Not (Logix.UserRoles.EditOffer And EditRegardlessOfBuyer1 And Not (FromTemplate And Disallow_Edit)) Then
        Send("  disableAll();")
      End If
    Else
      If Not (Logix.UserRoles.EditTemplates And EditRegardlessOfBuyer1) Then
        Send("  disableAll();")
      End If
    End If
  %>
}

function disableAll() {
  document.getElementById('select1').disabled=true;
  document.getElementById('deselect1').disabled=true;
  document.getElementById('functionselect').disabled=true;
  document.getElementById('selected').disabled=true;
}
</script>
<%
  Send("<script type=""text/javascript"">")
  
  Send("function handleExcludedClick() { ")
  Send("  var elem = document.getElementById('selected');")
  Send("  var selectedValue = '';")
  Send("  var selectedName = '';")
  Send("")
  Send("  if (elem!=null && elem.options.length > 0) {")
  Send("    selectedValue = elem.options[0].value;")
  Send("    selectedName = elem.options[0].text;")
  Send("  }")
  If Not ExcludedTender Then
    Send(" window.location = 'UEoffer-con-tender.aspx?OfferID=" & OfferID & "&IncentiveTenderID=" & IncentiveTenderID & "&ExcludedTender=1" & _
         "&selGroups=' + selectedValue + '&selectedName=' + selectedName;")
  ElseIf ExcludedTender Then
    Send(" window.location = 'UEoffer-con-tender.aspx?OfferID=" & OfferID & "&IncentiveTenderID=" & IncentiveTenderID & "&ExcludedTender=0" & _
         "&selGroups=' + selectedValue + '&selectedName=' + selectedName;")
  End If
  Send("} ")
  
  Send("function ChangeParentDocument() { ")
    If (EngineID = 3) Then
        Send("  if (opener != null) {")
        Send("    var newlocation = '/logix/web-offer-con.aspx?OfferID=" & OfferID & "'; ")
        Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
        Send("  opener.location = '/logix/web-offer-con.aspx?OfferID=" & OfferID & "'; ")
    ElseIf (EngineID = 5) Then
        Send("  if (opener != null) {")
        Send("    var newlocation = '/logix/email-offer-con.aspx?OfferID=" & OfferID & "'; ")
        Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
        Send("  opener.location = '/logix/email-offer-con.aspx?OfferID=" & OfferID & "'; ")
    ElseIf (EngineID = 6) Then
        Send("  if (opener != null) {")
        Send("    var newlocation = '/logix/CAM/CAM-offer-con.aspx?OfferID=" & OfferID & "'; ")
        Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
        Send("  opener.location = '/logix/CAM/CAM-offer-con.aspx?OfferID=" & OfferID & "'; ")
    ElseIf (EngineID = 9) Then
        Send("  if (opener != null) {")
        Send("    var newlocation = '/logix/UE/UEoffer-con.aspx?OfferID=" & OfferID & "'; ")
        Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
        Send("  opener.location = '/logix/UE/UEoffer-con.aspx?OfferID=" & OfferID & "'; ")
    Else
        Send("  if (opener != null) {")
        Send("    var newlocation = '/logix/CPEoffer-con.aspx?OfferID=" & OfferID & "'; ")
        Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
        Send("  opener.location = '/logix/CPEoffer-con.aspx?OfferID=" & OfferID & "'; ")
    End If
    Send("} ")
    Send("} ")
  Send("} ")
  Send("</script>")
  Send_HeadEnd()
  
  If (isTemplate) Then
    Send_BodyBegin(12)
  Else
    Send_BodyBegin(2)
  End If
  If (Logix.UserRoles.AccessOffers = False AndAlso Not isTemplate) Then
    Send_Denied(2, "perm.offers-access")
    GoTo done
  End If
  If (Logix.UserRoles.AccessTemplates = False AndAlso isTemplate) Then
    Send_Denied(2, "perm.offers-access-templates")
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
%>
<form action="#" id="mainform" name="mainform" onsubmit="return saveForm();">
  <span id="hiddenVals"></span>
  <input type="hidden" id="OfferID" name="OfferID" value="<% sendb(OfferID) %>" />
  <input type="hidden" id="Name" name="Name" value="<% sendb(Name) %>" />
  <input type="hidden" id="ConditionID" name="ConditionID" value="<% sendb(ConditionID) %>" />
  <input type="hidden" id="roid" name="roid" value="<%sendb(roid) %>" />
  <input type="hidden" id="excluded" name="excluded" value="<%sendb(IIf(ExcludedTender, 1,0)) %>" />
  <input type="hidden" id="EngineID" name="EngineID" value="<% Sendb(EngineID) %>" />
  <input type="hidden" id="IsTemplate" name="IsTemplate" value="<% 
        If (IsTemplate) Then
          Sendb("IsTemplate")
        Else
          Sendb("Not")
        End If
        %>" />
  <div id="intro">
    <%
      If (isTemplate) Then
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.tendercondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
      Else
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.tendercondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
      End If
    %>
    <div id="controls">
      <% If (isTemplate) Then%>
      <span class="temp">
        <input type="checkbox" class="tempcheck" id="Disallow_Edit" name="Disallow_Edit"<% if(Disallow_Edit)then sendb(" checked=""checked""") %> />
        <label for="Disallow_Edit">
          <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
        </label>
      </span>
      <% End If%>
      <% 
          m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
      If ((Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
        If Not IsTemplate Then
                  If (Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And Disallow_Edit) And Not IsOfferWaitingForApproval(OfferID)) Then Send_Save()
        Else
                If (Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer) Then Send_Save()
        End If
      End If
      %>
    </div>
  </div>
  <div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <div id="column1">
      <div class="box" id="types">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.types", LanguageID))%>
          </span>
        </h2>
        <input type="radio" id="functionradio1" name="functionradio" checked="checked"<% sendb(DisabledAttribute) %> /><label for="functionradio1"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
        <input type="radio" id="functionradio2" name="functionradio"<% sendb(DisabledAttribute) %> /><label for="functionradio2"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
        <input type="text" class="medium" id="functioninput" name="functioninput" maxlength="100" onkeyup="handleKeyUp(200);" value=""<% sendb(DisabledAttribute) %> /><br />
        <select class="longer" id="functionselect" name="functionselect" size="16"<% sendb(DisabledAttribute) %>>
        </select>
        <br />
        <br class="half" />
        <input type="button" class="regular select" id="select1" name="select1" value="&#9660; <% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID)) %>" onclick="handleSelectClick('select1');"<% sendb(DisabledAttribute) %> />&nbsp;
        <input type="button" class="regular deselect" id="deselect1" name="deselect1" value="<% Sendb(Copient.PhraseLib.Lookup("term.deselect", LanguageID)) %> &#9650;" disabled="disabled" onclick="handleSelectClick('deselect1');" /><br />
        <br class="half" />
        <select class="longer" id="selected" name="selected" size="5"<% sendb(DisabledAttribute) %>>
          <%
            If IncentiveTenderID > 0 Then
              MyCommon.QueryStr = "select distinct TT.TenderTypeID, TT.Name, ITT.Value from CPE_TenderTypes as TT with (NoLock) " & _
                                  "left join CPE_IncentiveTenderTypes as ITT with (NoLock) on TT.TenderTypeID=ITT.TenderTypeID and ITT.RewardOptionID = " & roid & " and ITT.Deleted=0 " & _
                                  "where TT.Deleted=0 and ITT.IncentiveTenderID=" & IncentiveTenderID & " order by Name;"
              rst = MyCommon.LRT_Select
              For Each row In rst.Rows
                If Not IsDBNull(row.Item("Value")) Then
                  Send("<option value=""" & MyCommon.NZ(row.Item("TenderTypeID"), 0) & """>" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
                End If
              Next
            ElseIf IncentiveTenderID = 0 Then
              If Request.QueryString("selGroups") <> "" Then
                MyCommon.QueryStr = "select Name from CPE_TenderTypes where TenderTypeID=" & MyCommon.Extract_Val(Request.QueryString("selGroups"))
                rst = MyCommon.LRT_Select
                If rst.Rows.Count > 0 Then
                  Send("<option value=""" & MyCommon.Extract_Val(Request.QueryString("selGroups")) & """>" & MyCommon.NZ(rst.Rows(0).Item("Name"), "") & "</option>")
                End If
              End If
            End If
          %>
        </select>
        <hr class="hidden" />
      </div>
    </div>
    
    <div id="gutter">
    </div>
    
    <div id="column2">
      <div class="box" id="value">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.value", LanguageID))%>
          </span>
        </h2>
        <div id="results">
        </div>
        <%
          MyCommon.QueryStr = "select ITT.IncentiveTenderID, TEN.TenderTypeID, TEN.Name, ITT.Value from " & _
                              "(select TT.TenderTypeID, TT.Name from CPE_TenderTypes TT with (NoLock) " & _
                              "where deleted=0 and TenderTypeID in (" & SelGrp & ")) TEN " & _
                              "left join CPE_IncentiveTenderTypes ITT with (NoLock) on ITT.TenderTypeID = TEN.TenderTypeID " & _
                              "where ITT.Deleted=0 and ITT.RewardOptionID=" & roid & " and ITT.IncentiveTenderID=" & IncentiveTenderID
          dt = MyCommon.LRT_Select
          
          Send("<table summary=""" & Copient.PhraseLib.Lookup("term.values", LanguageID) & """>")
          Send("  <thead>")
          Send("    <tr>")
          Send("      <th class=""th-value"" style=""" & IIf(ExcludedTender, "display:none;", "") & """ scope=""col"">" & Copient.PhraseLib.Lookup("term.value", LanguageID) & "</th>")
          Send("    </tr>")
          Send("  </thead>")
          Send("  <tbody>")
          If (dt.Rows.Count > 0) Then
            Send("    <tr>")
            Send("      <td " & IIf(ExcludedTender, "style=""display:none;""", "") & ">")
            tempTenderValue = Math.Round(CDec(MyCommon.NZ(dt.Rows(0).Item("Value"), 0)), CurPrecision)
            If IncentiveTenderID = 0 Then
            
              For t = 1 To TierLevel
                If (SupportGlobalAndTieredConditions = 1 And t = 1 And TierLevel > 1) Then
                  Send("  <input type=""checkbox"" name=""useSameTierValue"" id =""useSameTierValue"" style=""margin-left:0px;"" align='top' value=""1""" & IIf(UseSameTierValue = 1, " checked=""checked""", "") & " style=""margin-left:7px;"" onclick=""setSameTierValue(" & TierLevel & ")""/>")
                  Sendb("  <label for=""useThisValueForAllTiers"" align='top'>" & Copient.PhraseLib.Lookup("term.UseThisValueForAllTiers", LanguageID) & "</label><br>") 
                End If

                If TierLevel = 1 Then
                  If Request.QueryString("t1_tVal") <> "" Then
                    Send("        " & CurSymbol & "&nbsp;<input type=""text"" class=""shorter"" maxlength=""16"" name=""t1_tVal"" id=""t1_tVal"" value=""" & Math.Round(CDec(MyCommon.Extract_Val(Request.QueryString("t1_tVal"))), CurPrecision) & """" & DisabledAttribute & " />&nbps;" & CurAbbr)
                  Else
                    Send("        " & CurSymbol & "&nbsp;<input type=""text"" class=""shorter"" maxlength=""16"" name=""t1_tVal"" id=""t1_tVal"" value=""0""" & DisabledAttribute & " />&nbps;" & CurAbbr)
                  End If
                End If
                If Request.QueryString("t" & t & "_tVal") <> "" Then
                  Send("        " & CurSymbol & "&nbsp;<input type=""text"" class=""shorter"" maxlength=""16"" name=""t1_tVal"" id=""t1_tVal"" value=""" & Math.Round(CDec(MyCommon.Extract_Val(Request.QueryString("t" & t & "_tVal"))), CurPrecision) & """" &  IIf(UseSameTierValue = 1, " disabled=""disabled""", DisabledAttribute) & " />&nbsp;" & CurAbbr & "<br />")
                Else
                  Send("        " & CurSymbol & "&nbsp;<input type=""text"" class=""shorter"" maxlength=""16"" name=""t1_tVal"" id=""t1_tVal"" value=""0""" & DisabledAttribute & " />&nbsp;" & CurAbbr & "<br />")
                End If
              Next
            Else
              For t = 1 To TierLevel
                If (SupportGlobalAndTieredConditions = 1 And t = 1 And TierLevel > 1) Then
                  Send("  <input type=""checkbox"" name=""useSameTierValue"" id =""useSameTierValue"" style=""margin-left:0px;"" align='top' value=""1""" & IIf(UseSameTierValue = 1, " checked=""checked""", "") & " style=""margin-left:7px;"" onclick=""setSameTierValue(" & TierLevel & ")""/>")
                  Sendb("  <label for=""useThisValueForAllTiers"" align='top'>" & Copient.PhraseLib.Lookup("term.UseThisValueForAllTiers", LanguageID) & "</label><br>") 
                End If
                MyCommon.QueryStr = "select Value from CPE_IncentiveTenderTypeTiers with (NoLock) where RewardOptionID=" & roid & " and IncentiveTenderID=" & IncentiveTenderID & " and TierLevel=" & t & ";"
                TierDT = MyCommon.LRT_Select()
                If TierLevel = 1 Then
                  Send("        " & CurSymbol & "&nbsp;<input type=""text"" class=""shorter"" maxlength=""16"" name=""t1_tVal"" id=""t1_tVal"" value=""" & tempTenderValue.ToString(MyCommon.GetAdminUser.Culture) & """" & DisabledAttribute & " />&nbsp;" & CurAbbr)
                Else
                  Send("        <label for=""t" & t & "_tVal"">" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</label>")
                  If TierDT.Rows.Count > 0 Then
                    Send("        " & CurSymbol & "&nbsp;<input type=""text"" class=""shorter"" maxlength=""16"" name=""t" & t & "_tVal"" id=""t" & t & "_tVal"" value=""" & Math.Round(CDec(MyCommon.NZ(TierDT.Rows(0).Item("Value"), 0)), CurPrecision) & """" &  IIf(UseSameTierValue = 1, " disabled=""disabled""", DisabledAttribute) & " />&nbsp;" & CurAbbr & "<br />")
                  Else
                    Send("        " & CurSymbol & "&nbsp;<input type=""text"" class=""shorter"" maxlength=""16"" name=""t" & t & "_tVal"" id=""t" & t & "_tVal"" value=""0""" & DisabledAttribute & " />&nbsp;" & CurAbbr & "<br />")
                  End If
                End If
              Next
            End If
            Send("        <input type=""hidden"" name=""IncentiveTenderID"" id=""IncentiveTenderID"" value=""" & IncentiveTenderID & """ />")
            Send("        <input type=""hidden"" name=""ttID"" id=""ttID"" value=""" & MyCommon.NZ(dt.Rows(0).Item("TenderTypeID"), "0") & """ />")
            Send("      </td>")
            Send("    </tr>")
            Send("  </tbody>")
            Send("</table>")
          Else
            Send("    <tr>")
            Send("      <td " & IIf(ExcludedTender, "style=""display:none;""", "") & ">")
            If IncentiveTenderID = 0 Then
              If TierLevel = 1 Then
                If Request.QueryString("t1_tVal") <> "" Then
                  Send("        " & CurSymbol & "&nbsp;<input type=""text"" class=""shorter"" maxlength=""16"" name=""t1_tVal"" id=""t1_tVal"" value=""" & Math.Round(CDec(MyCommon.Extract_Val(Request.QueryString("t1_tVal"))), CurPrecision) & """" & DisabledAttribute & " />&nbsp;" & CurAbbr)
                Else
                  Send("        " & CurSymbol & "&nbsp;<input type=""text"" class=""shorter"" maxlength=""16"" name=""t1_tVal"" id=""t1_tVal"" value=""0""" & DisabledAttribute & " />&nbsp;" & CurAbbr)
                End If
              Else
                For t = 1 To TierLevel
                  If (SupportGlobalAndTieredConditions = 1 And t = 1 And TierLevel > 1) Then
                    Send("  <input type=""checkbox"" name=""useSameTierValue"" id =""useSameTierValue"" style=""margin-left:0px;"" align='top' value=""1""" & IIf(UseSameTierValue = 1, " checked=""checked""", "") & " style=""margin-left:7px;"" onclick=""setSameTierValue(" & TierLevel & ")""/>")
                    Sendb("  <label for=""useThisValueForAllTiers"" align='top'>" & Copient.PhraseLib.Lookup("term.UseThisValueForAllTiers", LanguageID) & "</label><br>") 
                  End If
                  If Request.QueryString("t" & t & "_tVal") <> "" Then
                    Send("        <label for=""t" & t & "_tVal"">" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</label>")
                    Send("        " & CurSymbol & "&nbsp;<input type=""text"" class=""shorter"" maxlength=""16"" name=""t" & t & "_tVal"" id=""t" & t & "_tVal"" value=""" & Math.Round(CDec(MyCommon.Extract_Val(Request.QueryString("t" & t & "_tVal"))), CurPrecision) & """" & IIf(UseSameTierValue = 1, " disabled=""disabled""", DisabledAttribute) & " />&nbsp;" & CurAbbr & "<br />")
                  Else
                    Send("        <label for=""t" & t & "_tVal"">" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</label>")
                    Send("        " & CurSymbol & "&nbsp;<input type=""text"" class=""shorter"" maxlength=""16"" name=""t" & t & "_tVal"" id=""t" & t & "_tVal"" value=""0""" & DisabledAttribute & " />&nbsp;" & CurAbbr & "<br />")
                  End If
                Next
              End If
            Else
              For t = 1 To TierLevel
                If TierLevel = 1 Then
                Else
                  If (SupportGlobalAndTieredConditions = 1 And t = 1 And TierLevel > 1) Then
                    Send("  <input type=""checkbox"" name=""useSameTierValue"" id =""useSameTierValue"" style=""margin-left:0px;"" align='top' value=""1""" & IIf(UseSameTierValue = 1, " checked=""checked""", "") & " style=""margin-left:7px;"" onclick=""setSameTierValue(" & TierLevel & ")""/>")
                    Sendb("  <label for=""useThisValueForAllTiers"" align='top'>" & Copient.PhraseLib.Lookup("term.UseThisValueForAllTiers", LanguageID) & "</label><br>") 
                  End If
                  Send("        <label for=""t" & t & "_tVal"">" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</label>")
                End If
                Send("        " & CurSymbol & "&nbsp;<input type=""text"" class=""shorter"" maxlength=""16"" name=""t" & t & "_tVal"" id=""t" & t & "_tVal"" value=""0""" &  IIf(UseSameTierValue = 1, " disabled=""disabled""", DisabledAttribute) & " />&nbsp;" & CurAbbr & "<br />")
              Next
            End If
            Send("        <input type=""hidden"" name=""IncentiveTenderID"" id=""IncentiveTenderID"" value=""" & IncentiveTenderID & """ />")
            Send("        <input type=""hidden"" name=""ttID"" id=""ttID"" value=""" & SelGrp & """ />")
            Send("      </td>")
            Send("    </tr>")
            Send("  </tbody>")
            Send("</table>")
          End If
          Send("<br class=""half"" />")
          Send("<div id=""exclusion"">")
          Send("  <input type=""checkbox"" id=""useasexcluded"" name=""useasexcluded""" & IIf(ExcludedTender, " checked=""checked""", "") & " onclick=""handleExcludedClick();""" & DisabledAttribute & " />")
          Send("  <label for=""useasexcluded"">" & Copient.PhraseLib.Lookup("condition.useasexcluded", LanguageID) & "</label>")
          If ExcludedTender Then
            Send("  <div id=""alltenders"" style=""" & IIf(Not ExcludedTender, "display:none;", "") & """>")
            Send("    <br class=""half"" />")
            Send("    <table summary="""">")
            Send("      <thead>")
            Send("        <tr>")
            Send("          <th class=""th-name"" scope=""col"">" & Copient.PhraseLib.Lookup("term.tender", LanguageID) & "</th>")
            Send("          <th scope=""col"" style=""width:70px;"">" & Copient.PhraseLib.Lookup("term.value", LanguageID) & "</th>")
            Send("        </tr>")
            Send("      </thead>")
            Send("      <tbody>")
            Send("        <tr>")
            Send("          <td><label for=""exVal"">" & Copient.PhraseLib.Lookup("ueoffer-con-tender.OtherRequired", LanguageID) & "</label></td>")
            Send("          <td>" & CurSymbol & "&nbsp;<input type=""text"" class=""shorter"" maxlength=""16"" name=""exVal"" id=""exVal"" value=""" & Math.Round(ExcludedValue, CurPrecision) & """" & DisabledAttribute & " />&nbsp;" & CurAbbr & "</td>")
            Send("        </tr>")
            Send("      </tbody>")
            Send("    </table>")
            Send("  </div>")            
          End If
          Send("</div>")
        %>
        <hr class="hidden" />
      </div>
    </div>
  </div>
</form>

<script type="text/javascript">
    function setSameTierValue(tierLevels){
      var box = document.getElementById("useSameTierValue");
      var text;
      if(box.checked){
        for (i=1; i < (tierLevels + 1); i++){
          text = "t" + i.toString() + "_tVal";
          //alert(document.getElementById("tier1_l1discountamt").value.toString());
          document.getElementById(text).value = document.getElementById("t1_tVal").value;
          document.getElementById(text).setAttribute('disabled', 'disabled');
        } 
      }
      else{
        for (i=1; i < (tierLevels + 1); i++){
          text = "t" + i.toString() + "_tVal";
          document.getElementById(text).disabled = false;
        } 
      }
    }

     
    function enableTiers(){
      var t = 1
      qtyElem = document.getElementById("t" + t + "_tVal");
  
      while (qtyElem != null) 
      {
	    qtyElem.disabled = false
        t++;
        qtyElem = document.getElementById("t" + t + "_tVal");
      }
    }
 
<% If (CloseAfterSave) Then %>
    window.close();
<% Else %>
    removeUsed();
    updateButtons();
<% End If %>
</script>

<script runat="server">
  Function GetDecimalPlaces(ByVal value As String) As Integer
    Dim Places As Integer = 0
    Dim TempDec As Decimal
    Dim NumberAsString As String
    Dim IndexOfDecimalPoint As Integer
    
    If value IsNot Nothing AndAlso Decimal.TryParse(value, TempDec) Then
      NumberAsString = TempDec.ToString()
      IndexOfDecimalPoint = NumberAsString.IndexOf(".")
      If IndexOfDecimalPoint > -1 Then
         Places = NumberAsString.Substring(IndexOfDecimalPoint + 1).Length
      End If
    End If
    
    Return Places
  End Function
</script>

<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd("mainform", "functioninput")
  MyCommon = Nothing
  Logix = Nothing
%>
