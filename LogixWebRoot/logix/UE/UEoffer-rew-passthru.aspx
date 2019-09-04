﻿<%@ Page Language="vb" Debug="true" CodeFile="ue-cb.vb" Inherits="UECB" %>

<%@ Import Namespace="Copient" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
    ' *****************************************************************************
    ' * FILENAME: UEoffer-rew-passthru.aspx 
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
    Dim Localization As Copient.Localization
    Dim rst As DataTable
    Dim rst2 As DataTable
    Dim rst3 As DataTable
    Dim dt As DataTable
    Dim dt2 As DataTable
    Dim row As DataRow

    Dim OfferID As Long
    Dim OfferName As String
    Dim Phase As Integer
    Dim RewardID As String
    Dim DeliverableID As Long
    Dim PassThruRewardID As Long
    Dim PassThruPKID As Integer = 0
    Dim PassThruRewardName As String = ""
    Dim TierPKID As Integer = 0
    Dim bIsErrorMsg As Boolean = False
    Dim TouchPoint As Integer = 0
    Dim TpROID As Integer = 0
    Dim CreateROID As Integer = 0
    Dim LSInterfaceID As Integer = 0
    Dim ActionTypeID As Integer = 0
    Dim DataTemplate As String = ""
    Dim TierString As String = ""

    Dim PresString As String = ""
    Dim PresIndex As Integer = 0

    Dim CloseAfterSave As Boolean = False
    Dim Disallow_Edit As Boolean = True
    Dim FromTemplate As Boolean = False
    Dim IsTemplate As Boolean = False
    Dim IsTemplateVal As String = "Not"
    Dim DisabledAttribute As String = ""
    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False
    Dim EngineID As Integer = 2
    Dim BannersEnabled As Boolean = True
    Dim TierLevels As Integer = 1
    Dim t As Integer = 1
    Dim i As Integer = 1
    Dim ValidTiers As Boolean = True

    Dim StartIndex As Integer = 0
    Dim EndIndex As Integer = 0
    Dim TokenLength As Integer = 0
    Dim TokenID As Integer = 0
    Dim TokenValue As String = ""
    Dim Tokens(-1) As String
    Dim TokensList As String = ""
    Dim CurrentValue As String = ""
    Dim ReplacementText As String = ""
    Dim FinalString As String = ""
    Dim BadTagValues As Integer = 0
    Dim RewardValue As Decimal
    Dim RewardRequired As Boolean = True

    Dim MultiLanguageEnabled As Boolean = False
    Dim MultiLanguagePresTag As Boolean = False
    Dim DefaultLanguageID As Integer = 0
    Dim DefaultLanguageCode As String = ""
    Dim LanguagesDT As DataTable
    Dim lrow As DataRow
    Dim HasMLTag As Boolean = False
    Dim MLI As New Copient.Localization.MultiLanguageRec

    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If

    Response.Expires = 0
    MyCommon.AppName = "UEoffer-rew-passthru.aspx"
    MyCommon.Open_LogixRT()
    MyCommon.Open_LogixXS()
    MyCommon.Open_LogixWH()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
    Localization = New Copient.Localization(MyCommon)

    BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")

    MultiLanguageEnabled = IIf(MyCommon.Fetch_SystemOption(124) = "1", True, False)
    Integer.TryParse(MyCommon.Fetch_SystemOption(1), DefaultLanguageID)

    If DefaultLanguageID > 0 Then
        MyCommon.QueryStr = "select MSNetCode from Languages with (NoLock) where LanguageID=@DefaultLanguageID"
        MyCommon.DBParameters.Add("@DefaultLanguageID", SqlDbType.Int).Value = DefaultLanguageID
        rst = MyCommon.ExecuteQuery(DataBases.LogixRT)
        If rst.Rows.Count > 0 Then
            DefaultLanguageCode = MyCommon.NZ(rst.Rows(0).Item("MSNetCode"), "")
        End If
    End If

    OfferID = Request.QueryString("OfferID")
    'This would redirect to summary page if offer is awaiting deployment or if offer is awaiting recommendations.
    CheckIfValidOffer(MyCommon, OfferID)
    RewardID = Request.QueryString("RewardID")
    DeliverableID = MyCommon.Extract_Val(Request.QueryString("DeliverableID"))
    RewardRequired = (MyCommon.Extract_Val(GetCgiValue("requiredToDeliver")) = 1)
    Phase = MyCommon.Extract_Val(Request.QueryString("Phase"))
    PassThruRewardID = MyCommon.Extract_Val(Request.QueryString("PassThruRewardID"))
    If (Phase = 0) Then Phase = MyCommon.Extract_Val(Request.Form("Phase"))
    If (Phase = 0) Then Phase = 3
    If (Request.QueryString("EngineID") <> "") Then
        EngineID = MyCommon.Extract_Val(Request.QueryString("EngineID"))
    Else
        MyCommon.QueryStr = "select EngineID from OfferIDs where OfferID= @OfferID"
        MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
        rst = MyCommon.ExecuteQuery(DataBases.LogixRT)
        If rst.Rows.Count > 0 Then
            EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), 0)
        End If
    End If

    Dim isTranslatedOffer As Boolean =MyCommon.IsTranslatedUEOffer(OfferID,  MyCommon)
    Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249)="1",True,False)

    Dim bEnableAdditionalLockoutRestrictionsOnOffers As Boolean = IIf(MyCommon.Fetch_SystemOption(260) = "1", True, False)
    Dim bOfferEditable As Boolean = MyCommon.IsOfferEditablePastLockOutPeriod(Logix.UserRoles.EditOfferPastLockoutPeriod, MyCommon, OfferID)

    MyCommon.QueryStr = "select LSInterfaceID, ActionTypeID, Name, PhraseID from PassThruRewards with (NoLock) " & _
                        "where PassThruRewardID= @PassThruRewardID"
    MyCommon.DBParameters.Add("@PassThruRewardID", SqlDbType.BigInt).Value = PassThruRewardID
    rst = MyCommon.ExecuteQuery(DataBases.LogixRT)

    If rst.Rows.Count > 0 Then
        LSInterfaceID = MyCommon.NZ(rst.Rows(0).Item("LSInterfaceID"), 0)
        ActionTypeID = MyCommon.NZ(rst.Rows(0).Item("ActionTypeID"), 0)
        If IsDBNull(rst.Rows(0).Item("PhraseID")) Then
            PassThruRewardName = MyCommon.NZ(rst.Rows(0).Item("Name"), "")
        Else
            PassThruRewardName = Copient.PhraseLib.Lookup(rst.Rows(0).Item("PhraseID"), LanguageID)
        End If
    End If

    MyCommon.QueryStr = "select TierLevels from CPE_RewardOptions where RewardOptionID=@RewardID"
    MyCommon.DBParameters.Add("@RewardID", SqlDbType.BigInt).Value = RewardID
    rst = MyCommon.ExecuteQuery(DataBases.LogixRT)
    If rst.Rows.Count > 0 Then
        TierLevels = MyCommon.NZ(rst.Rows(0).Item("TierLevels"), 1)
    End If

    TouchPoint = MyCommon.Extract_Val(Request.QueryString("TP"))
    If (TouchPoint > 0) Then
        TpROID = MyCommon.Extract_Val(Request.QueryString("ROID"))
    End If

    'Fetch the offer name and template details
    MyCommon.QueryStr = "Select IncentiveName, IsTemplate, FromTemplate from CPE_Incentives with (NoLock) where IncentiveID= @OfferID"
    MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
    rst = MyCommon.ExecuteQuery(DataBases.LogixRT)
    If (rst.Rows.Count > 0) Then
        OfferName = MyCommon.NZ(rst.Rows(0).Item("IncentiveName"), "")
        IsTemplate = MyCommon.NZ(rst.Rows(0).Item("IsTemplate"), False)
        FromTemplate = MyCommon.NZ(rst.Rows(0).Item("FromTemplate"), False)
    End If
    IsTemplateVal = IIf(IsTemplate, "IsTemplate", "Not")

    Dim DTLen As Integer = 0
    Dim TVLen As Integer = 0
    Dim TempDate As Date

    'Save logic
    If (Request.QueryString("save") <> "") Then
        'First, validate the save data
        MyCommon.QueryStr = "select DataTemplate from PassThruRewards with (NoLock) where PassThruRewardID= @PassThruRewardID"
        MyCommon.DBParameters.Add("@PassThruRewardID", SqlDbType.BigInt).Value = PassThruRewardID
        rst2 = MyCommon.ExecuteQuery(DataBases.LogixRT)
        If rst2.Rows.Count > 0 Then
            DataTemplate = MyCommon.NZ(rst2.Rows(0).Item("DataTemplate"), "")
            DTLen = Len(Regex.Replace(DataTemplate, "\^[0-9]+;", ""))
            Tokens = ParseTokenValues(DataTemplate)
            For t = 1 To TierLevels
                For i = 0 To Tokens.GetUpperBound(0)
                    TokenID = Tokens(i)
                    TokenValue = Request.QueryString("t" & t & "_token" & TokenID)
                    TVLen = TVLen + Len(TokenValue)
                    If TokenValue = "" Then
                        infoMessage = Copient.PhraseLib.Lookup("inputValidator.emptyStringNotPermitted", LanguageID)
                    ElseIf TokenID = 33 Then
                        Dim tmp As Integer
                        If (Not UInteger.TryParse(TokenValue, tmp)) OrElse (TokenValue = 0) Then
                            infoMessage = Copient.PhraseLib.Lookup("term.rew-maxselectionerror", LanguageID)
                        End If
                    End If

                    MyCommon.QueryStr = "select ParamDataTypeID, ParamName from PassThruPresTags with (NoLock) where PassThruPresTagID= @TokenID"
                    MyCommon.DBParameters.Add("@TokenID", SqlDbType.Int).Value = TokenID
                    rst3 = MyCommon.ExecuteQuery(DataBases.LogixRT)
                    If rst3.Rows(0).Item("ParamDataTypeID") = 1 Then
                        If Not (IsNumeric(TokenValue)) OrElse (InStr(TokenValue, ".") OrElse InStr(TokenValue, "-")) Then
                            BadTagValues += 1
                            infoMessage = Copient.PhraseLib.Detokenize("CPEoffer-rew.passthru-invalidvalue", LanguageID, rst3.Rows(0).Item("ParamName"))
                        End If
                    ElseIf rst3.Rows(0).Item("ParamDataTypeID") = 2 Then
                        If Not (IsNumeric(TokenValue)) OrElse (Double.Parse(TokenValue) > 999999999999.99) OrElse (Math.Round(Double.Parse(TokenValue), 2) <> Double.Parse(TokenValue)) Then
                            BadTagValues += 1
                            infoMessage = Copient.PhraseLib.Detokenize("CPEoffer-rew.passthru-invalidvalue", LanguageID, rst3.Rows(0).Item("ParamName"))
                        End If
                    ElseIf rst3.Rows(0).Item("ParamDataTypeID") = 3 Then
                        If Not (IsNumeric(TokenValue)) OrElse (Double.Parse(TokenValue) > 999999999999.999) OrElse (Math.Round(Double.Parse(TokenValue), 3) <> Double.Parse(TokenValue)) Then
                            BadTagValues += 1
                            infoMessage = Copient.PhraseLib.Detokenize("CPEoffer-rew.passthru-invalidvalue", LanguageID, rst3.Rows(0).Item("ParamName"))
                        End If
                    ElseIf rst3.Rows(0).Item("ParamDataTypeID") = 5 Then
                        If Not (Date.TryParse(TokenValue, MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, TempDate)) Then
                            BadTagValues += 1
                            infoMessage = Copient.PhraseLib.Detokenize("CPEoffer-rew.passthru-invalidvalue", LanguageID, rst3.Rows(0).Item("ParamName"))
                        Else
                            If Not (TokenValue Like "####[-]##[-]##") Then
                                BadTagValues += 1
                            End If
                        End If
                    End If
                Next
            Next
        End If

        'Determine if there's at least one multilanguage tag in the datatemplate for this passthru
        TokensList = "0"
        For i = 0 To Tokens.GetUpperBound(0)
            TokenID = Tokens(i)
            TokensList &= "," & TokenID
        Next
        MyCommon.QueryStr = "select PassThruPresTagID from PassThruPresTags with (NoLock) where MultiLanguage=1 and PassThruPresTagID in (" & TokensList & ");"
        rst3 = MyCommon.LRT_Select()
        If rst3.Rows.Count > 0 Then
            HasMLTag = True
        End If
        If infoMessage = "" Then
            If (BadTagValues = 0) AndAlso ((DTLen + TVLen) <= 2000) Then
                If DeliverableID = 0 Then
                    'It's a new deliverable, so create it
                    MyCommon.QueryStr = "dbo.pa_CPE_AddPassThruReward"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.Int, 4).Value = OfferID
                    MyCommon.LRTsp.Parameters.Add("@ROID", SqlDbType.Int, 4).Value = RewardID
                    MyCommon.LRTsp.Parameters.Add("@Phase", SqlDbType.Int, 4).Value = 3
                    MyCommon.LRTsp.Parameters.Add("@PassThruRewardID", SqlDbType.Int, 4).Value = PassThruRewardID
                    MyCommon.LRTsp.Parameters.Add("@LSInterfaceID", SqlDbType.Int, 4).Value = LSInterfaceID
                    MyCommon.LRTsp.Parameters.Add("@ActionTypeID", SqlDbType.Int, 4).Value = ActionTypeID
                    MyCommon.LRTsp.Parameters.Add("@Required", SqlDbType.Bit).Value = IIf(RewardRequired, 1, 0)
                    MyCommon.LRTsp.Parameters.Add("@DeliverableID", SqlDbType.Int, 4).Direction = ParameterDirection.Output
                    MyCommon.LRTsp.ExecuteNonQuery()
                    DeliverableID = MyCommon.LRTsp.Parameters("@DeliverableID").Value
                    MyCommon.Close_LRTsp()
                Else
                    'It's an existing deliverable, so blow away any old Tiers and TierValues
                    MyCommon.QueryStr = "delete from PassThruTierValues where PTPKID in (select PKID from PassThrus where DeliverableID=@DeliverableID)"
                    MyCommon.DBParameters.Add("@DeliverableID", SqlDbType.Int).Value = DeliverableID
                    MyCommon.ExecuteNonQuery(DataBases.LogixRT)
                    MyCommon.QueryStr = "delete from PassThruTiers where PTPKID in (select PKID from PassThrus where DeliverableID=@DeliverableID" & ")"
                    MyCommon.DBParameters.Add("@DeliverableID", SqlDbType.Int).Value = DeliverableID
                    MyCommon.ExecuteNonQuery(DataBases.LogixRT)

                    ' update the required status of this deliverable
                    MyCommon.QueryStr = "update CPE_Deliverables with (RowLock) set Required= @RewardRequired, LastUpdate=getdate() where DeliverableID= @DeliverableID"
                    MyCommon.DBParameters.Add("@RewardRequired", SqlDbType.Bit).Value = IIf(RewardRequired, 1, 0)
                    MyCommon.DBParameters.Add("@DeliverableID", SqlDbType.Int).Value = DeliverableID
                    MyCommon.ExecuteNonQuery(DataBases.LogixRT)
                End If

                'Get the PKID of the pass-thru
                MyCommon.QueryStr = "select OutputID from CPE_Deliverables with (NoLock) where DeliverableID= @DeliverableID"
                MyCommon.DBParameters.Add("@DeliverableID", SqlDbType.Int).Value = DeliverableID
                rst = MyCommon.ExecuteQuery(DataBases.LogixRT)
                If rst.Rows.Count > 0 Then
                    PassThruPKID = MyCommon.NZ(rst.Rows(0).Item("OutputID"), 0)
                End If
                'Get the active languages, for use during tier creation
                MyCommon.QueryStr = "SELECT L.LanguageID, L.MSNetCode FROM Languages AS L " & _
                                    "WHERE L.LanguageID in (" & IIf(MultiLanguageEnabled, "SELECT TLV.LanguageID FROM TransLanguagesCF_UE AS TLV", DefaultLanguageID) & ") " & _
                                    "ORDER BY CASE WHEN L.LanguageID=" & DefaultLanguageID & " THEN 1 ELSE 2 END, L.LanguageID;"
                LanguagesDT = MyCommon.LRT_Select
                'Create the tiers
                For t = 1 To TierLevels
                    RewardValue = MyCommon.Extract_Val(GetCgiValue("t" & t & "_value"))
                    If RewardValue < 0 Then RewardValue = 0
                    If MultiLanguageEnabled And HasMLTag Then
                        'We need to make a tier record for each language
                        For Each lrow In LanguagesDT.Rows
                            MyCommon.QueryStr = "dbo.pa_CPE_AddPassThruRewardTiers"
                            MyCommon.Open_LRTsp()
                            MyCommon.LRTsp.Parameters.Add("@PTPKID", SqlDbType.Int, 4).Value = PassThruPKID
                            MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int, 4).Value = t
                            MyCommon.LRTsp.Parameters.Add("@Data", SqlDbType.NVarChar).Value = ""
                            MyCommon.LRTsp.Parameters.Add("@Value", SqlDbType.Decimal, 9).Value = RewardValue
                            MyCommon.LRTsp.Parameters.Add("@LanguageID", SqlDbType.Int, 4).Value = lrow.Item("LanguageID")
                            MyCommon.LRTsp.Parameters.Add("@TierPKID", SqlDbType.Int, 4).Direction = ParameterDirection.Output
                            MyCommon.LRTsp.ExecuteNonQuery()
                            TierPKID = MyCommon.LRTsp.Parameters("@TierPKID").Value
                            MyCommon.Close_LRTsp()
                        Next
                    Else
                        'Make a single tier record with a language of 0
                        MyCommon.QueryStr = "dbo.pa_CPE_AddPassThruRewardTiers"
                        MyCommon.Open_LRTsp()
                        MyCommon.LRTsp.Parameters.Add("@PTPKID", SqlDbType.Int, 4).Value = PassThruPKID
                        MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int, 4).Value = t
                        MyCommon.LRTsp.Parameters.Add("@Data", SqlDbType.NVarChar).Value = ""
                        MyCommon.LRTsp.Parameters.Add("@Value", SqlDbType.Decimal, 9).Value = RewardValue
                        MyCommon.LRTsp.Parameters.Add("@LanguageID", SqlDbType.Int, 4).Value = DefaultLanguageID
                        MyCommon.LRTsp.Parameters.Add("@TierPKID", SqlDbType.Int, 4).Direction = ParameterDirection.Output
                        MyCommon.LRTsp.ExecuteNonQuery()
                        TierPKID = MyCommon.LRTsp.Parameters("@TierPKID").Value
                        MyCommon.Close_LRTsp()
                    End If
                    'Populate the TierValues table
                    MyCommon.QueryStr = "select DataTemplate from PassThruRewards with (NoLock) where PassThruRewardID= @PassThruRewardID"
                    MyCommon.DBParameters.Add("@PassThruRewardID", SqlDbType.BigInt).Value = PassThruRewardID
                    rst2 = MyCommon.ExecuteQuery(DataBases.LogixRT)
                    If rst2.Rows.Count > 0 Then
                        DataTemplate = MyCommon.NZ(rst2.Rows(0).Item("DataTemplate"), "")
                        Tokens = ParseTokenValues(DataTemplate)
                        For i = 0 To Tokens.GetUpperBound(0)
                            TokenID = Tokens(i)
                            'Determine if the token is multilanguage
                            MyCommon.QueryStr = "select MultiLanguage,ParamDataTypeID,isnull(MaxLength,0) as MaxLength from PassThruPresTags with (NoLock) where PassThruPresTagID= @TokenID"
                            MyCommon.DBParameters.Add("@TokenID", SqlDbType.Int).Value = TokenID
                            rst3 = MyCommon.ExecuteQuery(DataBases.LogixRT)
                            If rst3.Rows.Count > 0 Then
                                If rst3.Rows(0).Item("MultiLanguage") Then
                                    MultiLanguagePresTag = True
                                Else
                                    MultiLanguagePresTag = False
                                End If
                            End If
                            Dim intMaxLength As Integer = IIf(Convert.ToInt32(rst3.Rows(0).Item("MaxLength")) > 0, Convert.ToInt32(rst3.Rows(0).Item("MaxLength")), 999)
                            'Insert the token values into PassThruTierValues
                            If MultiLanguageEnabled And MultiLanguagePresTag Then
                                If LanguagesDT.Rows.Count > 0 Then
                                    For Each lrow In LanguagesDT.Rows
                                        TokenValue = Request.QueryString("t" & t & "_token" & TokenID & "_" & MyCommon.NZ(lrow.Item("MSNetCode"), ""))
                                        TokenValue = Left(TokenValue, intMaxLength)
                                        If rst3.Rows(0).Item("ParamDataTypeID") = 4 Then
                                            TokenValue = Logix.TrimAll(TokenValue)
                                        End If
                                        MyCommon.QueryStr = "insert into PassThruTierValues (PassThruPresTagID, Value, LanguageID, PTPKID, TierLevel) " & _
                                                            "values ( @TokenID, @TokenValue, @LanguageID, @PassThruPKID, @TierLevel)"
                                        MyCommon.DBParameters.Add("@TokenID", SqlDbType.Int).Value = TokenID
                                        MyCommon.DBParameters.Add("@TokenValue", SqlDbType.NVarChar).Value = TokenValue
                                        MyCommon.DBParameters.Add("@LanguageID", SqlDbType.Int).Value = MyCommon.NZ(lrow.Item("LanguageID"), 1)
                                        MyCommon.DBParameters.Add("@PassThruPKID", SqlDbType.Int).Value = PassThruPKID
                                        MyCommon.DBParameters.Add("@TierLevel", SqlDbType.Int).Value = t
                                        MyCommon.ExecuteNonQuery(DataBases.LogixRT)
                                        'End If
                                    Next
                                End If
                            Else
                                TokenValue = Request.QueryString("t" & t & "_token" & TokenID)
                                TokenValue = Left(TokenValue, intMaxLength)
                                If rst3.Rows(0).Item("ParamDataTypeID") = 4 Then
                                    TokenValue = Logix.TrimAll(TokenValue)
                                End If



                                MyCommon.QueryStr = "insert into PassThruTierValues (PassThruPresTagID, Value, LanguageID, PTPKID, TierLevel) " & _
                                                    "values (@TokenID, @TokenValue, @LanguageID, @PassThruPKID, @TierLevel)"
                                MyCommon.DBParameters.Add("@TokenID", SqlDbType.Int).Value = TokenID
                                MyCommon.DBParameters.Add("@TokenValue", SqlDbType.NVarChar).Value = TokenValue
                                If rst3.Rows(0).Item("ParamDataTypeID") = 4 Then
                                    TokenValue = Logix.TrimAll(TokenValue)
                                End If
                                MyCommon.DBParameters.Add("@LanguageID", SqlDbType.Int).Value = IIf(MultiLanguagePresTag, DefaultLanguageID, 0)
                                MyCommon.DBParameters.Add("@PassThruPKID", SqlDbType.Int).Value = PassThruPKID
                                MyCommon.DBParameters.Add("@TierLevel", SqlDbType.Int).Value = t
                                MyCommon.ExecuteNonQuery(DataBases.LogixRT)

                            End If
                        Next

                        If MultiLanguageEnabled And HasMLTag Then
                            For Each lrow In LanguagesDT.Rows
                                'Now we look at the PassThruTierValues table and build up the Data column for PassThruTiers
                                MyCommon.QueryStr = "select PTPT.PassThruPresTagID, PTPT.MultiLanguage, PTTV.LanguageID, Replace(PTPT.ReplacementText, '^', PTTV.Value) as ReplacementText " & _
                                                    "from PassThruTierValues as PTTV with (NoLock) " & _
                                                    "inner join PassThruPresTags as PTPT with (NoLock) on PTPT.PassThruPresTagID=PTTV.PassThruPresTagID " & _
                                                    "where PTTV.PTPKID=(select PKID from PassThrus where DeliverableID= @DeliverableID ) and PTTV.TierLevel=@TierLevel " & _
                                                    " and LanguageID in (0, @LanguageID) " & _
                                                    " order by PassThruPresTagID;"

                                MyCommon.DBParameters.Add("@DeliverableID", SqlDbType.Int).Value = DeliverableID
                                MyCommon.DBParameters.Add("@LanguageID", SqlDbType.Int).Value = MyCommon.NZ(lrow.Item("LanguageID"), 1)
                                MyCommon.DBParameters.Add("@TierLevel", SqlDbType.Int).Value = t
                                rst3 = MyCommon.ExecuteQuery(DataBases.LogixRT)
                                TierString = DataTemplate
                                If rst3.Rows.Count > 0 Then
                                    For Each row In rst3.Rows
                                        TierString = TierString.Replace("^" & MyCommon.NZ(row.Item("PassThruPresTagID"), "") & ";", MyCommon.NZ(row.Item("ReplacementText"), ""))
                                    Next
                                End If
                                MyCommon.QueryStr = "select PKID from PassThruTiers where PTPKID=@PassThruPKID and TierLevel=@TierLevel and LanguageID=@LanguageID"
                                MyCommon.DBParameters.Add("@PassThruPKID", SqlDbType.Int).Value = PassThruPKID
                                MyCommon.DBParameters.Add("@LanguageID", SqlDbType.Int).Value = MyCommon.NZ(lrow.Item("LanguageID"), 1)
                                MyCommon.DBParameters.Add("@TierLevel", SqlDbType.Int).Value = t
                                rst3 = MyCommon.ExecuteQuery(DataBases.LogixRT)
                                If rst3.Rows.Count > 0 Then
                                    TierPKID = rst3.Rows(0).Item("PKID")
                                End If
                                MyCommon.QueryStr = "update PassThruTiers set Data=@Data where PKID=@PKID and LanguageID=@LanguageID"
                                MyCommon.DBParameters.Add("@Data", SqlDbType.NVarChar).Value = TierString
                                MyCommon.DBParameters.Add("@PKID", SqlDbType.Int).Value = TierPKID
                                MyCommon.DBParameters.Add("@LanguageID", SqlDbType.Int).Value = lrow.Item("LanguageID")
                                MyCommon.ExecuteNonQuery(DataBases.LogixRT)
                            Next
                        Else
                            'Now we look at the PassThruTierValues table and build up the Data column for PassThruTiers
                            MyCommon.QueryStr = "select PTPT.PassThruPresTagID, PTPT.MultiLanguage, PTTV.LanguageID, Replace(PTPT.ReplacementText, '^', PTTV.Value) as ReplacementText " & _
                                                "from PassThruTierValues as PTTV with (NoLock) " & _
                                                "inner join PassThruPresTags as PTPT with (NoLock) on PTPT.PassThruPresTagID=PTTV.PassThruPresTagID " & _
                                                "where PTTV.PTPKID=(select PKID from PassThrus where DeliverableID=@DeliverableID) and PTTV.TierLevel=@TierLevel " & _
                                                "order by PassThruPresTagID;"
                            MyCommon.DBParameters.Add("@DeliverableID", SqlDbType.Int).Value = DeliverableID
                            MyCommon.DBParameters.Add("@TierLevel", SqlDbType.Int).Value = t
                            rst3 = MyCommon.ExecuteQuery(DataBases.LogixRT)
                            TierString = DataTemplate
                            If rst3.Rows.Count > 0 Then
                                For Each row In rst3.Rows
                                    TierString = TierString.Replace("^" & MyCommon.NZ(row.Item("PassThruPresTagID"), "") & ";", MyCommon.NZ(row.Item("ReplacementText"), ""))
                                Next

                                MyCommon.QueryStr = "update PassThruTiers set Data=@Data where PKID=@PKID"
                                MyCommon.DBParameters.Add("@Data", SqlDbType.NVarChar).Value = TierString
                                MyCommon.DBParameters.Add("@PKID", SqlDbType.Int).Value = TierPKID
                                MyCommon.ExecuteNonQuery(DataBases.LogixRT)
                            End If
                        End If

                    End If
                Next
                MyCommon.QueryStr = "update CPE_Incentives set LastUpdate=getdate(), LastUpdatedByAdminID=@LastUpdatedByAdminID, StatusFlag=1 where IncentiveID=@IncentiveID"
                MyCommon.DBParameters.Add("@LastUpdatedByAdminID", SqlDbType.Int).Value = AdminUserID
                MyCommon.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = OfferID
                MyCommon.ExecuteNonQuery(DataBases.LogixRT)
                ResetOfferApprovalStatus(OfferID)

                MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Reward.editpassthru", LanguageID))
                CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")

            Else
                bIsErrorMsg = True
                If BadTagValues > 0 Then
                    infoMessage = Copient.PhraseLib.Lookup("error.invalidvalue", LanguageID)
                ElseIf ((DTLen + TVLen) > 2000) Then
                    infoMessage = Copient.PhraseLib.Lookup("ueoffer-rew-passthru.Exceeds2000Chars", LanguageID)
                End If
            End If
        End If
    End If

    'Update the templates permission if necessary
    If (Request.QueryString("save") <> "" AndAlso Request.QueryString("IsTemplate") = "IsTemplate") Then
        ' time to update the status bits for the templates
        Dim form_Disallow_Edit As Integer = 0
        If (Request.QueryString("Disallow_Edit") = "on") Then
            form_Disallow_Edit = 1
        End If
        MyCommon.QueryStr = "update CPE_Deliverables with (RowLock) set DisallowEdit=" & form_Disallow_Edit & " where DeliverableID=" & DeliverableID & ";"
        MyCommon.LRT_Execute()
    End If

    If (IsTemplate Or FromTemplate) Then
        MyCommon.QueryStr = "select DisallowEdit from CPE_Deliverables with (NoLock) where DeliverableID=" & DeliverableID & ";"
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
            Disallow_Edit = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit"), False)
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
    Dim strSubTitle As String = ""
    If PassThruRewardID = 11 Then
        strSubTitle = "term.giftwithpurchaserewards"
    ElseIf PassThruRewardID = 12 Then
        strSubTitle = "term.purchasewithpurchasereward"
    ElseIf PassThruRewardID = 13 Then
        strSubTitle = "term.proximitymessagereward"
    End If
    Send_HeadBegin("term.offer", strSubTitle, OfferID)
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send("<style type=""text/css"">")
    Send("  td:first-child {")
    Send("    width:50%;")
    Send("  }")
    Send("</style>")
    Send_Scripts()
    Send("<script type=""text/javascript"">")
    Send("function ChangeParentDocument() { ")
    If (EngineID = 3) Then
        Send("  if (opener != null) {")
        Send("    var newlocation = '/logix/web-offer-rew.aspx?OfferID=" & OfferID & "'; ")
        Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
        Send("  opener.location = '/logix/web-offer-rew.aspx?OfferID=" & OfferID & "'; ")
    ElseIf (EngineID = 5) Then
        Send("  if (opener != null) {")
        Send("    var newlocation = '/logix/email-offer-rew.aspx?OfferID=" & OfferID & "'; ")
        Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
        Send("  opener.location = '/logix/email-offer-rew.aspx?OfferID=" & OfferID & "'; ")
    ElseIf (EngineID = 6) Then
        Send("  if (opener != null) {")
        Send("    var newlocation = '/logix/CAM/CAM-offer-rew.aspx?OfferID=" & OfferID & "'; ")
        Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
        Send("  opener.location = '/logix/CAM/CAM-offer-rew.aspx?OfferID=" & OfferID & "'; ")
    ElseIf (EngineID = 9) Then
        If (CloseAfterSave) Then
            Send("  if (opener != null) {")
            Send("    var newlocation = 'UEoffer-rew.aspx?OfferID=" & OfferID & "'; ")
            Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
            Send("  opener.location = 'UEoffer-rew.aspx?OfferID=" & OfferID & "'; ")
        End If
    Else
        Send("  if (opener != null) {")
        Send("    var newlocation = '/logix/CPEoffer-rew.aspx?OfferID=" & OfferID & "'; ")
        Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
        Send("  opener.location = '/logix/CPEoffer-rew.aspx?OfferID=" & OfferID & "'; ")
    End If
    Send("} ")
    Send("} ")
    Send("} ")
    Sendb("</")
    Send("script>")
    Send_HeadEnd()

    If (IsTemplate) Then
        Send_BodyBegin(12)
    Else
        Send_BodyBegin(2)
    End If
    If (Logix.UserRoles.AccessOffers = False AndAlso Not IsTemplate) Then
        Send_Denied(2, "perm.offers-access")
        GoTo done
    End If
    If (Logix.UserRoles.AccessTemplates = False AndAlso IsTemplate) Then
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
<form action="UEoffer-rew-passthru.aspx" id="mainform" name="mainform">
  <div id="intro">
    <input type="hidden" id="OfferID" name="OfferID" value="<% Sendb(OfferID)%>" />
    <input type="hidden" id="RewardID" name="RewardID" value="<% Sendb(RewardID)%>" />
    <input type="hidden" id="DeliverableID" name="DeliverableID" value="<% Sendb(DeliverableID) %>" />
    <input type="hidden" id="PassThruRewardID" name="PassThruRewardID" value="<% Sendb(PassThruRewardID) %>" />
    <input type="hidden" id="Phase" name="Phase" value="<% Sendb(Phase) %>" />
    <input type="hidden" id="ROID" name="ROID" value="<%Sendb(TpROID) %>" />
    <input type="hidden" id="TP" name="TP" value="<%Sendb(TouchPoint) %>" />
    <input type="hidden" id="IsTemplate" name="IsTemplate" value="<% Sendb(IsTemplateVal)%>" />
    <input type="hidden" id="EngineID" name="EngineID" value="<%Sendb(EngineID) %>" />
    <input type="hidden" id="LSInterfaceID" name="LSInterfaceID" value="<%Sendb(LSInterfaceID) %>" />
    <input type="hidden" id="ActionTypeID" name="ActionTypeID" value="<%Sendb(ActionTypeID) %>" />
    <%
      If (IsTemplate) Then
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & " " & PassThruRewardName & "</h1>")
      Else
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & PassThruRewardName & "</h1>")
      End If
    %>
    <div id="controls">
      <%
        If (IsTemplate) Then
          Send("<span class=""temp"">")
          Send("  <input type=""checkbox"" class=""tempcheck"" id=""Disallow_Edit"" name=""Disallow_Edit""" & IIf(Disallow_Edit, " checked=""checked""", "") & " />")
          Send("  <label for=""Disallow_Edit"">" & Copient.PhraseLib.Lookup("term.locked", LanguageID) & "</label>")
          Send("</span>")
          End If
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
    <%
      If (infoMessage <> "" And bIsErrorMsg) Then
                Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
            ElseIf (infoMessage <> "" And CloseAfterSave = False) Then
        Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
      ElseIf (infoMessage <> "") Then
        Send("<div id=""infobar"" class=""green-background"">" & infoMessage & "</div>")
      End If
    %>
    <div id="column">
      <%
        For t = 1 To TierLevels
          Send("<div class=""box"" id=""data" & IIf(TierLevels > 1, t, "") & """>")
          Send("  <h2>")
          Send("    <span>")
          If TierLevels > 1 Then
            Send("      " & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & " " & StrConv(Copient.PhraseLib.Lookup("term.data", LanguageID), VbStrConv.Lowercase))
          Else
            Send("      " & Copient.PhraseLib.Lookup("term.data", LanguageID))
          End If
          Send("    </span>")
          Send("  </h2>")
          MyCommon.QueryStr = "select PTR.Presentation, PTR.PresentationPhraseID, PTR.DataTemplate from PassThruRewards as PTR with (NoLock) " & _
                              "where PassThruRewardID=@PassThruRewardID"
          
          MyCommon.DBParameters.Add("@PassThruRewardID", SqlDbType.BigInt).Value = PassThruRewardID
          rst = MyCommon.ExecuteQuery(DataBases.LogixRT)
          If rst.Rows.Count > 0 Then
            If IsDBNull(rst.Rows(0).Item("PresentationPhraseID")) Then
              PresString = MyCommon.NZ(rst.Rows(0).Item("Presentation"), "")
            Else
              PresString = Copient.PhraseLib.Lookup(rst.Rows(0).Item("PresentationPhraseID"), LanguageID)
            End If
            If DeliverableID > 0 Then
              MyCommon.QueryStr = "select PTT.TierLevel, PTT.PKID, PTT.PTPKID, PTT.TierLevel, PTT.Data from PassThruTiers as PTT with (NoLock) " & _
                                  "left join PassThrus as PT on PT.PKID=PTT.PTPKID " & _
                                  "where PT.PassThruRewardID=@PassThruRewardID and PT.DeliverableID=@DeliverableID and PTT.TierLevel=@t"
              
              MyCommon.DBParameters.Add("@PassThruRewardID", SqlDbType.BigInt).Value = PassThruRewardID
              MyCommon.DBParameters.Add("@DeliverableID", SqlDbType.Int).Value = DeliverableID
              MyCommon.DBParameters.Add("@t", SqlDbType.Int).Value = t
              rst2 = MyCommon.ExecuteQuery(DataBases.LogixRT)
              If rst2.Rows.Count > 0 Then
                Send(InterpretPresentation(MyCommon, DeliverableID, t, MyCommon.NZ(rst2.Rows(0).Item("PTPKID"), ""), PresString, MyCommon.NZ(rst2.Rows(0).Item("Data"), ""), MultiLanguageEnabled, DefaultLanguageID) & "<br />")
              Else
                Send(InterpretPresentation(MyCommon, DeliverableID, t, 0, PresString, "", MultiLanguageEnabled, DefaultLanguageID) & "<br />")
              End If
            Else
              Send(InterpretPresentation(MyCommon, 0, t, 0, PresString, "", MultiLanguageEnabled, DefaultLanguageID) & "<br />")
            End If
          End If
          Send("  <hr class=""hidden"" />")
          Send("</div>")
        Next
      %>
      </div>
  </div>
</form>

<script type="text/javascript">
<% If (CloseAfterSave) Then %>
  window.close();
<% End If %>
</script>

<script runat="server">
  Function ParseTokenValues(ByVal DataTemplate As String) As String()
    Dim Tokens(-1) As String
    Dim StartIndex As Integer = 0
    Dim EndIndex As Integer = 0
    Dim TokenID As String = ""
    
    StartIndex = DataTemplate.IndexOf("^", 0)
    While StartIndex > -1
      EndIndex = DataTemplate.IndexOf(";", StartIndex)
      If EndIndex > -1 Then
        TokenID = DataTemplate.Substring(StartIndex + 1, (EndIndex - StartIndex) - 1)
        If TokenID.Trim <> "" Then
          ReDim Preserve Tokens(Tokens.Length)
          Tokens(Tokens.Length - 1) = TokenID
        End If
      End If
      StartIndex = DataTemplate.IndexOf("^", EndIndex)
    End While
    
    Return Tokens
  End Function
  Function InterpretPresentation(ByRef MyCommon As Copient.CommonInc, _
                                 ByVal DeliverableID As Integer, _
                                 ByVal TierLevel As Integer, _
                                 ByVal PTPKID As Integer, _
                                 ByVal Presentation As String, _
                                 ByVal Data As String, _
                                 ByVal MultiLanguageEnabled As Boolean, _
                                 ByVal DefaultLanguageID As Integer) As String
    Dim dt As System.Data.DataTable
    Dim dt2 As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim i As Integer = 0
    Dim Localization As Copient.Localization
    
    Dim StartIndex As Integer = 0
    Dim EndIndex As Integer = 0
    Dim TokenLength As String = ""
    Dim TokenID As String = ""
    Dim Tokens(-1) As String
    Dim CurrentValue As String = ""
    Dim ReplacementText As String = ""
    Dim FinalString As String = Presentation
    Dim sourceDB As String = ""
    Dim sourceDT As System.Data.DataTable
    Dim sourceRow As System.Data.DataRow
    Dim EXInstalled As Boolean = False
    Dim EXAccessible As Boolean = False
    Dim MLI As New Copient.Localization.MultiLanguageRec
    
    Localization = New Copient.Localization(MyCommon)
    
    If (System.Environment.GetEnvironmentVariable("LEXSERVER") <> "") AndAlso (System.Environment.GetEnvironmentVariable("LEXDATABASE") <> "") Then
      EXInstalled = True
    End If
    
    MultiLanguageEnabled = IIf(MyCommon.Fetch_SystemOption(124) = "1", True, False)
    
    If EXInstalled Then
      Try
        MyCommon.QueryStr = "select database_id from master.sys.databases where database_id=db_id();"
        dt = MyCommon.ExecuteQuery(DataBases.LogixEX)
        If dt.Rows.Count > 0 Then
          EXAccessible = True
        Else
          EXAccessible = False
        End If
      Catch ex As Exception
        EXAccessible = False
      End Try
    End If
    Tokens = ParseTokenValues(Presentation)
    
    Dim RedemptionPresTag As Integer = 0
    Dim DescriptionPresTag As Integer = 0
    Dim MaximumSelectionsPresTag As Integer = 0
    Dim tempTable As DataTable
    Dim OfferID As Long = MyCommon.NZ(Request.QueryString("OfferID"), 0)
    Dim strParamName As String
    Dim strPhrase As String
    For i = 0 To Tokens.GetUpperBound(0)
      TokenID = Tokens(i)
      
      'Next, look up the token and interpret it
      MyCommon.QueryStr = "select ReplacementText, TokenValueSelector, MaxLength, MultiLanguage, ParamName, PhraseID " & _
                          "from PassThruPresTags as PTPT with (NoLock) " & _
                          "where PassThruPresTagID=@TokenID"
      MyCommon.DBParameters.Add("@TokenID", SqlDbType.Int).Value = TokenID
      dt = MyCommon.ExecuteQuery(DataBases.LogixRT)
      If dt.Rows.Count > 0 Then
        strParamName = MyCommon.NZ(dt.Rows(0).Item("ParamName"), "")
        strPhrase = Copient.PhraseLib.Lookup(MyCommon.NZ(dt.Rows(0).Item("PhraseID"), 0), LanguageID)
        
        If (Request.QueryString("t" & TierLevel & "_token" & TokenID) <> "") Then
          CurrentValue = Request.QueryString("t" & TierLevel & "_token" & TokenID)
                
        ElseIf DeliverableID > 0 Then 'Get the value from the PassThruTierValues table
          MyCommon.QueryStr = "select Value from PassThruTierValues with (NoLock) " & _
                              "where PTPKID in (select PKID from PassThrus where DeliverableID=" & DeliverableID & ") " & _
                              " and TierLevel=@TierLevel" & _
                              " and PassThruPresTagID=@TokenID" & _
                              " and LanguageID in (0,@DefaultLanguageID) " & _
                              "order by LanguageID desc;"
          MyCommon.DBParameters.Add("@TierLevel", SqlDbType.Int).Value = TierLevel
          MyCommon.DBParameters.Add("@TokenID", SqlDbType.Int).Value = TokenID
          MyCommon.DBParameters.Add("@DefaultLanguageID", SqlDbType.Int).Value = DefaultLanguageID
          dt2 = MyCommon.ExecuteQuery(DataBases.LogixRT)
          If dt2.Rows.Count > 0 Then
            CurrentValue = MyCommon.NZ(dt2.Rows(0).Item("Value"), "")
          Else
            CurrentValue = "" '' If element is missing in PassThruTierValues table's defined XML then set the CurrentValue to Empty
          End If
        Else
          CurrentValue = ""
        End If

        MyCommon.QueryStr = "select PassThruPresTagID, ParamName from PassThruPresTags where ParamName in ('Redemption Text','Description (tokenized)','Maximum Selections');"
        tempTable = MyCommon.ExecuteQuery(DataBases.LogixRT)
        If tempTable.Rows.Count > 0 Then
          For Each row In tempTable.Rows
            If MyCommon.NZ(row.Item("ParamName"), "") = "Redemption Text" Then
              RedemptionPresTag = MyCommon.NZ(row.Item("PassThruPresTagID"), 0)
            ElseIf MyCommon.NZ(row.Item("ParamName"), "") = "Description (tokenized)" Then
              DescriptionPresTag = MyCommon.NZ(row.Item("PassThruPresTagID"), 0)
            ElseIf MyCommon.NZ(row.Item("ParamName"), "") = "Maximum Selections" Then
              MaximumSelectionsPresTag = MyCommon.NZ(row.Item("PassThruPresTagID"), 0)
            End If
          Next
        End If
        
        If TokenID = RedemptionPresTag Then
          'This is the Redemption Text token, so we must do some special calculations for location text
          MyCommon.QueryStr = "select O.LocationGroupID, LG.Name, LG.PhraseID from OfferLocations as O with (NoLock) " & _
                              "left join LocationGroups as LG with (NoLock) on LG.LocationGroupID=O.LocationGroupID " & _
                              "where Excluded=0 and O.Deleted=0 and O.OfferID=@OfferID order by LG.Name;"
          MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
          tempTable = MyCommon.ExecuteQuery(DataBases.LogixRT)
          'All Locations are selected for this offer
          If (tempTable.Rows.Count > 0) AndAlso MyCommon.NZ(tempTable.Rows(0).Item("Name"), "") = "All Locations" Then
            CurrentValue = "Redeemable at Any Location"
          End If
          MyCommon.QueryStr = "select Address1, Address2, City, State, Zip, ExtLocationCode from CPE_IncentiveLocationsView as c " & _
                              "inner join Locations as l on l.LocationID = c.LocationID " & _
                              "where incentiveid =&OfferID"
          MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
          tempTable = MyCommon.ExecuteQuery(DataBases.LogixRT)
          
          If tempTable.Rows.Count > 1 AndAlso CurrentValue <> "Redeemable at Any Location" Then
            CurrentValue = "Redeemable at Participating Locations"
          ElseIf tempTable.Rows.Count = 1 AndAlso CurrentValue <> "Redeemable at Any Location" Then
            CurrentValue = "Location " & MyCommon.NZ(tempTable.Rows(0).Item("ExtLocationCode"), "") & ": " & _
                        MyCommon.NZ(tempTable.Rows(0).Item("Address1"), "") & _
                        MyCommon.NZ(tempTable.Rows(0).Item("Address2"), "") & ", " & _
                        MyCommon.NZ(tempTable.Rows(0).Item("City"), "") & ", " & _
                        MyCommon.NZ(tempTable.Rows(0).Item("State"), "") & ", " & _
                        MyCommon.NZ(tempTable.Rows(0).Item("Zip"), "")
          Else
            'offer valid at no locations
          End If
        ElseIf TokenID = DescriptionPresTag And CurrentValue = "" Then
          'This is the "AmountAway" description token, so set the default value
          CurrentValue = "You are [AmountAway] from your reward."
        ElseIf TokenID = MaximumSelectionsPresTag And CurrentValue = "" Then
          CurrentValue = ""
        End If

        'Build the element
        If Not MyCommon.NZ(dt.Rows(0).Item("TokenValueSelector"), False) Then
          CurrentValue = Replace(CurrentValue, Chr(34), "&quot;")
          If MyCommon.NZ(dt.Rows(0).Item("MultiLanguage"), False) Then
            MLI.MLTableName = "PassThruTierValues"
            MLI.MLColumnName = "Value"
            MLI.ItemID = TokenID
            MLI.ItemID2 = PTPKID
            MLI.ItemID3 = TierLevel
            MLI.MLIdentifierName = "PassThruPresTagID"
            MLI.MLIdentifierName2 = "PTPKID"
            MLI.MLIdentifierName3 = "TierLevel"
            MLI.StandardValue = CurrentValue
            MLI.InputName = "t" & TierLevel & "_token" & TokenID
            MLI.InputID = "t" & TierLevel & "_token" & TokenID
            MLI.InputType = "text"
            MLI.CSSStyle = "width:290px;"
            MLI.MaxLength = MyCommon.NZ(dt.Rows(0).Item("MaxLength"), 0)
            ReplacementText = Localization.SendTranslationInputs(MyCommon, MLI, Request.QueryString, 9)
          Else
            ReplacementText = "<input type=""text"" id=""t" & TierLevel & "_token" & TokenID & """ name=""t" & TierLevel & "_token" & TokenID & """ value=""" & CurrentValue & """" & IIf(MyCommon.NZ(dt.Rows(0).Item("MaxLength"), 0) > 0, " maxlength=""" & MyCommon.NZ(dt.Rows(0).Item("MaxLength"), 0) & """", "") & " />"
          End If
        Else
          ReplacementText = "<select id=""t" & TierLevel & "_token" & TokenID & """ name=""t" & TierLevel & "_token" & TokenID & """>"
          MyCommon.QueryStr = "select * from PassThruTokenValues as PTTV with (NoLock) " & _
                              "where PassThruPresTagID=@TokenID order by DisplayOrder;"
          MyCommon.DBParameters.Add("@TokenID", SqlDbType.Int).Value = TokenID
          dt2 = MyCommon.ExecuteQuery(DataBases.LogixRT)
          
          For Each row In dt2.Rows
            If (MyCommon.NZ(row.Item("SourceDatabase"), "") <> "" AndAlso MyCommon.NZ(row.Item("SourceTable"), "") <> "" AndAlso MyCommon.NZ(row.Item("OptionDescription"), "") <> "" AndAlso MyCommon.NZ(row.Item("OptionValue"), "") <> "") Then
              sourceDB = MyCommon.NZ(row.Item("SourceDatabase"), "")
              'Option value(s) derive not from PassThruTokenValues but from some other table (the SourceTable)
              MyCommon.QueryStr = "select" & _
                                  IIf(MyCommon.NZ(row.Item("SourceLimit"), 0) > 0, " top " & MyCommon.NZ(row.Item("SourceLimit"), 0), "") & " " & _
                                  MyCommon.NZ(row.Item("OptionDescription"), "") & " as OptionDescription, " & _
                                  MyCommon.NZ(row.Item("OptionValue"), "") & " as OptionValue " & _
                                  "from " & MyCommon.NZ(row.Item("SourceTable"), "") & " with (NoLock)" & _
                                  IIf(MyCommon.NZ(row.Item("SourceWhereClause"), "") <> "", " WHERE " & MyCommon.NZ(row.Item("SourceWhereClause"), ""), "") & _
                                  IIf(MyCommon.NZ(row.Item("SourceOrderByClause"), "") <> "", " ORDER BY " & MyCommon.NZ(row.Item("SourceOrderByClause"), ""), "") & _
                                  ";"
              If sourceDB = "RT" Then
                sourceDT = MyCommon.LRT_Select
              ElseIf sourceDB = "XS" Then
                sourceDT = MyCommon.LXS_Select
              ElseIf sourceDB = "WH" Then
                sourceDT = MyCommon.LWH_Select
              ElseIf sourceDB = "EX" AndAlso EXInstalled AndAlso EXAccessible Then
                sourceDT = MyCommon.LEX_Select
              Else
                sourceDT = Nothing
              End If
              If sourceDT Is Nothing Then
                ReplacementText &= "<option disabled=""disabled"">" & Copient.PhraseLib.Lookup("term.error", LanguageID) & "</option>"
              Else
                If sourceDT.Rows.Count > 0 Then
                  For Each sourceRow In sourceDT.Rows
                    ReplacementText &= "<option value=""" & MyCommon.NZ(sourceRow.Item("OptionValue"), "").ToString & """" & IIf(CurrentValue = MyCommon.NZ(sourceRow.Item("OptionValue"), "").ToString, " selected=""selected""", "") & ">"
                    ReplacementText &= MyCommon.NZ(sourceRow.Item("OptionDescription"), "&nbsp;")
                    ReplacementText &= "</option>"
                  Next
                End If
              End If
            Else
              'Conventional option value
              ReplacementText &= "<option value=""" & MyCommon.NZ(row.Item("OptionValue"), "") & """" & IIf(CurrentValue = MyCommon.NZ(row.Item("OptionValue"), ""), " selected=""selected""", "") & ">"
              If IsDBNull(row.Item("PhraseID")) Then
                ReplacementText &= MyCommon.NZ(row.Item("OptionDescription"), "&nbsp;")
              Else
                ReplacementText &= Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID)
              End If
              ReplacementText &= "</option>"
            End If
          Next
          ReplacementText &= "</select>"
        End If
        FinalString = Replace(FinalString, "^" & TokenID & ";", ReplacementText)
        If strPhrase <> "" Then
          FinalString = Regex.Replace(FinalString, strParamName, strPhrase, RegexOptions.IgnoreCase)
        End If
      End If
    Next
    Try
      MyCommon.Close_LogixEX()
    Catch ex As Exception
      'EX failed
    End Try
    
    Return FinalString
  End Function
</script>
<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd("mainform", "t1_tag1")
  Logix = Nothing
  MyCommon = Nothing
%>
