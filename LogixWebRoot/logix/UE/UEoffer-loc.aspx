<%@ Page Language="vb" Debug="true" CodeFile="ue-cb.vb" Inherits="UECB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
    ' *****************************************************************************
    ' * FILENAME: offer-loc.aspx 
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
    Dim rst As DataTable
    Dim rst3 As DataTable
    Dim rstSelected As DataTable
    Dim rstExcluded As DataTable
    Dim row As DataRow
    Dim row2 As DataRow
    Dim OfferID As Long
    Dim PromoID As String
    Dim Name As String = ""
    Dim IsTemplate As Boolean = False
    Dim FromTemplate As Boolean = False
    'Dim Disallow_Stores As Boolean = False
    'Dim Disallow_Terminals As Boolean = False
    'Dim PaddingValue As Integer = 0
    'Dim BtnPaddingTop As Integer = 30
    Dim EngineID As Integer = 0
    Dim EngineSubTypeID As Integer = 0
    Dim SelectSize As Integer = 6
    Dim infoMessage As String = ""
    Dim modMessage As String = ""
    Dim Handheld As Boolean = False
    Dim BannersEnabled As Boolean = False
    Dim NameValue As String = ""
    Dim PrevNameValue As String = ""
    Dim BannerList As New ArrayList(10)
    Dim BannerListStr As String = ""
    Dim LoopCtr As Integer = 0
    Dim AllBannersPermission As Boolean = False
    Dim StatusCode As Copient.LogixInc.STATUS_FLAGS = Copient.LogixInc.STATUS_FLAGS.STATUS_UNKNOWN
    Dim StatusText As String = ""
    Dim Popup As Boolean = False
    Dim FocusType As Integer = 0
    Dim LocationGroupID As Integer = 0
    Dim ExcludedStores As String
    Dim TierLevels As Integer = 1
    Dim CloseAfterSave As Boolean = False
    Dim RECORD_LIMIT As Integer = GroupRecordLimit '500


    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If

    Response.Expires = 0
    MyCommon.AppName = "offer-loc.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)

    OfferID = Request.QueryString("OfferID")

    BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
    Popup = IIf((Request.QueryString("Popup") <> "") AndAlso (Request.QueryString("Popup") <> "0"), True, False)
    FocusType = IIf(Request.QueryString("Focus") = "1", 1, 0)

    'Set the EngineID to 9 ... since UE is the only engine that this page will handle
    EngineID = 9
    'the only engine sub type for UE is zero
    EngineSubTypeID = 0

    If (OfferID = 0) Then
        Response.Status = "301 Moved Permanently"
        Response.AddHeader("Location", "offer-gen.aspx?new=New")
        GoTo done
    End If


    'This would redirect to summary page if offer is awaiting deployment or if offer is awaiting recommendations.
    CheckIfValidOffer(MyCommon, OfferID)

    Dim isTranslatedOffer As Boolean =MyCommon.IsTranslatedUEOffer(OfferID,  MyCommon)
    Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249)="1",True,False)

    Dim bEnableAdditionalLockoutRestrictionsOnOffers As Boolean = IIf(MyCommon.Fetch_SystemOption(260) = "1", True, False)
    Dim bOfferEditable As Boolean = MyCommon.IsOfferEditablePastLockOutPeriod(Logix.UserRoles.EditOfferPastLockoutPeriod, MyCommon, OfferID)

    'Get the tier level
    MyCommon.QueryStr = "select TierLevels from CPE_RewardOptions with (NoLock) where TouchResponse=0 and IncentiveID=" & OfferID & ";"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
        TierLevels = rst.Rows(0).Item("TierLevels")
    End If

    MyCommon.QueryStr = "select IncentiveName as Name, IsTemplate, FromTemplate,buy.ExternalBuyerId as BuyerID from CPE_Incentives CPE with (NoLock) " & _
                          "left outer join Buyers as buy with (nolock) on buy.BuyerId= CPE.BuyerId " & _
                          "where IncentiveID=" & OfferID & " and Deleted=0;"

    rst = MyCommon.LRT_Select()
    For Each row In rst.Rows
        If (MyCommon.Fetch_UE_SystemOption(168) = "1" AndAlso MyCommon.NZ(row.Item("BuyerID"), "") <> "") Then
            Name = "Buyer " + row.Item("BuyerID").ToString() + " - " + MyCommon.NZ(row.Item("Name"), "").ToString()
        Else
            Name = MyCommon.NZ(row.Item("Name"), Copient.PhraseLib.Lookup("term.unknown", LanguageID))
        End If
        'Name = MyCommon.NZ(row.Item("Name"), "")
        IsTemplate = row.Item("IsTemplate")
        FromTemplate = row.Item("FromTemplate")
    Next

    If (Request.QueryString("stores-add1") <> "" And Request.QueryString("sgroups-available") <> "") Then
        Dim i As String = Request.QueryString("sgroups-available")
        Dim a() As String
        Dim j As Integer
        a = i.Split(",")
        MyCommon.QueryStr = "select PKID from OfferLocations with (NoLock) where LocationGroupID=1 and OfferID=" & OfferID
        rst = MyCommon.LRT_Select

        If (rst.Rows.Count > 0) Then
            ' All cardholders is already in the selected box, so lose it
            MyCommon.QueryStr = "update OfferLocations with (RowLock) set Deleted=1, StatusFlag=2, TCRMAStatusFlag=3 where LocationGroupID=1 and OfferID=" & OfferID
            MyCommon.LRT_Execute()
            MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=1, LastUpdate=getdate() where IncentiveID=" & OfferID
            MyCommon.LRT_Execute()
            ResetOfferApprovalStatus(OfferID)
        End If

        If (a(0) = 1) Then
            MyCommon.QueryStr = "delete from OfferLocations with (RowLock) where Deleted=1 and OfferID=" & OfferID
            MyCommon.LRT_Execute()
            MyCommon.QueryStr = "update OfferLocations with (RowLock) set Deleted=1, TCRMAStatusFlag=3 where OfferID=" & OfferID
            MyCommon.LRT_Execute()

            MyCommon.QueryStr = "dbo.pt_OfferLocations_Insert"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
            MyCommon.LRTsp.Parameters.Add("@LocationGroupID", SqlDbType.BigInt).Value = a(0)
            MyCommon.LRTsp.Parameters.Add("@Excluded", SqlDbType.Bit).Value = 0
            MyCommon.LRTsp.Parameters.Add("@TCRMAStatusFlag", SqlDbType.Int).Value = 3
            MyCommon.LRTsp.Parameters.Add("@PKID", SqlDbType.BigInt).Direction = ParameterDirection.Output
            MyCommon.LRTsp.ExecuteNonQuery()
            MyCommon.Close_LRTsp()

            MyCommon.QueryStr = "update OfferLocations with (RowLock) set StatusFlag=2,TCRMAStatusFlag=3 where OfferID=" & OfferID & " and LocationGroupID=" & a(0)
            MyCommon.LRT_Execute()
            MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where IncentiveID=" & OfferID

            MyCommon.LRT_Execute()
            ResetOfferApprovalStatus(OfferID)
            MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-addstore", LanguageID))
        Else
            For j = 0 To a.GetUpperBound(0)
                MyCommon.QueryStr = "dbo.pt_OfferLocations_Insert"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
                MyCommon.LRTsp.Parameters.Add("@LocationGroupID", SqlDbType.BigInt).Value = a(j)
                MyCommon.LRTsp.Parameters.Add("@Excluded", SqlDbType.Bit).Value = 0
                MyCommon.LRTsp.Parameters.Add("@TCRMAStatusFlag", SqlDbType.Int).Value = 3
                MyCommon.LRTsp.Parameters.Add("@PKID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                MyCommon.LRTsp.ExecuteNonQuery()
                MyCommon.Close_LRTsp()

                MyCommon.QueryStr = "update OfferLocations with (RowLock) set StatusFlag=2,TCRMAStatusFlag=3 where OfferID=" & OfferID & " and LocationGroupID=" & a(j)
                MyCommon.LRT_Execute()

                MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=1, LastUpdate=getdate() where IncentiveID=" & OfferID
                MyCommon.LRT_Execute()
                ResetOfferApprovalStatus(OfferID)
                'MyCommon.QueryStr = "insert into OfferLocations (OfferID,LocationGroupID,LastUpdate) values(" & OfferID & "," & a(j) & "," & "getdate())"
                'MyCommon.LRT_Execute()            
            Next

            ' clear out the excluded during an add 
            MyCommon.QueryStr = "select PKID from OfferLocations with (NoLock) where OfferID=" & OfferID & " and Excluded=1"
            rst = MyCommon.LRT_Select
            For Each row In rst.Rows
                MyCommon.QueryStr = "dbo.pt_OfferLocations_Delete"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@PKID", SqlDbType.BigInt).Value = row.Item("PKID")
                MyCommon.LRTsp.Parameters.Add("@TCRMAStatusFlag", SqlDbType.Int).Value = 3
                MyCommon.LRTsp.ExecuteNonQuery()
                MyCommon.Close_LRTsp()
            Next
            MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-addstore", LanguageID))
        End If

    ElseIf (Request.QueryString("stores-rem1") <> "" And Request.QueryString("sgroups-select") <> "") Then
        MyCommon.QueryStr = "select PKID from OfferLocations with (NoLock) where OfferID=" & OfferID & " and LocationGroupID in (" & Request.QueryString("sgroups-select") & ")"
        rst = MyCommon.LRT_Select

        For Each row In rst.Rows
            MyCommon.QueryStr = "dbo.pt_OfferLocations_Delete"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@PKID", SqlDbType.BigInt).Value = row.Item("PKID")
            MyCommon.LRTsp.Parameters.Add("@TCRMAStatusFlag", SqlDbType.Int).Value = 3
            MyCommon.LRTsp.ExecuteNonQuery()
            MyCommon.Close_LRTsp()
        Next

        MyCommon.QueryStr = "update OfferLocations with (RowLock) set StatusFlag=2,TCRMAStatusFlag=3 where OfferID=" & OfferID & " and LocationGroupID in (" & Request.QueryString("sgroups-select") & ")"
        MyCommon.LRT_Execute()
        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=1, lastupdate=getdate() where IncentiveID=" & OfferID
        MyCommon.LRT_Execute()
        ResetOfferApprovalStatus(OfferID)
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-removestore", LanguageID))

    ElseIf (Request.QueryString("stores-add2") <> "" And Request.QueryString("sgroups-available") <> "" And Request.QueryString("sgroups-available") <> "1") Then
        ' Someone clicked excluded, so force all locations into the selected box
        MyCommon.QueryStr = "dbo.pt_OfferLocations_Insert"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
        MyCommon.LRTsp.Parameters.Add("@LocationGroupID", SqlDbType.BigInt).Value = 1
        MyCommon.LRTsp.Parameters.Add("@Excluded", SqlDbType.Bit).Value = 0
        MyCommon.LRTsp.Parameters.Add("@TCRMAStatusFlag", SqlDbType.Int).Value = 3
        MyCommon.LRTsp.Parameters.Add("@PKID", SqlDbType.BigInt).Direction = ParameterDirection.Output
        MyCommon.LRTsp.ExecuteNonQuery()
        MyCommon.Close_LRTsp()

        MyCommon.QueryStr = "select PKID from OfferLocations with (NoLock) where Excluded<>1 and Deleted=0 and OfferID=" & OfferID
        rst = MyCommon.LRT_Select

        If (rst.Rows.Count > 0) Then
            ' All cardholders is already in the selected box, so lose it
            MyCommon.QueryStr = "update OfferLocations with (RowLock) set Deleted=1, StatusFlag=2, TCRMAStatusFlag=3 where LocationGroupID<>1 and OfferID=" & OfferID
            MyCommon.LRT_Execute()
            MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=1, LastUpdate=getdate() where IncentiveID=" & OfferID
            MyCommon.LRT_Execute()
            ResetOfferApprovalStatus(OfferID)
        End If

        'Response.Write("Adding " & Request.QueryString("sgroups-avail") & " to excluded")
        Dim i As String = Request.QueryString("sgroups-available")
        Dim a() As String
        Dim j As Integer = 0
        Dim locGroupID As String = ""
        a = i.Split(",")

        'find the first group that's not All Locations(i.e. LocationGroupdID= "1") and use it
        For j = 0 To a.GetUpperBound(0)
            If (a(j) <> "1") Then
                locGroupID = a(j)
                Exit For
            End If
        Next

        If (locGroupID <> "") Then
            MyCommon.QueryStr = "dbo.pt_OfferLocations_Insert"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
            MyCommon.LRTsp.Parameters.Add("@LocationGroupID", SqlDbType.BigInt).Value = locGroupID
            MyCommon.LRTsp.Parameters.Add("@Excluded", SqlDbType.Bit).Value = 1
            MyCommon.LRTsp.Parameters.Add("@TCRMAStatusFlag", SqlDbType.Int).Value = 3
            MyCommon.LRTsp.Parameters.Add("@PKID", SqlDbType.BigInt).Direction = ParameterDirection.Output
            MyCommon.LRTsp.ExecuteNonQuery()
            MyCommon.Close_LRTsp()
            MyCommon.QueryStr = "update OfferLocations with (RowLock) set StatusFlag=2, TCRMAStatusFlag=3 where OfferID=" & OfferID & " and LocationGroupID=" & a(j)
            MyCommon.LRT_Execute()
            MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where IncentiveID=" & OfferID
            MyCommon.LRT_Execute()
            ResetOfferApprovalStatus(OfferID)
        End If
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-excludestore", LanguageID))
    ElseIf (Request.QueryString("stores-rem2") <> "" And Request.QueryString("sgroups-exclude") <> "") Then
        MyCommon.QueryStr = "select PKID from OfferLocations where OfferID=" & OfferID & " and LocationGroupID=" & Request.QueryString("sgroups-exclude")
        rst = MyCommon.LRT_Select

        For Each row In rst.Rows
            MyCommon.QueryStr = "dbo.pt_OfferLocations_Delete"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@PKID", SqlDbType.BigInt).Value = row.Item("PKID")
            MyCommon.LRTsp.Parameters.Add("@TCRMAStatusFlag", SqlDbType.Int).Value = 3
            MyCommon.LRTsp.ExecuteNonQuery()
            MyCommon.Close_LRTsp()
        Next

        MyCommon.QueryStr = "update OfferLocations with (RowLock) set StatusFlag=2, TCRMAStatusFlag=3 where OfferID=" & OfferID & " and LocationGroupID=" & Request.QueryString("sgroups-exclude")
        MyCommon.LRT_Execute()
        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where IncentiveID=" & OfferID
        MyCommon.LRT_Execute()
        ResetOfferApprovalStatus(OfferID)
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-unexcludestore", LanguageID))
    ElseIf (Request.QueryString("terminals-add1") <> "" And Request.QueryString("terminals-available") <> "") Then
        ' Response.Write("terminals-available:" & Request.QueryString("terminals-available"))
        Dim i As String = Request.QueryString("terminals-available")
        Dim a() As String
        Dim j As Integer

        ' First delete the any terminals, if present
        MyCommon.QueryStr = "delete from OfferTerminals with (RowLock) where OfferID=" & OfferID & " and Excluded=0 and TerminalTypeID in " & _
                            "(select TerminalTypeID from TerminalTypes with (NoLock) where AnyTerminal=1);"
        MyCommon.LRT_Execute()
        a = i.Split(",")
        For j = 0 To a.GetUpperBound(0)
            MyCommon.QueryStr = "select 1 from OfferTerminals with (NoLock) where OfferID =" & OfferID & " and TerminalTypeId =" & a(j) & ""
            rst = MyCommon.LRT_Select()
            If Not rst.Rows.Count > 0 Then
                MyCommon.QueryStr = "insert into OfferTerminals with (RowLock) (OfferID,TerminalTypeID,LastUpdate) values(" & OfferID & "," & a(j) & "," & "getdate())"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where IncentiveID=" & OfferID
                MyCommon.LRT_Execute()
                ResetOfferApprovalStatus(OfferID)
            End If
        Next

        ' clear out all other terminals if anyterminals terminal type is selected
        MyCommon.QueryStr = "select TT.TerminalTypeId from OfferTerminals OT with (NoLock) " & _
                            "inner join TerminalTypes TT with (NoLock) on TT.TerminalTypeId=OT.TerminalTypeID " & _
                            "where OT.Excluded=0 and OfferID=" & OfferID & " and TT.AnyTerminal=1;"
        rst = MyCommon.LRT_Select()
        If (rst.Rows.Count > 0) Then
            MyCommon.QueryStr = "delete from OfferTerminals with (RowLock) where OfferID=" & OfferID & " and Excluded=0 and TerminalTypeID not in " & _
                                "(select TerminalTypeID from TerminalTypes where AnyTerminal=1);"
            MyCommon.LRT_Execute()
        End If

        ' clear out all excluded terminals if a regular terminal type is selected
        MyCommon.QueryStr = "select TT.TerminalTypeId from OfferTerminals OT with (NoLock) " & _
                            "inner join TerminalTypes TT with (NoLock) on TT.TerminalTypeId=OT.TerminalTypeID " & _
                            "where OT.Excluded=0 and OfferID=" & OfferID & " and TT.AnyTerminal=0;"
        rst = MyCommon.LRT_Select()
        If (rst.Rows.Count > 0) Then
            MyCommon.QueryStr = "delete from OfferTerminals with (RowLock) where OfferID=" & OfferID & " and Excluded=1;"
            MyCommon.LRT_Execute()
        End If
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-addterminal", LanguageID))
    ElseIf (Request.QueryString("terminals-rem1") <> "" And Request.QueryString("terminals-select") <> "") Then
        Dim i As String = Request.QueryString("terminals-select")
        Dim a() As String
        Dim j As Integer

        ' first delete all excluded terminals
        MyCommon.QueryStr = "delete from OfferTerminals with (RowLock) where OfferID=" & OfferID & " and Excluded=1;"
        MyCommon.LRT_Execute()

        a = i.Split(",")
        For j = 0 To a.GetUpperBound(0)
            MyCommon.QueryStr = "delete from OfferTerminals with (RowLock) where OfferID=" & OfferID & " and TerminalTypeID=" & a(j)
            MyCommon.LRT_Execute()
            MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where IncentiveID=" & OfferID
            MyCommon.LRT_Execute()
            ResetOfferApprovalStatus(OfferID)
        Next
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-removeterminal", LanguageID))
    ElseIf (Request.QueryString("terminals-add2") <> "" And Request.QueryString("terminals-available") <> "") Then
        Dim i As String = Request.QueryString("terminals-available")
        Dim a() As String
        Dim j As Integer

        ' first check if any terminals is selected before allowing an excluded terminal
        MyCommon.QueryStr = "select OT.* from OfferTerminals OT with (NoLock) " & _
                            "inner join TerminalTypes TT with (NoLock) on OT.TerminalTypeID = TT.TerminalTypeID " & _
                            "where OfferID=" & OfferID & " and Excluded=0 and AnyTerminal=1;"
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
            a = i.Split(",")
            For j = 0 To a.GetUpperBound(0)
                MyCommon.QueryStr = "select 1 from OfferTerminals with (NoLock) where Excluded=1 and OfferID =" & OfferID & " and TerminalTypeId =" & a(j) & ""
                rst = MyCommon.LRT_Select()
                If Not rst.Rows.Count > 0 Then
                    MyCommon.QueryStr = "insert into OfferTerminals with (RowLock) (OfferID, TerminalTypeID, LastUpdate, Excluded) values(" & OfferID & ", " & a(j) & ", " & "getdate(), 1)"
                    MyCommon.LRT_Execute()
                    MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where IncentiveID=" & OfferID
                    MyCommon.LRT_Execute()
                    ResetOfferApprovalStatus(OfferID)
                End If
            Next
            ' delete just in case the any terminal was selected
            MyCommon.QueryStr = "delete from OfferTerminals with (RowLock) where OfferID=" & OfferID & " and Excluded=1 and TerminalTypeID in " & _
                                "(select TerminalTypeID from TerminalTypes with (NoLock) where AnyTerminal=1);"
            MyCommon.LRT_Execute()
            MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-excludeterminal", LanguageID))
        Else
            infoMessage = Copient.PhraseLib.Lookup("offer-loc.excludewarning", LanguageID)
        End If
    ElseIf (Request.QueryString("terminals-rem2") <> "" And Request.QueryString("terminals-exclude") <> "") Then
        Dim i As String = Request.QueryString("terminals-exclude")
        Dim a() As String
        Dim j As Integer
        a = i.Split(",")
        For j = 0 To a.GetUpperBound(0)
            MyCommon.QueryStr = "delete from OfferTerminals with (RowLock) where PKID=" & a(j)
            MyCommon.LRT_Execute()
            MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where IncentiveID=" & OfferID
            MyCommon.LRT_Execute()
            ResetOfferApprovalStatus(OfferID)
        Next
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-unexcludeterminal", LanguageID))
    End If

    Send_HeadBegin("term.offer", "term.locations", OfferID)
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts()
    Send_HeadEnd()
    If (IsTemplate) Then
        Send_BodyBegin(IIf(Popup, 13, 11))
    Else
        Send_BodyBegin(IIf(Popup, 3, 1))
    End If

    If (Not Popup) Then
        Send_Bar(Handheld)
        Send_Help(CopientFileName)
        Send_Logos()
        Send_Tabs(Logix, 2)

        If (IsTemplate) Then
            'send template sub tabs for the UE Engine
            Send_Subtabs(Logix, 25, 8, , OfferID)
        Else
            'send offer subtabs for the UE Engine
            Send_Subtabs(Logix, 24, 8, , OfferID)
        End If
    End If

    If (Logix.UserRoles.AccessOffers = False AndAlso Not IsTemplate) Then
        Send_Denied(1, "perm.offers-access")
        GoTo done
    End If
    If (Logix.UserRoles.AccessTemplates = False AndAlso IsTemplate) Then
        Send_Denied(1, "perm.offers-access-templates")
        GoTo done
    End If

    If (BannersEnabled AndAlso Not Logix.IsAccessibleOffer(AdminUserID, OfferID)) Then
        Send("<script type=""text/javascript"" language=""javascript"">")
        Send("  function updateCookie() { return true; } ")
        Send("</script>")
        Send_Denied(1, "banners.access-denied-offer")
        Send_BodyEnd()
        GoTo done
    End If

    If (infoMessage = "") Then
        CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
    Else
        CloseAfterSave = False
    End If
    If (Request.QueryString("SaveThenClose") = "true") Then
        CloseAfterSave = True
    End If
%>
<script type="text/javascript">
  var isIE =(navigator.appName.indexOf("Microsoft")!=-1) ? true : false;
  var isOpera = (navigator.appName.indexOf("Opera")!=-1) ? true : false;
  var fullAvailList = null;
  var hasUnsavedChanges = <%Sendb(IIf(Request.QueryString("condChanged") = "true", "true","false")) %>;
  

  // This is the javascript array holding the function list
  // The PrintJavascriptArray ASP function can be used to print this array.
  // Locations array
  <% 
    If (BannersEnabled) Then
      If (EngineID = 0) Then
        MyCommon.QueryStr = "select BE.BannerID from AdminUserBanners AUB with (NoLock) " & _
                            "inner join Banners BAN with (NoLock) on BAN.BannerID = AUB.BannerID " & _
                            "inner join BannerEngines BE with (NoLock) on BE.BannerID = BAN.BannerID " & _
                            "inner join BannerOffers BO with (NoLock) on BO.BannerID = BE.BannerID " & _
                            "where AdminUserID = " & AdminUserID & " and BE.EngineID = " & EngineID & _
                            " and BAN.AllBanners=1 and BO.OfferId=" & OfferID & ";" 
      Else
        MyCommon.QueryStr = "select BE.BannerID from AdminUserBanners AUB with (NoLock) " & _
                            "inner join Banners BAN with (NoLock) on BAN.BannerID = AUB.BannerID " & _
                            "inner join BannerEngines BE with (NoLock) on BE.BannerID = BAN.BannerID " & _
                            "where AdminUserID = " & AdminUserID & " and BE.EngineID = " & EngineID & " and BAN.AllBanners=1;" 
      End If
      rst = MyCommon.LRT_Select
      AllBannersPermission = (rst.Rows.Count > 0)
      
      ' build up a string with the Banner IDs, if all banners then get all the banners assigned to that engine
      MyCommon.QueryStr  = "select Distinct AUB.BannerID, BAN.AllBanners from Banners BAN with (NoLock)	" & _
                           " inner join AdminUserBanners AUB with (NoLock) on AUB.BannerID = BAN.BannerID " & _
                           " inner join BannerOffers BO with (NoLock) on BO.BannerID = BAN.BannerID " & _
                           "where AUB.AdminUserID =" & AdminUserID & " and BO.OfferID=" & OfferID & " and BAN.Deleted=0"
      rst = MyCommon.LRT_Select
      If (rst.rows.count>0 )
        For Each row In rst.Rows
          BannerList.Add(MyCommon.NZ(row.Item("BannerID"), -1))
          If (MyCommon.NZ(row.Item("AllBanners"), False)) then
            ' get all the banners assigned to the engine
            MyCommon.QueryStr = "select BannerID from BannerEngines BE with (NoLock) where EngineID = " & EngineID & " " & _
                                " and BannerID <> " & MyCommon.NZ(row.Item("BannerID"), -1) & ";"
            rst3 = MyCommon.LRT_Select
            For Each row2 In rst3.Rows
              BannerList.Add(MyCommon.NZ(row2.Item("BannerID"), -1))
            Next
          End If
        Next
      Else
        BannerList.Add("-1")
      End If
      
      For LoopCtr = 0 to BannerList.Count-1
        If (LoopCtr >0) Then BannerListStr &= ","
        BannerListStr &= BannerList(LoopCtr)
      Next LoopCtr
      
      If EngineID = 0 Then
        MyCommon.QueryStr = "select LocationGroupID, LG.Name, PhraseID, AllLocations, BAN.BannerID, BAN.Name as BannerName from LocationGroups as LG with (NoLock) " & _
                            "left join Banners BAN with (NoLock) on BAN.BannerID = LG.BannerID and BAN.Deleted=0 " & _
                            "where LG.Deleted=0 and LocationGroupID not in (select LocationGroupID from OfferLocations with (NoLock) where Deleted=0 and OfferID=" & OfferID & ") " & _
                            "and (EngineID=" & EngineID & " and AllLocations=0) and (" & IIf(AllBannersPermission, "LG.BannerID is NULL or", "") & " LG.BannerID in (" & BannerListStr & ")) " & _
                            "union " & _
                            "select LocationGroupID, Name, PhraseID, AllLocations, 0 as BannerID, '" & Copient.PhraseLib.Lookup("term.all", LanguageID) & "' as BannerName from LocationGroups as LG with (NoLock) where Deleted=0 and LocationGroupID not in " & _
                            "(select LocationGroupID from OfferLocations with (NoLock) where Deleted=0 and OfferID=" & OfferID & ")" & " " & _
                            "and Deleted=0 and AllLocations=1 " & _
                            "order by AllLocations desc, BannerName, Name;"
      Else
        MyCommon.QueryStr = "select LocationGroupID, LG.Name,PhraseID, BAN.BannerID, BAN.Name as BannerName from LocationGroups as LG with (NoLock) " & _
                            "left join Banners BAN with (NoLock) on BAN.BannerID = LG.BannerID and BAN.Deleted=0 " & _
                            "where LG.Deleted=0 and LocationGroupID not in (select LocationGroupID from OfferLocations with (NoLock) where Deleted=0 and OfferID=" & OfferID & ")" & _
                            "   and (EngineID=" & EngineID & " and AllLocations=0) and (" & IIf(AllBannersPermission, "LG.BannerID is NULL or", "") & " LG.BannerID in (" & BannerListStr & ")) " & _
                            "order by AllLocations desc, BannerName, Name;"
      End If
    Else
    
      MyCommon.QueryStr = "select LocationGroupID,Name,PhraseID from LocationGroups as LG with (NoLock) where Deleted=0 and LocationGroupID not in " & _
                          "(select LocationGroupID from OfferLocations with (NoLock) where Deleted=0 and OfferID=" & OfferID & ")" & " " & _
                          "and Deleted=0 and (EngineID=" & EngineID & " or AllLocations=1) order by AllLocations Desc, Name"
    End If
    rst = MyCommon.LRT_Select
    
    If (rst.rows.count>0 )
      Sendb("var functionlist = Array(")
      For Each row In rst.Rows
        If IsDBNull(row.Item("PhraseID")) Then
          Sendb("""" & MyCommon.NZ(row.item("Name"), "").ToString().Replace("""", "\""") & """,")
        Else
          If (row.Item("PhraseID") = 0) Then
            Sendb("""" & MyCommon.NZ(row.item("Name"), "").ToString().Replace("""", "\""") & """,")
          Else
            Sendb("""" & Copient.PhraseLib.Lookup(row.item("PhraseID"), LanguageID) & """,")
          End If
        End If
      Next
      Send(""""");")
      Sendb("  var vallist = Array(")
      For Each row In rst.Rows
        Sendb("""" & row.item("LocationGroupID") & """,")
      Next
      Send(""""");")
      If (BannersEnabled) Then
        Sendb("  var bannerlist = Array(")
        For Each row In rst.Rows
          Sendb("""" & MyCommon.NZ(row.item("BannerName"), "") & """,")
        Next
        Send(""""");")
      End If
    Else
      Sendb("var functionlist = Array(")
      Send("""" & "" & """);")
      Sendb("  var vallist = Array(")
      Send("""" & "" & """);")
      If (BannersEnabled) Then
        Sendb("  var bannerlist = Array(")
        Send("""" & "" & """);")
      End If
    End If
  %>
  
  // Terminals Array 
  <%
    If (BannersEnabled) Then
      If EngineID = 0 Then
        MyCommon.QueryStr = "select TT.TerminalTypeID, TT.Name, TT.AnyTerminal, TT.FuelProcessing, BAN.Name as BannerName from TerminalTypes TT with (NoLock) " & _ 
                            "left join Banners BAN with (NoLock) on BAN.BannerID = TT.BannerID and BAN.Deleted=0 " & _
                            "where TT.Deleted=0 and TT.TerminalTypeID not in " & _
                            "  (select TerminalTypeID from OfferTerminals with (NoLock) where OfferID=" & OfferID & ") and EngineID=" & EngineID & " " & _
                            "and (TT.BannerID in (" & BannerListStr & ")) " & _
                            "union " & _
                            "select TT.TerminalTypeID, TT.Name, TT.AnyTerminal, TT.FuelProcessing, '" & Copient.PhraseLib.Lookup("term.allbanners", LanguageID) & "' as BannerName from TerminalTypes TT with (NoLock) " & _ 
                            "where TT.Deleted=0 and TT.TerminalTypeID not in " & _
                            "  (select TerminalTypeID from OfferTerminals with (NoLock) where OfferID=" & OfferID & ") and EngineID=" & EngineID & " " & _
                            "and ((TT.BannerID=0 and AnyTerminal=0) or AnyTerminal=1)" & _
                            "order by BannerName, AnyTerminal desc, TT.Name;"
      Else
        MyCommon.QueryStr = "select TT.TerminalTypeID, TT.Name, TT.AnyTerminal, TT.FuelProcessing, BAN.Name as BannerName from TerminalTypes TT with (NoLock) " & _ 
                            "left join Banners BAN with (NoLock) on BAN.BannerID = TT.BannerID and BAN.Deleted=0 " & _
                            "where TT.Deleted=0 and TT.TerminalTypeID not in " & _
                            "  (select TerminalTypeID from OfferTerminals with (NoLock) where OfferID=" & OfferID & ") and EngineID=" & EngineID & " " & _
                            "and (" & IIf(AllBannersPermission, "TT.BannerID=0 or", "") & " (TT.BannerID in ( " & BannerListStr & ") or AnyTerminal=1) ) " & _
                            "order by BannerName, AnyTerminal desc, TT.Name;"
      End If
    Else
      MyCommon.QueryStr = "select TerminalTypeID, FuelProcessing, Name from TerminalTypes with (NoLock) " & _
                          "where Deleted=0 and TerminalTypeID not in " & _
                          "(select TerminalTypeID from OfferTerminals with (NoLock) where OfferID=" & OfferID & ") " & _
                          "and EngineID=" & EngineID & " order by AnyTerminal desc, Name"
    End If
    
    rst = MyCommon.LRT_Select
    If (rst.rows.count>0)
      Sendb("var functionlist2 = Array(")
      For Each row In rst.Rows
        Sendb("""" & MyCommon.NZ(row.item("Name"), "").ToString().Replace("""", "\""") & """,")
      Next
      Send(""""");")
      If (BannersEnabled) Then
        Sendb("  var bannerlist2 = Array(")
        For Each row In rst.Rows
          Sendb("""" & MyCommon.NZ(row.item("BannerName"), "").ToString().Replace("""", "\""") & """,")
        Next
        Send(""""");")
      End If
      Sendb("var vallist2 = Array(")
      For Each row In rst.Rows
        Sendb("""" & row.item("TerminalTypeID") & """,")
      Next
      Send(""""");")
    Else
      Sendb("var functionlist2 = Array(")
      Send("""" & "" & """);")
      Sendb("var vallist2 = Array(")
      Send("""" & "" & """);")
      If (BannersEnabled) Then
        Sendb("  var bannerlist2 = Array(")
        Send("""" & "" & """);")
      End If
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
    var newOpt, optGp;
    
    document.getElementById("sgroups-available").size = "10";
    
    // Set references to the form elements
    selectObj = document.getElementById("sgroups-available");
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
    
    // Clear the options list (for IE, cloning the select box without its option and then replacing the 
    // existing one, is significantly faster than removing each option.
    if (textObj.value == '' && fullAvailList != null)  {
      document.getElementById('sgrouplist').replaceChild(fullAvailList, selectObj);
    } else {
      //var newSelectBox = selectObj.cloneNode(false);
      if (isIE) {
        <%
          Send("var newSelectBox = document.createElement('select');")
          Send("newSelectBox.setAttribute('id', 'sgroups-available');")
          Send("newSelectBox.setAttribute('name', 'sgroups-available');")
          Send("newSelectBox.setAttribute('class', 'longest');")
          Send("newSelectBox.setAttribute('size', '10');")
          Send("newSelectBox.setAttribute('multiple', 'multiple');")
        %>
      } else {
        var newSelectBox = document.createElement('select');
        newSelectBox.id = 'sgroups-available';
        newSelectBox.name = 'sgroups-available';
        newSelectBox.className = 'longest';
        newSelectBox.size = '10';
        newSelectBox.multiple = true;
        <%

        %>
      }
      
      document.getElementById('sgrouplist').replaceChild(newSelectBox, selectObj);
      selectObj = document.getElementById("sgroups-available");
     
      // Loop through the array and re-add matching options
      numShown = 0;
      for(i = 0; i < functionListLength; i++) {
        if(functionlist[i].search(re) != -1) {
          if (vallist[i] != "") {
            var newOpt = document.createElement('OPTION');
            newOpt.value = vallist[i];
            if (isIE) { newOpt.innerText = functionlist[i]}; 
            newOpt.text =  functionlist[i]; 
            
            <% If (BannersEnabled) Then %>
              if (!isOpera) {
                optGp = GetOptionGroup(bannerlist[i], selectObj);
                if (optGp != null) {
                  optGp.appendChild(newOpt);
                  selectObj.appendChild(optGp);
                } else {
                  selectObj[numShown] = newOpt;
                }                
              } else {
                selectObj[numShown] = newOpt;
              }
            <% Else %>
              selectObj[numShown] = new Option(newOpt.text, newOpt.value);
            <% End If %>
            if (vallist[i] == 1) {
              selectObj[numShown].style.fontWeight = 'bold';
              selectObj[numShown].style.color = 'brown';
            }
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
        try {
          selectObj.options[0].selected = true;
          enableEditButton(1);
        } catch (ex) {
          // ignore if unable to select (workaround for problem in IE 6)
        }
      }
    }
  }
  
  function GetOptionGroup(bannerName, elemSlct) {
    var elemGroup = null;
    
    if (bannerName == "") {
      elemGroup = document.getElementById("gpAllBanners");
      if (elemGroup == null) {
        elemGroup = document.createElement("OPTGROUP");
        elemGroup.label = "<% Sendb(Copient.PhraseLib.Lookup("term.unassigned", LanguageID)) %>";
        elemGroup.id = "gpAllBanners";
        elemSlct.appendChild(elemGroup);
      }
    } else {
      elemGroup = document.getElementById("gp" + bannerName);
      if (elemGroup == null) {
        elemGroup = document.createElement("OPTGROUP");
        elemGroup.label = bannerName;
        elemGroup.id = "gp" + bannerName;
        elemSlct.appendChild(elemGroup);
      }
    }
    return elemGroup;
  }
  
  // This is the function that refreshes the list after a keypress.
  // The maximum number to show can be limited to improve performance with
  // huge lists (1000s of entries).
  // The function clears the list, and then does a linear search through the
  // globally defined array and adds the matches back to the list.
  function handleKeyUp2(maxNumToShow) {
    var selectObj, textObj, functionList2Length;
    var i,  numShown;
    var searchPattern;
    var newOpt, optGp;
    var isAll = false;
    
    document.getElementById("terminals-available").size = "10";
    
    // Set references to the form elements
    selectObj = document.getElementById("terminals-available");
    textObj = document.forms[0].functioninput2;
    
    // Remember the function list length for loop speedup
    functionList2Length = functionlist2.length;
    
    // Set the search pattern depending
    if(document.forms[0].functionradio3[0].checked == true) {
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
    for(i = 0; i < functionList2Length; i++) {
      if(functionlist2[i].search(re) != -1) {
        if (vallist2[i] != "") {
          isAll = (functionlist2[i] == "All CPE Terminals" || functionlist2[i] == "All CM Terminals" || functionlist2[i] == "All DP Terminals")
          var newOpt = document.createElement('OPTION');
          newOpt.value = vallist2[i];
          if (isIE) { newOpt.innerText = functionlist2[i]}; 
          newOpt.text =  functionlist2[i]; 
          <% If (BannersEnabled) Then %>
            if (!isOpera) {
              optGp = GetTerminalOptionGroup(bannerlist2[i], selectObj, isAll);
              if (optGp != null) {
                optGp.appendChild(newOpt);
                selectObj.appendChild(optGp);
              } else {
                selectObj[numShown] = newOpt
              }                
            } else {
              selectObj[numShown] = newOpt
            }
          <% Else %>
            selectObj[numShown] = new Option(newOpt.text, newOpt.value);
          <% End If %>
          
          if (isAll) {
            selectObj[numShown].style.fontWeight = 'bold';
            selectObj[numShown].style.color = 'brown';
          }
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
  
  function GetTerminalOptionGroup(bannerName, elemSlct, isAll) {
    var elemGroup = null;
    
    if (bannerName == "") {
      elemGroup = document.getElementById("gpTermAllBanners");
      if (elemGroup == null) {
        elemGroup = document.createElement("OPTGROUP");
        if (isAll) {
          elemGroup.label = "<% Sendb(Copient.PhraseLib.Lookup("term.allbanners", LanguageID)) %>";
        } else {
          elemGroup.label = "<% Sendb(Copient.PhraseLib.Lookup("term.unassigned", LanguageID)) %>";
        }
        elemGroup.id = "gpTermAllBanners";
        elemSlct.appendChild(elemGroup);
      }
    } else {
      elemGroup = document.getElementById("gpTerm" + bannerName);
      if (elemGroup == null) {
        elemGroup = document.createElement("OPTGROUP");
        elemGroup.label = bannerName;
        elemGroup.id = "gpTerm" + bannerName;
        elemSlct.appendChild(elemGroup);
      }
    }
    return elemGroup;
  }
  
  function handleKeyDown(e, slctName) {
    var key = e.which ? e.which : e.keyCode;
    
    if (key == 40) {
      var elemSlct = document.getElementById(slctName);
      if (elemSlct != null) { elemSlct.focus(); }
    }
  }
  
//  function handleSlctKeyDown(e, slctName, btnName) {
//    var key = e.which ? e.which : e.keyCode;
//    
//    if (key == 13) {
//      var elemSlct = document.getElementById(slctName);
//      if (elemSlct != null && elemSlct.disabled == false) { 
//        var fireOnThis = document.getElementById(btnName);
//        if( document.createEvent ) {
//          var evObj = document.createEvent('MouseEvents');
//          evObj.initEvent( 'click', true, false );
//          fireOnThis.dispatchEvent(evObj);
//        } else if( document.createEventObject ) {
//          fireOnThis.fireEvent('onclick');
//        }
//      }
//      e.returnValue=false;
//      return false;
//    }
//  }
  
  function ConfirmGroupExclude() { 
    var elem = document.getElementById('sgroups-available');
    var selectedCt = 0;
    var i = 0;
    var excludedGroup = "";
    
    if (elem != null) {
      for (i=0; i<elem.length; i++) {
        if(elem.options[i].selected) {
          selectedCt++;
          if (excludedGroup=="" && elem.options[i].value != "1") {
      	    excludedGroup = elem.options[i].text;
          }
        }
      }
      if (selectedCt > 1) {
        var response = confirm("<% Sendb(Copient.PhraseLib.Lookup("message.alertexcludegroup", LanguageID)) %>" + " '" + excludedGroup + "'?");
        if (response) {
          //' do something
        } else {
          return false;
        }
      }
    }
  }
    
<%
  If EngineID <> 7 Then
    goto skipPDEjavascript
  End If
%>
  function saveForm() {
    var funcSel = document.getElementById('sgroups-available');
    var elSel = document.getElementById('sgroups-select');
    var exSel = document.getElementById('sgroups-exclude');
    var i,j;
    var selectList = "";
    var excludedList = "";
    var htmlContents = "";
    
    // assemble the list of values from the selected box
    for (i = elSel.length - 1; i>=0; i--) {
      if(elSel.options[i].value != ""){
        if(selectList != "") { selectList = selectList + ","; }
        selectList = selectList + elSel.options[i].value;
      }
    }
    for (i = exSel.length - 1; i>=0; i--) {
      if(exSel.options[i].value != ""){
        if(excludedList != "") { excludedList = excludedList + ","; }
        excludedList = excludedList + exSel.options[i].value;
      }
    }
    
    // time to build up the hidden variables to pass for saving
    htmlContents = "<input type=\"hidden\" name=\"selGroups\" value=" + selectList + "> ";
    htmlContents = htmlContents + "<input type=\"hidden\" name=\"exGroups\" value=" + excludedList + ">";
    document.getElementById("hiddenVals").innerHTML = htmlContents;
    
    document.getElementById("LocationGroupID").value = selectList;
    document.getElementById("ExcludedStores").value = excludedList; 
    
    // alert(htmlContents);
    return true;
  }
  
  function isExceptionGroup(groupID) {
    var bRetVal = false;
    
    for (var i=0; i < exceptlist.length && !bRetVal; i++) {
      bRetVal = (exceptlist[i] == groupID) 
    }
    
    return bRetVal;
  }
  
  function selectCreatedGroup(groupID) {
    var funcSel = document.getElementById('sgroups-available');
    var i,j;
    
    for(j=funcSel.length-1;j>=0; j--) {
      if(funcSel.options[j].value == groupID){
        funcSel.options[j].selected = true;
      }
    }
  }
  
  function createGroup(groupID) {
    document.location = "lgroup-edit.aspx?OfferID=<%Sendb(OfferID) %>&EngineID=7" + getGroupTokens();
  }
  
  function editGroup(mode) {
    var elem = null;
    var index = -1;
    var groupID = 0;
    
    if (mode == 1) {
      elem = document.getElementById("sgroups-available")
    } else if (mode == 2) {
      elem = document.getElementById("sgroups-select")
    } else if (mode == 3) {
      elem = document.getElementById("sgroups-exclude")
    }
    
    if (elem != null) {
      index = elem.selectedIndex;
      if (index > -1) {
        groupID = elem.options[index].value;
        
        if (groupID > 1) {
          document.location = "lgroup-edit.aspx?LocationGroupID=" + groupID + "&OfferID=<%Sendb(OfferID) %>&EngineID=7" + getGroupTokens();
        }
      }
    }
  }
  
  function enableEditButton(mode) {
    var elem = null;
    var elemButton = null;
    var index = -1;
    var groupID = 0;
    
    if (mode == 1) {
      elem = document.getElementById("sgroups-available");
      elemButton = document.getElementById("editAvailable");
      elemButton.disabled = false;
    } else if (mode == 2) {
      elem = document.getElementById("sgroups-select");
      elemButton = document.getElementById("editSelected");
    } else if (mode == 3) {
      elem = document.getElementById("sgroups-exclude");
      elemButton = document.getElementById("editExcluded");
    }
    
    if (elem != null) {
      index = elem.selectedIndex;
      
      if (index > -1) {
        groupID = elem.options[index].value;
        if (groupID <= 1) {
          elemButton.disabled = true;
        } else {
          elemButton.disabled = false;
        }
      } else {
        elemButton.disabled = true;
      }
    }
  }
  
  function handleCloseClick() {
    var saveChanges = false;
    var elem = null;
    
    if (hasUnsavedChanges) { 
      saveChanges = confirm('<%Sendb(Copient.PhraseLib.Lookup("PDEoffer-loc.save-on-close", LanguageID)) %>')
    }
    
    if (saveChanges) {
      // save the changes
      saveForm();
      elem = document.getElementById('SaveThenClose');
      if (elem != null) {
        elem.value = "true";
      }
      document.mainform.submit();
    } else {
      // close the customer window.
      top.closeIframePopup(true);
    }
  }
  
  function getGroupTokens() {
    var funcSel = document.getElementById('sgroups-available');
    var elSel = document.getElementById('sgroups-select');
    var exSel = document.getElementById('sgroups-exclude');
    var i,j;
    var selectList = "-1";
    var excludedList = "-1";
    var groupTokens = "";
    
    // assemble the list of values from the selected box
    for (i = elSel.length - 1; i>=0; i--) {
      if(elSel.options[i].value != ""){
        if(selectList != "") { selectList = selectList + ","; }
        selectList = selectList + elSel.options[i].value;
      }
    }
    for (i = exSel.length - 1; i>=0; i--) {
      if(exSel.options[i].value != ""){
        if(excludedList != "") { excludedList = excludedList + ","; }
        excludedList = excludedList + exSel.options[i].value;
      }
    }
    
    if (selectList != '') {
      groupTokens = '&slct=' + selectList;
    }
    if (excludedList != '') {
      groupTokens += '&ex=' + excludedList
    }
    if (hasUnsavedChanges) {
      groupTokens += '&condChanged=true';
    }
    
    return groupTokens;
  }
<%
skipPDEjavascript:
%>
</script>
<form action="#" id="mainform" name="mainform">
<input type="hidden" id="OfferID" name="OfferID" value="<% Sendb(OfferID) %>" />
<input type="hidden" id="Name" name="Name" value="<% Sendb(Name) %>" />
<input type="hidden" id="Popup" name="Popup" value="<% Sendb(IIf(Popup, 1, 0)) %>" />
<input type="hidden" id="PromoID" name="PromoID" value="<% Sendb(PromoID) %>" />
<input type="hidden" id="EngineID" name="EngineID" value="<% Sendb(EngineID) %>" />
    <input type="hidden" id="CustomerGroupID" name="CustomerGroupID" value="<% Sendb(LocationGroupID)%>" />
    <input type="hidden" id="ExcludedStores" name="ExcludedStores" value="<% Sendb(ExcludedStores)%>" />
    <input type="hidden" id="SaveThenClose" name="SaveThenClose" value="false" />
    <input type="hidden" id="savedTime" name="savedTime" value="<%=DateTime.Now()%>" />
    <%
        Send("<div id=""intro"">")
        If (IsTemplate) Then
        Sendb("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID)
    Else
        Sendb("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID)
    End If
    Send(": " & MyCommon.TruncateString(Name, 50) & "</h1>")
    Send("<div id=""controls"">")
    'If (IsTemplate) Then
    'Send_Save()
    'End If
    If MyCommon.Fetch_SystemOption(75) Then
        If (OfferID > 0 AndAlso Logix.UserRoles.AccessNotes AndAlso Not Popup AndAlso (Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse  bOfferEditable)) Then
            Send_NotesButton(3, OfferID, AdminUserID)
        End If
    End If
    Send("</div>")
    Send("</div>")
%>
<a name="h00" id="h00"></a>
<div id="main">
    <%
        MyCommon.QueryStr = "select StatusFlag from CPE_Incentives where IncentiveID=" & OfferID & ";"
        rst = MyCommon.LRT_Select
        StatusText = Logix.GetOfferStatus(OfferID, LanguageID, StatusCode)
        If Not IsTemplate Then
            If (MyCommon.NZ(rst.Rows(0).Item("StatusFlag"), 0) <> 2) Then
                If (StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_ACTIVE) AndAlso (MyCommon.NZ(rst.Rows(0).Item("StatusFlag"), 0) > 0) Then
                    modMessage = Copient.PhraseLib.Lookup("alert.modpostdeploy", LanguageID)
                    Send("<div id=""modbar"">" & modMessage & "</div>")
                End If
            End If
        End If
        If (infoMessage <> "") Then
            Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
        End If
      
        ' Send the status bar, but not if it's a new offer or a template, or if there's already a modMessage being shown.
        If (Not IsTemplate AndAlso modMessage = "") Then
            MyCommon.QueryStr = "select IncentiveID from CPE_Incentives with (NoLock) where CreatedDate=LastUpdate and IncentiveID=" & OfferID & ";"
            rst3 = MyCommon.LRT_Select
            If (rst3.Rows.Count = 0) Then
                Send_Status(OfferID, 9)
            End If
        End If
    %>
    <div id="<%Sendb(IIf(EngineID = 7, "column", "column1")) %>">
        <div class="box" id="storegroups">
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.storegroups", LanguageID))%>
                </span>
            </h2>
            <% If (IsTemplate) Then%>
            <!--
        <span class="temp">
          <input type="checkbox" class="tempcheck" id="temp-stores" name="Disallow_Stores" />
          <label for="temp-stores"><% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID)) %></label>
        </span>
        -->
            <% End If%>
            <label for="sgroups-available">
                <b>
                    <% Sendb(Copient.PhraseLib.Lookup("term.available", LanguageID) & ":")%>
                </b>
            </label>
            <br clear="all" />
            <input type="radio" id="functionradio1" name="functionradio" <% if(MyCommon.Fetch_SystemOption(175)= "1" ) then sendb(" checked=""checked""") %> /><label
                for="functionradio1"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
            <input type="radio" id="functionradio2" name="functionradio" <% if(MyCommon.Fetch_SystemOption(175)= "2" ) then sendb(" checked=""checked""") %> /><label
                for="functionradio2"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
            <%
            %>
            <input class="longer" onkeydown="handleKeyDown(event, 'sgroups-available');" onkeyup="handleKeyUp(200);"
                id="functioninput" name="functioninput" maxlength="100" type="text" value="" /><br />
            <br class="half" />
            <%
                Send("<span id=""sgrouplist"">")
                Send("<select class=""longest"" multiple=""multiple"" id=""sgroups-available"" name=""sgroups-available"" size=""10""" & IIf(EngineID = 7, " onchange=""enableEditButton(1)"" style=""width:650px;""", "") & ">")
                Dim topString As String = ""
                If RECORD_LIMIT > 0 Then topString = "top " & RECORD_LIMIT
                    If BannersEnabled Then
                        MyCommon.QueryStr = "select " & topString & " LocationGroupID, LG.Name,PhraseID, BAN.BannerID, BAN.Name as BannerName from LocationGroups as LG with " & _
                                            "(NoLock) left join Banners BAN with (NoLock) on BAN.BannerID = LG.BannerID and BAN.Deleted=0 where " & _
                                             "LG.Deleted=0 and LocationGroupID not in (select LocationGroupID from OfferLocations with (NoLock) where Deleted=0 and OfferID=" & OfferID & ") " & _
                                                " and (EngineID=" & EngineID & " and AllLocations=0) and (" & IIf(AllBannersPermission, "LG.BannerID is NULL or", "") & " LG.BannerID in (" & BannerListStr & "))  order by AllLocations desc, BannerName, Name;"
                    Else
                MyCommon.QueryStr = "select  " & topString & "LocationGroupID,Name,PhraseID from LocationGroups as LG with (NoLock) where Deleted=0 and LocationGroupID not in " & _
                                    "(select LocationGroupID from OfferLocations with (NoLock) where Deleted=0 and OfferID=" & OfferID & ")" & " and Deleted=0 and (EngineID=" & EngineID & " or AllLocations=1) order by AllLocations desc, Name"
                    End If
                rst = MyCommon.LRT_Select
                If (rst.Rows.Count > 0) Then
                        If BannersEnabled Then
                            For Each item In BannerList
                                MyCommon.QueryStr = "select Name from Banners where Bannerid=" & item & " and deleted=0"
                                Dim rst1 = MyCommon.LRT_Select
                                Send("<optgroup label=" & rst1.Rows(0).Item("Name") & ">")
                    For Each row In rst.Rows
                                    If Not IsDBNull(row.Item("BannerID")) AndAlso row.Item("BannerID") = item Then
                                        If IsDBNull(row.Item("PhraseID")) Then
                                            Send("<option value=""" & row.Item("LocationGroupID") & """>" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
                                        Else
                                            If (row.Item("PhraseID") = 0) Then
                                                Send("<option value=""" & row.Item("LocationGroupID") & """>" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
                                            Else
                                                Send("<option value=""" & row.Item("LocationGroupID") & """ style=""font-weight:bold;color:brown;"">" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID) & "</option>")
                                            End If
                                        End If
                                    End If
                                Next
                                Send("</optgroup>")
                            Next
                        Else
                            For Each row In rst.Rows
                        If IsDBNull(row.Item("PhraseID")) Then
                            Send("<option value=""" & row.Item("LocationGroupID") & """>" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
                        Else
                            If (row.Item("PhraseID") = 0) Then
                                Send("<option value=""" & row.Item("LocationGroupID") & """>" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
                            Else
                                Send("<option value=""" & row.Item("LocationGroupID") & """ style=""font-weight:bold;color:brown;"">" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID) & "</option>")
                            End If
                        End If
                    Next
                        End If
                End If
                Send("</select>")
                Send("</span>")
                Send("<br />")
                Send("<br />")
          
                ' First off: queries to get both the selected and excluded store groups
                If (BannersEnabled) Then
                    MyCommon.QueryStr = "select O.LocationGroupID, LG.Name, LG.PhraseID, BAN.Name as BannerName from OfferLocations as O with (NoLock) " & _
                                        "left join LocationGroups as LG with (NoLock) on LG.LocationGroupID=O.LocationGroupID " & _
                                        "left join Banners as BAN with (NoLock) on LG.BannerID = BAN.BannerID and BAN.Deleted=0 " & _
                                        "where Excluded=0 and O.Deleted=0 and O.OfferID=" & OfferID & " order by BAN.Name, LG.Name"
                Else
                    MyCommon.QueryStr = "select O.LocationGroupID, LG.Name, LG.PhraseID from OfferLocations as O with (NoLock) " & _
                                        "left join LocationGroups as LG with (NoLock) on LG.LocationGroupID=O.LocationGroupID " & _
                                        "where Excluded=0 and O.Deleted=0 and O.OfferID=" & OfferID & " order by LG.Name"
                End If
                rstSelected = MyCommon.LRT_Select
                MyCommon.QueryStr = "select O.LocationGroupID,LG.Name,LG.PhraseID from OfferLocations as O with (NoLock) left join LocationGroups as LG with (NoLock) on " & _
                                    "LG.LocationGroupID=O.LocationGroupID where Excluded=1 and O.deleted=0 and O.OfferID=" & OfferID & " order by LG.Name"
                rstExcluded = MyCommon.LRT_Select
          
                ' SELECTED STORE GROUPS
                Send("<label for=""sgroups-select""><b>" & Copient.PhraseLib.Lookup("term.selected", LanguageID) & ":</b></label>")
                Send("<br />")
          
                ' Buttons
                'If (EngineID = 7 And Logix.UserRoles.EditCustomerGroups) Then
                '  Send("<div style=""float:right;"">")
                '  Send("  <input type=""button"" class=""short"" id=""editSelected"" name=""editSelected"" value=""" & Copient.PhraseLib.Lookup("term.edit", LanguageID) & """ onclick=""editGroup(2)"" disabled=""disabled"" />")
                '  Send("</div>")
                'End If
                Sendb("<input type=""submit"" class=""regular select"" id=""stores-add1"" name=""stores-add1"" title=""" & Copient.PhraseLib.Lookup("term.select", LanguageID) & """ value=""&#9660; " & Copient.PhraseLib.Lookup("term.select", LanguageID) & """")
                If Not IsTemplate Then
                    If (Not Logix.UserRoles.CRUDLocationsToOffers) OrElse (bEnableRestrictedAccessToUEOfferBuilder AndAlso isTranslatedOffer) OrElse (bEnableAdditionalLockoutRestrictionsOnOffers AndAlso Not bOfferEditable) Then
                        Sendb(" disabled=""disabled""")
                    End If
                Else
                    If (Not Logix.UserRoles.CRUDLocationsToTemplates) OrElse (bEnableRestrictedAccessToUEOfferBuilder AndAlso isTranslatedOffer) OrElse (bEnableAdditionalLockoutRestrictionsOnOffers AndAlso Not bOfferEditable) Then
                        Sendb(" disabled=""disabled""")
                    End If
                End If
                If (Logix.UserRoles.EditOffer = False Or IsOfferWaitingForApproval(OfferID) Or Not (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID)) OrElse (bEnableRestrictedAccessToUEOfferBuilder AndAlso isTranslatedOffer) OrElse (bEnableAdditionalLockoutRestrictionsOnOffers AndAlso Not bOfferEditable)) Then
                    Sendb("disabled=""disabled""")
                End If
                Send(" />")
                Sendb("<input type=""submit"" class=""regular deselect"" id=""stores-rem1"" name=""stores-rem1"" title=""" & Copient.PhraseLib.Lookup("term.unselect", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.deselect", LanguageID) & " &#9650;""")
                If Not IsTemplate Then
                    If (Not Logix.UserRoles.CRUDLocationsToOffers) OrElse (rstSelected.Rows.Count = 0) OrElse (bEnableRestrictedAccessToUEOfferBuilder AndAlso isTranslatedOffer) OrElse (bEnableAdditionalLockoutRestrictionsOnOffers AndAlso Not bOfferEditable) Then
                        Sendb(" disabled=""disabled""")
                    End If
                Else
                    If (Not Logix.UserRoles.CRUDLocationsToTemplates) OrElse (rstSelected.Rows.Count = 0) OrElse (bEnableRestrictedAccessToUEOfferBuilder AndAlso isTranslatedOffer) OrElse (bEnableAdditionalLockoutRestrictionsOnOffers AndAlso Not bOfferEditable)  Then
                        Sendb(" disabled=""disabled""")
                    End If
                End If
                If (Logix.UserRoles.EditOffer = False Or IsOfferWaitingForApproval(OfferID) Or Not (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID)) OrElse (bEnableRestrictedAccessToUEOfferBuilder AndAlso isTranslatedOffer) OrElse (bEnableAdditionalLockoutRestrictionsOnOffers AndAlso Not bOfferEditable)) Then
                    Sendb("disabled=""disabled""")
                End If
                Send(" />")
                Send("<br />")
          
                ' List
                Send("<select class=""longest"" multiple=""multiple"" id=""sgroups-select"" name=""sgroups-select"" size=""" & SelectSize & """" & IIf(EngineID = 7, " onchange=""enableEditButton(2)"" style=""height:38px; margin-top:2px; width:650px;""", "") & ">")
                If (rstSelected.Rows.Count > 0) Then
                    For Each row In rstSelected.Rows
                        If (BannersEnabled) Then
                            If row.Item("LocationGroupID") = 1 Then
                                NameValue = Copient.PhraseLib.Lookup("term.all", LanguageID)
                            Else
                                NameValue = MyCommon.NZ(row.Item("BannerName"), Copient.PhraseLib.Lookup("term.unassigned", LanguageID))
                            End If
                            If (NameValue <> PrevNameValue) Then
                                If (PrevNameValue <> "") Then Send("</optgroup>")
                                Send("   <optgroup label=""" & NameValue & """>  ")
                            End If
                            PrevNameValue = NameValue
                        End If
                        If IsDBNull(row.Item("PhraseID")) Then
                            Send("<option value=""" & row.Item("LocationGroupID") & """>" & row.Item("Name") & "</option>")
                        Else
                            If (row.Item("PhraseID") = 0) Then
                                Send("<option value=""" & row.Item("LocationGroupID") & """>" & row.Item("Name") & "</option>")
                            Else
                                Send("<option value=""" & row.Item("LocationGroupID") & """ style=""font-weight:bold;color:brown;"">" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID) & "</option>")
                            End If
                        End If
                    Next
                    If (BannersEnabled) Then Send("</optgroup>")
                End If
                Send("</select>")
                Send("<br />")
                Send("<br class=""half"" />")
          
                ' EXCLUDED STORE GROUPS
                Send("<label for=""sgroups-exclude""><b>" & Copient.PhraseLib.Lookup("term.excluded", LanguageID) & ":</b></label>")
                Send("<br />")
          
                ' Buttons
                'If (EngineID = 7 And Logix.UserRoles.EditCustomerGroups) Then
                '  Send("<div style=""float:right;"">")
                '  Send("  <input type=""button"" class=""short"" id=""editExcluded"" name=""editExcluded"" value=""" & Copient.PhraseLib.Lookup("term.edit", LanguageID) & """ onclick=""editGroup(3)"" disabled=""disabled"" />")
                '  Send("</div>")
                'End If
                Sendb("<input type=""submit"" class=""regular select"" id=""stores-add2"" name=""stores-add2"" onclick=""return ConfirmGroupExclude();"" title=""" & Copient.PhraseLib.Lookup("term.exclude", LanguageID) & """ value=""&#9660; " & Copient.PhraseLib.Lookup("term.select", LanguageID) & """")
                If Not IsTemplate Then
                    If (Not Logix.UserRoles.CRUDLocationsToOffers) OrElse (rstSelected.Rows.Count = 0) OrElse (rstSelected.Rows(0).Item("LocationGroupID") <> 1) OrElse (bEnableRestrictedAccessToUEOfferBuilder AndAlso isTranslatedOffer) OrElse (bEnableAdditionalLockoutRestrictionsOnOffers AndAlso Not bOfferEditable)  Then
                        Sendb(" disabled=""disabled""")
                    End If
                Else
                    If (Not Logix.UserRoles.CRUDLocationsToTemplates) OrElse (rstSelected.Rows.Count = 0) OrElse (rstSelected.Rows(0).Item("LocationGroupID") <> 1 OrElse (bEnableRestrictedAccessToUEOfferBuilder AndAlso isTranslatedOffer) OrElse (bEnableAdditionalLockoutRestrictionsOnOffers AndAlso Not bOfferEditable)) Then
                        Sendb(" disabled=""disabled""")
                    End If
                End If
                If (Logix.UserRoles.EditOffer = False Or IsOfferWaitingForApproval(OfferID) Or Not (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID)) OrElse (bEnableRestrictedAccessToUEOfferBuilder AndAlso isTranslatedOffer) OrElse (bEnableAdditionalLockoutRestrictionsOnOffers AndAlso Not bOfferEditable)) Then
                    Sendb("disabled=""disabled""")
                End If
                Send(" />")
                Sendb("<input type=""submit"" class=""regular deselect"" id=""stores-rem2"" name=""stores-rem2"" title=""" & Copient.PhraseLib.Lookup("term.unexclude", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.deselect", LanguageID) & " &#9650;""")
                If Not IsTemplate Then
                    If (Not Logix.UserRoles.CRUDLocationsToOffers) OrElse (rstExcluded.Rows.Count = 0 )  OrElse (bEnableRestrictedAccessToUEOfferBuilder AndAlso isTranslatedOffer) OrElse (bEnableAdditionalLockoutRestrictionsOnOffers AndAlso Not bOfferEditable) Then
                        Sendb(" disabled=""disabled""")
                    End If
                Else
                    If (Not Logix.UserRoles.CRUDLocationsToTemplates) OrElse (rstExcluded.Rows.Count = 0) OrElse (bEnableRestrictedAccessToUEOfferBuilder AndAlso isTranslatedOffer) OrElse (bEnableAdditionalLockoutRestrictionsOnOffers AndAlso Not bOfferEditable) Then
                        Sendb(" disabled=""disabled""")
                    End If
                End If
                If (Logix.UserRoles.EditOffer = False Or IsOfferWaitingForApproval(OfferID) Or Not (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID)) OrElse (bEnableRestrictedAccessToUEOfferBuilder AndAlso isTranslatedOffer) OrElse (bEnableAdditionalLockoutRestrictionsOnOffers AndAlso Not bOfferEditable)) Then
                    Sendb("disabled=""disabled""")
                End If
                Send(" />")
                Send("<br />")
          
                ' List
                Send("<select class=""longest"" id=""sgroups-exclude"" name=""sgroups-exclude"" size=""5""" & IIf(EngineID = 7, " onchange=""enableEditButton(3)"" style=""height:38px; margin-top:2px; width:650px;""", "") & ">")
                For Each row In rstExcluded.Rows
                    If IsDBNull(row.Item("PhraseID")) Then
                        Send("<option value=""" & row.Item("LocationGroupID") & """>" & row.Item("Name") & "</option>")
                    Else
                        If (row.Item("PhraseID") = 0) Then
                            Send("<option value=""" & row.Item("LocationGroupID") & """>" & row.Item("Name") & "</option>")
                        Else
                            Send("<option value=""" & row.Item("LocationGroupID") & """>" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID) & "</option>")
                        End If
                    End If
                Next
                Send("</select>")
            %>
            <hr class="hidden" />
        </div>
        <a name="h01"></a>
        <%
            'If (EngineID = 4) Then
            '  PaddingValue = 50
            '  BtnPaddingTop += PaddingValue + 30
            '  SelectSize = 10
            'End If
        %>
    </div>
    <div id="gutter" <%Sendb(IIf(EngineID = 7, " style=""display:none;""", "")) %>>
    </div>
    <div id="column2" <%Sendb(IIf(EngineID = 7, " style=""display:none;""", "")) %>>
        <div class="box" id="terminals">
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.terminals", LanguageID))%>
                </span>
            </h2>
            <% If (IsTemplate) Then%>
            <!--
        <span class="temp">
          <input type="checkbox" class="tempcheck" id="temp-terminals" name="Disallow_Terminals" />
          <label for="temp-terminals"><% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID)) %></label>
        </span>
        -->
            <% End If%>
            <label for="terminals-available">
                <b>
                    <% Sendb(Copient.PhraseLib.Lookup("term.available", LanguageID) & ":")%>
                </b>
            </label>
            <br clear="all" />
            <input type="radio" id="functionradio3" name="functionradio3" <% if(MyCommon.Fetch_SystemOption(175)= "1" ) then sendb(" checked=""checked""") %> /><label
                for="functionradio3"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
            <input type="radio" id="functionradio4" name="functionradio3" <% if(MyCommon.Fetch_SystemOption(175)= "2" ) then sendb(" checked=""checked""") %> /><label
                for="functionradio4"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
            <input class="longer" onkeydown="handleKeyDown(event, 'terminals-available');" onkeyup="handleKeyUp2(9999);"
                id="functioninput2" name="functioninput2" maxlength="100" type="text" value="" /><br />
            <br class="half" />
            <select class="longest" multiple="multiple" id="terminals-available" name="terminals-available"
                size="10">
                <%
                    Dim topStringTerm As String = ""
                    If RECORD_LIMIT > 0 Then topStringTerm = "top " & RECORD_LIMIT
                    If (BannersEnabled) Then
                        If EngineID = 0 Then
                            MyCommon.QueryStr = "select " & topStringTerm & "  TT.TerminalTypeID, TT.Name,TT.PhraseID TT.AnyTerminal, TT.FuelProcessing, BAN.Name as BannerName from TerminalTypes TT with (NoLock) " & _
                                            "left join Banners BAN with (NoLock) on BAN.BannerID = TT.BannerID and BAN.Deleted=0 " & _
                                            "where TT.Deleted=0 and TT.TerminalTypeID not in " & _
                                            "  (select TerminalTypeID from OfferTerminals with (NoLock) where OfferID=" & OfferID & ") and EngineID=" & EngineID & " " & _
                                            "and (TT.BannerID in (" & BannerListStr & ")) " & _
                                            "union " & _
                                            "select TT.TerminalTypeID, TT.Name, TT.PhraseID,TT.AnyTerminal, TT.FuelProcessing, '" & Copient.PhraseLib.Lookup("term.allbanners", LanguageID) & "' as BannerName from TerminalTypes TT with (NoLock) " & _
                                            "where TT.Deleted=0 and TT.TerminalTypeID not in " & _
                                            "  (select TerminalTypeID from OfferTerminals with (NoLock) where OfferID=" & OfferID & ") and EngineID=" & EngineID & " " & _
                                            "and ((TT.BannerID=0 and AnyTerminal=0) or AnyTerminal=1)" & _
                                            "order by BannerName, AnyTerminal desc, TT.Name;"
                        Else
                            MyCommon.QueryStr = "select " & topStringTerm & "  TT.TerminalTypeID, TT.Name,TT.PhraseID, TT.AnyTerminal, TT.FuelProcessing, BAN.Name as BannerName from TerminalTypes TT with (NoLock) " & _
                                                "left join Banners BAN with (NoLock) on BAN.BannerID = TT.BannerID and BAN.Deleted=0 " & _
                                                "where TT.Deleted=0 and TT.TerminalTypeID not in " & _
                                                "  (select TerminalTypeID from OfferTerminals with (NoLock) where OfferID=" & OfferID & ") and EngineID=" & EngineID & " " & _
                                                "and (" & IIf(AllBannersPermission, "TT.BannerID=0 or", "") & " (TT.BannerID in ( " & BannerListStr & ") or AnyTerminal=1) ) " & _
                                                "order by BannerName, AnyTerminal desc, TT.Name;"
                        End If
                    Else
                        MyCommon.QueryStr = "select " & topStringTerm & " TerminalTypeID, FuelProcessing, Name ,PhraseID from TerminalTypes with (NoLock) " & _
                                            "where Deleted=0 and TerminalTypeID not in " & _
                                            "(select TerminalTypeID from OfferTerminals with (NoLock) where OfferID=" & OfferID & ") " & _
                                            "and EngineID=" & EngineID & " order by AnyTerminal desc, Name"
                    End If
                    rst=MyCommon.LRT_Select()
                    If (rst.Rows.Count > 0) Then
                        For Each row In rst.Rows
                            If IsDBNull(row.Item("PhraseID")) Then
                                Send("<option value=""" & row.Item("TerminalTypeID") & """>" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
                            Else
                                If (row.Item("PhraseID") = 0) Then
                                    Send("<option value=""" & row.Item("TerminalTypeID") & """>" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
                                Else
                                    Send("<option value=""" & row.Item("TerminalTypeID") & """ style=""font-weight:bold;color:brown;"">" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID) & "</option>")
                                End If
                            End If
                        Next
                    End If
                    %>
            </select>
            <br />
            <br />
            <%
                ' First off: queries to get both the selected and excluded terminals
                If (BannersEnabled) Then
                    MyCommon.QueryStr = "select OT.TerminalTypeID as TID, T.Name, T.FuelProcessing, T.PhraseID, BAN.Name as BannerName, T.AnyTerminal,T.BannerID from OfferTerminals as OT with (NoLock) " & _
                                        "left join TerminalTypes as T with (NoLock) on OT.TerminalTypeID=T.TerminalTypeID " & _
                                        "left join Banners BAN with (NoLock) on BAN.BannerID = T.BannerID and BAN.Deleted=0 " & _
                                        "where OfferID=" & OfferID & " and Excluded=0" & _
                                        "order by AnyTerminal desc, BAN.Name, T.Name"
                Else
                    MyCommon.QueryStr = "select OT.TerminalTypeID as TID, T.Name, T.FuelProcessing, T.PhraseID, T.AnyTerminal from OfferTerminals as OT with (NoLock) " & _
                                        "left join TerminalTypes as T with (NoLock) on OT.TerminalTypeID=T.TerminalTypeID " & _
                                        "where OfferID=" & OfferID & " and Excluded=0" & _
                                        "order by AnyTerminal desc, T.Name"
                End If
                rstSelected = MyCommon.LRT_Select
                MyCommon.QueryStr = "select OT.PKID, T.Name, T.FuelProcessing from OfferTerminals as OT with (nolock) " & _
                                    "left join TerminalTypes as T with (NoLock) on T.TerminalTypeID=OT.TerminalTypeID " & _
                                    "where OT.OfferID=" & OfferID & " and OT.Excluded=1" & _
                                    "order by T.Name"
                rstExcluded = MyCommon.LRT_Select
          
                ' SELECTED TERMINALS
                Send("<label for=""terminals-select""><b>" & Copient.PhraseLib.Lookup("term.selected", LanguageID) & ":</b></label>")
                Send("<br />")
          
                ' Buttons
                Sendb("<input type=""submit"" class=""regular select"" id=""terminals-add1"" name=""terminals-add1"" title=""" & Copient.PhraseLib.Lookup("term.select", LanguageID) & """ value=""&#9660; " & Copient.PhraseLib.Lookup("term.select", LanguageID) & """")
                If Not IsTemplate Then
                    If (Not Logix.UserRoles.CRUDLocationsToOffers) OrElse (bEnableRestrictedAccessToUEOfferBuilder AndAlso isTranslatedOffer) OrElse (bEnableAdditionalLockoutRestrictionsOnOffers AndAlso Not bOfferEditable) Then
                        Sendb(" disabled=""disabled""")
                    End If
                Else
                    If (Not Logix.UserRoles.CRUDLocationsToTemplates) OrElse (bEnableRestrictedAccessToUEOfferBuilder AndAlso isTranslatedOffer) OrElse (bEnableAdditionalLockoutRestrictionsOnOffers AndAlso Not bOfferEditable) Then
                        Sendb(" disabled=""disabled""")
                    End If
                End If
                If (Logix.UserRoles.EditOffer = False Or IsOfferWaitingForApproval(OfferID) Or Not (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID)) OrElse (bEnableRestrictedAccessToUEOfferBuilder AndAlso isTranslatedOffer) OrElse (bEnableAdditionalLockoutRestrictionsOnOffers AndAlso Not bOfferEditable)) Then
                    Sendb("disabled=""disabled""")
                End If
                Send(" />")
                Sendb("<input type=""submit"" class=""regular deselect"" id=""terminals-rem1"" name=""terminals-rem1"" title=""" & Copient.PhraseLib.Lookup("term.unselect", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.deselect", LanguageID) & " &#9650;""")
                If Not IsTemplate Then
                    If (Not Logix.UserRoles.CRUDLocationsToOffers) OrElse (rstSelected.Rows.Count = 0) OrElse (bEnableRestrictedAccessToUEOfferBuilder AndAlso isTranslatedOffer) OrElse (bEnableAdditionalLockoutRestrictionsOnOffers AndAlso Not bOfferEditable) Then
                        Sendb(" disabled=""disabled""")
                    End If
                Else
                    If (Not Logix.UserRoles.CRUDLocationsToTemplates) OrElse (rstSelected.Rows.Count = 0) OrElse (bEnableRestrictedAccessToUEOfferBuilder AndAlso isTranslatedOffer) OrElse (bEnableAdditionalLockoutRestrictionsOnOffers AndAlso Not bOfferEditable) Then
                        Sendb(" disabled=""disabled""")
                    End If
                End If
                If (Logix.UserRoles.EditOffer = False Or IsOfferWaitingForApproval(OfferID) Or Not (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID)) OrElse (bEnableRestrictedAccessToUEOfferBuilder AndAlso isTranslatedOffer) OrElse (bEnableAdditionalLockoutRestrictionsOnOffers AndAlso Not bOfferEditable)) Then
                    Sendb("disabled=""disabled""")
                End If
                Send(" />")
                Send("<br />")
          
                ' List
                Send("<select class=""longest"" multiple=""multiple"" id=""terminals-select"" name=""terminals-select"" size=""" & SelectSize & """>")
                If rstSelected.Rows.Count = 0 Then
                    'Send("<option value="""">All Terminals</option>")
                Else
                    PrevNameValue = ""
                    For Each row In rstSelected.Rows
                        If (BannersEnabled) Then
                            If (MyCommon.NZ(row.Item("AnyTerminal"), False)) Then
                                NameValue = Copient.PhraseLib.Lookup("term.allbanners", LanguageID)
                            Else
                                If (EngineID = 0 And (MyCommon.NZ(row.Item("BannerId"), 0)) = 0) Then
                                    NameValue = Copient.PhraseLib.Lookup("term.allbanners", LanguageID)
                                Else
                                    NameValue = MyCommon.NZ(row.Item("BannerName"), Copient.PhraseLib.Lookup("term.unassigned", LanguageID))
                                End If
                            End If
                            If (NameValue <> PrevNameValue) Then
                                If (PrevNameValue <> "") Then Send("</optgroup>")
                                Send("   <optgroup label=""" & NameValue & """>  ")
                            End If
                            PrevNameValue = NameValue
                        End If
                        If (row.Item("Name").ToString.Trim = "All CPE Terminals" OrElse row.Item("Name").ToString.Trim = "All CM Terminals" OrElse row.Item("Name").ToString.Trim = "All DP Terminals" OrElse row.Item("Name").ToString.Trim = "All UE Terminals") Then
                            Send("<option value=""" & row.Item("TID") & """ style=""font-weight:bold;color:brown;"">" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID) & "</option>")
                        Else
                            Send("<option value=""" & row.Item("TID") & """>" & row.Item("Name") & "</option>")
                        End If
                    Next
                    If (BannersEnabled) Then Send("</optgroup>")
                End If
                Send("</select>")
                Send("<br />")
                Send("<br class=""half"" />")
          
                ' EXCLUDED TERMINALS
                Send("<div>")
                Send("<label for=""terminals-exclude""><b>" & Copient.PhraseLib.Lookup("term.excluded", LanguageID) & ":</b></label>")
                Send("<br />")
          
                'Buttons
                Sendb("<input type=""submit"" class=""regular select"" id=""terminals-add2"" name=""terminals-add2"" title=""" & Copient.PhraseLib.Lookup("term.exclude", LanguageID) & """ value=""&#9660; " & Copient.PhraseLib.Lookup("term.select", LanguageID) & """")
                If Not IsTemplate Then
                    If (Not Logix.UserRoles.CRUDLocationsToOffers) OrElse (rstSelected.Rows.Count = 0) OrElse (rstSelected.Rows(0).Item("AnyTerminal") = False) OrElse (bEnableRestrictedAccessToUEOfferBuilder AndAlso isTranslatedOffer) OrElse (bEnableAdditionalLockoutRestrictionsOnOffers AndAlso Not bOfferEditable) Then
                        Sendb(" disabled=""disabled""")
                    End If
                Else
                    If (Not Logix.UserRoles.CRUDLocationsToTemplates) OrElse (rstSelected.Rows.Count = 0) OrElse (rstSelected.Rows(0).Item("AnyTerminal") = False) OrElse (bEnableRestrictedAccessToUEOfferBuilder AndAlso isTranslatedOffer) OrElse (bEnableAdditionalLockoutRestrictionsOnOffers AndAlso Not bOfferEditable) Then
                        Sendb(" disabled=""disabled""")
                    End If
                End If
                If (Logix.UserRoles.EditOffer = False Or IsOfferWaitingForApproval(OfferID) Or Not (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID)) OrElse (bEnableRestrictedAccessToUEOfferBuilder AndAlso isTranslatedOffer) OrElse (bEnableAdditionalLockoutRestrictionsOnOffers AndAlso Not bOfferEditable)) Then
                    Sendb("disabled=""disabled""")
                End If
                Send(" />")
                Sendb("<input type=""submit"" class=""regular deselect"" id=""terminals-rem2"" name=""terminals-rem2"" title=""" & Copient.PhraseLib.Lookup("term.unexclude", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.deselect", LanguageID) & " &#9650;""")
                If Not IsTemplate Then
                    If (Not Logix.UserRoles.CRUDLocationsToOffers) OrElse (rstExcluded.Rows.Count = 0) OrElse (bEnableRestrictedAccessToUEOfferBuilder AndAlso isTranslatedOffer) OrElse (bEnableAdditionalLockoutRestrictionsOnOffers AndAlso Not bOfferEditable) OrElse Not IsOfferWaitingForApproval(OfferID) Then
                        Sendb(" disabled=""disabled""")
                    End If
                Else
                    If (Not Logix.UserRoles.CRUDLocationsToTemplates) OrElse (rstExcluded.Rows.Count = 0) OrElse (bEnableRestrictedAccessToUEOfferBuilder AndAlso isTranslatedOffer) OrElse (bEnableAdditionalLockoutRestrictionsOnOffers AndAlso Not bOfferEditable) Then
                        Sendb(" disabled=""disabled""")
                    End If
                End If
                Send(" />")
                Send("<br />")
          
                ' List
                Send("<select class=""longest"" multiple=""multiple"" id=""terminals-exclude"" name=""terminals-exclude"" size=""5"">")
                For Each row In rstExcluded.Rows
                    If (row.Item("Name").ToString.Trim = "All CPE Terminals" OrElse row.Item("Name").ToString.Trim = "All CM Terminals" OrElse row.Item("Name").ToString.Trim = "All DP Terminals" OrElse row.Item("Name").ToString.Trim = "All UE Terminals") Then
                        Send("<option value=""" & row.Item("PKID") & """ style=""font-weight:bold;color:brown;"">" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID) & "</option>")
                    Else
                        Send("<option value=""" & row.Item("PKID") & """>" & row.Item("Name") & "</option>")
                    End If
                Next
                Send("</select>")
                Send("</div>")
            %>
        </div>
    </div>
    <br clear="all" />
</div>
</form>
<script type="text/javascript">

    $(document).ready(function() {
        var savedTimeVal = document.getElementById('savedTime');
        var offerIDVal = document.getElementById('OfferID');
        if(savedTimeVal != null && offerIDVal != null) {
            var savedTime = new Date(savedTimeVal.value).getTime();
            var presentTime = new Date().getTime();
            var seconds = (presentTime - savedTime) / 1000;
            if(seconds > 2){
                $.support.cors = true;
                $.ajax({
                    type: "POST",
                    url: "/Connectors/AjaxProcessingFunctions.asmx/GetLockedSystemOptions",
                    data: JSON.stringify({ offerID : offerIDVal.value }),
                    contentType: "application/json; charset=utf-8",
                    dataType: "json"
                })
                .done(function (data) {
                    if(data.d == "true"){
                        window.location.href = window.location.href.replace("UEOffer-loc.aspx", "UEOffer-sum.aspx");
                    }
                });
            }
        }
    });

<% If (Request.QueryString("save") <> "" OrElse Request.QueryString("SaveThenClose") = "true") Then%>
    top.drawStores(<%Sendb(OfferID)%>);
  <%If (CloseAfterSave) Then%>
   top.closeIframePopup(false);
  <% End If %>
<%  End If %>

<% If (BannersEnabled) Then %>
    //handleKeyUp(9999);
<% End If %>
  //handleKeyUp2(9999);
  
  if (document.getElementById("sgroups-available") != null) {
    fullAvailList = document.getElementById("sgroups-available").cloneNode(true);
  }
  
  <% If (infoMessage <> "") Then %>
    // DOM2
    if (typeof window.addEventListener != "undefined")
      window.addEventListener( "load", JumpToTop, false );
    // IE
    else if (typeof window.attachEvent != "undefined") {
      window.attachEvent( "onload", JumpToTop );
    } else {
      if (window.onload != null) {
        var oldOnload = window.onload;
        window.onload = function ( e ) {
        oldOnload(e);
          JumpToTop();
        };
      } else
        window.onload = InitialiseScrollableArea;
    }
    function JumpToTop() {
      try {
        window.location.hash = 'h00';
        document.getElementById('main').scrollTop = 0;
      } catch (err) {
        // ignore
      }
    }
  <% End If %>
</script>
<%  
    If MyCommon.Fetch_SystemOption(75) Then
        If (OfferID > 0 AndAlso Logix.UserRoles.AccessNotes) Then
            Send_Notes(3, OfferID, AdminUserID)
        End If
    End If
done:
    Send_BodyEnd()
    'Send_BodyEnd("mainform", IIf(FocusType = 0, "functioninput", "functioninput2"))
    MyCommon.Close_LogixRT()
    MyCommon = Nothing
    Logix = Nothing
%>
