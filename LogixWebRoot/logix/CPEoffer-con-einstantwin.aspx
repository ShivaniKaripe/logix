﻿<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: CPEoffer-con-einstantwin.aspx 
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
  Dim MyCPEOffer As New Copient.EIW
  Dim Logix As New Copient.LogixInc
  Dim OfferID As Long
  Dim UpdateLevel As Integer = 0
  Dim Name As String = ""
  Dim IncentiveEIWID As Long
  Dim IsTemplate As Boolean = False
  Dim Disallow_Edit As Boolean = False
  Dim DisabledAttribute As String = ""
  Dim FromTemplate As Boolean = False
  Dim row As DataRow
  Dim rst As DataTable
  Dim rst2 As DataTable
  Dim x As Integer = 0
  Dim i As Integer = 0
  Dim CloseAfterSave As Boolean = False
  Dim historyString As String = ""
  Dim roid As Integer
  Dim infoMessage As String = ""
  Dim noteMessage As String = ""
  Dim Handheld As Boolean = False
  Dim EngineID As Integer = 2
  Dim EngineSubTypeID As Integer = 1
  Dim BannersEnabled As Boolean = True
  Dim TempInt As Integer
  Dim NumberOfPrizes As Integer = 0
  Dim FrequencyID As Integer = 1
  Dim StartDateTime As DateTime
  Dim EndDateTime As DateTime
  Dim CurrentDateTime As DateTime
  
  Dim TriggersTotal As Integer = 0
  Dim TriggersUsed As Integer = 0
  Dim TriggersRemaining As Integer = 0
  
  Dim Days As Integer = 0
  Dim Weeks As Integer = 0
  
  Dim PrevTriggersRemaining As Integer = 0
  Dim AddQty As Integer = 0
  Dim RemoveQty As Integer = 0
  Dim RemainingRandomized As Integer = 0
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "CPEoffer-con-instantwin.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  OfferID = Request.QueryString("OfferID")
  IncentiveEIWID = Request.QueryString("IncentiveEIWID")
  'Get the EngineID
  If (Request.QueryString("EngineID") <> "") Then
    EngineID = MyCommon.Extract_Val(Request.QueryString("EngineID"))
  Else
    MyCommon.QueryStr = "select EngineID from OfferIDs where OfferID=" & OfferID & ";"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
      EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), 0)
    End If
  End If
  'Get the ROID
  MyCommon.QueryStr = "select RewardOptionID from CPE_RewardOptions with (NoLock) " & _
                      "where IncentiveID=" & OfferID & " and TouchResponse=0 and Deleted=0;"
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    roid = MyCommon.NZ(rst.Rows(0).Item("RewardOptionID"), 0)
  End If
  'Get the offer's update level (to indicate if it's ever been deployed)
  MyCommon.QueryStr = "select UpdateLevel from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID & ";"
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    UpdateLevel = MyCommon.NZ(rst.Rows(0).Item("UpdateLevel"), 0)
  End If
  'Get the number of triggers and the current server time
  MyCommon.QueryStr = "select" & _
                      "  (select count(IncentiveEIWID) from CPE_EIWTriggers with (NoLock) where RewardOptionID=" & roid & " and IncentiveEIWID=" & IncentiveEIWID & " and Removed=0) as TriggersTotal," & _
                      "  (select count(IncentiveEIWID) from CPE_EIWTriggers as T with (NoLock) inner join CPE_EIWTriggersUsed as TU on TU.TriggerID=T.TriggerID where RewardOptionID=" & roid & " and IncentiveEIWID=" & IncentiveEIWID & " and T.Removed=0) as TriggersUsed, " & _
                      "  getdate() as CurrentDateTime;"
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    TriggersTotal = rst.Rows(0).Item("TriggersTotal")
    TriggersUsed = rst.Rows(0).Item("TriggersUsed")
    TriggersRemaining = TriggersTotal - TriggersUsed
    CurrentDateTime = rst.Rows(0).Item("CurrentDateTime")
    'Bump the current time out an hour, to allow a little buffer time for deployment, etc.
    CurrentDateTime = CurrentDateTime.AddHours(1)
  End If
  'Get the offer start and end dates to use as the trigger range; if the start is already past, use the current date/time.
  MyCommon.QueryStr = "select StartDate, EndDate from CPE_Incentives with (NoLock) " & _
                      "where IncentiveID=" & OfferID & ";"
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    StartDateTime = MyCommon.NZ(rst.Rows(0).Item("StartDate"), "1/1/1900 00:00:00")
    EndDateTime = MyCommon.NZ(rst.Rows(0).Item("EndDate"), "1/1/1900 00:00:00")
    EndDateTime = EndDateTime.AddHours(23)
    EndDateTime = EndDateTime.AddMinutes(59)
    EndDateTime = EndDateTime.AddSeconds(59)
  End If
  If StartDateTime < CurrentDateTime Then
    StartDateTime = CurrentDateTime
  End If
  'Get prizes and frequency
  MyCommon.QueryStr = "select NumberOfPrizes, FrequencyID from CPE_IncentiveEIW with (NoLock) " & _
                      "where IncentiveEIWID=" & IncentiveEIWID & ";"
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    NumberOfPrizes = MyCommon.NZ(rst.Rows(0).Item("NumberOfPrizes"), 0)
    FrequencyID = MyCommon.NZ(rst.Rows(0).Item("FrequencyID"), 0)
  End If
  'Determine how many days and weeks are affected
  Days = DateDiff(DateInterval.Day, StartDateTime, EndDateTime.AddDays(1))
  Weeks = System.Math.Ceiling(Days / 7)
  
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' User clicked save
  If (Request.QueryString("save") <> "") Then
    If IsNumeric(Request.QueryString("NumberOfPrizes")) Then
      NumberOfPrizes = MyCommon.Extract_Val(Request.QueryString("NumberOfPrizes"))
    End If
    FrequencyID = MyCommon.Extract_Val(Request.QueryString("FrequencyID"))
    Disallow_Edit = IIf(Request.QueryString("Disallow_Edit") = "on", True, False)
    
    'Validate entries
    If Not Integer.TryParse(NumberOfPrizes, TempInt) OrElse TempInt <= 0 Then
      infoMessage = Copient.PhraseLib.Lookup("ueoffer-con-einstantwin.PrizeNumberError", LanguageID)
    Else
      If IncentiveEIWID = 0 Then
        'Create a new condition
        MyCommon.QueryStr = "dbo.pa_CPE_AddEIW"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@ROID", SqlDbType.Int).Value = roid
        MyCommon.LRTsp.Parameters.Add("@NumberOfPrizes", SqlDbType.Int).Value = NumberOfPrizes
        MyCommon.LRTsp.Parameters.Add("@FrequencyID", SqlDbType.Int).Value = FrequencyID
        MyCommon.LRTsp.Parameters.Add("@DisallowEdit", SqlDbType.Bit).Value = Disallow_Edit
        MyCommon.LRTsp.Parameters.Add("@IncentiveEIWID", SqlDbType.Int).Direction = ParameterDirection.Output
        MyCommon.LRTsp.ExecuteNonQuery()
        IncentiveEIWID = MyCommon.LRTsp.Parameters("@IncentiveEIWID").Value
        MyCommon.Close_LRTsp()
        'Create new triggers for the condition
        If (EndDateTime > StartDateTime) And (EndDateTime > Now) Then
          MyCPEOffer.GenerateTriggers(OfferID, roid, IncentiveEIWID, NumberOfPrizes, FrequencyID, StartDateTime, EndDateTime, True)
        End If
        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1, EnableImpressRpt=1, EnableRedeemRpt=1, SendIssuance=1 where IncentiveID=" & OfferID & ";"
        MyCommon.LRT_Execute()
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-instantwin", LanguageID))
        'Response.Redirect("CPEoffer-con-einstantwin.aspx?OfferID=" & OfferID & "&IncentiveEIWID=" & IncentiveEIWID)
        CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
      Else
        'Update an existing condition
        'NB: Now that blackout periods have been dropped, updating is not currently allowed, but I'm retaining the logic just in case.
        'Update the CPE_IncentiveEIW record
        MyCommon.QueryStr = "update CPE_IncentiveEIW set LastUpdate=getdate(), DisallowEdit=" & IIf(Disallow_Edit, "1", "0") & ", RequiredFromTemplate=0 " & _
                            "where RewardOptionID=" & roid & " and IncentiveEIWID=" & IncentiveEIWID & ";"
        MyCommon.LRT_Execute()
        'Remove all existing triggers for the condition before adding new ones
        MyCommon.QueryStr = "update CPE_EIWTriggers with (RowLock) set Removed=1, LastUpdate=getdate() " & _
                            "where IncentiveEIWID=" & IncentiveEIWID & " and TriggerID not in (select TriggerID from CPE_EIWTriggersUsed);"
        MyCommon.LRT_Execute()
        'Create new triggers for the condition
        If (EndDateTime > StartDateTime) And (EndDateTime > Now) Then
          MyCPEOffer.GenerateTriggers(OfferID, roid, IncentiveEIWID, NumberOfPrizes, FrequencyID, StartDateTime, EndDateTime, False, TriggersRemaining)
        End If
        'MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1, EnableImpressRpt=1, EnableRedeemRpt=1, SendIssuance=1 where IncentiveID=" & OfferID & ";"
        'MyCommon.LRT_Execute()
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-instantwin", LanguageID))
        CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1") AndAlso infoMessage = ""
      End If
    End If
  End If
  
  'Trigger manipulation submissions
  If (Request.QueryString("randomize") <> "") Then
    'Drop all existing triggers for the condition before adding new ones
    MyCommon.QueryStr = "update CPE_EIWTriggers with (RowLock) set Removed=1, LastUpdate=getdate() " & _
                        "where IncentiveEIWID=" & IncentiveEIWID & " and TriggerID not in (select TriggerID from CPE_EIWTriggersUsed);"
    MyCommon.LRT_Execute()
    'Create new triggers
    If (EndDateTime > StartDateTime) And (EndDateTime > Now) Then
      If (TriggersRemaining = 0 And UpdateLevel = 0) Then 'we treat randomization as all-new trigger creation
        MyCPEOffer.GenerateTriggers(OfferID, roid, IncentiveEIWID, NumberOfPrizes, FrequencyID, StartDateTime, EndDateTime, True)
        noteMessage = Copient.PhraseLib.Detokenize("ueoffer-con-einstantwin.TriggersGenerated", LanguageID, NumberOfPrizes)
      Else 'we handle this as normal randomization
        MyCPEOffer.GenerateTriggers(OfferID, roid, IncentiveEIWID, NumberOfPrizes, FrequencyID, StartDateTime, EndDateTime, False, TriggersRemaining)
        noteMessage = Copient.PhraseLib.Detokenize("ueoffer-con-einstantwin.TriggersRandomized", LanguageID, TriggersRemaining)
      End If
    End If
    'Update associated records
    MyCommon.QueryStr = "update CPE_IncentiveEIW with (RowLock) set LastUpdate=getdate() where IncentiveEIWID=" & IncentiveEIWID & ";"
    MyCommon.LRT_Execute()
    'MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1, EnableImpressRpt=1, EnableRedeemRpt=1, SendIssuance=1 where IncentiveID=" & OfferID & ";"
    'MyCommon.LRT_Execute()
    'Log activity and redirect
    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-instantwin-randomize", LanguageID))
   'The number of triggers remaining that account for any added or removed triggers before rerandomizing.     
   PrevTriggersRemaining = TriggersRemaining 'value passed in
   
   'Determine the new total for TriggersRemaining     
   MyCommon.QueryStr = "select" & _
                      "  (select count(IncentiveEIWID) from CPE_EIWTriggers with (NoLock) where RewardOptionID=" & roid & " and IncentiveEIWID=" & IncentiveEIWID & " and Removed=0) as TriggersTotal," & _
                      "  (select count(IncentiveEIWID) from CPE_EIWTriggers as T with (NoLock) inner join CPE_EIWTriggersUsed as TU on TU.TriggerID=T.TriggerID where RewardOptionID=" & roid & " and IncentiveEIWID=" & IncentiveEIWID & " and T.Removed=0) as TriggersUsed;"
 
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
        TriggersTotal = rst.Rows(0).Item("TriggersTotal")
        TriggersUsed = rst.Rows(0).Item("TriggersUsed")
        TriggersRemaining = TriggersTotal - TriggersUsed
    End If
    
    'Adjust the TriggersRemaining after rerandomization to include previously added/removed triggers.      
    If PrevTriggersRemaining > TriggersRemaining 'previously added triggers
        AddQty = PrevTriggersRemaining - TriggersRemaining
        RemainingRandomized = TriggersRemaining + AddQty    
        If AddQty > 0 Then
            MyCPEOffer.GenerateTriggers(OfferID, roid, IncentiveEIWID, AddQty, 1, StartDateTime, EndDateTime, False)
            'Update associated records
            MyCommon.QueryStr = "update CPE_IncentiveEIW with (RowLock) set LastUpdate=getdate() where IncentiveEIWID=" & IncentiveEIWID & ";"
            MyCommon.LRT_Execute()
            'Log activity and redirect
            MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-instantwin", LanguageID))
            noteMessage = Copient.PhraseLib.Detokenize("ueoffer-con-einstantwin.TriggersRandomized", LanguageID, RemainingRandomized)
            Response.Redirect("CPEoffer-con-einstantwin.aspx?OfferID=" & OfferID & "&IncentiveEIWID=" & IncentiveEIWID & "&noteMessage=" & noteMessage)
        End If
    ElseIf PrevTriggersRemaining < TriggersRemaining 'previously removed triggers
        RemoveQty = TriggersRemaining - PrevTriggersRemaining
        RemainingRandomized = TriggersRemaining - RemoveQty    
        If RemoveQty > 0 Then
          MyCommon.QueryStr = "update top ("& RemoveQty &") ct set ct.Removed=1, ct.UpdateLevel=1, ct.LastUpdate=getdate() " & _
          "  from CPE_EIWTriggers ct with (nolock) left outer join CPE_EIWTriggersUsed ctu with (nolock) " & _
                              "  on  ct.TriggerID=ctu.TriggerID where ctu.TriggerID is null and ct.IncentiveEIWID=" & IncentiveEIWID & " and ct.Removed=0 "                     
          MyCommon.LRT_Execute()
          MyCommon.QueryStr = "update CPE_IncentiveEIW with (RowLock) set LastUpdate=getdate() where IncentiveEIWID=" & IncentiveEIWID & ";"
          MyCommon.LRT_Execute()
          MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-instantwin", LanguageID))
          noteMessage = Copient.PhraseLib.Detokenize("ueoffer-con-einstantwin.TriggersRandomized", LanguageID, RemainingRandomized)
          Response.Redirect("CPEoffer-con-einstantwin.aspx?OfferID=" & OfferID & "&IncentiveEIWID=" & IncentiveEIWID & "&noteMessage=" & noteMessage)
        End If
    Else
        Response.Redirect("CPEoffer-con-einstantwin.aspx?OfferID=" & OfferID & "&IncentiveEIWID=" & IncentiveEIWID & "&noteMessage=" & noteMessage)
    End If
  ElseIf (Request.QueryString("remove") <> "") Then
    If IsNumeric(Request.QueryString("RemoveQty")) Then
      x = MyCommon.Extract_Val(Request.QueryString("RemoveQty"))
    End If
    If x <= TriggersRemaining Then
      ' CLOUDSOL-297: Enterprise Instant Win not working as expected rewards were not issuing the full 100 prizes per day 
      MyCommon.QueryStr = "update top ("& x &") ct set ct.Removed=1, ct.UpdateLevel=1, ct.LastUpdate=getdate() " & _
      "  from CPE_EIWTriggers ct with (nolock) left outer join CPE_EIWTriggersUsed ctu with (nolock) " & _
                          "  on  ct.TriggerID=ctu.TriggerID where ctu.TriggerID is null and ct.IncentiveEIWID=" & IncentiveEIWID & " and ct.Removed=0 "                     
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "update CPE_IncentiveEIW with (RowLock) set LastUpdate=getdate() where IncentiveEIWID=" & IncentiveEIWID & ";"
      MyCommon.LRT_Execute()
      'MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1, EnableImpressRpt=1, EnableRedeemRpt=1, SendIssuance=1 where IncentiveID=" & OfferID & ";"
      'MyCommon.LRT_Execute()
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-instantwin", LanguageID))
      noteMessage = Copient.PhraseLib.Detokenize("ueoffer-con-einstantwin.TriggerRemoved", LanguageID, x)
      Response.Redirect("CPEoffer-con-einstantwin.aspx?OfferID=" & OfferID & "&IncentiveEIWID=" & IncentiveEIWID & "&noteMessage=" & noteMessage)
    Else
      infoMessage = Copient.PhraseLib.Detokenize("ueoffer-con-einstantwin.RemoveTriggerError", LanguageID, x, TriggersRemaining)
    End If
  ElseIf (Request.QueryString("add") <> "") Then
    If IsNumeric(Request.QueryString("AddQty")) Then
      x = MyCommon.Extract_Val(Request.QueryString("AddQty"))
    End If
    If x > 0 Then
      MyCPEOffer.GenerateTriggers(OfferID, roid, IncentiveEIWID, x, 1, StartDateTime, EndDateTime, False)
      'Update associated records
      MyCommon.QueryStr = "update CPE_IncentiveEIW with (RowLock) set LastUpdate=getdate() where IncentiveEIWID=" & IncentiveEIWID & ";"
      MyCommon.LRT_Execute()
      'MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1, EnableImpressRpt=1, EnableRedeemRpt=1, SendIssuance=1 where IncentiveID=" & OfferID & ";"
      'MyCommon.LRT_Execute()
      'Log activity and redirect
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-instantwin", LanguageID))
      noteMessage = Copient.PhraseLib.Detokenize("ueoffer-con-einstantwin.TriggerAdded", LanguageID, x)
      Response.Redirect("CPEoffer-con-einstantwin.aspx?OfferID=" & OfferID & "&IncentiveEIWID=" & IncentiveEIWID & "&noteMessage=" & noteMessage)
    End If
  End If
  
  ' Load offer data
  MyCommon.QueryStr = "select IncentiveID, IsTemplate, ClientOfferID, IncentiveName as Name, CPE.Description, PromoClassID, CRMEngineID, Priority," & _
                      "StartDate, EndDate, EveryDOW, EligibilityStartDate, EligibilityEndDate, TestingStartDate, TestingEndDate, " & _
                      "P1DistQtyLimit, P1DistTimeType, P1DistPeriod, P2DistQtyLimit, P2DistTimeType, P2DistPeriod, P3DistQtyLimit, P3DistTimeType, P3DistPeriod, " & _
                      "EnableImpressRpt, EnableRedeemRpt, CreatedDate, CPE.LastUpdate, CPE.Deleted, CPEOARptDate, CPEOADeploySuccessDate, CPEOADeployRpt, " & _
                      "CRMRestricted, StatusFlag, OC.Description as CategoryName, IsTemplate, FromTemplate, EngineID, EngineSubTypeID " & _
                      "from CPE_Incentives as CPE with (NoLock) " & _
                      "left join OfferCategories as OC with (NoLock) on CPE.PromoClassID=OfferCategoryID " & _
                      "where IncentiveID=" & Request.QueryString("OfferID") & ";"
  rst = MyCommon.LRT_Select
  For Each row In rst.Rows
    Name = MyCommon.NZ(row.Item("Name"), "")
    IsTemplate = MyCommon.NZ(row.Item("IsTemplate"), False)
    FromTemplate = MyCommon.NZ(row.Item("FromTemplate"), False)
    EngineID = MyCommon.NZ(row.Item("EngineID"), 0)
    EngineSubTypeID = MyCommon.NZ(row.Item("EngineSubTypeID"), 0)
  Next
  
  ' Update templates permission if necessary
  If (Request.QueryString("save") <> "" AndAlso Request.QueryString("IsTemplate") = "IsTemplate") Then
    ' update the status bits for the templates
    Dim form_Disallow_Edit As Integer = 0
    If (Request.QueryString("Disallow_Edit") = "on") Then
      form_Disallow_Edit = 1
    End If
    MyCommon.QueryStr = "update CPE_IncentiveEIW with (RowLock) set DisallowEdit=" & form_Disallow_Edit & _
                        " where RewardOptionID=" & roid & ";"
    MyCommon.LRT_Execute()
  End If
  
  If (IsTemplate Or FromTemplate) Then
    ' Load permissions if it's a template
    MyCommon.QueryStr = "select DisallowEdit from CPE_IncentiveEIW with (NoLock) where RewardOptionID=" & roid & ";"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      Disallow_Edit = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit"), False)
    Else
      Disallow_Edit = False
    End If
  End If
  
  If infoMessage = "" Then
    noteMessage = Request.QueryString("infoMessage")
  End If
  If noteMessage = "" Then
    noteMessage = Request.QueryString("noteMessage")
  End If
  
  If Not isTemplate Then
    DisabledAttribute = IIf(Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Edit), "", " disabled=""disabled""")
  Else
    DisabledAttribute = IIf(Logix.UserRoles.EditTemplates, "", " disabled=""disabled""")
  End If
  
  Send_HeadBegin("term.offer", "term.epriseinstantwincondition", OfferID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
  Send("<script type=""text/javascript"">")
  Send("function ChangeParentDocument() { ")
  If (EngineID = 3) Then
    Send("  opener.location = 'web-offer-con.aspx?OfferID=" & OfferID & "'; ")
  ElseIf (EngineID = 5) Then
    Send("  opener.location = 'email-offer-con.aspx?OfferID=" & OfferID & "'; ")
  ElseIf (EngineID = 6) Then
    Send("  opener.location = 'CAM/CAM-offer-con.aspx?OfferID=" & OfferID & "'; ")
  Else
    Send("  opener.location = 'CPEoffer-con.aspx?OfferID=" & OfferID & "'; ")
  End If
  Send("} ")
  Send("</script>")
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
<form action="#" id="mainform" name="mainform">
<div id="intro">
  <input type="hidden" id="OfferID" name="OfferID" value="<% Sendb(OfferID) %>" />
  <input type="hidden" id="IsTemplate" name="IsTemplate" value="<% Sendb(IIf(IsTemplate, "IsTemplate", "Not")) %>" />
  <input type="hidden" id="IncentiveEIWID" name="IncentiveEIWID" value="<% Sendb(IncentiveEIWID) %>" />
  <%
    If (IsTemplate) Then
      Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.epriseinstantwincondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
    Else
      Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.epriseinstantwincondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
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
      If Not IsTemplate Then
        If (Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Edit)) Then
          If IncentiveEIWID = 0 Then
            Send_Save()
          End If
        End If
      Else
        If (Logix.UserRoles.EditTemplates) Then
          If IncentiveEIWID = 0 Then
            Send_Save()
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
    ElseIf (noteMessage <> "") Then
      Send("<div id=""infobar"" class=""green-background"">" & noteMessage & "</div>")
    End If
  %>
  <div id="column1">
    <div class="box" id="general">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.general", LanguageID))%>
        </span>
      </h2>
      <%
        MyCommon.QueryStr = "select NumberOfPrizes, FrequencyID, LastUpdate from CPE_IncentiveEIW as EIW with (NoLock) " & _
                            "where RewardOptionID=" & roid & " and IncentiveEIWID=" & IncentiveEIWID & ";"
        rst = MyCommon.LRT_Select
        Send("<table summary=""" & Copient.PhraseLib.Lookup("term.general", LanguageID) & """>")
        Send("  <tr>")
        Send("    <td><label for=""NumberOfPrizes"">" & Copient.PhraseLib.Lookup("CPEoffer-con-einstantwin.NumberOfPrizes", LanguageID) & ":</label></td>")
        Send("    <td>")
        If IncentiveEIWID = 0 Then
          Sendb("      <input type=""text"" class=""shorter"" id=""NumberOfPrizes"" name=""NumberOfPrizes"" maxlength=""4"" value=""")
          If rst.Rows.Count > 0 Then
            Sendb(MyCommon.NZ(rst.Rows(0).Item("NumberOfPrizes"), 0))
          End If
          Send(""" />")
          Send("      <select id=""FrequencyID"" name=""FrequencyID"">")
          MyCommon.QueryStr = "select FrequencyID, Description, PhraseID from CPE_IncentiveEIWFrequency with (NoLock) order by FrequencyID;"
          rst2 = MyCommon.LRT_Select
          If rst2.Rows.Count > 0 Then
            For Each row In rst2.Rows
              Send("        <option value=""" & MyCommon.NZ(row.Item("FrequencyID"), 0) & """>" & MyCommon.NZ(row.Item("Description"), "") & "</option>")
            Next
          End If
          Send("      </select>")
        Else
          Send("      " & MyCommon.NZ(rst.Rows(0).Item("NumberOfPrizes"), 0) & " ")
          MyCommon.QueryStr = "select Description, PhraseID from CPE_IncentiveEIWFrequency with (NoLock) where FrequencyID=" & MyCommon.NZ(rst.Rows(0).Item("FrequencyID"), 0) & ";"
          rst2 = MyCommon.LRT_Select
          If rst2.Rows.Count > 0 Then
            Send(StrConv(Copient.PhraseLib.Lookup(MyCommon.NZ(rst2.Rows(0).Item("PhraseID"), 0), LanguageID, MyCommon.NZ(rst2.Rows(0).Item("Description"), "")), VbStrConv.Lowercase))
          End If
          Send("      <input type=""hidden"" id=""NumberOfPrizes"" name=""NumberOfPrizes"" value=""" & MyCommon.NZ(rst.Rows(0).Item("NumberOfPrizes"), 0) & """ />")
        End If
        Send("    </td>")
        Send("  </tr>")
        If rst.Rows.Count > 0 Then
          Send("  <tr>")
          Send("    <td>" & Copient.PhraseLib.Lookup("term.lastupdated", LanguageID) & ":</td>")
          Send("    <td>" & MyCommon.NZ(rst.Rows(0).Item("LastUpdate"), "") & "</td>")
          Send("  </tr>")
        End If
        Send("</table>")
      %>
      <hr class="hidden" />
    </div>
    
    <div class="box" id="triggerrange">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.trigger", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.range", LanguageID), VbStrConv.Lowercase))%>
        </span>
      </h2>
    <%
      Send("<p>" & Copient.PhraseLib.Lookup("ueoffer-con-einstantwin.TriggerMessage", LanguageID) & "</p>")
      Send("<table summary=""" & Copient.PhraseLib.Lookup("term.range", LanguageID) & """>")
      Send("  <tr>")
      Send("    <td>" & Copient.PhraseLib.Lookup("ueoffer-con-einstantwin.TriggerRangeBegins", LanguageID) & "</td>")
      Send("    <td>" & Logix.ToLongDateTimeString(StartDateTime, MyCommon) & "</td>")
      Send("  </tr>")
      Send("  <tr>")
      Send("    <td>" & Copient.PhraseLib.Lookup("ueoffer-con-einstantwin.TriggerRangeEnds", LanguageID) & "</td>")
      Send("    <td>" & Logix.ToLongDateTimeString(EndDateTime, MyCommon) & "</td>")
      Send("  </tr>")
      Send("  <tr>")
      Send("    <td>" & Copient.PhraseLib.Lookup("ueoffer-con-einstantwin.TriggerRangeExtent", LanguageID) & "</td>")
      Send("    <td>" & Copient.PhraseLib.Detokenize("ueoffer-con-einstantwin.TriggerMessage2", LanguageID, Days, Weeks) & "</td>")
      Send("  </tr>")
      MyCommon.QueryStr = "select I.DOWID, D.DayName from CPE_IncentiveDOW as I with (NoLock) " & _
                          "left join CPE_DaysOfWeek as D on D.DOWID=I.DOWID " & _
                          "where IncentiveID=" & OfferID & " and Deleted=0;"
      rst = MyCommon.LRT_Select
      MyCommon.QueryStr = "select * from CPE_IncentiveTOD with (NoLock) " & _
                          "where IncentiveID=" & OfferID & ";"
      rst2 = MyCommon.LRT_Select
      If (rst.Rows.Count > 0) Or (rst2.Rows.Count > 0) Then
        Send("<tr>")
        Send("  <td colspan=""2""><i>" & Copient.PhraseLib.Lookup("term.limitations", LanguageID) & ":</i></td>")
        Send("</tr>")
        If rst.Rows.Count > 0 Then
          Send("  <tr>")
          Send("    <td valign=""top"">" & Copient.PhraseLib.Lookup("ueoffer-con-einstantwin.OnlyTheseDays", LanguageID) & "</td>")
          Sendb("    <td>")
          i = 1
          For Each row In rst.Rows
            Sendb(MyCommon.NZ(row.Item("DayName"), ""))
            If i < rst.Rows.Count Then
              Sendb(", ")
            End If
            i += 1
          Next
          Send("</td>")
          Send("  </tr>")
        End If
        If rst2.Rows.Count > 0 Then
          Send("  <tr>")
          Send("    <td valign=""top"">" & Copient.PhraseLib.Lookup("ueoffer-con-einstantwin.OnlyTheseTimes", LanguageID) & "</td>")
          Sendb("    <td>")
          i = 1
          For Each row In rst2.Rows
            Sendb(MyCommon.NZ(row.Item("StartTime"), "") & " - " & MyCommon.NZ(row.Item("EndTime"), ""))
            If i < rst2.Rows.Count Then
              Sendb("<br />")
            End If
            i += 1
          Next
          Send("</td>")
          Send("  </tr>")
        End If
      End If
      Send("</table>")
    %>
    </div>
  </div>
  
  <div id="gutter">
  </div>
  
  <div id="column2"<% Sendb(IIf(IncentiveEIWID = 0, " style=""display:none;""", "")) %>>
    <div class="box" id="triggers">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.triggers", LanguageID))%>
        </span>
      </h2>
      <%
        Send("<table summary=""" & Copient.PhraseLib.Lookup("term.triggers", LanguageID) & """>")
        Send("  <tr style=""background-color:#eeeeff;"">")
        Send("    <td>" & Copient.PhraseLib.Lookup("term.TotalTriggers", LanguageID) & ":</td>")
        Send("    <td>" & TriggersTotal & "</td>")
        Send("  </tr>")
        Send("  <tr style=""background-color:#ffeeee;"">")
        Send("    <td>" & Copient.PhraseLib.Lookup("term.UsedTriggers", LanguageID) & ":</td>")
        Send("    <td>" & TriggersUsed & "</td>")
        Send("  </tr>")
        Send("  <tr style=""background-color:#eeffee;"">")
        Send("    <td>" & Copient.PhraseLib.Lookup("term.UnusedTriggers", LanguageID) & ":</td>")
        Send("    <td>" & TriggersRemaining & "</td>")
        Send("  </tr>")
        Send("</table>")
          
        Send("<table summary=""" & Copient.PhraseLib.Lookup("term.administration", LanguageID) & """>")
        Send("  <tr>")
        Send("<hr />")
        Send("    <td><label for=""AddQty"">" & Copient.PhraseLib.Lookup("term.TriggersToAdd", LanguageID) & ":</label></td>")
        Send("    <td><input type=""text"" class=""shortest"" id=""AddQty"" name=""AddQty"" maxlength=""4"" value="""" /></td>")
        Send("    <td><input type=""submit"" id=""Add"" name=""Add"" style=""width:85px;"" value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """ /></td>")
        Send("  </tr>")
        Send("  <tr>")
        Send("    <td><label for=""RemoveQty"">" & Copient.PhraseLib.Lookup("term.TriggersToRemove", LanguageID) & ":</label></td>")
        Send("    <td><input type=""text"" class=""shortest"" id=""RemoveQty"" name=""RemoveQty"" maxlength=""9"" value="""" /></td>")
        Send("    <td><input type=""submit"" id=""Remove"" name=""Remove"" style=""width:85px;"" value=""" & Copient.PhraseLib.Lookup("term.remove", LanguageID) & """ /></td>")
        Send("  </tr>")
        Send("  <tr>")
        Send("    <td colspan=""2""><label for=""Randomize"">" & Copient.PhraseLib.Lookup("ueoffer-con-einstantwin.RandomizeUnused", LanguageID) & "</label></td>")
        Send("    <td><input type=""submit"" id=""Randomize"" name=""Randomize"" style=""width:85px;"" value=""" & Copient.PhraseLib.Lookup("term.randomize", LanguageID) & """ /></td>")
        Send("  </tr>")
        Send("</table>")
      %>
      <hr class="hidden" />
    </div>
    
  </div>
</div>
</form>

<script type="text/javascript">
<% If (CloseAfterSave) Then %>
    window.close();
<% End If %>
</script>
<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd("mainform", "NumPrizesAllowed")
  MyCommon = Nothing
  Logix = Nothing
%>
