<%@ Page Language="vb" Debug="true" CodeFile="ue-cb.vb" Inherits="UECB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
    ' *****************************************************************************
    ' * FILENAME: UEoffer-con-time.aspx 
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
    Dim MyCPEOffer As New Copient.EIW
    Dim Logix As New Copient.LogixInc
    Dim OfferID As Long
    Dim Name As String = ""
    Dim ConditionID As String
    Dim IsTemplate As Boolean = False
    Dim Disallow_Edit As Boolean = True
    Dim DisabledAttribute As String = ""
    Dim FromTemplate As Boolean = False
    Dim i As Integer
    Dim row As DataRow
    Dim rst As DataTable
    Dim rst2 As DataTable
    Dim row2 As DataRow
    Dim CloseAfterSave As Boolean = False
    Dim historyString As String = ""
    Dim roid As Integer
    Dim slot As Integer = 0
    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False
    Dim EngineID As Integer = 2
    Dim BannersEnabled As Boolean = True
    Dim StartTime As String
    Dim StartHour As Integer
    Dim StartMinute As Integer
    Dim StartHourEntry As String = ""
    Dim StartMinEntry As String = ""
    Dim EndTime As String
    Dim EndHour As Integer
    Dim EndMinute As Integer
    Dim EndHourEntry As String = ""
    Dim EndMinEntry As String = ""
    Dim TimeDisplay As String = "none"
    Dim TimeChecked As String = ""
    Dim TimeEntry1 As String = ""
    Dim TimeEntry2 As String = ""
    Dim TimeEntryType As Integer = 1
    Dim IncentiveTODID As Long
    Dim StartMeridiem As String = ""
    Dim EndMeridiem As String = ""
    Dim TimeMap As New BitArray(1440, False)
    Dim RangeStart As Integer
    Dim RangeEnd As Integer
    Dim Overlaps As Boolean
    Dim OfferHasEIW As Boolean = False

    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If

    Response.Expires = 0
    MyCommon.AppName = "UEoffer-con-time.aspx"
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

    MyCommon.QueryStr = "select RewardOptionID from CPE_RewardOptions with (NoLock) where IncentiveID=" & OfferID & " and touchresponse=0 and deleted=0;"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
        roid = rst.Rows(0).Item("RewardOptionID")
    End If

    'Determine if the offer has an enterprise instant win condition
    MyCommon.QueryStr = "select IncentiveEIWID from CPE_IncentiveEIW with (NoLock) where RewardOptionID=" & roid & ";"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
        OfferHasEIW = True
    End If

    ' see if someone is saving
    If (Request.QueryString("save") <> "" OrElse Request.QueryString("saveTimeSlot") = "1") Then

        'store the existing locking value for use in newly-created records
        MyCommon.QueryStr = "select DisallowEdit from CPE_IncentiveTOD with (NoLock) where IncentiveID=" & OfferID
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
            Disallow_Edit = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit"), False)
        Else
            Disallow_Edit = False
        End If

        ' delete all the time slots for this incentive
        If (Request.QueryString("startHour") <> "" OrElse Request.QueryString("startMinute") <> "" OrElse Request.QueryString("endHour") <> "" OrElse Request.QueryString("endMinute") <> "") Then

            ' check for valid time character entry

            If Request.QueryString("timeentry") = 1 AndAlso (Not Integer.TryParse(Request.QueryString("startHour"), StartHour) OrElse (StartHour < 0 OrElse StartHour > 12)) Then
                infoMessage = Copient.PhraseLib.Lookup("CPEoffer-con-time.start-12hour-msg", LanguageID)
            ElseIf Request.QueryString("timeentry") = 2 AndAlso (Not Integer.TryParse(Request.QueryString("startHour"), StartHour) OrElse (StartHour < 0 OrElse StartHour > 23)) Then
                infoMessage = Copient.PhraseLib.Lookup("CPEoffer-con-time.start-hour-msg", LanguageID)
            ElseIf Not Integer.TryParse(Request.QueryString("startMinute"), StartMinute) OrElse StartMinute < 0 OrElse StartMinute > 59 Then
                infoMessage = Copient.PhraseLib.Lookup("CPEoffer-con-time.start-minute-msg", LanguageID)
            ElseIf Request.QueryString("timeentry") = 1 AndAlso (Not Integer.TryParse(Request.QueryString("endHour"), EndHour) OrElse (EndHour < 0 OrElse EndHour > 12)) Then
                infoMessage = Copient.PhraseLib.Lookup("CPEoffer-con-time.end-12hour-msg", LanguageID)
            ElseIf Request.QueryString("timeentry") = 2 AndAlso (Not Integer.TryParse(Request.QueryString("endHour"), EndHour) OrElse (EndHour < 0 OrElse EndHour > 23)) Then
                infoMessage = Copient.PhraseLib.Lookup("CPEoffer-con-time.end-hour-msg", LanguageID)
            ElseIf Not Integer.TryParse(Request.QueryString("endMinute"), EndMinute) OrElse EndMinute < 0 OrElse EndMinute > 59 Then
                infoMessage = Copient.PhraseLib.Lookup("CPEoffer-con-time.end-minute-msg", LanguageID)
            End If

            If (infoMessage = "" AndAlso Request.QueryString("timeentry") = "1") Then
                If (Request.QueryString("startMeridiem") = "1" AndAlso StartHour = 12) Then
                    StartHour = 0
                ElseIf (Request.QueryString("startMeridiem") = "2" AndAlso StartHour < 12) Then
                    StartHour += 12
                End If
                If (Request.QueryString("endMeridiem") = "1" AndAlso EndHour = 12) Then
                    EndHour = 0
                ElseIf (Request.QueryString("endMeridiem") = "2" AndAlso EndHour < 12) Then
                    EndHour += 12
                End If
            End If

            ' check that end time is after start time
            If (infoMessage = "" AndAlso Integer.Parse(StartHour.ToString.PadLeft(2, "0") & StartMinute.ToString.PadLeft(2, "0")) >= Integer.Parse(EndHour.ToString.PadLeft(2, "0") & EndMinute.ToString.PadLeft(2, "0"))) Then
                infoMessage = Copient.PhraseLib.Lookup("CPEoffer-con-time.end-before-start-msg", LanguageID)
            End If
            If (infoMessage = "") Then
                ' check for any overlap of existing time ranges
                MyCommon.QueryStr = "select cast(left(IsNull(StartTime,'00'), 2) as int) * 60 + cast (right(IsNull(StartTime,'00'), 2) as int) as StartMinute, " & _
                                    "  cast(left(IsNull(EndTime,'00'), 2) as int) * 60 + cast (right(IsNull(EndTime,'00'), 2) as int) as EndMinute " & _
                                    "from CPE_IncentiveTOD with (NoLock) where IncentiveId = " & OfferID
                rst = MyCommon.LRT_Select
                If (rst.Rows.Count > 0) Then
                    For Each row In rst.Rows
                        RangeStart = MyCommon.NZ(row.Item("StartMinute"), -1)
                        RangeEnd = MyCommon.NZ(row.Item("EndMinute"), -1)
                        If (RangeStart > -1 AndAlso RangeStart < TimeMap.Length) AndAlso (RangeEnd > -1 AndAlso RangeEnd < TimeMap.Length) AndAlso (RangeStart <= RangeEnd) Then
                            For i = RangeStart To RangeEnd
                                TimeMap.Set(i, True)
                            Next
                        End If
                    Next
                End If
                ' check for any overlap of existing time ranges
                For i = (StartHour * 60) + StartMinute To (EndHour * 60) + EndMinute
                    Overlaps = TimeMap.Get(i)
                    If Overlaps Then
                        infoMessage = Copient.PhraseLib.Lookup("CPEoffer-con-time.overlap", LanguageID)
                        Exit For
                    End If
                Next
            End If
            If (infoMessage = "") Then
                StartTime = Left(StartHour.ToString, 2).PadLeft(2, "0") & ":" & Left(StartMinute.ToString, 2).PadLeft(2, "0")
                EndTime = Left(EndHour.ToString, 2).PadLeft(2, "0") & ":" & Left(EndMinute.ToString, 2).PadLeft(2, "0")

                MyCommon.QueryStr = "dbo.pt_CPE_IncentiveTOD_Insert"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@IncentiveID", SqlDbType.BigInt).Value = OfferID
                MyCommon.LRTsp.Parameters.Add("@StartTime", SqlDbType.NVarChar, 5).Value = StartTime
                MyCommon.LRTsp.Parameters.Add("@EndTime", SqlDbType.NVarChar, 5).Value = EndTime
                MyCommon.LRTsp.Parameters.Add("@IncentiveTODID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                MyCommon.LRTsp.ExecuteNonQuery()
                MyCommon.Close_LRTsp()
                MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-time-add", LanguageID))
            End If
        End If
        If (infoMessage = "") Then
            'Set LastUpdate and record history:
            MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID
            MyCommon.LRT_Execute()
            MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-time-edit", LanguageID))
            ResetOfferApprovalStatus(OfferID)
            'Randomize any EIW triggers associated with this offer:
            If OfferHasEIW Then
                MyCPEOffer.RandomizeTriggersByOffer(OfferID)
                MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-instantwin-randomize", LanguageID))
            End If
            CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1") AndAlso Request.QueryString("saveTimeSlot") <> "1" AndAlso infoMessage = ""
        End If

    ElseIf MyCommon.Extract_Val(Request.QueryString("DeleteSlot")) > 0 Then
        MyCommon.QueryStr = "dbo.pt_CPE_IncentiveTOD_Delete"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@IncentiveTODID", SqlDbType.BigInt).Value = MyCommon.Extract_Val(Request.QueryString("DeleteSlot"))
        MyCommon.LRTsp.ExecuteNonQuery()
        MyCommon.Close_LRTsp()
        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID
        MyCommon.LRT_Execute()
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-time-delete", LanguageID))
        ResetOfferApprovalStatus(OfferID)
        'Randomize any EIW triggers associated with this offer:
        If OfferHasEIW Then
            MyCPEOffer.RandomizeTriggersByOffer(OfferID)
            MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-instantwin-randomize", LanguageID))
        End If
    End If

    ' reset the EveryTOD column to reflect the change, if no time is selected then set to 1, otherwise set to 0
    MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set EveryTOD = " & _
                        "   case (select count(*) TimeCount from CPE_IncentiveTOD where IncentiveID=" & OfferID & ") " & _
                        "       when 0 then 1 " & _
                        "       else 0 " & _
                        "   end " & _
                        "where incentiveid = " & OfferID & " and deleted=0;"
    MyCommon.LRT_Execute()

    ' dig the offer info out of the database
    ' no one clicked anything
    MyCommon.QueryStr = "Select IncentiveID,IsTemplate,ClientOfferID,IncentiveName as Name,CPE.Description,PromoClassID,CRMEngineID,Priority," & _
                        "StartDate,EndDate,EveryDOW,EligibilityStartDate,EligibilityEndDate,TestingStartDate,TestingEndDate,P1DistQtyLimit,P1DistTimeType,P1DistPeriod," & _
                        "P2DistQtyLimit,P2DistTimeType,P2DistPeriod,P3DistQtyLimit,P3DistTimeType,P3DistPeriod,EnableImpressRpt,EnableRedeemRpt,CreatedDate,CPE.LastUpdate,CPE.Deleted, CPEOARptDate, CPEOADeploySuccessDate, CPEOADeployRpt," & _
                        "CRMRestricted,StatusFlag,OC.Description as CategoryName,IsTemplate,FromTemplate from CPE_Incentives as CPE with (NoLock) left join OfferCategories as OC with (NoLock) on CPE.PromoClassID=OfferCategoryID where IncentiveID=" & Request.QueryString("OfferID")
    rst = MyCommon.LRT_Select
    For Each row In rst.Rows
        Name = MyCommon.NZ(row.Item("Name"), "")
        IsTemplate = MyCommon.NZ(row.Item("IsTemplate"), False)
        FromTemplate = MyCommon.NZ(row.Item("FromTemplate"), False)
    Next

    'update the templates permission if necessary
    If (Request.QueryString("save") <> "" AndAlso Request.QueryString("IsTemplate") = "IsTemplate") Then
        ' time to update the status bits for the templates
        Dim form_Disallow_Edit As Integer = 0
        If (Request.QueryString("Disallow_Edit") = "on") Then
            form_Disallow_Edit = 1
        End If
        MyCommon.QueryStr = "update CPE_IncentiveTOD with (RowLock) set DisallowEdit=" & form_Disallow_Edit & _
                            " where IncentiveID=" & OfferID & ";"
        MyCommon.LRT_Execute()
    End If

    If (IsTemplate Or FromTemplate) Then
        ' lets dig the permissions if its a template
        MyCommon.QueryStr = "select DisallowEdit from CPE_IncentiveTOD with (NoLock) where IncentiveID=" & OfferID & ";"
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

    Send_HeadBegin("term.offer", "term.timecondition", OfferID)
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts()
%>
<script type="text/javascript" language="javascript">
//    function showTimeSlots(bChecked) {
//      var elemTimeSlot = document.getElementById("timeslotspan");
//      
//      if (elemTimeSlot != null) {
//        elemTimeSlot.style.display = (bChecked) ? "" : "none";
//      }
//    }
    
    function addTimeSlot() {
      var elemTime = document.getElementById("saveTimeSlot");
      
      if (elemTime != null) {
        elemTime.value = "1";
      }
      
      document.mainform.submit();        
    }
    
    function deleteTimeSlot(slotId) {
      var elemTime = document.getElementById("deleteSlot");
      
      if (elemTime != null) {
        elemTime.value = slotId;
      }
      
      document.mainform.submit();
    }
    
    function handleClockType(type) {
      var startMeridiem = document.getElementById("startMeridiem");
      var endMeridiem =  document.getElementById("endMeridiem");
      
      if (startMeridiem != null && endMeridiem != null) {
        startMeridiem.style.display = (type == 1) ? "inline" : "none";              
        endMeridiem.style.display = (type == 1) ? "inline" : "none";              
      }
      
      adjustTimeSlots(type, 'start');
      adjustTimeSlots(type, 'end');
    }
    
    function adjustTimeSlots(clockType, prefix) {
      var i = 0;
      var newHour = 0;
      var elemMeridiem = document.getElementById(prefix + "MeridiemSpan" + i);
      var elemHour = document.getElementById(prefix + "HourSpan" + i);
      
      while (elemMeridiem != null && elemHour != null) {
        newHour = parseInt(elemHour.innerHTML, 10)
                        
        if (clockType == 1) {
          // convert from 24 hour notation to 12 hour notation
          if (newHour == 0) {
            newHour = 12;
            elemMeridiem.innerHTML = "AM";
          } else if (newHour > 0 && newHour < 12) {
            elemMeridiem.innerHTML = "AM";
          } else if (newHour == 12) {
            elemMeridiem.innerHTML = "PM";
          } else if (newHour > 12) {
            newHour -= 12;
            elemMeridiem.innerHTML = "PM";
          } 
        } else if (clockType == 2) {
          // convert from 12 hour notation to 24 hour notation
          if (elemMeridiem.innerHTML == "PM" && newHour < 12) {
            newHour += 12
            elemHour.innerHTML = newHour;
          } else if (elemMeridiem.innerHTML == "AM" && newHour == 12) {
            newHour =0            
          }
          elemMeridiem.innerHTML = "";
        }
        elemHour.innerHTML = newHour;
        
        i++;  
        elemMeridiem = document.getElementById(prefix + "MeridiemSpan" + i);
        elemHour = document.getElementById(prefix + "HourSpan" + i);
      }
    }
    
    function handleTimeSlotEntry(e) {
      var keycode;
      
      if (window.event) keycode = window.event.keyCode;
      else if (e) keycode = e.which;
      else return true;
      
      if (keycode == 13) {
        addTimeSlot();
        return false;
      }
      
      return true;    
    }
</script>
<%
  Send("<script type=""text/javascript"">")
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
    <input type="hidden" id="OfferID" name="OfferID" value="<% sendb(OfferID) %>" />
    <input type="hidden" id="Name" name="Name" value="<% sendb(Name) %>" />
    <input type="hidden" id="ConditionID" name="ConditionID" value="<% sendb(ConditionID) %>" />
    <input type="hidden" id="IsTemplate" name="IsTemplate" value="<% 
        If (IsTemplate) Then
        Sendb("IsTemplate")
        Else
        Sendb("Not")
        End If
        %>" />
    <%If (IsTemplate) Then
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.timecondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
      Else
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.timecondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
      End If
    %>
    <div id="controls">
      <% If (IsTemplate) Then%>
      <span class="temp">
        <input type="checkbox" class="tempcheck" id="Disallow_Edit" name="Disallow_Edit"
          <% if(disallow_edit)then sendb(" checked=""checked""") %> />
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
      <div class="box" id="time" style="min-height: 300px; height: auto !important; height: 300px;">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.time", LanguageID))%>
          </span>
        </h2>
        <%
          i = 0
          
          MyCommon.QueryStr = "select IncentiveTODID, StartTime, EndTime from CPE_IncentiveTOD with (NoLock) where IncentiveID=" & OfferID & " order by StartTime, EndTime;"
          rst = MyCommon.LRT_Select
          
          TimeDisplay = IIf(rst.Rows.Count > 0, "block", "none")
          TimeChecked = IIf(rst.Rows.Count > 0, " checked=""checked"" ", "")
          
          If Request.QueryString("timeentry") <> "" Then
            TimeDisplay = "block"
            TimeChecked = " checked=""checked"""
            If (infoMessage <> "") Then
              StartHourEntry = Request.QueryString("startHour")
              StartMinEntry = Request.QueryString("startMinute")
              EndHourEntry = Request.QueryString("endHour")
              EndMinEntry = Request.QueryString("endMinute")
              StartMeridiem = Request.QueryString("startMeridiem")
              EndMeridiem = Request.QueryString("endMeridiem")
            End If
          End If
          
          TimeEntry1 = IIf(Request.QueryString("timeentry") = "1", " checked=""checked""", "")
          TimeEntry2 = IIf(Request.QueryString("timeentry") = "2", " checked=""checked""", "")
          If (TimeEntry1 = "" AndAlso TimeEntry2 = "") Then TimeEntry1 = " checked=""checked"""
          TimeEntryType = IIf(Request.QueryString("timeentry") = "2", 2, 1)
          
          Send("<div id=""timeslotspan"" name=""timeslotspan"">")
          Send("<input type=""radio"" name=""timeentry"" id=""timeentry1"" value=""1""" & TimeEntry1 & " onclick=""handleClockType(1);""" & DisabledAttribute & " /><label for=""timeentry1"">" & Copient.PhraseLib.Lookup("term.12HourNotation", LanguageID) & "</label>&nbsp;&nbsp;")
          Send("<input type=""radio"" name=""timeentry"" id=""timeentry2"" value=""2""" & TimeEntry2 & " onclick=""handleClockType(2);""" & DisabledAttribute & " /><label for=""timeentry2"">" & Copient.PhraseLib.Lookup("term.24HourNotation", LanguageID) & "</label><br/><br/>")
          Sendb("<input type=""text"" class=""shortest"" id=""startHour"" name=""startHour"" maxlength=""2"" value=""" & StartHourEntry & """ onkeypress=""handleTimeSlotEntry(event);""" & DisabledAttribute & " />:")
          Send("<input type=""text"" class=""shortest"" id=""startMinute"" name=""startMinute"" maxlength=""2"" value=""" & StartMinEntry & """ onkeypress=""handleTimeSlotEntry(event);""" & DisabledAttribute & " />")
          Send("<select class=""shorter"" id=""startMeridiem"" name=""startMeridiem"" onkeypress=""handleTimeSlotEntry(event);""" & DisabledAttribute & " >")
          Send("  <option value=""1""" & IIf(StartMeridiem = "1", " selected=""selected""", "") & ">AM</option>")
          Send("  <option value=""2""" & IIf(StartMeridiem = "2", " selected=""selected""", "") & ">PM</option>")
          Send("</select>")
          Send("&nbsp;" & Copient.PhraseLib.Lookup("term.to", LanguageID) & "&nbsp;")
          Sendb("<input type=""text"" class=""shortest"" id=""endHour"" name=""endHour"" maxlength=""2"" value=""" & EndHourEntry & """ onkeypress=""handleTimeSlotEntry(event);""" & DisabledAttribute & " />:")
          Send("<input type=""text"" class=""shortest"" id=""endMinute"" name=""endMinute"" maxlength=""2"" value=""" & EndMinEntry & """ onkeypress=""handleTimeSlotEntry(event);""" & DisabledAttribute & " />")
          Send("<select class=""shorter"" id=""endMeridiem"" name=""endMeridiem"" onkeypress=""handleTimeSlotEntry(event);""" & DisabledAttribute & " >")
          Send("  <option value=""1""" & IIf(EndMeridiem = "1", " selected=""selected""", "") & ">AM</option>")
          Send("  <option value=""2""" & IIf(EndMeridiem = "2", " selected=""selected""", "") & ">PM</option>")
          Send("</select>")
          Send("&nbsp;&nbsp;")
          If(Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) Then
           Send("<input type=""button"" id=""addSlot"" name=""addSlot"" onclick=""addTimeSlot();"" value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """ " & DisabledAttribute & " />")
          End If
          Send("<input type=""hidden"" id=""saveTimeSlot"" name=""saveTimeSlot"" value=""0"" />")
          Send("<input type=""hidden"" id=""deleteSlot"" name=""deleteSlot"" value=""0"" />")
          
          If (rst.Rows.Count > 0) Then
            Send("<br /><br />")
            Send("<table summary=""" & Copient.PhraseLib.Lookup("term.timeday", LanguageID) & """>")
            Send("<tr>")
            Send("  <th class=""th-del"">Delete</th>")
            Send("  <th class=""th-date"">Start Time</th>")
            Send("  <th>End Time</th>")
            Send("</tr>")
            
            For slot = 0 To rst.Rows.Count - 1
              StartTime = MyCommon.NZ(rst.Rows(slot).Item("StartTime"), "00:00")
              StartHour = Left(StartTime, 2)
              StartMinute = Right(StartTime, 2)
              If (StartHour > 12) Then
                StartHour -= 12
                StartMeridiem = "PM"
              ElseIf (StartHour = 12) Then
                StartMeridiem = "PM"
              ElseIf (StartHour = 0) Then
                StartHour = 12
                StartMeridiem = "AM"
              Else
                StartMeridiem = "AM"
              End If
              
              EndTime = MyCommon.NZ(rst.Rows(slot).Item("EndTime"), "00:00")
              EndHour = Left(EndTime, 2)
              EndMinute = Right(EndTime, 2)
              
              If (EndHour > 12) Then
                EndHour -= 12
                EndMeridiem = "PM"
              ElseIf (EndHour = 12) Then
                EndMeridiem = "PM"
              ElseIf (EndHour = 0) Then
                EndHour = 12
                EndMeridiem = "AM"
              Else
                EndMeridiem = "AM"
              End If
              IncentiveTODID = MyCommon.NZ(rst.Rows(slot).Item("IncentiveTODID"), -1)
              
              Send("<tr>")
              Send("  <td><input type=""button"" class=""ex"" id=""btnDeleteSlot"" name=""btnDeleteSlot"" onclick=""deleteTimeSlot(" & IncentiveTODID & ");"" value=""X"" " & DisabledAttribute & "  /></td>")
              Send("  <td>")
              Send("    <span id=""startHourSpan" & slot & """>" & StartHour.ToString.PadLeft(2, "0") & "</span>:" & StartMinute.ToString.PadLeft(2, "0") & "&nbsp;<span id=""startMeridiemSpan" & slot & """>" & StartMeridiem & "</span>")
              Send("  </td>")
              Send("  <td>")
              Send("    <span id=""endHourSpan" & slot & """>" & EndHour.ToString.PadLeft(2, "0") & "</span>:" & EndMinute.ToString.PadLeft(2, "0") & "&nbsp;<span id=""endMeridiemSpan" & slot & """>" & EndMeridiem & "</span>")
              Send("  </td>")
              Send("</tr>")
            Next
            Send("</table>")
          End If
          
          Send("</div>")
        %>
        <br />
      </div>
    </div>
    
    <div id="gutter">
    </div>
    
    <div id="column2">
    </div>
    
  </div>
</form>

<script type="text/javascript">
<% If (CloseAfterSave) Then %>
    window.close();
<% End If %>
<% If (TimeEntryType = 2) Then %>
    handleClockType(2);
<% End If %>
</script>
<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd("mainform", "startHour")
  MyCommon = Nothing
  Logix = Nothing
%>
