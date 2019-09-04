<%@ Page Language="vb" Debug="true" CodeFile="ue-cb.vb" Inherits="UECB" %>

<%@ Import Namespace="CMS.AMS.Models" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
    ' *****************************************************************************
    ' * FILENAME: UEoffer-con-point.aspx 
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
    Dim row As DataRow
    Dim rst2 As DataTable
    Dim OfferID As Long
    Dim Name As String = ""
    Dim isTemplate As Boolean
    Dim FromTemplate As Boolean
    Dim Disallow_Edit As Boolean = True
    Dim DisabledAttribute As String = ""
    Dim roid As Integer
    Dim Ids() As String
    Dim i As Integer
    Dim historyString As String = ""
    Dim CloseAfterSave As Boolean = False
    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False
    Dim RequirePP As Boolean = False
    Dim HasRequiredPP As Boolean = False
    Dim EngineID As Integer = 2
    Dim BannersEnabled As Boolean = True
    Dim QtyForIncentive As String = ""
    Dim TierLevels As Integer = 1
    Dim t As Integer = 1
    Dim TierQty As Decimal
    Dim ValidTier As Boolean = False
    Dim TierDT As DataTable
    Dim ProgramID As Integer = 0
    Dim PointsID As Integer = 0
    Dim ValildTier As Boolean = False
    Dim NegTiers As Boolean = False
    Dim NumArray(99) As Decimal
    Dim NonNumericInput As Boolean = False
    Dim NegativeInput As Boolean = False
    Dim ReloadFromBrowser As Boolean = False
    Dim ValidProximityMessage As Boolean = False
    Dim IsAnyCustomer As Boolean = False
    Dim iRadioValue As Int16 = 3
    Dim granttypeid As Int16 = 2
    Dim bEnablePointCondition As Boolean = IIf(MyCommon.Fetch_UE_SystemOption(181) = "1", True, False)
    Dim SupportGlobalAndTieredConditions As Integer = MyCommon.Fetch_UE_SystemOption(197)
    Dim UseSameTierValue As Integer = 0

    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If

    Response.Expires = 0
    MyCommon.AppName = "UEoffer-con-point.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)

    BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")

    OfferID = Request.QueryString("OfferID")

    'This would redirect to summary page if offer is awaiting deployment or if offer is awaiting recommendations.
    CheckIfValidOffer(MyCommon, OfferID)

    PointsID = MyCommon.Extract_Val(Request.QueryString("IncentivePointsID"))

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

    MyCommon.QueryStr = "select RewardOptionID, TierLevels from CPE_RewardOptions with (NoLock) where IncentiveID=" & OfferID & " and TouchResponse=0 and Deleted=0;"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
        roid = rst.Rows(0).Item("RewardOptionID")
        TierLevels = rst.Rows(0).Item("TierLevels")
    End If

    MyCommon.QueryStr = "select ProgramID, IncentivePointsID from CPE_IncentivePointsGroups with (NoLock) where IncentivePointsID=" & PointsID & " and Deleted=0;"
    rst = MyCommon.LRT_Select
    If MyCommon.Extract_Val(Request.QueryString("selGroups")) <> 0 Then
        ProgramID = MyCommon.Extract_Val(Request.QueryString("selGroups"))
    ElseIf rst.Rows.Count > 0 Then
        ProgramID = MyCommon.NZ(rst.Rows(0).Item("ProgramID"), 0)
        PointsID = MyCommon.NZ(rst.Rows(0).Item("IncentivePointsID"), 0)
    End If

    'Get UseSameTierValue
    If TierLevels > 1 And SupportGlobalAndTieredConditions = 1 Then
        MyCommon.QueryStr = "select IPGT.TierLevel, IPGT.Quantity from CPE_IncentivePointsGroupTiers As IPGT with (NoLock) left join CPE_IncentivePointsGroups as IPG with (NoLock) " & _
                            "on IPGT.IncentivePointsID=IPG.IncentivePointsID where IPG.IncentivePointsID=" & PointsID & " and IPG.Deleted=0;"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
            Dim lastQuantity As Integer = 0
            For Each row In rst.Rows
                If (MyCommon.NZ(row.Item("TierLevel"), 0) > 1) and MyCommon.NZ(row.Item("Quantity"), 0) <> lastQuantity Then
                    UseSameTierValue = 0
                    Exit For
                Else
                    lastQuantity = MyCommon.NZ(row.Item("Quantity"), 0)
                End If
            Next
            If MyCommon.NZ(row.Item("TierLevel"), 0) = TierLevels Then
                UseSameTierValue = 1
            End If
        End If
    Else
        UseSameTierValue = 0
    End If

    ' see if someone is saving
    If (Request.QueryString("save") <> "" And roid > 0) Then
        For t = 1 To TierLevels
            If IsNumeric(Request.QueryString("t" & t & "_QtyForIncentive")) Then
                If MyCommon.Extract_Val(Request.QueryString("t" & t & "_QtyForIncentive")) < 0 Then
                    NegativeInput = True
                Else
                    NumArray(t - 1) = MyCommon.Extract_Val(Request.QueryString("t" & t & "_QtyForIncentive"))
                End If
            Else
                NonNumericInput = True
            End If
        Next
        If NumArray(0) < 0 Then
            NegTiers = True
        End If
        ValidTier = ValidateTier(TierLevels, NumArray,MyCommon)
        If (TierLevels = 1 AndAlso Request.QueryString("t1_QtyForIncentive") = "") OrElse (NonNumericInput) OrElse (NegativeInput) Then
            ValidTier = False
        End If

        Dim TierValues As String = ""
        For t = 1 To TierLevels
            TierValues &= MyCommon.Extract_Val(Request.QueryString("t" & t & "_QtyForIncentive")) & ","
        Next
        If ValidProximityMessagePointCondtionExist(MyCommon, roid, TierLevels, TierValues, PointsID)
            ValidProximityMessage = True
        End If

        If ValidTier AndAlso ValidProximityMessage Then
            'Delete tiers records
            MyCommon.QueryStr = "delete from CPE_IncentivePointsGroupTiers where IncentivePointsID=" & PointsID & ";"
            MyCommon.LRT_Execute()

            ' check to see if a points condition is required by the template, if applicable
            MyCommon.QueryStr = "select ProgramID from CPE_IncentivePointsGroups with (NoLock) where IncentivePointsID=" & PointsID & _
                                " and RequiredFromTemplate=1 and Deleted=0;"
            rst = MyCommon.LRT_Select
            HasRequiredPP = (rst.Rows.Count > 0)
            'Check for points condition to update
            MyCommon.QueryStr = "select IncentivePointsID from CPE_IncentivePointsGroups where Deleted=0 and IncentivePointsID=" & PointsID & ";"
            rst = MyCommon.LRT_Select()
            Dim QtyString As String = ""
            If TierLevels = 1 Then
                QtyString = " requires " & Request.QueryString("t1_QtyForIncentive")
            Else
                For t = 1 To TierLevels
                    QtyString += " TierLevel " & t & " requires " & Request.QueryString("t" & t & "_QtyForIncentive") & ";"
                Next
            End If
            historyString = "Altered points condition group: Point Program " & Request.QueryString("selGroups") & QtyString
            If rst.Rows.Count > 0 Then
                If (bEnablePointCondition) Then
                    MyCommon.QueryStr = "update CPE_IncentivePointsGroups set ProgramID=" & ProgramID & ", QtyForIncentive=" & MyCommon.NZ(Request.QueryString("t1_QtyForIncentive"), 0) & ", " & _
                                             "PointsAlgorithmTypeID =" & MyCommon.NZ(Request.QueryString("radioValue"), 3) & ", "
                    If (TierLevels = 1) Then MyCommon.QueryStr &= "RewardGrantConditionID =" & MyCommon.NZ(Request.QueryString("granted"), 2) & ", "
                    MyCommon.QueryStr &= "Deleted=0, LastUpdate=getdate(), RequiredFromTemplate=" & IIf(HasRequiredPP, "1", "0") & _
                                          " where IncentivePointsID=" & PointsID & ";"
                Else
                    MyCommon.QueryStr = "update CPE_IncentivePointsGroups set ProgramID=" & ProgramID & ", QtyForIncentive=" & MyCommon.NZ(Request.QueryString("t1_QtyForIncentive"), 0) & ", " & _
                                "Deleted=0, LastUpdate=getdate(), RequiredFromTemplate=" & IIf(HasRequiredPP, "1", "0") & " where IncentivePointsID=" & PointsID & ";"
                End If
                MyCommon.LRT_Execute()
            Else
                If (Request.QueryString("selGroups") <> "") Then
                    ' ok we need to do some work to set the limit values if there are any otherwise just set to 0
                    ' in theory there should be one set of limit values for each selected groups and possibly an accumulation infos
                    If (bEnablePointCondition) Then
                        If (TierLevels = 1) Then
                            MyCommon.QueryStr = "insert into CPE_IncentivePointsGroups (RewardOptionID, ProgramID, QtyForIncentive, Deleted, LastUpdate, RequiredFromTemplate, PointsAlgorithmTypeID, RewardGrantConditionID) " & _
                                            " values(" & roid & ", " & ProgramID & ", " & IIf(Request.QueryString("t1_QtyForIncentive") = "", 0, Request.QueryString("t1_QtyForIncentive")) & ", 0, getdate(), " & IIf(HasRequiredPP, "1", "0") & _
                                            "," & MyCommon.NZ(Request.QueryString("radioValue"), 3) & "," & MyCommon.NZ(Request.QueryString("granted"), 2) & ");"
                        Else
                            MyCommon.QueryStr = "insert into CPE_IncentivePointsGroups (RewardOptionID, ProgramID, QtyForIncentive, Deleted, LastUpdate, RequiredFromTemplate, PointsAlgorithmTypeID) " & _
                                            " values(" & roid & ", " & ProgramID & ", " & IIf(Request.QueryString("t1_QtyForIncentive") = "", 0, Request.QueryString("t1_QtyForIncentive")) & ", 0, getdate(), " & IIf(HasRequiredPP, "1", "0") & _
                                            "," & MyCommon.NZ(Request.QueryString("radioValue"), 3) & ");"
                        End If

                    Else
                        MyCommon.QueryStr = "insert into CPE_IncentivePointsGroups (RewardOptionID, ProgramID, QtyForIncentive, Deleted, LastUpdate, RequiredFromTemplate) " & _
                        " values(" & roid & ", " & ProgramID & ", " & IIf(Request.QueryString("t1_QtyForIncentive") = "", 0, Request.QueryString("t1_QtyForIncentive")) & ", 0, getdate(), " & IIf(HasRequiredPP, "1", "0") & ");"
                    End If
                    MyCommon.LRT_Execute()

                    MyCommon.QueryStr = "select IncentivePointsID from CPE_IncentivePointsGroups where RewardOptionID=" & roid & " and Deleted=0 order by IncentivePointsID desc;"
                    rst = MyCommon.LRT_Select()
                    If rst.Rows.Count > 0 Then
                        PointsID = rst.Rows(0).Item("IncentivePointsID")
                    End If
                End If
            End If

            For t = 1 To TierLevels
                TierQty = MyCommon.Extract_Val(Request.QueryString("t" & t & "_QtyForIncentive"))
                MyCommon.QueryStr = "insert into CPE_IncentivePointsGroupTiers (RewardOptionID,IncentivePointsID,TierLevel,Quantity) " & _
                                    "values (" & roid & "," & PointsID & "," & t & "," & TierQty & ");"
                MyCommon.LRT_Execute()
            Next

            MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
            MyCommon.LRT_Execute()
            If (infoMessage = "") Then
                CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
            Else
                CloseAfterSave = False
            End If
            ResetOfferApprovalStatus(OfferID)
            MyCommon.Activity_Log(3, OfferID, AdminUserID, historyString)
        Else
            If NonNumericInput Then
                infoMessage = Copient.PhraseLib.Lookup("error.invalidvalue", LanguageID)
                ReloadFromBrowser = True
                'PointsID = 0
            ElseIf NegativeInput Then
                infoMessage = Copient.PhraseLib.Lookup("condition.badvalue", LanguageID)
                ReloadFromBrowser = True
                'PointsID = 0
            ElseIf Not ValidProximityMessage Then
                infoMessage = Copient.PhraseLib.Lookup("condition.affectspmrpointvalue", LanguageID)
                ReloadFromBrowser = True
            Else
                If NegTiers Then
                    infoMessage = Copient.PhraseLib.Lookup("error.tiervalues-negative", LanguageID)
                    ReloadFromBrowser = True
                    'PointsID = 0
                Else
                    infoMessage = Copient.PhraseLib.Lookup("error.tiervalues", LanguageID)
                    ReloadFromBrowser = True
                    'PointsID = 0
                End If
            End If
        End If
    End If

    ' dig the offer info out of the database
    ' no one clicked anything
    MyCommon.QueryStr = "Select IncentiveID,IsTemplate,ClientOfferID,IncentiveName as Name,CPE.Description,PromoClassID,CRMEngineID,Priority," & _
                        "StartDate,EndDate,EveryDOW,EligibilityStartDate,EligibilityEndDate,TestingStartDate,TestingEndDate,P1DistQtyLimit,P1DistTimeType,P1DistPeriod," & _
                        "P2DistQtyLimit,P2DistTimeType,P2DistPeriod,P3DistQtyLimit,P3DistTimeType,P3DistPeriod,EnableImpressRpt,EnableRedeemRpt,CreatedDate,CPE.LastUpdate,CPE.Deleted, CPEOARptDate, CPEOADeploySuccessDate, CPEOADeployRpt," & _
                        "CRMRestricted,StatusFlag,OC.Description as CategoryName,IsTemplate,FromTemplate from CPE_Incentives as CPE with (NoLock) left join OfferCategories as OC with (NoLock) on CPE.PromoClassID=OfferCategoryID " & _
                        "where IncentiveID=" & Request.QueryString("OfferID") & ";"
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
            MyCommon.QueryStr = "update CPE_IncentivePointsGroups with (RowLock) set DisallowEdit=" & form_Disallow_Edit & _
                                ", RequiredFromTemplate=0 where RewardOptionID=" & roid & " and Deleted=0;"
        Else
            MyCommon.QueryStr = "update CPE_IncentivePointsGroups with (RowLock) set DisallowEdit=" & form_Disallow_Edit & _
                                ", RequiredFromTemplate=" & form_Require_PP & " " & _
                                " where RewardOptionID=" & roid & " and Deleted=0;"
        End If
        MyCommon.LRT_Execute()

        ' if necessary, create an empty condition
        If (form_Require_PP = 1) Then
            MyCommon.QueryStr = "select ProgramID from CPE_IncentivePointsGroups with (NoLock) where RewardOptionID=" & roid & " and Deleted=0;"
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count = 0) Then
                MyCommon.QueryStr = "insert into CPE_IncentivePointsGroups (RewardOptionID,QtyForIncentive, Deleted,LastUpdate,RequiredFromTemplate) " & _
                                    " values(" & roid & "," & IIf(Request.QueryString("t1_QtyForIncentive") <> "", Request.QueryString("t1_QtyForIncentive"), "0") & ",0,getdate(),1);"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "select top 1 IncentivePointsID from CPE_IncentivePointsGroups where RewardOptionID=" & roid & " order by LastUpdate DESC;"
                rst2 = MyCommon.LRT_Select()
                If rst2.Rows.Count > 0 Then
                    For t = 1 To TierLevels
                        TierQty = MyCommon.Extract_Val(Request.QueryString("t" & t & "_QtyForIncentive"))
                        MyCommon.QueryStr = "insert into CPE_IncentivePointsGroupTiers (IncentivePointsID,RewardOptionID,TierLevel,Quantity) " & _
                                            "values (" & rst2.Rows(0).Item("IncentivePointsID") & "," & roid & "," & t & "," & TierQty & ");"
                        MyCommon.LRT_Execute()
                    Next
                End If
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
        MyCommon.QueryStr = "select DisallowEdit, RequiredFromTemplate from CPE_IncentivePointsGroups with (NoLock) where RewardOptionID=" & roid & " and Deleted=0;"
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

    Send_HeadBegin("term.offer", "term.pointscondition", OfferID)
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts()
%>
<script type="text/javascript">
// This is the javascript array holding the function list
// The PrintJavascriptArray ASP function can be used to print this array.
<%
 If UEOffer_Has_AnyCustomer(MyCommon, OfferID) Then
  IsAnyCustomer = true
 End If
  MyCommon.QueryStr = "Select ProgramID, ProgramName from PointsPrograms with (NoLock) where Deleted=0 and ProgramID is not null "
  If EngineID = 6 Then
    MyCommon.QueryStr &= " and CAMProgram=1"
  Else
    MyCommon.QueryStr &= " and CAMProgram=0"
  End If
  If IsAnyCustomer Then
    MyCommon.QueryStr &= " and ProgramID in (SELECT ProgramID FROM PointsProgramsPromoEngineSettings WHERE AllowAnyCustomer = 1)"
  End If
  MyCommon.QueryStr &= " order by ProgramName;"
  rst = MyCommon.LRT_Select
  If (rst.rows.count>0)
    Sendb("var functionlist = Array(")
    For Each row In rst.Rows
      Sendb("""" & MyCommon.NZ(row.item("ProgramName"), "").ToString().Replace("""", "\""") & """,")
    Next
    Sendb(""""");")
    Sendb("var vallist = Array(")
    For Each row In rst.Rows
      Sendb("""" & MyCommon.NZ(row.item("ProgramID"), 0) & """,")
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
  searchPattern = cleanSpecialChar(textObj.value);
  if (document.forms[0].functionradio[0].checked == true){
      searchPattern = "^" + searchPattern;
  }
 
  
  // Create a regular expression
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
  removeUsed(true);
}

function removeUsed(handleKeyUpAlreadyhandled) {
    if (!handleKeyUpAlreadyhandled) handleKeyUp(99999);
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
  if (selectedValue != "") {
    selectedText = selectObj[document.getElementById("functionselect").selectedIndex].text;
  }
  
  selectboxObj = document.forms[0].selected;
  selectedboxValue = document.getElementById("selected").value;
  if (selectedboxValue != ""){
    selectedboxText = selectboxObj[document.getElementById("selected").selectedIndex].text;
  }
  
  if (itemSelected == "select1") {
    if (selectedValue != "") {
      // add items to selected box
      document.getElementById('deselect1').disabled=false;
      selectboxObj[selectboxObj.length] = new Option(selectedText,selectedValue);
      if (selectboxObj.length == 1) {
        document.getElementById('select1').disabled=true;      
      }
      if (selectboxObj.length == 0) {
        // nothing in the select box so disable deselect
        document.getElementById('deselect1').disabled=true;
      }
    }
  }
  
  if (itemSelected == "deselect1") {
    if (selectedboxValue != "") {
      // remove items from selected box
      document.getElementById("selected").remove(document.getElementById("selected").selectedIndex)
      if (selectboxObj.length == 1) {
        document.getElementById('select1').disabled=true;      
      }
      if (selectboxObj.length == 0) {
        // nothing in the select box so disable deselect
        document.getElementById('deselect1').disabled=true;
        document.getElementById('select1').disabled=false;      
      }
    }
  }
  // remove items from large list that are in the other lists
  removeUsed(false);
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
  enableTiers();
  return true;
}

function validateEntry() {
  var retVal = true;
  var elemPP = document.getElementById("require_pp");
  var elem = document.getElementById("selected");   
  var qtyElem = document.getElementById("t1_QtyForIncentive");
  var elemProgram = document.getElementById("ProgramID");
  var msg = '';
  var tierLevel = 1;
  var bEnablePointsConditions='<% Sendb(IIf(MyCommon.Fetch_UE_SystemOption(181) = "1", True, False)) %>'
  
  if (elemPP == null || !elemPP.checked) {
    if (elem != null && elem.options.length == 0) {
      retVal = false;
      msg += '* <% Sendb(Copient.PhraseLib.Lookup("CPE-rew-membership.selectpoints", LanguageID)) %>'
      elem.focus();
    } else if (elem !=null && elemProgram != null) {
      elemProgram.value = elem.options[0].value;
    }
     while (qtyElem != null && tierLevel <= 4) {
      // trim the string
      var qtyVal = qtyElem.value.replace(/^\s+|\s+$/g, ''); 

      if(bEnablePointsConditions == 'True')
      {
          if (qtyVal == "" || isNaN(qtyVal) || !isInteger(qtyVal) || parseInt(qtyVal) < 0) {
            retVal = false;
            if (msg != '') { msg += '\n\r\n\r'; }
            msg += '* <% Sendb(Copient.PhraseLib.Lookup("CPEoffer-con-point.positiveinteger", LanguageID)) %>';
            qtyElem.focus();
            qtyElem.select();
          }
      }
      else
      {
          if (qtyVal == "" || isNaN(qtyVal) || !isInteger(qtyVal) || parseInt(qtyVal) == 0) {
            retVal = false;
            if (msg != '') { msg += '\n\r\n\r'; }
            msg += '* <% Sendb(Copient.PhraseLib.Lookup("CPEoffer-con-point.positiveinteger", LanguageID)) %>';
            qtyElem.focus();
            qtyElem.select();
          }
      }
      tierLevel +=1;
      qtyElem = document.getElementById("t" + tierLevel + "_QtyForIncentive");
    }
  }
  if (msg != '') {
    alert(msg);
  }
  
  return retVal;
}

function updateButtons() {
  var elemSelect1 = document.getElementById('select1');
  var elemDeselect1 = document.getElementById('deselect1');
  var elemSave = document.getElementById('save');
  var elemSelected = document.forms[0].selected;
  
  if (elemSelected != null) {
    if (elemSelected.length == 0) {
      elemSelect1.disabled = false;
      elemDeselect1.disabled = true;
      if (document.getElementById('require_pp') != null) {
        if (document.getElementById('require_pp').checked == true) {
          if(elemSave!=null) elemSave.disabled = false;
        } else {
          if(elemSave!=null) elemSave.disabled = true;
        }
      } else {
        if(elemSave!=null) elemSave.disabled = true;
      }
    } else {
      elemSelect1.disabled = true;
      elemDeselect1.disabled = false;
      if(elemSave!=null) elemSave.disabled = false;
    }
  }
  <%
   m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
    If Not IsTemplate Then
      If Not (Logix.UserRoles.EditOffer  And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And Disallow_Edit)) Then
        Send("  disableAll();")
      End If
    Else
      If Not (Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer) Then
        Send("  disableAll();")
      End If
    End If
  %>
}

function handleRequiredToggle() {
  if(document.forms[0].selected.length == 0) {
    if (document.getElementById("require_pp").checked == false) {
      document.getElementById('save').disabled=true;
    } else {
      document.getElementById('save').disabled=false;
    }
  }
  if (document.getElementById("require_pp").checked == true) {
    document.getElementById("Disallow_Edit").checked=false;
  }
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
  <input type="hidden" id="IncentivePointsID" name="IncentivePointsID" value="<% sendb(PointsID) %>" />
  <input type="hidden" id="roid" name="roid" value="<%sendb(roid) %>" />
  <input type="hidden" id="EngineID" name="EngineID" value="<% Sendb(EngineID) %>" />
  <input type="hidden" id="IsTemplate" name="IsTemplate" value="<% 
        If (IsTemplate) Then
        Sendb("IsTemplate")
        Else
        Sendb("Not")
        End If
        %>" />
  <div id="intro">
    <%If (isTemplate) Then
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.pointscondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
      Else
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.pointscondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
      End If
    %>
    <div id="controls">
      <% If (isTemplate) Then%>
      <span class="temp">
        <input type="checkbox" class="tempcheck" id="Disallow_Edit" name="Disallow_Edit"<% if(disallow_edit)then sendb(" checked=""checked""") %> />
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
      <div class="box" id="selector">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.pointsprogram", LanguageID))%>
          </span>
          <% If (isTemplate) Then%>
          <span class="tempRequire">
            <input type="checkbox" class="tempcheck" id="require_pp" name="require_pp" onclick="handleRequiredToggle();"<% if(requirepp)then sendb(" checked=""checked""") %> />
            <label for="require_pp">
              <% Sendb(Copient.PhraseLib.Lookup("term.required", LanguageID))%>
            </label>
          </span>
          <% ElseIf (FromTemplate And RequirePP) Then%>
          <span class="tempRequire">*
            <%Sendb(Copient.PhraseLib.Lookup("term.required", LanguageID))%>
          </span>
          <% End If%>
        </h2>
        <input type="radio" id="functionradio1" name="functionradio" <% if(MyCommon.Fetch_SystemOption(175)= "1" ) then sendb(" checked=""checked""") %> <% sendb(disabledattribute) %> /><label for="functionradio1"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
        <input type="radio" id="functionradio2" name="functionradio" <% if(MyCommon.Fetch_SystemOption(175)= "2" ) then sendb(" checked=""checked""") %> <% sendb(disabledattribute) %> /><label for="functionradio2"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
        <input type="text" class="medium" id="functioninput" name="functioninput" maxlength="100" onkeyup="handleKeyUp(200);" value=""<% sendb(disabledattribute) %> /><br />
        <select class="longer" id="functionselect" name="functionselect" size="16"<% sendb(disabledattribute) %>>
          <%
            If UEOffer_Has_AnyCustomer(MyCommon, OfferID) Then
              IsAnyCustomer = True
            End If
            MyCommon.QueryStr = "Select ProgramID, ProgramName from PointsPrograms with (NoLock) where Deleted=0 and ProgramID is not null "
            If EngineID = 6 Then
              MyCommon.QueryStr &= " and CAMProgram=1"
            Else
              MyCommon.QueryStr &= " and CAMProgram=0"
            End If
            If IsAnyCustomer Then
              MyCommon.QueryStr &= " and ProgramID in (SELECT ProgramID FROM PointsProgramsPromoEngineSettings WHERE AllowAnyCustomer = 1)"
            End If
            MyCommon.QueryStr &= " order by ProgramName;"
            rst = MyCommon.LRT_Select
            For Each row In rst.Rows
              Send("<option value=" & MyCommon.NZ(row.Item("ProgramID"), 0) & ">" & MyCommon.NZ(row.Item("ProgramName"), "") & "</option>")
            Next
          %>
        </select>
        <br />
        <br class="half" />
        <input type="button" class="regular select" id="select1" name="select1" value="&#9660; <% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID)) %>" onclick="handleSelectClick('select1');"<% sendb(disabledattribute) %> />&nbsp;
        <input type="button" class="regular deselect" id="deselect1" name="deselect1" value="<% Sendb(Copient.PhraseLib.Lookup("term.deselect", LanguageID)) %> &#9650;" disabled="disabled" onclick="handleSelectClick('deselect1');"<% sendb(disabledattribute) %> /><br />
        <br class="half" />
        <select class="longer" id="selected" name="selected" size="2"<% sendb(disabledattribute) %>>
          <%
            If PointsID <> 0 AndAlso Not ReloadFromBrowser Then
              MyCommon.QueryStr = "Select IPG.ProgramID,ProgramName from CPE_IncentivePointsGroups as IPG with (NoLock) left join PointsPrograms as PP with (NoLock) on PP.ProgramID=IPG.ProgramID where IPG.deleted=0 and IPG.ProgramID is not null and IncentivePointsID=" & PointsID
              rst = MyCommon.LRT_Select
              For Each row In rst.Rows
                Send("<option value=""" & MyCommon.NZ(row.Item("ProgramID"), 0) & """>" & MyCommon.NZ(row.Item("ProgramName"), "") & "</option>")
              Next
            Else
              If Request.QueryString("selGroups") <> "" Then
                MyCommon.QueryStr = "select ProgramName from PointsPrograms where ProgramID=" & MyCommon.Extract_Val(Request.QueryString("selGroups"))
                rst = MyCommon.LRT_Select
                If rst.Rows.Count > 0 Then
                  Send("<option value=""" & MyCommon.Extract_Val(Request.QueryString("selGroups")) & """>" & MyCommon.NZ(rst.Rows(0).Item("ProgramName"), "") & "</option>")
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
        <%
            Send("         <div class=""box"" id=""value"">")
            Send("         <h2>")
            Send("         <span>")
            Sendb(Copient.PhraseLib.Lookup("term.value", LanguageID))
            Send("         </span>")
            Send("         </h2>")
            Send("        <label for=""t1_QtyForIncentive"">" & Copient.PhraseLib.Lookup("condition.valueneeded", LanguageID) & "</label>")
            Send(" <br /> ")
            
          If TierLevels > 1 Then
            If PointsID = 0 OrElse ReloadFromBrowser Then
              For t = 1 To TierLevels
                If (SupportGlobalAndTieredConditions = 1 And t = 1 And TierLevels > 1) Then
                  Send("  <input type=""checkbox"" name=""useSameTierValue"" id =""useSameTierValue"" style=""margin-left:0px;"" align='top' value=""1""" & IIf(UseSameTierValue = 1, " checked=""checked""", "") & " style=""margin-left:7px;"" onclick=""setSameTierValue(" & TierLevels & ")""/>")
                  Sendb("  <label for=""useThisValueForAllTiers"" align='top'>" & Copient.PhraseLib.Lookup("term.UseThisValueForAllTiers", LanguageID) & "</label><br>") 
                End If

                If Request.QueryString("t" & t & "_QtyForIncentive") <> "" Then
                  Send("        <label for=""t" & t & "_QtyForIncentive"">" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</label>")
                            Send("        <input type=""text"" class=""short"" id=""t" & t & "_QtyForIncentive"" name=""t" & t & "_QtyForIncentive"" value=""" & Request.QueryString("t" & t & "_QtyForIncentive") & """ " & IIf(UseSameTierValue = 1, " disabled=""disabled""", DisabledAttribute) & " maxlength=""6"" /><br />")
                Else
                  Send("        <label for=""t" & t & "_QtyForIncentive"">" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</label>")
                            Send("        <input type=""text"" class=""short"" id=""t" & t & "_QtyForIncentive"" name=""t" & t & "_QtyForIncentive"" value=""0"" " & DisabledAttribute & " maxlength=""6"" /><br />")
                End If
              Next
                   
                    If (bEnablePointCondition) Then
                        If (Request.QueryString("radioValue") <> "" AndAlso Int16.TryParse((Request.QueryString("radioValue")), iRadioValue)) Then iRadioValue = Int16.Parse(Request.QueryString("radioValue"))
                        If (Request.QueryString("granted") <> "" AndAlso Int16.TryParse((Request.QueryString("granted")), granttypeid)) Then granttypeid = Int16.Parse(Request.QueryString("granted"))
                        Send(" <br class=""half"" />")
                        Sendb(Copient.PhraseLib.Lookup("condition.satisfy", LanguageID))
                        Send(" <br /> ")
                        Send("  <input type=""radio"" class=""radioValue"" id=""RadioValue1"" name=""radioValue"" value=""1""" & IIf(iRadioValue = 1, " checked=""checked""", "") & "/>")
                        Send(" <label for=""RadioValue1""> " & Copient.PhraseLib.Lookup("condition.PreviousValue", LanguageID) & "</label>")
                        Send(" <br /> ")
                        Send("  <input type=""radio"" class=""radioValue"" id=""RadioValue2"" name=""radioValue"" value=""2""" & IIf(iRadioValue = 2, " checked=""checked""", "") & "/>")
                        Send(" <label for=""RadioValue2""> " & Copient.PhraseLib.Lookup("condition.CurrentValue", LanguageID) & "</label>")
                        Send(" <br /> ")
                        Send("  <input type=""radio""  class=""radioValue"" id=""RadioValue3"" name=""radioValue"" value=""3""" & IIf(iRadioValue = 3, " checked=""checked""", "") & "/>")
                        Send(" <label for=""RadioValue3""> " & Copient.PhraseLib.Lookup("condition.TotalValue", LanguageID) & "</label>")
                        Send(" <br /> ")
                        Send("  <input type=""radio""  class=""radioValue"" id=""RadioValue4"" name=""radioValue"" value=""4""" & IIf(iRadioValue = 4, " checked=""checked""", "") & "/>")
                        Send(" <label for=""RadioValue4""> " & Copient.PhraseLib.Lookup("condition.NetValue", LanguageID) & "</label>")
                        Send(" <br /> ")
                        Send("<hr class=""hidden"" />")
                        Send(" </div>")
                        'GrantType div is disbaled when Tiers > 1
                    End If
            Else
              For t = 1 To TierLevels
                If (SupportGlobalAndTieredConditions = 1 And t = 1 And TierLevels > 1) Then
                  Send("  <input type=""checkbox"" name=""useSameTierValue"" id =""useSameTierValue"" style=""margin-left:0px;"" align='top' value=""1""" & IIf(UseSameTierValue = 1, " checked=""checked""", "") & " style=""margin-left:7px;"" onclick=""setSameTierValue(" & TierLevels & ")""/>")
                  Sendb("  <label for=""useThisValueForAllTiers"" align='top'>" & Copient.PhraseLib.Lookup("term.UseThisValueForAllTiers", LanguageID) & "</label><br>") 
                End If

                MyCommon.QueryStr = "select Quantity from CPE_IncentivePointsGroupTiers with (NoLock) where RewardOptionID=" & roid & " and TierLevel=" & t & " and IncentivePointsID=" & PointsID & ";"
                TierDT = MyCommon.LRT_Select()
                If TierDT.Rows.Count > 0 Then
                  TierQty = MyCommon.NZ(TierDT.Rows(0).Item("Quantity"), 0)
                  Send("        <label for=""t" & t & "_QtyForIncentive"">" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</label>")
                  Send("        <input type=""text"" class=""short"" id=""t" & t & "_QtyForIncentive"" name=""t" & t & "_QtyForIncentive"" value=""" & TierQty & """" & IIf(UseSameTierValue = 1, " disabled=""disabled""", DisabledAttribute) & " maxlength=""9"" /><br />")
                Else
                  Send("        <label for=""t" & t & "_QtyForIncentive"">" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</label>")
                            Send("        <input type=""text"" class=""short"" id=""t" & t & "_QtyForIncentive"" name=""t" & t & "_QtyForIncentive"" value=""0"" " & DisabledAttribute & " maxlength=""6"" /><br />")
                End If
              Next
                    
                    If (bEnablePointCondition) Then
                        Send(" <br class=""half"" />")
                        MyCommon.QueryStr = "select ProgramID, IncentivePointsID,PointsAlgorithmTypeID,RewardGrantConditionID from CPE_IncentivePointsGroups with (NoLock) where IncentivePointsID=" & PointsID
                        rst = MyCommon.LRT_Select
                        For Each row In rst.Rows
                            iRadioValue = MyCommon.NZ(row.Item("PointsAlgorithmTypeID"), 3)
                            granttypeid = MyCommon.NZ(row.Item("RewardGrantConditionID"), 2)
                        Next
                        Sendb(Copient.PhraseLib.Lookup("condition.satisfy", LanguageID))
                        Send(" <br /> ")
                        Send("  <input type=""radio"" class=""radioValue"" id=""RadioValue1"" name=""radioValue"" value=""1""" & IIf(iRadioValue = 1, " checked=""checked""", "") & "/>")
                        Send(" <label for=""RadioValue1""> " & Copient.PhraseLib.Lookup("condition.PreviousValue", LanguageID) & "</label>")
                        Send(" <br /> ")
                        Send("  <input type=""radio"" class=""radioValue"" id=""RadioValue2"" name=""radioValue"" value=""2""" & IIf(iRadioValue = 2, " checked=""checked""", "") & "/>")
                        Send(" <label for=""RadioValue2""> " & Copient.PhraseLib.Lookup("condition.CurrentValue", LanguageID) & "</label>")
                        Send(" <br /> ")
                        Send("  <input type=""radio"" class=""radioValue"" id=""RadioValue3"" name=""radioValue"" value=""3""" & IIf(iRadioValue = 3, " checked=""checked""", "") & "/>")
                        Send(" <label for=""RadioValue3""> " & Copient.PhraseLib.Lookup("condition.TotalValue", LanguageID) & "</label>")
                        Send(" <br /> ")
                        Send("  <input type=""radio"" class=""radioValue"" id=""RadioValue4"" name=""radioValue"" value=""4""" & IIf(iRadioValue = 4, " checked=""checked""", "") & "/>")
                        Send(" <label for=""RadioValue4""> " & Copient.PhraseLib.Lookup("condition.NetValue", LanguageID) & "</label>")
                        Send(" <br /> ")
                        Send("<hr class=""hidden"" />")
                        Send(" </div>")
                        'GrantType div is disbaled when Tiers > 1
                    End If
            End If
          Else
            If PointsID = 0 OrElse ReloadFromBrowser Then
              If Request.QueryString("t1_QtyForIncentive") <> "" Then
                        Send("        <input type=""text"" class=""short"" id=""t1_QtyForIncentive"" name=""t1_QtyForIncentive"" value=""" & Request.QueryString("t1_QtyForIncentive") & """" & DisabledAttribute & " maxlength=""6"" /><br />")
              Else
                Send("        <input type=""text"" class=""short"" id=""t1_QtyForIncentive"" name=""t1_QtyForIncentive""" & DisabledAttribute & " maxlength=""6"" /><br />")
                    End If
                    
                    If (bEnablePointCondition) Then
                        If (Request.QueryString("radioValue") <> "" AndAlso Int16.TryParse((Request.QueryString("radioValue")), iRadioValue)) Then iRadioValue = Int16.Parse(Request.QueryString("radioValue"))
                        If (Request.QueryString("granted") <> "" AndAlso Int16.TryParse((Request.QueryString("granted")), granttypeid)) Then granttypeid = Int16.Parse(Request.QueryString("granted"))
                        Sendb(Copient.PhraseLib.Lookup("condition.satisfy", LanguageID))
                        Send(" <br /> ")
                        Send("  <input type=""radio"" class=""radioValue"" id=""RadioValue1"" name=""radioValue"" value=""1""" & IIf(iRadioValue = 1, " checked=""checked""", "") & "/>")
                        Send(" <label for=""RadioValue1""> " & Copient.PhraseLib.Lookup("condition.PreviousValue", LanguageID) & "</label>")
                        Send(" <br /> ")
                        Send("  <input type=""radio"" class=""radioValue"" id=""RadioValue2"" name=""radioValue"" value=""2""" & IIf(iRadioValue = 2, " checked=""checked""", "") & "/>")
                        Send(" <label for=""RadioValue2""> " & Copient.PhraseLib.Lookup("condition.CurrentValue", LanguageID) & "</label>")
                        Send(" <br /> ")
                        Send("  <input type=""radio"" class=""radioValue"" id=""RadioValue3"" name=""radioValue"" value=""3""" & IIf(iRadioValue = 3, " checked=""checked""", "") & "/>")
                        Send(" <label for=""RadioValue3""> " & Copient.PhraseLib.Lookup("condition.TotalValue", LanguageID) & "</label>")
                        Send(" <br /> ")
                        Send("  <input type=""radio"" class=""radioValue"" id=""RadioValue4"" name=""radioValue"" value=""4""" & IIf(iRadioValue = 4, " checked=""checked""", "") & "/>")
                        Send(" <label for=""RadioValue4""> " & Copient.PhraseLib.Lookup("condition.NetValue", LanguageID) & "</label>")
                        Send(" <br /> ")
                        Send("<hr class=""hidden"" />")
                        Send(" </div>")
                        Send(" <div class=""box"" id=""grants"" >")
                        Send("         <h2>")
                        Send("         <span>")
                        Sendb(Copient.PhraseLib.Lookup("term.rewards", LanguageID))
                        Send("         </span>")
                        Send("         </h2>")
                        Sendb(Copient.PhraseLib.Lookup("condition.rewardsgranted", LanguageID))
                        Send(" <br /> ")
                        Send("<input type=""radio"" class=""radio"" id=""eachtime"" name=""granted"" value=""3""" & IIf(granttypeid = 3, " checked=""checked""", "") & "/>")
                        Send("<label for=""eachtime"">" & Copient.PhraseLib.Lookup("condition.eachtime", LanguageID) & "</label>")
                        Send(" <br /> ")
                        Send("<input type=""radio"" class=""radio"" id=""equalto"" name=""granted"" value=""1""" & IIf(granttypeid = 1, " checked=""checked""", "") & "/>")
                        Send("<label for=""equalto"">" & Copient.PhraseLib.Lookup("condition.equalto", LanguageID) & "</label>")
                        Send(" <br /> ")
                        Send("<input type=""radio"" class=""radio"" id=""greaterthan"" name=""granted"" value=""2""" & IIf(granttypeid = 2, " checked=""checked""", "") & "/>")
                        Send("<label for=""greaterthan"">" & Copient.PhraseLib.Lookup("condition.greaterthan", LanguageID) & "</label>")
                        Send(" <br /> ")
                        Send(" </div>")
                    End If
            Else
              MyCommon.QueryStr = "Select QtyForIncentive from CPE_IncentivePointsGroups with (NoLock) where deleted=0 and IncentivePointsID=" & PointsID
              rst = MyCommon.LRT_Select
              If (rst.Rows.Count > 0) Then
                QtyForIncentive = rst.Rows(0).Item("QtyForIncentive").ToString
              Else
                QtyForIncentive = ""
              End If
              Send("        <input type=""text"" class=""short"" id=""t1_QtyForIncentive"" name=""t1_QtyForIncentive"" value=""" & QtyForIncentive & """" & DisabledAttribute & " maxlength=""7"" /><br />")
                    
                    If (bEnablePointCondition) Then
                        Send(" <br class=""half"" />")
                             
                        MyCommon.QueryStr = "select ProgramID, IncentivePointsID,PointsAlgorithmTypeID,RewardGrantConditionID from CPE_IncentivePointsGroups with (NoLock) where IncentivePointsID=" & PointsID
                        rst = MyCommon.LRT_Select
                        For Each row In rst.Rows
                            iRadioValue = MyCommon.NZ(row.Item("PointsAlgorithmTypeID"), 3)
                            granttypeid = MyCommon.NZ(row.Item("RewardGrantConditionID"), 2)
                        Next
                        Sendb(Copient.PhraseLib.Lookup("condition.satisfy", LanguageID))
                        Send(" <br /> ")
                        Send("  <input type=""radio"" class=""radioValue"" id=""RadioValue1"" name=""radioValue"" value=""1""" & IIf(iRadioValue = 1, " checked=""checked""", "") & "/>")
                        Send(" <label for=""RadioValue1""> " & Copient.PhraseLib.Lookup("condition.PreviousValue", LanguageID) & "</label>")
                        Send(" <br /> ")
                        Send("  <input  type=""radio""  class=""radioValue"" id=""RadioValue2"" name=""radioValue"" value=""2""" & IIf(iRadioValue = 2, " checked=""checked""", "") & "/>")
                        Send(" <label for=""RadioValue2""> " & Copient.PhraseLib.Lookup("condition.CurrentValue", LanguageID) & "</label>")
                        Send(" <br /> ")
                        Send("  <input type=""radio"" class=""radioValue"" id=""RadioValue3"" name=""radioValue"" value=""3""" & IIf(iRadioValue = 3, " checked=""checked""", "") & "/>")
                        Send(" <label for=""RadioValue3""> " & Copient.PhraseLib.Lookup("condition.TotalValue", LanguageID) & "</label>")
                        Send(" <br /> ")
                        Send("  <input type=""radio"" class=""radioValue"" id=""RadioValue4"" name=""radioValue"" value=""4""" & IIf(iRadioValue = 4, " checked=""checked""", "") & "/>")
                        Send(" <label for=""RadioValue4""> " & Copient.PhraseLib.Lookup("condition.NetValue", LanguageID) & "</label>")
                        Send(" <br /> ")
                        Send("<hr class=""hidden"" />")
                        Send(" </div>")
                        Send(" <div class=""box"" id=""grants"" >")
                        Send("         <h2>")
                        Send("         <span>")
                        Sendb(Copient.PhraseLib.Lookup("term.rewards", LanguageID))
                        Send("         </span>")
                        Send("         </h2>")
                        Sendb(Copient.PhraseLib.Lookup("condition.rewardsgranted", LanguageID))
                        Send(" <br /> ")
                        Send("<input type=""radio"" class=""radio"" id=""eachtime"" name=""granted"" value=""3""" & IIf(granttypeid = 3, " checked=""checked""", "") & "/>")
                        Send("<label for=""eachtime"">" & Copient.PhraseLib.Lookup("condition.eachtime", LanguageID) & "</label>")
                        Send(" <br /> ")
                        Send("<input type=""radio"" class=""radio"" id=""equalto"" name=""granted"" value=""1""" & IIf(granttypeid = 1, " checked=""checked""", "") & "/>")
                        Send("<label for=""equalto"">" & Copient.PhraseLib.Lookup("condition.equalto", LanguageID) & "</label>")
                        Send(" <br /> ")
                        Send("<input type=""radio"" class=""radio"" id=""greaterthan"" name=""granted"" value=""2""" & IIf(granttypeid = 2, " checked=""checked""", "") & "/>")
                        Send("<label for=""greaterthan"">" & Copient.PhraseLib.Lookup("condition.greaterthan", LanguageID) & "</label>")
                        Send(" <br /> ")
                        Send(" </div>")
                    End If
            End If
          End If
        %>
        <br class="half" />
        <hr class="hidden" />
      </div>
    </div>
  </div>
</form>

<script runat="server">
  Function ValidateTier(ByVal TierLevels As Integer, ByVal NumArray() As Decimal,ByRef MyCommon As Copient.CommonInc)
    Dim ValidTier As Boolean = False
    Dim AllNeg As Boolean = False
    Dim Cont As Boolean = False
    Dim t1, t2 As Decimal
    Dim t As Integer = 0
    Dim bEnablePointCondition As Boolean = IIf(MyCommon.Fetch_UE_SystemOption(181) = "1", True, False)
  
        
    'Tier level validation code
    If TierLevels > 1 Then
      ValidTier = True
      'Are the tiers negative
      For t = 1 To TierLevels
        t1 = NumArray(t - 1)
        If bEnablePointCondition Then
            If t1 < 0 Then
              AllNeg = True
            Else
              AllNeg = False
              Exit For
            End If       
        Else
            If t1 = 0 Then
              ValidTier = False
              Exit For
            ElseIf t1 < 0 Then
              AllNeg = True
            Else
              AllNeg = False
              Exit For
            End If   
        End If
      Next
      If ValidTier Then
        If AllNeg Then
          For t = 1 To TierLevels - 1
            t2 = NumArray(t)
            t1 = NumArray(t - 1)
            Send(t2 & "<" & t1)
            If t2 <= t1 Then
              ValidTier = True
            Else
              ValidTier = False
              Exit For
            End If
          Next
        Else
          For t = 1 To TierLevels - 1
            t2 = NumArray(t)
            t1 = NumArray(t - 1)
            If t2 >= t1 Then
              ValidTier = True
            Else
              ValidTier = False
              Exit For
            End If
          Next
        End If
      End If
    Else
      ValidTier = True
    End If
    
    Return ValidTier
  End Function
    
  Function ValidProximityMessagePointCondtionExist(ByRef Common As Copient.CommonInc, ByVal ROID As Integer, ByVal TierLevels As Integer, ByVal NewValues As String, ByVal IncentivePointGroupId As Integer) As Boolean
    Dim validPMRPointCondition As Boolean = True
    Dim dt As Data.DataTable
    Dim dpg As DataTable
    Dim thresholdType As Integer
    Dim count As Integer = 0
                
    Common.QueryStr = "Select PM.ThresholdTypeID from ProximityMessage as PM " & _
                " inner join CPE_Deliverables as CPED " & _
                " on CPED.OutputID = PM.ID where CPED.DeliverableTypeID = 14 and CPED.Deleted = 0 and CPED.RewardOptionID = " & ROID
    dt = Common.LRT_Select
    
    For Each row As DataRow In dt.Rows            
        thresholdType = Common.NZ(dt.Rows(count).Item("ThresholdTypeID"), 0)
        
        If thresholdType = CPEUnitTypes.Points Then                'Validate for PMR with Point condtions only
              Common.QueryStr = "select IncentivePointsID from CPE_IncentivePointsGroups where Deleted=0 And RewardOptionID= " & ROID
              dpg = Common.LRT_Select

              If dpg.Rows.Count > 0 And IncentivePointGroupId=0 Then 'New one is being added
                validPMRPointCondition = False
              Else If IncentivePointGroupId <> 0   'Existing one being edited
                If Not ValidPMRTierValues(Common, NewValues, ROID, TierLevels)
                    validPMRPointCondition = False
                End If
              End If
        End If
        count += 1
    Next
    Return validPMRPointCondition
  End Function
    
  Function ValidPMRTierValues(Common As Copient.CommonInc, NewValues As String, ByVal ROID As Integer, ByVal TierLevels As Integer) As Boolean
    Dim Values As String() = NewValues.Split(New Char() {","c})
    Dim dst As DataTable
    Dim dt As DataTable
    Dim validTierValues As Boolean = True        
    Dim tier0Value As Decimal
        
    Common.QueryStr = "select PM.ThresholdTypeID, PMT.TriggerValue from ProximityMessageTier as PMT " & _
        "inner join ProximityMessage as PM " & _
        "on PM.ID = PMT.ProximityMessageId " & _
        "inner join CPE_Deliverables as CPED " & _
        "on CPED.OutputID = PM.ID where CPED.DeliverableTypeID = 14 and CPED.Deleted = 0 and CPED.RewardOptionID = " & ROID
    dt = Common.LRT_Select
    tier0Value = Common.NZ(dt.Rows(0).Item("TriggerValue"), 0)
    Common.QueryStr = "select IPGT.Quantity from CPE_IncentivePointsGroupTiers as IPGT " & _
                        "where IPGT.RewardOptionID = " & ROID

    dst = Common.LRT_Select
    If (dst.Rows.Count = TierLevels AndAlso TierLevels > 1) Then
        If Decimal.Parse(Values(0)) <= tier0Value Then
        validTierValues = False
        End If
                    
        For i As Integer = 1 To TierLevels - 1 Step 1
        Dim newDiff As Integer = Decimal.Parse(Values(i)) - Decimal.Parse(Values(i - 1))
        If (newDiff < Decimal.Parse(Common.NZ(dt.Rows(i).Item("TriggerValue"), 0))) Then
            validTierValues = False
        End If
        Next
    ElseIf (TierLevels = 1) Then
        If (Decimal.Parse(Values(0)) < tier0Value) Then
        validTierValues = False
        End If
    End If
    Return validTierValues        
    End Function	

</script>

<script type="text/javascript">
    function setSameTierValue(tierLevels){
      var box = document.getElementById("useSameTierValue");
      var text;
      if(box.checked){
        for (i=1; i < (tierLevels + 1); i++){
          text = "t" + i.toString() + "_QtyForIncentive";
          //alert(document.getElementById("tier1_l1discountamt").value.toString());
          document.getElementById(text).value = document.getElementById("t1_QtyForIncentive").value;
          document.getElementById(text).setAttribute('disabled', 'disabled');
        } 
      }
      else{
        for (i=1; i < (tierLevels + 1); i++){
          text = "t" + i.toString() + "_QtyForIncentive";
          document.getElementById(text).disabled = false;
        } 
      }
    }

     
    function enableTiers(){
      var t = 1
      qtyElem = document.getElementById("t" + t + "_QtyForIncentive");
  
      while (qtyElem != null) 
      {
	    qtyElem.disabled = false
        t++;
        qtyElem = document.getElementById("t" + t + "_QtyForIncentive");
      }
    }
 
<% If (CloseAfterSave) Then %>
    window.close();
<% End If %>
removeUsed(false);
updateButtons();
</script>
<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd("mainform", "functioninput")
  MyCommon = Nothing
  Logix = Nothing
%>
