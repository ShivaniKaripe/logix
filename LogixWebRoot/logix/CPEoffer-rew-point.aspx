<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: CPEoffer-rew-point.aspx 
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
  Dim Logix As New Copient.LogixInc
  Dim Localization As Copient.Localization
  Dim rst As DataTable
  Dim rst2 As DataTable
  Dim dt As DataTable
  Dim row As DataRow
  Dim OfferID As Long
  Dim RewardID As Long
  Dim ProgramID As Long
  Dim DeliverableID As Long
  Dim DPPKID As Long
  Dim Phase As Integer
  Dim Quantity As Long = 0
  Dim QuantityDisplay As String = ""
  Dim TierQuantity As Long = 0
  Dim MaxAdjustment As Long = 0
  Dim MaxAdjustmentDisplay As String = ""
  Dim MaxAdjustmentEnabled As Boolean = False
  Dim Name As String = ""
  Dim bError As Boolean = False
  Dim TouchPoint As Integer = 0
  Dim TpROID As Integer = 0
  Dim CreateROID As Integer = 0
  Dim CloseAfterSave As Boolean = False
  Dim Disallow_Edit As Boolean = True
  Dim FromTemplate As Boolean = False
  Dim IsTemplate As Boolean = False
  Dim IsTemplateVal As String = "Not"
  Dim DisabledAttribute As String = ""
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim BannersEnabled As Boolean = True
  Dim EngineID As Integer = 2
  Dim ScorecardAtProgram As Integer = 0
  Dim ScorecardDescAtProgram As String = ""
  Dim ScorecardBoldAtProgram As Boolean = False
  Dim ScorecardID As Integer = 0
  Dim ScorecardDesc As String = ""
  Dim ScorecardBold As Boolean = False
  Dim FocusField As String = "functioninput"
  Dim EdiscountType As Integer  '1=Marsh style 2=Specified PLU  3=IBM serial integration
  Dim sQuery As String = ""
  Dim AllBanners As Boolean = False
  Dim DefaultChrgBack As Integer = -1
  Dim ChargebackDeptID As Integer = -1
  Dim GlobalDepts As String = ""
  Dim TierLevels As Integer = 1
  Dim t As Integer = 1
  Dim ValidTiers As Boolean = False
  Dim NegTiers As Boolean = False
  Dim NumArray(99) As Decimal
  Dim HasZeroValEntry As Boolean = False
  
  Dim InvalidNumericEntry As Boolean = False
  Dim MLI As New Copient.Localization.MultiLanguageRec
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "CPEoffer-rew-point.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  Localization = New Copient.Localization(MyCommon)
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  OfferID = Request.QueryString("OfferID")
  RewardID = MyCommon.Extract_Val(Request.QueryString("RewardID"))
  ProgramID = MyCommon.Extract_Val(Request.QueryString("ProgramID"))
  QuantityDisplay = MyCommon.Extract_Val(Request.QueryString("t1_quantity"))
  ChargebackDeptID = MyCommon.Extract_Val(Request.QueryString("chargeback"))
  MaxAdjustmentDisplay = Request.QueryString("maxadjustment")


  If (Request.QueryString("EngineID") <> "" Or Request.Form("EngineID") <> "") Then
    EngineID = MyCommon.Extract_Val(Request.QueryString("EngineID"))
    If EngineID = 0 Then
      EngineID = MyCommon.Extract_Val(Request.Form("EngineID"))
    End If
  Else
    MyCommon.QueryStr = "select EngineID from OfferIDs where OfferID=" & OfferID & ";"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
      EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), 0)
    End If
  End If
  
  ' Ensure that neither the quantity nor the max adjustment are fractional.
  If (QuantityDisplay IsNot Nothing AndAlso QuantityDisplay.IndexOf(".") > -1) Then
    infoMessage = Copient.PhraseLib.Lookup("CPE-rew-points.integers-only", LanguageID)
    bError = True
    FocusField = "quantity"
  End If
  If (MaxAdjustmentDisplay IsNot Nothing AndAlso MaxAdjustmentDisplay.IndexOf(".") > -1) Then
    infoMessage = Copient.PhraseLib.Lookup("CPE-rew-points.integers-only", LanguageID)
    bError = True
    FocusField = "maxadjustment"
  End If
  If Request.QueryString("MaxAdjSwitch") <> 0 AndAlso infoMessage = String.Empty Then
    If MaxAdjustmentDisplay = "0" Then
      infoMessage = Copient.PhraseLib.Lookup("error.zeronotvalid", LanguageID)
      bError = True
      FocusField = "maxadjustment"
    ElseIf Long.TryParse(MaxAdjustmentDisplay, TierQuantity) = False OrElse TierQuantity < 0 Then
      infoMessage = Copient.PhraseLib.Lookup("CPEoffer-con-tender.positivevalue", LanguageID)
      bError = True
      FocusField = "maxadjustment"
    End If
  End If
  
  If (Integer.TryParse(QuantityDisplay, Quantity)) Then
    Quantity = MyCommon.Extract_Val(QuantityDisplay)
    QuantityDisplay = Quantity.ToString()
  ElseIf (infoMessage = "") Then
    infoMessage = Copient.PhraseLib.Lookup("CPE-rew-points.enterquantity", LanguageID)
    bError = True
    FocusField = "quantity"
    Quantity = 0
  End If
  If (Integer.TryParse(MaxAdjustmentDisplay, MaxAdjustment)) Then
    MaxAdjustment = MyCommon.Extract_Val(MaxAdjustmentDisplay)
    MaxAdjustmentDisplay = MaxAdjustment.ToString()
  Else
    MaxAdjustment = 0
  End If
  
  ScorecardID = MyCommon.Extract_Val(Request.QueryString("ScorecardID"))
  ScorecardDesc = IIf(Request.QueryString("ScorecardDesc") <> "", Request.QueryString("ScorecardDesc"), "")
  ScorecardBold = IIf(Request.QueryString("ScorecardBold") <> "", True, False)
  
  If EngineID = 6 Then
    MyCommon.QueryStr = "select ScorecardID, ScorecardDesc, ScorecardBold from PointsPrograms with (NoLock) " & _
                        "where Deleted=0 and ProgramID=" & ProgramID & " and CAMProgram=1;"
  Else
    MyCommon.QueryStr = "select ScorecardID, ScorecardDesc, ScorecardBold from PointsPrograms with (NoLock) " & _
                        "where Deleted=0 and ProgramID=" & ProgramID & " and CAMProgram=0;"
  End If
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    ScorecardAtProgram = MyCommon.NZ(rst.Rows(0).Item("ScorecardID"), 0)
    ScorecardDescAtProgram = MyCommon.NZ(rst.Rows(0).Item("ScorecardDesc"), "")
    ScorecardBoldAtProgram = MyCommon.NZ(rst.Rows(0).Item("ScorecardBold"), False)
  End If
  
  DeliverableID = MyCommon.Extract_Val(Request.QueryString("DeliverableID"))
  Phase = MyCommon.Extract_Val(Request.QueryString("phase"))
  If (Phase = 0) Then Phase = MyCommon.Extract_Val(Request.Form("Phase"))
  If (Phase = 0) Then Phase = 3
  
  TouchPoint = MyCommon.Extract_Val(Request.QueryString("tp"))
  If (TouchPoint > 0) Then TpROID = MyCommon.Extract_Val(Request.QueryString("roid"))
  
  MyCommon.QueryStr = "select PKID from CPE_DeliverablePoints with (NoLock) where DeliverableID=" & DeliverableID & ";"
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    DPPKID = MyCommon.NZ(rst.Rows(0).Item("PKID"), 0)
  End If
  
  MyCommon.QueryStr = "select MaxAdjustment from CPE_DeliverablePoints with (NoLock) where PKID=" & DPPKID & ";"
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    If IsDBNull(rst.Rows(0).Item("MaxAdjustment")) Then
      MaxAdjustmentEnabled = False
    Else
      MaxAdjustmentEnabled = True
    End If
  End If
  
  If Request.QueryString("MaxAdjSwitch") <> "" Then
    If Request.QueryString("MaxAdjSwitch") = "0" Then
      MaxAdjustmentEnabled = False
    ElseIf Request.QueryString("MaxAdjSwitch") = "1" Then
      MaxAdjustmentEnabled = True
    End If
  End If
  
  MyCommon.QueryStr = "select TierLevels from CPE_RewardOptions with (NoLock) where RewardOptionID=" & RewardID & ";"
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    TierLevels = MyCommon.NZ(rst.Rows(0).Item("TierLevels"), 1)
  End If
  
  ' fetch the name
  MyCommon.QueryStr = "Select IncentiveName, IsTemplate, FromTemplate from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID
  rst = MyCommon.LRT_Select
  If (rst.Rows.Count > 0) Then
    Name = MyCommon.NZ(rst.Rows(0).Item("IncentiveName"), "")
    IsTemplate = MyCommon.NZ(rst.Rows(0).Item("IsTemplate"), False)
    FromTemplate = MyCommon.NZ(rst.Rows(0).Item("FromTemplate"), False)
  End If
  IsTemplateVal = IIf(IsTemplate, "IsTemplate", "Not")
  
  If (Request.QueryString("save") <> "" AndAlso infoMessage = "") Then
    ValidTiers = True
    For t = 1 To TierLevels
      If Not IsNumeric(Request.QueryString("t" & t & "_quantity")) Then
        ValidTiers = False
        InvalidNumericEntry = True
      End If
      NumArray(t - 1) = MyCommon.Extract_Val(Request.QueryString("t" & t & "_quantity"))
    Next
    If NumArray(0) < 0 Then NegTiers = True
    
    If ValidTiers Then
      ValidTiers = ValidateTier(TierLevels, NumArray, HasZeroValEntry)
      If TierLevels = 1 AndAlso Request.QueryString("t1_quantity") = "" Then
        ValidTiers = False
      End If
    End If

    If ValidTiers = True Then
      If (OfferID > 0 AndAlso RewardID > 0 AndAlso ProgramID > 0) Then
        CreateROID = IIf(TpROID > 0, TpROID, RewardID)
        ' Preliminary check for valid scorecard print line:
        If (ScorecardID > 0 OrElse ScorecardAtProgram > 0) Then
          If ScorecardDesc = "" Then
            MyCommon.QueryStr = "select ScorecardDesc from PointsPrograms with (NoLock) where ProgramID=" & ProgramID & ";"
            rst = MyCommon.LRT_Select
            If rst.Rows.Count > 0 Then
              If MyCommon.NZ(rst.Rows(0).Item("ScorecardDesc"), "") = "" Then
                infoMessage = Copient.PhraseLib.Lookup("ue-offer-rew-point.MissingScorecardText", LanguageID)
                bError = True
                GoTo NoSave
              End If
            End If
          End If
        End If
        If MyCommon.Extract_Val(Request.QueryString("DeliverableID")) = 0 Then
          ' Create a new reward
          If (Create_Reward(OfferID, EngineID, CreateROID, ProgramID, Quantity, Phase, ScorecardID, ScorecardDesc, ScorecardBold, MaxAdjustmentEnabled, MaxAdjustment, ChargebackDeptID, DeliverableID)) Then
            ' Insert tier values
            MyCommon.QueryStr = "select PKID from CPE_DeliverablePoints with (NoLock) where DeliverableID=" & DeliverableID & ";"
            rst = MyCommon.LRT_Select()
            If rst.Rows.Count > 0 Then
              DPPKID = MyCommon.NZ(rst.Rows(0).Item("PKID"), 0)
            End If
            MyCommon.QueryStr = "delete from CPE_DeliverablePointTiers with (RowLock) where DPPKID=" & DPPKID & ";"
            MyCommon.LRT_Execute()
            For t = 1 To TierLevels
              TierQuantity = MyCommon.Extract_Val(Request.QueryString("t" & t & "_quantity"))
              Create_RewardTiers(DPPKID, t, TierQuantity)
            Next
            infoMessage = Copient.PhraseLib.Lookup("term.ChangesWereSaved", LanguageID)
            MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Reward.createpoints", LanguageID))
          Else
            infoMessage = Copient.PhraseLib.Lookup("ueoffer-rew-points.ErrorOnSave", LanguageID)
            bError = True
          End If
        Else
          ' Update an existing reward
          If (Update_Reward(OfferID, EngineID, CreateROID, DeliverableID, ProgramID, Quantity, Phase, ScorecardID, ScorecardDesc, ScorecardBold, MaxAdjustmentEnabled, MaxAdjustment, ChargebackDeptID)) Then
            ' Delete existing tier values from the tiers table, then insert new values
            MyCommon.QueryStr = "select PKID from CPE_DeliverablePoints with (NoLock) where DeliverableID=" & DeliverableID & ";"
            rst = MyCommon.LRT_Select()
            If rst.Rows.Count > 0 Then
              DPPKID = MyCommon.NZ(rst.Rows(0).Item("PKID"), 0)
            End If
            MyCommon.QueryStr = "delete from CPE_DeliverablePointTiers with (RowLock) where DPPKID=" & DPPKID & ";"
            MyCommon.LRT_Execute()
            For t = 1 To TierLevels
              TierQuantity = MyCommon.Extract_Val(Request.QueryString("t" & t & "_quantity"))
              Create_RewardTiers(DPPKID, t, TierQuantity)
            Next
            infoMessage = Copient.PhraseLib.Lookup("term.ChangesWereSaved", LanguageID)
            MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Reward.editpoints", LanguageID))
          Else
            infoMessage = Copient.PhraseLib.Lookup("ueoffer-rew-points.ErrorOnSave", LanguageID)
            bError = True
          End If
        End If
        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
        MyCommon.LRT_Execute()
        'Update multi-language:
        'Name
        MLI.ItemID = DPPKID
        MLI.MLTableName = "CPE_DeliverablePointsTranslations"
        MLI.MLColumnName = "ScorecardDesc"
        MLI.MLIdentifierName = "DeliverablePointsID"
        MLI.StandardTableName = "CPE_DeliverablePoints"
        MLI.StandardColumnName = "ScorecardDesc"
        MLI.StandardIdentifierName = "PKID"
        MLI.StandardValue = ScorecardDesc
        MLI.InputName = "ScorecardDesc"
        MLI.InputID = "ScorecardDesc"
        Localization.SaveTranslationInputs(MyCommon, MLI, Request.QueryString, 2)
      End If
      CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
    Else
      If NegTiers Then
        infoMessage = Copient.PhraseLib.Lookup("error.tiervalues-negative", LanguageID)
        DeliverableID = 0
      ElseIf InvalidNumericEntry Then
        infoMessage = Copient.PhraseLib.Lookup("term.invalidnumericentry", LanguageID)
        DeliverableID = 0
      ElseIf HasZeroValEntry Then
        infoMessage = Copient.PhraseLib.Lookup("error.zeronotvalid", LanguageID)
        DeliverableID = 0
      Else
        infoMessage = Copient.PhraseLib.Lookup("error.tiervalues", LanguageID)
        DeliverableID = 0
      End If
      bError = True
    End If
  End If
NoSave:
  
  'Update the templates permission if necessary
  If (Request.QueryString("save") <> "" AndAlso IsTemplate) Then
    ' time to update the status bits for the templates
    Dim form_Disallow_Edit As Integer = 0
    If (Request.QueryString("Disallow_Edit") = "on") Then
      form_Disallow_Edit = 1
    End If
    MyCommon.QueryStr = "update CPE_Deliverables with (RowLock) set DisallowEdit=" & form_Disallow_Edit & " " & _
                        "where DeliverableID=" & DeliverableID & ";"
    MyCommon.LRT_Execute()
  End If
  
  If (IsTemplate Or FromTemplate) Then
    ' lets dig the permissions if its a template
    MyCommon.QueryStr = "select DisallowEdit from CPE_Deliverables with (NoLock) " & _
                        "where DeliverableID=" & DeliverableID & ";"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      Disallow_Edit = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit"), False)
    Else
      Disallow_Edit = False
    End If
  End If
  
  EdiscountType = 1
  'figure out what type of Ediscount we should be dealing with
  MyCommon.QueryStr = "select OptionValue from CPE_SystemOptions with (NoLock) where OptionID=8;"
  rst = MyCommon.LRT_Select
  If Not (rst.Rows.Count = 0) Then
    EdiscountType = MyCommon.NZ(rst.Rows(0).Item("OptionValue"), 1)
  End If
  
  If Not IsTemplate Then
    DisabledAttribute = IIf(Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Edit), "", " disabled=""disabled""")
  Else
    DisabledAttribute = IIf(Logix.UserRoles.EditTemplates, "", " disabled=""disabled""")
  End If
  
  Send_HeadBegin("term.offer", "term.pointsreward", OfferID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
%>
<script type="text/javascript" language="javascript">
// This is the javascript array holding the function list
// The PrintJavascriptArray ASP function can be used to print this array.
<%
  MyCommon.QueryStr = "Select PP.ProgramID, PP.ProgramName from PointsPrograms PP with (NoLock) where PP.Deleted=0 "
  If EngineID = 6 Then
    MyCommon.QueryStr &= "and PP.CAMProgram=1 "
  Else
    MyCommon.QueryStr &= "and PP.CAMProgram=0 " & _
                         "  AND NOT EXISTS " & _
                         "    (SELECT CDP.ProgramID " & _
                         "     FROM CPE_DeliverablePoints CDP WITH (NoLock) " & _
                         "     WHERE CDP.RewardOptionID = @RewardOptionID " & _
                         "       AND CDP.Deleted=0 " & _
                         "       AND CDP.DeliverableID <> @DeliverableID " & _
                         "       AND PP.ProgramID = CDP.ProgramID) "
    MyCommon.DBParameters.Add("@RewardOptionID", SqlDbType.BigInt).Value = RewardID
    MyCommon.DBParameters.Add("@DeliverableID", SqlDbType.Int).Value = DeliverableID
  End If
  MyCommon.QueryStr &= " ORDER BY PP.ProgramName"
  rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
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
  var selectedList;
  
  document.getElementById("functionselect").size = "16";
  
  // Set references to the form elements
  selectObj = document.forms[0].functionselect;
  textObj = document.forms[0].functioninput;
  selectedList = document.getElementById("selected");
  
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
      if (vallist[i] != "" && (selectedList.options.length < 1 || vallist[i] != selectedList.options[0].value)) {
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

// this function gets the selected value and loads the appropriate
// php reference page in the display frame
// it can be modified to perform whatever action is needed, or nothing
function handleSelectClick() {
  selectObj = document.forms[0].functionselect;
  textObj = document.forms[0].functioninput;
  selectedValue = document.getElementById("functionselect").value;
  if(selectedValue != "") {
  }
}

function addToSelect() {
  var elemSource = document.getElementById("functionselect");
  var elemDest = document.getElementById("selected");
  var selOption = null;
  var selText ="", selVal = "";
  var selIndex = -1;
  
  if (elemSource != null && elemSource.options.selectedIndex == -1) {
    alert('<% Sendb(Copient.PhraseLib.Lookup("CPE-rew-membership.selectpoints", LanguageID)) %>');
    elemSource.focus();
  } else {
    selIndex = elemSource.options.selectedIndex;
    selOption = elemSource.options[selIndex];
    selText = selOption.text;
    selVal = selOption.value;
    if (elemDest != null && elemDest.options.length > 0) {
      removeFromSelect();
    }
    elemDest.options[0] = new Option(selText, selVal);
    elemSource.options[selIndex] = null;
    handleKeyUp(99999); 
  }
  generateSCInputs(selVal);
}

function removeFromSelect() {
  var elem = document.getElementById("selected");   
  var elemList = document.getElementById("functionselect");
  
  if (elem != null && elem.options.length > 0) {
    elemList.options[elemList.options.length] = new Option(elem.options[0].text, elem.options[0].value);
    elem.options[0] = null;
    handleKeyUp(99999);
  }
  generateSCInputs(0);
}

function validateEntry() {
  var retVal = true;
  var elem = document.getElementById("selected");   
  var qtyElem = document.getElementById("t1_quantity");
  var maxadjElem = document.getElementById("maxadjustment");
  var elemProgram = document.getElementById("ProgramID");
  var saveElem = document.getElementById("save");
  var msg = '';
  var tierLevel = 1;
  
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
    if (qtyVal == "" || isNaN(qtyVal) || !isSignedInteger(qtyVal)) {
      retVal = false;
      if (msg != '') { msg += '\n\r\n\r'; }
      msg += '* <% Sendb(Copient.PhraseLib.Lookup("CPE-rew-membership.enterquantity", LanguageID)) %>';
      qtyElem.focus();
      qtyElem.select();
    }
    tierLevel +=1;
    qtyElem = document.getElementById("t" + tierLevel + "_quantity");
  }
//  Commenting out this check for now, because we want users to have this field null. -hjw
//  if (maxadjElem != null) {
//    // trim the string
//    var maxadjVal = maxadjElem.value.replace(/^\s+|\s+$/g, ''); 
//    if (maxadjVal == "" || isNaN(maxadjVal)) {
//      retVal = false;
//      if (msg != '') { msg += '\n\r\n\r'; }
//      msg += '* <% Sendb(Copient.PhraseLib.Lookup("CPE-rew-membership.enterquantity", LanguageID)) %>';
//      maxadjElem.focus();
//      maxadjElem.select();
//    }
//  }
  
  if (msg != '') {
    alert(msg);
    if (saveElem != null) {
      if (saveElem.style.visibility=='hidden') {
        saveElem.style.visibility='visible';
      }
    }
  }
  
  return retVal;
}

function selectByPointsID(pointsID) {
  var elemSelect = document.getElementById("functionselect");
  var selectLength = elemSelect.length;
  
  for(i = 0; i < selectLength; i++) {
    if (elemSelect.options[i] != null && elemSelect.options[i].value == pointsID) {
      elemSelect.options[i].selected = true;
      addToSelect();
    }
  }    
}

function handleKeyDown(e) {
  var key = e.which ? e.which : e.keyCode;
  
  if (key == 40) {
    var elemSlct = document.getElementById("functionselect");
    if (elemSlct != null) { elemSlct.focus(); }
  }
}

function handleSlctKeyDown(e) {
  var key = e.which ? e.which : e.keyCode;
  
  if (key == 13) {
    var elemSlct = document.getElementById("functionselect");
    if (elemSlct != null && elemSlct.disabled == false) { 
      addToSelect();
      clearEntry();
      var elemQty = document.getElementById("quantity");
      if (elemQty != null && elemQty.disabled == false) {
        elemQty.focus();
        elemQty.select();
      } 
    }
    e.returnValue=false;
    return false;
  }
}

function clearEntry() {
  var elemInput = document.getElementById("functioninput");
  
  if (elemInput != null) {
    elemInput.value = "";
    handleKeyUp(200);
    elemInput.focus();
  }
}

function toggleMaxAdjustment(value) {
  var elemInput = document.getElementById("maxadjustment");
  var elemHidden = document.getElementById("MaxAdjustmentEnabled");
  
  if (value == 1) {
    elemInput.disabled = '';
    elemHidden.value = '1';
  } else {
    elemInput.value = '';
    elemInput.disabled = 'disabled';
    elemHidden.value = '0';
  }
}

function toggleScorecardText() {
  if (document.getElementById("ScorecardID") != null) {
    if (document.getElementById("ScorecardID").value == 0) {
      if(document.getElementById("ScorecardDescLine") != null) { document.getElementById("ScorecardDescLine").style.display = 'none'; }
      if(document.getElementById("ScorecardDesc") != null) { document.getElementById("ScorecardDesc").value = ''; }
    } else {
      if(document.getElementById("ScorecardDescLine") != null) { document.getElementById("ScorecardDescLine").style.display = ''; }
    }
  }
}

function generateSCInputs(ProgramID) {
  xmlhttpPost('ScorecardFeeds.aspx', 'ScorecardFieldsForReward=1&EngineID=<%Sendb(EngineID)%>&OfferID=<%Sendb(OfferID)%>&DeliverableID=<%Sendb(DeliverableID)%>&ProgramID=' + ProgramID + '&ScorecardTypeID=1');
  toggleScorecardText();
}

function xmlhttpPost(strURL, qryStr) {
  var xmlHttpReq = false;
  var self = this;
  var respTxt = '';
  var i = 0;
  var scInputs = document.getElementById("scorecardinputs");
  
  if (window.XMLHttpRequest) { // Mozilla/Safari
    self.xmlHttpReq = new XMLHttpRequest();
  } else if (window.ActiveXObject) { // IE
    self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
  }
  
  strURL += "?" + qryStr;
  self.xmlHttpReq.open('POST', strURL, true);
  self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
  self.xmlHttpReq.onreadystatechange = function() {
    if (self.xmlHttpReq !=null && self.xmlHttpReq.readyState == 4) {
      respTxt = self.xmlHttpReq.responseText
      if (scInputs != null) {
        scInputs.innerHTML = respTxt;
      }
    }
  }
  self.xmlHttpReq.send(qryStr);
}
</script>
<%
  Send("<script type=""text/javascript"">")
  Send("function ChangeParentDocument() { ")
  If (EngineID = 3) Then
    Send("  opener.location = 'web-offer-rew.aspx?OfferID=" & OfferID & "'; ")
  ElseIf (EngineID = 5) Then
    Send("  opener.location = 'email-offer-rew.aspx?OfferID=" & OfferID & "'; ")
  ElseIf (EngineID = 6) Then
    Send("  opener.location = 'CAM/CAM-offer-rew.aspx?OfferID=" & OfferID & "'; ")
  Else
    Send("  opener.location = 'CPEoffer-rew.aspx?OfferID=" & OfferID & "'; ")
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
<form action="CPEoffer-rew-point.aspx" id="mainform" name="mainform" onsubmit="return validateEntry();">
  <div id="intro">
    <input type="hidden" id="OfferID" name="OfferID" value="<% Sendb(OfferID) %>" />
    <input type="hidden" id="Name" name="Name" value="<% Sendb(Name) %>" />
    <input type="hidden" id="RewardID" name="RewardID" value="<% Sendb(RewardID) %>" />
    <input type="hidden" id="DeliverableID" name="DeliverableID" value="<% Sendb(DeliverableID) %>" />
    <input type="hidden" id="Phase" name="Phase" value="<% Sendb(Phase) %>" />
    <input type="hidden" id="ProgramID" name="ProgramID" value="<% Sendb(ProgramID) %>" />
    <input type="hidden" id="roid" name="roid" value="<% Sendb(IIf(TpROID > 0, TpROID, RewardID)) %>" />
    <input type="hidden" id="id" name="tp" value="<%Sendb(TouchPoint) %>" />
    <input type="hidden" id="IsTemplate" name="IsTemplate" value="<% Sendb(IsTemplateVal)%>" />
    <input type="hidden" id="EngineID" name="EngineID" value="<% Sendb(EngineID) %>" />
    <input type="hidden" id="MaxAdjustmentEnabled" name="MaxAdjustmentEnabled" value="<% Sendb(IIf(MaxAdjustmentEnabled, 1, 0)) %>" />
    <%
      If (IsTemplate) Then
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.pointsreward", LanguageID), VbStrConv.Lowercase) & "</h1>")
      Else
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.pointsreward", LanguageID), VbStrConv.Lowercase) & "</h1>")
      End If
    %>
    <div id="controls">
      <% If (IsTemplate) Then%>
      <span class="temp">
        <input type="checkbox" class="tempcheck" id="Disallow_Edit" name="Disallow_Edit"<% if(disallow_edit)then sendb(" checked=""checked""") %> />
        <label for="Disallow_Edit">
          <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
        </label>
      </span>
      <% End If%>
      <%
      If Not IsTemplate Then
        If (Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Edit)) Then
          If DeliverableID = 0 Then
            Send_Save(" onclick=""this.style.visibility='hidden';""")
          Else
            Send_Save()
          End If
        End If
      Else
        If (Logix.UserRoles.EditTemplates) Then
          If DeliverableID = 0 Then
            Send_Save(" onclick=""this.style.visibility='hidden';""")
          Else
            Send_Save()
          End If
        End If
      End If
      
        
      %>
    </div>
  </div>
  <div id="main">
    <%
      If (infoMessage <> "" And bError) Then
        Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
      ElseIf (infoMessage <> "") Then
        Send("<div id=""infobar"" class=""green-background"">" & infoMessage & "</div>")
      End If
    %>
    <div id="column1">
      <div class="box" id="selector">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.pointsprogram", LanguageID))%>
          </span>
        </h2>
        <input type="radio" id="functionradio1" name="functionradio" checked="checked"<% sendb(DisabledAttribute) %> /><label for="functionradio1"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
        <input type="radio" id="functionradio2" name="functionradio" /><label for="functionradio2"<% sendb(DisabledAttribute) %>><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
        <input type="text" class="medium" id="functioninput" name="functioninput" maxlength="100" onkeydown="handleKeyDown(event);" onkeyup="handleKeyUp(200);" value=""<% sendb(DisabledAttribute) %> /><br />
        <select class="longer" id="functionselect" name="functionselect" onkeydown="return handleSlctKeyDown(event);" onclick="handleSelectClick();" ondblclick="addToSelect();" size="16"<% sendb(DisabledAttribute) %>>
          <%
            If EngineID = 6 Then
              MyCommon.QueryStr = "select ProgramID, ProgramName from PointsPrograms with (NoLock) " & _
                                  "where Deleted=0 and CAMProgram=1 and ProgramID is not null order by ProgramName;"
            Else
              MyCommon.QueryStr = "SELECT PP.ProgramID, PP.ProgramName " & _
                                  "FROM PointsPrograms PP WITH (NoLock) " & _
                                  "WHERE PP.Deleted=0 " & _
                                  "  AND PP.CAMProgram=0 " & _
                                  "  AND PP.ProgramID IS NOT NULL " & _
                                  "  AND NOT EXISTS " & _
                                  "    (SELECT CDP.ProgramID " & _
                                  "     FROM CPE_DeliverablePoints CDP WITH (NoLock) " & _
                                  "     WHERE CDP.RewardOptionID = @RewardOptionID " & _
                                  "       AND CDP.Deleted=0 " & _
                                  "       AND CDP.DeliverableID <> @DeliverableID " & _
                                  "       AND PP.ProgramID = CDP.ProgramID) " & _
                                  "ORDER BY PP.ProgramName"
              MyCommon.DBParameters.Add("@RewardOptionID", SqlDbType.BigInt).Value = RewardID
              MyCommon.DBParameters.Add("@DeliverableID", SqlDbType.Int).Value = DeliverableID
            End If
            rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            For Each row In rst.Rows
              Send("<option value=""" & MyCommon.NZ(row.Item("ProgramID"), 0) & """>" & MyCommon.NZ(row.Item("ProgramName"), "") & "</option>")
            Next
          %>
        </select>
        <br />
        <br class="half" />
        <input type="button" class="regular" id="select" name="select" value="&#9660; <% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID))%>" onclick="addToSelect();"<% sendb(DisabledAttribute) %> />&nbsp;
        <input type="button" class="regular" id="deselect" name="deselect" value="<% Sendb(Copient.PhraseLib.Lookup("term.deselect", LanguageID))%> &#9650;" onclick="removeFromSelect();"<% sendb(DisabledAttribute) %> /><br />
        <br class="half" />
        <select class="longer" id="selected" name="selected" size="2" ondblclick="removeFromSelect();"<% sendb(DisabledAttribute) %>>
        </select>
        <hr class="hidden" />
      </div>
    </div>
    
    <div id="gutter">
    </div>
    
    <div id="column2">
      <div class="box" id="distribution">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.distribution", LanguageID))%>
          </span>
        </h2>
        <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.distribution", LanguageID))%>">
          <tr>
            <td>
              <label for="t1_quantity"><% Sendb(Copient.PhraseLib.Lookup("term.quantityawarded", LanguageID))%>:</label>
            </td>
            <td>
              <%
                If DeliverableID = 0 Then
                  For t = 1 To TierLevels
                    If Request.QueryString("t" & t & "_quantity") <> "" Then
                      If TierLevels > 1 Then
                        Send("<label for=""t" & t & "_quantity"">" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & ":</label>")
                      End If
                              Sendb("<input type=""text"" class=""shorter"" id=""t" & t & "_quantity"" name=""t" & t & "_quantity"" maxlength=""6"" value=""" & MyCommon.Extract_Val(Request.QueryString("t" & t & "_quantity")) & """/>")
                    Else
                      If TierLevels > 1 Then
                        Send("<label for=""t" & t & "_quantity"">" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & ":</label>")
                      End If
                              Sendb("<input type=""text"" class=""shorter"" id=""t" & t & "_quantity"" name=""t" & t & "_quantity"" maxlength=""6"" value=""0""/>")
                    End If
                    Send("<br />")
                  Next
                Else
                  MyCommon.QueryStr = "Select TierLevel, Quantity from CPE_DeliverablePointTiers with (NoLock) " & _
                                      "where DPPKID=" & DPPKID & " order by TierLevel;"
                  rst = MyCommon.LRT_Select
                  For t = 1 To TierLevels
                    If TierLevels > 1 Then
                      Send("<label for=""t" & t & "_quantity"">" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & ":</label>")
                    End If
                          Sendb("<input type=""text"" class=""shorter"" id=""t" & t & "_quantity"" name=""t" & t & "_quantity"" maxlength=""6""")
                    If t > rst.Rows.Count Then
                      Sendb(" value=""0""" & DisabledAttribute & " />")
                    Else
                      Sendb(" value=""" & MyCommon.NZ(rst.Rows(t - 1).Item("Quantity"), 0) & """" & DisabledAttribute & " />")
                    End If
                    Send("<br />")
                  Next
                End If
              %>
            </td>
          </tr>
          <tr>
            <td>
              <label for="maxadjustment"><% Sendb(Copient.PhraseLib.Lookup("term.maxadjustment", LanguageID))%>:</label>
            </td>
            <td>
              <%
                If DeliverableID = 0 Then
                  Send("<input type=""radio"" id=""MaxAdjOff"" name=""MaxAdjSwitch"" value=""0"" " & IIf(MyCommon.Extract_Val(Request.QueryString("MaxAdjSwitch")) = 0, "checked=""checked""", "") & " onclick=""toggleMaxAdjustment(0);"" /><label for=""MadAdjOff"" >" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "</label>")
                  Send("<br />")
                  Send("<input type=""radio"" id=""MaxAdjOn"" name=""MaxAdjSwitch"" value=""1"" " & IIf(MyCommon.Extract_Val(Request.QueryString("MaxAdjSwitch")) = 1, "checked=""checked""", "") & " onclick=""toggleMaxAdjustment(1);"" />")
                      Send("<input type=""text"" class=""shorter"" id=""maxadjustment"" name=""maxadjustment"" maxlength=""6"" value=""" & IIf(MyCommon.Extract_Val(Request.QueryString("MaxAdjSwitch")) = 1, MaxAdjustmentDisplay, "") & """ " & IIf(DisabledAttribute = "" And MyCommon.Extract_Val(Request.QueryString("MaxAdjSwitch")) = 1, "", " disabled=""disabled""") & "/>")
                Else
                  Send("<input type=""radio"" id=""MaxAdjOff"" name=""MaxAdjSwitch"" value=""0""" & IIf(MaxAdjustmentEnabled = False, " checked=""checked""", "") & " onclick=""toggleMaxAdjustment(0);"" /><label for=""MaxAdjOff"">" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "</label>")
                  Send("<br />")
                  Send("<input type=""radio"" id=""MaxAdjOn"" name=""MaxAdjSwitch"" value=""1""" & IIf(MaxAdjustmentEnabled = True, " checked=""checked""", "") & " onclick=""toggleMaxAdjustment(1);"" />")
                      Send("<input type=""text"" class=""shorter"" id=""maxadjustment"" name=""maxadjustment"" maxlength=""6"" value=""" & IIf(MaxAdjustmentEnabled, MaxAdjustmentDisplay, "") & """" & IIf(DisabledAttribute = "" And MaxAdjustmentEnabled, "", " disabled=""disabled""") & " />")
                End If
              %>
            </td>
          </tr>
        </table>
      </div>
      
      <%If EngineID = 6 Or EngineID = 2 Then%>
      <div class="box" id="department">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.chargebackdepartment", LanguageID))%>
          </span>
        </h2>
        <%
          MyCommon.QueryStr = "select ChargebackDeptID from CPE_DeliverablePoints where DeliverableID=" & DeliverableID
          dt = MyCommon.LRT_Select
          If (dt.Rows.Count > 0) Then
            ChargebackDeptID = MyCommon.NZ(dt.Rows(0).Item("ChargebackDeptID"), 0)
          Else
            ChargebackDeptID = 0
          End If
          
          If (EdiscountType = 1) Or (EdiscountType = 3) Or (EdiscountType = 4) Then
            sQuery = "select ChargeBackDeptID, ExternalID, Name, PhraseID from ChargebackDepts with (NoLock) where ChargeBackDeptID not in (0,10,14) "
            GlobalDepts = "0"
            If (BannersEnabled) Then
              MyCommon.QueryStr = "select BO.BannerID, BAN.AllBanners from BannerOffers BO with (NoLock) " & _
                                  "inner join Banners BAN with (NoLock) on BAN.BannerID=BO.BannerID and BAN.Deleted=0 " & _
                                  "where OfferID=" & OfferID
              rst2 = MyCommon.LRT_Select
              AllBanners = (rst2.Rows.Count = 1 AndAlso MyCommon.NZ(rst2.Rows(0).Item("AllBanners"), False))

              If (rst2.Rows.Count = 1 AndAlso Not AllBanners) Then
                sQuery &= " and BannerID=" & MyCommon.NZ(rst2.Rows(0).Item("BannerID"), -1) & " or ChargebackDeptID in (" & GlobalDepts & ")"
                ' Find the default chargeback dept ID so that, if one isn't already assigned, the default will be preselected.
                MyCommon.QueryStr = "select DefaultChargebackDeptID from Banners where BannerID=" & MyCommon.NZ(rst2.Rows(0).Item("BannerID"), -1) & ";"
                rst = MyCommon.LRT_Select
                If (rst.Rows.Count = 1) Then
                  DefaultChrgBack = MyCommon.NZ(rst.Rows(0).Item("DefaultChargebackDeptID"), -1)
                End If
              Else
                sQuery &= " and (BannerID=0 or BannerID IS NULL) " & " or ChargebackDeptID in (" & GlobalDepts & ")"
              End If
            End If
            MyCommon.QueryStr = sQuery & " order by ExternalID;"
            rst = MyCommon.LRT_Select
            Send("<select class=""longer"" id=""chargeback"" name=""chargeback"" " & DisabledAttribute & ">")
            If Not (rst.Rows.Count = 0) Then
              For Each row In rst.Rows
                Sendb("<option value=""" & MyCommon.NZ(row.Item("ChargeBackDeptID"), 0) & """")
                If MyCommon.NZ(row.Item("ChargeBackDeptID"), -1) = ChargebackDeptID OrElse (ChargebackDeptID = -1 AndAlso MyCommon.NZ(row.Item("ChargeBackDeptID"), -1) = DefaultChrgBack) Then
                  Sendb(" selected=""selected""")
                ElseIf DeliverableID = 0 Then
                  If (Request.QueryString("chargeback") = MyCommon.NZ(row.Item("ChargeBackDeptID"), -1)) Then
                    Sendb(" selected=""selected""")
                  End If
                End If
                Sendb(">")
                If ((row.Item("ExternalID") = "") Or (row.Item("ExternalID") = "0")) Then
                Else
                  Sendb(row.Item("ExternalID") & " - ")
                End If
                If (IsDBNull(row.Item("PhraseID"))) Then
                  Send(MyCommon.NZ(row.Item("Name"), ""))
                Else
                  If (row.Item("PhraseID") = 0) Then
                    Send(MyCommon.NZ(row.Item("Name"), ""))
                  Else
                    Sendb(Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID))
                  End If
                End If
                Send("</option>")
              Next
            End If
              Send("</select>")
          Else
            Send("<input type=""hidden"" id=""chargeback"" name=""chargeback"" value=""1"" />")
          End If
        %>
        <br />
        <hr class="hidden" />
      </div>
      <%End If%>
      
      <div class="box" id="scorecards">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.scorecard", LanguageID))%>
          </span>
        </h2>
        <div id="scorecardinputs">
        </div>
        <hr class="hidden" />
      </div>
    </div>
  </div>

<script runat="server">
  Function Create_Reward(ByVal OfferID As Long, ByVal EngineID As Long, ByVal ROID As Long, ByVal ProgramID As Long, ByVal Quantity As Long, ByVal Phase As Long, ByVal ScorecardID As Long, ByVal ScorecardDesc As String, ByVal ScorecardBold As Boolean, ByVal MaxAdjustmentEnabled As Boolean, ByVal MaxAdjustment As Long, ByVal ChargebackDeptID As Integer, ByRef DeliverableID As Long) As Boolean
    Dim MyCommon As New Copient.CommonInc
    Dim Status As Integer = 0
    
    Try
      MyCommon.QueryStr = "dbo.pa_CPE_AddPointsReward"
      MyCommon.Open_LogixRT()
      MyCommon.Open_LRTsp()
      
      MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.Int, 4).Value = OfferID
      MyCommon.LRTsp.Parameters.Add("@ROID", SqlDbType.Int, 4).Value = ROID
      MyCommon.LRTsp.Parameters.Add("@ProgramID", SqlDbType.Int, 4).Value = ProgramID
      MyCommon.LRTsp.Parameters.Add("@Quantity", SqlDbType.Int, 4).Value = Quantity
      MyCommon.LRTsp.Parameters.Add("@Phase", SqlDbType.Int, 4).Value = Phase
      MyCommon.LRTsp.Parameters.Add("@ScorecardID", SqlDbType.Int, 4).Value = ScorecardID
      MyCommon.LRTsp.Parameters.Add("@ScorecardDesc", SqlDbType.NVarChar, 100).Value = ScorecardDesc
      MyCommon.LRTsp.Parameters.Add("@ScorecardBold", SqlDbType.Bit, 1).Value = 0 ' Hardcoded to zero/false; to restore, set this back to 'ScorecardBold'
      If MaxAdjustmentEnabled Then
        MyCommon.LRTsp.Parameters.Add("@MaxAdjustment", SqlDbType.Int, 4).Value = MaxAdjustment
      End If
      MyCommon.LRTsp.Parameters.Add("@ChargebackDeptID", SqlDbType.Int, 4).Value = ChargebackDeptID
      MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.Int, 4).Direction = ParameterDirection.Output
      MyCommon.LRTsp.Parameters.Add("@DeliverableID", SqlDbType.Int, 4).Direction = ParameterDirection.Output
      
      MyCommon.LRTsp.ExecuteNonQuery()
      Status = MyCommon.LRTsp.Parameters("@Status").Value
      DeliverableID = MyCommon.LRTsp.Parameters("@DeliverableID").Value
      
      MyCommon.Close_LRTsp()
    Catch ex As Exception
      Status = -1
      Send(ex.ToString)
    Finally
      MyCommon.Close_LogixRT()
      MyCommon = Nothing
    End Try
    
    Return (Status = 0)
  End Function
  
  Function Update_Reward(ByVal OfferID As Long, ByVal EngineID As Long, ByVal ROID As Long, ByVal DeliverableID As Long, ByVal ProgramID As Long, ByVal Quantity As Long, ByVal Phase As Long, ByVal ScorecardID As Long, ByVal ScorecardDesc As String, ByVal ScorecardBold As Boolean, ByVal MaxAdjustmentEnabled As Boolean, ByVal MaxAdjustment As Long, ByVal ChargebackDeptID As Integer) As Boolean
    Dim MyCommon As New Copient.CommonInc
    Dim Status As Integer = 0
    
    Try
      MyCommon.QueryStr = "dbo.pa_CPE_UpdatePointsReward"
      MyCommon.Open_LogixRT()
      MyCommon.Open_LRTsp()
      
      MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.Int, 4).Value = OfferID
      MyCommon.LRTsp.Parameters.Add("@ROID", SqlDbType.Int, 4).Value = ROID
      MyCommon.LRTsp.Parameters.Add("@DeliverableID", SqlDbType.Int, 4).Value = DeliverableID
      MyCommon.LRTsp.Parameters.Add("@ProgramID", SqlDbType.Int, 4).Value = ProgramID
      MyCommon.LRTsp.Parameters.Add("@Quantity", SqlDbType.Int, 4).Value = Quantity
      MyCommon.LRTsp.Parameters.Add("@Phase", SqlDbType.Int, 4).Value = Phase
      MyCommon.LRTsp.Parameters.Add("@ScorecardID", SqlDbType.Int, 4).Value = ScorecardID
      MyCommon.LRTsp.Parameters.Add("@ScorecardDesc", SqlDbType.NVarChar, 100).Value = ScorecardDesc
      MyCommon.LRTsp.Parameters.Add("@ScorecardBold", SqlDbType.Bit, 1).Value = 0 ' Hardcoded to zero/false; to restore, set this back to 'ScorecardBold'
      If MaxAdjustmentEnabled Then
        MyCommon.LRTsp.Parameters.Add("@MaxAdjustment", SqlDbType.Int, 4).Value = MaxAdjustment
      End If
      MyCommon.LRTsp.Parameters.Add("@ChargebackDeptID", SqlDbType.Int, 4).Value = ChargebackDeptID
    
      MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.Int, 4).Direction = ParameterDirection.Output
      
      MyCommon.LRTsp.ExecuteNonQuery()
      Status = MyCommon.LRTsp.Parameters("@Status").Value
      
      MyCommon.Close_LRTsp()
    Catch ex As Exception
      Status = -1
      Send(ex.ToString)
    Finally
      MyCommon.Close_LogixRT()
      MyCommon = Nothing
    End Try
    
    Return (Status = 0)
  End Function
  
  Sub Create_RewardTiers(ByVal DPPKID As Long, ByVal TierLevel As Long, ByVal Quantity As Long)
    Dim MyCommon As New Copient.CommonInc
    
    MyCommon.QueryStr = "dbo.pa_CPE_AddPointsRewardTiers"
    
    MyCommon.Open_LogixRT()
    MyCommon.Open_LRTsp()
    
    MyCommon.LRTsp.Parameters.Add("@DPPKID", SqlDbType.Int, 4).Value = DPPKID
    MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int, 4).Value = TierLevel
    MyCommon.LRTsp.Parameters.Add("@Quantity", SqlDbType.Int, 4).Value = Quantity
    MyCommon.LRTsp.ExecuteNonQuery()
    
    MyCommon.Close_LRTsp()
    MyCommon.Close_LogixRT()
    MyCommon = Nothing
    
  End Sub
  
  Function ValidateTier(ByVal TierLevels As Integer, ByVal NumArray() As Decimal, ByRef HasZeroVal As Boolean)
    Dim ValidTier As Boolean = False
    Dim AllNeg As Boolean = False
    Dim Cont As Boolean = False
    Dim t1, t2 As Decimal
    Dim t As Integer = 0

    'Tier level validation code
    If TierLevels > 1 Then
      ValidTier = True
      'Are the tiers negative
      For t = 1 To TierLevels
        t1 = NumArray(t - 1)
        If t1 = 0 Then
          ValidTier = False
          HasZeroVal = True
          Exit For
        ElseIf t1 < 0 Then
          AllNeg = True
        Else
          AllNeg = False
          Exit For
        End If
      Next

      If ValidTier Then
        If AllNeg Then
          For t = 1 To TierLevels - 1
            t2 = NumArray(t)
            t1 = NumArray(t - 1)
            Send(t2 & "<" & t1)
            If t2 < t1 Then
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
            If t2 > t1 Then
              ValidTier = True
            Else
              ValidTier = False
              Exit For
            End If
          Next
        End If
      End If
    ElseIf TierLevels = 1 AndAlso NumArray.Length > 0 Then
      HasZeroVal = (NumArray(0) = 0)
      ValidTier = Not HasZeroVal
    Else
      ValidTier = True
    End If

    Return ValidTier
  End Function
</script>

<script type="text/javascript">
<% If (CloseAfterSave) Then %>
  window.location = 'close.aspx';
<% End If %>
</script>
<%
  If (ProgramID > 0) Then
    Send("<script type=""text/javascript"">")
    Send("  selectByPointsID(" & ProgramID & ");")
    Send("</script>")
  End If
%>
</form>
<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd("mainform", FocusField)
  Logix = Nothing
  MyCommon = Nothing
%>
