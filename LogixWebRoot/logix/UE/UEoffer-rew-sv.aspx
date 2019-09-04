<%@ Page Language="vb" Debug="true" CodeFile="ue-cb.vb" Inherits="UECB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
    ' *****************************************************************************
    ' * FILENAME: UEoffer-rew-sv.aspx 
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
    Dim row As DataRow
    Dim rst2 As DataTable
    Dim OfferID As Long
    Dim RewardID As Long
    Dim ProgramID As Long
    Dim DeliverableID As Long
    Dim DSVPKID As Long
    Dim Phase As Integer
	Dim Quantity As Integer = 0
	Dim Multiplier As Decimal = 1
	Dim FractionalQuantity As Decimal = 0.0
    Dim QuantityDisplay As String = ""
    Dim TierQuantity As Decimal = 0.0
    Dim MaxAdjustment As Long = 0
    Dim MaxAdjustmentDisplay As String = ""
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
    Dim EngineID As Integer = 2
    Dim ScorecardAtProgram As Integer = 0
    Dim ScorecardDescAtProgram As String = ""
    Dim ScorecardBoldAtProgram As Boolean = False
    Dim ScorecardID As Integer = 0
    Dim ScorecardDesc As String = ""
    Dim ScorecardBold As Boolean = False
    Dim BannersEnabled As Boolean = True
    Dim TierLevels As Integer = 1
    Dim t As Integer = 1
    Dim ValidTiers As Boolean = False
    Dim NegTiers As Boolean = False
    Dim ZeroTier As Boolean = False
    Dim NumArray(99) As Decimal
    Dim RewardRequired As Boolean = True
    Dim MLI As New Copient.Localization.MultiLanguageRec
    Dim IsAnyCustomer As Boolean = False
    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If

    Response.Expires = 0
    MyCommon.AppName = "UEoffer-rew-sv.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
    Localization = New Copient.Localization(MyCommon)

    BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")

    OfferID = Request.QueryString("OfferID")
    'This would redirect to summary page if offer is awaiting deployment or if offer is awaiting recommendations.
    CheckIfValidOffer(MyCommon, OfferID)
    RewardID = MyCommon.Extract_Val(Request.QueryString("RewardID"))
    ProgramID = MyCommon.Extract_Val(Request.QueryString("ProgramID"))
    QuantityDisplay = Request.QueryString("t1_quantity")
    RewardRequired = (MyCommon.Extract_Val(GetCgiValue("requiredToDeliver")) = 1)
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

    Dim isTranslatedOffer As Boolean =MyCommon.IsTranslatedUEOffer(OfferID,  MyCommon)
    Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249)="1",True,False)

    Dim bEnableAdditionalLockoutRestrictionsOnOffers As Boolean = IIf(MyCommon.Fetch_SystemOption(260) = "1", True, False)
    Dim bOfferEditable As Boolean = MyCommon.IsOfferEditablePastLockOutPeriod(Logix.UserRoles.EditOfferPastLockoutPeriod, MyCommon, OfferID)

    Dim tmp As Boolean
    If (QuantityDisplay IsNot Nothing AndAlso Not Int32.TryParse(QuantityDisplay, tmp)) Then
        infoMessage = Copient.PhraseLib.Lookup("error.sv-quantity-invalid", LanguageID)
        bError = True
    End If
	
	If (Decimal.TryParse(QuantityDisplay, System.Globalization.NumberStyles.AllowDecimalPoint, MyCommon.GetAdminUser.Culture, Quantity)) Then
		FractionalQuantity = MyCommon.Extract_Decimal(QuantityDisplay, MyCommon.GetAdminUser.Culture)
		QuantityDisplay = FractionalQuantity.ToString()
	  Else
		FractionalQuantity = 0
		QuantityDisplay = ""
	  End If
	  
	 if QuantityDisplay.IndexOf(".") > -1 Then
		Quantity = FractionalQuantity * 100
		Multiplier = 0.01
	Else
		Quantity = FractionalQuantity
		Multiplier = 1
	End If
	
    ScorecardID = MyCommon.Extract_Val(Request.QueryString("ScorecardID"))
    ScorecardDesc = IIf(Request.QueryString("ScorecardDesc") <> "", Request.QueryString("ScorecardDesc"), "")
    ScorecardBold = IIf(Request.QueryString("ScorecardBold") <> "", True, False)

    DeliverableID = MyCommon.Extract_Val(Request.QueryString("DeliverableID"))
    Phase = MyCommon.Extract_Val(Request.QueryString("phase"))
    If (Phase = 0) Then Phase = MyCommon.Extract_Val(Request.Form("Phase"))
    If (Phase = 0) Then Phase = 3

    TouchPoint = MyCommon.Extract_Val(Request.QueryString("tp"))
    If (TouchPoint > 0) Then TpROID = MyCommon.Extract_Val(Request.QueryString("roid"))

    MyCommon.QueryStr = "select PKID from CPE_DeliverableStoredValue with (NoLock) where DeliverableID=" & DeliverableID & ";"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
        DSVPKID = MyCommon.NZ(rst.Rows(0).Item("PKID"), 0)
    End If

    MyCommon.QueryStr = "select TierLevels from CPE_RewardOptions with (NoLock) where RewardOptionID=" & RewardID & ";"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
        TierLevels = MyCommon.NZ(rst.Rows(0).Item("TierLevels"), 1)
    End If

    ' Fetch the name
    MyCommon.QueryStr = "Select IncentiveName, IsTemplate, FromTemplate from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
        Name = MyCommon.NZ(rst.Rows(0).Item("IncentiveName"), "")
        IsTemplate = MyCommon.NZ(rst.Rows(0).Item("IsTemplate"), False)
        FromTemplate = MyCommon.NZ(rst.Rows(0).Item("FromTemplate"), False)
    End If
    IsTemplateVal = IIf(IsTemplate, "IsTemplate", "Not")

    If (Request.QueryString("save") <> "" AndAlso infoMessage = "") Then
        'Tier level validation code
        For t = 1 To TierLevels
            NumArray(t - 1) = MyCommon.Extract_Val(Request.QueryString("t" & t & "_quantity"))
        Next
        If NumArray(0) < 0 Then
            NegTiers = True
        End If
        If (TierLevels = 1) AndAlso (NumArray(0) = 0) Then
            ZeroTier = True
        End If
        ValidTiers = ValidateTier(TierLevels, NumArray)

        If ValidTiers Then
            ' Preliminary check for valid scorecard print line:
            If (ScorecardID > 0 OrElse ScorecardAtProgram > 0) Then
                If ScorecardDesc = "" Then
                    MyCommon.QueryStr = "select ScorecardDesc from StoredValuePrograms with (NoLock) where SVProgramID=" & ProgramID & ";"
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
            If (OfferID > 0 AndAlso RewardID > 0 AndAlso ProgramID > 0 AndAlso DeliverableID = 0) Then
                CreateROID = IIf(TpROID > 0, TpROID, RewardID)
				If (Create_Reward(OfferID, CreateROID, ProgramID, Quantity, Multiplier, Phase, ScorecardID, ScorecardDesc, ScorecardBold, MaxAdjustment, RewardRequired, DeliverableID)) Then
                    ' Ensure we have the proper newly-made PKID
                    MyCommon.QueryStr = "select PKID from CPE_DeliverableStoredValue with (NoLock) where RewardOptionID=" & RewardID & ";"
                    rst = MyCommon.LRT_Select
                    If rst.Rows.Count > 0 Then
                        DSVPKID = MyCommon.NZ(rst.Rows(0).Item("PKID"), 0)
                    End If
                    ' Delete existing tier values from the tiers table, then insert new values
                    MyCommon.QueryStr = "delete from CPE_DeliverableStoredValueTiers with (RowLock) where DSVPKID in (0, " & DSVPKID & ");"
                    MyCommon.LRT_Execute()
                    t = 1
                    For t = 1 To TierLevels
						TierQuantity = Decimal.Parse(Request.QueryString("t" & t & "_quantity"))
                        Create_RewardTiers(DSVPKID, t, TierQuantity)
                    Next
                    infoMessage = Copient.PhraseLib.Lookup("term.ChangesWereSaved", LanguageID)
                    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Reward.createstoredvalue", LanguageID))
                Else
                    infoMessage = Copient.PhraseLib.Lookup("ueoffer-rew-points.ErrorOnSave", LanguageID)
                    bError = True
                End If
            ElseIf (OfferID > 0 AndAlso RewardID > 0 AndAlso ProgramID > 0 AndAlso DeliverableID > 0) Then
                ' Modify existing deliverable
                MyCommon.QueryStr = "update CPE_DeliverableStoredValue with (RowLock) set SVProgramID=" & ProgramID & ", Quantity=" & Quantity & ", " & _
                                    " ScorecardID=" & ScorecardID & ", ScorecardDesc='" & ScorecardDesc & "', ScorecardBold=" & IIf(ScorecardBold, 1, 0) & _
                                    " where RewardOptionID=" & RewardID & " and DeliverableID=" & DeliverableID
                MyCommon.LRT_Execute()

                ' update the required status for this deliverable
                MyCommon.QueryStr = "update CPE_Deliverables with (RowLock) set Required=" & IIf(RewardRequired, 1, 0) & ", LastUpdate=getdate() " & _
                                    " where DeliverableID=" & DeliverableID
                MyCommon.LRT_Execute()

                ' Delete existing tier values from the tiers table, then insert new values
                MyCommon.QueryStr = "delete from CPE_DeliverableStoredValueTiers with (RowLock) where DSVPKID in (0, " & DSVPKID & ");"
                MyCommon.LRT_Execute()
                t = 1
                For t = 1 To TierLevels
					 TierQuantity = Decimal.Parse(Request.QueryString("t" & t & "_quantity"))
                    Create_RewardTiers(DSVPKID, t, TierQuantity)
                Next
                MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Reward.updatestoredvalue", LanguageID))
            End If
            MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
            MyCommon.LRT_Execute()
            ResetOfferApprovalStatus(OfferID)
            'Update multi-language:
            'Name
            MLI.ItemID = DSVPKID
            MLI.MLTableName = "CPE_DeliverableSVTranslations"
            MLI.MLColumnName = "ScorecardDesc"
            MLI.MLIdentifierName = "DeliverableSVID"
            MLI.StandardTableName = "CPE_DeliverableStoredValue"
            MLI.StandardColumnName = "ScorecardDesc"
            MLI.StandardIdentifierName = "PKID"
            MLI.StandardValue = ScorecardDesc
            MLI.InputName = "ScorecardDesc"
            MLI.InputID = "ScorecardDesc"
            Localization.SaveTranslationInputs(MyCommon, MLI, Request.QueryString, 9)
            CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
        Else
            If NegTiers Then
                infoMessage = Copient.PhraseLib.Lookup("error.tiervalues-negative", LanguageID)
            ElseIf ZeroTier Then
                infoMessage = Copient.PhraseLib.Lookup("error.nonzero", LanguageID)
            Else
                infoMessage = Copient.PhraseLib.Lookup("error.tiervalues", LanguageID)
            End If
            bError = True
        End If
    End If
NoSave:

    'Update the templates permission if necessary
    If (Request.QueryString("save") <> "" AndAlso Request.QueryString("IsTemplate") = "IsTemplate") Then
        ' time to update the status bits for the templates
        Dim form_Disallow_Edit As Integer = 0
        If (Request.QueryString("Disallow_Edit") = "on") Then
            form_Disallow_Edit = 1
        End If
        MyCommon.QueryStr = "update CPE_Deliverables with (RowLock) set DisallowEdit=" & form_Disallow_Edit & _
                            " where DeliverableID=" & DeliverableID
        MyCommon.LRT_Execute()
    End If

    If (IsTemplate Or FromTemplate) Then
        ' lets dig the permissions if its a template
        MyCommon.QueryStr = "select DisallowEdit from CPE_Deliverables with (NoLock) where DeliverableID=" & DeliverableID
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
            Disallow_Edit = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit"), True)
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

    Send_HeadBegin("term.offer", "term.storedvaluereward", OfferID)
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts()
%>
<script type="text/javascript">
// This is the javascript array holding the function list
// The PrintJavascriptArray ASP function can be used to print this array.
<%
If  UEOffer_Has_AnyCustomer(MyCommon, OfferID) Then
    IsAnyCustomer = true
End If
  If EngineID = 2 Then
        MyCommon.querystr = "Select SVProgramID as ProgramID,Name as ProgramName from StoredValuePrograms with (NoLock) " & _
                        "where deleted=0 and SVTypeID not in (3) and SVPRogramID Not in  " & _
                        "( select SVProgramID from CPE_Deliverables D with (NoLock)  " & _
                        "  inner join CPE_DeliverableStoredValue DSV with (NoLock) on DSV.PKID = D.OutputID " & _
                        "  where D.Deleted=0 and D.RewardOptionID = " & RewardID & " and D.DeliverableTypeID=11 " & _
                        "     and D.DeliverableID <> " & DeliverableID & " and DSV.Deleted=0) " & _
                        "order by Name;"
  Else
    MyCommon.querystr = "Select SVProgramID as ProgramID,Name as ProgramName from StoredValuePrograms with (NoLock) " & _
                        "where deleted=0" & _ 
                        IIf(IsAnyCustomer, " and SVPRogramID in ( Select SVProgramID from SVProgramsPromoEngineSettings Where AllowAnyCustomer = 1)","" ) & _
                        " and SVPRogramID Not in  ( select SVProgramID from CPE_Deliverables D with (NoLock)  " & _
                        "  inner join CPE_DeliverableStoredValue DSV with (NoLock) on DSV.PKID = D.OutputID " & _
                        "  where D.Deleted=0 and D.RewardOptionID = " & RewardID & " and D.DeliverableTypeID=11 " & _
                        "     and D.DeliverableID <> " & DeliverableID & " and DSV.Deleted=0) " & _
                        "order by Name; "
  End If
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
  
  // Create a regular expression
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
  
  if(selectedValue != "")
  {
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
  var qtyElem = document.getElementById("quantity");
  var elemProgram = document.getElementById("ProgramID");
  var msg = '';
  
  if (elem != null && elem.options.length == 0) {
    retVal = false;
    msg += '* <% Sendb(Copient.PhraseLib.Lookup("CPE-rew-membership.selectpoints", LanguageID)) %>'
    elem.focus();
  } else if (elem !=null && elemProgram != null) {
    elemProgram.value = elem.options[0].value;
  }
  if (qtyElem != null) {
    // trim the string
    var qtyVal = qtyElem.value.replace(/^\s+|\s+$/g, ''); 
    if (qtyVal == "" || isNaN(qtyVal) || !isSignedInteger(qtyVal)) {
      retVal = false;
      if (msg != '') { msg += '\n\r\n\r'; }
      msg += '* <% Sendb(Copient.PhraseLib.Lookup("CPE-rew-membership.enterquantity", LanguageID)) %>';
      qtyElem.focus();
      qtyElem.select();
    }
  }
  if (msg != '') {
    alert(msg);
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
  xmlhttpPost('/logix/ScorecardFeeds.aspx', 'ScorecardFieldsForReward=1&EngineID=<%Sendb(EngineID)%>&OfferID=<%Sendb(OfferID)%>&DeliverableID=<%Sendb(DeliverableID)%>&ProgramID=' + ProgramID + '&ScorecardTypeID=2');
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
        Send("  if (opener != null) {")
        Send("    var newlocation = 'UEoffer-rew.aspx?OfferID=" & OfferID & "'; ")
        Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
        Send("  opener.location = 'UEoffer-rew.aspx?OfferID=" & OfferID & "'; ")
    Else
        Send("  if (opener != null) {")
        Send("    var newlocation = '/logix/CPEoffer-rew.aspx?OfferID=" & OfferID & "'; ")
        Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
        Send("  opener.location = '/logix/CPEoffer-rew.aspx?OfferID=" & OfferID & "'; ")
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
<form action="UEoffer-rew-sv.aspx" id="mainform" name="mainform" onsubmit="return validateEntry();">
  <div id="intro">
    <input type="hidden" id="OfferID" name="OfferID" value="<% Sendb(OfferID) %>" />
    <input type="hidden" id="Name" name="Name" value="<% Sendb(Name) %>" />
    <input type="hidden" id="RewardID" name="RewardID" value="<% Sendb(RewardID) %>" />
    <input type="hidden" id="DeliverableID" name="DeliverableID" value="<% sendb(DeliverableID) %>" />
    <input type="hidden" id="Phase" name="Phase" value="<% sendb(Phase) %>" />
    <input type="hidden" id="ProgramID" name="ProgramID" value="<% Sendb(ProgramID) %>" />
    <input type="hidden" id="roid" name="roid" value="<%Sendb(TpROID) %>" />
    <input type="hidden" id="id" name="tp" value="<%Sendb(TouchPoint) %>" />
    <input type="hidden" id="IsTemplate" name="IsTemplate" value="<% Sendb(IsTemplateVal)%>" />
    <input type="hidden" id="EngineID" name="EngineID" value="<%Sendb(EngineID) %>" />
    <%
      If (IsTemplate) Then
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.storedvaluereward", LanguageID), VbStrConv.Lowercase) & "</h1>")
      Else
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.storedvaluereward", LanguageID), VbStrConv.Lowercase) & "</h1>")
      End If
    %>
    <div id="controls">
      <% If (IsTemplate) Then%>
      <span class="temp">
        <input type="checkbox" class="tempcheck" id="Disallow_Edit" name="Disallow_Edit"<% if(disallow_edit)then sendb(" checked=""checked""") %> />
        <label for="Disallow_Edit"><% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
      </span>
      <% End If%>      
      <% 
          m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
          If((Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
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
            <% Sendb(Copient.PhraseLib.Lookup("term.storedvalueprogram", LanguageID))%>
          </span>
        </h2>
        <input type="radio" id="functionradio1" name="functionradio" <% if(MyCommon.Fetch_SystemOption(175)= "1" ) then sendb(" checked=""checked""") %> <% sendb(DisabledAttribute) %> /><label for="functionradio1"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
        <input type="radio" id="functionradio2" name="functionradio" <% if(MyCommon.Fetch_SystemOption(175)= "2" ) then sendb(" checked=""checked""") %> /><label for="functionradio2"<% sendb(DisabledAttribute) %>><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
        <input type="text" class="medium" id="functioninput" name="functioninput" maxlength="100" onkeydown="handleKeyDown(event);" onkeyup="handleKeyUp(200);" value=""<% sendb(DisabledAttribute) %> /><br />
        <select class="longer" id="functionselect" name="functionselect" onkeydown="return handleSlctKeyDown(event);" onclick="handleSelectClick();" ondblclick="addToSelect();" size="16"<% sendb(DisabledAttribute) %>>
          <%
            If UEOffer_Has_AnyCustomer(MyCommon, OfferID) Then
              IsAnyCustomer = True
            End If
            
            If EngineID = 2 Then
              MyCommon.QueryStr = "Select SVProgramID as ProgramID,Name as ProgramName from StoredValuePrograms with (NoLock) " & _
                              "where deleted=0 and SVTypeID not in (3) and SVPRogramID Not in  " & _
                              "( select SVProgramID from CPE_Deliverables D with (NoLock)  " & _
                              "  inner join CPE_DeliverableStoredValue DSV with (NoLock) on DSV.PKID = D.OutputID " & _
                              "  where D.Deleted=0 and D.RewardOptionID = " & RewardID & " and D.DeliverableTypeID=11 " & _
                              "     and D.DeliverableID <> " & DeliverableID & " and DSV.Deleted=0) " & _
                              "order by Name;"
            Else
              MyCommon.QueryStr = "Select SVProgramID as ProgramID,Name as ProgramName from StoredValuePrograms with (NoLock) " & _
                                  "where deleted=0 " & _
                                 IIf(IsAnyCustomer, " and SVPRogramID in ( Select SVProgramID from SVProgramsPromoEngineSettings Where AllowAnyCustomer = 1)", "") & _
                                  "and SVPRogramID Not in  ( select SVProgramID from CPE_Deliverables D with (NoLock)  " & _
                                  "  inner join CPE_DeliverableStoredValue DSV with (NoLock) on DSV.PKID = D.OutputID " & _
                                  "  where D.Deleted=0 and D.RewardOptionID = " & RewardID & " and D.DeliverableTypeID=11 " & _
                                  "     and D.DeliverableID <> " & DeliverableID & " and DSV.Deleted=0) " & _
                                  "order by Name; "
            End If
            rst = MyCommon.LRT_Select
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
        <%
          'Max adjustment (not using this now)
          Send("<input type=""hidden"" id=""maxadjustment"" name=""maxadjustment"" value=""" & MaxAdjustment & """ />")
          
          Send("<table summary=""" & Copient.PhraseLib.Lookup("term.distribution", LanguageID) & """>")
		  MyCommon.QueryStr = "Select TierLevel, Quantity, Multiplier from CPE_DeliverableStoredValueTiers with (NoLock) " & _
                              "where DSVPKID=" & DSVPKID & " order by TierLevel;"
          rst = MyCommon.LRT_Select
          If DeliverableID = 0 Then
            For t = 1 To TierLevels
              Send("<tr>")
              Send("  <td>")
              Sendb("    <label for=""t" & t & "_quantity"">")
              If TierLevels > 1 Then
                Sendb(Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & " " & StrConv(Copient.PhraseLib.Lookup("term.quantity", LanguageID), VbStrConv.Lowercase) & ":")
              Else
                Sendb(Copient.PhraseLib.Lookup("term.quantityawarded", LanguageID) & ":")
              End If
              Send("</label>")
              Send("  </td>")
              Send("  <td>")
              If Request.QueryString("t" & t & "_quantity") <> "" Then
                Sendb("    <input type=""text"" class=""shorter"" id=""t" & t & "_quantity"" name=""t" & t & "_quantity"" maxlength=""6"" value=""" & IIf(MyCommon.Extract_Val(Request.QueryString("t" & t & "_quantity")) <> "", MyCommon.Extract_Val(Request.QueryString("t" & t & "_quantity")), 0) & """" & DisabledAttribute & " />")
              Else
                Sendb("    <input type=""text"" class=""shorter"" id=""t" & t & "_quantity"" name=""t" & t & "_quantity"" maxlength=""6"" value=""0""" & DisabledAttribute & " />")
              End If
              Send("    <br />")
              Send("  </td>")
              Send("</tr>")
            Next
          Else
            For t = 1 To TierLevels
              Send("<tr>")
              Send("  <td>")
              Sendb("    <label for=""t" & t & "_quantity"">")
              If TierLevels > 1 Then
                Sendb(Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & " " & StrConv(Copient.PhraseLib.Lookup("term.quantity", LanguageID), VbStrConv.Lowercase) & ":")
              Else
                Sendb(Copient.PhraseLib.Lookup("term.quantityawarded", LanguageID) & ":")
              End If
              Send("</label>")
              Send("  </td>")
              Send("  <td>")
              Sendb("    <input type=""text"" class=""shorter"" id=""t" & t & "_quantity"" name=""t" & t & "_quantity"" maxlength=""6""")
              If t > rst.Rows.Count Then
                Sendb(" value=""0""" & DisabledAttribute & " />")
              Else
				Sendb(" value=""" & MyCommon.NZ(rst.Rows(t - 1).Item("Quantity"), 0) * MyCommon.NZ(rst.Rows(t - 1).Item("Multiplier"), 1) & """" & DisabledAttribute & " />")
              End If
              Send("    <br />")
              Send("  </td>")
              Send("</tr>")
            Next
          End If
          
          RewardRequired = True
          If DeliverableID > 0 Then
            MyCommon.QueryStr = "select Required from CPE_Deliverables with (NoLock) where DeliverableID=" & DeliverableID
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count > 0) Then
              RewardRequired = MyCommon.NZ(rst.Rows(0).Item("Required"), True)
            End If
          End If
          
          Send("<tr><td colspan=""2""><hr /></td></tr>")
          Send("<tr>")
          Send("  <td colspan=""2"">")
          Send("    <input type=""checkbox"" id=""requiredToDeliver"" name=""requiredToDeliver"" value=""1""" & IIf(RewardRequired, " checked=""checked""", "") & " />")
          Send("    <label for=""requiredToDeliver"">" & Copient.PhraseLib.Lookup("ue-reward.reward-required", LanguageID) & "</label>")
          Send("  </td>")
          Send("</tr>")

          Send("</table>")
        %>
        <hr class="hidden" />
      </div>
      
      <div class="box" id="scorecards" style="display:none;">
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
     Function Create_Reward(ByVal OfferID As Long, ByVal ROID As Long, ByVal ProgramID As Long, ByVal Quantity As Integer, ByVal Multiplier As Decimal, ByVal Phase As Long, _
                           ByVal ScorecardID As Long, ByVal ScorecardDesc As String, ByVal ScorecardBold As Boolean, ByVal MaxAdjustment As Long, _
                           ByVal RewardRequired As Boolean, ByRef DeliverableID As Long) As Boolean
      Dim MyCommon As New Copient.CommonInc
      Dim Status As Integer = 0
      
      Try
        MyCommon.QueryStr = "dbo.pa_CPE_AddStoredValueReward"
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
        MyCommon.LRTsp.Parameters.Add("@MaxAdjustment", SqlDbType.Int, 4).Value = MaxAdjustment
        MyCommon.LRTsp.Parameters.Add("@Required", SqlDbType.Bit).Value = IIf(RewardRequired, 1, 0)
		MyCommon.LRTsp.Parameters.Add("@Multiplier", SqlDbType.Real, 9).Value = Multiplier		
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
    
    Sub Create_RewardTiers(ByVal DSVPKID As Long, ByVal TierLevel As Long, ByVal Quantity As Decimal)
      Dim MyCommon As New Copient.CommonInc
      Dim Multiplier As Decimal
	  
      MyCommon.QueryStr = "dbo.pa_CPE_AddStoredValueRewardTiers"
      
      MyCommon.Open_LogixRT()
      MyCommon.Open_LRTsp()
	  
	  if Math.Round(Quantity) - Quantity <> 0 then 
		Multiplier = 0.01
		Quantity = Quantity * 100
	  Else 
		Multiplier = 1
	  End If
	  
      
      MyCommon.LRTsp.Parameters.Add("@DSVPKID", SqlDbType.Int, 4).Value = DSVPKID
      MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int, 4).Value = TierLevel
      MyCommon.LRTsp.Parameters.Add("@Quantity", SqlDbType.Int, 4).Value = Quantity
	  MyCommon.LRTsp.Parameters.Add("@Multiplier", SqlDbType.Real, 9).Value = Multiplier
      MyCommon.LRTsp.ExecuteNonQuery()
      
      MyCommon.Close_LRTsp()
      MyCommon.Close_LogixRT()
      MyCommon = Nothing
      
    End Sub
    
    Function ValidateTier(ByVal TierLevels As Integer, ByVal NumArray() As Decimal)
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
      Else
        If NumArray(0) <> 0 Then
          ValidTier = True
        End If
      End If

      Return ValidTier
    End Function

  </script>
  <script type="text/javascript">
  <% If (CloseAfterSave) Then %>
    window.close();
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
  Send_BodyEnd("mainform", "functioninput")
  Logix = Nothing
  MyCommon = Nothing
%>
