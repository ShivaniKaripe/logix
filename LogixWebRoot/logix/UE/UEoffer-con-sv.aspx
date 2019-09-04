<%@ Page Language="vb" Debug="true" CodeFile="ue-cb.vb" Inherits="UECB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
    ' *****************************************************************************
    ' * FILENAME: UEoffer-con-sv.aspx 
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
    Dim EngineID As Integer = 2
    Dim BannersEnabled As Boolean = True
    Dim TierLevels As Integer = 1
    Dim t As Integer = 1
    Dim TierQty As Decimal
    Dim ValidTier As Boolean = False
    Dim TierDT As DataTable
    Dim ProgramID As Integer = 0
    Dim SVID As Integer = 0
    Dim ValildTier As Boolean = False
    Dim NegTiers As Boolean = False
    Dim IncentiveSVID As Double = 0
    Dim NumArray(99) As Decimal
    Dim IsAnyCustomer As Boolean = False
    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If

    Response.Expires = 0
    MyCommon.AppName = "UEoffer-con-sv.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)

    BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")

    OfferID = Request.QueryString("OfferID")

    'This would redirect to summary page if offer is awaiting deployment or if offer is awaiting recommendations.
    CheckIfValidOffer(MyCommon, OfferID)

    IncentiveSVID = Request.QueryString("IncentiveStoredValueID")

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

    MyCommon.QueryStr = "select IncentiveStoredValueID, SVProgramID from CPE_IncentiveStoredValuePrograms where IncentiveStoredValueID=" & IncentiveSVID & " and Deleted=0;"
    rst = MyCommon.LRT_Select
    If Request.QueryString("selGroups") <> "" Then
        SVID = MyCommon.Extract_Val(Request.QueryString("selGroups"))
    ElseIf rst.Rows.Count > 0 Then
        SVID = MyCommon.NZ(rst.Rows(0).Item("SVProgramID"), 0)
        IncentiveSVID = MyCommon.NZ(rst.Rows(0).Item("IncentiveStoredValueID"), 0)
    End If

    ' see if someone is saving
    If (Request.QueryString("save") <> "" And roid > 0) Then
        'Tier level validation code
        For t = 1 To TierLevels
            NumArray(t - 1) = MyCommon.Extract_Val(Request.QueryString("t" & t & "_QtyForIncentive"))
        Next
        If NumArray(0) < 0 Then NegTiers = True
        ValidTier = ValidateTier(TierLevels, NumArray)

        If ValidTier Then
            'Delete tiers records
            MyCommon.QueryStr = "delete from CPE_IncentiveStoredValueProgramTiers where IncentiveStoredValueID=" & IncentiveSVID & ";"
            MyCommon.LRT_Execute()

            If (Request.QueryString("selGroups") <> "") Then
                Dim QtyString As String = ""
                If TierLevels = 1 Then
                    QtyString = Copient.PhraseLib.Lookup("term.requires", LanguageID) & " " & Request.QueryString("t1_QtyForIncentive")
                Else
                    For t = 1 To TierLevels
                        QtyString += " " & Copient.PhraseLib.Lookup("term.tierlevel", LanguageID) & " " & t & " " & Copient.PhraseLib.Lookup("term.requires", LanguageID) & " " & Request.QueryString("t" & t & "_QtyForIncentive") & ";"
                    Next
                End If
                historyString = Copient.PhraseLib.Detokenize("CPEoffer-con-sv.AlteredGroup", LanguageID, Request.QueryString("selGroups"))
                historyString = historyString + " " + QtyString
                'Check for SV program to update
                MyCommon.QueryStr = "select IncentiveStoredValueID from CPE_IncentiveStoredValuePrograms with (NoLock) where IncentiveStoredValueID=" & IncentiveSVID & " and Deleted=0;"
                rst = MyCommon.LRT_Select()
                If rst.Rows.Count = 0 Then
                    ' ok we need to do some work to set the limit values if there are any otherwise just set to 0
                    ' in theory there should be one set of limit values for each selected groups and possibly an accumulation infos
                    MyCommon.QueryStr = "insert into CPE_IncentiveStoredValuePrograms with (RowLock) (RewardOptionID,SVProgramID,QtyForIncentive,Deleted,LastUpdate) " & _
                                        " values(" & roid & "," & SVID & "," & Request.QueryString("t1_QtyForIncentive") & ",0,getdate());"
                    MyCommon.LRT_Execute()

                    MyCommon.QueryStr = "select IncentiveStoredValueID from CPE_IncentiveStoredValuePrograms where RewardOptionID=" & roid & " and Deleted=0 order by IncentiveStoredValueID desc;"
                    rst = MyCommon.LRT_Select()
                    If rst.Rows.Count > 0 Then
                        IncentiveSVID = rst.Rows(0).Item("IncentiveStoredValueID")
                    End If
                Else
                    MyCommon.QueryStr = "update CPE_IncentiveStoredValuePrograms with (RowLock) set SVProgramID=" & SVID & ",QtyForIncentive=" & MyCommon.NZ(Request.QueryString("t1_QtyForIncentive"), 0) & "," & _
                                        "Deleted=0,LastUpdate=getdate() where IncentiveStoredValueID=" & IncentiveSVID & " and Deleted=0;"
                    MyCommon.LRT_Execute()
                End If

                For t = 1 To TierLevels
                    TierQty = MyCommon.Extract_Val(Request.QueryString("t" & t & "_QtyForIncentive"))
                    MyCommon.QueryStr = "insert into CPE_IncentiveStoredValueProgramTiers (RewardOptionID,IncentiveStoredValueID,TierLevel,Quantity) " & _
                                        "values (" & roid & "," & IncentiveSVID & "," & t & "," & TierQty & ");"
                    MyCommon.LRT_Execute()
                Next
            End If
            MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
            MyCommon.LRT_Execute()
            CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
            ResetOfferApprovalStatus(OfferID)
            MyCommon.Activity_Log(3, OfferID, AdminUserID, historyString)
        Else
            If NegTiers Then
                infoMessage = Copient.PhraseLib.Lookup("error.tiervalues-negative", LanguageID)
                'IncentiveSVID = 0
            Else
                infoMessage = Copient.PhraseLib.Lookup("error.tiervalues", LanguageID)
                'IncentiveSVID = 0
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
        If (Request.QueryString("Disallow_Edit") = "on") Then
            form_Disallow_Edit = 1
        End If
        MyCommon.QueryStr = "update CPE_IncentiveStoredValuePrograms with (RowLock) set DisallowEdit=" & form_Disallow_Edit & _
                            " where RewardOptionID=" & roid & " and Deleted=0;"
        MyCommon.LRT_Execute()
    End If

    If (isTemplate Or FromTemplate) Then
        ' lets dig the permissions if its a template
        MyCommon.QueryStr = "select DisallowEdit from CPE_IncentiveStoredValuePrograms with (NoLock) where RewardOptionID=" & roid & " and Deleted=0;"
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

    Send_HeadBegin("term.offer", "term.storedvaluecondition", OfferID)
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
    IsAnyCustomer = True
  End If
If EngineID = 2 Then
  MyCommon.QueryStr = "Select SVProgramID as ProgramID,Name as ProgramName from StoredValuePrograms with (NoLock) where Deleted=0 and SVTypeID not in (2,3) order by Name;"
Else
If IsAnyCustomer = true Then
 MyCommon.QueryStr ="SELECT SVP.SVProgramID AS ProgramID, SVP.Name AS ProgramName FROM StoredValuePrograms SVP INNER JOIN SVProgramsPromoEngineSettings SPP ON SVP.SVProgramID = SPP.SVProgramID WHERE SVP.Deleted = 0 AND SPP.AllowAnyCustomer = 1"
Else
  MyCommon.QueryStr = "Select SVProgramID as ProgramID,Name as ProgramName from StoredValuePrograms with (NoLock) where Deleted=0 order by Name;"
  End If
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
  
  document.getElementById("functionselect").size = "16";
  
  // Set references to the form elements
  selectObj = document.forms[0].functionselect;
  textObj = document.forms[0].functioninput;
  
  // Remember the function list length for loop speedup
  functionListLength = functionlist.length;
  
  // Set the search pattern depending
  searchPattern = cleanSpecialChar(textObj.value);
  if (document.forms[0].functionradio[0].checked == true) {
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
  if (selectedboxValue != "") {
    selectedboxText = selectboxObj[document.getElementById("selected").selectedIndex].text;
  }
  if (itemSelected == "select1") {
    if (selectedValue != "") {
      // add items to selected box
      document.getElementById('deselect1').disabled=false;
      selectboxObj[selectboxObj.length] = new Option(selectedText,selectedValue);
      if(selectboxObj.length == 1) {
        document.getElementById('select1').disabled=true;      
      }
      if(selectboxObj.length == 0) {
        // nothing in the select box so disable deselect
        document.getElementById('deselect1').disabled=true;
      }
    }
  }
  
  if (itemSelected == "deselect1") {
     if (selectedboxValue != "") {
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
  }
  // remove items from large list that are in the other lists
  removeUsed(false);
  updateButtons();
  return true;
}

function saveForm(){
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
    if (elSel.options[i].value != "") {
      if(selectList != "") { selectList = selectList + ","; }
      selectList = selectList + elSel.options[i].value;
    }
  }
  
  // ok time to build up the hidden variables to pass for saving
  htmlContents = "<input type=\"hidden\" name=\"selGroups\" value=" + selectList + "> ";
  document.getElementById("hiddenVals").innerHTML = htmlContents;
  // alert(htmlContents);
  return true;
}

function validateEntry() {
  var retVal = true;
  var elem = document.getElementById("selected");   
  var qtyElem = document.getElementById("t1_QtyForIncentive");
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
    if (qtyVal == "" || isNaN(qtyVal) || !isInteger(qtyVal) || qtyVal==0) {
      retVal = false;
      if (msg != '') { msg += '\n\r\n\r'; }
      msg += '* <% Sendb(Copient.PhraseLib.Lookup("CPEoffer-con-point.positiveinteger", LanguageID)) %>';
      qtyElem.focus();
      qtyElem.select();
    }
  }
  if (msg != '') {
    alert(msg);
  }
  return retVal;
}

function updateButtons(){
  var elemSelect1 = document.getElementById('select1');
  var elemDeselect1 = document.getElementById('deselect1');
  var elemSave = document.getElementById('save');
  var elemSelected = document.forms[0].selected;
  
  if (elemSelected.length > 0) {
    elemSelect1.disabled=true; 
    elemDeselect1.disabled=false;
    if (elemSave != null) {
      elemSave.disabled=false;
    }
  } else {
    if (elemSave != null) {
      elemSave.disabled=true;
    }
  }
  <%
  m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
    If Not IsTemplate Then
      If Not (Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And Disallow_Edit)) Then
        Send("  disableAll();")
      End If
    Else
      If Not (Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer) Then
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
  <input type="hidden" id="IncentiveStoredValueID" name="IncentiveStoredValueID" value="<% sendb(IncentiveSVID) %>" />
  <input type="hidden" id="roid" name="roid" value="<%sendb(roid) %>" />
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
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.storedvaluecondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
      Else
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.storedvaluecondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
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
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <div id="column1">
      <div class="box" id="selector">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.storedvalueprogram", LanguageID))%>
          </span>
        </h2>
        <input type="radio" id="functionradio1" name="functionradio" <% if(MyCommon.Fetch_SystemOption(175)= "1" ) then sendb(" checked=""checked""") %> <% sendb(disabledattribute) %> /><label for="functionradio1"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
        <input type="radio" id="functionradio2" name="functionradio" <% if(MyCommon.Fetch_SystemOption(175)= "2" ) then sendb(" checked=""checked""") %> <% sendb(disabledattribute) %> /><label for="functionradio2"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
        <input type="text" class="medium" id="functioninput" name="functioninput" maxlength="100" onkeyup="handleKeyUp(200);" value=""<% sendb(disabledattribute) %> /><br />
        <select class="longer" id="functionselect" name="functionselect" size="16"<% sendb(disabledattribute) %>>
          <%
            If UEOffer_Has_AnyCustomer(MyCommon, OfferID) Then
              IsAnyCustomer = True
            End If
            If EngineID = 2 Then
              MyCommon.QueryStr = "Select SVProgramID as ProgramID,Name as ProgramName from StoredValuePrograms with (NoLock) where deleted=0 and SVTypeID not in (2,3) order by Name"
            Else
              If IsAnyCustomer = True Then
                MyCommon.QueryStr = "SELECT SVP.SVProgramID AS ProgramID, SVP.Name AS ProgramName FROM StoredValuePrograms SVP INNER JOIN SVProgramsPromoEngineSettings SPP ON SVP.SVProgramID = SPP.SVProgramID WHERE SVP.Deleted = 0 AND SPP.AllowAnyCustomer = 1"
              Else
                MyCommon.QueryStr = "Select SVProgramID as ProgramID,Name as ProgramName from StoredValuePrograms with (NoLock) where deleted=0 order by Name"
              End If

            End If
            rst = MyCommon.LRT_Select
            For Each row In rst.Rows
              Send("<option value=" & MyCommon.NZ(row.Item("ProgramID"), 0) & ">" & MyCommon.NZ(row.Item("ProgramName"), "") & "</option>")
            Next
          %>
        </select>
        <br />
        <br class="half" />
        <input type="button" class="regular" id="select1" name="select1" value="&#9660; <% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID)) %>" onclick="handleSelectClick('select1');"<% sendb(disabledattribute) %> />&nbsp;
        <input type="button" class="regular" id="deselect1" name="deselect1" value="<% Sendb(Copient.PhraseLib.Lookup("term.deselect", LanguageID)) %> &#9650;" disabled="disabled" onclick="handleSelectClick('deselect1');"<% sendb(disabledattribute) %> /><br />
        <br class="half" />
        <select class="longer" id="selected" name="selected" size="2"<% sendb(disabledattribute) %>>
          <%
            If IncentiveSVID = 0 Then
              If Request.QueryString("selGroups") <> "" Then
                MyCommon.QueryStr = "select Name from StoredValuePrograms where SVProgramID=" & MyCommon.Extract_Val(Request.QueryString("selGroups"))
                rst = MyCommon.LRT_Select
                If rst.Rows.Count > 0 Then
                  Send("<option value=""" & MyCommon.Extract_Val(Request.QueryString("selGroups")) & """ >" & MyCommon.NZ(rst.Rows(0).Item("Name"), "") & "</option>")
                End If
              End If
            Else
              MyCommon.QueryStr = "Select IPG.SVProgramID as ProgramID,Name as ProgramName from CPE_IncentiveStoredValuePrograms as IPG with (NoLock) left join StoredValuePrograms as PP with (NoLock) on PP.SVProgramID=IPG.SVProgramID where IPG.deleted=0 and IncentiveStoredValueID=" & IncentiveSVID
              rst = MyCommon.LRT_Select
              For Each row In rst.Rows
                Send("<option value=""" & MyCommon.NZ(row.Item("ProgramID"), 0) & """>" & MyCommon.NZ(row.Item("ProgramName"), "") & "</option>")
              Next
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
        <label for="t1_QtyForIncentive">
          <% Sendb(Copient.PhraseLib.Lookup("condition.valueneeded", LanguageID))%>
        </label>
        <br />
        <%
          If TierLevels > 1 Then
            If IncentiveSVID = 0 Then
              For t = 1 To TierLevels
                If Request.QueryString("t" & t & "_QtyForIncentive") <> "" Then
                  Send("        <label for=""t" & t & "_QtyForIncentive"">" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</label>")
                            Send("        <input type=""text"" class=""short"" id=""t" & t & "_QtyForIncentive"" name=""t" & t & "_QtyForIncentive"" value=""" & MyCommon.Extract_Val(Request.QueryString("t" & t & "_QtyForIncentive")) & """" & DisabledAttribute & " maxlength=""6"" /><br />")
                Else
                  Send("        <label for=""t" & t & "_QtyForIncentive"">" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</label>")
                            Send("        <input type=""text"" class=""short"" id=""t" & t & "_QtyForIncentive"" name=""t" & t & "_QtyForIncentive"" value=""0""" & DisabledAttribute & " maxlength=""6"" /><br />")
                End If
              Next
            Else
              For t = 1 To TierLevels
                MyCommon.QueryStr = "Select Quantity from CPE_IncentiveStoredValueProgramTiers with (NoLock) where RewardOptionID=" & roid & " and TierLevel=" & t & " and IncentiveStoredValueID=" & IncentiveSVID
                TierDT = MyCommon.LRT_Select()
                If TierDT.Rows.Count > 0 Then
                  TierQty = MyCommon.NZ(TierDT.Rows(0).Item("Quantity"), 0)
                  Send("        <label for=""t" & t & "_QtyForIncentive"">" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</label>")
                            Send("        <input type=""text"" class=""short"" id=""t" & t & "_QtyForIncentive"" name=""t" & t & "_QtyForIncentive"" value=""" & TierQty & """" & DisabledAttribute & " maxlength=""6"" /><br />")
                Else
                  Send("        <label for=""t" & t & "_QtyForIncentive"">" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</label>")
                  Send("        <input type=""text"" class=""short"" id=""t" & t & "_QtyForIncentive"" name=""t" & t & "_QtyForIncentive"" value=""0"" " & DisabledAttribute & " maxlength=""6"" /><br />")
                End If
              Next
            End If
          Else
            If IncentiveSVID = 0 Then
              If Request.QueryString("t1_QtyForIncentive") <> "" Then
                Send("        <input type=""text"" class=""short"" id=""t1_QtyForIncentive"" name=""t1_QtyForIncentive"" value=""" & MyCommon.Extract_Val(Request.QueryString("t1_QtyForIncentive")) & """" & DisabledAttribute & " maxlength=""6"" /><br />")
              Else
                Send("        <input type=""text"" class=""short"" id=""t1_QtyForIncentive"" name=""t1_QtyForIncentive"" value=""""" & DisabledAttribute & " maxlength=""6"" /><br />")
              End If
            Else
              Dim QtyForIncentive As String = ""
              MyCommon.QueryStr = "Select QtyForIncentive from CPE_IncentiveStoredValuePrograms with (NoLock) where deleted=0 and IncentiveStoredValueID=" & IncentiveSVID
              rst = MyCommon.LRT_Select
              If (rst.Rows.Count > 0) Then
                QtyForIncentive = CInt(MyCommon.NZ(rst.Rows(0).Item("QtyForIncentive"), 0)).ToString
              Else
                QtyForIncentive = ""
              End If
              Send("        <input type=""text"" class=""short"" id=""t1_QtyForIncentive"" name=""t1_QtyForIncentive"" value=""" & QtyForIncentive & """" & DisabledAttribute & " maxlength=""6"" /><br />")
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
      ValidTier = True
    End If
    
    Return ValidTier
  End Function
</script>

<script type="text/javascript">
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
