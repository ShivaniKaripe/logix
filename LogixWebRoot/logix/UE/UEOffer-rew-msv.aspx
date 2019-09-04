<%@ Page Language="vb" Debug="true" CodeFile="ue-cb.vb" Inherits="UECB" EnableEventValidation="false" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: UEoffer-rew-msv.aspx 
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
  ' * Version : 5.99.1.84223 
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
  Dim OfferID As Long
  Dim RewardID As Long
  Dim ProgramID As Long
  Dim DeliverableID As Long
  Dim Phase As Integer
  Dim Name As String = ""
  Dim bError As Boolean = False
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
  Dim SVDesc As String = ""
  Dim BannersEnabled As Boolean = True
  Dim TierLevels As Integer = 1
  Dim t As Integer = 1
  Dim NumArray(99) As Decimal
  Dim RewardRequired As Boolean = True
  Dim MLI As New Copient.Localization.MultiLanguageRec
  Dim IsAnyCustomer As Boolean = False

  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If

  Response.Expires = 0
  MyCommon.AppName = "UEoffer-rew-msv.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  Localization = New Copient.Localization(MyCommon)

  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")

  OfferID = Request.QueryString("OfferID")
  'This would redirect to summary page if offer is awaiting deployment or if offer is awaiting recommendations.
  CheckIfValidOffer(MyCommon, OfferID)
  RewardID = MyCommon.Extract_Val(Request.QueryString("RewardID"))
  ProgramID = MyCommon.Extract_Val(Request.QueryString("ProgramID"))
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

  Dim bEnableAdditionalLockoutRestrictionsOnOffers As Boolean = IIf(MyCommon.Fetch_SystemOption(260) = "1", True, False)
  Dim bOfferEditable As Boolean = MyCommon.IsOfferEditablePastLockOutPeriod(Logix.UserRoles.EditOfferPastLockoutPeriod, MyCommon, OfferID)
  Dim Descfromspecial As String = ""
  Dim DescriptLength As Boolean = False
  Dim dtDesc As DataTable
  MyCommon.QueryStr = "Select Description from StoredValuePrograms with (NoLock) where SVProgramID =" & ProgramID
  dtDesc = MyCommon.LRT_Select
  If dtDesc.Rows.Count > 0 Then
    Descfromspecial = MyCommon.NZ(dtDesc.Rows(0)(0), "")
    MLI.StandardValue = Descfromspecial
  End If
  If Descfromspecial <> "" Then
    SVDesc = Descfromspecial.Replace("&quot;", Chr(34))
    If SVDesc.Length <= 1000 Then
      DescriptLength = True
    End If
  Else
    DescriptLength = True
  End If

  DeliverableID = MyCommon.Extract_Val(Request.QueryString("DeliverableID"))
  Phase = MyCommon.Extract_Val(Request.QueryString("phase"))
  If (Phase = 0) Then Phase = MyCommon.Extract_Val(Request.Form("Phase"))
  If (Phase = 0) Then Phase = 3


  ' Fetch the name
  MyCommon.QueryStr = "Select IncentiveName, IsTemplate, FromTemplate from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID
  rst = MyCommon.LRT_Select
  If (rst.Rows.Count > 0) Then
    Name = MyCommon.NZ(rst.Rows(0).Item("IncentiveName"), "")
    IsTemplate = MyCommon.NZ(rst.Rows(0).Item("IsTemplate"), False)
    FromTemplate = MyCommon.NZ(rst.Rows(0).Item("FromTemplate"), False)
  End If
  IsTemplateVal = IIf(IsTemplate, "IsTemplate", "Not")

  If (Request.QueryString("save") <> "") Then
    SVDesc = Request.QueryString("SVdesc")
    If SVDesc.Trim() = "" Then
      infoMessage = Copient.PhraseLib.Lookup("term.emptymsv", LanguageID)
    ElseIf (SVDesc.Length > 1000) Then
      infoMessage = Copient.PhraseLib.Lookup("CPEoffergen.description", LanguageID)
    End If
    If infoMessage = "" Then
      If (OfferID > 0 AndAlso RewardID > 0 AndAlso ProgramID > 0 AndAlso DeliverableID = 0) Then

        CreateROID = IIf(TpROID > 0, TpROID, RewardID)
        If (Create_Reward(OfferID, CreateROID, ProgramID, Phase, SVDesc, RewardRequired, DeliverableID)) Then
          infoMessage = Copient.PhraseLib.Lookup("term.ChangesWereSaved", LanguageID)
          DisableDeferCalcToEOS(OfferID)
          MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Reward.createmonstoredvalue", LanguageID))
        Else
          infoMessage = Copient.PhraseLib.Lookup("ueoffer-rew-points.ErrorOnSave", LanguageID)
          bError = True
        End If
      ElseIf (OfferID > 0 AndAlso RewardID > 0 AndAlso ProgramID > 0 AndAlso DeliverableID > 0) Then
        ' Modify existing deliverable
        MyCommon.QueryStr = "update CPE_DeliverableMonStoredValue with (RowLock) set SVProgramID=" & ProgramID & _
                            "where RewardOptionID=" & RewardID & " and DeliverableID=" & DeliverableID
        MyCommon.LRT_Execute()

        ' update the required status for this deliverable
        MyCommon.QueryStr = "update CPE_Deliverables with (RowLock) set Required=" & IIf(RewardRequired, 1, 0) & ", LastUpdate=getdate() " & _
                            " where DeliverableID=" & DeliverableID
        MyCommon.LRT_Execute()

        'Updating the description of stored value program
        MyCommon.QueryStr = "update StoredValuePrograms with (RowLock) set Description = @SVDesc where SVProgramID =@ProgramID"
        MyCommon.DBParameters.Add("@SVDesc", SqlDbType.NVarChar, 1000).Value = SVDesc
        MyCommon.DBParameters.Add("@ProgramID", SqlDbType.BigInt).Value = ProgramID
        MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
        
        DisableDeferCalcToEOS(OfferID)

        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Reward.updatemonstoredvalue", LanguageID))
      End If
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
      MyCommon.LRT_Execute()
            ResetOfferApprovalStatus(OfferID)
      'Update multi-language:
      'Description
      MLI.ItemID = ProgramID
      MLI.MLTableName = "CPE_DeliverableMonSVTranslations"
      MLI.MLColumnName = "Description"
      MLI.MLIdentifierName = "SVProgramID"
      MLI.StandardTableName = "StoredValuePrograms"
      MLI.StandardColumnName = "Description"
      MLI.StandardIdentifierName = "SVProgramId"
      MLI.InputName = "SVdesc"
      MLI.InputID = "SVdesc"
      MLI.InputType = "textarea"
      MLI.LabelPhrase = "term.custfacingdescription"
      MLI.MaxLength = 1000
      MLI.CSSClass = "longest"
      MLI.CSSStyle = "width:90%;"
      MLI.TRows = "4"
      Localization.SaveTranslationInputs(MyCommon, MLI, Request.QueryString, 9)
      CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
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

  Send_HeadBegin("term.offer", "term.monetarystoredvalue", OfferID)
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
  If EngineID = 9 Then

    MyCommon.QueryStr = "Select SVProgramID as ProgramID,Name as ProgramName from StoredValuePrograms with (NoLock) where deleted=0 and SVTypeID = 2 " & _
                         IIf(IsAnyCustomer, " and SVPRogramID in ( Select SVProgramID from SVProgramsPromoEngineSettings Where AllowAnyCustomer = 1)", "") & _
                         "and SVPRogramID Not in " & _
                         "(select SVProgramID from CPE_Deliverables D with (NoLock) inner join CPE_DeliverableMonStoredValue DSV with (NoLock) on DSV.RewardOptionID = D.RewardOptionID " & _
                         " inner Join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID = DSV.RewardOptionID " & _
                         " inner Join CPE_Incentives CI with (NoLock) on CI.IncentiveID = RO.IncentiveID and CI.Deleted = 0 " & _
                         "where D.Deleted=0 and D.DeliverableTypeID=16  and D.DeliverableID <> " & DeliverableID & " and DSV.Deleted=0) and SVExpireType <> 1 " & _
                         "union " & _
                         "Select SVProgramID as ProgramID,Name as ProgramName from StoredValuePrograms with (NoLock) where deleted=0 and SVTypeID = 2 " & _
                         IIf(IsAnyCustomer, " and SVPRogramID in ( Select SVProgramID from SVProgramsPromoEngineSettings Where AllowAnyCustomer = 1)", "") & _
                         "and SVPRogramID Not in " & _
                         "(select SVProgramID from CPE_Deliverables D with (NoLock) inner join CPE_DeliverableMonStoredValue DSV with (NoLock) on DSV.RewardOptionID = D.RewardOptionID " & _
                         " inner Join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID = DSV.RewardOptionID " & _
                         " inner Join CPE_Incentives CI with (NoLock) on CI.IncentiveID = RO.IncentiveID and CI.Deleted = 0 " & _
                         "where D.Deleted=0 and D.DeliverableTypeID=16  and D.DeliverableID <> " & DeliverableID & " and DSV.Deleted=0) and SVExpireType = 1 and ExpireDate is not null and ExpireDate > GETDATE() order by Name; "

  End If
  rst = MyCommon.LRT_Select
  If (rst.Rows.Count > 0) Then
   
    Sendb("var functionlist = Array(")
    For Each row In rst.Rows
      Sendb("""" & MyCommon.NZ(row.Item("ProgramName"), "").ToString().Replace("""", "\""") & """,")
    Next
    Sendb(""""");")
    Sendb("var vallist = Array(")
    For Each row In rst.Rows
      Sendb("""" & MyCommon.NZ(row.Item("ProgramID"), 0) & """,")
    Next
    Sendb(""""");")
  Else
    infoMessage = Copient.PhraseLib.Lookup("sv-list.notavailable", LanguageID)
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
    var i, numShown;
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
    if (document.forms[0].functionradio[0].checked == true) {
      searchPattern = "^" + textObj.value;
    } else {
      searchPattern = textObj.value;
    }
    searchPattern = cleanRegExpString(searchPattern);

    // Create a regular expression
    re = new RegExp(searchPattern, "gi");
    // Clear the options list
    selectObj.length = 0;

    // Loop through the array and re-add matching options
    numShown = 0;
    for (i = 0; i < functionListLength; i++) {
      if (functionlist[i].search(re) != -1) {
        if (vallist[i] != "" && (selectedList.options.length < 1 || vallist[i] != selectedList.options[0].value)) {
          selectObj[numShown] = new Option(functionlist[i], vallist[i]);
          numShown++;
        }
      }
      // Stop when the number to show is reached
      if (numShown == maxNumToShow) {
        break;
      }
    }
    // When options list whittled to one, select that entry
    if (selectObj.length == 1) {
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

    if (selectedValue != "") {
    }
  }

  function addToSelect() {
    var elemSource = document.getElementById("functionselect");
    var elemDest = document.getElementById("selected");
    var selOption = null;
    var selText = "", selVal = "";
    var selIndex = -1;

    if (elemSource != null && elemSource.options.selectedIndex == -1) {
      alert('<% Sendb(Copient.PhraseLib.Lookup("CPE-rew-membership.selectpoints", LanguageID))%>');
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
    if (elemDest.options[0] != null)
      GetSVProgDescription();
    handleKeyUp(99999);
  }
}

function removeFromSelect() {
  var elem = document.getElementById("selected");
  var elemList = document.getElementById("functionselect");


  if (elem != null && elem.options.length > 0) {
    elemList.options[elemList.options.length] = new Option(elem.options[0].text, elem.options[0].value);
    elem.options[0] = null;
    handleKeyUp(99999);
  }

  //  showMultiLanguageInput('SVdesc', event, true);
  if (document.getElementById("mlwrap_SVdesc") != null)
    $("div#mlwrap_SVdesc textarea").val("");
}

//Ajax method to get the stored value programs for the country in drop down change event
function GetSVProgDescription() {
  xmlhttpPost('../OfferFeeds.aspx', 'GetSVProgDescription');
}

function xmlhttpPost(strURL, mode) {
  var xmlHttpReq = false;
  var self = this;
  if (window.XMLHttpRequest) {
    self.xmlHttpReq = new XMLHttpRequest();
  }
    // IE
  else if (window.ActiveXObject) {
    self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
  }
  var qryStr = getprogramquery(mode);
  self.xmlHttpReq.open('POST', strURL + "?" + qryStr, true);
  self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
  self.xmlHttpReq.onreadystatechange = function () {
    if (self.xmlHttpReq != null && self.xmlHttpReq.readyState == 4) {
      FillSVDesp(self.xmlHttpReq.responseText);
    }
  }
  self.xmlHttpReq.send(qryStr);
}

function getprogramquery(mode) {
  return "Mode=" + mode + "&svProgID=" + document.getElementById('selected').options[0].value;
}

// Callback method that gets the values in the response text
function FillSVDesp(str) {
  var resArrLan = new Array();
  var resArrDesc = new Array();
  if (str.length > 0) {
    str = str.trim();
    if (str.indexOf("~") != -1) {
      var responseArray = new Array();
      responseArray = str.split('~');
      for (var i = 0; i < responseArray.length - 1; i++) {
        var tempresArray = new Array();
        tempresArray = responseArray[i].split('|');
        resArrLan[i] = tempresArray[0];
        resArrDesc[i] = tempresArray[1];
      }
      var defaultLanDesc = "";
      for (var j = 0; j < resArrLan.length; j++) {
        defaultLanDesc = resArrDesc[0];
        document.getElementById('SVdesc_' + resArrLan[j]).value = resArrDesc[j];

      }
      $("div#mlwrap_SVdesc textarea#SVdesc").val(defaultLanDesc);
    }
    else {
      if (document.getElementById("ml_SVdesc") != null) {
        $("div#mlwrap_SVdesc textarea#SVdesc").val(str);
        var defaultInput = document.getElementById('ml_' + 'SVdesc' + '_default').firstChild;
        defaultInput.value = str;
      }
      else {
        $("div#mlwrap_SVdesc textarea#SVdesc").val(str);
      }
    }

  }
  else {
    if (document.getElementById("ml_SVdesc") != null) {
      $("div#mlwrap_SVdesc textarea#SVdesc").val("");
      var defaultInput = document.getElementById('ml_' + 'SVdesc' + '_default').firstChild;
      defaultInput.value = "";
    }
    else {
      $("div#mlwrap_SVdesc textarea#SVdesc").val("");
    }
  }
}


function validateEntry() {
  var retVal = true;
  var elem = document.getElementById("selected");
  var elemProgram = document.getElementById("ProgramID");
  var msg = '';

  if (elem != null && elem.options.length == 0) {
    retVal = false;
    msg += '* <% Sendb(Copient.PhraseLib.Lookup("CPE-rew-membership.selectpoints", LanguageID))%>'
    elem.focus();
  } else if (elem != null && elemProgram != null) {
    elemProgram.value = elem.options[0].value;
  }
  if (msg != '') {
    alert(msg);
  }
  return retVal;
}

function selectByPointsID(pointsID) {
  var elemSelect = document.getElementById("functionselect");
  var selectLength = elemSelect.length;

  for (i = 0; i < selectLength; i++) {
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
    e.returnValue = false;
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

</script>
<%
  Send("<script type=""text/javascript"">")
  Send("function ChangeParentDocument() { ")
  If (EngineID = 9) Then
    Send("  if (opener != null) {")
    Send("    var newlocation = 'UEoffer-rew.aspx?OfferID=" & OfferID & "'; ")
    Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
    Send("  opener.location = 'UEoffer-rew.aspx?OfferID=" & OfferID & "'; ")
    Send("  }")
    Send("  }")
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
<form action="UEoffer-rew-msv.aspx" id="mainform" name="mainform" onsubmit="return validateEntry();">
  <div id="intro">
    <input type="hidden" id="OfferID" name="OfferID" value="<% Sendb(OfferID)%>" />
    <input type="hidden" id="Name" name="Name" value="<% Sendb(Name)%>" />
    <input type="hidden" id="RewardID" name="RewardID" value="<% Sendb(RewardID)%>" />
    <input type="hidden" id="DeliverableID" name="DeliverableID" value="<% Sendb(DeliverableID)%>" />
    <input type="hidden" id="Phase" name="Phase" value="<% Sendb(Phase)%>" />
    <input type="hidden" id="ProgramID" name="ProgramID" value="<% Sendb(ProgramID)%>" />
    <input type="hidden" id="roid" name="roid" value="<%Sendb(TpROID)%>" />
    <input type="hidden" id="IsTemplate" name="IsTemplate" value="<% Sendb(IsTemplateVal)%>" />
    <input type="hidden" id="EngineID" name="EngineID" value="<%Sendb(EngineID)%>" />
    <%
      If (IsTemplate) Then
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.monetarystoredvalue", LanguageID), VbStrConv.Lowercase) & "</h1>")
      Else
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.monetarystoredvalue", LanguageID), VbStrConv.Lowercase) & "</h1>")
      End If
    %>
    <div id="controls">
      <% If (IsTemplate) Then%>
      <span class="temp">
        <input type="checkbox" class="tempcheck" id="Disallow_Edit" name="Disallow_Edit" <% If (Disallow_Edit) Then Sendb(" checked=""checked""")%> />
        <label for="Disallow_Edit"><% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
      </span>
      <% End If%>
      <% 
        If ((Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
          If Not IsTemplate Then
            If (Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Edit)) Then Send_Save()
          Else
            If (Logix.UserRoles.EditTemplates) Then Send_Save()
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
    %>
    <div id="column1">
      <div class="box" id="selector">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.storedvalueprogram", LanguageID))%>
          </span>
        </h2>
        <input type="radio" id="functionradio1" name="functionradio" <% If (MyCommon.Fetch_SystemOption(175) = "1") Then Sendb(" checked=""checked""")%> <% Sendb(DisabledAttribute)%> /><label for="functionradio1"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
        <input type="radio" id="functionradio2" name="functionradio" <% If (MyCommon.Fetch_SystemOption(175) = "2") Then Sendb(" checked=""checked""")%> /><label for="functionradio2" <% Sendb(DisabledAttribute)%>><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
        <input type="text" class="medium" id="functioninput" name="functioninput" maxlength="1000" onkeydown="handleKeyDown(event);" onkeyup="handleKeyUp(200);" value="" <% Sendb(DisabledAttribute)%> /><br />
        <select class="longer" id="functionselect" name="functionselect" onkeydown="return handleSlctKeyDown(event);" onclick="handleSelectClick();" ondblclick="addToSelect();" size="16" <% Sendb(DisabledAttribute)%>>
          <%
            If UEOffer_Has_AnyCustomer(MyCommon, OfferID) Then
              IsAnyCustomer = True
            End If
            
            If EngineID = 9 Then
                
              MyCommon.QueryStr = "Select SVProgramID as ProgramID,Name as ProgramName from StoredValuePrograms with (NoLock) where deleted=0 and SVTypeID = 2 " & _
                                   IIf(IsAnyCustomer, " and SVPRogramID in ( Select SVProgramID from SVProgramsPromoEngineSettings Where AllowAnyCustomer = 1)", "") & _
                                   "and SVPRogramID Not in " & _
                                   "(select SVProgramID from CPE_Deliverables D with (NoLock) inner join CPE_DeliverableMonStoredValue DSV with (NoLock) on DSV.RewardOptionID = D.RewardOptionID " & _
                                   " inner Join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID = DSV.RewardOptionID " & _
                                   " inner Join CPE_Incentives CI with (NoLock) on CI.IncentiveID = RO.IncentiveID and CI.Deleted = 0 " & _
                                   "where D.Deleted=0 and D.DeliverableTypeID=16  and D.DeliverableID <> " & DeliverableID & " and DSV.Deleted=0) and SVExpireType <> 1 " & _
                                   "union " & _
                                   "Select SVProgramID as ProgramID,Name as ProgramName from StoredValuePrograms with (NoLock) where deleted=0 and SVTypeID = 2 " & _
                                   IIf(IsAnyCustomer, " and SVPRogramID in ( Select SVProgramID from SVProgramsPromoEngineSettings Where AllowAnyCustomer = 1)", "") & _
                                   "and SVPRogramID Not in " & _
                                   "(select SVProgramID from CPE_Deliverables D with (NoLock) inner join CPE_DeliverableMonStoredValue DSV with (NoLock) on DSV.RewardOptionID = D.RewardOptionID " & _
                                   " inner Join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID = DSV.RewardOptionID " & _
                                   " inner Join CPE_Incentives CI with (NoLock) on CI.IncentiveID = RO.IncentiveID and CI.Deleted = 0 " & _
                                   "where D.Deleted=0 and D.DeliverableTypeID=16  and D.DeliverableID <> " & DeliverableID & " and DSV.Deleted=0) and SVExpireType = 1 and ExpireDate is not null and ExpireDate > GETDATE() order by Name; "
                                       
            End If
            rst = MyCommon.LRT_Select
            If rst.Rows.Count = 0 Then
              infoMessage = Copient.PhraseLib.Lookup("sv-list.notavailable", LanguageID)
              Send("<script type=""text/javascript"" language=""javascript"">")
              Send("  document.getElementById(""save"").disabled = 'disabled' ")
              Send("</script>")
            Else
              Send("<script type=""text/javascript"" language=""javascript"">")
              Send("  if (document.getElementById(""save"") != null ) {")
              Send(" document.getElementById(""save"").disabled = '' ; }")
              Send("</script>")
              For Each row In rst.Rows
                Send("<option value=""" & MyCommon.NZ(row.Item("ProgramID"), 0) & """>" & MyCommon.NZ(row.Item("ProgramName"), "") & "</option>")
              Next
            End If
          %>
        </select>
        <br />
        <br class="half" />
        <input type="button" class="regular" id="select" name="select" value="&#9660; <% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID))%>" onclick="addToSelect();" <% Sendb(DisabledAttribute)%> />&nbsp;
        <input type="button" class="regular" id="deselect" name="deselect" value="<% Sendb(Copient.PhraseLib.Lookup("term.deselect", LanguageID))%> &#9650;" onclick="removeFromSelect();" <% Sendb(DisabledAttribute)%> /><br />
        <br class="half" />
        <select class="longer" id="selected" name="selected" size="2" ondblclick="removeFromSelect();" <% Sendb(DisabledAttribute)%>>
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
            <% Sendb(Copient.PhraseLib.Lookup("term.details", LanguageID))%>
          </span>
        </h2>
        <br />
        <%   
          'Description
          MLI.ItemID = ProgramID
          MLI.MLTableName = "CPE_DeliverableMonSVTranslations"
          MLI.MLColumnName = "Description"
          MLI.MLIdentifierName = "SVProgramID"
          MLI.StandardTableName = "StoredValuePrograms"
          MLI.StandardColumnName = "Description"
          MLI.StandardIdentifierName = "SVProgramID"
          MLI.StandardValue = SVDesc
          MLI.InputName = "SVdesc"
          MLI.InputID = "SVdesc"
          MLI.InputType = "textarea"
          MLI.LabelPhrase = "term.custfacingdescription"
          MLI.MaxLength = 1000
          MLI.CSSClass = "longest"
          MLI.CSSStyle = "width:90%;"
          MLI.TRows = "4"
             
          Send(Localization.SendTranslationInputs(MyCommon, MLI, Request.QueryString, 9))
        %>
        <small>
          <%Send("(" & Copient.PhraseLib.Lookup("CPEoffergen.description", LanguageID) & ")")%></small><br
            class="half" />
        <br class="half" />
        <hr class="hidden" />
      </div>
    </div>
  </div>

  <script runat="server">
    Function Create_Reward(ByVal OfferID As Long, ByVal ROID As Long, ByVal ProgramID As Long, ByVal Phase As Long, _
                         ByVal SVDesc As String, ByVal RewardRequired As Boolean, ByRef DeliverableID As Long) As Boolean
      Dim MyCommon As New Copient.CommonInc
      Dim Status As Integer = 0
      
      Try
        MyCommon.QueryStr = "dbo.pa_CPE_AddMonStoredValueReward"
        MyCommon.Open_LogixRT()
        MyCommon.Open_LRTsp()
        
        MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.Int, 4).Value = OfferID
        MyCommon.LRTsp.Parameters.Add("@ROID", SqlDbType.Int, 4).Value = ROID
        MyCommon.LRTsp.Parameters.Add("@ProgramID", SqlDbType.Int, 4).Value = ProgramID
        MyCommon.LRTsp.Parameters.Add("@Phase", SqlDbType.Int, 4).Value = Phase
        MyCommon.LRTsp.Parameters.Add("@MSVdesc", SqlDbType.NVarChar, 1000).Value = SVDesc
        MyCommon.LRTsp.Parameters.Add("@Required", SqlDbType.Bit).Value = IIf(RewardRequired, 1, 0)
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
    
    Sub DisableDeferCalcToEOS(ByVal OfferID As Long)
      Dim MyCommon As New Copient.CommonInc
      Dim Status As Integer = 0
      Dim rst As New DataTable()
      
      Try
        MyCommon.Open_LogixRT()
        
        ' monetary stored value disable the DeferCalcToEOS option of an offer
        MyCommon.QueryStr = "Select DeferCalcToEOS from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID & " and Deleted=0 and DeferCalcToEOS=1;"
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
          MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set DeferCalcToEOS=0 where IncentiveID=" & OfferID & ";"
          MyCommon.LRT_Execute()
        End If
        
      Catch ex As Exception
      Finally
        MyCommon.Close_LogixRT()
        MyCommon = Nothing
      End Try
      
    End Sub
   
  </script>
  <script type="text/javascript">
  <% If (CloseAfterSave) Then%>
    window.close();
    <% End If%>
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

