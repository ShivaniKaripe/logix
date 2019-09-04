<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: offer-con-SV.aspx 
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
  Dim rst As DataTable
  Dim row As DataRow
  Dim OfferID As Long
  Dim Name As String = ""
  Dim ConditionID As String
  Dim Tiered As Boolean
  Dim NumTiers As Integer
  Dim Disallow_Edit As Boolean = True
  Dim bUseTemplateLocks As Boolean
  Dim IsTemplate As Boolean = False
  Dim CloseAfterSave As Boolean = False
  Dim DisabledAttribute As String = ""
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim RequirePP As Boolean = False
  Dim ProgramID As Integer = 0
  Dim BannersEnabled As Boolean = False
  Dim bDisallowEditValue As Boolean = False
  Dim bDisallowEditRewards As Boolean = False
  Dim bDisallowEditPp As Boolean = False
  Dim iRadioValue As Int16
  Dim sDisabled As String
  
  Dim objTemp As Object
  Dim decTemp As Decimal
  Dim intNumDecimalPlaces As Integer
  Dim decFactor As Decimal
  Dim sTemp As String
  Dim SVTypeID As Integer
  Dim bNeedToFormat As Boolean
  Dim RECORD_LIMIT As Integer = GroupRecordLimit '500
  Dim bCreateGroupOrProgramFromOffer As Boolean = IIf(MyCommon.Fetch_CM_SystemOption(134) ="1",True,False)
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "offer-con-SV.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  Tiered = False
  OfferID = Request.QueryString("OfferID")
  ConditionID = Request.QueryString("ConditionID")
  NumTiers = Request.QueryString("NumTiers")
  ProgramID = MyCommon.Extract_Val(Request.QueryString("ProgramID"))
  SVTypeID = MyCommon.Extract_Val(Request.QueryString("SVTypeID"))
  
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")

  objTemp = MyCommon.Fetch_CM_SystemOption(41)
  If Not (Integer.TryParse(objTemp.ToString, intNumDecimalPlaces)) Then
    intNumDecimalPlaces = 0
  End If
  decFactor = (10 ^ intNumDecimalPlaces)
  
  MyCommon.QueryStr = "Select Name,IsTemplate,FromTemplate from Offers with (NoLock) where OfferID=" & OfferID
  rst = MyCommon.LRT_Select
  If (rst.Rows.Count > 0) Then
    Name = rst.Rows(0).Item("Name")
    IsTemplate = MyCommon.NZ(rst.Rows(0).Item("IsTemplate"), False)
    If IsTemplate Then
      bUseTemplateLocks = False
    Else
      bUseTemplateLocks = MyCommon.NZ(rst.Rows(0).Item("FromTemplate"), False)
    End If
  End If
  
  If (IsTemplate Or bUseTemplateLocks) Then
    ' lets dig the permissions if its a template
    MyCommon.QueryStr = "select Tiered,Disallow_Edit,RequiredFromTemplate,DisallowEdit1,DisallowEdit2,DisallowEdit3 from OfferConditions with (NoLock) where ConditionID=" & ConditionID & ";"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      Disallow_Edit = MyCommon.NZ(rst.Rows(0).Item("Disallow_Edit"), True)
      RequirePP = MyCommon.NZ(rst.Rows(0).Item("RequiredFromTemplate"), False)
      bDisallowEditPp = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit1"), False)
      bDisallowEditValue = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit2"), False)
      bDisallowEditRewards = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit3"), False)
      Tiered = MyCommon.NZ(rst.Rows(0).Item("Tiered"), False)
      If Tiered Then
        bDisallowEditRewards = True
      End If
      If bUseTemplateLocks Then
        If Disallow_Edit Then
          bDisallowEditPp = True
          bDisallowEditValue = True
          bDisallowEditRewards = True
        Else
          Disallow_Edit = bDisallowEditPp And bDisallowEditValue And bDisallowEditRewards
        End If
      End If
    End If
  End If
  
  Send_HeadBegin("term.offer", "term.storedvaluecondition", OfferID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
%>

<script type="text/javascript" language="javascript">
var fullSelect = null;
var isFireFox = (navigator.appName.indexOf('Mozilla') !=-1) ? true : false;

/**************************************************************************************************************************/
//Script to call server method through JavaScript
//to load stored value programs based on search criteria.
var timer;
function xmlPostTimer(strURL,mode)
{
  clearTimeout(timer);
  timer=setTimeout("xmlhttpPost('" + strURL + "','" + mode + "')", 250);
}

function xmlhttpPost(strURL,mode) {
  var xmlHttpReq = false;
  var self = this;
  document.getElementById("searchLoadDiv").innerHTML = '<% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID)) %>';
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
  self.xmlHttpReq.onreadystatechange = function() {
    if (self.xmlHttpReq !=null && self.xmlHttpReq.readyState == 4) {
      updatepage(self.xmlHttpReq.responseText);
    }
    }
    self.xmlHttpReq.send(qryStr);
  }

function getprogramquery(mode) {
  var radioString;
  if(document.getElementById('functionradio2').checked) {
    radioString = 'functionradio2';
  }
  else {
    radioString = 'functionradio1';
  }
  var selected = document.getElementById('selected');
  var selectedProgram = 0;
  if(selected.options[0] != null){
    selectedProgram = selected.options[0].value;
  }
  return "Mode=" + mode + "&ProgramSearch=" + document.getElementById('functioninput').value + "&SelectedProgram=" + selectedProgram + "&SearchRadio=" + radioString;
 
}

function updatepage(str) {
  if(str.length > 0){
    if(!isFireFox){
      document.getElementById("List").innerHTML = '<select class="longer" id="functionselect" name="functionselect" size="12"<% sendb(disabledattribute) %>>' + str + '</select>';
    }
    else{
      document.getElementById("functionselect").innerHTML = str;
    }
    document.getElementById("searchLoadDiv").innerHTML = '&nbsp;';
    if (document.getElementById("functionselect").options.length > 0) {
      document.getElementById("functionselect").options[0].selected = true;
    }
  }
  else if(str.length == 0){
    if(!isFireFox){
      document.getElementById("List").innerHTML = '<select class="longer" id="functionselect" name="functionselect" size="12"<% sendb(disabledattribute) %>>&nbsp;</select>';
    }
    else{
      document.getElementById("functionselect").innerHTML = '&nbsp;';
    }
    document.getElementById("searchLoadDiv").innerHTML = '&nbsp;';
  }
}

/***************************************************************************************************************************/

// this function gets the selected value and loads the appropriate
// php reference page in the display frame
// it can be modified to perform whatever action is needed, or nothing
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
    // alert(htmlContents);
    return true;
}

function validateEntry() {
    var retVal = true;
    var elemPP = document.getElementById("require_pp");
    var elem = document.getElementById("selected");   
    var qtyElem = document.getElementById("QtyForIncentive");
    var elemProgram = document.getElementById("ProgramID");
    var msg = '';
    
    if (elemPP == null || !elemPP.checked) {
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
          if (qtyVal == "" || isNaN(qtyVal)) {
              retVal = false;
              if (msg != '') { msg += '\n\r\n\r'; }
              msg += '* <% Sendb(Copient.PhraseLib.Lookup("CPE-rew-membership.enterquantity", LanguageID)) %>';
              qtyElem.focus();
              qtyElem.select();
            }
        }
    }
    if (msg != '') {
        alert(msg);
    }
    return retVal;
}

function updateButtons(){
  var elemDisallowEditPpOpt = document.getElementById("DisallowEditPpOpt");
  var selectObj = document.getElementById('selected');
  
  if (elemDisallowEditPpOpt != null && elemDisallowEditPpOpt.value == '1') {
      document.getElementById('select1').disabled=true;
      document.getElementById('deselect1').disabled=true;
      if(document.getElementById("btncreate")!=undefined && document.getElementById("btncreate")!=null){
        document.getElementById("btncreate").disabled=true;
      }
  } else {
    if (selectObj.length == 0) {
      document.getElementById('select1').disabled=false;
      document.getElementById('deselect1').disabled=true;
    } else {
      document.getElementById('select1').disabled=false;
      document.getElementById('deselect1').disabled=false;
    }
     if(document.getElementById("btncreate")!=undefined && document.getElementById("btncreate")!=null){
        document.getElementById("btncreate").disabled=false;
      }
  }
}

function removeUsed() {
  xmlhttpPost('OfferFeeds.aspx', 'StoredValueProgramsCM');
  // this function will remove items from the functionselect box that are used in 
  // selected and excluded boxes

  var funcSel = document.getElementById('functionselect');
  var elSel = document.getElementById('selected');
  var i, j;

  for (i = elSel.length - 1; i >= 0; i--) {
    for (j = funcSel.length - 1; j >= 0; j--) {
      if (funcSel.options[j].value == elSel.options[i].value) {
        document.getElementById("SVTypeID").value = '';
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
  if (selectedValue != "") { selectedText = selectObj[document.getElementById("functionselect").selectedIndex].text; }

  selectboxObj = document.forms[0].selected;
  selectedboxValue = document.getElementById("selected").value;
  if (selectedboxValue != "") { selectedboxText = selectboxObj[document.getElementById("selected").selectedIndex].text; }

  if (itemSelected == "select1") {
    if (selectedValue != "") {
      // empty the select box
      for (i = selectboxObj.length - 1; i >= 0; i--) {
        selectboxObj.options[i] = null;
      }
      // The SVType corresponds to the same position as the selected program
      document.getElementById("SVTypeID").value = document.getElementById("functionselect").value;

      // add items to selected box
      selectboxObj[selectboxObj.length] = new Option(selectedText, selectedValue);
    }
  }

  if (itemSelected == "deselect1") {
    if (selectedboxValue != "") {
      // remove items from selected box
      document.getElementById("selected").remove(document.getElementById("selected").selectedIndex)
      document.getElementById("SVTypeID").value = '';
    }
  }

  updateButtons();
  // remove items from large list that are in the other lists
  removeUsed();
  return true;
}

 function handleCreateClick(createbtn)
    {
    
        var alertMessage="";
        if(document.getElementById(createbtn)!= undefined && document.getElementById(createbtn) != null)
        {
            var searchText= document.getElementById('functioninput').value;
            if(searchText != null && searchText!="")
            {
                xmlhttpPost_CreateGroupOrProgramFromOffer('OfferFeeds.aspx', 'CreateGroupOrProgramFromOffer');
            }
            else
            {
                alertMessage='<% Sendb(Copient.PhraseLib.Lookup("term.enter", LanguageID))%>' +' ' + '<% Sendb(Copient.PhraseLib.Lookup("term.valid", LanguageID).ToLower())%>'+' ' + '<% Sendb(Copient.PhraseLib.Lookup("term.storedvalueprogram", LanguageID).ToLower())%>'+' ' + '<% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID).ToLower())%>';
                alert(alertMessage);
                document.getElementById('functioninput').focus();
                return false;
            }
        }
        return true;
    }

 function xmlhttpPost_CreateGroupOrProgramFromOffer(strURL,mode) {
      var xmlHttpReq = false;
      var self = this;
      document.getElementById("searchLoadDiv").innerHTML = '<% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID)) %>';
      if (window.XMLHttpRequest) {
        self.xmlHttpReq = new XMLHttpRequest();
      }
      // IE
      else if (window.ActiveXObject) {
        self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
      }
      var qryStr = getcreatequery(mode);
      self.xmlHttpReq.open('POST', strURL + "?" + qryStr, true);
      self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
      self.xmlHttpReq.onreadystatechange = function() {
        if (self.xmlHttpReq !=null && self.xmlHttpReq.readyState == 4) {
          updatepage_creategroupOrprogramfromoffer(self.xmlHttpReq.responseText);
        }
      }

      self.xmlHttpReq.send(qryStr);
    }
    function getcreatequery(mode)
    {
        return "Mode=" + mode + "&CreateType=StoredValue&Name=" + document.getElementById('functioninput').value
    }
    function updatepage_creategroupOrprogramfromoffer(str) 
    {
       if(str.length > 0)
      {
          var status ="";
          var responseArr = str.split('~');
          if(responseArr.length >0)
          {
            status=responseArr[0];
            if(status=="Ok")
            {
                var resultArr=responseArr[1].split('|');
                 if(document.getElementById('selected').options[0] !=undefined && document.getElementById('selected').options[0]!=null){
                    document.getElementById('selected').options[0].selected = 'selected';
                }
                handleSelectClick('deselect1');
                addNewGrouptoSelectbox(resultArr[0],resultArr[1]);
                document.getElementById('functionselect').value=resultArr[1];
                handleSelectClick('select1') ;
            }
            else if(status =="Fail")
            {
                var resultArr=responseArr[1].split('|');
                var selectedGroupValue= -1;
                if(document.getElementById('selected').options[0] !=undefined && document.getElementById('selected').options[0]!=null)
                    selectedGroupValue= document.getElementById('selected').options[0].value
                if(parseInt(selectedGroupValue) != parseInt(resultArr[1] ))
                {
                    alert(responseArr[2]);
                    if(selectedGroupValue != -1){
                        document.getElementById('selected').options[0].selected = 'selected';
                    }
                    handleSelectClick('deselect1');
                    document.getElementById('functionselect').value=resultArr[1];
                    handleSelectClick('select1') ;
                }
                else
                {
                    alert('<% Sendb(Copient.PhraseLib.Lookup("term.storedvalueprogram", LanguageID))%>' + ": '"+ resultArr[0] + "' " + '<% Sendb(Copient.PhraseLib.Lookup("term.is", LanguageID).ToLower())%>'+'  '+'<% Sendb(Copient.PhraseLib.Lookup("term.already", LanguageID).ToLower())%>'+'  '+'<% Sendb(Copient.PhraseLib.Lookup("term.selected", LanguageID).ToLower())%>');
                }
            }
            else if(status =="Error")
            {
                alert(responseArr[1]);
                return false;
            }
         }
      }
       document.getElementById("searchLoadDiv").innerHTML = '&nbsp;';
   }

   function addNewGrouptoSelectbox(text,val)
   {
        var sel = document.getElementById('functionselect');
        var opt = document.createElement('option'); // create new option element
        // create text node to add to option element (opt)
        opt.appendChild( document.createTextNode(text) );
        opt.value = val; // set value property of opt
        sel.appendChild(opt); // add opt to end of select box (sel)
   }
</script>

<%
  Send_Scripts()
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
  
  If (Request.QueryString("save") <> "") Then
    
    If Not (bUseTemplateLocks And bDisallowEditPp) Then
      If (ProgramID <> 0) Then
        MyCommon.QueryStr = "update OfferConditions with (RowLock) set LinkID=" & ProgramID & " where ConditionID=" & ConditionID & ";"
        MyCommon.LRT_Execute()
      End If
    End If
    
    If Not (bUseTemplateLocks And bDisallowEditRewards) Then
      If (Request.QueryString("granted") <> "") Then
        MyCommon.QueryStr = "update OfferConditions with (RowLock) set GrantTypeID=" & Request.QueryString("granted") & " where ConditionID=" & ConditionID & ";"
        MyCommon.LRT_Execute()
      End If
    End If

    If Not (bUseTemplateLocks And bDisallowEditValue) Then
      If (Request.QueryString("radioValue") <> "") Then
        iRadioValue = Int16.Parse(Request.QueryString("radioValue"))
        If iRadioValue = 4 Then
          MyCommon.QueryStr = "update OfferConditions with (RowLock) set PointsRedeemInstant=1,PointsUseEarnedValue=0,PointsUseNetValue=0 where ConditionID=" & ConditionID & ";"
        ElseIf iRadioValue = 2 Then
          MyCommon.QueryStr = "update OfferConditions with (RowLock) set PointsRedeemInstant=0,PointsUseEarnedValue=1,PointsUseNetValue=0 where ConditionID=" & ConditionID & ";"
        ElseIf iRadioValue = 1 Then
          MyCommon.QueryStr = "update OfferConditions with (RowLock) set PointsRedeemInstant=0,PointsUseEarnedValue=0,PointsUseNetValue=1 where ConditionID=" & ConditionID & ";"
        Else
          MyCommon.QueryStr = "update OfferConditions with (RowLock) set PointsRedeemInstant=0,PointsUseEarnedValue=0,PointsUseNetValue=0 where ConditionID=" & ConditionID & ";"
        End If
        MyCommon.LRT_Execute()
      End If

      If (Request.QueryString("valuetype") <> "") Then
        MyCommon.QueryStr = "update OfferConditions with (RowLock) set QtyUnitType=5 where ConditionID=" & ConditionID & ";"
        MyCommon.LRT_Execute()
      End If

      If SVTypeID = 1 And intNumDecimalPlaces > 0 Then
        bNeedToFormat = True
      Else
        bNeedToFormat = False
      End If

      If (Request.QueryString("tier0") <> "" And (Request.QueryString("Tiered") = "False")) Then
        If bNeedToFormat Then
          decTemp = MyCommon.Extract_Val(Request.QueryString("tier0")) * decFactor
        Else
          decTemp = MyCommon.Extract_Val(Request.QueryString("tier0"))
        End If
        decTemp = Int(decTemp + 0.5)
        MyCommon.QueryStr = "dbo.pt_ConditionTiers_Update"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@ConditionID", SqlDbType.BigInt).Value = ConditionID
        MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int).Value = 0
        MyCommon.LRTsp.Parameters.Add("@AmtRequired", SqlDbType.Decimal).Value = decTemp
        MyCommon.LRTsp.ExecuteNonQuery()
        MyCommon.Close_LRTsp()
      ElseIf (Request.QueryString("Tiered") = "True") Then
        ' delete the current tier ammounts
        MyCommon.QueryStr = "delete from ConditionTiers with (RowLock) where ConditionID=" & ConditionID
        MyCommon.LRT_Execute()
        Dim x As Integer
        Dim sTier As String
        For x = 1 To NumTiers
          sTier = "tier" & x
          If bNeedToFormat Then
            decTemp = MyCommon.Extract_Val(Request.QueryString(sTier)) * decFactor
          Else
            decTemp = MyCommon.Extract_Val(Request.QueryString(sTier))
          End If
          decTemp = Int(decTemp + 0.5)
          MyCommon.QueryStr = "dbo.pt_ConditionTiers_Update"
          MyCommon.Open_LRTsp()
          MyCommon.LRTsp.Parameters.Add("@ConditionID", SqlDbType.BigInt).Value = MyCommon.Extract_Val(ConditionID)
          MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int).Value = x
          MyCommon.LRTsp.Parameters.Add("@AmtRequired", SqlDbType.Decimal).Value = decTemp
          If (x > 1) And (Int(MyCommon.Extract_Val(Request.QueryString("tier" & x))) < Int(MyCommon.Extract_Val(Request.QueryString("tier" & (x - 1))))) Then
            infoMessage = Copient.PhraseLib.Lookup("condition.tiervalues", LanguageID)
          End If
          MyCommon.LRTsp.ExecuteNonQuery()
          MyCommon.Close_LRTsp()
        Next
      End If

    End If

    If (Request.QueryString("IsTemplate") = "IsTemplate") Then
      ' time to update the status bits for the templates
      Dim form_Disallow_Edit As Integer = 0
      Dim form_require_pp As Integer = 0
      Dim iDisallowEditPp As Integer = 0
      Dim iDisallowEditValue As Integer = 0
      Dim iDisallowEditRewards As Integer = 0
      
      Disallow_Edit = False
      RequirePP = False
      bDisallowEditValue = False
      bDisallowEditRewards = False
      bDisallowEditPp = False
      
      If (Request.QueryString("Disallow_Edit") = "on") Then
        form_Disallow_Edit = 1
        Disallow_Edit = True
      End If
      If (Request.QueryString("require_pp") <> "") Then
        form_require_pp = 1
        RequirePP = True
      End If
      If (Request.QueryString("DisallowEditPp") = "on") Then
        iDisallowEditPp = 1
        bDisallowEditPp = True
      End If
      If (Request.QueryString("DisallowEditValue") = "on") Then
        iDisallowEditValue = 1
        bDisallowEditValue = True
      End If
      If (Request.QueryString("DisallowEditRewards") = "on") Then
        iDisallowEditRewards = 1
        bDisallowEditRewards = True
      End If
      ' both requiring and locking the points program is not permitted 
      If (form_require_pp = 1 AndAlso (form_Disallow_Edit = 1 Or iDisallowEditPp = 1)) Then
        infoMessage = Copient.PhraseLib.Lookup("offer-con.lockeddenied", LanguageID)
      Else
        MyCommon.QueryStr = "update OfferConditions with (RowLock) set Disallow_Edit=" & form_Disallow_Edit & _
        ",RequiredFromTemplate=" & form_require_pp & _
        ",DisallowEdit1=" & iDisallowEditPp & _
        ",DisallowEdit2=" & iDisallowEditValue & _
        ",DisallowEdit3=" & iDisallowEditRewards & _
        " where ConditionID=" & ConditionID
        MyCommon.LRT_Execute()
        ' update the points program requirement
        MyCommon.QueryStr = "update OfferConditions with (RowLock) set RequiredFromTemplate = " & _
                            IIf(Request.QueryString("require_pp") <> "", 1, 0) & "where ConditionID = " & ConditionID
        MyCommon.LRT_Execute()
      End If
    End If
    ' udpate the flags for this condition
    ' update the flags
    MyCommon.QueryStr = "update OfferConditions with (RowLock) set TCRMAStatusFlag=2,CRMAStatusFlag=2 where ConditionID=" & ConditionID
    MyCommon.LRT_Execute()
    MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where offerid=" & OfferID
    MyCommon.LRT_Execute()
    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-sv", LanguageID))
    If (infoMessage = "") Then
      CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
    End If
  End If
  
  Send("<script type=""text/javascript"">")
    Send("function ChangeParentDocument() { ")
    Send("  if (opener != null) {")
    Send("    var newlocation = 'offer-con.aspx?OfferID=" & OfferID & "'; ")
    Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
  Send("    opener.location = 'offer-con.aspx?OfferID=" & OfferID & "'; ")
    Send("    } ")
    Send("  }")
    Send("}")
  Send("</script>")
%>
<form action="#" id="mainform" name="mainform" onkeydown="if (event.keyCode == 13) return false;" onsubmit="return validateEntry();">
  <div id="intro">
    <input type="hidden" id="OfferID" name="OfferID" value="<% sendb(OfferID) %>" />
    <input type="hidden" id="Name" name="Name" value="<% sendb(Name) %>" />
    <input type="hidden" id="ConditionID" name="ConditionID" value="<% sendb(ConditionID) %>" />
    <input type="hidden" name="ProgramID" id="ProgramID" value="<% Sendb(ProgramID) %>" />
    <input type="hidden" id="IsTemplate" name="IsTemplate" value="<% 
            if(istemplate)then 
            sendb("IsTemplate")
            else 
            sendb("Not") 
            end if
        %>" />
    <%If (IsTemplate) Then
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.storedvaluecondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
      Else
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.storedvaluecondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
      End If
      If (bUseTemplateLocks And bDisallowEditPp) Then
        Send("<input type=""hidden"" id=""DisallowEditPpOpt"" name=""DisallowEditPpOpt"" value=""1"" />")
      Else
        Send("<input type=""hidden"" id=""DisallowEditPpOpt"" name=""DisallowEditPpOpt"" value=""0"" />")
      End If
      If (bUseTemplateLocks And bDisallowEditValue) Then
        Send("<input type=""hidden"" id=""DisallowEditValueOpt"" name=""DisallowEditValueOpt"" value=""1"" />")
      Else
        Send("<input type=""hidden"" id=""DisallowEditValueOpt"" name=""DisallowEditValueOpt"" value=""0"" />")
      End If
    %>
    <div id="controls">
      <% If (IsTemplate) Then%>
      <span class="temp">
        <input type="checkbox" class="tempcheck" id="temp-employees" name="Disallow_Edit"
          <% if(disallow_edit)then sendb(" checked=""checked""") %> />
        <label for="temp-employees">
          <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
        </label>
      </span>
      <% End If%>
      <% If Not (IsTemplate) Then
           If (Logix.UserRoles.EditOffer And Not (bUseTemplateLocks And Disallow_Edit)) Then Send_Save()
         Else
           If (Logix.UserRoles.EditTemplates) Then Send_Save()
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
          <% If (IsTemplate) Then%>
          <span class="tempRequire">
            <input type="checkbox" class="tempcheck" id="require_pp1" name="require_pp" <% if(requirepp)then sendb(" checked=""checked""") %> />
            <label for="require_pp">
              <% Sendb(Copient.PhraseLib.Lookup("term.required", LanguageID))%>
            </label>
          </span><span class="tempLocked">
            <input type="checkbox" class="tempcheck" id="DisallowEditPp1" name="DisallowEditPp"
              <% if(bDisallowEditPp)then send(" checked=""checked""") %> />
            <label for="temp-Tiers">
              <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
            </label>
          </span>
          <% ElseIf (bUseTemplateLocks) Then%>
          <% If (RequirePP) Then%>
          <span class="tempRequire">
            <input type="checkbox" class="tempcheck" id="require_pp2" name="require_pp" disabled="disabled"
              checked="checked" />
            <label for="require_pp">
              <% Sendb(Copient.PhraseLib.Lookup("term.required", LanguageID))%>
            </label>
          </span>
          <% ElseIf (bDisallowEditPp) Then%>
          <span class="tempRequire">
            <input type="checkbox" class="tempcheck" id="DisallowEditPp2" name="DisallowEditPp"
              disabled="disabled" checked="checked" />
            <label for="temp-Tiers">
              <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
            </label>
          </span>
          <% End If%>
          <% End If%>
        </h2>
        <input type="radio" id="functionradio1" name="functionradio" <% if(MyCommon.Fetch_SystemOption(175)= "1" ) then sendb(" checked=""checked""") %> <% sendb(disabledattribute) %> /><label
          for="functionradio1"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
        <input type="radio" id="functionradio2" name="functionradio" <% if(MyCommon.Fetch_SystemOption(175)= "2" ) then sendb(" checked=""checked""") %> <% sendb(disabledattribute) %> /><label
          for="functionradio2"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
        <input class="medium" onkeyup="javascript:xmlPostTimer('OfferFeeds.aspx','StoredValueProgramsCM');" id="functioninput" name="functioninput" type="text" maxlength="100" value="" <% sendb(disabledattribute) %> />
         <% If (bCreateGroupOrProgramFromOffer AndAlso Logix.UserRoles.CreateStoredValuePrograms) Then%>
        <input class="regular" name="btncreate" id="btncreate" type="button" value="<% Sendb(Copient.PhraseLib.Lookup("term.create", LanguageID))%>" onclick="javascript:handleCreateClick('btncreate');" <% sendb(disabledattribute) %>/>
        <% End If%> 
         <br />
        <div id="searchLoadDiv" style="display: block;">&nbsp;</div>
        <div id="List">
        <select class="longer" id="functionselect" name="functionselect" size="16" <% sendb(disabledattribute) %>>
          <%
              Dim topString As String = ""
              If RECORD_LIMIT > 0 Then topString = "top " & RECORD_LIMIT
              MyCommon.QueryStr = "Select " & topString & "SVProgramID as ProgramID, Name as ProgramName, SVTypeID from StoredValuePrograms as PP with (NoLock) where Deleted=0 and SVProgramID is not null and SVTypeID not in (4,5) order by ProgramName"
            rst = MyCommon.LRT_Select
            For Each row In rst.Rows
              Send("<option value=" & MyCommon.NZ(row.Item("ProgramID"), -1) & ">" & MyCommon.NZ(row.Item("ProgramName"), "") & "</option>")
            Next
          %>
        </select>
        </div>
        <br />
        <%If (RECORD_LIMIT > 0) Then
          Send(Copient.PhraseLib.Lookup("groups.display", LanguageID) & ": " & RECORD_LIMIT.ToString() & "<br />")
        End If
        %>
        <br class="half" />
        <input class="regular select" id="select1" name="select1" type="button" value="&#9660; <% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID)) %>"
          onclick="handleSelectClick('select1');" <% sendb(disabledattribute) %> />&nbsp;
        <input class="regular deselect" id="deselect1" name="deselect1" type="button" value="<% Sendb(Copient.PhraseLib.Lookup("term.deselect", LanguageID)) %> &#9650;"
          onclick="handleSelectClick('deselect1');" <% sendb(disabledattribute) %> /><br />
        <br class="half" />
        <select class="longer" id="selected" name="selected" size="2" <% sendb(disabledattribute) %>>
          <%
            MyCommon.QueryStr = "select SVProgramID as ProgramID, Name as ProgramName, SVTypeID, OC.LinkID from StoredValuePrograms as SVP with (NoLock) " & _
                                "left join OfferConditions as OC with (NoLock) on OC.LinkID=SVP.SVProgramID where OC.ConditionID=" & ConditionID & ";"
            rst = MyCommon.LRT_Select
            For Each row In rst.Rows
              Send("<option value=""" & row.Item("ProgramID") & """>" & row.Item("ProgramName") & "</option>")
              SVTypeID = Int(MyCommon.NZ(row.Item("SVTypeID"), 0))
              If SVTypeID = 1 And intNumDecimalPlaces > 0 Then
                bNeedToFormat = True
              Else
                bNeedToFormat = False
              End If
            Next
          %>
        </select>
        <hr class="hidden" />
      </div>
    </div>
    <input type="hidden" id="SVTypeID" name="SVTypeID" value="<% sendb(SVTypeID) %>" />
    <div id="gutter">
    </div>
    <div id="column2">
      <div class="box" id="value">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.value", LanguageID))%>
          </span>
        </h2>
        <% If (IsTemplate Or (bUseTemplateLocks And bDisallowEditValue)) Then%>
        <span class="temp">
          <input type="checkbox" class="tempcheck" id="DisallowEditValue" name="DisallowEditValue"
            <% if(bDisallowEditValue)then send(" checked=""checked""") %> <% If(bUseTemplateLocks) Then sendb(" disabled=""disabled""") %> />
          <label for="temp-Tiers">
            <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
          </label>
        </span>
        <br class="printonly" />
        <% End If%>
        <label for="tier0">
          <% Sendb(Copient.PhraseLib.Lookup("condition.valueneeded", LanguageID))%>
        </label>
        <br />
        <%
          MyCommon.QueryStr = "select LinkID,Tiered,O.Numtiers,QtyUnitType,O.OfferID,CT.TierLevel,CT.AmtRequired from OfferConditions as OC with (NoLock) left  join Offers as O with (NoLock) on O.OfferID=OC.OfferID left  join ConditionTiers as CT with (NoLock) on OC.ConditionID=CT.ConditionID where OC.ConditionID=" & ConditionID
          rst = MyCommon.LRT_Select()
          Dim q As Integer
          q = 1
          If (bUseTemplateLocks And bDisallowEditValue) Then
            sDisabled = " disabled=""disabled"""
          Else
            sDisabled = ""
          End If
          For Each row In rst.Rows
            If bNeedToFormat Then
              decTemp = (Int(MyCommon.NZ(row.Item("AmtRequired"), 0)) * 1.0) / decFactor
              sTemp = FormatNumber(decTemp, intNumDecimalPlaces)
            Else
              sTemp = Int(MyCommon.NZ(row.Item("AmtRequired"), 0)).ToString()
            End If

            If (MyCommon.NZ(row.Item("Tiered"), 0) = 0) Then
              Send("<input class=""shorter"" id=""tier0"" name=""tier0"" type=""text"" maxlength=""9"" value=""" & sTemp & """" & sDisabled & " /><br />")
            Else
              Tiered = True
              Send("<label for=""tier" & q & """><b>" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & q & ":</b></label> <input class=""shorter"" id=""tier" & q & """ name=""tier" & q & """ type=""text"" value=""" & sTemp & """" & sDisabled & " /><br />")
            End If
            q = q + 1
          Next
          Send("<input type=""hidden"" name=""NumTiers"" value=""" & MyCommon.NZ(row.Item("NumTiers"), "") & """ />")
          Send("<input type=""hidden"" name=""Tiered"" value=""" & MyCommon.NZ(row.Item("Tiered"), "") & """ />")
          MyCommon.QueryStr = "select LinkID,ExcludedID,PointsRedeemInstant,MinOrderItemsOnly,GrantTypeID,DoNotItemDistribute,PointsUseEarnedValue,PointsUseNetValue from OfferConditions with (NoLock) where ConditionID=" & ConditionID & ";"
          rst = MyCommon.LRT_Select
          For Each row In rst.Rows
          Next
        %>
        <br class="half" />
        <%
          Sendb(Copient.PhraseLib.Lookup("condition.satisfy", LanguageID))
          If (MyCommon.NZ(row.Item("pointsredeeminstant"), False)) Then
            iRadioValue = 4
          ElseIf (MyCommon.NZ(row.Item("pointsuseearnedvalue"), False)) Then
            iRadioValue = 2
          ElseIf (MyCommon.NZ(row.Item("pointsusenetvalue"), False)) Then
            iRadioValue = 1
          Else
            iRadioValue = 0
          End If
        %>
        <br />
        <input class="radioValue" id="RadioValue1" name="radioValue" value="0" type="radio"
          <% if(iRadioValue=0)then sendb(" checked=""checked""") %><% If(bUseTemplateLocks and bDisallowEditValue) Then sendb(" disabled=""disabled""") %> />
        <label for="RadioValue1">
          <% Sendb(Copient.PhraseLib.Lookup("condition.PreviousValue", LanguageID))%>
        </label>
        <br />
        <input class="radioValue" id="RadioValue2" name="radioValue" value="2" type="radio"
          <% if(iRadioValue=2)then sendb(" checked=""checked""") %><% If(bUseTemplateLocks and bDisallowEditValue) Then sendb(" disabled=""disabled""") %> />
        <label for="RadioValue2">
          <% Sendb(Copient.PhraseLib.Lookup("condition.CurrentValue", LanguageID))%>
        </label>
        <br />
        <input class="radioValue" id="RadioValue3" name="radioValue" value="4" type="radio"<% if(iRadioValue=4)then sendb(" checked=""checked""") %><% If(bUseTemplateLocks and bDisallowEditValue) Then sendb(" disabled=""disabled""") %> />
        <label for="RadioValue3"><% Sendb(Copient.PhraseLib.Lookup("condition.TotalValue", LanguageID))%></label>
        <br />
        <input class="radioValue" id="RadioValue4" name="radioValue" value="1" type="radio"<% if(iRadioValue=1)then sendb(" checked=""checked""") %><% If(bUseTemplateLocks and bDisallowEditValue) Then sendb(" disabled=""disabled""") %> />
        <label for="RadioValue4"><% Sendb(Copient.PhraseLib.Lookup("condition.NetValue", LanguageID))%></label>
        <br />
        <hr class="hidden" />
      </div>
      <div class="box" id="grants" <%if(tiered)then send(" style=""visibility: hidden;""") %>>
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.rewards", LanguageID))%>
          </span>
        </h2>
        <% If (IsTemplate Or IsTemplate Or (bUseTemplateLocks And bDisallowEditRewards)) Then%>
        <span class="temp">
          <input type="checkbox" class="tempcheck" id="DisallowEditRewards" name="DisallowEditRewards"
            <% if(bDisallowEditRewards)then send(" checked=""checked""") %><% If(bUseTemplateLocks) Then sendb(" disabled=""disabled""") %> />
          <label for="temp-Tiers">
            <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
          </label>
        </span>
        <br class="printonly" />
        <% End If%>
        <% Sendb(Copient.PhraseLib.Lookup("condition.rewardsgranted", LanguageID))%>
        <br />
        <input class="radio" id="eachtime" name="granted" value="3" type="radio" <% if(MyCommon.NZ(row.item("granttypeid"), 0)=3)then sendb(" checked=""checked""") %><% If(bUseTemplateLocks and bDisallowEditRewards) Then sendb(" disabled=""disabled""") %> />
        <label for="eachtime">
          <% Sendb(Copient.PhraseLib.Lookup("condition.eachtime", LanguageID))%>
        </label>
        <br />
        <input class="radio" id="equalto" name="granted" value="1" type="radio" <% if(MyCommon.NZ(row.item("granttypeid"),0)=1)then sendb(" checked=""checked""") %><% If(bUseTemplateLocks and bDisallowEditRewards) Then sendb(" disabled=""disabled""") %> />
        <label for="equalto">
          <% Sendb(Copient.PhraseLib.Lookup("condition.equalto", LanguageID))%>
        </label>
        <br />
        <input class="radio" id="greaterthan" name="granted" value="2" type="radio" <% if(MyCommon.NZ(row.item("granttypeid"),0)=2)then sendb(" checked=""checked""") %><% If(bUseTemplateLocks and bDisallowEditRewards) Then sendb(" disabled=""disabled""") %> />
        <label for="greaterthan">
          <% Sendb(Copient.PhraseLib.Lookup("condition.greaterthan", LanguageID))%>
        </label>
        <br />
      </div>
    </div>
  </div>
</form>

<script type="text/javascript">
<% If (CloseAfterSave) Then %>
    window.close();
<% End If %>
//handleKeyUp(9999);
updateButtons();

</script>

<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd("mainform", "functioninput")
  MyCommon = Nothing
  Logix = Nothing
%>
