<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: offer-rew-membership.aspx 
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
  Dim rst2 As DataTable
  Dim row2 As DataRow
  Dim OfferID As Long
  Dim Name As String = ""
  Dim RewardID As String
  Dim NumTiers As Integer
  Dim RewardTypeID As Integer
  Dim Tiered As Integer
  Dim TriggerQty As Integer
  Dim x As Integer
  Dim ExcludedItem As Integer
  Dim SelectedItem As Integer
  Dim transactionlevelselected = False
  Dim Disallow_Edit As Boolean = True
  Dim FromTemplate As Boolean
  Dim IsTemplate As Boolean = False
  Dim CloseAfterSave As Boolean = False
  Dim DisabledAttribute As String = ""
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim BannersEnabled As Boolean = False
  Dim RECORD_LIMIT As Integer = GroupRecordLimit '500
  Dim bCreateGroupOrProgramFromOffer As Boolean = IIf(MyCommon.Fetch_CM_SystemOption(134) ="1",True,False)
    
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "offer-rew-membership.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  OfferID = Request.QueryString("OfferID")
  RewardID = Request.QueryString("RewardID")
  
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  
  If (Request.QueryString("save") <> "") Then
    CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
  End If
  
  ' fetch the name
  MyCommon.QueryStr = "Select Name,IsTemplate,FromTemplate from Offers with (NoLock) where OfferID=" & OfferID
  rst = MyCommon.LRT_Select
  If (rst.Rows.Count > 0) Then
    Name = rst.Rows(0).Item("Name")
    IsTemplate = MyCommon.NZ(rst.Rows(0).Item("IsTemplate"), False)
    FromTemplate = MyCommon.NZ(rst.Rows(0).Item("FromTemplate"), False)
  End If
  
  ' ok we need to know if we are tiered and if so draw this thing tier times
  MyCommon.QueryStr = "Select Tiered,O.NumTiers,RewardTypeID,TriggerQty from OfferRewards as OFR with (NoLock) left join Offers as O with (NoLock) on O.OfferID=OFR.OfferID where RewardID=" & RewardID
  rst = MyCommon.LRT_Select
  For Each row In rst.Rows
    NumTiers = row.Item("NumTiers")
    Tiered = row.Item("Tiered")
    RewardTypeID = row.Item("RewardTypeID")
    TriggerQty = row.Item("TriggerQty")
  Next
  
  If (Request.QueryString("pgroup-add1") <> "" And Request.QueryString("pgroup-avail") <> "") Then
    MyCommon.QueryStr = "update OfferRewards with (RowLock) set ProductGroupID=0 where RewardID=" & RewardID & ";"
    MyCommon.LRT_Execute()
    MyCommon.QueryStr = "update OfferRewards with (RowLock) set ProductGroupID=" & Request.QueryString("pgroup-avail") & " where RewardID=" & RewardID & ";"
    MyCommon.LRT_Execute()
  ElseIf (Request.QueryString("pgroup-rem1") <> "") Then
    MyCommon.QueryStr = "update OfferRewards with (RowLock) set ProductGroupID=0 where RewardID=" & RewardID & ";"
    MyCommon.LRT_Execute()
  ElseIf (Request.QueryString("pgroup-add2") <> "" And Request.QueryString("pgroup-avail") <> "" And Request.QueryString("pgroup-avail") <> "1") Then
    MyCommon.QueryStr = "update OfferRewards with (RowLock) set ExcludedProdGroupID=0 where RewardID=" & RewardID & ";"
    MyCommon.LRT_Execute()
    MyCommon.QueryStr = "update OfferRewards with (RowLock) set ExcludedProdGroupID=" & Request.QueryString("pgroup-avail") & " where RewardID=" & RewardID & ";"
    MyCommon.LRT_Execute()
  ElseIf (Request.QueryString("pgroup-rem2") <> "") Then
    MyCommon.QueryStr = "update OfferRewards with (RowLock) set ExcludedProdGroupID=0 where RewardID=" & RewardID & ";"
    MyCommon.LRT_Execute()
  End If
  
  If (Request.QueryString("addvalue") <> "") Then
    MyCommon.QueryStr = "select max(TierLevel) as maxtier from RewardTiers with (RowLock) where RewardID=" & RewardID
    rst = MyCommon.LRT_Select
    ' ok we know the highest tier now so we need to add one
    'dbo.pt_ConditionTiers_Update @ConditionID bigint, @TierLevel int,@AmtRequired decimal(12,3)
    MyCommon.QueryStr = "dbo.pt_RewardTiers_Update"
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add("@RewardID", SqlDbType.BigInt).Value = RewardID
    MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int).Value = MyCommon.Extract_Val(MyCommon.NZ(rst.Rows(0).Item("maxtier"), 0)) + 1
    MyCommon.LRTsp.Parameters.Add("@RewardAmount", SqlDbType.Decimal).Value = 0
    MyCommon.LRTsp.ExecuteNonQuery()
    MyCommon.Close_LRTsp()
  ElseIf (Request.QueryString("save") <> "" Or _
  Request.QueryString("pgroup-add1") <> "" Or _
  Request.QueryString("pgroup-rem1") <> "" Or _
  Request.QueryString("pgroup-add2") <> "" Or _
  Request.QueryString("pgroup-rem2") <> "") Then
    ' minorderitems nodistribute
    ' delete any existing for this offer
    MyCommon.QueryStr = "delete from RewardCustomerGroupTiers with (RowLock) where RewardID=" & RewardID
        MyCommon.LRT_Execute()
        If (Request.QueryString("selected2") <> "") Then
            For x = 1 To NumTiers
                MyCommon.QueryStr = "insert into RewardCustomerGroupTiers with (RowLock) (RewardID,TierLevel,CustomerGroupID) values(" & RewardID & "," & x & "," & MyCommon.Extract_Val(Request.QueryString("selected" & x)) & ")"
                MyCommon.LRT_Execute()
                '"t" & t & "
            Next
        Else
            MyCommon.QueryStr = "insert into RewardCustomerGroupTiers with (RowLock) (RewardID,TierLevel,CustomerGroupID) values(" & RewardID & ",0," & MyCommon.Extract_Val(Request.QueryString("selected1")) & ")"
            MyCommon.LRT_Execute()
        End If
    TriggerQty = Int(Request.QueryString("Xbox"))
    If (TriggerQty < 0) Then
      TriggerQty = 0
    End If
    MyCommon.QueryStr = "update OfferRewards with (RowLock) set TriggerQty=" & TriggerQty & ",TCRMAStatusFlag=2,CMOAStatusFlag=2 where RewardID=" & RewardID
    MyCommon.LRT_Execute()
    MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where offerid=" & OfferID
    MyCommon.LRT_Execute()
    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.rew-membership", LanguageID))
    If (Request.QueryString("IsTemplate") = "IsTemplate") Then
      ' time to update the status bits for the templates
      Dim form_Disallow_Edit As Integer = 0
      If (Request.QueryString("Disallow_Edit") = "on") Then
        form_Disallow_Edit = 1
      End If
      MyCommon.QueryStr = "update OfferRewards with (RowLock) set Disallow_Edit=" & form_Disallow_Edit & _
                          " where RewardID=" & RewardID
      MyCommon.LRT_Execute()
    End If
  End If
  
  If (IsTemplate Or FromTemplate) Then
    ' lets dig the permissions if its a template
    MyCommon.QueryStr = "select Disallow_Edit from OfferRewards with (NoLock) where RewardID=" & RewardID
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      For Each row In rst.Rows
        ' ok there are some rows for the template
        Disallow_Edit = MyCommon.NZ(row.Item("Disallow_Edit"), True)
      Next
    End If
  End If
  
  Send_HeadBegin("term.offer", "term.membershipreward", OfferID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
%>
<script type="text/javascript">
  // This is the javascript array holding the function list
  // The PrintJavascriptArray ASP function can be used to print this array.
  var isFireFox = (navigator.appName.indexOf('Mozilla') != -1) ? true : false;
  var timer;
  function xmlPostTimer(strURL, mode) {
    clearTimeout(timer);
    timer = setTimeout("xmlhttpPost('" + strURL + "','" + mode + "')", 250);
  }

  function xmlhttpPost(strURL, mode) {
    var xmlHttpReq = false;
    var self = this;

    //document.getElementById("functionselect").style.display = "none";
    document.getElementById("searchLoadDiv").innerHTML = '<% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID)) %>';

    // Mozilla/Safari
    if (window.XMLHttpRequest) {
      self.xmlHttpReq = new XMLHttpRequest();
    }
    // IE
    else if (window.ActiveXObject) {
      self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
    }
    var qryStr = getgroupquery(mode);
    self.xmlHttpReq.open('POST', strURL + "?" + qryStr, true);
    self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    self.xmlHttpReq.onreadystatechange = function () {
      if (self.xmlHttpReq != null && self.xmlHttpReq.readyState == 4) {
        updatepage(self.xmlHttpReq.responseText);
      }
    }

    self.xmlHttpReq.send(qryStr);
    //self.xmlHttpReq.send(getquerystring());
  }

  function getgroupquery(mode) {
    var radioString;
    if (document.getElementById('functionradio2').checked) {
      radioString = 'functionradio2';
    }
    else {
      radioString = 'functionradio1';
    }

  return "Mode=" + mode + "&OfferID=" + document.getElementById('OfferID').value + "&RewardID=" + document.getElementById('RewardID').value + "&Search=" + document.getElementById('functioninput').value + "&EngineID=0&SearchRadio=" + radioString + "" + GetQueryString();

  }

  function updatepage(str) {
    if (str.length > 0) {
      if (!isFireFox) {
        document.getElementById("cgList").innerHTML = '<select class="longer" id="functionselect" name="functionselect" onkeydown="return handleSlctKeyDown(event);" onclick="handleSelectClick();" ondblclick="addGroupToSelect();" size="12"<% Sendb(DisabledAttribute) %>>' + str + '</select>';
      }
      else {
        document.getElementById("functionselect").innerHTML = str;
      }

      document.getElementById("searchLoadDiv").innerHTML = '&nbsp;';
      if (document.getElementById("functionselect").options.length > 0) {
        document.getElementById("functionselect").options[0].selected = true;
      }
    }
    else if (str.length == 0) {
      if (!isFireFox) {
        document.getElementById("cgList").innerHTML = '<select class="longer" id="functionselect" name="functionselect" onkeydown="return handleSlctKeyDown(event);" onclick="handleSelectClick();" ondblclick="addGroupToSelect();" size="12"<% Sendb(DisabledAttribute) %>></select>';
      }
      else {
        document.getElementById("functionselect").innerHTML = '';
      }

      document.getElementById("searchLoadDiv").innerHTML = '&nbsp;';
    }
  }

  function GetQueryString() {
      var selectedGroups = ""
      var qString = "";
      var tiered = document.getElementById("Tiered").value;
      var tier = document.getElementById("NumTiers").value;
      var SelectCount = 0;
      if (tiered == 0) {
          var selElem = document.getElementById("selected1");
          if (selElem.options.length > 0) {
              SelectCount = 1;
              qString = qString + "Group1=" + selElem.options[0].value;
          }
      }
      else if (tier > 0) {
       
          for (i=1; i <= tier; i++) {
              var selElem = document.getElementById("selected" + i);
              if (i != tier) {
                  if (selElem.options.length > 0) {
                      SelectCount++;
                      qString = qString + "Group" + SelectCount + "=" + selElem.options[0].value + "&";
                  }
              }
              else {
                  if (selElem.options.length > 0) {
                      SelectCount++
                      qString = qString + "Group" + SelectCount + "=" + selElem.options[0].value;
                  }
              }
          }
      }
      if (qString != "") {
          qString = "&GroupCount=" + SelectCount + "&" + qString;
      }
      return qString;
  }
  

  // This is the function that refreshes the list after a keypress.
  // The maximum number to show can be limited to improve performance with
  // huge lists (1000s of entries).
  // The function clears the list, and then does a linear search through the
  // globally defined array and adds the matches back to the list.
  function handleKeyUp(maxNumToShow, TierLevels) {
    var selectObj, textObj, functionListLength;
    var i;
    var t;
    var numShown;
    var searchPattern;
    var selectedList;

    document.getElementById("functionselect").size = "12";

    // Set references to the form elements
    selectObj = document.forms[0].functionselect;
    textObj = document.forms[0].functioninput;
    selectedList = document.getElementById("t1_selected");

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
        //      if (vallist[i] != "" && (selectedList.options.length < 1 || vallist[i] != selectedList.options[0].value)) {      if (vallist[i] != "" && (selectedList.options.length < 1 || vallist[i] != selectedList.options[0].value)) {
        if (vallist[i] != "") {
          selectObj[numShown] = new Option(functionlist[i], vallist[i]);
          if (vallist[i] == 2) {
            selectObj[numShown].style.fontWeight = 'bold';
            selectObj[numShown].style.color = 'brown';
          }
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

  function addGroupToSelect(t) {
    var elemSource = document.getElementById("functionselect");
    var elemDest = document.getElementById("selected" + t);
    var selOption = null;
    var selText = "", selVal = "";
    var selIndex = -1;

    if (elemSource != null && elemSource.options.selectedIndex == -1) {
      alert('<% Sendb(Copient.PhraseLib.Lookup("CPE-rew-membership.selectgroup", LanguageID)) %>');
      elemSource.focus();
    } else {
      selIndex = elemSource.options.selectedIndex;
      selOption = elemSource.options[selIndex];
      selText = selOption.text;
      selVal = selOption.value;
      if (elemDest != null && elemDest.options.length > 0) {
        removeGroupFromSelect(t);
      }
      elemDest.options[0] = new Option(selText, selVal);
      document.getElementById("deselect" + t).disabled = false;
      elemDest.options[0].style.fontweight = 'bold';
      if (selVal == 2) {
        elemDest.options[0].style.color = 'brown';
        elemSource.options[selIndex] = null;
      }
      //handleKeyUp(99999); 
      removeUsed(t);
    }
  }

  
  function removeGroupFromSelect(t) {
      var elem;
      var elemList = document.getElementById("functionselect");

      if (t > 0) {
          elem = document.getElementById('selected' + t);
      }
      else {
          elem = document.getElementById('selected1');
      }
      if (elem != null && elem.options.length > 0) {
          elem.options[0] = null;
          document.getElementById("deselect" + t).disabled = true;
      }
      removeUsed(t);
  }

  function validateEntry(tierLevels) {
      var retVal = true;
      var i;
      var elem = document.getElementById("t0_selected");
      var elemGroup = document.getElementById("t0_CustomerGroupID");
      // Loop through the tiers
      for (i = 0; i <= tierLevels; i++) {
          elem = document.getElementById("t" + i + "_selected");
          elemGroup = document.getElementById("t" + i + "_CustomerGroupID");
          if (elem != null && elemGroup != null) {
              if (elem.options.length == 0) {
                  retVal = false;
                  alert('<% Sendb(Copient.PhraseLib.Lookup("CPE-rew-membership.selectgroup", LanguageID)) %>');
                  elem.focus();
              } else {
                  elemGroup.value = elem.options[0].value;
              }
          }
      }
      return retVal;
  }
  // this function will remove items from the functionselect box that are used in 
  // selected and excluded boxes
  function removeUsed(t) {
      xmlPostTimer('OfferFeeds.aspx', 'GrantMembership');
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
        addGroupToSelect();
        clearEntry();
      }
      e.returnValue = false;
      return false;
    }
  }

  function clearEntry() {
    var elemInput = document.getElementById("functioninput");

    if (elemInput != null) {
      elemInput.value = "";
      //handleKeyUp(200);
      elemInput.focus();
    }
  }


function handleCreateClick(createbtn) {
    var alertMessage = "";
    if (document.getElementById(createbtn) != undefined && document.getElementById(createbtn) != null) {
        var searchText = document.getElementById('functioninput').value;
        if (searchText != null && searchText != "") {
            xmlhttpPost_CreateGroupOrProgramFromOffer('OfferFeeds.aspx', 'CreateGroupOrProgramFromOffer');
        }
        else {
            alertMessage = '<% Sendb(Copient.PhraseLib.Lookup("term.enter", LanguageID))%>' + ' ' + '<% Sendb(Copient.PhraseLib.Lookup("term.valid", LanguageID).ToLower())%>' + ' ' + '<% Sendb(Copient.PhraseLib.Lookup("term.customergroup", LanguageID).ToLower())%>' + ' ' + '<% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID).ToLower())%>';
            alert(alertMessage);
            document.getElementById('functioninput').focus();
            return false;
        }
    }
    return true;
}

function xmlhttpPost_CreateGroupOrProgramFromOffer(strURL, mode) {
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
    self.xmlHttpReq.onreadystatechange = function () {
        if (self.xmlHttpReq != null && self.xmlHttpReq.readyState == 4) {
            updatepage_creategroupOrprogramfromoffer(self.xmlHttpReq.responseText);
        }
    }

    self.xmlHttpReq.send(qryStr);
}
function getcreatequery(mode) {
    return "Mode=" + mode + "&CreateType=CustomerGroup&Name=" + document.getElementById('functioninput').value
}
function updatepage_creategroupOrprogramfromoffer(str) {

    var tiered = document.getElementById("Tiered").value;
    if (str.length > 0) {
        var status = "";
        var responseArr = str.split('~');
        if (responseArr.length > 0) {
            status = responseArr[0];
            if (status == "Ok") {
                var resultArr = responseArr[1].split('|');
                if (tiered != 0) {
                    addNewGrouptoSelectbox(resultArr[0], resultArr[1]);
                    document.getElementById('functionselect').value = resultArr[1];
                }
                else {
                    removeGroupFromSelect('1');
                    addNewGrouptoSelectbox(resultArr[0], resultArr[1]);
                    document.getElementById('functionselect').value = resultArr[1];
                    addGroupToSelect(1);
                }
            }
            else if (status == "Fail") {
                var resultArr = responseArr[1].split('|');
                var selectedGroupValue = -1;
                var groupExist=false;
                if (tiered != 0) {//more than one tier and should check any of the selected dropdown contains the value
                var numtiers=document.getElementById('NumTiers').value;
                for(var i=1 ;i<=numtiers;i++)
                {
                if(document.getElementById('selected'+ i)!=null)
                    {
                        if (parseInt(document.getElementById('selected'+ i).value) == parseInt(resultArr[1]))
                        {
                            groupExist=true; break;
                        }
                    }
                }
                 
                    if (groupExist == false) 
                    {
                        alert(responseArr[2]);
                        document.getElementById('functionselect').value = resultArr[1];
                    }
                    else 
                    {
                        alert('<% Sendb(Copient.PhraseLib.Lookup("term.customergroup", LanguageID))%>' + ": '" + resultArr[0] + "' " + '<% Sendb(Copient.PhraseLib.Lookup("term.is", LanguageID).ToLower())%>' + '  ' + '<% Sendb(Copient.PhraseLib.Lookup("term.already", LanguageID).ToLower())%>' + '  ' + '<% Sendb(Copient.PhraseLib.Lookup("term.selected", LanguageID).ToLower())%>');
                    }
                }
                else {
                
                    if (document.getElementById('selected1')!=null)
                        selectedGroupValue = document.getElementById('selected1').value;
                    if (parseInt(selectedGroupValue) == parseInt(resultArr[1])) {
                      alert('<% Sendb(Copient.PhraseLib.Lookup("term.customergroup", LanguageID))%>' + ": '" + resultArr[0] + "' " + '<% Sendb(Copient.PhraseLib.Lookup("term.is", LanguageID).ToLower())%>' + '  ' + '<% Sendb(Copient.PhraseLib.Lookup("term.already", LanguageID).ToLower())%>' + '  ' + '<% Sendb(Copient.PhraseLib.Lookup("term.selected", LanguageID).ToLower())%>');
                    }
                    else {
                        alert(responseArr[2]);
                        removeGroupFromSelect('1');
                        document.getElementById('functionselect').value = resultArr[1];
                        addGroupToSelect(1);
                        
                    }
                }
            }
            else if (status == "Error") {
                alert(responseArr[1]);
                return false;
            }
          }
        }
    
    document.getElementById("searchLoadDiv").innerHTML = '&nbsp;';
  
}

function addNewGrouptoSelectbox(text, val) {
    var sel = document.getElementById('functionselect');
    var opt = document.createElement('option'); // create new option element
    // create text node to add to option element (opt)
    opt.appendChild(document.createTextNode(text));
    opt.value = val; // set value property of opt
    sel.appendChild(opt); // add opt to end of select box (sel)

}
</script>
<%
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
 
  Send("<script type=""text/javascript"">")
    Send("function ChangeParentDocument() { ")
    Send("  if (opener != null) {")
    Send("    var newlocation = 'offer-rew.aspx?OfferID=" & OfferID & "'; ")
    Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
  Send("    opener.location = 'offer-rew.aspx?OfferID=" & OfferID & "'; ")
    Send("    } ")
    Send("  }")
    Send("}")
  Send("</script>")
%>
<form action="#" id="mainform" name="mainform" onkeydown="if (event.keyCode == 13) return false;">
<div id="intro">
  <input type="hidden" id="OfferID" name="OfferID" value="<% sendb(OfferID) %>" />
  <input type="hidden" id="Name" name="Name" value="<% sendb(Name) %>" />
  <input type="hidden" id="RewardID" name="RewardID" value="<% sendb(RewardID) %>" />
  <input type="hidden" id="IsTemplate" name="IsTemplate" value="<% 
            if(istemplate)then 
            sendb("IsTemplate")
            else 
            sendb("Not") 
            end if
        %>" />
  <%If (IsTemplate) Then
      Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.membershipreward", LanguageID), VbStrConv.Lowercase) & "</h1>")
    Else
      Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.membershipreward", LanguageID), VbStrConv.Lowercase) & "</h1>")
    End If
  %>
  <div id="controls">
    <% If (IsTemplate) Then%>
    <span class="temp">
      <input type="checkbox" class="tempcheck" id="temp-employees" name="Disallow_Edit"
        <% if(disallow_edit)then sendb(" checked=""checked""") %> />
      <label for="temp-employees">
        <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
    </span>
    <% End If%>
    <%If Not (IsTemplate) Then
        If (Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Edit)) Then Send_Save()
      Else
        If (Logix.UserRoles.EditTemplates) Then Send_Save()
      End If
    %>
  </div>
</div>
<div id="main">
  <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
  <div id="column1">
    <div class="box" id="groupselect">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.group", LanguageID))%>
        </span>
      </h2>
      <%
        If (RewardTypeID = 5) Then
          Send("<label for=""functionselect"">" & Copient.PhraseLib.Lookup("reward.groupgrant", LanguageID) & "</label><br />")
        ElseIf (RewardTypeID = 6) Then
          Send("<label for=""functionselect"">" & Copient.PhraseLib.Lookup("reward.groupremove", LanguageID) & "</label><br />")
        End If
      %>
      <br class="half" />
      <input type="radio" id="functionradio1" name="functionradio" <% if(MyCommon.Fetch_SystemOption(175)= "1" ) then sendb(" checked=""checked""") %> <% sendb(disabledattribute) %> /><label
        for="functionradio1"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
      <input type="radio" id="functionradio2" name="functionradio" <% if(MyCommon.Fetch_SystemOption(175)= "2" ) then sendb(" checked=""checked""") %> <% sendb(disabledattribute) %> /><label
        for="functionradio2"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
      <input type="text" class="medium" id="functioninput" name="functioninput" onkeyup="javascript:xmlPostTimer('OfferFeeds.aspx','GrantMembership');"
        maxlength="100" value="" <% Sendb(DisabledAttribute) %> />
     <% If (bCreateGroupOrProgramFromOffer AndAlso Logix.UserRoles.CreateCustomerGroups) Then%>
        <input class="regular" name="btncreate" id="btncreate" type="button" value="<% Sendb(Copient.PhraseLib.Lookup("term.create", LanguageID))%>" onclick="javascript:handleCreateClick('btncreate');" <% if(FromTemplate And Disallow_Edit) then sendb(" disabled=""disabled""") %>/>
     <% End If%>  
      <br />
      <div id="searchLoadDiv" style="display: block;">
        &nbsp;</div>
      <div id="cgList">
        <select class="longer" id="functionselect" name="functionselect" size="10" <% sendb(disabledattribute) %>>
          <%
            Dim topString As String = ""
            If (RECORD_LIMIT > 0) Then topString = " top " & RECORD_LIMIT
                If (RewardTypeID = 5) Then
                  MyCommon.QueryStr = "Select distinct" & topString & " CustomerGroupID,Name from CustomerGroups as CG with (NoLock) Left Outer Join OfferConditions OC ON CG.CustomerGroupID <> OC.LinkID " & _
                            "where CG.AnyCardholder <> 1 and CG.AnyCustomer <> 1 and CG.NewCardholders <> 1 and CG.Deleted = 0 " & _
                            "and OC.ConditionTypeID = 1 and OC.OfferID = " & OfferID & " " & _
                            "ORDER BY CG.CustomerGroupID desc, Name"
                Else
                  MyCommon.QueryStr = "Select distinct " & topString & " CustomerGroupID,Name from CustomerGroups with (NoLock) " & _
                            "where AnyCardholder <> 1 and AnyCustomer <> 1  and NewCardholders <> 1 and deleted=0 " & _
                            "ORDER BY CustomerGroupID desc, Name"
                End If 

            rst2 = MyCommon.LRT_Select
            Dim RowSelected As Integer
            If (rst2.Rows.Count > 0) Then
              RowSelected = rst2.Rows(0).Item("CustomerGroupID")
            Else
              RowSelected = 0
            End If
            For Each row In rst2.Rows
              Send("<option value=" & row.Item("CustomerGroupID") & ">" & row.Item("Name") & "</option>")
            Next
          %>
        </select>
      </div>
      <%If (RECORD_LIMIT > 0) Then
          Send(Copient.PhraseLib.Lookup("groups.display", LanguageID) & ": " & RECORD_LIMIT.ToString() & "<br />")
        End If
      %>
      <br class="half" />
      <% 
          Send("<input type=""hidden"" id=""Tiered"" name=""Tiered"" value=""" & Tiered & """ />")
          Send("<input type=""hidden""  id=""NumTiers"" name=""NumTiers"" value=""" & NumTiers & """ />")
        If (Tiered = 0) Then
          MyCommon.QueryStr = "Select RewardID,RC.CustomerGroupID,CG.Name,CG.CustomerGroupID from RewardCustomerGroupTiers as RC with (NoLock) left join CustomerGroups as CG with (NoLock) on CG.CustomerGroupID=RC.CustomerGroupID where RewardID=" & RewardID & " and RC.CustomerGroupID<>0 and TierLevel=0;"
          'MyCommon.QueryStr = "Select CG.CustomerGroupID,CG.Name,ORW.LinkID from CustomerGroups as CG with (NoLock) left join OfferRewards as ORW with (NoLock) on ORW.LinkID=CG.CustomerGroupID where ORW.RewardID=" & RewardID & ";"
          rst = MyCommon.LRT_Select
            
          MyCommon.QueryStr = "Select RewardID,RC.CustomerGroupID,CG.Name,CG.CustomerGroupID from RewardCustomerGroupTiers as RC with (NoLock) left join CustomerGroups as CG with (NoLock) on CG.CustomerGroupID=RC.CustomerGroupID where RewardID=" & RewardID & " and RC.CustomerGroupID<>0 and TierLevel=" & x & ";"
          'MyCommon.QueryStr = "Select CG.CustomerGroupID,CG.Name,ORW.LinkID from CustomerGroups as CG with (NoLock) left join OfferRewards as ORW with (NoLock) on ORW.LinkID=CG.CustomerGroupID where ORW.RewardID=" & RewardID & ";"
          rst = MyCommon.LRT_Select
                    
          Send("<input class=""regular select"" id=""select1"" name=""select1"" type=""button"" value=""&#9660; " & Copient.PhraseLib.Lookup("term.select", LanguageID) & """ onclick=""addGroupToSelect(1);""" & IIf(rst.Rows.Count < 0, " disabled=""disabled""", "") & " />&nbsp;")
          Send("<input class=""regular deselect"" id=""deselect1"" name=""deselect1"" type=""button"" value=""" & Copient.PhraseLib.Lookup("term.deselect", LanguageID) & " &#9650;"" onclick=""removeGroupFromSelect('1');""" & IIf(rst.Rows.Count = 0, " disabled=""disabled""", "") & " /><br />")
          Send("<br class=""half"" />")
          Send("<select class=""longer"" id=""selected1"" name=""selected1"" size=""1""" & DisabledAttribute & ">")
          For Each row In rst.Rows
            Send("<option value=""" & row.Item("CustomerGroupID") & """>" & row.Item("Name") & "</option>")
          Next
          Send("</select>")
        Else
          For x = 1 To NumTiers
            MyCommon.QueryStr = "Select RewardID,RC.CustomerGroupID,CG.Name,CG.CustomerGroupID from RewardCustomerGroupTiers as RC with (NoLock) left join CustomerGroups as CG with (NoLock) on CG.CustomerGroupID=RC.CustomerGroupID where RewardID=" & RewardID & " and RC.CustomerGroupID<>0 and TierLevel=" & x & ";"
            'MyCommon.QueryStr = "Select CG.CustomerGroupID,CG.Name,ORW.LinkID from CustomerGroups as CG with (NoLock) left join OfferRewards as ORW with (NoLock) on ORW.LinkID=CG.CustomerGroupID where ORW.RewardID=" & RewardID & ";"
            rst = MyCommon.LRT_Select
            Send("<label for=""tiergroup" & x & """><b>Tier " & x & ":</b></label><br />")
            Send("<input class=""regular select"" id=""select" & x & """ name=""select" & x & """ type=""button"" value=""&#9660; " & Copient.PhraseLib.Lookup("term.select", LanguageID) & """ onclick=""addGroupToSelect('" & x & "');""" & IIf(rst.Rows.Count < 0, " disabled=""disabled""", "") & " />&nbsp;")
            Send("<input class=""regular deselect"" id=""deselect" & x & """ name=""deselect" & x & """ type=""button"" value=""" & Copient.PhraseLib.Lookup("term.deselect", LanguageID) & " &#9650;"" onclick=""removeGroupFromSelect('" & x & "');""" & IIf(rst.Rows.Count = 0, " disabled=""disabled""", "") & " /><br />")
            Send("<br class=""half"" />")
            Send("<select class=""longer"" id=""selected" & x & """ name=""selected" & x & """ size=""1""" & DisabledAttribute & ">")
            For Each row In rst.Rows
              Send("<option value=""" & row.Item("CustomerGroupID") & """>" & row.Item("Name") & "</option>")
            Next
            Send("</select>")
          Next
        End If
      %>
      <hr class="hidden" />
    </div>
    
  </div>
  <div id="gutter">
  </div>
  <div id="column2">
    <div class="box" id="distribution" style="display: none;">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.distribution", LanguageID))%>
        </span>
      </h2>
      <label for="Xbox">
        <% Sendb(Copient.PhraseLib.Lookup("term.mustpurchase", LanguageID))%></label>
      <input class="shortest" id="Xbox" name="Xbox" maxlength="9" type="text" <% sendb(" value=""" & triggerqty & """ ") %> />
      <% Sendb(Copient.PhraseLib.Lookup("term.item(s)", LanguageID))%>
    </div>
  </div>
</div>
</form>
<script type="text/javascript">
var t = document.getElementsByName("NumTiers")[0].value;
var k;
<% If (CloseAfterSave) Then %>
    window.close();
<% End If %>
if(t > 0) 
  for(k=1; k<=t; k++) {
    removeUsed(k);
  }
else
  removeUsed(1);
</script>
<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd("mainform", "functioninput")
  Logix = Nothing
  MyCommon = Nothing
%>
