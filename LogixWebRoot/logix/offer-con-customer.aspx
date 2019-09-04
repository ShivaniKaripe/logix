<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%@ Import Namespace="CMS.AMS" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: offer-con-customer.aspx 
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
  Dim rstBannerCGs As DataTable = Nothing
  Dim row As DataRow
  Dim OfferID As Long
  Dim Name As String = ""
  Dim ConditionID As String
  Dim ExcludedItem As Integer
  Dim SelectedItem As Integer
  Dim Disallow_Edit As Boolean = True
  Dim FromTemplate As Boolean
  Dim IsTemplate As Boolean = False
  Dim CloseAfterSave As Boolean = False
  Dim DisabledAttribute As String = ""
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim RequireCG As Boolean = False
  Dim CustomerGroupID As Integer = 0
  ' RD-AMSPS-548 Card Required Store Attribute Begin
  Dim ConditionAllowsAnyCardholderOverride As Integer = 0
  ' RD-AMSPS-548 Card Required Store Attribute End
  Dim ExcludedID As Integer = 0
  Dim BannersEnabled As Boolean = False
  Dim i As Integer
  Dim NewCardholdersID As Integer = 0
  Dim EngineID As Integer = 0
  Dim RECORD_LIMIT As Integer = GroupRecordLimit '500
  Dim AllCAM As Integer = 0
  Dim EligibleIncludedcustomergroups As String = String.Empty
  Dim EligibleExcludedcustomergroups As String = String.Empty
  Dim IsEligibilityConditionExistForOffer As String = "False"
  Dim bCreateGroupOrProgramFromOffer As Boolean = IIf(MyCommon.Fetch_CM_SystemOption(134) = "1", True, False)
    
  CMS.AMS.CurrentRequest.Resolver.AppName = "offer-con-customer.aspx"
  Dim m_Condition As CMS.AMS.Contract.ICustomerGroupCondition = CMS.AMS.CurrentRequest.Resolver.Resolve(Of CMS.AMS.Contract.ICustomerGroupCondition)()
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "offer-con-customer.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  OfferID = Request.QueryString("OfferID")
  If Request.QueryString("mode") = "optout" Then
    Dim objCondition As Models.CustomerGroupConditions = m_Condition.GetOfferCustomerCondition(OfferID, EngineID)
    ConditionID = objCondition.ConditionID
  Else
    ConditionID = Request.QueryString("ConditionID")
  End If
  CustomerGroupID = MyCommon.Extract_Val(Request.QueryString("CustomerGroupID"))
  ExcludedID = MyCommon.Extract_Val(Request.QueryString("ExcludedID"))
  
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  ' get all the any banner cardholder customer groups for this offer 
  'If (BannersEnabled) Then
  '  MyCommon.QueryStr = "select CustomerGroupID, Name from CustomerGroups with (NoLock) " & _
  '                      "where BannerID in (select BannerID from BannerOffers with (NoLock) where OfferID=" & OfferID & ") and deleted =0;"
  '  rstBannerCGs = MyCommon.LRT_Select
  'End If
  
  ' fetch the name
  MyCommon.QueryStr = "Select Name,IsTemplate,FromTemplate from Offers with (NoLock) where OfferID=" & OfferID
  rst = MyCommon.LRT_Select
  If (rst.Rows.Count > 0) Then
    Name = rst.Rows(0).Item("Name")
    IsTemplate = MyCommon.NZ(rst.Rows(0).Item("IsTemplate"), False)
    FromTemplate = MyCommon.NZ(rst.Rows(0).Item("FromTemplate"), False)
  End If
  
  Send_HeadBegin("term.offer", "term.customercondition", OfferID)
  If (Request.QueryString("mode") = "optout") Then
    Send("<base target='_self'/>")
  End If
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
%>
<script type="text/javascript">
    <%
  MyCommon.QueryStr = "select CustomerGroupID from CustomerGroups with (NoLock) where NewCardholders = 1 and Deleted=0;"
  rst = MyCommon.LRT_Select
  If (rst.Rows.Count > 0) Then
    NewCardholdersID = MyCommon.NZ(rst.Rows(0).Item("CustomerGroupID"), -1)
  Else
    NewCardholdersID = -1
  End If
  Send("var newCardholdersID = " & NewCardholdersID & ";")
    
  MyCommon.QueryStr = "select CustomerGroupID from CustomerGroups with (NoLock) where AnyCAMCardholder=1 and Deleted=0;"
  rst = MyCommon.LRT_Select
  
  If rst.Rows.Count > 0 Then
    AllCAM = MyCommon.NZ(rst.Rows(0).Item("CustomerGroupID"), -1)
  Else
    AllCAM = -1
  End If
  MyCommon.QueryStr = "Select count(*) from CustomerGroups with (NoLock) where deleted=0 and AnyCustomer<>1 and CustomerGroupID <> 2 and NewCardholders=0 and CustomerGroupID is not null and BannerID is null"
  rst = MyCommon.LRT_Select
    
  If (rst IsNot Nothing AndAlso rst.Rows.Count > 0 AndAlso rst.Rows(0)(0) > 0) Then
    Sendb("var exceptlist = new Array(")
    MyCommon.QueryStr = "select CustomerGroupID from CustomerGroups with (NoLock) where (AnyCardholder=1 or AnyCustomer=1) and Deleted=0 order by CustomerGroupID;"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      i = 1
      For Each row In rst.Rows
        If (i > 1) Then Sendb(",")
        Sendb(row.Item("CustomerGroupID"))
        i += 1
      Next
      Send(");")
    End If
  Else
    Sendb("var functionlist = Array(")
    Sendb("""" & StrConv(Copient.PhraseLib.Lookup("term.anycustomer", LanguageID), VbStrConv.ProperCase) & """,")
    Sendb("""" & StrConv(Copient.PhraseLib.Lookup("term.anycardholder", LanguageID), VbStrConv.ProperCase) & """,")
    Send("""" & StrConv(Copient.PhraseLib.Lookup("term.newcardholders", LanguageID), VbStrConv.ProperCase) & """);")
        
    Sendb("var vallist = Array(")
    Sendb("""" & "1"",""2" & """,")
    Send("""" & NewCardholdersID & """);")
        
    Sendb("var exceptlist = new Array(")
    MyCommon.QueryStr = "select CustomerGroupID from CustomerGroups with (NoLock) where (AnyCardholder=1 or AnyCustomer=1) and Deleted=0 order by CustomerGroupID;"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      i = 1
      For Each row In rst.Rows
        If (i > 1) Then Sendb(",")
        Sendb(row.Item("CustomerGroupID"))
        i += 1
      Next
      Send(");")
    Else
      Sendb("var exceptlist = Array(")
      Send("""" & "-99" & """);")
    End If
  End If
     %>
  var isFireFox = (navigator.appName.indexOf('Mozilla') != -1) ? true : false;
  var timer;
  function xmlPostTimer(strURL, mode) {
    clearTimeout(timer);
    timer = setTimeout("xmlhttpPost('" + strURL + "','" + mode + "')", 250);
  }

  function xmlhttpPost(strURL, mode) {
    var xmlHttpReq = false;
    var self = this;

    document.getElementById("searchLoadDiv").innerHTML = '<% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID))%>';

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
    var select = document.getElementById('selected').options;
    var exclude = document.getElementById('excluded').options;

    var opt = 0;
    var selectedGroups = '';
    for (opt = 0; opt < select.length; opt++) {
      if (select[opt] != null) {
        if (selectedGroups != '') {
          selectedGroups += ",";
        }
        selectedGroups += select[opt].value;
      }
    }

    var excludedGroups = '';
    for (opt = 0; opt < exclude.length; opt++) {
      if (exclude[opt] != null) {
        if (excludedGroups != '') {
          excludedGroups += ",";
        }
        excludedGroups += exclude[opt].value;
      }
    }
    return "Mode=" + mode + "&Search=" + document.getElementById('functioninput').value + "&EngineID=" + '<% Sendb(EngineID)%>' + "&AnyCustomerEnabled=true"  +
  "&SelectedGroups=" + selectedGroups + "&ExcludedGroups=" + excludedGroups + "&OfferID=" + '<% Sendb(OfferID)%>' + "&SearchRadio=" + radioString;

      }

      function updatepage(str) {
        if (str.length > 0) {
          if (!isFireFox) {
            document.getElementById("cgList").innerHTML = '<select class="longer" id="functionselect" name="functionselect" size="12" multiple="multiple">' + str + '</select>';
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
            document.getElementById("cgList").innerHTML = '';
          }
          else {
            document.getElementById("functionselect").innerHTML = '<select class="longer" id="functionselect" name="functionselect" size="12" multiple="multiple"></select>';
          }
          document.getElementById("searchLoadDiv").innerHTML = '&nbsp;';
        }
      }
      // This is the javascript array holding the function list
      // The PrintJavascriptArray ASP function can be used to print this array.

      // This is the function that refreshes the list after a keypress.
      // The maximum number to show can be limited to improve performance with



      function removeUsed(bSkipKeyUp) {
        if(!bSkipKeyUp) xmlhttpPost('OfferFeeds.aspx','ConditionCustomerGroups');
        // this function will remove items from the functionselect box that are used in 
        // selected and excluded boxes

        var funcSel = document.getElementById('functionselect');
        var exSel = document.getElementById('selected');
        var elSel = document.getElementById('excluded');
        var i, j;

        for (i = elSel.length - 1; i >= 0; i--) {
          for (j = funcSel.length - 1; j >= 0; j--) {
            if (funcSel.options[j].value == elSel.options[i].value) {
              funcSel.options[j] = null;
            }
          }
        }

        for (i = exSel.length - 1; i >= 0; i--) {
          for (j = funcSel.length - 1; j >= 0; j--) {
            if (funcSel.options[j].value == exSel.options[i].value) {
              funcSel.options[j] = null;
            }
          }
        }
      }


      // this function gets the selected value and loads the appropriate
      // php reference page in the display frame
      // it can be modified to perform whatever action is needed, or nothing
      function handleSelectClick(itemSelected) {
           <%
  CMS.AMS.CurrentRequest.Resolver.AppName = "offer-con-customer.aspx"
  Dim m_offer As CMS.AMS.Contract.IOffer = CMS.AMS.CurrentRequest.Resolver.Resolve(Of CMS.AMS.Contract.IOffer)()
  Dim m_defaultcust As CMS.AMS.Models.CustomerGroup = m_offer.GetOfferDefaultCustomerGroup(OfferID, EngineID)
  Dim CustGroupId As String = ""
  If (m_defaultcust IsNot Nothing) Then
    CustGroupId = m_defaultcust.CustomerGroupID
  End If
           %>
        var IsEligibleConditionExist = '<%= m_offer.IsOfferOptable(OfferID)%>';
        var defaultcusgroupid= '<%=CustGroupId%>'
        var pagemode = '<%= Request.QueryString("mode")%>'
        // RD-AMSPS-548 Card Required Store Attribute Begin
        var elemChkBoxPane = document.getElementById("chkBoxPane");
        var elemCheckBox = document.getElementById("allowStoreACOverride");
        var DefaultValueForAnyCardholderOverride = <% Sendb(MyCommon.Fetch_CM_SystemOption(127))%>;
        // RD-AMSPS-548 Card Required Store Attribute End
      textObj = document.forms[0].functioninput;
      selectObj = document.forms[0].functionselect;
      selectedValue = document.getElementById("functionselect").value;
      if (selectedValue != "") { selectedText = selectObj[document.getElementById("functionselect").selectedIndex].text; }

      selectboxObj = document.forms[0].selected;
      selectedboxValue = document.getElementById("selected").value;
      if (selectedboxValue != "") { 
        selectedboxText = selectboxObj[document.getElementById("selected").selectedIndex].text; 
      }

      excludedbox = document.forms[0].excluded;
      excludedboxValue = document.getElementById("excluded").value;
      if (excludedboxValue != "") { 
        excludeboxText = excludedbox[document.getElementById("excluded").selectedIndex].text; 
      }

      if (itemSelected == "select1") {
        if (selectedValue != "") {
          // add items to selected box
          if (selectedValue == 2) {
            // RD-AMSPS-548 Card Required Store Attribute Begin
            if (elemChkBoxPane != null)
            {
              elemChkBoxPane.style.visibility = 'visible';
              elemChkBoxPane.disabled = false;
              if (DefaultValueForAnyCardholderOverride == 1)
              {
                elemCheckBox.checked = true;
              }
              else
              {
                elemCheckBox.checked = false;
              }
            }
            // RD-AMSPS-548 Card Required Store Attribute End
            document.getElementById('select2').disabled = false;
          }
            // RD-AMSPS-548 Card Required Store Attribute Begin
          else
          {
            if (elemChkBoxPane != null)
            {
              elemChkBoxPane.disabled = true;
              elemChkBoxPane.style.visibility = 'hidden';
              elemCheckBox.checked = false;
            }            
          }           
          // RD-AMSPS-548 Card Required Store Attribute End
          document.getElementById('deselect1').disabled = false;   
          if (selectedValue < 5)
          {              
            document.getElementById('select1').disabled = true;
          }
          for (i = selectboxObj.length - 1; i >= 0; i--) {
            selectboxObj.options[i] = null;
          }
          selectboxObj[selectboxObj.length] = new Option(selectedText, selectedValue);
        }
      }

      if (itemSelected == "deselect1") {
        if (selectedboxValue != "") {
          if(IsEligibleConditionExist == "True" && pagemode != 'optout')
          {
            if(defaultcusgroupid == selectedboxValue)
            {
              alert('<% Sendb(Copient.PhraseLib.Lookup("term.deselectdefaultgroup", LanguageID))%>');
                  return;
                }
              }
            // RD-AMSPS-548 Card Required Store Attribute Begin
              if (selectedboxValue == 2)
              {
                if (elemChkBoxPane != null)
                {
                  elemChkBoxPane.disabled = true;
                  elemChkBoxPane.style.visibility = 'hidden';
                  elemCheckBox.checked = false;
                }
              }
            // RD-AMSPS-548 Card Required Store Attribute End
            // remove items from selected box
              document.getElementById("selected").remove(document.getElementById("selected").selectedIndex)
              if (selectedboxValue == 2) {
                document.getElementById('select1').disabled = false;
                document.getElementById('select2').disabled = true;
                document.getElementById('deselect2').disabled = true;
              }
              if (selectboxObj.length == 0) {
                // nothing in the select box so disable deselect
                document.getElementById('deselect1').disabled = true;
              }

            }
            if (document.getElementById("selected").options.length == 0) {
              document.getElementById('select1').disabled = false;
            }
          }

          if (itemSelected == "select2") {
            if (selectedValue != "") {
              if (selectedValue == newCardholdersID) {
                alert('<% Sendb(Copient.PhraseLib.Lookup("offer-con.newcardholderexcluded", LanguageID))%>');
                     } else if (document.getElementById("selected").options[0] != null && document.getElementById("selected").options[0].value == 1 && selectedValue != 2) {
                       alert('<% Sendb(Copient.PhraseLib.Lookup("offer-con.onlyanycardholder", LanguageID))%>');
                } else if (document.getElementById("selected").options[0] != null && document.getElementById("selected").options[0].value == 2 && selectedValue == 1) {
                  alert('<% Sendb(Copient.PhraseLib.Lookup("offer-con.anycustomerexclude", LanguageID))%>');
                } else {
                  // add items to excluded box
                  excludedbox[excludedbox.length] = new Option(selectedText, selectedValue);
                  if (excludedbox.length == 1) {
                    document.getElementById('select2').disabled = true;
                    // need to disable deselection on the selected box also since we added an exclude
                    document.getElementById('deselect1').disabled = true;
                    document.getElementById('deselect2').disabled = false;
                  }
                }
          }
        }

        if (itemSelected == "deselect2") {
          if (excludedboxValue != "") {
            // remove items from excluded box    
            document.getElementById("excluded").remove(document.getElementById("excluded").selectedIndex)
            if (excludedbox.length == 0) {
              document.getElementById('select2').disabled = false;
              document.getElementById('deselect1').disabled = false;
              document.getElementById('deselect2').disabled = true;
            }
          }
        }

        updateExceptionButtons();

        // remove items from large list that are in the other lists
        removeUsed(false);
        return true;
      }

      function saveForm() {
        var funcSel = document.getElementById('functionselect');
        var exSel = document.getElementById('excluded');
        var elSel = document.getElementById('selected');
        var i, j;
        var selectList = "";
        var excludededList = "";
        var htmlContents = "";

        // assemble the list of values from the selected box
        for (i = elSel.length - 1; i >= 0; i--) {
          if (elSel.options[i].value != "") {
            if (selectList != "") { selectList = selectList + ","; }
            selectList = selectList + elSel.options[i].value;
          }
        }
        for (i = exSel.length - 1; i >= 0; i--) {
          if (exSel.options[i].value != "") {
            if (excludededList != "") { excludededList = excludededList + ","; }
            excludededList = excludededList + exSel.options[i].value;
          }
        }

        document.getElementById("CustomerGroupID").value = selectList;
        document.getElementById("ExcludedID").value = excludededList;

        // alert(htmlContents);
        return true;
      }

      function updateButtons() {
        // RD-AMSPS-548 Card Required Store Attribute Begin
        var elemChkBoxPane = document.getElementById("chkBoxPane");
        // RD-AMSPS-548 Card Required Store Attribute End
        if (document.getElementById('selected').length > 0) {
          if (document.getElementById('selected').options[0].value == 2) {
            // all customers is in the selected box so disable adding another
            //document.getElementById('select1').disabled = true;
            // RD-AMSPS-548 Card Required Store Attribute Begin
            if (elemChkBoxPane != null)
            {
              elemChkBoxPane.style.visibility = 'visible';
              elemChkBoxPane.disabled = false;
            }
            // RD-AMSPS-548 Card Required Store Attribute End

            if (document.forms[0].selected.length == 1) {
              // there is already an selected group so lets disable selection
              document.getElementById('select2').disabled = true;

              if (document.forms[0].excluded.length == 0) {
                // nothing is excluded so allow excluding one

                document.getElementById('select2').disabled = false;
                document.getElementById('deselect2').disabled = true;
                document.getElementById('deselect1').disabled = false;
              }
              else {
                document.getElementById('select2').disabled = true;
                document.getElementById('deselect2').disabled = false;
                document.getElementById('deselect1').disabled = true;
              }
            }
          }
          else {
            // something is selected but its not all customers
            // RD-AMSPS-548 Card Required Store Attribute Begin
            if (elemChkBoxPane != null)
            {
              elemChkBoxPane.disabled = true;
              elemChkBoxPane.style.visibility = 'hidden';
            }
            // RD-AMSPS-548 Card Required Store Attribute End
            //document.getElementById('select1').disabled = true;
            document.getElementById('deselect1').disabled = false;
          }
          if (document.getElementById('selected').options[0].value < 5) 
          {            
            document.getElementById('select1').disabled = true;
          }
        }
          // RD-AMSPS-548 Card Required Store Attribute Begin
        else
        {
          if (elemChkBoxPane != null)
          {
            elemChkBoxPane.disabled = true;
            elemChkBoxPane.style.visibility = 'hidden';
          }
        }
        // RD-AMSPS-548 Card Required Store Attribute End
      }

      function updateExceptionButtons() {
        var exSel = document.getElementById('excluded');
        var elSel = document.getElementById('selected');
        var bEligible = false;

        // check if there already is an excluded group, if so disable select and enable deselect
        if (exSel != null && elSel != null && exSel.options.length == 0) {
          // check if a exception-qualify customer group is in the selected list
          for (var i = 0; i < elSel.options.length && !bEligible; i++) {
            bEligible = isExceptionGroup(elSel.options[i].value)
          }
          document.getElementById('select2').disabled = (bEligible ? false : true);
          document.getElementById('deselect2').disabled = (bEligible || (!bEligible && exSel.options.length == 0) ? true : false);
        } else if (exSel != null && exSel.options.length > 0) {
          document.getElementById('select2').disabled = true;
          document.getElementById('deselect2').disabled = false;
        } else {
          document.getElementById('select2').disabled = true;
          document.getElementById('deselect2').disabled = true;
        }
      }

      function isExceptionGroup(groupID) {
        var bRetVal = false;

        for (var i = 0; i < exceptlist.length && !bRetVal; i++) {
          bRetVal = (exceptlist[i] == groupID)
        }

        return bRetVal;
      }

   
      function handleCreateClick(createbtn)
      {
        var alertMessage="";
        if(document.getElementById(createbtn)!= undefined && document.getElementById(createbtn) != null)
        {
          var searchText= document.getElementById('functioninput').value;
          if(searchText != null && searchText!="")
          {
            if (searchText.toLowerCase() == '<% Sendb(Copient.PhraseLib.Lookup("term.anycardholder", LanguageID).ToLower())%>' || searchText.toLowerCase() == '<% Sendb(Copient.PhraseLib.Lookup("term.anycustomer", LanguageID).ToLower())%>' || searchText.toLowerCase() == '<%Sendb(Copient.PhraseLib.Lookup("term.newcardholders", LanguageID).ToLower())%>')
            {
              alertMessage='<% Sendb(Copient.PhraseLib.Lookup("term.enter", LanguageID))%>' +' ' + '<% Sendb(Copient.PhraseLib.Lookup("term.valid", LanguageID).ToLower())%>'+' ' + '<% Sendb(Copient.PhraseLib.Lookup("term.customergroup", LanguageID).ToLower())%>'+' ' + '<% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID).ToLower())%>';
              alert(alertMessage);
              return false;
            }
            else
            {
              xmlhttpPost_CreateGroupOrProgramFromOffer('OfferFeeds.aspx', 'CreateGroupOrProgramFromOffer');
            }
          }
          else
          {
            alertMessage='<% Sendb(Copient.PhraseLib.Lookup("term.enter", LanguageID))%>' +' ' + '<% Sendb(Copient.PhraseLib.Lookup("term.valid", LanguageID).ToLower())%>'+' ' + '<% Sendb(Copient.PhraseLib.Lookup("term.customergroup", LanguageID).ToLower())%>'+' ' + '<% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID).ToLower())%>';
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
        document.getElementById("searchLoadDiv").innerHTML = '<% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID))%>';
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
          return "Mode=" + mode + "&CreateType=CustomerGroup&Name=" + document.getElementById('functioninput').value
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
                var selectedGroupValue=-1;
                if(document.getElementById('selected').options[0] !=undefined && document.getElementById('selected').options[0]!=null)
                  selectedGroupValue= document.getElementById('selected').options[0].value
                if(parseInt(selectedGroupValue) != parseInt(resultArr[1]))
                {
                  alert(responseArr[2]);
                  if(selectedGroupValue != -1 ){
                    document.getElementById('selected').options[0].selected = 'selected';
                  }
                  handleSelectClick('deselect1');
                  document.getElementById('functionselect').value=resultArr[1];
                  handleSelectClick('select1') ;
                }
                else
                {
                  alert('<% Sendb(Copient.PhraseLib.Lookup("term.customergroup", LanguageID))%>' + ": '"+ resultArr[0] + "' " + '<% Sendb(Copient.PhraseLib.Lookup("term.is", LanguageID).ToLower())%>'+'  '+'<% Sendb(Copient.PhraseLib.Lookup("term.already", LanguageID).ToLower())%>'+'  '+'<% Sendb(Copient.PhraseLib.Lookup("term.selected", LanguageID).ToLower())%>');
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
    Select Case CustomerGroupID
      Case 0
        MyCommon.QueryStr = "update OfferConditions with (RowLock) set LinkID=0, ExcludedID=0, AllowAnycardholderOverride=0 where ConditionID=" & ConditionID & " and deleted=0;"
        MyCommon.LRT_Execute()
      Case Else
        If (Request.QueryString("allowStoreACOverride") <> "") Then
          ConditionAllowsAnyCardholderOverride = 1
        Else
          ConditionAllowsAnyCardholderOverride = 0
        End If

        MyCommon.QueryStr = "update OfferConditions with (RowLock) set LinkID=" & CustomerGroupID & ", ExcludedId=" & ExcludedID & ", AllowAnycardholderOverride=" & ConditionAllowsAnyCardholderOverride & " where ConditionID=" & ConditionID & " and deleted=0;"
        MyCommon.LRT_Execute()
    End Select

    If (Request.QueryString("IsTemplate") = "IsTemplate") Then
      ' time to update the status bits for the templates
      Dim form_Disallow_Edit As Integer = 0
      Dim form_require_cg As Integer = 0
      
      If (Request.QueryString("Disallow_Edit") = "on") Then
        form_Disallow_Edit = 1
      End If
      If (Request.QueryString("require_cg") <> "") Then
        form_require_cg = 1
      End If
      ' both requiring and locking the customer group is not permitted 
      If (form_require_cg = 1 AndAlso form_Disallow_Edit = 1) Then
        infoMessage = Copient.PhraseLib.Lookup("offer-con.lockeddenied", LanguageID)
      Else
        MyCommon.QueryStr = "update OfferConditions with (RowLock) set Disallow_Edit=" & form_Disallow_Edit & _
        " where ConditionID=" & ConditionID
        MyCommon.LRT_Execute()
        ' update the customer group requirement
        MyCommon.QueryStr = "update OfferConditions with (RowLock) set RequiredFromTemplate = " & _
                            IIf(Request.QueryString("require_cg") <> "", 1, 0) & "where ConditionID = " & ConditionID
        MyCommon.LRT_Execute()
      End If
    End If

    If (infoMessage = "") Then
      CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
    End If
    
    MyCommon.QueryStr = "update OfferConditions with (RowLock) set TCRMAStatusFlag=3,CRMAStatusFlag=2 where ConditionID=" & ConditionID
    MyCommon.LRT_Execute()
    MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where offerid=" & OfferID
    MyCommon.LRT_Execute()
    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-customer", LanguageID))
    
  End If
  
  If (IsTemplate Or FromTemplate) Then
    ' lets dig the permissions if its a template
    MyCommon.QueryStr = "select Disallow_Edit from OfferConditions with (NoLock) where ConditionID=" & ConditionID
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      For Each row In rst.Rows
        ' ok there are some rows for the template
        Disallow_Edit = MyCommon.NZ(row.Item("Disallow_Edit"), True)
      Next
    End If
    'lets dig the customer group requirement for the template
    MyCommon.QueryStr = "select RequiredFromTemplate from OfferConditions with (NoLock) where ConditionID=" & ConditionID
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      RequireCG = MyCommon.NZ(rst.Rows(0).Item("RequiredFromTemplate"), False)
    End If
  End If
  
  Send("<script type=""text/javascript"">")
  Send("function ChangeParentDocument() { ")
  Send("  if (opener != null) {")
  Send("    var newlocation = 'offer-con.aspx?OfferID=" & OfferID & "'; ")
  Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
  Send("     opener.location = 'offer-con.aspx?OfferID=" & OfferID & "'; ")
  Send("  }")
  Send("  }")
  Send("}")
  Send("</script>")
%>
<form action="#" id="mainform" name="mainform" onkeydown="if (event.keyCode == 13) return false;"
  onsubmit="saveForm()">
  <div id="intro">
    <input type="hidden" id="OfferID" name="OfferID" value="<% Sendb(OfferID)%>" />
    <input type="hidden" id="Name" name="Name" value="<% Sendb(Name)%>" />
    <input type="hidden" id="ConditionID" name="ConditionID" value="<% Sendb(ConditionID)%>" />
    <input type="hidden" id="CustomerGroupID" name="CustomerGroupID" value="<% Sendb(CustomerGroupID)%>" />
    <input type="hidden" id="ExcludedID" name="ExcludedID" value="<% Sendb(ExcludedID)%>" />
    <input type="hidden" id="IsTemplate" name="IsTemplate" value="<% 
      If (IsTemplate) Then
        Sendb("IsTemplate")
      Else
        Sendb("Not")
      End If
        %>" />
    <%If (IsTemplate) Then
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.customercondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
      Else
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.customercondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
      End If
    %>
    <div id="controls">
      <% If (IsTemplate) Then%>
      <span class="temp">
        <input type="checkbox" class="tempcheck" id="temp-employees" name="Disallow_Edit"
          <% If (Disallow_Edit) Then Sendb(" checked=""checked""")%> />
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
      <div class="box" id="selector">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.customergroup", LanguageID))%>
          </span>
          <% If (IsTemplate) Then%>
          <span class="tempRequire">
            <input type="checkbox" class="tempcheck" id="require_cg" name="require_cg" <% If (RequireCG) Then Sendb(" checked=""checked""")%> />
            <label for="require_cg">
              <% Sendb(Copient.PhraseLib.Lookup("term.required", LanguageID))%></label>
          </span>
          <% ElseIf (FromTemplate And RequireCG) Then%>
          <span class="tempRequire">*
          <%Sendb(Copient.PhraseLib.Lookup("term.required", LanguageID))%>
          </span>
          <% End If%>
        </h2>
        <input type="radio" id="functionradio1" name="functionradio" <% If (MyCommon.Fetch_SystemOption(175) = "1") Then Sendb(" checked=""checked""")%> <% Sendb(DisabledAttribute)%> /><label
          for="functionradio1"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
        <input type="radio" id="functionradio2" name="functionradio" <% If (MyCommon.Fetch_SystemOption(175) = "2") Then Sendb(" checked=""checked""")%> <% Sendb(DisabledAttribute)%> /><label
          for="functionradio2"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
        <input type="text" class="medium" id="functioninput" name="functioninput" maxlength="100"
          onkeyup="javascript:xmlPostTimer('OfferFeeds.aspx','ConditionCustomerGroups');"
          value="" <% Sendb(DisabledAttribute)%> />
        <% If (bCreateGroupOrProgramFromOffer AndAlso Logix.UserRoles.CreateCustomerGroups) Then%>
        <input class="regular" name="btncreate" id="btncreate" type="button" value="<% Sendb(Copient.PhraseLib.Lookup("term.create", LanguageID))%>" onclick="javascript:handleCreateClick('btncreate');" <% If ((FromTemplate AndAlso Disallow_Edit) OrElse m_offer.IsOfferOptable(OfferID) = True) Then Sendb(" disabled=""disabled""")%> />
        <% End If%>
        <br />
        <div id="searchLoadDiv" style="display: block;">
          &nbsp;
        </div>
        <div id="cgList">
          <select class="longer" id="functionselect" name="functionselect" size="15" <% Sendb(DisabledAttribute)%>>
            <%
                If (EngineID <> 6) Then
                    'add "special" customer groups

                    'see if the offer conditions/rewards allow us to display AnyCustomer group.  (This is not allowed if conditions/rewards require a known customer ex: Points, Stored Value, etc.)
                    MyCommon.QueryStr = "dbo.pa_Check_AnyCustomer_Violation"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@IncentiveID", SqlDbType.BigInt).Value = OfferID
                    Dim anyCustomerDT As DataTable = MyCommon.LRTsp_select
                    MyCommon.Close_LRTsp()
                    If anyCustomerDT.Rows.Count = 0 Then
                        Send("<option value=""1"" style=""color:brown;font-weight:bold;"">" & StrConv(Copient.PhraseLib.Lookup("term.anycustomer", LanguageID), VbStrConv.ProperCase) & "</option>")
                    End If
                    ' dst = Nothing
                    Send("<option value=""2"" style=""color:brown;font-weight:bold;"">" & Copient.PhraseLib.Lookup("term.anycardholder", LanguageID) & "</option>")
                    Send("<option value=""" & NewCardholdersID & """ style=""color:brown;font-weight:bold;"">" & Copient.PhraseLib.Lookup("term.newcardholders", LanguageID) & "</option>")
                Else
                    Send("<option value=""" & AllCAM & """ style=""color:brown;font-weight:bold;"">" & Copient.PhraseLib.Lookup("term.allcam", LanguageID) & "</option>")
                End If
                Dim topString As String = ""
                If (RECORD_LIMIT > 0) Then topString = " top " & RECORD_LIMIT & " "


                'MyCommon.QueryStr = "Select " & topString & " CustomerGroupID,Name from CustomerGroups with (NoLock) where deleted=0 and AnyCustomer<>1 and CustomerGroupID <> 2 and NewCardholders=0 and CustomerGroupID is not null and BannerID is null "
                MyCommon.QueryStr = "SELECT DISTINCT " & topString & " CG.CustomerGroupID, CG.Name " &
                                    "FROM CustomerGroups CG With (NOLOCK) " &
                                    "LEFT JOIN ExtSegmentMap ER on CG.CustomerGroupID = ER.InternalId " &
                                    "WHERE CG.Deleted = 0 And AnyCustomer <> 1 And CustomerGroupID <> 2 " &
                                    "AND CustomerGroupID IS NOT NULL AND BannerID IS NULL " &
                                    "And NewCardholders = 0 And CAMCustomerGroup <> 1 AND CG.Deleted = 0 " &
                                    "AND ((ER.SegmentTypeID IS NULL OR ER.SegmentTypeID = 1 OR ER.SegmentTypeID = 3) AND " &
                                    "(ER.IncentiveId = " & OfferID & " OR ER.ExtSegmentID > 0 OR ER.ExtSegmentID IS NULL)) " &
                                    " And CG.isOptInGroup = 0 order by CG.CustomerGroupID desc, CG.Name;"

                rst = MyCommon.LRT_Select
                For Each row In rst.Rows
                    Send("<option value=""" & row.Item("CustomerGroupID") & """>" & row.Item("Name") & "</option>")
                Next
            %>
          </select>
        </div>
        <%If (RECORD_LIMIT > 0) Then
            Send(Copient.PhraseLib.Lookup("groups.display", LanguageID) & ": " & RECORD_LIMIT.ToString() & "<br />")
          End If
        %>
        <br class="half" />
        <b>
          <% Sendb(Copient.PhraseLib.Lookup("term.selectedcustomers", LanguageID))%>:</b><br />
        <input class="regular select" name="select1" id="select1" type="button" value="&#9660; <% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID))%>"
          onclick="handleSelectClick('select1');" <% Sendb(DisabledAttribute)%> />&nbsp;
      <input class="regular deselect" name="deselect1" id="deselect1" type="button" value="<% Sendb(Copient.PhraseLib.Lookup("term.deselect", LanguageID))%> &#9650;"
        disabled="disabled" onclick="handleSelectClick('deselect1');" <% Sendb(DisabledAttribute)%> /><br />
        <br class="half" />
        <select class="longer" id="selected" name="selected" size="2" <% Sendb(DisabledAttribute)%>>
          <%
            ' find the currently selected groups on page load
            ' RD-AMSPS-548 Card Required Store Attribute Begin
            MyCommon.QueryStr = "select LinkID,ExcludedID,C.Name,AllowAnycardholderOverride from OfferConditions with (NoLock) join CustomerGroups as C with (NoLock) on LinkID=CustomerGroupID and ConditionID=" & ConditionID
            ' RD-AMSPS-548 Card Required Store Attribute End
            If Request.QueryString("mode") = "optout" Then
              MyCommon.QueryStr = MyCommon.QueryStr & " and IsOptInGroup=0 order by C.Name;"
            Else
              MyCommon.QueryStr = MyCommon.QueryStr & " order by C.Name;"
            End If
            rst = MyCommon.LRT_Select
            For Each row In rst.Rows
              If (row.Item("LinkID") = 1) Then
                Send("<option value=""1"" style=""color:brown;font-weight:bold;"">" & StrConv(Copient.PhraseLib.Lookup("term.anycustomer", LanguageID), VbStrConv.ProperCase) & "</option>")
              ElseIf (row.Item("LinkID") = 2) Then
                Send("<option value=""2"" style=""color:brown;font-weight:bold;"">" & Copient.PhraseLib.Lookup("term.anycardholder", LanguageID) & "</option>")
                If (row.Item("AllowAnycardholderOverride")) Then
                  ConditionAllowsAnyCardholderOverride = 1
                Else
                  ConditionAllowsAnyCardholderOverride = 0
                End If
              ElseIf (row.Item("LinkID") = NewCardholdersID) Then
                Send("<option value=""" & NewCardholdersID & """ style=""color:brown;font-weight:bold;"">" & Copient.PhraseLib.Lookup("term.newcardholders", LanguageID) & "</option>")
              ElseIf (row.Item("LinkID") = AllCAM) Then
                Send("<option value=""" & AllCAM & """ style=""color:brown;font-weight:bold;"">" & Copient.PhraseLib.Lookup("term.allcam", LanguageID) & "</option>")
              Else
                Send("<option value=""" & row.Item("LinkID") & """>" & row.Item("Name") & "</option>")
              End If
              SelectedItem = row.Item("LinkID")
            Next
          %>
        </select>
        <br />
        <br class="half" />
        <b>
          <% Sendb(Copient.PhraseLib.Lookup("term.excludedcustomers", LanguageID))%>:</b><br />
        <input class="regular select" id="select2" name="select2" type="button" value="&#9660; <% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID))%>"
          disabled="disabled" onclick="handleSelectClick('select2');" <% Sendb(DisabledAttribute)%> />&nbsp;
      <input class="regular deselect" id="deselect2" name="deselect2" type="button" value="<% Sendb(Copient.PhraseLib.Lookup("term.deselect", LanguageID))%> &#9650;"
        disabled="disabled" onclick="handleSelectClick('deselect2');" <% Sendb(DisabledAttribute)%> /><br />
        <br class="half" />
        <select class="longer" id="excluded" name="excluded" size="2" <% Sendb(DisabledAttribute)%>>
          <%
            ' find the excluded groups on page load
            MyCommon.QueryStr = "select LinkID,ExcludedID,C.Name  from OfferConditions with (NoLock) join CustomerGroups as C with (NoLock) on CustomerGroupID=ExcludedID  where not(ExcludedID=0) and ConditionID=" & ConditionID
            If Request.QueryString("mode") = "optout" Then
              MyCommon.QueryStr = MyCommon.QueryStr & " and IsOptInGroup=0 order by C.Name;"
            Else
              MyCommon.QueryStr = MyCommon.QueryStr & " order by C.Name;"
            End If
            rst = MyCommon.LRT_Select
            For Each row In rst.Rows
              Send("<option value=""" & row.Item("ExcludedID") & """>" & row.Item("Name") & "</option>")
              ExcludedItem = row.Item("ExcludedID")
            Next
          %>
        </select>
        <hr class="hidden" />
      </div>
    </div>
    <div id="gutter">
    </div>
    <% ' RD-AMSPS-548 Card Required Store Attribute Begin %>
    <% 
      If (MyCommon.Fetch_CM_SystemOption(126) = 1) Then
    %>
    <div id="col2" style="display: block;">
      <div class="panel" id="chkBoxPane" style="'margin-top:350px; height: 10%; width: 20%;'">
        <%
          Dim sChecked As String
          sChecked = ""
          If (ConditionAllowsAnyCardholderOverride = 1) Then
            sChecked = " checked=""checked"""
          End If
          Send("<input class=""checkbox""" & sChecked & " id=""allowStoreACOverride""            name=""allowStoreACOverride"" type=""checkbox"" style=""margin-top:380px;"" />")

          Send("<label for=""allowStoreACOverride"">" & Copient.PhraseLib.Lookup("term.AllowAnyCardholderOverride", LanguageID) & "</label><br />")
        %>
      </div>
    </div>
    <%
    End If
    %>
    <% ' RD-AMSPS-548 Card Required Store Attribute End %>
    <div id="column2" style="display: none;">
      <div class="box" id="options">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.advancedoptions", LanguageID))%>
          </span>
        </h2>
      </div>
    </div>
  </div>
</form>
<script type="text/javascript">
<% If (CloseAfterSave) Then%>
  window.close();
<% Else%>
  removeUsed(false);
  updateButtons();
  updateExceptionButtons();
  <% End If%>
</script>
<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd("mainform", "functioninput")
  MyCommon = Nothing
  Logix = Nothing
%>
