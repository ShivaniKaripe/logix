<%@ Page Language="vb" Debug="true" CodeFile="ue-cb.vb" Inherits="UECB" EnableEventValidation="false" %>

<%@ Import Namespace="Copient.CommonIncConfigurable" %>

<%@ Import Namespace="System.ServiceModel.Web" %>

<%@ Import Namespace="System.Web.Services" %>

<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Globalization" %>
<%@ Register Src="../UserControls/ProductAttributeFilter.ascx" TagName="ProductAttributeFilter"
    TagPrefix="uc1" %>
<%  ' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Web.UI.WebControls" %>
<%@ Import Namespace="CMS" %>
<%@ Import Namespace="CMS.AMS" %>
<%@ Import Namespace="CMS.AMS.Contract" %>
<%@ Import Namespace="CMS.AMS.Models" %>
<%@ Import Namespace="ProductGroup" %>
<%
    ' *****************************************************************************
    ' * FILENAME: UEoffer-con-product.aspx
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
%>
<script type="text/javascript">
    // This is the function that refreshes the list after a keypress.
    // The maximum number to show can be limited to improve performance with
    // huge lists (1000s of entries).
    // The function clears the list, and then does a linear search through the
    // globally defined array and adds the matches back to the list.
    var SaveClicked = false;
    var NodeID= "";

    function handleKeyUp(maxNumToShow) {
        var selectObj, textObj, functionListLength;
        var i,  numShown;
        var searchPattern;

        document.getElementById("functionselect").size = "20";

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

        // Loop through the array and re-add matching options
        numShown = 0;
        if (textObj.value == '' && fullSelect != null) {
            var newSelectBox = fullSelect.cloneNode(true);
            document.getElementById('pgList').replaceChild(newSelectBox, selectObj);
            RegisterPGAsyncLoadHandler();  //Attach the scroll event handler after replacement
        } else {
            var newSelectBox = selectObj.cloneNode(false);
            document.getElementById('pgList').replaceChild(newSelectBox, selectObj);
            RegisterPGAsyncLoadHandler();  //Attach the scroll event handler after replacement
            selectObj = document.getElementById("functionselect");
            for(i = 0; i < functionListLength; i++) {
                if(functionlist[i].search(re) != -1) {
                    if (vallist[i] != "") {
                        selectObj[numShown] = new Option(functionlist[i], vallist[i]);
                        selectObj[numShown].title = functionlist[i];
                        if (vallist[i] == 1) {
                            selectObj[numShown].style.fontWeight = 'bold';            
                            selectObj[numShown].style.color = 'brown';
                        }
                        else if((vallist.length > 0 && attributepglist.length > 0) && ($.inArray(parseInt(vallist[i]), attributepglist ) > -1)) {
                            selectObj[numShown].style.color = 'blue';
                        }
                        numShown++;
                    }
                }
                // Stop when the number to show is reached
                if(numShown == maxNumToShow) {
                    break;
                }
            }
        }
        removeUsed(true);
        // When options list whittled to one, select that entry
        if(selectObj.length == 1) {
            selectObj.options[0].selected = true;
        }
        updateButtons();
    }
    function ValidateSave() {
        if ($("#save").length > 0) {
            $('#save').bind("click", function(e) { handleSaveClick(); });
        }
    }
    function handleSaveClick() {
        if(typeof Groupgrid !== "undefined")
            UpdateProductChanges();//For groupgrid update the product and level exclude details
        SaveClicked = true;
    }
    function removeUsed(bSkipKeyUp) {
        if (!bSkipKeyUp) handleKeyUp(99999);
        // this function will remove items from the functionselect box that are used in
        // selected and excluded boxes
        var funcSel = document.getElementById('functionselect');
        var exSel = document.getElementById('selected');
        var elSel = document.getElementById('excluded');
        var i,j;
        for (i = elSel.length - 1; i>=0; i--) {
            for(j=funcSel.length-1;j>=0; j--) {
                if(funcSel.options[j].value == elSel.options[i].value){
                    funcSel.options[j] = null;
                }
            }
        }
        for (i = exSel.length - 1; i>=0; i--) {
            for(j=funcSel.length-1;j>=0; j--) {
                if(funcSel.options[j].value == exSel.options[i].value){
                    funcSel.options[j] = null;
                }
            }
        }
    }

    // this function gets the selected value and loads the appropriate
    // php reference page in the display frame
    // it can be modified to perform whatever action is needed, or nothing
    function handleSelectClick(itemSelected) {
        var isAttributeTypeSwitch = false;

        textObj = document.forms[0].functioninput;

        selectObj = document.forms[0].functionselect;

        selectboxObj = document.forms[0].selected;
        excludedbox = document.forms[0].excluded;
        selectboxObj_attributebased = document.forms[0].functionselect_attr;
        excludedbox_attributebased = document.forms[0].excluded_attr;

        if(itemSelected == "select3" || itemSelected == "deselect3") {
            selectedValue = selectboxObj_attributebased.value;
            if(selectedValue != ""){ selectedText = selectboxObj_attributebased[selectboxObj_attributebased.selectedIndex].text; }
            excludedboxValue = excludedbox_attributebased.value;
            if(excludedboxValue != ""){ excludeboxText = excludedbox[excludedbox_attributebased.selectedIndex].text; }
        }
        else {
            selectedValue = document.getElementById("functionselect").value;
            if(selectedValue != ""){ selectedText = selectObj[document.getElementById("functionselect").selectedIndex].text; }
            selectedboxValue = document.getElementById("selected").value;
            if(selectedboxValue != ""){ selectedboxText = selectboxObj[document.getElementById("selected").selectedIndex].text; }
            excludedboxValue = document.getElementById("excluded").value;
            if(excludedboxValue != ""){ excludeboxText = excludedbox[document.getElementById("excluded").selectedIndex].text; }
        }

        if(itemSelected == "select1") {
            if(selectedValue != "") {
                // add items to selected box
                //    while (selectObj.selectedIndex != -1) {
                //      selectedText = selectObj.options[selectObj.selectedIndex].text;
                //      selectedValue = selectObj.options[selectObj.selectedIndex].value;
                //      selectboxObj[selectboxObj.length] = new Option(selectedText,selectedValue);
                //      selectObj[selectObj.selectedIndex].selected = false;
                //      document.getElementById("select1").disabled = true;
                //    }
                selectedText = selectObj.options[selectObj.selectedIndex].text;
                selectedValue = selectObj.options[selectObj.selectedIndex].value;
                selectboxObj[selectboxObj.length] = new Option(selectedText,selectedValue);
      
                selectObj[selectObj.selectedIndex].selected = false;
                if (selectedValue == "1") {
                    selectboxObj[selectboxObj.length-1].style.color = 'brown';
                    selectboxObj[selectboxObj.length-1].style.fontWeight = 'bold';

        
                }
                else  if(attributepglist.length > 0){
                    if($.inArray(parseInt(selectedValue), attributepglist ) > -1){
                        selectboxObj[selectboxObj.length-1].style.color = 'blue';
                        isAttributeTypeSwitch = true;
                        $("#AttributeSwitchType").val("SelectedAttributeGroup");
                    }
                }
                else{
                    if(parseInt(selectedValue) > -1){
                        //        selectboxObj[selectboxObj.length-1].style.color = 'blue';
                        //        isAttributeTypeSwitch = true;
                        //        $("#AttributeSwitchType").val("SelectedAttributeGroup");
                    }
                }
                document.getElementById("select1").disabled = true;
            }
        }

        if(itemSelected == "deselect1") {
            if(selectedboxValue != "") {
                // remove items from selected box
                while (document.getElementById("selected").selectedIndex != -1) {
                    //AMS-684: Removal of multiple exclusion is not required
<%--        if(selectedboxValue == 1) {
          if (excludedbox.length > 0) {
            if (confirm('<%Sendb(Copient.PhraseLib.Lookup("CPEoffer-con-product.anyproductconfirm", LanguageID))%>')) {
              document.getElementById("excluded").remove(0);
              //Remove items from attributes
             if (excludedbox_attributebased != null) {
                for(var i=0;i<excludedbox_attributebased.options.length;i++)
                  excludedbox_attributebased.options[i] = null;
                  }
              document.getElementById("selected").remove(document.getElementById("selected").selectedIndex);
              document.getElementById("select1").disabled = false;
            } else {
              return;
            }
          } else {
            document.getElementById("selected").remove(document.getElementById("selected").selectedIndex);
            document.getElementById("select1").disabled = false;
          }--%>
            //} else {
            if(attributepglist.length > 0){
                if($.inArray(parseInt(document.getElementById("selected").value), attributepglist ) > -1) {
                    isAttributeTypeSwitch = true;
                    $("#AttributeSwitchType").val("DeSelectedAttributeGroup");
                }
            }
            document.getElementById("selected").remove(document.getElementById("selected").selectedIndex);
            document.getElementById("select1").disabled = false;
            //}
        }
    }
    selectboxObj.options.selectedIndex = -1;
    document.getElementById("select1").disabled = false;
    document.getElementById("IncentiveProductGroupID").value = "0";
}

    if(itemSelected == "select2") {
        if(selectedValue != ""){
            // add items to excluded box
            excludedbox[excludedbox.length] = new Option(selectedText,selectedValue);
            if(attributepglist.length > 0)
            {
                if($.inArray(parseInt(selectedValue), attributepglist ) > -1) {
                    excludedbox[excludedbox.length-1].style.color = 'blue';
                }
            }
            // copy item to excluded box in Attribute Based Section
            if (excludedbox_attributebased != null) {
                excludedbox_attributebased[excludedbox_attributebased.length] = new Option(selectedText,selectedValue);
                excludedbox_attributebased[excludedbox_attributebased.length-1].style.color = excludedbox[excludedbox.length-1].style.color;
                updateButtons_attr();
                removeUsed_attr(true);
            }
        }
    }

    if(itemSelected == "deselect2") {
        if(excludedboxValue != ""){
            // remove items from excluded box
            var selectedIndex = excludedbox.selectedIndex;
            excludedbox.options[selectedIndex] = null;
            // Remove item from excluded box in Attribute Based Section
            if (excludedbox_attributebased != null) {
                excludedbox_attributebased.options[selectedIndex] = null;
                updateButtons_attr();
                removeUsed_attr(false);
            }

        }
    }
    if(itemSelected == "select3") {
        if(selectedValue != ""){
            // add items to excluded box in Attribute Based Section
            excludedbox_attributebased[excludedbox_attributebased.length] = new Option(selectedText,selectedValue);
            if($.inArray(parseInt(selectedValue), attributepglist ) > -1) {
                excludedbox_attributebased[excludedbox_attributebased.length-1].style.color = 'blue';
            }
            // copy item to excluded box in Normal product Group Section
            if (excludedbox != null) {
                excludedbox[excludedbox.length] = new Option(selectedText,selectedValue);
                excludedbox[excludedbox.length-1].style.color = excludedbox_attributebased[excludedbox_attributebased.length-1].style.color;
                updateButtons();
                removeUsed(true);
            }
        }
    }

    if(itemSelected == "deselect3") {
        if(excludedboxValue != ""){
            // remove items from excluded box
            var selectedIndex = excludedbox_attributebased.selectedIndex;
            excludedbox_attributebased.options[selectedIndex] = null;
            // Remove item from excluded box in Normal product Group Section
            if (excludedbox != null) {
                excludedbox.options[selectedIndex] = null;
                updateButtons();
                removeUsed(false);
            }
        }
    }

    if(itemSelected == "select3" || itemSelected == "deselect3") {
        removeUsed_attr(false);
        updateButtons_attr();
        ProductGroupTypeSelection();
    }
    else {
        // remove items from large list that are in the other lists
        removeUsed(false);
        updateButtons();
        if (isAttributeTypeSwitch) {
            $("#mainform").submit();
        }
    }
    ShowOrHideTenderType();
    return true;
}

function ShoworHideDivsWithoutPostback(){
    if($("#radiobtnlistpgselection input[type=radio]:checked").val() == 2 ){
        ShoworHideDivs();
    }
}

function ProductGroupTypeSelection() {
    if ($("#radiobtnlistpgselection").length == 0) {
        document.getElementById('selector').style.display = 'block';
        if(document.getElementById('divAttributeBuilder') != null)
            document.getElementById('divAttributeBuilder').style.display = 'none' ;
        return;
    
    }
    var radiobtn = $("#radiobtnlistpgselection input[type=radio]:checked").val();
    var attributecontainer = $("#filter");
  <% If (String.IsNullOrWhiteSpace(DisabledAttribute)) Then %>
    if(radiobtn == 2 )
        $("#save").removeAttr("disabled")
        //  if (attributecontainer != null && radiobtn == 2 && attributecontainer.children().length > 0) {
        //    if (attributecontainer.children().length > 0) {
        //      $("#save").removeAttr("disabled")
        //    }
        //    else {
        //     $("#save").attr("disabled", "");
        //    }
        //  }
    else {
        updateButtons();
    }
  <% End If %>
    document.getElementById('selector').style.display = (radiobtn == 1 ? 'block' : 'none') ;
    document.getElementById('attributeExcludeGroupSelector').style.display = (radiobtn == 2 ? 'block' : 'none');       
    document.getElementById('divAttributeBuilder').style.display = (radiobtn == 2 ? 'block' : 'none');
 
}
function DisableButton_attr() {
    var selectedItem = document.getElementById('functionselect_attr').value;
    var AttributePGID = document.getElementById('AttributeProductGroupID').value;
    if (selectedItem != 'undefined') {
        if (selectedItem == 1) {
            document.getElementById('select3').disabled = true;
        }
        else if(selectedItem == AttributePGID) {
            document.getElementById('select3').disabled = true;
        }
        else {
            updateButtons_attr();
        }
    }
}
function updateButtons_attr() {
    var elemSelect3 = document.getElementById('select3');
    var elemDeselect3 = document.getElementById('deselect3');
    var functionSelectBoxAttr = document.forms[0].functionselect_attr
    var excludedbox = document.forms[0].excluded_attr;

    if (functionSelectBoxAttr != null)
    {
        if(functionSelectBoxAttr.length == 0)
            elemSelect3.disabled = true;
    }
    if (excludedbox != null) {    
        elemDeselect3.disabled = (excludedbox.length > 0) ? false : true;
    }
}

function handleKeyUp_attr(maxNumToShow) {
    var selectObj, textObj, functionListLength;
    var i,  numShown;
    var searchPattern;
    var AttributePGID = document.getElementById("AttributeProductGroupID").value
    document.getElementById("functionselect_attr").size = "10";

    // Set references to the form elements
    selectObj = document.forms[0].functionselect_attr;
    textObj = document.forms[0].functioninput_attr;

    // Remember the function list length for loop speedup
    functionListLength = functionlist.length;

    // Set the search pattern depending
    if(document.forms[0].functionradio_attr[0].checked == true) {
        searchPattern = "^"+textObj.value;
    } else {
        searchPattern = textObj.value;
    }
    searchPattern = cleanRegExpString(searchPattern);

    // Create a regular expression
    re = new RegExp(searchPattern,"gi");

    // Loop through the array and re-add matching options
    numShown = 0;
    if (textObj.value == '' && fullSelect_attr != null) {
        var newSelectBox = fullSelect_attr.cloneNode(true);
        document.getElementById('pgList_attributebased').replaceChild(newSelectBox, selectObj);
        RegisterPGAsyncLoadHandler();  //Attach the scroll event handler after replacement
    } else {
        var newSelectBox = selectObj.cloneNode(false);
        document.getElementById('pgList_attributebased').replaceChild(newSelectBox, selectObj);
        RegisterPGAsyncLoadHandler();  //Attach the scroll event handler after replacement
        selectObj = document.getElementById("functionselect_attr");
        for(i = 0; i < functionListLength; i++) {
            if(functionlist[i].search(re) != -1) {
                if (vallist[i] != "" && vallist[i] != 1 && vallist[i] != AttributePGID) {
                    selectObj[numShown] = new Option(functionlist[i], vallist[i]);
                    if($.inArray(parseInt(vallist[i]), attributepglist ) > -1) {
                        selectObj[numShown].style.color = 'blue';
                    }
                    numShown++;
                }
            }
            // Stop when the number to show is reached
            if(numShown == maxNumToShow) {
                break;
            }
        }
    }
    removeUsed_attr(true);
    // When options list whittled to one, select that entry
    if(selectObj.length == 1) {
        selectObj.options[0].selected = true;
    }
}

function removeUsed_attr(bSkipKeyUp) {

    if (!bSkipKeyUp) handleKeyUp_attr(99999);
    // this function will remove items from the functionselect box that are used in
    // selected and excluded boxes
    var funcSel = document.getElementById('functionselect_attr');
    var elSel = document.getElementById('excluded_attr');
    var i,j;
    for (i = elSel.length - 1; i>=0; i--) {
        for(j=funcSel.length-1;j>=0; j--) {
            if(funcSel.options[j].value == elSel.options[i].value){
                funcSel.options[j] = null;
            }
        }
    }
}

function saveForm(){
    //if (typeof  spin === 'function' ) 
    //spin('divAttributeBuilder');//PAB
    var dqElem = document.getElementById('Disqualifier');
    var funcSel = document.getElementById('functionselect');
    var exSel = document.getElementById('excluded');
    var elSel = document.getElementById('selected');
    var exSel_attr = document.getElementById('excluded_attr');
    var nodelist = document.getElementById('NodeListID');
    var i,j;
    var selectList = "";
    var excludededList = "";
    var excludedList_attr = ""
    var htmlContents = "";
    var bValidEntry = false;
    var isDisqualifier = false;

    if(typeof idList !="undefined"){
        nodelist.value=idList;
    }
    if (SaveClicked) {
        if (dqElem!=null) { isDisqualifier = (dqElem.value=="1") }
        if(!ValidSave(isDisqualifier)) {
            SaveClicked = false;
            return false;
        }
    }
    // assemble the list of values from the selected box
    for (i = elSel.length - 1; i>=0; i--) {
        if(elSel.options[i].value != ""){
            if(selectList != "") { selectList = selectList + ","; }
            selectList = selectList + elSel.options[i].value;

            bValidEntry = checkEntries(elSel.options[i].value);
            if (!bValidEntry) { SaveClicked = false;return false; }
        }
    }
    for (i = exSel.length - 1; i>=0; i--) {
        if(exSel.options[i].value != ""){
            if(excludededList != "") { excludededList = excludededList + ","; }
            excludededList = excludededList + exSel.options[i].value;
        }
    }
    for (i = exSel_attr.length - 1; i>=0; i--) {
        if(exSel_attr.options[i].value != ""){
            if(excludedList_attr != "") { excludedList_attr = excludedList_attr + ","; }
            excludedList_attr = excludedList_attr + exSel_attr.options[i].value;
        }
    }

    // ok time to build up the hidden variables to pass for saving
    htmlContents = "<input type=\"hidden\" name=\"selGroups\" value=" + selectList + "> ";
    htmlContents = htmlContents + "<input type=\"hidden\" name=\"exGroups\" value=" + excludededList + ">";
    htmlContents = htmlContents + "<input type=\"hidden\" name=\"exGroups_attr\" value=" + excludedList_attr + ">";
    document.getElementById("hiddenVals").innerHTML = htmlContents;
    SaveClicked = false;
    EnableTiers();
    return true;
}
//AMS-1055 - Disabled this function as now grouping of options is taking care of enabling\disabling options
<%--function CheckForSelection(val)
{
    //Any one of the check box should be checked "SameItem" or "Unique".
    //"SameItem" checkbox will be enabled only when UE systemoption #198 is enabled
    if(document.getElementById("SameItem")!=null && document.getElementById("Unique")!=null)
    {
        if(val=="unique")
        {
            if(document.getElementById("SameItem").checked == true)
            {
                document.getElementById("SameItem").checked = false;
             
                ////Minimum Group Purchase GVC
                //document.getElementById("MinPurch").style.display = '';
                  
                ////Minimum Item Price GVC
                //document.getElementById("MinItem").style.display = '';
            }
            if(document.getElementById("NoRestriction").checked == true)
                document.getElementById("NoRestriction").checked = false;
             
            //if(document.getElementById("unique").checked == true)
            //    {
            //    document.getElementById("NoRestriction").checked = false;
            //    document.getElementById("SameItem").checked = false;
            //    }
             
        }
        else if(val=="sameitem")
        {
            if(document.getElementById("Unique").checked == true)
                document.getElementById("Unique").checked = false;
            if(document.getElementById("NoRestriction").checked == true)
                document.getElementById("NoRestriction").checked = false;
            //if (document.getelementbyid("sameitem").checked == false)
            //{
            //    //minimum group purchase
            //    document.getelementbyid("minpurchamt").value = '0';
            //    document.getelementbyid("minpurch").style.display = '';
            //    document.getelementbyid("minpurchlbl").style.display = '';
            //    document.getelementbyid("minpurchcurrsymbol").style.display = '';
            //    document.getelementbyid("minpurchamt").style.display = '';
            //    document.getelementbyid("minpurchcurrabbrev").style.display = '';
              
            //    //minimum item price
            //    document.getelementbyid("minitemprice").value = '0';
            //    document.getelementbyid("minitem").style.display = '';
            //    document.getelementbyid("minitemlbl").style.display = '';
            //    document.getelementbyid("minitemcurrsymbol").style.display = '';
            //    document.getelementbyid("minitemprice").style.display = '';
            //    document.getelementbyid("minitemcurrabbrev").style.display = '';
            //}
            //else 
          
            //    {
            //    if(confirmsameitem())
            //    {
            //        //minimum group purchase gvc
            //        //document.getelementbyid("minpurch").style.display = 'none';
                  
            //        //minimum item price gvc
            //       // document.getelementbyid("minitem").style.display = 'none';
            //    }
            //    else
            //    {
            //        document.getelementbyid("sameitem").checked = false;
            //        checkforselection("sameitem");
            //    }
            //}

        }
        else if(val=="NoRestriction")
        {
            if(document.getElementById("Unique").checked == true)
                document.getElementById("Unique").checked = false;
            if(document.getElementById("SameItem").checked == true)
                document.getElementById("SameItem").checked = false;
        }
    }
}

function ConfirmSameItem()
{
    //var text = 
    if(confirm('<%Sendb(Copient.PhraseLib.Lookup("confirm.sameitem", LanguageID))%>'))
    {
        return true;
    }
    else{
        return false;
    } 
    //'Same Item has been selected. This Product Condition type will restrict some Reward Discount options. Click OK to continue.'
}--%>


function updateButtons() {
    var elemSelect1 = document.getElementById('select1');
    var elemSelect2 = document.getElementById('select2');
    var elemDeselect1 = document.getElementById('deselect1');
    var elemDeselect2 = document.getElementById('deselect2');
    var elemSave = document.getElementById('save');

    var functionSelectList = document.forms[0].functionselect;
    var selectboxObj = document.forms[0].selected;
    var excludedbox = document.forms[0].excluded;
    var isAnyProductSelected = false;
    var isAnyProduct = false;

    if (selectboxObj != null) {
        elemDeselect1.disabled = (selectboxObj.length == 0) ? true : false;
        elemSelect1.disabled = (selectboxObj.length > 0) ? true : false;
        if (selectboxObj.length == 0) {
            if (document.getElementById('require_pg') != null) {
                if (document.getElementById('require_pg').checked == true) {
                    if(elemSave!=null) elemSave.disabled = false;
                } else {
                    if(elemSave!=null) elemSave.disabled = true;
                }
            } else {
                if(elemSave!=null) elemSave.disabled = false;
            }
        } else {
            if(elemSave!=null) elemSave.disabled = false;
        }
        for (var i=0; i < selectboxObj.length; i++) {
            if (selectboxObj.options[i].value == '1') {
                isAnyProductSelected = true;
            }
        }
    } else {
        elemSelect1.disabled = false;
    }

    if(functionSelectList != null)
    {
        for (var i=0; i < functionSelectList.length; i++) {
            if (functionSelectList.options[i].value == '1' && functionSelectList.options[i].selected) {
                isAnyProduct = true;
            }
        }
        if(functionSelectList.length == 0 || functionSelectList.selectedItem=="1" )
            elemSelect2.disabled = true; 
        else 
            elemSelect2.disabled = false; 
    }
    if (excludedbox != null) {
        elemDeselect2.disabled = (excludedbox.length > 0) ? false : true;
       
    }
    if( isAnyProduct )
    {
        elemSelect2.disabled = true;
    }
  <%
  Dim m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
    If Not IsTemplate Then
      If Not (Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer  And Not(FromTemplate And Disallow_Edit)) Then
            Send("  disableAll();")
      End If
    Else
      If Not (Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer) Then
            Send("  disableAll();")
      End If
    End If
  %>
}

//function xmlhttpPost(strURL) {
//  var xmlHttpReq = false;
//  var self = this;

////  document.getElementById("results").innerHTML = "<div class=\"loading\"><img id=\"clock\" src=\"../images/clock22.png\" \ /><br \ />" + '<% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID)) %><\/div>';
    ////  handleSaveButton(true);

    //  // Mozilla/Safari
    //  if (window.XMLHttpRequest) {
    //    self.xmlHttpReq = new XMLHttpRequest();
    //  }
    //  // IE
    //  else if (window.ActiveXObject) {
    //    self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
    //  }
    //  self.xmlHttpReq.open('POST', strURL+"?NodeID="+NodeID, true);
    //  self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    //  self.xmlHttpReq.onreadystatechange = function() {
    //    if (self.xmlHttpReq !=null && self.xmlHttpReq.readyState == 4) {
    ////      updatepage(self.xmlHttpReq.responseText);
    //    }
    //  }

    //  self.xmlHttpReq.send(getquerystring());
    //}

    function getquerystring() {
        var elSel = document.getElementById('selected');
        var i;
        var selectList = "";

        // assemble the list of values from the selected box
        for (i = elSel.length - 1; i>=0; i--) {
            if(elSel.options[i].value != ""){
                if(selectList != "") { selectList = selectList + ","; }
                selectList = selectList + elSel.options[i].value;
            }
        }
        qstr = '<%Sendb("LanguageID=" & LanguageID)%>' + '&CPEProductConditionLimits=' + escape(selectList) + '&RewardOptionID=' + document.getElementById('roid').value + '&Disqualifier=' + document.getElementById('Disqualifier').value+'&NodeID='+NodeID;  // NOTE: no '?' before querystring
  return qstr;
}

function updatepage(str) {
    document.getElementById("results").innerHTML = str;
    handleSaveButton(false);
}

function handleSaveButton(bDisabled) {
    var elemSave = document.getElementById("save");

    if (elemSave != null) {
        elemSave.disabled = bDisabled
    }
}

function checkEntries(pgID) {
    var bValidEntry = false;
    var invalidElemID = '';
    var elemInvalid = null;

    // check for invalid entries in the textboxes
    bValidEntry = isValidEntry('limit', 'select');
    if (!bValidEntry) { invalidElemID = 'limit'; }
    bValidEntry = bValidEntry && isValidEntry('accummin', 'select');
    if (!bValidEntry && invalidElemID == '') { invalidElemID = 'accummin'; }
    bValidEntry = bValidEntry && isValidEntry('accumlimit', 'select');
    if (!bValidEntry && invalidElemID == '') { invalidElemID = 'accumlimit'; }
    bValidEntry = bValidEntry && isValidEntry('accumperiod', 'select');
    if (!bValidEntry && invalidElemID == '') { invalidElemID = 'accumperiod'; }

    if (!bValidEntry) {
        alert('<%Sendb(Copient.PhraseLib.Lookup("term.invalidnumericentry", LanguageID))%>');
      elemInvalid = document.getElementById(invalidElemID);
      if (elemInvalid != null) {
          elemInvalid.focus();
          elemInvalid.select();
      }
  }

    return bValidEntry;
}

function isValidEntry(elemID, elemType) {
    var bValid = true;
    var elem = document.getElementById(elemID);
    var elemUnitType = document.getElementById(elemType);
    var unitType = 0;

    if (elem != null) {
        bValid = !isNaN(elem.value);
        if (bValid && elemUnitType != null) {
            unitType = elemUnitType.value;
            if (unitType == "1") {  //Items
                if (isInteger(elem.value) || elem.value != Math.round(elem.value)) {
                    bValid = isInteger(elem.value);
                }
            } else if ((unitType == "2")||(unitType == "4")) {  //Currency
                if ((decimalPlaces(elem.value, '.') > <% sendb(CurrencyPrecision) %> && elem.value != (Math.round(elem.value*<% sendb(10^CurrencyPrecision) %>)/<% sendb(10^CurrencyPrecision)%>)) || parseFloat(elem.Value) < 0) {
            bValid = false;
        }
} else if (unitType == "3") {  //Legacy weight-volume
    if ((decimalPlaces(elem.value, '.') > 3 && elem.value != (Math.round(elem.value*1000)/1000)) || parseFloat(elem.Value) < 0) {
        bValid = false;
    }
} else if (unitType == "5") {  //Weight
    if ((decimalPlaces(elem.value, '.') > <% sendb(Localization.GetCached_UOM_Precision(roid, 5, Copient.Localization.UOMUsageEnum.UnitType)) %> && elem.value != (Math.round(elem.value*<% sendb(10^Localization.GetCached_UOM_Precision(roid, 5, Copient.Localization.UOMUsageEnum.UnitType)) %>) / <% sendb(10^Localization.GetCached_UOM_Precision(roid, 5, Copient.Localization.UOMUsageEnum.UnitType))%>)) || parseFloat(elem.Value) < 0) {
              bValid = false;
          }
      } else if (unitType == "6") {  //Volume
          if ((decimalPlaces(elem.value, '.') > <% sendb(Localization.GetCached_UOM_Precision(roid, 6, Copient.Localization.UOMUsageEnum.UnitType)) %> && elem.value != (Math.round(elem.value*<% sendb(10^Localization.GetCached_UOM_Precision(roid, 5, Copient.Localization.UOMUsageEnum.UnitType)) %>) / <% sendb(10^Localization.GetCached_UOM_Precision(roid, 5, Copient.Localization.UOMUsageEnum.UnitType))%>)) || parseFloat(elem.Value) < 0) {
              bValid = false;
          }
      } else if (unitType == "7") {  //length
          if ((decimalPlaces(elem.value, '.') > <% sendb(Localization.GetCached_UOM_Precision(roid, 7, Copient.Localization.UOMUsageEnum.UnitType)) %> && elem.value != (Math.round(elem.value*<% sendb(10^Localization.GetCached_UOM_Precision(roid, 5, Copient.Localization.UOMUsageEnum.UnitType)) %>) / <% sendb(10^Localization.GetCached_UOM_Precision(roid, 5, Copient.Localization.UOMUsageEnum.UnitType))%>)) || parseFloat(elem.Value) < 0) {
              bValid = false;
          }
      } else if (unitType == "8") {  //SurfaceArea
          if ((decimalPlaces(elem.value, '.') > <% sendb(Localization.GetCached_UOM_Precision(roid, 8, Copient.Localization.UOMUsageEnum.UnitType)) %> && elem.value != (Math.round(elem.value*<% sendb(10^Localization.GetCached_UOM_Precision(roid, 5, Copient.Localization.UOMUsageEnum.UnitType)) %>) / <% sendb(10^Localization.GetCached_UOM_Precision(roid, 5, Copient.Localization.UOMUsageEnum.UnitType))%>)) || parseFloat(elem.Value) < 0) {
              bValid = false;
          }

      }
}
}

return bValid;
}

function ValidateEntry(evt, src) {
    var elem = document.getElementById("valuetype");

    if (elem != null) {
        if (elem.value == "1") {
            return NumberCheck(evt, src, false);
        } else {
            return NumberCheck(evt, src, true);
        }
    }
}

function NumberCheck(evt, src, allowDecimal) {
    var nkeycode=(window.event) ? window.event.keyCode : evt.which;
    var exceptionKeycodes = new Array(8, 9, 46, 13, 16, 36);

    if (nkeycode >= 48 && nkeycode <= 57) {
        return true;
    } else {
        for (var i=0; i < exceptionKeycodes.length; i++) {
            if (nkeycode == exceptionKeycodes[i]) {
                return true;
            }
        }
        if (allowDecimal == true && nkeycode == 190) {
            if (src !=null && src.value.indexOf(".") < 0) {
                return true;
            }
        }
        return false;
    }
}

function ReformatForQty(isTiered) {
    var elem = document.getElementById("valuetype");
    var txtElem = null;
    var maxCt = 999, decPt = -1;
    var txtVal = "";

    if (elem != null) {
        if (elem.value == "1") {
            var i = (isTiered) ? 1 : 0;
            while (i < maxCt) {
                txtElem = document.getElementById("tier" + i);
                if (txtElem != null) {
                    txtVal = txtElem.value;
                    decPt = txtVal.indexOf(".");
                    if (decPt >= 0) {
                        txtElem.value = txtVal.substring(0, decPt);
                    }
                } else {
                    break;
                }
                i++;
            }
        }
    }
}

function ValidSave(isDisqualifier){
    var retVal = true;
    var elem = document.getElementById("selected");
    var unitElem = document.getElementById("select");
    var qtyElem = document.getElementById("t1_limit");
    var minElem = document.getElementById("MinPurchAmt");
    var minItemElem = document.getElementById("MinItemPrice");
    var elemProgram = document.getElementById("IncentiveProductGroupID");
    var IsNormalPGCondition = (($("#radiobtnlistpgselection").length == 0) || ($("#radiobtnlistpgselection input[type=radio]:checked").val() == 1));
    var msg = '';
    var t = 1;
    var unitType = 1;

    if (elem != null && elem.options.length == 0 && IsNormalPGCondition) {
        if (document.getElementById('require_pg') != null && document.getElementById('require_pg').checked == false) {
            retVal = false;
            msg += '* <% Sendb(Copient.PhraseLib.Lookup("CPEoffer-con-product.selectgroup", LanguageID)) %>'
      elem.focus();
  }
  else if(elemProgram.value == "0"){
        <% If(Not isTemplate) then %>
        var maindiv = document.getElementById('main'); 
        var testdiv = "<div id='infobar' class='red-background'></div>"; 
        var infodiv = document.getElementById('infobar'); 
        if(infodiv == null) 
        { 
            main.innerHTML = testdiv + main.innerHTML; 
        }
        document.getElementById('infobar').innerHTML = '<% Sendb(Copient.PhraseLib.Lookup("reward.groupselect", LanguageID)) %>';
      return;
        <% End If%>
    }
  } 

    if (!isDisqualifier) {
        if (unitElem != null) {
            unitType = parseInt(unitElem.value);
        }

        while (qtyElem != null) {
            // trim the string
            var qtyVal = qtyElem.value.replace(/^\s+|\s+$/g, '');
            qtyVal = qtyVal.replace(",", ".");
            if (unitType==1 && (!isInt(qtyVal) || parseInt(qtyVal)<= 0) || parseInt(qtyVal)>= 1000000000) {
                retVal = false;
                if (msg != '') { msg += '\n\r\n\r'; }
                msg += '* <% Sendb(Copient.PhraseLib.Lookup("CPEoffer-con-point.positiveinteger", LanguageID)) %>';
        qtyElem.focus();
        qtyElem.select();
        break;
    } else if ((unitType!=1 && (qtyVal=="" || isNaN(qtyVal))) || (!isNaN(qtyVal) && unitType!=1 && parseFloat(qtyVal)<= 0)) {
        retVal = false;
        if (msg != '') { msg += '\n\r\n\r'; }
        msg += '* <% Sendb(Copient.PhraseLib.Lookup("CPEoffer-con-point.positivedecimal", LanguageID)) %>';
        qtyElem.focus();
        qtyElem.select();
        break;
    }
    t++;
    qtyElem = document.getElementById("t" + t + "_limit");
}

      //Check for Minimum Purchase Amount
    var minVal = minElem.value.replace(/^\s+|\s+$/g, '');
    minVal = minVal.replace(",", ".");
    if ((minVal=="" || isNaN(minVal)) || (!isNaN(minVal) && parseFloat(minVal)< 0)) {
        retVal = false;
        if (msg != '') { msg += '\n\r\n\r'; }
        msg += '* <% Sendb(Copient.PhraseLib.Lookup("CPEoffer-con-point.positivedecimal", LanguageID)) %>';
        minElem.focus();
        minElem.select();
    }

    if (minItemElem != null) {
        // Check for Minimum Item Price
        minVal = minItemElem.value.replace(/^\s+|\s+$/g, '');
        minVal = minVal.replace(",", ".");
        if (isNaN(minVal) || (!isNaN(minVal) && parseFloat(minVal)< 0)) {
            retVal = false;
            if (msg != '') { msg += '\n\r\n\r'; }
            msg += '* <% Sendb(Copient.PhraseLib.Lookup("CPEoffer-con-point.positivedecimal", LanguageID)) %>';
          minItemElem.focus();
          minItemElem.select();
      }
  }

}

    if (msg != '') {
        alert(msg);
    }
    return retVal;
}

function SetTiersToValue(checked){
    var qtyElem = document.getElementById("t1_limit");
    var firstTierValue = qtyElem.value
    qtyElem.disabled = checked
    var t = 2
    qtyElem = document.getElementById("t" + t + "_limit");
  
    while (qtyElem != null) 
    {
        if (checked) 
        {
            qtyElem.value = firstTierValue
            qtyElem.disabled = true
        }
        else 
        {
            qtyElem.disabled = false
        }
        t++;
        qtyElem = document.getElementById("t" + t + "_limit");
    }
}

function EnableTiers(){
    var t = 1
    qtyElem = document.getElementById("t" + t + "_limit");
  
    while (qtyElem != null) 
    {
        qtyElem.disabled = false
        t++;
        qtyElem = document.getElementById("t" + t + "_limit");
    }
}

function isInt(sNum) {
    return (sNum!="" && !isNaN(sNum) && (sNum/1)==parseInt(sNum));
}
</script>
<%
    Send("<script " & "type=""text/javascript"">")
    Send("function ChangeUnit(Type) { ")
    Send("  var prod = document.getElementById('Unique');")
    Send("  var same = document.getElementById('SameItem');")
    Send("  var nor = document.getElementById('NoRestriction');")
    Send("  var accum = document.getElementById('EnableAccum');")
    Send("  var accumlabel = document.getElementById('AccumChecklbl');")
    Send("  var rounding = document.getElementById('rounding');")
    Send("  var roundingspan = document.getElementById('roundingspan');")
    Send("  var minpurch = document.getElementById('MinPurch');")
    Send("  var minpurchamt = document.getElementById('MinPurchAmt');")
    Send("  var minitem = document.getElementById('MinItem');")
    Send("  var minitemprice = document.getElementById('MinItemPrice');")
    Send("  var accumulation = document.getElementById('accumulation');")
    Send("  var unittypedesc = document.getElementById('UnitTypeDesc');")
    Send("  ShowOrHideTenderType();")
    For t = 1 To TierLevels
        Send("  var currsymbolt" & t & " = document.getElementById('currsymbolt" & t & "');")
        Send("  var currabbrevt" & t & " = document.getElementById('currabbrevt" & t & "');")
        Send("  var weightabbrevt" & t & " = document.getElementById('weightabbrevt" & t & "');")
        Send("  var volumeabbrevt" & t & " = document.getElementById('volumeabbrevt" & t & "');")
        Send("  var lengthabbrevt" & t & " = document.getElementById('lengthabbrevt" & t & "');")
        Send("  var surfareaabbrevt" & t & " = document.getElementById('surfareaabbrevt" & t & "');")
    Next t

    Send("  //turn off all of the input box lables")
    For t = 1 To TierLevels
        Send("    if (currsymbolt" & t & " != null) { currsymbolt" & t & ".style.display = 'none'; }")
        Send("    if (currabbrevt" & t & " != null) { currabbrevt" & t & ".style.display = 'none'; }")
        Send("    if (weightabbrevt" & t & " != null) { weightabbrevt" & t & ".style.display = 'none'; }")
        Send("    if (volumeabbrevt" & t & " != null) { volumeabbrevt" & t & ".style.display = 'none'; }")
        Send("    if (lengthabbrevt" & t & " != null) { lengthabbrevt" & t & ".style.display = 'none'; }")
        Send("    if (surfareaabbrevt" & t & " != null) { surfareaabbrevt" & t & ".style.display = 'none'; }")
    Next t

    Send("  if (Type == ""1"") {") 'items
    Send("    if (accum == null || accum.checked == false) {")
    Send("      prod.disabled = false; ")
    Send("      same.disabled = false; ")
    Send("      nor.disabled = false; ")
    Send("    }")
    Send("    if (roundingspan != null) { roundingspan.style.display = 'none'; }")
    Send("    if (rounding != null) { rounding.checked = false; } ")
    Send("    if (accum != null) { accum.style.display = 'inline'; }")
    Send("    if (accumlabel != null) { accumlabel.style.display = 'inline'; }")
    Send("    if (minpurch != null) { minpurch.style.display = ''; }")
    Send("    if (minitem != null) { minitem.style.display = ''; }")
    Send("    if (unittypedesc != null) { unittypedesc.innerHTML = '" & Copient.PhraseLib.Lookup("term.quantity", LanguageID) & "'; }")
    Send("  } else if (Type == ""4"" || Type == ""10"") {") 'Qty1 at price or Price Threshold 'vtopol_DT13
    Send("    if (prod != null) { prod.checked = false; }")
    Send("    if (prod != null) { prod.disabled = true; }")
    Send("    if (same != null) { same.checked = false; }")
    Send("    if (same != null) { same.disabled = true; }")
    Send("    if (nor != null) { nor.checked = true; }")
    'Send("    if (nor != null) { nor.disabled = true; }")
    Send("    if (roundingspan != null) { roundingspan.style.display = 'none'; }")
    Send("    if (rounding != null) { rounding.checked = false; }")
    Send("    if (accum != null) { accum.style.display = 'none'; }")
    Send("    if (accumlabel != null) { accumlabel.style.display = 'none'; }")
    Send("    if (minpurch != null) { minpurch.style.display = 'none'; }")
    Send("    if (minpurchamt != null) { minpurchamt.value = '0'; }")
    Send("    if (minitem != null) { minitem.style.display = 'none'; }")
    Send("    if (minitemprice != null) { minitemprice.value = '0'; }")
    Send("    if (accum != null) { accum.checked = false; }")
    Send("    if (accumulation != null) { accumulation.style.display = 'none'; }")
    Send("    if (unittypedesc != null) { unittypedesc.innerHTML = '" & Copient.PhraseLib.Lookup("term.price", LanguageID) & "'; }")
    For t = 1 To TierLevels
        Send("    if (currsymbolt" & t & " != null) { currsymbolt" & t & ".style.display = ''; }")
        Send("    if (currabbrevt" & t & " != null) { currabbrevt" & t & ".style.display = ''; }")
    Next t
    Send("  } else {")
    Send("    if (prod != null) { prod.checked = false; }")
    Send("    if (prod != null) { prod.disabled = true; }")
    Send("    if (same != null) { same.checked = false; }")
    Send("    if (same != null) { same.disabled = true; }")
    Send("    if (nor != null) { nor.checked = true; }")
    'Send("    if (nor != null) { nor.disabled = true; }")
    Send("    if (roundingspan != null) { roundingspan.style.display = 'inline'; }")
    Send("    if (rounding != null) { rounding.style.display = 'inline'; }")
    Send("    if (accum != null) { accum.style.display = 'inline'; }")
    Send("    if (accumlabel != null) { accumlabel.style.display = 'inline'; }")
    Send("    if (minitem != null) { minitem.style.display = ''; }")
    Send("    if (minpurch != null) { minpurch.style.display = ''; }")
    Send("    if (Type == ""2"") {")  'Currency
    Send("      if (unittypedesc != null) { unittypedesc.innerHTML = '" & Copient.PhraseLib.Lookup("term.price", LanguageID) & "'; }")
    For t = 1 To TierLevels
        Send("      if (currsymbolt" & t & " != null) { currsymbolt" & t & ".style.display = ''; }")
        Send("      if (currabbrevt" & t & " != null) { currabbrevt" & t & ".style.display = ''; }")
    Next t
    Send("    } else  {")
    Send("      if (unittypedesc != null) { unittypedesc.innerHTML = '" & Copient.PhraseLib.Lookup("term.quantity", LanguageID) & "'; }")
    Send("      if (Type == ""5"") { ") 'Weight
    For t = 1 To TierLevels
        Send("        if (weightabbrevt" & t & " != null) { weightabbrevt" & t & ".style.display = ''; }")
    Next t
    Send("      } else if (Type == ""6"") { ") 'Volume
    For t = 1 To TierLevels
        Send("        if (volumeabbrevt" & t & " != null) { volumeabbrevt" & t & ".style.display = ''; }")
    Next t
    Send("      } else if (Type == ""7"") { ") 'Length
    For t = 1 To TierLevels
        Send("        if (lengthabbrevt" & t & " != null) { lengthabbrevt" & t & ".style.display = ''; }")
    Next t
    Send("      } else if (Type == ""8"") { ") 'Surface area
    For t = 1 To TierLevels
        Send("        if (surfareaabbrevt" & t & " != null) { surfareaabbrevt" & t & ".style.display = ''; }")
    Next t
    Send("      } ")
    Send("    }")
    Send("  }")
    Send("}")
    Send("</script>")

%>
<script type="text/javascript">

    function handleRequiredToggle(cb) {
        //Update other required checkbox also.
        if (cb.id == "require_pg")
            $("#require_pg")[0].checked = cb.checked;
        else
            $("#require_pg_Attr")[0].checked = cb.checked;

        if (document.forms[0].selected.length == 0) {
            if (document.getElementById("require_pg").checked == false) {
                document.getElementById('save').disabled = true;
            } else {
                document.getElementById('save').disabled = false;
            }
        }
        if ($("#require_pg")[0].checked == true || $("#require_pg_Attr")[0].checked==true) {
            document.getElementById("Disallow_Edit").checked = false;
        }
    }

    function disableAll() {
        document.getElementById('select1').disabled = true;
        document.getElementById('select2').disabled = true;
        document.getElementById('deselect1').disabled = true;
        document.getElementById('deselect2').disabled = true;
        document.getElementById('functionselect').disabled = true;
        document.getElementById('selected').disabled = true;
        document.getElementById('excluded').disabled = true;
    }
    function DisableButton() {
        var selectedItem = document.getElementById('functionselect').value;
        if (selectedItem != 'undefined') {
            if (selectedItem == 1)
            { document.getElementById('select2').disabled = true; }
            else
            { document.getElementById('select2').disabled = false; }
        }
    }
  
</script>
<%

    Send("<script type=""text/javascript"">")

    Send("function toggleAccum() {")
    Send("  var elem = document.getElementById(""EnableAccum"");")
    Send("  var elemDiv = document.getElementById(""accumulation"");")
    Send("  var elemMin = document.getElementById(""accummin"");")
    Send("  var elemLimit = document.getElementById(""accumlimit"");")
    Send("  var elemPeriod = document.getElementById(""accumperiod"");")
    Send("  var elemMinPurch = document.getElementById(""MinPurchAmt"");")
    Send("  var elemMinPurchlbl = document.getElementById(""MinPurchlbl"");")
    Send("  var elemMinItem = document.getElementById(""MinItemPrice"");")
    Send("  var elemMinItemlbl = document.getElementById(""MinItemlbl"");")
    Send("  var elemMinItemCurrSymbol = document.getElementById(""MinItemCurrSymbol"");")
    Send("  var elemMinPurchCurrSymbol = document.getElementById(""MinItemPurchCurrSymbol"");")
    Send("  var elemMinItemCurrAbbrev = document.getElementById(""MinItemCurrAbbrev"");")
    Send("  var elemMinPurchCurrAbbrev = document.getElementById(""MinPurchCurrAbbrev"");")

    Send("  if (elem != null && elemDiv != null) { ")
    Send("    if (elem.checked) { ")
    Send("      elemDiv.style.display = '';")
    Send("      elemMinPurch.style.display = 'none';")
    Send("      elemMinPurchlbl.style.display = 'none';")
    Send("      elemMinItem.style.display = 'none';")
    Send("      elemMinItemlbl.style.display = 'none';")
    Send("      elemMinItemCurrSymbol.style.display = 'none';")
    Send("      elemMinPurchCurrSymbol.style.display = 'none';")
    Send("      elemMinItemCurrAbbrev.style.display = 'none';")
    Send("      elemMinPurchCurrAbbrev.style.display = 'none';")
    Send("    } else { ")
    Send("      elemDiv.style.display = 'none';")
    Send("      elemMinPurch.style.display = '';")
    Send("      elemMinPurchlbl.style.display = '';")
    Send("      elemMinItem.style.display = '';")
    Send("      elemMinItemlbl.style.display = '';")
    Send("      elemMinItemCurrSymbol.style.display = '';")
    Send("      elemMinPurchCurrSymbol.style.display = '';")
    Send("      elemMinItemCurrAbbrev.style.display = '';")
    Send("      elemMinPurchCurrAbbrev.style.display = '';")
    Send("      if (elemMin != null) { elemMin.value = '0'; }")
    Send("      if (elemLimit != null) { elemLimit.value = '0'; }")
    Send("      if (elemPeriod != null) { elemPeriod.value = '0'; }")
    Send("      if (elemMinPurch != null) { elemMinPurch.value = '0'; }")
    Send("      if (elemMinItem != null) { elemMinItem.value = '0'; }")
    Send("    }")
    Send("  }")
    Send("} ")

    Send("function ChangeAccum() {")
    If Not EnableAccum Then
        Send("  document.location = 'UEoffer-con-product.aspx?OfferID=" & OfferID & "&EnableAccum=1&Disqualifier=" & IIf(Disqualifier, 1, 0) & IIf(IncentiveProdGroupID > 0, "&IncentiveProductGroupID=" & IncentiveProdGroupID, "") & "';")
    ElseIf EnableAccum Then
        Send("  document.location = 'UEoffer-con-product.aspx?OfferID=" & OfferID & "&EnableAccum=0&Disqualifier=" & IIf(Disqualifier, 1, 0) & IIf(IncentiveProdGroupID > 0, "&IncentiveProductGroupID=" & IncentiveProdGroupID, "") & "';")
    End If
    Send("} ")

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
<form action="#" id="mainform" name="mainform"  onsubmit="return saveForm();" runat="server">
<div id="divNotification" style="display:none"><label id="lblAjaxNotification" ></label></div>
<div id="intro">
    <span id="hiddenVals"></span>
    <input type="hidden" id="OfferID" name="OfferID" value="<% sendb(OfferID) %>" />
    <input type="hidden" id="Name" name="Name" value="<% sendb(Name) %>" />
    <input type="hidden" id="ConditionID" name="ConditionID" value="<% Sendb(ConditionID) %>" />
    <input type="hidden" id="IncentiveProductGroupID" name="IncentiveProductGroupID"
        value="<% Sendb(IncentiveProdGroupID) %>" />
    <input type="hidden" id="AttributeProductGroupID" name="AttributeProductGroupID"
        value="<% Sendb(AttributeProductGroupID) %>" />
    <input type="hidden" id="hdnBuyerID" name="hdnBuyerID" value="<%
    If(objOffer Is Nothing OrElse objOffer.BuyerID Is Nothing) Then
      Sendb("0")
    Else
      Sendb(objOffer.BuyerID)
    End If %>" />
    <input type="hidden" id="AttributeSwitchType" name="AttributeSwitchType" value="" />
    
            <input id="NodeListID" type="hidden" name="NodeListID" value="" runat="server" />
    <%
        If Not Disqualifier Then
            Sendb("<input type=""hidden"" id=""Disqualifier"" name=""Disqualifier"" value=""0"" />")
        Else
            Sendb("<input type=""hidden"" id=""Disqualifier"" name=""Disqualifier"" value=""1"" />")
        End If
    %>
    <input type="hidden" id="roid" name="roid" value="<%sendb(roid) %>" />
    <input type="hidden" id="IsTemplate" name="IsTemplate" value="<%
          If (IsTemplate) Then
          Sendb("IsTemplate")
          Else
          Sendb("Not")
          End If
          %>" />
    <%
        If (isTemplate) Then
            Sendb("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID)
        Else
            Sendb("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID)
        End If
        If Not Disqualifier Then
            Send(" " & StrConv(Copient.PhraseLib.Lookup("term.productcondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
        Else
            Send(" " & StrConv(Copient.PhraseLib.Lookup("term.productdisqualifier", LanguageID), VbStrConv.Lowercase) & "</h1>")
        End If
    %>
    <div id="controls">
        <% If (isTemplate) Then%>
        <span class="temp">
            <input type="checkbox" class="tempcheck" id="Disallow_Edit" name="Disallow_Edit"
                <% if(disallow_edit)then sendb(" checked=""checked""") %> />
            <label for="Disallow_Edit">
                <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
            </label>
        </span>
        <% End If%>
        <%
            Dim isTranslatedOffer As Boolean =MyCommon.IsTranslatedUEOffer(OfferID,  MyCommon)
            Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249)="1",True,False)
            
            Dim bEnableAdditionalLockoutRestrictionsOnOffers As Boolean = IIf(MyCommon.Fetch_SystemOption(260) = "1", True, False)
            Dim bOfferEditable As Boolean = MyCommon.IsOfferEditablePastLockOutPeriod(Logix.UserRoles.EditOfferPastLockoutPeriod, MyCommon, OfferID)
            
            Dim m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
              
            If ((Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
              If Not isTemplate Then
                    If (Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And Disallow_Edit) And Not IsOfferWaitingForApproval(OfferID)) Then Send_Save()
              Else
                  If (Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer) Then Send_Save()
              End If
            End If
        %>
    </div>
</div>

<div id="main" >
    <%  If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <% If AttributePGEnabled Then%>
    <div>
        <asp:radiobuttonlist id="radiobtnlistpgselection" name="radiobtnlistpgselection"
            runat="server" />
    </div>
  <%--  <div class="box" id="toolbar"  style="float:left;position: relative; width: 90%; height: 80%;">
        <span style="float: left;">Hi </span><br /><br /><br /><br /><br />
       
    </div>--%>
    <% End If%>
  
    <div style="float: left;position: relative; width: 108%; height: 80%;"">
   
        <div class="box column3x" id="selector" style="overflow: auto;">
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.productgroup", LanguageID))%>
                </span>
                
                <% If (isTemplate) Then%>
                <span class="tempRequire">
                    <input type="checkbox" class="tempcheck" id="require_pg" name="require_pg" onclick="handleRequiredToggle(this);"
                        <% if(RequirePG)then sendb(" checked=""checked""") %> />
                    <label for="require_pg">
                        <% Sendb(Copient.PhraseLib.Lookup("term.required", LanguageID))%>
                    </label>
                </span>
                <% ElseIf (FromTemplate And RequirePG) Then%>
                <span class="tempRequire">
                    <%Sendb("*" & Copient.PhraseLib.Lookup("term.required", LanguageID))%>
                </span>
                <% End If%>
            </h2>
            <input type="radio" id="functionradio1" name="functionradio" <% if(MyCommon.Fetch_SystemOption(175)= "1" ) then sendb(" checked=""checked""") %>
                <% sendb(disabledattribute) %> /><label for="functionradio1"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
            <input type="radio" id="functionradio2" name="functionradio" <% if(MyCommon.Fetch_SystemOption(175)= "2" ) then sendb(" checked=""checked""") %>
                <% sendb(disabledattribute) %> /><label for="functionradio2"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
            <input type="text" class="medium" id="functioninput" name="functioninput" maxlength="100"
                onkeyup="handleKeyUp(200);" value="" <% sendb(disabledattribute) %> /><br />
            <div id="pgList" class="column3x1">
                <select class="long" id="functionselect" onchange="DisableButton()" name="functionselect"
                    size="20" <% sendb(disabledattribute) %>>
                    <%
                        dtAllPGList = GetProductGroupListHTML(MyCommon, bEnableRestrictedAccessToUEOfferBuilder, "", AdminUserID, roid, Disqualifier, Logix.UserRoles.ViewProductgroupRegardlessBuyer, LanguageID, false, shouldFetchPGAsync)
                        Send(PrepareProductGroupHTML(MyCommon, dtAllPGList, LanguageID))
                        'PopulateProductGroupList(dtAllPGList)
                    %>
                </select>
            </div>
            <div class="column3x2">
                <center>
                    <br />
                    <br />
                    <%
                        'AMS-684 removed bUseMultipleProductExclusionGroups
                        If Disqualifier Then
                            Send("<br />")
                            Send("<br />")
                            Send("<br />")
                            Send("<br />")
                            Send("<br />")
                        End If
                    %>
                    <input type="button" class="regular select" id="select1" name="select1" value="<% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID))%> &#9658;"
                        onclick="handleSelectClick('select1');" <% sendb(disabledattribute) %> />
                    <br />
                    <br />
                    <input type="button" class="regular deselect" id="deselect1" name="deselect1" value="&#9668; <% Sendb(Copient.PhraseLib.Lookup("term.deselect", LanguageID))%>"
                        disabled="disabled" onclick="handleSelectClick('deselect1');" /><br />
                    <br />
                    <br />
                    <br />
                    <br />
                    <br />
                    <br />
                    <br />
                    <%
                        'AMS-684 removed bUseMultipleProductExclusionGroups, removed conditions other than disqualifier to enable the select\deselect buttons for exclusion
                        If Disqualifier Then
                            Send("<div style=""display:none;"">")
                        End If
                    %>
                    <input type="button" class="regular select" name="select2" id="select2" value="<% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID))%> &#9658;"
                        onclick="handleSelectClick('select2');" <% sendb(disabledattribute) %> /><br />
                    <br />
                    <br />
                    <input type="button" class="regular deselect" name="deselect2" id="deselect2" value="&#9668; <% Sendb(Copient.PhraseLib.Lookup("term.deselect", LanguageID))%>"
                       onclick="handleSelectClick('deselect2');" <% sendb(disabledattribute) %> /><br />
                    <%
                        'AMS-684 removed bUseMultipleProductExclusionGroups, removed conditions other than disqualifier to enable the select\deselect buttons for exclusion for 
                        If Disqualifier Then
                            Send("</div>")
                        End If
                    %>
                </center>
            </div>
            <div class="column3x3">
                <div class="boxwrap">
                    <h3>
                        <span>
                            <% Sendb(Copient.PhraseLib.Lookup("term.includedgroups", LanguageID))%>
                        </span>
                    </h3>
                </div>
                <select class="long" id="selected" name="selected" multiple="multiple" <%Sendb(IIf(Disqualifier, " size=""18""", " size=""7"""))%>
                    <% sendb(disabledattribute) %>>
                    <%
                        'If ProdID > 0 And ExProdID = -1 Then
                        If IncentiveProdGroupID > 0 AndAlso Not Page.IsPostBack Then
                            ' alright lets find the currently selected groups on page load
                            MyCommon.QueryStr = "select Name,buyerid from ProductGroups where ProductGroupID=" & ProdID
                            rst = MyCommon.LRT_Select
                            If rst.Rows.Count > 0 Then
                                If ProdID = 1 Then
                                    Send("<option value=""1"" style=""color: brown; font-weight: bold;"">" & Copient.PhraseLib.Lookup("term.anyproduct", LanguageID) & "</option>")
                                Else
                                    If (MyCommon.IsEngineInstalled(9) And SystemCacheData.GetSystemOption_UE_ByOptionId(168) = "1" And Not IsDBNull(rst.Rows(0).Item("Buyerid"))) Then
                                        Dim buyerid As Integer = rst.Rows(0).Item("Buyerid")
                                        Dim externalBuyerid = MyCommon.GetExternalBuyerId(buyerid)
                                        Send("<option value=""" & ProdID & """ " & IIf(lstAttributeProductGroups.Contains(ProdID), "style=""color: blue;""", "") & ">" & "Buyer " & externalBuyerid & " - " & MyCommon.NZ(rst.Rows(0).Item("Name"), "") & "</option>")
                                    Else
                                        Send("<option value=""" & ProdID & """ " & IIf(lstAttributeProductGroups.Contains(ProdID), "style=""color: blue;""", "") & ">" & MyCommon.NZ(rst.Rows(0).Item("Name"), "") & "</option>")
                                    End If

                                End If

                            End If

                        Else
                            If GetCgiValue("selGroups") <> "" Then
                                MyCommon.QueryStr = "select Name,buyerid from ProductGroups where ProductGroupID=" & MyCommon.Extract_Val(GetCgiValue("selGroups"))
                                rst = MyCommon.LRT_Select
                                If rst.Rows.Count > 0 Then
                                    If MyCommon.Extract_Val(GetCgiValue("selGroups")) = 1 Then
                                        Send("<option value=""1"" style=""color: brown; font-weight: bold;"">" & Copient.PhraseLib.Lookup("term.anyproduct", LanguageID) & "</option>")
                                    ElseIf (MyCommon.IsEngineInstalled(9) And SystemCacheData.GetSystemOption_UE_ByOptionId(168) = "1" And Not IsDBNull(rst.Rows(0).Item("Buyerid"))) Then
                                        Dim buyerid As Integer = rst.Rows(0).Item("Buyerid")
                                        Dim externalBuyerid = MyCommon.GetExternalBuyerId(buyerid)
                                        Send("<option value=""" & MyCommon.Extract_Val(GetCgiValue("selGroups")) & """ " & IIf(lstAttributeProductGroups.Contains(MyCommon.Extract_Val(GetCgiValue("selGroups"))), "style=""color: blue;""", "") & ">" & "Buyer " & externalBuyerid & " - " & MyCommon.NZ(rst.Rows(0).Item("Name"), "") & "</option>")
                                    Else
                                        Send("<option value=""" & MyCommon.Extract_Val(GetCgiValue("selGroups")) & """ " & IIf(lstAttributeProductGroups.Contains(MyCommon.Extract_Val(GetCgiValue("selGroups"))), "style=""color: blue;""", "") & ">" & MyCommon.NZ(rst.Rows(0).Item("Name"), "") & "</option>")
                                    End If
                                End If
                            End If
                        End If
                    %>
                </select>
                <br />
                <%
                    'AMS-684 removed bUseMultipleProductExclusionGroups and conditions other than disqualifier to enable the excluded group box
                    If Disqualifier Then
                        Send("<div style=""display:none;"">")
                    End If
                %>
                <br class="half" />
                <br class="half" />
                <div class="boxwrap">
                    <h3>
                        <span>
                            <% Sendb(Copient.PhraseLib.Lookup("term.excludedgroups", LanguageID))%>
                        </span>
                    </h3>
                </div>
                  <select class="long" id="excluded" name="excluded" multiple="multiple" <%Sendb(IIf(Disqualifier, " size=""18""", " size=""7"""))%>
                     <% sendb(disabledattribute) %>>
                    <%
                        If IncentiveProdGroupID > 0 AndAlso Not Page.IsPostBack Then
                            ' alright lets find the currently excluded groups on page load
                            'AMS-684 removed condition: ExcludedProducts=1. Added new table ProductConditionProductGroups in join to get excluded product groups
                            MyCommon.QueryStr = "select PG.ProductGroupID,Name,PG.buyerid from CPE_IncentiveProductGroups as IPG (nolock) " & _
                                                " Inner Join ProductConditionProductGroups PCPG with(nolock) on PCPG.IncentiveProductGroupId = IPG.IncentiveProductGroupId" & _
                                                " inner join ProductGroups as PG (nolock) on PG.ProductGroupID=PCPG.ProductGroupID " & _
                                                " where PG.Deleted=0 and PCPG.Excluded=1 and RewardOptionID=@ROID and PCPG.IncentiveProductGroupId=@IncentiveProductGroupId" & _
                                                " and IPG.deleted=0 and IPG.ProductGroupID is not null"
                            MyCommon.DBParameters.Add("@ROID", SqlDbType.BigInt).Value = roid
                            MyCommon.DBParameters.Add("@IncentiveProductGroupId", SqlDbType.Int).Value = IncentiveProdGroupID
                            If Not Disqualifier Then
                                MyCommon.QueryStr = MyCommon.QueryStr & " and Disqualifier=0"
                            Else
                                MyCommon.QueryStr = MyCommon.QueryStr & " and Disqualifier=1"
                            End If
                            'AMS-684 removed bUseMultipleProductExclusionGroups
                            'If bUseMultipleProductExclusionGroups Then
                            '    MyCommon.QueryStr = MyCommon.QueryStr & " and  IPG.InclusionIncentiveProductGroupSet =" & IncentiveProdGroupID &" "
                            'End If
                            
                            dtExcludedPG = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                            For Each row In dtExcludedPG.Rows
                                If (MyCommon.IsEngineInstalled(9) And SystemCacheData.GetSystemOption_UE_ByOptionId(168) = "1" And Not IsDBNull(row.Item("Buyerid"))) Then
                                    Dim buyerid As Integer = row.Item("Buyerid")
                                    Dim externalBuyerid = MyCommon.GetExternalBuyerId(buyerid)
                                    Send("<option title=""" & MyCommon.NZ(row.Item("Name"), "") & """ value=""" & MyCommon.NZ(row.Item("ProductGroupID"), 0) & """ " & IIf(lstAttributeProductGroups.Contains(row.Item("ProductGroupID")), "style=""color: blue;""", "") & ">" & "Buyer " & externalBuyerid & " - " & MyCommon.NZ(row.Item("Name"), "") & "</option>")
                                Else
                                    Send("<option title=""" & MyCommon.NZ(row.Item("Name"), "") & """ value=""" & MyCommon.NZ(row.Item("ProductGroupID"), 0) & """ " & IIf(lstAttributeProductGroups.Contains(row.Item("ProductGroupID")), "style=""color: blue;""", "") & ">" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
                                End If

                            Next
                        Else
                            If GetCgiValue("exGroups") <> "" Then
                                MyCommon.QueryStr = "select Name,buyerid, ProductGroupID from ProductGroups PG " & _
                                  "INNER JOIN dbo.Split(@ExcludedGroups, ',') excludeitem ON PG.ProductGroupID = excludeitem.items " & _
                                  "ORDER BY ProductGroupID ASC"
                                MyCommon.DBParameters.Add("@ExcludedGroups", SqlDbType.VarChar).Value = GetCgiValue("exGroups")

                                rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                                For Each rowExcluded As DataRow In rst.Rows
                                    If (MyCommon.IsEngineInstalled(9) And SystemCacheData.GetSystemOption_UE_ByOptionId(168) = "1" And Not IsDBNull(rowExcluded.Item("Buyerid"))) Then
                                        Dim buyerid As Integer = rowExcluded.Item("Buyerid")
                                        Dim externalBuyerid = MyCommon.GetExternalBuyerId(buyerid)
                                        Send("<option title=""" & MyCommon.NZ(row.Item("Name"), "") & """ value=""" & MyCommon.NZ(rowExcluded.Item("ProductGroupID"), 0) & """ " & IIf(lstAttributeProductGroups.Contains(rowExcluded.Item("ProductGroupID")), "style=""color: blue;""", "") & ">" & "Buyer " & externalBuyerid & " - " & MyCommon.NZ(rowExcluded.Item("Name"), "") & "</option>")
                                    Else
                                        Send("<option title=""" & MyCommon.NZ(row.Item("Name"), "") & """ value=""" & MyCommon.NZ(rowExcluded.Item("ProductGroupID"), 0) & """ " & IIf(lstAttributeProductGroups.Contains(rowExcluded.Item("ProductGroupID")), "style=""color: blue;""", "") & ">" & MyCommon.NZ(rowExcluded.Item("Name"), "") & "</option>")
                                    End If
                                Next
                                'If rst.Rows.Count > 0 Then

                                'End If
                            End If
                        End If
                    %>
                </select>
            </div>
            <%
                'AMS-684 removed bUseMultipleProductExclusionGroups
                If Disqualifier Then
                    Send("</div>")
                End If
            %>

             <%           
                 'Dim locateHierarchyURL As String = ""
                 'If Not String.IsNullOrWhiteSpace(Request.QueryString("LocateHierarchyURL")) Then
                 '    locateHierarchyURL = HttpUtility.UrlDecode(Request.QueryString("LocateHierarchyURL"))
                 'End If
                 
                 If Not IsPostBack Then
                     If AttributeProductGroupID <> 0 Then
                         ucProductAttributeFilter.HierarchyTreeURL = "/logix/phierarchytree.aspx?ProductGroupID=" & IIf(AttributeProductGroupID > 0, AttributeProductGroupID, 0) & "&PAB=1&OfferID=" & OfferID & "&ConditionID=" & ConditionID & "&Disqualifier=" & Disqualifier
                     Else
                         ucProductAttributeFilter.HierarchyTreeURL = "/logix/phierarchytree.aspx?ProductGroupID=-1&PAB=1&OfferID=" & OfferID & "&ConditionID=" & ConditionID & "&Disqualifier=" & Disqualifier 
                     End If
                 
                 End If
            		%> 
                  
            <hr class="hidden" />
        </div>   
        <%If AttributePGEnabled Then %>
                
           
          <div class="box column3x" id="divAttributeBuilder"  style="float:left;position:relative; width: 90%; height: auto;">
           <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.includedproducts", LanguageID))%>
                </span>
				    <% If (isTemplate) Then%>
                <span class="tempRequire">
                    <input type="checkbox" class="tempcheck" id="require_pg_Attr" name="require_pg_Attr" onclick="handleRequiredToggle(this);"
                        <% if(RequirePG)then sendb(" checked=""checked""") %> />
                    <label for="require_pg_Attr">
                        <% Sendb(Copient.PhraseLib.Lookup("term.required", LanguageID))%>
                    </label>
                </span>
                <% ElseIf (FromTemplate And RequirePG) Then%>
                <span class="tempRequire">
                    <%Sendb("*" & Copient.PhraseLib.Lookup("term.required", LanguageID))%>
                </span>
                <% End If%>
            </h2>
            <div style="float:left;position:relative;width: 90%; height: auto;padding-left:12px">
            <label >Product Group Name:   </label> 
            <input id="txtProductGroupName" runat="server" type="text" style="width:60%" value=""/>
            </div>
            <div id="attributeSelector" style="float:left;position:relative; width: 100%; height: auto;min-height:250px;">
            <br />
                     <uc1:ProductAttributeFilter ID="ucProductAttributeFilter" runat="server" AppName="pgroup-edit.aspx"  />
            </div>
        
           
          </div>     
           <%End If%> 
        
        <%
            'AMS-684 removed conditions other than disqualifier to show exclude box 
            If Disqualifier Then
                Send("<div style=""display:none;"">")
            End If
        %>
        <div class="box column3x" id="attributeExcludeGroupSelector" style="overflow: auto;
            <% Sendb(IIf(AttributePGEnabled AndAlso ProductGroupTypeID <> 2,"display:block""","display:none")) %>">
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.excludedproductgroups", LanguageID))%>
                </span>
            </h2>
            <div id="pgList_attributebased" class="column3x1">
                <input type="radio" id="functionradio1_attr" name="functionradio_attr" <% if(MyCommon.Fetch_SystemOption(175)= "1" ) then sendb(" checked=""checked""") %>
                    <% sendb(disabledattribute) %> /><label for="functionradio1_attr"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
                <input type="radio" id="functionradio2_attr" name="functionradio_attr" <% if(MyCommon.Fetch_SystemOption(175)= "2" ) then sendb(" checked=""checked""") %>
                    <% sendb(disabledattribute) %> /><label for="functionradio2_attr"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
                <input type="text" class="medium" id="functioninput_attr" name="functioninput_attr"
                    maxlength="100" onkeyup="handleKeyUp_attr(200);" value="" <% sendb(disabledattribute) %> /><br />
                <select class="long" id="functionselect_attr" name="functionselect_attr" onchange="DisableButton_attr()"
                    size="10" <% sendb(disabledattribute) %>>
                    <%
                        For Each row In dtAllPGList.Rows
                            If (MyCommon.NZ(row.Item("ProductGroupID"), 0) <> 1 AndAlso MyCommon.NZ(row.Item("ProductGroupID"), 0) <> AttributeProductGroupID) Then
                                If (row.Item("ProductGroupID") IsNot Nothing AndAlso lstAttributeProductGroups.Contains(row.Item("ProductGroupID"))) Then
                                    If (MyCommon.IsEngineInstalled(9) And SystemCacheData.GetSystemOption_UE_ByOptionId(168) = "1" And Not IsDBNull(row.Item("Buyerid"))) Then
                                        Dim pbuyerid As Integer = row.Item("Buyerid")
                                        Dim pexternalBuyerid = MyCommon.GetExternalBuyerId(pbuyerid)
                                        Send("<option value=""" & MyCommon.NZ(row.Item("ProductGroupID"), 0) & """ title=""" & MyCommon.NZ(row.Item("Name"), "") & """ style=""color: blue;"">" & "Buyer " & pexternalBuyerid & " - " & MyCommon.NZ(row.Item("Name"), "") & "</option>")
                                    Else
                                        Send("<option value=""" & MyCommon.NZ(row.Item("ProductGroupID"), 0) & """ title=""" & MyCommon.NZ(row.Item("Name"), "") & """ style=""color: blue;"">" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
                                    End If

                                Else
                                    If (MyCommon.IsEngineInstalled(9) And SystemCacheData.GetSystemOption_UE_ByOptionId(168) = "1" And Not IsDBNull(row.Item("Buyerid"))) Then
                                        Dim pbuyerid As Integer = row.Item("Buyerid")
                                        Dim pexternalBuyerid = MyCommon.GetExternalBuyerId(pbuyerid)
                                        Send("<option  title=""" & MyCommon.NZ(row.Item("Name"), "") & """ value=""" & MyCommon.NZ(row.Item("ProductGroupID"), 0) & """>" & "Buyer " & pexternalBuyerid & " - " & MyCommon.NZ(row.Item("Name"), "") & "</option>")
                                    Else
                                        Send("<option title=""" & MyCommon.NZ(row.Item("Name"), "") & """  value=""" & MyCommon.NZ(row.Item("ProductGroupID"), 0) & """>" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
                                    End If

                                End If
                            End If
                        Next
                    %>
                </select>
            </div>
            <div class="column3x2">
                <center>
                    <br />
                    <br />
                    <br />
                    <input type="button" class="regular select" name="select3" id="select3" value="<% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID))%> &#9658;"
                        onclick="handleSelectClick('select3');" <% sendb(disabledattribute) %> />
                    <br />
                    <br />
                    <input type="button" class="regular deselect" name="deselect3" id="deselect3" value="&#9668; <% Sendb(Copient.PhraseLib.Lookup("term.deselect", LanguageID))%>"
                        onclick="handleSelectClick('deselect3');" <% sendb(disabledattribute) %> /><br />
                </center>
            </div>
            <div class="column3x3">
                <br />
                <div class="boxwrap">
                    <h3>
                        <span>
                            <% Sendb(Copient.PhraseLib.Lookup("term.excludedgroups", LanguageID))%>
                        </span>
                    </h3>
                </div>
                <select class="long" id="excluded_attr" name="excluded_attr" size="10" <% sendb(disabledattribute) %>>
                    <%
                        If IncentiveProdGroupID > 0 AndAlso Not Page.IsPostBack Then
                            For Each row In dtExcludedPG.Rows
                                Send("<option value=""" & MyCommon.NZ(row.Item("ProductGroupID"), 0) & """ " & IIf(lstAttributeProductGroups.Contains(row.Item("ProductGroupID")), "style=""color: blue;""", "") & ">" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
                            Next
                        Else
                            If GetCgiValue("exGroups_attr") <> "" Then
                                MyCommon.QueryStr = "select Name, ProductGroupID from ProductGroups PG " & _
                                  "INNER JOIN dbo.Split(@ExcludedGroups, ',') excludeitem ON PG.ProductGroupID = excludeitem.items " & _
                                  "ORDER BY ProductGroupID ASC"
                                MyCommon.DBParameters.Add("@ExcludedGroups", SqlDbType.VarChar).Value = GetCgiValue("exGroups_attr")

                                rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                                For Each rowExcluded As DataRow In rst.Rows
                                    Send("<option value=""" & MyCommon.NZ(rowExcluded.Item("ProductGroupID"), 0) & """  " & IIf(lstAttributeProductGroups.Contains(rowExcluded.Item("ProductGroupID")), "style=""color: blue;""", "") & ">" & MyCommon.NZ(rowExcluded.Item("Name"), "") & "</option>")
                                Next
                            End If
                        End If
                    %>
                </select>
            </div>
        </div>
        <%
            'AMS-684 removed conditions other than disqualifier to show exclude box 
            If Disqualifier Then
                Send("</div>")
            End If
        %>
        <%If (Disqualifier = False) Then%>
        <div id="column1">
            <div class="box" id="value">
                <h2>
                    <span>
                        <% Sendb(Copient.PhraseLib.Lookup("term.value", LanguageID))%>
                    </span>
                </h2>
                <div id="results">
                    <%
                        If GetCgiValue("selGroups") <> "" Then
                            limits = GetCgiValue("selGroups")
                        ElseIf ProdID = 1 AndAlso (Not exclusionPGList Is Nothing AndAlso exclusionPGList.Count > 0) Then
                            limits = exclusionPGList(0).ProductGroupId
                        Else
                            If ProdID <> 0 Then
                                limits = ProdID
                            Else
                                limits = -1
                            End If
                        End If

                        ' determines which group of amount types to display depending on whether or not UOM is enabled.
                        ' Preload the unit types
                        If (MyCommon.Fetch_UE_SystemOption(203) <> "1") Then
                            UOMCriteria = " UnitTypeID<>10 and (MultiUOMState = " & UOM_ALWAYS & " or MultiUOMState = " & IIf(MyCommon.Fetch_UE_SystemOption(UOM_OPTION_ID) = "1", UOM_ENABLED_ONLY, UOM_DISABLED_ONLY) & ")"
                        Else
                            UOMCriteria = "(MultiUOMState = " & UOM_ALWAYS & " or MultiUOMState = " & IIf(MyCommon.Fetch_UE_SystemOption(UOM_OPTION_ID) = "1", UOM_ENABLED_ONLY, UOM_DISABLED_ONLY) & ")"
                            
                        End If
                        MyCommon.QueryStr = "select UnitTypeID, PhraseID, Description " & _
                                            "from CPE_UnitTypes UT with (NoLock) " & _
                                            "where " & UOMCriteria & " or UnitTypeID=" & Type & ";"
                        rst3 = MyCommon.LRT_Select
                        If IncentiveProdGroupID <> 0 Then
							' KB250202 - CLOUDSOL-3443
                            ' If (limits <> -1) OrElse (limits = -1 AndAlso isTemplate AndAlso RequirePG) OrElse (limits = -1 AndAlso FromTemplate AndAlso RequirePG) Then
                                Send("<table summary=""" & Copient.PhraseLib.Lookup("term.groups", LanguageID) & """>")
                                Send("  <thead>")
                                Send("    <tr>")
                                Send("      <th class=""th-group"" scope=""col"" id=""UnitTypeDesc"">" & Copient.PhraseLib.Lookup("term.quantity", LanguageID) & "</th>")
                                Send("      <th class=""th-unit"" scope=""col"">" & Copient.PhraseLib.Lookup("term.unit", LanguageID) & "</th>")
                                Send("    </tr>")
                                Send("  </thead>")
                                Send("  <tbody>")
                                ' Determine limits
                                ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                                MyCommon.QueryStr = "select IncentiveProductGroupID,ProductGroupID,QtyForIncentive,QtyUnitType,AccumMin,AccumLimit,AccumPeriod,UniqueProduct,MinPurchAmt,ReturnedItemGroup,MinItemPrice,TenderType,NetPriceProduct,SameItem  "
                                'xxx MyCommon.QueryStr &= IIf(SupportSameItemProductCondition, ",SameItem  ", "")
                                MyCommon.QueryStr &= "from CPE_IncentiveProductGroups with (NoLock) " & _
                               "where Deleted=0 and RewardOptionID=" & roid & " and IncentiveProductGroupID=" & IncentiveProdGroupID & ";"
                                rst2 = MyCommon.LRT_Select
                                If rst2.Rows.Count > 0 Then
                                    If MyCommon.NZ(rst2.Rows(0).Item("ProductGroupID"), 0) > 0 Then
                                        MyCommon.QueryStr = "select Name,buyerid from ProductGroups with (NoLock) where ProductGroupID=" & rst2.Rows(0).Item("ProductGroupID") & ";"
                                        rst = MyCommon.LRT_Select
                                        If rst.Rows.Count > 0 Then
                                            If (MyCommon.IsEngineInstalled(9) And MyCommon.Fetch_UE_SystemOption(168) = "1" And Not IsDBNull(rst.Rows(0).Item("Buyerid"))) Then
                                                Dim buyerid As Integer = rst.Rows(0).Item("Buyerid")
                                                Dim externalBuyerid = MyCommon.GetExternalBuyerId(buyerid)
                                                CleanGroupName = MyCommon.SplitNonSpacedString("Buyer " & externalBuyerid & " - " & MyCommon.NZ(rst.Rows(0).Item("Name"), ""), 15)
                                            Else
                                                CleanGroupName = MyCommon.SplitNonSpacedString(MyCommon.NZ(rst.Rows(0).Item("Name"), ""), 15)
                                            End If

                                        End If
                                    End If
                                    Qty = MyCommon.NZ(rst2.Rows(0).Item("QtyForIncentive"), 1)
                                    Type = MyCommon.NZ(rst2.Rows(0).Item("QtyUnitType"), 1)
                                    TenderType = MyCommon.NZ(rst2.Rows(0).Item("TenderType"), 0)
                                    AccumMin = MyCommon.NZ(rst2.Rows(0).Item("AccumMin"), 0)
                                    AccumLimit = MyCommon.NZ(rst2.Rows(0).Item("AccumLimit"), 0)
                                    AccumPeriod = MyCommon.NZ(rst2.Rows(0).Item("AccumPeriod"), 0)
                                    UniqueChecked = MyCommon.NZ(rst2.Rows(0).Item("UniqueProduct"), False)
                                    NetPriceChecked = MyCommon.NZ(rst2.Rows(0).Item("NetPriceProduct"), False)
                                    MinPurchAmt = MyCommon.NZ(rst2.Rows(0).Item("MinPurchAmt"), 0)
                                    ReturnedItemChecked = MyCommon.NZ(rst2.Rows(0).Item("ReturnedItemGroup"), False)
                                    MinItemPrice = MyCommon.NZ(rst2.Rows(0).Item("MinItemPrice"), 0)
                                    'xxx If (SupportSameItemProductCondition) Then
                                        SameItemChecked = MyCommon.NZ(rst2.Rows(0).Item("SameItem"), False)
                                   'xxx End If
                                    'If (rst2.Rows(0).Item("SameItem") = 0) Then
                                    '    NoRestrictionChecked = MyCommon.NZ(rst2.Rows(0).Item("SameItem"), False)
                                    'End If
                                    If Type = 1 Then
                                        IsItem = True
                                    ElseIf Type = 2 Then
                                        IsDollar = True
                                    ElseIf Type = 4 Then
                                        IsQty1 = True
                                    End If
                                Else
                                    Qty = 1
                                    Type = 1
                                    AccumMin = 0
                                    AccumLimit = 0
                                    AccumPeriod = 0
                                End If

                                If Page.IsPostBack Then
                                    If GetCgiValue("itemReturnedGroup") <> String.Empty Then
                                        ReturnedItemChecked = (GetCgiValue("itemReturnedGroup") = "on")
                                    Else
                                        ReturnedItemChecked = False
                                    End If
                                    If GetCgiValue("select") <> String.Empty Then
                                        Type = GetCgiValue("select")
                                    End If
                                    If GetCgiValue("selectTenderType") <> String.Empty Then
                                        TenderType = GetCgiValue("selectTenderType")
                                    End If
                                    If Type = 1 AndAlso Not String.IsNullOrWhiteSpace(GetCgiValue("NetPrice")) Then
                                        NetPriceChecked = True
                                    Else
                                        NetPriceChecked = False
                                    End If
                                    If Type = 1 AndAlso GetCgiValue("ItemRestriction")=ProductFilterEnum.Unique.ToString() Then
                                        UniqueChecked = True
                                    Else
                                        UniqueChecked = False
                                    End If
                                    'xxx If SupportSameItemProductCondition AndAlso Not String.IsNullOrWhiteSpace(GetCgiValue("SameItem")) Then
                                    If Type = 1 AndAlso GetCgiValue("ItemRestriction") = ProductFilterEnum.SameItem.ToString() Then
                                        SameItemChecked = True
                                    Else
                                        SameItemChecked = False
                                    End If
                                    'If Type = 1 AndAlso GetCgiValue("ItemRestriction") = ProductFilterEnum.NoRestriction.ToString() Then
                                    '    NoRestrictionChecked = True
                                    'Else
                                    '    NoRestrictionChecked = False
                                    'End If
                                    If Not String.IsNullOrWhiteSpace(GetCgiValue("MinPurchAmt")) Then
                                        MinPurchAmt = GetCgiValue("MinPurchAmt")
                                    End If
                                    If Not String.IsNullOrWhiteSpace(GetCgiValue("MinItemPrice")) Then
                                        MinItemPrice = GetCgiValue("MinItemPrice")
                                    End If
                                End If

                                Send("    <tr " & Shaded & ">")
                                Send("      <td>")
                                If TierLevels = 1 Or Disqualifier Then
                                    GlobalTier = 0
                                    If rst2.Rows.Count > 0 Then
                                        MyCommon.QueryStr = "Select Quantity from CPE_IncentiveProductGroupTiers with (NoLock) where RewardOptionID=" & roid & " and TierLevel=1 and IncentiveProductGroupID=" & rst2.Rows(0).Item("IncentiveProductGroupID")
                                        TierDT = MyCommon.LRT_Select()
                                        If TierDT.Rows.Count > 0 Then
                                            TierQty = MyCommon.NZ(TierDT.Rows(0).Item("Quantity"), 0)
                                        Else
                                            TierQty = 0
                                        End If
                                    Else
                                        TierQty = 0
                                    End If

                                    If Page.IsPostBack Then
                                        If GetCgiValue("t1_limit") <> "" Then
                                            TierQty = GetCgiValue("t1_limit")
                                        End If
                                    End If

                                    MinPurchAmt = Localization.Round_Currency(MinPurchAmt, roid)
                                    MinItemPrice = Localization.Round_Currency(MinItemPrice, roid)
                                    TierQty = Localization.Round_Quantity(TierQty, roid, Type)
                    								
                                    Send("<span id=""currsymbolt1"" style=""display:none;"">" & CurrencySymbol & "</span>")
                                    Send("<input type=""text"" class=""shorter"" maxlength=""16"" name=""t1_limit"" id=""t1_limit"" value=""" & TierQty.ToString(MyCommon.GetAdminUser.Culture) & """ />")
                                    Send("<span id=""currabbrevt1"" style=""display:none;"">" & CurrencyAbbreviation & "</span>")
                                    Send("<span id=""weightabbrevt1"" style=""display:none;"">" & Localization.GetCached_UOM_Abbreviation(roid, 5, Copient.Localization.UOMUsageEnum.UnitType) & "</span>")
                                    Send("<span id=""volumeabbrevt1"" style=""display:none;"">" & Localization.GetCached_UOM_Abbreviation(roid, 6, Copient.Localization.UOMUsageEnum.UnitType) & "</span>")
                                    Send("<span id=""lengthabbrevt1"" style=""display:none;"">" & Localization.GetCached_UOM_Abbreviation(roid, 7, Copient.Localization.UOMUsageEnum.UnitType) & "</span>")
                                    Send("<span id=""surfareaabbrevt1"" style=""display:none;"">" & Localization.GetCached_UOM_Abbreviation(roid, 8, Copient.Localization.UOMUsageEnum.UnitType) & "</span>")
                                Else
                                    GlobalTier = 1
                                    If MyCommon.Fetch_UE_SystemOption(197) = "1" Then
                                        Send("<input type=""checkbox"" style=""margin-bottom:11px;"" id=""GlobalTierCheckbox"" name=""GlobalTierCheckbox"" value=""1"" onclick=""javascript:SetTiersToValue(this.checked);"" /><label for=""GlobalTierCheckbox"" id=""SelectUnique"">")
                                        Sendb(Copient.PhraseLib.Lookup("term.UseThisValueForAllTiers", LanguageID))
                                        Send("</label>")
                                    End If

                                    For t = 1 To TierLevels
                                        If rst2.Rows.Count > 0 Then
                                            MyCommon.QueryStr = "Select Quantity from CPE_IncentiveProductGroupTiers with (NoLock) where RewardOptionID=" & roid & " and TierLevel=" & t & " and IncentiveProductGroupID=" & rst2.Rows(0).Item("IncentiveProductGroupID")
                                            TierDT = MyCommon.LRT_Select()
                                            If TierDT.Rows.Count > 0 Then
                                                TierQty = MyCommon.NZ(TierDT.Rows(0).Item("Quantity"), 0)
                                            Else
                                                TierQty = 0
                                            End If
                                        Else
                                            TierQty = 0
                                        End If
                                        'Save the values that were incorrect
                                        If Page.IsPostBack Then
                                            If GetCgiValue("t" & t & "_limit") <> "" Then
                                                TierQty = GetCgiValue("t" & t & "_limit")
                                            End If
                                        End If
                                        MinPurchAmt = Math.Round(MinPurchAmt, CurrencyPrecision)
                                        MinItemPrice = Math.Round(MinItemPrice, CurrencyPrecision)
                                        TierQty = Localization.Round_Quantity(TierQty, roid, Type)

                                        If t = 1 Then
                                            TierFirstQty = TierQty
                                        Else
                                            If TierQty <> TierFirstQty Or TierFirstQty = 0 Then
                                                GlobalTier = 0
                                            End If
                                        End If


                                        If TierLevels > 1 Then
                                            Send("<label style=""margin-left:15px;margin-right:15px;"" for=""t" & t & "_limit"">" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</label>")
                                            Send("<span id=""currsymbolt" & t & """ style=""display:none;"">" & CurrencySymbol & "</span>")
                                            Send("<input type=""text"" class=""shorter"" maxlength=""16"" name=""t" & t & "_limit"" id=""t" & t & "_limit"" value=""" & TierQty & """ />")
                                            Send("<span id=""currabbrevt" & t & """ style=""display:none;"">" & CurrencyAbbreviation & "</span>")
                                            Send("<span id=""weightabbrevt" & t & """ style=""display:none;"">" & Localization.Get_UOM_Abbreviation(roid, 5, Copient.Localization.UOMUsageEnum.UnitType) & "</span>")
                                            Send("<span id=""volumeabbrevt" & t & """ style=""display:none;"">" & Localization.Get_UOM_Abbreviation(roid, 6, Copient.Localization.UOMUsageEnum.UnitType) & "</span>")
                                            Send("<span id=""lengthabbrevt" & t & """ style=""display:none;"">" & Localization.Get_UOM_Abbreviation(roid, 7, Copient.Localization.UOMUsageEnum.UnitType) & "</span>")
                                            Send("<span id=""surfareaabbrevt" & t & """ style=""display:none;"">" & Localization.Get_UOM_Abbreviation(roid, 8, Copient.Localization.UOMUsageEnum.UnitType) & "</span>")
                                            Send(IIf(TierLevels > 1, "<br />", ""))
                                        End If
                                    Next
                                    Send("<script type = 'text/javascript'>document.getElementById('GlobalTierCheckbox').checked = " & GlobalTier & ";SetTiersToValue(" & GlobalTier & ")</script>")
                                End If
                                Send("</td>")
                                Send("      <td>")
                                Send("        <select name=""select"" id=""select""  STYLE=""width:120px;"" onchange=""ChangeUnit(this.options[this.selectedIndex].value);"" >")
                                For Each row3 In rst3.Rows
                                    Sendb("          <option")
                                    If (Type = row3.Item("UnitTypeID")) Then
                                        Sendb(" selected=""selected""")
                                    End If
                                    Sendb(" value=""" & row3.Item("UnitTypeID") & """>")
                                    If row3.Item("UnitTypeID") = 2 Then  'for Localization - instead of displaying "Dollars" from the UnitTypes table, display the currency name selected at the offer level
                                        Sendb(CurrencyName)
                                    Else
                                        Sendb(Copient.PhraseLib.Lookup(row3.Item("PhraseID"), LanguageID, MyCommon.NZ(row3.Item("Description"), "!Unknown!")))
                                    End If
                                    Send("    </option>")
                                Next
                                Send("        </select>")
                                Send("      </td>")
                                Send("    </tr>")
                                'Tender Type
                                MyCommon.QueryStr = "select TenderTypeID, Name from CPE_TenderTypes with (NoLock) Where Deleted=0 AND TenderTypeID Not in " & _
                                                    " (select p.TenderType from CPE_IncentiveProductGroups p INNER JOIN CPE_RewardOptions r on p.rewardoptionid=r.rewardoptionid " & _
                                                    " WHERE incentiveid=@OfferID and p.IncentiveProductGroupID != @IncentiveProdGroupID and r.deleted=0 and p.deleted=0 and p.TenderType IS NOT NULL and p.TenderType > 0) ORDER BY CPE_TenderTypes.Name"
                                MyCommon.DBParameters.Add("@OfferID", SqlDbType.Int).Value = OfferID
                                MyCommon.DBParameters.Add("@IncentiveProdGroupID", SqlDbType.Int).Value = IncentiveProdGroupID
                                rstTenderTypes = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                                'AMS-684 replaced ExProdID with exclusionPGList
                                Send("<tr id=""trtendertype""" & Shaded & IIf(ProdID = 1 AndAlso (exclusionPGList Is Nothing Or (Not exclusionPGList Is Nothing AndAlso exclusionPGList.Count <= 0)) AndAlso Type = 2, ">", " style=""display:none;"">") & "<td>")
                                Send(Copient.PhraseLib.Lookup("term.tendertype", LanguageID))
                                Send("</td><td>")
                                Send("        <select style=""width:150px;"" name=""selectTenderType"" id=""selectTenderType"" >")
                                Sendb("          <option value=""0"" ")
                                If (TenderType = 0) Then
                                    Sendb(" selected=""selected""")
                                End If
                                Sendb(">")
                                Send(Copient.PhraseLib.Lookup("term.any", LanguageID) & " " & Copient.PhraseLib.Lookup("term.tendertype", LanguageID))
                                Send("    </option>")
                                For Each row1 In rstTenderTypes.Rows
                                    Sendb("<option value=""" & row1.Item("TenderTypeID") & """")
                                    If (TenderType = row1.Item("TenderTypeID")) Then
                                        Sendb(" selected=""selected""")
                                    End If
                                    Sendb(">")
                                    Sendb(row1.Item("Name"))
                                    Send("    </option>")
                                Next
                                Send("        </select>")
                                Send("</td></tr>")
                                'Send("    <tr " & Shaded & ">")
                                'Send("      <td></td>")
                                'Send("      <td colspan=""2"">")
                                'Send("        <input type=""checkbox"" id=""Unique"" name=""Unique""" & IIf(UniqueChecked And Not EnableAccum, " checked=""checked""", "") & " value=""1"" onclick=""javascript:CheckForSelection('unique')"" " & IIf(Type > 1 OrElse EnableAccum OrElse IsQty1, " disabled=""disabled""", "") & " /><label for=""Unique"" id=""SelectUnique"">" & Copient.PhraseLib.Lookup("term.uniqueproduct", LanguageID) & "</label>")
                                'Send("      </td>")
                                'Send("    </tr>")
                                'If SupportSameItemProductCondition Then
                                '    Send("    <tr " & Shaded & ">")
                                '    Send("      <td></td>")
                                '    Send("      <td colspan=""2"">")
                                '    Send(" <legend>Personalia:</legend>")
                                '    Send("        <input type=""checkbox"" id=""SameItem"" name=""SameItem""" & IIf(SameItemChecked, " checked=""checked""", "") & " value=""1"" onclick=""javascript:CheckForSelection('sameitem')"" " & IIf(Type > 1 OrElse IsQty1, " disabled=""disabled""", "") & " /><label for=""SameItem"" id=""SelectSameItem"">" & Copient.PhraseLib.Lookup("term.sameitem", LanguageID) & "</label>")
                                '    Send("      </td>")
                                '    Send("    </tr>")
                                'End If
                                If MyCommon.Fetch_UE_SystemOption(210) = "1" Then
                                  Send("    <tr " & Shaded & ">")
                                  Send("      <td></td>")
                                  Send("      <td colspan=""2"">")
                                  Send("        <input type=""checkbox"" id=""NetPrice"" name=""NetPrice""" & IIf(TierLevels > 1, " style=""display:none;"" ", "") & IIf(NetPriceChecked, " checked=""checked""", "") & " value=""1""" & " /><label for=""NetPrice"" id=""SelectNetPrice""" &  IIf(TierLevels > 1, " style=""display:none;"" ", "") & ">" & Copient.PhraseLib.Lookup("term.netpriceproduct", LanguageID) & "</label>")
                                  Send("      </td>")
                                  Send("    </tr>")
                                End If

                                Send("    <tr " & Shaded & " id=""MinPurch""" & IIf(IsQty1, " style=""display:none;"" ", "") & ">")
                                Send("      <td><label id=""MinPurchlbl"" for=""MinPurchAmt"" " & IIf(EnableAccum OrElse TierLevels > 1, " style=""display:none;"" ", "") & " >" & Copient.PhraseLib.Lookup("term.minimumgrouppurchase", LanguageID) & "</label></td>")
                                Send("      <td><span id=""MinPurchCurrSymbol"" " & IIf(EnableAccum Or TierLevels > 1, " style=""display:none;""", "") & ">" & CurrencySymbol & "</span><input type=""text"" class=""short"" id=""MinPurchAmt"" name=""MinPurchAmt"" " & IIf(EnableAccum Or TierLevels > 1, " style=""display:none;"" ", "") & " value=""" & MinPurchAmt.ToString(MyCommon.GetAdminUser.Culture) & """ maxlength=""16"" /><span id=""MinPurchCurrAbbrev"" " & IIf(EnableAccum Or TierLevels > 1, " style=""display:none;""", "") & ">" & CurrencyAbbreviation & "</span>")
                                Send("    </tr>")
                                Send("    <tr " & Shaded & " id=""MinItem""" & IIf(IsQty1, " style=""display:none;"" ", "") & ">")
                                Send("      <td><label id=""MinItemlbl"" for=""MinItemPrice"" " & IIf(EnableAccum OrElse TierLevels > 1, " style=""display:none;"" ", "") & " >" & Copient.PhraseLib.Lookup("term.minimumitemprice", LanguageID) & "</label></td>")
                                Send("      <td><span id=""MinItemCurrSymbol"" " & IIf(EnableAccum Or TierLevels > 1, " style=""display:none;""", "") & ">" & CurrencySymbol & "</span><input type=""text"" class=""short"" id=""MinItemPrice"" name=""MinItemPrice"" " & IIf(EnableAccum Or TierLevels > 1, " style=""display:none;"" ", "") & " value=""" & MinItemPrice.ToString(MyCommon.GetAdminUser.Culture) & """ maxlength=""16"" /><span id=""MinItemCurrAbbrev"" " & IIf(EnableAccum Or TierLevels > 1, " style=""display:none;""", "") & ">" & CurrencyAbbreviation & "</span>")
                                Send("    </tr>")
                                rst2 = Nothing
                                If Shaded = " class=""shaded""" Then
                                    Shaded = ""
                                Else
                                    Shaded = " class=""shaded"""
                                End If
                                ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                                Send("  </tbody>")
                                Send("</table>")
                                Send("<br />")
                            ' End If
                        Else
                            Send("<table summary=""" & Copient.PhraseLib.Lookup("term.groups", LanguageID) & """>")
                            Send("  <thead>")
                            Send("    <tr>")
                            Send("      <th class=""th-group"" scope=""col"" id=""UnitTypeDesc"">" & Copient.PhraseLib.Lookup("term.quantity", LanguageID) & "</th>")
                            Send("      <th class=""th-unit"" scope=""col"">" & Copient.PhraseLib.Lookup("term.unit", LanguageID) & "</th>")
                            Send("    </tr>")
                            Send("  </thead>")
                            Send("  <tbody>")
                            Send("    <tr " & Shaded & ">")
                            Send("      <td>")
                            If Disqualifier Then
                                If GetCgiValue("t1_limit") <> "" Then
                                    Send("<span id=""currsymbolt1"" style=""display:none;"">" & CurrencySymbol & "</span>")
                                    Send("<input type=""text"" class=""shorter"" maxlength=""16"" name=""t1_limit"" id=""t1_limit"" value=""" & MyCommon.Extract_Val(GetCgiValue("t1_limit")) & """ />")
                                    Send("<span id=""currabbrevt1"" style=""display:none;"">" & CurrencyAbbreviation & "</span>")
                                    Send("<span id=""weightabbrevt" & t & """ style=""display:none;"">" & Localization.Get_UOM_Abbreviation(roid, 5, Copient.Localization.UOMUsageEnum.UnitType) & "</span>")
                                    Send("<span id=""volumeabbrevt" & t & """ style=""display:none;"">" & Localization.Get_UOM_Abbreviation(roid, 6, Copient.Localization.UOMUsageEnum.UnitType) & "</span>")
                                    Send("<span id=""lengthabbrevt" & t & """ style=""display:none;"">" & Localization.Get_UOM_Abbreviation(roid, 7, Copient.Localization.UOMUsageEnum.UnitType) & "</span>")
                                    Send("<span id=""surfareaabbrevt" & t & """ style=""display:none;"">" & Localization.Get_UOM_Abbreviation(roid, 8, Copient.Localization.UOMUsageEnum.UnitType) & "</span>")
                                Else
                                    Send("<span id=""currsymbolt1"" style=""display:none;"">" & CurrencySymbol & "</span>")
                                    Send("<input type=""text"" class=""shorter"" maxlength=""16"" name=""t1_limit"" id=""t1_limit"" value=""0"" />")
                                    Send("<span id=""currabbrevt1"" style=""display:none;"">" & CurrencyAbbreviation & "</span>")
                                    Send("<span id=""weightabbrevt" & t & """ style=""display:none;"">" & Localization.Get_UOM_Abbreviation(roid, 5, Copient.Localization.UOMUsageEnum.UnitType) & "</span>")
                                    Send("<span id=""volumeabbrevt" & t & """ style=""display:none;"">" & Localization.Get_UOM_Abbreviation(roid, 6, Copient.Localization.UOMUsageEnum.UnitType) & "</span>")
                                    Send("<span id=""lengthabbrevt" & t & """ style=""display:none;"">" & Localization.Get_UOM_Abbreviation(roid, 7, Copient.Localization.UOMUsageEnum.UnitType) & "</span>")
                                    Send("<span id=""surfareaabbrevt" & t & """ style=""display:none;"">" & Localization.Get_UOM_Abbreviation(roid, 8, Copient.Localization.UOMUsageEnum.UnitType) & "</span>")
                                End If
                            Else
                                If TierLevels = 1 Then
                                    If GetCgiValue("t1_limit") <> "" Then
                                        Send("<span id=""currsymbolt1"" style=""display:none;"">" & CurrencySymbol & "</span>")
                                        Send("<input type=""text"" class=""shorter"" maxlength=""16"" name=""t1_limit"" id=""t1_limit"" value=""" & MyCommon.Extract_Val(GetCgiValue("t1_limit")) & """ />")
                                        Send("<span id=""currabbrevt1"" style=""display:none;"">" & CurrencyAbbreviation & "</span>")
                                        Send("<span id=""weightabbrevt1"" style=""display:none;"">" & Localization.Get_UOM_Abbreviation(roid, 5, Copient.Localization.UOMUsageEnum.UnitType) & "</span>")
                                        Send("<span id=""volumeabbrevt1"" style=""display:none;"">" & Localization.Get_UOM_Abbreviation(roid, 6, Copient.Localization.UOMUsageEnum.UnitType) & "</span>")
                                        Send("<span id=""lengthabbrevt1"" style=""display:none;"">" & Localization.Get_UOM_Abbreviation(roid, 7, Copient.Localization.UOMUsageEnum.UnitType) & "</span>")
                                        Send("<span id=""surfareaabbrevt1"" style=""display:none;"">" & Localization.Get_UOM_Abbreviation(roid, 8, Copient.Localization.UOMUsageEnum.UnitType) & "</span>")
                                    Else
                                        Send("<span id=""currsymbolt1"" style=""display:none;"">" & CurrencySymbol & "</span>")
                                        Send("<input type=""text"" class=""shorter"" maxlength=""16"" name=""t1_limit"" id=""t1_limit"" value=""0"" />")
                                        Send("<span id=""currabbrevt1"" style=""display:none;"">" & CurrencyAbbreviation & "</span>")
                                        Send("<span id=""weightabbrevt1"" style=""display:none;"">" & Localization.Get_UOM_Abbreviation(roid, 5, Copient.Localization.UOMUsageEnum.UnitType) & "</span>")
                                        Send("<span id=""volumeabbrevt1"" style=""display:none;"">" & Localization.Get_UOM_Abbreviation(roid, 6, Copient.Localization.UOMUsageEnum.UnitType) & "</span>")
                                        Send("<span id=""lengthabbrevt1"" style=""display:none;"">" & Localization.Get_UOM_Abbreviation(roid, 7, Copient.Localization.UOMUsageEnum.UnitType) & "</span>")
                                        Send("<span id=""surfareaabbrevt1"" style=""display:none;"">" & Localization.Get_UOM_Abbreviation(roid, 8, Copient.Localization.UOMUsageEnum.UnitType) & "</span>")
                                    End If
                                Else
                                    If MyCommon.Fetch_UE_SystemOption(197) = "1" Then
                                        Send("<input type=""checkbox"" style=""margin-left:15px;margin-bottom:11px;"" id=""GlobalTierCheckbox"" name=""GlobalTierCheckbox"" value=""1"" onclick=""javascript:SetTiersToValue(this.checked);"" /><label for=""GlobalTierCheckbox"" id=""SelectUnique"">Use this value for all tiers</label>")
                                    End If

                                    For t = 1 To TierLevels
                                        If TierLevels > 1 Then
                                            If GetCgiValue("t" & t & "_limit") <> "" Then
                                                Send("<label style=""margin-left:15px;margin-right:15px;"" for=""t" & t & "_limit"">" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</label>")
                                                Send("<span id=""currsymbolt" & t & """ style=""display:none;"">" & CurrencySymbol & "</span>")
                                                Send("<input type=""text"" class=""shorter"" maxlength=""16"" name=""t" & t & "_limit"" id=""t" & t & "_limit"" value=""" & MyCommon.Extract_Val(GetCgiValue("t" & t & "_limit")) & """ />")
                                                Send("<span id=""currabbrevt" & t & """ style=""display:none;"">" & CurrencyAbbreviation & "</span>")
                                                Send("<span id=""weightabbrevt" & t & """ style=""display:none;"">" & Localization.Get_UOM_Abbreviation(roid, 5, Copient.Localization.UOMUsageEnum.UnitType) & "</span>")
                                                Send("<span id=""volumeabbrevt" & t & """ style=""display:none;"">" & Localization.Get_UOM_Abbreviation(roid, 6, Copient.Localization.UOMUsageEnum.UnitType) & "</span>")
                                                Send("<span id=""lengthabbrevt" & t & """ style=""display:none;"">" & Localization.Get_UOM_Abbreviation(roid, 7, Copient.Localization.UOMUsageEnum.UnitType) & "</span>")
                                                Send("<span id=""surfareaabbrevt" & t & """ style=""display:none;"">" & Localization.Get_UOM_Abbreviation(roid, 8, Copient.Localization.UOMUsageEnum.UnitType) & "</span>")
                                                Send(IIf(TierLevels > 1, "<br />", ""))
                                            Else
                                                Send("<label style=""margin-left:15px;margin-right:15px;"" for=""t" & t & "_limit"">" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & "</label>")
                                                Send("<span id=""currsymbolt" & t & """ style=""display:none;"">" & CurrencySymbol & "</span>")
                                                Send("<input type=""text"" class=""shorter"" maxlength=""16"" name=""t" & t & "_limit"" id=""t" & t & "_limit"" value=""0"" />")
                                                Send("<span id=""currabbrevt" & t & """ style=""display:none;"">" & CurrencyAbbreviation & "</span>")
                                                Send("<span id=""weightabbrevt" & t & """ style=""display:none;"">" & Localization.Get_UOM_Abbreviation(roid, 5, Copient.Localization.UOMUsageEnum.UnitType) & "</span>")
                                                Send("<span id=""volumeabbrevt" & t & """ style=""display:none;"">" & Localization.Get_UOM_Abbreviation(roid, 6, Copient.Localization.UOMUsageEnum.UnitType) & "</span>")
                                                Send("<span id=""lengthabbrevt" & t & """ style=""display:none;"">" & Localization.Get_UOM_Abbreviation(roid, 7, Copient.Localization.UOMUsageEnum.UnitType) & "</span>")
                                                Send("<span id=""surfareaabbrevt" & t & """ style=""display:none;"">" & Localization.Get_UOM_Abbreviation(roid, 8, Copient.Localization.UOMUsageEnum.UnitType) & "</span>")
                                                Send(IIf(TierLevels > 1, "<br />", ""))
                                            End If
                                        End If
                                    Next
                                End If
                            End If
                            Send("</td>")
                            Send("      <td><select name=""select"" id=""select"" STYLE=""width:120px;"" onchange=""ChangeUnit(this.options[this.selectedIndex].value);"">")
                            For Each row3 In rst3.Rows
                                Sendb("<option")
                                If GetCgiValue("select") <> "" Then
                                    Type = MyCommon.Extract_Val(GetCgiValue("select"))
                                Else
                                    Type = 1
                                End If
                                If (Type = row3.Item("UnitTypeID")) Then
                                    Sendb(" selected=""selected""")
                                End If
                                Sendb(" value=""" & row3.Item("UnitTypeID") & """>")
                                If row3.Item("UnitTypeID") = 2 Then  'for Localization - instead of displaying "Dollars" from the UnitTypes table, display the currency name selected at the offer level
                                    Sendb(CurrencyName)
                                Else
                                    Sendb(Copient.PhraseLib.Lookup(row3.Item("PhraseID"), LanguageID, MyCommon.NZ(row3.Item("Description"), "!Unknown!")))
                                End If

                                Send("</option>")
                            Next
                            Send("</select>")
                            Send("      </td>")
                            Send("    </tr>")
                            'Tender Type
                            MyCommon.QueryStr = "select TenderTypeID, Name from CPE_TenderTypes with (NoLock) Where Deleted=0 AND TenderTypeID Not in " & _
                                                   " (select p.TenderType from CPE_IncentiveProductGroups p INNER JOIN CPE_RewardOptions r on p.rewardoptionid=r.rewardoptionid " & _
                                                   " WHERE incentiveid=@OfferID and p.IncentiveProductGroupID != @IncentiveProdGroupID and r.deleted=0 and p.deleted=0 and p.TenderType IS NOT NULL and p.TenderType > 0) ORDER BY CPE_TenderTypes.Name"
                            MyCommon.DBParameters.Add("@OfferID", SqlDbType.Int).Value = OfferID
                            MyCommon.DBParameters.Add("@IncentiveProdGroupID", SqlDbType.Int).Value = IncentiveProdGroupID
                            rstTenderTypes = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                                
                            If GetCgiValue("selectTenderType") <> "" Then
                                TenderType = MyCommon.Extract_Val(GetCgiValue("selectTenderType"))
                            Else
                                TenderType = 0
                            End If
                            
                            Send("<tr id=""trtendertype""" & Shaded & " style=""display:none;""><td>")
                            Send(Copient.PhraseLib.Lookup("term.tendertype", LanguageID))
                            Send("</td><td>")
                            Send("        <select style=""width:150px;"" name=""selectTenderType"" id=""selectTenderType"" >")
                            Sendb("          <option value=""0"" ")
                            If (TenderType = 0) Then
                                Sendb(" selected=""selected""")
                            End If
                            Sendb(">")
                            Send(Copient.PhraseLib.Lookup("term.any", LanguageID) & " " & Copient.PhraseLib.Lookup("term.tendertype", LanguageID))
                            Send("    </option>")
                            For Each row1 In rstTenderTypes.Rows
                                Sendb("<option value=""" & row1.Item("TenderTypeID") & """")
                                If (TenderType = row1.Item("TenderTypeID")) Then
                                    Sendb(" selected=""selected""")
                                End If
                                Sendb(">")
                                Sendb(row1.Item("Name"))
                                Send("    </option>")
                            Next
                            Send("        </select>")
                            ''need to be decided 
                            'Send("</td></tr>")
                            'Send("    <tr " & Shaded & ">")
                            'Send("      <td></td>")
                            'Send("      <td colspan=""2"">")
                            'Send("        <input type=""radio"" id=""unique"" name=""unique""" & IIf(GetCgiValue("unique") <> "" And Not EnableAccum, " checked=""checked""", "") & " value=""1"" onclick=""javascript:checkforselection('unique')"" " & IIf(Type > 1 OrElse EnableAccum, " disabled ", "") & " /><label for=""unique"" id=""selectunique"">" & Copient.PhraseLib.Lookup("term.uniqueitems", LanguageID) & "</label>")
                            'Send("      </td>")
                            'Send("    </tr>")
                            'If SupportSameItemProductCondition Then
                            '    Send("    <tr " & Shaded & ">")
                            '    Send("      <td></td>")
                            '    Send("      <td colspan=""2"">")
                            '    Send("        <input type=""radio"" id=""sameitem"" name=""sameitem""" & IIf(GetCgiValue("sameitem") <> "", " checked=""checked""", "") & " value=""1"" onclick=""javascript:checkforselection('sameitem')""  /><label for=""sameitem"" id=""selectsameitem"">" & Copient.PhraseLib.Lookup("term.sameitems", LanguageID) & "</label>")
                            '    Send("      </td>")
                            '    Send("    </tr>")
                            'End If
                            Send("</td></tr>")
                            If MyCommon.Fetch_UE_SystemOption(210) = "1" Then
                              Send("    <tr " & Shaded & ">")
                              Send("      <td></td>")
                              Send("      <td colspan=""2"">")
                              Send("        <input type=""checkbox"" id=""NetPrice"" name=""NetPrice""" & IIf(TierLevels > 1, " style=""display:none;"" ", "") & IIf(NetPriceChecked, " checked=""checked""", "") & " value=""1""" & " /><label for=""NetPrice"" id=""SelectNetPrice""" &  IIf(TierLevels > 1, " style=""display:none;"" ", "") & ">" & Copient.PhraseLib.Lookup("term.netpriceproduct", LanguageID) & "</label>")
                              Send("      </td>")
                              Send("    </tr>")
                            End If

                            Send("    <tr " & Shaded & " id=""MinPurch"">")
                            Send("      <td><label id=""MinPurchlbl"" for=""MinPurchAmt"" " & IIf(EnableAccum Or TierLevels > 1, " style=""display:none;"" ", "") & " >" & Copient.PhraseLib.Lookup("term.minimumgrouppurchase", LanguageID) & "</label></td>")
                            Send("      <td><span id=""MinPurchCurrSymbol"" " & IIf(EnableAccum Or TierLevels > 1, " style=""display:none;""", "") & ">" & CurrencySymbol & "</span><input type=""text"" class=""short"" id=""MinPurchAmt"" name=""MinPurchAmt"" " & IIf(EnableAccum Or TierLevels > 1, " style=""display:none;"" ", "") & " value=""" & MinPurchAmt.ToString(MyCommon.GetAdminUser.Culture) & """ maxlength=""16"" /><span id=""MinPurchCurrAbbrev"" " & IIf(EnableAccum Or TierLevels > 1, " style=""display:none;""", "") & ">" & CurrencyAbbreviation & "</span>")
                            Send("    </tr>")
                            Send("    <tr " & Shaded & " id=""MinItem"">")
                            Send("      <td><label id=""MinItemlbl"" for=""MinItemPrice"" " & IIf(EnableAccum Or TierLevels > 1, " style=""display:none;"" ", "") & " >" & Copient.PhraseLib.Lookup("term.minimumitemprice", LanguageID) & "</label></td>")
                            Send("      <td><span id=""MinItemCurrSymbol"" " & IIf(EnableAccum Or TierLevels > 1, " style=""display:none;""", "") & ">" & CurrencySymbol & "</span><input type=""text"" class=""short"" id=""MinItemPrice"" name=""MinItemPrice"" " & IIf(EnableAccum Or TierLevels > 1, " style=""display:none;"" ", "") & " value=""" & MinItemPrice.ToString(MyCommon.GetAdminUser.Culture) & """ maxlength=""16"" /><span id=""MinItemCurrAbbrev"" " & IIf(EnableAccum Or TierLevels > 1, " style=""display:none;""", "") & ">" & CurrencyAbbreviation & "</span>")
                            Send("    </tr>")
                            rst2 = Nothing
                            If Shaded = " class=""shaded""" Then
                                Shaded = ""
                            Else
                                Shaded = " class=""shaded"""
                            End If
                            Send("  </tbody>")
                            Send("</table>")
                            Send("<br />")
                        End If

                        'Rounding
                        If MyCommon.Fetch_UE_SystemOption(89) = "1" And EngineID = 2 Then
                            Send("<span id=""roundingspan""" & IIf(Type = 1 OrElse Type = 4, " style=""display:none;""", "") & ">")
                            Send("<input type=""checkbox"" id=""rounding"" name=""rounding"" value=""on""" & IIf(Rounding, " checked=""checked""", "") & " /><label for=""rounding"">" & Copient.PhraseLib.Lookup("CPEoffer-con-product.rounding", LanguageID) & "</label><br />")
                            Send("<br class=""half"" />")
                            Send("</span>")
                        End If

                        'Accumulation checkbox
                        ' Disabled accumulation for UE per RT 4604

                        'If (ShowAccum And Not Disqualifier And Not HasDisqualifier And Not HasAnyCustomer) AndAlso (TierLevels = 1) Then
                        '  Send("<input type=""checkbox"" id=""EnableAccum"" onclick=""toggleAccum();"" name=""EnableAccum"" value=""1"" " & IIf(EnableAccum, "checked=""checked""", "") & IIf(IsQty1, " style=""display:none;""", "") & " /><label id=""AccumChecklbl"" for=""EnableAccum""" & IIf(IsQty1, " style=""display:none;""", "") & ">" & Copient.PhraseLib.Lookup("term.enable", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.accumulation", LanguageID), VbStrConv.Lowercase) & "</label>")
                        '  Send("<br />")
                        '  ' Do the accumulation stuff here
                        '  Send("<div id=""accumulation"" style=""width:310px;" & IIf(EnableAccum AndAlso Not IsQty1, "", "display:none;") & """>")
                        '  Send("<table summary=""" & Copient.PhraseLib.Lookup("term.accumulation", LanguageID) & """>")
                        '  Send("  <thead>")
                        '  Send("    <tr>")
                        '  Send("      <th class=""th-minimum"" scope=""col"">" & Copient.PhraseLib.Lookup("term.minimum", LanguageID) & "</th>")
                        '  Send("      <th class=""th-limit"" scope=""col"">" & Copient.PhraseLib.Lookup("term.limit", LanguageID) & "</th>")
                        '  Send("      <th class=""th-period"" scope=""col"">" & Copient.PhraseLib.Lookup("term.period", LanguageID) & "</th>")
                        '  Send("    </tr>")
                        '  Send("  </thead>")
                        '  Send("  <tbody>")
                        '  Send("    <tr>")
                        '  If IsItem Then
                        '    AccumMin = Math.Truncate(AccumMin)
                        '  ElseIf IsDollar Then
                        '    AccumMin = Math.Round(AccumMin, 2)
                        '  End If
                        '  If Not EnableAccum Then AccumMin = 0
                        '  Send("      <td><input type=""text"" class=""shorter"" maxlength=""9"" name=""accummin"" id=""accummin"" value=""" & AccumMin & """ /></td>")
                        '  If IsItem Then
                        '    AccumLimit = Math.Truncate(AccumLimit)
                        '  ElseIf IsDollar Then
                        '    AccumLimit = Math.Round(AccumLimit, 2)
                        '  End If
                        '  If Not EnableAccum Then
                        '    AccumLimit = 0
                        '    AccumPeriod = 0
                        '  End If
                        '  Send("      <td><input type=""text"" class=""shorter"" maxlength=""9"" name=""accumlimit"" id=""accumlimit"" value=""" & AccumLimit & """ /> " & StrConv(Copient.PhraseLib.Lookup("term.per", LanguageID), VbStrConv.Lowercase) & "</td>")
                        '  Send("      <td><input type=""text"" class=""shorter"" maxlength=""9"" name=""accumperiod"" id=""accumperiod"" value=""" & AccumPeriod & """ /> " & StrConv(Copient.PhraseLib.Lookup("term.days", LanguageID), VbStrConv.Lowercase) & "</td>")
                        '  Send("    </tr>")
                        '  Send("  </tbody>")
                        '  Send("</table>")
                        '  Send("</div>")
                        '  If EnableAccum Then ProductComboID = 0
                        'End If
                    %>
                </div>
                <br class="half" />
                <hr class="hidden" />
            </div>
        </div>
        <div id="gutter">
        </div>
        <div id="column2">
            <div class="box" id="options" <% If restrictRewardforRPOS Then Sendb("style=""display:none;""") %>>
                <h2>
                    <span>
                        <% Sendb(Copient.PhraseLib.Lookup("term.advancedoptions", LanguageID))%>
                    </span>
                </h2>
                <%

                    'NCR 55 ReturnedItemGroup functionality allowing cashier messages to be displayed when items are returned. Can be turned off by setting UE_SystemOption 182 to 0.                              
        			If MyCommon.Fetch_UE_SystemOption(182) = "1" Then
						MyCommon.QueryStr = "select DeliverableTypeID from CPE_Deliverables with (NoLock) where Deleted=0 and RewardOptionID=" & roid & ";" 
        				rst2 = MyCommon.LRT_Select
	        			If rst2.Rows.Count < 1 Then
	              			Sendb("<input type=""checkbox"" id=""itemReturnedGroup""  name=""itemReturnedGroup"" value=""on""")
        	      			Sendb(If(ReturnedItemChecked, " checked=""checked""", "") & " />")
	              			Send(Copient.PhraseLib.Lookup("term.ReturnedItemGroup", LanguageID))
        	      			Send("<br /><br />")
	        			Else If rst2.Rows.Count = "1" And rst2.Rows(0).Item("DeliverableTypeID") = 9 Then
        	      			Sendb("<input type=""checkbox"" id=""itemReturnedGroup""  name=""itemReturnedGroup"" value=""on""")
	              			Sendb(If(ReturnedItemChecked, " checked=""checked""", "") & " />")
	              			Send(Copient.PhraseLib.Lookup("term.ReturnedItemGroup", LanguageID))
	              			Send("<br /><br />")
	        			End If
        			End If

                    MyCommon.QueryStr = "select convert(nvarchar(10), FullPrice) + convert(nvarchar(10),ClearanceState) + CONVERT(nvarchar(10), ClearanceLevel) as PriceFilter " & _
                                        "from CPE_IncentiveProductGroups with (NoLock) " & _
                                        "where Deleted=0 and RewardOptionID=" & roid & " and IncentiveProductGroupID=" & IncentiveProdGroupID & ";"
                    rst2 = MyCommon.LRT_Select
                    If rst2.Rows.Count > 0 Then
                        PriceFilter = MyCommon.NZ(rst2.Rows(0).Item("PriceFilter"), "000")
                    End If
                    If Page.IsPostBack AndAlso Not String.IsNullOrWhiteSpace(GetCgiValue("priceFilter")) Then
                        PriceFilter = GetCgiValue("priceFilter").Trim()
                    End If

                    Send("<label for=""priceFilter"">" & Copient.PhraseLib.Lookup("pricefilter.selectedproducts", LanguageID) & ":</label><br />")
                    Send("<select id=""priceFilter"" name=""priceFilter"" class=""long"">")
                    Send("  <option value=""000""" & If(PriceFilter = "000", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("pricefilter.no-filter", LanguageID) & "</option>")
                    Send("  <option value=""100""" & If(PriceFilter = "100", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("pricefilter.fullprice", LanguageID) & "</option>")
                    Send("  <option value=""010""" & If(PriceFilter = "010", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("pricefilter.notonclearance", LanguageID) & "</option>")
                    Send("  <option value=""020""" & If(PriceFilter = "020", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("pricefilter.anyclearancelevel", LanguageID) & "</option>")

                    MyCommon.QueryStr = "select ClearanceLevelValue, Description, PhraseID from UE_ClearanceLevels with (NoLock) order by ClearanceLevelValue;"
                    rst2 = MyCommon.LRT_Select
                    For Each row2 As DataRow In rst2.Rows
                        Send("  <option value=""03" & MyCommon.NZ(row2.Item("ClearanceLevelValue"), 0).ToString & """" & If(PriceFilter = "03" & MyCommon.NZ(row2.Item("ClearanceLevelValue"), 0), " selected=""selected""", "") & ">")
                        Send(Copient.PhraseLib.Lookup(MyCommon.NZ(row2.Item("PhraseID"), 0), LanguageID, MyCommon.NZ(row2.Item("Description"), "Unknown")))
                        Send("  </option>")
                    Next
                    Send("</select>")
                    Send("      <br>")
                    Send("      <br>")
                    Send(" <fieldset>")
                    Send("<legend> " & Copient.PhraseLib.Lookup("term.combofilter", LanguageID) & "</legend>")
                    Send("    <tr " & Shaded & ">")
                    Send("      <td></td>")
                    Send("      <td colspan=""2"">")
                    Send("        <input type=""radio"" id=""NoRestriction"" name=""ItemRestriction""" & IIf(SameItemChecked = False And UniqueChecked = False, " checked=""checked""", "") & " value=""NoRestriction"" " & IIf(Type > 1 OrElse EnableAccum OrElse IsQty1, " disabled=""disabled""", "") & " /><label for=""NoRestriction"" id=""SelectNoRestriction"">" & Copient.PhraseLib.Lookup("sv.redemption.norestriction", LanguageID) & "</label>")
                    Send("      </td>")
                    Send("    </tr>")
                    Send("      <br>")
                'xxx    If SupportSameItemProductCondition Then
                        Send("    <tr " & Shaded & ">")
                        Send("      <td></td>")
                        Send("      <td colspan=""2"">")
                    Send("        <input type=""radio"" id=""SameItem"" name=""ItemRestriction""" & IIf(SameItemChecked, " checked=""checked""", "") & " value=""SameItem"" " & IIf(Type > 1 OrElse IsQty1, " disabled=""disabled""", "") & " /><label for=""SameItem"" id=""SelectSameItem"">" & Copient.PhraseLib.Lookup("term.sameitems", LanguageID) & "</label>")
                        Send("      </td>")
                        Send("    </tr>")
                    'xxx End If
                    Send("      <br>")
                    Send("    <tr " & Shaded & ">")
                    Send("      <td></td>")
                    Send("      <td colspan=""2"">")
                    Send("        <input type=""radio"" id=""Unique"" name=""ItemRestriction""" & IIf(UniqueChecked And Not EnableAccum, " checked=""checked""", "") & " value=""Unique"" " & IIf(Type > 1 OrElse EnableAccum OrElse IsQty1, " disabled=""disabled""", "") & " /><label for=""Unique"" id=""SelectUnique"">" & Copient.PhraseLib.Lookup("term.uniqueitems", LanguageID) & "</label>")
                    Send("      </td>")
                    Send("    </tr>")
                    Send("  </fieldset>")
                %>
                <hr class="hidden" />
            </div>
            <%Else%>
            <input type="hidden" name="t1_limit" id="t1_limit" value="0" />
            <%End If%>
        </div>
    </div>
</div>
</form>
<script runat="server">
    '------------------------------------------------------------------------------------------------------------------------------------------
    Dim CopientFileName As String = String.Empty
    Dim CopientFileVersion As String = "7.3.1.138972"
    Dim CopientProject As String = "Copient Logix"
    Dim CopientNotes As String = ""
    Dim restrictRewardforRPOS As Boolean = False
    Dim MyCommon As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    Dim Localization As Copient.Localization
    Dim Hierarchy As New Copient.HierarchyResync(MyCommon, "UI", "HierarchyFeeds.txt")
    Dim rst As DataTable
    Dim dt As DataTable
    Dim row As DataRow
    Dim OfferID As Long
    Dim Name As String = ""
    Dim ConditionID As String
    Dim isTemplate As Boolean
    Dim FromTemplate As Boolean
    Dim Disallow_Edit As Boolean = True
    Dim Disqualifier As Boolean = False
    Dim DisabledAttribute As String = ""
    Dim roid As Integer
    Dim ProductGroups As Integer
    Dim historyString As String
    Dim CloseAfterSave As Boolean = False
    Dim Qty As Decimal
    Dim Type As Integer
    Dim TenderType As Integer = 0
    Dim AccumMin As Decimal
    Dim AccumLimit As Decimal
    Dim AccumPeriod As Integer
    Dim AccumEligible As Boolean
    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False
    Dim RequirePG As Boolean = False
    Dim HasRequiredPG As Boolean = False
    Dim EngineID As Integer = 2
    Dim BannersEnabled As Boolean = True
    Dim HasTenderCondition As Boolean = False
    Dim HasDisqualifier As Boolean = False
    Dim UniqueProduct As Integer = 0
    Dim NetPriceProduct as Integer = 0
    Dim IsUniqueProd As Boolean = False
    Dim TierLevels As Integer = 1
    Dim t As Integer = 1
    Dim TierQty As Decimal
    Dim ValidTier As Boolean = False
    Dim TierFirstQty As Decimal 'vtopol_DT13
    Dim TierTempQty As Decimal 'vtopol_DT13
    Dim GlobalTier As Integer = 0 'vtopol_DT13
    Dim ProdID As Integer = 0
    'Dim ExProdID As Integer = 0
    Dim IncentiveProdGroupID As Integer = 0
    Dim TempIncentiveProdGroupID As Integer=0
    Dim rst2 As DataTable
    Dim rst3 As DataTable
    Dim rstTenderTypes As DataTable
    Dim UniqueChecked As Boolean = False
    Dim NetPriceChecked as Boolean = False
    Dim IsItem As Boolean = False
    Dim IsDollar As Boolean = False
    Dim IsQty1 As Boolean = False
    Dim CleanGroupName As String = ""
    Dim Shaded As String = " class=""shaded"""
    Dim TierDT As DataTable
    Dim row3 As DataRow
    Dim limits As String = ""
    Dim ShowAccum, ValidAccum As Boolean
    Dim EnableAccum As Boolean = False
    Dim t1, t2 As Decimal
    Dim AnyProductUsed As Boolean = False
    Dim Rounding As Boolean = False
    Dim ValidRounding As Boolean = True
    Dim MinPurchAmt As Decimal = 0
    Dim ReturnedItemChecked As Boolean = False
    Dim MinItemPrice As Decimal = 0
    Dim HasAnyCustomer As Boolean
    Dim PriceFilter As String = "000"
    Dim UOMCriteria As String = ""
    Dim ValidPMR As Boolean = False
    Dim ValidGCR As Boolean = True
    Dim gcrErrorPhrase As String
    Const UOM_ALWAYS As Integer = -1
    Const UOM_DISABLED_ONLY As Integer = 0
    Const UOM_ENABLED_ONLY As Integer = 1
    Const UOM_OPTION_ID As Integer = 135
    Dim AttributePGEnabled As Boolean = False
    Dim ProductGroupTypeID As Byte = 1
    Dim m_ProductGroupService As IProductGroupService
    Dim m_ActivityLogService As IActivityLogService
    Dim m_OfferService As IOffer
    Dim CurrencyID As Integer = 0
    Dim CurrencyPrecision As Integer = 2
    Dim CurrencySymbol As String = ""
    Dim CurrencyName As String = "Dollars"
    Dim CurrencyAbbreviation As String = ""
    Dim AttributeProductGroupID As Integer = 0
    Dim dtAllPGList As New DataTable
    Dim dtExcludedPG As New DataTable
    Dim lstAttributeProductGroups As New List(Of Long)
    Dim AttributeSwitchType As String = String.Empty
    Dim BuyerID As Integer = 0
    Dim objOffer As Offer
    Dim NodeID As String=""
    Dim ProductGroupName As String=""
    Dim Selected As String=""
    Dim Linking As String=""
    Dim LinkedItems As String=""
    Dim ProductGroupID As long
    Dim ValidPMRAwayValues As Boolean = True
    'Dim locateHierarchyURL As String = ""
    Dim PABStage As Int16 = 1
    Dim ValidNodeId As Boolean = True
    Dim ValidMultipleExclusionProdCondition As Boolean =True
    'xxx  Dim SupportSameItemProductCondition As Boolean = IIf(MyCommon.Fetch_UE_SystemOption(198) = "1", True, False)
    Dim SameItem As Integer = 0
    Dim SameItemChecked As Boolean = False
    'Dim NoRestriction As Integer = 2
    'Dim NoRestrictionChecked As Boolean = False
    'Dim bUseMultipleProductExclusionGroups As Boolean = True
    Dim m_ProductConditionPGService As IProductConditionService
    Dim exclusionPGList As List(Of ProductConditionProductGroup)
    Dim resultExcludedPGList As AMSResult(Of List(Of ProductConditionProductGroup))
    'AMS-2223 Product group list paging settings
    Dim shouldFetchPGAsync As Boolean = MyCommon.Fetch_SystemOption(289)
    'Dim Shared listSize As integer = 50
    Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249) = "1", True, False)
    Dim SystemCacheData As ICacheData

    Enum ProductFilterEnum
        NoRestriction
        SameItem
        Unique
    End Enum


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
        CopientFileName = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))

        If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
            Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
        End If

        Response.Expires = 0
        MyCommon.AppName = "UEoffer-con-product.aspx"
        CurrentRequest.Resolver.AppName = MyCommon.AppName
        m_ProductGroupService = CurrentRequest.Resolver.Resolve(Of IProductGroupService)()
        m_ActivityLogService = CurrentRequest.Resolver.Resolve(Of IActivityLogService)()
        m_OfferService = CurrentRequest.Resolver.Resolve(Of IOffer)()
        m_ProductConditionPGService = CurrentRequest.Resolver.Resolve(Of IProductConditionService)()
        SystemCacheData = CurrentRequest.Resolver.Resolve(Of ICacheData)()

        MyCommon.Open_LogixRT()
        AdminUserID = Verify_AdminUser(MyCommon, Logix)
        Localization = New Copient.Localization(MyCommon)
        restrictRewardforRPOS = (SystemCacheData.GetSystemOption_UE_ByOptionId(234) = "1")
        BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
        OfferID = MyCommon.Extract_Val(GetCgiValue("OfferID"))

        'This would redirect to summary page if offer is awaiting deployment or if offer is awaiting recommendations.
        CheckIfValidOffer(MyCommon, OfferID)

        ConditionID = GetCgiValue("ConditionID")
        Disqualifier = IIf(GetCgiValue("Disqualifier") = "1", True, False)
        HasAnyCustomer = UEOffer_Has_AnyCustomer(MyCommon, OfferID)
        NodeID = GetCgiValue("NodeListID")
        Selected = GetCgiValue("selected")
        Linking = GetCgiValue("Linking")
        If GetCgiValue("IncentiveProductGroupID") <> "" Then IncentiveProdGroupID = MyCommon.Extract_Val(GetCgiValue("IncentiveProductGroupID"))
        If GetCgiValue("AttributeProductGroupID") <> "" Then AttributeProductGroupID = MyCommon.Extract_Val(GetCgiValue("AttributeProductGroupID"))
        If GetCgiValue("hdnBuyerID") <> "" Then BuyerID = MyCommon.Extract_Val(GetCgiValue("hdnBuyerID"))
        If Not String.IsNullOrWhiteSpace(GetCgiValue("AttributeSwitchType")) Then AttributeSwitchType = GetCgiValue("AttributeSwitchType")
        If (GetCgiValue("EngineID") <> "") Then
            EngineID = MyCommon.Extract_Val(GetCgiValue("EngineID"))
        Else
            MyCommon.QueryStr = "select EngineID from OfferIDs with (NoLock) where OfferID=" & OfferID & ";"
            rst = MyCommon.LRT_Select
            If rst.Rows.Count > 0 Then
                EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), 0)
            End If
        End If



        AttributePGEnabled = (MyCommon.Fetch_UE_SystemOption(157) = "1") AndAlso (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE))
        If (AttributePGEnabled) Then
            Dim AttributeBasedPGIds As AMSResult(Of List(Of Int64)) = m_ProductGroupService.GetAllAttributeBasedProductGroupIDs
            If AttributeBasedPGIds.ResultType <> AMSResultType.Success Then
                infoMessage = AttributeBasedPGIds.MessageString
            Else
                lstAttributeProductGroups = AttributeBasedPGIds.Result
            End If

            If Not Page.IsPostBack Then
                Dim ProductGroupTypes As AMSResult(Of List(Of ProductGroupTypes)) = m_ProductGroupService.GetProductGroupTypes()
                If (ProductGroupTypes.ResultType <> AMSResultType.Success) Then
                    infoMessage = ProductGroupTypes.MessageString
                Else
                    radiobtnlistpgselection.DataSource = ProductGroupTypes.Result
                    radiobtnlistpgselection.DataTextField = "Name"
                    radiobtnlistpgselection.DataValueField = "ProductGroupTypeID"
                    radiobtnlistpgselection.DataBind()
                    radiobtnlistpgselection.RepeatDirection = RepeatDirection.Horizontal
                    radiobtnlistpgselection.Attributes.Add("onclick", "javascript:ProductGroupTypeSelection();ShoworHideDivsWithoutPostback();ShowOrHideTenderType();")

                    If (radiobtnlistpgselection.Items.Count > 0) Then
                        If radiobtnlistpgselection.Items(0) IsNot Nothing Then
                            radiobtnlistpgselection.Items(0).Text = Copient.PhraseLib.Lookup("term.selectexistingprodgroups", LanguageID)
                            radiobtnlistpgselection.Items(0).Attributes.Add("style", "margin-right:50px;")
                        End If
                        If radiobtnlistpgselection.Items(1) IsNot Nothing Then
                            radiobtnlistpgselection.Items(1).Text = Copient.PhraseLib.Lookup("term.useattributeprodgroups", LanguageID)
                        End If

                        ProductGroupTypeID = MyCommon.Fetch_UE_SystemOption(156)
                        SelectAttributeType(ProductGroupTypeID)
                    End If
                End If
            Else
                If radiobtnlistpgselection.Items.Count > 0 Then radiobtnlistpgselection.Items(0).Attributes.Add("style", "margin-right:50px;")
                ProductGroupTypeID = radiobtnlistpgselection.SelectedItem.Value.ConvertToByte()
                SelectAttributeType(ProductGroupTypeID)
            End If
        Else
            radiobtnlistpgselection.Visible = False
        End If
        ProductGroupName = ""

        If AttributePGEnabled Then
            If MyCommon.Extract_Val(GetCgiValue("selGroups")) <> "0" And m_ProductGroupService.GetProductGroupType(GetCgiValue("selGroups").ConvertToLong()).Result = 1 Then
                If GetCgiValue("txtProductGroupName") <> m_ProductGroupService.GetProductGroupName(m_ProductGroupService.GetProductGroupID(IncentiveProdGroupID).Result).Result Then
                    ProductGroupName = GetCgiValue("txtProductGroupName")
                End If
            Else
                If MyCommon.Extract_Val(GetCgiValue("selGroups")) <> "0" And m_ProductGroupService.GetProductGroupType(GetCgiValue("selGroups").ConvertToLong()).Result = 2 Then
                    If GetCgiValue("txtProductGroupName") <> "" AndAlso GetCgiValue("txtProductGroupName") = m_ProductGroupService.GetProductGroupName(MyCommon.Extract_Val(GetCgiValue("selGroups"))).Result AndAlso AttributeSwitchType = "SelectedAttributeGroup" Then
                        ProductGroupName = GetCgiValue("txtProductGroupName")
                    ElseIf GetCgiValue("txtProductGroupName") <> "" AndAlso GetCgiValue("txtProductGroupName") <> m_ProductGroupService.GetProductGroupName(MyCommon.Extract_Val(GetCgiValue("selGroups"))).Result AndAlso AttributeSwitchType <> "SelectedAttributeGroup" Then
                        ProductGroupName = GetCgiValue("txtProductGroupName")
                    Else
                        ProductGroupName = m_ProductGroupService.GetProductGroupName(GetCgiValue("selGroups").ConvertToLong()).Result
                    End If
                Else
                    If GetCgiValue("txtProductGroupName") <> "" AndAlso AttributeSwitchType <> "DeSelectedAttributeGroup"  Then
                        ProductGroupName = GetCgiValue("txtProductGroupName")
                    Else
                        If AttributeSwitchType <> "DeSelectedAttributeGroup" AndAlso m_ProductGroupService.GetProductGroupType(IIf(MyCommon.Extract_Val(GetCgiValue("selGroups")) <> "0", MyCommon.Extract_Val(GetCgiValue("selGroups")), m_ProductGroupService.GetProductGroupID(IncentiveProdGroupID).Result)).Result = 2 Then
                            ProductGroupName = m_ProductGroupService.GetProductGroupName(m_ProductGroupService.GetProductGroupID(IIf(AttributeProductGroupID = 0, IncentiveProdGroupID, AttributeProductGroupID)).Result).Result
                        End If
                    End If
                End If
            End If

        End If
        If ProductGroupName = "" Then
            ProductGroupName = String.Concat(Copient.PhraseLib.Lookup("term.offer", LanguageID), " ", OfferID.ToString(), " ", Copient.PhraseLib.Lookup("term.conditionalproducts", LanguageID).ToLower())
        End If
        txtProductGroupName.Value = ProductGroupName
        MyCommon.QueryStr = "select RO.RewardOptionID, RO.TierLevels, RO.CurrencyID, isnull(C.Symbol, '') as Symbol, isnull(C.Precision, 2) as Precision, isnull(C.AbbreviationPhraseTerm, '') as AbbreviationPhraseTerm, " & _
                            "isnull(C.NamePhraseTerm, 'term.dollars') as NamePhraseTerm " & _
                            "from CPE_RewardOptions as RO with (NoLock) " & _
                            "left Join Currencies as C on C.CurrencyID=RO.CurrencyID " & _
                            "where RO.IncentiveID=" & OfferID & " and RO.TouchResponse=0 and RO.Deleted=0;"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
            roid = rst.Rows(0).Item("RewardOptionID")
            TierLevels = rst.Rows(0).Item("TierLevels")
            CurrencyID = rst.Rows(0).Item("CurrencyID")
            CurrencySymbol = rst.Rows(0).Item("Symbol")
            CurrencyPrecision = rst.Rows(0).Item("Precision")
            CurrencyName = Copient.PhraseLib.Lookup(rst.Rows(0).Item("NamePhraseTerm"), LanguageID)
            CurrencyAbbreviation = IIf(rst.Rows(0).Item("AbbreviationPhraseTerm") = "", Copient.PhraseLib.Lookup("term.dollars", LanguageID), Copient.PhraseLib.Lookup(rst.Rows(0).Item("AbbreviationPhraseTerm"), LanguageID))
        Else
            infoMessage = Copient.PhraseLib.Lookup("term.errornorewardoption", LanguageID)
        End If

        MyCommon.QueryStr = "select IncentiveProductGroupID from CPE_IncentiveProductGroups with (NoLock) " & _
                            "where RewardOptionID=" & roid & " and ProductGroupID=1 and Deleted=0;"
        dt = MyCommon.LRT_Select
        If dt.Rows.Count > 0 Then
            AnyProductUsed = True
        End If

        'Find the product groups for this ROID
        MyCommon.QueryStr = "select IncentiveProductGroupID, ProductGroupID from CPE_IncentiveProductGroups with (NoLock) where Deleted=0 and ExcludedProducts=0 and RewardOptionID=" & roid & ";"
        dt = MyCommon.LRT_Select()
        If dt.Rows.Count > 0 Then
            ProductGroups = dt.Rows.Count
            If ProductGroups >= 2 Then
                'There's more than one group, so don't allow accumulation
                ShowAccum = False
            ElseIf ProductGroups = 1 Then
                'There's one group, and this is it, so allow accumulation
                If IncentiveProdGroupID = dt.Rows(0).Item("IncentiveProductGroupID") Then
                    ShowAccum = True
                Else
                    ShowAccum = False
                End If
            End If
        Else
            'There are no groups so allow accumulation
            ProductGroups = 0
            ShowAccum = True
        End If

        If IncentiveProdGroupID = 0 Then
            ProdID = -1
            'ExProdID = -1
            If GetCgiValue("EnableAccum") = "1" Then
                EnableAccum = True
            ElseIf GetCgiValue("EnableAccum") = "0" Then
                EnableAccum = False
            End If
        Else
            'AMS-684 Get exclusion product groups
            resultExcludedPGList = m_ProductConditionPGService.GetExclusionProductGroups(IncentiveProdGroupID)
            If (resultExcludedPGList.ResultType = AMSResultType.Success) Then
                exclusionPGList = resultExcludedPGList.Result
            Else
                infoMessage = resultExcludedPGList.PhraseString
            End If

            MyCommon.QueryStr = "select ProductGroupID,AccumMin from CPE_IncentiveProductGroups with (NoLock) where RewardOptionID=" & roid & " and ExcludedProducts=0 and Deleted=0 and IncentiveProductGroupID=" & IncentiveProdGroupID & ";"
            rst = MyCommon.LRT_Select()
            If rst.Rows.Count > 0 Then
                ProdID = MyCommon.NZ(rst.Rows(0).Item("ProductGroupID"), -1)

                'ExProdID = -1
                If MyCommon.NZ(rst.Rows(0).Item("AccumMin"), 0) > 0 Then
                    EnableAccum = True
                Else
                    EnableAccum = False
                End If
            Else
                'AMS-684 Extracted the Any product group setting logic which was based on if there is any excluded product group present and included group not present in above if
                If Not exclusionPGList Is Nothing And exclusionPGList.Count > 0 Then
                    ProdID = 1
                End If
                'MyCommon.QueryStr = "select ProductGroupID, AccumMin from CPE_IncentiveProductGroups with (NoLock) where RewardOptionID=" & roid & " and ExcludedProducts=1 and IncentiveProductGroupID=" & IncentiveProdGroupID & ";"
                'dt = MyCommon.LRT_Select
                'If dt.Rows.Count > 0 Then
                '    ProdID = 1
                '    ExProdID = dt.Rows(0).Item("ProductGroupID")
                '    If MyCommon.NZ(dt.Rows(0).Item("AccumMin"), 0) > 0 Then
                '        EnableAccum = True
                '    Else
                '        EnableAccum = False
                '    End If
                'Else
                '    EnableAccum = True
                'End If
            End If

            If GetCgiValue("EnableAccum") = "1" Then
                EnableAccum = True
            ElseIf GetCgiValue("EnableAccum") = "0" Then
                EnableAccum = False
            End If
        End If

        If IncentiveProdGroupID > 0 Then
            MyCommon.QueryStr = "select Rounding from CPE_IncentiveProductGroups where RewardOptionID=" & roid & " and IncentiveProductGroupID=" & IncentiveProdGroupID & ";"
            dt = MyCommon.LRT_Select
            If dt.Rows.Count > 0 Then
                If dt.Rows(0).Item("Rounding") = "1" Then
                    Rounding = True
                End If
            End If
        End If
        objOffer = m_OfferService.GetOffer(OfferID, LoadOfferOptions.None)
        If (objOffer IsNot Nothing) Then
            If (AttributePGEnabled AndAlso objOffer.BuyerID IsNot Nothing) Then
                ucProductAttributeFilter.BuyerID = objOffer.BuyerID
            End If
            Name = objOffer.OfferName
            isTemplate = objOffer.IsTemplate
            FromTemplate = objOffer.FromTemplate
        End If
        ' see if someone is saving
        If (GetCgiValue("save") <> "" And roid > 0) Then
            'Tier level validation code
            If TierLevels > 1 Then
                For t = 2 To TierLevels
                    t2 = MyCommon.Extract_Val(GetCgiValue("t" & t & "_limit"))
                    t1 = MyCommon.Extract_Val(GetCgiValue("t" & t - 1 & "_limit"))
                    If t2 > t1 Then
                        ValidTier = True
                    Else
                        ValidTier = False
                        Exit For
                    End If
                Next
            Else
                ValidTier = True
            End If
            If Disqualifier Then
                ValidTier = True
            End If
            'AMS-684 repalced ExProdID<=0 below
            TenderType = IIf(ProdID = 1 AndAlso (exclusionPGList Is Nothing Or (Not exclusionPGList Is Nothing And exclusionPGList.Count <= 0)) AndAlso Type = 2 AndAlso Not Disqualifier, TenderType, 0)

            'Rounding validation code
            If GetCgiValue("Rounding") = "on" Then
                Rounding = True
            Else
                Rounding = False
            End If
            'ReturnedItemGroup validation code
            If GetCgiValue("itemReturnedGroup") = "on" Then
                ReturnedItemChecked = True
            Else
                ReturnedItemChecked = False
            End If
            If Rounding Then
                For t = 1 To TierLevels
                    If CInt(MyCommon.Extract_Val(GetCgiValue("t" & t & "_limit"))) <> MyCommon.Extract_Val(GetCgiValue("t" & t & "_limit")) Then
                        ValidRounding = False
                    End If
                Next
            End If
            'Accumulation validation code
            If GetCgiValue("EnableAccum") = "" Then
                EnableAccum = False
            End If
            ValidAccum = False   'Set default for the valid accumulation
            If EnableAccum Then
                AccumMin = If(MyCommon.Extract_Val(GetCgiValue("accummin")) <> "", MyCommon.Extract_Val(GetCgiValue("accummin")), 0)
                If AccumMin > 0 Then
                    ValidAccum = True
                End If
            Else
                ValidAccum = True
            End If

            ValidPMR = False
            Dim newValues As String = ""
            For t = 1 To TierLevels
                newValues &= MyCommon.Extract_Val(GetCgiValue("t" & t & "_limit")) & ","
            Next

            Type = If(MyCommon.Extract_Val(GetCgiValue("select")) <> 0, MyCommon.Extract_Val(GetCgiValue("select")), 0)
            TenderType = MyCommon.Extract_Val(GetCgiValue("selectTenderType"))
            If ValidProximityMessageProductCondtionExist(MyCommon, roid, TierLevels, newValues, Type, IncentiveProdGroupID) Then
                ValidPMR = True
            End If
            'Al-5916
            If ExistProductPriceConditionGCRConflict(MyCommon, roid, IncentiveProdGroupID, Type, ExistGCRPercentOff(MyCommon, roid), gcrErrorPhrase) Then
                ValidGCR = False
            End If

            ' Validation for Node 
            If AttributePGEnabled Then
                If NodeID = "" AndAlso radiobtnlistpgselection.SelectedItem.Value = "2" AndAlso GetCgiValue("selGroups") = "" AndAlso Not istemplate Then
                    ValidNodeId = False
                End If
            End If
            'AMS-684 removed bUseMultipleProductExclusionGroups
            If (((GetCgiValue("selGroups") = "" AndAlso GetCgiValue("exGroups") <> "") OrElse (GetCgiValue("selGroups") = "" AndAlso GetCgiValue("exGroups") = "")) And (Not isTemplate And GetCgiValue("require_pg")="")) Then
                If AttributePGEnabled AndAlso radiobtnlistpgselection.SelectedItem.Value <> "2" Then    'Exclude the case of Value=2 as then an attribute based group is being saved which has default name
                    ValidMultipleExclusionProdCondition = False
                End If
            End If
            'CLOUDSOL-2939:Bay UAT/Saks UAT -    Template Value disappears when locking a required product condition
            If (GetCgiValue("selGroups") = "" And GetCgiValue("require_pg") <> "" And GetCgiValue("Disallow_Edit") <> "") Then
                ValidMultipleExclusionProdCondition = False
            End If
               
            If roid > 0 And ValidTier And ValidAccum And ValidRounding And ValidPMR And ValidGCR And ValidNodeId And ValidMultipleExclusionProdCondition Then
                If (Not IsValidEntry(MyCommon, Localization, roid, TierLevels)) Then
                    infoMessage = Copient.PhraseLib.Lookup("term.invalidnumericentry", LanguageID)
                Else
                    If (AttributePGEnabled) AndAlso (radiobtnlistpgselection.SelectedItem.Value = "2") Then
                        If (AttributeProductGroupID = 0) Then
                            Dim productgroup As New ProductGroup
                            'productgroup.ProductGroupName = String.Concat(Copient.PhraseLib.Lookup("term.offer", LanguageID), " ", OfferID.ToString(), " ", Copient.PhraseLib.Lookup("term.conditionalproducts", LanguageID).ToLower())
                            productgroup.ProductGroupName = ProductGroupName
                            productgroup.AnyProduct = False
                            productgroup.ProductGroupTypeID = 2
                            productgroup.BuyerID = BuyerID
                            Dim ProductID As AMSResult(Of Int64) = m_ProductGroupService.CreateProductGroup(productgroup)
                            If ProductID.ResultType <> AMSResultType.Success Then
                                infoMessage = ProductID.MessageString
                            Else
                                AttributeProductGroupID = ProductID.Result
                                m_ActivityLogService.Activity_Log(ActivityTypes.ProductGroup, ProductID.Result, AdminUserID, Copient.PhraseLib.Lookup("history.pgroup-create", LanguageID))
                            End If

                        Else
                            If ProductGroupName <> m_ProductGroupService.GetProductGroupName(AttributeProductGroupID).Result Then
                                m_ProductGroupService.UpdateProductGroupName(AttributeProductGroupID, ProductGroupName)
                            End If
                        End If
                        If (IncentiveProdGroupID > 0 AndAlso radiobtnlistpgselection.SelectedItem.Value = "2") Then
                            If AttributeProductGroupID = m_ProductGroupService.GetProductGroupID(IncentiveProdGroupID).Result Then
                                If ProductGroupName <> m_ProductGroupService.GetProductGroupName(m_ProductGroupService.GetProductGroupID(IncentiveProdGroupID).Result).Result Then
                                    m_ProductGroupService.UpdateProductGroupName(m_ProductGroupService.GetProductGroupID(IncentiveProdGroupID).Result, ProductGroupName)
                                End If
                            End If
                        End If
                        ProdID = AttributeProductGroupID
                        ucProductAttributeFilter.ProductGroupID = ProdID
                        ucProductAttributeFilter.SelectedNodeIDs = NodeID
                        ucProductAttributeFilter.SaveData = True

                    ElseIf GetCgiValue("selGroups") <> "" Then
                        ProdID = MyCommon.Extract_Val(GetCgiValue("selGroups"))
                    Else
                        ProdID = -1
                    End If

                    ' Check to see if a product condition is required by the template, if applicable
                    MyCommon.QueryStr = "select ProductGroupID from CPE_IncentiveProductGroups with (NoLock) where RewardOptionID=" & roid & _
                                        " and RequiredFromTemplate=1 and Deleted=0 and ExcludedProducts=0;"
                    rst = MyCommon.LRT_Select
                    HasRequiredPG = (rst.Rows.Count > 0)

                    MyCommon.QueryStr = "select IncentiveTenderID from CPE_IncentiveTenderTypes with (NoLock) where Deleted = 0 and RewardOptionID = " & roid
                    rst = MyCommon.LRT_Select
                    HasTenderCondition = (rst.Rows.Count > 0)

                    MyCommon.QueryStr = "select IPG.IncentiveProductGroupID, PG.ProductGroupID, PG.Name, PG.PhraseID, PG.AnyProduct, UT.PhraseID as UnitPhraseID, " & _
                                        " ExcludedProducts, ProductComboID, QtyForIncentive, QtyUnitType, AccumMin, AccumLimit, AccumPeriod, DisallowEdit, " & _
                                        " RequiredFromTemplate, Disqualifier, Rounding, MinPurchAmt, ReturnedItemGroup, MinItemPrice from CPE_IncentiveProductGroups as IPG with (NoLock) " & _
                                        " left join ProductGroups as PG with (NoLock) on PG.ProductGroupID=IPG.ProductGroupID " & _
                                        " left join CPE_RewardOptions as RO with (NoLock) on IPG.RewardOptionID=RO.RewardOptionID " & _
                                        " left join CPE_UnitTypes as UT with (NoLock) on UT.UnitTypeID=IPG.QtyUnitType " & _
                                        " where IPG.RewardOptionID=" & roid & "and IPG.Deleted=0 and Disqualifier=1 " & _
                                        " order by Name;"
                    rst = MyCommon.LRT_Select
                    HasDisqualifier = (rst.Rows.Count > 0)

                    If (HasTenderCondition AndAlso GetCgiValue("selGroups") = "") AndAlso (ProductGroupTypeID <> 2) Then
                        infoMessage = Copient.PhraseLib.Lookup("ueoffer-con-product.TenderRequirement", LanguageID)
                    ElseIf (HasDisqualifier AndAlso GetCgiValue("selGroups") = "") And Not Disqualifier AndAlso (ProductGroupTypeID <> 2) Then
                        infoMessage = Copient.PhraseLib.Lookup("ueoffer-con-product.ProdDisqualierRequirement", LanguageID)
                    Else
                        'Find IncentiveProductGroupID and use it to delete the records from the tiers table
                        If IncentiveProdGroupID <> 0 Then
                            MyCommon.QueryStr = "delete from CPE_IncentiveProductGroupTiers where IncentiveProductGroupID=" & IncentiveProdGroupID
                            MyCommon.LRT_Execute()
                        End If
                    End If

                    ' we need to do some work to set the limit values if there are any, otherwise just set to 0
                    ' in theory there should be one set of limit values for each selected groups and possibly an accumulation infos
                    If Not EnableAccum OrElse Disqualifier Then
                        AccumMin = 0
                        AccumLimit = 0
                        AccumPeriod = 0
                    Else
                        AccumMin = If(MyCommon.Extract_Val(GetCgiValue("accummin")) <> 0, MyCommon.Extract_Val(GetCgiValue("accummin")), 0)
                        AccumLimit = If(MyCommon.Extract_Val(GetCgiValue("accumlimit")) <> 0, MyCommon.Extract_Val(GetCgiValue("accumlimit")), 0)
                        AccumPeriod = If(MyCommon.Extract_Val(GetCgiValue("accumperiod")) <> 0, MyCommon.Extract_Val(GetCgiValue("accumperiod")), 0)
                    End If
                    Qty = If(MyCommon.Extract_Val(GetCgiValue("t1_limit")) <> 0, MyCommon.Extract_Val(GetCgiValue("t1_limit")), 0)
                    UniqueProduct = IIf(GetCgiValue("ItemRestriction") = ProductFilterEnum.Unique.ToString(),1,0)
                    NetPriceProduct = If(MyCommon.Extract_Val(GetCgiValue("NetPrice")) <> 0, MyCommon.Extract_Val(GetCgiValue("NetPrice")), 0)
                    'xxx If SupportSameItemProductCondition Then
                    SameItem = IIf(GetCgiValue("ItemRestriction") = ProductFilterEnum.SameItem.ToString(),1,0)
                    'xxx End If
                    'AMS-1055
                    'NoRestriction = If(MyCommon.Extract_Val(GetCgiValue("NoRestriction")) <> 0, MyCommon.Extract_Val(GetCgiValue("NoRestriction")), 0)
                    If TierLevels > 1 Or EnableAccum Then
                        MinPurchAmt = 0
                        MinItemPrice = 0
                    Else
                        MinPurchAmt = MyCommon.Extract_Decimal(GetCgiValue("MinPurchAmt"), MyCommon.GetAdminUser.Culture)
                        MinItemPrice = MyCommon.Extract_Decimal(GetCgiValue("MinItemPrice"), MyCommon.GetAdminUser.Culture)
                    End If

                    MinPurchAmt = Math.Round(MinPurchAmt, CurrencyPrecision)
                    MinItemPrice = Math.Round(MinItemPrice, CurrencyPrecision)
                    Send("<!-- Qty pre-round: " & Qty & " -->")
                    Qty = Localization.Round_Quantity(Qty, roid, Type)
                    Send("<!-- Qty post-round: " & Qty & " -->")

                    '     System.IO.File.AppendAllText("F:\UELogix\logs\conprod.txt", "selGroups " & (GetCgiValue("selGroups") <> "") & ControlChars.CrLf)
                    ' lets handle the selected first
                    If (GetCgiValue("selGroups") <> "") OrElse (GetCgiValue("require_pg") <> "") OrElse (ProductGroupTypeID = 2 AndAlso AttributeProductGroupID > 0) Then
                        '      System.IO.File.AppendAllText("F:\UELogix\logs\conprod.txt", "point 1 " & ControlChars.CrLf)
                        If (AttributePGEnabled AndAlso ProductGroupTypeID = 2) Then
                            historyString = Copient.PhraseLib.Lookup("term.alteredproductgroups", LanguageID) & ": " & GetCgiValue("selGroups")
                        Else
                            historyString = Copient.PhraseLib.Lookup("term.alteredproductgroups", LanguageID) & ": " & AttributeProductGroupID
                        End If

                        If (UniqueProduct = 1) Then
                            IsUniqueProd = True
                        End If

                        '     System.IO.File.AppendAllText("F:\UELogix\logs\conprod.txt", "IncentiveProdGroupID=" & IncentiveProdGroupID & ControlChars.CrLf)
                        TempIncentiveProdGroupID = IncentiveProdGroupID
                        If IncentiveProdGroupID = 0 Then 'AndAlso (Not GetCgiValue("require_pg") <> "") Then
                            'If this is a product disqualifier then then Qty=0 and Type=1
                            MyCommon.QueryStr = "dbo.pa_CPE_AddProductGroup"
                            MyCommon.Open_LRTsp()
                            MyCommon.LRTsp.Parameters.Add("@ROID", SqlDbType.Int, 4).Value = roid
                            MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.Int, 4).Value = ProdID
                            MyCommon.LRTsp.Parameters.Add("@QtyForIncentive", SqlDbType.Decimal, 12).Value = IIf(Disqualifier, 1, Qty)
                            MyCommon.LRTsp.Parameters.Add("@QtyUnitType", SqlDbType.Int, 4).Value = IIf(Disqualifier, 1, Type)
                            MyCommon.LRTsp.Parameters.Add("@AccumMin", SqlDbType.Decimal, 12).Value = AccumMin
                            MyCommon.LRTsp.Parameters.Add("@AccumLimit", SqlDbType.Decimal, 12).Value = AccumLimit
                            MyCommon.LRTsp.Parameters.Add("@AccumPeriod", SqlDbType.Decimal, 12).Value = AccumPeriod
                            MyCommon.LRTsp.Parameters.Add("@UniqueProduct", SqlDbType.Bit).Value = UniqueProduct
                            MyCommon.LRTsp.Parameters.Add("@NetPriceProduct", SqlDbType.Bit).Value = NetPriceProduct
                            MyCommon.LRTsp.Parameters.Add("@RequiredFromTemplate", SqlDbType.Bit).Value = IIf(HasRequiredPG, 1, 0)
                            MyCommon.LRTsp.Parameters.Add("@Disqualifier", SqlDbType.Bit).Value = IIf(Disqualifier, 1, 0)
                            MyCommon.LRTsp.Parameters.Add("@Rounding", SqlDbType.Bit).Value = IIf(Rounding, 1, 0)
                            MyCommon.LRTsp.Parameters.Add("@MinPurchAmt", SqlDbType.Decimal, 15).Value = MinPurchAmt
                            MyCommon.LRTsp.Parameters.Add("@ReturnedItemGroup", SqlDbType.Bit).Value = IIf(ReturnedItemChecked, 1, 0)
                            MyCommon.LRTsp.Parameters.Add("@TenderType", SqlDbType.Int).Value = TenderType
                            MyCommon.LRTsp.Parameters.Add("@IncentiveProductGroupID", SqlDbType.Int, 4).Direction = ParameterDirection.Output
                            MyCommon.LRTsp.ExecuteNonQuery()
                            IncentiveProdGroupID = MyCommon.LRTsp.Parameters("@IncentiveProductGroupID").Value
                            If (TempIncentiveProdGroupID = 0) Then TempIncentiveProdGroupID = IncentiveProdGroupID
                            MyCommon.Close_LRTsp()

                            ' save the price filter settings.
                            If IncentiveProdGroupID > 0 Then
                                MyCommon.QueryStr = "update CPE_IncentiveProductGroups with (RowLock) set FullPrice=" & GetPriceFilterToken("FullPrice") & ", " & _
                                                    "  ClearanceState=" & GetPriceFilterToken("ClearanceState") & ", ClearanceLevel=" & GetPriceFilterToken("ClearanceLevel") & ", " & _
                                                    "  MinItemPrice= " & MinItemPrice & " " & " , SameItem=" & SameItem & " "
                                'xxx MyCommon.QueryStr &= IIf(SupportSameItemProductCondition, " , SameItem=" & SameItem & " ", "")

                                MyCommon.QueryStr &= "where IncentiveProductGroupID=" & IncentiveProdGroupID & " and RewardOptionID=" & roid & ";"
                                MyCommon.LRT_Execute()
                            End If
                        Else
                            MyCommon.QueryStr = "update CPE_IncentiveProductGroups set ProductGroupID=" & If(ProdID = -1, "NULL", ProdID) & ", QtyForIncentive=" & If(Disqualifier, 1, Qty) & ", QtyUnitType=" & If(Disqualifier, 1, Type) & ", " & _
                                                "AccumMin=" & AccumMin & ", AccumLimit=" & AccumLimit & ", AccumPeriod=" & AccumPeriod & ", ExcludedProducts=0, " & _
                                                "RequiredFromTemplate=" & If(HasRequiredPG, "1", "0") & ", TCRMAStatusFlag=3, Disqualifier=" & If(Disqualifier, "1", "0") & ", " & _
                                                "UniqueProduct=" & UniqueProduct & ", Rounding=" & If(Rounding, "1", "0") & ", MinPurchAmt=" & MinPurchAmt & ", NetPriceProduct=" & NetPriceProduct & ", ReturnedItemGroup=" & If(ReturnedItemChecked, "1", "0") & ", " & _
                                                "FullPrice=" & GetPriceFilterToken("FullPrice") & ", ClearanceState=" & GetPriceFilterToken("ClearanceState") & ", " & _
                                                "ClearanceLevel=" & GetPriceFilterToken("ClearanceLevel") & ", MinItemPrice=" & MinItemPrice & ",TenderType= " & TenderType &" , SameItem=" & SameItem & " "
                            'xxx  MyCommon.QueryStr &= IIf(SupportSameItemProductCondition, " , SameItem=" & SameItem & " ", "")
                            MyCommon.QueryStr &= "where IncentiveProductGroupID=" & IncentiveProdGroupID & " and RewardOptionID=" & roid & ";"
                            MyCommon.LRT_Execute()
                        End If
                        MyCommon.QueryStr = "SELECT PG.ProductGroupID FROM ProductGroups PG  WITH (NoLock)INNER JOIN CPE_IncentiveProductGroups CPG WITH (Nolock) on PG.ProductGroupID= CPG.ProductGroupID WHERE cpg.IncentiveProductGroupID=" & IncentiveProdGroupID & ";"
                        rst = MyCommon.LRT_Select
                        If rst.Rows.Count > 0 Then
                            ProductGroupID = Convert.ToInt64(MyCommon.NZ(rst.Rows(0).Item("ProductGroupID"), 0))
                        End If



                        'Saving Tiers
                        If IncentiveProdGroupID <> 0 Then
                            MyCommon.QueryStr = "delete from CPE_IncentiveProductGroupTiers where RewardOptionID=" & roid & " and IncentiveProductGroupID=" & IncentiveProdGroupID
                            MyCommon.LRT_Execute()
                        Else
                            'MyCommon.QueryStr = "select IncentiveProductGroupID from CPE_IncentiveProductGroups where RewardOptionID=" & roid & " and ProductGroupID=" & ProdID & " and deleted=0"
                            'TierDT = MyCommon.LRT_Select()
                            'IncentiveProdGroupID = TierDT.Rows(0).Item("IncentiveProductGroupID")
                        End If
                        If Disqualifier Then
                            'Qty is now set to 0 for a disqualifier
                            TierQty = MyCommon.Extract_Val(GetCgiValue("t1_limit"))
                            MyCommon.QueryStr = "dbo.pa_CPE_AddProductGroupTiers"
                            MyCommon.Open_LRTsp()
                            MyCommon.LRTsp.Parameters.Add("@ROID", SqlDbType.Int, 4).Value = roid
                            MyCommon.LRTsp.Parameters.Add("@IncentiveProductGroupID", SqlDbType.Int, 4).Value = IncentiveProdGroupID
                            MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int, 4).Value = 1
                            MyCommon.LRTsp.Parameters.Add("@Qty", SqlDbType.Decimal, 12).Value = 1 'TierQty
                            MyCommon.LRTsp.ExecuteNonQuery()
                            MyCommon.Close_LRTsp()
                        Else
                            For t = 1 To TierLevels
                                TierQty = MyCommon.Extract_Decimal(GetCgiValue("t" & t & "_limit"), MyCommon.GetAdminUser.Culture)
                                TierQty = Localization.Round_Quantity(TierQty, roid, Type)
                                MyCommon.QueryStr = "dbo.pa_CPE_AddProductGroupTiers"
                                MyCommon.Open_LRTsp()
                                MyCommon.LRTsp.Parameters.Add("@ROID", SqlDbType.Int, 4).Value = roid
                                MyCommon.LRTsp.Parameters.Add("@IncentiveProductGroupID", SqlDbType.Int, 4).Value = IncentiveProdGroupID
                                MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int, 4).Value = t
                                MyCommon.LRTsp.Parameters.Add("@Qty", SqlDbType.Decimal, 15).Value = TierQty
                                MyCommon.LRTsp.ExecuteNonQuery()
                                MyCommon.Close_LRTsp()
                            Next
                        End If
                        'Set offer to limit frequency to once per transaction
                        If GetCgiValue("ItemRestriction") = ProductFilterEnum.Unique.ToString Then
                            MyCommon.QueryStr = "update CPE_Incentives set P3DistQtyLimit=1, P3DistTimeType=2, P3DistPeriod=1 where IncentiveID=" & OfferID
                            MyCommon.LRT_Execute()
                        End If
                    ElseIf HasRequiredPG Then
                        If IncentiveProdGroupID = 0 Then
                            MyCommon.QueryStr = "insert into CPE_IncentiveProductGroups (RewardOptionID,ProductGroupID,ExcludedProducts,Deleted,LastUpdate,QtyForIncentive,QtyUnitType,AccumMin,AccumLimit,AccumPeriod," & _
                                                "  RequiredFromTemplate,TCRMAStatusFlag,Disqualifier,UniqueProduct,Rounding,MinPurchAmt, ReturnedItemGroup,FullPrice,ClearanceState,ClearanceLevel, MinItemPrice,TenderType,NetPriceProduct,SameItem)"
                            'xxx MyCommon.QueryStr &= IIf(SupportSameItemProductCondition, "SameItem", "")
                            'xxx MyCommon.QueryStr &= ")"
                            MyCommon.QueryStr &= " values (" & roid & "," & ProdID & ",0,0,getdate()," & IIf(Disqualifier, 1, Qty) & "," & IIf(Disqualifier, 1, Type) & "," & AccumMin & "," & AccumLimit & "," & AccumPeriod & _
                                       ",1,3," & IIf(Disqualifier, "1", "0") & "," & UniqueProduct & "," & IIf(Rounding, "1", "0") & "," & MinPurchAmt & "," & IIf(ReturnedItemChecked, "1", "0") & "," & GetPriceFilterToken("FullPrice") & _
                                       "," & GetPriceFilterToken("ClearanceState") & "," & GetPriceFilterToken("ClearanceLevel") & "," & MinItemPrice & TenderType & "," & NetPriceProduct & "," & SameItem & ")"
                            'xxxMyCommon.QueryStr &= IIf(SupportSameItemProductCondition, "," & SameItem, "")
                            'xxx MyCommon.QueryStr &= ")"
                            MyCommon.LRT_Execute()
                        Else
                            MyCommon.QueryStr = "update CPE_IncentiveProductGroups set ProductGroupID=" & ProdID & ",QtyForIncentive=" & IIf(Disqualifier, 1, Qty) & "," & _
                                                "QtyUnitType=" & IIf(Disqualifier, 1, Type) & ",AccumMin=" & AccumMin & ",AccumLimit=" & AccumLimit & ",AccumPeriod=" & AccumPeriod & "," & _
                                                "RequiredFromTemplate=" & IIf(HasRequiredPG, "1", "0") & ",TCRMAStatusFlag=3,Disqualifier=" & IIf(Disqualifier, "1", "0") & "," & _
                                                "UniqueProduct=" & UniqueProduct & ", NetPriceProduct=" & NetPriceProduct & ", Rounding=" & IIf(Rounding, "1", "0") & ", MinPurchAmt=" & MinPurchAmt & ", ReturnedItemGroup=" & IIf(ReturnedItemChecked, "1", "0") & ", " & _
                                                "FullPrice=" & GetPriceFilterToken("FullPrice") & ", ClearanceState=" & GetPriceFilterToken("ClearanceState") & ", " & _
                                                "ClearanceLevel=" & GetPriceFilterToken("ClearanceLevel") & ", MinItemPrice= " & MinItemPrice & ",TenderType= " & TenderType & " , SameItem=" & SameItem
                            'xxx MyCommon.QueryStr &= IIf(SupportSameItemProductCondition, " , SameItem=" & SameItem & " ", "")
                            MyCommon.QueryStr &= "where IncentiveProductGroupID=" & IncentiveProdGroupID & " and RewardOptionID=" & roid & ";"
                            MyCommon.LRT_Execute()
                        End If
                    End If

                    ' check to see if a product condition is required by the template, if applicable
                    MyCommon.QueryStr = "select ProductGroupID from CPE_IncentiveProductGroups where RewardOptionID=" & roid & _
                                        " and RequiredFromTemplate=1 and Deleted=0 and ExcludedProducts=1;"
                    rst = MyCommon.LRT_Select
                    HasRequiredPG = (rst.Rows.Count > 0)

                    'AMS-684 Handle the excluded product groups
                    Dim excludedPGlist As String = If((AttributePGEnabled AndAlso radiobtnlistpgselection.SelectedItem.Value = "2"), GetCgiValue("exGroups_attr"), GetCgiValue("exGroups"))
                    If (excludedPGlist <> "") Then
                        historyString = historyString & " " & Copient.PhraseLib.Lookup("term.excluding", LanguageID) & ": " & excludedPGlist
                    End If
                    'Ids = GetCgiValue("exGroups").Split(",")
                    Dim dtExclusionPG As DataTable = GetExclusionPGDataTable(excludedPGlist)

                    Dim resultExPGSave As AMSResult(Of Boolean) = m_ProductConditionPGService.SaveExclusionGroups(IncentiveProdGroupID, dtExclusionPG)
                    If resultExPGSave.ResultType = AMSResultType.Exception Then
                        infoMessage = resultExPGSave.PhraseString
                    End If

                    MyCommon.QueryStr = "update CPE_Incentives set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID
                    MyCommon.LRT_Execute()
                    ResetOfferApprovalStatus(OfferID)

                    'check if accumulation message needs to be removed
                    MyCommon.QueryStr = "dbo.pa_CPE_AccumMsgEligible"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@ROID", SqlDbType.Int, 4).Value = roid
                    MyCommon.LRTsp.Parameters.Add("@AccumEligible", SqlDbType.Bit, 1).Direction = ParameterDirection.Output
                    MyCommon.LRTsp.ExecuteNonQuery()
                    AccumEligible = MyCommon.LRTsp.Parameters("@AccumEligible").Value
                    MyCommon.Close_LRTsp()

                    If Not (AccumEligible) Then
                        'Mark any accumulation messages as deleted
                        MyCommon.QueryStr = "update CPE_Deliverables set Deleted = 1 where DeliverableID in " & _
                                            "(select D.DeliverableID from CPE_RewardOptions RO inner join CPE_Deliverables D on RO.RewardOptionID = D.RewardOptionID " & _
                                            "where RO.Deleted = 0 and D.Deleted = 0 and RO.IncentiveID = " & OfferID & " and RewardOptionPhase = 2 and DeliverableTypeID = 4);"
                        'MyCommon.QueryStr = "update CPE_Deliverables set Deleted=1 where RewardOptionId=" & roid & " and RewardOptionPhase=2 and deleted=0;"
                        MyCommon.LRT_Execute()
                    End If

                    'Finally: accumulation and disqualification are mutually incompatible within an offer, so:
                    'if this condition is a disqualifier, disable accumulation on every product condition in this offer, or
                    'if this condition has accumulation, delete any associated disqualifiers from this offer.
                    If Disqualifier Then
                        MyCommon.QueryStr = "update CPE_IncentiveProductGroups set AccumMin=0, AccumLimit=0, AccumPeriod=0 " & _
                                            "where RewardOptionID=" & roid & ";"
                        MyCommon.LRT_Execute()
                    ElseIf (AccumMin > 0 OrElse AccumLimit > 0 OrElse AccumPeriod > 0) Then
                        MyCommon.QueryStr = "delete from CPE_IncentiveProductGroupTiers where IncentiveProductGroupID in " & _
                                            "(select IncentiveProductGroupID from CPE_IncentiveProductGroups where RewardOptionID=" & roid & " and Deleted=0 and Disqualifier=1) " & _
                                            "and RewardOptionID=" & roid & ";"
                        MyCommon.LRT_Execute()
                        MyCommon.QueryStr = "update CPE_IncentiveProductGroups with (RowLock) set Deleted=1, LastUpdate=getdate(), TCRMAStatusFlag=3 " & _
                                            "where RewardOptionID=" & roid & " and Disqualifier=1;"
                        MyCommon.LRT_Execute()
                    End If

                    MyCommon.Activity_Log(3, OfferID, AdminUserID, historyString)
                End If
            Else
                If Not ValidTier Then
                    infoMessage = Copient.PhraseLib.Lookup("error.tiervalues", LanguageID)
                    'IncentiveProdGroupID = 0
                ElseIf Not ValidAccum Then
                    infoMessage = Copient.PhraseLib.Lookup("error.validaccum", LanguageID)
                    'IncentiveProdGroupID = 0
                ElseIf Not ValidRounding Then
                    infoMessage = Copient.PhraseLib.Lookup("error.rounding", LanguageID)
                    'IncentiveProdGroupID = 0
                ElseIf Not ValidPMR Then
                    If ValidPMRAwayValues Then
                        infoMessage = Copient.PhraseLib.Lookup("condition.affectspmrvalue", LanguageID)
                    Else
                        infoMessage = Copient.PhraseLib.Lookup("prod-con.undefinedquantityaway", LanguageID)
                    End If
                ElseIf Not ValidGCR Then
                    infoMessage = Copient.PhraseLib.Lookup(gcrErrorPhrase, LanguageID)
                ElseIf Not ValidNodeId Then
                    infoMessage = Copient.PhraseLib.Lookup("error.noselectednode", LanguageID)
                ElseIf Not ValidMultipleExclusionProdCondition Then
                    infoMessage = Copient.PhraseLib.Lookup("term.select", LanguageID) & " " & Copient.PhraseLib.Lookup("term.includedgroups", LanguageID)
                Else
                    infoMessage = Copient.PhraseLib.Lookup("term.errornorewardoption", LanguageID)
                    'IncentiveProdGroupID = 0
                End If
            End If

            If (infoMessage = "") Then
                CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
            Else
                CloseAfterSave = False
            End If
        End If

        ' dig the offer info out of the database
        ' no one clicked anything
        'objOffer = m_OfferService.GetOffer(OfferID, LoadOfferOptions.None)
        'If (objOffer IsNot Nothing) Then
        '    If (AttributePGEnabled AndAlso objOffer.BuyerID IsNot Nothing) Then
        '        ucProductAttributeFilter.BuyerID = objOffer.BuyerID
        '    End If
        '    Name = objOffer.OfferName
        '    isTemplate = objOffer.IsTemplate
        '    FromTemplate = objOffer.FromTemplate
        'End If

        'update the templates permission if necessary
        If (GetCgiValue("save") <> "" AndAlso GetCgiValue("IsTemplate") = "IsTemplate") Then

            ' time to update the status bits for the templates
            Dim form_Disallow_Edit As Integer = 0
            Dim form_Require_PG As Integer = 0

            If (GetCgiValue("Disallow_Edit") = "on") Then
                form_Disallow_Edit = 1
            End If

            If (GetCgiValue("require_pg") <> "" Or GetCgiValue("require_pg_Attr") <> "") Then
                form_Require_PG = 1
            End If
            If (form_Disallow_Edit = 1 AndAlso form_Require_PG = 1) Then
                infoMessage = Copient.PhraseLib.Lookup("offer-con.lockeddenied", LanguageID)
                MyCommon.QueryStr = "update CPE_IncentiveProductGroups set DisallowEdit=1, RequiredFromTemplate=0 " & _
                                    " where RewardOptionID=" & roid & " and Deleted=0 and Disqualifier=" & IIf(Disqualifier, "1", "0")
                'AMS-684 removed bUseMultipleProductExclusionGroups        
                'If (bUseMultipleProductExclusionGroups) Then
                MyCommon.QueryStr &= " and IncentiveProductGroupID =" & TempIncentiveProdGroupID
                'End If
                MyCommon.QueryStr &= ";"
            Else
                MyCommon.QueryStr = "update CPE_IncentiveProductGroups set DisallowEdit=" & form_Disallow_Edit & ", RequiredFromTemplate=" & form_Require_PG & _
                                    " where RewardOptionID=" & roid & " and Deleted=0 and Disqualifier=" & IIf(Disqualifier, "1", "0")
                'AMS-684 removed bUseMultipleProductExclusionGroups
                'If (bUseMultipleProductExclusionGroups) Then
                MyCommon.QueryStr &= " and IncentiveProductGroupID =" & TempIncentiveProdGroupID
                'End If
                MyCommon.QueryStr &= ";"
            End If
            MyCommon.LRT_Execute()

            ' If necessary, create an empty product condition
            If (form_Require_PG = 1) Then
                MyCommon.QueryStr = "select ProductGroupID from CPE_IncentiveProductGroups with (NoLock) " & _
                                    " where RewardOptionID=" & roid & " and Deleted=0;"
                rst = MyCommon.LRT_Select
                If (rst.Rows.Count = 0) Then
                    ' Create the product condition record, then the product condition tier record(s)
                    MyCommon.QueryStr = "insert into CPE_IncentiveProductGroups (RewardOptionID,QtyForIncentive,QtyUnitType,ExcludedProducts,Deleted,LastUpdate,RequiredFromTemplate,TCRMAStatusFlag)" & _
                                        " values(" & roid & "," & MyCommon.Extract_Val(GetCgiValue("t1_limit")) & "," & MyCommon.Extract_Val(GetCgiValue("select")) & ",0,0,getdate(),1,3);"
                    MyCommon.LRT_Execute()
                    MyCommon.QueryStr = "select top 1 IncentiveProductGroupID from CPE_IncentiveProductGroups where RewardOptionID=" & roid & " order by LastUpdate DESC;"
                    rst = MyCommon.LRT_Select()
                    If rst.Rows.Count > 0 Then
                        For t = 1 To TierLevels
                            TierQty = MyCommon.Extract_Val(GetCgiValue("t" & t & "_limit"))
                            MyCommon.QueryStr = "dbo.pa_CPE_AddProductGroupTiers"
                            MyCommon.Open_LRTsp()
                            MyCommon.LRTsp.Parameters.Add("@ROID", SqlDbType.Int, 4).Value = roid
                            MyCommon.LRTsp.Parameters.Add("@IncentiveProductGroupID", SqlDbType.Int, 4).Value = rst.Rows(0).Item("IncentiveProductGroupID")
                            MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int, 4).Value = t
                            MyCommon.LRTsp.Parameters.Add("@Qty", SqlDbType.Decimal, 12).Value = TierQty
                            MyCommon.LRTsp.ExecuteNonQuery()
                            MyCommon.Close_LRTsp()
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
            MyCommon.QueryStr = "select DisallowEdit, RequiredFromTemplate from CPE_IncentiveProductGroups with (NoLock) " & _
                                " where RewardOptionID=" & roid & " and Deleted=0 and Disqualifier=" & IIf(Disqualifier, "1", "0")
            'AMS-684 removed bUseMultipleProductExclusionGroups                                
            'If (bUseMultipleProductExclusionGroups) Then
            MyCommon.QueryStr &= " and IncentiveProductGroupID =" & IncentiveProdGroupID
            'End If

            MyCommon.QueryStr &= ";"

            rst = MyCommon.LRT_Select
            If (rst.Rows.Count > 0) Then
                Disallow_Edit = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit"), False)
                RequirePG = MyCommon.NZ(rst.Rows(0).Item("RequiredFromTemplate"), False)
            Else
                Disallow_Edit = False
            End If
        End If
        Dim m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
        If Not isTemplate Then
            DisabledAttribute = IIf(Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And Disallow_Edit), "", "disabled=""disabled""")
        Else
            DisabledAttribute = IIf(Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer, "", "disabled=""disabled""")
        End If
        If (Not String.IsNullOrEmpty(DisabledAttribute)) Then
            txtProductGroupName.Disabled = True
        End If

        If (AttributePGEnabled) Then
            If String.IsNullOrWhiteSpace(AttributeSwitchType) AndAlso AttributeProductGroupID = 0 And Not Page.IsPostBack AndAlso ProdID > 0 Then
                Dim productidtemp As AMSResult(Of Byte) = m_ProductGroupService.GetProductGroupType(ProdID)
                If productidtemp.ResultType <> AMSResultType.Success Then
                    infoMessage = productidtemp.MessageString
                ElseIf (productidtemp.Result = 2) Then
                    AttributeProductGroupID = ProdID
                    ProductGroupTypeID = 2
                Else
                    ProductGroupTypeID = 1
                End If
            ElseIf AttributeSwitchType = "SelectedAttributeGroup" AndAlso String.IsNullOrWhiteSpace(GetCgiValue("save")) Then
                AttributeProductGroupID = MyCommon.Extract_Val(GetCgiValue("selGroups"))
                ProductGroupTypeID = 2
                ucProductAttributeFilter.HierarchyTreeURL = "/logix/phierarchytree.aspx?ProductGroupID=" & AttributeProductGroupID & "&PAB=1&OfferID=" & OfferID & "&ConditionID=" & ConditionID & "&Disqualifier=" & Disqualifier & "&AttributeProductGroupID=" & AttributeProductGroupID '& locateHierarchyURL
                ucProductAttributeFilter.IsAttributeSwitch = True
            ElseIf (AttributeSwitchType = "DeSelectedAttributeGroup" AndAlso String.IsNullOrWhiteSpace(GetCgiValue("save"))) Or (GetCgiValue("save") <> "" AndAlso AttributeSwitchType="") Then
                AttributeProductGroupID = 0
                ucProductAttributeFilter.HierarchyTreeURL = "/logix/phierarchytree.aspx?ProductGroupID=-1&PAB=1&OfferID=" & OfferID & "&ConditionID=" & ConditionID & "&Disqualifier=" & Disqualifier & "&AttributeProductGroupID=" & AttributeProductGroupID '& locateHierarchyURL
                ProductGroupTypeID = 1
                ucProductAttributeFilter.IsAttributeSwitch = True
                'below Line is added to reload Product Hierarchy '
            ElseIf AttributeSwitchType = "" AndAlso Not ValidNodeId Then
                AttributeProductGroupID = 0
                ucProductAttributeFilter.HierarchyTreeURL = "/logix/phierarchytree.aspx?ProductGroupID=-1&PAB=1&OfferID=" & OfferID & "&ConditionID=" & ConditionID & "&Disqualifier=" & Disqualifier & "&AttributeProductGroupID=" & AttributeProductGroupID '& locateHierarchyURL
                ProductGroupTypeID = 2
                ucProductAttributeFilter.IsAttributeSwitch = True
            End If
            SelectAttributeType(ProductGroupTypeID)
            ucProductAttributeFilter.IsPGAttributeType = True
            ucProductAttributeFilter.ProductGroupID = AttributeProductGroupID
            ucProductAttributeFilter.LanguageID = LanguageID
            m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, OfferID))
            If Not isTemplate Then
                ucProductAttributeFilter.IsEditPermitted = (Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And Disallow_Edit))
                radiobtnlistpgselection.Enabled = (Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (FromTemplate And Disallow_Edit))
            Else
                ucProductAttributeFilter.IsEditPermitted = Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer
                radiobtnlistpgselection.Enabled = Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer
            End If
            '         MyCommon.QueryStr = "SELECT PG.ProductGroupID FROM ProductGroups PG  WITH (NoLock)INNER JOIN CPE_IncentiveProductGroups CPG WITH (Nolock) on PG.ProductGroupID= CPG.ProductGroupID WHERE cpg.IncentiveProductGroupID=" & IncentiveProdGroupID & ";"
            '   rst = MyCommon.LRT_Select
            '   If rst.Rows.Count > 0 Then
            '       ProductGroupID = Convert.ToInt64(MyCommon.NZ(rst.Rows(0).Item("ProductGroupID"), 0))
            '   End If
            rst = Hierarchy.GetNodesLinkedToProductGroupID(IIf(AttributeProductGroupID = 0, IIf(GetCgiValue("selGroups") <> "", MyCommon.Extract_Val(GetCgiValue("selGroups")), 0), AttributeProductGroupID))
            LinkedItems = ""
            If (rst.Rows.Count > 0) Then
                For i As Integer = 0 To rst.Rows.Count - 1
                    LinkedItems = LinkedItems & rst.Rows(i)("ExtNodeID").ToString()
                    LinkedItems = LinkedItems & If((i < rst.Rows.Count - 1), ",", String.Empty)
                Next
            End If
            'If Not String.IsNullOrWhiteSpace(GetCgiValue("LocateHierarchyURL")) Then
            '    locateHierarchyURL = HttpUtility.UrlDecode(GetCgiValue("LocateHierarchyURL"))
            '    'PABStage = Convert.ToInt32(GetCgiValue("PABStage"))
            '    SelectAttributeType(2)
            'End If
        End If

        If Not Disqualifier Then
            Send_HeadBegin("term.offer", "term.productcondition", OfferID)
        Else
            Send_HeadBegin("term.offer", "term.productdisqualifier", OfferID)
        End If
        Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
        Send_Metas()
        Send_Links(Handheld)
        Send_Scripts()

        Send("<script " & "type=""text/javascript"">")
        Send("var fullSelect = null;")
        Send("var fullSelect_attr = null;")
        Send("  // This is the javascript array holding the function list")
        Send("  // The PrintJavascriptArray ASP function can be used to print this array.")
        MyCommon.QueryStr = "select ProductGroupID,Buyerid, Name from ProductGroups where ProductGroupID is not null " & _
                             "and Deleted=0 and NonDiscountableGroup=0 and PointsNotApplyGroup=0 "
        If (MyCommon.IsUserAssociatedWithAnyBuyer(AdminUserID) AndAlso Not Logix.UserRoles.ViewProductgroupRegardlessBuyer) Then
            MyCommon.QueryStr &= "and (BuyerId in(select BuyerId from BuyerRoleUsers where AdminUserID=" & AdminUserID & ") or BuyerId is null)"
        End If

        If Disqualifier Then
            MyCommon.QueryStr &= " and ProductGroupID <> 1  and ProductGroupID not in " & _
                                 "   (select ProductGroupID from CPE_IncentiveProductGroups with (NoLock) where Deleted=0 and RewardOptionID=" & roid & " and Disqualifier=0 and ExcludedProducts=0)"
        Else
            MyCommon.QueryStr &= " and ProductGroupID not in " & _
                                 "   (select ProductGroupID from CPE_IncentiveProductGroups with (NoLock) where Deleted=0 and RewardOptionID=" & roid & " and Disqualifier=1 and ExcludedProducts=0)"
        End If
        MyCommon.QueryStr &= " order by AnyProduct desc, Name asc"

        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
            Sendb("var functionlist = Array(")
            'If Not Disqualifier AndAlso Not AnyProductUsed Then
            '  Sendb("""" & Copient.PhraseLib.Lookup("term.anyproduct", LanguageID) & """,")
            'End If
            For Each row In rst.Rows
                Sendb("""" & MyCommon.NZ(row.Item("Name"), "").ToString().Replace("""", "\""") & """,")
            Next
            Send(""""");")
            Sendb("var vallist = Array(")
            'If Not Disqualifier AndAlso Not AnyProductUsed Then
            '  Sendb("""" & "1" & """,")
            'End If
            For Each row In rst.Rows
                Sendb("""" & MyCommon.NZ(row.Item("ProductGroupID"), 0) & """,")
            Next
            Send(""""");")
        Else
            Sendb("var functionlist = Array(")
            If Not Disqualifier AndAlso Not AnyProductUsed Then
                Send("""" & Copient.PhraseLib.Lookup("term.anyproduct", LanguageID) & """);")
            Else
                Send(");")
            End If
            Sendb("var vallist = Array(")
            If Not Disqualifier AndAlso Not AnyProductUsed Then
                Send("""" & "1" & """);")
            Else
                Send(");")
            End If
        End If
        'Populate Attribute based Product Group IDs so that Logix Knows which Group Type is being selected/deselected and it can perform appropriate action for the same.
        Sendb("var attributepglist = [")
        Sendb(String.Join(",", lstAttributeProductGroups))
        Send("];")
        Send("</" & "script>")
    End Sub

    Sub SelectAttributeType(ByVal ProductGroupTypeID As String)
        If (radiobtnlistpgselection.Visible = False) Then
            Return
        End If
        radiobtnlistpgselection.ClearSelection()
        Dim radioBtn As ListItem = radiobtnlistpgselection.Items.FindByValue(ProductGroupTypeID)
        If radioBtn Is Nothing Then
            radiobtnlistpgselection.Items(0).Selected = True
        Else
            radioBtn.Selected = True
        End If

    End Sub
    Private Shared Function PrepareProductGroupHTML(ByRef MyCommon As Copient.CommonInc, ByVal dtPG As DataTable, ByVal LanguageID as integer) As String
        Dim pgHTML As New StringBuilder()
        Dim dt As New DataTable()
        dt = dtPG
        For Each pgRow In dt.Rows
            If MyCommon.NZ(pgRow.Item("ProductGroupID"), 0) = 1 Then
                pgHTML.AppendLine("<option value=""1"" style=""color: brown; font-weight: bold;"">" & Copient.PhraseLib.Lookup("term.anyproduct", LanguageID) & "</option>")
            ElseIf (pgRow.Item("ProductGroupID") IsNot Nothing AndAlso pgRow.Item("ProductGroupTypeId") = 2) Then
                If (MyCommon.IsEngineInstalled(9) And MyCommon.Fetch_UE_SystemOption(168) = "1" And Not IsDBNull(pgRow.Item("Buyerid"))) Then
                    Dim buyerid As Integer = pgRow.Item("Buyerid")
                    Dim externalBuyerid = MyCommon.GetExternalBuyerId(buyerid)
                    pgHTML.AppendLine("<option value=""" & MyCommon.NZ(pgRow.Item("ProductGroupID"), 0) & """  title=""" & MyCommon.NZ(pgRow.Item("Name"), "") & """ style=""color: blue;"">" & "Buyer " & externalBuyerid & " - " & MyCommon.NZ(pgRow.Item("Name"), "") & "</option>")
                Else
                    pgHTML.AppendLine("<option value=""" & MyCommon.NZ(pgRow.Item("ProductGroupID"), 0) & """  title=""" & MyCommon.NZ(pgRow.Item("Name"), "") & """ style=""color: blue;"">" & MyCommon.NZ(pgRow.Item("Name"), "") & "</option>")
                End If

            Else
                If (MyCommon.IsEngineInstalled(9) And MyCommon.Fetch_UE_SystemOption(168) = "1" And Not IsDBNull(pgRow.Item("Buyerid"))) Then
                    Dim buyerid As Integer = pgRow.Item("Buyerid")
                    Dim externalBuyerid = MyCommon.GetExternalBuyerId(buyerid)
                    pgHTML.AppendLine("<option value=""" & MyCommon.NZ(pgRow.Item("ProductGroupID"), 0) & """ title=""" & MyCommon.NZ(pgRow.Item("Name"), "") & """ >" & "Buyer " & externalBuyerid & " - " & MyCommon.NZ(pgRow.Item("Name"), "") & "</option>")
                Else
                    pgHTML.AppendLine("<option value=""" & MyCommon.NZ(pgRow.Item("ProductGroupID"), 0) & """ title=""" & MyCommon.NZ(pgRow.Item("Name"), "") & """>" & MyCommon.NZ(pgRow.Item("Name"), "") & "</option>")
                End If
            End If
        Next
        Return pgHTML.ToString
    End Function

    <WebMethod>
    Public Shared Function GetProductGroupListJSON(ByVal bEnableRestrictedAccessToUEOfferBuilder As Boolean, ByVal lastPGName As String, ByVal AdminUserID As Integer, ByVal roid As Long, ByVal Disqualifier As Boolean, ByVal viewProductgroupRegardlessBuyer As Boolean, ByVal LanguageID As Integer) As String
        Dim MyCommon As New Copient.CommonInc()
        Dim dt As DataTable = GetProductGroupListHTML(MyCommon, bEnableRestrictedAccessToUEOfferBuilder, lastPGName, AdminUserID, roid, Disqualifier, viewProductgroupRegardlessBuyer, LanguageID, True, True)
        Return PrepareProductGroupHTML(MyCommon, dt, LanguageID)
    End Function
    Public Shared Function GetProductGroupListHTML(ByRef MyCommon As Copient.CommonInc, ByVal bEnableRestrictedAccessToUEOfferBuilder As Boolean, ByVal lastPGName As String,
    ByVal AdminUserID As Integer, ByVal roid As Long, ByVal Disqualifier As Boolean, ByVal viewProductgroupRegardlessBuyer As Boolean, ByVal LanguageID As Integer, ByVal IsAjaxCall As Boolean, ByVal shouldFetchPGAsync As Boolean) As DataTable
        Dim listSize As Integer = MyCommon.Fetch_SystemOption(290)
        If (shouldFetchPGAsync) Then
            MyCommon.QueryStr = "select top " & listSize & " ProductGroupID,buyerid,ProductGroupTypeId, Name from ProductGroups where ProductGroupID is not null " & _
                        "and Deleted=0 and NonDiscountableGroup=0 and PointsNotApplyGroup=0 "
        Else
            MyCommon.QueryStr = "select ProductGroupID,buyerid,ProductGroupTypeId, Name from ProductGroups where ProductGroupID is not null " & _
                     "and Deleted=0 and NonDiscountableGroup=0 and PointsNotApplyGroup=0 "
        End If
        If (MyCommon.IsUserAssociatedWithAnyBuyer(AdminUserID) AndAlso Not viewProductgroupRegardlessBuyer) Then
            MyCommon.QueryStr &= "and (BuyerId in(select BuyerId from BuyerRoleUsers where AdminUserID=" & AdminUserID & ") or BuyerId is null)"
        End If

        If Disqualifier Then
            MyCommon.QueryStr &= " and ProductGroupID <> 1  and ProductGroupID not in " & _
                                 "   (select ProductGroupID from CPE_IncentiveProductGroups with (NoLock) where Deleted=0 and RewardOptionID=" & roid & " and Disqualifier=0 and ExcludedProducts=0)"
        Else
            MyCommon.QueryStr &= " and ProductGroupID not in " & _
                                 "   (select ProductGroupID from CPE_IncentiveProductGroups with (NoLock) where Deleted=0 and RewardOptionID=" & roid & " and Disqualifier=1 and ExcludedProducts=0)"
        End If
        If (bEnableRestrictedAccessToUEOfferBuilder) Then MyCommon.QueryStr &= " and isnull(TranslatedFromOfferID,0) = 0 "
        MyCommon.QueryStr &= " and Name > '" & lastPGName & "'"
        If Not IsAjaxCall Then
            MyCommon.QueryStr &= " order by AnyProduct desc, Name asc"
        Else
            MyCommon.QueryStr &= " order by Name asc"
        End If
        Dim dt As DataTable = MyCommon.LRT_Select
        Return dt

    End Function
    Function GetExclusionPGDataTable(ByVal excludedPGList As String) As DataTable
        Dim dt As New DataTable
        Dim dc As New DataColumn("ProductGroupId")
        Dim dr As DataRow
        Dim  testInt As Integer
        dt.Columns.Add(dc)
        For Each exPGID In excludedPGList.Split(",")
            If (Integer.TryParse(exPGID.Trim, testInt)) Then    'For preventing XSS attack
                dr = dt.NewRow
                dr("ProductGroupId") = exPGID.Trim
                dt.Rows.Add(dr)
            End If
        Next

        Return dt
    End Function
    Function IsValidEntry(ByRef MyCommon As Copient.CommonInc, ByRef Localization As Copient.Localization, ByVal RewardOptionID As Long, ByVal TierLevels As Integer) As Boolean
        Dim Qty As Decimal
        Dim Type As Integer
        Dim AccumMin As Decimal
        Dim AccumLimit As Decimal
        Dim AccumPeriod As Integer
        Dim MinPurchAmt As Decimal
        Dim MinItemPrice As Decimal
        Dim IsValid As Boolean = True
        Dim ProdID As Integer = 0
        Dim t As Integer
        Dim TierQty As Decimal

        If (GetCgiValue("selGroups") <> "") Then
            'Ids = GetCgiValue("selGroups").Split(",")
            ProdID = MyCommon.Extract_Val(GetCgiValue("selGroups"))

            ' we need to do some work to set the limit values if there are any, otherwise just set to 0
            ' in theory there should be one set of limit values for each selected groups and possibly an accumulation infos
            Qty = MyCommon.Extract_Decimal(GetCgiValue("limit"), MyCommon.GetAdminUser.Culture)
            Type = MyCommon.Extract_Decimal(GetCgiValue("select"), MyCommon.GetAdminUser.Culture)
            AccumMin = MyCommon.Extract_Decimal(GetCgiValue("accummin"), MyCommon.GetAdminUser.Culture)
            AccumLimit = MyCommon.Extract_Decimal(GetCgiValue("accumlimit"), MyCommon.GetAdminUser.Culture)
            AccumPeriod = MyCommon.Extract_Decimal(GetCgiValue("accumperiod"), MyCommon.GetAdminUser.Culture)
            MinPurchAmt = MyCommon.Extract_Decimal(GetCgiValue("MinPurchAmt"), MyCommon.GetAdminUser.Culture)
            MinItemPrice = MyCommon.Extract_Decimal(GetCgiValue("MinItemPrice"), MyCommon.GetAdminUser.Culture)

            IsValid = IsValid AndAlso IsProperFormat(Localization, Type, Qty, RewardOptionID)
            IsValid = IsValid AndAlso IsProperFormat(Localization, Type, AccumMin, RewardOptionID)
            IsValid = IsValid AndAlso IsProperFormat(Localization, Type, AccumLimit, RewardOptionID)
            IsValid = IsValid AndAlso IsProperFormat(Localization, Type, AccumPeriod, RewardOptionID)
            IsValid = IsValid AndAlso IsProperFormat(Localization, 2, MinPurchAmt, RewardOptionID)
            IsValid = IsValid AndAlso IsProperFormat(Localization, 2, MinItemPrice, RewardOptionID)

            If TierLevels > 1 Then
                For t = 1 To TierLevels
                    TierQty = MyCommon.Extract_Val(GetCgiValue("t" & t & "_limit"))
                    IsValid = IsValid AndAlso IsProperFormat(Localization, Type, TierQty, RewardOptionID)
                Next
            End If

            'If (Not IsValid) Then Exit For
        End If

        Return IsValid
    End Function

    '------------------------------------------------------------------------------------------------------------------------------------------

    Function IsProperFormat(ByRef Localization As Copient.Localization, ByVal UnitType As Integer, ByVal Value As Double, ByVal RewardOptionID As Long) As Boolean
        Dim FormatOk As Boolean = True
        Dim StrValue As String
        Dim DecPtPos, CharAfterDec As Integer

        'Send("<!-- comparing " & Value & "  with 1000000000 ... UnitType=" & UnitType & "  RewardOptionID=" & RewardOptionID & " -->")
        If Value >= 1000000000 Then 'numbers >= 1 billion will blow up the insert into the database which is defined as decimal (15,6)
            Return False
        End If
        StrValue = Value.ToString
        DecPtPos = StrValue.IndexOf(".")
        If (DecPtPos > -1) Then
            CharAfterDec = (StrValue.Length - (DecPtPos + 1))
        End If

        Select Case UnitType
            Case 1 ' ###,##0
                FormatOk = (DecPtPos = -1)
            Case 2 ' ###,##0.00
                FormatOk = (DecPtPos = -1) OrElse (CharAfterDec >= 0 AndAlso CharAfterDec <= CurrencyPrecision)
            Case 3 ' ###,##0.000
                FormatOk = (DecPtPos = -1) OrElse (CharAfterDec >= 0 AndAlso CharAfterDec <= 3)
            Case 5, 6, 7, 8
                FormatOk = (DecPtPos = -1) OrElse (CharAfterDec >= 0 AndAlso CharAfterDec <= Localization.GetCached_UOM_Precision(RewardOptionID, UnitType, Copient.Localization.UOMUsageEnum.UnitType))
            Case Else
                FormatOk = True
        End Select

        Return FormatOk
    End Function

    '------------------------------------------------------------------------------------------------------------------------------------------

    Function GetPriceFilterToken(ByVal TokenName As String) As Integer
        Dim Value As Integer = 0
        Dim PriceFilter As String

        PriceFilter = GetCgiValue("pricefilter")

        If PriceFilter IsNot Nothing AndAlso PriceFilter.Length >= 3 Then
            Select Case TokenName.ToUpper
                Case "FULLPRICE"
                    Integer.TryParse(PriceFilter.Substring(0, 1), Value)
                Case "CLEARANCESTATE"
                    Integer.TryParse(PriceFilter.Substring(1, 1), Value)
                Case "CLEARANCELEVEL"
                    Integer.TryParse(PriceFilter.Substring(2), Value)
            End Select
        End If

        Return Value
    End Function

    '------------------------------------------------------------------------------------------------------------------------------------------
    Function ExistGCRPercentOff(ByRef MyCommon As Copient.CommonInc, ByVal ROID As Long) As Boolean
        Dim percentOffGCRLinked As Boolean = False
        Dim dt As DataTable
        MyCommon.QueryStr = "Select distinct gc.Id From CPE_Deliverables cd join GiftCard gc on (cd.OutputId=gc.Id) join GiftCardTier gct on (gc.id=gct.GiftCardId) " & _
                                                " Where gct.AmountTypeId=@AmountTypeId And cd.RewardOptionid=@ROID And DeliverableTypeId=@DeliverableTypeId And cd.Deleted=0;"
        MyCommon.DBParameters.Add("@ROID", SqlDbType.BigInt).Value = ROID
        MyCommon.DBParameters.Add("@DeliverableTypeId", SqlDbType.BigInt).Value = DELIVERABLE_TYPES.GIFTCARD
        MyCommon.DBParameters.Add("@AmountTypeId", SqlDbType.BigInt).Value = CPEAmountTypes.PercentageOff
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            percentOffGCRLinked = True
        End If
        Return percentOffGCRLinked
    End Function
    Function ExistGcrPercentOffMinimumConditional(ByRef MyCommon As Copient.CommonInc, ByVal ROID As Long) As Boolean
        Dim percentOffGCRMinimumConditionalLinked As Boolean = False
        Dim dt As DataTable
        MyCommon.QueryStr = "Select gc.Id From CPE_Deliverables cd join GiftCard gc on (cd.OutputId=gc.Id) join GiftCardTier gct on (gc.id=gct.GiftCardId) " & _
                            " Where gct.AmountTypeId=@AmountTypeId And cd.RewardOptionid=@ROID And DeliverableTypeId=@DeliverableTypeId And gct.ProrationTypeId=@ProrationTypeId And cd.Deleted=0;"
        MyCommon.DBParameters.Add("@ROID", SqlDbType.BigInt).Value = ROID
        MyCommon.DBParameters.Add("@DeliverableTypeId", SqlDbType.BigInt).Value = DELIVERABLE_TYPES.GIFTCARD
        MyCommon.DBParameters.Add("@AmountTypeId", SqlDbType.BigInt).Value = CPEAmountTypes.PercentageOff
        MyCommon.DBParameters.Add("@ProrationTypeId", SqlDbType.Int).Value = UEProrationTypes.MinimumConditionalItems
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            percentOffGCRMinimumConditionalLinked = True
        End If
        Return percentOffGCRMinimumConditionalLinked
    End Function
    Function ExistProductPriceConditionGCRConflict(ByRef MyCommon As Copient.CommonInc, ByVal ROID As Long, ByVal IncentiveProductGroupId As Long,
                                                   ByVal NewUnitType As Int32, ByVal GcrPercentOffExists As Boolean, ByRef GCRErrorPhrase As String) As Boolean
        Dim priceConditionGCRConflictFlag As Boolean = False
        Dim dt As DataTable
        MyCommon.QueryStr = "Select QtyUnitType from dbo.CPE_IncentiveProductGroups WHERE RewardOptionID=@ROID AND Deleted=0"
        MyCommon.DBParameters.Add("@ROID", SqlDbType.BigInt).Value = ROID
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        Dim productConditionPriceTypeCount As Int16 = 0
        Dim productConditionOthersCount As Int16 = 0
        Dim productCondition As Boolean = False
        If dt.Rows.Count > 0 Then
            productCondition = True

            For Each row1 In dt.Rows
                If CInt(row1("QtyUnitType")) = CPEUnitTypes.Dollars Then
                    productConditionPriceTypeCount += 1
                End If
                If CInt(row1("QtyUnitType")) <> CPEUnitTypes.Dollars Then
                    productConditionOthersCount += 1
                End If
            Next
            If IncentiveProductGroupId = 0 Then 'New condition
                If productConditionPriceTypeCount = 0 AndAlso NewUnitType = CPEUnitTypes.Dollars AndAlso GcrPercentOffExists AndAlso GetOfferLimit(MyCommon, ROID) <> 1 Then
                    priceConditionGCRConflictFlag = True
                    GCRErrorPhrase = "offer-gen.badlimitduetoGCRPercentOff"
                End If
                If productConditionPriceTypeCount > 0 AndAlso NewUnitType <> CPEUnitTypes.Dollars AndAlso GcrPercentOffExists Then
                    priceConditionGCRConflictFlag = True
                    GCRErrorPhrase = "error.productconditionadd"
                End If
                If productConditionOthersCount > 0 AndAlso NewUnitType = CPEUnitTypes.Dollars AndAlso GcrPercentOffExists Then    'Mixing of Price condition with others not allowed for %off gcr
                    priceConditionGCRConflictFlag = True
                    GCRErrorPhrase = "error.productconditionadd"
                End If
                If productConditionOthersCount > 0 AndAlso NewUnitType = CPEUnitTypes.Dollars AndAlso GcrPercentOffExists Then    'Mixing of Price condition with others not allowed for %off gcr
                    priceConditionGCRConflictFlag = True
                    GCRErrorPhrase = "error.productconditionadd"
                End If
            Else
                If productConditionPriceTypeCount > 1 AndAlso NewUnitType <> CPEUnitTypes.Dollars AndAlso GcrPercentOffExists Then
                    priceConditionGCRConflictFlag = True
                    GCRErrorPhrase = "error.affectsGCR"
                End If
                If productConditionOthersCount > 1 AndAlso NewUnitType = CPEUnitTypes.Dollars AndAlso GcrPercentOffExists Then
                    priceConditionGCRConflictFlag = True
                    GCRErrorPhrase = "error.affectsGCR"
                End If
                If (productConditionPriceTypeCount = 1 Or productConditionOthersCount = 1) AndAlso NewUnitType = CPEUnitTypes.Dollars AndAlso GcrPercentOffExists Then
                    If GetOfferLimit(MyCommon, ROID) <> 1 Then
                        priceConditionGCRConflictFlag = True
                        GCRErrorPhrase = "error.pricecondition-offerlimit-conflict"
                    ElseIf ExistGcrPercentOffMinimumConditional(MyCommon, ROID) Then
                        priceConditionGCRConflictFlag = True
                        GCRErrorPhrase = "error.pricecondition-prorationtype-conflict"
                    End If
                End If
            End If
        End If

        Return priceConditionGCRConflictFlag
    End Function
    Function GetOfferLimit(ByRef MyCommon As Copient.CommonInc, ByVal ROID As Long) As Int32
        Dim offerLimit As Int32
        Dim rst2 As DataTable
        MyCommon.QueryStr = "Select P3DistQtyLimit from dbo.CPE_Incentives ci (nolock) join dbo.CPE_RewardOptions ro (nolock) on (ci.IncentiveId=ro.IncentiveId) WHERE ro.RewardOptionId = @ROID AND ro.Deleted=0"
        MyCommon.DBParameters.Add("@ROID", SqlDbType.Int).Value = ROID
        rst2 = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)

        If rst2.Rows.Count > 0 AndAlso rst2.Rows(0)(0) <> Nothing Then
            offerLimit = CInt(rst2.Rows(0)(0))
        End If
        Return offerLimit
    End Function

    Function ValidProximityMessageProductCondtionExist(ByRef Common As Copient.CommonInc, ByVal ROID As Integer, ByVal TierLevels As Integer, ByVal NewValues As String, ByVal NewQtyType As Integer, ByVal IncentiveProductGroupId As Integer) As Boolean
        Dim validPMRProdCondition As Boolean = True
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

            If thresholdType <> CPEUnitTypes.Points Then                'Validate for PMR with Product conditions only
                Common.QueryStr = "select IncentiveProductGroupID from CPE_IncentiveProductGroups where Deleted=0 And RewardOptionID= " & ROID
                dpg = Common.LRT_Select

                If dpg.Rows.Count > 0 And IncentiveProductGroupId = 0 Then 'New one is being added
                    validPMRProdCondition = False
                ElseIf IncentiveProductGroupId <> 0 Then   'Existing one being edited
                    If thresholdType <> NewQtyType Then
                        validPMRProdCondition = False
                    ElseIf Not ValidPMRTierValues(Common, NewValues) Then
                        validPMRProdCondition = False
                    End If
                End If
            End If
            count += 1
        Next
        Return validPMRProdCondition
    End Function

    Function ValidPMRTierValues(ByVal Common As Copient.CommonInc, ByVal NewValues As String) As Boolean
        Dim Values As String() = NewValues.Split(New Char() {","c})
        Dim dst As DataTable
        Dim dt As DataTable
        Dim validTierValues As Boolean = True
        Dim tier0Value As Decimal
        Dim validTriggers As Boolean = True
        Common.QueryStr = "select PM.ThresholdTypeID, PMT.TriggerValue from ProximityMessageTier as PMT " & _
            "inner join ProximityMessage as PM " & _
            "on PM.ID = PMT.ProximityMessageId " & _
            "inner join CPE_Deliverables as CPED " & _
            "on CPED.OutputID = PM.ID where CPED.DeliverableTypeID = 14 and CPED.Deleted = 0 and CPED.RewardOptionID = " & roid
        dt = Common.LRT_Select
        If (TierLevels > 1) Then
            If (dt.Rows.Count > 0 AndAlso dt.Rows.Count < TierLevels) Then
                validTriggers = False
            End If
        End If
        tier0Value = Common.NZ(dt.Rows(0).Item("TriggerValue"), 0)
        Common.QueryStr = "select IPGT.Quantity from CPE_IncentiveProductGroupTiers as IPGT " & _
                            "where IPGT.RewardOptionID = " & roid

        dst = Common.LRT_Select
        If Decimal.Parse(Values(0)) <= tier0Value Then
            validTierValues = False
        End If
        If (dst.Rows.Count = TierLevels AndAlso TierLevels > 1) Then
            If validTriggers Then
                For i As Integer = 1 To TierLevels - 1 Step 1
                    Dim newDiff As Integer = Decimal.Parse(Values(i)) - Decimal.Parse(Values(i - 1))
                    If (newDiff < Decimal.Parse(Common.NZ(dt.Rows(i).Item("TriggerValue"), 0))) Then
                        validTierValues = False
                    End If
                Next
            Else
                validTierValues = False
                ValidPMRAwayValues = False
            End If
        End If
        Return validTierValues
    End Function
    'Function to be called asynchronously by javascript for fetching paged attribute values
    <WebMethod>
    <WebInvoke(Method:="POST")>
    Public Shared Function GetAttributes(term As String, attributetype As Int16, pageindex As Int32, excludeattr As String, keyValue As String) As List(Of Attributes)
        CurrentRequest.Resolver.AppName = "UEoffer-con-product.aspx"
        Dim attributeService As IAttributeService = CurrentRequest.Resolver.Resolve(Of IAttributeService)()
        Dim nodeIdsDT As DataTable = HttpContext.Current.Session("AllChildNodes")
        If HttpContext.Current.Session("AllChildNodes") IsNot Nothing Then
            nodeIdsDT = CType(HttpContext.Current.Session("AllChildNodes"), DataTable)
        End If
        Dim amsResultAttributes As AMSResult(Of List(Of Attributes)) = attributeService.GetAttributesInChunks(term, attributetype, pageindex, nodeIdsDT, excludeattr, keyValue)
        If amsResultAttributes.ResultType <> AMSResultType.Success AndAlso amsResultAttributes.MessageString <> String.Empty Then
            Dim activityFields As ActivityLogFields = New ActivityLogFields
            Dim myCommon As New Copient.CommonInc
            myCommon.Open_LogixRT()
            activityFields.LinkID = attributetype
            activityFields.Description = amsResultAttributes.MessageString
            myCommon.Activity_Log3(activityFields)
            myCommon.Close_LogixRT()
        End If
        Return amsResultAttributes.Result
    End Function

</script>
<%If Not AttributePGEnabled Then %>
<script type="text/javascript" src="../../javascript/jquery.min.js"></script>
<% End If %>
<script type="text/javascript">
    var fetchPGAsync = <%= IIf(shouldFetchPGAsync, 1, 0)%>;
<% If (CloseAfterSave) Then %>
    // window.opener.location.reload();
    //AMS-5916:Product Condition page will popup after clicking the Save button
    opener.location = "/logix/UE/UEoffer-con.aspx?OfferID=<%Sendb(OfferID)%>";
    window.close();
<% Else %>
    if (document.getElementById("functionselect") != null) {
        fullSelect = document.getElementById("functionselect").cloneNode(true);
        fullSelect_attr = document.getElementById("functionselect_attr").cloneNode(true)
    }
    removeUsed(true);
    removeUsed_attr(true);
  <% If (String.IsNullOrWhiteSpace(DisabledAttribute)) Then %>
    updateButtons();
    updateButtons_attr();
  <% End If %>
    ProductGroupTypeSelection();
    ValidateSave();
    ChangeUnit(document.getElementById("select").options[document.getElementById("select").selectedIndex].value);
<% End If %>

    window.onload = function () { 
        if( $("#radiobtnlistpgselection input[type=radio]:checked").val() == 2){  
            ShoworHideDivs();
        }
    }

    function DisableMinItemPrice(){
        CheckTenderType();
        CheckMinItemPrice();
        CheckpriceFilter();
    }

    function CheckMinItemPrice(){
        if(parseFloat($('#MinItemPrice').val()) > 0)
            $("#selectTenderType").attr('disabled',true);
        else
        {
            if($("#priceFilter option:selected").val() >0)
                $("#selectTenderType").attr('disabled',true);
            else
                $("#selectTenderType").attr('disabled',false);
        }
    }

    //Update the item price and Price filter controls based on tender type value 
    function CheckTenderType(){
        if($("#selectTenderType option:selected").val() >0){
            $('#MinItemPrice').val("");
            $("#priceFilter").val("000");
            $("#MinItemPrice").attr('disabled',true);
            $("#priceFilter").attr('disabled',true);
        }
        else{
            $("#MinItemPrice").attr('disabled',false);
            $("#priceFilter").attr('disabled',false);
        }
    }

    //AMS-832 Tender Type field and Price filter field should be mutually exclusive
    function CheckpriceFilter(){
        if($("#priceFilter option:selected").val() >0){
            $("#selectTenderType").val("0");
            $("#selectTenderType").attr('disabled',true);
        }
        else{
            CheckMinItemPrice();
        }
    }
    //Get the last PG name in the select list box
    function GetLastPGName()
    {        
        var len = $('select#functionselect option').length-1;   //Get length
        if(len > 0)   
            return $("#functionselect option")[len].text;   //Get the PG name

        return "";
    }
    $(document).ready( function ()
    {
        DisableMinItemPrice();
        $('#MinItemPrice').blur(function() {
            CheckMinItemPrice();
        });

        $('#selectTenderType').change(function() {
            CheckTenderType();
        });

        $('#priceFilter').change(function() {
            CheckpriceFilter();
        });
        RegisterPGAsyncLoadHandler();
    });
    function RegisterPGAsyncLoadHandler()
    {
        if(fetchPGAsync == 1)
        {
            $('#functionselect').on("scroll", ProductGroupLoader);
            $('#functionselect_attr').on("scroll", ProductGroupLoaderAttr);
        }
    }
    //Used during Ajax call to load Product groups
    var lblNotifierText = "";
    var loadInProgress = false;
    //Loader for the functionselect 
    function ProductGroupLoader(event){            
        if($('#functioninput').val()=="")   //There is nothing in the search box
        {
            //Only new parameter is lastPGName rest are taken from server due to the existing code reuse for AJAX call needs it
            var data = JSON.stringify({bEnableRestrictedAccessToUEOfferBuilder: <%= bEnableRestrictedAccessToUEOfferBuilder.ToString.ToLower %>, lastPGName: GetLastPGName(), AdminUserID: <%= AdminUserID %>, roid: <%= roid %>, Disqualifier: <%= Disqualifier.ToString.ToLower %>, viewProductgroupRegardlessBuyer: <%= Logix.UserRoles.ViewProductgroupRegardlessBuyer.ToString.ToLower %>, LanguageID: <%= LanguageID %>});
            LoadItemsOnScroll("functionselect", "<%=Request.Url.AbsolutePath%>/GetProductGroupListJSON", data);
        }
    }
    //Loader for the functionselect_attr
    function ProductGroupLoaderAttr(event){            
        if($('#functioninput_attr').val()=="")
        {
            var data = JSON.stringify({bEnableRestrictedAccessToUEOfferBuilder: <%= bEnableRestrictedAccessToUEOfferBuilder.ToString.ToLower %>, lastPGName: GetLastPGName(), AdminUserID: <%= AdminUserID %>, roid: <%= roid %>, Disqualifier: <%= Disqualifier.ToString.ToLower %>, viewProductgroupRegardlessBuyer: <%= Logix.UserRoles.ViewProductgroupRegardlessBuyer.ToString.ToLower %>, LanguageID: <%= LanguageID %>});
            LoadItemsOnScroll("functionselect_attr", "<%=Request.Url.AbsolutePath%>/GetProductGroupListJSON", data);
        }
    }

    function OnLoadError(response, status, error){
        lblNotifierText += "<%=Copient.PhraseLib.Lookup("term.erroronpgload", LanguageID)%>"; //"Error occurred while loading product groups";
        if(error)
            lblNotifierText += error;
        $('#lblAjaxNotification').text(lblNotifierText);
    }
    function BeforeSendSetup(request){
        loadInProgress = true;

        $('#divNotification').show();
        lblNotifierText += "<%=Copient.PhraseLib.Lookup("term.loadingpg", LanguageID)%>";//"Loading product groups...";
    }
    function OnLoadSuccess(response){
        if(response.d)
            lblNotifierText += "<%=Copient.PhraseLib.Lookup("term.done", LanguageID)%>";//"Done";
        else
            lblNotifierText = "<%=Copient.PhraseLib.Lookup("term.nomorepg", LanguageID)%>";//"No more product groups to load";
        $('#functionselect').append($.parseHTML(response.d));
        $('#functionselect_attr').append($.parseHTML(response.d));
        //Update FullSelect variable which is used in HandleKeyUp for search box text changes
        fullSelect = document.getElementById("functionselect").cloneNode(true);
        fullSelect_attr = document.getElementById("functionselect_attr").cloneNode(true);

        $('#lblAjaxNotification').text(lblNotifierText);
        loadInProgress = false;
        $('#divNotification').delay(500).fadeOut();
        lblNotifierText = "";
    }
    function ShowOrHideTenderType(){
        var selectedboxValue;
        if($('#selected').children('option').length > 0) selectedboxValue = $('#selected option:first-child').val();
        var unittype;
        if ($('#select').children('option').length > 0) unittype= $("#select option:selected").val();
        var IsNormalPGCondition = (($("#radiobtnlistpgselection").length == 0) || ($("#radiobtnlistpgselection input[type=radio]:checked").val() == 1));

        if(IsNormalPGCondition && selectedboxValue==1 && unittype==2 && $('#excluded').children('option').length<=0)
            $("#trtendertype").css('display',"");
        else{
            $("#trtendertype").css('display',"none");
            $("#MinItemPrice").attr('disabled',false);
            $("#selectTenderType").find('option:eq(0)').prop('selected', true);
            $("#priceFilter").attr('disabled',false);
        }
    }

</script>
<%
done:
    MyCommon.Close_LogixRT()
    Send_BodyEnd("mainform", "functioninput")
    MyCommon = Nothing
    Logix = Nothing
%>