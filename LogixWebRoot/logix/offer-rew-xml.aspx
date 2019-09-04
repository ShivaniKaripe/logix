<%@ Page Language="vb" validateRequest="false" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
	' *****************************************************************************
	' * FILENAME: offer-rew-xml.aspx 
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
	Dim RewardID As String
	Dim ExcludedItem As Integer
	Dim SelectedItem As Integer
	Dim NumTiers As Integer
	Dim LinkID As Long
	Dim VarID As String
	Dim RewardAmountTypeID As Integer
	Dim TriggerQty As Integer
	Dim ApplyToLimit As Integer
	Dim DoNotItemDistribute As Boolean
	Dim TransactionLevelSelected As Boolean = False
  Dim CloseAfterSave As Boolean = False
	Dim DistPeriod As Integer
	Dim UseSpecialPricing As Boolean
	Dim SPRepeatAtOccur As Integer
	Dim ValueRadio As Integer
  Dim q As Integer
	Dim x As Integer
  Dim Tiered As Integer
	Dim SponsorID As Integer
	Dim PromoteToTransLevel As Boolean
	Dim RewardLimit As Integer
	Dim RewardLimitTypeID As Integer
	Dim Disallow_Edit As Boolean = True
	Dim FromTemplate As Boolean
  Dim IsTemplate As Boolean = False
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim ProductGroupID As Integer = 0
  Dim ExcludedID As Integer = 0
  Dim BannersEnabled As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  OfferID = Request.QueryString("OfferID")
  
	RewardID = Request.QueryString("RewardID")
	NumTiers = Request.QueryString("NumTiers")
  ProductGroupID = MyCommon.Extract_Val(Request.QueryString("ProductGroupID"))
  ExcludedID = MyCommon.Extract_Val(Request.QueryString("ExcludedID"))
  
  Response.Expires = 0
  MyCommon.AppName = "offer-rew-xml.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  
  ' fetch the name
  MyCommon.QueryStr = "Select Name,IsTemplate,FromTemplate from Offers with (NoLock) where OfferID=" & OfferID
	rst = MyCommon.LRT_Select
	If (rst.Rows.Count > 0) Then
		Name = rst.Rows(0).Item("Name")
		IsTemplate = MyCommon.NZ(rst.Rows(0).Item("IsTemplate"), False)
		FromTemplate = MyCommon.NZ(rst.Rows(0).Item("FromTemplate"), False)
  End If
  
	Send_HeadBegin("term.offer", "term.xmlgenericreward", OfferID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
%>
  
<script type="text/javascript">
// This is the javascript array holding the function list
// The PrintJavascriptArray ASP function can be used to print this array.
<%
    MyCommon.QueryStr = "select ProductGroupID,Name from ProductGroups with (NoLock) where Deleted=0 order by AnyProduct desc, Name"
    rst = MyCommon.LRT_Select
    
    If (rst.rows.count>0)
        Sendb("var functionlist = Array(")
        For Each row In rst.Rows
            Sendb("""" & MyCommon.NZ(row.item("Name"), "").ToString().Replace("""", "\""") & """,")
        Next
        Send(""""");")
        Sendb("var vallist = Array(")
        For Each row In rst.Rows
            Sendb("""" & row.item("ProductGroupID") & """,")
        Next
        Send(""""");")
    Else
        Sendb("var functionlist = Array(")
        Send("""" & Copient.PhraseLib.Lookup("term.anyproduct", LanguageID) & """);")
        Sendb("var vallist = Array(")
        Send("""" & "1" & """);")
    End If
%>

// This is the function that refreshes the list after a keypress.
// The maximum number to show can be limited to improve performance with
// huge lists (1000s of entries).
// The function clears the list, and then does a linear search through the
// globally defined array and adds the matches back to the list.
function handleKeyUp(maxNumToShow)
{
    var selectObj, textObj, functionListLength;
    var i,  numShown;
    var searchPattern;
    var selectedList, excludedList;

    document.getElementById("functionselect").size = "10";
    
    // Set references to the form elements
    selectObj = document.forms[0].functionselect;
    textObj = document.forms[0].functioninput;
    selectedList = document.getElementById("selected");
    excludedList = document.getElementById("excluded");

    // Remember the function list length for loop speedup
    functionListLength = functionlist.length;

    // Set the search pattern depending
    if(document.forms[0].functionradio[0].checked == true)
    {
        searchPattern = "^"+textObj.value;
    }
    else
    {
        searchPattern = textObj.value;
    }
    searchPattern = cleanRegExpString(searchPattern);

    // Create a regulare expression

    re = new RegExp(searchPattern,"gi");
    // Clear the options list
    selectObj.length = 0;

    // Loop through the array and re-add matching options
    numShown = 0;
    for(i = 0; i < functionListLength; i++)
    {
        if(functionlist[i].search(re) != -1)
        {
            if (vallist[i] != "" && (selectedList.options.length < 1 || vallist[i] != selectedList.options[0].value)
              && (excludedList.options.length < 1 || vallist[i] != excludedList.options[0].value) ) {
                selectObj[numShown] = new Option(functionlist[i],vallist[i]);
                if (vallist[i] == 1) {
                    selectObj[numShown].style.fontWeight = 'bold';
                    selectObj[numShown].style.color = 'brown';
                }
                numShown++;
            }
        }
        // Stop when the number to show is reached
        if(numShown == maxNumToShow)
        {
            break;
        }
    }
    // When options list whittled to one, select that entry
    if(selectObj.length == 1)
    {
        selectObj.options[0].selected = true;
    }
}

function removeUsed()
{
    handleKeyUp(99999);
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
function handleSelectClick(itemSelected)
{
    textObj = document.forms[0].functioninput;
     
    selectObj = document.forms[0].functionselect;
    selectedValue = document.getElementById("functionselect").value;
    if(selectedValue != ""){ selectedText = selectObj[document.getElementById("functionselect").selectedIndex].text; }
    
    selectboxObj = document.forms[0].selected;
    selectedboxValue = document.getElementById("selected").value;
    if(selectedboxValue != ""){ selectedboxText = selectboxObj[document.getElementById("selected").selectedIndex].text; }
    
    excludedbox = document.forms[0].excluded;
    excludedboxValue = document.getElementById("excluded").value;
    if(excludedboxValue != ""){ excludeboxText = excludedbox[document.getElementById("excluded").selectedIndex].text; }
    
    if(itemSelected == "select1") {
        if(selectedValue != ""){
            // add items to selected box
            //document.getElementById('select1').disabled=true; //Disabled per RT#2919
            document.getElementById('select2').disabled=false;
            // someones adding all customers we need to empty the select box
            for (i = selectboxObj.length - 1; i>=0; i--) {
                selectboxObj.options[i] = null;
            }
            document.getElementById('deselect1').disabled=false;
            //document.getElementById('select1').disabled=true; //Disabled per RT#2919
            selectboxObj[selectboxObj.length] = new Option(selectedText,selectedValue);
        }
    }
    
    if(itemSelected == "deselect1") {
        if(selectedboxValue != ""){
            // remove items from selected box
            document.getElementById("selected").remove(document.getElementById("selected").selectedIndex)

            document.getElementById('select1').disabled=false;
            document.getElementById('select2').disabled=true;
            document.getElementById('deselect2').disabled=true;

            if(selectboxObj.length == 0) {
                // nothing in the select box so disable deselect
                document.getElementById('deselect1').disabled=true;
            }
        }
        if (document.getElementById("selected").options.length == 0) {
          document.getElementById('select1').disabled=false;
        }
    }
    
    if(itemSelected == "select2") {
      if(selectedValue != "" && selectedValue != "1"){
        // add items to excluded box
        excludedbox[excludedbox.length] = new Option(selectedText,selectedValue);
        if(excludedbox.length == 1) { 
            document.getElementById('select2').disabled=true; 
            // need to disable deselection on the selected box also since we added an exclude
            document.getElementById('deselect1').disabled=true; 
            document.getElementById('deselect2').disabled=false; 
        }
      } else if (selectedValue == "1") {
        alert('<%Sendb(Copient.PhraseLib.Lookup("term.anyproduct-not-excluded", LanguageID)) %>');
      }
    }
    
    if(itemSelected == "deselect2") {
        if(excludedboxValue != ""){
            // remove items from excluded box    
            document.getElementById("excluded").remove(document.getElementById("excluded").selectedIndex)
            if(excludedbox.length == 0) { 
                document.getElementById('select2').disabled=false; 
                document.getElementById('deselect1').disabled=false; 
                document.getElementById('deselect2').disabled=true; 
            }
        }
    }
    
    // remove items from large list that are in the other lists
    removeUsed();
    return true;
}

function saveForm(){
    var funcSel = document.getElementById('functionselect');
    var exSel = document.getElementById('excluded');
    var elSel = document.getElementById('selected');
    var i,j;
    var selectList = "";
    var excludededList = "";
    var htmlContents = "";

// assemble the list of values from the selected box
    for (i = elSel.length - 1; i>=0; i--) {
        if(elSel.options[i].value != ""){
            if(selectList != "") { selectList = selectList + ","; }
            selectList = selectList + elSel.options[i].value;
        }
    }
    for (i = exSel.length - 1; i>=0; i--) {
        if(exSel.options[i].value != ""){
            if(excludededList != "") { excludededList = excludededList + ","; }
            excludededList = excludededList + exSel.options[i].value;
        }
    }
    
    document.getElementById("ProductGroupID").value = selectList;
    document.getElementById("ExcludedID").value = excludededList;
    // alert(htmlContents);
    return true;
}

function updateButtons(){
  if(document.getElementById('selected').length > 0){
      if(document.forms[0].excluded.length == 0) { 
        //document.getElementById('select1').disabled=true; //Disabled per RT#2919
        document.getElementById('deselect1').disabled=false; 
        document.getElementById('select2').disabled=false; 
        document.getElementById('deselect2').disabled=true; 
      } else {
        //document.getElementById('select1').disabled=true; //Disabled per RT#2919
        document.getElementById('deselect1').disabled=true; 
        document.getElementById('select2').disabled=true; 
        document.getElementById('deselect2').disabled=false; 
      }
  } else {
    document.getElementById('select1').disabled=false; 
    document.getElementById('deselect1').disabled=true; 
    document.getElementById('select2').disabled=true; 
    document.getElementById('deselect2').disabled=true; 
  }
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
  
  ' we need to determine our linkid for updates and tiered
  MyCommon.QueryStr = "select LinkID,Tiered,SponsorID,PromoteToTransLevel,RewardDistPeriod,RewardLimit,RewardLimitTypeID,TriggerQty,RewardAmountTypeID, " & _
     "UseSpecialPricing, SPRepeatAtOccur,ApplyToLimit,DoNotItemDistribute from OfferRewards with (NoLock) where RewardID=" & RewardID
	rst = MyCommon.LRT_Select
  For Each row In rst.Rows
    DistPeriod = MyCommon.NZ(row.Item("RewardDistPeriod"), 0)
    LinkID = row.Item("LinkID")
    Tiered = MyCommon.NZ(row.Item("Tiered"), 0)
    SponsorID = MyCommon.NZ(row.Item("SponsorID"), 0)
    PromoteToTransLevel = MyCommon.NZ(row.Item("PromoteToTransLevel"), 0)
    RewardLimit = MyCommon.NZ(row.Item("RewardLimit"), 0)
    RewardLimitTypeID = MyCommon.NZ(row.Item("RewardLimitTypeID"), 2)
    RewardAmountTypeID = MyCommon.NZ(row.Item("RewardAmountTypeID"), 1)
    'ExItemLevelDist = MyCommon.NZ(row.Item("ExItemLevelDist"), 0)
    TriggerQty = MyCommon.NZ(row.Item("TriggerQty"), 1)
    ApplyToLimit = MyCommon.NZ(row.Item("ApplyToLimit"), 1)
    UseSpecialPricing = MyCommon.NZ(row.Item("UseSpecialPricing"), 0)
    SPRepeatAtOccur = MyCommon.NZ(row.Item("SPRepeatAtOccur"), 1)
    DoNotItemDistribute = row.Item("DoNotItemDistribute")
  Next
  
  If (TriggerQty = ApplyToLimit And TriggerQty <> 0) Then
    ValueRadio = 1
  Else
    ValueRadio = 2
  End If
  
  If (Request.QueryString("save") <> "") Then
    Select Case ProductGroupID
      Case 0
        MyCommon.QueryStr = "update OfferRewards with (RowLock) set ProductGroupID=0, ExcludedProdGroupID=0 where RewardID=" & RewardID & " and deleted=0;"
        MyCommon.LRT_Execute()
      Case 1
        MyCommon.QueryStr = "update OfferRewards with (RowLock) set ProductGroupID=" & ProductGroupID & ", ExcludedProdGroupID=" & ExcludedID & " where RewardID=" & RewardID & " and deleted=0;"
        MyCommon.LRT_Execute()
      Case Else
        If (ExcludedID > 0) Then
          MyCommon.QueryStr = "update OfferRewards with (RowLock) set ProductGroupID=" & ProductGroupID & ", ExcludedProdGroupID=" & ExcludedID & " where RewardID=" & RewardID & " and deleted=0;"
        Else
          MyCommon.QueryStr = "update OfferRewards with (RowLock) set ProductGroupID=" & ProductGroupID & ", ExcludedProdGroupID=0 where RewardID=" & RewardID & " and deleted=0;"
        End If
        MyCommon.LRT_Execute()
    End Select

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
    ' save the limit and also create the var id while saving it if the limit is greater than zero
    If (MyCommon.Extract_Val(Request.QueryString("limitvalue")) <> "") Then
      If (MyCommon.Extract_Val(Request.QueryString("limitvalue")) < 0) Then
        infoMessage = Copient.PhraseLib.Lookup("reward.badlimit", LanguageID)
      End If
      MyCommon.QueryStr = "update offerRewards with (RowLock) set RewardLimit=" & MyCommon.Extract_Val(Request.QueryString("limitvalue")) & " where RewardID=" & RewardID
      MyCommon.LRT_Execute()
    End If
    If (Request.QueryString("sponsor") <> "") Then
      MyCommon.QueryStr = "update offerRewards with (RowLock) set SponsorID=" & MyCommon.Extract_Val(Request.QueryString("sponsor")) & " where RewardID=" & RewardID
      MyCommon.LRT_Execute()
    End If
    If (Request.QueryString("form_DistPeriod") <> "") Then
      If (Request.QueryString("form_DistPeriod") < 0) Then
        infoMessage = Copient.PhraseLib.Lookup("reward.badlimit", LanguageID)
      End If
      MyCommon.QueryStr = "update OfferRewards with (RowLock) set RewardDistPeriod ='" & Int(MyCommon.Extract_Val(Request.QueryString("form_DistPeriod"))) & "' where RewardID=" & RewardID
      MyCommon.LRT_Execute()
      DistPeriod = MyCommon.Extract_Val(Request.QueryString("form_DistPeriod"))
    End If
    ' someone saves - let's do the special case and set a promo variable if the
    ' distribution's greater than zero and the promo variable doesn't already exist
    If (Request.QueryString("form_DistPeriod") > 0) Then
      MyCommon.QueryStr = "select RewardDistLimitVarID from OfferRewards with (NoLock) where RewardID=" & RewardID
      rst = MyCommon.LRT_Select
      For Each row In rst.Rows
        If (MyCommon.NZ(row.Item("RewardDistLimitVarID"), 0) = 0) Then
          'dbo.pa_DistributionVar_Create @OfferID bigint, @VarID bigint OUTPUT
          MyCommon.Open_LogixXS()
          MyCommon.QueryStr = "dbo.pc_RewardLimitVar_Create "
          MyCommon.Open_LXSsp()
          MyCommon.LXSsp.Parameters.Add("@RewardID", SqlDbType.BigInt).Value = RewardID
          MyCommon.LXSsp.Parameters.Add("@VarID", SqlDbType.BigInt).Direction = ParameterDirection.Output
          MyCommon.LXSsp.ExecuteNonQuery()
          VarID = MyCommon.LXSsp.Parameters("@VarID").Value
          MyCommon.Close_LXSsp()
          MyCommon.Close_LogixXS()
          MyCommon.QueryStr = "update OfferRewards with (RowLock) set RewardDistLimitVarID=" & VarID & " where RewardID=" & RewardID
          MyCommon.LRT_Execute()
        End If
      Next
    End If
    
    ' ok here we need to handle the tiering stuffs
    If (Tiered = 0) Then
      MyCommon.QueryStr = "dbo.pt_XmlTiers_Update"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@RewardID", SqlDbType.BigInt).Value = RewardID
      MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int).Value = 0
      MyCommon.LRTsp.Parameters.Add("@XmlText", SqlDbType.NVarChar, 4000).Value = Request.QueryString("tier0")
      MyCommon.LRTsp.ExecuteNonQuery()
      MyCommon.Close_LRTsp()
    Else
      ' delete the current tier ammounts
      MyCommon.QueryStr = "delete from RewardXmlTiers with (RowLock) where RewardID=" & RewardID
      MyCommon.LRT_Execute()
      For x = 1 To NumTiers
        MyCommon.QueryStr = "dbo.pt_XmlTiers_Update"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@RewardID", SqlDbType.BigInt).Value = RewardID
        MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int).Value = x
        MyCommon.LRTsp.Parameters.Add("@XmlText", SqlDbType.NVarChar, 4000).Value = Request.QueryString("tier" & x)
        MyCommon.LRTsp.ExecuteNonQuery()
        MyCommon.Close_LRTsp()
      Next
    End If
    If (Request.QueryString("trigger") <> "") Then
      If (Request.QueryString("trigger") = "1") Then
        ' set  TriggerQty=Xbox
        Dim tempTrigger As Integer
        tempTrigger = MyCommon.Extract_Val(Request.QueryString("Xbox"))
        If (tempTrigger = 0) Then
          tempTrigger = 1
        End If
        MyCommon.QueryStr = "update OfferRewards with (RowLock) set TriggerQty=" & tempTrigger & "," & _
                            "ApplyToLimit=" & tempTrigger & " where RewardID=" & RewardID
        MyCommon.LRT_Execute()
        TriggerQty = tempTrigger
        ValueRadio = 1
      ElseIf (Request.QueryString("trigger") = "2") Then
        ' Set  and TriggerQty=Xbox2+Ybox2 and ApplyToLimit=Ybox2
        MyCommon.QueryStr = "update OfferRewards with (RowLock) set TriggerQty=" & Int(MyCommon.Extract_Val(Request.QueryString("Xbox2"))) + Int(MyCommon.Extract_Val(Request.QueryString("Ybox2"))) & "," & _
                            "ApplyToLimit=" & MyCommon.Extract_Val(Request.QueryString("Ybox2")) & " where RewardID=" & RewardID
        'Response.Write(MyCommon.QueryStr)
        MyCommon.LRT_Execute()
        ValueRadio = 2
        TriggerQty = Int(MyCommon.Extract_Val(Request.QueryString("Xbox2"))) + Int(MyCommon.Extract_Val(Request.QueryString("Ybox2")))
        ApplyToLimit = MyCommon.Extract_Val(Request.QueryString("Ybox2"))
        'If (TriggerQty = ApplyToLimit) Then ValueRadio = 1
      End If
    End If
    SponsorID = MyCommon.Extract_Val(Request.QueryString("sponsor"))
    RewardLimit = MyCommon.Extract_Val(Request.QueryString("limitvalue"))
    MyCommon.QueryStr = "update OfferRewards with (RowLock) set TCRMAStatusFlag=2,CMOAStatusFlag=2 where RewardID=" & RewardID
    MyCommon.LRT_Execute()
    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.xmlgeneric", LanguageID))

    MyCommon.QueryStr = "update OfferRewards with (RowLock) set TCRMAStatusFlag=3,CMOAStatusFlag=2 where RewardID=" & RewardID
    MyCommon.LRT_Execute()
    MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where offerid=" & OfferID
    MyCommon.LRT_Execute()

    CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
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
<form action="offer-rew-xml.aspx" id="mainform" name="mainform" onsubmit="return saveForm();">
	<div id="intro">
		<input type="hidden" name="OfferID" value="<% sendb(OfferID) %>" />
		<input type="hidden" name="Name" value="<% sendb(Name) %>" />
		<input type="hidden" name="RewardID" value="<% sendb(RewardID) %>" />
    <input type="hidden" id="ProductGroupID" name="ProductGroupID" value="<% sendb(ProductGroupID) %>" />
    <input type="hidden" id="ExcludedID" name="ExcludedID" value="<% sendb(ExcludedID) %>" />
		<input type="hidden" name="IsTemplate" value="<%
		 If (istemplate) Then 
    sendb("IsTemplate")
    Else 
    sendb("Not") 
    End If
    %>" />
		<h1 id="title"><% Sendb(Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & Copient.PhraseLib.Lookup("term.xmlgenericreward", LanguageID))%></h1>
		<%If (IsTemplate) Then%>
		<span class="temp2">
			<input type="checkbox" class="tempcheck" id="temp-employees" name="Disallow_Edit"<% if(disallow_edit)then sendb(" checked=""checked""") %> />
			<label for="temp-employees"><% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
		</span>
		<%End If%>
		<div id="controls">
			<% If Not (IsTemplate) Then
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
      <% Send_ProductConditionSelector(Logix, TransactionLevelSelected, FromTemplate, Disallow_Edit, SelectedItem, ExcludedItem, RewardID, Copient.CommonInc.InstalledEngines.CM)%>
		</div>
		
		<div id="gutter">
		</div>
		
		<div id="column2">
			<%If Not TransactionLevelSelected Then%>
			<div class="box" id="distribution">
				<h2><span><% Sendb(Copient.PhraseLib.Lookup("term.distribution", LanguageID))%></span></h2>
				<input id="triggerbogo" name="trigger" type="radio"<% if(valueradio=1)then sendb(" checked=""checked""") %> value="1" />
				<label for="triggerbogo"><% Sendb(Copient.PhraseLib.Lookup("reward.messageevery", LanguageID))%></label>
				<br />
				&nbsp; &nbsp; &nbsp; &nbsp;
				<label for="Xbox"><% Sendb(Copient.PhraseLib.Lookup("term.mustpurchase", LanguageID))%></label>
				<input class="shortest" id="Xbox" name="Xbox" type="text"<% if(valueradio=1)then sendb(" value=""" & triggerqty & """ ") %> />
				<% Sendb(Copient.PhraseLib.Lookup("term.item(s)", LanguageID))%>
				<br />
				<input id="triggerbxgy" name="trigger" type="radio" value="2"<% if(valueradio=2)then sendb(" checked=""checked""") %> />
				<label for="triggerbxgy"><% Sendb(Copient.PhraseLib.Lookup("term.buy", LanguageID))%></label>
				<input class="shortest" id="bxgy1" name="Xbox2" type="text"<% if(valueradio=2)then sendb(" value=""" & triggerqty-applytolimit & """ ") %> />
				<% Sendb(Copient.PhraseLib.Lookup("term.item(s)", LanguageID))%>, <% Sendb(Copient.PhraseLib.Lookup("reward.givemessageto", LanguageID))%>
				<input class="shortest" id="bxgy2" name="Ybox2" type="text"<% if(valueradio=2)then sendb(" value=""" & applytolimit & """ ") %> /><br />
				<hr class="hidden" />
			</div>
			<% End If%>
      <div class="box" id="limits">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.limits", LanguageID))%>
          </span>
        </h2>
        <input class="shorter" id="limitvalue" maxlength="9" name="limitvalue" type="text" value="<% sendb(RewardLimit) %>" />
        <% Sendb(Copient.PhraseLib.Lookup("term.count", LanguageID))%>
        &nbsp;<% Sendb(Copient.PhraseLib.Lookup("term.per", LanguageID))%>
        <input class="shorter" type="text" maxlength="4" id="limitperiod" name="form_DistPeriod"<% if(distperiod=0)then sendb(" style=""visibility: hidden;""")%> value="<% sendb(DistPeriod) %>" />
        <select onchange="javascript:if(this.value==1){document.mainform.form_DistPeriod.style.visibility ='visible';document.mainform.form_DistPeriod.value=0}else{document.mainform.form_DistPeriod.style.visibility ='hidden';document.mainform.form_DistPeriod.value=0;}">
          <option value="1"<% if(distperiod>0)then sendb(" selected=""selected""") %>>
            <% Sendb(Copient.PhraseLib.Lookup("term.days", LanguageID))%>
          </option>
          <option value="2"<% if(distperiod=0)then sendb(" selected=""selected""") %>>
            <% Sendb(Copient.PhraseLib.Lookup("term.transaction", LanguageID))%>
          </option>
        </select>
        <hr class="hidden" />
      </div>
			<div class="box" id="sponsor">
				<h2><span><% Sendb(Copient.PhraseLib.Lookup("term.sponsor", LanguageID))%></span></h2>
        <%
          MyCommon.QueryStr = "select SponsorID, Description, PhraseID from Sponsors with (NoLock)"
          rst = MyCommon.LRT_Select()
          For Each row In rst.Rows
            Sendb("<input class=""radio"" id=""" & StrConv(row.Item("Description"), VbStrConv.Lowercase) & """ name=""sponsor"" type=""radio"" value=""" & row.Item("SponsorID") & """")
            If SponsorID = row.Item("SponsorID") Then
              Sendb(" checked=""checked""")
            End If
            Send(" /><label for=""" & StrConv(row.Item("Description"), VbStrConv.Lowercase) & """>" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID) & "</label>")
          Next
        %>
				<hr class="hidden" />
			</div>
			<div class="box" id="message">
				<h2>
				  <span>
				    <% Sendb("XML")%>
				  </span>
				</h2>
				<%
				  MyCommon.QueryStr = "select OFR.RewardID,Tiered,O.Numtiers,O.OfferID,XT.TierLevel,XT.XmlText from OfferRewards as OFR with (NoLock) " & _
				                      "left join Offers as O with (NoLock) on O.OfferID=OFR.OfferID left join RewardXmlTiers as XT with (NoLock) on OFR.RewardID=XT.RewardID where OFR.RewardID=" & RewardID
				  rst = MyCommon.LRT_Select()
					q = 1
					For Each row In rst.Rows
						If (row.Item("Tiered") = False) Then
							Send("    <div class=""pmsgwrap"">")
							Send("     <label for=""tier0""><b>" & "XML Text" & "</b></label><br />")
				      Send("     <textarea id=""tier0"" name=""tier0"" cols=""35"" rows=""17"" wrap=""hard"">" & row.Item("XmlText") & "</textarea><br />")
				      'Send("     <script type=""text/javascript"">edToolbar();</script>")
							Send("    </div>")
							Send("    <script type=""text/javascript"">var edCanvas = document.getElementById('tier0');</script>")
						Else
							Send("    <div class=""pmsgwrap"">")
							Send("     <label for=""tier" & q & """><b>" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & q & " " & "XML Text" & ":</b></label><br />")
				      Send("     <textarea id=""tier" & q & """ name=""tier" & q & """ cols=""35"" rows=""17"" wrap=""hard"">" & row.Item("XmlText") & "</textarea><br />")
							If q = 1 Then
				        'Send("     <script type=""text/javascript"">edToolbar();</script>")
							End If
							Send("    </div>")
							If q = 1 Then
								Send("    <script type=""text/javascript"">var edCanvas = document.getElementById('tier" & q & "');</script>")
							End If
						End If
						q = q + 1
					Next
				  Send("     <input type=""hidden"" id=""NumTiers"" name=""NumTiers"" value=""" & row.Item("NumTiers") & """ />")
				%>
				<br class="half" />
				<br />
			</div>
		</div>
	</div>
</form>
<script type="text/javascript">
<% If (CloseAfterSave) Then %>
    window.close();
<% End If %>
  updateButtons();
  removeUsed();
</script>

<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd()
  Logix = Nothing
  MyCommon = Nothing
%>
